"""
Outlook to Gmail Forwarder

Installation requirements (run once):
    pip install selenium webdriver-manager google-api-python-client google-auth-httplib2 \
                google-auth-oauthlib streamlit python-dateutil

Run the Streamlit UI:
    streamlit run app.py

Test on a non-production (dummy) account before using with your primary accounts.
"""
import base64
import json
import os
import random
import shutil
import threading
import time
from collections import deque
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Any, Deque, Dict, List, Optional, Tuple

import streamlit as st
import streamlit.components.v1 as components
from dateutil import tz
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


# --------------------------------------------------------------------------------------
# Paths & Configuration
# --------------------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "automation_state"
DATA_DIR.mkdir(exist_ok=True)

CHROME_PROFILE_DIR = DATA_DIR / "chrome_profile"
PROFILE_READY_PATH = DATA_DIR / "profile_ready.txt"
FORWARDED_LOG_PATH = DATA_DIR / "forwarded.json"
FORWARD_STATE_PATH = DATA_DIR / "daily_counter.json"
SETTINGS_PATH = DATA_DIR / "settings.json"
GOOGLE_TOKEN_PATH = DATA_DIR / "token.json"

OUTLOOK_LOGIN_URL = "https://outlook.office.com/mail/"
OUTLOOK_INBOX_URL = "https://outlook.office.com/mail/inbox"
OUTLOOK_JUNK_URL = "https://outlook.office.com/mail/junkemail"
OUTLOOK_SENT_URL = "https://outlook.office.com/mail/sentitems"
OUTLOOK_DRAFTS_URL = "https://outlook.office.com/mail/drafts"
OUTLOOK_DELETED_URL = "https://outlook.office.com/mail/deleteditems"
OUTLOOK_ARCHIVE_URL = "https://outlook.office.com/mail/archive"
OUTLOOK_OUTBOX_URL = "https://outlook.office.com/mail/outbox"

OUTLOOK_FOLDERS = [
    ("Inbox", OUTLOOK_INBOX_URL),
    ("Junk Email", OUTLOOK_JUNK_URL),
    ("Sent Items", OUTLOOK_SENT_URL),
    ("Drafts", OUTLOOK_DRAFTS_URL),
    ("Deleted Items", OUTLOOK_DELETED_URL),
    ("Archive", OUTLOOK_ARCHIVE_URL),
    ("Outbox", OUTLOOK_OUTBOX_URL),
]

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

if "MANUAL_LOGIN_EVENT" not in globals():
    MANUAL_LOGIN_EVENT = threading.Event()
if "STOP_EVENT" not in globals():
    STOP_EVENT = threading.Event()
if "WORKER_THREAD" not in globals():
    WORKER_THREAD: Optional[threading.Thread] = None

try:
    session_holder = st.session_state
except RuntimeError:
    if "MANUAL_DRIVER_HOLDER" not in globals():
        MANUAL_DRIVER_HOLDER: Dict[str, Optional[webdriver.Chrome]] = {"driver": None}
else:
    if "MANUAL_DRIVER_HOLDER" not in session_holder:
        session_holder["MANUAL_DRIVER_HOLDER"] = {"driver": None}
    MANUAL_DRIVER_HOLDER = session_holder["MANUAL_DRIVER_HOLDER"]

if "LOG_BUFFER" not in globals():
    LOG_BUFFER: Deque[str] = deque(maxlen=500)
LOG_LOCK = threading.Lock()


# --------------------------------------------------------------------------------------
# Utility helpers
# --------------------------------------------------------------------------------------
def log_message(message: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry = f"[{timestamp}] {message}"
    with LOG_LOCK:
        LOG_BUFFER.append(entry)
    print(entry, flush=True)


def human_delay(min_seconds: float = 1.0, max_seconds: float = 2.0) -> None:
    time.sleep(random.uniform(min_seconds, max_seconds))


def _coerce_minutes(value: Any, fallback: int, minimum: int = 1, maximum: int = 30) -> int:
    try:
        coerced = int(float(value))
    except (TypeError, ValueError):
        return fallback
    return max(minimum, min(maximum, coerced))


class ManualLoginRequired(Exception):
    """Raised when a manual login is required."""


class ManualLoginPending(ManualLoginRequired):
    """Raised when automation is waiting for the user to finish logging in."""


class CaptchaDetected(Exception):
    """Raised when Outlook displays a CAPTCHA or block."""


class SettingsManager:
    """Persist lightweight settings (Gmail target address, polling interval, etc.)."""

    def __init__(self) -> None:
        self.path = SETTINGS_PATH
        self._settings: Dict[str, Any] = {}
        if self.path.exists():
            try:
                self._settings = json.loads(self.path.read_text(encoding="utf-8"))
            except json.JSONDecodeError:
                self._settings = {}

    def get(self, key: str, default: Any = None) -> Any:
        return self._settings.get(key, default)

    def set(self, key: str, value: Any) -> None:
        self._settings[key] = value
        self.path.write_text(json.dumps(self._settings, indent=2), encoding="utf-8")


class ForwardedRegistry:
    """Keep track of Outlook message IDs that have already been forwarded."""

    def __init__(self, path: Path) -> None:
        self.path = path
        if path.exists():
            try:
                self.registry = set(json.loads(path.read_text(encoding="utf-8")))
            except json.JSONDecodeError:
                self.registry = set()
        else:
            self.registry = set()
        self.lock = threading.Lock()

    def has(self, message_id: str) -> bool:
        with self.lock:
            return message_id in self.registry

    def add(self, message_id: str) -> None:
        with self.lock:
            self.registry.add(message_id)
            self.path.write_text(json.dumps(sorted(self.registry)), encoding="utf-8")


class DailyCounter:
    """Track number of forwarded emails per day."""

    def __init__(self, path: Path) -> None:
        self.path = path
        self.lock = threading.Lock()
        if path.exists():
            try:
                data = json.loads(path.read_text(encoding="utf-8"))
                self.day = data.get("day")
                self.count = data.get("count", 0)
            except json.JSONDecodeError:
                self.day = datetime.now(timezone.utc).date().isoformat()
                self.count = 0
        else:
            self.day = datetime.now(timezone.utc).date().isoformat()
            self.count = 0

    def increment(self) -> int:
        with self.lock:
            today = datetime.now(timezone.utc).date().isoformat()
            if today != self.day:
                self.day = today
                self.count = 0
            self.count += 1
            self._persist()
            return self.count

    def get_count(self) -> int:
        with self.lock:
            today = datetime.now(timezone.utc).date().isoformat()
            if today != self.day:
                self.day = today
                self.count = 0
                self._persist()
            return self.count

    def _persist(self) -> None:
        payload = {"day": self.day, "count": self.count}
        self.path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


@dataclass
class EmailContent:
    message_id: str
    sender: str
    subject: str
    body_html: str
    body_text: str


class GmailForwarder:
    """Send emails using Gmail API."""

    def __init__(self, settings: SettingsManager) -> None:
        self.settings = settings
        self._service = None
        self.lock = threading.Lock()
        self._client_secret_path: Optional[Path] = None

    def _resolve_client_secret_file(self) -> Path:
        if self._client_secret_path and self._client_secret_path.exists():
            return self._client_secret_path

        default_path = BASE_DIR / "credentials.json"
        if default_path.exists():
            self._client_secret_path = default_path
            log_message(f"Using Google client secret file: {default_path.name}")
            return default_path

        candidate_dirs = [BASE_DIR, DATA_DIR]
        candidate_files: List[Path] = []
        for directory in candidate_dirs:
            if not directory.exists():
                continue
            for path in directory.glob("*.json"):
                if path.name.endswith(".apps.googleusercontent.com.json"):
                    candidate_files.append(path)
            for path in directory.glob("*.apps.googleusercontent.com"):
                if path.is_file():
                    candidate_files.append(path)

        if candidate_files:
            candidate_files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
            self._client_secret_path = candidate_files[0]
            log_message(f"Using Google client secret file: {self._client_secret_path.name}")
            return self._client_secret_path

        raise FileNotFoundError(
            "Google OAuth client secret JSON not found. Place your downloaded client secret file "
            "(e.g., client_secret_XXXX.apps.googleusercontent.com.json) in the project folder."
        )

    def _build_service(self):
        creds = None
        if GOOGLE_TOKEN_PATH.exists():
            creds = Credentials.from_authorized_user_file(str(GOOGLE_TOKEN_PATH), SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                client_secret_path = self._resolve_client_secret_file()
                flow = InstalledAppFlow.from_client_secrets_file(str(client_secret_path), SCOPES)
                creds = flow.run_local_server(port=0)
            GOOGLE_TOKEN_PATH.parent.mkdir(exist_ok=True)
            with GOOGLE_TOKEN_PATH.open("w", encoding="utf-8") as token:
                token.write(creds.to_json())
        return build("gmail", "v1", credentials=creds, cache_discovery=False)

    def _ensure_service(self):
        with self.lock:
            if self._service is None:
                self._service = self._build_service()
        return self._service

    def send_email(self, to_email: str, subject: str, body_html: str, body_text: str) -> None:
        service = self._ensure_service()
        message = MIMEMultipart("alternative")
        message["to"] = to_email
        message["subject"] = subject or "(no subject)"
        message.attach(MIMEText(body_text or "(no body)", "plain", "utf-8"))
        message.attach(MIMEText(body_html or body_text or "(no body)", "html", "utf-8"))
        raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
        try:
            service.users().messages().send(userId="me", body={"raw": raw_message}).execute()
            log_message(f"Sent email to {to_email} | subject='{subject}'")
        except HttpError as exc:  # noqa: BLE001
            log_message(f"Gmail API error: {exc}")
            raise

    def send_alert(self, to_email: str, reason: str) -> None:
        subject = "Outlook check failed - possible block. Check manually."
        text_body = f"Automation failed with reason: {reason}\nPlease sign in manually."
        html_body = f"<p>Automation failed with reason: <strong>{reason}</strong></p><p>Please sign in manually.</p>"
        self.send_email(to_email, subject, html_body, text_body)

    def send_test_email(self, to_email: str) -> None:
        timestamp = datetime.now().astimezone(tz.tzlocal()).strftime("%Y-%m-%d %H:%M:%S %Z")
        subject = "Outlook ➜ Gmail Forwarder test message"
        text_body = (
            "Hello!\n\n"
            "This is a verification email sent by the Outlook ➜ Gmail Forwarder to confirm that "
            "your Gmail connection is working.\n\n"
            f"Timestamp: {timestamp}\n"
        )
        html_body = (
            "<p>Hello!</p>"
            "<p>This is a verification email sent by the <strong>Outlook ➜ Gmail Forwarder</strong> to confirm "
            "that your Gmail connection is working.</p>"
            f"<p>Timestamp: <code>{timestamp}</code></p>"
        )
        log_message(f"Sending Gmail connectivity test email to {to_email}.")
        self.send_email(to_email, subject, html_body, text_body)


class OutlookAutomation:
    """Selenium automation for Outlook Web."""

    def _get_existing_driver(self) -> Optional[webdriver.Chrome]:
        driver = MANUAL_DRIVER_HOLDER.get("driver")
        if not driver:
            return None
        try:
            driver.execute_script("return document.readyState")
            return driver
        except Exception:  # noqa: BLE001
            try:
                driver.quit()
            except Exception:  # noqa: BLE001
                pass
            MANUAL_DRIVER_HOLDER["driver"] = None
            return None

    def _create_driver(self, headless: bool = True, use_profile: bool = False) -> webdriver.Chrome:
        options = Options()
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--disable-extensions")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--lang=en-US")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
        if use_profile:
            CHROME_PROFILE_DIR.mkdir(parents=True, exist_ok=True)
            options.add_argument(f"--user-data-dir={CHROME_PROFILE_DIR}")
            options.add_argument("--profile-directory=Default")
        if headless:
            options.add_argument("--headless=new")
        binary_candidates = [
            os.environ.get("CHROME_BINARY"),
            os.environ.get("GOOGLE_CHROME_SHIM"),
            shutil.which("google-chrome"),
            shutil.which("chrome"),
            shutil.which("chromium"),
            shutil.which("chromium-browser"),
            shutil.which("msedge"),
            shutil.which("msedge.exe"),
        ]
        windows_defaults = [
            Path(os.environ.get("PROGRAMFILES", "")) / "Google/Chrome/Application/chrome.exe",
            Path(os.environ.get("PROGRAMFILES(X86)", "")) / "Google/Chrome/Application/chrome.exe",
            Path(os.environ.get("LOCALAPPDATA", "")) / "Google/Chrome/Application/chrome.exe",
        ]
        for candidate in windows_defaults:
            if candidate and candidate.is_file():
                binary_candidates.append(str(candidate))
        binary_location = next((candidate for candidate in binary_candidates if candidate), None)
        if binary_location:
            options.binary_location = binary_location
            log_message(f"Using Chrome binary at: {binary_location}")
        else:
            log_message(
                "Chrome binary not found automatically. Attempting to start ChromeDriver without an explicit binary path."
            )
        try:
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        except WebDriverException as exc:
            message = str(exc)
            lowered = message.lower()
            if "user data directory is already in use" in lowered:
                raise RuntimeError(
                    "Chrome profile already in use. Close other Chrome windows that use this profile and try again."
                ) from exc
            if "cannot find chrome binary" in lowered or "chrome binary" in lowered:
                raise RuntimeError(
                    "Unable to locate a Chrome or Chromium browser binary. Install Chrome/Chromium or set the CHROME_BINARY "
                    "environment variable to the browser executable before launching the manual login."
                ) from exc
            raise RuntimeError(f"Unable to start Chrome: {message}") from exc
        driver.set_window_size(1400, 900)
        return driver

    @staticmethod
    def _first_present(wait: WebDriverWait, locators: List[Tuple[str, str]]):
        """Return the first element located from the provided strategies."""

        for by, value in locators:
            try:
                element = wait.until(EC.presence_of_element_located((by, value)))
                if element:
                    return element
            except TimeoutException:
                continue
            except Exception:  # noqa: BLE001
                continue
        return None

    @staticmethod
    def _human_mouse_move(driver: webdriver.Chrome, target_element: Optional[Any] = None) -> None:
        """Simulate a short sequence of mouse movements to appear more human."""

        try:
            width, height = driver.execute_script(
                "return [window.innerWidth || document.documentElement.clientWidth || 0, "
                "window.innerHeight || document.documentElement.clientHeight || 0];"
            )
            if not width or not height:
                return

            width = max(int(width), 1)
            height = max(int(height), 1)

            start_x = random.randint(int(width * 0.2), int(width * 0.8))
            start_y = random.randint(int(height * 0.2), int(height * 0.8))

            steps = random.randint(3, 6)
            points: List[Tuple[int, int]] = [(start_x, start_y)]
            current_x, current_y = start_x, start_y
            for _ in range(steps - 1):
                current_x = max(5, min(width - 5, current_x + random.randint(-90, 90)))
                current_y = max(5, min(height - 5, current_y + random.randint(-70, 70)))
                points.append((current_x, current_y))

            body = driver.find_element(By.TAG_NAME, "body")
            actions = ActionChains(driver)
            actions.move_to_element_with_offset(body, points[0][0], points[0][1])
            prev_x, prev_y = points[0]
            for x, y in points[1:]:
                actions.pause(random.uniform(0.05, 0.18))
                actions.move_by_offset(x - prev_x, y - prev_y)
                prev_x, prev_y = x, y

            if target_element is not None:
                actions.pause(random.uniform(0.08, 0.25))
                actions.move_to_element(target_element)
                actions.pause(random.uniform(0.05, 0.15))

            actions.perform()
        except Exception:  # noqa: BLE001
            if target_element is not None:
                try:
                    ActionChains(driver).move_to_element(target_element).perform()
                except Exception:  # noqa: BLE001
                    pass

    def _safe_click(self, driver: webdriver.Chrome, element: Any) -> bool:
        """Click an element with fallbacks to reduce flakiness."""

        try:
            self._human_mouse_move(driver, element)
            ActionChains(driver).pause(random.uniform(0.1, 0.3)).click().perform()
            return True
        except WebDriverException:
            try:
                driver.execute_script("arguments[0].click();", element)
                return True
            except Exception:  # noqa: BLE001
                return False

    @staticmethod
    def _wait_for_folder_navigation(driver: webdriver.Chrome, folder_url: str) -> bool:
        """Wait for the current URL to reflect the requested folder."""

        try:
            target_fragment = folder_url.rstrip("/").split("/mail/")[-1].lower()
            if not target_fragment:
                return False
            WebDriverWait(driver, 10).until(
                lambda d: target_fragment in ((d.current_url or "").lower())
            )
            return True
        except TimeoutException:
            return False

    def _open_folder_by_click(
        self, driver: webdriver.Chrome, folder_name: str, folder_url: str
    ) -> bool:
        """Attempt to switch Outlook folders via sidebar clicks instead of direct URL loads."""

        sidebar_wait = WebDriverWait(driver, 12)
        normalized_name = folder_name.strip()
        # Common selectors for Outlook sidebar nodes. We try multiple strategies
        # because Microsoft frequently tweaks the DOM structure.
        locator_candidates: List[Tuple[str, str]] = [
            (By.CSS_SELECTOR, f"button[title='{normalized_name}']"),
            (By.CSS_SELECTOR, f"div[role='treeitem'][aria-label='{normalized_name}']"),
            (By.XPATH, f"//div[@role='treeitem']//span[normalize-space(text())='{normalized_name}']"),
            (By.XPATH, f"//span[@title='{normalized_name}']"),
            (By.XPATH, f"//button[@aria-label='{normalized_name}']"),
        ]

        for by, value in locator_candidates:
            try:
                element = sidebar_wait.until(EC.presence_of_element_located((by, value)))
            except TimeoutException:
                continue
            except Exception:  # noqa: BLE001
                continue

            clickable = element
            try:
                if clickable.tag_name.lower() == "span":
                    clickable = clickable.find_element(
                        By.XPATH, "ancestor::*[@role='treeitem' or @role='button'][1]"
                    )
            except Exception:  # noqa: BLE001
                pass

            if clickable is None:
                continue

            if not self._safe_click(driver, clickable):
                continue

            human_delay(0.6, 1.2)
            if self._wait_for_folder_navigation(driver, folder_url):
                return True

        return False

    def profile_ready(self) -> bool:
        if not PROFILE_READY_PATH.exists():
            return False
        if not CHROME_PROFILE_DIR.exists():
            return False
        return any(CHROME_PROFILE_DIR.iterdir())

    def launch_manual_login(self, *, auto_open: bool = False) -> None:
        existing_driver = self._get_existing_driver()
        if existing_driver:
            log_message("Persistent Outlook window already running.")
            return
        driver = self._create_driver(headless=False, use_profile=True)
        MANUAL_DRIVER_HOLDER["driver"] = driver
        target_url = OUTLOOK_INBOX_URL if self.profile_ready() else OUTLOOK_LOGIN_URL
        if auto_open:
            log_message("Outlook window opened automatically using the saved profile.")
        else:
            log_message(
                "Manual Chrome window launched with a persistent profile. Log in, solve any prompts, and then click 'Save session'."
            )
        driver.get(target_url)

    def complete_manual_login(self) -> None:
        driver = self._get_existing_driver()
        if not driver:
            log_message("No manual Chrome session is active.")
            return
        try:
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '[role="main"]'))
            )
        except TimeoutException as exc:
            raise RuntimeError(
                "Outlook web app not detected yet. Finish the login process before saving the session."
            ) from exc
        PROFILE_READY_PATH.write_text(datetime.now().isoformat())
        log_message(
            "Saved persistent Outlook session. Leave the Outlook window open so automation can reuse it."
        )
        MANUAL_LOGIN_EVENT.clear()

    def _detect_captcha(self, driver: webdriver.Chrome) -> bool:
        """Detect whether Outlook is presenting a bot or verification challenge."""

        page_text = driver.page_source.lower()
        title = (driver.title or "").lower()
        captcha_keywords = ["captcha", "enter the characters you see", "verification challenge"]
        alert_phrases = [
            "help us protect your account",
            "verify your identity",
            "unusual activity",
        ]
        if any(keyword in page_text or keyword in title for keyword in captcha_keywords):
            return True
        return any(phrase in page_text for phrase in alert_phrases)

    def _is_login_page(self, driver: webdriver.Chrome) -> bool:
        url = (driver.current_url or "").lower()
        if any(domain in url for domain in ("login.live.com", "login.microsoftonline.com")):
            return True
        try:
            driver.find_element(By.CSS_SELECTOR, "input[name='loginfmt']")
            return True
        except Exception:  # noqa: BLE001
            return False

    def ensure_session(self) -> webdriver.Chrome:
        driver = self._get_existing_driver()
        if not driver:
            self.launch_manual_login(auto_open=True)
            raise ManualLoginPending(
                "Outlook window opened automatically. Waiting for you to finish signing in."
            )
        try:
            driver.get(OUTLOOK_INBOX_URL)
            if self._is_login_page(driver):
                raise ManualLoginPending(
                    "Outlook is prompting for a login. Complete the sign-in flow in the open window."
                )
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CSS_SELECTOR, '[role="main"]')))
            if self._detect_captcha(driver):
                raise CaptchaDetected("CAPTCHA detected after loading the Outlook profile")
            log_message("Authenticated to Outlook using the persistent browser profile.")
            return driver
        except CaptchaDetected:
            raise
        except ManualLoginRequired:
            raise
        except Exception as exc:  # noqa: BLE001
            raise ManualLoginRequired(
                "Outlook session not authenticated. Launch the manual login window to refresh the Outlook profile."
            ) from exc

    def fetch_new_emails(
        self,
        driver: webdriver.Chrome,
        registry: ForwardedRegistry,
        folders: Optional[List[Tuple[str, str]]] = None,
    ) -> List[EmailContent]:
        wait = WebDriverWait(driver, 30)
        folders_to_check = folders or OUTLOOK_FOLDERS
        new_emails: List[EmailContent] = []
        for folder_name, folder_url in folders_to_check:
            log_message(f"Scanning folder: {folder_name}")
            navigated = self._open_folder_by_click(driver, folder_name, folder_url)
            if not navigated:
                log_message(
                    f"Falling back to direct navigation for {folder_name} due to missing sidebar element."
                )
                try:
                    driver.get(folder_url)
                except WebDriverException as exc:  # noqa: BLE001
                    log_message(f"Failed to open Outlook folder '{folder_name}': {exc}")
                    continue
            human_delay(1.0, 1.6)
            try:
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div[role="main"]')))
            except TimeoutException:
                log_message(f"{folder_name} appears empty or failed to load.")
                continue
            if random.random() < 0.85:
                self._human_mouse_move(driver)
            human_delay()
            email_rows = driver.find_elements(By.CSS_SELECTOR, 'div[role="option"]')
            for row in email_rows:
                aria_label = row.get_attribute("aria-label") or ""
                class_name = row.get_attribute("class") or ""
                data_is_read = (row.get_attribute("data-isread") or "").lower()
                is_unread = "unread" in class_name.lower() or "unread" in aria_label.lower() or data_is_read in {"false", "0"}
                item_id = row.get_attribute("data-itemid") or row.get_attribute("aria-labelledby") or aria_label
                if not item_id:
                    item_id = str(hash(aria_label + str(time.time())))
                if registry.has(item_id):
                    continue
                if not is_unread:
                    continue
                if not self._safe_click(driver, row):
                    log_message(f"Failed to select email row in {folder_name}. Skipping message.")
                    continue
                human_delay()
                content_wait = WebDriverWait(driver, 12)
                sender_elem = self._first_present(
                    content_wait,
                    [
                        (By.CSS_SELECTOR, 'div[role="main"] [data-test-id="message-header-from"] span'),
                        (By.CSS_SELECTOR, 'div[role="main"] span[title]'),
                        (By.CSS_SELECTOR, 'div[role="main"] [data-test-id="sender-name"]'),
                    ],
                )
                subject_elem = self._first_present(
                    content_wait,
                    [
                        (By.CSS_SELECTOR, 'div[role="main"] h1'),
                        (By.CSS_SELECTOR, 'div[role="main"] [data-test-id="message-subject"]'),
                        (By.CSS_SELECTOR, 'div[role="main"] [role="heading"]'),
                    ],
                )
                body_elem = self._first_present(
                    content_wait,
                    [
                        (By.CSS_SELECTOR, 'div[role="document"]'),
                        (By.CSS_SELECTOR, 'div[aria-label="Message body"]'),
                        (By.CSS_SELECTOR, 'div[data-test-id="message-body-container"]'),
                    ],
                )

                if not subject_elem or not body_elem:
                    log_message(
                        f"Failed to parse email content in {folder_name}: missing subject or body elements."
                    )
                    continue

                sender = sender_elem.text.strip() if sender_elem else "Unknown sender"
                subject = subject_elem.text.strip()
                body_html = body_elem.get_attribute("innerHTML") or ""
                body_text = body_elem.text or ""
                new_emails.append(EmailContent(item_id, sender, subject, body_html, body_text))
                log_message(f"Captured email '{subject}' from {sender} in {folder_name}")
                human_delay(0.5, 1.0)
        return new_emails


@dataclass
class AutomationState:
    registry: ForwardedRegistry
    counter: DailyCounter
    gmail_forwarder: GmailForwarder
    outlook: OutlookAutomation
    settings: SettingsManager
    running: bool = False
    last_run: Optional[str] = None
    cooldown_until: Optional[datetime] = None


AUTOMATION_STATE = AutomationState(
    registry=ForwardedRegistry(FORWARDED_LOG_PATH),
    counter=DailyCounter(FORWARD_STATE_PATH),
    gmail_forwarder=GmailForwarder(SettingsManager()),
    outlook=OutlookAutomation(),
    settings=SettingsManager(),
)


def worker_loop(stop_event: threading.Event, manual_event: threading.Event) -> None:
    registry = AUTOMATION_STATE.registry
    counter = AUTOMATION_STATE.counter
    gmail = AUTOMATION_STATE.gmail_forwarder
    outlook = AUTOMATION_STATE.outlook
    settings = AUTOMATION_STATE.settings
    target_email = settings.get("target_email")
    if not target_email:
        log_message("Target Gmail address not set. Stopping worker.")
        AUTOMATION_STATE.running = False
        return
    if not outlook.profile_ready():
        log_message("Outlook profile not ready. Launch manual login and save the session before starting automation.")
        manual_event.set()
        AUTOMATION_STATE.running = False
        return
    log_message("Automation worker started.")
    while not stop_event.is_set():
        now = datetime.now(timezone.utc)
        if AUTOMATION_STATE.cooldown_until and now < AUTOMATION_STATE.cooldown_until:
            wait_seconds = (AUTOMATION_STATE.cooldown_until - now).total_seconds()
            log_message(f"Cooling down for {int(wait_seconds)} seconds before retrying.")
            stop_event.wait(wait_seconds)
            continue
        driver: Optional[webdriver.Chrome] = None
        try:
            driver = outlook.ensure_session()
            emails = outlook.fetch_new_emails(driver, registry)
            if emails:
                for email in emails:
                    gmail.send_email(target_email, f"FWD: {email.subject}", email.body_html, email.body_text)
                    registry.add(email.message_id)
                    count = counter.increment()
                    log_message(f"Forwarded email #{count} today.")
            else:
                log_message("No new unread emails detected.")
            AUTOMATION_STATE.last_run = datetime.now().astimezone(tz.tzlocal()).strftime("%Y-%m-%d %H:%M:%S %Z")
        except ManualLoginPending as exc:
            manual_event.set()
            log_message(f"Waiting for manual login: {exc}")
            AUTOMATION_STATE.cooldown_until = datetime.now(timezone.utc) + timedelta(seconds=30)
        except ManualLoginRequired as exc:
            manual_event.set()
            log_message(f"Manual login required: {exc}")
            try:
                gmail.send_alert(target_email, str(exc))
            except Exception:
                pass
            AUTOMATION_STATE.cooldown_until = datetime.now(timezone.utc) + timedelta(minutes=30)
        except CaptchaDetected as exc:
            manual_event.set()
            log_message(f"CAPTCHA detected: {exc}")
            try:
                gmail.send_alert(target_email, str(exc))
            except Exception:
                pass
            AUTOMATION_STATE.cooldown_until = datetime.now(timezone.utc) + timedelta(minutes=30)
        except FileNotFoundError as exc:
            log_message(str(exc))
            manual_event.set()
            AUTOMATION_STATE.running = False
            return
        except Exception as exc:  # noqa: BLE001
            log_message(f"Unexpected error: {exc}")
            AUTOMATION_STATE.cooldown_until = datetime.now(timezone.utc) + timedelta(minutes=10)
        if stop_event.is_set():
            break
        min_window = _coerce_minutes(settings.get("polling_min_minutes", 5), 5)
        max_window = _coerce_minutes(settings.get("polling_max_minutes", 10), 10)
        if max_window < min_window:
            max_window = max(min_window, min_window + 1)
        sleep_minutes = random.uniform(min_window, max_window)
        log_message(f"Sleeping for {sleep_minutes:.1f} minutes before next check.")
        stop_event.wait(sleep_minutes * 60)
    AUTOMATION_STATE.running = False
    log_message("Automation worker stopped.")


def run_single_check() -> Tuple[bool, str]:
    settings = AUTOMATION_STATE.settings
    target_email = settings.get("target_email")
    if not target_email:
        return False, "Please save your target Gmail address first."
    outlook = AUTOMATION_STATE.outlook
    if not outlook.profile_ready():
        MANUAL_LOGIN_EVENT.set()
        return False, "Outlook profile not found. Launch the manual login, save the session, and try again."

    registry = AUTOMATION_STATE.registry
    counter = AUTOMATION_STATE.counter
    gmail = AUTOMATION_STATE.gmail_forwarder

    log_message("Manual check triggered from the UI.")
    driver: Optional[webdriver.Chrome] = None
    try:
        driver = outlook.ensure_session()
        emails = outlook.fetch_new_emails(driver, registry)
        if emails:
            for email in emails:
                gmail.send_email(target_email, f"FWD: {email.subject}", email.body_html, email.body_text)
                registry.add(email.message_id)
                counter.increment()
            log_message(f"Manual check forwarded {len(emails)} new email(s).")
            message = f"Manual check complete. Forwarded {len(emails)} new email(s)."
        else:
            log_message("Manual check: no new unread emails detected.")
            message = "Manual check complete. No unread emails detected."
        AUTOMATION_STATE.last_run = datetime.now().astimezone(tz.tzlocal()).strftime("%Y-%m-%d %H:%M:%S %Z")
        AUTOMATION_STATE.cooldown_until = None
        return True, message
    except ManualLoginPending as exc:
        MANUAL_LOGIN_EVENT.set()
        log_message(f"Manual check waiting for login: {exc}")
        return False, "Manual login in progress. Complete the Outlook sign-in window and try again."
    except ManualLoginRequired as exc:
        MANUAL_LOGIN_EVENT.set()
        log_message(f"Manual check requires login: {exc}")
        try:
            gmail.send_alert(target_email, str(exc))
        except Exception:  # noqa: BLE001
            pass
        AUTOMATION_STATE.cooldown_until = datetime.now(timezone.utc) + timedelta(minutes=30)
        return False, f"Manual login required: {exc}"
    except CaptchaDetected as exc:
        MANUAL_LOGIN_EVENT.set()
        log_message(f"Manual check detected CAPTCHA: {exc}")
        AUTOMATION_STATE.cooldown_until = datetime.now(timezone.utc) + timedelta(minutes=30)
        return False, "CAPTCHA detected. Please complete the manual login flow."
    except FileNotFoundError as exc:
        log_message(str(exc))
        return False, str(exc)
    except Exception as exc:  # noqa: BLE001
        log_message(f"Manual check failed: {exc}")
        return False, f"Manual check failed: {exc}"


def send_gmail_test_email(target_email: str, success_message: str) -> None:
    try:
        AUTOMATION_STATE.gmail_forwarder.send_test_email(target_email)
    except FileNotFoundError as exc:
        message = str(exc)
        st.warning(message)
        log_message(message)
    except HttpError as exc:
        message = f"Gmail API error while sending test email: {exc}"
        st.error(message)
        log_message(message)
    except Exception as exc:  # noqa: BLE001
        message = f"Failed to send test email: {exc}"
        st.error(message)
        log_message(message)
    else:
        st.success(success_message)
        log_message(f"Test email sent to {target_email}.")


def handle_open_outlook_and_start() -> None:
    """Single entrypoint for: open Outlook → save session → start automation."""
    global WORKER_THREAD

    settings_manager = AUTOMATION_STATE.settings
    outlook = AUTOMATION_STATE.outlook

    target_email = settings_manager.get("target_email")
    if not target_email:
        st.error("Set your target Gmail address first.")
        return

    profile_ready = outlook.profile_ready()
    driver = MANUAL_DRIVER_HOLDER.get("driver")

    # Phase 1: No profile yet → open Outlook window so user can log in
    if not profile_ready and not driver:
        try:
            outlook.launch_manual_login()
            MANUAL_LOGIN_EVENT.set()
            st.info("Outlook window opened. Sign in fully, then click this button again.")
        except Exception as exc:  # noqa: BLE001
            st.error(f"Could not launch Outlook: {exc}")
        return

    # Phase 2: Driver exists but profile not saved yet → try to persist the session
    if not profile_ready and driver:
        try:
            outlook.complete_manual_login()
            profile_ready = True
            MANUAL_LOGIN_EVENT.clear()
            st.success("Outlook session saved. Starting the forwarder…")
        except Exception as exc:  # noqa: BLE001
            st.error(
                "Outlook inbox not detected yet. Make sure you're fully signed in, "
                "then click this button again."
            )
            return

    # Phase 3: Profile ready → start automation if not already running
    if profile_ready:
        if AUTOMATION_STATE.running:
            st.info("Automation is already running.")
            return

        if not settings_manager.get("target_email"):
            st.error("Set your target Gmail address first.")
            return

        STOP_EVENT.clear()
        MANUAL_LOGIN_EVENT.clear()
        AUTOMATION_STATE.running = True
        AUTOMATION_STATE.cooldown_until = None

        WORKER_THREAD = threading.Thread(
            target=worker_loop,
            args=(STOP_EVENT, MANUAL_LOGIN_EVENT),
            name="OutlookForwarder",
            daemon=True,
        )
        WORKER_THREAD.start()
        st.success("Automation started. New Outlook mail will be forwarded to Gmail.")


# --------------------------------------------------------------------------------------
# Streamlit UI (light theme)
# --------------------------------------------------------------------------------------

st.set_page_config(
    page_title="Outlook ➜ Gmail Forwarder",
    page_icon="📬",
    layout="wide",
)

# --- Global light styling ---
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    html, body, .stApp {
        font-family: 'Inter', system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
        background-color: #f5f5f8;
    }

    [data-testid="stAppViewContainer"] > .main {
        padding-top: 0.5rem;
    }

    div.block-container {
        max-width: 1100px;
        padding-top: 0.75rem;
        padding-bottom: 2.5rem;
    }

    div[data-testid="stHeader"] {
        background: transparent;
    }

    .app-title {
        font-size: 1.9rem;
        font-weight: 700;
        color: #111827;
        margin-bottom: 0.1rem;
        letter-spacing: -0.01em;
    }

    .app-subtitle {
        font-size: 0.95rem;
        color: #4b5563;
        margin-bottom: 0.9rem;
    }

    .card {
        background-color: #ffffff;
        border-radius: 0.9rem;
        padding: 1rem 1.2rem;
        border: 1px solid #e5e7eb;
        box-shadow: 0 8px 20px rgba(15, 23, 42, 0.04);
        margin-bottom: 1rem;
    }

    .section-title {
        font-size: 0.8rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.12em;
        color: #6b7280;
        margin-bottom: 0.4rem;
    }

    .small-note {
        font-size: 0.8rem;
        color: #6b7280;
        margin-top: 0.25rem;
    }

    .status-row {
        display: flex;
        flex-wrap: wrap;
        gap: 0.35rem;
        margin-top: 0.15rem;
        margin-bottom: 0.4rem;
    }

    .status-pill {
        display: inline-flex;
        align-items: center;
        border-radius: 999px;
        padding: 0.12rem 0.6rem;
        font-size: 0.75rem;
        font-weight: 500;
        border: 1px solid #e5e7eb;
        background-color: #f9fafb;
        color: #374151;
    }

    .status-pill.ok {
        border-color: #bbf7d0;
        background-color: #f0fdf4;
        color: #166534;
    }

    .status-pill.warn {
        border-color: #fed7aa;
        background-color: #fffbeb;
        color: #92400e;
    }

    .primary-action button {
        background-color: #2563eb !important;
        color: #ffffff !important;
        border-radius: 0.6rem !important;
        border: 1px solid #1d4ed8 !important;
        font-weight: 600 !important;
        font-size: 0.9rem !important;
    }

    .primary-action button:hover:not(:disabled) {
        background-color: #1d4ed8 !important;
    }

    .secondary-action button {
        background-color: #ffffff !important;
        color: #111827 !important;
        border-radius: 0.6rem !important;
        border: 1px solid #d1d5db !important;
        font-weight: 500 !important;
        font-size: 0.88rem !important;
    }

    .secondary-action button:hover:not(:disabled) {
        background-color: #f3f4f6 !important;
    }

    .danger-action button {
        background-color: #fee2e2 !important;
        color: #b91c1c !important;
        border-radius: 0.6rem !important;
        border: 1px solid #fecaca !important;
        font-weight: 500 !important;
        font-size: 0.88rem !important;
    }

    .danger-action button:hover:not(:disabled) {
        background-color: #fecaca !important;
    }

    div[data-baseweb="input"] input {
        border-radius: 0.55rem !important;
        border: 1px solid #d1d5db !important;
        padding: 0.45rem 0.75rem !important;
        font-size: 0.9rem !important;
    }

    div[data-baseweb="input"] input:focus {
        border-color: #2563eb !important;
        box-shadow: 0 0 0 1px rgba(37, 99, 235, 0.25);
    }

    div[data-baseweb="slider"] {
        padding-top: 0.15rem;
    }

    pre, code {
        border-radius: 0.5rem !important;
    }

    .stExpander {
        border-radius: 0.7rem !important;
        border: 1px solid #e5e7eb !important;
        background-color: #ffffff !important;
    }

    a, .stMarkdown a {
        color: #2563eb;
        text-decoration: none;
        font-weight: 500;
    }

    a:hover, .stMarkdown a:hover {
        text-decoration: underline;
    }

    /* Metric cards */
    div[data-testid="stMetric"] {
        background-color: #f9fafb;
        border-radius: 0.7rem;
        border: 1px solid #e5e7eb;
        padding: 0.4rem 0.7rem;
    }

    div[data-testid="stMetricLabel"] > div {
        font-size: 0.78rem;
        color: #6b7280;
    }

    div[data-testid="stMetricValue"] > div {
        font-size: 1.25rem;
        color: #111827;
        font-weight: 600;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# --- Tab focus → optional live refresh of logs ---
focus_state = components.html(
    """
    <script>
    const sendFocusState = () => {
        const isFocused = document.visibilityState === 'visible';
        Streamlit.setComponentValue(isFocused);
    };
    document.addEventListener('visibilitychange', sendFocusState);
    window.addEventListener('focus', sendFocusState);
    window.addEventListener('blur', sendFocusState);
    setInterval(sendFocusState, 5000);
    sendFocusState();
    </script>
    """,
    height=0,
)
is_focused = True if focus_state is None else bool(focus_state)
st.session_state["tab_focused"] = is_focused

# --- Title ---
st.markdown("<div class='app-title'>📬 Outlook ➜ Gmail Forwarder</div>", unsafe_allow_html=True)
st.markdown(
    "<div class='app-subtitle'>Forward new Outlook messages into Gmail while Outlook runs in the background.</div>",
    unsafe_allow_html=True,
)

with st.expander("How to set this up", expanded=False):
    st.markdown(
        """
        1. Install the Python packages listed at the top of the script.  
        2. Place your Gmail API `credentials.json` next to this file.  
        3. Enter the Gmail address that should receive forwarded mail.  
        4. Click **Open Outlook & Start** and follow the prompts in the Outlook window.
        """
    )

settings_manager = AUTOMATION_STATE.settings

# --- Load persisted settings ---
target_email = settings_manager.get("target_email", "") or ""
polling_min_saved = _coerce_minutes(settings_manager.get("polling_min_minutes", 5), 5)
polling_max_saved = _coerce_minutes(settings_manager.get("polling_max_minutes", 10), 10)
if polling_max_saved < polling_min_saved:
    polling_max_saved = max(polling_min_saved, polling_min_saved + 1)

profile_ready = AUTOMATION_STATE.outlook.profile_ready()
running = AUTOMATION_STATE.running

# === Top section: Gmail + status ===
with st.container():
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    top_cols = st.columns([1.6, 1.1])

    # Left: Gmail + cadence
    with top_cols[0]:
        st.markdown("<div class='section-title'>Step 1 · Where to send mail</div>", unsafe_allow_html=True)

        updated_email = st.text_input(
            "Gmail address to receive forwarded messages",
            value=target_email,
            placeholder="you@example.com",
        )

        normalized_saved = (target_email or "").strip().lower()
        normalized_input = (updated_email or "").strip()
        save_disabled = normalized_input.lower() == normalized_saved

        btn_row = st.columns([1, 1])
        with btn_row[0]:
            st.markdown("<div class='secondary-action'>", unsafe_allow_html=True)
            if st.button("Save Gmail address", disabled=save_disabled, use_container_width=True):
                if normalized_input:
                    settings_manager.set("target_email", normalized_input)
                    target_email = normalized_input
                    st.success("Gmail address saved.")
                else:
                    st.error("Enter a valid Gmail address first.")
            st.markdown("</div>", unsafe_allow_html=True)

        with btn_row[1]:
            st.markdown("<div class='secondary-action'>", unsafe_allow_html=True)
            if st.button("Send test email", use_container_width=True):
                if normalized_input:
                    settings_manager.set("target_email", normalized_input)
                    target_email = normalized_input
                    send_gmail_test_email(
                        normalized_input,
                        "Verification email sent. Check your Gmail inbox.",
                    )
                else:
                    st.error("Enter a Gmail address before sending a test.")
            st.markdown("</div>", unsafe_allow_html=True)

        st.markdown(
            "<div class='small-note'>Tip: use a dedicated Gmail account for this forwarder.</div>",
            unsafe_allow_html=True,
        )

        st.markdown("<div class='section-title' style='margin-top:0.7rem;'>Scan interval</div>", unsafe_allow_html=True)
        polling_min, polling_max = st.slider(
            "Minutes between Outlook checks (random within this range)",
            min_value=1,
            max_value=30,
            value=(polling_min_saved, polling_max_saved),
            help="The worker sleeps a random duration in this range between checks.",
        )
        if (polling_min, polling_max) != (polling_min_saved, polling_max_saved):
            settings_manager.set("polling_min_minutes", polling_min)
            settings_manager.set("polling_max_minutes", polling_max)
            note = "Updated polling interval saved."
        else:
            note = f"Currently pausing between {polling_min} and {polling_max} minutes."
        st.markdown(f"<div class='small-note'>{note}</div>", unsafe_allow_html=True)

    # Right: status / metrics
    with top_cols[1]:
        st.markdown("<div class='section-title'>Status</div>", unsafe_allow_html=True)

        st.markdown("<div class='status-row'>", unsafe_allow_html=True)
        gmail_ok = bool(target_email)
        outlook_ok = profile_ready

        gmail_class = "ok" if gmail_ok else "warn"
        outlook_class = "ok" if outlook_ok else "warn"
        running_class = "ok" if running else ""

        st.markdown(
            f"<span class='status-pill {gmail_class}'>Gmail: {'set' if gmail_ok else 'not set'}</span>",
            unsafe_allow_html=True,
        )
        st.markdown(
            f"<span class='status-pill {outlook_class}'>Outlook profile: {'ready' if outlook_ok else 'not ready'}</span>",
            unsafe_allow_html=True,
        )
        st.markdown(
            f"<span class='status-pill {running_class}'>Automation: {'running' if running else 'stopped'}</span>",
            unsafe_allow_html=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)

        metric_cols = st.columns(3)
        with metric_cols[0]:
            st.metric("Forwarded today", AUTOMATION_STATE.counter.get_count())
        with metric_cols[1]:
            st.metric("Last run", AUTOMATION_STATE.last_run or "N/A")
        with metric_cols[2]:
            cooldown_value = (
                AUTOMATION_STATE.cooldown_until.astimezone(tz.tzlocal()).strftime("%Y-%m-%d %H:%M:%S %Z")
                if AUTOMATION_STATE.cooldown_until
                else "Ready"
            )
            st.metric("Cooldown until", cooldown_value)

    st.markdown("</div>", unsafe_allow_html=True)

# === Automation controls ===
with st.container():
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.markdown("<div class='section-title'>Step 2 · Start forwarding</div>", unsafe_allow_html=True)

    ctrl_cols = st.columns([1.4, 0.9, 0.9])

    # Main CTA
    with ctrl_cols[0]:
        st.markdown("<div class='primary-action'>", unsafe_allow_html=True)
        if st.button("Open Outlook & Start", disabled=running, use_container_width=True):
            handle_open_outlook_and_start()
        st.markdown("</div>", unsafe_allow_html=True)

    # Stop
    with ctrl_cols[1]:
        st.markdown("<div class='danger-action'>", unsafe_allow_html=True)
        if st.button("Stop automation", disabled=not running, use_container_width=True):
            STOP_EVENT.set()
            if WORKER_THREAD and WORKER_THREAD.is_alive():
                WORKER_THREAD.join(timeout=2)
            AUTOMATION_STATE.running = False
            st.info("Automation stop requested.")
        st.markdown("</div>", unsafe_allow_html=True)

    # One-off check
    with ctrl_cols[2]:
        st.markdown("<div class='secondary-action'>", unsafe_allow_html=True)
        if st.button("Run one check", disabled=running, use_container_width=True):
            if not AUTOMATION_STATE.outlook.profile_ready():
                MANUAL_LOGIN_EVENT.set()
                st.error("Outlook profile not found. Use “Open Outlook & Start” first.")
            else:
                success, message = run_single_check()
                if success:
                    st.success(message)
                else:
                    st.error(message)
        st.markdown("</div>", unsafe_allow_html=True)

    # Helper text about manual login / profile state
    if MANUAL_LOGIN_EVENT.is_set():
        st.warning(
            "Manual Outlook login required. When the Outlook window opens, sign in fully, "
            "then click “Open Outlook & Start” again."
        )
    elif not profile_ready:
        st.markdown(
            "<div class='small-note'>On first run, the button opens Outlook so you can sign in. "
            "After saving the session, automation can run in the background.</div>",
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            "<div class='small-note'>Outlook profile is ready. Use the buttons above to start, stop, "
            "or run a single check.</div>",
            unsafe_allow_html=True,
        )

    st.markdown("</div>", unsafe_allow_html=True)

# === Activity log ===
with st.container():
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    header_cols = st.columns([1.3, 0.8, 0.7])

    with header_cols[0]:
        st.markdown("<div class='section-title'>Activity log</div>", unsafe_allow_html=True)

    with header_cols[1]:
        live_refresh = st.checkbox(
            "Auto-refresh when tab is focused",
            value=st.session_state.get("tab_focused", True) and running,
            help="If enabled, the log view auto-refreshes while this tab is active and automation is running.",
        )
        st.session_state["live_refresh"] = live_refresh

    with header_cols[2]:
        st.markdown("<div class='secondary-action'>", unsafe_allow_html=True)
        if st.button("Refresh log", use_container_width=True):
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    log_container = st.empty()
    with LOG_LOCK:
        log_lines = list(LOG_BUFFER)
    log_container.code("\n".join(log_lines[-200:]) or "No logs yet.", language="text")

    st.markdown("</div>", unsafe_allow_html=True)

# Auto-refresh ping for logs when running & live-refresh is on
if (
    AUTOMATION_STATE.running
    and st.session_state.get("live_refresh")
    and st.session_state.get("tab_focused", True)
):
    components.html(
        """
        <script>
        setTimeout(() => {
            Streamlit.setComponentValue(Date.now());
        }, 5000);
        </script>
        """,
        height=0,
    )

st.markdown("[View the project on GitHub](https://github.com/Skytheredhead/outlookscrape)")
