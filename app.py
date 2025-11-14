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
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
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
if "MANUAL_DRIVER_HOLDER" not in globals():
    MANUAL_DRIVER_HOLDER: Dict[str, Optional[webdriver.Chrome]] = {"driver": None}

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
            raise RuntimeError(
                "Unable to locate a Chrome or Chromium browser binary. Install Chrome/Chromium or set the CHROME_BINARY "
                "environment variable to the browser executable before launching the manual login."
            ) from exc
        driver.set_window_size(1400, 900)
        return driver

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
            raise ManualLoginRequired(
                "Outlook window opened automatically. Complete the login flow and click 'Save session' before starting automation."
            )
        try:
            driver.get(OUTLOOK_INBOX_URL)
            if self._is_login_page(driver):
                raise ManualLoginRequired(
                    "Outlook asked for login again. Launch the manual login window to refresh the saved session."
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
            try:
                driver.get(folder_url)
            except WebDriverException as exc:  # noqa: BLE001
                log_message(f"Failed to open Outlook folder '{folder_name}': {exc}")
                continue
            human_delay(1.0, 1.6)
            tail = folder_url.rsplit("/", 1)[-1].lower()
            current_url = (driver.current_url or "").lower()
            if tail not in current_url:
                try:
                    nav_button = driver.find_element(
                        By.XPATH,
                        "//span[normalize-space()='{0}']/ancestor::button[1] | "
                        "//span[contains(@title, '{0}')]/ancestor::button[1]".format(folder_name),
                    )
                    driver.execute_script("arguments[0].click();", nav_button)
                    human_delay(0.8, 1.2)
                except Exception as exc:  # noqa: BLE001
                    log_message(f"Unable to switch to folder '{folder_name}' via navigation: {exc}")
            try:
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div[role="option"]')))
            except TimeoutException:
                log_message(f"{folder_name} appears empty or failed to load.")
                continue
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
                try:
                    ActionChains(driver).move_to_element(row).pause(random.uniform(0.5, 1.2)).click().perform()
                except WebDriverException as exc:  # noqa: BLE001
                    log_message(f"Failed to select email row in {folder_name}: {exc}")
                    continue
                human_delay()
                try:
                    sender_elem = wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'div[role="main"] span[title]'))
                    )
                    subject_elem = driver.find_element(By.CSS_SELECTOR, 'div[role="main"] h1')
                    body_container = driver.find_element(By.CSS_SELECTOR, 'div[role="document"]')
                    sender = sender_elem.text.strip()
                    subject = subject_elem.text.strip()
                    body_html = body_container.get_attribute("innerHTML")
                    body_text = body_container.text
                    new_emails.append(EmailContent(item_id, sender, subject, body_html, body_text))
                    log_message(f"Captured email '{subject}' from {sender} in {folder_name}")
                except NoSuchElementException as exc:
                    log_message(f"Failed to parse email content in {folder_name}: {exc}")
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


# --------------------------------------------------------------------------------------
# Streamlit UI
# --------------------------------------------------------------------------------------
st.set_page_config(
    page_title="Outlook ➜ Gmail Forwarder",
    page_icon="📬",
    layout="wide",
)

st.markdown(
    """
    <style>
    :root {
        color-scheme: dark;
        --accent: #38bdf8;
        --accent-strong: #6366f1;
    }
    div[data-testid="stAppViewContainer"] > .main {
        background: radial-gradient(circle at top, rgba(99, 102, 241, 0.2), transparent 60%), #0b1120;
        color: #e2e8f0;
        padding-top: 1rem;
    }
    div.block-container {
        padding-top: 1rem;
        padding-bottom: 2.4rem;
        max-width: 1120px;
    }
    div[data-testid="stHeader"] {background: transparent;}
    .page-title {
        font-size: 1.9rem;
        font-weight: 700;
        color: #f8fafc;
        margin-bottom: 0.2rem;
    }
    .page-subtitle {
        color: #94a3b8;
        font-size: 0.95rem;
        margin-bottom: 1.1rem;
    }
    .card {
        background: rgba(15, 23, 42, 0.88);
        border-radius: 0.9rem;
        padding: 1rem 1.2rem;
        box-shadow: 0 18px 40px rgba(2, 6, 23, 0.45);
        margin-bottom: 1rem;
        border: 1px solid rgba(148, 163, 184, 0.12);
    }
    .section-title {
        font-size: 0.92rem;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        color: #bae6fd;
        margin-bottom: 0.55rem;
    }
    .stButton>button {
        background: linear-gradient(135deg, var(--accent) 0%, var(--accent-strong) 100%);
        color: #0b1120;
        border-radius: 0.65rem;
        border: none;
        font-weight: 600;
        font-size: 0.9rem;
        padding: 0.45rem 0.85rem;
        transition: transform 0.15s ease, filter 0.15s ease;
        box-shadow: 0 10px 24px rgba(56, 189, 248, 0.35);
    }
    .stButton>button:disabled {
        background: rgba(30, 41, 59, 0.85);
        color: #64748b;
        box-shadow: none;
    }
    .stButton>button:not(:disabled):hover {
        transform: translateY(-1px);
        filter: brightness(1.05);
    }
    .stButton>button:not(:disabled):active {
        transform: translateY(0);
    }
    div[data-baseweb="input"] input {
        background: rgba(15, 23, 42, 0.65);
        border-radius: 0.65rem !important;
        border: 1px solid rgba(148, 163, 184, 0.28);
        color: #e2e8f0 !important;
        padding: 0.48rem 0.85rem !important;
    }
    div[data-baseweb="input"] input::placeholder {
        color: #64748b;
    }
    div[data-baseweb="slider"] {
        padding: 0.3rem 0.35rem 0.1rem;
    }
    .small-note {
        font-size: 0.78rem;
        color: #94a3b8;
        margin-top: 0.3rem;
    }
    .pill {
        display: inline-flex;
        align-items: center;
        gap: 0.35rem;
        background: rgba(56, 189, 248, 0.18);
        color: #e0f2fe;
        border-radius: 999px;
        padding: 0.18rem 0.75rem;
        font-size: 0.75rem;
        font-weight: 600;
        letter-spacing: 0.03em;
    }
    div[data-testid="stMetric"] {
        background: rgba(15, 23, 42, 0.78);
        border-radius: 0.75rem;
        border: 1px solid rgba(148, 163, 184, 0.18);
    }
    div[data-testid="stMetricLabel"] > div {
        color: #94a3b8;
        font-size: 0.75rem;
        letter-spacing: 0.08em;
    }
    div[data-testid="stMetricValue"] > div {
        color: #f8fafc;
        font-size: 1.45rem;
        font-weight: 700;
    }
    .stCheckbox>label {
        color: #cbd5f5;
        font-size: 0.84rem;
    }
    pre, code {
        background-color: rgba(15, 23, 42, 0.85) !important;
        color: #e2e8f0 !important;
    }
    a, .stMarkdown a {
        color: #38bdf8;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown("<div class='page-title'>📬 Outlook ➜ Gmail Forwarder</div>", unsafe_allow_html=True)
st.markdown(
    "<div class='page-subtitle'>A compact command center for relaying Outlook messages into Gmail on your terms.</div>",
    unsafe_allow_html=True,
)

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
if focus_state is None:
    is_focused = True
else:
    is_focused = bool(focus_state)
st.session_state["tab_focused"] = is_focused

with st.expander("Setup checklist", expanded=False):
    st.markdown(
        """
        1. Install the dependencies listed at the top of this script.
        2. Download your `credentials.json` (or the `client_secret_*.apps.googleusercontent.com.json` file) from the [Google Cloud Console](https://console.cloud.google.com/apis/credentials) for the Gmail API and place it alongside this script.
        3. Run `streamlit run app.py` (or double-click `run_app.bat` on Windows).
        4. Enter the Gmail address that should receive forwarded messages and save it.
        5. Click **Login to Outlook** to open Chrome (non-headless), sign in completely, then click **Save session** to persist the profile and keep the window open.
        6. Press **Start scanning** (or **Run a check**) once the status indicators show that everything is ready.
        """
    )

settings_manager = AUTOMATION_STATE.settings

if "auto_outlook_open_attempted" not in st.session_state:
    st.session_state["auto_outlook_open_attempted"] = False

profile_ready = AUTOMATION_STATE.outlook.profile_ready()
if profile_ready and not st.session_state["auto_outlook_open_attempted"]:
    st.session_state["auto_outlook_open_attempted"] = True
    try:
        AUTOMATION_STATE.outlook.launch_manual_login(auto_open=True)
    except Exception as exc:  # noqa: BLE001
        st.warning(f"Unable to auto-open Outlook window: {exc}")
elif not profile_ready:
    st.session_state["auto_outlook_open_attempted"] = False

target_email = settings_manager.get("target_email", "") or ""
polling_min_saved = _coerce_minutes(settings_manager.get("polling_min_minutes", 5), 5)
polling_max_saved = _coerce_minutes(settings_manager.get("polling_max_minutes", 10), 10)
if polling_max_saved < polling_min_saved:
    polling_max_saved = max(polling_min_saved, polling_min_saved + 1)

with st.container():
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.markdown("<div class='section-title'>Notifications & cadence</div>", unsafe_allow_html=True)
    notif_col, cadence_col = st.columns([1.35, 1])
    with notif_col:
        updated_email = st.text_input(
            "Main Gmail address",
            value=target_email,
            placeholder="you@example.com",
        )
        normalized_saved = (target_email or "").strip().lower()
        normalized_input = (updated_email or "").strip()
        save_disabled = normalized_input.lower() == normalized_saved
        button_cols = st.columns(2)
        with button_cols[0]:
            if st.button(
                "Save Gmail address",
                disabled=save_disabled,
            ):
                if normalized_input:
                    settings_manager.set("target_email", normalized_input)
                    target_email = normalized_input
                    st.success("Saved! Use the test button to confirm Gmail delivery when ready.")
                else:
                    st.error("Please provide a valid Gmail address.")
        with button_cols[1]:
            if st.button("Send test email"):
                if normalized_input:
                    settings_manager.set("target_email", normalized_input)
                    target_email = normalized_input
                    send_gmail_test_email(
                        normalized_input,
                        "Sent a verification email. Check your Gmail inbox to confirm delivery.",
                    )
                else:
                    st.error("Please provide a valid Gmail address before sending a test.")
    with cadence_col:
        st.markdown("<span class='pill'>⏱️ Polling window</span>", unsafe_allow_html=True)
        polling_min, polling_max = st.slider(
            "Choose how often to scan (minutes)",
            min_value=1,
            max_value=30,
            value=(polling_min_saved, polling_max_saved),
            help="The worker waits a random duration within this range before each Outlook scan.",
        )
        poll_message = f"Currently pausing between {polling_min} and {polling_max} minutes."
        if (polling_min, polling_max) != (polling_min_saved, polling_max_saved):
            settings_manager.set("polling_min_minutes", polling_min)
            settings_manager.set("polling_max_minutes", polling_max)
            poll_message = "Updated polling cadence saved instantly."
        st.markdown(f"<div class='small-note'>{poll_message}</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

profile_ready = AUTOMATION_STATE.outlook.profile_ready()

with st.container():
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.markdown("<div class='section-title'>Outlook session</div>", unsafe_allow_html=True)
    status_label = "✅ Profile ready" if profile_ready else "⚠️ Profile missing"
    st.markdown(
        f"<span class='pill'>{status_label}</span>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<div class='small-note'>Launch the dedicated Chrome window whenever Microsoft asks you to verify the session. Keep it open and signed in.</div>",
        unsafe_allow_html=True,
    )
    manual_cols = st.columns(2)
    with manual_cols[0]:
        if st.button("Login to Outlook"):
            try:
                AUTOMATION_STATE.outlook.launch_manual_login()
            except Exception as exc:  # noqa: BLE001
                st.error(f"Failed to launch manual login: {exc}")
    with manual_cols[1]:
        if st.button("Save session"):
            try:
                AUTOMATION_STATE.outlook.complete_manual_login()
                st.success("Session saved. Leave the Outlook window open so scans can reuse it instantly.")
            except Exception as exc:  # noqa: BLE001
                st.error(f"Failed to persist profile: {exc}")
    if MANUAL_LOGIN_EVENT.is_set():
        st.error("Manual login required. Launch the Outlook window, sign in, and save the session.")
    st.markdown("</div>", unsafe_allow_html=True)

with st.container():
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.markdown("<div class='section-title'>Automation status</div>", unsafe_allow_html=True)
    running = AUTOMATION_STATE.running
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

    control_cols = st.columns([1.1, 1, 1, 1])
    with control_cols[0]:
        if st.button("▶️ Start scanning", disabled=AUTOMATION_STATE.running):
            if not settings_manager.get("target_email"):
                st.error("Please save your target Gmail address first.")
            elif not AUTOMATION_STATE.outlook.profile_ready():
                MANUAL_LOGIN_EVENT.set()
                st.error("Outlook profile not found. Launch the manual login window and save the session before starting.")
            else:
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
                st.success("Automation started. Logs will appear below.")
    with control_cols[1]:
        if st.button("⏹️ Stop", disabled=not AUTOMATION_STATE.running):
            STOP_EVENT.set()
            if WORKER_THREAD and WORKER_THREAD.is_alive():
                WORKER_THREAD.join(timeout=2)
            AUTOMATION_STATE.running = False
            st.info("Automation stop requested.")
    with control_cols[2]:
        if st.button("🔍 Run a check", disabled=AUTOMATION_STATE.running):
            if not AUTOMATION_STATE.outlook.profile_ready():
                MANUAL_LOGIN_EVENT.set()
                st.error("Outlook profile not found. Launch the manual login window and save the session before running a check.")
            else:
                success, message = run_single_check()
                if success:
                    st.success(message)
                else:
                    st.error(message)
    with control_cols[3]:
        live_refresh = st.checkbox(
            "Live refresh",
            value=st.session_state.get("tab_focused", True) and running,
            help="When enabled, the log view auto-refreshes whenever this tab stays in focus.",
        )
        st.session_state["live_refresh"] = live_refresh
    st.markdown("</div>", unsafe_allow_html=True)

st.markdown("---")

log_container = st.empty()
with LOG_LOCK:
    log_lines = list(LOG_BUFFER)
log_container.code("\n".join(log_lines[-200:]) or "No logs yet.", language="text")

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

if st.button("Refresh logs", help="Manual refresh to keep resource usage low when unfocused."):
    st.rerun()

st.markdown("[View the project on GitHub](https://github.com/Skytheredhead/outlookscrape)")

