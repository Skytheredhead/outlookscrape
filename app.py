"""
Outlook to Gmail Forwarder

Installation requirements (run once):
    pip install selenium webdriver-manager google-api-python-client google-auth-httplib2 \
                google-auth-oauthlib cryptography streamlit python-dateutil

Run the Streamlit UI:
    streamlit run app.py

Test on a non-production (dummy) account before using with your primary accounts.
"""
import base64
import json
import os
import pickle
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
from typing import Deque, Dict, List, Optional, Tuple

import streamlit as st
import streamlit.components.v1 as components
from cryptography.fernet import Fernet
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

COOKIES_PATH = DATA_DIR / "cookies.pkl"
FERNET_KEY_PATH = DATA_DIR / "fernet.key"
ENCRYPTED_CREDENTIALS_PATH = DATA_DIR / "outlook.enc"
FORWARDED_LOG_PATH = DATA_DIR / "forwarded.json"
FORWARD_STATE_PATH = DATA_DIR / "daily_counter.json"
SETTINGS_PATH = DATA_DIR / "settings.json"

OUTLOOK_LOGIN_URL = "https://outlook.office.com/mail/"
OUTLOOK_INBOX_URL = "https://outlook.office.com/mail/inbox"
OUTLOOK_JUNK_URL = "https://outlook.office.com/mail/junkemail"
OUTLOOK_FOLDERS = [
    ("Inbox", OUTLOOK_INBOX_URL),
    ("Junk Email", OUTLOOK_JUNK_URL),
]

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

MANUAL_LOGIN_EVENT = threading.Event()
STOP_EVENT = threading.Event()
WORKER_THREAD: Optional[threading.Thread] = None
MANUAL_DRIVER_HOLDER: Dict[str, Optional[webdriver.Chrome]] = {"driver": None}

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


class ManualLoginRequired(Exception):
    """Raised when a manual login is required."""


class CaptchaDetected(Exception):
    """Raised when Outlook displays a CAPTCHA or block."""


class CredentialManager:
    """Manage encrypted Outlook credentials."""

    def __init__(self) -> None:
        self.key_path = FERNET_KEY_PATH
        self.secret_path = ENCRYPTED_CREDENTIALS_PATH

    def _get_cipher(self) -> Fernet:
        if not self.key_path.exists():
            key = Fernet.generate_key()
            self.key_path.write_bytes(key)
        else:
            key = self.key_path.read_bytes()
        return Fernet(key)

    def save_credentials(self, username: str, password: str) -> None:
        payload = json.dumps({"username": username, "password": password}).encode("utf-8")
        cipher = self._get_cipher()
        token = cipher.encrypt(payload)
        self.secret_path.write_bytes(token)
        log_message("Encrypted Outlook credentials saved.")

    def load_credentials(self) -> Optional[Dict[str, str]]:
        if not self.secret_path.exists():
            return None
        cipher = self._get_cipher()
        try:
            payload = cipher.decrypt(self.secret_path.read_bytes())
            data = json.loads(payload.decode("utf-8"))
            return data
        except Exception as exc:  # noqa: BLE001
            log_message(f"Failed to decrypt Outlook credentials: {exc}")
            return None


class SettingsManager:
    """Persist lightweight settings (Gmail target address, polling interval, etc.)."""

    def __init__(self) -> None:
        self.path = SETTINGS_PATH
        self._settings: Dict[str, str] = {}
        if self.path.exists():
            try:
                self._settings = json.loads(self.path.read_text(encoding="utf-8"))
            except json.JSONDecodeError:
                self._settings = {}

    def get(self, key: str, default: Optional[str] = None) -> Optional[str]:
        return self._settings.get(key, default)

    def set(self, key: str, value: str) -> None:
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

    def _build_service(self):
        creds = None
        if Path("token.json").exists():
            creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                # The credentials.json file must be downloaded from Google Cloud console.
                flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
                creds = flow.run_local_server(port=0)
            with open("token.json", "w", encoding="utf-8") as token:
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


class OutlookAutomation:
    """Selenium automation for Outlook Web."""

    def __init__(self, credential_manager: CredentialManager):
        self.credential_manager = credential_manager

    @staticmethod
    def _create_driver(headless: bool = True) -> webdriver.Chrome:
        options = Options()
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--disable-extensions")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--lang=en-US")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
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
        binary_location = next((candidate for candidate in binary_candidates if candidate), None)
        if binary_location:
            options.binary_location = binary_location
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

    def _apply_cookies(self, driver: webdriver.Chrome) -> None:
        if not COOKIES_PATH.exists():
            raise ManualLoginRequired("Cookie file not found")
        with COOKIES_PATH.open("rb") as file:
            cookies = pickle.load(file)
        driver.get("https://outlook.office.com")
        for cookie in cookies:
            cookie_dict = {k: v for k, v in cookie.items() if k in {"name", "value", "domain", "path", "expiry", "secure", "httpOnly"}}
            try:
                driver.add_cookie(cookie_dict)
            except Exception:  # noqa: BLE001
                continue
        driver.get(OUTLOOK_INBOX_URL)

    def save_cookies(self, driver: webdriver.Chrome) -> None:
        cookies = driver.get_cookies()
        with COOKIES_PATH.open("wb") as file:
            pickle.dump(cookies, file)
        log_message("Saved Outlook cookies to disk.")

    def launch_manual_login(self) -> None:
        if MANUAL_DRIVER_HOLDER["driver"]:
            log_message("Manual login window already open.")
            return
        driver = self._create_driver(headless=False)
        MANUAL_DRIVER_HOLDER["driver"] = driver
        log_message("Manual Chrome window launched. Complete login and then click 'Save cookies'.")
        driver.get(OUTLOOK_LOGIN_URL)

    def complete_manual_login(self) -> None:
        driver = MANUAL_DRIVER_HOLDER.get("driver")
        if not driver:
            log_message("No manual Chrome session is active.")
            return
        try:
            self.save_cookies(driver)
        finally:
            try:
                driver.quit()
            except Exception:  # noqa: BLE001
                pass
            MANUAL_DRIVER_HOLDER["driver"] = None
        MANUAL_LOGIN_EVENT.clear()

    def _detect_captcha(self, driver: webdriver.Chrome) -> bool:
        page_text = driver.page_source.lower()
        keywords = ["captcha", "verify", "identity", "stay signed in", "blocked"]
        return any(keyword in page_text for keyword in keywords)

    def ensure_session(self) -> webdriver.Chrome:
        driver = self._create_driver(headless=True)
        try:
            self._apply_cookies(driver)
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CSS_SELECTOR, '[role="main"]')))
            if self._detect_captcha(driver):
                raise CaptchaDetected("CAPTCHA detected after applying cookies")
            log_message("Authenticated to Outlook using saved cookies.")
            return driver
        except CaptchaDetected:
            driver.quit()
            raise
        except ManualLoginRequired:
            driver.quit()
            raise
        except Exception as exc:  # noqa: BLE001
            driver.quit()
            raise ManualLoginRequired(
                "Outlook session not authenticated. Launch the manual login window to refresh cookies."
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
            try:
                driver.get(folder_url)
            except WebDriverException as exc:  # noqa: BLE001
                log_message(f"Failed to open Outlook folder '{folder_name}': {exc}")
                continue
            human_delay(1.0, 2.0)
            try:
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div[role="option"]')))
            except TimeoutException:
                log_message(f"{folder_name} appears empty or failed to load.")
                continue
            human_delay()
            email_rows = driver.find_elements(By.CSS_SELECTOR, 'div[role="option"]')
            for row in email_rows:
                aria_label = row.get_attribute("aria-label") or ""
                is_unread = "unread" in (row.get_attribute("class") or "") or "unread" in aria_label.lower()
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
    outlook=OutlookAutomation(CredentialManager()),
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
        except Exception as exc:  # noqa: BLE001
            log_message(f"Unexpected error: {exc}")
            AUTOMATION_STATE.cooldown_until = datetime.now(timezone.utc) + timedelta(minutes=10)
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:  # noqa: BLE001
                    pass
        if stop_event.is_set():
            break
        sleep_minutes = random.uniform(5, 10)
        log_message(f"Sleeping for {sleep_minutes:.1f} minutes before next check.")
        stop_event.wait(sleep_minutes * 60)
    AUTOMATION_STATE.running = False
    log_message("Automation worker stopped.")


# --------------------------------------------------------------------------------------
# Streamlit UI
# --------------------------------------------------------------------------------------
st.set_page_config(
    page_title="Outlook ➜ Gmail Forwarder",
    page_icon="📬",
    layout="wide",
)

st.title("📬 Outlook ➜ Gmail Forwarder")
st.caption(
    "Automate copying new Outlook emails into Gmail. The UI is optimized to stay idle when unfocused."
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
        2. Download your `credentials.json` from the [Google Cloud Console](https://console.cloud.google.com/apis/credentials) for the Gmail API.
        3. Run `streamlit run app.py`.
        4. Provide your Outlook credentials below (encrypted with Fernet).
        5. Click **Launch Manual Login** to open Chrome (non-headless), sign in, then click **Save Cookies**.
        6. Click **Start scanning** to begin the background watcher.
        """
    )

settings_manager = AUTOMATION_STATE.settings
cred_manager = AUTOMATION_STATE.outlook.credential_manager

col1, col2 = st.columns(2)
with col1:
    target_email = st.text_input(
        "Main Gmail address to receive copies",
        value=settings_manager.get("target_email", ""),
        placeholder="you@example.com",
    )
    if st.button("Save Gmail address"):
        if target_email:
            settings_manager.set("target_email", target_email)
            st.success("Saved target Gmail address.")
        else:
            st.error("Please provide a valid Gmail address.")

with col2:
    st.markdown("**Polling interval**: 5-10 minutes (randomized). Human-like delays applied to scraping.")
    st.markdown("**Cookies** stored at: `automation_state/cookies.pkl`")

st.markdown("---")

credentials = cred_manager.load_credentials()
if credentials:
    st.success("Encrypted Outlook credentials are stored securely. Use the button below to update if needed.")
else:
    st.warning("Outlook credentials not saved yet. Enter them below and click save.")

with st.form("outlook_credentials_form", clear_on_submit=False):
    username = st.text_input("Outlook email", value=credentials.get("username") if credentials else "")
    password = st.text_input("Outlook password", type="password")
    submitted = st.form_submit_button("Encrypt & Save Outlook credentials")
    if submitted:
        if username and password:
            cred_manager.save_credentials(username, password)
            st.success("Encrypted credentials saved.")
        else:
            st.error("Both fields are required.")

col_manual1, col_manual2 = st.columns(2)
with col_manual1:
    if st.button("Launch Manual Login (Chrome)"):
        try:
            AUTOMATION_STATE.outlook.launch_manual_login()
        except Exception as exc:  # noqa: BLE001
            st.error(f"Failed to launch manual login: {exc}")
with col_manual2:
    if st.button("Save Cookies & Close Browser"):
        try:
            AUTOMATION_STATE.outlook.complete_manual_login()
            st.success("Cookies saved. You can now run in headless mode.")
        except Exception as exc:  # noqa: BLE001
            st.error(f"Failed to save cookies: {exc}")

if MANUAL_LOGIN_EVENT.is_set():
    st.error("Manual login required. Launch the manual window, sign in, and save cookies.")

st.markdown("---")

status_col, stats_col = st.columns([2, 1])
with status_col:
    running = AUTOMATION_STATE.running
    st.metric("Forwarded today", AUTOMATION_STATE.counter.get_count())
    st.metric("Last run", AUTOMATION_STATE.last_run or "N/A")
    st.metric("Cooldown until", AUTOMATION_STATE.cooldown_until.astimezone(tz.tzlocal()).strftime("%Y-%m-%d %H:%M:%S %Z") if AUTOMATION_STATE.cooldown_until else "Ready")

with stats_col:
    live_refresh = st.checkbox("Live refresh while focused", value=st.session_state.get("tab_focused", True) and running)
    st.session_state["live_refresh"] = live_refresh

start_col, stop_col = st.columns(2)
with start_col:
    if st.button("▶️ Start scanning", disabled=AUTOMATION_STATE.running):
        if not settings_manager.get("target_email"):
            st.error("Please save your target Gmail address first.")
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

with stop_col:
    if st.button("⏹️ Stop", disabled=not AUTOMATION_STATE.running):
        STOP_EVENT.set()
        if WORKER_THREAD and WORKER_THREAD.is_alive():
            WORKER_THREAD.join(timeout=2)
        AUTOMATION_STATE.running = False
        st.info("Automation stop requested.")

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
        key="log-refresh",
    )

if st.button("Refresh logs", help="Manual refresh to keep resource usage low when unfocused."):
    st.experimental_rerun()

st.caption(
    "This interface minimizes resource usage by refreshing logs only on demand or when the tab is focused with live refresh enabled."
)

