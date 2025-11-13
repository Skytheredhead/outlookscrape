# Outlook ➜ Gmail Forwarder

Automate copying new Outlook webmail into your primary Gmail account while staying within Outlook's web interface restrictions.

## Features
- Streamlit dashboard with start/stop controls, focus-aware live logging, and dependency self-checks.
- Securely encrypted storage of Outlook credentials (Fernet encryption, decrypted only at use time).
- Manual and headless Selenium sessions with cookie reuse and CAPTCHA detection.
- Gmail API integration for forwarding and alert notifications.
- Human-like polling cadence, cooldowns on failure, and persistent forwarding history.

## Prerequisites
- Python 3.9 or later on Windows (tested) or other desktop OS.
- Google Cloud project with the Gmail API enabled (instructions below).
- Chrome browser installed (ChromeDriver is downloaded automatically via `webdriver-manager`).

## Installation
Install the required Python packages before running the app. The easiest cross-platform command is:

```bash
python -m pip install -r requirements.txt
```

On Windows you can replace `python` with `py -3`, `py`, or whichever Python command you normally use.

## Gmail API setup
1. Visit the [Google Cloud Console](https://console.cloud.google.com/apis/credentials).
2. Create a project (or reuse an existing one) and enable the **Gmail API**.
3. Under **APIs & Services → Credentials**, create an **OAuth client ID** of type **Desktop app**.
4. Download the `credentials.json` file and place it in the same directory as `app.py`.
5. The Streamlit UI provides a **Login to Gmail API** button to launch the OAuth consent screen. The resulting `token.json` is stored locally for reuse.

## Running the dashboard

### Windows one-click launcher
1. Double-click `run_app.bat`.
2. The script verifies that the required Python packages (listed in `requirements.txt`) are installed. If any are missing it attempts to download them using whichever Python command works first out of `py -3`, `py`, `python`, or `python3`.
3. If the automatic install cannot run (for example because `pip` is unavailable) the launcher still tries to open the app so you can continue when the packages are already present. When Streamlit cannot start it will prompt you to install the requirements manually with the exact Python command it detected.
4. Once the dependencies are ready, the Streamlit dashboard starts in the same window and opens your browser to `http://localhost:8501`.
5. Keep the Command Prompt window open while you use the tool. Close it to stop the server.

### Manual launch (any platform)
1. Ensure the dependencies above are installed. You can use the provided `requirements.txt` file:
   ```bash
   python -m pip install -r requirements.txt
   ```
2. Place `credentials.json` next to `app.py`.
3. Start the UI with:
   ```bash
   python -m streamlit run app.py
   ```
4. The dashboard will automatically open in your default browser at `http://localhost:8501`.
5. Enter your Outlook username and password once via the UI to encrypt and save them. The plaintext values are discarded immediately after encryption.
6. Use **Launch Manual Login** to open a non-headless Chrome window, sign into Outlook manually, then click **Save Cookies & Close Browser**.
7. Press **Start scanning** to begin forwarding unread Outlook emails to the Gmail address you saved in settings.

## Notes
- Forwarded emails are tracked to avoid duplicates and daily counts reset automatically.
- If Outlook prompts for CAPTCHA or blocks the session, the tool pauses for 30 minutes and emails you an alert through Gmail.
- Logs and state files live under the `automation_state/` folder.
- Always test with dummy accounts before using production mailboxes.
