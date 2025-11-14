# Outlook ➜ Gmail Forwarder

Automate copying new Outlook webmail into your primary Gmail account while staying within Outlook's web interface restrictions.

## Manual installation (quick start)
1. Install Python 3.9 or later. On Windows, either the `python` command or the `py` launcher works.
2. Install the dependencies listed in `requirements.txt`:
   ```bash
   python -m pip install -r requirements.txt
   ```
   On Windows you can also use the launcher syntax:
   ```bash
   py -3 -m pip install -r requirements.txt
   ```
3. Place your `credentials.json` (downloaded from the Google Cloud Console) next to `app.py`.
4. Launch the Streamlit dashboard:
   ```bash
   python -m streamlit run app.py
   ```
   or, on Windows:
   ```bash
   py -3 -m streamlit run app.py
   ```
5. Follow the on-screen prompts to set your Gmail destination, log into Outlook via the dedicated Chrome window, and start the background watcher.

## Features
- Streamlit control center with start/stop controls, focus-aware live logging, and dependency self-checks.
- Card-based layout with gradient styling, compact status metrics, and streamlined buttons.
- Customizable polling cadence with human-like randomization between the limits you choose.
- Manual and headless Selenium sessions backed by a persistent Chrome profile to minimize CAPTCHAs.
- Gmail API integration for forwarding and alert notifications.
- Persistent forwarding history and automatic cooldowns when Outlook objects.

## Prerequisites
- Python 3.9 or later on Windows (tested) or other desktop OS.
- Google Cloud project with the Gmail API enabled (instructions below).
- Chrome browser installed (ChromeDriver is downloaded automatically via `webdriver-manager`).

## Installation
Install the required Python packages before running the app. The easiest cross-platform command is:

```bash
python -m pip install -r requirements.txt
```

On Windows you can replace `python` with `py -3` if you prefer the Python launcher.

## Gmail API setup
1. Visit the [Google Cloud Console](https://console.cloud.google.com/apis/credentials).
2. Create a project (or reuse an existing one) and enable the **Gmail API**.
3. Under **APIs & Services → Credentials**, create an **OAuth client ID** of type **Desktop app**.
4. Download the `credentials.json` file and place it in the same directory as `app.py`.
5. The Streamlit UI provides a **Login to Gmail API** button to launch the OAuth consent screen. The resulting `token.json` is stored locally for reuse.

If you see `Error 403: access_denied` while authorizing Gmail sending access:

1. Open the [Google Cloud Console OAuth consent screen](https://console.cloud.google.com/apis/credentials/consent).
2. Scroll to the **Test users** section.
3. Click **Add users**, enter the Gmail address you are using for sending, and click **Save**.
4. Retry the Gmail OAuth flow from the app.

## Running the dashboard

### Windows one-click launcher
1. Double-click `run_app.bat`.
2. The script automatically installs/updates the required Python packages (listed in `requirements.txt`) using the `py` launcher if available, or the `python` command otherwise.
3. Once the dependencies are ready, the Streamlit dashboard starts in the same window and opens your browser to `http://localhost:8501`.
4. Keep the Command Prompt window open while you use the tool. Close it to stop the server.

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
5. Use **Login to Outlook** to open a non-headless Chrome window, sign into Outlook manually, then click **Save & Close** to store the profile.
6. Adjust the Gmail destination, run a test email, fine-tune the polling window, and press **Start scanning** to begin forwarding unread Outlook emails.

## Notes
- Forwarded emails are tracked to avoid duplicates and daily counts reset automatically.
- If Outlook prompts for CAPTCHA or blocks the session, the tool pauses for 30 minutes and emails you an alert through Gmail.
- Logs and state files live under the `automation_state/` folder.
- Always test with dummy accounts before using production mailboxes.
