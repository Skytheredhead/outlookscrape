"""Convenience launcher for the Outlook ➜ Gmail Forwarder Streamlit app."""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def main() -> None:
    app_path = Path(__file__).resolve().parent / "app.py"
    cmd = [sys.executable, "-m", "streamlit", "run", str(app_path)]
    subprocess.run(cmd, check=True)


if __name__ == "__main__":
    main()
