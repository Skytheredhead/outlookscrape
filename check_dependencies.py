"""Utility to verify that required Python packages are importable.

This script is primarily used by run_app.bat to determine whether the
project's dependencies are already installed. When the modules are present it
exits with status code 0, otherwise it prints the missing imports and exits
with status code 1.
"""
from __future__ import annotations

import importlib.util
import sys
from typing import Iterable, List, Tuple

# (package_name, module_name)
REQUIRED_MODULES: Tuple[Tuple[str, str], ...] = (
    ("selenium", "selenium"),
    ("webdriver-manager", "webdriver_manager"),
    ("google-api-python-client", "googleapiclient"),
    ("google-auth-httplib2", "google_auth_httplib2"),
    ("google-auth-oauthlib", "google_auth_oauthlib"),
    ("streamlit", "streamlit"),
    ("python-dateutil", "dateutil"),
)


def find_missing_modules(modules: Iterable[Tuple[str, str]]) -> List[str]:
    missing: List[str] = []
    for package_name, module_name in modules:
        if importlib.util.find_spec(module_name) is None:
            missing.append(f"{package_name} (import '{module_name}')")
    return missing


def main(argv: List[str]) -> int:
    quiet = "--quiet" in argv
    missing = find_missing_modules(REQUIRED_MODULES)

    if missing:
        if not quiet:
            print("Missing Python packages detected:")
            for item in missing:
                print(f"  - {item}")
            print("\nInstall them by running:\n    python -m pip install -r requirements.txt")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
