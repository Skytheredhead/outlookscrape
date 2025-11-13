@echo off
setlocal ENABLEDELAYEDEXPANSION
cd /d "%~dp0"

:: Determine Python launcher
where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON=py -3"
) else (
    where python >nul 2>nul
    if %errorlevel%==0 (
        set "PYTHON=python"
    ) else (
        echo.^
        echo Python 3 is required but was not found on your PATH.^
        echo Install it from https://www.python.org/downloads/ and retry.
        echo.^
        pause
        exit /b 1
    )
)

echo Ensuring required Python packages are installed...
"!PYTHON!" -m pip install --upgrade pip >nul
if %errorlevel% neq 0 goto pip_fail
"!PYTHON!" -m pip install --disable-pip-version-check --no-warn-script-location -r requirements.txt
if %errorlevel% neq 0 goto pip_fail

echo.
echo Launching the Outlook to Gmail Forwarder dashboard.
echo Close this window to stop the server when you are done.
"!PYTHON!" -m streamlit run app.py
if %errorlevel% neq 0 goto streamlit_fail
exit /b 0

:pip_fail
echo.
echo Failed to install the Python dependencies. Check your internet connection and try again.
pause
exit /b 1

:streamlit_fail
echo.
echo Streamlit failed to start. Review the messages above for details.
pause
exit /b 1
