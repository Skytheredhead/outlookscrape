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

echo Checking required Python packages...
"!PYTHON!" "%~dp0check_dependencies.py" --quiet
if %errorlevel%==0 goto launch_ready

echo Missing packages detected. Attempting installation (internet access required)...
"!PYTHON!" -m pip install --disable-pip-version-check --no-warn-script-location -r requirements.txt
if %errorlevel% neq 0 goto pip_fail

"!PYTHON!" "%~dp0check_dependencies.py" --quiet
if %errorlevel% neq 0 goto missing_after_install

:launch_ready
echo.
echo Launching the Outlook to Gmail Forwarder dashboard.
echo Close this window to stop the server when you are done.
:launch
"!PYTHON!" -m streamlit run app.py
if %errorlevel% neq 0 goto streamlit_fail
exit /b 0

:pip_fail
echo.
echo Failed to install the Python dependencies. Check your internet connection and try again.
echo You can also install them manually by running:
echo     python -m pip install -r requirements.txt
pause
exit /b 1

:missing_after_install
echo.
echo The launcher could not verify that the required packages are installed.
echo Run the following command in this folder and check for errors:
echo     python -m pip install -r requirements.txt
echo After it succeeds, double-click run_app.bat again.
pause
exit /b 1

:streamlit_fail
echo.
echo Streamlit failed to start. Review the messages above for details.
pause
exit /b 1
