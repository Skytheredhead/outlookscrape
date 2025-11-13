@echo off
setlocal ENABLEDELAYEDEXPANSION
cd /d "%~dp0"

:: Determine Python launcher
set "PY_ARGS="
where py >nul 2>nul
if %errorlevel%==0 (
    py -3 --version >nul 2>nul
    if %errorlevel%==0 (
        set "PYTHON=py"
        set "PY_ARGS=-3"
    ) else (
        set "PYTHON=py"
    )
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
set "DEPENDENCY_STATUS=ok"
"!PYTHON!" !PY_ARGS! -m pip install --upgrade pip >nul
if %errorlevel% neq 0 (
    set "DEPENDENCY_STATUS=failed"
    echo.
    echo Unable to update pip automatically. Continuing without reinstalling dependencies.
    echo Install the required packages manually by running:
    echo    "!PYTHON!" !PY_ARGS! -m pip install -r requirements.txt
) else (
    "!PYTHON!" !PY_ARGS! -m pip install --disable-pip-version-check --no-warn-script-location -r requirements.txt
    if %errorlevel% neq 0 (
        set "DEPENDENCY_STATUS=failed"
        echo.
        echo Failed to install the Python dependencies automatically. Continuing anyway...
        echo Install them manually by running:
        echo    "!PYTHON!" !PY_ARGS! -m pip install -r requirements.txt
    )
)

echo.
echo Launching the Outlook to Gmail Forwarder dashboard.
echo Close this window to stop the server when you are done.
"!PYTHON!" !PY_ARGS! -m streamlit run app.py
if %errorlevel% neq 0 goto streamlit_fail
exit /b 0

:streamlit_fail
echo.
if /I "!DEPENDENCY_STATUS!"=="failed" (
    echo Streamlit failed to start. This is often caused by missing dependencies.
    echo Install them manually by running:
    echo    "!PYTHON!" !PY_ARGS! -m pip install -r requirements.txt
) else (
    echo Streamlit failed to start. Review the messages above for details.
)
pause
exit /b 1
