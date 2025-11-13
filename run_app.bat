@echo off
setlocal ENABLEDELAYEDEXPANSION
cd /d "%~dp0"

:: Determine Python launcher
set "PYTHON_EXECUTABLE="
set "PYTHON_ARGS="
set "PYTHON_DISPLAY="
set "DEPENDENCY_WARNING=0"

call :detect_python "py" "-3"
if defined PYTHON_EXECUTABLE goto python_found
call :detect_python "py" ""
if defined PYTHON_EXECUTABLE goto python_found
call :detect_python "python" ""
if defined PYTHON_EXECUTABLE goto python_found
call :detect_python "python3" ""
if defined PYTHON_EXECUTABLE goto python_found

echo.^
echo Python 3 is required but was not found on your PATH.^
echo Install it from https://www.python.org/downloads/ and retry.
echo.^
pause
exit /b 1

:python_found
set "PYTHON_DISPLAY=%PYTHON_EXECUTABLE%"
if not "%PYTHON_ARGS%"=="" set "PYTHON_DISPLAY=%PYTHON_EXECUTABLE% %PYTHON_ARGS%"

echo Checking required Python packages...
call :run_python "%~dp0check_dependencies.py" --quiet
if %errorlevel%==0 goto launch_ready

set "DEPENDENCY_WARNING=1"
echo Missing packages detected. Attempting installation (internet access required)...
call :run_python -m pip install --disable-pip-version-check --no-warn-script-location -r requirements.txt
if %errorlevel% neq 0 goto pip_fail

call :run_python "%~dp0check_dependencies.py" --quiet
if %errorlevel% neq 0 goto missing_after_install

:launch_ready
set "DEPENDENCY_WARNING=0"
echo.
echo Launching the Outlook to Gmail Forwarder dashboard.
echo Close this window to stop the server when you are done.
goto launch

:launch_with_warning
echo.
echo Launching the Outlook to Gmail Forwarder dashboard (dependencies not verified).
echo Close this window to stop the server when you are done.

:launch
call :run_python -m streamlit run app.py
if %errorlevel% neq 0 goto streamlit_fail
exit /b 0

:pip_fail
echo.
echo Failed to install the Python dependencies automatically.
echo The app will still attempt to start in case everything is already installed.
echo If it fails to open, install the packages manually by running:
echo     %PYTHON_DISPLAY% -m pip install -r requirements.txt
goto launch_with_warning

:missing_after_install
echo.
echo The launcher could not verify that the required packages are installed.
echo It will still try to open the app in case the modules are already present.
echo If the app does not start, run:
echo     %PYTHON_DISPLAY% -m pip install -r requirements.txt
goto launch_with_warning

:streamlit_fail
echo.
echo Streamlit failed to start.
if "%DEPENDENCY_WARNING%"=="1" (
    echo This usually means required Python packages are missing.
    echo Please install them by running:
    echo     %PYTHON_DISPLAY% -m pip install -r requirements.txt
) else (
    echo Review the messages above for details.
)
pause
exit /b 1

:run_python
setlocal
if "%PYTHON_ARGS%"=="" (
    "%PYTHON_EXECUTABLE%" %*
) else (
    "%PYTHON_EXECUTABLE%" %PYTHON_ARGS% %*
)
set "RUN_EXIT=%errorlevel%"
endlocal & exit /b %RUN_EXIT%

:detect_python
set "CAND_EXEC=%~1"
set "CAND_ARGS=%~2"
where %CAND_EXEC% >nul 2>nul
if errorlevel 1 goto :eof
call "%CAND_EXEC%" %CAND_ARGS% -V >nul 2>nul
if errorlevel 1 goto :eof
set "PYTHON_EXECUTABLE=%CAND_EXEC%"
set "PYTHON_ARGS=%CAND_ARGS%"
goto :eof
