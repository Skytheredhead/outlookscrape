@echo off
setlocal ENABLEDELAYEDEXPANSION
cd /d "%~dp0"

:: Determine Python launcher
set "PYTHON_EXECUTABLE="
set "PYTHON_ARGS="
set "PYTHON_DISPLAY="

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
call "%PYTHON_EXECUTABLE%" %PYTHON_ARGS% "%~dp0check_dependencies.py" --quiet
if %errorlevel%==0 goto launch_ready

echo Missing packages detected. Attempting installation (internet access required)...
call "%PYTHON_EXECUTABLE%" %PYTHON_ARGS% -m pip install --disable-pip-version-check --no-warn-script-location -r requirements.txt
if %errorlevel% neq 0 goto pip_fail

call "%PYTHON_EXECUTABLE%" %PYTHON_ARGS% "%~dp0check_dependencies.py" --quiet
if %errorlevel% neq 0 goto missing_after_install

:launch_ready
echo.
echo Launching the Outlook to Gmail Forwarder dashboard.
echo Close this window to stop the server when you are done.
:launch
call "%PYTHON_EXECUTABLE%" %PYTHON_ARGS% -m streamlit run app.py
if %errorlevel% neq 0 goto streamlit_fail
exit /b 0

:pip_fail
echo.
echo Failed to install the Python dependencies. Check your internet connection and try again.
echo You can also install them manually by running:
echo     %PYTHON_DISPLAY% -m pip install -r requirements.txt
pause
exit /b 1

:missing_after_install
echo.
echo The launcher could not verify that the required packages are installed.
echo Run the following command in this folder and check for errors:
echo     %PYTHON_DISPLAY% -m pip install -r requirements.txt
echo After it succeeds, double-click run_app.bat again.
pause
exit /b 1

:streamlit_fail
echo.
echo Streamlit failed to start. Review the messages above for details.
pause
exit /b 1

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
