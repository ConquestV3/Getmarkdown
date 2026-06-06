@echo off
echo ==============================================================
echo             GetMarkdown Local - Unlimited Converter
echo ==============================================================
echo.

set PYTHON_CMD=python
python --version >nul 2>&1
if errorlevel 1 (
    set PYTHON_CMD=
)

:: If standard python fails, look in Local AppData (typical Windows install location)
if "%PYTHON_CMD%"=="" (
    echo [+] Standard python command failed. Searching installation directories...
    for /d %%d in ("%LOCALAPPDATA%\Python\pythoncore-*") do (
        if exist "%%d\python.exe" (
            set PYTHON_CMD="%%d\python.exe"
            goto :found_python
        )
    )
    for /d %%d in ("C:\Program Files\Python*") do (
        if exist "%%d\python.exe" (
            set PYTHON_CMD="%%d\python.exe"
            goto :found_python
        )
    )
    for /d %%d in ("C:\Program Files (x86)\Python*") do (
        if exist "%%d\python.exe" (
            set PYTHON_CMD="%%d\python.exe"
            goto :found_python
        )
    )
)

:found_python
if "%PYTHON_CMD%"=="" (
    echo [ERROR] Python was not found in common installation directories.
    echo Please install Python 3.8 or newer and check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b
)

echo [+] Using Python executable: %PYTHON_CMD%
%PYTHON_CMD% --version

:: Create Virtual Environment if not exists
if not exist .venv (
    echo [+] Creating virtual environment (.venv)...
    %PYTHON_CMD% -m venv .venv
    if errorlevel 1 (
        echo [ERROR] Failed to initialize virtual environment.
        pause
        exit /b
    )
)

:: Activate Virtual Environment
echo [+] Activating virtual environment...
call .venv\Scripts\activate

:: Upgrade pip and install libraries
echo [+] Installing requirements from requirements.txt...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Failed to install backend dependencies.
    pause
    exit /b
)

echo.
echo [+] GetMarkdown Local successfully configured!
echo [+] Starting server at http://127.0.0.1:8000
echo [+] Automatically opening web browser...
echo.

:: Open default browser
start "" "http://127.0.0.1:8000"

:: Boot FastAPI app
python main.py

pause
