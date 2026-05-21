@echo off
setlocal

cd /d "%~dp0"

set "PYTHON_CMD="

set "PYTHON_CMD=py -3"
if defined PYTHON_CMD (
    %PYTHON_CMD% --version >nul 2>nul
    if errorlevel 1 (
        set "PYTHON_CMD="
    )
)

if not defined PYTHON_CMD (
    set "PYTHON_CMD=python"
)

if defined PYTHON_CMD (
    %PYTHON_CMD% --version >nul 2>nul
    if errorlevel 1 (
        set "PYTHON_CMD="
    )
)

if not defined PYTHON_CMD (
    echo [ERROR] Python was not found. Please install Python 3 and try again.
    pause
    exit /b 1
)

echo Checking required Python packages...
%PYTHON_CMD% -c "import importlib.util, sys; missing=[m for m in ['pandas','openpyxl','docxtpl','tkinterdnd2','win32com'] if importlib.util.find_spec(m) is None]; print(','.join(missing)); sys.exit(1 if missing else 0)"

if errorlevel 1 (
    echo Missing packages found. Installing from requirements.txt...
    %PYTHON_CMD% -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [ERROR] Failed to install dependencies.
        pause
        exit /b 1
    )
) else (
    echo All required packages are installed.
)

echo Starting Contract Generator...
%PYTHON_CMD% app.py

if errorlevel 1 (
    echo [ERROR] The application exited with an error.
    pause
    exit /b 1
)

endlocal
