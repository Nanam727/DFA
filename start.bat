@echo off
setlocal
cd /d "%~dp0"

set "PY_CMD="
where python >nul 2>nul
if %errorlevel%==0 set "PY_CMD=python"

if not defined PY_CMD (
    where py >nul 2>nul
    if %errorlevel%==0 set "PY_CMD=py -3"
)

if not defined PY_CMD (
    echo [ERROR] Python 3.10+ not found in PATH.
    echo         Install Python, then run fix.bat.
    pause
    exit /b 1
)

echo [INFO] Using interpreter: %PY_CMD%
%PY_CMD% -c "import customtkinter, docx, PIL" >nul 2>nul
if errorlevel 1 (
    echo [INFO] Missing dependencies. Installing from requirements.txt ...
    %PY_CMD% -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [ERROR] Dependency install failed.
        echo         Please check your network, then run: pip install -r requirements.txt
        pause
        exit /b 1
    )
)

%PY_CMD% run.py
endlocal
