@echo off
setlocal
cd /d "%~dp0"
set "VENV=%LOCALAPPDATA%\Sklad_win\venv"
if not exist "%VENV%\Scripts\python.exe" (
    where py >nul 2>&1
    if errorlevel 1 (
        python -m venv "%VENV%" || goto error
    ) else (
        py -3 -m venv "%VENV%" || goto error
    )
)
"%VENV%\Scripts\python.exe" -c "import openpyxl, keyring" >nul 2>&1
if errorlevel 1 "%VENV%\Scripts\python.exe" -m pip install -r requirements.txt || goto error
"%VENV%\Scripts\python.exe" app.py || goto error
exit /b 0

:error
echo Sklad_win could not start. Install Python 3.10+ from python.org and try again.
pause
exit /b 1
