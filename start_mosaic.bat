@echo off
cd /d "%~dp0"

where pyw.exe >nul 2>nul
if not errorlevel 1 (
    start "" pyw.exe -3 "%~dp0launcher.pyw"
    exit /b
)

where pythonw.exe >nul 2>nul
if not errorlevel 1 (
    start "" pythonw.exe "%~dp0launcher.pyw"
    exit /b
)

echo Python could not be found. Install Python 3 and try again.
pause
