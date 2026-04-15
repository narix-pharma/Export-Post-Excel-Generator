@echo off
title Narix — Export Bill Generator
color 0A
cls

echo.
echo  ==========================================
echo    NARIX PHARMACEUTICALS
echo    Export Bill Generator
echo  ==========================================
echo.

:: Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo  [!] Python not found. Installing now...
    echo      This may take 2-3 minutes on first run.
    echo.
    winget install -e --id Python.Python.3.11 --accept-source-agreements --accept-package-agreements
    if %errorlevel% neq 0 (
        echo.
        echo  [ERROR] Could not auto-install Python.
        echo  Please go to https://www.python.org/downloads/
        echo  Download and install Python, then run this file again.
        echo.
        pause
        exit /b 1
    )
    echo  Restarting to apply changes...
    start "" "%~f0"
    exit /b
)

echo  [1/3] Python found... OK
echo.

:: Kill any existing instance on port 5000
echo  Checking for existing server...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":5000 " 2^>nul') do (
    taskkill /F /PID %%a >nul 2>&1
)
timeout /t 1 /nobreak >nul

:: Install dependencies quietly (only installs if missing)
echo  [2/3] Checking dependencies...
python -m pip install flask pdfplumber openpyxl python-dateutil --quiet --disable-pip-version-check 2>nul
echo        Ready.
echo.

echo  [3/3] Starting server...
echo.

:: Start Flask
cd /d "%~dp0"
start /B python app.py > app_log.txt 2>&1

:: Wait up to 60 seconds — product DB loading can take a moment
echo  Waiting for server to be ready...
set /a attempts=0
:waitloop
set /a attempts+=1
if %attempts% gtr 60 (
    echo.
    echo  [ERROR] Server failed to start after 60 seconds.
    echo  Check app_log.txt in this folder for details.
    echo.
    pause
    exit /b 1
)
python -c "import urllib.request; urllib.request.urlopen('http://localhost:5000/health', timeout=3)" >nul 2>&1
if %errorlevel% neq 0 (
    timeout /t 1 /nobreak >nul
    goto waitloop
)

:: Open browser
start "" "http://localhost:5000"

echo  ==========================================
echo   Server running at: http://localhost:5000
echo.
echo   Browser should have opened automatically.
echo   KEEP THIS WINDOW OPEN while using the app.
echo   Close this window to stop the server.
echo  ==========================================
echo.

:: Keep alive — check every 15 seconds (generous window for slow generates)
:keepalive
timeout /t 15 /nobreak >nul
python -c "import urllib.request; urllib.request.urlopen('http://localhost:5000/health', timeout=10)" >nul 2>&1
if %errorlevel% equ 0 goto keepalive

echo.
echo  [!] Server stopped. Restarting...
echo.
start /B python app.py >> app_log.txt 2>&1
timeout /t 5 /nobreak >nul
goto keepalive
