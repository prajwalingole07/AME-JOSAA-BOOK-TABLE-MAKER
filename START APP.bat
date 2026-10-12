@echo off
title JoSAA IIT / NIT / IIIT Cut-off Book Table Maker
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python not found on PATH. Install Python 3 and re-run this file.
    pause
    exit /b 1
)

echo Checking dependencies...
python -c "import flask, pandas, docx, openpyxl, lxml" >nul 2>nul
if errorlevel 1 (
    echo Installing packages from requirements.txt ...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [ERROR] Package install failed. Run manually: python -m pip install -r requirements.txt
        pause
        exit /b 1
    )
)

echo.
echo Starting server at http://localhost:5000
echo Keep this window open. Press Ctrl+C to stop.
echo.
start "browser" /min cmd /c "timeout /t 3 >nul & start http://localhost:5000"
python web_app.py

pause
