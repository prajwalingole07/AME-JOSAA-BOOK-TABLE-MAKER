@echo off
title JoSAA CUTOFF BOOK MAKER - IIT / NIT / IIIT
color 0B
cd /d "%~dp0"

echo ============================================================
echo     STARTING JoSAA CUTOFF BOOK MAKER LOCAL WEB APP
echo     IIT / NIT / IIIT Cut-off Book Maker
echo ============================================================
echo.
echo Server running at: http://localhost:5000
echo Login: josaa2027 / josaa@2027
echo Opening browser...
echo.

:: Open browser after 2 seconds invisibly without creating a second CMD window
start "" powershell -windowstyle hidden -Command "Start-Sleep -Seconds 2; Start-Process 'http://localhost:5000'"

:: Run the Flask server directly in this single CMD window
python web_app.py
