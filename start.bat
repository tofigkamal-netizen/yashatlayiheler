@echo off
chcp 65001 >nul
cd /d "%~dp0"
title YASAT Layiheler - lokal server
where python >nul 2>nul || (echo Python tapilmadi. python.org-dan Python qurasdirin. & pause & exit /b 1)
echo Server isleyir: http://localhost:8000
echo Baza: %~dp0layiheler.xlsx
echo Dayandirmaq ucun bu pencereni baglayin.
echo.
start "" cmd /c "timeout /t 2 /nobreak >nul & start http://localhost:8000"
python server.py
pause
