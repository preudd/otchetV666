@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Перезапуск бота...
call "%~dp0close.bat" nopause
timeout /t 2 /nobreak >nul
start "" "%~dp0openbot.bat"
echo Открыто новое окно с ботом.
pause
