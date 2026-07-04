@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Останавливаю бота (python ... bot.py)...
powershell -NoProfile -Command "Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'python' -and $_.CommandLine -match 'bot\.py' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"
echo Готово.
if /i not "%~1"=="nopause" pause
