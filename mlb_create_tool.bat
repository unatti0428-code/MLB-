@echo off
echo MLB成績ツールを起動しています...
powershell -NoProfile -Command "Get-NetTCPConnection -LocalPort 3940 -ErrorAction SilentlyContinue | ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }"
timeout /t 1 /nobreak > nul
"C:\Program Files\nodejs\node.exe" "%~dp0mlb_create_tool.js"
pause
