@echo off
echo MLB成績ツール（投手・野手統合）を起動しています...
powershell -NoProfile -Command "Get-NetTCPConnection -LocalPort 3942 -ErrorAction SilentlyContinue | ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }"
timeout /t 1 /nobreak > nul
"C:\Program Files\nodejs\node.exe" "%~dp0mlb_stats_tool.js"
pause
