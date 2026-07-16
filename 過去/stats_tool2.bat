@echo off
echo MLB投手成績 能力値追加ツールを起動しています...
powershell -NoProfile -Command "Get-NetTCPConnection -LocalPort 3942 -ErrorAction SilentlyContinue | ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }"
timeout /t 1 /nobreak > nul
"C:\Program Files\nodejs\node.exe" "%~dp0stats_tool2.js"
pause
