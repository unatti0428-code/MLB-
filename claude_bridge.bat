@echo off
echo Claude 中継サーバー（列伝AI強化用）を起動しています...
powershell -NoProfile -Command "Get-NetTCPConnection -LocalPort 3950 -ErrorAction SilentlyContinue | ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }"
timeout /t 1 /nobreak > nul
"C:\Program Files\nodejs\node.exe" "%~dp0claude_bridge.js"
pause
