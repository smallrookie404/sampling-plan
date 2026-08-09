@echo off
setlocal
set "FOUND="
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8017" ^| findstr "LISTENING"') do (
  set "FOUND=1"
  taskkill /f /pid %%a >nul 2>nul
)
if defined FOUND (
  echo Service stopped.
) else (
  echo Service is not running.
)
pause
endlocal
