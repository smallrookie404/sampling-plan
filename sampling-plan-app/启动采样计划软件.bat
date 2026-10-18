@echo off
rem Thin wrapper - the real logic lives in start-app.bat
rem (kept ASCII-only for cmd.exe codepage safety).
if not exist "%~dp0start-app.bat" (
  echo [ERROR] start-app.bat not found next to this file.
  pause
  exit /b 1
)
call "%~dp0start-app.bat"
