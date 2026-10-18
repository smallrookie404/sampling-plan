@echo off
rem ============================================================
rem  ONE-CLICK LAUNCHER (repo root)
rem
rem  Double-click this file (or the desktop shortcut that points
rem  to it) to start the Sampling Plan app:
rem    1. starts the local service on 127.0.0.1:8017 (hidden)
rem    2. waits until it answers /api/health
rem    3. opens the browser
rem
rem  Safe to click again while the app is running - it just
rem  re-opens the browser instead of starting a second copy.
rem
rem  Real logic: sampling-plan-app\start-app.bat
rem  NOTE: keep this file ASCII-only (cmd.exe codepage safety).
rem ============================================================
setlocal
if not exist "%~dp0sampling-plan-app\start-app.bat" (
  echo.
  echo [ERROR] sampling-plan-app\start-app.bat not found.
  echo Make sure this file sits in the repository root.
  echo.
  pause
  exit /b 1
)
call "%~dp0sampling-plan-app\start-app.bat"
endlocal
