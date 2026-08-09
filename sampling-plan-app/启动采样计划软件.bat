@echo off
setlocal
cd /d "%~dp0"

rem Find Node.js: system node first, then the bundled runtime on this machine
set "NODE_EXE="
where node >nul 2>nul && set "NODE_EXE=node"
if not defined NODE_EXE if exist "C:\Users\Administrator\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe" set "NODE_EXE=C:\Users\Administrator\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe"
if not defined NODE_EXE (
  echo Node.js not found. Please install Node.js from https://nodejs.org
  pause
  exit /b 1
)

rem Start local service in a hidden window (exits silently if already running)
powershell -NoProfile -WindowStyle Hidden -Command "Start-Process -FilePath '%NODE_EXE%' -ArgumentList '%~dp0server.mjs' -WorkingDirectory '%~dp0' -WindowStyle Hidden -RedirectStandardError '%~dp0server-error.log' -RedirectStandardOutput '%~dp0server.log'"

rem Wait for the service and self-check
timeout /t 2 /nobreak >nul
powershell -NoProfile -Command "try { $r = Invoke-WebRequest -UseBasicParsing 'http://127.0.0.1:8017/api/health' -TimeoutSec 3; if ($r.StatusCode -eq 200) { exit 0 } else { exit 1 } } catch { exit 1 }"
if errorlevel 1 (
  echo.
  echo Service failed to start: no response on http://127.0.0.1:8017
  echo See server-error.log for details.
  if exist server-error.log type server-error.log
  pause
  exit /b 1
)

start "" "http://127.0.0.1:8017"
endlocal
