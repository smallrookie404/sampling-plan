@echo off
rem ============================================================
rem  Sampling Plan - one-click launcher (real implementation)
rem
rem  Called by (both are thin Chinese-named wrappers):
rem    <repo root>\<one-click>.bat   - one click from repo root
rem    .\<start app>.bat             - same folder as this file
rem
rem  Behaviour:
rem    * service already up  -> just opens the browser (safe to
rem      double-click repeatedly, never starts a second copy)
rem    * service not up      -> starts it hidden, waits for
rem      /api/health, then opens the browser
rem
rem  NOTE: keep this file ASCII-only (cmd.exe codepage safety).
rem        Chinese text belongs in the FILE NAME only.
rem
rem  NOTE: proxy-proofing. When Clash / V2Ray exports BOTH
rem        HTTP_PROXY and http_proxy, Start-Process throws
rem        "duplicate dictionary key", and Invoke-WebRequest sends
rem        127.0.0.1 through the proxy ("access denied"). So:
rem        1) drop the 4 proxy vars in the launcher's own
rem           PowerShell child before Start-Process;
rem        2) health-check with HttpWebRequest and Proxy = null.
rem        The local server never needs a proxy.
rem
rem  NOTE: the whole start/wait sequence runs in ONE PowerShell
rem        child (spawning it repeatedly costs ~5s each on some
rem        machines). Exit codes: 0 = up, 1 = started but silent,
rem        2 = could not launch node.
rem  NOTE: no "timeout /t" - it fails when stdin is redirected.
rem ============================================================
setlocal
cd /d "%~dp0"
set "URL=http://127.0.0.1:8017"
set "PREP=foreach($k in 'HTTP_PROXY','HTTPS_PROXY','http_proxy','https_proxy'){ Remove-Item ('Env:'+$k) -ErrorAction SilentlyContinue };"
set "UP=function Up { try { $r=[Net.HttpWebRequest]::Create('%URL%/api/health'); $r.Proxy=$null; $r.Timeout=1500; $r.GetResponse().Close(); $true } catch { $false } };"

rem --- 1. locate Node.js: system PATH first (full path), then bundled ---
set "NODE_EXE="
for /f "delims=" %%N in ('where node 2^>nul') do if not defined NODE_EXE set "NODE_EXE=%%N"
if not defined NODE_EXE for %%P in (
  "%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe"
  "%USERPROFILE%\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\node\bin\node.exe"
  "%ProgramFiles%\nodejs\node.exe"
  "%ProgramFiles(x86)%\nodejs\node.exe"
  "%LOCALAPPDATA%\Programs\nodejs\node.exe"
) do if not defined NODE_EXE if exist %%P set "NODE_EXE=%%~P"
if not defined NODE_EXE (
  echo.
  echo [ERROR] Node.js not found. Please install it from https://nodejs.org
  echo.
  pause
  exit /b 1
)

rem --- 2. already running? start it if needed, then wait (single PS child) ---
powershell -NoProfile -WindowStyle Hidden -Command "%PREP% %UP% if (Up) { exit 0 }; try { Start-Process -FilePath '%NODE_EXE%' -ArgumentList '%~dp0server.mjs' -WorkingDirectory '%~dp0' -WindowStyle Hidden -RedirectStandardError '%~dp0server-error.log' -RedirectStandardOutput '%~dp0server.log' } catch { exit 2 }; for ($i=0; $i -lt 25; $i++) { Start-Sleep -Milliseconds 700; if (Up) { exit 0 } }; exit 1"
if errorlevel 2 goto :launchfail
if errorlevel 1 goto :timeout
goto :open

:launchfail
echo.
echo [ERROR] Could not launch node.exe:
echo   %NODE_EXE%
echo.
pause
exit /b 1

:timeout
echo.
echo [ERROR] Service did not respond on %URL%
echo See server-error.log in this folder for details.
echo.
if exist "%~dp0server-error.log" type "%~dp0server-error.log"
pause
exit /b 1

:open
start "" "%URL%"
endlocal
exit /b 0
