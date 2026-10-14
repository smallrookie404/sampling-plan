@echo off
setlocal
cd /d "%~dp0"

rem =====================================================================
rem  Account Bridge - required by "Data Report Rollback" (sampling-plan)
rem
rem  The App's rollback runs in two steps and each step must be executed
rem  under its own identity:
rem    step 1 (jcbg / test result report) -> btc-gxz
rem    step 2 (jcdg / test working paper) -> btc-wcy
rem  The browser cannot hold those tokens, so this local bridge serves them
rem  from the account pool (pool.py). Cached tokens cost no captcha.
rem
rem  Listens on http://127.0.0.1:8787
rem  KEEP THIS WINDOW OPEN while using "Data Report Rollback".
rem =====================================================================

echo ============================================================
echo  Account Bridge  -  http://127.0.0.1:8787
echo  Keep this window OPEN while using "Data Report Rollback".
echo  Press Ctrl+C to stop.
echo ============================================================
echo.

set "PY_EXE="
where py >nul 2>nul && set "PY_EXE=py"
if not defined PY_EXE (
  where python >nul 2>nul && set "PY_EXE=python"
)
if not defined PY_EXE for %%P in (
  "%LOCALAPPDATA%\Programs\Python\Python314\python.exe"
  "%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
  "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
  "%ProgramFiles%\Python314\python.exe"
  "%ProgramFiles%\Python313\python.exe"
  "%ProgramFiles%\Python312\python.exe"
) do if not defined PY_EXE if exist %%P set "PY_EXE=%%~P"

if not defined PY_EXE (
  echo Python not found. Please install Python 3 from https://www.python.org
  echo.
  pause
  exit /b 1
)

echo Using Python: %PY_EXE%
echo.
"%PY_EXE%" "%~dp0bridge.py"

echo.
echo Bridge stopped.
pause
