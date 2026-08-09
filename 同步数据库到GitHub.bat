@echo off
setlocal
cd /d "%~dp0"

where git >nul 2>nul
if errorlevel 1 (
  echo git not found. Please install Git: https://git-scm.com
  pause
  exit /b 1
)

git add "sampling-plan-app/data/records.json"
if errorlevel 1 (
  echo Failed to stage records.json.
  pause
  exit /b 1
)

for /f "tokens=1-3 delims=/ " %%a in ("%date%") do set D=%%a%%b%%c
git commit -m "db sync %D% %time%"

git push
if errorlevel 1 (
  echo.
  echo Push failed. Please configure the GitHub remote first, for example:
  echo   git remote add origin https://github.com/yourname/repo.git
  echo Then sign in to GitHub when prompted.
  pause
  exit /b 1
)

echo.
echo Database synced to GitHub.
pause
endlocal
