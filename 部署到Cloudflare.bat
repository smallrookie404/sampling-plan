@echo off
setlocal EnableExtensions
cd /d "%~dp0"

where npx >nul 2>nul
if errorlevel 1 (
  echo [ERROR] npx not found. Please install Node.js: https://nodejs.org
  pause
  exit /b 1
)

echo ============================================================
echo  Step 1/4  Login to Cloudflare (a browser window will open)
echo  If CLOUDFLARE_API_TOKEN is already set, this step is skipped.
echo ============================================================
if not defined CLOUDFLARE_API_TOKEN (
  call npx --yes wrangler login
  if errorlevel 1 (
    echo Login failed. Please try again.
    pause
    exit /b 1
  )
)

echo.
echo ============================================================
echo  Step 2/4  Create KV namespace "SAMPLING_RECORDS" if needed
echo ============================================================
findstr /C:"REPLACE_WITH_KV_NAMESPACE_ID" wrangler.jsonc >nul 2>nul
if not errorlevel 1 (
  echo wrangler.jsonc still contains the placeholder KV id.
  call npx --yes wrangler kv namespace create SAMPLING_RECORDS
  echo.
  echo Please copy the "id" from the output above, open wrangler.jsonc
  echo and replace both REPLACE_WITH_KV_NAMESPACE_ID values with it.
  pause
)

echo.
echo ============================================================
echo  Step 3/4  Build Pages Functions into a Worker bundle
echo ============================================================
set "BUILD_DIR=.wrangler\worker-build"
set "DIST_DIR=.wrangler\pages-dist"
if exist "%BUILD_DIR%" rmdir /s /q "%BUILD_DIR%"
if exist "%DIST_DIR%" rmdir /s /q "%DIST_DIR%"
mkdir "%BUILD_DIR%" >nul 2>nul
mkdir "%DIST_DIR%" >nul 2>nul

pushd sampling-plan-app
call npx --yes wrangler pages functions build --outdir "..\%BUILD_DIR%"
popd
if not exist "%BUILD_DIR%\index.js" (
  echo [ERROR] Function build failed: index.js was not generated.
  pause
  exit /b 1
)

echo  Assembling deployment directory...
copy /y "sampling-plan-app\index.html" "%DIST_DIR%\" >nul
copy /y "%BUILD_DIR%\index.js" "%DIST_DIR%\_worker.js" >nul
xcopy /e /i /y "sampling-plan-app\css" "%DIST_DIR%\css" >nul
xcopy /e /i /y "sampling-plan-app\js" "%DIST_DIR%\js" >nul
xcopy /e /i /y "sampling-plan-app\data" "%DIST_DIR%\data" >nul

echo.
echo ============================================================
echo  Step 4/4  Deploy to Cloudflare Pages
echo ============================================================
call npx --yes wrangler pages deploy "%DIST_DIR%" --project-name sampling-plan
if errorlevel 1 (
  echo Deploy failed. See messages above.
  pause
  exit /b 1
)

echo.
echo Deployed. Visit https://sampling-plan.pages.dev
pause
endlocal
