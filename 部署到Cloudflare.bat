@echo off
setlocal
cd /d "%~dp0"

where npx >nul 2>nul
if errorlevel 1 (
  echo npx not found. Please install Node.js: https://nodejs.org
  pause
  exit /b 1
)

echo ============================================================
echo  Step 1/3  Login to Cloudflare (a browser window will open)
echo ============================================================
call npx --yes wrangler login
if errorlevel 1 (
  echo Login failed. Please try again.
  pause
  exit /b 1
)

echo.
echo ============================================================
echo  Step 2/3  Create KV namespace (for data records)
echo ============================================================
echo Running: npx wrangler kv namespace create SAMPLING_RECORDS
echo If it succeeds, copy the returned "id" into wrangler.jsonc.
echo If it reports the namespace already exists, continue.
call npx --yes wrangler kv namespace create SAMPLING_RECORDS
echo.
echo IMPORTANT: open wrangler.jsonc and replace REPLACE_WITH_KV_NAMESPACE_ID
echo with the id above. Then, in the Cloudflare dashboard under
echo Pages -^> sampling-plan -^> Settings -^> Functions -^> KV namespace bindings,
echo add a binding named SAMPLING_RECORDS pointing to that namespace.
echo.
pause

echo.
echo ============================================================
echo  Step 3/3  Deploy to Cloudflare Pages
echo ============================================================
call npx --yes wrangler pages deploy "sampling-plan-app" --project-name sampling-plan
if errorlevel 1 (
  echo Deploy failed. See messages above.
  pause
  exit /b 1
)

echo.
echo Deployed. Visit https://sampling-plan.pages.dev
pause
endlocal
