@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

where git >nul 2>nul
if errorlevel 1 (
  echo 未找到 git，请先安装 Git：https://git-scm.com
  pause
  exit /b 1
)

git add "sampling-plan-app/data/records.json"
if errorlevel 1 (
  echo 添加文件失败。
  pause
  exit /b 1
)

for /f "tokens=1-3 delims=/ " %%a in ("%date%") do set D=%%a%%b%%c
git commit -m "数据库同步 %D% %time%"

git push
if errorlevel 1 (
  echo.
  echo 推送失败：请确认已在 git 中配置 GitHub 远程仓库地址，例如：
  echo   git remote add origin https://github.com/你的用户名/仓库名.git
  echo 并在首次推送时登录 GitHub 授权。
  pause
  exit /b 1
)

echo.
echo 数据库已同步到 GitHub。
pause
endlocal
