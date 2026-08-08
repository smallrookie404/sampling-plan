@echo off
chcp 65001 >nul
setlocal
set "FOUND="
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8017" ^| findstr "LISTENING"') do (
  set "FOUND=1"
  taskkill /f /pid %%a >nul 2>nul
)
if defined FOUND (
  echo 采样计划软件服务已停止。
) else (
  echo 采样计划软件服务未在运行。
)
pause
endlocal
