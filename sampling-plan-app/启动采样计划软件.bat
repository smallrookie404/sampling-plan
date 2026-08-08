@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

rem 查找 Node.js：优先使用系统安装的 node，找不到则使用本机 Codex 内置运行时
set "NODE_EXE="
where node >nul 2>nul && set "NODE_EXE=node"
if not defined NODE_EXE if exist "C:\Users\Administrator\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe" set "NODE_EXE=C:\Users\Administrator\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe"
if not defined NODE_EXE (
  echo 未找到 Node.js，请先安装 Node.js（https://nodejs.org）后重试。
  pause
  exit /b 1
)

rem 隐藏窗口启动本地服务（若已在运行则自动退出，不影响使用）
powershell -NoProfile -WindowStyle Hidden -Command "Start-Process -FilePath '%NODE_EXE%' -ArgumentList '\"%~dp0server.mjs\"' -WindowStyle Hidden"
timeout /t 1 /nobreak >nul

start "" "http://127.0.0.1:8017"
endlocal
