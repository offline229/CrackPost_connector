@echo off
rem 在脚本所在目录启动一个独立的 http.server（端口8000），并在默认浏览器中打开可视化页面
pushd "%~dp0"

rem 如果存在虚拟环境且想用虚拟环境的 python，请取消下面两行注释并调整路径
rem if exist "%~dp0venv\Scripts\python.exe" (
rem     start "CrackPost HTTP" "%~dp0venv\Scripts\python.exe" -m http.server 8000
rem ) else (
start "CrackPost HTTP" cmd /c "python -m http.server 8000"
rem )

rem 等待1秒让服务器启动
timeout /t 1 /nobreak >nul

start "" "http://localhost:8000/visualization_private.html"

popd
exit /b 0