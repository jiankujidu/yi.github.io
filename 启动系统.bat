@echo off
chcp 65001 >nul
echo 正在启动ELISA代测表填写系统...
echo.
echo [1/2] 启动后端服务...
start /b python server.py
timeout /t 2 >nul
echo [2/2] 打开填写页面...
start index.html
echo.
echo 系统已启动！
echo 后端服务运行在 http://127.0.0.1:8765
echo 填写完成后点击"导出Excel表格"即可下载
echo.
echo 按任意键停止服务并退出...
pause >nul
taskkill /f /im python.exe 2>nul
