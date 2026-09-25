@echo off
chcp 65001 >nul
echo ========================================================
echo   LINE 連線代購系統 - 本機測試伺服器啟動中...
echo ========================================================
echo.
echo 網頁首頁請打開: http://localhost:6666/
echo 買家查單頁面: http://localhost:6666/my-orders.html
echo 賣家發卡工具: http://localhost:6666/admin-card-generator.html
echo.
echo 按 Ctrl + C 可停止本機伺服器
echo ========================================================
echo.

python -m http.server 6666
pause
