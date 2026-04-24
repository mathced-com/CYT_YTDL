@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==============================================
echo      開始打包 CED_YTDL.exe
echo ==============================================
echo.

echo 正在檢查是否已安裝 PyInstaller...
py -3 -m pip install pyinstaller

echo.
echo 準備打包中，這可能會花費 1-2 分鐘...
REM 使用 --onefile 打包為單一執行檔，並使用 --windowed 隱藏終端機黑視窗
py -3 -m PyInstaller --noconfirm --onefile --windowed --name "CED_YTDL" main.py

echo.
echo ==============================================
echo      打包完成！
echo      請至「dist」資料夾尋找 CED_YTDL.exe
echo ==============================================
pause
