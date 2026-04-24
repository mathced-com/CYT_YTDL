@echo off
cd /d "%~dp0"
echo ==============================================
echo      Start Downloading (Sync from GitHub)
echo ==============================================
echo.

echo Pulling latest changes from GitHub...
git pull

echo.
echo ==============================================
echo      Download Complete!
echo ==============================================
pause
