@echo off
cd /d "%~dp0"
echo ==============================================
echo      Start Uploading (Sync to GitHub)
echo ==============================================
echo.

echo [1/3] Adding files to Git...
git add .

echo [2/3] Committing changes...
git commit -m "Auto sync progress: %date% %time%"

echo [3/3] Pushing to GitHub...
git push

echo.
echo ==============================================
echo      Upload Complete!
echo ==============================================
pause
