@echo off
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"

echo ==============================================
echo       Starting CYT_YTDL Downloader...
echo ==============================================

set PYTHON_CMD=

REM 1. Try parent directory virtual environment .venv
if exist "..\.venv\Scripts\python.exe" (
    set "PYTHON_CMD=..\.venv\Scripts\python.exe"
    goto :FOUND
)

REM 2. Try current directory virtual environment .venv
if exist ".venv\Scripts\python.exe" (
    set "PYTHON_CMD=.venv\Scripts\python.exe"
    goto :FOUND
)

REM 3. Try global py launcher
py -3 --version >nul 2>&1
if %errorlevel% equ 0 (
    set "PYTHON_CMD=py -3"
    goto :FOUND
)

REM 4. Try global python
python --version >nul 2>&1
if %errorlevel% equ 0 (
    set "PYTHON_CMD=python"
    goto :FOUND
)

:FOUND
if "%PYTHON_CMD%"=="" (
    echo [ERROR] Python not found!
    echo Please install Python and add it to your PATH environment variable.
    echo Or create a .venv virtual environment in the workspace.
    echo.
    pause
    exit /b 1
)

echo [OK] Using Python: %PYTHON_CMD%
echo.
echo Launching GUI...
%PYTHON_CMD% main.py

if %errorlevel% neq 0 (
    echo.
    echo =========================================
    echo [ERROR] Failed to start! Please screenshot this window.
    echo =========================================
    pause
)
