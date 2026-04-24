@echo off
set PYTHONIOENCODING=utf-8
chcp 65001 >nul
cd /d "%~dp0"
py -3 release_helper.py
if %errorlevel% neq 0 pause
