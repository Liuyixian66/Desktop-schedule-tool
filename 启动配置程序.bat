@echo off
cd /d "%~dp0"
if exist "dist\配置程序.exe" (
    start "" "dist\配置程序.exe"
) else (
    python config_app.py
)
