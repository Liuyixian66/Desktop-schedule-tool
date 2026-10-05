@echo off
cd /d "%~dp0"
if exist "dist\桌面课表.exe" (
    start "" "dist\桌面课表.exe"
) else (
    start "" pythonw main_app.py
)
