@echo off
cd /d "%~dp0"
py licence_tool.py
if errorlevel 1 python licence_tool.py
pause