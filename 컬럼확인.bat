@echo off
chcp 65001 >nul
cd /d "%~dp0"
python check_columns.py
pause

