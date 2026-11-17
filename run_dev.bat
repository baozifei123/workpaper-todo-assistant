@echo off
chcp 65001 >nul
cd /d "%~dp0"
pythonw -X utf8 -m workpaper_todo.app
