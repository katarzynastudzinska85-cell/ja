@echo off
chcp 65001 >nul
"%LOCALAPPDATA%\Programs\Python\Python312\python.exe" -I "%~dp0szukaj.py" %*
