@echo off
REM Uruchamia agenta e-Connectivity (bez podstron).
cd /d "%~dp0"
node live_diagnostic_agent.js
pause
