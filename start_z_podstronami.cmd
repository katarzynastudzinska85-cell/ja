@echo off
REM Uruchamia agenta e-Connectivity razem z podstronami AAA (np. TD-016837).
cd /d "%~dp0"
set AAA_FOLLOW_LINKS=1
node live_diagnostic_agent.js
pause
