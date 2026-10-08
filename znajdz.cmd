@echo off
REM Wyszukuje klienta po numerze J. Uzycie: znajdz.cmd 76000978
cd /d "%~dp0"
node znajdz.js %*
pause
