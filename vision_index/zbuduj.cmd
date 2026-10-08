@echo off
chcp 65001 >nul
rem Buduje / aktualizuje indeks dokumentacji Vision w C:\ServiceIndex_Vision (folder źródłowy tylko do odczytu)
set PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe
"%PY%" -I "%~dp0build_index.py" %* || goto :blad
"%PY%" -I "%~dp0build_db.py" || goto :blad
echo.
echo Gotowe. Raport: C:\ServiceIndex_Vision\raport_indeksu.txt
pause
exit /b 0
:blad
echo BLAD - sprawdz komunikat powyzej.
pause
exit /b 1
