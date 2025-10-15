@echo off
REM Auto-generated launcher for NoteBook
setlocal
cd /d %~dp0
IF NOT EXIST .venv (echo Virtual environment missing. Run scripts\install_windows.ps1 & exit /b 1)
call .venv\Scripts\activate.bat
"%~dp0.venv\Scripts\python.exe" "%~dp0main.py" %*
endlocal
