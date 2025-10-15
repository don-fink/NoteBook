@echo off
rem Minimal, robust installer wrapper with logging.
setlocal

set "SCRIPT_DIR=%~dp0"
set "PS1=%SCRIPT_DIR%scripts\install_windows.ps1"
set "LOG=%SCRIPT_DIR%install_log.txt"

echo ===============================================
echo NoteBook Installer Wrapper
echo Script Dir: %SCRIPT_DIR%
echo PS1 Path  : %PS1%
echo Log File  : %LOG%
echo Args      : %*
echo ===============================================

if not exist "%PS1%" (
  echo [ERROR] scripts\install_windows.ps1 not found next to install.cmd.
  echo Aborting.
  pause
  exit /b 1
)

echo Starting PowerShell installer... (this may take a minute)
echo (A transcript will be written to install_log.txt)
echo. > "%LOG%"
echo [LOG] %date% %time% Launching installer with arguments: %* >> "%LOG%"

REM Use cmd /c powershell so any quoting issues are minimized, and capture exit code.
powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%PS1%" %* 1>>"%LOG%" 2>&1
set "ERR=%ERRORLEVEL%"
echo.>>"%LOG%"
echo [LOG] %date% %time% Installer exit code: %ERR%>>"%LOG%"

if not "%ERR%"=="0" (
  echo Installer failed (exit %ERR%). See install_log.txt for details.
) else (
  echo Installer completed successfully. See install_log.txt for transcript.
)

echo.
echo Press any key to close this window...
pause >nul
exit /b %ERR%
