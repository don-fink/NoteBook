@echo off
REM Optional installer script for NoteBook
REM Creates Start Menu shortcut for the portable executable

echo.
echo ================================================
echo          NoteBook Start Menu Setup
echo ================================================
echo.
echo This script will add NoteBook to your Start Menu.
echo The program will remain portable - you can still
echo move the folder anywhere you want.
echo.
pause

set "EXE_PATH=%~dp0NoteBook.exe"
set "SHORTCUT_PATH=%APPDATA%\Microsoft\Windows\Start Menu\Programs\NoteBook.lnk"

if not exist "%EXE_PATH%" (
    echo ERROR: NoteBook.exe not found in this folder.
    echo Make sure this script is in the same folder as NoteBook.exe
    pause
    exit /b 1
)

echo Creating Start Menu shortcut...

REM Use PowerShell to create the shortcut
powershell -NoProfile -Command ^
    "$ws = New-Object -ComObject WScript.Shell; " ^
    "$s = $ws.CreateShortcut('%SHORTCUT_PATH%'); " ^
    "$s.TargetPath = '%EXE_PATH%'; " ^
    "$s.WorkingDirectory = '%~dp0'; " ^
    "$s.Description = 'NoteBook - Rich Text Note-Taking Application'; " ^
    "$s.Save()"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo SUCCESS: NoteBook added to Start Menu!
    echo.
    echo You can now:
    echo   • Search for "NoteBook" in Start Menu
    echo   • Pin it to taskbar if desired
    echo.
    echo To remove: Delete the shortcut from Start Menu
) else (
    echo.
    echo ERROR: Failed to create Start Menu shortcut.
    echo You may need to run as administrator.
)

echo.
pause