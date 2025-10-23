@echo off
REM Clear Windows icon cache to fix icon display issues

echo Clearing Windows icon cache...
echo.

REM Stop Explorer
echo Stopping Windows Explorer...
taskkill /f /im explorer.exe >nul 2>&1

REM Clear icon cache files
echo Clearing icon cache files...
del /a /q "%userprofile%\AppData\Local\IconCache.db" >nul 2>&1
del /a /f /q "%userprofile%\AppData\Local\Microsoft\Windows\Explorer\iconcache*" >nul 2>&1

REM Restart Explorer
echo Restarting Windows Explorer...
start explorer.exe

echo.
echo Icon cache cleared! Icons should refresh shortly.
echo If icons still don't show, try restarting your computer.
echo.
pause