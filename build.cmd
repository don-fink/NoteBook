@echo off
setlocal
pushd "%~dp0"

echo Building NoteBook executable...

REM Prefer project virtual environment's Python if present
set "VENV_PY=.venv\Scripts\python.exe"
if exist "%VENV_PY%" (
	echo Using venv: %VENV_PY%
	"%VENV_PY%" -m PyInstaller notebook.spec --clean
) else (
	echo No venv found, falling back to system PyInstaller on PATH
	pyinstaller notebook.spec --clean
)

if errorlevel 1 (
	echo.
	echo PyInstaller build failed. See output above for details.
	popd
	endlocal
	exit /b 1
)

echo.
echo Build complete! Check the dist folder for NoteBook.exe
echo To create a release package, run .\create_release_simple.ps1
popd
endlocal
pause