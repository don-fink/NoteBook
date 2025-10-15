@echo off
echo Building NoteBook executable...
pyinstaller notebook.spec --clean
echo.
echo Build complete! Check the dist folder for NoteBook.exe
echo To create a release package, run scripts\create_release_simple.ps1
pause