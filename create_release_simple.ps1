<#
 NoteBook Distribution Package Creator (simple)
 Location-agnostic: resolves paths relative to this script file.
#>
$ErrorActionPreference = 'Stop'

# Resolve repository root relative to this script and make it the working directory
$Root = Split-Path -Parent $PSCommandPath
Set-Location $Root

$distFolder = Join-Path $Root 'NoteBook_Release'
Write-Host "Creating PyInstaller distribution package..." -ForegroundColor Cyan

# Clean and create distribution folder
if (Test-Path $distFolder) { Remove-Item -Recurse -Force $distFolder }
New-Item -ItemType Directory $distFolder | Out-Null

# Safety: ensure no machine-specific pointer file is accidentally bundled
if (Test-Path (Join-Path $Root 'settings.loc')) {
	try {
		Remove-Item (Join-Path $Root 'settings.loc') -Force -ErrorAction SilentlyContinue
		Write-Host "  Removed stray settings.loc from working directory" -ForegroundColor Yellow
	} catch {}
}

# Copy the full PyInstaller output (includes NoteBook.exe and any _internal or support files)
$distSource = Join-Path $Root 'dist\NoteBook'
if (-not (Test-Path $distSource)) {
	Write-Host "Build output folder not found. Attempting to build with PyInstaller..." -ForegroundColor Yellow
	$venvPy = Join-Path $Root '.venv\Scripts\python.exe'
	try {
		if (Test-Path $venvPy) {
			& $venvPy -m PyInstaller 'notebook.spec' --clean
		} else {
			Write-Host "No venv detected. Using system PyInstaller on PATH..." -ForegroundColor Yellow
			pyinstaller 'notebook.spec' --clean
		}
		if ($LASTEXITCODE -ne 0) {
			throw "PyInstaller exited with code $LASTEXITCODE"
		}
	} catch {
		Write-Error "Failed to build with PyInstaller: $($_.Exception.Message)"
		exit 1
	}
}
Copy-Item (Join-Path $distSource '*') $distFolder -Recurse
Write-Host "  Copied PyInstaller output (NoteBook + dependencies)" -ForegroundColor Green

# Copy optional Start Menu installer
Copy-Item (Join-Path $Root 'add_to_start_menu.cmd') $distFolder
Write-Host "  Copied Start Menu installer" -ForegroundColor Green

# Create README
$readme = @'
# NoteBook - Rich Text Note-Taking Application

## Quick Start

1. Double-click NoteBook.exe to launch
2. No Python installation required!
3. On first run, create a new database or open existing

## Optional: Add to Start Menu

Run `add_to_start_menu.cmd` to add NoteBook to your Start Menu.
This makes it searchable and easier to launch, while keeping the app portable.

## Features

- Multiple notebooks with sections and pages
- Rich text editing with formatting
- Image and video support
- Planning registers and tables
- Multiple themes

## System Requirements

- Windows 10 or Windows 11
- No additional software required

## First Run

1. Choose "New Database" to start
2. Pick storage location (Documents recommended)
3. Start creating content!

## Uninstall

Simply delete the NoteBook folder. If you used the Start Menu installer,
delete the shortcut from Start Menu → All Apps.

Enjoy!
'@

Set-Content -Path (Join-Path $distFolder 'README.txt') -Value $readme -Encoding UTF8
Add-Content -Path (Join-Path $distFolder 'README.txt') -Value "`n---`nNOTE: Settings are stored per-user at %LOCALAPPDATA%\NoteBook\settings.json. Do NOT distribute any settings.loc file; it is a machine-specific pointer and will break on other systems." -Encoding UTF8
Write-Host "  Created README.txt" -ForegroundColor Green

# Show size info (exe only)
$exeSize = [math]::Round((Get-Item (Join-Path $distFolder 'NoteBook.exe')).Length / 1MB, 1)
Write-Host ""
Write-Host "Distribution ready!" -ForegroundColor Green
Write-Host "Location: $distFolder" -ForegroundColor Cyan
Write-Host "Size: $exeSize MB" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host ("  1. Test: Double-click {0}" -f (Join-Path $distFolder 'NoteBook.exe')) -ForegroundColor White
Write-Host ("  2. Optional: Run {0}" -f (Join-Path $distFolder 'add_to_start_menu.cmd')) -ForegroundColor White
Write-Host "  3. Zip the folder for distribution" -ForegroundColor White