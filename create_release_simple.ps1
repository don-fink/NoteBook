# NoteBook Distribution Package Creator
$ErrorActionPreference = 'Stop'

$distFolder = 'NoteBook_Release'
Write-Host "Creating PyInstaller distribution package..." -ForegroundColor Cyan

# Clean and create distribution folder
if (Test-Path $distFolder) { Remove-Item -Recurse -Force $distFolder }
New-Item -ItemType Directory $distFolder | Out-Null

# Copy the executable
Copy-Item 'dist\NoteBook.exe' $distFolder
Write-Host "  Copied NoteBook.exe" -ForegroundColor Green

# Copy optional Start Menu installer
Copy-Item 'add_to_start_menu.cmd' $distFolder
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

Set-Content -Path "$distFolder\README.txt" -Value $readme -Encoding UTF8
Write-Host "  Created README.txt" -ForegroundColor Green

# Show size info
$exeSize = [math]::Round((Get-Item "$distFolder\NoteBook.exe").Length / 1MB, 1)
Write-Host ""
Write-Host "Distribution ready!" -ForegroundColor Green
Write-Host "Location: $distFolder" -ForegroundColor Cyan
Write-Host "Size: $exeSize MB" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "  1. Test: Double-click $distFolder\NoteBook.exe" -ForegroundColor White
Write-Host "  2. Optional: Run $distFolder\add_to_start_menu.cmd" -ForegroundColor White
Write-Host "  3. Zip the folder for distribution" -ForegroundColor White