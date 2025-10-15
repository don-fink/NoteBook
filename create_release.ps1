# NoteBook Distribution Package Creator
# Creates a clean distribution folder with the PyInstaller executable

$ErrorActionPreference = 'Stop'

$distFolder = 'NoteBook_Release'
Write-Host "Creating PyInstaller distribution package..." -ForegroundColor Cyan

# Clean and create distribution folder
if (Test-Path $distFolder) { Remove-Item -Recurse -Force $distFolder }
New-Item -ItemType Directory $distFolder | Out-Null

# Copy the executable
Copy-Item 'dist\NoteBook.exe' $distFolder
Write-Host "  ✓ NoteBook.exe" -ForegroundColor Green

# Create user-friendly README
$readmeContent = @'
# NoteBook - Rich Text Note-Taking Application

## Quick Start

1. Double-click `NoteBook.exe` to launch the application
2. No Python installation required!
3. On first run, create a new database or open an existing one

## Features

- Multiple notebooks with sections and pages
- Rich text editing with formatting
- Image and video support with thumbnails
- Planning registers and tables
- Multiple themes
- Cross-platform database format

## System Requirements

- Windows 10 or Windows 11
- No additional software required

## First Time Setup

When you first run NoteBook:
1. Choose "New Database" to create your first notebook collection
2. Pick a location to store your notebooks (recommended: Documents folder)
3. Start creating notebooks, sections, and pages!

## Data Location

Your notebooks are stored as SQLite database files (.db) that you can:
- Back up easily
- Share between computers
- Open from any location

## Uninstall

Simply delete the NoteBook.exe file - no registry entries or system files are created.

## Support

This is a portable application that doesn't modify your system.
All settings and data are stored in your user profile.

Enjoy taking notes!
'@

Set-Content -Path "$distFolder\README.txt" -Value $readmeContent -Encoding UTF8
Write-Host "  ✓ README.txt" -ForegroundColor Green

# Get executable size
$exeSize = [math]::Round((Get-Item "$distFolder\NoteBook.exe").Length / 1MB, 1)

Write-Host ""
Write-Host "Distribution package created!" -ForegroundColor Green
Write-Host "Location: $distFolder" -ForegroundColor Cyan
Write-Host "Executable size: $exeSize MB" -ForegroundColor Cyan
Write-Host ""
Write-Host "Ready to distribute:" -ForegroundColor Yellow
Write-Host "  • Zip the '$distFolder' folder" -ForegroundColor White
Write-Host "  • Users just extract and run NoteBook.exe" -ForegroundColor White
Write-Host "  • No Python installation needed!" -ForegroundColor White
Write-Host ""

# Optional: Create ZIP
$createZip = Read-Host "Create zip file for distribution? (Y/n)"
if ($createZip -notmatch '^n|N|no|No$') {
    $zipPath = "$distFolder.zip"
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
    
    try {
        Compress-Archive -Path "$distFolder\*" -DestinationPath $zipPath -CompressionLevel Optimal
        $zipSize = [math]::Round((Get-Item $zipPath).Length / 1MB, 1)
        Write-Host "  ✓ $zipPath ($zipSize MB)" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to create zip: $($_.Exception.Message)"
    }
}

Write-Host ""
Write-Host "Test the executable:" -ForegroundColor Yellow
Write-Host "  Double-click: $distFolder\NoteBook.exe" -ForegroundColor White