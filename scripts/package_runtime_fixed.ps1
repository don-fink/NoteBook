param(
    [string]$Destination = 'NoteBook_Distribution',
    [switch]$CreateZip
)

$ErrorActionPreference = 'Stop'
Write-Host "Creating distribution package..." -ForegroundColor Cyan

# Clean and create destination
if (Test-Path $Destination) { Remove-Item -Recurse -Force $Destination }
New-Item -ItemType Directory -Path $Destination | Out-Null

# Core application files
$files = @(
    'main.py','db_access.py','db_pages.py','db_sections.py','db_version.py',
    'media_store.py','settings_manager.py','ui_loader.py','ui_logic.py','ui_richtext.py',
    'ui_sections.py','ui_tabs.py','ui_planning_register.py','left_tree.py','page_editor.py',
    'schema.sql','requirements.txt'
)

Write-Host "Copying application files..." -ForegroundColor Yellow
foreach($f in $files){
    if(Test-Path $f){ 
        Copy-Item $f -Destination (Join-Path $Destination $f) -Force 
        Write-Host "  ✓ $f" -ForegroundColor Green
    } else { 
        Write-Warning "Missing: $f" 
    }
}

# UI files
Write-Host "Copying UI files..." -ForegroundColor Yellow
Get-ChildItem -Filter '*.ui' -File | ForEach-Object {
    Copy-Item $_.FullName -Destination (Join-Path $Destination $_.Name) -Force
    Write-Host "  ✓ $($_.Name)" -ForegroundColor Green
}

# Themes directory
if(Test-Path 'themes'){
    Copy-Item 'themes' -Destination (Join-Path $Destination 'themes') -Recurse -Force
    Write-Host "  ✓ themes/" -ForegroundColor Green
}

# Scripts and installer
Write-Host "Copying installer..." -ForegroundColor Yellow
if(-not (Test-Path (Join-Path $Destination 'scripts'))){
    New-Item -ItemType Directory -Path (Join-Path $Destination 'scripts') | Out-Null
}

$installerFiles = @('install_professional.ps1', 'Notebook_icon.ico')
foreach($file in $installerFiles) {
    $src = Join-Path 'scripts' $file
    if(Test-Path $src){
        $dst = Join-Path $Destination "scripts/$file"
        Copy-Item $src $dst -Force
        Write-Host "  ✓ scripts/$file" -ForegroundColor Green
    }
}

# Create simple launcher batch for manual runs
$launcherContent = @'
@echo off
rem Simple launcher for development/testing
if not exist .venv (
    echo Virtual environment not found. Please run the installer first.
    echo   PowerShell: scripts\install_professional.ps1
    pause
    exit /b 1
)
call .venv\Scripts\activate.bat
"%~dp0.venv\Scripts\pythonw.exe" "%~dp0main.py" %*
'@

Set-Content -Path (Join-Path $Destination 'launch.cmd') -Value $launcherContent -Encoding ASCII
Write-Host "  ✓ launch.cmd" -ForegroundColor Green

# Create README for distribution
$readmeContent = @'
# NoteBook Application

## Installation

1. Extract this archive to a temporary folder
2. Run installer (right-click, Run with PowerShell):
   scripts\install_professional.ps1
3. Follow the wizard to choose installation location and options
4. Launch from Start Menu or desktop shortcut

## Requirements

- Windows 10/11
- Python 3.11+ (installer will guide you if missing)

## Manual Installation

If the installer doesn't work:

1. Install Python 3.11+ from python.org (check "Add to PATH")
2. Extract files to desired location
3. Open PowerShell in that folder:
   python -m venv .venv
   .\.venv\Scripts\activate
   pip install -r requirements.txt
4. Run: launch.cmd

## Uninstall

Run uninstall.ps1 in the installation directory, or use Add/Remove Programs.

## Support

This is a portable rich-text note-taking application with:
- Multiple notebooks and sections
- Rich text editing with images/videos
- Planning registers and tables
- Theme support

Enjoy!
'@

Set-Content -Path (Join-Path $Destination 'README.txt') -Value $readmeContent -Encoding UTF8
Write-Host "  ✓ README.txt" -ForegroundColor Green

# Create zip if requested
if ($CreateZip) {
    $zipPath = "$Destination.zip"
    Write-Host "Creating zip archive..." -ForegroundColor Yellow
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
    
    try {
        Compress-Archive -Path "$Destination\*" -DestinationPath $zipPath -CompressionLevel Optimal
        Write-Host "  ✓ $zipPath" -ForegroundColor Green
        
        # Show size info
        $size = [math]::Round((Get-Item $zipPath).Length / 1MB, 1)
        Write-Host "  Size: $size MB" -ForegroundColor Cyan
    } catch {
        Write-Warning "Failed to create zip: $($_.Exception.Message)"
    }
}

Write-Host ""
Write-Host "Distribution package ready!" -ForegroundColor Green
Write-Host "Location: $Destination" -ForegroundColor Cyan
if ($CreateZip -and (Test-Path "$Destination.zip")) {
    Write-Host "Archive: $Destination.zip" -ForegroundColor Cyan
}
Write-Host ""
Write-Host "To test:" -ForegroundColor Yellow
Write-Host "  1. Copy '$Destination' folder to another location" -ForegroundColor White
Write-Host "  2. Run: scripts\install_professional.ps1" -ForegroundColor White