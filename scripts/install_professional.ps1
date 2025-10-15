<#
 NoteBook Professional Windows Installer
 ========================================
 
 Complete installation wizard for NoteBook application.
 
 Features:
 - Python 3.11+ detection with download assistance
 - Custom installation directory selection  
 - Start Menu integration
 - Desktop shortcut (optional)
 - Custom icon support
 - Clean uninstaller generation
 - Professional user experience
 
 Usage: Run from the distribution folder containing the app files
   powershell -ExecutionPolicy Bypass -File install_professional.ps1
#>

param(
    [string]$InstallPath = "",
    [switch]$Force,
    [switch]$NoStartMenu, 
    [switch]$NoDesktopPrompt,
    [switch]$Quiet,
    [switch]$Verbose
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --- Helper Functions ---
function Write-Info($msg) { Write-Host "[INFO] $msg" -ForegroundColor Cyan }
function Write-Warn($msg) { Write-Host "[WARN] $msg" -ForegroundColor Yellow }
function Write-Err($msg)  { Write-Host "[ERROR] $msg" -ForegroundColor Red }
function Write-Success($msg) { Write-Host "[SUCCESS] $msg" -ForegroundColor Green }
function Write-Step($msg) { Write-Host "[STEP] $msg" -ForegroundColor Magenta }

$ScriptRoot = $PSScriptRoot
$SourceRoot = (Resolve-Path (Join-Path $ScriptRoot '..')).Path
$AppName = "NoteBook"
$AppDescription = "Rich text note-taking application"

# --- Welcome Banner ---
Write-Host ""
Write-Host "==========================================================" -ForegroundColor Magenta
Write-Host "           $AppName Installation Wizard" -ForegroundColor Magenta  
Write-Host "==========================================================" -ForegroundColor Magenta
Write-Host ""
Write-Info "This installer will set up $AppName on your computer."
Write-Host ""

# --- Step 1: Python Detection ---
Write-Step "Checking for Python 3.11+"

function Get-Python311Plus {
    # Try py launcher for recent versions
    foreach ($ver in @('3.13', '3.12', '3.11')) {
        try {
            $output = (& py "-$ver" --version 2>$null)
            if ($LASTEXITCODE -eq 0) { 
                return @{Cmd="py"; Args=@("-$ver"); Version=$ver; DisplayName="Python $ver (py launcher)"}
            }
        } catch {}
    }
    # Try system python
    try {
        $verLine = (& python --version 2>&1)
        if ($verLine -match 'Python 3\.(\d+)\.(\d+)') {
            $major = [int]$matches[1] 
            $minor = [int]$matches[2]
            if ($major -eq 3 -and $minor -ge 11) { 
                return @{Cmd="python"; Args=@(); Version="3.$minor"; DisplayName=$verLine.Trim()}
            }
        }
    } catch {}
    return $null
}

function Prompt-PythonInstall {
    Write-Warn 'Python 3.11+ is required but not found.'
    Write-Host ""
    Write-Host "Please install Python from the official website:" -ForegroundColor Yellow
    Write-Host "  • Visit: https://www.python.org/downloads/windows/" -ForegroundColor White
    Write-Host "  • Download Python 3.11 or newer" -ForegroundColor White  
    Write-Host "  • During installation, check 'Add Python to PATH'" -ForegroundColor White
    Write-Host "  • Restart this installer after installation" -ForegroundColor White
    Write-Host ""
    
    if (-not $Quiet) {
        $openBrowser = Read-Host "Open download page in browser? (y/N)"
        if ($openBrowser -match '^y|Y$') {
            try {
                Start-Process "https://www.python.org/downloads/windows/"
                Write-Info "Opened browser. Please download and install Python, then re-run this installer."
            } catch {
                Write-Warn "Could not open browser. Please visit the URL manually."
            }
        }
    }
    
    exit 1
}

$pythonInfo = Get-Python311Plus
if (-not $pythonInfo) {
    Prompt-PythonInstall
}
Write-Success "Found: $($pythonInfo.DisplayName)"

# --- Step 2: Installation Directory ---
Write-Step "Selecting installation directory"

if (-not $InstallPath) {
    $defaultPath = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) "Programs\$AppName"
    
    if (-not $Quiet) {
        Write-Host "Choose installation location:" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  1. User folder (RECOMMENDED): $defaultPath" -ForegroundColor Green
        Write-Host "     • No administrator rights needed" -ForegroundColor Gray
        Write-Host "     • Private to your user account" -ForegroundColor Gray
        Write-Host "     • Easy updates and removal" -ForegroundColor Gray
        Write-Host ""
        Write-Host "  2. Custom location" -ForegroundColor Yellow
        Write-Host "     • Choose your own directory" -ForegroundColor Gray
        Write-Host "     • May require admin rights in some locations" -ForegroundColor Gray
        Write-Host ""
        
        do {
            $choice = Read-Host "Enter choice (1 or 2) [1 - Recommended]"
            if ([string]::IsNullOrWhiteSpace($choice)) { $choice = "1" }
        } while ($choice -notin @("1", "2"))
        
        if ($choice -eq "2") {
            Write-Host ""
            Write-Host "Examples of custom locations:" -ForegroundColor Gray
            Write-Host "  • C:\Program Files\$AppName (requires admin)" -ForegroundColor Gray  
            Write-Host "  • C:\Tools\$AppName" -ForegroundColor Gray
            Write-Host "  • D:\Applications\$AppName" -ForegroundColor Gray
            Write-Host ""
            do {
                $customPath = Read-Host "Enter full installation path"
                if (-not [string]::IsNullOrWhiteSpace($customPath)) {
                    $InstallPath = $customPath
                    break
                }
                Write-Host "Please enter a valid directory path." -ForegroundColor Red
            } while ($true)
        } else {
            $InstallPath = $defaultPath
            Write-Host ""
            Write-Info "Using recommended user location (no admin rights needed)"
        }
    } else {
        $InstallPath = $defaultPath
    }
}

# Test write permissions
try {
    $testDir = New-Item -ItemType Directory -Path $InstallPath -Force -ErrorAction Stop
    Write-Success "Installation directory: $InstallPath"
} catch {
    Write-Err "Cannot create installation directory: $InstallPath"
    Write-Err "Error: $($_.Exception.Message)"
    exit 1
}

# --- Step 3: Copy Application Files ---
Write-Step "Copying application files"

$filesToCopy = @(
    'main.py', 'db_access.py', 'db_pages.py', 'db_sections.py', 'db_version.py',
    'media_store.py', 'settings_manager.py', 'ui_loader.py', 'ui_logic.py', 
    'ui_richtext.py', 'ui_sections.py', 'ui_tabs.py', 'ui_planning_register.py',
    'left_tree.py', 'page_editor.py', 'schema.sql', 'requirements.txt'
)

foreach ($file in $filesToCopy) {
    $src = Join-Path $SourceRoot $file
    $dst = Join-Path $InstallPath $file
    if (Test-Path $src) {
        Copy-Item $src $dst -Force
        if ($Verbose) { Write-Info "Copied: $file" }
    } else {
        Write-Warn "Missing source file: $file"
    }
}

# Copy UI files
Get-ChildItem -Path $SourceRoot -Filter "*.ui" | ForEach-Object {
    $dst = Join-Path $InstallPath $_.Name
    Copy-Item $_.FullName $dst -Force
    if ($Verbose) { Write-Info "Copied: $($_.Name)" }
}

# Copy themes directory
$themesSource = Join-Path $SourceRoot "themes"
$themesTarget = Join-Path $InstallPath "themes"
if (Test-Path $themesSource) {
    Copy-Item $themesSource $themesTarget -Recurse -Force
    Write-Info "Copied themes directory"
}

# Copy icon
$iconSource = Join-Path $ScriptRoot "Notebook_icon.ico" 
$iconTarget = Join-Path $InstallPath "Notebook_icon.ico"
if (Test-Path $iconSource) {
    Copy-Item $iconSource $iconTarget -Force
    Write-Info "Copied application icon"
}

Write-Success "Application files copied successfully"

# --- Step 4: Create Virtual Environment ---
Write-Step "Setting up Python virtual environment"

$venvPath = Join-Path $InstallPath '.venv'
if ($Force -and (Test-Path $venvPath)) {
    Write-Info "Removing existing virtual environment (Force mode)"
    Remove-Item -Recurse -Force $venvPath
}

if (-not (Test-Path $venvPath)) {
    Write-Info "Creating virtual environment..."
    if ($pythonInfo.Args.Count -gt 0) {
        & $pythonInfo.Cmd @($pythonInfo.Args) -m venv $venvPath
    } else {
        & $pythonInfo.Cmd -m venv $venvPath
    }
    
    if ($LASTEXITCODE -ne 0) {
        Write-Err "Failed to create virtual environment"
        exit 1
    }
} else {
    Write-Info "Virtual environment already exists"
}

$venvPython = Join-Path $venvPath 'Scripts\python.exe'
$venvPythonW = Join-Path $venvPath 'Scripts\pythonw.exe'

if (-not (Test-Path $venvPython)) {
    Write-Err "Virtual environment creation failed - python.exe not found"
    exit 1
}

# --- Step 5: Install Dependencies ---
Write-Step "Installing dependencies"

Write-Info "Upgrading pip..."
& $venvPython -m pip install --upgrade pip --quiet

$reqFile = Join-Path $InstallPath 'requirements.txt'
if (Test-Path $reqFile) {
    Write-Info "Installing packages from requirements.txt..."
    & $venvPython -m pip install -r $reqFile --quiet
    if ($LASTEXITCODE -ne 0) {
        Write-Err "Failed to install dependencies"
        exit 1
    }
    Write-Success "Dependencies installed successfully"
} else {
    Write-Warn "requirements.txt not found - skipping package installation"
}

# --- Step 6: Create Start Menu Entry ---
if (-not $NoStartMenu) {
    Write-Step "Creating Start Menu entry"
    
    try {
        $startMenuPath = Join-Path ([Environment]::GetFolderPath('StartMenu')) "Programs\$AppName.lnk"
        $wsh = New-Object -ComObject WScript.Shell
        $shortcut = $wsh.CreateShortcut($startMenuPath)
        $shortcut.TargetPath = $venvPythonW
        $shortcut.Arguments = "`"$(Join-Path $InstallPath 'main.py')`""
        $shortcut.WorkingDirectory = $InstallPath
        $shortcut.Description = $AppDescription
        if (Test-Path $iconTarget) {
            $shortcut.IconLocation = "$iconTarget,0"
        }
        $shortcut.Save()
        Write-Success "Start Menu entry created"
    } catch {
        Write-Warn "Failed to create Start Menu entry: $($_.Exception.Message)"
    }
}

# --- Step 7: Desktop Shortcut (Optional) ---
if (-not $NoDesktopPrompt) {
    Write-Step "Desktop shortcut"
    
    $createDesktop = $true
    if (-not $Quiet) {
        $response = Read-Host "Create desktop shortcut? (Y/n)"
        $createDesktop = -not ($response -match '^n|N|no|No$')
    }
    
    if ($createDesktop) {
        try {
            $desktopPath = Join-Path ([Environment]::GetFolderPath('Desktop')) "$AppName.lnk"
            $wsh = New-Object -ComObject WScript.Shell
            $shortcut = $wsh.CreateShortcut($desktopPath)
            $shortcut.TargetPath = $venvPythonW
            $shortcut.Arguments = "`"$(Join-Path $InstallPath 'main.py')`""
            $shortcut.WorkingDirectory = $InstallPath
            $shortcut.Description = $AppDescription
            if (Test-Path $iconTarget) {
                $shortcut.IconLocation = "$iconTarget,0"
            }
            $shortcut.Save()
            Write-Success "Desktop shortcut created"
        } catch {
            Write-Warn "Failed to create desktop shortcut: $($_.Exception.Message)"
        }
    }
}

# --- Step 8: Create Uninstaller ---
Write-Step "Creating uninstaller"

$uninstallerPath = Join-Path $InstallPath "uninstall.ps1"
$uninstallerContent = @"
# $AppName Uninstaller
param([switch]`$Quiet)

`$InstallPath = Split-Path -Parent `$MyInvocation.MyCommand.Path
`$AppName = "$AppName"

if (-not `$Quiet) {
    `$confirm = Read-Host "Remove `$AppName and all its files? (y/N)"
    if (`$confirm -notmatch '^y|Y|yes|Yes$') {
        Write-Host "Uninstall cancelled."
        exit 0
    }
}

Write-Host "Removing `$AppName..." -ForegroundColor Yellow

# Remove shortcuts
try {
    `$startMenu = Join-Path ([Environment]::GetFolderPath('StartMenu')) "Programs\`$AppName.lnk"
    if (Test-Path `$startMenu) { Remove-Item `$startMenu -Force }
    
    `$desktop = Join-Path ([Environment]::GetFolderPath('Desktop')) "`$AppName.lnk"  
    if (Test-Path `$desktop) { Remove-Item `$desktop -Force }
} catch {}

# Remove installation directory
try {
    Set-Location ([Environment]::GetFolderPath('UserProfile'))
    Remove-Item -Recurse -Force "`$InstallPath"
    Write-Host "`$AppName has been removed." -ForegroundColor Green
} catch {
    Write-Host "Error removing files: `$(`$_.Exception.Message)" -ForegroundColor Red
    Write-Host "You may need to manually delete: `$InstallPath" -ForegroundColor Yellow
}

if (-not `$Quiet) {
    Read-Host "Press Enter to close"
}
"@

Set-Content -Path $uninstallerPath -Value $uninstallerContent -Encoding UTF8
Write-Success "Uninstaller created: uninstall.ps1"

# --- Installation Complete ---
Write-Host ""
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "           Installation Complete!" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
Write-Host ""
Write-Success "$AppName has been installed to: $InstallPath"
Write-Host ""
Write-Host "Launch options:" -ForegroundColor Cyan
if (-not $NoStartMenu) {
    Write-Host "  • Start Menu: Search for '$AppName'" -ForegroundColor White
}
if (-not $NoDesktopPrompt -and (Test-Path (Join-Path ([Environment]::GetFolderPath('Desktop')) "$AppName.lnk"))) {
    Write-Host "  • Desktop: Double-click the $AppName icon" -ForegroundColor White  
}
Write-Host "  • Direct: Run `"$venvPythonW`" `"$(Join-Path $InstallPath 'main.py')`"" -ForegroundColor White
Write-Host ""
Write-Host "To uninstall: Run uninstall.ps1 in the installation directory" -ForegroundColor Yellow
Write-Host ""

if (-not $Quiet) {
    $launch = Read-Host "Launch $AppName now? (Y/n)"
    if (-not ($launch -match '^n|N|no|No$')) {
        try {
            Start-Process -FilePath $venvPythonW -ArgumentList "`"$(Join-Path $InstallPath 'main.py')`"" -WorkingDirectory $InstallPath
            Write-Success "$AppName is starting..."
        } catch {
            Write-Warn "Could not launch automatically. Use the Start Menu or desktop shortcut."
        }
    }
}