<#
 NoteBook Windows Installation Script
 ------------------------------------
 Professional installer for the NoteBook application.

 What it does:
 1. Detects or prompts to install Python 3.11+ (with download assistance).
 2. Allows user to choose installation directory (defaults to %LOCALAPPDATA%\Programs\NoteBook).
 3. Copies application files to the chosen directory.
 4. Creates a virtual environment and installs dependencies.
 5. Creates Start Menu entry and optionally desktop shortcut.
 6. Associates custom icon with shortcuts.
 7. Generates uninstaller for clean removal.

 Usage (PowerShell):
   ./scripts/install_windows.ps1

 Parameters:
   -InstallPath <path> : Custom installation directory (optional).
   -Force : Recreate virtual environment and overwrite existing installation.
   -NoStartMenu : Skip Start Menu entry creation.
   -NoDesktopPrompt : Skip asking about desktop shortcut.
   -Quiet : Minimal prompts (uses defaults).
   -Verbose : Show detailed progress.
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

function Write-Info($msg) { Write-Host "[INFO] $msg" -ForegroundColor Cyan }
function Write-Warn($msg) { Write-Host "[WARN] $msg" -ForegroundColor Yellow }
function Write-Err($msg)  { Write-Host "[ERROR] $msg" -ForegroundColor Red }
function Write-Success($msg) { Write-Host "[SUCCESS] $msg" -ForegroundColor Green }

$ScriptRoot = $PSScriptRoot
$SourceRoot = (Resolve-Path (Join-Path $ScriptRoot '..')).Path

Write-Host ""
Write-Host "===============================================" -ForegroundColor Magenta
Write-Host "       NoteBook Installation Wizard" -ForegroundColor Magenta  
Write-Host "===============================================" -ForegroundColor Magenta
Write-Host ""

# --- Determine installation directory ---
if (-not $InstallPath) {
    $defaultPath = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'Programs\NoteBook'
    if (-not $Quiet) {
        Write-Host "Choose installation directory:"
        Write-Host "  1. User Programs (recommended): $defaultPath"
        Write-Host "  2. Custom location"
        Write-Host ""
        do {
            $choice = Read-Host "Enter choice (1 or 2) [1]"
            if ([string]::IsNullOrWhiteSpace($choice)) { $choice = "1" }
        } while ($choice -notin @("1", "2"))
        
        if ($choice -eq "2") {
            do {
                $InstallPath = Read-Host "Enter installation path"
            } while ([string]::IsNullOrWhiteSpace($InstallPath))
        } else {
            $InstallPath = $defaultPath
        }
    } else {
        $InstallPath = $defaultPath
    }
}

Write-Info "Installation target: $InstallPath"

# --- 1. Enhanced Python detection with download assistance ---
function Get-Python311Plus {
    # Try py launcher for 3.11+
    foreach ($ver in @('3.13', '3.12', '3.11')) {
        try {
            $output = (& py "-$ver" --version 2>$null)
            if ($LASTEXITCODE -eq 0) { 
                Write-Info "Found Python $ver via py launcher"
                return "py -$ver", $ver
            }
        } catch {}
    }
    # Try system python
    try {
        $verLine = (& python --version 2>&1)
        if ($verLine -match 'Python 3\.(\d+)\.') {
            $minorVer = [int]$matches[1]
            if ($minorVer -ge 11) { 
                Write-Info "Found compatible system Python: $verLine"
                return 'python', $verLine
            }
        }
    } catch {}
    return $null, $null
}

function Prompt-PythonInstall {
    Write-Warn 'Python 3.11+ was not found on this system.'
    Write-Host ""
    Write-Host "NoteBook requires Python 3.11 or newer. Please:"
    Write-Host "  1. Visit: https://www.python.org/downloads/windows/"
    Write-Host "  2. Download Python 3.11+ (recommended: latest stable)"  
    Write-Host "  3. During installation, check 'Add Python to PATH'"
    Write-Host "  4. Restart this installer after Python is installed"
    Write-Host ""
    
    if (-not $Quiet) {
        $openBrowser = Read-Host "Open download page in browser? (y/N)"
        if ($openBrowser -eq 'y' -or $openBrowser -eq 'Y') {
            try {
                Start-Process "https://www.python.org/downloads/windows/"
            } catch {
                Write-Warn "Could not open browser. Please visit the URL manually."
            }
        }
    }
    
    Write-Host "Installation cannot continue without Python 3.11+." -ForegroundColor Red
    exit 1
}

$pythonCmd, $pythonVer = Get-Python311Plus
if (-not $pythonCmd) {
    Prompt-PythonInstall
}
Write-Success "Using Python: $pythonCmd ($pythonVer)"

# Normalize interpreter: if py launcher, build a command to run modules
if ($pythonCmd -eq 'py -3.11') {
    # Store base executable and argument separately so invocations use call operator with array
    $pythonExe = 'py'
    $pythonArgs = @('-3.11')
} else {
    $pythonExe = 'python'
    $pythonArgs = @()
}

# --- 2. (Optional) Recreate venv ---
$venvPath = Join-Path $ProjectRoot '.venv'
if ($Force -and (Test-Path $venvPath)) {
    Write-Warn 'Removing existing virtual environment due to -Force.'
    Remove-Item -Recurse -Force $venvPath
}

if (-not (Test-Path $venvPath)) {
    Write-Info 'Creating virtual environment (.venv)'
    if ($pythonArgs.Count -gt 0) {
        & $pythonExe @pythonArgs -m venv .venv
    } else {
        & $pythonExe -m venv .venv
    }
} else {
    Write-Info 'Virtual environment already exists; skipping creation.'
}

# Paths inside venv
$venvPython = Join-Path $venvPath 'Scripts/python.exe'
$venvPythonW = Join-Path $venvPath 'Scripts/pythonw.exe'
if (-not (Test-Path $venvPython)) {
    Write-Err 'Virtual environment python.exe not found; creation may have failed.'
    exit 1
}

# --- 3. Upgrade pip and install dependencies ---
Write-Info 'Upgrading pip'
& $venvPython -m pip install --upgrade pip | Out-Null

# Decide which requirements file to use (runtime only)
$reqFile = Join-Path $ProjectRoot 'requirements.txt'
if (-not (Test-Path $reqFile)) {
    Write-Err 'requirements.txt not found.'
    exit 1
}
Write-Info 'Installing runtime dependencies'
& $venvPython -m pip install -r $reqFile

# --- 4. Create launcher batch file ---
$launcher = Join-Path $ProjectRoot 'launch_notebook.cmd'
Write-Info "Creating launcher $launcher"
@"
@echo off
REM Auto-generated launcher for NoteBook
setlocal
cd /d %~dp0
IF NOT EXIST .venv (echo Virtual environment missing. Run scripts\install_windows.ps1 & exit /b 1)
call .venv\Scripts\activate.bat
"%~dp0.venv\Scripts\python.exe" "%~dp0main.py" %*
endlocal
"@ | Set-Content -Encoding ASCII $launcher

# --- 5. Desktop shortcut (use pythonw.exe to avoid console window) ---
if (-not $NoShortcut) {
    try {
        $desktop = [Environment]::GetFolderPath('Desktop')
        $shortcutPath = Join-Path $desktop 'NoteBook.lnk'
        Write-Info "Creating desktop shortcut: $shortcutPath"
        $wsh = New-Object -ComObject WScript.Shell
        $sc = $wsh.CreateShortcut($shortcutPath)
        # Point directly to pythonw.exe so no console window appears
        $sc.TargetPath = $venvPythonW
        $sc.Arguments = '"' + (Join-Path $ProjectRoot 'main.py') + '"'
        $sc.WorkingDirectory = $ProjectRoot
        $sc.IconLocation = "$venvPython,0"
        $sc.WindowStyle = 1
        $sc.Description = 'NoteBook (PyQt5)'
        $sc.Save()
    } catch {
        Write-Warn "Failed to create desktop shortcut: $($_.Exception.Message)"
    }
} else {
    Write-Info 'Skipping desktop shortcut creation (-NoShortcut specified).'
}

Write-Host "`nInstallation complete." -ForegroundColor Green
Write-Host 'Launch via:  launch_notebook.cmd  (or the desktop shortcut).' -ForegroundColor Green
Write-Host 'Re-run this script with -Force to recreate the environment if needed.' -ForegroundColor Green

Pop-Location