# Cross-Platform Compatibility

> **Note**: This application has been developed and tested on Windows only. Mac and Linux support is theoretical based on code analysis and has **not been tested** on either platform. Community testing and feedback is welcome.

## Overview

NoteBook is a PyQt6 application that uses cross-platform libraries and should run on Mac and Linux with minimal changes. The core application code requires no modifications—only system dependencies and packaging scripts differ by platform.

### What's Already Cross-Platform

| Component | Implementation |
|-----------|----------------|
| Settings storage | Platform-specific paths in `settings_manager.py` |
| UI loading | Qt's `uic` with `os.path` resolution |
| Database | SQLite (works identically everywhere) |
| File operations | Uses `os.path` throughout |
| Desktop integration | `QDesktopServices.openUrl()` |

---

## macOS

### System Dependencies

```bash
# Install Homebrew if not present
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# Install enchant for spell checking
brew install enchant
```

### Python Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install PyQt6 pyenchant
```

### Running

```bash
source .venv/bin/activate
python main.py
```

### Packaging Notes

- Application icon: Convert `scripts/Notebook_icon.ico` to `.icns` format
- PyInstaller spec: Update `icon=` parameter and add `bundle_identifier`
- Create a proper `.app` bundle for distribution

### Known Considerations

- Qt Designer setup is best handled via Qt's official installer on Mac
- Settings stored in: `~/Library/Application Support/NoteBook/`

---

## Linux

### System Dependencies

**Debian/Ubuntu:**
```bash
sudo apt install python3-venv libenchant-2-dev aspell-en
```

**Fedora/RHEL:**
```bash
sudo dnf install python3-devel enchant2-devel aspell-en
```

**Arch Linux:**
```bash
sudo pacman -S python enchant aspell-en
```

### Python Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install PyQt6 pyenchant
```

### Running

```bash
source .venv/bin/activate
python main.py
```

### Desktop Integration

Create a `.desktop` file for menu integration:

```ini
# Save as: ~/.local/share/applications/notebook.desktop
[Desktop Entry]
Name=NoteBook
Comment=Note-taking application
Exec=/path/to/NoteBook/.venv/bin/python /path/to/NoteBook/main.py
Icon=/path/to/NoteBook/scripts/notebook_icon.png
Type=Application
Categories=Office;Utility;
Terminal=false
```

Then update the desktop database:
```bash
update-desktop-database ~/.local/share/applications/
```

### Packaging Notes

- Application icon: Use `.png` or `.svg` format (not `.ico`)
- PyInstaller works on Linux; remove Windows-specific icon references from spec
- Consider AppImage or Flatpak for distribution

### Known Considerations

- Install Qt Designer via system packages instead of Windows-only helper bundles:
  ```bash
  sudo apt install qt6-tools-dev-tools  # Debian/Ubuntu
  ```
- Settings stored in: `~/.config/NoteBook/`

---

## Development Dependencies

The current `requirements-dev.txt` already targets PyQt6. For cross-platform development, install Qt Designer separately using the official Qt installer or your platform package manager.

---

## Contributing

If you test NoteBook on Mac or Linux, please:

1. Open an issue reporting your experience
2. Note your OS version and Python version
3. Document any issues encountered and workarounds found

Your feedback will help improve cross-platform support!
