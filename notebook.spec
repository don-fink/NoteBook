# -*- mode: python ; coding: utf-8 -*-

# PyInstaller spec file for NoteBook application

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Include UI files
        ('*.ui', '.'),
        # Include themes directory
        ('themes', 'themes'),
        # Include schema
        ('schema.sql', '.'),
        # Include icon
        ('scripts/Notebook_icon.ico', 'scripts'),
    ],
    hiddenimports=[
        # PyQt5 modules that might not be auto-detected
        'PyQt5.QtCore',
        'PyQt5.QtWidgets', 
        'PyQt5.QtGui',
        'PyQt5.uic',
        # Database modules
        'sqlite3',
        # Other potentially hidden imports
        'settings_manager',
        'db_access',
        'db_pages', 
        'db_sections',
        'db_version',
        'media_store',
        'ui_loader',
        'ui_logic',
        'ui_richtext',
        'ui_sections',
        'ui_tabs',
        'ui_planning_register',
        'left_tree',
        'page_editor',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Exclude modules we don't need to reduce size
        'tkinter',
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
    ],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='NoteBook',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Hide console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='scripts/Notebook_icon.ico',  # Use your custom icon
    version_file=None,
)