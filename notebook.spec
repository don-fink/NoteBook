# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('*.ui', '.'), ('themes', 'themes'), ('schema.sql', '.'), ('add_to_start_menu.cmd', '.'), ('clear_icon_cache.cmd', '.'), ('README.md', '.')],
    hiddenimports=['PyQt5.QtCore', 'PyQt5.QtWidgets', 'PyQt5.QtGui', 'PyQt5.uic', 'sqlite3', 'settings_manager', 'db_access', 'db_pages', 'db_sections', 'db_version', 'media_store', 'ui_loader', 'ui_logic', 'ui_richtext', 'ui_sections', 'ui_planning_register', 'left_tree', 'page_editor', 'two_pane_core'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'numpy', 'pandas', 'scipy'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='NoteBook',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['scripts\\Notebook_icon.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='NoteBook',
)

# Copy additional files to the root of the distribution folder
import os
import shutil

dist_dir = os.path.join('dist', 'NoteBook')

for file in ['add_to_start_menu.cmd', 'clear_icon_cache.cmd', 'README.md']:
    src = os.path.join('.', file)
    dst = os.path.join(dist_dir, file)
    if os.path.exists(src):
        shutil.copy(src, dst)
