# -*- mode: python ; coding: utf-8 -*-

# Collect pyenchant data files (DLLs and dictionaries)
from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs
import os

# Get enchant package location
import enchant
enchant_pkg_dir = os.path.dirname(enchant.__file__)
enchant_data_dir = os.path.join(enchant_pkg_dir, 'data')

# Manually collect all enchant data
enchant_datas = []
enchant_binaries = []

if os.path.exists(enchant_data_dir):
    for root, dirs, files in os.walk(enchant_data_dir):
        for f in files:
            src = os.path.join(root, f)
            # Destination path relative to enchant package
            rel_path = os.path.relpath(root, enchant_pkg_dir)
            dest = os.path.join('enchant', rel_path)
            if f.endswith('.dll'):
                enchant_binaries.append((src, dest))
            else:
                enchant_datas.append((src, dest))

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=enchant_binaries,
    datas=[('main_window_2_column.ui', '.'), ('settings_dialog.ui', '.'), ('themes', 'themes'), ('schema.sql', '.'), ('add_to_start_menu.cmd', '.'), ('clear_icon_cache.cmd', '.'), ('README.md', '.')] + enchant_datas,
    hiddenimports=['PyQt5.QtCore', 'PyQt5.QtWidgets', 'PyQt5.QtGui', 'PyQt5.uic', 'sqlite3', 'settings_manager', 'db_access', 'db_pages', 'db_sections', 'db_version', 'media_store', 'ui_loader', 'ui_logic', 'ui_richtext', 'ui_sections', 'ui_planning_register', 'left_tree', 'page_editor', 'two_pane_core', 'spell_check', 'enchant'],
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
