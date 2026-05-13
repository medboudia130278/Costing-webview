# -*- mode: python ; coding: utf-8 -*-
import sys
from PyInstaller.utils.hooks import collect_all, collect_data_files

_icon = ['app.ico']

datas    = []
binaries = []
hiddenimports = ['_cffi_backend']

# Dossier web (HTML / CSS / JS)
datas += [('web', 'web')]

# Assets
datas += [('logo.png', '.')]

# nacl / cffi
tmp = collect_all('nacl')
datas += tmp[0]; binaries += tmp[1]; hiddenimports += tmp[2]
tmp = collect_all('cffi')
datas += tmp[0]; binaries += tmp[1]; hiddenimports += tmp[2]

# PuLP + CBC solver (même fix que la version tkinter)
tmp = collect_all('pulp')
datas += tmp[0]; binaries += tmp[1]; hiddenimports += tmp[2]
binaries += [(
    'C:/Users/medbo/AppData/Roaming/Python/Python313/site-packages/pulp/solverdir/cbc/win/i64/cbc.exe',
    'pulp/solverdir/cbc/win/i64'
)]

# PyWebView + pythonnet
tmp = collect_all('webview')
datas += tmp[0]; binaries += tmp[1]; hiddenimports += tmp[2]
tmp = collect_all('pythonnet')
datas += tmp[0]; binaries += tmp[1]; hiddenimports += tmp[2]

hiddenimports += [
    'webview',
    'webview.platforms.winforms',
    'clr',
    'pythonnet',
]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='MaintenanceControlPanel_WebView',
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
    icon=_icon,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='MaintenanceControlPanel_WebView',
)
