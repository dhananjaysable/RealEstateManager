# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['modern_gui_app.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\Dhanajay.s\\AppData\\Roaming\\Python\\Python313\\site-packages\\customtkinter', 'customtkinter/'), ('D:\\Excel Byforgation\\live work\\live work\\reslivemain', 'reslivemain/'), ('D:\\Excel Byforgation\\live work\\live work\\resvaduvlive', 'resvaduvlive/')],
    hiddenimports=['pandas', 'openpyxl', 'PIL', 'tqdm'],
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
    name='RealEstateManager',
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
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='RealEstateManager',
)
