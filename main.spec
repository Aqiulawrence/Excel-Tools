# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['ImageSearch.py', 'ExcelCompare.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe1 = EXE(
    pyz,
    a.scripts[:1],
    [],
    exclude_binaries=True,
    name='ImageSearch',
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
exe2 = EXE(
    pyz,
    a.scripts[1:],
    [],
    exclude_binaries=True,
    name='ExcelCompare',
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
    exe1,
	exe2,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Excel-Tools',
)
