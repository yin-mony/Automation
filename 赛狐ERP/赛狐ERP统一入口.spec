# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['OnlyRun.py'],
    pathex=['c:\\Users\\admin\\Desktop\\Automation\\赛狐ERP'],
    binaries=[],
    datas=[],
    hiddenimports=['OnlyMain', 'LowMain', 'DewMain', 'SaihuERPLogin'],
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
    a.binaries,
    a.datas,
    [],
    name='赛狐ERP统一入口',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
