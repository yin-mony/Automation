# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

block_cipher = None

ddddocr_datas, ddddocr_binaries, ddddocr_hiddenimports = collect_all('ddddocr')
dp_datas, dp_binaries, dp_hiddenimports = collect_all('DrissionPage')

a = Analysis(
    ['run.py'],
    pathex=[],
    binaries=ddddocr_binaries + dp_binaries,
    datas=ddddocr_datas + dp_datas,
    hiddenimports=[
        'openpyxl',
        'main',
        'NewSet',
        'Variant',
        'SaihuERPLogin',
        *ddddocr_hiddenimports,
        *dp_hiddenimports,
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='赛狐ERP自动化',
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
