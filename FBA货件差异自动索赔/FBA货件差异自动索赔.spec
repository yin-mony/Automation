# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all, collect_submodules

ddddocr_datas, ddddocr_binaries, ddddocr_hiddenimports = collect_all('ddddocr')

hiddenimports = [
    'main',
    'test',
    'auto',
    'export',
    'email_util',
    'SaihuERPLogin',
    'YidekeLogin',
    'docx',
    'docx2pdf',
    'DrissionPage',
    'psutil',
    'pywinauto',
]
hiddenimports += collect_submodules('DrissionPage')
hiddenimports += collect_submodules('docx')
hiddenimports += ddddocr_hiddenimports

a = Analysis(
    ['run.py'],
    pathex=[],
    binaries=ddddocr_binaries,
    datas=[
        ('db53060fa183_发票模板.docx', '.'),
        ('服务商模板.docx', '.'),
        ('AWD亚马逊分销POD.pdf', '.'),
        ('FBA直发POD.pdf', '.'),
    ] + ddddocr_datas,
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
    a.binaries,
    a.datas,
    [],
    name='FBA货件差异自动索赔',
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
