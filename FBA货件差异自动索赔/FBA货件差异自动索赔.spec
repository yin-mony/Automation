# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all, collect_submodules

ddddocr_datas, ddddocr_binaries, ddddocr_hiddenimports = collect_all('ddddocr')
tkcalendar_datas, tkcalendar_binaries, tkcalendar_hiddenimports = collect_all('tkcalendar')
babel_datas, babel_binaries, babel_hiddenimports = collect_all('babel')

hiddenimports = [
    'tkinter',
    'tkinter.ttk',
    'tkinter.filedialog',
    'tkinter.messagebox',
    'tkinter.scrolledtext',
    '_tkinter',
    'tkcalendar',
    'main',
    'saihu',
    'test',
    'auto',
    'export',
    'email_util',
    'SaihuERPLogin',
    'YidekeLogin',
    'docx',
    'docx2pdf',
    'pdfplumber',
    'reportlab',
    'fitz',
    'DrissionPage',
    'psutil',
    'pywinauto',
]
hiddenimports += collect_submodules('DrissionPage')
hiddenimports += collect_submodules('docx')
hiddenimports += collect_submodules('pdfplumber')
hiddenimports += collect_submodules('reportlab')
hiddenimports += collect_submodules('fitz')
hiddenimports += ddddocr_hiddenimports
hiddenimports += tkcalendar_hiddenimports
hiddenimports += babel_hiddenimports

a = Analysis(
    ['run.py'],
    pathex=[],
    binaries=ddddocr_binaries + tkcalendar_binaries + babel_binaries,
    datas=[
        ('db53060fa183_发票模板.docx', '.'),
        ('服务商模板.docx', '.'),
        ('AWD_POD.pdf', '.'),
        ('FBA_POD.pdf', '.'),
    ] + ddddocr_datas + tkcalendar_datas + babel_datas,
    hiddenimports=hiddenimports,
    hookspath=['hooks'],
    hooksconfig={},
    runtime_hooks=['tk_runtime_hook.py'],
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
