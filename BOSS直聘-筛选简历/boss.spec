# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all, collect_submodules

block_cipher = None

dpDatas, dpBinaries, dpHidden = collect_all('DrissionPage')

hiddenimports = [
    'run',
    'boss_web',
    'boss_web.auto',
    'boss_web.db',
    'boss_web.login',
    'boss_web.job',
    'boss_web.template',
    'boss_web.report',
    'boss_web.reply',
    'boss_web.parse',
    'boss_web.match',
    'DrissionPage',
] + dpHidden + collect_submodules('boss_web') + collect_submodules('DrissionPage')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=dpBinaries,
    datas=dpDatas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='BOSS直聘筛选简历',
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
