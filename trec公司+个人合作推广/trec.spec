# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules


# 生成目录版 EXE；大型 GUI 和浏览器依赖无需每次启动时临时解压。
hiddenimports = [
    "main",
    "data",
    "serp",
    "proxy",
    "browser",
    "mail",
    "socks",
    "pproxy",
    "openpyxl",
    "lxml",
    "DrissionPage",
    "PySide6",
]
hiddenimports += collect_submodules("DrissionPage")
hiddenimports += collect_submodules("openpyxl")
hiddenimports += collect_submodules("pproxy")

a = Analysis(
    ["run.py"],
    pathex=[],
    binaries=[],
    datas=[("file/time2renew-logo.png", "file")],
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
    name="TREC推广工具",
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

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="TREC推广工具",
)
