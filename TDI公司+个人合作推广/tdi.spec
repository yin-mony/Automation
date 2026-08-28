# -*- mode: python ; coding: utf-8 -*-

import sys
from pathlib import Path

from PyInstaller.utils.hooks import collect_submodules


# 生成目录版 EXE；GUI、直连抓取和大表文件保持外置目录更稳定。
hiddenimports = [
    "main",
    "data",
    "serp",
    "browser",
    "mail",
    "requests",
    "certifi",
    "ssl",
    "_ssl",
    "openpyxl",
    "lxml",
    "PySide6",
]
hiddenimports += collect_submodules("openpyxl")

python_dlls = Path(sys.base_prefix) / "DLLs"
ssl_binaries = [
    (str(python_dlls / "libcrypto-3-x64.dll"), "."),
    (str(python_dlls / "libssl-3-x64.dll"), "."),
]

a = Analysis(
    ["run.py"],
    pathex=[],
    binaries=ssl_binaries,
    datas=[
        ("file/time2renew-logo.png", "file"),
        ("使用说明.md", "."),
    ],
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
    name="TDI推广工具",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
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
    upx=False,
    upx_exclude=[],
    name="TDI推广工具",
)
