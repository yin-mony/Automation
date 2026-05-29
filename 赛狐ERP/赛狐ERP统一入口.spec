# -*- mode: python ; coding: utf-8 -*-

import os
from PyInstaller.utils.hooks import collect_all, collect_submodules

ddddocr_datas, ddddocr_binaries, ddddocr_hiddenimports = collect_all('ddddocr')
onnx_datas, onnx_binaries, onnx_hiddenimports = collect_all('onnxruntime')

conda_bin_dir = r"C:\Users\admin\miniconda3\envs\saihu312\Library\bin"
required_conda_dlls = [
    "liblzma.dll",
    "libbz2.dll",
    "ffi.dll",
    "libexpat.dll",
    "sqlite3.dll",
]
extra_binaries = []
for dll_name in required_conda_dlls:
    dll_path = os.path.join(conda_bin_dir, dll_name)
    if os.path.exists(dll_path):
        extra_binaries.append((dll_path, "."))

extra_hiddenimports = []
extra_hiddenimports.extend(collect_submodules('ddddocr'))
extra_hiddenimports.extend(collect_submodules('onnxruntime'))

all_hiddenimports = [
    'OnlyMain',
    'LowMain',
    'DewMain',
    'SaihuERPLogin',
]
all_hiddenimports.extend(ddddocr_hiddenimports)
all_hiddenimports.extend(onnx_hiddenimports)
all_hiddenimports.extend(extra_hiddenimports)
all_hiddenimports = list(dict.fromkeys(all_hiddenimports))

a = Analysis(
    ['OnlyRun.py'],
    pathex=['c:\\Users\\admin\\Desktop\\Automation\\赛狐ERP'],
    binaries=ddddocr_binaries + onnx_binaries + extra_binaries,
    datas=ddddocr_datas + onnx_datas,
    hiddenimports=all_hiddenimports,
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
