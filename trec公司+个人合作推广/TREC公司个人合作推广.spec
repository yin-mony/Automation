# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules


# hiddenimports：补充 PyInstaller 静态分析不容易自动识别的动态依赖。
hiddenimports = [
    "main",
    "email_util",
    "openpyxl",
    "openpyxl.cell._writer",
    "DrissionPage",
]

# DrissionPage：二级页面浏览器抓取依赖，打包时完整收集子模块。
hiddenimports += collect_submodules("DrissionPage")

# openpyxl：读取和导出 xlsx 依赖，打包时完整收集子模块。
hiddenimports += collect_submodules("openpyxl")


a = Analysis(
    ["run.py"],
    pathex=[],
    binaries=[],
    # datas：把当前子项目 file 目录作为内置数据资源打进 exe。
    datas=[("file", "file")],
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
    name="TREC公司个人合作推广",
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
