"""手动收集 Tcl/Tk 运行资源。

PyInstaller 在当前机器上无法自动判断 Tcl/Tk 可用性，所以这里显式收集：
1. _tkinter.pyd
2. tcl86t.dll / tk86t.dll
3. tcl8.6 / tk8.6 脚本目录
"""

import sys
from pathlib import Path

pythonRoot = Path(sys.base_prefix)
dllDir = pythonRoot / "DLLs"
tclRoot = pythonRoot / "tcl"

binaries = []
datas = []

for dllName in ("_tkinter.pyd", "tcl86t.dll", "tk86t.dll"):
    dllPath = dllDir / dllName
    if dllPath.is_file():
        binaries.append((str(dllPath), "."))


def collectTclDir(sourceDir, targetDir):
    """按 PyInstaller 运行时钩子要求收集 Tcl/Tk 脚本目录。"""
    if not sourceDir.is_dir():
        return
    for path in sourceDir.rglob("*"):
        if not path.is_file():
            continue
        if path.suffix.lower() == ".lib":
            continue
        target = Path(targetDir) / path.relative_to(sourceDir).parent
        datas.append((str(path), str(target)))


collectTclDir(tclRoot / "tcl8.6", "_tcl_data")
collectTclDir(tclRoot / "tk8.6", "_tk_data")
collectTclDir(tclRoot / "tcl8", "tcl8")
collectTclDir(tclRoot / "dde1.4", "dde1.4")
collectTclDir(tclRoot / "reg1.3", "reg1.3")
