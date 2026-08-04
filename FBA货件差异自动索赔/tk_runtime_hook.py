"""设置打包后 Tcl/Tk 资源路径。"""

import os
import sys

baseDir = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
tclDir = os.path.join(baseDir, "_tcl_data")
tkDir = os.path.join(baseDir, "_tk_data")

if os.path.isdir(tclDir):
    os.environ["TCL_LIBRARY"] = tclDir

if os.path.isdir(tkDir):
    os.environ["TK_LIBRARY"] = tkDir
