"""覆盖 PyInstaller 默认 tkinter 预查找钩子。

当前打包机的 Tcl/Tk 自动探测会失败，但 tkinter 模块文件本身可用。
默认钩子会在探测失败时清空搜索目录，导致 exe 缺失 tkinter。
这里保留正常搜索目录，Tcl/Tk 资源由 hook-_tkinter.py 手动收集。
"""


def pre_find_module_path(hook_api):
    """保留 tkinter 模块搜索路径。"""
    return
