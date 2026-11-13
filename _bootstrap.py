# -*- coding: utf-8 -*-
"""
运行时自举：把工具目录内自带的 lib 依赖提升到 sys.path 最前，
并让 pywin32 的 DLL 可被加载。使本工具完全不依赖：
  - WorkBuddy 私有 Python 及其 PYTHONPATH shim
  - 用户级 site-packages (AppData\\Roaming\\Python\\...)
  - 系统 Python 是否已安装第三方包
用法：任何脚本在 import 第三方库之前 `import _bootstrap`。
"""
import os
import sys

_TOOL_DIR = os.path.dirname(os.path.abspath(__file__))
_LIB = os.path.join(_TOOL_DIR, "lib")

if os.path.isdir(_LIB):
    # pywin32 的 pythoncom314.dll / pywintypes314.dll
    _DLL = os.path.join(_LIB, "pywin32_system32")
    if os.path.isdir(_DLL):
        os.environ["PATH"] = _DLL + os.pathsep + os.environ.get("PATH", "")
        if hasattr(os, "add_dll_directory"):
            try:
                os.add_dll_directory(_DLL)
            except Exception:
                pass

    # 等价于 pywin32.pth 的路径注入（win32 / win32\lib / pythonwin），
    # 外加 lib 自身与其它 pywin32 包目录
    _sub = [
        _LIB,
        os.path.join(_LIB, "win32"),
        os.path.join(_LIB, "win32", "lib"),
        os.path.join(_LIB, "win32com"),
        os.path.join(_LIB, "win32comext"),
        os.path.join(_LIB, "pythonwin"),
    ]
    for _p in reversed(_sub):
        if os.path.isdir(_p) and _p not in sys.path:
            sys.path.insert(0, _p)


def dependency_report():
    """返回依赖来源报告，用于验收取证。"""
    import PIL, fitz, win32print, win32ui
    return {
        "python": sys.executable,
        "PIL": PIL.__file__,
        "fitz": fitz.__file__,
        "win32print": win32print.__file__,
        "win32ui": win32ui.__file__,
        "uses_tool_lib": all(_LIB.lower() in p.lower() for p in
                             (PIL.__file__, fitz.__file__, win32print.__file__, win32ui.__file__)),
    }
