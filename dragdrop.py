# -*- coding: utf-8 -*-
"""
Windows 原生文件拖放（WM_DROPFILES），纯 ctypes + 标准库，无第三方依赖。

【为什么不再用 COM IDropTarget】
上一版用纯 ctypes 注册 IDropTarget。真机验证发现两个致命问题：
  1) 把文件拖到 A3 区时程序直接关闭 —— COM 回调可能在 RPC 工作线程上调用
     Tkinter（widget.after），Tk 不是线程安全的，会原生崩溃；
  2) 用真实 OLE DoDragDrop 复现时，源端永久阻塞在 IDropTarget::DragEnter，
     连 IDropSource::QueryContinueDrag 都没有被调用 —— Tk 的 Tcl 事件循环
     无法可靠地为跨进程 COM STA 调用泵消息。
因此改为 Win32 经典机制：DragAcceptFiles + 窗口过程子类化拦截 WM_DROPFILES。
这是资源管理器向普通窗口投递文件拖放时使用的真实 Windows 消息通道。

【线程安全约定】
  - 原生回调（窗口过程）里绝对不碰 Tkinter，只做 Win32 调用 + 写日志 + queue.put
  - Tk 主线程用 widget.after(80, ...) 轮询队列，在主线程里执行真正的加载回调
  - 所有原生回调都包 try/except，异常写日志，绝不抛出到 Win32 调用栈
"""
import os
import queue
import ctypes
import datetime
import threading
from ctypes import wintypes, WINFUNCTYPE, POINTER, byref, cast, c_void_p

WM_DROPFILES = 0x0233
CF_HDROP = 15
GWLP_WNDPROC = -4
GWL_EXSTYLE = -20
WS_EX_ACCEPTFILES = 0x00000010

_shell32 = ctypes.WinDLL("shell32", use_last_error=True)
_user32 = ctypes.WinDLL("user32", use_last_error=True)

_WPARAM = ctypes.c_size_t
_LPARAM = ctypes.c_ssize_t
_LRESULT = ctypes.c_ssize_t
_WNDPROC = WINFUNCTYPE(_LRESULT, wintypes.HWND, wintypes.UINT, _WPARAM, _LPARAM)
_ENUMPROC = WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, _LPARAM)

_shell32.DragQueryFileW.argtypes = [c_void_p, wintypes.UINT, wintypes.LPWSTR, wintypes.UINT]
_shell32.DragQueryFileW.restype = wintypes.UINT
_shell32.DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]
_shell32.DragAcceptFiles.restype = None
_shell32.DragFinish.argtypes = [c_void_p]
_shell32.DragFinish.restype = None

_user32.GetParent.argtypes = [wintypes.HWND]
_user32.GetParent.restype = wintypes.HWND
_user32.CallWindowProcW.argtypes = [c_void_p, wintypes.HWND, wintypes.UINT,
                                   _WPARAM, _LPARAM]
_user32.CallWindowProcW.restype = _LRESULT
_user32.EnumChildWindows.argtypes = [wintypes.HWND, _ENUMPROC, _LPARAM]
_user32.GetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int]
_user32.GetWindowLongPtrW.restype = c_void_p

try:
    _SetWindowLongPtrW = _user32.SetWindowLongPtrW          # 64 位
except AttributeError:
    _SetWindowLongPtrW = _user32.SetWindowLongW             # 32 位
_SetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int, c_void_p]
_SetWindowLongPtrW.restype = c_void_p

BASE = os.path.dirname(os.path.abspath(__file__))
NOTE_LOG = os.path.join(BASE, "A3SplitTool.log")
_LOCK = threading.Lock()

_STATE = {"callback": None, "widget": None, "hwnd": None, "registered": False,
          "mechanism": "WM_DROPFILES(DragAcceptFiles + WndProc)",
          "drops": 0, "last_paths": None, "hr": 0, "errors": []}
_QUEUE = queue.Queue()
_OLDPROC = {}                  # hwnd(int) -> 原窗口过程地址
_KEEP = []                     # 防止 WNDPROC / 结构被 GC


def _note(msg):
    """原生回调里唯一允许写的日志：纯文件 IO，不碰 Tk。"""
    try:
        with _LOCK:
            with open(NOTE_LOG, "a", encoding="utf-8") as f:
                f.write("[%s] DRAGDROP %s\n" %
                        (datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), msg))
    except Exception:
        pass


def extract_paths(hdrop):
    """从 HDROP 解出路径，并调用 DragFinish 释放（WM_DROPFILES 接收方义务）。"""
    h = c_void_p(hdrop)
    paths = []
    try:
        count = _shell32.DragQueryFileW(h, 0xFFFFFFFF, None, 0)
        for i in range(count):
            buf = ctypes.create_unicode_buffer(2048)
            _shell32.DragQueryFileW(h, i, buf, 2048)
            if buf.value:
                paths.append(buf.value)
    except Exception as e:
        _STATE["errors"].append(repr(e))
    finally:
        try:
            _shell32.DragFinish(h)
        except Exception:
            pass
    return paths


@_WNDPROC
def _wndproc(hwnd, msg, wparam, lparam):
    """真实 Windows 消息回调：只做 Win32 + 队列，绝不调用 Tkinter。"""
    try:
        if msg == WM_DROPFILES:
            paths = extract_paths(wparam)
            _STATE["drops"] += 1
            _STATE["last_paths"] = paths
            _note("WM_DROPFILES hwnd=%s paths=%s" % (int(hwnd or 0), paths))
            if paths:
                _QUEUE.put(paths)
            return 0
        old = _OLDPROC.get(int(hwnd or 0))
        if old:
            return _user32.CallWindowProcW(c_void_p(old), hwnd, msg, wparam, lparam)
    except Exception:
        _note("WndProc 异常: " + _exc())
    return 0


def _exc():
    import traceback
    return traceback.format_exc().replace("\n", " | ")


def _poll():
    """Tk 主线程轮询：把原生回调收到的路径交给 GUI 加载（线程安全）。"""
    try:
        while True:
            paths = _QUEUE.get_nowait()
            cb = _STATE["callback"]
            if cb:
                try:
                    cb(paths)
                except Exception:
                    _note("加载回调异常: " + _exc())
    except queue.Empty:
        pass
    except Exception:
        _note("轮询异常: " + _exc())
    if _STATE["registered"]:
        try:
            _STATE["widget"].after(80, _poll)
        except Exception:
            _STATE["registered"] = False


def _collect_hwnds(widget):
    widget.update_idletasks()
    own = widget.winfo_id()
    try:
        top = _user32.GetParent(wintypes.HWND(own)) or own
    except Exception:
        top = own
    hwnds = [int(top)]
    if int(own) != int(top):
        hwnds.append(int(own))

    found = []

    @_ENUMPROC
    def _each(h, _lparam):
        found.append(int(h or 0))
        return 1

    try:
        _user32.EnumChildWindows(wintypes.HWND(int(top)), _each, 0)
    except Exception:
        pass
    _KEEP.append(_each)
    for h in found:
        if h not in hwnds:
            hwnds.append(h)
    return hwnds, int(top)


def hook_dropfiles(widget, on_files):
    """注册文件拖放目标。返回是否成功；失败时 GUI 自动降级为「打开文件」。"""
    try:
        hwnds, top = _collect_hwnds(widget)
        cb = cast(_wndproc, c_void_p)
        _KEEP.append(_wndproc)
        for h in hwnds:
            try:
                H = wintypes.HWND(h)
                _shell32.DragAcceptFiles(H, True)
                old = _SetWindowLongPtrW(H, GWLP_WNDPROC, cb)
                if old:
                    _OLDPROC[h] = int(old)
            except Exception as e:
                _STATE["errors"].append("hwnd=%s %r" % (h, e))
        _STATE.update({"callback": on_files, "widget": widget, "hwnd": top,
                       "registered": bool(_OLDPROC), "hr": 0})
        widget.after(80, _poll)
        _note("注册=%s 顶层hwnd=%s 子类化窗口数=%d 机制=%s" %
              (_STATE["registered"], top, len(_OLDPROC), _STATE["mechanism"]))
        return _STATE["registered"]
    except Exception:
        _note("注册异常: " + _exc())
        return False


def unhook_dropfiles():
    for h, old in list(_OLDPROC.items()):
        try:
            H = wintypes.HWND(h)
            _SetWindowLongPtrW(H, GWLP_WNDPROC, c_void_p(old))
            _shell32.DragAcceptFiles(H, False)
        except Exception:
            pass
    _OLDPROC.clear()
    _STATE["registered"] = False


def status():
    """返回注册状态，供独立验证脚本读取（不含回调对象）。"""
    return {k: v for k, v in _STATE.items() if k not in ("callback", "widget")}


def pop_paths(timeout=0.2):
    """测试辅助：从队列取一批路径（不触发 Tk）。"""
    try:
        return _QUEUE.get(timeout=timeout)
    except queue.Empty:
        return None


def debug_inject_paths(paths):
    """测试辅助：模拟一次已收到的 WM_DROPFILES 负载（仅用于静态回归，非真实验证）。"""
    _STATE["drops"] += 1
    _STATE["last_paths"] = list(paths)
    _QUEUE.put(list(paths))
    return True


def debug_drop_callback():
    """兼容旧静态验证脚本的 5 参数回调签名：把负载投递到队列。"""
    def _fake(self, pdo, key_state, pt, effect):
        if effect:
            effect[0] = 1
        return 0
    return _fake
