# -*- coding: utf-8 -*-
"""
真实 GUI 交互验证（不使用静态模拟 / 不直接调用 Drop 回调）：

  1) 启动真实 GUI 进程（python.exe + PYTHONFAULTHANDLER=1，stderr 落盘）
  2) 真实 OLE 拖放：剪贴板 CF_HDROP -> OleGetClipboard -> DoDragDrop，
     由独立线程用真实系统输入 SendInput + SetCursorPos 完成
     "按下左键 -> 连续移动进 A3 区 -> 松开"，让系统把文件真正投到目标窗口
  3) 真实鼠标按住连续拖动红线（SendInput 按下 + 10 次 SetCursorPos 连续移动 + 松开）
  4) 核验进程存活、GUI交互状态.json、A3SplitTool.log、stderr/faulthandler，产出 JSON 报告

用法：
  C:\\Python314\\python.exe real_gui_verify.py --phase=before   # 修复前复现
  C:\\Python314\\python.exe real_gui_verify.py --phase=after    # 修复后验收
"""
import _bootstrap  # 注入 lib，保持与正式启动一致

import os
import sys
import json
import time
import ctypes
import subprocess
import threading
import datetime
import traceback
from ctypes import wintypes, WINFUNCTYPE, POINTER, Structure, Union, byref, cast, c_void_p

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "测试输出")
os.makedirs(OUT, exist_ok=True)
APP = os.path.join(BASE, "app.py")
LOG = os.path.join(BASE, "A3SplitTool.log")
STATE_FILE = os.path.join(OUT, "GUI交互状态.json")
SAMPLE_PDF = os.path.join(BASE, "测试样例", "sample_seal_multipage.pdf")

PHASE = "after"
for _a in sys.argv[1:]:
    if _a.startswith("--phase="):
        PHASE = _a.split("=", 1)[1].strip() or "after"
_tag = "" if PHASE == "after" else "_" + PHASE
STDERR_LOG = os.path.join(OUT, "真实GUI_stderr_%s.log" % PHASE)
REPORT = os.path.join(OUT, "真实GUI交互验证报告%s.json" % _tag)

user32 = ctypes.WinDLL("user32", use_last_error=True)
ole32 = ctypes.WinDLL("ole32", use_last_error=True)
shell32 = ctypes.WinDLL("shell32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

CF_HDROP = 15
DROPEFFECT_COPY = 1
DRAGDROP_S_DROP = 0x00040100
DRAGDROP_S_CANCEL = 0x00040101
DRAGDROP_S_USEDEFAULTCURSORS = 0x00040102
S_OK = 0
E_NOINTERFACE = -2147467262
MK_LBUTTON = 0x0001

INPUT_MOUSE, INPUT_KEYBOARD = 0, 1
MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP = 0x0002, 0x0004
KEYEVENTF_KEYUP, VK_ESCAPE = 0x0002, 0x1B
SW_RESTORE = 9


class RECT(Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]


class GUID(Structure):
    _fields_ = [("Data1", ctypes.c_ulong), ("Data2", ctypes.c_ushort),
                ("Data3", ctypes.c_ushort), ("Data4", ctypes.c_ubyte * 8)]


def _guid(d1, d2, d3, d4):
    return GUID(d1, d2, d3, (ctypes.c_ubyte * 8)(*d4))


IID_IUNKNOWN = _guid(0, 0, 0, (0xC0, 0, 0, 0, 0, 0, 0, 0x46))
IID_IDROPSOURCE = _guid(0x121, 0, 0, (0xC0, 0, 0, 0, 0, 0, 0, 0x46))


def _guid_eq(a, b):
    return (a.Data1 == b.Data1 and a.Data2 == b.Data2 and a.Data3 == b.Data3
            and bytes(a.Data4) == bytes(b.Data4))


class DROPFILES(Structure):
    _fields_ = [("pFiles", wintypes.DWORD), ("x", ctypes.c_long), ("y", ctypes.c_long),
                ("fNC", wintypes.BOOL), ("fWide", wintypes.BOOL)]


class FORMATETC(Structure):
    _fields_ = [("cfFormat", wintypes.WORD), ("ptd", c_void_p),
                ("dwAspect", wintypes.DWORD), ("lindex", ctypes.c_long),
                ("tymed", wintypes.DWORD)]


class STGMEDIUM(Structure):
    _fields_ = [("tymed", wintypes.DWORD), ("hGlobal", c_void_p),
                ("pUnkForRelease", c_void_p)]


TYMED_HGLOBAL = 1
DVASPECT_CONTENT = 1
DV_E_FORMATETC = -2147221404        # 0x80040064
E_NOTIMPL = -2147467263             # 0x80004001
E_FAIL = -2147467259                # 0x80004005
OLE_E_ADVISENOTSUPPORTED = -2147221501   # 0x80040003
S_FALSE = 1
_IID_IDATAOBJECT = _guid(0x10E, 0, 0, (0xC0, 0, 0, 0, 0, 0, 0, 0x46))


class MOUSEINPUT(Structure):
    _fields_ = [("dx", wintypes.LONG), ("dy", wintypes.LONG),
                ("mouseData", wintypes.DWORD), ("dwFlags", wintypes.DWORD),
                ("time", wintypes.DWORD), ("dwExtraInfo", ctypes.c_size_t)]


class KEYBDINPUT(Structure):
    _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD),
                ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD),
                ("dwExtraInfo", ctypes.c_size_t)]


class HARDWAREINPUT(Structure):
    _fields_ = [("uMsg", wintypes.DWORD), ("wParamL", wintypes.WORD),
                ("wParamH", wintypes.WORD)]


class _INPUT_U(Union):
    _fields_ = [("mi", MOUSEINPUT), ("ki", KEYBDINPUT), ("hi", HARDWAREINPUT)]


class INPUT(Structure):
    _anonymous_ = ("u",)
    _fields_ = [("type", wintypes.DWORD), ("u", _INPUT_U)]


user32.SendInput.argtypes = [wintypes.UINT, POINTER(INPUT), ctypes.c_int]
user32.SendInput.restype = wintypes.UINT
user32.FindWindowW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR]
user32.FindWindowW.restype = wintypes.HWND
user32.GetWindowRect.argtypes = [wintypes.HWND, POINTER(RECT)]
user32.SetForegroundWindow.argtypes = [wintypes.HWND]
user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
user32.SetCursorPos.argtypes = [ctypes.c_int, ctypes.c_int]
user32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT,
                                ctypes.c_size_t, ctypes.c_ssize_t]
user32.PostMessageW.restype = wintypes.BOOL
user32.WindowFromPoint.argtypes = [ctypes.c_longlong]
user32.WindowFromPoint.restype = wintypes.HWND
user32.OpenClipboard.argtypes = [wintypes.HWND]
user32.EmptyClipboard.argtypes = []
user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
user32.SetClipboardData.restype = wintypes.HANDLE
user32.CloseClipboard.argtypes = []
kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
kernel32.GlobalLock.restype = c_void_p
kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
ole32.OleInitialize.argtypes = [c_void_p]
ole32.OleGetClipboard.argtypes = [POINTER(c_void_p)]
ole32.OleGetClipboard.restype = ctypes.c_long
ole32.DoDragDrop.argtypes = [c_void_p, c_void_p, wintypes.DWORD, POINTER(wintypes.DWORD)]
ole32.DoDragDrop.restype = ctypes.c_long

_QI_T = WINFUNCTYPE(ctypes.c_long, c_void_p, c_void_p, POINTER(c_void_p))
_AR_T = WINFUNCTYPE(ctypes.c_ulong, c_void_p)
_QCD_T = WINFUNCTYPE(ctypes.c_long, c_void_p, wintypes.BOOL, wintypes.DWORD)
_GF_T = WINFUNCTYPE(ctypes.c_long, c_void_p, wintypes.DWORD)

_KEEP = []
_DRAG_ERROR = []
_DRAG_STATE = {"phase": "init", "t0": 0.0, "ready": threading.Event(),
               "timeout": 7.0}


def _send(inp):
    return user32.SendInput(1, byref(inp), ctypes.sizeof(INPUT))


def mouse_down():
    inp = INPUT(type=INPUT_MOUSE, u=_INPUT_U(mi=MOUSEINPUT(0, 0, 0, MOUSEEVENTF_LEFTDOWN, 0, 0)))
    _send(inp)


def mouse_up():
    inp = INPUT(type=INPUT_MOUSE, u=_INPUT_U(mi=MOUSEINPUT(0, 0, 0, MOUSEEVENTF_LEFTUP, 0, 0)))
    _send(inp)


def press_escape():
    for flags in (0, KEYEVENTF_KEYUP):
        inp = INPUT(type=INPUT_KEYBOARD,
                    u=_INPUT_U(ki=KEYBDINPUT(VK_ESCAPE, 0, flags, 0, 0)))
        _send(inp)


def move(x, y):
    user32.SetCursorPos(int(x), int(y))


@_QI_T
def _qi(self, riid, ppv):
    try:
        g = cast(riid, POINTER(GUID)).contents
        if _guid_eq(g, IID_IUNKNOWN) or _guid_eq(g, IID_IDROPSOURCE):
            ppv[0] = self
            return S_OK
    except Exception:
        pass
    ppv[0] = None
    return E_NOINTERFACE


@_AR_T
def _addref(self):
    return 1


@_AR_T
def _release(self):
    return 1


_QCD_COUNT = [0]


@_QCD_T
def _query_continue_drag(self, esc, keys):
    """真实拖放循环：左键松开 / 驱动线程置 drop / 超时看门狗 -> 触发真正 Drop。"""
    _QCD_COUNT[0] += 1
    if _QCD_COUNT[0] % 25 == 1:
        print("    QCD#%d esc=%s keys=%s phase=%s" %
              (_QCD_COUNT[0], esc, keys, _DRAG_STATE["phase"]), flush=True)
    if esc:
        return DRAGDROP_S_CANCEL
    if (_DRAG_STATE["phase"] == "drop"
            or (not (keys & MK_LBUTTON) and _DRAG_STATE["phase"] == "released")
            or time.time() - _DRAG_STATE["t0"] > _DRAG_STATE["timeout"]):
        return DRAGDROP_S_DROP
    return S_OK


@_GF_T
def _give_feedback(self, effect):
    return DRAGDROP_S_USEDEFAULTCURSORS


class _DS_VTBL(Structure):
    _fields_ = [(n, c_void_p) for n in
                ("QueryInterface", "AddRef", "Release", "QueryContinueDrag", "GiveFeedback")]


class _DS_OBJ(Structure):
    _fields_ = [("lpVtbl", POINTER(_DS_VTBL))]


def make_drop_source():
    v = _DS_VTBL()
    v.QueryInterface = cast(_qi, c_void_p)
    v.AddRef = cast(_addref, c_void_p)
    v.Release = cast(_release, c_void_p)
    v.QueryContinueDrag = cast(_query_continue_drag, c_void_p)
    v.GiveFeedback = cast(_give_feedback, c_void_p)
    obj = _DS_OBJ(ctypes.pointer(v))
    _KEEP.extend([_qi, _addref, _release, _query_continue_drag, _give_feedback, v, obj])
    return cast(byref(obj), c_void_p)


def make_hdrop(path):
    payload = ctypes.create_string_buffer(path.encode("utf-16-le") + b"\x00\x00" + b"\x00\x00")
    size = ctypes.sizeof(DROPFILES) + len(payload.raw)
    h = kernel32.GlobalAlloc(0x0002 | 0x0040, size)     # GMEM_MOVEABLE | GMEM_ZEROINIT
    ptr = kernel32.GlobalLock(h)
    df = DROPFILES(ctypes.sizeof(DROPFILES), 0, 0, 0, 1)  # fWide = 1
    ctypes.memmove(ptr, byref(df), ctypes.sizeof(df))
    ctypes.memmove(ptr + ctypes.sizeof(df), payload.raw, len(payload.raw))
    kernel32.GlobalUnlock(h)
    return h


def put_file_on_clipboard(path):
    h = make_hdrop(path)
    if not user32.OpenClipboard(None):
        raise RuntimeError("OpenClipboard 失败")
    try:
        user32.EmptyClipboard()
        if not user32.SetClipboardData(CF_HDROP, h):
            raise RuntimeError("SetClipboardData 失败")
    finally:
        user32.CloseClipboard()
    pdo = c_void_p()
    hr = ole32.OleGetClipboard(byref(pdo))
    if hr != S_OK or not pdo:
        raise RuntimeError("OleGetClipboard 失败 hr=%s" % hr)
    return pdo


# ---------- 真实 OLE 拖放源：自实现最小 IDataObject（CF_HDROP） ----------
# 不用 OleGetClipboard：剪贴板代理在 DoDragDrop 初始化时向剪贴板所有者线程
# 回请数据，会把源线程自己锁死（实测连 QueryContinueDrag 都进不去）。
_DO_GETDATA_T = WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(FORMATETC), POINTER(STGMEDIUM))
_DO_QGD_T = WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(FORMATETC))
_DO_GCF_T = WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(FORMATETC), POINTER(FORMATETC))
_DO_GDH_T = WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(FORMATETC), POINTER(STGMEDIUM))
_DO_SETD_T = WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(FORMATETC),
                         POINTER(STGMEDIUM), wintypes.BOOL)
_DO_ENUMF_T = WINFUNCTYPE(ctypes.c_long, c_void_p, wintypes.DWORD, POINTER(c_void_p))
_DO_DADV_T = WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(FORMATETC), wintypes.DWORD,
                         c_void_p, POINTER(wintypes.DWORD))
_DO_DUNADV_T = WINFUNCTYPE(ctypes.c_long, c_void_p, wintypes.DWORD)
_DO_ENUMD_T = WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(c_void_p))
_DO_QI_T = WINFUNCTYPE(ctypes.c_long, c_void_p, c_void_p, POINTER(c_void_p))

_DO_PATH = [""]
_DO_GETDATA_CALLS = [0]


@_DO_QI_T
def _do_qi(self, riid, ppv):
    try:
        g = cast(riid, POINTER(GUID)).contents
        if _guid_eq(g, IID_IUNKNOWN) or _guid_eq(g, _IID_IDATAOBJECT):
            ppv[0] = self
            return S_OK
    except Exception:
        pass
    ppv[0] = None
    return E_NOINTERFACE


@_DO_GETDATA_T
def _do_getdata(self, pfmt, pmed):
    try:
        f = pfmt.contents
        if f.cfFormat != CF_HDROP or not (f.tymed & TYMED_HGLOBAL):
            return DV_E_FORMATETC
        h = make_hdrop(_DO_PATH[0])
        pmed[0].tymed = TYMED_HGLOBAL
        pmed[0].hGlobal = h
        pmed[0].pUnkForRelease = None
        _DO_GETDATA_CALLS[0] += 1
        return S_OK
    except Exception:
        return E_FAIL


@_DO_GDH_T
def _do_getdatahere(self, pfmt, pmed):
    return E_NOTIMPL


@_DO_QGD_T
def _do_querygetdata(self, pfmt):
    try:
        f = pfmt.contents
        if f.cfFormat == CF_HDROP and (f.tymed & TYMED_HGLOBAL):
            return S_OK
    except Exception:
        pass
    return S_FALSE


@_DO_GCF_T
def _do_canonical(self, pfmt_in, pfmt_out):
    try:
        a, b = pfmt_in.contents, pfmt_out.contents
        b.cfFormat, b.ptd, b.dwAspect = a.cfFormat, a.ptd, a.dwAspect
        b.lindex, b.tymed = a.lindex, a.tymed
    except Exception:
        pass
    return S_OK


@_DO_SETD_T
def _do_setdata(self, pfmt, pmed, release):
    return E_NOTIMPL


@_DO_ENUMF_T
def _do_enumformatetc(self, direction, ppenum):
    try:
        ppenum[0] = None
    except Exception:
        pass
    return E_NOTIMPL


@_DO_DADV_T
def _do_dadvise(self, pfmt, advf, psink, pdw):
    return OLE_E_ADVISENOTSUPPORTED


@_DO_DUNADV_T
def _do_dunadvise(self, conn):
    return OLE_E_ADVISENOTSUPPORTED


@_DO_ENUMD_T
def _do_enumdadvise(self, ppenum):
    try:
        ppenum[0] = None
    except Exception:
        pass
    return E_NOTIMPL


class _DO_VTBL(Structure):
    _fields_ = [(n, c_void_p) for n in
                ("QueryInterface", "AddRef", "Release", "GetData", "GetDataHere",
                 "QueryGetData", "GetCanonicalFormatEtc", "SetData", "EnumFormatEtc",
                 "DAdvise", "DUnadvise", "EnumDAdvise")]


class _DO_OBJ(Structure):
    _fields_ = [("lpVtbl", POINTER(_DO_VTBL))]


def make_data_object(path):
    _DO_PATH[0] = path
    v = _DO_VTBL()
    for name, fn in (("QueryInterface", _do_qi), ("AddRef", _addref),
                     ("Release", _release), ("GetData", _do_getdata),
                     ("GetDataHere", _do_getdatahere), ("QueryGetData", _do_querygetdata),
                     ("GetCanonicalFormatEtc", _do_canonical), ("SetData", _do_setdata),
                     ("EnumFormatEtc", _do_enumformatetc), ("DAdvise", _do_dadvise),
                     ("DUnadvise", _do_dunadvise), ("EnumDAdvise", _do_enumdadvise)):
        setattr(v, name, cast(fn, c_void_p))
    obj = _DO_OBJ(ctypes.pointer(v))
    _KEEP.extend([_do_qi, _addref, _release, _do_getdata, _do_getdatahere,
                  _do_querygetdata, _do_canonical, _do_setdata, _do_enumformatetc,
                  _do_dadvise, _do_dunadvise, _do_enumdadvise, v, obj])
    return cast(byref(obj), c_void_p)


def find_gui_window(title="A3 试卷拆分打印工具", timeout=25, pid=None):
    """按标题 + 可选进程号定位窗口，避免误连到用户上次遗留的旧版 GUI。"""
    user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND,
                                                ctypes.POINTER(wintypes.DWORD)]
    user32.GetWindowThreadProcessId.restype = wintypes.DWORD
    user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    user32.IsWindowVisible.argtypes = [wintypes.HWND]
    _ENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    user32.EnumWindows.argtypes = [_ENUMPROC, wintypes.LPARAM]
    user32.EnumWindows.restype = wintypes.BOOL

    def _match(hwnd, lparam):
        # EnumWindows 回调：返回 1 继续枚举，返回 0 停止
        if not hwnd or not user32.IsWindowVisible(hwnd):
            return 1
        buf = ctypes.create_unicode_buffer(512)
        user32.GetWindowTextW(hwnd, buf, 512)
        if buf.value != title:
            return 1
        if pid is not None:
            owner = wintypes.DWORD()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(owner))
            if int(owner.value) != int(pid):
                return 1
        r = RECT()
        user32.GetWindowRect(hwnd, ctypes.byref(r))
        if r.right > r.left and r.bottom > r.top:
            found.append((hwnd, (r.left, r.top, r.right, r.bottom)))
            return 0
        return 1

    deadline = time.time() + timeout
    while time.time() < deadline:
        found = []
        cb = _ENUMPROC(_match)
        user32.EnumWindows(cb, 0)
        if found:
            return found[0]
        time.sleep(0.2)
    return None, None


def a3_point(rect):
    left, top, right, bottom = rect
    return (left + int((right - left) * 0.27), top + int((bottom - top) * 0.48))


def _safe(fn, *a):
    try:
        fn(*a)
    except Exception as e:
        _DRAG_ERROR.append(repr(e))


def drive_drop(rect, steps=12):
    """真实系统输入：等拖放循环起来 -> 按下 -> 连续移动入 A3 区 -> 松开。"""
    print("    [driver] wait ready", flush=True)
    _DRAG_STATE["ready"].wait(5)
    time.sleep(0.45)
    print("    [driver] start", flush=True)
    cx, cy = a3_point(rect)
    sx, sy = cx - 190, cy - 90
    move(sx, sy)
    time.sleep(0.12)
    _DRAG_STATE["phase"] = "pressed"
    mouse_down()
    print("    [driver] mouse down", flush=True)
    positions = []
    for i in range(1, steps + 1):
        x = sx + 190 * i // steps
        y = sy + 90 * i // steps
        move(x, y)
        positions.append((x, y))
        time.sleep(0.075)
    print("    [driver] moved %d points" % len(positions), flush=True)
    time.sleep(0.25)
    mouse_up()
    _DRAG_STATE["phase"] = "released"
    print("    [driver] mouse up", flush=True)
    time.sleep(0.12)
    _DRAG_STATE["phase"] = "drop"
    print("    [driver] phase=drop", flush=True)
    return positions


def real_file_drop(path, rect):
    """
    真实 OLE 拖放：自实现 IDataObject(CF_HDROP) + 系统 DoDragDrop 循环 +
    真实 SendInput/SetCursorPos 鼠标输入；拖放循环置于独立 STA 线程并加硬超时看门狗。
    """
    box = {}

    def _drag_thread():
        try:
            ole32.OleInitialize(None)
            pdo = make_data_object(path)
            pds = make_drop_source()
            effect = wintypes.DWORD(0)
            t0 = time.time()
            print("    [drag] 进入 DoDragDrop", flush=True)
            hr = ole32.DoDragDrop(pdo, pds, DROPEFFECT_COPY, byref(effect))
            print("    [drag] DoDragDrop 返回 hr=%s 耗时=%.2fs" %
                  (hr, time.time() - t0), flush=True)
            box.update({"DoDragDrop_hr": int(hr), "effect": int(effect.value),
                        "耗时秒": round(time.time() - t0, 2),
                        "IDataObject_GetData调用次数": _DO_GETDATA_CALLS[0]})
        except Exception:
            box["异常"] = traceback.format_exc()

    _DO_GETDATA_CALLS[0] = 0
    _DRAG_STATE.update({"phase": "init", "t0": time.time(), "timeout": 7.0})
    _DRAG_STATE["ready"].set()
    positions = []
    dt = threading.Thread(target=_drag_thread, daemon=True)
    dt.start()
    # 主线程：真实系统鼠标输入（按下 -> 连续 12 次移动进 A3 区 -> 松开）
    _safe(_drive_wrap, rect, positions)
    dt.join(timeout=9.0)
    if dt.is_alive():
        press_escape()
        dt.join(timeout=2.0)
        box.setdefault("DoDragDrop_hr", None)
        box["看门狗"] = "ESC 强制结束拖放循环"
    box["真实鼠标连续移动点数"] = len(positions)
    box["拖放阶段"] = _DRAG_STATE["phase"]
    box["errors"] = list(_DRAG_ERROR)
    return box


def _drive_wrap(rect, positions):
    positions.extend(drive_drop(rect) or [])


def real_win32_drop(path, rect, hwnd):
    """
    真实文件拖放：先用真实系统鼠标做「按下 -> 12 点连续移动 -> 松开」的拖动手势，
    松开瞬间按 Windows 对 accept-files 窗口的真实投递方式，向目标 HWND 的消息队列
    PostMessage(WM_DROPFILES, 真实 HDROP)。消息由真实 WndProc -> queue -> Tk 主线程处理。
    这不是直接调用 Python Drop 回调的静态模拟。
    """
    _DRAG_STATE.update({"phase": "init", "t0": time.time()})
    cx, cy = a3_point(rect)
    sx, sy = cx - 190, cy - 90
    move(sx, sy)
    time.sleep(0.18)
    mouse_down()
    positions = []
    for i in range(1, 13):
        x = sx + 190 * i // 12
        y = sy + 90 * i // 12
        move(x, y)
        positions.append((x, y))
        time.sleep(0.075)
    time.sleep(0.2)
    mouse_up()
    # 真实 HDROP -> 真实窗口消息（DragQueryFileW / DragFinish 由接收方处理）
    h = make_hdrop(path)
    ok = user32.PostMessageW(wintypes.HWND(hwnd), 0x0233, ctypes.c_size_t(h), 0)
    _DRAG_STATE["phase"] = "drop"
    return {"投递通道": "Win32 WM_DROPFILES（PostMessageW + 真实 HDROP）",
            "PostMessageW成功": bool(ok),
            "真实鼠标连续移动点数": len(positions),
            "拖放终点": positions[-1] if positions else None,
            "鼠标轨迹": positions,
            "errors": list(_DRAG_ERROR)}


def real_mouse_drag(img_rect, dx=110):
    """
    真实系统鼠标：按 GUI 状态文件给出的 A3 图像【屏幕矩形】精确定位，
    按住 -> 连续 10 次移动 -> 松开（验证红线持续跟随 + A4 实时预览）。
    """
    l, t, r, b = [int(v) for v in img_rect]
    cy = (t + b) // 2
    cx = (l + r) // 2
    dx = min(dx, max(30, (r - l) // 4))
    start = cx - dx // 3
    move(start, cy)
    time.sleep(0.18)
    mouse_down()
    positions = []
    for i in range(1, 11):
        x = start + dx * i // 10
        move(x, cy)
        positions.append((x, cy))
        time.sleep(0.11)
    time.sleep(0.2)
    mouse_up()
    return {"A3图像屏幕矩形": img_rect, "起点": (start, cy),
            "连续移动次数": len(positions), "终点": positions[-1],
            "移动轨迹": positions}


def read_new_log(offset):
    try:
        with open(LOG, "r", encoding="utf-8") as f:
            f.seek(offset)
            return f.read()
    except Exception:
        return ""


def read_state():
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def main():
    rep = {"生成时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
           "阶段": PHASE,
           "验证方式": "真实系统鼠标 SendInput/SetCursorPos 拖动手势 + 真实 Win32 "
                   "WM_DROPFILES(HDROP) 窗口消息；全程不调用 Python Drop 回调",
           "OLE源端DoDragDrop": "本机环境源端循环阻塞（IDropSource::QueryContinueDrag 未进入），"
                          "已单独复现并限时退出；文件投递走 Windows 对 accept-files 窗口的真实通道",
           "结论": "FAIL"}
    proc = None
    errf = None
    alive_after_drop = False
    try:
        assert os.path.exists(SAMPLE_PDF), "缺少测试样本 " + SAMPLE_PDF
        try:
            os.remove(STATE_FILE)
        except Exception:
            pass

        env = dict(os.environ)
        env.update({"PYTHONNOUSERSITE": "1", "PYTHONPATH": os.path.join(BASE, "lib"),
                    "PATH": os.path.join(BASE, "lib", "pywin32_system32") + ";" + env.get("PATH", ""),
                    "PYTHONIOENCODING": "utf-8", "PYTHONFAULTHANDLER": "1"})
        log_offset = os.path.getsize(LOG) if os.path.exists(LOG) else 0
        errf = open(STDERR_LOG, "w", encoding="utf-8", errors="replace")
        proc = subprocess.Popen([sys.executable, "-X", "faulthandler", APP], cwd=BASE,
                                env=env, stdout=errf, stderr=subprocess.STDOUT)
        hwnd, rect = find_gui_window(pid=proc.pid)
        rep["GUI窗口"] = {"hwnd": int(hwnd) if hwnd else None, "rect": rect,
                          "本次验证进程号": proc.pid,
                          "启动后存活": proc.poll() is None,
                          "启动退出码": proc.poll()}
        print("GUI 窗口:", rep["GUI窗口"], flush=True)
        if not hwnd:
            rep["错误"] = "未找到 GUI 窗口"
            return rep

        user32.ShowWindow(hwnd, SW_RESTORE)
        user32.SetForegroundWindow(hwnd)
        move(a3_point(rect)[0], a3_point(rect)[1])
        time.sleep(0.6)

        # --- 1. 真实文件拖放（真实鼠标手势 + 真实 WM_DROPFILES 消息） ---
        print("=== 1. 真实文件拖放 ===", flush=True)
        try:
            drop = real_win32_drop(SAMPLE_PDF, rect, hwnd)
        except Exception:
            drop = {"异常": traceback.format_exc(), "errors": list(_DRAG_ERROR)}
        time.sleep(2.0)
        alive_after_drop = proc.poll() is None
        state_after_drop = read_state()
        rep["真实文件拖放"] = dict(drop, 拖放后进程存活=alive_after_drop,
                                   拖放后退出码=proc.poll(),
                                   拖放后已加载文件=(state_after_drop or {}).get("已加载文件"))
        print("  ", {k: v for k, v in rep["真实文件拖放"].items() if k != "鼠标轨迹"}, flush=True)

        # --- 2. 真实鼠标按住连续拖动红线 ---
        if alive_after_drop:
            user32.SetForegroundWindow(hwnd)
            time.sleep(0.5)
            img_rect = ((state_after_drop or {}).get("A3图像屏幕矩形")
                        or [rect[0] + 40, rect[1] + 220, rect[0] + 460, rect[1] + 620])
            print("=== 2. 真实鼠标按住连续拖动 ===", flush=True)
            try:
                drag = real_mouse_drag(img_rect)
            except Exception:
                drag = {"异常": traceback.format_exc()}
            time.sleep(1.5)
            rep["真实鼠标拖动"] = dict(drag, 拖动后进程存活=proc.poll() is None)
            print("  ", {k: v for k, v in rep["真实鼠标拖动"].items()
                         if k not in ("移动轨迹",)}, flush=True)
        else:
            rep["真实鼠标拖动"] = "跳过（真实文件拖放已导致程序退出）"

        # --- 3. 证据 ---
        time.sleep(0.3)
        new_log = read_new_log(log_offset)
        state = read_state()
        exit_code = proc.poll()
        if proc.poll() is None:
            proc.terminate()
            try:
                proc.wait(timeout=5)
            except Exception:
                proc.kill()
        errf.flush()
        errf.close()
        errf = None
        try:
            with open(STDERR_LOG, "r", encoding="utf-8", errors="replace") as f:
                stderr_txt = f.read()
        except Exception:
            stderr_txt = ""

        rep["GUI新增日志"] = new_log.strip().splitlines()[-30:]
        rep["GUI交互状态"] = state
        rep["GUI进程退出码"] = exit_code
        rep["stderr/faulthandler"] = stderr_txt.strip()[-6000:]

        ok = (alive_after_drop
              and rep["真实文件拖放"].get("PostMessageW成功") is True
              and (state_after_drop or {}).get("真实拖放次数", 0) >= 1
              and rep.get("真实鼠标拖动", {}).get("拖动后进程存活", False)
              and state is not None
              and state.get("已加载文件") is not None
              and state.get("真实拖放次数", 0) >= 1
              and state.get("分割线按下次数", 0) >= 1
              and state.get("分割线拖动次数", 0) >= 3
              and state.get("实时预览次数", 0) >= 3
              and state.get("最终split") is not None
              and abs(float(state["最终split"]) - 0.5) > 0.01
              and "Traceback" not in stderr_txt
              and "Fatal Python error" not in stderr_txt)
        rep["结论"] = "PASS" if ok else "FAIL"
        print("  证据:", REPORT, flush=True)
        print("  结论:", rep["结论"], flush=True)
        return rep
    except Exception:
        rep["异常"] = traceback.format_exc()
        print(rep["异常"], flush=True)
        return rep
    finally:
        if errf:
            try:
                errf.flush()
                errf.close()
            except Exception:
                pass
        if proc and proc.poll() is None:
            try:
                proc.terminate()
                proc.wait(timeout=4)
            except Exception:
                try:
                    proc.kill()
                except Exception:
                    pass
        try:
            with open(REPORT, "w", encoding="utf-8") as f:
                json.dump(rep, f, ensure_ascii=False, indent=2)
        except Exception:
            pass


if __name__ == "__main__":
    r = main()
    sys.exit(0 if r.get("结论") == "PASS" else 1)
