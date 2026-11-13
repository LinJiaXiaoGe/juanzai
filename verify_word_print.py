# -*- coding: utf-8 -*-
"""
v8 无纸验收：一键直印 + 隐藏 Word COM 直印 + 失败回退（2026-11-06）

老板退回整改的验收点，全部在「不出纸、不创建真实 Word、不发送任何打印任务」的前提下覆盖：
  1) 打印参数：每张 A4 独立一页，PrintOut(Pages="1", Background=False, Copies=1)
  2) 指定打印机端口：ActivePrinter = 「打印机名 on 端口名」，端口来自只读 Win32 事实
  3) Word 不外显：独立隐藏实例、Visible=False、DisplayAlerts=0
  4) 文档不保存：Close(SaveChanges=0) / Quit(SaveChanges=0)，从不 Documents.Open，用户文档哈希不变
  5) 临时文件清理：临时 .docx 与全部临时图片在返回前删除，零残留
  6) 失败回退：Word 提交异常 / 842 事件 0x6BE → 中文提示 + 自动保留 A4 PDF
  7) 一键直印：删除「选择打印通道」弹窗，do_print 中不再有任何「是/否」询问
  8) 不把提交当成功：只有核验「失败」才弹中文提示；非失败结论不弹任何成功提示
  9) 其它常规驱动实体打印机自动走 GDI

运行：C:\\Python314\\python.exe verify_word_print.py
"""
import _bootstrap  # 必须在任何第三方库之前

import os
import sys
import json
import hashlib
import datetime
import subprocess
import traceback

from PIL import Image

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "测试输出")
TEMP = os.path.join(OUT, "v8无纸临时",
                    datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f"))
SAMPLE = os.path.join(BASE, "测试样例", "sample_seal_multipage.pdf")
HP = "HP DeskJet 4900 series"
GUI_JSON = os.path.join(OUT, "v8无纸验收_GUI.json")
REPORT = os.path.join(OUT, "v8无纸验收报告.json")

import app
import printguard as pg
import wordprint as wp


# ---------------- 假 Word COM（绝不创建真实 Word、绝不提交打印任务） ----------------
class _Rec(object):
    """记录所有属性赋值，便于断言打印参数。"""

    def __init__(self, calls, prefix):
        object.__setattr__(self, "_calls", calls)
        object.__setattr__(self, "_prefix", prefix)
        object.__setattr__(self, "_v", {})

    def __setattr__(self, k, v):
        if k.startswith("_"):
            object.__setattr__(self, k, v)
            return
        self._calls.append((self._prefix + "." + k, v))
        self._v[k] = v

    def __getattr__(self, k):
        if k.startswith("_"):
            raise AttributeError(k)
        return self._v.get(k)


class FakeContent(object):
    End = 100


class FakeRange(object):
    def __init__(self, calls):
        self._calls = calls

    def InsertBreak(self, kind):
        self._calls.append(("Range.InsertBreak", kind))


class FakeShape(_Rec):
    def __init__(self, calls, prefix):
        super().__init__(calls, prefix)
        object.__setattr__(self, "WrapFormat", _Rec(calls, prefix + ".WrapFormat"))


class FakeShapes(object):
    def __init__(self, calls):
        self._calls = calls

    def AddPicture(self, *args):
        self._calls.append(("Shapes.AddPicture",
                            {"文件": os.path.basename(str(args[0])),
                             "Left": args[3], "Top": args[4],
                             "Width": args[5], "Height": args[6]}))
        return FakeShape(self._calls, "Shape")


class FakeDoc(object):
    def __init__(self, calls, fail_print=False):
        self._calls = calls
        self._fail = fail_print
        self.PageSetup = _Rec(calls, "PageSetup")
        self.Content = FakeContent()
        self.Shapes = FakeShapes(calls)

    def Range(self, a=None, b=None):
        self._calls.append(("Document.Range", (a, b)))
        return FakeRange(self._calls)

    def SaveAs2(self, path, fmt=16):
        self._calls.append(("Document.SaveAs2", {"路径": path, "格式": fmt}))
        with open(path, "wb") as f:
            f.write(b"PK\x03\x04FAKE-TEMP-DOCX")

    def SaveAs(self, path, fmt=16):
        return self.SaveAs2(path, fmt)

    def PrintOut(self, **kwargs):
        self._calls.append(("Document.PrintOut", dict(kwargs)))
        if self._fail:
            raise RuntimeError("模拟：Word 提交打印失败")

    def Repaginate(self):
        self._calls.append(("Document.Repaginate", True))

    def ComputeStatistics(self, kind):
        self._calls.append(("Document.ComputeStatistics", kind))
        return 1

    def Close(self, **kwargs):
        self._calls.append(("Document.Close", dict(kwargs)))


class FakeDocuments(object):
    def __init__(self, calls, fail_print=False):
        self._calls = calls
        self._fail = fail_print

    def Add(self):
        self._calls.append(("Documents.Add", {}))
        return FakeDoc(self._calls, self._fail)

    def Open(self, *a, **k):
        self._calls.append(("Documents.Open", list(a)))
        raise AssertionError("禁止打开任何用户文档")


class FakeWord(object):
    def __init__(self, calls, fail_print=False):
        object.__setattr__(self, "_calls", calls)
        object.__setattr__(self, "_v", {"打开用户文档次数": 0})
        object.__setattr__(self, "Documents", FakeDocuments(calls, fail_print))

    def __setattr__(self, k, v):
        if k.startswith("_") or k == "Documents":
            object.__setattr__(self, k, v)
            return
        self._calls.append(("Word." + k, v))
        self._v[k] = v

    def __getattr__(self, k):
        if k.startswith("_"):
            raise AttributeError(k)
        return self._v.get(k)

    def Quit(self, **kwargs):
        self._calls.append(("Word.Quit", dict(kwargs)))


def make_factory(calls, fail_print=False):
    def factory():
        calls.append(("DispatchEx", "Word.Application"))   # 独立隐藏实例
        return FakeWord(calls, fail_print)
    return factory


def _find(calls, name):
    return [v for k, v in calls if k == name]


def _images(n=2):
    return [Image.new("RGB", (120, 170), (255, 255, 255)) for _ in range(n)]


def _sha(path):
    with open(path, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest()


# ---------------- 1) 模块级无纸测试 ----------------
def run_module_checks():
    os.makedirs(TEMP, exist_ok=True)
    rep = {}

    # --- 1.1 Word 直印成功路径：打印参数 / 端口 / 不外显 / 不保存 / 清理 ---
    calls = []
    rec = {}
    user_doc = os.path.join(TEMP, "用户文档_原件.docx")
    with open(user_doc, "wb") as f:
        f.write(b"USER-DOCUMENT-ORIGINAL-CONTENT")
    user_before = (_sha(user_doc), os.stat(user_doc).st_mtime_ns)

    r = wp.print_via_word(HP, _images(2), port="USB001",
                          word_factory=make_factory(calls), temp_dir=TEMP, recorder=rec)

    ps = dict((k, v) for k, v in calls if k.startswith("PageSetup."))
    po = _find(calls, "Document.PrintOut")
    close = _find(calls, "Document.Close")
    quit_ = _find(calls, "Word.Quit")
    adds = _find(calls, "Shapes.AddPicture")
    shapes = {}
    for k, v in calls:
        if k.startswith("Shape.") and k not in shapes:
            shapes[k] = v
    printers = _find(calls, "Word.ActivePrinter")

    assert po == [{"Background": False, "Copies": 1, "Pages": "1"}] * 2, po
    assert ps.get("PageSetup.PaperSize") == wp.WD_PAPER_A4, ps
    assert ps.get("PageSetup.Orientation") == wp.WD_ORIENT_PORTRAIT, ps
    assert abs(ps.get("PageSetup.PageWidth") - wp.A4_W_PT) < 0.001, ps
    assert abs(ps.get("PageSetup.PageHeight") - wp.A4_H_PT) < 0.001, ps
    assert all(abs(ps.get("PageSetup." + k) - wp.SAFE_MARGIN_PT) < 0.001 for k in
               ("TopMargin", "BottomMargin", "LeftMargin", "RightMargin")), ps
    assert ps.get("PageSetup.Gutter") == 0, ps
    assert len(adds) == 2 and shapes.get("Shape.LockAspectRatio") == 0, (adds, shapes)
    assert abs(shapes.get("Shape.Width") - (wp.A4_W_PT - 2 * wp.SAFE_MARGIN_PT)) < 0.001, shapes
    assert abs(shapes.get("Shape.Height") - (wp.A4_H_PT - 2 * wp.SAFE_MARGIN_PT)) < 0.001, shapes
    assert _find(calls, "Document.ComputeStatistics") == [2, 2], calls
    assert printers == [f"{HP} on USB001"], printers
    assert _find(calls, "Word.Visible") == [False], calls
    assert _find(calls, "Word.DisplayAlerts") == [wp.WD_ALERTS_NONE], calls
    assert _find(calls, "DispatchEx") == ["Word.Application"], calls
    assert _find(calls, "Documents.Add") == [{}, {}], calls
    assert not _find(calls, "Documents.Open"), "绝不打开用户文档"
    assert close == [{"SaveChanges": 0}] * 2, close
    assert quit_ == [{"SaveChanges": 0}], quit_
    assert r["Word不外显"] is True and r["未打开用户文档"] is True, r
    assert r["临时文件清理"]["全部清理"] is True, r["临时文件清理"]
    assert r["临时文件清理"]["创建"] >= 2, r["临时文件清理"]
    assert not any(os.path.exists(p) for p in r["临时图片"]), "临时文件必须全部删除"
    assert [x["实际打印页"] for x in r["每页排版"]] == ["1", "1"], r["每页排版"]
    assert (_sha(user_doc), os.stat(user_doc).st_mtime_ns) == user_before, "用户文档被改动"
    assert rec["临时文件已清理"] is True and rec["提交完成"] is True, rec
    rep["打印参数与A4页面"] = {
        "PrintOut": po[0], "纸张": "A4", "方向": "纵向",
        "页面pt": [round(wp.A4_W_PT, 4), round(wp.A4_H_PT, 4)],
        "安全边距pt": round(wp.SAFE_MARGIN_PT, 4), "嵌入A4页数": len(adds),
        "图片宽pt": round(shapes["Shape.Width"], 4),
        "图片高pt": round(shapes["Shape.Height"], 4),
        "每个临时文档排版页数": [x["排版总页数"] for x in r["每页排版"]],
        "实际打印页": [x["实际打印页"] for x in r["每页排版"]],
        "判定": "PASS"}
    rep["指定打印机端口"] = {"ActivePrinter": printers[0], "端口": "USB001",
                          "无端口时": wp.build_active_printer(HP, None)[0],
                          "虚拟端口时": wp.build_active_printer(HP, "PORTPROMPT:")[0],
                          "判定": "PASS"}
    rep["Word不外显与不保存"] = {"Visible": _find(calls, "Word.Visible")[0],
                             "DisplayAlerts": _find(calls, "Word.DisplayAlerts")[0],
                             "实例": "DispatchEx 独立隐藏实例",
                             "文档关闭": close[0], "Word退出": quit_[0],
                             "Documents.Open次数": 0,
                             "用户文档哈希不变": True, "判定": "PASS"}
    rep["临时文件清理"] = dict(r["临时文件清理"], 判定="PASS")

    # --- 1.2 Word 提交失败：仍关闭不保存、仍清理临时文件 ---
    calls2 = []
    rec2 = {}
    try:
        wp.print_via_word(HP, _images(1), port="USB001",
                          word_factory=make_factory(calls2, fail_print=True),
                          temp_dir=TEMP, recorder=rec2)
        raise AssertionError("Word 提交失败应当抛出异常")
    except RuntimeError:
        pass
    assert _find(calls2, "Document.Close") == [{"SaveChanges": 0}], calls2
    assert _find(calls2, "Word.Quit") == [{"SaveChanges": 0}], calls2
    assert rec2["临时文件已清理"] is True and rec2["提交完成"] is False, rec2
    assert not rec2["残留临时文件"], rec2
    rep["Word提交异常收尾"] = {"异常已上抛": True, "关闭不保存": True,
                           "临时文件零残留": True, "判定": "PASS"}

    # --- 1.3 一键直印：源码级删除通道选择弹窗 ---
    with open(os.path.join(BASE, "app.py"), "r", encoding="utf-8") as f:
        src = f.read()
    i = src.index("        def do_print(self):")
    j = src.find("\n    root = tk.Tk()", i)
    body = src[i:j if j > 0 else len(src)]
    assert "选择打印通道" not in src, "仍残留通道选择弹窗文案"
    assert "askyesno" not in body and "askokcancel" not in body, "do_print 仍有是/否询问"
    assert "pdf_channel_print()" not in body, "do_print 仍会把打印改道为只生成文件"
    assert "self.direct_print_with_verification(pn)" in body, "点打印未直接提交"
    rep["删除通道选择弹窗"] = {"源码含选择打印通道": False,
                          "do_print含是/否询问": False,
                          "点打印直接提交": True, "判定": "PASS"}

    # --- 1.4 自动判定：V4/失败史 → Word；常规驱动 → GDI ---
    v4 = {"驱动": "HP DeskJet 4900 series PCL-3 (V4)", "端口": "USB001"}
    normal = {"驱动": "HP Universal Printing PCL 6", "端口": "USB002"}
    b1, why1, _ = app.choose_print_backend("模拟V4实体打印机", v4)
    b2, why2, _ = app.choose_print_backend("模拟常规驱动实体打印机", normal)
    assert b1 == "word", (b1, why1)
    assert b2 == "gdi", (b2, why2)
    rep["自动选择提交方式"] = {"V4或XPS驱动": b1, "常规驱动": b2,
                          "V4原因": why1, "常规原因": why2, "判定": "PASS"}

    # --- 1.5 提交后核验：重放 842 / MS_XPS_PROC / 0x6BE，绝不把提交当成功 ---
    saved = {"word": wp.print_via_word, "gdi": app.print_a4_images,
             "record": pg.record_channel_result}
    gdi_calls = []
    try:
        wp.print_via_word = lambda printer, imgs, port=None: {
            "提交方式": "Word COM 直印（桩）", "页数": len(imgs), "端口": port}
        app.print_a4_images = lambda *a, **k: gdi_calls.append(1) or {"pages": 2}
        pg.record_channel_result = lambda *a, **k: {}

        def events():
            now = datetime.datetime.now(datetime.timezone.utc).replace(tzinfo=None)
            return [{"事件ID": "842", "记录号": "4895",
                     "时间UTC": now.strftime("%Y-%m-%d %H:%M:%S"),
                     "_时间": now, "作业号": 7,
                     "打印机": HP, "驱动": "HP DeskJet 4900 series PCL-3 (V4)",
                     "打印处理器": "MS_XPS_PROC", "错误码": "0X6BE", "是否错误": True}]

        r2 = app.print_with_verification(
            HP, _images(2), wait_seconds=0,
            facts_provider=lambda n: dict(v4),
            jobs_provider=lambda: [], events_provider=events)
        ver = r2["核验"]
        assert r2["提交方式"] == "word", r2
        assert not gdi_calls, "Word 通道不得再走 GDI 提交"
        assert ver["结论"] == "失败", ver
        assert ver["842事件数"] == 1 and "0X6BE" in ver["错误代码"], ver
        assert "842 事件" in ver["中文说明"], ver["中文说明"]
        rep["提交后核验842事件"] = {
            "提交方式": r2["提交方式"], "端口": r2["打印机端口"],
            "结论": ver["结论"], "842事件数": ver["842事件数"],
            "错误代码": ver["错误代码"], "中文说明": ver["中文说明"],
            "GDI提交次数": len(gdi_calls), "判定": "PASS"}
    finally:
        wp.print_via_word = saved["word"]
        app.print_a4_images = saved["gdi"]
        pg.record_channel_result = saved["record"]

    rep["真实打印任务数"] = 0
    rep["真实Word实例数"] = 0
    rep["实际出纸张数"] = 0
    return rep


# ---------------- 2) 真实 App.do_print 的无纸 GUI 验收 ----------------
def run_gui_checks(mod, AppClass, out_path, sample):
    """由 app.py --word-print-probe 调用：真实 Tk App，提交/弹窗/打开全部桩化。"""
    import tkinter as tk
    import tkinter.messagebox as mb

    captured = []
    originals = {}
    for name, default in (("showerror", "ok"), ("showwarning", "ok"),
                          ("showinfo", "ok"), ("askyesno", True),
                          ("askokcancel", True)):
        originals[name] = getattr(mb, name)

        def rec(n=name, d=default):
            def _f(*a, **k):
                captured.append({"函数": n,
                                 "标题": str(a[0] if a else k.get("title", "")),
                                 "正文": str(a[1] if len(a) > 1 else k.get("message", ""))})
                return d
            return _f
        setattr(mb, name, rec())

    saved = {"list": mod.list_printers, "facts": mod.read_printer_facts,
             "submit": mod.print_with_verification,
             "state": mod.PRINTER_STATE_FILE, "gui_state": mod.GUI_STATE_PATH,
             "openfolder": pg.open_folder, "startfile": os.startfile}
    result = {"生成时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
              "判定": "FAIL"}
    root = None
    try:
        mod.PRINTER_STATE_FILE = os.path.join(OUT, "v8模拟_打印机状态.json")
        mod.GUI_STATE_PATH = os.path.join(OUT, "v8模拟_GUI交互状态.json")
        mod.list_printers = lambda: [HP]
        mod.read_printer_facts = lambda n: {
            "端口": "USB001", "驱动": "HP DeskJet 4900 series PCL-3 (V4)",
            "备注": "", "位置": "", "系统属性": 0x40, "状态位": 0,
            "打印处理器": "MS_XPS_PROC", "打开成功": True}
        pg.open_folder = lambda p: True
        os.startfile = lambda *a, **k: True
        submits = []
        conclusion = {"结论": "失败",
                      "中文说明": "系统打印服务报告打印作业失败，页面内容没有到达打印机。"
                                  " 已核验打印队列与打印服务 842 事件，本次作业被报告为失败。"}

        def fake_submit(printer, imgs):
            submits.append({"打印机": printer, "页数": len(imgs)})
            if conclusion["结论"] == "抛出异常":
                raise RuntimeError("模拟：Word 直印提交失败")
            return {"提交结果": {"提交方式": "Word COM 直印（桩）"},
                    "核验": dict(conclusion), "提交方式": "word"}
        mod.print_with_verification = fake_submit

        root = tk.Tk()
        root.withdraw()
        gui = AppClass(root)
        gui._write_state = lambda: None
        gui.pages = mod.load_document(sample)[:1]

        # 场景 A：核验失败 → 只弹中文错误、自动保留 A4 PDF、绝不出现「选择是否」或成功提示
        gui.do_print()
        fb = gui.last_fallback
        assert fb and len(fb["文件"]) >= 3, fb
        assert all(os.path.exists(f["路径"]) for f in fb["文件"]), "应自动保留 A4 PDF"
        assert all(f["页数"] >= 1 for f in fb["文件"]), fb
        assert len(submits) == 1, "点一次打印只提交一次"
        assert not [c for c in captured if c["函数"] in ("askyesno", "askokcancel")], \
            "仍存在通道选择 / 是或否弹窗"
        assert not [c for c in captured if c["函数"] == "showinfo"], \
            "失败时不得出现任何成功提示"
        errs = [c for c in captured if c["函数"] == "showerror"]
        assert errs and all(c["标题"] and c["正文"] for c in errs), captured
        import re
        assert all(re.search(r"[\u4e00-\u9fff]", c["正文"]) for c in captured), \
            "错误提示必须是简体中文"
        result["失败场景"] = {
            "提交次数": len(submits), "弹窗函数": [c["函数"] for c in captured],
            "是否询问选择": False, "自动保留A4文件数": len(fb["文件"]),
            "保留文件均存在": True, "中文错误提示": errs[0]["标题"],
            "成功提示数": 0, "判定": "PASS"}

        # 场景 B：非失败结论 → 不弹任何成功提示，不把提交当成功
        captured.clear()
        conclusion.update({"结论": "未检出错误",
                           "中文说明": "打印作业已离开系统队列，且未在队列或打印服务事件中检出错误。"
                                       "本工具无法确认纸张是否真正输出，请检查打印机出纸口。"})
        gui.do_print()
        assert len(submits) == 2, submits
        assert not captured, "非失败结论不得弹出成功提示（不把提交当成功）"
        assert "请检查打印机出纸口" in gui.status.cget("text") or \
            "提交并核验" in gui.status.cget("text"), gui.status.cget("text")
        result["非失败场景"] = {"提交次数": len(submits), "弹窗数": 0,
                           "状态栏": gui.status.cget("text"),
                           "宣称打印成功": False, "判定": "PASS"}

        # 场景 C：Word 提交异常 → 中文提示 + 自动保留 A4 PDF
        captured.clear()
        conclusion.update({"结论": "抛出异常"})
        gui.do_print()
        assert len(submits) == 3, submits
        assert any(c["函数"] == "showerror" for c in captured), captured
        assert not [c for c in captured if c["函数"] in ("askyesno", "askokcancel")], captured
        assert gui.last_fallback and all(os.path.exists(f["路径"])
                                         for f in gui.last_fallback["文件"])
        result["提交异常场景"] = {"提交次数": len(submits),
                           "中文错误提示": True, "自动保留A4文件": True,
                           "是否询问选择": False, "判定": "PASS"}

        result["真实打印任务数"] = 0
        result["实际出纸张数"] = 0
        result["判定"] = "PASS"
        return result
    except Exception as e:
        result["异常"] = f"{type(e).__name__}: {e}"
        result["堆栈"] = traceback.format_exc()
        raise
    finally:
        for name, fn in originals.items():
            setattr(mb, name, fn)
        mod.list_printers = saved["list"]
        mod.read_printer_facts = saved["facts"]
        mod.print_with_verification = saved["submit"]
        mod.PRINTER_STATE_FILE = saved["state"]
        mod.GUI_STATE_PATH = saved["gui_state"]
        pg.open_folder = saved["openfolder"]
        os.startfile = saved["startfile"]
        if root is not None:
            try:
                root.destroy()
            except Exception:
                pass
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)


# ---------------- 3) 汇总 ----------------
def main():
    print("=" * 72)
    print("v8 无纸验收：一键直印 + 隐藏 Word COM 直印 + 842 事件核验 + 失败回退")
    print("=" * 72)
    os.makedirs(OUT, exist_ok=True)
    rep = {"生成时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
           "版本": "v8", "判定": "FAIL"}
    rep["模块级测试"] = run_module_checks()
    for k, v in rep["模块级测试"].items():
        if isinstance(v, dict):
            print(f"  [模块] {k}: {v.get('判定', '')}")

    code = subprocess.call([sys.executable, os.path.join(BASE, "app.py"),
                            "--word-print-probe"], cwd=BASE, timeout=180)
    with open(GUI_JSON, "r", encoding="utf-8") as f:
        gui = json.load(f)
    assert code == 0 and gui.get("判定") == "PASS", (code, gui.get("异常"))
    rep["真实GUI测试"] = gui
    for k in ("失败场景", "非失败场景", "提交异常场景"):
        print(f"  [GUI] {k}: {gui[k]['判定']}")
    rep.update({"真实打印任务数": 0, "真实Word实例数": 0, "实际出纸张数": 0,
                "判定": "PASS"})
    with open(REPORT, "w", encoding="utf-8") as f:
        json.dump(rep, f, ensure_ascii=False, indent=2)
    print("\n结论: PASS | 真实打印任务 0 | 真实 Word 实例 0 | 出纸 0")
    print("证据:", REPORT, "与", GUI_JSON)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        traceback.print_exc()
        sys.exit(2)
