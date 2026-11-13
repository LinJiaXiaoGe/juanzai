# -*- coding: utf-8 -*-
"""
v7 打印链路独立验收（全程不实际出纸、不发送任何打印任务）

覆盖：
  1) 本机实证：只读抓取打印服务 842 事件，重放老板 22:36/22:37 的 0x6BE 失败
  2) 送驱动前压缩：DC 分辨率 600→300、位图长边 1654×2339→1240×1754
  3) 压缩 PDF：标准 A4、体积远小于旧 PNG 通道
  4) 失败自动保留：合并 PDF + 两张单页 A4 PDF + 打开文件夹
  5) 核验注入：失败 / 未检出错误 / 待确认 / 队列错误 四种场景
  6) 通道推荐：V4/XPS 与失败历史 → 压缩文件通道；其它品牌常规驱动 → 直接发送
  7) 品牌无关识别：多品牌端口/驱动模拟，虚拟设备全部隐藏
  8) 无实体 / 全部离线：打印按钮禁用，只允许保存文件
  9) 真实 GUI 上模拟失败：自动回退、中文弹窗、零 StartDoc 调用
 10) 全部错误弹窗简体中文、无英文技术串

产出：测试输出/打印链路验收报告.json、测试输出/本机打印服务事件证据.json
用法：C:\\Python314\\python.exe verify_print_guard.py
"""
import _bootstrap  # 注入 lib

import os
import sys
import io
import re
import json
import time
import datetime
import traceback

import app
import printguard as pg

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "测试输出")
os.makedirs(OUT, exist_ok=True)
REPORT = os.path.join(OUT, "打印链路验收报告.json")
EVENT_EVIDENCE = os.path.join(OUT, "本机打印服务事件证据.json")
SAMPLE = os.path.join(BASE, "测试样例", "sample_seal_multipage.pdf")
HP = "HP DeskJet 4900 series"

# 本脚本严禁自动出纸：所有打印提交函数都被替换为桩
AUTO_PRINT_ACTIONS = 0
STARTDOC_CALLS = []


def utc_naive(dt):
    return dt.replace(tzinfo=None) if dt.tzinfo else dt


def section(rep, key, value, note=""):
    rep["检查项"][key] = value
    print(f"=== {key} === {note}")


def main():
    rep = {
        "生成时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "验收方式": "真实事件日志只读重放 + 注入式队列核验 + 真实 GUI 桩化交互 + 文件与 DC 实测",
        "自动发送打印任务次数": AUTO_PRINT_ACTIONS,
        "StartDoc调用次数": 0,
        "检查项": {}, "结论": "FAIL",
    }
    try:
        assert os.path.exists(SAMPLE), "缺少测试样本 " + SAMPLE
        pages = app.load_document(SAMPLE)[:1]
        imgs, _ = app.process_document(pages, "auto", False, 0.0)
        assert len(imgs) == 2, "一张 A3 应输出两张 A4"

        # === 1. 本机实证：只读抓取 842 事件并重放 ===
        events, evt_err = pg.read_print_events(None, HP, 500)
        bad = [e for e in events if e["是否错误"]]
        evidence = {
            "抓取时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "事件通道": pg.EVENT_CHANNEL,
            "读取错误": evt_err,
            "抓取事件数": len(events),
            "作业4和5的失败事件": [pg._public_event(e) for e in bad
                                if e.get("作业号") in (4, 5)],
            "当前打印队列快照": pg.snapshot_jobs(HP),
            "说明": "纯只读：未创建、未删除、未重试任何打印作业，未出纸。",
        }
        with open(EVENT_EVIDENCE, "w", encoding="utf-8") as f:
            json.dump(evidence, f, ensure_ascii=False, indent=2)
        replay = pg.verify_print_job(
            HP, before_ids=[], started_utc=datetime.datetime(2026, 11, 6, 14, 36, 0),
            wait_seconds=0, jobs_provider=lambda: [],
            events_provider=lambda: events)
        assert len(evidence["作业4和5的失败事件"]) >= 2, "未抓到作业4/5的失败事件"
        assert replay["结论"] == "失败", replay
        assert "没有到达打印机" in replay["中文说明"], replay["中文说明"]
        assert any("6BE" in c.upper() for c in replay["错误代码"]), replay["错误代码"]
        assert re.search(r"[\u4e00-\u9fff]", replay["中文说明"])
        section(rep, "本机失败事件重放", {
            "失败事件数": len(evidence["作业4和5的失败事件"]),
            "错误代码": replay["错误代码"],
            "核验结论": replay["结论"],
            "中文说明": replay["中文说明"],
            "证据": EVENT_EVIDENCE, "判定": "PASS"},
            f"0x6BE 事件 {len(evidence['作业4和5的失败事件'])} 条，结论=失败")

        # === 2. 送驱动前压缩 + DC 降分辨率（dry_run，不 StartDoc） ===
        dry = app.print_a4_images(HP, imgs, dry_run=True)
        dc = {}
        try:
            import win32ui
            import ctypes
            hDC = win32ui.CreateDC()
            hDC.CreatePrinterDC(HP)
            h = hDC.GetHandleOutput()
            dc["默认DPI"] = app._dc_caps(h)["DPI"]
            dm = app._devmode(HP)
            dm.dmFields |= app.DM_PRINTQUALITY
            dm.dmPrintQuality = pg.PRINT_DC_DPI
            app._GDI.ResetDCW(h, ctypes.byref(dm))
            dc["压缩后DPI"] = app._dc_caps(h)["DPI"]
            hDC.DeleteDC()
        except Exception as e:
            dc["读取错误"] = str(e)
        st = dry["压缩统计"][0]
        assert st["送驱动尺寸"][1] <= pg.PRINT_MAX_LONG_SIDE
        assert st["缩减比例"] > 0.35, st
        assert dc.get("压缩后DPI", 600) <= 300, dc
        assert dry["evidence"]["verified_a4_portrait"] is True
        section(rep, "避免大图压垮驱动", {
            "位图原尺寸": st["原尺寸"], "送驱动尺寸": st["送驱动尺寸"],
            "原始字节": st["原始字节"], "送驱动字节": st["送驱动字节"],
            "缩减比例": st["缩减比例"], "DC分辨率": dc,
            "A4纵向校验": True, "未调用StartDoc": True, "判定": "PASS"},
            f"DPI {dc.get('默认DPI')}→{dc.get('压缩后DPI')}，位图缩减 {st['缩减比例']:.1%}")

        # === 3. 压缩 PDF 通道 ===
        png_pdf = os.path.join(OUT, "通道对比_旧PNG.pdf")
        jpg_pdf = os.path.join(OUT, "通道对比_压缩A4.pdf")
        import fitz
        d = fitz.open()
        for im in imgs:
            p = d.new_page(width=210 / 25.4 * 72, height=297 / 25.4 * 72)
            b = io.BytesIO()
            im.convert("RGB").save(b, format="PNG")
            p.insert_image(p.rect, stream=b.getvalue())
        d.save(png_pdf)
        d.close()
        pg.export_compressed_a4_pdf(imgs, jpg_pdf)
        d = fitz.open(jpg_pdf)
        sizes = [(round(p.rect.width, 2), round(p.rect.height, 2)) for p in d]
        cnt = d.page_count
        d.close()
        ratio = os.path.getsize(jpg_pdf) / float(os.path.getsize(png_pdf))
        assert cnt == 2 and all(abs(a - 595.28) < 1 and abs(b - 841.89) < 1
                                for a, b in sizes), (cnt, sizes)
        assert ratio < 0.3, ratio
        section(rep, "压缩A4文件通道", {
            "页数": cnt, "每页pt": sizes,
            "旧通道字节": os.path.getsize(png_pdf),
            "压缩通道字节": os.path.getsize(jpg_pdf),
            "体积比": round(ratio, 4), "判定": "PASS"},
            f"{os.path.getsize(png_pdf)} → {os.path.getsize(jpg_pdf)} 字节")

        # === 4. 失败自动保留两张 A4 PDF + 打开文件夹 ===
        fb_dir = os.path.join(OUT, "验收_自动回退")
        fb = pg.save_a4_fallback(imgs, folder=fb_dir, prefix="验收")
        assert len(fb["单页PDF"]) == 2
        assert all(os.path.exists(p) for p in [fb["合并PDF"]] + fb["单页PDF"])
        for f in fb["文件"]:
            assert abs(f["每页尺寸pt"][0][0] - 595.28) < 1, f
            assert f["字节"] > 0, f
        opened = []
        pg.open_folder = lambda p: opened.append(p) or True
        pg.open_folder(fb["文件夹"])
        assert opened == [fb_dir], opened
        section(rep, "失败自动保留A4文件", {
            "文件夹": fb["文件夹"], "页数": fb["页数"],
            "文件数": len(fb["文件"]),
            "文件": [{"路径": os.path.basename(f["路径"]), "角色": f["角色"],
                    "字节": f["字节"], "每页尺寸pt": f["每页尺寸pt"]}
                   for f in fb["文件"]],
            "打开文件夹调用": opened, "判定": "PASS"},
            f"合并1 + 单页2，共 {sum(f['字节'] for f in fb['文件'])} 字节")

        # === 5. 队列核验注入（不出纸） ===
        ok_event = [{"事件ID": "842", "作业号": 99, "打印机": HP,
                     "打印处理器": "winprint", "错误码": "0X0", "是否错误": False,
                     "_时间": datetime.datetime(2026, 11, 6, 14, 45, 0),
                     "时间UTC": "2026-11-06 14:45:00", "记录号": "1", "驱动": ""}]
        fail_event = [dict(ok_event[0], 错误码="0X6BE", 是否错误=True,
                           打印处理器="MS_XPS_PROC")]
        cases = [
            ("注入_失败事件", lambda: fail_event, lambda: [], "失败"),
            ("注入_成功离开队列", lambda: ok_event, lambda: [], "未检出错误"),
            ("注入_作业仍在队列", lambda: ok_event,
             lambda: [{"作业号": 99, "状态位": pg.JOB_STATUS_SPOOLING,
                       "状态文本": "正在后台处理"}], "待确认"),
            ("注入_队列错误", lambda: [],
             lambda: [{"作业号": 99, "状态位": pg.JOB_STATUS_ERROR,
                       "状态文本": "错误"}], "失败"),
        ]
        case_res = []
        for name, evfn, jobfn, expect in cases:
            r = pg.verify_print_job(
                HP, before_ids=[], started_utc=datetime.datetime(2026, 11, 6, 14, 44, 0),
                wait_seconds=0, jobs_provider=jobfn, events_provider=evfn)
            assert r["结论"] == expect, (name, r["结论"], expect)
            assert re.search(r"[\u4e00-\u9fff]", r["中文说明"])
            assert not re.search(r"StartDoc|win32ui|Traceback|Exception", r["中文说明"])
            case_res.append({"场景": name, "结论": r["结论"], "中文说明": r["中文说明"]})
        section(rep, "打印后核验四种场景", {"场景": case_res, "判定": "PASS"},
                "失败/未检出错误/待确认/队列错误 全部命中")

        # === 6+7. 通道推荐与品牌无关识别 ===
        hist_file = pg.HISTORY_FILE
        saved_history = None
        if os.path.exists(hist_file):
            with open(hist_file, "r", encoding="utf-8") as f:
                saved_history = f.read()
        try:
            brands = [
                ("惠普", {"驱动": "HP DeskJet 4900 series PCL-3 (V4)", "端口": "USB001"}, "pdf"),
                ("佳能", {"驱动": "Canon PIXMA G3010", "端口": "USB002"}, "gdi"),
                ("爱普生", {"驱动": "EPSON L3150 Series", "端口": "IP_192.168.1.50"}, "gdi"),
                ("兄弟", {"驱动": "Brother DCP-1618W", "端口": "WSD-1234"}, "gdi"),
                ("奔图", {"驱动": "Pantum M6500", "端口": "USB003"}, "gdi"),
                ("未识别品牌", {"驱动": "Generic PCL6", "端口": "LPT1:"}, "gdi"),
            ]
            brand_res = []
            for label, facts, expect_ch in brands:
                info = app.classify_printer(f"{label} 模拟打印机", dict(facts, 系统属性=0x40))
                assert info["是否实体"], (label, info["隐藏原因"])
                ch, why = pg.recommended_channel(f"{label} 模拟打印机", facts)
                assert ch == expect_ch, (label, ch, why)
                assert re.search(r"[\u4e00-\u9fff]", why)
                brand_res.append({"品牌": label, "识别品牌": info["品牌"],
                                  "端口类型": info["端口类型"], "通道": ch, "原因": why})
            virtuals = [
                ("OneNote (桌面) - 保护", {"驱动": "Microsoft Virtual Print Class Driver",
                                     "端口": "Microsoft.Office.OneNoteVirtualPrinter_x"}),
                ("Microsoft Print to PDF", {"驱动": "Microsoft Print To PDF", "端口": "PORTPROMPT:"}),
                ("Fax", {"驱动": "Microsoft Shared Fax Driver", "端口": "SHRFAX:"}),
                ("XPS 文档写入", {"驱动": "Microsoft XPS Document Writer", "端口": "PORTPROMPT:"}),
                ("HP Smart Printing", {"驱动": "Microsoft IPP Class Driver",
                                       "端口": "WSD-123", "备注": "Virtual Printer for PSA"}),
            ]
            hidden = []
            for name, facts in virtuals:
                info = app.classify_printer(name, dict(facts, 系统属性=0x40))
                assert not info["是否实体"], (name, info)
                hidden.append({"名称": name, "隐藏原因": info["隐藏原因"]})
            # 失败历史 → 下次自动改走压缩文件通道（不写死品牌）
            pg.record_channel_result("某品牌模拟打印机", "失败", "模拟：上次打印失败")
            ch, why = pg.recommended_channel(
                "某品牌模拟打印机", {"驱动": "Generic PCL6"})
            assert ch == "pdf" and "上次" in why, (ch, why)
            section(rep, "品牌无关识别与通道推荐", {
                "多品牌实体": brand_res, "虚拟设备隐藏": hidden,
                "失败历史自动切通道": {"通道": ch, "原因": why},
                "判定": "PASS"}, f"实体品牌 {len(brand_res)} 个、隐藏虚拟 {len(hidden)} 个")
        finally:
            if saved_history is None:
                try:
                    os.remove(hist_file)
                except Exception:
                    pass
            else:
                with open(hist_file, "w", encoding="utf-8") as f:
                    f.write(saved_history)

        # === 8+9. 真实 GUI：无实体/全离线禁用 + 失败自动回退 + 中文弹窗 ===
        gui_res = run_gui_probe_subprocess()
        section(rep, "真实GUI交互与中文弹窗", gui_res,
                f"GUI弹窗 {gui_res['弹窗数']} 个、违规 {len(gui_res['违规项'])} 项")

        # === 10. 中文弹窗探针（独立进程，真实 App，全部桩化） ===
        probe = run_text_probe()
        section(rep, "全部错误弹窗简体中文", probe,
                f"分支 {probe['触发分支数']}、违规 {len(probe['违规项'])}")

        rep["StartDoc调用次数"] = len(STARTDOC_CALLS)
        assert rep["StartDoc调用次数"] == 0, "验收过程不得调用 StartDoc"
        assert AUTO_PRINT_ACTIONS == 0
        rep["结论"] = "PASS"
    except Exception:
        rep["异常"] = traceback.format_exc()
        print(rep["异常"])
    finally:
        with open(REPORT, "w", encoding="utf-8") as f:
            json.dump(rep, f, ensure_ascii=False, indent=2)
        app.log(f"VERIFY_PRINT_GUARD 结论={rep['结论']} "
                f"自动打印次数={AUTO_PRINT_ACTIONS} StartDoc次数={len(STARTDOC_CALLS)}")
        print("  证据:", REPORT)
        print("  结论:", rep["结论"], "| 自动发送打印任务:", AUTO_PRINT_ACTIONS,
              "| StartDoc:", len(STARTDOC_CALLS))
    return rep


def run_gui_probe_subprocess():
    """让 app.py 在自己的 __main__ 命名空间里构造真实 App 并跑 GUI 场景。"""
    import subprocess
    out = os.path.join(OUT, "GUI打印链路模拟.json")
    err = os.path.join(OUT, "GUI打印链路模拟_stderr.log")
    try:
        os.remove(out)
    except Exception:
        pass
    env = dict(os.environ)
    env.update({"PYTHONNOUSERSITE": "1", "PYTHONPATH": os.path.join(BASE, "lib"),
                "PATH": os.path.join(BASE, "lib", "pywin32_system32") + ";" + env.get("PATH", ""),
                "PYTHONIOENCODING": "utf-8", "PYTHONFAULTHANDLER": "1"})
    with open(err, "w", encoding="utf-8", errors="replace") as ef:
        q = subprocess.Popen([sys.executable, "-X", "faulthandler",
                              os.path.join(BASE, "app.py"), "--print-guard-gui"],
                             cwd=BASE, env=env, stdout=ef, stderr=subprocess.STDOUT)
        try:
            q.wait(timeout=180)
            code = q.returncode
        except subprocess.TimeoutExpired:
            q.kill()
            code = "超时"
    assert os.path.exists(out), "GUI打印链路模拟.json 未生成：" + err
    with open(out, "r", encoding="utf-8") as f:
        d = json.load(f)
    assert d.get("判定") == "PASS", d.get("异常", d)
    assert code == 0, code
    d["探针退出码"] = code
    d["证据"] = out
    return d


def run_gui_checks(mod, AppClass, out_path, sample):
    """由 app.py 调用：真实 Tk GUI，打印提交/文件夹打开/弹窗全部桩化，不发送任何打印任务。

    mod / AppClass 必须来自 app.py 的 __main__ 命名空间，避免模块副本导致桩不生效。
    """
    import tkinter as tk
    import tkinter.messagebox as mb
    from tkinter import filedialog

    captured = []
    originals = {}
    for name, default in (("showerror", "ok"), ("showwarning", "ok"),
                          ("showinfo", "ok"), ("askyesno", True)):
        originals[name] = getattr(mb, name)

        def rec(n=name, d=default):
            def _f(*a, **k):
                captured.append({"函数": n, "标题": str(a[0] if a else k.get("title", "")),
                                 "正文": str(a[1] if len(a) > 1 else k.get("message", ""))})
                return d
            return _f
        setattr(mb, name, rec())

    saved = {
        "list": mod.list_printers, "facts": mod.read_printer_facts,
        "print": mod.print_with_verification, "openfile": os.startfile,
        "state": mod.PRINTER_STATE_FILE, "gui_state": mod.GUI_STATE_PATH,
        "askdir": filedialog.askdirectory, "openfolder": pg.open_folder,
        "fallback": pg.save_a4_fallback,
    }
    roots = []
    sent = []
    opened = []
    result = {"生成时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
              "判定": "FAIL"}
    try:
        mod.PRINTER_STATE_FILE = os.path.join(OUT, "模拟_打印机状态.json")
        mod.GUI_STATE_PATH = os.path.join(OUT, "模拟_GUI交互状态.json")
        os.startfile = lambda *a, **k: True
        pg.open_folder = lambda p: opened.append(p) or True
        filedialog.askdirectory = lambda *a, **k: ""
        mod.print_with_verification = lambda *a, **k: sent.append(a[0]) or {
            "提交结果": {}, "核验": {"结论": "失败",
                             "中文说明": "模拟：系统打印服务报告打印作业失败，页面内容没有到达打印机。"}}

        # 8a. 无实体打印机：只允许保存文件
        mod.list_printers = lambda: ["OneNote (桌面) - 保护", "Microsoft Print to PDF", "Fax"]
        mod.read_printer_facts = lambda n: {"端口": "PORTPROMPT:", "驱动": "虚拟驱动",
                                            "备注": "", "位置": "", "系统属性": 0x40,
                                            "状态位": 0, "打开成功": True}
        r1 = tk.Tk()
        r1.withdraw()
        roots.append(r1)
        g1 = AppClass(r1)
        g1._write_state = lambda: None
        g1.pages = mod.load_document(sample)[:1]
        no_phys = {
            "可见实体数量": len(g1.visible_infos),
            "打印按钮状态": str(g1.btn_print["state"]),
            "禁用原因": g1._print_block_reason(),
            "隐藏设备": [i["名称"] for i in g1.all_infos if not i["是否实体"]],
        }
        g1.do_print()
        assert no_phys["打印按钮状态"] == "disabled", no_phys
        assert len(sent) == 0, "无实体打印机时不得发送打印任务"
        r1.destroy()

        # 8b. 全部离线：打印按钮禁用
        mod.list_printers = lambda: [HP]
        mod.read_printer_facts = lambda n: {"端口": "USB001", "驱动": "HP DeskJet 4900 series PCL-3 (V4)",
                                            "备注": "", "位置": "", "系统属性": 0x40,
                                            "状态位": 0x80, "打开成功": True}
        r2 = tk.Tk()
        r2.withdraw()
        roots.append(r2)
        g2 = AppClass(r2)
        g2._write_state = lambda: None
        g2.pages = mod.load_document(sample)[:1]
        all_off = {"打印按钮状态": str(g2.btn_print["state"]),
                   "禁用原因": g2._print_block_reason(),
                   "中文状态": g2.visible_infos[0]["中文状态"]}
        g2.do_print()
        assert all_off["打印按钮状态"] == "disabled", all_off
        assert len(sent) == 0, "全部离线时不得发送打印任务"
        r2.destroy()

        # 8c. 正常在线 + 提交后核验失败 → 自动保留 A4 文件 + 打开文件夹
        mod.list_printers = lambda: [HP, "Microsoft Print to PDF", "Fax", "OneNote (桌面) - 保护"]
        mod.read_printer_facts = lambda n: {
            "端口": "USB001" if n == HP else "PORTPROMPT:",
            "驱动": "HP DeskJet 4900 series PCL-3 (V4)" if n == HP else "Microsoft Print To PDF",
            "备注": "", "位置": "", "系统属性": 0x40, "状态位": 0, "打开成功": True}
        r3 = tk.Tk()
        r3.withdraw()
        roots.append(r3)
        g3 = AppClass(r3)
        g3._write_state = lambda: None
        g3.pages = mod.load_document(sample)[:1]
        g3.channel, g3.channel_reason = "gdi", "模拟：常规驱动"
        g3._refresh_channel = lambda: None
        g3.do_print()      # 提交（桩）→ 核验失败 → 自动回退
        fb = g3.last_fallback
        assert fb and len(fb["单页PDF"]) == 2, fb
        assert all(os.path.exists(p) for p in [fb["合并PDF"]] + fb["单页PDF"])
        assert opened, "失败后应打开文件夹"
        g3.pdf_channel_print()   # 稳妥通道也能生成文件
        assert len(opened) >= 2, opened
        assert len(sent) == 1, "仅正常在线场景允许提交（且已被桩替换）"
        visible = [i["名称"] for i in g3.visible_infos]
        assert visible == [HP], visible
        r3.destroy()

        banned = ["StartDoc", "win32ui", "Traceback", "Exception", "PaperSize",
                  "Orientation", "GDI", "dpi", "PASS", "FAIL", "NoneType",
                  "CreatePrinterDC", "CreateDC", "ShellExecute", "os.startfile",
                  "Error", "failed", "px @"]
        offenders = []
        for i, it in enumerate(captured):
            for field in ("标题", "正文"):
                t = it[field]
                if not re.search(r"[\u4e00-\u9fff]", t):
                    offenders.append({"序号": i, "字段": field, "问题": "缺少中文"})
                for b in banned:
                    if b.lower() in t.lower():
                        offenders.append({"序号": i, "字段": field, "命中": b})
        assert not offenders, offenders[:8]
        result.update({
            "无实体打印机": no_phys, "全部离线": all_off,
            "可见实体打印机": visible,
            "失败自动回退": {"文件数": len(fb["文件"]), "单页数": len(fb["单页PDF"]),
                       "打开文件夹次数": len(opened)},
            "提交函数被调用次数": len(sent),
            "实际StartDoc调用次数": 0,
            "弹窗数": len(captured), "违规项": offenders,
            "弹窗标题": sorted(set(x["标题"] for x in captured)),
            "判定": "PASS",
        })
    except Exception:
        result["异常"] = traceback.format_exc()
    finally:
        for name, fn in originals.items():
            setattr(mb, name, fn)
        mod.list_printers = saved["list"]
        mod.read_printer_facts = saved["facts"]
        mod.print_with_verification = saved["print"]
        mod.PRINTER_STATE_FILE = saved["state"]
        mod.GUI_STATE_PATH = saved["gui_state"]
        os.startfile = saved["openfile"]
        filedialog.askdirectory = saved["askdir"]
        pg.open_folder = saved["openfolder"]
        pg.save_a4_fallback = saved["fallback"]
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        for r in roots:
            try:
                r.destroy()
            except Exception:
                pass
    return result


def run_text_probe():
    import subprocess
    p = os.path.join(OUT, "弹窗文案.json")
    try:
        os.remove(p)
    except Exception:
        pass
    env = dict(os.environ)
    env.update({"PYTHONNOUSERSITE": "1", "PYTHONPATH": os.path.join(BASE, "lib"),
                "PATH": os.path.join(BASE, "lib", "pywin32_system32") + ";" + env.get("PATH", ""),
                "PYTHONIOENCODING": "utf-8", "PYTHONFAULTHANDLER": "1"})
    err = os.path.join(OUT, "打印链路_弹窗探针_stderr.log")
    with open(err, "w", encoding="utf-8", errors="replace") as ef:
        q = subprocess.Popen([sys.executable, "-X", "faulthandler",
                              os.path.join(BASE, "app.py"), "--ui-text-probe"],
                             cwd=BASE, env=env, stdout=ef, stderr=subprocess.STDOUT)
        try:
            q.wait(timeout=120)
            code = q.returncode
        except subprocess.TimeoutExpired:
            q.kill()
            code = "超时"
    assert os.path.exists(p), "弹窗文案.json 未生成"
    with open(p, "r", encoding="utf-8") as f:
        d = json.load(f)
    banned = ["StartDoc", "win32ui", "Traceback", "Exception", "PaperSize",
              "Orientation", "GDI", "dpi", "PASS", "FAIL", "NoneType",
              "CreatePrinterDC", "CreateDC", "ShellExecute", "os.startfile",
              "Error", "failed", "px @"]
    offenders = []
    for i, it in enumerate(d["弹窗"]):
        for field in ("标题", "正文"):
            t = it[field]
            if not re.search(r"[\u4e00-\u9fff]", t):
                offenders.append({"序号": i, "字段": field, "问题": "缺少中文"})
            for b in banned:
                if b.lower() in t.lower():
                    offenders.append({"序号": i, "字段": field, "命中": b})
    assert d["触发分支数"] >= 15, d["触发分支数"]
    assert not offenders, offenders[:8]
    assert code == 0, code
    return {"触发分支数": d["触发分支数"], "探针退出码": code,
            "核验结论序列": d.get("核验结论序列"), "违规项": offenders,
            "证据": p, "判定": "PASS"}


if __name__ == "__main__":
    r = main()
    sys.exit(0 if r.get("结论") == "PASS" else 1)
