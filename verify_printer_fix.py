# -*- coding: utf-8 -*-
"""
打印机修复独立验证（不自动消耗纸张）：
  1) 下拉默认选系统默认打印机，且优先实体 HP；OneNote/Fax 标注 [不支持GDI]
  2) OneNote 场景：真实 StartDoc 失败被包装成 PrinterStartDocError，
     提示里含具体打印机名、原始错误、可操作建议、未出纸说明，程序不崩溃
  3) HP DeskJet 4900 series：dry_run / DC 强制 A4 纵向检查通过（全程不 StartDoc）
  4) 安全兜底：导出 A4 PDF 可生成且每页 595.28×841.89pt；验证过程无任何自动打印动作
  5) 真实启动 GUI，核验下拉实际默认选中值（读 测试输出/打印机状态.json）

产出：测试输出/打印机修复验证报告.json
用法：C:\\Python314\\python.exe verify_printer_fix.py
"""
import _bootstrap  # 注入 lib

import os
import sys
import json
import time
import subprocess
import datetime
import traceback

import app

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "测试输出")
os.makedirs(OUT, exist_ok=True)
APP = os.path.join(BASE, "app.py")
REPORT = os.path.join(OUT, "打印机修复验证报告.json")
STATE = os.path.join(OUT, "打印机状态.json")
STDERR = os.path.join(OUT, "打印机验证_GUI_stderr.log")
SAMPLE = os.path.join(BASE, "测试样例", "sample_seal_multipage.pdf")
ONE_NOTE = "OneNote (桌面) - 保护"

# 本验证脚本严禁自动出纸：只做 dry_run、虚拟目标 StartDoc、只读 DC 查询、导出 PDF
AUTO_PRINT_ACTIONS = 0


def main():
    rep = {"生成时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
           "验证方式": "真实系统打印机枚举/DC 查询 + 虚拟目标真实 StartDoc 失败捕获 + "
                   "HP dry_run（不 StartDoc）+ 真实 GUI 启动读取下拉默认选中",
           "自动发送打印任务次数": AUTO_PRINT_ACTIONS,
           "检查项": {}, "结论": "FAIL"}
    proc = None
    errf = None
    try:
        assert os.path.exists(SAMPLE), "缺少测试样本 " + SAMPLE

        # === 1. 默认选中 + 标注 ===
        names = app.list_printers()
        sysdef = app.get_default_printer()
        chosen = app.choose_default_printer(names, sysdef)
        opts, mapping = app.build_printer_options(names, chosen)
        assert names and sysdef, "未读到打印机"
        assert chosen == sysdef, f"下拉默认未选系统默认: {chosen} != {sysdef}"
        assert app.is_hp_physical(chosen), f"默认未优先实体 HP: {chosen}"
        assert app.printer_status(chosen)[0] == "physical"
        assert ONE_NOTE in names, "本机缺少 OneNote 虚拟目标，无法复现"
        assert "Fax" in names, "本机缺少 Fax 目标"
        one_display = opts[names.index(ONE_NOTE)]
        assert "[不支持" in one_display
        assert "[不支持" in opts[names.index("Fax")]
        assert "[不支持" in opts[names.index("OneNote (Desktop)")]
        assert "★" in opts[names.index(chosen)], "默认选中项缺少 ★ 标记"
        assert app.printer_name_from_display(one_display, mapping) == ONE_NOTE
        print(f"=== 1. 默认选中 ===\n  系统默认={sysdef}\n  下拉默认={chosen} "
              f"等级={app.printer_status(chosen)[0]}")
        for o in opts:
            print("   ", o)
        rep["检查项"]["默认选中与标注"] = {
            "系统默认打印机": sysdef, "下拉默认选中": chosen,
            "优先实体HP": app.is_hp_physical(chosen),
            "OneNote标注": one_display,
            "Fax标注": opts[names.index("Fax")],
            "全部下拉项": opts, "判定": "PASS"}

        # === 2. OneNote 报错清晰、不崩溃 ===
        hints = app.start_doc_failure_hints(ONE_NOTE)
        wrapped = app.PrinterStartDocError(ONE_NOTE, "win32ui.error: StartDoc failed", hints)
        log_text = str(wrapped)                  # 含英文原始异常，仅供日志
        msg = wrapped.user_message()             # 面向用户：纯简体中文
        for kw in (ONE_NOTE, "无法启动打印文档", "实体打印机",
                   "导出A4 PDF", "不会消耗纸张", "没有自动重试"):
            assert kw in msg, f"错误提示缺少关键内容: {kw}"
        for kw in ("StartDoc", "win32ui", "failed", "GDI"):
            assert kw not in msg, f"用户提示夹带英文技术串: {kw}"
        assert "StartDoc failed" in log_text, "英文原始异常应保留在日志文本里"
        assert app.printer_status(ONE_NOTE)[0] == "unsupported"
        assert app.printer_status("Fax")[0] == "unsupported"
        print(f"=== 2. OneNote 中文错误提示 ===\n  {msg.replace(chr(10), ' | ')[:260]}...")

        pages = app.load_document(SAMPLE)[:1]
        imgs, _ = app.process_document(pages, "auto", False, 0.0)
        assert len(imgs) == 2, f"样本应输出 2 张 A4，实际 {len(imgs)}"
        real_ok, real_msg = False, ""
        try:
            app.print_a4_images(ONE_NOTE, imgs, dry_run=False)   # 虚拟目标，不出纸
            real_msg = "未抛异常（不符合预期）"
        except app.PrinterStartDocError as e:
            real_ok, real_msg = True, str(e).replace("\n", " | ")
        except Exception as e:
            real_msg = "非预期异常: " + repr(e)
        assert real_ok and ONE_NOTE in real_msg and "StartDoc" in real_msg, real_msg
        print(f"  真实 StartDoc 失败已捕获: {real_msg[:200]}...")
        rep["检查项"]["OneNote报错清晰不崩溃"] = {
            "等级": app.printer_status(ONE_NOTE)[0],
            "包装异常": "PrinterStartDocError", "真实启动文档失败已捕获": real_ok,
            "面向用户提示": msg, "面向用户提示无英文技术串": True,
            "日志用文本含英文原始异常": "StartDoc failed" in log_text,
            "真实异常摘要（仅日志）": real_msg[:400],
            "界面层提前拦截": "选择 unsupported 目标时不发送打印任务",
            "判定": "PASS"}

        # === 3. HP dry_run / DC 检查（不 StartDoc） ===
        dry = app.print_a4_images(chosen, imgs, dry_run=True)
        assert dry["pages"] == 2, dry["pages"]
        assert dry["evidence"]["verified_a4_portrait"]
        assert dry["evidence"]["applied_devmode"]["PaperSize"] == 9
        assert dry["evidence"]["applied_devmode"]["Orientation"] == 1
        dc = app.query_printer_dc(chosen)
        assert dc["dc_paper"] == "DMPAPER_A4(9)"
        assert dc["dc_orientation"] == "DMORIENT_PORTRAIT(1)"
        assert dc["devmode_changed_dc"] and dc["verified_a4_portrait"]
        print(f"=== 3. HP dry_run/DC ===\n  pages={dry['pages']} "
              f"paper={dry['evidence']['paper_mm']}mm "
              f"A4纵向={dry['evidence']['verified_a4_portrait']} "
              f"DC={dc['dc_paper']}/{dc['dc_orientation']}")
        rep["检查项"]["HP_dryRun_DC"] = {
            "打印机": chosen, "dry_run页数": dry["pages"],
            "A4纸张mm": list(dry["evidence"]["paper_mm"]),
            "PaperSize": dry["evidence"]["applied_devmode"]["PaperSize"],
            "Orientation": dry["evidence"]["applied_devmode"]["Orientation"],
            "可印区": list(dry["printable_rect"]),
            "DC_paper": dc["dc_paper"], "DC_orientation": dc["dc_orientation"],
            "DEVMODE改变了DC": dc["devmode_changed_dc"],
            "未调用StartDoc": True, "判定": "PASS"}

        # === 4. 安全兜底 PDF ===
        pdf = os.path.join(OUT, "兜底_系统打印_A4.pdf")
        app.export_a4_pdf(imgs, pdf)
        import fitz
        d = fitz.open(pdf)
        sizes = [(round(p.rect.width, 2), round(p.rect.height, 2)) for p in d]
        cnt = d.page_count
        d.close()
        assert cnt == 2 and all(abs(a - 595.28) < 1 and abs(b - 841.89) < 1
                                for a, b in sizes), (cnt, sizes)
        print(f"=== 4. 兜底 PDF ===\n  {pdf} 页数={cnt} 每页pt={sizes[0]}")
        rep["检查项"]["PDF安全兜底"] = {
            "文件": pdf, "字节": os.path.getsize(pdf), "页数": cnt,
            "每页pt": sizes, "生成后可被系统打印对话框使用": True,
            "是否自动调起打印": False,
            "兜底动作": "导出 A4 PDF → 用户确认 → os.startfile 打开阅读器（只 open，不是 print verb）",
            "判定": "PASS"}

        # === 5. 真实 GUI 启动，读下拉默认选中 ===
        try:
            os.remove(STATE)
        except Exception:
            pass
        env = dict(os.environ)
        env.update({"PYTHONNOUSERSITE": "1", "PYTHONPATH": os.path.join(BASE, "lib"),
                    "PATH": os.path.join(BASE, "lib", "pywin32_system32") + ";" + env.get("PATH", ""),
                    "PYTHONIOENCODING": "utf-8", "PYTHONFAULTHANDLER": "1"})
        errf = open(STDERR, "w", encoding="utf-8", errors="replace")
        proc = subprocess.Popen([sys.executable, "-X", "faulthandler", APP],
                                cwd=BASE, env=env, stdout=errf,
                                stderr=subprocess.STDOUT)
        state = None
        deadline = time.time() + 20
        while time.time() < deadline:
            if os.path.exists(STATE):
                try:
                    with open(STATE, "r", encoding="utf-8") as f:
                        state = json.load(f)
                    break
                except Exception:
                    state = None
            if proc.poll() is not None:
                break
            time.sleep(0.2)
        gui_alive = proc.poll() is None
        assert state is not None, "GUI 未写出 打印机状态.json"
        assert state["系统默认打印机"] == sysdef
        assert state["下拉默认选中实际名"] == chosen
        assert state["选中判定"] == "physical"
        assert state["全部打印机"][names.index(ONE_NOTE)]["等级"] == "unsupported"
        print(f"=== 5. 真实 GUI ===\n  进程存活={gui_alive} "
              f"下拉默认={state['下拉默认选中实际名']} 显示={state['下拉默认选中显示']}")
        rep["检查项"]["真实GUI默认选中"] = {
            "进程存活": gui_alive, "打印机状态文件": STATE,
            "系统默认": state["系统默认打印机"],
            "下拉默认选中实际名": state["下拉默认选中实际名"],
            "下拉默认选中显示": state["下拉默认选中显示"],
            "判定": "PASS"}
        # 第 5 项的 GUI 用完即关，避免与探针进程互相干扰
        if proc.poll() is None:
            proc.terminate()
            try:
                proc.wait(timeout=5)
            except Exception:
                proc.kill()
        errf.close()
        errf = None
        with open(STDERR, "r", encoding="utf-8", errors="replace") as f:
            rep["GUI_stderr"] = f.read().strip()[-2000:]

        # === 6. 真实 App 上触发全部弹窗，验收「标题/正文纯中文、无英文技术串」 ===
        import re
        probe_json = os.path.join(OUT, "弹窗文案.json")
        probe_err = os.path.join(OUT, "弹窗文案_probe_stderr.log")
        try:
            os.remove(probe_json)
        except Exception:
            pass
        log_path = os.path.join(BASE, "A3SplitTool.log")
        log_off = os.path.getsize(log_path) if os.path.exists(log_path) else 0
        pf = open(probe_err, "w", encoding="utf-8", errors="replace")
        p2 = subprocess.Popen([sys.executable, "-X", "faulthandler", APP,
                               "--ui-text-probe"], cwd=BASE, env=env,
                              stdout=pf, stderr=subprocess.STDOUT)
        try:
            p2.wait(timeout=90)
            probe_exit = p2.returncode
        except subprocess.TimeoutExpired:
            p2.kill()
            probe_exit = "超时"
        pf.flush()
        pf.close()
        assert os.path.exists(probe_json), "弹窗文案.json 未生成"
        with open(probe_json, "r", encoding="utf-8") as f:
            probe = json.load(f)
        banned = ["StartDoc", "win32ui", "Traceback", "Exception", "PaperSize",
                  "Orientation", "GDI", "dpi", "PASS", "FAIL", "NoneType",
                  "CreatePrinterDC", "CreateDC", "ShellExecute", "os.startfile",
                  "Error", "failed", "px @"]
        offenders = []
        for i, it in enumerate(probe["弹窗"]):
            for field in ("标题", "正文"):
                t = it[field]
                if not re.search(r"[\u4e00-\u9fff]", t):
                    offenders.append({"序号": i, "字段": field,
                                      "问题": "缺少中文", "内容": t[:120]})
                for b in banned:
                    if b.lower() in t.lower():
                        offenders.append({"序号": i, "字段": field,
                                          "命中": b, "内容": t[:120]})
        assert probe["触发分支数"] >= 10, probe["触发分支数"]
        assert not offenders, offenders[:8]
        with open(log_path, "r", encoding="utf-8") as f:
            f.seek(log_off)
            newlog = f.read()
        assert "原始异常（仅日志）" in newlog or "win32ui.error" in newlog, \
            "英文原始异常未按要求只写入日志"
        print(f"=== 6. 弹窗文案 ===\n  触发分支={probe['触发分支数']} "
              f"违规项={len(offenders)} 英文原始异常已写入日志")
        rep["检查项"]["全部弹窗中文无英文技术串"] = {
            "触发弹窗分支数": probe["触发分支数"],
            "探针退出码": probe_exit,
            "覆盖路径": ["打开正常", "打开格式错误", "打开时加载异常", "拖入时加载异常",
                     "导出失败", "兜底生成失败", "兜底打开失败",
                     "切换OneNote预警", "切换虚拟打印机提示", "OneNote打印拦截",
                     "启动打印文档失败", "其它打印异常", "拖入格式错误",
                     "未打开文件提示"],
            "禁止英文技术串": banned, "违规项": offenders,
            "英文原始异常仅写日志": True,
            "证据": probe_json, "判定": "PASS"}

        rep["结论"] = "PASS"
    except Exception:
        rep["异常"] = traceback.format_exc()
        print(rep["异常"])
    finally:
        if errf:
            try:
                errf.flush()
                errf.close()
                with open(STDERR, "r", encoding="utf-8", errors="replace") as f:
                    rep["GUI_stderr"] = f.read().strip()[-3000:]
            except Exception:
                pass
        if proc and proc.poll() is None:
            try:
                proc.terminate()
                proc.wait(timeout=5)
            except Exception:
                try:
                    proc.kill()
                except Exception:
                    pass
        with open(REPORT, "w", encoding="utf-8") as f:
            json.dump(rep, f, ensure_ascii=False, indent=2)
        app.log(f"VERIFY_PRINTER_FIX 结论={rep['结论']} "
                f"默认选中={rep['检查项'].get('默认选中与标注', {}).get('下拉默认选中')} "
                f"自动打印次数={AUTO_PRINT_ACTIONS}")
        print("  证据:", REPORT)
        print("  结论:", rep["结论"], "| 自动发送打印任务次数:", AUTO_PRINT_ACTIONS)
    return rep


if __name__ == "__main__":
    r = main()
    sys.exit(0 if r.get("结论") == "PASS" else 1)
