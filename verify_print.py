# -*- coding: utf-8 -*-
"""
独立验收脚本：验证"打印 DC 强制 A4 纵向"是否真实生效。
不需要打开 GUI，直接输出可核验证据并写入：
  - 控制台
  - 测试输出/验证报告_打印DC.json
  - A3SplitTool.log

运行（纯净环境示例）：
  set PYTHONNOUSERSITE=1
  C:\\Python314\\python.exe verify_print.py
"""
import _bootstrap  # 必须在第三方库之前

import os
import json
import datetime
import app


def main():
    base = os.path.dirname(os.path.abspath(__file__))
    out_dir = os.path.join(base, "测试输出")
    os.makedirs(out_dir, exist_ok=True)

    print("=" * 70)
    print("打印 DC 强制 A4 纵向 — 独立验收")
    print("=" * 70)
    rep = _bootstrap.dependency_report()
    print("[运行环境]")
    print("  python     =", rep["python"])
    print("  PIL        =", rep["PIL"])
    print("  fitz       =", rep["fitz"])
    print("  win32print =", rep["win32print"])
    print("  win32ui    =", rep["win32ui"])
    print("  全部依赖来自工具目录 lib =", rep["uses_tool_lib"])
    assert rep["uses_tool_lib"], "依赖仍来自工具目录外部"

    printers = app.list_printers()
    print("\n[本机打印机]", printers)
    targets = [p for p in ("Microsoft Print to PDF", "HP DeskJet 4900 series", "HP Smart Printing")
               if p in printers]
    assert targets, "未找到可验证的打印机"

    evs = []
    for t in targets:
        ev = app.query_printer_dc(t)
        evs.append(ev)
        print("\n" + "-" * 70)
        print("打印机:", t)
        print("  DEVMODEW 结构大小      =", ev["DEVMODEW_size"], "(Windows 标准 220)")
        print("  ResetDCW 调用成功      =", ev["resetdcw_ok"])
        print("  应用的 DEVMODE         = PaperSize=9 (DMPAPER_A4), "
              "Orientation=1 (DMORIENT_PORTRAIT), dmFields=3")
        print("  ① 默认 DC        纸张 =", ev["default"]["paper_mm"], "mm",
              " caps =", ev["default"]["caps"])
        print("  ② 强制 Letter/横向    =", ev["letter_landscape"]["paper_mm"], "mm")
        print("  ③ 强制 A4/纵向        =", ev["a4_portrait"]["paper_mm"], "mm",
              " caps =", ev["a4_portrait"]["caps"])
        print("  DEVMODE 确实改变了 DC =", ev["devmode_changed_dc"])
        print("  DC 实际 PaperSize     =", ev["dc_paper"])
        print("  DC 实际 Orientation   =", ev["dc_orientation"])
        print("  A4 纵向校验           =", ev["verified_a4_portrait"])
        app.log("VERIFY_PRINT " + json.dumps(ev, ensure_ascii=False))

    ok = all(e["resetdcw_ok"] and e["devmode_changed_dc"] and e["verified_a4_portrait"]
             for e in evs)
    report = {
        "生成时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "说明": "ResetDCW 对照实验：把同一 DC 先设为 Letter 横向，再设为 A4 纵向，"
                "查询 GetDeviceCaps 物理尺寸变化，证明 DEVMODE 真正作用到打印 DC。",
        "依赖来源": rep,
        "本机打印机": printers,
        "DC证据": evs,
        "结论": "PASS" if ok else "FAIL",
        "总结论": "PASS — 所有被测打印机 DC 均被强制为 A4 纵向" if ok
                 else "FAIL — 存在未通过项",
        "判定布尔": ok,
    }
    path = os.path.join(out_dir, "验证报告_打印DC.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)

    print("\n" + "=" * 70)
    print("证据文件:", path)
    print("总结论:", report["总结论"])
    app.log("VERIFY_PRINT 总结论=" + report["总结论"])
    assert ok, "打印 DC 强制 A4 纵向验证未通过"
    return True


if __name__ == "__main__":
    main()
