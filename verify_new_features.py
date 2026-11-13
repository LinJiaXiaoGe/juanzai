# -*- coding: utf-8 -*-
"""
三项新功能独立验证（只读本地文件，不联网）：
  1) 两页使用不同分割位置，导出/打印均使用该位置
  2) 拖动裁切：画布坐标 -> 每页分割位置，clamp 生效，重算结果变化
  3) A4 预览：两张之间有间距、各自边框、页号，不再连成一张
  4) 拖放加载：真实 Win32 DragAcceptFiles/WM_DROPFILES 消息链 + 队列轮询 + 格式错误提示
  5) 旧功能回归：split=0.5 且关闭去密封线时逐像素一致、A4 强制打印 DC 不变

产出：测试输出/新功能验证报告.json、新功能_A4预览_间距边框页号.png、
      新功能_两页不同分割_A4.pdf、新功能_拖放格式样本/
"""
import _bootstrap

import os
import io
import json
import time
import ctypes
from ctypes import wintypes, Structure, byref, c_void_p
import datetime

from PIL import Image, ImageChops

import app
import dragdrop
from verify_seal import legacy_make_a4_page

SPLIT_A, SPLIT_B = 0.35, 0.65


# ---------- Win32 WM_DROPFILES 消息链：构造真实 HDROP 并投递到真实窗口 ----------
WM_DROPFILES = 0x0233


class DROPFILES(Structure):
    _fields_ = [("pFiles", wintypes.DWORD), ("x", ctypes.c_long), ("y", ctypes.c_long),
                ("fNC", wintypes.BOOL), ("fWide", wintypes.BOOL)]


_kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
_kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
_kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
_kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
_kernel32.GlobalLock.restype = c_void_p
_kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
_kernel32.GlobalUnlock.restype = wintypes.BOOL

_user32 = ctypes.WinDLL("user32", use_last_error=True)
_user32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT,
                                 ctypes.c_size_t, ctypes.c_ssize_t]
_user32.PostMessageW.restype = wintypes.BOOL


def _make_hdrop(path):
    payload = ctypes.create_string_buffer(path.encode("utf-16-le") + b"\x00\x00" + b"\x00\x00")
    size = ctypes.sizeof(DROPFILES) + len(payload.raw)
    h = _kernel32.GlobalAlloc(0x0002 | 0x0040, size)   # GMEM_MOVEABLE | GMEM_ZEROINIT
    ptr = _kernel32.GlobalLock(h)
    df = DROPFILES(ctypes.sizeof(DROPFILES), 0, 0, 0, 1)   # fWide=1 表示 Unicode 路径
    ctypes.memmove(ptr, byref(df), ctypes.sizeof(df))
    ctypes.memmove(ptr + ctypes.sizeof(df), payload.raw, len(payload.raw))
    _kernel32.GlobalUnlock(h)
    return h


def _post_drop(hwnd, path):
    """把真实 HDROP 作为 WM_DROPFILES 投递到真实 HWND（走完整 WndProc + 队列 + 轮询链）。"""
    h = _make_hdrop(path)
    ok = _user32.PostMessageW(wintypes.HWND(hwnd), WM_DROPFILES,
                             ctypes.c_size_t(h), 0)
    return bool(ok)


def _pump(root, seconds=1.0):
    end = time.time() + seconds
    while time.time() < end:
        root.update()
        time.sleep(0.025)


# ---------- 主验证 ----------
def run(base, skip_gui=False):
    sample_dir = os.path.join(base, "测试样例")
    out_dir = os.path.join(base, "测试输出")
    fmt_dir = os.path.join(out_dir, "新功能_拖放格式样本")
    os.makedirs(fmt_dir, exist_ok=True)

    rep = {"生成时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
           "检查项": {}, "结论": "FAIL"}
    pdf = os.path.join(sample_dir, "sample_seal_multipage.pdf")
    assert os.path.exists(pdf), f"缺少测试样本: {pdf}"

    # === 1. 两页不同分割位置 ===
    print("=== 1. 两页不同分割位置 ===")
    pages = app.load_document(pdf)[:2]
    assert len(pages) == 2
    app.set_page_split(pages, 0, SPLIT_A)
    app.set_page_split(pages, 1, SPLIT_B)
    imgs, results = app.process_document(pages, "auto", False, 0.0)
    assert len(imgs) == 4, len(imgs)
    assert abs(results[0]["split"] - SPLIT_A) < 1e-6
    assert abs(results[1]["split"] - SPLIT_B) < 1e-6
    assert all(im.size == (app.A4_W_PX, app.A4_H_PX) for im in imgs)

    w0 = pages[0]["img"].width
    w1 = pages[1]["img"].width
    h1a, h2a = app.split_image(pages[0]["img"], "landscape", SPLIT_A)
    h1b, h2b = app.split_image(pages[1]["img"], "landscape", SPLIT_B)
    ratio_a, ratio_b = h1a.width / w0, h1b.width / w1
    assert abs(ratio_a - SPLIT_A) < 0.005 and abs(ratio_b - SPLIT_B) < 0.005
    assert h1a.width + h2a.width == w0 and h1b.width + h2b.width == w1, "中缝有缺口/重叠"
    assert ImageChops.difference(imgs[0], imgs[2]).getbbox() is not None, "两页分割位置未产生差异"

    out_pdf = os.path.join(out_dir, "新功能_两页不同分割_A4.pdf")
    app.export_a4_pdf(imgs, out_pdf)
    import fitz
    d = fitz.open(out_pdf)
    sizes = [(round(p.rect.width, 2), round(p.rect.height, 2)) for p in d]
    assert d.page_count == 4 and all(abs(a - 595.28) < 1 and abs(b - 841.89) < 1
                                     for a, b in sizes)
    d.close()

    printers = app.list_printers()
    target = next((p for p in ("Microsoft Print to PDF", "HP DeskJet 4900 series",
                               "HP Smart Printing") if p in printers), None)
    assert target, "无可验证打印机"
    dry = app.print_a4_images(target, imgs, dry_run=True)
    assert dry["pages"] == 4 and dry["evidence"]["verified_a4_portrait"], "打印未使用这批 A4"
    print(f"  第1页 split={ratio_a:.3f} / 第2页 split={ratio_b:.3f}; 导出4页A4; "
          f"打印dry_run={dry['pages']}页 A4纵向={dry['evidence']['verified_a4_portrait']}")
    rep["检查项"]["两页不同分割位置"] = {
        "第1页": round(ratio_a, 4), "第2页": round(ratio_b, 4),
        "输出A4页数": len(imgs), "导出PDF": out_pdf, "每页pt": sizes[0],
        "打印dry_run页数": dry["pages"], "A4纵向": dry["evidence"]["verified_a4_portrait"],
        "判定": "PASS",
    }

    # === 2. 拖动裁切 ===
    print("=== 2. 拖动裁切（画布坐标 -> 每页分割位置） ===")
    img = pages[0]["img"]
    iw, ih = img.size
    nw, nh = 600, max(1, round(600 * ih / iw))
    disp = (20, 28, nw / iw, nw, nh)
    cx = disp[0] + nw * 0.35
    r_drag = app.split_from_canvas_event(cx, disp[1] + nh * 0.5, "landscape", disp, iw, ih)
    assert abs(r_drag - SPLIT_A) < 0.01, r_drag
    r_low = app.split_from_canvas_event(disp[0] - 10000, 0, "landscape", disp, iw, ih)
    r_high = app.split_from_canvas_event(disp[0] + 10000, 0, "landscape", disp, iw, ih)
    assert r_low == app.SPLIT_MIN and r_high == app.SPLIT_MAX, (r_low, r_high)
    r_portrait = app.split_from_canvas_event(0, disp[1] + nh * 0.65, "portrait",
                                             disp, iw, ih)
    assert abs(r_portrait - 0.65) < 0.01, r_portrait

    app.set_page_split(pages, 0, 0.5)
    before = app.process_document(pages, "auto", False, 0.0)[0][0]
    saved = app.set_page_split(pages, 0, r_drag)
    after = app.process_document(pages, "auto", False, 0.0)[0][0]
    assert ImageChops.difference(before, after).getbbox() is not None, "拖动后裁切结果未变化"
    assert abs(pages[0]["split"] - saved) < 1e-9, "未按页保存"
    assert abs(pages[1]["split"] - SPLIT_B) < 1e-9, "第1页拖动不应影响第2页"
    print(f"  拖到35% -> split={r_drag:.4f}; clamp=[{r_low}, {r_high}]; "
          f"纵向映射={r_portrait:.4f}; 页1已保存={saved:.4f}, 页2仍={pages[1]['split']:.4f}")
    rep["检查项"]["拖动裁切"] = {"画布映射": round(r_drag, 4), "下限": r_low, "上限": r_high,
                              "纵向映射": round(r_portrait, 4), "页1保存": round(saved, 4),
                              "页2保持": round(pages[1]["split"], 4), "判定": "PASS"}

    # === 3. A4 预览间距 / 边框 / 页号 ===
    print("=== 3. A4 预览间距 / 边框 / 页号 ===")
    a4a, a4b = imgs[0], imgs[1]
    combo, layout = app.compose_a4_preview(a4a, a4b, 1, 0.0)
    (x0a, y0a, x1a, y1a), (x0b, y0b, x1b, y1b) = layout["boxes"]
    assert layout["gap_between_pages"] == app.PREVIEW_GAP > 0
    assert x0b - x1a == app.PREVIEW_GAP
    assert combo.size == (2 * app.A4_W_PX + 3 * app.PREVIEW_GAP,
                          app.A4_H_PX + 2 * app.PREVIEW_GAP + app.PREVIEW_HEADER)
    # 间隙列必须是深色背景，不能是白色/A4 内容
    mid = (x1a + x0b) // 2
    bg_ok = all(sum(abs(combo.getpixel((mid, yy))[i] - app.PREVIEW_BG[i]) for i in range(3)) < 12
                for yy in range(y0a, y1a, 120))
    assert bg_ok, "两页预览仍连在一起"
    # 边框像素
    border_ok = combo.getpixel((x0a - 1, y0a + 100)) == (215, 215, 215)
    assert border_ok, "缺少边框"
    # 页号：页号区存在亮色文字像素
    head = combo.crop((x0a, 0, x0a + 500, app.PREVIEW_GAP + app.PREVIEW_HEADER))
    bright = any(max(px) > 190 for px in head.getdata())
    assert bright, "缺少页号"
    shot = os.path.join(out_dir, "新功能_A4预览_间距边框页号.png")
    combo.resize((combo.width // 2, combo.height // 2), Image.LANCZOS).save(shot)
    # 去密封线开启时布局不变
    combo_s, layout_s = app.compose_a4_preview(a4a, a4b, 1, 8.0)
    assert combo_s.size == combo.size and layout_s["boxes"] == layout["boxes"]
    print(f"  间距={layout['gap_between_pages']}px 边框={border_ok} 页号={bright} "
          f"画布={combo.size} 去密封线后布局不变=PASS")
    rep["检查项"]["A4预览间距边框页号"] = {
        "间距px": layout["gap_between_pages"], "边框": border_ok, "页号": bright,
        "画布尺寸": combo.size, "两页box": layout["boxes"],
        "截图": shot, "去密封线开启后布局不变": "PASS", "判定": "PASS",
    }

    # === 4. 拖放加载 + 格式错误提示 ===
    print("=== 4. 拖放加载 + 格式错误提示 ===")
    src_jpg = os.path.join(sample_dir, "sample_a3_portrait.jpg")
    probe = Image.new("RGB", (600, 420), (250, 250, 250))
    fmt_paths = [pdf, src_jpg]
    for ext, fmt in (("png", "PNG"), ("bmp", "BMP"), ("tiff", "TIFF")):
        p = os.path.join(fmt_dir, f"拖放样本.{ext}")
        probe.save(p, format=fmt)
        fmt_paths.append(p)
    load_ok = {}
    for p in fmt_paths:
        loaded, err = app.load_dropped_document([p])
        assert err is None and loaded, f"{os.path.basename(p)} 加载失败: {err}"
        load_ok[os.path.splitext(p)[1].lower()] = len(loaded)

    bad_cases = [
        ([r"C:\fake\试卷.docx"], "不支持的文件格式"),
        ([r"C:\fake\病毒.exe"], "不支持的文件格式"),
        ([r"C:\fake\说明.txt"], "不支持的文件格式"),
        ([pdf, src_jpg], "一次只能拖入 1 个文件"),
        ([], "未检测到拖入的文件"),
        ([r"C:\fake\不存在的.pdf"], "文件不存在"),
    ]
    err_msgs = {}
    for i, (paths, keyword) in enumerate(bad_cases, 1):
        _, err = app.load_dropped_document(paths)
        assert err and keyword in err, f"未提示预期错误: {paths} -> {err}"
        err_msgs[f"{i:02d}_{keyword}"] = err.splitlines()[0]

    gui_part = {"DragAcceptFiles注册": "SKIP(selftest)",
                "WM_DROPFILES消息链": "SKIP(selftest)"}
    if not skip_gui:
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        root.update()
        received = []

        def on_files(paths):
            loaded, err = app.load_dropped_document(paths)
            received.append({"paths": paths,
                             "pages": None if loaded is None else len(loaded),
                             "error": err})

        reg_ok = dragdrop.hook_dropfiles(root, on_files)
        assert reg_ok, f"DragAcceptFiles 注册失败: {dragdrop.status()}"
        status = dragdrop.status()
        assert status["registered"] and status["mechanism"].startswith("WM_DROPFILES")

        # 真实 Win32 消息：拖入合法 PDF
        assert _post_drop(status["hwnd"], pdf), "PostMessage(WM_DROPFILES) 失败"
        _pump(root, 0.9)
        # 真实 Win32 消息：拖入非法 .docx -> 应产生格式错误提示
        bad = r"C:\fake\试卷.docx"
        assert _post_drop(status["hwnd"], bad), "PostMessage(WM_DROPFILES) 失败"
        _pump(root, 0.9)
        assert len(received) >= 2, received
        assert received[0]["pages"] == 3 and received[0]["error"] is None, received[0]
        assert received[1]["pages"] is None and "不支持的文件格式" in received[1]["error"], received[1]
        assert dragdrop.status()["last_paths"] == [bad]
        drops = dragdrop.status()["drops"]
        assert drops >= 2, drops
        dragdrop.unhook_dropfiles()
        root.destroy()
        gui_part = {
            "DragAcceptFiles注册": "PASS",
            "机制": status["mechanism"], "HWND": status["hwnd"],
            "WM_DROPFILES消息链": "PASS", "收到消息数": drops,
            "拖入PDF": received[0], "拖入docx错误提示": received[1]["error"],
        }
        print(f"  DragAcceptFiles/WM_DROPFILES 消息数={drops}；"
              f"PDF->{received[0]['pages']}页；docx->格式错误提示")
    else:
        print("  (selftest 模式：跳过 Tk 注册，完整链路由独立运行本脚本验证)")
    print(f"  支持格式加载={load_ok}；错误提示 {len(err_msgs)} 种")
    rep["检查项"]["拖放加载"] = {
        "支持格式": load_ok, "错误提示": err_msgs, "GUI链路": gui_part, "判定": "PASS",
    }

    # === 5. 旧功能回归 ===
    print("=== 5. 旧功能回归 ===")
    pg0 = app.load_document(pdf)[0]
    raw1, raw2 = app.split_image(pg0["img"], "landscape")     # 旧签名：无 split 参数
    res0 = app.process_page(pg0["img"], pg0["orientation"], "auto", 0.0, 1)
    d1 = ImageChops.difference(legacy_make_a4_page(raw1, "landscape"),
                               res0["halves"][0][1]).getbbox()
    d2 = ImageChops.difference(legacy_make_a4_page(raw2, "landscape"),
                               res0["halves"][1][1]).getbbox()
    assert d1 is None and d2 is None, "默认 0.5 分割时与原算法不一致"
    # 旧版 annotate_seal_preview 调用签名仍可用
    old_combo = Image.new("RGB", (app.A4_W_PX * 2, app.A4_H_PX), (40, 40, 40))
    old_combo.paste(raw1, (0, 0))
    old_shot = app.annotate_seal_preview(old_combo, 8.0, 1, app.A4_W_PX, app.A4_H_PX)
    assert old_shot.size == old_combo.size
    # A4 强制打印 DC 回归
    ev = app.query_printer_dc(target)
    assert ev["verified_a4_portrait"] and ev["devmode_changed_dc"]
    print(f"  默认split=0.5 逐像素回归=PASS；旧预览签名兼容=PASS；"
          f"打印DC A4纵向={ev['dc_paper']}/{ev['dc_orientation']}")
    rep["检查项"]["旧功能回归"] = {
        "默认分割逐像素一致": "PASS", "旧预览签名兼容": "PASS",
        "打印DC": f"{ev['dc_paper']}/{ev['dc_orientation']}", "判定": "PASS",
    }

    rep["结论"] = "PASS"
    # selftest 的无 GUI 版本写独立文件，不覆盖含真实拖放链路证据的完整报告
    report_name = "新功能验证报告_自测.json" if skip_gui else "新功能验证报告.json"
    path = os.path.join(out_dir, report_name)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(rep, f, ensure_ascii=False, indent=2)
    app.log("VERIFY_NEW_FEATURES PASS " + json.dumps(
        {k: v.get("判定") if isinstance(v, dict) else v for k, v in rep["检查项"].items()},
        ensure_ascii=False))
    print("  证据文件:", path)
    print("  结论: PASS")
    return rep


if __name__ == "__main__":
    run(os.path.dirname(os.path.abspath(__file__)))
