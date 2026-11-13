# -*- coding: utf-8 -*-
"""
A3 试卷拆分打印工具
功能：
  1) 导入 PDF / JPG / PNG / BMP / TIFF（多页文件逐页处理），A3 区可直接拖入文件
  2) 自动判断 A3 横向(左右拆分) / 纵向(上下拆分)，允许手动切换
  3) 红色分割线可鼠标拖动，按【每页】保存位置，导出/打印均使用该位置
  4) 拆分预览、A4 结果预览（两页独立：间距 + 边框 + 页号）、导出 A4 PDF
  5) 去密封线（奇数输出页裁左 / 偶数输出页裁右，先裁后等比适配）
  6) 枚举本机打印机并直接打印，强制 A4 纵向、适配安全边距、不从中缝留缝

运行依赖：全部位于本目录 lib/（Pillow / PyMuPDF / pywin32），
拖放为纯 ctypes 的 Win32 WM_DROPFILES 实现（dragdrop.py，标准库），
不依赖 WorkBuddy 私有 Python 或用户级 site-packages。
"""
import _bootstrap  # 必须在任何第三方库之前

import os
import re
import sys
import io
import json
import time
import ctypes
import traceback
import datetime
from ctypes import wintypes

from PIL import Image, ImageDraw, ImageFont, ImageTk, ImageWin

try:
    import dragdrop
except Exception:                       # 拖放不可用时降级为「打开文件」按钮，不影响启动
    dragdrop = None
    _DRAGDROP_IMPORT_ERROR = traceback.format_exc()

try:
    import printguard as pg              # v7：打印结果核验 + 压缩回退
except Exception:
    pg = None
    _PRINTGUARD_IMPORT_ERROR = traceback.format_exc()

try:
    import wordprint as wp              # v8：隐藏 Word COM 一键直印
except Exception:
    wp = None
    _WORDPRINT_IMPORT_ERROR = traceback.format_exc()

# ---------- 常量 ----------
RENDER_DPI = 200
A4_W_MM, A4_H_MM = 210.0, 297.0
A4_W_PX = round(A4_W_MM / 25.4 * RENDER_DPI)   # 1654
A4_H_PX = round(A4_H_MM / 25.4 * RENDER_DPI)   # 2339
SAFE_MARGIN_MM = 3.0
DEFAULT_SEAL_MM = 8.0        # 去密封线默认裁切宽度(mm)，可调

DEFAULT_SPLIT = 0.5          # 默认从正中间拆分
SPLIT_MIN, SPLIT_MAX = 0.20, 0.80   # 分割线可拖动范围，避免半张过窄

PREVIEW_GAP = 28             # 两张 A4 预览之间的间距(px)
PREVIEW_HEADER = 48          # 预览页号区高度(px)
PREVIEW_BORDER = 2           # 每张 A4 预览的边框宽度(px)
PREVIEW_BG = (40, 40, 40)

SUPPORTED_EXTS = (".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff")
SUPPORTED_TEXT = "PDF / JPG / PNG / BMP / TIFF"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_PATH = os.path.join(BASE_DIR, "A3SplitTool.log")
GUI_STATE_PATH = os.path.join(BASE_DIR, "测试输出", "GUI交互状态.json")
# 模拟场景探针会临时改这两个文件名，避免覆盖真机验收证据
PRINTER_STATE_FILE = os.path.join(BASE_DIR, "测试输出", "打印机状态.json")


def log(msg):
    try:
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")
    except Exception:
        pass


# ---------- GDI 原生 DEVMODE：可靠强制 A4 纵向 ----------
# 本机 pywin32 的 CreatePrinterDC 只接受 1 个参数且无 ResetDC 包装，
# 因此改用 gdi32 原生 ResetDCW(hdc, DEVMODEW*) 在创建 DC 后真正应用纸张设置。
DM_ORIENTATION = 0x1
DM_PAPERSIZE = 0x2
DM_PRINTQUALITY = 0x00000400          # v7：降低送驱动分辨率，避免大图压垮 V4/XPS 驱动
DMPAPER_LETTER = 1
DMPAPER_A4 = 9
DMORIENT_PORTRAIT = 1
DMORIENT_LANDSCAPE = 2


class DEVMODEW(ctypes.Structure):
    _fields_ = [
        ("dmDeviceName", ctypes.c_wchar * 32),
        ("dmSpecVersion", ctypes.c_ushort),
        ("dmDriverVersion", ctypes.c_ushort),
        ("dmSize", ctypes.c_ushort),
        ("dmDriverExtra", ctypes.c_ushort),
        ("dmFields", ctypes.c_ulong),
        ("dmOrientation", ctypes.c_short),
        ("dmPaperSize", ctypes.c_short),
        ("dmPaperLength", ctypes.c_short),
        ("dmPaperWidth", ctypes.c_short),
        ("dmScale", ctypes.c_short),
        ("dmCopies", ctypes.c_short),
        ("dmDefaultSource", ctypes.c_short),
        ("dmPrintQuality", ctypes.c_short),
        ("dmColor", ctypes.c_short),
        ("dmDuplex", ctypes.c_short),
        ("dmYResolution", ctypes.c_short),
        ("dmTTOption", ctypes.c_short),
        ("dmCollate", ctypes.c_short),
        ("dmFormName", ctypes.c_wchar * 32),
        ("dmLogPixels", ctypes.c_ushort),
        ("dmBitsPerPel", ctypes.c_ulong),
        ("dmPelsWidth", ctypes.c_ulong),
        ("dmPelsHeight", ctypes.c_ulong),
        ("dmDisplayFlags", ctypes.c_ulong),
        ("dmDisplayFrequency", ctypes.c_ulong),
        ("dmICMMethod", ctypes.c_ulong),
        ("dmICMIntent", ctypes.c_ulong),
        ("dmMediaType", ctypes.c_ulong),
        ("dmDitherType", ctypes.c_ulong),
        ("dmReserved1", ctypes.c_ulong),
        ("dmReserved2", ctypes.c_ulong),
        ("dmPanningWidth", ctypes.c_ulong),
        ("dmPanningHeight", ctypes.c_ulong),
    ]


_GDI = ctypes.WinDLL("gdi32", use_last_error=True)
_GDI.CreateDCW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR, ctypes.c_void_p]
_GDI.CreateDCW.restype = wintypes.HDC
_GDI.ResetDCW.argtypes = [wintypes.HDC, ctypes.c_void_p]
_GDI.ResetDCW.restype = wintypes.HDC
_GDI.DeleteDC.argtypes = [wintypes.HDC]
_GDI.GetDeviceCaps.argtypes = [wintypes.HDC, ctypes.c_int]
_GDI.GetDeviceCaps.restype = ctypes.c_int


def _devmode(device, paper=DMPAPER_A4, orient=DMORIENT_PORTRAIT):
    dm = DEVMODEW()
    dm.dmSize = ctypes.sizeof(DEVMODEW)       # 220
    dm.dmDeviceName = device[:31]
    dm.dmFields = DM_PAPERSIZE | DM_ORIENTATION
    dm.dmPaperSize = paper
    dm.dmOrientation = orient
    return dm


def _dc_caps(hdc):
    g = lambda i: _GDI.GetDeviceCaps(hdc, i)
    return {
        "HORZSIZE_mm": g(4), "VERTSIZE_mm": g(6),
        "PHYSICALWIDTH_px": g(110), "PHYSICALHEIGHT_px": g(111),
        "PHYSICALOFFSETX_px": g(112), "PHYSICALOFFSETY_px": g(113),
        "DPI": g(88),
    }


def _paper_mm_from_caps(c):
    if not c["DPI"]:
        return (0.0, 0.0)
    return (c["PHYSICALWIDTH_px"] / c["DPI"] * 25.4,
            c["PHYSICALHEIGHT_px"] / c["DPI"] * 25.4)


def query_printer_dc(printer_name):
    """创建打印机 DC，应用 PaperSize=9 / Orientation=1 并查询实际设置（含对照实验）。"""
    import win32ui
    hDC = win32ui.CreateDC()
    hDC.CreatePrinterDC(printer_name)          # 本机 pywin32 只接受 1 个参数
    h = hDC.GetHandleOutput()
    caps_default = _dc_caps(h)
    r1 = _GDI.ResetDCW(h, ctypes.byref(_devmode(printer_name, DMPAPER_LETTER, DMORIENT_LANDSCAPE)))
    caps_letter_landscape = _dc_caps(h)
    r2 = _GDI.ResetDCW(h, ctypes.byref(_devmode(printer_name, DMPAPER_A4, DMORIENT_PORTRAIT)))
    caps_a4_portrait = _dc_caps(h)
    hDC.DeleteDC()

    mm_default = _paper_mm_from_caps(caps_default)
    mm_letter = _paper_mm_from_caps(caps_letter_landscape)
    mm_a4 = _paper_mm_from_caps(caps_a4_portrait)
    return {
        "printer": printer_name,
        "resetdcw_ok": bool(r1 and r2),
        "DEVMODEW_size": ctypes.sizeof(DEVMODEW),
        "applied_devmode": {
            "PaperSize": DMPAPER_A4, "Orientation": DMORIENT_PORTRAIT,
            "dmFields": DM_PAPERSIZE | DM_ORIENTATION,
            "meaning": "DMPAPER_A4=9 / DMORIENT_PORTRAIT=1",
        },
        "default": {"caps": caps_default, "paper_mm": tuple(round(x, 1) for x in mm_default)},
        "letter_landscape": {"caps": caps_letter_landscape, "paper_mm": tuple(round(x, 1) for x in mm_letter)},
        "a4_portrait": {"caps": caps_a4_portrait, "paper_mm": tuple(round(x, 1) for x in mm_a4)},
        "devmode_changed_dc": caps_letter_landscape != caps_a4_portrait,
        "verified_a4_portrait": (abs(mm_a4[0] - 210.0) < 3.0
                                 and abs(mm_a4[1] - 297.0) < 3.0
                                 and mm_a4[0] < mm_a4[1]),
        "dc_paper": "DMPAPER_A4(9)" if abs(mm_a4[0] - 210) < 3 and abs(mm_a4[1] - 297) < 3 else "UNKNOWN",
        "dc_orientation": "DMORIENT_PORTRAIT(1)" if mm_a4[0] < mm_a4[1] else "DMORIENT_LANDSCAPE(2)",
    }


def create_a4_printer_dc(printer_name, max_dpi=None):
    """创建并强制 A4 纵向的打印机 DC；验证不通过直接抛错（不静默回退）。

    max_dpi：可选，用 DEVMODE 的 dmPrintQuality 把驱动渲染分辨率降到该值，
    显著降低后台打印文件体积，避免 V4/XPS 打印处理器处理大图失败。
    """
    import win32ui
    hDC = win32ui.CreateDC()
    hDC.CreatePrinterDC(printer_name)
    h = hDC.GetHandleOutput()
    dm = _devmode(printer_name, DMPAPER_A4, DMORIENT_PORTRAIT)
    if max_dpi:
        dm.dmFields |= DM_PRINTQUALITY
        dm.dmPrintQuality = int(max_dpi)
    ret = _GDI.ResetDCW(h, ctypes.byref(dm))
    if not ret:
        hDC.DeleteDC()
        raise RuntimeError(f"ResetDCW 应用 A4 DEVMODE 失败, err={ctypes.get_last_error()}")

    c = _dc_caps(h)
    mm = _paper_mm_from_caps(c)
    ok = (abs(mm[0] - 210.0) < 3.0 and abs(mm[1] - 297.0) < 3.0 and mm[0] < mm[1])
    ev = {"printer": printer_name,
          "applied_devmode": {"PaperSize": dm.dmPaperSize, "Orientation": dm.dmOrientation,
                              "dmFields": dm.dmFields,
                              "PrintQuality": int(max_dpi) if max_dpi else None},
          "caps": c, "paper_mm": (round(mm[0], 1), round(mm[1], 1)),
          "verified_a4_portrait": ok}
    log(f"打印DC强制A4纵向: printer={printer_name} DEVMODE PaperSize=9 Orientation=1 "
        f"PrintQuality={max_dpi or '默认'} 实际DPI={c['DPI']} "
        f"物理={c['PHYSICALWIDTH_px']}x{c['PHYSICALHEIGHT_px']}px "
        f"={mm[0]:.1f}x{mm[1]:.1f}mm HORZ/VERTSIZE={c['HORZSIZE_mm']}x{c['VERTSIZE_mm']}mm "
        f"校验={'PASS' if ok else 'FAIL'}")
    if not ok:
        hDC.DeleteDC()
        raise RuntimeError(f"打印机 DC 未被强制为 A4 纵向: {mm[0]:.1f}x{mm[1]:.1f}mm")
    return hDC, ev, c


# ---------- 文档加载 ----------
def load_document(path):
    ext = os.path.splitext(path)[1].lower()
    pages = []
    if ext == ".pdf":
        import fitz
        doc = fitz.open(path)
        for i, page in enumerate(doc):
            pix = page.get_pixmap(dpi=RENDER_DPI)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            w, h = img.size
            pages.append({"img": img, "orientation": "landscape" if w >= h else "portrait",
                          "src": path, "page": i, "split": DEFAULT_SPLIT})
        doc.close()
    elif ext in SUPPORTED_EXTS:
        img = Image.open(path).convert("RGB")
        w, h = img.size
        pages.append({"img": img, "orientation": "landscape" if w >= h else "portrait",
                      "src": path, "page": 0, "split": DEFAULT_SPLIT})
    else:
        raise ValueError(f"不支持的文件类型：{ext}")
    return pages


def validate_dropped_files(paths):
    """拖入文件的格式校验：返回 (单一文件路径, 错误提示)。"""
    if not paths:
        return None, "未检测到拖入的文件。"
    paths = [p for p in paths if p]
    if len(paths) > 1:
        return None, f"一次只能拖入 1 个文件。\n支持格式：{SUPPORTED_TEXT}"
    p = paths[0]
    ext = os.path.splitext(p)[1].lower()
    if ext not in SUPPORTED_EXTS:
        return None, (f"不支持的文件格式：{ext or '（无扩展名）'}\n"
                      f"仅支持 {SUPPORTED_TEXT}")
    if not os.path.isfile(p):
        return None, f"文件不存在或不是普通文件：\n{p}"
    return p, None


def load_dropped_document(paths):
    """拖入 -> 校验格式 -> 加载页面。返回 (pages, 错误提示)。"""
    p, err = validate_dropped_files(paths)
    if err:
        return None, err
    try:
        pages = load_document(p)
    except Exception as e:
        return None, f"文件无法读取：{os.path.basename(p)}\n{e}"
    if not pages:
        return None, f"文件没有可读取的页面：{os.path.basename(p)}"
    return pages, None


# ---------- 拆分 ----------
def clamp_split(v):
    try:
        v = float(v)
    except Exception:
        return DEFAULT_SPLIT
    return max(SPLIT_MIN, min(v, SPLIT_MAX))


def split_image(img, mode, split=DEFAULT_SPLIT):
    """按 split 比例切分；split=0.5 时与原版逐像素一致。"""
    r = clamp_split(split)
    w, h = img.size
    if mode == "landscape":
        mid = max(1, min(w - 1, round(w * r)))
        return img.crop((0, 0, mid, h)), img.crop((mid, 0, w, h))
    mid = max(1, min(h - 1, round(h * r)))
    return img.crop((0, 0, w, mid)), img.crop((0, mid, w, h))


def set_page_split(pages, index, ratio):
    """按页保存分割位置，返回实际保存值。"""
    if not pages:
        return None
    i = max(0, min(index, len(pages) - 1))
    pages[i]["split"] = clamp_split(ratio)
    return pages[i]["split"]


def split_from_canvas_event(canvas_x, canvas_y, mode, disp, img_w, img_h):
    """把 A3 预览画布上的鼠标坐标换算成该页的分割比例（拖动红线用）。"""
    if not disp or len(disp) < 5:
        return DEFAULT_SPLIT
    x0, y0, scale, nw, nh = disp
    if scale <= 0:
        return DEFAULT_SPLIT
    if mode == "landscape":
        return clamp_split(((canvas_x - x0) / scale) / max(1, img_w))
    return clamp_split(((canvas_y - y0) / scale) / max(1, img_h))


def seal_crop_px(seal_mm, width_px):
    """mm -> 200DPI 像素；上限保护：不超过半张宽度的 1/4。"""
    if not seal_mm or seal_mm <= 0:
        return 0
    return max(0, min(round(seal_mm / 25.4 * RENDER_DPI), width_px // 4))


def make_a4_page(half_img, mode, seal_mm=0.0, out_index=1):
    """
    半张 -> 标准 A4 纵向页。
    seal_mm>0 时【先裁切再等比适配】：按输出 A4 页序号，
    奇数页裁掉左侧密封线、偶数页裁掉右侧密封线，随后 contain 等比适配，
    绝不拉伸。seal_mm=0 时走与原版逐像素完全一致的逻辑（回归保证）。
    """
    half = half_img.rotate(90, expand=True) if mode == "portrait" else half_img
    px = seal_crop_px(seal_mm, half.width)
    if px > 0:
        if out_index % 2 == 1:                       # 奇数页：裁左
            half = half.crop((px, 0, half.width, half.height))
        else:                                        # 偶数页：裁右
            half = half.crop((0, 0, half.width - px, half.height))
    canvas = Image.new("RGB", (A4_W_PX, A4_H_PX), (255, 255, 255))
    hw, hh = half.size
    scale = min(A4_W_PX / hw, A4_H_PX / hh)
    nw, nh = max(1, round(hw * scale)), max(1, round(hh * scale))
    canvas.paste(half.resize((nw, nh), Image.LANCZOS),
                 ((A4_W_PX - nw) // 2, (A4_H_PX - nh) // 2))
    return canvas


def process_page(page_img, auto_orient, manual_mode, seal_mm=0.0, first_out_index=1,
                 split=DEFAULT_SPLIT):
    """处理单页；first_out_index 为该页第一张 A4 在整批输出中的序号(1-based)。"""
    orient = auto_orient if manual_mode == "auto" else manual_mode
    h1, h2 = split_image(page_img, orient, split)
    names = ["左半(A4-1)", "右半(A4-2)"] if orient == "landscape" else ["上半(A4-1)", "下半(A4-2)"]
    idx1, idx2 = first_out_index, first_out_index + 1
    return {
        "orientation": orient,
        "split": clamp_split(split),
        "out_index": [idx1, idx2],
        "halves": [(names[0], make_a4_page(h1, orient, seal_mm, idx1)),
                   (names[1], make_a4_page(h2, orient, seal_mm, idx2))],
    }


def process_document(pages, manual_mode, seal_enabled, seal_mm):
    """
    批量处理：同一批文件/多页 PDF 共用同一套【去密封线】设置，
    但【分割位置按每页独立读取】，输出 A4 页序号全局连续累加。
    关闭开关时 seal_mm 传 0，行为与原逻辑完全一致。
    """
    mm = float(seal_mm) if seal_enabled else 0.0
    a4_images, results = [], []
    idx = 1
    for pg in pages:
        res = process_page(pg["img"], pg["orientation"], manual_mode, mm, idx,
                           pg.get("split", DEFAULT_SPLIT))
        results.append(res)
        a4_images.extend(a4 for _, a4 in res["halves"])
        idx += 2
    return a4_images, results


def _cjk_font(size):
    for p in ("C:/Windows/Fonts/msyh.ttc", "C:/Windows/Fonts/simhei.ttf",
              "C:/Windows/Fonts/simsun.ttc"):
        try:
            return ImageFont.truetype(p, size)
        except Exception:
            pass
    return None


def annotate_seal_preview(combo, seal_mm, first_out_index, a4_w=None, a4_h=None, boxes=None):
    """
    A4 结果预览：半透明灰 + 45° 斜纹标出将裁掉的区域，并写明奇偶裁切方向。
    boxes 为 [(x0,y0,x1,y1), ...]（新版带间距布局）；不传时沿用旧版紧贴布局，
    以保持 verify_seal.py 等旧脚本的调用兼容。
    """
    if boxes is None:
        a4_w = a4_w or A4_W_PX
        a4_h = a4_h or A4_H_PX
        boxes = [(0, 0, a4_w, a4_h), (a4_w, 0, 2 * a4_w, a4_h)]
    px = seal_crop_px(seal_mm, a4_w or (boxes[0][2] - boxes[0][0]))
    if px <= 0:
        return combo
    overlay = Image.new("RGBA", combo.size, (0, 0, 0, 0))
    for k, idx in enumerate((first_out_index, first_out_index + 1)):
        x0, y0, x1, y1 = boxes[k]
        bh = y1 - y0
        bx = x0 if idx % 2 == 1 else x1 - px        # 奇数页裁左 / 偶数页裁右
        band = Image.new("RGBA", (px, bh), (70, 70, 70, 165))
        bd = ImageDraw.Draw(band)
        for s in range(-bh, px + bh, 18):
            bd.line([(s, bh), (s + bh, 0)], fill=(205, 205, 205, 190), width=3)
        overlay.alpha_composite(band, (int(bx), int(y0)))
    out = Image.alpha_composite(combo.convert("RGBA"), overlay).convert("RGB")
    dd = ImageDraw.Draw(out)
    f = _cjk_font(20)
    for k, idx in enumerate((first_out_index, first_out_index + 1)):
        x0, y0, x1, y1 = boxes[k]
        if f is not None:
            label = (f"奇数页{idx} 裁左 {seal_mm:g}mm" if idx % 2 == 1
                     else f"偶数页{idx} 裁右 {seal_mm:g}mm")
        else:
            label = (f"ODD {idx}: cut LEFT {seal_mm:g}mm" if idx % 2 == 1
                     else f"EVEN {idx}: cut RIGHT {seal_mm:g}mm")
        if y0 >= 40:                                 # 新版：写在页号区，不遮挡 A4 内容
            dd.text((x0 + 220, 10), label, font=f, fill=(235, 195, 120))
        else:                                        # 旧版：图内白底红字标签
            dd.rectangle([x0 + 4, 4, x0 + 300, 40], fill=(255, 255, 255))
            dd.text((x0 + 10, 12), label, font=f, fill=(170, 0, 0))
    return out


def annotate_orig_seal(orig_img, orient, seal_mm, first_out_index):
    """原图预览：把「将裁掉的区域」用半透明灰 + 斜纹标出（横向模式映射精确）。"""
    px = seal_crop_px(seal_mm, orig_img.width)
    if px <= 0 or orient != "landscape":
        return orig_img
    overlay = Image.new("RGBA", orig_img.size, (0, 0, 0, 0))
    boxes = [(0, 0, px, orig_img.height),
             (orig_img.width - px, 0, orig_img.width, orig_img.height)]
    for x0, y0, x1, y1 in boxes:
        bw, bh = x1 - x0, y1 - y0
        band = Image.new("RGBA", (bw, bh), (70, 70, 70, 170))
        bd = ImageDraw.Draw(band)
        for s in range(-bh, bw + bh, 18):
            bd.line([(s, bh), (s + bh, 0)], fill=(210, 210, 210, 200), width=3)
        overlay.alpha_composite(band, (x0, y0))
    return Image.alpha_composite(orig_img.convert("RGBA"), overlay).convert("RGB")


def compose_a4_preview(a4a, a4b, first_out_index=1, seal_mm=0.0):
    """
    两张 A4 结果预览：彼此留出间距、各带边框和页号，不再连成一张。
    返回 (预览图, layout)；layout 供独立验证脚本断言间距/边框/页号。
    """
    a4_w, a4_h = a4a.size
    gap, header, border = PREVIEW_GAP, PREVIEW_HEADER, PREVIEW_BORDER
    W, H = a4_w * 2 + gap * 3, a4_h + gap * 2 + header
    combo = Image.new("RGB", (W, H), PREVIEW_BG)
    dd = ImageDraw.Draw(combo)
    f = _cjk_font(26)
    boxes = []
    for k, img in enumerate((a4a, a4b)):
        x, y = gap + k * (a4_w + gap), gap + header
        combo.paste(img, (x, y))
        dd.rectangle([x - border, y - border, x + a4_w + border - 1, y + a4_h + border - 1],
                     outline=(215, 215, 215), width=border)
        idx = first_out_index + k
        label = f"A4 第 {idx} 页  ·  {'奇数页 裁左' if idx % 2 == 1 else '偶数页 裁右'}"
        dd.text((x + 2, gap + 10), label, font=f, fill=(240, 240, 240))
        boxes.append((x, y, x + a4_w, y + a4_h))
    if seal_mm and seal_mm > 0:
        combo = annotate_seal_preview(combo, seal_mm, first_out_index, a4_w, a4_h, boxes)
    layout = {
        "boxes": boxes,
        "gap": gap, "header": header, "border": border,
        "size": combo.size,
        "page_numbers": [first_out_index, first_out_index + 1],
        "gap_between_pages": boxes[1][0] - boxes[0][2],
    }
    return combo, layout


# ---------- 导出 A4 PDF ----------
def export_a4_pdf(a4_images, out_path):
    """导出标准 A4 纵向 PDF；v7 起默认用压缩通道（JPEG + deflate），文件小一个数量级。"""
    if pg is not None:
        return pg.export_compressed_a4_pdf(a4_images, out_path)
    import fitz
    doc = fitz.open()
    for img in a4_images:
        page = doc.new_page(width=A4_W_MM / 25.4 * 72, height=A4_H_MM / 25.4 * 72)
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        page.insert_image(fitz.Rect(0, 0, page.rect.width, page.rect.height), stream=buf.getvalue())
    doc.save(out_path)
    doc.close()
    return out_path


# ---------- 打印机 ----------
def list_printers():
    try:
        import win32print
        flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        return [name for (f, desc, name, comment) in win32print.EnumPrinters(flags)]
    except Exception as e:
        log(f"枚举打印机失败: {e}")
        return []


def get_default_printer():
    """系统默认打印机名；读取失败返回 None。"""
    try:
        import win32print
        return win32print.GetDefaultPrinter()
    except Exception as e:
        log(f"读取系统默认打印机失败: {e}")
        return None


# ---------- v6：实体打印机识别（端口 / 驱动 / 系统属性，多品牌，不写死惠普） ----------
# 设计原则：以「端口 + 驱动 + 名称 + 备注/位置 + 系统属性」多重证据识别虚拟设备，
# 只要命中任一虚拟特征就隐藏；未命中虚拟特征的一律按实体处理，
# 品牌通过多品牌关键字表从名称/驱动/端口/备注中推断，识别不出品牌也照样可用。

_VIRTUAL_PORT_EXACT = {"porprompt:", "nul:", "file:", "shrfax:", "fax:", "null:"}
_VIRTUAL_PORT_PREFIX = ("microsoft.", "porprompt:", "nul:", "shrfax:", "file:",
                        "onenote", "xps", "pdf", "virtual")
_VIRTUAL_TEXT_KEYWORDS = (
    "onenote", "print to pdf", "printtopdf", "microsoft xps", "xps document writer",
    "pdf", "xps", "fax", "传真", "send to", "virtual print", "virtual printer",
    "虚拟", "ipp class driver", "shared fax", "smart printing", "print to file",
    "pdfcreator", "bullzip", "cutepdf", "adobe pdf", "foxit", "dopdf",
)

# 多品牌识别：短品牌名用词边界，避免把别的词误判成品牌；识别不出也不影响实体判定
_BRAND_PATTERNS = (
    ("惠普", r"\bhp\b|hewlett|deskjet|laserjet|officejet|photosmart|\benvy\b|惠普"),
    ("佳能", r"\bcanon\b|佳能|pixma|imageclass|bjnp"),
    ("爱普生", r"\bepson\b|爱普生|stylus|workforce"),
    ("兄弟", r"\bbrother\b|兄弟|\bdcp-|\bmfc-|\bhl-"),
    ("三星", r"\bsamsung\b|三星|\bscx-|\bclp-|\bml-"),
    ("奔图", r"\bpantum\b|奔图"),
    ("联想", r"\blenovo\b|联想"),
    ("京瓷", r"\bkyocera\b|京瓷|taskalfa"),
    ("理光", r"\bricoh\b|理光|\bsavin\b|lanier"),
    ("施乐", r"\bxerox\b|施乐|docuprint|docucentre"),
    ("柯尼卡美能达", r"konica|minolta|bizhub|柯尼卡"),
    ("利盟", r"\blexmark\b|利盟"),
    ("戴尔", r"\bdell\b|戴尔"),
    ("得力", r"\bdeli\b|得力"),
    ("佳博", r"gprinter|佳博"),
    ("斑马", r"\bzebra\b|斑马"),
    ("富士通", r"fujitsu|富士通"),
    ("冲电气", r"\boki\b"),
    ("夏普", r"\bsharp\b|夏普"),
    ("东芝", r"\btoshiba\b|东芝"),
)

# win32print 打印机状态位（只用数值，避免不同 pywin32 版本常量名差异）
_STATUS_BITS = (
    (0x00000001, "已暂停", True), (0x00000002, "报告错误", True),
    (0x00000004, "正在删除", True), (0x00000008, "卡纸", True),
    (0x00000010, "缺纸", True), (0x00000020, "需手动进纸", False),
    (0x00000040, "纸张问题", True), (0x00000080, "离线", True),
    (0x00000100, "正在读写", False), (0x00000200, "繁忙", False),
    (0x00000400, "正在打印", False), (0x00000800, "出纸盒已满", True),
    (0x00001000, "当前不可用", True), (0x00002000, "等待打印", False),
    (0x00004000, "正在处理", False), (0x00008000, "正在初始化", False),
    (0x00010000, "正在预热", False), (0x00020000, "墨粉不足", False),
    (0x00040000, "无墨粉", True), (0x00080000, "页面被跳过", False),
    (0x00100000, "需要人工处理", True), (0x00200000, "内存不足", True),
    (0x00400000, "机盖未关", True), (0x00800000, "服务器未知", False),
    (0x01000000, "省电状态", False),
)
_ATTRIBUTE_WORK_OFFLINE = 0x00000400


def read_printer_facts(name):
    """用 Win32 读取端口、驱动、系统属性、状态位。任何失败都只写日志，不抛给用户。"""
    f = {"端口": "", "驱动": "", "备注": "", "位置": "", "系统属性": 0,
         "状态位": 0, "打印处理器": "", "打开成功": False, "读取错误": ""}
    try:
        import win32print
        h = win32print.OpenPrinter(name)
        try:
            d = win32print.GetPrinter(h, 2)
        finally:
            win32print.ClosePrinter(h)
        f.update({
            "端口": str(d.get("pPortName") or ""),
            "驱动": str(d.get("pDriverName") or ""),
            "备注": str(d.get("pComment") or ""),
            "位置": str(d.get("pLocation") or ""),
            "系统属性": int(d.get("Attributes") or 0),
            "状态位": int(d.get("Status") or 0),
            "打印处理器": str(d.get("pPrintProcessor") or ""),
            "打开成功": True,
        })
    except Exception as e:
        f["读取错误"] = f"{type(e).__name__}: {e}"
        log(f"读取打印机属性失败（仅日志）: {name}: {e}")
    return f


def _port_kind(port):
    p = (port or "").strip().lower()
    if p.startswith("usb"):
        return "USB"
    if re.match(r"^lpt\d*:?$", p) or "parallel" in p:
        return "并口"
    if re.match(r"^com\d*:?$", p) or "serial" in p:
        return "串口"
    if p.startswith("wsd-"):
        return "网络(WSD)"
    if p.startswith("ip_") or p.startswith("\\\\") or "tcp/ip" in p or p.startswith("standard tcp"):
        return "网络(TCP/IP)"
    if p.startswith("dot4") or "bjnp" in p:
        return "USB/复合端口"
    if "bluetooth" in p:
        return "蓝牙"
    return "其它端口"


def _attribute_text(attrs):
    names = []
    for bit, txt in ((0x00000040, "本地"), (0x00000010, "网络"), (0x00000008, "共享"),
                     (0x00000004, "系统默认"), (0x00000020, "隐藏"),
                     (_ATTRIBUTE_WORK_OFFLINE, "脱机工作")):
        if attrs & bit:
            names.append(txt)
    return "、".join(names) or "常规"


def detect_brand(name, facts):
    """从名称/驱动/端口/备注/位置推断品牌；识别不出返回「未识别品牌」但仍是实体。"""
    fields = [("名称", name), ("驱动", facts.get("驱动", "")),
              ("端口", facts.get("端口", "")), ("备注", facts.get("备注", "")),
              ("位置", facts.get("位置", ""))]
    for brand, pat in _BRAND_PATTERNS:
        for label, val in fields:
            if val and re.search(pat, val.lower()):
                return brand, f"{label}「{val}」"
    return "未识别品牌", "端口/驱动/系统属性中未匹配到已知品牌，仍按实体打印机处理"


def classify_printer(name, facts):
    """纯分类函数：给定名称与 Win32 事实，返回是否实体 / 隐藏原因 / 品牌 / 判定依据。"""
    port = facts.get("端口", "") or ""
    driver = facts.get("驱动", "") or ""
    comment = facts.get("备注", "") or ""
    location = facts.get("位置", "") or ""
    plow = port.strip().lower()
    reasons = []
    if plow in _VIRTUAL_PORT_EXACT or any(plow.startswith(x) for x in _VIRTUAL_PORT_PREFIX):
        reasons.append(f"端口「{port}」是虚拟输出端口")
    for label, val in (("名称", name), ("驱动", driver), ("备注", comment), ("位置", location)):
        low = (val or "").lower()
        hit = next((k for k in _VIRTUAL_TEXT_KEYWORDS if k in low), None)
        if hit:
            reasons.append(f"{label}含虚拟设备特征「{hit}」")
    physical = not reasons
    attrs = int(facts.get("系统属性", 0) or 0)
    evidence = []
    if physical:
        evidence.append(f"端口类型={_port_kind(port)}（{port or '未知'}）")
        evidence.append(f"驱动={driver or '未知'}")
        evidence.append(f"系统属性={_attribute_text(attrs)}（{hex(attrs)}）")
    brand, brand_src = detect_brand(name, facts)
    return {
        "名称": name, "端口": port, "驱动": driver, "备注": comment, "位置": location,
        "端口类型": _port_kind(port), "系统属性": attrs,
        "属性说明": _attribute_text(attrs), "状态位": int(facts.get("状态位", 0) or 0),
        "是否实体": physical, "隐藏原因": reasons, "实体判定依据": evidence,
        "品牌": brand, "品牌识别依据": brand_src,
    }


def printer_runtime_state(facts):
    """由状态位 + 系统属性 + 能否打开句柄，判定在线/离线，输出简体中文状态。"""
    status = int(facts.get("状态位", 0) or 0)
    attrs = int(facts.get("系统属性", 0) or 0)
    blocking, warning = [], []
    if not facts.get("打开成功", False):
        blocking.append("无法连接打印机，设备可能已断开或驱动不可用")
    if attrs & _ATTRIBUTE_WORK_OFFLINE:
        blocking.append("打印机被系统设为脱机工作")
    for bit, txt, is_block in _STATUS_BITS:
        if status & bit:
            (blocking if is_block else warning).append(txt)
    if blocking:
        cn = "不可用：" + "、".join(blocking)
        if warning:
            cn += "；另有提醒：" + "、".join(warning)
    elif warning:
        cn = "在线·" + "、".join(warning)
    else:
        cn = "在线·就绪"
    return {"在线": not blocking, "中文状态": cn,
            "阻断原因": blocking, "提醒": warning}


def collect_printers(names=None, facts_provider=None):
    """枚举并逐台分析，返回全部设备（含被隐藏者）的完整信息，供界面与验收脚本核验。"""
    if names is None:
        names = list_printers()
    provider = facts_provider or read_printer_facts
    out = []
    for n in names:
        facts = provider(n)
        info = classify_printer(n, facts)
        info.update(printer_runtime_state(facts))
        out.append(info)
    return out


def visible_printers(infos, default=None):
    """只保留真实实体打印机；排序：系统默认优先 → 惠普实体优先 → 名称。"""
    vis = [i for i in infos if i.get("是否实体")]
    dl = (default or "").lower()
    vis.sort(key=lambda i: (0 if dl and i["名称"].lower() == dl else 1,
                            0 if i["品牌"] == "惠普" else 1,
                            i["名称"]))
    return vis


def choose_visible_printer(infos, default=None):
    """默认选中：可用（在线）实体中，系统默认优先 → 惠普实体优先 → 第一台。"""
    vis = visible_printers(infos, default)
    if not vis:
        return None
    online = [i for i in vis if i.get("在线")]
    pool = online or vis
    dl = (default or "").lower()
    if dl:
        for i in pool:
            if i["名称"].lower() == dl:
                return i["名称"]
    for i in pool:
        if i["品牌"] == "惠普":
            return i["名称"]
    return pool[0]["名称"]


def can_print(infos):
    """是否存在「可见且在线」的实体打印机。"""
    return any(i.get("是否实体") and i.get("在线") for i in infos)


def recommend_printer(infos, default=None):
    name = choose_visible_printer(infos, default)
    return name or "下拉列表中的实体打印机"


def build_visible_options(infos, chosen=None):
    """实体打印机下拉项：品牌 + 中文在线状态 + 默认星标。"""
    opts, mapping = [], {}
    for i in infos:
        brand_tag = "[实体·%s]" % i["品牌"] if i["品牌"] != "未识别品牌" else "[实体打印机]"
        display = (("★ " if i["名称"] == chosen else "") + i["名称"] +
                   "  " + brand_tag + "  [" + i["中文状态"] + "]")
        opts.append(display)
        mapping[display] = i["名称"]
    return opts, mapping


# 旧版按名称分级（保留给 v5 回归脚本与低层调用），面向用户的文案已去掉写死的惠普型号
UNSUPPORTED_PRINTER_KEYWORDS = ("onenote", "fax", "传真", "send to", "print to onenote")
VIRTUAL_PRINTER_KEYWORDS = ("print to pdf", "xps", "smart printing", "pdfcreator")
PHYSICAL_HP_KEYWORDS = ("hp", "deskjet", "laserjet", "officejet", "envy", "惠普")


def printer_status(name):
    """返回 (等级, 标注, 建议)；等级 = unsupported / virtual / physical。

    建议文案会直接进入界面与弹窗，必须使用简体中文，不得夹带英文技术串。
    """
    low = (name or "").lower()
    for k in UNSUPPORTED_PRINTER_KEYWORDS:
        if k in low:
            return ("unsupported", "[不支持直接打印·勿选]",
                    f"「{name}」不是实体打印机，而是虚拟或转发目标；"
                    "它虽然能创建打印画布，却无法启动打印文档。"
                    "请切换到下拉列表中的实体打印机，"
                    "或改用「导出A4 PDF → 系统打印」兜底。")
    for k in VIRTUAL_PRINTER_KEYWORDS:
        if k in low:
            return ("virtual", "[虚拟·会弹保存框]",
                    f"「{name}」是虚拟打印机，会弹出保存或转发对话框，而不是直接出纸；"
                    "需要实体打印请切换到下拉列表中的实体打印机。")
    hp = any(k in low for k in PHYSICAL_HP_KEYWORDS)
    return ("physical", "[实体打印机·推荐]" if hp else "[实体打印机]",
            f"「{name}」是实体打印机，可直接按 A4 纵向打印。")


def is_hp_physical(name):
    return (printer_status(name)[0] == "physical"
            and any(k in (name or "").lower() for k in PHYSICAL_HP_KEYWORDS))


def choose_default_printer(names, default=None):
    """
    下拉默认：系统默认打印机（若为实体）；系统默认不可用/虚拟时优先实际 HP 实体，
    再退到任意实体，最后才用系统默认 / 第一项。绝不默认选中 OneNote/Fax。
    """
    if not names:
        return None
    if default is None:
        default = get_default_printer()
    d = None
    if default:
        d = next((n for n in names if n == default), None)
        if d is None:
            d = next((n for n in names if n.lower() == default.lower()), None)
    if d and printer_status(d)[0] == "physical":
        return d
    for n in names:
        if is_hp_physical(n):
            return n
    for n in names:
        if printer_status(n)[0] == "physical":
            return n
    return d or names[0]


def build_printer_options(names, chosen=None):
    """下拉显示项（带标注）与「显示文本 -> 真实打印机名」映射。"""
    opts, mapping = [], {}
    for n in names:
        level, tag, hint = printer_status(n)
        display = ("★ " if n == chosen else "") + f"{n}  {tag}"
        opts.append(display)
        mapping[display] = n
    return opts, mapping


def printer_name_from_display(display, mapping):
    """显示文本还原成真实打印机名；兼容旧代码直接传名字。"""
    return mapping.get(display, display)


class PrinterStartDocError(RuntimeError):
    """启动打印文档失败的专用异常。

    __str__ 保留英文原始异常，仅供写入日志；面向用户一律用 user_message()。
    """

    def __init__(self, printer, original, hints):
        self.printer = printer
        self.original = str(original)
        self.hints = list(hints)
        super().__init__(
            f"打印机「{printer}」启动打印文档失败。\n\n"
            f"原始错误（仅日志）：{original}\n\n可操作建议：\n" +
            "\n".join(f"  {i+1}. {h}" for i, h in enumerate(self.hints)) +
            "\n\n本次未向打印机发送任何页面，不会消耗纸张；工具也没有自动重试。")

    def user_message(self):
        """纯简体中文的用户提示，不含任何英文异常串。"""
        return ("打印机「" + self.printer + "」无法启动打印文档。\n\n"
                "常见原因：所选目标不是实体打印机，或打印机驱动暂时不可用。\n\n"
                "建议按下面顺序处理：\n" +
                "\n".join(f"  {i+1}）{h}" for i, h in enumerate(self.hints)) +
                "\n\n本次没有向打印机发送任何页面，不会消耗纸张，也没有自动重试。")


def start_doc_failure_hints(printer, default_recommend=None):
    rec = default_recommend or "本机在线的实体打印机"
    return [
        printer_status(printer)[2],
        f"把「打印机」下拉切换到实体打印机（推荐「{rec}」），再点「打印」。",
        "或点「导出A4 PDF → 系统打印」，先生成 A4 文件，再用系统打印对话框手动确认。",
        "不要反复重试同一目标，避免误触发虚拟目标的保存或转发对话框。",
    ]


def print_a4_images(printer_name, a4_images, safe_margin_mm=SAFE_MARGIN_MM,
                    dry_run=False, compress=True):
    """直接打印：强制 A4 纵向、适配安全边距、每半张独立占满一页。

    v7：送驱动前把 DC 分辨率降到 300dpi、把每页位图降采样到长边 1754px，
    避免大位图让 V4/XPS 打印处理器渲染失败（本机实证 0x6BE）。
    """
    max_dpi = (pg.PRINT_DC_DPI if pg else None) if compress else None
    hDC, ev, c = create_a4_printer_dc(printer_name, max_dpi=max_dpi)
    DPI = c["DPI"] or 600
    safe = int(safe_margin_mm / 25.4 * DPI)
    x0 = c["PHYSICALOFFSETX_px"] + safe
    y0 = c["PHYSICALOFFSETY_px"] + safe
    pw = max(1, c["PHYSICALWIDTH_px"] - 2 * c["PHYSICALOFFSETX_px"] - 2 * safe)
    ph = max(1, c["PHYSICALHEIGHT_px"] - 2 * c["PHYSICALOFFSETY_px"] - 2 * safe)
    rect = (x0, y0, pw, ph)
    send_images, compression = list(a4_images), []
    if compress and pg is not None:
        send_images = []
        for img in a4_images:
            small, st = pg.compress_for_print(img)
            send_images.append(small)
            compression.append(st)
    log(f"可印区={pw}x{ph}px 安全边距={safe_margin_mm}mm 待打印页数={len(a4_images)} "
        f"dry_run={dry_run} 驱动DPI={DPI} 位图压缩统计={compression}")
    result = {"pages": len(a4_images), "evidence": ev,
              "printable_rect": rect, "压缩统计": compression,
              "送驱动DPI": DPI}

    if dry_run:
        hDC.DeleteDC()
        return result
    try:
        try:
            hDC.StartDoc("A3拆分A4打印")
        except Exception as e:
            # StartDoc 失败（OneNote/Fax 等虚拟目标）：给出具体打印机名 + 可操作建议，
            # 不重试、不向打印机送页、不消耗纸张。
            log(f"STARTDOC 失败: printer={printer_name} err={e} 未发送任何页面")
            raise PrinterStartDocError(
                printer_name, e, start_doc_failure_hints(printer_name)) from e
        for img in send_images:
            hDC.StartPage()
            dib = ImageWin.Dib(img)
            iw, ih = img.size
            scale = min(pw / iw, ph / ih)
            dw, dh = max(1, int(iw * scale)), max(1, int(ih * scale))
            dx, dy = x0 + (pw - dw) // 2, y0 + (ph - dh) // 2
            dib.draw(hDC.GetHandleOutput(), (dx, dy, dx + dw, dy + dh))
            hDC.EndPage()
        hDC.EndDoc()
    finally:
        hDC.DeleteDC()
    return result


def choose_print_backend(printer, facts=None):
    """v8：自动判定提交方式，绝不弹窗询问用户。

    系统文档渲染通道（V4 / XPS / Class Driver）或该打印机有 GDI 直打失败历史
    （打印服务 842 / 0x6BE）→ 隐藏 Word COM 直印；
    其它常规驱动的实体打印机 → 仍走 GDI。返回 (后端, 简体中文原因, 打印机事实)。
    """
    facts = facts or read_printer_facts(printer)
    if wp is None:
        return "gdi", "Word 直印模块不可用，按常规方式打印", facts
    if pg is None:
        return "gdi", "打印结果核验模块不可用，按常规方式打印", facts
    try:
        ch, why = pg.recommended_channel(printer, facts)
    except Exception as e:
        log(f"PRINTGUARD 提交方式判定失败（仅日志）: {e}")
        ch, why = "gdi", "提交方式判定不可用，按常规方式打印"
    if ch == "pdf":
        why = why.replace("已自动改走更稳妥的文件通道", "直接送驱动曾失败")
        return "word", (why + "；自动改用隐藏 Word COM 直印"
                        "（点「打印」即直接打印，无需选择）"), facts
    return "gdi", "常规驱动，可直接发送打印任务", facts


def print_with_verification(printer, a4_images, wait_seconds=None, backend=None,
                            facts_provider=None, jobs_provider=None,
                            events_provider=None):
    """一键提交（Word 直印 / GDI 二选一，自动判定）后核验「队列 + 842 事件」。

    绝不把「提交成功」当作「打印成功」；只有核验结论「失败」才由 GUI 层
    用中文提示并自动保留 A4 PDF。全部数据源可注入，便于无纸测试。
    """
    if pg is None:
        raise RuntimeError("printguard 模块不可用")
    facts_fn = facts_provider or read_printer_facts
    facts = facts_fn(printer)
    if backend is None:
        backend, reason, _ = choose_print_backend(printer, facts)
    else:
        reason = "由调用方指定提交方式（测试/诊断用）"
    before = []
    try:
        jobs_fn = jobs_provider or (lambda: pg.snapshot_jobs(printer))
        before = [j["作业号"] for j in (jobs_fn() or [])]
    except Exception as e:
        log("PRINTGUARD 读取打印前队列失败（仅日志）: " + str(e))
    started = datetime.datetime.now(datetime.timezone.utc).replace(tzinfo=None)
    port = facts.get("端口", "") or ""
    log(f"PRINTGUARD 准备提交: printer={printer} 端口={port} 提交方式={backend} "
        f"原因={reason} 队列前作业={before}")
    if backend == "word":
        if wp is None:
            raise RuntimeError("Word 直印模块不可用")
        submitted = wp.print_via_word(printer, a4_images, port=port)
    else:
        submitted = print_a4_images(printer, a4_images)
    ver = pg.verify_print_job(printer, before, started,
                              wait_seconds=(pg.VERIFY_WAIT_SECONDS if wait_seconds is None
                                            else wait_seconds),
                              jobs_provider=jobs_provider,
                              events_provider=events_provider)
    try:
        pg.record_channel_result(printer, ver["结论"], ver["中文说明"],
                                 {"驱动": facts.get("驱动", ""), "提交方式": backend})
    except Exception as e:
        log("PRINTGUARD 记录通道历史失败（仅日志）: " + str(e))
    return {"提交结果": submitted, "核验": ver, "提交方式": backend,
            "提交方式原因": reason, "打印机端口": port}


# ---------- 自测模式 ----------
def selftest():
    base = os.path.dirname(os.path.abspath(__file__))
    sample_dir = os.path.join(base, "测试样例")
    out_dir = os.path.join(base, "测试输出")
    os.makedirs(out_dir, exist_ok=True)

    print("=== 0. 运行环境 / 依赖来源 ===")
    rep = _bootstrap.dependency_report()
    for k, v in rep.items():
        print(f"  {k} = {v}")
    assert rep["uses_tool_lib"], "依赖仍来自工具目录之外（不合格）"
    print("  -> PASS: 全部依赖来自本目录 lib，不依赖 WorkBuddy 私有 Python / 用户级 site-packages")
    log("SELFTEST 依赖来源校验 PASS: " + str(rep))

    from make_samples import make_all
    print("\n=== 1. 生成 A3 测试样本 ===")
    paths = make_all(sample_dir)
    for p in paths:
        print("  ", p)

    total_a4 = 0
    a4_imgs_all = []
    for sp in paths:
        print(f"\n=== 2. 拆分校验 {os.path.basename(sp)} ===")
        pages = load_document(sp)
        print(f"  页数={len(pages)} 原图方向={pages[0]['orientation']}")
        a4_imgs = []
        for pi, pg in enumerate(pages):
            res = process_page(pg["img"], pg["orientation"], "auto")
            print(f"  第{pi+1}页 -> 模式={res['orientation']} 半张尺寸={res['halves'][0][1].size}")
            for nm, a4 in res["halves"]:
                assert a4.size == (A4_W_PX, A4_H_PX), f"A4 尺寸错误: {a4.size}"
                a4_imgs.append(a4)
                total_a4 += 1
            h1, h2 = res["halves"][0][1], res["halves"][1][1]
            c1 = h1.getpixel((h1.width // 2, h1.height // 2))
            c2 = h2.getpixel((h2.width // 2, h2.height // 2))
            left_blue, right_orange = (c1[2] > c1[0]), (c2[0] > c2[2])
            print(f"    半张1中心色={c1} 半张2中心色={c2} 归属正确={left_blue and right_orange}")
            assert left_blue and right_orange, f"内容归属错误: {c1},{c2}"
            raw1, raw2 = split_image(pg["img"], res["orientation"])
            ow, oh = pg["img"].size
            seam_ok = ((raw1.width + raw2.width == ow) if res["orientation"] == "landscape"
                       else (raw1.height + raw2.height == oh))
            print(f"    中缝无缝(两半拼回原图)={seam_ok}")
            assert seam_ok, "拆分存在缺口或重叠"
        a4_imgs_all.extend(a4_imgs)
        pdf_path = os.path.join(out_dir, os.path.splitext(os.path.basename(sp))[0] + "_A4.pdf")
        export_a4_pdf(a4_imgs, pdf_path)
        print(f"  导出 A4 PDF: {pdf_path} ({os.path.getsize(pdf_path)} bytes)")

    print("\n=== 2.5 去密封线功能 + 原功能回归 ===")
    from verify_seal import run as verify_seal_run
    seal_rep = verify_seal_run(base)
    print(f"  去密封线验证结论: {seal_rep['结论']}")
    assert seal_rep["结论"] == "PASS"
    log("SELFTEST 去密封线/回归 PASS")

    print("\n=== 2.6 三项新功能（分割线/预览/拖放）+ 旧功能回归 ===")
    from verify_new_features import run as verify_new_run
    new_rep = verify_new_run(base, skip_gui=True)
    assert new_rep["结论"] == "PASS", "新功能验证未通过"
    print(f"  新功能验证结论: {new_rep['结论']}")
    log("SELFTEST 新功能验证 PASS: " + str(new_rep["检查项"]))

    print("\n=== 3. 打印机枚举 + DC 强制 A4 纵向验证（真实查询） ===")
    printers = list_printers()
    print("  本机打印机:", printers)
    log("SELFTEST 打印机枚举: " + str(printers))
    sys_def = get_default_printer()
    chosen = choose_default_printer(printers, sys_def)
    print(f"  系统默认={sys_def}  下拉默认选中={chosen}  等级={printer_status(chosen)[0]}")
    log(f"SELFTEST 默认打印机选择: 系统默认={sys_def} 下拉默认={chosen} "
        f"等级={printer_status(chosen)[0]}")
    assert printer_status(chosen)[0] == "physical", "默认打印机不应是虚拟/不支持GDI目标"
    targets = [p for p in ("Microsoft Print to PDF", "HP DeskJet 4900 series", "HP Smart Printing")
               if p in printers]
    assert targets, "未找到可验证的打印机"
    for t in targets:
        ev = query_printer_dc(t)
        print(f"\n  [{t}]")
        print(f"    ResetDCW 成功={ev['resetdcw_ok']} DEVMODEW size={ev['DEVMODEW_size']}")
        print(f"    应用 DEVMODE: PaperSize=9(DMPAPER_A4) Orientation=1(PORTRAIT) dmFields=3")
        print(f"    默认 DC          物理纸={ev['default']['paper_mm']}mm")
        print(f"    强制 Letter/横向 物理纸={ev['letter_landscape']['paper_mm']}mm")
        print(f"    强制 A4/纵向     物理纸={ev['a4_portrait']['paper_mm']}mm caps={ev['a4_portrait']['caps']}")
        print(f"    DEVMODE 改变了 DC={ev['devmode_changed_dc']}  校验 A4纵向={ev['verified_a4_portrait']}")
        print(f"    DC 实际 => PaperSize={ev['dc_paper']}  Orientation={ev['dc_orientation']}")
        log("SELFTEST DC验证 " + str(ev))
        assert ev["resetdcw_ok"] and ev["devmode_changed_dc"] and ev["verified_a4_portrait"], \
            f"打印机 {t} 未通过 A4 纵向强制验证"

    print("\n=== 4. 打印路径 dry_run（创建 DC + 校验，不弹保存框） ===")
    res = print_a4_images(targets[0], a4_imgs_all[:2], dry_run=True)
    print(f"  目标={targets[0]} 将打印页数={res['pages']} 可印区={res['printable_rect']}")
    print(f"  DC 证据: {res['evidence']}")
    assert res["evidence"]["verified_a4_portrait"]
    log("SELFTEST 打印dry_run PASS: " + str(res))

    print(f"\n自测完成：共 {total_a4} 张 A4，全部 {A4_W_PX}x{A4_H_PX}；"
          f"拆分归属/中缝/导出/去密封线/新功能/打印DC强制A4 全部 PASS。")
    return True


# ---------- 弹窗文案探针（仅验收用，正常启动不触发） ----------
def run_ui_text_probe(gui, root):
    """
    在真实 App 上触发「打开 / 导出 / 打印 / PDF 兜底 / 拖入 / 切换打印机」
    的全部弹窗分支，把标题与正文写入 测试输出/弹窗文案.json。
    文件对话框、导出、打印、打开程序全部替换为桩：不弹真实窗口、
    不产生大文件、不发送任何打印任务、不消耗纸张。
    """
    import tkinter.messagebox as mb
    from tkinter import filedialog

    captured = []
    originals = {}

    def recorder(name, default):
        def _f(*args, **kwargs):
            title = args[0] if args else kwargs.get("title", "")
            text = args[1] if len(args) > 1 else kwargs.get("message", "")
            captured.append({"函数": name, "标题": str(title), "正文": str(text)})
            return default
        return _f

    for name, default in (("showerror", "ok"), ("showwarning", "ok"),
                          ("showinfo", "ok"), ("askyesno", True),
                          ("askokcancel", True), ("askretrycancel", False)):
        originals[name] = getattr(mb, name)
        setattr(mb, name, recorder(name, default))

    g = globals()
    saved = {"open": filedialog.askopenfilename,
             "save": filedialog.asksaveasfilename,
             "dir": filedialog.askdirectory,
             "export": g["export_a4_pdf"],
             "print": g["print_a4_images"],
             "load": g["load_dropped_document"],
             "startfile": os.startfile,
             "write_state": gui._write_state,
             "refresh_channel": gui._refresh_channel,
             "snapshot": pg.snapshot_jobs if pg else None,
             "verify": pg.verify_print_job if pg else None,
             "fallback": pg.save_a4_fallback if pg else None,
             "openfolder": pg.open_folder if pg else None,
             "record": pg.record_channel_result if pg else None,
             "word_submit": wp.print_via_word if wp else None}
    sample = os.path.join(BASE_DIR, "测试样例", "sample_seal_multipage.pdf")
    tmp_pdf = os.path.join(BASE_DIR, "测试输出", "弹窗探测_A4.pdf")
    out = os.path.join(BASE_DIR, "测试输出", "弹窗文案.json")
    hp_display = next((d for d, n in gui.printer_map.items()
                       if n == gui.chosen_printer), None)
    verification = {"结论": "失败"}
    verify_calls = []

    def _boom(*a, **k):
        raise RuntimeError("win32ui.error: StartDoc failed")

    def _fake_export(imgs, path):
        with open(path, "wb") as f:
            f.write(b"%PDF-1.4\n")
        return path

    def _startdoc(*a, **k):
        raise PrinterStartDocError(
            gui.chosen_printer, "win32ui.error: StartDoc failed",
            start_doc_failure_hints(gui.chosen_printer))

    def _generic(*a, **k):
        raise OSError("win32ui.error: CreatePrinterDC failed")

    def _load_boom(*a, **k):
        raise RuntimeError("win32ui.error: image decode failed")

    def _fake_fallback(*a, **k):
        return {"文件夹": os.path.join(BASE_DIR, "测试输出", "打印失败回退"),
                "合并PDF": tmp_pdf, "单页PDF": [tmp_pdf, tmp_pdf],
                "页数": 2, "文件": [{"路径": tmp_pdf, "角色": "模拟文件"}]}

    def _fake_verify(*a, **k):
        v = dict(verification)
        v.setdefault("中文说明", "模拟：打印作业核验结果。")
        verify_calls.append(v["结论"])
        return v

    def _force_gdi():
        gui.channel, gui.channel_reason = "gdi", "模拟：常规驱动，可直接发送打印任务"

    try:
        gui._write_state = lambda: None       # 不覆盖真实交互证据文件
        filedialog.askopenfilename = lambda *a, **k: sample
        filedialog.asksaveasfilename = lambda *a, **k: tmp_pdf
        filedialog.askdirectory = lambda *a, **k: ""     # 不实际写文件
        os.startfile = lambda *a, **k: (_ for _ in ()).throw(
            RuntimeError("simulated open failure"))
        gui._refresh_channel = _force_gdi
        _force_gdi()
        if pg:
            pg.snapshot_jobs = lambda *a, **k: []
            pg.verify_print_job = _fake_verify
            pg.save_a4_fallback = _fake_fallback
            pg.open_folder = lambda *a, **k: True
            pg.record_channel_result = lambda *a, **k: {}
        # 关键：探针里绝不调用真实打印，Word 直印与 GDI 两条提交路径都替换为桩
        g["print_a4_images"] = lambda *a, **k: {
            "pages": 2, "evidence": {}, "printable_rect": (0, 0, 1, 1),
            "压缩统计": [], "送驱动DPI": 300}
        if wp:
            # v8：Word 直印提交委托给同一个桩，异常模拟（StartDoc/通用异常）同样生效
            wp.print_via_word = lambda *a, **k: g["print_a4_images"](*a, **k)

        gui.open_file()                                   # 正常加载
        filedialog.askopenfilename = lambda *a, **k: r"C:\fake\试卷.docx"
        gui.open_file()                                   # 打开格式错误
        filedialog.askopenfilename = lambda *a, **k: sample
        g["load_dropped_document"] = _load_boom
        gui.open_file()                                   # 打开时加载异常
        gui.on_drop_files([sample])                       # 拖入时加载异常
        g["load_dropped_document"] = saved["load"]

        g["export_a4_pdf"] = _boom
        gui.export_pdf()                                  # 导出失败
        gui.safe_fallback_print()                         # 兜底生成失败
        g["export_a4_pdf"] = _fake_export
        gui.safe_fallback_print()                         # 兜底生成成功但打开失败
        g["export_a4_pdf"] = saved["export"]

        # v6：无实体打印机 / 打印按钮禁用时的中文拦截
        saved_infos = gui.visible_infos
        gui.visible_infos = []
        gui.do_print()                                    # 无法打印 + 选择文件夹
        gui.visible_infos = saved_infos

        # v7：PDF 稳妥通道、打开回退文件夹（startfile 被桩化为失败）
        gui.pdf_channel_print()                           # 稳妥通道成功
        gui.open_fallback_folder()                        # 打开文件夹失败（中文）

        # v7：打印后核验的三种结论 + 两类提交异常，全部不出纸
        if hp_display:
            gui.cmb_printer.set(hp_display)
        for conclusion in ("失败", "未检出错误", "待确认"):
            verification.clear()
            verification.update({"结论": conclusion,
                                 "中文说明": "模拟：系统打印队列核验结果为" + conclusion + "。"})
            gui.do_print()
        g["print_a4_images"] = _startdoc
        gui.do_print()                                    # 启动打印文档失败 → 自动回退
        g["print_a4_images"] = _generic
        gui.do_print()                                    # 其它打印异常 → 自动回退

        gui.on_drop_files([r"C:\fake\试卷.docx"])         # 拖入格式错误
        old_pages = gui.pages
        gui.pages = []
        gui.export_pdf()                                  # 未打开文件的提示
        gui.safe_fallback_print()
        gui.pdf_channel_print()
        gui.do_print()
        gui.pages = old_pages
    finally:
        for name, fn in originals.items():
            setattr(mb, name, fn)
        filedialog.askopenfilename = saved["open"]
        filedialog.asksaveasfilename = saved["save"]
        filedialog.askdirectory = saved["dir"]
        g["export_a4_pdf"] = saved["export"]
        g["print_a4_images"] = saved["print"]
        g["load_dropped_document"] = saved["load"]
        os.startfile = saved["startfile"]
        gui._write_state = saved["write_state"]
        gui._refresh_channel = saved["refresh_channel"]
        if pg:
            pg.snapshot_jobs = saved["snapshot"]
            pg.verify_print_job = saved["verify"]
            pg.save_a4_fallback = saved["fallback"]
            pg.open_folder = saved["openfolder"]
            pg.record_channel_result = saved["record"]
        if wp:
            wp.print_via_word = saved["word_submit"]
        os.makedirs(os.path.dirname(out), exist_ok=True)
        with open(out, "w", encoding="utf-8") as f:
            json.dump({"生成时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                       "触发分支数": len(captured), "核验结论序列": verify_calls,
                       "弹窗": captured},
                      f, ensure_ascii=False, indent=2)
        log(f"UI_TEXT_PROBE 弹窗分支数={len(captured)} 证据={out} "
            f"核验结论序列={verify_calls} 自动打印任务=0")


# ---------- GUI ----------
def main():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    class App:
        def __init__(self, root):
            self.root = root
            self.root.title("A3 试卷拆分打印工具")
            self.root.geometry("1280x840")
            self.pages = []
            self.cur = 0
            self.manual = tk.StringVar(value="auto")
            self.seal_on = tk.BooleanVar(value=False)
            self.seal_mm = tk.StringVar(value=str(DEFAULT_SEAL_MM))
            self.cur_a4 = []
            self._photos = []
            self._disp = {}                 # canvas id -> (x0, y0, scale, nw, nh)
            self._dragging = False
            self._drag_ratio = None
            self._preview_layout = None
            self._dnd_ok = False
            self._drag_start_split = None
            self._drag_start_pos = None
            self._drag_span = 1
            self._last_drop = (None, 0.0)
            # 真实 GUI 交互证据：由鼠标/拖放事件写入，供独立验证脚本核验
            self._state = {"已加载文件": None, "真实拖放次数": 0,
                           "分割线按下次数": 0, "分割线拖动次数": 0,
                           "实时预览次数": 0, "拖动轨迹": [], "最终split": None}
            self._build()

        def _write_state(self):
            """把真实交互状态写盘，供 real_gui_verify.py 读取核验。"""
            try:
                self._state["更新时间"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                # A3 原图在屏幕上的实际矩形，供真实鼠标拖动验证精确定位
                disp = self._disp.get(id(self.orig_cv))
                if disp:
                    x0, y0, scale, nw, nh = disp
                    rx, ry = self.orig_cv.winfo_rootx(), self.orig_cv.winfo_rooty()
                    self._state["A3图像屏幕矩形"] = [rx + x0, ry + y0,
                                                 rx + x0 + nw, ry + y0 + nh]
                os.makedirs(os.path.dirname(GUI_STATE_PATH), exist_ok=True)
                with open(GUI_STATE_PATH, "w", encoding="utf-8") as f:
                    json.dump(self._state, f, ensure_ascii=False, indent=2)
            except Exception:
                log("写 GUI交互状态 失败: " + traceback.format_exc())

        def _build(self):
            frm_top = ttk.Frame(self.root)
            frm_top.pack(fill="x", padx=8, pady=6)
            ttk.Button(frm_top, text="打开文件 (PDF/JPG/PNG/BMP/TIFF)",
                       command=self.open_file).pack(side="left")
            self.lbl_file = ttk.Label(frm_top, text="未打开文件")
            self.lbl_file.pack(side="left", padx=10)
            ttk.Label(frm_top, text="拆分方式:").pack(side="left", padx=(10, 2))
            for txt, val in (("自动", "auto"), ("横向左右", "landscape"), ("纵向上下", "portrait")):
                ttk.Radiobutton(frm_top, text=txt, variable=self.manual, value=val,
                                command=self.recompute).pack(side="left")
            self.lbl_orient = ttk.Label(frm_top, text="")
            self.lbl_orient.pack(side="left", padx=10)

            # --- 去密封线（开关 + 裁切宽度）---
            frm_seal = ttk.Frame(self.root)
            frm_seal.pack(fill="x", padx=8, pady=(0, 4))
            ttk.Checkbutton(frm_seal, text="去密封线", variable=self.seal_on,
                            command=self.recompute).pack(side="left")
            ttk.Label(frm_seal, text="裁切宽度(mm):").pack(side="left", padx=(10, 2))
            ent = ttk.Entry(frm_seal, textvariable=self.seal_mm, width=7)
            ent.pack(side="left")
            ent.bind("<KeyRelease>", lambda e: self.recompute())
            self.lbl_seal_rule = ttk.Label(frm_seal, text="关闭时保持原拆分行为不变")
            self.lbl_seal_rule.pack(side="left", padx=12)

            # --- 分割线 + 拖放提示 ---
            frm_split = ttk.Frame(self.root)
            frm_split.pack(fill="x", padx=8, pady=(0, 4))
            ttk.Label(frm_split, text="分割线:").pack(side="left")
            self.lbl_split = ttk.Label(frm_split, text="50%（拖动 A3 区红线，按页保存）")
            self.lbl_split.pack(side="left", padx=8)
            ttk.Button(frm_split, text="重置本页分割线", command=self.reset_split).pack(side="left")
            self.lbl_dnd = ttk.Label(frm_split, text="")
            self.lbl_dnd.pack(side="left", padx=12)

            frm_mid = ttk.Frame(self.root)
            frm_mid.pack(fill="both", expand=True, padx=8)
            frm_mid.columnconfigure(0, weight=1)
            frm_mid.columnconfigure(1, weight=1)
            frm_mid.rowconfigure(0, weight=1)
            frm_orig = ttk.LabelFrame(frm_mid, text="原图预览 (A3，红色分割线可拖动 | 可拖入文件)")
            frm_orig.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
            self.orig_cv = tk.Canvas(frm_orig, bg="#222")
            self.orig_cv.pack(fill="both", expand=True)
            self.orig_cv.bind("<ButtonPress-1>", self._orig_press)
            self.orig_cv.bind("<B1-Motion>", self._orig_motion)
            self.orig_cv.bind("<ButtonRelease-1>", self._orig_release)
            frm_res = ttk.LabelFrame(frm_mid, text="A4 拆分结果预览（两张独立：间距 / 边框 / 页号）")
            frm_res.grid(row=0, column=1, sticky="nsew", padx=4, pady=4)
            self.res_cv = tk.Canvas(frm_res, bg="#222")
            self.res_cv.pack(fill="both", expand=True)

            frm_nav = ttk.Frame(self.root)
            frm_nav.pack(fill="x", padx=8)
            ttk.Button(frm_nav, text="上一页", command=self.prev).pack(side="left")
            ttk.Button(frm_nav, text="下一页", command=self.next).pack(side="left")
            self.lbl_page = ttk.Label(frm_nav, text="")
            self.lbl_page.pack(side="left", padx=10)

            frm_bot = ttk.Frame(self.root)
            frm_bot.pack(fill="x", padx=8, pady=6)
            ttk.Button(frm_bot, text="生成 A4 PDF", command=self.export_pdf).pack(side="left")
            ttk.Button(frm_bot, text="选择文件夹保存 A4 PDF",
                       command=self.save_a4_to_folder).pack(side="left", padx=(8, 0))
            ttk.Button(frm_bot, text="打开回退文件夹",
                       command=self.open_fallback_folder).pack(side="left", padx=(8, 0))
            ttk.Label(frm_bot, text="打印机:").pack(side="left", padx=(12, 2))
            self.cmb_printer = ttk.Combobox(frm_bot, values=[], width=58, state="readonly")
            self.cmb_printer.bind("<<ComboboxSelected>>", self.on_printer_selected)
            self.cmb_printer.pack(side="left")
            self.btn_print = ttk.Button(frm_bot, text="打印", command=self.do_print)
            self.btn_print.pack(side="left", padx=(8, 0))
            ttk.Button(frm_bot, text="刷新状态", command=self.reload_printers).pack(
                side="left", padx=(6, 0))

            self.lbl_printer = ttk.Label(self.root, text="", wraplength=1200,
                                         foreground="#e8a33d")
            self.lbl_printer.pack(fill="x", padx=8, pady=(0, 2))
            self.last_fallback = None
            # 启动时就落盘真实打印机状态，供独立验收脚本核验下拉默认选中
            self.reload_printers(write_state=True)

            self.status = tk.Label(self.root, text="就绪", anchor="w", relief="sunken")
            self.status.pack(fill="x", side="bottom")

            # --- Windows 原生文件拖放（纯 ctypes，离线可用）---
            if dragdrop is not None:
                try:
                    self._dnd_ok = dragdrop.hook_dropfiles(self.root, self.on_drop_files)
                    log(f"DRAGDROP 注册={self._dnd_ok} 状态={dragdrop.status()}")
                except Exception:
                    self._dnd_ok = False
                    log("DRAGDROP 注册异常: " + traceback.format_exc())
            else:
                log("DRAGDROP 模块导入失败: " + globals().get("_DRAGDROP_IMPORT_ERROR", ""))
            self.lbl_dnd.config(text=("A3 区可拖入 PDF/JPG/PNG/BMP/TIFF" if self._dnd_ok
                                      else "拖放不可用，请用「打开文件」"))
            log(f"GUI 初始化完成: python={sys.executable} "
                f"依赖={_bootstrap.dependency_report()} 打印机数={len(self.printers)} "
                f"拖放={self._dnd_ok}")

        def set_status(self, msg):
            self.status.config(text=msg)
            self.root.update_idletasks()

        def _fail(self, title, cn_text, exc=None):
            """统一错误弹窗：标题与正文只用简体中文；英文原始异常只写日志。"""
            if exc is not None:
                log(f"错误弹窗【{title}】原始异常（仅日志）: "
                    f"{type(exc).__name__}: {exc}")
                log(traceback.format_exc())
            messagebox.showerror(title, cn_text)

        def open_file(self):
            path = filedialog.askopenfilename(
                title="选择 A3 试卷文件",
                filetypes=[("常见文件", "*.pdf *.jpg *.jpeg *.png *.bmp *.tif *.tiff"),
                           ("PDF", "*.pdf"), ("图片", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff")])
            if not path:
                return
            try:
                pages, err = load_dropped_document([path])
            except Exception as e:
                self._fail("文件打开失败",
                           "无法读取选中的文件。文件可能已损坏、正被其它程序占用，"
                           "或格式不受支持。\n\n"
                           "请确认文件是 PDF、JPG、PNG、BMP 或 TIFF 格式后重试。", e)
                self.set_status("加载失败")
                return
            if err:
                self._fail("格式错误", err)
                self.set_status("打开失败")
                return
            try:
                self.set_status("正在加载...")
                self.pages = pages
                self.cur = 0
                self.lbl_file.config(text=os.path.basename(path))
                self.recompute()
                self.set_status(f"已加载 {len(pages)} 页；可拖动红线调整分割位置")
            except Exception as e:
                self._fail("文件打开失败",
                           "无法读取选中的文件。文件可能已损坏、正被其它程序占用，"
                           "或格式不受支持。\n\n"
                           "请确认文件是 PDF、JPG、PNG、BMP 或 TIFF 格式后重试。", e)
                self.set_status("加载失败")

        def on_drop_files(self, paths):
            """拖放回调：由 dragdrop 的 Tk 主线程轮询队列调用，绝不来自 COM/RPC 线程。"""
            key = paths[0] if paths else None
            now = time.time()
            # 顶层窗口与子窗口可能都收到 WM_DROPFILES，1.5 秒内同路径只处理一次
            if key and self._last_drop[0] == key and now - self._last_drop[1] < 1.5:
                log("DRAGDROP 重复消息已去重: " + str(key))
                return
            self._last_drop = (key, now)
            self._state["真实拖放次数"] += 1
            self._state["已加载文件"] = key
            self._write_state()
            try:
                pages, err = load_dropped_document(paths)
            except Exception as e:
                self._fail("拖入文件加载失败",
                           "拖入的文件无法读取。请确认文件完整、未被其它程序占用，"
                           "并且是 PDF、JPG、PNG、BMP 或 TIFF 格式。", e)
                self.set_status("拖放失败：文件无法读取")
                return
            if err:
                log("DRAGDROP 格式错误: " + err.replace("\n", " | "))
                self._fail("格式错误", err)
                self.set_status("拖放失败：" + err.splitlines()[0])
                return
            try:
                self.pages = pages
                self.cur = 0
                self.lbl_file.config(text=os.path.basename(pages[0]["src"]))
                self.recompute()
                self.set_status(f"拖入加载成功：{len(pages)} 页；红线可拖动")
                log(f"DRAGDROP 加载成功: {paths} 页数={len(pages)}")
            except Exception as e:
                self._fail("拖入文件加载失败",
                           "拖入的文件无法读取。请确认文件完整、未被其它程序占用，"
                           "并且是 PDF、JPG、PNG、BMP 或 TIFF 格式。", e)

        def seal_value(self):
            """全局去密封线设置；同一批文件/多页 PDF 内保持一致，切换文件不重置。"""
            if not self.seal_on.get():
                return 0.0
            try:
                v = float(self.seal_mm.get())
            except Exception:
                v = DEFAULT_SEAL_MM
            return max(0.0, min(v, 60.0))

        def _orig_mode(self):
            pg = self.pages[self.cur]
            return pg["orientation"] if self.manual.get() == "auto" else self.manual.get()

        def _orig_press(self, ev):
            """
            鼠标按下：只要在 A3 图像区域内即进入【持续拖动】状态，
            不要求必须精确点中红线；记录起点，之后由 <B1-Motion> 连续跟随。
            """
            if not self.pages:
                return
            disp = self._disp.get(id(self.orig_cv))
            if not disp:
                return
            x0, y0, scale, nw, nh = disp
            inside = (x0 - 8 <= ev.x <= x0 + nw + 8 and
                      y0 - 8 <= ev.y <= y0 + nh + 8)
            self._dragging = bool(inside)
            self._drag_ratio = None
            if not self._dragging:
                return
            mode = self._orig_mode()
            self._drag_start_split = self.pages[self.cur].get("split", DEFAULT_SPLIT)
            self._drag_start_pos = ev.x if mode == "landscape" else ev.y
            self._drag_span = max(1, nw if mode == "landscape" else nh)
            self.orig_cv.config(cursor=("sb_h_double_arrow" if mode == "landscape"
                                        else "sb_v_double_arrow"))
            self._state["分割线按下次数"] += 1
            self._write_state()
            self.set_status("已按住：左右/上下连续移动鼠标，红线实时跟随并刷新 A4 预览；松开保存")
            log(f"分割线按下: 第{self.cur+1}页 起始split={self._drag_start_split:.4f} "
                f"mode={mode} 起点={self._drag_start_pos}")

        def _orig_motion(self, ev):
            """
            按住后连续移动：按位移增量持续更新分割比例，限制在 20%-80%，
            并【实时】重算两张 A4 预览（不是等松开才更新）。
            """
            if not self._dragging or self._drag_start_split is None:
                return
            mode = self._orig_mode()
            cur = ev.x if mode == "landscape" else ev.y
            delta = (cur - self._drag_start_pos) / self._drag_span
            r = clamp_split(self._drag_start_split + delta)
            self._drag_ratio = r
            self._state["分割线拖动次数"] += 1
            if len(self._state["拖动轨迹"]) < 40:
                self._state["拖动轨迹"].append(round(r, 4))
            self._state["最终split"] = r
            try:
                self._live_preview(r)
            except Exception:
                log("实时预览异常: " + traceback.format_exc())
            self._write_state()

        def _live_preview(self, r):
            """拖动过程中的实时 A4 预览：与导出/打印使用完全相同的处理链路。"""
            pg = self.pages[self.cur]
            mm = self.seal_value()
            idx = self.cur * 2 + 1
            set_page_split(self.pages, self.cur, r)
            res = process_page(pg["img"], pg["orientation"], self.manual.get(), mm, idx, r)
            self.cur_a4 = res["halves"]
            self.show_preview(pg["img"], res, mm, idx)
            self._state["实时预览次数"] += 1
            self.lbl_split.config(text=f"分割位置 {r*100:.1f}%（拖动中 · A4 实时预览）")

        def _orig_release(self, ev):
            """松开：保存本页分割位置，全量重算，写状态与日志。"""
            if not self._dragging:
                return
            self._dragging = False
            if self._drag_ratio is not None:
                set_page_split(self.pages, self.cur, self._drag_ratio)
            self._drag_ratio = None
            self._drag_start_split = None
            self.orig_cv.config(cursor="")
            try:
                self.orig_cv.delete("dragline")
            except Exception:
                pass
            self.recompute()
            sp = self.pages[self.cur].get("split", DEFAULT_SPLIT)
            self._state["最终split"] = sp
            self._write_state()
            log(f"分割线拖动保存: 第{self.cur+1}页 split={sp:.4f} "
                f"拖动次数={self._state['分割线拖动次数']} "
                f"实时预览次数={self._state['实时预览次数']}")
            self.set_status(f"第 {self.cur+1} 页分割位置已保存 {sp*100:.1f}%；导出/打印均使用该位置")

        def reset_split(self):
            if not self.pages:
                return
            set_page_split(self.pages, self.cur, DEFAULT_SPLIT)
            self.recompute()
            self.set_status(f"第 {self.cur+1} 页分割位置已重置为 50%")

        def recompute(self):
            if not self.pages:
                return
            mm = self.seal_value()
            idx = self.cur * 2 + 1
            pg = self.pages[self.cur]
            split = pg.get("split", DEFAULT_SPLIT)
            self.lbl_orient.config(text=f"原图方向: {pg['orientation']}  | 使用: {self.manual.get()}")
            res = process_page(pg["img"], pg["orientation"], self.manual.get(), mm, idx, split)
            self.cur_a4 = res["halves"]
            self.show_preview(pg["img"], res, mm, idx)
            self.lbl_page.config(text=f"第 {self.cur+1} / {len(self.pages)} 页")
            self.lbl_split.config(text=f"第 {self.cur+1} 页分割位置 {split*100:.1f}%（已按页保存）")
            if mm > 0:
                i1, i2 = res["out_index"]
                self.lbl_seal_rule.config(
                    text=f"去密封线 {mm:g}mm：本页 A4-{i1}(奇数页 裁左) / "
                         f"A4-{i2}(偶数页 裁右)；全批一致")
                self.set_status(f"已拆分并去密封线 {mm:g}mm，分割位置 {split*100:.1f}%")
            else:
                self.lbl_seal_rule.config(text="关闭时保持原拆分行为不变")
                self.set_status(f"已拆分：{res['orientation']}，分割位置 {split*100:.1f}%，生成 2 张 A4")
            self._write_state()

        def show_preview(self, orig_img, res, seal_mm=0.0, first_out_index=1):
            self._photos = []
            split = self.pages[self.cur].get("split", DEFAULT_SPLIT)
            od = orig_img.copy()
            dw, dh = od.size
            draw = ImageDraw.Draw(od)
            lw = max(2, dw // 400)
            if res["orientation"] == "landscape":
                x = round(dw * split)
                draw.line([(x, 0), (x, dh)], fill=(255, 0, 0), width=lw)
            else:
                y = round(dh * split)
                draw.line([(0, y), (dw, y)], fill=(255, 0, 0), width=lw)
            if seal_mm and seal_mm > 0:
                od = annotate_orig_seal(od, res["orientation"], seal_mm, first_out_index)
            self._show_on_canvas(self.orig_cv, od)
            a4a, a4b = res["halves"][0][1], res["halves"][1][1]
            combo, layout = compose_a4_preview(a4a, a4b, first_out_index, seal_mm)
            self._preview_layout = layout
            self._show_on_canvas(self.res_cv, combo)

        def _show_on_canvas(self, canvas, img):
            cw, ch = canvas.winfo_width() or 480, canvas.winfo_height() or 480
            scale = min(cw / img.width, ch / img.height)
            nw, nh = max(1, int(img.width * scale)), max(1, int(img.height * scale))
            photo = ImageTk.PhotoImage(img.resize((nw, nh), Image.LANCZOS))
            self._photos.append(photo)
            canvas.delete("all")
            cx, cy = cw // 2, ch // 2
            canvas.create_image(cx, cy, image=photo, anchor="center")
            self._disp[id(canvas)] = (cx - nw // 2, cy - nh // 2, nw / img.width, nw, nh)

        def prev(self):
            if self.pages and self.cur > 0:
                self.cur -= 1
                self.recompute()

        def next(self):
            if self.pages and self.cur < len(self.pages) - 1:
                self.cur += 1
                self.recompute()

        def all_a4(self):
            # 批量处理：每页使用自己保存的分割位置，去密封线设置全批一致
            imgs, _ = process_document(self.pages, self.manual.get(), True, self.seal_value())
            return imgs

        def _info_by_name(self, name):
            return next((i for i in self.visible_infos if i["名称"] == name), None)

        def reload_printers(self, write_state=True):
            """重新枚举并判定；下拉只显示真实实体打印机，按系统默认 / 惠普实体排序。"""
            self.sys_default = get_default_printer()
            self.all_infos = collect_printers()
            self.printers = [i["名称"] for i in self.all_infos]
            self.visible_infos = visible_printers(self.all_infos, self.sys_default)
            self.chosen_printer = choose_visible_printer(self.all_infos, self.sys_default)
            self.printer_options, self.printer_map = build_visible_options(
                self.visible_infos, self.chosen_printer)
            self.cmb_printer["values"] = self.printer_options
            if self.chosen_printer:
                idx = next((k for k, i in enumerate(self.visible_infos)
                            if i["名称"] == self.chosen_printer), 0)
                self.cmb_printer.current(idx)
            else:
                self.cmb_printer.set("")
            self._refresh_channel()
            self._refresh_printer_hint()
            self._update_print_availability()
            if write_state:
                self._write_printer_state()
            hidden = [i["名称"] for i in self.all_infos if not i["是否实体"]]
            log(f"打印机刷新: 系统默认={self.sys_default} 实体可见={len(self.visible_infos)} "
                f"隐藏虚拟设备={hidden} 排序={[i['名称'] for i in self.visible_infos]} "
                f"默认选中={self.chosen_printer}")

        def _refresh_channel(self):
            """v8：自动判定提交方式，只用于状态栏说明，绝不弹窗让用户选。"""
            name = printer_name_from_display(self.cmb_printer.get(), self.printer_map)
            if not name:
                self.channel, self.channel_reason = "gdi", "尚未选择实体打印机"
                return
            self.channel, self.channel_reason, facts = choose_print_backend(name)
            log(f"PRINTGUARD 提交方式判定: 打印机={name} 通道={self.channel} "
                f"原因={self.channel_reason} 驱动={facts.get('驱动', '')} "
                f"端口={facts.get('端口', '')}")

        def _refresh_printer_hint(self):
            name = printer_name_from_display(self.cmb_printer.get(), self.printer_map)
            info = self._info_by_name(name)
            if not info:
                self.lbl_printer.config(
                    text=f"系统默认打印机：{self.sys_default or '（未知）'}  |  "
                         "未检测到可用的真实实体打印机；虚拟设备（OneNote、传真、生成 PDF 等）已隐藏。")
                return
            ch = "隐藏 Word 直印（自动）" if self.channel == "word" else "直接送驱动（自动）"
            self.lbl_printer.config(
                text=f"系统默认打印机：{self.sys_default or '（未知）'}  |  "
                     f"当前：{name}（实体·{info['品牌']}）  |  状态：{info['中文状态']}\n"
                     f"识别依据：端口 {info['端口'] or '未知'}（{info['端口类型']}）  |  "
                     f"驱动 {info['驱动'] or '未知'}  |  系统属性 {info['属性说明']}\n"
                     f"打印方式：{ch}（{self.channel_reason}）。点「打印」即直接打印，"
                     "不再选择通道；提交后自动核验打印队列与 842 事件，失败才提示并保留 A4 文件。")

        def _print_block_reason(self):
            """返回打印禁用原因（简体中文）；可用时返回 None。"""
            if not self.visible_infos:
                return "没有检测到真实实体打印机"
            if not can_print(self.all_infos):
                return "所有实体打印机都处于离线或不可用状态"
            name = printer_name_from_display(self.cmb_printer.get(), self.printer_map)
            info = self._info_by_name(name)
            if not info:
                return "请从下拉列表中选择一台实体打印机"
            if not info["在线"]:
                return f"「{name}」{info['中文状态']}"
            return None

        def _update_print_availability(self):
            reason = self._print_block_reason()
            if reason:
                self.btn_print.config(state="disabled")
                self.cmb_printer.config(state="disabled" if not self.visible_infos else "readonly")
                self.lbl_printer.config(
                    text=self.lbl_printer.cget("text") +
                         f"\n打印已禁用：{reason}。请点「选择文件夹保存 A4 PDF」，"
                         "把 A4 文件保存后自行打印；本工具不会自动发送打印任务、不会消耗纸张。",
                    foreground="#e05c5c")
            else:
                self.btn_print.config(state="normal")
                self.cmb_printer.config(state="readonly")
                self.lbl_printer.config(foreground="#e8a33d")

        def _write_printer_state(self):
            """把识别结果落盘，供独立验证脚本核验（可复现证据）。"""
            try:
                os.makedirs(os.path.dirname(PRINTER_STATE_FILE), exist_ok=True)
                data = {
                    "生成时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "系统默认打印机": self.sys_default,
                    "下拉默认选中实际名": self.chosen_printer,
                    "下拉默认选中显示": self.cmb_printer.get(),
                    "选中判定": printer_status(self.chosen_printer)[0] if self.chosen_printer else "无",
                    "打印按钮状态": str(self.btn_print["state"]),
                    "打印禁用原因": self._print_block_reason(),
                    "可见实体打印机": [
                        {"名称": i["名称"], "品牌": i["品牌"], "端口": i["端口"],
                         "端口类型": i["端口类型"], "驱动": i["驱动"],
                         "系统属性": i["属性说明"], "中文状态": i["中文状态"],
                         "在线": i["在线"], "品牌识别依据": i["品牌识别依据"],
                         "实体判定依据": i["实体判定依据"]}
                        for i in self.visible_infos],
                    "已隐藏虚拟设备": [
                        {"名称": i["名称"], "端口": i["端口"], "驱动": i["驱动"],
                         "隐藏原因": i["隐藏原因"]}
                        for i in self.all_infos if not i["是否实体"]],
                    "排序规则": "系统默认优先 → 惠普实体优先 → 名称",
                    "全部打印机": [
                        {"名称": n, "显示文本": self.printer_options[i]
                         if i < len(self.printer_options) else "（已隐藏）",
                         "等级": printer_status(n)[0], "标注": printer_status(n)[1],
                         "建议": printer_status(n)[2]}
                        for i, n in enumerate(self.printers)],
                }
                with open(PRINTER_STATE_FILE, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                log(f"打印机状态: 系统默认={self.sys_default} "
                    f"下拉默认选中={self.chosen_printer} "
                    f"打印按钮={self.btn_print['state']}")
            except Exception:
                log("写打印机状态失败: " + traceback.format_exc())

        def on_printer_selected(self, ev=None):
            """选中实体打印机后刷新中文状态；离线/不可用则禁用打印并给出中文提示。"""
            name = printer_name_from_display(self.cmb_printer.get(), self.printer_map)
            info = self._info_by_name(name)
            self._refresh_printer_hint()
            self._update_print_availability()
            if not info:
                return
            log(f"打印机切换: {name} 品牌={info['品牌']} 状态={info['中文状态']}")
            if not info["在线"]:
                messagebox.showwarning(
                    "打印机当前不可用",
                    f"「{name}」目前无法打印。\n\n"
                    f"状态：{info['中文状态']}\n\n"
                    "本工具已禁用「打印」按钮，不会向该设备发送任何任务，也不会消耗纸张。\n"
                    "请检查电源、连接线、是否设为脱机、驱动是否正常，然后点「刷新状态」；\n"
                    "也可以点「选择文件夹保存 A4 PDF」，保存后自行打印。")

        def save_a4_to_folder(self):
            """无可用实体打印机时的替代路径：选择文件夹并保存 A4 PDF，绝不发送打印任务。"""
            if not self.pages:
                messagebox.showwarning("提示", "请先打开文件")
                return
            folder = filedialog.askdirectory(
                title="选择保存 A4 PDF 的文件夹", mustexist=False)
            if not folder:
                return
            path = os.path.join(folder, "A3拆分_A4.pdf")
            try:
                os.makedirs(folder, exist_ok=True)
                self.set_status("正在保存 A4 PDF 到文件夹...")
                imgs = self.all_a4()
                export_a4_pdf(imgs, path)
                log(f"文件夹保存 A4 PDF: {path} 页数={len(imgs)} 未发送打印任务")
                self.set_status(f"已保存 A4 文件（未打印）：{path}")
                messagebox.showinfo(
                    "已保存 A4 文件",
                    f"A4 文件已保存到：\n{path}\n\n"
                    f"共 {len(imgs)} 页，每页均为标准 A4 纵向。\n"
                    "本工具没有发送任何打印任务，也不会消耗纸张。\n"
                    "请在连接好实体打印机后自行打开该文件打印。")
            except Exception as e:
                self._fail("保存 A4 文件失败",
                           "无法把 A4 文件保存到所选文件夹。\n\n"
                           "请确认文件夹可写、磁盘空间充足，或改选其它文件夹后重试。", e)

        def _make_fallback_files(self, prefix="A3拆分_A4"):
            """生成压缩的合并 PDF + 每张 A4 一个单页 PDF，并记录到状态，供打开文件夹。"""
            if pg is None:
                raise RuntimeError("printguard 模块不可用")
            fb = pg.save_a4_fallback(self.all_a4(), prefix=prefix)
            self.last_fallback = fb
            return fb

        def _auto_fallback(self, title, chinese_reason, evidence=None):
            """任何打印失败都自动保留两张 A4 PDF 并打开文件夹，中文说明原因。"""
            self.set_status("打印未完成：正在保留 A4 文件...")
            fb = None
            try:
                fb = self._make_fallback_files("A3拆分_打印失败")
            except Exception as e:
                self._fail("保留 A4 文件失败",
                           "打印没有完成，而且自动保留 A4 文件时也出错了。\n\n"
                           "请改用「生成 A4 PDF」或「选择文件夹保存 A4 PDF」手动保存，"
                           "再自行打印。", e)
                return
            try:
                pg.open_folder(fb["文件夹"])
                folder_note = "保存文件夹已自动打开。"
            except Exception as e:
                log(f"打开回退文件夹失败（仅日志）: {e}")
                folder_note = "请点「打开回退文件夹」查看文件。"
            files = "\n".join(os.path.basename(f["路径"]) for f in fb["文件"])
            log(f"PRINTGUARD 自动回退完成: 原因={chinese_reason} 文件数={len(fb['文件'])} "
                f"文件夹={fb['文件夹']} 核验={evidence}")
            messagebox.showerror(
                title,
                chinese_reason + "\n\n"
                f"已自动保留 {fb['页数']} 张 A4 页的压缩文件：\n{files}\n\n"
                f"保存位置：{fb['文件夹']}\n{folder_note}\n"
                "本工具不会自动重试、不会再次发送打印任务，也不会消耗纸张。\n"
                "请在文件夹里打开 A4 文件，用阅读器的打印功能手动确认后打印。")
            self.set_status("打印未完成：已自动保留 A4 文件并打开文件夹（未重复发送任务）")

        def pdf_channel_print(self):
            """稳妥通道：只生成压缩 A4 文件并打开文件夹，不向 V4/XPS 驱动送大图。"""
            if not self.pages:
                messagebox.showwarning("提示", "请先打开文件")
                return
            try:
                self.set_status("正在生成压缩 A4 文件（稳妥通道）...")
                fb = self._make_fallback_files("A3拆分_稳妥通道")
                pg.open_folder(fb["文件夹"])
                files = "\n".join(os.path.basename(f["路径"]) for f in fb["文件"])
                self.set_status(f"已生成 {fb['页数']} 张 A4 文件（未发送打印任务）")
                messagebox.showinfo(
                    "已生成压缩 A4 文件",
                    f"已生成 {fb['页数']} 张标准 A4 纵向页面，并保留为压缩文件：\n{files}\n\n"
                    f"保存位置：{fb['文件夹']}（文件夹已打开）\n\n"
                    "这条路不经过容易失败的文档渲染通道，文件体积小、兼容各品牌打印机。\n"
                    "请用阅读器打开 A4 文件，选择实体打印机、A4 纸张、纵向后打印。\n"
                    "本工具没有发送任何打印任务，不会消耗纸张。")
            except Exception as e:
                self._fail("生成 A4 文件失败",
                           "无法生成压缩 A4 文件。\n\n"
                           "请确认磁盘可写、空间充足；也可以改用「生成 A4 PDF」按钮另存。", e)

        def open_fallback_folder(self):
            """打开最近一次自动保留 A4 文件的文件夹；没有则打开默认回退目录。"""
            folder = (self.last_fallback or {}).get("文件夹") or (
                pg.FALLBACK_DIR if pg else os.path.join(BASE_DIR, "测试输出"))
            try:
                os.makedirs(folder, exist_ok=True)
                os.startfile(folder)
                self.set_status(f"已打开文件夹：{folder}")
            except Exception as e:
                self._fail("打开文件夹失败",
                           f"无法打开文件夹：\n{folder}\n\n请手动在资源管理器中打开该位置。", e)

        def direct_print_with_verification(self, pn):
            """一键提交后核验队列与 842 事件；只有失败才中文提示并自动保留 A4 PDF。

            「提交成功」不等于「已打印」：非失败结论也不弹任何成功提示，
            只在状态栏与日志如实写明核验结论，请人工确认出纸。
            """
            self.set_status("正在直接打印并核验打印队列与 842 事件，请稍候...")
            try:
                r = print_with_verification(pn, self.all_a4())
                ver = r["核验"]
                if ver["结论"] == "失败":
                    self._auto_fallback("打印未完成", ver["中文说明"], ver)
                    return
                self.set_status("提交并核验：" + ver["中文说明"])
                log(f"PRINTGUARD 非失败结论: 提交方式={r.get('提交方式')} "
                    f"结论={ver['结论']} 说明={ver['中文说明']} 未把提交当作成功")
            except PrinterStartDocError as e:
                log("启动打印文档失败（已用中文提示）: " + str(e).replace("\n", " | "))
                self._auto_fallback("打印失败：无法启动打印文档", e.user_message())
            except Exception as e:
                self._auto_fallback("打印失败",
                                    "打印过程中发生错误，本次没有完成打印，页面内容可能没有送达打印机。",
                                    {"原始异常": f"{type(e).__name__}: {e}"})
                log("打印异常（原始信息仅日志）: " + traceback.format_exc())

        def safe_fallback_print(self):
            """
            安全兜底：导出 A4 PDF → 用户确认 → 用系统默认程序打开 PDF。
            只调 open，绝不调用 ShellExecute 的 print verb，也不自动重试 GDI 打印，
            因此不会自动消耗纸张；是否出纸完全由用户在系统打印对话框里手动确认。
            """
            if not self.pages:
                messagebox.showwarning("提示", "请先打开文件")
                return
            path = filedialog.asksaveasfilename(
                title="保存 A4 PDF（系统打印兜底）", defaultextension=".pdf",
                filetypes=[("PDF", "*.pdf")], initialfile="A3拆分_A4_系统打印.pdf")
            if not path:
                return
            try:
                self.set_status("正在生成 A4 PDF（兜底）...")
                imgs = self.all_a4()
                export_a4_pdf(imgs, path)
                log(f"兜底A4 PDF 已生成: {path} 页数={len(imgs)} 未自动打印")
            except Exception as e:
                self._fail("生成 A4 文件失败",
                           "无法生成用于系统打印的 A4 文件。\n\n"
                           "请确认保存位置可写、磁盘空间充足，然后重试；"
                           "也可以改用「生成 A4 PDF」按钮另存到其它位置。", e)
                return
            rec = recommend_printer(self.all_infos, self.sys_default)
            open_it = messagebox.askyesno(
                "安全兜底：用系统打印",
                f"A4 文件已生成：\n{path}\n\n"
                "本工具不会自动发送打印任务，因此不会消耗纸张。\n"
                "点「是」用系统默认程序打开该文件，然后在阅读器里选择「打印」，"
                f"选中实体打印机（推荐「{rec}」）、纸张 A4、纵向，手动确认后再打印。\n\n"
                "点「否」只保留文件，稍后自行打印。")
            if open_it:
                try:
                    os.startfile(path)          # 只打开阅读器，不发送打印任务
                    self.set_status("已打开 A4 文件；请在阅读器里手动确认打印（未自动出纸）")
                    log("兜底：已用系统默认程序打开 A4 文件（未发送打印任务、未出纸）")
                except Exception as e:
                    self._fail("打开文件失败",
                               "无法用系统默认程序打开已生成的文件。\n\n"
                               f"文件已经保存在：\n{path}\n\n"
                               "请手动双击该文件，在阅读器里选择「打印」。", e)
            else:
                self.set_status(f"已生成 A4 文件（未打开、未打印）：{path}")

        def export_pdf(self):
            if not self.pages:
                messagebox.showwarning("提示", "请先打开文件")
                return
            path = filedialog.asksaveasfilename(
                title="保存 A4 PDF", defaultextension=".pdf",
                filetypes=[("PDF", "*.pdf")], initialfile="A3拆分_A4.pdf")
            if not path:
                return
            try:
                self.set_status("正在生成 A4 PDF...")
                export_a4_pdf(self.all_a4(), path)
                self.set_status(f"已生成: {path}")
                messagebox.showinfo("完成", f"A4 文件已保存到：\n{path}")
            except Exception as e:
                self._fail("生成 A4 文件失败",
                           "无法生成 A4 文件。\n\n"
                           "请确认保存位置可写、磁盘空间充足，然后重试。", e)

        def do_print(self):
            """v8 一键直印：不弹通道选择、不问「是/否」，点「打印」就直接打印。

            自动判定提交方式（V4/XPS 或有 842/0x6BE 失败史 → 隐藏 Word 直印；
            其它实体打印机 → GDI），提交后核验队列与 842 事件；
            只有失败才用中文提示并自动保留 A4 PDF，绝不把提交当成成功。
            """
            if not self.pages:
                messagebox.showwarning("提示", "请先打开文件")
                return
            display = self.cmb_printer.get()
            pn = printer_name_from_display(display, self.printer_map)
            if not pn:
                messagebox.showwarning("提示", "未选择打印机")
                return

            block = self._print_block_reason()
            if block:
                log(f"打印被拦截: {pn} 原因={block}")
                self.set_status("打印已禁用：" + block)
                self._auto_fallback(
                    "打印未完成",
                    f"目标打印机「{pn}」当前无法打印：{block}。\n"
                    "本工具没有向打印机发送任务，也没有消耗纸张。")
                return
            self.direct_print_with_verification(pn)

    root = tk.Tk()
    try:
        if "--word-print-probe" in sys.argv:
            # v8 无纸验收：真实 App 类 + 真实 do_print，所有提交/弹窗/打开动作均桩化
            import verify_word_print as _vwp
            _vwp.run_gui_checks(
                sys.modules[__name__], App,
                os.path.join(BASE_DIR, "测试输出", "v8无纸验收_GUI.json"),
                os.path.join(BASE_DIR, "测试样例", "sample_seal_multipage.pdf"))
            try:
                root.destroy()
            except Exception:
                pass
            return
        gui = App(root)
        if "--ui-text-probe" in sys.argv:
            run_ui_text_probe(gui, root)
            try:
                root.destroy()
            except Exception:
                pass
            return
        if "--print-guard-gui" in sys.argv:
            # 独立验收：把 __main__ 命名空间与真实 App 类交给探针，避免模块副本导致桩失效
            import verify_print_guard as _vpg
            _vpg.run_gui_checks(
                sys.modules[__name__], App,
                os.path.join(BASE_DIR, "测试输出", "GUI打印链路模拟.json"),
                os.path.join(BASE_DIR, "测试样例", "sample_seal_multipage.pdf"))
            try:
                root.destroy()
            except Exception:
                pass
            return
        root.mainloop()
    except Exception:
        log(traceback.format_exc())


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        try:
            ok = selftest()
            sys.exit(0 if ok else 1)
        except Exception:
            traceback.print_exc()
            log(traceback.format_exc())
            sys.exit(2)
    else:
        main()
