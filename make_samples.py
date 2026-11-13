# -*- coding: utf-8 -*-
"""
生成 A3 测试样本：
  - sample_a3_landscape.pdf : A3 横向(420x297mm)，2 页，每页左右各标注 左/右 + 页码
  - sample_a3_portrait.jpg  : A3 纵向(297x420mm)，1 页，上下各标注 上/下
供 selftest 与实际验证使用。
"""
import _bootstrap  # 必须在第三方库之前，注入工具自带 lib
import os
from PIL import Image, ImageDraw, ImageFont


def _font(size):
    try:
        return ImageFont.truetype("simhei.ttf", size)
    except Exception:
        try:
            return ImageFont.truetype("C:/Windows/Fonts/msyh.ttc", size)
        except Exception:
            return ImageFont.load_default()


def make_landscape_pdf(path, dpi=200):
    import fitz
    # A3 横向 pt: 1191 x 842
    W, H = 1191, 842
    doc = fitz.open()
    for pg in range(1, 3):
        page = doc.new_page(width=W, height=H)
        shape = page.new_shape()
        # 左半
        shape.draw_rect(fitz.Rect(0, 0, W / 2, H))
        shape.finish(color=(0.2, 0.2, 0.2), fill=(0.85, 0.9, 1.0))
        # 右半
        shape.draw_rect(fitz.Rect(W / 2, 0, W, H))
        shape.finish(color=(0.2, 0.2, 0.2), fill=(1.0, 0.92, 0.8))
        # 中缝
        shape.draw_line((W / 2, 0), (W / 2, H))
        shape.finish(color=(1, 0, 0), width=2)
        shape.commit()
        page.insert_text((W / 4 - 90, H / 2 - 40), f"左 LEFT  P{pg}", fontsize=40, color=(0, 0, 0))
        page.insert_text((3 * W / 4 - 90, H / 2 - 40), f"右 RIGHT  P{pg}", fontsize=40, color=(0, 0, 0))
        page.insert_text((20, 30), f"A3横向 第{pg}页 左半在左/右半在右", fontsize=14, color=(0, 0, 0))
    doc.save(path)
    doc.close()


def make_portrait_jpg(path, dpi=200):
    W = round(297 / 25.4 * dpi)   # 2339
    H = round(420 / 25.4 * dpi)   # 3307
    img = Image.new("RGB", (W, H), (255, 255, 255))
    d = ImageDraw.Draw(img)
    # 上半
    d.rectangle([0, 0, W, H // 2], fill=(200, 230, 255))
    # 下半
    d.rectangle([0, H // 2, W, H], fill=(255, 230, 200))
    # 中缝
    d.line([(0, H // 2), (W, H // 2)], fill=(255, 0, 0), width=max(3, W // 400))
    f1 = _font(W // 14)
    d.text((W // 2 - W // 8, H // 4), "上 TOP", font=f1, fill=(0, 0, 0))
    d.text((W // 2 - W // 8, 3 * H // 4), "下 BOTTOM", font=f1, fill=(0, 0, 0))
    d.text((20, 20), "A3纵向 上半在上方/下半在下方", font=_font(W // 28), fill=(0, 0, 0))
    img.save(path, "JPEG", quality=92)


def make_all(out_dir):
    os.makedirs(out_dir, exist_ok=True)
    pdf = os.path.join(out_dir, "sample_a3_landscape.pdf")
    jpg = os.path.join(out_dir, "sample_a3_portrait.jpg")
    make_landscape_pdf(pdf)
    make_portrait_jpg(jpg)
    return [pdf, jpg]


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    out = os.path.join(here, "测试样例")
    for p in make_all(out):
        print("生成:", p, os.path.getsize(p), "bytes")
