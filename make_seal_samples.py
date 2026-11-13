# -*- coding: utf-8 -*-
"""
生成「去密封线」专用测试样本：3 页 A3 横向 PDF。
每页：
  - 左半最左侧 8mm  = 红色密封区
  - 右半最右侧 8mm  = 蓝色密封区
  - 左右内容区各有一个 40mm×40mm 绿色正方形（用于检验等比、不拉伸）
  - 内容区有页码文字、中间有中缝红线
"""
import _bootstrap  # 必须在第三方库之前

import os
import fitz

SEAL_MM = 8.0
PT = 72.0 / 25.4                     # 1mm 对应 pt
RED = (0.86, 0.12, 0.12)
BLUE = (0.12, 0.12, 0.86)
GREEN = (0.12, 0.70, 0.12)
BLACK = (0, 0, 0)
WHITE = (1, 1, 1)


def make(path, pages=3):
    W, H = 420 * PT, 297 * PT        # A3 横向 1190.55 × 841.89 pt
    seal = SEAL_MM * PT              # 8mm = 22.68pt -> @200dpi ≈ 63px
    sq = 40 * PT                     # 40mm 正方形 -> @200dpi ≈ 315px
    doc = fitz.open()
    for pg in range(1, pages + 1):
        page = doc.new_page(width=W, height=H)
        sh = page.new_shape()
        # 内容区白底（左右各半）
        sh.draw_rect(fitz.Rect(0, 0, W / 2, H))
        sh.finish(color=WHITE, fill=WHITE)
        sh.draw_rect(fitz.Rect(W / 2, 0, W, H))
        sh.finish(color=WHITE, fill=WHITE)
        # 左密封区（红）、右密封区（蓝）
        sh.draw_rect(fitz.Rect(0, 0, seal, H))
        sh.finish(color=RED, fill=RED)
        sh.draw_rect(fitz.Rect(W - seal, 0, W, H))
        sh.finish(color=BLUE, fill=BLUE)
        # 绿色正方形（左右各一）
        for cx in (W / 4, 3 * W / 4):
            sh.draw_rect(fitz.Rect(cx - sq / 2, H / 2 - sq / 2,
                                   cx + sq / 2, H / 2 + sq / 2))
            sh.finish(color=GREEN, fill=GREEN)
        # 中缝
        sh.draw_line((W / 2, 0), (W / 2, H))
        sh.finish(color=(1, 0, 0), width=1)
        sh.commit()
        page.insert_text((W / 4 - 70, 60), f"PAGE {pg} LEFT", fontsize=16, color=BLACK)
        page.insert_text((3 * W / 4 - 70, 60), f"PAGE {pg} RIGHT", fontsize=16, color=BLACK)
        page.insert_text((seal + 6, H - 30), f"SEAL-L R", fontsize=9, color=BLACK)
        page.insert_text((W - seal - 60, H - 30), f"SEAL-R B", fontsize=9, color=BLACK)
    doc.save(path)
    doc.close()
    return path


def make_all(out_dir):
    os.makedirs(out_dir, exist_ok=True)
    p = os.path.join(out_dir, "sample_seal_multipage.pdf")
    make(p)
    return [p]


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    for p in make_all(os.path.join(here, "测试样例")):
        print("生成:", p, os.path.getsize(p), "bytes")
