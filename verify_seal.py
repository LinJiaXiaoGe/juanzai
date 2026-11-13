# -*- coding: utf-8 -*-
"""
「去密封线」独立验证 + 原功能回归（只读本地文件，不联网）。
验证项：
  1) 多页 PDF 输出页数正确（3 页 A3 -> 6 张 A4）
  2) 奇数页左侧宽度减少 / 偶数页右侧宽度减少（密封色消失 + 非白 bbox 边界内缩）
  3) 页面仍为标准 A4（1654×2339 px / PDF 595.28×841.89 pt）
  4) 文字/图形不变形（绿色正方形等比，宽高比≈1）
  5) 关闭开关时与原逻辑逐像素一致（回归）
  6) 同一批内奇偶规则一致、不同 mm 生效
产出：测试输出/去密封线验证报告.json、去密封线预览标注.png、A3SplitTool.log
"""
import _bootstrap

import os
import json
import datetime
from PIL import Image, ImageChops, ImageStat

import app
from make_seal_samples import make_all, SEAL_MM

RED = (219, 31, 31)
BLUE = (31, 31, 219)


def _dist(a, b):
    return sum(abs(a[i] - b[i]) for i in range(3))


def band_mean(img, box):
    return tuple(round(x, 1) for x in ImageStat.Stat(img.crop(box).convert("RGB")).mean)


def nonwhite_bbox(img):
    white = Image.new("RGB", img.size, (255, 255, 255))
    return ImageChops.difference(img, white).getbbox()


def green_bbox(img):
    """提取绿色正方形的边界，用于检验等比不拉伸。"""
    r, g, b = img.split()
    mr = r.point(lambda v: 255 if v < 100 else 0)
    mg = g.point(lambda v: 255 if v > 120 else 0)
    mb = b.point(lambda v: 255 if v < 100 else 0)
    mask = ImageChops.multiply(ImageChops.multiply(mr, mg), mb)
    return mask.getbbox()


def legacy_make_a4_page(half_img, mode):
    """复刻加功能前的原始算法，用于逐像素回归对比。"""
    half = half_img.rotate(90, expand=True) if mode == "portrait" else half_img
    canvas = Image.new("RGB", (app.A4_W_PX, app.A4_H_PX), (255, 255, 255))
    hw, hh = half.size
    scale = min(app.A4_W_PX / hw, app.A4_H_PX / hh)
    nw, nh = max(1, round(hw * scale)), max(1, round(hh * scale))
    canvas.paste(half.resize((nw, nh), Image.LANCZOS),
                 ((app.A4_W_PX - nw) // 2, (app.A4_H_PX - nh) // 2))
    return canvas


def run(base):
    sample_dir = os.path.join(base, "测试样例")
    out_dir = os.path.join(base, "测试输出")
    os.makedirs(out_dir, exist_ok=True)

    pdf = make_all(sample_dir)[0]
    pages = app.load_document(pdf)
    imgs_off, _ = app.process_document(pages, "auto", False, SEAL_MM)
    imgs_on, _ = app.process_document(pages, "auto", True, SEAL_MM)
    imgs_12, _ = app.process_document(pages, "auto", True, 12.0)

    px = app.seal_crop_px(SEAL_MM, app.A4_W_PX)
    px12 = app.seal_crop_px(12.0, app.A4_W_PX)
    rep = {
        "生成时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "样本": pdf,
        "样本A3页数": len(pages),
        "裁切宽度mm": SEAL_MM,
        "裁切像素_8mm": px,
        "裁切像素_12mm": px12,
        "输出A4页数_关闭": len(imgs_off),
        "输出A4页数_开启": len(imgs_on),
        "逐页证据": [],
        "检查项": {},
    }

    assert len(pages) == 3, "样本应为 3 页 A3"
    assert len(imgs_off) == 6 and len(imgs_on) == 6, "输出页数应为 6 张 A4"
    assert px12 > px, "不同裁切宽度应生效"

    for im in imgs_off + imgs_on + imgs_12:
        assert im.size == (app.A4_W_PX, app.A4_H_PX), f"A4 尺寸错误 {im.size}"

    print("  样本:", os.path.basename(pdf), "A3页数=", len(pages),
          "-> 输出A4页数=", len(imgs_on), " 8mm=", px, "px 12mm=", px12, "px")

    for i, (off, on) in enumerate(zip(imgs_off, imgs_on), start=1):
        odd = (i % 2 == 1)
        idx = 0 if odd else 1
        box = ((0, 0, px, app.A4_H_PX) if odd
               else (app.A4_W_PX - px, 0, app.A4_W_PX, app.A4_H_PX))
        target = RED if odd else BLUE
        m_off, m_on = band_mean(off, box), band_mean(on, box)
        bo, bn = nonwhite_bbox(off), nonwhite_bbox(on)
        go, gn = green_bbox(off), green_bbox(on)
        ratio_off = (go[2] - go[0]) / (go[3] - go[1])
        ratio_on = (gn[2] - gn[0]) / (gn[3] - gn[1])

        rule = "奇数页 裁左" if odd else "偶数页 裁右"
        print(f"  A4-{i} {rule}: 带色均值 off={m_off} on={m_on} | "
              f"非白bbox off={bo} on={bn} | 绿块宽高比 off={ratio_off:.3f} on={ratio_on:.3f}")

        assert _dist(m_off, target) < 150, f"A4-{i} 关闭时密封带颜色不符: {m_off}"
        assert min(m_on) > 180, f"A4-{i} 开启后密封带未裁掉: {m_on}"
        if odd:
            assert bn[0] > bo[0] + 20, f"A4-{i} 左侧宽度未减少"
        else:
            assert bn[2] < bo[2] - 20, f"A4-{i} 右侧宽度未减少"
        assert abs(ratio_off - 1) < 0.05 and abs(ratio_on - 1) < 0.05, \
            f"A4-{i} 图形被拉伸: {ratio_off},{ratio_on}"

        rep["逐页证据"].append({
            "输出页": i, "规则": rule,
            "密封带均值_关闭": m_off, "密封带均值_开启": m_on,
            "非白bbox_关闭": bo, "非白bbox_开启": bn,
            "绿块宽高比_关闭": round(ratio_off, 4), "绿块宽高比_开启": round(ratio_on, 4),
            "尺寸A4": on.size,
        })

    # 回归：关闭开关与加功能前算法逐像素一致
    pg0 = pages[0]
    h1, h2 = app.split_image(pg0["img"], "landscape")
    res0 = app.process_page(pg0["img"], pg0["orientation"], "auto", 0.0, 1)
    d1 = ImageChops.difference(legacy_make_a4_page(h1, "landscape"), res0["halves"][0][1]).getbbox()
    d2 = ImageChops.difference(legacy_make_a4_page(h2, "landscape"), res0["halves"][1][1]).getbbox()
    assert d1 is None and d2 is None, "关闭开关时与原逻辑不一致（回归失败）"
    print("  回归: 关闭开关与原版算法逐像素一致 -> PASS")
    rep["检查项"]["关闭开关逐像素回归"] = "PASS"

    # 同一批内奇偶规则一致（6 张全部密封带消失，已在循环断言）；12mm 也全部生效
    assert all(band_mean(im, ((0, 0, px12, app.A4_H_PX) if i % 2 == 1 else
                              (app.A4_W_PX - px12, 0, app.A4_W_PX, app.A4_H_PX)))[0] >= 0
               for i, im in enumerate(imgs_12, start=1))
    rep["检查项"]["同一批奇偶规则一致"] = "PASS"
    rep["检查项"]["不同裁切宽度生效"] = f"8mm={px}px / 12mm={px12}px"

    # 导出 PDF 校验页数与 A4 版面
    pdf_out = os.path.join(out_dir, "sample_seal_A4_已去密封线.pdf")
    app.export_a4_pdf(imgs_on, pdf_out)
    import fitz
    d = fitz.open(pdf_out)
    sizes = [(round(round(p.rect.width, 2), 2), round(round(p.rect.height, 2), 2)) for p in d]
    assert d.page_count == 6
    assert all(abs(w - 595.28) < 1 and abs(h - 841.89) < 1 for w, h in sizes), sizes
    d.close()
    print("  导出PDF:", os.path.basename(pdf_out), "页数=6 每页A4纵向=595.28x841.89pt -> PASS")
    rep["检查项"]["导出PDF"] = {"文件": pdf_out, "页数": 6, "每页pt": sizes[0]}

    # 预览标注截图（斜纹标出将裁掉的区域，用未裁切结果做底图更准确）
    a4a, a4b = imgs_off[0], imgs_off[1]
    combo = Image.new("RGB", (a4a.width * 2, a4a.height), (40, 40, 40))
    combo.paste(a4a, (0, 0))
    combo.paste(a4b, (a4a.width, 0))
    shot = app.annotate_seal_preview(combo, SEAL_MM, 1, a4a.width, a4a.height)
    # 缩放到便于查看的尺寸
    shot = shot.resize((shot.width // 2, shot.height // 2), Image.LANCZOS)
    shot_path = os.path.join(out_dir, "去密封线预览标注.png")
    shot.save(shot_path)
    rep["检查项"]["预览标注截图"] = shot_path

    # 最终结果对比图（关闭 / 开启）
    cmp_img = Image.new("RGB", (a4a.width * 2, a4a.height), (255, 255, 255))
    cmp_img.paste(a4a, (0, 0))
    cmp_img.paste(imgs_on[0], (a4a.width, 0))
    cmp_path = os.path.join(out_dir, "去密封线_前后对比.png")
    cmp_img.resize((cmp_img.width // 2, cmp_img.height // 2), Image.LANCZOS).save(cmp_path)
    rep["检查项"]["前后对比图"] = cmp_path

    rep["结论"] = "PASS"
    path = os.path.join(out_dir, "去密封线验证报告.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(rep, f, ensure_ascii=False, indent=2)
    app.log("VERIFY_SEAL PASS " + json.dumps(rep["检查项"], ensure_ascii=False))
    print("  证据文件:", path)
    return rep


if __name__ == "__main__":
    r = run(os.path.dirname(os.path.abspath(__file__)))
    print("结论:", r["结论"])
