# -*- coding: utf-8 -*-
"""
Word COM 直印（v8，2026-11-06）

背景（本机实证）：HP DeskJet 4900 series / PCL-3 (V4) 走 MS_XPS_PROC，
GDI 直接送大位图即便已降到 300dpi + 位图压缩，23:13:24 的作业 7 仍在
23:13:28 与 23:14:44 的打印服务 842 事件里返回 0x6BE，作业被丢弃、未出纸；
而 Word 向同一台实体打印机打印正常。因此本模块提供「隐藏 Word 直印」：

  1) 每张 A4 页生成一个独立临时 Word 文档（PaperSize=A4 / 纵向 / 3mm 安全边距）
  2) 页面图固定放在该文档第 1 页，只打印第 1 页，杜绝多出空白纸
  3) 用隐藏的 Word.Application 设置所选实体打印机与端口
  4) PrintOut(Background=False, Copies=1)
  5) 不打开、不修改任何用户文档；关闭不保存；清理全部临时文件

可注入 word_factory / recorder / temp_dir，便于在完全不出纸的前提下做无纸测试。
所有用户可见文案均为简体中文；英文原始异常只写日志。
"""
import os
import datetime
import traceback

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_PATH = os.path.join(BASE_DIR, "A3SplitTool.log")
TEMP_DIR = os.path.join(BASE_DIR, "测试输出", "Word直印临时")

# 精确 A4 纵向（pt）：210mm × 297mm
A4_W_PT = 210.0 / 25.4 * 72.0        # 595.2756
A4_H_PT = 297.0 / 25.4 * 72.0        # 841.8898

# Word 常量（避免依赖 makepy 生成的常量模块）
WD_PAPER_A4 = 7                       # wdPaperA4
WD_ORIENT_PORTRAIT = 0                # wdOrientPortrait
WD_DO_NOT_SAVE_CHANGES = 0            # wdDoNotSaveChanges
WD_ALERTS_NONE = 0                    # wdAlertsNone
WD_PAGE_BREAK = 7                     # wdPageBreak
WD_FORMAT_DOCX = 16                   # wdFormatDocumentDefault
WD_RELATIVE_HORIZONTAL_PAGE = 1        # wdRelativeHorizontalPositionPage
WD_RELATIVE_VERTICAL_PAGE = 1          # wdRelativeVerticalPositionPage
WD_WRAP_BEHIND = 5                     # wdWrapBehind
SAFE_MARGIN_PT = 3.0 / 25.4 * 72.0    # 3mm，避免图片占满页面后挤出一页空白页

# PrintOut 参数：老板明确要求 Background=False / Copies=1
PRINT_OUT_KWARGS = {"Background": False, "Copies": 1}


def log(msg):
    try:
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")
    except Exception:
        pass


def build_active_printer(printer, port):
    """Word 的 ActivePrinter 需要「打印机名 on 端口名」；无可用端口时只用打印机名。"""
    p = (port or "").strip()
    if not p:
        return printer, ""
    upper = p.upper()
    # 这些不是实体打印机的真实输出端口，拼进去反而会让 Word 找不到设备
    if upper.startswith("PORTPROMPT") or upper.startswith("FILE:") or p == "":
        return printer, ""
    return f"{printer} on {p}", p


def _default_word_factory():
    """独立 Word 实例：DispatchEx 不复用用户已打开的 Word，保证不外显、不碰用户窗口。"""
    import win32com.client
    return win32com.client.DispatchEx("Word.Application")


def _cleanup(paths):
    removed, remaining = [], []
    for p in paths:
        try:
            if os.path.exists(p):
                os.remove(p)
            if os.path.exists(p):
                remaining.append(p)
            else:
                removed.append(p)
        except Exception as e:
            log(f"WORD 临时文件清理失败（仅日志）: {p} {type(e).__name__}: {e}")
            remaining.append(p)
    return removed, remaining


def print_via_word(printer, a4_images, port=None, word_factory=None,
                   temp_dir=None, recorder=None):
    """
    隐藏 Word COM 直印。返回简体中文键名的证据字典；任何异常都向上抛出，
    由调用方的「提交后核验 + 失败自动保留 A4 PDF」链路处理。

    注入 word_factory 时跳过 COM 初始化，便于无纸测试；真实调用使用
    DispatchEx 新建独立实例。绝不调用 Documents.Open，绝不保存用户文档。
    """
    rec = recorder if recorder is not None else {}
    temp_dir = temp_dir or TEMP_DIR
    os.makedirs(temp_dir, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    created = []
    word = doc = None
    opened_docs = []
    com_started = False
    active_printer, used_port = build_active_printer(printer, port)

    rec.update({"提交方式": "Word COM 直印", "打印机": printer, "端口": port or "",
                "页面尺寸pt": [round(A4_W_PT, 4), round(A4_H_PT, 4)],
                "打开用户文档次数": 0, "临时文件": [], "临时文件已清理": False,
                "每页排版": [], "每页PrintOut": []})
    # result 为可变对象：finally 中的收尾结果会反映到调用方拿到的返回值上
    result = {
        "提交方式": "Word COM 直印", "页数": len(a4_images),
        "打印机": printer, "端口": used_port, "ActivePrinter": active_printer,
        "页面": {}, "PrintOut参数": {}, "Word不外显": None,
        "独立实例": "", "未打开用户文档": True,
        "临时文档": [], "临时图片": [],
        "打印机设置证据": rec,
    }

    try:
        if word_factory is None:
            import pythoncom
            pythoncom.CoInitialize()
            com_started = True
            factory = _default_word_factory
            rec["实例来源"] = "DispatchEx(Word.Application) 独立隐藏实例"
        else:
            factory = word_factory
            rec["实例来源"] = "注入的 Word 工厂（无纸测试，不创建真实 COM）"

        word = factory()
        # 隐藏 Word：绝不外显窗口，也抑制一切提示框
        word.Visible = False
        rec["Word.Visible"] = False
        try:
            word.DisplayAlerts = WD_ALERTS_NONE
            rec["Word.DisplayAlerts"] = WD_ALERTS_NONE
        except Exception as e:
            log(f"WORD 设置 DisplayAlerts 失败（仅日志）: {e}")

        # 指定所选实体打印机与端口
        try:
            word.ActivePrinter = active_printer
        except Exception as e:
            log(f"WORD 按端口设置打印机失败，改用打印机名（仅日志）: {e}")
            active_printer = printer
            used_port = ""
            word.ActivePrinter = active_printer
        rec["ActivePrinter"] = active_printer
        rec["已指定端口"] = used_port

        images = []
        body_w = A4_W_PT - 2 * SAFE_MARGIN_PT
        body_h = A4_H_PT - 2 * SAFE_MARGIN_PT
        for i, img in enumerate(a4_images, 1):
            png = os.path.join(temp_dir, f"A4页_{stamp}_{i}.png")
            img.convert("RGB").save(png, "PNG")
            created.append(png)

            # 每一张 A4 使用独立临时文档，只打印第 1 页。这样绝不因整页内联图片
            # 后的段落标记或分页符多出一张空白纸。
            doc = word.Documents.Add()
            opened_docs.append(doc)
            ps = doc.PageSetup
            ps.PaperSize = WD_PAPER_A4
            ps.Orientation = WD_ORIENT_PORTRAIT
            ps.PageWidth = A4_W_PT
            ps.PageHeight = A4_H_PT
            ps.TopMargin = SAFE_MARGIN_PT
            ps.BottomMargin = SAFE_MARGIN_PT
            ps.LeftMargin = SAFE_MARGIN_PT
            ps.RightMargin = SAFE_MARGIN_PT
            ps.Gutter = 0
            ps.HeaderDistance = 0
            ps.FooterDistance = 0

            shape = doc.Shapes.AddPicture(
                png, False, True, SAFE_MARGIN_PT, SAFE_MARGIN_PT, body_w, body_h)
            shape.LockAspectRatio = 0
            shape.Width = body_w
            shape.Height = body_h
            shape.WrapFormat.Type = WD_WRAP_BEHIND
            shape.RelativeHorizontalPosition = WD_RELATIVE_HORIZONTAL_PAGE
            shape.RelativeVerticalPosition = WD_RELATIVE_VERTICAL_PAGE
            shape.Left = SAFE_MARGIN_PT
            shape.Top = SAFE_MARGIN_PT

            doc.Repaginate()
            page_count = int(doc.ComputeStatistics(2))  # wdStatisticPages
            layout = {
                "页序": i, "临时图片": png,
                "纸张": "A4", "方向": "纵向",
                "页面pt": [round(A4_W_PT, 4), round(A4_H_PT, 4)],
                "图片宽pt": round(body_w, 4), "图片高pt": round(body_h, 4),
                "安全边距pt": round(SAFE_MARGIN_PT, 4),
                "排版总页数": page_count, "实际打印页": "1",
            }
            rec["每页排版"].append(layout)
            images.append(layout)

            kwargs = dict(PRINT_OUT_KWARGS)
            kwargs["Pages"] = "1"
            log(f"WORD 直印提交: 打印机={printer} 端口={used_port or '无'} "
                f"ActivePrinter={active_printer} 第{i}页/"
                f"文档总页数={page_count} A4纵向={A4_W_PT:.2f}x{A4_H_PT:.2f}pt "
                f"PrintOut={kwargs}")
            doc.PrintOut(**kwargs)
            rec["每页PrintOut"].append(dict(kwargs))
            doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
            opened_docs.pop()
            doc = None

        rec["嵌入A4页面"] = images
        rec["页数"] = len(images)
        rec["PrintOut参数"] = dict(PRINT_OUT_KWARGS)
        rec["提交完成"] = True
        result.update({
            "页面": {"纸张": "A4", "方向": "纵向",
                     "宽pt": round(A4_W_PT, 4), "高pt": round(A4_H_PT, 4),
                     "安全边距pt": round(SAFE_MARGIN_PT, 4)},
            "PrintOut参数": dict(PRINT_OUT_KWARGS),
            "Word不外显": not bool(word.Visible),
            "独立实例": rec.get("实例来源", ""),
            "页数": len(images), "端口": used_port,
            "ActivePrinter": active_printer,
            "临时图片": [x["临时图片"] for x in images],
            "每页排版": list(images),
            "每页PrintOut": list(rec["每页PrintOut"]),
        })
        return result
    except Exception as e:
        rec["提交完成"] = False
        log("WORD 直印异常（原始信息仅日志）: " + repr(e) + " | " + traceback.format_exc())
        raise
    finally:
        # 关闭不保存：文档与 Word 实例都不保存任何更改
        for pending_doc in list(opened_docs):
            try:
                pending_doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
                rec["文档关闭"] = "Close(SaveChanges=0) 不保存"
            except Exception as e:
                rec["文档关闭"] = "关闭失败（仅日志）"
                log(f"WORD 关闭残留文档失败（仅日志）: {e}")
            if pending_doc in opened_docs:
                opened_docs.remove(pending_doc)
        if word is not None:
            try:
                word.Quit(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
                rec["Word退出"] = "Quit(SaveChanges=0) 不保存"
            except Exception as e:
                rec["Word退出"] = "退出失败（仅日志）"
                log(f"WORD 退出失败（仅日志）: {e}")
        removed, remaining = _cleanup(created)
        rec["临时文件"] = list(created)
        rec["已删除临时文件"] = removed
        rec["残留临时文件"] = remaining
        rec["临时文件已清理"] = (len(remaining) == 0)
        try:
            rec["打开用户文档次数"] = int(getattr(word, "打开用户文档次数", 0) or 0) \
                if word is not None else 0
        except Exception:
            rec["打开用户文档次数"] = 0
        result["关闭"] = {"文档": rec.get("文档关闭"), "Word": rec.get("Word退出")}
        result["临时文件清理"] = {"创建": len(created), "已删除": len(removed),
                                "残留": remaining, "全部清理": len(remaining) == 0}
        result["未打开用户文档"] = (rec.get("打开用户文档次数", 0) == 0)
        if com_started:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass
        log(f"WORD 收尾: 文档={rec.get('文档关闭')} Word={rec.get('Word退出')} "
            f"临时文件={len(created)} 已删={len(removed)} 残留={len(remaining)}")
