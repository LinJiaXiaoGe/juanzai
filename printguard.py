# -*- coding: utf-8 -*-
"""
打印结果核验与可靠回退（v7，2026-11-06）

背景（本机实证）：HP DeskJet 4900 series / PCL-3 (V4) 驱动使用系统 XPS 渲染通道，
GDI 直接提交的大位图会让打印处理器返回 0x6BE，作业被后台处理程序丢弃，
StartDoc / EndDoc 却全部返回成功——所以「StartDoc 成功」绝不能当作「已打印」。

本模块提供：
  1) 送驱动前的位图降采样（避免大图压垮驱动）
  2) 压缩 A4 PDF（JPEG + deflate，比 PNG 版小一个数量级）
  3) 只读的打印队列 / 打印服务 842 事件核验（能抓到 0x6BE 这类异步失败）
  4) V4 / XPS / 类驱动识别 + 失败历史，按品牌无关的证据推荐更稳妥的 PDF 通道
  5) 失败时自动保留合并 PDF + 两张单页 A4 PDF，并提供打开文件夹

所有用户可见文案均为简体中文；英文原始异常/事件字段只写入 A3SplitTool.log。
核验函数支持注入 jobs_provider / events_provider，可在完全不出纸的前提下自动化测试。
"""
import io
import os
import re
import json
import time
import datetime
import traceback

from PIL import Image

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_PATH = os.path.join(BASE_DIR, "A3SplitTool.log")
FALLBACK_DIR = os.path.join(BASE_DIR, "测试输出", "打印失败回退")
HISTORY_FILE = os.path.join(BASE_DIR, "测试输出", "打印通道历史.json")
EVENT_CHANNEL = "Microsoft-Windows-PrintService/Operational"

PRINT_DC_DPI = 300               # 送驱动时把 DC 分辨率从 600 降到 300
PRINT_MAX_LONG_SIDE = 1754       # A4 长边 150dpi，约为原 200dpi 渲染的 75%
PRINT_JPEG_QUALITY = 80
VERIFY_WAIT_SECONDS = 8.0
VERIFY_POLL_SECONDS = 0.5

A4_W_PT = 210.0 / 25.4 * 72
A4_H_PT = 297.0 / 25.4 * 72

# win32print 作业状态位
JOB_STATUS_ERROR = 0x00000002
JOB_STATUS_SPOOLING = 0x00000008
JOB_STATUS_PRINTING = 0x00000010
JOB_STATUS_PRINTED = 0x00000080
JOB_STATUS_DELETED = 0x00000100
JOB_STATUS_RETAINING = 0x00002000


def log(msg):
    try:
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")
    except Exception:
        pass


# ---------- 1. 避免大图压垮驱动 ----------
def compress_for_print(img, max_long_side=PRINT_MAX_LONG_SIDE):
    """送 GDI 之前压小位图并统一成 RGB；返回 (新图, 统计)。"""
    src = img.convert("RGB")
    ow, oh = src.size
    if max(ow, oh) <= max_long_side:
        small = src.copy()
    else:
        r = max_long_side / float(max(ow, oh))
        small = src.resize((max(1, round(ow * r)), max(1, round(oh * r))),
                           Image.LANCZOS)
    before = ow * oh * 3
    after = small.size[0] * small.size[1] * 3
    return small, {
        "原尺寸": [ow, oh], "送驱动尺寸": list(small.size),
        "原始字节": before, "送驱动字节": after,
        "缩减比例": round(1.0 - after / float(before or 1), 4),
    }


# ---------- 2. 压缩 A4 PDF 通道 ----------
def export_compressed_a4_pdf(a4_images, out_path, quality=PRINT_JPEG_QUALITY):
    """每页用 JPEG 流嵌入，保持标准 A4 纵向；文件远小于 PNG 版本。"""
    import fitz
    doc = fitz.open()
    try:
        for img in a4_images:
            buf = io.BytesIO()
            img.convert("RGB").save(buf, format="JPEG", quality=quality,
                                    optimize=True, progressive=False)
            page = doc.new_page(width=A4_W_PT, height=A4_H_PT)
            page.insert_image(page.rect, stream=buf.getvalue())
        os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
        doc.save(out_path, garbage=4, deflate=True, deflate_images=True, clean=True)
    finally:
        doc.close()
    return out_path


def _pdf_facts(path):
    import fitz
    d = fitz.open(path)
    try:
        sizes = [(round(p.rect.width, 2), round(p.rect.height, 2)) for p in d]
        return {"页数": d.page_count, "每页尺寸pt": sizes,
                "字节": os.path.getsize(path)}
    finally:
        d.close()


def save_a4_fallback(a4_images, folder=None, prefix="A3拆分_A4"):
    """失败时自动保留：一个合并 PDF + 每张 A4 各一个单页 PDF。"""
    folder = folder or FALLBACK_DIR
    os.makedirs(folder, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    merged = os.path.join(folder, f"{prefix}_回退_{stamp}.pdf")
    export_compressed_a4_pdf(a4_images, merged)
    singles = []
    for i, img in enumerate(a4_images, 1):
        p = os.path.join(folder, f"{prefix}_第{i}页_{stamp}.pdf")
        export_compressed_a4_pdf([img], p)
        singles.append(p)
    files = [{"路径": p, "角色": role, **_pdf_facts(p)}
             for p, role in [(merged, "全部页合并")] +
             [(p, f"第{i}页单页") for i, p in enumerate(singles, 1)]]
    log(f"PRINTGUARD 失败回退文件已生成 文件夹={folder} 页数={len(a4_images)} "
        f"文件数={len(files)} 总字节={sum(f['字节'] for f in files)}")
    return {"文件夹": folder, "合并PDF": merged, "单页PDF": singles,
            "页数": len(a4_images), "文件": files}


def open_folder(folder):
    """只打开文件夹，不发送任何打印任务。"""
    os.startfile(folder)
    return True


# ---------- 3. 只读核验：打印队列 + 打印服务事件 ----------
def snapshot_jobs(printer):
    """只读枚举队列；返回简体中文键名的作业列表。"""
    import win32print
    h = win32print.OpenPrinter(printer)
    try:
        jobs = win32print.EnumJobs(h, 0, 999, 2) or []
    finally:
        win32print.ClosePrinter(h)
    out = []
    for j in jobs:
        st = int(j.get("Status") or 0)
        out.append({
            "作业号": int(j.get("JobId") or 0),
            "状态位": st, "状态文本": str(j.get("pStatus") or ""),
            "文档名": str(j.get("pDocument") or ""),
            "打印处理器": str(j.get("pPrintProcessor") or ""),
            "数据类型": str(j.get("pDatatype") or ""),
            "总页数": int(j.get("TotalPages") or 0),
            "已打页数": int(j.get("PagesPrinted") or 0),
        })
    return out


def _xml_field(xml, tag):
    m = re.search(r"<%s>(.*?)</%s>" % (tag, tag), xml, re.S | re.I)
    return m.group(1).strip() if m else ""


def _parse_event_time(xml):
    m = re.search(r"SystemTime='([^']+)'", xml)
    if not m:
        return None
    try:
        return datetime.datetime.strptime(m.group(1)[:19], "%Y-%m-%dT%H:%M:%S")
    except Exception:
        return None


def read_print_events(since_utc=None, printer=None, max_scan=200):
    """只读读取 PrintService/Operational 事件，重点抓 842（打印处理器返回错误码）。"""
    import win32evtlog
    events, err = [], None
    try:
        q = win32evtlog.EvtQuery(EVENT_CHANNEL, win32evtlog.EvtQueryReverseDirection,
                                 "*", None)
        raw = win32evtlog.EvtNext(q, max_scan)
        for handle in raw:
            xml = win32evtlog.EvtRender(handle, win32evtlog.EvtRenderEventXml)
            t = _parse_event_time(xml)
            if since_utc and (t is None or t < since_utc):
                continue
            ev_printer = _xml_field(xml, "Printer")
            if printer and ev_printer and printer.lower() not in ev_printer.lower():
                continue
            code = _xml_field(xml, "ErrorCode").lower()
            bad = code not in ("", "0", "0x0", "0x00000000")
            jid = _xml_field(xml, "JobId")
            events.append({
                "事件ID": _xml_field(xml, "EventID"),
                "记录号": _xml_field(xml, "EventRecordID"),
                "时间UTC": t.strftime("%Y-%m-%d %H:%M:%S") if t else "",
                "_时间": t,
                "作业号": int(jid) if jid.isdigit() else None,
                "打印机": ev_printer,
                "驱动": _xml_field(xml, "Driver"),
                "打印处理器": _xml_field(xml, "Processor"),
                "错误码": code.upper() if code else "",
                "是否错误": bad,
            })
    except Exception as e:
        err = f"{type(e).__name__}: {e}"
        log("PRINTGUARD 读取打印服务事件失败（仅日志）: " + err)
    return events, err


def _public_event(e):
    return {k: v for k, v in e.items() if not k.startswith("_")}


def verify_print_job(printer, before_ids=(), started_utc=None,
                     wait_seconds=VERIFY_WAIT_SECONDS, poll=VERIFY_POLL_SECONDS,
                     jobs_provider=None, events_provider=None):
    """
    打印后只读追踪队列与事件，判断作业是「失败 / 已离开队列 / 仍待确认」。
    jobs_provider / events_provider 可注入；注入时 wait_seconds 建议设 0，不出纸。
    """
    started_utc = started_utc or datetime.datetime.now(datetime.timezone.utc).replace(tzinfo=None)
    jobs_fn = jobs_provider or (lambda: snapshot_jobs(printer))
    events_fn = events_provider or (lambda: read_print_events(started_utc, printer)[0])
    before = set(before_ids or ())
    deadline = time.time() + wait_seconds
    new_ids, errors, last_jobs, event_errors = set(), [], [], []
    note = ""

    while True:
        try:
            last_jobs = jobs_fn() or []
            for j in last_jobs:
                jid = j.get("作业号")
                if jid and jid not in before:
                    new_ids.add(jid)
                st = int(j.get("状态位") or 0)
                if st & JOB_STATUS_ERROR:
                    errors.append({"来源": "打印队列", "作业号": jid,
                                   "状态文本": j.get("状态文本", "")})
        except Exception as e:
            note = "读取打印队列受限（仅日志）：" + str(e)
            log("PRINTGUARD " + note)

        try:
            evs = events_fn() or []
            for e in evs:
                t = e.get("_时间")
                if t and t < started_utc:
                    continue
                if e.get("打印机") and printer.lower() not in e["打印机"].lower():
                    continue
                if e.get("作业号"):
                    new_ids.add(e["作业号"])
                if e.get("是否错误"):
                    event_errors.append(_public_event(e))
        except Exception as e:
            log("PRINTGUARD 读取打印服务事件异常（仅日志）: " + str(e))

        if errors or event_errors:
            break
        if new_ids and not (set(j.get("作业号") for j in last_jobs) & new_ids) \
                and time.time() - (deadline - wait_seconds) >= 0.5:
            break
        if time.time() >= deadline:
            break
        time.sleep(poll)

    codes = sorted(set(e["错误码"] for e in event_errors if e.get("错误码")))
    xps = any("XPS" in (e.get("打印处理器") or "").upper() for e in event_errors)
    event842 = [e for e in event_errors if str(e.get("事件ID")) == "842"]
    if errors or event_errors:
        cn = "系统打印服务报告打印作业失败，页面内容没有到达打印机。"
        if event842:
            cn += " 已核验打印队列与打印服务 842 事件，本次作业被报告为失败。"
        if xps:
            cn += " 当前驱动的文档渲染通道处理本次作业失败"
            if codes:
                cn += "（错误代码 " + "、".join(codes) + "）"
            cn += "，通常与驱动处理大尺寸图像有关。"
        if errors:
            cn += " 队列中该作业也被标记为错误。"
        conclusion = "失败"
    elif new_ids and not (set(j.get("作业号") for j in last_jobs) & new_ids):
        conclusion = "未检出错误"
        cn = ("打印作业已离开系统队列，且未在队列或打印服务事件中检出错误。"
              "本工具无法确认纸张是否真正输出，请检查打印机出纸口。")
    elif new_ids:
        conclusion = "待确认"
        cn = ("打印作业仍在系统队列中，暂未检出错误，但也无法确认是否已经出纸。"
              "请稍后查看打印机状态；本工具不会自动重试，以免重复消耗纸张。")
    else:
        conclusion = "待确认"
        cn = ("未在打印队列中捕获到本次作业编号，无法确认打印结果。"
              "请检查打印机是否联机，或使用下面保留的 A4 文件打印。")

    result = {
        "结论": conclusion, "中文说明": cn,
        "打印机": printer, "核验起始UTC": started_utc.strftime("%Y-%m-%d %H:%M:%S"),
        "新作业号": sorted(new_ids), "队列前作业号": sorted(before),
        "队列错误": errors, "事件错误": event_errors,
        "错误代码": codes, "842事件数": len(event842),
        "最终队列快照": last_jobs,
        "队列读取备注": note,
    }
    log("PRINTGUARD 核验结论=" + conclusion + " 作业号=" + str(sorted(new_ids)) +
        " 事件错误数=" + str(len(event_errors)) + " 错误代码=" + str(codes) +
        " 中文说明=" + cn)
    return result


# ---------- 4. 品牌无关的通道推荐 ----------
# 只看驱动特征与历史失败证据，不写死任何品牌
V4_DRIVER_HINTS = ("(v4)", " v4 ", "class driver", "xps", "microsoft ipp",
                   "microsoft pcl6", "v4 driver")


def _load_history():
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def record_channel_result(printer, conclusion, chinese_note, extra=None):
    """记录每台打印机的直打结果，供下次自动选通道（与品牌无关）。"""
    data = _load_history()
    rec = data.get(printer, {})
    rec.update({"上次核验时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "上次结论": conclusion, "上次说明": chinese_note,
                "推荐通道": "pdf" if conclusion == "失败" else "gdi"})
    if extra:
        rec.update(extra)
    data[printer] = rec
    try:
        os.makedirs(os.path.dirname(HISTORY_FILE), exist_ok=True)
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log("PRINTGUARD 写通道历史失败（仅日志）: " + str(e))
    return rec


def recommended_channel(printer, facts=None):
    """返回 (通道, 简体中文原因)：pdf = 压缩 PDF 文件通道；gdi = 直接送驱动。"""
    facts = facts or {}
    driver = (facts.get("驱动") or "").lower()
    hist = _load_history().get(printer, {})
    reasons = []
    if hist.get("推荐通道") == "pdf":
        reasons.append("这台打印机上次直接打印失败，已自动改走更稳妥的文件通道")
    for k in V4_DRIVER_HINTS:
        if k in driver:
            reasons.append("当前驱动使用系统文档渲染通道，处理大尺寸图像容易失败")
            break
    if not reasons:
        return "gdi", "常规驱动，可直接发送打印任务"
    return "pdf", "；".join(reasons)
