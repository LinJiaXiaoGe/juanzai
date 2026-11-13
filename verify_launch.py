# -*- coding: utf-8 -*-
"""
独立验收：通过【桌面快捷方式】实际启动 GUI 并核验，不依赖 WorkBuddy 私有 Python。
核验点：
  1) .lnk 指向 启动工具.bat 且工作目录正确
  2) os.startfile(.lnk) 等同双击，能拉起 pythonw.exe 且命令行含 app.py
  3) GUI 完整初始化并写下日志（含 sys.executable 与依赖来源 uses_tool_lib=True）
  4) 结束后关闭 GUI 进程，输出 测试输出/验证报告_快捷方式启动.json
"""
import _bootstrap

import os
import sys
import json
import time
import subprocess
import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
LNK = os.path.join(os.environ["USERPROFILE"], "Desktop", "A3试卷拆分打印工具.lnk")
LOG = os.path.join(BASE, "A3SplitTool.log")
OUT = os.path.join(BASE, "测试输出")
os.makedirs(OUT, exist_ok=True)

PS_CMD = """Get-CimInstance Win32_Process -Filter "Name = 'pythonw.exe'" | Select-Object ProcessId,CommandLine | ConvertTo-Json -Compress"""


def gui_processes():
    try:
        raw = subprocess.check_output(["powershell", "-NoProfile", "-Command", PS_CMD],
                                      timeout=20)
        # PowerShell may emit Chinese paths in the OEM code page; decode with
        # replacement so a non-UTF8 byte never breaks the JSON parsing.
        text = raw.decode("utf-8", errors="replace").strip() if raw else ""
        data = json.loads(text) if text else []
        if isinstance(data, dict):
            data = [data]
        return [d for d in data if "app.py" in (d.get("CommandLine") or "")]
    except Exception as e:
        print("  进程查询异常:", e)
        return []


def kill(pid):
    subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)


def main():
    print("=" * 70)
    print("桌面快捷方式启动 GUI — 独立验收")
    print("=" * 70)
    assert os.path.exists(LNK), f"桌面快捷方式不存在: {LNK}"

    # 0) 清理此前残留的本工具 GUI 进程
    for d in gui_processes():
        print("  清理残留进程 PID=", d["ProcessId"])
        kill(d["ProcessId"])

    # 1) 读取 .lnk 属性
    import win32com.client
    ws = win32com.client.Dispatch("WScript.Shell")
    sc = ws.CreateShortcut(LNK)
    print("[快捷方式]")
    print("  路径              =", LNK)
    print("  TargetPath        =", sc.TargetPath)
    print("  WorkingDirectory  =", sc.WorkingDirectory)
    print("  IconLocation      =", sc.IconLocation)
    assert sc.TargetPath.lower().endswith("启动工具.bat"), "快捷方式目标不是启动工具.bat"
    assert os.path.isdir(sc.WorkingDirectory), "工作目录无效"

    # 2) 记录启动前日志大小
    before = os.path.getsize(LOG) if os.path.exists(LOG) else 0

    # 3) 等同"双击快捷方式"
    print("\n[启动] os.startfile(桌面快捷方式) ...")
    os.startfile(LNK)

    # 4) 轮询：GUI 日志 + 进程
    gui_log = None
    procs = []
    deadline = time.time() + 25
    while time.time() < deadline:
        if os.path.exists(LOG) and os.path.getsize(LOG) > before:
            with open(LOG, "r", encoding="utf-8") as f:
                f.seek(before)
                new = f.read()
            if "GUI 初始化完成" in new:
                gui_log = new.strip()
        procs = gui_processes()
        if gui_log and procs:
            break
        time.sleep(0.5)

    print("\n[核验结果]")
    print("  GUI 完整初始化日志 =", "已写入" if gui_log else "未写入")
    if gui_log:
        for line in gui_log.splitlines():
            print("    ", line)
    print("  GUI 进程数 =", len(procs))
    for d in procs:
        print("    PID=", d["ProcessId"], " CMD=", d["CommandLine"])

    uses_lib = bool(gui_log) and ("uses_tool_lib': True" in gui_log or 'uses_tool_lib": True' in gui_log)
    is_system_py = bool(gui_log) and ("C:\\Python314\\pythonw.exe" in gui_log)
    no_workbuddy = bool(gui_log) and ("workbuddy" not in gui_log.lower())
    print("  依赖来自工具目录 lib =", uses_lib)
    print("  使用系统 Python 3.14 =", is_system_py)
    print("  未依赖 WorkBuddy 私有 Python =", no_workbuddy)

    ok = bool(gui_log) and bool(procs) and uses_lib and is_system_py and no_workbuddy

    report = {
        "生成时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "快捷方式": LNK,
        "TargetPath": sc.TargetPath,
        "WorkingDirectory": sc.WorkingDirectory,
        "GUI启动日志": gui_log,
        "GUI进程": procs,
        "依赖来自工具目录lib": uses_lib,
        "使用系统Python314": is_system_py,
        "未依赖WorkBuddy私有Python": no_workbuddy,
        "结论": "PASS" if ok else "FAIL",
    }
    path = os.path.join(OUT, "验证报告_快捷方式启动.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    print("\n证据文件:", path)
    print("结论:", report["结论"])

    # 5) 关闭 GUI 进程，避免残留窗口
    for d in procs:
        kill(d["ProcessId"])
        print("  已关闭 GUI 进程 PID=", d["ProcessId"])

    assert ok, "桌面快捷方式启动 GUI 验收未通过"
    return True


if __name__ == "__main__":
    main()
