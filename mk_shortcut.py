# -*- coding: utf-8 -*-
"""在桌面创建指向启动工具.bat 的 .lnk 快捷方式（用 win32com 调用 Shell 接口）。"""
import _bootstrap  # 必须在第三方库之前，注入工具自带 lib
import os, sys
import win32com.client

desktop = os.path.join(os.environ["USERPROFILE"], "Desktop")
lnk = os.path.join(desktop, "A3试卷拆分打印工具.lnk")
tool_dir = os.path.dirname(os.path.abspath(__file__))
target = os.path.join(tool_dir, "启动工具.bat")

ws = win32com.client.Dispatch("WScript.Shell")
sc = ws.CreateShortcut(lnk)
sc.TargetPath = target
sc.WorkingDirectory = tool_dir
sc.Description = "A3 试卷拆分打印工具 v8（一键直印：点打印即直接打印，无需选择通道）"
sc.IconLocation = r"C:\Windows\System32\shell32.dll,14"
sc.Save()

print("SHORTCUT_CREATED:", lnk)
print("EXISTS:", os.path.exists(lnk))
