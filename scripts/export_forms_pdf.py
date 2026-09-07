# -*- coding: utf-8 -*-
"""
Step 2: 將分項報批表 xlsx 轉 A4 縱向 PDF。
關鍵：強制 Orientation=1（縱向），否則易跟範本/預設出橫向。

用法：改 CONFIG.SRC 為含 xlsx 嘅文件夾，跑：
  <WORKBUDDY_DIR>/binaries/python/envs/default/Scripts/python.exe export_forms_pdf.py
"""
import os
import win32com.client

# ===================== CONFIG =====================
SRC = r"D:\Users\user\Desktop\2026-07-10-某學校-2025-2026學校年度某學校中學部暑期校舍工程\材料文件\材料審批表（分項）"
# =================================================

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
try:
    for fn in sorted(os.listdir(SRC)):
        if not fn.lower().endswith(".xlsx") or fn.startswith("~$"):
            continue
        src = os.path.join(SRC, fn)
        out = os.path.splitext(src)[0] + ".pdf"
        if os.path.exists(out):
            try:
                os.remove(out)
            except Exception:
                pass
        wb = excel.Workbooks.Open(src)
        try:
            ws = wb.ActiveSheet
            ps = ws.PageSetup
            ps.Orientation = 1          # xlPortrait 縱向
            try:
                ps.PaperSize = 9        # xlPaperA4
            except Exception:
                pass
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = 1
            ps.Zoom = False
            wb.ExportAsFixedFormat(0, out)   # xlTypePDF=0
            print("PDF OK:", os.path.basename(out))
        finally:
            wb.Close(False)
    print("DONE")
finally:
    excel.Quit()
