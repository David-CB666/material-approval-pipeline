# -*- coding: utf-8 -*-
"""
Step 3: 將工程標書 BQ 頁附到每份材料 PDF，順序 = 審批表(P1) + BQ頁面 + 產品資料(P2..Pn)。
⚠️ 順序關鍵：BQ 插喺審批表之後、產品資料之前（唔係擺最尾）。

⚠️ serial->page 映射係項目相關！某學校項目（12頁=3樓×4頁）映射如下；
   其他項目必須 render + 肉眼讀工程標書推導自己嘅映射，唔好硬抄。

用法：改 CONFIG，跑：
  <WORKBUDDY_DIR>/binaries/python/envs/default/Scripts/python.exe append_bq.py
"""
import fitz, os, re, shutil
import openpyxl

# ===================== CONFIG（每個項目改呢度） =====================
BASE         = r"D:\Users\user\Desktop\2026-07-10-某學校-2025-2026學校年度某學校中學部暑期校舍工程"
MATERIAL_DIR = os.path.join(BASE, "材料文件")
TENDER_PATH  = os.path.join(BASE, "工程標書.pdf")
TOTAL_XLSX   = os.path.join(BASE, "材料審批", "材料審批總表-澳門某學校中學部暑期校舍工程_已填.xlsx")
BACKUP_DIR   = r"<TEMP_DIR>\材料文件_加BQ前備份"
CODE_RE      = r'((?:AR|EL|AG|EG|AC)-\d{3})'   # 注意：成個碼 capture，唔可以漏 -001
# 某學校項目映射（其他項目自己 render 推）：序號 1-6→[1,5,9] 7-13→[2,6,10] 14-21→[3,7,11] 22-23→[4,8,12]
# =====================================================================

# 1. 總表 item-code -> serial
wb = openpyxl.load_workbook(TOTAL_XLSX)
ws = wb.active
code2serial = {}
for row in ws.iter_rows(values_only=True):
    code, bq = row[0], row[2]
    if code and re.match(r'^(AR|EL|AG|EG|AC)-\d{3}$', str(code).strip()):
        first = str(bq).strip().split('/')[0].strip()   # "A.14"
        serial = int(first.split('.')[1])
        code2serial[str(code).strip()] = serial
print(f"Mapped {len(code2serial)} item codes")

def serial_to_pages(serial):
    if serial <= 6:    g = 1
    elif serial <= 13: g = 2
    elif serial <= 21: g = 3
    else:              g = 4
    return [g, g+4, g+8]   # A/B/C 樓層頁（1-based）

# 2. 備份
os.makedirs(BACKUP_DIR, exist_ok=True)
pdfs = [f for f in os.listdir(MATERIAL_DIR)
        if f.lower().endswith('.pdf') and '_sample' not in f and '_test' not in f]
for f in pdfs:
    shutil.copy2(os.path.join(MATERIAL_DIR, f), os.path.join(BACKUP_DIR, f))
print(f"Backed up {len(pdfs)} PDFs")

# 3. 處理
ten_doc = fitz.open(TENDER_PATH)
log = []
for f in sorted(pdfs):
    src = os.path.join(MATERIAL_DIR, f)
    mat = fitz.open(src)
    codes = re.findall(CODE_RE, f)
    pages = set()
    for c in codes:
        s = code2serial.get(c)
        if s:
            pages.update(serial_to_pages(s))
    pages = sorted(pages)
    if not pages:
        log.append(f"SKIP (no serial): {f}")
        mat.close(); continue
    out = fitz.open()
    out.insert_pdf(mat, from_page=0, to_page=0)               # 審批表 P1
    for pg in pages:
        out.insert_pdf(ten_doc, from_page=pg-1, to_page=pg-1)  # BQ 頁（中間）
    if mat.page_count > 1:
        out.insert_pdf(mat, from_page=1, to_page=mat.page_count-1)  # 產品資料
    tmp = src + ".tmp"
    out.save(tmp); out.close(); mat.close()
    os.replace(tmp, src)
    log.append(f"OK  {f:40s} -> {fitz.open(src).page_count:2d} pages | BQ {pages}")
ten_doc.close()
print("\n".join(log))
