# -*- coding: utf-8 -*-
"""
Step 4: 由材料 PDF 提取型號，產生「型號草稿」俾 agent 覆核（唔直接硬寫總表）。
策略：
  - 有文字層嘅 PDF（潔具/機電類）：逐頁 get_text()，regex 抽型號候選。
  - 圖像型目錄（水泥/防水/天花/石膏磚/仿石面）：渲染 P5-P8 成 PNG，交 agent Read 肉眼讀。
  - 無材料文件嘅 item：標註「無對應 PDF」。

用法：改 CONFIG，跑：
  <WORKBUDDY_DIR>/binaries/python/envs/default/Scripts/python.exe extract_models.py
產出：<TEMP_DIR>\型號草稿.md + 渲染圖（如有）
"""
import fitz, os, re
import openpyxl

# ===================== CONFIG =====================
BASE         = r"D:\Users\user\Desktop\2026-07-10-某學校-2025-2026學校年度某學校中學部暑期校舍工程"
MATERIAL_DIR = os.path.join(BASE, "材料文件")
TOTAL_XLSX   = os.path.join(BASE, "材料審批", "材料審批總表-澳門某學校中學部暑期校舍工程_已填.xlsx")
RENDER_DIR   = r"<TEMP_DIR>\render\models"
DRAFT_OUT    = r"<TEMP_DIR>\型號草稿.md"
CODE_RE      = r'((?:AR|EL|AG|EG|AC)-\d{3})'
# =================================================

# 讀總表 code->name
wb = openpyxl.load_workbook(TOTAL_XLSX, data_only=True)
ws = wb.active
code2name = {}
for row in ws.iter_rows(values_only=True):
    code = row[0]
    if code and re.match(r'^(AR|EL|AG|EG|AC)-\d{3}$', str(code).strip()):
        code2name[str(code).strip()] = row[1]
wb.close()

pdfs = sorted(f for f in os.listdir(MATERIAL_DIR) if f.lower().endswith('.pdf'))
draft = ["# 型號草稿（待 agent + 用戶確認）\n"]
need_visual = []

for f in pdfs:
    codes = re.findall(CODE_RE, f)
    code = codes[0] if codes else "?"
    name = code2name.get(code, "?")
    path = os.path.join(MATERIAL_DIR, f)
    doc = fitz.open(path)
    text_chars = 0
    candidates = []
    for i, page in enumerate(doc):
        t = page.get_text()
        text_chars += len(t)
        if t:
            # 抽型號候選：含 Model/型號/產品編號 行，或長度 3-20 嘅英文數字混合 token
            for line in t.splitlines():
                if re.search(r'(model|型號|product|item|ref|型號|編號)', line, re.I):
                    candidates.append(line.strip()[:80])
    doc.close()
    if text_chars > 200:
        draft.append(f"## {code} {name}\n- 來源：PDF 文字層\n- 候選：{candidates[:5]}\n")
    else:
        # 圖像型：渲染 P5-P8 備查
        os.makedirs(RENDER_DIR, exist_ok=True)
        d = fitz.open(path)
        for idx in range(4, min(d.page_count, 8)):
            pix = d[idx].get_pixmap(dpi=110)
            pix.save(os.path.join(RENDER_DIR, f"{code}_p{idx+1}.png"))
        d.close()
        need_visual.append(code)
        draft.append(f"## {code} {name}\n- 來源：圖像型目錄（已渲染 P5-P8 到 {RENDER_DIR}）\n- 候選：⏳ 等 agent Read 渲染圖\n")

# 無材料文件嘅 item
have_codes = {re.findall(CODE_RE, f)[0] for f in pdfs if re.findall(CODE_RE, f)}
missing = [c for c in code2name if c not in have_codes]
if missing:
    draft.append("## ⚠️ 總表有但材料文件夾無 PDF（型號留空，等用戶補檔）\n- " + ", ".join(missing) + "\n")

with open(DRAFT_OUT, "w", encoding="utf-8") as fh:
    fh.write("\n".join(draft))
print("草稿 ->", DRAFT_OUT)
print("需肉眼讀圖嘅 item:", need_visual)
print("材料文件夾缺失嘅 item:", missing)
