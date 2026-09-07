"""helpers.py — 某學校材料報批共用 helper（2026-08-05 自我複查後固化）

解決 3 個昨日踩過嘅坑：
1. win32com 寫 Python datetime → 日期 off-by-one（UTC+8 轉錯）
2. BQ 格位估錯（H5 vs 某學校真實 G5:L5+M5:Q5）→ flip-flop
3. 冇先讀 known-good 行 cache 格位 → 每個文件重新試錯
"""
from datetime import date

EPOCH = date(1899, 12, 30)


def excel_serial(y, m, d):
    """Win32COM 寫日期要用 raw serial，唔好寫 Python datetime（會 off-by-one）。
    例：excel_serial(2026,8,5) -> 46239.0，cell 數字格式設 yyyy/mm/dd 即顯示 2026/08/05。"""
    return float((date(y, m, d) - EPOCH).days)


def set_bq_cell(ws, value, label="合約中之項目編號 Item Contratual："):
    """某學校模板 BQ 格：G5:L5 標籤 + M5:Q5 值。
    先拆晒 G5:R5 所有合併，再重組，避免有殘留 hidden 值（昨日 AR-018 就係中咗呢招）。"""
    for c in range(7, 19):  # G..R
        cell = ws.Cells(5, c)
        try:
            if cell.MergeCells:
                cell.MergeArea.UnMerge()
        except Exception:
            pass
    ws.Range("G5:L5").Merge()
    ws.Range("M5:Q5").Merge()
    ws.Range("G5").Value = label
    ws.Range("M5").Value = value


def fingerprint_bq(ws, row=5):
    """讀一條 known-good 材料行嘅 BQ 合併結構，return 座標，確保新行 copy 相同 format。
    每次生成前 call 一次確認，唔好硬編碼。"""
    return {
        "label_merge": ws.Range("G5").MergeArea.Address,
        "value_merge": ws.Range("M5").MergeArea.Address,
        "label_text": ws.Range("G5").Value,
    }


def category_checkbox(category):
    """某學校類別勾選格座標。"""
    return {"AR": "I7", "AG": "F8", "EG": "F8", "EL": "O7"}.get(category, "I7")
