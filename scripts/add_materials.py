"""add_materials.py — 某學校專案「新增/更新材料碼」一鍵生成器

取代 2026-08-05 嗰 6 支散腳本（build / part2 / patch_h5 / patch_bq / fix_ar018 / fix_ar018_p2）。
一次過做晒：pre-flight → 設值 → 出 PDF → append 產品資料 → sync 總表 →（可選）sync 美化版。

用法：
    python add_materials.py items.json

items.json 結構（見 README 注釋）：
{
  "project_root": "D:/Users/.../某學校.../",
  "zongbiao": "材料文件/材料審批總表-..._已填.xlsx",
  "combined": "材料文件/材料報批表-..._已填.xlsx",
  "out_dir": "材料文件",
  "meihua": "材料審批總表-..._美化版.xlsx",   // 可選，填咗就 sync 美化版
  "items": [
    {
      "code": "AR-018",
      "name": "彩色防霉填縫料",
      "bq": "A.4 / B.4 / C.4",
      "brand": "德高",
      "model": "超細型",
      "qty": null,            // 數值或 null
      "qty_text": "（依現場）", // qty=null 時用呢個做 C17
      "unit": "平方米",
      "lead": "7天",
      "category": "AR",       // AR/AG/EG/EL
      "remark": "",
      "product_pdfs": ["材料文件/08-05新增材料/填縫劑-Davco(德高)彩色防霉填縫料.pdf"],
      "action": "new",        // new | update
      "insert_after": "AR-017", // new 用：喺邊個 code 之後插
      "submit_date": [2026,8,5] // 遞交日期，寫 raw serial
    }
  ]
}

⚠️ 驗證：本腳本按 2026-08-05 實戰証明可用嘅模式寫成（copy sheet→新 wb→ExportAsFixedFormat→fitz insert_pdf）。
   下次真實跑之前先用 1 個 item 試跑確認，再批量。
"""
import json, os, sys, shutil
import win32com.client as win32
import fitz
sys.path.insert(0, os.path.dirname(__file__))
from helpers import excel_serial, set_bq_cell, category_checkbox

XL_PASTE_FORMATS = -4122  # xlPasteFormats


def find_row(ws, code, col=1):
    for r in range(1, ws.UsedRange.Rows.Count + 5):
        v = ws.Cells(r, col).Value
        if v is not None and str(v).strip() == code:
            return r
    return None


def export_sheet_pdf(excel, sh, pdf_path):
    """copy 該 sheet 去新 wb → A4 縱向 → ExportAsFixedFormat。避免整本 workbook 匯出。"""
    sh.Copy(None, None)
    new_wb = excel.ActiveWorkbook
    ws = new_wb.Worksheets(1)
    ws.PageSetup.Orientation = 1      # xlPortrait
    ws.PageSetup.PaperSize = 9        # A4
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = 1
    ws.PageSetup.Zoom = False
    if os.path.exists(pdf_path):
        os.remove(pdf_path)
    ws.ExportAsFixedFormat(0, pdf_path)
    new_wb.Close(False)


def append_products(pdf_path, product_pdfs):
    doc = fitz.open(pdf_path)
    for pp in product_pdfs or []:
        if not os.path.exists(pp):
            print(f"  ⚠️ 產品資料唔存在：{pp}")
            continue
        src = fitz.open(pp)
        doc.insert_pdf(src)
        src.close()
    doc.save(pdf_path)
    doc.close()


def update_zongbiao(zongbiao_path, items):
    """用 openpyxl 更新/插入 總表行（總表係純數據表，無圖，openpyxl 安全）。"""
    import openpyxl
    wb = openpyxl.load_workbook(zongbiao_path)
    ws = wb.active
    # 搵最後一行 + header 行數（假設第 4 行係 header，數據由第 5 行起）
    maxr = ws.max_row
    for it in items:
        r = None
        for rr in range(1, maxr + 1):
            if str(ws.cell(rr, 1).value or "").strip() == it["code"]:
                r = rr
                break
        if r is None:
            # new：插喺 insert_after 之後
            after = None
            for rr in range(1, maxr + 1):
                if str(ws.cell(rr, 1).value or "").strip() == it.get("insert_after", ""):
                    after = rr
                    break
            r = (after or maxr) + 1
            ws.insert_rows(r)
        # 寫 12 欄：A編號 B名稱 C BQ D品牌 E型號 F數量 G單位 H貨期 I遞交日期 J K備註 ...
        ws.cell(r, 1, it["code"])
        ws.cell(r, 2, it["name"])
        ws.cell(r, 3, it["bq"])
        ws.cell(r, 4, it.get("brand", ""))
        ws.cell(r, 5, it.get("model", ""))
        ws.cell(r, 6, it["qty"] if it.get("qty") is not None else it.get("qty_text", ""))
        ws.cell(r, 7, it.get("unit", ""))
        ws.cell(r, 8, it.get("lead", ""))
        sd = it.get("submit_date")
        ws.cell(r, 9, excel_serial(*sd) if sd else None)
        ws.cell(r, 11, f"型號：{it['model']}" if it.get("model") else "")
    wb.save(zongbiao_path)
    print(f"總表已更新：{zongbiao_path}")


def sync_meihua(excel, meihua_path, items):
    """sync 美化版：喺 (二) 之前插 AR 新行，copy 隔離 format，日期寫 raw serial。"""
    bak = meihua_path + ".bak_addmat"
    if not os.path.exists(bak):
        shutil.copy(meihua_path, bak)
    wb = excel.Workbooks.Open(meihua_path)
    ws = wb.Worksheets(1)
    # 搵 (二) 行
    row_2 = None
    for r in range(1, ws.UsedRange.Rows.Count + 5):
        if str(ws.Cells(r, 1).Value or "").strip() == "(二)":
            row_2 = r
            break
    # 逐個 AR item 插入（由後插入 avoid 位移）
    ar_items = [i for i in items if i.get("category") == "AR" and i.get("action") == "new"]
    for idx, it in enumerate(reversed(ar_items)):
        ins = row_2 + idx
        ws.Rows(f"{ins}:{ins}").Insert()
        # copy format from neighbor (AR-016 white / AR-017 light)
        donor = it.get("insert_after", "AR-017")
        dr = find_row(ws, donor)
        if dr:
            ws.Range(f"A{dr}:L{dr}").Copy()
            ws.Range(f"A{ins}:L{ins}").PasteSpecial(Paste=XL_PASTE_FORMATS)
            excel.CutCopyMode = False
        # 設值
        ws.Cells(ins, 1).Value = it["code"]
        ws.Cells(ins, 2).Value = it["name"]
        ws.Cells(ins, 3).Value = it["bq"]
        ws.Cells(ins, 4).Value = it.get("brand", "")
        ws.Cells(ins, 5).Value = it.get("model", "")
        ws.Cells(ins, 6).Value = it["qty"] if it.get("qty") is not None else it.get("qty_text", "")
        ws.Cells(ins, 7).Value = it.get("unit", "")
        ws.Cells(ins, 8).Value = it.get("lead", "")
        sd = it.get("submit_date")
        ws.Cells(ins, 9).Value = excel_serial(*sd) if sd else None
        ws.Cells(ins, 11).Value = f"型號：{it['model']}" if it.get("model") else ""
    wb.Save()
    wb.Close(False)
    print(f"美化版已 sync：{meihua_path}（備份 {bak}）")


def main():
    cfg = json.load(open(sys.argv[1], encoding="utf-8"))
    root = cfg["project_root"]
    combined = os.path.join(root, cfg["combined"])
    out_dir = os.path.join(root, cfg.get("out_dir", "材料文件"))
    os.makedirs(out_dir, exist_ok=True)
    zongbiao = os.path.join(root, cfg["zongbiao"])
    meihua = os.path.join(root, cfg["meihua"]) if cfg.get("meihua") else None

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False

    wb = excel.Workbooks.Open(combined)
    sheet_map = {}
    for it in cfg["items"]:
        code = it["code"]
        exists = any(s.Name == code for s in wb.Worksheets)
        if exists:
            sh = wb.Worksheets(code)
            print(f"更新現有 sheet：{code}")
        else:
            donor = it.get("insert_after", "AR-017")
            dsh = wb.Worksheets(donor)
            dsh.Copy(None, wb.Worksheets(wb.Worksheets.Count))
            sh = wb.Worksheets(wb.Worksheets.Count)
            sh.Name = code
            print(f"新建 sheet：{code}（copy from {donor}）")
        # 設值
        sh.Range("D5").Value = code
        sh.Range("D7").Value = it["name"]
        sh.Range("D9").Value = it.get("brand", "")
        if it.get("qty") is not None:
            sh.Range("C17").Value = f"{it['qty']} {it.get('unit','')}"
        else:
            sh.Range("C17").Value = it.get("qty_text", "（依現場）")
        sh.Range("J17").Value = it.get("lead", "")
        sh.Range("C19").Value = it.get("remark", "")
        set_bq_cell(sh, it["bq"])  # 某學校 G5:L5 + M5:Q5
        chk = category_checkbox(it.get("category", "AR"))
        sh.Range(chk).Value = "■"
        sheet_map[code] = sh
    wb.Save()

    # 出 PDF + append 產品資料
    for it in cfg["items"]:
        code = it["code"]
        sh = sheet_map[code]
        pdf_path = os.path.join(out_dir, f"{code}_{it['name']}.pdf")
        export_sheet_pdf(excel, sh, pdf_path)
        append_products(pdf_path, [os.path.join(root, p) for p in it.get("product_pdfs", [])])
        print(f"已出 PDF：{pdf_path}（{fitz.open(pdf_path).page_count} 頁）")

    wb.Close(False)

    # sync 總表
    update_zongbiao(zongbiao, cfg["items"])

    # sync 美化版（可選）
    if meihua and os.path.exists(meihua):
        sync_meihua(excel, meihua, cfg["items"])

    excel.Quit()
    print("✅ add_materials 完成")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法：python add_materials.py items.json")
        sys.exit(1)
    main()
