# -*- coding: utf-8 -*-
"""
merge_bq.py — 統一 BQ 頁合併工具（material-approval-pipeline v2.0）
═══════════════════════════════════════════════════════════════════════
將工程標書 BQ 頁**正確插入**每份材料報批 PDF。

═══ 頁順序（用戶確認，唔可以擺錯位）═══
    審批表(P1)  →  BQ 頁(中間)  →  產品資料(P2..Pn)

═══ 兩種 BQ 來源 ═══
  情況 A 有文字層  ：--mode auto ，正則掃 BQ PDF 建立「BQ編號→[頁碼]」自動映射
  情況 B 圖像型掃描：--mode manual --mapping code2pages.json （人手提供 code→[頁]）

  兩者都支援 **跨頁 BQ 項目**（自動/手動附加全部相關頁，唔會漏頁）。

═══ CLI ═══
  python merge_bq.py \
    --zongbiao 總表.xlsx \
    --bq BQ.pdf \
    --input 報批表PDF目錄 \
    --output 輸出目錄 \
    [--product 產品資料PDF目錄] \
    [--mode auto|manual] \
    [--mapping code2pages.json] \
    [--src-sheet 0] [--src-start 7] \
    [--code-col 1] [--bq-col 3] [--name-col 2] \
    [--front-has-product] [--backup] [--render-bq DIR]

依賴：openpyxl, PyMuPDF(fitz)。若 BQ 係圖像型，先跑 --render-bq 渲染俾人讀序號，
再人手寫 mapping JSON 用 --mode manual 合併。

輸出目錄會生成 merge_summary.txt 記低每項成功/跳過原因。
"""
import sys, io, os, re, json, shutil, argparse
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

try:
    import openpyxl
    import fitz  # PyMuPDF
except ImportError as e:
    print("缺少依賴：", e); sys.exit(1)


# ───────────────────────── 工具函數 ─────────────────────────
def find_pdf(directory, code, code_map=None):
    """喺 directory 搵以 code 開頭嘅 PDF（支援 'EL-001 配電箱.pdf' 或 'EL-001.pdf'）。
    若提供 code_map（序號→文件名），優先按映射搵（解決總表序號≠PDF文件名，實戰常見）。"""
    if code_map and code in code_map:
        fn = code_map[code]
        if not fn.lower().endswith('.pdf'):
            fn += '.pdf'
        fp = os.path.join(directory, fn)
        if os.path.exists(fp):
            return fp
    for f in sorted(os.listdir(directory)):
        if f.lower().endswith('.pdf') and f.startswith(code):
            return os.path.join(directory, f)
    return None


def build_bq_text_index(bq_path):
    """
    掃 BQ PDF 文字層，建立 {bq編號: [頁碼(1-based)...]}。
    記低每個編號出現嘅**所有**頁（支援跨頁項目）。
    無文字層 → 返空 dict（代表圖像型，需 manual）。
    """
    doc = fitz.open(bq_path)
    idx = {}
    for i in range(doc.page_count):
        t = doc[i].get_text()
        if not t.strip():
            continue
        # 支援 1.1 / 2.1.4 / A.2 / B.2 / D1.1.1 / E.1.2.1（字母前綴可帶點，如 E.）等格式
        for m in re.findall(r'\b((?:[A-Za-z]+\.?)?\d+\.\d+(?:[-.]\d+)?)\b', t):
            idx.setdefault(m, []).append(i + 1)
    doc.close()
    return idx


def resolve_bq_pages(bq_code, text_index):
    """由 BQ 編號（如 1.1.1 / 2.1 / A.2）搵頁碼，支援去尾綴回退；跨頁返全部相關頁。"""
    if not bq_code:
        return None
    if bq_code in text_index:
        return sorted(set(text_index[bq_code]))
    parts = bq_code.split('.')
    for n in range(len(parts) - 1, 0, -1):
        base = '.'.join(parts[:n])
        if base in text_index:
            return sorted(set(text_index[base]))
    return None


def split_bq_codes(bq_str):
    """將總表 BQ 欄拆成多個編號，支援兩種實戰格式：
    - 分隔符式：'B3.3.2、B3.3.4、B3.3.5' / '1.1.1, 2.1'（某政府項目2020）
    - 括號補充式：'E.1.2.1（間接包括1.3.1-1.3.14項）'（某政府學院2023，主編號 + 括號內多編號/範圍）
    """
    if not bq_str:
        return []
    raw = str(bq_str).strip()
    codes = []
    if '（' in raw or '(' in raw:
        # 括號補充式：E.1.2.1（間接包括1.3.1-1.3.14項）
        seg = re.split(r'[（(]', raw)[0].strip()
        if seg and re.search(r'\d', seg):
            codes.append(seg)
        for paren in re.findall(r'[（(]([^）)]*)[）)]', raw):
            # 範圍 X.X.X - Y.Y.Y（同 prefix 最後一段遞增）→ 展開全部
            for a, b in re.findall(r'(\d+\.\d+(?:\.\d+)*)\s*[-–—~]\s*(\d+\.\d+(?:\.\d+)*)', paren):
                codes.extend(expand_range(a, b))
            # 括號內其它獨立編號
            for m in re.findall(r'\d+\.\d+(?:\.\d+)*', paren):
                codes.append(m)
    else:
        # 分隔符式：B3.3.2、B3.3.4、B3.3.5
        for p in re.split(r'[、,;，；\s]+', raw):
            p = p.strip()
            if p and re.search(r'\d', p):
                codes.append(p)
    # 去重保序
    seen = set(); out = []
    for c in codes:
        if c not in seen:
            seen.add(c); out.append(c)
    return out


def expand_range(a, b):
    """展開 '1.3.1'~'1.3.14' → ['1.3.1'...'1.3.14']（僅最後一段遞增、prefix 相同）。"""
    ap = a.split('.'); bp = b.split('.')
    if len(ap) != len(bp):
        return [a, b]
    try:
        sa = int(ap[-1]); sb = int(bp[-1])
    except ValueError:
        return [a, b]
    if sb < sa:
        return [a, b]
    prefix = ap[:-1]
    return ['.'.join(prefix + [str(i)]) for i in range(sa, sb + 1)]


def resolve_multi_bq(bq_str, text_index):
    """支援多 BQ 編號，返回去重排序後嘅全部相關頁（跨頁自動合併，唔會漏）。"""
    pages = []
    for code in split_bq_codes(bq_str):
        ps = resolve_bq_pages(code, text_index)
        if ps:
            pages.extend(ps)
    return sorted(set(pages)) if pages else None


def main():
    ap = argparse.ArgumentParser(description='材料報批 PDF 合併 BQ 頁')
    ap.add_argument('--zongbiao', required=True, help='材料審批總表 xlsx')
    ap.add_argument('--bq', required=True, help='工程標書 BQ PDF')
    ap.add_argument('--input', required=True, help='報批表 PDF 目錄（按 code 命名）')
    ap.add_argument('--output', required=True, help='合併後輸出目錄')
    ap.add_argument('--product', default=None, help='產品資料 PDF 目錄（可省略）')
    ap.add_argument('--mode', choices=['auto', 'manual'], default='auto')
    ap.add_argument('--mapping', default=None, help='manual 模式嘅 code→[頁] JSON')
    ap.add_argument('--code-map', default=None, help='序號→PDF文件名映射 JSON（總表序號與PDF文件名唔同時用）')
    ap.add_argument('--src-sheet', type=int, default=0, help='總表 Sheet 索引')
    ap.add_argument('--src-start', type=int, default=7, help='總表數據起始行')
    ap.add_argument('--code-col', type=int, default=1, help='總表「編號」列(1-based)')
    ap.add_argument('--bq-col', type=int, default=3, help='總表「BQ編號」列(1-based)')
    ap.add_argument('--name-col', type=int, default=2, help='總表「材料名」列(1-based)')
    ap.add_argument('--front-has-product', action='store_true',
                    help='報批表 PDF 已含產品資料（P1=審批表, P2+=產品），BQ 插喺中間')
    ap.add_argument('--backup', action='store_true', help='合併前備份 input 目錄到 output/_backup_input')
    ap.add_argument('--render-bq', default=None,
                    help='將 BQ PDF 每頁渲染成 PNG 到指定目錄（供人手讀序號後寫 mapping）')
    args = ap.parse_args()

    os.makedirs(args.output, exist_ok=True)

    # 渲染 BQ 供人手讀（圖像型 fallback 第一步）
    if args.render_bq:
        os.makedirs(args.render_bq, exist_ok=True)
        d = fitz.open(args.bq)
        for i in range(d.page_count):
            d[i].get_pixmap(dpi=110).save(os.path.join(args.render_bq, f'bq_p{i+1:02d}.png'))
        d.close()
        print(f'已渲染 BQ 共 {d.page_count} 頁到 {args.render_bq}，請人手讀序號後用 --mode manual --mapping')

    # 讀總表
    print('=== 讀取總表 ===')
    wb = openpyxl.load_workbook(args.zongbiao, data_only=True)
    ws = wb.worksheets[args.src_sheet]
    # 表頭關鍵詞：第 src-start 行可能係表頭（如「審批表格編號」），要跳過
    HEADER_HINTS = ('編號', '材料', '名稱', 'BQ', '審批', '审批', '品牌', '數量', '單位')
    items = []
    for r in range(args.src_start, ws.max_row + 1):
        code = ws.cell(r, args.code_col).value
        bq = ws.cell(r, args.bq_col).value
        name = ws.cell(r, args.name_col).value
        if (code is None) and (bq is None):
            break
        code_s = str(code).strip() if code else ''
        # 跳過表頭行
        if code_s and any(h in code_s for h in HEADER_HINTS):
            continue
        items.append({
            'code': str(code).strip() if code else '',
            'bq': str(bq).strip() if bq else '',
            'name': str(name).strip() if name else '',
        })
    print(f'共 {len(items)} 項')

    # 建立 code→BQ頁
    text_index = build_bq_text_index(args.bq)
    print(f'BQ 文字層索引：{len(text_index)} 個編號（空=圖像型，需 manual）')

    manual_map = {}
    if args.mode == 'manual':
        if not args.mapping:
            print('ERROR: manual 模式必須提供 --mapping JSON'); sys.exit(1)
        with open(args.mapping, 'r', encoding='utf-8') as f:
            manual_map = json.load(f)

    code_map = {}
    if args.code_map:
        with open(args.code_map, 'r', encoding='utf-8') as f:
            code_map = json.load(f)

    for it in items:
        if args.mode == 'auto':
            it['pages'] = resolve_multi_bq(it['bq'], text_index) if it['bq'] else None
        else:
            it['pages'] = manual_map.get(it['code'])

    unmatched = [it['code'] for it in items if not it['pages']]
    if args.mode == 'auto' and unmatched:
        print(f'⚠️  auto 模式有 {len(unmatched)} 項搵唔到 BQ 頁（圖像型？）：{unmatched[:10]}')
        print('     → 建議加 --render-bq 渲染後用手動 mapping，或轉 --mode manual')

    # 備份
    if args.backup:
        bk = os.path.join(args.output, '_backup_input')
        os.makedirs(bk, exist_ok=True)
        for f in os.listdir(args.input):
            if f.lower().endswith('.pdf'):
                shutil.copy2(os.path.join(args.input, f), os.path.join(bk, f))
        print(f'已備份 input 到 {bk}')

    # 合併
    print('=== 合併（頁序：審批表 → BQ → 產品資料）===')
    bq_doc = fitz.open(args.bq)
    ok = skip = 0
    summary = []

    for it in items:
        code = it['code']
        bp = find_pdf(args.input, code, code_map)
        if not bp:
            print(f'  跳過（無報批表PDF）: {code}'); skip += 1
            summary.append(f'✗ {code} 無報批表PDF'); continue
        if not it['pages']:
            print(f'  跳過（無BQ頁）: {code} (bq={it["bq"]})'); skip += 1
            summary.append(f'✗ {code} 無BQ頁 (bq={it["bq"]})'); continue

        out = fitz.open()
        src = fitz.open(bp)
        if args.front_has_product:
            # 報批表 PDF 已含 產品資料：P1=審批表, P2+=產品
            out.insert_pdf(src, from_page=0, to_page=0)
            for pg in it['pages']:
                out.insert_pdf(bq_doc, from_page=pg - 1, to_page=pg - 1)
            if src.page_count > 1:
                out.insert_pdf(src, from_page=1, to_page=src.page_count - 1)
        else:
            # 報批表(P1..) → BQ 頁 → 產品資料
            out.insert_pdf(src)
            for pg in it['pages']:
                out.insert_pdf(bq_doc, from_page=pg - 1, to_page=pg - 1)
            if args.product:
                pp = find_pdf(args.product, code)
                if pp:
                    out.insert_pdf(fitz.open(pp))
        src.close()

        safe = re.sub(r'[\\/:*?"<>|]', '-', it['name'])
        oname = f'{code} {safe}.pdf' if safe else f'{code}.pdf'
        out.save(os.path.join(args.output, oname)); out.close()
        print(f'  ✓ {code} {it["name"][:18]} → BQ p{it["pages"]}')
        ok += 1
        summary.append(f'✓ {code} {it["name"]} → BQ p{it["pages"]}')

    bq_doc.close()
    with open(os.path.join(args.output, 'merge_summary.txt'), 'w', encoding='utf-8') as f:
        f.write(f'BQ PDF : {args.bq}\n總表   : {args.zongbiao}\n模式   : {args.mode}\n')
        f.write(f'成功   : {ok}  跳過: {skip}\n\n' + '\n'.join(summary))
    print(f'\n完成：{ok} 成功，{skip} 跳過；輸出 {args.output}')
    print(f'匯總：{os.path.join(args.output, "merge_summary.txt")}')


if __name__ == '__main__':
    main()
