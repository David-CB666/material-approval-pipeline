"""
batch_brand_fill.py
批量填充工具：讀取材料清單Excel，自動查詢品牌後寫入新列
"""

import sys
import json
from pathlib import Path
from openpyxl import load_workbook, Workbook
from brand_recommender import BrandRecommender

DEFAULT_DB = Path(__file__).parent.parent / "references" / "brand_db.json"


def batch_fill(
    input_xlsx: str,
    output_xlsx: str | None = None,
    db_path: str | None = None,
    sheet_index: int = 0,
    spec_col: int = 2,  # 默認第2列（B列）為標書型號
    name_col: int | None = None,  # 默認不填材料名稱列
    add_recommend_col: bool = True,
    add_fill_text_col: bool = True,
    add_category_col: bool = True,
    add_confidence_col: bool = False,
    dry_run: bool = False,
) -> dict:
    """
    批量填充品牌

    參數：
        input_xlsx: 輸入Excel路徑
        output_xlsx: 輸出Excel路徑（None時自動在原檔案名後加 _品牌填充）
        db_path: 品牌知識庫路徑
        sheet_index: 工作表索引（默認第1個）
        spec_col: 標書型號列（1-based，默認2=B列）
        name_col: 材料名稱列（1-based，None時不填）
        add_recommend_col: 是否新增「推荐品牌」列
        add_fill_text_col: 是否新增「報批填入文本」列
        add_category_col: 是否新增「分類」列
        add_confidence_col: 是否新增「匹配度」列
        dry_run: 是否僅預覽不寫入

    返回：
        {
            "total": int,
            "matched": int,
            "fuzzy": int,
            "none": int,
            "results": [(row, spec, recommended, fill_text), ...]
        }
    """
    recommender = BrandRecommender(db_path)

    wb = load_workbook(input_xlsx)
    ws = wb.worksheets[sheet_index]

    # 確定輸出路徑
    if output_xlsx is None:
        p = Path(input_xlsx)
        output_xlsx = str(p.parent / f"{p.stem}_品牌填充{p.suffix}")

    # 找到最大列
    max_col = ws.max_column

    # 新增列位置
    new_col_base = max_col + 1
    col_map = {}
    if add_category_col:
        col_map["category"] = new_col_base
        new_col_base += 1
    if add_recommend_col:
        col_map["recommended"] = new_col_base
        new_col_base += 1
    if add_fill_text_col:
        col_map["fill_text"] = new_col_base
        new_col_base += 1
    if add_confidence_col:
        col_map["confidence"] = new_col_base

    # 寫入表頭
    header_row = 1
    for col_name, col_idx in col_map.items():
        headers = {
            "category": "分類",
            "recommended": "推薦品牌",
            "fill_text": "報批填入文本",
            "confidence": "匹配度",
        }
        ws.cell(row=header_row, column=col_idx, value=headers.get(col_name, col_name))

    # 遍歷數據行
    results = []
    stats = {"total": 0, "matched": 0, "fuzzy": 0, "none": 0, "category": 0}

    for row in range(header_row + 1, ws.max_row + 1):
        spec_cell = ws.cell(row=row, column=spec_col)
        spec_value = spec_cell.value

        if not spec_value or str(spec_value).strip() == "":
            continue

        spec = str(spec_value).strip()
        stats["total"] += 1

        # 查詢品牌
        result = recommender.search(spec)

        match_type = result["match_type"]
        stats[match_type] = stats.get(match_type, 0) + 1

        # 寫入結果
        if not dry_run:
            if "category" in col_map:
                ws.cell(row=row, column=col_map["category"], value=result["category"])
            if "recommended" in col_map:
                ws.cell(row=row, column=col_map["recommended"],
                        value=", ".join(result["recommended"]) if result["recommended"] else "（待確認）")
            if "fill_text" in col_map:
                ws.cell(row=row, column=col_map["fill_text"], value=result["fill_text"])
            if "confidence" in col_map:
                ws.cell(row=row, column=col_map["confidence"],
                        value=f"{result['confidence']:.0%}")

        results.append({
            "row": row,
            "spec": spec,
            "match_type": match_type,
            "recommended": result["recommended"],
            "fill_text": result["fill_text"],
            "confidence": result["confidence"],
        })

    # 保存
    if not dry_run:
        wb.save(output_xlsx)
        print(f"✅ 已保存至：{output_xlsx}")

    return {**stats, "results": results, "output": output_xlsx}


def print_summary(stats: dict):
    """打印統計摘要"""
    total = stats["total"]
    print()
    print("=" * 50)
    print(f"📊 批量填充統計（共 {total} 項）")
    print("-" * 50)
    print(f"  ✅ 精確匹配：{stats.get('exact', 0)} 項")
    print(f"  🔍 模糊匹配：{stats.get('fuzzy', 0)} 項")
    print(f"  📂 按類別推薦：{stats.get('category', 0)} 項")
    print(f"  ❌ 無匹配：{stats.get('none', 0)} 項")
    print("=" * 50)


def main():
    import argparse

    parser = argparse.ArgumentParser(description="批量填充材料品牌")
    parser.add_argument("input", help="輸入Excel文件路徑")
    parser.add_argument("-o", "--output", help="輸出Excel文件路徑")
    parser.add_argument("-d", "--db", default=str(DEFAULT_DB), help="品牌知識庫路徑")
    parser.add_argument("-s", "--sheet", type=int, default=0, help="工作表索引（默認0）")
    parser.add_argument("-c", "--col", type=int, default=2, help="標書型號列（1-based，默認2=B）")
    parser.add_argument("--dry-run", action="store_true", help="僅預覽不寫入")
    parser.add_argument("--no-category", action="store_true", help="不新增分類列")
    parser.add_argument("--no-fill-text", action="store_true", help="不新增報批填入文本列")
    parser.add_argument("--show-confidence", action="store_true", help="新增匹配度列")

    args = parser.parse_args()

    print(f"📂 輸入文件：{args.input}")
    if args.dry_run:
        print("🔍 預覽模式（不寫入文件）")

    result = batch_fill(
        input_xlsx=args.input,
        output_xlsx=args.output,
        db_path=args.db,
        sheet_index=args.sheet,
        spec_col=args.col,
        add_category_col=not args.no_category,
        add_fill_text_col=not args.no_fill_text,
        add_confidence_col=args.show_confidence,
        dry_run=args.dry_run,
    )

    print_summary(result)

    # 顯示前10條結果
    print()
    print("📋 前10條結果預覽：")
    print(f"{'行':<4} {'標書型號':<30} {'匹配':<6} {'推薦品牌':<20} {'填入文本'}")
    print("-" * 100)
    for r in result["results"][:10]:
        match_icon = {"exact": "✅", "fuzzy": "🔍", "category": "📂", "none": "❌"}
        icon = match_icon.get(r["match_type"], "?")
        brands = ", ".join(r["recommended"]) if r["recommended"] else "—"
        print(f"{r['row']:<4} {r['spec'][:28]:<30} {icon:<6} {brands[:18]:<20} {r['fill_text']}")

    if len(result["results"]) > 10:
        print(f"... 還有 {len(result['results']) - 10} 項")


if __name__ == "__main__":
    main()
