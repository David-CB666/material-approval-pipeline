"""
brand_recommender.py
澳門電機工程品牌推薦引擎
基於162條材料審批黃金清單
"""

import json
import re
from difflib import SequenceMatcher
from pathlib import Path

DB_PATH = Path(__file__).parent.parent / "references" / "brand_db.json"


class BrandRecommender:
    """品牌推薦引擎"""

    def __init__(self, db_path: str | None = None):
        path = Path(db_path) if db_path else DB_PATH
        with open(path, "r", encoding="utf-8") as f:
            self.db = json.load(f)
        self._build_index()

    def _build_index(self):
        """建立快速查找索引"""
        self._exact_index: dict[str, dict] = {}
        self._all_items: list[tuple[str, dict, str]] = []  # (spec, item, cat_label)

        for cat_key, cat_data in self.db["categories"].items():
            label = cat_data["label"]
            for item in cat_data.get("items", []):
                spec = item["tender_spec"].strip()
                # 精確索引（小寫脫敏）
                key = spec.lower().replace(" ", "")
                self._exact_index[key] = (item, label)
                self._all_items.append((spec, item, label))

    def search(self, query: str) -> dict:
        """
        輸入型號 → 返回推薦結果
        返回格式：
        {
            "match_type": "exact" | "fuzzy" | "category" | "none",
            "spec": str,
            "category": str,
            "recommended": [str],
            "alternative": [str],
            "origin": str,
            "model": str,
            "bq": str,
            "source": str,
            "notes": str,
            "fill_text": str,
            "confidence": float,
            "similar_items": [dict]  # fuzzy match 時返回
        }
        """
        query = query.strip()
        if not query:
            return self._empty_result("請輸入型號或材料名稱")

        # Step 1: 精確匹配
        exact = self._exact_match(query)
        if exact:
            return exact

        # Step 2: 模糊匹配
        fuzzy = self._fuzzy_match(query, threshold=0.5)
        if fuzzy:
            return fuzzy

        # Step 3: 分類關鍵詞匹配
        category = self._category_match(query)
        if category:
            return category

        # Step 4: 無匹配
        return self._no_match_result(query)

    def _exact_match(self, query: str) -> dict | None:
        """精確匹配 Tender Spec"""
        q_lower = query.lower().replace(" ", "")

        for spec, item, label in self._all_items:
            spec_key = spec.lower().replace(" ", "")
            if spec_key == q_lower or q_lower in spec_key or spec_key in q_lower:
                return self._build_result(item, label, "exact", 1.0, query)

        return None

    def _fuzzy_match(self, query: str, threshold: float = 0.5) -> dict | None:
        """模糊匹配（基於關鍵詞重疊）"""
        q_clean = re.sub(r"[^\w\u4e00-\u9fff]", " ", query).lower()
        q_words = set(q_clean.split())

        scored: list[tuple[float, str, dict, str]] = []

        for spec, item, label in self._all_items:
            spec_clean = re.sub(r"[^\w\u4e00-\u9fff]", " ", spec).lower()
            spec_words = set(spec_clean.split())

            # 重疊關鍵詞數量
            overlap = q_words & spec_words
            if not overlap:
                # 使用 SequenceMatcher
                ratio = SequenceMatcher(None, q_clean, spec_clean).ratio()
            else:
                # 有重疊，提高分數
                overlap_ratio = len(overlap) / max(len(q_words), len(spec_words))
                ratio = 0.5 + 0.5 * overlap_ratio

            if ratio >= threshold:
                scored.append((ratio, spec, item, label))

        if not scored:
            return None

        scored.sort(key=lambda x: -x[0])
        best_score, best_spec, best_item, best_label = scored[0]

        # 取前三個相似項
        similar = [
            {"spec": s, "brands": i["brands"], "score": sc}
            for sc, s, i, l in scored[:3]
        ]

        result = self._build_result(best_item, best_label, "fuzzy", best_score, query)
        result["similar_items"] = similar
        return result

    def _category_match(self, query: str) -> dict | None:
        """分類關鍵詞匹配"""
        q_lower = query.lower()
        keywords_map = {
            # 強電/電氣
            "MCB": ("EL_electrical", "微型断路器"),
            "RCCB": ("EL_electrical", "漏电断路器"),
            "RCBO": ("EL_electrical", "漏电断路器"),
            "微型断路器": ("EL_electrical", "微型断路器"),
            "漏电断路器": ("EL_electrical", "漏电断路器"),
            "漏電開關": ("EL_electrical", "漏电断路器"),
            "隔離開關": ("EL_electrical", "隔离开关"),
            "斷路器": ("EL_electrical", "断路器"),
            "配電箱": ("EL_electrical", "配电箱"),
            "電箱": ("EL_electrical", "配电箱"),
            "開關": ("EL_electrical", "开关插座"),
            "插座": ("EL_electrical", "插座"),
            "電纜": ("EL_electrical", "电缆电线"),
            "電線": ("EL_electrical", "电缆电线"),
            "開關掣": ("EL_electrical", "开关插座"),
            # 燈具/照明
            "燈": ("EL_lighting", "灯具照明"),
            "燈具": ("EL_lighting", "灯具照明"),
            "光管": ("EL_lighting", "灯具照明"),
            "LED": ("EL_lighting", "灯具照明"),
            "照明": ("EL_lighting", "灯具照明"),
            # 弱電/ELV
            "網線": ("EL_elv", "弱电网络"),
            "網絡": ("EL_elv", "弱电网络"),
            "交換機": ("EL_elv", "弱电网络"),
            "門禁": ("EL_elv", "弱电门禁"),
            "CCTV": ("EL_elv", "弱电监控"),
            "閉路": ("EL_elv", "弱电监控"),
            "監控": ("EL_elv", "弱电监控"),
            # 給排水/潔具
            "潔具": ("PLUMBING", "给排水洁具"),
            "龍頭": ("PLUMBING", "给排水洁具"),
            "座廁": ("PLUMBING", "给排水洁具"),
            "小便": ("PLUMBING", "给排水洁具"),
            "洗手盆": ("PLUMBING", "给排水洁具"),
            "排水": ("PLUMBING", "给排水"),
            "供水": ("PLUMBING", "给排水"),
            "閥": ("PLUMBING", "阀门管件"),
            # 空調/通風
            "空调": ("HVAC", "空调通风"),
            "冷氣": ("HVAC", "空调通风"),
            "抽氣扇": ("HVAC", "空调通风"),
            "風管": ("HVAC", "通风"),
            "天花風扇": ("HVAC", "空调用品"),
            "風扇": ("HVAC", "空调用品"),
            "VRF": ("HVAC", "空调通风"),
            "分體機": ("HVAC", "空调通风"),
            "保溫": ("HVAC", "保温材料"),
            # 牆身/飾面
            "天花": ("FINISHING", "天花/饰面"),
            "油漆": ("FINISHING", "油漆涂料"),
            "磁磚": ("FINISHING", "磁砖地材"),
            "牆磚": ("FINISHING", "磁砖地材"),
            "地磚": ("FINISHING", "磁砖地材"),
            "防水": ("FINISHING", "防水泥水"),
            "膠板": ("FINISHING", "飾面膠板"),
            "冰火板": ("FINISHING", "飾面膠板"),
            # 門/五金
            "門扇": ("DOORS_HARDWARE", "门/五金"),
            "趟門": ("DOORS_HARDWARE", "门/五金"),
            "木門": ("DOORS_HARDWARE", "门/木作"),
            "防火門": ("DOORS_HARDWARE", "门/五金"),
            "間隔": ("DOORS_HARDWARE", "卫生间隔断"),
            "扶手": ("DOORS_HARDWARE", "五金/扶手"),
            # 消防
            "滅火": ("FIRE", "消防"),
            "花灑頭": ("FIRE", "消防"),
            "感應器": ("FIRE", "消防"),
            "消防": ("FIRE", "消防"),
            "警鐘": ("FIRE", "消防"),
            "防火閘": ("FIRE", "消防"),
        }

        for kw, (cat_key, cat_label) in keywords_map.items():
            if kw.lower() in q_lower or kw in q_lower:
                cat = self.db["categories"].get(cat_key, {})
                items = cat.get("items", [])
                if items:
                    # 返回该类别最常见的品牌
                    brand_count: dict[str, int] = {}
                    for item in items:
                        for b in item.get("brands", []):
                            brand_count[b] = brand_count.get(b, 0) + 1
                    top_brands = sorted(brand_count.items(), key=lambda x: -x[1])[:3]
                    recommended = [b for b, _ in top_brands]

                    return {
                        "match_type": "category",
                        "spec": query,
                        "category": cat_label,
                        "recommended": recommended,
                        "alternative": self._get_alternatives(recommended),
                        "origin": "—",
                        "model": "—",
                        "bq": "—",
                        "source": f"品牌知识库 v{self.db['version']}（按类别推荐）",
                        "notes": f"根據「{cat_label}」類別歷史數據推薦",
                        "fill_text": f"{recommended[0]}（{cat_label}）或同級" if recommended else "（待確認）",
                        "confidence": 0.4,
                        "similar_items": []
                    }

        return None

    def _get_alternatives(self, recommended: list[str]) -> list[str]:
        """获取备选品牌"""
        premium = self.db["global_alternatives"]["premium"]
        alternatives = [b for b in premium if b not in recommended]
        return alternatives[:3]

    def _build_result(
        self, item: dict, label: str, match_type: str, confidence: float, query: str
    ) -> dict:
        """构建结果字典"""
        brands = item.get("brands", [])
        main_brand = brands[0] if brands else "—"

        # 报批表填入文本
        if brands:
            # 尝试判断产地
            origin = self._guess_origin(main_brand)
            fill_text = f"{main_brand}（{origin}）或同級"
        else:
            fill_text = "（待確認）"

        return {
            "match_type": match_type,
            "spec": item.get("tender_spec", query),
            "category": label,
            "recommended": brands,
            "alternative": self._get_alternatives(brands),
            "origin": self._guess_origin(main_brand),
            "model": item.get("model", "—"),
            "bq": item.get("bq", "—"),
            "source": item.get("source", "—"),
            "notes": self._build_notes(item),
            "fill_text": fill_text,
            "confidence": confidence,
            "similar_items": []
        }

    def _guess_origin(self, brand: str) -> str:
        """猜测品牌产地"""
        brand_lower = brand.lower()
        origin_map = {
            "legrand": "法國", "schneider": "法國", "siemens": "德國",
            "abb": "瑞士", "philips": "荷蘭", "toto": "日本",
            "cotto": "泰國", "american standard": "美國",
            "mk": "英國", "commscope": "美國", "nexans": "法國",
            "hitachi": "日本", "daikin": "日本", "mitsubishi": "日本",
            "panasonic": "日本", "samsung": "韓國", "epson": "日本",
            "施耐德": "法國", "飛利浦": "荷蘭", "華為": "中國",
            "海康威視": "中國", "綠聯": "中國", "正泰": "中國",
            "德力西": "中國", "美的": "中國", "日立": "日本",
            "大金": "日本", "kdk": "日本", "人民電器": "中國",
        }
        return origin_map.get(brand_lower, "—")

    def _build_notes(self, item: dict) -> str:
        """构建备注"""
        notes_parts = []
        if item.get("model") and item["model"] not in ("—", "詳見附件"):
            notes_parts.append(f"型號：{item['model']}")
        if item.get("bq") and item["bq"] not in ("—",):
            notes_parts.append(f"BQ：{item['bq']}")
        return "；".join(notes_parts) if notes_parts else "—"

    def _empty_result(self, msg: str) -> dict:
        return {
            "match_type": "none",
            "spec": "",
            "category": "—",
            "recommended": [],
            "alternative": [],
            "origin": "—",
            "model": "—",
            "bq": "—",
            "source": "—",
            "notes": msg,
            "fill_text": "（請輸入型號）",
            "confidence": 0.0,
            "similar_items": []
        }

    def _no_match_result(self, query: str) -> dict:
        return {
            "match_type": "none",
            "spec": query,
            "category": "—",
            "recommended": [],
            "alternative": [],
            "origin": "—",
            "model": "—",
            "bq": "—",
            "source": "—",
            "notes": "知識庫中未找到匹配項，需人工確認品牌",
            "fill_text": "（待確認）",
            "confidence": 0.0,
            "similar_items": []
        }

    def get_fill_text(self, result: dict) -> str:
        """生成报批表填入文本"""
        if result["recommended"]:
            main = result["recommended"][0]
            origin = result.get("origin", "")
            if origin and origin != "—":
                return f"{main}（{origin}）或同級"
            return f"{main}或同級"
        return "（待確認）"

    def list_categories(self) -> list[dict]:
        """列出所有分类"""
        return [
            {"key": k, "label": v["label"], "count": len(v.get("items", []))}
            for k, v in self.db["categories"].items()
        ]


def print_recommendation(result: dict):
    """格式化打印推荐结果"""
    match_icon = {
        "exact": "✅ 精確匹配",
        "fuzzy": "🔍 模糊匹配",
        "category": "📂 按類別推薦",
        "none": "❌ 無匹配"
    }

    print()
    print("=" * 50)
    print(f"型號：{result['spec']}")
    print(f"分類：{result['category']}")
    print(f"匹配：{match_icon.get(result['match_type'], '')}")
    print("-" * 50)

    if result["recommended"]:
        print(f"✅ 推荐品牌：{', '.join(result['recommended'])}")
        if result.get("origin") and result["origin"] != "—":
            print(f"   產地：{result['origin']}")
        if result["alternative"]:
            print(f"   備選：{', '.join(result['alternative'])}")
    else:
        print(f"⚠️  {result['notes']}")

    if result["bq"] and result["bq"] not in ("—",):
        print(f"   BQ編號：{result['bq']}")

    print()
    print("📝 報批表填入文本：")
    print(f"   「{result['fill_text']}」")
    print()

    if result.get("similar_items") and result["match_type"] == "fuzzy":
        print("📋 相似項目：")
        for i, sim in enumerate(result["similar_items"], 1):
            print(f"   {i}. {sim['spec']} → {', '.join(sim['brands'])}")
        print()

    print("=" * 50)


def main():
    print("🔍 澳門工程品牌推薦引擎")
    print(f"   知識庫版本：v{BrandRecommender().db['version']}")
    print("   輸入型號或材料名稱查詢（如：13A 插座 / Schneider / 飛利浦）")
    print("   輸入 exit 退出")
    print()

    recommender = BrandRecommender()

    while True:
        try:
            query = input("輸入標書型號：").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n已退出")
            break

        if query.lower() in ("exit", "quit", "q"):
            print("已退出")
            break

        if not query:
            continue

        result = recommender.search(query)
        print_recommendation(result)


if __name__ == "__main__":
    main()
