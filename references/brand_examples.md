# brand-recommender 使用示例

## 测试用例

### Test 1: 精确匹配 - 13A 插座
```
输入: 13A 插座
期望: 施耐德
```

### Test 2: 精确匹配 - LED灯盘
```
输入: LED燈盤
期望: 飛利浦
```

### Test 3: 模糊匹配 - Schneider开关
```
输入: Schneider 開關掣
期望: 施耐德 (fuzzy match)
```

### Test 4: 分类匹配 - MCB
```
输入: MCB 微型断路器
期望: 按類別推薦 (EL_electrical)
```

### Test 5: 分类匹配 - 抽氣扇
```
输入: 抽氣扇
期望: KDK / 金羚 / 美的
```

### Test 6: 无匹配
```
输入: XXX特種閥門
期望: （待確認）
```

---

## 命令行测试脚本

```bash
cd ~/.項目負責人/skills/brand-recommender
python -c "
from brand_recommender import BrandRecommender, print_recommendation

r = BrandRecommender()
tests = [
    '13A 插座',
    'LED燈盤',
    'Schneider 開關掣',
    'MCB',
    '抽氣扇',
    'XXX特種閥門',
]
for t in tests:
    result = r.search(t)
    print_recommendation(result)
"
```
