# Data Adapter 层设计

## 核心思想

所有数据源的差异在 adapter 层消化掉，向上输出统一的中间数据。
填充引擎不再关心数据从哪来。

```
用户 Excel ──→ AutoDetect ──→ TaozhangAdapter ──┐
                               GLAdapter        ├──→ UnifiedData ──→ 填充引擎
                               SimpleAdapter    ─┘
```

## UnifiedData 结构

所有 adapter 输出此格式:

```python
{
  "accounts": {
    "库存现金": {"期末": 1000.0, "期初": 1000.0},
    "银行存款": {"期末": 12500.0, "期初": 11000.0},
    "应收账款": {"期末": 41249020.0, "期初": 30658017.0},
  },
  "balance_sheet": {"货币资金": {"期末": 13500.0, ...}},
  "income_statement": {"营业收入": {"本期": 143768157.0, ...}},
  "cash_flow": {"现金期末余额": {"金额": 13500.0}},
  "entity": {"name": "保定吉达...", "uscc": "..."},
}
```

## 实现顺序

1. UnifiedData 数据模型 + BaseAdapter 接口
2. TaozhangAdapter（把当前提取逻辑搬进去）
3. GLAdapter（科目余额表）
4. SimpleAdapter（仅三大主表的情况）
5. AutoDetect（自动选适配器）
6. 改 fill_notes.py 对接 UnifiedData
