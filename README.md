# Audit Note Filler v2 — 配置驱动的附注填充引擎

从决算套表 Excel 自动提取数据，填入审计报告附注 Word 模板。所有映射规则外置为 YAML 配置文件，更换数据源和模板格式只需改配置，不改代码。

## 设计

引擎逻辑（行匹配、列映射、合计计算、主表稽核）和业务规则（哪张 sheet 的哪一列填到模板的哪个位置）完全分离。

```
                  ┌─────────────────────┐
                  │  config/mappings.yaml │
                  │  config/name_synonyms│
                  │  config/auto_sum_*   │
                  └─────────┬───────────┘
                            │ 加载
决算套表 Excel ──►  填充引擎  ──► 已填充附注 Word
                            │
                  ┌─────────┴───────────┐
                  │  match_name()       │
                  │  extract_tz_data()  │
                  │  fill_table()       │
                  │  auto_sum_totals()  │
                  │  validate_sums()    │
                  │  reconcile()        │
                  └─────────────────────┘
```

## 和 v1 的区别

| | v1（原型） | v2（通用版） |
|---|---|---|
| 映射规则 | 硬编码在 fill_notes.py | config/*.yaml 外部配置 |
| 换一家企业 | 改代码 | 换配置文件 |
| 死代码 | SEMANTIC_TABLES 等未使用常量 | 已清理 |
| 定位 | 单项目验证 | 可复用的工具框架 |

## 配置结构

```
config/
├── mappings.yaml          # 46 张附注表的映射规则
├── name_synonyms.yaml     # 76 组名称同义词
├── semantic_tables.yaml   # 会计语义定义（参考文档）
├── cross_validations.yaml # 跨表勾稽关系
├── concept_map.yaml       # 会计概念映射
└── auto_sum_rules.yaml    # 合计规则/税率/稽核配置
```

### mappings.yaml 格式示例

```yaml
- cat: "货币资金-明细"
  sheet_kw: "货币资金_原始数据"
  table_idx: 5
  tz_cols: [3, 4]
  col_map:
    - [0, 1, false]   # 套表第0列 → 模板第1列
    - [1, 2, false]   # 套表第1列 → 模板第2列

- cat: "应收账款-账龄"
  sheet_kw: "应收款项计提坏账准备情况表_原始数据"
  table_idx: 23
  tz_cols: [3, 5, 8, 10]
  accumlate: true
  aggregations:
    - target: "1年以内（含1年）"
      sources: ["1年以内", "1年以内（含1年）"]
      op: "sum"
```

## 使用方式

```bash
# 安装依赖
pip install pandas openpyxl python-docx pyyaml

# 运行
python fill_notes.py

# 在 fill_notes.py 末尾修改:
# - 决算套表路径
# - 附注模板路径
# - 输出路径
# - config_dir（默认 "config"）
```

## 自定义适配

换一家企业的套表和模板，做两件事：

1. 把 config/ 目录复制一份，比如 config/other_company/
2. 修改 config/other_company/ 里的 mappings.yaml 和 name_synonyms.yaml，把 sheet 名、列索引、行名对应关系改成新企业的

然后在代码里传入 `config_dir="config/other_company"` 即可。

## 已验证的能力

45 张附注表自动填充，349 个单元格，33 项合计自动计算。4 项主表稽核通过（货币资金、应收账款、固定资产、营业收入）。

## License

MIT
