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

## 架构

```
                  ┌──────────────────────────┐
                  │  智能适配层 (新增)         │
                  │  template_analyzer.py     │
                  │  excel_profiler.py        │
                  │  smart_matcher.py         │
                  │  config_generator.py      │
                  └──────────┬───────────────┘
                             │ 分析后自动生成
                  ┌──────────▼───────────────┐
                  │  config/*.yaml            │
                  │  外部配置文件               │
                  └──────────┬───────────────┘
                             │ 加载
决算套表 Excel ──►  填充引擎  ──► 已填充附注 Word
                    (fill_notes.py)
```

三层架构：
- **智能适配层**：上传任意套表和模板，自动分析两边结构并生成配置文件
- **配置层**：YAML 格式，可手动调整、按企业复用
- **填充引擎层**：执行实际填充，不关心配置来源

## 模块说明

| 模块 | 功能 |
|------|------|
| `auto_adapt.py` | 一键入口：模板分析 + Excel 分析 + 匹配 + 配置生成 |
| `template_analyzer.py` | 分析 Word 模板结构：表结构、行标签、列头、合计行、占位符 |
| `excel_profiler.py` | 分析 Excel 套表：Sheet 分类、列类型识别、科目名提取 |
| `smart_matcher.py` | 智能匹配：表名/列模式/行标签三维评分，自动生成映射关系 |
| `config_generator.py` | 生成 YAML 配置文件 + 适配报告，支持从已有配置学习 |
| `config_loader.py` | 运行时加载 YAML 配置 |
| `fill_notes.py` | 填充引擎核心：行匹配、列映射、合计计算、主表稽核 |

## 智能适配用法

```bash
# 安装依赖
pip install pandas openpyxl python-docx pyyaml

# 一键适配：上传模板和套表，自动生成配置
python auto_adapt.py --template 模板路径.docx --excel 套表路径.xlsx

# 指定输出目录
python auto_adapt.py --template 模板.docx --excel 套表.xlsx --output ./my_config

# 加载历史配置进行学习（适配效果逐渐提升）
python auto_adapt.py --template 模板.docx --excel 套表.xlsx --existing ./config

# 输出 JSON 格式结果
python auto_adapt.py --template 模板.docx --excel 套表.xlsx --json
```

## 填充引擎用法

```bash
# 使用已有配置执行填充
python fill_notes.py
# 在 fill_notes.py 末尾修改:
# - 决算套表路径
# - 附注模板路径
# - 输出路径
# - config_dir（默认 "config"）
```

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

## 验证状态

- 填充引擎：45 张附注表自动填充，349 个单元格，33 项合计自动计算，4 项主表稽核通过
- 智能适配：已完成 4 个模块开发，需用实际模板+套表验证匹配效果
- 配置外置：mappings.yaml + name_synonyms.yaml 已从代码抽离

## License

MIT
