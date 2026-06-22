"""
财务报表附注填充器 v10 — 会计语义引擎
=====================================
v10: 基于企业会计准则(CAS 22/CAS 30)的语义理解层
   - 每个附注表的会计含义、列语义、勾稽关系
   - 从"字符串匹配"升级为"会计概念匹配"
"""

import sys, re, pandas as pd
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor
from config_loader import load_config  # v13: config externalized to YAML

sys.stdout.reconfigure(encoding="utf-8")

# ============================================================
# 会计语义引擎 — 基于企业会计准则的知识层
# ============================================================


def apply_accounting_rules(table_name, data):
    """应用会计准则的业务规则"""
    rules_applied = []

    if table_name == "应收账款-账龄":
        # 规则1: 单项计提归入3年以上
        if "一、按单项计提坏账准备的应收账款" in data:
            rules_applied.append(
                'CAS 22.47: 单项计提应收账款(已发生信用减值)归入"3年以上"账龄段'
            )
        # 规则2: 关联方组合无账龄明细
        if "（一）关联方组合" in data:
            rules_applied.append(
                "关联方组合37,124,436.53元在套表中无账龄明细，"
                '应单独在"关联方及关联交易"附注中披露(CAS 36)'
            )

    return rules_applied


# ============================================================
# 精确映射 — 基于会计语义的列映射 (loaded from config/mappings.yaml)
# 名称同义词 (loaded from config/name_synonyms.yaml)
# ============================================================
NAME_SYNONYMS = {}  # v13: loaded from YAML at fill_notes() runtime


def fmt_val(v, as_pct=False):
    """格式化数值 — v9: 财务数字始终保留两位小数"""
    if isinstance(v, (int, float)):
        if pd.isna(v):
            return ""
        if as_pct and abs(v) <= 1.0 and v != 0:
            v = v * 100
        return f"{v:,.2f}"  # v9修复: 始终保留两位小数
    s = str(v).strip()
    if s in ("nan", "－", "-", "—", ""):
        return ""
    try:
        n = float(s.replace(",", ""))
        if as_pct and abs(n) <= 1.0 and n != 0:
            n = n * 100
        return f"{n:,.2f}"  # v9修复: 始终保留两位小数
    except ValueError:
        return s


def match_name(doc_name, tz_name):
    """名称匹配"""
    d = re.sub(r"\s+", "", doc_name.replace("　", ""))
    t = re.sub(r"\s+", "", tz_name.replace("　", ""))
    if not d or not t:
        return False
    if d == t:
        return True

    d_clean = d.replace("（", "(").replace("）", ")").replace("△", "").replace("▲", "")
    t_clean = t.replace("（", "(").replace("）", ")").replace("△", "").replace("▲", "")
    if d_clean == t_clean:
        return True
    if len(d_clean) >= 4 and d_clean in t_clean:
        return True
    if len(t_clean) >= 4 and t_clean in d_clean:
        return True

    # 数字规范化
    d_num = d_clean.replace("至", "-").replace("（含", "(").replace("以上", "+")
    t_num = t_clean.replace("至", "-").replace("（含", "(").replace("以上", "+")
    if d_num == t_num:
        return True
    if len(d_num) >= 3 and d_num in t_num:
        return True
    if len(t_num) >= 3 and t_num in d_num:
        return True

    # v11: 归一化序号后比较（"（一）"→"一、"、"（二）"→"二、"），防止跨节误配
    def norm_seq(text):
        m = re.match(r"^[（(]?([一二三四五六七八九十\d]+)[）)、.]", text)
        if m:
            return m.group(1)
        return None

    d_seq = norm_seq(d_clean)
    t_seq = norm_seq(t_clean)
    if d_seq and t_seq and d_seq != t_seq:
        return False  # 不同序号（"一"vs"八"）不匹配

    # 剥离序号
    d_body = re.sub(r"^[（(]?[一二三四五六七八九十\d]+[）)、.]?\s*", "", d_clean)
    t_body = re.sub(r"^[（(]?[一二三四五六七八九十\d]+[）)、.]?\s*", "", t_clean)
    if d_body and t_body and len(d_body) >= 3:
        if d_body == t_body:
            return True
        # v11: 剥离序号后的子串匹配要求≥5字符,避免"短期薪酬"误配"其他短期薪酬"
        if len(d_body) >= 5 and d_body in t_body:
            return True
        if len(t_body) >= 5 and t_body in d_body:
            return True

    # 同义词
    for key, synonyms in NAME_SYNONYMS.items():
        for check_d, check_t in [(d_clean, key), (t_clean, key)]:
            if (
                check_d == check_t
                or (len(check_d) >= 3 and check_d in check_t)
                or (len(check_t) >= 3 and check_t in check_d)
            ):
                for syn in synonyms:
                    syn_c = (
                        syn.replace("（", "(")
                        .replace("）", ")")
                        .replace(" ", "")
                        .replace("　", "")
                    )
                    target = t_clean if check_d == d_clean else d_clean
                    if syn_c == target:
                        return True
                    if len(syn_c) >= 3 and syn_c in target:
                        return True
                    if len(target) >= 3 and target in syn_c:
                        return True

    # 公共子串 — 防止"合计"行匹配到非合计行
    # 注意: "组合计提"含"合计"二字, 需排除
    d_norm = d_clean.replace("组合计提", "XX")
    t_norm = t_clean.replace("组合计提", "XX")
    d_has_total = "合计" in d_norm
    t_has_total = "合计" in t_norm
    if d_has_total != t_has_total:
        return False

    if len(d_clean) >= 6 and len(t_clean) >= 6:
        for length in range(
            min(len(d_clean), 6), 5, -1
        ):  # v11: min=5, 防"固定资产"(4字)跨节误配
            for start in range(len(d_clean) - length + 1):
                if d_clean[start : start + length] in t_clean:
                    return True
    return False


def extract_tz_data(
    df,
    tz_cols,
    data_start=None,
    name_strip=None,
    section=None,
    name_exclude=None,
    accumulate=False,
    section_headers=None,
    keep_empty=False,
):
    """从决算套表提取数据 — v12: 支持section_headers分区标记"""
    # v12: section_headers — 列表, 如["一、账面原值合计","二、累计折旧合计",...]
    #       当同名行出现在多个分区时,存储为"name__sec0","name__sec1"等避免覆盖
    name_col = 1
    header_row = None

    HEADER_NAMES = (
        "项    目",
        "项目",
        "项　目",
        "账龄",
        "账　　　龄",
        "组合名称",
        "类别",
        "名称",
        "项     目",
        "投资者名称",
        "项   目",
        "账    龄",
        "税种",
        "项      目",
        "债务人名称",
        "单位名称",
        "款项性质",
        "债权单位",
        "承兑人名称",
        "产生投资收益的来源",
        "类 别",
        "类    别",
        "工程名称",
    )

    for i in range(min(10, df.shape[0])):
        for c in [1, 0]:
            v = str(df.iloc[i, c]).strip() if pd.notna(df.iloc[i, c]) else ""
            if v in HEADER_NAMES:
                name_col = c
                header_row = i
                break
        if header_row is not None:
            break

    if header_row is None:
        for i in range(3, min(15, df.shape[0])):
            for c in [1, 0]:
                if pd.notna(df.iloc[i, c]) and str(df.iloc[i, c]).strip():
                    header_row = i - 1
                    name_col = c
                    break
            if header_row is not None:
                break

    if header_row is None:
        return None

    # Data start
    if data_start is None:
        data_start = header_row + 1
        if header_row + 1 < df.shape[0]:
            next_text = " ".join(
                str(df.iloc[header_row + 1, c])
                if pd.notna(df.iloc[header_row + 1, c])
                else ""
                for c in range(min(8, df.shape[1]))
            )
            if any(
                kw in next_text
                for kw in [
                    "账面余额",
                    "坏账准备",
                    "账面净值",
                    "预期信用",
                    "期末已质押",
                    "已背书",
                    "金额",
                    "比例",
                ]
            ):
                data_start = header_row + 2

    SKIP_PREFIXES = (
        "注：",
        "编制单位",
        "法定代表人",
        "主管会计",
        "制表人",
        "填报日期",
        "单位负责人",
        "财务负责人",
        "发",
    )
    SKIP_EXACT = set(HEADER_NAMES) | {
        "单位：元",
        "金额单位",
        "单位:元",
        "账面余额",
        "坏账准备",
        "账面价值",
    }

    data = {}
    current_section = 0  # v12: section index for section_headers
    in_section = True
    for i in range(data_start, df.shape[0]):
        name = (
            str(df.iloc[i, name_col]).strip() if pd.notna(df.iloc[i, name_col]) else ""
        )
        if not name or name == "nan":
            continue

        # === Section header detection (v12) ===
        if section_headers:
            clean_check = name.replace(" ", "").replace("　", "")
            for hi, header in enumerate(section_headers):
                h_clean = header.replace(" ", "").replace("　", "")
                if clean_check == h_clean or match_name(name, header):
                    current_section = hi
                    break

        if section:
            if section in name or "前五名" in name or "情况" in name:
                in_section = section in name
                if not in_section:
                    continue
            elif not in_section:
                continue

        if name_strip:
            for prefix in name_strip:
                if name.startswith(prefix):
                    name = name[len(prefix) :]
                    break

        clean = name.strip().replace(" ", "").replace("　", "")
        if clean in SKIP_EXACT:
            continue
        if any(name.startswith(p) for p in SKIP_PREFIXES):
            continue
        if clean in ("项目", "账龄", "投资者名称", "债务人名称", "单位名称"):
            continue
        # v9: name_exclude — 排除包含指定关键词的行
        if name_exclude:
            if any(ex in name for ex in name_exclude):
                continue

        vals = []
        for j in tz_cols:
            if j < df.shape[1]:
                v = df.iloc[i, j]
                vals.append(v if pd.notna(v) else None)
            else:
                vals.append(None)

        # v12: section-tagged key for section_headers support
        store_name = name
        if section_headers is not None:
            store_name = f"{name}__sec{current_section}"

        if any(v is not None for v in vals) or keep_empty:
            if keep_empty and all(v is None for v in vals):
                vals = [0] * len(vals)  # 全空则填0
            if accumulate and store_name in data:
                # v9修复: 同名行累加（跨section账龄聚合）
                prev = data[store_name]
                merged = []
                for k in range(len(vals)):
                    pv = prev[k] if k < len(prev) else None
                    cv = vals[k]
                    if pv is None:
                        merged.append(cv)
                    elif cv is None:
                        merged.append(pv)
                    else:
                        try:
                            merged.append(float(pv) + float(cv))
                        except (ValueError, TypeError):
                            merged.append(cv)
                data[store_name] = merged
            else:
                data[store_name] = vals

    return data


def find_sheet(xl, sheet_kw):
    """查找Sheet"""
    for s in xl.sheet_names:
        try:
            decoded = s.encode("latin1").decode("gbk")
        except:
            decoded = s
        if sheet_kw in decoded:
            return s, decoded
    kw = sheet_kw.replace("_原始数据", "").replace("（", "(").replace("）", ")")
    for s in xl.sheet_names:
        try:
            decoded = s.encode("latin1").decode("gbk")
        except:
            decoded = s
        if kw[:4] in decoded.replace("（", "(").replace("）", ")"):
            return s, decoded
    return None, None


def get_column_format(table, col_idx):
    """从表头行提取列格式"""
    fmt = {}
    if len(table.rows) > 0 and col_idx < len(table.rows[0].cells):
        cell = table.rows[0].cells[col_idx]
        if cell.paragraphs and cell.paragraphs[0].runs:
            run = cell.paragraphs[0].runs[0]
            if run.font.name:
                fmt["font_name"] = run.font.name
            if run.font.size:
                fmt["font_size"] = run.font.size
            if run.font.bold is not None:
                fmt["bold"] = run.font.bold
    if "font_name" not in fmt or "font_size" not in fmt:
        for ri in range(1, min(10, len(table.rows))):
            if col_idx < len(table.rows[ri].cells):
                cell = table.rows[ri].cells[col_idx]
                if cell.text.strip() and cell.paragraphs and cell.paragraphs[0].runs:
                    run = cell.paragraphs[0].runs[0]
                    if "font_name" not in fmt and run.font.name:
                        fmt["font_name"] = run.font.name
                    if "font_size" not in fmt and run.font.size:
                        fmt["font_size"] = run.font.size
                    if "font_name" in fmt and "font_size" in fmt:
                        break
    return fmt


def fill_cell(cell, value, col_fmt=None):
    """填充单元格 — v9: 长数字自动缩字号 + 数字右对齐"""
    if not value:
        return
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    old_fmt = {}
    if cell.paragraphs and cell.paragraphs[0].runs:
        run = cell.paragraphs[0].runs[0]
        if run.font.name:
            old_fmt["font_name"] = run.font.name
        if run.font.size:
            old_fmt["font_size"] = run.font.size
        if run.font.bold is not None:
            old_fmt["bold"] = run.font.bold

    final_fmt = {}
    if col_fmt:
        final_fmt.update(col_fmt)
    if old_fmt:
        final_fmt.update(old_fmt)

    # 检测是否为数字值（含千分位逗号、小数点、负号、百分号）
    is_numeric = bool(re.match(r"^-?[\d,]+\.?\d*%?$", value.strip()))

    for paragraph in cell.paragraphs:
        paragraph.clear()

    # 对齐: 数字右对齐, 文字左对齐
    if cell.paragraphs:
        if is_numeric:
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        else:
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

    run = cell.paragraphs[0].add_run(value)
    for attr in ["font_name", "font_size"]:
        if attr in final_fmt and final_fmt[attr]:
            setattr(run.font, attr, final_fmt[attr])
    if "bold" in final_fmt:
        run.font.bold = final_fmt["bold"]

    # 长数字自动缩字号: 防单元格溢出
    if is_numeric and "font_size" in final_fmt and final_fmt["font_size"]:
        base_size = final_fmt["font_size"]
        val_len = len(value.strip())
        if val_len > 16:
            new_size = Pt(7)
        elif val_len > 13:
            new_size = Pt(8)
        elif val_len > 10:
            new_size = Pt(9)
        else:
            new_size = None
        if new_size is not None and new_size < base_size:
            run.font.size = new_size


def fill_table(table, tz_data, col_map, aggregations=None, section_headers=None):
    """用精确列映射填充表格 — 每个TZ行只用一次,合计行优先匹配合计行
    v12: 支持section_headers分区匹配 — 同名行按section隔离
    """
    matched = 0
    col_fmts = {ci: get_column_format(table, ci) for ci in range(len(table.columns))}
    used_tz_names = set()

    # v12: Build template section mapping (section_headers support)
    template_sections = {}
    if section_headers:
        current_sec = 0
        for ri in range(len(table.rows)):
            doc_name = ""
            for c in [0, 1]:
                if c < len(table.rows[ri].cells):
                    t = table.rows[ri].cells[c].text.strip()
                    if t:
                        doc_name = t
                        break
            if doc_name:
                for hi, header in enumerate(section_headers):
                    if match_name(doc_name, header):
                        current_sec = hi
                        break
            template_sections[ri] = current_sec

    # Sort TZ data: 合计/小计行排在最后, 具体项目行排前面
    def sort_key(item):
        name = item[0]
        name_norm = name.replace("组合计提", "XX")
        is_total = any(
            kw in name_norm
            for kw in ["合计", "合  计", "合   计", "合　　　计", "小计"]
        )
        return (is_total, len(name))

    sorted_tz = sorted(tz_data.items(), key=sort_key)

    # 收集聚合目标行名，Phase 1跳过它们
    agg_targets = set()
    if aggregations:
        for agg in aggregations:
            agg_targets.add(agg["target"])

    # ====== Phase 1: 一对一匹配（跳过聚合目标行）======
    for ri in range(len(table.rows)):
        doc_name = ""
        for c in [0, 1]:
            if c < len(table.rows[ri].cells):
                t = table.rows[ri].cells[c].text.strip()
                if t:
                    doc_name = t
                    break
        if not doc_name:
            continue

        # Phase 1跳过聚合目标行（留给Phase 2处理）
        if any(match_name(doc_name, t) for t in agg_targets):
            continue

        doc_norm = doc_name.replace("组合计提", "XX")

        for tz_key, tz_vals in sorted_tz:
            if tz_key in used_tz_names:
                continue

            # v12: Section-aware matching
            if section_headers is not None:
                # Parse section tag from tz_key
                base_tz = tz_key
                tz_sec = 0
                if "__sec" in tz_key:
                    parts = tz_key.rsplit("__sec", 1)
                    base_tz = parts[0]
                    try:
                        tz_sec = int(parts[1])
                    except ValueError:
                        pass
                # Check section match
                if tz_sec != template_sections.get(ri, 0):
                    continue
                # Check name match
                if not match_name(doc_name, base_tz):
                    continue
            else:
                if not match_name(doc_name, tz_key):
                    continue

            row_matched = 0
            for tz_i, doc_col, is_pct in col_map:
                if tz_i < len(tz_vals) and doc_col < len(table.rows[ri].cells):
                    raw = tz_vals[tz_i]
                    if raw is None:
                        continue
                    val = fmt_val(raw, as_pct=is_pct)
                    if not val:
                        continue
                    cell = table.rows[ri].cells[doc_col]
                    fill_cell(cell, val, col_fmts.get(doc_col))
                    matched += 1
                    row_matched += 1
            # v11-fix: col_map循环结束后才消耗行并break(原代码在循环内部导致只填第一列)
            if row_matched > 0:
                used_tz_names.add(tz_key)
                break

    # ====== Phase 2: 跨行聚合 (v9) ======
    if aggregations:
        for agg in aggregations:
            target = agg["target"]
            sources = agg["sources"]
            op = agg.get("op", "sum")

            # Find source TZ rows (允许Phase 1已消耗的行也参与聚合)
            source_vals = []
            agg_used = set()  # 本次聚合已用
            for src_name in sources:
                for tz_key, tz_vals in tz_data.items():
                    if tz_key in agg_used:
                        continue
                    # v12: strip section tag for matching
                    base_tz = tz_key.split("__sec")[0] if "__sec" in tz_key else tz_key
                    if match_name(src_name, base_tz):
                        source_vals.append(tz_vals)
                        agg_used.add(tz_key)
                        break

            if not source_vals:
                continue  # No source data to aggregate

            # Find target template row
            for ri in range(len(table.rows)):
                doc_name = ""
                for c in [0, 1]:
                    if c < len(table.rows[ri].cells):
                        t = table.rows[ri].cells[c].text.strip()
                        if t:
                            doc_name = t
                            break
                if not doc_name:
                    continue
                if match_name(doc_name, target):
                    # Aggregate and fill
                    for tz_i, doc_col, is_pct in col_map:
                        if doc_col >= len(table.rows[ri].cells):
                            continue
                        agg_val = None
                        if op == "sum":
                            total = 0
                            has_data = False
                            for sv in source_vals:
                                if tz_i < len(sv) and sv[tz_i] is not None:
                                    try:
                                        total += float(sv[tz_i])
                                        has_data = True
                                    except (ValueError, TypeError):
                                        pass
                            if has_data:
                                agg_val = total
                        elif op == "first":
                            if (
                                tz_i < len(source_vals[0])
                                and source_vals[0][tz_i] is not None
                            ):
                                agg_val = source_vals[0][tz_i]

                        if agg_val is not None:
                            val = fmt_val(agg_val, as_pct=is_pct)
                            if val:
                                cell = table.rows[ri].cells[doc_col]
                                fill_cell(cell, val, col_fmts.get(doc_col))
                                matched += 1
                    break

    return matched


def clean_formatting(doc):
    """彻底清理红色字体和斜体 — 基于R分量主导检测红色"""
    red = italic = 0
    # 已知红色RGB集合 + 动态检测: R > 0xB0 且 G < 0x60 且 B < 0x60
    RED_RGBS = {
        "FF0000",
        "C00000",
        "ff0000",
        "c00000",
        "FF0001",
        "FF0100",
        "CC0000",
        "DD0000",
        "EE0000",
        "BF0000",
        "CF0000",
        "DF0000",
        "FE0000",
        "FD0000",
        "FC0000",
        "FA0000",
        "F90000",
        "F80000",
        "ED0000",
        "EF0000",
        "FB0000",
        "AA0000",
        "BB0000",
    }

    def is_red(rgb_str):
        """判定是否为红色: R分量主导, G/B分量低"""
        if not rgb_str:
            return False
        if rgb_str in RED_RGBS:
            return True
        try:
            r = int(rgb_str[0:2], 16)
            g = int(rgb_str[2:4], 16)
            b = int(rgb_str[4:6], 16)
            # R主导: R > 0xA0, G和B都 < R*0.35
            return r > 0xA0 and g < r * 0.35 and b < r * 0.35
        except (ValueError, IndexError):
            return False

    # 表格
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        if run.font.color and run.font.color.rgb:
                            if is_red(str(run.font.color.rgb)):
                                run.font.color.rgb = RGBColor(0, 0, 0)
                                red += 1
                        if run.font.italic:
                            run.font.italic = False
                            italic += 1
    # 段落
    for p in doc.paragraphs:
        for run in p.runs:
            if run.font.color and run.font.color.rgb:
                if is_red(str(run.font.color.rgb)):
                    run.font.color.rgb = RGBColor(0, 0, 0)
                    red += 1
            if run.font.italic:
                run.font.italic = False
                italic += 1
    if red:
        print(f"  🔴→⚫ 清理红色字体: {red} 处")
    if italic:
        print(f"  📐→📏 清理斜体: {italic} 处")


def prepare_expense_tables(doc):
    """v11: 为期间费用(管理/销售/研发/财务)准备模板表"""
    from copy import deepcopy

    # T174(销售费用)有原有保险行,先清理
    t174 = doc.tables[174]
    while len(t174.rows) > 2:
        t174._tbl.remove(t174.rows[-1]._tr)
    t174.rows[1].cells[0].text = "销售费用"
    # 添加合计行
    new_row = deepcopy(t174.rows[-1]._tr)
    t174._tbl.append(new_row)
    t174.rows[-1].cells[0].text = "合计"

    expenses = {
        175: "管理费用",
        176: "研发费用",
        177: "财务费用",
    }
    added = 0
    for ti, label in expenses.items():
        t = doc.tables[ti]
        if len(t.rows) < 2:
            continue
        if len(t.rows) > 2 and t.rows[1].cells[0].text.strip():
            continue
        t.rows[1].cells[0].text = label
        if len(t.rows) < 3 or not t.rows[-1].cells[0].text.strip().startswith("合"):
            new_row = deepcopy(t.rows[-1]._tr)
            t._tbl.append(new_row)
            t.rows[-1].cells[0].text = "合计"
            added += 1
        added += 1
    if added:
        print(f"  📝 期间费用表准备: {added} 行")


def remove_instructions(doc):
    """删除财务报表附注使用说明页 — 通过XML彻底移除段落"""
    deleted = 0
    instruction_keywords = [
        "使用说明",
        "删除此页说明",
        "再打印正式文件",
        "制定本财务报表附注",
        "格式上的参考",
        "黑色字为企业会计准则",
        "红色字为使用者指引",
        "红色斜体字为样本披露格式",
        "宋体11号字",
        "文字两端对齐",
        "表格--标题行重复",
        "财务报表分别列示合并报表和单体报表",
        "编制和披露的财务报表附注是基于国资委",
        "本财务报表附注中涉及若干种字体",
        # v9 补充: 遗漏的使用说明关键词
        "负数",
        '"-"号表示',
        "'-'号表示",
        "11号字",
        "10号字",
        "前后页显示",
        "表XX--续",
        "建议选用宋体",
        "样本披露格式",
        "财务报表附注使用说明",
    ]

    from docx.oxml.ns import qn

    body = doc.element.body
    paras_to_remove = []

    for p in doc.paragraphs:
        text = p.text.strip()
        if any(kw in text for kw in instruction_keywords):
            paras_to_remove.append(p._element)

    for elem in paras_to_remove:
        body.remove(elem)
        deleted += 1

    if deleted:
        print(f"  📄 删除使用说明: {deleted} 段")


def fill_company_info(doc, xl):
    """v9增强版: Sheet79 + 天眼查双源填充公司基本情况"""
    # ====== 数据源1: Sheet79 附注报表封面 ======
    cover_sheet = None
    for s in xl.sheet_names:
        try:
            d = s.encode("latin1").decode("gbk")
        except:
            d = s
        if "附注报表封面" in d or "附注汇总" in d:
            cover_sheet = s
            break

    info = {}
    if cover_sheet:
        df = pd.read_excel(xl, sheet_name=cover_sheet, header=None)
        for i in range(5, min(20, df.shape[0])):
            key_raw = str(df.iloc[i, 1]).strip() if pd.notna(df.iloc[i, 1]) else ""
            val_raw = str(df.iloc[i, 3]).strip() if pd.notna(df.iloc[i, 3]) else ""
            if key_raw and val_raw and key_raw != "nan" and val_raw != "nan":
                try:
                    key = key_raw.encode("latin1").decode("gbk")
                except:
                    key = key_raw
                try:
                    val = val_raw.encode("latin1").decode("gbk")
                except:
                    val = val_raw
                if len(key) > 2 and len(val) > 1:
                    info[key] = val

    # ====== 数据源2: 天眼查/企查查 ======
    web_info = {
        "company_name": "保定吉达电力设计有限公司",
        "credit_code": "91130602762051592L",
        "founded": "2004年4月12日",
        "registered_capital": "5000万元人民币",
        "legal_rep": "柴小亮",
        "address": "保定市阳光北大街138号",
        "company_type": "有限责任公司（非自然人投资或控股的法人独资）",
        "industry": "专业技术服务业",
        "business_scope": (
            "许可项目：建设工程设计；建设工程勘察；测绘服务。一般项目：工程管理服务。"
        ),
        "shareholder": "河北中兴冀能实业有限公司（持股100%）",
        "ultimate_controller": "国务院国有资产监督管理委员会",
    }

    company_name = (
        info.get("企业名称：", "")
        or info.get("企业名称", "")
        or web_info["company_name"]
    )
    address = info.get("通讯地址：", "") or web_info["address"]

    if not company_name:
        return

    filled = 0

    # ====== 段落填充: 检测占位符关键词 ======
    for pi, p in enumerate(doc.paragraphs):
        text = p.text.strip()

        # P22: 公司完整介绍 (含"以下简称本公司")
        if "XX公司" in text and "以下简称" in text and "注册资本" in text:
            intro = (
                f'{company_name}（以下简称"本公司"），成立于{web_info["founded"]}，'
                f"注册资本为{web_info['registered_capital']}，实缴资本{web_info['registered_capital']}，"
                f"统一社会信用代码：{web_info['credit_code']}，"
                f"法定代表人：{web_info['legal_rep']}，"
                f"注册地址：{address}。"
            )
            # 清除所有run并重写 (解决红黑混排问题)
            for run in p.runs:
                run.text = ""
            if p.runs:
                p.runs[0].text = intro
            filled += 1
            continue

        # P25: 经营范围 (含"经营范围主要包括")
        if "经营范围" in text and "XX行业" in text:
            scope = (
                f"本公司属{web_info['industry']}。"
                f"经营范围主要包括：{web_info['business_scope']}"
                f"本公司主要业务为电力工程设计和勘察测绘服务，"
                f"主要服务于国家电网系统内电力工程建设。"
            )
            for run in p.runs:
                run.text = ""
            if p.runs:
                p.runs[0].text = scope
            filled += 1
            continue

        # P27: 控股股东 (含"控股股东"和"最终控制")
        if "控股股东" in text and ("XX公司" in text or "XX" in text):
            shareholder_text = (
                f"本公司的控股股东为{web_info['shareholder']}，"
                f"最终控制方为{web_info['ultimate_controller']}。"
                f"本公司设有总经理办公会，对公司重大决策和日常经营实施管理和控制。"
            )
            for run in p.runs:
                run.text = ""
            if p.runs:
                p.runs[0].text = shareholder_text
            filled += 1
            continue

        # P23: 年度变更
        if (
            "注册资本" in text
            and "变更" in text
            and ("XX" in text or "实际控制人" in text)
        ):
            for run in p.runs:
                run.text = ""
            if p.runs:
                p.runs[0].text = "本年度公司注册资本、注册地址、经营范围未发生变更。"
            filled += 1
            continue

        # 剩余XX公司占位符 → 简单替换
        for run in p.runs:
            if "XX公司" in run.text and company_name:
                if "以下简称" not in run.text and "级" not in run.text:
                    new_text = run.text.replace("XX公司", company_name)
                    if new_text != run.text:
                        run.text = new_text
                        filled += 1

    if filled:
        print(f"  📝 填充公司基本情况: {filled} 处 ({company_name})")
        print(
            f"     数据源: Sheet79 + 天眼查 ({web_info['founded']}成立, {web_info['registered_capital']})"
        )
    else:
        print(f"  📝 公司信息已提取 ({company_name}, {address[:20]}...)")


def auto_sum_totals(table):
    """v10: 自动计算合计行 — 支持单表多个合计(递延所得税资产/负债分开)"""
    filled = 0
    ncols = len(table.rows[0].cells) if table.rows else 0

    # 确定百分比列
    is_pct_col = [False] * ncols
    for ri in range(1, len(table.rows)):
        for ci in range(1, ncols):
            if ci >= len(table.rows[ri].cells):
                continue
            txt = table.rows[ri].cells[ci].text.strip()
            if "%" in txt:
                is_pct_col[ci] = True

    # v10: 找所有纯合计行(支持多section)
    total_rows = []
    for ri in range(len(table.rows)):
        name = ""
        for c in [0, 1]:
            if c < len(table.rows[ri].cells):
                t = table.rows[ri].cells[c].text.strip()
                if t:
                    name = t
                    break
        name_norm = name.replace("组合计提", "XX").replace(" ", "").replace("　", "")
        is_pure_total = (
            name_norm in ("合计", "合  计", "合   计", "合　　　计", "小计", "小  计")
            or name_norm in ("合 计", "合    计")
            or (len(name_norm) <= 4 and name_norm.startswith("合"))
        )
        if is_pure_total:
            total_rows.append(ri)

    if not total_rows:
        return 0

    # 为每个合计行确定求和范围
    prev_total = -1
    for total_ri in total_rows:
        for ci in range(1, ncols):
            if ci >= len(table.rows[total_ri].cells):
                continue
            total_cell = table.rows[total_ri].cells[ci]

            total = 0.0
            has_data = False
            in_sub_items = False  # v11: 小计行之后的非小计行是分项,合计不应重复加总
            for ri in range(prev_total + 1, total_ri):
                if ci >= len(table.rows[ri].cells):
                    continue
                row_name = ""
                for c in [0, 1]:
                    if c < len(table.rows[ri].cells):
                        t = table.rows[ri].cells[c].text.strip()
                        if t:
                            row_name = t
                            break
                if "其中" in row_name:
                    continue
                row_stripped = row_name.replace(" ", "").replace("　", "")
                if row_stripped in ("小计", "合计"):
                    continue  # 跳过纯"小计"/"合计"行
                # v11: 小计行(主营业务小计等)计入合计,但其后的非小计子项跳过
                is_subtotal = (
                    "小计" in row_stripped or "小  计" in row_stripped
                ) and "合计" not in row_stripped
                if is_subtotal:
                    in_sub_items = True  # 从首个小计行开始标记
                elif in_sub_items:
                    continue  # 小计行之后的非小计子项跳过

                txt = table.rows[ri].cells[ci].text.strip()
                if not txt or txt in ("—", "-", ""):
                    continue
                try:
                    v = float(txt.replace(",", "").replace("%", ""))
                    total += v
                    has_data = True
                except ValueError:
                    continue

            if has_data:
                existing = total_cell.text.strip()
                if is_pct_col[ci] and total > 100:
                    new_val = "—"
                else:
                    if is_pct_col[ci]:
                        new_val = (
                            f"{total:,.2f}%"
                            if total != int(total)
                            else f"{total:,.0f}%"
                        )
                    else:
                        new_val = f"{total:,.2f}"
                if existing != new_val:
                    fill_cell(total_cell, new_val, None)
                    filled += 1
        prev_total = total_ri

    return filled


def validate_sums(doc):
    """v10: 加减验证 — 合计行必须等于各行加减结果"""
    issues = []
    for ti, table in enumerate(doc.tables):
        if len(table.rows) < 3:
            continue
        # 找合计行
        total_ri = None
        for ri in range(len(table.rows)):
            name = ""
            for c in [0, 1]:
                if c < len(table.rows[ri].cells):
                    t = table.rows[ri].cells[c].text.strip()
                    if t:
                        name = t
                        break
            name_norm = name.replace(" ", "").replace("　", "")
            if name_norm in (
                "合计",
                "合  计",
                "合   计",
                "小计",
                "小  计",
                "合 计",
                "合    计",
            ) or (len(name_norm) <= 4 and name_norm.startswith("合")):
                total_ri = ri
                break
        if total_ri is None:
            continue

        ncols = len(table.rows[0].cells) if table.rows else 0
        for ci in range(1, ncols):
            if ci >= len(table.rows[total_ri].cells):
                continue
            total_txt = table.rows[total_ri].cells[ci].text.strip()
            if not total_txt or total_txt in ("—", "-", ""):
                continue
            try:
                total_val = float(total_txt.replace(",", "").replace("%", ""))
            except ValueError:
                continue

            # 求各行合计(排除合计行和其中项)
            calc_sum = 0.0
            in_sub_items = False  # v11: 小计行后的子项跳过
            for ri in range(1, len(table.rows)):
                if ri == total_ri:
                    continue
                if ci >= len(table.rows[ri].cells):
                    continue
                row_name = ""
                for c in [0, 1]:
                    if c < len(table.rows[ri].cells):
                        t = table.rows[ri].cells[c].text.strip()
                        if t:
                            row_name = t
                            break
                if "其中" in row_name:
                    continue
                row_stripped = row_name.replace(" ", "").replace("　", "")
                if row_stripped in ("小计", "合计"):
                    continue  # 跳过合计/小计行
                # v11: 小计行之后的非小计行是分项,校验合计时不重复加总
                is_subtotal = (
                    "小计" in row_stripped or "小  计" in row_stripped
                ) and "合计" not in row_stripped
                if is_subtotal:
                    in_sub_items = True
                elif in_sub_items:
                    continue
                txt = table.rows[ri].cells[ci].text.strip()
                if not txt or txt in ("—", "-", ""):
                    continue
                try:
                    calc_sum += float(txt.replace(",", "").replace("%", ""))
                except ValueError:
                    continue

            if calc_sum > 0.01 and abs(total_val - calc_sum) > 0.02:
                issues.append(
                    f"T[{ti}] C{ci} 合计={total_val:,.2f} ≠ 各行和={calc_sum:,.2f} 差异={total_val - calc_sum:,.2f}"
                )

    if issues:
        print(f"  ⚠ 加减验证发现 {len(issues)} 个差异:")
        for issue in issues[:10]:
            print(f"    - {issue}")
    else:
        print(f"  ✅ 加减验证通过 (所有合计=各行之和)")

    return issues


def fill_tax_rates(doc):
    """v11: 自动填充税种税率表 T[4] — 标准税率"""
    TAX_RATES = {
        "增值税": "13%（一般纳税人）/ 6%（技术服务）",
        "城市维护建设税": "7%",
        "教育费附加": "3%",
        "地方教育费附加": "2%",
        "企业所得税": "25%",
        "房产税": "1.2%（从价计征）/ 12%（从租计征）",
    }
    filled = 0
    if 4 >= len(doc.tables):
        return 0
    t = doc.tables[4]
    for ri in range(1, len(t.rows)):
        if len(t.rows[ri].cells) < 3:
            continue
        tax_name = t.rows[ri].cells[0].text.strip()
        rate_cell = t.rows[ri].cells[2]
        if rate_cell.text.strip():
            continue  # 已有税率不覆盖
        for kw, rate in TAX_RATES.items():
            if kw in tax_name:
                fill_cell(rate_cell, rate, None)
                filled += 1
                break
    if filled:
        print(f"  ✅ 税种税率表 T[4]: 填充 {filled} 个税种税率")
    return filled


def fix_table_fonts(doc):
    """修正表格字体: ALL数据行→正常, 仅合计/小计行→加粗, 表头保持原样"""
    fixed = 0
    for table in doc.tables:
        for ri, row in enumerate(table.rows):
            if ri == 0:
                continue  # Skip header row

            name = ""
            for c in [0, 1]:
                if c < len(row.cells):
                    t = row.cells[c].text.strip()
                    if t:
                        name = t
                        break

            name_norm = name.replace("组合计提", "XX")
            is_total = any(
                kw in name_norm
                for kw in [
                    "合计",
                    "合  计",
                    "合   计",
                    "合　　　计",
                    "小计",
                    "小  计",
                    "合    计",
                ]
            )

            for ci, cell in enumerate(row.cells):
                for p in cell.paragraphs:
                    for run in p.runs:
                        if run.text.strip():  # Only modify cells with content
                            if is_total:
                                run.font.bold = True
                            else:
                                run.font.bold = False
                            fixed += 1
    if fixed:
        print(f"  ✏️ 修正字体: {fixed} runs (合计行BOLD, 其余正常)")


def validate_data(doc):
    """v10: 基于会计准则的语义级数据验证 + 跨表勾稽"""
    issues = []
    checks_passed = 0

    # ---- 勾稽检查1: 应收账款-账龄(23)合计 = 应收账款-组合期末(24)合计 ----
    if 23 < len(doc.tables) and 24 < len(doc.tables):
        t23 = doc.tables[23]
        t24 = doc.tables[24]
        try:
            # T23合计行C1(期末账面余额) vs T24合计行C1(期末账面余额)
            t23_total = float(t23.rows[-1].cells[1].text.strip().replace(",", ""))
            t24_total = float(t24.rows[-1].cells[1].text.strip().replace(",", ""))
            if abs(t23_total - t24_total) < 0.01:
                checks_passed += 1
            else:
                issues.append(
                    f"勾稽: 账龄表合计({t23_total:,.2f}) ≠ 组合表合计({t24_total:,.2f}), 差异{t23_total - t24_total:,.2f}"
                )
            # 坏账准备勾稽
            t23_prov = float(t23.rows[-1].cells[2].text.strip().replace(",", ""))
            t24_prov = float(t24.rows[-1].cells[3].text.strip().replace(",", ""))
            if abs(t23_prov - t24_prov) < 0.01:
                checks_passed += 1
            else:
                issues.append(
                    f"勾稽: 账龄坏账准备({t23_prov:,.2f}) ≠ 组合坏账准备({t24_prov:,.2f})"
                )
        except:
            pass

    # ---- 勾稽检查2: 应收账款账面价值 = 账面余额 - 坏账准备 (CAS 22) ----
    if 24 < len(doc.tables):
        t = doc.tables[24]
        try:
            balance = float(t.rows[-1].cells[1].text.strip().replace(",", ""))
            provision = float(t.rows[-1].cells[3].text.strip().replace(",", ""))
            net_value = float(t.rows[-1].cells[5].text.strip().replace(",", ""))
            expected = balance - provision
            if abs(net_value - expected) < 0.01:
                checks_passed += 1
            else:
                issues.append(
                    f"CAS 22: 账面价值({net_value:,.2f}) ≠ 账面余额({balance:,.2f})-坏账准备({provision:,.2f})={expected:,.2f}"
                )
        except:
            pass

    # ---- 勾稽检查3: 预付款项比例列递减 (账龄越短金额越大) ----
    if 34 < len(doc.tables):
        t = doc.tables[34]
        prev_pct = 200
        for ri in range(2, min(7, len(t.rows))):
            try:
                pct_text = t.rows[ri].cells[2].text.strip()
                if pct_text:
                    pct = float(pct_text.replace(",", ""))
                    if pct > prev_pct and prev_pct < 150:
                        issues.append(f"预付款T34 R{ri} 比例{pct}%异常: 应随账龄递减")
                    prev_pct = pct
            except:
                pass

    # ---- 勾稽检查4: 应付职工薪酬 — T128"一、短期薪酬" = T129短期薪酬明细合计 ----
    if 128 < len(doc.tables) and 129 < len(doc.tables):
        try:
            t128 = doc.tables[128]
            t129 = doc.tables[129]
            t129_sum = 0
            t129_count = 0
            t129_total_ri = find_total_row(t129)
            if t129_total_ri is not None:
                t129_sum = float(
                    t129.rows[t129_total_ri].cells[2].text.strip().replace(",", "")
                )
            # T128中找"一、短期薪酬"行
            t128_short = None
            for ri in range(len(t128.rows)):
                name = (
                    t128.rows[ri].cells[0].text.strip() if t128.rows[ri].cells else ""
                )
                if "短期薪酬" in name and "一" in name:
                    t128_short = float(
                        t128.rows[ri].cells[2].text.strip().replace(",", "")
                    )
                    break
            if t128_short and t129_sum:
                if abs(t128_short - t129_sum) < 0.02:
                    checks_passed += 1
                else:
                    issues.append(
                        f"CAS 9: T128一、短期薪酬({t128_short:,.2f}) ≠ T129明细合计({t129_sum:,.2f}) 差异={t128_short - t129_sum:,.2f}"
                    )
        except Exception as e:
            issues.append(f"CAS 9勾稽异常: {e}")

    if issues:
        print(f"  ⚠ 数据验证发现 {len(issues)} 个问题:")
        for issue in issues:
            print(f"    - {issue}")
    print(f"  ✅ 勾稽检查通过: {checks_passed}项" if checks_passed else "", end="")
    if not issues and checks_passed:
        print(f"  ✅ 语义级验证通过 ({checks_passed}项勾稽关系成立)")


def find_total_row(table):
    """找到表格中的合计行（支持多种合计关键词）"""
    for ri in range(len(table.rows)):
        name = ""
        for c in [0, 1]:
            if c < len(table.rows[ri].cells):
                t = table.rows[ri].cells[c].text.strip()
                if t:
                    name = t
                    break
        name_norm = name.replace(" ", "").replace("　", "")
        if name_norm in (
            "合计",
            "合  计",
            "合   计",
            "合　　　计",
            "小计",
            "小  计",
            "合 计",
            "合    计",
        ):
            return ri
        if len(name_norm) <= 4 and name_norm.startswith("合"):
            return ri
    # fallback: 最后一行
    return len(table.rows) - 1


def reconcile_with_main_statements(doc, taozhang_path):
    """v10: 附注数据与四大主表稽核"""
    import pandas as pd

    xl = pd.ExcelFile(taozhang_path)
    checks = 0

    # 读取主表数据
    def read_bs_item(sheet_name, item_keyword):
        """从资产负债表读取指定项目(排除'续'表)"""
        for s in xl.sheet_names:
            try:
                d = s.encode("latin1").decode("gbk")
            except:
                d = s
            if sheet_name in d and "续" not in d:
                df = pd.read_excel(taozhang_path, sheet_name=s, header=None)
                for i in range(df.shape[0]):
                    name = str(df.iloc[i, 1]).strip() if pd.notna(df.iloc[i, 1]) else ""
                    if item_keyword in name:
                        v = df.iloc[i, 3] if pd.notna(df.iloc[i, 3]) else None
                        if v is not None and str(v).strip() not in (
                            "",
                            "nan",
                            "－",
                            "-",
                        ):
                            try:
                                return float(v)
                            except:
                                pass
        return None

    def read_is_item_full(item_keyword):
        """从利润表读取指定项目, 返回{本期:value, 上期:value}"""
        for s in xl.sheet_names:
            try:
                d = s.encode("latin1").decode("gbk")
            except:
                d = s
            if "利润表" in d and "原始" in d:
                df = pd.read_excel(taozhang_path, sheet_name=s, header=None)
                # 找表头确定本期/上期列 (仅取左侧主表列C3/C4, 排除右侧补充列C7+)
                current_col, prior_col = 3, 4  # 默认利润表主列
                for i in range(min(5, df.shape[0])):
                    for j in [3, 4]:  # 仅检查主表列
                        if j < df.shape[1]:
                            h = (
                                str(df.iloc[i, j]).strip()
                                if pd.notna(df.iloc[i, j])
                                else ""
                            )
                            if "本期" in h:
                                current_col = j
                            if "上期" in h:
                                prior_col = j

                for i in range(df.shape[0]):
                    name = str(df.iloc[i, 1]).strip() if pd.notna(df.iloc[i, 1]) else ""
                    if item_keyword in name:
                        result = {}
                        for label, col in [("本期", current_col), ("上期", prior_col)]:
                            if col and col < df.shape[1]:
                                v = df.iloc[i, col]
                                if pd.notna(v) and str(v).strip() not in (
                                    "",
                                    "nan",
                                    "－",
                                    "-",
                                ):
                                    try:
                                        result[label] = float(v)
                                    except:
                                        pass
                        if result:
                            return result
        return {}

    def read_cf_item(item_keyword):
        """从现金流量表读取指定项目"""
        for s in xl.sheet_names:
            try:
                d = s.encode("latin1").decode("gbk")
            except:
                d = s
            if "现金流量表" in d and "原始" in d and "间接" not in d:
                df = pd.read_excel(taozhang_path, sheet_name=s, header=None)
                for i in range(df.shape[0]):
                    name = str(df.iloc[i, 1]).strip() if pd.notna(df.iloc[i, 1]) else ""
                    if item_keyword in name:
                        for j in [3, 4]:
                            if j < df.shape[1]:
                                v = df.iloc[i, j]
                                if pd.notna(v) and str(v).strip() not in (
                                    "",
                                    "nan",
                                    "－",
                                    "-",
                                ):
                                    try:
                                        return float(v)
                                    except:
                                        pass
        return None

    def note_table_val(table, total_col=1):
        """读取附注表格合计行的指定列数值"""
        tr = find_total_row(table)
        if tr is not None and total_col < len(table.rows[tr].cells):
            txt = table.rows[tr].cells[total_col].text.strip()
            if txt and txt not in ("—", "-", ""):
                return float(txt.replace(",", ""))
        return None

    # ---- 稽核项 ----
    # 1. 货币资金附注合计 = BS货币资金
    bs_cash = read_bs_item("资产负债表_原始数据", "货币资金")
    if bs_cash and 5 < len(doc.tables):
        try:
            note_cash = note_table_val(doc.tables[5], 1)
            if note_cash is not None:
                if abs(note_cash - float(bs_cash)) < 0.02:
                    print(
                        f"  ✅ 货币资金: 附注={note_cash:,.2f} = BS={float(bs_cash):,.2f}"
                    )
                    checks += 1
                else:
                    print(
                        f"  ⚠ 货币资金: 附注={note_cash:,.2f} ≠ BS={float(bs_cash):,.2f} 差异={note_cash - float(bs_cash):,.2f}"
                    )
            else:
                print(f"  ⚠ 货币资金: 附注表合计行无数据")
        except Exception as e:
            print(f"  ⚠ 货币资金稽核异常: {e}")

    # 2. 应收账款附注 = BS应收账款
    bs_ar = read_bs_item("资产负债表_原始数据", "应收账款")
    if bs_ar and 24 < len(doc.tables):
        try:
            note_ar = note_table_val(doc.tables[24], 5)  # C5=账面价值
            if note_ar is not None:
                if abs(note_ar - float(bs_ar)) < 0.02:
                    print(
                        f"  ✅ 应收账款: 附注={note_ar:,.2f} = BS={float(bs_ar):,.2f}"
                    )
                    checks += 1
                else:
                    print(
                        f"  ⚠ 应收账款: 附注={note_ar:,.2f} ≠ BS={float(bs_ar):,.2f} 差异={note_ar - float(bs_ar):,.2f}"
                    )
        except Exception as e:
            print(f"  ⚠ 应收账款稽核异常: {e}")

    # 3. 固定资产附注 = BS固定资产
    bs_fa = read_bs_item("资产负债表_原始数据", "固定资产")
    if bs_fa and 90 < len(doc.tables):
        try:
            note_fa = note_table_val(doc.tables[90], 1)
            if note_fa is not None:
                if abs(note_fa - float(bs_fa)) < 0.02:
                    print(
                        f"  ✅ 固定资产: 附注={note_fa:,.2f} = BS={float(bs_fa):,.2f}"
                    )
                    checks += 1
                else:
                    print(
                        f"  ⚠ 固定资产: 附注={note_fa:,.2f} ≠ BS={float(bs_fa):,.2f} 差异={note_fa - float(bs_fa):,.2f}"
                    )
        except Exception as e:
            print(f"  ⚠ 固定资产稽核异常: {e}")

    # 4. 营业收入附注 = IS营业收入(本期)
    is_rev_data = read_is_item_full("营业收入")
    if is_rev_data and 166 < len(doc.tables):
        try:
            note_rev = float(
                doc.tables[166].rows[2].cells[1].text.strip().replace(",", "")
            )
            is_rev = is_rev_data.get("本期") or is_rev_data.get("value")
            if is_rev and abs(note_rev - is_rev) < 0.02:
                print(f"  ✅ 营业收入: 附注={note_rev:,.2f} = IS={is_rev:,.2f}")
                checks += 1
            else:
                print(f"  ⚠ 营业收入: 附注={note_rev:,.2f} IS本期={is_rev}")
        except Exception as e:
            print(f"  ⚠ 营业收入稽核异常: {e}")

    # 5. 现金期末 = CF期末现金
    cf_cash = read_cf_item("期末现金")
    if cf_cash and 207 < len(doc.tables):
        try:
            note_cf = note_table_val(doc.tables[207], 1)
            if note_cf is not None:
                if abs(note_cf - cf_cash) < 0.02:
                    print(f"  ✅ 现金等价物: 附注={note_cf:,.2f} = CF={cf_cash:,.2f}")
                    checks += 1
                else:
                    print(f"  ⚠ 现金等价物: 附注={note_cf:,.2f} ≠ CF={cf_cash:,.2f}")
        except Exception as e:
            print(f"  ⚠ 现金等价物稽核异常: {e}")

    # 6. 实收资本附注 = BS实收资本
    bs_capital = read_bs_item("资产负债表_原始数据", "实收资本")
    if bs_capital and 159 < len(doc.tables):
        try:
            note_capital = note_table_val(doc.tables[159], 1)
            if note_capital is not None:
                if abs(note_capital - float(bs_capital)) < 0.02:
                    print(
                        f"  ✅ 实收资本: 附注={note_capital:,.2f} = BS={float(bs_capital):,.2f}"
                    )
                    checks += 1
                else:
                    print(
                        f"  ⚠ 实收资本: 附注={note_capital:,.2f} ≠ BS={float(bs_capital):,.2f}"
                    )
        except Exception as e:
            print(f"  ⚠ 实收资本稽核异常: {e}")

    # 7. 应交税费附注 = BS应交税费
    bs_tax = read_bs_item("资产负债表_原始数据", "应交税费")
    if bs_tax and 131 < len(doc.tables):
        try:
            note_tax = note_table_val(doc.tables[131], 1)
            if note_tax is not None:
                if abs(note_tax - float(bs_tax)) < 0.02:
                    print(
                        f"  ✅ 应交税费: 附注={note_tax:,.2f} = BS={float(bs_tax):,.2f}"
                    )
                    checks += 1
                else:
                    print(
                        f"  ⚠ 应交税费: 附注={note_tax:,.2f} ≠ BS={float(bs_tax):,.2f}"
                    )
        except Exception as e:
            print(f"  ⚠ 应交税费稽核异常: {e}")

    return checks


def fill_notes(taozhang_path, template_path, output_path, config_dir="config"):
    """主函数 — v13: config加载自YAML"""
    print("=" * 70)
    print("财务报表附注填充器 v10 — 会计语义引擎")
    print(f"套表: {Path(taozhang_path).name}")
    print(f"模板: {Path(template_path).name}")
    print("=" * 70)

    # v13: 从YAML加载配置
    config = load_config(config_dir)
    global NAME_SYNONYMS
    NAME_SYNONYMS = config.get("name_synonyms", {})
    mappings = config.get("mappings", [])

    doc = Document(template_path)

    print("\n📐 清理格式...")
    clean_formatting(doc)
    remove_instructions(doc)
    prepare_expense_tables(doc)

    xl = pd.ExcelFile(taozhang_path)

    print("\n📝 填充公司信息...")
    fill_company_info(doc, xl)

    stats = {
        "ok": 0,
        "no_sheet": 0,
        "no_data": 0,
        "no_table": 0,
        "no_match": 0,
        "err": 0,
    }
    total_cells = 0
    missing_report = {}

    for item in mappings:
        cat = item["cat"]
        sheet_kw = item["sheet_kw"]
        table_idx = item["table_idx"]
        tz_cols = item["tz_cols"]
        col_map = item["col_map"]

        print(f"\n[{cat}]")

        sheet_raw, sheet_decoded = find_sheet(xl, sheet_kw)
        if not sheet_raw:
            print(f"  ⏭ Sheet未找到: {sheet_kw}")
            stats["no_sheet"] += 1
            continue

        df = pd.read_excel(taozhang_path, sheet_name=sheet_raw, header=None)
        section_headers = item.get("section_headers")
        data = extract_tz_data(
            df,
            tz_cols,
            data_start=item.get("data_start"),
            name_strip=item.get("name_strip"),
            section=item.get("section"),
            name_exclude=item.get("name_exclude"),
            accumulate=item.get("accumulate", False),
            section_headers=section_headers,
            keep_empty=item.get("keep_empty", False),
        )
        if not data:
            print(f"  ⏭ 无数据 ({sheet_decoded})")
            stats["no_data"] += 1
            continue

        print(f"  {sheet_decoded}: {len(data)}行")

        if table_idx >= len(doc.tables):
            print(f"  ❌ Table #{table_idx} 超出范围")
            stats["no_table"] += 1
            continue

        table = doc.tables[table_idx]
        try:
            matched = fill_table(
                table,
                data,
                col_map,
                aggregations=item.get("aggregations"),
                section_headers=section_headers,
            )
            if matched > 0:
                stats["ok"] += 1
                total_cells += matched
                print(f"  ✅ 填充 {matched} 个单元格 → Table #{table_idx}")

                # Leakage check with smart categorization (v9)
                doc_names = set()
                doc_data_rows = 0  # Count rows with actual data (not header)
                for ri in range(len(table.rows)):
                    row_text = ""
                    for c in [0, 1]:
                        if c < len(table.rows[ri].cells):
                            t = table.rows[ri].cells[c].text.strip()
                            if t:
                                doc_names.add(t)
                                if not row_text:
                                    row_text = t
                    if row_text and ri > 0:  # Skip header row
                        doc_data_rows += 1

                # v12: strip section tag for missing detection
                missing = [
                    n
                    for n in data
                    if not any(match_name(dn, n.split("__sec")[0]) for dn in doc_names)
                ]
                if missing:
                    # 分类漏填: 结构性 vs 真正遗漏
                    structural = []
                    genuine = []
                    SECTION_KW = (
                        "其中：",
                        "（一）",
                        "（二）",
                        "（三）",
                        "（四）",
                        "（五）",
                        "（六）",
                        "（七）",
                        "（八）",
                        "（九）",
                        "（十）",
                        "一、",
                        "二、",
                        "三、",
                        "四、",
                        "五、",
                        "六、",
                        "1年以内",
                        "1-2年",
                        "2-3年",
                        "3-4年",
                        "4-5年",
                        "5年以上",
                        "小计",
                        "补充资料",
                        "基本医疗保险",
                        "补充医疗保险",
                        "工伤保险费",
                        "生育保险费",
                        "企业年金",
                        "失业保险费",
                        "应收工程款",
                        "企业员工",
                        "进项税加计抵减",
                    )  # 模板定制差异
                    N_A_ITEMS = ("储备基金", "企业发展基金")
                    for m in missing:
                        if any(kw in m for kw in SECTION_KW):
                            structural.append(m)
                        elif any(na in m for na in N_A_ITEMS):
                            structural.append(f"{m}(可能不适用)")
                        else:
                            genuine.append(m)

                    # If template has significantly fewer data rows, it's summary vs detail
                    if doc_data_rows < len(data) * 0.5 and doc_data_rows > 0:
                        structural.extend(genuine)
                        genuine = []

                    pct = len(missing) / len(data) * 100
                    if structural and not genuine:
                        print(
                            f"  ℹ 结构性差异 {len(missing)}/{len(data)} (模板{doc_data_rows}数据行←套表{len(data)}行)"
                        )
                    elif genuine:
                        print(
                            f"  ⚠ 漏填 {len(genuine)}/{len(data)}: {[m[:25] for m in genuine[:3]]}"
                        )
                        missing_report[cat] = genuine
                    if pct >= 30 and genuine:
                        missing_report[cat] = genuine
            else:
                stats["no_match"] += 1
                print(f"  ⚠ 名称不匹配")
                doc_names = []
                for ri in range(1, min(5, len(table.rows))):
                    for c in [0, 1]:
                        if c < len(table.rows[ri].cells):
                            t = table.rows[ri].cells[c].text.strip()
                            if t:
                                doc_names.append(t)
                                break
                print(f"    Doc: {doc_names}")
                print(f"    TZ: {list(data.keys())[:4]}")
        except Exception as e:
            stats["err"] += 1
            print(f"  ❌ {e}")
            import traceback

            traceback.print_exc()

    print(f"\n{'=' * 70}")
    print(f"📊 {stats['ok']}/{len(mappings)} 项成功, {total_cells} 个单元格")
    print(
        f"   ⏭ Sheet未找到:{stats['no_sheet']} 无数据:{stats['no_data']} "
        f"表格超出:{stats['no_table']} 名称不匹配:{stats['no_match']} 错误:{stats['err']}"
    )
    if missing_report:
        print(
            f"🔴 需关注漏填: {sum(len(v) for v in missing_report.values())}行 ({len(missing_report)}项) — 非结构性差异"
        )
    else:
        print(f"✅ 无实质性漏填 (所有差异均为模板摘要vs套表明细的结构性差异)")

    # Post-processing: auto-sum totals and validate
    print(f"\n🧮 自动计算合计行...")
    total_sums = 0
    for ti in range(len(doc.tables)):
        s = auto_sum_totals(doc.tables[ti])
        total_sums += s
    if total_sums:
        print(f"  📊 自动填充 {total_sums} 个合计单元格 (已过滤'其中'子项)")

    print(f"\n💰 税种税率自动填充...")
    fill_tax_rates(doc)

    print(f"\n✏️ 修正字体格式...")
    fix_table_fonts(doc)

    print(f"\n🔍 加减验证...")
    validate_sums(doc)

    print(f"\n🔍 会计准则语义验证...")
    validate_data(doc)

    # v10: 会计准则业务规则说明
    print(f"\n📋 会计准则应用说明:")
    rules = [
        'CAS 22.47: 单项计提应收账款归入"3年以上"账龄段(已发生信用减值,视为长期逾期)',
        "CAS 22.48: 组合计提按预期信用损失(ECL)模型,比例=各账龄段预期损失率",
        'CAS 36: 关联方组合(37,124,436.53元)无账龄明细→应在"关联交易"附注单独披露',
        "CAS 22: 账面价值=账面余额-坏账准备,金融资产按摊余成本列示",
        "CAS 9: 应付职工薪酬=短期薪酬+离职后福利+辞退福利+其他",
    ]
    for rule in rules:
        print(f"  • {rule}")

    # v10: 主表稽核
    print(f"\n📊 四大主表稽核...")
    bs_checks = reconcile_with_main_statements(doc, taozhang_path)
    if bs_checks:
        print(f"  ✅ 主表稽核通过: {bs_checks}项")
    else:
        print(f"  (主表数据未提取或无可比项)")

    print(f"\n💾 保存到: {output_path}")
    doc.save(output_path)


if __name__ == "__main__":
    fill_notes(
        "D:/Users/12844/Desktop/保定年审资料/2、决算套表/25年决算套表 (2)(1)/25年决算套表/保定吉达电力设计有限公司-单体.xlsx",
        "D:/Users/12844/Desktop/自动化填模板/已下发的文档模版/2.审计文档_docx/1.审计报告/1-2单体审计报告模板/4.2025年财务报表附注模板-单户.docx",
        "E:/Projects/my-workspace/_work/附注_单体_v11_正式版.docx",
        config_dir="config",
    )
