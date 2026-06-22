"""
Excel 决算套表剖析器 — 自动提取所有工作表的结构化信息
=====================================================
功能：
  - 遍历 Excel 所有 sheet，提取每个 sheet 的结构元数据
  - 自动检测表头行位置、列类型、科目名称、合计行
  - 按 资产负债表/利润表/现金流量表/附表明细/其他 分类
  - 支持 standalone 模式输出剖析报告

用法：
  from excel_profiler import ExcelProfiler
  profiler = ExcelProfiler()
  result = profiler.profile("path/to/taozhang.xlsx")
"""

import re
import sys
from pathlib import Path

import pandas as pd


# ============================================================
# 常量定义
# ============================================================

# 表头检测关键词 — 出现在表头行中的典型标记
HEADER_KEYWORDS = [
    "项目",
    "科目",
    "行次",
    "期末余额",
    "期初余额",
    "年初余额",
    "年末余额",
    "本期金额",
    "上期金额",
    "本年累计",
    "上年累计",
    "本期数",
    "上期数",
    "期末数",
    "期初数",
    "本年数",
    "上年数",
]

# 列类型检测关键词
COL_TYPE_RULES = [
    ("比例%", ["比例", "占比", "百分比"]),
    ("期末余额", ["期末余额", "期末数", "年末余额", "年末数", "期末"]),
    ("期初余额", ["期初余额", "期初数", "年初余额", "年初数", "期初"]),
    ("本年增加", ["本年增加", "本期增加", "本期借方"]),
    ("本年减少", ["本年减少", "本期减少", "本期贷方"]),
    ("本年累计", ["本年累计", "本年数", "本期金额"]),
    ("上年累计", ["上年累计", "上年数", "上期金额"]),
]

# 附注表名模式 — 用于 notes_detail 分类
NOTES_SHEET_PATTERNS = [
    "货币资金",
    "交易性金融资产",
    "应收票据",
    "应收账款",
    "预付账款",
    "其他应收款",
    "存货",
    "合同资产",
    "持有待售资产",
    "固定资产",
    "在建工程",
    "无形资产",
    "开发支出",
    "商誉",
    "长期待摊费用",
    "递延所得税资产",
    "其他非流动资产",
    "应付票据",
    "应付账款",
    "预收账款",
    "合同负债",
    "应付职工薪酬",
    "应交税费",
    "其他应付款",
    "长期应付款",
    "递延收益",
    "实收资本",
    "资本公积",
    "其他权益工具",
    "盈余公积",
    "营业收入",
    "营业成本",
    "税金及附加",
    "销售费用",
    "管理费用",
    "研发费用",
    "财务费用",
    "信用减值损失",
    "资产减值损失",
    "投资收益",
    "其他收益",
    "营业外收入",
    "营业外支出",
    "所得税费用",
    "专项应付款",
    "预计负债",
    "租赁负债",
    "使用权资产",
    "长期应付款",
    "递延所得税负债",
]

# 已知一级科目名 — 用于 detected_accounts 过滤
KNOWN_ACCOUNTS = set(
    [
        "库存现金",
        "银行存款",
        "其他货币资金",
        "交易性金融资产",
        "应收票据",
        "应收账款",
        "预付账款",
        "应收股利",
        "应收利息",
        "其他应收款",
        "坏账准备",
        "存货",
        "原材料",
        "库存商品",
        "发出商品",
        "周转材料",
        "合同资产",
        "合同履约成本",
        "持有待售资产",
        "债权投资",
        "长期股权投资",
        "投资性房地产",
        "固定资产",
        "累计折旧",
        "固定资产减值准备",
        "在建工程",
        "工程物资",
        "固定资产清理",
        "无形资产",
        "累计摊销",
        "研发支出",
        "商誉",
        "长期待摊费用",
        "递延所得税资产",
        "短期借款",
        "应付票据",
        "应付账款",
        "预收账款",
        "应付职工薪酬",
        "应交税费",
        "应付股利",
        "应付利息",
        "其他应付款",
        "合同负债",
        "持有待售负债",
        "长期借款",
        "应付债券",
        "长期应付款",
        "专项应付款",
        "预计负债",
        "递延收益",
        "递延所得税负债",
        "实收资本",
        "股本",
        "资本公积",
        "减：库存股",
        "其他综合收益",
        "盈余公积",
        "未分配利润",
        "主营业务收入",
        "其他业务收入",
        "营业外收入",
        "主营业务成本",
        "其他业务成本",
        "税金及附加",
        "销售费用",
        "管理费用",
        "研发费用",
        "财务费用",
        "其他收益",
        "投资收益",
        "信用减值损失",
        "资产减值损失",
        "营业利润",
        "营业外支出",
        "所得税费用",
        "净利润",
    ]
)


def clean_sheet_name(name: str) -> str:
    """清洗 sheet 名称：移除 '_原始数据' 等后缀"""
    name = str(name).strip()
    name = re.sub(r"_原始数据$", "", name)
    name = re.sub(r"_明细$", "", name)
    name = re.sub(r"_明细表$", "", name)
    return name


def is_numeric(val) -> bool:
    """判断一个值是否为数值类型（含 NaN 保护）"""
    if val is None:
        return False
    if isinstance(val, (int, float)):
        return not pd.isna(val)
    if isinstance(val, str):
        try:
            float(val.replace(",", "").replace(" ", ""))
            return True
        except (ValueError, TypeError):
            return False
    return False


def try_parse_num(val):
    """尝试将值转为 float，失败返回 None"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    try:
        return float(str(val).replace(",", "").replace(" ", ""))
    except (ValueError, TypeError):
        return None


# ============================================================
# ExcelProfiler 主类
# ============================================================


class ExcelProfiler:
    """
    决算套表剖析器 — 提取 Excel 工作簿中所有 sheet 的结构化信息。

    典型用法：
        profiler = ExcelProfiler()
        report = profiler.profile("决算套表.xlsx")
        print(report["num_sheets"], "个 sheet 已分析")
        for s in report["sheets"]:
            print(f"  {s['name']} → {s['category']} ({s['num_cols']}列×{s['num_rows']}行)")
    """

    def __init__(self, header_scan_rows: int = 15):
        """
        Args:
            header_scan_rows: 检测表头时扫描的最大行数
        """
        self.header_scan_rows = header_scan_rows

    # -------------------------------------------------------
    # 主入口
    # -------------------------------------------------------

    def profile(self, excel_path: str) -> dict:
        """
        剖析整个 Excel 文件，返回结构化报告。

        Args:
            excel_path: Excel 文件路径 (.xlsx)

        Returns:
            dict: 包含所有 sheet 元数据的报告，见模块文档
        """
        path = Path(excel_path)
        if not path.exists():
            raise FileNotFoundError(f"Excel 文件不存在: {excel_path}")
        if path.suffix not in (".xlsx", ".xls"):
            raise ValueError(f"不支持的文件格式: {path.suffix}，仅支持 .xlsx")

        # 读取 Excel
        try:
            xls = pd.ExcelFile(excel_path, engine="openpyxl")
        except Exception as e:
            raise RuntimeError(f"无法打开 Excel 文件: {e}")

        sheet_names = xls.sheet_names
        if not sheet_names:
            raise ValueError("Excel 文件中没有任何 sheet")

        sheets_report = []
        categories_map = {
            "balance_sheet": [],
            "income_statement": [],
            "cash_flow": [],
            "notes_detail": [],
            "other": [],
        }

        for sheet_name in sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            except Exception as e:
                # 读取失败时，以最小信息记录
                sheets_report.append(
                    {
                        "name": sheet_name,
                        "clean_name": clean_sheet_name(sheet_name),
                        "category": "other",
                        "header_row": -1,
                        "num_cols": 0,
                        "num_rows": 0,
                        "columns": [],
                        "rows": [],
                        "accounts": set(),
                        "has_total": False,
                        "error": str(e),
                    }
                )
                categories_map["other"].append(sheet_name)
                continue

            sheet_profile = self.profile_sheet(df, sheet_name)
            sheets_report.append(sheet_profile)

            cat = sheet_profile["category"]
            categories_map[cat].append(sheet_name)

        # 组装最终报告
        report = {
            "file": str(path.resolve()),
            "num_sheets": len(sheet_names),
            "sheets": sheets_report,
            "balance_sheets": categories_map["balance_sheet"],
            "income_statements": categories_map["income_statement"],
            "cash_flows": categories_map["cash_flow"],
            "notes_sheets": categories_map["notes_detail"],
            "other_sheets": categories_map["other"],
        }
        return report

    # -------------------------------------------------------
    # 单个 Sheet 剖析
    # -------------------------------------------------------

    def profile_sheet(self, df: pd.DataFrame, sheet_name: str) -> dict:
        """
        剖析单个 sheet，返回结构化描述。

        Args:
            df: pandas DataFrame（header=None 读取）
            sheet_name: 原始 sheet 名

        Returns:
            dict: sheet 元数据
        """
        clean_name = clean_sheet_name(sheet_name)
        num_cols = df.shape[1]
        total_rows = df.shape[0]

        # 如果 sheet 为空，返回空结构
        if total_rows == 0 or num_cols == 0:
            return {
                "name": sheet_name,
                "clean_name": clean_name,
                "category": self.classify_sheet(clean_name, df),
                "header_row": -1,
                "num_cols": num_cols,
                "num_rows": 0,
                "columns": [],
                "rows": [],
                "accounts": set(),
                "has_total": False,
            }

        # 1) 检测表头行
        header_row = self.detect_header_row(df)
        if header_row < 0:
            # 未找到表头，保守取第 0 行
            header_row = 0

        # 2) 提取列信息
        columns = self._extract_columns(df, header_row)

        # 3) 提取数据行
        data_start = header_row + 1
        rows_info = self._extract_rows(df, data_start)

        # 4) 提取科目名称
        accounts = self._extract_account_names(df, header_row)

        # 5) 检测是否有合计行
        has_total = self._detect_total(df, data_start)

        # 6) 分类
        category = self.classify_sheet(clean_name, df)

        data_row_count = max(0, total_rows - data_start)

        return {
            "name": sheet_name,
            "clean_name": clean_name,
            "category": category,
            "header_row": header_row,
            "num_cols": int(num_cols),
            "num_rows": int(data_row_count),
            "columns": columns,
            "rows": rows_info,
            "accounts": accounts,
            "has_total": has_total,
        }

    # -------------------------------------------------------
    # 表头行检测
    # -------------------------------------------------------

    def detect_header_row(self, df: pd.DataFrame) -> int:
        """
        检测表头行。

        策略：扫描前 N 行，对每行计算"表头匹配分"，
        匹配分 = 该行中包含 HEADER_KEYWORDS 的单元格数 + 数值单元格占比惩罚。

        Returns:
            检测到的表头行索引（0-based），未找到返回 -1
        """
        scan_limit = min(self.header_scan_rows, df.shape[0])
        best_row = -1
        best_score = -1

        for row_idx in range(scan_limit):
            row = df.iloc[row_idx]
            score = 0
            numeric_count = 0
            total_cells = 0

            for cell_val in row:
                if pd.isna(cell_val):
                    continue
                total_cells += 1
                s = str(cell_val).strip()
                if not s:
                    continue

                # 表头关键词命中加分
                for kw in HEADER_KEYWORDS:
                    if kw in s:
                        score += 2
                        break

                # 纯数值（非日期）惩罚 — 表头行不应有数值数据
                if is_numeric(cell_val) and not isinstance(cell_val, str):
                    numeric_count += 1

            # 表头行不应有大量数值单元格
            if total_cells > 0:
                numeric_ratio = numeric_count / total_cells
                if numeric_ratio > 0.5:
                    score -= 10

            # 空行惩罚
            if total_cells == 0:
                score -= 5

            # 记录最佳
            if score > best_score:
                best_score = score
                best_row = row_idx

        # 如果最佳行得分太低，返回 -1
        if best_score <= 0:
            return -1
        return best_row

    # -------------------------------------------------------
    # 列类型检测
    # -------------------------------------------------------

    def classify_column(self, header_text: str, sample_values: list) -> str:
        """
        检测列类型。

        规则优先级：
          1. 比例/占比关键词 → "比例%"
          2. 期末/年末关键词 → "期末余额"
          3. 期初/年初关键词 → "期初余额"
          4. 本年增加/本期增加 → "本年增加"
          5. 本年减少/本期减少 → "本年减少"
          6. 样本值以文本为主 → "文本"
          7. 默认 → "数值"

        Args:
            header_text: 列头文本
            sample_values: 该列前 N 个数据值（去空后）

        Returns:
            列类型字符串
        """
        h = str(header_text).strip() if header_text else ""

        # 对规则进行优先级匹配
        for col_type, keywords in COL_TYPE_RULES:
            for kw in keywords:
                if kw in h:
                    return col_type

        # 比例检测 — 样本值中有 % 或值为 0~1 小数
        if sample_values and not h:
            # 无表头的列，根据样本推断
            return self._infer_type_from_samples(sample_values)

        # 文本检测 — 检查表头本身
        text_indicators = ["项目", "名称", "注释", "备注", "说明", "类别", "类型"]
        for ti in text_indicators:
            if ti in h:
                return "文本"

        # 根据样本值辅助判断
        if sample_values:
            return self._infer_type_from_samples(sample_values)

        return "数值"

    def _infer_type_from_samples(self, sample_values: list) -> str:
        """从样本值推断列类型"""
        # 过滤 NaN/None
        valid = [
            v
            for v in sample_values
            if v is not None and not (isinstance(v, float) and pd.isna(v))
        ]
        if not valid:
            return "文本"

        # 检查是否有 % 值
        str_vals = [str(v).strip() for v in valid if isinstance(v, str)]
        if any("%" in s for s in str_vals):
            return "比例%"

        # 检查是否为 0~1 间的小数（比例特征）
        numeric_vals = [try_parse_num(v) for v in valid]
        numeric_vals = [n for n in numeric_vals if n is not None]
        if numeric_vals:
            all_small = all(0 <= n <= 1 for n in numeric_vals)
            any_large = any(n > 100 for n in numeric_vals)
            if all_small and len(numeric_vals) >= 2 and not any_large:
                return "比例%"
            return "数值"

        # 检查是否全为文本
        text_count = sum(1 for v in valid if isinstance(v, str) and not is_numeric(v))
        total = len(valid)
        if total > 0 and text_count / total > 0.8:
            return "文本"

        return "文本"

    # -------------------------------------------------------
    # Sheet 分类
    # -------------------------------------------------------

    def classify_sheet(self, clean_name: str, df: pd.DataFrame) -> str:
        """
        按 sheet 名称分类。

        Args:
            clean_name: 清洗后的 sheet 名称
            df: sheet 数据（用于辅助判断）

        Returns:
            分类值: balance_sheet / income_statement / cash_flow / notes_detail / other
        """
        name = clean_name

        # 资产负债表
        if re.search(r"资产|负债|资产负债", name):
            # 排除附注中包含"资产"二字的表
            # 如果同时匹配 notes_detail 模式，优先作为附注
            if not self._is_notes_name(name):
                return "balance_sheet"

        # 利润表 / 损益表
        if re.search(r"利润|损益", name):
            return "income_statement"

        # 现金流量表
        if re.search(r"现金", name):
            return "cash_flow"

        # 附注表
        if self._is_notes_name(name):
            return "notes_detail"

        # 根据数据内容辅助判断
        if df is not None and df.shape[0] > 2 and df.shape[1] >= 3:
            sample_text = " ".join(
                str(v) for v in df.iloc[:3, 0].values if not pd.isna(v)
            )
            # 检查是否有典型的资产负债表/利润表科目
            if any(k in sample_text for k in ["流动资产", "非流动资产", "流动负债"]):
                return "balance_sheet"
            if any(k in sample_text for k in ["营业收入", "营业利润", "净利润"]):
                return "income_statement"

        return "other"

    def _is_notes_name(self, name: str) -> bool:
        """判断 sheet 名是否为附注表名"""
        for pattern in NOTES_SHEET_PATTERNS:
            if pattern in name:
                return True
        return False

    # -------------------------------------------------------
    # 科目名提取
    # -------------------------------------------------------

    def extract_account_names(self, df: pd.DataFrame) -> list:
        """
        从 sheet 数据中提取科目名称列表。
        扫描第一列（通常是"项目"或"科目"列）提取已知会计科目。

        Args:
            df: sheet 数据（header=None 读取）

        Returns:
            list: 科目名称列表（去重）
        """
        header_row = self.detect_header_row(df)
        if header_row < 0:
            header_row = 0

        return self._extract_account_names(df, header_row)

    def _extract_account_names(self, df, header_row) -> set:
        """内部实现：从数据区提取科目名"""
        accounts = set()
        data_start = header_row + 1
        if data_start >= df.shape[0]:
            return accounts

        # 扫描第 0 列和第 1 列，寻找科目名
        for col_idx in (0, 1):
            if col_idx >= df.shape[1]:
                continue
            for row_idx in range(data_start, df.shape[0]):
                val = df.iloc[row_idx, col_idx]
                if pd.isna(val):
                    continue
                s = str(val).strip()
                if not s or len(s) < 2:
                    continue

                # 清理序号前缀
                cleaned = self._clean_account_name(s)
                if cleaned and len(cleaned) >= 2:
                    accounts.add(cleaned)

        return accounts

    def _clean_account_name(self, name: str) -> str:
        """清洗科目名：去除序号前缀、特殊字符"""
        s = name.strip()
        # 去除序号前缀，如 "一、", "（一）", "1.", "1、", "(1)" 等
        s = re.sub(r"^[一二三四五六七八九十]+、", "", s)
        s = re.sub(r"^[（(][一二三四五六七八九十]+[）)]", "", s)
        s = re.sub(r"^\d+[\.、．]", "", s)
        s = re.sub(r"^[（(]\d+[）)]", "", s)
        s = s.strip()
        # 跳过纯数字、空字符串
        if not s or re.match(r"^[\d\s,.%\-—－]+$", s):
            return ""
        return s

    # -------------------------------------------------------
    # 辅助方法
    # -------------------------------------------------------

    def _extract_columns(self, df: pd.DataFrame, header_row: int) -> list:
        """提取列信息列表"""
        columns = []
        num_cols = df.shape[1]

        for col_idx in range(num_cols):
            # 提取表头
            header_val = df.iloc[header_row, col_idx] if header_row >= 0 else None
            header_text = str(header_val).strip() if not pd.isna(header_val) else ""

            # 提取样本值（从数据区取前 3 个）
            sample_values = []
            data_start = header_row + 1
            for r in range(data_start, min(data_start + 10, df.shape[0])):
                v = df.iloc[r, col_idx]
                if not pd.isna(v):
                    sample_values.append(v)
                if len(sample_values) >= 3:
                    break

            col_type = self.classify_column(header_text, sample_values)

            columns.append(
                {
                    "index": int(col_idx),
                    "header": header_text if header_text else f"Col_{col_idx}",
                    "type": col_type,
                    "sample_values": sample_values[:3],
                }
            )

        return columns

    def _extract_rows(self, df: pd.DataFrame, data_start: int) -> list:
        """提取数据行信息（行名 + 值列表）"""
        rows = []
        if data_start >= df.shape[0]:
            return rows

        for row_idx in range(data_start, df.shape[0]):
            # 行名取自第 0 列
            row_name_val = df.iloc[row_idx, 0] if df.shape[1] > 0 else None
            if pd.isna(row_name_val):
                continue
            row_name = str(row_name_val).strip()
            if not row_name:
                continue

            # 收集该行所有值
            values = []
            for col_idx in range(1, df.shape[1]):
                v = df.iloc[row_idx, col_idx]
                values.append(None if pd.isna(v) else v)

            rows.append(
                {
                    "row_name": row_name,
                    "row_index": int(row_idx),
                    "values": values,
                }
            )

        return rows

    def _detect_total(self, df: pd.DataFrame, data_start: int) -> bool:
        """检测是否有合计行"""
        if data_start >= df.shape[0]:
            return False
        for row_idx in range(data_start, df.shape[0]):
            for col_idx in range(min(2, df.shape[1])):
                v = df.iloc[row_idx, col_idx]
                if pd.isna(v):
                    continue
                s = str(v).strip()
                if s in ("合计", "合  计", "合   计", "总计", "小计"):
                    return True
        return False


# ============================================================
# 便捷函数
# ============================================================


def quick_profile(excel_path: str) -> dict:
    """
    快速剖析 Excel 文件，一行调用。

    Args:
        excel_path: Excel 文件路径

    Returns:
        dict: 结构化报告
    """
    profiler = ExcelProfiler()
    return profiler.profile(excel_path)


def print_profile_summary(report: dict) -> None:
    """打印剖析报告摘要"""
    print("=" * 60)
    print(f"文件: {report['file']}")
    print(f"Sheet 总数: {report['num_sheets']}")
    print(f"=" * 60)
    print(f"  资产负债表: {len(report['balance_sheets'])}个")
    for s in report["balance_sheets"]:
        print(f"    - {s}")
    print(f"  利润表: {len(report['income_statements'])}个")
    for s in report["income_statements"]:
        print(f"    - {s}")
    print(f"  现金流量表: {len(report['cash_flows'])}个")
    for s in report["cash_flows"]:
        print(f"    - {s}")
    print(f"  附注明细表: {len(report['notes_sheets'])}个")
    for s in report["notes_sheets"]:
        print(f"    - {s}")
    print(f"  其他: {len(report['other_sheets'])}个")
    for s in report["other_sheets"]:
        print(f"    - {s}")
    print(f"=" * 60)
    print()

    for sheet in report["sheets"]:
        print(f"[{sheet['category']:16s}] {sheet['name']}")
        print(
            f"      表头行={sheet['header_row']}, "
            f"列数={sheet['num_cols']}, 数据行={sheet['num_rows']}, "
            f"合计={'Y' if sheet['has_total'] else 'N'}"
        )
        cols_info = ", ".join(
            f"{c['header'][:8].ljust(8)}({c['type']})" for c in sheet["columns"][:6]
        )
        if sheet["columns"]:
            print(f"      前列: {cols_info}")
        if sheet["accounts"]:
            accts = list(sheet["accounts"])[:8]
            print(f"      科目: {', '.join(accts)}")
        print()


# ============================================================
# Standalone 模式
# ============================================================


def main():
    """命令行入口——剖析 Excel 并输出结构摘要"""
    import argparse

    parser = argparse.ArgumentParser(
        description="Excel 决算套表剖析器 — 提取所有工作表的结构化信息"
    )
    parser.add_argument("excel_path", help="Excel 文件路径 (.xlsx)")
    parser.add_argument(
        "--json", action="store_true", help="以 JSON 格式输出（含完整列信息）"
    )
    parser.add_argument(
        "--sheet", "-s", type=str, default=None, help="只分析指定 sheet（名称或索引）"
    )
    args = parser.parse_args()

    try:
        profiler = ExcelProfiler()
        report = profiler.profile(args.excel_path)
    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)
        sys.exit(1)

    if args.json:
        import json

        class SetEncoder(json.JSONEncoder):
            def default(self, obj):
                if isinstance(obj, set):
                    return sorted(obj)
                return super().default(obj)

        # 过滤指定 sheet
        if args.sheet:
            filtered = [
                s
                for s in report["sheets"]
                if s["name"] == args.sheet or s["clean_name"] == args.sheet
            ]
            report["sheets"] = filtered
            report["num_sheets"] = len(filtered)

        print(json.dumps(report, ensure_ascii=False, indent=2, cls=SetEncoder))
    else:
        # 过滤指定 sheet
        if args.sheet:
            filtered = [
                s
                for s in report["sheets"]
                if s["name"] == args.sheet or s["clean_name"] == args.sheet
            ]
            report["sheets"] = filtered
            report["num_sheets"] = len(filtered)
        print_profile_summary(report)


if __name__ == "__main__":
    main()
