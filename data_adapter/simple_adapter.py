"""
简易适配器 — Simple Adapter
仅处理标准三表（资产负债表、利润表、现金流量表），不做明细科目提取。
适用场景：只有 BS/IS/CF 主表，无科目余额表或辅助明细表的 Excel 文件。
"""

import re
from typing import Dict, List, Optional, Tuple
from pathlib import Path

import pandas as pd

from .base import BaseAdapter, UnifiedData

# ============================================================
# Sheet 名称匹配模式
# ============================================================
_SHEET_PATTERNS: Dict[str, List[str]] = {
    "balance_sheet": [
        r"^.*资产负债.*$",
        r"^.*资产.*负债.*$",
        r"^.*资产负.*$",
        r"^.*balanc.*sheet.*$",
    ],
    "income_statement": [
        r"^.*利润表.*$",
        r"^.*利润.*$",
        r"^.*income.*statement.*$",
        r"^.*损益表.*$",
    ],
    "cash_flow": [
        r"^.*现金流量.*$",
        r"^.*现金.*流量.*$",
        r"^.*cash.*flow.*$",
    ],
}

# 列类型关键词（用于识别表头列含义）
_COL_PROJECT_KW = ["项目", "名称", "科目", "项目名称", "报表项目"]
_COL_CURRENT_KW = ["本期", "期末", "年末", "本年", "期末数", "期末余额", "期末金额"]
_COL_PREVIOUS_KW = ["上期", "期初", "年初", "上年", "期初数", "期初余额", "期初金额"]

# 行内容跳过关键词
_SKIP_ROW_KW = [
    "合计",
    "总计",
    "单位",
    "负责人",
    "制表",
    "会计主管",
    "补充资料",
    "附注",
    "法定代表人",
    "主管会计",
]

# 常见报表项目前缀（用于辅助定位数据起始行）
_BS_ITEM_PREFIX = [
    "一、",
    "二、",
    "三、",
    "四、",
    "五、",
    "六、",
    "流动资产",
    "非流动资产",
    "流动负债",
    "非流动负债",
    "所有者权益",
    "营业收入",
    "营业利润",
    "利润总额",
]
_IS_ITEM_PREFIX = [
    "一、",
    "二、",
    "三、",
    "四、",
    "五、",
    "六、",
    "营业收入",
    "减：",
    "加：",
    "营业利润",
    "利润总额",
]
_CF_ITEM_PREFIX = [
    "一、",
    "二、",
    "三、",
    "四、",
    "五、",
    "经营活动",
    "投资活动",
    "筹资活动",
    "现金流出",
    "现金流入",
]


# ============================================================
# SimpleAdapter
# ============================================================


class SimpleAdapter(BaseAdapter):
    """简易适配器 — 仅提取标准三表"""

    @property
    def name(self) -> str:
        return "simple_adapter"

    def accept(self, filepath: str) -> bool:
        """
        检查是否至少包含"利润表"或"资产负债表"其中之一。
        但不接受科目余额表（应由 GLAdapter 处理）。
        """
        try:
            xl = pd.ExcelFile(filepath)
            sheets = xl.sheetnames
        except Exception:
            return False

        # 排除科目余额表
        gl_kw = ["科目余额", "科目汇总", "余额表", "总账", "总分类账"]
        for s in sheets:
            sl = s.lower()
            if any(kw.lower() in sl for kw in gl_kw):
                return False

        # 检查三表
        for s in sheets:
            sl = s.lower()
            for patterns in _SHEET_PATTERNS.values():
                if any(re.search(p, sl) for p in patterns):
                    return True
        return False

    # ---- 内部方法 ----

    def _find_sheet(self, xl: pd.ExcelFile, patterns: List[str]) -> Optional[str]:
        """按模式匹配 Sheet 名称"""
        for s in xl.sheetnames:
            if any(re.search(p, s, re.IGNORECASE) for p in patterns):
                return s
        return None

    def _detect_header(self, df: pd.DataFrame) -> Tuple[int, int, int, int]:
        """
        检测表头与列位置。

        返回 (header_row, project_col, current_col, previous_col)：
          header_row: 表头所在行索引
          project_col: 项目名称列索引
          current_col: 本期/期末数列索引
          previous_col: 上期/期初数列索引
        """
        n_rows = min(10, len(df))
        n_cols = df.shape[1]

        best_header = 0
        best_project = 0
        best_current = 1
        best_previous = 2

        for r in range(n_rows):
            row = df.iloc[r]
            texts = [
                str(row[c]).strip() if pd.notna(row[c]) else "" for c in range(n_cols)
            ]

            # 统计关键词命中情况
            project_hits = sum(
                1
                for c in range(n_cols)
                if any(kw in texts[c] for kw in _COL_PROJECT_KW)
            )
            current_hits = sum(
                1
                for c in range(n_cols)
                if any(kw in texts[c] for kw in _COL_CURRENT_KW)
            )
            previous_hits = sum(
                1
                for c in range(n_cols)
                if any(kw in texts[c] for kw in _COL_PREVIOUS_KW)
            )

            # 该行有表头特征
            if project_hits > 0 or current_hits > 0 or previous_hits > 0:
                # 识别各列
                proj_col = self._find_col(texts, _COL_PROJECT_KW)
                curr_col = self._find_col(texts, _COL_CURRENT_KW)
                prev_col = self._find_col(texts, _COL_PREVIOUS_KW)

                if proj_col is not None:
                    best_project = proj_col
                if curr_col is not None:
                    best_current = curr_col
                if prev_col is not None:
                    best_previous = prev_col

                best_header = r
                break  # 取第一个表头行

        return best_header, best_project, best_current, best_previous

    def _find_col(self, texts: List[str], keywords: List[str]) -> Optional[int]:
        """在文本列表中查找第一个包含任意关键词的列"""
        for i, text in enumerate(texts):
            for kw in keywords:
                if kw in text:
                    return i
        return None

    def _find_data_start(
        self,
        df: pd.DataFrame,
        header_row: int,
        project_col: int,
        item_prefixes: List[str],
    ) -> int:
        """从表头之后找到第一条数据行"""
        for r in range(header_row + 1, len(df)):
            row = df.iloc[r]
            name = str(row[project_col]).strip() if pd.notna(row[project_col]) else ""
            if not name:
                continue
            # 报表项目通常以"一、"等前缀开头，或包含中文
            has_chinese = any("\u4e00" <= ch <= "\u9fff" for ch in name)
            if has_chinese:
                return r
            # 或匹配常见前缀
            if any(name.startswith(p) for p in item_prefixes):
                return r
        return header_row + 1

    def _extract_single_statement(
        self,
        df: pd.DataFrame,
        header_row: int,
        project_col: int,
        current_col: int,
        previous_col: int,
        item_prefixes: List[str],
    ) -> Dict[str, Dict[str, float]]:
        """从单张报表 Sheet 提取数据行"""
        result: Dict[str, Dict[str, float]] = {}
        n_cols = df.shape[1]
        data_start = self._find_data_start(df, header_row, project_col, item_prefixes)
        prev_name = ""  # 用于合并跨行项目名

        for r in range(data_start, len(df)):
            row = df.iloc[r]
            project_text = (
                str(row[project_col]).strip() if pd.notna(row[project_col]) else ""
            )

            # 空行 → 清空上次合并的上下文
            if not project_text:
                prev_name = ""
                continue

            # 跳过汇总/无效行
            if any(kw.lower() in project_text.lower() for kw in _SKIP_ROW_KW):
                continue

            # 跳过纯数字行（页码等）
            if re.match(r"^\d+$", project_text):
                continue

            # 跳过空项目名
            if not project_text or project_text.lower() == "nan":
                continue

            # 当前值
            curr_val = None
            prev_val = None
            if current_col < n_cols:
                curr_val = self.safe_float(row[current_col])
            if previous_col < n_cols:
                prev_val = self.safe_float(row[previous_col])

            # 仅当至少一个值有效时才记录
            if curr_val is not None or prev_val is not None:
                entry: Dict[str, float] = {}
                if curr_val is not None:
                    entry["期末"] = curr_val
                if prev_val is not None:
                    entry["期初"] = prev_val
                result[project_text] = entry

        return result

    # ---- 主入口 ----

    def extract(self, filepath: str) -> UnifiedData:
        """提取三表数据"""
        data = UnifiedData()
        data.source_files.append(str(Path(filepath).resolve()))

        try:
            xl = pd.ExcelFile(filepath)
        except Exception as e:
            data.warnings.append(f"无法打开文件: {e}")
            return data

        sheet_type_map = {
            "balance_sheet": (_SHEET_PATTERNS["balance_sheet"], _BS_ITEM_PREFIX),
            "income_statement": (_SHEET_PATTERNS["income_statement"], _IS_ITEM_PREFIX),
            "cash_flow": (_SHEET_PATTERNS["cash_flow"], _CF_ITEM_PREFIX),
        }

        for stype, (patterns, prefixes) in sheet_type_map.items():
            sheet_name = self._find_sheet(xl, patterns)
            if not sheet_name:
                data.warnings.append(f"未找到{stype}对应的 Sheet")
                continue

            df = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
            if df.shape[1] < 2:
                data.warnings.append(f"Sheet「{sheet_name}」列数不足，跳过")
                continue

            header_row, proj_col, curr_col, prev_col = self._detect_header(df)

            extracted = self._extract_single_statement(
                df, header_row, proj_col, curr_col, prev_col, prefixes
            )

            # 输出到 UnifiedData 对应字段
            if stype == "balance_sheet":
                data.balance_sheet = extracted
                data.entity["资产负债表Sheet"] = sheet_name
            elif stype == "income_statement":
                data.income_statement = extracted
                data.entity["利润表Sheet"] = sheet_name
            elif stype == "cash_flow":
                data.cash_flow = extracted
                data.entity["现金流量表Sheet"] = sheet_name

        # 实体信息
        data.entity["来源文件"] = Path(filepath).stem

        return data
