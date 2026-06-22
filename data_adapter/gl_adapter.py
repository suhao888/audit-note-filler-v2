"""
科目余额表适配器 — GL Adapter
从科目余额表/总分类账提取明细科目数据，按标准汇总规则（企业会计准则）输出 UnifiedData。

支持常见格式变化：
  - 全格式：编码 | 名称 | 期初借方 | 期初贷方 | 本期借方 | 本期贷方 | 期末借方 | 期末贷方
  - 简格式：编码 | 名称 | 期初余额 | 本期借方 | 本期贷方 | 期末余额
  - 多级表头合并、自动列类型检测、无编码列时以名称匹配
"""

import re
from typing import Dict, List, Optional, Tuple
from pathlib import Path

import pandas as pd

from .base import BaseAdapter, UnifiedData

# ============================================================
# 标准科目汇总规则（企业会计准则）
# ============================================================
# key: 合并后报表项目名, value: 参与汇总的科目编码前缀列表
STANDARD_ACCOUNT_GROUPS: Dict[str, List[str]] = {
    "货币资金": ["1001", "1002", "1012"],
    "应收账款": ["1122", "1231"],
    "存货": ["1401", "1402", "1403", "1405", "1471"],
    "固定资产": ["1601", "1602", "1603"],
    "无形资产": ["1701", "1702", "1703"],
    "短期借款": ["2001"],
    "应付职工薪酬": ["2211"],
    "应交税费": ["2221"],
    "长期借款": ["2501"],
    "营业收入": ["6001"],
    "营业成本": ["6401"],
    "管理费用": ["6602"],
    "销售费用": ["6601"],
    "财务费用": ["6603"],
    "研发费用": ["6604", "6701"],
}

# 分类：哪些汇总条目进入 BS / IS
_BS_GROUPS = {
    "货币资金",
    "应收账款",
    "存货",
    "固定资产",
    "无形资产",
    "短期借款",
    "应付职工薪酬",
    "应交税费",
    "长期借款",
}
_IS_GROUPS = {
    "营业收入",
    "营业成本",
    "管理费用",
    "销售费用",
    "财务费用",
    "研发费用",
}

# 收入类科目编码前缀（损益类中正常余额在贷方）
_REVENUE_PREFIXES = {
    "6001",
    "6011",
    "6021",
    "6031",
    "6041",
    "6051",
    "6061",
    "6101",
    "6111",
    "6301",
}

# 无编码列时，科目名称 → 标准科目名映射
_NAME_TO_STANDARD: Dict[str, str] = {
    "库存现金": "货币资金",
    "银行存款": "货币资金",
    "其他货币资金": "货币资金",
    "应收账款": "应收账款",
    "坏账准备": "应收账款",
    "原材料": "存货",
    "库存商品": "存货",
    "固定资产": "固定资产",
    "累计折旧": "固定资产",
    "固定资产减值准备": "固定资产",
    "无形资产": "无形资产",
    "累计摊销": "无形资产",
    "无形资产减值准备": "无形资产",
    "短期借款": "短期借款",
    "应付职工薪酬": "应付职工薪酬",
    "应交税费": "应交税费",
    "长期借款": "长期借款",
    "营业收入": "营业收入",
    "主营业务收入": "营业收入",
    "其他业务收入": "营业收入",
    "营业成本": "营业成本",
    "主营业务成本": "营业成本",
    "其他业务成本": "营业成本",
    "管理费用": "管理费用",
    "销售费用": "销售费用",
    "财务费用": "财务费用",
    "研发费用": "研发费用",
    "研发支出": "研发费用",
}

# 列类型关键词检测
_COL_KEYWORDS: Dict[str, List[str]] = {
    "code": ["编码", "编号", "代码", "科目号", "科目代码"],
    "name": ["项目名称", "科目名称", "名称"],
    "begin": ["期初", "年初", "上年"],
    "end": ["期末", "年末", "年终"],
    "current": ["本期", "本年"],
    "debit": ["借方", "借"],
    "credit": ["贷方", "贷"],
    "balance": ["余额"],
    "amount": ["发生额"],
}

# 科目余额表 Sheet 名称关键词
_SHEET_KW = ["科目余额", "科目汇总", "余额表", "总分类账", "总账", "明细账"]
# 需跳过的非数据行关键词
_SKIP_KW = [
    "合计",
    "小计",
    "总计",
    "累计",
    "单位：",
    "单位:",
    "负责人",
    "制表",
    "会计主管",
    "打印日期",
    "第",
    "页",
]
# 科目编码正则：至少1位数字开头
_CODE_PATTERN = re.compile(r"^\d+")


# ============================================================
# 工具函数
# ============================================================


def _first_digit(code: str) -> str:
    """取科目编码首位"""
    return code[0] if code else "0"


def _compute_net_balance(code: str, debit: float, credit: float) -> float:
    """
    计算科目净余额（有符号值）。

    资产类(1xxx)   : 借方 - 贷方，正常余额为正
    负债类(2xxx)   : 贷方 - 借方，正常余额为正
    权益类(4xxx)   : 贷方 - 借方
    成本类(5xxx)   : 借方 - 贷方
    损益类(6xxx)   : 收入科目(贷方-借方), 费用科目(借方-贷方)
    共同类(3xxx)   : 贷方 - 借方（默认）
    """
    head = _first_digit(code)
    if head == "1":
        return debit - credit
    elif head in "234":
        return credit - debit
    elif head == "5":
        return debit - credit
    elif head == "6":
        # 收入类科目正常余额在贷方
        if code[:4] in _REVENUE_PREFIXES:
            return credit - debit
        else:
            return debit - credit
    else:
        return debit - credit


def _normalize_code(code: str) -> str:
    """标准化科目编码，去除分隔符后缀等"""
    # 去掉 "." "-" "_" 后的子编码，如 "1001.01" → "1001"
    code = re.sub(r"[\.\-_].*$", "", code.strip())
    return code.strip()


def _merge_headers(df: pd.DataFrame, max_rows: int = 5) -> List[str]:
    """
    合并多级表头：将前 N 行逐层串联（从上到下），空单元格继承上面非空值。
    返回每列的合并标题字符串。
    """
    n_cols = df.shape[1]
    merged = [""] * n_cols

    for r in range(min(max_rows, len(df))):
        for c in range(n_cols):
            val = str(df.iloc[r, c]).strip() if pd.notna(df.iloc[r, c]) else ""
            if val and val.lower() != "nan":
                if merged[c]:
                    merged[c] += (" " + val) if not merged[c].endswith(" ") else val
                else:
                    merged[c] = val

    return merged


# ============================================================
# GLAdapter
# ============================================================


class GLAdapter(BaseAdapter):
    """科目余额表适配器"""

    @property
    def name(self) -> str:
        return "gl_adapter"

    def accept(self, filepath: str) -> bool:
        """检查是否包含科目余额表特征的 Sheet"""
        try:
            xl = pd.ExcelFile(filepath)
            sheets = xl.sheetnames
            for s in sheets:
                sl = s.lower()
                if any(kw.lower() in sl for kw in _SHEET_KW):
                    return True
            return False
        except Exception:
            return False

    # ---- 内部方法 ----

    def _find_gl_sheet(self, filepath: str) -> Tuple[Optional[pd.DataFrame], str]:
        """找到科目余额表所在 Sheet"""
        xl = pd.ExcelFile(filepath)
        sheets = xl.sheetnames

        # 按关键词匹配
        for s in sheets:
            sl = s.lower()
            if any(kw.lower() in sl for kw in _SHEET_KW):
                df = pd.read_excel(filepath, sheet_name=s, header=None)
                if df.shape[1] >= 4:  # 至少 4 列才是财务表
                    return df, s

        # 无匹配时选列数最多的 sheet（兜底）
        best_sheet, best_cols = "", 0
        for s in sheets:
            df = pd.read_excel(filepath, sheet_name=s, header=None)
            if df.shape[1] > best_cols:
                best_sheet, best_cols = s, df.shape[1]
        if best_sheet and best_cols >= 4:
            return pd.read_excel(
                filepath, sheet_name=best_sheet, header=None
            ), best_sheet

        return None, ""

    def _detect_columns(
        self, df: pd.DataFrame, max_header_rows: int = 5
    ) -> Tuple[Optional[Dict], int]:
        """
        自动检测列布局。

        返回 (col_map, header_end_row):
            col_map = {
                'code': int | None, 'name': int | None,
                'begin_debit': int | None, 'begin_credit': int | None,
                'curr_debit': int | None, 'curr_credit': int | None,
                'end_debit': int | None, 'end_credit': int | None,
                'begin_bal': int | None, 'end_bal': int | None,
            }
            header_end_row: 数据起始行索引
        """
        n_cols = df.shape[1]
        merged = _merge_headers(df, max_rows=max_header_rows)

        # ---- 对每列进行类型评分 ----
        scores: List[Dict[str, int]] = []
        for col_idx, header_text in enumerate(merged):
            by_type: Dict[str, int] = {}
            text_lower = header_text.lower()
            for ctype, patterns in _COL_KEYWORDS.items():
                cnt = 0
                for pat in patterns:
                    if pat in text_lower:
                        cnt += 1
                if cnt > 0:
                    by_type[ctype] = cnt
            scores.append(by_type)

        col_map: Dict[str, Optional[int]] = {
            "code": None,
            "name": None,
            "begin_debit": None,
            "begin_credit": None,
            "curr_debit": None,
            "curr_credit": None,
            "end_debit": None,
            "end_credit": None,
            "begin_bal": None,
            "end_bal": None,
        }

        # 1) 识别编码列
        code_idx = self._best_match(scores, "code", n_cols)
        if code_idx is not None:
            col_map["code"] = code_idx

        # 2) 识别名称列（排除已占用的编码列）
        name_idx = self._best_match(scores, "name", n_cols, exclude={code_idx})
        if name_idx is not None:
            col_map["name"] = name_idx
        elif code_idx is not None:
            # 兜底：名称紧邻编码右侧
            if code_idx + 1 < n_cols:
                col_map["name"] = code_idx + 1

        # 编码和名称都无识别结果 → 无法处理
        if col_map["code"] is None and col_map["name"] is None:
            return None, 0

        # 3) 识别金额列：按 (period, side) 组合
        used_cols = {col_map["code"], col_map["name"]}
        for col_idx in range(n_cols):
            if col_idx in used_cols:
                continue
            s = scores[col_idx]

            # 判断 period
            period = None
            for p in ["begin", "end", "current"]:
                if s.get(p, 0) > 0:
                    period = p
                    break
            if period is None:
                # 无明确时期标记 → 跳过（可能不是金额列）
                continue

            # 判断 side
            if (
                s.get("balance", 0) > 0
                and s.get("debit", 0) == 0
                and s.get("credit", 0) == 0
            ):
                # 单列余额（只有"余额"，无借/贷方向）
                key = f"{period}_bal"
                if col_map.get(key) is None:
                    col_map[key] = col_idx
                    used_cols.add(col_idx)
            elif s.get("debit", 0) > 0 and s.get("credit", 0) == 0:
                key = f"{period}_debit"
                if col_map.get(key) is None:
                    col_map[key] = col_idx
                    used_cols.add(col_idx)
            elif s.get("credit", 0) > 0 and s.get("debit", 0) == 0:
                key = f"{period}_credit"
                if col_map.get(key) is None:
                    col_map[key] = col_idx
                    used_cols.add(col_idx)
            elif (
                s.get("amount", 0) > 0
                and s.get("debit", 0) == 0
                and s.get("credit", 0) == 0
            ):
                # 只有"发生额"标记，无借/贷 → 跳过
                continue
            else:
                # 兼具借/贷标记，或无法判断 → 跳过，后面用备用规则处理
                pass

        # ---- 备用规则：如果 end_debit / end_credit 未识别到，尝试根据列位置推导 ----
        # 常见布局: 编码 | 名称 | 期初借 | 期初贷 | 本期借 | 本期贷 | 期末借 | 期末贷
        # 此时金额列有 6 列，按位置依次排列
        if col_map["end_debit"] is None or col_map["end_credit"] is None:
            self._fallback_column_detect(col_map, merged, scores, n_cols)

        # 确定数据起始行
        header_end = self._find_first_data_row(df, col_map, max_header_rows + 2)

        return col_map, header_end

    def _best_match(
        self,
        scores: List[Dict[str, int]],
        col_type: str,
        n_cols: int,
        exclude: Optional[set] = None,
    ) -> Optional[int]:
        """根据评分选出最佳匹配列"""
        if exclude is None:
            exclude = set()
        best_score = -1
        best_idx = None
        for i in range(n_cols):
            if i in exclude:
                continue
            sc = scores[i].get(col_type, 0)
            if sc > best_score:
                best_score = sc
                best_idx = i
        return best_idx

    def _fallback_column_detect(
        self, col_map: Dict, merged: List[str], scores: List[Dict], n_cols: int
    ):
        """
        备用列检测规则：
        当基于关键词分类不全时，根据列位置推断金额列含义。
        """
        # 收集已识别 + 名称/编码列之后的金额列
        known_cols = {v for v in col_map.values() if v is not None}
        # 按从左到右顺序扫描
        amount_cols = [i for i in range(n_cols) if i not in known_cols]

        # 尝试按 6 列布局匹配: 期初借|期初贷|本期借|本期贷|期末借|期末贷
        if len(amount_cols) >= 6:
            order = [
                "begin_debit",
                "begin_credit",
                "curr_debit",
                "curr_credit",
                "end_debit",
                "end_credit",
            ]
            for idx, col_idx in enumerate(amount_cols[:6]):
                key = order[idx]
                if col_map.get(key) is None:
                    col_map[key] = col_idx
        elif len(amount_cols) >= 4:
            # 4 列布局: 期初余额|本期借方|本期贷方|期末余额
            # 或: 期初借|期初贷|期末借|期末贷
            # 检查是否有 "发生额" 开头
            first_text = merged[amount_cols[0]].lower()
            if (
                "期初" in first_text
                and ("借" in first_text or "贷" in first_text)
                and "发生" not in first_text
            ):
                # 期初借方|期初贷方|期末借方|期末贷方
                if len(amount_cols) >= 4:
                    if col_map.get("begin_debit") is None:
                        col_map["begin_debit"] = amount_cols[0]
                    if col_map.get("begin_credit") is None:
                        col_map["begin_credit"] = amount_cols[1]
                    if col_map.get("end_debit") is None:
                        col_map["end_debit"] = amount_cols[2]
                    if col_map.get("end_credit") is None:
                        col_map["end_credit"] = amount_cols[3]
            else:
                # 期初余额|本期借方|本期贷方|期末余额
                if col_map.get("begin_bal") is None:
                    col_map["begin_bal"] = amount_cols[0]
                if col_map.get("curr_debit") is None:
                    col_map["curr_debit"] = amount_cols[1]
                if col_map.get("curr_credit") is None:
                    col_map["curr_credit"] = amount_cols[2]
                if col_map.get("end_bal") is None:
                    col_map["end_bal"] = amount_cols[3]
        elif len(amount_cols) >= 2:
            # 2 列布局: 期初余额 | 期末余额
            if col_map.get("begin_bal") is None:
                col_map["begin_bal"] = amount_cols[0]
            if col_map.get("end_bal") is None:
                col_map["end_bal"] = amount_cols[1]

    def _find_first_data_row(
        self, df: pd.DataFrame, col_map: Dict, start: int = 3
    ) -> int:
        """从 start 开始扫描，找到第一条数据行"""
        name_col = col_map.get("name")
        code_col = col_map.get("code")
        for r in range(start, len(df)):
            row = df.iloc[r]
            if self._is_data_row(row, name_col, code_col):
                return r
        return max(start, 0)

    def _is_data_row(
        self, row, name_col: Optional[int], code_col: Optional[int]
    ) -> bool:
        """
        判断是否为有效数据行。
        规则：有科目编码（以数字开头）或有科目名称，且非合计/小计行。
        """
        name_val = ""
        code_val = ""
        if name_col is not None and name_col < len(row):
            name_val = (
                str(row.iloc[name_col]).strip() if pd.notna(row.iloc[name_col]) else ""
            )
        if code_col is not None and code_col < len(row):
            code_val = (
                str(row.iloc[code_col]).strip() if pd.notna(row.iloc[code_col]) else ""
            )

        if not name_val and not code_val:
            return False

        # 跳过汇总/说明行
        combined = (name_val + " " + code_val).lower()
        if any(kw.lower() in combined for kw in _SKIP_KW):
            return False

        # 有数字编码
        if code_val and _CODE_PATTERN.match(code_val):
            return True
        # 无编码但有有意义的中文名称
        if name_val and any("\u4e00" <= ch <= "\u9fff" for ch in name_val[:2]):
            return True

        return False

    def _get_col_val(self, row, col_idx: Optional[int]) -> Optional[float]:
        """安全获取列值"""
        if col_idx is None or col_idx >= len(row):
            return None
        return self.safe_float(row.iloc[col_idx])

    def _collect_raw_accounts(
        self, df: pd.DataFrame, col_map: Dict, header_end: int
    ) -> List[dict]:
        """
        遍历数据行，收集原始科目数据。
        返回: [{code, name, begin_signed, end_signed, begin_debit, ...}, ...]
        """
        accounts_raw = []
        name_col = col_map.get("name")

        for r in range(header_end, len(df)):
            row = df.iloc[r]
            if not self._is_data_row(row, name_col, col_map.get("code")):
                continue

            name_val = ""
            if name_col is not None and name_col < len(row):
                name_val = (
                    str(row.iloc[name_col]).strip()
                    if pd.notna(row.iloc[name_col])
                    else ""
                )
            code_val = ""
            code_col = col_map.get("code")
            if code_col is not None and code_col < len(row):
                code_val = (
                    str(row.iloc[code_col]).strip()
                    if pd.notna(row.iloc[code_col])
                    else ""
                )

            if not name_val and not code_val:
                continue

            # 标准化编码
            code_val = _normalize_code(code_val) if code_val else code_val

            # 读取金额列
            begin_debit = self._get_col_val(row, col_map.get("begin_debit"))
            begin_credit = self._get_col_val(row, col_map.get("begin_credit"))
            curr_debit = self._get_col_val(row, col_map.get("curr_debit"))
            curr_credit = self._get_col_val(row, col_map.get("curr_credit"))
            end_debit = self._get_col_val(row, col_map.get("end_debit"))
            end_credit = self._get_col_val(row, col_map.get("end_credit"))
            begin_bal = self._get_col_val(row, col_map.get("begin_bal"))
            end_bal = self._get_col_val(row, col_map.get("end_bal"))

            # 计算净额
            if begin_bal is not None and end_bal is not None:
                # 单列余额模式：值已为净额
                begin_signed = begin_bal
                end_signed = end_bal
            else:
                bd = begin_debit if begin_debit is not None else 0.0
                bc = begin_credit if begin_credit is not None else 0.0
                ed = end_debit if end_debit is not None else 0.0
                ec = end_credit if end_credit is not None else 0.0
                code_for_calc = code_val if code_val else "1"
                begin_signed = _compute_net_balance(code_for_calc, bd, bc)
                end_signed = _compute_net_balance(code_for_calc, ed, ec)

            account = {
                "code": code_val,
                "name": name_val,
                "begin_signed": begin_signed,
                "end_signed": end_signed,
            }
            if curr_debit is not None:
                account["curr_debit"] = curr_debit
            if curr_credit is not None:
                account["curr_credit"] = curr_credit

            accounts_raw.append(account)

        return accounts_raw

    def _aggregate_standard_groups(self, accounts_raw: List[dict], data: UnifiedData):
        """
        按标准科目汇总规则聚合，输出到 unified_data.accounts / balance_sheet / income_statement。

        step 1: 按编码前缀匹配，将明细科目归入标准组
        step 2: 组内求和净额
        step 3: 输出正数化后的报表项目值
        """
        # ---- Step 1: 编码前缀匹配 ----
        used_raw: set = set()  # 记录已归组的原始科目索引，防止重复
        group_results: List[
            Tuple[str, float, float, List[str]]
        ] = []  # (组名, 期初净额, 期末净额, 明细列表)

        for group_name, code_prefixes in STANDARD_ACCOUNT_GROUPS.items():
            matched = []
            matched_names = []
            for idx, acct in enumerate(accounts_raw):
                if idx in used_raw:
                    continue
                code = acct["code"]
                if not code:
                    continue
                for prefix in code_prefixes:
                    if code.startswith(prefix):
                        matched.append(acct)
                        matched_names.append(acct["name"] or code)
                        used_raw.add(idx)
                        break

            if not matched:
                # 无编码匹配时，按名称匹配（兜底）
                for idx, acct in enumerate(accounts_raw):
                    if idx in used_raw:
                        continue
                    name = acct["name"]
                    if (
                        name
                        and name in _NAME_TO_STANDARD
                        and _NAME_TO_STANDARD[name] == group_name
                    ):
                        matched.append(acct)
                        matched_names.append(name)
                        used_raw.add(idx)

            if not matched:
                data.warnings.append(f"标准科目「{group_name}」未匹配到明细科目")
                continue

            total_begin = sum(m["begin_signed"] for m in matched)
            total_end = sum(m["end_signed"] for m in matched)

            group_results.append((group_name, total_begin, total_end, matched_names))

        # ---- Step 2: 将未归组科目直接写入 accounts ----
        for idx, acct in enumerate(accounts_raw):
            if idx not in used_raw:
                name = acct["name"] or acct["code"] or f"未命名_{idx}"
                data.accounts[name] = {
                    "期初": acct["begin_signed"],
                    "期末": acct["end_signed"],
                }

        # ---- Step 3: 输出到 accounts（标准组也放入 accounts 便于按名取数） ----
        for gname, t_begin, t_end, names in group_results:
            data.accounts[gname] = {"期初": t_begin, "期末": t_end}

        # ---- Step 4: 正数化输出到 BS / IS ----
        for gname, t_begin, t_end, names in group_results:
            if gname in _BS_GROUPS:
                # 负债类正常余额为正（因已按 credit-debit 计算），但取 abs 防御
                data.balance_sheet[gname] = {
                    "期初": abs(t_begin) if t_begin < 0 else t_begin,
                    "期末": abs(t_end) if t_end < 0 else t_end,
                }
                if t_begin < 0 or t_end < 0:
                    data.warnings.append(
                        f"标准科目「{gname}」净额为负({t_begin:.2f}/{t_end:.2f})，已取绝对值"
                    )
            elif gname in _IS_GROUPS:
                # IS 使用"本期"、"上期"命名
                data.income_statement[gname] = {
                    "上期": abs(t_begin) if t_begin < 0 else t_begin,
                    "本期": abs(t_end) if t_end < 0 else t_end,
                }
                if t_begin < 0 or t_end < 0:
                    data.warnings.append(
                        f"标准科目「{gname}」净额为负({t_begin:.2f}/{t_end:.2f})，已取绝对值"
                    )

    # ---- 主入口 ----

    def extract(self, filepath: str) -> UnifiedData:
        """提取科目余额表数据"""
        data = UnifiedData()
        data.source_files.append(str(Path(filepath).resolve()))

        # 1. 找到科目余额表 Sheet
        df, sheet_name = self._find_gl_sheet(filepath)
        if df is None:
            data.warnings.append("未找到科目余额表 Sheet")
            return data
        data.entity["数据来源Sheet"] = sheet_name

        # 2. 检测列布局
        col_map, header_end = self._detect_columns(df)
        if col_map is None:
            data.warnings.append("无法检测科目余额表列结构（未找到编码/名称列）")
            return data

        # 3. 提取明细科目
        accounts_raw = self._collect_raw_accounts(df, col_map, header_end)

        if not accounts_raw:
            data.warnings.append("未提取到任何科目数据")
            return data

        # 4. 按标准规则汇总
        self._aggregate_standard_groups(accounts_raw, data)

        # 5. 实体信息（文件名）
        data.entity["来源文件"] = Path(filepath).stem

        return data
