"""
国资委决算套表适配器

处理国资委标准决算套表格式（.xlsx），提取为 UnifiedData 统一数据模型。

特征检测:
  - 含有 `资产负债表_原始数据` / `利润表_原始数据` 等 sheet
  - 列布局: col B(1)=科目名, col D(3)=期末/本期, col E(4)=期初/上期
  - Sheet 名可能出现 GBK 编码问题（需 latin1→gbk 转换）
"""

import re
import pandas as pd
from pathlib import Path
from typing import Dict, List, Optional, Set

from .base import BaseAdapter, UnifiedData


# ── 全局排除清单 ──────────────────────────────────────────────
# 这些 sheet 不会被纳入 accounts 明细表
_MAIN_SHEET_KEYWORDS = [
    "资产负债表",
    "利润表",
    "现金流量表",
    "所有者权益变动",
    "所有者权益变动表",
]

_ENTITY_SHEET_KEYWORDS = [
    "境内企业基础信息",
    "企业基础信息",
    "Sheet79",
    "附注报表封面",
    "封面",
]

_IGNORE_SHEET_KEYWORDS = [
    "辅助信息",
    "参数设置",
    "公式",
    "存储",
    "说明",
    "模板",
    "合计行",
    "校核公式",
    "审核公式",
]

# 明细表中常见的无用行名（合计/小计在有些表不需要，但保留）
_NOISE_ROW_NAMES = {"", "nan", "none", "合计", "小计", "总计"}

# ── Sheet 名编/解码用 ──────────────────────────────────────────
_GBK_ENCODE_CACHE: Dict[str, str] = {}


def _decode_sheet_name(raw: str) -> str:
    """处理决算套表的 GBK 编码问题：latin1→gbk 回退"""
    if raw in _GBK_ENCODE_CACHE:
        return _GBK_ENCODE_CACHE[raw]
    try:
        decoded = raw.encode("latin1").decode("gbk")
    except Exception:
        decoded = raw
    _GBK_ENCODE_CACHE[raw] = decoded
    return decoded


def _norm_sheet_name(name: str) -> str:
    """标准化 sheet 名（去空格、小写）"""
    return _decode_sheet_name(name).replace(" ", "").lower()


def _is_sheet_match(sheet_name: str, keyword: str) -> bool:
    """模糊匹配 sheet 名是否包含 keyword"""
    return keyword.lower() in _norm_sheet_name(sheet_name)


def _extract_sheet_keywords(sheet_name: str) -> str:
    """去掉 _原始数据 / _明细 等后缀得到语义关键词"""
    name = _decode_sheet_name(sheet_name)
    name = re.sub(r"_原始数据$", "", name)
    name = re.sub(r"_原始_data$", "", name)
    name = re.sub(r"_明细表$", "", name)
    name = re.sub(r"_明细$", "", name)
    return name


def _safe_val(df: pd.DataFrame, row: int, col: int) -> Optional[float]:
    """安全取值并转 float；越界 / NaN / 公式残留 返回 None"""
    if row >= df.shape[0] or col >= df.shape[1]:
        return None
    v = df.iloc[row, col]
    if v is None:
        return None
    if isinstance(v, (int, float)):
        if pd.isna(v):
            return None
        return float(v)
    try:
        s = str(v).replace(",", "").replace(" ", "").strip()
        if s in ("", "nan", "－", "-", "—", "none", "null"):
            return None
        return float(s)
    except (ValueError, TypeError):
        return None


def _get_row_name(df: pd.DataFrame, row: int, col: int = 1) -> str:
    """取行名并清洗"""
    if row >= df.shape[0] or col >= df.shape[1]:
        return ""
    v = df.iloc[row, col]
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    # 移除不可见字符
    s = re.sub(r"[\u0000-\u001f]", "", s)
    return s


def _is_noise_row(name: str) -> bool:
    """是否为噪音行（空行、纯符号等）"""
    return name.lower() in _NOISE_ROW_NAMES or not name or len(name) <= 1


# ══════════════════════════════════════════════════════════════
# TaozhangAdapter
# ══════════════════════════════════════════════════════════════


class TaozhangAdapter(BaseAdapter):
    """处理国资委标准决算套表格式"""

    @property
    def name(self) -> str:
        return "taozhang"

    # ── accept ──────────────────────────────────────────────

    def accept(self, filepath: str) -> bool:
        """检查是否为决算套表格式

        判断依据: 含有 资产负债表_原始数据 / 利润表_原始数据 等 sheet
        """
        try:
            xl = pd.ExcelFile(filepath)
        except Exception:
            return False

        for sheet in xl.sheet_names:
            ns = _norm_sheet_name(sheet)
            for kw in [
                "资产负债表_原始数据",
                "资产负债表_原始",
                "利润表_原始数据",
                "利润表_原始",
            ]:
                if kw in ns:
                    return True
        return False

    # ── extract（主入口）────────────────────────────────────

    def extract(self, filepath: str) -> UnifiedData:
        """提取为统一数据"""
        data = UnifiedData()
        data.source_files.append(filepath)

        try:
            xl = pd.ExcelFile(filepath)
        except Exception as e:
            data.warnings.append(f"无法打开文件: {e}")
            return data

        # 1. 提取资产负债表
        self._extract_bs(xl, data)

        # 2. 提取利润表
        self._extract_is(xl, data)

        # 3. 提取现金流量表
        self._extract_cf(xl, data)

        # 4. 提取明细表 → accounts
        self._extract_detail_sheets(xl, data)

        # 5. 提取企业信息
        self._extract_entity(xl, data)

        return data

    # ── _find_sheet ─────────────────────────────────────────

    def _find_sheet(self, xl: pd.ExcelFile, candidates: List[str]) -> Optional[str]:
        """模糊查找匹配的 sheet 名（返回实际 sheet 名，非 decode 后）"""
        for sheet in xl.sheet_names:
            ns = _norm_sheet_name(sheet)
            for c in candidates:
                if c.lower() in ns:
                    return sheet
        return None

    def _find_sheets(self, xl: pd.ExcelFile, candidates: List[str]) -> List[str]:
        """模糊查找所有匹配的 sheet（返回实际 sheet 名列表）"""
        matched: List[str] = []
        for sheet in xl.sheet_names:
            ns = _norm_sheet_name(sheet)
            for c in candidates:
                if c.lower() in ns:
                    matched.append(sheet)
                    break
        return matched

    def _classify_sheet(self, sheet_name: str) -> str:
        """将 sheet 分类: bs / is / cf / entity / detail / ignore"""
        ns = _norm_sheet_name(sheet_name)

        for kw in _IGNORE_SHEET_KEYWORDS:
            if kw.lower() in ns:
                return "ignore"

        for kw in _ENTITY_SHEET_KEYWORDS:
            if kw.lower() in ns:
                return "entity"

        for kw in _MAIN_SHEET_KEYWORDS:
            if kw.lower() in ns:
                # 进一步区分
                if "利润表" in ns:
                    return "is"
                if "现金流量表" in ns:
                    return "cf"
                if "资产负债" in ns:
                    return "bs"
                return "other_main"

        # 含有 _原始数据 / _原始_data 后缀的视为明细表
        decoded = _decode_sheet_name(sheet_name)
        if re.search(r"_原始数据|_原始_data|_明细表|_明细$", decoded):
            return "detail"

        # 包含基础信息/信息的视为 entity
        if "基础信息" in ns or "基本信息" in ns or "企业信息" in ns:
            return "entity"

        return "ignore"

    # ── 数据提取子方法 ──────────────────────────────────────

    # ---- 资产负债表 ----

    def _extract_bs(self, xl: pd.ExcelFile, data: UnifiedData) -> None:
        """从资产负债表_原始数据 提取"""
        sheet = self._find_sheet(xl, ["资产负债表_原始数据", "资产负债表"])
        if not sheet:
            data.warnings.append("未找到资产负债表 sheet")
            return

        try:
            df = pd.read_excel(xl, sheet_name=sheet, header=None)
        except Exception as e:
            data.warnings.append(f"读取资产负债表失败: {e}")
            return

        header_row = 5  # 表头通常在第5行（0-based）
        count = 0
        for i in range(header_row, len(df)):
            name = _get_row_name(df, i, 1)  # col B
            if _is_noise_row(name):
                continue

            end_val = _safe_val(df, i, 3)  # col D = 期末
            beg_val = _safe_val(df, i, 4)  # col E = 期初

            val: Dict[str, float] = {}
            if end_val is not None:
                val["期末"] = end_val
            if beg_val is not None:
                val["期初"] = beg_val
            if val:
                data.balance_sheet[name] = val
                count += 1

        if count == 0:
            data.warnings.append("资产负债表提取 0 行数据（可能表头行号不匹配）")

    # ---- 利润表 ----

    def _extract_is(self, xl: pd.ExcelFile, data: UnifiedData) -> None:
        """从利润表_原始数据 提取

        利润表的列是 本期/上期（不是 期末/期初）。
        同时检测表头中列的实际位置（部分套表有补充列）。
        """
        sheet = self._find_sheet(xl, ["利润表_原始数据", "利润表"])
        if not sheet:
            data.warnings.append("未找到利润表 sheet")
            return

        try:
            df = pd.read_excel(xl, sheet_name=sheet, header=None)
        except Exception as e:
            data.warnings.append(f"读取利润表失败: {e}")
            return

        # 检测表头，确定本期/上期列位置
        current_col, prior_col = 3, 4  # 默认位置
        for i in range(min(5, df.shape[0])):
            for j in range(min(10, df.shape[1])):
                v = str(df.iloc[i, j]).strip() if pd.notna(df.iloc[i, j]) else ""
                if "本期" in v or "本年" in v:
                    current_col = j
                elif "上期" in v or "上年" in v:
                    prior_col = j

        count = 0
        for i in range(5, len(df)):
            name = _get_row_name(df, i, 1)
            if _is_noise_row(name):
                continue

            cur_val = _safe_val(df, i, current_col) if current_col is not None else None
            pri_val = _safe_val(df, i, prior_col) if prior_col is not None else None

            val: Dict[str, float] = {}
            if cur_val is not None:
                val["本期"] = cur_val
            if pri_val is not None:
                val["上期"] = pri_val
            if val:
                data.income_statement[name] = val
                count += 1

        if count == 0:
            data.warnings.append("利润表提取 0 行数据")

    # ---- 现金流量表 ----

    def _extract_cf(self, xl: pd.ExcelFile, data: UnifiedData) -> None:
        """从现金流量表_原始数据 提取（排除间接法表）"""
        candidates = ["现金流量表_原始数据", "现金流量表"]
        sheet = None
        for s in xl.sheet_names:
            ns = _norm_sheet_name(s)
            for c in candidates:
                if c.lower() in ns and "间接" not in ns:
                    sheet = s
                    break
            if sheet:
                break

        if not sheet:
            data.warnings.append("未找到现金流量表 sheet")
            return

        try:
            df = pd.read_excel(xl, sheet_name=sheet, header=None)
        except Exception as e:
            data.warnings.append(f"读取现金流量表失败: {e}")
            return

        # 检测列位置
        current_col, prior_col = 3, 4
        for i in range(min(5, df.shape[0])):
            for j in range(min(10, df.shape[1])):
                v = str(df.iloc[i, j]).strip() if pd.notna(df.iloc[i, j]) else ""
                if "本期" in v or "本年" in v or "期末" in v:
                    current_col = j
                elif "上期" in v or "上年" in v or "期初" in v:
                    prior_col = j

        count = 0
        for i in range(5, len(df)):
            name = _get_row_name(df, i, 1)
            if _is_noise_row(name):
                continue

            cur_val = _safe_val(df, i, current_col) if current_col is not None else None
            pri_val = _safe_val(df, i, prior_col) if prior_col is not None else None

            val: Dict[str, float] = {}
            # UnifiedData.cash_flow 约定用 期末/期初 键名
            if cur_val is not None:
                val["期末"] = cur_val
            if pri_val is not None:
                val["期初"] = pri_val
            if val:
                data.cash_flow[name] = val
                count += 1

        if count == 0:
            data.warnings.append("现金流量表提取 0 行数据")

    # ---- 明细表提取 ----

    def _extract_detail_sheets(self, xl: pd.ExcelFile, data: UnifiedData) -> None:
        """提取所有明细表（_原始数据 / _原始_data 等）到 data.accounts

        策略:
          - 对所有非主表/非entity sheet 尝试提取
          - 每张 sheet 生成一个命名空间前缀（基于 sheet 名语义）
          - 提取 col B(1) = 行名, col D(3) = 期末/本期, col E(4) = 期初/上期
          - 忽略空行/合计行/纯符号行
        """
        extracted: Set[str] = set()  # 记录已提取的 sheet，避免重复

        for sheet in xl.sheet_names:
            cat = self._classify_sheet(sheet)
            if cat not in ("detail", "other_main"):
                continue
            if sheet in extracted:
                continue
            extracted.add(sheet)

            try:
                df = pd.read_excel(xl, sheet_name=sheet, header=None)
            except Exception as e:
                data.warnings.append(
                    f"读取明细表 [{_decode_sheet_name(sheet)}] 失败: {e}"
                )
                continue

            if df.shape[0] < 3 or df.shape[1] < 4:
                # 行/列太少，跳过
                continue

            # 生成命名空间前缀: "货币资金_原始数据" → "货币资金"
            prefix = _extract_sheet_keywords(sheet)
            # 清理多余字符
            prefix = prefix.replace("（", "(").replace("）", ")").strip()

            # 检测列: 找第一个非空行确定列标签
            # 尝试找 header 行（第 3~6 行中 col D 含"期末"/"本期"的）
            detected_current_col = 3
            detected_prior_col = 4
            current_label = "期末"
            prior_label = "期初"

            for i in range(min(6, df.shape[0])):
                for j in range(3, min(10, df.shape[1])):
                    v = str(df.iloc[i, j]).strip() if pd.notna(df.iloc[i, j]) else ""
                    if v in (
                        "期末",
                        "期末余额",
                        "期末数",
                        "期末账面余额",
                        "本期",
                        "本期数",
                    ):
                        detected_current_col = j
                        if "期" in v:
                            current_label = "期末"
                        if "本" in v:
                            current_label = "本期"
                    elif v in (
                        "期初",
                        "期初余额",
                        "期初数",
                        "期初账面余额",
                        "上期",
                        "上期数",
                    ):
                        detected_prior_col = j
                        if "期" in v:
                            prior_label = "期初"
                        if "上" in v:
                            prior_label = "上期"

            # 数据起始行（表头 +1）
            data_start = 6  # 默认
            for i in range(min(6, df.shape[0])):
                name = _get_row_name(df, i, 1)
                if _is_noise_row(name):
                    continue
                # 找到第一个有数值的行
                v1 = _safe_val(df, i, detected_current_col)
                if v1 is not None:
                    data_start = i
                    break

            # 如果 data_start 不是从 6 找到的，说明表头行不标准，用 i 作为起始
            rows_added = 0
            for i in range(data_start, len(df)):
                name = _get_row_name(df, i, 1)
                if _is_noise_row(name):
                    continue

                # 排除明显不是数据行的行名
                decoded_name = _decode_sheet_name(name)
                if decoded_name.lower() in ("合计", "小计", "总计", "合计行", "其中"):
                    continue
                # 排除空括号行
                if re.match(r"^[\s\(\（\)\）]+$", decoded_name):
                    continue

                cur_val = _safe_val(df, i, detected_current_col)
                pri_val = _safe_val(df, i, detected_prior_col)

                val: Dict[str, float] = {}
                if cur_val is not None:
                    val[current_label] = cur_val
                if pri_val is not None:
                    val[prior_label] = pri_val
                if not val:
                    # 两个列都没值，跳过
                    continue

                # 构建带命名空间的 key: "货币资金:库存现金"
                key = f"{prefix}:{name}"
                data.accounts[key] = val
                rows_added += 1

            # 同时写入简写版（不含前缀）如果有冲突则保留前缀版
            # 已在上方写入前缀版，不必重复写入非前缀版避免覆盖

    # ---- 企业信息提取 ----

    def _extract_entity(self, xl: pd.ExcelFile, data: UnifiedData) -> None:
        """提取企业基本信息

        数据源顺序:
          1. 境内企业基础信息_原始数据（最正式）
          2. 附注报表封面 / Sheet79 / 封面
          3. 任何含"基础信息"的 sheet
        """
        entity_sources = [
            "境内企业基础信息_原始数据",
            "企业基础信息_原始数据",
            "附注报表封面",
            "附注汇总",
            "Sheet79",
        ]

        # 先尝试查找数据源 sheet
        best_sheet = None
        for kw in entity_sources:
            s = self._find_sheet(xl, [kw])
            if s:
                best_sheet = s
                break

        if not best_sheet:
            # 兜底：查找含"基础信息"的 sheet
            for s in xl.sheet_names:
                if "基础信息" in _norm_sheet_name(s) or "基本信息" in _norm_sheet_name(
                    s
                ):
                    best_sheet = s
                    break

        if not best_sheet:
            best_sheet = self._find_sheet(xl, ["Sheet79", "封面"])

        if not best_sheet:
            data.warnings.append(
                "未找到企业信息 sheet（境内企业基础信息_原始数据/Sheet79/封面）"
            )
            return

        try:
            df = pd.read_excel(xl, sheet_name=best_sheet, header=None)
        except Exception as e:
            data.warnings.append(f"读取企业信息 sheet 失败: {e}")
            return

        # 定义已知字段的映射（key关键词 → entity field name）
        field_map: Dict[str, str] = {
            "企业名称": "name",
            "统一社会信用代码": "uscc",
            "信用代码": "uscc",
            "法定代表人": "legal_rep",
            "法人代表": "legal_rep",
            "注册资本": "registered_capital",
            "注册地址": "address",
            "通讯地址": "address",
            "住所": "address",
            "成立日期": "founded_date",
            "成立时间": "founded_date",
            "企业类型": "company_type",
            "经济性质": "company_type",
            "所属行业": "industry",
            "经营范围": "business_scope",
            "控股股东": "shareholder",
            "最终控制方": "ultimate_controller",
        }

        # 模式A: "境内企业基础信息" 专用格式 — col B=项目名, col D=值
        # 通常 5 行之后开始
        found_field = False
        for i in range(3, min(40, df.shape[0])):
            key_raw = _get_row_name(df, i, 1)
            if not key_raw:
                continue

            # 去掉末尾冒号
            key = key_raw.rstrip("：:")
            val_raw = (
                str(df.iloc[i, 3]).strip()
                if df.shape[1] > 3 and pd.notna(df.iloc[i, 3])
                else ""
            )
            if not val_raw or val_raw in ("nan", "None"):
                continue

            # 匹配已知字段
            for kw_tag, field_name in field_map.items():
                if kw_tag in key:
                    data.entity[field_name] = val_raw
                    found_field = True
                    break

        # 如果境内企业基础信息未找到字段，尝试附注报表封面格式
        if not found_field:
            # 模式B: 封面/Sheet79 格式 — 遍历所有行 col B=key, col D=value
            for i in range(5, min(30, df.shape[0])):
                key_raw = _get_row_name(df, i, 1)
                val_raw = (
                    str(df.iloc[i, 3]).strip()
                    if df.shape[1] > 3 and pd.notna(df.iloc[i, 3])
                    else ""
                )

                if not key_raw or not val_raw or val_raw in ("nan", "None"):
                    continue

                key = key_raw.replace("：", ":").rstrip(":").strip()
                if len(key) <= 2 or len(val_raw) <= 1:
                    continue

                for kw_tag, field_name in field_map.items():
                    if kw_tag in key:
                        data.entity[field_name] = val_raw
                        break

        # 补充: 尝试从文件名提取企业名（当 entity.name 为空时）
        if "name" not in data.entity and data.source_files:
            fp = Path(data.source_files[0])
            stem = fp.stem
            # 剥离常见后缀
            for suffix in ["单体", "合并", "决算套表", "报表"]:
                stem = stem.replace(suffix, "")
            stem = stem.strip("-_（）() ")
            if stem and len(stem) >= 4:
                data.entity["name"] = stem

        if not data.entity:
            data.warnings.append("企业信息提取为空（sheet 格式不标准）")
