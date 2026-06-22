#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
配置生成器 — 将 SmartMatcher 的匹配结果转换为完整的 config/ 目录
================================================================

功能:
  1. 读取智能匹配结果（含置信度评分）
  2. 生成 mappings.yaml（列映射配置）
  3. 生成 name_synonyms.yaml（名称同义词映射）
  4. 生成 auto_sum_rules.yaml（自动求和规则）
  5. 生成 adaptation_report.md（适配报告供用户审核）
  6. learn_from_existing() — 从已有配置学习模式

输出格式与 config_loader.py 完全兼容。

用法:
  from config_generator import ConfigGenerator
  gen = ConfigGenerator(output_dir="config_generated")
  result = gen.generate(match_result, template_analysis, excel_profile)
"""

import sys
import re
import json
import shutil
from pathlib import Path
from datetime import datetime
from collections import defaultdict
from typing import Dict, List, Optional, Tuple, Any, Set

try:
    import yaml
except ImportError:
    try:
        from ruamel import yaml
    except ImportError:
        print("ERROR: 需要 PyYAML 库。请运行: pip install pyyaml")
        sys.exit(1)

# ✅ 引擎使用的标准会计科目（用于生成 name_synonyms 时的参考集）
KNOWN_ACCOUNTS: Set[str] = {
    # 资产类
    "货币资金",
    "库存现金",
    "银行存款",
    "其他货币资金",
    "交易性金融资产",
    "应收票据",
    "应收账款",
    "应收款项融资",
    "预付账款",
    "应收股利",
    "应收利息",
    "其他应收款",
    "存货",
    "合同资产",
    "持有待售资产",
    "一年内到期的非流动资产",
    "其他流动资产",
    "债权投资",
    "其他债权投资",
    "长期应收款",
    "长期股权投资",
    "其他权益工具投资",
    "其他非流动金融资产",
    "投资性房地产",
    "固定资产",
    "在建工程",
    "使用权资产",
    "无形资产",
    "开发支出",
    "商誉",
    "长期待摊费用",
    "递延所得税资产",
    "其他非流动资产",
    # 负债类
    "短期借款",
    "应付票据",
    "应付账款",
    "预收账款",
    "合同负债",
    "应付职工薪酬",
    "应交税费",
    "应付利息",
    "应付股利",
    "其他应付款",
    "持有待售负债",
    "一年内到期的非流动负债",
    "其他流动负债",
    "长期借款",
    "应付债券",
    "长期应付款",
    "租赁负债",
    "预计负债",
    "递延收益",
    "递延所得税负债",
    "其他非流动负债",
    # 权益类
    "实收资本",
    "股本",
    "资本公积",
    "减：库存股",
    "其他综合收益",
    "专项储备",
    "盈余公积",
    "未分配利润",
    # 损益类
    "营业收入",
    "营业成本",
    "税金及附加",
    "销售费用",
    "管理费用",
    "研发费用",
    "财务费用",
    "其他收益",
    "投资收益",
    "公允价值变动收益",
    "信用减值损失",
    "资产减值损失",
    "资产处置收益",
    "营业外收入",
    "营业外支出",
    "所得税费用",
    # 现金流量
    "现金流量表",
    "现金流量表补充资料",
}

# 中文序号前缀
CN_NUM_PREFIXES = [
    "一、",
    "二、",
    "三、",
    "四、",
    "五、",
    "六、",
    "七、",
    "八、",
    "九、",
    "十、",
    "（一）",
    "（二）",
    "（三）",
    "（四）",
    "（五）",
    "（六）",
    "（七）",
    "（八）",
    "(一)",
    "(二)",
    "(三)",
    "(四)",
    "(五)",
    "(六)",
    "(七)",
    "(八)",
    "1．",
    "2．",
    "3．",
    "4．",
    "5．",
    "6．",
    "7．",
    "8．",
    "9．",
    "10．",
    "1.",
    "2.",
    "3.",
    "4.",
    "5.",
    "6.",
    "7.",
    "8.",
    "9.",
    "10.",
    "1、",
    "2、",
    "3、",
    "4、",
    "5、",
    "6、",
    "7、",
    "8、",
    "9、",
    "10、",
    "（1）",
    "（2）",
    "（3）",
    "（4）",
    "（5）",
    "(1)",
    "(2)",
    "(3)",
    "(4)",
    "(5)",
]

# 表类型后缀
TABLE_SUFFIXES = [
    "-账龄",
    "-组合",
    "-明细",
    "-分类",
    "-补充资料",
    "-情况",
    "-增减",
    "-前五名",
    "-情况表",
    "-总表",
]


# ============================================================
# 辅助函数
# ============================================================


def _normalize_whitespace(text: str) -> str:
    """标准化空白字符：全角→半角，多个空格→单个"""
    if not text:
        return ""
    text = str(text)
    text = text.replace("\u3000", " ")  # 全角空格 → 半角
    text = text.replace("\xa0", " ")  # &nbsp; → 空格
    text = re.sub(r"\s+", "", text)  # 所有空白 → 空（用于名称比较）
    return text


def _strip_prefix(name: str) -> str:
    """去除名称前的序号前缀"""
    for prefix in CN_NUM_PREFIXES:
        if name.startswith(prefix):
            return name[len(prefix) :].strip()
    return name.strip()


def _extract_core_name(name: str) -> str:
    """提取核心名称：去前缀+标准化空白"""
    return _strip_prefix(_normalize_whitespace(name))


def _confidence_label(score: float) -> str:
    """置信度标签"""
    if score >= 0.8:
        return "高"
    elif score >= 0.5:
        return "中"
    else:
        return "低"


def _ensure_dir(path: Path):
    """确保目录存在"""
    path.mkdir(parents=True, exist_ok=True)


def _write_yaml(filepath: Path, data, header_comment: str = ""):
    """写 YAML 文件（utf-8，含可选文件头注释）"""
    content = yaml.dump(
        data,
        allow_unicode=True,
        default_flow_style=False,
        sort_keys=False,
        width=120,
    )
    if header_comment:
        content = header_comment + "\n" + content
    filepath.write_text(content, encoding="utf-8")


def _sorted_unique(seq) -> list:
    """去重并保持顺序"""
    seen = set()
    result = []
    for item in seq:
        if item not in seen:
            seen.add(item)
            result.append(item)
    return result


# ============================================================
# ConfigGenerator 主类
# ============================================================


class ConfigGenerator:
    """
    配置生成器 — 将 SmartMatcher 匹配结果结构化输出为 YAML 配置 + 适配报告。

    典型用法:
        gen = ConfigGenerator(output_dir="config_generated")
        result = gen.generate(match_result, template_analysis, excel_profile)
        if result["has_issues"]:
            print("部分映射需要人工审核，见适配报告。")
    """

    def __init__(self, output_dir: str = "config_generated"):
        """
        参数:
            output_dir: 生成配置文件的输出目录（相对或绝对路径）
        """
        self.output_dir = Path(output_dir)
        self._learned_patterns: Dict[str, Any] = {
            # 从已有配置学到的模式缓存
            "template_to_sheet": {},  # {模板表名: [常用套表Sheet关键词]}
            "col_map_patterns": [],  # [常见列映射模式]
        }

    # ── 主入口 ──────────────────────────────────────────────

    def generate(
        self,
        match_result: dict,
        template_analysis: dict,
        excel_profile: dict,
    ) -> dict:
        """
        生成所有配置文件。

        参数:
            match_result: SmartMatcher 的匹配结果，结构如下:
                {
                    "matches": [
                        {
                            "template_table_index": 5,
                            "template_table_name": "货币资金-明细",
                            "excel_sheet": "货币资金_原始数据",
                            "confidence": 0.95,
                            "tz_cols": [3, 4],
                            "col_map": [[0, 1, false], [1, 2, false]],
                            "row_matches": [...],
                        },
                        ...
                    ],
                    "unmatched_templates": [{"index": 99, "name": "..."}],
                    "unmatched_sheets": ["sheet_name_1"],
                    "extra_sheets": [...],
                }
            template_analysis: TemplateAnalyzer.analyze() 的返回值
            excel_profile: ExcelProfiler.profile() 的返回值

        返回:
            dict: 生成结果报告
                {
                    "config_dir": str,          # 输出目录路径
                    "files": {str: str},        # {文件名: 状态}
                    "stats": {str: int},        # 统计信息
                    "has_issues": bool,         # 是否有需人工审核项
                    "issues": [str],            # 问题列表
                    "report_content": str,      # 适配报告 Markdown 内容
                    "status": str,              # ok / needs_review / error
                }
        """
        # 1) 验证输入
        validation = self._validate_inputs(
            match_result, template_analysis, excel_profile
        )
        if not validation["valid"]:
            return {
                "config_dir": str(self.output_dir),
                "files": {},
                "stats": {"error": 1},
                "has_issues": True,
                "issues": validation["errors"],
                "report_content": "",
                "status": "error",
            }

        # 2) 清空/创建输出目录
        self._prepare_output_dir()

        # 3) 分别生成各文件
        files_generated = {}
        issues: List[str] = []
        stats: Dict[str, int] = defaultdict(int)

        # --- mappings.yaml ---
        mappings_yaml = self._generate_mappings_yaml(match_result)
        mappings_path = self.output_dir / "mappings.yaml"
        _write_yaml(mappings_path, mappings_yaml, self._MAPPINGS_HEADER)
        num_mappings = len(mappings_yaml) if isinstance(mappings_yaml, list) else 0
        files_generated["mappings.yaml"] = f"已写入 ({num_mappings} 条映射)"
        stats["mappings"] = num_mappings

        # --- name_synonyms.yaml ---
        synonyms_yaml = self._generate_name_synonyms_yaml(
            template_analysis, excel_profile, match_result
        )
        synonyms_path = self.output_dir / "name_synonyms.yaml"
        _write_yaml(synonyms_path, synonyms_yaml, self._SYNONYMS_HEADER)
        num_synonyms = len(synonyms_yaml) if isinstance(synonyms_yaml, dict) else 0
        files_generated["name_synonyms.yaml"] = f"已写入 ({num_synonyms} 条同义词)"
        stats["synonyms"] = num_synonyms

        # --- auto_sum_rules.yaml ---
        sum_rules_yaml = self._generate_auto_sum_rules(template_analysis, match_result)
        sum_rules_path = self.output_dir / "auto_sum_rules.yaml"
        _write_yaml(sum_rules_path, sum_rules_yaml, self._SUM_RULES_HEADER)
        files_generated["auto_sum_rules.yaml"] = "已写入"
        stats["auto_sum_rules"] = (
            len(sum_rules_yaml) if isinstance(sum_rules_yaml, dict) else 0
        )

        # 统计匹配置信度分布
        matches = match_result.get("matches", [])
        for m in matches:
            conf = m.get("confidence", 0)
            if conf >= 0.8:
                stats["high_confidence"] += 1
            elif conf >= 0.5:
                stats["medium_confidence"] += 1
            else:
                stats["low_confidence"] += 1

        # 未匹配统计
        unmatched_tables = match_result.get("unmatched_templates", [])
        unmatched_sheets = match_result.get("unmatched_sheets", [])
        stats["unmatched_tables"] = len(unmatched_tables)
        stats["unmatched_sheets"] = len(unmatched_sheets)

        # 收集问题
        for m in matches:
            conf = m.get("confidence", 0)
            tname = m.get(
                "template_table_name", f"表{m.get('template_table_index', '?')}"
            )
            if conf < 0.5:
                issues.append(f"{tname}: 置信度过低 ({conf:.2f}) — 请手动确认配对")
            elif conf < 0.8:
                detail = ""
                col_map = m.get("col_map", [])
                if not col_map:
                    detail = "，无列映射"
                issues.append(f"{tname}: 中等置信度 ({conf:.2f}){detail} — 建议审核")

        for ut in unmatched_tables:
            tname = ut.get("name", f"表{ut.get('index', '?')}")
            issues.append(f"{tname}: 未匹配到任何 Excel Sheet — 需要手动设置")

        for us in unmatched_sheets:
            issues.append(
                f"Excel Sheet「{us}」: 未匹配到任何模板表 — 可能为多余或需新增映射"
            )

        has_issues = (
            stats["medium_confidence"] > 0
            or stats["low_confidence"] > 0
            or stats["unmatched_tables"] > 0
        )

        # --- adaptation_report.md ---
        report_content = self._generate_adaptation_report(
            match_result,
            template_analysis,
            excel_profile,
            stats,
            issues,
        )
        report_path = self.output_dir / "adaptation_report.md"
        report_path.write_text(report_content, encoding="utf-8")
        files_generated["adaptation_report.md"] = "已写入"

        # 5) 返回结果
        return {
            "config_dir": str(self.output_dir.resolve()),
            "files": files_generated,
            "stats": dict(stats),
            "has_issues": has_issues,
            "issues": issues,
            "report_content": report_content,
            "status": "needs_review" if has_issues else "ok",
        }

    # ── 输入验证 ────────────────────────────────────────────

    def _validate_inputs(self, match_result, template_analysis, excel_profile) -> dict:
        """验证三个输入的结构完整性"""
        errors = []

        if not isinstance(match_result, dict):
            errors.append("match_result 必须为 dict 类型")
        else:
            required_keys = ["matches"]
            for k in required_keys:
                if k not in match_result:
                    errors.append(f"match_result 缺少关键字段: {k}")

        if not isinstance(template_analysis, dict):
            errors.append("template_analysis 必须为 dict 类型")
        else:
            if (
                "tables" not in template_analysis
                and "num_tables" not in template_analysis
            ):
                errors.append("template_analysis 缺少关键字段: tables 或 num_tables")

        if not isinstance(excel_profile, dict):
            errors.append("excel_profile 必须为 dict 类型")
        else:
            if "sheets" not in excel_profile and "num_sheets" not in excel_profile:
                errors.append("excel_profile 缺少关键字段: sheets 或 num_sheets")

        return {
            "valid": len(errors) == 0,
            "errors": errors,
        }

    # ── 输出目录准备 ────────────────────────────────────────

    def _prepare_output_dir(self):
        """清空并重建输出目录"""
        if self.output_dir.exists():
            shutil.rmtree(str(self.output_dir))
        _ensure_dir(self.output_dir)

    # ── 文件头注释 ──────────────────────────────────────────

    _MAPPINGS_HEADER = """# ============================================================
# 自动生成的列映射配置 — 由 ConfigGenerator 从匹配结果生成
# 生成时间: {timestamp}
# ============================================================
# 字段说明：
#   cat:         附注类别标识（对应模板表格用途）
#   sheet_kw:   决算套表中Sheet名称关键词（模糊匹配）
#   table_idx:  模板中表格索引号（python-docx tables序号）
#   tz_cols:    套表中数据列的索引列表（0-based）
#   col_map:    列映射列表 [(套表列索引, 模板列索引, 是否百分比)]
#   data_start: 套表数据起始行（默认为表头+1）
#   name_strip: 行名前缀剥离列表
#   name_exclude: 排除的套表行名关键词
#   section_headers: 分区标题列表
#   confidence:  匹配置信度（生成时参考）
# ============================================================
""".format(timestamp=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    _SYNONYMS_HEADER = """# ============================================================
# 自动生成的名称同义词映射 — 由 ConfigGenerator 从名称变异分析生成
# 生成时间: {timestamp}
# ============================================================
# 格式: "模板行名": ["套表行名1", "套表行名2", ...]
# 将套表中的数据行名称映射到模板中的对应行名称
# ============================================================
""".format(timestamp=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    _SUM_RULES_HEADER = """# ============================================================
# 自动生成的合计行求和规则 — 由 ConfigGenerator 从合计行检测生成
# 生成时间: {timestamp}
# ============================================================
# 格式:
#   <表索引>:
#     totals: [<合计行索引>, ...]
#     subtotals: [<小计行索引>, ...]
# ============================================================
""".format(timestamp=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    # ── mappings.yaml 生成 ──────────────────────────────────

    def _generate_mappings_yaml(self, match_result: dict) -> list:
        """
        从匹配结果生成 MAPPINGS 条目列表。

        每条映射包含:
        - cat:        表名（从模板分析结果取）
        - sheet_kw:   Sheet 名称关键词（用于后续模糊匹配）
        - table_idx:  模板中表格索引
        - tz_cols:    套表数据列的索引列表（排序后）
        - col_map:    (excel_col, template_col, is_pct) 三元组
        - data_start: 数据起始行（可选）
        - name_strip: 行名前缀剥离列表（可选）
        - name_exclude: 排除行关键词（可选）
        - section_headers: 分区标题（可选）
        """
        mappings: list = []
        matches = match_result.get("matches", [])

        # 按表索引排序，保持模板中的表格顺序
        matches_sorted = sorted(matches, key=lambda m: m.get("template_table_index", 0))

        for m in matches_sorted:
            entry = self._build_mapping_entry(m)
            if entry:
                mappings.append(entry)

        return mappings

    def _build_mapping_entry(self, match: dict) -> Optional[dict]:
        """从单条匹配结果构建一个 MAPPINGS 条目"""
        try:
            template_name = match.get("template_table_name", "")
            sheet_name = match.get("excel_sheet", "")
            table_idx = match.get("template_table_index")
            tz_cols_raw = match.get("tz_cols", [])
            col_map_raw = match.get("col_map", [])
            confidence = match.get("confidence", 0)

            if table_idx is None:
                return None

            # tz_cols: 排序去重
            tz_cols = _sorted_unique(sorted(tz_cols_raw))

            # col_map: 转为标准格式
            col_map = []
            for cm in col_map_raw:
                if isinstance(cm, (list, tuple)) and len(cm) >= 3:
                    col_map.append(
                        {
                            "excel_col": int(cm[0]),
                            "tmpl_col": int(cm[1]),
                            "is_pct": bool(cm[2]),
                        }
                    )
                elif isinstance(cm, dict):
                    col_map.append(
                        {
                            "excel_col": int(cm.get("excel_col", cm.get("source", 0))),
                            "tmpl_col": int(cm.get("tmpl_col", cm.get("target", 0))),
                            "is_pct": bool(cm.get("is_pct", False)),
                        }
                    )

            # 构建基础条目
            entry = {
                "cat": template_name or f"表{table_idx}",
                "sheet_kw": sheet_name,
                "table_idx": table_idx,
                "tz_cols": tz_cols,
                "col_map": [
                    [c["excel_col"], c["tmpl_col"], c["is_pct"]] for c in col_map
                ],
            }

            # 可选字段（如有则追加）
            data_start = match.get("data_start")
            if data_start is not None:
                entry["data_start"] = data_start

            name_strip = match.get("name_strip", [])
            if name_strip:
                entry["name_strip"] = name_strip

            name_exclude = match.get("name_exclude", [])
            if name_exclude:
                entry["name_exclude"] = name_exclude

            section_headers = match.get("section_headers", [])
            if section_headers:
                entry["section_headers"] = section_headers

            aggregations = match.get("aggregations", [])
            if aggregations:
                entry["aggregations"] = aggregations

            accumulate = match.get("accumulate", False)
            if accumulate:
                entry["accumulate"] = True

            keep_empty = match.get("keep_empty", False)
            if keep_empty:
                entry["keep_empty"] = True

            # 置信度过低时添加说明注释（不影响 YAML dump，仅作参考）
            if confidence < 0.5 and "confidence" not in entry:
                entry["_note"] = f"置信度: {confidence:.2f} — 需人工确认"

            return entry

        except Exception as e:
            print(f"  [WARN] 构建映射条目失败: {e}", file=sys.stderr)
            return None

    # ── name_synonyms.yaml 生成 ─────────────────────────────

    def _generate_name_synonyms_yaml(
        self,
        template_analysis: dict,
        excel_profile: dict,
        match_result: dict,
    ) -> dict:
        """
        扫描模板和套表，检测名称变异并生成同义词映射。

        检测策略:
          1. 已知会计科目 → 直接添加标准映射
          2. 含/不含前缀（"一、""（一）"等）的同名行
          3. 全角/半角/多空格变异（"合   计" vs "合计"）
          4. 来自匹配结果中 row_matches 的对应关系
          5. 常用同义词组（"期初" vs "年初"等）
        """
        synonyms: Dict[str, list] = {}

        # --- 策略1: 从匹配结果的 row_matches 提取 ---
        for m in match_result.get("matches", []):
            row_matches = m.get("row_matches", [])
            for rm in row_matches:
                if not isinstance(rm, dict):
                    continue
                tmpl_name = rm.get("template_name", "")
                excel_name = rm.get("excel_name", "")
                if tmpl_name and excel_name and tmpl_name != excel_name:
                    tmpl_clean = _normalize_whitespace(tmpl_name)
                    excel_clean = _normalize_whitespace(excel_name)
                    if tmpl_clean != excel_clean:
                        if tmpl_name not in synonyms:
                            synonyms[tmpl_name] = []
                        if excel_name not in synonyms[tmpl_name]:
                            synonyms[tmpl_name].append(excel_name)

        # --- 策略2: 从模板和套表的行名交叉分析 ---
        tmpl_names = self._collect_template_row_names(template_analysis)
        excel_names = self._collect_excel_row_names(excel_profile, match_result)

        # 2a: 核心名称相同但前缀不同的行
        core_to_tmpl = defaultdict(list)  # 核心名称 → [模板行名列表]
        core_to_excel = defaultdict(list)  # 核心名称 → [套表行名列表]

        for tn in tmpl_names:
            core = _extract_core_name(tn)
            if core and len(core) >= 2:
                core_to_tmpl[core].append(tn)

        for en in excel_names:
            core = _extract_core_name(en)
            if core and len(core) >= 2:
                core_to_excel[core].append(en)

        # 对每个核心名称，如果模板行和套表行名称不同，生成映射
        for core, t_names in core_to_tmpl.items():
            e_names = core_to_excel.get(core, [])
            for tn in t_names:
                if tn not in synonyms:
                    synonyms[tn] = []
                for en in e_names:
                    if en != tn and en not in synonyms[tn]:
                        synonyms[tn].append(en)

        # 2b: 已知会计科目 → 添加标准同义词
        known_synonyms = {
            "货币资金": ["库存现金", "银行存款", "其他货币资金"],
            "合　计": ["合计", "合  计", "合   计"],
            "合      计": ["合计", "合  计"],
            "合  计": ["合计", "合   计"],
            "应收账款": ["应收账"],
            "其他应收款": ["其他应收", "其他应收款"],
            "固定资产": ["固定"],
            "预付账款": ["预付"],
            "应付职工薪酬": ["应付职工", "应付薪酬"],
            "应交税费": ["应交税", "应交税金"],
            "实收资本": ["实收"],
            "营业收入": ["主营业务收入", "营业收"],
            "营业成本": ["主营业务成本", "营业成"],
            "管理费用": ["管理费"],
            "销售费用": ["销售费"],
            "研发费用": ["研发费"],
            "财务费用": ["财务费"],
            "信用减值损失": ["信用减值"],
            "资产减值损失": ["资产减值"],
            "年初余额": ["期初余额", "年初余额", "上年年末余额"],
            "年末余额": ["期末余额", "年末余额"],
            "本期增加": ["本年增加"],
            "本期减少": ["本年减少"],
            "坏账准备": ["坏账准备金额"],
            "账面价值": ["账面净额", "净额"],
            "合   计": ["合计", "小计"],
            "小计": ["合计"],
        }
        for tmpl_name, aliases in known_synonyms.items():
            if tmpl_name not in synonyms:
                synonyms[tmpl_name] = []
            for alias in aliases:
                if alias not in synonyms[tmpl_name]:
                    synonyms[tmpl_name].append(alias)

        # --- 策略3: 从 learn_from_existing 缓存补充 ---
        for tmpl_name, aliases in self._learned_patterns.get("synonyms", {}).items():
            if tmpl_name not in synonyms:
                synonyms[tmpl_name] = []
            for a in aliases:
                if a not in synonyms[tmpl_name]:
                    synonyms[tmpl_name].append(a)

        # --- 排序输出 ---
        sorted_synonyms = {}
        for key in sorted(synonyms.keys(), key=lambda x: x):
            sorted_synonyms[key] = _sorted_unique(synonyms[key])

        return sorted_synonyms

    def _collect_template_row_names(self, template_analysis: dict) -> List[str]:
        """从模板分析结果中提取所有行名"""
        names: List[str] = []

        tables = template_analysis.get("tables", [])
        if not tables and "sheets" in template_analysis:
            # 兼容：可能是 ExcelProfile 格式
            return names

        for table in tables:
            rows = table.get("rows", [])
            for row in rows:
                name = row.get("name", "").strip()
                if name:
                    names.append(name)
            # 同时收集小节标题行
            for sh in table.get("section_headers", []):
                if sh.strip() and sh.strip() not in names:
                    names.append(sh.strip())

        return _sorted_unique(names)

    def _collect_excel_row_names(
        self,
        excel_profile: dict,
        match_result: dict,
    ) -> List[str]:
        """从套表分析结果中提取所有行名"""
        names: List[str] = []

        sheets = excel_profile.get("sheets", [])
        for sheet in sheets:
            rows = sheet.get("rows", [])
            for row in rows:
                name = row.get("row_name", "").strip()
                if name:
                    names.append(name)
            # 从 accounts 中提取
            for acct in sheet.get("accounts", []):
                if acct.strip() and acct.strip() not in names:
                    names.append(acct.strip())

        # 从匹配结果补充
        for m in match_result.get("matches", []):
            row_matches = m.get("row_matches", [])
            for rm in row_matches:
                if isinstance(rm, dict):
                    en = rm.get("excel_name", "").strip()
                    if en and en not in names:
                        names.append(en)

        return _sorted_unique(names)

    # ── auto_sum_rules.yaml 生成 ────────────────────────────

    def _generate_auto_sum_rules(
        self,
        template_analysis: dict,
        match_result: dict,
    ) -> dict:
        """
        从模板分析结果生成自动求和规则。

        规则格式:
            <table_idx>:
                totals: [<行索引>, ...]    # 需要自动求和的合计行
                subtotals: [<行索引>, ...]  # 小计行
        """
        rules: Dict[str, dict] = {}

        if not template_analysis:
            return rules

        tables = template_analysis.get("tables", [])
        for table in tables:
            if table.get("error"):
                continue

            idx = table.get("index")
            if idx is None:
                continue

            total_rows = table.get("total_rows", [])
            subtotal_rows = table.get("subtotal_rows", [])

            # 只生成有合计行的规则
            if total_rows or subtotal_rows:
                key = str(idx)
                entry = {}
                if total_rows:
                    entry["totals"] = total_rows
                if subtotal_rows:
                    entry["subtotals"] = subtotal_rows
                rules[key] = entry

        # 从匹配结果补充聚合规则
        for m in match_result.get("matches", []):
            aggregations = m.get("aggregations", [])
            if aggregations:
                idx = str(m.get("template_table_index"))
                if idx not in rules:
                    rules[idx] = {}
                rules[idx]["aggregations"] = aggregations

        return rules

    # ── adaptation_report.md 生成 ───────────────────────────

    def _generate_adaptation_report(
        self,
        match_result: dict,
        template_analysis: dict,
        excel_profile: dict,
        stats: dict,
        issues: list,
    ) -> str:
        """
        生成适配报告（可读 MD 格式，中文）。

        包含章节:
          1. 摘要: 总表数/匹配数/置信度分布
          2. 高置信度匹配 (>=0.8): 自动接受
          3. 中置信度匹配 (0.5-0.8): 需审核
          4. 低置信度/无匹配 (<0.5): 需手动设置
          5. 列映射详情
          6. 下一步操作指引
        """
        lines = []
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        lines.append("# 智能适配报告")
        lines.append("")
        lines.append(f"**生成时间**: {ts}")
        lines.append(f"**生成工具**: ConfigGenerator")
        lines.append("")
        lines.append("---")
        lines.append("")

        # ── 章节1: 摘要 ──
        lines.append("## 一、适配摘要")
        lines.append("")
        num_template_tables = template_analysis.get("num_tables", 0) or len(
            template_analysis.get("tables", [])
        )
        num_excel_sheets = excel_profile.get("num_sheets", 0) or len(
            excel_profile.get("sheets", [])
        )
        total_matches = sum(
            stats.get(k, 0)
            for k in ["high_confidence", "medium_confidence", "low_confidence"]
        )

        lines.append(f"- **模板表格数**: {num_template_tables}")
        lines.append(f"- **Excel Sheet 数**: {num_excel_sheets}")
        lines.append(f"- **成功匹配**: {total_matches}")
        lines.append(f"- **未匹配模板表**: {stats.get('unmatched_tables', 0)}")
        lines.append(f"- **未匹配 Sheet**: {stats.get('unmatched_sheets', 0)}")
        lines.append("")
        lines.append("| 置信度级别 | 数量 | 处理方式 |")
        lines.append("|-----------|------|---------|")
        lines.append(f"| 高 (≥0.8) | {stats.get('high_confidence', 0)} | 自动接受 |")
        lines.append(
            f"| 中 (0.5-0.8) | {stats.get('medium_confidence', 0)} | 建议人工审核 |"
        )
        lines.append(f"| 低 (<0.5) | {stats.get('low_confidence', 0)} | 需手动确认 |")
        lines.append("")

        # ── 章节2: 高置信度匹配 ──
        lines.append("## 二、高置信度匹配（自动接受）")
        lines.append("")
        high_matches = [
            m for m in match_result.get("matches", []) if m.get("confidence", 0) >= 0.8
        ]
        if high_matches:
            lines.append("| 模板表 | Excel Sheet | 置信度 | 列映射 |")
            lines.append("|--------|------------|--------|--------|")
            for m in high_matches:
                tname = m.get(
                    "template_table_name", f"表{m.get('template_table_index', '?')}"
                )
                sname = m.get("excel_sheet", "—")
                conf = m.get("confidence", 0)
                col_count = len(m.get("col_map", []))
                lines.append(f"| {tname} | {sname} | {conf:.0%} | {col_count} 列 |")
        else:
            lines.append("（无）")
        lines.append("")

        # ── 章节3: 中等置信度匹配 ──
        lines.append("## 三、中等置信度匹配（建议审核）")
        lines.append("")
        mid_matches = [
            m
            for m in match_result.get("matches", [])
            if 0.5 <= m.get("confidence", 0) < 0.8
        ]
        if mid_matches:
            lines.append("| 模板表 | Excel Sheet | 置信度 | 列映射 | 说明 |")
            lines.append("|--------|------------|--------|--------|------|")
            for m in mid_matches:
                tname = m.get(
                    "template_table_name", f"表{m.get('template_table_index', '?')}"
                )
                sname = m.get("excel_sheet", "—")
                conf = m.get("confidence", 0)
                col_count = len(m.get("col_map", []))
                reason = m.get("review_reason", "列映射需确认")
                lines.append(
                    f"| {tname} | {sname} | {conf:.0%} | {col_count} 列 | {reason} |"
                )
        else:
            lines.append("（无）")
        lines.append("")

        # ── 章节4: 低置信度/无匹配 ──
        lines.append("## 四、低置信度/无匹配（需手动设置）")
        lines.append("")

        # 低置信度匹配
        low_matches = [
            m for m in match_result.get("matches", []) if m.get("confidence", 0) < 0.5
        ]
        if low_matches:
            lines.append("### 4.1 低置信度匹配")
            lines.append("")
            lines.append("| 模板表 | 建议 Sheet | 置信度 | 建议操作 |")
            lines.append("|--------|-----------|--------|---------|")
            for m in low_matches:
                tname = m.get(
                    "template_table_name", f"表{m.get('template_table_index', '?')}"
                )
                sname = m.get("excel_sheet", "—")
                conf = m.get("confidence", 0)
                lines.append(
                    f"| {tname} | {sname} | {conf:.0%} | 手动确认配对和列映射 |"
                )
            lines.append("")

        # 未匹配模板表
        unmatched_tables = match_result.get("unmatched_templates", [])
        if unmatched_tables:
            lines.append("### 4.2 未匹配的模板表")
            lines.append("")
            lines.append("| 表索引 | 模板表名 | 建议操作 |")
            lines.append("|--------|---------|---------|")
            for ut in unmatched_tables:
                tidx = ut.get("index", "?")
                tname = ut.get("name", "（未识别）")
                lines.append(f"| {tidx} | {tname} | 手动指定对应的 Excel Sheet |")
            lines.append("")

        # 未匹配 Sheet
        unmatched_sheets = match_result.get("unmatched_sheets", [])
        if unmatched_sheets:
            lines.append("### 4.3 未匹配的 Excel Sheet")
            lines.append("")
            lines.append(
                "以下 Excel Sheet 未能匹配到任何模板表，可能是多余数据或遗漏的映射："
            )
            lines.append("")
            for us in unmatched_sheets:
                lines.append(f"- **{us}**")
            lines.append("")

        # ── 章节5: 列映射详情 ──
        lines.append("## 五、列映射详情")
        lines.append("")
        for m in match_result.get("matches", []):
            tname = m.get(
                "template_table_name", f"表{m.get('template_table_index', '?')}"
            )
            sname = m.get("excel_sheet", "—")
            conf = m.get("confidence", 0)
            col_map = m.get("col_map", [])

            lines.append(f"### {tname} → {sname} (置信度: {conf:.0%})")
            lines.append("")
            if col_map:
                lines.append("| Excel 列 | → 模板列 | 是否比例 |")
                lines.append("|---------|---------|---------|")
                for cm in col_map:
                    if isinstance(cm, (list, tuple)) and len(cm) >= 3:
                        ec, tc, is_pct = cm[0], cm[1], cm[2]
                    elif isinstance(cm, dict):
                        ec = cm.get("excel_col", cm.get("source", "?"))
                        tc = cm.get("tmpl_col", cm.get("target", "?"))
                        is_pct = cm.get("is_pct", False)
                    else:
                        continue
                    pct_mark = "是 ✓" if is_pct else "否"
                    lines.append(f"| 第 {ec} 列 | 第 {tc} 列 | {pct_mark} |")
            else:
                lines.append("（无列映射，需手动配置）")
            lines.append("")

        # ── 章节6: 下一步操作 ──
        lines.append("## 六、下一步操作指引")
        lines.append("")
        lines.append("根据适配结果，请按以下步骤处理：")
        lines.append("")
        lines.append(
            "1. **审核配置文件** — 检查 `config_generated/` 目录下的 YAML 文件"
        )

        if stats.get("unmatched_tables", 0) > 0:
            lines.append(
                "2. **处理未匹配表** — 对未匹配的模板表，在 mappings.yaml 中补充映射"
            )
            lines.append("   - 编辑 mappings.yaml，添加新的映射条目")
            lines.append("   - 指定 sheet_kw（Excel Sheet 关键词）")
            lines.append("   - 配置 tz_cols 和 col_map 列映射")
            step_offset = 3
        else:
            step_offset = 2

        if stats.get("medium_confidence", 0) > 0:
            lines.append(
                f"{step_offset}. **审核中置信度映射** — 检查中等置信度项，确认列映射正确"
            )
            step_offset += 1

        if stats.get("low_confidence", 0) > 0:
            lines.append(
                f"{step_offset}. **确认低置信度映射** — 逐一确认低置信度项的配对和列映射"
            )
            step_offset += 1

        lines.append(
            f"{step_offset}. **运行验证** — 使用填充引擎在小范围数据上测试生成的配置"
        )
        lines.append(f"   - 选取 1-2 个表做测试填充")
        lines.append(f"   - 检查合计行计算是否正确")
        lines.append(f"   - 检查比例%列是否显示正确")
        lines.append(
            f"{step_offset + 1}. **确认后应用** — 将 `config_generated/` 目录复制为"
        )
        lines.append(f"   `config/`，开始正式填充")
        lines.append("")

        # ── 文件清单 ──
        lines.append("## 七、生成的文件")
        lines.append("")
        lines.append(f"配置文件已输出到 `{self.output_dir.resolve()}/`：")
        lines.append("")
        lines.append("| 文件 | 用途 |")
        lines.append("|------|------|")
        lines.append("| `mappings.yaml` | 表-列映射配置 |")
        lines.append("| `name_synonyms.yaml` | 行名同义词映射 |")
        lines.append("| `auto_sum_rules.yaml` | 合计行求和规则 |")
        lines.append("| `adaptation_report.md` | 本适配报告 |")
        lines.append("")

        return "\n".join(lines)

    # ── 从已有配置学习 ──────────────────────────────────────

    def learn_from_existing(self, existing_config_dir: str) -> dict:
        """
        加载已有的手工调优配置，作为未来适配的参考。

        读取 mappings.yaml 和 name_synonyms.yaml 中的模式，
        存储在 self._learned_patterns 中以供后续 generate() 使用。

        参数:
            existing_config_dir: 已有配置目录的路径

        返回:
            dict: 学习到的模式统计
                { "patterns_learned": int, "synonyms_learned": int }
        """
        config_path = Path(existing_config_dir)
        if not config_path.exists() or not config_path.is_dir():
            print(f"  [WARN] 配置目录不存在: {config_path}", file=sys.stderr)
            return {"patterns_learned": 0, "synonyms_learned": 0}

        patterns_learned = 0
        synonyms_learned = 0

        # 学习 mappings.yaml 中的模式
        mappings_file = config_path / "mappings.yaml"
        if mappings_file.exists():
            try:
                with open(mappings_file, "r", encoding="utf-8") as f:
                    mappings = yaml.safe_load(f) or []

                for m in mappings:
                    if not isinstance(m, dict):
                        continue
                    cat = m.get("cat", "")
                    sheet_kw = m.get("sheet_kw", "")
                    if cat and sheet_kw:
                        # 学习: 表名 → Sheet 关键词
                        self._learned_patterns.setdefault("template_to_sheet", {})[
                            cat
                        ] = sheet_kw
                        patterns_learned += 1

                print(
                    f"  [LEARN] 从 {mappings_file.name} 学习到 {patterns_learned} 个映射模式"
                )

            except Exception as e:
                print(f"  [WARN] 读取 mappings.yaml 失败: {e}", file=sys.stderr)

        # 学习 name_synonyms.yaml 中的模式
        synonyms_file = config_path / "name_synonyms.yaml"
        if synonyms_file.exists():
            try:
                with open(synonyms_file, "r", encoding="utf-8") as f:
                    synonyms = yaml.safe_load(f) or {}

                for tmpl_name, aliases in synonyms.items():
                    if isinstance(aliases, list) and aliases:
                        self._learned_patterns.setdefault("synonyms", {})[tmpl_name] = (
                            aliases
                        )
                        synonyms_learned += 1

                print(
                    f"  [LEARN] 从 {synonyms_file.name} 学习到 {synonyms_learned} 个同义词模式"
                )

            except Exception as e:
                print(f"  [WARN] 读取 name_synonyms.yaml 失败: {e}", file=sys.stderr)

        # 学习 auto_sum_rules.yaml 中的模式（可选）
        sum_file = config_path / "auto_sum_rules.yaml"
        if sum_file.exists():
            try:
                with open(sum_file, "r", encoding="utf-8") as f:
                    sum_rules = yaml.safe_load(f) or {}
                if sum_rules:
                    self._learned_patterns["sum_rules"] = sum_rules
            except Exception as e:
                print(f"  [WARN] 读取 auto_sum_rules.yaml 失败: {e}", file=sys.stderr)

        return {
            "patterns_learned": patterns_learned,
            "synonyms_learned": synonyms_learned,
        }

    # ── 便捷工具方法 ─────────────────────────────────────────

    def preview(self, match_result: dict) -> str:
        """
        预览生成的配置摘要（不写文件）。

        参数:
            match_result: SmartMatcher 匹配结果

        返回:
            str: 格式化的文本摘要
        """
        lines = []
        lines.append("=" * 60)
        lines.append("  配置生成预览")
        lines.append("=" * 60)

        matches = match_result.get("matches", [])
        lines.append(f"  总匹配数: {len(matches)}")

        # 分组
        high = [m for m in matches if m.get("confidence", 0) >= 0.8]
        mid = [m for m in matches if 0.5 <= m.get("confidence", 0) < 0.8]
        low = [m for m in matches if m.get("confidence", 0) < 0.5]

        lines.append(f"  高置信度 (≥0.8): {len(high)}")
        lines.append(f"  中置信度 (0.5-0.8): {len(mid)}")
        lines.append(f"  低置信度 (<0.5): {len(low)}")

        unmatched_tables = match_result.get("unmatched_templates", [])
        unmatched_sheets = match_result.get("unmatched_sheets", [])
        lines.append(f"  未匹配模板表: {len(unmatched_tables)}")
        lines.append(f"  未匹配 Sheet: {len(unmatched_sheets)}")

        lines.append("")
        lines.append("  --- 匹配详情 ---")

        for m in matches:
            tname = m.get(
                "template_table_name", f"表{m.get('template_table_index', '?')}"
            )
            sname = m.get("excel_sheet", "?")
            conf = m.get("confidence", 0)
            col_count = len(m.get("col_map", []))
            label = _confidence_label(conf)
            lines.append(
                f"  [{label}] {tname} ↔ {sname} (conf={conf:.2f}, {col_count}列)"
            )

        if unmatched_tables:
            lines.append("")
            lines.append("  --- 未匹配模板表 ---")
            for ut in unmatched_tables:
                lines.append(f"  - 表{ut.get('index', '?')}: {ut.get('name', '?')}")

        if unmatched_sheets:
            lines.append("")
            lines.append("  --- 未匹配 Sheet ---")
            for us in unmatched_sheets:
                lines.append(f"  - {us}")

        lines.append("")
        lines.append("=" * 60)

        return "\n".join(lines)

    @property
    def learned_patterns(self) -> dict:
        """已学习的模式（只读）"""
        return dict(self._learned_patterns)


# ============================================================
# CLI 入口
# ============================================================


def main():
    """
    命令行入口 — 从 JSON 文件读取匹配结果并生成配置。

    用法:
        python config_generator.py --match match_result.json \\
            --template template_analysis.json \\
            --excel excel_profile.json \\
            --output config_generated
    """
    import argparse

    parser = argparse.ArgumentParser(
        description="配置生成器 — 从智能匹配结果生成 YAML 配置文件",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python config_generator.py \\
    --match match_result.json \\
    --template template_analysis.json \\
    --excel excel_profile.json \\
    --output config_generated

  python config_generator.py \\
    --match match_result.json \\
    --template template_analysis.json \\
    --excel excel_profile.json \\
    --learn-from existing_config/ \\
    --output config_generated_v2

  python config_generator.py --preview match_result.json
        """,
    )
    parser.add_argument("--match", "-m", required=True, help="智能匹配结果 JSON 文件")
    parser.add_argument("--template", "-t", default=None, help="模板分析结果 JSON 文件")
    parser.add_argument("--excel", "-e", default=None, help="Excel 剖析报告 JSON 文件")
    parser.add_argument("--output", "-o", default="config_generated", help="输出目录")
    parser.add_argument(
        "--learn-from", "-l", default=None, help="已有配置目录（用于学习模式）"
    )
    parser.add_argument("--preview", "-p", action="store_true", help="仅预览，不写文件")
    parser.add_argument("--quiet", "-q", action="store_true", help="静默模式，最小输出")

    args = parser.parse_args()

    # 读取输入
    def load_json(filepath: str, desc: str) -> dict:
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"错误: 读取 {desc} 失败 — {e}", file=sys.stderr)
            sys.exit(1)

    match_result = load_json(args.match, "匹配结果")

    template_analysis = {}
    if args.template:
        template_analysis = load_json(args.template, "模板分析结果")

    excel_profile = {}
    if args.excel:
        excel_profile = load_json(args.excel, "Excel 剖析报告")

    # 预览模式
    if args.preview:
        gen = ConfigGenerator()
        print(gen.preview(match_result))
        return

    # 完整生成模式
    gen = ConfigGenerator(output_dir=args.output)

    # 从已有配置学习
    if args.learn_from:
        if not args.quiet:
            print(f"正在从 {args.learn_from} 学习模式...")
        learned = gen.learn_from_existing(args.learn_from)
        if not args.quiet:
            print(
                f"  已学习: {learned['patterns_learned']} 个映射模式, {learned['synonyms_learned']} 个同义词"
            )

    # 生成配置
    if not args.quiet:
        print(f"正在生成配置到 {args.output}/ ...")

    result = gen.generate(match_result, template_analysis, excel_profile)

    # 输出结果
    if not args.quiet:
        print(f"  生成状态: {result['status']}")
        print(f"  生成文件:")
        for fname, fstatus in result["files"].items():
            print(f"    - {fname}: {fstatus}")
        print(f"  统计:")
        for k, v in result["stats"].items():
            print(f"    {k}: {v}")
        if result["has_issues"]:
            print(f"  ⚠️ 发现 {len(result['issues'])} 个需关注的问题:")
            for issue in result["issues"][:5]:
                print(f"    - {issue}")
            if len(result["issues"]) > 5:
                print(f"    ... 还有 {len(result['issues']) - 5} 个（详见适配报告）")
        print(f"  适配报告: {result['config_dir']}/adaptation_report.md")
    else:
        # 静默模式：输出 JSON 格式结果
        compact = {
            "status": result["status"],
            "config_dir": result["config_dir"],
            "files": list(result["files"].keys()),
            "stats": result["stats"],
            "has_issues": result["has_issues"],
            "issue_count": len(result["issues"]),
        }
        print(json.dumps(compact, ensure_ascii=False))


if __name__ == "__main__":
    main()
