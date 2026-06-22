#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
智能匹配引擎 — SmartMatcher
============================
核心功能：将 TemplateAnalyzer（模板分析）的输出与 ExcelProfiler（套表剖析）的输出
进行智能匹配，自动生成驱动填表引擎的 MAPPINGS 配置。

工作流：
  1. 接收 template_analysis (dict) + excel_profile (dict)
  2. 对每个模板表格，找到最匹配的 Excel sheet（加权评分）
  3. 生成列映射（col_map）
  4. 推断行匹配策略
  5. 输出完整 mappings 配置，附带置信度评分和警告

用法:
  from smart_matcher import SmartMatcher
  matcher = SmartMatcher()
  result = matcher.match(template_analysis, excel_profile)
  for m in result["mappings"]:
      print(f"  {m['cat']} -> sheet_kw={m['sheet_kw']} confidence={m['confidence']}")

输出格式:
  {
    "status": "ok" | "partial" | "low_confidence",
    "mappings": [...],       # 驱动填表引擎的配置
    "unmatched_tables": [],  # 无匹配的模板表
    "unmatched_sheets": [],  # 无匹配的 Excel sheet
    "confidence": 0.85,      # 整体置信度
    "warnings": [],          # 警告列表
  }

依赖:
  - re (标准库)
  - 无第三方依赖（输入已由上游处理好）
"""

import re
import json
import logging
from typing import Dict, List, Optional, Tuple, Any

logger = logging.getLogger(__name__)


# ============================================================
# 共享工具函数
# ============================================================


def match_name(doc_name: str, tz_name: str) -> bool:
    """
    模糊名称匹配 — 与 fill_notes_config.py 中的 match_name 一致。

    匹配规则：
      1. 完全一致（修剪空白后）
      2. 全半角括号归一化后一致
      3. 长串（>=4 字）包含关系
      4. 去除"及/与/和/、/，"等连接词后一致

    Args:
        doc_name: 文档侧名称（模板表名）
        tz_name: 套表侧名称（Excel sheet 名或行名）

    Returns:
        bool: 是否匹配
    """
    d = re.sub(r"\s+", "", str(doc_name).replace("\u3000", ""))
    t = re.sub(r"\s+", "", str(tz_name).replace("\u3000", ""))
    if not d or not t:
        return False
    if d == t:
        return True
    # 全半角括号归一化
    d2 = d.replace("（", "(").replace("）", ")")
    t2 = t.replace("（", "(").replace("）", ")")
    if d2 == t2:
        return True
    # 长串包含
    if len(d2) >= 4 and d2 in t2:
        return True
    if len(t2) >= 4 and t2 in d2:
        return True
    # 去除连接词后比较
    for ch in "及与和、，":
        d = d.replace(ch, "")
        t = t.replace(ch, "")
    if d == t:
        return True
    return False


def clean_text(text: str) -> str:
    """清洗文本：去空白、全角空格、不可见字符"""
    if not text:
        return ""
    text = re.sub(r"\s+", "", str(text))
    text = text.replace("\u3000", "")
    text = text.replace("\n", "").replace("\r", "")
    return text.strip()


def normalize_punctuation(text: str) -> str:
    """归一化标点：全角括号→半角，逗号→统一"""
    text = text.replace("（", "(").replace("）", ")")
    text = text.replace("，", ",").replace("、", ",")
    text = text.replace("；", ";").replace("：", ":")
    text = re.sub(r"\s+", "", text)
    return text


def extract_keywords(text: str) -> List[str]:
    """从文本中提取有意义的匹配关键词（>=2 字的中文或英文词）"""
    text = clean_text(text)
    if not text:
        return []
    # 中文词
    chinese = re.findall(r"[\u4e00-\u9fff]{2,}", text)
    # 英文词
    english = re.findall(r"[a-zA-Z]{2,}", text)
    # 数字+单位组合
    numeric = re.findall(r"\d+[\.\d]*", text)
    return chinese + english + numeric


# ============================================================
# 标准附注表模式库
# ============================================================

STANDARD_TABLE_PATTERNS = [
    {
        "name": "货币资金",
        "expected_cols": ["期末余额", "期初余额"],
        "row_keywords": ["库存现金", "银行存款", "其他货币资金"],
    },
    {
        "name": "交易性金融资产",
        "expected_cols": ["期末余额", "期初余额"],
        "row_keywords": ["指定以公允价值计量", "分类为以公允价值计量"],
    },
    {
        "name": "应收票据",
        "expected_cols": [
            "期末账面余额",
            "期末坏账准备",
            "期末账面价值",
            "期初账面余额",
            "期初坏账准备",
            "期初账面价值",
        ],
        "row_keywords": ["银行承兑汇票", "商业承兑汇票"],
    },
    {
        "name": "应收账款",
        "expected_cols": [
            "期末账面余额",
            "期末坏账准备",
            "期末账面价值",
            "期初账面余额",
            "期初坏账准备",
            "期初账面价值",
        ],
        "row_keywords": ["1年以内", "1-2年", "2-3年", "3年以上"],
    },
    {
        "name": "应收账款-账龄",
        "expected_cols": [
            "期末账面余额",
            "期末坏账准备",
            "期初账面余额",
            "期初坏账准备",
        ],
        "row_keywords": ["1年以内", "1-2年", "2-3年", "3年以上"],
    },
    {
        "name": "应收账款-组合",
        "expected_cols": [
            "期末账面余额",
            "期末坏账准备",
            "期初账面余额",
            "期初坏账准备",
        ],
        "row_keywords": ["账龄组合", "关联方组合", "低风险组合"],
    },
    {
        "name": "预付款项",
        "expected_cols": ["期末金额", "期末比例", "期初金额", "期初比例"],
        "row_keywords": ["1年以内", "1-2年", "2-3年", "3年以上"],
    },
    {
        "name": "其他应收款",
        "expected_cols": [
            "期末账面余额",
            "期末坏账准备",
            "期末账面价值",
            "期初账面余额",
            "期初坏账准备",
            "期初账面价值",
        ],
        "row_keywords": ["1年以内", "1-2年", "2-3年", "3年以上"],
    },
    {
        "name": "其他应收款-账龄",
        "expected_cols": [
            "期末账面余额",
            "期末坏账准备",
            "期初账面余额",
            "期初坏账准备",
        ],
        "row_keywords": ["1年以内", "1-2年", "2-3年", "3年以上"],
    },
    {
        "name": "存货",
        "expected_cols": ["期末余额", "期初余额"],
        "row_keywords": ["原材料", "库存商品", "在产品", "发出商品", "周转材料"],
    },
    {
        "name": "合同资产",
        "expected_cols": ["期末余额", "期末坏账准备", "期初余额", "期初坏账准备"],
        "row_keywords": ["已完工未结算", "质保金"],
    },
    {
        "name": "固定资产",
        "expected_cols": ["期初余额", "本年增加", "本年减少", "期末余额"],
        "row_keywords": [
            "房屋",
            "建筑物",
            "机器设备",
            "运输设备",
            "办公设备",
            "电子设备",
        ],
    },
    {
        "name": "固定资产-明细",
        "expected_cols": ["期初余额", "本年增加", "本年减少", "期末余额"],
        "row_keywords": ["房屋", "建筑物", "机器设备"],
    },
    {
        "name": "在建工程",
        "expected_cols": ["期末余额", "期初余额"],
        "row_keywords": ["建筑工程", "安装工程", "技术改造"],
    },
    {
        "name": "无形资产",
        "expected_cols": ["期初余额", "本年增加", "本年减少", "期末余额"],
        "row_keywords": ["土地使用权", "专利权", "非专利技术", "软件"],
    },
    {
        "name": "开发支出",
        "expected_cols": ["期初余额", "本年增加", "本年减少", "期末余额"],
        "row_keywords": ["资本化支出", "费用化支出"],
    },
    {
        "name": "长期待摊费用",
        "expected_cols": ["期初余额", "本年增加", "本年摊销", "期末余额"],
        "row_keywords": ["装修费", "改造费"],
    },
    {
        "name": "递延所得税资产",
        "expected_cols": ["期末余额", "期初余额"],
        "row_keywords": ["资产减值准备", "内部交易未实现利润", "可抵扣亏损"],
    },
    {
        "name": "短期借款",
        "expected_cols": ["期末余额", "期初余额"],
        "row_keywords": ["信用借款", "抵押借款", "保证借款"],
    },
    {
        "name": "应付票据",
        "expected_cols": ["期末余额", "期初余额"],
        "row_keywords": ["银行承兑汇票", "商业承兑汇票"],
    },
    {
        "name": "应付账款",
        "expected_cols": ["期末余额", "期初余额"],
        "row_keywords": ["1年以内", "1-2年", "2-3年", "3年以上"],
    },
    {
        "name": "应付账款-账龄",
        "expected_cols": ["期末余额", "期初余额"],
        "row_keywords": ["1年以内", "1-2年", "2-3年", "3年以上"],
    },
    {
        "name": "预收账款",
        "expected_cols": ["期末余额", "期初余额"],
        "row_keywords": ["货款", "工程款"],
    },
    {
        "name": "合同负债",
        "expected_cols": ["期末余额", "期初余额"],
        "row_keywords": ["预收货款", "预收工程款"],
    },
    {
        "name": "应付职工薪酬",
        "expected_cols": ["期初余额", "本期增加", "本期减少", "期末余额"],
        "row_keywords": [
            "工资",
            "奖金",
            "津贴",
            "补贴",
            "社保",
            "住房公积金",
            "工会经费",
            "职工教育经费",
        ],
    },
    {
        "name": "应交税费",
        "expected_cols": ["期初余额", "期末余额"],
        "row_keywords": [
            "增值税",
            "企业所得税",
            "个人所得税",
            "城建税",
            "教育费附加",
            "房产税",
            "土地使用税",
        ],
    },
    {
        "name": "其他应付款",
        "expected_cols": ["期末余额", "期初余额"],
        "row_keywords": ["押金", "保证金", "往来款"],
    },
    {
        "name": "长期借款",
        "expected_cols": ["期末余额", "期初余额"],
        "row_keywords": ["信用借款", "抵押借款", "保证借款"],
    },
    {
        "name": "长期应付款",
        "expected_cols": ["期末余额", "期初余额"],
        "row_keywords": ["应付融资租赁款", "专项应付款"],
    },
    {
        "name": "租赁负债",
        "expected_cols": ["期末余额", "期初余额"],
        "row_keywords": ["租赁付款额", "未确认融资费用"],
    },
    {
        "name": "实收资本",
        "expected_cols": ["期初余额", "本年增加", "本年减少", "期末余额"],
        "row_keywords": ["国家资本", "法人资本", "个人资本"],
    },
    {
        "name": "资本公积",
        "expected_cols": ["期初余额", "本年增加", "本年减少", "期末余额"],
        "row_keywords": ["资本溢价", "其他资本公积"],
    },
    {
        "name": "盈余公积",
        "expected_cols": ["期初余额", "本年增加", "本年减少", "期末余额"],
        "row_keywords": ["法定盈余公积", "任意盈余公积"],
    },
    {
        "name": "未分配利润",
        "expected_cols": ["期初余额", "本年增加", "本年减少", "期末余额"],
        "row_keywords": ["期初未分配利润", "本期净利润"],
    },
    {
        "name": "营业收入",
        "expected_cols": ["本期金额", "上期金额"],
        "row_keywords": ["主营业务收入", "其他业务收入"],
    },
    {
        "name": "营业成本",
        "expected_cols": ["本期金额", "上期金额"],
        "row_keywords": ["主营业务成本", "其他业务成本"],
    },
    {
        "name": "税金及附加",
        "expected_cols": ["本期金额", "上期金额"],
        "row_keywords": ["城建税", "教育费附加", "房产税"],
    },
    {
        "name": "销售费用",
        "expected_cols": ["本期金额", "上期金额"],
        "row_keywords": ["职工薪酬", "折旧费", "办公费", "差旅费"],
    },
    {
        "name": "管理费用",
        "expected_cols": ["本期金额", "上期金额"],
        "row_keywords": ["职工薪酬", "折旧费", "办公费", "差旅费", "中介费"],
    },
    {
        "name": "研发费用",
        "expected_cols": ["本期金额", "上期金额"],
        "row_keywords": ["人员人工", "直接投入", "折旧摊销", "设计试验"],
    },
    {
        "name": "财务费用",
        "expected_cols": ["本期金额", "上期金额"],
        "row_keywords": ["利息费用", "利息收入", "手续费", "汇兑损益"],
    },
    {
        "name": "信用减值损失",
        "expected_cols": ["本期金额", "上期金额"],
        "row_keywords": ["应收账款坏账", "其他应收款坏账"],
    },
    {
        "name": "资产减值损失",
        "expected_cols": ["本期金额", "上期金额"],
        "row_keywords": ["存货跌价", "固定资产减值"],
    },
    {
        "name": "营业外收入",
        "expected_cols": ["本期金额", "上期金额"],
        "row_keywords": ["政府补助", "罚没收入"],
    },
    {
        "name": "营业外支出",
        "expected_cols": ["本期金额", "上期金额"],
        "row_keywords": ["罚款支出", "捐赠支出"],
    },
    {
        "name": "所得税费用",
        "expected_cols": ["本期金额", "上期金额"],
        "row_keywords": ["当期所得税", "递延所得税"],
    },
    {
        "name": "现金流量表",
        "expected_cols": ["本期金额", "上期金额"],
        "row_keywords": ["销售商品", "购买商品", "支付给职工"],
    },
    {
        "name": "现金流量表补充资料",
        "expected_cols": ["本期金额", "上期金额"],
        "row_keywords": ["净利润", "资产减值准备", "折旧", "摊销"],
    },
    {
        "name": "成本费用管理情况",
        "expected_cols": ["本期金额", "上期金额"],
        "row_keywords": ["材料成本", "人工成本", "制造费用"],
    },
]

# 列头类型检测关键词（用于列匹配）
COL_TYPE_KEYWORDS = {
    "期末余额": ["期末余额", "期末数", "年末余额", "年末数", "期末"],
    "期初余额": ["期初余额", "期初数", "年初余额", "年初数", "期初"],
    "本期金额": ["本期金额", "本期数", "本年累计", "本年数"],
    "上期金额": ["上期金额", "上期数", "上年累计", "上年数"],
    "本年增加": ["本年增加", "本期增加", "本年借方", "本期借方"],
    "本年减少": ["本年减少", "本期减少", "本年贷方", "本期贷方"],
    "账面余额": ["账面余额", "账面原值", "原值", "余额"],
    "账面价值": ["账面价值", "净值"],
    "坏账准备": ["坏账准备", "减值准备"],
    "比例%": ["比例", "占比", "百分比", "%"],
}


# ============================================================
# 智能匹配器
# ============================================================


class SmartMatcher:
    """
    智能匹配引擎 — 将模板表格与 Excel 套表进行自动配对。

    使用加权评分综合以下信号：
      - 名称相似度（weight=0.5）
      - 列兼容性（weight=0.3）
      - 行标签重叠（weight=0.2）

    根据总分确定匹配状态：
      >= 0.75  → 自动接受
      >= 0.50  → 需人工确认
      <  0.50  → 无匹配
    """

    # 匹配置信度阈值
    AUTO_ACCEPT_THRESHOLD = 0.75
    MANUAL_REVIEW_THRESHOLD = 0.50
    HIGH_CONFIDENCE_THRESHOLD = 0.85

    # 各信号权重
    WEIGHT_NAME = 0.50
    WEIGHT_COLUMNS = 0.30
    WEIGHT_ROWS = 0.20

    def __init__(
        self,
        table_patterns: Optional[List[dict]] = None,
        semantic_patterns: Optional[List[dict]] = None,
    ):
        """
        初始化匹配器。

        Args:
            table_patterns: 自定义表模式列表，默认使用 STANDARD_TABLE_PATTERNS
            semantic_patterns: 从 semantic_tables.yaml 加载的语义模式，会与 table_patterns 合并
        """
        base = table_patterns or STANDARD_TABLE_PATTERNS
        # 合并语义模式（优先用语义模式，同名的覆盖内置模式）
        if semantic_patterns:
            # 建立索引：同名覆盖
            base_by_name = {}
            for p in base:
                base_by_name[clean_text(p["name"])] = p
            for sp in semantic_patterns:
                name_key = clean_text(sp["name"])
                if name_key in base_by_name:
                    # 合并：语义模式中的字段覆盖内置的
                    base_by_name[name_key].update(sp)
                    base_by_name[name_key]["source"] = "semantic_merged"
                else:
                    # 新增：语义模式中定义的但内置没有的
                    base_by_name[name_key] = sp
            self.table_patterns = list(base_by_name.values())
            semantic_count = len(semantic_patterns)
            total_count = len(self.table_patterns)
            print(
                f"  SmartMatcher: {semantic_count} 条语义模式已合并, 总计 {total_count} 条"
            )
        else:
            self.table_patterns = base

        # 建立模式查找索引（表名→模式）
        self._pattern_index = {}
        for p in self.table_patterns:
            name_key = clean_text(p["name"])
            self._pattern_index[name_key] = p
            # 也索引短名称（如"应收账款"索引"应收账款-账龄"和"应收账款-组合"）
            base_name = clean_text(p["name"]).split("-")[0]
            if base_name != name_key and base_name not in self._pattern_index:
                self._pattern_index[base_name] = p

    # -------------------------------------------------------
    # 主入口
    # -------------------------------------------------------

    def match(self, template_analysis: dict, excel_profile: dict) -> dict:
        """
        主入口：匹配模板表格到 Excel sheet，生成 mappings 配置。

        Args:
            template_analysis: TemplateAnalyzer.analyze() 的输出
                {
                    "tables": [
                        {
                            "index": int,
                            "name": str,          # 推断的表名
                            "confidence": float,  # 表名置信度
                            "headers": [str],     # 列头
                            "rows": [{"index", "name", "is_total", "values"}],
                            "num_rows": int,
                            "num_cols": int,
                            "section_headers": [str],
                            "total_rows": [int],
                            ...
                        },
                    ],
                    ...
                }
            excel_profile: ExcelProfiler.profile() 的输出
                {
                    "sheets": [
                        {
                            "name": str,        # 原始 sheet 名
                            "clean_name": str,  # 清洗后的 sheet 名
                            "category": str,    # balance_sheet / income_statement / ...
                            "columns": [
                                {"index": int, "header": str, "type": str, ...}
                            ],
                            "rows": [{"row_name", "row_index", "values"}],
                            "accounts": set[str],
                            "header_row": int,
                            "has_total": bool,
                            ...
                        },
                    ],
                    ...
                }

        Returns:
            dict: 匹配结果，含 mappings / unmatched / confidence / warnings
        """
        result = {
            "status": "ok",
            "mappings": [],
            "unmatched_tables": [],
            "unmatched_sheets": [],
            "confidence": 0.0,
            "warnings": [],
        }

        # 参数校验
        if not template_analysis or "tables" not in template_analysis:
            result["status"] = "low_confidence"
            result["warnings"].append("模板分析结果为空或格式无效")
            return result

        if not excel_profile or "sheets" not in excel_profile:
            result["status"] = "low_confidence"
            result["warnings"].append("Excel 套表剖析结果为空或格式无效")
            return result

        tables = template_analysis.get("tables", [])
        sheets = excel_profile.get("sheets", [])

        if not tables:
            result["status"] = "low_confidence"
            result["warnings"].append("模板中未发现任何表格")
            return result

        if not sheets:
            result["status"] = "low_confidence"
            result["warnings"].append("Excel 中未发现任何 sheet")
            return result

        # 记录哪些 sheet 已被匹配（防止重复匹配）
        matched_sheet_indices = set()
        matched_table_indices = set()

        # 为每一张模板表找最佳匹配的 Excel sheet
        for tbl in tables:
            tbl_idx = tbl.get("index", -1)
            if tbl_idx < 0:
                continue

            # 跳过无有效行数据的表
            num_rows = tbl.get("num_rows", 0)
            if num_rows < 2:
                continue

            # 为当前模板表查找最佳 sheet 匹配
            match_result = self._match_table_to_sheet(
                tbl, sheets, matched_sheet_indices
            )

            if match_result is None:
                # 无匹配
                result["unmatched_tables"].append(
                    {
                        "index": tbl_idx,
                        "name": tbl.get("name", ""),
                        "reason": "未找到匹配的 Excel sheet",
                    }
                )
                continue

            sheet_idx = match_result["sheet_index"]
            sheet_obj = sheets[sheet_idx]
            score = match_result["score"]

            # 标记已匹配
            matched_sheet_indices.add(sheet_idx)
            matched_table_indices.add(tbl_idx)

            # 列匹配
            col_map = self._match_columns(
                template_cols=tbl.get("headers", []),
                excel_columns=sheet_obj.get("columns", []),
            )

            # 确定 Excel 数据列索引
            tz_cols = self._get_tz_cols(col_map, sheet_obj)

            # 行匹配策略
            row_match = self._match_rows(
                template_rows=tbl.get("rows", []),
                excel_rows=sheet_obj.get("rows", []),
            )

            # 聚合推断
            aggregations = self._infer_aggregations(
                template_table=tbl,
                excel_data=sheet_obj,
            )

            # 构建最终配置
            cat_name = tbl.get("name", f"表{tbl_idx}")
            if cat_name and tbl.get("confidence", 0) < 0.5:
                # 表名置信度低，用索引代替
                cat_name = f"表{tbl_idx}"

            mapping = {
                "cat": cat_name,
                "sheet_kw": sheet_obj.get("name", ""),
                "table_idx": tbl_idx,
                "tz_cols": tz_cols,
                "col_map": col_map,
                "data_start": row_match.get("data_start"),
                "data_end": row_match.get("data_end"),
                "name_col": 0,  # 第一列通常是科目名
                "row_match_mode": row_match.get("mode", "exact"),
                "confidence": round(score, 4),
                "aggregate": aggregations.get("needed", False),
                "aggregate_cols": aggregations.get("cols", []),
                "has_total": sheet_obj.get("has_total", False),
            }

            result["mappings"].append(mapping)

            # 生成低置信度警告
            if score < self.AUTO_ACCEPT_THRESHOLD:
                result["warnings"].append(
                    f"低置信度匹配: 模板表「{cat_name}」(索引{tbl_idx}) "
                    f"→ Excel sheet「{sheet_obj['name']}」(得分{score:.2f})"
                )

        # 记录未匹配的 sheet
        for si, sheet in enumerate(sheets):
            if si not in matched_sheet_indices:
                sheet_name = sheet.get("name", f"sheet_{si}")
                # 排除空 sheet 和无数据 sheet
                if sheet.get("num_rows", 0) >= 2 and sheet.get("num_cols", 0) >= 2:
                    result["unmatched_sheets"].append(
                        {
                            "index": si,
                            "name": sheet_name,
                            "category": sheet.get("category", "other"),
                        }
                    )

        # 计算整体置信度
        result["confidence"] = self._compute_overall_confidence(result)

        # 确定状态
        result["status"] = self._determine_status(result)

        return result

    # -------------------------------------------------------
    # 表→Sheet 匹配（加权评分核心）
    # -------------------------------------------------------

    def _match_table_to_sheet(
        self,
        template_table: dict,
        excel_sheets: List[dict],
        used_indices: set,
    ) -> Optional[dict]:
        """
        为单个模板表寻找最佳的 Excel sheet 匹配。

        使用三组信号加权评分：
          1. 名称相似度 (weight=0.5)
          2. 列兼容性 (weight=0.3)
          3. 行标签重叠 (weight=0.2)

        Args:
            template_table: 模板表结构
            excel_sheets: 所有 Excel sheet 列表
            used_indices: 已被使用的 sheet 索引集合

        Returns:
            dict: {"sheet_index": int, "score": float} 或 None（无合适匹配）
        """
        best_sheet = None
        best_score = 0.0

        tbl_name = template_table.get("name", "")
        tbl_headers = template_table.get("headers", [])
        tbl_rows = template_table.get("rows", [])

        for si, sheet in enumerate(excel_sheets):
            # 跳过已被匹配的 sheet
            if si in used_indices:
                continue

            # 跳过空 sheet 或明显不是数据源的 sheet
            num_rows = sheet.get("num_rows", 0)
            num_cols = sheet.get("num_cols", 0)
            if num_rows < 2 or num_cols < 2:
                continue

            # 计算各维度得分
            name_score = self._name_similarity_score(tbl_name, sheet)
            col_score = self._column_compatibility_score(tbl_headers, sheet)
            row_score = self._row_label_overlap_score(tbl_rows, sheet)

            # 加权总分
            total = (
                name_score * self.WEIGHT_NAME
                + col_score * self.WEIGHT_COLUMNS
                + row_score * self.WEIGHT_ROWS
            )

            logger.debug(
                "Table '%s' vs sheet '%s': name=%.3f col=%.3f row=%.3f total=%.3f",
                tbl_name,
                sheet.get("name", "?"),
                name_score,
                col_score,
                row_score,
                total,
            )

            if total > best_score:
                best_score = total
                best_sheet = si

        # 得分过低 → 无匹配
        if best_score < self.MANUAL_REVIEW_THRESHOLD or best_sheet is None:
            return None

        return {"sheet_index": best_sheet, "score": best_score}

    # -------------------------------------------------------
    # 名称相似度评分 (weight=0.5)
    # -------------------------------------------------------

    def _name_similarity_score(self, tbl_name: str, sheet: dict) -> float:
        """
        计算模板表名与 Excel sheet 的名称相似度。

        匹配策略：
          1. 精确匹配 → 1.0
          2. 全半角归一化后精确匹配 → 1.0
          3. match_name() 模糊匹配通过 → 0.85
          4. 关键词包含关系 → 0.6-0.75
          5. 无匹配 → 0.0

        Args:
            tbl_name: 模板表名（如"货币资金"）
            sheet: Excel sheet 结构

        Returns:
            float: 0.0-1.0 的得分
        """
        if not tbl_name:
            return 0.0

        # 使用多种名称来源进行匹配
        candidate_names = [
            sheet.get("name", ""),
            sheet.get("clean_name", ""),
        ]
        # 如果有科目列表，尝试从中寻找匹配
        accounts = sheet.get("accounts", set())
        if accounts:
            candidate_names.extend(list(accounts)[:5])  # 取前5个作为候选

        cleaned_tbl = clean_text(tbl_name)
        if not cleaned_tbl:
            return 0.0

        best_score = 0.0

        for name in candidate_names:
            if not name:
                continue
            cleaned_name = clean_text(name)
            if not cleaned_name:
                continue

            # 精确匹配
            if cleaned_tbl == cleaned_name:
                return 1.0

            # 全半角归一化后精确匹配
            norm_tbl = normalize_punctuation(cleaned_tbl)
            norm_name = normalize_punctuation(cleaned_name)
            if norm_tbl == norm_name:
                best_score = max(best_score, 1.0)
                continue

            # fuzzy match 通过
            if match_name(tbl_name, name):
                best_score = max(best_score, 0.85)
                continue

            # 关键词包含
            if len(cleaned_tbl) >= 3 and cleaned_tbl in cleaned_name:
                best_score = max(best_score, 0.75)
                continue
            if len(cleaned_name) >= 3 and cleaned_name in cleaned_tbl:
                best_score = max(best_score, 0.70)
                continue

            # 中文关键词部分重叠
            tbl_keywords = extract_keywords(tbl_name)
            name_keywords = extract_keywords(name)
            if tbl_keywords and name_keywords:
                intersection = set(tbl_keywords) & set(name_keywords)
                if intersection:
                    ratio = len(intersection) / max(
                        len(tbl_keywords), len(name_keywords)
                    )
                    best_score = max(best_score, 0.4 + 0.3 * ratio)

        # 尝试通过标准表模式匹配
        pattern = self._find_matching_pattern(cleaned_tbl)
        if pattern:
            # 检查 sheet 名是否包含模式名
            for name in candidate_names:
                cleaned_name = clean_text(name)
                if pattern["name"] in cleaned_name or cleaned_name in pattern["name"]:
                    best_score = max(best_score, 0.75)
                # 检查行关键词
                pattern_kws = pattern.get("row_keywords", [])
                if pattern_kws and accounts:
                    kw_hits = sum(
                        1 for kw in pattern_kws for acct in accounts if kw in acct
                    )
                    if kw_hits >= 2:
                        best_score = max(best_score, 0.70)

        return min(best_score, 1.0)

    def _find_matching_pattern(self, name: str) -> Optional[dict]:
        """通过名称查找匹配的标准表模式"""
        cleaned = clean_text(name)
        if not cleaned:
            return None

        # 直接命中
        if cleaned in self._pattern_index:
            return self._pattern_index[cleaned]

        # 尝试包含匹配
        for key, pattern in self._pattern_index.items():
            if key in cleaned or cleaned in key:
                return pattern

        return None

    # -------------------------------------------------------
    # 列兼容性评分 (weight=0.3)
    # -------------------------------------------------------

    def _column_compatibility_score(self, tbl_headers: List[str], sheet: dict) -> float:
        """
        计算模板列头与 Excel 列头的兼容性。

        策略：
          1. 列数是否匹配（相差越小越好）
          2. 列类型模式匹配
          3. 列头关键词重叠

        Args:
            tbl_headers: 模板表头列表（如["项目","期末余额","期初余额"]）
            sheet: Excel sheet 结构

        Returns:
            float: 0.0-1.0 得分
        """
        sheet_cols = sheet.get("columns", [])

        if not tbl_headers or not sheet_cols:
            return 0.0

        # 排除第一列"项目"类文本列，只比较数据列
        tbl_data_cols = [h for h in tbl_headers if not self._is_label_column(h)]
        sheet_data_cols = [
            c for c in sheet_cols if not self._is_label_column(c.get("header", ""))
        ]

        if not tbl_data_cols or not sheet_data_cols:
            # 如果只有标签列，给予基础分
            return 0.3

        # 1) 列数匹配度（负指数衰减）
        n_tbl = len(tbl_data_cols)
        n_sheet = len(sheet_data_cols)
        col_count_diff = abs(n_tbl - n_sheet)
        count_score = max(0.0, 1.0 - col_count_diff * 0.25)

        # 2) 列类型模式匹配
        type_score = self._match_column_types(tbl_data_cols, sheet_data_cols)

        # 3) 列头关键词重叠
        keyword_score = self._match_column_keywords(tbl_data_cols, sheet_data_cols)

        # 加权综合
        return 0.3 * count_score + 0.4 * type_score + 0.3 * keyword_score

    def _is_label_column(self, header: str) -> bool:
        """判断列头是否为标签列而非数据列"""
        if not header:
            return True
        clean = clean_text(header)
        label_kw = ["项目", "科目", "名称", "行次", "序号", "类别", "注释"]
        for kw in label_kw:
            if kw == clean or kw in clean:
                return True
        # 空列头或无意义列头
        if not clean or clean.startswith("Col_"):
            return True
        return False

    def _match_column_types(self, tbl_cols: List[str], sheet_cols: List[dict]) -> float:
        """
        匹配列类型模式。

        例如模板列类型为[期末余额, 期初余额]，Excel 也匹配到同样的类型序列 → 高分。

        Returns:
            float: 0.0-1.0
        """
        if not tbl_cols or not sheet_cols:
            return 0.0

        # 检测每列的 COL_TYPE 类型
        tbl_types = []
        for h in tbl_cols:
            detected = self._detect_col_type(h)
            tbl_types.append(detected)

        sheet_types = []
        for c in sheet_cols:
            header = c.get("header", "")
            detected = self._detect_col_type(header)
            sheet_types.append(detected)

        # 比较类型序列（允许跳过未知类型）
        matches = 0
        total = 0
        tbl_i = 0

        for s_type in sheet_types:
            if tbl_i >= len(tbl_types):
                break
            t_type = tbl_types[tbl_i]
            if t_type == "未知":
                # 模板列类型未知，尝试匹配下一个
                tbl_i += 1
                continue
            if s_type != "未知":
                total += 1
                if s_type == t_type:
                    matches += 1
                tbl_i += 1
            # Excel 侧未知类型列跳过（可能是中间列）

        if total == 0:
            # 双方都未知，给基础分
            return 0.3 if len(tbl_types) == len(sheet_types) else 0.1

        return matches / total if total > 0 else 0.0

    def _match_column_keywords(
        self, tbl_cols: List[str], sheet_cols: List[dict]
    ) -> float:
        """
        计算列头关键词匹配度。

        提取模板和 Excel 列头中的"期初/期末/增加/减少/比例"等关键词，
        计算交集比例。
        """
        # 从模板列头提取关键词
        tbl_kws = set()
        for h in tbl_cols:
            kws = self._extract_col_keywords(h)
            tbl_kws.update(kws)

        # 从 Excel 列头提取关键词
        sheet_kws = set()
        for c in sheet_cols:
            header = c.get("header", "")
            kws = self._extract_col_keywords(header)
            sheet_kws.update(kws)

        if not tbl_kws or not sheet_kws:
            return 0.0

        intersection = tbl_kws & sheet_kws
        # 查全率（模板的关键词有多少被覆盖）
        recall = len(intersection) / len(tbl_kws)
        # 查准率
        precision = len(intersection) / len(sheet_kws)

        # 调和平均
        if recall + precision == 0:
            return 0.0
        f1 = 2 * recall * precision / (recall + precision)
        return f1

    def _detect_col_type(self, header: str) -> str:
        """检测列头的财务类型"""
        if not header:
            return "未知"
        clean = clean_text(header)
        for col_type, keywords in COL_TYPE_KEYWORDS.items():
            for kw in keywords:
                if kw in clean:
                    return col_type
        return "未知"

    def _extract_col_keywords(self, header: str) -> set:
        """从列头提取财务关键词"""
        if not header:
            return set()
        clean = clean_text(header)
        found = set()
        for col_type, keywords in COL_TYPE_KEYWORDS.items():
            for kw in keywords:
                if kw in clean:
                    found.add(col_type)
                    break
        return found

    # -------------------------------------------------------
    # 行标签重叠评分 (weight=0.2)
    # -------------------------------------------------------

    def _row_label_overlap_score(self, tbl_rows: List[dict], sheet: dict) -> float:
        """
        计算模板行名与 Excel 行名的重叠度。

        策略：
          - 比较第一列行名（科目名）
          - 使用 match_name() 模糊匹配
          - 计算重叠比例

        Args:
            tbl_rows: 模板行列表
            sheet: Excel sheet 结构

        Returns:
            float: 0.0-1.0 得分
        """
        # 提取模板行名（跳过合计/小计/空行）
        tbl_labels = set()
        for r in tbl_rows:
            name = r.get("name", "")
            if not name:
                continue
            clean = clean_text(name)
            if not clean or len(clean) < 2:
                continue
            # 跳过合计行
            if "合计" in clean or "小计" in clean or "总计" in clean:
                continue
            # 跳过小节标题（如"一、账面原值"）
            if re.match(r"^[一二三四五六七八九十]+[、．\.]", clean):
                continue
            tbl_labels.add(clean)

        # 提取 Excel 行名
        excel_rows = sheet.get("rows", [])
        excel_labels = set()
        for r in excel_rows:
            name = r.get("row_name", "")
            if not name:
                continue
            clean = clean_text(name)
            if not clean or len(clean) < 2:
                continue
            if "合计" in clean or "小计" in clean or "总计" in clean:
                continue
            excel_labels.add(clean)

        if not tbl_labels or not excel_labels:
            return 0.0

        # 计算 match_name 匹配的个数
        matched = set()
        for tl in tbl_labels:
            for el in excel_labels:
                if match_name(tl, el):
                    matched.add(tl)
                    break

        # 也可以通过科目列表辅助
        accounts = sheet.get("accounts", set())
        if accounts:
            for tl in tbl_labels:
                for acct in accounts:
                    if match_name(tl, acct):
                        matched.add(tl)
                        break

        # 计算覆盖率（模板行中有多少在 Excel 中找到匹配）
        coverage = len(matched) / len(tbl_labels) if tbl_labels else 0.0

        # 针对短表（3行以下）给予调整
        if len(tbl_labels) <= 3:
            # 短表如果全匹配，加分
            if coverage >= 1.0:
                coverage = min(1.0, coverage + 0.1)
            # 短表如果完全不匹配，不减太多（可能是汇总表）
            elif coverage == 0.0:
                coverage = 0.1

        return min(coverage, 1.0)

    # -------------------------------------------------------
    # 列映射
    # -------------------------------------------------------

    def _match_columns(
        self,
        template_cols: List[str],
        excel_columns: List[dict],
    ) -> List[list]:
        """
        生成模板列 → Excel 列的映射。

        返回格式: [[excel_col_idx, template_col_idx, is_percentage], ...]

        匹配策略：
          1. 按列头关键词匹配（期末余额→期末余额列）
          2. 按位置降级匹配
          3. 百分比列特殊标记

        Args:
            template_cols: 模板列头列表（含第一列"项目"类）
            excel_columns: Excel 列信息列表

        Returns:
            list[list]: col_map，每个元素为 [excel_col_idx, template_col_idx, is_percentage]
        """
        col_map = []

        if not template_cols or not excel_columns:
            return col_map

        # 排除模板的第一列（通常是"项目"标签列）
        # 从索引 1 开始映射（模板侧）
        tmpl_data_start = 0
        for i, h in enumerate(template_cols):
            if self._is_label_column(h):
                tmpl_data_start = i + 1
            else:
                break

        # Excel 侧也跳过标签列
        excel_data_start = 0
        for i, c in enumerate(excel_columns):
            if self._is_label_column(c.get("header", "")):
                excel_data_start = i + 1
            else:
                break

        # 提取数据列
        tmpl_data_cols = template_cols[tmpl_data_start:]
        excel_data_cols = excel_columns[excel_data_start:]

        if not tmpl_data_cols:
            return col_map

        # 记录已映射的 Excel 列索引
        used_excel_indices = set()

        # 第一轮：按关键词匹配
        for ti, t_col in enumerate(tmpl_data_cols):
            t_type = self._detect_col_type(t_col)
            if t_type == "未知":
                continue

            best_ei = None
            for ei, e_col in enumerate(excel_data_cols):
                real_ei = excel_data_start + ei
                if real_ei in used_excel_indices:
                    continue
                e_type = self._detect_col_type(e_col.get("header", ""))
                if e_type == t_type:
                    best_ei = real_ei
                    # 如果同时匹配了比例，优先精确匹配
                    if "比例" in t_type or self._is_pct_col(e_col):
                        if self._is_pct_col(e_col):
                            best_ei = real_ei
                            break
                    break

            if best_ei is not None:
                is_pct = (
                    self._is_pct_col(excel_columns[best_ei - excel_data_start])
                    if (best_ei - excel_data_start) < len(excel_columns)
                    else False
                )
                col_map.append([best_ei, tmpl_data_start + ti, is_pct])
                used_excel_indices.add(best_ei)

        # 第二轮：剩余未匹配的模板列，按位置降级匹配
        matched_t_indices = {item[1] for item in col_map}
        unmatched_t_indices = [
            tmpl_data_start + i
            for i in range(len(tmpl_data_cols))
            if (tmpl_data_start + i) not in matched_t_indices
        ]

        for ti in unmatched_t_indices:
            # 找下一个未使用的 Excel 列
            for ei in range(excel_data_start, len(excel_columns)):
                if ei in used_excel_indices:
                    continue
                # 跳过明显的标签列
                if self._is_label_column(excel_columns[ei].get("header", "")):
                    continue
                is_pct = self._is_pct_col(excel_columns[ei])
                col_map.append([ei, ti, is_pct])
                used_excel_indices.add(ei)
                break

        # 按模板列索引排序
        col_map.sort(key=lambda x: x[1])

        return col_map

    def _is_pct_col(self, col: dict) -> bool:
        """判断 Excel 列是否为百分比列"""
        col_type = col.get("type", "")
        if "比例" in col_type:
            return True
        header = str(col.get("header", ""))
        if "比例" in header or "占比" in header or "%" in header:
            return True
        return False

    def _get_tz_cols(self, col_map: List[list], sheet: dict) -> List[int]:
        """
        从 col_map 提取 Excel 数据列索引列表供 tz_cols 使用。

        Args:
            col_map: _match_columns 的输出
            sheet: Excel sheet 结构

        Returns:
            list[int]: Excel 列索引列表（去重、排序）
        """
        cols = set()
        for entry in col_map:
            if len(entry) >= 2:
                cols.add(entry[0])  # excel_col_idx
        # 如果没有通过 col_map 匹配到任何列，尝试自动推断
        if not cols:
            sheet_cols = sheet.get("columns", [])
            for c in sheet_cols:
                ci = c.get("index", -1)
                if ci >= 0 and not self._is_label_column(c.get("header", "")):
                    cols.add(ci)
        return sorted(cols)

    # -------------------------------------------------------
    # 行匹配策略
    # -------------------------------------------------------

    def _match_rows(self, template_rows: List[dict], excel_rows: List[dict]) -> dict:
        """
        确定行匹配策略。

        Returns:
            dict: {
                "mode": "exact" | "fuzzy" | "section",  # 匹配模式
                "data_start": int | None,                # 数据起始行
                "data_end": int | None,                  # 数据结束行
                "has_total": bool,                       # 是否有合计行
            }
        """
        if not template_rows or not excel_rows:
            return {
                "mode": "exact",
                "data_start": None,
                "data_end": None,
                "has_total": False,
            }

        # 检查是否有合计行
        has_total = any("合计" in clean_text(r.get("name", "")) for r in excel_rows)

        # 检查是否有小节标题（如"一、账面原值"）
        has_sections = any(
            bool(re.match(r"^[一二三四五六七八九十]+[、．\.]", r.get("name", "")))
            for r in excel_rows
        )

        # 检查行名是否精确匹配
        exact_matches = 0
        fuzzy_matches = 0
        for tr in template_rows:
            t_name = clean_text(tr.get("name", ""))
            if not t_name or "合计" in t_name:
                continue
            for er in excel_rows:
                e_name = clean_text(er.get("row_name", ""))
                if not e_name or "合计" in e_name:
                    continue
                if t_name == e_name:
                    exact_matches += 1
                    break
                elif match_name(t_name, e_name):
                    fuzzy_matches += 1
                    break

        n_tbl_rows = len(
            [
                r
                for r in template_rows
                if clean_text(r.get("name", "")) and "合计" not in r.get("name", "")
            ]
        )

        if n_tbl_rows == 0:
            return {
                "mode": "exact",
                "data_start": None,
                "data_end": None,
                "has_total": has_total,
            }

        # 判断模式
        if exact_matches / n_tbl_rows >= 0.7:
            mode = "exact"
        elif fuzzy_matches / n_tbl_rows >= 0.5:
            mode = "fuzzy"
        elif has_sections:
            mode = "section"
        else:
            mode = "fuzzy"

        return {
            "mode": mode,
            "data_start": None,  # 留空由 fill engine 自动检测
            "data_end": None,
            "has_total": has_total,
        }

    # -------------------------------------------------------
    # 聚合推断
    # -------------------------------------------------------

    def _infer_aggregations(
        self,
        template_table: dict,
        excel_data: dict,
    ) -> dict:
        """
        推断是否需要聚合（如多日行程求和、多行合并等）。

        判断条件：
          1. Excel 数据行数 > 模板行数 × 1.5
          2. 模板有合计行但 Excel 数据行分散
          3. 模板列包含"本期金额""本年累计"等累计列

        Args:
            template_table: 模板表结构
            excel_data: Excel sheet 结构

        Returns:
            dict: {
                "needed": bool,       # 是否需要聚合
                "cols": List[int],    # 需要聚合的列索引
            }
        """
        needed = False
        agg_cols = []

        tbl_rows = template_table.get("rows", [])
        excel_rows = excel_data.get("rows", [])

        n_tbl = len([r for r in tbl_rows if clean_text(r.get("name", ""))])
        n_excel = len([r for r in excel_rows if clean_text(r.get("row_name", ""))])

        # 条件 1：行数差异大
        if n_tbl > 0 and n_excel > n_tbl * 1.5:
            needed = True

        # 条件 2：模板列包含累计关键词
        headers = template_table.get("headers", [])
        for i, h in enumerate(headers):
            if any(kw in h for kw in ["本期金额", "本年累计", "合计", "总额"]):
                if i not in agg_cols:
                    agg_cols.append(i)

        # 条件 3：检测到合计行但 Excel 中有多个明细行
        has_total = any("合计" in clean_text(r.get("name", "")) for r in tbl_rows)
        if has_total and n_excel > n_tbl:
            needed = True

        return {
            "needed": needed,
            "cols": sorted(agg_cols),
        }

    # -------------------------------------------------------
    # 置信度计算
    # -------------------------------------------------------

    def _compute_overall_confidence(self, result: dict) -> float:
        """
        计算整体匹配置信度。

        综合考量：
          - 已匹配表的平均得分
          - 未匹配表的比例（惩罚项）
          - 是否所有关键表都已匹配

        Args:
            result: match() 的输出（部分填充）

        Returns:
            float: 0.0-1.0 的整体置信度
        """
        mappings = result.get("mappings", [])
        unmatched_tables = result.get("unmatched_tables", [])
        total_tables = len(mappings) + len(unmatched_tables)

        if total_tables == 0:
            return 1.0  # 没有需要匹配的表

        # 已匹配表的平均分
        if mappings:
            avg_score = sum(m.get("confidence", 0.0) for m in mappings) / len(mappings)
        else:
            avg_score = 0.0

        # 未匹配惩罚
        unmatch_ratio = len(unmatched_tables) / total_tables
        unmatch_penalty = 1.0 - unmatch_ratio * 0.5  # 最多罚 0.5

        # 低分惩罚
        low_score_count = sum(1 for m in mappings if m.get("confidence", 0) < 0.6)
        low_score_penalty = 1.0 - (low_score_count / max(len(mappings), 1)) * 0.3

        final = avg_score * unmatch_penalty * low_score_penalty
        return max(0.0, min(1.0, final))

    def _determine_status(self, result: dict) -> str:
        """
        根据匹配结果确定整体状态。

        Returns:
            "ok" | "partial" | "low_confidence"
        """
        confidence = result.get("confidence", 0.0)
        warnings = result.get("warnings", [])
        unmatched = result.get("unmatched_tables", [])

        if confidence >= self.HIGH_CONFIDENCE_THRESHOLD and not unmatched:
            return "ok"

        if confidence >= self.MANUAL_REVIEW_THRESHOLD and len(unmatched) <= 3:
            return "partial"

        return "low_confidence"

    # -------------------------------------------------------
    # 批量/辅助方法
    # -------------------------------------------------------

    def match_multiple(
        self,
        analyses: List[dict],
        profiles: List[dict],
    ) -> List[dict]:
        """
        批量匹配多组模板与套表（适用于多个实体的场景）。

        Args:
            analyses: List[TemplateAnalyzer 输出]
            profiles: List[ExcelProfiler 输出]

        Returns:
            List[dict]: 每个实体对应的匹配结果
        """
        results = []
        n = min(len(analyses), len(profiles))
        for i in range(n):
            result = self.match(analyses[i], profiles[i])
            results.append(result)
        return results

    def to_json(self, result: dict, indent: int = 2, ensure_ascii: bool = False) -> str:
        """
        将匹配结果序列化为 JSON。

        Args:
            result: match() 的输出
            indent: JSON 缩进
            ensure_ascii: 是否转义非 ASCII

        Returns:
            str: JSON 字符串
        """
        return json.dumps(result, ensure_ascii=ensure_ascii, indent=indent, default=str)

    def print_summary(self, result: dict) -> None:
        """打印匹配结果摘要"""
        status = result.get("status", "?")
        confidence = result.get("confidence", 0.0)
        mappings = result.get("mappings", [])
        unmatched_tables = result.get("unmatched_tables", [])
        unmatched_sheets = result.get("unmatched_sheets", [])
        warnings = result.get("warnings", [])

        print("=" * 60)
        print(f"  智能匹配结果 — 状态: {status}  置信度: {confidence:.2%}")
        print("=" * 60)
        print(f"  已匹配: {len(mappings)} 张表")
        for m in mappings:
            ok_mark = "+"
            print(f"    {ok_mark} [{m['table_idx']}] {m['cat']}")
            print(f"       sheet: {m['sheet_kw']}")
            print(f"       得分: {m['confidence']:.2%}")
            print(f"       列映射: {m['col_map']}")
            print(f"       行模式: {m.get('row_match_mode', 'exact')}")

        if unmatched_tables:
            print(f"\n  未匹配模板表: {len(unmatched_tables)} 张")
            for ut in unmatched_tables:
                fail_mark = "x"
                print(
                    f"    {fail_mark} [{ut['index']}] {ut.get('name', '(无名)')} - {ut.get('reason', '')}"
                )

        if unmatched_sheets:
            print(f"\n  未匹配 Excel Sheet: {len(unmatched_sheets)} 个")
            for us in unmatched_sheets:
                print(f"    ? [{us['index']}] {us['name']} ({us.get('category', '?')})")

        if warnings:
            print(f"\n  警告 ({len(warnings)}):")
            for w in warnings:
                warn_mark = "!"
                print(f"    {warn_mark} {w}")

        print("=" * 60)


# ============================================================
# 快捷入口
# ============================================================


def quick_match(template_path: str, excel_path: str) -> dict:
    """
    快捷匹配：一行调用完成模板→套表的智能匹配。

    Args:
        template_path: Word 模板路径
        excel_path: Excel 套表路径

    Returns:
        dict: 匹配结果
    """
    from template_analyzer import TemplateAnalyzer
    from excel_profiler import ExcelProfiler

    # 分析模板
    analyzer = TemplateAnalyzer()
    template_analysis = analyzer.analyze(template_path)

    # 剖析套表
    profiler = ExcelProfiler()
    excel_profile = profiler.profile(excel_path)

    # 智能匹配
    matcher = SmartMatcher()
    return matcher.match(template_analysis, excel_profile)


# ============================================================
# 独立执行
# ============================================================


def main():
    """命令行入口：智能匹配并输出结果"""
    import argparse

    parser = argparse.ArgumentParser(
        description="智能匹配引擎 — 模板表格 ↔ Excel 套表自动配对",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python smart_matcher.py 模板.docx 套表.xlsx
  python smart_matcher.py 模板.docx 套表.xlsx --json
  python smart_matcher.py 模板.docx 套表.xlsx --summary-only
        """,
    )
    parser.add_argument("template", help="Word 审计附注模板路径 (.docx)")
    parser.add_argument("excel", help="Excel 决算套表路径 (.xlsx)")
    parser.add_argument("--json", "-j", action="store_true", help="以 JSON 格式输出")
    parser.add_argument("--summary-only", "-s", action="store_true", help="仅输出摘要")
    parser.add_argument(
        "--indent", "-i", type=int, default=2, help="JSON 缩进（默认 2）"
    )
    parser.add_argument(
        "--output", "-o", type=str, default=None, help="输出到文件（代替 stdout）"
    )

    args = parser.parse_args()

    try:
        result = quick_match(args.template, args.excel)
    except FileNotFoundError as e:
        print(f"错误: 文件不存在 — {e}", file=sys.stderr)
        sys.exit(1)
    except ValueError as e:
        print(f"错误: 文件格式不支持 — {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"错误: 匹配过程异常 — {e}", file=sys.stderr)
        sys.exit(1)

    if args.json:
        output = json.dumps(result, ensure_ascii=False, indent=args.indent, default=str)
        if args.output:
            with open(args.output, "w", encoding="utf-8") as f:
                f.write(output)
            print(f"结果已保存: {args.output}")
        else:
            print(output)
    elif args.summary_only:
        matcher = SmartMatcher()
        matcher.print_summary(result)
    else:
        matcher = SmartMatcher()
        matcher.print_summary(result)
        print("\n完整匹配结果:")
        print(json.dumps(result, ensure_ascii=False, indent=args.indent, default=str))


if __name__ == "__main__":
    main()
