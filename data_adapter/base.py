"""
统一数据模型 + 适配器基类
所有数据源适配器继承 BaseAdapter，输出 UnifiedData
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional
from abc import ABC, abstractmethod


@dataclass
class UnifiedData:
    """
    统一中间数据，完全与数据源格式解耦。

    accounts: {科目名: {"期末": 数值, "期初": 数值}}
        或 {科目名: {"本期": 数值, "上期": 数值}}

    balance_sheet: {科目名: {"期末": 数值, "期初": 数值}}
    income_statement: {科目名: {"本期": 数值, "上期": 数值}}
    cash_flow: {科目名: {"期末": 数值, "期初": 数值}}

    entity: {字段名: 字符串值}
    """

    accounts: Dict[str, Dict[str, float]] = field(default_factory=dict)
    balance_sheet: Dict[str, Dict[str, float]] = field(default_factory=dict)
    income_statement: Dict[str, Dict[str, float]] = field(default_factory=dict)
    cash_flow: Dict[str, Dict[str, float]] = field(default_factory=dict)
    entity: Dict[str, str] = field(default_factory=dict)

    source_files: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    def get_account(self, name: str, period: str = "期末") -> Optional[float]:
        """从 accounts 取科目余额"""
        if name in self.accounts and period in self.accounts[name]:
            return self.accounts[name][period]
        return None

    def get_bs(self, name: str, period: str = "期末") -> Optional[float]:
        """从资产负债表取数"""
        if name in self.balance_sheet and period in self.balance_sheet[name]:
            return self.balance_sheet[name][period]
        return None

    def get_is(self, name: str, period: str = "本期") -> Optional[float]:
        """从利润表取数"""
        if name in self.income_statement and period in self.income_statement[name]:
            return self.income_statement[name][period]
        return None

    def get_cf(self, name: str, period: str = "期末") -> Optional[float]:
        """从现金流量表取数"""
        if name in self.cash_flow and period in self.cash_flow[name]:
            return self.cash_flow[name][period]
        return None

    def print_summary(self):
        """打印摘要"""
        print(f"  accounts: {len(self.accounts)} 项")
        print(f"  balance_sheet: {len(self.balance_sheet)} 项")
        print(f"  income_statement: {len(self.income_statement)} 项")
        print(f"  cash_flow: {len(self.cash_flow)} 项")
        if self.warnings:
            for w in self.warnings:
                print(f"  WARN: {w}")


class BaseAdapter(ABC):
    """适配器基类"""

    @property
    @abstractmethod
    def name(self) -> str:
        """适配器名称"""
        pass

    @abstractmethod
    def accept(self, filepath: str) -> bool:
        """
        判断是否能处理此文件。
        检查 sheet 名、文件结构等。
        """
        pass

    @abstractmethod
    def extract(self, filepath: str) -> UnifiedData:
        """
        提取为统一数据。
        """
        pass

    def safe_float(self, v) -> Optional[float]:
        """安全转浮点数"""
        if v is None:
            return None
        if isinstance(v, (int, float)):
            return float(v)
        try:
            return float(str(v).replace(",", "").replace(" ", ""))
        except (ValueError, TypeError):
            return None
