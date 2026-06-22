"""
统一数据模型 + 适配器基类
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional
from abc import ABC, abstractmethod


@dataclass
class UnifiedData:
    """
    统一中间数据，完全与数据源格式解耦。

    accounts:     {科目名: {"期末": 数值, "期初": 数值}}
    balance_sheet:   {科目名: {"期末": 数值, "期初": 数值}}
    income_statement: {科目名: {"本期": 数值, "上期": 数值}}
    cash_flow:       {科目名: {"期末": 数值, "期初": 数值}}
    entity:          {字段名: 值}
    """

    accounts: Dict[str, Dict[str, float]] = field(default_factory=dict)
    balance_sheet: Dict[str, Dict[str, float]] = field(default_factory=dict)
    income_statement: Dict[str, Dict[str, float]] = field(default_factory=dict)
    cash_flow: Dict[str, Dict[str, float]] = field(default_factory=dict)
    entity: Dict[str, str] = field(default_factory=dict)

    sources: List[str] = field(default_factory=list)  # 数据来源文件
    adapters: List[str] = field(default_factory=list)  # 用到的适配器
    warnings: List[str] = field(default_factory=list)

    def get_account(self, name: str, period: str = "期末") -> Optional[float]:
        if name in self.accounts and period in self.accounts[name]:
            return self.accounts[name][period]
        return None

    def get_bs(self, name: str, period: str = "期末") -> Optional[float]:
        if name in self.balance_sheet and period in self.balance_sheet[name]:
            return self.balance_sheet[name][period]
        return None

    def get_is(self, name: str, period: str = "本期") -> Optional[float]:
        if name in self.income_statement and period in self.income_statement[name]:
            return self.income_statement[name][period]
        return None

    def merge(self, other: "UnifiedData"):
        """
        合并另一个 UnifiedData 到当前对象。
        accounts 相加互不覆盖，主表选条目数多的保留。
        """
        # accounts: 同名科目不覆盖（先到先得）
        for k, v in other.accounts.items():
            if k not in self.accounts:
                self.accounts[k] = dict(v)

        # 主表：选条目数更多的版本
        for field_name in ["balance_sheet", "income_statement", "cash_flow"]:
            self_dict = getattr(self, field_name)
            other_dict = getattr(other, field_name)
            if len(other_dict) > len(self_dict):
                setattr(self, field_name, dict(other_dict))
            elif not self_dict:
                setattr(self, field_name, dict(other_dict))

        # entity: 追加缺失项
        for k, v in other.entity.items():
            if k not in self.entity:
                self.entity[k] = v

        self.adapters.extend(other.adapters)
        self.sources.extend(other.sources)
        self.warnings.extend(other.warnings)

    def print_summary(self):
        print(f"  accounts: {len(self.accounts)} 项")
        print(f"  balance_sheet: {len(self.balance_sheet)} 项")
        print(f"  income_statement: {len(self.income_statement)} 项")
        print(f"  cash_flow: {len(self.cash_flow)} 项")
        print(f"  entity: {len(self.entity)} 项")
        if self.adapters:
            print(f"  用到的适配器: {', '.join(set(self.adapters))}")
        if self.warnings:
            for w in self.warnings:
                print(f"  WARN: {w}")


class BaseAdapter(ABC):
    """适配器基类"""

    @property
    @abstractmethod
    def name(self) -> str:
        pass

    @abstractmethod
    def accept(self, filepath: str) -> bool:
        """是否包含本适配器能处理的 sheet"""
        pass

    @abstractmethod
    def extract(self, filepath: str) -> UnifiedData:
        pass

    def safe_float(self, v) -> Optional[float]:
        if v is None:
            return None
        if isinstance(v, (int, float)):
            return float(v)
        try:
            return float(str(v).replace(",", "").replace(" ", ""))
        except (ValueError, TypeError):
            return None
