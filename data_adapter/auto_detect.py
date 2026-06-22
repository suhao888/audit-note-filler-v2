"""
自动检测并选择合适的 Data Adapter

用法:
    from data_adapter.auto_detect import select_adapter

    adapter = select_adapter("决算套表.xlsx")
    data = adapter.extract("决算套表.xlsx")

适配器优先级:
    1. TaozhangAdapter — 国资委标准决算套表
    2. GLAdapter — 科目余额表
    3. SimpleAdapter — 仅有三大主表
"""

from .base import BaseAdapter, UnifiedData
from .taozhang_adapter import TaozhangAdapter
from .gl_adapter import GLAdapter
from .simple_adapter import SimpleAdapter


ADAPTERS = [TaozhangAdapter(), GLAdapter(), SimpleAdapter()]


def select_adapter(filepath: str) -> BaseAdapter:
    """
    遍历所有适配器，返回第一个 accept() 返回 True 的。
    无匹配时返回 SimpleAdapter（兜底）。
    """
    for adapter in ADAPTERS:
        try:
            if adapter.accept(filepath):
                return adapter
        except Exception as e:
            print(f"  WARN: {adapter.name} accept() 失败: {e}")
            continue
    # 兜底
    return SimpleAdapter()


def extract_all(filepath: str) -> UnifiedData:
    """自动选适配器并提取数据"""
    adapter = select_adapter(filepath)
    print(f"  选用适配器: {adapter.name}")
    data = adapter.extract(filepath)
    return data
