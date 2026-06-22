"""
自动检测并运行所有匹配的 Data Adapter，合并结果。

一个 Excel 可能同时包含：
  - 明细表（_原始数据 sheet）→ TaozhangAdapter
  - 科目余额表 → GLAdapter
  - 三大主表 → SimpleAdapter

各适配器各自 accept()，各自提取能处理的数据，最后 merge。

用法:
    from data_adapter.auto_detect import extract_all
    data = extract_all("决算套表.xlsx")
"""

from .base import UnifiedData
from .taozhang_adapter import TaozhangAdapter
from .gl_adapter import GLAdapter
from .simple_adapter import SimpleAdapter


ADAPTERS = [TaozhangAdapter(), GLAdapter(), SimpleAdapter()]


def extract_all(filepath: str) -> UnifiedData:
    """
    遍历所有适配器，对 accept() 返回 True 的逐个执行 extract()，
    最终合并为一个 UnifiedData。
    """
    combined = UnifiedData()
    combined.sources.append(filepath)
    used = []

    for adapter in ADAPTERS:
        try:
            if adapter.accept(filepath):
                print(f"  运行适配器: {adapter.name}")
                partial = adapter.extract(filepath)
                combined.merge(partial)
                used.append(adapter.name)
        except Exception as e:
            msg = f"{adapter.name}.extract() 异常: {e}"
            print(f"  WARN: {msg}")
            combined.warnings.append(msg)

    if not used:
        print("  无适配器匹配，返回空数据")

    combined.adapters = used
    return combined
