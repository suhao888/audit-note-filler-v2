"""
智能适配入口 — 一键完成 模板分析 + Excel 分析 + 匹配 + 配置生成

用法:
    python auto_adapt.py --template 模板路径.docx --excel 套表路径.xlsx
    python auto_adapt.py --template 模板.docx --excel 套表.xlsx --output ./my_config
    python auto_adapt.py --template 模板.docx --excel 套表.xlsx --existing ./config
"""

import sys, argparse, json
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
sys.stdout.reconfigure(encoding="utf-8")

from template_analyzer import TemplateAnalyzer
from excel_profiler import ExcelProfiler
from smart_matcher import SmartMatcher
from config_generator import ConfigGenerator


def main():
    parser = argparse.ArgumentParser(
        description="审计附注智能适配 — 自动生成 MAPPINGS 配置"
    )
    parser.add_argument("--template", required=True, help="Word 附注模板路径 (.docx)")
    parser.add_argument("--excel", required=True, help="决算套表 Excel 路径 (.xlsx)")
    parser.add_argument(
        "--output", default="config_auto", help="输出配置目录 (默认: config_auto)"
    )
    parser.add_argument("--existing", help="已有 config/ 目录路径 (用于学习历史调优)")
    parser.add_argument("--json", action="store_true", help="输出 JSON 格式结果")
    parser.add_argument(
        "--min-confidence",
        type=float,
        default=0.5,
        help="最低置信度阈值 (低于此值标为需人工确认, 默认: 0.5)",
    )

    args = parser.parse_args()

    template_path = Path(args.template)
    excel_path = Path(args.excel)
    output_dir = Path(args.output)

    # 文件检查
    if not template_path.exists():
        print(f"ERROR: 模板文件不存在: {template_path}")
        sys.exit(1)
    if not excel_path.exists():
        print(f"ERROR: 套表文件不存在: {excel_path}")
        sys.exit(1)

    print("=" * 60)
    print("审计附注智能适配 v2")
    print("=" * 60)

    # Step 1: 分析模板
    print("\n[1/4] 分析 Word 模板...")
    ta = TemplateAnalyzer()
    template_result = ta.analyze(str(template_path))
    if template_result.get("error"):
        print(f"  错误: {template_result['error']}")
        sys.exit(1)
    print(f"  表格数量: {template_result['num_tables']}")
    table_names = [
        t.get("name", f"表{t['index']}") for t in template_result.get("tables", [])
    ]
    for i, name in enumerate(table_names[:10]):
        print(f"    表[{i}]: {name}")
    if len(table_names) > 10:
        print(f"    ... 还有 {len(table_names) - 10} 张表")

    # Step 2: 分析 Excel
    print("\n[2/4] 分析 Excel 套表...")
    ep = ExcelProfiler()
    excel_result = ep.profile(str(excel_path))
    if excel_result.get("error"):
        print(f"  错误: {excel_result['error']}")
        sys.exit(1)
    print(f"  Sheet 数量: {excel_result['num_sheets']}")
    notes_sheets = excel_result.get("notes_sheets", [])
    print(f"  附注类 Sheet: {len(notes_sheets)} 张")
    for s in notes_sheets[:10]:
        print(f"    {s}")
    if len(notes_sheets) > 10:
        print(f"    ... 还有 {len(notes_sheets) - 10} 张")

    # Step 3: 智能匹配
    print("\n[3/4] 执行智能匹配...")
    sm = SmartMatcher()
    match_result = sm.match(template_result, excel_result)

    status = match_result.get("status", "error")
    confidence = match_result.get("confidence", 0.0)
    print(f"  匹配状态: {status}")
    print(f"  全局置信度: {confidence * 100:.1f}%")
    print(f"  成功匹配: {len(match_result.get('mappings', []))} 张表")
    unmatched = match_result.get("unmatched_tables", [])
    if unmatched:
        print(f"  未匹配: {len(unmatched)} 张表")
        for t in unmatched[:5]:
            print(
                f"    - {t.get('name', '未知')} (置信度: {t.get('confidence', 0) * 100:.0f}%)"
            )
    warnings = match_result.get("warnings", [])
    if warnings:
        print(f"  警告: {len(warnings)} 条")
        for w in warnings[:5]:
            print(f"    - {w}")

    # Step 4: 生成配置
    print("\n[4/4] 生成配置...")
    cg = ConfigGenerator(output_dir=str(output_dir))

    # 如果有已有配置，学习历史调优
    if args.existing:
        existing_path = Path(args.existing)
        if existing_path.exists():
            print(f"  加载历史配置: {existing_path}")
            cg.learn_from_existing(str(existing_path))

    gen_result = cg.generate(match_result, template_result, excel_result)
    print(f"  配置目录: {gen_result.get('config_dir', output_dir)}")
    files = gen_result.get("files", {})
    for fname, fstatus in files.items():
        print(f"    {fname}: {fstatus}")

    issues = gen_result.get("issues", [])
    if issues:
        print(f"\n  需关注的问题 ({len(issues)} 项):")
        for issue in issues:
            print(f"    ⚠ {issue}")

    # 输出 JSON
    if args.json:
        output = {
            "template": str(template_path),
            "excel": str(excel_path),
            "match_status": status,
            "confidence": round(confidence, 4),
            "mappings_count": len(match_result.get("mappings", [])),
            "unmatched_tables": len(match_result.get("unmatched_tables", [])),
            "warnings_count": len(warnings),
            "issues_count": len(issues),
            "config_dir": gen_result.get("config_dir", str(output_dir)),
        }
        print("\n" + json.dumps(output, ensure_ascii=False, indent=2))

    print("\n" + "=" * 60)
    print("完成! 请检查生成结果并运行测试填充。")
    print("=" * 60)


if __name__ == "__main__":
    main()
