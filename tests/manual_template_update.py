"""本地测试模板复制和更新，不登录 JDE；结果保留供 Excel 人工检查。"""
from datetime import datetime
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import openpyxl

from core.context import TaskContext
from workflows.jde_flow.property_tax import PropertyTaxFlow


class LocalConfig:
    def __init__(self, values, output):
        self.values = values
        self.output = output

    def get(self, key, default=None):
        return self.values.get(key, default)

    def get_path(self, kind):
        return self.output if kind == "output" else ROOT / "data"


def main():
    # ================== 在这里修改你的文件路径 ==================
    # 建议使用绝对路径，或者相对于当前脚本的路径
    TEMPLATE_PATH = Path(r"C:\你的路径\房产税模板.xlsx")       # 原始计税表模板
    LEDGER_PATH = Path(r"C:\你的路径\JDE明细账.xlsx")         # 导出的 JDE 明细账
    REPORT_PATH = Path(r"C:\你的路径\R55GL1005.xlsx")         # 导出的 R55GL1005
    OUTPUT_DIR = Path(r"C:\你的路径\output\template_tests")   # 输出目录
    # ========================================================

    # 检查文件是否存在
    for path in (TEMPLATE_PATH, LEDGER_PATH, REPORT_PATH):
        if not path.is_file():
            raise FileNotFoundError(f"输入文件不存在：{path}")

    # 每次运行使用独立目录，业务入口也会拒绝覆盖输入文件。
    output = OUTPUT_DIR.resolve() / datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    
    values = {
        "模板文件": str(TEMPLATE_PATH.resolve()),
        "R55GL1005文件": str(REPORT_PATH.resolve()),
        "公司": "复制更新测试",          # 默认值
        "会计期间": "测试期间",           # 默认值
        "输出目录": str(output),
        "明细账标题行": 1,               # 默认值
        # "明细账工作表": "Sheet1",      # 如果需要指定工作表，取消注释并修改
        # "R55GL1005工作表": "Sheet1",   # 如果需要指定工作表，取消注释并修改
    }

    context = TaskContext()
    context.config = LocalConfig(
        {f"房产税配置.{key}": value for key, value in values.items()}, output
    )
    
    flow = PropertyTaxFlow(context)
    path = flow.create_template(LEDGER_PATH.resolve())
    print(f"复制完成：{path}")
    print(f"写入明细账：{context.get('jde_ledger_rows')} 行")
    
    flow.update_template(path)

    book = openpyxl.load_workbook(path, data_only=False)
    try:
        audit = book["房产税（从租）审核底稿"]
        rows = [r for r in range(8, audit.max_row + 1)
                if audit.cell(r, 6).value is not None]
        if not rows:
            raise AssertionError("审核底稿未生成科目行")
        for r in rows:
            actual = audit.cell(r, 9).value
            if actual != f"=H{r}-G{r}":
                raise AssertionError(f"I{r} 公式不完整：{actual!r}")
        if "蓝色" not in book.defined_names:
            raise AssertionError("缺少蓝色区域命名范围")
        print(f"更新完成：{len(context.get('property_tax_bu_totals'))} 条 Check 汇总")
        print(f"公式检查通过：审核底稿 {len(rows)} 行")
        print(f"蓝色范围：{book.defined_names['蓝色'].attr_text}")
        print(f"N14：{audit['N14'].value}")
    finally:
        book.close()
    print("请用 Excel 打开结果查看重算金额；本脚本仅检查公式，不执行金额核对。")


if __name__ == "__main__":
    main()