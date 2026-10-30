"""真实环境验证：登录 JDE，导出数据、更新模板并核对金额。

本脚本不参与 unittest 自动收集，手动执行。
可不带任何参数运行，全部取 Config.xlsx（默认当前月）：
    python tests/manual_login_export.py

也可按需覆盖单项参数：
    python tests/manual_login_export.py --from-company 02501 --to-company 02504 --period 2026-10
    python tests/manual_login_export.py --template "C:\\code\\tgh-jde-download\\data\\房产税计税表.xlsx" --company "太古汇" --accounting-period 2026-10
    python tests/manual_login_export.py --from-company 02501 --to-company 02504 \
        --start 2026-10-01 --end 2026-10-31

未传 --start/--end 时按自然月计算：
    传 --period 2026-10 → 该期间的自然月；
    未传 --period     → 会计期间（Config.xlsx）的自然月，未配置会计期间则取今天所在月份。
--period 只写入本次上下文 jde_ledger_period，不修改配置文件的 会计期间。
--accounting-period 用于输出模板名称及模板业务期间，与导出日期期间分开；未传时按自然月。

脚本登录后执行完整验证：
    明细账：选择页面 → 所有页面 → 总账页签 → 按科目帐的帐户查询
           → 填写条件 → 查找 → 输出网格 → 下载
    R55GL1005：返回主页 → 选择页面 → 所有页面 → 总账页签
               → 中方科目总账 → 提交报表 → 下载 → 启用模板宏 → 转换 xlsx
    模板：复制模板 → 按标题写入明细账 → 复制 R55GL1005 工作表
          → 更新会计期间与底稿内容 → compare_values 重算并比对金额
          （新口径：通过返回 True，不通过抛 ConfigError）
定位失败时去掉 --headless 重新运行，便于观察菜单与对话框状态。
"""
import argparse
from pathlib import Path
import sys

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from config import Config
from core.context import TaskContext
from log.loggingConfig import logger
from utils.helpers import get_timestamp
from web.browser import BrowserManager
from workflows.jde_flow.login_task import JDELoginTask
from workflows.jde_flow.property_tax import PropertyTaxFlow


def _as_bool(value):
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in ("true", "1", "yes")


def _report_shape(path):
    """读取转换后 R55GL1005 的工作表与主表尺寸，用于人工确认宏已生效。"""
    book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = book[book.sheetnames[0]]
        return list(book.sheetnames), sheet.max_row, sheet.max_column
    finally:
        book.close()


def main():
    parser = argparse.ArgumentParser(
        description="登录 JDE，导出明细账和 R55GL1005，更新模板并比对金额（真实环境）"
    )
    parser.add_argument("--credential-name", help="Windows 通用凭据名；默认读取 Config.xlsx")
    parser.add_argument("--ledger-type", help="分类账类型，例如 AA")
    parser.add_argument("--from-company", help="自公司，例如 02501")
    parser.add_argument("--to-company", help="至公司，例如 02504")
    parser.add_argument("--period", help="导出日期期间，支持 2026-10 / 2026/10 / 202610 / 2026年10月；"
                                        "未传 --start/--end 时用于计算自然月，缺省取会计期间或当前月")
    parser.add_argument("--start", help="起始日期，可选，例如 2026-10-01")
    parser.add_argument("--end", help="至日期，可选，例如 2026-10-31")
    parser.add_argument("--file", help="下载文件名、相对输出目录路径或绝对路径；默认按时间戳生成")
    parser.add_argument("--r55gl1005-file", help="R55GL1005 输出文件名或路径")
    parser.add_argument("--template", help="房产税模板文件名或路径；默认读取 Config.xlsx")
    parser.add_argument("--company", help="模板输出名称使用的公司；默认读取 Config.xlsx")
    parser.add_argument("--accounting-period", help="模板业务会计期间；未传时按自然月（当月）")
    parser.add_argument("--headless", action="store_true", help="无界面运行")
    args = parser.parse_args()

    config = Config()
    config.load(str(ROOT / "data/UserConfig.xlsx"), str(ROOT / "data/Config.xlsx"))
    config.ensure_directories()

    # 命令行参数只覆盖本次运行，不写回配置工作簿。
    overrides = {
        "JDE配置.jde_credential_name": args.credential_name,
        "房产税配置.明细账分类账类型": args.ledger_type,
        "房产税配置.明细账自公司": args.from_company,
        "房产税配置.明细账至公司": args.to_company,
        "房产税配置.明细账文件": args.file,
        "房产税配置.R55GL1005文件": args.r55gl1005_file,
        "房产税配置.模板文件": args.template,
        "房产税配置.公司": args.company,
        "房产税配置.会计期间": args.accounting_period,
    }
    for key, value in overrides.items():
        if value:
            config.set(key, value)

    # 未指定明细账文件名时按时间戳生成（支持绝对路径，写入指定目录）。
    if not config.get("房产税配置.明细账文件"):
        config.set("房产税配置.明细账文件", f"JDE明细账 {get_timestamp()}.xlsx")

    if not config.get("JDE配置.jde_credential_name"):
        raise SystemExit("缺少 JDE配置.jde_credential_name：请填写 Config.xlsx 或传入 --credential-name")
    for key in ("明细账自公司", "明细账至公司"):
        if not str(config.get(f"房产税配置.{key}") or "").strip():
            # 允许无传值运行；确实缺失时由 export_ledger 给出“缺少导出参数”的明确错误。
            logger.warning(f"未提供 房产税配置.{key}（命令行与 Config.xlsx 均为空）")
    if bool(args.start) != bool(args.end):
        logger.warning("仅提供起始或至日期时，另一端将取期间（--period 未传时为当前月）的月初/月末")

    context = TaskContext()
    context.config = config
    # 日期规则：start/end 都在上下文时直接使用；否则用 jde_ledger_period，缺省取当前月。
    if args.period:
        context.set("jde_ledger_period", args.period)
    if args.start:
        context.set("jde_ledger_start", args.start)
    if args.end:
        context.set("jde_ledger_end", args.end)
    browser = BrowserManager(
        headless=args.headless or _as_bool(config.get("JDE配置.headless", False)))
    context.set_browser("JDE", browser)
    if not any((args.credential_name, args.ledger_type, args.from_company, args.to_company,
                args.period, args.start, args.end, args.file, args.headless)):
        logger.info("未传入任何参数：全部使用 Config.xlsx 配置，期间取当前月")
    logger.info(f"本次参数：凭据 {config.get('JDE配置.jde_credential_name')}，"
                f"分类账类型 {config.get('房产税配置.明细账分类账类型', 'AA')}，"
                f"公司 {config.get('房产税配置.明细账自公司')}-{config.get('房产税配置.明细账至公司')}，"
                f"期间 {args.period or '当前月'}，输出 {config.get('房产税配置.明细账文件')}")
    try:
        browser.start()
        if not JDELoginTask(context).run():
            raise SystemExit("JDE 登录失败，未执行导出")
        logger.info("JDE 登录成功，开始导出明细账")
        logger.info("界面步骤：选择页面 → 所有页面 → 总账页签 → "
                    "按科目帐的帐户查询 → 填写条件 → 查找 → 输出网格 → 下载")
        if args.headless or _as_bool(config.get("JDE配置.headless", False)):
            logger.warning("当前为 headless 模式；菜单/对话框定位失败时请去掉 --headless 复现")
        flow = PropertyTaxFlow(context)
        try:
            start_date, end_date = flow.ledger_dates()
        except Exception as exc:
            raise SystemExit(f"无法确定导出期间：{exc}")
        period_first = flow.accounting_period()
        logger.info(f"会计期间：{period_first:%Y-%m}（{period_first:%Y%m}）；"
                    f"未配置 房产税配置.会计期间 时按自然月")
        if args.start and args.end:
            date_source = "显式 --start/--end"
        elif args.period:
            date_source = f"--period {args.period}"
        elif config.get("房产税配置.会计期间"):
            date_source = f"房产税配置.会计期间 {config.get('房产税配置.会计期间')}"
        else:
            date_source = "当前月"
        logger.info(f"查询条件：分类账类型 {config.get('房产税配置.明细账分类账类型', 'AA')}，"
                    f"公司 {config.get('房产税配置.明细账自公司')}-{config.get('房产税配置.明细账至公司')}，"
                    f"日期 {start_date} 至 {end_date}（来源：{date_source}）")
        path = flow.export_ledger()
        if not Path(path).is_file():
            raise SystemExit(f"导出文件不存在：{path}")
        logger.info(f"明细账导出成功：{path}")
        logger.info("明细账已完成，开始导出 R55GL1005")
        report_path = flow.export_r55gl1005()
        if not Path(report_path).is_file():
            raise SystemExit(f"R55GL1005 导出文件不存在：{report_path}")
        raw_report = Path(report_path).with_suffix(".xls")
        if not raw_report.is_file():
            raise SystemExit(f"R55GL1005 原始 .xls 不存在：{raw_report}")
        sheets, rows, columns = _report_shape(report_path)
        logger.info(
            f"R55GL1005 导出成功：{report_path}；"
            f"原始文件 {raw_report.name} 已启用模板宏并转换，"
            f"工作表={sheets}，主表 {rows} 行 × {columns} 列"
        )
        if rows <= 1 or columns <= 1:
            raise SystemExit(
                f"R55GL1005 转换结果异常（{rows} 行 × {columns} 列），请检查模板宏是否成功运行"
            )
        print(f"明细账：{path}")
        print(f"文件大小：{Path(path).stat().st_size} 字节")
        print(f"R55GL1005 原始：{raw_report}")
        print(f"文件大小：{raw_report.stat().st_size} 字节")
        print(f"R55GL1005 转换后：{report_path}")
        print(f"文件大小：{Path(report_path).stat().st_size} 字节")
        print(f"工作表：{sheets}，主表 {rows} 行 × {columns} 列")
        logger.info("两个 JDE 导出文件已就绪，开始生成并更新房产税模板")
        template_path = flow.create_template(path)
        if not Path(template_path).is_file():
            raise SystemExit(f"模板生成失败，输出文件不存在：{template_path}")
        logger.info(
            f"模板明细账写入完成：{template_path}；"
            f"明细行数={context.get('jde_ledger_rows')}，"
            f"标题映射={context.get('jde_ledger_mapping')}"
        )

        flow.update_template(template_path)
        logger.info(f"会计期间与底稿内容已更新：{template_path}")
        logger.info(
            f"重算前模板文件检查：存在={Path(template_path).is_file()}，"
            f"大小={Path(template_path).stat().st_size if Path(template_path).is_file() else 0} 字节"
        )

        # 重算走 compare_values 新口径：内部完成重算，通过返回 True，不通过抛 ConfigError。
        if flow.compare_values(template_path) is not True:
            raise SystemExit("模板金额比对未返回通过")

        check_status = context.get("property_tax_check_status")
        if check_status != "passed":
            raise SystemExit(f"模板金额比对未通过或未完成，状态：{check_status}")
        results = context.get("property_tax_check_results") or []
        check_report = context.get("property_tax_check_report")
        logger.info(f"模板金额比对通过：{len(results)} 项，结果报告：{check_report}")
        print(f"更新后模板：{template_path}")
        print(f"会计期间：{period_first:%Y-%m}")
        print(f"金额比对状态：{check_status}")
        print(f"金额比对报告：{check_report}")
    finally:
        context.close_all_browsers()


if __name__ == "__main__":
    main()
