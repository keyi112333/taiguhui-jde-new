"""房产税模块二至六；未确认的业务规则明确阻断后续处理。"""
import calendar
import json
import re
from copy import copy
from datetime import date,datetime
from pathlib import Path
from tempfile import NamedTemporaryFile

import openpyxl
from openpyxl.workbook.properties import CalcProperties

from config.config_loader import ConfigError
from log.loggingConfig import logger
from utils.helpers import get_timestamp
from web.JDE.jde_function import download_gl_movement, download_r55gl1005
from workflows.jde_flow.template_update import (
    external_link_parts,
    restore_external_links,
    update_workbook,
)
from workflows.jde_flow.template_check import (
    check_sheet,
    convert_legacy_excel_to_xlsx,
    recalculate_file,
)


HEADER_DEFINITION = Path(__file__).resolve().parents[2] / "config/property_tax_ledger_headers.json"


def normalize_header_name(value):
    if value is None:
        return ""
    text = str(value).strip().lower()
    return re.sub(r"[\s_\-+/()（）]+", "", text)


def build_header_pairs(source_titles, target_titles, explicit_mapping=None):
    """按标题名构建源->模板映射；明确映射优先，随后按规范化标题自动补齐。"""
    source_by_norm = {normalize_header_name(title): title for title in source_titles}
    target_by_norm = {normalize_header_name(title): title for title in target_titles}
    pairs = {}
    mapping = {} if explicit_mapping is None else dict(explicit_mapping)

    def resolve_source(title):
        if title in source_by_norm:
            return source_by_norm[title]
        if title in source_titles:
            return title
        norm = normalize_header_name(title)
        return source_by_norm.get(norm)

    def resolve_target(title):
        if title in target_by_norm:
            return target_by_norm[title]
        if title in target_titles:
            return title
        norm = normalize_header_name(title)
        return target_by_norm.get(norm)

    for source_title, target_title in mapping.items():
        resolved_source = resolve_source(source_title)
        resolved_target = resolve_target(target_title)
        if resolved_source and resolved_target:
            pairs[resolved_source] = resolved_target

    for source_title in source_titles:
        if source_title in pairs:
            continue
        target_title = target_by_norm.get(normalize_header_name(source_title))
        if target_title and target_title not in pairs.values():
            pairs[source_title] = target_title

    for target_title in target_titles:
        if target_title in pairs.values():
            continue
        source_title = source_by_norm.get(normalize_header_name(target_title))
        if source_title and source_title not in pairs:
            pairs[source_title] = target_title

    return pairs


class PendingModule(Exception):
    """流程等待录制、数据或业务规则，不代表已完成。"""

def format_jde_date(value,label):

    if isinstance(value, (datetime,date)):
        return value.strftime("%m/%d/%Y")
    text = str(value).strip()
    for pattern in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%m-%d-%Y"):
        try:
            return datetime.strptime(text, pattern).strftime("%m/%d/%Y")
        except ValueError:
            continue
    raise ConfigError(f"{label} 日期格式不正确，应为 YYYY-MM-DD 或 MM/DD/YYYY 等可识别格式；当前值：{text}")

def month_bounds(period):
    """会计期间自然月第一天和最后一天"""
    if isinstance(period, (datetime,date)):
        year, month = period.year, period.month
    else:
        text = "" if period is None else str(period).strip()
        matchd = re.fullmatch(r"(\d{4})[-/年]?(\d{1,2})月?", text)
        if not matchd:
            raise ConfigError(f"会计期间格式不正确，应为 YYYY-MM 或 YYYY/MM 或 YYYYMM 等可识别格式；当前值：{text}")
        year, month = int(matchd.group(1)), int(matchd.group(2))
    if not (1 <= month <= 12):
        raise ConfigError(f"会计期间月份不正确，应为 1-12；当前值：{month}")
    return date(year, month, 1), date(year, month, calendar.monthrange(year, month)[1])



XL_PASTE_FORMATS = -4122


class ExcelUnavailable(Exception):
    """本机无法使用 Excel COM，用于回退到 openpyxl 写入。"""


def write_template_with_excel(template, destination, engine, sheet_name,
                              header_row, column_pairs, rows):
    """复制模板后由 Excel 写入明细并重算保存，模板中的图表、数据验证等原样保留。"""
    try:
        import pythoncom
        import shutil
        import win32com.client
    except ImportError as exc:
        raise ExcelUnavailable(f"缺少 pywin32：{exc}") from exc

    first_col = min(col for col, _ in column_pairs)
    last_col = max(col for col, _ in column_pairs)
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template, destination)
    pythoncom.CoInitialize()
    app = book = None
    saved = False
    try:
        try:
            app = win32com.client.DispatchEx(engine)
        except Exception as exc:
            raise ExcelUnavailable(f"无法启动 {engine}：{exc}") from exc
        app.Visible = False
        app.DisplayAlerts = False
        app.AutomationSecurity = 3
        try:
            book = app.Workbooks.Open(
                Filename=str(destination),
                UpdateLinks=0,
                ReadOnly=False,
                IgnoreReadOnlyRecommended=True,
                AddToMru=False,
            )
        except Exception as exc:
            raise ExcelUnavailable(f"Excel 无法打开输出文件：{exc}") from exc
        try:
            sheet = book.Worksheets(sheet_name)
        except Exception as exc:
            raise ConfigError(f"模板缺少工作表：{sheet_name}") from exc
        used = sheet.UsedRange
        last_row = max(used.Row + used.Rows.Count - 1, header_row)
        # 只清映射列的历史数据，保持模板其他列（公式等）不动。
        for target_col, _ in column_pairs:
            sheet.Range(
                sheet.Cells(header_row + 1, target_col),
                sheet.Cells(last_row, target_col),
            ).ClearContents()
        # 逐列写入：纯文本列先设为文本格式，避免 "00100" 被 Excel 转成 100。
        for target_col, source_col in column_pairs:
            target_range = sheet.Range(
                sheet.Cells(header_row + 1, target_col),
                sheet.Cells(header_row + len(rows), target_col),
            )
            column = [row[source_col - 1] for row in rows]
            text_column = all(value is None or isinstance(value, str) for value in column)
            original_format = target_range.NumberFormat if text_column else None
            if text_column:
                target_range.NumberFormat = "@"
            target_range.Value = [[value if value is not None else ""] for value in column]
            if text_column and original_format is not None:
                target_range.NumberFormat = original_format
        existing = last_row - header_row
        if len(rows) > existing:
            first = sheet.Range(sheet.Cells(header_row + 1, first_col),
                                sheet.Cells(header_row + 1, last_col))
            extra = sheet.Range(sheet.Cells(header_row + existing + 1, first_col),
                                sheet.Cells(header_row + len(rows), last_col))
            first.Copy()
            extra.PasteSpecial(XL_PASTE_FORMATS)
            app.CutCopyMode = False
            extra.EntireRow.RowHeight = first.EntireRow.RowHeight
        app.CalculateFullRebuild()
        book.Save()
        saved = True
    finally:
        try:
            if book is not None:
                book.Close(SaveChanges=False)
        finally:
            try:
                if app is not None:
                    app.Quit()
            finally:
                pythoncom.CoUninitialize()
                if not saved:
                    try:
                        destination.unlink()
                    except OSError:
                        pass
    if not destination.is_file() or destination.stat().st_size == 0:
        raise ConfigError(f"生成的输出文件无效：{destination}")


def write_template_with_openpyxl(template, destination, sheet_name,
                                 header_row, column_pairs, rows):
    """无 Excel 时的降级写入：读取模板后另存，保留其他工作表与单元格样式。"""
    links = external_link_parts(template)
    target = None
    temp_path = None
    try:
        target = openpyxl.load_workbook(template)
        sheet = target[sheet_name]
        columns = sorted({col for col, _ in column_pairs})
        # 保存首条明细格式，扩展新行时延续日期、金额等显示格式。
        styles = {col: copy(sheet.cell(header_row + 1, col)._style)
                  for col in columns}
        for col in columns:
            for row in range(header_row + 1, sheet.max_row + 1):
                sheet.cell(row, col).value = None
        for offset, row in enumerate(rows, 1):
            for col, source_col in column_pairs:
                cell = sheet.cell(header_row + offset, col)
                cell._style = copy(styles[col])
                value = row[source_col - 1]
                cell.value = value
                if isinstance(value, str):
                    cell.data_type = "s"
        target.calculation = CalcProperties(fullCalcOnLoad=True, forceFullCalc=True)
        destination.parent.mkdir(parents=True, exist_ok=True)
        with NamedTemporaryFile(dir=destination.parent, suffix=".xlsx", delete=False) as tmp:
            temp_path = Path(tmp.name)
        target.save(temp_path)
        restore_external_links(temp_path, links)
        temp_path.replace(destination)
        temp_path = None
    finally:
        if target is not None:
            target.close()
        if temp_path is not None and temp_path.exists():
            temp_path.unlink()


class PropertyTaxFlow:
    def __init__(self, context):
        self.context = context
        self.config = context.config

    def get(self, key, default=None):
        return self.config.get(f"房产税配置.{key}", default)

    def ledger_dates(self):
        """返回明细账起始和结束日期，优先使用配置项；未配置时按会计期间自然月计算。"""
        start = self.context.get("jde_ledger_start")
        end = self.context.get("jde_ledger_end")
        if start and end:
            return(format_jde_date(start, "明细账起始日期"), format_jde_date(end, "明细账至日期"))
        # first, last = month_bounds(self.context.get("jde_ledger_period") or self.get("会计期间"))
        first,last = month_bounds(self.context.get("jde_ledger_period") or date.today())
        return (format_jde_date(first, "明细账起始日期") if start else format_jde_date(first, "会计期间起始日期"),
                format_jde_date(last, "明细账至日期") if end else format_jde_date(last, "会计期间结束日期"))

    def export_ledger(self):
        """模块二：从 JDE 导出明细账，返回下载文件路径。"""
        required = ("明细账自公司", "明细账至公司")
        missing =[name for name in required 
                  if self.get(name) is None or not str(self.get(name)).strip()]
        if missing:
            raise ConfigError(f"模块二缺少导出参数：{', '.join(missing)}")
        browser = self.context.get_browser("JDE")
        page = browser.get_page() if browser else None
        if page is None:
            raise RuntimeError("JDE 浏览器未打开；请先登录 JDE")
        start_date, end_date = self.ledger_dates()

        destination = self.get("明细账文件")
        if destination:
            destination = Path(destination)
            if not destination.is_absolute():
                destination = self.config.get_path("output") / destination
            else:
                destination =self.config.get_path("output") / f"房产税明细账 {get_timestamp()}"
            path = download_gl_movement(
                page,
                ledger_type=self.get("明细账分类账类型", "AA"),
                from_company=self.get("明细账自公司"),
                to_company=self.get("明细账至公司"),
                start_date=start_date,
                end_date=end_date,
                destination=destination,
            )
            self.context.set("jde_ledger_path", path)
        return path

    def export_r55gl1005(self):
        """明细账导出后运行 R55GL1005 并返回本期报表路径。"""
        browser = self.context.get_browser("JDE")
        page = browser.get_page() if browser else None
        if page is None:
            raise RuntimeError("JDE 浏览器未打开；无法导出 R55GL1005")

        company = self.get("明细账自公司")
        if company is None or not str(company).strip():
            raise ConfigError("导出 R55GL1005 缺少 房产税配置.明细账自公司")

        destination = (
            self.context.get("jde_r55gl1005_path")
            or self.get("R55GL1005文件")
        )
        if destination:
            destination = Path(destination)
            if not destination.is_absolute():
                destination = self.config.get_path("input") / destination
        else:
            destination = self.config.get_path("output") / (
                f"R55GL1005_{get_timestamp()}.xlsx"
            )

        destination = Path(destination)
        if destination.suffix.lower() != ".xlsx":
            destination = destination.with_suffix(".xlsx")
        raw_path = download_r55gl1005(
            page,
            company=company,
            destination=destination.with_suffix(".xls"),
        )
        convert_legacy_excel_to_xlsx(raw_path, destination)
        logger.info(f"R55GL1005 已转换为工作簿格式：{destination}")
        path = destination
        self.context.set("jde_r55gl1005_path", path)
        return path

    def create_template(self, ledger_path):
        """模块三：复制模板生成输出文件，按标题写入明细并重算。"""
        input_dir = self.config.get_path("input")
        template_name = self.get("模板文件")
        if template_name:
            template = Path(template_name)
            if not template.is_absolute():
                template = input_dir / template
        else:
            candidates = [p for p in input_dir.glob("*房产税计税表*.xlsx")
                          if not p.name.startswith("~$") and p.resolve() != Path(ledger_path).resolve()]
            if len(candidates) != 1:
                raise ConfigError("输入目录需有唯一包含‘房产税计税表’的 xlsx 模板，或配置 模板文件")
            template = candidates[0]
        company, period = self.get("公司"), self.get("会计期间")
        for label, value in (("公司", company), ("会计期间", period)):
            if value is None or not str(value).strip() or re.search(r'[<>:"/\\|?*\x00-\x1f]', str(value)):
                raise ConfigError(f"房产税配置.{label} 必填且不能包含文件名非法字符")
        output_dir = Path(self.get("输出目录") or self.config.get_path("output"))
        if not output_dir.is_absolute():
            output_dir = self.config.get_path("output") / output_dir
        destination = output_dir / f"房产税计税表 {str(company).strip()} {str(period).strip()}.xlsx"
        if destination.exists() or destination.resolve() in (template.resolve(), Path(ledger_path).resolve()):
            raise ConfigError(f"输出文件已存在或与输入重合：{destination}")
        mapping = self.get("标题映射", "")
        try:
            mapping = json.loads(mapping) if mapping else {}
        except (TypeError, ValueError) as exc:
            raise ConfigError("标题映射须为 JSON 对象：源标题 -> 模板标题") from exc
        if not isinstance(mapping, dict) or not all(isinstance(k, str) and isinstance(v, str) for k, v in mapping.items()):
            raise ConfigError("标题映射须为字符串到字符串的 JSON 对象")
        definition = json.loads(HEADER_DEFINITION.read_text(encoding="utf-8"))
        expected = [column["template_header"] for column in definition["columns"]]
        source = openpyxl.load_workbook(ledger_path, data_only=True)
        target = None
        try:
            sheet_name = self.get("明细账工作表") or source.sheetnames[0]
            source_sheet = source[sheet_name]
            target = openpyxl.load_workbook(template)
            target_sheet = target[definition["sheet"]]
            source_row = int(self.get("明细账标题行", 1))
            target_row = int(self.get("模板标题行", definition["header_row"]))
            if min(source_row, target_row) < 1:
                raise ConfigError("标题行必须为正整数")
            def headers(sheet, row):
                result = {}
                for cell in sheet[row]:
                    if cell.value is None or not str(cell.value).strip():
                        continue
                    title = str(cell.value).strip()
                    if title in result:
                        raise ConfigError(f"{sheet.title} 标题重复：{title}")
                    result[title] = cell.column
                return result
            src = headers(source_sheet, source_row)
            dst = headers(target_sheet, target_row)
            missing_expected = sorted(set(expected) - set(dst))
            if missing_expected:
                raise ConfigError(f"模板标题与记录不符；缺少：{missing_expected}")
            aliases = {column["source_header"]: column["template_header"]
                       for column in definition["columns"] if column["source_header"]}
            aliases.update(mapping)
            pairs = build_header_pairs(list(src), list(dst), aliases)
            if not pairs:
                raise ConfigError("没有匹配的标题，请补充标题映射")
            if len(set(pairs.values())) != len(pairs):
                raise ConfigError("多个源标题不能写入同一个模板标题")
            for origin, name in pairs.items():
                if origin not in src or name not in dst:
                    raise ConfigError(f"标题不存在：{origin} -> {name}")
            unmatched_source = sorted(set(src) - set(pairs))
            if unmatched_source:
                raise ConfigError(f"源标题尚未映射：{unmatched_source}")
            missing_template = sorted(set(expected) - set(pairs.values()))
            if missing_template:
                raise ConfigError(f"明细账缺少模板所需列：{missing_template}")
            rows = [row for row in source_sheet.iter_rows(min_row=source_row + 1, values_only=True)
                    if not all(value is None or (isinstance(value, str) and not value.strip())
                               for value in row)]
            if not rows:
                raise ConfigError("明细账没有数据，未生成输出")
            target.close()
            target = None
            column_pairs = [(dst[name], src[origin]) for origin, name in pairs.items()]
            engine = str(self.get("重算引擎", "Excel.Application")).strip()
            if engine.lower() == "openpyxl":
                write_template_with_openpyxl(template, destination, definition["sheet"],
                                             target_row, column_pairs, rows)
            else:
                try:
                    write_template_with_excel(template, destination, engine,
                                              definition["sheet"], target_row, column_pairs, rows)
                    logger.info(f"已复制模板并由 {engine} 写入明细：{destination}")
                except ExcelUnavailable as exc:
                    logger.warning(f"Excel 不可用，改用 openpyxl 写入模板：{exc}")
                    write_template_with_openpyxl(template, destination, definition["sheet"],
                                                 target_row, column_pairs, rows)
            self.context.set("property_tax_path", destination)
            self.context.set("jde_ledger_rows", len(rows))
            self.context.set("jde_ledger_mapping", pairs)
            return destination
        finally:
            source.close()
            if target is not None:
                target.close()

    def update_template(self, path):
        """模块四：更新 BU 汇总；公式由 Excel 打开时重算。"""
        mode = self.get("收入科目匹配方式", "精确")
        if mode not in ("精确", "包含子科目"):
            raise ConfigError("收入科目匹配方式须为 精确 或 包含子科目")
        book = openpyxl.load_workbook(path)
        links = external_link_parts(path)
        temp_path = None
        try:
            report_path = self.context.get("jde_r55gl1005_path") or self.get("R55GL1005文件")
            if not report_path:
                raise PendingModule("R55GL1005 导出待录制；请配置 房产税配置.R55GL1005文件，避免使用模板历史报表")
            report_path = Path(report_path)
            if not report_path.is_absolute():
                report_path = self.config.get_path("input") / report_path
            report = openpyxl.load_workbook(report_path, data_only=False)
            try:
                source_name = self.get("R55GL1005工作表") or report.sheetnames[0]
                if source_name not in report.sheetnames:
                    raise ConfigError(f"R55GL1005 工作表不存在：{source_name}")
                source_sheet = report[source_name]
                if "R55GL1005" in book.sheetnames:
                    book.remove(book["R55GL1005"])
                destination = book.create_sheet("R55GL1005")
                for merged in source_sheet.merged_cells.ranges:
                    destination.merge_cells(str(merged))
                for row in range(1, source_sheet.max_row + 1):
                    for col in range(1, source_sheet.max_column + 1):
                        source_cell = source_sheet.cell(row=row, column=col)
                        target_cell = destination.cell(row=row, column=col)
                        target_cell.value = source_cell.value
                        target_cell._style = copy(source_cell._style)
                        target_cell.number_format = source_cell.number_format
                        target_cell.alignment = copy(source_cell.alignment)
                        target_cell.protection = copy(source_cell.protection)
                        if source_cell.comment:
                            target_cell.comment = source_cell.comment
            finally:
                report.close()
            records = update_workbook(book, include_children=mode == "包含子科目")
            book.calculation = CalcProperties(fullCalcOnLoad=True, forceFullCalc=True)
            with NamedTemporaryFile(dir=Path(path).parent, suffix=".xlsx", delete=False) as tmp:
                temp_path = Path(tmp.name)
            book.save(temp_path)
            restore_external_links(temp_path, links)
            temp_path.replace(path)
            self.context.set("property_tax_bu_totals", records)
        finally:
            book.close()
            if temp_path and temp_path.exists():
                temp_path.unlink()

    def compare_values(self, path):
        """模块五：重算并核对；是否写分录仍待业务规则确认。"""
        self.context.set("property_tax_check_status", "running")
        try:
            recalculate_file(path, self.get("重算引擎", "Excel.Application"))
            book = openpyxl.load_workbook(path, data_only=True)
            try:
                results = check_sheet(book["房产税（从租）计算底稿"],
                                      tolerance=self.get("核对容差", "0.00"),
                                      audit=book["房产税（从租）审核底稿"] if "R55GL1005" in book.sheetnames else None)
            finally:
                book.close()
            passed = all(item["通过"] for item in results)
            self.context.set("property_tax_check_results", results)
            self.context.set("property_tax_check_status", "passed" if passed else "failed")
            report = Path(path).with_suffix(".核对结果.json")
            report.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
            self.context.set("property_tax_check_report", report)
            if not passed:
                raise ConfigError("模板核对失败，详见核对结果：" + str(report))
        except Exception:
            self.context.set("property_tax_check_status", "failed")
            raise
        raise PendingModule("模板核对通过；是否进入 JDE 写入分录的规则待补充")

    def write_journal(self, path):
        """模块六：待补充分录字段及 JDE 操作。"""
        raise PendingModule("模块六待补充：JDE 分录写入操作")

    def run(self):
        self.context.set("property_tax_status", "running")
        try:
            ledger_path = self.export_ledger()
            self.export_r55gl1005()
            path = self.create_template(ledger_path)
            self.update_template(path)
            decision = self.compare_values(path)
            if type(decision) is not bool:
                raise ConfigError("模块五必须明确返回是否需要分录")
            self.context.set("journal_required", decision)
            if decision:
                self.write_journal(path)
            self.context.set("property_tax_status", "completed")
        except PendingModule as exc:
            self.context.set("property_tax_status", "pending")
            self.context.set("property_tax_pending_reason", str(exc))
            logger.warning(str(exc))
        except Exception:
            self.context.set("property_tax_status", "failed")
            raise
