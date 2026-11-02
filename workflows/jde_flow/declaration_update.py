"""更新房产税（从租）申报底稿：所属期、从租计征明细与税款所属期。"""
import calendar
from copy import copy
from datetime import date

from config.config_loader import ConfigError
from log.loggingConfig import logger

DECLARATION_SHEET = "房产税（从租）申报底稿"
CALC_SHEET = "房产税（从租）计算底稿"
TAX_OFFICE = "国家税务总局广州市税务局第三税务分局税源管理二科"
PRESERVED_COLUMNS = ("B", "C", "G", "H", "I")
DETAIL_LAST_COLUMN = 12


def _text(value):
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip() if value is not None else ""


def _find_cell(sheet, title, max_row=40, max_column=6):
    """按标题文字定位单元格（忽略大小写和首尾空格）。"""
    target = title.strip().lower()
    for row in range(1, min(sheet.max_row, max_row) + 1):
        for column in range(1, min(sheet.max_column, max_column) + 1):
            if _text(sheet.cell(row, column).value).lower() == target:
                return sheet.cell(row, column)
    return None


def _total_row(sheet, column, start_row):
    for row in range(start_row, sheet.max_row + 1):
        if _text(sheet.cell(row, column).value) == "合计":
            return row
    raise ConfigError(f"{sheet.title} 未找到“合计”行")


def _declaration_bu(value):
    """计算底稿 BU 转申报底稿 BU：6 位取前 2 位 + 后 2 位（250104 -> 2504）。"""
    text = _text(value)
    if text.isdigit() and len(text) == 6:
        return text[:2] + text[-2:]
    return text


def _quarter(period):
    year, month = period.year, period.month
    first = (month - 1) // 3 * 3 + 1
    return year, first, first + 2


def _calc_buses(calc):
    header = _find_cell(calc, "BU")
    if header is None:
        raise ConfigError(f"{calc.title} 未找到 BU 表头")
    try:
        total = _total_row(calc, header.column, header.row + 1)
    except ConfigError:
        total = calc.max_row + 1
    buses = []
    for row in range(header.row + 1, total):
        value = _text(calc.cell(row, header.column).value)
        if not value or value == "合计":
            continue
        bu = _declaration_bu(value)
        if bu and bu not in buses:
            buses.append(bu)
    if not buses:
        raise ConfigError("计算底稿没有可申报的 BU 数据")
    return buses


def _existing_rows(sheet, start, total):
    """记录已有 BU 对应的房产信息，用于本次保留。"""
    rows = {}
    for row in range(start, total):
        bu = _text(sheet.cell(row, 1).value)
        if not bu:
            continue
        rows.setdefault(bu, {col: sheet[f"{col}{row}"].value for col in PRESERVED_COLUMNS})
    return rows


def _copy_row_style(sheet, source, target):
    for column in range(1, DETAIL_LAST_COLUMN + 1):
        sheet.cell(target, column)._style = copy(sheet.cell(source, column)._style)
    height = sheet.row_dimensions[source].height
    if height is not None:
        sheet.row_dimensions[target].height = height


def _resize_rows(sheet, at_row, count):
    """在 at_row 前插入/删除行，并同步移动下方合并区域。"""
    if count == 0:
        return
    shifted = []
    for area in list(sheet.merged_cells.ranges):
        if area.min_row >= at_row:
            shifted.append(
                (area.min_col, area.min_row + count, area.max_col, area.max_row + count)
            )
            sheet.unmerge_cells(str(area))
    if count > 0:
        sheet.insert_rows(at_row, count)
    else:
        sheet.delete_rows(at_row, -count)
    for min_col, min_row, max_col, max_row in shifted:
        if min_row >= 1 and max_row >= min_row:
            sheet.merge_cells(
                start_row=min_row, start_column=min_col,
                end_row=max_row, end_column=max_col,
            )


def update_declaration(book, period):
    """按计算底稿 BU 重建申报明细，并更新所属期、有效期和税款所属期。

    返回供 Excel 保存使用的结构变更说明；模板缺少申报页时返回 None。
    """
    if DECLARATION_SHEET not in book.sheetnames:
        logger.warning(f"模板缺少工作表 {DECLARATION_SHEET}，跳过申报底稿更新")
        return None
    if CALC_SHEET not in book.sheetnames:
        raise ConfigError(f"模板缺少工作表：{CALC_SHEET}")

    buses = _calc_buses(book[CALC_SHEET])
    sheet = book[DECLARATION_SHEET]
    header = _find_cell(sheet, "BU")
    if header is None:
        raise ConfigError(f"{sheet.title} 未找到 BU 表头")
    old_total = _total_row(sheet, header.column, header.row + 1)
    start = header.row + 1
    old_count = old_total - start
    existing = _existing_rows(sheet, start, old_total)
    new_count = len(buses)

    insert_at = insert_count = delete_at = delete_count = 0
    if new_count > old_count:
        insert_at, insert_count = old_total, new_count - old_count
        _resize_rows(sheet, insert_at, insert_count)
    elif new_count < old_count:
        delete_at, delete_count = start + new_count, old_count - new_count
        _resize_rows(sheet, delete_at, -delete_count)

    total = start + new_count
    year, first, last = _quarter(period)
    quarter_start = date(year, first, 1)
    quarter_end = date(year, last, calendar.monthrange(year, last)[1])

    for offset in range(old_count, new_count):
        _copy_row_style(sheet, start, start + offset)

    writes = []
    for offset, bu in enumerate(buses):
        row = start + offset
        preserved = existing.get(bu, {})
        bus_value = int(bu) if bu.isdigit() else bu
        sheet.cell(row, 1).value = bus_value
        writes.append((f"A{row}", bus_value))
        for column in PRESERVED_COLUMNS:
            value = preserved.get(column)
            sheet[f"{column}{row}"].value = value
            writes.append((f"{column}{row}", value))
        for column, value in ((4, TAX_OFFICE), (5, date(year, 1, 1)),
                              (6, date(year, 12, 31)),
                              (11, f"=$K${total}"), (12, f"=$L${total}")):
            sheet.cell(row, column).value = value
            writes.append((f"{sheet.cell(row, column).coordinate}", value))
        for column in (5, 6, 11, 12):
            sheet.cell(row, column).number_format = "yyyy/m/d"

    sheet.cell(total, 1).value = "合计"
    sheet.cell(total, 11).value = quarter_start
    sheet.cell(total, 12).value = quarter_end
    for column in (11, 12):
        sheet.cell(total, column).number_format = "yyyy/m/d"
    writes.extend([
        (f"A{total}", "合计"),
        (f"K{total}", quarter_start),
        (f"L{total}", quarter_end),
    ])

    subject_header = _find_cell(sheet, "所属期：")
    if subject_header is None:
        raise ConfigError(f"{sheet.title} 未找到“所属期：”单元格")
    subject_cell = sheet.cell(subject_header.row, subject_header.column + 1)
    subject = f"{year}年{first}-{last}月"
    subject_cell.value = subject
    writes.append((subject_cell.coordinate, subject))

    logger.info(
        f"申报底稿已更新：所属期 {subject}，BU {buses}，"
        f"明细行 {old_count} -> {new_count}"
    )
    return {
        "sheet": DECLARATION_SHEET,
        "insert_at": insert_at,
        "insert_count": insert_count,
        "delete_at": delete_at,
        "delete_count": delete_count,
        "style_row": start,
        "writes": writes,
        "format_ranges": [f"E{start}:F{total}", f"K{start}:L{total}"],
    }
