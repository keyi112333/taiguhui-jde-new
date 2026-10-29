"""新版 Check 条件驱动的计税及公式更新。"""
from collections import defaultdict
from copy import copy
from decimal import Decimal, InvalidOperation

from config.config_loader import ConfigError


def text(value):
    return str(value).strip() if value is not None else ""


def update_from_check(book, include_children=False):
    sheet = book['房产税（从租）计算底稿']
    ledger = book['JDE明细账']
    cols = {text(c.value): c.column for c in ledger[1]}#固定了取模板表第二行的信息
    conditions = [(r, text(sheet.cell(r, 15).value), text(sheet.cell(r, 13).value))
                  for r in range(5, sheet.max_row + 1) if text(sheet.cell(r, 15).value)]
    if not conditions:
        raise ConfigError('计算底稿 Check 区没有账号筛选条件，请使用新版模板')
    if len({(a, d) for _, a, d in conditions}) != len(conditions):
        raise ConfigError('Check 筛选条件重复')
    totals = defaultdict(Decimal)
    refs = defaultdict(list)
    records = []
    end = max(ledger.max_row, 2)
    ranges = {key: f"JDE明细账!${ledger.cell(1, cols[key]).column_letter}$2:${ledger.cell(1, cols[key]).column_letter}${end}"
              for key in ('帐号', '单据 类型', '分类帐类型 1 金额')}
    for r, account, doc in conditions:
        if doc == '非':
            raise ConfigError(f'Check 第 {r} 行排除类型为空')
        amount = Decimal(0)
        for values in ledger.iter_rows(min_row=2, values_only=True):
            actual = text(values[cols['帐号'] - 1])
            if actual != account and not (include_children and actual.startswith(account + '.')):
                continue
            dtype = text(values[cols['单据 类型'] - 1])
            if doc and not (dtype != doc[1:] if doc.startswith('非') else dtype == doc):
                continue
            raw = values[cols['分类帐类型 1 金额'] - 1]
            try:
                value = Decimal(str(raw))
                if not value.is_finite():
                    raise InvalidOperation
            except (ValueError, InvalidOperation):
                raise ConfigError(f'科目 {actual} 的金额不是有效数值：{raw}')
            amount += value
        key = (account.split('.')[0], doc)
        totals[key] += amount
        refs[key].append(r)
        ar = f"TRIM({ranges['帐号']})"
        match = f'({ar}=TRIM(O{r}))'
        if include_children:
            match = f'(({ar}=TRIM(O{r}))+(LEFT({ar},LEN(TRIM(O{r}))+1)=TRIM(O{r})&"."))'
        dt = ''
        if doc:
            criterion = f'MID(M{r},2,99)' if doc.startswith('非') else f'M{r}'
            dt = f'*({ranges["单据 类型"]}{"<>" if doc.startswith("非") else "="}{criterion})'
        # SUMPRODUCT 不会把布尔数组当 1/0，乘 1 转成数值后再求和。
        sheet.cell(r, 16).value = (
            f'=SUMPRODUCT({match}{dt}*1,{ranges["分类帐类型 1 金额"]})')
        records.append({'BU': key[0], '收入科目': account, '单据类型': doc, '汇总金额': str(amount)})
    sheet.cell(max(r for r, _, _ in conditions) + 1, 16).value = '=' + '+'.join(f'P{r}' for r, _, _ in conditions)
    old_total = [r for r in range(9, sheet.max_row + 1) if text(sheet.cell(r, 2).value) == '合计']
    if len(old_total) != 1:
        raise ConfigError('计算底稿必须包含唯一合计行')
    old_total = old_total[0]
    total_styles = [copy(sheet.cell(old_total, c)._style) for c in (2, 3, 4)]
    old_rows = defaultdict(list)
    for r in range(9, old_total):
        if text(sheet.cell(r, 2).value):
            old_rows[text(sheet.cell(r, 2).value)].append(r)
    styles = [copy(sheet.cell(9, c)._style) for c in (2, 3, 4)]
    total = 9 + len(totals) + 2
    for row in range(9, max(total, old_total) + 1):
        for col in (2, 3, 4):
            sheet.cell(row, col).value = None
    new_rows = defaultdict(list)
    for row, ((bu, doc), amount) in enumerate(totals.items(), 9):
        sheet.cell(row, 2).value = int(bu) if bu.isdigit() else bu
        sheet.cell(row, 3).value = '=-SUM(' + ','.join(f'P{r}' for r in refs[bu, doc]) + ')'
        new_rows[bu].append(row)
    fixed = 9 + len(totals)
    sheet.cell(fixed, 2).value = 250108
    sheet.cell(fixed, 3).value = '=LOOKUP($C$4,MOH!$E$6:$AF$6,MOH!$E$7:$AF$7)'
    new_rows['250108'].append(fixed)
    for r in range(9, fixed + 1):
        for col, style in zip((2, 3, 4), styles):
            sheet.cell(r, col)._style = copy(style)
        sheet.cell(r, 4).value = f'=ROUND(C{r}*$C$8,2)'
    sheet.cell(total, 2).value = '合计'
    for c, style in zip((2, 3, 4), total_styles):
        sheet.cell(total, c)._style = copy(style)
    for c in (3, 4):
        letter = 'C' if c == 3 else 'D'
        sheet.cell(total, c).value = f'=SUM({letter}9:{letter}{total-1})'
    relocation = {old_total: total}
    for bu, rows in old_rows.items():
        for i, old in enumerate(rows):
            relocation[old] = new_rows[bu][i] if i < len(new_rows[bu]) else None
    import re
    for r in range(8, sheet.max_row + 1):
        account = text(sheet.cell(r, 7).value)
        if account.endswith('.917013'):
            sheet.cell(r, 8).value = '=' + '+'.join(f'$D${n}' for n in new_rows[account.split('.')[0]]) if new_rows[account.split('.')[0]] else '=0'
        cell = sheet.cell(r, 9)
        if cell.data_type == 'f':
            def move(m):
                old = int(m[2])
                return m[0] if old not in relocation else (m[1] + str(relocation[old]) if relocation[old] else '0')
            cell.value = re.sub(r'(\$?D\$?)(\d+)\b', move, cell.value)
    if 'R55GL1005' in book.sheetnames:
        from workflows.jde_flow.audit_update import update_audit
        update_audit(book, total)
    return records
