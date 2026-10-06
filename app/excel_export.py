"""Build a right-to-left Excel report.

Persian is written as ordinary Unicode. Glyphs are not reshaped and strings are
not reversed: that treatment is for images and PDFs, and it makes Excel text
unreadable. Excel itself shapes the letters because the sheet is right-to-left
and each text cell uses reading order 2.
"""

from __future__ import annotations

import io
from decimal import Decimal

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from app.fa import format_jalali, quantize_money, to_persian_digits
from app.reconcile import (
    STATUS_OVER,
    STATUS_PAID_ONLY,
    STATUS_SETTLED,
    STATUS_SHORT,
    STATUS_UNPAID,
    Person,
    Report,
)

FONT_NAME = "Tahoma"
CHARSET_ARABIC = 178

INK = "17343A"
TEAL = "0F6E68"
GOLD = "A6843D"
PAPER = "F7F4EE"
LINE = "E6DDD0"
WHITE = "FFFFFF"
ZEBRA = "FBF8F4"
RED = "9C2F2F"
GREEN = "1F7A4D"
BLUE = "1D4E89"
AMBER = "8A5A12"

STATUS_STYLE = {
    STATUS_SETTLED: ("E5F4EC", GREEN),
    STATUS_SHORT: ("FDECEC", RED),
    STATUS_UNPAID: ("FDECEC", RED),
    STATUS_OVER: ("FFF4E5", AMBER),
    STATUS_PAID_ONLY: ("E8F1FA", BLUE),
}

THIN = Border(
    left=Side(style="thin", color=LINE),
    right=Side(style="thin", color=LINE),
    top=Side(style="thin", color=LINE),
    bottom=Side(style="thin", color=LINE),
)
TOTAL_BORDER = Border(
    left=Side(style="thin", color=LINE),
    right=Side(style="thin", color=LINE),
    top=Side(style="medium", color=GOLD),
    bottom=Side(style="thin", color=LINE),
)


def build_workbook(report: Report) -> bytes:
    workbook = Workbook()
    workbook.properties.title = "گزارش مغایرت‌گیری"
    workbook.properties.creator = "سامانه مغایرت‌گیری"
    workbook.properties.subject = "تطبیق مبلغ قابل‌پرداخت با پرداخت‌شده"
    workbook.calculation.calcMode = "auto"
    workbook.calculation.fullCalcOnLoad = True

    summary = workbook.active
    summary.title = "خلاصه"
    _write_summary(summary, report)

    detail_columns = _detail_columns()
    _write_table(
        workbook.create_sheet("تطبیق و مانده"),
        "تطبیق و مانده",
        "مانده یعنی مبلغی که هنوز باید پرداخت شود. اگر فقط بعضی ماه‌ها پرداخت شده باشد، در ستون مغایرت و ماه‌های مانده نوشته می‌شود.",
        detail_columns,
        _detail_rows(report.people),
        tab_color=TEAL,
    )
    _write_table(
        workbook.create_sheet("مانده ماه‌ها"),
        "مانده ماه‌ها",
        "هر ردیف یک ماه پرداخت‌نشده است. اگر فقط مثلاً شهریور پرداخت شده باشد، بقیه ماه‌ها اینجا می‌آیند.",
        _month_columns(),
        _month_rows(report),
        tab_color=RED,
    )
    _write_table(
        workbook.create_sheet("جزئیات منابع"),
        "جزئیات منابع",
        "سهم هر فایل برای هر بیمه‌گذار.",
        _audit_columns(),
        _audit_rows(report),
        tab_color="5E6A71",
    )

    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def _font(size: int = 11, bold: bool = False, color: str = INK) -> Font:
    return Font(name=FONT_NAME, size=size, bold=bold, color=color, charset=CHARSET_ARABIC)


def _align(horizontal: str = "right", wrap: bool = False) -> Alignment:
    return Alignment(horizontal=horizontal, vertical="center", wrap_text=wrap, readingOrder=2)


def _fill(color: str) -> PatternFill:
    return PatternFill("solid", fgColor=color)


def _prepare(worksheet: Worksheet, tab_color: str) -> None:
    worksheet.sheet_view.rightToLeft = True
    worksheet.sheet_view.showGridLines = False
    worksheet.sheet_view.zoomScale = 110
    worksheet.sheet_properties.tabColor = tab_color
    worksheet.page_setup.orientation = "landscape"
    worksheet.page_setup.paperSize = worksheet.PAPERSIZE_A4
    worksheet.page_setup.fitToWidth = 1
    worksheet.page_setup.fitToHeight = 0
    worksheet.page_setup.horizontalCentered = True
    worksheet.sheet_properties.pageSetUpPr.fitToPage = True
    worksheet.page_margins.left = 0.45
    worksheet.page_margins.right = 0.45
    worksheet.page_margins.top = 0.6
    worksheet.page_margins.bottom = 0.6
    worksheet.page_margins.header = 0.25
    worksheet.page_margins.footer = 0.3
    worksheet.oddFooter.right.text = "سامانه مغایرت‌گیری"
    worksheet.oddFooter.right.font = FONT_NAME
    worksheet.oddFooter.left.text = "صفحه &P از &N"
    worksheet.oddFooter.left.font = FONT_NAME
    worksheet.sheet_properties.pageSetUpPr.fitToPage = True
    worksheet.print_options.horizontalCentered = True
    worksheet.sheet_format.defaultRowHeight = 22


def _put(worksheet: Worksheet, row: int, column: int, value, kind: str, fill: PatternFill | None = None, border: Border = THIN):
    cell = worksheet.cell(row, column)
    if kind == "text":
        text = "" if value is None else str(value).strip()
        if text[:1] in "=-+@":
            text = "'" + text
        cell.value = text or None
        if text:
            cell.number_format = "@"
        cell.alignment = _align("right", wrap=column_is_note(value, kind))
    elif kind == "id":
        text = "" if value is None else str(value).strip()
        cell.value = text
        cell.number_format = "@"
        cell.alignment = _align("center")
    elif kind == "index":
        cell.value = value
        cell.number_format = "#,##0"
        cell.alignment = _align("center")
    elif kind == "money":
        if value is None:
            cell.value = None
        else:
            cell.value = _excel_number(value)
            cell.number_format = "#,##0" if _is_whole(value) else "#,##0.00"
        cell.alignment = _align("center")
    elif kind == "percent":
        if value is None:
            cell.value = None
        else:
            cell.value = float(value)
            cell.number_format = "0.0%"
        cell.alignment = _align("center")
    elif kind == "status":
        text = "" if value is None else str(value)
        cell.value = text
        cell.alignment = _align("center")
        colors = STATUS_STYLE.get(text)
        if colors:
            cell.fill = _fill(colors[0])
            cell.font = _font(bold=True, color=colors[1])
        else:
            cell.font = _font()
    elif kind == "formula":
        cell.value = value
        cell.number_format = "#,##0"
        cell.alignment = _align("center")
    else:
        cell.value = value
        cell.alignment = _align()

    if kind != "status":
        cell.font = _font(bold=kind in {"index"})
    if fill is not None and kind != "status":
        cell.fill = fill
    cell.border = border
    return cell


def column_is_note(value, kind: str) -> bool:
    return False


def _is_whole(value: Decimal) -> bool:
    return value == value.to_integral_value()


def _excel_number(value: Decimal):
    amount = quantize_money(Decimal(value))
    if _is_whole(amount):
        return int(amount)
    return float(amount)


def _paint_money_font(cell, value: Decimal | None) -> None:
    if value is None:
        return
    if value > 0:
        cell.font = _font(bold=True, color=RED)
    elif value < 0:
        cell.font = _font(bold=True, color=BLUE)
    else:
        cell.font = _font(bold=True, color=GREEN)


def _write_summary(worksheet: Worksheet, report: Report) -> None:
    _prepare(worksheet, GOLD)
    widths = [34, 24, 28, 22, 18, 16, 14, 18, 22]
    for index, width in enumerate(widths, start=1):
        worksheet.column_dimensions[get_column_letter(index)].width = width

    worksheet.merge_cells("A1:I1")
    title = worksheet["A1"]
    title.value = "گزارش مغایرت‌گیری پرداخت‌ها"
    title.font = _font(20, True, WHITE)
    title.fill = _fill(INK)
    title.alignment = _align("right")
    worksheet.row_dimensions[1].height = 42

    worksheet.merge_cells("A2:I2")
    stamp = worksheet["A2"]
    clock = report.created_at.strftime("%H:%M")
    stamp.value = (
        f"تاریخ تهیه: {format_jalali(report.created_at)} ساعت {to_persian_digits(clock)}"
        f"  |  حد چشم‌پوشی: {to_persian_digits(format(report.tolerance, 'f'))} ریال"
    )
    stamp.font = _font(11, False, INK)
    stamp.fill = _fill("F3E6C8")
    stamp.alignment = _align("right")
    worksheet.row_dimensions[2].height = 24

    counts = report.counts()
    due_total = quantize_money(sum((person.due for person in report.people), Decimal("0")))
    paid_total = quantize_money(sum((person.paid for person in report.people), Decimal("0")))
    balance = quantize_money(due_total - paid_total)
    metrics = [
        ("تعداد بیمه‌گذار", len(report.people), "index"),
        ("مبلغ کل", due_total, "money"),
        ("تطبیق پرداخت", paid_total, "money"),
        ("مانده", balance, "money"),
        ("تسویه‌شده", counts[STATUS_SETTLED], "index"),
        ("کسری پرداخت", counts[STATUS_SHORT], "index"),
        ("فاقد پرداخت", counts[STATUS_UNPAID], "index"),
        ("اضافه‌پرداخت", counts[STATUS_OVER], "index"),
        ("پرداخت بدون هزینه", counts[STATUS_PAID_ONLY], "index"),
        ("ردیف‌های کنار گذاشته‌شده", sum(source.skipped_rows for source in report.sources), "index"),
    ]

    worksheet.merge_cells("E4:I4")
    guide_title = worksheet["E4"]
    guide_title.value = "راهنمای خواندن گزارش"
    guide_title.font = _font(13, True, WHITE)
    guide_title.fill = _fill(TEAL)
    guide_title.alignment = _align("right")

    worksheet.merge_cells("E5:I14")
    guide = worksheet["E5"]
    guide.value = (
        "مبلغ کل، مبلغی است که باید پرداخت شود.\n"
        "تطبیق پرداخت، مبلغی است که واقعاً پرداخت شده.\n"
        "مانده = مبلغ کل − تطبیق پرداخت. مانده مثبت یعنی هنوز باید پرداخت شود.\n"
        "برگهٔ «مانده ماه‌ها» نشان می‌دهد هر کس کدام ماه‌ها را نپرداخته است."
    )
    guide.font = _font(11)
    guide.alignment = Alignment(horizontal="right", vertical="top", wrap_text=True, readingOrder=2)
    guide.fill = _fill(PAPER)
    for row in range(5, 15):
        for column in range(5, 10):
            worksheet.cell(row, column).border = THIN
            worksheet.cell(row, column).fill = _fill(PAPER)

    section = worksheet["A4"]
    section.value = "نمای کلی"
    section.font = _font(13, True, WHITE)
    section.fill = _fill(INK)
    section.alignment = _align("right")
    worksheet["B4"].fill = _fill(INK)
    worksheet["B4"].border = THIN
    section.border = THIN

    for offset, (label, value, kind) in enumerate(metrics):
        row = 5 + offset
        label_cell = worksheet.cell(row, 1, label)
        label_cell.font = _font(11, True)
        label_cell.alignment = _align("right")
        label_cell.fill = _fill(WHITE if offset % 2 == 0 else ZEBRA)
        label_cell.border = THIN
        value_cell = _put(worksheet, row, 2, value, kind, fill=_fill(WHITE if offset % 2 == 0 else ZEBRA))
        if label == "مانده":
            _paint_money_font(value_cell, balance)
        worksheet.row_dimensions[row].height = 22

    header_row = 16
    headers = [
        "عنوان فهرست",
        "نقش",
        "نام فایل",
        "برگه",
        "ردیف عنوان",
        "ستون مبلغ",
        "ضریب",
        "ردیف‌های محاسبه‌شده",
        "جمع مبلغ",
    ]
    worksheet.merge_cells("A15:I15")
    files_title = worksheet["A15"]
    files_title.value = "فایل‌هایی که در تطبیق آمده‌اند"
    files_title.font = _font(13, True, WHITE)
    files_title.fill = _fill(INK)
    files_title.alignment = _align("right")
    worksheet.row_dimensions[15].height = 26
    for column in range(1, 10):
        worksheet.cell(15, column).fill = _fill(INK)

    for column, header in enumerate(headers, start=1):
        cell = worksheet.cell(header_row, column, header)
        cell.font = _font(10, True, WHITE)
        cell.fill = _fill(TEAL)
        cell.alignment = _align("center", wrap=True)
        cell.border = THIN
    worksheet.row_dimensions[header_row].height = 30
    worksheet.auto_filter.ref = f"A{header_row}:I{header_row + max(len(report.sources), 1)}"

    for index, source in enumerate(report.sources):
        row = header_row + 1 + index
        role = "قابل‌پرداخت" if source.role == "due" else "پرداخت‌شده"
        values = [
            (source.label, "text"),
            (role, "text"),
            (source.filename, "text"),
            (source.sheet, "text"),
            (source.header_row, "index"),
            (source.amount_header, "text"),
            (_excel_number(source.factor), "index"),
            (source.used_rows, "index"),
            (source.total, "money"),
        ]
        fill = _fill(WHITE if index % 2 == 0 else ZEBRA)
        for column, (value, kind) in enumerate(values, start=1):
            _put(worksheet, row, column, value, kind, fill=fill)
        worksheet.row_dimensions[row].height = 22

    warning_row = header_row + len(report.sources) + 3
    worksheet.merge_cells(start_row=warning_row, start_column=1, end_row=warning_row, end_column=9)
    warning_title = worksheet.cell(warning_row, 1, "هشدارها و توضیح‌ها")
    warning_title.font = _font(13, True, WHITE)
    warning_title.fill = _fill("6B4E16")
    warning_title.alignment = _align("right")
    for column in range(1, 10):
        worksheet.cell(warning_row, column).fill = _fill("6B4E16")
    lines = report.warnings or ["هشداری ثبت نشده است. ردیف‌های دارای کد ملی یا نام، در محاسبه آمده‌اند."]
    for offset, line in enumerate(lines):
        row = warning_row + 1 + offset
        worksheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
        cell = worksheet.cell(row, 1, line)
        cell.font = _font(10)
        cell.alignment = _align("right", wrap=True)
        cell.fill = _fill("FFF9F0")
        for column in range(1, 10):
            worksheet.cell(row, column).border = THIN
            worksheet.cell(row, column).fill = _fill("FFF9F0")
        worksheet.row_dimensions[row].height = 22

    legend_row = warning_row + len(lines) + 2
    worksheet.cell(legend_row, 1, "معنی وضعیت‌ها").font = _font(12, True, WHITE)
    worksheet.cell(legend_row, 1).fill = _fill(INK)
    worksheet.cell(legend_row, 1).alignment = _align("right")
    legends = [
        (STATUS_SETTLED, "مبلغ پرداخت‌شده با مبلغ قابل‌پرداخت، در حد چشم‌پوشی، برابر است."),
        (STATUS_SHORT, "بخشی از مبلغ پرداخت شده و هنوز مانده دارد."),
        (STATUS_UNPAID, "در فهرست هزینه هست و هیچ پرداختی برایش نیامده است."),
        (STATUS_OVER, "بیشتر از مبلغ قابل‌پرداخت، پرداخت شده است."),
        (STATUS_PAID_ONLY, "پرداخت ثبت شده ولی در فهرست قابل‌پرداخت ردیفی ندارد."),
    ]
    for offset, (status, meaning) in enumerate(legends):
        row = legend_row + 1 + offset
        _put(worksheet, row, 1, status, "status")
        worksheet.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
        cell = worksheet.cell(row, 2, meaning)
        cell.font = _font(10)
        cell.alignment = _align("right")
        cell.border = THIN
        for column in range(2, 10):
            worksheet.cell(row, column).border = THIN
        worksheet.row_dimensions[row].height = 22

    worksheet.freeze_panes = "A17"
    worksheet.page_setup.fitToWidth = 1
    worksheet.oddHeader.right.text = "خلاصهٔ مغایرت"
    worksheet.print_title_rows = "1:2"
    worksheet.sheet_view.showGridLines = False


def _detail_columns() -> list[tuple[str, str, float]]:
    return [
        ("ردیف", "index", 8),
        ("نام بیمه‌گذار", "text", 24),
        ("شماره بیمه", "id", 18),
        ("کد ملی", "id", 16),
        ("نوع بیمه", "text", 16),
        ("مبلغ کل", "money", 16),
        ("تطبیق پرداخت", "money", 16),
        ("مانده", "money", 16),
        ("مغایرت", "text", 42),
        ("ماه‌های پرداخت‌شده", "text", 28),
        ("ماه‌های مانده", "text", 36),
    ]


def _detail_rows(people: list[Person]) -> list[list[tuple]]:
    rows = []
    for person in people:
        rows.append(
            [
                (None, "index"),
                (person.holder_name, "text"),
                (person.insurance_no, "id"),
                (person.national_id, "id"),
                (person.insurance_type, "text"),
                (person.due, "money"),
                (person.paid, "money"),
                (person.balance, "money"),
                (person.discrepancy_text(), "text"),
                ("، ".join(person.paid_month_list()), "text"),
                ("، ".join(person.unpaid_month_list()), "text"),
            ]
        )
    return rows


def _month_columns() -> list[tuple[str, str, float]]:
    return [
        ("ردیف", "index", 8),
        ("نام بیمه‌گذار", "text", 24),
        ("شماره بیمه", "id", 18),
        ("نوع بیمه", "text", 16),
        ("ماه", "text", 14),
        ("مبلغ ماه", "money", 16),
        ("پرداخت‌شده", "money", 16),
        ("مانده", "money", 16),
    ]


def _month_rows(report: Report) -> list[list[tuple]]:
    rows = []
    for person in report.people:
        for month in person.unpaid_month_list():
            due = person.due_months.get(month, Decimal("0"))
            paid = person.paid_months.get(month, Decimal("0"))
            rows.append(
                [
                    (None, "index"),
                    (person.holder_name, "text"),
                    (person.insurance_no, "id"),
                    (person.insurance_type, "text"),
                    (month, "text"),
                    (due, "money"),
                    (paid, "money"),
                    (quantize_money(due - paid), "money"),
                ]
            )
    return rows


def _audit_columns() -> list[tuple[str, str, float]]:
    return [
        ("ردیف", "index", 8),
        ("نام بیمه‌گذار", "text", 24),
        ("شماره بیمه", "id", 18),
        ("فهرست", "text", 28),
        ("نقش", "text", 16),
        ("مبلغ", "money", 16),
    ]


def _audit_rows(report: Report) -> list[list[tuple]]:
    rows = []
    for person in report.people:
        for source in report.sources:
            if source.key not in person.by_source:
                continue
            role = "قابل‌پرداخت" if source.role == "due" else "پرداخت‌شده"
            rows.append(
                [
                    (None, "index"),
                    (person.holder_name, "text"),
                    (person.insurance_no, "id"),
                    (source.label, "text"),
                    (role, "text"),
                    (person.by_source[source.key], "money"),
                ]
            )
    return rows


def _write_table(
    worksheet: Worksheet,
    title: str,
    note: str,
    columns: list[tuple[str, str, float]],
    rows: list[list[tuple]],
    tab_color: str,
) -> None:
    _prepare(worksheet, tab_color)
    last_column = len(columns)
    for index, (_, _, width) in enumerate(columns, start=1):
        worksheet.column_dimensions[get_column_letter(index)].width = width

    worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_column)
    title_cell = worksheet.cell(1, 1, title)
    title_cell.font = _font(18, True, WHITE)
    title_cell.fill = _fill(INK)
    title_cell.alignment = _align("right")
    worksheet.row_dimensions[1].height = 36
    for column in range(1, last_column + 1):
        worksheet.cell(1, column).fill = _fill(INK)

    worksheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_column)
    note_cell = worksheet.cell(2, 1, note)
    note_cell.font = _font(10, False, INK)
    note_cell.fill = _fill("F3E6C8")
    note_cell.alignment = _align("right", wrap=False)
    worksheet.row_dimensions[2].height = 22
    for column in range(1, last_column + 1):
        worksheet.cell(2, column).fill = _fill("F3E6C8")

    header_row = 3
    for column, (header, _, _) in enumerate(columns, start=1):
        cell = worksheet.cell(header_row, column, header)
        cell.font = _font(10, True, WHITE)
        cell.fill = _fill(TEAL)
        cell.alignment = _align("center", wrap=False)
        cell.border = THIN
    worksheet.row_dimensions[header_row].height = 22
    worksheet.freeze_panes = "A4"
    worksheet.auto_filter.ref = f"A{header_row}:{get_column_letter(last_column)}{header_row + max(len(rows), 1)}"
    worksheet.print_title_rows = "1:3"
    worksheet.page_setup.fitToWidth = 1
    worksheet.oddHeader.right.text = title

    if not rows:
        worksheet.merge_cells(start_row=4, start_column=1, end_row=4, end_column=last_column)
        empty = worksheet.cell(4, 1, "ردیفی برای نمایش وجود ندارد.")
        empty.font = _font(12, True, GREEN)
        empty.alignment = _align("right")
        empty.fill = _fill("E5F4EC")
        worksheet.row_dimensions[4].height = 22
        return

    balance_index = next((index for index, (header, _, _) in enumerate(columns, start=1) if header == "مانده"), None)
    money_indexes = [index for index, (_, kind, _) in enumerate(columns, start=1) if kind == "money"]
    percent_index = next((index for index, (_, kind, _) in enumerate(columns, start=1) if kind == "percent"), None)
    due_index = next((index for index, (header, _, _) in enumerate(columns, start=1) if header == "جمع قابل‌پرداخت"), None)
    paid_index = next((index for index, (header, _, _) in enumerate(columns, start=1) if header == "جمع پرداخت‌شده"), None)

    for row_offset, values in enumerate(rows):
        row = 4 + row_offset
        fill = _fill(WHITE if row_offset % 2 == 0 else ZEBRA)
        for column, (value, kind) in enumerate(values, start=1):
            written = value
            this_kind = kind
            if column == 1:
                written = row_offset + 1
                this_kind = "index"
            cell = _put(worksheet, row, column, written, this_kind, fill=fill)
            if balance_index == column and isinstance(value, Decimal):
                _paint_money_font(cell, value)
        worksheet.row_dimensions[row].height = 22

    total_row = 4 + len(rows)
    label = worksheet.cell(total_row, 1, "جمع")
    label.font = _font(11, True, INK)
    label.fill = _fill("F3E6C8")
    label.alignment = _align("center")
    label.border = TOTAL_BORDER
    for column in range(2, last_column + 1):
        cell = worksheet.cell(total_row, column, None)
        cell.fill = _fill("F3E6C8")
        cell.border = TOTAL_BORDER
        cell.font = _font(11, True)
        cell.alignment = _align("center")
        if column in money_indexes:
            letter = get_column_letter(column)
            cell.value = f"=SUBTOTAL(109,{letter}4:{letter}{total_row - 1})"
            cell.number_format = "#,##0"
        elif percent_index == column and due_index and paid_index:
            due_letter = get_column_letter(due_index)
            paid_letter = get_column_letter(paid_index)
            cell.value = f'=IF({due_letter}{total_row}=0,"",{paid_letter}{total_row}/{due_letter}{total_row})'
            cell.number_format = "0.0%"
    if balance_index:
        _paint_money_font(worksheet.cell(total_row, balance_index), None)
        worksheet.cell(total_row, balance_index).font = _font(11, True, INK)
    worksheet.row_dimensions[total_row].height = 22
