"""Match payable amounts with payments across any number of files."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal

from app.errors import UserError
from app.fa import (
    display_fa,
    fa_sort_key,
    is_blank,
    loose_name,
    national_id_is_valid,
    normalize_national_id,
    parse_amount,
    parse_decimal_input,
    preview_cell,
    quantize_money,
    safe_label,
    split_full_name,
    to_persian_digits,
)

STATUS_SETTLED = "تسویه‌شده"
STATUS_SHORT = "کسری پرداخت"
STATUS_OVER = "اضافه‌پرداخت"
STATUS_UNPAID = "فاقد پرداخت"
STATUS_PAID_ONLY = "پرداخت بدون هزینه"

ROLE_DUE = "due"
ROLE_PAID = "paid"


@dataclass
class ColumnMap:
    national_id: int | None = None
    first_name: int | None = None
    last_name: int | None = None
    full_name: int | None = None
    amount: int | None = None
    amount_due: int | None = None
    amount_paid: int | None = None
    note: int | None = None


@dataclass
class FileJob:
    file_id: str
    filename: str
    sheet: str
    header_row: int
    label: str
    role: str
    factor: Decimal
    columns: ColumnMap
    rows: list[list]


@dataclass
class SourceInfo:
    key: str
    title: str
    role: str
    label: str
    filename: str
    sheet: str
    header_row: int
    amount_header: str
    factor: Decimal
    used_rows: int = 0
    skipped_rows: int = 0
    total: Decimal = Decimal("0")


@dataclass
class Person:
    national_id: str = ""
    first_name: str = ""
    last_name: str = ""
    full_name: str = ""
    due: Decimal = Decimal("0")
    paid: Decimal = Decimal("0")
    by_source: dict[str, Decimal] = field(default_factory=dict)
    notes: list[str] = field(default_factory=list)
    matched_by_name: bool = False
    status: str = ""

    @property
    def balance(self) -> Decimal:
        return quantize_money(self.due - self.paid)

    @property
    def full_display(self) -> str:
        if self.full_name:
            return self.full_name
        return display_fa(f"{self.first_name} {self.last_name}")


@dataclass
class Report:
    people: list[Person]
    sources: list[SourceInfo]
    warnings: list[str]
    tolerance: Decimal
    created_at: datetime

    def counts(self) -> dict[str, int]:
        result = {
            STATUS_SETTLED: 0,
            STATUS_SHORT: 0,
            STATUS_OVER: 0,
            STATUS_UNPAID: 0,
            STATUS_PAID_ONLY: 0,
        }
        for person in self.people:
            result[person.status] = result.get(person.status, 0) + 1
        return result


def status_for(due: Decimal, paid: Decimal, tolerance: Decimal) -> str:
    balance = due - paid
    if due == 0 and paid > 0:
        if abs(paid) <= tolerance:
            return STATUS_SETTLED
        return STATUS_PAID_ONLY
    if abs(balance) <= tolerance:
        return STATUS_SETTLED
    if paid == 0 and due > 0:
        return STATUS_UNPAID
    if balance > 0:
        return STATUS_SHORT
    return STATUS_OVER


def reconcile(jobs: list[FileJob], tolerance: Decimal) -> Report:
    if tolerance < 0:
        raise UserError("حد چشم‌پوشی نمی‌تواند منفی باشد.")
    if not jobs:
        raise UserError("حداقل یک فایل بارگذاری کنید.")

    sources = _expand_sources(jobs)
    if not any(source.role == ROLE_DUE for source in sources):
        raise UserError(
            "هیچ فهرستی برای مبلغ قابل‌پرداخت مشخص نشده است. نقش یکی از فایل‌ها را «قابل‌پرداخت» بگذارید."
        )
    if not any(source.role == ROLE_PAID for source in sources):
        raise UserError(
            "هیچ فهرستی برای مبلغ پرداخت‌شده مشخص نشده است. نقش یکی از فایل‌ها را «پرداخت‌شده» بگذارید، یا اگر هر دو مبلغ در یک فایل است نقش «هر دو در یک فایل» را انتخاب کنید."
        )

    people: dict[str, Person] = {}
    warnings: list[str] = []
    for job in jobs:
        job_sources = [source for source in sources if source.key.startswith(f"{job.file_id}:")]
        _consume_job(job, job_sources, people, warnings)

    ready: list[Person] = []
    for person in people.values():
        person.due = quantize_money(
            sum((amount for key, amount in person.by_source.items() if _role_of(sources, key) == ROLE_DUE), Decimal("0"))
        )
        person.paid = quantize_money(
            sum((amount for key, amount in person.by_source.items() if _role_of(sources, key) == ROLE_PAID), Decimal("0"))
        )
        if person.due == 0 and person.paid == 0:
            continue
        person.status = status_for(person.due, person.paid, tolerance)
        if person.matched_by_name:
            _add_note(person, "این شخص با نام تطبیق شد، چون کد ملی در همهٔ فهرست‌ها نبود.")
        if person.status == STATUS_SETTLED and person.balance != 0:
            _add_note(person, "اختلاف مبلغ در حد چشم‌پوشی است.")
        person.notes = person.notes[:8]
        ready.append(person)

    ready.sort(key=lambda person: (fa_sort_key(person.last_name), fa_sort_key(person.first_name), person.national_id))
    return Report(
        people=ready,
        sources=sources,
        warnings=warnings[:40],
        tolerance=quantize_money(tolerance),
        created_at=datetime.now(),
    )


def _role_of(sources: list[SourceInfo], key: str) -> str:
    for source in sources:
        if source.key == key:
            return source.role
    return ROLE_DUE


def _expand_sources(jobs: list[FileJob]) -> list[SourceInfo]:
    used_titles: set[str] = set()
    sources: list[SourceInfo] = []
    for job in jobs:
        headers = _headers_of(job)
        _require_identity(job, headers)
        if job.factor <= 0:
            raise UserError(f"ضریب مبلغ در «{job.label}» باید بزرگ‌تر از صفر باشد.")
        if job.role == ROLE_DUE:
            sources.append(_make_source(job, ROLE_DUE, job.columns.amount, "قابل‌پرداخت", headers, used_titles))
        elif job.role == ROLE_PAID:
            sources.append(_make_source(job, ROLE_PAID, job.columns.amount, "پرداخت‌شده", headers, used_titles))
        elif job.role == "both":
            if job.columns.amount_due is None or job.columns.amount_paid is None:
                raise UserError(f"در «{job.label}» هر دو ستون مبلغ را مشخص کنید.")
            if job.columns.amount_due == job.columns.amount_paid:
                raise UserError(f"در «{job.label}» دو ستون مبلغ نباید یکی باشند.")
            sources.append(_make_source(job, ROLE_DUE, job.columns.amount_due, "قابل‌پرداخت", headers, used_titles))
            sources.append(_make_source(job, ROLE_PAID, job.columns.amount_paid, "پرداخت‌شده", headers, used_titles))
        else:
            raise UserError("نقش فایل نامعتبر است.")
    return sources


def _headers_of(job: FileJob) -> list[str]:
    index = job.header_row - 1
    if index < 0 or index >= len(job.rows):
        raise UserError(f"ردیف عنوان در «{job.label}» پیدا نشد.")
    width = len(job.rows[index])
    headers = []
    for cell in job.rows[index]:
        headers.append(display_fa(preview_cell(cell)) if not is_blank(cell) else "")
    if width == 0:
        raise UserError(f"برگهٔ «{job.label}» عنوان ستون ندارد.")
    return headers


def _require_identity(job: FileJob, headers: list[str]) -> None:
    columns = job.columns
    chosen = [columns.national_id, columns.first_name, columns.last_name, columns.full_name]
    if all(item is None for item in chosen):
        raise UserError(f"در «{job.label}» حداقل کد ملی یا نام را مشخص کنید.")
    for column in chosen:
        _check_index(column, headers, job.label)
    for column in (columns.amount, columns.amount_due, columns.amount_paid, columns.note):
        _check_index(column, headers, job.label)
    if job.role in {ROLE_DUE, ROLE_PAID} and columns.amount is None:
        raise UserError(f"در «{job.label}» ستون مبلغ را مشخص کنید.")


def _check_index(column: int | None, headers: list[str], label: str) -> None:
    if column is None:
        return
    if column < 0 or column >= len(headers):
        raise UserError(f"یکی از ستون‌های «{label}» خارج از محدوده است.")


def _make_source(
    job: FileJob,
    role: str,
    column: int | None,
    role_title: str,
    headers: list[str],
    used_titles: set[str],
) -> SourceInfo:
    if column is None:
        raise UserError(f"در «{job.label}» ستون مبلغ مشخص نشده است.")
    amount_header = headers[column] or f"ستون {to_persian_digits(column + 1)}"
    base = f"{role_title} | {job.label}"
    title = base
    suffix = 2
    while title in used_titles:
        title = f"{base} ({suffix})"
        suffix += 1
    used_titles.add(title)
    return SourceInfo(
        key=f"{job.file_id}:{role}:{column}",
        title=title,
        role=role,
        label=job.label,
        filename=job.filename,
        sheet=job.sheet,
        header_row=job.header_row,
        amount_header=amount_header,
        factor=job.factor,
    )


def _amount_column(job: FileJob, source: SourceInfo) -> int:
    if job.role == "both" and source.role == ROLE_DUE:
        assert job.columns.amount_due is not None
        return job.columns.amount_due
    if job.role == "both" and source.role == ROLE_PAID:
        assert job.columns.amount_paid is not None
        return job.columns.amount_paid
    assert job.columns.amount is not None
    return job.columns.amount


def _consume_job(job: FileJob, sources: list[SourceInfo], people: dict[str, Person], warnings: list[str]) -> None:
    missing_identity = 0
    bad_by_source: dict[str, list[str]] = {source.key: [] for source in sources}
    data = job.rows[job.header_row :]
    for offset, row in enumerate(data):
        if all(is_blank(cell) for cell in row):
            continue
        excel_row = job.header_row + offset + 1
        parsed: dict[str, Decimal] = {}
        for source in sources:
            column = _amount_column(job, source)
            try:
                amount = parse_amount(_cell(row, column))
            except ValueError:
                source.skipped_rows += 1
                if len(bad_by_source[source.key]) < 12:
                    bad_by_source[source.key].append(to_persian_digits(excel_row))
                continue
            parsed[source.key] = quantize_money((amount or Decimal("0")) * job.factor)
        if not parsed:
            continue
        identity = _identity(job, row)
        if identity is None:
            missing_identity += 1
            for source in sources:
                if source.key in parsed:
                    source.skipped_rows += 1
            continue
        key, person_seed, name_matched = identity
        for source in sources:
            if source.key not in parsed:
                continue
            person = people.get(key)
            if person is None:
                person = person_seed
                people[key] = person
            else:
                _merge_identity(person, person_seed, source.role)
            if name_matched:
                person.matched_by_name = True
            person.by_source[source.key] = quantize_money(
                person.by_source.get(source.key, Decimal("0")) + parsed[source.key]
            )
            for item in person_seed.notes:
                _add_note(person, item)
            source.used_rows += 1
            source.total = quantize_money(source.total + parsed[source.key])
        note = _note_text(job, row)
        if note and key in people:
            _add_note(people[key], note)

    if missing_identity:
        warnings.append(
            f"در «{job.label}»، {to_persian_digits(missing_identity)} ردیف به‌دلیل نداشتن کد ملی و نام وارد محاسبه نشد."
        )
    for source in sources:
        lines = bad_by_source[source.key]
        if lines:
            shown = "، ".join(lines)
            warnings.append(
                f"در فهرست «{source.title}» مبلغ ردیف {shown} نامعتبر بود و در جمع نیامد."
            )
        if source.used_rows == 0:
            warnings.append(f"از فهرست «{source.title}» ردیف قابل استفاده‌ای به دست نیامد.")


def _cell(row: list, index: int | None):
    if index is None or index >= len(row):
        return None
    return row[index]


def _identity(job: FileJob, row: list) -> tuple[str, Person, bool] | None:
    columns = job.columns
    national_id, id_note = normalize_national_id(_cell(row, columns.national_id))
    first = display_fa(preview_cell(_cell(row, columns.first_name))) if columns.first_name is not None else ""
    last = display_fa(preview_cell(_cell(row, columns.last_name))) if columns.last_name is not None else ""
    full = display_fa(preview_cell(_cell(row, columns.full_name))) if columns.full_name is not None else ""
    if full and not (first and last):
        guessed_first, guessed_last = split_full_name(full)
        first = first or guessed_first
        last = last or guessed_last
    if not full:
        full = display_fa(f"{first} {last}")
    person = Person(national_id=national_id, first_name=first, last_name=last, full_name=full)
    if id_note:
        _add_note(person, id_note)
    if national_id:
        if len(national_id) == 10 and not national_id_is_valid(national_id):
            _add_note(person, "رقم کنترلی کد ملی نادرست است")
        return f"id:{national_id}", person, False
    if loose_name(full):
        return f"name:{loose_name(full)}", person, True
    return None


def _merge_identity(person: Person, incoming: Person, role: str) -> None:
    if person.national_id and incoming.national_id and person.national_id != incoming.national_id:
        _add_note(person, "کد ملی در فایل‌ها یکسان نیست")
    elif not person.national_id:
        person.national_id = incoming.national_id
    _take_name(person, incoming, role)


def _take_name(person: Person, incoming: Person, role: str) -> None:
    incoming_full = incoming.full_display
    current_full = person.full_display
    conflict = bool(current_full and incoming_full and loose_name(current_full) != loose_name(incoming_full))
    if conflict:
        _add_note(person, "نام در فایل‌ها یکسان نیست")
    replace = role == ROLE_DUE or not current_full
    if not replace:
        return
    if incoming.first_name or not person.first_name:
        if role == ROLE_DUE or not person.first_name:
            person.first_name = incoming.first_name or person.first_name
    if incoming.last_name or not person.last_name:
        if role == ROLE_DUE or not person.last_name:
            person.last_name = incoming.last_name or person.last_name
    if incoming.full_name and (role == ROLE_DUE or not person.full_name):
        person.full_name = incoming.full_name
    elif not person.full_name:
        person.full_name = display_fa(f"{person.first_name} {person.last_name}")


def _note_text(job: FileJob, row: list) -> str:
    if job.columns.note is None:
        return ""
    text = display_fa(preview_cell(_cell(row, job.columns.note)))
    if len(text) > 180:
        text = text[:179].rstrip() + "…"
    return text


def _add_note(person: Person, text: str) -> None:
    cleaned = display_fa(text)
    if cleaned and cleaned not in person.notes:
        person.notes.append(cleaned)


def job_from_payload(item: dict, sheets: dict[str, list[list]], filename: str) -> FileJob:
    sheet_name = str(item.get("sheet") or "")
    if sheet_name not in sheets:
        raise UserError(f"برگه «{sheet_name}» در فایل «{filename}» نیست.")
    try:
        header_row = int(item.get("header_row") or 1)
        factor = parse_decimal_input(item.get("factor", "1"))
    except (TypeError, ValueError) as exc:
        raise UserError(f"ردیف عنوان یا ضریب مبلغ در «{filename}» عدد نیست.") from exc
    columns_raw = item.get("columns") or {}

    def optional(name: str) -> int | None:
        value = columns_raw.get(name)
        if value is None or value == "":
            return None
        try:
            return int(value)
        except (TypeError, ValueError) as exc:
            raise UserError(f"ستون انتخاب‌شده در «{filename}» نامعتبر است.") from exc

    label = safe_label(str(item.get("label") or ""), filename.rsplit(".", 1)[0])
    role = str(item.get("role") or ROLE_DUE)
    return FileJob(
        file_id=str(item["id"]),
        filename=filename,
        sheet=sheet_name,
        header_row=header_row,
        label=label,
        role=role,
        factor=factor,
        columns=ColumnMap(
            national_id=optional("national_id"),
            first_name=optional("first_name"),
            last_name=optional("last_name"),
            full_name=optional("full_name"),
            amount=optional("amount"),
            amount_due=optional("amount_due"),
            amount_paid=optional("amount_paid"),
            note=optional("note"),
        ),
        rows=sheets[sheet_name],
    )
