"""Persian text helpers.

Excel must receive normal Unicode Persian. Do not reshape glyphs and do not
reverse strings; Excel lays the text out itself when the sheet is right-to-left.
"""

from __future__ import annotations

import math
import re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

PERSIAN_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
TO_PERSIAN_DIGITS = str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")
ARABIC_LETTERS = str.maketrans(
    {
        "ي": "ی",
        "ك": "ک",
        "ة": "ه",
        "ؤ": "و",
        "أ": "ا",
        "إ": "ا",
        "ٱ": "ا",
        "ٰ": "",
    }
)

TWOPLACES = Decimal("0.01")
CURRENCY_WORDS = ("ریال", "تومان", "rial", "irr")


def to_persian_digits(value: object) -> str:
    return str(value).translate(TO_PERSIAN_DIGITS)


def fold_digits(value: str) -> str:
    return value.translate(PERSIAN_DIGITS)


def display_fa(value: object) -> str:
    """Clean a piece of text for showing, without reversing it."""
    text = str(value).translate(ARABIC_LETTERS)
    text = text.replace("\u200b", "").replace("\ufeff", "")
    text = re.sub(r"[ \t\r\f\v]+", " ", text)
    return text.strip()


def fold(value: object) -> str:
    text = display_fa(value).replace("\u200c", "").replace("\u200d", "")
    text = re.sub(r"\s+", " ", text)
    return text.strip().casefold()


def loose_name(value: object) -> str:
    text = fold(value).replace("ئ", "ی").replace("آ", "ا").replace(" ", "")
    return text


def fa_sort_key(value: object) -> tuple:
    alphabet = "آابپتثجچحخدذرزژسشصضطظعغفقکگلمنوهی"
    order = {char: index for index, char in enumerate(alphabet)}
    return tuple(order.get(char, 200 + ord(char)) for char in fold(value))


def is_blank(value: object) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        return True
    return str(value).strip() == ""


def preview_cell(value: object) -> str:
    if is_blank(value):
        return ""
    if isinstance(value, bool):
        return "بله" if value else "خیر"
    if isinstance(value, int) and not isinstance(value, bool):
        return str(value)
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return format(value, "f").rstrip("0").rstrip(".")
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, Decimal):
        return format(value, "f")
    return str(value).strip()


def quantize_money(value: Decimal) -> Decimal:
    amount = value.quantize(TWOPLACES, rounding=ROUND_HALF_UP)
    if amount == amount.to_integral_value():
        return amount.quantize(Decimal("1"))
    return amount


def parse_amount(value: object) -> Decimal | None:
    """Return None for an empty cell. Raise ValueError when the text is not a number."""
    if is_blank(value):
        return None
    if isinstance(value, Decimal):
        return quantize_money(value)
    if isinstance(value, bool):
        raise ValueError("boolean")
    if isinstance(value, int):
        return quantize_money(Decimal(value))
    if isinstance(value, float):
        return quantize_money(Decimal(str(value)))

    text = fold_digits(str(value)).strip()
    if text.lower() in {"nan", "none", "null"}:
        return None
    for word in CURRENCY_WORDS:
        text = re.sub(word, "", text, flags=re.IGNORECASE)
    text = (
        text.replace("\u200c", "")
        .replace(" ", "")
        .replace("٬", ",")
        .replace("،", ",")
        .replace("٫", ".")
        .replace("−", "-")
        .replace("ـ", "")
    )
    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1]
    if text[:1] in "+-":
        negative = text[0] == "-"
        text = text[1:]
    if text == "":
        return None

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        parts = text.split(",")
        if len(parts) > 1 and all(len(part) == 3 for part in parts[1:]) and 1 <= len(parts[0]) <= 3:
            text = "".join(parts)
        elif len(parts) == 2 and 1 <= len(parts[1]) <= 2:
            text = parts[0] + "." + parts[1]
        else:
            text = text.replace(",", "")
    elif text.count(".") > 1:
        parts = text.split(".")
        if all(len(part) == 3 for part in parts[1:]):
            text = "".join(parts)
    elif text.count(".") == 1:
        whole, frac = text.split(".")
        if len(frac) == 3 and 1 <= len(whole) <= 3:
            text = whole + frac

    if not re.fullmatch(r"\d+(\.\d+)?", text):
        raise ValueError(text)
    amount = Decimal(text)
    if negative:
        amount = -amount
    return quantize_money(amount)


def normalize_national_id(value: object) -> tuple[str, str | None]:
    """Return (code, note). Empty code means the cell had no identifier."""
    if is_blank(value):
        return "", None
    if isinstance(value, float):
        if not value.is_integer():
            raw = preview_cell(value)
        else:
            raw = str(int(value))
    elif isinstance(value, int) and not isinstance(value, bool):
        raw = str(value)
    else:
        raw = preview_cell(value)
        if re.fullmatch(r"\d+\.0+", fold_digits(raw)):
            raw = fold_digits(raw).split(".")[0]

    digits = re.sub(r"\D", "", fold_digits(raw))
    if not digits:
        return "", None
    note = None
    if len(digits) in {8, 9}:
        digits = digits.zfill(10)
        note = "صفر ابتدای کد ملی اضافه شد تا ده رقم شود"
    elif len(digits) != 10:
        note = "کد ملی ده رقم نیست"
    return digits, note


def national_id_is_valid(code: str) -> bool:
    if len(code) != 10 or not code.isdigit() or code == code[0] * 10:
        return False
    total = sum(int(code[index]) * (10 - index) for index in range(9))
    remainder = total % 11
    check = int(code[9])
    if remainder < 2:
        return check == remainder
    return check == 11 - remainder


def with_check_digit(first_nine: str) -> str:
    if len(first_nine) != 9 or not first_nine.isdigit():
        raise ValueError("nine digits required")
    total = sum(int(first_nine[index]) * (10 - index) for index in range(9))
    remainder = total % 11
    check = remainder if remainder < 2 else 11 - remainder
    return first_nine + str(check)


def split_full_name(full_name: str) -> tuple[str, str]:
    parts = display_fa(full_name).split()
    if not parts:
        return "", ""
    if len(parts) == 1:
        return parts[0], ""
    return " ".join(parts[:-1]), parts[-1]


def gregorian_to_jalali(gy: int, gm: int, gd: int) -> tuple[int, int, int]:
    g_days_in_month = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    j_days_in_month = [31, 31, 31, 31, 31, 31, 30, 30, 30, 30, 30, 29]
    gy -= 1600
    gm -= 1
    gd -= 1
    g_day_no = 365 * gy + (gy + 3) // 4 - (gy + 99) // 100 + (gy + 399) // 400
    for index in range(gm):
        g_day_no += g_days_in_month[index]
    leap = (gy + 1600) % 4 == 0 and ((gy + 1600) % 100 != 0 or (gy + 1600) % 400 == 0)
    if gm > 1 and leap:
        g_day_no += 1
    g_day_no += gd
    j_day_no = g_day_no - 79
    j_np = j_day_no // 12053
    j_day_no %= 12053
    jy = 979 + 33 * j_np + 4 * (j_day_no // 1461)
    j_day_no %= 1461
    if j_day_no >= 366:
        jy += (j_day_no - 1) // 365
        j_day_no = (j_day_no - 1) % 365
    for index in range(11):
        if j_day_no < j_days_in_month[index]:
            return jy, index + 1, j_day_no + 1
        j_day_no -= j_days_in_month[index]
    return jy, 12, j_day_no + 1


def format_jalali(moment: datetime) -> str:
    year, month, day = gregorian_to_jalali(moment.year, moment.month, moment.day)
    text = f"{year:04d}/{month:02d}/{day:02d}"
    return to_persian_digits(text)


def parse_decimal_input(value: object) -> Decimal:
    if isinstance(value, Decimal):
        return value
    text = fold_digits(str(value)).strip().replace(",", "").replace("٬", "")
    try:
        return Decimal(text)
    except (InvalidOperation, ValueError) as exc:
        raise ValueError("number") from exc


def fold_header(value: object) -> str:
    return fold(value).replace(" ", "").replace("\u200c", "")


MONTHS = (
    "فروردین",
    "اردیبهشت",
    "خرداد",
    "تیر",
    "مرداد",
    "شهریور",
    "مهر",
    "آبان",
    "آذر",
    "دی",
    "بهمن",
    "اسفند",
)
_MONTH_ALIASES = {
    "farvardin": "فروردین",
    "ordibehesht": "اردیبهشت",
    "khordad": "خرداد",
    "tir": "تیر",
    "mordad": "مرداد",
    "shahrivar": "شهریور",
    "mehr": "مهر",
    "aban": "آبان",
    "azar": "آذر",
    "dey": "دی",
    "day": "دی",
    "bahman": "بهمن",
    "esfand": "اسفند",
}


def parse_month(value: object) -> str | None:
    """Read a Persian month from a cell. Empty cells stay empty."""
    if is_blank(value):
        return None
    if isinstance(value, datetime):
        _, month, _ = gregorian_to_jalali(value.year, value.month, value.day)
        return MONTHS[month - 1]
    if isinstance(value, date):
        _, month, _ = gregorian_to_jalali(value.year, value.month, value.day)
        return MONTHS[month - 1]
    text = display_fa(preview_cell(value))
    folded = fold(text)
    compact = folded.replace("ماه", "").strip()
    for name in MONTHS:
        if fold(name) == compact:
            return name
    for alias, name in _MONTH_ALIASES.items():
        if alias == compact:
            return name
    for name in MONTHS:
        if len(name) >= 4 and fold(name) in folded:
            return name
    digits = re.sub(r"\D", "", fold_digits(text))
    if re.fullmatch(r"\d{1,2}", digits):
        number = int(digits)
        if 1 <= number <= 12:
            return MONTHS[number - 1]
    if re.fullmatch(r"\d{4}\d{2}\d{2}", digits) and len(digits) == 8:
        number = int(digits[4:6])
        if 1 <= number <= 12:
            return MONTHS[number - 1]
    return None


def normalize_code(value: object) -> str:
    """Keep an insurance number as text, including letters, without reversing it."""
    if is_blank(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        text = str(int(value))
    elif isinstance(value, int) and not isinstance(value, bool):
        text = str(value)
    else:
        text = fold_digits(preview_cell(value))
        if re.fullmatch(r"\d+\.0+", text):
            text = text.split(".")[0]
    text = display_fa(text)
    return re.sub(r"\s+", "", text)


def suggest_columns(headers: list[str]) -> dict[str, int | None]:
    folded = [fold_header(header) for header in headers]

    def find(predicates: tuple[str, ...], avoid: tuple[str, ...] = ()) -> int | None:
        needles = tuple(fold_header(item) for item in predicates)
        blocked = tuple(fold_header(item) for item in avoid)
        for index, header in enumerate(folded):
            if any(block in header for block in blocked):
                continue
            if any(needle in header for needle in needles):
                return index
        return None

    full_name = find(("نام و نام خانوادگی", "نام کامل", "نام نام خانوادگی"))
    # «و» داخل «خانوادگی» هست؛ فقط الگوی نامِ کامل را کنار می‌گذاریم.
    last_name = find(
        ("نام خانوادگی", "نامخانوادگی", "فامیل"),
        avoid=("نامونام", "نامکامل", "نامنام"),
    )
    first_name = find(("نام", "اسم"), avoid=("خانواد", "فامیل", "کامل", "پدر", "مادر", "بیمه"))
    if full_name is not None and first_name == full_name:
        first_name = None
    if full_name is not None and last_name == full_name:
        last_name = None
    if last_name is not None and first_name == last_name:
        first_name = None

    return {
        "policyholder": find(("نام بیمه گذار", "بیمه گذار", "بیمهگزار")),
        "insurance_no": find(("شماره بیمه", "شماره بیمه نامه", "کد بیمه")),
        "insurance_type": find(("نوع بیمه", "رشته بیمه", "نوع بیمه نامه")),
        "month": find(("نام ماه", "ماه پرداخت", "ماه"), avoid=("شماره", "مبلغ", "بیمه")),
        "national_id": find(("کد ملی", "کدملی", "شناسه ملی", "شماره ملی", "کد شناسایی")),
        "first_name": first_name,
        "last_name": last_name,
        "full_name": full_name,
        "amount": find(("مبلغ", "هزینه", "شهریه", "بدهکار", "بستانکار", "واریز", "پرداخت")),
        "amount_due": find(("قابل پرداخت", "مبلغ هزینه", "هزینه", "شهریه", "بدهکار")),
        "amount_paid": find(
            ("مبلغ پرداخت", "پرداخت شده", "واریز", "پرداختی"),
            avoid=("قابل",),
        ),
        "note": find(("توضیح", "شرح", "بابت", "یادداشت")),
    }


def suggest_role(filename: str) -> str:
    name = fold(filename)
    paid = any(word in name for word in ("پرداخت", "واریز", "دریافت", "فیش", "صندوق"))
    due = any(word in name for word in ("هزینه", "شهریه", "بده", "قابل"))
    if paid and not due:
        return "paid"
    return "due"


def safe_label(value: str, fallback: str) -> str:
    text = display_fa(value)
    text = re.sub(r"[\[\]\:\*\?\/\\]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        text = fallback
    if len(text) > 40:
        text = text[:39].rstrip() + "…"
    return text
