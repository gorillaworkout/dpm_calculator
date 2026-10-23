"""One number parser for everything a person types or a file carries (PR #26).

Before this module every tool had its own `float(...)` / `Decimal(...)` with its
own rules. `--jwallet-opening '3.5E+09'` was refused, '1,5%' read as 5%, and a
balance cell stored as text was skipped (so the balance silently became 0).

    parse_number(value, ambiguous=DOT_DECIMAL, as_decimal=False, default=RAISE)

Accepted
    plain numbers and decimals    1234   1234.5   .5   1234.
    scientific notation           3.5E+09   3.5e9   3,5E+09   -1.2e-3
    Indonesian thousands          3.500.000.000   3.500.000,50
    English thousands             3,500,000,000   3,500,000.50
    Indian grouping (',' only)    1,00,00,000.50
    negatives                     -1234   (1,234)   1234-   −1234 (unicode minus)
    spaces, NBSP, thin spaces, apostrophes and '_' between digits
    a currency code (KODE_MATA_UANG) or symbol before or after the number:
        Rp 3.500.000   Rp. 3.500.000   IDR 3,5E+09   USD 12.50   12.50 USDT   $1,234   US$ 5
    Excel numbers (int / float / Decimal) are returned unchanged.

Rules for '.' and ','
  1. Both appear: the LAST one is the decimal point, the other is thousands.
  2. One kind appears more than once: thousands ('3.500.000' -> 3500000).
  3. One single separator:
       - in the mantissa of scientific notation      -> decimal ('3,5E+09')
       - integer part is empty or 0 ('0.500', ',5')   -> decimal
       - followed by 1-2 or 4+ digits                 -> decimal ('377,23', '0,0275')
       - followed by EXACTLY 3 digits: AMBIGUOUS ('1.234', '1,234'). Resolved by
         the caller's convention, which is what the code already did there:
           DOT_DECIMAL (files, rates; the old `to_float`): '.' is the decimal
               point, ',' is thousands. '1.234' -> 1.234, '1,234' -> 1234.
           THOUSANDS (amounts typed by a person on the upload page; the old
               `baca_angka_saldo`): both are thousands. '1.234' / '1,234' -> 1234.
         A value marked 'Rp' / 'IDR' always uses THOUSANDS (rupiah amounts are
         written '1.234' = one thousand two hundred thirty-four), so
         'Rp 1.000' -> 1000 even in a DOT_DECIMAL field.
  Thousands groups must be well formed: 1-3 digits, then groups of exactly 3
  ('1.2.3' and '12,34.5' are refused rather than guessed).

Refused (NumberError, or `default` when given): empty-after-cleanup text, letters
other than a currency code / exponent, '%', two decimal points, bad grouping,
inf / nan, and booleans. Empty input (None, '', spaces) returns None.
"""
import math
import re
from decimal import Decimal, InvalidOperation
from pathlib import Path

DOT_DECIMAL = "dot-decimal"
THOUSANDS = "thousands"
RAISE = object()

APP_NAME = "DW Calculator"


class NumberError(ValueError):
    """The value is not a number this parser can read."""

    def __init__(self, value, reason=""):
        self.value = value
        self.reason = reason
        super().__init__(f"bukan angka yang bisa dibaca: {value!r}" + (f" ({reason})" if reason else ""))


_SPASI = re.compile(r"[\s\u00a0\u2007\u2009\u202f'\u2019_]")
_SIMBOL = "$€£¥₫฿₱₹₩"
# Currency codes that may stand before or after a number (case-insensitive).
# Anything else with letters ('12abc', '1.5k') is refused, not guessed.
KODE_MATA_UANG = {
    "RP", "IDR", "USD", "USDT", "USDC", "US", "EUR", "GBP", "SGD", "MYR", "THB",
    "VND", "PHP", "CNY", "RMB", "JPY", "KRW", "INR", "HKD", "AUD", "NZD", "CAD",
    "CHF", "TWD", "LAK", "TRY", "PKR", "EGP", "KES", "AED", "BRL", "MXN", "NGN",
    "ZAR", "KHR", "UZS", "LKR", "BDT", "MMK", "SAR",
}
_KODE = "|".join(sorted(KODE_MATA_UANG, key=len, reverse=True))
_MATA_UANG_DEPAN = re.compile(r"^(?:(?:US|S|A|HK|NZ|C)\$|(?:" + _KODE + r")\.?|[" + _SIMBOL + r"])", re.I)
_MATA_UANG_BELAKANG = re.compile(r"(?:" + _KODE + r"|[" + _SIMBOL + r"])$", re.I)
_RUPIAH = {"RP", "IDR"}
_POLA = re.compile(r"(?P<m>[\d.,]*\d[\d.,]*|[.,]\d+)(?:[eE](?P<e>[+-]?\d+))?")
_POLOS = re.compile(r"[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?")


def _kelompok_sah(bagian, sep):
    """Integer part split on the thousands separator: 1-3 digits, then 3s
    (or Indian 2s ending in a 3, only for ',')."""
    if not bagian[0] or len(bagian[0]) > 3 or not all(b.isdigit() for b in bagian):
        return False
    ekor = bagian[1:]
    if all(len(b) == 3 for b in ekor):
        return True
    return (sep == "," and len(ekor) >= 2 and len(ekor[-1]) == 3
            and all(len(b) == 2 for b in ekor[:-1]))


def _normal(mantissa, ambiguous, ilmiah):
    """Mantissa with '.'/',' -> plain '1234.5' text, or None when unreadable."""
    titik, koma = mantissa.count("."), mantissa.count(",")
    if titik and koma:
        des = "." if mantissa.rfind(".") > mantissa.rfind(",") else ","
        rib = "," if des == "." else "."
        if mantissa.count(des) > 1:
            return None
        bulat, pecahan = mantissa.split(des)
        if rib not in bulat or not _kelompok_sah(bulat.split(rib), rib) or not pecahan.isdigit():
            return None
        return bulat.replace(rib, "") + "." + pecahan
    if not titik and not koma:
        return mantissa
    sep = "." if titik else ","
    if mantissa.count(sep) > 1:
        bagian = mantissa.split(sep)
        return "".join(bagian) if _kelompok_sah(bagian, sep) else None
    bulat, pecahan = mantissa.split(sep)
    if not (bulat.isdigit() or bulat == "") or not (pecahan.isdigit() or pecahan == ""):
        return None
    if pecahan == "":                       # '1234.' -> 1234
        return bulat
    decimal = (ilmiah or bulat in ("", "0") or len(pecahan) != 3
               or (ambiguous == DOT_DECIMAL and sep == "."))
    if not decimal and len(bulat) > 3:      # '1234,567' is not a thousands group
        return None
    return (bulat or "0") + "." + pecahan if decimal else bulat + pecahan


def parse_number(value, ambiguous=DOT_DECIMAL, as_decimal=False, default=RAISE):
    """-> float (or Decimal), None for empty input. See the module docstring."""
    def gagal(reason=""):
        if default is RAISE:
            raise NumberError(value, reason)
        return default

    if value is None:
        return None
    if isinstance(value, bool):
        return gagal("boolean")
    if isinstance(value, (int, float, Decimal)):
        if isinstance(value, float) and not math.isfinite(value):
            return gagal("not finite")
        return Decimal(str(value)) if as_decimal else float(value)
    if ambiguous not in (DOT_DECIMAL, THOUSANDS):
        raise ValueError(f"unknown convention {ambiguous!r}")
    s = str(value).strip()
    if not s:
        return None

    # Fast path: what float() already reads the same way under the rules above.
    if _POLOS.fullmatch(s) and (ambiguous == DOT_DECIMAL or "." not in s
                                or "e" in s.lower()):
        return Decimal(s) if as_decimal else float(s)

    s = _SPASI.sub("", s).replace("\u2212", "-").replace("\u2012", "-").replace("\u2013", "-")
    if not s:
        return None
    neg = False
    tanda = kurung = 0
    kode = []
    for _ in range(8):
        sebelum = s
        if s.startswith("(") and s.endswith(")") and len(s) > 2:
            s, kurung = s[1:-1], kurung + 1
        if s[:1] in "+-":
            tanda += 1
            neg ^= s[0] == "-"
            s = s[1:]
        elif s.endswith("-") and len(s) > 1 and s[-2].isdigit():
            tanda += 1
            neg = not neg
            s = s[:-1]
        m = _MATA_UANG_DEPAN.match(s)
        if m and len(m.group()) < len(s):
            kode.append(m.group())
            s = s[m.end():]
        m = _MATA_UANG_BELAKANG.search(s)
        if m and m.start() > 0 and (s[m.start() - 1].isdigit() or s[m.start() - 1] in ".,)"):
            kode.append(m.group())
            s = s[:m.start()]
        if s == sebelum or not s:
            break
    if not s:
        return gagal("no digits")
    if kurung > 1 or tanda > 1 or (kurung and tanda):
        return gagal("more than one minus sign")
    if kurung:
        neg = True
    if any(k.rstrip(".").upper() in _RUPIAH for k in kode):
        ambiguous = THOUSANDS
    m = _POLA.fullmatch(s)
    if not m:
        return gagal()
    teks = _normal(m.group("m"), ambiguous, m.group("e") is not None)
    if teks is None:
        return gagal("thousands/decimal separators do not fit")
    if m.group("e") is not None:
        teks += "e" + m.group("e")
    try:
        hasil = Decimal(teks)
    except InvalidOperation:
        return gagal()
    if not hasil.is_finite():
        return gagal("not finite")
    if neg:
        hasil = -hasil
    if as_decimal:
        return hasil
    f = float(hasil)
    return f if math.isfinite(f) else gagal("too large")


def version_label(path=None):
    """'DW Calculator v1.26 · 2026-10-23' from the VERSION file next to this module."""
    path = Path(path) if path else Path(__file__).resolve().parent / "VERSION"
    info = {"version": "unknown", "build": ""}
    try:
        teks = path.read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError):
        teks = ""
    for baris in teks.splitlines():
        baris = baris.strip()
        if not baris or baris.startswith("#"):
            continue
        if "=" in baris:
            k, _, v = baris.partition("=")
            if k.strip().lower() in info and v.strip():
                info[k.strip().lower()] = v.strip()
        elif info["version"] == "unknown":
            info["version"] = baris
    v = info["version"]
    label = f"{APP_NAME} v{v}" if v != "unknown" else f"{APP_NAME} version unknown"
    return f"{label} · {info['build']}" if info["build"] else label
