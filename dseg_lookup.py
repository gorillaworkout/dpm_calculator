"""UID / KYC Country lookup for the DPM Deal Segregator (v1.50).

The DPM Malaysia team wants two columns at the front of both Monthly Summary
sheets: ``KYC Country`` and ``UID``. Neither is in the MT5 Deals History, so
they come from two files that rarely change and are therefore kept on the
server (not uploaded with every run, never committed to git):

* account file  -- columns ``user_id`` and ``account``  (one row per trading login)
* KYC file      -- columns ``user_id`` and ``Country of KYC``

Lookup (same as the team's XLOOKUP): Login -> account -> user_id = UID, then
UID -> user_id -> Country of KYC. The first matching row wins, like XLOOKUP.

Location: ``DW_DSEG_LOOKUP_DIR`` or ``dseg_lookup/`` next to users.db (the same
rule as approval_files/ and invitation_files/), so a deploy that copies code
files never touches it. Inside it the files are stored as
``account_uid.xlsx`` and ``uid_kyc_country.xlsx`` (.csv also accepted).

Parsing 107k + 374k rows of xlsx takes ~5 s, so the parsed dicts are cached in
``.lookup_cache.pkl`` in the same folder, keyed by each file's size + mtime.
A changed file rebuilds the cache on the next run.
"""
import csv
import datetime
import io
import os
import pickle
import re
import tempfile
import warnings
from pathlib import Path

HERE = Path(__file__).resolve().parent

KIND_ACCOUNT = "account"
KIND_KYC = "kyc"
STEM = {KIND_ACCOUNT: "account_uid", KIND_KYC: "uid_kyc_country"}
# Normalised header names (lower case, single spaces) each file must have.
HEADERS = {KIND_ACCOUNT: ("user_id", "account"), KIND_KYC: ("user_id", "country of kyc")}
SHOW = {KIND_ACCOUNT: "user_id, account", KIND_KYC: "user_id, Country of KYC"}
LABEL = {KIND_ACCOUNT: "Account file (UID)", KIND_KYC: "KYC file (Country of KYC)"}
EXTS = (".xlsx", ".xlsm", ".csv")
CACHE = ".lookup_cache.pkl"
_CACHE_VERSION = 1

_memo = {}


def lookup_dir() -> Path:
    env = os.environ.get
    if env("DW_DSEG_LOOKUP_DIR"):
        return Path(env("DW_DSEG_LOOKUP_DIR")).expanduser()
    auth_db = Path(env("DW_AUTH_DB") or HERE / "users.db").expanduser()
    return auth_db.with_name("dseg_lookup")


def _norm_header(v) -> str:
    return re.sub(r"\s+", " ", str(v if v is not None else "")).strip().lower()


_INT_FLOAT = re.compile(r"^[+-]?\d+\.0*$")


def norm_id(v) -> str:
    """'503617', 503617, 503617.0 and ' 503617 ' are the same id."""
    if v is None or isinstance(v, bool):
        return ""
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        if v != v:
            return ""
        return str(int(v)) if v.is_integer() else repr(v)
    s = str(v).strip()
    if _INT_FLOAT.match(s):
        s = s.split(".", 1)[0]
    return s


def _uid_value(key: str):
    """UID written as a number when it is a plain integer (as XLOOKUP returns it)."""
    if key.isdigit() and (key == "0" or not key.startswith("0")) and len(key) <= 15:
        return int(key)
    return key


def stored_path(kind, folder=None):
    folder = Path(folder) if folder else lookup_dir()
    for ext in EXTS:
        p = folder / f"{STEM[kind]}{ext}"
        if p.is_file():
            return p
    return None


# ------------------------------------------------------------------ reading
def _rows_csv(path: Path):
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "utf-16", "cp1252"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError("the file is not readable text")
    first = text.split("\n", 1)[0]
    delim = max((",", ";", "\t"), key=first.count)
    yield from csv.reader(io.StringIO(text), delimiter=delim)


def _rows_xlsx(path: Path):
    """Rows of the first sheet whose first row has the needed headers (any sheet)."""
    from openpyxl import load_workbook
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = load_workbook(path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            yield ws.title
            for row in ws.iter_rows(values_only=True):
                yield row
            yield None          # end of sheet
    finally:
        wb.close()


def _header_index(row, kind):
    names = [_norm_header(v) for v in (row or ())]
    need = HEADERS[kind]
    if all(n in names for n in need):
        return [names.index(n) for n in need]
    return None


def _iter_pairs(path: Path, kind):
    """(header_ok, iterator of (key, value)) for the first sheet with the right header."""
    path = Path(path)
    if path.suffix.lower() == ".csv":
        rows = _rows_csv(path)
        header = next(rows, None)
        idx = _header_index(header, kind)
        if idx is None:
            return None, header
        return idx, rows
    gen = _rows_xlsx(path)
    first_header = None

    def sheet_rows():
        for row in gen:
            if row is None:
                return
            yield row

    for item in gen:
        if not isinstance(item, str):
            continue
        rows = sheet_rows()
        header = next(rows, None)
        if first_header is None:
            first_header = header
        idx = _header_index(header, kind)
        if idx is not None:
            return idx, rows
        for _ in rows:          # skip the rest of this sheet
            pass
    return None, first_header


def parse(path, kind):
    """Parse one lookup file into {key: value}. Raises ValueError on a bad header."""
    try:
        idx, rows = _iter_pairs(path, kind)
    except ValueError:
        raise
    except Exception as exc:   # corrupt workbook, wrong format
        raise ValueError(f"the file could not be read ({type(exc).__name__})") from None
    if idx is None:
        found = ", ".join(str(v) for v in (rows or ()) if v not in (None, "")) or "(empty)"
        raise ValueError(f"the header must have the columns {SHOW[kind]}; found: {found[:200]}")
    i_user, i_other = idx
    out, duplicates, conflicts, rows_n = {}, 0, 0, 0
    width = max(idx) + 1
    for row in rows:
        if row is None:
            continue
        row = list(row) + [None] * (width - len(row))
        user, other = row[i_user], row[i_other]
        if kind == KIND_ACCOUNT:
            key, val = norm_id(other), norm_id(user)
        else:
            key = norm_id(user)
            val = "" if other is None else str(other).strip()
        if not key:
            continue
        rows_n += 1
        if key in out:
            duplicates += 1
            if out[key] != val:
                conflicts += 1
            continue
        out[key] = val
    return {"map": out, "rows": rows_n, "duplicates": duplicates, "conflicts": conflicts}


def _signature(path: Path):
    st = path.stat()
    return (path.name, st.st_size, st.st_mtime_ns)


def file_date(path: Path) -> str:
    ts = datetime.datetime.fromtimestamp(path.stat().st_mtime)
    return ts.strftime("%Y-%m-%d %H:%M")


def load(folder=None):
    """Both lookups, or a dict with 'missing' set. Cached by file size + mtime."""
    folder = Path(folder) if folder else lookup_dir()
    paths = {k: stored_path(k, folder) for k in STEM}
    missing = [k for k, p in paths.items() if p is None]
    if missing:
        return {"ok": False, "missing": missing, "folder": str(folder), "paths": paths}
    sig = tuple(_signature(paths[k]) for k in (KIND_ACCOUNT, KIND_KYC))
    key = (str(folder), sig)
    if key in _memo:
        return _memo[key]
    cache = folder / CACHE
    data = None
    try:
        with open(cache, "rb") as fh:
            got = pickle.load(fh)
        if got.get("v") == _CACHE_VERSION and got.get("sig") == sig:
            data = got["data"]
    except Exception:
        data = None
    if data is None:
        data = {k: parse(paths[k], k) for k in (KIND_ACCOUNT, KIND_KYC)}
        try:
            fd, tmp = tempfile.mkstemp(dir=folder, prefix=".cache-", suffix=".tmp")
            with os.fdopen(fd, "wb") as fh:
                pickle.dump({"v": _CACHE_VERSION, "sig": sig, "data": data}, fh,
                            protocol=pickle.HIGHEST_PROTOCOL)
            os.replace(tmp, cache)
        except OSError:
            pass
    result = {"ok": True, "missing": [], "folder": str(folder), "paths": paths,
              "data": data, "dates": {k: file_date(paths[k]) for k in paths}}
    _memo.clear()
    _memo[key] = result
    return result


def status(folder=None):
    """For the tool page: which files are stored and their dates (no parsing)."""
    folder = Path(folder) if folder else lookup_dir()
    out = []
    for kind in (KIND_ACCOUNT, KIND_KYC):
        p = stored_path(kind, folder)
        out.append({"kind": kind, "label": LABEL[kind], "columns": SHOW[kind],
                    "present": p is not None, "name": p.name if p else None,
                    "date": file_date(p) if p else None})
    return out


def save(kind, src, original_name, folder=None):
    """Validate an uploaded lookup file and store it over the current one.

    Returns the parse stats. Raises ValueError (English text) when the file is
    not usable; the stored file is then left as it was.
    """
    if kind not in STEM:
        raise ValueError("unknown lookup file")
    ext = Path(original_name or "").suffix.lower()
    if ext not in EXTS:
        raise ValueError(f"{LABEL[kind]}: use a .xlsx, .xlsm or .csv file")
    folder = Path(folder) if folder else lookup_dir()
    folder.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=folder, prefix=".upload-", suffix=ext)
    os.close(fd)
    tmp = Path(tmp)
    try:
        if hasattr(src, "save"):
            src.save(str(tmp))
        else:
            tmp.write_bytes(Path(src).read_bytes())
        try:
            stats = parse(tmp, kind)
        except ValueError as exc:
            raise ValueError(f"{LABEL[kind]}: {exc}") from None
        if not stats["map"]:
            raise ValueError(f"{LABEL[kind]}: the file has the right header but no data rows")
        target = folder / f"{STEM[kind]}{ext}"
        os.replace(tmp, target)
        for other in EXTS:
            if other != ext:
                old = folder / f"{STEM[kind]}{other}"
                if old.exists():
                    old.unlink()
        _memo.clear()
        return stats
    finally:
        if tmp.exists():
            tmp.unlink()


# ------------------------------------------------------------------ applying
def enrich(rows_out, rows_in, folder=None):
    """Add 'KYC Country', 'UID' to both Monthly Summary row lists and 'Total Volume' to out.

    Returns (summary dict for Verifikasi, set of labels to highlight yellow).
    """
    lk = load(folder)
    for d in rows_out:
        vol = d.get("Volume")
        if isinstance(vol, (int, float)) and str(d.get("Currency") or "").strip().upper() == "USC":
            d["Total Volume"] = vol / 100
        else:
            d["Total Volume"] = vol
    label = "UID / KYC lookup (KYC Country, UID columns)"
    if not lk["ok"]:
        for d in list(rows_out) + list(rows_in):
            d["KYC Country"] = None
            d["UID"] = None
        names = " and the ".join(f"{LABEL[k]} ({SHOW[k]})" for k in lk["missing"])
        verb = "are" if len(lk["missing"]) > 1 else "is"
        note = (f"NOT AVAILABLE: the {names} {verb} not stored on the server, so the KYC Country "
                "and UID columns are blank. An admin can upload it on the DPM Deal "
                "Segregator page (Update UID / KYC lookup files).")
        return {label: note}, {label}
    acc = lk["data"][KIND_ACCOUNT]["map"]
    kyc = lk["data"][KIND_KYC]["map"]
    logins, no_uid, no_kyc = set(), set(), set()
    for d in list(rows_out) + list(rows_in):
        login = norm_id(d.get("Login"))
        logins.add(login)
        uid = acc.get(login)
        if not uid:
            d["UID"], d["KYC Country"] = None, None
            no_uid.add(login)
            continue
        d["UID"] = _uid_value(uid)
        country = kyc.get(uid)
        d["KYC Country"] = country if country else None
        if not country:
            no_kyc.add(uid)
    matched = len(logins) - len(no_uid)
    summary = {
        label: (f"from the server: {LABEL[KIND_ACCOUNT]} {lk['paths'][KIND_ACCOUNT].name} "
                f"(updated {lk['dates'][KIND_ACCOUNT]}, {len(acc):,} accounts), "
                f"{LABEL[KIND_KYC]} {lk['paths'][KIND_KYC].name} "
                f"(updated {lk['dates'][KIND_KYC]}, {len(kyc):,} user_ids). "
                "Login -> account -> user_id = UID; UID -> Country of KYC. First match wins."),
        "Logins matched to a UID": f"{matched} of {len(logins)}",
        "Logins with no UID (KYC Country and UID left blank)":
            _list(no_uid) if no_uid else "(none)",
        "UIDs with no KYC row (KYC Country left blank)": _list(no_kyc) if no_kyc else "(none)",
        "Total Volume (Monthly Summary)": "Volume / 100 when Currency is USC, otherwise Volume.",
    }
    yellow = set()
    if no_uid:
        yellow.add("Logins with no UID (KYC Country and UID left blank)")
    if no_kyc:
        yellow.add("UIDs with no KYC row (KYC Country left blank)")
    return summary, yellow


def _list(items, limit=30):
    items = sorted(items, key=lambda s: (len(s), s))
    text = ", ".join(items[:limit])
    if len(items) > limit:
        text += f" (+{len(items) - limit} more)"
    return f"{len(items)}: {text}"
