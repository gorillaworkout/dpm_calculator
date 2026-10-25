"""Penentuan tanggal transaksi per baris -- DIPAKAI BERSAMA oleh isi_template.py
dan hitung_dw.py (dan lewat keduanya, alur KVB), supaya kedua tahap SELALU sepakat
baris mana masuk bulan laporan yang mana.

MASALAH YANG DISELESAIKAN (25 Okt 2026)
---------------------------------------
Sheet W memakai 'Completed Date' sebagai tanggal transaksi. Di ekspor baru,
~101 ribu baris 'Rebate Withdrawal' punya 'Completed Date' KOSONG (atau teks, atau
1970), padahal waktu selesainya ada di 'Completed Time' (angka serial Excel
beserta jam) dan 'Settlement Date' terisi. Dulu baris itu dibuang DIAM-DIAM oleh
saringan bulan -- ~7 ribu baris Rebate September hilang dari laporan.

URUTAN (hanya kalau kolom tanggal utamanya 'Completed Date'):
  1. 'Completed Date'   kalau berisi tanggal yang sah (1990-2100).
  2. 'Completed Time'   kalau berisi tanggal+jam (serial > 1000, datetime, atau
                        teks ISO) -> bagian tanggalnya.
  3. 'Settlement Date'  kalau sah. TIDAK dipakai kalau 'Completed Date'-nya
                        1970 (artinya "belum dibayar", lihat mtoatd_spec) --
                        baris itu tetap DILAPORKAN, bukan diselamatkan.
  'Apply Date' TIDAK PERNAH dipakai sebagai cadangan: itu tanggal PENGAJUAN,
  selalu lebih awal dari penyelesaian, dan ekspor back office disaring pakai
  tanggal itu -- memakainya menarik baris bulan lalu masuk bulan laporan.
Kolom utama lain (mis. 'Paid Date' untuk D) TIDAK berubah: tanpa cadangan.

'Completed Date' yang SAH selalu menang. Baris Rebate yang 'Completed Date'-nya
sah tapi bulannya beda dari 'Completed Time' + 'Settlement Date' (yang sama-sama
sepakat) dihitung dan dilaporkan sebagai KONFLIK; kalau tim memutuskan waktu yang
harus menang, balik saja UTAMAKAN_WAKTU.
"""
import datetime
import re
from collections import Counter, namedtuple

# Di luar batas ini (mis. 1970-01-01 = sel tanggal 'nol') tanggal dianggap TIDAK SAH.
TAHUN_MIN, TAHUN_MAX = 1990, 2100
SERIAL_MIN, SERIAL_MAX = 32874, 73415          # 1990-01-01 .. 2100-12-31 (serial Excel)
SERIAL_WAKTU_MIN = 1000     # di bawah ini = jam saja (0.68) atau angka biasa, bukan tanggal+jam

KOLOM_WAKTU = "COMPLETED TIME"
KOLOM_COMPLETED = "COMPLETED DATE"
KOLOM_SETTLE = "SETTLEMENT DATE"
# kolom utama -> kolom cadangan, urut. Kolom utama lain: tanpa cadangan.
RANTAI_CADANGAN = {KOLOM_COMPLETED: (KOLOM_WAKTU, KOLOM_SETTLE)}

# False = 'Completed Date' yang sah menang atas Completed Time / Settlement Date.
# True  = kalau keduanya sepakat dan bulannya beda, pakai itu (BELUM dipilih).
UTAMAKAN_WAKTU = False

AMBANG_TANPA_TANGGAL = 0.01     # peringatan kalau > 1% baris sheet tanpa tanggal sah

SHEET_RINGKASAN = "Date Check"
KATA_STATUS = ("FINISH", "SUCCESS", "REFUSE", "PROCESS", "CANCEL", "CANCELLED", "FAILED",
               "PENDING", "REJECT", "REJECTED")

_EPOCH = datetime.datetime(1899, 12, 30)
_RE_ISO = re.compile(r"^\s*(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T]+\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?)?\s*$")


def _angka(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _serial_ke_date(v):
    try:
        return (_EPOCH + datetime.timedelta(days=float(v))).date()
    except (OverflowError, ValueError):
        return None


def _sah(d):
    return d is not None and TAHUN_MIN <= d.year <= TAHUN_MAX


def _mentah(v):
    """Tanggal apa adanya (BELUM dicek rentang tahunnya), atau None."""
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, datetime.time):
        return None
    if _angka(v):
        f = float(v)
        if f != f:                                   # NaN
            return None
        if f.is_integer():
            s = str(int(f))
            if len(s) == 8:                          # YYYYMMDD (sheet D Juli 2026)
                try:
                    return datetime.datetime.strptime(s, "%Y%m%d").date()
                except ValueError:
                    pass
        if f >= 1:                                   # < 1 = jam saja
            return _serial_ke_date(f)
        return None
    if isinstance(v, str):
        t = v.strip()
        if not t:
            return None
        m = _RE_ISO.match(t)
        if m:
            try:
                return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                return None
        for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%Y%m%d"):
            try:
                return datetime.datetime.strptime(t, fmt).date()
            except ValueError:
                pass
    return None


def ke_tanggal(v):
    """Tanggal SAH (1990-2100) atau None."""
    d = _mentah(v)
    return d if _sah(d) else None


def kosong(v):
    return v is None or (isinstance(v, str) and not v.strip())


def kelas_nilai(v):
    """'kosong' | 'sah' | 'tidak_sah' (tanggal terbaca tapi di luar 1990-2100, mis. 1970)
    | 'bukan_tanggal' (teks / angka yang bukan tanggal, mis. 'success')."""
    if kosong(v):
        return "kosong"
    d = _mentah(v)
    if d is not None:
        return "sah" if _sah(d) else "tidak_sah"
    return "bukan_tanggal"


def waktu_ke_tanggal(v):
    """Bagian TANGGAL dari 'Completed Time'. Hanya yang benar-benar tanggal+jam:
    datetime, serial Excel > 1000, atau teks ISO ('2026-09-03 10:15:00').
    Jam saja (datetime.time, 0.68) -> None. Hasil harus sah (1990-2100)."""
    if isinstance(v, datetime.datetime):
        d = v.date()
    elif isinstance(v, (datetime.time, bool)) or kosong(v):
        return None
    elif isinstance(v, datetime.date):
        d = v
    elif _angka(v):
        if not float(v) > SERIAL_WAKTU_MIN:
            return None
        d = _serial_ke_date(v)
    elif isinstance(v, str):
        t = v.strip()
        try:
            f = float(t)
        except ValueError:
            m = _RE_ISO.match(t)
            if not m:
                return None
            try:
                d = datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                return None
        else:
            if not f > SERIAL_WAKTU_MIN:
                return None
            d = _serial_ke_date(f)
    else:
        return None
    return d if _sah(d) else None


HasilTanggal = namedtuple(
    "HasilTanggal",
    "tanggal sumber alasan_utama konflik konflik_sumber")
# tanggal         : date sah, atau None (= tidak ada tanggal yang bisa dipakai)
# sumber          : nama kolom (HURUF BESAR) asal tanggal, atau None
# alasan_utama    : 'sah' | 'kosong' | 'tidak_sah' | 'bukan_tanggal' -- keadaan kolom UTAMA
# konflik         : date dari Completed Time / Settlement kalau berbeda BULAN dari
#                   'Completed Date' yang sah (lihat UTAMAKAN_WAKTU), selain itu None
# konflik_sumber  : kolom asal 'konflik'


def tentukan_tanggal(nilai, utama):
    """Tanggal transaksi SATU baris.

    nilai : {NAMA KOLOM HURUF BESAR: isi sel} -- hanya kolom yang ada di sheet.
    utama : nama kolom tanggal utama sheet itu (mis. 'COMPLETED DATE', 'PAID DATE').
    """
    mentah = nilai.get(utama)
    kelas = kelas_nilai(mentah)
    cadangan = RANTAI_CADANGAN.get(utama, ())
    d_utama = ke_tanggal(mentah) if kelas == "sah" else None

    # Pembanding (hanya relevan untuk 'Completed Date'): tanggal dari Completed Time
    # dan Settlement Date.
    d_waktu = waktu_ke_tanggal(nilai.get(KOLOM_WAKTU)) if KOLOM_WAKTU in cadangan else None
    d_settle = ke_tanggal(nilai.get(KOLOM_SETTLE)) if KOLOM_SETTLE in cadangan else None

    if d_utama is not None:
        konflik = konflik_sumber = None
        # Konflik hanya dinilai kalau Completed Time (penentu waktu selesai) sah, dan
        # Settlement -- kalau ada -- sebulan dengannya. Settlement sering sehari
        # sebelum Completed, jadi dia sendirian tidak cukup sebagai bukti.
        if (d_waktu is not None
                and (d_waktu.year, d_waktu.month) != (d_utama.year, d_utama.month)
                and (d_settle is None
                     or (d_settle.year, d_settle.month) == (d_waktu.year, d_waktu.month))):
            konflik, konflik_sumber = d_waktu, KOLOM_WAKTU
        if konflik is not None and UTAMAKAN_WAKTU:
            return HasilTanggal(konflik, konflik_sumber, "sah", konflik, konflik_sumber)
        return HasilTanggal(d_utama, utama, "sah", konflik, konflik_sumber)

    # kolom utama tidak bisa dipakai -> cadangan
    if KOLOM_WAKTU in cadangan and d_waktu is not None:
        return HasilTanggal(d_waktu, KOLOM_WAKTU, kelas, None, None)
    # Completed Date 1970 = "belum dibayar": tidak diselamatkan lewat Settlement.
    if KOLOM_SETTLE in cadangan and kelas != "tidak_sah" and d_settle is not None:
        return HasilTanggal(d_settle, KOLOM_SETTLE, kelas, None, None)
    return HasilTanggal(None, None, kelas, None, None)


NAMA_ALASAN = {
    "kosong": "blank",
    "tidak_sah": f"outside {TAHUN_MIN}-{TAHUN_MAX} (e.g. 1970)",
    "bukan_tanggal": "not a date (text)",
}


class Ringkasan:
    """Hitungan per sheet: dipakai, dilewati (per alasan / bulan), cadangan, konflik."""

    def __init__(self, sheet):
        self.sheet = sheet
        self.total = 0
        self.dipakai = 0
        self.luar_bulan = Counter()      # 'YYYY-MM' -> baris
        self.tanpa_tanggal = Counter()   # alasan_utama -> baris
        self.cadangan = Counter()        # kolom cadangan -> baris yang tanggalnya dari situ
        self.konflik = Counter()         # (bulan Completed Date, bulan Completed Time) -> baris
        self.konflik_dipakai = 0
        self.status_geser = 0            # baris dengan kata status di 'Pay Channel'

    def catat(self, hasil, utama, masuk):
        """masuk: True/False = baris dipakai / dilewati (atau None kalau tak ada saringan)."""
        self.total += 1
        if hasil.tanggal is None:
            self.tanpa_tanggal[hasil.alasan_utama] += 1
        else:
            if hasil.sumber != utama:
                self.cadangan[hasil.sumber] += 1
            if hasil.konflik is not None:
                self.konflik[(hasil.tanggal.strftime("%Y-%m") if hasil.sumber == utama
                              else "(replaced)", hasil.konflik.strftime("%Y-%m"))] += 1
                if hasil.sumber != utama:
                    self.konflik_dipakai += 1
            if masuk is False:
                self.luar_bulan[hasil.tanggal.strftime("%Y-%m")] += 1
        if masuk:
            self.dipakai += 1

    @property
    def n_tanpa(self):
        return sum(self.tanpa_tanggal.values())

    def lewat_ambang(self):
        return self.total > 0 and self.n_tanpa / self.total > AMBANG_TANPA_TANGGAL

    def baris_ringkasan(self, saring_label=None):
        """[(sheet, kategori, rincian, jumlah)] untuk sheet 'Date Check' / Missing Data."""
        s = self.sheet
        out = [(s, "rows read", "rows with a currency", self.total)]
        out.append((s, "rows used", "", self.dipakai))
        for bln, n in sorted(self.luar_bulan.items()):
            out.append((s, "skipped: outside report month", bln, n))
        for alasan, n in sorted(self.tanpa_tanggal.items()):
            out.append((s, "skipped: no usable date", NAMA_ALASAN.get(alasan, alasan), n))
        for kol, n in sorted(self.cadangan.items()):
            out.append((s, "date taken from fallback column", kol.title(), n))
        for (b_utama, b_waktu), n in sorted(self.konflik.items()):
            out.append((s, "Completed Date disagrees with Completed Time + Settlement Date",
                        f"Completed Date month {b_utama} / Completed Time month {b_waktu}"
                        + (" -> Completed Time used" if b_utama == "(replaced)"
                           else " -> Completed Date kept"), n))
        if self.status_geser:
            out.append((s, "shifted columns", "status word found under 'Pay Channel'",
                        self.status_geser))
        if self.lewat_ambang():
            pct = self.n_tanpa / self.total * 100
            out.append((s, "WARNING", f"{pct:.1f}% of rows have no usable date "
                                      f"(limit {AMBANG_TANPA_TANGGAL * 100:.0f}%)", self.n_tanpa))
        return out
