# Dupoin DPM Tools

D&W Calculator plus Deal Segregator for the DPM Malaysia team.

GitHub: https://github.com/gorillaworkout/dpm_calculator

## D&W Calculator — auto hitung Handling Fee, Xero Rate, Xero USD, Forex Gain/Loss

Jalan di **macOS maupun Windows**. Yang menghitung adalah `hitung_dw.py`
(Python, lintas-platform). File launcher-nya beda per OS:

| File | OS | Cara pakai |
|---|---|---|
| `Hitung DW.app` | macOS | double-click → pilih file Excel |
| `Hitung DW.bat` | Windows | drag file Excel ke file .bat, atau double-click lalu drag |
| `Hitung DW.command` | macOS (cadangan) | klik kanan → Open With → Terminal |

Paket ini **khusus macOS** — hanya `Hitung DW.app` (+ `Hitung DW.command` sebagai
cadangan). Semua file `.py`-nya sendiri lintas-platform dan tidak bergantung macOS,
jadi untuk dipakai di server atau website cukup panggil `hitung_dw.py` langsung.
| `hitung_dw.py` | dua-duanya | lewat Terminal / Command Prompt |

---

## Setup Windows (sekali saja)

1. Download Python: <https://www.python.org/downloads/windows/>
2. Saat install, **CENTANG "Add python.exe to PATH"** (ini yang paling sering kelupaan).
3. Buka **Command Prompt**, jalankan:
   ```
   py -m pip install openpyxl
   ```
   **Pakai `py -m pip`, jangan `pip` saja.** Kalau `pip install openpyxl` menghasilkan
   `'pip' is not recognized as an internal or external command`, itu bukan berarti Python
   gagal terinstall — folder `Scripts` milik Python saja yang belum masuk PATH.
   `py -m pip` tidak butuh itu karena `py` dipasang langsung ke `C:\Windows`.
   Kalau `py` juga tidak dikenali: tutup semua Command Prompt dan buka yang baru
   (PATH tidak ter-refresh di jendela yang dibuka sebelum install).

   Langkah ini boleh dilewati — `Hitung DW.bat` otomatis memeriksa openpyxl
   dan menawarkan menginstallnya.
4. Copy seluruh folder `dw-calculator` ke Windows, mis. ke `C:\dw-calculator`.
5. Double-click `Hitung DW.bat`.

Kalau muncul `Python belum terpasang / belum masuk PATH`, ulangi langkah 2
(install ulang Python dan centang kotak PATH-nya).

## Setup macOS (sekali saja)
```bash
python3 -m pip install --user openpyxl
```
Lalu double-click `Hitung DW.app`.

---

## Lewat Terminal / Command Prompt

macOS:
```bash
cd ~/Documents/Dupoin/dw-calculator
python3 hitung_dw.py ~/Downloads/"DPM-D&W.xlsx"
```

Windows:
```
cd C:\dw-calculator
py hitung_dw.py "C:\Users\NAMAMU\Downloads\DPM-D&W.xlsx"
```

Nama file wajib di dalam tanda kutip karena ada karakter `&`.
Hasil: `DPM-D&W-hasil.xlsx` di folder yang sama dengan file input. File asli tidak diubah.

---

## Bulan laporan — `--period` / dropdown di web app

Ekspor back office **selalu punya ekor**. File "Juni" yang diunduh dari back office
biasanya berisi **25 Mei s/d 5 Juli**: ekspornya disaring pakai *Apply Date*,
sementara laporan ini dibangun di atas *Paid Date* (deposit) dan *Completed Date*
(withdrawal). Kalau ekor itu ikut, `MTOATD` keluar 42 blok tanggal bukan 30, dan
totalnya memuat uang bulan lain.

Di web app ada dropdown **Report month** yang **wajib** diisi (default: bulan lalu).
Lewat Terminal:

```bash
python3 isi_template.py "sumber.xlsx" --bulan 2026-06 -o kerja.xlsx
python3 hitung_dw.py    "kerja.xlsx"  --period 2026-06
```

Basis tanggalnya sama dengan yang dipakai seluruh laporan — kolom `Date` per baris,
yaitu `Paid Date` untuk sheet `D` dan `Completed Date` untuk sheet `W` — jadi satu
baris masuk bulan yang sama di `D&W Report`, `MTOATD` dan `Payment Channel Balance`.
Yang ikut disaring: sheet **D**, sheet **W**, dan **Fund Transfer Table**
(J Wallet + kolom *Fund transfer* di Payment Channel Balance).

Baris yang dibuang **tetap ada** di sheet D/W, hanya 4 kolom hitungannya dikosongkan
dan diarsir abu-abu — filter `(Blanks)` untuk melihatnya. Rinciannya di sheet
`Missing Data` bagian **0**, aturannya di sheet `Legend` bagian **REPORT MONTH**.

Kalau tidak ada satu pun baris yang cocok, tool **berhenti** dan menyebut bulan apa
saja yang sebenarnya ada di file itu. Sengaja gagal: workbook kosong yang kelihatan
normal lebih berbahaya daripada pesan error.

Tanpa `--period`, semua tanggal dihitung (perilaku lama) dan itu ditulis merah di
sheet `Legend`.

Terbukti pada file uji 38 MB (`June - ..._Test 1.xlsx`, 25 Mei – 5 Jul):

| | tanpa `--period` | `--period 2026-06` |
|---|---|---|
| baris dihitung | 168.732 | 128.210 |
| blok tanggal MTOATD | 42 (25 Mei – 5 Jul) | **30 (Juni saja)** |
| baris dibuang | – | 40.901 (D: 17.805 Mei + 7.177 Jul; W: 11.114 Mei + 4.804 Jul + 1 tanpa tanggal) |

---

## Yang dihitung

4 kolom baru di kanan sheet `D` (dan `W` kalau ada):

| Kolom | Rumus | Format |
|---|---|---|
| Handling Fee | `Transaction × rate` dari sheet `D&W FEE`, dicocokkan Currency + Payment Gateway. Deposit pakai kolom Deposit, withdrawal pakai kolom Withdrawal (dibaca dari kolom `Source Name` tiap baris) | `#,##0.##` → `8,100`, `0.18` |
| Xero Rate | kurs di sheet `XERO` pada `Paid Date` + `Currency`. Kuning = tanggal belum ada di XERO, dipakai kurs tanggal terdekat sebelumnya | `#,##0.######` → `26,317.2`, `61.5507`, `1` |
| Xero USD | `Transaction ÷ Xero Rate` | `#,##0.00` |
| Forex Gain/Loss | deposit: `Xero USD − USD` · withdrawal: `USD − Xero USD` (tanda dibalik karena uang keluar — atur di `FLIP_FOREX_ON_WITHDRAWAL`) | `#,##0.00` |

### Sheet yang dihasilkan

| Sheet | Isi |
|---|---|
| `D` / `W` (dan sheet transaksi lain) | data aslimu + 4 kolom hitungan, ditandai warna. Kalau file sudah punya kolom `Handling Fee` / `Xero Rate` / `Xero USD` / `Forex Gain/-Loss` / `Currency Gain/Loss`, **kolom itu yang diisi** — judulnya tidak diubah, jadi tampilan file tetap sama |
| **`D&W Report`** | **ringkasan Currency × Payment Gateway, layout sama dengan report D&W yang sudah beredar** — bagian Deposit di atas, Withdrawal di bawah, masing-masing dengan Grand Total |
| `D&W Detail` | tabel datar gabungan semua sheet transaksi, satu baris per transaksi + kolom audit |
| `Missing Data` | rincian data yang belum ada: tanggal kurs yang tidak ada di `XERO` (per currency, per tanggal, berapa baris, kurs tanggal berapa yang dipakai, selisih berapa hari), rate fee yang belum ada, masalah data lain |
| `Legend` | penjelasan kolom & arti warna (English) |

Keempat sheet hasil **ditulis ulang tiap run** — jangan diedit manual.
Matikan semuanya dengan `--no-report`.

#### Sheet `D&W Report` — rumusnya

Dikelompokkan per `Currency` → `Payment Gateway`, memakai nilai **apa adanya**
(spasi di ujung nama gateway tidak dibuang, supaya `TRC20` dan `TRC20 ` tetap
terpisah seperti di pivot Excel).

**Bagian Deposit**

| Kolom | Rumus |
|---|---|
| `CRM USD ` | jumlah `USD` |
| `TRANSACTIONS` | jumlah `Transaction` |
| `Xero USD ` | jumlah `Xero USD` |
| `Total FE Gain/-Loss (USD)` | `Xero USD − CRM USD` |
| `CRM Average Rate` | `TRANSACTIONS ÷ CRM USD` |
| `Xero average Rate` | `TRANSACTIONS ÷ Xero USD` |
| `Recorded in Xero` | dikosongkan — diisi manual |
| `Rounding Adjustment` | `−Xero USD` |

**Bagian Withdrawal** (ada satu kolom tambahan, `Handling Charges Income`)

| Kolom | Rumus |
|---|---|
| `CRM USD ` | jumlah `USD` |
| `Handling Charges Income` | jumlah kolom `Charges` dari sheet `W` |
| `Transactions` | jumlah `Transaction` |
| `Xero USD ` | jumlah `Xero USD` |
| `Total FE Gain/-Loss (USD)` | `CRM USD − Xero USD` (tanda berlawanan dengan deposit) |
| `CRM average` | `Transactions ÷ CRM USD` |
| `Xero average` | `Transactions ÷ Xero USD` |
| `Recorded in Xero` | dikosongkan — diisi manual |
| `Rounding Adjustment` | `+Xero USD` |

Judul, label, warna header, fill, format akuntansi, dan lebar kolom disamakan
dengan report aslinya. Periode di bawah judul bagian diambil dari bulan yang
paling banyak muncul di data — atur manual dengan `--period 2026-06`.

**Urutan grup** tidak bisa direproduksi dari report lama (urutannya tersimpan di
cache pivot Excel). Atur dengan `REPORT_GROUP_ORDER` di config:
`"value"` (default, CRM USD terbesar dulu), `"appearance"` (urutan kemunculan di
sheet), atau `"alpha"` (A–Z).

### Sheet `Legend` di file hasil
Tiap run, file hasil dapat sheet **`Legend`** berisi penjelasan keempat kolom
dan arti setiap warna, **dalam bahasa Inggris**, lengkap dengan contoh sel berwarna.
Sheet ini ditulis ulang setiap run, jadi jangan diedit manual.
Penjelasan yang sama juga terpasang sebagai komentar di sel header kolomnya
(arahkan kursor ke header `Handling Fee` / `Xero Rate` / dst).

### Arti warna sel
| Warna | Kolom | Arti |
|---|---|---|
| **Merah** `Handling Fee` | rate Currency + Payment Gateway belum ada di tabel fee. Sel sengaja **dikosongkan, bukan 0**, supaya tidak terhitung diam-diam. Tambahkan barisnya di sheet `D&W FEE` lalu hitung ulang |
| **Merah** `Xero Rate` / `Xero USD` / `Forex` | kurs tidak ditemukan sama sekali untuk tanggal + currency itu |
| **Kuning** `Xero Rate` | tanggalnya belum ada di sheet `XERO`, dipakai kurs tanggal terdekat sebelumnya. Lengkapi sheet XERO untuk menghilangkannya |

Keterangan ini juga terpasang sebagai komentar di sel header masing-masing kolom di Excel.

Baris `<CURRENCY>/Other` di tabel fee **hanya** dipakai kalau nama gateway di data
memang literal `Other`. Gateway lain yang tidak ketemu dibiarkan merah — jadi tidak
ada lagi rate yang diam-diam dianggap 0%.

### Kolom tanggal untuk cari kurs
Dicari berurutan sesuai `DATE_COLUMNS`: `Paid Date` → `Completed Date` → `Settlement Date` → `Apply Date`.
Sheet `D` punya `Paid Date` → dipakai. Sheet `W` tidak punya → turun ke **`Completed Date`**.
Kolom yang benar-benar dipakai selalu dicetak di laporan tiap run.

### Nama gateway yang otomatis dikenali
- suffix currency dibuang: `PA-PHP` → `PA`, `BeckPay-VND` → `BeckPay`, `MNTX-UZS` → `MNTX`
- tanda `（Manual）` diabaikan: `PA-PHP（Manual）` → `PA`
- alias manual di `GATEWAY_ALIAS`: `VN77PAY` → `77 Pay`, `BEP20`/`USDT-BEP20` → `Other`, `MNTX-BinancePay` → `Binance Pay`

---

## File lambat? Rapikan dulu — `rapikan_file.py`

```bash
python3 rapikan_file.py ~/Downloads/"DPM-D&W.xlsx" --buang-pivot
```

Dua sumber pemborosan yang dibersihkan, keduanya tidak kelihatan dari Excel:

| Penyebab | Contoh nyata di `DPM-D&W.xlsx` |
|---|---|
| **Baris hantu** — baris kosong yang pernah diformat | sheet `D` tercatat 918.498 baris, isinya 182. `sheet2.xml` = 178 MB |
| **Pivot cache** — salinan data sumber PivotTable | 28,7 MB. Menyumbang **99,8%** waktu buka |

Hasil terukur:

| | Sebelum | Sesudah |
|---|---|---|
| ukuran file | 20,7 MB | **0,2 MB** |
| waktu buka (openpyxl) | 313 detik | **0,2 detik** |
| waktu hitung penuh | 145 detik | **1 detik** |

Merapikannya sendiri hanya **2,3 detik** — dikerjakan langsung di level XML,
jadi nilai angkanya **identik byte-per-byte dengan aslinya** (sudah diuji
sel per sel; justru lebih setia daripada rute openpyxl yang membulatkan
`22.400000000000002` menjadi `22.4`).

File asli tidak diubah; hasilnya `<nama>-rapi.xlsx`.

`--buang-pivot` membuat PivotTable lama tidak bisa di-refresh lagi (angka
terakhirnya tetap ada sebagai angka biasa). Aman kalau report sudah
di-generate tool ini. Tanpa opsi itu file tetap lambat.

## Berapa lama prosesnya

| File | Waktu |
|---|---|
| di bawah 5 MB | beberapa detik |
| 47 MB / 45.000 baris | **2–3 menit** |

Sebagian besar waktu bukan untuk menghitung, tapi untuk I/O Excel:

```
membuka file 47 MB  ~ 75 detik
menghitung          ~ 10 detik
menyimpan           ~ 60 detik
```

Progress dicetak selama proses (`[m:ss]` + persentase per sheet), jadi
kalau masih ada tulisan berjalan berarti belum stuck. `Hitung DW.app`
menjalankannya di jendela Terminal supaya progress ini terlihat —
tanpa itu aplikasinya tampak diam selama ±75 detik pertama.

Opsi `--open` membuka file hasil otomatis setelah selesai.

## Menambahkan rate ke sheet `D&W FEE` lewat script

Kalau tidak mau mengedit Excel manual:
```bash
python3 tambah_rate.py "DPM-D&W.xlsx" VND BeckPay 1.6% 0%
python3 tambah_rate.py "DPM-D&W.xlsx" THB Other 0 0
```
Bikin backup otomatis dulu (`DPM-D&W-backup.xlsx`), lalu menambah baris baru --
atau meng-update kalau kombinasi Currency + Payment Gateway itu sudah ada.
File harus ditutup dari Excel.

## MTOATD — sekali jalan, tanpa input manual

MTOATD adalah rekonsiliasi harian antara **empat sistem** — MT4, Wallet, CRM, dan TD.
Dulu baris MT4/Wallet/CRM dianggap harus diisi manusia dan alurnya dua tahap. **Itu
keliru.** Workbook `22.08 Bayu - D&W-Dupoin Markets-JUN 2026v3.xlsx` menyimpan rumusnya
di sheet `MTOATD (MAY'26)`: semuanya `SUMIFS` ke sheet `Deposits` / `Withdrawals`, yang
isi dan susunan kolomnya **identik** dengan sheet `D` / `W`.

Jadi sekarang satu perintah saja:

```
python3 hitung_dw.py "D&W JUN 2026.xlsx"
   -> D&W Report, D&W Detail, D&W FEE, Channel Balance, J Wallet (calc),
      MTOATD, Missing Data, Legend
```

`hitung_mtoatd.py` sudah dihapus — tidak ada tahap kedua lagi.

### Rumusnya

Semua ada di `mtoatd_spec.py`. Per tanggal `t`, per currency:

```
MT4入金            = Σ USD          [Settlement Date = t]
Wallet入金          = (tidak ada rumusnya -> KOSONG)
CRM入金            = MT4入金 + Wallet入金
TD入金             = Σ USD          [Paid Date = t]
——crm已入,TD未入    = Σ USD          [Settlement = t, Paid > t]
——调上日差异 (入金)   = -Σ USD         [Paid = t, Settlement < t]
MT4出金            = Σ USD          [Settlement = t, Source Name = 'Withdrawal']
Wallet出金          = Σ USD          [Settlement = t, Source Name = 'Rebate Withdrawal']
CRM出金            = MT4出金 + Wallet出金
TD出金             = Σ USD          [Completed = t, Status ≠ refuse]
——crm已出,TD未出    = Σ USD          [Settlement = t, Completed > t, Status ≠ refuse]
                    + Σ USD          [Settlement = t, Completed = 1/1/1970]
——调上日差异 (出金)   = -Σ USD         [Completed = t, Settlement < t, Status ≠ refuse]
CRM 入金（原币种）   = Σ Transaction  [Settlement = t]
CRM 出金（原币种）   = Σ OC           [Settlement = t]
实收（原币种）       = Σ Transaction  [Paid = t]
实出（原币种）       = Σ Transaction  [Completed = t]
——未打款            = Σ OC           [Settlement = t, Completed > t, Status ≠ refuse]
                    + Σ OC           [Settlement = t, Completed = 1/1/1970]
——手续费/转账费      = Σ Charges      [Completed = t]     (hanya USDT & USD)
——调上日差异 (原币种) = -Σ Transaction [Completed = t, Settlement < t]
本日 MT4+錢包淨入金  = MT4入金 + Wallet入金 - MT4出金 - Wallet出金
月累计              = akumulasi 本日 sepanjang bulan
合計                = jumlah antar currency
```

`OC` = *original currency*: untuk `USDT` dan `USD` mata uang aslinya memang USD, jadi
rumus mereka memakai kolom **USD**; currency lain memakai kolom **Transaction**.

Basis tanggalnya penting: MT4/CRM pakai **Settlement Date**, TD入金/实收 pakai
**Paid Date**, TD出金/实出 pakai **Completed Date**.

### Baris yang dibiarkan kosong

Baris `——...` yang di workbook sumber **tidak punya rumus** (`——作废单` voided,
`——小数点差异` rounding, `——MYR改USDT出`, `——分拆出金`, dan `Wallet入金`) ditulis
**KOSONG** dan diberi warna abu-abu — bukan nol. Itu memang rincian yang diketik
manusia kalau perlu; tidak ada yang dihitung ulang dari sana.

### Blok 当月累计

Di sebelah kanan kolom `合計` ada blok kedua dengan susunan baris yang sama, berisi
akumulasi sejak awal bulan sampai tanggal itu. Matikan dengan
`MTOATD_BLOK_AKUMULASI = False`.

### Tiga kejanggalan yang ditiru

Di workbook sumber ada tiga baris yang rumusnya diedit **hanya di kolom USDT / USD**:

| Baris | Kejanggalan |
|---|---|
| `实出（原币种）` | hanya kolom USDT yang menyaring `Status ≠ refuse` |
| `——手续费/转账费` | hanya ada di kolom USDT & USD; USDT menyaring refuse, USD tidak |
| `——调上日差异` (原币种) | USDT menyaring `finish` + `≠ refuse`; USD pakai kolom USD; VND menunjuk kolom `AK` yang kosong (jadi tanpa saringan) |

Semuanya ditiru apa adanya supaya angkanya nyambung dengan laporan yang sudah beredar.
Setel `QUIRKS_ASLI = False` di `mtoatd_spec.py` untuk memakai satu aturan seragam.

### Verifikasi

Dibandingkan sel per sel dengan sheet `MTOATD (MAY'26)` tanggal 1 Jun 2026:
**523 dari 523 sel cocok, 0 beda** (19 currency + kolom `合計`, semua baris yang
punya rumus). Baris tanpa rumus di sheet mereka bernilai 0 atau kosong — di sini
memang dikosongkan.

Baris tanpa `Payment Gateway` **tidak** dikeluarkan: rumus asli mereka juga tidak
menyaring gateway, jadi angkanya cocok apa adanya.

## Template — cara paling mudah

`Template D&W.xlsx` adalah file kerja bersih: tempel data, jalankan, selesai.

| Sheet | Isi |
|---|---|
| `Guide` | penjelasan singkat (English) — apa ditempel di mana |
| `D` | header deposit, kosong. Tempel data mulai baris 2 |
| `W` | header withdrawal, kosong. Tempel data mulai baris 2 |
| `Xero` | kurs harian, sudah terisi (759 tanggal) |
| `Fund Transfer Table` | pemindahan dana antar channel — sumber `Channel Balance` & `J Wallet (calc)` |
| `Opening Balance` | saldo penutup hari terakhir bulan sebelumnya, per currency + channel |
| `D&W FEE` | **satu** tabel fee, sudah tergabung dari `D&W TD FEE` + `D&W TD FEE -add.` |

Empat kolom ungu (`Handling Fee`, `Xero Rate`, `Xero USD`, `Forex Gain/Loss`) di sheet
`D` dan `W` dibiarkan kosong — kalkulator yang mengisinya.

Bikin ulang template dari workbook mana pun:
```bash
python3 buat_template.py "workbook-sumber.xlsx" -o "Template D&W.xlsx"
```

### Mengisi template dari workbook lain — `isi_template.py`

```bash
python3 isi_template.py "21.08 Bayu - D&W....xlsx" --bulan 2026-06 -o "D&W JUN 2026.xlsx"
```

Menyalin baris dari sheet `D` dan `W` di workbook sumber ke template.
**Kolom dicocokkan berdasarkan nama header**, bukan posisi — jadi urutan kolom di
sumber boleh berbeda, dan kolom hitungan di sumber diabaikan (kalkulator yang
mengisinya). Isi lama template dibersihkan lebih dulu.

Selain `D` dan `W`, dua sheet lagi ikut diisi supaya **sekali jalan langsung lengkap**:

- **`Fund Transfer Table`** — disalin dari sheet dengan nama yang sama di sumber.
  Sheet itu punya **dua sisi bersebelahan dengan nama kolom identik** (`金额`, `手续费`,
  `Xero (USD)` muncul dua kali), jadi kolom dicocokkan **per sisi** dengan batas kolom
  `收款日期`. Kalau dicocokkan pakai nama saja, kunci yang kembar saling menimpa.
- **`Opening Balance`** — saldo penutup hari terakhir sebelum data ini, dibaca dari
  sheet `J Wallet` dan `Payment Channel Balance` di workbook sumber. Tanpa ini saldo di
  `Channel Balance` dan `J Wallet (calc)` mulai dari nol. Basis tanggalnya beda:
  channel pakai tanggal transaksi paling awal (`Paid` / `Completed Date`), J Wallet
  pakai tanggal `Fund Transfer Table` paling awal.

Nama sheet dicoba berkelompok dari yang paling spesifik: `D` → `Deposit` → `Deposits`
(dan `W` → `Withdrawal`/`WD` → `Withdrawals`), jadi sheet mentah menang atas sheet
yang sudah berisi hasil hitungan. Tentukan manual dengan `--sheet-d` / `--sheet-w`.

`--bulan YYYY-MM` menyaring berdasarkan kolom tanggal (`Paid Date` untuk deposit,
`Completed Date` untuk withdrawal).

### Kenapa pakai template, bukan workbook penuh

Terukur di file `21.08 Bayu - D&W-Dupoin Markets-JUN 2026v3(1).xlsx` (20 MB, 21 sheet):

| | Workbook penuh | Template |
|---|---|---|
| ukuran file kerja | 20 MB | **105 KB** |
| waktu satu putaran | **314 detik** | **1,5 detik** |
| PivotTable | **rusak** di file hasil — 4 sheet kehilangan koneksi pivot, dan file hasilnya bahkan tidak bisa dibuka ulang (referensi `pivotCacheRecords1.xml` menggantung) | tidak ada pivot yang bisa rusak |
| angka | acuan | **identik** — 12.000 nilai dibanding, 0 beda, Grand Total sama |

Workbook penuh tetap bisa dipakai kalau perlu, tapi simpan file sumbermu sebagai
master dan jangan pakai file hasilnya untuk apa pun selain dibaca.

## Tabel fee: digabung otomatis dari beberapa sheet

Semua sheet yang namanya dimulai `D&W FEE` / `D&W TD FEE` / `DW FEE` **digabung
otomatis** setiap kali script jalan. Sheet yang mengandung `-add` diterapkan
**terakhir**, jadi nilainya menimpa sheet dasar. Tidak perlu menyalin apa pun.

```
sheet fee 'D&W TD FEE '        : 35 baris -> 35 baru,  0 menimpa
sheet fee 'D&W TD FEE -add.'   : 58 baris -> 33 baru, 25 menimpa
Fee : 69 kombinasi currency+gateway
```

Baris header dan **urutan kolom dideteksi dari nama header**, jadi layout
`Currency | Payment Gateway | Deposit | Withdrawal` (header di baris 3) maupun
`No | Payment Gateway | Currency | Deposit | Withdrawal` (header di baris 2,
gateway di depan currency) dua-duanya kebaca.

Nama gateway dicocokkan setelah semua tanda baca dibuang, sehingga
`77 Pay` = `77Pay`, `1-2-PAY` = `1-2Pay`, `SHUNFA PAY` = `SHUNFAPAY`.
Tanpa ini tabel dasar dan tabel `-add.` tidak akan saling menimpa.

### Bentuk rate yang didukung

| Ditulis di sheet | Artinya |
|---|---|
| `0.027` | 2,7% |
| `1.5%+50` | 1,5% + biaya tetap 50 |
| `3.5%+6` | 3,5% + biaya tetap 6 |
| `+8` | 0% + biaya tetap 8 |
| `0.45%,min 90` | 0,45%, minimum 90 |

Rumusnya: `fee = max(Transaction × persen, minimum) + biaya_tetap`.
Biaya tetap dan minimum yang berasal dari rate **selalu** diterapkan.
Notasi lama `+10000` di kolom tambahan sheet dasar tetap digating
`INCLUDE_FIXED_FEE` (default `False`).

`gabung_fee.py` menghasilkan file tinjauan `D&W FEE gabungan.xlsx` berisi tabel
tergabung + sheet `Konflik` — untuk diperiksa manusia, bukan syarat perhitungan.

## Sheet mana yang dihitung

Nama sheet dicoba berkelompok, dari yang paling spesifik:
`("D","W")` → `("Deposit","Withdrawal","WD")` → `("Deposit Data","Withdrawal Data")`.
**Kelompok pertama yang ketemu dipakai, sisanya diabaikan.** Jadi di workbook yang
punya `D`, `W`, *dan* `Deposits`, `Withdrawals`, `Withdrawal`, `Withdrawal (2)`,
`Rebate Withdrawal`, hanya `D` + `W` yang dihitung — yang dilewati dilaporkan:

```
Lewati : 6 sheet transaksi lain tidak dihitung ('Deposits', 'Withdrawals', ...)
         tambahkan dengan --sheet "NAMA SHEET" kalau perlu
```

Kalau tidak ada nama yang dikenali, baru sheet dideteksi dari header kolomnya.

Sheet kurs dicari bernama `XERO` tanpa peduli huruf besar/kecil (`Xero` juga kebaca).

## Syarat file yang di-upload

Sheet transaksi **tidak perlu bernama `D` / `W`**. Skrip mendeteksi sheet transaksi
dari header-nya: sheet apa pun yang punya kolom `Currency`, `Transaction`,
`Payment Gateway`, `USD` + salah satu kolom tanggal (`Paid Date` / `Completed Date` /
`Settlement Date` / `Apply Date`) akan ikut dihitung.

Deposit vs withdrawal ditentukan **per baris** dari kolom `Source Name`
(`deposit` / `withdrawal` / `WD` / `Penarikan` / dst). Kalau kolom itu kosong,
barulah nama sheet dipakai: sheet bernama `W` / `WD` / `Withdrawal` dianggap
withdrawal, selain itu deposit.

Selain sheet transaksi, file harus punya:

| Sheet | Isi | Kalau tidak ada |
|---|---|---|
| `D&W FEE` | header di baris 3: `Currency \| Payment Gateway \| Deposit \| Withdrawal` | pakai `--fee-from "file-lain.xlsx"` |
| `XERO` | baris 1: `Date` + kode currency; baris berikutnya tanggal + kurs | pakai `--xero-from "file-lain.xlsx"` |

Kalau ada yang tidak sesuai, skrip berhenti dengan pesan yang menyebutkan
header apa yang dicari dan opsi apa yang harus dipakai — bukan traceback.

## Menambah data

Skrip membaca semua baris sampai baris terakhir yang ada isinya — tanpa batas.
Sudah dites 21.000 baris: ±7 detik.

**Cara 1 — tempel ke sheet `D` file yang sama**
1. Paste baris baru di bawah baris terakhir sheet `D` (nama header di baris 1 harus tetap:
   `Currency`, `Transaction`, `Payment Gateway`, `Paid Date`, `USD`, `Source Name`).
2. Kalau masuk bulan baru, tambah baris tanggal baru di sheet `XERO`.
3. Save → **tutup Excel** → jalankan lagi.

Dijalankan ulang di file yang sudah punya 4 kolom hasil: kolom itu **ditimpa**, tidak menumpuk.

**Cara 2 — file export terpisah yang belum punya sheet `D&W FEE` / `XERO`**
```bash
python3 hitung_dw.py export-juli.xlsx --fee-from DPM-D&W.xlsx --xero-from DPM-D&W.xlsx
```

---

## Opsi

| Opsi | Guna |
|---|---|
| `-o nama.xlsx` | tentukan nama file hasil |
| `--in-place` | timpa file input (harus ditutup dari Excel) |
| `--sheet D W` | pilih sheet data manual |
| `--fee-from FILE` | ambil tabel fee dari workbook lain |
| `--xero-from FILE` | ambil kurs dari workbook lain |
| `--period YYYY-MM` | **bulan laporan** — baris di luar bulan itu dibuang |
| `--jwallet-opening ANGKA` | saldo penutup J Wallet bulan sebelumnya (titik mulai saldo) |

## Kalau ada gateway atau rate baru

Edit bagian `KONFIGURASI` di atas `hitung_dw.py` (pakai editor teks apa saja):

- `EXTRA_FEES` — rate yang belum ada di sheet `D&W FEE`: `("VND","NOVOLINK"): (deposit, withdrawal)`
- `GATEWAY_ALIAS` — kalau nama gateway di data beda dari tabel fee, mis. `VN77PAY` → `77 PAY`
- `INCLUDE_FIXED_FEE` — `True` untuk ikut menambah fixed fee (mis. 10.000 VND/trx untuk 77 Pay)
- `XERO_RATE_MULTIPLIER` — `True` kalau mau Xero Rate disimpan 1/kurs dan Xero USD = Transaction × Xero Rate
- `MISSING_DATE_MODE` — `"nearest"` (default) atau `"exact"`
- `FMT_FEE` / `FMT_RATE` / `FMT_MONEY` — format tampilan angka di Excel

Tiap run, skrip mencetak tabel rate yang dipakai. Cek kolom `MATCH`:
`exact` / `alias` / `strip-suffix` = aman. `fallback-other` atau blok
`!! GATEWAY TANPA RATE` = gateway itu belum ada rate-nya, perlu ditambah ke `EXTRA_FEES`.

## Rate yang sudah masuk lewat EXTRA_FEES

| Currency + Gateway | Deposit | Withdrawal |
|---|---|---|
| `VND + Novolink` | 0.9% | 0% |
| `USD + BuziPay` | 0% | 0% |
| `THB + VPay` | 2.5% | 0.2% |
| `THB + NEPay` | 2.7% | 0.5% |
| `THB + Berry` | 2.4% | 0% |
| `UZS + MNTX` | 5.5% | 3.5% |
| `USD + MNTX-LAK USD` | (dari sheet) | 0% |
| `USD + Pay247-KH USD` | (dari sheet) | 2% |
| `VND + BeckPay` | 1.6% | 0% |

`EXTRA_FEES` sifatnya **menambal**, bukan menimpa: nilai yang sudah ada di sheet
`D&W FEE` selalu menang, dan `None` di `EXTRA_FEES` berarti "jangan sentuh".
Jadi `MNTX-LAK USD` tetap pakai deposit 3% dari sheet, cuma kolom Withdrawal-nya
yang kosong ditambal dari sini. Sudah diuji.

## Yang masih belum dikonfirmasi

| Currency + Gateway | Catatan |
|---|---|
| `PHP + Other` | belum ada baris `PHP/Other` di tabel fee -> Handling Fee kosong + merah (1 baris) |
| `THB + Other` | belum ada baris `THB/Other` di tabel fee -> Handling Fee kosong + merah (2 baris) |

Fixed fee 10.000 VND untuk 77 Pay belum dihitung (`INCLUDE_FIXED_FEE = False`).

## Withdrawal — sudah dites, siap dipakai

Skrip menentukan deposit/withdrawal **per baris** dari kolom `Source Name`
(`deposit` -> pakai rate kolom Deposit, `withdrawal` -> kolom Withdrawal).
Dua cara ini dua-duanya sudah diuji dan hasilnya benar:

1. Baris withdrawal **dicampur di sheet `D`** bersama deposit.
2. **Sheet terpisah bernama `W`** -> otomatis terdeteksi, tidak perlu opsi tambahan.

Kalau nama sheet-nya bukan `D`/`W` (mis. `Withdrawal`), pakai:
```bash
python3 hitung_dw.py file.xlsx --sheet D Withdrawal
```
