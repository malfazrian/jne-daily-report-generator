# Bot Report Gabungan

Automasi generate laporan harian (Excel) dari data shipment, termasuk:
- Ambil data source (folder bulanan/Parquet)
- Transformasi kolom + normalisasi tanggal
- Pembuatan sheet dan pivot

Repo ini dipakai untuk workflow report operasional dengan struktur data dan path jaringan internal.

## Ringkasan Alur

1. Load daftar project dari `list.py` (`criteria_lists`).
2. Enrich `criteria` dengan category dan id_account terbaru dari referensi.
3. Ambil data master per category:
   - Mode legacy: baca file bulanan (`main.py` + `master_data_reader.py`)
   - Mode parquet (direkomendasikan): query Parquet dengan DuckDB (`main_pq.py` + `master_data_pq_reader.py`)
4. Proses edit file (split sheet, styling, pivot, dan penyesuaian lain).
5. Group output per contact frontline.
6. (Opsional) kirim file via WhatsApp Web.

## Struktur Folder Utama

```text
.
|- main.py
|- main_pq.py
|- list.py
|- data_loader/
|- data_transform/
|- data_submitter/
|- utils/
|- data/
```

Keterangan cepat:
- `main.py`: entrypoint legacy (reader non-parquet).
- `main_pq.py`: entrypoint utama (reader parquet, paralel, lebih cepat).
- `list.py`: konfigurasi report per project (`criteria_lists`).
- `data_loader/`: logic pengambilan dan pembentukan master data.
- `data_transform/`: logic edit workbook, split sheet, pivot, transform kolom.
- `data_submitter/send_wa.py`: kirim pesan + lampiran ke WhatsApp Web.
- `utils/helper.py`: history processing, lock file, rule `should_send`.
- `utils/utils.py`: helper umum, enrich category/id, update referensi project.

## Prasyarat

- OS Windows.
- Python 3.10+ (disarankan 3.10/3.11).
- Google Chrome terpasang.
- Akses ke path jaringan internal (UNC path) yang dipakai pada konfigurasi.
- Microsoft Excel (untuk fitur pivot berbasis COM di beberapa alur).

## Setup Environment

### 1. Buat virtual environment

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 2. Install dependency

Install dependency dari file requirements:

```powershell
pip install -r requirements.txt
```

### 3. Siapkan konfigurasi environment

Copy template konfigurasi lalu sesuaikan path sesuai mesin kamu:

```powershell
Copy-Item .env.example .env
```

Jika dibutuhkan, update `pip` terlebih dulu:

```powershell
python -m pip install --upgrade pip
```

## Konfigurasi Penting

Path utama sekarang dibaca dari file `.env`.

Key yang paling penting:

- `OUTPUT_BASE`
- `BASE_DIR_LEGACY`
- `BASE_DIR_PARQUET`
- `TABLE_REFERENCE_PATH`
- `PROJECT_REFERENCE_SOURCE_XLSX`
- `PROJECT_REFERENCE_TARGET_CSV`
- `MANUALS_PATH`
- `RT_REFERENCE_PATH`
- `WA_USER_DATA_DIR`
- `WA_PROFILE_DIRECTORY`
- `WA_HEADLESS`

Contoh nilai default tersedia di file `.env.example`.

## Menjalankan Program

### Opsi A (Direkomendasikan): mode parquet

```powershell
python main_pq.py
```

Catatan:
- Proses `get_data_from_master_pq` berjalan paralel per category/chunk.
- Bagian kirim WhatsApp di akhir `main_pq.py` saat ini dikomentari (nonaktif), jadi default hanya generate file.

### Opsi B: mode legacy

```powershell
python main.py
```

Catatan:
- Alur ini menjalankan pengiriman WA jika ada file yang lolos rule `should_send()`.

## Konfigurasi Report per Project (`list.py`)

Setiap item di `criteria_lists` adalah 1 konfigurasi report untuk 1 customer/project.

### Field Utama (Top-Level)

- `group_name` (wajib): nama project/customer.
- `id_account` (opsional): list account yang difilter. Jika tidak diisi, sistem bisa mapping dari referensi.
- `selected_cols` (wajib): daftar kolom final yang disimpan ke output.
- `save_as` (opsional): override nama file output.
- `pic_frontline` (opsional): nama contact penerima file (string atau list).
- `period` (opsional): saat ini dipakai untuk mode yearly (`"year"`).
- `jumlah_bulan` (opsional): jumlah bulan ke belakang untuk pengambilan data.

### Field Filter, Split, dan Output

- `split_by_id` (opsional, bool): split output menjadi file per `ID_ACCOUNT`.
- `split_by_col_val` (opsional): split output berdasarkan nilai suatu kolom.
- `selected_statuses` (opsional): filter khusus status POD.
- `filter_cols` (opsional): rule filter tambahan.
- `exclude_cols` (opsional): kolom yang dibuang sebelum save.
- `rename_cols` (opsional): mapping rename nama kolom output.
- `move_to_sheet_of_file` (opsional): append hasil ke file Excel target sebagai sheet.
- `clean_receiver` (opsional): aktifkan normalisasi receiver tertentu.
- `manuals_path` (opsional): override path file manual correction untuk project ini.

### Konfigurasi `edit_file`

`edit_file` adalah dictionary untuk proses formatting workbook.

- `replace_values`: merge override dari `manuals.csv`.
- `pivots`: definisi PivotTable yang akan dibuat.
- `date_format`: format tanggal output.
- `rename_data_sheet`: rename sheet master (default `Master`).
- `split_sheet_by_column`: split sheet berdasarkan satu kolom.
  - `column`: nama kolom pemisah.
  - `include_master`: apakah sheet master tetap dipertahankan.
- `split_sheet_by_col_val`: split sheet berdasarkan nilai tertentu.
  - Bisa `dict` atau `list` rule.
- `awb_rt_sheet`: menambahkan sheet AWB RT dari file referensi RT.

### Konfigurasi `cust_ref` (Opsional)

Dipakai jika report butuh merge ke reference customer tambahan.

- `ref_path`: lokasi file/folder referensi customer.
- `ref_sheet`: nama sheet referensi.
- `key_col`: pasangan key join `[key_di_ref, key_di_master]`.
- `selected_ref_cols`: kolom referensi yang ikut dibawa.
- `date_col`: kolom tanggal aktif untuk proses berbasis tanggal.
- `jumlah_bulan`: jangkauan data referensi.
- `key_format`: pattern bantuan untuk ekstraksi key jika key kosong.
- `key_fallback_cols`: kolom fallback untuk bantu isi key.
- `col_order`: urutan kolom final output.

### Contoh Minimal

```python
{
    "group_name": "CONTOH PROJECT",
    "id_account": ["'80000001", "'80000002"],
    "selected_cols": ["AWB", "ID_ACCOUNT", "TGL_ENTRY", "STATUS_POD"],
    "save_as": "CONTOH PROJECT DAILY",
    "edit_file": {
        "date_format": "%Y-%m-%d",
        "rename_data_sheet": "Master",
        "pivots": []
    },
    "pic_frontline": ["Nama Frontline 1", "Nama Frontline 2"]
}
```

Tips:
- Pastikan kolom di `selected_cols` benar-benar tersedia di source data.
- Jika pakai `key_col` di `cust_ref`, urutan yang direkomendasikan adalah `[key_di_ref, key_di_master]`.
- Gunakan `save_as` untuk mencegah nama file bentrok antar konfigurasi yang mirip.

## Mekanisme History & Skip Process

- Status history file: `utils/data/status_history.json`
- Process history harian: `utils/data/process_history.json`
- Rule kirim file: `utils.helper.should_send()`
  - Kirim jika nama file mengandung bulan berjalan/tahun berjalan, atau status berubah.

Di `main_pq.py` ada cleanup otomatis process history lama (`clean_old_process_history(days=14)`).

## Troubleshooting

- Gagal akses folder network:
  - Pastikan VPN/LAN aktif dan path UNC bisa diakses dari Explorer.
- File Excel gagal disimpan (`PermissionError`):
  - Tutup workbook yang masih terbuka, script punya fallback nama file `(1)`, `(2)`, dst.
- Project kategori `UNKNOWN`:
  - Cek dan update `data/project_reference.csv` (atau source referensinya).
- Pivot/COM error:
  - Pastikan Excel terinstal dan `pywin32` sudah terpasang.

## Saran Pengembangan Lanjutan

- Pertimbangkan lockfile terpisah untuk dev/prod jika kebutuhan environment mulai berbeda.
- Pertimbangkan validasi startup untuk key `.env` wajib agar gagal lebih cepat saat path belum benar.
- Tambahkan logging terstruktur (mis. `logging` module) untuk monitoring batch.
