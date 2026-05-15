import pandas as pd
import numpy as np
import math
import datetime
import os
import glob
import re
from pandas.tseries.offsets import CustomBusinessDay
from dateutil.relativedelta import relativedelta

INDO_MONTHS = [
    "", "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI",
    "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"
]

from data_transform.status_pod import add_status_pod

def add_AJCar_status(
    df: pd.DataFrame,
    ref=r"\\192.168.9.76\D\RYAN\1. References\Table Reference.xlsx",
    debug=False
) -> pd.DataFrame:
    df = df.copy()
    df = add_status_pod(df)
    
    # hapus AJCar_Status lama kalau ada
    for col in ["AJCar_Status"]:
        if col in df.columns:
            df.drop(columns=[col], inplace=True)

    if "STATUS_POD" not in df.columns:
        print("Kolom 'STATUS_POD' tidak ditemukan.")
        return df
    
    # Buat kolom
    try:
        ref_df = pd.read_excel(ref, sheet_name="posisisis")
        required = {"Status", "Grouping BIG Status", "AJ Car Status", "KETERANGAN AJ CAR"}
        if not required.issubset(ref_df.columns):
            print("Sheet posisisis tidak memiliki kolom Status, Grouping BIG Status, AJ Car Status atau KETERANGAN AJ CAR.")
            return df

        df = df.merge(
            ref_df[["Status", "Grouping BIG Status", "AJ Car Status", "KETERANGAN AJ CAR"]],
            how="left",
            left_on="STATUS_POD",
            right_on="Status"
        )

        # drop Status join key
        df.drop(columns=["Status"], inplace=True)

        # rename Grouping BIG Status → Remark Return
        df.rename(columns={"Grouping BIG Status": "Remark Return"}, inplace=True)

    except Exception as e:
        print(f"Gagal menambahkan kolom AJ Car Status & KETERANGAN AJ CAR: {e}")

    return df

def load_uob_pickup_data(bulan: int, tahun: int) -> pd.DataFrame:
    """
    Baca semua file .txt pickup UOB untuk bulan+tahun tertentu.
    Return DataFrame dengan kolom:
    REFF NUM, CUSTOMER NAME, PICK UP DATE, CYCLE NAME, 
    NAMA FILE PICK UP, CYCLE & JENIS
    """
    all_rows = []
    UOB_BASE = r"D:\RYAN\Python Scripts\Bot Report Gabungan\data\uob"

    bulan_id = {
        1: "JANUARI", 2: "FEBRUARI", 3: "MARET", 4: "APRIL",
        5: "MEI", 6: "JUNI", 7: "JULI", 8: "AGUSTUS",
        9: "SEPTEMBER", 10: "OKTOBER", 11: "NOVEMBER", 12: "DESEMBER"
    }
    bulan_str = bulan_id[bulan].upper()

    # mapping kategori singkatan
    CATEGORY_MAP = {
        "BS CREDIT CARD": "BSCC",
        "BS UOB CASHPLUS": "BSCP",
    }

    # cari folder sesuai bulan & tahun
    pattern = os.path.join(UOB_BASE, f"*{bulan_str} {tahun}*")
    folders = glob.glob(pattern)

    for folder in folders:
        folder_name = os.path.basename(folder)

        # --- cari cycle name ---
        match = re.search(r"(CY\s[0-9\-]+\s+[A-Z]+\s+\d{4})", folder_name.upper())
        cycle_name = match.group(1).upper() if match else ""

        # --- ambil kategori ---
        kategori_raw = folder_name.upper().replace("DATA ", "").split("CY")[0].strip()
        kategori_singkat = CATEGORY_MAP.get(kategori_raw, kategori_raw.replace(" ", ""))

        # --- buat base nama file ---
        nama_file_base = f"JNE_{kategori_singkat} {cycle_name.rsplit(' ', 1)[0]}"

        for txt_file in glob.glob(os.path.join(folder, "*.txt")):
            txt_name = os.path.basename(txt_file).upper()

            if "SD_OUTPUT" in txt_name:
                nama_file_pickup = f"{nama_file_base}_SD.txt"
            else:
                nama_file_pickup = f"{nama_file_base}.txt"

            # --- tentukan CYCLE & JENIS ---
            jenis_prefix = "BS" if kategori_singkat == "BSCC" else "CP"
            # --- buang bulan & tahun dari cycle_name ---
            cycle_part = re.sub(r"\s+[A-Z]+\s+\d{4}$", "", cycle_name).strip()

            if nama_file_pickup.endswith("_SD.txt"):
                cycle_jenis = f"{jenis_prefix} {cycle_part} STAMP"
            else:
                cycle_jenis = f"{jenis_prefix} {cycle_part} NON STAMP"

            with open(txt_file, "r", encoding="latin-1") as f:
                lines = f.readlines()

            for line in lines[1:]:
                parts = [p.strip() for p in line.split(";") if p.strip()]
                if len(parts) < 3:
                    continue

                all_rows.append({
                    "REFF NUM": parts[0],
                    "CUSTOMER NAME": parts[1],
                    "PICK UP DATE": parts[-1],
                    "CYCLE NAME": cycle_name,
                    "NAMA FILE PICK UP": nama_file_pickup,
                    "CYCLE & JENIS": cycle_jenis,
                    "PERIODE UOB": bulan_str
                })

    if not all_rows:
        return pd.DataFrame(columns=[
            "REFF NUM", "CUSTOMER NAME", "PICK UP DATE", 
            "CYCLE NAME", "NAMA FILE PICK UP", "CYCLE & JENIS"
        ])

    df_uob = pd.DataFrame(all_rows)

    # normalisasi tanggal
    df_uob["PICK UP DATE"] = pd.to_datetime(
        df_uob["PICK UP DATE"], errors="coerce", format="%m/%d/%Y"
    )

    # --- tambah kolom JENIS BILLING ---
    def get_jenis_billing(nama_file: str) -> str:
        nama_file = str(nama_file).upper()
        if "BSCC" in nama_file:
            return "BS STAMP" if "_SD" in nama_file else "BS NON STAMP"
        elif "BSCP" in nama_file:
            return "BSCP STAMP" if "_SD" in nama_file else "BSCP NON STAMP"
        return pd.NA

    df_uob["JENIS BILLING"] = df_uob["NAMA FILE PICK UP"].map(get_jenis_billing)

    return df_uob

def add_uob_pickup_data_cols(df: pd.DataFrame, bulan_ke_belakang: int = 2) -> pd.DataFrame:
    """
    Join pickup UOB data ke df berdasarkan REFNO_UOB = REFF NUM.
    Tambah kolom: 
      - REFF NUM
      - CUSTOMER NAME
      - PICK UP DATE
      - CYCLE NAME
      - RECIPIENT
      - RECIPIENT RELATION (hanya untuk STATUS UOB = 'S')
      - STATUS UOB
      - REASON 2 (hanya untuk STATUS UOB = 'R')
      - REMARK (AWB jika STATUS UOB = 'S', atau AWB + REASON RETURN jika STATUS UOB = 'R')
    """
    today = datetime.datetime.today()
    all_uob = []

    # --- load data UOB beberapa bulan terakhir ---
    for i in range(bulan_ke_belakang):
        target_date = today - relativedelta(months=i)
        bulan = target_date.month
        tahun = target_date.year
        df_uob = load_uob_pickup_data(bulan, tahun)
        if not df_uob.empty:
            all_uob.append(df_uob)

    if not all_uob:
        print(f"Tidak ada data UOB ditemukan {bulan_ke_belakang} bulan terakhir")
        for col in [
            "REFF NUM", "CUSTOMER NAME", "PICK UP DATE", "CYCLE NAME",
            "RECIPIENT", "RECIPIENT RELATION", "STATUS UOB", "REASON 2", "REMARK"
        ]:
            df[col] = pd.NA
        return df

    df_uob_all = (
        pd.concat(all_uob, ignore_index=True)
        .drop_duplicates(subset=["REFF NUM"])
        .copy()
    )

    # --- normalisasi key ---
    if "REFNO_UOB" not in df.columns:
        print("Kolom REFNO_UOB tidak ditemukan di dataframe utama")
        df["REFNO_UOB"] = ""

    df["REFNO_UOB"] = (
        df["REFNO_UOB"].astype(str)
        .str.strip()
        .str.replace("'", "", regex=False)
        .str.replace("*", "", regex=False)
    )
    df_uob_all["REFF NUM"] = df_uob_all["REFF NUM"].astype(str).str.strip()

    # --- merge ---
    df_merged = df.merge(
        df_uob_all,
        how="right",
        left_on="REFNO_UOB",
        right_on="REFF NUM"
    )

    if "REFF NUM_x" in df_merged.columns:
        df_merged.drop(columns=["REFF NUM_x"], inplace=True, errors="ignore")
        df_merged.rename(columns={"REFF NUM_y": "REFF NUM"}, inplace=True)

    # --- kolom RECIPIENT ---
    if {"STATUS_POD", "RECEIVED/REASON"}.issubset(df_merged.columns):
        df_merged["RECIPIENT"] = np.where(
            df_merged["STATUS_POD"].astype(str).str.lower().eq("success"),
            df_merged["RECEIVED/REASON"],
            pd.NA
        )
    else:
        df_merged["RECIPIENT"] = pd.NA

    # --- mapping CODING ---
    coding_map = {
        "CR1": ("Retur Origin", "R"),
        "D01": ("01", "S"),
        "D02": ("07", "S"),
        "D03": ("08", "S"),
        "D04": ("06", "S"),
        "D05": ("08", "S"),
        "D06": ("02", "S"),
        "D07": ("05", "S"),
        "D08": ("08", "S"),
        "D09": ("03", "S"),
        "D10": ("10", "S"),
        "D11": ("08", "S"),
        "D12": ("08", "S"),
        "DB1": ("08", "S"),
        "DB2": ("08", "S"),
        "R10": ("06", "R"),
        "U01": ("03", "R"),
        "U02": ("04", "R"),
        "U03": ("02", "R"),
        "U04": ("06", "R"),
        "U05": ("07", "R"),
        "U06": ("08", "R"),
        "U07": ("07", "R"),
        "U08": ("08", "R"),
        "U09": ("07", "R"),
        "U10": ("10", "R"),
        "U11": ("06", "R"),
        "U12": ("06", "R"),
        "U13": ("08", "R"),
        "U14": ("07", "R"),
        "U21": ("08", "R"),
        "U22": ("07", "R"),
        "U23": ("07", "R"),
        "U24": ("06", "R"),
        "U25": ("07", "R"),
        "OPC": ("On Process", "On Process"),
    }

    df_merged["CODING"] = df_merged["CODING"].astype(str)
    
    # --- perbaikan untuk coding kosong ---
    df_merged["CODING"] = (
        df_merged["CODING"]
        .fillna("OPC")             # isi NaN dengan OPC
        .replace(r"^\s*$", "OPC", regex=True)  # isi string kosong / spasi jadi OPC
    )

    # hasil mapping
    df_merged["MAP_RELATION"] = df_merged["CODING"].map(lambda x: coding_map.get(x, (pd.NA, pd.NA))[0])
    df_merged["STATUS UOB"] = df_merged["CODING"].map(lambda x: coding_map.get(x, (pd.NA, pd.NA))[1])

    # isi RECIPIENT RELATION hanya untuk STATUS UOB = "S"
    df_merged["RECIPIENT RELATION"] = np.where(
        df_merged["STATUS UOB"] == "S",
        df_merged["MAP_RELATION"],
        pd.NA
    )

    # isi REASON 2 hanya untuk STATUS UOB = "R"
    df_merged["REASON 2"] = np.where(
        df_merged["STATUS UOB"] == "R",
        df_merged["MAP_RELATION"],
        pd.NA
    )

    # --- isi kolom REMARK ---
    df_merged["REMARK_UOB"] = np.where(
        df_merged["STATUS UOB"] == "S",
        df_merged["AWB"],
        np.where(
            df_merged["STATUS UOB"] == "R",
            df_merged["AWB"].astype(str) + " " + df_merged["REASON RETURN"].astype(str),
            pd.NA
        )
    )

    # --- isi kolom COURIER ID
    df_merged["COURIER ID"] = "26"

    # --- isi kolom UPLOAD DATE
    df_merged["UPLOAD DATE"] = df_merged["PICK UP DATE"]

    # --- isi kolom KET AWB ---
    df_merged["KET AWB"] = np.where(
        df_merged["AWB"].notna() & df_merged["AWB"].astype(str).str.strip().ne(""),
        "ADA AWB",
        "TIDAK ADA AWB"
    )

    # --- isi kolom CLOSING DATE ---
    df_merged["CLOSING DATE"] = df_merged["TGL_RECEIVED"]

    # --- isi kolom STATUS CLOSING ---
    df_merged["STATUS CLOSING"] = np.where(
        df_merged["CLOSING DATE"].notna() & df_merged["CLOSING DATE"].astype(str).str.strip().ne(""),
        "SUDAH CLOSING",
        "BELUM CLOSING"
    )

    # drop kolom intermediate
    df_merged.drop(columns=["MAP_RELATION"], inplace=True)

    return df_merged

def add_sociolla_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Tambah kolom DIM dll
    df["DIM"] = 0
    df["DIM - Copy"] = df["DIM"]

    return df

def add_SPK(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    pattern = r"(SPK\/[A-Z0-9\-]+\/\d+\/[A-Z]+\/\d+)"
    
    df["SPK"] = df["INTRUCTION"].str.extract(pattern)

    return df

def add_young_living_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Ensure expected input columns exist to avoid KeyError
    for _c in ("AWB", "NOREF", "AMOUNT", "WEIGHT", "PROVINSI", "OTMS Order", "Buyer PO", "ETD"):
        if _c not in df.columns:
            df[_c] = pd.NA

    # Basic columns / defaults
    df["OTMS Order"] = df.get("OTMS Order", '0').fillna('0')
    df["Target"] = df.get("Target", "99%")
    df["Origin_WH"] = df.get("Origin_WH", "PDU REG")
    df["Vendor"] = df.get("Vendor", "JNE")

    # No_YL from NOREF; coerce to str and strip
    df["No_YL"] = df["NOREF"].astype(str).str.strip().replace("nan", "")

    # No DSV: last 7 chars of AWB (if present)
    df["No DSV"] = df["AWB"].fillna("").astype(str).str[-7:]

    df["Buyer PO"] = df.get("Buyer PO", "")

    df["City Code"] = df["OTMS Order"]

    # More robust mapping for Indonesia Bagian using province keywords
    def _map_bagian(prov):
        p = str(prov).lower()
        if p in ("nan", "none"):
            return pd.NA

        west_kw = (
            "sumatera", "aceh", "riau", "kepulauan riau", "kepri", "bengkulu",
            "lampung", "jambi", "riau", "banten", "jakarta", "jawa barat", "jawa",
            "yogyakarta", "yogya", "bali"
        )
        central_kw = ("sulawesi", "kalimantan", "borneo")
        east_kw = ("papua", "maluku", "nusa tenggara", "ntb", "ntt")

        if any(k in p for k in west_kw):
            return "Indonesia Bagian Barat"
        if any(k in p for k in central_kw):
            return "Indonesia Bagian Tengah"
        if any(k in p for k in east_kw):
            return "Indonesia Bagian Timur"
        return "Indonesia Bagian Timur"

    df["Indonesia Bagian"] = df["PROVINSI"].map(_map_bagian)

    # Harga Per KG: safe numeric division, handle zero/inf
    amt = pd.to_numeric(df["AMOUNT"], errors="coerce")
    wt = pd.to_numeric(df["WEIGHT"], errors="coerce")

    # raw price per kg
    raw = amt / wt
    raw = raw.replace([np.inf, -np.inf], pd.NA)

    # Round up to next 1000 step (e.g., 1 -> 1000, 1001 -> 2000)
    rounded = np.floor(raw / 1000.0) * 1000.0

    # Apply only where raw is valid; keep NA otherwise
    df["Harga Per KG"] = pd.Series(rounded).where(raw.notna())

    def _coerce_datetime_with_fallback(series):
        raw = pd.Series(series, index=df.index)
        raw_text = raw.astype(str).str.strip()
        dayfirst_mask = raw_text.str.match(r"^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}")

        parsed = pd.to_datetime(raw, errors="coerce", format="mixed")
        if dayfirst_mask.any():
            parsed.loc[dayfirst_mask] = pd.to_datetime(
                raw.loc[dayfirst_mask], errors="coerce", format="mixed", dayfirst=True
            )

        missing = parsed.isna() & raw.notna() & raw.astype(str).str.strip().ne("")
        if missing.any():
            parsed_dayfirst = pd.to_datetime(
                raw.loc[missing], errors="coerce", format="mixed", dayfirst=True
            )
            parsed.loc[missing] = parsed_dayfirst
        return parsed

    # Ensure datetime columns before arithmetic
    df["TGL_RECEIVED_RAW_DEBUG"] = df.get("TGL_RECEIVED", pd.NA)
    df["TGL_ENTRY_RAW_DEBUG"] = df.get("TGL_ENTRY", pd.NA)
    df["TGL_RECEIVED"] = _coerce_datetime_with_fallback(df.get("TGL_RECEIVED", pd.NA))
    df["TGL_ENTRY"] = _coerce_datetime_with_fallback(df.get("TGL_ENTRY", pd.NA))

    mask_success = (
        df["STATUS_POD"].astype(str).str.strip().str.lower().eq("success")
        if "STATUS_POD" in df.columns
        else pd.Series(False, index=df.index)
    )
    mask_success_dates = mask_success & df["TGL_RECEIVED"].notna() & df["TGL_ENTRY"].notna()

    # ATA and LEAD TIME only for successful deliveries
    df["ATA"] = pd.NA
    df["LEAD TIME"] = pd.NA
    if mask_success_dates.any():
        if "TIME_RECEIVED" in df.columns:
            df.loc[mask_success_dates, "ATA"] = df.loc[mask_success_dates, "TIME_RECEIVED"]
        df.loc[mask_success_dates, "LEAD TIME"] = (
            df.loc[mask_success_dates, "TGL_RECEIVED"] - df.loc[mask_success_dates, "TGL_ENTRY"]
        ).dt.days

    # SLA_: ensure ETD numeric
    df["SLA_"] = pd.to_numeric(df.get("ETD", pd.NA), errors="coerce")

    df["REMARK_YL"] = np.where(mask_success, "OK", pd.NA)

    df["AWB 2"] = df["AWB"].astype(str).str.strip().replace("nan", "")

    df["TODAY DATE"] = pd.Timestamp.today().normalize()

    df["Overday Before Holiday"] = pd.Series(pd.NA, index=df.index, dtype="Int64")
    df.loc[mask_success_dates, "Overday Before Holiday"] = (
        df.loc[mask_success_dates, "TGL_RECEIVED"] - df.loc[mask_success_dates, "TGL_ENTRY"]
    ).dt.days.add(1).astype("Int64")

    # --- Load holiday list ---
    holiday_path = r"//192.168.9.76/d/RYAN/1. References/Holiday.xlsx"
    try:
        holidays = pd.read_excel(holiday_path, sheet_name="Holiday", usecols="A")
        holidays = pd.to_datetime(holidays.iloc[:, 0], errors="coerce").dropna().dt.normalize()
        holidays = holidays.values.astype("datetime64[D]")
    except Exception as e:
        print(f"[!] Gagal baca holiday file: {e}")
        holidays = np.array([], dtype="datetime64[D]")

    success_count = int(mask_success.sum())
    success_valid_dates_count = int(mask_success_dates.sum())
    success_missing_entry_count = int((mask_success & df["TGL_ENTRY"].isna()).sum())
    success_missing_received_count = int((mask_success & df["TGL_RECEIVED"].isna()).sum())
    print(
        "[YL DEBUG] "
        f"rows={len(df)}, success={success_count}, "
        f"success_valid_dates={success_valid_dates_count}, "
        f"success_missing_TGL_ENTRY={success_missing_entry_count}, "
        f"success_missing_TGL_RECEIVED={success_missing_received_count}, "
        f"holidays_loaded={len(holidays)}"
    )
    if success_count and success_valid_dates_count == 0:
        sample_cols = [
            "AWB", "STATUS_POD", "TGL_ENTRY_RAW_DEBUG", "TGL_ENTRY",
            "TGL_RECEIVED_RAW_DEBUG", "TGL_RECEIVED"
        ]
        sample_cols = [col for col in sample_cols if col in df.columns]
        print("[YL DEBUG] Sample success tanpa tanggal valid:")
        print(df.loc[mask_success, sample_cols].head(10).to_string(index=False))

    # --- Helper: networkdays ---
    def networkdays(start, end):
        if pd.isna(start) or pd.isna(end):
            return pd.NA
        try:
            start_dt = pd.to_datetime(start)
            end_dt = pd.to_datetime(end)

            # NORMALISASI KE TANGGAL SAJA
            start_date = np.datetime64(start_dt.normalize().date(), "D")
            end_date = np.datetime64(end_dt.normalize().date(), "D")

            if end_date < start_date:
                return 0

            return np.busday_count(
                start_date,
                end_date + np.timedelta64(1, "D"),
                holidays=holidays,
            )
        except Exception:
            return pd.NA

    df["Overday Success After Holiday"] = pd.Series(pd.NA, index=df.index, dtype="Int64")
    df.loc[mask_success_dates, "Overday Success After Holiday"] = (
        df.loc[mask_success_dates]
        .apply(lambda row: networkdays(row["TGL_ENTRY"], row["TGL_RECEIVED"]), axis=1)
        .astype("Int64")
    )

    df["Total Tgl Merah"] = pd.Series(pd.NA, index=df.index, dtype="Int64")
    mask_total_tgl_merah = (
        df["Overday Success After Holiday"].notna()
        & df["Overday Before Holiday"].notna()
    )
    df.loc[mask_total_tgl_merah, "Total Tgl Merah"] = (
        df.loc[mask_total_tgl_merah, "Overday Before Holiday"]
        - df.loc[mask_total_tgl_merah, "Overday Success After Holiday"]
    ).astype("Int64")
    negative_total_count = int((df["Total Tgl Merah"].dropna() < 0).sum())

    print(
        "[YL DEBUG] "
        f"overday_before_filled={int(df['Overday Before Holiday'].notna().sum())}, "
        f"overday_after_filled={int(df['Overday Success After Holiday'].notna().sum())}, "
        f"total_tgl_merah_filled={int(df['Total Tgl Merah'].notna().sum())}, "
        f"total_tgl_merah_negative={negative_total_count}"
    )
    if success_valid_dates_count and df["Total Tgl Merah"].isna().all():
        sample_cols = [
            "AWB", "STATUS_POD", "TGL_ENTRY", "TGL_RECEIVED",
            "Overday Before Holiday", "Overday Success After Holiday", "Total Tgl Merah"
        ]
        sample_cols = [col for col in sample_cols if col in df.columns]
        print("[YL DEBUG] Sample overday yang gagal isi Total Tgl Merah:")
        print(df.loc[mask_success_dates, sample_cols].head(10).to_string(index=False))

    df.drop(columns=["TGL_RECEIVED_RAW_DEBUG", "TGL_ENTRY_RAW_DEBUG"], inplace=True, errors="ignore")
    mask_fail = (
        mask_success
        & df["SLA_"].notna()
        & df["LEAD TIME"].notna()
        & (df["LEAD TIME"] > df["SLA_"])
    )

    df["Status Delivery"] = np.select(
        [
            ~mask_success,
            mask_success & df["TGL_RECEIVED"].isna(),
            mask_fail,
        ],
        [
            pd.NA,
            "On Process",
            "Fail",
        ],
        default="Achieve",
    )

    # MOD: LEAD TIME <= 0 → SDS, LEAD TIME = 1 → NDS, LEAD TIME = 2 → H+2, LEAD TIME > 2 → H++
    df["MOD"] = np.where(
        df["LEAD TIME"].notna(),
        np.where(
            df["LEAD TIME"] <= 0, "SDS",
            np.where(
                df["LEAD TIME"] == 1, "NDS",
                np.where(df["LEAD TIME"] == 2, "H+2", "H++")
            )
        ),
        pd.NA
    )

    return df

def add_rodamas_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    df["NAMA ORIGIN"] = "JNE DKI JAKARTA"
    df["CUSTOMER"] = "HO"

    return df

def add_bni_kategori(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    if "GOODS_DESCR" not in df.columns:
        print("Kolom 'GOODS_DESCR' tidak ditemukan, tidak bisa menambahkan KATEGORI BNI.")
        return df

    pattern = r"ULANG TAHUN|ULTAH|UCAPAN"

    mask = (
        df["GOODS_DESCR"]
        .fillna("")
        .astype(str)
        .str.upper()
        .str.contains(pattern, regex=True, na=False)
    )

    df["BNI_KATEGORI"] = "REGULER"
    df.loc[mask, "BNI_KATEGORI"] = "ATENSI ULANG TAHUN"

    return df

def add_mandiri_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    df_area = pd.DataFrame()
    try:
        df_area = pd.read_excel(r"D:\RYAN\1. References\Mandiri\Table Area OOB.xlsx")
    except Exception as e:
        print(f"Gagal membaca file area Mandiri: {e}")
        return df

    # Tambah kolom untuk AREA MANDIRI dan SLA OOB berdasarkan REGIONAL = REGION di df_area
    if "REGIONAL" not in df.columns:
        print("Kolom 'REGIONAL' tidak ditemukan, tidak bisa menambahkan AREA MANDIRI.")
        df["AREA MANDIRI"] = pd.NA
        return df

    required_cols = {"REGION", "AREA MANDIRI", "SLA"}
    if not required_cols.issubset(df_area.columns):
        print("File area Mandiri tidak memiliki kolom REGION, AREA MANDIRI, dan SLA OOB.")
        return df
    
    df = df.merge(
        df_area[["REGION", "AREA MANDIRI", "SLA"]],
        how="left",
        left_on="REGIONAL",
        right_on="REGION",
    )
    df.drop(columns=["REGION"], inplace=True)
    df.rename(columns={"SLA": "SLA OOB"}, inplace=True)

    # Hitung kolom CARRER berdasarkan SLA OOB - AGING < 0 → "OVER SLA", else "On Time SLA"
    df["CARRER"] = np.where(
        df["SLA OOB"] - df["AGING"] < 0,
        "Over SLA",
        "On Time SLA"
    )

    return df
