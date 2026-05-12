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

def add_address(df: pd.DataFrame, name: str) -> pd.DataFrame:
    df = df.copy()
    addr_cols = ["ADDR1", "ADDR2", "ADDR3"]
    for col in addr_cols:
        if col not in df.columns:
            df[col] = ""

    # hanya buat kolom ADDRESS kalau belum ada
    if name not in df.columns or df[name].isna().all():
        df[name] = (
            df[addr_cols]
            .fillna("")
            .astype(str)
            .apply(lambda row: " ".join(x for x in row if x.strip()), axis=1)
        )
        
    return df

def add_cust_name(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if "GROUPING_SHIPPER" in df.columns:
        df = df.rename(columns={"GROUPING_SHIPPER": "CUST_NAME"})
    return df

def add_periode(df: pd.DataFrame, debug=False) -> pd.DataFrame:
    df = df.copy()
    for col in ["PERIODE", "WEEK", "WEEK_OF_YEAR", "DAY"]:
        if col in df.columns:
            df.drop(columns=[col], inplace=True)

    if "TGL_ENTRY" not in df.columns:
        print("Kolom 'TGL_ENTRY' tidak ditemukan.")
        return df

    # Pastikan TGL_ENTRY datetime
    if not pd.api.types.is_datetime64_any_dtype(df["TGL_ENTRY"]):
        if debug:
            print("add_periode: TGL_ENTRY belum datetime, mencoba konversi fallback...")
        df["TGL_ENTRY"] = pd.to_datetime(df["TGL_ENTRY"], errors="coerce")

    # Buat kolom PERIODE (nama bulan)
    df["PERIODE"] = df["TGL_ENTRY"].apply(
        lambda val: INDO_MONTHS[val.month] if pd.notna(val) and 1 <= val.month <= 12 else ""
    )

    # Buat kolom WEEK (minggu ke berapa dalam bulan)
    def week_of_month(dt):
        if pd.isna(dt):
            return None
        first_day = dt.replace(day=1)
        return ((dt - first_day).days // 7) + 1

    df["WEEK"] = df["TGL_ENTRY"].apply(week_of_month)

    # Buat kolom WEEK_OF_YEAR (minggu ISO, 1â€“52/53)
    df["WEEK_OF_YEAR"] = df["TGL_ENTRY"].dt.isocalendar().week

    # Buat kolom DAY (tanggal saja, 1â€“31)
    df["DAY"] = df["TGL_ENTRY"].dt.day

    # Buat kolom TIME_ENTRY (jam:menit)
    df["TIME_ENTRY"] = df["TGL_ENTRY"].dt.strftime("%H:%M")

    if debug:
        print(f"add_periode: sample: {df[['TGL_ENTRY','PERIODE','WEEK','WEEK_OF_YEAR','DAY','TIME_ENTRY']].head(3).to_dict('records')}")

    return df

def add_time_received(df: pd.DataFrame, debug=False) -> pd.DataFrame:
    df = df.copy()
    for col in ["TIME_RECEIVED"]:
        if col in df.columns:
            df.drop(columns=[col], inplace=True)

    if "TGL_RECEIVED" not in df.columns:
        print("Kolom 'TGL_RECEIVED' tidak ditemukan.")
        return df

    # Pastikan TGL_RECEIVED datetime
    if not pd.api.types.is_datetime64_any_dtype(df["TGL_RECEIVED"]):
        if debug:
            print("add_time_received: TGL_RECEIVED belum datetime, mencoba konversi fallback...")
        df["TGL_RECEIVED"] = pd.to_datetime(df["TGL_RECEIVED"], errors="coerce")

    # Buat kolom TIME_RECEIVED
    df["TIME_RECEIVED"] = df["TGL_RECEIVED"].dt.strftime("%H:%M:%S")

    if debug:
        print(f"add_time_received: sample: {df[['TGL_RECEIVED','TIME_RECEIVED']].head(3).to_dict('records')}")

    return df

def add_status_latlong(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if "STATUS_LATITUDE" not in df.columns or "STATUS_LONGITUDE" not in df.columns:
        return df

    def combine_latlong(row):
        lat = str(row["STATUS_LATITUDE"]).strip() if pd.notna(row["STATUS_LATITUDE"]) else ""
        lon = str(row["STATUS_LONGITUDE"]).strip() if pd.notna(row["STATUS_LONGITUDE"]) else ""
        if lat and lon and lat.lower() != "nan" and lon.lower() != "nan":
            return f"{lat}, {lon}"
        return ""

    df["STATUS_LATLONG"] = df.apply(combine_latlong, axis=1)
    return df

def resolve_col(df, options, default=pd.NA):
    """
    Cari nama kolom dari daftar `options` yang ada di df.
    Kalau tidak ketemu â†’ buat kolom baru pakai nama pertama.
    """
    for col in options:
        if col in df.columns:
            return col
    # kalau semua ga ada â†’ buat kolom pertama
    df[options[0]] = default
    return options[0]

def add_aging_carrer(df: pd.DataFrame, 
                     holiday_path="//192.168.9.76/d/RYAN/1. References/Holiday.xlsx",
                     debug=False) -> pd.DataFrame:

    df = df.copy()

    debug = True

    # --- Resolve kolom dengan dual name ---
    col_1st_attempt      = resolve_col(df, ["1ST ATTEMPT", "1ST_ATTEMPT_DATE"])
    col_aging_1st        = resolve_col(df, ["AGING 1st ATTEMPT", "AGING_1ST"])
    col_career_1st       = resolve_col(df, ["CAREER 1st ATTEMPT", "CAREER_1ST"])
    col_aging_pod        = resolve_col(df, ["AGING POD", "AGING_POD"])
    col_carrer_pod       = resolve_col(df, ["CARRER POD", "CARRER_POD"])

    # --- Load holiday list ---
    try:
        holidays = pd.read_excel(holiday_path, sheet_name="Holiday", usecols="A")
        holidays = pd.to_datetime(holidays.iloc[:, 0], errors="coerce").dropna().dt.normalize()
        holidays = holidays.values.astype("datetime64[D]")
        if debug:
            print(f"[i] Holidays loaded: {len(holidays)} hari libur")
    except Exception as e:
        print(f"[!] Gagal baca holiday file: {e}")
        holidays = np.array([], dtype="datetime64[D]")

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

            # SAME DAY â†’ 0 hari kerja
            if end_date <= start_date:
                return 0

            return np.busday_count(start_date, end_date, holidays=holidays) + 1
        except Exception:
            return pd.NA

    # --- Define mask for Success ---
    mask_success = df["STATUS_POD"].astype(str).str.lower().eq("success")
    mask_non_success = ~mask_success

    # --- Clear untuk non-success ---
    for col in [col_1st_attempt, col_aging_1st, col_career_1st, col_aging_pod, col_carrer_pod]:
        if col in df.columns:
            df.loc[mask_non_success, col] = pd.NA

    # --- AGING 1st ATTEMPT ---
    mask_empty = df[col_aging_1st].isna() & mask_success
    df.loc[mask_empty, col_aging_1st] = df[mask_empty].apply(
        lambda row: networkdays(row["TGL_ENTRY"], row[col_1st_attempt]), axis=1
    )

    # --- CAREER 1st ATTEMPT ---
    mask_empty = df[col_career_1st].isna() & mask_success
    df.loc[mask_empty, col_career_1st] = np.where(
        (df.loc[mask_empty, col_aging_1st].notna()) & (df.loc[mask_empty, "ETD"].notna()),
        np.where(df.loc[mask_empty, col_aging_1st] <= df.loc[mask_empty, "ETD"], "Ontime SLA", "Over SLA"),
        pd.NA
    )

    # --- AGING POD ---
    mask_empty = df[col_aging_pod].isna() & mask_success
    df.loc[mask_empty, col_aging_pod] = df[mask_empty].apply(
        lambda row: networkdays(row["TGL_ENTRY"], row["TGL_RECEIVED"]), axis=1
    )

    # --- CARRER POD ---
    mask_empty = df[col_carrer_pod].isna() & mask_success
    df.loc[mask_empty, col_carrer_pod] = np.where(
        (df.loc[mask_empty, col_aging_pod].notna()) & (df.loc[mask_empty, "ETD"].notna()),
        np.where(df.loc[mask_empty, col_aging_pod] <= df.loc[mask_empty, "ETD"], "Ontime SLA", "Over SLA"),
        pd.NA
    )

    return df

def add_PIC(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["PIC"] = "JNE"  

    return df

def add_SBS(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if "ORIGIN" not in df.columns:
        print("Kolom 'ORIGIN' tidak ditemukan, tidak bisa menambahkan ORIGIN 2.")
        return df

    df["ORIGIN 2"] = np.where(
        df["ORIGIN"].astype(str).str.startswith("CGK"),
        "Non SBS",
        "SBS"
    )
    return df

def add_date_time_received(df: pd.DataFrame, debug=False) -> pd.DataFrame:
    df = df.copy()
    for col in ["DATE_RECEIVED", "TIME_RECEIVED"]:
        if col in df.columns:
            df.drop(columns=[col], inplace=True)

    if "TGL_RECEIVED" not in df.columns:
        print("Kolom 'TGL_RECEIVED' tidak ditemukan.")
        return df

    # Pastikan TGL_RECEIVED datetime
    if not pd.api.types.is_datetime64_any_dtype(df["TGL_RECEIVED"]):
        if debug:
            print("add_date_time_received: TGL_RECEIVED belum datetime, mencoba konversi fallback...")
        df["TGL_RECEIVED"] = pd.to_datetime(df["TGL_RECEIVED"], errors="coerce")

    # Buat kolom
    df["DATE_RECEIVED"] = df["TGL_RECEIVED"].dt.date
    df["TIME_RECEIVED"] = df["TGL_RECEIVED"].dt.strftime("%H:%M:%S")

    if debug:
        print(f"add_time_received: sample: {df[['TGL_RECEIVED','DATE_RECEIVED','TIME_RECEIVED']].head(3).to_dict('records')}")

    return df

def add_dept_PZC(df: pd.DataFrame, debug=False) -> pd.DataFrame:
    df = df.copy()
    df = add_cust_name(df)
    for col in ["DEPT PZC"]:
        if col in df.columns:
            df.drop(columns=[col], inplace=True)

    if "ID_ACCOUNT" not in df.columns or "CUST_NAME" not in df.columns:
        print("Kolom 'ID_ACCOUNT' atau 'CUST_NAME' tidak ditemukan.")
        return df

    # Buat kolom
    df["DEPT PZC"] = df["ID_ACCOUNT"].astype(str) + " " + df["CUST_NAME"].astype(str)

    if debug:
        print(f"add_dept_PZC: sample: {df[['ID_ACCOUNT','CUST_NAME','DEPT PZC']].head(3).to_dict('records')}")

    return df

def add_is_close(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if "STATUS_POD" not in df.columns:
        print("Kolom 'STATUS_POD' tidak ditemukan, tidak bisa menambahkan IS_CLOSE.")
        return df

    # Tambah kolom IS_CLOSE (string "x" atau kosong)
    df["IS_CLOSE"] = df["STATUS_POD"].apply(
        lambda x: "x" if str(x).strip().lower() in ["success", "return shipper"] else ""
    )

    # Urutkan: yang IS_CLOSE = "x" dulu, lalu STATUS_POD A-Z
    df = df.sort_values(by=["IS_CLOSE", "STATUS_POD"], ascending=[False, True]).reset_index(drop=True)

    return df

def add_reason(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    if "STATUS_POD" not in df.columns:
        print("Kolom 'STATUS_POD' tidak ditemukan, tidak bisa menambahkan REASON.")
        return df

    # pastikan kolom AGING_POD ada, kalau tidak isi dengan None
    if "AGING_POD" not in df.columns:
        df["AGING_POD"] = None

    def map_reason(row):
        if row["STATUS_POD"] == "Success":
            return "CLOSE CABANG"
        elif row["STATUS_POD"] == "Return Shipper":
            return "CLOSE CABANG"
        elif row["STATUS_POD"] != "Success":
            return "ON FOLLOW UP CABANG"
        elif row["AGING_POD"] != "Return Shipper":
            return "ON FOLLOW UP CABANG"
        else:
            return None

    df["REASON"] = df.apply(map_reason, axis=1)
    return df

def add_no(df: pd.DataFrame, debug: bool = False) -> pd.DataFrame:
    """
    Adds a 'NO' column with sequential numbers (1, 2, 3, ...) to the DataFrame.
    If DataFrame is empty, adds an empty 'NO' column.
    
    Args:
        df (pd.DataFrame): Input DataFrame.
        debug (bool): If True, prints debug information.
    
    Returns:
        pd.DataFrame: DataFrame with 'NO' column added at index 0.
    """
    df = df.copy()
    if debug:
        print(f"[DEBUG] add_no: DataFrame size: {len(df)} rows")
    if "NO" in df.columns:
        df = df.drop(columns=["NO"])
        if debug:
            print("[DEBUG] add_no: Dropped existing 'NO' column")
    if len(df) > 0:
        df.insert(0, "NO", range(1, len(df) + 1))
        if debug:
            print(f"[DEBUG] add_no: Added 'NO' column with values 1..{len(df)}")
            print(f"[DEBUG] add_no: First few 'NO' values: {df['NO'].head().tolist()}")
    else:
        df.insert(0, "NO", [])
        if debug:
            print("[DEBUG] add_no: DataFrame empty, added empty 'NO' column")
    return df

def add_date_receive_request(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    def compute_value(row):
        val = row.get("CONNOTE_RETURN_RT")
        if pd.isna(val):
            return row.get("TGL_RECEIVED")
        elif isinstance(val, str) and "RT" in val:
            return row.get("DATE_CONNOTE_RETURN_RT")
        else:
            return None

    df["DATE_RECEIVE_REQUEST"] = df.apply(compute_value, axis=1)

    return df

def add_descr_return(df: pd.DataFrame) -> pd.DataFrame:
    """
    Tambahkan kolom DESCR_RETURN:
    - "RETURN SICEPAT" jika STATUS_POD == "Return Shipper"
    - else kosong (None)
    """
    df = df.copy()
    if "STATUS_POD" not in df.columns:
        print("Kolom 'STATUS_POD' tidak ditemukan, tidak bisa menambahkan DESCR_RETURN.")
        return df

    df["DESCR_RETURN"] = np.where(
        df["STATUS_POD"].astype(str).str.strip().eq("Return Shipper"),
        "RETURN SICEPAT",
        None
    )

    return df

def add_update_time(df: pd.DataFrame) -> pd.DataFrame:
    """
    Tambahkan kolom UPDATE_TIME:
    - Isi dengan tanggal hari ini (tanpa jam).
    """
    df = df.copy()
    today = datetime.datetime.today().date()  # hanya tanggal
    df["UPDATE_TIME"] = today
    return df

def fix_regional_cols(df: pd.DataFrame) -> pd.DataFrame:
    """
    Revisi semua kolom yang mengandung kata 'REGIONAL'
    - Replace 'Sumatara' (case-insensitive) menjadi 'Sumatera'
    """
    for col in df.columns:
        if "REGIONAL" in col.upper():
            # Pastikan kolom bertipe string
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(r"(?i)sumatara", "Sumatera", regex=True)
            )
    return df

def fix_contact_notelp_col(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    if "CONTACT" not in df.columns or "NOTELP" not in df.columns:
        print("Kolom 'CONTACT' atau 'NOTELP' tidak ditemukan, tidak bisa memperbaiki CONTACT_NOTELP.")
        return df

    df["CONTACT"] = "*****"
    df["NOTELP"] = "*****"

    return df

def fix_empty_date_1st_attempt(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    if "DATE_1ST_ATTEMPT" not in df.columns:
        print("Kolom 'DATE_1ST_ATTEMPT' tidak ditemukan, tidak bisa memperbaiki DATE_1ST_ATTEMPT.")
        return df

    # Ganti nilai kosong atau NaN di DATE_1ST_ATTEMPT menjadi TGL_RECEIVED hanya jika STATUS_POD = "Success"
    if "STATUS_POD" in df.columns and "TGL_RECEIVED" in df.columns:
        mask = (
            df["STATUS_POD"].astype(str).str.strip().eq("Success") &
            df["DATE_1ST_ATTEMPT"].isna()
        )
        df.loc[mask, "DATE_1ST_ATTEMPT"] = df.loc[mask, "TGL_RECEIVED"]

    return df

def add_3lc_dest_fw(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    if "DEST_FW" not in df.columns:
        print("Kolom 'DEST_FW' tidak ditemukan, tidak bisa menambahkan 3 LC DEST FW.")
        return df
    
    df["3 LC DEST FW"] = df["DEST_FW"].str[:3]

    return df

def add_rounded_weight(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    if "WEIGHT" not in df.columns:
        print("Kolom 'WEIGHT' tidak ditemukan.")
        return df

    df["WEIGHT"] = pd.to_numeric(df["WEIGHT"], errors="coerce")

    rounded = np.where(
        (df["WEIGHT"] % 1) > 0.3,
        np.ceil(df["WEIGHT"]),
        np.floor(df["WEIGHT"])
    )

    # Pastikan minimal 1
    df["WEIGHT"] = np.maximum(1, rounded)

    return df

