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

def add_reason_undel(df: pd.DataFrame, ref=r"\\192.168.9.76\D\RYAN\1. References\Table Reference.xlsx") -> pd.DataFrame:
    df = df.copy()

    try:
        if "CODING_UNDEL" not in df.columns:
            print("Kolom CODING_UNDEL tidak ditemukan")
            return df

        ref_df = pd.read_excel(ref, sheet_name="Coding Undel")

        # Validasi kolom
        if not {"Coding Undel", "Remark_Bahasa"}.issubset(ref_df.columns):
            print("Kolom referensi tidak lengkap")
            return df

        # Normalize key di kolom sementara supaya CODING_UNDEL asli tetap blank.
        # astype(str) langsung ke kolom asli akan mengubah None/NaN menjadi teks.
        def normalize_key(series: pd.Series) -> pd.Series:
            key = series.astype("string").str.strip()
            return key.mask(key.str.upper().isin(["", "NONE", "NAN", "NAT"]))

        df["__CODING_UNDEL_KEY"] = normalize_key(df["CODING_UNDEL"])
        ref_df["__CODING_UNDEL_KEY"] = normalize_key(ref_df["Coding Undel"])
        ref_df = ref_df.dropna(subset=["__CODING_UNDEL_KEY"])

        # Merge
        df = df.merge(
            ref_df[["__CODING_UNDEL_KEY", "Remark_Bahasa"]],
            how="left",
            on="__CODING_UNDEL_KEY"
        )

        df["REASON UNDEL"] = df["Remark_Bahasa"]

        # Cleanup
        df.drop(columns=["__CODING_UNDEL_KEY", "Remark_Bahasa"], inplace=True, errors="ignore")

    except Exception as e:
        print(f"Error in add_reason_undel: {e}")
    
    return df

def add_coding_delivery(df: pd.DataFrame, ref=r"\\192.168.9.76\D\RYAN\1. References\Table Reference.xlsx") -> pd.DataFrame:
    df = df.copy()
    
    if "CODING" not in df.columns:
        print("Kolom 'CODING' tidak ditemukan, tidak bisa menambahkan CODING_DELIVERY.")
        return df

    try:
        ref_df = pd.read_excel(ref, sheet_name="STATUS_CCC", header=1)
        if not {"CODE", "STATUS"}.issubset(ref_df.columns):
            print("Kolom referensi tidak lengkap")
            return df

        # Buat kolom default kosong
        df["CODING_DELIVERY"] = None  

        # Ambil hanya baris dengan CODING diawali 'D'
        mask_d = df["CODING"].astype(str).str.startswith("D", na=False)
        df_d = df.loc[mask_d].merge(
            ref_df[["CODE", "STATUS"]],
            how="left",
            left_on="CODING",
            right_on="CODE"
        )

        # Update kolom CODING_DELIVERY hanya untuk yang CODING diawali D
        df.loc[mask_d, "CODING_DELIVERY"] = df_d["STATUS"].values

    except Exception as e:
        print(f"Gagal menambahkan informasi CODING_DELIVERY: {e}")

    return df

def add_coding_remarks(df: pd.DataFrame, ref=r"\\192.168.9.76\D\RYAN\1. References\Table Reference.xlsx") -> pd.DataFrame:
    df = df.copy()
    
    if "CODING" not in df.columns:
        print("Kolom 'CODING' tidak ditemukan, tidak bisa menambahkan CODING_REMARKS.")
        return df

    try:
        ref_df = pd.read_excel(ref, sheet_name="STATUS_CCC", header=1)
        if not {"CODE", "STATUS"}.issubset(ref_df.columns):
            print("Kolom referensi tidak lengkap")
            return df

        # Buat kolom default kosong
        df["CODING_REMARKS"] = None  

        # Ambil semua baris dengan CODING yang tidak kosong
        df_d = df.loc[df["CODING"].notnull()].merge(
            ref_df[["CODE", "STATUS"]],
            how="left",
            left_on="CODING",
            right_on="CODE"
        )

        df.loc[df["CODING"].notnull(), "CODING_REMARKS"] = df_d["STATUS"].values

    except Exception as e:
        print(f"Gagal menambahkan informasi CODING_REMARKS: {e}")

    return df

def add_reason_1st_attempt(df: pd.DataFrame, ref_path: str, sheet_name: str = "Coding Undel") -> pd.DataFrame:
    df = df.copy()

    if "RESULT_1ST_ATTEMPT" not in df.columns:
        print("Kolom 'RESULT_1ST_ATTEMPT' tidak ditemukan, tidak bisa menambahkan REASON_1ST_ATTEMPT.")
        return df

    ref = pd.read_excel(ref_path, sheet_name=sheet_name)

    if "Coding Undel" not in ref.columns or "Remark_Bahasa" not in ref.columns:
        print("Sheet referensi tidak punya kolom 'Coding Undel' dan/atau 'Remark_Bahasa'")
        return df

    ref = ref[["Coding Undel", "Remark_Bahasa"]].drop_duplicates()

    df = df.merge(
        ref,
        left_on="RESULT_1ST_ATTEMPT",
        right_on="Coding Undel",
        how="left",
        validate="m:1"
    )

    df = df.drop(columns=["Coding Undel"])
    df = df.rename(columns={"Remark_Bahasa": "REASON_1ST_ATTEMPT"})

    # hanya isi jika diawali 'U'
    mask = df["RESULT_1ST_ATTEMPT"].astype(str).str.startswith("U")
    df.loc[~mask, "REASON_1ST_ATTEMPT"] = np.nan

    return df

def add_reason_2nd_attempt(df: pd.DataFrame, ref_path: str, sheet_name: str = "Coding Undel") -> pd.DataFrame:
    df = df.copy()

    if "RESULT_2ND_ATTEMPT" not in df.columns:
        print("Kolom 'RESULT_2ND_ATTEMPT' tidak ditemukan, tidak bisa menambahkan REASON_2ND_ATTEMPT.")
        return df

    # Load referensi
    ref = pd.read_excel(ref_path, sheet_name=sheet_name)

    # Pastikan ada kolom kunci di ref
    if "Coding Undel" not in ref.columns or "Remark_Bahasa" not in ref.columns:
        print("Sheet referensi tidak punya kolom 'Coding Undel' dan/atau 'Remark_Bahasa'")
        return df

    # Ambil hanya kolom kunci + Remark_Bahasa
    ref = ref[["Coding Undel", "Remark_Bahasa"]].drop_duplicates()

    # Join df.RESULT_2ND_ATTEMPT dengan ref["Coding Undel"]
    df = df.merge(ref, left_on="RESULT_2ND_ATTEMPT", right_on="Coding Undel", how="left", validate="m:1")

    # Drop kolom kunci dari ref biar tidak duplikat
    df = df.drop(columns=["Coding Undel"])

    # Rename Remark_Bahasa â†’ REASON_2ND_ATTEMPT
    df = df.rename(columns={"Remark_Bahasa": "REASON_2ND_ATTEMPT"})

    # hanya isi jika diawali 'U'
    mask = df["RESULT_2ND_ATTEMPT"].astype(str).str.startswith("U")
    df.loc[~mask, "REASON_2ND_ATTEMPT"] = np.nan

    return df

def add_reason_last_attempt(df: pd.DataFrame, ref_path: str, sheet_name: str = "Coding Undel") -> pd.DataFrame:
    df = df.copy()

    if "RESULT_LAST_ATTEMPT" not in df.columns:
        print("Kolom 'RESULT_LAST_ATTEMPT' tidak ditemukan, tidak bisa menambahkan REASON_LAST_ATTEMPT.")
        return df

    # Load referensi
    ref = pd.read_excel(ref_path, sheet_name=sheet_name)

    # Pastikan ada kolom kunci di ref
    if "Coding Undel" not in ref.columns or "Remark_Bahasa" not in ref.columns:
        print("Sheet referensi tidak punya kolom 'Coding Undel' dan/atau 'Remark_Bahasa'")
        return df

    # Ambil hanya kolom kunci + Remark_Bahasa
    ref = ref[["Coding Undel", "Remark_Bahasa"]].drop_duplicates()

    # Join df.RESULT_LAST_ATTEMPT dengan ref["Coding Undel"]
    df = df.merge(ref, left_on="RESULT_LAST_ATTEMPT", right_on="Coding Undel", how="left", validate="m:1")

    # Drop kolom kunci dari ref biar tidak duplikat
    df = df.drop(columns=["Coding Undel"])

    # Rename Remark_Bahasa â†’ REASON_LAST_ATTEMPT
    df = df.rename(columns={"Remark_Bahasa": "REASON_LAST_ATTEMPT"})

    # hanya isi jika diawali 'U'
    mask = df["RESULT_LAST_ATTEMPT"].astype(str).str.startswith("U")
    df.loc[~mask, "REASON_LAST_ATTEMPT"] = np.nan

    return df

