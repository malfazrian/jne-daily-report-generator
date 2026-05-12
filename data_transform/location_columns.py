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

def add_origin_city(df: pd.DataFrame, ref=r"\\192.168.9.76\D\RYAN\1. References\Table Reference.xlsx") -> pd.DataFrame:
    # Pastikan kolom ORIGIN ada
    if "ORIGIN" not in df.columns:
        return df

    try:
        # Baca sheet ZONA dari file referensi
        ref_df = pd.read_excel(ref, sheet_name="ZONA")

        # Pastikan kolom referensi lengkap
        if not {"DEST", "3 LC", "NAMA KAB/KOTA", "NAMA KAB/KOTA 2", "REGION", "PROVINSI", "KODE POS"}.issubset(ref_df.columns):
            print("Sheet ZONA tidak memiliki kolom DEST, 3 LC, NAMA KAB/KOTA, NAMA KAB/KOTA 2, REGION, PROVINSI atau KODE POS.")
            return df

        # Join left: df.ORIGIN -> ref.DEST (bedakan suffix)
        df = df.merge(
            ref_df[["DEST", "3 LC", "NAMA KAB/KOTA", "NAMA KAB/KOTA 2", "REGION", "PROVINSI", "KODE POS"]],
            how="left",
            left_on="ORIGIN",
            right_on="DEST",
            suffixes=("", "_REF")
        )

        # Rename hasil join
        df.rename(
            columns={
                "NAMA KAB/KOTA": "ORIGIN_CITY",
                "NAMA KAB/KOTA 2_REF": "ORIGIN_CITY_2",
                "PROVINSI": "PROVINSI_ORIGIN",
                "3 LC": "3_LC_ORIGIN",
                "KODE POS": "KODE_POS_ORI"
            },
            inplace=True
        )

        # Drop kolom DEST dari referensi saja
        df.drop(columns=["DEST_REF"], inplace=True, errors="ignore")

    except Exception as e:
        print(f"Gagal menambahkan 3_LC_ORIGIN, ORIGIN_CITY, ORIGIN_CITY_2, REGION, PROVINSI_ORIGIN dan KODE_POS_ORI: {e}")

    return df

def add_province_zipcode(
    df: pd.DataFrame,
    ref=r"\\192.168.9.76\D\RYAN\1. References\Table Reference.xlsx",
    debug: bool = False
) -> pd.DataFrame:
    df = df.copy()

    if "DEST" not in df.columns:
        print("Sheet MASTER tidak memiliki kolom DEST.")
        return df

    try:
        ref_df = pd.read_excel(ref, sheet_name="ZONA")

        if not {"DEST", "KECAMATAN", "NAMA KAB/KOTA", "PROVINSI", "KODE POS"}.issubset(ref_df.columns):
            print("Sheet ZONA tidak memiliki kolom DEST, KECAMATAN, NAMA KAB/KOTA, PROVINSI atau KODE POS.")
            return df

        # Pastikan sama-sama string dan tanpa spasi
        df["DEST"] = df["DEST"].astype(str).str.strip().str.upper()
        ref_df["DEST"] = ref_df["DEST"].astype(str).str.strip().str.upper()

        # Hilangkan duplikat di referensi
        ref_df_unique = ref_df.drop_duplicates(subset=["DEST"])[["DEST", "KECAMATAN", "NAMA KAB/KOTA", "PROVINSI", "KODE POS"]]

        # Merge berdasarkan DEST
        df = df.merge(
            ref_df_unique,
            how="left",
            on="DEST"
        )

        # Normalisasi nilai kosong di hasil merge (agar '' / 'nan' ikut dianggap kosong)
        if "KODE POS" in df.columns:
            df["KODE POS"] = (
                df["KODE POS"]
                .astype("string")
                .str.strip()
                .str.replace(r"\.0+$", "", regex=True)
                .replace({"": pd.NA, "NAN": pd.NA, "nan": pd.NA, "None": pd.NA, "NONE": pd.NA, "0": pd.NA, "0.0": pd.NA})
            )

        # Fallback KODE POS: ambil dari DEST terdekat berikutnya (forward)
        # Contoh: CKR10000 -> CKR10001, jika kosong lanjut ke CKR10020, dst.
        kode_pos_clean = (
            ref_df_unique["KODE POS"]
            .astype("string")
            .str.strip()
            .str.replace(r"\.0+$", "", regex=True)
            .replace({"": pd.NA, "NAN": pd.NA, "nan": pd.NA, "None": pd.NA, "NONE": pd.NA, "0": pd.NA, "0.0": pd.NA})
        )

        ref_tmp = ref_df_unique[["DEST"]].copy()
        ref_tmp["KODE_POS_REF"] = kode_pos_clean
        ref_extract = ref_tmp["DEST"].str.extract(r"^([A-Z]+)\s*(\d+)$")
        ref_tmp["DEST_PREFIX"] = ref_extract[0]
        ref_tmp["DEST_NUM"] = pd.to_numeric(ref_extract[1], errors="coerce")

        # Source DEST yang benar-benar punya KODE POS
        ref_zip = ref_tmp[
            ref_tmp["KODE_POS_REF"].notna() &
            ref_tmp["DEST_PREFIX"].notna() &
            ref_tmp["DEST_NUM"].notna()
        ][["DEST_PREFIX", "DEST_NUM", "KODE_POS_REF"]].sort_values(["DEST_PREFIX", "DEST_NUM"])

        if not ref_zip.empty:
            ref_zip["DEST_NUM"] = ref_zip["DEST_NUM"].astype("int64")

        if not ref_zip.empty and "KODE POS" in df.columns:
            df_extract = df["DEST"].str.extract(r"^([A-Z]+)\s*(\d+)$")
            df_fallback = df[["DEST", "KODE POS"]].copy()
            df_fallback["__row_id"] = df_fallback.index
            df_fallback["DEST_PREFIX"] = df_extract[0]
            df_fallback["DEST_NUM"] = pd.to_numeric(df_extract[1], errors="coerce")

            # Hanya untuk baris KODE POS kosong + DEST yang valid format prefix+angka
            mask_need_fill = (
                df_fallback["KODE POS"].isna() &
                df_fallback["DEST_PREFIX"].notna() &
                df_fallback["DEST_NUM"].notna()
            )

            if mask_need_fill.any():
                left = df_fallback.loc[mask_need_fill, ["__row_id", "DEST_PREFIX", "DEST_NUM"]].copy()
                left["DEST_NUM"] = left["DEST_NUM"].astype("int64")

                right = ref_zip[["DEST_PREFIX", "DEST_NUM", "KODE_POS_REF"]].copy()
                right["DEST_NUM"] = right["DEST_NUM"].astype("int64")

                # Bangun index per prefix agar fallback selalu ambil DEST >= current DEST terdekat
                right_groups = {}
                for prefix, grp in right.groupby("DEST_PREFIX", sort=False):
                    grp_sorted = grp.sort_values("DEST_NUM")
                    right_groups[prefix] = (
                        grp_sorted["DEST_NUM"].to_numpy(),
                        grp_sorted["KODE_POS_REF"].to_numpy()
                    )

                fill_pairs = []
                for row_id, prefix, cur_num in left.itertuples(index=False, name=None):

                    if prefix not in right_groups:
                        continue

                    nums, zips = right_groups[prefix]
                    pos = np.searchsorted(nums, cur_num, side="left")
                    if pos < len(nums):
                        fill_pairs.append((row_id, zips[pos]))

                if fill_pairs:
                    fill_map = pd.Series({rid: z for rid, z in fill_pairs}, name="KODE_POS_REF")
                    df.loc[fill_map.index, "KODE POS"] = df.loc[fill_map.index, "KODE POS"].fillna(fill_map)

        # Final normalize: paksa KODE POS tetap teks dan buang akhiran .0
        if "KODE POS" in df.columns:
            df["KODE POS"] = (
                df["KODE POS"]
                .astype("string")
                .str.strip()
                .str.replace(r"\.0+$", "", regex=True)
                .replace({"": pd.NA, "NAN": pd.NA, "nan": pd.NA, "None": pd.NA, "NONE": pd.NA, "0": pd.NA, "0.0": pd.NA})
            )

        if debug and "KODE POS" in df.columns:
            missing_after = int(df["KODE POS"].isna().sum())
            total_rows = len(df)
            print(f"[add_province_zipcode] Total rows: {total_rows}")
            print(f"[add_province_zipcode] Missing KODE POS after fill: {missing_after}")

            # Tampilkan sample baris yang masih kosong untuk investigasi format DEST
            if missing_after > 0:
                sample_missing = df.loc[df["KODE POS"].isna(), ["DEST"]].head(10)
                print("[add_province_zipcode] Sample DEST yang masih kosong KODE POS:")
                print(sample_missing.to_string(index=False))

    except Exception as e:
        print(f"Gagal menambahkan KECAMATAN, NAMA KAB/KOTA, PROVINSI dan KODE POS: {e}")

    return df

def add_wilayah(
    df: pd.DataFrame,
    ref: str = r"\\192.168.9.76\D\RYAN\1. References\Table Reference.xlsx"
) -> pd.DataFrame:
    """
    Menambahkan kolom WILAYAH berdasarkan DEST dan referensi REGION (sheet ZONA)
    """

    df = df.copy()

    # ===============================
    # Validasi kolom utama
    # ===============================
    if "DEST" not in df.columns:
        print("Sheet MASTER tidak memiliki kolom DEST.")
        return df

    try:
        # ===============================
        # Load reference
        # ===============================
        ref_df = pd.read_excel(ref, sheet_name="ZONA")

        required_cols = {"DEST", "REGION"}
        if not required_cols.issubset(ref_df.columns):
            print("Sheet ZONA tidak memiliki kolom DEST atau REGION.")
            return df

        # ===============================
        # Normalisasi data
        # ===============================
        df["DEST"] = df["DEST"].astype(str).str.strip()
        ref_df["DEST"] = ref_df["DEST"].astype(str).str.strip()

        # ===============================
        # Ambil DEST unik
        # ===============================
        ref_df = (
            ref_df[["DEST", "REGION"]]
            .drop_duplicates(subset="DEST")
            .rename(columns={"REGION": "REGION_REF"})
        )

        # ===============================
        # Merge
        # ===============================
        df = df.merge(ref_df, how="left", on="DEST")

        # ===============================
        # Mapping wilayah
        # ===============================
        wilayah_map = {
            "Regional Jakarta": "Jabodetabekcil",
            "Regional Bodetabekcil": "Jabodetabekcil",
            "Regional Jawa Barat": "Pulau Jawa",
            "Regional Jateng Diy": "Pulau Jawa",
        }

        df["WILAYAH"] = (
            df["REGION_REF"]
            .map(wilayah_map)
            .fillna("Luar Pulau Jawa")
        )

        # ===============================
        # Cleanup
        # ===============================
        df.drop(columns=["REGION_REF"], inplace=True)

    except Exception as e:
        print(f"Gagal menambahkan WILAYAH: {e}")

    return df

def add_eta(
    df: pd.DataFrame, 
    holiday_path=r"\\192.168.9.76\D\RYAN\1. References\Holiday.xlsx", 
    holidayExcl=True
) -> pd.DataFrame:
    df = df.copy()
    
    if "ETD" not in df.columns or "TGL_ENTRY" not in df.columns:
        print("Kolom 'ETD' atau 'TGL_ENTRY' tidak ditemukan, tidak bisa menambahkan ETA.")
        return df

    # Pastikan TGL_ENTRY datetime
    df["TGL_ENTRY"] = pd.to_datetime(df["TGL_ENTRY"], errors="coerce")

    # Default tidak ada holidays
    holidays = []
    if holidayExcl:
        try:
            holidays = pd.read_excel(
                holiday_path, 
                sheet_name="Holiday", 
                usecols=[0], 
                header=None
            )[0]
            holidays = pd.to_datetime(holidays, errors="coerce").dropna().tolist()
        except Exception as e:
            print(f"Gagal load Holiday.xlsx: {e}")
            holidays = []

    try:
        df["ETD_days"] = pd.to_numeric(df["ETD"], errors="coerce")

        if holidayExcl:
            # Custom business day offset (skip weekend + holiday)
            cbd = CustomBusinessDay(holidays=holidays)
            df["ETA"] = df.apply(
                lambda row: row["TGL_ENTRY"] + cbd * (int(row["ETD_days"]) + 1)
                if pd.notna(row["TGL_ENTRY"]) and pd.notna(row["ETD_days"]) else pd.NaT,
                axis=1
            )
        else:
            # Kalender normal (termasuk weekend & libur)
            df["ETA"] = df["TGL_ENTRY"] + pd.to_timedelta(df["ETD_days"] + 1, unit="D")

    except Exception as e:
        print(f"Gagal menghitung ETA: {e}")
        df["ETA"] = pd.NaT

    df.drop(columns=["ETD_days"], inplace=True, errors="ignore")
    
    return df

