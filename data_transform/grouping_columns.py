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

from data_transform.location_columns import add_origin_city

def add_grouping_late(df: pd.DataFrame) -> pd.DataFrame:
    # ---------- Helper ----------
    def dur_hours(td):
        if pd.isna(td): return np.nan
        return td.total_seconds() / 3600
    def dur_days(td):
        if pd.isna(td): return np.nan
        return td.days
    def round_up(x):
        if pd.isna(x): return np.nan
        return math.ceil(x)
    def safe_str(x): return "" if pd.isna(x) else str(x)

    # ---------- Step by Step ----------
    df = add_origin_city(df)
    # ORG_DEST = perbandingan 3_LC_ORIGIN vs 3 LC DEST
    df["ORG_DEST"] = df["3_LC_ORIGIN"] == df["3 LC DEST"]

    # Combine OM HVO
    df["OM_HVO"] = np.where(
        df["OUTBOUND_MANIFEST_DATE"].isna() & df["ORG_DEST"],
        df["1ST_HVO_DATE"], df["OUTBOUND_MANIFEST_DATE"]
    )

    # Combine OM HVO HBG
    df["OM_HVO_HBG"] = np.where(df["OM_HVO"].isna(), df["HBG_DATE"], df["OM_HVO"])

    # Combine PU & Entry
    df["PU_ENTRY"] = np.where(
        (df["PICKUP_STATUS"]=="S01") & (df["PICKUP_DATE"] < df["TGL_ENTRY"]),
        df["PICKUP_DATE"], df["TGL_ENTRY"]
    )

    # Entry - OM
    df["ENTRY_TO_OM"] = (df["OM_HVO_HBG"] - df["TGL_ENTRY"]).apply(dur_hours).apply(round_up)

    # PU - HBG
    df["PU_HBG"] = (df["OM_HVO_HBG"] - df["PU_ENTRY"]).apply(dur_hours).apply(round_up)

    # Remarks ENTRY_OM
    df["REMARKS_ENTRY_OM"] = df["ENTRY_TO_OM"].apply(
        lambda x: None if pd.isna(x) else "FM - Late" if x > 5 else "FM - Ontime"
    )

    # Remarks PU_HBG
    df["REMARKS_PU_HBG"] = df["PU_HBG"].apply(
        lambda x: None if pd.isna(x) else "FM - Late" if x > 5 else "FM - Ontime"
    )

    # IM_HVI
    df["IM_HVI"] = np.where(
        df["INBOUND_MANIFEST_DATE"].isna() & df["ORG_DEST"],
        df["HVI_DATE"], df["INBOUND_MANIFEST_DATE"]
    )

    # IM_HVI_PRA
    df["IM_HVI_PRA"] = np.where(df["IM_HVI"].isna(), df["1ST_RUNSHEET_DATE"], df["IM_HVI"])

    # HBG_IM
    df["HBG_IM"] = (df["IM_HVI_PRA"] - df["OM_HVO_HBG"]).apply(dur_days)

    # Check LT-MM ETD
    df["CHECK_LT_MM_ETD"] = df.apply(
        lambda r: None if pd.isna(r["HBG_IM"])
        else "MM - Late ETD Max" if r["HBG_IM"] > r["ETD"]
        else "check MM", axis=1
    )

    # Check LT-MM
    df["CHECK_LT_MM"] = df["HBG_IM"].apply(
        lambda x: None if pd.isna(x) else "MM - Late > 1 Hari" if x > 1 else "MM - Ontime"
    )

    # REMARKS MM
    df["REMARKS_MM"] = df.apply(
        lambda r: None if pd.isna(r["CHECK_LT_MM_ETD"]) else
                  r["CHECK_LT_MM"] if r["CHECK_LT_MM_ETD"]=="check MM" else
                  r["CHECK_LT_MM_ETD"], axis=1
    )

    # JAM IM
    df["JAM_IM"] = pd.to_datetime(df["IM_HVI"], errors="coerce").dt.hour

    # REMARK IM 5 AB
    df["REMARK_IM_5_1"] = df.apply(
        lambda r: None if pd.isna(r["JAM_IM"]) else
                  "IM < 5 Sore AB" if (r["JAM_IM"]<17 and r["ZONA"] in ["A","B"]) else
                  "IM > 5 Sore AB", axis=1
    )
    # REMARK IM 12 CD
    df["REMARK_IM_5_2"] = df.apply(
        lambda r: None if pd.isna(r["JAM_IM"]) else
                  "IM < 12 Siang CD" if (r["JAM_IM"]<12 and r["ZONA"] in ["C","D"]) else
                  "IM > 12 Siang CD", axis=1
    )

    # RESULT JAM ZONA
    df["REMARK_IM_5"] = df.apply(
        lambda r: r["REMARK_IM_5_1"] if r["ZONA"] in ["A","B"]
        else r["REMARK_IM_5_2"] if r["ZONA"] in ["C","D"]
        else None, axis=1
    )

    # 1st Runsheet_IM HVI PRA
    df["1ST_RUNSHEET_HVI"] = (df["1ST_RUNSHEET_DATE"] - df["IM_HVI_PRA"]).apply(dur_days).apply(round_up)

    # REMARKS DAY
    def remark_day(x):
        if pd.isna(x): return None
        if x==0: return "Same Day"
        if x==1: return "Next Day"
        if x<0: return "Back Date"
        return "Another Day"
    df["REMARKS_DAY"] = df["1ST_RUNSHEET_HVI"].apply(remark_day)

    # REMARKS <5 Zona A
    df["REMARKS_<5_A"] = df.apply(
        lambda r: None if pd.isna(r["REMARK_IM_5"]) else
                  "LM - Ontime Zona AB" if (r["REMARK_IM_5_1"]=="IM < 5 Sore AB" and r["REMARKS_DAY"]=="Same Day" and r["ZONA"] in ["A","B"])
                  else "LM - Late Zona AB", axis=1
    )

    # REMARKS >5 A
    df["REMARKS_>5_A"] = df.apply(
        lambda r: None if pd.isna(r["REMARK_IM_5"]) else
                  "LM - IM Backdate" if r["REMARKS_DAY"]=="Back Date" else
                  "LM - Ontime Zona AB" if (r["REMARK_IM_5"]=="IM > 5 Sore AB" and r["ZONA"] in ["A","B"] and r["REMARKS_DAY"] in ["Same Day","Next Day"])
                  else "LM - Late Zona AB", axis=1
    )

    # REMARKS FINAL ZONA A
    df["REMARKS_FINAL_ZONA_A"] = df.apply(
        lambda r: None if pd.isna(r["REMARK_IM_5"]) else
                  r["REMARKS_<5_A"] if r["REMARK_IM_5"] in ["IM < 5 Sore AB","IM < 12 Siang CD"] else
                  r["REMARKS_>5_A"], axis=1
    )

    # MTS_IM
    df["MTS_IM"] = (df["MANIFEST_TRANSIT_SUBAGEN_DATE"] - df["IM_HVI_PRA"]).apply(dur_days)

    # REMARKS DAY MTS
    def remark_day_mts(x):
        if pd.isna(x): return None
        if x==0: return "Same Day"
        if x==1: return "Next Day"
        if x<0: return "LM - IM Backdate"
        return "Another Day"
    df["REMARKS_DAY_MTS"] = df["MTS_IM"].apply(remark_day_mts)

    # REMARKS <5 C D
    df["REMARKS_<5_CD"] = df.apply(
        lambda r: None if pd.isna(r["REMARK_IM_5"]) else
                  "LM - IM Backdate" if r["REMARKS_DAY_MTS"]=="LM - IM Backdate" else
                  "Ontime Cabang Utama FW Sub Agen" if (r["REMARK_IM_5"]=="IM < 12 Siang CD" and r["REMARKS_DAY_MTS"]=="Same Day") else
                  "LM - Late FW to Sub Agen", axis=1
    )

    # REMARKS >5 C D
    df["REMARKS_>5_CD"] = df.apply(
        lambda r: None if pd.isna(r["REMARK_IM_5"]) else
                  "LM - IM Backdate" if r["REMARKS_DAY_MTS"]=="LM - IM Backdate" else
                  "Ontime Cabang Utama FW Sub Agen" if (r["REMARK_IM_5"]=="IM > 12 Siang CD" and r["REMARKS_DAY_MTS"] in ["Same Day","Next Day"]) else
                  "LM - Late FW to Sub Agen", axis=1
    )

    # FINAL MTS
    df["FINAL_MTS"] = np.where(df["REMARK_IM_5"]=="IM < 5 Sore", df["REMARKS_<5_CD"], df["REMARKS_>5_CD"])

    # FINAL SUBAGEN
    df["FINAL_FW_SUBAGEN"] = np.where(df["ZONA"].isin(["A","B"]), df["REMARKS_FINAL_ZONA_A"], df["FINAL_MTS"])

    # MTS_MTI
    df["MTS_MTI"] = (df["MANIFEST_INBOUND_SUBAGEN_DATE"] - df["MANIFEST_TRANSIT_SUBAGEN_DATE"]).apply(dur_days)

    # REMARK_MTS_MTI
    def remark_mts_mti(r):
        if pd.isna(r): return None
        if r<0: return "LM - MTI Backdate"
        if r>=2: return "LM - Late Inbound Sub Agen"
        return "Ontime Inbound Sub Agen"
    df["REMARK_MTS_MTI"] = df["MTS_MTI"].apply(remark_mts_mti)

    # FINAL FW MTI
    df["FINAL_FW_MTI"] = df.apply(
        lambda r: None if pd.isna(r["FINAL_FW_SUBAGEN"]) else
                  r["REMARK_MTS_MTI"] if r["FINAL_FW_SUBAGEN"]=="Ontime Cabang Utama FW Sub Agen" and pd.notna(r["REMARK_MTS_MTI"]) else
                  r["FINAL_FW_SUBAGEN"], axis=1
    )

    # JAM MTI
    df["JAM_MTI"] = pd.to_datetime(df["MANIFEST_INBOUND_SUBAGEN_DATE"], errors="coerce").dt.hour

    # REMARKS MTI 5
    df["REMARKS_MTI_5"] = df["JAM_MTI"].apply(
        lambda x: None if pd.isna(x) else "MTI < 5 Sore" if x==17 else "MTI > 5 Sore"
    )

    # 1St Runsheet_MTI
    df["1ST_RUNSHEET_MTI"] = (df["1ST_RUNSHEET_DATE"] - df["MANIFEST_INBOUND_SUBAGEN_DATE"]).apply(dur_days)

    # REMARKS DAY MTI
    def remark_day_mti(x):
        if pd.isna(x): return None
        if x==0: return "Same Day"
        if x==1: return "Next Day"
        if x<0: return "LM - MTI Backdate"
        return "Another Day"
    df["REMARKS_DAY_MTI"] = df["1ST_RUNSHEET_MTI"].apply(remark_day_mti)

    # REMARKS <5 MTI
    df["REMARKS_<5_MTI"] = df.apply(
        lambda r: None if pd.isna(r["REMARKS_MTI_5"]) else
                  "LM - MTI Backdate" if r["REMARKS_DAY_MTI"]=="LM - MTI Backdate" else
                  "LM - Ontime Runsheet Zona C D" if (r["REMARKS_MTI_5"]=="MTI < 5 Sore" and r["REMARKS_DAY_MTI"]=="Same Day") else
                  "LM - Late Runsheet Zona C D", axis=1
    )

    # REMARKS >5 MTI
    df["REMARKS_>5_MTI"] = df.apply(
        lambda r: None if pd.isna(r["REMARKS_MTI_5"]) else
                  "LM - MTI Backdate" if r["REMARKS_DAY_MTI"]=="LM - MTI Backdate" else
                  "LM - Ontime Runsheet Zona C D" if (r["REMARKS_MTI_5"]=="MTI > 5 Sore" and r["REMARKS_DAY_MTI"] in ["Same Day","Next Day"]) else
                  "LM - Late Runsheet Zona C D", axis=1
    )

    # FINAL MTI
    df["FINAL_MTI"] = np.where(df["REMARKS_MTI_5"]=="MTI < 5 Sore", df["REMARKS_<5_MTI"], df["REMARKS_>5_MTI"])

    # FINAL MTI SUBAGEN
    df["FINAL_MTI_SUBAGEN"] = np.where(
        df["FINAL_FW_MTI"].isna(), df["FINAL_MTI"],
        np.where(df["FINAL_FW_MTI"]=="Ontime Inbound Sub Agen", df["FINAL_MTI"], df["FINAL_FW_MTI"])
    )

    # FINAL COMBINE MTI SUBAGEN
    df["FINAL_COMBINE_MTI_SUBAGEN"] = np.where(
        df["FINAL_MTI_SUBAGEN"].isna() & (df["FINAL_FW_MTI"]=="Ontime Cabang Utama FW Sub Agen"),
        "Un MTI", df["FINAL_MTI_SUBAGEN"]
    )

    # FINAL COMBINE LASTMILE
    df["FINAL_COMBINE_LASTMILE"] = np.where(
        (df["FINAL_FW_MTI"]=="Ontime Cabang Utama FW Sub Agen") & (df["FINAL_MTI"].isna()),
        "LM - UnMTI", df["FINAL_COMBINE_MTI_SUBAGEN"]
    )

    # FM / MM / LM flags
    df["FM"] = df["REMARKS_ENTRY_OM"].apply(lambda x: None if pd.isna(x) or "Ontime" in str(x) else x)
    df["MM"] = df["REMARKS_MM"].apply(lambda x: None if pd.isna(x) or "Ontime" in str(x) else x)
    df["LM"] = df["FINAL_COMBINE_LASTMILE"].apply(lambda x: None if pd.isna(x) or "Late" not in str(x) else x)

    # Un-series
    df["Un Rcc"] = df["RECEIVING_DATE"].apply(lambda x: "FM - Unreceiving" if pd.isna(x) else None)
    df["Un Om"] = df["OM_HVO_HBG"].apply(lambda x: "FM - Un Manifest" if pd.isna(x) else None)
    df["Un Im"] = df["IM_HVI_PRA"].apply(lambda x: "MM - Un Inbound" if pd.isna(x) else None)
    df["Un MTS"] = df["MANIFEST_TRANSIT_SUBAGEN_DATE"].apply(lambda x: "LM - Un MTS" if pd.isna(x) else None)
    df["Un MTI"] = df["MANIFEST_INBOUND_SUBAGEN_DATE"].apply(lambda x: "LM - Un MTI" if pd.isna(x) else None)
    df["Un Runsheet"] = df["1ST_RUNSHEET_DATE"].apply(lambda x: "LM - Un Runsheet" if pd.isna(x) else None)

    # Aging Delivery Courier
    df["AGING_DELIVERY_COURIER"] = (df["TGL_RECEIVED"] - df["1ST_RUNSHEET_DATE"]).apply(dur_days)
    df["POD_COURIER"] = df["AGING_DELIVERY_COURIER"].apply(
        lambda x: None if pd.isna(x) else "LM - Late Update POD" if x > 1 else None
    )

    # FD Undel
    df["FD_UNDEL"] = df["DATE_LAST_ATTEMPT"].apply(lambda x: "LM - Late FD (Undel)" if pd.notna(x) else None)

    # Cancel
    df["CANCEL"] = df["STATUS_POD"].apply(
        lambda x: None if pd.isna(x) else "FM - AWB Cancel" if "Cancel" in str(x) else None
    )

    # FW
    df["FW"] = df["NO_CNOTE_FW"].apply(
        lambda x: None if pd.isna(x) else "LM - Problem Misroute" if "FW" in str(x) else None
    )

    # Cukai
    df["CUKAI"] = df["HOLD_REASON"].apply(
        lambda x: None if pd.isna(x) else "MM - Late Problem Bea Cukai"
        if ("Cukai" in str(x) or "Bea" in str(x)) else None
    )

    cols = ["FM","MM","LM","Un Rcc","Un Om","Un Im","Un MTS","Un MTI",
        "Un Runsheet","CARRER_POD","POD_COURIER","FD_UNDEL","CANCEL","FW","CUKAI"]

    df["GABUNGAN_ANALISA"] = (
        df[cols]
        .apply(lambda col: col.map(safe_str)) 
        .agg(", ".join, axis=1)
    )

    # Flagging FM/MM/LM late
    df["FM_LATE"] = df["GABUNGAN_ANALISA"].apply(
        lambda x: "FM - LATE" if "FM" in x else None
    )
    df["MM_LATE"] = df["GABUNGAN_ANALISA"].apply(
        lambda x: "MM - LATE" if "MM" in x else None
    )
    df["LM_LATE"] = df["GABUNGAN_ANALISA"].apply(
        lambda x: "LM - LATE" if "LM" in x else None
    )

    df["GROUPING_LATE"] = (
        df[["FM_LATE", "MM_LATE", "LM_LATE"]]
        .apply(lambda row: ", ".join([v for v in row if v not in [None, ""]]), axis=1)
    )

    # Hanya isi kalau CARRER_POD == "Over SLA", selain itu kosong
    df.loc[~df["CARRER"].isin(["Over SLA", "Over SLA (became)"]), "GROUPING_LATE"] = ""
    df.loc[(df["CARRER"] == "Over SLA") & (df["GROUPING_LATE"].isna() | (df["GROUPING_LATE"] == "")), "GROUPING_LATE"] = "LM - LATE"
    
    # Bersihkan kolom bantu
    df = df.drop(
        columns=[
            # gabungan akhir
            "GABUNGAN_ANALISA","FM_LATE","MM_LATE","LM_LATE",
            # remark FM/MM/LM awal
            "FM","MM","LM","Un Rcc","Un Om","Un Im","Un MTS","Un MTI","Un Runsheet","POD_COURIER","FD_UNDEL","CANCEL","FW","CUKAI",
            "FINAL_COMBINE_LASTMILE","FINAL_COMBINE_MTI_SUBAGEN","FINAL_MTI_SUBAGEN",
            "FINAL_MTI","REMARKS_>5_MTI","REMARKS_<5_MTI","REMARKS_DAY_MTI",
            "1ST_RUNSHEET_MTI","REMARKS_MTI_5","JAM_MTI","FINAL_FW_MTI",
            "REMARK_MTS_MTI","MTS_MTI","FINAL_FW_SUBAGEN",
            "POD_COURIER","AGING_DELIVERY_COURIER",
            # blok ORG_DEST sampai FINAL_MTS
            "ORG_DEST","OM_HVO","OM_HVO_HBG","PU_ENTRY","ENTRY_TO_OM","PU_HBG",
            "REMARKS_ENTRY_OM","REMARKS_PU_HBG",
            "IM_HVI","IM_HVI_PRA","HBG_IM","CHECK_LT_MM_ETD","CHECK_LT_MM",
            "REMARKS_MM","JAM_IM","REMARK_IM_5_1","REMARK_IM_5_2","REMARK_IM_5",
            "1ST_RUNSHEET_HVI","REMARKS_DAY","REMARKS_<5_A","REMARKS_>5_A",
            "REMARKS_FINAL_ZONA_A","MTS_IM","REMARKS_DAY_MTS",
            "REMARKS_<5_CD","REMARKS_>5_CD","FINAL_MTS"
        ],
        errors="ignore"
    )

    return df

def add_grouping_sla(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    if "SLA" not in df.columns:
        print("Kolom 'SLA' tidak ditemukan, tidak bisa menambahkan GROUPING_SLA.")
        return df

    def categorize_sla(val):
        try:
            v = int(val)
        except (ValueError, TypeError):
            return ""
        
        if v == 0:
            return "TODAY"
        elif v == -1:
            return "H-1"
        elif v == -2:
            return "H-2"
        elif v == -3:
            return "H-3"
        elif v <= -4:
            return ">H-4"
        elif v >= 1:
            return "OVER SLA"
        else:
            return ""

    df["GROUPING_SLA"] = df["SLA"].apply(categorize_sla)
    return df

def add_grouping_status(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    if "STATUS_POD" not in df.columns:
        print("Kolom 'STATUS_POD' tidak ditemukan, tidak bisa menambahkan GROUPING_STATUS.")
        return df

    def categorize_status_pod(val):
        try:
            v = str(val).strip()
        except (ValueError, TypeError):
            return ""
        
        if v == "Success":
            return "Success"
        elif v in ["Return Shipper", "RU Origin"]:
            return "Process Return"
        elif v in ["Missing", "AWB Cancel", "Hold WH Destination", "Undel"]:
            return "Undel"
        else:
            return "On Process"

    df["GROUPING_STATUS"] = df["STATUS_POD"].apply(categorize_status_pod)
    return df

