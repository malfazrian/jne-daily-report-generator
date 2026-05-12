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

def add_status_pod_c01(df: pd.DataFrame, debug: bool = False) -> pd.DataFrame:
    df = df.copy()

    if "STATUS_POD" not in df.columns:
        raise ValueError("Missing required columns: ['STATUS_POD']")
    if "CODING" not in df.columns:
        df["CODING"] = pd.NA
    if "STATUS_POD_UPDATE" not in df.columns:
        df["STATUS_POD_UPDATE"] = df["STATUS_POD"]
    
    if debug:
        print("Kolom awal:", df.columns.tolist())
        print(f"STATUS_POD_UPDATE unique values: {df['STATUS_POD_UPDATE'].unique().tolist()}")

    # Tambahkan kolom STATUS_POD_UPDATE berdasarkan logika khusus untuk CODING 'C01' dan 'C02'
    def status_pod_update(row):
        try:
            coding = str(row["CODING"]).strip() if pd.notna(row["CODING"]) else None
            status_pod = row["STATUS_POD_UPDATE"] if pd.notna(row["STATUS_POD_UPDATE"]) else ""
            if coding == "C01":
                return "Claim Missing"
            elif coding == "C02":
                return "Claim Damage"
            else:
                return status_pod
        except Exception as e:
            print(f"Error in status_pod_update for row: {e}")
            return "Error"
    try:
        df["STATUS_POD_UPDATE"] = df.apply(status_pod_update, axis=1)
        if debug: print("STATUS_POD_UPDATE done")
        if debug: print(f"STATUS_POD_UPDATE unique values: {df['STATUS_POD_UPDATE'].unique().tolist()}")
    except Exception as e:
        raise ValueError(f"Error in STATUS_POD_UPDATE: {e}")
    
    # -----------------------------
    # Replace STATUS_POD with STATUS_POD_UPDATE
    # -----------------------------
    if "STATUS_POD_UPDATE" in df.columns:
        if "STATUS_POD" in df.columns:
            df = df.drop(columns=["STATUS_POD"])
        df = df.rename(columns={"STATUS_POD_UPDATE": "STATUS_POD"})

    return df

def add_status_pod(df: pd.DataFrame, debug: bool = False) -> pd.DataFrame:
    df = df.copy()

    # Check if required columns exist
    required_columns = [
        'CODING', 'AWB_CANCEL', 'TGL_ENTRY', 'CONNOTE_RETURN_RT', 'CONNOTE_RETURN_RF',
        'RESULT_1ST_ATTEMPT', 'STATUS_POD', 'INBOUND_MANIFEST', 'MANIFEST_TRANSIT_AGEN',
        'HVI_NO', 'RUNSHEET_NO', 'OUTBOUND_MANIFEST', 'RECEIVING', 'HVO_NO', 'ORIGIN',
        'DEST', 'CONFIRM_SHIPMENT_UNDEL', 'DATE_RUNSHEET'
    ]
    optional_columns = ['NO_CNOTE_FW', 'DEST_FW', 'CODING_STATUS_FW', 'DESC_STATUS_FW', 'IREG_CODE']
    missing_required = [col for col in required_columns if col not in df.columns]
    if missing_required:
        raise ValueError(f"Missing required columns: {missing_required}")

    # Add missing optional columns with pd.NA
    for col in optional_columns:
        if col not in df.columns:
            df[col] = pd.NA
            if debug:
                print(f"Added missing optional column {col} with pd.NA")

    if "STATUS_POD_UPDATE" in df.columns:
        # Terapkan STATUS_POD_UPDATE ke STATUS_POD sebelum dibuang,
        # agar nilai seperti "Return Shipper" tidak hilang saat cek_d24_d25 membaca STATUS_POD.
        mask = df["STATUS_POD_UPDATE"].notna() & (df["STATUS_POD_UPDATE"].astype(str).str.strip() != "")
        df.loc[mask, "STATUS_POD"] = df.loc[mask, "STATUS_POD_UPDATE"]
        df = df.drop(columns=["STATUS_POD_UPDATE"])

    if debug:
        print("Kolom awal:", df.columns.tolist())
        print(f"STATUS_POD unique values: {df['STATUS_POD'].unique().tolist()}")

    # -----------------------------
    # CEK D24 D25
    # -----------------------------
    def cek_d24_d25(row):
        try:
            coding = str(row["CODING"]).strip() if pd.notna(row["CODING"]) else None
            if coding in ["D25", "D26", "D99"]:
                return "Auto Closed Breach (D25,D26)"
            elif coding == "D24":
                return "Claim Damage"
            elif coding == "C01":
                return "Claim Missing"
            elif coding == "C02":
                return "Claim Damage"
            elif coding == "R26":
                return "Return Shipper"
            elif coding == "CL1":
                return "Auto Closed Origin (CL1)"
            elif coding in ["CL2", "CL3", "CL4"]:
                return "Auto Closed Destination (CL2,CL4)"
            elif coding in ["R24", "R25"]:
                return "Rejection Return"
            elif coding == "CR4":
                return "Destroyed"
            elif pd.isna(row["CODING"]):
                # Map RU Shipper/Origin to RU Origin
                status_pod = row["STATUS_POD"] if pd.notna(row["STATUS_POD"]) else "Unknown"
                return "RU Origin" if status_pod == "RU Shipper/Origin" else status_pod
            else:
                status_pod = row["STATUS_POD"] if pd.notna(row["STATUS_POD"]) else "Unknown"
                return "RU Origin" if status_pod == "RU Shipper/Origin" else status_pod
        except Exception as e:
            print(f"Error in cek_d24_d25 for row: {e}")
            return "Error"
        
    # helper untuk conditional kolom
    def add_column(df, column_name, func):
        try:
            df[column_name] = df.apply(func, axis=1)
            return df
        except Exception as e:
            raise ValueError(f"Error adding column {column_name}: {e}")

    try:
        df = add_column(df, "CEK_D24_D25", cek_d24_d25)
        if debug: print("CEK_D24_D25 done")
        if debug: print(f"CEK_D24_D25 unique values: {df['CEK_D24_D25'].unique().tolist()}")
    except Exception as e:
        raise ValueError(f"Error in CEK_D24_D25: {e}")

    # -----------------------------
    # Status1
    # -----------------------------
    def status1(row):
        try:
            # Map RU Shipper/Origin to RU Origin
            cek_d24_d25 = "RU Origin" if row["CEK_D24_D25"] == "RU Shipper/Origin" else row["CEK_D24_D25"]
            if cek_d24_d25 in [
                "Success", "Return Shipper", "RU Origin",
                "Auto Closed Breach (D25,D26)", "Claim Damage", "Claim Missing",
                "Auto Closed Origin (CL1)", "Auto Closed Destination (CL2,CL4)",
                "Rejection Return", "Destroyed"
            ]:
                return cek_d24_d25
            elif str(row["AWB_CANCEL"]).strip() == "Y":
                return "AWB Cancel"
            elif pd.isna(row["TGL_ENTRY"]):
                return "Invalid AWB"
            elif str(row["CODING"]).strip() == "CR8":
                return "RU Origin"
            elif str(row["CODING"]).strip() == "U14":
                return "Missing"
            elif str(row["CODING"]).strip() == "U11":
                return "Damage"
            elif cek_d24_d25 == "Un Receiving":
                return "Un Receiving"
            else:
                return "Next Cek"
        except Exception as e:
            print(f"Error in status1 for row: {e}")
            return "Error"

    try:
        df = add_column(df, "Status1", status1)
        if debug: print("Status1 done")
        if debug: print(f"Status1 unique values: {df['Status1'].unique().tolist()}")
    except Exception as e:
        raise ValueError(f"Error in Status1: {e}")

    # -----------------------------
    # Status2 - Status12
    # -----------------------------
    try:
        df["Status2"] = np.where(
            df["CONNOTE_RETURN_RT"].notna() & (df["Status1"] == "Next Cek"),
            "RU Origin", df["Status1"]
        )
        df["Status3"] = np.where(df["CONNOTE_RETURN_RT"].isna(), df["Status1"], df["Status2"])

        df["Status2RF"] = np.where(
            df["CONNOTE_RETURN_RF"].notna() & (df["Status3"] == "Next Cek"),
            "RU Origin", df["Status3"]
        )
        df["Status_3RF"] = np.where(df["CONNOTE_RETURN_RF"].isna(), df["Status3"], df["Status2RF"])

        df["Status4"] = np.where(
            (df["Status_3RF"] == "Next Cek") & df["CODING"].astype(str).str.contains("U0", na=False),
            "Undel", df["Status_3RF"]
        )
        df["Status5"] = np.where(df["CODING"].isna(), df["Status_3RF"], df["Status4"])

        df["Status6"] = np.where(
            (df["Status5"] == "Next Cek") & df["CODING"].astype(str).str.contains("CR3", na=False),
            "Undel", df["Status5"]
        )
        df["Status7"] = np.where(df["CODING"].isna(), df["Status5"], df["Status6"])

        df["Status8"] = np.where(
            df["CODING"].isin(["CR2", "U21", "U22", "U23", "U10"]),
            "Undel", df["Status7"]
        )
        df["Status9"] = np.where(df["Status7"] == "Next Cek", df["Status8"], df["Status7"])

        df["Status10"] = np.where(
            df["RESULT_1ST_ATTEMPT"].notna() & ~df["RESULT_1ST_ATTEMPT"].astype(str).str.contains("U12", na=False),
            "Undel", df["Status9"]
        )
        df["Status10_1"] = np.where(df["Status9"] != "Next Cek", df["Status9"], df["Status10"])

        df["Status11"] = np.where(
            (df["STATUS_POD"].notna()) | (df["INBOUND_MANIFEST"].notna()) |
            (df["MANIFEST_TRANSIT_AGEN"].notna()) | (df["HVI_NO"].notna()) |
            (df["RUNSHEET_NO"].notna()),
            "On Process", df["Status10_1"]
        )
        df["Status12"] = np.where(df["Status10"] == "Next Cek", df["Status11"], df["Status10_1"])
        if debug: print("Status2-Status12 done")
        if debug: print(f"Status12 unique values: {df['Status12'].unique().tolist()}")
    except Exception as e:
        raise ValueError(f"Error in Status2-Status12: {e}")

    # -----------------------------
    # Status13 - Status_18
    # -----------------------------
    try:
        df["3LC_DEST"] = df["DEST"].astype(str).str[:3]
        df["3LC_HVO"] = df["HVO_NO"].astype(str).str[:3]
        df["HVOSamaDest"] = df["3LC_DEST"] == df["3LC_HVO"]
        df["Status13"] = np.where(
            (df["Status12"] == "Next Cek") & (df["HVOSamaDest"] == True),
            "On Process", df["Status12"]
        )

        df["Status14"] = np.where(
            (df["OUTBOUND_MANIFEST"].notna()) | (df["HVO_NO"].notna()),
            "On Forward Destination", df["Status13"]
        )
        df["Status15"] = np.where(df["Status13"] != "Next Cek", df["Status13"], df["Status14"])

        df["Status16"] = np.where(df["RECEIVING"].notna(), "Un Manifest", df["Status15"])
        df["Status17"] = np.where(df["Status15"] != "Next Cek", df["Status15"], df["Status16"])

        df["Status_18"] = np.where(df["Status17"] == "Next Cek", "Un Receiving", df["Status17"])
        if debug: print("Status13–Status_18 done")
        if debug: print(f"Status_18 unique values: {df['Status_18'].unique().tolist()}")
    except Exception as e:
        raise ValueError(f"Error in Status13-Status_18: {e}")

    # -----------------------------
    # CONFIRM_SHIPMENT_UNDEL Check
    # -----------------------------
    try:
        df["Status_Update"] = np.where(
            df["CONFIRM_SHIPMENT_UNDEL"].astype(str).str.contains("WH1", na=False),
            "Hold WH Destination",
            np.where(
                df["Status_18"].isin([
                    "Success", "Return Shipper", "RU Origin", "AWB Cancel",
                    "Auto Closed Breach (D25,D26)", "Claim Damage", "Invalid AWB",
                    "Rejection Return", "Auto Closed Origin (CL1)",
                    "Auto Closed Destination (CL2,CL4)", "Destroyed"
                ]),
                df["Status_18"],
                df["Status_18"]
            )
        )
        if debug: print("CONFIRM_SHIPMENT_UNDEL check done")
        if debug: print(f"Status_Update unique values: {df['Status_Update'].unique().tolist()}")
    except Exception as e:
        raise ValueError(f"Error in CONFIRM_SHIPMENT_UNDEL check: {e}")

    # -----------------------------
    # CGK Intra-City Check
    # -----------------------------
    try:
        df["CGK_INTRA"] = np.where(
            (df["ORIGIN"].astype(str).str.contains("CGK", na=False)) &
            (df["DEST"].astype(str).str.contains("CGK", na=False)),
            "Y", "N"
        )
        df["CGK_INTRA2"] = np.where(df["ORIGIN"].isna(), pd.NA, df["CGK_INTRA"])
        df["Status_CGK"] = np.where(
            (df["Status_Update"] == "On Process") & (df["CGK_INTRA2"] == "Y"),
            "Harus di Cek", "Tidak Harus di Cek"
        )
        df["CGK_1"] = np.where(
            (df["Status_CGK"] == "Harus di Cek") & (df["RECEIVING"].isna()),
            "Belum Receiving", df["Status_Update"]
        )
        df["CGK_2_HVO"] = np.where(
            (df["Status_CGK"] == "Harus di Cek") & (df["HVO_NO"].isna()) & (df["RECEIVING"].notna()),
            "Sudah RCC", df["Status_Update"]
        )
        df["CGK3_HVI"] = np.where(
            (df["Status_CGK"] == "Harus di Cek") & (df["HVI_NO"].isna()) & (df["HVO_NO"].notna()),
            "Sudah HVO", df["Status_Update"]
        )
        df["CGK_HVI"] = np.where(
            (df["Status_CGK"] == "Harus di Cek") & (df["HVI_NO"].notna()),
            "Sudah HVI", df["Status_Update"]
        )
        df["Status_Val_CGK1"] = np.where(df["CGK_1"] == "Belum Receiving", "Un Receiving", df["Status_Update"])
        df["Status_Val_CGK2"] = np.where(df["CGK_2_HVO"] == "Sudah RCC", "Un Manifest", df["Status_Val_CGK1"])
        df["Status_Val_CGK3"] = np.where(df["CGK3_HVI"] == "Sudah HVO", "On Forward Destination", df["Status_Val_CGK2"])
        df["Status_Val_CGK4"] = np.where(df["CGK_HVI"] == "Sudah HVI", "On Process", df["Status_Val_CGK3"])
        if debug: print("CGK Intra-City check done")
        if debug: print(f"Status_Val_CGK4 unique values: {df['Status_Val_CGK4'].unique().tolist()}")
    except Exception as e:
        raise ValueError(f"Error in CGK Intra-City check: {e}")

    # -----------------------------
    # Forwarding (FW) Columns
    # -----------------------------
    try:
        df["CHECK_NO_FW"] = np.where(df["NO_CNOTE_FW"].notna(), "Y", "N")
        status_mapping = {
            "Success": "Success",
            "Return Shipper": "Return Shipper",
            "RU Shipper/Origin": "RU Origin",
            "Cancel": "Undel",
            "Undel": "Undel",
            "Destroyed": "Destroyed",
            "Damage Case": "Damage",
            "Missing": "Missing",
            "Rejection Return": "Rejection Return",
            "On Process": "On Process"
        }
        df["Status_FW_Bantu"] = df["DESC_STATUS_FW"].map(status_mapping).fillna("On Process").astype("object")
        df["Status_FW_Bantu"] = np.where(
            df["DESC_STATUS_FW"].isna(), "On Process",
            np.where(
                df["DESC_STATUS_FW"].isin(status_mapping.keys()),
                df["Status_FW_Bantu"],
                pd.NA
            )
        )
        df["NEW_STATUS"] = np.where(
            df["CHECK_NO_FW"] == "Y",
            np.where(
                df["Status_FW_Bantu"].isin([
                    "Success", "Return Shipper", "RU Origin", "Undel", "Damage",
                    "Missing", "Rejection Return", "Destroyed"
                ]),
                df["Status_FW_Bantu"],
                df["Status_Val_CGK4"]
            ),
            df["Status_Val_CGK4"]
        )
        df["DEST"] = np.where(df["CHECK_NO_FW"] == "Y", df["DEST_FW"], df["DEST"])
        df["CODING"] = np.where(df["NO_CNOTE_FW"].notna(), df["CODING_STATUS_FW"], df["CODING"])
        if debug: print("Forwarding columns done")
        if debug: print(f"Status_FW_Bantu unique values: {df['Status_FW_Bantu'].unique().tolist()}")
        if debug: print(f"NEW_STATUS unique values: {df['NEW_STATUS'].unique().tolist()}")
    except Exception as e:
        raise ValueError(f"Error in Forwarding columns: {e}")

    # -----------------------------
    # DB1 Check
    # -----------------------------
    try:
        df["DB1_Check"] = np.where(df["CODING"] == "DB1", "Success", df["NEW_STATUS"])
        if debug: print("DB1 Check done")
        if debug: print(f"DB1_Check unique values: {df['DB1_Check'].unique().tolist()}")
    except Exception as e:
        raise ValueError(f"Error in DB1 Check: {e}")

    # -----------------------------
    # CONNOTE_RETURN_RT Check
    # -----------------------------
    try:
        df["Check_RT"] = np.where(
            df["DB1_Check"].isin([
                "On Process", "Undel", "Hold WH Destination",
                "Auto Closed Destination (CL2,CL4)", "Auto Closed Breach (D25,D26)"
            ]) & (df["CONNOTE_RETURN_RT"].notna()),
            "RU Origin", df["DB1_Check"]
        )
        if debug: print("CONNOTE_RETURN_RT Check done")
        if debug: print(f"Check_RT unique values: {df['Check_RT'].unique().tolist()}")
    except Exception as e:
        raise ValueError(f"Error in CONNOTE_RETURN_RT Check: {e}")

    # -----------------------------
    # Cek_OP and Runsheet Checks
    # -----------------------------
    try:
        df["Cek_OP"] = np.where(df["Check_RT"] == "On Process", "Cek OP", df["Check_RT"])
        df["3LC_RUNSHEET"] = df["RUNSHEET_NO"].astype(str).str[:3]
        df["3LC_CEK_DEST"] = df["DEST"].astype(str).str[:3]
        df["3LC_CEK_ORG"] = df["ORIGIN"].astype(str).str[:3]
        df["RUNSHEET_DEST"] = df["3LC_RUNSHEET"] == df["3LC_CEK_DEST"]
        df["RUNSHEET_DEST_2"] = np.where(
            (df["RUNSHEET_DEST"] == False) & (df["3LC_RUNSHEET"].isna()), pd.NA, df["RUNSHEET_DEST"]
        )
        df["RUNSHEET_ORIGIN"] = df["3LC_RUNSHEET"] == df["3LC_CEK_ORG"]
        df["RUNSHEET_ORIGIN2"] = np.where(
            (df["RUNSHEET_ORIGIN"] == False) & (df["3LC_RUNSHEET"].isna()), pd.NA, df["RUNSHEET_ORIGIN"]
        )
        df["Cek_OP_2"] = np.where(
            (df["Cek_OP"] == "Cek OP") & (df["RUNSHEET_DEST_2"] == True),
            "Open POD", df["Cek_OP"]
        )
        df["Cek_OP_3"] = np.where(
            (df["Cek_OP"] == "Cek OP") & (df["RUNSHEET_DEST_2"] == False) & (df["RUNSHEET_ORIGIN2"] == True),
            "On Forward Destination", df["Cek_OP_2"]
        )
        df["Cek_OP_4"] = np.where(
            (df["Cek_OP"] == "Cek OP") & (df["RUNSHEET_DEST_2"] == False) & (df["RUNSHEET_ORIGIN2"] == False),
            "Problem Missroute", df["Cek_OP_3"]
        )
        df["Cek_OP_5"] = np.where(
            (df["Cek_OP"] == "Cek OP") & (df["3LC_RUNSHEET"].isna()),
            "Un Runsheet", df["Cek_OP_4"]
        )
        current_date = pd.to_datetime("today").normalize()
        df["DATE_RUNSHEET_2"] = pd.to_datetime(df["DATE_RUNSHEET"], errors="coerce").dt.normalize()
        df["RUNSHEET_VS_NOW"] = df["DATE_RUNSHEET_2"] == current_date
        df["Cek_OP_6"] = np.where(
            (df["Cek_OP_5"] == "Open POD") & (df["RUNSHEET_VS_NOW"] == True),
            "On Delivery", df["Cek_OP_5"]
        )
        if debug: print("Cek_OP and Runsheet checks done")
        if debug: print(f"Cek_OP_6 unique values: {df['Cek_OP_6'].unique().tolist()}")
    except Exception as e:
        raise ValueError(f"Error in Cek_OP and Runsheet checks: {e}")

    # -----------------------------
    # X-Ray Check
    # -----------------------------
    try:
        df["X_Ray_Cek1"] = np.where(
            df["IREG_CODE"].astype(str).str.startswith("X", na=False),
            "Problem X-Ray", pd.NA
        )
        df["Cek_OP_7"] = np.where(
            (df["X_Ray_Cek1"] == "Problem X-Ray") & 
            (df["Cek_OP_6"].isin(["Un Receiving", "Un Manifest", "On Forward Destination"])),
            "Problem X-Ray", df["Cek_OP_6"]
        )
        if debug: print("X-Ray Check done")
        if debug: print(f"Cek_OP_7 unique values: {df['Cek_OP_7'].unique().tolist()}")
    except Exception as e:
        raise ValueError(f"Error in X-Ray Check: {e}")

    # -----------------------------
    # Final Translation to STATUS_POD_UPDATE
    # -----------------------------
    def translate(value):
        try:
            if value == "Auto Closed Breach (D25,D26)":
                return "Claim Breach (Over SLA)"
            elif value == "Auto Closed Destination (CL2,CL4)":
                return "Claim Breach (Over SLA)"
            elif value == "Claim Missing":
                return "Claim Missing"
            elif value == "Claim Damage":
                return "Claim Damage"
            elif value in ["Open POD", "Problem Missroute", "Un Runsheet"]:
                return "On Process"
            elif value == "Auto Closed Origin (CL1)":
                return "Un Receiving"
            elif value == "RU Shipper/Origin":  
                return "RU Origin"
            else:
                return value
        except Exception as e:
            print(f"Error in translate for value {value}: {e}")
            return "Error"

    try:
        df["STATUS_POD_UPDATE"] = df["Cek_OP_7"].map(translate)
        if debug: print("STATUS_POD_UPDATE dibuat")
        if debug: print(f"STATUS_POD_UPDATE unique values: {df['STATUS_POD_UPDATE'].unique().tolist()}")
    except Exception as e:
        raise ValueError(f"Error in translate to STATUS_POD_UPDATE: {e}")

    # -----------------------------
    # Cleanup temporary columns
    # -----------------------------
    drop_cols = [
        col for col in df.columns 
        if col.startswith(("Status", "CEK_D24_D25", "CGK_", "Status_", "Cek_OP", "3LC_", 
                          "HVOSamaDest", "RUNSHEET_", "X_Ray_Cek1", "CHECK_NO_FW", 
                          "Status_FW_Bantu", "NEW_STATUS", "DB1_Check", "Check_RT", "DATE_RUNSHEET_2", 
                          "RUNSHEET_VS_NOW"))
    ]
    drop_cols = [col for col in drop_cols if col not in ["STATUS_POD", "STATUS_POD_UPDATE"]]
    df = df.drop(columns=drop_cols, errors="ignore")

    # -----------------------------
    # Replace STATUS_POD with STATUS_POD_UPDATE
    # -----------------------------
    if "STATUS_POD_UPDATE" in df.columns:
        if "STATUS_POD" in df.columns:
            df = df.drop(columns=["STATUS_POD"])
        df = df.rename(columns={"STATUS_POD_UPDATE": "STATUS_POD"})

    if debug:
        print("Kolom akhir:", df.columns.tolist())
        print(f"STATUS_POD values: {df['STATUS_POD'].unique().tolist()}")

    return df

