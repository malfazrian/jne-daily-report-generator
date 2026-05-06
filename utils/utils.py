import os
from pathlib import Path
import pandas as pd
from datetime import datetime
from data.bad_words import AMBIGUOUS_RECEIVERS
from rapidfuzz import fuzz, process
from config import (
    OUTPUT_BASE,
    PROJECT_REFERENCE_SOURCE_XLSX,
    PROJECT_REFERENCE_SHEET,
    PROJECT_REFERENCE_TARGET_CSV,
)

def find_files(base_dir, filename_prefix: str) -> list[str]:
    """
    Cari semua file di base_dir (rekursif) yang namanya diawali dengan prefix.
    """
    matched_files = []
    for root, dirs, files in os.walk(base_dir):
        for file in files:
            if file.lower().startswith(filename_prefix.lower()):
                matched_files.append(os.path.join(root, file))
    return matched_files

def get_gabungan_path() -> str:
    """
    Menghasilkan path output gabungan berdasarkan tanggal hari ini.
    Struktur folder: 
    \\192.168.9.74\f\ALL REPORT\01. JANUARI 2024\ALL REPORT GABUNGAN
        \<bulan_digit>. <NAMA_BULAN> <tahun_2digit>\<dd mm yy>
    """
    bulan_id = {
        1: "JANUARI", 2: "FEBRUARI", 3: "MARET", 4: "APRIL",
        5: "MEI", 6: "JUNI", 7: "JULI", 8: "AGUSTUS",
        9: "SEPTEMBER", 10: "OKTOBER", 11: "NOVEMBER", 12: "DESEMBER"
    }

    output_base = OUTPUT_BASE
    today = datetime.today()

    bulan_digit = f"{today.month:02d}"       # contoh: 08
    nama_bulan = bulan_id[today.month]       # contoh: AGUSTUS
    tahun_2digit = str(today.year)[2:]       # contoh: 25
    tanggal_str = today.strftime("%d %m %y") # contoh: 27 08 25

    output_dir = fr"{output_base}\{bulan_digit}. {nama_bulan} {tahun_2digit}\{tanggal_str}"
    return output_dir

def enrich_criteria_with_category(criteria_list, ref_path, debug=False):
    """
    Tambahkan CATEGORY ke tiap criteria di criteria_list
    berdasarkan group_name atau id_account dari file referensi.
    """

    ref_path = Path(ref_path)
    if not ref_path.exists():
        raise FileNotFoundError(f"Reference file tidak ditemukan: {ref_path}")

    # --- Load file referensi ---
    try:
        df_ref = pd.read_csv(ref_path, dtype=str, encoding="utf-8")
    except UnicodeDecodeError:
        df_ref = pd.read_csv(ref_path, dtype=str, encoding="latin1")

    # Normalisasi kolom
    df_ref = df_ref.rename(columns=lambda x: x.strip().upper())

    # Bersihkan nilai string pada kolom bertipe object
    for col in df_ref.select_dtypes(include=["object"]).columns:
        df_ref[col] = (
            df_ref[col]
            .str.replace("\xa0", " ", regex=False)
            .str.strip()
        )

    # Validasi kolom wajib
    if not {"CUST_ID", "BIG_GROUPING_CUST", "CATEGORY"}.issubset(df_ref.columns):
        raise ValueError("File referensi wajib punya kolom: CUST_ID, BIG_GROUPING_CUST, CATEGORY")

    enriched = []
    unknowns = []

    for criteria in criteria_list:
        # --- Bersihkan group_name ---
        group_name = criteria.get("group_name", "")
        if isinstance(group_name, str):
            group_name = group_name.replace("\xa0", " ").strip().upper()
        else:
            group_name = ""

        # --- Bersihkan id_accounts ---
        id_accounts = set()
        for acc in criteria.get("id_account", []):
            if isinstance(acc, str):
                id_accounts.add(acc.replace("\xa0", " ").strip())

        # --- Cari CATEGORY ---
        cat = None

        # 1. Cocokkan group_name
        if group_name:
            match = df_ref[df_ref["BIG_GROUPING_CUST"].str.upper() == group_name]
            if not match.empty:
                cat = match["CATEGORY"].iloc[0]

        # 2. Kalau belum ketemu, coba pakai id_account
        if not cat and id_accounts:
            match = df_ref[df_ref["CUST_ID"].isin(id_accounts)]
            if not match.empty:
                cat = match["CATEGORY"].iloc[0]

        # 3. Kalau tetap nggak ketemu
        if not cat:
            cat = "UNKNOWN"
            unknowns.append(criteria)

        # --- Tambah ke hasil ---
        new_entry = criteria.copy()
        new_entry["category"] = cat
        enriched.append(new_entry)

    # --- Warning kalau ada UNKNOWN ---
    if unknowns:
        print("\n[WARNING] Ditemukan project kategori UNKNOWN:")
        for crit in unknowns:
            print(f"   • group_name={crit.get('group_name','')} , id_account={','.join(crit.get('id_account',[]))}")
        print(">>> Update project_reference.csv agar tidak UNKNOWN <<<\n")

    if debug:
        print(f"[DEBUG] Total criteria: {len(criteria_list)}, UNKNOWN: {len(unknowns)}")

    return enriched

def enrich_criteria_with_latest_id_account(criteria_list, ref_path, debug=False):
    """
    Isi / update id_account pada tiap criteria berdasarkan BIG_GROUPING_CUST
    dari project_reference.csv agar selalu pakai referensi terbaru.

    Aturan:
    - Jika criteria tidak punya id_account -> pakai semua CUST_ID dari group_name.
    - Jika criteria punya id_account dan masih subset valid dari referensi -> pertahankan
      (berguna untuk report yang memang split by subset account).
    - Jika criteria punya id_account tapi tidak cocok referensi terbaru -> override
      dengan semua CUST_ID dari group_name.
    """

    ref_path = Path(ref_path)
    if not ref_path.exists():
        raise FileNotFoundError(f"Reference file tidak ditemukan: {ref_path}")

    try:
        df_ref = pd.read_csv(ref_path, dtype=str, encoding="utf-8")
    except UnicodeDecodeError:
        df_ref = pd.read_csv(ref_path, dtype=str, encoding="latin1")

    df_ref = df_ref.rename(columns=lambda x: x.strip().upper())
    for col in df_ref.select_dtypes(include=["object"]).columns:
        df_ref[col] = (
            df_ref[col]
            .str.replace("\xa0", " ", regex=False)
            .str.strip()
        )

    required_cols = {"CUST_ID", "BIG_GROUPING_CUST"}
    if not required_cols.issubset(df_ref.columns):
        raise ValueError("File referensi wajib punya kolom: CUST_ID, BIG_GROUPING_CUST")

    if "CUST_NAME" not in df_ref.columns:
        df_ref["CUST_NAME"] = ""

    # Buat mapping group_name -> list cust_id unik (pertahankan urutan awal)
    ref_group_map = {}
    for _, row in df_ref[["BIG_GROUPING_CUST", "CUST_ID"]].dropna().iterrows():
        grp = str(row["BIG_GROUPING_CUST"]).replace("\xa0", " ").strip().upper()
        cid = str(row["CUST_ID"]).strip()
        if cid and not cid.startswith("'"):
            cid = f"'{cid}"
        if not grp or not cid:
            continue
        ref_group_map.setdefault(grp, [])
        if cid not in ref_group_map[grp]:
            ref_group_map[grp].append(cid)

    enriched = []
    for criteria in criteria_list:
        new_entry = criteria.copy()

        group_name = str(criteria.get("group_name", "")).replace("\xa0", " ").strip().upper()
        latest_ids = ref_group_map.get(group_name, [])

        # fallback jika group_name di list adalah nama brand / cust_name, bukan BIG_GROUPING_CUST
        if not latest_ids and group_name:
            match_cust = df_ref[
                df_ref["CUST_NAME"].astype(str).str.upper().str.contains(group_name, na=False)
            ]
            if not match_cust.empty:
                latest_ids = []
                for cid in match_cust["CUST_ID"].dropna().astype(str):
                    cid = cid.strip()
                    if cid and not cid.startswith("'"):
                        cid = f"'{cid}"
                    if cid and cid not in latest_ids:
                        latest_ids.append(cid)

        existing_ids = []
        for x in criteria.get("id_account", []):
            s = str(x).strip()
            if not s:
                continue
            if not s.startswith("'"):
                s = f"'{s}"
            existing_ids.append(s)

        final_ids = latest_ids
        reason = "mapped_from_group"

        if existing_ids:
            latest_set = set(latest_ids)
            existing_set = set(existing_ids)

            if latest_ids and existing_set.issubset(latest_set):
                final_ids = existing_ids
                reason = "kept_existing_subset"
            elif latest_ids and not existing_set.issubset(latest_set):
                final_ids = latest_ids
                reason = "override_outdated_existing"
            else:
                # group_name belum ada di referensi: tetap pakai existing agar tidak kosong total
                final_ids = existing_ids
                reason = "fallback_existing_no_group_match"

        new_entry["id_account"] = final_ids
        enriched.append(new_entry)

        if debug:
            print(
                f"[DEBUG] id_account enrich | group={group_name} | reason={reason} | "
                f"count={len(new_entry['id_account'])}"
            )

    return enriched

def ensure_list(x):
    if isinstance(x, (list, tuple, set)):
        return list(x)
    return [x]

def is_ambiguous(receiver, threshold=85):
    receiver = receiver.strip().lower()
    if not receiver:
        return False
    
    # Jangan proses kalau nama panjang (anggap valid)
    if len(receiver) > 15:  
        return False

    for bad in AMBIGUOUS_RECEIVERS:
        bad = bad.lower().strip()
        
        # exact match cepat
        if receiver == bad:
            return True
        
        # fuzzy match (lebih aman token_set_ratio)
        score = fuzz.token_set_ratio(receiver, bad)
        if score >= threshold:
            return True
    return False

def clean_receiver_column(row):
    receiver = str(row.get("RECEIVED/REASON", "")).strip()
    consignee = str(row.get("CONSIGNEE_NAME", "")).strip()
    
    if is_ambiguous(receiver):
        return consignee if consignee else receiver
    else:
        return receiver
    
def auto_update_project_reference():
    # Path sumber dan target
    src_excel = PROJECT_REFERENCE_SOURCE_XLSX
    sheet_name = PROJECT_REFERENCE_SHEET
    target_csv = PROJECT_REFERENCE_TARGET_CSV

    # Baca sheet
    df = pd.read_excel(src_excel, sheet_name=sheet_name, dtype=str)
    # Rename kolom CUST_ID_2 ke CUST_ID, kolom lain biarkan
    df = df.rename(columns={"CUST_ID_2": "CUST_ID"})
    # Pilih kolom yang dibutuhkan (pastikan urutan sesuai kebutuhan)
    needed_cols = ["CUST_ID", "CUST_NAME", "BIG_GROUPING_CUST", "CATEGORY"]
    df_out = df[needed_cols]
    # Simpan ke CSV
    df_out.to_csv(target_csv, index=False, encoding="utf-8")
    print(f"✅ project_reference.csv berhasil diupdate dari {src_excel}")

#PIVOT CONFIGS
pivot_standard = [
                {
                    "name": "PivotStatus",
                    "dest": "B3",
                    "rows": ["STATUS_POD"],
                    "columns": [],
                    "filters": [],
                    "values": [
                        {"field": "AWB", "name": "JUMLAH", "func": "count"}
                    ]
                },
                {
                    "name": "Pivot3LC",
                    "dest": "E3",
                    "rows": ["3 LC DEST"],
                    "columns": [],
                    "filters": [],
                    "values": [
                        {"field": "AWB", "name": "JUMLAH", "func": "count"},
                        {"field": "AMOUNT", "func": "sum", "caption": "JUMLAH_AMOUNT", "num_format": '"IDR" #,##0.00'}
                    ]
                },
                {
                    "name": "PivotCarrer",
                    "dest": "I3",
                    "rows": ["CARRER"],
                    "columns": [],
                    "filters": [],
                    "values": [
                        {"field": "AWB", "name": "JUMLAH", "func": "count"}
                    ]
                }
            ]

pivot_status_by_awb = [
    {
        "name": "PivotStatusByAWB",
        "dest": "B3",
        "rows": ["STATUS_POD"],
        "columns": [],
        "filters": [],
        "values": [
            {"field": "AWB", "name": "JUMLAH", "func": "count"},
            {"field": "AWB", "name": "JUMLAH %", "func": "count", "as_percentage": True, "percentage_of": "column"}
        ]
    }
]

pivot_status_by_periode = [
                          {
                            "name": "PivotStatus",
                            "dest": "B3",
                            "rows": ["STATUS_POD"],
                            "columns": ["PERIODE"],
                            "filters": [],
                            "values": [
                                {"field": "AWB", "name": "JUMLAH", "func": "count"}
                            ]
                        }
                      ]

pivot_aj_car = [
    {
        "name": "AllSummary",
        "dest": "B2",
        "rows": ["AJ Car Status", "KETERANGAN AJ CAR", "CODING_UNDEL", "REASON UNDEL"],
        "columns": [],
        "filters": [],
        "values": [
            {"field": "AWB", "name": "Total", "func": "count"}
        ]
    },
    {
        "name": "ReasonReturnSummary",
        "dest": "H2",
        "rows": ["REASON UNDEL"],
        "hide_blank": True,
        "columns": [],
        "filters": [],
        "values": [
            {"field": "AWB", "name": "Total", "func": "count"}
        ]
    },
    {
        "name": "StatusbyTglEntry",
        "dest": "L2",
        "rows": ["TGL_ENTRY"],
        "columns": ["AJ Car Status"],
        "filters": [],
        "values": [
            {"field": "AWB", "name": "Total", "func": "count"}
        ]
    },
    {
        "name": "StatusbyTglEntryPercent",
        "dest": "T2",
        "rows": ["TGL_ENTRY"],
        "columns": ["AJ Car Status"],
        "filters": [],
        "values": [
            {"field": "AWB", "name": "% of Row Total", "func": "count", "as_percentage": True, "percentage_of": "row"}
        ]
    }
]

pivot_status_by_dest = [
                          {
                            "name": "PivotStatusByDest",
                            "dest": "B3",
                            "rows": ["3 LC DEST"],
                            "columns": ["STATUS_POD"],
                            "filters": [],
                            "values": [
                                {"field": "AWB", "name": "JUMLAH", "func": "count"}
                            ]
                        }
                      ]

pivot_status_by_id = [
                          {
                            "name": "PivotStatusById",
                            "dest": "B2",
                            "rows": ["STATUS_POD"],
                            "columns": ["ID_ACCOUNT"],
                            "filters": [],
                            "values": [
                                {"field": "AWB", "name": "JUMLAH", "func": "count"}
                            ]
                        }
                      ]

pivot_status_by_id_tgl_entry = [
                          {
                            "name": "PivotStatusByIdTglEntry",
                            "dest": "B2",
                            "rows": ["ID_ACCOUNT", "STATUS_POD"],
                            "columns": ["TGL_ENTRY"],
                            "filters": [],
                            "values": [
                                {"field": "AWB", "name": "JUMLAH", "func": "count"}
                            ]
                        }
                      ]

pivot_status_by_3LC = [{
                    "name": "PivotStatusBy3LC",
                    "dest": "B2",
                    "rows": ["3 LC DEST"],
                    "columns": ["STATUS_POD"],
                    "filters": [],
                    "values": [
                        {"field": "AWB", "name": "JUMLAH", "func": "count"}
                    ]
                }]

pivot_pzc = [
    {
        "name": "PivotPZC",
        "dest": "B2",
        "rows": ["DEPT PZC"],
        "columns": ["STATUS_POD"],
        "filters": [],
        "values": [
            {"field": "AWB", "name": "JUMLAH", "func": "count"}
        ]
    }
]

pivot_watson = [
    {
        "name": "PivotStatusWatson",
        "dest": "B2",
        "rows": ["STATUS_POD"],
        "columns": [],
        "filters": [],
        "values": [
            {"field": "AWB", "name": "JUMLAH", "func": "count"}
        ]
    },
    {
        "name": "PivotOriginWatson",
        "dest": "E2",
        "rows": ["ORIGIN 2", "ORIGIN"],
        "columns": [],
        "filters": [],
        "values": [
            {"field": "AWB", "name": "JUMLAH", "func": "count"}
        ]
    },
    {
        "name": "PivotCarrerPercentWatson",
        "dest": "I2",
        "rows": ["STATUS_POD"],
        "columns": ["CARRER"],
        "filters": [],
        "values": [
            {"field": "AWB", "name": "% of Row Total", "func": "count", "as_percentage": True, "percentage_of": "row"}
        ]
    },
    {
        "name": "PivotCarrerWatson",
        "dest": "I22",
        "rows": ["STATUS_POD"],
        "columns": ["CARRER"],
        "filters": [],
        "values": [
            {"field": "AWB", "name": "JUMLAH", "func": "count"}
        ]
    }
]

pivot_reason_return = [
    {
        "name": "PivotReasonReturn",
        "dest": "B2",
        "rows": ["REASON RETURN"],
        "columns": [],
        "filters": [
            {"field": "STATUS_POD", "value": "Undel"}
        ],
        "values": [
            {"field": "AWB", "name": "JUMLAH", "func": "count"}
        ]
    }
]

pivot_uob = [
    {
        "name": "PivotCycleNameByStatus",
        "dest": "B2",
        "rows": ["CYCLE NAME"],
        "columns": ["STATUS"],
        "filters": [],
        "values": [
            {"field": "REFF NUM", "name": "JUMLAH", "func": "count"}
        ]
    },
    {
        "name": "PivotCycleNameByKetAWB",
        "dest": "J2",
        "rows": ["CYCLE NAME"],
        "columns": ["KET AWB"],
        "filters": [],
        "values": [
            {"field": "REFF NUM", "name": "JUMLAH", "func": "count"}
        ]
    }
]

pivot_map_aktif = [
    {
        "name": "PivotStatusMapAktif",
        "dest": "B2",
        "rows": ["PERIODE", "WEEK_OF_YEAR"],
        "columns": ["CARRER"],
        "filters": [],
        "values": [
            {"field": "AWB", "name": "JUMLAH", "func": "count"},
            {"field": "AWB", "name": "% of Row Total", "func": "count", "as_percentage": True, "percentage_of": "row"}
        ]
    }
]

pivot_status_by_tgl_entry = [
    {
        "name": "PivotStatusByTglEntry",
        "dest": "B2",
        "rows": ["STATUS_POD"],
        "columns": ["TGL_ENTRY"],
        "filters": [],
        "values": [
            {"field": "AWB", "name": "JUMLAH", "func": "count"},
            {
                "field": "AWB",
                "name": "% of Row Total",
                "func": "count",
                "as_percentage": True,
                "percentage_of": "column"
            }
        ],
        "sort": {
            "row": "STATUS_POD",
            "by": "% of Row Total",
            "order": "desc"
        }
    }
]

pivot_rodamas = [
    {
        "name": "PivotStatusRodamas",
        "dest": "B2",
        "rows": ["STATUS_POD"],
        "columns": [],
        "filters": [],
        "values": [
            {"field": "AWB", "name": "JUMLAH", "func": "count"},
            {"field": "AWB", "name": "% of Row Total", "func": "count", "as_percentage": True, "percentage_of": "column"}
        ]
    },
    {
        "name": "PivotServiceRodamas",
        "dest": "F2",
        "rows": ["SERVICE"],
        "columns": [],
        "filters": [],
        "values": [
            {"field": "AWB", "name": "JUMLAH", "func": "count"},
            {"field": "AWB", "name": "% of Row Total", "func": "count", "as_percentage": True, "percentage_of": "column"}
        ]
    },
    {
        "name": "PivotCustomerRodamas",
        "dest": "J2",
        "rows": ["CUSTOMER"],
        "columns": [],
        "filters": [],
        "values": [
            {"field": "AWB", "name": "JUMLAH", "func": "count"},
            {"field": "AWB", "name": "% of Row Total", "func": "count", "as_percentage": True, "percentage_of": "column"}
        ]
    },
    {
        "name": "PivotStatusByTglEntry",
        "dest": "N2",
        "rows": ["TGL_ENTRY"],
        "columns": [],
        "filters": [],
        "values": [
            {"field": "AWB", "name": "JUMLAH", "func": "count"},
            {
                "field": "AWB",
                "name": "% of Row Total",
                "func": "count",
                "as_percentage": True,
                "percentage_of": "column"
            }
        ]
    }
]

pivot_template_zilong = [
        {
            "name": "PivotStatusZilong",
            "dest": "B2",
            "rows": ["STATUS_POD"],
            "columns": [],
            "filters": [],
            "values": [
                {"field": "AWB", "name": "JUMLAH", "func": "count"}
            ]
        },
        {
            "name": "PivotZonaZilong",
            "dest": "B13",
            "rows": ["ZONA"],
            "columns": [],
            "filters": [],
            "values": [
                {"field": "AWB", "name": "JUMLAH", "func": "count"},
            ]
        },
        {
            "name": "PivotCustomerNameZilong",
            "dest": "E2",
            "rows": ["CUST_NAME"],
            "columns": [],
            "filters": [],
            "values": [
                {"field": "AWB", "name": "JUMLAH", "func": "count"}
            ]
        },
        {
            "name": "PivotRegionalZilong",
            "dest": "E12",
            "rows": ["REGIONAL"],
            "columns": [],
            "filters": [],
            "values": [
                {"field": "AWB", "name": "JUMLAH", "func": "count"}
            ]
        },
        {
            "name": "PivotCarrer1stZilong",
            "dest": "H2",
            "rows": ["CAREER_1ST"],
            "columns": [],
            "hide_blank": True,
            "filters": [],
            "values": [
                {"field": "AWB", "name": "JUMLAH", "func": "count"}
            ]
        },
        {
            "name": "PivotWilayahZilong",
            "dest": "H12",
            "rows": ["WILAYAH"],
            "columns": [],
            "filters": [],
            "values": [
                {"field": "AWB", "name": "JUMLAH", "func": "count"}
            ]
        },
        {
            "name": "PivotCarrerPODZilong",
            "dest": "K2",
            "rows": ["CARRER_POD"],
            "columns": [],
            "hide_blank": True,
            "filters": [],
            "values": [
                {"field": "AWB", "name": "JUMLAH", "func": "count"}
            ]
        },
        {
            "name": "PivotPeriodeZilong",
            "dest": "K12",
            "rows": ["TGL_ENTRY"],
            "columns": [],
            "filters": [],
            "values": [
                {"field": "AWB", "name": "JUMLAH", "func": "count"}
            ]
        }
    ]