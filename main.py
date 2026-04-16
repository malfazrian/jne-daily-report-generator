import os
import pandas as pd
from data_loader.master_data_reader import get_data_from_master
from list import criteria_lists
from datetime import datetime
from data_transform.edit_file import process_edit_file
# from data_submitter.send_wa import send_all_files
from utils.utils import enrich_criteria_with_category, ensure_list, auto_update_project_reference
from utils.helper import should_send, is_processed_today, mark_processed_today
from collections import defaultdict
from config import (
    OUTPUT_BASE,
    BASE_DIR_LEGACY,
    TABLE_REFERENCE_PATH,
    MANUALS_PATH,
    PROJECT_REFERENCE_TARGET_CSV,
)

# Mapping bulan ke bahasa Indonesia
bulan_id = {
    1: "JANUARI", 2: "FEBRUARI", 3: "MARET", 4: "APRIL",
    5: "MEI", 6: "JUNI", 7: "JULI", 8: "AGUSTUS",
    9: "SEPTEMBER", 10: "OKTOBER", 11: "NOVEMBER", 12: "DESEMBER"
}

# Base folder tetap
output_base = OUTPUT_BASE

# Ambil tanggal hari ini
today = datetime.today()
bulan_digit = f"{today.month:02d}"         # 08
nama_bulan = bulan_id[today.month]         # AGUSTUS
tahun_2digit = str(today.year)[2:]         # 25
tanggal_str = today.strftime("%d %m %y")   # 08 08 25

base_dir = BASE_DIR_LEGACY
output_dir = fr"{output_base}\{bulan_digit}. {nama_bulan} {tahun_2digit}\{tanggal_str}\trial BOT"
table_reference_path = TABLE_REFERENCE_PATH
manuals_path = MANUALS_PATH
ref_path = PROJECT_REFERENCE_TARGET_CSV

auto_update_project_reference()

# --- enrich criteria dengan category ---
criteria_with_category = enrich_criteria_with_category(criteria_lists, ref_path)

# 🔍 Cek siapa saja yang UNKNOWN
unknowns = [c for c in criteria_with_category if c.get("category", "").upper() == "UNKNOWN"]
if unknowns:
    print("\n[WARNING] Ditemukan project kategori UNKNOWN:")
    for crit in unknowns:
        print(f"   • group_name={crit.get('group_name', '')}, "
              f"id_account={','.join(crit.get('id_account', []))}")
    print(">>> Update project_reference.csv agar tidak UNKNOWN <<<\n")

# --- group by category ---
criteria_by_category = {}
for crit in criteria_with_category:
    cat = crit.get("category", "UNKNOWN").upper()
    criteria_by_category.setdefault(cat, []).append(crit)

all_file_list = []

# --- loop per category ---
for category, criteria_list in criteria_by_category.items():
    print(f"[INFO] Proses category: {category} ({len(criteria_list)} projects)")

    # === Cek apakah kategori ini sudah diproses hari ini ===
    process_key_category = f"query_{category}"
    if is_processed_today(process_key_category):
        print(f"⏩ Skip get_data_from_master untuk kategori '{category}', sudah diproses hari ini.")
    else:
        try:
            get_data_from_master(
                base_dir,
                criteria_list,
                category=category,
                output_dir=output_dir,
                reference_path=table_reference_path,
                debug=False,
                bypass_history=False
            )
            # Tandai kategori sudah selesai diproses
            mark_processed_today(process_key_category)
            print(f"✅ Kategori '{category}' ditandai sudah diproses hari ini.")
        except Exception as e:
            print(f"❌ Gagal proses kategori {category}: {e}")
            continue

    # === Tetap jalankan proses edit file (agar file_list tetap lengkap) ===
    frontline_files = process_edit_file(criteria_list, output_dir, manuals_path, bypass_history=False)

    # Masukkan frontline ke all_file_list
    file_list = [
        {"contact": c, "files": files}
        for contact, files in frontline_files.items()
        for c in ensure_list(contact)
    ]
    all_file_list.extend(file_list)

# --- gabungkan by contact ---
grouped_files = defaultdict(list)
for entry in all_file_list:
    grouped_files[entry["contact"]].extend(entry["files"])

# pastikan tidak ada duplikat file per contact
final_file_list = []
for contact, files in grouped_files.items():
    unique_files = list(set(files))
    send_files = []

    for f in unique_files:
        if should_send(f): # cek apakah perlu dikirim
            send_files.append(f)
        else:
            print(f"⏩ Skip kirim {os.path.basename(f)}, STATUS_POD tidak berubah")

    if send_files:
        final_file_list.append({"contact": contact, "files": send_files})

# # --- kirim ---
# if final_file_list:
#     send_all_files(final_file_list)