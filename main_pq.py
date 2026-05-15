import os
import re
from datetime import datetime
from collections import defaultdict
from concurrent.futures import ProcessPoolExecutor, as_completed
import math
import pandas as pd

from data_loader.master_data_pq_reader import get_data_from_master_pq
from data_transform.edit_file import process_edit_file
# from data_submitter.send_wa import send_all_files

from utils.utils import (
    enrich_criteria_with_category,
    enrich_criteria_with_latest_id_account,
    ensure_list,
    auto_update_project_reference,
)
from utils.helper import should_send
from utils.helper import clean_old_process_history
from list import criteria_lists
from config import (
    OUTPUT_BASE,
    BASE_DIR_PARQUET,
    TABLE_REFERENCE_PATH,
    MANUALS_PATH,
    PROJECT_REFERENCE_TARGET_CSV,
)


# =========================
# CONFIG & CONSTANT
# =========================

bulan_id = {
    1: "JANUARI", 2: "FEBRUARI", 3: "MARET", 4: "APRIL",
    5: "MEI", 6: "JUNI", 7: "JULI", 8: "AGUSTUS",
    9: "SEPTEMBER", 10: "OKTOBER", 11: "NOVEMBER", 12: "DESEMBER"
}

output_base = OUTPUT_BASE
base_dir = BASE_DIR_PARQUET
table_reference_path = TABLE_REFERENCE_PATH

manuals_path = MANUALS_PATH
ref_path = PROJECT_REFERENCE_TARGET_CSV

LOCAL_TEMP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "duckdb_temp")
os.makedirs(LOCAL_TEMP_DIR, exist_ok=True)
os.environ["TEMP"] = LOCAL_TEMP_DIR
os.environ["TMP"] = LOCAL_TEMP_DIR
os.environ["TMPDIR"] = LOCAL_TEMP_DIR

MAX_WORKERS = int(os.getenv("BOT_MAX_WORKERS", "1"))


# =========================
# DATE & PATH
# =========================

today = datetime.today()
bulan_digit = f"{today.month:02d}"
nama_bulan = bulan_id[today.month]
tahun_2digit = str(today.year)[2:]
tanggal_str = today.strftime("%d %m %y")

output_dir = (
    fr"{output_base}\{bulan_digit}. {nama_bulan} {tahun_2digit}"
    fr"\{tanggal_str}\trial BOT"
)

os.makedirs(output_dir, exist_ok=True)


# =========================
# HELPER: PROCESS CATEGORY
# =========================

def _chunk_list(items, chunk_size):
    for i in range(0, len(items), chunk_size):
        yield items[i:i + chunk_size]


def _sanitize_dirname(name: str) -> str:
    if not name:
        return "UNKNOWN"
    cleaned = re.sub(r"[<>:\\/?*\"|]", " ", str(name))
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned or "UNKNOWN"


def _criteria_label(criteria: dict) -> str:
    return str(
        criteria.get("save_as")
        or criteria.get("group_name")
        or criteria.get("groups_name")
        or "UNKNOWN"
    )


def expand_grouped_report_criteria(criteria_list):
    """
    Ubah criteria dengan format:
        {"groups_name": "Nama File", "reports": [{"group_name": "..."}]}
    menjadi beberapa criteria sheet yang semuanya disimpan ke workbook yang sama.
    """
    expanded = []

    for criteria in criteria_list:
        reports = criteria.get("reports")
        if not reports:
            expanded.append(criteria)
            continue

        report_file_name = (
            criteria.get("save_as")
            or criteria.get("groups_name")
            or criteria.get("group_name")
        )
        if not report_file_name:
            raise ValueError("Criteria dengan 'reports' wajib punya 'groups_name' atau 'save_as'.")

        parent_defaults = {
            key: value
            for key, value in criteria.items()
            if key not in {"reports", "group_name", "groups_name", "save_as"}
        }

        for report in reports:
            child = parent_defaults.copy()
            child.update(report)
            child.setdefault("selected_cols", criteria.get("selected_cols"))
            child.setdefault("jumlah_bulan", criteria.get("jumlah_bulan"))
            child.setdefault("period", criteria.get("period"))
            child["move_to_sheet_of_file"] = report_file_name
            child["move_to_sheet_name"] = criteria.get("sheet_name") or report_file_name
            child["append_to_sheet"] = True
            child.setdefault("save_as", report.get("save_as") or report.get("group_name"))
            child["_combined_report_name"] = report_file_name
            child["_combined_report_industry"] = criteria.get("industry") or "UNKNOWN"
            expanded.append(child)

    return expanded


def reset_grouped_report_workbooks(criteria_list, output_root):
    removed = set()
    for crit in criteria_list:
        target_name = crit.get("move_to_sheet_of_file")
        if not target_name:
            continue
        industry = _sanitize_dirname(crit.get("industry", "UNKNOWN"))
        target_file = os.path.join(output_root, industry, f"{target_name}.xlsx")
        if target_file in removed or not os.path.exists(target_file):
            continue
        try:
            os.remove(target_file)
            removed.add(target_file)
            print(f"[INFO] Reset grouped workbook: {target_file}")
        except PermissionError:
            print(f"⚠️ Tidak bisa reset grouped workbook karena sedang dibuka: {target_file}")


def load_industry_map(reference_path: str, sheet_name: str = "ACC & SHIPPER GROUPING") -> dict:
    try:
        df = pd.read_excel(reference_path, sheet_name=sheet_name, dtype=str)
    except Exception as e:
        print(f"âš ï¸ Gagal baca reference untuk industry: {e}")
        return {}

    df.columns = df.columns.str.upper().str.strip()
    if "BIG_GROUPING_CUST" not in df.columns or "CUST_INDUSTRY_NEW" not in df.columns:
        print("âš ï¸ Kolom BIG_GROUPING_CUST / CUST_INDUSTRY_NEW tidak ditemukan")
        return {}

    df["BIG_GROUPING_CUST"] = (
        df["BIG_GROUPING_CUST"].astype(str).str.replace("\xa0", " ").str.strip().str.upper()
    )
    df["CUST_INDUSTRY_NEW"] = (
        df["CUST_INDUSTRY_NEW"].astype(str).str.replace("\xa0", " ").str.strip()
    )

    mapping = (
        df.dropna(subset=["BIG_GROUPING_CUST", "CUST_INDUSTRY_NEW"])
        .drop_duplicates(subset=["BIG_GROUPING_CUST"], keep="first")
        .set_index("BIG_GROUPING_CUST")["CUST_INDUSTRY_NEW"]
        .to_dict()
    )
    return mapping


def process_one_category(args, bypass_history=False):
    """
    HANYA get_data_from_master_pq
    Tidak boleh edit file / pivot di sini
    """
    category, criteria_list, chunk_idx, chunk_total, output_root = args
    failed_projects = []

    try:
        print(
            f"[INFO] START category {category} "
            f"chunk {chunk_idx + 1}/{chunk_total} "
            f"({len(criteria_list)} projects)"
        )

        for crit in criteria_list:
            project_name = _criteria_label(crit)
            try:
                industry = _sanitize_dirname(crit.get("industry", "UNKNOWN"))
                output_dir_industry = os.path.join(output_root, industry)
                os.makedirs(output_dir_industry, exist_ok=True)

                ok = get_data_from_master_pq(
                    base_dir=base_dir,
                    criteria=crit,
                    category=category,
                    output_dir=output_dir_industry,
                    reference_path=table_reference_path,
                    manuals_path=manuals_path,
                    debug=False,
                    bypass_history=bypass_history
                )
                if ok is False:
                    failed_projects.append(project_name)
                    print(f"âš ï¸ Project gagal/skip: {category} | {project_name}")

            except Exception as e:
                failed_projects.append(f"{project_name}: {e}")
                print(f"âŒ Project error: {category} | {project_name}: {e}")

        if failed_projects:
            shown_failures = ", ".join(failed_projects[:10])
            remaining_count = len(failed_projects) - 10
            if remaining_count > 0:
                shown_failures = f"{shown_failures}, ... (+{remaining_count} lainnya)"
            return category, True, f"{len(failed_projects)} project gagal/skip: {shown_failures}"

        return category, True, None

    except Exception as e:
        return category, False, str(e)

# =========================
# MAIN
# =========================

if __name__ == "__main__":

    removed_history_count = clean_old_process_history(days=14)
    if removed_history_count:
        print(f"ðŸ§¹ Hapus {removed_history_count} log process_history lama")

    auto_update_project_reference()

    criteria_lists_expanded = expand_grouped_report_criteria(criteria_lists)

    # --- enrich criteria ---
    criteria_with_category = enrich_criteria_with_category(criteria_lists_expanded, ref_path)

    # --- resolve latest id_account from reference by group_name ---
    criteria_with_category = enrich_criteria_with_latest_id_account(
        criteria_with_category,
        ref_path,
    )

    # --- enrich industry ---
    industry_map = load_industry_map(table_reference_path)
    for crit in criteria_with_category:
        if crit.get("_combined_report_name"):
            industry = crit.get("_combined_report_industry") or "UNKNOWN"
        else:
            group_name = str(crit.get("group_name", "")).replace("\xa0", " ").strip().upper()
            industry = industry_map.get(group_name) or "UNKNOWN"
        crit["industry"] = industry

    # --- cek UNKNOWN ---
    unknowns = [
        c for c in criteria_with_category
        if c.get("category", "").upper() == "UNKNOWN"
    ]
    if unknowns:
        print("\n[WARNING] Ditemukan project kategori UNKNOWN:")
        for crit in unknowns:
            print(
                f"   â€¢ group_name={crit.get('group_name', '')}, "
                f"id_account={','.join(crit.get('id_account', []))}"
            )
        print(">>> Update project_reference.csv agar tidak UNKNOWN <<<\n")

    # --- group by category ---
    criteria_by_category = defaultdict(list)
    for crit in criteria_with_category:
        cat = crit.get("category", "UNKNOWN").upper()
        criteria_by_category[cat].append(crit)

    # =========================
    # PARALLEL: GET DATA ONLY
    # =========================

    bypass_history = False

    print(f"Parallel get_data_from_master_pq | workers={MAX_WORKERS}\n")
 
    combined_criteria_by_category = defaultdict(list)
    normal_criteria_by_category = defaultdict(list)
    for category, criteria_list in criteria_by_category.items():
        for crit in criteria_list:
            if crit.get("move_to_sheet_of_file"):
                combined_criteria_by_category[category].append(crit)
            else:
                normal_criteria_by_category[category].append(crit)

    if combined_criteria_by_category:
        print("\n[INFO] Proses grouped report (serial, 1 workbook 1 sheet gabungan)\n")
        for category, criteria_list in combined_criteria_by_category.items():
            reset_grouped_report_workbooks(criteria_list, output_dir)
            _, ok, err = process_one_category(
                (category, criteria_list, 0, 1, output_dir),
                bypass_history=bypass_history,
            )
            if ok:
                print(f"Grouped report category {category} selesai")
                if err:
                    print(f"Grouped report category {category} selesai dengan catatan: {err}")
            else:
                print(f"Grouped report category {category} gagal: {err}")

    tasks = []
    for category, criteria_list in normal_criteria_by_category.items():
        if not criteria_list:
            continue
        chunk_size = max(2, math.ceil(len(criteria_list) / MAX_WORKERS))
        chunks = list(_chunk_list(criteria_list, chunk_size))
        chunk_total = len(chunks)
        for idx, chunk in enumerate(chunks):
            tasks.append((category, chunk, idx, chunk_total, output_dir))

    if tasks:
        with ProcessPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = [executor.submit(process_one_category, t, bypass_history) for t in tasks]

            for fut in as_completed(futures):
                category, ok, err = fut.result()
                if ok:
                    print(f"✅ Category {category} selesai")
                    if err:
                        print(f"⚠️ Category {category} selesai dengan catatan: {err}")
                else:
                    print(f"❌ Category {category} gagal: {err}")

    # =========================
    # SINGLE THREAD: EDIT FILE
    # =========================

    print("Proses edit file (SINGLE THREAD, AMAN PIVOT)\n")

    all_file_list = []

    for category, criteria_list in criteria_by_category.items():
        print(f"[INFO] Edit file category: {category}")

        try:
            criteria_list = [
                crit for crit in criteria_list
                if not crit.get("move_to_sheet_of_file")
            ]
            if not criteria_list:
                continue

            frontline_files = process_edit_file(
                criteria_list,
                output_dir,
                manuals_path,
                bypass_history=bypass_history
            )

            file_list = [
                {"contact": c, "files": files}
                for contact, files in frontline_files.items()
                for c in ensure_list(contact)
            ]
            all_file_list.extend(file_list)

        except Exception as e:
            print(f"âŒ Gagal edit file category {category}: {e}")

    # =========================
    # GROUP & SEND
    # =========================

    # grouped_files = defaultdict(list)
    # for entry in all_file_list:
    #     grouped_files[entry["contact"]].extend(entry["files"])

    # final_file_list = []

    # for contact, files in grouped_files.items():
    #     unique_files = list(set(files))
    #     send_files = []

    #     for f in unique_files:
    #         if should_send(f):
    #             send_files.append(f)
    #         else:
    #             print(f"â© Skip kirim {os.path.basename(f)}, STATUS_POD tidak berubah")

    #     if send_files:
    #         final_file_list.append({"contact": contact, "files": send_files})

    # if final_file_list:
    #     print("\nðŸ“¤ Kirim file WhatsApp\n")
    #     send_all_files(final_file_list)
    # else:
    #     print("\nðŸ“­ Tidak ada file untuk dikirim\n")
