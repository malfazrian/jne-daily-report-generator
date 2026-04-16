import json
import os
from datetime import datetime, timedelta
import time

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
HISTORY_FILE = os.path.join(BASE_DIR, "data", "status_history.json")
PROCESS_HISTORY_PATH = os.path.join(os.path.dirname(__file__), "data", "process_history.json")
LOCK_FILE = PROCESS_HISTORY_PATH + ".lock"

def acquire_lock():
    while True:
        try:
            # buat lock file (exclusive)
            fd = os.open(LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.close(fd)
            return
        except FileExistsError:
            time.sleep(0.05)

def release_lock():
    if os.path.exists(LOCK_FILE):
        os.remove(LOCK_FILE)

def load_history():
    if not os.path.exists(HISTORY_FILE):
        return {}
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError:
        corrupt_path = HISTORY_FILE + ".corrupt"
        os.replace(HISTORY_FILE, corrupt_path)
        print(f"⚠️ WARNING: status_history.json corrupt → moved to {corrupt_path}")
        return {}

def save_history(history):
    # Pastikan folder ada
    os.makedirs(os.path.dirname(HISTORY_FILE), exist_ok=True)
    with open(HISTORY_FILE, "w") as f:
        json.dump(history, f, indent=2)

def should_send(file_path: str) -> bool:
    basename = os.path.basename(file_path)
    lower_name = basename.lower()

    BULAN_ID_MAP = {
        1: "januari",
        2: "februari",
        3: "maret",
        4: "april",
        5: "mei",
        6: "juni",
        7: "juli",
        8: "agustus",
        9: "september",
        10: "oktober",
        11: "november",
        12: "desember",
    }

    now = datetime.now()
    bulan_sekarang = BULAN_ID_MAP[now.month]
    tahun_sekarang = str(now.year)

    # ✅ RULE 1: nama file mengandung BULAN SEKARANG (Indonesia)
    if bulan_sekarang in lower_name:
        return True

    # ✅ RULE 2: nama file mengandung TAHUN BERJALAN
    if tahun_sekarang in lower_name:
        return True

    history = load_history()
    today_counts = history.get(basename)
    last_counts = history.get(f"{basename}__last")

    if today_counts is None:
        # belum pernah disimpan snapshot → anggap baru
        return True

    statuses = set(today_counts.keys())
    closed_statuses = {"Success", "Return Shipper"}

    # Tidak ada perubahan
    if last_counts == today_counts:
        # Kalau masih ada status di luar closed → tetap kirim
        if not statuses.issubset(closed_statuses):
            return True
        return False

    # Ada perubahan → kirim dan update snapshot
    history[f"{basename}__last"] = today_counts
    save_history(history)
    return True

def load_process_history():
    if not os.path.exists(PROCESS_HISTORY_PATH):
        return {}
    try:
        with open(PROCESS_HISTORY_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError:
        # Backup file corrupt
        corrupt_path = PROCESS_HISTORY_PATH + ".corrupt"
        os.rename(PROCESS_HISTORY_PATH, corrupt_path)
        print(f"⚠️ WARNING: process_history.json corrupt → moved to {corrupt_path}")
        return {}

def save_process_history(history):
    acquire_lock()
    try:
        tmp_path = PROCESS_HISTORY_PATH + ".tmp"
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(history, f, ensure_ascii=False, indent=2)
        os.replace(tmp_path, PROCESS_HISTORY_PATH)  # atomic overwrite
    finally:
        release_lock()

def clean_old_process_history(days=14):
    history = load_process_history()
    if not history:
        return 0

    cutoff_date = datetime.today().date() - timedelta(days=days)
    cleaned_history = {}
    removed_count = 0

    for key, value in history.items():
        try:
            processed_date = datetime.strptime(str(value), "%Y-%m-%d").date()
        except (TypeError, ValueError):
            removed_count += 1
            continue

        if processed_date >= cutoff_date:
            cleaned_history[key] = value
        else:
            removed_count += 1

    if removed_count:
        save_process_history(cleaned_history)

    return removed_count

def is_processed_today(key):
    history = load_process_history()
    today = datetime.today().strftime("%Y-%m-%d")
    return history.get(key, "") == today

def mark_processed_today(key):
    history = load_process_history()
    today = datetime.today().strftime("%Y-%m-%d")
    history[key] = today
    save_process_history(history)