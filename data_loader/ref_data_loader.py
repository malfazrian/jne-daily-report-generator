import os
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from data_transform.parse_dates import normalize_all_dates
import glob

def read_any_file(path):
    """Baca file Excel/CSV fleksibel."""
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext in [".xlsx", ".xlsm", ".xlsb"]:
            return pd.read_excel(path, dtype=str)
        elif ext == ".xls":
            return pd.read_excel(path, engine="xlrd", dtype=str)
        elif ext == ".csv":
            return pd.read_csv(path, dtype=str, sep=None, engine="python")
        else:
            print(f"⚠️ Format tidak didukung: {path}")
            return None
    except Exception as e:
        print(f"❌ Gagal baca {path}: {e}")
        return None

def load_cust_ref(
    ref_base_dir: str,
    jumlah_bulan: int,
    ref_sheet: str = None,
    date_col: str = None
) -> pd.DataFrame:

    today = datetime.today()
    start_date = today - relativedelta(months=jumlah_bulan - 1)

    # =====================================================
    # 🟢 MODE 1: SINGLE FILE (FULL PATH)
    # =====================================================
    if os.path.isfile(ref_base_dir):

        print(f"📄 Membaca single cust_ref file: {ref_base_dir}")

        ext = os.path.splitext(ref_base_dir)[1].lower()

        try:
            if ext == ".csv":
                df = read_any_file(ref_base_dir)
            else:
                df = pd.read_excel(ref_base_dir, sheet_name=ref_sheet) if ref_sheet else pd.read_excel(ref_base_dir)
        except Exception as e:
            print(f"❌ Gagal baca file: {e}")
            return pd.DataFrame()

        if df is None or df.empty:
            return pd.DataFrame()

        df.columns = df.columns.str.upper().str.strip()
        df = normalize_all_dates(df)

        if not date_col:
            raise ValueError("⚠️ date_col wajib diisi jika menggunakan single file")

        date_col = date_col.upper()

        if date_col not in df.columns:
            raise ValueError(f"⚠️ Kolom date_col '{date_col}' tidak ditemukan")

        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")

        # filter bulan
        df = df[df[date_col] >= start_date]

        df["_source_file"] = os.path.basename(ref_base_dir)
        df["_source_period"] = df[date_col].dt.strftime("%y%m")

        print(f"📊 Total cust_ref rows (filtered): {len(df)}")

        return df.reset_index(drop=True)

    # =====================================================
    # 🟢 MODE 2: FOLDER PERIODIK (LOGIC LAMA)
    # =====================================================
    if not os.path.isdir(ref_base_dir):
        print(f"❌ Path tidak valid: {ref_base_dir}")
        return pd.DataFrame()

    periods = []
    for i in range(jumlah_bulan):
        d = today - relativedelta(months=i)
        periods.append(d.strftime("%y%m"))
    periods = sorted(periods)

    print(f"📌 Periode cust_ref yg dibaca: {periods}")

    all_data = []

    for p in periods:
        folder = os.path.join(ref_base_dir, p)

        if not os.path.exists(folder):
            print(f"⚠️ Folder tidak ditemukan: {folder}")
            continue

        print(f"📂 Membaca folder cust_ref: {folder}")

        files = glob.glob(os.path.join(folder, "*.*"))

        for f in files:
            ext = os.path.splitext(f)[1].lower()
            if ext not in [".csv", ".xls", ".xlsx", ".xlsm", ".xlsb"]:
                continue

            try:
                if ext == ".csv":
                    df = read_any_file(f)
                else:
                    df = pd.read_excel(f, sheet_name=ref_sheet) if ref_sheet else pd.read_excel(f)
            except Exception as e:
                print(f"⚠️ Gagal baca file {f}: {e}")
                continue

            if df is not None and not df.empty:
                df["_source_period"] = p
                df["_source_file"] = os.path.basename(f)
                df.columns = df.columns.str.upper().str.strip()
                all_data.append(df)

    if not all_data:
        print("❌ Tidak ada file cust_ref valid.")
        return pd.DataFrame()

    combined = pd.concat(all_data, ignore_index=True, join="outer")
    print(f"📊 Total cust_ref rows: {len(combined)}")

    return combined

