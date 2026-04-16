import os, shutil, tempfile, uuid, time, re
import win32com.client as win32
from win32com.client import constants as c
import psutil

def work_on_local_copy(src_path: str) -> tuple[str, str]:
    """Copy file ke temp lokal dgn nama pendek. Return (tmp_path, tmp_dir)."""
    tmp_dir = os.path.join(tempfile.gettempdir(), "pivot_tmp")
    os.makedirs(tmp_dir, exist_ok=True)
    tmp_path = os.path.join(tmp_dir, f"{uuid.uuid4().hex}.xlsx")
    shutil.copy2(src_path, tmp_path)
    return tmp_path, tmp_dir

def copy_back(src_tmp: str, dst_path: str, retries: int = 5, delay: float = 1.5) -> bool:
    """Copy balik hasil temp ke path asli, retry kalau ada PermissionError."""
    for _ in range(retries):
        try:
            # gunakan copy2 supaya preserve metadata (ctime/mtime)
            shutil.copy2(src_tmp, dst_path)
            return True
        except PermissionError:
            time.sleep(delay)
    return False

def is_writable(path: str) -> bool:
    """Cek bisa buka handle write tanpa mengubah isi."""
    try:
        with open(path, "r+b"):
            return True
    except Exception:
        return False

def ensure_unique_pivot_name(wb, base: str) -> str:
    """Pastikan nama PivotTables unik di workbook."""
    safe = re.sub(r'[^A-Za-z0-9_]', '_', base or "Pivot")
    existing = set()
    for i in range(1, wb.Sheets.Count + 1):
        sh = wb.Sheets(i)
        try:
            pt_count = sh.PivotTables().Count
            for j in range(1, pt_count + 1):
                existing.add(sh.PivotTables(j).Name)
        except Exception:
            pass
    if safe not in existing:
        return safe
    # tambah suffix timestamp kecil
    suffix = str(int(time.time() * 1000) % 100000)
    new_name = f"{safe}_{suffix}"
    return new_name if new_name not in existing else f"{safe}_{suffix}_x"

def safe_create_pivotcache(wb, src_range):
    for attempt in range(5):
        try:
            return wb.PivotCaches().Create(SourceType=c.xlDatabase, SourceData=src_range)
        except Exception as e:
            print(f"[!] Gagal membuat PivotCache (attempt {attempt+1}): {e}")
            time.sleep(0.8)
    raise RuntimeError("Gagal membuat PivotCache setelah 5x percobaan.")

def make_pivot(cache, dest_cell, table_name, ws):
    return cache.CreatePivotTable(
        TableDestination=ws.Range(dest_cell),
        TableName=table_name
    )

def is_file_unlocked(path: str) -> bool:
    """Cek apakah file bisa diakses (tidak ke-lock Excel)."""
    try:
        if os.path.exists(path) and os.access(path, os.W_OK):
            with open(path, "a"):
                return True
    except PermissionError:
        return False
    return False

def kill_excel_processes():
    """Paksa matikan semua proses EXCEL.EXE."""
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] and proc.info['name'].lower() == "excel.exe":
            try:
                proc.kill()
                print(f"💀 Killed EXCEL process PID={proc.pid}")
            except Exception as e:
                print(f"⚠️ Gagal kill PID={proc.pid}: {e}")

def try_open_workbook(file_path, retries=5, delay=2):
    """Coba buka workbook dengan retry, kalau gagal return None"""
    for attempt in range(retries):
        try:
            xl = win32.gencache.EnsureDispatch("Excel.Application")
            xl.DisplayAlerts = False
            xl.Visible = False
            wb = xl.Workbooks.Open(file_path, ReadOnly=False)
            return xl, wb
        except Exception as e:
            print(f"⏳ Workbook masih locked? retry {attempt+1}/{retries} ... {e}")
            time.sleep(delay)
    return None, None

def wait_until_unlocked(path, timeout=3*60):
    """Tunggu sampai file bisa ditulis (tidak ke-lock Excel)."""
    start = time.time()
    while time.time() - start < timeout:
        try:
            with open(path, "a"):
                return True
        except PermissionError:
            time.sleep(1)
    return False

def safe_clear_pivots(wb):
    """Hapus semua PivotTable lama dari workbook."""
    try:
        for ws in wb.Worksheets:
            try:
                pt_count = ws.PivotTables().Count
                for j in range(pt_count, 0, -1):
                    try:
                        ws.PivotTables(j).TableRange2.Clear()
                        ws.PivotTables(j).Delete()
                    except Exception:
                        pass
            except Exception:
                pass
    except Exception:
        pass
