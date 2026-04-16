import os, time, shutil, pythoncom, win32com.client as win32, psutil
from win32com.client import constants as c 
from .pivot_helper import work_on_local_copy, copy_back, safe_create_pivotcache, ensure_unique_pivot_name, make_pivot, is_writable, safe_clear_pivots
from datetime import datetime, date
import threading


PIVOT_CONFIGS = [
    {
        "name": "PivotStatus",             # nama pivot
        "sheet": "Summary",                # sheet tujuan
        "dest_cell": "B3",                 # anchor cell
        "rows": ["STATUS_POD"],            # baris
        "columns": [],                     # kolom
        "values": [                        # data values
            {"field": "AWB", "func": "count", "caption": "JUMLAH"}
        ],
        "filters": []                      # filter (opsional)
    },
    {
        "name": "PivotDestAmt",
        "sheet": "Summary",
        "dest_cell": "E2",
        "rows": ["3 LC DEST"],
        "columns": [],
        "values": [
            {"field": "AWB", "func": "count", "caption": "JUMLAH"},
            {"field": "AMOUNT", "func": "sum", "caption": "JUMLAH_AMOUNT", "num_format": '"IDR" #,##0.00'}
        ],
        "filters": []
    },
    {
        "name": "PivotCarrier",
        "sheet": "Summary",
        "dest_cell": "I3",
        "rows": ["CARRER"],
        "columns": [],
        "values": [
            {"field": "AWB", "func": "count", "caption": "JUMLAH"}
        ],
        "filters": []
    }
]

def _group_pivot_date_by_month_year(pt, field_name: str) -> bool:
    """
    Group field tanggal di PivotTable per Bulan + Tahun.
    Kembalikan True kalau sukses, False kalau gagal.
    """
    try:
        pf = pt.PivotFields(field_name)
        # Pastikan layout sudah jadi
        try:
            pt.RefreshTable()
        except Exception:
            pass

        rng = None
        # DataRange biasanya ada untuk Row/Column fields
        try:
            rng = pf.DataRange
        except Exception:
            rng = None
        # Fallback ke LabelRange kalau DataRange None
        if rng is None:
            try:
                rng = pf.LabelRange
            except Exception:
                rng = None

        if rng is None:
            print(f"[!] Tidak bisa akses range untuk grouping pada field: {field_name}")
            return False

        # Periods = [sec, min, hour, day, month, quarter, year]
        rng.Cells(1, 1).Group(Start=True, End=True,
                              Periods=[False, False, False, False, True, False, True])
        return True
    except Exception as e:
        print(f"[!] Gagal group {field_name} per bulan+tahun: {e}")
        return False

def sort_pivot_rows_by_value(pt, row_field_name, data_field_caption, order="desc"):
    """
    Sort RowField PivotTable berdasarkan DataField tertentu.
    order: 'asc' | 'desc'
    """
    try:
        rf = pt.PivotFields(row_field_name)

        # pastikan pivot up-to-date
        try:
            pt.RefreshTable()
        except Exception:
            pass

        xl_order = c.xlDescending if str(order).lower() == "desc" else c.xlAscending

        rf.AutoSort(
            xl_order,
            data_field_caption
        )
        return True
    except Exception as e:
        print(f"[!] Gagal sort pivot row '{row_field_name}' by '{data_field_caption}': {e}")
        return False

def add_pivots(file_path: str, data_sheet_name: str | None = None, pivots: list | None = None, summary_sheet_name: str = "Summary", date_format: str = "dd/mm/yyyy",
    max_retries: int = 5) -> bool:
    """
    Buat PivotTables asli via Excel COM **di salinan lokal** lalu copy balik.
    Return True kalau sukses (pivot terbuat), False kalau gagal.
    """
    if not os.path.exists(file_path):
        print(f"[!] File tidak ditemukan: {file_path}")
        return False

    # 1) Kerja di temp lokal (hindari UNC & path panjang/locking)
    tmp_path, tmp_dir = work_on_local_copy(file_path)

    def close_excel_processes(force=False):
        """Close Excel processes, force terminate if force=True"""
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] and proc.info['name'].lower() == "excel.exe":
                try:
                    if force:
                        proc.terminate()
                    else:
                        proc.kill()
                except Exception:
                    pass
    
    ok = False
    for attempt in range(max_retries):
        xl = wb = None
        try:
            pythoncom.CoInitialize()
            xl = win32.gencache.EnsureDispatch("Excel.Application")
            xl.DisplayAlerts = False
            xl.Visible = False
            xl.ScreenUpdating = False
            
            # 2) Open temp file dgn parameter aman
            open_kwargs = dict(
                Filename=os.path.abspath(tmp_path),
                UpdateLinks=0,
                ReadOnly=False,
                AddToMru=False,
                IgnoreReadOnlyRecommended=True
            )
            wb = xl.Workbooks.Open(**open_kwargs)

            # baru setelah wb terbuka, coba set Calculation
            try:
                xl.Calculation = c.xlCalculationManual
            except Exception:
                print("[!] Gagal set Calculation ke Manual, lanjut default...")

            # 3) Ambil worksheet data
            src_ws = wb.Worksheets(data_sheet_name) if data_sheet_name else wb.Worksheets(1)

            # Pastikan ada data
            last_row = src_ws.Cells(src_ws.Rows.Count, 1).End(c.xlUp).Row
            last_col = src_ws.Cells(1, src_ws.Columns.Count).End(c.xlToLeft).Column
            if last_row < 2 or last_col < 1:
                print("[!] Data terlalu sedikit untuk pivot.")
                return False

            src_range = src_ws.Range(src_ws.Cells(1, 1), src_ws.Cells(last_row, last_col))

            # Normalisasi kolom tanggal → hilangkan time (jam:menit:detik)
            for col in range(1, last_col + 1):
                header = str(src_ws.Cells(1, col).Value).strip().upper()
                if header == "TGL_ENTRY":
                    rng = src_ws.Range(src_ws.Cells(2, col), src_ws.Cells(last_row, col))
                    try:
                        rng.NumberFormat = date_format

                        for r in range(2, last_row + 1):
                            val = src_ws.Cells(r, col).Value
                            # Excel COM kasih pywintypes.Time (subclass dari datetime)
                            if isinstance(val, datetime):
                                # buang jam → set ke 00:00:00 tapi pertahankan tzinfo
                                new_val = val.replace(hour=0, minute=0, second=0, microsecond=0)
                                src_ws.Cells(r, col).Value = new_val

                        for r in range(2, min(last_row, 6)):
                            new_val = src_ws.Cells(r, col).Value

                    except Exception as e:
                        pass

            # 4) Hapus Summary lama jika ada
            safe_clear_pivots(wb)
            try:
                wb.Worksheets(summary_sheet_name).Delete()
            except Exception:
                pass

            sum_ws = wb.Worksheets.Add(After=wb.Sheets(wb.Sheets.Count))
            sum_ws.Name = summary_sheet_name

            if not pivots:
                print("[!] Tidak ada konfigurasi pivot yang diberikan")
                return False

            # 5) Build pivots
            for pivot in pivots:
                pc = safe_create_pivotcache(wb, src_range)
                # nama unik
                table_name = ensure_unique_pivot_name(wb, pivot.get("name", "Pivot"))
                pt = make_pivot(pc, pivot.get("dest", "B3"), table_name, sum_ws)

                # rows
                for i, row_field in enumerate(pivot.get("rows", []), start=1):
                    pf = pt.PivotFields(row_field)
                    pf.Orientation = c.xlRowField
                    pf.Position = i

                    # ganti "Row Labels" jadi nama field aslinya
                    try:
                        pt.RowAxisLayout(1)  # pastikan tabular layout
                    except Exception:
                        pass
                    pf.Caption = row_field  

                    # --- jika kolom tanggal, group per bulan ---
                    hdr = str(row_field).upper()
                    if ("TGL" in hdr) or ("DATE" in hdr and "UPDATE" not in hdr):
                        _group_pivot_date_by_month_year(pt, row_field)
                        
                # columns
                for i, col_field in enumerate(pivot.get("columns", []), start=1):
                    pf = pt.PivotFields(col_field)
                    pf.Orientation = c.xlColumnField
                    pf.Position = i

                    # --- jika kolom tanggal, group per bulan ---
                    hdr = str(col_field).upper()
                    if ("TGL" in hdr) or ("DATE" in hdr and "UPDATE" not in hdr):
                        _group_pivot_date_by_month_year(pt, col_field)

                # filters
                for i, flt in enumerate(pivot.get("filters", []), start=1):
                    if isinstance(flt, str):
                        # mode lama: cuma field
                        pf = pt.PivotFields(flt)
                        pf.Orientation = c.xlPageField
                        pf.Position = i
                    elif isinstance(flt, dict):
                        # mode baru: ada field + value
                        pf = pt.PivotFields(flt["field"])
                        pf.Orientation = c.xlPageField
                        pf.Position = i
                        try:
                            pf.CurrentPage = flt["value"]   # set ke "Undel"
                        except Exception as e:
                            print(f"[!] Gagal set filter {flt['field']} ke {flt['value']}: {e}")

                # values
                for val in pivot.get("values", []):
                    field = pt.PivotFields(val["field"])
                    fn = val.get("func", "sum").lower()
                    caption = val.get("name") or val.get("caption") or f"{fn.title()} of {val['field']}"

                    if fn == "count":
                        df = pt.AddDataField(field, caption, c.xlCount)
                    elif fn in ("avg", "average"):
                        df = pt.AddDataField(field, caption, c.xlAverage)
                    elif fn == "max":
                        df = pt.AddDataField(field, caption, c.xlMax)
                    elif fn == "min":
                        df = pt.AddDataField(field, caption, c.xlMin)
                    else:
                        df = pt.AddDataField(field, caption, c.xlSum)

                    if val.get("as_percentage"):
                        perc_type = (val.get("percentage_of") or "row").lower()
                        if perc_type == "row":
                            df.Calculation = c.xlPercentOfRow
                        elif perc_type == "column":
                            df.Calculation = c.xlPercentOfColumn
                        elif perc_type == "total":
                            df.Calculation = c.xlPercentOfTotal
                        else:
                            print(f"[!] Jenis percentage '{perc_type}' tidak dikenali, fallback ke row")
                            df.Calculation = c.xlPercentOfRow
                        df.NumberFormat = "0.00%"
                    elif "num_format" in val:
                        df.NumberFormat = val["num_format"]

                # --- SORT ---
                sort_cfg = pivot.get("sort")
                if sort_cfg:
                    sort_pivot_rows_by_value(
                        pt,
                        row_field_name=sort_cfg.get("row"),
                        data_field_caption=sort_cfg.get("by"),
                        order=sort_cfg.get("order", "desc")
                    )

            # 6) Save temp & close Excel
            wb.Save()
            ok = True
            break  

        except Exception as e:
            print(f"[!] Attempt {attempt+1}/{max_retries} gagal: {e}")
            # Auto-clear corrupt gen_py cache
            if "gen_py" in str(e):
                gen_py_dir = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Temp", "gen_py")
                if os.path.isdir(gen_py_dir):
                    try:
                        shutil.rmtree(gen_py_dir)
                        print(f"[i] gen_py cache dihapus: {gen_py_dir}")
                    except Exception as rm_err:
                        print(f"[!] Gagal hapus gen_py cache: {rm_err}")
                # Also clear the in-memory cache so next EnsureDispatch rebuilds it
                try:
                    win32.gencache.__init__()
                except Exception:
                    pass
            close_excel_processes(force=(attempt == max_retries-1))
            time.sleep(10)

        finally:
            try:
                if wb: wb.Close(SaveChanges=True)
            except Exception:
                pass
            try:
                if xl:
                    try:
                        xl.Calculation = c.xlCalculationAutomatic
                    except Exception:
                        pass
                    xl.ScreenUpdating = True
                    xl.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()
            time.sleep(0.8)

    # 7) Copy balik hasil temp → path asli (jangan SaveAs UNC dari Excel)
    if ok:
        if not is_writable(file_path):
            # mencoba lagi sebentar kalau share sempat nge-lock
            for _ in range(5):
                time.sleep(0.8)
                if is_writable(file_path):
                    break
        if copy_back(tmp_path, file_path):
            return True
        else:
            print("[!] Pivot berhasil dibuat di temp, tapi gagal copy balik ke lokasi asli.")
            return False
    else:
        return False