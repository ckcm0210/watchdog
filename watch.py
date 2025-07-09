import os
import time
import csv
import hashlib
import gc
import psutil
import shutil
import tempfile
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from openpyxl import load_workbook

# =========== User Config ============
WATCH_FOLDERS = [
    r"\\network_drive\your_folder1",
    r"\\network_drive\your_folder2"
]
LOG_FOLDER = r".\excel_watch_log"
CSV_LOG_FILE = os.path.join(LOG_FOLDER, "excel_change_log.csv")
SUPPORTED_EXTS = ('.xlsx', '.xlsm')

# Smart retry config
MAX_RETRY = 10           # user可改
RETRY_INTERVAL_SEC = 2   # user可改
USE_TEMP_COPY = True     # user可改（True = 用temp copy方法，False = 直接開原檔案）

# =========== End User Config ============

def get_all_excel_files(folders):
    all_files = []
    for folder in folders:
        for dirpath, _, filenames in os.walk(folder):
            for f in filenames:
                if f.lower().endswith(SUPPORTED_EXTS) and not f.startswith('~$'):
                    all_files.append(os.path.join(dirpath, f))
    return all_files

def get_memory_mb():
    gc.collect()
    process = psutil.Process(os.getpid())
    mem = process.memory_info().rss / 1024 / 1024
    return mem

def get_excel_last_author(path):
    try:
        wb = load_workbook(path, read_only=True)
        author = wb.properties.lastModifiedBy
        wb.close()
        return author
    except Exception as e:
        print(f"[ERROR] 無法讀取 last author: {e}")
        return None

def dump_excel_cells_with_formula(path):
    try:
        wb_formula = load_workbook(path, data_only=False)
        wb_value = load_workbook(path, data_only=True)
        result = {}
        for ws_formula, ws_value in zip(wb_formula.worksheets, wb_value.worksheets):
            ws_data = {}
            for row_formula, row_value in zip(ws_formula.iter_rows(), ws_value.iter_rows()):
                for cell_formula, cell_value in zip(row_formula, row_value):
                    formula = cell_formula.value if cell_formula.data_type == "f" else None
                    value = cell_value.value
                    # 一律包裝成 dict，保證 downstream 一定係 dict
                    ws_data[cell_formula.coordinate] = {
                        "formula": formula,
                        "value": value
                    }
            result[ws_formula.title] = ws_data
        wb_formula.close()
        wb_value.close()
        return result
    except Exception as e:
        print(f"[ERROR] 無法讀取 Excel cell: {e}")
        return {}

def hash_excel_content(cells_dict):
    try:
        flat = [
            (ws, coord, dct.get("formula"), dct.get("value"))
            for ws, ws_dict in cells_dict.items()
            for coord, dct in ws_dict.items()
        ]
        return hashlib.md5(str(sorted(flat)).encode('utf-8')).hexdigest()
    except Exception as e:
        print(f"[ERROR] hash 失敗: {e}")
        return None

def load_baseline(baseline_file):
    if os.path.exists(baseline_file):
        with open(baseline_file, 'r', encoding='utf-8') as f:
            import json
            return json.load(f)
    return None

def save_baseline(baseline_file, data):
    with open(baseline_file, 'w', encoding='utf-8') as f:
        import json
        json.dump(data, f, ensure_ascii=False, indent=2)

def safe_get(val, key, default=None):
    if isinstance(val, dict):
        return val.get(key, default)
    return default

def compare_cells(old, new):
    changes = []
    old = old or {}
    new = new or {}
    for ws in new:
        old_ws = old.get(ws, {})
        new_ws = new[ws]
        all_cells = set(new_ws.keys()) | set(old_ws.keys())
        for cell in all_cells:
            old_val = old_ws.get(cell, {"formula": None, "value": None})
            new_val = new_ws.get(cell, {"formula": None, "value": None})
            # 用 safe_get 防止 float/string error
            if old_val != new_val:
                changes.append({
                    "worksheet": ws,
                    "cell": cell,
                    "old_formula": safe_get(old_val, "formula"),
                    "old_value": safe_get(old_val, "value"),
                    "new_formula": safe_get(new_val, "formula"),
                    "new_value": safe_get(new_val, "value")
                })
    for ws in old:
        if ws not in new:
            for cell, old_val in old[ws].items():
                changes.append({
                    "worksheet": ws,
                    "cell": cell,
                    "old_formula": safe_get(old_val, "formula"),
                    "old_value": safe_get(old_val, "value"),
                    "new_formula": None,
                    "new_value": None
                })
    return changes

def log_changes_csv(csv_log_file, file_path, last_author, changes):
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    os.makedirs(os.path.dirname(csv_log_file), exist_ok=True)
    file_exists = os.path.exists(csv_log_file)
    with open(csv_log_file, 'a', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        if not file_exists:
            writer.writerow([
                "timestamp", "file", "worksheet", "cell", "old_formula", "old_value", "new_formula", "new_value", "last_author"
            ])
        for change in changes:
            writer.writerow([
                now,
                file_path,
                change['worksheet'],
                change['cell'],
                change['old_formula'],
                change['old_value'],
                change['new_formula'],
                change['new_value'],
                last_author
            ])

def print_console_header():
    print("\n" + "="*80)
    print(" Excel File Change Watcher ".center(80, "-"))
    print("="*80 + "\n")

def print_event(msg, char="-"):
    print(char*80)
    print(msg)
    print(char*80)

def print_cell_changes_summary(changes, max_show=10):
    print(f"  變更 cell 數量：{len(changes)}")
    for i, change in enumerate(changes[:max_show]):
        ws = change['worksheet']
        cell = change['cell']
        oldf = change['old_formula']
        oldv = change['old_value']
        newf = change['new_formula']
        newv = change['new_value']
        print(f"    [{ws}] {cell}: [公式:{oldf}] [值:{oldv}]  →  [公式:{newf}] [值:{newv}]")
    if len(changes) > max_show:
        print(f"    ... 其餘 {len(changes) - max_show} 個 cell 省略 ...")

def create_baseline_for_files(xlsx_files):
    total = len(xlsx_files)
    for idx, file_path in enumerate(xlsx_files, 1):
        base_name = os.path.basename(file_path)
        baseline_file = os.path.join(LOG_FOLDER, f"{base_name}.baseline.json")
        if os.path.exists(baseline_file):
            continue  # 已有 baseline
        print(f"[baseline] {idx}/{total}: {file_path}")
        mem_before = get_memory_mb()
        cell_data = dump_excel_cells_with_formula(file_path)
        curr_author = get_excel_last_author(file_path)
        curr_hash = hash_excel_content(cell_data)
        save_baseline(baseline_file, {
            "last_author": curr_author,
            "content_hash": curr_hash,
            "cells": cell_data
        })
        mem_after = get_memory_mb()
        print(f"    [memory] 用咗 {mem_after-mem_before:.2f} MB, 目前 process 共用 {mem_after:.2f} MB")

class ExcelChangeHandler(FileSystemEventHandler):
    def on_modified(self, event):
        if not event.is_directory and event.src_path.lower().endswith(SUPPORTED_EXTS):
            filename = os.path.basename(event.src_path)
            if filename.startswith('~$'):
                return  # skip Excel temp files

            print_event(f"[{datetime.now().strftime('%H:%M:%S')}] 檔案有更動：{event.src_path}")

            base_name = filename
            baseline_file = os.path.join(LOG_FOLDER, f"{base_name}.baseline.json")

            for attempt in range(MAX_RETRY):
                temp_path = None
                try:
                    # === temp copy workaround ===
                    file_to_open = event.src_path
                    if USE_TEMP_COPY:
                        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(event.src_path)[1]) as tmpf:
                            shutil.copy2(event.src_path, tmpf.name)
                            temp_path = tmpf.name
                        file_to_open = temp_path

                    is_first_time = not os.path.exists(baseline_file)
                    baseline = load_baseline(baseline_file)
                    prev_cells = baseline['cells'] if baseline else {}
                    prev_hash = baseline['content_hash'] if baseline else None

                    curr_author = get_excel_last_author(file_to_open)
                    curr_cells = dump_excel_cells_with_formula(file_to_open)
                    curr_hash = hash_excel_content(curr_cells)

                    if is_first_time:
                        # 首次偵測：不 print，只 log
                        changes = compare_cells({}, curr_cells)
                        log_changes_csv(CSV_LOG_FILE, event.src_path, curr_author, changes)
                        save_baseline(baseline_file, {
                            "last_author": curr_author,
                            "content_hash": curr_hash,
                            "cells": curr_cells
                        })
                        break

                    if curr_hash != prev_hash:
                        changes = compare_cells(prev_cells, curr_cells)
                        if changes:
                            print(f"  [INFO] 檢測到 cell 內容有更動！")
                            print(f"  Last Author: {curr_author or '未知'}")
                            print_cell_changes_summary(changes)
                            log_changes_csv(CSV_LOG_FILE, event.src_path, curr_author, changes)
                        else:
                            print("  [INFO] 檔案內容有 hash 變，但未 detect cell 差異。")
                        save_baseline(baseline_file, {
                            "last_author": curr_author,
                            "content_hash": curr_hash,
                            "cells": curr_cells
                        })
                    else:
                        print("  [INFO] 檔案有更動，但 cell 內容無改變。")
                    break
                except PermissionError:
                    print(f"  [WARN] 檔案 lock 緊，等一等再讀... (retry {attempt+1}/{MAX_RETRY})")
                    time.sleep(RETRY_INTERVAL_SEC)
                except Exception as e:
                    print(f"[ERROR] 輸出出錯: {e}")
                    break
                finally:
                    if temp_path and os.path.exists(temp_path):
                        os.remove(temp_path)

if __name__ == "__main__":
    print_console_header()
    print("  監控資料夾:")
    for folder in WATCH_FOLDERS:
        print(f"    - {folder}")
    print(f"  支援副檔名: {SUPPORTED_EXTS}")
    print(f"  Log/baseline 儲存位置: {os.path.abspath(LOG_FOLDER)}")
    print(f"  變更 Log (CSV): {os.path.abspath(CSV_LOG_FILE)}")
    print(f"  Smart retry config:")
    print(f"    MAX_RETRY = {MAX_RETRY}")
    print(f"    RETRY_INTERVAL_SEC = {RETRY_INTERVAL_SEC}")
    print(f"    USE_TEMP_COPY = {USE_TEMP_COPY}")
    os.makedirs(LOG_FOLDER, exist_ok=True)

    choice = input("\n要唔要 scan 晒所有 Excel 做 baseline？(y/n): ").strip().lower()
    if choice == "y":
        all_files = get_all_excel_files(WATCH_FOLDERS)
        print(f"總共 find 到 {len(all_files)} 個 Excel file.")
        create_baseline_for_files(all_files)
        print("baseline scan 完成！\n")

    event_handler = ExcelChangeHandler()
    observer = Observer()
    for folder in WATCH_FOLDERS:
        observer.schedule(event_handler, folder, recursive=True)
    print("\n  [INFO] 開始監控...\n")
    try:
        observer.start()
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        print("\n  [INFO] 停止監控，程式結束。")
    observer.join()
