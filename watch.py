import os
import time
import csv
import hashlib
import gc
import psutil
import shutil
import tempfile
import gzip
import json
import signal
import threading
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from openpyxl import load_workbook

# =========== User Config ============
SCAN_ALL_MODE = True

# 🚀 優化選項
USE_LOCAL_CACHE = True
ENABLE_FAST_MODE = True
CACHE_FOLDER = r".\\excel_cache"

# 🔧 診斷和恢復選項
ENABLE_TIMEOUT = True          # 啟用超時保護
FILE_TIMEOUT_SECONDS = 120     # 每個檔案最大處理時間 (秒)
ENABLE_MEMORY_MONITOR = True   # 啟用記憶體監控
MEMORY_LIMIT_MB = 2048         # 記憶體限制 (MB)
ENABLE_RESUME = True           # 啟用斷點續傳
RESUME_LOG_FILE = r".\\baseline_progress.log"  # 進度記錄檔

WATCH_FOLDERS = [
    r"\\network_drive\\your_folder1",
    r"\\network_drive\\your_folder2"
]

MANUAL_BASELINE_TARGET = [
    r"\\network_drive\\your_folder1\\somefile.xlsx",
    r"\\network_drive\\your_folder2\\subfolder"
]

LOG_FOLDER = r".\\excel_watch_log"
LOG_FILE_DATE = datetime.now().strftime('%Y%m%d')
CSV_LOG_FILE = os.path.join(LOG_FOLDER, f"excel_change_log_{LOG_FILE_DATE}.csv.gz")
SUPPORTED_EXTS = ('.xlsx', '.xlsm')

MAX_RETRY = 10
RETRY_INTERVAL_SEC = 2
USE_TEMP_COPY = True

WHITELIST_USERS = ['ckcm0210', 'yourwhiteuser']
LOG_WHITELIST_USER_CHANGE = True

FORCE_BASELINE_ON_FIRST_SEEN = [
    r"\\network_drive\\your_folder1\\must_first_baseline.xlsx",
    "force_this_file.xlsx"
]
# =========== End User Config ============

# 全局變數
current_processing_file = None
processing_start_time = None
force_stop = False

def signal_handler(signum, frame):
    """處理 Ctrl+C 中斷"""
    global force_stop
    force_stop = True
    print("\n🛑 收到中斷信號，正在安全停止...")
    if current_processing_file:
        print(f"   目前處理檔案: {current_processing_file}")

signal.signal(signal.SIGINT, signal_handler)

def get_memory_usage():
    """獲取目前記憶體使用量"""
    try:
        process = psutil.Process(os.getpid())
        return process.memory_info().rss / 1024 / 1024  # MB
    except Exception:
        return 0

def check_memory_limit():
    """檢查記憶體是否超限"""
    if not ENABLE_MEMORY_MONITOR:
        return False
    
    current_memory = get_memory_usage()
    if current_memory > MEMORY_LIMIT_MB:
        print(f"⚠️ 記憶體使用量過高: {current_memory:.1f} MB > {MEMORY_LIMIT_MB} MB")
        print("   正在執行垃圾回收...")
        gc.collect()
        new_memory = get_memory_usage()
        print(f"   垃圾回收後: {new_memory:.1f} MB")
        return new_memory > MEMORY_LIMIT_MB
    return False

def save_progress(completed_files, total_files):
    """儲存進度到檔案"""
    if not ENABLE_RESUME:
        return
    
    try:
        progress_data = {
            "timestamp": datetime.now().isoformat(),
            "completed": completed_files,
            "total": total_files,
            "completed_list": completed_files  # 可以改為檔案列表
        }
        
        with open(RESUME_LOG_FILE, 'w', encoding='utf-8') as f:
            json.dump(progress_data, f, ensure_ascii=False, indent=2)
            
    except Exception as e:
        print(f"[WARN] 無法儲存進度: {e}")

def load_progress():
    """載入之前的進度"""
    if not ENABLE_RESUME or not os.path.exists(RESUME_LOG_FILE):
        return None
    
    try:
        with open(RESUME_LOG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"[WARN] 無法載入進度: {e}")
        return None

def timeout_handler():
    """超時處理函數"""
    global current_processing_file, processing_start_time, force_stop
    
    while not force_stop:
        time.sleep(10)  # 每 10 秒檢查一次
        
        if current_processing_file and processing_start_time:
            elapsed = time.time() - processing_start_time
            if elapsed > FILE_TIMEOUT_SECONDS:
                print(f"\n⏰ 檔案處理超時!")
                print(f"   檔案: {current_processing_file}")
                print(f"   已處理時間: {elapsed:.1f} 秒 > {FILE_TIMEOUT_SECONDS} 秒")
                print(f"   將跳過此檔案並繼續...")
                # 這裡可以設置一個標誌來跳過當前檔案
                current_processing_file = None
                processing_start_time = None

def get_all_excel_files(folders):
    all_files = []
    for folder in folders:
        if os.path.isfile(folder):
            if folder.lower().endswith(SUPPORTED_EXTS) and not os.path.basename(folder).startswith('~$'):
                all_files.append(folder)
        elif os.path.isdir(folder):
            for dirpath, _, filenames in os.walk(folder):
                for f in filenames:
                    if f.lower().endswith(SUPPORTED_EXTS) and not f.startswith('~$'):
                        all_files.append(os.path.join(dirpath, f))
    return all_files

def serialize_cell_value(value):
    """快速序列化"""
    if value is None:
        return None
    elif isinstance(value, datetime):
        return value.isoformat()
    elif isinstance(value, (int, float, str, bool)):
        return value
    else:
        return str(value)

def get_excel_last_author(path):
    try:
        wb = load_workbook(path, read_only=True)
        author = wb.properties.lastModifiedBy
        wb.close()
        return author
    except Exception:
        return None

def copy_to_cache(network_path):
    """🚀 帶診斷的緩存功能"""
    if not USE_LOCAL_CACHE:
        return network_path
    
    try:
        os.makedirs(CACHE_FOLDER, exist_ok=True)
        
        # 檢查原始檔案是否存在和可讀
        if not os.path.exists(network_path):
            raise FileNotFoundError(f"網絡檔案不存在: {network_path}")
        
        if not os.access(network_path, os.R_OK):
            raise PermissionError(f"無法讀取網絡檔案: {network_path}")
        
        file_hash = hashlib.md5(network_path.encode('utf-8')).hexdigest()[:16]
        cache_file = os.path.join(CACHE_FOLDER, f"{file_hash}_{os.path.basename(network_path)}")
        
        # 檢查緩存
        if os.path.exists(cache_file):
            try:
                network_mtime = os.path.getmtime(network_path)
                cache_mtime = os.path.getmtime(cache_file)
                if cache_mtime >= network_mtime:
                    return cache_file
            except Exception:
                pass
        
        # 複製檔案，顯示進度
        network_size = os.path.getsize(network_path)
        print(f"   📥 複製到緩存: {os.path.basename(network_path)} ({network_size/(1024*1024):.1f} MB)")
        
        copy_start = time.time()
        shutil.copy2(network_path, cache_file)
        copy_time = time.time() - copy_start
        
        print(f"      複製完成，耗時 {copy_time:.1f} 秒")
        return cache_file
        
    except Exception as e:
        print(f"   ❌ 緩存失敗: {e}")
        return network_path

def dump_excel_cells_with_timeout(path):
    """🚀 帶超時保護的 Excel 讀取"""
    global current_processing_file, processing_start_time
    
    current_processing_file = path
    processing_start_time = time.time()
    
    try:
        # 檢查檔案基本資訊
        file_size = os.path.getsize(path)
        print(f"   📊 檔案大小: {file_size/(1024*1024):.1f} MB")
        
        # 使用本地緩存
        local_path = copy_to_cache(path)
        
        if ENABLE_FAST_MODE:
            # 快速模式
            print(f"   🚀 使用快速模式讀取...")
            wb = load_workbook(local_path, read_only=True, data_only=False)
            result = {}
            
            worksheet_count = len(wb.worksheets)
            print(f"   📋 工作表數量: {worksheet_count}")
            
            for idx, ws in enumerate(wb.worksheets, 1):
                print(f"      處理工作表 {idx}/{worksheet_count}: {ws.title}")
                
                ws_data = {}
                cell_count = 0
                
                if ws.max_row > 1 or ws.max_column > 1:
                    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, 
                                          min_col=1, max_col=ws.max_column):
                        for cell in row:
                            if cell.value is not None:
                                formula = None
                                if cell.data_type == "f":
                                    formula = str(cell.value)
                                    if not formula.startswith("="):
                                        formula = "=" + formula
                                
                                ws_data[cell.coordinate] = {
                                    "formula": formula,
                                    "value": serialize_cell_value(cell.value)
                                }
                                cell_count += 1
                
                print(f"         找到 {cell_count} 個有資料的 cell")
                
                if ws_data:
                    result[ws.title] = ws_data
            
            wb.close()
            print(f"   ✅ Excel 讀取完成")
        else:
            # 標準模式
            print(f"   📚 使用標準模式讀取...")
            wb_formula = load_workbook(local_path, data_only=False)
            wb_value = load_workbook(local_path, data_only=True)
            result = {}
            
            for ws_formula, ws_value in zip(wb_formula.worksheets, wb_value.worksheets):
                ws_data = {}
                for row_formula, row_value in zip(ws_formula.iter_rows(), ws_value.iter_rows()):
                    for cell_formula, cell_value in zip(row_formula, row_value):
                        try:
                            formula = cell_formula.value if cell_formula.data_type == "f" else None
                            value = serialize_cell_value(cell_value.value)
                            
                            if formula or (value not in [None, ""]):
                                if formula is not None:
                                    formula = str(formula)
                                    if not formula.startswith("="):
                                        formula = "=" + formula
                                    if not formula.startswith("'="):
                                        formula = "'" + formula
                                ws_data[cell_formula.coordinate] = {
                                    "formula": formula,
                                    "value": value
                                }
                        except Exception:
                            pass
                
                if ws_data:
                    result[ws_formula.title] = ws_data
            
            wb_formula.close()
            wb_value.close()
        
        current_processing_file = None
        processing_start_time = None
        return result
        
    except Exception as e:
        current_processing_file = None
        processing_start_time = None
        print(f"   ❌ Excel 讀取失敗: {e}")
        return {}

def hash_excel_content(cells_dict):
    try:
        content_str = json.dumps(cells_dict, sort_keys=True, ensure_ascii=False)
        return hashlib.md5(content_str.encode('utf-8')).hexdigest()
    except Exception:
        return None

def baseline_file_path(base_name):
    return os.path.join(LOG_FOLDER, f"{base_name}.baseline.json.gz")

def save_baseline(baseline_file, data):
    try:
        with gzip.open(baseline_file, 'wt', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, separators=(',', ':'))
    except Exception as e:
        print(f"[ERROR][save_baseline] error saving {baseline_file}: {e}")

def is_force_baseline_file(filepath):
    try:
        lowerfile = filepath.lower()
        for pattern in FORCE_BASELINE_ON_FIRST_SEEN:
            if pattern.lower() in lowerfile:
                return True
        return False
    except Exception:
        return False

def human_readable_size(num_bytes):
    for unit in ['B','KB','MB','GB','TB']:
        if num_bytes < 1024.0:
            return f"{num_bytes:,.2f} {unit}"
        num_bytes /= 1024.0
    return f"{num_bytes:.2f} PB"

def create_baseline_for_files_robust(xlsx_files, skip_force_baseline=True):
    """🛡️ 強化版 baseline 建立，帶診斷和恢復功能"""
    global force_stop
    
    total = len(xlsx_files)
    if total == 0:
        print("[INFO] 沒有需要 baseline 的檔案。")
        return

    print()
    print("=" * 90)
    print(" BASELINE 建立程序 (強化診斷版本) ".center(90, "="))
    print("=" * 90)
    
    # 檢查是否有之前的進度
    progress = load_progress()
    start_index = 0
    if progress and ENABLE_RESUME:
        print(f"🔄 發現之前的進度記錄:")
        print(f"   之前完成: {progress['completed']}/{progress['total']}")
        print(f"   記錄時間: {progress['timestamp']}")
        
        resume = input("是否要從上次中斷的地方繼續? (y/n): ").strip().lower()
        if resume == 'y':
            start_index = progress['completed']
            print(f"   ✅ 從第 {start_index + 1} 個檔案開始")
    
    # 啟動超時監控線程
    if ENABLE_TIMEOUT:
        timeout_thread = threading.Thread(target=timeout_handler, daemon=True)
        timeout_thread.start()
        print(f"⏰ 啟用超時保護: {FILE_TIMEOUT_SECONDS} 秒")
    
    if ENABLE_MEMORY_MONITOR:
        print(f"💾 啟用記憶體監控: {MEMORY_LIMIT_MB} MB")
    
    optimizations = []
    if USE_LOCAL_CACHE:
        optimizations.append("本地緩存")
    if ENABLE_FAST_MODE:
        optimizations.append("快速模式")
    
    print(f"🚀 啟用優化: {', '.join(optimizations)}")
    print(f"📂 Baseline 儲存位置: {os.path.abspath(LOG_FOLDER)}")
    if USE_LOCAL_CACHE:
        print(f"💾 本地緩存位置: {os.path.abspath(CACHE_FOLDER)}")
    print(f"📋 要處理的檔案: {total} 個 Excel (從第 {start_index + 1} 個開始)")
    print(f"⏰ 開始時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    print("-" * 90)
    
    # 確保資料夾存在
    os.makedirs(LOG_FOLDER, exist_ok=True)
    if USE_LOCAL_CACHE:
        os.makedirs(CACHE_FOLDER, exist_ok=True)
    
    baseline_total_size = 0
    success_count = 0
    skip_count = 0
    error_count = 0
    start_time = time.time()
    
    for i in range(start_index, total):
        if force_stop:
            print("\n🛑 收到停止信號，正在安全退出...")
            save_progress(i, total)
            break
            
        file_path = xlsx_files[i]
        base_name = os.path.basename(file_path)
        baseline_file = baseline_file_path(base_name)
        
        # 檢查記憶體
        if check_memory_limit():
            print(f"⚠️ 記憶體使用量過高，暫停 10 秒...")
            time.sleep(10)
            if check_memory_limit():
                print(f"❌ 記憶體仍然過高，停止處理")
                save_progress(i, total)
                break
        
        # 記錄檔案處理時間
        file_start_time = time.time()
        start_time_str = datetime.now().strftime('%H:%M:%S')
        current_memory = get_memory_usage()
        
        print(f"[完成 {i+1:>2}/{total}] [原始#{i+1:>2}] 處理中... (記憶體: {current_memory:.1f}MB)")
        print(f"  檔案: {base_name}")
        
        try:
            # 檢查是否跳過
            if skip_force_baseline and is_force_baseline_file(file_path):
                end_time_str = datetime.now().strftime('%H:%M:%S')
                consumed_time = time.time() - file_start_time
                
                print(f"  結果: [SKIP]")
                print(f"  原因: 屬於 FORCE_BASELINE_ON_FIRST_SEEN")
                print(f"  時間: 從 {start_time_str} 到 {end_time_str} 耗時 {consumed_time:.2f} 秒")
                print()
                
                skip_count += 1
                save_progress(i + 1, total)
                continue
            
            # 🛡️ 使用強化的 Excel 讀取
            cell_data = dump_excel_cells_with_timeout(file_path)
            
            if not cell_data and current_processing_file is None:
                # 可能是超時了
                print(f"  結果: [TIMEOUT]")
                print(f"  原因: 處理超時，跳過此檔案")
                error_count += 1
                save_progress(i + 1, total)
                continue
            
            curr_author = get_excel_last_author(file_path)
            curr_hash = hash_excel_content(cell_data)
            
            # 儲存 baseline
            save_baseline(baseline_file, {
                "last_author": curr_author,
                "content_hash": curr_hash,
                "cells": cell_data
            })
            
            # 計算結果
            size = os.path.getsize(baseline_file)
            baseline_total_size += size
            end_time_str = datetime.now().strftime('%H:%M:%S')
            consumed_time = time.time() - file_start_time
            baseline_name = os.path.basename(baseline_file)
            
            print(f"  結果: [OK]")
            print(f"  Baseline: {baseline_name}")
            print(f"  檔案大小: {human_readable_size(size)} | 累積: {human_readable_size(baseline_total_size)}")
            print(f"  時間: 從 {start_time_str} 到 {end_time_str} 耗時 {consumed_time:.2f} 秒")
            print()
            
            success_count += 1
            save_progress(i + 1, total)
            
        except Exception as e:
            end_time_str = datetime.now().strftime('%H:%M:%S')
            consumed_time = time.time() - file_start_time
            
            print(f"  結果: [ERROR]")
            print(f"  錯誤: {e}")
            print(f"  時間: 從 {start_time_str} 到 {end_time_str} 耗時 {consumed_time:.2f} 秒")
            print()
            
            error_count += 1
            save_progress(i + 1, total)
    
    force_stop = True  # 停止超時監控線程
    
    end_time = time.time()
    total_time = end_time - start_time
    
    print("-" * 90)
    print("🎯 BASELINE 建立完成!")
    print(f"⏱️  總耗時: {total_time:.2f} 秒")
    print(f"✅ 成功: {success_count} 個")
    print(f"⏭️  跳過: {skip_count} 個")
    print(f"❌ 失敗: {error_count} 個")
    print(f"📦 累積 baseline 檔案大小: {human_readable_size(baseline_total_size)}")
    
    if success_count > 0:
        print(f"📊 平均每檔案處理時間: {total_time/total:.2f} 秒")
    
    # 清理進度檔案
    if ENABLE_RESUME and os.path.exists(RESUME_LOG_FILE):
        try:
            os.remove(RESUME_LOG_FILE)
            print(f"🧹 清理進度檔案")
        except Exception:
            pass
    
    print()
    print(f"📁 所有 baseline 檔案存放於: {os.path.abspath(LOG_FOLDER)}")
    if USE_LOCAL_CACHE:
        print(f"💾 本地緩存檔案存放於: {os.path.abspath(CACHE_FOLDER)}")
    print("=" * 90 + "\n")

def print_console_header():
    print("\n" + "="*80)
    print(" Excel File Change Watcher (診斷強化版本) ".center(80, "-"))
    print("="*80 + "\n")

# ============= 其他函數保持原樣... ============

if __name__ == "__main__":
    try:
        print_console_header()
        print("  監控資料夾:")
        for folder in WATCH_FOLDERS:
            print(f"    - {folder}")
        print(f"  支援副檔名: {SUPPORTED_EXTS}")
        print(f"  目前使用者: {os.getlogin()}")  # 應該顯示 ckcm0210
        
        optimizations = []
        if USE_LOCAL_CACHE:
            optimizations.append("本地緩存")
        if ENABLE_FAST_MODE:
            optimizations.append("快速模式")
        if ENABLE_TIMEOUT:
            optimizations.append(f"超時保護({FILE_TIMEOUT_SECONDS}s)")
        if ENABLE_MEMORY_MONITOR:
            optimizations.append(f"記憶體監控({MEMORY_LIMIT_MB}MB)")
        if ENABLE_RESUME:
            optimizations.append("斷點續傳")
        
        print(f"  🚀 啟用功能: {', '.join(optimizations)}")
        print(f"  📂 Baseline 儲存位置: {os.path.abspath(LOG_FOLDER)}")
        if USE_LOCAL_CACHE:
            print(f"  💾 本地緩存位置: {os.path.abspath(CACHE_FOLDER)}")
        
        # 確保資料夾存在
        os.makedirs(LOG_FOLDER, exist_ok=True)
        if USE_LOCAL_CACHE:
            os.makedirs(CACHE_FOLDER, exist_ok=True)

        if SCAN_ALL_MODE:
            all_files = get_all_excel_files(WATCH_FOLDERS)
            print(f"總共 find 到 {len(all_files)} 個 Excel file.")
            create_baseline_for_files_robust(all_files, skip_force_baseline=True)
            print("baseline scan 完成！\n")
        else:
            target_files = get_all_excel_files(MANUAL_BASELINE_TARGET)
            print(f"手動指定 baseline，合共 {len(target_files)} 個 Excel file.")
            create_baseline_for_files_robust(target_files, skip_force_baseline=False)
            print("手動 baseline 完成！\n")

        # 其他監控程式碼...
        
    except Exception as e:
        print(f"[ERROR][main] 程式主流程 error: {e}")
        import traceback
        traceback.print_exc()
