import pandas as pd
import glob
import os
import sys

# ================= 設定區 =================

# 1. 取得腳本所在的基本路徑 (Base Path)
if getattr(sys, 'frozen', False):
    base_path = os.path.dirname(sys.executable)
elif '__file__' in locals():
    base_path = os.path.dirname(os.path.abspath(__file__))
else:
    base_path = os.getcwd()

# 2. 指定資料來源資料夾名稱 (根據你的描述是這個)
target_folder_name = "data-2015-2024"
input_path = os.path.join(base_path, target_folder_name)

# 3. 設定輸出檔名 (存放在腳本旁邊，也就是 datastream 資料夾)
final_excel_name = "all-countries.xlsx"
output_excel_path = os.path.join(base_path, final_excel_name)

# 4. 暫存檔與紀錄檔
temp_csv_name = "temp_merged_data.csv"
output_csv_path = os.path.join(base_path, temp_csv_name)

log_file_name = "processed_log.txt"
log_path = os.path.join(base_path, log_file_name)

# ==========================================

print(f"程式位置: {base_path}")
print(f"正在搜尋資料夾: {input_path}")
print("-" * 30)

# 檢查資料夾是否存在
if not os.path.exists(input_path):
    print(f"[錯誤] 找不到資料夾：{target_folder_name}")
    print(f"請確認你的目錄結構如下：")
    print(f"{base_path}\\")
    print(f"  └── {target_folder_name}\\ (Excel要放在這裡)")
    input("按 Enter 離開...")
    sys.exit()

# 讀取「已完成清單」
processed_files = set()
if os.path.exists(log_path):
    with open(log_path, "r", encoding="utf-8") as f:
        processed_files = set(line.strip() for line in f)

# 抓取 target_folder_name 內所有的 .xlsx
all_files = glob.glob(os.path.join(input_path, "*.xlsx"))
print(f"發現 {len(all_files)} 個 Excel 檔案。")

count = 0
for filename in all_files:
    file_basename = os.path.basename(filename)
    
    # === 過濾區 ===
    if file_basename.startswith("~$"): continue
    if file_basename in processed_files: continue
    # =============

    print(f"正在合併: {file_basename}")
    
    try:
        df = pd.read_excel(filename)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]  # 去掉多餘的空白欄
        print(f"{file_basename}: {len(df.columns)} 欄")
        
        # 寫入 CSV (存放在外面那一層，避免汙染資料夾)
        file_exists = os.path.isfile(output_csv_path)
        df.to_csv(output_csv_path, mode='a', index=False, header=not file_exists, encoding='utf-8-sig')
        
        # 寫入 Log
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(file_basename + "\n")
            
        count += 1
        
    except Exception as e:
        print(f"[錯誤] 讀取 {file_basename} 失敗: {e}")

print("-" * 30)
print(f"本次新增合併 {count} 個檔案。")

# ================= 轉存 Excel =================
if os.path.exists(output_csv_path):
    print(f"正在建立最終檔案: {final_excel_name}...")
    try:
        df_final = pd.read_csv(output_csv_path)
        df_final.to_excel(output_excel_path, index=False)
        print(f"\n★ 成功！檔案位置: {output_excel_path}")
        
        # 合併完成後刪除暫存 CSV (保持乾淨)
        # os.remove(output_csv_path) 
        
    except Exception as e:
        print(f"轉存 Excel 失敗: {e}")
else:
    if count == 0:
        print("沒有新檔案需要合併。")

input("\n程式執行完畢，按 Enter 離開...")
