import pandas as pd
import glob
import os
import sys
import config

if not os.path.exists(config.YEAR_OUTPUT_FOLDER):
    print(f"{config.YEAR_OUTPUT_FOLDER} not found.")
    input("Press Enter to leave...")
    sys.exit()

all_files = [f for f in glob.glob(os.path.join(config.YEAR_OUTPUT_FOLDER, "*.xlsx"))
             if not os.path.basename(f).startswith("~$")]

count = len(all_files)
print(f"Found {count} Excel files...")

if count > 0:
    final_excel_name = f"all-{count}countries.xlsx"
    final_csv_name   = f"all-{count}countries.csv"
    log_file_name    = f"all-{count}countries_integrate_log.txt"

    final_files = [
        final_excel_name,
        final_csv_name,
        log_file_name
    ]

    existing_files = [f for f in final_files if os.path.exists(f)]

    if existing_files:
        print("Same files detected, possibly generated during the last execution...")
        for f in existing_files:
            print(f" - {os.path.basename(f)}")

        ans = input("Overwrite these files? (y/n): ").strip().lower()
        if ans != "y":
            print("Operation cancelled. Retain old files and avoid overwriting.")
            sys.exit()
        else:
            for f in existing_files:
                try:
                    os.remove(f)
                    print(f"已刪除舊檔：{os.path.basename(f)}")
                except Exception as e:
                    print(f"[錯誤] 無法刪除 {f}: {e}")
    
    processed_files = set()
    if os.path.exists(log_file_name):
        with open(log_file_name, "r", encoding="utf-8") as f:
            processed_files = set(line.strip() for line in f)

    actual_merge_count = 0
    for filename in all_files:
        file_basename = os.path.basename(filename)
        
        # === 過濾區 ===
        if file_basename in processed_files:
            continue
        # =============
        
        try:
            df = pd.read_excel(filename, dtype=str)
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]  # 去掉多餘的空白欄

            print(f"正在合併: {file_basename} ({len(df.columns)} 欄)")

            file_exists = os.path.isfile(final_csv_name)
            df.to_csv(final_csv_name, mode='a', index=False, header=not file_exists, encoding='utf-8-sig')

            with open(log_file_name, "a", encoding="utf-8") as f:
                f.write(file_basename + "\n")
                
            actual_merge_count += 1
            
        except Exception as e:
            print(f"[錯誤] 讀取 {file_basename} 失敗: {e}")

    print("-" * 30)
    print(f"本次新增合併 {actual_merge_count} 個檔案。")

# ================= 轉存 Excel =================
if os.path.exists(final_csv_name):
    print(f"即將建立最終檔案: {final_excel_name}，可能會花幾分鐘...")
    
    user_input = input("是否要轉存為 Excel？(y/n): ").strip().lower()
    if user_input == "y":
        try:
            df_final = pd.read_csv(final_csv_name, dtype=str)  # Specify type as string to avoid DtypeWarning.
            df_final.to_excel(final_excel_name, index=False)
            print(f"\n★ Saved to: {final_excel_name}")
        except Exception as e:
            print(f"Failed to convert to Excel: {e}")
    else:
        print("跳過轉存 Excel，只輸出 CSV 檔案。")
else:
    if count == 0:
        print("沒有新檔案需要合併。")
