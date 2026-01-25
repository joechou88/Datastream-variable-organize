import re
import os
import openpyxl
import glob

# ========= 自動抓資料夾裡 all- 開頭的 xlsx，但排除 -renamed  =========
xlsx_files = [f for f in glob.glob("all-*.xlsx") if "-renamed" not in f]  # 抓所有以 all- 開頭的 xlsx
if not xlsx_files:
    print("找不到 all- 開頭的 xlsx 檔案")
    exit()

# ========= 如果找到多個檔案，列出給使用者選 =========
if len(xlsx_files) > 1:
    print("找到多個符合條件的 xlsx 檔案：")
    for i, f in enumerate(xlsx_files, 1):
        print(f"{i}. {f}")
    while True:
        choice = input(f"請輸入要處理的檔案的國家數量: ").strip()
        matched_files = [f for f in xlsx_files if re.search(rf"all-{choice}countries\.xlsx", f)]
        if matched_files:
            original_file = matched_files[0]
            break
        print("找不到符合條件的檔案，請重新輸入")
else:
    original_file = xlsx_files[0]

print(f"你選擇的檔案是: {original_file}")

# 去掉副檔名
base_name = os.path.splitext(os.path.basename(original_file))[0]

# 抓 all- 後面的部分
m = re.match(r"all-(.+)", base_name)
country_count = m.group(1) if m else "all"

# 自動生成 input / output 檔名
input_file = original_file  # 直接用找到的檔名
output_file = f"all-{country_count}-renamed.xlsx"

# ========= 檢查檔案是否存在 =========
if os.path.exists(output_file):
    ans = input(f"檔案 '{output_file}' 已存在，是否刪除並生成新檔？(y/n): ").strip().lower()
    if ans != 'y':
        print("取消操作，程式結束。")
        exit()
    else:
        os.remove(output_file)
        print(f"已刪除舊檔 '{output_file}'。")


wb = openpyxl.load_workbook(input_file)
ws = wb.active  # 假設只改第一個工作表

# ========= 欄位名稱重命名函數 =========
def rename_col(col):    
    """
    處理欄位名稱：
    1. X(WC01254)           → WC01254
    2. X(WC06705)~U         → XWC06705U
    3. X(WC02051)~U$.1      → XWC02051U
    4. X(WC18545)~U$        → XWC18545U
    5. X(WC04601)~US        → XWC04601U
    6. Type                 → DSCD
    其他欄位保持不變
    """

    if col == "Type":
        return "DSCD"
    
    # 處理美金欄位
    m = re.match(r"^([A-Z])\((WC\d+)\)~([A-Z]+)(\$(?:\.\d+)?)?$", col)
    if m:
        return f"{m.group(1)}{m.group(2)}U"
    
    # 處理 X(WC01254) → WC01254
    m2 = re.match(r"^[A-Z]\((WC\d+)\)$", col)
    if m2:
        return m2.group(1)
    
    return col

for col_cell in ws[1]:
    old_value = col_cell.value
    new_value = rename_col(col_cell.value)
    if old_value != new_value:
        col_cell.value = new_value
        print(f"{old_value} → {new_value}")

# ========= 儲存新檔案 =========
wb.save(output_file)
print(f"已生成新檔案 '{output_file}'。")
