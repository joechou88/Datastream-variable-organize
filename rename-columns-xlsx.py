import openpyxl
import re
import os

input_file = "all-countries.xlsx"

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

# ========= 計算國家數量 =========
# 國家欄在第二欄（B列），且從第2列開始
countries = set()
for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
    val = row[0].value
    if val:
        countries.add(val)

num_countries = len(countries)

# ========= 自動生成輸出檔名 =========
output_file = f"{num_countries}countries.xlsx"

# ========= 檢查檔案是否存在 =========
if os.path.exists(output_file):
    ans = input(f"檔案 '{output_file}' 已存在，是否刪除並生成新檔？(y/n): ").strip().lower()
    if ans != 'y':
        print("取消操作，程式結束。")
        exit()
    else:
        os.remove(output_file)
        print(f"已刪除舊檔 '{output_file}'。")

# ========= 儲存新檔案 =========
wb.save(output_file)
print(f"已生成新檔案 '{output_file}'。")
