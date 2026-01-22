import pandas as pd
import re
import os

input_file = "all-countries.csv"
output_file = "all-countries-renamed.csv"

# ========= 檢查檔案是否存在 =========
if os.path.exists(output_file):
    ans = input(f"檔案 '{output_file}' 已存在，是否刪除並生成新檔？(y/n): ").strip().lower()
    if ans != 'y':
        print("取消操作，程式結束。")
        exit()
    else:
        os.remove(output_file)
        print(f"已刪除舊檔 '{output_file}'。")

df = pd.read_csv(input_file, dtype=str)

# ========= Type → DSCD =========
df = df.rename(columns={"Type": "DSCD"})

# ========= 美金相關欄位重新命名 & 處理 X(...) =========
def rename_col(col):
    """
    處理欄位名稱：
    1. X(WC01254)           → WC01254
    2. X(WC06705)~U         → XWC06705U
    3. X(WC02051)~U$.1      → XWC02051U
    4. X(WC18545)~U$        → XWC18545U
    5. X(WC04601)~US        → XWC04601U
    其他欄位保持不變
    """
    # 處理美金欄位
    m = re.match(r"^([A-Z])\((WC\d+)\)~([A-Z]+)(\$(?:\.\d+)?)?$", col)
    if m:
        return f"{m.group(1)}{m.group(2)}U"
    
    # 處理 X(WC01254) → WC01254
    m2 = re.match(r"^[A-Z]\((WC\d+)\)$", col)
    if m2:
        return m2.group(1)
    
    return col

df = df.rename(columns=rename_col)

df.to_csv(output_file, index=False)
print(f"已生成新檔案 '{output_file}'。")
