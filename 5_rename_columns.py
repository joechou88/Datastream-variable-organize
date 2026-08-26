import pandas as pd
import openpyxl
import re
import os
import glob

def rename_col(col):    
    """
    Rename rules:
    1. Type                 -> DSCD
    2. X(WC01254)           -> WC01254
    3. X(WC06705)~U         -> XWC06705U
    4. X(WC02051)~U$.1      -> XWC02051U
    5. X(WC18545)~U$        -> XWC18545U
    6. X(WC04601)~US        -> XWC04601U
    """

    if not isinstance(col, str):
        return col
    
    if col == "Type":
        return "DSCD"
    
    # Regex for USD-related columns
    m = re.match(r"^([A-Z])\((WC\d+)\)~([A-Z]+)(\$(?:\.\d+)?)?$", col)
    if m:
        return f"{m.group(1)}{m.group(2)}U"
    
    # Regex for standard X(...) columns
    m2 = re.match(r"^[A-Z]\((WC\d+)\)$", col)
    if m2:
        return m2.group(1)
    
    return col

while True:
    file_type = input("Select file format (csv or xlsx): ").strip().lower()
    if file_type in ['csv', 'xlsx']:
        break
    print("Invalid input. Please enter csv or xlsx.")

search_pattern = f"all-*.{file_type}"
files = [f for f in glob.glob(search_pattern) if "-renamed" not in f]

if not files:
    print(f"No files matching {search_pattern} found.")
    exit()

# Handle multiple files selection
if len(files) > 1:
    print(f"Multiple matching {file_type} files found:")
    for i, f in enumerate(files, 1):
        print(f"{i}. {f}")
    while True:
        choice = input("Enter the country count for the file to process: ").strip()
        matched_files = [f for f in files if re.search(rf"all-{choice}countries\.{file_type}", f)]
        if matched_files:
            input_file = matched_files[0]
            break
        print("No matching file found. Please try again.")
else:
    input_file = files[0]

print(f"Selected file: {input_file}")

base_name = os.path.splitext(os.path.basename(input_file))[0]
m = re.match(r"all-(.+)", base_name)
country_count = m.group(1) if m else "all"
output_file = f"all-{country_count}-renamed.{file_type}"

if os.path.exists(output_file):
    ans = input(f"File '{output_file}' already exists. Overwrite? (y/n): ").strip().lower()
    if ans != 'y':
        print("Operation cancelled. Exiting.")
        exit()
    else:
        os.remove(output_file)
        print(f"Deleted existing file '{output_file}'.")

print("\nColumn renaming summary:")

if file_type == 'csv':
    df = pd.read_csv(input_file, dtype=str)
    original_columns = df.columns.tolist() 
    df = df.rename(columns=rename_col)
    new_columns = df.columns.tolist()      
    
    for old_value, new_value in zip(original_columns, new_columns):
        if old_value != new_value:
            print(f"{old_value} -> {new_value}")
            
    df.to_csv(output_file, index=False)

elif file_type == 'xlsx':
    wb = openpyxl.load_workbook(input_file)
    ws = wb.active 
    
    for col_cell in ws[1]:
        old_value = col_cell.value
        new_value = rename_col(old_value)
        if old_value != new_value:
            col_cell.value = new_value
            print(f"{old_value} -> {new_value}")
            
    wb.save(output_file)

print(f"\nSuccessfully generated '{output_file}'.")
