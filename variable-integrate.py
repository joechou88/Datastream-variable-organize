import os
import re
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from collections import defaultdict

DATA_SRC = "./data-split-by-variable"
DATA_OUT = "./data"

os.makedirs(DATA_OUT, exist_ok=True)

def parse_filename(fname):
    """
    解析檔名，例如：
    Denmark-2015A.xlsx
    Denmark-2015-2018B.xlsm
    回傳 (country, year_start, year_end, variable_tag)
    """
    name = os.path.splitext(fname)[0]
    m = re.match(r"(.+?)-(\d{4})(?:-(\d{4}))?([A-Z]+)$", name)

    if not m:
        return None

    country, y1, y2, vars_ = m.groups()
    y2 = y2 if y2 else y1

    return [
        (country, int(y1), int(y2), var, fname)
        for var in vars_
    ]

def get_expected_output_files(parsed, country_year_spans):
    outputs = {}  # out_path -> (country, year_label)

    for country, spans in country_year_spans.items():
        is_consistent, year_span_list = check_year_span_consistency(
            country, spans
        )
        if not is_consistent:
            continue

        for start_year, end_year in year_span_list:
            year_label = (
                f"{start_year}"
                if start_year == end_year
                else f"{start_year}-{end_year}"
            )
            fname = f"{country}-{year_label}.xlsx"
            out_path = os.path.join(DATA_OUT, fname)
            outputs[out_path] = (country, year_label)

    return outputs

def create_output_file(country, start_year, end_year):
    year_label = (
        f"{start_year}"
        if start_year == end_year
        else f"{start_year}-{end_year}"
    )

    fname = f"{country}-{year_label}.xlsx"

    out_path = os.path.join(DATA_OUT, fname)
    
    files = [f for f in os.listdir(DATA_SRC) if f.endswith((".xlsx", ".xlsm"))]

    try:
        template_fname = find_excel_file(country, start_year, "A", files)
    except FileNotFoundError:
        raise FileNotFoundError(
            f"❌ 找不到 {country}-{start_year}A.xlsx 或 .xlsm 作為模板"
        )

    template_path = os.path.join(DATA_SRC, template_fname)

    if not os.path.exists(template_path):
        raise FileNotFoundError(f"找不到檔案：{template_path}")
    
    wb = load_workbook(template_path)

    if "REQUEST_TABLE" not in wb.sheetnames:
        raise ValueError(f"{template_fname} 中沒有 REQUEST_TABLE 工作表")
    
    wb.save(out_path)
    
    return out_path

def find_excel_file(country, start_year, var_tag, files):
    """
    找出指定國家、年份、變數的檔案（A/B/C...）
    支援單年或跨年
    """
    # 精確匹配 country-startyear(-endyear)var_tag
    pattern = re.compile(
        rf"^{re.escape(country)}-{start_year}(?:-\d{{4}})?[A-Z]*{var_tag}[A-Z]*\.(xlsx|xlsm)$"
    )
    candidates = [f for f in files if pattern.match(f)]

    if not candidates:
        raise FileNotFoundError(
            f"❌ 找不到 {country}-{start_year}{var_tag}.xlsx 或 .xlsm"
        )

    # 如果剛好有兩個（理論上不應該），優先用 .xlsx
    candidates.sort(key=lambda x: x.endswith(".xlsm"))
    return candidates[0]


def check_year_span_consistency(country, year_spans):
    """
    1) 將所有年份標準化，單一年 -> (year, year)
    2) 依開始年排序，找出各個 year_span
    3) 確保同一 year_span 的 A/B/C/... 年段完全一致
    4) 後續 year_span 不能重疊先前 year_span 年份
    回傳：
        - is_consistent: True/False
        - year_span_list: list of (start_year, end_year)
    """
    # 標準化：只有一年 -> (year, year)
    normalized_year_span = [(s, s) if e is None else (s, e) for s, e in year_spans]

    # 依開始年排序
    normalized_year_span = sorted(normalized_year_span, key=lambda x: x[0])

    year_span_list = []
    current_start, current_end = normalized_year_span[0]

    for s, e in normalized_year_span[1:]:
        if s <= current_end:  # 屬於同一個 year_span
            if (s, e) != (current_start, current_end):
                print(f"\n🚨 {country}：同一 year_span A/B/C 年段不一致")
                print(f"  Expected：{current_start}-{current_end}")
                print(f"  Found：{s}-{e}")
                return False, None
        else:   # 新 year_span
            year_span_list.append((current_start, current_end))
            current_start, current_end = s, e

    year_span_list.append((current_start, current_end))  # 加最後一個 year_span

    # 檢查 year_span 之間不重疊
    for i in range(1, len(year_span_list)):
        prev_s, prev_e = year_span_list[i-1]
        curr_s, curr_e = year_span_list[i]
        if curr_s <= prev_e:
            print(f"\n🚨 {country}：與前一個 year_span 重疊")
            print(f"  前一個 year_span：{prev_s}-{prev_e}")
            print(f"  當前 year_span：{curr_s}-{curr_e}")
            return False, None

    return True, year_span_list

def read_request_table(xls_path):
    """讀取 REQUEST_TABLE，回傳 dataframe（row=7 開始）"""
    return pd.read_excel(
        xls_path, sheet_name="REQUEST_TABLE", engine="openpyxl", header=None
    )

def get_sheet_for_year(req_df, year):
    """根據 REQUEST_TABLE 找到對應年份的工作表位置"""
    
    # 從 row7 開始抓 G欄（index=6）
    df_years = pd.to_numeric(req_df.iloc[6:, 6], errors='coerce')
    matches = df_years[df_years == year]

    if matches.empty:
        print("🔍 DEBUG：REQUEST_TABLE G 欄 'Start Date'（前 5 筆）內容如下：")
        print(df_years.head(5).tolist())
        raise ValueError(f"⚠️ REQUEST_TABLE 找不到年份 {year}")

    # 取第一個符合年份的列索引
    row_idx = matches.index[0]
    row_series = req_df.iloc[row_idx]

    sheet_ref = row_series[10]      # K欄
    expected_rows = row_series[13]  # N欄
    expected_cols = row_series[14]  # O欄

    # sheet_ref 形如: 工作表1'!$A$1
    sheet_name = sheet_ref.split("!") [0].replace("'", "")

    return sheet_name, int(expected_rows), int(expected_cols), row_idx + 1

def read_variable_data(xls_path, sheet_name):
    """從指定 sheet 讀資料"""
    df = pd.read_excel(xls_path, sheet_name=sheet_name, engine="openpyxl")
    return df

def append_column(out_path, df, sheet_name):
    """
    將單一變數資料貼到 Merged 工作表
    df: 原工作表 dataframe
    var_tag: 變數組數（A/B/C…）
    """
    wb = load_workbook(out_path)

    # 如果工作表已存在就用原名，否則建立
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.create_sheet(title=sheet_name)
    
    if ws.max_column == 1 and ws.cell(row=1, column=1).value is None:
        new_col_idx = 1   # 如果是空表，就從 A 欄開始  
    else:
        new_col_idx = ws.max_column + 1     # 找「下一個空欄」，避免覆蓋

    rows_to_write = dataframe_to_rows(df.iloc[:, 1:], index=False, header=True)
    
    for r_idx, row in enumerate(rows_to_write, start=1):
        for c_idx, value in enumerate(row, start=new_col_idx):
            ws.cell(row=r_idx, column=c_idx, value=value)

    wb.save(out_path)

def update_request_table(out_path, src_path, excel_row):
    """
    將 src_path 的 REQUEST_TABLE 中
    對應 year 的 O/P 欄，加到 out_path 的 REQUEST_TABLE
    並印出加總過程
    """
    wb_out = load_workbook(out_path)
    wb_src = load_workbook(src_path)

    ws_out = wb_out["REQUEST_TABLE"]
    ws_src = wb_src["REQUEST_TABLE"]

    # ========= 先檢查 N 欄 (Rows) =========
    N_COL = 14  # column N

    n_out = ws_out.cell(row=excel_row, column=N_COL).value
    n_src = ws_src.cell(row=excel_row, column=N_COL).value

    if n_out != n_src:
        print(
            f"❌ ROWS 不一致 | "
            f"{os.path.basename(out_path)} N{excel_row}={n_out} | "
            f"{os.path.basename(src_path)} N{excel_row}={n_src}"
        )

    for col, label in [(15, "O"), (16, "P")]:
        v_out = ws_out.cell(row=excel_row, column=col).value or 0
        v_src = ws_src.cell(row=excel_row, column=col).value or 0
        new_v = v_out + v_src

        print(
            f"🧮 {os.path.basename(out_path)} "
            f"REQUEST_TABLE {label}{excel_row}: "
            f"{v_out} + {v_src} = {new_v}"
        )

        ws_out.cell(row=excel_row, column=col, value=new_v)

    wb_out.save(out_path)

def main():
    files = [f for f in os.listdir(DATA_SRC) if f.endswith((".xlsx", ".xlsm"))]

    parsed = []
    for f in files:
        parsed.extend(parse_filename(f))

    # 依國家 -> 年度 -> 變數排序（A, B, C...）
    grouped = defaultdict(lambda: defaultdict(list))  # country -> year -> list of (var, fname)
    country_year_spans = defaultdict(list)
    
    for country, y1, y2, var, fname in parsed:
        country_year_spans[country].append((y1, y2))
        for y in range(y1, y2 + 1):
            grouped[country][y].append((var, fname))
    
    # 檢查之前是否已輸出過
    expected_outputs = get_expected_output_files(parsed, country_year_spans)

    existing_outputs = {
        path: meta
        for path, meta in expected_outputs.items()
        if os.path.exists(path)
    }

    if existing_outputs:
        print("\n⚠️ 發現以下輸出檔已存在 ./data：")
        for i, (path, (country, year_label)) in enumerate(existing_outputs.items(), 1):
            print(f"{i}. {country} ({year_label}) → {os.path.basename(path)}")

        while True:
            ans = input(
                "\n👉 是否【全部刪除】後重新產生？ (y/n): "
            ).strip().lower()

            if ans == "y":
                for path in existing_outputs:
                    print(f"🗑️ 刪除 {os.path.basename(path)}")
                    os.remove(path)
                break

            elif ans == "n":
                print(
                    "\n⏭️ 未刪除任何檔案。\n"
                    "請自行至 ./data 刪除欲重新產生的檔案後再執行。"
                )
                return

            else:
                print("請輸入 y 或 n")

    for country, spans in grouped.items():

        # 先檢查該國所有檔案的年段是否一致
        is_consistent, year_span_list = check_year_span_consistency(
            country, country_year_spans[country]
        )
        if not is_consistent:
            continue   # 整個國家直接跳過，不輸出

        print(f"\n========== ▶ 開始處理 {country} ==========")

        for start_year, end_year in year_span_list:
            out_xlsx = create_output_file(country, start_year, end_year)
            if out_xlsx is None:
                continue   # 這個年度已做過，直接跳過
            skip_country = False

            # 篩選這個 block 的檔案
            block_files = [
                (y1, y2, var, fname)
                for parsed_country, y1, y2, var, fname in parsed
                if parsed_country == country and y1 >= start_year and y2 <= end_year
            ]
            block_files = sorted(block_files, key=lambda x: x[2])  # A/B/C 排序，以第一個最小字母先處理
            
            processed_files = set() # 記錄已處理 Excel

            for s, e, var, _ in block_files:
                fname = find_excel_file(country, s, var, files)
                if fname in processed_files:
                    continue  # 否則 Hong-Kong-2015CD 會被併 2 次
                processed_files.add(fname)  # 標記 Hong-Kong-2015CD 已處理

                src_path = os.path.join(DATA_SRC, fname)
                vars_in_file = [v for _, _, v, f in block_files if f == fname]
                is_first_variable = ("A" in vars_in_file)
                print(f"📂 處理 {src_path}")

                req_df = read_request_table(src_path)

                for year in range(s, e+1):
                    try:
                        sheet_name, exp_rows, exp_cols, excel_row = get_sheet_for_year(req_df, year)
                        df = read_variable_data(src_path, sheet_name)
                        df_rows, df_cols = df.shape  # DataFrame 不含 header，會少一 row

                        actual_rows = df_rows + 1
                        actual_cols = df_cols

                        # 檢查尺寸
                        if actual_rows != exp_rows or actual_cols != exp_cols:
                            print(f"⚠️ {country}-{start_year}-{end_year}{var} rows/cols 不符"
                                f"   Expected: {exp_rows} rows x {exp_cols} cols\n"
                                f"   Actual:   {actual_rows} rows x {actual_cols} cols"
                            )
                            skip_country = True
                            break
                        else:
                            print(f"🔹 工作表: {sheet_name}, shape: {exp_rows} rows x {exp_cols} columns")

                        if is_first_variable:   # A 組變數作為模板，已經在新檔裡，skip
                            continue

                        append_column(
                            out_path=out_xlsx,
                            df=df,
                            sheet_name=sheet_name
                        )

                        if not is_first_variable:
                            update_request_table(
                                out_path=out_xlsx,
                                src_path=src_path,
                                excel_row=excel_row
                            )
                    except Exception as e:
                        print(f"⚠️ ERROR: {e}")
                        skip_country = True
                        break   # 跳出 var 迴圈，外層會處理刪檔 + 換國

            if skip_country:
                if os.path.exists(out_xlsx):
                    print(f"🗑️ 刪除檔案 {out_xlsx}")
                    os.remove(out_xlsx)
                break   # 跳出 year 迴圈 (略過後續年度)，換下一國

    print("🎉 所有國家/年度整合完成！")

if __name__ == "__main__":
    main()
