## Worldspace（原 Datastream）下載 Excel 須知

用 Excel 下載 Worldspace（原 Datastream）的教學，請參考：[Worldspace(原Datastream) Data Collection SOP](https://drive.google.com/file/d/1gW6l87DrgSTm3RZfPdKsc6xqmHRGRZBR/view?usp=drive_link)

由於 Worldspace（原 Datastream）有下載限制，無法一次輸出多國多年資料，甚至有些變數太多的國家同一年也沒辦法一次載下來，所以需要分段下載後，再利用本工具做資料整合。

下載完成後，請**避免手動編輯 Excel 內容**。若直接修改儲存格，可能會殘留不可見的空白字元，進而導致程式執行錯誤。 

請將 Excel 檔案存放至：

- 分年 + 分變數：`./data-split-by-variable`
- 分年：`./data`

---

## 程式執行流程

本工具的資料整合流程如下：

1. **同一國家多個變數合併（`variable-integrate.py`）**  
   `./data-split-by-variable → ./data`

2. **同一國家多年資料合併（`year-integrate.py`）**  
   `./data → ./data-2015-2024`

3. **整合所有國家資料（`country-integrate.py`）**  
   `./data-2015-2024 → ./all-countries.csv`

4. **重新命名欄位**  
   1. `rename-columns-csv.py`: `./all-{國家數量}countries.csv → ./all-{國家數量}countries-renamed.csv`  
   2. `rename-columns-xlsx.py`: `./all-{國家數量}countries.xlsx → ./all-{國家數量}countries-renamed.xlsx`

---

## 輸入檔案命名規則

程式支援 `.xlsm` 與 `.xlsx` 兩種格式。檔案命名規則如下：

### 【分年 + 分變數】（存放於 `./data-split-by-variable`）

- **國家-開始年-結束年(第幾組變數)**  
  - 例：`Switzerland-2015-2018A`

- **國家-年分(第幾組變數)**（僅單一年份）  
  - 例：`South-Korea-2015A`

### 【僅分年】（存放於 `./data`）

- **國家-開始年-結束年**  
  - 例：`Finland-2019-2020`

- **國家-年分**（僅單一年份）  
  - 例：`Finland-2021`

---

## 重新命名欄位
| 原欄位名稱           | 轉換後名稱   | 說明 |
|--------------------|------------|------|
| Type               | DSCD       | 固定重命名 |
| X(WC01254)         | WC01254    | 去掉外層 X() |
| X(WC06705)~U       | XWC06705U  | 去掉括號 & ~，保留末碼 |
| X(WC02051)~U$.1    | XWC02051U  | 去掉 .1 和 $，保留末碼 |
| X(WC18545)~U$      | XWC18545U  | 去掉 ~ 和 $ |
| X(WC04601)~US      | XWC04601U  | 尾碼統一為單個 U |

---

## 維護注意事項

1. 下載 Worldspace（原 Datastream）時，要確保每個國家各年度的變數，最後併起來應一致

2. 【分年 + 分變數】（存放於 `./data-split-by-variable`）合併時，沒有更新 REQUEST_TABLE N 欄的 Rows 數、P 欄的 Total 數，後面的開發者如果要拿這兩欄的值去做其他事，務必要先更新
