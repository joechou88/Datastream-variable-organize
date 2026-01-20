## Workspace（原 Datastream）下載 Excel 須知

用 Excel 下載 Workspace（原 Datastream）的教學，請參考：

由於 Workspace（原 Datastream）有下載限制，無法一次輸出多國多年資料，甚至有些變數太多的國家同一年也沒辦法一次載下來，所以需要分段下載後，再利用本工具做資料整合。

下載 Workspace（原 Datastream）後，請**避免手動編輯 Excel 內容**。若直接修改儲存格，可能會殘留不可見的空白字元，進而導致程式執行錯誤。 

下載完成後，請將 Excel 檔案存放至：

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
   `./data-2015-2024 → ./all-countries.xlsx`

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
