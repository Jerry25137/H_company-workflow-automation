# Tempature Convert v3.2.2

本工具支援 **PICO TC-08** 與 **88598 AZ EB 溫度計** 所輸出的 `.TXT` / `.CSV` 檔案，進行溫度資料整併、通道命名、時間修正與 Excel 圖表輸出。  
目標是 **「簡化溫度資料處理流程，快速完成報表視覺化」**，協助現場工程人員更有效率地產出分析成果。  

This tool supports `.TXT` / `.CSV` files exported from **PICO TC-08** and **88598 AZ EB thermometers**, performing data merging, channel naming, time correction, and Excel chart output.  
Its goal is **"Simplify temperature data processing and quickly generate visual reports"**, helping on-site engineers produce analysis results more efficiently.

---

## 開發背景
### Development Background

1. 測溫資料往往格式不一致、通道多、難以對齊與繪圖，增加工程人員處理成本。  
   Temperature data often has inconsistent formats, multiple channels, and alignment difficulties, increasing processing effort for engineers.

2. 現場常用設備包括 **PICO TC-08** 與 **AZ EB 88598**，為節省轉檔與整理時間，本工具統一支援並自動辨識格式。  
   Common on-site equipment includes **PICO TC-08** and **AZ EB 88598**. This tool automatically detects formats to save conversion and organization time.

3. 採用 GUI 圖形介面操作，降低門檻，讓沒有寫程式經驗的使用者也能快速上手。  
   Uses a GUI interface to lower the barrier, allowing users without programming experience to quickly operate the tool.

4. 溫度資料的時間修正，以解決溫度計時間失準問題。  
   Time correction ensures accurate timestamps, resolving clock drift issues in thermometers.

---

## 功能總覽
### Features Overview

1. 📥 **自動辨識設備格式**（PICO `.CSV` / AZ `.TXT`），自動解析資料欄位。  
   📥 **Automatic device format detection** (PICO `.TXT` / AZ `.CSV`), automatically parses data columns.

2. ✅ **通道選擇與自訂標籤**，支援最多 8 組熱電偶通道輸出。  
   ✅ **Channel selection and custom labels**, supports up to 8 thermocouple channels.

3. ⏱️ **起始時間修正功能**（限 AZ `.TXT`），自動依照時間間隔補齊時間軸。  
   ⏱️ **Start time correction** (AZ `.TXT` only), automatically fills the time axis according to intervals.

4. 📊 匯出資料與圖表至 Excel（`Tempature_Output.xlsx`），內含平滑曲線、多通道色彩區分。  
   📊 Export data and charts to Excel (`Tempature_Output.xlsx`) with smooth curves and multi-channel color differentiation.

5. 🧩 對 PICO `.CSV` 格式會自動補足缺失的通道欄，保證 Excel 格式一致。  
   🧩 For PICO `.CSV` files, missing channels are automatically filled to ensure consistent Excel formatting.

---

## 支援格式對照
### Supported File Formats

| 檔案類型 | 設備            | 特徵欄位                                | 時間處理方式          |
|----------|-----------------|----------------------------------------|----------------------|
| `.TXT`   | PICO TC-08      | 有 `時間間隔欄`（如 30s、1m）           | 可設定起始時間，自動推算 |
| `.CSV`   | 88598 AZ EB     | 第一欄為 `Date time`，後續為溫度欄位   | 已內含絕對時間        |

| File Type | Device          | Key Columns                             | Time Handling         |
|-----------|-----------------|----------------------------------------|----------------------|
| `.TXT`    | PICO TC-08      | Interval column (e.g., 30s, 1m)        | Start time adjustable, automatically calculated |
| `.CSV`    | 88598 AZ EB     | First column `Date time`, subsequent temperature columns | Absolute time included |

---

## 操作方式
### How to Use

1. 執行主程式（例如 `Tempature_convert v3.2.1 (UI).py`）。  
   Run the main program (e.g., `Tempature_convert v3.2.1 (UI).py`).

2. 點選【瀏覽】，選擇 `.TXT`（PICO TC-08）或 `.CSV`（AZ EB 88598）資料檔案。  
   Click [Browse] and select `.TXT` (PICO TC-08) or `.CSV` (AZ EB 88598) data files.

3. 勾選要輸出的通道，可修改每個通道顯示名稱。  
   Select the channels to export and optionally rename each channel.

4. 若為 `.TXT`，可展開視窗啟用「起始時間修正」。  
   For `.TXT` files, expand the window to enable "Start Time Correction."

5. 點選【執行】，自動產出 `Tempature_Output.xlsx` 報表，儲存在原始資料夾。  
   Click [Run], automatically generate `Tempature_Output.xlsx` report, saved in the original folder.

---

## 注意事項
### Notes / Precautions

1. 匯出前請關閉 `Tempature_Output.xlsx`，避免儲存失敗。  
   Close `Tempature_Output.xlsx` before exporting to avoid save failures.

2. `.TXT` 資料來源需為 PICO TC-08 原始格式，時間欄需帶有秒 / 分單位（如 30s、1m）。  
   `.TXT` source files must be in PICO TC-08 original format, with time column in seconds/minutes (e.g., 30s, 1m).

3. `.CSV` 欄位須包含 `Date time` 與至少一筆溫度資料（會自動補足缺漏通道）。  
   `.CSV` files must include `Date time` column and at least one temperature column (missing channels are auto-filled).

4. 若未選擇任何通道或檔案格式錯誤，程式將跳出提示訊息。  
   If no channels are selected or file format is incorrect, the program will show a warning.

---

## 執行需求
### Requirements

- 作業系統：Windows（建議中文化介面） / OS: Windows (Chinese locale recommended)  
- Python 3.8 以上版本 / Python 3.8 or higher  
- 相依套件 / Dependencies:  
  - `openpyxl`  
  - `tkcalendar`  
  - `tkinter` (built-in with Python)

---
