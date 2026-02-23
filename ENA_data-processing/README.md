# ENA Merge Files v4.8.1

本工具專為 E5061B ENA 系列網路分析儀 (Network Analyzer) 數據處理而設計，目標是 **「用最少的時間，達成最大的效益」**，快速完成資料合併、圖表產出與峰谷點分析。  
This tool is designed for E5061B ENA Series Network Analyzer data processing. Its goal is **"Achieve maximum efficiency in minimal time"**, quickly completing data merging, chart generation, and peak/trough analysis.

---

## 開發背景
### Development Background

1. 為解決繁瑣的人工數據處理流程，提升處理速度與準確度，自動完成資料合併與分析。  
   To solve cumbersome manual data processing, improve speed and accuracy, and automatically perform data merging and analysis.

2. 考量團隊中部分成員無程式背景，特別設計為單一腳本，並整合簡易 GUI，方便直接點擊操作，無需修改任何程式碼，即可直接上手。  
   Considering some team members lack programming experience, the tool is designed as a single script with a simple GUI for easy point-and-click operation without modifying any code.

3. 使用者只需選擇資料夾，一鍵即可輸出完整報告與圖表。  
   Users only need to select a folder, and with one click, a complete report and charts are generated.

---

## 功能總覽
### Features Overview

1. 內建 Tkinter 圖形化操作介面，工程端可直接使用。  
   Built-in Tkinter GUI for visual operation, engineers can use it directly.

2. 批次合併多筆 ENA 測試資料，支援多種數據格式。  
   Batch merge multiple ENA test datasets, supporting various data formats.

3. 自動分類、彙整與命名輸出檔案。  
   Automatically categorize, consolidate, and name output files.

4. 自動繪製 Excel 中的數據點與圖表。  
   Automatically plot data points and charts in Excel.

5. 峰值與谷值自動分析，結合三點法、距離 (Distance) 與顯著性 (Prominence) 篩選邏輯。  
   Automatic peak/trough analysis using three-point method, Distance, and Prominence filtering logic.

6. 產生峰谷點視覺化分析圖表，方便判讀。  
   Generate visual peak/trough analysis charts for easy interpretation.

7. 支援 Unicode 中文及中文變數命名，適用於在地使用環境。  
   Supports Unicode Chinese and Chinese variable names, suitable for local environments.

---

## 操作方式
### How to Use

1. 執行主程式（例如 `ENA_merge_files_v4.8.0.py`）。  
   Run the main program (e.g., `ENA_merge_files_v4.8.0.py`).

2. 點選「瀏覽」，選擇包含 ENA 測試資料的資料夾。  
   Click "Browse" and select the folder containing ENA test data.

3. 勾選所需參數（依據需求調整）。  
   Check the required parameters (adjust according to your needs).

4. 點擊「執行」，程式會自動：  
   Click "Run", the program will automatically:  
   - 讀取並合併資料 / Read and merge data  
   - 計算並分析峰值/峰谷點，分析圖表 / Calculate and analyze peaks/troughs, generate analysis charts  
   - 產出 Excel 報告與圖表，檔案會儲存在同資料夾內 / Export Excel reports and charts, saved in the same folder

---

## 注意事項
### Notes / Precautions

1. 請勿在 Excel 檔案開啟狀態下執行本工具，避免儲存失敗。  
   Do not run the tool while Excel files are open to prevent saving failure.

2. Excel 圖表最多支援 255 條資料線，超過時會顯示警告。  
   Excel charts support up to 255 data lines; a warning will appear if exceeded.

3. 峰谷點分析僅適用於具掃頻結構的資料（頻率 + 數值欄位）。  
   Peak/trough analysis is only applicable to data with a sweep structure (frequency + value columns).

4. 本工具為單一腳本，尚未模組化，若需整合至大型系統，建議重構封裝。  
   This tool is a single script and not modularized; for integration into large systems, refactoring and encapsulation are recommended.

---

## 執行需求
### Requirements

- 作業系統：Windows / OS: Windows

- Python 3.8 以上 / Python 3.8 or higher

- 依賴套件 / Dependencies:  
  - `tkinter`  
  - `openpyxl==3.1.3`  
  - `pywin32`  
  - `matplotlib`
