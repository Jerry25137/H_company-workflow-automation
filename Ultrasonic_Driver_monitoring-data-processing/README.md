# UD7 HMI Convert v1.2.5

本工具專為 H公司 UD7 系列驅動器監控軟體，監控數據之 CSV 資料進行整併、篩選與圖表分析設計，目標是 **「簡化工程流程，加速資料處理」**，快速完成數據整合、視覺化與報表輸出。  
This tool is designed for HCompany UD7 series drives monitoring software. It processes CSV monitoring data for consolidation, filtering, and chart analysis. The goal is **"Simplify engineering workflow and accelerate data processing"**, quickly completing data integration, visualization, and report generation.

---

## 開發背景
### Development Background

1. 解決手動處理與判讀追頻資料繁瑣、容易出錯的問題，提升處理準確率與效率。  
   Solve the cumbersome and error-prone manual handling of frequency-tracking data, improving accuracy and efficiency.

2. 對象涵蓋具工程背景但不熟悉程式語言之使用者，因此採用單一腳本 + 圖形化界面（Tkinter）形式，方便點擊操作。  
   Target users include engineers not familiar with programming, so it uses a single script + GUI (Tkinter) for easy point-and-click operation.

3. 使用者只需選擇資料夾、設定時間區間與資料欄位，即可一鍵完成分析報告與圖表輸出。  
   Users only need to select the folder, set the time range and data columns, then generate analysis reports and charts with one click.

---

## 功能總覽
### Features Overview

1. 內建 Tkinter GUI 視覺化操作界面，無需編程背景即可使用。  
   Built-in Tkinter GUI for visual operation, no programming knowledge required.

2. 批次合併 `.csv` 驅動器監控資料資料，支援 `FREQ (頻率)`、`IFB (電流)`、`VFB (功率)` 三種欄位選擇。  
   Batch merge `.csv` drive monitoring data, supporting `FREQ`, `IFB`, and `VFB` column selection.

3. 自動依據 `追頻啟動` 與 `停止命令` 特徵訊號切分資料。  
   Automatically split data based on `frequency tracking start` and `stop command` signals.

4. 圖表以 `matplotlib` 實時預覽，或匯出至 Excel 完整呈現。  
   Charts can be previewed in real-time using `matplotlib` or exported to Excel for full presentation.

5. 每組追頻資料自動命名工作表，並在 Excel 中對應繪製多組雙 Y 軸圖表。  
   Each frequency-tracking dataset auto-names its worksheet and plots multiple dual Y-axis charts in Excel.

6. 內建錯誤追蹤機制，自動辨識未正常停止、錯誤警報與異常切換。  
   Built-in error tracking automatically identifies abnormal stops, alarms, and mode switching errors.

---

## 操作方式
### How to Use

1. 執行主程式（例如 `UD7_HMI_convert_v1.2.3 (UI).py`）。  
   Run the main program (e.g., `UD7_HMI_convert_v1.2.3 (UI).py`).

2. 點選「瀏覽」，選擇包含 UD7系列監控資料之 `.CSV` 檔案的資料夾。  
   Click "Browse" and select the folder containing UD7 series monitoring `.CSV` files.

3. 設定「搜尋時間區間」與欲輸出的資料欄位（FREQ / IFB / VFB）。  
   Set the "Search Time Range" and desired data columns to export (FREQ / IFB / VFB).

4. 點擊：  
   - 「**預覽**」→ 以 `matplotlib` 顯示即時圖形。  
     "Preview" → Display real-time charts using `matplotlib`.  
   - 「**執行**」→ 匯出含圖表之 Excel 報表。  
     "Run" → Export Excel report with charts.

5. 執行完成後，報表將自動儲存為 `UD7_HMI_Output.xlsx`，並保存在資料來源資料夾內。  
   After execution, the report will be automatically saved as `UD7_HMI_Output.xlsx` in the source folder.

---

## 注意事項
### Notes / Precautions

1. 請勿在 Excel 檔案開啟狀態下執行「執行」功能，否則可能導致儲存失敗。  
   Do not run the "Run" function while the Excel file is open; otherwise, saving may fail.

2. 每次操作僅處理一組資料夾內容，請確認 CSV 格式符合原始設備輸出結構。  
   Each operation processes only one folder at a time. Ensure CSV format matches the original device output structure.

3. 若系統偵測到下列情況，將跳出提示訊息：  
   If the system detects the following, a prompt will appear:  
   - UD7 Alarm 警報 / UD7 alarm  
   - 操作模式切換異常 / Mode switching anomaly  
   - 驅動器/HMI 程式未正常關閉 / Drive/HMI program not properly closed  
   - 提示訊息不影響資料合併 / Prompts do not affect data merging

4. 匯出圖表為雙 Y 軸設計，最多支援 FREQ、IFB、VFB 三種欄位並行顯示。  
   Exported charts use dual Y-axis design, supporting up to FREQ, IFB, and VFB simultaneously.

---

## 執行需求
### Requirements

- 作業系統：Windows（建議搭配中文化環境）  
  OS: Windows (Chinese locale recommended)

- Python 3.8 以上版本  
  Python 3.8 or higher

- 相依套件 / Dependencies:  
  - `tkinter`  
  - `tkcalendar`  
  - `openpyxl==3.1.3`  
  - `matplotlib`
