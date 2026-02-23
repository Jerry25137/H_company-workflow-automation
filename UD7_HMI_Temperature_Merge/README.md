# UD7 HMI & Temperature Merge v1.0.1

本工具支援合併 **UD7 HMI** 資料與溫度資料（Excel `.xlsx`），自動生成合併報表與圖表。  
目標是 **「簡化資料整併流程，快速產生可視化分析」**，協助現場工程人員或測試人員更有效率地完成報告。

This tool merges **UD7 HMI** data and temperature data (Excel `.xlsx`) into a single report with charts.  
Its goal is **"Simplify data merging and quickly generate visual analysis"**, helping engineers or test staff produce reports efficiently.

---

## 開發背景
### Development Background

1. HMI 與溫度資料通常分散在不同檔案中，人工合併耗時且容易出錯。  
   HMI and temperature data are often stored in separate files; manual merging is time-consuming and error-prone.

2. 透過自動化程式，合併數據、產生 Excel 報表與圖表，減少手動操作成本。  
   This tool automatically merges data, generates Excel reports and charts, reducing manual effort.

3. 內建 GUI 檔案選取功能，使用者無需修改程式碼即可操作。  
   Built-in GUI allows users to select files without modifying code.

4. 支援雙資料來源合併（HMI 與溫度），自動調整圖表顏色與線型。  
   Supports merging two data sources (HMI and Temperature) and automatically adjusts chart colors and line styles.

---

## 功能總覽
### Features

1. 📥 **讀取 Excel 資料**，自動跳過空白列。  
   📥 **Read Excel files**, automatically skip empty rows.

2. 🔀 **自動合併 HMI 與溫度資料**，時間對齊。  
   🔀 **Automatically merge HMI and temperature data**, align by time.

3. 📊 **自動生成 Excel 圖表**：
   - 支援多線資料顏色循環
   - 左右 Y 軸分離
   - 自動設定圖表標題與 X/Y 軸
   - 支援溫度倍率 (x10)
   
   📊 **Automatically generate Excel charts**:
   - Multi-line color cycling
   - Separate left/right Y axes
   - Auto chart titles and axis labels
   - Supports temperature scaling (x10)

4. 🧩 **自動調整 Excel 圖表選項**：
   - 連接空白資料點
   - 關閉格線或設定刻度位置
   - 使用 pywin32 控制 Excel 後處理

   🧩 **Automatically configure Excel chart options**:
   - Connect blank data points
   - Customize gridlines and tick marks
   - Post-process charts using pywin32

5. 💾 **儲存合併報表**：
   - Excel 檔案自動命名為 `UD7_HMI+Tempature_Output.xlsx`
   - 存於使用者指定資料夾

   💾 **Save merged report**:
   - Excel file automatically named `UD7_HMI+Tempature_Output.xlsx`
   - Saved in user-selected folder

---

## 使用方式
### How to Use

1. 執行程式 `UD7_HMI+Temp_Merge.py`  
   Run the script `UD7_HMI+Temp_Merge.py`

2. 選擇第一個檔案（UD7 HMI 或溫度資料）  
   Select the first file (UD7 HMI or Temperature)

3. 選擇第二個檔案（UD7 HMI 或溫度資料）  
   Select the second file (UD7 HMI or Temperature)

4. 選擇資料儲存資料夾  
   Choose the folder to save merged report

5. 程式會自動產生合併報表與圖表，完成後跳出提示訊息  
   The script will automatically generate merged Excel report and charts, with completion message

---

## 注意事項
### Notes / Precautions

1. 請關閉 `UD7_HMI+Tempature_Output.xlsx` 以避免儲存失敗。  
   Close `UD7_HMI+Tempature_Output.xlsx` to avoid save errors.

2. 請確保輸入檔案為 `.xlsx` 格式。  
   Ensure input files are in `.xlsx` format.

3. 若合併失敗或出現錯誤，請檢查檔案是否有空列或格式異常。  
   If merging fails or errors occur, check for empty rows or invalid formats in the files.

---

## 執行需求
### Requirements

- 作業系統：Windows  
- Python 3.8 以上  
- 套件依賴：
  - `openpyxl`
  - `pywin32`
  - `tkinter`（Python 內建）

---
