# H公司程式工具合集 / HCompany Python Utilities

這個專案包含五個由我在 H 公司開發的資料處理與分析工具，主要用於網路分析儀、溫度計、以及 UD 系列超音波驅動器資料的整併、分析與圖表產出。  
The repository contains five Python-based data processing and analysis tools developed at H Company, used for Network Analyzers, thermometers, and UD series ultrasonic driver data consolidation, analysis, and chart generation.

---

## 工具清單 / Tools Overview

| 工具名稱 / Tool | 版本 / Version | 功能概覽 / Features | 檔案位置 / Folder |
| --------------- | -------------- | ------------------ | ---------------- |
| ENA Merge Files | v4.8.1 | ENA 資料合併、Excel 圖表、峰谷點分析<br>ENA data merging, Excel chart generation, peak/trough analysis | ENA_merge_files/ |
| Tempature Convert | v3.2.2 | 溫度資料整併、時間修正、Excel 圖表<br>Temperature data consolidation, time correction, Excel charts | Tempature_convert/ |
| UD7 HMI & Temperature Merge | v1.0.1 | HMI + 溫度資料合併報表與圖表<br>Merge HMI & Temperature data, generate reports and charts | UD7_HMI_Temp_Merge/ |
| UD2 & UD7 Merge Files | v3.1.5 | UD2/UD7 掃頻資料合併與共振點分析<br>UD2/UD7 sweep data merging & resonance point analysis | UD2_UD7_Merge/ |
| UD7 HMI Convert | v1.2.5 | UD7 監控 CSV 資料整併與圖表分析<br>UD7 monitoring CSV data consolidation & chart analysis | UD7_HMI_Convert/ |

---

## 使用方式 / How to Use

每個工具資料夾內皆有獨立 `README.md`，包含詳細的開發背景、功能說明、操作步驟與套件需求。  
Each tool folder contains its own `README.md` with detailed development background, features, usage instructions, and dependencies.

1. 進入工具資料夾 / Enter the tool folder  
2. 參考資料夾內 README.md / Refer to the README.md in the folder  
3. 執行主程式，依照說明操作 / Run the main script and follow the instructions  

---

## 系統需求 / Requirements

- 作業系統 / OS: Windows  
- Python 3.8 以上 / Python 3.8 or higher  
- 各工具所需套件 / Dependencies: 請參考各資料夾 README.md  

---

## 注意事項 / Notes

- 請勿在 Excel 開啟狀態下執行工具，以免儲存失敗。  
  Do not run the tools while Excel files are open to avoid save failures.  
- 每次操作僅處理指定資料夾或檔案，請確認格式正確。  
  Each operation processes only the specified folder or files; ensure formats are correct.  
- 詳細操作與範例請參考各工具 README.md  
  Refer to each tool's README.md for detailed usage and examples.
