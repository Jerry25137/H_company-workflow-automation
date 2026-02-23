# UD2 & UD7 Merge Files v3.1.5

本工具專為處理 H公司 UD2 系列與 UD7 系列超音波驅動器掃頻數據而設計，旨在 **「簡化操作流程，提升分析效率」**，可快速完成資料合併、圖表產出與共振點分析。  
This tool is designed for processing Hcompany UD2 and UD7 series ultrasonic driver sweep data. Its goal is **"Simplify workflow and improve analysis efficiency"**, quickly completing data merging, chart generation, and resonance point analysis.

---

## 開發背景
### Development Background

1. 處理自動化：為協助工程端進行大量掃頻資料整合與視覺化，透過自動化合併與圖表產出，省去繁瑣的手動處理與比對。  
   Automation: To help engineers consolidate and visualize large amounts of sweep data, automating merging and chart generation to save manual effort.

2. GUI 介面：考量團隊中部分成員無程式背景，特別設計為單一腳本，並整合簡易 GUI，方便直接點擊操作，無需修改任何程式碼，即可直接上手。  
   GUI: Considering some team members lack programming experience, designed as a single script with simple GUI for easy point-and-click operation without modifying code.

3. 減少學習成本：使用者只需選擇資料夾與參數，即可自動完成資料處理與報告輸出，降低學習曲線。  
   Reduce learning curve: Users only need to select folders and parameters to automatically process data and generate reports.

4. 峰值偵測策略：針對現行「移動平均 + 斜率法」在驅動器(MCU)上資源耗用過高問題，自行設計並實作替代演算法，採用「三點法 + Distance + Prominence」策略進行峰值辨識，僅需比大小與減法計算，免去浮點運算與減少數據處理步驟，保有良好辨識準確度。  
   Peak detection strategy: Due to excessive MCU resource consumption of the current "moving average + slope method," a custom algorithm using "Three-point method + Distance + Prominence" was implemented, requiring only comparison and subtraction without floating-point computation, maintaining high detection accuracy.

---

## 峰值偵測開發挑戰與心得
### Peak Detection Challenges & Insights

1. `scipy.find_peaks` 提供的峰值偵測方式具有高度辨識準確度，其核心邏輯──結合「三點法 + Prominence + Distance」──非常適用於掃頻資料的共振點判定。  
   `scipy.find_peaks` provides accurate peak detection using "Three-point method + Prominence + Distance," suitable for resonance detection in sweep data.

2. 然而該套件流程封閉、彈性有限，無法直接應用於驅動器嵌入式環境（MCU）中對資源的高度要求，亦不利於後續客製化調整與優化。  
   However, its closed process and limited flexibility prevent direct use in resource-constrained MCU environments and hinder customization and optimization.

3. 為此，自行重構並實作可嵌入 MCU 的峰值偵測演算法，採用「三點法 + Distance + Prominence」邏輯，並模組化設計，增設可調參數，提升可讀性、維護性與適應彈性。  
   Therefore, a custom MCU-friendly peak detection algorithm was implemented using "Three-point + Distance + Prominence," modularized with adjustable parameters to improve readability, maintainability, and adaptability.

4. 同步導入視覺化分析工具，方便比對不同設定參數對峰值偵測結果的影響，提升非程式背景成員的理解與使用效率；完整專案已開源於 GitHub。  
   Integrated visualization tools allow comparison of parameter settings on peak detection results, improving understanding for non-programmers. The full project is open-sourced on GitHub.

5. 自訂演算法避免斜率法的浮點計算與掃頻預處理，僅透過簡單的比大小與減法操作，顯著降低 MCU 資源耗用，並保有良好準確率與反應速度，適合實際嵌入式應用。  
   Custom algorithm avoids floating-point computation and preprocessing, using simple comparisons and subtraction, reducing MCU resource usage while maintaining accuracy and responsiveness.

6. 專案中原已採用斜率法進行峰值判定，因具備一定物理意涵且流程已建立，主管對替代方案持保留態度。雖已完成新方法的測試與優劣比較說明，最終未被納入產品開發流程。  
   The original slope-based method was kept due to physical relevance and established workflow; management did not adopt the alternative method despite testing and comparison.

7. 雖感遺憾與挫折，仍保留完整成果於研發文件中，視為個人獨立開發的重要案例，也展現對嵌入式環境下演算法優化的理解與實踐能力。  
   Despite regret, the full results are preserved in documentation as an important personal development case, demonstrating understanding of algorithm optimization in embedded environments.

8. 期望未來能將這類更貼近實務限制與效能平衡的設計思維，應用於真正重視工程現實與效率導向的團隊與產品中。  
   Hope to apply this practical, performance-balanced design thinking in teams and products that value engineering reality and efficiency.

---

## 功能總覽
### Features Overview

1. 內建 Tkinter GUI 圖形化操作介面，所有設定皆透過點選完成。  
   Built-in Tkinter GUI, all settings configured via menu selection.

2. 支援匯入多筆漢鼎 UD2 系列與 UD7 系列超音波驅動器掃頻資料（CSV、TXT 格式均可）。  
   Supports importing multiple HanDing UD2 and UD7 ultrasonic driver sweep datasets (CSV and TXT formats).

3. 自動判別資料格式，區分兩款超音波驅動器數據（支援：頻率 / 電流、頻率 / 相位）。  
   Automatically detects data format, distinguishing two driver types (supports Frequency/Current and Frequency/Phase).

4. 設定刀把數量、刀具數量與循環次數後，可依順序自動整併資料。  
   After setting blade count, tool count, and cycles, data is automatically merged in sequence.

5. 產出合併後的 Excel，內含：  
   - 原始資料 / Original data  
   - 自動產生的散佈圖（含線條樣式、顏色、線寬調整） / Automatically generated scatter plots with style, color, and width adjustments  
   - 可選擇是否產出 Current / Phase 單獨分頁，或合併圖 / Optionally generate separate Current/Phase sheets or merged charts

6. 支援共振點分析（UD2 相位、UD7 電流）：  
   - 自動計算數據峰值與峰谷（Max/Min），相位或電流最大點為共振點  
     Automatically calculate data peaks/troughs; max phase or current = resonance point  
   - 使用三點法 + Distance 過濾 + Prominence 邏輯篩選顯著峰值與峰谷  
     Use Three-point + Distance + Prominence logic to filter significant peaks and troughs  
   - 生成單獨圖表或四個分析圖，完整展示判斷過程  
     Generate individual or four analysis charts, showing full judgment process

7. 支援空值補線（Excel 自動補點）、Unicode 中文顯示、檔名自動命名。  
   Supports missing value interpolation (Excel auto-fill), Unicode Chinese display, and automatic file naming.

---

## 操作方式
### How to Use

1. 執行 `UD2_&_UD7_ Merge_files v3.1.4 (UI).py`  
   Run `UD2_&_UD7_ Merge_files v3.1.4 (UI).py`

2. 點選「瀏覽」，選擇掃頻資料所在的資料夾（支援 CSV / TXT）  
   Click "Browse" and select the folder containing sweep data (CSV/TXT supported)

3. 設定 / Configure:  
   - 勾選要匯出的資料類型（電流 Current / 相位 Phase） / Select data type to export (Current / Phase)  
   - 是否啟用掃頻最大點分析（可調整 UD2/UD7 閾值條件篩選） / Enable sweep max point analysis (adjust UD2/UD7 threshold filter)  
   - 是否顯示共振點分析圖與數據標籤 / Show resonance analysis charts and labels  
   - 刀把數量、刀具數量、循環次數（控制 Excel 圖表線條顏色） / Blade count, tool count, cycles (control chart line color)  
   - 線條樣式（單色、實線、線寬） / Line style (single color, solid, width)

4. 點選「執行」，程式會自動：  
   Click "Run", the program will automatically:  
   - 讀取資料、合併 / Read and merge data  
   - 計算共振點(峰值/峰谷)、生成分析圖表（如啟用） / Calculate resonance points (peaks/troughs), generate charts if enabled  
   - 產出 Excel 檔（存於同資料夾，檔名自動命名） / Export Excel file (saved in same folder, auto-named)

---

## 注意事項
### Notes / Precautions

1. Excel 圖表最多支援 255 條資料線，如超出將顯示警告，只能合併數據，無法生成圖表。  
   Excel charts support up to 255 data lines; if exceeded, only merge data, charts cannot be generated.

2. 若開啟 Excel 檔案，儲存會失敗，請關閉後再執行。  
   If the Excel file is open, saving will fail. Close the file before running.

3. 此工具尚未模組化，如需整合至大型系統建議重構封裝。  
   This tool is not modular; for integration into large systems, refactoring is recommended.

---

## 執行需求
### Requirements

- 作業系統：Windows / OS: Windows  
- Python 3.8 以上 / Python 3.8 or higher  
- 安裝套件 / Dependencies:  
  - `tkinter`  
  - `openpyxl==3.1.3`  
  - `pywin32`  
  - `matplotlib`
