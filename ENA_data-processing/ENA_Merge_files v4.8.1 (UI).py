"""
Date：2025.05.29

Version：4.8.0

@author: Hsiao Yu-Chieh
"""

import os
import csv
import shutil
import time
import copy

from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import ParagraphProperties, CharacterProperties
import win32com.client as win32

import matplotlib
import matplotlib.pyplot as plt
plt.rcParams['font.family'] = 'Microsoft JhengHei' # 設定中文字體
plt.rcParams['axes.unicode_minus'] = False         # 設定普通的減號
matplotlib.use('Qt5Agg')   

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

version = "EE"

# 主程式模組----------------------------------------------------------------------------------------

# 10種線型
linetypes = ["solid", "sysDash", "sysDashDot", "sysDashDotDot", "sysDot",
             "dash", "dashDot", "dot", "lgDash", "lgDashDot", "lgDashDotDot",]

# 54種循環顏色
colors = ["4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646", "2C4D75", "772C2A", "5F7530", "4D3B62", 
          "276A7C", "B65708", "729ACA", "CD7371", "AFC97A", "9983B5", "6FBDD1", "F9AB6B", "3A679C", "9F3B38", 
          "7E9D40", "664F83", "358EA6", "F3740B", "95B3D7", "D99694", "C3D69B", "B3A2C7", "93CDDD", "FAC090", 
          "254061", "632523", "4F6228", "403152", "215968", "984807", "84A7D1", "D38482", "B9CF8B", "A692BE", 
          "81C5D7", "F9B67E", "335A88", "8B3431", "6F8938", "594573", "2E7C91", "D56509", "A7C0DE", "DFA8A6", 
          "CDDDAC", "BFB2D0", "A5D6E2", "FBCBA3", ]

# ENA 量測數據類型
data_types  = ["Z.csv", "O.csv", "C.csv", "X.csv", "R.csv", "L.csv", "Q.csv",]  # 數據類型
data_titles = ["阻抗", "相位", "電容", "感抗", "電阻", "電感", "品質因素"]         # 數據名稱
data_units  = ["ohm", "deg.", "nF", "ohm", "ohm", "mH", ""]                     # 數據單位

data_titles_en = ["Impedance", "Phase", "Capacitance", "Inductive Reactance", 
                  "Resistance", "Inductance", "Quality Factor"]

# 搜尋方向/顯著值設定 (刀把/變幅桿)
freq_set     = ((2000, 4000), (2000, 3000), (None, None), (2000, 2000), (2000, 2000), (None, None), (2000, 2000))

promi_finds  = (("valley", "valley"), ("valley", "peak"), (None, None), ("valley", "valley"), 
                ("valley", "valley"), (None, None), ("peak", "peak"))

promi_direct = (("left", "right"), ("right", "right"),  (None, None), ("left", "left"),  
                ("left", "left"),  (None, None), ("left", "left"))

# 讀取CSV檔(資料夾路徑, 控制模式)
def get_files_in_directory(f_path, mode):
    # 建立分類空集合，儲存名稱
    class_files_tag = []
    for i in range(len(data_types)):
        class_files_tag.append([])

    if mode == "dir_mode":
        # 取得目前目錄下的所有資料夾的路徑
        class_files_tag = [f.name for f in os.scandir(f_path) if f.is_dir()]
        

    elif mode == "file_mode":
        # 取得目前目錄中的所有檔案名，不分種類
        org_files = os.listdir(f_path)
    
        # 篩選出以 .csv 結尾的檔名，並將它們儲存到清單中
        csv_files = [filename for filename in org_files if filename.endswith(".csv") or filename.endswith(".CSV")]
         
        # 儲存檔案清單至分類內
        for f in csv_files:
            for i in range(len(data_types)):                
                if f[0] == data_types[i][0] or f[0] == chr(ord(data_types[i][0]) + 32):
                    class_files_tag[i].append(f)
    
    return class_files_tag

# 取得檔案路徑(資料夾路徑, 控制模式, 檔案清單)
def get_files_path(f_path, mode, f_tag):
    #  建立分類空集合，儲存路徑
    class_files_path = []
    for i in range(len(data_types)):
        class_files_path.append({})
    
    if mode == "dir_mode":
        for subfolder in f_tag:
            for i in range(len(data_types)):
                # 建立檔案路徑
                csv_file_path = os.path.join(f_path, subfolder)
                csv_file_path = os.path.join(csv_file_path, data_types[i])
                
                # 使用字典儲存檔案路徑
                if os.path.isfile(csv_file_path):
                    class_files_path[i][csv_file_path] = subfolder
                    
    elif mode == "file_mode":
        for f in f_tag:
            for i in range(len(f)):
                # 建立檔案路徑
                csv_file_path = os.path.join(f_path, f[i])

                # 使用字典儲存檔案路徑
                if os.path.isfile(csv_file_path):                    
                    class_files_path[f_tag.index(f)][csv_file_path] = f[i][:-4]
                                       
    return class_files_path

# ENA資料處理(資料夾路徑, 控制模式)
def ENA(f_path, mode):
    
    # 讀取路徑下檔案清單
    csv_files = get_files_in_directory(f_path, mode)
    
    # 取得檔案的路徑與資料標籤
    class_files_path = get_files_path(f_path, mode, csv_files)

    # 儲存全部資料用
    all_data = []
    
    # 用來追蹤是否有補充資料
    data_filled = False
      
    try:
        for n in range(len(class_files_path)):
            DATA = []
            
            for f in class_files_path[n]:
                # 檢查 .csv 文件是否存在
                if os.path.isfile(f):
                    # 讀取 .csv 的內容
                    with open(f, 'r', newline = '', encoding = 'utf-8-sig') as file:
                        reader = csv.reader(file)
                        
                        # 跳過前兩行
                        next(reader)
                        next(reader)
                        
                        if mode == "file_mode" and (class_files_path[n][f] + str(".csv")) == data_types[n]:
                            rows = [["Frequency", os.path.basename(f_path)]]
                        
                        else:
                            rows = [["Frequency", class_files_path[n][f]]]
                        
                        # 讀取並處理每一行
                        for row in reader:
                            try:
                                frequency      = float(row[0])  # 將 Frequency 欄位轉為 float
                                formatted_data = float(row[1])  # 將 Formatted Data 欄位轉為 float
                                rows.append([frequency, formatted_data])
                                
                            except ValueError:
                                continue  # 如果轉換失敗，跳過該行

                    if len(DATA) == 0: # 寫入第一個檔案的資料
                        DATA += rows
 
                    else: # 寫入第一個檔案之後的資料
                        DATA, filled = Merge_data(DATA, rows)
                        
                        if filled:  # 如果資料有補充
                            data_filled = True

            # 儲存全部資料
            all_data.append(DATA)
            #print(all_data) 
                   
    except Exception as e:
        print(f"ENA資料處理錯誤：{e}")

    return all_data, data_filled

# 資料合併模組------------------------------------------------------------------------------------------------

# 聯集合併資料(List A, List B)
def Merge_data(A, B):
    headers = A[0] + B[0][1:]
    body = {}
    data_filled = False

    freq_A = set(row[0] for row in A[1:])
    freq_B = set(row[0] for row in B[1:])

    # 只要 freq 不完全相同，就需要補空白
    if freq_A != freq_B:
        data_filled = True

    # 合併 freq：A、B 都有的 freq，以及獨有的 freq 都要
    all_freqs = sorted(freq_A.union(freq_B))
    default_row = ["" for _ in range(len(headers))]

    # 先處理 A 資料
    for row in A[1:]:
        freq = row[0]
        full_row = default_row[:]
        full_row[:len(row)] = row
        body[freq] = full_row

    # 再處理 B 資料
    for row in B[1:]:
        freq = row[0]
        if freq not in body:
            full_row = default_row[:]
            full_row[0] = freq
            body[freq] = full_row
        body[freq][len(A[0]):] = row[1:]

    # 組合結果
    merged = [headers] + [body[freq] for freq in all_freqs]
    return merged, data_filled

# 資料分類模組------------------------------------------------------------------------------------------------

# 檔案分類至資料夾
def Organize_files_to_folde(f_path, class_files):
    # 讀取分類List
    n = 0 # 控制碼
    for i in class_files:
        if i:
            new_folder_path = f"{f_path}/{data_types[n][0]}"
            
            # 建立分類資料夾
            if not os.path.exists(new_folder_path):                
                os.makedirs(new_folder_path)
                print(f'資料夾 {new_folder_path} 已建立')
            
            # 取得分類後檔案檔名，並移動至新資料夾
            for j in i:
                org_file_path = f"{f_path}/{j}"
                new_file_path = f"{f_path}/{data_types[n][0]}/{j}"
                
                if os.path.exists(new_folder_path):                    
                    shutil.move(org_file_path, new_file_path)                    
                    print(f"檔案 '{org_file_path}' 已移動到 '{new_file_path}'。")     
                    
        n += 1 # 控制碼

# 掃頻最大點--------------------------------------------------------------------------------------------------

# 找刀具接點 (數據類型, 數據, ΔmA)
def Find_peaks_max(data_type, no, DATA, ScanErr, show_analysis = None, labels = None):
    # 最大點 / 最小點
    Max_point = ["Max_point"]
    Min_point = ["Min_point"]
    for i2 in range(int(len(DATA[0]) - 1)):
        Max_point.append([])
        Min_point.append([])
        
    # 繪圖控制
    if show_analysis:
        if version == "EE":
            find_peaks_single_drw = False
            find_peaks_all_drw    = True
        
        elif version == "CE":
            find_peaks_single_drw = True
            find_peaks_all_drw    = False
            
    else:
        find_peaks_single_drw = False
        find_peaks_all_drw    = False
    
    # Step 1：初始化資料 ---------------------------------------------------------------------------
    Step0_time_start = time.perf_counter()
    
    orig_DATA, DATA, find_mode_list, frq_gap, direction, prominence = find_peaks_or_valleys(no, DATA, ScanErr)
    #print(orig_DATA)
    print(DATA)
    #print(find_mode_list)
    #print(frq_gap)
    #print(direction)
    #print(prominence)
        
    Step0_time_end = time.perf_counter()
    Step0_time = Step0_time_end - Step0_time_start
    print(f"Step 0：初始化資料，處理耗時 = {Step0_time:.10f}")
    
    try:
        for n in range(1, len(DATA[0])):
            print(f"資料：{DATA[0][n]}")
            # Step 1：建立資料 ---------------------------------------------------------------------------
            Step1_time_start = time.perf_counter()
            
            dStart, dEnd = find_data_ranges(DATA, n) # 取得資料位置
            valleys = []     # 波谷
            prominences = [] # 顯著值
            
            # 建立限制範圍
            I = dStart + 1
            while DATA[I][n] == "" and I <= dEnd:
                I += 1
            range_limit = int(frq_gap[n] / (DATA[I][0] - DATA[dStart][0]))
            
            Step1_time_end = time.perf_counter()
            Step1_time = Step1_time_end - Step1_time_start
            print(f"Step 1：建立參數，處理耗時 = {Step1_time:.10f}")
            
            # Step 2：3點法，找潛在峰值 -------------------------------------------------------------------
            Step2_time_start = time.perf_counter()
            
            gap = find_consecutive_max_length(DATA, n)
            peaks_3 = auto_find_best_gap(find_peak_3point, DATA, n, dStart, dEnd, max_gap = gap)
            peaks = peaks_3

            Step2_time_end = time.perf_counter()
            Step2_time = Step2_time_end - Step2_time_start
            print(f"Step 2：3點法，處理耗時 = {Step2_time:.10f}")
                        
            # Step 3：Distance 排除靠近者，保留區間主峰 -----------------------------------------------------
            Step3_time_start = time.perf_counter()
            
            peaks, valleys, prominences = Distance(DATA, n, peaks, valleys, prominences, dStart, dEnd, distance = range_limit)
            peaks_dist = peaks
            
            Step3_time_end = time.perf_counter()
            Step3_time = Step3_time_end - Step3_time_start
            print(f"Step 3：Distance，處理耗時 = {Step3_time:.10f}")
            
            # Step 4：Prominence 篩選顯著主峰---------------------------------------------------------------
            Step4_time_start = time.perf_counter()
            
            # 搜尋方向/顯著值
            P_side  = direction[n]
            P_promi = prominence[n]

            peaks, valleys, prominences = Prominence(DATA, n, peaks, dStart, dEnd, P_promi, search_range = range_limit, side = P_side)
            peaks_promi   = peaks
            valleys_promi = valleys
            
            Step4_time_end = time.perf_counter()
            Step4_time = Step4_time_end - Step4_time_start
            print(f"Step 4：Prominence，處理耗時 = {Step4_time:.10f}")
            print(f"找峰值，總共耗時 = {Step1_time + Step2_time + Step3_time + Step4_time:.10f}")
            print()
            print("---------------")
            print()
            
            # Step 5：還原反轉資料 -------------------------------------------------------------------------
            
            print(peaks_promi, valleys_promi)
            print(peaks, valleys)
            
            if find_mode_list[n] == "valley":
                peaks_promi, valleys_promi = rev_peaks_and_valleys(peaks_promi, valleys_promi)
                peaks, valleys = rev_peaks_and_valleys(peaks, valleys)
            
            # Step 6：儲存 Max / Min point ---------------------------------------------------------------
            Max_point[n] = " / ".join(f"{orig_DATA[p][0]}:{orig_DATA[p][n]}" for p in peaks)
            Min_point[n] = " / ".join(f"{orig_DATA[p][0]}:{orig_DATA[p][n]}" for p in valleys)
            
            # 繪圖分析 -----------------------------------------------------------------------------------
            # 獨立圖表
            if find_peaks_single_drw:
                plot_final_peak(data_type, no, orig_DATA, n, dStart, dEnd, 
                                peaks, valleys, show_data_labels = labels, find_mode = find_mode_list[n])
            
            # 完整圖表
            if find_peaks_all_drw:
                plot_find_max_steps(data_type, no, orig_DATA, n, dStart, dEnd,
                                    peaks_3,
                                    peaks_dist,
                                    peaks_promi, valleys_promi,
                                    show_data_labels = labels,
                                    find_mode = find_mode_list[n]
                                    )
            
        # 確認資料
        print(f"{data_type} 掃頻最大點：")
        print(DATA[0])
        print(Max_point)
        print(Min_point)
            
    except Exception as e:
        print("⚠️找共振點發生錯誤：", e)
        messagebox.showerror("錯誤", "找掃頻最大點時，發生錯誤！")
    
    return Max_point, Min_point

# 初期資料處理
def find_peaks_or_valleys(no, DATA, ScanErr):
    # 保留原始資料
    orig_DATA = copy.deepcopy(DATA)

    # 初始化儲存表
    rev_list        = [None]
    frq_gap         = [None]
    direction       = [None]
    prominence      = [None]
    find_mode_list  = [None]

    # 搜尋峰值或峰谷
    keywords = "HN"
    for i in range(1, len(DATA[0])):
        if keywords.lower() in DATA[0][i].lower():
            x = 1
        
        else:
            x = 0

        # 根據設定決定找 peak / valley，並記錄對應資訊
        mode = promi_finds[no][x]
        find_mode_list.append(mode)

        if mode == "peak":
            rev_list.append(False)
        elif mode == "valley":
            rev_list.append(True)
        else:
            rev_list.append(False)  # 預設不反轉（保護）

        frq_gap.append(freq_set[no][x])
        direction.append(promi_direct[no][x])
        prominence.append(float(ScanErr[no][x].get()))

    # 對資料反轉（如果找的是谷）
    for idx, val in enumerate(rev_list):
        if val:
            for i in range(1, len(DATA)):
                if DATA[i][idx] != "":
                    DATA[i][idx] = - DATA[i][idx]

    # 回傳所有資料，包含 find_mode_list
    return orig_DATA, DATA, find_mode_list, frq_gap, direction, prominence
 

# 找出有資料的起始和終止位置(資料, 資料位置)
def find_data_ranges(DATA, column_index):
    Start, End = None, None
    for i, row in enumerate(DATA[1:]):
        value = row[column_index]
        if value != "":  # 如果有資料
            if Start is None:
                Start = i  # 記錄起始位置
            End = i        # 更新終止位置
    Start += 1
    End   += 1
    return Start, End

# 計算數據最大重複值(平台)
def find_consecutive_max_length(DATA, n):
    max_len = 0
    count = 0
    prev_value = None

    for i in range(len(DATA)):
        current = DATA[i][n]
        if current == "":
            continue  # 空值不計算也不重設

        if current == prev_value:
            count += 1
        else:
            count = 1
            prev_value = current

        max_len = max(max_len, count)

    return max_len

# 自動找最佳間距
def auto_find_best_gap(method_func, DATA, n, dStart, dEnd, max_gap = 5):
    get_peaks = set()
    
    for gap in range(1, max_gap + 1):
        peaks = method_func(DATA, n, dStart, dEnd, gap)
        get_peaks = get_peaks | set(peaks)

    get_peaks = list(get_peaks)        
    get_peaks.sort()
    
    return get_peaks

# 3點找PEAK
def find_peak_3point(DATA, n, dStart, dEnd, gap):
    peaks_index = []

    for i in range(dStart, dEnd):
        if DATA[i][n] == "":
            continue

        # 往左找 gap 個有效資料點
        left = i - 1
        left_count = 0
        while left >= dStart:
            if DATA[left][n] != "":
                left_count += 1
                if left_count == gap:
                    break
            left -= 1
        if left < dStart or DATA[left][n] == "":
            continue  # 左邊不足 gap 個有效點

        # 往右找 gap 個有效資料點
        right = i + 1
        right_count = 0
        while right < dEnd:
            if DATA[right][n] != "":
                right_count += 1
                if right_count == gap:
                    break
            right += 1
        if right >= dEnd or DATA[right][n] == "":
            continue  # 右邊不足 gap 個有效點

        # 比較大小
        if DATA[i][n] > DATA[left][n] and DATA[i][n] > DATA[right][n]:
            peaks_index.append(i)

    return peaks_index

# Distance 排除靠近者，保留區間主峰
def Distance(DATA, n, peaks, valleys, prominences, dStart, dEnd, distance = 1):
    # 將 peak 座標排序（保險）
    peaks = sorted(peaks)

    # 預先建立值陣列（避免重複存取 DATA）
    peak_values = {p: DATA[p][n] for p in peaks if isinstance(DATA[p][n], (int, float))}

    # 標記哪些 peak 要保留（預設全部保留）
    keep_flags = {p: True for p in peaks}

    # 主迴圈：滑動窗口 + 比較距離
    for i in range(len(peaks)):
        if not keep_flags[peaks[i]]:
            continue  # 若已經被淘汰，就跳過

        group  = [peaks[i]]
        values = [peak_values.get(peaks[i], -float('inf'))]

        for j in range(i + 1, len(peaks)):
            if abs(peaks[j] - peaks[i]) <= distance:
                if keep_flags[peaks[j]]:
                    group.append(peaks[j])
                    values.append(peak_values.get(peaks[j], -float('inf')))
            else:
                break  # 超過範圍就不比了（因為 peaks 是排序的）

        if len(group) > 1:
            # 找出該區域最大者
            ranked = ranks(values)
            winner_index = ranked.index(1)
            winner = group[winner_index]

            for k, p in enumerate(group):
                if p != winner:
                    keep_flags[p] = False

    # 根據 keep_flags 重建 peaks / valleys / prominences
    new_peaks       = []
    new_valleys     = []
    new_prominences = []

    for i, p in enumerate(peaks):
        if keep_flags[p]:
            new_peaks.append(p)
            if i < len(valleys):
                new_valleys.append(valleys[i])
            if i < len(prominences):
                new_prominences.append(prominences[i])

    return new_peaks, new_valleys, new_prominences

# Prominence
def Prominence(DATA, n, peaks, dStart, dEnd, prominence, search_range = None, side = 'both'):
    new_peaks = []
    valleys = []
    prominences = []

    for peak_index in peaks:
        peak_value = DATA[peak_index][n] # 峰值
        left_valley_value  = peak_value  # 左波谷值
        right_valley_value = peak_value  # 右波谷值
        left_valley_index  = peak_index  # 左波谷值序
        right_valley_index = peak_index  # 右波谷值序
        
        # 搜尋左谷底
        if side in ('left', 'both'):
            left = peak_index
            step_left = 0
            while left > dStart and (search_range is None or step_left < search_range):
                left -= 1

                if isinstance(DATA[left][n], (int, float)):
                    step_left += 1
                    
                    if DATA[left][n] > peak_value:
                        break
                    if DATA[left][n] < left_valley_value:
                        left_valley_index = left
                        left_valley_value = DATA[left][n]
            
        # 搜尋右谷底
        if side in ('right', 'both'):
            right = peak_index
            step_right = 0
            while right < dEnd and (search_range is None or step_right < search_range):
                right += 1

                if isinstance(DATA[right][n], (int, float)):
                    step_right += 1
                    
                    if DATA[right][n] > peak_value:
                        break
                    if DATA[right][n] < right_valley_value:
                        right_valley_index = right
                        right_valley_value = DATA[right][n]
        
        # 計算 prominence 值
        if side == 'left':
            prominence_value = peak_value - left_valley_value
            chosen_valley = left_valley_index
            
        elif side == 'right':
            prominence_value = peak_value - right_valley_value
            chosen_valley = right_valley_index
            
        else:  # both
            if left_valley_value >= right_valley_value:
                chosen_valley = left_valley_index
                
            elif left_valley_value < right_valley_value:
                chosen_valley = right_valley_index
                
            prominence_value = peak_value - max(left_valley_value, right_valley_value)
            
        if prominence_value >= prominence:
            new_peaks.append(peak_index)
            valleys.append(chosen_valley)
            prominences.append(prominence_value)

    return new_peaks, valleys, prominences

# 電流排名篩選
def ranks(Current):
    # 把原始資料配上索引
    indexed = list(enumerate(Current))

    # 根據數值從大到小排序
    sorted_indexed = sorted(indexed, key = lambda x: x[1], reverse = True)

    # 建立結果 list，初始全 0
    ranks = [0] * len(Current)

    # 記錄每個原始位置的排名
    for rank, (i, val) in enumerate(sorted_indexed, start = 1):
        ranks[i] = rank
    
    return ranks

# 反轉峰值和峰谷
def rev_peaks_and_valleys(peaks, valleys):    
    temp    = copy.deepcopy(peaks)
    peaks   = copy.deepcopy(valleys)
    valleys = copy.deepcopy(temp)
    
    return peaks, valleys

# Find Peaks analysis ---------------------------------------------------------------------------------------
# 只顯示最終共振點結果
def plot_final_peak(data_type, all_data_index, DATA, channel_index, dStart, dEnd, 
                    peaks_final, valleys_final, show_data_labels = None, find_mode = "peak"):
    # 標記樣式依據主目標 find_mode 來決定（模仿 plot_find_max_steps 方式）
    if find_mode == "peak":
        p_color, p_marker, p_s = "gold", "*", 100
        v_color, v_marker, v_s = "gray", "x", 65
        tag = "Peaks"
    else:  # valley
        p_color, p_marker, p_s = "green", "^", 65
        v_color, v_marker, v_s = "gold",  "*", 100
        tag = "Valleys"

    # 擷取資料
    x = [row[0] for row in DATA[dStart:dEnd] if row[channel_index] not in [None, ""]]
    y = [row[channel_index] for row in DATA[dStart:dEnd] if row[channel_index] not in [None, ""]]

    fig, ax = plt.subplots(figsize=(14, 10))
    ax.plot(x, y, label='Original Curve', color='lightgray')

    # 畫出標記：星號 (*) 是主目標
    ax.scatter([DATA[p][0] for p in peaks_final], [DATA[p][channel_index] for p in peaks_final],
               color = p_color, marker = p_marker, s = p_s, label = 'Peaks' if find_mode == "peak" else 'Peaks (non-target)')

    ax.scatter([DATA[v][0] for v in valleys_final], [DATA[v][channel_index] for v in valleys_final],
               color = v_color, marker = v_marker, s = v_s, label = 'Valleys' if find_mode == "valley" else 'Valleys (non-target)')

    # 標記數值與垂直線
    y_min, y_max = ax.get_ylim()
    peak_y_offset   = (y_max - y_min) * 0.02
    valley_y_offset = (y_max - y_min) * 0.03
    label_ys = []

    for i in range(min(len(peaks_final), len(valleys_final))):
        peak = peaks_final[i]
        valley = valleys_final[i]
        peak_y   = DATA[peak][channel_index]   + peak_y_offset
        valley_y = DATA[valley][channel_index] - valley_y_offset

        ax.text(DATA[peak][0], peak_y,
                f'{smart_format(DATA[peak][0])}, {smart_format(DATA[peak][channel_index])}',
                color = 'black', fontsize = 9, ha = 'center', clip_on = True)
        ax.text(DATA[valley][0], valley_y,
                f'{smart_format(DATA[valley][0])}, {smart_format(DATA[valley][channel_index])}',
                color = 'black', fontsize = 9, ha = 'center', clip_on = True)

        label_ys.extend([peak_y, valley_y])

    if label_ys:
        ax.set_ylim(
            min(y_min, min(label_ys) - valley_y_offset),
            max(y_max, max(label_ys) + peak_y_offset)
        )

    # 標題與座標軸
    ax.set_title(f'Final Result - Target: {tag} (*)')
    fig.suptitle(DATA[0][channel_index], fontsize=12, fontweight='bold')
    ax.set_xlabel('Frequency [Hz]')

    if data_type == "ENA":
        ax.set_ylabel(f"{data_titles_en[all_data_index]} [{data_units[all_data_index]}]")

    ax.grid(True)
    ax.legend()
    plt.tight_layout()
    plt.get_current_fig_manager().window.showMaximized()
    plt.show()



# 3點法 + Distance + Prominence 圖表
def plot_find_max_steps(
                        data_type, all_data_index, DATA, channel_index, dStart, dEnd,
                        peaks_3, 
                        peaks_distance,
                        peaks_promi, valleys_promi,
                        show_data_labels = None, find_mode = "peak"
                        ):
    
    # 標記峰值或峰谷
    if find_mode == "peak":
        tag = "Peaks"
        p_color, p_maker, p_s = "gold", "*", 60
        v_color, v_maker, v_s = "gray", "x", 30

    else:
        tag = "Valleys"
        p_color, p_maker, p_s = "green", "^", 30
        v_color, v_maker, v_s = "gold",  "*", 60

    x = [row[0] for row in DATA[dStart:dEnd] if row[channel_index] not in [None, ""]]
    y = [row[channel_index] for row in DATA[dStart:dEnd] if row[channel_index] not in [None, ""]]

    fig, axes = plt.subplots(3, figsize = (14, 10), sharex = True, sharey = True)
    fig.suptitle(DATA[0][channel_index], fontsize = 12, fontweight = 'bold')

    # Step 1: 3點與5點法找出的 peaks
    axes[0].plot(x, y, label = 'Original Curve', color = 'lightgray')
    axes[0].scatter([DATA[p][0] for p in peaks_3], [DATA[p][channel_index] for p in peaks_3],
                    color = 'red', marker = 'o', label = f'{tag} (by 3-point)')
    axes[0].set_title(f'Step 1: 3-Point Filtered {tag}', fontweight = 'bold')
    axes[0].legend()

    # Step 2: Distance 篩選後保留的 peaks
    axes[1].plot(x, y, label = 'Original Curve', color = 'lightgray')
    axes[1].scatter([DATA[p][0] for p in peaks_distance], [DATA[p][channel_index] for p in peaks_distance],
                    color = 'blue', marker = 's', label = f'{tag} (by Distance)')
    axes[1].set_title(f'Step 2: Distance-Filtered {tag}', fontweight = 'bold')
    axes[1].legend()

    # Step 3: Prominence 結果
    axes[2].plot(x, y, label = 'Original Curve', color = 'lightgray')
    axes[2].scatter([DATA[p][0] for p in peaks_promi], [DATA[p][channel_index] for p in peaks_promi],
                    color = p_color, marker = p_maker, s = p_s, label = 'Peaks (by Prominence)')
    axes[2].scatter([DATA[v][0] for v in valleys_promi], [DATA[v][channel_index] for v in valleys_promi],
                    color = v_color, marker = v_maker, s = v_s, label = 'Valleys (by Prominence)')
    axes[2].set_title('Step 3: Prominence-Filter Peaks & Valleys', fontweight = 'bold')
    axes[2].legend()
    
    if show_data_labels:
        # 取得目前 y 軸範圍（原始資料範圍）
        y_min, y_max = axes[2].get_ylim()
        peak_y_offset   = (y_max - y_min) * 0.04
        valley_y_offset = (y_max - y_min) * 0.10
    
        # 放標籤，順便記錄標籤最高和最低 y 座標
        label_ys = []
        for i, peak in enumerate(peaks_promi):
            valley = valleys_promi[i]

            peak_y = DATA[peak][channel_index] + peak_y_offset
            axes[2].text(DATA[peak][0], peak_y,
                         f'{smart_format(DATA[peak][0])}, {smart_format(DATA[peak][channel_index])}',
                         color = 'black', fontsize = 9, ha = 'center', clip_on = True)
            label_ys.append(peak_y)
        
            valley_y = DATA[valley][channel_index] - valley_y_offset
            axes[2].text(DATA[valley][0], valley_y,
                         f'{smart_format(DATA[valley][0])}, {smart_format(DATA[valley][channel_index])}',
                         color = 'black', fontsize = 9, ha = 'center', clip_on = True)
            label_ys.append(valley_y)
        
        if label_ys:
            max_label_y = max(label_ys)
            min_label_y = min(label_ys)
            new_y_min   = min(y_min, min_label_y - valley_y_offset * 1.5)
            new_y_max   = max(y_max, max_label_y + peak_y_offset * 2.5)
            axes[2].set_ylim(new_y_min, new_y_max)
        
        else:
            axes[2].set_ylim(y_min, y_max)

    # Y軸
    for i in range(3):
        ax = axes[i]
        if all_data_index < 6:
            ax.set_ylabel(f"{data_titles_en[all_data_index]} [{data_units[all_data_index]}]")
        else:
            ax.set_ylabel(f"{data_titles_en[all_data_index]}")
        
        if i == 1:
            ax.set_xlabel('Frequency [Hz]') 
        
        ax.grid(True)

    plt.tight_layout()
    plt.get_current_fig_manager().window.showMaximized()
    plt.show()

def smart_format(val):
    return f'{val:.2f}' if abs(val) >= 1e-3 else f'{val:.2e}'

# 繪圖模組------------------------------------------------------------------------------------------

# Excel 圖表繪製
def Drawing(title, unit, DATA, Lcolors, Ltypes, Lwidth, ws):
    # 圖表標題設定
    def set_chart_title_size(chart, size = 1400):
        paraprops = ParagraphProperties()
        paraprops.defRPr = CharacterProperties(sz = size)

        for para in chart.title.tx.rich.paragraphs:
            para.pPr = paraprops
    
    # 圖表存檔位置
    def Drawing_adress(n):
        adress = []
        n += 1
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            adress.append(chr(remainder + ord('A')))
        return ''.join(reversed(adress))
    
    if DATA != []:
        # XY散佈圖
        chart = ScatterChart()
        chart.title = title
        set_chart_title_size(chart, size = 1400)
        chart.style = 13
        chart.y_axis.title = unit
        chart.x_axis.title = DATA[0][0]
        
        # X軸
        xMin = lambda x: (x // 1000) * 1000 if (x % 1000) > 0 else x # X軸最小值調整
        xMax = lambda x: ((x // 1000) + 1) * 1000 if (x % 1000) > 0 else x # X軸最大值調整
        xvalues = Reference(ws, min_col = 1, min_row = 2, max_row = len(DATA))
        chart.x_axis.scaling.min = xMin(DATA[1][0])  # Minimum value for x-axis
        chart.x_axis.scaling.max = xMax(DATA[-1][0]) # Maximum value for x-axis
        
        # X軸刻度間距
        chart.x_axis.majorUnit = 5000
        
        # Y軸
        for y in range(2, len(DATA[0]) + 1):
            yvalues = Reference(ws, min_col = y, min_row = 1, max_row = len(DATA))
            series = Series(yvalues, xvalues, title_from_data = True)
            line_properties = LineProperties(w = Lwidth, solidFill = Lcolors[y - 2], prstDash = Ltypes[y - 2])
            series.graphicalProperties.line = line_properties
            # 資料存檔
            chart.series.append(series)
        
        # 設定線的樣式平滑曲線
        for x in range(len(DATA[0]) - 1):
            chart.series[x].smooth = True
        
        # 圖表儲存位置
        adress = Drawing_adress(len(DATA[0])) + str("1")
        
        # 圖表大小
        chart.height = 7.5  # 設置高度
        chart.width  = 17   # 設置寬度
        return chart, adress

# 隱藏的空格與空值補線
def set_excel_chart_options(file_path):
    excel = win32.DispatchEx("Excel.Application")  # 使用 DispatchEx 確保背景執行
    excel.Visible = False                          # 設定為不可見
    excel.DisplayAlerts = False                    # 關閉提示

    try:
        # 開啟文件
        workbook = excel.Workbooks.Open(file_path)
        
        # 取得所有分頁
        for sheet in workbook.Sheets:
            # 迭代該分頁中的所有圖表
            for chart_object in sheet.ChartObjects():
                chart = chart_object.Chart
                # 設定「連接資料點的線」
                chart.DisplayBlanksAs = 3  # 3 表示連接空白資料點 (xlInterpolated)

        # 儲存
        workbook.Save()

    finally:
        # 確保清理資源並關閉 Excel
        workbook.Close(SaveChanges = True)
        excel.Quit()

# Excel存檔----------------------------------------------------------------------------------------

# 建立Excel
def Excel_file(f_path, section_data_types, 
               all_data, Lcolors, Ltypes, Lwidth, 
               var_max, check_peak_drw_var, show_data_labels_var, 
               analysis_ck_vars, analysis_cbbs_vars):
    # 創建一個新的 Excel 工作簿
    wb = Workbook()  

    # 使用勾選選單，選擇要呈現的檔案
    any_data = False  # 用來檢查是否有勾選項被選中

    # 使用勾選選單，選擇要呈現的檔案
    for i in section_data_types:
        DATA = all_data[i]
      
        # 若資料為空，不生成Excel分頁
        if DATA != []:
            any_data = True  # 有勾選項被選中
            # 新增Excel分頁
            if section_data_types.index(i) == 0:
                ws = wb.active
                ws.title = data_types[i][0]

            else:
                ws = wb.create_sheet(title = data_types[i][0])
            
            # 儲存資料
            for row in DATA:
                ws.append(row)
                    
            # 執行圖表繪製
            chart, adress = Drawing(data_titles[i], data_units[i], DATA, Lcolors, Ltypes, Lwidth, ws)
            ws.add_chart(chart, adress)

            # 找峰值峰谷
            if var_max.get() and analysis_ck_vars[i].get():
                Max_point, Min_point = Find_peaks_max("ENA", i, DATA, analysis_cbbs_vars, 
                                                      show_analysis = check_peak_drw_var.get(),
                                                      labels = show_data_labels_var.get())
                ws.append(Max_point)
                ws.append(Min_point)
            
    return wb, any_data

# 檢查Excel是否成功生成
def Save_Excel_file(wb, any_data, f_path):
    try:
        # 儲存檔案
        if any_data:
            save_path = f'{f_path}/ENA_Output.xlsx'
            wb.save(save_path)
            
            # 成功訊息
            messagebox.showinfo("成功", "成功合併檔案：ENA_Output.xlsx")
        else:
            print("⚠️檢查 1：讀取數據方式，是否選擇正確！")
            print("⚠️檢查 2：ENA數據資料，至少選擇 1 項！")
            messagebox.showerror("錯誤", "合併檔案失敗！！")
            
    except Exception as e:
        print("⚠️檢查 1：讀確認ENA_Output.xlsx，檔案是否有開啟！")
        messagebox.showerror("錯誤", f"儲存檔案時，發生錯誤：{str(e)}")
        
    return save_path

# GUI介面------------------------------------------------------------------------------------------

class ENA_App:
    def __init__(self, root):
        self.root = root
        self.root.title("ENA merge files")
        self.window_width  = 260
        self.window_height = 425
        self.root.geometry(f"{self.window_width}x{self.window_height}")  # 設置窗口大小
        self.root.minsize(self.window_width, self.window_height) # 限制視窗大小
        self.root.maxsize(self.window_width, self.window_height) # 限制視窗大小

        # 瀏覽資料夾框架 -----------------------------------------------------------------------------------------
        self.f_path_frame = ttk.LabelFrame(self.root, text = "檔案路徑：", relief = "groove", borderwidth = 2)
        self.f_path_frame.place(x = 15, y = 10, width = 230, height = 60)
        self.f_path_frame.config(style = "Dashed.TFrame")
        
        # 創建Listbox
        self.listbox = tk.Listbox(self.f_path_frame, width = 23, height = 1)
        self.listbox.grid(row = 0, column = 0, padx = 5, pady = 5)

        # 創建瀏覽資料夾按鈕
        self.browse_button = ttk.Button(self.f_path_frame, text = "瀏覽...", command = self.browse_folder, width = 6.5)
        self.browse_button.grid(row = 0, column = 1, padx = 0, pady = 5)

        # 儲存最新選擇的資料夾路徑
        self.latest_folder = os.getcwd()
        self.update_listbox()


        # 讀取數據方式選項 ---------------------------------------------------------------------------------------
        self.radio_frame_file = ttk.LabelFrame(self.root, text = "讀取數據方式：", relief = "groove", borderwidth = 2)
        self.radio_frame_file.place(x = 15, y = 80, width = 230, height = 55)
        self.radio_frame_file.config(style = "Dashed.TFrame")
    
        self.radio_var_file = tk.StringVar()
        self.radio_var_file.set('dir_mode')  # 預設選項

        self.radio_dir = ttk.Radiobutton(self.radio_frame_file, 
                                         text     = '路徑/資料夾/檔案', 
                                         variable = self.radio_var_file, 
                                         value    = 'dir_mode',
                                         command  = self.update_checkbutton_state)
        self.radio_file = ttk.Radiobutton(self.radio_frame_file, 
                                          text     = '路徑/檔案', 
                                          variable = self.radio_var_file, 
                                          value    = 'file_mode',
                                          command  = self.update_checkbutton_state)

        self.radio_dir.grid (row = 0, column = 0, padx = 5, pady = 5)
        self.radio_file.grid(row = 0, column = 1, padx = 5, pady = 5)
        
        # ENA數據資料選項：第1行 ------------------------------------------------------------------------------
        self.frame1 = ttk.LabelFrame(self.root, text = "ENA數據資料：", relief = "groove", borderwidth = 2)
        self.frame1.place(x = 15, y = 145, width = 230, height = 115)
        self.frame1.config(style = "Dashed.TFrame")

        self.var_z = tk.BooleanVar(value = True)
        self.var_o = tk.BooleanVar(value = True)
        self.var_c = tk.BooleanVar(value = True)

        self.check_z = ttk.Checkbutton(self.frame1, text = "阻抗(Z)", variable = self.var_z, command = self.anlysis_check)
        self.check_o = ttk.Checkbutton(self.frame1, text = "相位(O)", variable = self.var_o, command = self.anlysis_check)
        self.check_c = ttk.Checkbutton(self.frame1, text = "電容(C)", variable = self.var_c, command = self.anlysis_check)

        self.check_z.grid(row = 0, column = 0, padx = 5, pady = 5)
        self.check_o.grid(row = 0, column = 1, padx = 5, pady = 5)
        self.check_c.grid(row = 0, column = 2, padx = 5, pady = 5)

        # ENA數據資料選項：第2行
        self.var_x = tk.BooleanVar()
        self.var_r = tk.BooleanVar()
        self.var_l = tk.BooleanVar()
        self.var_q = tk.BooleanVar()

        self.check_x = ttk.Checkbutton(self.frame1, text = "感抗(X)", variable = self.var_x, command = self.anlysis_check)
        self.check_r = ttk.Checkbutton(self.frame1, text = "電阻(R)", variable = self.var_r, command = self.anlysis_check)
        self.check_l = ttk.Checkbutton(self.frame1, text = "電感(L)", variable = self.var_l, command = self.anlysis_check)
        self.check_q = ttk.Checkbutton(self.frame1, text = "品質因子(Q)", variable = self.var_q, command = self.anlysis_check)

        self.check_x.grid(row = 1, column = 0, padx = 5, pady = 5)
        self.check_r.grid(row = 1, column = 1, padx = 5, pady = 5)
        self.check_l.grid(row = 1, column = 2, padx = 5, pady = 5)
        self.check_q.grid(row = 2, column = 0, padx = 5, pady = 5, columnspan = 2, sticky = "w")
        
        self.button_var = [ self.var_z, self.var_o, self.var_c, 
                            self.var_x, self.var_r, self.var_l, self.var_q ]
        
        self.button_check = [ self.check_z, self.check_o, self.check_c, 
                              self.check_x, self.check_r, self.check_l, self.check_q ]
        
        self.analysis_button = ttk.Button(self.frame1, text = "峰值分析", command = self.analysis_setting, width = 8)
        self.analysis_button.grid(row = 2, column = 2, padx = 0, pady = 5)
        
        # 掃頻設定 ---------------------------------------------------------------------------------------------
        self.var_max                 = tk.BooleanVar(value = False)
        self.check_peak_drw_var      = tk.BooleanVar()
        self.show_data_labels_var    = tk.BooleanVar()
                
        # 初始化勾選框變數
        self.analysis_ck_vars   = []
        self.analysis_ckbtn     = []
        self.analysis_ck_vars   = []
        self.analysis_cbbs_vars = []
        self.analysis_cbbs1     = []
        self.analysis_cbbs2     = []
  
        # 下拉選單 ----------------------------------------------------------------------------------------
        self.frame2 = ttk.LabelFrame(self.root, text = "請選擇數量：", relief = "groove", borderwidth = 2)
        self.frame2.place(x = 15, y = 270, width = 125, height = 110)
        self.frame2.config(style = "Dashed.TFrame")

        self.label1 = ttk.Label(self.frame2, text = "刀把數量")
        self.label1.grid(row = 0, column = 0, padx = 5, pady = 5)

        self.combo_var1 = tk.StringVar()
        self.combo = ttk.Combobox(self.frame2, 
                                  textvariable = self.combo_var1, 
                                  values = [str(i) for i in range(1, 1000)], 
                                  width  = 3,
                                  state  = 'readonly')
        self.combo.grid(row = 0, column = 1, padx = 5, pady = 5)
        self.combo.current(0)  # 預設選擇第一個選項

        self.label2 = ttk.Label(self.frame2, text = "刀具數量")
        self.label2.grid(row = 1, column = 0, padx = 5, pady = 5)

        self.combo_var2 = tk.StringVar()
        self.combo2 = ttk.Combobox(self.frame2, 
                                   textvariable = self.combo_var2, 
                                   values = [str(i) for i in range(0, int(len(linetypes) + 1))], 
                                   width = 3,
                                   state = 'readonly')
        self.combo2.grid(row = 1, column = 1, padx = 5, pady = 5)
        self.combo2.current(0)  # 預設選擇第一個選項

        # 進階選單 ----------------------------------------------------------------------------------------
        self.frame3 = ttk.LabelFrame(self.root, text = "進階選項：", relief = "groove", borderwidth = 2)
        self.frame3.place(x = 145, y = 270, width = 100, height = 110)
        self.frame3.config(style = "Dashed.TFrame")

        self.var_color = tk.BooleanVar()
        self.var_linetype = tk.BooleanVar()
        
        self.combo_var3 = tk.StringVar() # 線寬選單
        
        # 線寬清單
        combo3_values = []  # 線寬空集合
        Lw = 0.5            # 起始
        for i in range(10):
            combo3_values.append("線寬:" + str(Lw) + " pt")
            Lw += 0.5 # 線寬間距
            
        self.combo3 = ttk.Combobox(self.frame3, 
                                   textvariable = self.combo_var3, 
                                   values = combo3_values, 
                                   width  = 9,
                                   state  = 'readonly')
        self.combo3.current(1)

        self.check_color = ttk.Checkbutton(self.frame3, text = "單色線條", variable = self.var_color)
        self.check_linetype = ttk.Checkbutton(self.frame3, text = "實線線條", variable = self.var_linetype)

        self.combo3.grid        (row = 0, column = 1, padx = 5, pady = 5, sticky = "w")
        self.check_color.grid   (row = 1, column = 0, padx = 5, pady = 5, columnspan = 2, sticky = "w")
        self.check_linetype.grid(row = 2, column = 0, padx = 5, pady = 5, columnspan = 2, sticky = "w")
        
        
        # 執行/離開按鈕框架 ----------------------------------------------------------------------------------------
        self.button_frame = ttk.Frame(self.root)
        self.button_frame.place(x = 40, y = 385)

        self.execute_button = ttk.Button(self.button_frame, text = "執行", command = self.execute, width = 10)
        self.execute_button.grid(row = 0, column = 0, padx = 5, pady = 5)
        
        self.close_button = ttk.Button(self.button_frame, text = "離開", command = self.close_action, width = 10)
        self.close_button.grid(row = 0, column = 1, padx = 5, pady = 5)

    # 瀏覽資料夾
    def browse_folder(self):
        self.folder_path = filedialog.askdirectory()
        if self.folder_path:
            self.latest_folder = self.folder_path
            self.update_listbox()
            self.update_combobox()
               
    # 清除Listbox的內容，並插入最新選擇的資料夾路徑    
    def update_listbox(self):
        self.listbox.delete(0, tk.END)
        self.listbox.insert(tk.END, self.latest_folder)      
    
    # 分析設定
    def analysis_setting(self):
        if hasattr(self, 'settings_window') and self.settings_window.winfo_exists():
            self.settings_window.deiconify()
            self.settings_window.lift()
            return
    
        self.settings_window = tk.Toplevel(self.root)
        self.settings_window.title("Find Peaks")
        self.settings_width  = 366
        self.settings_height = 260
        self.settings_window.geometry(f"{self.settings_width}x{self.settings_height}")
        self.settings_window.resizable(False, False)
        self.settings_window.protocol("WM_DELETE_WINDOW", self.hide_settings_window)
        
        # 設定 -----------------------------------------------------------------------------
        self.show_drw_frame = ttk.LabelFrame(self.settings_window, text = "峰值分析設定：", relief = "groove", borderwidth = 2)
        self.show_drw_frame.place(x = 15, y = 10, width = 336, height = 60)
        self.show_drw_frame.config(style = "Dashed.TFrame")
    
        self.check_max = ttk.Checkbutton(self.show_drw_frame,
                                         text = "峰值峰谷分析",
                                         variable = self.var_max,
                                         command  = self.anlysis_check)
        self.check_max.grid(row = 0, column = 0, padx = 5, pady = 5, sticky = "w")
    
        self.check_peak_drw = ttk.Checkbutton(self.show_drw_frame,
                                              text = "顯示分析圖表",
                                              variable = self.check_peak_drw_var,
                                              command  = self.show_labels)
        self.check_peak_drw.grid(row = 0, column = 1, padx = 5, pady = 5, sticky = "w")
    
        self.show_data_labels = ttk.Checkbutton(self.show_drw_frame,
                                                text = "顯示數據標籤",
                                                variable = self.show_data_labels_var,
                                                state    = "disabled")
        self.show_data_labels.grid(row = 0, column = 2, padx = 5, pady = 5, sticky = "w")
    
        # 表格 --------------------------------------------------------------------------------
        data = [
                ("阻抗(Z)差值", 
                 [str(i) for i in range(10, 301, 5)],         [str(i) for i in range(500, 10001, 500)],    "Ω"),
                ("相位(O)差值", 
                 [str(i) for i in range(1, 51)],              [str(i) for i in range(5, 101, 5)],          "°"),
                ("電容(C)差值", 
                 [f"{i * 10**-8:.1e}" for i in range(1, 31)], [f"{i * 10**-8:.1e}" for i in range(1, 31)], "nF"),
                ("感抗(X)差值", 
                 [str(i) for i in range(10, 301, 5)],         [str(i) for i in range(500, 10001, 500)],    "Ω"),
                ("電阻(R)差值", 
                 [str(i) for i in range(10, 301, 5)],         [str(i) for i in range(500, 10001, 500)],    "Ω"),
                ("電感(L)差值", 
                 [str(i) for i in range(1, 21)],              [str(i) for i in range(1, 21)],              "nH"),
                ("品因(Q)差值", 
                 [str(i) for i in range(1, 51)],              [str(i) for i in range(10, 51, 5)],          "")
                ]
    
        table_frame = tk.Frame(self.settings_window)
        table_frame.place(x = 15, y = 80, width = 600, height = 300)
        
        headers = ["勾選", "項目", "刀把", "變幅桿", "單位"]
        for col, title in enumerate(headers):
            label = tk.Label(table_frame, text = title,
                             borderwidth = 1, relief = "solid",
                             bg = "#cccccc", font = ("Arial", 10, "bold"),
                             padx = 8, pady = 5)
            label.grid(row = 0, column = col, sticky = "nsew")
    
        combo_width = 7
    
        for i, (title, options1, options2, unit) in enumerate(data, start = 1):
            # 勾選框
            var = tk.BooleanVar(value = False)
            frame_chk = tk.Frame(table_frame, borderwidth = 1, relief = "solid")
            frame_chk.grid(row = i, column = 0, sticky = "nsew")
            chk = tk.Checkbutton(frame_chk, variable = var)
            chk.pack(expand = True)
            self.analysis_ck_vars.append(var)
            self.analysis_ckbtn.append(chk)
    
            # 項目
            lbl_title = tk.Label(table_frame, text = title,
                                 borderwidth = 1, relief = "solid",
                                 padx = 8, pady = 5)
            lbl_title.grid(row = i, column = 1, sticky = "nsew")
    
            # 選擇1
            frame_combo1 = tk.Frame(table_frame, borderwidth = 1, relief = "solid")
            frame_combo1.grid(row = i, column = 2, sticky = "nsew")
            
            combo_var1 = tk.StringVar(value = options1[0])
            combo1 = ttk.Combobox(frame_combo1, values = options1, textvariable = combo_var1, width = combo_width, state = "readonly")
            combo1.pack(padx = 2, pady = 2)
            self.analysis_cbbs1.append(combo1)
    
            # 選擇2
            frame_combo2 = tk.Frame(table_frame, borderwidth = 1, relief = "solid")
            frame_combo2.grid(row = i, column = 3, sticky = "nsew")
            
            combo_var2 = tk.StringVar(value = options2[0])
            combo2 = ttk.Combobox(frame_combo2, values = options2, textvariable = combo_var2, width = combo_width, state = "readonly")
            combo2.pack(padx = 2, pady = 2)
            self.analysis_cbbs2.append(combo2)
            
            self.analysis_cbbs_vars.append((combo_var1, combo_var2))
    
            # 單位
            lbl_unit = tk.Label(table_frame, text = unit,
                                borderwidth = 1, relief = "solid",
                                padx = 8, pady = 5)
            lbl_unit.grid(row = i, column = 4, sticky = "nsew")
            
            # 隱藏電容、電感設定
            if i == 3 or i == 6:
                frame_chk.grid_remove()
                lbl_title.grid_remove()
                frame_combo1.grid_remove()
                frame_combo2.grid_remove()
                lbl_unit.grid_remove()
            
    
        for col in range(5):
            table_frame.grid_columnconfigure(col, weight = 0)
        
        self.anlysis_check()
        
    # 分析啟動勾選
    def anlysis_check(self):
        # 刷新峰值分析可用狀態
        try:
            if self.var_max.get():
                self.check_peak_drw.config(state = "normal")
                
                for i in range(len(self.button_var)):
                    if self.button_var[i].get() and i != 2 and i != 5:
                        self.analysis_ck_vars[i].set(True)
                        self.analysis_ckbtn[i].config(state = "normal")
                        self.analysis_cbbs1[i].config(state = "readonly")
                        self.analysis_cbbs2[i].config(state = "readonly")
            
            else:
                self.check_peak_drw_var.set(False)
                self.show_data_labels_var.set(False)
                self.check_peak_drw.config(state = "disabled")
                self.show_data_labels.config(state = "disabled")
                
                for i in range(len(self.button_var)):
                    self.analysis_ck_vars[i].set(False)
                    self.analysis_ckbtn[i].config(state = "disabled")
                    self.analysis_cbbs1[i].config(state = "disabled")
                    self.analysis_cbbs2[i].config(state = "disabled")
        
            # 主頁面取消勾選時，峰值設定取消
            for i in range(len(self.button_var)):
                if not self.button_var[i].get():
                    self.analysis_ck_vars[i].set(False)
                    self.analysis_ckbtn[i].config(state = "disabled")
                    self.analysis_cbbs1[i].config(state = "disabled")
                    self.analysis_cbbs2[i].config(state = "disabled")
        
        except Exception as e:
            e
    
    # 數據標籤
    def show_labels(self):
        if self.check_peak_drw_var.get():
            self.show_data_labels.config(state = "normal")
            
        else:
            self.show_data_labels_var.set(False)
            self.show_data_labels.config(state = "disabled")
            
    # 隱藏彈出窗口
    def hide_settings_window(self):
        self.settings_window.withdraw()
        
    # 選擇資料合併方式
    def update_checkbutton_state(self):
        self.mode = self.radio_var_file.get() # 取得檔案處理模式
        if self.mode == "dir_mode":  # 當選擇資料夾合併
            for i in range(len(self.button_var)):
                if i < 3:
                    n = 1
                else:
                    n = 0
                    
                # 確保勾選框被選中
                self.button_var[i].set(n)
                
                # 確保勾選框可以用
                self.button_check[i].config(state = tk.NORMAL)
                        
        elif self.mode == "file_mode": # 當選擇檔案合併
            # 取得檔案清單
            files_list = get_files_in_directory(self.folder_path, self.mode)
            #print(files_list)
            
            # 檢測ENA量測的資料類型
            for i in range(len(files_list)):
                if files_list[i] == []:
                    self.button_var[i].set(0)  # 取消勾選框
                    self.button_check[i].config(state = tk.DISABLED)  # 使勾選框不可用            

        # 刷新刀把數量
        self.update_combobox()
    
    # 自動取代刀把數量
    def update_combobox(self):  
        self.mode = self.radio_var_file.get() # 取得檔案處理模式
        if self.mode == "dir_mode":
            # 取得路徑下的資料夾清單
            subfolders = get_files_in_directory(self.folder_path, self.mode)
            
            # 取得資料夾數量
            self.file_count = int(len(subfolders))
            
            # 創建一個包含從 1 到 file_count 的列表
            values = list(range(1, self.file_count + 1))
            
        elif self.mode == "file_mode":
            # 移除空的List，並取得檔案清單
            files_list = list(filter(lambda x: x, get_files_in_directory(self.folder_path, self.mode)))
            n = files_list[0]
            
            # 取得分類後的數量
            self.file_count = int(len(n))
                
            # 創建一個包含從 1 到 self.file_count 的列表
            values = list(range(1, self.file_count + 1))
        
        # 更新 Combobox 的值
        self.combo['values'] = values
        if values:
            self.combo.set(values[-1])
    
    # 執行
    def execute(self):
        self.mode = self.radio_var_file.get() # 取得檔案處理模式
        section_data_types = [] # 量測資料數據類型
        titles = []             # 量測資料數據標題
        units  = []             # 量測資料數據單位
        
        for x in range(len(self.button_var)):
            if self.button_var[x].get():
                section_data_types.append(x)
                titles.append(data_titles[x])
                units.append(data_units[x])  
                
        # 刀把數量
        USTH_number = int(self.combo_var1.get())
        
        # 刀具數量
        tool_number = int(self.combo_var2.get())
        
        Ncolors = [] # 增量預設顏色
        Lcolors = [] # 增量顏色
        Ltypes  = [] # 增量線型
        
        # 增量預設顏色
        if USTH_number > len(colors):
            Ncolors = (USTH_number // len(colors)) * colors
            for i1 in range(USTH_number % len(colors)):
               Ncolors.append(colors[i1])
        else:
            Ncolors = colors
        
        for i in range(USTH_number):
            if tool_number <= 1:
                tool_number = 1
                n = 1
            else:
                n = tool_number
            
            # 圖表顏色、線型增量
            for j in range(n):
                Lcolors.append(Ncolors[i])
                Ltypes.append(linetypes[j])
                
        # 圖表統一顏色
        if self.var_color.get():
            for i in range(len(Lcolors)):
                Lcolors[i] = Ncolors[0]
                
        # 圖表統一線型
        if self.var_linetype.get():
            for j in range(len(Ltypes)):
                Ltypes[j] = linetypes[0]
        
        # 線寬
        Lwidth = float(self.combo_var3.get()[3:6])
        Lwidth = int(Lwidth * 12700.2)
        
        try:        
            # 避免使用小於檔案數量的參數，繪圖會出錯
            if USTH_number * tool_number >= self.file_count:
                # 處理檔案
                all_data, data_filled = ENA(self.latest_folder, self.mode)
                wb, any_data = Excel_file(self.latest_folder, section_data_types, all_data, 
                                          Lcolors, Ltypes, Lwidth, 
                                          self.var_max, self.check_peak_drw_var, self.show_data_labels_var,
                                          self.analysis_ck_vars, self.analysis_cbbs_vars)
                
                if self.mode == "file_mode": # 當選擇檔案合併
                # 檔案分類至資料夾
                    response = messagebox.askyesno("選擇", "您要分類數據至資料夾嗎？")
                    if response:
                        csv_files = get_files_in_directory(self.latest_folder, self.mode)
                        Organize_files_to_folde(self.latest_folder, csv_files)
        
                save_path = Save_Excel_file(wb, any_data, self.latest_folder)
                
                if data_filled:
                    set_excel_chart_options(save_path)
                
            else:
                print("⚠️警告：檔案數量 ≠ 刀把數量 x 刀具數量")
                messagebox.showerror("錯誤", "檔案數量不匹配，合併檔案失敗！！")
        
        except Exception as e:
            print(f"⚠執行發生錯誤錯誤：{e}")
            messagebox.showerror("錯誤", "未選擇任何檔案！")
            
    # 離開        
    def close_action(self):
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = ENA_App(root)
    root.mainloop()