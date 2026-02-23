# -*- coding: utf-8 -*-
"""
Design for "88598 AZ EB" and "PICO TC-08"

@author: Hsiao, Yu-Chieh
"""
import os
import csv

import openpyxl
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.drawing.text import ParagraphProperties, CharacterProperties
from openpyxl.drawing.line import LineProperties

from datetime import datetime, timedelta
from tkcalendar import DateEntry

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# 主程式模組---------------------------------------------------------------------------------------

# 11種線型
linetypes = ["solid", "sysDash", "sysDashDot", "sysDashDotDot", "sysDot",
             "dash", "dashDot", "dot", "lgDash", "lgDashDot", "lgDashDotDot",]

# 54種循環顏色
colors = ["4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646", "2C4D75", "772C2A", "5F7530", "4D3B62", 
          "276A7C", "B65708", "729ACA", "CD7371", "AFC97A", "9983B5", "6FBDD1", "F9AB6B", "3A679C", "9F3B38", 
          "7E9D40", "664F83", "358EA6", "F3740B", "95B3D7", "D99694", "C3D69B", "B3A2C7", "93CDDD", "FAC090", 
          "254061", "632523", "4F6228", "403152", "215968", "984807", "84A7D1", "D38482", "B9CF8B", "A692BE", 
          "81C5D7", "F9B67E", "335A88", "8B3431", "6F8938", "594573", "2E7C91", "D56509", "A7C0DE", "DFA8A6", 
          "CDDDAC", "BFB2D0", "A5D6E2", "FBCBA3", ]

# PICO TECH 溫度計
CHAN = ["Channel 1 Last (C)", "Channel 2 Last (C)", "Channel 3 Last (C)", "Channel 4 Last (C)",
        "Channel 5 Last (C)", "Channel 6 Last (C)", "Channel 7 Last (C)", "Channel 8 Last (C)"]

# 取得當前日期和時間
current_time   = datetime.now()
current_date   = current_time.strftime('%Y-%m-%d')
current_hour   = current_time.strftime('%H')
current_minute = current_time.strftime('%M')
current_second = current_time.strftime('%S')

# 讀取txt檔案
def read_file(f_path):
    if f_path.endswith('.TXT') or f_path.endswith('.txt'):
        with open(f_path, 'r', encoding = 'utf-8-sig') as file:
            lines = file.readlines()
            rows  = [line.split() for line in lines] # 按照空白字符分割每行內容並存入列表

    elif f_path.endswith('.CSV') or f_path.endswith('.csv'):
        with open(f_path, 'r', newline = '') as file:
            reader = csv.reader(file)
            rows   = [row for row in reader]
            
    return rows

# 溫度資料處理
def Temperature(f_path):
    # 儲存數據
    DATA = [] 
    
    try:
        rows = read_file(f_path)

        for row in rows:
            if rows.index(row) == 0:
                row[4] = "Ch.1"
                row[5] = "Ch.2"
                row[6] = "Ch.3"
                row[7] = "Ch.4"
                
            else:
                row[4] = float(row[4])
                row[5] = float(row[5])
                row[6] = float(row[6])
                row[7] = float(row[7])
                
                # 日期
                date_str = row[1]
                
                # 時間
                time_str = row[2]
                
                # 日期 + 時間
                datetime_str = f"{date_str} {time_str}"
                row[2] = datetime.strptime(datetime_str, "%Y-%m-%d %H:%M:%S")
                
            del row[8], row[3], row[1], row[0]

            DATA.append(row)
      
    except Exception as e:
        print(e)
        messagebox.showerror("錯誤", "檔案內容格式錯誤！")
    
    #print(DATA)
    return DATA

def Temperature_csv(f_path):
    try:
        rows = read_file(f_path) # 讀取資料
        rows[0][0] = "Date time" # 替換空點
        head = rows[0]           # 分割開頭
        body = rows[1:]          # 分割數據
        
        body = [
                [datetime.fromisoformat(row[0]).replace(tzinfo = None)] + [float(value) for value in row[1:]]
                for row in body
                if all(cell.strip() != '' for cell in row)
                ]

        for i1 in range(len(CHAN)):
            if CHAN[i1] not in head:
                head.insert(i1 + 1, CHAN[i1])
                for j2 in range(len(body)):
                    body[j2].insert(i1 + 1, 0)
                    
        DATA = [head] + body # 合併數據
        
    except Exception as e:
        print(e)
        messagebox.showerror("錯誤", "檔案內容格式錯誤！")
    
    #print(DATA)
    return DATA

# 繪圖模組-----------------------------------------------------------------------------------------

# 設定圖表標題格式
def set_chart_title_size(chart, size = 1400):
    paraprops = ParagraphProperties()
    paraprops.defRPr = CharacterProperties(sz = size)

    for para in chart.title.tx.rich.paragraphs:
        para.pPr = paraprops    

# Excel 圖表繪製
def Drawing_for_temp(DATA, colors, linetype, ws):
    # XY散佈圖
    chart = LineChart()
    chart.title = "Burn-In Test"
    set_chart_title_size(chart, size = 1400)
    chart.style = 13
    chart.y_axis.title = "Temperature (°C)"
    
    # X軸
    chart.x_axis.title = DATA[0][0]
    chart.x_axis.number_format = "h:mm:ss"
    chart.x_axis.majorTimeUnit = "hours"
    x_values=Reference(ws, min_col = 1, max_col = 1, min_row = 2, max_row = len(DATA))
    
    # Y軸
    for y in range(2, len(DATA[0]) + 1):
        y_values = Reference(ws, min_col = y, min_row = 1, max_row = len(DATA))
        series = Series(y_values, title_from_data = True)
        line_properties = LineProperties(w = 12700, solidFill = colors[y - 2], prstDash = linetype[0])
        series.graphicalProperties.line = line_properties
        chart.append(series)
    chart.y_axis.majorGridlines = openpyxl.chart.axis.ChartLines() # 打開格線
    
    # 設定線的樣式平滑曲線
    for x in range(len(DATA[0]) - 1):
        chart.series[x].smooth = True
    
    # 設定X軸標籤
    chart.set_categories(x_values)
    
    # 將圖表新增至工作表中，並設定大小
    adress = chr(65 + len(DATA[0])) + str("1")    
    
    chart.height = 7.5 # 設置高度
    chart.width  = 17  # 設置寬度
    
    return chart, adress

# 儲存至Excel
def Excel_file(f_path, DATA):
    
    # 取得檔案目錄
    dir_path = os.path.dirname(f_path)
    
    # 創建一個新的 Excel 工作簿
    wb = Workbook()
    
    # 將資料存進Temp分頁
    ws = wb.active
    ws.title = "Tempature"
    
    # 存入資料進Excel
    for row in DATA:
        ws.append(row)
    
    # 繪製Excel圖表
    chart, adress = Drawing_for_temp(DATA, colors, linetypes, ws) 

    # 儲存Excel圖表       
    ws.add_chart(chart, adress)
    
    try:
        # 儲存檔案
        save_path = f'{dir_path}/Tempature_Output.xlsx'
        wb.save(save_path)
        
        # 提示保存成功
        messagebox.showinfo("成功", "成功合併檔案：Tempature_Output.xlsx")
        
    except Exception as e:
        print("⚠️檢查 1：讀確認Tempature_Output.xlsx，檔案是否有開啟！")
        messagebox.showerror("錯誤", f"合併檔案時出錯：{str(e)}")

# GUI介面------------------------------------------------------------------------------------------

class Temp_App:
    def __init__(self, root):
        self.root = root
        self.root.title("Temp. merge files")
        self.window_width  = 260
        self.window_height = 400
        self.root.geometry(f"{self.window_width}x{self.window_height}")  # 設置窗口大小
        self.root.resizable(False, False) # 限制視窗大小

        # 瀏覽資料夾框架-----------------------------------------------------------------------------------------------
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
        self.latest_file = os.getcwd()
        self.update_listbox()

        # 溫度數據資料選項-------------------------------------------------------------------------------------------
        self.frame1 = ttk.LabelFrame(self.root, text = "溫度數據 / 標籤名稱：", relief = "groove", borderwidth = 2)
        self.frame1.place(x = 15, y = 80, width = 230, height = 275)
        self.frame1.config(style = "Dashed.TFrame")

        self.var_ch1 = tk.BooleanVar(value = True)
        self.var_ch2 = tk.BooleanVar(value = True)
        self.var_ch3 = tk.BooleanVar(value = True)
        self.var_ch4 = tk.BooleanVar(value = True)
        self.var_ch5 = tk.BooleanVar()
        self.var_ch6 = tk.BooleanVar()
        self.var_ch7 = tk.BooleanVar()
        self.var_ch8 = tk.BooleanVar()

        self.check_ch1 = ttk.Checkbutton(self.frame1, text = "熱電偶 1：", variable = self.var_ch1)
        self.check_ch2 = ttk.Checkbutton(self.frame1, text = "熱電偶 2：", variable = self.var_ch2)
        self.check_ch3 = ttk.Checkbutton(self.frame1, text = "熱電偶 3：", variable = self.var_ch3)
        self.check_ch4 = ttk.Checkbutton(self.frame1, text = "熱電偶 4：", variable = self.var_ch4)
        self.check_ch5 = ttk.Checkbutton(self.frame1, text = "熱電偶 5：", variable = self.var_ch5)
        self.check_ch6 = ttk.Checkbutton(self.frame1, text = "熱電偶 6：", variable = self.var_ch6)
        self.check_ch7 = ttk.Checkbutton(self.frame1, text = "熱電偶 7：", variable = self.var_ch7)
        self.check_ch8 = ttk.Checkbutton(self.frame1, text = "熱電偶 8：", variable = self.var_ch8)

        self.check_ch1.grid(row = 0, column = 0, padx = 5, pady = 5)
        self.check_ch2.grid(row = 1, column = 0, padx = 5, pady = 5)
        self.check_ch3.grid(row = 2, column = 0, padx = 5, pady = 5)
        self.check_ch4.grid(row = 3, column = 0, padx = 5, pady = 5)
        self.check_ch5.grid(row = 4, column = 0, padx = 5, pady = 5)
        self.check_ch6.grid(row = 5, column = 0, padx = 5, pady = 5)
        self.check_ch7.grid(row = 6, column = 0, padx = 5, pady = 5)
        self.check_ch8.grid(row = 7, column = 0, padx = 5, pady = 5)
        
        self.check = [self.check_ch1, self.check_ch2, self.check_ch3, self.check_ch4,
                      self.check_ch5, self.check_ch6, self.check_ch7, self.check_ch8]
        
        self.var   = [self.var_ch1, self.var_ch2, self.var_ch3, self.var_ch4,
                      self.var_ch5, self.var_ch6, self.var_ch7, self.var_ch8]

        # 變更標籤
        self.tag   = ["預設標籤：Ch.1", "預設標籤：Ch.2", "預設標籤：Ch.3", "預設標籤：Ch.4",
                      "預設標籤：Ch.5", "預設標籤：Ch.6", "預設標籤：Ch.7", "預設標籤：Ch.8"]
        
        self.entry_ch1 = tk.Entry(self.frame1, width = 17)
        self.entry_ch2 = tk.Entry(self.frame1, width = 17)
        self.entry_ch3 = tk.Entry(self.frame1, width = 17)
        self.entry_ch4 = tk.Entry(self.frame1, width = 17)
        self.entry_ch5 = tk.Entry(self.frame1, width = 17)
        self.entry_ch6 = tk.Entry(self.frame1, width = 17)
        self.entry_ch7 = tk.Entry(self.frame1, width = 17)
        self.entry_ch8 = tk.Entry(self.frame1, width = 17)
        
        self.entry_ch1.grid(row = 0, column = 1, padx = 0, pady = 5)
        self.entry_ch2.grid(row = 1, column = 1, padx = 0, pady = 5)
        self.entry_ch3.grid(row = 2, column = 1, padx = 0, pady = 5)
        self.entry_ch4.grid(row = 3, column = 1, padx = 0, pady = 5)
        self.entry_ch5.grid(row = 4, column = 1, padx = 0, pady = 5)
        self.entry_ch6.grid(row = 5, column = 1, padx = 0, pady = 5)
        self.entry_ch7.grid(row = 6, column = 1, padx = 0, pady = 5)
        self.entry_ch8.grid(row = 7, column = 1, padx = 0, pady = 5)
        
        self.entry_ch1.insert(0, self.tag[0])
        self.entry_ch2.insert(0, self.tag[1])
        self.entry_ch3.insert(0, self.tag[2])
        self.entry_ch4.insert(0, self.tag[3])
        self.entry_ch5.insert(0, self.tag[4])
        self.entry_ch6.insert(0, self.tag[5])
        self.entry_ch7.insert(0, self.tag[6])
        self.entry_ch8.insert(0, self.tag[7])
        
        self.entry = [self.entry_ch1, self.entry_ch2, self.entry_ch3, self.entry_ch4,
                      self.entry_ch5, self.entry_ch6, self.entry_ch7, self.entry_ch8]
        
        # 時間修正--------------------------------------------------------------------------------------------------
        self.frame_time = ttk.LabelFrame(self.root, text = "日期時間軸修正：", relief = "groove", borderwidth = 2)
        self.frame_time.place_forget()
        self.frame_time.config(style = "Dashed.TFrame")
        
        # 日期選擇器
        tk.Label(self.frame_time, text = "起始日期:").grid(row = 0, column = 0, padx = 5, pady = 5)
        self.cal = DateEntry(self.frame_time, date_pattern = 'yyyy-mm-dd', 
                             year  = current_time.year, 
                             month = current_time.month, 
                             day   = current_time.day)
        self.cal.grid(row = 0, column = 1, padx = 5, pady = 5, columnspan = 4, sticky = "w")
        
        # 時間選擇器
        tk.Label(self.frame_time, text = "起始時間:").grid(row = 1, column = 0, padx = 5, pady = 5)

        self.hour_var   = tk.StringVar(value = current_hour)
        self.minute_var = tk.StringVar(value = current_minute)
        self.second_var = tk.StringVar(value = current_second)

        self.hour_spinbox = tk.Spinbox(self.frame_time, from_ = 0, to = 23, wrap = True, 
                                       textvariable = self.hour_var, width = 3, format = "%02.0f")
        self.hour_spinbox.grid(row = 1, column = 1, padx = 5, pady = 5)

        self.minute_spinbox = tk.Spinbox(self.frame_time, from_ = 0, to = 59, wrap = True, 
                                         textvariable = self.minute_var, width = 3, format = "%02.0f")
        self.minute_spinbox.grid(row = 1, column = 2, padx = 5, pady = 5)

        self.second_spinbox = tk.Spinbox(self.frame_time, from_ = 0, to = 59, wrap = True, 
                                         textvariable = self.second_var, width = 3, format = "%02.0f")
        self.second_spinbox.grid(row = 1, column = 3, padx = 5, pady = 5)
        
        # 時間間隔
        self.time_gap_status = tk.Label(self.frame_time, text = '時間間隔：0 s')
        self.time_gap_status.grid(row = 2, column = 0, padx = 5, pady = 5, columnspan = 3, sticky = "w")


        # 執行/離開按鈕框架---------------------------------------------------------------------------------------
        self.button_frame = ttk.Frame(self.root)
        self.button_frame.place(x = 40, y = 360)

        self.execute_button = ttk.Button(self.button_frame, text = "執行", command = self.execute, width = 10)
        self.execute_button.grid(row = 0, column = 0, padx = 5, pady = 5)
        
        self.close_button = ttk.Button(self.button_frame, text = "離開", command = self.close_action, width = 10)
        self.close_button.grid(row = 0, column = 1, padx = 5, pady = 5)
        
        self.window_button = ttk.Button(self.button_frame, text = "▽", command = self.toggle_size, width = 2)
        self.window_button.grid(row = 0, column = 2, padx = 0, pady = 5)
        
        self.expanded = False

    # 瀏覽檔案
    def browse_folder(self):
        # 選擇檔案
        self.folder_path = filedialog.askopenfilename()
        self.execute_button["state"] = "normal" # 執行按鈕可用
        
        # 判定是否為txt檔案
        if self.folder_path.endswith('.TXT') or self.folder_path.endswith('.txt'):
            if self.folder_path:
                self.latest_file = self.folder_path
                self.update_listbox()
                self.time_gap()
                self.window_button["state"] = "normal" # 可用時間修正
                
                # 刷新選項狀態
                for i1 in range(4):
                    self.var[i1].set(1)
                    self.check[i1].config(state = "normal")
                    self.entry[i1].config(state = "normal")
                
                # 禁用Ch5~Ch8
                for i2 in range(4, len(self.check)):
                    self.var[i2].set(0)
                    self.check[i2].config(state = "disabled")
                    self.entry[i2].config(state = "disabled")
        
        elif self.folder_path.endswith('.CSV') or self.folder_path.endswith('.csv'):
            if self.folder_path:
                self.latest_file = self.folder_path
                self.update_listbox()
                
                # 禁用時間修正
                self.window_button["state"] = "disabled"
                if self.expanded:
                    self.toggle_size()
                    self.expanded = not self.expanded
                
                # 刷新選項狀態
                for i3 in range(len(self.check)):
                    self.var[i3].set(1)
                    self.check[i3].config(state = "normal")
                    self.entry[i3].config(state = "normal")
                
                # 禁用未使用Channel
                head = read_file(self.latest_file)[0]
                self.block_ch = []
                
                for i4 in range(len(CHAN)):
                    if CHAN[i4] not in head:
                       self.block_ch.append(i4)
                       
                for i5 in self.block_ch:
                    self.var[i5].set(0)
                    self.check[i5].config(state = "disabled")
                    self.entry[i5].config(state = "disabled")
                
                if len(self.block_ch) == 8:
                    self.execute_button["state"] = "disabled" # 執行按鈕禁用
                    messagebox.showerror("錯誤", "檔案內容格式錯誤，請確認是否為溫度計的資料！")
                    
        else:
            self.execute_button["state"] = "disabled" # 執行按鈕禁用
            messagebox.showerror("錯誤", "檔案不是.txt或.csv檔！")
              
    # 清除Listbox的內容，並插入最新選擇的資料夾路徑    
    def update_listbox(self):
        self.listbox.delete(0, tk.END)
        self.listbox.insert(tk.END, self.latest_file)             
    
    # 擴展視窗大小，啟動時間修正功能
    def toggle_size(self):
        if self.expanded:
            # 變更視窗大小
            self.window_width  = 260
            self.window_height = 400
            self.root.geometry(f"{self.window_width}x{self.window_height}")  # 設置窗口大小
            self.root.minsize(self.window_width, self.window_height) # 限制視窗大小
            self.root.maxsize(self.window_width, self.window_height) # 限制視窗大小
            
            # 改變框架位置
            self.frame_time.place_forget()
            self.button_frame.place(x = 40, y = 360)
            
            # 變更按鈕圖示
            self.window_button.config(text = '▽')
            
        else:
            # 變更視窗大小
            self.window_width  = 260
            self.window_height = 525
            self.root.geometry(f"{self.window_width}x{self.window_height}")  # 設置窗口大小
            self.root.minsize(self.window_width, self.window_height) # 限制視窗大小
            self.root.maxsize(self.window_width, self.window_height) # 限制視窗大小
            
            # 改變框架位置
            self.frame_time.place  (x = 15, y = 365, width = 230, height = 115)
            self.button_frame.place(x = 40, y = 485)

            # 變更按鈕圖示            
            self.window_button.config(text = '△')
            
        self.expanded = not self.expanded
    
    
    # 取得時間間隔
    def time_gap(self):
        try:
            rows = read_file(self.latest_file)
            
            # 取得時間單位
            self.t_unit = rows[1][3][-1]
            
            # 取得時間數值
            self.t_gap = ""
            for i in range(len(rows[1][3]) - 1):
                self.t_gap += rows[1][3][i]
            
            self.t_gap = int(self.t_gap)
            self.time_gap_status.config(text = f"時間間隔：{self.t_gap} {self.t_unit}")
            
        except Exception as e:
            print(e)
            self.execute_button["state"] = "disabled" # 執行按鈕禁用
            messagebox.showerror("錯誤", "檔案內容格式錯誤，請確認是否為溫度計的資料！")

    # 執行
    def execute(self):
        # 取得原始資料
        if self.folder_path.endswith('.TXT') or self.folder_path.endswith('.txt'):
            org_DATA = Temperature(self.latest_file)
            
        elif self.folder_path.endswith('.CSV') or self.folder_path.endswith('.csv'):
            org_DATA = Temperature_csv(self.latest_file)
            
        # 時間修正啟動驗證碼
        time_ck = True
        
        # 時間修正
        if self.expanded:
            date   = self.cal.get_date().strftime('%Y/%m/%d')
            hour   = self.hour_var.get()
            minute = self.minute_var.get()
            second = self.second_var.get()
            
            try:
                # 組合日期和時間並嘗試轉換為 datetime 對象
                datetime_text = f"{date} {hour}:{minute}:{second}"
                
                # 起始時間存檔
                org_DATA[1][0] = datetime.strptime(datetime_text, '%Y/%m/%d %H:%M:%S')
                
                # 判別溫度間隔單位
                if self.t_unit == "s":
                    T_gap = timedelta(seconds = self.t_gap)
                    
                elif self.t_unit == "m":
                    T_gap = timedelta(minutes = self.t_gap)
                
                # 覆蓋原時間資料
                for i1 in range(2,len(org_DATA)):
                    org_DATA[i1][0] = org_DATA[i1 - 1][0] + T_gap
                
            except ValueError:
                # 如果轉換失敗，顯示錯誤訊息
                messagebox.showerror("輸入錯誤", "請輸入有效的日期和時間！")
                
                # 時間修正啟動驗證碼
                time_ck = False
        
        if (self.var_ch1.get() == 1 or self.var_ch2.get() == 1 or self.var_ch3.get() == 1 or self.var_ch4.get() == 1 or
            self.var_ch5.get() == 1 or self.var_ch6.get() == 1 or self.var_ch7.get() == 1 or self.var_ch8.get() == 1):
            if time_ck:
                # 建立合併的資料
                DATA = []
                for i2 in range(len(org_DATA)):
                    DATA.append([])
                    DATA[i2].append(org_DATA[i2][0])
                
                for i3 in range(len(self.var)):
                    # 塞選想合併的資料
                    if self.var[i3].get():
                        # 變更資料標籤
                        if self.entry[i3].get() != self.tag[i3]:
                            org_DATA[0][i3 + 1] = self.entry[i3].get()
                        
                        # 合併Channel資料
                        for j3 in range(len(org_DATA)):
                            DATA[j3].append(org_DATA[j3][i3 + 1])
                #print(DATA)
                # 儲存至Excel
                Excel_file(self.latest_file, DATA)
                
        else:
            messagebox.showerror("錯誤", "注意未選擇資料！")

    # 離開
    def close_action(self):
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = Temp_App(root)
    root.mainloop()