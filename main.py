#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
興大學習日誌自動填寫工具
National Chung Hsing University Learning Journal Auto-Fill Tool

功能：
- 自動登入EZ-Come系統
- 批量填寫學習日誌（每日模式）
- 可視化瀏覽器操作
- 智能表單識別

版本：5.1 (僅修改兩項)
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import time
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import logging
import random
from typing import List, Dict
import json
import os

# 嘗試導入tkcalendar，如果沒有則使用備用方案
try:
    from tkcalendar import DateEntry
    HAS_TKCALENDAR = True
except ImportError:
    HAS_TKCALENDAR = False

class JournalAutoFiller:
    """學習日誌自動填寫主程式"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.driver = None
        self.is_running = False
        self.config_file = "config.json"
        self.setup_gui()
        self.setup_logging()
        self.load_config()
        
    def load_config(self):
        """載入配置檔案"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    
                # 載入帳號密碼
                self.username_var.set(config.get('username', ''))
                self.password_var.set(config.get('password', ''))
                self.school_id_var.set(config.get('school_id'))
                self.url_var.set(config.get('url', 'https://psf.nchu.edu.tw/punch/Menu.jsp'))
                
                # 載入校內編號列表
                school_ids = config.get('school_ids')
                self.school_combo['values'] = school_ids
                
                # 載入工作內容
                work_content = config.get('work_content', '')
                if work_content:
                    self.content_text.delete('1.0', tk.END)
                    self.content_text.insert('1.0', work_content)
                    
                self.logger.info("✅ 已載入配置檔案")
            else:
                self.logger.info("💡 配置檔案不存在，將使用預設值")
                
        except Exception as e:
            self.logger.error(f"❌ 載入配置檔案失敗: {e}")
            messagebox.showwarning("警告", f"載入配置檔案失敗: {e}")
            
    def save_config(self):
        """儲存配置檔案"""
        try:
            config = {
                'username': self.username_var.get(),
                'password': self.password_var.get(),
                'school_ids': list(self.school_combo['values']),  # 儲存所有校內編號選項
                'url': self.url_var.get(),
                'work_content': self.content_text.get('1.0', tk.END).strip()
            }
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
                
            self.logger.info("✅ 已儲存配置檔案")
            
        except Exception as e:
            self.logger.error(f"❌ 儲存配置檔案失敗: {e}")
            messagebox.showerror("錯誤", f"儲存配置檔案失敗: {e}")
            
    def clear_config(self):
        """清除配置檔案"""
        try:
            if os.path.exists(self.config_file):
                os.remove(self.config_file)
                
            # 清除界面
            self.username_var.set('')
            self.password_var.set('')
            self.school_combo['values'] = ('')  # 重置為預設選項
            self.url_var.set('https://psf.nchu.edu.tw/punch/Menu.jsp')
            self.content_text.delete('1.0', tk.END)
            
            self.logger.info("✅ 已清除配置檔案")
            messagebox.showinfo("成功", "配置檔案已清除")
            
        except Exception as e:
            self.logger.error(f"❌ 清除配置檔案失敗: {e}")
            messagebox.showerror("錯誤", f"清除配置檔案失敗: {e}")
            
    def add_school_id(self):
        """新增校內編號"""
        # 建立輸入對話框
        dialog = tk.Toplevel(self.root)
        dialog.title("新增校內編號")
        dialog.geometry("300x150")
        dialog.resizable(False, False)
        
        # 置中顯示
        dialog.transient(self.root)
        dialog.grab_set()
        
        frame = ttk.Frame(dialog, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="請輸入新的校內編號:").pack(pady=(0, 10))
        
        entry_var = tk.StringVar()
        entry = ttk.Entry(frame, textvariable=entry_var, width=20)
        entry.pack(pady=(0, 20))
        entry.focus()
        
        def add_id():
            new_id = entry_var.get().strip()
            if not new_id:
                messagebox.showwarning("警告", "請輸入校內編號")
                return
                
            # 檢查是否已存在
            current_values = list(self.school_combo['values'])
            if new_id in current_values:
                messagebox.showwarning("警告", "此校內編號已存在")
                return
                
            # 新增到列表
            current_values.append(new_id)
            self.school_combo['values'] = current_values
            self.school_id_var.set(new_id)  # 自動選擇新增的編號
            
            self.logger.info(f"✅ 已新增校內編號: {new_id}")
            messagebox.showinfo("成功", f"已新增校內編號: {new_id}")
            dialog.destroy()
            
        def cancel():
            dialog.destroy()
            
        button_frame = ttk.Frame(frame)
        button_frame.pack()
        
        ttk.Button(button_frame, text="新增", command=add_id).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="取消", command=cancel).pack(side=tk.LEFT)
        
        # 綁定Enter鍵
        entry.bind('<Return>', lambda e: add_id())
        
    def setup_gui(self):
        """設定GUI界面"""
        self.root.title("🎓 興大學習日誌自動填寫工具")
        self.root.geometry("800x1000")
        
        # 主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # 標題
        title = ttk.Label(main_frame, text="🎓 興大學習日誌自動填寫工具", 
                         font=('Microsoft JhengHei', 16, 'bold'))
        title.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # 登入設定區
        login_frame = ttk.LabelFrame(main_frame, text="🔐 登入設定", padding="15")
        login_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
        
        ttk.Label(login_frame, text="系統網址:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.url_var = tk.StringVar(value="https://psf.nchu.edu.tw/punch/Menu.jsp")
        ttk.Entry(login_frame, textvariable=self.url_var, width=50).grid(row=0, column=1, sticky=(tk.W, tk.E))
        
        ttk.Label(login_frame, text="校內帳號:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        self.username_var = tk.StringVar()
        ttk.Entry(login_frame, textvariable=self.username_var, width=30).grid(row=1, column=1, sticky=tk.W, pady=(10, 0))
        
        ttk.Label(login_frame, text="密碼:").grid(row=2, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        self.password_var = tk.StringVar()
        ttk.Entry(login_frame, textvariable=self.password_var, show="*", width=30).grid(row=2, column=1, sticky=tk.W, pady=(10, 0))
        
        login_frame.columnconfigure(1, weight=1)
        
        # 填寫設定區
        config_frame = ttk.LabelFrame(main_frame, text="📅 填寫設定", padding="15")
        config_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
        
        ttk.Label(config_frame, text="校內編號:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.school_id_var = tk.StringVar(value="")
        self.school_combo = ttk.Combobox(config_frame, textvariable=self.school_id_var, width=15)
        self.school_combo['values'] = ('')
        self.school_combo.grid(row=0, column=1, sticky=tk.W)
        
        # 新增校內編號按鈕
        ttk.Button(config_frame, text="➕", width=3,
                  command=self.add_school_id).grid(row=0, column=2, padx=(5, 0))
        ttk.Label(config_frame, text="新增完成後點選「儲存配置」永久保存", 
                 font=('Microsoft JhengHei', 8), foreground='gray').grid(row=0, column=3, sticky=tk.W, padx=(5, 0))
        
        # 日期設定 - 使用日期選擇器
        date_frame = ttk.Frame(config_frame)
        date_frame.grid(row=1, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(10, 0))
        
        if HAS_TKCALENDAR:
            # 使用tkcalendar的DateEntry
            ttk.Label(date_frame, text="開始日期:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
            self.start_date_picker = DateEntry(date_frame, width=12, background='darkblue',
                                             foreground='white', borderwidth=2, 
                                             date_pattern='yyyy-mm-dd', locale='zh_TW')
            self.start_date_picker.grid(row=0, column=1, sticky=tk.W)
            
            ttk.Label(date_frame, text="結束日期:").grid(row=0, column=2, sticky=tk.W, padx=(30, 10))
            self.end_date_picker = DateEntry(date_frame, width=12, background='darkblue',
                                           foreground='white', borderwidth=2,
                                           date_pattern='yyyy-mm-dd', locale='zh_TW')
            self.end_date_picker.grid(row=0, column=3, sticky=tk.W)
            
            # 設定預設日期
            today = datetime.now().date()
            week_ago = today - timedelta(days=7)
            self.start_date_picker.set_date(week_ago)
            self.end_date_picker.set_date(today)
            
        else:
            # 備用方案：使用下拉選單
            ttk.Label(date_frame, text="開始日期:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
            self.start_date_frame = ttk.Frame(date_frame)
            self.start_date_frame.grid(row=0, column=1, sticky=tk.W)
            
            # 開始日期選擇器
            today = datetime.now()
            week_ago = today - timedelta(days=7)
            
            self.start_year_var = tk.StringVar(value=str(week_ago.year))
            self.start_month_var = tk.StringVar(value=str(week_ago.month))
            self.start_day_var = tk.StringVar(value=str(week_ago.day))
            
            ttk.Combobox(self.start_date_frame, textvariable=self.start_year_var, width=6,
                        values=[str(y) for y in range(2020, 2030)]).grid(row=0, column=0)
            ttk.Label(self.start_date_frame, text="年").grid(row=0, column=1)
            ttk.Combobox(self.start_date_frame, textvariable=self.start_month_var, width=4,
                        values=[str(m) for m in range(1, 13)]).grid(row=0, column=2, padx=(5, 0))
            ttk.Label(self.start_date_frame, text="月").grid(row=0, column=3)
            ttk.Combobox(self.start_date_frame, textvariable=self.start_day_var, width=4,
                        values=[str(d) for d in range(1, 32)]).grid(row=0, column=4, padx=(5, 0))
            ttk.Label(self.start_date_frame, text="日").grid(row=0, column=5)
            
            ttk.Label(date_frame, text="結束日期:").grid(row=0, column=2, sticky=tk.W, padx=(30, 10))
            self.end_date_frame = ttk.Frame(date_frame)
            self.end_date_frame.grid(row=0, column=3, sticky=tk.W)
            
            # 結束日期選擇器
            self.end_year_var = tk.StringVar(value=str(today.year))
            self.end_month_var = tk.StringVar(value=str(today.month))
            self.end_day_var = tk.StringVar(value=str(today.day))
            
            ttk.Combobox(self.end_date_frame, textvariable=self.end_year_var, width=6,
                        values=[str(y) for y in range(2020, 2030)]).grid(row=0, column=0)
            ttk.Label(self.end_date_frame, text="年").grid(row=0, column=1)
            ttk.Combobox(self.end_date_frame, textvariable=self.end_month_var, width=4,
                        values=[str(m) for m in range(1, 13)]).grid(row=0, column=2, padx=(5, 0))
            ttk.Label(self.end_date_frame, text="月").grid(row=0, column=3)
            ttk.Combobox(self.end_date_frame, textvariable=self.end_day_var, width=4,
                        values=[str(d) for d in range(1, 32)]).grid(row=0, column=4, padx=(5, 0))
            ttk.Label(self.end_date_frame, text="日").grid(row=0, column=5)
            
        # 安裝提示
        if not HAS_TKCALENDAR:
            ttk.Label(date_frame, text="💡 安裝 tkcalendar 可使用日期選擇器：pip install tkcalendar", 
                     font=('Microsoft JhengHei', 8), foreground='gray').grid(row=1, column=0, columnspan=4, sticky=tk.W, pady=(5, 0))
        
        ttk.Label(config_frame, text="操作延遲:").grid(row=2, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        self.delay_var = tk.StringVar(value="1")  # 預設改為1秒
        delay_combo = ttk.Combobox(config_frame, textvariable=self.delay_var, width=10, 
                                  values=["1", "2", "3", "5"], state="readonly")
        delay_combo.grid(row=2, column=1, sticky=tk.W, pady=(10, 0))
        ttk.Label(config_frame, text="秒").grid(row=2, column=2, sticky=tk.W, padx=(5, 0), pady=(10, 0))
        
        # 工作內容區 - 移除模板按鈕
        content_frame = ttk.LabelFrame(main_frame, text="📝 工作內容", padding="15")
        content_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        main_frame.rowconfigure(3, weight=1)
        
        ttk.Label(content_frame, text="工作內容:").grid(row=0, column=0, sticky=(tk.W, tk.N), pady=(0, 5))
        self.content_text = scrolledtext.ScrolledText(content_frame, height=6, width=70)
        self.content_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        content_frame.columnconfigure(0, weight=1)
        content_frame.rowconfigure(1, weight=1)
        
        # 說明文字
        ttk.Label(content_frame, text="請輸入工作內容，程式會直接使用您輸入的內容填寫（不做任何修改）", 
                 font=('Microsoft JhengHei', 9), foreground='gray').grid(row=2, column=0, sticky=tk.W)
        
        # 配置檔案管理按鈕
        config_button_frame = ttk.Frame(content_frame)
        config_button_frame.grid(row=3, column=0, pady=10)
        
        ttk.Button(config_button_frame, text="💾 儲存配置", 
                  command=self.save_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(config_button_frame, text="📁 載入配置", 
                  command=self.load_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(config_button_frame, text="🗑️ 清除配置", 
                  command=self.clear_config).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(content_frame, text="(配置檔案: config.json)", 
                 font=('Microsoft JhengHei', 8), foreground='gray').grid(row=4, column=0, sticky=tk.W)
        
        # 執行狀態區
        status_frame = ttk.LabelFrame(main_frame, text="📊 執行狀態", padding="15")
        status_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
        
        self.status_var = tk.StringVar(value="就緒")
        ttk.Label(status_frame, textvariable=self.status_var, 
                 font=('Microsoft JhengHei', 11)).grid(row=0, column=0, sticky=tk.W)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(status_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        status_frame.columnconfigure(0, weight=1)
        
        # 統計資訊
        stats_frame = ttk.Frame(status_frame)
        stats_frame.grid(row=2, column=0, pady=5)
        
        self.total_var = tk.StringVar(value="總計: 0")
        self.success_var = tk.StringVar(value="成功: 0")
        self.failed_var = tk.StringVar(value="失敗: 0")
        
        ttk.Label(stats_frame, textvariable=self.total_var).pack(side=tk.LEFT, padx=20)
        ttk.Label(stats_frame, textvariable=self.success_var, foreground='green').pack(side=tk.LEFT, padx=20)
        ttk.Label(stats_frame, textvariable=self.failed_var, foreground='red').pack(side=tk.LEFT, padx=20)
        
        # 日誌區
        log_frame = ttk.LabelFrame(main_frame, text="📜 執行日誌", padding="15")
        log_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        main_frame.rowconfigure(5, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, state='disabled')
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # 控制按鈕
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, columnspan=2, pady=20)
        
        self.start_btn = ttk.Button(button_frame, text="🚀 開始執行", 
                                   command=self.start_execution)
        self.start_btn.pack(side=tk.LEFT, padx=10)
        
        self.stop_btn = ttk.Button(button_frame, text="⏹️ 停止執行", 
                                  command=self.stop_execution, state='disabled')
        self.stop_btn.pack(side=tk.LEFT, padx=10)
        
        ttk.Button(button_frame, text="🗑️ 清除日誌", 
                  command=self.clear_log).pack(side=tk.LEFT, padx=10)
        
        ttk.Button(button_frame, text="❓ 說明", 
                  command=self.show_help).pack(side=tk.LEFT, padx=10)
        
    def setup_logging(self):
        """設定日誌系統"""
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.INFO)
        
        class GUILogHandler(logging.Handler):
            def __init__(self, text_widget):
                super().__init__()
                self.text_widget = text_widget
                
            def emit(self, record):
                msg = self.format(record)
                self.text_widget.config(state='normal')
                self.text_widget.insert(tk.END, msg + '\n')
                self.text_widget.see(tk.END)
                self.text_widget.config(state='disabled')
                self.text_widget.update()
        
        handler = GUILogHandler(self.log_text)
        formatter = logging.Formatter('%(asctime)s - %(message)s', datefmt='%H:%M:%S')
        handler.setFormatter(formatter)
        self.logger.addHandler(handler)
        
    def clear_log(self):
        """清除日誌"""
        self.log_text.config(state='normal')
        self.log_text.delete('1.0', tk.END)
        self.log_text.config(state='disabled')
        
    def get_start_date(self):
        """取得開始日期字串"""
        if HAS_TKCALENDAR:
            return self.start_date_picker.get_date().strftime('%Y-%m-%d')
        else:
            return f"{self.start_year_var.get()}-{self.start_month_var.get().zfill(2)}-{self.start_day_var.get().zfill(2)}"
        
    def get_end_date(self):
        """取得結束日期字串"""
        if HAS_TKCALENDAR:
            return self.end_date_picker.get_date().strftime('%Y-%m-%d')
        else:
            return f"{self.end_year_var.get()}-{self.end_month_var.get().zfill(2)}-{self.end_day_var.get().zfill(2)}"
        
    def validate_inputs(self):
        """驗證輸入資料"""
        if not self.username_var.get().strip():
            messagebox.showerror("錯誤", "請輸入校內帳號")
            return False
            
        if not self.password_var.get().strip():
            messagebox.showerror("錯誤", "請輸入密碼")
            return False
            
        if not self.school_id_var.get().strip():
            messagebox.showerror("錯誤", "請輸入校內編號")
            return False
            
        # 驗證日期
        try:
            if HAS_TKCALENDAR:
                start_date = self.start_date_picker.get_date()
                end_date = self.end_date_picker.get_date()
                
                # 轉換為datetime以便比較
                start_datetime = datetime.combine(start_date, datetime.min.time())
                end_datetime = datetime.combine(end_date, datetime.min.time())
                
                if start_datetime > end_datetime:
                    messagebox.showerror("錯誤", "開始日期不能晚於結束日期")
                    return False
            else:
                start_date = datetime(
                    int(self.start_year_var.get()),
                    int(self.start_month_var.get()),
                    int(self.start_day_var.get())
                )
                end_date = datetime(
                    int(self.end_year_var.get()),
                    int(self.end_month_var.get()),
                    int(self.end_day_var.get())
                )
                
                if start_date > end_date:
                    messagebox.showerror("錯誤", "開始日期不能晚於結束日期")
                    return False
                
        except ValueError:
            messagebox.showerror("錯誤", "日期設定錯誤，請檢查日期")
            return False
            
        if not self.content_text.get('1.0', tk.END).strip():
            messagebox.showerror("錯誤", "請輸入工作內容")
            return False
            
        return True
        
    def start_execution(self):
        """開始執行"""
        if not self.validate_inputs():
            return
            
        if self.is_running:
            messagebox.showwarning("警告", "程式正在執行中")
            return
            
        # 計算天數
        if HAS_TKCALENDAR:
            start_date = self.start_date_picker.get_date()
            end_date = self.end_date_picker.get_date()
            days = (end_date - start_date).days + 1
        else:
            start_date = datetime(
                int(self.start_year_var.get()),
                int(self.start_month_var.get()),
                int(self.start_day_var.get())
            )
            end_date = datetime(
                int(self.end_year_var.get()),
                int(self.end_month_var.get()),
                int(self.end_day_var.get())
            )
            days = (end_date - start_date).days + 1
        
        result = messagebox.askyesno(
            "確認執行",
            f"即將開始自動填寫學習日誌\n\n"
            f"帳號: {self.username_var.get()}\n"
            f"校內編號: {self.school_id_var.get()}\n"
            f"日期範圍: {self.get_start_date()} ~ {self.get_end_date()}\n"
            f"填寫天數: {days} 天\n\n"
            f"確定要執行嗎？"
        )
        
        if not result:
            return
            
        # 在執行前先儲存配置
        self.save_config()
            
        # 設定執行狀態
        self.is_running = True
        self.start_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        self.status_var.set("正在執行...")
        self.progress_var.set(0)
        
        # 重置統計
        self.total_var.set("總計: 0")
        self.success_var.set("成功: 0")
        self.failed_var.set("失敗: 0")
        
        # 在新執行緒中執行
        thread = threading.Thread(target=self.execute_auto_fill, daemon=True)
        thread.start()
        
    def stop_execution(self):
        """停止執行"""
        self.is_running = False
        self.status_var.set("正在停止...")
        self.logger.info("⏹️ 使用者要求停止執行")
        
    def execute_auto_fill(self):
        """執行自動填寫"""
        try:
            bot = SeleniumBot(
                url=self.url_var.get(),
                username=self.username_var.get(),
                password=self.password_var.get(),
                school_id=self.school_id_var.get(),
                delay=int(self.delay_var.get()),
                logger=self.logger
            )
            
            results = bot.auto_fill_journals(
                start_date=self.get_start_date(),
                end_date=self.get_end_date(),
                base_content=self.content_text.get('1.0', tk.END).strip(),
                progress_callback=self.update_progress,
                stop_callback=lambda: self.is_running
            )
            
            if self.is_running:
                self.status_var.set("執行完成")
                self.logger.info("🎉 自動填寫執行完成！")
                messagebox.showinfo("完成", f"學習日誌自動填寫已完成！\n成功: {results['success']}/{results['total']}")
            else:
                self.status_var.set("已停止")
                
        except Exception as e:
            self.logger.error(f"❌ 執行錯誤: {e}")
            self.status_var.set("執行錯誤")
            messagebox.showerror("錯誤", f"執行失敗: {e}")
        finally:
            self.is_running = False
            self.start_btn.config(state='normal')
            self.stop_btn.config(state='disabled')
            
    def update_progress(self, current, total, success, failed):
        """更新進度"""
        progress = (current / total) * 100 if total > 0 else 0
        self.progress_var.set(progress)
        
        self.total_var.set(f"總計: {total}")
        self.success_var.set(f"成功: {success}")
        self.failed_var.set(f"失敗: {failed}")
        
        self.root.update_idletasks()
        
    def show_help(self):
        """顯示說明"""
        help_text = """🎓 興大學習日誌自動填寫工具使用說明

📋 功能特色:
• 每日填寫模式：選擇日期範圍，每天填寫一筆記錄
• 自動轉換日期為民國年格式 (yyymmdd)
• 直接使用您輸入的工作內容，不做修改
• 配置檔案功能：自動儲存帳號密碼和工作內容
• 可視化瀏覽器操作過程

🚀 使用步驟:
1. 輸入校內帳號和密碼
2. 選擇校內編號
3. 選擇開始和結束日期
4. 輸入工作內容
5. 點擊「開始執行」

💾 配置檔案功能:
• 儲存配置：保存帳號密碼、校內編號和工作內容到 config.json
• 載入配置：從 config.json 載入之前儲存的設定
• 清除配置：刪除配置檔案並重置界面
• 自動載入：程式啟動時自動載入配置
• 自動儲存：執行前自動儲存當前配置
• 校內編號管理：可新增自訂校內編號，會一起儲存

📅 填寫邏輯:
• 自動填入日期（民國年格式）
• 填入您輸入的工作內容（不做變化）
• 選擇校內編號
• 點擊「新增」按鈕
• 重新載入頁面繼續填寫下一天

⚠️ 注意事項:
• 建議先測試1-2天確認無誤
• 配置檔案會保存在程式目錄下
• 工作內容會原樣填入，請確保內容正確
• 預設操作延遲為1秒

🔧 環境需求:
• Python 3.6+
• selenium套件: pip install selenium
• tkcalendar套件: pip install tkcalendar (可選)
• Chrome瀏覽器
• ChromeDriver (可自動下載)"""
        
        help_window = tk.Toplevel(self.root)
        help_window.title("使用說明")
        help_window.geometry("600x500")
        
        help_frame = ttk.Frame(help_window, padding="20")
        help_frame.pack(fill=tk.BOTH, expand=True)
        
        help_scroll = scrolledtext.ScrolledText(help_frame, wrap=tk.WORD)
        help_scroll.pack(fill=tk.BOTH, expand=True)
        help_scroll.insert('1.0', help_text)
        help_scroll.config(state='disabled')
        
        ttk.Button(help_frame, text="關閉", 
                  command=help_window.destroy).pack(pady=15)
                  
    def run(self):
        """啟動程式"""
        self.root.mainloop()


class SeleniumBot:
    """Selenium自動化機器人"""
    
    def __init__(self, url, username, password, school_id, delay, logger):
        self.url = url
        self.username = username
        self.password = password
        self.school_id = school_id
        self.delay = delay
        self.logger = logger
        self.driver = None
        
    def create_driver(self):
        """建立WebDriver"""
        try:
            options = Options()
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--disable-gpu')
            options.add_argument('--window-size=1280,720')
            
            # 指定同目錄下的 chromedriver.exe 路徑
            chromedriver_path = os.path.join(os.getcwd(), 'chromedriver.exe')
            
            # 檢查 chromedriver.exe 是否存在於同目錄下
            if not os.path.exists(chromedriver_path):
                self.logger.error("❌ 找不到 chromedriver.exe，請確保它位於程式的同一目錄下")
                return None
            
            # 使用 Service 來設置 chromedriver 路徑
            service = Service(executable_path=chromedriver_path)
            
            # 使用指定的 chromedriver 路徑啟動瀏覽器
            return webdriver.Chrome(service=service, options=options)
                
        except Exception as e:
            self.logger.error(f"❌ 建立WebDriver失敗: {e}")
            return None
            
    def login(self):
        """登入系統"""
        try:
            self.logger.info("🌐 正在開啟瀏覽器...")
            self.driver = self.create_driver()
            if not self.driver:
                return False
                
            self.logger.info(f"🔗 正在訪問: {self.url}")
            self.driver.get(self.url)
            time.sleep(3)
            
            self.logger.info("🔍 正在尋找登入表單...")
            wait = WebDriverWait(self.driver, 10)
            
            # 尋找帳號欄位 - 使用具體的欄位名稱
            try:
                username_input = wait.until(EC.presence_of_element_located((By.ID, "txtLoginID")))
                self.logger.info("✅ 找到帳號欄位 (txtLoginID)")
            except TimeoutException:
                try:
                    username_input = wait.until(EC.presence_of_element_located((By.NAME, "txtLoginID")))
                    self.logger.info("✅ 找到帳號欄位 (name=txtLoginID)")
                except TimeoutException:
                    self.logger.error("❌ 找不到帳號輸入欄位")
                    return False
                    
            # 尋找密碼欄位
            try:
                password_input = self.driver.find_element(By.ID, "txtLoginPWD")
                self.logger.info("✅ 找到密碼欄位 (txtLoginPWD)")
            except NoSuchElementException:
                try:
                    password_input = self.driver.find_element(By.NAME, "txtLoginPWD")
                    self.logger.info("✅ 找到密碼欄位 (name=txtLoginPWD)")
                except NoSuchElementException:
                    self.logger.error("❌ 找不到密碼輸入欄位")
                    return False
            
            # 輸入帳號密碼
            self.logger.info("⌨️ 正在輸入帳號密碼...")
            username_input.clear()
            username_input.send_keys(self.username)
            time.sleep(0.5)
            
            password_input.clear()
            password_input.send_keys(self.password)
            time.sleep(0.5)
            
            # 提交登入 - 使用具體的按鈕ID
            try:
                submit_btn = self.driver.find_element(By.ID, "button")
                self.logger.info("🖱️ 點擊登入按鈕")
                submit_btn.click()
            except NoSuchElementException:
                try:
                    submit_btn = self.driver.find_element(By.CSS_SELECTOR, "input[value='登入']")
                    self.logger.info("🖱️ 使用備用方式點擊登入按鈕")
                    submit_btn.click()
                except NoSuchElementException:
                    self.logger.info("🔍 按Enter鍵登入")
                    from selenium.webdriver.common.keys import Keys
                    password_input.send_keys(Keys.RETURN)
            
            time.sleep(3)
            
            # 檢查登入結果
            if any(keyword in self.driver.page_source for keyword in ["登出", "logout", self.username, "Menu"]):
                self.logger.info("✅ 登入成功")
                return True
            else:
                self.logger.error("❌ 登入失敗")
                return False
                
        except Exception as e:
            self.logger.error(f"❌ 登入過程發生錯誤: {e}")
            return False
    
    def navigate_to_journal(self):
        """導航到學習日誌頁面"""
        try:
            self.logger.info("📋 正在導航到學習日誌頁面...")
            
            # 方法1: 嘗試點擊左側選單中的學習日誌連結
            try:
                # 常見的學習日誌連結文字和選擇器
                journal_selectors = [
                    "//a[contains(text(), '學習日誌')]",
                    "//a[contains(text(), '日誌')]", 
                    "//a[contains(@href, 'PunchList_A')]",
                    "//li//a[contains(text(), '學習日誌')]",
                    "//ul//a[contains(text(), '學習日誌')]",
                    "//div//a[contains(text(), '學習日誌')]"
                ]
                
                for selector in journal_selectors:
                    try:
                        journal_link = WebDriverWait(self.driver, 3).until(
                            EC.element_to_be_clickable((By.XPATH, selector))
                        )
                        self.logger.info(f"✅ 找到學習日誌連結，準備點擊: {selector}")
                        journal_link.click()
                        time.sleep(2)
                        
                        # 檢查是否成功導航到學習日誌頁面
                        if "PunchList_A" in self.driver.current_url or "學習日誌" in self.driver.page_source:
                            self.logger.info("✅ 成功點擊學習日誌連結")
                            return True
                            
                    except (TimeoutException, NoSuchElementException):
                        continue
                        
            except Exception as e:
                self.logger.warning(f"點擊連結失敗: {e}")
            
            # 方法2: 嘗試在iframe中尋找連結
            try:
                # 檢查是否有iframe
                iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
                self.logger.info(f"找到 {len(iframes)} 個iframe")
                
                for i, iframe in enumerate(iframes):
                    try:
                        self.driver.switch_to.frame(iframe)
                        self.logger.info(f"切換到iframe {i+1}")
                        
                        # 在iframe中尋找學習日誌連結
                        for selector in journal_selectors:
                            try:
                                journal_link = self.driver.find_element(By.XPATH, selector)
                                self.logger.info(f"✅ 在iframe中找到學習日誌連結")
                                journal_link.click()
                                time.sleep(2)
                                self.driver.switch_to.default_content()
                                return True
                            except NoSuchElementException:
                                continue
                                
                        self.driver.switch_to.default_content()
                        
                    except Exception as e:
                        self.driver.switch_to.default_content()
                        continue
                        
            except Exception as e:
                self.logger.warning(f"iframe處理失敗: {e}")
            
            # 方法3: 直接訪問學習日誌頁面URL
            try:
                # 根據登入後的URL構建學習日誌頁面URL
                current_url = self.driver.current_url
                self.logger.info(f"當前URL: {current_url}")
                
                # 嘗試不同的學習日誌URL模式
                base_url = current_url.split('/punch/')[0] if '/punch/' in current_url else None
                if not base_url:
                    base_url = self.url.split('/punch/')[0] if '/punch/' in self.url else self.url.rstrip('/')
                
                possible_urls = [
                    f"{base_url}/punch/PunchList_A.jsp",
                    f"{base_url}/PunchList_A.jsp",
                    f"{base_url}/punch/journal.jsp",
                    f"{base_url}/journal.jsp"
                ]
                
                for journal_url in possible_urls:
                    try:
                        self.logger.info(f"🔗 嘗試直接訪問: {journal_url}")
                        self.driver.get(journal_url)
                        time.sleep(2)
                        
                        # 檢查頁面是否包含學習日誌相關內容
                        page_source = self.driver.page_source
                        if any(keyword in page_source for keyword in ["學習日誌", "工作內容", "date", "work"]):
                            self.logger.info("✅ 成功訪問學習日誌頁面")
                            return True
                            
                    except Exception as e:
                        self.logger.warning(f"訪問 {journal_url} 失敗: {e}")
                        continue
                        
            except Exception as e:
                self.logger.warning(f"直接訪問失敗: {e}")
            
            # 方法4: 手動指導用戶
            self.logger.info("🤖 自動導航失敗，嘗試手動協助...")
            self.logger.info("💡 請檢查頁面上是否有'學習日誌'選項")
            
            # 列出頁面上所有可能的連結供參考
            try:
                links = self.driver.find_elements(By.TAG_NAME, "a")
                link_texts = [link.text.strip() for link in links if link.text.strip()]
                self.logger.info(f"頁面上的連結: {link_texts[:10]}")  # 只顯示前10個
            except:
                pass
            
            # 給用戶一些時間手動點擊
            self.logger.info("⏳ 等待10秒，您可以手動點擊'學習日誌'連結...")
            time.sleep(10)
            
            # 檢查是否已經在學習日誌頁面
            if "PunchList_A" in self.driver.current_url or any(keyword in self.driver.page_source for keyword in ["學習日誌", "工作內容", "date", "work"]):
                self.logger.info("✅ 檢測到已在學習日誌頁面")
                return True
            
            self.logger.error("❌ 無法導航到學習日誌頁面")
            return False
            
        except Exception as e:
            self.logger.error(f"❌ 導航過程發生錯誤: {e}")
            return False
    
    def generate_dates(self, start_date: str, end_date: str) -> List[str]:
        """生成每日日期列表"""
        start = datetime.strptime(start_date, '%Y-%m-%d')
        end = datetime.strptime(end_date, '%Y-%m-%d')
        
        dates = []
        current = start
        while current <= end:
            dates.append(current.strftime('%Y-%m-%d'))
            current += timedelta(days=1)
            
        return dates
    
    def generate_content(self, base_content: str, index: int) -> str:
        """直接返回工作內容，不做變化"""
        return base_content.strip()
    
    def fill_journal_entry(self, date: str, content: str) -> bool:
        """填寫單筆學習日誌"""
        try:
            self.logger.info(f"📝 正在填寫 {date} 的學習日誌...")
            
            # 等待頁面完全載入
            self.logger.info("⏳ 等待頁面載入...")
            time.sleep(3)
            
            # 檢查是否在iframe中
            try:
                iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
                if iframes:
                    self.logger.info(f"發現 {len(iframes)} 個iframe，嘗試切換...")
                    for i, iframe in enumerate(iframes):
                        try:
                            self.driver.switch_to.frame(iframe)
                            self.logger.info(f"切換到iframe {i+1}")
                            
                            # 檢查是否有學習日誌表單
                            if self.driver.find_elements(By.ID, "date"):
                                self.logger.info("✅ 在iframe中找到學習日誌表單")
                                break
                            else:
                                self.driver.switch_to.default_content()
                        except:
                            self.driver.switch_to.default_content()
                            continue
            except:
                pass
            
            # 轉換日期格式為民國年
            date_obj = datetime.strptime(date, '%Y-%m-%d')
            roc_year = date_obj.year - 1911
            formatted_date = f"{roc_year:03d}{date_obj.month:02d}{date_obj.day:02d}"
            
            self.logger.info(f"📅 日期轉換: {date} -> {formatted_date}")
            
            # 等待並尋找日期欄位
            self.logger.info("🔍 正在尋找日期欄位...")
            date_input = None
            
            # 使用WebDriverWait等待元素出現
            try:
                wait = WebDriverWait(self.driver, 10)
                date_input = wait.until(EC.presence_of_element_located((By.ID, "date")))
                self.logger.info("✅ 找到日期欄位 (id=date)")
            except TimeoutException:
                try:
                    date_input = wait.until(EC.presence_of_element_located((By.NAME, "date")))
                    self.logger.info("✅ 找到日期欄位 (name=date)")
                except TimeoutException:
                    try:
                        date_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder*='民國yyymmdd']")))
                        self.logger.info("✅ 找到日期欄位 (placeholder)")
                    except TimeoutException:
                        self.logger.error("❌ 等待日期欄位超時")
                        
                        # 列出頁面上所有input欄位
                        try:
                            inputs = self.driver.find_elements(By.TAG_NAME, "input")
                            input_info = []
                            for inp in inputs:
                                inp_id = inp.get_attribute('id')
                                inp_name = inp.get_attribute('name')
                                inp_type = inp.get_attribute('type')
                                inp_placeholder = inp.get_attribute('placeholder')
                                input_info.append(f"id:{inp_id}, name:{inp_name}, type:{inp_type}, placeholder:{inp_placeholder}")
                            self.logger.info(f"頁面上的input欄位: {input_info}")
                        except:
                            pass
                            
                        return False
            
            if date_input:
                try:
                    # 確保欄位是可見和可互動的
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", date_input)
                    time.sleep(0.5)
                    
                    date_input.clear()
                    time.sleep(0.3)
                    date_input.send_keys(formatted_date)
                    self.logger.info(f"✅ 已填入日期: {formatted_date}")
                except Exception as e:
                    self.logger.error(f"❌ 填入日期失敗: {e}")
                    return False
            else:
                self.logger.error("❌ 找不到日期欄位")
                return False
            
            time.sleep(0.5)
            
            # 尋找工作內容欄位
            self.logger.info("🔍 正在尋找工作內容欄位...")
            work_input = None
            
            try:
                work_input = self.driver.find_element(By.ID, "work")
                self.logger.info("✅ 找到工作內容欄位 (id=work)")
            except NoSuchElementException:
                try:
                    work_input = self.driver.find_element(By.NAME, "work")
                    self.logger.info("✅ 找到工作內容欄位 (name=work)")
                except NoSuchElementException:
                    try:
                        work_input = self.driver.find_element(By.CSS_SELECTOR, "input[required='ture']")
                        # 確保不是日期欄位
                        if work_input.get_attribute('id') != 'date':
                            self.logger.info("✅ 找到工作內容欄位 (required=ture)")
                        else:
                            work_input = None
                    except NoSuchElementException:
                        pass
            
            if work_input:
                try:
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", work_input)
                    time.sleep(0.5)
                    
                    work_input.clear()
                    time.sleep(0.3)
                    work_input.send_keys(content)
                    self.logger.info("✅ 已填入工作內容")
                except Exception as e:
                    self.logger.error(f"❌ 填入工作內容失敗: {e}")
                    return False
            else:
                self.logger.error("❌ 找不到工作內容欄位")
                return False
            
            time.sleep(0.5)
            
            # 尋找校內編號選單
            self.logger.info("🔍 正在尋找校內編號選單...")
            school_select = None
            
            try:
                school_select = self.driver.find_element(By.ID, "schno")
                self.logger.info("✅ 找到校內編號選單 (id=schno)")
            except NoSuchElementException:
                try:
                    school_select = self.driver.find_element(By.NAME, "schno")
                    self.logger.info("✅ 找到校內編號選單 (name=schno)")
                except NoSuchElementException:
                    try:
                        school_select = self.driver.find_element(By.TAG_NAME, "select")
                        self.logger.info("✅ 找到校內編號選單 (select tag)")
                    except NoSuchElementException:
                        pass
            
            if school_select:
                try:
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", school_select)
                    time.sleep(0.5)
                    
                    select = Select(school_select)
                    select.select_by_value(self.school_id)
                    self.logger.info(f"✅ 已選擇校內編號: {self.school_id}")
                except Exception as e:
                    self.logger.warning(f"⚠️ 選擇校內編號失敗: {e}")
                    # 列出可用選項
                    try:
                        options = [opt.get_attribute('value') for opt in school_select.find_elements(By.TAG_NAME, "option")]
                        self.logger.info(f"可用的校內編號選項: {options}")
                    except:
                        pass
            else:
                self.logger.warning("⚠️ 找不到校內編號選單")
            
            time.sleep(1)
            
            # 尋找新增按鈕 (根據HTML結構，按鈕id是btnSent)
            self.logger.info("🔍 正在尋找新增按鈕...")
            submit_btn = None
            
            try:
                submit_btn = self.driver.find_element(By.ID, "btnSent")
                self.logger.info("✅ 找到新增按鈕 (id=btnSent)")
            except NoSuchElementException:
                try:
                    submit_btn = self.driver.find_element(By.NAME, "btnSent")
                    self.logger.info("✅ 找到新增按鈕 (name=btnSent)")
                except NoSuchElementException:
                    try:
                        submit_btn = self.driver.find_element(By.CSS_SELECTOR, "input[value*='新增']")
                        self.logger.info("✅ 找到新增按鈕 (value contains 新增)")
                    except NoSuchElementException:
                        try:
                            submit_btn = self.driver.find_element(By.CSS_SELECTOR, "input[onclick*='add']")
                            self.logger.info("✅ 找到新增按鈕 (onclick contains add)")
                        except NoSuchElementException:
                            pass
            
            if submit_btn:
                try:
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", submit_btn)
                    time.sleep(0.5)
                    
                    self.logger.info("🖱️ 點擊新增按鈕")
                    submit_btn.click()
                except Exception as e:
                    self.logger.error(f"❌ 點擊新增按鈕失敗: {e}")
                    return False
            else:
                self.logger.error("❌ 找不到新增按鈕")
                # 列出頁面上所有按鈕
                try:
                    buttons = self.driver.find_elements(By.TAG_NAME, "input") + self.driver.find_elements(By.TAG_NAME, "button")
                    button_info = []
                    for btn in buttons:
                        btn_id = btn.get_attribute('id')
                        btn_name = btn.get_attribute('name')
                        btn_type = btn.get_attribute('type')
                        btn_value = btn.get_attribute('value')
                        btn_onclick = btn.get_attribute('onclick')
                        button_info.append(f"id:{btn_id}, name:{btn_name}, type:{btn_type}, value:{btn_value}, onclick:{btn_onclick}")
                    self.logger.info(f"頁面上的按鈕: {button_info}")
                except:
                    pass
                return False
            
            time.sleep(3)
            
            # 切回主框架
            try:
                self.driver.switch_to.default_content()
            except:
                pass
            
            # 檢查提交結果
            page_source = self.driver.page_source
            if any(keyword in page_source for keyword in ["成功", "完成", "新增完成", "儲存成功", "success"]):
                self.logger.info(f"✅ {date} 學習日誌填寫成功")
                return True
            elif any(keyword in page_source for keyword in ["錯誤", "失敗", "重複", "已存在", "error"]):
                self.logger.warning(f"⚠️ {date} 學習日誌填寫失敗 - 可能已存在或有錯誤")
                return False
            else:
                self.logger.info(f"✅ {date} 學習日誌已提交（未確認狀態訊息）")
                return True
                
        except Exception as e:
            # 確保切回主框架
            try:
                self.driver.switch_to.default_content()
            except:
                pass
            self.logger.error(f"❌ 填寫 {date} 時發生錯誤: {e}")
            return False
    
    def auto_fill_journals(self, start_date: str, end_date: str, base_content: str, 
                          progress_callback=None, stop_callback=None) -> Dict:
        """自動批量填寫學習日誌 - 每日填寫模式"""
        results = {
            'total': 0,
            'success': 0,
            'failed': 0,
            'details': []
        }
        
        try:
            # 登入系統
            if not self.login():
                return results
            
            # 導航到學習日誌頁面
            if not self.navigate_to_journal():
                return results
            
            # 生成每日日期列表
            dates = self.generate_dates(start_date, end_date)
            results['total'] = len(dates)
            
            self.logger.info(f"🚀 開始每日填寫模式，共 {len(dates)} 天")
            self.logger.info(f"📅 日期範圍: {dates[0]} ~ {dates[-1]}")
            self.logger.info(f"🆔 校內編號: {self.school_id}")
            
            # 逐日填寫
            for i, date in enumerate(dates):
                # 檢查是否要停止
                if stop_callback and not stop_callback():
                    self.logger.info("⏹️ 執行被使用者停止")
                    break
                
                # 生成當日內容 (直接使用用戶輸入的內容)
                content = self.generate_content(base_content, i)
                
                # 更新進度
                if progress_callback:
                    progress_callback(i + 1, len(dates), results['success'], results['failed'])
                
                # 填寫當日學習日誌
                success = self.fill_journal_entry(date, content)
                
                # 記錄結果
                result_detail = {
                    'date': date,
                    'content': content,
                    'school_id': self.school_id,
                    'success': success
                }
                results['details'].append(result_detail)
                
                if success:
                    results['success'] += 1
                else:
                    results['failed'] += 1
                
                # 填寫完成後，需要重新導航到新增頁面（因為已提交）
                if i < len(dates) - 1:  # 不是最後一筆
                    self.logger.info(f"⏳ 等待 {self.delay} 秒後繼續下一筆...")
                    time.sleep(self.delay)
                    
                    # 重新導航到學習日誌新增頁面
                    try:
                        self.navigate_to_journal()
                        time.sleep(1)
                    except Exception as e:
                        self.logger.warning(f"⚠️ 重新導航失敗: {e}")
            
            # 最終進度更新
            if progress_callback:
                progress_callback(len(dates), len(dates), results['success'], results['failed'])
            
            # 輸出統計結果
            self.logger.info("="*50)
            self.logger.info("🎉 每日填寫執行完成！統計結果：")
            self.logger.info(f"📊 總計: {results['total']} 天")
            self.logger.info(f"✅ 成功: {results['success']} 天")
            self.logger.info(f"❌ 失敗: {results['failed']} 天")
            if results['total'] > 0:
                success_rate = results['success'] / results['total'] * 100
                self.logger.info(f"📈 成功率: {success_rate:.1f}%")
            self.logger.info("="*50)
            
        except Exception as e:
            self.logger.error(f"❌ 批量填寫過程發生錯誤: {e}")
        finally:
            # 關閉瀏覽器
            if self.driver:
                try:
                    self.logger.info("🔒 正在關閉瀏覽器...")
                    self.driver.quit()
                except:
                    pass
        
        return results


def main():
    """主程式入口"""
    print("🎓 興大學習日誌自動填寫工具 v5.2")
    print("="*50)
    
    try:
        # 檢查Selenium
        import selenium
        print("✅ Selenium 已安裝")
    except ImportError:
        print("❌ 請先安裝 Selenium: pip install selenium")
        input("按任意鍵退出...")
        return
    
    # 檢查tkcalendar
    if HAS_TKCALENDAR:
        print("✅ tkcalendar 已安裝（提供日期選擇器）")
    else:
        print("💡 建議安裝 tkcalendar 以使用日期選擇器: pip install tkcalendar")
        print("   目前使用備用的下拉選單")
    
    print("\n🚀 正在啟動程式...")
    
    try:
        app = JournalAutoFiller()
        app.run()
    except Exception as e:
        print(f"❌ 程式啟動失敗: {e}")
        try:
            messagebox.showerror("錯誤", f"程式啟動失敗: {e}")
        except:
            pass


if __name__ == "__main__":
    main()