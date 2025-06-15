import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import pyautogui
import pytesseract
from PIL import Image, ImageTk, ImageDraw, ImageFont, ImageEnhance
import cv2
import numpy as np
from googletrans import Translator
import threading
import time
import keyboard
from datetime import datetime
import json
import os
import sys
from collections import deque

# 語言配置
LANGUAGES = {
    # 東亞語言
    'jpn': {'name': '日文', 'google_code': 'ja'},
    'kor': {'name': '韓文', 'google_code': 'ko'},
    'chi_sim': {'name': '簡體中文', 'google_code': 'zh-cn'},
    'chi_tra': {'name': '繁體中文', 'google_code': 'zh-tw'},
    
    # 歐洲語言
    'eng': {'name': '英文', 'google_code': 'en'},
    'fra': {'name': '法文', 'google_code': 'fr'},
    'deu': {'name': '德文', 'google_code': 'de'},
    'spa': {'name': '西班牙文', 'google_code': 'es'},
    'ita': {'name': '義大利文', 'google_code': 'it'},
    'por': {'name': '葡萄牙文', 'google_code': 'pt'},
    'rus': {'name': '俄文', 'google_code': 'ru'},
    
    # 其他語言
    'ara': {'name': '阿拉伯文', 'google_code': 'ar'},
    'tha': {'name': '泰文', 'google_code': 'th'},
    'vie': {'name': '越南文', 'google_code': 'vi'},
    'ind': {'name': '印尼文', 'google_code': 'id'},
    'tur': {'name': '土耳其文', 'google_code': 'tr'},
    'pol': {'name': '波蘭文', 'google_code': 'pl'},
    'nld': {'name': '荷蘭文', 'google_code': 'nl'},
    'swe': {'name': '瑞典文', 'google_code': 'sv'},
}

# 目標語言選項
TARGET_LANGUAGES = {
    'zh-tw': '繁體中文',
    'zh-cn': '簡體中文',
    'en': '英文',
    'ja': '日文',
    'ko': '韓文',
    'es': '西班牙文',
    'fr': '法文',
    'de': '德文',
    'ru': '俄文',
    'ar': '阿拉伯文',
    'th': '泰文',
    'vi': '越南文',
}

class MultilingualGameTranslator:
    def __init__(self, root):
        self.root = root
        self.root.title("多語言遊戲翻譯器 v3.0")
        self.root.geometry("1200x800")
        
        # 初始化
        self.translator = Translator()
        self.is_capturing = False
        self.capture_region = None
        self.translation_cache = {}
        self.translation_history = deque(maxlen=500)
        
        # 檢查已安裝的語言
        self.check_installed_languages()
        
        # 設定
        self.settings = {
            'source_language': 'jpn',  # 預設日文
            'target_language': 'zh-tw',  # 預設繁中
            'ocr_mode': 'single',  # single 或 multi
            'update_interval': 0.5,
            'preprocessing': True,
            'auto_detect': False,
            'confidence_threshold': 60
        }
        
        # 載入設定
        self.load_settings()
        
        # 設定樣式
        self.setup_styles()
        
        # 建立UI
        self.create_ui()
        
        # 設定快捷鍵
        self.setup_hotkeys()
        
    def check_installed_languages(self):
        """檢查已安裝的 Tesseract 語言包"""
        try:
            # 取得可用語言列表
            languages = pytesseract.get_languages()
            self.installed_languages = {}
            
            for lang_code, lang_info in LANGUAGES.items():
                if lang_code in languages:
                    self.installed_languages[lang_code] = lang_info
                    
            print(f"已安裝的語言: {list(self.installed_languages.keys())}")
        except Exception as e:
            print(f"無法檢查語言包: {e}")
            self.installed_languages = LANGUAGES  # 假設全部都有
            
    def setup_styles(self):
        """設定視覺樣式"""
        self.root.configure(bg='#1e1e1e')
        
        style = ttk.Style()
        style.theme_use('clam')
        
        # 深色主題
        style.configure('TNotebook', background='#1e1e1e')
        style.configure('TNotebook.Tab', background='#2d2d2d', foreground='white')
        style.configure('TFrame', background='#1e1e1e')
        style.configure('TLabel', background='#1e1e1e', foreground='white')
        style.configure('TLabelframe', background='#1e1e1e', foreground='white')
        style.configure('TLabelframe.Label', background='#1e1e1e', foreground='white')
        
    def create_ui(self):
        """建立使用者介面"""
        # 建立筆記本（分頁）
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 主要翻譯頁面
        main_tab = ttk.Frame(notebook)
        notebook.add(main_tab, text="翻譯")
        self.create_main_tab(main_tab)
        
        # 語言設定頁面
        lang_tab = ttk.Frame(notebook)
        notebook.add(lang_tab, text="語言設定")
        self.create_language_tab(lang_tab)
        
        # 進階設定頁面
        settings_tab = ttk.Frame(notebook)
        notebook.add(settings_tab, text="進階設定")
        self.create_settings_tab(settings_tab)
        
        # 歷史記錄頁面
        history_tab = ttk.Frame(notebook)
        notebook.add(history_tab, text="歷史")
        self.create_history_tab(history_tab)
        
        # 狀態列
        self.create_status_bar()
        
    def create_main_tab(self, parent):
        """建立主要翻譯介面"""
        # 頂部控制面板
        control_frame = tk.Frame(parent, bg='#2d2d2d', height=100)
        control_frame.pack(fill=tk.X, padx=5, pady=5)
        control_frame.pack_propagate(False)
        
        # 語言選擇區域
        lang_frame = tk.Frame(control_frame, bg='#2d2d2d')
        lang_frame.pack(side=tk.LEFT, padx=10, pady=5)
        
        # 來源語言
        tk.Label(
            lang_frame,
            text="來源語言:",
            bg='#2d2d2d',
            fg='white',
            font=('Arial', 10)
        ).pack()
        
        self.source_lang_var = tk.StringVar(value=self.settings['source_language'])
        self.source_lang_combo = ttk.Combobox(
            lang_frame,
            textvariable=self.source_lang_var,
            values=[(f"{info['name']} ({code})") for code, info in self.installed_languages.items()],
            width=15,
            state='readonly'
        )
        self.source_lang_combo.pack(pady=5)
        
        # 箭頭
        tk.Label(
            control_frame,
            text="→",
            bg='#2d2d2d',
            fg='white',
            font=('Arial', 20)
        ).pack(side=tk.LEFT, padx=10)
        
        # 目標語言
        target_frame = tk.Frame(control_frame, bg='#2d2d2d')
        target_frame.pack(side=tk.LEFT, padx=10, pady=5)
        
        tk.Label(
            target_frame,
            text="目標語言:",
            bg='#2d2d2d',
            fg='white',
            font=('Arial', 10)
        ).pack()
        
        self.target_lang_var = tk.StringVar(value=self.settings['target_language'])
        self.target_lang_combo = ttk.Combobox(
            target_frame,
            textvariable=self.target_lang_var,
            values=list(TARGET_LANGUAGES.values()),
            width=15,
            state='readonly'
        )
        self.target_lang_combo.pack(pady=5)
        
        # 按鈕區域
        button_frame = tk.Frame(control_frame, bg='#2d2d2d')
        button_frame.pack(side=tk.RIGHT, padx=10, pady=5)
        
        buttons_data = [
            ("選擇區域", self.select_capture_region, '#4CAF50'),
            ("截圖翻譯", self.screenshot_translate, '#9C27B0'),
            ("開始偵測", self.toggle_capture, '#2196F3'),
            ("清除", self.clear_current, '#f44336'),
        ]
        
        for text, command, color in buttons_data:
            btn = tk.Button(
                button_frame,
                text=text,
                command=command,
                bg=color,
                fg='white',
                font=('Arial', 10, 'bold'),
                padx=15,
                pady=8,
                relief=tk.FLAT,
                cursor='hand2'
            )
            btn.pack(side=tk.LEFT, padx=5)
            
        # 自動偵測語言
        self.auto_detect_var = tk.BooleanVar(value=self.settings['auto_detect'])
        auto_detect_cb = tk.Checkbutton(
            control_frame,
            text="自動偵測語言",
            variable=self.auto_detect_var,
            bg='#2d2d2d',
            fg='white',
            selectcolor='#2d2d2d',
            activebackground='#2d2d2d',
            activeforeground='white',
            command=self.toggle_auto_detect
        )
        auto_detect_cb.pack(side=tk.LEFT, padx=20)
        
        # 顯示區域
        display_frame = tk.Frame(parent, bg='#1e1e1e')
        display_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 左側：擷取預覽
        left_frame = tk.Frame(display_frame, bg='#1e1e1e', width=400)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        left_frame.pack_propagate(False)
        
        preview_label = tk.Label(
            left_frame,
            text="擷取預覽",
            bg='#1e1e1e',
            fg='white',
            font=('Arial', 10, 'bold')
        )
        preview_label.pack()
        
        self.preview_canvas = tk.Canvas(
            left_frame,
            bg='#0d0d0d',
            highlightthickness=1,
            highlightbackground='#444'
        )
        self.preview_canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 右側：翻譯結果
        right_frame = tk.Frame(display_frame, bg='#1e1e1e')
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # 即時翻譯顯示
        instant_frame = tk.LabelFrame(
            right_frame,
            text="即時翻譯",
            bg='#1e1e1e',
            fg='white',
            font=('Arial', 10, 'bold')
        )
        instant_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 原文顯示
        source_frame = tk.Frame(instant_frame, bg='#1e1e1e')
        source_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.source_label = tk.Label(
            source_frame,
            text="原文 (未知):",
            bg='#1e1e1e',
            fg='#888',
            font=('Arial', 9)
        )
        self.source_label.pack(anchor=tk.W)
        
        self.instant_source = tk.Text(
            source_frame,
            height=6,
            bg='#0d0d0d',
            fg='#64B5F6',
            font=('Arial', 12),
            wrap=tk.WORD
        )
        self.instant_source.pack(fill=tk.X)
        
        # 譯文顯示
        target_frame = tk.Frame(instant_frame, bg='#1e1e1e')
        target_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.target_label = tk.Label(
            target_frame,
            text="譯文:",
            bg='#1e1e1e',
            fg='#888',
            font=('Arial', 9)
        )
        self.target_label.pack(anchor=tk.W)
        
        self.instant_target = tk.Text(
            target_frame,
            height=8,
            bg='#0d0d0d',
            fg='#81C784',
            font=('Microsoft JhengHei', 12),
            wrap=tk.WORD
        )
        self.instant_target.pack(fill=tk.X)
        
        # 信心度顯示
        self.confidence_label = tk.Label(
            instant_frame,
            text="識別信心度: --",
            bg='#1e1e1e',
            fg='#FFB74D',
            font=('Arial', 9)
        )
        self.confidence_label.pack(pady=5)
        
        # 翻譯歷史
        history_frame = tk.LabelFrame(
            right_frame,
            text="翻譯記錄",
            bg='#1e1e1e',
            fg='white',
            font=('Arial', 10, 'bold')
        )
        history_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.translation_display = scrolledtext.ScrolledText(
            history_frame,
            bg='#0d0d0d',
            fg='white',
            font=('Microsoft JhengHei', 10),
            wrap=tk.WORD
        )
        self.translation_display.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 設定文字標籤
        self.translation_display.tag_configure('time', foreground='#FFB74D')
        self.translation_display.tag_configure('lang', foreground='#B39DDB')
        self.translation_display.tag_configure('source', foreground='#64B5F6')
        self.translation_display.tag_configure('target', foreground='#81C784')
        self.translation_display.tag_configure('confidence', foreground='#FFA726')
        
    def create_language_tab(self, parent):
        """建立語言設定頁面"""
        main_frame = tk.Frame(parent, bg='#1e1e1e')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 已安裝語言
        installed_frame = tk.LabelFrame(
            main_frame,
            text="已安裝的語言包",
            bg='#1e1e1e',
            fg='white',
            font=('Arial', 11, 'bold')
        )
        installed_frame.pack(fill=tk.X, pady=10)
        
        # 顯示已安裝語言
        lang_text = tk.Text(
            installed_frame,
            height=10,
            bg='#0d0d0d',
            fg='white',
            font=('Arial', 10),
            wrap=tk.WORD
        )
        lang_text.pack(fill=tk.X, padx=10, pady=10)
        
        # 填入已安裝語言資訊
        lang_text.insert(tk.END, "已安裝的 OCR 語言包:\n\n")
        for code, info in self.installed_languages.items():
            lang_text.insert(tk.END, f"• {info['name']} ({code})\n")
            
        lang_text.insert(tk.END, "\n" + "="*50 + "\n")
        lang_text.insert(tk.END, "支援的翻譯目標語言:\n\n")
        for code, name in TARGET_LANGUAGES.items():
            lang_text.insert(tk.END, f"• {name} ({code})\n")
            
        lang_text.config(state=tk.DISABLED)
        
        # 語言對應設定
        mapping_frame = tk.LabelFrame(
            main_frame,
            text="常用語言組合",
            bg='#1e1e1e',
            fg='white',
            font=('Arial', 11, 'bold')
        )
        mapping_frame.pack(fill=tk.X, pady=10)
        
        # 預設組合按鈕
        presets = [
            ("日文→繁中", "jpn", "zh-tw"),
            ("韓文→繁中", "kor", "zh-tw"),
            ("英文→繁中", "eng", "zh-tw"),
            ("簡中→繁中", "chi_sim", "zh-tw"),
            ("日文→英文", "jpn", "en"),
            ("多語言模式", "multi", "zh-tw"),
        ]
        
        preset_frame = tk.Frame(mapping_frame, bg='#1e1e1e')
        preset_frame.pack(pady=10)
        
        for i, (label, source, target) in enumerate(presets):
            btn = tk.Button(
                preset_frame,
                text=label,
                command=lambda s=source, t=target: self.set_language_preset(s, t),
                bg='#3f51b5',
                fg='white',
                font=('Arial', 10),
                padx=15,
                pady=5,
                relief=tk.FLAT
            )
            btn.grid(row=i//3, column=i%3, padx=5, pady=5)
            
        # 安裝指南
        guide_frame = tk.LabelFrame(
            main_frame,
            text="安裝更多語言",
            bg='#1e1e1e',
            fg='white',
            font=('Arial', 11, 'bold')
        )
        guide_frame.pack(fill=tk.X, pady=10)
        
        guide_text = tk.Label(
            guide_frame,
            text="如需支援更多語言，請重新執行 Tesseract 安裝程式\n"
                 "並在「Additional language data」選項中勾選需要的語言包\n\n"
                 "常用遊戲語言建議:\n"
                 "• 日文遊戲: jpn, jpn_vert (直書)\n"
                 "• 韓文遊戲: kor, kor_vert (直書)\n"
                 "• 歐美遊戲: eng, fra, deu, spa, ita\n"
                 "• 中文遊戲: chi_sim (簡體), chi_tra (繁體)",
            bg='#1e1e1e',
            fg='#ccc',
            font=('Arial', 10),
            justify=tk.LEFT
        )
        guide_text.pack(padx=10, pady=10)
        
    def create_settings_tab(self, parent):
        """建立進階設定頁面"""
        settings_frame = tk.Frame(parent, bg='#1e1e1e')
        settings_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # OCR 設定
        ocr_frame = tk.LabelFrame(
            settings_frame,
            text="OCR 設定",
            bg='#1e1e1e',
            fg='white',
            font=('Arial', 11, 'bold')
        )
        ocr_frame.pack(fill=tk.X, pady=10)
        
        # OCR 模式
        mode_frame = tk.Frame(ocr_frame, bg='#1e1e1e')
        mode_frame.pack(pady=10)
        
        tk.Label(
            mode_frame,
            text="識別模式:",
            bg='#1e1e1e',
            fg='white'
        ).pack(side=tk.LEFT, padx=10)
        
        self.ocr_mode_var = tk.StringVar(value=self.settings['ocr_mode'])
        
        tk.Radiobutton(
            mode_frame,
            text="單一語言 (快速)",
            variable=self.ocr_mode_var,
            value='single',
            bg='#1e1e1e',
            fg='white',
            selectcolor='#1e1e1e'
        ).pack(side=tk.LEFT, padx=10)
        
        tk.Radiobutton(
            mode_frame,
            text="多語言混合 (慢)",
            variable=self.ocr_mode_var,
            value='multi',
            bg='#1e1e1e',
            fg='white',
            selectcolor='#1e1e1e'
        ).pack(side=tk.LEFT, padx=10)
        
        # 預處理選項
        self.preprocessing_var = tk.BooleanVar(value=self.settings['preprocessing'])
        tk.Checkbutton(
            ocr_frame,
            text="啟用影像預處理 (提高識別率)",
            variable=self.preprocessing_var,
            bg='#1e1e1e',
            fg='white',
            selectcolor='#1e1e1e'
        ).pack(pady=5)
        
        # 信心度門檻
        threshold_frame = tk.Frame(ocr_frame, bg='#1e1e1e')
        threshold_frame.pack(pady=10)
        
        tk.Label(
            threshold_frame,
            text="最低信心度門檻:",
            bg='#1e1e1e',
            fg='white'
        ).pack(side=tk.LEFT, padx=10)
        
        self.confidence_var = tk.IntVar(value=self.settings['confidence_threshold'])
        confidence_scale = tk.Scale(
            threshold_frame,
            from_=0,
            to=100,
            orient=tk.HORIZONTAL,
            variable=self.confidence_var,
            bg='#1e1e1e',
            fg='white',
            highlightthickness=0,
            troughcolor='#444',
            length=200
        )
        confidence_scale.pack(side=tk.LEFT)
        
        self.confidence_value_label = tk.Label(
            threshold_frame,
            text=f"{self.confidence_var.get()}%",
            bg='#1e1e1e',
            fg='white'
        )
        self.confidence_value_label.pack(side=tk.LEFT, padx=10)
        
        confidence_scale.config(command=lambda v: self.confidence_value_label.config(
            text=f"{int(float(v))}%"
        ))
        
        # 效能設定
        perf_frame = tk.LabelFrame(
            settings_frame,
            text="效能設定",
            bg='#1e1e1e',
            fg='white',
            font=('Arial', 11, 'bold')
        )
        perf_frame.pack(fill=tk.X, pady=10)
        
        # 更新間隔
        interval_frame = tk.Frame(perf_frame, bg='#1e1e1e')
        interval_frame.pack(pady=10)
        
        tk.Label(
            interval_frame,
            text="擷取間隔 (秒):",
            bg='#1e1e1e',
            fg='white'
        ).pack(side=tk.LEFT, padx=10)
        
        self.interval_var = tk.DoubleVar(value=self.settings['update_interval'])
        interval_scale = tk.Scale(
            interval_frame,
            from_=0.1,
            to=3.0,
            resolution=0.1,
            orient=tk.HORIZONTAL,
            variable=self.interval_var,
            bg='#1e1e1e',
            fg='white',
            highlightthickness=0,
            troughcolor='#444',
            length=200
        )
        interval_scale.pack(side=tk.LEFT)
        
        # 儲存按鈕
        tk.Button(
            settings_frame,
            text="儲存設定",
            command=self.save_settings,
            bg='#4CAF50',
            fg='white',
            font=('Arial', 10, 'bold'),
            padx=20,
            pady=10,
            relief=tk.FLAT
        ).pack(pady=20)
        
    def create_history_tab(self, parent):
        """建立歷史記錄頁面"""
        history_frame = tk.Frame(parent, bg='#1e1e1e')
        history_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 統計資訊
        stats_frame = tk.Frame(history_frame, bg='#2d2d2d', height=60)
        stats_frame.pack(fill=tk.X, pady=(0, 10))
        stats_frame.pack_propagate(False)
        
        self.stats_label = tk.Label(
            stats_frame,
            text="總翻譯數: 0 | 今日: 0 | 語言分布: --",
            bg='#2d2d2d',
            fg='white',
            font=('Arial', 10)
        )
        self.stats_label.pack(expand=True)
        
        # 工具列
        toolbar = tk.Frame(history_frame, bg='#2d2d2d')
        toolbar.pack(fill=tk.X, pady=(0, 10))
        
        tk.Button(
            toolbar,
            text="匯出歷史",
            command=self.export_history,
            bg='#4CAF50',
            fg='white',
            font=('Arial', 9),
            padx=15,
            pady=5
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            toolbar,
            text="清空歷史",
            command=self.clear_history,
            bg='#f44336',
            fg='white',
            font=('Arial', 9),
            padx=15,
            pady=5
        ).pack(side=tk.LEFT, padx=5)
        
        # 篩選器
        filter_frame = tk.Frame(toolbar, bg='#2d2d2d')
        filter_frame.pack(side=tk.RIGHT, padx=10)
        
        tk.Label(
            filter_frame,
            text="篩選語言:",
            bg='#2d2d2d',
            fg='white'
        ).pack(side=tk.LEFT, padx=5)
        
        self.filter_var = tk.StringVar(value="全部")
        filter_combo = ttk.Combobox(
            filter_frame,
            textvariable=self.filter_var,
            values=["全部"] + [info['name'] for info in LANGUAGES.values()],
            width=12,
            state='readonly'
        )
        filter_combo.pack(side=tk.LEFT)
        filter_combo.bind('<<ComboboxSelected>>', self.filter_history)
        
        # 歷史列表
        list_frame = tk.Frame(history_frame, bg='#1e1e1e')
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # 捲軸
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.history_listbox = tk.Listbox(
            list_frame,
            bg='#0d0d0d',
            fg='white',
            selectbackground='#444',
            font=('Microsoft JhengHei', 10),
            yscrollcommand=scrollbar.set
        )
        self.history_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.history_listbox.yview)
        
        # 雙擊顯示詳情
        self.history_listbox.bind('<Double-Button-1>', self.show_history_detail)
        
    def create_status_bar(self):
        """建立狀態列"""
        status_frame = tk.Frame(self.root, bg='#0d0d0d', height=30)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        self.status_label = tk.Label(
            status_frame,
            text="準備就緒",
            bg='#0d0d0d',
            fg='#4CAF50',
            font=('Arial', 9)
        )
        self.status_label.pack(side=tk.LEFT, padx=10)
        
        self.region_label = tk.Label(
            status_frame,
            text="未選擇區域",
            bg='#0d0d0d',
            fg='#888',
            font=('Arial', 9)
        )
        self.region_label.pack(side=tk.RIGHT, padx=10)
        
        self.current_lang_label = tk.Label(
            status_frame,
            text="--",
            bg='#0d0d0d',
            fg='#B39DDB',
            font=('Arial', 9)
        )
        self.current_lang_label.pack(side=tk.RIGHT, padx=20)
        
    def setup_hotkeys(self):
        """設定快捷鍵"""
        keyboard.add_hotkey('f2', self.select_capture_region)
        keyboard.add_hotkey('f3', self.toggle_capture)
        keyboard.add_hotkey('f5', self.quick_switch_language)
        keyboard.add_hotkey('ctrl+s', self.save_current_session)
        
    def toggle_auto_detect(self):
        """切換自動偵測語言"""
        self.settings['auto_detect'] = self.auto_detect_var.get()
        if self.auto_detect_var.get():
            self.source_lang_combo.config(state='disabled')
            self.status_label.config(text="已啟用自動語言偵測", fg='#4CAF50')
        else:
            self.source_lang_combo.config(state='readonly')
            self.status_label.config(text="已關閉自動語言偵測", fg='#FFC107')
            
    def set_language_preset(self, source, target):
        """設定語言預設組合"""
        if source == "multi":
            self.ocr_mode_var.set("multi")
            self.auto_detect_var.set(True)
            self.toggle_auto_detect()
        else:
            self.source_lang_var.set(source)
            self.ocr_mode_var.set("single")
            
        self.target_lang_var.set(target)
        self.update_language_display()
        self.status_label.config(text=f"已切換語言組合", fg='#4CAF50')
        
    def quick_switch_language(self):
        """快速切換常用語言 (F5)"""
        # 定義快速切換順序
        quick_langs = ['jpn', 'kor', 'eng', 'chi_sim']
        current = self.source_lang_var.get()
        
        try:
            current_index = quick_langs.index(current)
            next_index = (current_index + 1) % len(quick_langs)
            self.source_lang_var.set(quick_langs[next_index])
            self.update_language_display()
        except ValueError:
            self.source_lang_var.set(quick_langs[0])
            
    def update_language_display(self):
        """更新語言顯示"""
        source_code = self.source_lang_var.get()
        target_code = self.get_target_code()
        
        if source_code in self.installed_languages:
            source_name = self.installed_languages[source_code]['name']
            self.source_label.config(text=f"原文 ({source_name}):")
            
        target_name = TARGET_LANGUAGES.get(target_code, '未知')
        self.target_label.config(text=f"譯文 ({target_name}):")
        
        self.current_lang_label.config(
            text=f"{source_name}→{target_name}" if source_code in self.installed_languages else "--"
        )
        
    def get_target_code(self):
        """取得目標語言代碼"""
        target_name = self.target_lang_var.get()
        for code, name in TARGET_LANGUAGES.items():
            if name == target_name:
                return code
        return 'zh-tw'
        
    def select_capture_region(self):
        """選擇螢幕擷取區域"""
        self.status_label.config(text="選擇擷取區域中...", fg='#FFC107')
        self.root.withdraw()
        
        # 建立全螢幕選擇視窗
        selection_window = tk.Toplevel()
        selection_window.attributes('-fullscreen', True)
        selection_window.attributes('-alpha', 0.3)
        selection_window.attributes('-topmost', True)
        selection_window.configure(bg='black')
        
        # 擷取螢幕
        screenshot = pyautogui.screenshot()
        
        # 建立畫布
        canvas = tk.Canvas(
            selection_window,
            highlightthickness=0,
            cursor="cross"
        )
        canvas.pack(fill=tk.BOTH, expand=True)
        
        # 顯示螢幕截圖
        photo = ImageTk.PhotoImage(screenshot)
        canvas.create_image(0, 0, anchor=tk.NW, image=photo)
        
        # 選擇變數
        self.selection_coords = {}
        
        def on_mouse_down(event):
            self.selection_coords['start'] = (event.x, event.y)
            self.selection_coords['rect'] = canvas.create_rectangle(
                event.x, event.y, event.x, event.y,
                outline='red', width=2, fill='red', stipple='gray50'
            )
            
        def on_mouse_move(event):
            if 'rect' in self.selection_coords:
                x1, y1 = self.selection_coords['start']
                canvas.coords(self.selection_coords['rect'], x1, y1, event.x, event.y)
                
                # 顯示座標資訊
                if hasattr(self, 'coord_label'):
                    self.coord_label.destroy()
                self.coord_label = tk.Label(
                    selection_window,
                    text=f"區域: {abs(event.x-x1)}x{abs(event.y-y1)}",
                    bg='black',
                    fg='white',
                    font=('Arial', 12)
                )
                self.coord_label.place(x=event.x+10, y=event.y+10)
                
        def on_mouse_up(event):
            if 'start' in self.selection_coords:
                x1, y1 = self.selection_coords['start']
                x2, y2 = event.x, event.y
                
                # 確保座標正確
                x1, x2 = min(x1, x2), max(x1, x2)
                y1, y2 = min(y1, y2), max(y1, y2)
                
                if x2 - x1 > 10 and y2 - y1 > 10:
                    self.capture_region = (x1, y1, x2 - x1, y2 - y1)
                    self.region_label.config(
                        text=f"區域: {x2-x1}x{y2-y1} @ ({x1},{y1})",
                        fg='white'
                    )
                    self.status_label.config(text="已選擇區域", fg='#4CAF50')
                else:
                    self.status_label.config(text="選擇區域太小", fg='#f44336')
                    
            selection_window.destroy()
            self.root.deiconify()
            
        # 綁定事件
        canvas.bind('<Button-1>', on_mouse_down)
        canvas.bind('<B1-Motion>', on_mouse_move)
        canvas.bind('<ButtonRelease-1>', on_mouse_up)
        canvas.bind('<Escape>', lambda e: (selection_window.destroy(), self.root.deiconify()))
        
        # 提示文字
        hint_label = tk.Label(
            selection_window,
            text="拖曳滑鼠選擇區域 | ESC 取消 | 選擇遊戲文字區域",
            bg='black',
            fg='white',
            font=('Arial', 14, 'bold')
        )
        hint_label.place(relx=0.5, y=30, anchor=tk.CENTER)
        
    def toggle_capture(self):
        """開始/停止擷取"""
        if not self.capture_region:
            messagebox.showwarning("提示", "請先選擇擷取區域！")
            return
            
        self.is_capturing = not self.is_capturing
        
        if self.is_capturing:
            # 更新按鈕
            for widget in self.root.winfo_children():
                for child in widget.winfo_children():
                    if isinstance(child, tk.Button) and child['text'] == '開始偵測':
                        child.config(text='停止偵測', bg='#f44336')
                        
            self.status_label.config(text="偵測中...", fg='#4CAF50')
            self.update_language_display()
            
            # 啟動擷取執行緒
            self.capture_thread = threading.Thread(target=self.capture_loop, daemon=True)
            self.capture_thread.start()
        else:
            # 更新按鈕
            for widget in self.root.winfo_children():
                for child in widget.winfo_children():
                    if isinstance(child, tk.Button) and child['text'] == '停止偵測':
                        child.config(text='開始偵測', bg='#2196F3')
                        
            self.status_label.config(text="已停止", fg='#FFC107')
            
    def capture_loop(self):
        """擷取循環"""
        last_text_hash = ""
        
        while self.is_capturing:
            try:
                # 擷取指定區域
                x, y, w, h = self.capture_region
                screenshot = pyautogui.screenshot(region=(x, y, w, h))
                
                # 更新預覽
                self.root.after(0, self.update_preview, screenshot)
                
                # 預處理影像
                if self.preprocessing_var.get():
                    processed_img = self.preprocess_image(screenshot)
                else:
                    processed_img = np.array(screenshot)
                    
                # OCR 識別
                if self.auto_detect_var.get() or self.ocr_mode_var.get() == 'multi':
                    # 多語言模式
                    result = self.multi_language_ocr(processed_img)
                else:
                    # 單一語言模式
                    result = self.single_language_ocr(processed_img)
                    
                if result and result['text']:
                    current_hash = hash(result['text'])
                    if current_hash != last_text_hash:
                        last_text_hash = current_hash
                        
                        # 檢查信心度
                        if result['confidence'] >= self.confidence_var.get():
                            # 翻譯
                            translation = self.translate_text(
                                result['text'],
                                result['language']
                            )
                            
                            # 更新顯示
                            self.root.after(0, self.update_translation,
                                          result['text'],
                                          translation,
                                          result['language'],
                                          result['confidence'])
                        else:
                            self.root.after(0, self.update_confidence,
                                          result['confidence'])
                            
            except Exception as e:
                print(f"擷取錯誤: {e}")
                
            time.sleep(self.interval_var.get())
            
    def preprocess_image(self, image):
        """影像預處理"""
        # 轉換為 numpy 陣列
        img = np.array(image)
        
        # 轉換為灰階
        gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
        
        # 提高對比度
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
        enhanced = clahe.apply(gray)
        
        # 去噪
        denoised = cv2.fastNlMeansDenoising(enhanced)
        
        # 放大
        scaled = cv2.resize(denoised, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
        
        # 銳化
        kernel = np.array([[-1,-1,-1],
                          [-1, 9,-1],
                          [-1,-1,-1]])
        sharpened = cv2.filter2D(scaled, -1, kernel)
        
        # 二值化
        _, binary = cv2.threshold(sharpened, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        
        return binary
        
    def single_language_ocr(self, image):
        """單一語言 OCR"""
        lang_code = self.source_lang_var.get()
        
        try:
            # 使用 Tesseract 進行 OCR
            data = pytesseract.image_to_data(
                image,
                lang=lang_code,
                output_type=pytesseract.Output.DICT,
                config='--psm 6'
            )
            
            # 計算平均信心度
            confidences = [int(conf) for conf in data['conf'] if int(conf) > 0]
            avg_confidence = sum(confidences) / len(confidences) if confidences else 0
            
            # 組合文字
            text = ' '.join([data['text'][i] for i in range(len(data['text'])) 
                           if int(data['conf'][i]) > 0])
            
            return {
                'text': text.strip(),
                'language': lang_code,
                'confidence': avg_confidence
            }
        except Exception as e:
            print(f"OCR 錯誤: {e}")
            return None
            
    def multi_language_ocr(self, image):
        """多語言 OCR (自動偵測)"""
        best_result = None
        best_confidence = 0
        
        # 嘗試所有已安裝的語言
        for lang_code in self.installed_languages.keys():
            try:
                data = pytesseract.image_to_data(
                    image,
                    lang=lang_code,
                    output_type=pytesseract.Output.DICT,
                    config='--psm 6'
                )
                
                # 計算信心度
                confidences = [int(conf) for conf in data['conf'] if int(conf) > 0]
                avg_confidence = sum(confidences) / len(confidences) if confidences else 0
                
                # 如果信心度更高，更新最佳結果
                if avg_confidence > best_confidence:
                    text = ' '.join([data['text'][i] for i in range(len(data['text'])) 
                                   if int(data['conf'][i]) > 0])
                    
                    if text.strip():
                        best_confidence = avg_confidence
                        best_result = {
                            'text': text.strip(),
                            'language': lang_code,
                            'confidence': avg_confidence
                        }
                        
            except:
                continue
                
        return best_result
        
    def translate_text(self, text, source_lang_code):
        """翻譯文字"""
        try:
            # 取得語言代碼
            source_google = self.installed_languages[source_lang_code]['google_code']
            target_google = self.get_target_code()
            
            # 檢查快取
            cache_key = f"{text}_{source_google}_{target_google}"
            if cache_key in self.translation_cache:
                return self.translation_cache[cache_key]
                
            # 執行翻譯
            result = self.translator.translate(
                text,
                src=source_google,
                dest=target_google
            )
            
            # 儲存到快取
            self.translation_cache[cache_key] = result.text
            
            return result.text
        except Exception as e:
            return f"翻譯錯誤: {str(e)}"
            
    def update_preview(self, screenshot):
        """更新預覽圖片"""
        # 調整大小以適應畫布
        canvas_width = self.preview_canvas.winfo_width()
        canvas_height = self.preview_canvas.winfo_height()
        
        if canvas_width > 1 and canvas_height > 1:
            # 計算縮放比例
            img_width, img_height = screenshot.size
            scale = min(canvas_width / img_width, canvas_height / img_height, 1.0)
            
            # 調整圖片大小
            new_width = int(img_width * scale)
            new_height = int(img_height * scale)
            resized = screenshot.resize((new_width, new_height), Image.Resampling.LANCZOS)
            
            # 轉換為 PhotoImage
            photo = ImageTk.PhotoImage(resized)
            
            # 清除畫布並顯示圖片
            self.preview_canvas.delete("all")
            x = (canvas_width - new_width) // 2
            y = (canvas_height - new_height) // 2
            self.preview_canvas.create_image(x, y, anchor=tk.NW, image=photo)
            self.preview_canvas.image = photo  # 保持引用
            
    def update_translation(self, source_text, target_text, language, confidence):
        """更新翻譯顯示"""
        # 更新即時顯示
        self.instant_source.delete(1.0, tk.END)
        self.instant_source.insert(1.0, source_text)
        
        self.instant_target.delete(1.0, tk.END)
        self.instant_target.insert(1.0, target_text)
        
        # 更新語言和信心度
        lang_name = self.installed_languages[language]['name']
        self.source_label.config(text=f"原文 ({lang_name}):")
        self.confidence_label.config(text=f"識別信心度: {confidence:.1f}%")
        
        # 加入歷史記錄
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # 更新翻譯顯示區
        self.translation_display.insert(tk.END, f"[{timestamp}] ", 'time')
        self.translation_display.insert(tk.END, f"{lang_name} ", 'lang')
        self.translation_display.insert(tk.END, f"(信心度: {confidence:.1f}%)\n", 'confidence')
        self.translation_display.insert(tk.END, f"原文: {source_text}\n", 'source')
        self.translation_display.insert(tk.END, f"譯文: {target_text}\n", 'target')
        self.translation_display.insert(tk.END, "-" * 70 + "\n")
        self.translation_display.see(tk.END)
        
        # 儲存到歷史
        history_item = {
            'timestamp': timestamp,
            'date': datetime.now().strftime("%Y-%m-%d"),
            'language': language,
            'language_name': lang_name,
            'source': source_text,
            'target': target_text,
            'confidence': confidence
        }
        self.translation_history.append(history_item)
        
        # 更新歷史列表
        self.history_listbox.insert(
            tk.END,
            f"{timestamp} [{lang_name}] {source_text[:30]}..."
        )
        
        # 更新統計
        self.update_statistics()
        
    def update_confidence(self, confidence):
        """只更新信心度（文字未通過門檻）"""
        self.confidence_label.config(
            text=f"識別信心度: {confidence:.1f}% (低於門檻)",
            fg='#ff9800'
        )
        
    def update_statistics(self):
        """更新統計資訊"""
        total = len(self.translation_history)
        today = sum(1 for item in self.translation_history 
                   if item['date'] == datetime.now().strftime("%Y-%m-%d"))
        
        # 語言分布
        lang_count = {}
        for item in self.translation_history:
            lang = item['language_name']
            lang_count[lang] = lang_count.get(lang, 0) + 1
            
        # 找出最常用的語言
        if lang_count:
            top_lang = max(lang_count.items(), key=lambda x: x[1])
            lang_dist = f"{top_lang[0]} ({top_lang[1]}次)"
        else:
            lang_dist = "--"
            
        self.stats_label.config(
            text=f"總翻譯數: {total} | 今日: {today} | 最常: {lang_dist}"
        )
        
    def clear_current(self):
        """清除當前顯示"""
        self.instant_source.delete(1.0, tk.END)
        self.instant_target.delete(1.0, tk.END)
        self.translation_display.delete(1.0, tk.END)
        self.confidence_label.config(text="識別信心度: --")
        
    def filter_history(self, event=None):
        """篩選歷史記錄"""
        filter_lang = self.filter_var.get()
        self.history_listbox.delete(0, tk.END)
        
        for item in self.translation_history:
            if filter_lang == "全部" or item['language_name'] == filter_lang:
                display_text = f"{item['timestamp']} [{item['language_name']}] {item['source'][:30]}..."
                self.history_listbox.insert(tk.END, display_text)
                
    def show_history_detail(self, event):
        """顯示歷史詳情"""
        selection = self.history_listbox.curselection()
        if selection:
            # 找出對應的歷史項目
            display_text = self.history_listbox.get(selection[0])
            timestamp = display_text.split(' ')[0]
            
            for item in self.translation_history:
                if item['timestamp'] == timestamp:
                    # 建立詳情視窗
                    detail_window = tk.Toplevel(self.root)
                    detail_window.title("翻譯詳情")
                    detail_window.geometry("600x400")
                    detail_window.configure(bg='#1e1e1e')
                    
                    # 顯示詳細內容
                    detail_text = tk.Text(
                        detail_window,
                        bg='#0d0d0d',
                        fg='white',
                        font=('Microsoft JhengHei', 11),
                        wrap=tk.WORD
                    )
                    detail_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
                    
                    detail_text.insert(tk.END, f"時間: {item['date']} {item['timestamp']}\n")
                    detail_text.insert(tk.END, f"語言: {item['language_name']} ({item['language']})\n")
                    detail_text.insert(tk.END, f"信心度: {item['confidence']:.1f}%\n\n")
                    detail_text.insert(tk.END, f"原文:\n{item['source']}\n\n")
                    detail_text.insert(tk.END, f"譯文:\n{item['target']}")
                    detail_text.config(state=tk.DISABLED)
                    
                    # 複製按鈕
                    tk.Button(
                        detail_window,
                        text="複製譯文",
                        command=lambda: self.copy_to_clipboard(item['target']),
                        bg='#4CAF50',
                        fg='white',
                        font=('Arial', 10)
                    ).pack(pady=10)
                    
                    break
                    
    def copy_to_clipboard(self, text):
        """複製到剪貼簿"""
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.status_label.config(text="已複製到剪貼簿", fg='#4CAF50')
        
    def export_history(self):
        """匯出歷史記錄"""
        if not self.translation_history:
            messagebox.showinfo("提示", "沒有可匯出的歷史記錄")
            return
            
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[
                ("JSON files", "*.json"),
                ("Text files", "*.txt"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ],
            initialfile=f"translation_history_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        )
        
        if filename:
            try:
                if filename.endswith('.json'):
                    with open(filename, 'w', encoding='utf-8') as f:
                        json.dump(list(self.translation_history), f, ensure_ascii=False, indent=2)
                elif filename.endswith('.csv'):
                    import csv
                    with open(filename, 'w', encoding='utf-8-sig', newline='') as f:
                        writer = csv.DictWriter(f, fieldnames=['date', 'timestamp', 'language_name', 
                                                              'source', 'target', 'confidence'])
                        writer.writeheader()
                        writer.writerows(self.translation_history)
                else:  # txt
                    with open(filename, 'w', encoding='utf-8') as f:
                        for item in self.translation_history:
                            f.write(f"[{item['date']} {item['timestamp']}]\n")
                            f.write(f"語言: {item['language_name']} (信心度: {item['confidence']:.1f}%)\n")
                            f.write(f"原文: {item['source']}\n")
                            f.write(f"譯文: {item['target']}\n")
                            f.write("-" * 70 + "\n\n")
                            
                messagebox.showinfo("成功", "歷史記錄已匯出")
            except Exception as e:
                messagebox.showerror("錯誤", f"匯出失敗: {str(e)}")
                
    def clear_history(self):
        """清空歷史記錄"""
        if self.translation_history:
            if messagebox.askyesno("確認", "確定要清空所有歷史記錄嗎？"):
                self.translation_history.clear()
                self.history_listbox.delete(0, tk.END)
                self.translation_display.delete(1.0, tk.END)
                self.update_statistics()
                self.status_label.config(text="已清空歷史記錄", fg='#4CAF50')
                
    def save_current_session(self):
        """儲存當前工作階段"""
        if not self.translation_history:
            self.status_label.config(text="沒有可儲存的記錄", fg='#FFC107')
            return
            
        filename = f"session_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(list(self.translation_history), f, ensure_ascii=False, indent=2)
            self.status_label.config(text=f"已儲存: {filename}", fg='#4CAF50')
        except Exception as e:
            self.status_label.config(text=f"儲存失敗: {str(e)}", fg='#f44336')
            
    def save_settings(self):
        """儲存設定"""
        self.settings['source_language'] = self.source_lang_var.get()
        self.settings['target_language'] = self.get_target_code()
        self.settings['ocr_mode'] = self.ocr_mode_var.get()
        self.settings['preprocessing'] = self.preprocessing_var.get()
        self.settings['update_interval'] = self.interval_var.get()
        self.settings['confidence_threshold'] = self.confidence_var.get()
        self.settings['auto_detect'] = self.auto_detect_var.get()
        
        try:
            with open('translator_settings.json', 'w') as f:
                json.dump(self.settings, f, indent=2)
            messagebox.showinfo("成功", "設定已儲存")
        except Exception as e:
            messagebox.showerror("錯誤", f"儲存設定失敗: {str(e)}")
            
    def load_settings(self):
        """載入設定"""
        try:
            if os.path.exists('translator_settings.json'):
                with open('translator_settings.json', 'r') as f:
                    saved_settings = json.load(f)
                    self.settings.update(saved_settings)
        except Exception as e:
            print(f"載入設定失敗: {e}")
            
    def on_closing(self):
        """關閉程式時的處理"""
        self.is_capturing = False
        self.save_settings()
        self.root.destroy()
        """關閉程式時的處理"""
        self.is_capturing = False
        self.save_settings()
        self.root.destroy()

def main():
    """主程式入口"""
    # 檢查系統需求
    if os.name == 'nt':  # Windows
        # 設定 DPI 感知
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass
        
        # 檢查 Tesseract
        tesseract_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        if os.path.exists(tesseract_path):
            pytesseract.pytesseract.tesseract_cmd = tesseract_path
        else:
            # 嘗試其他常見路徑
            alt_paths = [
                r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
                r'C:\Tesseract-OCR\tesseract.exe',
                r'C:\tesseract\tesseract.exe'
            ]
            
            found = False
            for path in alt_paths:
                if os.path.exists(path):
                    pytesseract.pytesseract.tesseract_cmd = path
                    found = True
                    break
                    
            if not found:
                response = messagebox.askyesno(
                    "Tesseract OCR 未找到",
                    "未偵測到 Tesseract OCR。\n\n"
                    "請先安裝 Tesseract OCR:\n"
                    "https://github.com/UB-Mannheim/tesseract/wiki\n\n"
                    "安裝時請記得勾選需要的語言包。\n\n"
                    "是否仍要繼續執行程式？"
                )
                if not response:
                    return
    
    # 建立並執行應用程式
    root = tk.Tk()
    app = MultilingualGameTranslator(root)
    
    # 設定關閉事件
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    # 顯示啟動提示
    app.status_label.config(
        text="歡迎使用多語言遊戲翻譯器！請按 F2 選擇區域",
        fg='#4CAF50'
    )
    
    # 執行主循環
    root.mainloop()

if __name__ == "__main__":
    main()