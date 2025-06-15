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
import easyocr
import ctypes

class OverlayWindow:
    """透明覆蓋視窗，用於在遊戲上顯示翻譯"""
    def __init__(self, parent_app):
        self.parent_app = parent_app
        self.window = None
        self.is_showing = False
        
    def create_overlay(self):
        """建立覆蓋視窗"""
        self.window = tk.Toplevel()
        self.window.title("翻譯覆蓋")
        
        # 設定視窗屬性
        self.window.attributes('-topmost', True)
        self.window.attributes('-alpha', 0.8)
        self.window.overrideredirect(True)
        
        # 設定視窗樣式
        self.window.configure(bg='black')
        
        # 翻譯文字標籤
        self.text_label = tk.Label(
            self.window,
            text="",
            bg='black',
            fg='yellow',
            font=('Microsoft JhengHei', 14, 'bold'),
            wraplength=400,
            justify=tk.LEFT
        )
        self.text_label.pack(padx=10, pady=10)
        
        # 使視窗可拖動
        self.window.bind('<Button-1>', self.start_move)
        self.window.bind('<B1-Motion>', self.on_move)
        
        # 初始位置
        self.window.geometry("+100+100")
        
    def start_move(self, event):
        self.x = event.x
        self.y = event.y
        
    def on_move(self, event):
        deltax = event.x - self.x
        deltay = event.y - self.y
        x = self.window.winfo_x() + deltax
        y = self.window.winfo_y() + deltay
        self.window.geometry(f"+{x}+{y}")
        
    def update_text(self, text):
        """更新覆蓋文字"""
        if not self.window:
            self.create_overlay()
        self.text_label.config(text=text)
        
    def toggle(self):
        """切換顯示/隱藏"""
        if not self.window:
            self.create_overlay()
            self.is_showing = True
        else:
            if self.is_showing:
                self.window.withdraw()
            else:
                self.window.deiconify()
            self.is_showing = not self.is_showing

class GameTranslatorEnhanced:
    def __init__(self, root):
        self.root = root
        self.root.title("遊戲韓文翻譯器 Pro v2.0")
        self.root.geometry("1000x700")
        
        # 初始化元件
        self.translator = Translator()
        self.overlay = OverlayWindow(self)
        
        # 初始化 EasyOCR（可選）
        self.ocr_engine = 'tesseract'  # 預設使用 tesseract
        self.easyocr_reader = None
        
        # 狀態變數
        self.is_capturing = False
        self.capture_region = None
        self.translation_cache = {}
        self.translation_history = deque(maxlen=100)
        self.hotkey_enabled = True
        
        # 設定
        self.settings = {
            'ocr_engine': 'tesseract',
            'translation_api': 'google',
            'update_interval': 0.5,
            'preprocessing': True,
            'overlay_enabled': False,
            'auto_copy': False,
            'sound_notification': False
        }
        
        # 載入設定
        self.load_settings()
        
        # 設定樣式
        self.setup_styles()
        
        # 建立UI
        self.create_ui()
        
        # 設定快捷鍵
        self.setup_hotkeys()
        
    def setup_styles(self):
        """設定視覺樣式"""
        self.root.configure(bg='#1e1e1e')
        
        style = ttk.Style()
        style.theme_use('clam')
        
        # 深色主題配置
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
        
        # 設定頁面
        settings_tab = ttk.Frame(notebook)
        notebook.add(settings_tab, text="設定")
        self.create_settings_tab(settings_tab)
        
        # 歷史記錄頁面
        history_tab = ttk.Frame(notebook)
        notebook.add(history_tab, text="歷史")
        self.create_history_tab(history_tab)
        
        # 狀態列
        self.create_status_bar()
        
    def create_main_tab(self, parent):
        """建立主要翻譯介面"""
        # 工具列
        toolbar = tk.Frame(parent, bg='#2d2d2d', height=50)
        toolbar.pack(fill=tk.X, padx=5, pady=5)
        
        # 建立按鈕
        buttons_data = [
            ("選擇區域", self.select_capture_region, '#4CAF50', 'F2'),
            ("開始偵測", self.toggle_capture, '#2196F3', 'F3'),
            ("覆蓋顯示", self.toggle_overlay, '#FF9800', 'F4'),
            ("清除", self.clear_current, '#f44336', None),
            ("儲存", self.save_current_session, '#9C27B0', 'Ctrl+S')
        ]
        
        for text, command, color, hotkey in buttons_data:
            btn_frame = tk.Frame(toolbar, bg='#2d2d2d')
            btn_frame.pack(side=tk.LEFT, padx=5)
            
            btn = tk.Button(
                btn_frame,
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
            btn.pack()
            
            if hotkey:
                tk.Label(
                    btn_frame,
                    text=hotkey,
                    bg='#2d2d2d',
                    fg='#888',
                    font=('Arial', 8)
                ).pack()
        
        # 顯示區域
        display_frame = tk.Frame(parent, bg='#1e1e1e')
        display_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 左側：擷取預覽
        left_frame = tk.Frame(display_frame, bg='#1e1e1e')
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
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
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 即時翻譯顯示
        instant_frame = tk.LabelFrame(
            right_frame,
            text="即時翻譯",
            bg='#1e1e1e',
            fg='white',
            font=('Arial', 10, 'bold')
        )
        instant_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.instant_korean = tk.Text(
            instant_frame,
            height=3,
            bg='#0d0d0d',
            fg='#64B5F6',
            font=('Malgun Gothic', 12),
            wrap=tk.WORD
        )
        self.instant_korean.pack(fill=tk.X, padx=5, pady=5)
        
        self.instant_chinese = tk.Text(
            instant_frame,
            height=3,
            bg='#0d0d0d',
            fg='#81C784',
            font=('Microsoft JhengHei', 12),
            wrap=tk.WORD
        )
        self.instant_chinese.pack(fill=tk.X, padx=5, pady=5)
        
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
        self.translation_display.tag_configure('korean', foreground='#64B5F6')
        self.translation_display.tag_configure('chinese', foreground='#81C784')
        
    def create_settings_tab(self, parent):
        """建立設定頁面"""
        settings_frame = tk.Frame(parent, bg='#1e1e1e')
        settings_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # OCR 引擎設定
        ocr_frame = tk.LabelFrame(
            settings_frame,
            text="OCR 設定",
            bg='#1e1e1e',
            fg='white',
            font=('Arial', 11, 'bold')
        )
        ocr_frame.pack(fill=tk.X, pady=10)
        
        self.ocr_var = tk.StringVar(value=self.settings['ocr_engine'])
        
        tk.Radiobutton(
            ocr_frame,
            text="Tesseract (快速)",
            variable=self.ocr_var,
            value='tesseract',
            bg='#1e1e1e',
            fg='white',
            selectcolor='#1e1e1e',
            activebackground='#1e1e1e',
            activeforeground='white'
        ).pack(anchor=tk.W, padx=20, pady=5)
        
        tk.Radiobutton(
            ocr_frame,
            text="EasyOCR (準確)",
            variable=self.ocr_var,
            value='easyocr',
            bg='#1e1e1e',
            fg='white',
            selectcolor='#1e1e1e',
            activebackground='#1e1e1e',
            activeforeground='white'
        ).pack(anchor=tk.W, padx=20, pady=5)
        
        # 功能設定
        feature_frame = tk.LabelFrame(
            settings_frame,
            text="功能設定",
            bg='#1e1e1e',
            fg='white',
            font=('Arial', 11, 'bold')
        )
        feature_frame.pack(fill=tk.X, pady=10)
        
        self.preprocessing_var = tk.BooleanVar(value=self.settings['preprocessing'])
        tk.Checkbutton(
            feature_frame,
            text="啟用影像預處理",
            variable=self.preprocessing_var,
            bg='#1e1e1e',
            fg='white',
            selectcolor='#1e1e1e',
            activebackground='#1e1e1e',
            activeforeground='white'
        ).pack(anchor=tk.W, padx=20, pady=5)
        
        self.auto_copy_var = tk.BooleanVar(value=self.settings['auto_copy'])
        tk.Checkbutton(
            feature_frame,
            text="自動複製翻譯結果",
            variable=self.auto_copy_var,
            bg='#1e1e1e',
            fg='white',
            selectcolor='#1e1e1e',
            activebackground='#1e1e1e',
            activeforeground='white'
        ).pack(anchor=tk.W, padx=20, pady=5)
        
        # 更新間隔
        interval_frame = tk.Frame(feature_frame, bg='#1e1e1e')
        interval_frame.pack(anchor=tk.W, padx=20, pady=5)
        
        tk.Label(
            interval_frame,
            text="更新間隔 (秒):",
            bg='#1e1e1e',
            fg='white'
        ).pack(side=tk.LEFT)
        
        self.interval_var = tk.DoubleVar(value=self.settings['update_interval'])
        interval_scale = tk.Scale(
            interval_frame,
            from_=0.1,
            to=2.0,
            resolution=0.1,
            orient=tk.HORIZONTAL,
            variable=self.interval_var,
            bg='#1e1e1e',
            fg='white',
            highlightthickness=0,
            troughcolor='#444'
        )
        interval_scale.pack(side=tk.LEFT, padx=10)
        
        # 儲存按鈕
        save_btn = tk.Button(
            settings_frame,
            text="儲存設定",
            command=self.save_settings,
            bg='#4CAF50',
            fg='white',
            font=('Arial', 10, 'bold'),
            padx=20,
            pady=10
        )
        save_btn.pack(pady=20)
        
    def create_history_tab(self, parent):
        """建立歷史記錄頁面"""
        history_frame = tk.Frame(parent, bg='#1e1e1e')
        history_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
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
            text="匯入歷史",
            command=self.import_history,
            bg='#2196F3',
            fg='white',
            font=('Arial', 9),
            padx=15,
            pady=5
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            toolbar,
            text="清空歷史",
            command=self.clear_all_history,
            bg='#f44336',
            fg='white',
            font=('Arial', 9),
            padx=15,
            pady=5
        ).pack(side=tk.LEFT, padx=5)
        
        # 搜尋框
        search_frame = tk.Frame(toolbar, bg='#2d2d2d')
        search_frame.pack(side=tk.RIGHT, padx=5)
        
        tk.Label(
            search_frame,
            text="搜尋:",
            bg='#2d2d2d',
            fg='white'
        ).pack(side=tk.LEFT)
        
        self.search_var = tk.StringVar()
        search_entry = tk.Entry(
            search_frame,
            textvariable=self.search_var,
            bg='#444',
            fg='white',
            insertbackground='white'
        )
        search_entry.pack(side=tk.LEFT, padx=5)
        search_entry.bind('<Return>', self.search_history)
        
        # 歷史列表
        self.history_listbox = tk.Listbox(
            history_frame,
            bg='#0d0d0d',
            fg='white',
            selectbackground='#444',
            font=('Microsoft JhengHei', 10)
        )
        self.history_listbox.pack(fill=tk.BOTH, expand=True)
        
        # 綁定雙擊事件
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
        
    def setup_hotkeys(self):
        """設定全域快捷鍵"""
        keyboard.add_hotkey('f2', self.select_capture_region)
        keyboard.add_hotkey('f3', self.toggle_capture)
        keyboard.add_hotkey('f4', self.toggle_overlay)
        keyboard.add_hotkey('ctrl+s', self.save_current_session)
        keyboard.add_hotkey('ctrl+c', self.copy_latest_translation)
        
    def select_capture_region(self):
        """選擇螢幕擷取區域 - 改進版"""
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
            text="拖曳滑鼠選擇區域 | ESC 取消",
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
            # 初始化 OCR 引擎
            if self.ocr_var.get() == 'easyocr' and not self.easyocr_reader:
                self.status_label.config(text="正在載入 EasyOCR...", fg='#FFC107')
                self.root.update()
                try:
                    self.easyocr_reader = easyocr.Reader(['ko', 'ch_tra'])
                except:
                    messagebox.showerror("錯誤", "EasyOCR 載入失敗，切換至 Tesseract")
                    self.ocr_var.set('tesseract')
                    
            # 更新按鈕和狀態
            for widget in self.root.winfo_children():
                if isinstance(widget, tk.Button) and widget['text'] == '開始偵測':
                    widget.config(text='停止偵測', bg='#f44336')
                    
            self.status_label.config(text="偵測中...", fg='#4CAF50')
            
            # 啟動擷取執行緒
            self.capture_thread = threading.Thread(target=self.capture_loop, daemon=True)
            self.capture_thread.start()
        else:
            # 更新按鈕和狀態
            for widget in self.root.winfo_children():
                if isinstance(widget, tk.Button) and widget['text'] == '停止偵測':
                    widget.config(text='開始偵測', bg='#2196F3')
                    
            self.status_label.config(text="已停止", fg='#FFC107')
            
    def capture_loop(self):
        """改進的擷取循環"""
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
                    processed_img = self.advanced_preprocess(screenshot)
                else:
                    processed_img = np.array(screenshot)
                    
                # OCR 識別
                korean_text = ""
                if self.ocr_var.get() == 'tesseract':
                    korean_text = pytesseract.image_to_string(
                        processed_img,
                        lang='kor',
                        config='--psm 6'
                    ).strip()
                elif self.ocr_var.get() == 'easyocr' and self.easyocr_reader:
                    results = self.easyocr_reader.readtext(processed_img)
                    korean_text = ' '.join([result[1] for result in results])
                    
                # 檢查是否為新文字
                current_hash = hash(korean_text)
                if korean_text and current_hash != last_text_hash:
                    last_text_hash = current_hash
                    
                    # 檢查快取
                    if korean_text in self.translation_cache:
                        chinese_text = self.translation_cache[korean_text]
                    else:
                        # 翻譯
                        chinese_text = self.translate_text(korean_text)
                        self.translation_cache[korean_text] = chinese_text
                        
                    # 更新顯示
                    self.root.after(0, self.update_translation, korean_text, chinese_text)
                    
                    # 自動複製
                    if self.auto_copy_var.get():
                        self.root.clipboard_clear()
                        self.root.clipboard_append(chinese_text)
                        
            except Exception as e:
                print(f"擷取錯誤: {e}")
                
            time.sleep(self.interval_var.get())
            
    def advanced_preprocess(self, image):
        """進階影像預處理"""
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
        
    def translate_text(self, text):
        """翻譯文字"""
        try:
            result = self.translator.translate(text, src='ko', dest='zh-tw')
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
            
    def update_translation(self, korean_text, chinese_text):
        """更新翻譯顯示"""
        # 更新即時顯示
        self.instant_korean.delete(1.0, tk.END)
        self.instant_korean.insert(1.0, korean_text)
        
        self.instant_chinese.delete(1.0, tk.END)
        self.instant_chinese.insert(1.0, chinese_text)
        
        # 更新覆蓋視窗
        if self.overlay.is_showing:
            self.overlay.update_text(f"{korean_text}\n{chinese_text}")
        
        # 加入歷史記錄
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # 更新翻譯顯示區
        self.translation_display.insert(tk.END, f"[{timestamp}] ", 'time')
        self.translation_display.insert(tk.END, f"韓: {korean_text}\n", 'korean')
        self.translation_display.insert(tk.END, f"中: {chinese_text}\n", 'chinese')
        self.translation_display.insert(tk.END, "-" * 60 + "\n")
        self.translation_display.see(tk.END)
        
        # 儲存到歷史
        history_item = {
            'timestamp': timestamp,
            'date': datetime.now().strftime("%Y-%m-%d"),
            'korean': korean_text,
            'chinese': chinese_text
        }
        self.translation_history.append(history_item)
        
        # 更新歷史列表
        self.history_listbox.insert(
            tk.END,
            f"{timestamp} | {korean_text[:20]}... → {chinese_text[:20]}..."
        )
        
    def toggle_overlay(self):
        """切換覆蓋顯示"""
        self.overlay.toggle()
        self.settings['overlay_enabled'] = self.overlay.is_showing
        
    def clear_current(self):
        """清除當前顯示"""
        self.instant_korean.delete(1.0, tk.END)
        self.instant_chinese.delete(1.0, tk.END)
        self.translation_display.delete(1.0, tk.END)
        
    def save_current_session(self):
        """儲存當前工作階段"""
        if not self.translation_history:
            messagebox.showinfo("提示", "沒有可儲存的翻譯記錄")
            return
            
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialfile=f"translation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        )
        
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(list(self.translation_history), f, ensure_ascii=False, indent=2)
                self.status_label.config(text=f"已儲存: {os.path.basename(filename)}", fg='#4CAF50')
            except Exception as e:
                messagebox.showerror("錯誤", f"儲存失敗: {str(e)}")
                
    def copy_latest_translation(self):
        """複製最新翻譯"""
        if self.translation_history:
            latest = self.translation_history[-1]
            text = f"{latest['korean']}\n{latest['chinese']}"
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.status_label.config(text="已複製到剪貼簿", fg='#4CAF50')
            
    def search_history(self, event=None):
        """搜尋歷史記錄"""
        keyword = self.search_var.get().lower()
        self.history_listbox.delete(0, tk.END)
        
        for item in self.translation_history:
            if (keyword in item['korean'].lower() or 
                keyword in item['chinese'].lower()):
                display_text = (f"{item['timestamp']} | "
                              f"{item['korean'][:20]}... → "
                              f"{item['chinese'][:20]}...")
                self.history_listbox.insert(tk.END, display_text)
                
    def show_history_detail(self, event):
        """顯示歷史詳細內容"""
        selection = self.history_listbox.curselection()
        if selection:
            index = selection[0]
            if index < len(self.translation_history):
                item = self.translation_history[index]
                
                detail_window = tk.Toplevel(self.root)
                detail_window.title("翻譯詳情")
                detail_window.geometry("500x300")
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
                
                detail_text.insert(tk.END, f"時間: {item['date']} {item['timestamp']}\n\n")
                detail_text.insert(tk.END, f"韓文:\n{item['korean']}\n\n")
                detail_text.insert(tk.END, f"中文:\n{item['chinese']}")
                detail_text.config(state=tk.DISABLED)
                
    def export_history(self):
        """匯出歷史記錄"""
        if not self.translation_history:
            messagebox.showinfo("提示", "沒有可匯出的歷史記錄")
            return
            
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[
                ("Text files", "*.txt"),
                ("JSON files", "*.json"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        
        if filename:
            try:
                if filename.endswith('.json'):
                    with open(filename, 'w', encoding='utf-8') as f:
                        json.dump(list(self.translation_history), f, ensure_ascii=False, indent=2)
                elif filename.endswith('.csv'):
                    import csv
                    with open(filename, 'w', encoding='utf-8-sig', newline='') as f:
                        writer = csv.DictWriter(f, fieldnames=['date', 'timestamp', 'korean', 'chinese'])
                        writer.writeheader()
                        writer.writerows(self.translation_history)
                else:  # txt
                    with open(filename, 'w', encoding='utf-8') as f:
                        for item in self.translation_history:
                            f.write(f"[{item['date']} {item['timestamp']}]\n")
                            f.write(f"韓文: {item['korean']}\n")
                            f.write(f"中文: {item['chinese']}\n")
                            f.write("-" * 60 + "\n\n")
                            
                messagebox.showinfo("成功", "歷史記錄已匯出")
            except Exception as e:
                messagebox.showerror("錯誤", f"匯出失敗: {str(e)}")
                
    def import_history(self):
        """匯入歷史記錄"""
        filename = filedialog.askopenfilename(
            filetypes=[
                ("JSON files", "*.json"),
                ("All files", "*.*")
            ]
        )
        
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    imported_data = json.load(f)
                    
                # 合併歷史記錄
                for item in imported_data:
                    self.translation_history.append(item)
                    display_text = (f"{item['timestamp']} | "
                                  f"{item['korean'][:20]}... → "
                                  f"{item['chinese'][:20]}...")
                    self.history_listbox.insert(tk.END, display_text)
                    
                messagebox.showinfo("成功", f"已匯入 {len(imported_data)} 筆記錄")
            except Exception as e:
                messagebox.showerror("錯誤", f"匯入失敗: {str(e)}")
                
    def clear_all_history(self):
        """清空所有歷史記錄"""
        if messagebox.askyesno("確認", "確定要清空所有歷史記錄嗎？"):
            self.translation_history.clear()
            self.history_listbox.delete(0, tk.END)
            self.translation_display.delete(1.0, tk.END)
            self.status_label.config(text="已清空歷史記錄", fg='#4CAF50')
            
    def save_settings(self):
        """儲存設定"""
        self.settings['ocr_engine'] = self.ocr_var.get()
        self.settings['preprocessing'] = self.preprocessing_var.get()
        self.settings['auto_copy'] = self.auto_copy_var.get()
        self.settings['update_interval'] = self.interval_var.get()
        
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
                    self.settings.update(json.load(f))
        except Exception as e:
            print(f"載入設定失敗: {e}")
            
    def on_closing(self):
        """關閉程式時的處理"""
        self.is_capturing = False
        self.save_settings()
        self.root.destroy()

def main():
    """主程式入口"""
    # 檢查系統需求
    if os.name == 'nt':  # Windows
        # 設定 DPI 感知
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        
        # 檢查 Tesseract
        tesseract_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        if os.path.exists(tesseract_path):
            pytesseract.pytesseract.tesseract_cmd = tesseract_path
        else:
            response = messagebox.askyesno(
                "Tesseract 未安裝",
                "未偵測到 Tesseract OCR。\n"
                "是否要繼續執行？（只能使用 EasyOCR）"
            )
            if not response:
                return
                
    # 建立並執行應用程式
    root = tk.Tk()
    app = GameTranslatorEnhanced(root)
    
    # 設定關閉事件
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    # 執行主循環
    root.mainloop()

if __name__ == "__main__":
    main()