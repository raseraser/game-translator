import tkinter as tk
from tkinter import ttk, scrolledtext
import pyautogui
import pytesseract
from PIL import Image, ImageTk, ImageDraw, ImageFont
import cv2
import numpy as np
from googletrans import Translator
import threading
import time
import keyboard
from datetime import datetime
import json
import os

class GameTranslatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("遊戲韓文即時翻譯器")
        self.root.geometry("800x600")
        
        # 初始化翻譯器
        self.translator = Translator()
        
        # 狀態變數
        self.is_capturing = False
        self.capture_region = None
        self.translation_history = []
        
        # 設定樣式
        self.setup_styles()
        
        # 建立UI
        self.create_ui()
        
        # 設定快捷鍵
        self.setup_hotkeys()
        
    def setup_styles(self):
        """設定視覺樣式"""
        self.root.configure(bg='#2b2b2b')
        style = ttk.Style()
        style.theme_use('clam')
        
        # 配置深色主題
        style.configure('TLabel', background='#2b2b2b', foreground='white')
        style.configure('TButton', background='#404040', foreground='white')
        style.configure('TFrame', background='#2b2b2b')
        
    def create_ui(self):
        """建立使用者介面"""
        # 主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 控制面板
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 區域選擇按鈕
        self.select_btn = tk.Button(
            control_frame,
            text="選擇擷取區域",
            command=self.select_capture_region,
            bg='#4CAF50',
            fg='white',
            font=('Arial', 10, 'bold'),
            padx=20,
            pady=5
        )
        self.select_btn.pack(side=tk.LEFT, padx=5)
        
        # 開始/停止按鈕
        self.toggle_btn = tk.Button(
            control_frame,
            text="開始偵測",
            command=self.toggle_capture,
            bg='#2196F3',
            fg='white',
            font=('Arial', 10, 'bold'),
            padx=20,
            pady=5
        )
        self.toggle_btn.pack(side=tk.LEFT, padx=5)
        
        # 清除歷史按鈕
        self.clear_btn = tk.Button(
            control_frame,
            text="清除歷史",
            command=self.clear_history,
            bg='#f44336',
            fg='white',
            font=('Arial', 10, 'bold'),
            padx=20,
            pady=5
        )
        self.clear_btn.pack(side=tk.LEFT, padx=5)
        
        # 狀態標籤
        self.status_label = tk.Label(
            control_frame,
            text="準備就緒",
            bg='#2b2b2b',
            fg='#4CAF50',
            font=('Arial', 10)
        )
        self.status_label.pack(side=tk.RIGHT, padx=5)
        
        # 分隔線
        separator = ttk.Separator(main_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=10)
        
        # 顯示區域
        display_frame = ttk.Frame(main_frame)
        display_frame.pack(fill=tk.BOTH, expand=True)
        
        # 預覽圖片框架
        preview_frame = ttk.LabelFrame(display_frame, text="擷取預覽")
        preview_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        self.preview_label = tk.Label(
            preview_frame,
            text="尚未擷取",
            bg='#1e1e1e',
            fg='gray',
            font=('Arial', 12)
        )
        self.preview_label.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 翻譯結果框架
        translation_frame = ttk.LabelFrame(display_frame, text="翻譯結果")
        translation_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # 翻譯文字顯示區
        self.translation_text = scrolledtext.ScrolledText(
            translation_frame,
            wrap=tk.WORD,
            bg='#1e1e1e',
            fg='white',
            font=('Microsoft JhengHei', 11),
            insertbackground='white'
        )
        self.translation_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 設定文字標籤樣式
        self.translation_text.tag_configure('korean', foreground='#64B5F6')
        self.translation_text.tag_configure('chinese', foreground='#81C784')
        self.translation_text.tag_configure('timestamp', foreground='#FFB74D')
        
    def setup_hotkeys(self):
        """設定快捷鍵"""
        keyboard.add_hotkey('ctrl+shift+s', self.select_capture_region)
        keyboard.add_hotkey('ctrl+shift+t', self.toggle_capture)
        
    def select_capture_region(self):
        """選擇螢幕擷取區域"""
        self.status_label.config(text="請拖曳選擇區域...", fg='#FFC107')
        self.root.withdraw()  # 隱藏主視窗
        
        # 建立全螢幕選擇視窗
        selection_window = tk.Toplevel()
        selection_window.attributes('-fullscreen', True)
        selection_window.attributes('-alpha', 0.3)
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
        
        # 滑鼠事件變數
        start_x = start_y = None
        rect_id = None
        
        def on_mouse_down(event):
            nonlocal start_x, start_y, rect_id
            start_x, start_y = event.x, event.y
            if rect_id:
                canvas.delete(rect_id)
            rect_id = canvas.create_rectangle(
                start_x, start_y, start_x, start_y,
                outline='red', width=2
            )
            
        def on_mouse_move(event):
            nonlocal rect_id
            if start_x is not None and rect_id:
                canvas.coords(rect_id, start_x, start_y, event.x, event.y)
                
        def on_mouse_up(event):
            nonlocal start_x, start_y
            if start_x is not None:
                # 計算選擇區域
                x1, y1 = min(start_x, event.x), min(start_y, event.y)
                x2, y2 = max(start_x, event.x), max(start_y, event.y)
                
                if x2 - x1 > 10 and y2 - y1 > 10:  # 最小區域限制
                    self.capture_region = (x1, y1, x2 - x1, y2 - y1)
                    self.status_label.config(
                        text=f"已選擇區域: {x2-x1}x{y2-y1}",
                        fg='#4CAF50'
                    )
                
            selection_window.destroy()
            self.root.deiconify()  # 恢復主視窗
            
        # 綁定滑鼠事件
        canvas.bind('<Button-1>', on_mouse_down)
        canvas.bind('<B1-Motion>', on_mouse_move)
        canvas.bind('<ButtonRelease-1>', on_mouse_up)
        
        # ESC 鍵取消
        selection_window.bind('<Escape>', lambda e: (
            selection_window.destroy(),
            self.root.deiconify(),
            self.status_label.config(text="已取消選擇", fg='#f44336')
        ))
        
    def toggle_capture(self):
        """切換擷取狀態"""
        if not self.capture_region:
            self.status_label.config(text="請先選擇擷取區域！", fg='#f44336')
            return
            
        self.is_capturing = not self.is_capturing
        
        if self.is_capturing:
            self.toggle_btn.config(text="停止偵測", bg='#f44336')
            self.status_label.config(text="正在偵測中...", fg='#4CAF50')
            # 啟動擷取執行緒
            self.capture_thread = threading.Thread(target=self.capture_loop, daemon=True)
            self.capture_thread.start()
        else:
            self.toggle_btn.config(text="開始偵測", bg='#2196F3')
            self.status_label.config(text="已停止偵測", fg='#FFC107')
            
    def capture_loop(self):
        """擷取循環"""
        last_text = ""
        
        while self.is_capturing:
            try:
                # 擷取指定區域
                x, y, w, h = self.capture_region
                screenshot = pyautogui.screenshot(region=(x, y, w, h))
                
                # 轉換為OpenCV格式
                img_cv = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
                
                # 影像預處理
                processed_img = self.preprocess_image(img_cv)
                
                # OCR識別韓文
                korean_text = pytesseract.image_to_string(
                    processed_img,
                    lang='kor',
                    config='--psm 6'
                ).strip()
                
                # 如果識別到新文字
                if korean_text and korean_text != last_text:
                    last_text = korean_text
                    
                    # 翻譯
                    chinese_text = self.translate_text(korean_text)
                    
                    # 更新UI
                    self.root.after(0, self.update_display, screenshot, korean_text, chinese_text)
                    
                # 更新預覽（每秒更新）
                self.root.after(0, self.update_preview, screenshot)
                
            except Exception as e:
                print(f"擷取錯誤: {e}")
                
            time.sleep(0.5)  # 每0.5秒檢查一次
            
    def preprocess_image(self, img):
        """影像預處理以改善OCR效果"""
        # 轉換為灰階
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        # 放大影像
        scaled = cv2.resize(gray, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
        
        # 應用二值化
        _, binary = cv2.threshold(scaled, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        
        # 去噪
        denoised = cv2.medianBlur(binary, 3)
        
        return denoised
        
    def translate_text(self, korean_text):
        """翻譯韓文為中文"""
        try:
            result = self.translator.translate(korean_text, src='ko', dest='zh-tw')
            return result.text
        except Exception as e:
            return f"翻譯錯誤: {str(e)}"
            
    def update_preview(self, screenshot):
        """更新預覽圖片"""
        # 調整圖片大小以適應預覽區域
        img = screenshot.copy()
        img.thumbnail((300, 300), Image.Resampling.LANCZOS)
        
        # 轉換為PhotoImage
        photo = ImageTk.PhotoImage(img)
        
        # 更新標籤
        self.preview_label.config(image=photo, text="")
        self.preview_label.image = photo  # 保持引用
        
    def update_display(self, screenshot, korean_text, chinese_text):
        """更新顯示內容"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # 插入時間戳記
        self.translation_text.insert(tk.END, f"\n[{timestamp}] ", 'timestamp')
        
        # 插入韓文
        self.translation_text.insert(tk.END, "韓文: ", 'korean')
        self.translation_text.insert(tk.END, f"{korean_text}\n")
        
        # 插入中文翻譯
        self.translation_text.insert(tk.END, "中文: ", 'chinese')
        self.translation_text.insert(tk.END, f"{chinese_text}\n")
        
        # 插入分隔線
        self.translation_text.insert(tk.END, "-" * 50 + "\n")
        
        # 自動捲動到底部
        self.translation_text.see(tk.END)
        
        # 儲存到歷史記錄
        self.translation_history.append({
            'timestamp': timestamp,
            'korean': korean_text,
            'chinese': chinese_text
        })
        
    def clear_history(self):
        """清除翻譯歷史"""
        self.translation_text.delete(1.0, tk.END)
        self.translation_history.clear()
        self.status_label.config(text="已清除歷史", fg='#4CAF50')
        
    def save_history(self):
        """儲存翻譯歷史"""
        if not self.translation_history:
            return
            
        filename = f"translation_history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(self.translation_history, f, ensure_ascii=False, indent=2)
            
        self.status_label.config(text=f"已儲存至 {filename}", fg='#4CAF50')

def main():
    # 檢查必要的套件和設定
    try:
        # 設定Tesseract路徑（Windows）
        # 如果是其他系統，請調整路徑
        if os.name == 'nt':  # Windows
            pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    except:
        pass
    
    # 建立主視窗
    root = tk.Tk()
    app = GameTranslatorApp(root)
    
    # 啟動主循環
    root.mainloop()

if __name__ == "__main__":
    main()
