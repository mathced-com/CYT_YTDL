import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import yt_dlp
import threading
import os
import sys
import re
import time
import urllib.request
import io
import json
import subprocess
import zipfile
import shutil
import ssl
import ctypes
import math

ssl._create_default_https_context = ssl._create_unverified_context
APP_VERSION = "2.3.2"
GITHUB_REPO = "mathced-com/CYT_YTDL"


# ===========================================================================
# 使用者回饋系統的 Google Apps Script 後台網址
# 若未來有更換新的後台 URL，請直接在此處進行替換即可！
# ===========================================================================
FEEDBACK_API_URL = "https://script.google.com/macros/s/AKfycbztgKmePvfpuWwvNgKMJK_uRyUvanOpG0TkHpCwTGyDHn1k2SyGDNNnRJFEYmFmmOLA/exec"


try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# 移除 os.chdir 以免鎖定暫存資料夾
# 後續路徑皆改用 self.app_dir 等絕對路徑管理

class ScrollableFrame(ttk.Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # 讓內部 frame 寬度填滿 canvas，以便靠右排版正常運作
        self.canvas.bind(
            "<Configure>",
            lambda e: self.canvas.itemconfig(self.canvas_window, width=e.width)
        )
        
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

class CancelLogger:
    def __init__(self, gui):
        self.gui = gui
    def debug(self, msg):
        if not getattr(self.gui, 'is_analyzing', False):
            raise ValueError("USER_CANCELLED")
    def info(self, msg):
        if not getattr(self.gui, 'is_analyzing', False):
            raise ValueError("USER_CANCELLED")
    def warning(self, msg):
        if not getattr(self.gui, 'is_analyzing', False):
            raise ValueError("USER_CANCELLED")
    def error(self, msg):
        if not getattr(self.gui, 'is_analyzing', False):
            raise ValueError("USER_CANCELLED")

class YouTubeDownloaderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"CYT_網路影音下載器 v{APP_VERSION}")
        self.root.geometry("850x730")
        self.root.resizable(False, False)
        
        self.app_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
        
        try:
            self.root.iconbitmap(self.resource_path("icon.ico"))
        except Exception:
            pass
            
        # 清理更新時遺留的舊版檔案
        for f in os.listdir(self.app_dir):
            if f.endswith('.old'):
                try:
                    os.remove(os.path.join(self.app_dir, f))
                except Exception:
                    pass
        
        default_dl_dir = os.path.join(self.app_dir, "download")
        os.makedirs(default_dl_dir, exist_ok=True)
        self.download_path = tk.StringVar(value=default_dl_dir)
        self.format_choice = tk.StringVar(value="mp4")
        self.quality_choice = tk.StringVar()
        
        self.video_info = None
        self.is_playlist = False
        
        self.playlist_all_vars = []
        self.playlist_all_entries = []
        self.playlist_current_page = 0
        
        # 章節分割相關
        self.current_chapters = []
        self.split_by_chapters = tk.BooleanVar(value=False)
        self._downloaded_filepath = None  # 追蹤最後下載的完整檔案路徑
        
        # 載入持久化配置
        self.config_path = os.path.join(self.app_dir, "config.json")
        self.load_config()
        
        # 暫停與取消狀態標記
        self.is_paused = False
        self.is_cancelled = False
        self.is_analyzing = False
        self.playlist_status_labels = []
        
        self.create_widgets()
        self.update_quality_options()
        
        # 如果配置中有舊的品質設定，嘗試套用
        conf = self._get_config_data()
        if conf.get("quality_choice"):
            if conf["quality_choice"] in self.quality_combo['values']:
                self.quality_choice.set(conf["quality_choice"])

        self.check_ffmpeg_environment()
        
        if not HAS_PIL:
            messagebox.showwarning("缺少套件", "系統缺少 Pillow 套件，將無法顯示影片封面。")

        # 啟動 2 秒後自動背景靜默檢查更新
        self.root.after(2000, lambda: self.check_app_update(is_auto=True))

    def resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_path, relative_path)

    def load_config(self):
        default_dl_dir = os.path.join(self.app_dir, "download")
        data = self._get_config_data()
        
        # 讀取下載路徑，預設為主程式目錄下的 download 資料夾
        dl_path = data.get("download_path", default_dl_dir)
        
        # 智慧相容性與 fallback 檢查：
        # 1. 如果設定檔中的路徑含有歷史開發遺留路徑 (例如 YTDL_temp)
        # 2. 或者該路徑在當前電腦上並不存在 (例如以前儲存的特定外接碟磁碟機或特定目錄被刪除)
        # 則嘗試建立它；如果建立失敗 (如沒有該磁碟槽)，則強制 fallback 至預設的主程式 download 目錄。
        if "YTDL_temp" in dl_path or not os.path.exists(dl_path):
            try:
                os.makedirs(dl_path, exist_ok=True)
            except Exception:
                dl_path = default_dl_dir
        
        # 確保該資料夾最終 100% 存在於硬碟中，若沒有則自動生成
        try:
            os.makedirs(dl_path, exist_ok=True)
        except Exception:
            # 備用安全防線：降規至目前工作目錄下的 download
            dl_path = os.path.abspath("download")
            os.makedirs(dl_path, exist_ok=True)
            
        self.download_path = tk.StringVar(value=dl_path)
        self.format_choice = tk.StringVar(value=data.get("format_choice", "mp4"))
        self.quality_choice = tk.StringVar(value=data.get("quality_choice", ""))
        self.cookie_browser = tk.StringVar(value=data.get("cookie_browser", "無"))
        self.cookie_file_path = tk.StringVar(value=data.get("cookie_file_path", ""))
        self.threads_choice = tk.IntVar(value=data.get("threads_choice", 1))
        self.url_entry_var = tk.StringVar()
        self.url_entry_var.trace_add("write", self.on_url_change)

    def _get_config_data(self):
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {}

    def save_config(self, *args):
        data = {
            "download_path": self.download_path.get(),
            "format_choice": self.format_choice.get(),
            "quality_choice": self.quality_choice.get(),
            "cookie_browser": self.cookie_browser.get(),
            "cookie_file_path": self.cookie_file_path.get(),
            "threads_choice": self.threads_choice.get() if hasattr(self, 'threads_choice') else 1
        }
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception:
            pass

    def _on_tab_changed(self, event):
        sel = self.notebook.select()
        if not sel: return
        text = self.notebook.tab(sel, "text")
        if "MP3 裁剪" in text:
            self.trimmer._refresh_list()
        elif "影片裁剪" in text:
            self.video_trimmer._refresh_list()
        elif "合併" in text:
            self.merger._refresh_src_list()
        elif "轉檔" in text:
            self.converter._refresh_list()

    def create_widgets(self):
        # === 全局標題列 (在 Notebook 之上) ===
        header_frame = tk.Frame(self.root, bg="white")
        header_frame.pack(fill="x", pady=10)
        
        try:
            logo_img = Image.open(self.resource_path("icon.png"))
            logo_img = logo_img.resize((32, 32), Image.LANCZOS)
            self.logo_photo = ImageTk.PhotoImage(logo_img)
            tk.Label(header_frame, image=self.logo_photo, bg="white").pack(side="left", padx=(20, 10))
        except Exception:
            pass
            
        tk.Label(header_frame, text=f"CYT_網路影音下載器 v{APP_VERSION}", font=("Microsoft JhengHei", 18, "bold"), bg="white").pack(side="left")

        # 右側使用者回饋按鈕 (Flat UI, 暖橘色)
        self.feedback_btn = tk.Button(
            header_frame, 
            text="💡 使用者回饋", 
            command=self.open_feedback_dialog, 
            font=("Microsoft JhengHei", 10, "bold"), 
            bg="#FF9800", 
            fg="white", 
            relief="flat", 
            padx=10, 
            pady=3,
            cursor="hand2"
        )
        self.feedback_btn.pack(side="right", padx=(5, 20))

        # 右側使用說明按鈕 (Flat UI, 藍色)
        self.help_btn = tk.Button(
            header_frame, 
            text="📖 使用說明", 
            command=self.open_help_dialog, 
            font=("Microsoft JhengHei", 10, "bold"), 
            bg="#2196F3", 
            fg="white", 
            relief="flat", 
            padx=10, 
            pady=3,
            cursor="hand2"
        )
        self.help_btn.pack(side="right", padx=5)

        # === 標簿頁 (Notebook) ===
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Tab 1: YouTube 下載器
        tab_download = tk.Frame(self.notebook)
        self.notebook.add(tab_download, text="  ⬇️ 影音下載器  ")
        
        # Tab 2: MP3 裁剪工具
        tab_trim = tk.Frame(self.notebook)
        self.notebook.add(tab_trim, text="  ✂️ MP3 裁剪工具  ")
        self.trimmer = MP3TrimmerTab(tab_trim, self.download_path)
        
        # Tab 3: MP3 合併工具
        tab_merge = tk.Frame(self.notebook)
        self.notebook.add(tab_merge, text="  🔗 MP3 合併工具  ")
        self.merger = MP3MergerTab(tab_merge, self.download_path)
        
        # Tab 4: 影片裁剪工具
        tab_video_trim = tk.Frame(self.notebook)
        self.notebook.add(tab_video_trim, text="  🎬 影片裁剪工具  ")
        self.video_trimmer = VideoTrimmerTab(tab_video_trim, self.download_path)
        
        # Tab 5: 影音轉檔工具
        tab_converter = tk.Frame(self.notebook)
        self.notebook.add(tab_converter, text="  🔄 影音轉檔工具  ")
        self.converter = VideoConverterTab(tab_converter, self.download_path)
        
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        # === 以下為 Tab 1 (下載器) 的 UI ===
        parent = tab_download
        
        url_frame = tk.Frame(parent)
        url_frame.pack(fill="x", padx=15, pady=5)
        tk.Label(url_frame, text="網址：", font=("Microsoft JhengHei", 12)).pack(side="left")
        
        self.url_entry = tk.Entry(url_frame, width=35, font=("Microsoft JhengHei", 10))
        self.url_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        # 按鈕順序：貼上、清除、解析
        tk.Button(url_frame, text="貼上", command=self.paste_url, font=("Microsoft JhengHei", 10), bg="#FFEB3B").pack(side="left", padx=2)
        
        self.clear_btn = tk.Button(url_frame, text="清除網址", command=self.clear_url, font=("Microsoft JhengHei", 10))
        self.clear_btn.pack(side="left", padx=2)

        self.analyze_btn = tk.Button(url_frame, text="解析網址", command=self.start_analyze, bg="#2196F3", fg="white", font=("Microsoft JhengHei", 10, "bold"))
        self.analyze_btn.pack(side="left", padx=2)
        
        # 步驟提示
        hint_text = "執行步驟：\n一、貼上影音來源網址\n二、點擊「解析網址」\n三、點擊「開始下載」"
        hint_label = tk.Label(url_frame, text=hint_text, fg="#E91E63", font=("Microsoft JhengHei", 9, "bold"), justify="left")
        hint_label.pack(side="left", padx=5)
        
        # 先建立底部框架並鎖定在視窗最下方，保證不被清單擠出畫面
        bottom_frame = tk.Frame(parent)
        bottom_frame.pack(side="bottom", fill="x", pady=5)
        
        self.info_frame = tk.LabelFrame(parent, text="影片預覽 / 播放清單", font=("Microsoft JhengHei", 10))
        self.info_frame.pack(fill="both", expand=True, padx=15, pady=5)
        
        self.title_label = tk.Label(self.info_frame, text="請輸入網址並點選「解析網址」", fg="gray", wraplength=680, justify="left")
        self.title_label.pack(pady=5, padx=10)
        
        self.select_btn_frame = tk.Frame(self.info_frame)
        self.select_all_btn = tk.Button(self.select_btn_frame, text="全部勾選", command=self.select_all, font=("Microsoft JhengHei", 10), bg="#4CAF50", fg="white")
        self.select_all_btn.pack(side="left", padx=5)
        
        self.deselect_all_btn = tk.Button(self.select_btn_frame, text="取消全選", command=self.deselect_all, font=("Microsoft JhengHei", 10))
        self.deselect_all_btn.pack(side="left", padx=5)
        
        # 本頁全選與本頁取消按鈕，預先建立但不 pack，由 show_playlist 動態決定是否 pack
        self.select_page_btn = tk.Button(self.select_btn_frame, text="本頁全部勾選", command=self.select_page, font=("Microsoft JhengHei", 10), bg="#2196F3", fg="white")
        self.deselect_page_btn = tk.Button(self.select_btn_frame, text="本頁全部取消", command=self.deselect_page, font=("Microsoft JhengHei", 10))
        
        # 分頁控制
        self.prev_btn = tk.Button(self.select_btn_frame, text="◀ 上一頁", command=self.prev_page, font=("Microsoft JhengHei", 9))
        self.prev_btn.pack(side="left", padx=10)
        self.page_label = tk.Label(self.select_btn_frame, text="第 1 / 1 頁", font=("Microsoft JhengHei", 9, "bold"))
        self.page_label.pack(side="left")
        self.next_btn = tk.Button(self.select_btn_frame, text="下一頁 ▶", command=self.next_page, font=("Microsoft JhengHei", 9))
        self.next_btn.pack(side="left", padx=10)
 
        self.selection_label = tk.Label(self.select_btn_frame, text="已勾選: 0 / 0", font=("Microsoft JhengHei", 10, "bold"), fg="#E91E63")

        self.selection_label = tk.Label(self.select_btn_frame, text="已勾選: 0 / 0", font=("Microsoft JhengHei", 10, "bold"), fg="#E91E63")
        self.selection_label.pack(side="left", padx=15)
        
        self.select_btn_frame.pack(pady=3)
        self.select_btn_frame.pack_forget()
        
        self.list_frame = ScrollableFrame(self.info_frame)
        
        format_frame = tk.Frame(bottom_frame)
        format_frame.pack(fill="x", padx=15, pady=3)
        tk.Label(format_frame, text="格式：", font=("Microsoft JhengHei", 12)).pack(side="left")
        tk.Radiobutton(format_frame, text="MP4", variable=self.format_choice, value="mp4", command=self.on_format_change).pack(side="left", padx=2)
        tk.Radiobutton(format_frame, text="MKV", variable=self.format_choice, value="mkv", command=self.on_format_change).pack(side="left", padx=2)
        tk.Radiobutton(format_frame, text="MP3", variable=self.format_choice, value="mp3", command=self.on_format_change).pack(side="left", padx=2)
        tk.Radiobutton(format_frame, text="WAV", variable=self.format_choice, value="wav", command=self.on_format_change).pack(side="left", padx=2)
        
        tk.Label(format_frame, text="   品質：", font=("Microsoft JhengHei", 12)).pack(side="left")
        self.quality_combo = ttk.Combobox(format_frame, textvariable=self.quality_choice, state="readonly", width=18)
        self.quality_combo.pack(side="left", padx=5)
        self.quality_combo.bind("<<ComboboxSelected>>", self.save_config)

        tk.Label(format_frame, text="   同時下載數：", font=("Microsoft JhengHei", 10)).pack(side="left")
        self.threads_combo = ttk.Combobox(format_frame, textvariable=self.threads_choice, state="readonly", width=5)
        self.threads_combo['values'] = [1, 2, 3, 4, 5]
        self.threads_combo.pack(side="left", padx=5)
        self.threads_combo.bind("<<ComboboxSelected>>", self.save_config)
        
        path_frame = tk.Frame(bottom_frame)
        path_frame.pack(fill="x", padx=15, pady=3)
        tk.Label(path_frame, text="儲存：", font=("Microsoft JhengHei", 12)).pack(side="left")
        self.path_entry = tk.Entry(path_frame, textvariable=self.download_path, width=40, state="readonly", font=("Microsoft JhengHei", 10))
        self.path_entry.pack(side="left", padx=5, fill="x", expand=True)
        tk.Button(path_frame, text="選擇", command=self.browse_folder).pack(side="left", padx=2)
        tk.Button(path_frame, text="開啟", command=self.open_download_folder, bg="#9C27B0", fg="white").pack(side="left", padx=2)
        
        # Cookies 來源設定
        cookie_frame = tk.Frame(bottom_frame)
        cookie_frame.pack(fill="x", padx=15, pady=3)
        tk.Label(cookie_frame, text="Cookies 來源：", font=("Microsoft JhengHei", 10)).pack(side="left")
        self.cookie_browser_combo = ttk.Combobox(cookie_frame, textvariable=self.cookie_browser, state="readonly", width=16)
        self.cookie_browser_combo['values'] = ["無", "chrome", "edge", "選擇 .txt 檔案..."]
        self.cookie_browser_combo.pack(side="left", padx=5)
        self.cookie_browser_combo.bind("<<ComboboxSelected>>", self.on_cookie_browser_change)
        
        self.cookie_tip_label = tk.Label(cookie_frame, text="(下載私人/限定影片時必備)", font=("Microsoft JhengHei", 8), fg="gray")
        self.cookie_tip_label.pack(side="left")
        
        tk.Button(cookie_frame, text="💡 FB/私人影片下載教學", font=("Microsoft JhengHei", 8), command=self.show_fb_tutorial, fg="#1976D2", relief="flat", cursor="hand2").pack(side="left", padx=10)
        
        status_frame = tk.Frame(bottom_frame)
        status_frame.pack(fill="x", padx=15, pady=3)
        self.progress_bar = ttk.Progressbar(status_frame, orient="horizontal", length=740, mode="determinate")
        self.progress_bar.pack(pady=2)
        self.status_label = tk.Label(status_frame, text="等待解析...", fg="blue", font=("Microsoft JhengHei", 10))
        self.status_label.pack(pady=2)
        
        # 執行與控制按鈕區
        btn_frame = tk.Frame(bottom_frame)
        btn_frame.pack(pady=5)
        self.download_btn = tk.Button(btn_frame, text="開始下載", font=("Microsoft JhengHei", 12, "bold"), bg="#4CAF50", fg="white", width=12, command=self.start_download, state="disabled")
        self.download_btn.pack(side="left", padx=5)
        
        self.pause_btn = tk.Button(btn_frame, text="暫停", font=("Microsoft JhengHei", 10), command=self.toggle_pause, state="disabled", width=8)
        self.pause_btn.pack(side="left", padx=5)
        
        self.cancel_btn = tk.Button(btn_frame, text="取消", font=("Microsoft JhengHei", 10), command=self.cancel_download, state="disabled", bg="#f44336", fg="white", width=8)
        self.cancel_btn.pack(side="left", padx=5)
        
        # 移除 frozen 限制，讓 exe 使用者也能修復核心
        tk.Button(btn_frame, text="修復下載核心", command=self.update_ytdlp, bg="#FF9800", fg="white", font=("Microsoft JhengHei", 9)).pack(side="left", padx=5)
            
        tk.Button(btn_frame, text="檢查主程式更新", command=self.check_app_update, bg="#FF9800", fg="white", font=("Microsoft JhengHei", 9)).pack(side="left", padx=5)

    def select_all(self):
        for var in self.playlist_all_vars:
            var.set(True)
        self.update_selection_count()
            
    def deselect_all(self):
        for var in self.playlist_all_vars:
            var.set(False)
        self.update_selection_count()

    def select_page(self):
        if not hasattr(self, 'playlist_all_vars') or not self.playlist_all_vars:
            return
        start_idx = self.playlist_current_page * 50
        end_idx = min(start_idx + 50, len(self.playlist_all_vars))
        for i in range(start_idx, end_idx):
            self.playlist_all_vars[i].set(True)
        self.update_selection_count()

    def deselect_page(self):
        if not hasattr(self, 'playlist_all_vars') or not self.playlist_all_vars:
            return
        start_idx = self.playlist_current_page * 50
        end_idx = min(start_idx + 50, len(self.playlist_all_vars))
        for i in range(start_idx, end_idx):
            self.playlist_all_vars[i].set(False)
        self.update_selection_count()

    def update_selection_count(self):
        total = len(self.playlist_all_vars)
        selected = sum(1 for var in self.playlist_all_vars if var.get())
        self.selection_label.config(text=f"已勾選: {selected} / {total}")

    def paste_url(self):
        try:
            clipboard = self.root.clipboard_get()
            self.url_entry.delete(0, tk.END)
            self.url_entry.insert(0, clipboard)
        except Exception:
            messagebox.showwarning("貼上失敗", "剪貼簿中沒有可讀取的內容。")

    def open_download_folder(self):
        path = self.download_path.get()
        if os.path.exists(path):
            os.startfile(path)
        else:
            messagebox.showerror("錯誤", "找不到指定的資料夾路徑。")

    def toggle_pause(self):
        if self.is_paused:
            self.is_paused = False
            self.pause_btn.config(text="暫停", bg="SystemButtonFace")
            self.update_progress_ui(self.progress_bar['value'], "繼續下載...", "blue")
        else:
            self.is_paused = True
            self.pause_btn.config(text="繼續", bg="#FFC107")
            self.update_progress_ui(self.progress_bar['value'], "下載已暫停", "orange")

    def cancel_download(self):
        if messagebox.askyesno("確認取消", "確定要取消目前的下載任務嗎？"):
            self.is_cancelled = True
            self.is_paused = False # 釋放可能在暫停狀態的迴圈
            self.update_progress_ui(self.progress_bar['value'], "正在終止下載程序，請稍候...", "red")
            self.cancel_btn.config(state="disabled")
            self.pause_btn.config(state="disabled")

    def on_cookie_browser_change(self, event=None):
        if self.cookie_browser.get() == "選擇 .txt 檔案...":
            f = filedialog.askopenfilename(filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
            if f:
                self.cookie_file_path.set(f)
                self.save_config()
                messagebox.showinfo("成功", f"已載入 Cookies 檔案：\n{os.path.basename(f)}")
            else:
                self.cookie_browser.set("無")
        else:
            self.save_config()

    def on_url_change(self, *args):
        url = self.url_entry.get().strip().lower()
        if not url:
            self.cookie_tip_label.config(text="(下載私人/限定影片時必備)", fg="gray")
            return
            
        if "facebook.com" in url or "fb.watch" in url or "reel" in url:
            self.cookie_tip_label.config(text="💡 偵測到 FB 網址：建議用「.txt 檔案」載入 Cookies 確保成功。", fg="#D32F2F")
        elif "instagram.com" in url:
            self.cookie_tip_label.config(text="💡 偵測到 IG 網址：建議開啟 Cookies 支援。", fg="#C2185B")
        elif "youtube.com" in url or "youtu.be" in url:
            self.cookie_tip_label.config(text="💡 YouTube 影片通常選「無」即可順暢下載。", fg="#388E3C")
        elif "douyin.com" in url or "tiktok.com" in url:
            self.cookie_tip_label.config(text="⚠️ 抖音強制要求 Cookies 解析：請務必載入 .txt 檔案！", fg="#D32F2F")
        else:
            self.cookie_tip_label.config(text="(下載私人/限定影片時必備)", fg="gray")

    def show_fb_tutorial(self):
        tutorial = (
            "【Facebook 私人/限動影片下載指南】\n\n"
            "1. 在瀏覽器安裝外掛「Get cookies.txt LOCALLY」。\n"
            "2. 在瀏覽器打開 FB 並確認已登入。\n"
            "3. 點擊外掛圖示，點選「Export」下載 .txt 檔案。\n"
            "4. 在本程式「Cookies 來源」選「選擇 .txt 檔案...」並載入該檔。\n"
            "5. 貼上網址後即可解析下載。\n\n"
            "※ 提示：使用 .txt 檔案不需要關閉瀏覽器，且成功率最高！"
        )
        messagebox.showinfo("下載教學", tutorial)

    def on_format_change(self):
        self.update_quality_options()
        self.save_config()

    def update_quality_options(self):
        fmt = self.format_choice.get()
        if fmt in ["mp4", "mkv"]:
            options = ["最高畫質 (自動)", "1080p", "720p", "480p", "360p"]
        elif fmt == "wav":
            options = ["無損音質 (WAV)"]
        else: # mp3
            options = ["最高音質 (320k)", "標準音質 (192k)", "普通音質 (128k)"]
        self.quality_combo['values'] = options
        self.quality_combo.current(0)

    def browse_folder(self):
        folder = filedialog.askdirectory(initialdir=self.download_path.get())
        if folder:
            self.download_path.set(folder)
            self.save_config()

    def update_ytdlp(self):
        # 如果是打包好的 exe 版本，yt-dlp 已經被封裝在裡面，無法透過 pip 單獨更新
        if getattr(sys, 'frozen', False):
            messagebox.showinfo("提示", "您目前使用的是免安裝執行檔版本，yt-dlp 下載核心已直接整合於主程式中。\n\n如需更新下載核心，請直接點選旁邊的「檢查主程式更新」按鈕即可！")
            return
            
        self.update_progress_ui(0, "正在更新 yt-dlp... 請稍候", "orange")
        def run_update():
            result = os.system(f"{sys.executable} -m pip install -U yt-dlp")
            if result == 0:
                self.root.after(0, lambda: messagebox.showinfo("更新成功", "yt-dlp 已更新至最新版！"))
                self.root.after(0, lambda: self.update_progress_ui(0, "準備就緒", "blue"))
            else:
                self.root.after(0, lambda: self.update_progress_ui(0, "更新失敗", "red"))
        threading.Thread(target=run_update, daemon=True).start()

    def check_ffmpeg_environment(self):
        ffmpeg_exe = os.path.join(self.app_dir, "ffmpeg.exe")
        ffprobe_exe = os.path.join(self.app_dir, "ffprobe.exe")
        
        if os.path.exists(ffmpeg_exe) and os.path.exists(ffprobe_exe):
            return
            
        def download_ffmpeg():
            self.root.after(0, lambda: self.update_progress_ui(0, "首次啟動：準備下載 FFmpeg 元件...", "orange"))
            self.root.after(0, lambda: self.download_btn.config(state="disabled"))
            self.root.after(0, lambda: self.analyze_btn.config(state="disabled"))
            try:
                def reporthook(blocknum, blocksize, totalsize):
                    if totalsize > 0:
                        readsofar = blocknum * blocksize
                        percent = min(100.0, (readsofar / totalsize) * 100)
                        self.root.after(0, lambda: self.update_progress_ui(percent, f"首次啟動：正在下載 FFmpeg 元件... ({percent:.1f}%)", "orange"))

                url = "https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip"
                zip_path = os.path.join(self.app_dir, "ffmpeg.zip")
                urllib.request.urlretrieve(url, zip_path, reporthook=reporthook)
                
                self.root.after(0, lambda: self.update_progress_ui(100, "下載完成，正在提取元件 (這需要幾十秒，請耐心等候)...", "orange"))
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    bin_path = None
                    for name in zip_ref.namelist():
                        if name.endswith('bin/ffmpeg.exe'):
                            bin_path = os.path.dirname(name)
                            break
                    if bin_path:
                        for exe in ['ffmpeg.exe', 'ffprobe.exe']:
                            source = f"{bin_path}/{exe}"
                            target = os.path.join(self.app_dir, exe)
                            with zip_ref.open(source) as zf, open(target, 'wb') as f:
                                shutil.copyfileobj(zf, f)
                self.root.after(0, lambda: self.update_progress_ui(0, "元件配置完成，可以開始使用！", "green"))
            except Exception as e:
                err_str = str(e)
                self.root.after(0, lambda: messagebox.showerror("錯誤", f"FFmpeg 下載失敗：\n{err_str}"))
                self.root.after(0, lambda: self.update_progress_ui(0, "環境不完整，可能無法進行影片轉檔", "red"))
            finally:
                if os.path.exists(zip_path):
                    try:
                        os.remove(zip_path)
                    except:
                        pass
                self.root.after(0, lambda: self.download_btn.config(state="normal" if self.video_info else "disabled"))
                self.root.after(0, lambda: self.analyze_btn.config(state="normal"))

        threading.Thread(target=download_ffmpeg, daemon=True).start()

    def check_app_update(self, is_auto=False):
        if not is_auto:
            self.update_progress_ui(0, "正在檢查主程式更新...", "blue")
        def run_check():
            try:
                req = urllib.request.Request(f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest", headers={'User-Agent': 'Mozilla/5.0'})
                response = urllib.request.urlopen(req, timeout=10)
                data = json.loads(response.read().decode('utf-8'))
                latest_version = data.get("tag_name", "").replace("v", "")
                
                if not latest_version:
                    if not is_auto:
                        self.root.after(0, lambda: self.update_progress_ui(0, "無法取得版本資訊", "red"))
                    return
                    
                if latest_version != APP_VERSION:
                    assets = data.get("assets", [])
                    download_url = None
                    for asset in assets:
                        if asset.get("name") == "CYT_YTDL.zip":
                            download_url = asset.get("browser_download_url")
                            break
                            
                    if download_url:
                        self.root.after(0, lambda: self.prompt_update(latest_version, download_url))
                    else:
                        if not is_auto:
                            self.root.after(0, lambda: messagebox.showinfo("發現新版本", f"目前最新版本為 {latest_version}，但開發者尚未上傳執行檔。"))
                            self.root.after(0, lambda: self.update_progress_ui(0, "檢查完畢", "blue"))
                else:
                    if not is_auto:
                        self.root.after(0, lambda: messagebox.showinfo("檢查更新", "您目前使用的已經是最新版本！"))
                        self.root.after(0, lambda: self.update_progress_ui(0, "準備就緒", "blue"))
            except urllib.error.HTTPError as e:
                if not is_auto:
                    if e.code == 404:
                        self.root.after(0, lambda: messagebox.showinfo("檢查更新", "專案尚未發布任何版本 (Release)。"))
                        self.root.after(0, lambda: self.update_progress_ui(0, "無可用更新", "blue"))
                    else:
                        self.root.after(0, lambda: self.update_progress_ui(0, f"檢查失敗: {e}", "red"))
            except Exception as e:
                if not is_auto:
                    self.root.after(0, lambda: self.update_progress_ui(0, f"檢查失敗: {e}", "red"))
        
        threading.Thread(target=run_check, daemon=True).start()

    def prompt_update(self, latest_version, download_url):
        if messagebox.askyesno("發現新版本", f"發現新版本 v{latest_version}！\n是否要立即下載並更新？\n\n注意：更新程式後，需手動關閉並重新啟動，才會使用最新版本。"):
            self.perform_update(download_url)
        else:
            self.update_progress_ui(0, "已取消更新", "blue")

    def perform_update(self, download_url):
        self.download_btn.config(state="disabled")
        self.analyze_btn.config(state="disabled")
        self.update_progress_ui(0, "正在準備下載新版本...", "orange")
        
        def run_update():
            try:
                def reporthook(blocknum, blocksize, totalsize):
                    if totalsize > 0:
                        readsofar = blocknum * blocksize
                        percent = min(100.0, (readsofar / totalsize) * 100)
                        self.root.after(0, lambda: self.update_progress_ui(percent, f"正在下載新版本... ({percent:.1f}%)", "orange"))

                new_zip_name = "CYT_YTDL_update.zip"
                new_zip_path = os.path.join(self.app_dir, new_zip_name)
                new_exe_name = "CYT_YTDL_update.exe"
                new_exe_path = os.path.join(self.app_dir, new_exe_name)
                
                # 1. 下載 ZIP
                urllib.request.urlretrieve(download_url, new_zip_path, reporthook=reporthook)
                
                # 2. 解壓縮
                self.root.after(0, lambda: self.update_progress_ui(100, "下載完成，正在解壓縮元件...", "orange"))
                try:
                    with zipfile.ZipFile(new_zip_path, 'r') as zip_ref:
                        # 尋找 ZIP 內的 exe 與 程式說明.txt
                        for name in zip_ref.namelist():
                            lname = name.lower()
                            if lname.endswith('.exe'):
                                with zip_ref.open(name) as zf, open(new_exe_path, 'wb') as f:
                                    shutil.copyfileobj(zf, f)
                            elif "程式說明.txt" in name:
                                target_txt = os.path.join(self.app_dir, "程式說明.txt")
                                with zip_ref.open(name) as zf, open(target_txt, 'wb') as f:
                                    shutil.copyfileobj(zf, f)
                    # 刪除暫存的 ZIP
                    if os.path.exists(new_zip_path):
                        os.remove(new_zip_path)
                except Exception as e:
                    self.root.after(0, lambda: messagebox.showerror("錯誤", f"解壓縮更新檔失敗：\n{e}"))
                    return

                self.root.after(0, lambda: self.update_progress_ui(100, "新版本準備就緒！等待確認重啟...", "green"))
                
                def ask_restart():
                    if messagebox.askyesno("更新準備就緒", "新版本已下載完畢！\n\n需關閉程式後重新開啟，才會使用最新版本。\n\n請問是否立刻關閉程式？"):
                        if getattr(sys, 'frozen', False):
                            current_exe_path = sys.executable
                            old_exe_path = current_exe_path + ".old"
                            
                            try:
                                # 嘗試替換檔案
                                if os.path.exists(old_exe_path):
                                    try: os.remove(old_exe_path)
                                    except: pass
                                
                                os.rename(current_exe_path, old_exe_path)
                                os.rename(new_exe_path, current_exe_path)
                                
                                messagebox.showinfo("更新成功", "新版本已替換完成！\n\n請在關閉本視窗後，重新手動執行程式以使用最新版本。")
                                self.root.destroy()
                                return
                            except Exception as e:
                                messagebox.showerror("錯誤", f"替換檔案失敗，請檢查權限或嘗試手動更新：\n{e}")
                            except Exception as e:
                                messagebox.showerror("錯誤", f"替換檔案失敗，請檢查權限：\n{e}")
                                self.update_progress_ui(0, "更新失敗", "red")
                        else:
                            messagebox.showinfo("開發者模式", "您目前在開發環境下，請手動更新程式碼即可。")
                            self.update_progress_ui(0, "開發環境無需更新", "blue")
                    else:
                        if os.path.exists(new_exe_path):
                            try:
                                os.remove(new_exe_path)
                            except:
                                pass
                        self.update_progress_ui(0, "已取消更新安裝", "blue")
                        self.download_btn.config(state="normal" if self.video_info else "disabled")
                        self.analyze_btn.config(state="normal")
                        
                self.root.after(0, ask_restart)
                
            except Exception as e:
                err_str = str(e)
                self.root.after(0, lambda: messagebox.showerror("錯誤", f"更新失敗：\n{err_str}"))
                self.root.after(0, lambda: self.update_progress_ui(0, "更新失敗", "red"))
                self.root.after(0, lambda: self.download_btn.config(state="normal" if self.video_info else "disabled"))
                self.root.after(0, lambda: self.analyze_btn.config(state="normal"))

        threading.Thread(target=run_update, daemon=True).start()

    def clear_url(self):
        self.url_entry.delete(0, tk.END)
        self.title_label.config(text="請輸入網址並點選「解析網址」")
        for widget in self.list_frame.scrollable_frame.winfo_children():
            widget.destroy()
        self.list_frame.pack_forget()
        self.select_btn_frame.pack_forget()
        self.download_btn.config(state="disabled")
        self.video_info = None
        self.is_playlist = False
        self.update_progress_ui(0, "等待解析...", "blue")

    def start_analyze(self):
        if getattr(self, 'is_analyzing', False):
            self.is_analyzing = False
            self.analyze_btn.config(text="正在取消...", state="disabled")
            self.update_progress_ui(0, "正在中斷解析程序...", "orange")
            return

        url = self.url_entry.get().strip()
        if not url:
            messagebox.showwarning("警告", "請輸入 YouTube 網址！")
            return
            
        self.is_analyzing = True
        self.analyze_btn.config(text="停止解析", bg="#f44336", fg="white", activebackground="#d32f2f", activeforeground="white")
        self.download_btn.config(state="disabled")
        self.update_progress_ui(0, "正在解析網址與抓取標題，請稍候...", "blue")
        self.title_label.config(text="解析中...")
        self.list_frame.pack_forget()
        self.select_btn_frame.pack_forget()
        
        threading.Thread(target=self.process_analyze, args=(url,), daemon=True).start()

    def reset_analyze_button(self):
        self.is_analyzing = False
        self.analyze_btn.config(
            text="解析網址",
            state="normal",
            bg="#2196F3",
            fg="white",
            activebackground="#2196F3",
            activeforeground="white"
        )
        
    def process_analyze(self, url):
        import urllib.parse
        parsed_url = urllib.parse.urlparse(url)
        query_params = urllib.parse.parse_qs(parsed_url.query)
        if 'list' in query_params and 'youtube.com' in parsed_url.netloc:
            playlist_id = query_params['list'][0]
            # YT Mix (合輯) 的清單 ID 通常以 RD 開頭，這種清單不能直接轉換為 /playlist?list= 否則會報錯
            if not playlist_id.startswith("RD"):
                url = f"https://www.youtube.com/playlist?list={playlist_id}"
                
        # 預解析網址：追蹤重定向 (處理 FB share 或短網址)
        if "facebook.com" in url or "douyin.com" in url or "t.co" in url or "bit.ly" in url:
            try:
                headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'}
                req = urllib.request.Request(url, headers=headers)
                with urllib.request.urlopen(req, timeout=5) as response:
                    url = response.geturl()
                # 重新解析 query 以便後續處理
                parsed_url = urllib.parse.urlparse(url)
                query_params = urllib.parse.parse_qs(parsed_url.query)
            except:
                pass

        # 抖音網址優化處理 (強化參數提取)
        if "douyin.com" in url:
            # 優先搜尋 modal_id 參數
            match = re.search(r'modal_id=(\d+)', url)
            if match:
                url = f"https://www.douyin.com/video/{match.group(1)}"
            elif "/jingxuan" in url or "/video/" in url:
                # 確保網址格式正確
                pass

        ydl_opts = {
            'extract_flat': 'in_playlist', # 單一影片改用深層解析以應對加密，清單則用扁平快速解析
            'quiet': True,
            'no_warnings': True,
            'nocheckcertificate': True,
            'socket_timeout': 15,
            'user_agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
            'logger': CancelLogger(self),
        }
        
        # 抖音額外優化 (深度偽裝)
        if "douyin.com" in url:
            ydl_opts['http_headers'] = {
                'Referer': url,
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
                'Accept-Language': 'zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7',
                'Sec-Ch-Ua': '"Chromium";v="124", "Google Chrome";v="124", "Not-A.Brand";v="99"',
                'Sec-Ch-Ua-Mobile': '?0',
                'Sec-Ch-Ua-Platform': '"Windows"',
                'Sec-Fetch-Dest': 'document',
                'Sec-Fetch-Mode': 'navigate',
                'Sec-Fetch-Site': 'same-origin',
                'Sec-Fetch-User': '?1',
                'Upgrade-Insecure-Requests': '1',
            }

        # 增加 Cookies 支援
        browser_choice = self.cookie_browser.get()
        use_cookies = False
        if browser_choice == "選擇 .txt 檔案...":
            cookie_file = self.cookie_file_path.get()
            if os.path.exists(cookie_file):
                ydl_opts['cookiefile'] = cookie_file
                use_cookies = True
        elif browser_choice != "無":
            ydl_opts['cookiesfrombrowser'] = (browser_choice,)
            use_cookies = True

        try:
            try:
                with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                    info = ydl.extract_info(url, download=False)
            except Exception as e:
                # 智慧降級：如果帶 Cookies 失敗，嘗試不帶 Cookies 重解析
                if use_cookies:
                    temp_opts = ydl_opts.copy()
                    if 'cookiefile' in temp_opts: del temp_opts['cookiefile']
                    if 'cookiesfrombrowser' in temp_opts: del temp_opts['cookiesfrombrowser']
                    with yt_dlp.YoutubeDL(temp_opts) as ydl:
                        info = ydl.extract_info(url, download=False)
                        # 如果無 Cookies 成功了，發送提示
                        self.root.after(0, lambda: self.update_progress_ui(0, "Cookies 解析受阻，已自動切換至相容模式解析成功", "orange"))
                else:
                    raise e
                    
            self.video_info = info
            
            if 'entries' in info:
                self.is_playlist = True
                entries = list(info.get('entries') or [])
                total = len(entries)
                self.playlist_all_entries = entries
                self.playlist_all_vars = [tk.BooleanVar(value=True) for _ in range(total)]
                self.playlist_status_labels = [None] * total
                self.playlist_current_page = 0
                
                if total > 50:
                    def ask_playlist_action():
                        # 建立自定義彈窗以支援「1」與「2」按鈕
                        dialog = tk.Toplevel(self.root)
                        dialog.title("播放清單處理方式")
                        dialog.geometry("450x250")
                        dialog.resizable(False, False)
                        dialog.transient(self.root)
                        dialog.grab_set()
                        
                        # 讓視窗置中
                        dialog.update_idletasks()
                        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (dialog.winfo_width() // 2)
                        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (dialog.winfo_height() // 2)
                        dialog.geometry(f"+{x}+{y}")

                        msg = f"偵測到龐大的播放清單 (共 {total} 部影片)！\n\n請選擇後續動作：\n\n【 1 】載入前 50 筆清單讓我手動勾選。\n【 2 】分頁顯示全部清單進行勾選。\n【取消】取消解析。"
                        tk.Label(dialog, text=msg, justify="left", font=("Microsoft JhengHei", 11), padx=20, pady=20).pack()

                        btn_frame = tk.Frame(dialog)
                        btn_frame.pack(side="bottom", pady=20)

                        def on_choice(choice):
                            dialog.destroy()
                            if choice in [1, 2]:
                                if choice == 1:
                                    self.playlist_all_entries = self.playlist_all_entries[:50]
                                    self.playlist_all_vars = self.playlist_all_vars[:50]
                                    self.playlist_status_labels = self.playlist_status_labels[:50]
                                self.root.after(0, lambda: self.show_playlist(0))
                            else:
                                self.update_progress_ui(0, "已取消解析", "blue")
                                self.analyze_btn.config(state="normal")
                                self.title_label.config(text="請輸入網址並點選「解析網址」")

                        tk.Button(btn_frame, text=" 1 ", width=10, command=lambda: on_choice(1), font=("Microsoft JhengHei", 10, "bold"), bg="#2196F3", fg="white").pack(side="left", padx=10)
                        tk.Button(btn_frame, text=" 2 ", width=10, command=lambda: on_choice(2), font=("Microsoft JhengHei", 10, "bold"), bg="#4CAF50", fg="white").pack(side="left", padx=10)
                        tk.Button(btn_frame, text="取消", width=10, command=lambda: on_choice(0)).pack(side="left", padx=10)

                    self.root.after(0, ask_playlist_action)
                else:
                    self.root.after(0, lambda: self.show_playlist(0))
            else:
                self.is_playlist = False
                self.playlist_status_labels = []
                # 優先從多個欄位尋找標題
                title = info.get('title') or info.get('fulltitle') or f"影片_{info.get('id', '未知')}"
                dur_str = self.format_duration(info.get('duration'))
                
                # 取得章節資料
                chapters = info.get('chapters') or []
                self.current_chapters = chapters
                
                # 優先從多個欄位尋找圖片
                thumb_url = info.get('thumbnail')
                if not thumb_url and info.get('thumbnails'):
                    thumb_url = info['thumbnails'][-1].get('url') # 抓最後一張通常最大
                    
                self.root.after(0, lambda: self.show_single_video(title, dur_str, thumb_url, chapters))
                
        except Exception as e:
            err_str = str(e)
            if "USER_CANCELLED" in err_str or not getattr(self, 'is_analyzing', False):
                self.root.after(0, lambda: self.update_progress_ui(0, "解析已中斷", "orange"))
                self.root.after(0, lambda: self.title_label.config(text="解析已手動取消。"))
                return
            # 抖音專屬中文提示優化
            if "douyin" in url and "Fresh cookies" in err_str:
                msg = (
                    "抖音目前加強了防抓取機制，偵測到您尚未登入或 IP 異常。\n\n"
                    "建議處理方式：\n"
                    "1. 請在瀏覽器中「重新播放一次」該影片（觸發網站更新驗證）。\n"
                    "2. 重新匯出最新的 .txt Cookies 檔案。\n"
                    "3. 確保程式採用『選擇 .txt 檔案』的方式載入後再進行解析。"
                )
                self.root.after(0, lambda: messagebox.showwarning("抖音解析提示", msg))
                self.root.after(0, lambda: self.update_progress_ui(0, "需要最新的 Cookies 才能解析", "red"))
            elif "login.php" in err_str or "facebook.com/login" in err_str:
                msg = (
                    "此影片可能為私人影片或限時動態，且目前的 Cookies 已失效或權限不足。\n\n"
                    "建議處理方式：\n"
                    "1. 請在瀏覽器中「重新播放一次」該影片（確認可正常收看）。\n"
                    "2. 重新匯出最新的 .txt Cookies 檔案後再嘗試解析。"
                )
                self.root.after(0, lambda: messagebox.showwarning("FB 解析提示", msg))
                self.root.after(0, lambda: self.update_progress_ui(0, "FB 登入失效，請更新 Cookies", "red"))
            else:
                self.root.after(0, lambda: self.title_label.config(text="解析失敗，請確認網址是否正確。"))
                self.root.after(0, lambda: self.update_progress_ui(0, "發生錯誤", "red"))
                self.root.after(0, lambda: messagebox.showerror("錯誤", f"解析失敗：\n{err_str}"))
        finally:
            self.root.after(0, lambda: self.reset_analyze_button())

    def format_duration(self, seconds):
        if not seconds:
            return ""
        m, s = divmod(int(seconds), 60)
        h, m = divmod(m, 60)
        return f" [{h}:{m:02d}:{s:02d}]" if h else f" [{m:02d}:{s:02d}]"

    def show_playlist_summary(self, title, entries):
        self.title_label.config(text=f"【播放清單】\n{title} (共 {len(entries)} 部影片)")
        for widget in self.list_frame.scrollable_frame.winfo_children():
            widget.destroy()
            
        self.list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.select_btn_frame.pack_forget()
        
        row_frame = tk.Frame(self.list_frame.scrollable_frame, pady=5)
        row_frame.pack(fill="x", anchor="w")
        
        txt_label = tk.Label(row_frame, text=f"💡 為避免介面卡頓，已隱藏清單明細。\n\n共有 {len(entries)} 部影片已準備就緒！\n請確認下方的「格式」與「儲存資料夾」無誤後，點擊「開始下載」即可自動下載全集。", justify="left", wraplength=500, font=("Microsoft JhengHei", 11, "bold"), fg="#2196F3")
        txt_label.pack(side="left", anchor="w", padx=20, pady=20)
        
        self.playlist_entries = entries
        self.playlist_vars = []  # 空的 var 代表總結模式 (全選)
        self.download_btn.config(state="normal")
        self.update_progress_ui(0, "解析完成！點擊「開始下載」以下載全集", "green")

    def show_single_video(self, title, dur_str, thumb_url, chapters=None):
        self.title_label.config(text="【單一影片解析結果】")
        for widget in self.list_frame.scrollable_frame.winfo_children():
            widget.destroy()
            
        # 有章節時需要更多高度
        canvas_height = 200 if chapters else 150
        self.list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.list_frame.canvas.config(height=canvas_height)
        try:
            self.list_frame.scrollbar.pack_forget() # 隱藏捲軸
        except:
            pass
        
        row_frame = tk.Frame(self.list_frame.scrollable_frame, pady=10)
        row_frame.pack(fill="x", anchor="w")
        
        thumb_label = tk.Label(row_frame, text="無圖片", bg="#e0e0e0", width=14, height=3)
        thumb_label.pack(side="left", padx=5)
        
        txt_label = tk.Label(row_frame, text=f"{title}\n時間: {dur_str.strip() if dur_str else '未知'}", justify="left", wraplength=500, font=("Microsoft JhengHei", 10))
        txt_label.pack(side="left", anchor="w", padx=10)
        
        if HAS_PIL and thumb_url:
            # 處理 Bilibili 可能缺少的協定頭
            if thumb_url.startswith("//"):
                thumb_url = "https:" + thumb_url
            threading.Thread(target=self.load_thumbnail, args=(thumb_url, thumb_label), daemon=True).start()
        
        # 若偵測到章節，顯示分割選項
        if chapters:
            chapter_frame = tk.Frame(self.list_frame.scrollable_frame, pady=5, padx=5)
            chapter_frame.pack(fill="x", anchor="w")
            
            self.split_by_chapters.set(False)  # 每次解析重置
            chk = tk.Checkbutton(
                chapter_frame,
                text=f"✂️ 依章節分割為多個檔案（偵測到 {len(chapters)} 個章節）",
                variable=self.split_by_chapters,
                font=("Microsoft JhengHei", 10, "bold"),
                fg="#7B1FA2",
                command=self._on_chapter_split_toggle
            )
            chk.pack(side="left", padx=5)
            
            # 章節清單預覽
            self.chapter_preview_label = tk.Label(
                self.list_frame.scrollable_frame,
                text=self._build_chapter_preview(chapters),
                justify="left",
                font=("Microsoft JhengHei", 8),
                fg="#555555",
                wraplength=700
            )
            self.chapter_preview_label.pack(fill="x", padx=15, pady=(0, 5))
        
        self.update_progress_ui(0, "解析完成！請確認資訊後點擊「開始下載」", "green")
        self.download_btn.config(state="normal")
        
        # 強制更新捲動區域，解決預覽消失問題
        self.root.after(100, lambda: self.list_frame.canvas.configure(scrollregion=self.list_frame.canvas.bbox("all")))

    def _build_chapter_preview(self, chapters):
        """建立章節預覽文字"""
        parts = []
        for i, ch in enumerate(chapters[:5]):  # 最多預覽前5個
            start = int(ch.get('start_time', 0))
            h, rem = divmod(start, 3600)
            m, s = divmod(rem, 60)
            t = f"{h:02d}:{m:02d}:{s:02d}" if h else f"{m:02d}:{s:02d}"
            parts.append(f"{i+1}. {t} {ch.get('title', '')}")
        if len(chapters) > 5:
            parts.append(f"... 共 {len(chapters)} 個章節")
        return "  |  ".join(parts)

    def _on_chapter_split_toggle(self):
        """章節分割勾選框切換時的提示"""
        if self.split_by_chapters.get():
            self.update_progress_ui(0, f"✅ 下載後將自動分割為 {len(self.current_chapters)} 個章節檔案", "purple")
        else:
            self.update_progress_ui(0, "解析完成！請確認資訊後點擊「開始下載」", "green")


    def show_playlist(self, page=0):
        title = self.video_info.get('title', '播放清單')
        total_items = len(self.playlist_all_entries)
        total_pages = math.ceil(total_items / 50)
        self.playlist_current_page = page
        
        # 當總數超過 50 筆且具有分頁時，自動動態顯示「本頁全部勾選」與「本頁全部取消」
        if total_items > 50:
            self.select_page_btn.pack(side="left", padx=5, after=self.deselect_all_btn)
            self.deselect_page_btn.pack(side="left", padx=5, after=self.select_page_btn)
        else:
            self.select_page_btn.pack_forget()
            self.deselect_page_btn.pack_forget()
        
        self.title_label.config(text=f"【播放清單】\n{title} (共 {total_items} 部影片)")
        self.page_label.config(text=f"第 {page + 1} / {total_pages} 頁")
        
        # 控制翻頁按鈕可用性
        self.prev_btn.config(state="normal" if page > 0 else "disabled")
        self.next_btn.config(state="normal" if page < total_pages - 1 else "disabled")
        
        for widget in self.list_frame.scrollable_frame.winfo_children():
            widget.destroy()
            
        self.select_btn_frame.pack(pady=3)
        self.list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        try:
            self.list_frame.scrollbar.pack(side="right", fill="y") # 確保捲軸在清單模式下出現
        except:
            pass
        
        start_idx = page * 50
        end_idx = min(start_idx + 50, total_items)
        
        # 開始漸進式渲染
        self.update_progress_ui(0, "正在載入清單項目...", "blue")
        self._render_batch(start_idx, end_idx, start_idx)

    def _render_batch(self, start, end, current):
        # 每次渲染 5 筆
        batch_size = 5
        batch_end = min(current + batch_size, end)
        
        for i in range(current, batch_end):
            entry = self.playlist_all_entries[i]
            var = self.playlist_all_vars[i]
            
            row_frame = tk.Frame(self.list_frame.scrollable_frame, pady=3)
            row_frame.pack(fill="x", anchor="w")
            
            chk = tk.Checkbutton(row_frame, variable=var, command=self.update_selection_count)
            chk.pack(side="left", padx=5)
            
            status_lbl = tk.Label(row_frame, text="⏳ 等待下載", fg="gray", font=("Microsoft JhengHei", 9, "bold"))
            status_lbl.pack(side="left", padx=5)
            if hasattr(self, 'playlist_status_labels') and len(self.playlist_status_labels) > i:
                self.playlist_status_labels[i] = status_lbl
                
            thumb_label = tk.Label(row_frame, text="等候中", bg="#e0e0e0", width=14, height=3)
            thumb_label.pack(side="left", padx=5)
            
            dur_str = self.format_duration(entry.get('duration'))
            title_text = entry.get('title', f'影片 {i+1}')
            txt_label = tk.Label(row_frame, text=f"{i+1}. {title_text}\n時間: {dur_str.strip() if dur_str else '未知'}", justify="left", wraplength=450, font=("Microsoft JhengHei", 10))
            txt_label.pack(side="left", anchor="w")
            
            if HAS_PIL:
                url = entry.get('thumbnail')
                if not url and entry.get('thumbnails'):
                    url = entry['thumbnails'][0].get('url')
                if url:
                    threading.Thread(target=self.load_thumbnail, args=(url, thumb_label), daemon=True).start()

        self.update_selection_count()
        
        if batch_end < end:
            # 繼續下一批
            self.root.after(10, lambda: self._render_batch(start, end, batch_end))
        else:
            self.update_progress_ui(0, "解析完成！您可以切換分頁進行勾選。", "green")
            self.download_btn.config(state="normal")

    def prev_page(self):
        if self.playlist_current_page > 0:
            self.show_playlist(self.playlist_current_page - 1)

    def next_page(self):
        total_pages = math.ceil(len(self.playlist_all_entries) / 50)
        if self.playlist_current_page < total_pages - 1:
            self.show_playlist(self.playlist_current_page + 1)

    def load_thumbnail(self, url, label):
        try:
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            raw_data = urllib.request.urlopen(req, timeout=5).read()
            im = Image.open(io.BytesIO(raw_data))
            im.thumbnail((100, 56))
            photo = ImageTk.PhotoImage(im)
            self.root.after(0, lambda: self._set_image(label, photo))
        except Exception:
            pass
            
    def _set_image(self, label, photo):
        label.config(image=photo, text="", width=100, height=56)
        label.image = photo

    def update_progress_ui(self, value, text, color="blue"):
        self.progress_bar['value'] = value
        self.status_label.config(text=text, fg=color)

    def progress_hook(self, d):
        # 攔截下載封包，實現暫停與取消
        while self.is_paused:
            if self.is_cancelled:
                raise ValueError("USER_CANCELLED")
            time.sleep(0.5)
            
        if self.is_cancelled:
            raise ValueError("USER_CANCELLED")

        if d['status'] == 'downloading':
            downloaded = d.get('downloaded_bytes', 0)
            total = d.get('total_bytes') or d.get('total_bytes_estimate', 0)
            percent_val = (downloaded / total * 100) if total > 0 else 0.0
            
            ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')
            percent_str = ansi_escape.sub('', d.get('_percent_str', f'{percent_val:.1f}%')).strip()
            speed = ansi_escape.sub('', d.get('_speed_str', 'N/A')).strip()
            eta = ansi_escape.sub('', d.get('_eta_str', 'N/A')).strip()
            
            threads_limit = self.threads_choice.get() if hasattr(self, 'threads_choice') else 1
            if self.is_playlist and threads_limit > 1:
                pass
            else:
                self.root.after(0, lambda: self.update_progress_ui(percent_val, f"下載進度: {percent_str} (速度: {speed}, 剩餘: {eta})", "blue"))
            
        elif d['status'] == 'finished':
            threads_limit = self.threads_choice.get() if hasattr(self, 'threads_choice') else 1
            if not (self.is_playlist and threads_limit > 1):
                self.root.after(0, lambda: self.update_progress_ui(100.0, "單檔下載完成！正在合併影像或轉檔... (此階段無法暫停)", "orange"))

    def make_progress_hook(self, idx):
        def hook(d):
            self.item_progress_hook(idx, d)
        return hook

    def item_progress_hook(self, idx, d):
        while self.is_paused:
            if self.is_cancelled:
                raise ValueError("USER_CANCELLED")
            time.sleep(0.5)
            
        if self.is_cancelled:
            raise ValueError("USER_CANCELLED")

        if d['status'] == 'downloading':
            downloaded = d.get('downloaded_bytes', 0)
            total = d.get('total_bytes') or d.get('total_bytes_estimate', 0)
            percent_val = (downloaded / total * 100) if total > 0 else 0.0
            
            if self.is_playlist and hasattr(self, 'playlist_status_labels') and len(self.playlist_status_labels) > idx:
                lbl = self.playlist_status_labels[idx]
                if lbl and lbl.winfo_exists():
                    self.root.after(0, lambda: lbl.config(text=f"⚡ 下載中 ({percent_val:.1f}%)", fg="blue"))
                    
            threads_limit = self.threads_choice.get() if hasattr(self, 'threads_choice') else 1
            if self.is_playlist and threads_limit == 1:
                ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')
                percent_str = ansi_escape.sub('', d.get('_percent_str', f'{percent_val:.1f}%')).strip()
                speed = ansi_escape.sub('', d.get('_speed_str', 'N/A')).strip()
                eta = ansi_escape.sub('', d.get('_eta_str', 'N/A')).strip()
                cur_idx = getattr(self, 'current_download_index', 0) + 1
                tot_cnt = getattr(self, 'total_download_count', 1)
                self.root.after(0, lambda: self.update_progress_ui(
                    percent_val, 
                    f"({cur_idx}/{tot_cnt}) 下載進度: {percent_str} (速度: {speed}, 剩餘: {eta})", 
                    "blue"
                ))
            
        elif d['status'] == 'finished':
            if self.is_playlist and hasattr(self, 'playlist_status_labels') and len(self.playlist_status_labels) > idx:
                lbl = self.playlist_status_labels[idx]
                if lbl and lbl.winfo_exists():
                    self.root.after(0, lambda: lbl.config(text="🔄 合併轉檔中...", fg="orange"))

    def start_download(self):
        save_dir = self.download_path.get()
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)
            
        fmt = self.format_choice.get()
        quality = self.quality_combo.get()
        
        urls_to_download = []
        if self.is_playlist:
            selected_indices = [i for i, var in enumerate(self.playlist_all_vars) if var.get()]
                
            if not selected_indices:
                messagebox.showwarning("提示", "請至少在清單中勾選一部影片！")
                return
            
            entries = self.playlist_all_entries
            for i in selected_indices:
                vid_url = entries[i].get('url') or entries[i].get('webpage_url')
                if not vid_url:
                    vid_id = entries[i].get('id')
                    if vid_id:
                        vid_url = f"https://www.youtube.com/watch?v={vid_id}"
                if vid_url:
                    urls_to_download.append((i, vid_url))
        else:
            urls_to_download.append((0, self.url_entry.get().strip()))

        self.download_btn.config(state="disabled")
        self.analyze_btn.config(state="disabled")
        
        # 啟用控制按鈕並重置狀態
        self.is_cancelled = False
        self.is_paused = False
        self.pause_btn.config(state="normal", text="暫停", bg="SystemButtonFace")
        self.cancel_btn.config(state="normal")
        self.update_progress_ui(0, "準備開始下載...", "blue")
        
        threading.Thread(target=self.process_download, args=(urls_to_download, save_dir, fmt, quality), daemon=True).start()

    def process_download(self, items, save_dir, fmt, quality):
        threads_limit = self.threads_choice.get() if hasattr(self, 'threads_choice') else 1
        total = len(items)
        
        self.current_download_index = 0
        self.total_download_count = total
        self._downloaded_filepath = None
        
        # 先將所有被勾選的項目之狀態標記為「等待下載」
        if self.is_playlist and hasattr(self, 'playlist_status_labels'):
            for idx, _ in items:
                if idx < len(self.playlist_status_labels):
                    lbl = self.playlist_status_labels[idx]
                    if lbl and lbl.winfo_exists():
                        lbl.config(text="⏳ 等待下載", fg="gray")

        # 初始化 ydl_opts
        if fmt in ["mp4", "mkv"]:
            if "最高畫質" in quality:
                format_str = 'bestvideo[ext=mp4]+bestaudio[ext=m4a]/best[ext=mp4]/best'
            elif "1080" in quality:
                format_str = 'bestvideo[height<=1080][ext=mp4]+bestaudio[ext=m4a]/best[height<=1080][ext=mp4]/best'
            elif "720" in quality:
                format_str = 'bestvideo[height<=720][ext=mp4]+bestaudio[ext=m4a]/best[height<=720][ext=mp4]/best'
            elif "480" in quality:
                format_str = 'bestvideo[height<=480][ext=mp4]+bestaudio[ext=m4a]/best[height<=480][ext=mp4]/best'
            else:
                format_str = 'bestvideo[height<=360][ext=mp4]+bestaudio[ext=m4a]/best[height<=360][ext=mp4]/best'
                
            ydl_opts = {
                'outtmpl': os.path.join(save_dir, '%(title)s [%(id)s].%(ext)s'),
                'format': format_str,
                'merge_output_format': fmt,
                'ffmpeg_location': self.app_dir,
                'color': 'no_color',
                'nocheckcertificate': True,
                'no_warnings': True,
                'socket_timeout': 30,
                'user_agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'referer': 'https://www.facebook.com/',
            }
        elif fmt == "wav":
            ydl_opts = {
                'outtmpl': os.path.join(save_dir, '%(title)s [%(id)s].%(ext)s'),
                'format': 'bestaudio/best',
                'postprocessors': [{
                    'key': 'FFmpegExtractAudio',
                    'preferredcodec': 'wav',
                }],
                'ffmpeg_location': self.app_dir,
                'color': 'no_color',
                'nocheckcertificate': True,
                'socket_timeout': 30,
                'user_agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            }
        else:
            if "320" in quality:
                kbps = '320'
            elif "192" in quality:
                kbps = '192'
            else:
                kbps = '128'
                
            ydl_opts = {
                'outtmpl': os.path.join(save_dir, '%(title)s [%(id)s].%(ext)s'),
                'format': 'bestaudio/best',
                'postprocessors': [{
                    'key': 'FFmpegExtractAudio',
                    'preferredcodec': 'mp3',
                    'preferredquality': kbps,
                }],
                'ffmpeg_location': self.app_dir,
                'color': 'no_color',
                'nocheckcertificate': True,
                'socket_timeout': 30,
                'user_agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            }
            
        ydl_opts['noplaylist'] = True

        # 定義單個項目的下載執行函數，回傳 (idx, success_status)
        def download_single_item(item_info):
            idx, url = item_info
            
            # 若已取消，直接回傳
            if self.is_cancelled:
                return idx, False
                
            # 建立該項目的專屬配置
            item_opts = ydl_opts.copy()
            item_opts['progress_hooks'] = [self.make_progress_hook(idx)]
            
            # 增加 Cookies 支援
            browser_choice = self.cookie_browser.get()
            use_cookies = False
            if browser_choice == "選擇 .txt 檔案...":
                cookie_file = self.cookie_file_path.get()
                if os.path.exists(cookie_file):
                    item_opts['cookiefile'] = cookie_file
                    use_cookies = True
            elif browser_choice != "無":
                item_opts['cookiesfrombrowser'] = (browser_choice,)
                use_cookies = True
                
            # 在 UI 上標記為「下載中」
            if self.is_playlist and hasattr(self, 'playlist_status_labels') and len(self.playlist_status_labels) > idx:
                lbl = self.playlist_status_labels[idx]
                if lbl and lbl.winfo_exists():
                    self.root.after(0, lambda: lbl.config(text="⚡ 連線下載中", fg="blue"))
                    
            # 連線與下載
            filepath = None
            try:
                with yt_dlp.YoutubeDL(item_opts) as ydl:
                    download_info = ydl.extract_info(url, download=True)
                    if download_info:
                        if 'requested_downloads' in download_info and download_info['requested_downloads']:
                            filepath = download_info['requested_downloads'][0].get('filepath')
                        else:
                            filepath = download_info.get('filepath') or ydl.prepare_filename(download_info)
                ret_code = 0
            except Exception as e:
                # 智慧容錯：如果是 YouTube 且 Cookies 下載階段報錯，嘗試不帶 Cookies 重試
                if use_cookies and ("youtube.com" in url or "youtu.be" in url) and ("cookie" in str(e).lower() or "dpapi" in str(e).lower() or "not copy" in str(e).lower()):
                    try:
                        temp_opts = item_opts.copy()
                        if "cookiesfrombrowser" in temp_opts: del temp_opts['cookiesfrombrowser']
                        if "cookiefile" in temp_opts: del temp_opts['cookiefile']
                        with yt_dlp.YoutubeDL(temp_opts) as ydl:
                            download_info = ydl.extract_info(url, download=True)
                            if download_info:
                                if 'requested_downloads' in download_info and download_info['requested_downloads']:
                                    filepath = download_info['requested_downloads'][0].get('filepath')
                                else:
                                    filepath = download_info.get('filepath') or ydl.prepare_filename(download_info)
                        ret_code = 0
                    except Exception as ex:
                        ret_code = -1
                        err_msg = str(ex)
                else:
                    ret_code = -1
                    err_msg = str(e)
                    
            if ret_code == 0 and filepath and os.path.exists(filepath):
                # 標記為「已完成」
                if self.is_playlist and hasattr(self, 'playlist_status_labels') and len(self.playlist_status_labels) > idx:
                    lbl = self.playlist_status_labels[idx]
                    if lbl and lbl.winfo_exists():
                        self.root.after(0, lambda: lbl.config(text="✅ 已完成", fg="green"))
                if not self.is_playlist:
                    self._downloaded_filepath = filepath
                return idx, True
            else:
                # 標記為「失敗」
                if self.is_playlist and hasattr(self, 'playlist_status_labels') and len(self.playlist_status_labels) > idx:
                    lbl = self.playlist_status_labels[idx]
                    if lbl and lbl.winfo_exists():
                        self.root.after(0, lambda: lbl.config(text="❌ 下載失敗", fg="red"))
                return idx, False

        import concurrent.futures
        success_count = 0
        
        try:
            # 建立 ThreadPoolExecutor
            with concurrent.futures.ThreadPoolExecutor(max_workers=threads_limit) as executor:
                # 提交所有任務
                futures = {executor.submit(download_single_item, item): item for item in items}
                
                # 監聽完成情況
                for future in concurrent.futures.as_completed(futures):
                    idx, success = future.result()
                    self.current_download_index += 1
                    if success:
                        success_count += 1
                        
                    # 更新整體進度 UI
                    overall_percent = (self.current_download_index / total) * 100.0
                    if self.is_playlist:
                        if threads_limit > 1:
                            self.root.after(0, lambda p=overall_percent: self.update_progress_ui(
                                p, 
                                f"下載進度: ({self.current_download_index}/{total}) 已完成 {success_count} 部 (同時下載: {threads_limit})", 
                                "blue"
                            ))
                        else:
                            self.root.after(0, lambda p=overall_percent: self.progress_bar.config(value=p))
                    else:
                        self.root.after(0, lambda p=overall_percent: self.progress_bar.config(value=p))
                        
            # 下載完畢後的二次確認與後處理
            if self.is_cancelled:
                self.root.after(0, lambda: self.update_progress_ui(0, "下載任務已取消", "red"))
                self.root.after(0, lambda: messagebox.showinfo("取消", "已成功取消下載任務。\n(未完成的暫存檔已保留，未來重新下載可自動接續進度)"))
            else:
                # 章節分割邏輯
                if (self.split_by_chapters.get() and self.current_chapters
                        and not self.is_playlist and self._downloaded_filepath):
                    chapters_copy = list(self.current_chapters)
                    filepath_copy = self._downloaded_filepath
                    title_copy = self.video_info.get('title', '影片') if self.video_info else '影片'
                    self.root.after(0, lambda: self.update_progress_ui(0, f"開始依章節分割 ({len(chapters_copy)} 個章節)...", "purple"))
                    self._split_by_chapters(filepath_copy, chapters_copy, title_copy, save_dir)
                else:
                    self.root.after(0, lambda: self.update_progress_ui(100.0, "所有任務皆已處理完成！", "green"))
                    self.root.after(0, lambda: messagebox.showinfo("成功", f"全部下載完畢！\n共成功下載 {success_count} / {total} 部影音。\n儲存至：\n{save_dir}"))
                    
        except Exception as e:
            err_str = str(e)
            if "USER_CANCELLED" in err_str:
                self.root.after(0, lambda: self.update_progress_ui(0, "下載任務已取消", "red"))
                self.root.after(0, lambda: messagebox.showinfo("取消", "已成功取消下載任務。\n(未完成的暫存檔已保留，未來重新下載可自動接續進度)"))
            else:
                self.root.after(0, lambda: self.update_progress_ui(0, "下載過程發生錯誤", "red"))
                self.root.after(0, lambda: messagebox.showerror("錯誤", f"下載失敗，可能是網路問題或影片遭版權封鎖：\n{err_str}"))
        finally:
            self.root.after(0, lambda: self.download_btn.config(state="normal"))
            self.root.after(0, lambda: self.analyze_btn.config(state="normal"))
            self.root.after(0, lambda: self.pause_btn.config(state="disabled", text="暫停", bg="SystemButtonFace"))
            self.root.after(0, lambda: self.cancel_btn.config(state="disabled"))
    def _sanitize_filename(self, name):
        """清除檔名中的非法字元"""
        return re.sub(r'[\\/:*?"<>|]', '_', name).strip()

    def _split_by_chapters(self, src_path, chapters, video_title, save_dir):
        """依章節將完整檔案分割為多個子檔案"""
        # 建立輸出子目錄
        safe_title = self._sanitize_filename(video_title)
        out_dir = os.path.join(save_dir, f"{safe_title}_chapters")
        os.makedirs(out_dir, exist_ok=True)
        
        ext = os.path.splitext(src_path)[1]  # 取得副檔名如 .mp4 .mp3
        ffmpeg_path = os.path.join(self.app_dir, 'ffmpeg.exe')
        total = len(chapters)
        errors = []
        
        for i, ch in enumerate(chapters):
            if self.is_cancelled:
                break
            start_sec = ch.get('start_time', 0)
            end_sec = ch.get('end_time', None)
            ch_title = self._sanitize_filename(ch.get('title', f'chapter_{i+1}'))
            out_name = f"{i+1:02d}_{ch_title}{ext}"
            out_path = os.path.join(out_dir, out_name)
            
            # 即時更新進度
            self.root.after(0, lambda idx=i: self.update_progress_ui(
                int(idx / total * 100),
                f"分割章節 {idx+1}/{total}：{chapters[idx].get('title', '')}",
                "purple"
            ))
            
            cmd = [ffmpeg_path, '-y', '-i', src_path, '-ss', str(start_sec)]
            if end_sec is not None:
                cmd += ['-to', str(end_sec)]
            cmd += ['-c', 'copy', out_path]
            
            try:
                result = subprocess.run(cmd, capture_output=True, text=True)
                if result.returncode != 0:
                    errors.append(f"章節 {i+1}: {result.stderr[-200:]}")
            except Exception as e:
                errors.append(f"章節 {i+1}: {str(e)}")
        
        # 完成後通知
        if errors:
            err_msg = '\n'.join(errors[:3])
            self.root.after(0, lambda: messagebox.showwarning(
                "部分章節分割失敗",
                f"以下章節分割時發生錯誤：\n{err_msg}\n\n成功的章節已儲存至：\n{out_dir}"
            ))
        else:
            self.root.after(0, lambda: self.update_progress_ui(100, f"章節分割完成！已分割 {total} 個章節", "green"))
            self.root.after(0, lambda: messagebox.showinfo(
                "章節分割完成",
                f"成功將影片分割為 {total} 個章節檔案！\n\n輸出資料夾：\n{out_dir}\n\n原始完整檔案已保留在：\n{save_dir}"
            ))

    def open_help_dialog(self):
        txt_path = os.path.join(self.app_dir, "使用說明.txt")
        md_path = os.path.join(self.app_dir, "使用說明.md")
        
        content = ""
        # 優先讀取 txt 說明，若不存在則嘗試讀取 md 說明 (便於開發與打包環境相容)
        if os.path.exists(txt_path):
            try:
                with open(txt_path, "r", encoding="utf-8") as f:
                    content = f.read()
            except Exception as e:
                messagebox.showerror("讀取失敗", f"讀取「使用說明.txt」失敗：{e}")
                return
        elif os.path.exists(md_path):
            try:
                with open(md_path, "r", encoding="utf-8") as f:
                    content = f.read()
            except Exception as e:
                messagebox.showerror("讀取失敗", f"讀取「使用說明.md」失敗：{e}")
                return
        else:
            messagebox.showerror("找不到說明檔", "在程式目錄中找不到「使用說明.txt」或「使用說明.md」檔案。")
            return
            
        # 建立使用說明 Toplevel 視窗
        help_win = tk.Toplevel(self.root)
        help_win.title("📖 CYT_網路影音下載器 - 使用說明")
        help_win.geometry("680x580")
        help_win.resizable(True, True)
        help_win.transient(self.root)
        help_win.grab_set()
        
        # 視窗置中
        help_win.update_idletasks()
        w_win, h_win = 680, 580
        x = self.root.winfo_x() + (self.root.winfo_width() - w_win) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - h_win) // 2
        help_win.geometry(f"{w_win}x{h_win}+{x}+{y}")
        
        # 視窗容器
        main_frame = tk.Frame(help_win, bg="#F5F5F5", padx=15, pady=15)
        main_frame.pack(fill="both", expand=True)
        
        # 標題
        title_lbl = tk.Label(main_frame, text="📖 CYT_網路影音下載器 使用說明書", font=("Microsoft JhengHei", 14, "bold"), bg="#F5F5F5", fg="#1976D2")
        title_lbl.pack(anchor="w", pady=(0, 10))
        
        # 文字展示區 (加上 Scrollbar)
        text_frame = tk.Frame(main_frame, bg="white", bd=1, relief="solid")
        text_frame.pack(fill="both", expand=True, pady=5)
        
        scrollbar = tk.Scrollbar(text_frame)
        scrollbar.pack(side="right", fill="y")
        
        text_area = tk.Text(text_frame, font=("微軟正黑體", 10), wrap="word", yscrollcommand=scrollbar.set, bd=0, padx=10, pady=10)
        text_area.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=text_area.yview)
        
        # 寫入內容並設為唯讀
        text_area.insert(tk.END, content)
        text_area.config(state="disabled")
        
        # 關閉按鈕
        close_btn = tk.Button(main_frame, text="關閉說明", command=help_win.destroy, font=("Microsoft JhengHei", 10, "bold"), bg="#9E9E9E", fg="white", relief="flat", padx=15, pady=5, cursor="hand2")
        close_btn.pack(pady=(10, 0))

    def open_feedback_dialog(self):
        # 建立表單 Toplevel 視窗
        dialog = tk.Toplevel(self.root)
        dialog.title("💡 使用者回饋表單")
        dialog.geometry("480x430")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 視窗置中
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (dialog.winfo_width() // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # 頂部裝飾條與文字
        top_bar = tk.Frame(dialog, bg="#FF9800", height=4)
        top_bar.pack(fill="x")
        
        tk.Label(dialog, text="💡 用戶問題回報與功能許願", font=("Microsoft JhengHei", 14, "bold"), fg="#333", pady=10).pack()
        
        # 表單容器
        form_frame = tk.Frame(dialog, padx=20, pady=10)
        form_frame.pack(fill="both", expand=True)
        
        # 姓名
        row1 = tk.Frame(form_frame)
        row1.pack(fill="x", pady=5)
        tk.Label(row1, text="您的稱呼：*", font=("Microsoft JhengHei", 10, "bold"), width=10, anchor="w").pack(side="left")
        name_entry = tk.Entry(row1, font=("Microsoft JhengHei", 10))
        name_entry.pack(side="left", fill="x", expand=True)
        
        # E-mail
        row2 = tk.Frame(form_frame)
        row2.pack(fill="x", pady=5)
        tk.Label(row2, text="聯絡信箱：", font=("Microsoft JhengHei", 10), width=10, anchor="w").pack(side="left")
        email_entry = tk.Entry(row2, font=("Microsoft JhengHei", 10))
        email_entry.pack(side="left", fill="x", expand=True)
        
        # 類別
        row3 = tk.Frame(form_frame)
        row3.pack(fill="x", pady=5)
        tk.Label(row3, text="回饋類別：", font=("Microsoft JhengHei", 10), width=10, anchor="w").pack(side="left")
        type_combo = ttk.Combobox(row3, values=["問題回報 🐛", "功能建議 💡", "其他 ✉️"], state="readonly", font=("Microsoft JhengHei", 9))
        type_combo.set("問題回報 🐛")
        type_combo.pack(side="left", fill="x", expand=True)
        
        # 內容
        row4 = tk.Frame(form_frame)
        row4.pack(fill="both", expand=True, pady=5)
        tk.Label(row4, text="詳細說明：*", font=("Microsoft JhengHei", 10, "bold"), width=10, anchor="nw").pack(side="left", pady=(3,0))
        content_text = tk.Text(row4, font=("Microsoft JhengHei", 10), height=8, wrap="word")
        content_text.pack(side="left", fill="both", expand=True)
        
        # 按鈕區
        btn_frame = tk.Frame(dialog, pady=15)
        btn_frame.pack(side="bottom", fill="x")
        
        def on_submit():
            name = name_entry.get().strip()
            email = email_entry.get().strip()
            feedback_type = type_combo.get()
            content = content_text.get("1.0", tk.END).strip()
            
            if not name:
                messagebox.showerror("錯誤", "「您的稱呼」為必填欄位！", parent=dialog)
                return
            if not content:
                messagebox.showerror("錯誤", "「詳細說明」為必填欄位！", parent=dialog)
                return
            if email and not re.match(r"[^@]+@[^@]+\.[^@]+", email):
                messagebox.showerror("錯誤", "請輸入格式正確的聯絡信箱！", parent=dialog)
                return
            
            confirm_msg = f"確認要送出回饋嗎？\n\n【您的稱呼】：{name}\n【聯絡信箱】：{email or '未提供'}\n【回饋類別】：{feedback_type}\n\n說明內容：\n{content}"
            if messagebox.askyesno("送出確認", confirm_msg, parent=dialog):
                self.send_feedback(name, email, feedback_type, content, dialog)
                
        tk.Button(btn_frame, text="🚀 送出回饋", bg="#4CAF50", fg="white", font=("Microsoft JhengHei", 11, "bold"), width=15, command=on_submit).pack(side="right", padx=30)
        tk.Button(btn_frame, text="取消", font=("Microsoft JhengHei", 11), width=10, command=dialog.destroy).pack(side="left", padx=30)

    def send_feedback(self, name, email, feedback_type, content, dialog):
        # 顯示發送中提示並鎖定按鈕
        dialog.title("正在發送...")
        self.update_progress_ui(0, "正在發送回饋資訊給開發者...", "blue")
        
        data = {
            "appName": "CYT_網路影音下載器",
            "userName": name,
            "email": email or "未提供",
            "type": feedback_type,
            "content": content,
            "os": "Windows"
        }
        
        def _send_thread():
            try:
                req_data = json.dumps(data).encode("utf-8")
                req = urllib.request.Request(
                    FEEDBACK_API_URL, 
                    data=req_data, 
                    headers={'Content-Type': 'application/json'},
                    method="POST"
                )
                
                with urllib.request.urlopen(req, timeout=10) as response:
                    res_body = response.read().decode("utf-8")
                    res_json = json.loads(res_body)
                    
                    if res_json.get("status") == "success":
                        self.root.after(0, lambda: self.update_progress_ui(0, "🎉 感謝您的寶貴回饋！送出成功！", "green"))
                        self.root.after(0, lambda: messagebox.showinfo("送出成功", "🎉 您的回饋已成功送達開發者！\n感謝您對 CYT_網路影音下載器的支持與建議！", parent=self.root))
                        self.root.after(0, dialog.destroy)
                    else:
                        err_msg = res_json.get("message", "未知錯誤")
                        self.root.after(0, lambda: self.update_progress_ui(0, f"❌ 送出失敗：{err_msg}", "red"))
                        self.root.after(0, lambda: messagebox.showerror("送出失敗", f"❌ 後台處理失敗：\n{err_msg}", parent=dialog))
            except Exception as e:
                err_msg = str(e)
                self.root.after(0, lambda: self.update_progress_ui(0, f"❌ 發送失敗：{err_msg}", "red"))
                self.root.after(0, lambda: messagebox.showerror("連線錯誤", f"❌ 無法連線至回饋後台伺服器！\n請檢查 API 網址是否正確且已公開發布。\n\n詳細錯誤資訊：\n{err_msg}", parent=dialog))
                
        threading.Thread(target=_send_thread, daemon=True).start()

class MCIPlayer:
    def __init__(self, alias="cyt_mp3_player"):
        try:
            self._mci = ctypes.windll.winmm.mciSendStringW
            self._get_error = ctypes.windll.winmm.mciGetErrorStringW
            self._available = True
        except Exception:
            self._available = False
        self._alias = alias
        self._is_open = False

    def _get_short_path(self, path):
        """MCI 對長路徑或含空格路徑支援較差，轉換為短路徑 (8.3 格式)"""
        buf = ctypes.create_unicode_buffer(512)
        ctypes.windll.kernel32.GetShortPathNameW(path, buf, 512)
        return buf.value

    def _send(self, cmd):
        if not self._available:
            return (0, "")
        buf = ctypes.create_unicode_buffer(256)
        res = self._mci(cmd, buf, 256, 0)
        return (res, buf.value.strip())

    def open(self, path):
        self.close()
        short_path = self._get_short_path(path)
        # 如果短路徑獲取失敗（例如檔案不存在），則使用原始路徑並加雙引號
        p = f'"{short_path}"' if short_path else f'"{path}"'
        
        res, _ = self._send(f'open {p} type mpegvideo alias {self._alias}')
        if res != 0:
            err_buf = ctypes.create_unicode_buffer(256)
            self._get_error(res, err_buf, 256)
            print(f"MCI Open Error: {err_buf.value}")
            return False
            
        self._send(f'set {self._alias} time format milliseconds')
        self._is_open = True
        return True

    def play(self):
        if self._is_open:
            self._send(f'play {self._alias}')

    def pause(self):
        if self._is_open:
            self._send(f'pause {self._alias}')

    def resume(self):
        if self._is_open:
            self._send(f'resume {self._alias}')

    def stop(self):
        if self._is_open:
            self._send(f'stop {self._alias}')
            self._send(f'seek {self._alias} to start')

    def seek(self, ms):
        if self._is_open:
            was_playing = self.get_mode() == "playing"
            self._send(f'seek {self._alias} to {int(ms)}')
            if was_playing:
                self._send(f'play {self._alias}')

    def get_position(self):
        try:
            _, val = self._send(f'status {self._alias} position')
            return int(val)
        except Exception:
            return 0

    def get_length(self):
        try:
            _, val = self._send(f'status {self._alias} length')
            return int(val)
        except Exception:
            return 0

    def get_mode(self):
        _, val = self._send(f'status {self._alias} mode')
        return val

    def set_volume(self, vol):
        """設定音量 0-1000；Windows MCI 使用 setaudio volume"""
        if self._is_open:
            self._send(f'setaudio {self._alias} volume to {int(vol)}')

    def close(self):
        if self._is_open:
            self._send(f'close {self._alias}')
            self._is_open = False


# ===========================================================================
# MP3TrimmerTab：MP3 裁剪工具的完整 UI 類別
# ===========================================================================

class MP3TrimmerTab:
    def __init__(self, parent, download_path_var):
        self.parent = parent
        self.download_path_var = download_path_var
        self.player = MCIPlayer()
        self.current_file = None
        self.total_ms = 0
        self.is_playing = False
        self.is_paused = False
        self.start_time_str = tk.StringVar(value="0:00")
        self.end_time_str = tk.StringVar(value="0:00")
        self._update_job = None
        self._seeking = False
        self._preview_mode = False  # 試聽標記區段模式
        self._loop_var = tk.BooleanVar(value=False)  # 循環播放開關
        self._build_ui()

    def _build_ui(self):
        # === 左側：檔案列表區 ===
        left_frame = tk.Frame(self.parent, width=210)
        left_frame.pack(side="left", fill="y", padx=(10, 5), pady=10)
        left_frame.pack_propagate(False)

        tk.Label(left_frame, text="MP3 檔案列表", font=("Microsoft JhengHei", 11, "bold")).pack(anchor="w")

        folder_frame = tk.Frame(left_frame)
        folder_frame.pack(fill="x", pady=3)
        self.folder_entry = tk.Entry(folder_frame, state="readonly", font=("Microsoft JhengHei", 8))
        self.folder_entry.pack(side="left", fill="x", expand=True)
        tk.Button(folder_frame, text="選擇", command=self._browse_folder, font=("Microsoft JhengHei", 8)).pack(side="left", padx=2)
        tk.Button(folder_frame, text="開啟", command=self._open_folder, font=("Microsoft JhengHei", 8)).pack(side="left", padx=2)

        self.file_listbox = tk.Listbox(left_frame, font=("Microsoft JhengHei", 9), selectmode="single", activestyle="dotbox")
        self.file_listbox.pack(fill="both", expand=True, pady=5)
        self.file_listbox.bind("<<ListboxSelect>>", self._on_file_select)

        tk.Button(left_frame, text="🔄 重新整理", command=self._refresh_list, font=("Microsoft JhengHei", 9)).pack(fill="x")

        # === 右側：控制區 ===
        right_frame = tk.Frame(self.parent)
        right_frame.pack(side="left", fill="both", expand=True, padx=(5, 10), pady=10)

        # 檔名顯示
        self.file_label = tk.Label(right_frame, text="尚未選取檔案", font=("Microsoft JhengHei", 10, "bold"),
                                   fg="#1565C0", wraplength=480, justify="left")
        self.file_label.pack(anchor="w", pady=(0, 5))

        # 播放控制按鈕（含跳轉±1s/±5s）
        ctrl_frame = tk.Frame(right_frame)
        ctrl_frame.pack(anchor="w", pady=3)
        tk.Button(ctrl_frame, text="⏮ -5s", command=lambda: self._seek_relative(-5000),
                  font=("Microsoft JhengHei", 9), width=5).pack(side="left", padx=1)
        tk.Button(ctrl_frame, text="◀ -1s", command=lambda: self._seek_relative(-1000),
                  font=("Microsoft JhengHei", 9), width=5).pack(side="left", padx=1)
        self.play_btn = tk.Button(ctrl_frame, text="▶ 播放", command=self._toggle_play,
                                  font=("Microsoft JhengHei", 11, "bold"), bg="#4CAF50", fg="white", width=8, state="disabled")
        self.play_btn.pack(side="left", padx=4)
        self.stop_btn = tk.Button(ctrl_frame, text="⏹ 停止", command=self._stop,
                                  font=("Microsoft JhengHei", 11), width=7, state="disabled")
        self.stop_btn.pack(side="left", padx=2)
        tk.Button(ctrl_frame, text="+1s ▶", command=lambda: self._seek_relative(1000),
                  font=("Microsoft JhengHei", 9), width=5).pack(side="left", padx=1)
        tk.Button(ctrl_frame, text="+5s ⏭", command=lambda: self._seek_relative(5000),
                  font=("Microsoft JhengHei", 9), width=5).pack(side="left", padx=1)
        self.time_label = tk.Label(ctrl_frame, text="00:00 / 00:00", font=("Microsoft JhengHei", 11), fg="#333")
        self.time_label.pack(side="left", padx=10)

        # 自訂 Canvas 進度條（顯示裁剪範圍色塊與起終點標記）
        canvas_outer = tk.Frame(right_frame, bg="#aaaaaa", pady=1)
        canvas_outer.pack(fill="x", pady=(4, 0))
        self.trim_canvas = tk.Canvas(canvas_outer, height=26, bg="#e0e0e0",
                                     highlightthickness=0, cursor="hand2")
        self.trim_canvas.pack(fill="both", expand=True)
        self.trim_canvas.bind("<ButtonPress-1>", self._canvas_click)
        self.trim_canvas.bind("<B1-Motion>", self._canvas_drag)
        self.trim_canvas.bind("<ButtonRelease-1>", self._canvas_release)
        self.trim_canvas.bind("<Configure>", lambda e: self._draw_trim_canvas())

        # 色彩圖例
        legend_frame = tk.Frame(right_frame)
        legend_frame.pack(anchor="w", pady=(2, 2))
        for color, label in [("#81C784", "裁剪範圍"), ("#1976D2", "起點"), ("#E64A19", "終點"), ("#EF5350", "播放位置")]:
            tk.Frame(legend_frame, bg=color, width=12, height=12).pack(side="left", padx=2)
            tk.Label(legend_frame, text=label, font=("Microsoft JhengHei", 8), fg="#555").pack(side="left", padx=(0, 8))

        ttk.Separator(right_frame, orient="horizontal").pack(fill="x", pady=5)

        # === 裁剪設定 ===
        trim_lf = tk.LabelFrame(right_frame, text="✂️ 裁剪設定", font=("Microsoft JhengHei", 10, "bold"), padx=10, pady=6)
        trim_lf.pack(fill="x", pady=3)

        # 試聽與循環播放
        preview_row = tk.Frame(trim_lf)
        preview_row.pack(fill="x", pady=(2, 5))
        self.preview_btn = tk.Button(preview_row, text="▶ 播放標記區段", command=self._preview_section,
                                     font=("Microsoft JhengHei", 10, "bold"), bg="#7B1FA2", fg="white", state="disabled", width=14)
        self.preview_btn.pack(side="left", padx=(0, 5))
        
        self.preview_toggle_btn = tk.Button(preview_row, text="▶ 播放 / 暫停", command=self._preview_toggle,
                                            font=("Microsoft JhengHei", 10, "bold"), bg="#673AB7", fg="white", state="disabled", width=12)
        self.preview_toggle_btn.pack(side="left", padx=5)
        
        tk.Checkbutton(preview_row, text="🔁 循環播放", variable=self._loop_var,
                       font=("Microsoft JhengHei", 10), fg="#4A148C").pack(side="left")
        self.duration_label = tk.Label(preview_row, text="預計長度：0秒", font=("Microsoft JhengHei", 10, "bold"), fg="#E91E63")
        self.duration_label.pack(side="left", padx=15)
        ttk.Separator(trim_lf, orient="horizontal").pack(fill="x", pady=(0, 4))

        # 起點
        start_row = tk.Frame(trim_lf)
        start_row.pack(fill="x", pady=3)
        tk.Label(start_row, text="起點：", font=("Microsoft JhengHei", 11), width=5).pack(side="left")
        self.start_entry = tk.Entry(start_row, textvariable=self.start_time_str, width=10, font=("Microsoft JhengHei", 11))
        self.start_entry.pack(side="left", padx=4)
        tk.Button(start_row, text="◀ -0.1s", command=lambda: self._adjust('start', -0.1),
                  font=("Microsoft JhengHei", 9), width=6).pack(side="left", padx=2)
        tk.Button(start_row, text="+0.1s ▶", command=lambda: self._adjust('start', +0.1),
                  font=("Microsoft JhengHei", 9), width=6).pack(side="left", padx=2)
        tk.Button(start_row, text="📍 標記目前位置", command=self._mark_start,
                  font=("Microsoft JhengHei", 10), bg="#1976D2", fg="white").pack(side="left", padx=8)

        # 終點
        end_row = tk.Frame(trim_lf)
        end_row.pack(fill="x", pady=3)
        tk.Label(end_row, text="終點：", font=("Microsoft JhengHei", 11), width=5).pack(side="left")
        self.end_entry = tk.Entry(end_row, textvariable=self.end_time_str, width=10, font=("Microsoft JhengHei", 11))
        self.end_entry.pack(side="left", padx=4)
        tk.Button(end_row, text="◀ -0.1s", command=lambda: self._adjust('end', -0.1),
                  font=("Microsoft JhengHei", 9), width=6).pack(side="left", padx=2)
        tk.Button(end_row, text="+0.1s ▶", command=lambda: self._adjust('end', +0.1),
                  font=("Microsoft JhengHei", 9), width=6).pack(side="left", padx=2)
        tk.Button(end_row, text="📍 標記目前位置", command=self._mark_end,
                  font=("Microsoft JhengHei", 10), bg="#E64A19", fg="white").pack(side="left", padx=8)

        ttk.Separator(right_frame, orient="horizontal").pack(fill="x", pady=8)

        # 輸出設定
        # 儲存資料夾 Row
        self.out_folder_row = tk.Frame(right_frame)
        self.out_folder_row.pack(fill="x", pady=3)
        tk.Label(self.out_folder_row, text="儲存資料夾：", font=("Microsoft JhengHei", 10, "bold"), width=12, anchor="w").pack(side="left")
        self.out_folder_var = tk.StringVar()
        self.out_folder_entry = tk.Entry(self.out_folder_row, textvariable=self.out_folder_var, font=("Microsoft JhengHei", 10), state="readonly")
        self.out_folder_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        tk.Button(self.out_folder_row, text="選擇", command=self._browse_out_folder, bg="#E91E63", fg="white", font=("Microsoft JhengHei", 9)).pack(side="left")

        out_frame = tk.Frame(right_frame)
        out_frame.pack(fill="x", pady=3)
        tk.Label(out_frame, text="新檔名：", font=("Microsoft JhengHei", 11)).pack(side="left")
        self.out_entry = tk.Entry(out_frame, font=("Microsoft JhengHei", 10), width=35)
        self.out_entry.pack(side="left", padx=5, fill="x", expand=True)
        tk.Label(out_frame, text=".mp3", font=("Microsoft JhengHei", 11)).pack(side="left")

        # 裁剪按鈕
        trim_btn_frame = tk.Frame(right_frame)
        trim_btn_frame.pack(pady=10)
        self.trim_btn = tk.Button(trim_btn_frame, text="✂️  裁剪並儲存新檔案", command=self._do_trim,
                                  font=("Microsoft JhengHei", 13, "bold"), bg="#E53935", fg="white",
                                  width=25, height=2, state="disabled")
        self.trim_btn.pack()

        self.trim_status = tk.Label(right_frame, text="", font=("Microsoft JhengHei", 10), fg="green")
        self.trim_status.pack()

        # 初始化載入下載資料夾
        self._refresh_list()

    def _browse_folder(self):
        folder = filedialog.askdirectory(initialdir=self.download_path_var.get())
        if folder:
            self._folder_path = folder
            self._update_folder_entry(folder)
            self._refresh_list()

    def _open_folder(self):
        folder = (getattr(self, '_folder_path', None) or self.download_path_var.get()) or ""
        if folder and os.path.exists(folder):
            os.startfile(folder)
        else:
            messagebox.showerror("錯誤", "找不到指定的資料夾路徑。")

    def _update_folder_entry(self, path):
        self.folder_entry.config(state="normal")
        self.folder_entry.delete(0, tk.END)
        self.folder_entry.insert(0, path)
        self.folder_entry.config(state="readonly")

    def _browse_out_folder(self):
        initial_dir = self.out_folder_var.get() or self.download_path_var.get()
        folder = filedialog.askdirectory(initialdir=initial_dir)
        if folder:
            is_writable = True
            test_file = os.path.join(folder, ".write_test")
            try:
                with open(test_file, "w") as f:
                    pass
                os.remove(test_file)
            except Exception:
                is_writable = False
            
            if not is_writable:
                messagebox.showerror("權限錯誤", "所選資料夾為唯讀或無寫入權限，請選擇其他儲存位置。")
            else:
                self.out_folder_var.set(folder)

    def _refresh_list(self):
        folder = (getattr(self, '_folder_path', None) or self.download_path_var.get()) or ""
        self._folder_path = folder
        self._update_folder_entry(folder)
        self.file_listbox.delete(0, tk.END)
        if not folder or not os.path.exists(folder):
            return
        mp3_files = sorted([f for f in os.listdir(folder) if f.lower().endswith('.mp3')])
        for f in mp3_files:
            self.file_listbox.insert(tk.END, f)

    def _on_file_select(self, event):
        sel = self.file_listbox.curselection()
        if not sel:
            return
        filename = self.file_listbox.get(sel[0])
        full_path = os.path.join(self._folder_path, filename)
        
        # 智慧唯讀性與可寫性探針檢測，決定預設儲存路徑
        is_writable = True
        test_file = os.path.join(self._folder_path, ".write_test")
        try:
            with open(test_file, "w") as f:
                pass
            os.remove(test_file)
        except Exception:
            is_writable = False
            
        if is_writable:
            self.out_folder_var.set(self._folder_path)
        else:
            self.out_folder_var.set(self.download_path_var.get())
            
        self._load_file(full_path)

    def _load_file(self, path):
        self._stop()
        self.current_file = path
        self.player.open(path)
        self.total_ms = self.player.get_length()
        self.file_label.config(text=f"🎵 {os.path.basename(path)}")
        self.trim_canvas.delete("all")  # 清除舊進度條
        self.time_label.config(text=f"00:00 / {self._fmt(self.total_ms)}")
        # 預設起終點
        self.start_time_str.set("0:00.00")
        self.end_time_str.set(self._fmt_time_str(self.total_ms / 1000))
        self._draw_trim_canvas()
        # 建議輸出檔名
        base = os.path.splitext(os.path.basename(path))[0]
        self.out_entry.delete(0, tk.END)
        self.out_entry.insert(0, f"{base}_trim")
        self.play_btn.config(state="normal")
        self.stop_btn.config(state="normal")
        self.preview_btn.config(state="normal")
        self.preview_toggle_btn.config(state="normal")
        self.trim_btn.config(state="normal")
        self.trim_status.config(text="")

    def _toggle_play(self):
        if not self.current_file:
            return
        mode = self.player.get_mode()
        if mode == "playing":
            self.player.pause()
            self.play_btn.config(text="▶ 播放")
            self.is_paused = True
        elif mode == "paused":
            self.player.resume()
            self.play_btn.config(text="⏸ 暫停")
            self.is_paused = False
            self._preview_mode = False # 切換回普通播放
            self._start_update_loop()
        else:
            self.player.play()
            self.play_btn.config(text="⏸ 暫停")
            self.is_paused = False
            self._preview_mode = False # 切換回普通播放
            self._start_update_loop()

    def _stop(self):
        self.player.stop()
        self.play_btn.config(text="▶ 播放")
        self.preview_btn.config(text="▶ 播放標記區段")
        self.preview_toggle_btn.config(text="▶ 播放 / 暫停")
        self.is_paused = False
        if self.current_file:
            self.time_label.config(text=f"00:00 / {self._fmt(self.total_ms)}")
        if self._update_job:
            self.parent.after_cancel(self._update_job)
            self._update_job = None
        self._draw_trim_canvas()

    def _start_update_loop(self):
        if self._update_job:
            self.parent.after_cancel(self._update_job)
        self._do_update()

    def _do_update(self):
        mode = self.player.get_mode()
        if mode == "playing" and not self._seeking:
            pos = self.player.get_position()
            self.time_label.config(text=f"{self._fmt(pos)} / {self._fmt(self.total_ms)}")
            self._draw_trim_canvas(pos)
            # 試聽模式：到達終點時停止或循環
            if self._preview_mode:
                end_ms = int(self._parse_time(self.end_time_str.get()) * 1000)
                if pos >= end_ms:
                    if self._loop_var.get():
                        # 循環：跳回起點重播
                        s_ms = int(self._parse_time(self.start_time_str.get()) * 1000)
                        self.player.seek(s_ms)
                    else:
                        self._stop()
                        self._preview_mode = False
                        return
        if mode in ("playing", "paused"):
            self._update_job = self.parent.after(200, self._do_update)
            if self._preview_mode:
                if mode == "playing":
                    self.preview_toggle_btn.config(text="⏸ 暫停")
                else:
                    self.preview_toggle_btn.config(text="▶ 播放")
        else:
            self.play_btn.config(text="▶ 播放")
            self.preview_btn.config(text="▶ 播放標記區段")
            self.preview_toggle_btn.config(text="▶ 播放 / 暫停")
            self._preview_mode = False
            self._update_job = None

    def _on_seek_drag(self, event=None):
        """Canvas 拖曳時即時更新時間顯示"""
        pass

    def _on_seek_release(self, event):
        self._seeking = False
        # Canvas 釋放時已在 _canvas_release 處理

    # === Canvas 進度條相關方法 ===
    def _ms_to_x(self, ms):
        """ms 轉換為 Canvas x 座標"""
        w = self.trim_canvas.winfo_width()
        if self.total_ms <= 0 or w <= 0:
            return 0
        return int(ms / self.total_ms * w)

    def _x_to_ms(self, x):
        """Canvas x 座標轉換為 ms"""
        w = self.trim_canvas.winfo_width()
        if self.total_ms <= 0 or w <= 0:
            return 0
        ms = int(x / w * self.total_ms)
        return max(0, min(ms, self.total_ms))

    def _draw_trim_canvas(self, pos_ms=None):
        """重繪 Canvas：背景灰、綠色裁剪區、起終點球、紅色播放指標"""
        c = self.trim_canvas
        w = c.winfo_width()
        h = c.winfo_height()
        if w <= 1:
            return
        c.delete("all")
        # 背景
        c.create_rectangle(0, 0, w, h, fill="#d0d0d0", outline="")
        if self.total_ms > 0:
            s_sec = self._parse_time(self.start_time_str.get())
            e_sec = self._parse_time(self.end_time_str.get())
            xs = self._ms_to_x(int(s_sec * 1000))
            xe = self._ms_to_x(int(e_sec * 1000))
            # 裁剪範圍（綠色）
            c.create_rectangle(xs, 0, xe, h, fill="#81C784", outline="")
            
            # 更新預計長度文字
            dur = abs(e_sec - s_sec)
            self.duration_label.config(text=f"預計長度：{self._fmt_time_str(dur)}")
            
            # 起點線（藍色）
            c.create_rectangle(xs - 2, 0, xs + 2, h, fill="#1976D2", outline="")
            c.create_oval(xs - 6, h // 2 - 6, xs + 6, h // 2 + 6, fill="#1976D2", outline="white", width=1)
            # 終點線（橘色）
            c.create_rectangle(xe - 2, 0, xe + 2, h, fill="#E64A19", outline="")
            c.create_oval(xe - 6, h // 2 - 6, xe + 6, h // 2 + 6, fill="#E64A19", outline="white", width=1)
            # 播放位置指標（紅色）
            if pos_ms is None:
                pos_ms = self.player.get_position() if self.player._is_open else 0
            xp = self._ms_to_x(pos_ms)
            c.create_rectangle(xp - 2, 0, xp + 2, h, fill="#EF5350", outline="")
            c.create_oval(xp - 5, 1, xp + 5, h - 1, fill="#EF5350", outline="white", width=1)

    def _canvas_click(self, event):
        self._seeking = True
        ms = self._x_to_ms(event.x)
        self.player.seek(ms)
        self.time_label.config(text=f"{self._fmt(ms)} / {self._fmt(self.total_ms)}")
        self._draw_trim_canvas(ms)

    def _canvas_drag(self, event):
        if self._seeking:
            ms = self._x_to_ms(event.x)
            self.player.seek(ms)
            self.time_label.config(text=f"{self._fmt(ms)} / {self._fmt(self.total_ms)}")
            self._draw_trim_canvas(ms)

    def _canvas_release(self, event):
        self._seeking = False

    def _seek_relative(self, delta_ms):
        """相對跳轉（delta_ms 可為正負）"""
        if not self.player._is_open:
            return
        pos = self.player.get_position()
        new_pos = max(0, min(pos + delta_ms, self.total_ms))
        self.player.seek(new_pos)
        self.time_label.config(text=f"{self._fmt(new_pos)} / {self._fmt(self.total_ms)}")
        self._draw_trim_canvas(new_pos)

    def _mark_start(self):
        pos_ms = self.player.get_position() if self.player._is_open else 0
        self.start_time_str.set(self._fmt_time_str(pos_ms / 1000))
        self._draw_trim_canvas()

    def _mark_end(self):
        pos_ms = self.player.get_position() if self.player._is_open else self.total_ms
        self.end_time_str.set(self._fmt_time_str(pos_ms / 1000))
        self._draw_trim_canvas()

    def _parse_time(self, t_str):
        """解析 1:23.45 或 83.45 為秒數"""
        try:
            if ":" in t_str:
                parts = t_str.split(":")
                if len(parts) == 2:
                    return float(parts[0]) * 60 + float(parts[1])
            return float(t_str)
        except Exception:
            return 0.0

    def _fmt_time_str(self, sec):
        """將秒數格式化為 M:SS.ss"""
        m = int(sec) // 60
        s = sec % 60
        return f"{m}:{s:05.2f}"

    def _adjust(self, target, delta):
        """微調起點或終點 ±0.1 秒"""
        if target == 'start':
            cur = self._parse_time(self.start_time_str.get())
            val = max(0.0, round(cur + delta, 2))
            self.start_time_str.set(self._fmt_time_str(val))
        else:
            cur = self._parse_time(self.end_time_str.get())
            val = round(cur + delta, 2)
            self.end_time_str.set(self._fmt_time_str(val))
        self._draw_trim_canvas()

    def _preview_section(self):
        """從標記起點開始播放，到達終點時自動停止"""
        if not self.current_file:
            return
            
        s_sec = self._parse_time(self.start_time_str.get())
        e_sec = self._parse_time(self.end_time_str.get())
        
        # 自動對調
        if s_sec > e_sec:
            s_sec, e_sec = e_sec, s_sec
            self.start_time_str.set(self._fmt_time_str(s_sec))
            self.end_time_str.set(self._fmt_time_str(e_sec))

        s_ms = int(s_sec * 1000)
        e_ms = int(e_sec * 1000)
        
        if s_ms >= e_ms:
            return
            
        self._preview_mode = True
        self.player.seek(s_ms)
        self.player.play()
        self.play_btn.config(text="⏸ 暫停")
        self.preview_toggle_btn.config(text="⏸ 暫停")
        self._start_update_loop()

    def _preview_toggle(self):
        """在當前位置切換 播放/暫停，且受限於標記範圍"""
        if not self.current_file:
            return
            
        mode = self.player.get_mode()
        if mode == "playing":
            self.player.pause()
            self.play_btn.config(text="▶ 播放")
            self.preview_toggle_btn.config(text="▶ 播放")
        elif mode == "paused":
            self.player.resume()
            self.play_btn.config(text="⏸ 暫停")
            self.preview_toggle_btn.config(text="⏸ 暫停")
            self._preview_mode = True # 確保是在預覽模式
            self._start_update_loop()
        else:
            # 如果目前是停止狀態，從起點播放
            self._preview_section()

    def _update_displays(self):
        self._draw_trim_canvas()

    def _do_trim(self):
        if not self.current_file:
            messagebox.showwarning("警告", "請先選擇一個 MP3 檔案。")
            return
        s = self._parse_time(self.start_time_str.get())
        e = self._parse_time(self.end_time_str.get())
        
        # 自動對調
        if s > e:
            s, e = e, s
            self.start_time_str.set(self._fmt_time_str(s))
            self.end_time_str.set(self._fmt_time_str(e))
            self._draw_trim_canvas()

        if s >= e:
            messagebox.showerror("錯誤", "無效的裁剪範圍。")
            return
        base_out_name = self.out_entry.get().strip()
        if not base_out_name:
            messagebox.showerror("錯誤", "請輸入輸出檔名。")
            return

        out_name = base_out_name
        out_path = os.path.join(self._folder_path, out_name + ".mp3")
        counter = 1
        while os.path.exists(out_path):
            out_name = f"{base_out_name}({counter})"
            out_path = os.path.join(self._folder_path, out_name + ".mp3")
            counter += 1
        # 暫停播放以釋放檔案鎖定
        if self.player.get_mode() == "playing":
            self.player.pause()
        self.trim_btn.config(state="disabled")
        self.trim_status.config(text="裁剪中，請稍候...", fg="blue")
        threading.Thread(target=self._run_ffmpeg, args=(self.current_file, out_path, s, e), daemon=True).start()

    def _run_ffmpeg(self, in_path, out_path, start_sec, end_sec):
        try:
            cmd = [
                "ffmpeg", "-y",
                "-i", in_path,
                "-ss", str(start_sec),
                "-to", str(end_sec),
                "-c", "copy",
                out_path
            ]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
            if result.returncode == 0:
                self.parent.after(0, lambda: self.trim_status.config(
                    text=f"✅ 已成功儲存：{os.path.basename(out_path)}", fg="green"))
                self.parent.after(0, self._refresh_list)
            else:
                err = result.stderr[-300:] if result.stderr else "未知錯誤"
                self.parent.after(0, lambda: self.trim_status.config(
                    text=f"❌ 裁剪失敗：{err}", fg="red"))
        except FileNotFoundError:
            self.parent.after(0, lambda: messagebox.showerror(
                "ffmpeg 找不到",
                "找不到 ffmpeg 執行檔。\n請確認 ffmpeg 已安裝並加入系統 PATH。"))
            self.parent.after(0, lambda: self.trim_status.config(text="", fg="red"))
        except Exception as ex:
            self.parent.after(0, lambda: self.trim_status.config(
                text=f"❌ 錯誤：{ex}", fg="red"))
        finally:
            self.parent.after(0, lambda: self.trim_btn.config(state="normal"))

    @staticmethod
    def _fmt(ms):
        """將毫秒格式化為 MM:SS"""
        s = int(ms) // 1000
        return f"{s // 60:02d}:{s % 60:02d}"

    @staticmethod
    def _fmt_sec(sec):
        """將秒數格式化為 X分Y.Z秒"""
        sec = float(sec)
        m = int(sec) // 60
        s = sec % 60
        return f"{m}分{s:05.2f}秒"


# ===========================================================================
# VideoTrimmerTab：影片裁剪工具的完整 UI 類別 (支援 MP4/MKV)
# ===========================================================================
class VideoTrimmerTab:
    def __init__(self, parent, download_path_var):
        self.parent = parent
        self.download_path_var = download_path_var
        self.player = MCIPlayer(alias="cyt_video_player")
        self.current_file = None
        self.total_ms = 0
        self.is_playing = False
        self.is_paused = False
        self.start_time_str = tk.StringVar(value="0:00")
        self.end_time_str = tk.StringVar(value="0:00")
        self._update_job = None
        self._seeking = False
        self._preview_mode = False
        self._loop_var = tk.BooleanVar(value=False)
        self._build_ui()

    def _build_ui(self):
        # === 左側：檔案列表區 ===
        left_frame = tk.Frame(self.parent, width=210)
        left_frame.pack(side="left", fill="y", padx=(10, 5), pady=10)
        left_frame.pack_propagate(False)

        tk.Label(left_frame, text="影片檔案列表", font=("Microsoft JhengHei", 11, "bold")).pack(anchor="w")

        folder_frame = tk.Frame(left_frame)
        folder_frame.pack(fill="x", pady=3)
        self.folder_entry = tk.Entry(folder_frame, state="readonly", font=("Microsoft JhengHei", 8))
        self.folder_entry.pack(side="left", fill="x", expand=True)
        tk.Button(folder_frame, text="選擇", command=self._browse_folder, font=("Microsoft JhengHei", 8)).pack(side="left", padx=2)

        self.file_listbox = tk.Listbox(left_frame, font=("Microsoft JhengHei", 9), selectmode="single", activestyle="dotbox")
        self.file_listbox.pack(fill="both", expand=True, pady=5)
        self.file_listbox.bind("<<ListboxSelect>>", self._on_file_select)

        tk.Button(left_frame, text="🔄 重新整理", command=self._refresh_list, font=("Microsoft JhengHei", 9)).pack(fill="x")

        # === 右側：控制區 ===
        right_frame = tk.Frame(self.parent)
        right_frame.pack(side="left", fill="both", expand=True, padx=(5, 10), pady=10)

        # 影片播放容器
        self.video_container = tk.Frame(right_frame, bg="black", height=280)
        self.video_container.pack(fill="x", pady=(0, 5))
        self.video_container.pack_propagate(False)
        self.video_label = tk.Label(self.video_container, text="請從左側選取影片進行預覽", fg="white", bg="black", font=("Microsoft JhengHei", 10))
        self.video_label.place(relx=0.5, rely=0.5, anchor="center")

        # 播放控制按鈕
        ctrl_frame = tk.Frame(right_frame)
        ctrl_frame.pack(anchor="w", pady=3)
        tk.Button(ctrl_frame, text="⏮ -5s", command=lambda: self._seek_relative(-5000), font=("Microsoft JhengHei", 9), width=5).pack(side="left", padx=1)
        tk.Button(ctrl_frame, text="◀ -1s", command=lambda: self._seek_relative(-1000), font=("Microsoft JhengHei", 9), width=5).pack(side="left", padx=1)
        tk.Button(ctrl_frame, text="⏪ -0.1s", command=lambda: self._seek_relative(-100), font=("Microsoft JhengHei", 8), width=6).pack(side="left", padx=1)
        
        self.play_btn = tk.Button(ctrl_frame, text="▶ 播放", command=self._toggle_play,
                                  font=("Microsoft JhengHei", 11, "bold"), bg="#2196F3", fg="white", width=8, state="disabled")
        self.play_btn.pack(side="left", padx=4)
        
        tk.Button(ctrl_frame, text="+0.1s ⏩", command=lambda: self._seek_relative(100), font=("Microsoft JhengHei", 8), width=6).pack(side="left", padx=1)
        tk.Button(ctrl_frame, text="+1s ▶", command=lambda: self._seek_relative(1000), font=("Microsoft JhengHei", 9), width=5).pack(side="left", padx=1)
        tk.Button(ctrl_frame, text="+5s ⏭", command=lambda: self._seek_relative(5000), font=("Microsoft JhengHei", 9), width=5).pack(side="left", padx=1)
        
        self.time_label = tk.Label(ctrl_frame, text="00:00 / 00:00", font=("Microsoft JhengHei", 11), fg="#333")
        self.time_label.pack(side="left", padx=10)

        # Canvas 進度條
        canvas_outer = tk.Frame(right_frame, bg="#333", pady=1)
        canvas_outer.pack(fill="x", pady=(4, 0))
        self.trim_canvas = tk.Canvas(canvas_outer, height=20, bg="#e0e0e0", highlightthickness=0, cursor="hand2")
        self.trim_canvas.pack(fill="both", expand=True)
        self.trim_canvas.bind("<ButtonPress-1>", self._canvas_click)
        self.trim_canvas.bind("<B1-Motion>", self._canvas_drag)
        self.trim_canvas.bind("<Configure>", lambda e: self._draw_trim_canvas())

        # 影片裁剪設定區
        trim_lf = tk.LabelFrame(right_frame, text="✂️ 影片裁剪設定", font=("Microsoft JhengHei", 10, "bold"), padx=10, pady=6)
        trim_lf.pack(fill="x", pady=5)

        # 播放與循環播放
        row0 = tk.Frame(trim_lf)
        row0.pack(fill="x", pady=(2, 5))
        self.preview_btn = tk.Button(row0, text="▶ 播放標記區段", command=self._preview_section,
                                     font=("Microsoft JhengHei", 10, "bold"), bg="#7B1FA2", fg="white", state="disabled", width=14)
        self.preview_btn.pack(side="left", padx=(0, 5))
        
        self.preview_toggle_btn = tk.Button(row0, text="▶ 播放 / 暫停", command=self._preview_toggle,
                                            font=("Microsoft JhengHei", 10, "bold"), bg="#673AB7", fg="white", state="disabled", width=12)
        self.preview_toggle_btn.pack(side="left", padx=5)
        
        tk.Checkbutton(row0, text="🔁 循環播放", variable=self._loop_var,
                       font=("Microsoft JhengHei", 10), fg="#4A148C").pack(side="left")
        self.duration_label = tk.Label(row0, text="預計長度：0秒", font=("Microsoft JhengHei", 10, "bold"), fg="#E91E63")
        self.duration_label.pack(side="right", padx=10)

        # 起點
        row1 = tk.Frame(trim_lf)
        row1.pack(fill="x", pady=2)
        tk.Label(row1, text="起點：", font=("Microsoft JhengHei", 10), width=6).pack(side="left")
        tk.Entry(row1, textvariable=self.start_time_str, width=12, font=("Microsoft JhengHei", 10)).pack(side="left", padx=5)
        tk.Button(row1, text="📍 標記目前位置", command=self._mark_start, font=("Microsoft JhengHei", 9), bg="#1976D2", fg="white").pack(side="left")

        # 終點
        row2 = tk.Frame(trim_lf)
        row2.pack(fill="x", pady=2)
        tk.Label(row2, text="終點：", font=("Microsoft JhengHei", 10), width=6).pack(side="left")
        tk.Entry(row2, textvariable=self.end_time_str, width=12, font=("Microsoft JhengHei", 10)).pack(side="left", padx=5)
        tk.Button(row2, text="📍 標記目前位置", command=self._mark_end, font=("Microsoft JhengHei", 9), bg="#E64A19", fg="white").pack(side="left")
        
        # 儲存資料夾 Row
        self.out_folder_row = tk.Frame(right_frame)
        self.out_folder_row.pack(fill="x", pady=5)
        tk.Label(self.out_folder_row, text="儲存資料夾：", font=("Microsoft JhengHei", 10, "bold"), width=12, anchor="w").pack(side="left")
        self.out_folder_var = tk.StringVar()
        self.out_folder_entry = tk.Entry(self.out_folder_row, textvariable=self.out_folder_var, font=("Microsoft JhengHei", 10), state="readonly")
        self.out_folder_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        tk.Button(self.out_folder_row, text="選擇", command=self._browse_out_folder, bg="#E91E63", fg="white", font=("Microsoft JhengHei", 9)).pack(side="left")

        # 輸出設定
        out_row = tk.Frame(right_frame)
        out_row.pack(fill="x", pady=5)
        tk.Label(out_row, text="儲存檔名：", font=("Microsoft JhengHei", 10)).pack(side="left")
        self.out_entry = tk.Entry(out_row, font=("Microsoft JhengHei", 10))
        self.out_entry.pack(side="left", fill="x", expand=True, padx=5)
        self.ext_label = tk.Label(out_row, text=".mp4", font=("Microsoft JhengHei", 10))
        self.ext_label.pack(side="left")

        # 執行按鈕
        self.trim_btn = tk.Button(right_frame, text="🎬 執行影片裁剪 (無損快速模式)", command=self._do_trim,
                                  font=("Microsoft JhengHei", 12, "bold"), bg="#F44336", fg="white", height=2, state="disabled")
        self.trim_btn.pack(fill="x", pady=5)

        self.extract_btn = tk.Button(right_frame, text="🎵 影音分離 (一鍵提取高品質音軌)", command=self._do_extract_audio,
                                     font=("Microsoft JhengHei", 11, "bold"), bg="#FF9800", fg="white", height=2, state="disabled")
        self.extract_btn.pack(fill="x", pady=2)
        
        self.trim_status = tk.Label(right_frame, text="", font=("Microsoft JhengHei", 9), fg="green")
        self.trim_status.pack()

        self._refresh_list()

    def _browse_folder(self):
        folder = filedialog.askdirectory(initialdir=self.download_path_var.get())
        if folder:
            self._folder_path = folder
            self._update_folder_entry(folder)
            self._refresh_list()

    def _update_folder_entry(self, path):
        self.folder_entry.config(state="normal")
        self.folder_entry.delete(0, tk.END)
        self.folder_entry.insert(0, path)
        self.folder_entry.config(state="readonly")

    def _browse_out_folder(self):
        initial_dir = self.out_folder_var.get() or self.download_path_var.get()
        folder = filedialog.askdirectory(initialdir=initial_dir)
        if folder:
            is_writable = True
            test_file = os.path.join(folder, ".write_test")
            try:
                with open(test_file, "w") as f:
                    pass
                os.remove(test_file)
            except Exception:
                is_writable = False
            
            if not is_writable:
                messagebox.showerror("權限錯誤", "所選資料夾為唯讀或無寫入權限，請選擇其他儲存位置。")
            else:
                self.out_folder_var.set(folder)

    def _browse_out_folder(self):
        initial_dir = self.out_folder_var.get() or self.download_path_var.get()
        folder = filedialog.askdirectory(initialdir=initial_dir)
        if folder:
            is_writable = True
            test_file = os.path.join(folder, ".write_test")
            try:
                with open(test_file, "w") as f:
                    pass
                os.remove(test_file)
            except Exception:
                is_writable = False
            
            if not is_writable:
                messagebox.showerror("權限錯誤", "所選資料夾為唯讀或無寫入權限，請選擇其他儲存位置。")
            else:
                self.out_folder_var.set(folder)

    def _refresh_list(self):
        folder = (getattr(self, '_folder_path', None) or self.download_path_var.get()) or ""
        self._folder_path = folder
        self._update_folder_entry(folder)
        self.file_listbox.delete(0, tk.END)
        if not folder or not os.path.exists(folder): return
        files = sorted([f for f in os.listdir(folder) if f.lower().endswith(('.mp4', '.mkv'))])
        for f in files: self.file_listbox.insert(tk.END, f)

    def _on_file_select(self, event):
        sel = self.file_listbox.curselection()
        if not sel: return
        filename = self.file_listbox.get(sel[0])
        full_path = os.path.join(self._folder_path, filename)
        
        # 智慧唯讀性與可寫性探針檢測，決定預設儲存路徑
        is_writable = True
        test_file = os.path.join(self._folder_path, ".write_test")
        try:
            with open(test_file, "w") as f:
                pass
            os.remove(test_file)
        except Exception:
            is_writable = False
            
        if is_writable:
            self.out_folder_var.set(self._folder_path)
        else:
            self.out_folder_var.set(self.download_path_var.get())
            
        self._load_video(full_path)

    def _load_video(self, path):
        self._stop()
        self.current_file = path
        # 使用 MCI 播放器
        short_path = self.player._get_short_path(path)
        p = f'"{short_path}"' if short_path else f'"{path}"'
        
        # 開啟影片並嵌入視窗
        hwnd = self.video_container.winfo_id()
        self.player._send(f'open {p} type mpegvideo alias {self.player._alias} style child parent {hwnd}')
        self.player._send(f'set {self.player._alias} time format milliseconds')
        self.player._is_open = True
        
        self.total_ms = self.player.get_length()
        self.time_label.config(text=f"00:00 / {self._fmt(self.total_ms)}")
        self.start_time_str.set("0:00.00")
        self.end_time_str.set(self._fmt_time_str(self.total_ms / 1000))
        
        # 調整影片顯示區域
        w, h = self.video_container.winfo_width(), self.video_container.winfo_height()
        self.player._send(f'put {self.player._alias} window at 0 0 {w} {h}')
        
        base, ext = os.path.splitext(os.path.basename(path))
        self.out_entry.delete(0, tk.END)
        self.out_entry.insert(0, f"{base}_clip")
        self.ext_label.config(text=ext)
        self.play_btn.config(state="normal")
        self.preview_btn.config(state="normal")
        self.preview_toggle_btn.config(state="normal")
        self.trim_btn.config(state="normal")
        self.extract_btn.config(state="normal")
        self.video_label.place_forget()
        self._draw_trim_canvas()

    def _toggle_play(self):
        if not self.current_file: return
        mode = self.player.get_mode()
        if mode == "playing":
            self.player.pause()
            self.play_btn.config(text="▶ 播放")
        else:
            self.player.play()
            self.play_btn.config(text="⏸ 暫停")
            self._preview_mode = False # 一般播放模式
            self._start_update_loop()

    def _stop(self):
        self.player.close()
        self.play_btn.config(text="▶ 播放")
        self.preview_btn.config(text="▶ 播放標記區段")
        self.preview_toggle_btn.config(text="▶ 播放 / 暫停")
        if self._update_job: self.parent.after_cancel(self._update_job)
        self.video_label.place(relx=0.5, rely=0.5, anchor="center")

    def _start_update_loop(self):
        if self._update_job: self.parent.after_cancel(self._update_job)
        self._do_update()

    def _do_update(self):
        if self.player._is_open:
            mode = self.player.get_mode()
            pos = self.player.get_position()
            self.time_label.config(text=f"{self._fmt(pos)} / {self._fmt(self.total_ms)}")
            self._draw_trim_canvas(pos)
            
            # 試聽模式邏輯
            if self._preview_mode:
                if mode == "playing":
                    self.preview_toggle_btn.config(text="⏸ 暫停")
                else:
                    self.preview_toggle_btn.config(text="▶ 播放")
                
                e_ms = int(self._parse_time(self.end_time_str.get()) * 1000)
                if pos >= e_ms:
                    if self._loop_var.get():
                        s_ms = int(self._parse_time(self.start_time_str.get()) * 1000)
                        self.player.seek(s_ms)
                        # 強制重新整理畫面
                        self.player._send(f'update {self.player._alias}')
                    else:
                        self.player.pause()
                        self.play_btn.config(text="▶ 播放")
                        self.preview_btn.config(text="▶ 播放標記區段")
                        self.preview_toggle_btn.config(text="▶ 播放 / 暫停")
                        self._preview_mode = False
                        return

            if self.player.get_mode() == "playing":
                self._update_job = self.parent.after(200, self._do_update)
            else:
                self.play_btn.config(text="▶ 播放")
                self.preview_btn.config(text="▶ 播放標記區段")
                self.preview_toggle_btn.config(text="▶ 播放 / 暫停")
                self._preview_mode = False
                self._update_job = None

    def _preview_section(self):
        if not self.player._is_open: return
        
        # 全新試聽：從起點開始
        s = self._parse_time(self.start_time_str.get())
        self.player.seek(int(s * 1000))
        self._preview_mode = True
        self.player.play()
        self.play_btn.config(text="⏸ 暫停")
        self.preview_toggle_btn.config(text="⏸ 暫停")
        self._start_update_loop()
        # 強制刷新畫面
        self.player._send(f'update {self.player._alias}')

    def _preview_toggle(self):
        """在當前位置切換 播放/暫停，且受限於標記範圍"""
        if not self.player._is_open: return
        
        mode = self.player.get_mode()
        if mode == "playing":
            self.player.pause()
            self.play_btn.config(text="▶ 播放")
            self.preview_toggle_btn.config(text="▶ 播放")
        elif mode == "paused":
            self.player.play()
            self.play_btn.config(text="⏸ 暫停")
            self.preview_toggle_btn.config(text="⏸ 暫停")
            self._preview_mode = True # 確保是在預覽模式
            self._start_update_loop()
        else:
            # 如果目前是停止狀態，從起點播放
            self._preview_section()

    def _seek_relative(self, delta_ms):
        if not self.player._is_open: return
        pos = self.player.get_position()
        new_pos = max(0, min(pos + delta_ms, self.total_ms))
        self.player.seek(new_pos)
        
        # 強制重新整理畫面 (MCI 在暫停時 Seek 可能不會更新視窗，需要 update 或 put)
        if self.player.get_mode() != "playing":
            w, h = self.video_container.winfo_width(), self.video_container.winfo_height()
            self.player._send(f'put {self.player._alias} window at 0 0 {w} {h}')
            self.player._send(f'update {self.player._alias}')
            
        # 即時同步 UI
        self.time_label.config(text=f"{self._fmt(new_pos)} / {self._fmt(self.total_ms)}")
        self._draw_trim_canvas(new_pos)

    def _mark_start(self):
        pos = self.player.get_position() if self.player._is_open else 0
        self.start_time_str.set(self._fmt_time_str(pos / 1000))
        self._draw_trim_canvas()

    def _mark_end(self):
        pos = self.player.get_position() if self.player._is_open else self.total_ms
        self.end_time_str.set(self._fmt_time_str(pos / 1000))
        self._draw_trim_canvas()

    def _draw_trim_canvas(self, pos_ms=None):
        c = self.trim_canvas
        w, h = c.winfo_width(), c.winfo_height()
        if w <= 1: return
        c.delete("all")
        c.create_rectangle(0, 0, w, h, fill="#d0d0d0", outline="")
        if self.total_ms > 0:
            s = self._parse_time(self.start_time_str.get())
            e = self._parse_time(self.end_time_str.get())
            
            # 更新預計長度文字
            dur = abs(e - s)
            self.duration_label.config(text=f"預計長度：{self._fmt_time_str(dur)}")
            
            xs, xe = int(s*1000/self.total_ms*w), int(e*1000/self.total_ms*w)
            c.create_rectangle(xs, 0, xe, h, fill="#81C784", outline="")
            if pos_ms is None: pos_ms = self.player.get_position()
            xp = int(pos_ms/self.total_ms*w)
            c.create_rectangle(xp-1, 0, xp+1, h, fill="red", outline="")

    def _canvas_click(self, event):
        if self.total_ms <= 0: return
        self._seeking = True
        w = self.trim_canvas.winfo_width()
        ms = int(event.x / w * self.total_ms)
        self.player.seek(ms)
        # 暫停時強制更新畫面
        if self.player.get_mode() != "playing":
            w_v, h_v = self.video_container.winfo_width(), self.video_container.winfo_height()
            self.player._send(f'put {self.player._alias} window at 0 0 {w_v} {h_v}')
            self.player._send(f'update {self.player._alias}')
            
        self.time_label.config(text=f"{self._fmt(ms)} / {self._fmt(self.total_ms)}")
        self._draw_trim_canvas(ms)

    def _canvas_drag(self, event):
        self._canvas_click(event)

    def _do_trim(self):
        if not self.current_file: return
        s, e = self._parse_time(self.start_time_str.get()), self._parse_time(self.end_time_str.get())
        if s >= e: 
            messagebox.showerror("錯誤", "起點必須小於終點。")
            return
        out_name = self.out_entry.get().strip() + self.ext_label.cget("text")
        target_folder = self.out_folder_var.get() or self.download_path_var.get()
        os.makedirs(target_folder, exist_ok=True)
        out_path = os.path.join(target_folder, out_name)
        
        self.trim_btn.config(state="disabled")
        self.trim_status.config(text="影片裁剪中（無損模式速度極快）...", fg="blue")
        threading.Thread(target=self._run_ffmpeg, args=(self.current_file, out_path, s, e), daemon=True).start()

    def _run_ffmpeg(self, in_path, out_path, start, end):
        try:
            # 使用 -ss 在 -i 前面可實現快速跳轉，搭配 -c copy 實現無損剪輯
            cmd = ["ffmpeg", "-y", "-ss", str(start), "-i", in_path, "-to", str(end-start), "-c", "copy", out_path]
            res = subprocess.run(cmd, capture_output=True, text=True)
            if res.returncode == 0:
                self.parent.after(0, lambda: self.trim_status.config(text=f"✅ 裁剪成功：{os.path.basename(out_path)}", fg="green"))
                self.parent.after(0, self._refresh_list)
            else:
                self.parent.after(0, lambda: self.trim_status.config(text="❌ 裁剪失敗", fg="red"))
        except Exception as ex:
            self.parent.after(0, lambda: self.trim_status.config(text=f"❌ 錯誤：{ex}", fg="red"))
        finally:
            self.parent.after(0, lambda: self.trim_btn.config(state="normal"))

    def _fmt(self, ms):
        s = int(ms) // 1000
        return f"{s // 60:02d}:{s % 60:02d}"

    def _fmt_time_str(self, sec):
        return f"{int(sec)//60}:{sec%60:05.2f}"

    def _parse_time(self, t_str):
        try:
            if ":" in t_str:
                p = t_str.split(":")
                return float(p[0])*60 + float(p[1])
            return float(t_str)
        except: return 0.0

    def _do_extract_audio(self):
        if not self.current_file: return
        # 彈出小視窗選擇格式
        dialog = tk.Toplevel(self.parent)
        dialog.title("選擇音訊提取格式")
        dialog.geometry("320x150")
        dialog.resizable(False, False)
        dialog.transient(self.parent)
        dialog.grab_set()
        
        # 居中顯示
        dialog.update_idletasks()
        w = dialog.winfo_width()
        h = dialog.winfo_height()
        x = self.parent.winfo_rootx() + (self.parent.winfo_width() - w) // 2
        y = self.parent.winfo_rooty() + (self.parent.winfo_height() - h) // 2
        dialog.geometry(f"+{x}+{y}")
        
        tk.Label(dialog, text="請選擇要提取的音訊格式：", font=("Microsoft JhengHei", 10)).pack(pady=15)
        
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(fill="x", padx=10)
        
        def select_format(fmt):
            dialog.destroy()
            self._start_audio_extraction(fmt)
            
        tk.Button(btn_frame, text="高品質 MP3 (320k)", bg="#4CAF50", fg="white", font=("Microsoft JhengHei", 9, "bold"),
                   command=lambda: select_format("mp3")).pack(side="left", expand=True, fill="x", padx=5)
        tk.Button(btn_frame, text="無損 WAV 音軌", bg="#2196F3", fg="white", font=("Microsoft JhengHei", 9, "bold"),
                   command=lambda: select_format("wav")).pack(side="left", expand=True, fill="x", padx=5)

    def _start_audio_extraction(self, fmt):
        base, _ = os.path.splitext(os.path.basename(self.current_file))
        target_folder = self.out_folder_var.get() or self.download_path_var.get()
        os.makedirs(target_folder, exist_ok=True)
        counter = 1
        out_filename = f"{base}.{fmt}"
        out_path = os.path.join(target_folder, out_filename)
        while os.path.exists(out_path):
            out_filename = f"{base}_{counter}.{fmt}"
            out_path = os.path.join(folder, out_filename)
            counter += 1
            
        self.extract_btn.config(state="disabled")
        self.trim_status.config(text=f"正在提取高品質 {fmt.upper()} 音軌...", fg="blue")
        threading.Thread(target=self._run_ffmpeg_extract, args=(self.current_file, out_path, fmt), daemon=True).start()

    def _run_ffmpeg_extract(self, in_path, out_path, fmt):
        try:
            if fmt == "mp3":
                cmd = ["ffmpeg", "-y", "-i", in_path, "-vn", "-acodec", "libmp3lame", "-ab", "320k", out_path]
            else: # wav
                cmd = ["ffmpeg", "-y", "-i", in_path, "-vn", out_path]
                
            res = subprocess.run(cmd, capture_output=True, text=True)
            if res.returncode == 0:
                self.parent.after(0, lambda: self.trim_status.config(text=f"✅ 影音分離成功：{os.path.basename(out_path)}", fg="green"))
                self.parent.after(0, self._refresh_list)
            else:
                self.parent.after(0, lambda: self.trim_status.config(text="❌ 提取失敗", fg="red"))
        except Exception as ex:
            self.parent.after(0, lambda: self.trim_status.config(text=f"❌ 錯誤：{ex}", fg="red"))
        finally:
            self.parent.after(0, lambda: self.extract_btn.config(state="normal"))


# ===========================================================================
# VideoConverterTab：影音轉檔工具的完整 UI 類別
# ===========================================================================
class VideoConverterTab:
    def __init__(self, parent, download_path_var):
        self.parent = parent
        self.download_path_var = download_path_var
        self.current_file = None
        self._folder_path = ""
        
        self.target_format = tk.StringVar(value="MP4")
        self.scale_choice = tk.StringVar(value="保持原解析度")
        self.speed_choice = tk.StringVar(value="預設速度 (Medium)")
        
        self._build_ui()

    def _build_ui(self):
        # === 左側：檔案列表區 ===
        left_frame = tk.Frame(self.parent, width=220)
        left_frame.pack(side="left", fill="y", padx=(10, 5), pady=10)
        left_frame.pack_propagate(False)

        tk.Label(left_frame, text="待轉檔影音列表", font=("Microsoft JhengHei", 11, "bold")).pack(anchor="w")

        folder_frame = tk.Frame(left_frame)
        folder_frame.pack(fill="x", pady=3)
        self.folder_entry = tk.Entry(folder_frame, state="readonly", font=("Microsoft JhengHei", 8))
        self.folder_entry.pack(side="left", fill="x", expand=True)
        tk.Button(folder_frame, text="選擇", command=self._browse_folder, font=("Microsoft JhengHei", 8)).pack(side="left", padx=2)

        self.file_listbox = tk.Listbox(left_frame, font=("Microsoft JhengHei", 9), selectmode="single", activestyle="dotbox")
        self.file_listbox.pack(fill="both", expand=True, pady=5)
        self.file_listbox.bind("<<ListboxSelect>>", self._on_file_select)

        tk.Button(left_frame, text="🔄 重新整理", command=self._refresh_list, font=("Microsoft JhengHei", 9)).pack(fill="x")

        # === 右側：控制區 ===
        right_frame = tk.Frame(self.parent)
        right_frame.pack(side="left", fill="both", expand=True, padx=(5, 10), pady=10)

        # 選取影片資訊
        self.info_lf = tk.LabelFrame(right_frame, text="🎬 已選取影音檔案", font=("Microsoft JhengHei", 10, "bold"), padx=15, pady=10)
        self.info_lf.pack(fill="x", pady=(0, 10))
        self.info_label = tk.Label(self.info_lf, text="請先從左側選擇要轉換的影片或音訊檔案", font=("Microsoft JhengHei", 10), fg="gray", justify="left", anchor="w", wraplength=500)
        self.info_label.pack(fill="x")
        
        # DVD VOB 合併勾選框 (預設隱藏，檢測到連續 VOB 時才 pack)
        self.merge_vobs_var = tk.BooleanVar(value=False)
        self.merge_vobs_chk = tk.Checkbutton(self.info_lf, text="偵測到連續的 DVD VOB 檔案，是否一鍵無縫合併轉檔？", 
                                             variable=self.merge_vobs_var, font=("Microsoft JhengHei", 9, "bold"), fg="#FF5722",
                                             anchor="w", justify="left", wraplength=480)

        # 轉檔參數設定
        params_lf = tk.LabelFrame(right_frame, text="⚙️ 轉檔參數與畫質壓縮設定", font=("Microsoft JhengHei", 10, "bold"), padx=15, pady=15)
        params_lf.pack(fill="x", pady=5)

        # 目標格式
        row1 = tk.Frame(params_lf)
        row1.pack(fill="x", pady=5)
        tk.Label(row1, text="目標輸出格式：", font=("Microsoft JhengHei", 10, "bold"), width=14, anchor="w").pack(side="left")
        self.format_combo = ttk.Combobox(row1, values=["MP4 (相容性最高)", "MKV (多軌道支援)", "MP3 (高品質音軌)", "WAV (無損音軌)"], state="readonly", font=("Microsoft JhengHei", 9))
        self.format_combo.set("MP4 (相容性最高)")
        self.format_combo.pack(side="left", fill="x", expand=True)
        self.format_combo.bind("<<ComboboxSelected>>", self._on_format_combo_change)

        # 儲存位置設定 Row
        self.out_folder_row = tk.Frame(params_lf)
        self.out_folder_row.pack(fill="x", pady=5)
        tk.Label(self.out_folder_row, text="儲存資料夾：", font=("Microsoft JhengHei", 10, "bold"), width=14, anchor="w").pack(side="left")
        self.out_folder_var = tk.StringVar()
        self.out_folder_entry = tk.Entry(self.out_folder_row, textvariable=self.out_folder_var, font=("Microsoft JhengHei", 10), state="readonly")
        self.out_folder_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        tk.Button(self.out_folder_row, text="選擇", command=self._browse_out_folder, bg="#E91E63", fg="white", font=("Microsoft JhengHei", 9)).pack(side="left")

        # 解析度降低
        self.row2 = tk.Frame(params_lf)
        self.row2.pack(fill="x", pady=5)
        self.scale_label = tk.Label(self.row2, text="畫質壓縮/解析度：", font=("Microsoft JhengHei", 10, "bold"), width=14, anchor="w")
        self.scale_label.pack(side="left")
        self.scale_combo = ttk.Combobox(self.row2, values=["保持原解析度", "1080p (1920x1080)", "720p (1280x720) [推薦，體積減60%]", "480p (854x480) [快速壓縮]"], state="readonly", font=("Microsoft JhengHei", 9))
        self.scale_combo.set("保持原解析度")
        self.scale_combo.pack(side="left", fill="x", expand=True)

        # 轉檔速度 (CPU Preset)
        self.row3 = tk.Frame(params_lf)
        self.row3.pack(fill="x", pady=5)
        self.speed_label = tk.Label(self.row3, text="轉檔編碼速度：", font=("Microsoft JhengHei", 10, "bold"), width=14, anchor="w")
        self.speed_label.pack(side="left")
        self.speed_combo = ttk.Combobox(self.row3, values=["極速模式 (Veryfast) [速度極快，體積稍大]", "預設速度 (Medium)", "高品質模式 (Slow) [壓縮率最高，較慢]"], state="readonly", font=("Microsoft JhengHei", 9))
        self.speed_combo.set("預設速度 (Medium)")
        self.speed_combo.pack(side="left", fill="x", expand=True)

        # 外掛字幕設定
        self.sub_row = tk.Frame(params_lf)
        self.sub_row.pack(fill="x", pady=5)
        
        self.merge_sub_var = tk.BooleanVar(value=False)
        self.sub_chk = tk.Checkbutton(self.sub_row, text="🎬 合併外掛字幕 (.srt)：", variable=self.merge_sub_var, font=("Microsoft JhengHei", 10, "bold"), command=self._on_sub_chk_change)
        self.sub_chk.pack(side="left")
        
        self.sub_path_var = tk.StringVar()
        self.sub_entry = tk.Entry(self.sub_row, textvariable=self.sub_path_var, state="readonly", font=("Microsoft JhengHei", 9), width=25)
        self.sub_entry.pack(side="left", fill="x", expand=True, padx=5)
        
        self.sub_btn = tk.Button(self.sub_row, text="選擇 SRT", command=self._browse_srt, font=("Microsoft JhengHei", 8), state="disabled")
        self.sub_btn.pack(side="left", padx=2)

        # 轉檔說明提示
        hint_lbl = tk.Label(params_lf, text="💡 提示：降低解析度（如將 1080p 轉為 720p）能大幅縮小影片檔案體積，非常適合在手機儲存與分享傳送。", font=("Microsoft JhengHei", 9), fg="#666", justify="left", wraplength=500)
        hint_lbl.pack(fill="x", pady=(10, 0))

        # 執行轉檔大按鈕
        self.convert_btn = tk.Button(right_frame, text="🚀 開始影音轉檔與畫質壓縮", command=self._do_convert, font=("Microsoft JhengHei", 12, "bold"), bg="#4CAF50", fg="white", height=2, state="disabled")
        self.convert_btn.pack(fill="x", pady=15)

        # 狀態與進度控制區
        self.progress_frame = tk.Frame(right_frame)
        self.progress_frame.pack(fill="x", pady=5)

        self.status_label = tk.Label(self.progress_frame, text="", font=("Microsoft JhengHei", 10, "bold"), fg="blue")
        self.status_label.pack(pady=2)

        self.progress_bar = ttk.Progressbar(self.progress_frame, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill="x", pady=5)
        self.progress_bar.pack_forget()

        self.progress_label = tk.Label(self.progress_frame, text="", font=("Microsoft JhengHei", 9), fg="#555")
        self.progress_label.pack(pady=2)
        self.progress_label.pack_forget()

        self._refresh_list()

    def _browse_folder(self):
        folder = filedialog.askdirectory(initialdir=self.download_path_var.get())
        if folder:
            self._folder_path = folder
            self._update_folder_entry(folder)
            self._refresh_list()

    def _update_folder_entry(self, path):
        self.folder_entry.config(state="normal")
        self.folder_entry.delete(0, tk.END)
        self.folder_entry.insert(0, path)
        self.folder_entry.config(state="readonly")

    def _browse_out_folder(self):
        initial_dir = self.out_folder_var.get() or self.download_path_var.get()
        folder = filedialog.askdirectory(initialdir=initial_dir)
        if folder:
            is_writable = True
            test_file = os.path.join(folder, ".write_test")
            try:
                with open(test_file, "w") as f:
                    pass
                os.remove(test_file)
            except Exception:
                is_writable = False
            
            if not is_writable:
                messagebox.showerror("權限錯誤", "所選資料夾為唯讀或無寫入權限，請選擇其他儲存位置。")
            else:
                self.out_folder_var.set(folder)

    def _refresh_list(self):
        folder = (getattr(self, '_folder_path', None) or self.download_path_var.get()) or ""
        self._folder_path = folder
        self._update_folder_entry(folder)
        self.file_listbox.delete(0, tk.END)
        if not folder or not os.path.exists(folder): return
        
        # 支持主流的影片與音訊格式
        valid_exts = ('.mp4', '.mkv', '.avi', '.flv', '.mov', '.webm', '.ts', '.mp3', '.wav', '.m4a', '.vob', '.dat')
        files = sorted([f for f in os.listdir(folder) if f.lower().endswith(valid_exts)])
        for f in files: self.file_listbox.insert(tk.END, f)

    def _on_file_select(self, event):
        sel = self.file_listbox.curselection()
        if not sel: return
        filename = self.file_listbox.get(sel[0])
        full_path = os.path.join(self._folder_path, filename)
        self.current_file = full_path
        
        # 取得檔案大小
        size_bytes = os.path.getsize(full_path)
        size_mb = size_bytes / (1024 * 1024)
        
        self.info_label.config(text=f"📂 檔名：{filename}\n⚖️ 大小：{size_mb:.2f} MB\n📍 路徑：{full_path}", fg="#333", font=("Microsoft JhengHei", 9, "bold"))
        self.convert_btn.config(state="normal")

        # 智慧唯讀性與可寫性探針校驗，決定介面上預設展示的儲存路徑
        is_writable = True
        test_file = os.path.join(self._folder_path, ".write_test")
        try:
            with open(test_file, "w") as f:
                pass
            os.remove(test_file)
        except Exception:
            is_writable = False
            
        if is_writable:
            self.out_folder_var.set(self._folder_path)
        else:
            self.out_folder_var.set(self.download_path_var.get())

        # 自動搜尋同目錄下同名且副檔名為 .srt 的檔案
        base, _ = os.path.splitext(filename)
        srt_name = f"{base}.srt"
        srt_path = os.path.join(self._folder_path, srt_name)
        if os.path.exists(srt_path):
            self.sub_path_var.set(srt_path)
            self.merge_sub_var.set(True)
            self._on_sub_chk_change()
        else:
            self.sub_path_var.set("")
            self.merge_sub_var.set(False)
            self._on_sub_chk_change()

        # 智慧偵測 DVD 連續 VOB 檔案
        self.detected_vobs = []
        self.merge_vobs_var.set(False)
        self.merge_vobs_chk.pack_forget()
        
        filename_lower = filename.lower()
        if filename_lower.endswith(".vob"):
            import re
            # DVD 的正片檔案命名通常為 VTS_XX_Y.VOB (其中 Y >= 1)
            match = re.match(r'^(vts_\d+_)(\d+)\.vob$', filename_lower)
            if match:
                prefix = match.group(1)
                current_idx = int(match.group(2))
                if current_idx > 0:
                    vob_files = []
                    i = 1
                    while True:
                        target_name = f"{prefix}{i}.vob"
                        target_path = None
                        try:
                            for f in os.listdir(self._folder_path):
                                if f.lower() == target_name:
                                    target_path = os.path.join(self._folder_path, f)
                                    break
                        except Exception:
                            pass
                        if target_path:
                            vob_files.append(target_path)
                            i += 1
                        else:
                            break
                            
                    if len(vob_files) > 1:
                        self.detected_vobs = vob_files
                        first_name = os.path.basename(vob_files[0])
                        last_name = os.path.basename(vob_files[-1])
                        self.merge_vobs_chk.config(text=f"✨ 偵測到連續 DVD 影片檔案 ({first_name} ~ {last_name})，是否勾選此處進行一鍵無縫合併轉檔？")
                        self.merge_vobs_chk.pack(fill="x", anchor="w", padx=5, pady=5)

    def _on_format_combo_change(self, event):
        fmt = self.format_combo.get()
        if "MP3" in fmt or "WAV" in fmt:
            # 音訊格式隱藏解析度、編碼速度與外掛字幕設定
            self.row2.pack_forget()
            self.row3.pack_forget()
            self.sub_row.pack_forget()
        else:
            # 影片格式顯示解析度、編碼速度與外掛字幕設定
            self.row2.pack(fill="x", pady=5)
            self.row3.pack(fill="x", pady=5)
            self.sub_row.pack(fill="x", pady=5)

    def _on_sub_chk_change(self):
        state = "normal" if self.merge_sub_var.get() else "disabled"
        self.sub_btn.config(state=state)
        if self.merge_sub_var.get() and not self.sub_path_var.get():
            self._browse_srt()

    def _browse_srt(self):
        initial_dir = self._folder_path or self.download_path_var.get()
        f = filedialog.askopenfilename(
            initialdir=initial_dir,
            filetypes=[("字幕檔案", "*.srt"), ("所有檔案", "*.*")]
        )
        if f:
            self.sub_path_var.set(f)
            self.merge_sub_var.set(True)
            self._on_sub_chk_change()

    def _browse_out_folder(self):
        initial_dir = self.out_folder_var.get() or self.download_path_var.get()
        folder = filedialog.askdirectory(initialdir=initial_dir)
        if folder:
            # 智慧寫入測試，確保使用者選擇的目錄不是唯讀的！防呆防到底！
            is_writable = True
            test_file = os.path.join(folder, ".write_test")
            try:
                with open(test_file, "w") as f:
                    pass
                os.remove(test_file)
            except Exception:
                is_writable = False
            
            if not is_writable:
                messagebox.showerror("權限錯誤", "所選資料夾為唯讀或無寫入權限，請選擇其他儲存位置。")
            else:
                self.out_folder_var.set(folder)

    def _do_convert(self):
        if not self.current_file: return
        
        in_path = self.current_file
        target_folder = self.out_folder_var.get()
        if not target_folder:
            target_folder = self.download_path_var.get()
            
        # 確保儲存路徑實體存在
        os.makedirs(target_folder, exist_ok=True)
            
        # 判斷是否啟用 DVD 連續 VOB 一鍵合併轉檔
        is_merge = getattr(self, 'detected_vobs', None) and self.merge_vobs_var.get()
        
        if is_merge:
            first_vob = self.detected_vobs[0]
            last_vob = self.detected_vobs[-1]
            base = f"{os.path.splitext(os.path.basename(first_vob))[0]}_to_{os.path.splitext(os.path.basename(last_vob))[0]}"
        else:
            base, _ = os.path.splitext(os.path.basename(in_path))
        
        # 決定目標副檔名與格式
        fmt_sel = self.format_combo.get()
        if "MP4" in fmt_sel:
            out_ext = ".mp4"
        elif "MKV" in fmt_sel:
            out_ext = ".mkv"
        elif "MP3" in fmt_sel:
            out_ext = ".mp3"
        else:
            out_ext = ".wav"
            
        out_name = f"{base}_converted{out_ext}"
        out_path = os.path.join(target_folder, out_name)
        
        # 防撞名機制
        base_out = f"{base}_converted"
        counter = 1
        while os.path.exists(out_path):
            out_path = os.path.join(target_folder, f"{base_out}({counter}){out_ext}")
            counter += 1
            
        self.convert_btn.config(state="disabled")
        self.status_label.config(text="🎬 影音轉檔壓縮中，這可能需要幾分鐘，請稍候...", fg="blue")
            
        self.progress_bar['value'] = 0
        self.progress_label.config(text="正在分析影音結構，請稍候...")
        
        # 背景線程轉檔
        vobs = self.detected_vobs if is_merge else None
        threading.Thread(target=self._run_ffmpeg_convert, args=(in_path, out_path, fmt_sel, vobs), daemon=True).start()

    def _run_ffmpeg_convert(self, in_path, out_path, fmt_sel, vob_list=None):
        try:
            # 取得影片總時長
            if vob_list:
                total_seconds = sum(self._get_video_duration(p) for p in vob_list)
            else:
                total_seconds = self._get_video_duration(in_path)
                
            if total_seconds > 0:
                self.parent.after(0, lambda: self.progress_bar.pack(fill="x", pady=5))
                self.parent.after(0, lambda: self.progress_label.pack(pady=2))
            
            # 針對 DVD VOB 轉成 MKV 的智慧多音軌與多字幕無損保留
            is_dvd_mkv = "MKV" in fmt_sel and in_path.lower().endswith(".vob")
            
            # 基本命令
            if vob_list:
                vob_names = [os.path.basename(p) for p in vob_list]
                concat_str = "concat:" + "|".join(vob_names)
                cmd = ["ffmpeg", "-y", "-i", concat_str]
            else:
                cmd = ["ffmpeg", "-y", "-i", in_path]
            
            if "MP3" in fmt_sel:
                # 轉為高品質音訊
                cmd += ["-vn", "-acodec", "libmp3lame", "-ab", "320k", out_path]
            elif "WAV" in fmt_sel:
                # 轉為無損音訊
                cmd += ["-vn", out_path]
            else:
                # 影片格式轉換及畫質壓縮
                scale = self.scale_combo.get()
                vf_args = []
                
                # 解析度降低處理
                if "1080p" in scale:
                    vf_args.append("scale=-2:1080")
                elif "720p" in scale:
                    vf_args.append("scale=-2:720")
                elif "480p" in scale:
                    vf_args.append("scale=-2:480")
                
                # 設定轉檔速度 (Preset)
                speed = self.speed_combo.get()
                preset_val = "medium"
                if "極速" in speed:
                    preset_val = "veryfast"
                elif "高品質" in speed:
                    preset_val = "slow"
                    
                if is_dvd_mkv:
                    # 智慧 Remux 保留多軌道：映射所有軌道 (-map 0)，多音軌與多字幕直接無損複製 (-c:a copy, -c:s copy)
                    cmd += ["-map", "0"]
                    if vf_args:
                        cmd += ["-vf", ",".join(vf_args)]
                    cmd += ["-c:v", "libx264", "-preset", preset_val, "-crf", "22", "-c:a", "copy", "-c:s", "copy", out_path]
                else:
                    # 一般轉檔單軌處理
                    if self.merge_sub_var.get() and self.sub_path_var.get():
                        sub_path = self.sub_path_var.get()
                        if os.path.exists(sub_path):
                            escaped_sub = self._escape_ffmpeg_path(sub_path)
                            vf_args.append(f"subtitles='{escaped_sub}'")
                    
                    if vf_args:
                        cmd += ["-vf", ",".join(vf_args)]
                        
                    cmd += ["-c:v", "libx264", "-preset", preset_val, "-crf", "22", "-c:a", "aac", "-b:a", "192k", out_path]
                
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="ignore",
                bufsize=1,
                cwd=self._folder_path,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
            )

            start_time = time.time()
            time_pattern = re.compile(r"time=(\d+):(\d+):(\d+(?:\.\d+)?)")
            speed_pattern = re.compile(r"speed=\s*(\d+\.?\d*)x")
            last_lines = []

            while True:
                line = process.stdout.readline()
                if not line:
                    break
                line = line.strip()
                if line:
                    last_lines.append(line)
                    if len(last_lines) > 10:
                        last_lines.pop(0)
                
                # 解析時間進度
                time_match = time_pattern.search(line)
                if time_match and total_seconds > 0:
                    h, m, s = time_match.group(1), time_match.group(2), time_match.group(3)
                    curr_seconds = int(h) * 3600 + int(m) * 60 + float(s)
                    
                    # 計算百分比
                    percent = min(100.0, max(0.0, (curr_seconds / total_seconds) * 100))
                    
                    # 解析速度
                    speed_match = speed_pattern.search(line)
                    speed_str = f"{speed_match.group(1)}x" if speed_match else "N/A"
                    
                    # 計算剩餘時間 (ETA)
                    elapsed = time.time() - start_time
                    if curr_seconds > 0:
                        remaining_seconds = max(0.0, (total_seconds - curr_seconds) * (elapsed / curr_seconds))
                        eta_str = self._format_eta(remaining_seconds)
                    else:
                        eta_str = "計算中..."
                        
                    # 格式化顯示文字
                    curr_time_str = f"{int(curr_seconds)//60:02d}:{int(curr_seconds)%60:02d}"
                    total_time_str = f"{int(total_seconds)//60:02d}:{int(total_seconds)%60:02d}"
                    
                    progress_text = f"🔄 進度: {percent:.1f}% ({curr_time_str} / {total_time_str}) | 速度: {speed_str} | 剩餘時間: {eta_str}"
                    
                    # 更新主執行緒 UI
                    self.parent.after(0, lambda p=percent, t=progress_text: self._update_progress_ui(p, t))

            process.wait()

            if process.returncode == 0:
                self.parent.after(0, lambda: self.status_label.config(text=f"✅ 轉檔成功：{os.path.basename(out_path)}", fg="green"))
                self.parent.after(0, self._refresh_list)
            else:
                # 尋找錯誤訊息
                err_msg = "請確認檔案格式是否受支援"
                for l in reversed(last_lines):
                    if any(w in l for w in ["Error", "Invalid", "Unable", "Failed", "error"]):
                        err_msg = l
                        break
                self.parent.after(0, lambda e=err_msg: self.status_label.config(text=f"❌ 轉檔失敗：{e}", fg="red"))
        except Exception as e:
            self.parent.after(0, lambda: self.status_label.config(text=f"❌ 錯誤：{str(e)}", fg="red"))
        finally:
            self.parent.after(0, lambda: self.convert_btn.config(state="normal"))
            self.parent.after(0, lambda: self.progress_bar.pack_forget())
            self.parent.after(0, lambda: self.progress_label.pack_forget())

    def _get_video_duration(self, file_path):
        # 1. 嘗試使用 ffprobe 快速取得時長
        try:
            cmd = ["ffprobe", "-v", "error", "-show_entries", "format=duration", "-of", "default=noprint_wrappers=1:nokey=1", file_path]
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=5)
            if res.returncode == 0 and res.stdout.strip():
                return float(res.stdout.strip())
        except Exception:
            pass
        
        # 2. 如果 ffprobe 失敗，嘗試使用 ffmpeg 解析標頭
        try:
            cmd = ["ffmpeg", "-i", file_path]
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=5)
            match = re.search(r"Duration:\s*(\d+):(\d+):(\d+\.\d+)", res.stderr)
            if match:
                h, m, s = float(match.group(1)), float(match.group(2)), float(match.group(3))
                return h * 3600 + m * 60 + s
        except Exception:
            pass
        return 0.0

    def _update_progress_ui(self, value, text):
        self.progress_bar['value'] = value
        self.progress_label.config(text=text)

    def _format_eta(self, seconds):
        s = int(seconds)
        h = s // 3600
        m = (s % 3600) // 60
        sec = s % 60
        if h > 0:
            return f"{h:02d}:{m:02d}:{sec:02d}"
        return f"{m:02d}:{sec:02d}"

    def _escape_ffmpeg_path(self, path):
        # 將 Windows 反斜線 \ 轉換為正斜線 /
        p = path.replace("\\", "/")
        # 處理 Windows 磁碟機號的冒號 (例如 C: 轉換為 C\:)，防範 FFmpeg 濾鏡解析出錯
        if ":" in p:
            drive, rest = p.split(":", 1)
            p = f"{drive}\\:{rest}"
        return p


# ===========================================================================
# MP3MergerTab：MP3 合併工具的完整 UI 類別
# ===========================================================================
class MP3MergerTab:
    def __init__(self, parent, download_path_var):
        self.parent = parent
        self.download_path_var = download_path_var
        self.staged_files = []  # 存儲路徑
        self.staged_durations = []  # 存儲每首歌的長度 (ms)
        self.total_ms = 0
        self.fade_var = tk.BooleanVar(value=False)
        self.fade_sec = tk.IntVar(value=3)
        
        # 監聽融合設定，即時更新時間軸
        self.fade_var.trace_add("write", lambda *args: self._update_total())
        self.fade_sec.trace_add("write", lambda *args: self._update_total())
        
        self._update_job = None
        self._seeking = False
        self._current_song_idx = -1
        
        # 使用雙播放器以實現預覽重疊
        self.players = [MCIPlayer(alias="merger_p1"), MCIPlayer(alias="merger_p2")]
        self.active_player_idx = 0
        self._next_song_triggered = False 
        self._build_ui()

    def _build_ui(self):
        # 頂部：資料夾選擇 (與裁剪工具一致)
        folder_frame = tk.Frame(self.parent)
        folder_frame.pack(fill="x", padx=10, pady=(10, 0))
        tk.Label(folder_frame, text="📁 歌曲資料夾：", font=("Microsoft JhengHei", 10)).pack(side="left")
        self.path_entry = tk.Entry(folder_frame, textvariable=self.download_path_var, font=("Microsoft JhengHei", 10))
        self.path_entry.pack(side="left", fill="x", expand=True, padx=5)
        tk.Button(folder_frame, text="選擇", command=self._browse_folder).pack(side="left", padx=2)
        tk.Button(folder_frame, text="開啟", command=self._open_folder).pack(side="left", padx=2)

        # 使用三欄佈局
        main_body = tk.Frame(self.parent)
        main_body.pack(fill="both", expand=True)

        # 1. 左側：來源檔案列表
        left_frame = tk.Frame(main_body, width=220)
        left_frame.pack(side="left", fill="y", padx=(10, 5), pady=10)
        left_frame.pack_propagate(False)
        tk.Label(left_frame, text="1. 來源 MP3 (可多選)", font=("Microsoft JhengHei", 10, "bold")).pack(anchor="w")
        tk.Label(left_frame, text="💡 按住 Shift 點選前後可連續選取", font=("Microsoft JhengHei", 8), fg="#666").pack(anchor="w")
        
        # 先 pack 下方按鈕，確保不會被清單擠掉
        tk.Button(left_frame, text="🔄 重新整理", command=self._refresh_src_list, font=("Microsoft JhengHei", 9)).pack(side="bottom", fill="x", pady=2)
        tk.Button(left_frame, text="➕ 加入合併清單", command=self._add_to_merge, bg="#4CAF50", fg="white", font=("Microsoft JhengHei", 10, "bold")).pack(side="bottom", fill="x")
        
        # 支援自訂選取行為：單擊切換 + Shift 範圍選取
        self.src_listbox = tk.Listbox(left_frame, font=("Microsoft JhengHei", 9), selectmode="extended")
        self.src_listbox.pack(fill="both", expand=True, pady=5)
        self.src_listbox.bind("<Button-1>", self._on_listbox_click)
        self._last_idx = None

        # 2. 中間：待合併清單 (Staging Area)
        mid_frame = tk.Frame(main_body, width=260)
        mid_frame.pack(side="left", fill="y", padx=5, pady=10)
        mid_frame.pack_propagate(False)
        tk.Label(mid_frame, text="2. 合併清單", font=("Microsoft JhengHei", 10, "bold")).pack(anchor="w")
        
        # 先 pack 下方按鈕
        tk.Button(mid_frame, text="🧹 清除全部歌曲", command=self._clear_all, bg="#757575", fg="white").pack(side="bottom", fill="x", pady=2)
        tk.Button(mid_frame, text="🗑️ 移除選定歌曲", command=self._remove_from_merge, bg="#f44336", fg="white").pack(side="bottom", fill="x", pady=2)
        
        btn_grid = tk.Frame(mid_frame)
        btn_grid.pack(side="bottom", fill="x")
        tk.Button(btn_grid, text="🔼 上移", command=lambda: self._move_item(-1), width=10).pack(side="left", padx=2, expand=True, fill="x")
        tk.Button(btn_grid, text="🔽 下移", command=lambda: self._move_item(1), width=10).pack(side="left", padx=2, expand=True, fill="x")

        self.merge_listbox = tk.Listbox(mid_frame, font=("Microsoft JhengHei", 9))
        self.merge_listbox.pack(fill="both", expand=True, pady=5)

        # 3. 右側：預覽與執行
        right_frame = tk.Frame(self.parent)
        right_frame.pack(side="left", fill="both", expand=True, padx=(5, 10), pady=10)
        tk.Label(right_frame, text="3. 預覽與輸出", font=("Microsoft JhengHei", 10, "bold")).pack(anchor="w")

        # 播放控制
        ctrl_frame = tk.Frame(right_frame)
        ctrl_frame.pack(fill="x", pady=5)
        tk.Button(ctrl_frame, text="⏮ -5s", command=lambda: self._seek_relative(-5000)).pack(side="left", padx=1)
        tk.Button(ctrl_frame, text="◀ -1s", command=lambda: self._seek_relative(-1000)).pack(side="left", padx=1)
        self.play_btn = tk.Button(ctrl_frame, text="▶ 播放合併效果", command=self._toggle_play, bg="#2196F3", fg="white", width=15)
        self.play_btn.pack(side="left", padx=5)
        tk.Button(ctrl_frame, text="⏹ 停止", command=self._stop).pack(side="left", padx=2)
        tk.Button(ctrl_frame, text="+1s ▶", command=lambda: self._seek_relative(1000)).pack(side="left", padx=1)
        tk.Button(ctrl_frame, text="+5s ⏭", command=lambda: self._seek_relative(5000)).pack(side="left", padx=1)

        self.time_label = tk.Label(right_frame, text="00:00 / 00:00", font=("Microsoft JhengHei", 10))
        self.time_label.pack(pady=2)

        # 虛擬進度條 (Canvas)
        canvas_outer = tk.Frame(right_frame, bg="#888", pady=1)
        canvas_outer.pack(fill="x", pady=5)
        self.merge_canvas = tk.Canvas(canvas_outer, height=40, bg="#eee", highlightthickness=0, cursor="hand2")
        self.merge_canvas.pack(fill="both", expand=True)
        self.merge_canvas.bind("<ButtonPress-1>", self._canvas_click)
        self.merge_canvas.bind("<Configure>", lambda e: self._draw_canvas())

        # 輸出設定
        # 儲存資料夾 Row (預設為 download_path_var 的值，並可自主選擇)
        self.out_folder_row = tk.Frame(right_frame)
        self.out_folder_row.pack(fill="x", pady=5)
        tk.Label(self.out_folder_row, text="儲存資料夾：", font=("Microsoft JhengHei", 10, "bold"), width=12, anchor="w").pack(side="left")
        self.out_folder_var = tk.StringVar(value=self.download_path_var.get())
        self.out_folder_entry = tk.Entry(self.out_folder_row, textvariable=self.out_folder_var, font=("Microsoft JhengHei", 10), state="readonly")
        self.out_folder_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        tk.Button(self.out_folder_row, text="選擇", command=self._browse_out_folder, bg="#E91E63", fg="white", font=("Microsoft JhengHei", 9)).pack(side="left")

        tk.Label(right_frame, text="合併後檔名：").pack(anchor="w", pady=(10, 0))
        self.out_entry = tk.Entry(right_frame)
        self.out_entry.pack(fill="x", pady=5)
        self.out_entry.insert(0, "merged_audio")

        # 融合效果設定 (Crossfade)
        fade_frame = tk.Frame(right_frame)
        fade_frame.pack(fill="x", pady=5)
        tk.Checkbutton(fade_frame, text="✨ 啟用融合效果 (Crossfade)", variable=self.fade_var, 
                       font=("Microsoft JhengHei", 10, "bold"), fg="#1976D2").pack(side="left")
        tk.Label(fade_frame, text="  融合秒數：").pack(side="left")
        tk.Spinbox(fade_frame, from_=1, to=5, textvariable=self.fade_sec, width=5).pack(side="left")
        tk.Label(fade_frame, text="(註：融合需重新轉檔，速度較慢)", font=("Microsoft JhengHei", 8), fg="gray").pack(side="left", padx=5)

        self.merge_btn = tk.Button(right_frame, text="🚀 開始合併所有歌曲", command=self._do_merge, 
                                   bg="#4CAF50", fg="white", font=("Microsoft JhengHei", 12, "bold"), height=2)
        self.merge_btn.pack(fill="x", pady=10)
        
        self.status_label = tk.Label(right_frame, text="", fg="blue")
        self.status_label.pack()

        # 初始化載入檔案
        self._refresh_src_list()

    def _browse_folder(self):
        d = filedialog.askdirectory(initialdir=self.download_path_var.get())
        if d:
            self.download_path_var.set(d)
            self._refresh_src_list()

    def _open_folder(self):
        d = self.download_path_var.get()
        if os.path.exists(d):
            os.startfile(d)

    def _browse_out_folder(self):
        initial_dir = self.out_folder_var.get() or self.download_path_var.get()
        folder = filedialog.askdirectory(initialdir=initial_dir)
        if folder:
            is_writable = True
            test_file = os.path.join(folder, ".write_test")
            try:
                with open(test_file, "w") as f:
                    pass
                os.remove(test_file)
            except Exception:
                is_writable = False
            
            if not is_writable:
                messagebox.showerror("權限錯誤", "所選資料夾為唯讀或無寫入權限，請選擇其他儲存位置。")
            else:
                self.out_folder_var.set(folder)

    def _refresh_src_list(self):
        folder = self.download_path_var.get()
        self.src_listbox.delete(0, tk.END)
        self._last_idx = None
        if os.path.exists(folder):
            files = sorted([f for f in os.listdir(folder) if f.lower().endswith('.mp3')])
            for f in files:
                self.src_listbox.insert(tk.END, f)

    def _on_listbox_click(self, event):
        """自定義選取邏輯：單擊即切換(Toggle)，Shift 則執行範圍選取"""
        idx = self.src_listbox.nearest(event.y)
        if idx < 0: return
        
        # 判斷是否按住 Shift (state & 0x0001)
        if (event.state & 0x0001) and self._last_idx is not None:
            # 範圍選取
            start = min(self._last_idx, idx)
            end = max(self._last_idx, idx)
            # 先清除其他，再選取範圍（或根據需求決定是否保留舊有選取）
            # 這裡採標準 Shift 行為：選取該區間
            for i in range(start, end + 1):
                self.src_listbox.selection_set(i)
        else:
            # 單擊切換 (Toggle)
            if self.src_listbox.selection_includes(idx):
                self.src_listbox.selection_clear(idx)
            else:
                self.src_listbox.selection_set(idx)
            self._last_idx = idx
            
        return "break" # 阻止 Tkinter 預設行為

    def _add_to_merge(self):
        sel = self.src_listbox.curselection()
        if not sel: return
        
        folder = self.download_path_var.get()
        failed_count = 0
        
        for i in sel:
            fname = self.src_listbox.get(i)
            fpath = os.path.join(folder, fname)
            
            # 使用唯一 Alias 獲取時長，避免與主播放器或其他實例衝突
            unique_alias = f"info_{int(time.time()*1000)}_{i}"
            temp_player = MCIPlayer(alias=unique_alias)
            if temp_player.open(fpath):
                dur = temp_player.get_length()
                temp_player.close()
                self.staged_files.append(fpath)
                self.staged_durations.append(dur)
                self.merge_listbox.insert(tk.END, f"[{self._fmt_ms(dur)}] {fname}")
            else:
                failed_count += 1
        
        self._update_total()
        self._update_out_filename()
        if failed_count > 0:
            messagebox.showwarning("警告", f"有 {failed_count} 個檔案無法讀取資訊。")

    def _remove_from_merge(self):
        sel = self.merge_listbox.curselection()
        if not sel: return
        idx = sel[0]
        self._stop()
        self.staged_files.pop(idx)
        self.staged_durations.pop(idx)
        self.merge_listbox.delete(idx)
        self._update_total()
        self._update_out_filename()

    def _move_item(self, direction):
        sel = self.merge_listbox.curselection()
        if not sel: return
        idx = sel[0]
        new_idx = idx + direction
        if 0 <= new_idx < len(self.staged_files):
            self._stop()
            # 交換資料
            self.staged_files[idx], self.staged_files[new_idx] = self.staged_files[new_idx], self.staged_files[idx]
            self.staged_durations[idx], self.staged_durations[new_idx] = self.staged_durations[new_idx], self.staged_durations[idx]
            # 更新 Listbox
            txt = self.merge_listbox.get(idx)
            self.merge_listbox.delete(idx)
            self.merge_listbox.insert(new_idx, txt)
            self.merge_listbox.selection_set(new_idx)
            self._draw_canvas()
            self._update_out_filename()

    def _update_total(self):
        n = len(self.staged_files)
        if n == 0:
            self.total_ms = 0
        elif self.fade_var.get() and n > 1:
            fade_ms = self.fade_sec.get() * 1000
            self.total_ms = max(0, sum(self.staged_durations) - (n - 1) * fade_ms)
        else:
            self.total_ms = sum(self.staged_durations)
        self.time_label.config(text=f"00:00 / {self._fmt_ms(self.total_ms)}")
        self._draw_canvas()

    def _update_out_filename(self):
        """以合併清單第一首歌的檔名作為預設輸出檔名"""
        if self.staged_files:
            stem = os.path.splitext(os.path.basename(self.staged_files[0]))[0]
            self.out_entry.delete(0, tk.END)
            self.out_entry.insert(0, f"{stem}_merged")
        else:
            self.out_entry.delete(0, tk.END)
            self.out_entry.insert(0, "merged_audio")

    def _clear_all(self):
        if not self.staged_files: return
        if messagebox.askyesno("確認", "確定要清除合併清單中的所有歌曲嗎？"):
            self._stop()
            self.staged_files = []
            self.staged_durations = []
            self.merge_listbox.delete(0, tk.END)
            self._update_total()

    def _draw_canvas(self, current_ms=0):
        c = self.merge_canvas
        w = c.winfo_width()
        h = c.winfo_height()
        if w <= 1 or self.total_ms <= 0: return
        c.delete("all")
        if not self.staged_durations: return

        colors       = ["#81C784", "#64B5F6", "#FFD54F", "#BA68C8", "#FF8A65", "#4DB6AC"]
        fade_colors  = ["#43A047", "#1E88E5", "#F9A825", "#8E24AA", "#E64A19", "#00897B"]
        do_fade = self.fade_var.get() and len(self.staged_files) > 1
        fade_ms = (self.fade_sec.get() * 1000) if do_fade else 0

        acc_virtual = 0  # 虛擬時間軸累積位置 (ms)
        for i, dur in enumerate(self.staged_durations):
            is_last = (i == len(self.staged_durations) - 1)
            eff_dur = dur - fade_ms if (do_fade and not is_last) else dur

            x0 = int(acc_virtual / self.total_ms * w)
            x1 = int((acc_virtual + dur) / self.total_ms * w)
            c.create_rectangle(x0, 0, x1, h, fill=colors[i % len(colors)], outline="")

            # 融合重疊區塊：上半層使用較深色呈現漸變感
            if do_fade and i > 0:
                fade_x0 = x0
                fade_x1 = int((acc_virtual + fade_ms) / self.total_ms * w)
                # 以斜線漸層模擬重疊 (tkinter無漸層，用半透明窄條代替)
                step = max(1, (fade_x1 - fade_x0) // 10)
                prev_color = fade_colors[(i-1) % len(fade_colors)]
                cur_color  = fade_colors[i % len(fade_colors)]
                for s in range(fade_x0, fade_x1, step):
                    ratio = (s - fade_x0) / max(1, fade_x1 - fade_x0)
                    stripe_color = prev_color if ratio < 0.5 else cur_color
                    c.create_rectangle(s, 0, s + step, h // 2, fill=stripe_color, outline="")
                # 標示融合區
                mid = (fade_x0 + fade_x1) // 2
                c.create_text(mid, h // 2, text="↔", font=("Microsoft JhengHei", 8), fill="white", anchor="center")

            # 標示起始時間
            if x1 - x0 > 45:
                t_str = self._fmt_ms(int(acc_virtual))
                c.create_text(x0 + 3, h - 4, text=t_str, anchor="sw", font=("Microsoft JhengHei", 8), fill="#222")

            acc_virtual += eff_dur

        # 播放進度條
        xp = int(current_ms / self.total_ms * w) if self.total_ms > 0 else 0
        c.create_rectangle(xp - 2, 0, xp + 2, h, fill="#f44336", outline="")

    def _get_info_at(self, ms):
        """根據總時間點找到是對應哪首歌以及在該歌中的相對時間"""
        if not self.staged_durations: return -1, 0
        
        fade_ms = (self.fade_sec.get() * 1000) if self.fade_var.get() else 0
        acc = 0
        for i, dur in enumerate(self.staged_durations):
            # 該首歌在虛擬時間軸上的「可用」長度（扣除與下一首的重疊部分）
            # 最後一首不扣除
            effective_dur = dur - fade_ms if i < len(self.staged_durations)-1 else dur
            if acc <= ms < acc + effective_dur + (fade_ms if i < len(self.staged_durations)-1 else 0):
                return i, ms - acc
            acc += effective_dur
        return len(self.staged_durations) - 1, self.staged_durations[-1]

    def _toggle_play(self):
        if not self.staged_files: return
        p = self.players[self.active_player_idx]
        mode = p.get_mode()
        if mode == "playing":
            for player in self.players: player.pause()
            self.play_btn.config(text="▶ 播放合併效果")
        elif mode == "paused":
            for player in self.players: player.resume()
            self.play_btn.config(text="⏸ 暫停")
            self._start_loop()
        else:
            self._play_at(0)

    def _play_at(self, total_ms):
        if not self.staged_files: return
        idx, rel_ms = self._get_info_at(total_ms)
        if idx < 0: return
        
        # 停止所有播放
        for p in self.players: p.stop()
        
        self.active_player_idx = 0
        self._current_song_idx = idx
        self._next_song_triggered = False
        
        p = self.players[self.active_player_idx]
        if p.open(self.staged_files[idx]):
            p.seek(rel_ms)
            p.play()
            self.play_btn.config(text="⏸ 暫停")
            self._start_loop()

    def _stop(self):
        for p in self.players: p.stop()
        self._current_song_idx = -1
        self._next_song_triggered = False
        self.play_btn.config(text="▶ 播放合併效果")
        if self._update_job:
            self.parent.after_cancel(self._update_job)
            self._update_job = None
        self.time_label.config(text=f"00:00 / {self._fmt_ms(self.total_ms)}")
        self._draw_canvas(0)

    def _start_loop(self):
        if self._update_job: self.parent.after_cancel(self._update_job)
        self._do_update()

    def _do_update(self):
        p_active = self.players[self.active_player_idx]
        mode = p_active.get_mode()

        if mode not in ("playing", "paused"):
            if self._next_song_triggered and self._current_song_idx + 1 < len(self.staged_files):
                # 切換到預加載的下一首
                self.active_player_idx = 1 - self.active_player_idx
                self._current_song_idx += 1
                self._next_song_triggered = False
                self._update_job = self.parent.after(80, self._do_update)
            else:
                self._stop()
            return

        if mode == "playing":
            rel_pos = p_active.get_position()
            do_fade = self.fade_var.get()
            fade_ms = (self.fade_sec.get() * 1000) if do_fade else 0

            # 計算虛擬總進度
            acc = 0
            for i in range(self._current_song_idx):
                eff = self.staged_durations[i] - fade_ms if (do_fade and i < len(self.staged_durations)-1) else self.staged_durations[i]
                acc += eff
            total_pos = min(acc + rel_pos, self.total_ms)

            self.time_label.config(text=f"{self._fmt_ms(total_pos)} / {self._fmt_ms(self.total_ms)}")
            self._draw_canvas(total_pos)

            dur_current = self.staged_durations[self._current_song_idx]
            time_left = dur_current - rel_pos
            has_next = (self._current_song_idx + 1 < len(self.staged_files))

            # ---- 預加載下一首 ----
            PRELOAD_MS = max(fade_ms, 3000)  # 稍微提早一點預加載，確保流暢
            if has_next and time_left <= PRELOAD_MS and not self._next_song_triggered:
                self._next_song_triggered = True
                next_idx = self._current_song_idx + 1
                other_p = self.players[1 - self.active_player_idx]
                if do_fade:
                    if other_p.open(self.staged_files[next_idx]):
                        other_p.set_volume(0) # 融合模式初始音量 0
                        other_p.play()
                else:
                    if other_p.open(self.staged_files[next_idx]):
                        other_p.set_volume(1000) # 非融合模式確保音量 1000
                        # 暫不 play

            # ---- 融合音量漸變 ----
            if do_fade and self._next_song_triggered and time_left <= fade_ms:
                fade_ratio = max(0.0, min(1.0, time_left / max(fade_ms, 1)))
                try:
                    p_active.set_volume(int(fade_ratio * 1000))
                    self.players[1 - self.active_player_idx].set_volume(int((1.0 - fade_ratio) * 1000))
                except Exception: pass

            # ---- 當前歌曲結束：切換主播放器 ----
            if time_left <= 120: # 稍微調大容錯，減少卡頓感
                if self._next_song_triggered:
                    try:
                        self.players[1 - self.active_player_idx].set_volume(1000)
                    except Exception: pass
                    
                    if not do_fade:
                        self.players[1 - self.active_player_idx].play()
                        
                    self.active_player_idx = 1 - self.active_player_idx
                    self._current_song_idx += 1
                    self._next_song_triggered = False
                elif has_next:
                    # 安全備案：不應發生
                    self._play_at(total_pos + 1)
                    return
                else:
                    self._stop()
                    return

        self._update_job = self.parent.after(80, self._do_update)

    def _canvas_click(self, event):
        if self.total_ms <= 0: return
        w = self.merge_canvas.winfo_width()
        ms = int((event.x / w) * self.total_ms)
        self._play_at(ms)

    def _seek_relative(self, delta):
        if self._current_song_idx == -1: return
        fade_ms = (self.fade_sec.get() * 1000) if self.fade_var.get() else 0
        acc = 0
        for i in range(self._current_song_idx):
            eff = self.staged_durations[i] - fade_ms if (self.fade_var.get() and i < len(self.staged_durations)-1) else self.staged_durations[i]
            acc += eff
        p_active = self.players[self.active_player_idx]
        cur_total = acc + p_active.get_position()
        new_total = max(0, min(cur_total + delta, self.total_ms))
        self._play_at(new_total)

    def _do_merge(self):
        if not self.staged_files: return
        out_name = self.out_entry.get().strip()
        if not out_name: return
        target_folder = self.out_folder_var.get() or self.download_path_var.get()
        os.makedirs(target_folder, exist_ok=True)
        out_path = os.path.join(target_folder, out_name + ".mp3")
        
        # 處理檔名重複
        base = out_name
        counter = 1
        while os.path.exists(out_path):
            out_name = f"{base}({counter})"
            out_path = os.path.join(self.download_path_var.get(), out_name + ".mp3")
            counter += 1

        self._stop()
        self.merge_btn.config(state="disabled")
        self.status_label.config(text="合併中，請稍候...", fg="blue")
        threading.Thread(target=self._run_ffmpeg_merge, args=(self.staged_files, out_path), daemon=True).start()

    def _run_ffmpeg_merge(self, files, out_path):
        try:
            do_fade = self.fade_var.get()
            fade_d = self.fade_sec.get()

            if not do_fade:
                # 傳統高速合併 (concat)
                list_file = os.path.join(self.download_path_var.get(), "concat_list.txt")
                with open(list_file, "w", encoding="utf-8") as f:
                    for fp in files:
                        p = fp.replace("'", "'\\''")
                        f.write(f"file '{p}'\n")
                cmd = ["ffmpeg", "-y", "-f", "concat", "-safe", "0", "-i", list_file, "-c", "copy", out_path]
                res = subprocess.run(cmd, capture_output=True, text=True)
                if os.path.exists(list_file): os.remove(list_file)
            else:
                # 融合效果合併 (acrossfade)
                # 對於多個檔案，需要構建 complex_filter
                cmd = ["ffmpeg", "-y"]
                for f in files:
                    cmd.extend(["-i", f])
                
                # 構建濾鏡鏈：[0][1]acrossfade=d=3[a1]; [a1][2]acrossfade=d=3[a2]...
                filter_str = ""
                last_label = "[0]"
                for i in range(1, len(files)):
                    next_label = f"[a{i}]"
                    filter_str += f"{last_label}[{i}]acrossfade=d={fade_d}:c1=tri:c2=tri"
                    if i < len(files) - 1:
                        filter_str += f"{next_label};"
                        last_label = next_label
                
                cmd.extend(["-filter_complex", filter_str, "-b:a", "320k", out_path])
                res = subprocess.run(cmd, capture_output=True, text=True)

            if res.returncode == 0:
                self.parent.after(0, lambda: self.status_label.config(text=f"✅ 合併成功：{os.path.basename(out_path)}", fg="green"))
                self.parent.after(0, self._refresh_src_list)
            else:
                self.parent.after(0, lambda: self.status_label.config(text="❌ 合併失敗 (FFmpeg 錯誤)", fg="red"))
        except Exception as e:
            self.parent.after(0, lambda: self.status_label.config(text=f"❌ 錯誤: {e}", fg="red"))
        finally:
            self.parent.after(0, lambda: self.merge_btn.config(state="normal"))

    @staticmethod
    def _fmt_ms(ms):
        s = int(ms) // 1000
        return f"{s // 60:02d}:{s % 60:02d}"

if __name__ == "__main__":
    root = tk.Tk()
    app = YouTubeDownloaderGUI(root)
    root.mainloop()

