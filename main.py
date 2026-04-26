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

ssl._create_default_https_context = ssl._create_unverified_context
APP_VERSION = "2.0.10"
GITHUB_REPO = "mathced-com/CYT_YTDL"

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
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

class YouTubeDownloaderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"CYT_YouTube 下載器 v{APP_VERSION}")
        self.root.geometry("850x700")
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
        
        self.playlist_vars = []
        self.playlist_entries = []
        
        # 暫停與取消狀態標記
        self.is_paused = False
        self.is_cancelled = False
        
        self.create_widgets()
        self.update_quality_options()
        
        self.check_ffmpeg_environment()
        
        if not HAS_PIL:
            messagebox.showwarning("缺少套件", "系統缺少 Pillow 套件，將無法顯示影片封面。")

    def resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_path, relative_path)

    def _on_tab_changed(self, event):
        sel = self.notebook.select()
        if not sel: return
        text = self.notebook.tab(sel, "text")
        if "裁剪" in text:
            self.trimmer._refresh_list()
        elif "合併" in text:
            self.merger._refresh_src_list()

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
            
        tk.Label(header_frame, text=f"CYT_YouTube 下載器 v{APP_VERSION}", font=("Arial", 18, "bold"), bg="white").pack(side="left")

        # === 標簿頁 (Notebook) ===
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Tab 1: YouTube 下載器
        tab_download = tk.Frame(self.notebook)
        self.notebook.add(tab_download, text="  ⬇️ YouTube 下載器  ")
        
        # Tab 2: MP3 裁剪工具
        tab_trim = tk.Frame(self.notebook)
        self.notebook.add(tab_trim, text="  ✂️ MP3 裁剪工具  ")
        self.trimmer = MP3TrimmerTab(tab_trim, self.download_path)
        
        # Tab 3: MP3 合併工具
        tab_merge = tk.Frame(self.notebook)
        self.notebook.add(tab_merge, text="  🔗 MP3 合併工具  ")
        self.merger = MP3MergerTab(tab_merge, self.download_path)
        
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        # === 以下為 Tab 1 (下載器) 的 UI ===
        parent = tab_download
        
        url_frame = tk.Frame(parent)
        url_frame.pack(fill="x", padx=15, pady=5)
        tk.Label(url_frame, text="網址：", font=("Arial", 12)).pack(side="left")
        
        self.url_entry = tk.Entry(url_frame, width=35, font=("Arial", 10))
        self.url_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        # 按鈕順序：貼上、清除、解析
        tk.Button(url_frame, text="貼上", command=self.paste_url, font=("Arial", 10), bg="#FFEB3B").pack(side="left", padx=2)
        
        self.clear_btn = tk.Button(url_frame, text="清除網址", command=self.clear_url, font=("Arial", 10))
        self.clear_btn.pack(side="left", padx=2)

        self.analyze_btn = tk.Button(url_frame, text="解析網址", command=self.start_analyze, bg="#2196F3", fg="white", font=("Arial", 10, "bold"))
        self.analyze_btn.pack(side="left", padx=2)
        
        # 步驟提示
        hint_text = "執行步驟：\n一、貼上Youtube網址\n二、點擊「解析網址」\n三、點擊「開始下載」"
        hint_label = tk.Label(url_frame, text=hint_text, fg="#E91E63", font=("Arial", 9, "bold"), justify="left")
        hint_label.pack(side="left", padx=5)
        
        # 先建立底部框架並鎖定在視窗最下方，保證不被清單擠出畫面
        bottom_frame = tk.Frame(parent)
        bottom_frame.pack(side="bottom", fill="x", pady=5)
        
        self.info_frame = tk.LabelFrame(parent, text="影片預覽 / 播放清單", font=("Arial", 10))
        self.info_frame.pack(fill="both", expand=True, padx=15, pady=5)
        
        self.title_label = tk.Label(self.info_frame, text="請輸入網址並點選「解析網址」", fg="gray", wraplength=680, justify="left")
        self.title_label.pack(pady=5, padx=10)
        
        self.select_btn_frame = tk.Frame(self.info_frame)
        tk.Button(self.select_btn_frame, text="全部勾選", command=self.select_all, font=("Arial", 10), bg="#4CAF50", fg="white").pack(side="left", padx=5)
        tk.Button(self.select_btn_frame, text="取消全選", command=self.deselect_all, font=("Arial", 10)).pack(side="left", padx=5)
        self.select_btn_frame.pack(pady=3)
        self.select_btn_frame.pack_forget()
        
        self.list_frame = ScrollableFrame(self.info_frame)
        
        format_frame = tk.Frame(bottom_frame)
        format_frame.pack(fill="x", padx=15, pady=3)
        tk.Label(format_frame, text="格式：", font=("Arial", 12)).pack(side="left")
        tk.Radiobutton(format_frame, text="MP4", variable=self.format_choice, value="mp4", command=self.update_quality_options).pack(side="left", padx=2)
        tk.Radiobutton(format_frame, text="MP3", variable=self.format_choice, value="mp3", command=self.update_quality_options).pack(side="left", padx=2)
        
        tk.Label(format_frame, text="   品質：", font=("Arial", 12)).pack(side="left")
        self.quality_combo = ttk.Combobox(format_frame, textvariable=self.quality_choice, state="readonly", width=18)
        self.quality_combo.pack(side="left", padx=5)
        
        path_frame = tk.Frame(bottom_frame)
        path_frame.pack(fill="x", padx=15, pady=3)
        tk.Label(path_frame, text="儲存：", font=("Arial", 12)).pack(side="left")
        self.path_entry = tk.Entry(path_frame, textvariable=self.download_path, width=40, state="readonly", font=("Arial", 10))
        self.path_entry.pack(side="left", padx=5, fill="x", expand=True)
        tk.Button(path_frame, text="選擇", command=self.browse_folder).pack(side="left", padx=2)
        tk.Button(path_frame, text="開啟", command=self.open_download_folder, bg="#9C27B0", fg="white").pack(side="left", padx=2)
        
        status_frame = tk.Frame(bottom_frame)
        status_frame.pack(fill="x", padx=15, pady=3)
        self.progress_bar = ttk.Progressbar(status_frame, orient="horizontal", length=740, mode="determinate")
        self.progress_bar.pack(pady=2)
        self.status_label = tk.Label(status_frame, text="等待解析...", fg="blue", font=("Arial", 10))
        self.status_label.pack(pady=2)
        
        # 執行與控制按鈕區
        btn_frame = tk.Frame(bottom_frame)
        btn_frame.pack(pady=5)
        self.download_btn = tk.Button(btn_frame, text="開始下載", font=("Arial", 12, "bold"), bg="#4CAF50", fg="white", width=12, command=self.start_download, state="disabled")
        self.download_btn.pack(side="left", padx=5)
        
        self.pause_btn = tk.Button(btn_frame, text="暫停", font=("Arial", 10), command=self.toggle_pause, state="disabled", width=8)
        self.pause_btn.pack(side="left", padx=5)
        
        self.cancel_btn = tk.Button(btn_frame, text="取消", font=("Arial", 10), command=self.cancel_download, state="disabled", bg="#f44336", fg="white", width=8)
        self.cancel_btn.pack(side="left", padx=5)
        
        if not getattr(sys, 'frozen', False):
            tk.Button(btn_frame, text="更新 yt-dlp (開發者模式)", command=self.update_ytdlp).pack(side="left", padx=15)
            
        tk.Button(btn_frame, text="檢查主程式更新", command=self.check_app_update, bg="#FF9800", fg="white").pack(side="left", padx=5)

    def select_all(self):
        for var in self.playlist_vars:
            var.set(True)
            
    def deselect_all(self):
        for var in self.playlist_vars:
            var.set(False)

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

    def update_quality_options(self):
        if self.format_choice.get() == "mp4":
            options = ["最高畫質 (自動)", "1080p", "720p", "480p", "360p"]
        else:
            options = ["最高音質 (320k)", "標準音質 (192k)", "普通音質 (128k)"]
        self.quality_combo['values'] = options
        self.quality_combo.current(0)

    def browse_folder(self):
        folder = filedialog.askdirectory(initialdir=self.download_path.get())
        if folder:
            self.download_path.set(folder)

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
                self.root.after(0, lambda: messagebox.showerror("錯誤", f"FFmpeg 下載失敗：\n{e}"))
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

    def check_app_update(self):
        self.update_progress_ui(0, "正在檢查主程式更新...", "blue")
        def run_check():
            try:
                req = urllib.request.Request(f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest", headers={'User-Agent': 'Mozilla/5.0'})
                response = urllib.request.urlopen(req, timeout=10)
                data = json.loads(response.read().decode('utf-8'))
                latest_version = data.get("tag_name", "").replace("v", "")
                
                if not latest_version:
                    self.root.after(0, lambda: self.update_progress_ui(0, "無法取得版本資訊", "red"))
                    return
                    
                if latest_version != APP_VERSION:
                    assets = data.get("assets", [])
                    download_url = None
                    for asset in assets:
                        if asset.get("name") == "CYT_YTDL.exe":
                            download_url = asset.get("browser_download_url")
                            break
                            
                    if download_url:
                        self.root.after(0, lambda: self.prompt_update(latest_version, download_url))
                    else:
                        self.root.after(0, lambda: messagebox.showinfo("發現新版本", f"目前最新版本為 {latest_version}，但開發者尚未上傳執行檔。"))
                        self.root.after(0, lambda: self.update_progress_ui(0, "檢查完畢", "blue"))
                else:
                    self.root.after(0, lambda: messagebox.showinfo("檢查更新", "您目前使用的已經是最新版本！"))
                    self.root.after(0, lambda: self.update_progress_ui(0, "準備就緒", "blue"))
            except urllib.error.HTTPError as e:
                if e.code == 404:
                    self.root.after(0, lambda: messagebox.showinfo("檢查更新", "專案尚未發布任何版本 (Release)。"))
                    self.root.after(0, lambda: self.update_progress_ui(0, "無可用更新", "blue"))
                else:
                    self.root.after(0, lambda: self.update_progress_ui(0, f"檢查失敗: {e}", "red"))
            except Exception as e:
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

                new_exe_name = "CYT_YTDL_update.exe"
                new_exe_path = os.path.join(self.app_dir, new_exe_name)
                urllib.request.urlretrieve(download_url, new_exe_path, reporthook=reporthook)
                
                self.root.after(0, lambda: self.update_progress_ui(100, "新版本下載完成！等待確認重啟...", "green"))
                
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
                self.root.after(0, lambda: messagebox.showerror("錯誤", f"更新失敗：\n{e}"))
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
        url = self.url_entry.get().strip()
        if not url:
            messagebox.showwarning("警告", "請輸入 YouTube 網址！")
            return
            
        self.analyze_btn.config(state="disabled")
        self.download_btn.config(state="disabled")
        self.update_progress_ui(0, "正在解析網址與抓取標題，請稍候...", "blue")
        self.title_label.config(text="解析中...")
        self.list_frame.pack_forget()
        self.select_btn_frame.pack_forget()
        
        threading.Thread(target=self.process_analyze, args=(url,), daemon=True).start()
        
    def process_analyze(self, url):
        import urllib.parse
        parsed_url = urllib.parse.urlparse(url)
        query_params = urllib.parse.parse_qs(parsed_url.query)
        if 'list' in query_params:
            playlist_id = query_params['list'][0]
            # YT Mix (合輯) 的清單 ID 通常以 RD 開頭，這種清單不能直接轉換為 /playlist?list= 否則會報錯
            if not playlist_id.startswith("RD"):
                url = f"https://www.youtube.com/playlist?list={playlist_id}"

        ydl_opts = {
            'extract_flat': True, 
            'quiet': True
        }
        try:
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                info = ydl.extract_info(url, download=False)
                
            self.video_info = info
            
            if 'entries' in info:
                self.is_playlist = True
                entries = list(info.get('entries') or [])
                total = len(entries)
                if total > 50:
                    def ask_playlist_action():
                        msg = f"偵測到龐大的播放清單 (共 {total} 部影片)！\n\n請選擇後續動作：\n\n【是】載入前 50 筆清單讓我手動勾選。\n【否】不展開清單，直接下載全部。\n【取消】取消解析。"
                        res = messagebox.askyesnocancel("播放清單處理方式", msg)
                        if res is True:
                            self.show_playlist(info.get('title', '播放清單'), entries[:50])
                        elif res is False:
                            self.show_playlist_summary(info.get('title', '播放清單'), entries)
                        else:
                            self.update_progress_ui(0, "已取消解析", "blue")
                            self.analyze_btn.config(state="normal")
                            self.title_label.config(text="請輸入網址並點選「解析網址」")
                    self.root.after(0, ask_playlist_action)
                else:
                    self.root.after(0, lambda: self.show_playlist(info.get('title', '播放清單'), entries))
            else:
                self.is_playlist = False
                title = info.get('title', '未知影片標題')
                dur_str = self.format_duration(info.get('duration'))
                thumb_url = info.get('thumbnail')
                self.root.after(0, lambda: self.show_single_video(title, dur_str, thumb_url))
                
        except Exception as e:
            self.root.after(0, lambda: self.title_label.config(text="解析失敗，請確認網址是否正確。"))
            self.root.after(0, lambda: self.update_progress_ui(0, "發生錯誤", "red"))
            self.root.after(0, lambda: messagebox.showerror("錯誤", f"解析失敗：\n{str(e)}"))
        finally:
            self.root.after(0, lambda: self.analyze_btn.config(state="normal"))

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
        
        txt_label = tk.Label(row_frame, text=f"💡 為避免介面卡頓，已隱藏清單明細。\n\n共有 {len(entries)} 部影片已準備就緒！\n請確認下方的「格式」與「儲存資料夾」無誤後，點擊「開始下載」即可自動下載全集。", justify="left", wraplength=500, font=("Arial", 11, "bold"), fg="#2196F3")
        txt_label.pack(side="left", anchor="w", padx=20, pady=20)
        
        self.playlist_entries = entries
        self.playlist_vars = []  # 空的 var 代表總結模式 (全選)
        self.download_btn.config(state="normal")
        self.update_progress_ui(0, "解析完成！點擊「開始下載」以下載全集", "green")

    def show_single_video(self, title, dur_str, thumb_url):
        self.title_label.config(text="【單一影片解析結果】")
        for widget in self.list_frame.scrollable_frame.winfo_children():
            widget.destroy()
            
        self.list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        row_frame = tk.Frame(self.list_frame.scrollable_frame, pady=5)
        row_frame.pack(fill="x", anchor="w")
        
        thumb_label = tk.Label(row_frame, text="無圖片", bg="#e0e0e0", width=14, height=3)
        thumb_label.pack(side="left", padx=5)
        
        txt_label = tk.Label(row_frame, text=f"{title}\n時間: {dur_str.strip() if dur_str else '未知'}", justify="left", wraplength=500, font=("Arial", 10))
        txt_label.pack(side="left", anchor="w", padx=10)
        
        if HAS_PIL and thumb_url:
            threading.Thread(target=self.load_thumbnail, args=(thumb_url, thumb_label), daemon=True).start()
        
        self.update_progress_ui(0, "解析完成！請確認資訊後點擊「開始下載」", "green")
        self.download_btn.config(state="normal")

    def show_playlist(self, title, entries):
        self.title_label.config(text=f"【播放清單】\n{title} (共 {len(entries)} 部影片)")
        
        for widget in self.list_frame.scrollable_frame.winfo_children():
            widget.destroy()
            
        self.playlist_vars.clear()
        self.playlist_entries = entries
        
        for i, entry in enumerate(entries):
            var = tk.BooleanVar(value=True)
            self.playlist_vars.append(var)
            
            row_frame = tk.Frame(self.list_frame.scrollable_frame, pady=3)
            row_frame.pack(fill="x", anchor="w")
            
            chk = tk.Checkbutton(row_frame, variable=var)
            chk.pack(side="left", padx=5)
            
            thumb_label = tk.Label(row_frame, text="無圖片", bg="#e0e0e0", width=14, height=3)
            thumb_label.pack(side="left", padx=5)
            
            dur_str = self.format_duration(entry.get('duration'))
            title_text = entry.get('title', f'隱藏影片 {i+1}')
            txt_label = tk.Label(row_frame, text=f"{i+1}. {title_text}\n時間: {dur_str.strip() if dur_str else '未知'}", justify="left", wraplength=450, font=("Arial", 10))
            txt_label.pack(side="left", anchor="w")
            
            if HAS_PIL:
                url = entry.get('thumbnail')
                if not url and entry.get('thumbnails'):
                    url = entry['thumbnails'][0].get('url')
                if url:
                    threading.Thread(target=self.load_thumbnail, args=(url, thumb_label), daemon=True).start()

        self.select_btn_frame.pack(pady=3)
        self.list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.update_progress_ui(0, "解析完成！請勾選想下載的集數，點擊「開始下載」", "green")
        self.download_btn.config(state="normal")

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
            
            self.root.after(0, lambda: self.update_progress_ui(percent_val, f"下載進度: {percent_str} (速度: {speed}, 剩餘: {eta})", "blue"))
            
        elif d['status'] == 'finished':
            self.root.after(0, lambda: self.update_progress_ui(100.0, "單檔下載完成！正在合併影像或轉檔... (此階段無法暫停)", "orange"))

    def start_download(self):
        save_dir = self.download_path.get()
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)
            
        fmt = self.format_choice.get()
        quality = self.quality_combo.get()
        
        urls_to_download = []
        if self.is_playlist:
            if not self.playlist_vars:
                # 總結模式：全部下載
                selected_indices = list(range(len(self.playlist_entries)))
            else:
                selected_indices = [i for i, var in enumerate(self.playlist_vars) if var.get()]
                
            if not selected_indices:
                messagebox.showwarning("提示", "請至少在清單中勾選一部影片！")
                return
            entries = self.playlist_entries
            for i in selected_indices:
                vid_url = entries[i].get('url') or entries[i].get('webpage_url')
                if vid_url:
                    urls_to_download.append(vid_url)
                else:
                    vid_id = entries[i].get('id')
                    if vid_id:
                        urls_to_download.append(f"https://www.youtube.com/watch?v={vid_id}")
        else:
            urls_to_download.append(self.url_entry.get().strip())

        self.download_btn.config(state="disabled")
        self.analyze_btn.config(state="disabled")
        
        # 啟用控制按鈕並重置狀態
        self.is_cancelled = False
        self.is_paused = False
        self.pause_btn.config(state="normal", text="暫停", bg="SystemButtonFace")
        self.cancel_btn.config(state="normal")
        self.update_progress_ui(0, "準備開始下載...", "blue")
        
        threading.Thread(target=self.process_download, args=(urls_to_download, save_dir, fmt, quality), daemon=True).start()

    def process_download(self, urls, save_dir, fmt, quality):
        if fmt == "mp4":
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
                'outtmpl': os.path.join(save_dir, '%(title)s.%(ext)s'),
                'format': format_str,
                'merge_output_format': 'mp4',
                'progress_hooks': [self.progress_hook],
                'ffmpeg_location': self.app_dir,
                'color': 'no_color'
            }
        else:
            if "320" in quality:
                kbps = '320'
            elif "192" in quality:
                kbps = '192'
            else:
                kbps = '128'
                
            ydl_opts = {
                'outtmpl': os.path.join(save_dir, '%(title)s.%(ext)s'),
                'format': 'bestaudio/best',
                'postprocessors': [{
                    'key': 'FFmpegExtractAudio',
                    'preferredcodec': 'mp3',
                    'preferredquality': kbps,
                }],
                'progress_hooks': [self.progress_hook],
                'ffmpeg_location': self.app_dir,
                'color': 'no_color'
            }
            
        ydl_opts['noplaylist'] = True

        try:
            total = len(urls)
            for i, url in enumerate(urls):
                if self.is_cancelled:
                    break
                    
                if total > 1:
                    self.root.after(0, lambda idx=i: self.update_progress_ui(0, f"即將下載清單第 {idx+1}/{total} 部，請稍候...", "blue"))
                else:
                    self.root.after(0, lambda: self.update_progress_ui(0, "連線中，準備開始下載...", "blue"))
                    
                with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                    ydl.download([url])
                    
            if self.is_cancelled:
                self.root.after(0, lambda: self.update_progress_ui(0, "下載任務已取消", "red"))
                self.root.after(0, lambda: messagebox.showinfo("取消", "已成功取消下載任務。\n(未完成的暫存檔已保留，未來重新下載可自動接續進度)"))
            else:
                self.root.after(0, lambda: self.update_progress_ui(100.0, "所有任務皆已處理完成！", "green"))
                self.root.after(0, lambda: messagebox.showinfo("成功", f"全部下載完畢！\n檔案已成功儲存至：\n{save_dir}"))
                
        except Exception as e:
            # 判斷是否為我們主動拋出的取消例外
            if "USER_CANCELLED" in str(e):
                self.root.after(0, lambda: self.update_progress_ui(0, "下載任務已取消", "red"))
                self.root.after(0, lambda: messagebox.showinfo("取消", "已成功取消下載任務。\n(未完成的暫存檔已保留，未來重新下載可自動接續進度)"))
            else:
                self.root.after(0, lambda: self.update_progress_ui(0, "下載過程發生錯誤", "red"))
                self.root.after(0, lambda: messagebox.showerror("錯誤", f"下載失敗，可能是網路問題或影片遭版權封鎖：\n{str(e)}"))
        finally:
            self.root.after(0, lambda: self.download_btn.config(state="normal"))
            self.root.after(0, lambda: self.analyze_btn.config(state="normal"))
            self.root.after(0, lambda: self.pause_btn.config(state="disabled", text="暫停", bg="SystemButtonFace"))
            self.root.after(0, lambda: self.cancel_btn.config(state="disabled"))


# ===========================================================================
# MCIPlayer：使用 Windows 內建多媒體控制介面 (MCI) 播放 MP3，不需額外套件
# ===========================================================================
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

        tk.Label(left_frame, text="MP3 檔案列表", font=("Arial", 11, "bold")).pack(anchor="w")

        folder_frame = tk.Frame(left_frame)
        folder_frame.pack(fill="x", pady=3)
        self.folder_entry = tk.Entry(folder_frame, state="readonly", font=("Arial", 8))
        self.folder_entry.pack(side="left", fill="x", expand=True)
        tk.Button(folder_frame, text="選擇", command=self._browse_folder, font=("Arial", 8)).pack(side="left", padx=2)
        tk.Button(folder_frame, text="開啟", command=self._open_folder, font=("Arial", 8)).pack(side="left", padx=2)

        self.file_listbox = tk.Listbox(left_frame, font=("Arial", 9), selectmode="single", activestyle="dotbox")
        self.file_listbox.pack(fill="both", expand=True, pady=5)
        self.file_listbox.bind("<<ListboxSelect>>", self._on_file_select)

        tk.Button(left_frame, text="🔄 重新整理", command=self._refresh_list, font=("Arial", 9)).pack(fill="x")

        # === 右側：控制區 ===
        right_frame = tk.Frame(self.parent)
        right_frame.pack(side="left", fill="both", expand=True, padx=(5, 10), pady=10)

        # 檔名顯示
        self.file_label = tk.Label(right_frame, text="尚未選取檔案", font=("Arial", 10, "bold"),
                                   fg="#1565C0", wraplength=480, justify="left")
        self.file_label.pack(anchor="w", pady=(0, 5))

        # 播放控制按鈕（含跳轉±1s/±5s）
        ctrl_frame = tk.Frame(right_frame)
        ctrl_frame.pack(anchor="w", pady=3)
        tk.Button(ctrl_frame, text="⏮ -5s", command=lambda: self._seek_relative(-5000),
                  font=("Arial", 9), width=5).pack(side="left", padx=1)
        tk.Button(ctrl_frame, text="◀ -1s", command=lambda: self._seek_relative(-1000),
                  font=("Arial", 9), width=5).pack(side="left", padx=1)
        self.play_btn = tk.Button(ctrl_frame, text="▶ 播放", command=self._toggle_play,
                                  font=("Arial", 11, "bold"), bg="#4CAF50", fg="white", width=8, state="disabled")
        self.play_btn.pack(side="left", padx=4)
        self.stop_btn = tk.Button(ctrl_frame, text="⏹ 停止", command=self._stop,
                                  font=("Arial", 11), width=7, state="disabled")
        self.stop_btn.pack(side="left", padx=2)
        tk.Button(ctrl_frame, text="+1s ▶", command=lambda: self._seek_relative(1000),
                  font=("Arial", 9), width=5).pack(side="left", padx=1)
        tk.Button(ctrl_frame, text="+5s ⏭", command=lambda: self._seek_relative(5000),
                  font=("Arial", 9), width=5).pack(side="left", padx=1)
        self.time_label = tk.Label(ctrl_frame, text="00:00 / 00:00", font=("Arial", 11), fg="#333")
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
            tk.Label(legend_frame, text=label, font=("Arial", 8), fg="#555").pack(side="left", padx=(0, 8))

        ttk.Separator(right_frame, orient="horizontal").pack(fill="x", pady=5)

        # === 裁剪設定 ===
        trim_lf = tk.LabelFrame(right_frame, text="✂️ 裁剪設定", font=("Arial", 10, "bold"), padx=10, pady=6)
        trim_lf.pack(fill="x", pady=3)

        # 試聽與循環播放
        preview_row = tk.Frame(trim_lf)
        preview_row.pack(fill="x", pady=(2, 5))
        self.preview_btn = tk.Button(preview_row, text="▶ 試聽標記區段", command=self._preview_section,
                                     font=("Arial", 10, "bold"), bg="#7B1FA2", fg="white", state="disabled", width=16)
        self.preview_btn.pack(side="left", padx=(0, 12))
        tk.Checkbutton(preview_row, text="🔁 循環播放", variable=self._loop_var,
                       font=("Arial", 10), fg="#4A148C").pack(side="left")
        self.duration_label = tk.Label(preview_row, text="預計長度：0秒", font=("Arial", 10, "bold"), fg="#E91E63")
        self.duration_label.pack(side="left", padx=15)
        ttk.Separator(trim_lf, orient="horizontal").pack(fill="x", pady=(0, 4))

        # 起點
        start_row = tk.Frame(trim_lf)
        start_row.pack(fill="x", pady=3)
        tk.Label(start_row, text="起點：", font=("Arial", 11), width=5).pack(side="left")
        self.start_entry = tk.Entry(start_row, textvariable=self.start_time_str, width=10, font=("Arial", 11))
        self.start_entry.pack(side="left", padx=4)
        tk.Button(start_row, text="◀ -0.1s", command=lambda: self._adjust('start', -0.1),
                  font=("Arial", 9), width=6).pack(side="left", padx=2)
        tk.Button(start_row, text="+0.1s ▶", command=lambda: self._adjust('start', +0.1),
                  font=("Arial", 9), width=6).pack(side="left", padx=2)
        tk.Button(start_row, text="📍 標記目前位置", command=self._mark_start,
                  font=("Arial", 10), bg="#1976D2", fg="white").pack(side="left", padx=8)

        # 終點
        end_row = tk.Frame(trim_lf)
        end_row.pack(fill="x", pady=3)
        tk.Label(end_row, text="終點：", font=("Arial", 11), width=5).pack(side="left")
        self.end_entry = tk.Entry(end_row, textvariable=self.end_time_str, width=10, font=("Arial", 11))
        self.end_entry.pack(side="left", padx=4)
        tk.Button(end_row, text="◀ -0.1s", command=lambda: self._adjust('end', -0.1),
                  font=("Arial", 9), width=6).pack(side="left", padx=2)
        tk.Button(end_row, text="+0.1s ▶", command=lambda: self._adjust('end', +0.1),
                  font=("Arial", 9), width=6).pack(side="left", padx=2)
        tk.Button(end_row, text="📍 標記目前位置", command=self._mark_end,
                  font=("Arial", 10), bg="#E64A19", fg="white").pack(side="left", padx=8)

        ttk.Separator(right_frame, orient="horizontal").pack(fill="x", pady=8)

        # 輸出設定
        out_frame = tk.Frame(right_frame)
        out_frame.pack(fill="x", pady=3)
        tk.Label(out_frame, text="新檔名：", font=("Arial", 11)).pack(side="left")
        self.out_entry = tk.Entry(out_frame, font=("Arial", 10), width=35)
        self.out_entry.pack(side="left", padx=5, fill="x", expand=True)
        tk.Label(out_frame, text=".mp3", font=("Arial", 11)).pack(side="left")

        # 裁剪按鈕
        trim_btn_frame = tk.Frame(right_frame)
        trim_btn_frame.pack(pady=10)
        self.trim_btn = tk.Button(trim_btn_frame, text="✂️  裁剪並儲存新檔案", command=self._do_trim,
                                  font=("Arial", 13, "bold"), bg="#E53935", fg="white",
                                  width=25, height=2, state="disabled")
        self.trim_btn.pack()

        self.trim_status = tk.Label(right_frame, text="", font=("Arial", 10), fg="green")
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
        folder = getattr(self, '_folder_path', self.download_path_var.get())
        if os.path.exists(folder):
            os.startfile(folder)
        else:
            messagebox.showerror("錯誤", "找不到指定的資料夾路徑。")

    def _update_folder_entry(self, path):
        self.folder_entry.config(state="normal")
        self.folder_entry.delete(0, tk.END)
        self.folder_entry.insert(0, path)
        self.folder_entry.config(state="readonly")

    def _refresh_list(self):
        folder = getattr(self, '_folder_path', self.download_path_var.get())
        self._folder_path = folder
        self._update_folder_entry(folder)
        self.file_listbox.delete(0, tk.END)
        if not os.path.exists(folder):
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
        else:
            self.play_btn.config(text="▶ 播放")
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
        if self.player.get_mode() != "playing":
            self.player.play()
        self.play_btn.config(text="⏸ 暫停")
        self._start_update_loop()

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
        tk.Label(folder_frame, text="📁 歌曲資料夾：", font=("Arial", 10)).pack(side="left")
        self.path_entry = tk.Entry(folder_frame, textvariable=self.download_path_var, font=("Arial", 10))
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
        tk.Label(left_frame, text="1. 來源 MP3 (可多選)", font=("Arial", 10, "bold")).pack(anchor="w")
        tk.Label(left_frame, text="💡 按住 Shift 點選前後可連續選取", font=("Arial", 8), fg="#666").pack(anchor="w")
        
        # 先 pack 下方按鈕，確保不會被清單擠掉
        tk.Button(left_frame, text="🔄 重新整理", command=self._refresh_src_list, font=("Arial", 9)).pack(side="bottom", fill="x", pady=2)
        tk.Button(left_frame, text="➕ 加入合併清單", command=self._add_to_merge, bg="#4CAF50", fg="white", font=("Arial", 10, "bold")).pack(side="bottom", fill="x")
        
        # 支援自訂選取行為：單擊切換 + Shift 範圍選取
        self.src_listbox = tk.Listbox(left_frame, font=("Arial", 9), selectmode="extended")
        self.src_listbox.pack(fill="both", expand=True, pady=5)
        self.src_listbox.bind("<Button-1>", self._on_listbox_click)
        self._last_idx = None

        # 2. 中間：待合併清單 (Staging Area)
        mid_frame = tk.Frame(main_body, width=260)
        mid_frame.pack(side="left", fill="y", padx=5, pady=10)
        mid_frame.pack_propagate(False)
        tk.Label(mid_frame, text="2. 合併清單", font=("Arial", 10, "bold")).pack(anchor="w")
        
        # 先 pack 下方按鈕
        tk.Button(mid_frame, text="🧹 清除全部歌曲", command=self._clear_all, bg="#757575", fg="white").pack(side="bottom", fill="x", pady=2)
        tk.Button(mid_frame, text="🗑️ 移除選定歌曲", command=self._remove_from_merge, bg="#f44336", fg="white").pack(side="bottom", fill="x", pady=2)
        
        btn_grid = tk.Frame(mid_frame)
        btn_grid.pack(side="bottom", fill="x")
        tk.Button(btn_grid, text="🔼 上移", command=lambda: self._move_item(-1), width=10).pack(side="left", padx=2, expand=True, fill="x")
        tk.Button(btn_grid, text="🔽 下移", command=lambda: self._move_item(1), width=10).pack(side="left", padx=2, expand=True, fill="x")

        self.merge_listbox = tk.Listbox(mid_frame, font=("Arial", 9))
        self.merge_listbox.pack(fill="both", expand=True, pady=5)

        # 3. 右側：預覽與執行
        right_frame = tk.Frame(self.parent)
        right_frame.pack(side="left", fill="both", expand=True, padx=(5, 10), pady=10)
        tk.Label(right_frame, text="3. 預覽與輸出", font=("Arial", 10, "bold")).pack(anchor="w")

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

        self.time_label = tk.Label(right_frame, text="00:00 / 00:00", font=("Arial", 10))
        self.time_label.pack(pady=2)

        # 虛擬進度條 (Canvas)
        canvas_outer = tk.Frame(right_frame, bg="#888", pady=1)
        canvas_outer.pack(fill="x", pady=5)
        self.merge_canvas = tk.Canvas(canvas_outer, height=40, bg="#eee", highlightthickness=0, cursor="hand2")
        self.merge_canvas.pack(fill="both", expand=True)
        self.merge_canvas.bind("<ButtonPress-1>", self._canvas_click)
        self.merge_canvas.bind("<Configure>", lambda e: self._draw_canvas())

        # 輸出設定
        tk.Label(right_frame, text="合併後檔名：").pack(anchor="w", pady=(10, 0))
        self.out_entry = tk.Entry(right_frame)
        self.out_entry.pack(fill="x", pady=5)
        self.out_entry.insert(0, "merged_audio")

        # 融合效果設定 (Crossfade)
        fade_frame = tk.Frame(right_frame)
        fade_frame.pack(fill="x", pady=5)
        tk.Checkbutton(fade_frame, text="✨ 啟用融合效果 (Crossfade)", variable=self.fade_var, 
                       font=("Arial", 10, "bold"), fg="#1976D2").pack(side="left")
        tk.Label(fade_frame, text="  融合秒數：").pack(side="left")
        tk.Spinbox(fade_frame, from_=1, to=5, textvariable=self.fade_sec, width=5).pack(side="left")
        tk.Label(fade_frame, text="(註：融合需重新轉檔，速度較慢)", font=("Arial", 8), fg="gray").pack(side="left", padx=5)

        self.merge_btn = tk.Button(right_frame, text="🚀 開始合併所有歌曲", command=self._do_merge, 
                                   bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), height=2)
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
                c.create_text(mid, h // 2, text="↔", font=("Arial", 8), fill="white", anchor="center")

            # 標示起始時間
            if x1 - x0 > 45:
                t_str = self._fmt_ms(int(acc_virtual))
                c.create_text(x0 + 3, h - 4, text=t_str, anchor="sw", font=("Arial", 8), fill="#222")

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
        out_path = os.path.join(self.download_path_var.get(), out_name + ".mp3")
        
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

