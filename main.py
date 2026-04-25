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

ssl._create_default_https_context = ssl._create_unverified_context
APP_VERSION = "1.2.8"
GITHUB_REPO = "mathced-com/CYT_YTDL"

try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# 將工作目錄設定為程式所在資料夾 (如果是打包環境則設定為 exe 所在目錄)
if getattr(sys, 'frozen', False):
    os.chdir(os.path.dirname(sys.executable))
else:
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

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
        self.root.geometry("750x650")
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

    def create_widgets(self):
        header_frame = tk.Frame(self.root)
        header_frame.pack(pady=10)
        
        try:
            logo_img = Image.open(self.resource_path("icon.ico")).resize((32, 32), Image.Resampling.LANCZOS)
            self.logo_photo = ImageTk.PhotoImage(logo_img)
            logo_label = tk.Label(header_frame, image=self.logo_photo)
            logo_label.pack(side="left", padx=10)
        except Exception:
            pass
            
        title_label = tk.Label(header_frame, text=f"CYT_YouTube 下載器 v{APP_VERSION}", font=("Arial", 16, "bold"))
        title_label.pack(side="left")
        
        url_frame = tk.Frame(self.root)
        url_frame.pack(fill="x", padx=20, pady=5)
        tk.Label(url_frame, text="網址：", font=("Arial", 12)).pack(side="left")
        
        self.url_entry = tk.Entry(url_frame, width=35, font=("Arial", 10))
        self.url_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        self.analyze_btn = tk.Button(url_frame, text="解析網址", command=self.start_analyze, bg="#2196F3", fg="white", font=("Arial", 10, "bold"))
        self.analyze_btn.pack(side="left", padx=2)
        
        self.clear_btn = tk.Button(url_frame, text="清除網址", command=self.clear_url, font=("Arial", 10))
        self.clear_btn.pack(side="left", padx=2)
        
        # 步驟提示
        hint_text = "執行步驟：\n一、貼上Youtube網址\n二、點擊「解析網址」\n三、點擊「開始下載」"
        hint_label = tk.Label(url_frame, text=hint_text, fg="#E91E63", font=("Arial", 9, "bold"), justify="left")
        hint_label.pack(side="left", padx=5)
        
        # 先建立底部框架並鎖定在視窗最下方，保證不被清單擠出畫面
        bottom_frame = tk.Frame(self.root)
        bottom_frame.pack(side="bottom", fill="x", pady=5)
        
        self.info_frame = tk.LabelFrame(self.root, text="影片預覽 / 播放清單", font=("Arial", 10))
        self.info_frame.pack(fill="both", expand=True, padx=20, pady=5)
        
        self.title_label = tk.Label(self.info_frame, text="請輸入網址並點選「解析網址」", fg="gray", wraplength=650, justify="left")
        self.title_label.pack(pady=5, padx=10)
        
        self.list_frame = ScrollableFrame(self.info_frame)
        
        self.select_btn_frame = tk.Frame(self.info_frame)
        tk.Button(self.select_btn_frame, text="全部勾選", command=self.select_all).pack(side="left", padx=5)
        tk.Button(self.select_btn_frame, text="取消全選", command=self.deselect_all).pack(side="left", padx=5)
        
        format_frame = tk.Frame(bottom_frame)
        format_frame.pack(fill="x", padx=20, pady=5)
        tk.Label(format_frame, text="格式：", font=("Arial", 12)).pack(side="left")
        tk.Radiobutton(format_frame, text="MP4", variable=self.format_choice, value="mp4", command=self.update_quality_options).pack(side="left", padx=2)
        tk.Radiobutton(format_frame, text="MP3", variable=self.format_choice, value="mp3", command=self.update_quality_options).pack(side="left", padx=2)
        
        tk.Label(format_frame, text="   品質：", font=("Arial", 12)).pack(side="left")
        self.quality_combo = ttk.Combobox(format_frame, textvariable=self.quality_choice, state="readonly", width=18)
        self.quality_combo.pack(side="left", padx=5)
        
        path_frame = tk.Frame(bottom_frame)
        path_frame.pack(fill="x", padx=20, pady=5)
        tk.Label(path_frame, text="儲存：", font=("Arial", 12)).pack(side="left")
        self.path_entry = tk.Entry(path_frame, textvariable=self.download_path, width=40, state="readonly", font=("Arial", 10))
        self.path_entry.pack(side="left", padx=5, fill="x", expand=True)
        tk.Button(path_frame, text="選擇資料夾", command=self.browse_folder).pack(side="left")
        
        status_frame = tk.Frame(bottom_frame)
        status_frame.pack(fill="x", padx=20, pady=5)
        self.progress_bar = ttk.Progressbar(status_frame, orient="horizontal", length=700, mode="determinate")
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
        if messagebox.askyesno("發現新版本", f"發現新版本 v{latest_version}！\n是否要立即下載並更新？\n\n注意：更新時程式將會自動關閉並重啟。"):
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
                    if messagebox.askyesno("更新準備就緒", "新版本已下載完畢！\n\n程式將立刻重新啟動以套用更新。\n\n請問是否立即重啟？"):
                        if getattr(sys, 'frozen', False):
                            current_exe_path = sys.executable
                            old_exe_path = current_exe_path + ".old"
                            
                            try:
                                os.rename(current_exe_path, old_exe_path)
                                os.rename(new_exe_path, current_exe_path)
                                
                                # 使用 os.startfile 完全等同於使用者親手雙擊檔案，
                                # 它會透過 Windows Shell 啟動，徹底避免防毒軟體因為父子程序啟動而產生的攔截，
                                # 同時也不會繼承到任何舊版的暫存工作目錄或污染的環境變數。
                                # 為了防止環境變數汙染，手動清除當前進程的 PyInstaller 變數，確保新版程式能獨立解壓
                                for key in ['_MEIPASS2', '_MEIPASS', 'TCL_LIBRARY', 'TK_LIBRARY', 'PYTZ_TZDATADIR']:
                                    os.environ.pop(key, None)
                                
                                try:
                                    os.startfile(current_exe_path)
                                except AttributeError:
                                    subprocess.Popen([current_exe_path])
                                    
                                os._exit(0)
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

        self.list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.select_btn_frame.pack(pady=5)
        
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

if __name__ == "__main__":
    root = tk.Tk()
    app = YouTubeDownloaderGUI(root)
    root.mainloop()
