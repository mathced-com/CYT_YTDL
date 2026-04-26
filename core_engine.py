import yt_dlp
import threading
import os

class YTDownloaderEngine:
    def __init__(self):
        self.video_info = None

    def format_duration(self, seconds):
        if not seconds:
            return "00:00"
        try:
            seconds = int(seconds)
            m, s = divmod(seconds, 60)
            h, m = divmod(m, 60)
            return f"{h}:{m:02d}:{s:02d}" if h else f"{m:02d}:{s:02d}"
        except:
            return "00:00"

    def analyze_url(self, url, callback):
        """
        異步解析網址。
        """
        def task():
            # 針對播放清單的解析優化
            ydl_opts = {
                'extract_flat': True, # 先快速抓取列表
                'quiet': True,
                'no_warnings': True,
            }
            try:
                with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                    info = ydl.extract_info(url, download=False)
                
                if not info:
                    callback(False, "無法獲取影片資訊")
                    return

                is_playlist = 'entries' in info
                entries = []
                if is_playlist:
                    raw_entries = list(info.get('entries', []))
                    # 最多取前 50 筆
                    for i, entry in enumerate(raw_entries[:50]):
                        if entry:
                            entries.append({
                                "index": i,
                                "title": entry.get('title', f"影片 {i+1}"),
                                "url": entry.get('url') or entry.get('webpage_url'),
                                "duration": self.format_duration(entry.get('duration')),
                                "id": entry.get('id')
                            })

                result = {
                    "title": info.get('title', '未知標題'),
                    "thumbnail": info.get('thumbnail'),
                    "duration": self.format_duration(info.get('duration')),
                    "is_playlist": is_playlist,
                    "entries": entries,
                    "original_url": url
                }
                callback(True, result)
            except Exception as e:
                callback(False, str(e))

        threading.Thread(target=task, daemon=True).start()

    def download_video(self, tasks, options, progress_callback):
        """
        執行下載任務。
        tasks: 列表，包含要下載的網址或 entry 對象
        options: {"format": "mp3"/"mp4", "quality": "320"/"1080", "path": "save_path"}
        """
        def progress_hook(d):
            if d['status'] == 'downloading':
                p = d.get('_percent_str', '0%').replace('%', '')
                try:
                    percent = float(p) / 100
                except:
                    percent = 0
                speed = d.get('_speed_str', '未知')
                filename = os.path.basename(d.get('filename', '檔案'))
                progress_callback("downloading", {"percent": percent, "speed": speed, "filename": filename})
            elif d['status'] == 'finished':
                progress_callback("processing", None)

        def task():
            save_path = options.get("path", ".")
            fmt = options.get("format", "mp3")
            quality = options.get("quality", "192")
            
            for i, video_url in enumerate(tasks):
                progress_callback("start_item", {"index": i, "total": len(tasks)})
                
                ydl_opts = {
                    'progress_hooks': [progress_hook],
                    'outtmpl': f'{save_path}/%(title)s.%(ext)s',
                    'quiet': True,
                    'noplaylist': True, # 強制下載單一影片
                }

                if fmt == "mp3":
                    ydl_opts.update({
                        'format': 'bestaudio/best',
                        'postprocessors': [{
                            'key': 'FFmpegExtractAudio',
                            'preferredcodec': 'mp3',
                            'preferredquality': quality,
                        }],
                    })
                else:
                    # 解析度限制
                    res = "1080" if quality == "1080" else "720"
                    ydl_opts.update({
                        'format': f'bestvideo[height<={res}][ext=mp4]+bestaudio[ext=m4a]/best[height<={res}][ext=mp4]/best',
                    })

                try:
                    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                        ydl.download([video_url])
                except Exception as e:
                    progress_callback("error_item", str(e))
                    continue

            progress_callback("success_all", None)

        threading.Thread(target=task, daemon=True).start()

    def trim_audio(self, in_path, out_path, start_sec, end_sec, callback):
        """
        音訊裁剪邏輯。
        """
        def task():
            try:
                import subprocess
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
                    callback(True, os.path.basename(out_path))
                else:
                    callback(False, result.stderr[-200:])
            except Exception as e:
                callback(False, str(e))

        threading.Thread(target=task, daemon=True).start()

    def merge_audios(self, files, out_path, fade_sec, callback):
        """
        音訊合併邏輯。支援 Crossfade。
        """
        def task():
            try:
                import subprocess
                if fade_sec <= 0:
                    # 高速合併
                    list_file = "concat_list.txt"
                    with open(list_file, "w", encoding="utf-8") as f:
                        for fp in files:
                            # 轉義路徑中的單引號
                            p = fp.replace("'", "'\\''")
                            f.write(f"file '{p}'\n")
                    cmd = ["ffmpeg", "-y", "-f", "concat", "-safe", "0", "-i", list_file, "-c", "copy", out_path]
                    res = subprocess.run(cmd, capture_output=True, text=True)
                    if os.path.exists(list_file): os.remove(list_file)
                else:
                    # 融合合併 (acrossfade)
                    cmd = ["ffmpeg", "-y"]
                    for f in files:
                        cmd.extend(["-i", f])
                    
                    filter_str = ""
                    last_label = "[0]"
                    for i in range(1, len(files)):
                        next_label = f"[a{i}]"
                        filter_str += f"{last_label}[{i}]acrossfade=d={fade_sec}:c1=tri:c2=tri"
                        if i < len(files) - 1:
                            filter_str += f"{next_label};"
                            last_label = next_label
                    
                    cmd.extend(["-filter_complex", filter_str, "-b:a", "320k", out_path])
                    res = subprocess.run(cmd, capture_output=True, text=True)

                if res.returncode == 0:
                    callback(True, os.path.basename(out_path))
                else:
                    callback(False, "FFmpeg 合併失敗")
            except Exception as e:
                callback(False, str(e))

        threading.Thread(target=task, daemon=True).start()

engine = YTDownloaderEngine()
