# CYT_YTDL (CYT_YouTube 下載器) - 專案開發紀錄與 AI 備忘錄 (v2.4.1)

> **AI 讀取指示**：當使用者在新環境啟動對話並要求「接續開發」時，請優先閱讀此檔案與 `交接筆記.md` 以掌握專案的最新狀態與架構。

## 📌 專案目前狀態
- **最新版本**：`v2.4.1` (正式穩定版，已發布於 GitHub Releases)。
- **核心狀態**：四大功能分頁架構穩定，YouTube 下載核心已升級最新 yt-dlp 並相容 Android/Web 串流協定，可 100% 順暢下載與解析。
- **近期重要變動**：
  1. **YouTube 403 繞過防護**：於解析與下載參數中配置 `extractor_args={'youtube': {'player_client': ['android', 'web']}}`，解決 YouTube SABR 與 PO Token / 403 封鎖。
  2. **播放清單自訂筆數模式**：清單超過 50 筆時，提供「1. 前50筆」、「2. 全部」、「3. 前幾筆 (自訂輸入筆數)」、「4. 取消」四種選項。
  3. **下載結果彈窗與狀態精確化**：區分「全部成功」、「部分成功」、「全部失敗」的彈窗提示。
  4. **殘留暫存檔自動清除**：下載完成後自動清理目標目錄中同影片 ID 的 `.part` / `.ytdl` 歷史未完成暫存檔。
  5. **版本與說明書同步**：全面校正 `使用說明.md` 與 `使用說明.txt` 為 v2.4.1。

## 🏗️ 核心架構與技術棧
- **開發語言**：Python 3
- **GUI 介面**：`tkinter` + `ttk` (字型全域指定微軟正黑體 `Microsoft JhengHei`，解決 Windows Tcl 吃字 Bug)
- **核心下載引擎**：`yt-dlp` (內嵌於獨立 EXE 中，打包後使用者不需額外安裝)
- **多媒體處理**：`FFmpeg` (`ffmpeg.exe` 與 `ffprobe.exe`，負責影音合併、音訊轉檔、字幕硬燒錄與無損裁剪)
- **影像處理**：`Pillow` (PIL) (負責在介面中縮放與顯示影片縮圖)
- **本機播放引擎**：Windows 原生 `winmm.dll` (MCIPlayer) + `GetShortPathNameW` (8.3 短檔名轉換，相容中文與空格路徑)
- **非同步處理**：`threading` 與 `concurrent.futures.ThreadPoolExecutor` (支援自訂 1~5 同時下載數，介面絕不卡死)

## 📁 主要檔案結構說明
- **`main.py`**：核心程式碼。包含主控制器與四大分頁類別 (`YouTubeDownloaderGUI`, `AudioEditorTab`, `VideoEditorTab`, `VideoConverterTab`)。
- **`release_helper.py`** / **`一鍵發布新版本.bat`**：自動化編譯打包與 GitHub Release 發布工具。
- **`使用說明.md`** / **`使用說明.txt`**：繁體中文使用者操作指南。
- **`config.json`**：使用者下載路徑、預設格式、並行數等設定持久化檔案。
- **`Project_Log.md`** / **`交接筆記.md`** / **`優化方向.md`**：專案開發紀錄與交接技術文檔。

## 🚀 四大核心分頁功能
1. **⬇️ 影音下載器**：單片/清單解析、1~5 線程並行下載、MP4/MKV 畫質選擇、MP3/WAV 音質選擇、章節分割、Cookies 載入。
2. **🎵 音樂編輯工具**：MP3 無損裁剪 (MCI 即時播放預覽) + MP3 串接合併與 Crossfade 淡入淡出。
3. **🎬 影片編輯工具**：MP4/MKV 無損秒級裁剪 + 影片合併 + 影音分離 (一鍵提取 320k MP3/WAV)。
4. **🔄 影音轉檔器**：格式轉換、畫質降規壓縮、外掛 .srt 字幕硬燒錄、DVD 連續 VOB 無縫拼接合併。

## 📝 下一步計畫 (TODO)
- 待後續版本規劃：整合 `aria2c` 多線程加速、GPU 硬體加速轉檔 (NVENC/QSV)、軟字幕封裝 (Softcode Subtitles) 等。
