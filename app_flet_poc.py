import flet as ft
import os
import asyncio
from core_engine import engine

async def main(page: ft.Page):
    page.title = "CYT YouTube 下載器 v3.0"
    page.theme_mode = ft.ThemeMode.DARK
    page.window_width = 1100
    page.window_height = 950
    page.padding = 30
    page.theme = ft.Theme(font_family="Inter", visual_density=ft.VisualDensity.COMFORTABLE)

    # --- 全域狀態 ---
    base_dir = os.path.dirname(os.path.abspath(__file__))
    default_download_path = os.path.join(base_dir, "downloads")
    if not os.path.exists(default_download_path): os.makedirs(default_download_path)

    state = {
        "format": "mp3", "quality": "320", "save_path": default_download_path,
        "current_data": None, "selected_indices": set(),
        "trim_file": None, "trim_start": 0.0, "trim_end": 0.0,
        "merge_files": []
    }

    # --- 通用服務 (僅保留穩定的 FilePicker) ---
    directory_picker = ft.FilePicker()
    file_picker = ft.FilePicker()
    merge_picker = ft.FilePicker()
    
    # 重新掛載到 overlay，FilePicker 在 0.84.0 通常是穩定的
    page.overlay.append(directory_picker)
    page.overlay.append(file_picker)
    page.overlay.append(merge_picker)

    # --- 下載器邏輯 ---
    async def pick_folder(e):
        try:
            path = await directory_picker.get_directory_path()
            if path:
                state["save_path"] = path
                path_input.value = path
                page.update()
        except: pass

    url_input = ft.TextField(label="YouTube 網址", prefix_icon="link", border_radius=15, expand=True)
    path_input = ft.TextField(label="儲存路徑", value=state["save_path"], text_size=12, height=40, expand=True, read_only=True)
    video_title = ft.Text("尚未解析影片", size=20, weight=ft.FontWeight.BOLD, no_wrap=True, overflow="ellipsis")
    video_duration = ft.Text("", size=14, color=ft.Colors.CYAN_200)
    video_thumbnail = ft.Image(src="", width=160, height=120, fit="cover", border_radius=10, visible=False)
    status_text = ft.Text("請輸入網址並點擊解析", size=14, color=ft.Colors.GREY_400)

    quality_dropdown = ft.Dropdown(label="音質/畫質", width=150, value="320",
        options=[ft.dropdown.Option("128", "128kbps"), ft.dropdown.Option("192", "192kbps"), ft.dropdown.Option("320", "320kbps")])
    quality_dropdown.on_change = lambda e: state.update({"quality": e.control.value})

    async def on_format_change(e):
        state["format"] = e.control.value
        if e.control.value == "mp4":
            quality_dropdown.options = [ft.dropdown.Option("720", "720p"), ft.dropdown.Option("1080", "1080p")]
            quality_dropdown.value = "1080"
        else:
            quality_dropdown.options = [ft.dropdown.Option("128", "128"), ft.dropdown.Option("192", "192"), ft.dropdown.Option("320", "320")]
            quality_dropdown.value = "320"
        page.update()

    format_radio = ft.RadioGroup(value="mp3", on_change=on_format_change,
        content=ft.Row([ft.Radio(value="mp3", label="MP3"), ft.Radio(value="mp4", label="MP4")]))

    playlist_container = ft.Column(spacing=5, scroll=ft.ScrollMode.ALWAYS, height=250)
    playlist_card = ft.Container(visible=False, padding=15, bgcolor=ft.Colors.with_opacity(0.05, ft.Colors.WHITE), border_radius=10,
        content=ft.Column([ft.Row([ft.Text("播放清單", weight="bold"), ft.Row([ft.TextButton("全選", on_click=lambda _: select_all(True)), ft.TextButton("清空", on_click=lambda _: select_all(False))])], alignment="spaceBetween"), playlist_container]))

    def select_all(val):
        for c in playlist_container.controls:
            if isinstance(c, ft.Checkbox):
                c.value = val
                if val: state["selected_indices"].add(c.data)
                else: state["selected_indices"].discard(c.data)
        page.update()

    progress_bar = ft.ProgressBar(color=ft.Colors.CYAN_400, value=0, visible=False)
    progress_text = ft.Text("", size=12, color=ft.Colors.GREY_400)

    def on_dl_prog(status, data):
        if status == "downloading":
            progress_bar.value = data['percent']
            progress_text.value = f"下載中: {int(data['percent']*100)}% ({data['speed']})"
        elif status == "success_all":
            progress_bar.visible = False
            progress_text.value = "✅ 下載完成！"
            progress_text.color = ft.Colors.GREEN_400
            download_btn.disabled = False
        page.update()

    async def start_download(e):
        if not state["current_data"]: return
        urls = []
        if state["current_data"]["is_playlist"]:
            for idx in state["selected_indices"]: urls.append(state["current_data"]["entries"][idx]["url"])
        else: urls.append(state["current_data"]["original_url"])
        if not urls: return
        download_btn.disabled = True
        progress_bar.visible = True
        page.update()
        engine.download_video(urls, state, on_dl_prog)

    def on_analyze_res(success, data):
        analyze_btn.disabled = False
        analyze_btn.content = ft.Text("解析網址", weight="bold")
        if success:
            state["current_data"] = data
            video_title.value, video_duration.value = data['title'], f"時長: {data['duration']}"
            video_thumbnail.src = data['thumbnail'] or ""; video_thumbnail.visible = True
            if data['is_playlist']:
                playlist_container.controls.clear()
                for item in data['entries']:
                    playlist_container.controls.append(ft.Checkbox(label=f"{item['title']} ({item['duration']})", data=item['index'], on_change=lambda e: state["selected_indices"].add(e.control.data) if e.control.value else state["selected_indices"].discard(e.control.data)))
                playlist_card.visible = True
            else: playlist_card.visible = False
            download_options.visible = True
        page.update()

    async def handle_analyze(e):
        if not url_input.value: return
        analyze_btn.disabled = True
        analyze_btn.content = ft.ProgressRing(width=20, height=20, stroke_width=2)
        page.update()
        engine.analyze_url(url_input.value, on_analyze_res)

    analyze_btn = ft.ElevatedButton(content=ft.Text("解析網址", weight="bold"), on_click=handle_analyze)
    download_btn = ft.ElevatedButton(content=ft.Text("開始下載任務"), icon="download", on_click=start_download)
    download_options = ft.Container(visible=False, content=ft.Column([ft.Row([ft.Column([ft.Text("1. 格式"), format_radio]), ft.Column([ft.Text("2. 品質"), quality_dropdown]), ft.Column([ft.Text("3. 存檔"), ft.Row([path_input, ft.IconButton(icon=ft.Icons.FOLDER_OPEN, on_click=pick_folder, visible=not page.web)])])]), progress_bar, progress_text, download_btn]))

    download_view = ft.Column([ft.Row([url_input, analyze_btn]), ft.Container(padding=20, bgcolor=ft.Colors.with_opacity(0.1, ft.Colors.WHITE), border_radius=20, content=ft.Column([ft.Row([video_thumbnail, ft.Column([video_title, video_duration, status_text], expand=True)]), playlist_card, download_options]))])

    # --- 音訊裁剪邏輯 (移除 Audio 避免紅框) ---
    trim_file_label = ft.Text("請選擇 MP3 檔案", color=ft.Colors.GREY_500)
    trim_start_input = ft.TextField(label="起點 (秒)", value="0.0", width=150)
    trim_end_input = ft.TextField(label="終點 (秒)", value="10.0", width=150)
    trim_status = ft.Text("", size=12)

    async def on_trim_file_picked(e):
        if e.files:
            state["trim_file"] = e.files[0].path
            trim_file_label.value = f"已選擇: {e.files[0].name}"
            page.update()

    file_picker.on_result = on_trim_file_picked

    def do_trim(e):
        if not state["trim_file"]: return
        try:
            s, e_val = float(trim_start_input.value), float(trim_end_input.value)
            out_path = os.path.join(state["save_path"], "trimmed_" + os.path.basename(state["trim_file"]))
            trim_btn.disabled = True
            trim_status.value = "裁剪中..."
            page.update()
            engine.trim_audio(state["trim_file"], out_path, s, e_val, lambda success, res: 
                setattr(trim_status, "value", f"✅ 已存至: {res}" if success else f"❌ 錯誤: {res}") or 
                setattr(trim_btn, "disabled", False) or page.update())
        except ValueError:
            trim_status.value = "❌ 請輸入有效的數字"
            page.update()

    trim_btn = ft.ElevatedButton("執行裁剪", icon="content_cut", on_click=do_trim)
    
    cutter_view = ft.Column([
        ft.Row([ft.ElevatedButton("選擇檔案", icon="file_open", on_click=lambda _: file_picker.pick_files()), trim_file_label]),
        ft.Divider(),
        ft.Text("💡 提示：請輸入欲裁剪的時間點 (秒數)"),
        ft.Row([trim_start_input, trim_end_input]),
        ft.Divider(),
        trim_btn,
        trim_status
    ])

    # --- 音訊合併邏輯 ---
    merge_list = ft.Column()
    fade_input = ft.TextField(label="交叉淡入淡出 (秒)", value="3", width=150)
    merge_status = ft.Text("", size=12)

    async def on_merge_files_picked(e):
        if e.files:
            for f in e.files:
                state["merge_files"].append(f.path)
                merge_list.controls.append(ft.ListTile(title=ft.Text(f.name), leading=ft.Icon(ft.Icons.MUSIC_NOTE)))
            page.update()

    merge_picker.on_result = on_merge_files_picked

    def do_merge(e):
        if not state["merge_files"]: return
        out_path = os.path.join(state["save_path"], "merged_audio.mp3")
        merge_btn.disabled = True
        merge_status.value = "合併中 (加上融合效果可能需要較長時間)..."
        page.update()
        engine.merge_audios(state["merge_files"], out_path, float(fade_input.value), lambda success, res:
            setattr(merge_status, "value", f"✅ 合併成功: {res}" if success else f"❌ 失敗: {res}") or
            setattr(merge_btn, "disabled", False) or page.update())

    merge_btn = ft.ElevatedButton("開始合併", icon="call_merge", on_click=do_merge)

    merger_view = ft.Column([
        ft.ElevatedButton("選擇多個檔案", icon="library_music", on_click=lambda _: merge_picker.pick_files(allow_multiple=True)),
        ft.Container(content=merge_list, height=300, bgcolor=ft.Colors.with_opacity(0.05, ft.Colors.WHITE), border_radius=10, padding=10),
        ft.Row([fade_input, ft.Text("秒 (設為 0 則快速拼接)")]),
        merge_btn,
        merge_status
    ])

    # 確保 merge_list 本身支援滾動
    merge_list.scroll = ft.ScrollMode.ALWAYS

    # --- 最終組裝 ---
    tabs = ft.Tabs(selected_index=0, length=4, expand=True,
        content=ft.Column(expand=True, controls=[
            ft.TabBar(tabs=[ft.Tab("下載器", icon="download"), ft.Tab("音訊裁剪", icon="content_cut"), ft.Tab("音訊合併", icon="call_merge"), ft.Tab("設定", icon="settings")]),
            ft.TabBarView(expand=True, controls=[
                ft.Container(content=download_view, padding=20),
                ft.Container(content=cutter_view, padding=20),
                ft.Container(content=merger_view, padding=20),
                ft.Container(content=ft.Text("⚙️ 設定開發中"), alignment=ft.Alignment.CENTER),
            ])
        ]))

    page.add(ft.Column(expand=True, controls=[ft.Text("CYT YouTube 下載器 v3.0", size=32, weight="bold", color=ft.Colors.CYAN_400), tabs]))

if __name__ == "__main__":
    ft.app(target=main)
