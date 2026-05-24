import os
import re
import sys
import subprocess
import webbrowser

def get_next_version(current):
    parts = current.split('.')
    if len(parts) == 3 and parts[-1].isdigit():
        parts[-1] = str(int(parts[-1]) + 1)
        return '.'.join(parts)
    return current + "_new"

def check_gh_login():
    try:
        result = subprocess.run(["gh", "auth", "status"], capture_output=True, text=True, encoding='utf-8', errors='ignore')
        if result.returncode != 0:
            print("\n[!] 偵測到您尚未登入 GitHub 命令列工具。")
            print("為了能自動上傳檔案，現在將啟動一次性登入流程：")
            print("請在接下來的提示中選擇：")
            print("1. 選擇 GitHub.com")
            print("2. 選擇 HTTPS")
            print("3. 選擇 Y (Authenticate Git with your GitHub credentials)")
            print("4. 選擇 Login with a web browser")
            print("5. 複製一次性驗證碼並在瀏覽器貼上\n")
            subprocess.run(["gh", "auth", "login"])
            
            result2 = subprocess.run(["gh", "auth", "status"], capture_output=True, text=True, encoding='utf-8', errors='ignore')
            if result2.returncode != 0:
                print("\n登入失敗或取消，無法自動上傳 Release。")
                return False
        return True
    except FileNotFoundError:
        print("\n[!] 系統中找不到 gh 指令 (GitHub CLI)。請先安裝 GitHub CLI 才能完全自動化。")
        return False

def convert_md_to_txt(md_path, txt_path):
    try:
        with open(md_path, "r", encoding="utf-8") as f:
            content = f.read()
        
        # 簡單的去標記化：去除 GitHub 警示框語法
        content = re.sub(r'> \[!IMPORTANT\]', '【重要聲明】', content)
        content = re.sub(r'> \[!NOTE\]', '【備註】', content)
        content = re.sub(r'> \[!TIP\]', '【提示】', content)
        content = re.sub(r'# ', '', content)
        content = re.sub(r'## ', '■ ', content)
        content = re.sub(r'### ', '  - ', content)
        content = re.sub(r'\*\*', '', content)
        
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(content)
        return True
    except Exception as e:
        print(f"轉換說明檔失敗: {e}")
        return False

def main():
    print("==============================================")
    print("      CYT_YTDL 一鍵發布新版本助手 (全自動版)")
    print("==============================================\n")
    
    current_version = "未知"
    content = ""
    try:
        with open("main.py", "r", encoding="utf-8") as f:
            content = f.read()
            match = re.search(r'APP_VERSION\s*=\s*"([^"]+)"', content)
            if match:
                current_version = match.group(1)
    except Exception as e:
        print(f"讀取 main.py 失敗: {e}")
        input("請按 Enter 鍵結束...")
        return

    suggested_version = get_next_version(current_version)
    print(f"目前專案版本為: {current_version}")
    
    print("\n==============================================")
    print("      請選擇本次執行模式 (Action Mode)")
    print("==============================================")
    print("  [1] 🚀 發布「正式穩定版」 (自動更新版，所有使用者下載此最新版)")
    print("  [2] 🧪 發布「測試預覽版」 (Pre-release，GitHub 標記為測試版)")
    print("  [3] 💾 僅「上傳備份程式碼」 (僅同步目前代碼至 GitHub，不打包、不發布)")
    print("==============================================")
    
    mode_choice = input("請輸入選項 [預設為 1]: ").strip()
    if not mode_choice:
        mode_choice = "1"

    # ==========================================
    # 模式 3：僅備份程式碼
    # ==========================================
    if mode_choice == "3":
        print(f"\n[1/2] 正在準備備份程式碼 (目前版本保持為: {current_version})...")
        update_notes = input("請簡單輸入本次備份的修改說明 (直接 Enter 預設為 '備份與同步程式碼'): ").strip()
        if not update_notes:
            update_notes = "備份與同步程式碼"
            
        change_ver = input(f"是否需要順便變更程式版本號？(y/N) [目前為 {current_version}]: ").strip().lower()
        if change_ver == 'y':
            new_version = input(f"請輸入新版本號 [直接按 Enter 預設為 {suggested_version}]: ").strip()
            if not new_version:
                new_version = suggested_version
            
            print(f"➔ 正在將 main.py 與 使用說明.md 內部的版本號更新為: {new_version}...")
            try:
                new_content = re.sub(r'APP_VERSION\s*=\s*"[^"]+"', f'APP_VERSION = "{new_version}"', content)
                with open("main.py", "w", encoding="utf-8") as f:
                    f.write(new_content)
                if os.path.exists("使用說明.md"):
                    with open("使用說明.md", "r", encoding="utf-8") as f:
                        md_content = f.read()
                    md_content = re.sub(r'使用說明 \(v[^)]+\)', f'使用說明 (v{new_version})', md_content)
                    md_content = re.sub(r'Version [0-9.]+', f'Version {new_version}', md_content)
                    with open("使用說明.md", "w", encoding="utf-8") as f:
                        f.write(md_content)
                print("      [OK] 版本號已更新。")
                commit_version = new_version
            except Exception as e:
                print(f"更新版本號失敗: {e}")
                input("請按 Enter 鍵結束...")
                return
        else:
            commit_version = current_version
            print("      [OK] 版本號保持不變。")

        print("\n[2/2] 正在將最新程式碼備份到 GitHub...")
        subprocess.run(["git", "add", "."])
        subprocess.run(["git", "commit", "-m", f"備份程式碼 v{commit_version}: {update_notes}"])
        subprocess.run(["git", "push"])
        print("\n==============================================")
        print("   [OK] 備份完成！程式碼已成功推送上傳 GitHub！")
        print("   (未進行 EXE 打包與建立 Release 發布版)")
        print("==============================================")
        input("\n請按 Enter 鍵關閉視窗...")
        return

    # ==========================================
    # 模式 1 & 2：打包並發布 Release (穩定/預覽版)
    # ==========================================
    print("\n[1/6] 正在檢查 GitHub 授權狀態...")
    has_gh = check_gh_login()
    if not has_gh:
        print("\n無法使用自動上傳，請取消這次發布，或改用手動發布。")
        input("請按 Enter 鍵結束...")
        return
        
    new_version = input(f"\n請輸入本次發布的版本號 [直接按 Enter 預設為 {suggested_version}]: ").strip()
    if not new_version:
        new_version = suggested_version
        
    print(f"[OK] 本次設定發布版本: {new_version}")
    
    update_notes = input("\n請簡單輸入這次發布的更新內容 (例如: 修復閃退問題): ").strip()
    if not update_notes:
        update_notes = "一般更新與修復"
        
    print(f"\n[2/6] 正在更新 main.py 與 使用說明.md 內的版本號為 {new_version}...")
    try:
        new_content = re.sub(r'APP_VERSION\s*=\s*"[^"]+"', f'APP_VERSION = "{new_version}"', content)
        with open("main.py", "w", encoding="utf-8") as f:
            f.write(new_content)
            
        if os.path.exists("使用說明.md"):
            with open("使用說明.md", "r", encoding="utf-8") as f:
                md_content = f.read()
            md_content = re.sub(r'使用說明 \(v[^)]+\)', f'使用說明 (v{new_version})', md_content)
            md_content = re.sub(r'Version [0-9.]+', f'Version {new_version}', md_content)
            with open("使用說明.md", "w", encoding="utf-8") as f:
                f.write(md_content)
            print("      [OK] 使用說明.md 版本號已同步。")
    except Exception as e:
        print(f"更新版本號失敗: {e}")
        input("請按 Enter 鍵結束...")
        return

    print("\n[3/6] 正在打包成執行檔 (這需要 1~2 分鐘，請耐心等候)...")
    subprocess.run([sys.executable, "-m", "PyInstaller", "--noconfirm", "--onefile", "--windowed", "--icon=icon.ico", "--add-data", "icon.ico;.", "--name", "CYT_YTDL", "main.py"])
    
    exe_path = os.path.join("dist", "CYT_YTDL.exe")
    zip_path = os.path.join("dist", "CYT_YTDL.zip")
    txt_path = os.path.join("dist", "使用說明.txt")
    
    if not os.path.exists(exe_path):
        print(f"\n[Error] 打包失敗，找不到 {exe_path}")
        input("請按 Enter 鍵結束...")
        return

    print("\n[3.3/6] 正在同步產生文字版：使用說明.txt...")
    convert_md_to_txt("使用說明.md", txt_path)

    print("\n[3.5/6] 正在將執行檔與程式說明壓縮為 ZIP...")
    try:
        import zipfile
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            zipf.write(exe_path, os.path.basename(exe_path))
            if os.path.exists(txt_path):
                zipf.write(txt_path, os.path.basename(txt_path))
        print(f"      [OK] 壓縮完成: {zip_path} (含使用說明.txt)")
    except Exception as e:
        print(f"      [Error] 壓縮失敗: {e}")
        input("請按 Enter 鍵結束...")
        return

    print("\n[4/6] 正在將最新程式碼備份到 GitHub...")
    subprocess.run(["git", "add", "."])
    # 根據模式設定 Commit Message
    commit_msg = f"發布測試預覽版 v{new_version}: {update_notes}" if mode_choice == "2" else f"發布正式版本 v{new_version}: {update_notes}"
    subprocess.run(["git", "commit", "-m", commit_msg])
    subprocess.run(["git", "push"])
    
    print("\n[5/6] 正在建立 Release 並同時上傳所有檔案...")
    print("      (包含 EXE、ZIP 與 使用說明.txt，上傳可能需要 1~2 分鐘)")
    
    cmd_release = [
        "gh", "release", "create", f"v{new_version}", 
        exe_path, 
        zip_path,
        txt_path,
        "--title", f"v{new_version}", 
        "--notes", update_notes
    ]
    
    # 模式 2：發布測試版 (Pre-release)
    if mode_choice == "2":
        cmd_release.append("--prerelease")
        print("      [INFO] 本次將建立為 Pre-release (測試預覽版)")
        
    result = subprocess.run(cmd_release, capture_output=True, text=True, encoding='utf-8', errors='ignore')
    
    if result.returncode == 0:
        print("\n[Success] 發布成功！EXE 與 ZIP 檔案皆已自動上傳完畢！")
        print("\n[6/6] 正在為您開啟最終的發布網頁以供確認...")
        release_url = f"https://github.com/mathced-com/CYT_YTDL/releases/tag/v{new_version}"
        webbrowser.open(release_url)
    else:
        print(f"\n[Error] 自動發布失敗: {result.stderr}")
    
    print("\n==============================================")
    print("      流程結束，所有使用者已可接收自動更新！")
    print("==============================================")
    input("\n請按 Enter 鍵關閉視窗...")

if __name__ == "__main__":
    main()