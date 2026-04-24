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
    print(f"目前版本為: {current_version}")
    
    new_version = input(f"請輸入新的版本號 [直接按 Enter 預設為 {suggested_version}]: ").strip()
    if not new_version:
        new_version = suggested_version
        
    print(f"\n[OK] 將發布新版本: {new_version}")
    
    update_notes = input("\n請簡單輸入這次更新的內容 (例如: 修復閃退問題): ").strip()
    if not update_notes:
        update_notes = "一般更新與修復"
        
    print("\n[1/6] 正在檢查 GitHub 授權狀態...")
    has_gh = check_gh_login()
    if not has_gh:
        print("\n無法使用自動上傳，請取消這次發布，或改用手動發布。")
        input("請按 Enter 鍵結束...")
        return
        
    print(f"\n[2/6] 正在更新 main.py 內的版本號為 {new_version}...")
    try:
        new_content = re.sub(r'APP_VERSION\s*=\s*"[^"]+"', f'APP_VERSION = "{new_version}"', content)
        with open("main.py", "w", encoding="utf-8") as f:
            f.write(new_content)
    except Exception as e:
        print(f"更新版本號失敗: {e}")
        input("請按 Enter 鍵結束...")
        return

    print("\n[3/6] 正在打包成執行檔 (這需要 1~2 分鐘，請耐心等候)...")
    print("      (正在確認 PyInstaller 封裝套件是否安裝)")
    subprocess.run(["py", "-3", "-m", "pip", "install", "pyinstaller"], capture_output=True)
    subprocess.run(["py", "-3", "-m", "PyInstaller", "--noconfirm", "--onefile", "--windowed", "--name", "CYT_YTDL", "main.py"])
    
    exe_path = os.path.join("dist", "CYT_YTDL.exe")
    if not os.path.exists(exe_path):
        print(f"\n[Error] 打包失敗，找不到 {exe_path}")
        input("請按 Enter 鍵結束...")
        return

    print("\n[4/6] 正在將最新程式碼備份到 GitHub...")
    subprocess.run(["git", "add", "."])
    subprocess.run(["git", "commit", "-m", f"發布新版本 v{new_version}: {update_notes}"])
    subprocess.run(["git", "push"])
    
    print("\n[5/6] 正在自動建立 GitHub Release 並將 CYT_YTDL.exe 上傳至雲端...")
    print("      (檔案有點大，上傳可能需要幾十秒鐘，請勿關閉視窗)")
    result = subprocess.run([
        "gh", "release", "create", f"v{new_version}", 
        exe_path, 
        "--title", f"v{new_version}", 
        "--notes", update_notes
    ], capture_output=True, text=True, encoding='utf-8', errors='ignore')
    
    if result.returncode == 0:
        print("\n[Success] 發布成功！檔案已由程式幫您自動上傳完畢！")
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
