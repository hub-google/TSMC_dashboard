import subprocess
import os
import sys
from datetime import datetime

if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def trigger_cloud_update():
    # 確保不管從哪裡執行，工作目錄都在腳本所在的資料夾
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    print("🚀 正在通知 GitHub 雲端伺服器抓取最新資料...")
    
    # 建立一個空的 commit 來觸發 GitHub Actions
    commit_msg = f"Trigger cloud update: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    
    try:
        subprocess.run(["git", "commit", "--allow-empty", "-m", commit_msg], check=True)
    except subprocess.CalledProcessError as e:
        print(f"❌ 無法建立更新指令: {e}")
        return
        
    print("☁️ 正在發送指令到雲端...")
    try:
        subprocess.run(["git", "push"], check=True)
    except subprocess.CalledProcessError as e:
        print(f"❌ 無法連線到 GitHub: {e}")
        return
    
    print("\n🎉 更新指令發送成功！")
    print("✅ 現在 GitHub 會「自動」去 Google Drive 下載檔案並發布。")
    print("👉 請等待約 1~2 分鐘，然後直接進入網頁即可看到最新資料！")
    print("（你的電腦不需要再下載任何 Excel 檔案了！）")

if __name__ == "__main__":
    trigger_cloud_update()
