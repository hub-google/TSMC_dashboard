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
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    print("🚀 正在通知 GitHub 雲端伺服器抓取最新資料...")
    
    # 優先嘗試透過 GitHub CLI 直接觸發 workflow_dispatch
    try:
        res = subprocess.run(["gh", "workflow", "run", "deploy.yml"], capture_output=True, text=True, check=True)
        print("🎉 指令發送成功！(透過 GitHub CLI 觸發)")
        print("✅ 現在 GitHub 會「自動」去 Google Drive 下載檔案、去除個資並發布。")
        print("👉 請等待約 1~2 分鐘，直接進入網頁即可看到最新資料！")
        return
    except Exception:
        pass

    # 備用語音：透過 Git empty commit
    commit_msg = f"Trigger cloud update: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    try:
        subprocess.run(["git", "commit", "--allow-empty", "-m", commit_msg], check=True)
        print("☁️ 正在發送指令到雲端...")
        subprocess.run(["git", "push"], check=True)
        print("\n🎉 更新指令發送成功！")
        print("✅ 現在 GitHub 會「自動」去 Google Drive 下載檔案並發布。")
        print("👉 請等待約 1~2 分鐘，直接進入網頁即可看到最新資料！")
    except subprocess.CalledProcessError as e:
        print(f"❌ 更新失敗: {e}")

if __name__ == "__main__":
    trigger_cloud_update()
