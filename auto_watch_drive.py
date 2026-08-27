"""
本機 Google Drive 檔案監聽守護程式 (Local Watcher Daemon)
用途：當本地 G 槽或本專案中的 Excel / 相關資料更新時，自動觸發 GitHub Actions 進行清洗與網站發布。
"""

import time
import os
import sys
import subprocess
from datetime import datetime

# 目標監控目錄 (預設為本腳本所在資料夾或指定 G 槽路徑)
WATCH_DIR = os.path.dirname(os.path.abspath(__file__))
# 監控的副檔名
WATCH_EXTENSIONS = ('.xlsx', '.xls', '.csv')

def trigger_github_workflow():
    print(f"\n[{datetime.now().strftime('%H:%M:%S')}] 🚀 偵測到資料變更！正在通知 GitHub Actions 發布新版本...")
    try:
        # 使用 GitHub CLI 觸發 dispatch
        result = subprocess.run(
            ["gh", "workflow", "run", "deploy.yml"],
            capture_output=True,
            text=True,
            check=True
        )
        print("✅ 成功發送觸發指令至 GitHub Actions！")
        print("👉 GitHub 伺服器正在自動抓取資料、去除個資並更新網站中...")
    except subprocess.CalledProcessError as e:
        print(f"❌ 觸發失敗: {e.stderr.strip() if e.stderr else e}")
    except FileNotFoundError:
        print("❌ 找不到 gh 指令，請確認 GitHub CLI 是否已安裝。")

def get_file_mtimes(directory):
    """取得資料夾下所有 Excel 檔案的最後修改時間"""
    mtimes = {}
    for root, _, files in os.walk(directory):
        # 排除 .git 和 node_modules
        if '.git' in root or 'node_modules' in root or 'dist' in root:
            continue
        for f in files:
            if f.endswith(WATCH_EXTENSIONS) and not f.startswith('~$'):
                path = os.path.join(root, f)
                try:
                    mtimes[path] = os.path.getmtime(path)
                except OSError:
                    pass
    return mtimes

def main():
    if hasattr(sys.stdout, 'reconfigure'):
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except Exception:
            pass

    print("=" * 60)
    print("👀 Google Drive / 本地 Excel 自動監控守護程式已啟動")
    print(f"📁 監控目錄: {WATCH_DIR}")
    print("💡 只要在此目錄或 Google 雲端資料夾存檔 Excel，就會秒級自動觸發網站更新！")
    print("（按 Ctrl + C 即可停止監控）")
    print("=" * 60)

    last_mtimes = get_file_mtimes(WATCH_DIR)

    try:
        while True:
            time.sleep(3)  # 每 3 秒檢查一次
            current_mtimes = get_file_mtimes(WATCH_DIR)

            changed = False
            for path, mtime in current_mtimes.items():
                if path not in last_mtimes or mtime > last_mtimes[path]:
                    print(f"\n📂 偵測到檔案變更: {os.path.basename(path)}")
                    changed = True
                    break

            if changed:
                last_mtimes = current_mtimes
                trigger_github_workflow()
                # 觸發後冷卻 10 秒避免連續存檔重複觸發
                time.sleep(10)
                last_mtimes = get_file_mtimes(WATCH_DIR)

    except KeyboardInterrupt:
        print("\n🛑 已停止監控。")

if __name__ == "__main__":
    main()
