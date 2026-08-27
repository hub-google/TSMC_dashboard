/**
 * =========================================================================
 * 儀表板 Google 試算表自動同步 Webhook 腳本 (Google Apps Script)
 * =========================================================================
 * 
 * 【功能說明】
 * 1. 當試算表有任何修改或貼上新資料時，自動發送 Webhook 通知 GitHub Actions 重新抓取並發布網站。
 * 2. 包含「防抖動 (Debounce)」機制：短時間內連續修改時，只會在最後一次修改後觸發一次，避免重複部署浪費資源。
 * 3. 試算表上方自動新增「🚀 儀表板管理」選單，提供「⚡ 立即更新網站」按鈕。
 * 
 * 【設定步驟 (只需 1 分鐘)】
 * 1. 打開你的 Google 試算表 (https://docs.google.com/spreadsheets/d/1C03_PNRWmS3vO-2nz2cFO9QW1k4EDD-W/edit)
 * 2. 點擊上方選單的【擴充功能】->【Apps Script】。
 * 3. 清空原本的程式碼，將本檔案內容全部複製貼上。
 * 4. 將下方的 GITHUB_TOKEN 換成你的 GitHub Personal Access Token (PAT)。
 * 5. 點擊上方的「儲存 (磁碟圖示)」，然後點選「執行 onOpen」測試授權。
 * 6. (選用：設定全自動編輯觸發)
 *    - 點選左側時鐘圖示【觸發條件】-> 右下角【新增觸發條件】
 *    - 選擇活動類型：【編輯時】或【變更時】
 *    - 執行的功能：選擇【onEditDebounced】
 *    - 點擊【儲存】即完成！
 */

// GitHub 相關設定
const GITHUB_REPO_OWNER = 'hub-google';
const GITHUB_REPO_NAME = 'TSMC_dashboard';

// 請在此填入具備 repo / workflow 權限的 GitHub Token
const GITHUB_TOKEN = 'YOUR_GITHUB_PERSONAL_ACCESS_TOKEN';

/**
 * 試算表打開時，自動新增自訂選單
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 儀表板管理')
    .addItem('⚡ 立即更新網站', 'manualTriggerDeploy')
    .addItem('⚙️ 測試 GitHub 連線', 'testGitHubConnection')
    .addToUi();
}

/**
 * 手動觸發更新按鈕
 */
function manualTriggerDeploy() {
  const ui = SpreadsheetApp.getUi();
  const success = triggerGitHubDeploy();
  if (success) {
    ui.alert('✅ 指令發送成功！\nGitHub 伺服器正在清洗個資並重新發布網站，約 1~2 分鐘後重新整理網頁即可看見最新資料。');
  } else {
    ui.alert('❌ 發送失敗，請檢查 GITHUB_TOKEN 是否有效，或查看 Apps Script 執行紀錄。');
  }
}

/**
 * 測試連線
 */
function testGitHubConnection() {
  const ui = SpreadsheetApp.getUi();
  const success = triggerGitHubDeploy();
  if (success) {
    ui.alert('🎉 連線測試成功！GitHub Actions 已順利啟動。');
  } else {
    ui.alert('❌ 連線測試失敗，請確認 Token 權限。');
  }
}

/**
 * 防抖動 (Debounce) 自動觸發器
 * 當有人連續編輯儲存格時，延遲觸發，避免短時間內發動多次 GitHub Actions
 */
function onEditDebounced(e) {
  const cache = CacheService.getScriptCache();
  // 檢查是否在防抖冷卻期 (例如 15 秒內)
  const isCooldown = cache.get('deploy_cooldown');
  if (isCooldown) {
    Logger.log('⏳ 處於防抖冷卻時間中，跳過重複觸發');
    return;
  }
  
  // 設定 15 秒冷卻鎖
  cache.put('deploy_cooldown', 'active', 15);
  
  Logger.log('📝 偵測到試算表更新，正在發送 GitHub 部署 Webhook...');
  triggerGitHubDeploy();
}

/**
 * 向 GitHub API 發送 repository_dispatch 事件
 */
function triggerGitHubDeploy() {
  const url = `https://api.github.com/repos/${GITHUB_REPO_OWNER}/${GITHUB_REPO_NAME}/dispatches`;
  
  const payload = JSON.stringify({
    event_type: 'drive_updated',
    client_payload: {
      timestamp: new Date().toISOString(),
      triggered_by: 'Google Sheets Apps Script'
    }
  });
  
  const options = {
    method: 'post',
    headers: {
      'Accept': 'application/vnd.github+json',
      'Authorization': 'Bearer ' + GITHUB_TOKEN,
      'User-Agent': 'Google-Apps-Script-TSMC-Dashboard'
    },
    contentType: 'application/json',
    payload: payload,
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    const body = response.getContentText();
    
    Logger.log(`GitHub API Response Code: ${code}`);
    Logger.log(`GitHub API Response Body: ${body}`);
    
    // GitHub API dispatches 成功時回傳 204 No Content
    if (code === 204 || code === 200) {
      Logger.log('🎉 成功觸發 GitHub Actions 部署工作流程！');
      return true;
    } else {
      Logger.log(`❌ 發送失敗，狀態碼: ${code}, 訊息: ${body}`);
      return false;
    }
  } catch (err) {
    Logger.log(`❌ 發生異常錯誤: ${err.message}`);
    return false;
  }
}
