function doPost(e) {
  try {
    // デバッグ用ログシートの準備（最優先で確保）
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var debugSheet = ss.getSheetByName('デバッグログ');
    if (!debugSheet) {
      debugSheet = ss.insertSheet('デバッグログ');
      debugSheet.appendRow(['時刻', 'ログ内容']);
    }

    // バージョン確認用ログ (v3.0)
    debugSheet.appendRow([new Date(), '[INFO] doPost実行 (v3.0 - Monthly Approval)']);
    
    // APIエンドポイントの振り分け
    var jsonData = JSON.parse(e.postData.contents);
    var action = jsonData.action;
    
    // 月次確認機能のAPI
    if (action === 'getMonthlyData') {
      return getMonthlyData(e);
    } else if (action === 'recordApproval') {
      return recordApproval(e);
    } else if (action === 'getPersonalMonthlyData') {
      return getPersonalMonthlyData(e);
    } else if (action === 'getApproverDashboard') { // 追加: 承認者ダッシュボード用
      return getApproverDashboard(e);
    }
    
    // 以下、通常の打刻処理（action が 'in' または 'out' の場合のみ）
    if (!jsonData.employeeId || !jsonData.timestamp || (action !== 'in' && action !== 'out')) {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'error',
        message: '必要なパラメータが不足しているか、無効なアクションです'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // ... (既存の打刻処理コードは省略せずに保持する必要がありますが、
    // ここではwrite_to_fileで全体を置き換えるか、あるいは関数を追加する形にするか判断が必要。
    // Code.gsは大きいため、replace_file_content で関数を追加・変更する方が安全です)
    
  } catch (error) {
     var output = ContentService.createTextOutput(JSON.stringify({
      result: 'error', 
      message: error.toString()
    }));
    output.setMimeType(ContentService.MimeType.JSON);
    return output;
  }
}
