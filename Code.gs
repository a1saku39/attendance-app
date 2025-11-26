function doPost(e) {
  try {
    // デバッグ用ログシートの準備
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var debugSheet = ss.getSheetByName('デバッグログ');
    if (!debugSheet) {
      debugSheet = ss.insertSheet('デバッグログ');
      debugSheet.appendRow(['時刻', 'ログ内容']);
    }
    
    // 受信データのログ
    debugSheet.appendRow([new Date(), '受信データ: ' + e.postData.contents]);
    
    // POSTデータをパース
    var jsonData = JSON.parse(e.postData.contents);
    var employeeId = jsonData.employeeId;
    var action = jsonData.action; // 'in' or 'out'
    var timestamp = new Date(jsonData.timestamp);
    var remarks = jsonData.remarks || ''; // 備考
    
    debugSheet.appendRow([new Date(), '社員コード: ' + employeeId + ', アクション: ' + action]);
    
    // 1. 社員マスタから名前と部署を検索
    var masterSheet = ss.getSheetByName('社員マスタ');
    if (!masterSheet) {
      // マスタシートがない場合は作成してヘッダーを追加
      masterSheet = ss.insertSheet('社員マスタ');
      masterSheet.appendRow(['社員コード', '氏名', '部署']);
      debugSheet.appendRow([new Date(), '社員マスタシートを新規作成しました']);
    }
    
    var username = '未登録社員(' + employeeId + ')';
    var department = '未設定';
    var lastRow = masterSheet.getLastRow();
    
    if (lastRow > 1) {
      // 2行目以降のデータを取得（A列:コード, B列:氏名, C列:部署）
      var values = masterSheet.getRange(2, 1, lastRow - 1, 3).getValues();
      for (var i = 0; i < values.length; i++) {
        // 文字列として比較するためにString()を使用
        if (String(values[i][0]) === String(employeeId)) {
          username = values[i][1];
          department = values[i][2] || '未設定'; // 部署が空の場合は「未設定」
          debugSheet.appendRow([new Date(), '社員マスタから検索: ' + username + ' (部署: ' + department + ')']);
          break;
        }
      }
    }
    
    if (username.indexOf('未登録') !== -1) {
      debugSheet.appendRow([new Date(), '警告: 社員コード ' + employeeId + ' はマスタに未登録です']);
    }
    
    // 2. 部署ごとの打刻データシートに記録
    var departmentSheetName = '打刻_' + department;
    var logSheet = ss.getSheetByName(departmentSheetName);
    
    if (!logSheet) {
      // 部署別シートがない場合は作成
      logSheet = ss.insertSheet(departmentSheetName);
      logSheet.appendRow(['日時', '社員コード', '名前', '種別', '備考']);
      debugSheet.appendRow([new Date(), '部署別シート「' + departmentSheetName + '」を新規作成しました']);
    } else if (logSheet.getLastRow() === 0) {
      logSheet.appendRow(['日時', '社員コード', '名前', '種別', '備考']);
    }
    
    // 3. 全体の打刻データにも記録（オプション）
    var allLogSheet = ss.getSheetByName('打刻データ_全体');
    if (!allLogSheet) {
      allLogSheet = ss.insertSheet('打刻データ_全体');
      allLogSheet.appendRow(['日時', '社員コード', '名前', '部署', '種別', '備考']);
      debugSheet.appendRow([new Date(), '全体打刻データシートを新規作成しました']);
    } else if (allLogSheet.getLastRow() === 0) {
      allLogSheet.appendRow(['日時', '社員コード', '名前', '部署', '種別', '備考']);
    }
    
    // 日本時間のフォーマット
    var days = ['日', '月', '火', '水', '木', '金', '土'];
    var dayOfWeek = days[timestamp.getDay()];
    var formattedDate = Utilities.formatDate(timestamp, "Asia/Tokyo", "yyyy/MM/dd") + ' (' + dayOfWeek + ') ' + Utilities.formatDate(timestamp, "Asia/Tokyo", "HH:mm:ss");
    var actionText = action === 'in' ? '出勤' : '退勤';
    
    // 部署別シート用のデータ
    var recordData = [formattedDate, employeeId, username, actionText, remarks];
    debugSheet.appendRow([new Date(), '記録データ: ' + JSON.stringify(recordData)]);
    
    // 部署別シートに記録
    logSheet.appendRow(recordData);
    debugSheet.appendRow([new Date(), '部署別シート「' + departmentSheetName + '」に記録完了']);
    
    // 全体シート用のデータ（部署情報を含む）
    var allRecordData = [formattedDate, employeeId, username, department, actionText, remarks];
    allLogSheet.appendRow(allRecordData);
    debugSheet.appendRow([new Date(), '全体シートに記録完了']);
    
    // レスポンス作成
    var output = ContentService.createTextOutput(JSON.stringify({
      result: 'success', 
      username: username,
      department: department,
      action: actionText,
      timestamp: formattedDate
    }));
    output.setMimeType(ContentService.MimeType.JSON);
    return output;
    
  } catch (error) {
    // エラーもログに記録
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var debugSheet = ss.getSheetByName('デバッグログ');
      if (debugSheet) {
        debugSheet.appendRow([new Date(), 'エラー: ' + error.toString()]);
      }
    } catch (e) {
      // ログ記録自体が失敗した場合は無視
    }
    
    var output = ContentService.createTextOutput(JSON.stringify({
      result: 'error', 
      message: error.toString()
    }));
    output.setMimeType(ContentService.MimeType.JSON);
    return output;
  }
}
