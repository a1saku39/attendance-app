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
    
    // 2. 部署ごとの打刻データシートを取得
    var departmentSheetName = '打刻_' + department;
    var logSheet = ss.getSheetByName(departmentSheetName);

    // 履歴取得リクエストの場合
    if (action === 'get_history') {
      var historyData = [];
      if (logSheet && logSheet.getLastRow() > 1) {
        // 直近のデータを取得（最大30日分程度）
        var lastRow = logSheet.getLastRow();
        var startRow = Math.max(2, lastRow - 30);
        var numRows = lastRow - startRow + 1;
        
        // A:日にち, B:社員コード, C:名前, D:種別出勤, E:出勤時刻, F:種別退勤, G:退勤時刻, H:備考
        var range = logSheet.getRange(startRow, 1, numRows, 8);
        var values = range.getValues();
        
        // 新しい順に並べ替えるために逆順で走査
        for (var i = values.length - 1; i >= 0; i--) {
          var row = values[i];
          // 社員コードが一致するものだけ抽出
          if (String(row[1]) === String(employeeId)) {
            var dateVal = row[0];
            if (dateVal instanceof Date) {
              dateVal = Utilities.formatDate(dateVal, "Asia/Tokyo", "yyyy/MM/dd");
            }
            
            historyData.push({
              date: dateVal,
              clockInTime: row[4] || '',  // E列
              clockOutTime: row[6] || '', // G列
              remarks: row[7] || ''       // H列
            });
            
            // 最大10件まで
            if (historyData.length >= 10) break;
          }
        }
      }
      
      var output = ContentService.createTextOutput(JSON.stringify({
        result: 'success',
        username: username,
        department: department,
        history: historyData
      }));
      output.setMimeType(ContentService.MimeType.JSON);
      return output;
    }
    
    // 以下、打刻記録処理 (in/out)
    
    // 新しいヘッダー定義: A:日にち, B:社員コード, C:名前, D:種別出勤, E:出勤時刻, F:種別退勤, G:退勤時刻, H:備考
    if (!logSheet) {
      // 部署別シートがない場合は作成
      logSheet = ss.insertSheet(departmentSheetName);
      logSheet.appendRow(['日にち', '社員コード', '名前', '種別出勤', '出勤時刻', '種別退勤', '退勤時刻', '備考']);
      debugSheet.appendRow([new Date(), '部署別シート「' + departmentSheetName + '」を新規作成しました']);
    } else if (logSheet.getLastRow() === 0) {
      logSheet.appendRow(['日にち', '社員コード', '名前', '種別出勤', '出勤時刻', '種別退勤', '退勤時刻', '備考']);
    }
    
    // 日付と時刻のフォーマット
    var dateStr = Utilities.formatDate(timestamp, "Asia/Tokyo", "yyyy/MM/dd");
    var timeStr = Utilities.formatDate(timestamp, "Asia/Tokyo", "HH:mm");
    var actionText = action === 'in' ? '出勤' : '退勤';
    
    // 行の更新または追加
    updateOrAppendRow(logSheet, {
      date: dateStr,
      id: employeeId,
      name: username,
      action: action,
      time: timeStr,
      remarks: remarks
    });
    
    debugSheet.appendRow([new Date(), '部署別シート「' + departmentSheetName + '」に記録完了']);
    
    // 3. 全体の打刻データにも記録（バックアップとして従来の形式で追記）
    var allLogSheet = ss.getSheetByName('打刻データ_全体');
    if (!allLogSheet) {
      allLogSheet = ss.insertSheet('打刻データ_全体');
      allLogSheet.appendRow(['日時', '社員コード', '名前', '部署', '種別', '備考']);
    } else if (allLogSheet.getLastRow() === 0) {
      allLogSheet.appendRow(['日時', '社員コード', '名前', '部署', '種別', '備考']);
    }
    
    // 全体シートは従来のログ形式を維持
    var days = ['日', '月', '火', '水', '木', '金', '土'];
    var dayOfWeek = days[timestamp.getDay()];
    var formattedFullDate = Utilities.formatDate(timestamp, "Asia/Tokyo", "yyyy/MM/dd") + ' (' + dayOfWeek + ') ' + Utilities.formatDate(timestamp, "Asia/Tokyo", "HH:mm:ss");
    var allRecordData = [formattedFullDate, employeeId, username, department, actionText, remarks];
    allLogSheet.appendRow(allRecordData);
    
    // レスポンス作成
    var output = ContentService.createTextOutput(JSON.stringify({
      result: 'success', 
      username: username,
      department: department,
      action: actionText,
      timestamp: formattedFullDate
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

// 指定された日付と社員コードの行を探して更新、なければ新規追加する関数
function updateOrAppendRow(sheet, data) {
  var lastRow = sheet.getLastRow();
  var foundRow = -1;
  
  // データがある場合、既存の行を検索（直近100行程度を検索対象とする）
  if (lastRow > 1) {
    var startRow = Math.max(2, lastRow - 100);
    var numRows = lastRow - startRow + 1;
    var values = sheet.getRange(startRow, 1, numRows, 2).getValues(); // A列(日付)とB列(コード)を取得
    
    // 下から上に検索
    for (var i = values.length - 1; i >= 0; i--) {
      // 日付と社員コードが一致するか確認
      // シートの日付がDateオブジェクトの場合は文字列に変換して比較
      var sheetDate = values[i][0];
      if (sheetDate instanceof Date) {
        sheetDate = Utilities.formatDate(sheetDate, "Asia/Tokyo", "yyyy/MM/dd");
      }
      
      if (String(sheetDate) === String(data.date) && String(values[i][1]) === String(data.id)) {
        foundRow = startRow + i;
        break;
      }
    }
  }
  
  if (foundRow > 0) {
    // 既存の行を更新
    if (data.action === 'in') {
      sheet.getRange(foundRow, 4).setValue('出勤'); // D列: 種別出勤
      sheet.getRange(foundRow, 5).setValue(data.time); // E列: 出勤時刻
      sheet.getRange(foundRow, 8).setValue(data.remarks); // H列: 備考
    } else if (data.action === 'out') {
      sheet.getRange(foundRow, 6).setValue('退勤'); // F列: 種別退勤
      sheet.getRange(foundRow, 7).setValue(data.time); // G列: 退勤時刻
      // 退勤時も備考を更新（上書き）
      sheet.getRange(foundRow, 8).setValue(data.remarks); // H列: 備考
    }
  } else {
    // 新規行を追加
    // A:日にち, B:社員コード, C:名前, D:種別出勤, E:出勤時刻, F:種別退勤, G:退勤時刻, H:備考
    var rowData = [
      data.date,
      data.id,
      data.name,
      data.action === 'in' ? '出勤' : '',
      data.action === 'in' ? data.time : '',
      data.action === 'out' ? '退勤' : '',
      data.action === 'out' ? data.time : '',
      data.remarks
    ];
    sheet.appendRow(rowData);
  }
}
