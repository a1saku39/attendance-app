function doPost(e) {
  try {
    // デバッグログ出力の無効化（ユーザー要望により停止）
    // 既存のコードが debugSheet.appendRow を呼び出してもエラーにならないようにダミーオブジェクトを作成
    var debugSheet = {
      appendRow: function(row) {
        // ログはスプレッドシートには出力せず、Cloud Logging (console.log) に出力する
        // try-catchで囲んで安全に実行
        try {
           var msg = row.map(function(item){ return item instanceof Date ? Utilities.formatDate(item, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss") : item; }).join(' | ');
           console.log(msg);
        } catch(e) {}
      }
    };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // バージョン確認用ログ (v3.2 DebugDisabled)
    debugSheet.appendRow([new Date(), '[INFO] doPost実行 (v3.2 DebugDisabled)']);
    debugSheet.appendRow([new Date(), '受信データ: ' + e.postData.contents]);

    // APIエンドポイントの振り分け
    var jsonData;
    try {
        jsonData = JSON.parse(e.postData.contents);
    } catch (e) {
        debugSheet.appendRow([new Date(), 'エラー: JSONパース失敗']);
        return ContentService.createTextOutput(JSON.stringify({result:'error', message:'Invalid JSON'})).setMimeType(ContentService.MimeType.JSON);
    }
    var action = jsonData.action;
    
    // 月次確認機能のAPI（これらは打刻処理ではないので、ここで処理を完了させる）
    if (action === 'getMonthlyData') {
      debugSheet.appendRow([new Date(), 'API実行: getMonthlyData']);
      return getMonthlyData(e);
    } else if (action === 'recordApproval') {
      debugSheet.appendRow([new Date(), 'API実行: recordApproval']);
      return recordApproval(e);
    } else if (action === 'getDepartments') {
      debugSheet.appendRow([new Date(), 'API実行: getDepartments']);
      return getDepartments();
    } else if (action === 'getEmployeesByDepartmentAndMonth') {
      debugSheet.appendRow([new Date(), 'API実行: getEmployeesByDepartmentAndMonth']);
      return getEmployeesByDepartmentAndMonth(e);
    } else if (action === 'approveEmployee') {
      debugSheet.appendRow([new Date(), 'API実行: approveEmployee']);
      return approveEmployee(e);
    } else if (action === 'getPersonalMonthlyData') {
      debugSheet.appendRow([new Date(), 'API実行: getPersonalMonthlyData']);
      return getPersonalMonthlyData(e);
    } else if (action === 'getApproverDashboard') {
      debugSheet.appendRow([new Date(), 'API実行: getApproverDashboard']);
      return getApproverDashboard(e);
    } else if (action === 'updateAttendance') {
       debugSheet.appendRow([new Date(), 'API実行: updateAttendance']);
       return updateDailyAttendance(e);
    } else if (action === 'getHolidaySettings') { // Added handler
       debugSheet.appendRow([new Date(), 'API実行: getHolidaySettings']);
       return getHolidaySettings(e);
    } else if (action === 'saveHolidaySettings') { // Added handler
       debugSheet.appendRow([new Date(), 'API実行: saveHolidaySettings']);
       return saveHolidaySettings(e);
    }
    
    // 以下、通常の打刻処理（action が 'in'、'out'、'location' の場合のみ）
    
    // 打刻処理に必要なフィールドの存在チェック
    if (!jsonData.employeeId || !jsonData.timestamp || 
        (action !== 'in' && action !== 'out' && action !== 'location' && action !== 'holiday')) {
      debugSheet.appendRow([new Date(), 'エラー: 無効なアクションまたはパラメータ不足 (' + action + ')']);
      return ContentService.createTextOutput(JSON.stringify({
        result: 'error',
        message: '必要なパラメータが不足しているか、無効なアクションです'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    debugSheet.appendRow([new Date(), '打刻処理開始: ' + action]);
    
    // POSTデータから必要な情報を取得
    var employeeId = jsonData.employeeId;
    
    // タイムスタンプのバリデーション
    if (!jsonData.timestamp) {
         debugSheet.appendRow([new Date(), 'エラー: timestampが欠落しています']);
         return ContentService.createTextOutput(JSON.stringify({
            result: 'error',
            message: '打刻時刻情報が不足しています'
        })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var timestamp = new Date(jsonData.timestamp);
    var remarks = jsonData.remarks || ''; // 備考
    debugSheet.appendRow([new Date(), '備考受信確認: ' + remarks]); // Debug log

    
    // 時刻変換のデバッグログ
    var debugTimeStr = Utilities.formatDate(timestamp, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
    debugSheet.appendRow([new Date(), '社員コード: ' + employeeId + ', アクション: ' + action]);
    debugSheet.appendRow([new Date(), '受信タイムスタンプ(UTC): ' + jsonData.timestamp + ', JST変換後: ' + debugTimeStr]);

    // 【重要】不正なアクションの混入防止
    // getPersonalMonthlyDataなどがここに来てしまった場合の安全策
    var validWriteActions = ['in', 'out', 'location', 'holiday'];
    if (validWriteActions.indexOf(action) === -1) {
        debugSheet.appendRow([new Date(), 'エラー: 書き込みアクションではありません: ' + action]);
        return ContentService.createTextOutput(JSON.stringify({
            result: 'error',
            message: 'システムエラー: 不正なアクションです'
        })).setMimeType(ContentService.MimeType.JSON);
    }

    // 【重要】日付妥当性チェック (1970年問題対策)
    if (timestamp.getFullYear() < 2024) {
        debugSheet.appendRow([new Date(), 'エラー: 不正な日付(1970年等)のため無視しました: ' + debugTimeStr]);
        return ContentService.createTextOutput(JSON.stringify({
            result: 'error',
            message: '日付情報が不正です。再試行してください。'
        })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // 1. 社員マスタから名前と部署を検索
    var masterSheet = ss.getSheetByName('社員マスタ');
    if (!masterSheet) {
      // マスタシートがない場合は作成してヘッダーを追加
      masterSheet = ss.insertSheet('社員マスタ');
      masterSheet.appendRow(['社員コード', '氏名', '部署', '承認対象部署', '第1承認者', '第2承認者']);
      debugSheet.appendRow([new Date(), '社員マスタシートを新規作成しました']);
    }
    
    var username = '未登録社員(' + employeeId + ')';
    var department = '未設定';
    var lastRow = masterSheet.getLastRow();
    
    if (lastRow > 1) {
      // 2行目以降のデータを取得(A列:コード, B列:氏名, C列:部署)
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
      debugSheet.appendRow([new Date(), '【重要】未登録社員の打刻: 社員コード=' + employeeId]);
      // 未登録でも打刻_未設定シートに記録されるように処理は継続
    }
    
    // 2. 部署ごとの打刻データシートを取得
    var departmentSheetName = '打刻_' + department;
    var logSheet = ss.getSheetByName(departmentSheetName);


    
    // 以下、打刻記録処理 (in/out)
    
    // 新しいヘッダー定義: A:日にち, B:社員コード, C:名前, D:種別出勤, E:出勤時刻, F:種別退勤, G:退勤時刻, H:勤務時間, I:備考
    if (!logSheet) {
      // 部署別シートがない場合は作成
      logSheet = ss.insertSheet(departmentSheetName);
      logSheet.appendRow(['日にち', '社員コード', '名前', '種別出勤', '出勤時刻', '種別退勤', '退勤時刻', '勤務時間', '備考']);
      debugSheet.appendRow([new Date(), '部署別シート「' + departmentSheetName + '」を新規作成しました']);
    } else if (logSheet.getLastRow() === 0) {
      logSheet.appendRow(['日にち', '社員コード', '名前', '種別出勤', '出勤時刻', '種別退勤', '退勤時刻', '勤務時間', '備考']);
    }
    
    // 日付と時刻のフォーマット
    var dateStr = Utilities.formatDate(timestamp, "Asia/Tokyo", "yyyy/MM/dd");
    var timeStr = Utilities.formatDate(timestamp, "Asia/Tokyo", "HH:mm");
    var actionText = '';
    if (action === 'in') actionText = '出勤';
    else if (action === 'out') actionText = '退勤';
    else if (action === 'location') actionText = '位置情報記録';
    else if (action === 'holiday') actionText = jsonData.option === 'paid_leave' ? '有休休暇' : '代休';
    
    // 行の更新または追加
    updateOrAppendRow(logSheet, {
      date: dateStr,
      id: employeeId,
      name: username,
      action: action,
      time: timeStr,
      remarks: remarks,

      location: jsonData.location, // 位置情報
      option: jsonData.option // 休日オプション
    });
    
    debugSheet.appendRow([new Date(), '部署別シート「' + departmentSheetName + '」に記録完了']);
    
    // 3. 全体の打刻データにも記録(バックアップとして従来の形式で追記)
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

// ヘルパー: B列が「曜日」かどうかチェックし、なければ追加する
function ensureDayOfWeekColumn(sheet) {
  var header = sheet.getRange(1, 2).getValue(); // B1
  if (header !== '曜日') {
    sheet.insertColumnAfter(1); // A列の後ろ(B列)に挿入
    sheet.getRange(1, 2).setValue('曜日');
    return true; // 変更あり
  }
  return false; // 変更なし
}

// 勤怠データの更新処理（月次カレンダーなどからの修正用）
function updateDailyAttendance(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var debugSheet = ss.getSheetByName('デバッグログ');
    var jsonData = JSON.parse(e.postData.contents);
    
    // 必須パラメータ: employeeId, date(YYYY/MM/dd), clockInTime, clockOutTime
    var employeeId = jsonData.employeeId;
    var targetDateStr = jsonData.date; // 修正対象の日付
    var clockInTime = jsonData.clockInTime; // "HH:mm" or ""
    var clockOutTime = jsonData.clockOutTime; // "HH:mm" or ""
    var remarks = jsonData.remarks || "";
    
    if (debugSheet) {
        debugSheet.appendRow([new Date(), '勤怠修正リクエスト: ID=' + employeeId + ', Date=' + targetDateStr]);
    }

    // 1. 部署名の特定
    var masterSheet = ss.getSheetByName('社員マスタ');
    if (!masterSheet) throw new Error('社員マスタが見つかりません');
    
    var department = '未設定';
    var username = '不明';
    var lastRow = masterSheet.getLastRow();
    if (lastRow > 1) {
        var values = masterSheet.getRange(2, 1, lastRow - 1, 3).getValues();
        for(var i=0; i<values.length; i++) {
            if(String(values[i][0]) === String(employeeId)) {
                username = values[i][1];
                department = values[i][2];
                break;
            }
        }
    }
    
    var departmentSheetName = '打刻_' + department;
    var logSheet = ss.getSheetByName(departmentSheetName);
    
    if(!logSheet) {
        logSheet = ss.insertSheet(departmentSheetName);
        logSheet.appendRow(['日にち', '曜日', '社員コード', '名前', '種別出勤', '出勤時刻', '種別退勤', '退勤時刻', '勤務時間', '備考']);
    } else {
        ensureDayOfWeekColumn(logSheet);
    }
    
    // 2. 該当行の検索と更新
    var logLastRow = logSheet.getLastRow();
    var foundRow = -1;
    
    if (logLastRow > 1) {
         // 日付は A列(1列目), IDは C列(3列目) になる
         var dataRange = logSheet.getRange(2, 1, logLastRow - 1, 3).getValues();
         
         for(var i=0; i<dataRange.length; i++) {
             // 日付比較
             var rowDate = dataRange[i][0];
             var rowDateStr = '';
             if (rowDate instanceof Date) {
                 rowDateStr = Utilities.formatDate(rowDate, "Asia/Tokyo", "yyyy/MM/dd");
             } else {
                 rowDateStr = String(rowDate);
             }
             
             // ID比較 (C列 = index 2)
             var rowId = String(dataRange[i][2]);
             
             if (rowDateStr === targetDateStr && rowId === String(employeeId)) {
                 foundRow = i + 2; 
                 break;
             }
         }
    }
    
    if (foundRow > 0) {
        // 更新 (列番号が全体的に +1 される)
        // E:種別出勤(5), F:出勤(6), G:種別退勤(7), H:退勤(8), J:備考(10)
        
        if(clockInTime) {
            logSheet.getRange(foundRow, 5).setValue('出勤'); // E
            logSheet.getRange(foundRow, 6).setValue(clockInTime); // F
        } else {
            logSheet.getRange(foundRow, 5).clearContent();
            logSheet.getRange(foundRow, 6).clearContent();
        }
        
        if(clockOutTime) {
            logSheet.getRange(foundRow, 7).setValue('退勤'); // G
            logSheet.getRange(foundRow, 8).setValue(clockOutTime); // H
        } else {
            logSheet.getRange(foundRow, 7).clearContent();
            logSheet.getRange(foundRow, 8).clearContent();
        }
        
        logSheet.getRange(foundRow, 10).setValue(remarks); // J:備考

        // 勤務時間(I列=9)の計算式再設定
        logSheet.getRange(foundRow, 9).setFormula('=IF(AND(F' + foundRow + '<>"", H' + foundRow + '<>""), TEXT(H' + foundRow + '-F' + foundRow + ', "[h]:mm"), "")');
        
        if(debugSheet) debugSheet.appendRow([new Date(), '既存行を更新しました: ' + foundRow + '行目']);
        
    } else {
        // 新規追加
        // ['日にち', '曜日', '社員コード', '名前', '種別出勤', '出勤時刻', '種別退勤', '退勤時刻', '勤務時間', '備考']
        
        // 曜日の計算
        var targetDate = new Date(targetDateStr);
        var days = ['日', '月', '火', '水', '木', '金', '土'];
        var dayOfWeek = days[targetDate.getDay()];
        
        var newRowData = [
            targetDateStr,
            dayOfWeek, // B: 曜日
            employeeId,
            username,
            clockInTime ? '出勤' : '',
            clockInTime,
            clockOutTime ? '退勤' : '',
            clockOutTime,
            '', // I列: 勤務時間(数式)
            remarks,
            '', // K列: 承認
            '' // L列: 位置情報
        ];
        logSheet.appendRow(newRowData);
        var newRowNum = logSheet.getLastRow();
        // I列=9.  INPUT: F, H
        logSheet.getRange(newRowNum, 9).setFormula('=IF(AND(F' + newRowNum + '<>"", H' + newRowNum + '<>""), TEXT(H' + newRowNum + '-F' + newRowNum + ', "[h]:mm"), "")');

        if(debugSheet) debugSheet.appendRow([new Date(), '新規行を追加しました: ' + newRowNum + '行目']);
    }
    
    // ソート
    if (logSheet.getLastRow() > 1) {
        var sortRange = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logSheet.getLastColumn());
        sortRange.sort([
            {column: 1, ascending: true}, 
            {column: 6, ascending: true} // F列(出勤時刻)でソート
        ]);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      result: 'success', 
      message: '勤怠データを修正しました'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (e) {
      return ContentService.createTextOutput(JSON.stringify({
      result: 'error', 
      message: e.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 指定された日付と社員コードの行を探して更新、なければ新規追加する関数
function updateOrAppendRow(sheet, data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var debugSheet = ss.getSheetByName('デバッグログ');
  
  // 曜日列があるか確認し、なければ追加
  ensureDayOfWeekColumn(sheet);
  
  var lastRow = sheet.getLastRow();
  var foundRow = -1;
  
  debugSheet.appendRow([new Date(), '検索開始: 日付=' + data.date + ', 社員コード=' + data.id + ', アクション=' + data.action]);
  
  // データがある場合、既存の行を検索
  if (lastRow > 1) {
    // 日付(A), 曜日(B), コード(C)
    // 検索範囲: A列〜C列
    var startRow = 2;
    var numRows = lastRow - 1;
    var values = sheet.getRange(startRow, 1, numRows, 3).getValues(); 
    
    // 検索対象
    var formattedSearchDate = String(data.date).trim();
    var formattedSearchId = String(data.id).trim();
    
    // 下から上に検索
    for (var i = values.length - 1; i >= 0; i--) {
      var sheetDate = values[i][0]; // A列
      var sheetId = values[i][2];   // C列
      
      var formattedSheetDate = '';
      var isDateObject = (typeof sheetDate === 'object' && sheetDate !== null && 
                         (sheetDate instanceof Date || Object.prototype.toString.call(sheetDate) === '[object Date]'));
      
      if (isDateObject) {
        formattedSheetDate = Utilities.formatDate(sheetDate, "Asia/Tokyo", "yyyy/MM/dd");
      } else if (sheetDate) {
        formattedSheetDate = String(sheetDate).trim();
      }
      
      var formattedSheetId = String(sheetId).trim();
      
      if (formattedSheetDate === formattedSearchDate && formattedSheetId === formattedSearchId) {
        foundRow = startRow + i;
        debugSheet.appendRow([new Date(), '✓ 既存行を発見: ' + foundRow + '行目']);
        break;
      }
    }
  }

  // 位置情報文字列
  var locationStr = '';
  if (data.location) {
    try {
      locationStr = JSON.stringify(data.location);
    } catch (e) {
      locationStr = 'error: ' + e.message;
    }
  }

  if (foundRow > 0) {
    // 既存行更新 (列番号+1)
    
    if (data.action === 'in') {
      sheet.getRange(foundRow, 5).setValue('出勤'); // E:種別
      sheet.getRange(foundRow, 6).setValue(data.time); // F:出勤時刻
      
      // 備考(J列=10)
      if (data.remarks) {
        var remarksCell = sheet.getRange(foundRow, 10);
        var currentRemarks = String(remarksCell.getValue());
        if (currentRemarks && currentRemarks !== data.remarks && currentRemarks.indexOf(data.remarks) === -1) {
             remarksCell.setValue(currentRemarks + ' ' + data.remarks);
        } else if (!currentRemarks) {
             remarksCell.setValue(data.remarks);
        }
      }
      
      // I列(9): 勤務時間.  F(6), H(8)
      sheet.getRange(foundRow, 9).setFormula('=IF(AND(F' + foundRow + '<>"", H' + foundRow + '<>""), TEXT(H' + foundRow + '-F' + foundRow + ', "[h]:mm"), "")');
      debugSheet.appendRow([new Date(), '出勤時刻を記録: ' + data.time]);
    } else if (data.action === 'out') {
      sheet.getRange(foundRow, 7).setValue('退勤'); // G:種別
      sheet.getRange(foundRow, 8).setValue(data.time); // H:退勤時刻
      // I列(9): 勤務時間
      sheet.getRange(foundRow, 9).setFormula('=IF(AND(F' + foundRow + '<>"", H' + foundRow + '<>""), TEXT(H' + foundRow + '-F' + foundRow + ', "[h]:mm"), "")');
      
      // 備考(J列=10)
      if (data.remarks) {
        var remarksCell = sheet.getRange(foundRow, 10);
        var currentRemarks = String(remarksCell.getValue());
        if (currentRemarks && currentRemarks !== data.remarks && currentRemarks.indexOf(data.remarks) === -1) {
             remarksCell.setValue(currentRemarks + ' ' + data.remarks);
        } else if (!currentRemarks) {
             remarksCell.setValue(data.remarks);
        }
      }
      debugSheet.appendRow([new Date(), '退勤時刻を記録: ' + data.time]);
    } else if (data.action === 'holiday') {
      var holidayText = data.option === 'paid_leave' ? '有給休暇' : '代休';
      sheet.getRange(foundRow, 5).setValue(holidayText); // E
      sheet.getRange(foundRow, 6).clearContent(); // F
      sheet.getRange(foundRow, 7).clearContent(); // G
      sheet.getRange(foundRow, 8).clearContent(); // H
      
      // 備考(J)
      if (data.remarks) {
        var remarksCell = sheet.getRange(foundRow, 10);
        var currentRemarks = String(remarksCell.getValue());
        if (currentRemarks && currentRemarks !== data.remarks && currentRemarks.indexOf(data.remarks) === -1) {
            remarksCell.setValue(currentRemarks + ' ' + data.remarks);
        } else if (!currentRemarks) {
            remarksCell.setValue(data.remarks);
        }
      }
      
      sheet.getRange(foundRow, 9).setValue(''); // I(勤務時間)クリア
      debugSheet.appendRow([new Date(), '休日記録: ' + holidayText]);
    }

    // 位置情報 (L列=12)
    if (locationStr) {
      var currentVal = sheet.getRange(foundRow, 12).getValue(); 
      var locArray = [];
      if (currentVal) {
        try {
          if (typeof currentVal === 'string' && (currentVal.startsWith('[') || currentVal.startsWith('{'))) {
             locArray = JSON.parse(currentVal);
             if (!Array.isArray(locArray)) locArray = [locArray];
          }
        } catch (e) { locArray = []; }
      }
      locArray.push({
        action: data.action,
        time: data.time,
        lat: data.location.lat,
        lng: data.location.lng,
        timestamp: new Date().toISOString()
      });
      sheet.getRange(foundRow, 12).setValue(JSON.stringify(locArray));
    }
  } else {
    // 新規行
    debugSheet.appendRow([new Date(), '新規行を追加します']);
    var initialLocJson = '';
    if (locationStr) {
       initialLocJson = JSON.stringify([{
         action: data.action,
         time: data.time,
         lat: data.location.lat,
         lng: data.location.lng,
         timestamp: new Date().toISOString()
       }]);
    }
    
    // 曜日の計算
    var dateObj = new Date(data.date);
    var days = ['日', '月', '火', '水', '木', '金', '土'];
    var dayOfWeek = days[dateObj.getDay()];
    
    // A:日にち, B:曜日, C:社員コード, D:名前, E:種別出勤, F:出勤時刻, G:種別退勤, H:退勤時刻, I:勤務時間, J:備考, K:承認, L:位置情報
    var rowData = [
      data.date,         // A
      dayOfWeek,         // B (New)
      data.id,           // C
      data.name,         // D
      data.action === 'in' ? '出勤' : (data.action === 'holiday' ? (data.option === 'paid_leave' ? '有給休暇' : '代休') : ''), // E
      data.action === 'in' ? data.time : '', // F
      data.action === 'out' ? '退勤' : '',   // G
      data.action === 'out' ? data.time : '', // H
      '',                // I (数式)
      data.remarks,      // J
      '',                // K
      initialLocJson     // L
    ];
    sheet.appendRow(rowData);
    
    var newRow = sheet.getLastRow();
    // I列(9)数式
    sheet.getRange(newRow, 9).setFormula('=IF(AND(F' + newRow + '<>"", H' + newRow + '<>""), TEXT(H' + newRow + '-F' + newRow + ', "[h]:mm"), "")');
    foundRow = newRow; // For coloring
    debugSheet.appendRow([new Date(), '新規行を追加しました: ' + newRow + '行目']);
  }
  
  // ソート
  if (sheet.getLastRow() > 1) {
    var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    range.sort([
      {column: 1, ascending: true}, 
      {column: 6, ascending: true} // F:出勤時刻
    ]);
  }
  
  // === 着色処理 (Company Holiday / Saturday Logic) ===
  if (foundRow > 0) {
      applyRowColoring(sheet, foundRow, data.date);
  }
}

// 着色ロジック
function applyRowColoring(sheet, row, dateStr) {
    // 1. 休日判定
    var dateObj = new Date(dateStr);
    var yearMonth = Utilities.formatDate(dateObj, "Asia/Tokyo", "yyyy-MM");
    var holidaysMap = getAllHolidaysMap(yearMonth);
    
    // 土曜日判定
    var isSaturday = (dateObj.getDay() === 6);
    
    // 会社休日判定
    var formattedDate = Utilities.formatDate(dateObj, "Asia/Tokyo", "yyyy-MM-dd");
    var holidayType = holidaysMap[formattedDate];
    var isCompanyHoliday = (holidayType && (holidayType === 'company' || holidayType === 'holiday')); 
    
    // 会社の休日の場合、または土曜日に適用 (リクエスト: "会社の休日の場合は、土曜日にカレンダーと同じ色で着色")
    // 土曜日 または 会社休日 の場合に着色
    if (isSaturday || isCompanyHoliday) {
        var range = sheet.getRange(row, 1, 1, sheet.getLastColumn());
        range.setBackground('#dbeafe'); // Light Blue
        range.setFontColor('#2563eb');  // Blue Text
    } else {
        // 平日の場合
        var range = sheet.getRange(row, 1, 1, sheet.getLastColumn());
        range.setBackground(null); 
        range.setFontColor(null);
    }
    
    // 日曜日
    if (dateObj.getDay() === 0) {
        var range = sheet.getRange(row, 1, 1, sheet.getLastColumn());
        range.setBackground('#fee2e2'); // Light Red
        range.setFontColor('#dc2626');  // Red Text
    }
}



// ===== 休日処理用ヘルパー関数 =====

// 休日設定を一括取得するヘルパー関数
function getAllHolidaysMap(yearMonth) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('設定_休日');
  if (!sheet) return {};
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return {};
  
  var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var holidays = {};
  var targetYear = yearMonth ? yearMonth.split('-')[0] : null;

  for (var i = 0; i < values.length; i++) {
    var d = values[i][0];
    if (!d) continue;
    
    // 日付オブジェクトかどうか確認
    if (!(d instanceof Date)) {
      d = new Date(d);
    }
    
    // 年月が指定されているなら、その年のものだけ取得（パフォーマンス用）--> 一旦無効化（トラブルシュート）
    // if (targetYear && String(d.getFullYear()) !== String(targetYear)) continue;

    var dateStr = Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM-dd");
    holidays[dateStr] = values[i][1];
  }
  return holidays;
}

// 休日判定関数
// 指定された日付が休日（土日または設定された休日）かどうかを判定する
function isHoliday(date, holidays) {
  var dateStr = Utilities.formatDate(date, "Asia/Tokyo", "yyyy-MM-dd");
  
  // 1. 設定された休日かどうか
  // holidaysマップがあればそれを優先
  if (holidays && holidays[dateStr]) {
    // holidays[dateStr] には 'holiday', 'company' などのタイプが入っている想定
    return true; 
  }
  
  // 2. 土日判定
  var day = date.getDay();
  if (day === 0 || day === 6) {
    return true;
  }
  
  return false;
}

// 休日設定を取得する関数
function getHolidaySettings(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var jsonData = JSON.parse(e.postData.contents);
    var year = jsonData.year; // 数値または文字列の年
    
    var sheet = ss.getSheetByName('設定_休日');
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'success',
        holidays: {} // シートがなければ空で返す
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'success',
        holidays: {}
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // A列:日付, B:種別
    var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    var holidays = {};
    
    for (var i = 0; i < values.length; i++) {
      var dateVal = values[i][0];
      var type = values[i][1];
      
      if (!dateVal) continue;
      
      var dateObj = new Date(dateVal);
      // 指定年のデータのみ抽出
      if (dateObj.getFullYear() == year) {
        var dateStr = Utilities.formatDate(dateObj, "Asia/Tokyo", "yyyy-MM-dd");
        holidays[dateStr] = type;
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      result: 'success',
      holidays: holidays
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 休日設定を保存する関数
function saveHolidaySettings(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var jsonData = JSON.parse(e.postData.contents);
    var year = jsonData.year;
    var holidays = jsonData.holidays; // { "2024-01-01": "holiday", ... }
    
    var sheet = ss.getSheetByName('設定_休日');
    if (!sheet) {
      sheet = ss.insertSheet('設定_休日');
      sheet.appendRow(['日付', '種別', '名称']); // ヘッダー
    }
    
    var lastRow = sheet.getLastRow();
    var newRows = [];
    
    // 既存データから、指定年以外のデータを保持
    if (lastRow > 1) {
      var existingData = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
      for (var i = 0; i < existingData.length; i++) {
        var val = existingData[i][0];
        if (!val) continue; // Skip empty rows
        
        var d;
        if (val instanceof Date) {
            d = val;
        } else {
            d = new Date(val);
        }

        if (isNaN(d.getTime())) continue; // Skip invalid dates

        // 指定年以外のデータのみ保持
        if (d.getFullYear() != year) {
           newRows.push(existingData[i]);
        }
      }
    }
    
    // 新しいデータを追加
    for (var dateStr in holidays) {
      var type = holidays[dateStr];
      // 日付文字列からDateオブジェクト作成 (YYYY-MM-DD -> YYYY/MM/DD)
      // ローカル日付としてパースされるようにする
      var dateObj = new Date(dateStr.replace(/-/g, '/'));
      
      // 有効な日付かチェック
      if (!isNaN(dateObj.getTime())) {
        var name = (type === 'holiday') ? '祝日' : '会社休日';
        newRows.push([dateObj, type, name]);
      }
    }
    
    // ソート（日付順）
    newRows.sort(function(a, b) {
      return new Date(a[0]) - new Date(b[0]);
    });
    
    // シートをクリアして全書き換え（ヘッダー以外）
    if (lastRow > 1) {
      // データが存在した範囲をクリア
      sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
    }
    
    if (newRows.length > 0) {
      sheet.getRange(2, 1, newRows.length, 3).setValues(newRows);
      
      // A列の表示形式を YYYY/MM/DD に設定
      sheet.getRange(2, 1, newRows.length, 1).setNumberFormat('yyyy/mm/dd');
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      result: 'success',
      message: '保存しました'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ===== 月次確認・承認機能 =====

// GETリクエストの処理(HTMLページの表示用)
function doGet(e) {
  var page = e.parameter.page || 'attendance';
  
  if (page === 'monthly-review') {
    return HtmlService.createHtmlOutputFromFile('MonthlyReview')
      .setTitle('月次勤怠確認・承認')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (page === 'manager-dashboard') {
    return HtmlService.createHtmlOutputFromFile('ManagerDashboard')
      .setTitle('管理者ダッシュボード')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (page === 'holiday-settings') {
    return HtmlService.createHtmlOutputFromFile('HolidaySettings')
      .setTitle('年間休日設定')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  // デフォルトは打刻画面
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('勤怠管理')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 月次データを取得する関数
function getMonthlyData(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var jsonData = JSON.parse(e.postData.contents);
    var department = jsonData.department;
    var yearMonth = jsonData.yearMonth; // 'YYYY-MM' 形式
    
    // 部署別の打刻シートを取得
    var sheetName = '打刻_' + department;
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'error',
        message: '部署「' + department + '」のシートが見つかりません'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // 曜日列確認 (Migrate if needed)
    ensureDayOfWeekColumn(sheet);
    
    // データを取得
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'success',
        data: [],
        approvalStatus: getApprovalStatus(department, yearMonth)
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Read up to Col J (Remarks) = 10 columns
    var values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    var monthlyData = [];
    
    // 指定された年月のデータのみフィルタリング
    for (var i = 0; i < values.length; i++) {
      var dateValue = values[i][0];
      var dateStr = '';
      
      if (dateValue instanceof Date) {
        dateStr = Utilities.formatDate(dateValue, "Asia/Tokyo", "yyyy-MM");
      } else if (dateValue) {
        dateStr = String(dateValue).substring(0, 7); // 'YYYY-MM' 部分を取得
      }
      
      if (dateStr === yearMonth) {
        // Indices shifted by 1 due to DayOfWeek at index 1
        monthlyData.push({
          date: dateValue instanceof Date ? Utilities.formatDate(dateValue, "Asia/Tokyo", "yyyy/MM/dd") : dateValue,
          dayOfWeek: values[i][1], // B列
          employeeId: values[i][2], // C列
          name: values[i][3], // D列
          clockInType: values[i][4], // E列
          clockInTime: values[i][5], // F列
          clockOutType: values[i][6], // G列
          clockOutTime: values[i][7], // H列
          workingHours: values[i][8], // I列
          remarks: values[i][9] // J列
        });
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      result: 'success',
      data: monthlyData,
      approvalStatus: getApprovalStatus(department, yearMonth)
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 承認状態を取得する関数（個人単位）
function getApprovalStatus(employeeId, yearMonth) {
  console.log('getApprovalStatus呼び出し - employeeId:', employeeId, 'yearMonth:', yearMonth);
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var approvalSheet = ss.getSheetByName('月次承認記録_個人');
  
  if (!approvalSheet) {
    console.log('承認シートが存在しません');
    return {
      self: null, // 本人確認
      boss: null  // 上長承認
    };
  }
  
  var lastRow = approvalSheet.getLastRow();
  console.log('承認シート最終行:', lastRow);
  
  if (lastRow <= 1) {
    console.log('承認シートにデータがありません');
    return { self: null, boss: null };
  }
  
  var values = approvalSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  console.log('承認シートから読み取った行数:', values.length);
  
  // A:年月, B:社員ID, C:本人確認者, D:本人確認日時, E:承認者, F:承認日時
  
  for (var i = values.length - 1; i >= 0; i--) {
    var rowYearMonth = values[i][0];
    var rowEmployeeId = String(values[i][1]);
    var rowSelf = values[i][2];
    var rowSelfDate = values[i][3];
    var rowBoss = values[i][4];
    var rowBossDate = values[i][5];
    
    console.log('行' + (i+2) + '確認 - 年月:', rowYearMonth, '社員ID:', rowEmployeeId, '本人:', rowSelf, '承認者:', rowBoss);
    
    if (String(rowEmployeeId) === String(employeeId) && rowYearMonth === yearMonth) {
      console.log('✓ マッチング成功! - 社員ID:', employeeId, '年月:', yearMonth);
      var result = {
        self: rowSelf ? { name: rowSelf, date: rowSelfDate } : null,
        boss: rowBoss ? { name: rowBoss, date: rowBossDate } : null
      };
      console.log('返却する承認状態:', result);
      return result;
    }
  }
  
  console.log('✗ 該当データが見つかりませんでした - 社員ID:', employeeId, '年月:', yearMonth);
  return { self: null, boss: null };
}

// 承認を記録する関数（個人単位）
function recordApproval(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var debugSheet = ss.getSheetByName('デバッグログ');
    var jsonData = JSON.parse(e.postData.contents);
    
    if (debugSheet) {
      debugSheet.appendRow([new Date(), 'recordApproval開始: ' + JSON.stringify(jsonData)]);
    }
    
    var targetEmployeeId = jsonData.targetEmployeeId; // 承認対象の社員ID
    var yearMonth = jsonData.yearMonth;
    var approverId = jsonData.approverId; // 操作している人のID
    var type = jsonData.type; // 'self' (本人確認) or 'boss' (承認)
    
    // 1. 操作している人の情報を取得（印鑑用）
    var masterSheet = ss.getSheetByName('社員マスタ');
    if (!masterSheet) {
      if (debugSheet) debugSheet.appendRow([new Date(), 'エラー: 社員マスタシートが存在しません']);
      return ContentService.createTextOutput(JSON.stringify({
        result: 'error',
        message: '社員マスタシートが存在しません。先に社員マスタを作成してください。'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var approverInfo = getEmployeeInfo(masterSheet, approverId);
    
    if (debugSheet) {
      debugSheet.appendRow([new Date(), '承認者情報: ' + (approverInfo ? JSON.stringify(approverInfo) : 'null')]);
    }
    
    if (!approverInfo) {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'error',
        message: '操作者(社員コード: ' + approverId + ')の情報がマスタに見つかりません。社員マスタに登録してください。'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // 印鑑用テキスト（マスタに印鑑列があればそれを使う、なければ氏名）
    var stampName = approverInfo.stampName || approverInfo.name;
    
    if (debugSheet) {
      debugSheet.appendRow([new Date(), '印鑑名: ' + stampName]);
    }
    
    // 2. 権限チェック
    if (type === 'self') {
      // 本人確認は本人のみ
      if (String(targetEmployeeId) !== String(approverId)) {
        if (debugSheet) {
          debugSheet.appendRow([new Date(), 'エラー: 本人以外の担当者印 - target:' + targetEmployeeId + ', approver:' + approverId]);
        }
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error',
          message: '本人以外のデータに担当者印は押せません'
        })).setMimeType(ContentService.MimeType.JSON);
      }
    } else if (type === 'boss') {
      // 承認は、対象者の「第1承認者」として登録されている人のみ
      var targetInfo = getEmployeeInfo(masterSheet, targetEmployeeId);
      
      if (debugSheet) {
        debugSheet.appendRow([new Date(), '対象者情報: ' + (targetInfo ? JSON.stringify(targetInfo) : 'null')]);
      }
      
      if (!targetInfo) {
         return ContentService.createTextOutput(JSON.stringify({
          result: 'error',
          message: '対象者(社員コード: ' + targetEmployeeId + ')の情報が見つかりません'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      // マスタの第1承認者名と、操作者の氏名が一致するか確認
      // ※より厳密にはIDで紐付けるべきですが、現状のマスタ構造に合わせて名前一致で判定
      if (debugSheet) {
        debugSheet.appendRow([new Date(), '第1承認者チェック - 対象者の第1承認者: "' + targetInfo.firstApprover + '", 操作者名: "' + approverInfo.name + '"']);
      }
      
      if (targetInfo.firstApprover !== approverInfo.name) {
         return ContentService.createTextOutput(JSON.stringify({
          result: 'error',
          message: 'あなたはこの社員の承認権限を持っていません。社員マスタのE列(第1承認者)を確認してください。\n対象者の第1承認者: "' + targetInfo.firstApprover + '", あなたの名前: "' + approverInfo.name + '"'
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    // 3. 記録
    var approvalSheet = ss.getSheetByName('月次承認記録_個人');
    if (!approvalSheet) {
      approvalSheet = ss.insertSheet('月次承認記録_個人');
      approvalSheet.appendRow(['年月', '社員ID', '本人確認者', '本人確認日時', '承認者', '承認日時']);
      if (debugSheet) debugSheet.appendRow([new Date(), '月次承認記録_個人シート新規作成']);
    }
    
    var lastRow = approvalSheet.getLastRow();
    var foundRow = -1;
    
    // 既存レコード検索
    if (lastRow > 1) {
      var values = approvalSheet.getRange(2, 1, lastRow - 1, 2).getValues();
      for (var i = 0; i < values.length; i++) {
        if (values[i][0] === yearMonth && String(values[i][1]) === String(targetEmployeeId)) {
          foundRow = i + 2;
          break;
        }
      }
    }
    
    if (debugSheet) {
      debugSheet.appendRow([new Date(), '既存レコード検索結果: ' + (foundRow > 0 ? foundRow + '行目' : '見つからず')]);
    }
    
    var now = new Date();
    
    if (foundRow > 0) {
      // 更新
      if (type === 'self') {
        approvalSheet.getRange(foundRow, 3).setValue(stampName);
        approvalSheet.getRange(foundRow, 4).setValue(now);
        if (debugSheet) debugSheet.appendRow([new Date(), '担当者印更新: ' + foundRow + '行目']);
      } else if (type === 'boss') {
        // 本人確認がまだならエラー
        var selfCheck = approvalSheet.getRange(foundRow, 3).getValue();
        if (!selfCheck) {
          if (debugSheet) debugSheet.appendRow([new Date(), 'エラー: 本人確認未完了']);
          return ContentService.createTextOutput(JSON.stringify({
            result: 'error',
            message: '本人確認が完了していません'
          })).setMimeType(ContentService.MimeType.JSON);
        }
        approvalSheet.getRange(foundRow, 5).setValue(stampName);
        approvalSheet.getRange(foundRow, 6).setValue(now);
        if (debugSheet) debugSheet.appendRow([new Date(), '承認印更新: ' + foundRow + '行目']);
      }
    } else {
      // 新規
      if (type === 'boss') {
        if (debugSheet) debugSheet.appendRow([new Date(), 'エラー: 新規レコードに承認印は押せません']);
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error',
          message: '本人確認が完了していません'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      approvalSheet.appendRow([
        yearMonth,
        targetEmployeeId,
        stampName,
        now,
        '',
        ''
      ]);
      if (debugSheet) debugSheet.appendRow([new Date(), '新規レコード追加']);
    }
    
    
    console.log('recordApproval成功完了 - type:', type, 'targetEmployeeId:', targetEmployeeId, 'yearMonth:', yearMonth);
    console.log('書き込んだ行:', foundRow > 0 ? foundRow : '新規行');
    
    if (debugSheet) {
      debugSheet.appendRow([new Date(), 'recordApproval成功完了']);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      result: 'success',
      message: (type === 'self' ? '担当者印' : '承認印') + 'を押しました',
      stampData: { name: stampName, date: now }
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var debugSheet = ss.getSheetByName('デバッグログ');
      if (debugSheet) {
        debugSheet.appendRow([new Date(), 'recordApprovalエラー: ' + error.toString()]);
      }
    } catch (e) {}
    
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error',
      message: 'システムエラー: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ヘルパー: 社員情報を取得
function getEmployeeInfo(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;
  
  var values = sheet.getRange(2, 1, lastRow - 1, 8).getValues(); // H列(印鑑)まで取得想定
  
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      return {
        id: values[i][0],
        name: values[i][1],
        department: values[i][2],
        firstApprover: values[i][4],
        stampName: values[i][7] // H列を印鑑データと仮定
      };
    }
  }
  return null;
}

// 承認者の権限を確認する関数（今回はrecordApproval内で処理するためダミー化または削除）
function checkApproverPermission(department, approverName, approvalLevel) {
  return false; 
}


// 部署一覧を取得する関数
function getDepartments() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName('社員マスタ');
    
    if (!masterSheet) {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'error',
        message: '社員マスタが見つかりません'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var lastRow = masterSheet.getLastRow();
    if (lastRow <= 1) {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'success',
        departments: []
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var values = masterSheet.getRange(2, 3, lastRow - 1, 1).getValues();
    var departments = [];
    var uniqueDepts = {};
    
    for (var i = 0; i < values.length; i++) {
      var dept = values[i][0];
      if (dept && !uniqueDepts[dept]) {
        uniqueDepts[dept] = true;
        departments.push(dept);
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      result: 'success',
      departments: departments
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 部署と年月を指定して社員一覧と打刻状況を取得
function getEmployeesByDepartmentAndMonth(e) {
  try {
    var jsonData = JSON.parse(e.postData.contents);
    var department = jsonData.department;
    var yearMonth = jsonData.yearMonth; // "2024-11" 形式
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName('社員マスタ');
    
    if (!masterSheet) {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'error',
        message: '社員マスタが見つかりません'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // 部署の社員一覧を取得
    var lastRow = masterSheet.getLastRow();
    if (lastRow <= 1) {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'success',
        employees: []
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var values = masterSheet.getRange(2, 1, lastRow - 1, 3).getValues();
    var employees = [];
    
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][2]) === String(department)) {
        employees.push({
          employeeId: String(values[i][0]),
          name: String(values[i][1])
        });
      }
    }
    
    // 打刻シートから各社員の打刻状況をチェック
    var departmentSheetName = '打刻_' + department;
    var logSheet = ss.getSheetByName(departmentSheetName);
    
    if (!logSheet) {
      // 打刻シートがない場合は全員×
      for (var i = 0; i < employees.length; i++) {
        employees[i].attendanceStatus = '×';
        employees[i].approved = false;
      }
    } else {
      ensureDayOfWeekColumn(logSheet);
      
      // 年月から営業日数を計算（簡易版：土日を除く＋設定された休日を除く）
      var year = parseInt(yearMonth.split('-')[0]);
      var month = parseInt(yearMonth.split('-')[1]);
      var daysInMonth = new Date(year, month, 0).getDate();
      
      // 休日マップを取得
      var holidaysMap = getAllHolidaysMap(yearMonth);
      
      var workDays = 0;
      for (var day = 1; day <= daysInMonth; day++) {
        var date = new Date(year, month - 1, day);
        if (!isHoliday(date, holidaysMap)) {
          workDays++;
        }
      }
      
      // 打刻データを取得
      var logLastRow = logSheet.getLastRow();
      if (logLastRow > 1) {
        // Read up to K (Approval) = 11 columns
        var logValues = logSheet.getRange(2, 1, logLastRow - 1, 11).getValues();
        
        for (var i = 0; i < employees.length; i++) {
          var employeeId = employees[i].employeeId;
          var attendanceDays = {};
          var approved = false;
          
          // この社員の打刻データを集計
          for (var j = 0; j < logValues.length; j++) {
            var logDate = logValues[j][0];
            var logId = String(logValues[j][2]); // C列 (Index 2)
            var logApproval = logValues[j][10]; // K列 (Index 10)
            
            // 日付を年月でフィルタ
            var formattedDate = '';
            if (logDate instanceof Date) {
              formattedDate = Utilities.formatDate(logDate, "Asia/Tokyo", "yyyy/MM");
            } else if (typeof logDate === 'string') {
              formattedDate = logDate.substring(0, 7);
            }
            
            if (formattedDate === yearMonth && logId === employeeId) {
              var dateKey = '';
              if (logDate instanceof Date) {
                dateKey = Utilities.formatDate(logDate, "Asia/Tokyo", "yyyy/MM/dd");
              } else {
                dateKey = String(logDate);
              }
              attendanceDays[dateKey] = true;
              
              // 承認フラグをチェック
              if (logApproval === '○') {
                approved = true;
              }
            }
          }
          
          // 打刻日数と営業日数を比較
          var attendanceCount = Object.keys(attendanceDays).length;
          employees[i].attendanceStatus = (attendanceCount >= workDays) ? '○' : '×';
          employees[i].approved = approved;
          employees[i].attendanceDays = attendanceCount;
          employees[i].workDays = workDays;
        }
      } else {
        // データがない場合は全員×
        for (var i = 0; i < employees.length; i++) {
          employees[i].attendanceStatus = '×';
          employees[i].approved = false;
        }
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      result: 'success',
      employees: employees
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 社員の承認を記録（部署シートのK列に○を記録）
function approveEmployee(e) {
  try {
    var jsonData = JSON.parse(e.postData.contents);
    var department = jsonData.department;
    var employeeId = jsonData.employeeId;
    var yearMonth = jsonData.yearMonth; // "2024-11" 形式
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var departmentSheetName = '打刻_' + department;
    var logSheet = ss.getSheetByName(departmentSheetName);
    
    // デバッグログ
    var debugSheet = ss.getSheetByName('デバッグログ');
    if (debugSheet) {
      debugSheet.appendRow([new Date(), '承認処理開始: 部署=' + department + ', 社員ID=' + employeeId + ', 年月=' + yearMonth]);
    }
    
    if (!logSheet) {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'error',
        message: '部署シート「' + departmentSheetName + '」が見つかりません'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    ensureDayOfWeekColumn(logSheet);
    
    // K列に承認フラグを追加（ヘッダーがない場合は追加）
    var headerRow = logSheet.getRange(1, 1, 1, 11).getValues()[0];
    if (!headerRow[10]) { // Index 10 is K
      logSheet.getRange(1, 11).setValue('承認');
      if (debugSheet) {
        debugSheet.appendRow([new Date(), 'K列に「承認」ヘッダーを追加しました']);
      }
    }
    
    // 該当社員の該当月のデータを検索してK列に○を記録
    var lastRow = logSheet.getLastRow();
    if (lastRow > 1) {
      // Read A to C (Date, Day, ID)
      var values = logSheet.getRange(2, 1, lastRow - 1, 3).getValues();
      var updatedCount = 0;
      
      // 年月を正規化（"2024-11" → "2024/11" の両方に対応）
      var targetYearMonth1 = yearMonth; // "2024-11"
      var targetYearMonth2 = yearMonth.replace('-', '/'); // "2024/11"
      
      if (debugSheet) {
        debugSheet.appendRow([new Date(), '検索対象年月: ' + targetYearMonth1 + ' または ' + targetYearMonth2]);
        debugSheet.appendRow([new Date(), '検索範囲: ' + (lastRow - 1) + '行']);
      }
      
      for (var i = 0; i < values.length; i++) {
        var logDate = values[i][0];
        var logId = String(values[i][2]); // C列 (Index 2)
        
        // 日付を年月でフィルタ
        var formattedDate = '';
        if (logDate instanceof Date) {
          formattedDate = Utilities.formatDate(logDate, "Asia/Tokyo", "yyyy/MM");
        } else if (typeof logDate === 'string') {
          // 文字列の場合は最初の7文字を取得（"2024/11/01" → "2024/11"）
          formattedDate = String(logDate).substring(0, 7);
        }
        
        // 年月と社員IDが一致するかチェック
        var dateMatches = (formattedDate === targetYearMonth1 || formattedDate === targetYearMonth2);
        var idMatches = (logId === employeeId);
        
        if (dateMatches && idMatches) {
          logSheet.getRange(i + 2, 11).setValue('○'); // K列
          updatedCount++;
          
          if (debugSheet) {
            debugSheet.appendRow([new Date(), '承認記録: 行' + (i + 2) + ', 日付=' + formattedDate + ', ID=' + logId]);
          }
        }
      }
      
      if (debugSheet) {
        debugSheet.appendRow([new Date(), '承認処理完了: ' + updatedCount + '件更新']);
      }
      
      return ContentService.createTextOutput(JSON.stringify({
        result: 'success',
        message: updatedCount + '件の打刻データに承認フラグを記録しました',
        updatedCount: updatedCount
      })).setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'error',
        message: '打刻データが見つかりません（シートにデータ行がありません）'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
  } catch (error) {
    // エラーログ
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var debugSheet = ss.getSheetByName('デバッグログ');
      if (debugSheet) {
        debugSheet.appendRow([new Date(), '承認処理エラー: ' + error.toString()]);
      }
    } catch (e) {}
    
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 個人の月次打刻データを取得する関数
function getPersonalMonthlyData(e) {
  try {
    var jsonData = JSON.parse(e.postData.contents);
    var employeeId = jsonData.employeeId;
    var yearMonth = jsonData.yearMonth; // "2024-11" 形式
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. 社員マスタから部署を特定
    var masterSheet = ss.getSheetByName('社員マスタ');
    if (!masterSheet) {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'error',
        message: '社員マスタが見つかりません'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var department = '';
    var lastRow = masterSheet.getLastRow();
    if (lastRow > 1) {
      var values = masterSheet.getRange(2, 1, lastRow - 1, 3).getValues();
      for (var i = 0; i < values.length; i++) {
        if (String(values[i][0]) === String(employeeId)) {
          department = values[i][2];
          break;
        }
      }
    }
    
    if (!department) {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'error',
        message: '社員コードに対応する部署が見つかりません'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // 2. 部署の打刻シートからデータを取得
    var departmentSheetName = '打刻_' + department;
    var logSheet = ss.getSheetByName(departmentSheetName);
    
    if (!logSheet) {
      // シートがない場合はデータなしとして返す
      return ContentService.createTextOutput(JSON.stringify({
        result: 'success',
        data: {}
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    ensureDayOfWeekColumn(logSheet);
    
    var logLastRow = logSheet.getLastRow();
    if (logLastRow <= 1) {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'success',
        data: {}
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // L列(位置情報)まで取得するよう範囲を拡張
    var maxCols = logSheet.getLastColumn();
    var fetchCols = Math.max(12, maxCols); // 少なくとも12列目までは取得
    
    var logValues = logSheet.getRange(2, 1, logLastRow - 1, fetchCols).getValues(); 
    var personalData = {};
    
    // 年月を正規化（"2024-11" → "2024/11" の両方に対応）
    var targetYearMonth1 = yearMonth; // "2024-11"
    var targetYearMonth2 = yearMonth.replace('-', '/'); // "2024/11"
    
    for (var i = 0; i < logValues.length; i++) {
      var logDate = logValues[i][0];
      var logId = String(logValues[i][2]); // C列 (Index 2)
      
      // 社員IDチェック
      if (logId !== String(employeeId)) continue;
      
      // 日付チェック
      var formattedDate = '';
      var dateKey = '';
      
      if (logDate instanceof Date) {
        formattedDate = Utilities.formatDate(logDate, "Asia/Tokyo", "yyyy/MM");
        dateKey = Utilities.formatDate(logDate, "Asia/Tokyo", "yyyy/MM/dd");
      } else if (typeof logDate === 'string') {
        formattedDate = logDate.substring(0, 7);
        dateKey = logDate;
      }
      
      if (formattedDate === targetYearMonth1 || formattedDate === targetYearMonth2) {
        // 時刻データのフォーマット処理
        var formatTime = function(timeVal) {
          if (!timeVal) return '';
          if (timeVal instanceof Date) {
            return Utilities.formatDate(timeVal, "Asia/Tokyo", "HH:mm");
          }
          return String(timeVal);
        };

        // 位置情報のパース
        var locationLog = [];
        var kVal = logValues[i][11]; // index 11 = L列
        if (kVal) {
          try {
             if (typeof kVal === 'string' && (kVal.startsWith('[') || kVal.startsWith('{'))) {
               var parsed = JSON.parse(kVal);
               locationLog = Array.isArray(parsed) ? parsed : [parsed];
             }
          } catch(e) {
             // パースエラーは無視
          }
        }

        // 日付をキーにしてデータを格納
        // Indices: E:4, F:5, G:6, H:7, I:8, J:9
        personalData[dateKey] = {
          clockInType: logValues[i][4],
          clockInTime: formatTime(logValues[i][5]),
          clockOutType: logValues[i][6],
          clockOutTime: formatTime(logValues[i][7]),
          workingHours: logValues[i][8],
          remarks: logValues[i][9],
          locationLog: locationLog // 位置情報ログを追加
        };
      }
    }
    
    // 承認状態を取得
    var approvalStatus = getApprovalStatus(employeeId, yearMonth);

    // 休日情報を取得
    var holidays = getAllHolidaysMap(yearMonth);
    
    return ContentService.createTextOutput(JSON.stringify({
      result: 'success',
      data: personalData,
      approvalStatus: approvalStatus,
      holidays: holidays
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}



// ===== リマインダー通知機能 =====

// 打刻忘れチェックとメール送信を行う関数
// トリガーで定期実行することを想定（例: 10:00 と 19:00）
function checkAndSendReminders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName('社員マスタ');
  
  if (!masterSheet) {
    console.log('社員マスタが見つかりません');
    return;
  }

  // 現在時刻と日付情報の取得
  var now = new Date();
  var todayStr = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd");
  var hour = now.getHours();
  
  // 午前(12時前)なら出勤チェック、午後なら退勤チェック
  var isMorningCheck = (hour < 12);
  var checkType = isMorningCheck ? '出勤' : '退勤';
  
  console.log('リマインダーチェック開始: ' + todayStr + ' (' + checkType + '確認)');

  // 1. 土日チェック
  var dayOfWeek = now.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    console.log('土日のためリマインダーをスキップします');
    return;
  }

  // 2. 祝日チェック (Googleカレンダーの日本の祝日を使用)
  try {
    var calendarId = 'ja.japanese#holiday@group.v.calendar.google.com';
    var calendar = CalendarApp.getCalendarById(calendarId);
    var events = calendar.getEventsForDay(now);
    if (events.length > 0) {
      console.log('祝日のためリマインダーをスキップします: ' + events[0].getTitle());
      return;
    }
  } catch (e) {
    console.log('祝日チェックに失敗しました（処理は続行します）: ' + e.toString());
  }

  // 3. 社員マスタ取得
  // A:コード, B:氏名, C:部署, ..., G:メールアドレス(7列目)
  var lastRow = masterSheet.getLastRow();
  if (lastRow <= 1) return;
  
  var employees = masterSheet.getRange(2, 1, lastRow - 1, 7).getValues();
  
  // 4. 各社員の打刻状況をチェック
  employees.forEach(function(emp) {
    var empId = emp[0];
    var empName = emp[1];
    var department = emp[2];
    var email = emp[6]; // G列: メールアドレス
    
    // メールアドレスがない、または部署未設定の場合はスキップ
    if (!email || !department || department === '未設定') return;
    
    var logSheetName = '打刻_' + department;
    var logSheet = ss.getSheetByName(logSheetName);
    var isClockedIn = false;
    var isClockedOut = false;
    
    if (logSheet) {
      // 曜日列確認 (Migrate if needed, though likely not needed for read if we handle indices dynamically, but better safe)
      ensureDayOfWeekColumn(logSheet);
      
      // 今日のデータを検索（直近50行を確認）
      var logLastRow = logSheet.getLastRow();
      if (logLastRow > 1) {
        var startRow = Math.max(2, logLastRow - 50);
        var numRows = logLastRow - startRow + 1;
        // Read up to H (OutTime+1) or just use 8 columns. A to H.
        var logs = logSheet.getRange(startRow, 1, numRows, 8).getValues(); // H列まで
        
        for (var i = logs.length - 1; i >= 0; i--) {
          var logDate = logs[i][0]; // A列
          var logId = logs[i][2];   // C列 (Index 2)
          
          var logDateStr = '';
          if (logDate instanceof Date) {
            logDateStr = Utilities.formatDate(logDate, "Asia/Tokyo", "yyyy/MM/dd");
          } else {
            logDateStr = String(logDate);
          }
          
          // 日付とIDが一致するか
          if (logDateStr === todayStr && String(logId) === String(empId)) {
            if (logs[i][5]) isClockedIn = true; // F列: 出勤時刻 (Index 5)
            if (logs[i][7]) isClockedOut = true; // H列: 退勤時刻 (Index 7)
            break;
          }
        }
      }
    }
    
    // 5. 条件に応じてメール送信
    if (isMorningCheck) {
      // 出勤チェック: まだ出勤していない場合に通知
      if (!isClockedIn) {
        sendReminderEmail(email, empName, 'in');
      }
    } else {
      // 退勤チェック: 出勤しているが、退勤していない場合に通知
      if (isClockedIn && !isClockedOut) {
        sendReminderEmail(email, empName, 'out');
      }
    }
  });
}

// メール送信ヘルパー関数
function sendReminderEmail(to, name, type) {
  var subject = '';
  var body = '';
  
  if (type === 'in') {
    subject = '【勤怠連絡】出勤打刻の確認';
    body = name + ' さん\n\n' +
           'おはようございます。\n' +
           '本日の出勤打刻が確認できていません。\n' +
           '業務を開始されている場合は、打刻をお願いします。\n\n' +
           '※休暇等の場合はご放念ください。\n' +
           '※このメールはシステムより自動送信されています。';
  } else {
    subject = '【勤怠連絡】退勤打刻の確認';
    body = name + ' さん\n\n' +
           'お疲れ様です。\n' +
           '本日の退勤打刻が確認できていません。\n' +
           '業務を終了されている場合は、打刻をお願いします。\n\n' +
           '※残業中の場合は、業務終了後に打刻をお願いします。\n' +
           '※このメールはシステムより自動送信されています。';
  }
  
  try {
    MailApp.sendEmail({
      to: to,
      subject: subject,
      body: body
    });
    console.log('メール送信成功: ' + name + ' (' + to + ')');
  } catch (e) {
    console.log('メール送信失敗: ' + name + ' - ' + e.toString());
  }
}

// 承認者ダッシュボード用データ取得
function getApproverDashboard(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var jsonData = JSON.parse(e.postData.contents);
    var approverId = jsonData.approverId;
    var yearMonth = jsonData.yearMonth; // YYYY-MM
    
    // 1. 承認者の情報を取得
    var masterSheet = ss.getSheetByName('社員マスタ');
    var approverInfo = getEmployeeInfo(masterSheet, approverId);
    
    if (!approverInfo) {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'error',
        message: '承認者の情報が見つかりません'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var approverName = approverInfo.name;
    
    // 2. 部下（承認対象者）を検索
    // マスタ構造: A:コード, B:氏名, C:部署, D:承認対象部署, E:第1承認者
    var lastRow = masterSheet.getLastRow();
    var subordinates = [];
    
    if (lastRow > 1) {
      var values = masterSheet.getRange(2, 1, lastRow - 1, 5).getValues();
      for (var i = 0; i < values.length; i++) {
        // E列(index 4)が承認者名と一致するか
        if (values[i][4] === approverName) {
          subordinates.push({
            id: values[i][0],
            name: values[i][1],
            department: values[i][2]
          });
        }
      }
    }
    
    if (subordinates.length === 0) {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'success',
        isApprover: false,
        subordinates: []
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // 3. 各部下の勤怠状況をチェック
    // 月の営業日数を計算（土日除く＋設定された休日除く）
    var parts = yearMonth.split('-');
    var year = parseInt(parts[0]);
    var month = parseInt(parts[1]);
    var daysInMonth = new Date(year, month, 0).getDate();
    
    // 休日マップを取得
    var holidaysMap = getAllHolidaysMap(yearMonth);
    
    var workDays = 0;
    for (var d = 1; d <= daysInMonth; d++) {
      var dateObj = new Date(year, month - 1, d);
      if (!isHoliday(dateObj, holidaysMap)) {
        workDays++;
      }
    }
    
    
    var approvalSheet = ss.getSheetByName('月次承認記録_個人');
    var approvalData = [];
    if (approvalSheet && approvalSheet.getLastRow() > 1) {
      approvalData = approvalSheet.getRange(2, 1, approvalSheet.getLastRow() - 1, 6).getValues();
      console.log('承認シートから読み取ったデータ行数:', approvalData.length);
      console.log('対象年月:', yearMonth);
      // デバッグ: 最初の数行を出力
      for (var debugIdx = 0; debugIdx < Math.min(5, approvalData.length); debugIdx++) {
        console.log('承認データ[' + debugIdx + ']:', approvalData[debugIdx]);
      }
    } else {
      console.log('承認シートが存在しないか、データがありません');
    }
    
    // 各部下のデータ処理
    for (var i = 0; i < subordinates.length; i++) {
      var sub = subordinates[i];
      var sheetName = '打刻_' + sub.department;
      var logSheet = ss.getSheetByName(sheetName);
      
      var attendanceCount = 0;
      var incompleteDays = []; // 退勤未打刻の日付リスト
      
      if (logSheet && logSheet.getLastRow() > 1) {
        ensureDayOfWeekColumn(logSheet);
        
        // Date(A), Day(B), ID(C), Name(D), InType(E), InTime(F), OutType(G), OutTime(H)
        var logs = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 8).getValues();
        var countedDays = {};
        
        for (var j = 0; j < logs.length; j++) {
           var logDate = logs[j][0];
           var logId = String(logs[j][2]); // C列 (Index 2)
           var clockInTime = logs[j][5]; // F列
           var clockOutTime = logs[j][7]; // H列
           
           if (String(logId) === String(sub.id)) {
             var logYM = '';
             if (logDate instanceof Date) {
               logYM = Utilities.formatDate(logDate, "Asia/Tokyo", "yyyy-MM");
             } else {
               logYM = String(logDate).substring(0, 7);
             }
             
             if (logYM === yearMonth) {
               var dateKey = '';
               if (logDate instanceof Date) {
                  dateKey = Utilities.formatDate(logDate, "Asia/Tokyo", "yyyy/MM/dd");
               } else {
                  dateKey = String(logDate);
               }
               
               if (!countedDays[dateKey]) {
                 countedDays[dateKey] = true;
                 attendanceCount++;
                 
                 // 出勤はあるが退勤がない場合
                 if (clockInTime && !clockOutTime) {
                   incompleteDays.push(dateKey);
                 }
               }
             }
           }
        }
      }
      
      // 承認状態の確認
      var status = { self: false, boss: false };
      var foundApproval = false;
      for (var k = 0; k < approvalData.length; k++) {
        // A:年月, B:社員ID, C:本人, E:承認者
        var dataYearMonth = approvalData[k][0];
        var dataEmployeeId = String(approvalData[k][1]);
        var dataSelf = approvalData[k][2];
        var dataBoss = approvalData[k][4];
        
        if (dataYearMonth === yearMonth && dataEmployeeId === String(sub.id)) {
          if (dataSelf) status.self = true;
          if (dataBoss) status.boss = true;
          foundApproval = true;
          console.log('✓ 承認状態取得成功 - 社員ID:', sub.id, '名前:', sub.name, 'yearMonth:', yearMonth, 'self:', status.self, 'boss:', status.boss);
          break;
        }
      }
      
      if (!foundApproval) {
        console.log('✗ 承認データ未発見 - 社員ID:', sub.id, '名前:', sub.name, 'yearMonth:', yearMonth);
      }
      
      
      sub.attendanceDays = attendanceCount;
      sub.workDays = workDays;
      sub.status = status;
      sub.incompleteDays = incompleteDays; // 退勤未打刻の日付リスト
      // 承認可能条件: (本人確認済み) AND (未承認)
      // ※以前は (attendanceCount >= workDays) も条件に含めていましたが、有給等の扱いや途中承認の柔軟性のため除外しました
      // 社員本人が「担当者印」を押していれば、日数が不足していても承認プロセスに進めるようにします
      sub.canApprove = status.self && !status.boss;
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      result: 'success',
      isApprover: true,
      subordinates: subordinates
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
