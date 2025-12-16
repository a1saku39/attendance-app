function doPost(e) {
  try {
    // デバッグ用ログシートの準備（最優先で確保）
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var debugSheet = ss.getSheetByName('デバッグログ');
    if (!debugSheet) {
      debugSheet = ss.insertSheet('デバッグログ');
      debugSheet.appendRow(['時刻', 'ログ内容']);
    }

    // バージョン確認用ログ (v2.0)
    debugSheet.appendRow([new Date(), '[INFO] doPost実行 (v2.0)']);
    debugSheet.appendRow([new Date(), '受信データ: ' + e.postData.contents]);

    // APIエンドポイントの振り分け
    var jsonData = JSON.parse(e.postData.contents);
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
    var timestamp = new Date(jsonData.timestamp);
    var remarks = jsonData.remarks || ''; // 備考
    
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
        logSheet.appendRow(['日にち', '社員コード', '名前', '種別出勤', '出勤時刻', '種別退勤', '退勤時刻', '勤務時間', '備考']);
    }
    
    // 2. 該当行の検索と更新
    var logLastRow = logSheet.getLastRow();
    var foundRow = -1;
    
    if (logLastRow > 1) {
         // データ量が多いと遅くなるので、直近のデータから探すか、全件探すか。月次修正なので日付指定は正確。
         // 日付は A列(1列目), IDは B列(2列目)
         var dataRange = logSheet.getRange(2, 1, logLastRow - 1, 2).getValues();
         
         for(var i=0; i<dataRange.length; i++) {
             // 日付比較
             var rowDate = dataRange[i][0];
             var rowDateStr = '';
             if (rowDate instanceof Date) {
                 rowDateStr = Utilities.formatDate(rowDate, "Asia/Tokyo", "yyyy/MM/dd");
             } else {
                 rowDateStr = String(rowDate);
             }
             
             // ID比較
             var rowId = String(dataRange[i][1]);
             
             if (rowDateStr === targetDateStr && rowId === String(employeeId)) {
                 foundRow = i + 2; // ヘッダー分(+1)と0始まりのインデックス(+1)
                 break;
             }
         }
    }
    
    if (foundRow > 0) {
        // 更新
        // D:種別出勤, E:出勤, F:種別退勤, G:退勤, I:備考
        if(clockInTime) {
            logSheet.getRange(foundRow, 4).setValue('出勤');
            logSheet.getRange(foundRow, 5).setValue(clockInTime);
        } else {
            logSheet.getRange(foundRow, 4).clearContent();
            logSheet.getRange(foundRow, 5).clearContent();
        }
        
        if(clockOutTime) {
            logSheet.getRange(foundRow, 6).setValue('退勤');
            logSheet.getRange(foundRow, 7).setValue(clockOutTime);
        } else {
            logSheet.getRange(foundRow, 6).clearContent();
            logSheet.getRange(foundRow, 7).clearContent();
        }
        
        logSheet.getRange(foundRow, 9).setValue(remarks);

        // 勤務時間(H列)の計算式再設定
        logSheet.getRange(foundRow, 8).setFormula('=IF(AND(E' + foundRow + '<>"", G' + foundRow + '<>""), TEXT(G' + foundRow + '-E' + foundRow + ', "[h]:mm"), "")');
        
        if(debugSheet) debugSheet.appendRow([new Date(), '既存行を更新しました: ' + foundRow + '行目']);
        
    } else {
        // 新規追加
        // ['日にち', '社員コード', '名前', '種別出勤', '出勤時刻', '種別退勤', '退勤時刻', '勤務時間', '備考']
        var newRowData = [
            targetDateStr,
            employeeId,
            username,
            clockInTime ? '出勤' : '',
            clockInTime,
            clockOutTime ? '退勤' : '',
            clockOutTime,
            '', // 勤務時間(数式)
            remarks,
            '', // 承認
            ''  // 位置情報(手動修正なので空)
        ];
        logSheet.appendRow(newRowData);
        var newRowNum = logSheet.getLastRow();
        logSheet.getRange(newRowNum, 8).setFormula('=IF(AND(E' + newRowNum + '<>"", G' + newRowNum + '<>""), TEXT(G' + newRowNum + '-E' + newRowNum + ', "[h]:mm"), "")');

        if(debugSheet) debugSheet.appendRow([new Date(), '新規行を追加しました: ' + newRowNum + '行目']);
    }
    
    // ソート
    if (logSheet.getLastRow() > 1) {
        var sortRange = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logSheet.getLastColumn());
        sortRange.sort([
            {column: 1, ascending: true}, 
            {column: 5, ascending: true}
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
  
  var lastRow = sheet.getLastRow();
  var foundRow = -1;
  
  debugSheet.appendRow([new Date(), '検索開始: 日付=' + data.date + ', 社員コード=' + data.id + ', アクション=' + data.action]);
  
  // データがある場合、既存の行を検索(直近100行程度を検索対象とする)
  if (lastRow > 1) {
    var startRow = Math.max(2, lastRow - 500);
    var numRows = lastRow - startRow + 1;
    var values = sheet.getRange(startRow, 1, numRows, 2).getValues(); // A列(日付)とB列(コード)を取得
    
    debugSheet.appendRow([new Date(), '検索範囲: ' + startRow + '行目から' + lastRow + '行目まで(' + numRows + '行)']);
    
    // 下から上に検索
    for (var i = values.length - 1; i >= 0; i--) {
      // 日付と社員コードが一致するか確認
      var sheetDate = values[i][0];
      var sheetId = values[i][1];
      
      // デバッグ: 元の値を記録
      debugSheet.appendRow([new Date(), '行' + (startRow + i) + 'をチェック: A列=' + sheetDate + ' (型:' + typeof sheetDate + '), B列=' + sheetId]);
      
      // 日付を統一フォーマットに変換
      var formattedSheetDate = '';
      
      // Dateオブジェクトかどうかを判定(Google Apps Script対応)
      var isDateObject = (typeof sheetDate === 'object' && sheetDate !== null && 
                         (sheetDate instanceof Date || Object.prototype.toString.call(sheetDate) === '[object Date]'));
      
      if (isDateObject) {
        // Dateオブジェクトの場合はフォーマット
        formattedSheetDate = Utilities.formatDate(sheetDate, "Asia/Tokyo", "yyyy/MM/dd");
        debugSheet.appendRow([new Date(), 'Dateオブジェクトを変換: ' + formattedSheetDate]);
      } else if (sheetDate) {
        // 文字列の場合はそのまま使用
        formattedSheetDate = String(sheetDate).trim();
        debugSheet.appendRow([new Date(), '文字列として処理: ' + formattedSheetDate]);
      }
      
      // 検索対象の日付も統一
      var formattedSearchDate = String(data.date).trim();
      
      // 社員コードも統一
      var formattedSheetId = String(sheetId).trim();
      var formattedSearchId = String(data.id).trim();
      
      debugSheet.appendRow([new Date(), '比較: シート日付="' + formattedSheetDate + '" vs 検索日付="' + formattedSearchDate + '", シートID="' + formattedSheetId + '" vs 検索ID="' + formattedSearchId + '"']);
      
      // 空白や型の違いを考慮して比較
      var dateMatch = formattedSheetDate === formattedSearchDate;
      var idMatch = formattedSheetId === formattedSearchId;
      
      debugSheet.appendRow([new Date(), '一致判定: 日付=' + dateMatch + ', ID=' + idMatch]);
      
      if (dateMatch && idMatch) {
        foundRow = startRow + i;
        debugSheet.appendRow([new Date(), '✓ 既存行を発見: ' + foundRow + '行目']);
        break;
      }
    }
    
    if (foundRow === -1) {
      debugSheet.appendRow([new Date(), '✗ 既存行が見つかりませんでした。新規行を追加します。']);
    }
  } else {
    debugSheet.appendRow([new Date(), 'シートにデータがありません。新規行を追加します。']);
  }
  
  // 位置情報文字列の作成 (JSON)
  var locationStr = '';
  if (data.location) {
    try {
      locationStr = JSON.stringify(data.location);
    } catch (e) {
      locationStr = 'error: ' + e.message;
    }
  }
  
  if (foundRow > 0) {
    // 既存の行を更新
    debugSheet.appendRow([new Date(), '既存行(' + foundRow + '行目)を更新します']);
    
    if (data.action === 'in') {
      sheet.getRange(foundRow, 4).setValue('出勤'); // D列: 種別出勤
      sheet.getRange(foundRow, 5).setValue(data.time); // E列: 出勤時刻
      sheet.getRange(foundRow, 9).setValue(data.remarks); // I列: 備考
      // 出勤時にも勤務時間の計算式を設定(退勤時刻が既に入力されている場合に備えて)
      sheet.getRange(foundRow, 8).setFormula('=IF(AND(E' + foundRow + '<>"", G' + foundRow + '<>""), TEXT(G' + foundRow + '-E' + foundRow + ', "[h]:mm"), "")');
      debugSheet.appendRow([new Date(), '出勤時刻を記録: ' + data.time]);
    } else if (data.action === 'out') {
      sheet.getRange(foundRow, 6).setValue('退勤'); // F列: 種別退勤
      sheet.getRange(foundRow, 7).setValue(data.time); // G列: 退勤時刻
      // H列に勤務時間の計算式を設定
      sheet.getRange(foundRow, 8).setFormula('=IF(AND(E' + foundRow + '<>"", G' + foundRow + '<>""), TEXT(G' + foundRow + '-E' + foundRow + ', "[h]:mm"), "")');
      // 退勤時も備考を更新(上書き)
      sheet.getRange(foundRow, 9).setValue(data.remarks); // I列: 備考
      debugSheet.appendRow([new Date(), '退勤時刻を記録: ' + data.time]);
    } else if (data.action === 'holiday') {
      var holidayText = data.option === 'paid_leave' ? '有給休暇' : '代休';
      sheet.getRange(foundRow, 4).setValue(holidayText); // D列
      sheet.getRange(foundRow, 5).clearContent(); // E列クリア
      sheet.getRange(foundRow, 6).clearContent(); // F列クリア
      sheet.getRange(foundRow, 7).clearContent(); // G列クリア
      sheet.getRange(foundRow, 9).setValue(data.remarks); // I列
      sheet.getRange(foundRow, 8).setValue(''); // H列(勤務時間)クリア
      debugSheet.appendRow([new Date(), '休日記録: ' + holidayText]);
    }

    // 位置情報の追記 (K列)
    if (locationStr) {
      var currentVal = sheet.getRange(foundRow, 11).getValue(); // K列
      var locArray = [];
      if (currentVal) {
        try {
          // JSONパースを試みる
          if (typeof currentVal === 'string' && (currentVal.startsWith('[') || currentVal.startsWith('{'))) {
             locArray = JSON.parse(currentVal);
             if (!Array.isArray(locArray)) locArray = [locArray]; // 配列でない場合は配列にする
          }
        } catch (e) {
          locArray = [];
        }
      }
      
      // 今回のデータを追加
      locArray.push({
        action: data.action,
        time: data.time,
        lat: data.location.lat,
        lng: data.location.lng,
        timestamp: new Date().toISOString()
      });
      
      sheet.getRange(foundRow, 11).setValue(JSON.stringify(locArray));
    }

  } else {
    // 新規行を追加
    debugSheet.appendRow([new Date(), '新規行を追加します']);

    // 位置情報の初期配列
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
    
    // A:日にち, B:社員コード, C:名前, D:種別出勤, E:出勤時刻, F:種別退勤, G:退勤時刻, H:勤務時間, I:備考
    var rowData = [
      data.date,
      data.id,
      data.name,
      data.action === 'in' ? '出勤' : (data.action === 'holiday' ? (data.option === 'paid_leave' ? '有給休暇' : '代休') : ''),
      data.action === 'in' ? data.time : '',
      data.action === 'out' ? '退勤' : '',
      data.action === 'out' ? data.time : '',
      '', // H列: 勤務時間(後で数式を設定)
      data.remarks, // I列: 備考
      '', // J列: 承認
      initialLocJson // K列: 位置情報
    ];
    sheet.appendRow(rowData);
    
    // 新規行の場合も勤務時間の計算式を設定
    var newRow = sheet.getLastRow();
    sheet.getRange(newRow, 8).setFormula('=IF(AND(E' + newRow + '<>"", G' + newRow + '<>""), TEXT(G' + newRow + '-E' + newRow + ', "[h]:mm"), "")');
    debugSheet.appendRow([new Date(), '新規行を追加しました: ' + newRow + '行目']);
  }

  // 日付順、出勤時刻順に並び替え
  if (sheet.getLastRow() > 1) {
    var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    range.sort([
      {column: 1, ascending: true}, // A列: 日にち
      {column: 5, ascending: true}  // E列: 出勤時刻
    ]);
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
    
    // データを取得
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'success',
        data: [],
        approvalStatus: getApprovalStatus(department, yearMonth)
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var values = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
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
        monthlyData.push({
          date: dateValue instanceof Date ? Utilities.formatDate(dateValue, "Asia/Tokyo", "yyyy/MM/dd") : dateValue,
          employeeId: values[i][1],
          name: values[i][2],
          clockInType: values[i][3],
          clockInTime: values[i][4],
          clockOutType: values[i][5],
          clockOutTime: values[i][6],
          workingHours: values[i][7],
          remarks: values[i][8]
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var approvalSheet = ss.getSheetByName('月次承認記録_個人');
  
  if (!approvalSheet) {
    return {
      self: null, // 本人確認
      boss: null  // 上長承認
    };
  }
  
  var lastRow = approvalSheet.getLastRow();
  if (lastRow <= 1) {
    return { self: null, boss: null };
  }
  
  var values = approvalSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  // A:年月, B:社員ID, C:本人確認者, D:本人確認日時, E:承認者, F:承認日時
  
  for (var i = values.length - 1; i >= 0; i--) {
    if (String(values[i][1]) === String(employeeId) && values[i][0] === yearMonth) {
      return {
        self: values[i][2] ? { name: values[i][2], date: values[i][3] } : null,
        boss: values[i][4] ? { name: values[i][4], date: values[i][5] } : null
      };
    }
  }
  
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
      // 年月から営業日数を計算（簡易版：土日を除く）
      var year = parseInt(yearMonth.split('-')[0]);
      var month = parseInt(yearMonth.split('-')[1]);
      var daysInMonth = new Date(year, month, 0).getDate();
      var workDays = 0;
      
      for (var day = 1; day <= daysInMonth; day++) {
        var date = new Date(year, month - 1, day);
        var dayOfWeek = date.getDay();
        if (dayOfWeek !== 0 && dayOfWeek !== 6) { // 土日以外
          workDays++;
        }
      }
      
      // 打刻データを取得
      var logLastRow = logSheet.getLastRow();
      if (logLastRow > 1) {
        var logValues = logSheet.getRange(2, 1, logLastRow - 1, 10).getValues();
        
        for (var i = 0; i < employees.length; i++) {
          var employeeId = employees[i].employeeId;
          var attendanceDays = {};
          var approved = false;
          
          // この社員の打刻データを集計
          for (var j = 0; j < logValues.length; j++) {
            var logDate = logValues[j][0];
            var logId = String(logValues[j][1]);
            var logApproval = logValues[j][9]; // J列（10列目）
            
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

// 社員の承認を記録（部署シートのJ列に○を記録）
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
    
    // J列に承認フラグを追加（ヘッダーがない場合は追加）
    var headerRow = logSheet.getRange(1, 1, 1, 10).getValues()[0];
    if (!headerRow[9]) {
      logSheet.getRange(1, 10).setValue('承認');
      if (debugSheet) {
        debugSheet.appendRow([new Date(), 'J列に「承認」ヘッダーを追加しました']);
      }
    }
    
    // 該当社員の該当月のデータを検索してJ列に○を記録
    var lastRow = logSheet.getLastRow();
    if (lastRow > 1) {
      var values = logSheet.getRange(2, 1, lastRow - 1, 2).getValues();
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
        var logId = String(values[i][1]);
        
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
          logSheet.getRange(i + 2, 10).setValue('○');
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
    } catch (e) {
      // ログ記録失敗は無視
    }
    
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
    
    var logLastRow = logSheet.getLastRow();
    if (logLastRow <= 1) {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'success',
        data: {}
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // K列(位置情報)まで取得するよう範囲を拡張
    // A:1, I:9, J:10, K:11
    // データがある範囲の最大列数を取得して、K列が含まれているか確認
    var maxCols = logSheet.getLastColumn();
    var fetchCols = Math.max(11, maxCols); // 少なくとも11列目までは取得
    
    var logValues = logSheet.getRange(2, 1, logLastRow - 1, fetchCols).getValues(); 
    var personalData = {};
    
    // 年月を正規化（"2024-11" → "2024/11" の両方に対応）
    var targetYearMonth1 = yearMonth; // "2024-11"
    var targetYearMonth2 = yearMonth.replace('-', '/'); // "2024/11"
    
    for (var i = 0; i < logValues.length; i++) {
      var logDate = logValues[i][0];
      var logId = String(logValues[i][1]);
      
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
        var kVal = logValues[i][10]; // index 10 = K列
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
        personalData[dateKey] = {
          clockInType: logValues[i][3],
          clockInTime: formatTime(logValues[i][4]),
          clockOutType: logValues[i][5],
          clockOutTime: formatTime(logValues[i][6]),
          workingHours: logValues[i][7],
          remarks: logValues[i][8],
          locationLog: locationLog // 位置情報ログを追加
        };
      }
    }
    
    // 承認状態を取得
    var approvalStatus = getApprovalStatus(employeeId, yearMonth);
    
    return ContentService.createTextOutput(JSON.stringify({
      result: 'success',
      data: personalData,
      approvalStatus: approvalStatus
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
      // 今日のデータを検索（直近50行を確認）
      var logLastRow = logSheet.getLastRow();
      if (logLastRow > 1) {
        var startRow = Math.max(2, logLastRow - 50);
        var numRows = logLastRow - startRow + 1;
        var logs = logSheet.getRange(startRow, 1, numRows, 7).getValues(); // G列(退勤時刻)まで
        
        for (var i = logs.length - 1; i >= 0; i--) {
          var logDate = logs[i][0];
          var logId = logs[i][1];
          
          var logDateStr = '';
          if (logDate instanceof Date) {
            logDateStr = Utilities.formatDate(logDate, "Asia/Tokyo", "yyyy/MM/dd");
          } else {
            logDateStr = String(logDate);
          }
          
          // 日付とIDが一致するか
          if (logDateStr === todayStr && String(logId) === String(empId)) {
            if (logs[i][4]) isClockedIn = true; // E列: 出勤時刻
            if (logs[i][6]) isClockedOut = true; // G列: 退勤時刻
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
    // 月の営業日数を計算（土日除く）
    var parts = yearMonth.split('-');
    var year = parseInt(parts[0]);
    var month = parseInt(parts[1]);
    var daysInMonth = new Date(year, month, 0).getDate();
    var workDays = 0;
    for (var d = 1; d <= daysInMonth; d++) {
      var dayCheck = new Date(year, month - 1, d).getDay();
      if (dayCheck !== 0 && dayCheck !== 6) workDays++;
    }
    
    var approvalSheet = ss.getSheetByName('月次承認記録_個人');
    var approvalData = [];
    if (approvalSheet && approvalSheet.getLastRow() > 1) {
      approvalData = approvalSheet.getRange(2, 1, approvalSheet.getLastRow() - 1, 6).getValues();
    }
    
    // 各部下のデータ処理
    for (var i = 0; i < subordinates.length; i++) {
      var sub = subordinates[i];
      var sheetName = '打刻_' + sub.department;
      var logSheet = ss.getSheetByName(sheetName);
      
      var attendanceCount = 0;
      
      if (logSheet && logSheet.getLastRow() > 1) {
        // 打刻データをスキャン（効率化のため全取得は避けるべきだが、簡易実装として）
        var logs = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 2).getValues();
        var countedDays = {};
        
        for (var j = 0; j < logs.length; j++) {
           var logDate = logs[j][0];
           var logId = String(logs[j][1]);
           
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
               }
             }
           }
        }
      }
      
      // 承認状態の確認
      var status = { self: false, boss: false };
      for (var k = 0; k < approvalData.length; k++) {
        // A:年月, B:社員ID, C:本人, E:承認者
        if (approvalData[k][0] === yearMonth && String(approvalData[k][1]) === String(sub.id)) {
          if (approvalData[k][2]) status.self = true;
          if (approvalData[k][4]) status.boss = true;
          break;
        }
      }
      
      sub.attendanceDays = attendanceCount;
      sub.workDays = workDays;
      sub.status = status;
      // 承認可能条件: (平日日数 <= 打刻日数) AND (本人確認済み) AND (未承認)
      sub.canApprove = (attendanceCount >= workDays) && status.self && !status.boss;
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
