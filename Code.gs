function doPost(e) {
  try {
    // APIエンドポイントの振り分け
    var jsonData = JSON.parse(e.postData.contents);
    var action = jsonData.action;
    
    // 月次確認機能のAPI
    if (action === 'getMonthlyData') {
      return getMonthlyData(e);
    } else if (action === 'recordApproval') {
      return recordApproval(e);
    } else if (action === 'getDepartments') {
      return getDepartments();
    } else if (action === 'getEmployeesByDepartmentAndMonth') {
      return getEmployeesByDepartmentAndMonth(e);
    } else if (action === 'approveEmployee') {
      return approveEmployee(e);
    } else if (action === 'getPersonalMonthlyData') {
      return getPersonalMonthlyData(e);
    }
    
    // 以下、通常の打刻処理

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
      debugSheet.appendRow([new Date(), '警告: 社員コード ' + employeeId + ' はマスタに未登録です']);
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

// 指定された日付と社員コードの行を探して更新、なければ新規追加する関数
function updateOrAppendRow(sheet, data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var debugSheet = ss.getSheetByName('デバッグログ');
  
  var lastRow = sheet.getLastRow();
  var foundRow = -1;
  
  debugSheet.appendRow([new Date(), '検索開始: 日付=' + data.date + ', 社員コード=' + data.id + ', アクション=' + data.action]);
  
  // データがある場合、既存の行を検索(直近100行程度を検索対象とする)
  if (lastRow > 1) {
    var startRow = Math.max(2, lastRow - 100);
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
    }
  } else {
    // 新規行を追加
    debugSheet.appendRow([new Date(), '新規行を追加します']);
    
    // A:日にち, B:社員コード, C:名前, D:種別出勤, E:出勤時刻, F:種別退勤, G:退勤時刻, H:勤務時間, I:備考
    var rowData = [
      data.date,
      data.id,
      data.name,
      data.action === 'in' ? '出勤' : '',
      data.action === 'in' ? data.time : '',
      data.action === 'out' ? '退勤' : '',
      data.action === 'out' ? data.time : '',
      '', // H列: 勤務時間(後で数式を設定)
      data.remarks // I列: 備考
    ];
    sheet.appendRow(rowData);
    
    // 新規行の場合も勤務時間の計算式を設定
    var newRow = sheet.getLastRow();
    sheet.getRange(newRow, 8).setFormula('=IF(AND(E' + newRow + '<>"", G' + newRow + '<>""), TEXT(G' + newRow + '-E' + newRow + ', "[h]:mm"), "")');
    debugSheet.appendRow([new Date(), '新規行を追加しました: ' + newRow + '行目']);
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

// 承認状態を取得する関数
function getApprovalStatus(department, yearMonth) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var approvalSheet = ss.getSheetByName('月次承認記録');
  
  if (!approvalSheet) {
    return {
      status: 'not_approved',
      firstApprover: null,
      firstApprovalDate: null,
      secondApprover: null,
      secondApprovalDate: null
    };
  }
  
  var lastRow = approvalSheet.getLastRow();
  if (lastRow <= 1) {
    return {
      status: 'not_approved',
      firstApprover: null,
      firstApprovalDate: null,
      secondApprover: null,
      secondApprovalDate: null
    };
  }
  
  var values = approvalSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  
  // 該当する部署と年月のレコードを検索
  for (var i = values.length - 1; i >= 0; i--) {
    if (values[i][0] === department && values[i][1] === yearMonth) {
      var status = 'not_approved';
      if (values[i][4]) {
        status = 'fully_approved'; // 第2承認済み
      } else if (values[i][2]) {
        status = 'first_approved'; // 第1承認済み
      }
      
      return {
        status: status,
        firstApprover: values[i][2] || null,
        firstApprovalDate: values[i][3] ? Utilities.formatDate(new Date(values[i][3]), "Asia/Tokyo", "yyyy/MM/dd HH:mm") : null,
        secondApprover: values[i][4] || null,
        secondApprovalDate: values[i][5] ? Utilities.formatDate(new Date(values[i][5]), "Asia/Tokyo", "yyyy/MM/dd HH:mm") : null
      };
    }
  }
  
  return {
    status: 'not_approved',
    firstApprover: null,
    firstApprovalDate: null,
    secondApprover: null,
    secondApprovalDate: null
  };
}

// 承認を記録する関数
function recordApproval(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var jsonData = JSON.parse(e.postData.contents);
    var department = jsonData.department;
    var yearMonth = jsonData.yearMonth;
    var approverName = jsonData.approverName;
    var approvalLevel = jsonData.approvalLevel; // 'first' or 'second'
    
    // 承認者マスタで権限を確認
    var hasPermission = checkApproverPermission(department, approverName, approvalLevel);
    if (!hasPermission) {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'error',
        message: 'この部署の' + (approvalLevel === 'first' ? '第1' : '第2') + '承認者として登録されていません'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // 月次承認記録シートを取得または作成
    var approvalSheet = ss.getSheetByName('月次承認記録');
    if (!approvalSheet) {
      approvalSheet = ss.insertSheet('月次承認記録');
      approvalSheet.appendRow(['部署', '年月', '第1承認者', '第1承認日時', '第2承認者', '第2承認日時']);
    }
    
    var lastRow = approvalSheet.getLastRow();
    var foundRow = -1;
    
    // 既存のレコードを検索
    if (lastRow > 1) {
      var values = approvalSheet.getRange(2, 1, lastRow - 1, 2).getValues();
      for (var i = 0; i < values.length; i++) {
        if (values[i][0] === department && values[i][1] === yearMonth) {
          foundRow = i + 2;
          break;
        }
      }
    }
    
    var now = new Date();
    
    if (foundRow > 0) {
      // 既存レコードを更新
      if (approvalLevel === 'first') {
        approvalSheet.getRange(foundRow, 3).setValue(approverName);
        approvalSheet.getRange(foundRow, 4).setValue(now);
      } else if (approvalLevel === 'second') {
        // 第2承認の前に第1承認が必要
        var firstApprover = approvalSheet.getRange(foundRow, 3).getValue();
        if (!firstApprover) {
          return ContentService.createTextOutput(JSON.stringify({
            result: 'error',
            message: '第1承認が完了していません'
          })).setMimeType(ContentService.MimeType.JSON);
        }
        approvalSheet.getRange(foundRow, 5).setValue(approverName);
        approvalSheet.getRange(foundRow, 6).setValue(now);
      }
    } else {
      // 新規レコードを追加
      if (approvalLevel === 'second') {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error',
          message: '第1承認が完了していません'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      approvalSheet.appendRow([
        department,
        yearMonth,
        approverName,
        now,
        '',
        ''
      ]);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      result: 'success',
      message: (approvalLevel === 'first' ? '第1' : '第2') + '承認を記録しました',
      approvalStatus: getApprovalStatus(department, yearMonth)
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 承認者の権限を確認する関数
function checkApproverPermission(department, approverName, approvalLevel) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName('社員マスタ');
  
  if (!masterSheet) {
    return false;
  }
  
  var lastRow = masterSheet.getLastRow();
  if (lastRow <= 1) {
    return false;
  }
  
  // 社員マスタ: A:コード, B:氏名, C:部署, D:承認対象部署, E:第1承認者, F:第2承認者
  // データ範囲を取得（F列まで）
  var values = masterSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  
  // 承認対象の部署の設定を探す
  // ※社員マスタの各行に部署の承認者が設定されている前提
  // ※同じ部署の設定が複数行にある場合は、どれか1つでも一致すればOKとする
  
  for (var i = 0; i < values.length; i++) {
    var targetDept = values[i][3]; // D列: 承認対象部署
    
    // 承認対象部署が一致するか確認（空の場合はC列の所属部署を使うフォールバックも考慮可能だが、今回はD列厳守）
    if (String(targetDept) === String(department)) {
      var firstApprover = values[i][4];  // E列: 第1承認者
      var secondApprover = values[i][5]; // F列: 第2承認者
      
      if (approvalLevel === 'first') {
        // 第1承認者の名前と一致するか
        if (String(firstApprover).trim() === String(approverName).trim()) {
          return true;
        }
      } else if (approvalLevel === 'second') {
        // 第2承認者の名前と一致するか
        if (String(secondApprover).trim() === String(approverName).trim()) {
          return true;
        }
      }
    }
  }
  
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
    
    var logValues = logSheet.getRange(2, 1, logLastRow - 1, 9).getValues(); // I列(備考)まで取得
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
        // 日付をキーにしてデータを格納
        personalData[dateKey] = {
          clockInType: logValues[i][3],
          clockInTime: logValues[i][4] ? Utilities.formatDate(new Date(logValues[i][4]), "Asia/Tokyo", "HH:mm") : '',
          clockOutType: logValues[i][5],
          clockOutTime: logValues[i][6] ? Utilities.formatDate(new Date(logValues[i][6]), "Asia/Tokyo", "HH:mm") : '',
          workingHours: logValues[i][7],
          remarks: logValues[i][8]
        };
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      result: 'success',
      data: personalData
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}


