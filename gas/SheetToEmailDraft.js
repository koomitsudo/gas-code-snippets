const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_ID = spreadsheet.getId();
const ss = SpreadsheetApp.openById(SHEET_ID);
const dataSheet = ss.getSheetByName("月初送信");
const midMonthDataSheet = ss.getSheetByName("月中15日送信");
const configSheet = ss.getSheetByName("template");

// 日付フォーマット関数
function getFormattedDate() {
  const now = new Date();
  return Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

// バリデーション関数
function validateSheet(sheet, configSheet, configCells) {
  if (!sheet) {
    showErrorAlert(`シート "${sheet.getName()}" が無いかも？`);
    return false;
  }

  if (sheet.getLastRow() <= 1) {
    showErrorAlert(`シート "${sheet.getName()}" にデータが無いかも？`);
    return false;
  }

  for (let cell of configCells) {
    if (!configSheet.getRange(cell).getValue()) {
      showErrorAlert(`設定シートのセル ${cell} に値が無いかも？`);
      return false;
    }
  }

  return true;
}

// エラーアラート表示関数
function showErrorAlert(message) {
  SpreadsheetApp.getUi().alert('エラー', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// 成功アラート表示関数
function showSuccessAlert(message) {
  SpreadsheetApp.getUi().alert('成功', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// 月初用
function runMonthDraft() {
  if (validateSheet(dataSheet, configSheet, ['B2', 'B3', 'B4', 'B5'])) {
    const excelBlob = createExcelBlob(dataSheet);
    createDraftEmail(excelBlob, configSheet, 'B2', 'B3', 'B4', 'B5');
    showSuccessAlert('月初の下書きメールが作成されました。');
  }
}

// 月中用
function runMidMonthDraft() {
  if (validateSheet(midMonthDataSheet, configSheet, ['C2', 'C3', 'C4', 'C5'])) {
    const excelBlob = createExcelBlob(midMonthDataSheet);
    createDraftEmail(excelBlob, configSheet, 'C2', 'C3', 'C4', 'C5');
    showSuccessAlert('月中15日の下書きメールが作成されました。');
  }
}

// Excelファイルを作成
function createExcelBlob(sheet) {
  const data = sheet.getDataRange().getValues();
  const tempSpreadsheet = SpreadsheetApp.create("TempSpreadsheet");
  const tempSheet = tempSpreadsheet.getSheets()[0];

  // データを一時シートにコピー
  tempSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  // 日付形式を設定
  const dateColumn = 3; // C列
  const dateRange = tempSheet.getRange(2, dateColumn, tempSheet.getLastRow() - 1, 1);

  // カスタム日付形式を設定
  dateRange.setNumberFormat("yyyy/mm/dd hh:mm:ss");

  // 日付データを文字列として再フォーマット
  const dateValues = dateRange.getValues();
  const formattedDateValues = dateValues.map(row => {
    const date = row[0];
    if (date instanceof Date) {
      return [[Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss")]];
    }
    return row;
  });
  dateRange.setValues(formattedDateValues);

  // Excelファイルとしてエクスポート
  const url = "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + tempSpreadsheet.getId() + "&exportFormat=xlsx";
  const params = {
    method: "get",
    headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    muteHttpExceptions: true
  };

const blob = UrlFetchApp.fetch(url, params).getBlob();

// ファイル名に日付を追加
const formattedDate = getFormattedDate();
blob.setName(`data_${formattedDate}.xlsx`);

// 一時スプレッドシートを削除
DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);

return blob;
}

// Gmailの下書きを行う
function createDraftEmail(excelBlob, configSheet, toCell, senderCell, subjectCell, bodyCell) {
  const recipients = configSheet.getRange(toCell).getValue().split(',').map(email => email.trim());
  const sender = configSheet.getRange(senderCell).getValue().trim();
  const subject = configSheet.getRange(subjectCell).getValue();
  const body = configSheet.getRange(bodyCell).getValue();

  const userEmail = Session.getActiveUser().getEmail();
  let to = recipients.join(',');
  let cc = '';
  if (sender !== userEmail) {
    cc = sender;
  }
 
  let emailOptions = {
    cc: cc,
    attachments: [excelBlob]
  };
 
  if (sender === userEmail) {
    emailOptions.from = sender;
  }
 
  GmailApp.createDraft(to, subject, body, emailOptions);
}
