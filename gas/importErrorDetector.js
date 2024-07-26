function notifImportError() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const SHEET_ID = spreadsheet.getId();
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const hogeSheet = ss.getSheetByName("hoge");
    const shouldPosted = hogeSheet.getRange(1, 9).getValue();

    const properties = PropertiesService.getScriptProperties();
    const SHEET_URL = properties.getProperty('SHEET_URL');
    const WEBHOOK_URL = properties.getProperty('WEBHOOK_URL');

    // Googleスプレッドシート側でIMPORT系関数などのエラーを検知
    if (shouldPosted === "更新エラー疑惑") {
        const json = {
            'username': 'notif',
            'text': "更新元のAPI実行エラー疑惑\n" + "```" + "以下シートを確認\n" + SHEET_URL + "```"
        };
        const payload = JSON.stringify(json);
        const options = {
            'method': 'post',
            'contentType': 'application/json',
            'payload': payload
        };
        // Slackなどへ通知する
        UrlFetchApp.fetch(WEBHOOK_URL, options);
    } else {
        Logger.log("更新エラー無し");
    }
}
