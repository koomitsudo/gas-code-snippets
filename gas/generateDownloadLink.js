// プロパティサービスをオブジェクト化
const sp = PropertiesService.getScriptProperties();

// Googleスプレッドシート情報
const SHEET_ID = sp.getProperty('SHEET_ID');
const SHEET_NAME = sp.getProperty('SHEET_NAME');

// Google Drive情報
const DRIVE_ID = sp.getProperty('DRIVE_ID');

// クリックでファイルをDL可能にしたexport URLを転記
function genDownloadUrl() {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME); 
    let firstRow = 2;

    // Google Driveの特定フォルダに保存されているファイルを取得
    const files = DriveApp.getFolderById(DRIVE_ID).getFiles();
    while (files.hasNext()) {
        const file = files.next();
        const name = file.getName();

        // DL URLを生成
        const url = file.getUrl().replace("file/d/", "uc?export=download&id=").replace("/view?usp=drivesdk", "");

        sheet.getRange(firstRow, 1).setValue(name); 
        sheet.getRange(firstRow, 2).setValue(url);  
        firstRow += 1;
    }
}
