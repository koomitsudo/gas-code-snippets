// redashのQueryをINPORTしつつQUERY関数をInsertする
function updateImportData() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('import');
    const queryTimestamp = new Date().getTime();
    const importFunction = `=QUERY(IMPORTDATA("${REDASH_BASE_URL}&t=${queryTimestamp}"), "${sqlQuery}")`;
    sheet.getRange("A1").setValue(importFunction);
    SpreadsheetApp.flush();
    Utilities.sleep(500);
}
