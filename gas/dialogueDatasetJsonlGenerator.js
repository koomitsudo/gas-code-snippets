const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('dialogue');

function main() {
    const jsonl = generateJsonl();
    createExportLink(jsonl);
}

// Googleスプレッドシートから対話セットを取得してJSONL形式に変換
function generateJsonl() {
    const lastRow = sheet.getLastRow() - 1;
    const data = sheet.getRange(2, 1, lastRow, 2).getValues();
    const systemContent = sheet.getRange("C2").getValue();
    
    let jsonl = "";
    
    data.forEach(row => {
        const userContent = row[0];
        const assistantContent = row[1];
        const jsonObj = {
            "messages": [
                {
                    "role": "system",
                    "content": systemContent
                },
                {
                    "role": "user",
                    "content": userContent
                },
                {
                    "role": "assistant",
                    "content": assistantContent
                }
            ]
        };

        jsonl += JSON.stringify(jsonObj) + "\n";
    });
    return jsonl;
}

// Export用のリンクを生成
function createExportLink(jsonl) {
    const FOLDER_ID = sheet.getRange("E2").getValue();
    const currentDate = Utilities.formatDate(new Date(), "JST", "yyyyMMdd")
    const fileName = `dialogue_${currentDate}.jsonl`;
    const folder = DriveApp.getFolderById(FOLDER_ID); 
    const file = folder.createFile(fileName, jsonl, "text/plain");
    const fileId = file.getId();
    const exportLink = `https://drive.google.com/file/d/${fileId}/view`;
    sheet.getRange("D2").setValue(exportLink);
}
