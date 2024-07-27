// プロパティに関する情報
const sp = PropertiesService.getScriptProperties();

// Googleスプレッドシート情報
const SHEET_ID = sp.getProperty('SHEET_ID');
const SHEET_NAME = sp.getProperty('SHEET_NAME');
const ss = SpreadsheetApp.openById(SHEET_ID);
const sheet = ss.getSheetByName(SHEET_NAME);

const WEBHOOK_URL = sp.getProperty('WEBHOOK_URL');

// Gmailで受信するGithubのコメントをGoogle チャットに投稿
function getGithubCommentFromGmail() {
    const subject = "[hoge/fuga]";
    const mailMax = 5;
    const threads = GmailApp.search(subject, 0, mailMax);
    const messages = GmailApp.getMessagesForThreads(threads);

    let firstRow = 2;

    messages.forEach(thread => {
        const message = thread[0];
        const chatMessage = message.getSubject();
        const dateMessage = message.getDate();
        const date = Utilities.formatDate(dateMessage, "JST", "yyyy/MM/dd");

        const botMessage = { 'text': chatMessage };
        const options = {
            'method': 'POST',
            'headers': {
                'Content-Type': 'application/json; charset=UTF-8'
            },
            'payload': JSON.stringify(botMessage)
        };
        const _result = UrlFetchApp.fetch(WEBHOOK_URL, options);
        const result = JSON.parse(_result.getContentText());

        sheet.getRange(firstRow, 1).setValue(date);
        sheet.getRange(firstRow, 2).setValue(result.text);
        firstRow += 1;
    });
}
