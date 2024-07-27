/*
SHEET_ID : スプレッドシート自体のID 
TOKEN    : Zendesk APIの Access Token
EMAIL    : 自分のメアド
BASE_URL : ZendeskのサブドメインまでのURL
USER_ID  : Zendeskの自分のUser ID
BRAND_ID : ZendeskのブランドID
*/

const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_ID = spreadsheet.getId();
const ss = SpreadsheetApp.openById(SHEET_ID);

// Zendesk設定値
const configSheet = ss.getSheetByName("config");
const sp = PropertiesService.getScriptProperties();
// Zendesk設定値
const ZENDESK_TOKEN = sp.getProperty('TOKEN');
const EMAIL = configValues[1];
const BASE_URL = configValues[2];
const USER_ID = configValues[3].toString();
const BRAND_ID = configValues[4].toString();


// マクロ情報をシートへ連携
function getMacroValuesToSheet() {
    const result = isExecutable();
    if (result === "ok") {
        const macroValues = fetchMacroValues();
        setMacroValues(macroValues);
    } else if (result === "cancel") {
        Browser.msgBox("キャンセルしました");
    }
}

// API経由でMacro情報を取得
function fetchMacroValues() {
    const options = getOptions();
    const url = BASE_URL + "/api/v2/macros.json";
    const response = UrlFetchApp.fetch(url, options);
    console.log("GET /api/v2/macros.json ："+response.getResponseCode())
    const values = JSON.parse(response.getContentText("UTF-8"));
    return values
}

// シートへ転記
function setMacroValues(values) {
    const sh = ss.getSheetByName("macros");
    clearSheetContents(sh);
    let rowCounter = 2
    values.macros.forEach(value => {
        sh.getRange(rowCounter, 1).setValue(TODAY);
        sh.getRange(rowCounter, 2).setValue(value.id);
        sh.getRange(rowCounter, 3).setValue(value.title);
        sh.getRange(rowCounter, 4).setValue(value.description);
        // <br>を改行に変換後にHTMLタグを除去
        sh.getRange(rowCounter, 5).setValue(value.actions[0].value
            .replace(/<br>/g, "\n")
            .replace(/<("[^"]*"|'[^']*'|[^'">])*>/g,''));
        // agentがアクセス可能なurlに変換
        sh.getRange(rowCounter, 6).setValue(value.url
            .replace(".json", "")
            .replace("/api/v2/", "/admin/workspaces/agent-workspace/"));
        sh.getRange(rowCounter, 7).setValue(value.active);   
        sh.getRange(rowCounter, 8).setValue(value.restriction);
        rowCounter += 1;
    });
}

// Optionを取得
function getOptions() {
    const options = {
        method: "get",
        contentType: "application/json",
        headers: {
            Authorization: "Basic " + Utilities.base64Encode(EMAIL + "/token:" + TOKEN),
        },
    };
    return options;
}

// シートを初期化
function clearSheetContents(sh) {
    const lastRow = sh.getLastRow();
    const lastColum = sh.getLastColumn();
    sh.getRange(2,1,lastRow,lastColum).clearContent();
    Utilities.sleep(3000);
}

// 実行確認のポプアップ
function isExecutable() {
    return Browser.msgBox("実行しても良いですか？", Browser.Buttons.OK_CANCEL);
}