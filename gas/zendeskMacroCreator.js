const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_ID = spreadsheet.getId();
const ss = SpreadsheetApp.openById(SHEET_ID);

// プロパティサービスから設定値を取得
const sp = PropertiesService.getScriptProperties();
const TOKEN = sp.getProperty("TOKEN");

function createMacros() {
    const configSheet = ss.getSheetByName("config");
    const EMAIL = configSheet.getRange(2, 2).getValue();
    const BASE_URL = configSheet.getRange(2, 3).getValue();

    const result = isExecutable();
    if (result === "ok") {
        const url = BASE_URL + "/api/v2/macros.json";
        const macroParams = getCreateParams();
        macroParams.forEach((e) => {
            Utilities.sleep(500);
            const title = e[1];
            const description = e[2];
            const body = e[3];
            const options = getCreateOptions(EMAIL, TOKEN, title, description, body);
            const response = UrlFetchApp.fetch(url, options);
            Logger.log("POST /api/v2/macros.json ：" + response.getResponseCode());
        });
    } else if (result === "cancel") {
        Browser.msgBox("キャンセルしました");
    }
}

// シートから新規作成Macroの設定値を取得
function getCreateParams() {
    const sh = ss.getSheetByName("createParams");
    const lastRow = sh.getLastRow() - 1;
    const lastColumn = sh.getLastColumn();
    return sh.getRange(2, 1, lastRow, lastColumn).getValues();
}

// 各種オプションを取得する関数を定義
function getCreateOptions(email, token, title, description, body) {
    return {
        method: "post",
        payload: setParams(title, description, body),
        contentType: "application/json",
        headers: {
            Authorization: "Basic " + Utilities.base64Encode(email + "/token:" + token),
        },
    };
}

// マクロの設定値（JSON）をエンコードして渡す関数を定義
function setParams(title, description, body) {
    const params = {
        macro: {
            title: title,
            active: true,
            description: description,
            actions: [
                {
                    field: "comment_value",
                    value: body,
                },
            ],
            restriction: {
                type: "User",
                id: USER_ID,
            },
        },
    };
    return JSON.stringify(params);
}

// シートの初期化の関数を定義
function clearSheetContents(sh) {
    const lastRow = sh.getLastRow();
    const lastColumn = sh.getLastColumn();
    sh.getRange(2, 1, lastRow, lastColumn).clearContent();
    Utilities.sleep(3000);
}

// 実行確認のためのポップアップの関数を定義
function isExecutable() {
    return Browser.msgBox("実行しても良いですか？", Browser.Buttons.OK_CANCEL);
}
