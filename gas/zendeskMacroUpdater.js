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

// 作成済みのマクロ設定を更新する
function updateMacros() {
    const result = isExecutable();
    if (result === "ok") {
        const macroParams = getUpdateParams();
        macroParams.forEach((param) => {
            const macroId = param[1];
            const title = param[2];
            const body = param[4];
            const active = param[9];
            const restriction = param[10];
            if (active !== "" || restriction !== "") {
                Utilities.sleep(100);
                const url = BASE_URL + `/api/v2/macros/${macroId}`;
                const options = getUpdateOptions(title, body, active, restriction);
                const response = UrlFetchApp.fetch(url, options);
                Logger.log(`PUT /api/v2/macros/${macroId}：${title}` + response.getResponseCode());
            }
        });
    } else if (result === "cancel") {
        Browser.msgBox("キャンセルしました");
    }
}

// シートからMacroの設定値を取得（アクティブ化と利用制限のフラグを取得）
function getUpdateParams() {
    const sh = ss.getSheetByName("macros");
    const lastRow = sh.getLastRow() - 1;
    const lastColumn = sh.getLastColumn();
    const macroParams = sh.getRange(2, 1, lastRow, lastColumn).getValues();
    return macroParams;
}

// Optionを取得
function getUpdateOptions(title, body, active, restriction) {
    if (restriction === "personal") {
        // update後のマクロ使用制限を個人
        const personalOptions = {
            "method": "put",
            "payload": personalRestriction(title, body, active),
            "contentType": "application/json",
            "headers": {
                "Authorization": "Basic " + Utilities.base64Encode(EMAIL + "/token:" + TOKEN)
            },
        };
        return personalOptions;
    } else {
        // update後のマクロ使用制限をメンバー全員に
        const allMemberOptions = {
            "method": "put",
            "payload": allMemberRestriction(title, body, active),
            "contentType": "application/json",
            "headers": {
                "Authorization": "Basic " + Utilities.base64Encode(EMAIL + "/token:" + TOKEN)
            },
        };
        return allMemberOptions;
    }
}

// Macroを個人のみ利用可能に変更
function personalRestriction(title, body, active) {
    const restriction = { "type": "User", "id": USER_ID };
    const params = {
        "macro": {
            "active": active,
            "actions": [
                {
                    "field": "comment_value",
                    "value": body,
                }
            ],
            "restriction": restriction,
            "title": title,
        }
    };
    return JSON.stringify(params);
}

// Macroを全員使用可能に変更
function allMemberRestriction(title, body, active) {
    const restriction = null;
    const params = {
        "macro": {
            "active": active,
            "actions": [
                {
                    "field": "comment_value",
                    "value": body,
                }
            ],
            "restriction": restriction,
            "title": title,
        }
    };
    return JSON.stringify(params);
}
