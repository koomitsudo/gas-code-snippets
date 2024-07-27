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


// 取得したいフィールドがあれば
const FIELD_01 = "<field_id_01>"
const FIELD_02 = "<field_id_02>"
const FIELD_03 = "<field_id_03>"


// 取得したいフィールドがあれば
const FORM_IDS = {
    "<form_id_01>": "<form_title_01>",
    "<form_id_02>": "<form_title_02>",
    "<form_id_03>": "<form_title_03>",
}


// 作成済みチケット情報を取得して転記
function getTicketValuesToSheet(){
    const result = isExecutable()
    if (result==="ok"){
        const ticketValues = fetchTicketValues();
        setTicketValues(ticketValues);
    } else if (result==="cancel"){
        Browser.msgBox("キャンセルしました")
    }
}


// API経由でTicket情報を取得
function fetchTicketValues() {
    const options = getOptions();
    const url = BASE_URL + "/api/v2/tickets.json";
    const response = UrlFetchApp.fetch(url, options);
    Logger.log("GET /api/v2/tickets.json ："+response.getResponseCode());

    let valuesArray = [];
    const values = JSON.parse(response.getContentText("UTF-8"));
    valuesArray.push(JSON.parse(response.getContentText("UTF-8")));

    let pageCounter = 2
    while(values.next_page !== null){
        const url = BASE_URL + "/api/v2/tickets.json?page="+pageCounter;
        const response = UrlFetchApp.fetch(url, options);
        Logger.log("/api/v2/tickets.json="+pageCounter + " ："+response.getResponseCode())
        valuesArray.push(JSON.parse(response.getContentText("UTF-8")));
        pageCounter += 1;
        const next_page_values = JSON.parse(response.getContentText("UTF-8"));
        if (next_page_values.next_page === null || pageCounter > 100) {
            break;
        }
    }
    return valuesArray;
}

// 取得した情報をスプレッドシートへ転記する
function setTicketValues(valuesArray) {
    const sh = ss.getSheetByName("tickets");
    clearSheetContents(sh);

    const _START_DATE = configSheet.getRange(2, 5).getValue();
    const _END_DATE = configSheet.getRange(2, 6).getValue();
    const START_DATE = Utilities.formatDate(new Date(_START_DATE), "UTC", "yyyy-MM-dd")
    const END_DATE = Utilities.formatDate(new Date(_END_DATE), "UTC", "yyyy-MM-dd")

    let rowCounter = 2
    for (const value of valuesArray) {
        value.tickets.forEach(val => {
            const tmpDate = Utilities.formatDate(new Date(val.created_at), "UTC", "yyyy-MM-dd")
            if(tmpDate > START_DATE && tmpDate <=  END_DATE){
                sh.getRange(rowCounter, 1).setValue(val.subject);
                // <br>を改行に変換後にHTMLタグを除去
                sh.getRange(rowCounter, 2).setValue(val.description
                    .replace(/<br>/g, "\n")
                    .replace(/<("[^"]*"|'[^']*'|[^'">])*>/g,''));
                // agentがアクセス可能なurlに変換
                sh.getRange(rowCounter, 3).setValue(val.url
                    .replace(".json", "")
                    .replace("/api/v2/", "/agent/"));
                sh.getRange(rowCounter, 4).setValue(val.id)
                sh.getRange(rowCounter, 5).setValue(Utilities.formatDate(new Date(val.created_at), "UTC", "yyyy-MM-dd"));
                for (let i = 0; i < val.custom_fields.length; i++) {
                    if (val.custom_fields[i].id == FIELD_01){
                        sh.getRange(rowCounter, 6).setValue(val.custom_fields[i].value);
                    } else if (val.custom_fields[i].id == FIELD_02){
                        sh.getRange(rowCounter, 7).setValue(val.custom_fields[i].value);
                    } else if (val.custom_fields[i].id == FIELD_03){
                        sh.getRange(rowCounter, 8).setValue(val.custom_fields[i].value);
                }
                sh.getRange(rowCounter, 9).setValue(FORM_IDS[val.ticket_form_id]);
                sh.getRange(rowCounter, 10).setValue(val.tags[0]);
                sh.getRange(rowCounter, 11).setValue(val.tags[1]);
                sh.getRange(rowCounter, 12).setValue(val.tags[2]);
                sh.getRange(rowCounter, 13).setValue(val.via.source.from.address);
                sh.getRange(rowCounter, 14).setValue(val.requester_id);
                rowCounter += 1;
            }
      });
    }
}