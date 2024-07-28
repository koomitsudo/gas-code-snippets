const ss = SpreadsheetApp.getActiveSpreadsheet();
const qaSheet = ss.getSheetByName("Q&A");
const retrievalSheet = ss.getSheetByName("retrieval");
const customInstructionSheet = ss.getSheetByName("custom_instruction");

// APIキーやトークン検出用
const apiKeyRegex = /[A-Za-z0-9_\-]{20,64}/g;

// クレジットカード番号検出用
const creditCardRegex = /\b(?:\d{4}[-\s]?){3,4}\d{1,4}\b/g;

// パスワード検出用（8文字以上、英数字と特殊文字を含む）
const passwordRegex = /\b(?=.*[A-Za-z])(?=.*\d)(?=.*[@$!%*?&])[A-Za-z\d@$!%*?&]{8,}\b/g;

function main() {
    // 現在のカウントを取得
    const currentCnt = qaSheet.getRange("C3");
    const currentCntValue = currentCnt.getValue();
    
    // カウントが201以上の場合、アラートを表示して終了
    if (currentCntValue >= 201) {
        SpreadsheetApp.getUi().alert("回数上限です");
        return;
    }
    
    // カウントが空の場合は1を設定
    if (currentCntValue === "") {
        currentCnt.setValue(1);
    // カウントが数値の場合は1つ増加
    } else if (typeof currentCntValue === "number") {
        let counter = currentCntValue;
        counter += 1;
        currentCnt.setValue(counter);
    // カウントが文字列の場合、アラート表示で終了
    } else {
        SpreadsheetApp.getUi().alert("文字列が入っている");
        return;
    }

    const instruction = getValueInstruction();
    // 入力テキストを取得
    const inputText = qaSheet.getRange("A3").getValue();

    // 入力テキストのセキュリティチェック
    if (!validateInput(inputText)) {
        // セキュリティチェックに失敗した場合、処理を中断
        return;
    }

    // 指示種別を取得
    const instructionType = qaSheet.getRange("A1").getValue();

    if (instructionType === "一般質問" || instructionType === "読解") {
        executeInquiryProcess(instruction, inputText);
    } else {
        executeRequestProcess(instruction, inputText);
    }
}


function getBlacklistedUrls() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('black_list');
    const urls = sheet.getRange('A2:A').getValues().flat().filter(String);
    return urls;
}

function createUrlRegex(urls) {
    const escapedUrls = urls.map(url => url.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
    return new RegExp('\\b(' + escapedUrls.join('|') + ')\\b', 'i');
}

function validateInput(inputText) {
    if (apiKeyRegex.test(inputText)) {
        SpreadsheetApp.getUi().alert("APIキーらしき文字列を検知したので確認して");
        return false;
    }

    if (creditCardRegex.test(inputText)) {
        SpreadsheetApp.getUi().alert("カード番号らしき文字列を検知したので確認して");
        return false;
    }

    if (passwordRegex.test(inputText)) {
        SpreadsheetApp.getUi().alert("パスワードらしき文字列を検知したので確認して");
        return false;
    }

    const blacklistedUrls = getBlacklistedUrls();
    const urlRegex = createUrlRegex(blacklistedUrls);
    if (urlRegex.test(inputText)) {
        SpreadsheetApp.getUi().alert("NGのURLを検知したので確認して");
        return false;
    }

    return true; // すべてのチェックをパスした場合
}


// 各種依頼
function getValueInstruction(){
    const instructionTypeSheet = ss.getSheetByName("instruction_type");
    const instructionType = qaSheet.getRange("A1").getValue();
    if (instructionType == "添削") {
        const instruction = instructionTypeSheet.getRange("B2").getValue();
        Logger.log("添削")
        return instruction;
    } else if (instructionType == "文章作成") {
        const instruction = instructionTypeSheet.getRange("B3").getValue();
        Logger.log("文章作成")
        return instruction;
    } else if (instructionType == "英文作成") {
        const instruction = instructionTypeSheet.getRange("B4").getValue();
        Logger.log("英文作成")
        return instruction;
    } else if (instructionType == "外国語翻訳") {
        const instruction = instructionTypeSheet.getRange("B5").getValue();
        Logger.log("外国語翻訳")
        return instruction;
    } else if (instructionType == "一般質問") {
        const instruction = instructionTypeSheet.getRange("B6").getValue();
        Logger.log("一般質問")
        return instruction;
    } else if (instructionType == "読解") {
        const instruction = instructionTypeSheet.getRange("B7").getValue();
        return instruction;
    } else {
        return;
    }
}

// 質問プロセス
function executeInquiryProcess(instruction, inputText) {
    const customInstructionInfo = customInstructionSheet.getRange("A2").getValue();
    const prompt = instruction + "\n" + "Refer to the information in 【Custom Instructions】 and create a final, simple answer:\n\n【Original Question】\n" + inputText
                               + "\n\n【Custom Instructions】\n" + customInstructionInfo + "\n\nlang:ja"
    const response = callGPTAPI(prompt);
    qaSheet.getRange("B3").setValue(response);
}

// 依頼プロセス
function executeRequestProcess(instruction, inputText) {
    const response = callGPTAPI(instruction + "\n" + inputText);
    qaSheet.getRange("B3").setValue(response);
}

function callGPTAPI(prompt) {
    const apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
    const model = "gpt-4o-mini-2024-07-18";
    const payload = {
        messages: [{
            "role": "system",
            "content": "You are a helpful assistant."
        }, {
            "role": "user",
            "content": prompt
        }],
        model: model
    };

    const options = {
        "method": "post",
        "contentType": "application/json",
        "headers": {
            "Authorization": "Bearer " + apiKey
        },
        "payload": JSON.stringify(payload)
    };

    const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", options);
    const responseText = JSON.parse(response.getContentText());
    return responseText.choices[0].message.content;
}
