const ss = SpreadsheetApp.getActiveSpreadsheet();
const qaSheet = ss.getSheetByName("Q&A");
const retrievalSheet = ss.getSheetByName("retrieval");
const customInstructionSheet = ss.getSheetByName("custom_instruction");


function main() {
    // 現在のカウント値を取得。目安のため回数制限を設置
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
    // 指示種別を取得
    const instructionType = qaSheet.getRange("A1").getValue();
    if (instructionType === "一般質問" || instructionType === "読解") {
        executeInquiryProcess(instruction, inputText);
    } else {
        executeRequestProcess(instruction, inputText);
    }
}

// 各種依頼
function getValueInstruction(){
    // 指示種別に応じたプラスαの指示を追加
    const instructionTypeSheet = ss.getSheetByName("instruction_type");
    // Googleスプレッドシート側のプルダウンで指示種別を選択している
    const instructionType = qaSheet.getRange("A1").getValue();
    if (instructionType == "添削") {
        const instruction = instructionTypeSheet.getRange("B2").getValue();
        return instruction;
    } else if (instructionType == "文章作成") {
        const instruction = instructionTypeSheet.getRange("B3").getValue();
        return instruction;
    } else if (instructionType == "英文作成") {
        const instruction = instructionTypeSheet.getRange("B4").getValue();
        return instruction;
    } else if (instructionType == "外国語翻訳") {
        const instruction = instructionTypeSheet.getRange("B5").getValue();
        return instruction;
    } else if (instructionType == "一般質問") {
        const instruction = instructionTypeSheet.getRange("B6").getValue();
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
    // const retrievalInfo = retrievalSheet.getRange("A2").getValue();
    const customInstructionInfo = customInstructionSheet.getRange("A2").getValue();
    const prompt = instruction + "\n" + "Refer to the information in 【Custom Instructions】 and create a final, simple answer:\n\n【Original Question】\n" + inputText
                               + "\n\n【Custom Instructions】\n" + customInstructionInfo + "\n\nlang:ja"
    const response = callGPToAPI(prompt);
    // GPT-4oの回答を転記
    qaSheet.getRange("B3").setValue(response);
}

// 依頼プロセス
function executeRequestProcess(instruction, inputText) {
    const response = callGPToAPI(instruction + "\n" + inputText);
    // GPT-4oの回答を転記
    qaSheet.getRange("B3").setValue(response);
}

function callGPToAPI(prompt) {
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
