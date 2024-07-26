const ss = SpreadsheetApp.getActiveSpreadsheet();
const qaSheet = ss.getSheetByName("Q&A");

// 指示種別を取得
function getInstructionType(){
    // A1セルの値を取得（指示種別）
    const instructionType = qaSheet.getRange("A1").getValue();
    // 指示種別に応じた追加指示も返す
    if (instructionType == "添削") {
        const instruction = "You are a Japanese customer service representative. Please create 2 versions of the following text, keeping the style mostly unchanged, but making it simple, clear, and suitable for business use. Use headings like [Version 1] and [Version 2], without using Markdown. Please explain in Japanese only.";
        return instruction;
    } else if (instructionType == "文章作成") {
        const instruction = "You are a Japanese customer service representative. Based on the following request, please create 2 versions of a simple, clear, and business-appropriate text. Use headings like [Version 1] and [Version 2], without using Markdown. Please explain in Japanese only. No explanation is needed.";
        return instruction;
    } else if (instructionType == "英文作成") {
        const instruction = "Please create 2 versions of the following Japanese text in simple and business-appropriate English. Use headings like [Version 1] and [Version 2], without using Markdown. [Version 1] and [Version 2] should be in Japanese, with the text in English. No explanation is needed.";
        return instruction;
    } else if (instructionType == "外国語翻訳") {
        const instruction = "Please translate the following. If it\"s in English, translate it to Japanese. For other foreign languages, also translate to Japanese. Create 2 versions of the translation. Use headings like [Literal Translation] and [Easy-to-understand Interpretation], without using Markdown. No explanation is needed.";
        return instruction;
    } else if (instructionType == "一般質問") {
        const instruction = "You are a knowledgeable Japanese businessperson. For the following question, please consider what should be assumed as a premise, provide a simple and clear answer, present specific examples if necessary, briefly mention any points to note if any, and explain in detail why that is the case. Do not use Markdown. Please explain in Japanese only.";
        return instruction;
    } else if (instructionType == "読解") {
        const instruction = "You are a Japanese customer service representative. Please explain the following text in an easy-to-understand manner. If there are any unclear points, present two possible interpretations and make the most logical guess. If interpretations differ, use headings like [Interpretation 1] and [Interpretation 2], without using Markdown. Please explain in Japanese only.";
        return instruction;
    } else {
        return;
    }
}

// GPT-4を呼び出す
function callGPT4() {
    // 目安のため回数制限を設置。現在のカウント値を取得
    const currentCnt = qaSheet.getRange("C3");
    const currentCntValue = currentCnt.getValue();
    // 上限に到達したらアラートを表示
    if (currentCntValue >= 201) {
        SpreadsheetApp.getUi().alert("回数上限です");
        return;
    }
    if (currentCntValue === "") {
        currentCnt.setValue(1);
    } else if (typeof currentCntValue === "number") {
        let counter = currentCntValue;
        counter += 1;
        currentCnt.setValue(counter);
    } else {
        SpreadsheetApp.getUi().alert("文字列が入っている");
        return;
    }
   
    // 指示文を取得
    const instruction = getInstructionType();
    // 入力プロンプトを取得し、指示文をプロンプトへ追加
    const _inputText = qaSheet.getRange("A3").getValue();
    const inputText = instruction + "\n" + _inputText + "\n\nlang:ja";

    const apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
    const model = "gpt-4o-mini-2024-07-18";

    const payload = {
        messages: [{
            "role": "system",
            "content": "You are a helpful assistant."
        }, {
            "role": "user",
            "content": inputText
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

    // GPT-4の回答を転記
    qaSheet.getRange("B3").setValue(responseText.choices[0].message.content);
}