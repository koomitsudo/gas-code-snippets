// プロパティに関する情報
const sp = PropertiesService.getScriptProperties();
const FOLDER_ID = sp.getProperty('FOLDER_ID');
const GOOGLE_API_KEY = sp.getProperty('GOOGLE_API_KEY');

function main() {
    const imageFiles = getImgFileFromDrive();

    // 画像ファイルをVision APIで処理
    imageFiles.forEach(image => {
        const resultMessage = analyzeImage(image);
        console.log(resultMessage);
    });
}

// Googleドライブからファイル取得
function getImgFileFromDrive() {
    const images = [];
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const files = folder.getFiles();
    while (files.hasNext()) {
        const file = files.next();
        const blob = file.getBlob();
        const base64EncodedFile = Utilities.base64Encode(blob.getBytes());
        images.push(base64EncodedFile);
    }

    return images;
}

// Vision APIで画像を解析して結果を取得
function analyzeImage(image) {
    const apiKey = GOOGLE_API_KEY;
    const url = 'https://vision.googleapis.com/v1/images:annotate?key=' + apiKey;

    // 画像からテキストの検出
    const body = {
        "requests": [
            {
                "image": {
                    "content": image
                },
                "features": [
                    {
                        "type": "DOCUMENT_TEXT_DETECTION",
                    }
                ],
                "imageContext": {
                    "languageHints": ["jp-t-i0-handwrit"]
                }
            }
        ]
    };

    const head = {
        "method": "post",
        "contentType": "application/json",
        "payload": JSON.stringify(body),
        "muteHttpExceptions": true
    };

    const response = UrlFetchApp.fetch(url, head);
    const obj = JSON.parse(response.getContentText());
    const result = obj.responses[0].textAnnotations[0].description;

    return result;
}
