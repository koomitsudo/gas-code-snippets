function convertStyle() { 
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('修正'); 
    // 修正前の文章の取得 
    const originalText = sheet.getRange('A2').getValue(); 
    const convertedText = convertToPlainStyle(originalText); 
    sheet.getRange('B2').setValue(convertedText); 
} 

function convertToPlainStyle(text) { 
    // 変換対象のパターンと変換後の対応 
    const patterns = [ 
        // 文末の「ます」「です」の削除および変換 
        { pattern: /ます。/g, replacement: '。' }, 
        { pattern: /です。/g, replacement: '。' }, 
        { pattern: /ます/g, replacement: '' }, 
        { pattern: /です/g, replacement: '' }, 
        { pattern: /だが、/g, replacement: 'が、' }, 
        { pattern: /だが。/g, replacement: 'が。' }, 
        { pattern: /でしょうか。/g, replacement: 'だろうか。' }, 
        { pattern: /でしょう。/g, replacement: 'だろう。' }, 
        // 連用形の変換 
        { pattern: /してい。/g, replacement: 'している。' }, 
        { pattern: /されてい。/g, replacement: 'されている。' }, 
        { pattern: /されてい/g, replacement: 'されている' }, 
        { pattern: /されている[【（\(\[].+[】）\)\]]/g, replacement: 'されている' }, // 注釈対応 
        { pattern: /あります。/g, replacement: 'ある。' }, 
        { pattern: /あります/g, replacement: 'ある' }, 
        { pattern: /ある[【（\(\[].+[】）\)\]]/g, replacement: 'ある' }, // 注釈対応 
        { pattern: /なります。/g, replacement: 'なる。' }, 
        { pattern: /なります/g, replacement: 'なる' }, 
        { pattern: /なり。/g, replacement: 'なる。' }, 
        // その他 
        { pattern: /してください。/g, replacement: 'すること。' }, 
        { pattern: /しないでください。/g, replacement: 'しないこと。' }, 
        { pattern: /いません。/g, replacement: 'いない。' }, 
        { pattern: /しましょう。/g, replacement: 'しよう。' }, 
        { pattern: /している。/g, replacement: 'している。' }, 
        { pattern: /しています。/g, replacement: 'している。' }, 
        { pattern: /います。/g, replacement: 'いる。' }, 
        { pattern: /いませんか。/g, replacement: 'いないか。' }, 
        { pattern: /しましょうか。/g, replacement: 'しようか。' } 
    ]; 

    // テキストに対して変換を適用 
    let newText = text; 
    patterns.forEach(pair => { 
        newText = newText.replace(pair.pattern, pair.replacement); 
    }); 

    // 変換後の不正パターンを修正 
    newText = newText.replace(/であるる。/g, 'である。'); 
    newText = newText.replace(/られるる。/g, 'られる。'); 
    newText = newText.replace(/されているる。/g, 'されている。'); 
    newText = newText.replace(/ているる。/g, 'ている。'); 
    newText = newText.replace(/しているる。/g, 'している。'); 
    newText = newText.replace(/られ。/g, 'られる。'); 
    newText = newText.replace(/られてい。/g, 'られている。'); 
    return newText; 
}
