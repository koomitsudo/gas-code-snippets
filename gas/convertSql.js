function convertSqlKeywordsToUpper() { 
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); 
    const sql = sheet.getRange('A2').getValue(); 
    const regex = /(?:\s|^)(select|from|where|left join|inner join|cross join|outer join|join|group by|order by|having|case|when|then|else|with|and|as|end|on|union|limit|is null|is not null|ilike|like|distinct|at time zone 'jst')(?=\s|,|;|$)/gi; 

    const convertedSQL = sql.replace(regex, function(match, p1) { 
        return match.toUpperCase(); // 全体を大文字に変換 
    }); 

    sheet.getRange('B2').setValue(convertedSQL); 
}
