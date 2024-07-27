// 最初に実行されるプロセス
function precedeProcess() {
    try {
        // なんらかの処理を記述
        execute();
        configSheet.getRange("A2").setValue("実行完了");
        // 後続処理(GAS)のためのトリガーをセット
        setTrigger();
    } catch (e) {
        GmailApp.sendEmail("slack@example.slack.com", "件名", e, options);
    }
}

// 次に実行されるプロセス
function succeedProcess() {
    const datetime = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd (E) HH:mm:ss Z");
    configSheet.getRange("C2").setValue(datetime);
    deleteTriggers();
    try {
        const result = configSheet.getRange("A2").getValue();

        if (result === "実行完了") {
            // 処理を記述
            executeAgain();
            configSheet.getRange("A2").setValue("");
        } 
    } catch (e) {
        GmailApp.sendEmail("slack@example.slack.com", "件名", e, options);
    }
}

// -------------------------
//  トリガー 
// -------------------------

// トリガーを作成
function setTrigger() {
    // トリガーをセットしたい関数を記載
    const funcName = "succeedProcess";
    // triggerTime分後に実行されるトリガー作成
    const triggerTime = configSheet.getRange("D2").getValue();
    ScriptApp.newTrigger(funcName).timeBased().after(triggerTime * 60 * 1000).create();
    const trigger = ScriptApp.getProjectTriggers();
    const id = trigger[0].getUniqueId();
    // シートへ triggerId をセット
    configSheet.getRange("B2").setValue(id);
    GmailApp.sendEmail("slack@example.slack.com", "件名", "トリガーセットの完了", options);
}

// トリガー削除
function deleteTriggers() {
    // シートから triggerId を取得
    const triggerId = configSheet.getRange("B2").getValue();
    const currentTriggers = ScriptApp.getProjectTriggers();
    // トリガーを削除
    currentTriggers.some(currentTrigger => {
        if (currentTrigger.getUniqueId() === triggerId) {
            ScriptApp.deleteTrigger(currentTrigger);
            GmailApp.sendEmail("slack@example.slack.com", "件名", "トリガー削除の完了", options);
        }
    });
}
