/**
 * Code.gs
 * Webアプリケーションへのアクセス時の処理
 */
function doGet(e) {
  // Index.html をレンダリングして返す
  var htmlOutput = HtmlService.createTemplateFromFile('Index').evaluate();
  htmlOutput.setTitle('太陽光発電設備 年次点検報告書システム');
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return htmlOutput;
}

/**
 * フロントエンドから送信されたデータを受け取る関数
 * @param {Object} formData フォームの入力データ
 * @return {String} 処理結果メッセージ
 */
function processFormSubmission(formData) {
  try {
    // ----------------------------------------------------
    // ここにスプレッドシートへの転記やPDF生成処理を実装します
    // 例: var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Report');
    // ----------------------------------------------------
    
    Logger.log("受信データ: " + JSON.stringify(formData));
    
    // 処理成功メッセージを返す
    return "点検データの送信と登録が完了しました。";
  } catch (error) {
    Logger.log("エラー: " + error.toString());
    throw new Error("データの送信に失敗しました: " + error.message);
  }
}
