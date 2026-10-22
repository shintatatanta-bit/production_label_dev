/**
 * 製造ラベル発行Webアプリの入口 (GETリクエスト処理)
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile("LabelPrinter")
    .evaluate()
    .setTitle("A-one 62295 ラベルシール発行システム")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}