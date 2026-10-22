/**
 * 商品リストおよびロット台帳からラベル発行用のマスタ一覧を取得（食品トレース側を参照）
 */
function getMasterProductList() {
  // 食品トレースのスプレッドシートID
  const traceSpreadsheetId = "1aBT8g4vPYv2jTNE9qS-mTGVZl3wkmc186yIn7EAg5gY";
  const ss = SpreadsheetApp.openById(traceSpreadsheetId);
  
  const prodSheet = ss.getSheetByName("商品リスト");
  const lotSheet = ss.getSheetByName("ロット台帳");

  const productMap = {};

  // ① 商品リストの読み込み（A列: レコード番号, B列: 商品名）
  if (prodSheet) {
    const prodData = prodSheet.getDataRange().getValues();
    for (let r = 1; r < prodData.length; r++) {
      const rec = String(prodData[r][0] || "").trim();
      const name = String(prodData[r][1] || "").trim();
      if (rec && name) {
        productMap[rec] = {
          recordNo: rec,
          name: name,
          defaultExpiry: ""
        };
      }
    }
  }

  // ② ロット台帳から最新の賞味期限目安を補完
  if (lotSheet) {
    const lotData = lotSheet.getDataRange().getValues();
    for (let r = 1; r < lotData.length; r++) {
      const lotId = String(lotData[r][0] || "").trim();
      const recNo = lotId.split("-")[0];
      const rawExp = lotData[r][3];

      let expStr = "";
      if (rawExp instanceof Date) {
        expStr = Utilities.formatDate(rawExp, "Asia/Tokyo", "yyyy/MM/dd");
      } else if (rawExp) {
        expStr = String(rawExp).trim();
      }

      if (productMap[recNo] && expStr) {
        productMap[recNo].defaultExpiry = expStr;
      }
    }
  }

  return Object.values(productMap);
}