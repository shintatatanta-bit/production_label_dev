/**
 * 商品リストおよびロット台帳からラベル発行用のマスタ一覧を取得
 */
function getMasterProductList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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

  // ② ロット台帳から最新の賞味期限目安を補完（A列: ロットID, B列: 商品名, D列: 賞味期限）
  if (lotSheet) {
    const lotData = lotSheet.getDataRange().getValues();
    for (let r = 1; r < lotData.length; r++) {
      const lotId = String(lotData[r][0] || "").trim(); // 5962-001 など
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