/**
 * 全商品（レシピ）マスタを取得する関数（食品トレース側を参照）
 */
function getAllProductMaster() {
  const traceSpreadsheetId = "1aBT8g4vPYv2jTNE9qS-mTGVZl3wkmc186yIn7EAg5gY";
  const ss = SpreadsheetApp.openById(traceSpreadsheetId);
  
  const sheet = ss.getSheetByName("全商品マスタ");
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  return data
    .filter(row => row[0] !== "" && row[1] !== "")
    .map(row => ({
      recordNo: String(row[0]).trim(),
      productName: String(row[1]).trim()
    }));
}