/**
 *全商品（レシピ）マスタを取得する関数
 */
function getAllProductMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("全商品マスタ"); // ← 実際の全商品シート名に変更してください
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues(); // A列:番号, B列:商品名 を想定
  return data
    .filter(row => row[0] !== "" && row[1] !== "")
    .map(row => ({
      recordNo: String(row[0]).trim(),
      productName: String(row[1]).trim()
    }));
}