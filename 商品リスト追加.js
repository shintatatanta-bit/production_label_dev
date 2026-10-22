/**
 * ラベル発行画面からの商品追加申請を保存する関数（申請者なし）
 */
function requestProductAddition(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("商品追加申請");
  
  // 「商品追加申請」シートがなければ自動作成する
  if (!sheet) {
    sheet = ss.insertSheet("商品追加申請");
    sheet.getRange(1, 1, 1, 4).setValues([["申請日時", "レコード番号", "商品名", "状態"]]);
    sheet.getRange("A1:D1").setBackground("#1a73e8").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  const recordNo = String(data.recordNo || "").trim();
  const productName = String(data.productName || "").trim();

  if (!recordNo || !productName) {
    return { success: false, message: "レコード番号と商品名は必須です。" };
  }

  const nowStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
  
  // 申請シートへ追記（申請者カラムなし：日時、レコード番号、商品名、状態）
  sheet.appendRow([nowStr, recordNo, productName, "未承認"]);
  SpreadsheetApp.flush();

  return { success: true };
}

function getPendingProductRequests() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("商品追加申請");
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getDataRange().getValues();
  const requests = [];

  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][3] || "").trim();
    if (status === "未承認") {
      // 日付データを確実に「YYYY/MM/DD HH:mm:ss」の文字列に変換
      let dateVal = data[i][0];
      if (dateVal instanceof Date) {
        dateVal = Utilities.formatDate(dateVal, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
      }

      requests.push({
        rowIndex: i + 1,
        date: String(dateVal),
        recordNo: String(data[i][1]).trim(),
        productName: String(data[i][2]).trim()
      });
    }
  }
  return requests;
}

/**
 * 申請を承認し、「商品リスト」へ自動追加する（最終行の直下に追加する改善版）
 */
function approveProductRequest(rowIndex, recordNo, productName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reqSheet = ss.getSheetByName("商品追加申請");
  const listSheet = ss.getSheetByName("商品リスト");

  if (!reqSheet || !listSheet) {
    return { success: false, message: "必要なシートが見つかりません。" };
  }

  try {
    // 1. 「商品追加申請」シートの状態を「承認済」に更新
    reqSheet.getRange(Number(rowIndex), 4).setValue("承認済");

    // 2. 「商品リスト」の実際のデータ最終行を安全に取得してその次に追加
    // ※保護の影響で下方に空行がある場合でも、データが入っている最後の行を正しく探します
    const listData = listSheet.getDataRange().getValues();
    let lastDataRow = 1;
    for (let r = listData.length - 1; r >= 0; r--) {
      // A列またはB列にデータが入っている最初の行を探索
      if (String(listData[r][0]).trim() !== "" || String(listData[r][1]).trim() !== "") {
        lastDataRow = r + 1;
        break;
      }
    }

    // データのすぐ次の行に書き込み
    listSheet.getRange(lastDataRow + 1, 1, 1, 2).setValues([[recordNo, productName]]);

    SpreadsheetApp.flush();
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  }
}