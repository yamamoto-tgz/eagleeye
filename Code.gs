function main() {
  // 楽天ペイ
  const rakutenPayPayments = process("label:rakuten/rakutenpaypayment", (message) => {
    const body = message.getPlainBody();
    const date = RegExp("ご利用日時.*([0-9]{4}/[0-9]{2}/[0-9]{2}).*\r\n").exec(body)[1];
    const amount = RegExp("カード　+(.*)円").exec(body)[1].replace(",", "");
    const description = RegExp("ご利用店舗　+(.*)\r\n").exec(body)[1];
    const source = "RakutenPay";

    message.moveToTrash();

    return [date, amount, description, source];
  });

  const rows = rakutenPayPayments;
  if (rows.length == 0) return;

  // Spreadsheetへ保存
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  const sheetId = PropertiesService.getScriptProperties().getProperty("SHEET_ID");
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetById(sheetId);

  const row = sheet.getLastRow() + 1;
  const col = 1;

  const range = sheet.getRange(row, col, rows.length, rows[0].length);
  range.setValues(rows);
}

function process(query, extractor) {
  return GmailApp.search(query).flatMap((thread) => {
    return thread.getMessages().map(extractor);
  });
}
