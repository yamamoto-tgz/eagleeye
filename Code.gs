function main() {
  // 楽天カード
  const rakutenCardPayments = process("label:rakuten/rakutencardpayment", (message) => {
    const body = message.getPlainBody();
    const date = RegExp("■利用日: (.*)\r\n").exec(body)[1];
    const amount = RegExp("■利用金額: (.*)円").exec(body)[1].replace(",", "");
    const description = RegExp("■利用先: (.*)\r\n").exec(body)[1];
    const source = "RakutenCard";

    message.moveToTrash();

    return [date, amount, description, source];
  });

  // 楽天デビット
  const rakutenDebitPayments = process("label:rakuten/rakutendebitpayment", (message) => {
    const date = Utilities.formatDate(message.getDate(), "JST", "yyyy/MM/dd");

    const body = message.getPlainBody();
    const amount = RegExp("口座引落分：(.*)円").exec(body)[1];
    const description = "ﾗｸﾃﾝ ﾃﾞﾋﾞｯﾄ";
    const source = "RakutenDebit";

    message.moveToTrash();

    return [date, amount, description, source];
  });

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

  const rows = [].concat(rakutenCardPayments, rakutenDebitPayments, rakutenPayPayments);
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
