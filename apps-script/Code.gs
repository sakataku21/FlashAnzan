/**
 * フラッシュ暗算アプリの記録先となる Google Apps Script です。
 *
 * スプレッドシートに紐づけて「ウェブアプリ」としてデプロイし、
 * 発行された /exec の URL をアプリの「記録設定」に貼り付けてください。
 * 手順は README.md を参照してください。
 */

var HEADERS = ['日時', '正誤', '桁数', '口数', '入力した答', '正答'];

function doPost(e) {
  // 同時に複数の結果が届いても行が壊れないよう直列化する
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getUserSheet(sheetNameFor(data.user));

    sheet.appendRow([
      parseDate(data.at),
      data.result,
      data.digits,
      data.count,
      data.answer === '' || data.answer === null ? '' : Number(data.answer),
      Number(data.correct)
    ]);

    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** 動作確認用。ブラウザで /exec を開くと応答が見えます。 */
function doGet() {
  return jsonResponse({ ok: true, message: 'flash anzan recorder is running' });
}

/**
 * ユーザー名からシートを取得し、無ければ見出し付きで作成します。
 */
function getUserSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (sheet) { return sheet; }

  sheet = ss.insertSheet(name);
  sheet.appendRow(HEADERS);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.getRange('A:A').setNumberFormat('yyyy/MM/dd HH:mm:ss');
  sheet.setColumnWidth(1, 160);
  return sheet;
}

/**
 * シート名に使えない文字を除きます。
 * Google スプレッドシートでは [ ] : * ? / \ が使えず、最大 100 文字です。
 */
function sheetNameFor(user) {
  var name = String(user == null ? '' : user).replace(/[\[\]:*?\/\\]/g, '_').trim();
  if (!name) { name = 'unknown'; }
  return name.slice(0, 100);
}

function parseDate(value) {
  var date = value ? new Date(value) : new Date();
  return isNaN(date.getTime()) ? new Date() : date;
}

function jsonResponse(body) {
  return ContentService
    .createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON);
}
