/**
 * フラッシュ暗算アプリの記録先となる Google Apps Script です。
 *
 * スプレッドシートに紐づけて「ウェブアプリ」としてデプロイし、
 * 発行された /exec の URL をアプリの「記録設定」に貼り付けてください。
 * 手順は README.md を参照してください。
 */

var HEADERS = ['日時', '正誤', '桁数', '口数', '4桁回数', '秒数', '入力した答', '正答'];

// 列を足す前の見出し。既存シートを見分けて移行するために使う。
var HEADERS_V1 = ['日時', '正誤', '桁数', '口数', '入力した答', '正答'];
var HEADERS_V2 = ['日時', '正誤', '桁数', '口数', '秒数', '入力した答', '正答'];

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
      Number(data.fourCount),
      Number(data.seconds),
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
  if (sheet) {
    migrateSheet(sheet);
    return sheet;
  }

  sheet = ss.insertSheet(name);
  sheet.appendRow(HEADERS);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.getRange('A:A').setNumberFormat('yyyy/MM/dd HH:mm:ss');
  sheet.setColumnWidth(1, 160);
  return sheet;
}

/**
 * 古い見出しのシートを、今の列構成へ順に引き上げます。
 * 旧版から続けて呼ぶので、6 列のシートも 1 回で 8 列になります。
 */
function migrateSheet(sheet) {
  insertColumnIfHeaderMatches(sheet, HEADERS_V1, 5, '秒数');      // 6 列 → 7 列
  insertColumnIfHeaderMatches(sheet, HEADERS_V2, 5, '4桁回数');   // 7 列 → 8 列
}

/**
 * 見出しが expected と完全に一致するときだけ、column の位置へ空の列を差し込みます。
 * 一致しなければ何もしないので、何度呼んでも、無関係なシートに対しても安全です。
 * 既存の行は右にずれるだけなので、見出しとの対応は保たれます。
 */
function insertColumnIfHeaderMatches(sheet, expected, column, label) {
  if (sheet.getLastColumn() !== expected.length) { return; }

  var header = sheet.getRange(1, 1, 1, expected.length).getValues()[0];
  for (var i = 0; i < expected.length; i++) {
    if (String(header[i]).trim() !== expected[i]) { return; }
  }

  sheet.insertColumnBefore(column);
  sheet.getRange(1, column).setValue(label).setFontWeight('bold');
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
