// ダンス霊幻のコメント受付（Google Apps Script）
// スプレッドシートの「comments」シートに溜まる。C列「表示」にチェックを入れたものだけ公開される。
const SHEET = 'comments';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET);
  if (!s) {
    s = ss.insertSheet(SHEET);
    s.appendRow(['日時', 'コメント', '表示']);
    s.setFrozenRows(1);
  }
  return s;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// 送信：未承認（チェックなし）で追加
function doPost(e) {
  let text = '';
  try { text = String(JSON.parse(e.postData.contents).text || ''); } catch (_) {}
  text = text.replace(/[\r\n\t]/g, ' ').trim().slice(0, 40);
  if (!text) return json_({ ok: false });
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const s = sheet_();
    const row = s.getLastRow() + 1;
    s.getRange(row, 2).setNumberFormat('@'); // =で始まる文字も数式にしない（先に書式を文字列に）
    s.getRange(row, 1, 1, 2).setValues([[new Date(), text]]);
    s.getRange(row, 3).insertCheckboxes();
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

// 取得：チェックが入ったコメントだけ返す
function doGet() {
  const v = sheet_().getDataRange().getValues().slice(1);
  return json_({ comments: v.filter(r => r[2] === true).map(r => String(r[1])).slice(-200) });
}
