/**
 * Life in Hawaii ブログ コメント機能 バックエンド
 * ------------------------------------------------------------------
 * Googleスプレッドシートを「みんなのコメント置き場」にするための
 * Google Apps Script です。設定手順は README.md を見てください。
 *
 * 生徒の端末からは JSONP（?callback=...）で呼び出されます。
 */

/* ====== 先生が変更する設定 ====== */
const TEACHER_PASSCODE = 'sensei';   // コメント削除に使うパスコード（必ず変更してください）
const SHEET_NAME       = 'comments'; // 記録するシート名（自動で作られます）
const MAX_TEXT_LENGTH  = 400;        // コメント1件の最大文字数
/* ================================ */

const HEADERS = ['id', 'ts', 'room', 'cls', 'num', 'name', 'answer', 'text', 'likes', 'clientId', 'deleted'];

function doGet(e) {
  const p = (e && e.parameter) || {};
  let out;
  try {
    out = handleRequest(p);
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  const json = JSON.stringify(out);
  if (p.callback) {
    return ContentService
      .createTextOutput(p.callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  return doGet(e);
}

function handleRequest(p) {
  const room = String(p.room || 'default').slice(0, 30);
  switch (p.action) {
    case 'list':   return { ok: true, comments: listComments(room) };
    case 'add':    return { ok: true, comment: addComment(room, p) };
    case 'like':   return { ok: true, likes: likeComment(room, String(p.id), Number(p.delta) || 0) };
    case 'delete': return deleteComment(room, String(p.id), String(p.pass || ''));
    default:       return { ok: false, error: 'unknown action' };
  }
}

/* ---------- シート ---------- */
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readAll(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
}

/* ---------- 一覧 ---------- */
function listComments(room) {
  const rows = readAll(getSheet());
  const out = [];
  rows.forEach(function (r) {
    if (String(r[2]) !== room) return;
    if (String(r[10]) === '1') return;
    out.push({
      id:       String(r[0]),
      ts:       r[1] instanceof Date ? r[1].toISOString() : String(r[1]),
      cls:      String(r[3]),
      num:      String(r[4]),
      name:     String(r[5]),
      answer:   String(r[6]) === 'no' ? 'no' : 'yes',
      text:     String(r[7]),
      likes:    Number(r[8]) || 0,
      clientId: String(r[9])
    });
  });
  return out;
}

/* ---------- 追加 ---------- */
function addComment(room, p) {
  const text = String(p.text || '').slice(0, MAX_TEXT_LENGTH).trim();
  if (!text) throw new Error('empty comment');

  const c = {
    id:       String(p.id || ('m' + Date.now())),
    ts:       new Date().toISOString(),
    cls:      String(p.cls || '').slice(0, 10),
    num:      String(p.num || '').slice(0, 5),
    name:     String(p.name || '').slice(0, 20),
    answer:   String(p.answer) === 'no' ? 'no' : 'yes',
    text:     text,
    likes:    0,
    clientId: String(p.clientId || '').slice(0, 20)
  };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet();
    // 同じidが二重送信された場合は追加しない
    const rows = readAll(sh);
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === c.id) return c;
    }
    sh.appendRow([c.id, c.ts, room, c.cls, c.num, c.name, c.answer, c.text, 0, c.clientId, '']);
  } finally {
    lock.releaseLock();
  }
  return c;
}

/* ---------- いいね ---------- */
function likeComment(room, id, delta) {
  if (delta !== 1 && delta !== -1) throw new Error('bad delta');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet();
    const rows = readAll(sh);
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === id && String(rows[i][2]) === room) {
        const likes = Math.max(0, (Number(rows[i][8]) || 0) + delta);
        sh.getRange(i + 2, 9).setValue(likes);
        return likes;
      }
    }
  } finally {
    lock.releaseLock();
  }
  throw new Error('not found');
}

/* ---------- 削除（先生用） ---------- */
function deleteComment(room, id, pass) {
  if (pass !== TEACHER_PASSCODE) return { ok: false, error: 'bad passcode' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet();
    const rows = readAll(sh);
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === id && String(rows[i][2]) === room) {
        sh.getRange(i + 2, 11).setValue('1'); // 行は残したまま非表示にします
        return { ok: true };
      }
    }
  } finally {
    lock.releaseLock();
  }
  return { ok: false, error: 'not found' };
}

/* ---------- 動作確認用（エディタから実行できます） ---------- */
function testSetup() {
  const sh = getSheet();
  Logger.log('シート「' + sh.getName() + '」を準備しました。行数: ' + sh.getLastRow());
}
