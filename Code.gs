/**
 * NeoRhythm Web - GAS API (ログイン方式: ID・パスワード管理システム + 端末数の制限)
 *
 * 【初期設定】
 *  1. データ用スプレッドシート(NeoRhythm DB)を開き、メニュー「拡張機能 → Apps Script」でこのコードを貼る
 *     (スプレッドシートに紐づいたスクリプトにすること。ユーザー表の自動ハッシュ化が動くため)
 *  2. 関数 initUsers を1回実行(「users」シートが作られる。権限の承認が出たら許可)
 *  3. users シートの A列にID、B列にパスワードを入力 → 自動でハッシュ化される(動かない場合はメニュー「NeoRhythm → ユーザーを反映」)
 *  4. 「デプロイ → 新しいデプロイ → ウェブアプリ」(実行ユーザー=自分 / アクセスできるユーザー=全員)
 *     コードを変更したら必ず「デプロイを管理 → 新バージョン」で更新する
 *  5. デプロイURL(…/exec)を config.js の API に設定する
 *
 * users シートの列: A=ID / B=パスワード(新規入力欄) / C=ハッシュ / D=ソルト / E=有効期限(yyyy-MM-dd、空欄=無期限)
 *                   F=状態(有効/停止、空欄=有効) / G=最終ログイン / H=端末数の上限(空欄=2) / I=登録端末(自動)
 */
const USERS = 'users', TZ = 'Asia/Tokyo', TOKEN_DAYS = 90, DEFAULT_DEVICES = 2;

/* ---------- スプレッドシート ---------- */
function ss_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (ss) return ss;
  const p = PropertiesService.getScriptProperties(), id = p.getProperty('SSID');
  if (id) return SpreadsheetApp.openById(id);
  ss = SpreadsheetApp.create('NeoRhythm DB'); p.setProperty('SSID', ss.getId());
  return ss;
}
function usersSheet_() {
  const ss = ss_();
  let s = ss.getSheetByName(USERS);
  if (!s) {
    s = ss.insertSheet(USERS);
    s.getRange('A:I').setNumberFormat('@');
    s.getRange(1, 1, 1, 9).setValues([['ID', 'パスワード（新規入力欄）', 'ハッシュ', 'ソルト', '有効期限（空欄=無期限）', '状態（有効/停止）', '最終ログイン', '端末数の上限（空欄=' + DEFAULT_DEVICES + '）', '登録端末（自動。消すとリセット）']]);
  } else if (!String(s.getRange(1, 8).getValue())) { // 他アプリのusersシートをコピーした場合に、端末管理の列を補う
    s.getRange('H:I').setNumberFormat('@');
    s.getRange(1, 8, 1, 2).setValues([['端末数の上限（空欄=' + DEFAULT_DEVICES + '）', '登録端末（自動。消すとリセット）']]);
  }
  return s;
}
function initUsers() { usersSheet_(); } // 最初に1回だけ実行

/* ---------- ユーザー管理 ---------- */
function hash_(salt, pw) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pw, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
// usersシートのA列(ID)・B列(パスワード)に入力 → 自動でハッシュ化(スマホ入力でも動作)
function onEdit(e) {
  try {
    const r = e.range, s = r.getSheet();
    if (s.getName() !== USERS || r.getColumn() > 2) return;
    for (let i = Math.max(2, r.getRow()); i <= r.getLastRow(); i++) {
      const id = String(s.getRange(i, 1).getValue()).trim(), pw = String(s.getRange(i, 2).getValue());
      if (!id || pw === '') continue;
      const salt = Utilities.getUuid();
      s.getRange(i, 3, 1, 2).setValues([[hash_(salt, pw), salt]]);
      s.getRange(i, 2).clearContent();
      if (!s.getRange(i, 6).getValue()) s.getRange(i, 6).setValue('有効');
    }
  } catch (x) {}
}
function applyUsers() { // onEditが効かない場合の手動反映
  const s = usersSheet_(), n = s.getLastRow() - 1;
  if (n < 1) return;
  const r = s.getRange(2, 1, n, 7), v = r.getValues();
  let c = 0;
  v.forEach(x => {
    if (String(x[0]).trim() && String(x[1]) !== '') {
      const salt = Utilities.getUuid();
      x[3] = salt; x[2] = hash_(salt, String(x[1])); x[1] = '';
      if (!x[5]) x[5] = '有効';
      c++;
    }
  });
  r.setValues(v);
  ss_().toast(c + '件のパスワードを反映しました。');
}
function resetDevices() { // 選択中の行(usersシート)の登録端末をすべて解除する
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== USERS) { ss_().toast('usersシートのユーザー行を選択してから実行してください。'); return; }
  const rg = sh.getActiveRange();
  let c = 0;
  for (let i = Math.max(2, rg.getRow()); i <= rg.getLastRow(); i++) { sh.getRange(i, 9).clearContent(); c++; }
  ss_().toast(c + '行の登録端末を解除しました。');
}
function onOpen() {
  try { SpreadsheetApp.getUi().createMenu('NeoRhythm').addItem('ユーザーを反映', 'applyUsers').addItem('選択行の端末をリセット', 'resetDevices').addToUi(); } catch (e) {}
}

function normDate_(e) {
  if (e instanceof Date) return Utilities.formatDate(e, TZ, 'yyyy-MM-dd');
  const m = String(e || '').match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : '';
}
function parseDevs_(v) {
  try { const a = JSON.parse(String(v || '[]')); return Array.isArray(a) ? a.filter(x => x && x.d) : []; } catch (e) { return []; }
}
function user_(id) {
  if (!id) return null;
  const s = usersSheet_(), n = s.getLastRow() - 1;
  if (n < 1) return null;
  const v = s.getRange(2, 1, n, 9).getValues();
  for (let i = 0; i < n; i++) {
    const x = v[i];
    if (String(x[0]).trim() === id && x[2]) {
      return { row: i + 2, hash: String(x[2]), salt: String(x[3]), exp: normDate_(x[4]), st: String(x[5] || '').trim(), max: Number(x[7]) > 0 ? Number(x[7]) : DEFAULT_DEVICES, devs: parseDevs_(x[8]) };
    }
  }
  return null;
}
function status_(u) {
  if (u.st && u.st !== '有効') return 'このIDは現在ご利用いただけません。';
  if (u.exp && Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd') > u.exp) return 'ご利用期限（' + u.exp + '）を過ぎています。';
  return '';
}

/* ---------- トークン ---------- */
function secret_() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRET', s); }
  return s;
}
function sig_(payload, salt) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payload + '|' + salt, secret_()));
}
function makeToken_(u, id, dev) { // IDの有効期限(なければ90日)まで有効。パスワード変更(ソルト更新)・端末解除で失効
  let exp = Date.now() + TOKEN_DAYS * 86400000;
  if (u.exp) { const e = new Date(u.exp + 'T23:59:59+09:00').getTime(); if (e < exp) exp = e; }
  const payload = Utilities.base64EncodeWebSafe(JSON.stringify({ i: id, e: exp, d: dev }));
  return payload + '.' + sig_(payload, u.salt);
}
function auth_(t) {
  const bad = { err: 'ログインし直してください。' };
  try {
    const parts = String(t || '').split('.');
    if (parts.length !== 2) return bad;
    const o = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString());
    if (!o.i || !o.d || Date.now() > o.e) return bad;
    const u = user_(o.i);
    if (!u || sig_(parts[0], u.salt) !== parts[1]) return bad;
    const st = status_(u);
    if (st) return { err: st };
    if (!u.devs.some(x => x.d === o.d)) return { err: 'この端末の登録が解除されています。ログインし直してください。' };
    return { id: o.i, u: u };
  } catch (e) { return bad; }
}

/* ---------- 端末の登録(同時に使える端末数の制限) ---------- */
function regDevice_(row, dev, name) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const s = usersSheet_(), max = Number(s.getRange(row, 8).getValue()) > 0 ? Number(s.getRange(row, 8).getValue()) : DEFAULT_DEVICES;
    const cell = s.getRange(row, 9), list = parseDevs_(cell.getValue());
    const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), i = list.findIndex(x => x.d === dev);
    if (i >= 0) { list[i].at = now; if (name) list[i].n = name; }
    else {
      if (list.length >= max) return { err: 'ご利用できる端末数（' + max + '台）に達しています。別の端末では、すでに使用中の端末を解除してからログインしてください。（解除は管理者にご連絡ください）' };
      list.push({ d: dev, n: name, at: now });
    }
    cell.setValue(JSON.stringify(list));
    return { n: list.length, max: max };
  } finally { lock.releaseLock(); }
}

/* ---------- ログイン ---------- */
function login_(id, pw, dev, name) {
  id = String(id || '').trim(); pw = String(pw || ''); dev = String(dev || '');
  if (!/^[A-Za-z0-9_-]{8,64}$/.test(dev)) return { err: '端末の情報を確認できません。ページを再読み込みしてください。' };
  const cache = CacheService.getScriptCache(), fk = 'f:' + id;
  if (Number(cache.get(fk) || 0) >= 5) return { err: 'ログインの失敗が続いたため、15分ほどお待ちください。' };
  const u = user_(id);
  if (!u || u.hash !== hash_(u.salt, pw)) {
    cache.put(fk, String(Number(cache.get(fk) || 0) + 1), 900);
    return { err: 'IDまたはパスワードが違います。' };
  }
  const st = status_(u);
  if (st) return { err: st };
  const r = regDevice_(u.row, dev, String(name || '').slice(0, 30)); // 端末数オーバーは、パスワード失敗には数えない
  if (r.err) return r;
  cache.remove(fk);
  usersSheet_().getRange(u.row, 7).setValue(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'));
  return { token: makeToken_(u, id, dev), exp: u.exp };
}

/* ---------- データ(スプレッドシート = DB) ---------- */
function sheet_(name) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(['id', 'json', 'at']); }
  return sh;
}
function upsert_(name, id, json) {
  const sh = sheet_(name), ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().flat();
  const i = ids.indexOf(id);
  if (i > 0) sh.getRange(i + 1, 1, 1, 3).setValues([[id, json, new Date()]]);
  else sh.appendRow([id, json, new Date()]);
  return true;
}
function list_(name, limit) {
  const sh = sheet_(name), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const from = Math.max(2, sh.getLastRow() - limit + 1);
  return sh.getRange(from, 2, sh.getLastRow() - from + 1, 1).getValues().flat();
}
function getSongs() { return list_('songs', 500); }
function saveSong(json) { const o = JSON.parse(json); return upsert_('songs', o.id, json); }
function saveChart(json) { const o = JSON.parse(json); return upsert_('charts', o.songId + ':' + o.difficulty, json); }
function getLogs(uid) { // プレイ記録はユーザーごとに分けて返す
  return list_('logs', 1000).filter(j => { try { return JSON.parse(j).uid === uid; } catch (e) { return false; } }).slice(-200);
}
function saveLog(json, uid) { const o = JSON.parse(json); o.uid = uid; return upsert_('logs', uid + ':' + o.at, JSON.stringify(o)); }

/* ---------- API入口 ---------- */
function api_(b) {
  if (b.a === 'login') return login_(b.id, b.pw, b.dev, b.dn);
  const a = auth_(b.t);
  if (a.err) return { err: a.err };
  switch (b.a) {
    case 'me': return { id: a.id, exp: a.u.exp, ads: false, dev: { n: a.u.devs.length, max: a.u.max } }; // ads:false = 有効なIDなので広告なし
    case 'getSongs': return { data: getSongs() };
    case 'saveSong': return { data: saveSong(b.arg) };
    case 'saveChart': return { data: saveChart(b.arg) };
    case 'getLogs': return { data: getLogs(a.id) };
    case 'saveLog': return { data: saveLog(b.arg, a.id) };
    default: return { err: '不正なリクエストです。' };
  }
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doPost(e) {
  let r;
  try { r = api_(JSON.parse(e.postData.contents)); } catch (x) { r = { err: 'エラーが発生しました。' }; }
  return out_(r);
}
function doGet() { return out_({ ok: true, name: 'NeoRhythm API' }); }
