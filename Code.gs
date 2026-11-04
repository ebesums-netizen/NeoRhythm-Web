/**
 * NeoRhythm Web - GAS API (Spreadsheet = DB)
 * 初回のみ: setToken() の合言葉を書き換えて1回実行 → デプロイ(自分として実行 / アクセス: 全員)
 */
function setToken() { PropertiesService.getScriptProperties().setProperty('TOKEN', 'ここに合言葉を入れる'); }

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet() { return json_({ ok: true, name: 'NeoRhythm API' }); }
function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents), tok = PropertiesService.getScriptProperties().getProperty('TOKEN');
    if (!tok || p.token !== tok) return json_({ ok: false, error: 'auth' });
    const fn = { getSongs: getSongs, saveSong: saveSong, saveChart: saveChart, getLogs: getLogs, saveLog: saveLog }[p.fn];
    if (!fn) return json_({ ok: false, error: 'fn' });
    return json_({ ok: true, data: fn(p.arg) });
  } catch (err) { return json_({ ok: false, error: String(err) }); }
}

function sheet_(name) {
  const p = PropertiesService.getScriptProperties();
  let id = p.getProperty('SSID'), ss;
  if (id) ss = SpreadsheetApp.openById(id);
  else { ss = SpreadsheetApp.create('NeoRhythm DB'); p.setProperty('SSID', ss.getId()); }
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

// ---- REST相当 (google.script.run から呼ぶ) ----
function getSongs() { return list_('songs', 500); }
function saveSong(json) { const o = JSON.parse(json); return upsert_('songs', o.id, json); }   // POST /api/songs
function saveChart(json) { const o = JSON.parse(json); return upsert_('charts', o.songId + ':' + o.difficulty, json); } // POST /api/charts
function getLogs() { return list_('logs', 200); }
function saveLog(json) { const o = JSON.parse(json); return upsert_('logs', String(o.at), json); }
