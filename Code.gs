/** NeoRhythm Web - GAS backend (Spreadsheet = DB). Deploy: ウェブアプリ / 自分として実行 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('NeoRhythm Web')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
