/**
 * 立威個案管理｜雲端同步（Google Apps Script）
 * 使用方式：
 *   1. 在 Google 試算表 →「擴充功能」→「Apps Script」，把這整份貼上、儲存。
 *   2. 上方選單選函式「setup」→ 按「執行」→ 依指示授權。
 *   3. 執行完到「執行記錄」複製「金鑰」。
 *   4. 右上「部署」→「新增部署作業」→ 類型「網頁應用程式」
 *      執行身分：我；誰可以存取：所有人 → 部署 → 複製「網頁應用程式網址」。
 *   5. 到 app「設定 → Google 雲端同步」貼上網址和金鑰。
 */
const CHUNK = 45000; // 每格最多 5 萬字，分段存

function setup() {
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('TOKEN')) p.setProperty('TOKEN', Utilities.getUuid().replace(/-/g, '').slice(0, 24));
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('data');
  if (!sh) {
    sh = ss.insertSheet('data');
    sh.appendRow(['col', 'id', 'ts', 'deleted', 'by', 'json']);
    sh.setFrozenRows(1);
  }
  if (!p.getProperty('FOLDER')) {
    const f = DriveApp.createFolder('立威個案管理_照片');
    p.setProperty('FOLDER', f.getId());
  }
  Logger.log('金鑰：' + p.getProperty('TOKEN'));
  Logger.log('照片資料夾：https://drive.google.com/drive/folders/' + p.getProperty('FOLDER'));
}

function doGet() { return out({ ok: true, msg: '立威個案管理 雲端同步運作中' }); }

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (x) { return out({ ok: false, err: '資料格式錯誤' }); }
  const p = PropertiesService.getScriptProperties();
  if (!b.token || b.token !== p.getProperty('TOKEN')) return out({ ok: false, err: '金鑰不正確' });
  try {
    if (b.action === 'ping') return out({ ok: true, now: Date.now() });
    if (b.action === 'pull') return out(pull_(+b.since || 0));
    if (b.action === 'push') return out(push_(b.changes || [], b.by || ''));
    if (b.action === 'img') return out(img_(b));
    if (b.action === 'getimg') return out(getimg_(b.ids || []));
    return out({ ok: false, err: '未知的動作' });
  } catch (err) { return out({ ok: false, err: String(err) }); }
}

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sheet_() { return SpreadsheetApp.getActive().getSheetByName('data'); }

function pull_(since) {
  const v = sheet_().getDataRange().getValues();
  const rows = []; let max = since;
  for (let i = 1; i < v.length; i++) {
    const r = v[i]; const ts = +r[2];
    if (ts > max) max = ts;
    if (ts > since) rows.push({ col: r[0], id: String(r[1]), ts: ts, deleted: !!r[3],
      json: r[3] ? '' : r.slice(5).filter(function (x) { return x !== ''; }).map(function (x) { return String(x).slice(1); }).join('') });
  }
  return { ok: true, rows: rows, ts: max };
}

function push_(changes, by) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const sh = sheet_();
    const v = sh.getDataRange().getValues();
    const idx = {}; for (let i = 1; i < v.length; i++) idx[v[i][0] + '|' + v[i][1]] = i + 1;
    let ts = Date.now();
    changes.forEach(function (c) {
      ts++;
      const j = c.deleted ? '' : String(c.json || '');
      const parts = []; for (let k = 0; k < j.length; k += CHUNK) parts.push('~' + j.slice(k, k + CHUNK));
      let row = [c.col, String(c.id), ts, c.deleted ? 1 : '', by].concat(parts.length ? parts : ['']);
      const width = Math.max(row.length, sh.getLastColumn());
      while (row.length < width) row.push('');
      const r = idx[c.col + '|' + c.id];
      if (r) sh.getRange(r, 1, 1, width).setNumberFormat('@').setValues([row]);
      else { sh.getRange(sh.getLastRow() + 1, 1, 1, width).setNumberFormat('@').setValues([row]); idx[c.col + '|' + c.id] = sh.getLastRow(); }
    });
    return { ok: true, ts: ts };
  } finally { lock.releaseLock(); }
}

function img_(b) {
  const folder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER'));
  const m = String(b.data).match(/^data:([^;]+);base64,(.*)$/);
  if (!m) return { ok: false, err: '圖片格式錯誤' };
  const ext = (m[1].split('/')[1] || 'jpg').replace('jpeg', 'jpg');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], (b.name || '照片') + '.' + ext);
  return { ok: true, id: folder.createFile(blob).getId() };
}

function getimg_(ids) {
  const o = {};
  ids.slice(0, 20).forEach(function (id) {
    try { const bl = DriveApp.getFileById(id).getBlob(); o[id] = 'data:' + bl.getContentType() + ';base64,' + Utilities.base64Encode(bl.getBytes()); }
    catch (e) { o[id] = null; }
  });
  return { ok: true, imgs: o };
}
