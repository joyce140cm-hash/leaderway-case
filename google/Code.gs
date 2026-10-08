/**
 * 立威個案管理｜雲端同步（Google Apps Script）
 * 更新方式：整份貼上 → 儲存 → 部署 → 管理部署作業 → 編輯 → 版本選「新版本」→ 部署（網址不變）
 */
const CHUNK = 45000;    // 每格最多 5 萬字，分段存
const MIN_VER = 3;      // 低於這個版本的 app 只能讀、不能寫（避免舊版蓋掉新資料）
const KEEP_DAYS = 30;   // 每日備份保留天數
const TZ = 'Asia/Taipei';

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
    if (b.action === 'ping') return out({ ok: true, now: Date.now(), backup: dailyBackup_() });
    if (b.action === 'pull') { const r = pull_(+b.since || 0); r.backup = dailyBackup_(); return out(r); }
    if (b.action === 'push') {
      if ((+b.ver || 0) < MIN_VER) return out({ ok: false, old: true, err: 'app 版本太舊，請重新整理頁面' });
      return out(push_(b.changes || [], b.by || ''));
    }
    if (b.action === 'img') return out(img_(b));
    if (b.action === 'getimg') return out(getimg_(b.ids || []));
    return out({ ok: false, err: '未知的動作' });
  } catch (err) { return out({ ok: false, err: String(err) }); }
}

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sheet_() { return SpreadsheetApp.getActive().getSheetByName('data'); }
function rowJson_(r) { return r.slice(5).filter(function (x) { return x !== ''; }).map(function (x) { return String(x).slice(1); }).join(''); }

function pull_(since) {
  // 沒有新資料就不讀整張表（大家每 45 秒同步一次，這樣雲端很省力）
  const last = +(PropertiesService.getScriptProperties().getProperty('LAST_TS') || 0);
  if (since > 0 && last > 0 && since >= last) return { ok: true, rows: [], ts: since };
  const v = sheet_().getDataRange().getValues();
  const rows = []; let max = since;
  for (let i = 1; i < v.length; i++) {
    const r = v[i]; const ts = +r[2];
    if (ts > max) max = ts;
    if (ts > since) rows.push({ col: r[0], id: String(r[1]), ts: ts, deleted: !!r[3], json: r[3] ? '' : rowJson_(r) });
  }
  return { ok: true, rows: rows, ts: max };
}

// 帳號名單：雲端只合併、不減少。同一個帳號，角色／對應人員／權限以最新一次管理員設定為準。
const AUTH_F = ['role', 'link', 'priceEdit', 'disabled', 'setAt', 'setBy'];
function mergeUsers_(oldJson, newJson) {
  let a = [], b = [];
  try { a = JSON.parse(oldJson || '[]'); } catch (e) { a = []; }
  try { b = JSON.parse(newJson || '[]'); } catch (e) { return oldJson; }
  if (!Array.isArray(a) || !Array.isArray(b)) return newJson;
  const old = {}; a.forEach(function (u) { if (u && u.acc) old[u.acc] = u; });
  const seen = {};
  const outList = b.filter(function (u) { return u && u.acc; }).map(function (u) {
    seen[u.acc] = 1; const o = old[u.acc]; if (!o) return u;
    if ((o.setAt || 0) > (u.setAt || 0)) { AUTH_F.forEach(function (f) { if (o[f] === undefined) delete u[f]; else u[f] = o[f]; }); }
    return u;
  });
  a.forEach(function (u) { if (u && u.acc && !seen[u.acc]) outList.push(u); });
  return JSON.stringify(outList);
}

function push_(changes, by) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const sh = sheet_();
    const props = PropertiesService.getScriptProperties();
    const v = sh.getDataRange().getValues();
    const idx = {}; for (let i = 1; i < v.length; i++) idx[v[i][0] + '|' + v[i][1]] = i + 1;
    // 時間戳記一定往前（不會因為上一批很大而倒退），其他裝置才不會漏拉
    let ts = Math.max(Date.now(), +(props.getProperty('LAST_TS') || 0));
    const toRow = function (c) {
      ts++;
      const j = c.deleted ? '' : String(c.json || '');
      const parts = []; for (let k = 0; k < j.length; k += CHUNK) parts.push('~' + j.slice(k, k + CHUNK));
      return [c.col, String(c.id), ts, c.deleted ? 1 : '', by].concat(parts.length ? parts : ['']);
    };
    const updates = [], adds = [];
    changes.forEach(function (c) {
      if (c.col === 'kv' && c.id === 'users') {
        if (c.deleted) return; // 帳號名單永遠不刪
        const r0 = idx['kv|users'];
        if (r0) c.json = mergeUsers_(rowJson_(v[r0 - 1]), c.json);
      }
      const r = idx[c.col + '|' + c.id]; (r ? updates : adds).push({ c: c, r: r });
    });
    updates.forEach(function (u) { u.row = toRow(u.c); });
    adds.forEach(function (u) { u.row = toRow(u.c); });
    let width = sh.getLastColumn();
    updates.concat(adds).forEach(function (u) { if (u.row.length > width) width = u.row.length; });
    const pad = function (row) { while (row.length < width) row.push(''); return row; };
    // 已存在的列：逐列更新；新的列：一次寫入（大量匯入快很多）
    updates.forEach(function (u) { sh.getRange(u.r, 1, 1, width).setNumberFormat('@').setValues([pad(u.row)]); });
    if (adds.length) {
      const start = sh.getLastRow() + 1;
      sh.getRange(start, 1, adds.length, width).setNumberFormat('@').setValues(adds.map(function (u) { return pad(u.row); }));
    }
    props.setProperty('LAST_TS', String(ts));
    return { ok: true, ts: ts };
  } finally { lock.releaseLock(); }
}

// 每天第一次有人同步時，自動把整份試算表複製一份到「立威個案管理_備份」資料夾，保留 30 天。
function dailyBackup_() {
  const p = PropertiesService.getScriptProperties();
  const day = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const done = p.getProperty('BACKUP_DAY');
  if (done === day) return p.getProperty('BACKUP_AT') || day;
  // 先「佔位」今天（很快放開鎖），複製檔案時不鎖住，不會卡到同事上傳
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return p.getProperty('BACKUP_AT') || '';
  try { if (p.getProperty('BACKUP_DAY') === day) return p.getProperty('BACKUP_AT') || day; p.setProperty('BACKUP_DAY', day); }
  finally { lock.releaseLock(); }
  try {
    let fid = p.getProperty('BACKUP_FOLDER'), folder = null;
    try { if (fid) folder = DriveApp.getFolderById(fid); } catch (e) { folder = null; }
    if (!folder) { folder = DriveApp.createFolder('立威個案管理_備份'); p.setProperty('BACKUP_FOLDER', folder.getId()); }
    DriveApp.getFileById(SpreadsheetApp.getActive().getId()).makeCopy('立威個案管理_備份_' + day, folder);
    const limit = Date.now() - KEEP_DAYS * 864e5;
    const it = folder.getFiles();
    while (it.hasNext()) { const f = it.next(); if (f.getDateCreated().getTime() < limit) f.setTrashed(true); }
    const at = Utilities.formatDate(new Date(), TZ, 'M/d HH:mm');
    p.setProperty('BACKUP_AT', at);
    return at;
  } catch (e) { p.setProperty('BACKUP_DAY', done || ''); return p.getProperty('BACKUP_AT') || ''; } // 失敗 → 下次再試
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
