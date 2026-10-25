/**
 * 3M Calendar – backend (Google Apps Script)
 * ------------------------------------------------------------
 * วางไฟล์นี้ใน Apps Script ของไฟล์ "3m_calendar ปีงบ70_ต้นทาง"
 * (Extensions > Apps Script) แล้วรัน setup() หนึ่งครั้ง จากนั้น Deploy เป็น Web app
 *
 * - ข้อมูลจริงเก็บในไฟล์ "ต้นทาง" (sheets: events, months, settings, logs, snapshots)
 * - ไฟล์ "แสดงผล" ถูกวาดใหม่จากข้อมูลต้นทางทุกครั้งที่มีการแก้ไข (อ่านอย่างเดียว)
 * - กด SAVE แล้วเก็บไฟล์ Excel ในโฟลเดอร์ Backup ไว้ 20 ไฟล์ล่าสุด
 */

const CFG = {
  SOURCE_ID: '1gB7hh2Zpw7u8_pHcpHujVlfb6JkmGSNLVKx182hVU_g',
  DISPLAY_ID: '10Bj4UuJ6r86p9_bgYRaeUUCsGELGtWt_ktB-4mF3dLI',
  BACKUP_FOLDER_ID: '1czRN-LXGfumODTUmZcyovqM4QQSkAjUj',
  DEFAULT_CODE: 'operation.slide',   // รหัสเปิดใช้งานเริ่มต้น (เปลี่ยนได้ด้วย setAccessCode)
  KEEP_BACKUPS: 20,
  KEEP_SNAPS: 20,
  SESSION_SEC: 3600,                 // ใส่รหัสแล้วใช้ได้ 1 ชั่วโมง
  LOCK_MS: 90000,                    // สิทธิ์แก้ไขหมดอายุถ้าไม่ต่ออายุภายใน 90 วินาที
  PRESENCE_SEC: 25,
  LOG_DAYS: 365,
  TZ: 'Asia/Bangkok',
  DISPLAY_SHEET: 'แผน 3 เดือน'
};

const EV_COLS = ['id', 'month', 'start', 'end', 'lane', 'rows', 'text', 'type', 'status', 'size', 'bold', 'italic', 'underline', 'color', 'fill', 'ts', 'by'];
const EV_NUM = { start: 1, end: 1, lane: 1, rows: 1, size: 1, ts: 1 };
const EV_BOOL = { bold: 1, italic: 1 };

/* =========================== HTTP =========================== */

function doGet() {
  return json_({ ok: true, app: '3mcalendar', time: new Date().toISOString() });
}

function doPost(e) {
  let req = {};
  try {
    req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = API[req.action];
    if (!fn) throw err_('bad_action', 'ไม่รู้จักคำสั่ง ' + req.action);
    const sess = req.action === 'login' ? null : auth_(req.token);
    return json_(fn(req, sess) || { ok: true });
  } catch (x) {
    const code = (x && x.code) || 'server';
    return json_({ error: { code: code, message: (x && x.message) || String(x) } });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function err_(code, message) { const e = new Error(message); e.code = code; return e; }

/* =========================== API =========================== */

const API = {
  login: function (r) {
    const name = String(r.name || '').trim().slice(0, 60);
    if (!name) throw err_('no_name', 'กรุณาใส่ชื่อผู้ใช้งาน');
    const want = PropertiesService.getScriptProperties().getProperty('CODE_HASH') || sha256_(CFG.DEFAULT_CODE);
    if (sha256_(String(r.code || '').trim()) !== want) { Utilities.sleep(700); throw err_('bad_code', 'รหัสไม่ถูกต้อง'); }
    const token = Utilities.getUuid().replace(/-/g, '');
    CacheService.getScriptCache().put('tok:' + token, JSON.stringify({ name: name }), CFG.SESSION_SEC);
    return { token: token, name: name, exp: Date.now() + CFG.SESSION_SEC * 1000 };
  },

  load: function () {
    return { ver: ver_(), state: readState_(), lock: lockGet_() };
  },

  poll: function (r, s) {
    const tab = String(r.tab || '').slice(0, 40);
    if (tab) presencePut_(tab, s.name, r.presence || {});
    return { ver: ver_(), lock: lockGet_(), peers: presenceAll_(), now: Date.now() };
  },

  leave: function (r) {
    const tab = String(r.tab || '');
    if (tab) CacheService.getScriptCache().remove('pres:' + tab);
    lockRelease_(tab);
    return { ok: true };
  },

  lock: function (r, s) {
    const tab = String(r.tab || '');
    return withLock_(function () {
      const now = Date.now(), cur = lockGet_();
      const busy = cur && cur.until > now && cur.tab !== tab;
      if (r.mode === 'release') { if (cur && cur.tab === tab) lockSet_(null); return { ok: true, lock: null }; }
      if (busy) return { ok: false, lock: cur };
      if (r.mode === 'renew' && !(cur && cur.tab === tab)) return { ok: false, lock: cur };
      const v = { tab: tab, token: r.token, by: s.name, since: (cur && cur.tab === tab) ? cur.since : now, until: now + CFG.LOCK_MS };
      lockSet_(v);
      return { ok: true, lock: lockPublic_(v) };
    });
  },

  write: function (r, s) {
    return withLock_(function () {
      requireLock_(r);
      const st = readState_();
      const touched = { events: false, months: false, settings: false };
      const logs = [];
      (r.ops || []).forEach(function (o) {
        if (o.c === 'events') {
          const id = String(o.id || '');
          if (!/^[A-Za-z0-9_\-]{1,60}$/.test(id)) throw err_('bad_id', 'รหัสภารกิจไม่ถูกต้อง');
          if (o.op === 'delete') delete st.events[id];
          else if (o.op === 'update') { if (st.events[id]) st.events[id] = Object.assign(st.events[id], cleanEv_(o.data), { by: s.name }); }
          else st.events[id] = Object.assign(cleanEv_(o.data), { id: id, by: s.name });
          touched.events = true;
        } else if (o.c === 'months') {
          st.months[String(o.id)] = { lanes: Math.max(1, Math.min(60, Number(o.data && o.data.lanes) || 1)) };
          touched.months = true;
        } else if (o.c === 'meta') {
          st.settings = o.data || {};
          touched.settings = true;
        } else if (o.c === 'logs') {
          const d = o.data || {};
          logs.push([Number(d.at) || Date.now(), s.name, String(d.action || '').slice(0, 80), String(d.detail || '').slice(0, 600)]);
        }
      });
      writeState_(st, touched);
      if (logs.length) appendLogs_(logs);
      const v = bump_();
      return { ok: true, ver: v };
    });
  },

  save: function (r, s) {
    return withLock_(function () {
      requireLock_(r);
      const st = readState_();
      st.settings = Object.assign({}, st.settings, { updated: String(r.iso || '') });
      writeState_(st, { settings: true });
      const at = Date.now();
      const snap = { id: 's' + at, at: at, label: String(r.label || ''), by: s.name, undoneAt: 0 };
      snapAppend_(snap, { events: objValues_(st.events), months: st.months, settings: st.settings });
      snapPrune_();
      appendLogs_([[at, s.name, 'SAVE', 'Update ' + snap.label]]);
      let file = null;
      if (r.xlsx) file = backupExcel_(String(r.filename || ('backup_' + at)) + '.xlsx', r.xlsx);
      bump_();
      return { ok: true, snap: snap, backup: file };
    });
  },

  lastSnap: function (r, s) {
    let snap = snapList_(1, false)[0] || null;
    if (!snap) {
      withLock_(function () {
        if (snapList_(1, false).length) return;
        const st = readState_(), at = Date.now();
        snap = { id: 's' + at, at: at, label: Utilities.formatDate(new Date(), CFG.TZ, 'dd_MM_') + String(Number(Utilities.formatDate(new Date(), CFG.TZ, 'yyyy')) + 543).slice(-2), by: 'ฉบับเริ่มต้น', undoneAt: 0 };
        snapAppend_(snap, { events: objValues_(st.events), months: st.months, settings: st.settings });
      });
      snap = snapList_(1, false)[0] || snap;
    }
    const since = Math.max(snap.at || 0, snap.undoneAt || 0) + 2000;
    return { snap: snap, dirty: logsSince_(since) };
  },

  listSnaps: function (r) {
    return { list: snapList_(Math.min(20, Number(r.limit) || 20), true) };
  },

  restore: function (r, s) {
    return withLock_(function () {
      requireLock_(r);
      const snap = snapGet_(String(r.id || ''));
      if (!snap) throw err_('not_found', 'ไม่พบฉบับที่บันทึก');
      const st = { events: {}, months: snap.data.months || {}, settings: snap.data.settings || {} };
      (snap.data.events || []).forEach(function (e) { st.events[e.id] = e; });
      writeState_(st, { events: true, months: true, settings: true });
      const at = Date.now();
      if (r.how === 'undo') {
        snapSetUndone_(snap.id, at);
        appendLogs_([[at, s.name, 'ยกเลิกการแก้ไข', 'ย้อนกลับไปเหมือน SAVE ล่าสุด (Update ' + snap.label + (snap.by ? ' โดย ' + snap.by : '') + ')']]);
      } else {
        appendLogs_([[at, s.name, 'เรียกคืนฉบับ', 'Update ' + snap.label + (snap.by ? ' (บันทึกโดย ' + snap.by + ')' : '')]]);
      }
      bump_();
      return { ok: true, at: at };
    });
  },

  listLogs: function (r) {
    return { list: logsList_(Math.min(1000, Number(r.limit) || 300)) };
  },

  render: function () {
    const l = LockService.getScriptLock();
    if (!l.tryLock(1000)) { setDirty_(true); return { ok: false, queued: true }; }
    try { renderDisplay_(); } finally { l.releaseLock(); }
    return { ok: true };
  }
};

/* =========================== auth / lock =========================== */

function auth_(token) {
  const raw = token && CacheService.getScriptCache().get('tok:' + token);
  if (!raw) throw err_('auth', 'หมดเวลาใช้งาน กรุณาใส่รหัสใหม่');
  return JSON.parse(raw);
}
function sha256_(t) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, t, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}
function withLock_(fn) {
  const l = LockService.getScriptLock();
  if (!l.tryLock(20000)) throw err_('busy', 'ระบบกำลังทำงาน ลองใหม่อีกครั้ง');
  try { return fn(); } finally { l.releaseLock(); }
}
function lockGet_() {
  const raw = PropertiesService.getScriptProperties().getProperty('EDIT_LOCK');
  const v = raw ? JSON.parse(raw) : null;
  return v && v.until > Date.now() ? lockPublic_(v) : null;
}
function lockRaw_() {
  const raw = PropertiesService.getScriptProperties().getProperty('EDIT_LOCK');
  return raw ? JSON.parse(raw) : null;
}
function lockPublic_(v) { return v ? { tab: v.tab, by: v.by, since: v.since, until: v.until } : null; }
function lockSet_(v) {
  const p = PropertiesService.getScriptProperties();
  if (v) p.setProperty('EDIT_LOCK', JSON.stringify(v)); else p.deleteProperty('EDIT_LOCK');
}
function lockRelease_(tab) {
  withLock_(function () { const c = lockRaw_(); if (c && c.tab === tab) lockSet_(null); });
}
function requireLock_(r) {
  const c = lockRaw_();
  if (!c || c.until < Date.now() || c.tab !== r.tab || c.token !== r.token)
    throw err_('not_lock_holder', 'คุณไม่ได้ถือสิทธิ์แก้ไขอยู่ อาจมีผู้อื่นกำลังแก้ไข');
  c.until = Date.now() + CFG.LOCK_MS; lockSet_(c);
}

/* =========================== presence =========================== */

function presencePut_(tab, name, presence) {
  const c = CacheService.getScriptCache();
  const p = { tab: tab, name: name, presence: sanitizePresence_(presence, name), at: Date.now() };
  c.put('pres:' + tab, JSON.stringify(p), CFG.PRESENCE_SEC);
  const idx = JSON.parse(c.get('presidx') || '[]');
  if (idx.indexOf(tab) < 0) {
    const l = LockService.getScriptLock();
    if (l.tryLock(1500)) {
      try {
        const cur = JSON.parse(c.get('presidx') || '[]');
        if (cur.indexOf(tab) < 0) cur.push(tab);
        c.put('presidx', JSON.stringify(cur.slice(-60)), 21600);
      } finally { l.releaseLock(); }
    }
  }
}
function sanitizePresence_(p, name) {
  return {
    name: String(p.name || name || '').slice(0, 60),
    mode: p.mode === 'edit' ? 'edit' : 'view',
    ev: p.ev ? String(p.ev).slice(0, 60) : null,
    cell: p.cell ? String(p.cell).slice(0, 30) : null
  };
}
function presenceAll_() {
  const c = CacheService.getScriptCache();
  const idx = JSON.parse(c.get('presidx') || '[]');
  if (!idx.length) return [];
  const got = c.getAll(idx.map(function (t) { return 'pres:' + t; }));
  const out = [];
  idx.forEach(function (t) { const v = got['pres:' + t]; if (v) out.push(JSON.parse(v)); });
  return out;
}

/* =========================== state in the source sheet =========================== */

function src_() { return SpreadsheetApp.openById(CFG.SOURCE_ID); }
function sheet_(name, header) {
  const ss = src_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (header) {
      sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#D6DCE4');
      sh.setFrozenRows(1);
    }
    sh.getRange(1, 1, sh.getMaxRows(), Math.max(header ? header.length : 1, 1)).setNumberFormat('@');
  }
  return sh;
}
function ver_() { return Number(PropertiesService.getScriptProperties().getProperty('VER') || '1'); }
function bump_() {
  const p = PropertiesService.getScriptProperties();
  const v = ver_() + 1; p.setProperty('VER', String(v)); setDirty_(true);
  return v;
}
function setDirty_(on) { PropertiesService.getScriptProperties().setProperty('DISPLAY_DIRTY', on ? '1' : '0'); }

function readState_() {
  const cache = CacheService.getScriptCache(), key = 'state:' + ver_();
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  const st = { events: {}, months: {}, settings: {} };
  const ev = sheet_('events', EV_COLS).getDataRange().getValues();
  for (let i = 1; i < ev.length; i++) {
    const row = ev[i]; if (!row[0]) continue;
    const o = {};
    EV_COLS.forEach(function (k, j) {
      let v = row[j];
      if (EV_NUM[k]) v = Number(v) || 0;
      else if (EV_BOOL[k]) v = v === true || String(v).toLowerCase() === 'true';
      else v = v == null ? '' : String(v);
      o[k] = v;
    });
    st.events[o.id] = o;
  }
  const mo = sheet_('months', ['month', 'lanes']).getDataRange().getValues();
  for (let i = 1; i < mo.length; i++) if (mo[i][0]) st.months[String(mo[i][0])] = { lanes: Number(mo[i][1]) || 1 };
  const se = sheet_('settings', ['key', 'value']).getDataRange().getValues();
  for (let i = 1; i < se.length; i++) if (se[i][0]) { try { st.settings[String(se[i][0])] = JSON.parse(se[i][1]); } catch (x) { st.settings[String(se[i][0])] = se[i][1]; } }
  try { cache.put(key, JSON.stringify(st), 600); } catch (x) { /* too large to cache */ }
  return st;
}

function writeState_(st, touched) {
  if (touched.events) {
    const sh = sheet_('events', EV_COLS);
    const ids = Object.keys(st.events).sort(function (a, b) {
      const x = st.events[a], y = st.events[b];
      return (x.month > y.month ? 1 : x.month < y.month ? -1 : 0) || (x.lane - y.lane) || (x.start - y.start);
    });
    const rows = [EV_COLS].concat(ids.map(function (id) { const e = st.events[id]; return EV_COLS.map(function (k) { const v = e[k]; return v == null ? '' : String(v); }); }));
    rewrite_(sh, rows);
  }
  if (touched.months) {
    rewrite_(sheet_('months', ['month', 'lanes']), [['month', 'lanes']].concat(Object.keys(st.months).sort().map(function (k) { return [k, String(st.months[k].lanes)]; })));
  }
  if (touched.settings) {
    rewrite_(sheet_('settings', ['key', 'value']), [['key', 'value']].concat(Object.keys(st.settings).sort().map(function (k) { return [k, JSON.stringify(st.settings[k])]; })));
  }
}
function rewrite_(sh, rows) {
  const w = rows[0].length;
  const old = sh.getLastRow();
  if (old > rows.length) sh.getRange(rows.length + 1, 1, old - rows.length, Math.max(w, sh.getLastColumn())).clearContent();
  if (sh.getMaxRows() < rows.length) sh.insertRowsAfter(sh.getMaxRows(), rows.length - sh.getMaxRows());
  const rg = sh.getRange(1, 1, rows.length, w);
  rg.setNumberFormat('@');
  rg.setValues(rows);
}
function cleanEv_(d) {
  d = d || {};
  const o = {};
  EV_COLS.forEach(function (k) {
    if (k === 'id' || k === 'by' || !(k in d)) return;
    let v = d[k];
    if (EV_NUM[k]) v = Number(v) || 0;
    else if (EV_BOOL[k]) v = !!v;
    else v = String(v == null ? '' : v).slice(0, k === 'text' ? 500 : 40);
    o[k] = v;
  });
  return o;
}
function objValues_(o) { return Object.keys(o).map(function (k) { return o[k]; }); }

/* =========================== logs =========================== */

function appendLogs_(rows) {
  const sh = sheet_('logs', ['at', 'time', 'by', 'action', 'detail']);
  const out = rows.map(function (r) { return [String(r[0]), Utilities.formatDate(new Date(r[0]), CFG.TZ, 'yyyy-MM-dd HH:mm:ss'), r[1], r[2], r[3]]; });
  sh.getRange(sh.getLastRow() + 1, 1, out.length, 5).setNumberFormat('@').setValues(out);
}
function logsList_(n) {
  const sh = sheet_('logs', ['at', 'time', 'by', 'action', 'detail']);
  const last = sh.getLastRow(); if (last < 2) return [];
  const from = Math.max(2, last - n + 1);
  return sh.getRange(from, 1, last - from + 1, 5).getValues().reverse().map(function (r) { return { at: Number(r[0]), by: String(r[2]), action: String(r[3]), detail: String(r[4]) }; });
}
function logsSince_(t) {
  const sh = sheet_('logs', ['at', 'time', 'by', 'action', 'detail']);
  const last = sh.getLastRow(); if (last < 2) return false;
  const from = Math.max(2, last - 50);
  return sh.getRange(from, 1, last - from + 1, 1).getValues().some(function (r) { return Number(r[0]) > t; });
}
function pruneLogs() {
  const sh = sheet_('logs', ['at', 'time', 'by', 'action', 'detail']);
  const last = sh.getLastRow(); if (last < 2) return;
  const lim = Date.now() - CFG.LOG_DAYS * 864e5;
  const col = sh.getRange(2, 1, last - 1, 1).getValues();
  let n = 0; while (n < col.length && Number(col[n][0]) < lim) n++;
  if (n) sh.deleteRows(2, n);
}

/* =========================== snapshots (ฉบับที่ SAVE) =========================== */

const SNAP_HEAD = ['id', 'at', 'label', 'by', 'undoneAt', 'data'];
const CHUNK = 45000;
function snapSheet_() { return sheet_('snapshots', SNAP_HEAD); }
function snapAppend_(meta, data) {
  const sh = snapSheet_();
  const js = JSON.stringify(data), parts = [];
  for (let i = 0; i < js.length; i += CHUNK) parts.push(js.slice(i, i + CHUNK));
  const row = [meta.id, String(meta.at), meta.label, meta.by, String(meta.undoneAt || 0)].concat(parts);
  if (sh.getMaxColumns() < row.length) sh.insertColumnsAfter(sh.getMaxColumns(), row.length - sh.getMaxColumns());
  sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setNumberFormat('@').setValues([row]);
}
function snapRows_() {
  const sh = snapSheet_(); const last = sh.getLastRow(); if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, Math.max(6, sh.getLastColumn())).getValues().map(function (r, i) { return { row: i + 2, r: r }; });
}
function snapMeta_(x) { const r = x.r; return { id: String(r[0]), at: Number(r[1]), label: String(r[2]), by: String(r[3]), undoneAt: Number(r[4]) || 0 }; }
function snapData_(x) { try { return JSON.parse(x.r.slice(5).join('')); } catch (e) { return { events: [], months: {}, settings: {} }; } }
function snapList_(n, withData) {
  return snapRows_().sort(function (a, b) { return Number(b.r[1]) - Number(a.r[1]); }).slice(0, n)
    .map(function (x) { const m = snapMeta_(x); if (withData) { const d = snapData_(x); m.events = d.events; m.months = d.months; m.settings = d.settings; } return m; });
}
function snapGet_(id) {
  const x = snapRows_().filter(function (x) { return String(x.r[0]) === id; })[0];
  if (!x) return null;
  const m = snapMeta_(x); m.data = snapData_(x); return m;
}
function snapSetUndone_(id, at) {
  const x = snapRows_().filter(function (x) { return String(x.r[0]) === id; })[0];
  if (x) snapSheet_().getRange(x.row, 5).setValue(String(at));
}
function snapPrune_() {
  const rows = snapRows_().sort(function (a, b) { return Number(b.r[1]) - Number(a.r[1]); });
  const drop = rows.slice(CFG.KEEP_SNAPS).map(function (x) { return x.row; }).sort(function (a, b) { return b - a; });
  const sh = snapSheet_();
  drop.forEach(function (r) { sh.deleteRow(r); });
}

/* =========================== Excel backup =========================== */

function backupExcel_(filename, b64) {
  const folder = DriveApp.getFolderById(CFG.BACKUP_FOLDER_ID);
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), MimeType.MICROSOFT_EXCEL, filename);
  const f = folder.createFile(blob);
  const mine = [];
  const it = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
  while (it.hasNext()) { const x = it.next(); if (x.getName().indexOf('_รวม_ภารกิจ+แผนซ่อมบำรุง_') >= 0) mine.push(x); }
  mine.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  mine.slice(CFG.KEEP_BACKUPS).forEach(function (x) { x.setTrashed(true); });
  return { name: f.getName(), url: f.getUrl() };
}

/* =========================== "แสดงผล" sheet =========================== */

const TYPES = {
  domestic: { label: 'ภารกิจภายในประเทศ', fill: '#E2F0DB' }, a32x: { label: 'A319/A320', fill: '#B4C6E7' },
  a340: { label: 'A340', fill: '#F5B084' }, ssj: { label: 'SSJ', fill: '#FFFF00' },
  maint: { label: 'แผนซ่อมบำรุง / บ.ไม่ FMC', fill: '#FF0000', text: '#FFFFFF' },
  repair: { label: 'บ.เข้าตรวจซ่อมอื่นๆ', fill: '#FF0000' }, other: { label: 'อื่นๆ', fill: '#FFFFFF' }
};
const LEGEND = ['domestic', 'a32x', 'a340', 'ssj', 'maint', 'repair'];
const MONTHS = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
const THD = '๐๑๒๓๔๕๖๗๘๙';
function thd_(n) { return String(n).replace(/\d/g, function (d) { return THD[d]; }); }
function hex_(c) { c = String(c || '#FFFFFF').replace('#', ''); if (c.length === 3) c = c.split('').map(function (x) { return x + x; }).join(''); return '#' + c.toUpperCase(); }
function pastBlend_(h) {
  h = hex_(h).slice(1);
  const r = parseInt(h.substr(0, 2), 16), g = parseInt(h.substr(2, 2), 16), b = parseInt(h.substr(4, 2), 16);
  return '#' + [r * 0.7 + 255 * 0.3, g * 0.7, b * 0.7].map(function (v) { return ('0' + Math.round(v).toString(16)).slice(-2); }).join('').toUpperCase();
}
function richRuns_(t) { return String(t).split(/\*\*(.+?)\*\*/).map(function (x, i) { return { t: x, b: i % 2 === 1 }; }).filter(function (r) { return r.t; }); }
function noteFmt_(n, part) {
  const base = { size: n.size || 12, color: n.color || '', bg: n.bg || '', bold: !!n.bold, italic: !!n.italic, underline: n.underline || '' };
  return Object.assign(base, n[part === 'l' ? 'lf' : 'tf'] || {});
}
// TH SarabunPSK is not in Google Sheets; Sarabun at ~75% of the size looks the same
function gsz_(pt) { return Math.max(6, Math.round((Number(pt) || 12) * 0.75)); }

function renderIfDirty() {
  if (PropertiesService.getScriptProperties().getProperty('DISPLAY_DIRTY') !== '1') return;
  const l = LockService.getScriptLock();
  if (!l.tryLock(5000)) return;
  try { renderDisplay_(); } finally { l.releaseLock(); }
}
function renderDaily() { // past-day shading moves every day
  const l = LockService.getScriptLock(); if (!l.tryLock(20000)) return;
  try { renderDisplay_(); } finally { l.releaseLock(); }
}

function renderDisplay_() {
  setDirty_(false);
  const st = readState_();
  const now = new Date();
  const ty = Number(Utilities.formatDate(now, CFG.TZ, 'yyyy')), tm = Number(Utilities.formatDate(now, CFG.TZ, 'M')), td = Number(Utilities.formatDate(now, CFG.TZ, 'd'));
  const ms = [];
  for (let i = 0; i < 3; i++) { let m = tm - 1 + i; const y = ty + Math.floor(m / 12); m = ((m % 12) + 12) % 12; ms.push({ y: y, m: m + 1, key: y + '-' + ('0' + (m + 1)).slice(-2), days: new Date(y, m + 1, 0).getDate() }); }
  const S = st.settings || {};
  const hol = S.holidays || [];
  const fy = (tm >= 10 ? ty + 1 : ty) + 543;
  let up = S.updated || '';
  let maxTs = 0; objValues_(st.events).forEach(function (e) { if (e.ts > maxTs) maxTs = e.ts; });
  let upD = up ? new Date(up + 'T00:00:00+07:00') : null; if (maxTs && (!upD || maxTs > upD.getTime())) upD = new Date(maxTs);
  const upStr = upD ? Utilities.formatDate(upD, CFG.TZ, 'dd_MM_') + String(Number(Utilities.formatDate(upD, CFG.TZ, 'yyyy')) + 543).slice(-2) : '';

  // ---- build a cell model
  const C = 32; const cells = {}; const merges = []; const heights = {}; const blocks = [];
  function cell(r, c) { const k = r + ':' + c; return cells[k] || (cells[k] = { v: '', bg: null, fc: '#000000', b: false, i: false, u: false, sz: 12, ha: 'center', wrap: true }); }
  function lanesOf(k) { let need = 0; objValues_(st.events).forEach(function (e) { if (e.month === k) need = Math.max(need, e.lane + (e.rows || 1)); }); return Math.max((st.months[k] && st.months[k].lanes) || 6, need, 1); }
  function isOff(M, d) { const g = new Date(M.y, M.m - 1, d).getDay(); return g === 0 || g === 6 || hol.indexOf(M.key + '-' + ('0' + d).slice(-2)) >= 0; }
  function textColor(e) { if (e.status === 'cancelled') return '#FF0000'; if (e.color) return e.color; return (TYPES[e.type] || TYPES.other).text || '#000000'; }
  function fillColor(e) { return e.fill || (TYPES[e.type] || TYPES.other).fill; }
  function label(e) { return e.text + (e.status === 'cancelled' ? ' (ยกเลิก)' : e.status === 'pending' ? ' (รอยืนยัน)' : ''); }

  let r = 1;
  Object.assign(cell(r, 1), { v: (S.unit || 'แผนภารกิจ') + ' ปีงบประมาณ ' + String(fy).slice(-2) + (upStr ? ' (Update ' + upStr + ')' : ''), b: true, sz: 16, wrap: false });
  merges.push([r, 1, 1, C]); heights[r] = 30; r++;
  Object.assign(cell(r, 1), { v: 'Master Plan ' + (fy - 543), b: true, sz: 16, wrap: false }); merges.push([r, 1, 1, C]); heights[r] = 28; r++;
  heights[r] = 6; r++;
  ms.forEach(function (M) {
    const hr = r;
    Object.assign(cell(r, 1), { v: 'เดือน', bg: '#D6DCE4', b: true });
    for (let d = 1; d <= 31; d++) Object.assign(cell(r, 1 + d), { v: thd_(d), bg: d > M.days ? '#000000' : '#D6DCE4', b: true });
    heights[r] = 24; r++;
    const n = lanesOf(M.key), r0 = r;
    let pastTo = 0; if (M.y < ty || (M.y === ty && M.m < tm)) pastTo = M.days; else if (M.y === ty && M.m === tm) pastTo = td - 1;
    Object.assign(cell(r0, 1), { v: MONTHS[M.m - 1] + '\n' + thd_(String(M.y + 543).slice(-2)), bg: '#DBDBDB', b: true, sz: 16 });
    merges.push([r0, 1, n, 1]);
    for (let L = 0; L < n; L++) {
      heights[r0 + L] = 26;
      for (let d = 1; d <= 31; d++) { const base = isOff(M, d) ? '#E3E3E3' : '#FFFFFF'; cell(r0 + L, 1 + d).bg = d > M.days ? '#000000' : (d <= pastTo ? pastBlend_(base) : base); }
    }
    const occ = {};
    objValues_(st.events).filter(function (e) { return e.month === M.key; }).sort(function (a, b) { return a.lane - b.lane || a.start - b.start; }).forEach(function (e) {
      const s = Math.max(1, e.start), en = Math.min(Math.max(e.end, s), M.days), rr = Math.max(1, Math.min(e.rows || 1, n - e.lane));
      if (e.lane >= n) return;
      let free = true; for (let i = 0; i < rr; i++) for (let d = s; d <= en; d++) if (occ[(e.lane + i) + ':' + d]) free = false;
      if (!free && occ[e.lane + ':' + s]) return;
      const lastD = free ? en : s, lastR = free ? rr : 1;
      const bg = en <= pastTo ? pastBlend_(fillColor(e)) : hex_(fillColor(e));
      for (let i = 0; i < lastR; i++) for (let d = s; d <= lastD; d++) { occ[(e.lane + i) + ':' + d] = 1; cell(r0 + e.lane + i, 1 + d).bg = bg; }
      Object.assign(cell(r0 + e.lane, 1 + s), { v: label(e), bg: bg, fc: hex_(textColor(e)), b: !!(e.bold || e.status === 'cancelled'), i: !!e.italic, u: !!e.underline, sz: e.size || 12 });
      if (lastR > 1 || lastD > s) merges.push([r0 + e.lane, 1 + s, lastR, lastD - s + 1]);
      if (lastR === 1) { const lines = label(e).split('\n').length; heights[r0 + e.lane] = Math.max(heights[r0 + e.lane], Math.ceil(lines * gsz_(e.size || 12) * 1.6 + 6)); }
    });
    blocks.push({ hr: hr, r0: r0, n: n, days: M.days });
    r = r0 + n; heights[r] = 8; r++;
  });
  Object.assign(cell(r, 1), { v: 'หมายเหตุ', b: true, ha: 'left', wrap: false });
  LEGEND.forEach(function (k, i) { const c = 2 + i * 5; cell(r, c).bg = TYPES[k].fill; Object.assign(cell(r, c + 1), { v: TYPES[k].label, ha: 'left', wrap: false }); merges.push([r, c + 1, 1, 4]); });
  const legendRow = r; heights[r] = 24; r++;
  heights[r] = 6; r++;
  const noteRows = [];
  (S.notes || []).forEach(function (nn) {
    if (!nn.text && !nn.label) return;
    const L = noteFmt_(nn, 'l'), T = noteFmt_(nn, 't');
    Object.assign(cell(r, 1), { v: nn.label || '', fc: hex_(L.color || '#000000'), bg: L.bg ? hex_(L.bg) : null, b: !!L.bold, i: !!L.italic, u: !!L.underline, sz: L.size, ha: 'left', wrap: false });
    Object.assign(cell(r, 2), { v: String(nn.text || '').replace(/\*\*/g, ''), fc: hex_(T.color || '#000000'), bg: T.bg ? hex_(T.bg) : null, b: !!T.bold, i: !!T.italic, u: !!T.underline, sz: T.size, ha: 'left', wrap: false });
    for (let c = 3; c <= C; c++) cell(r, c).bg = T.bg ? hex_(T.bg) : null;
    merges.push([r, 2, 1, C - 1]); noteRows.push({ r: r, text: nn.text || '', T: T });
    heights[r] = Math.max(22, Math.ceil(gsz_(Math.max(L.size, T.size)) * 1.9)); r++;
  });
  r++;
  Object.assign(cell(r, 1), { v: 'ไฟล์นี้สร้างอัตโนมัติจากไฟล์ต้นทาง · อัปเดต ' + Utilities.formatDate(now, CFG.TZ, 'dd/MM/') + String(ty + 543).slice(-2) + ' ' + Utilities.formatDate(now, CFG.TZ, 'HH:mm') + ' น. · แก้ไขผ่านหน้าเว็บเท่านั้น', fc: '#888888', sz: 10, ha: 'left', wrap: false, i: true });
  merges.push([r, 1, 1, C]);
  const R = r;

  // ---- write to the display spreadsheet
  const ss = SpreadsheetApp.openById(CFG.DISPLAY_ID);
  let sh = ss.getSheetByName(CFG.DISPLAY_SHEET);
  if (!sh) { const first = ss.getSheets()[0]; if (ss.getSheets().length === 1 && first.getLastRow() === 0) { sh = first.setName(CFG.DISPLAY_SHEET); } else sh = ss.insertSheet(CFG.DISPLAY_SHEET, 0); }
  if (sh.getMaxRows() < R) sh.insertRowsAfter(sh.getMaxRows(), R - sh.getMaxRows());
  if (sh.getMaxColumns() < C) sh.insertColumnsAfter(sh.getMaxColumns(), C - sh.getMaxColumns());
  const all = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  all.breakApart(); all.clear(); all.setBorder(false, false, false, false, false, false);
  sh.setHiddenGridlines(true);
  const vals = [], bgs = [], fcs = [], fws = [], fss = [], szs = [], uls = [], has = [], wraps = [];
  for (let i = 1; i <= R; i++) {
    const a = [], b = [], c2 = [], d = [], e = [], f = [], g = [], h = [], w = [];
    for (let j = 1; j <= C; j++) {
      const x = cells[i + ':' + j] || { v: '', bg: null, fc: '#000000', b: false, i: false, u: false, sz: 12, ha: 'center', wrap: true };
      a.push(x.v); b.push(x.bg); c2.push(x.fc); d.push(x.b ? 'bold' : 'normal'); e.push(x.i ? 'italic' : 'normal'); f.push(gsz_(x.sz)); g.push(x.u ? 'underline' : 'none'); h.push(x.ha); w.push(x.wrap);
    }
    vals.push(a); bgs.push(b); fcs.push(c2); fws.push(d); fss.push(e); szs.push(f); uls.push(g); has.push(h); wraps.push(w);
  }
  const rg = sh.getRange(1, 1, R, C);
  rg.setNumberFormat('@').setValues(vals).setBackgrounds(bgs).setFontColors(fcs).setFontWeights(fws).setFontStyles(fss)
    .setFontSizes(szs).setFontLines(uls).setHorizontalAlignments(has).setWraps(wraps).setVerticalAlignment('middle').setFontFamily('Sarabun');
  merges.forEach(function (m) { sh.getRange(m[0], m[1], m[2], m[3]).merge(); });
  noteRows.forEach(function (n) { // **bold** parts inside a note
    const runs = richRuns_(n.text); if (!runs.some(function (x) { return x.b; })) return;
    const plain = runs.map(function (x) { return x.t; }).join('');
    const b = SpreadsheetApp.newRichTextValue().setText(plain);
    let pos = 0;
    runs.forEach(function (x) { if (x.b) b.setTextStyle(pos, pos + x.t.length, SpreadsheetApp.newTextStyle().setBold(true).build()); pos += x.t.length; });
    sh.getRange(n.r, 2).setRichTextValue(b.build());
  });
  const SOL = SpreadsheetApp.BorderStyle.SOLID, MED = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
  blocks.forEach(function (bk) {
    sh.getRange(bk.r0, 2, bk.n, 31).setBorder(true, true, true, true, true, true, '#000000', SOL);
    sh.getRange(bk.hr, 1, 1, C).setBorder(true, true, true, true, true, true, '#000000', MED);
    sh.getRange(bk.r0, 1, bk.n, 1).setBorder(true, true, true, true, null, null, '#000000', MED);
    sh.getRange(bk.hr, 1, bk.n + 1, C).setBorder(true, true, true, true, null, null, '#000000', MED);
  });
  LEGEND.forEach(function (k, i) { sh.getRange(legendRow, 2 + i * 5).setBorder(true, true, true, true, null, null, '#000000', SOL); });
  sh.setColumnWidth(1, 120); sh.setColumnWidths(2, 31, 54);
  Object.keys(heights).forEach(function (k) { sh.setRowHeight(Number(k), heights[k]); });
  if (sh.getMaxRows() > R + 2) sh.deleteRows(R + 3, sh.getMaxRows() - R - 2);
  if (sh.getMaxColumns() > C) sh.deleteColumns(C + 1, sh.getMaxColumns() - C);
  // a gentle warning for anyone who tries to type here
  const pr = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if (!pr.length) sh.protect().setDescription('สร้างอัตโนมัติจากไฟล์ต้นทาง แก้ไขผ่านหน้าเว็บเท่านั้น').setWarningOnly(true);
  SpreadsheetApp.flush();
}

/* =========================== one-time setup =========================== */

/** รันครั้งเดียวหลังวางโค้ด: สร้างแท็บ ตั้งรหัสเริ่มต้น ตั้งตัวตั้งเวลา และวาดไฟล์แสดงผล */
function setup() {
  sheet_('events', EV_COLS); sheet_('months', ['month', 'lanes']); sheet_('settings', ['key', 'value']);
  sheet_('logs', ['at', 'time', 'by', 'action', 'detail']); snapSheet_();
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('CODE_HASH')) p.setProperty('CODE_HASH', sha256_(CFG.DEFAULT_CODE));
  if (!p.getProperty('VER')) p.setProperty('VER', '1');
  // first install: copy the current plan from Seed.gs (if that file was added) into the empty source sheet
  if (typeof SEED !== 'undefined' && !Object.keys(readState_().events).length) {
    const st = { events: {}, months: SEED.months || {}, settings: SEED.settings || {} };
    (SEED.events || []).forEach(function (e) { st.events[e.id] = Object.assign(cleanEv_(e), { id: e.id, by: e.by || '' }); });
    writeState_(st, { events: true, months: true, settings: true });
    appendLogs_([[Date.now(), 'ระบบ', 'นำเข้าข้อมูลเริ่มต้น', (SEED.events || []).length + ' ภารกิจ']]);
    Logger.log('นำเข้าข้อมูลเริ่มต้น ' + (SEED.events || []).length + ' ภารกิจ');
  }
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('renderIfDirty').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('renderDaily').timeBased().atHour(0).nearMinute(5).everyDays(1).inTimezone(CFG.TZ).create();
  ScriptApp.newTrigger('pruneLogs').timeBased().atHour(2).everyDays(1).inTimezone(CFG.TZ).create();
  DriveApp.getFolderById(CFG.BACKUP_FOLDER_ID).getName(); // ask for Drive permission now
  bump_();
  renderDisplay_();
  Logger.log('ตั้งค่าเสร็จแล้ว');
}

/** เปลี่ยนรหัสเปิดใช้งาน: แก้ข้อความในวงเล็บแล้วรันฟังก์ชันนี้ */
function setAccessCode() {
  const NEW_CODE = 'operation.slide';
  PropertiesService.getScriptProperties().setProperty('CODE_HASH', sha256_(NEW_CODE));
  Logger.log('เปลี่ยนรหัสแล้ว');
}

/** ปลดสิทธิ์แก้ไขที่ค้าง (กรณีฉุกเฉิน) */
function clearEditLock() { lockSet_(null); Logger.log('ปลดสิทธิ์แก้ไขแล้ว'); }
