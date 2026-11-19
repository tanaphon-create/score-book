/**
 * สมุดคะแนนรายห้อง — ฐานข้อมูล Google Sheet
 * วิธีติดตั้ง: ดู README.md ในโฟลเดอร์ score-book
 *
 * 1) ใส่อีเมลครูที่มีสิทธิ์ใน TEACHER_EMAILS (หน้าเว็บล็อกอิน Google แล้วส่ง token มาให้ตรวจ)
 * 2) Deploy > New deployment > Web app
 *    Execute as: Me   ·   Who has access: Anyone
 */
const TEACHER_EMAILS = ['tanaphon@wangchan.ac.th'];
const FIREBASE_API_KEY = 'AIzaSyAtzOh4utxj9-ry2PNWH7YrqMb_UjuaNOs'; // โปรเจกต์ quizlive-wcw (ใช้ตรวจ token เท่านั้น)

// ตรวจ Firebase ID token กับ Google: ต้องเป็นอีเมลครูที่ยืนยันแล้ว (จำผลไว้ 30 นาที ไม่ต้องถามซ้ำทุกครั้ง)
function authOk_(p) {
  const tok = String(p.idToken || '');
  if (tok.length < 100) return false;
  const cache = CacheService.getScriptCache();
  const key = 'tk_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, tok));
  const hit = cache.get(key);
  if (hit) return hit === '1';
  const r = UrlFetchApp.fetch('https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + FIREBASE_API_KEY, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify({idToken: tok}), muteHttpExceptions: true
  });
  let ok = false;
  if (r.getResponseCode() === 200) {
    const u = (JSON.parse(r.getContentText()).users || [])[0];
    ok = !!u && u.emailVerified === true && TEACHER_EMAILS.indexOf(String(u.email || '').toLowerCase()) >= 0;
  }
  cache.put(key, ok ? '1' : '0', ok ? 1800 : 60);
  return ok;
}

// ถ้าสคริปต์ไม่ได้สร้างจากเมนู ส่วนขยาย > Apps Script ของชีตโดยตรง ให้ใส่รหัสไฟล์ชีตที่นี่ (ส่วนที่อยู่ระหว่าง /d/ กับ /edit ในลิงก์)
const SPREADSHEET_ID = '';
function ss_() { return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }

const DATA_SHEET = '_data';      // ชีตเก็บข้อมูลดิบ (ซ่อนไว้) — อย่าแก้ไขด้วยมือ
const CHUNK = 45000;             // 1 เซลล์เก็บได้ไม่เกิน 50,000 ตัวอักษร
const ATT = {P: 'มา', S: 'สาย', A: 'ขาด', L: 'ลากิจ', H: 'ลาป่วย', E: 'กิจกรรม'};

function doGet() {
  return out_({ok: true, app: 'scorebook', hint: 'ใช้ POST จากหน้าเว็บสมุดคะแนน'});
}

function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (_) {}
  try {
    // หน้านักเรียน: ค้นด้วยเลขประจำตัว ไม่ต้องล็อกอิน (ส่งเฉพาะข้อมูลของคนนั้น และเฉพาะเมื่อครูเปิดไว้)
    if (p.action === 'student') return out_(studentLookup_(p.code));
    if (!authOk_(p)) return out_({ok: false, error: 'auth'});
    if (p.action === 'ping') return out_({ok: true, publish: isPublished_()});
    if (p.action === 'setpublish') {
      PropertiesService.getScriptProperties().setProperty('PUBLISH', p.on ? '1' : '0');
      return out_({ok: true, publish: isPublished_()});
    }
    if (p.action === 'load') return out_(Object.assign({ok: true}, load_()));
    if (p.action === 'save') {
      const lock = LockService.getScriptLock();
      lock.waitLock(25000);
      try { save_(p); } finally { lock.releaseLock(); }
      return out_({ok: true, at: Date.now()});
    }
    if (p.action === 'classes') return out_(classesBrief_());
    if (p.action === 'addbonus') {
      const lock = LockService.getScriptLock();
      lock.waitLock(25000);
      try { return out_(addBonus_(p)); } finally { lock.releaseLock(); }
    }
    return out_({ok: false, error: 'action'});
  } catch (err) {
    return out_({ok: false, error: String(err && err.message || err)});
  }
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- หน้านักเรียนดูคะแนน (สาธารณะ) ---------------- */
function isPublished_() { return PropertiesService.getScriptProperties().getProperty('PUBLISH') === '1'; }

function studentLookup_(code) {
  if (!isPublished_()) return {ok: false, error: 'closed'};
  code = String(code == null ? '' : code).trim();
  if (!code || code.length > 30) return {ok: false, error: 'code'};
  // กันการไล่เดาเลข: จำกัดจำนวนครั้งต่อนาทีของทั้งระบบ
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get('lk') || 0);
  if (n >= 90) return {ok: false, error: 'busy'};
  cache.put('lk', String(n + 1), 60);

  const found = [];
  load_().classes.forEach(c => {
    const st = (c.students || []).find(x => String(x.code).trim() === code);
    if (!st) return;
    const sc = {};
    (c.tasks || []).forEach(t => { const v = c.scores && c.scores[t.id] && c.scores[t.id][st.id]; sc[t.id] = (v === undefined || v === '') ? null : v; });
    const att = [];
    Object.keys(c.attendance || {}).sort().forEach(d => { const v = c.attendance[d][st.id]; if (v) att.push({d: d, v: v}); });
    found.push({
      cls: c.name, subject: c.subject || '',
      student: {no: st.no, code: st.code, name: st.name},
      tasks: (c.tasks || []).map(t => ({name: t.name, max: t.max, cat: t.cat, score: sc[t.id]})),
      bonus: (c.bonus || []).filter(b => b.sid === st.id).map(b => ({pts: b.pts, reason: b.reason, date: b.date})),
      attendance: att
    });
  });
  return found.length ? {ok: true, results: found} : {ok: false, error: 'notfound'};
}

/* ---------------- raw storage ---------------- */
// คอลัมน์: A id · B อัปเดตล่าสุด · C ชื่อชีตที่สร้าง (JSON) · D… ข้อมูล JSON แบ่งเป็นช่วง
function dataSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(DATA_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DATA_SHEET);
    sh.getRange(1, 1, 1, 4).setValues([['id', 'updated', 'sheets', 'json…']]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.hideSheet();
  }
  return sh;
}

function readRows_() {
  const sh = dataSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const w = Math.max(4, sh.getLastColumn());
  return sh.getRange(2, 1, last - 1, w).getValues().map((r, i) => ({
    row: i + 2,
    id: String(r[0]),
    sheets: r[2] ? JSON.parse(r[2]) : [],
    json: r.slice(3).join('')
  })).filter(r => r.id);
}

function load_() {
  const classes = [];
  let meta = null;
  readRows_().forEach(r => {
    if (!r.json) return;
    const v = JSON.parse(r.json);
    if (r.id === '__meta') meta = v; else classes.push(v);
  });
  return {classes: classes, meta: meta};
}

function writeRow_(rows, id, json, sheets) {
  const sh = dataSheet_();
  const parts = [];
  for (let i = 0; i < json.length; i += CHUNK) parts.push(json.slice(i, i + CHUNK));
  const vals = [id, new Date(), JSON.stringify(sheets || [])].concat(parts);
  let r = rows.find(x => x.id === id);
  let rowNo = r ? r.row : sh.getLastRow() + 1;
  if (r) {
    const w = sh.getLastColumn();
    if (w > 0) sh.getRange(rowNo, 1, 1, w).clearContent();
  }
  sh.getRange(rowNo, 1, 1, vals.length).setNumberFormat('@').setValues([vals]);
  sh.getRange(rowNo, 2).setNumberFormat('yyyy-mm-dd hh:mm');
  if (!r) rows.push({row: rowNo, id: id, sheets: sheets || []});
}

function save_(p) {
  const ss = ss_();
  const rows = readRows_();

  (p.classes || []).forEach(c => {
    if (!c || !c.id) return;
    const old = rows.find(x => x.id === c.id);
    const names = writeReadable_(ss, c);
    if (old) old.sheets.filter(n => names.indexOf(n) < 0).forEach(n => {
      const s = ss.getSheetByName(n); if (s) ss.deleteSheet(s);
    });
    writeRow_(rows, c.id, JSON.stringify(c), names);
  });

  if (p.meta) writeRow_(rows, '__meta', JSON.stringify(p.meta), []);

  const sh = dataSheet_();
  (p.deleted || []).forEach(id => {
    const fresh = readRows_();
    const r = fresh.find(x => x.id === id);
    if (!r) return;
    r.sheets.forEach(n => { const s = ss.getSheetByName(n); if (s) ss.deleteSheet(s); });
    sh.deleteRow(r.row);
  });
}

/* ---------------- readable tabs ---------------- */
function safeName_(s) { return String(s).replace(/[\[\]\*\?\/\\:]/g, '-').slice(0, 80); }

function grade_(p) {
  if (p >= 80) return 4; if (p >= 75) return 3.5; if (p >= 70) return 3; if (p >= 65) return 2.5;
  if (p >= 60) return 2; if (p >= 55) return 1.5; if (p >= 50) return 1; return 0;
}

function sheetFor_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clear();
  return sh;
}

function writeReadable_(ss, c) {
  const base = safeName_(c.name || 'ห้อง');
  const scoreName = base + ' คะแนน';
  const attName = base + ' เช็คชื่อ';
  const students = (c.students || []).slice().sort((a, b) => ((+a.no || 9999) - (+b.no || 9999)));
  const tasks = c.tasks || [];
  const dates = Object.keys(c.attendance || {}).sort();
  const max = tasks.reduce((a, t) => a + (+t.max || 0), 0);

  // --- คะแนน ---
  const head = ['เลขที่', 'รหัส', 'ชื่อ-สกุล'].concat(tasks.map(t => t.name + ' (' + t.max + ')'),
    ['คะแนนเสริม', 'รวม (' + max + ')', 'ร้อยละ', 'เกรด', 'มาเรียน %', 'ขาด', 'ลากิจ', 'ลาป่วย', 'สาย', 'กิจกรรม']);
  const body = students.map(s => {
    let raw = 0;
    const cells = tasks.map(t => {
      const v = c.scores && c.scores[t.id] ? c.scores[t.id][s.id] : null;
      if (v != null && v !== '') { raw += +v; return +v; }
      return '';
    });
    const bonus = (c.bonus || []).filter(b => b.sid === s.id).reduce((a, b) => a + (+b.pts || 0), 0);
    const total = raw + bonus;
    const pct = max ? Math.min(100, total / max * 100) : 0;
    const att = {P: 0, A: 0, L: 0, S: 0, H: 0, E: 0, n: 0};
    dates.forEach(d => { const x = c.attendance[d][s.id]; if (x && att[x] != null) { att[x]++; att.n++; } });
    const attPct = att.n ? (att.P + att.S + att.E) / att.n * 100 : '';
    const ms = attPct !== '' && attPct < 80;
    return [s.no, s.code, s.name].concat(cells,
      [bonus || '', Math.round(total * 100) / 100, Math.round(pct * 100) / 100, ms ? 'มส' : (max ? grade_(pct) : ''),
       attPct === '' ? '' : Math.round(attPct * 10) / 10, att.A, att.L, att.H, att.S, att.E]);
  });
  const sh = sheetFor_(ss, scoreName);
  sh.getRange(1, 1).setValue((c.name || '') + (c.subject ? ' · ' + c.subject : '') + '   (อัปเดตจากเว็บ ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd/M/yyyy HH:mm') + ' — แก้ไขคะแนนผ่านหน้าเว็บเท่านั้น)')
    .setFontWeight('bold');
  sh.getRange(2, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#eef0ff').setWrap(true);
  if (body.length) sh.getRange(3, 1, body.length, head.length).setValues(body);
  sh.setFrozenRows(2); sh.setFrozenColumns(3);
  sh.setColumnWidth(3, 220);

  // --- เช็คชื่อ ---
  const ah = ['เลขที่', 'ชื่อ-สกุล'].concat(dates, ['มา', 'สาย', 'กิจกรรม', 'ขาด', 'ลากิจ', 'ลาป่วย', 'มาเรียน %']);
  const ab = students.map(s => {
    const n = {P: 0, A: 0, L: 0, S: 0, H: 0, E: 0};
    const cells = dates.map(d => { const x = c.attendance[d][s.id]; if (x && n[x] != null) n[x]++; return x ? (ATT[x] || x) : ''; });
    const all = n.P + n.A + n.L + n.S + n.H + n.E;
    return [s.no, s.name].concat(cells, [n.P, n.S, n.E, n.A, n.L, n.H, all ? Math.round((n.P + n.S + n.E) / all * 1000) / 10 : '']);
  });
  const sa = sheetFor_(ss, attName);
  sa.getRange(1, 1, 1, ah.length).setNumberFormat('@').setValues([ah]).setFontWeight('bold').setBackground('#eef0ff');
  if (ab.length) sa.getRange(2, 1, ab.length, ah.length).setValues(ab);
  sa.setFrozenRows(1); sa.setFrozenColumns(2);
  sa.setColumnWidth(2, 220);
  if (ab.length && dates.length) {
    const rng = sa.getRange(2, 3, ab.length, dates.length);
    sa.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ขาด').setBackground('#fdecec').setFontColor('#dc2626').setRanges([rng]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ลากิจ').setBackground('#e4f3fb').setFontColor('#0284c7').setRanges([rng]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ลาป่วย').setBackground('#fae8ff').setFontColor('#a21caf').setRanges([rng]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('กิจกรรม').setBackground('#dcf5f1').setFontColor('#0f766e').setRanges([rng]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('สาย').setBackground('#fdf3e2').setFontColor('#d97706').setRanges([rng]).build()
    ]);
  }
  band_(sh, 2, body.length, head.length);
  band_(sa, 1, ab.length, ah.length);

  // --- คะแนนเสริม ---
  const bonusName = base + ' คะแนนเสริม';
  const byId = {}; students.forEach(s => byId[s.id] = s);
  const log = (c.bonus || []).filter(b => byId[b.sid]).sort((a, b) => a.ts - b.ts);
  const bh = ['วันที่', 'เลขที่', 'ชื่อ-สกุล', 'คะแนน', 'เหตุผล'];
  const bb = log.map(b => [b.date, byId[b.sid].no, byId[b.sid].name, +b.pts, b.reason || '']);
  const sb = sheetFor_(ss, bonusName);
  sb.getRange(1, 1, 1, bh.length).setValues([bh]).setFontWeight('bold').setBackground('#eef0ff');
  if (bb.length) {
    sb.getRange(2, 1, bb.length, 1).setNumberFormat('@');
    sb.getRange(2, 1, bb.length, bh.length).setValues(bb);
    sb.getRange(bb.length + 2, 3, 1, 2).setValues([['รวมทั้งหมด', '=SUM(D2:D' + (bb.length + 1) + ')']]).setFontWeight('bold');
  }
  sb.setFrozenRows(1); sb.setColumnWidth(3, 220); sb.setColumnWidth(5, 260);
  band_(sb, 1, bb.length, bh.length);

  return [scoreName, attName, bonusName];
}

// สลับสีแถวให้อ่านง่าย + จัดกึ่งกลาง
function band_(sh, headRow, n, w) {
  sh.getBandings().forEach(b => b.remove());
  if (n > 0) {
    sh.getRange(headRow, 1, n + 1, w).applyRowBanding(SpreadsheetApp.BandingTheme.INDIGO, true, false);
    sh.getRange(headRow, 1, n + 1, w).setHorizontalAlignment('center').setVerticalAlignment('middle');
  }
  sh.getRange(headRow, 1, 1, w).setFontColor('#ffffff').setBackground('#4f46e5').setFontWeight('bold');
}


// กดเรียกใช้ครั้งเดียวเพื่อให้สิทธิ์ "เชื่อมต่อบริการภายนอก" (ใช้ตรวจการล็อกอินครูกับ Google)
function setupLoginCheck() {
  const r = UrlFetchApp.fetch('https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + FIREBASE_API_KEY, {method: 'post', contentType: 'application/json', payload: '{"idToken":"x"}', muteHttpExceptions: true});
  Logger.log('ตรวจการล็อกอินพร้อมใช้งาน (HTTP ' + r.getResponseCode() + ')');
}


/* ---------------- รับคะแนนพิเศษจากเกมคำถามฟลุ๊กถูก ---------------- */
// รายชื่อห้อง + เลขประจำตัวนักเรียน (ใช้จับคู่ในหน้าเกม ไม่ส่งคะแนนออกไป)
function classesBrief_() {
  return {ok: true, classes: load_().classes.map(c => ({
    id: c.id, name: c.name, subject: c.subject || '',
    students: (c.students || []).map(s => ({code: String(s.code || '').trim(), name: s.name})),
    sets: setsBrief_(c)
  }))};
}
// ชุดกลุ่มของห้อง (ใช้ในโหมดห้องเรียน) — สมาชิกส่งเป็นเลขประจำตัว
function setsBrief_(c) {
  const sets = (c.gsets && c.gsets.length) ? c.gsets : ((c.groups && c.groups.length) ? [{id: 'legacy', name: 'ชุดกลุ่ม 1', groups: c.groups}] : []);
  const codeOf = {};
  (c.students || []).forEach(s => { codeOf[s.id] = String(s.code || '').trim(); });
  return sets.map(gs => ({id: gs.id, name: gs.name, groups: (gs.groups || []).map(g => ({name: g.name, members: (g.members || []).map(id => codeOf[id]).filter(Boolean)}))}));
}

// เพิ่มคะแนนเสริมลงห้องที่เลือก จับคู่ด้วยเลขประจำตัว — กดซ้ำเกมเดิมจะไม่บวกซ้ำ (เช็ก gk)
function addBonus_(p) {
  const all = load_();
  const c = all.classes.find(x => x.id === p.classId);
  if (!c) return {ok: false, error: 'class'};
  c.bonus = c.bonus || [];
  const date = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  const gk = String(p.gameKey || '').slice(0, 60);
  const added = [], missing = [], dup = [];
  let ts = Date.now();
  (p.entries || []).slice(0, 60).forEach(e => {
    const code = String(e.code || '').trim();
    const pts = Math.max(-100, Math.min(100, Number(e.pts) || 0));
    const st = (c.students || []).find(s => String(s.code || '').trim() === code);
    if (!code || !st || !pts) { missing.push(code); return; }
    if (gk && c.bonus.some(b => b.gk === gk && b.sid === st.id)) { dup.push(code); return; }
    c.bonus.push({id: 'qz' + Math.random().toString(36).slice(2, 9), sid: st.id, pts: pts,
      reason: String(e.reason || p.reason || 'ตอบคำถาม').slice(0, 120), date: date, ts: ts++, gk: gk});
    added.push({code: code, name: st.name, pts: pts});
  });
  if (added.length) save_({classes: [c]});
  return {ok: true, cls: c.name, added: added, missing: missing, dup: dup};
}
