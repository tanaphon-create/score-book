/**
 * สมุดคะแนนรายห้อง — ฐานข้อมูล Google Sheet
 * วิธีติดตั้ง: ดู README.md ในโฟลเดอร์ score-book
 *
 * 1) เปลี่ยน PIN ด้านล่างเป็นรหัสของคุณเอง (ใช้ตอนเชื่อมต่อจากหน้าเว็บ)
 * 2) Deploy > New deployment > Web app
 *    Execute as: Me   ·   Who has access: Anyone
 */
const PIN = 'เปลี่ยนรหัสนี้';

const DATA_SHEET = '_data';      // ชีตเก็บข้อมูลดิบ (ซ่อนไว้) — อย่าแก้ไขด้วยมือ
const CHUNK = 45000;             // 1 เซลล์เก็บได้ไม่เกิน 50,000 ตัวอักษร
const ATT = {P: 'มา', A: 'ขาด', L: 'ลา', S: 'สาย'};

function doGet() {
  return out_({ok: true, app: 'scorebook', hint: 'ใช้ POST จากหน้าเว็บสมุดคะแนน'});
}

function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (_) {}
  try {
    if (String(p.pin) !== String(PIN)) return out_({ok: false, error: 'pin'});
    if (p.action === 'ping') return out_({ok: true});
    if (p.action === 'load') return out_(Object.assign({ok: true}, load_()));
    if (p.action === 'save') {
      const lock = LockService.getScriptLock();
      lock.waitLock(25000);
      try { save_(p); } finally { lock.releaseLock(); }
      return out_({ok: true, at: Date.now()});
    }
    return out_({ok: false, error: 'action'});
  } catch (err) {
    return out_({ok: false, error: String(err && err.message || err)});
  }
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- raw storage ---------------- */
// คอลัมน์: A id · B อัปเดตล่าสุด · C ชื่อชีตที่สร้าง (JSON) · D… ข้อมูล JSON แบ่งเป็นช่วง
function dataSheet_() {
  const ss = SpreadsheetApp.getActive();
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
  const ss = SpreadsheetApp.getActive();
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
    ['คะแนนเสริม', 'รวม (' + max + ')', 'ร้อยละ', 'เกรด', 'มาเรียน %', 'ขาด', 'ลา', 'สาย']);
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
    const att = {P: 0, A: 0, L: 0, S: 0, n: 0};
    dates.forEach(d => { const x = c.attendance[d][s.id]; if (x) { att[x]++; att.n++; } });
    const attPct = att.n ? (att.P + att.S) / att.n * 100 : '';
    const ms = attPct !== '' && attPct < 80;
    return [s.no, s.code, s.name].concat(cells,
      [bonus || '', Math.round(total * 100) / 100, Math.round(pct * 100) / 100, ms ? 'มส' : (max ? grade_(pct) : ''),
       attPct === '' ? '' : Math.round(attPct * 10) / 10, att.A, att.L, att.S]);
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
  const ah = ['เลขที่', 'ชื่อ-สกุล'].concat(dates, ['มา', 'ขาด', 'ลา', 'สาย', 'มาเรียน %']);
  const ab = students.map(s => {
    const n = {P: 0, A: 0, L: 0, S: 0};
    const cells = dates.map(d => { const x = c.attendance[d][s.id]; if (x) n[x]++; return x ? ATT[x] : ''; });
    const all = n.P + n.A + n.L + n.S;
    return [s.no, s.name].concat(cells, [n.P, n.A, n.L, n.S, all ? Math.round((n.P + n.S) / all * 1000) / 10 : '']);
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
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ลา').setBackground('#e4f3fb').setFontColor('#0284c7').setRanges([rng]).build(),
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
