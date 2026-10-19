// Google Apps Script backend: receives daily-log jobs and overtime entries from phones.
const SECRET = "GST12345";   // must match KEY in config.js
const COLS = ["id","submittedAt","building","logDate","supervisor","shift","wo","location","complaint","technician","action","category","priority","received","start","finish","materials","status","pending","remarks"];
const KEEP_BACKUPS = 14;      // how many daily backup copies to keep
const U_COLS = ["id","submittedAt","technician","date","type","area","meter","account","reading"];
const DEFAULT_METERS = [["Electricity","Common","21172699","2888542999"],["Electricity","101","21135023","4254023806"],["Electricity","205","21135492","20000282308"],
  ["Water","Common","10W708361","2888542999"],["Water","101","12A009001","4254023806"],["Water","205","12A009016","4254023806".replace("4254023806","20000282308")],["Water","Fire Water","110194605","7004047472"]];
const OT_COLS = ["id","submittedAt","technician","otDate","from","to","type","location","work"];

function tab_(name, cols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight("bold");
    sh.getRange(2, 1, sh.getMaxRows() - 1, cols.length).setNumberFormat("@"); // keep times/dates as text
    sh.setFrozenRows(1);
  }
  return sh;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function append_(sh, cols, items) {            // adds new items only (by id); returns ids now stored
  const last = sh.getLastRow();
  const have = new Set(last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])) : []);
  const rows = [], saved = [];
  (items || []).forEach(en => {
    if (!have.has(en.id)) { en.submittedAt = new Date().toISOString(); rows.push(cols.map(c => en[c] == null ? "" : String(en[c]))); have.add(en.id); }
    saved.push(en.id);
  });
  if (rows.length) sh.getRange(last + 1, 1, rows.length, cols.length).setNumberFormat("@").setValues(rows);
  return saved;
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: "bad request" }); }
  if (d.key !== SECRET) return out_({ ok: false, error: "wrong key" });
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const pins = readLists_().pins, rejected = [];
    const allow = en => { const ok = pinOk_(pins, en.technician, d.pin); if (!ok) rejected.push(en.id); return ok; };
    const saved = append_(tab_("Logs", COLS), COLS, (d.entries || []).filter(allow));
    const savedOt = append_(tab_("Overtime", OT_COLS), OT_COLS, (d.overtime || []).filter(allow));
    const savedU = append_(tab_("Readings", U_COLS), U_COLS, (d.readings || []).filter(allow));
    return out_({ ok: true, saved: saved, savedOt: savedOt, savedU: savedU, rejected: rejected });
  } finally { lock.releaseLock(); }
}

function listsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Lists");
  if (!sh) {
    sh = ss.insertSheet("Lists");
    sh.getRange(1, 1, 1, 2).setValues([["Technicians", "PIN"]]).setFontWeight("bold");
    sh.getRange("B:B").setNumberFormat("@");               // keep PINs as text (e.g. 0123)
    sh.setColumnWidth(1, 200);
  }
  return sh;
}
function metersSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Meters");
  if (!sh) {
    sh = ss.insertSheet("Meters");
    sh.getRange(1, 1, 1, 4).setValues([["Type", "Unit/Area", "Meter No", "Account No"]]).setFontWeight("bold");
    sh.getRange("A2:D300").setNumberFormat("@");
    sh.getRange(2, 1, DEFAULT_METERS.length, 4).setValues(DEFAULT_METERS);
    sh.setColumnWidths(1, 4, 140);
  }
  return sh;
}
function readMeters_() {
  return metersSheet_().getDataRange().getValues().slice(1).filter(r => String(r[2]).trim())
    .map(r => ({ type: String(r[0]).trim(), area: String(r[1]).trim(), meter: String(r[2]).trim(), account: String(r[3]).trim() }));
}
function lastReadings_() {                       // latest reading per meter, so phones can warn about typos
  const out = {};
  read_("Readings", U_COLS, {}).forEach(r => {
    const o = out[r.meter];
    if (!o || r.date > o.date || (r.date === o.date && String(r.submittedAt) > o.s)) out[r.meter] = { date: r.date, reading: r.reading, s: String(r.submittedAt) };
  });
  return out;
}
function normPin_(v) { let s = String(v == null ? "" : v).trim(); if (/^\d+$/.test(s) && s.length < 4) s = s.padStart(4, "0"); return s; }
function readLists_() {
  const v = listsSheet_().getDataRange().getValues();
  const head = v[0].map(x => String(x).trim().toLowerCase()), i = head.indexOf("technicians"), p = head.indexOf("pin");
  const rows = i < 0 ? [] : v.slice(1).filter(r => String(r[i] || "").trim());
  const pins = {}; if (p >= 0) rows.forEach(r => { const x = normPin_(r[p]); if (x) pins[String(r[i]).trim()] = x; });
  const ap = head.indexOf("admin pin"), admin = (ap >= 0 && v.length > 1) ? normPin_(v[1][ap]) : "";
  return { technicians: rows.map(r => String(r[i]).trim()), pins: pins, admin: admin };
}
function adminOk_(pin) { const a = readLists_().admin; return !a || a === normPin_(pin); }
function pinOk_(pins, tech, pin) { const x = pins[String(tech || "").trim()]; return !x || x === normPin_(pin); }

// ---- run ONCE from the Apps Script editor: writes the technician list, their PINs (up to 6 digits) and the admin PIN ----
function setupPins() {
  const sh = listsSheet_();
  sh.clear();
  const rows = [["Technicians", "PIN"], ["Mehedi Hassan", "3765"], ["Kaushar Ali", "711312"], ["Ajeesh Mohankumar", "712371"], ["Nagendra Vishwakarma", "712892"],
    ["Ram Jatan", "713052"], ["Mohammad Jahangeer", "713150"], ["Ali Hassan", "713502"], ["Imtiyaz Kamruddin", "713520"], ["Madhab Charan", "713542"]];
  sh.getRange("B:B").setNumberFormat("@"); sh.getRange("D:D").setNumberFormat("@");
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(1, 4, 2, 1).setValues([["Admin PIN"], ["713361"]]);
  sh.getRange("A1:D1").setFontWeight("bold"); sh.setColumnWidth(1, 200);
}

// ---- daily backup: run setupBackup ONCE from the Apps Script editor ----
function backupNow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), it = DriveApp.getFoldersByName("Technician Log Backups");
  const dir = it.hasNext() ? it.next() : DriveApp.createFolder("Technician Log Backups");
  const name = ss.getName() + " backup " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const same = dir.getFilesByName(name); while (same.hasNext()) same.next().setTrashed(true);
  DriveApp.getFileById(ss.getId()).makeCopy(name, dir);
  const files = [], all = dir.getFiles(); while (all.hasNext()) files.push(all.next());
  files.sort((a, b) => b.getDateCreated() - a.getDateCreated()).slice(KEEP_BACKUPS).forEach(f => f.setTrashed(true));
}
function setupBackup() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "backupNow").forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("backupNow").timeBased().everyDays(1).atHour(2).create();
  backupNow();
}

function read_(name, cols, p) {                // rows as objects; dates/times turned back into text if Sheets converted them
  const sh = tab_(name, cols), last = sh.getLastRow();
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const dateCols = ["logDate", "otDate", "date"], timeCols = ["received", "start", "finish", "from", "to"];
  const fix = (c, v) => !(v instanceof Date) ? v :
    Utilities.formatDate(v, tz, dateCols.indexOf(c) >= 0 ? "yyyy-MM-dd" : timeCols.indexOf(c) >= 0 ? "HH:mm" : "yyyy-MM-dd'T'HH:mm:ss");
  let rows = last > 1 ? sh.getRange(2, 1, last - 1, cols.length).getValues() : [];
  rows = rows.map(r => { const o = {}; cols.forEach((c, i) => o[c] = fix(c, r[i])); return o; });
  const dc = cols.indexOf("logDate") >= 0 ? "logDate" : cols.indexOf("otDate") >= 0 ? "otDate" : "date";
  if (p.from) rows = rows.filter(r => String(r[dc]) >= p.from);
  if (p.to) rows = rows.filter(r => String(r[dc]) <= p.to);
  return rows;
}

function doGet(e) {
  const p = e.parameter || {};
  if (p.key !== SECRET) return out_({ ok: false, error: "wrong key" });
  if (p.action === "lists") { const l = readLists_(); return out_({ ok: true, lists: { technicians: l.technicians, pinRequired: Object.keys(l.pins).length > 0, meters: readMeters_(), last: lastReadings_() } }); }
  if (p.action === "adminlogin") return out_({ ok: true, valid: adminOk_(p.apin) });
  if (["lists", "login"].indexOf(p.action) < 0 && !adminOk_(p.apin)) return out_({ ok: false, error: "admin PIN required" });
  if (p.action === "login") return out_({ ok: true, valid: pinOk_(readLists_().pins, p.name, p.pin) });
  if (p.action === "readings") return out_({ ok: true, rows: read_("Readings", U_COLS, p) });
  if (p.action === "overtime") return out_({ ok: true, rows: read_("Overtime", OT_COLS, p) });
  return out_({ ok: true, rows: read_("Logs", COLS, p) });
}
