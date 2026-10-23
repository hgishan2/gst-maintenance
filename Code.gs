// Google Apps Script backend: receives daily-log jobs, overtime entries, and utility readings from phones.
const SECRET = "GST12345";   // must match KEY in config.js
const SCRIPT_VERSION = "2026-10-25";   // open YOUR_SCRIPT_URL?key=...&action=version to check which version is deployed
let REASONS = {};                      // why an item was refused: id -> "pin" | "name" | "notfound" | "invalid"
const COLS = ["id","submittedAt","building","logDate","supervisor","shift","wo","location","complaint","technician","action","category","priority","received","start","finish","materials","status","pending","remarks"];
const KEEP_BACKUPS = 14;      // how many daily backup copies to keep
const U_COLS = ["id","submittedAt","technician","date","type","area","meter","account","reading"];
const DEFAULT_METERS = [["Electricity","Common","21172699","2888542999"],["Electricity","101","21135023","4254023806"],["Electricity","205","21135492","20000282308"],
  ["Water","Common","10W708361","2888542999"],["Water","101","12A009001","4254023806"],["Water","205","12A009016","20000282308"],["Water","Fire Water","110194605","7004047472"]];
const OT_COLS = ["id","submittedAt","technician","otDate","from","to","type","location","work","otCategoryType"];
const S_COLS = ["id","createdAt","technician","date","location","category","desc","status","updatedAt"];

function tab_(name, cols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight("bold");
    sh.getRange(2, 1, sh.getMaxRows() - 1, cols.length).setNumberFormat("@"); // keep times/dates as text
    sh.setFrozenRows(1);
  } else {
    if (sh.getMaxColumns() < cols.length) sh.insertColumnsAfter(sh.getMaxColumns(), cols.length - sh.getMaxColumns());
    const have = sh.getRange(1, 1, 1, cols.length).getValues()[0];
    have.forEach((v, i) => {                      // fill in headers for columns added in newer versions
      if (String(v).trim() === "") {
        sh.getRange(1, i + 1).setValue(cols[i]).setFontWeight("bold");
        sh.getRange(2, i + 1, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat("@");
      }
    });
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
    REASONS = {};
    const pins = readLists_().pins, rejected = [];
    const allow = en => { const chk = pinCheck_(pins, en.technician, d.pin); if (chk !== "ok") { rejected.push(en.id); REASONS[en.id] = chk; } return chk === "ok"; };
    const saved = append_(tab_("Logs", COLS), COLS, (d.entries || []).filter(allow));
    const savedOt = append_(tab_("Overtime", OT_COLS), OT_COLS, (d.overtime || []).filter(allow));
    const savedU = append_(tab_("Readings", U_COLS), U_COLS, (d.readings || []).filter(allow));
    const sres = applySchedOps_(d.schedOps, d, pins);
    const tres = applyTechOps_(d.techOps, d);
    const lres = applyLogOps_(d.logOps, d, pins);
    return out_({ ok: true, saved: saved, savedOt: savedOt, savedU: savedU, rejected: rejected, schedDone: sres.done, schedRejected: sres.rejected, techDone: tres.done, techRejected: tres.rejected, logDone: lres.done, logRejected: lres.rejected, reasons: REASONS });
  } finally { lock.releaseLock(); }
}

// ---- Schedules: admin can assign to anyone (apin); a technician can only change their own (name + PIN) ----
function applySchedOps_(ops, d, pins) {
  const done = [], rejected = [];
  if (!ops || !ops.length) return { done: done, rejected: rejected };
  const sh = tab_("Schedules", S_COLS);
  const adminMode = !!d.apin && adminOk_(d.apin);
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  ops.forEach(op => {
    const it = op.item || {};
    const id = String(op.id || it.id || "");
    if (!id) return;
    const last = sh.getLastRow();
    const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])) : [];
    const pos = ids.indexOf(id), rowNo = pos >= 0 ? pos + 2 : -1;
    const existing = rowNo > 0 ? sh.getRange(rowNo, 1, 1, S_COLS.length).getValues()[0] : null;
    const owner = existing ? String(existing[2]) : String(it.technician || "");
    let allowed = adminMode || pinOk_(pins, owner, d.pin);
    if (allowed && !adminMode && op.op !== "delete" && normName_(it.technician) !== normName_(owner)) allowed = false;   // a technician cannot reassign
    if (!allowed) {
      const chk = adminMode ? "ok" : pinCheck_(pins, owner, d.pin);
      REASONS[id] = chk !== "ok" ? chk : "invalid";
      rejected.push(id); return;
    }

    if (op.op === "delete") {
      if (rowNo > 0) sh.deleteRow(rowNo);
      done.push(id);
      return;
    }
    const now = new Date().toISOString();
    const rec = {
      id: id, createdAt: existing ? existing[1] : now, technician: owner,
      date: String(it.date || ""), location: String(it.location || ""), category: String(it.category || ""),
      desc: String(it.desc || ""), status: String(it.status || "Pending"), updatedAt: now
    };
    const row = S_COLS.map(c => rec[c] == null ? "" : String(rec[c]));
    if (rowNo > 0) sh.getRange(rowNo, 1, 1, S_COLS.length).setNumberFormat("@").setValues([row]);
    else sh.getRange(last + 1, 1, 1, S_COLS.length).setNumberFormat("@").setValues([row]);
    done.push(id);
  });
  return { done: done, rejected: rejected };
}

// ---- Technician management (admin only). Works on the Lists sheet but NEVER touches the Admin PIN column. ----
function colFor_(sh, header) {                   // finds a header column in row 1, creating it at the end if missing
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(x => String(x).trim().toLowerCase());
  const i = head.indexOf(header.toLowerCase());
  if (i >= 0) return i + 1;
  const col = lastCol + 1;
  sh.getRange(1, col).setValue(header).setFontWeight("bold");
  sh.getRange(2, col, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat("@");
  return col;
}
function applyTechOps_(ops, d) {
  const done = [], rejected = [];
  if (!ops || !ops.length) return { done: done, rejected: rejected };
  if (!(d.apin && adminOk_(d.apin))) { ops.forEach(o => rejected.push(normName_(o && o.name))); return { done: done, rejected: rejected }; }
  const sh = listsSheet_();
  const cName = colFor_(sh, "Technicians"), cPin = colFor_(sh, "PIN"), cEmp = colFor_(sh, "Emp Code"), cDes = colFor_(sh, "Designation");
  ops.forEach(op => {
    const name = normName_(op.name);
    if (!name) return;
    const last = Math.max(sh.getLastRow(), 2);
    const names = sh.getRange(2, cName, last - 1, 1).getValues().map(r => normName_(r[0]));
    const pos = names.indexOf(name), rowNo = pos >= 0 ? pos + 2 : -1;

    if (op.op === "delete") {
      if (rowNo > 0) [cName, cPin, cEmp, cDes].forEach(c => sh.getRange(rowNo, c).clearContent());   // clear cells only: row 2 also holds the Admin PIN
      done.push(name);
      return;
    }
    const pin = String(op.pin == null ? "" : op.pin).trim();
    if (pin && (pin.length < 4 || pin.length > 10 || /\s/.test(pin))) { rejected.push(name); return; }
    let target = rowNo;
    if (target < 0) {
      if (!pin) { rejected.push(name); return; }                      // a new technician needs a PIN
      const blank = names.indexOf("");                                 // reuse an emptied row, else add at the end
      let lastName = 0; names.forEach((n, k) => { if (n) lastName = k + 1; });
      target = blank >= 0 && blank + 1 <= lastName ? blank + 2 : lastName + 2;
      sh.getRange(target, cName).setNumberFormat("@").setValue(name);
    }
    if (pin) sh.getRange(target, cPin).setNumberFormat("@").setValue(pin);
    if (op.empCode !== undefined) sh.getRange(target, cEmp).setNumberFormat("@").setValue(String(op.empCode || "").trim());
    if (op.designation !== undefined) sh.getRange(target, cDes).setNumberFormat("@").setValue(String(op.designation || "").trim());
    done.push(name);
  });
  return { done: done, rejected: rejected };
}

// ---- Update an existing job log (e.g. Pending / On Hold -> Completed). Admin: any job. Technician: only their own. ----
const LOG_EDITABLE = ["status", "action", "finish", "pending", "materials"];
function applyLogOps_(ops, d, pins) {
  const done = [], rejected = [];
  if (!ops || !ops.length) return { done: done, rejected: rejected };
  const sh = tab_("Logs", COLS);
  const adminMode = !!d.apin && adminOk_(d.apin);
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const iTech = COLS.indexOf("technician"), iStatus = COLS.indexOf("status"), iRem = COLS.indexOf("remarks");
  ops.forEach(op => {
    const id = String(op.id || "");
    if (!id) return;
    const last = sh.getLastRow();
    const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])) : [];
    const pos = ids.indexOf(id);
    if (pos < 0) { REASONS[id] = "notfound"; rejected.push(id); return; }                       // not on the server (yet)
    const rowNo = pos + 2;
    const row = sh.getRange(rowNo, 1, 1, COLS.length).getValues()[0];
    const owner = String(row[iTech]);
    if (!adminMode) { const chk = pinCheck_(pins, owner, d.pin); if (chk !== "ok") { REASONS[id] = chk; rejected.push(id); return; } }
    const u = op.updates || {};
    if (u.status !== undefined && ["Completed", "Pending", "On Hold"].indexOf(String(u.status)) < 0) { REASONS[id] = "invalid"; rejected.push(id); return; }

    const oldStatus = String(row[iStatus] || "Pending");
    const newStatus = u.status !== undefined ? String(u.status) : oldStatus;
    LOG_EDITABLE.forEach(c => {
      if (u[c] !== undefined) sh.getRange(rowNo, COLS.indexOf(c) + 1).setNumberFormat("@").setValue(String(u[c]));
    });
    if (newStatus === "Completed" && u.pending === undefined) sh.getRange(rowNo, COLS.indexOf("pending") + 1).setValue("");
    if (newStatus !== oldStatus) {                                     // audit trail in the Remarks column
      const note = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm") + " " + (adminMode ? "Admin" : owner) + ": " + oldStatus + " -> " + newStatus;
      const prev = String(row[iRem] || "");
      sh.getRange(rowNo, iRem + 1).setNumberFormat("@").setValue(prev ? prev + " | " + note : note);
    }
    done.push(id);
  });
  return { done: done, rejected: rejected };
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
function normName_(v) { return String(v == null ? "" : v).replace(/\s+/g, " ").trim().toUpperCase(); }
function readLists_() {
  const v = listsSheet_().getDataRange().getValues();
  const head = v[0].map(x => String(x).trim().toLowerCase()), i = head.indexOf("technicians"), p = head.indexOf("pin");
  const rows = i < 0 ? [] : v.slice(1).filter(r => String(r[i] || "").trim());
  const pins = {}; if (p >= 0) rows.forEach(r => { const x = normPin_(r[p]); if (x) pins[String(r[i]).trim()] = x; });
  const ap = head.indexOf("admin pin"), admin = (ap >= 0 && v.length > 1) ? normPin_(v[1][ap]) : "";
  const ec = head.indexOf("emp code"), ds = head.indexOf("designation"), profiles = {};
  rows.forEach(r => { profiles[String(r[i]).trim()] = { empCode: ec >= 0 ? String(r[ec] || "").trim() : "", designation: ds >= 0 ? String(r[ds] || "").trim() : "" }; });
  return { technicians: rows.map(r => String(r[i]).trim()), pins: pins, admin: admin, profiles: profiles };
}

// SECURITY: both checks FAIL CLOSED. A missing/blank PIN or an unknown technician name is never accepted.
function adminOk_(pin) {
  const a = readLists_().admin;
  return a !== "" && a === normPin_(pin);
}
function pinCheck_(pins, tech, pin) {            // "ok" | "name" (not in the Lists sheet) | "pin" (wrong or missing PIN)
  const key = normName_(tech);
  if (!key) return "name";
  const given = normPin_(pin);
  for (const name in pins) {
    if (normName_(name) === key) return (given && pins[name] === given) ? "ok" : "pin";
  }
  return "name";
}
function pinOk_(pins, tech, pin) { return pinCheck_(pins, tech, pin) === "ok"; }

// ---- run ONCE from the Apps Script editor: writes the technician list, their PINs and the admin PIN ----
// !! Before running, replace CHANGE-ME below with your NEW admin PIN. Running this overwrites the whole Lists sheet.
function setupPins() {
  const sh = listsSheet_();
  sh.clear();
  const rows = [["Technicians", "PIN"], ["MEHEDI HASAN LIYAKAT ALI", "3765"], ["MOHAMMAD KAUSHAR ALI", "711312"], ["AJEESH MOHANKUMAR GEETHA", "712371"],
    ["NAGENDRA VISHWAKARMA", "712892"], ["RAM JATAN MAHADEV", "713052"], ["MUHAMMAD JAHANGEER", "713150"], ["ALI HASSAN", "713502"],
    ["IMTIYAZ KAMRUDDIN", "713520"], ["MADHAB CHARAN", "713542"]];
  sh.getRange("B:B").setNumberFormat("@"); sh.getRange("D:D").setNumberFormat("@");
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(1, 4, 2, 1).setValues([["Admin PIN"], ["CHANGE-ME"]]);
  sh.getRange("A1:D1").setFontWeight("bold"); sh.setColumnWidth(1, 200);
}

// ---- ONE-TIME REPAIR for readings saved with a different area / meter spelling than the Meters sheet ----
// (for example area "Common Area" where the Meters sheet says "Common").
// The Meters sheet is the source of truth. A reading is matched to a meter by meter number, then by a meter number
// that starts with the other, then by area name ignoring the word "Area".
// Step 1: run previewMeterFix() and read the log (View -> Logs). Nothing is changed.
// Step 2: if every line looks right, run applyMeterFix().
function meterFix_(apply) {
  const meters = readMeters_();
  if (!meters.length) { Logger.log("The Meters sheet is empty - nothing to match against."); return; }
  const sh = tab_("Readings", U_COLS), last = sh.getLastRow();
  if (last < 2) { Logger.log("No readings found."); return; }
  const iType = U_COLS.indexOf("type"), iArea = U_COLS.indexOf("area"), iMeter = U_COLS.indexOf("meter"), iAcc = U_COLS.indexOf("account");
  const normArea = v => String(v || "").toLowerCase().replace(/\barea\b/g, "").replace(/\s+/g, " ").trim();
  const match = (type, meter, area) => {
    const cands = meters.filter(m => m.type === type);
    const mm = String(meter).trim().toUpperCase(), a = normArea(area);
    return cands.find(m => m.meter.toUpperCase() === mm)
      || cands.find(m => { const t = m.meter.toUpperCase(); return mm.length >= 6 && t.length >= 6 && (mm.indexOf(t) === 0 || t.indexOf(mm) === 0); })
      || cands.find(m => normArea(m.area) === a);
  };
  const data = sh.getRange(2, 1, last - 1, U_COLS.length).getValues();
  const changes = {}, unmatched = {};
  let changed = 0;
  data.forEach(row => {
    const type = String(row[iType]).trim(), area = String(row[iArea]).trim(), meter = String(row[iMeter]).trim(), acc = String(row[iAcc]).trim();
    const m = match(type, meter, area);
    if (!m) { const k = type + ' | area "' + area + '" | meter ' + meter; unmatched[k] = (unmatched[k] || 0) + 1; return; }
    if (area !== m.area || meter !== m.meter || acc !== m.account) {
      const k = type + ': area "' + area + '" / meter ' + meter + ' / account ' + acc + '   ->   area "' + m.area + '" / meter ' + m.meter + ' / account ' + m.account;
      changes[k] = (changes[k] || 0) + 1;
      changed++;
      row[iArea] = m.area; row[iMeter] = m.meter; row[iAcc] = m.account;
    }
  });
  Object.keys(changes).forEach(k => Logger.log((apply ? "FIXED " : "WOULD FIX ") + changes[k] + " row(s): " + k));
  Object.keys(unmatched).forEach(k => Logger.log("NOT MATCHED (left as is) " + unmatched[k] + " row(s): " + k));
  if (!changed) Logger.log("Nothing to fix: every reading already matches the Meters sheet.");
  if (apply && changed) {
    const block = data.map(r => [r[iArea], r[iMeter], r[iAcc]]);
    sh.getRange(2, iArea + 1, data.length, 3).setNumberFormat("@").setValues(block);   // area, meter, account are neighbouring columns
  }
}
function previewMeterFix() { meterFix_(false); }
function applyMeterFix() { meterFix_(true); }

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
  if (p.action === "lists") { const l = readLists_(); return out_({ ok: true, lists: { technicians: l.technicians, pinRequired: true, meters: readMeters_(), last: lastReadings_() } }); }
  if (p.action === "version") return out_({ ok: true, version: SCRIPT_VERSION });
  if (p.action === "meters") return out_({ ok: true, meters: readMeters_() });                  // meter list from the Meters sheet
  if (p.action === "techs") return out_({ ok: true, technicians: readLists_().technicians });   // names only, no PINs
  if (p.action === "adminlogin") return out_({ ok: true, valid: adminOk_(p.apin) });
  if (p.action === "login") return out_({ ok: true, valid: pinOk_(readLists_().pins, p.name, p.pin) });
  if (p.action === "mydaily") {                  // a technician's own logs + overtime for one day
    if (!pinOk_(readLists_().pins, p.name, p.pin)) return out_({ ok: false, error: "login required" });
    const day = p.date || Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd");
    const key = normName_(p.name), q = { from: day, to: day };
    const mine = r => normName_(r.technician) === key;
    const tz0 = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const since = Utilities.formatDate(new Date(Date.now() - 90 * 86400000), tz0, "yyyy-MM-dd");
    const recent = read_("Logs", COLS, { from: since, to: day }).filter(mine);
    return out_({ ok: true,
      logs: recent.filter(r => String(r.logDate) === day),
      open: recent.filter(r => String(r.status || "Pending") !== "Completed"),
      overtime: read_("Overtime", OT_COLS, q).filter(mine) });
  }
  if (p.action === "myschedules") {
    if (!pinOk_(readLists_().pins, p.name, p.pin)) return out_({ ok: false, error: "login required" });
    const key = normName_(p.name);
    return out_({ ok: true, rows: read_("Schedules", S_COLS, {}).filter(r => normName_(r.technician) === key) });
  }
  // everything below returns data, so it requires the admin PIN
  if (!adminOk_(p.apin)) return out_({ ok: false, error: "admin PIN required" });
  if (p.action === "techlist") {                 // admin only: names + profile, never the PINs
    const l = readLists_();
    return out_({ ok: true, rows: l.technicians.map(n => ({ name: n, hasPin: !!l.pins[n], empCode: (l.profiles[n] || {}).empCode || "", designation: (l.profiles[n] || {}).designation || "" })) });
  }
  if (p.action === "allschedules") return out_({ ok: true, rows: read_("Schedules", S_COLS, {}) });
  if (p.action === "readings") return out_({ ok: true, rows: read_("Readings", U_COLS, p) });
  if (p.action === "overtime") return out_({ ok: true, rows: read_("Overtime", OT_COLS, p) });
  return out_({ ok: true, rows: read_("Logs", COLS, p) });
}
