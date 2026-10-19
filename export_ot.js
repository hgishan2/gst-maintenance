// Fills your Overtime template: one "Individual OT Sheet" per technician (per month), Compatibility Report kept as the last sheet.
const MONTHS_UP = ["JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY","AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER"];
const MON3 = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const frac = t => { const p = String(t).split(":"); return (Number(p[0]) * 60 + Number(p[1])) / 1440; };
const clone = o => JSON.parse(JSON.stringify(o));

function otClone(wb, src, name) {
  const ws = wb.addWorksheet(name), m = clone(src.model);
  m.name = name; m.mergeCells = m.merges; ws.model = m; ws.name = name;
  return ws;
}
const dayRows = ws => { const out = []; for (let r = 17; r < 400; r++) { const v = ws.getCell("A" + r).value; if (v === "Grand Total") break; if (typeof v === "number") out.push(r); } return out; };
const totalRow = ws => { for (let r = 17; r < 400; r++) if (ws.getCell("A" + r).value === "Grand Total") return r; };

function fillOt(ws, items, st, tech, ym) {
  const [y, m] = ym.split("-").map(Number);
  ws.getCell("A12").value = "Name: " + (st.full || tech).replace(/\s+/g, " ");
  ws.getCell("F12").value = "Emp Code: " + (st.code || "");
  ws.getCell("A13").value = "Designation: " + (st.designation || "");
  ws.getCell("H12").value = "Month : " + MONTHS_UP[m - 1] + " - " + y;

  const lower = ws.model.merges.filter(x => parseInt(x.match(/\d+/)[0], 10) > 50);   // merges below the table
  lower.forEach(x => ws.unMergeCells(x));
  const byDay = {}; items.forEach(e => (byDay[Number(e.otDate.slice(8, 10))] = byDay[Number(e.otDate.slice(8, 10))] || []).push(e));
  Object.keys(byDay).map(Number).sort((a, b) => b - a).forEach(d => {          // add rows only where a day has more entries than rows
    const rows = dayRows(ws).filter(r => ws.getCell("A" + r).value === d);
    for (let k = rows.length; k < byDay[d].length; k++) {
      const last = dayRows(ws).filter(r => ws.getCell("A" + r).value === d).pop();
      ws.spliceRows(last + 1, 0, []);
      for (let c = 1; c <= 8; c++) ws.getRow(last + 1).getCell(c).style = clone(ws.getRow(last).getCell(c).style);
      ws.getRow(last + 1).height = ws.getRow(last).height; ws.getCell("A" + (last + 1)).value = d;
    }
  });
  const ins = totalRow(ws) - 51;
  lower.forEach(x => ws.mergeCells(x.replace(/\d+/g, n => Number(n) + ins)));
  const col = { N: 4, WO: 5, PH: 6, NS: 7 }, sums = { 4: 0, 5: 0, 6: 0, 7: 0 };
  Object.keys(byDay).map(Number).forEach(d => {
    const rows = dayRows(ws).filter(r => ws.getCell("A" + r).value === d);
    byDay[d].sort((a, b) => String(a.from).localeCompare(String(b.from))).forEach((e, i) => {
      const r = rows[i], f = frac(e.from), t = frac(e.to), h = Math.round((((t - f) % 1) + 1) % 1 * 24 * 100) / 100, c = col[e.type] || 4;
      ws.getCell("B" + r).value = f; ws.getCell("C" + r).value = t;
      const L = "ABCDEFGH"[c - 1];
      ws.getCell(L + r).value = { formula: `IF(OR(B${r}="",C${r}=""),"",ROUND(MOD(C${r}-B${r},1)*24,2))`, result: h };
      ws.getCell(L + r).numFmt = "0.00"; sums[c] += h;
      const txt = (e.location + " – " + e.work).replace(/\s+/g, " ");
      const rc = ws.getCell("H" + r); rc.value = txt;
      rc.alignment = Object.assign({}, rc.alignment, { wrapText: true, vertical: "middle" });
      ws.getRow(r).height = Math.max(13, 13 * Math.ceil(txt.length / 36));
    });
  });
  const T = totalRow(ws);
  [4, 5, 6, 7].forEach(c => { const L = "ABCDEFGH"[c - 1], cell = ws.getCell(L + T);
    cell.value = { formula: `SUM(${L}17:${L}${T - 1})`, result: Math.round(sums[c] * 100) / 100 }; cell.numFmt = "0.00"; });
}

// rows = [{technician, otDate, from, to, type, location, work}]; staff = CONFIG.STAFF
async function otBuild(ExcelJS, templateBuf, rows, staff) {
  const wb = new ExcelJS.Workbook(); await wb.xlsx.load(templateBuf);
  const base = wb.getWorksheet("Individual OT Sheet"), compat = wb.worksheets.find(w => w !== base);
  const g = {}; rows.forEach(r => (g[r.technician + "|" + r.otDate.slice(0, 7)] = g[r.technician + "|" + r.otDate.slice(0, 7)] || []).push(r));
  const ord = n => { const i = staff.findIndex(s => s.name === n); return i < 0 ? 999 : i; };
  const keys = Object.keys(g).sort((a, b) => { const [ta, ma] = a.split("|"), [tb, mb] = b.split("|"); return ord(ta) - ord(tb) || ta.localeCompare(tb) || ma.localeCompare(mb); });
  const used = new Set(), sheets = [];
  keys.forEach((k, i) => {
    const [tech, ym] = k.split("|"), [y, m] = ym.split("-");
    let name = (tech.split(" ")[0] + " " + MON3[Number(m) - 1] + "-" + y.slice(2)).slice(0, 31), n = 2;
    while (used.has(name)) name = tech.slice(0, 22) + " " + MON3[Number(m) - 1] + "-" + y.slice(2) + (n++ > 2 ? n : "");
    used.add(name); sheets.push(i === 0 ? base : otClone(wb, base, name));
    sheets[i].__name = name;
  });
  sheets.forEach((ws, i) => { const [tech, ym] = keys[i].split("|"); ws.name = ws.__name; fillOt(ws, g[keys[i]], staff.find(s => s.name === tech) || {}, tech, ym); ws.orderNo = i + 1; });
  if (compat) compat.orderNo = 9999;
  return wb.xlsx.writeBuffer();
}
if (typeof module !== "undefined") module.exports = { otBuild };
