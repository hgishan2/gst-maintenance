// Utility readings report: Summary + one sheet per type (Electricity / Water) + Daily Readings (with consumption formulas).
const dObj = s => { const [y, m, d] = s.split("-").map(Number); return new Date(Date.UTC(y, m - 1, d)); };
const addDays = (s, n) => { const d = dObj(s); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); };
const NAVY = "FF0F2234", GRID = { style: "thin", color: { argb: "FFD0D7DE" } };
const DR = "'Daily Readings'!";

async function utilBuild(ExcelJS, rows, meters, from, to, building) {
  const best = {};                                            // one reading per meter per day (latest entry wins)
  rows.forEach(r => { const v = parseFloat(String(r.reading).replace(/,/g, "")); if (isNaN(v)) return;
    const k = r.meter + "|" + r.date, o = best[k]; if (!o || String(r.submittedAt) >= String(o.submittedAt)) best[k] = Object.assign({}, r, { val: v }); });
  const per = {}; Object.values(best).forEach(r => (per[r.meter] = per[r.meter] || []).push(r));
  Object.values(per).forEach(a => a.sort((x, y) => x.date.localeCompare(y.date)));
  const flat = [];
  meters.forEach(m => (per[m.meter] || []).forEach((r, i, a) => { if (r.date >= from && r.date <= to) flat.push({ m, r, p: i > 0 ? a[i - 1] : null }); }));
  if (!flat.length) return null;
  const key = m => m.type + "|" + m.area, at = {};
  flat.forEach(f => { at[key(f.m) + "|" + f.r.date] = f; });
  const days = f => f.p ? Math.round((dObj(f.r.date) - dObj(f.p.date)) / 864e5) : "";
  const cons = f => f.p ? Math.round((f.r.val - f.p.val) * 1e4) / 1e4 : "";

  const wb = new ExcelJS.Workbook();
  const sSum = wb.addWorksheet("Summary");
  const types = [...new Set(meters.filter(m => flat.some(f => f.m === m)).map(m => m.type))];
  const sMat = types.map(t => wb.addWorksheet(t));
  const ws = wb.addWorksheet("Daily Readings");
  const head = (sh, row, from_, to_) => { for (let c = from_; c <= to_; c++) { const x = sh.getCell(row, c);
    x.font = { name: "Arial", bold: true, color: { argb: "FFFFFFFF" }, size: 10 }; x.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
    x.alignment = { horizontal: "center", vertical: "middle", wrapText: true }; x.border = { top: GRID, bottom: GRID, left: GRID, right: GRID }; } };
  const body = (cell, fmt) => { cell.font = { name: "Arial", size: 10 }; cell.border = { top: GRID, bottom: GRID, left: GRID, right: GRID }; if (fmt) cell.numFmt = fmt; };
  const title = (sh, sub) => { sh.getCell("A1").value = "UTILITY READINGS – " + building; sh.getCell("A1").font = { name: "Arial", bold: true, size: 14, color: { argb: NAVY } };
    sh.getCell("A2").value = sub + "  |  " + from + " to " + to; sh.getCell("A2").font = { name: "Arial", size: 10, color: { argb: "FF5E7083" } }; };
  const NUM = "#,##0.00", DATE = "dd-mmm-yyyy";

  // ---------- Daily Readings ----------
  title(ws, "Daily readings and consumption (today's reading minus previous reading)");
  ["Date","Type","Unit / Area","Meter No","Account No","Reading","Previous reading","Previous date","Days","Consumption","Entered by","Check","Key"].forEach((h, i) => ws.getCell(4, i + 1).value = h);
  head(ws, 4, 1, 13);
  flat.forEach((f, i) => {
    const r = 5 + i, c = cons(f), d = days(f);
    const set = (col, v, fmt) => { const x = ws.getCell(r, col); x.value = v; body(x, fmt); };
    set(1, dObj(f.r.date), DATE); set(2, f.m.type); set(3, f.m.area, "@"); set(4, f.m.meter, "@"); set(5, f.m.account, "@");
    set(6, f.r.val, NUM); set(7, f.p ? f.p.val : null, NUM); set(8, f.p ? dObj(f.p.date) : null, DATE);
    set(9, { formula: `IF(H${r}="","",A${r}-H${r})`, result: d }, "0");
    set(10, { formula: `IF(G${r}="","",F${r}-G${r})`, result: c }, NUM);
    set(11, f.r.technician || ""); 
    set(12, { formula: `IF(J${r}="","",IF(J${r}<0,"Lower than previous - check",IF(I${r}>1,I${r}&"-day gap","")))`, result: c === "" ? "" : c < 0 ? "Lower than previous - check" : d > 1 ? d + "-day gap" : "" });
    set(13, key(f.m));
    ws.getCell(r, 10).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEAF2F8" } };
  });
  const last = 4 + flat.length;
  ws.addConditionalFormatting({ ref: `J5:J${last}`, rules: [{ type: "cellIs", operator: "lessThan", formulae: [0], style: { font: { color: { argb: "FFC00000" }, bold: true } } }] });
  [12, 12, 13, 14, 15, 14, 16, 13, 7, 14, 24, 28, 22].forEach((w, i) => ws.getColumn(i + 1).width = w);
  ws.views = [{ state: "frozen", ySplit: 4 }]; ws.getColumn(13).hidden = true;

  // ---------- Summary ----------
  title(sSum, "Consumption summary per meter");
  ["Type","Unit / Area","Meter No","Account No","Opening reading","Closing reading","Total consumption","Readings in period","Avg per day"].forEach((h, i) => sSum.getCell(4, i + 1).value = h);
  head(sSum, 4, 1, 9);
  meters.forEach((m, i) => {
    const r = 5 + i, fl = flat.filter(f => f.m === m), k = key(m), tot = fl.reduce((s, f) => s + (f.p ? f.r.val - f.p.val : 0), 0), dd = fl.reduce((s, f) => s + (f.p ? days(f) : 0), 0);
    const set = (col, v, fmt) => { const x = sSum.getCell(r, col); x.value = v; body(x, fmt); };
    set(1, m.type); set(2, m.area, "@"); set(3, m.meter, "@"); set(4, m.account, "@");
    set(5, fl.length ? (fl[0].p ? fl[0].p.val : fl[0].r.val) : null, NUM); set(6, fl.length ? fl[fl.length - 1].r.val : null, NUM);
    set(7, { formula: `SUMIFS(${DR}$J:$J,${DR}$M:$M,"${k}")`, result: Math.round(tot * 1e4) / 1e4 }, NUM);
    set(8, { formula: `COUNTIFS(${DR}$M:$M,"${k}")`, result: fl.length }, "0");
    set(9, { formula: `IFERROR(G${r}/SUMIFS(${DR}$I:$I,${DR}$M:$M,"${k}"),"")`, result: dd ? Math.round(tot / dd * 100) / 100 : "" }, NUM);
    sSum.getCell(r, 7).font = { name: "Arial", size: 10, bold: true };
  });
  const sn = 6 + meters.length;
  sSum.getCell(sn, 1).value = "Consumption is today's reading minus the previous reading. Where a day was skipped, the next reading covers several days (see 'Days' and 'Check' in Daily Readings).";
  sSum.getCell(sn, 1).font = { name: "Arial", size: 9, italic: true, color: { argb: "FF5E7083" } };
  [14, 14, 16, 16, 16, 16, 18, 16, 14].forEach((w, i) => sSum.getColumn(i + 1).width = w);
  sSum.views = [{ state: "frozen", ySplit: 4 }];

  // ---------- Electricity / Water: date x meter ----------
  const dates = []; for (let d = from; d <= to; d = addDays(d, 1)) dates.push(d);
  types.forEach((t, ti) => {
    const sh = sMat[ti], ms = meters.filter(m => m.type === t);
    title(sh, t + " readings and daily consumption");
    sh.getCell(4, 1).value = "Date"; sh.mergeCells(4, 1, 5, 1);
    ms.forEach((m, j) => { const c0 = 2 + 2 * j; sh.getCell(4, c0).value = m.area + "  (" + m.meter + ")"; sh.mergeCells(4, c0, 4, c0 + 1);
      sh.getCell(5, c0).value = "Reading"; sh.getCell(5, c0 + 1).value = "Consumption"; sh.getColumn(c0).width = 14; sh.getColumn(c0 + 1).width = 15; });
    head(sh, 4, 1, 1 + 2 * ms.length); head(sh, 5, 1, 1 + 2 * ms.length); sh.getColumn(1).width = 14;
    const L = n => { let s = ""; while (n > 0) { s = String.fromCharCode(65 + (n - 1) % 26) + s; n = Math.floor((n - 1) / 26); } return s; };
    dates.forEach((d, i) => {
      const r = 6 + i, a = sh.getCell(r, 1); a.value = dObj(d); body(a, DATE);
      ms.forEach((m, j) => {
        const k = key(m), f = at[k + "|" + d], c0 = 2 + 2 * j;
        const rd = sh.getCell(r, c0), cs = sh.getCell(r, c0 + 1);
        rd.value = { formula: `IF(COUNTIFS(${DR}$M:$M,"${k}",${DR}$A:$A,$A${r})=0,"",SUMIFS(${DR}$F:$F,${DR}$M:$M,"${k}",${DR}$A:$A,$A${r}))`, result: f ? f.r.val : "" };
        cs.value = { formula: `IF(COUNTIFS(${DR}$M:$M,"${k}",${DR}$A:$A,$A${r},${DR}$J:$J,">-1E+300")=0,"",SUMIFS(${DR}$J:$J,${DR}$M:$M,"${k}",${DR}$A:$A,$A${r}))`, result: f ? cons(f) : "" };
        body(rd, NUM); body(cs, NUM); cs.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEAF2F8" } };
      });
    });
    const tr = 6 + dates.length; sh.getCell(tr, 1).value = "Total"; body(sh.getCell(tr, 1)); sh.getCell(tr, 1).font = { name: "Arial", bold: true, size: 10 };
    ms.forEach((m, j) => { const c1 = 3 + 2 * j, col = L(c1), x = sh.getCell(tr, c1);
      const tot = dates.reduce((s, d) => { const f = at[key(m) + "|" + d]; return s + (f && f.p ? f.r.val - f.p.val : 0); }, 0);
      x.value = { formula: `SUM(${col}6:${col}${tr - 1})`, result: Math.round(tot * 1e4) / 1e4 }; body(x, NUM); x.font = { name: "Arial", bold: true, size: 10 };
      body(sh.getCell(tr, c1 - 1)); });
    sh.addConditionalFormatting({ ref: `B6:${L(1 + 2 * ms.length)}${tr - 1}`, rules: [{ type: "cellIs", operator: "lessThan", formulae: [0], style: { font: { color: { argb: "FFC00000" }, bold: true } } }] });
    sh.views = [{ state: "frozen", xSplit: 1, ySplit: 5 }];
  });
  [sSum, ...sMat, ws].forEach(sh => { sh.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 }; });
  return wb.xlsx.writeBuffer();
}
if (typeof module !== "undefined") module.exports = { utilBuild, addDays };
