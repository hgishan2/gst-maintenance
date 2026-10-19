// Fills your original Excel template: one sheet per day, all days in ONE workbook.
const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const sheetName = d => { const [y,m,dd] = d.split("-"); return dd + "-" + MONTHS[Number(m)-1] + "-" + y; };
function toFrac(t){ if(!t) return null; const p=String(t).split(":"); return (Number(p[0])*60+Number(p[1]))/1440; }
function hrs(s,f){ if(s==null||f==null) return ""; return Math.round((((f-s)%1)+1)%1*24*100)/100; }

function fillSheet(ws, rows, meta) {
  ws.getCell("C3").value = meta.building;
  const [y,m,d] = meta.logDate.split("-").map(Number);
  ws.getCell("E3").value = new Date(Date.UTC(y, m-1, d));
  ws.getCell("E3").numFmt = "dd-mmm-yyyy";
  ws.getCell("G3").value = meta.supervisor || null;
  ws.getCell("I3").value = meta.shift || null;

  const FIRST = 5, TMPL_LAST = 23;
  const lastRow = Math.max(TMPL_LAST, FIRST + rows.length - 1);
  for (let r = TMPL_LAST + 1; r <= lastRow; r++) {            // extend table if more than 19 jobs
    ws.getRow(r).height = ws.getRow(TMPL_LAST).height;
    for (let c = 1; c <= 16; c++)
      ws.getRow(r).getCell(c).style = JSON.parse(JSON.stringify(ws.getRow(TMPL_LAST).getCell(c).style));
  }
  for (let r = FIRST; r <= lastRow; r++) {
    const e = rows[r - FIRST], row = ws.getRow(r);
    const s = e ? toFrac(e.start) : null, f = e ? toFrac(e.finish) : null;
    row.getCell(1).value  = { formula: `IF(D${r}="","",ROW()-4)`, result: e ? r - 4 : "" };   // S.No. (template had ROW()-8)
    row.getCell(12).value = { formula: `IF(OR(J${r}="",K${r}=""),"",ROUND(MOD(K${r}-J${r},1)*24,2))`, result: hrs(s, f) };
    if (!e) continue;
    row.getCell(2).value  = e.wo || null;
    row.getCell(3).value  = e.location || null;
    row.getCell(4).value  = e.complaint || null;
    row.getCell(5).value  = e.technician || null;
    row.getCell(6).value  = e.action || null;
    row.getCell(7).value  = e.category || null;
    row.getCell(8).value  = e.priority || null;
    row.getCell(9).value  = toFrac(e.received);
    row.getCell(10).value = s;
    row.getCell(11).value = f;
    row.getCell(13).value = e.materials || null;
    row.getCell(14).value = e.status || null;
    row.getCell(15).value = e.pending || null;
    row.getCell(16).value = e.remarks || null;
  }
  return lastRow;
}

function cloneSheet(wb, src, name) {            // exact copy of the (still empty) template sheet
  const ws = wb.addWorksheet(name);
  const model = JSON.parse(JSON.stringify(src.model));
  model.name = name; model.mergeCells = model.merges;
  ws.model = model; ws.name = name;
  return ws;
}

// ExcelJS writes overlapping dropdown ranges (Excel then says "repair"); replace with one clean block per sheet.
async function fixDropdowns(JSZip, buf, lastRows) {
  const LISTS = { G: "AC / HVAC,Plumbing,Electrical,Civil / Painting,Fire &amp; Life Safety,Carpentry,General Maintenance,PPM,Other",
                  H: "Emergency,High,Medium,Low", N: "Completed,Pending,In Progress,On Hold" };
  const zip = await JSZip.loadAsync(buf);
  for (let i = 0; i < lastRows.length; i++) {
    const f = `xl/worksheets/sheet${i + 1}.xml`;
    let xml = await zip.file(f).async("string");
    const dv = '<dataValidations count="3">' + Object.keys(LISTS).map(c =>
      `<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="${c}5:${c}${lastRows[i]}"><formula1>"${LISTS[c]}"</formula1></dataValidation>`).join("") + "</dataValidations>";
    xml = xml.replace(/<dataValidations[\s\S]*?<\/dataValidations>/, "");
    xml = xml.replace("<pageMargins", dv + "<pageMargins");
    zip.file(f, xml);
  }
  return zip.generateAsync({ type: "uint8array" });
}

// days = [{date:"YYYY-MM-DD", rows:[...]}] sorted by date; meta = {building, shift, supervisor(optional override)}
const startKey = r => { const f = toFrac(r.start); return f == null ? 9 : (f < 0.25 ? f + 1 : f); };
const byStart = rows => rows.map((r, i) => [r, i]).sort((a, b) => startKey(a[0]) - startKey(b[0]) || a[1] - b[1]).map(x => x[0]);

async function buildWorkbook(ExcelJS, JSZip, templateBuf, days, meta) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(templateBuf);
  const base = wb.worksheets[0];
  const sheets = [base];
  for (let i = 1; i < days.length; i++) sheets.push(cloneSheet(wb, base, sheetName(days[i].date)));
  base.name = sheetName(days[0].date);
  const lastRows = days.map((day, i) => {
    const sup = meta.supervisor || [...new Set(day.rows.map(r => String(r.supervisor || "").trim()).filter(Boolean))].join(", ");
    return fillSheet(sheets[i], byStart(day.rows), { building: meta.building, shift: meta.shift, logDate: day.date, supervisor: sup });
  });
  const buf = await wb.xlsx.writeBuffer();
  return fixDropdowns(JSZip, buf, lastRows);
}
if (typeof module !== "undefined") module.exports = { buildWorkbook };
