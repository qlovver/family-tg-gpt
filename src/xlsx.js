// Designed Excel workbooks from Markdown tables: coloured header, banded rows,
// number formats (thousands, ₽, %, dates), totals row, filters, frozen header,
// native charts (column, bar, line, pie).
//
// ## Расходы                      <- sheet name
// | Статья | Сумма, ₽ | Доля |
// |---|---|---|
// | Еда | 15000 | 30% |
// | Итого | =SUM(B2:B3) | =SUM(C2:C3) |
// @chart: pie; columns: B; title: Структура расходов

import { createZip } from "./zip.js";
import { xml, prepare, isTableRow, isSeparator, splitRow } from "./markdown.js";
import { resolveTheme, tint } from "./theme.js";

/* ------------------------------ parsing ------------------------------ */

function safeSheetName(name, used) {
  let n = String(name || "Лист").replace(/[[\]:*?/\\]/g, "").replace(/^'+|'+$/g, "").trim().slice(0, 31) || "Лист";
  let k = 2;
  const base = n;
  while (used.has(n.toLowerCase())) n = `${base.slice(0, 28)} ${k++}`;
  used.add(n.toLowerCase());
  return n;
}

function parseChart(spec) {
  const parts = spec.split(/[;|]/).map(s => s.trim()).filter(Boolean);
  const chart = { type: "column", columns: null, title: "" };
  for (const p of parts) {
    let m;
    if ((m = /^(?:type\s*[:=]\s*)?(column|columns|bar|bars|line|lines|pie|donut|столбцы|столбчатая|гистограмма|линия|линейная|график|круговая|пирог)$/i.exec(p))) {
      const t = m[1].toLowerCase();
      chart.type = /^(line|lines|линия|линейная|график)$/.test(t) ? "line" : /^(pie|donut|круговая|пирог)$/.test(t) ? "pie" : /^(bar|bars)$/.test(t) ? "bar" : "column";
    } else if ((m = /^(?:columns?|cols?|столбцы|данные)\s*[:=]\s*(.+)$/i.exec(p))) {
      chart.columns = m[1].split(/[\s,]+/).map(x => x.toUpperCase()).filter(x => /^[A-Z]{1,2}$/.test(x));
    } else if ((m = /^(?:title|название|заголовок)\s*[:=]\s*(.+)$/i.exec(p))) {
      chart.title = m[1];
    }
  }
  return chart;
}

/** Parses Markdown into sheets: "## Имя листа" + markdown table (+ "@chart:" lines). */
export function parseSheets(md) {
  const sheets = [];
  let cur = null;
  const ensure = () => { if (!cur) { cur = { name: "Лист1", rows: [], charts: [], notes: [] }; sheets.push(cur); } return cur; };
  for (const line of String(md).replace(/\r\n?/g, "\n").split("\n")) {
    let m;
    if ((m = /^\s{0,3}#{1,6}\s+(?:Лист\s*[:«]\s*)?(.+?)»?\s*$/.exec(line))) {
      cur = { name: m[1].replace(/\*\*/g, "").trim(), rows: [], charts: [], notes: [] };
      sheets.push(cur);
      continue;
    }
    if ((m = /^\s*@(?:chart|диаграмма|график)\s*:?\s*(.*)$/i.exec(line))) { ensure().charts.push(parseChart(m[1])); continue; }
    if (isTableRow(line)) { if (!isSeparator(line)) ensure().rows.push(splitRow(line)); continue; }
    if (line.includes("\t")) { ensure().rows.push(line.split("\t").map(c => c.trim())); continue; }
    if (/;/.test(line) && line.split(";").length >= 2 && !/[.!?]$/.test(line.trim()) && !/^\s*@/.test(line)) {
      ensure().rows.push(line.split(";").map(c => c.trim()));
      continue;
    }
    if (line.trim() && cur) cur.notes.push(line.trim());
  }
  return sheets.filter(s => s.rows.length);
}

/* ------------------------------ values ------------------------------ */

const MONEY_HEAD = /₽|руб|rub|\$|usd|€|eur|цена|стоимость|сумма|выручка|доход|расход|прибыль|бюджет|зарплата|оклад|платёж|платеж/i;

export function parseValue(raw) {
  let v = String(raw ?? "").trim();
  let bold = false;
  if (/^\*\*[\s\S]*\*\*$/.test(v)) { bold = true; v = v.slice(2, -2).trim(); }
  v = v.replace(/<br\s*\/?>/gi, "\n").replace(/\\\|/g, "|");
  if (!v) return { kind: "empty", bold };
  if (v.startsWith("=") && v.length > 1) return { kind: "formula", f: v.slice(1).replace(/^=+/, ""), bold };
  const n = v.replace(/[\s  ]/g, "").replace(/^−/, "-");
  let m;
  if ((m = /^(-?\d+(?:[.,]\d+)?)%$/.exec(n))) return { kind: "pct", n: Number(m[1].replace(",", ".")) / 100, bold };
  if ((m = /^(-?\d+(?:[.,]\d+)?)(₽|руб\.?|р\.)$/i.exec(n))) return { kind: "cur", n: Number(m[1].replace(",", ".")), bold };
  if ((m = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/.exec(n)) || (m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(n))) {
    const [y, mo, d] = m[1].length === 4 ? [m[1], m[2], m[3]] : [m[3], m[2], m[1]];
    const serial = (Date.UTC(+y, +mo - 1, +d) - Date.UTC(1899, 11, 30)) / 86400000;
    if (+mo >= 1 && +mo <= 12 && +d >= 1 && +d <= 31) return { kind: "date", n: serial, bold };
  }
  if (/^-?\d+(?:[.,]\d+)?$/.test(n) && !/^-?0\d/.test(n) && n.replace(/\D/g, "").length <= 15) {
    const num = Number(n.replace(",", "."));
    return { kind: Number.isInteger(num) ? "int" : "dec", n: num, bold };
  }
  return { kind: "text", s: v, bold };
}

function colName(i) {
  let s = "";
  for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}
const colIndex = letters => [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;

/* ------------------------------ styles ------------------------------ */

// cellXfs are generated from (format × variant) so every cell gets a consistent look.
const FORMATS = { text: 0, int: 3, dec: 4, pct: 164, cur: 165, curdec: 166, date: 167 };
const VARIANTS = ["body", "band", "total"];

function stylesXml(t) {
  const fmtKeys = Object.keys(FORMATS);
  const xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
    `<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>`];
  for (const v of VARIANTS) for (const f of fmtKeys) {
    const fill = v === "band" ? 3 : v === "total" ? 4 : 0;
    const font = v === "total" ? 2 : 0;
    const border = v === "total" ? 2 : 1;
    const align = f === "text" ? '<alignment vertical="center" wrapText="1"/>' : '<alignment vertical="center"/>';
    xfs.push(`<xf numFmtId="${FORMATS[f]}" fontId="${font}" fillId="${fill}" borderId="${border}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">${align}</xf>`);
  }
  const bold = '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>';
  const title = '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>';
  xfs.push(bold, title);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="4"><numFmt numFmtId="164" formatCode="0.0%"/><numFmt numFmtId="165" formatCode="#,##0 &quot;₽&quot;"/><numFmt numFmtId="166" formatCode="#,##0.00 &quot;₽&quot;"/><numFmt numFmtId="167" formatCode="dd.mm.yyyy"/></numFmts><fonts count="4"><font><sz val="11"/><color rgb="FF${t.text.length === 6 && !t.darkContent ? t.text : "1F2937"}"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><color rgb="FF${t.darkContent ? "1F2937" : t.heading}"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="14"/><color rgb="FF${t.darkContent ? "1F2937" : t.heading}"/><name val="Calibri"/><family val="2"/></font></fonts><fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF${t.darkContent ? "1E293B" : t.primary}"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF${tint(t.darkContent ? "334155" : t.primary, 0.07)}"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF${tint(t.accent, 0.2)}"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="3"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FF${tint(t.darkContent ? "334155" : t.primary, 0.22)}"/></left><right style="thin"><color rgb="FF${tint(t.darkContent ? "334155" : t.primary, 0.22)}"/></right><top style="thin"><color rgb="FF${tint(t.darkContent ? "334155" : t.primary, 0.22)}"/></top><bottom style="thin"><color rgb="FF${tint(t.darkContent ? "334155" : t.primary, 0.22)}"/></bottom><diagonal/></border><border><left style="thin"><color rgb="FF${tint(t.primary, 0.22)}"/></left><right style="thin"><color rgb="FF${tint(t.primary, 0.22)}"/></right><top style="medium"><color rgb="FF${t.darkContent ? "1E293B" : t.primary}"/></top><bottom style="thin"><color rgb="FF${tint(t.primary, 0.22)}"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join("")}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
}

function xfIndex(format, variant) {
  return 2 + VARIANTS.indexOf(variant) * Object.keys(FORMATS).length + Object.keys(FORMATS).indexOf(format);
}

/* ------------------------------ sheet ------------------------------ */

const TOTAL_RE = /^(итого|всего|сумма|total|итог|в среднем|среднее)(?!\p{L})/iu;

function columnFormats(rows) {
  const cols = Math.max(...rows.map(r => r.length));
  return Array.from({ length: cols }, (_, c) => {
    const head = String(rows[0]?.[c] ?? "");
    const kinds = rows.slice(1).map(r => parseValue(r[c]).kind).filter(k => k !== "empty" && k !== "formula");
    const count = k => kinds.filter(x => x === k).length;
    let fmt = "text";
    if (kinds.length) {
      const top = ["pct", "cur", "date", "dec", "int", "text"].sort((a, b) => count(b) - count(a))[0];
      fmt = top;
      if (top === "int" && count("dec")) fmt = "dec";
    } else if (/%|доля|процент/i.test(head)) fmt = "pct";
    else if (MONEY_HEAD.test(head)) fmt = "dec";
    if ((fmt === "int" || fmt === "dec") && /₽|руб/i.test(head)) fmt = fmt === "int" ? "cur" : "curdec";
    if (fmt === "cur" && kinds.some(k => k === "dec")) fmt = "curdec";
    return fmt;
  });
}

function sheetXml(sheet, t, drawingRel) {
  const rows = sheet.rows;
  const cols = Math.max(...rows.map(r => r.length));
  const header = rows.length > 1;
  const formats = columnFormats(rows);
  const widths = Array.from({ length: cols }, (_, c) => {
    const lens = rows.map((r, ri) => {
      const v = parseValue(r[c]);
      if (v.kind === "text") return Math.max(...String(v.s).split("\n").map(x => [...x].length));
      if (v.kind === "formula") return 12;
      return String(r[c] ?? "").length + (formats[c].startsWith("cur") ? 3 : 1);
    });
    return Math.min(50, Math.max(9, ...lens.map(l => l + 2)));
  });
  const totals = new Set();
  rows.forEach((r, ri) => { if (ri > 0 && TOTAL_RE.test(String(r[0] ?? "").replace(/\*\*/g, "").trim())) totals.add(ri); });

  const rowXml = rows.map((r, ri) => {
    const isHead = header && ri === 0;
    const variant = totals.has(ri) ? "total" : ri % 2 === 0 ? "band" : "body";
    const cells = Array.from({ length: cols }, (_, c) => {
      const ref = colName(c) + (ri + 1);
      const v = isHead ? { kind: "text", s: String(r[c] ?? "").replace(/\*\*/g, "") } : parseValue(r[c]);
      if (isHead) return `<c r="${ref}" t="inlineStr" s="1"><is><t xml:space="preserve">${xml(v.s)}</t></is></c>`;
      const fmt = v.kind === "text" ? "text" : v.kind === "empty" ? formats[c] : v.kind === "formula" ? formats[c] : v.kind === "pct" ? "pct" : v.kind === "date" ? "date" : v.kind === "cur" ? (formats[c] === "curdec" ? "curdec" : "cur") : formats[c] === "text" ? (v.kind === "dec" ? "dec" : "int") : formats[c];
      const s = xfIndex(fmt, variant);
      if (v.kind === "empty") return `<c r="${ref}" s="${s}"/>`;
      if (v.kind === "formula") return `<c r="${ref}" s="${s}"><f>${xml(v.f)}</f></c>`;
      if (v.kind === "text") return `<c r="${ref}" t="inlineStr" s="${xfIndex("text", variant)}"><is><t xml:space="preserve">${xml(v.s)}</t></is></c>`;
      return `<c r="${ref}" s="${s}"><v>${v.n}</v></c>`;
    }).join("");
    const ht = isHead ? ' ht="28" customHeight="1"' : "";
    return `<row r="${ri + 1}"${ht}>${cells}</row>`;
  }).join("");

  // notes below the table
  const notes = sheet.notes.slice(0, 20).map((n, k) =>
    `<row r="${rows.length + 2 + k}"><c r="A${rows.length + 2 + k}" t="inlineStr" s="${k === 0 && n.length < 60 ? 0 : 0}"><is><t xml:space="preserve">${xml(n.replace(/\*\*/g, ""))}</t></is></c></row>`).join("");

  const colsXml = widths.map((w, c) => `<col min="${c + 1}" max="${c + 1}" width="${w}" customWidth="1"/>`).join("");
  const views = header
    ? '<sheetViews><sheetView workbookViewId="0" zoomScale="110" zoomScaleNormal="110"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>'
    : '<sheetViews><sheetView workbookViewId="0"/></sheetViews>';
  const lastData = rows.length - [...totals].filter(x => x === rows.length - 1).length;
  const filter = header && rows.length > 3 ? `<autoFilter ref="A1:${colName(cols - 1)}${Math.max(2, lastData)}"/>` : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${views}<sheetFormatPr defaultRowHeight="18" customHeight="1"/><cols>${colsXml}</cols><sheetData>${rowXml}${notes}</sheetData>${filter}<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="${cols > 6 ? "landscape" : "portrait"}" fitToWidth="1" fitToHeight="0"/>${drawingRel ? `<drawing r:id="${drawingRel}"/>` : ""}</worksheet>`;
}

/* ------------------------------ charts ------------------------------ */

const C_NS = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function chartColors(t) {
  return t.darkContent ? ["38BDF8", "A78BFA", "F472B6", "34D399", "FBBF24", "94A3B8"]
    : [t.primary, t.accent, t.accent2, tint(t.primary, 0.55), tint(t.accent, 0.6), t.muted];
}

function chartXml(sheetName, sheet, chart, t) {
  const rows = sheet.rows;
  const cols = Math.max(...rows.map(r => r.length));
  const formats = columnFormats(rows);
  let last = rows.length;
  while (last > 2 && TOTAL_RE.test(String(rows[last - 1][0] ?? "").replace(/\*\*/g, ""))) last--;
  const numericCols = formats.map((f, c) => (c > 0 && f !== "text" && f !== "date" ? c : -1)).filter(c => c >= 0);
  let series = chart.columns ? chart.columns.map(colIndex).filter(c => c > 0 && c < cols) : numericCols;
  if (!series.length) series = numericCols.length ? numericCols : [1];
  if (chart.type === "pie") series = series.slice(0, 1);
  series = series.slice(0, 6);
  const q = `'${sheetName.replace(/'/g, "''")}'`;
  const cat = `<c:cat><c:strRef><c:f>${xml(q)}!$A$2:$A$${last}</c:f></c:strRef></c:cat>`;
  const colors = chartColors(t);
  const text = (sz, bold, color = "404040") => `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${sz}"${bold ? ' b="1"' : ""}><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:pPr><a:endParaRPr lang="ru-RU"/></a:p></c:txPr>`;
  const ser = series.map((c, k) => {
    const col = colName(c);
    const tx = `<c:tx><c:strRef><c:f>${xml(q)}!$${col}$1</c:f></c:strRef></c:tx>`;
    const val = `<c:val><c:numRef><c:f>${xml(q)}!$${col}$2:$${col}$${last}</c:f></c:numRef></c:val>`;
    const fill = `<a:solidFill><a:srgbClr val="${colors[k % colors.length]}"/></a:solidFill>`;
    if (chart.type === "line") {
      return `<c:ser><c:idx val="${k}"/><c:order val="${k}"/>${tx}<c:spPr><a:ln w="31750" cap="rnd"><a:solidFill><a:srgbClr val="${colors[k % colors.length]}"/></a:solidFill><a:round/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="7"/><c:spPr>${fill}<a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:marker>${cat}${val}<c:smooth val="0"/></c:ser>`;
    }
    if (chart.type === "pie") {
      const pts = Array.from({ length: Math.max(0, last - 1) }, (_, i) => `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${colors[i % colors.length]}"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`).join("");
      return `<c:ser><c:idx val="0"/><c:order val="0"/>${tx}${pts}<c:dLbls><c:numFmt formatCode="0%" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${text(1100, true, "FFFFFF")}<c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/><c:showLeaderLines val="0"/></c:dLbls>${cat}${val}</c:ser>`;
    }
    return `<c:ser><c:idx val="${k}"/><c:order val="${k}"/>${tx}<c:spPr>${fill}</c:spPr><c:invertIfNegative val="0"/>${cat}${val}</c:ser>`;
  }).join("");
  const axes = `<c:catAx><c:axId val="5001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${chart.type === "bar" ? "l" : "b"}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln></c:spPr>${text(1000)}<c:crossAx val="5002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx><c:valAx><c:axId val="5002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${chart.type === "bar" ? "b" : "l"}"/><c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="E7E7E7"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="${formats[series[0]] === "pct" ? "0%" : "#,##0"}" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>${text(1000)}<c:crossAx val="5001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  let plot;
  if (chart.type === "pie") plot = `<c:pieChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/></c:pieChart>`;
  else if (chart.type === "line") plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${ser}<c:marker val="1"/><c:axId val="5001"/><c:axId val="5002"/></c:lineChart>${axes}`;
  else plot = `<c:barChart><c:barDir val="${chart.type === "bar" ? "bar" : "col"}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${ser}<c:gapWidth val="70"/><c:overlap val="-10"/><c:axId val="5001"/><c:axId val="5002"/></c:barChart>${axes}`;
  const title = chart.title || (series.length === 1 ? String(rows[0][series[0]] ?? "") : sheetName);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace ${C_NS}><c:roundedCorners val="0"/><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="1"><a:solidFill><a:srgbClr val="${t.darkContent ? "1F2937" : t.heading}"/></a:solidFill></a:defRPr></a:pPr><a:r><a:rPr lang="ru-RU" sz="1400" b="1"><a:solidFill><a:srgbClr val="${t.darkContent ? "1F2937" : t.heading}"/></a:solidFill></a:rPr><a:t>${xml(title.replace(/\*\*/g, ""))}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>${plot}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/>${text(1000)}</c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="E5E5E5"/></a:solidFill></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:latin typeface="Calibri"/></a:defRPr></a:pPr><a:endParaRPr lang="ru-RU"/></a:p></c:txPr></c:chartSpace>`;
}

function drawingXml(anchors) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${anchors.map((a, k) => `<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${a.col}</xdr:col><xdr:colOff>152400</xdr:colOff><xdr:row>${a.row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>${a.col + 8}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${a.row + 17}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${k + 2}" name="Диаграмма ${k + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId${k + 1}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`).join("")}</xdr:wsDr>`;
}

/* ------------------------------ workbook ------------------------------ */

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export function buildXlsx(markdown, attrs = {}) {
  const t = resolveTheme(attrs);
  let sheets = parseSheets(prepare(markdown));
  if (!sheets.length) {
    const lines = String(markdown).split("\n").map(l => l.trim()).filter(Boolean);
    sheets = [{ name: "Лист1", rows: lines.map(l => [l]), charts: [], notes: [] }];
  }
  sheets = sheets.slice(0, 20);
  const used = new Set();
  const names = sheets.map(s => safeSheetName(s.name, used));
  const files = [];
  const overrides = [];
  let chartNo = 0, drawingNo = 0;
  sheets.forEach((s, k) => {
    const charts = s.charts.slice(0, 4);
    let drawingRel = null;
    if (charts.length && s.rows.length > 1) {
      drawingNo++;
      drawingRel = "rId1";
      const cols = Math.max(...s.rows.map(r => r.length));
      const anchors = charts.map((_, i) => ({ col: cols + 1, row: i * 18 }));
      const chartRels = [];
      charts.forEach((ch, i) => {
        chartNo++;
        files.push({ name: `xl/charts/chart${chartNo}.xml`, data: chartXml(names[k], s, ch, t) });
        overrides.push(`<Override PartName="/xl/charts/chart${chartNo}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`);
        chartRels.push(`<Relationship Id="rId${i + 1}" Type="${REL}/chart" Target="../charts/chart${chartNo}.xml"/>`);
      });
      files.push({ name: `xl/drawings/drawing${drawingNo}.xml`, data: drawingXml(anchors) });
      files.push({ name: `xl/drawings/_rels/drawing${drawingNo}.xml.rels`, data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${chartRels.join("")}</Relationships>` });
      files.push({ name: `xl/worksheets/_rels/sheet${k + 1}.xml.rels`, data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/drawing" Target="../drawings/drawing${drawingNo}.xml"/></Relationships>` });
      overrides.push(`<Override PartName="/xl/drawings/drawing${drawingNo}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`);
    }
    files.push({ name: `xl/worksheets/sheet${k + 1}.xml`, data: sheetXml(s, t, drawingRel) });
  });
  const defined = sheets.map((s, k) => {
    const cols = Math.max(...s.rows.map(r => r.length));
    return s.rows.length > 3 ? `<definedName name="_xlnm._FilterDatabase" localSheetId="${k}" hidden="1">'${xml(names[k].replace(/'/g, "''"))}'!$A$1:$${colName(cols - 1)}$${s.rows.length}</definedName>` : "";
  }).join("");
  return createZip([
    { name: "[Content_Types].xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_, k) => `<Override PartName="/xl/worksheets/sheet${k + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}${overrides.join("")}</Types>` },
    { name: "_rels/.rels", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>` },
    { name: "xl/workbook.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>${names.map((n, k) => `<sheet name="${xml(n)}" sheetId="${k + 1}" r:id="rId${k + 1}"/>`).join("")}</sheets>${defined ? `<definedNames>${defined}</definedNames>` : ""}<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>` },
    { name: "xl/_rels/workbook.xml.rels", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${names.map((_, k) => `<Relationship Id="rId${k + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${k + 1}.xml"/>`).join("")}<Relationship Id="rId${names.length + 1}" Type="${REL}/styles" Target="styles.xml"/></Relationships>` },
    { name: "xl/styles.xml", data: stylesXml(t) },
    ...files,
  ]);
}
