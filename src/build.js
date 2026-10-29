// Building Word / Excel / PowerPoint files from the model's Markdown.
// The model asks for a file with a block:
//   <file type="docx|xlsx|pptx" name="Имя.docx"> ...markdown... </file>

import { createZip } from "./zip.js";
import { xml, prepare, inlineRuns, parseBlocks, isTableRow, isSeparator, splitRow } from "./markdown.js";
import { buildPptx, parseDeck, THEMES } from "./slides.js";

export { xml, inlineRuns, parseBlocks, buildPptx, parseDeck, THEMES };

export const MIME = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
};
const DEFAULT_NAMES = { docx: "Документ.docx", xlsx: "Таблица.xlsx", pptx: "Презентация.pptx" };

/* ------------------------------ helpers ------------------------------ */

export function safeFileName(name, type) {
  let n = String(name || "").replace(/[\\/:*?"<>|\u0000-\u001F]/g, "").replace(/\s+/g, " ").trim().slice(0, 80);
  if (!n) n = DEFAULT_NAMES[type];
  if (!n.toLowerCase().endsWith("." + type)) n = n.replace(/\.[A-Za-z0-9]{2,5}$/, "") + "." + type;
  return n;
}

/** Extracts <file …>…</file> blocks from the answer. */
export function extractFileBlocks(answer) {
  const files = [];
  const re = /(?:```[a-z]*\s*\n)?<file\s+([^>]*)>\s*\n?([\s\S]*?)(?:<\/file>|$)(?:\s*\n```)?/gi;
  const text = String(answer).replace(re, (_, attrs, content) => {
    const attr = k => new RegExp(`${k}\\s*=\\s*"([^"]*)"`, "i").exec(attrs)?.[1];
    const type = (attr("type") || "").toLowerCase();
    if (!MIME[type]) return _;
    files.push({
      type, name: safeFileName(attr("name"), type), content: content.replace(/\s+$/, ""),
      font: attr("font"), size: Number(attr("size")) || undefined, theme: attr("theme"),
    });
    return `\n[[FILE:${files.length - 1}]]\n`;
  });
  return { text: text.replace(/\n?\[\[FILE:\d+\]\]\n?/g, "\n").replace(/\n{3,}/g, "\n\n").trim(), files };
}

/* ------------------------------ Word ------------------------------ */

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function wRuns(text, extra = {}) {
  return inlineRuns(text, extra).map(r => {
    const pr = [
      r.code ? '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>' : "",
      r.b ? "<w:b/><w:bCs/>" : "", r.i ? "<w:i/><w:iCs/>" : "", r.s ? "<w:strike/>" : "",
    ].join("");
    const parts = r.text.split("\t").map(t => `<w:t xml:space="preserve">${xml(t)}</w:t>`).join("<w:tab/>");
    return `<w:r>${pr ? `<w:rPr>${pr}</w:rPr>` : ""}${parts}</w:r>`;
  }).join("");
}

function wPara(runs, { style, numId, ilvl = 0, align, extraPPr = "" } = {}) {
  const ppr = [
    style ? `<w:pStyle w:val="${style}"/>` : "",
    numId ? `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>` : "",
    extraPPr,
    align ? `<w:jc w:val="${align}"/>` : "",
  ].join("");
  return `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ""}${runs}</w:p>`;
}

function wTable(rows) {
  const cols = Math.max(...rows.map(r => r.length));
  const width = 9355; // text width in twips for A4 with our margins
  const colW = Math.floor(width / cols);
  const grid = Array.from({ length: cols }, () => `<w:gridCol w:w="${colW}"/>`).join("");
  const trs = rows.map((r, ri) => {
    const tcs = Array.from({ length: cols }, (_, ci) => {
      const shade = ri === 0 ? '<w:shd w:val="clear" w:color="auto" w:fill="E7EEF7"/>' : "";
      const text = r[ci] ?? "";
      return `<w:tc><w:tcPr><w:tcW w:w="${colW}" w:type="dxa"/>${shade}</w:tcPr>${wPara(wRuns(text, ri === 0 ? { b: true } : {}), { style: "TableText" })}</w:tc>`;
    }).join("");
    return `<w:tr>${ri === 0 ? "<w:trPr><w:tblHeader/></w:trPr>" : ""}${tcs}</w:tr>`;
  }).join("");
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="${colW * cols}" w:type="dxa"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${trs}</w:tbl>${wPara("")}`;
}

function wStyles(font, size) {
  const sz = Math.round(size * 2);
  const h = (id, name, pt, before, color = "1F3864") =>
    `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${before}" w:after="120"/><w:outlineLvl w:val="${Number(id.slice(-1)) - 1 || 0}"/></w:pPr><w:rPr><w:b/><w:bCs/><w:color w:val="${color}"/><w:sz w:val="${Math.round(pt * 2)}"/><w:szCs w:val="${Math.round(pt * 2)}"/></w:rPr></w:style>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W_NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${xml(font)}" w:hAnsi="${xml(font)}" w:eastAsia="${xml(font)}" w:cs="${xml(font)}"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/><w:lang w:val="ru-RU" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="240"/><w:jc w:val="center"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="${Math.round(size * 2 * 1.8)}"/></w:rPr></w:style>
${h("Heading1", "heading 1", size * 1.45, 360)}${h("Heading2", "heading 2", size * 1.2, 240, "2F5496")}${h("Heading3", "heading 3", size * 1.05, 200, "2F5496")}
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="60"/><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:sz w:val="${Math.max(16, sz - 4)}"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="567"/><w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="A5A5A5"/></w:pBdr></w:pPr><w:rPr><w:i/><w:color w:val="404040"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="TableText"><w:name w:val="Table Text"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="40" w:after="40" w:line="240" w:lineRule="auto"/></w:pPr></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="8EA9C1"/><w:left w:val="single" w:sz="4" w:space="0" w:color="8EA9C1"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="8EA9C1"/><w:right w:val="single" w:sz="4" w:space="0" w:color="8EA9C1"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="8EA9C1"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="8EA9C1"/></w:tblBorders></w:tblPr></w:style>
</w:styles>`;
}

function wNumbering(numbered) {
  const lvls = (fmt) => Array.from({ length: 5 }, (_, l) => {
    const bullet = ["•", "◦", "▪", "•", "◦"][l];
    const text = fmt === "bullet" ? bullet : `%${l + 1}.`;
    const font = fmt === "bullet" ? '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>' : "";
    return `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="${fmt === "bullet" ? "bullet" : "decimal"}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 + l * 360}" w:hanging="360"/></w:pPr>${font}</w:lvl>`;
  }).join("");
  const nums = [`<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`];
  numbered.forEach((start, k) => nums.push(
    `<w:num w:numId="${k + 2}"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="${start}"/></w:lvlOverride></w:num>`));
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering ${W_NS}><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${lvls("bullet")}</w:abstractNum><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${lvls("decimal")}</w:abstractNum>${nums.join("")}</w:numbering>`;
}

export function buildDocx(markdown, { font = "Calibri", size = 11, title } = {}) {
  size = Math.min(Math.max(Number(size) || 11, 8), 28);
  font = String(font || "Calibri").slice(0, 40);
  const blocks = parseBlocks(prepare(markdown));
  const body = [];
  const numbered = [];
  let currentNum = 0, prevType = "";
  let first = true;
  for (const b of blocks) {
    if (b.type !== "number") currentNum = b.type === "blank" ? currentNum : 0;
    switch (b.type) {
      case "heading": {
        const style = b.level === 1 && first ? "Title" : `Heading${Math.min(b.level === 1 ? 1 : b.level - (first ? 0 : 0), 3)}`;
        body.push(wPara(wRuns(b.text.replace(/\*\*/g, "")), { style }));
        break;
      }
      case "para": body.push(wPara(wRuns(b.text))); break;
      case "bullet": body.push(wPara(wRuns(b.text), { style: "ListParagraph", numId: 1, ilvl: b.level })); break;
      case "number": {
        if (!currentNum || prevType !== "number" && prevType !== "blank" && prevType !== "bullet") {
          numbered.push(b.n || 1);
          currentNum = numbered.length + 1;
        }
        body.push(wPara(wRuns(b.text), { style: "ListParagraph", numId: currentNum, ilvl: b.level }));
        break;
      }
      case "quote": body.push(wPara(wRuns(b.text), { style: "Quote" })); break;
      case "code": for (const l of b.lines.length ? b.lines : [""]) body.push(wPara(`<w:r><w:t xml:space="preserve">${xml(l)}</w:t></w:r>`, { style: "Code" })); body.push(wPara("")); break;
      case "table": if (b.rows.length) body.push(wTable(b.rows)); break;
      case "rule": body.push(wPara("", { extraPPr: '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="A5A5A5"/></w:pBdr>' })); break;
      case "blank": break;
    }
    if (b.type !== "blank") { first = false; prevType = b.type; }
  }
  if (!body.length) body.push(wPara(""));
  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W_NS}><w:body>${body.join("")}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  const now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
  return createZip([
    { name: "[Content_Types].xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>` },
    { name: "_rels/.rels", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>` },
    { name: "docProps/core.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${xml(title || "")}</dc:title><dc:creator>Telegram AI bot</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created></cp:coreProperties>` },
    { name: "word/document.xml", data: document },
    { name: "word/styles.xml", data: wStyles(font, size) },
    { name: "word/numbering.xml", data: wNumbering(numbered) },
    { name: "word/_rels/document.xml.rels", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>` },
  ]);
}

/* ------------------------------ Excel ------------------------------ */

function cellRef(c, r) {
  let s = "";
  for (let i = c + 1; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + ((i - 1) % 26)) + s;
  return s + (r + 1);
}

function safeSheetName(name, used) {
  let n = String(name || "Лист").replace(/[[\]:*?/\\]/g, "").trim().slice(0, 31) || "Лист";
  let k = 2, base = n;
  while (used.has(n.toLowerCase())) n = `${base.slice(0, 28)} ${k++}`;
  used.add(n.toLowerCase());
  return n;
}

/** Parses Markdown into sheets: "## Имя листа" + markdown tables (or CSV/TSV lines). */
export function parseSheets(md) {
  const sheets = [];
  let cur = null;
  const lines = String(md).replace(/\r\n?/g, "\n").split("\n");
  for (const line of lines) {
    const h = /^\s{0,3}#{1,6}\s+(?:Лист\s*[:«]\s*)?(.+?)»?\s*$/.exec(line);
    if (h) { cur = { name: h[1].replace(/\*\*/g, "").trim(), rows: [] }; sheets.push(cur); continue; }
    if (isTableRow(line)) {
      if (isSeparator(line)) continue;
      if (!cur) { cur = { name: "Лист1", rows: [] }; sheets.push(cur); }
      cur.rows.push(splitRow(line));
      continue;
    }
    if (line.includes("\t") || /;/.test(line) && !/[.!?]$/.test(line.trim())) {
      if (!cur) { cur = { name: "Лист1", rows: [] }; sheets.push(cur); }
      cur.rows.push(line.split(line.includes("\t") ? "\t" : ";").map(c => c.trim()));
    }
  }
  return sheets.filter(s => s.rows.length);
}

function xlsxCell(raw, c, r, header) {
  const ref = cellRef(c, r);
  let v = String(raw ?? "").trim();
  let bold = header;
  if (/^\*\*[\s\S]*\*\*$/.test(v)) { bold = true; v = v.slice(2, -2).trim(); }
  v = v.replace(/<br\s*\/?>/gi, "\n");
  if (!v) return "";
  if (!header && v.startsWith("=") && v.length > 1) {
    return `<c r="${ref}" s="${bold ? 4 : 3}"><f>${xml(v.slice(1).replace(/^=+/, ""))}</f></c>`;
  }
  const num = v.replace(/[\s ]/g, "").replace(/^−/, "-");
  if (!header && /^-?\d+(?:[.,]\d+)?%$/.test(num)) {
    return `<c r="${ref}" s="2"><v>${Number(num.slice(0, -1).replace(",", ".")) / 100}</v></c>`;
  }
  if (!header && /^-?\d+(?:[.,]\d+)?$/.test(num) && !/^0\d/.test(num) && num.replace(/\D/g, "").length <= 15) {
    return `<c r="${ref}" s="${bold ? 4 : 3}"><v>${Number(num.replace(",", "."))}</v></c>`;
  }
  return `<c r="${ref}" t="inlineStr" s="${header ? 1 : bold ? 4 : 3}"><is><t xml:space="preserve">${xml(v)}</t></is></c>`;
}

export function buildXlsx(markdown) {
  let sheets = parseSheets(prepare(markdown));
  if (!sheets.length) {
    const lines = String(markdown).split("\n").map(l => l.trim()).filter(Boolean);
    sheets = [{ name: "Лист1", rows: lines.map(l => [l]) }];
  }
  const used = new Set();
  const names = sheets.map(s => safeSheetName(s.name, used));
  const sheetXml = sheets.map(s => {
    const cols = Math.max(...s.rows.map(r => r.length));
    const widths = Array.from({ length: cols }, (_, c) =>
      Math.min(60, Math.max(8, ...s.rows.map(r => [...String(r[c] ?? "").replace(/\*\*/g, "")].length + 2))));
    const rows = s.rows.map((r, ri) => `<row r="${ri + 1}">${r.map((v, ci) => xlsxCell(v, ci, ri, ri === 0 && s.rows.length > 1)).join("")}</row>`).join("");
    const colsXml = widths.map((w, c) => `<col min="${c + 1}" max="${c + 1}" width="${w}" customWidth="1"/>`).join("");
    const freeze = s.rows.length > 1 ? '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' : "";
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${freeze}<cols>${colsXml}</cols><sheetData>${rows}</sheetData></worksheet>`;
  });
  const files = [
    { name: "[Content_Types].xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheetXml.map((_, k) => `<Override PartName="/xl/worksheets/sheet${k + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}</Types>` },
    { name: "_rels/.rels", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>` },
    { name: "xl/workbook.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>${names.map((n, k) => `<sheet name="${xml(n)}" sheetId="${k + 1}" r:id="rId${k + 1}"/>`).join("")}</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>` },
    { name: "xl/_rels/workbook.xml.rels", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${names.map((_, k) => `<Relationship Id="rId${k + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${k + 1}.xml"/>`).join("")}<Relationship Id="rId${names.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>` },
    { name: "xl/styles.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFB4C6E7"/></left><right style="thin"><color rgb="FFB4C6E7"/></right><top style="thin"><color rgb="FFB4C6E7"/></top><bottom style="thin"><color rgb="FFB4C6E7"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="10" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>` },
  ];
  sheetXml.forEach((x, k) => files.push({ name: `xl/worksheets/sheet${k + 1}.xml`, data: x }));
  return createZip(files);
}

/* ------------------------------ entry ------------------------------ */

export function buildFile(block) {
  if (block.type === "docx") return buildDocx(block.content, { font: block.font, size: block.size, title: block.name.replace(/\.docx$/i, "") });
  if (block.type === "xlsx") return buildXlsx(block.content);
  if (block.type === "pptx") return buildPptx(block.content, { theme: block.theme });
  throw new Error("unknown_file_type");
}
