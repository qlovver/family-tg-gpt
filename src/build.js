// Building Word / Excel / PowerPoint files from the model's Markdown.
// The model asks for a file with a block:
//   <file type="docx|xlsx|pptx" name="Имя.docx"> ...markdown... </file>

import { createZip } from "./zip.js";
import { convertLatex } from "./format.js";
import * as T from "./pptx-template.js";

export const MIME = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
};
const DEFAULT_NAMES = { docx: "Документ.docx", xlsx: "Таблица.xlsx", pptx: "Презентация.pptx" };

/* ------------------------------ helpers ------------------------------ */

export function xml(s) {
  return String(s ?? "")
    // characters not allowed in XML 1.0
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, "")
    .replace(/[-]/g, "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

export function safeFileName(name, type) {
  let n = String(name || "").replace(/[\\/:*?"<>|\u0000-\u001F]/g, "").replace(/\s+/g, " ").trim().slice(0, 80);
  if (!n) n = DEFAULT_NAMES[type];
  if (!n.toLowerCase().endsWith("." + type)) n = n.replace(/\.[A-Za-z0-9]{2,5}$/, "") + "." + type;
  return n;
}

/** Markdown text with LaTeX converted to Unicode; boxed answers become **bold**. */
function prepare(md) {
  const parts = String(md).replace(/\r\n?/g, "\n").split(/(```[\s\S]*?(?:```|$))/);
  return parts.map((p, k) => (k % 2 ? p : convertLatex(p)))
    .join("")
    .replace(/[]/g, "**").replace(/[]/g, "*").replace(/[]/g, "~~");
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
      font: attr("font"), size: Number(attr("size")) || undefined,
    });
    return `\n[[FILE:${files.length - 1}]]\n`;
  });
  return { text: text.replace(/\n?\[\[FILE:\d+\]\]\n?/g, "\n").replace(/\n{3,}/g, "\n\n").trim(), files };
}

/** Inline Markdown -> [{text, b, i, code, s}] */
export function inlineRuns(line, base = {}) {
  const runs = [];
  const re = /\*\*(?=\S)([\s\S]+?)\*\*|`([^`]+)`|~~([^~]+)~~|(?<![\p{L}\p{N}*])\*(?=[^\s*])([^*\n]+?)\*(?![\p{L}\p{N}*])|(?<![\p{L}\p{N}_])_(?=[^\s_])([^_\n]+?)_(?![\p{L}\p{N}_])|\[([^\]]+)\]\((https?:\/\/[^)\s]+)\)/gu;
  let last = 0, m;
  while ((m = re.exec(line))) {
    if (m.index > last) runs.push({ ...base, text: line.slice(last, m.index) });
    if (m[1] !== undefined) runs.push(...inlineRuns(m[1], { ...base, b: true }));
    else if (m[2] !== undefined) runs.push({ ...base, text: m[2], code: true });
    else if (m[3] !== undefined) runs.push({ ...base, text: m[3], s: true });
    else if (m[4] !== undefined || m[5] !== undefined) runs.push(...inlineRuns(m[4] ?? m[5], { ...base, i: true }));
    else runs.push({ ...base, text: `${m[6]} (${m[7]})` });
    last = re.lastIndex;
  }
  if (last < line.length) runs.push({ ...base, text: line.slice(last) });
  return runs.filter(r => r.text);
}

const isTableRow = l => /^\s*\|.*\|\s*$/.test(l);
const isSeparator = l => /^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$/.test(l);
const splitRow = l => l.trim().replace(/^\|/, "").replace(/\|$/, "").split("|").map(c => c.trim());

/** Generic Markdown block parser shared by Word and PowerPoint. */
export function parseBlocks(md) {
  const lines = md.split("\n");
  const blocks = [];
  for (let k = 0; k < lines.length; k++) {
    const line = lines[k];
    if (/^\s*```/.test(line)) {
      const code = [];
      for (k++; k < lines.length && !/^\s*```/.test(lines[k]); k++) code.push(lines[k]);
      blocks.push({ type: "code", lines: code });
      continue;
    }
    if (isTableRow(line)) {
      const rows = [];
      for (; k < lines.length && isTableRow(lines[k]); k++) if (!isSeparator(lines[k])) rows.push(splitRow(lines[k]));
      k--;
      blocks.push({ type: "table", rows });
      continue;
    }
    let m;
    if ((m = /^\s{0,3}(#{1,6})\s+(.+?)\s*#*\s*$/.exec(line))) { blocks.push({ type: "heading", level: m[1].length, text: m[2] }); continue; }
    if (/^\s*([-*_])(\s*\1){2,}\s*$/.test(line)) { blocks.push({ type: "rule" }); continue; }
    if ((m = /^(\s*)[-*+•]\s+(.*)$/.exec(line))) { blocks.push({ type: "bullet", level: Math.min(Math.floor(m[1].replace(/\t/g, "  ").length / 2), 4), text: m[2] }); continue; }
    if ((m = /^(\s*)(\d{1,3})[.)]\s+(.*)$/.exec(line))) { blocks.push({ type: "number", level: Math.min(Math.floor(m[1].length / 2), 4), n: Number(m[2]), text: m[3] }); continue; }
    if ((m = /^\s*>\s?(.*)$/.exec(line))) { blocks.push({ type: "quote", text: m[1] }); continue; }
    if (!line.trim()) { blocks.push({ type: "blank" }); continue; }
    blocks.push({ type: "para", text: line });
  }
  return blocks;
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

/* ------------------------------ PowerPoint ------------------------------ */

const P_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function aRuns(text) {
  const runs = inlineRuns(text).map(r =>
    `<a:r><a:rPr lang="ru-RU" dirty="0"${r.b ? ' b="1"' : ""}${r.i ? ' i="1"' : ""}${r.s ? ' strike="sngStrike"' : ""}>${r.code ? '<a:latin typeface="Consolas"/>' : ""}</a:rPr><a:t>${xml(r.text)}</a:t></a:r>`).join("");
  return runs || '<a:endParaRPr lang="ru-RU" dirty="0"/>';
}

/** Markdown -> [{title, items: [{text, level, bullet}]}] */
export function parseSlides(md) {
  const slides = [];
  let cur = null;
  const push = () => { cur = { title: "", items: [] }; slides.push(cur); };
  for (const b of parseBlocks(md)) {
    if (b.type === "rule") { cur = null; continue; }
    if (b.type === "heading" && b.level <= 2) {
      if (!cur || cur.title || cur.items.length) push();
      cur.title = b.text.replace(/\*\*/g, "");
      continue;
    }
    if (!cur) push();
    if (b.type === "bullet") cur.items.push({ text: b.text, level: b.level, bullet: true });
    else if (b.type === "number") cur.items.push({ text: `${b.n}. ${b.text}`, level: b.level, bullet: false });
    else if (b.type === "heading") cur.items.push({ text: `**${b.text.replace(/\*\*/g, "")}**`, level: 0, bullet: false });
    else if (b.type === "para" || b.type === "quote") cur.items.push({ text: b.text, level: 0, bullet: false });
    else if (b.type === "table") for (const row of b.rows) cur.items.push({ text: row.join("  —  "), level: 0, bullet: true });
    else if (b.type === "code") for (const l of b.lines) cur.items.push({ text: "`" + l + "`", level: 0, bullet: false });
  }
  return slides.filter(s => s.title || s.items.length);
}

function slideXml(slide, isTitleSlide) {
  const sp = (id, name, ph, paras, bodyPr = "<a:bodyPr/>") =>
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${ph}</p:nvPr></p:nvSpPr><p:spPr/><p:txBody>${bodyPr}<a:lstStyle/>${paras}</p:txBody></p:sp>`;
  const titlePara = `<a:p>${aRuns(slide.title || " ")}</a:p>`;
  let shapes;
  if (isTitleSlide) {
    const sub = slide.items.map(it => `<a:p>${aRuns(it.text)}</a:p>`).join("") || "<a:p><a:endParaRPr lang=\"ru-RU\"/></a:p>";
    shapes = sp(2, "Title 1", '<p:ph type="ctrTitle"/>', titlePara) + sp(3, "Subtitle 2", '<p:ph type="subTitle" idx="1"/>', sub);
  } else {
    const chars = slide.items.reduce((n, it) => n + it.text.length, 0);
    const lines = slide.items.length;
    const scale = chars > 900 || lines > 12 ? 55000 : chars > 650 || lines > 10 ? 65000 : chars > 450 || lines > 8 ? 77500 : chars > 300 || lines > 6 ? 90000 : 0;
    const bodyPr = scale ? `<a:bodyPr><a:normAutofit fontScale="${scale}" lnSpcReduction="10000"/></a:bodyPr>` : "<a:bodyPr><a:normAutofit/></a:bodyPr>";
    const paras = slide.items.map(it => {
      const lvl = it.level ? ` lvl="${it.level}"` : "";
      const pPr = it.bullet ? (lvl ? `<a:pPr${lvl}/>` : "") : `<a:pPr marL="0" indent="0"${lvl}><a:buNone/></a:pPr>`;
      return `<a:p>${pPr}${aRuns(it.text)}</a:p>`;
    }).join("") || '<a:p><a:endParaRPr lang="ru-RU"/></a:p>';
    shapes = sp(2, "Title 1", '<p:ph type="title"/>', titlePara) + sp(3, "Content Placeholder 2", '<p:ph idx="1"/>', paras, bodyPr);
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld ${P_NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}

export function buildPptx(markdown) {
  let slides = parseSlides(prepare(markdown));
  if (!slides.length) slides = [{ title: "Презентация", items: [] }];
  slides = slides.slice(0, 60);
  // A first slide with only a title and up to two short lines becomes the title slide.
  const titleFirst = slides.length > 1 && slides[0].items.length <= 2 && slides[0].items.every(i => i.text.length < 120);
  const rels = (target, type) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/${type}" Target="${target}"/></Relationships>`;
  const sldIds = slides.map((_, k) => `<p:sldId id="${256 + k}" r:id="rId${10 + k}"/>`).join("");
  const files = [
    { name: "[Content_Types].xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/><Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/><Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>${slides.map((_, k) => `<Override PartName="/ppt/slides/slide${k + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join("")}</Types>` },
    { name: "_rels/.rels", data: rels("ppt/presentation.xml", "officeDocument") },
    { name: "ppt/presentation.xml", data: T.PRESENTATION.replace("{{SLIDES}}", `<p:sldIdLst>${sldIds}</p:sldIdLst>`) },
    { name: "ppt/_rels/presentation.xml.rels", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId3" Type="${REL}/presProps" Target="presProps.xml"/><Relationship Id="rId4" Type="${REL}/viewProps" Target="viewProps.xml"/><Relationship Id="rId5" Type="${REL}/theme" Target="theme/theme1.xml"/><Relationship Id="rId6" Type="${REL}/tableStyles" Target="tableStyles.xml"/>${slides.map((_, k) => `<Relationship Id="rId${10 + k}" Type="${REL}/slide" Target="slides/slide${k + 1}.xml"/>`).join("")}</Relationships>` },
    { name: "ppt/presProps.xml", data: T.PRES_PROPS },
    { name: "ppt/viewProps.xml", data: T.VIEW_PROPS },
    { name: "ppt/tableStyles.xml", data: T.TABLE_STYLES },
    { name: "ppt/theme/theme1.xml", data: T.THEME },
    { name: "ppt/slideMasters/slideMaster1.xml", data: T.MASTER },
    { name: "ppt/slideMasters/_rels/slideMaster1.xml.rels", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout2.xml"/><Relationship Id="rId3" Type="${REL}/theme" Target="../theme/theme1.xml"/></Relationships>` },
    { name: "ppt/slideLayouts/slideLayout1.xml", data: T.LAYOUT_TITLE },
    { name: "ppt/slideLayouts/slideLayout2.xml", data: T.LAYOUT_CONTENT },
    { name: "ppt/slideLayouts/_rels/slideLayout1.xml.rels", data: rels("../slideMasters/slideMaster1.xml", "slideMaster") },
    { name: "ppt/slideLayouts/_rels/slideLayout2.xml.rels", data: rels("../slideMasters/slideMaster1.xml", "slideMaster") },
  ];
  slides.forEach((s, k) => {
    const isTitle = k === 0 && titleFirst;
    files.push({ name: `ppt/slides/slide${k + 1}.xml`, data: slideXml(s, isTitle) });
    files.push({ name: `ppt/slides/_rels/slide${k + 1}.xml.rels`, data: rels(`../slideLayouts/slideLayout${isTitle ? 1 : 2}.xml`, "slideLayout") });
  });
  return createZip(files);
}

/* ------------------------------ entry ------------------------------ */

export function buildFile(block) {
  if (block.type === "docx") return buildDocx(block.content, { font: block.font, size: block.size, title: block.name.replace(/\.docx$/i, "") });
  if (block.type === "xlsx") return buildXlsx(block.content);
  if (block.type === "pptx") return buildPptx(block.content);
  throw new Error("unknown_file_type");
}
