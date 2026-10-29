// Designed Word documents from Markdown: colour themes or the Russian
// academic "ГОСТ" style, title block, styled headings and tables,
// answer/callout boxes, centred formulas, page numbers, contents.
//
// <file type="docx" name="Реферат.docx" style="gost">  — учебная работа
// <file type="docx" name="Решение.docx" theme="ocean"> — цветной документ
// Directives inside: "@toc" (содержание), "@pagebreak" (новая страница).

import { createZip } from "./zip.js";
import { xml, prepare, inlineRuns, parseBlocks } from "./markdown.js";
import { resolveTheme, tint, luminance } from "./theme.js";

export { tint };

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function makeStyle(opts) {
  const o = Object.fromEntries(Object.entries(opts || {}).map(([k, v]) => [String(k).toLowerCase().replace(/[-_]/g, ""), v]));
  const gost = /^(gost|гост)$/i.test(String(o.style || ""));
  let t = resolveTheme(o);
  if (t.darkContent) t = resolveTheme({ ...o, theme: "ocean", bg: undefined, background: undefined }); // no dark paper
  const font = String(o.font || (gost ? "Times New Roman" : t.headFont === "Georgia" ? "Georgia" : "Calibri")).slice(0, 40);
  const headFont = String(o.headfont || o.headingfont || (gost ? font : t.headFont !== "Calibri" ? t.headFont : font)).slice(0, 40);
  const size = Math.min(Math.max(Number(o.size) || (gost ? 14 : 11), 8), 28);
  const spacing = Math.min(Math.max(Number(String(o.spacing ?? o.linespacing ?? "").replace(",", ".")) || (gost ? 1.5 : 1.2), 1), 3);
  const alignMap = { justify: "both", both: "both", ширине: "both", "по ширине": "both", left: "left", влево: "left", "по левому краю": "left", center: "center", центру: "center", "по центру": "center", right: "right" };
  const align = alignMap[String(o.align ?? "").toLowerCase()] || (gost ? "both" : "left");
  const indentCm = Number(String(o.indent ?? "").replace(",", "."));
  const indent = Math.round((Number.isFinite(indentCm) && String(o.indent ?? "") !== "" ? indentCm : gost ? 1.25 : 0) * 567);
  const base = { gost, font, headFont, size, spacing, align, indent };
  if (gost) {
    const hc = o.heading || o.headingcolor || o.color || o.primary ? (t.heading) : "000000";
    return { ...base, heading: hc, sub: hc, accent: o.accent ? t.accent : "000000", muted: "000000",
      line: "000000", head: o.primary || o.color ? tint(t.primary, 0.15) : "FFFFFF", headText: "000000", band: "FFFFFF", box: "FFFFFF", boxLine: "000000", quote: "FFFFFF", text: o.text ? t.text : "000000" };
  }
  return { ...base,
    heading: t.heading, sub: t.primary, accent: t.accent, muted: t.muted, line: tint(t.primary, 0.25), text: t.text === "1F1F1F" || !o.text ? "000000" : t.text,
    head: t.primary, headText: luminance(t.primary) > 0.6 ? "111111" : "FFFFFF", band: tint(t.primary, 0.06), box: tint(t.accent, 0.14), boxLine: t.accent, quote: tint(t.primary, 0.05) };
}

/* ------------------------------ runs & paragraphs ------------------------------ */

function wRuns(text, extra = {}) {
  return inlineRuns(text, extra).map(r => {
    const pr = [
      r.font ? `<w:rFonts w:ascii="${xml(r.font)}" w:hAnsi="${xml(r.font)}" w:cs="${xml(r.font)}"/>` : "",
      r.code ? '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>' : "",
      r.b ? "<w:b/><w:bCs/>" : "", r.i ? "<w:i/><w:iCs/>" : "", r.s ? "<w:strike/>" : "",
      r.color ? `<w:color w:val="${r.color}"/>` : "",
      r.sz ? `<w:sz w:val="${r.sz}"/><w:szCs w:val="${r.sz}"/>` : "",
    ].join("");
    const parts = r.text.split("\t").map(x => `<w:t xml:space="preserve">${xml(x)}</w:t>`).join("<w:tab/>");
    return `<w:r>${pr ? `<w:rPr>${pr}</w:rPr>` : ""}${parts}</w:r>`;
  }).join("");
}

function wPara(runs, { style, numId, ilvl = 0, align, pPr = "", keepNext } = {}) {
  const ppr = [
    style ? `<w:pStyle w:val="${style}"/>` : "",
    keepNext ? "<w:keepNext/>" : "",
    numId ? `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>` : "",
    pPr,
    align ? `<w:jc w:val="${align}"/>` : "",
  ].join("");
  return `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ""}${runs}</w:p>`;
}

const PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';

// "k̇ = s√k − (n + g + δ)k" — a line that is a formula rather than prose
export function isFormulaLine(text) {
  const s = text.replace(/\*\*/g, "").trim();
  if (s.length < 3 || s.length > 140) return false;
  if (!/[=≈≤≥<>]/.test(s)) return false;
  const words = s.match(/[А-Яа-яЁё]{4,}/g) || [];
  if (words.length > 1) return false;
  return /[√∑∫²³⁰-⁹₀-₉πΔδαβγλμσωθ·×÷−±∞→⇒]|[a-zA-Z]\s*[=+\-*/^]|\d\s*[+\-*/×·]\s*\d/.test(s);
}

function wTable(rows, st) {
  const cols = Math.max(...rows.map(r => r.length));
  const width = st.gost ? 9355 : 9638;
  const lens = Array.from({ length: cols }, (_, ci) => Math.max(3, ...rows.map(r => String(r[ci] ?? "").length)));
  const sum = lens.reduce((a, b) => a + b, 0);
  const colW = lens.map(l => Math.max(900, Math.round(width * l / sum)));
  const k = width / colW.reduce((a, b) => a + b, 0);
  for (let i = 0; i < colW.length; i++) colW[i] = Math.round(colW[i] * k);
  const grid = colW.map(w => `<w:gridCol w:w="${w}"/>`).join("");
  const numeric = v => /^[-−+]?[\d\s ]+([.,]\d+)?\s*(%|₽|руб\.?)?$/.test(String(v).trim());
  const trs = rows.map((r, ri) => {
    const head = ri === 0 && rows.length > 1;
    const fill = head ? st.head : ri % 2 === 0 ? st.band : "FFFFFF";
    const tcs = Array.from({ length: cols }, (_, ci) => {
      const v = r[ci] ?? "";
      const shade = fill !== "FFFFFF" ? `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>` : "";
      const runs = wRuns(v, head ? { b: true, color: st.headText } : {});
      return `<w:tc><w:tcPr><w:tcW w:w="${colW[ci]}" w:type="dxa"/>${shade}<w:vAlign w:val="center"/></w:tcPr>${wPara(runs, { style: "TableText", align: !head && ci && numeric(v) ? "right" : head ? "center" : undefined })}</w:tc>`;
    }).join("");
    return `<w:tr>${head ? '<w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>' : '<w:trPr><w:cantSplit/></w:trPr>'}${tcs}</w:tr>`;
  }).join("");
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="${width}" w:type="dxa"/><w:jc w:val="center"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${trs}</w:tbl>${wPara("", { style: "AfterTable" })}`;
}

/* ------------------------------ styles ------------------------------ */

function wStyles(st) {
  const sz = Math.round(st.size * 2);
  const f = st.font, hf = st.headFont;
  const fonts = name => `<w:rFonts w:ascii="${xml(name)}" w:hAnsi="${xml(name)}" w:eastAsia="${xml(name)}" w:cs="${xml(name)}"/>`;
  const body = `<w:spacing w:after="${st.gost ? 0 : 140}" w:line="${Math.round(st.spacing * 240)}" w:lineRule="auto"/>${st.indent ? `<w:ind w:firstLine="${st.indent}"/>` : ""}<w:jc w:val="${st.align}"/>`;
  const heading = (id, name, pt, before, after, color, extra = "") =>
    `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${before}" w:after="${after}" w:line="${st.gost ? 360 : 264}" w:lineRule="auto"/><w:ind w:firstLine="0"/>${extra}${/w:jc/.test(extra) ? "" : '<w:jc w:val="left"/>'}<w:outlineLvl w:val="${Number(id.slice(-1)) - 1}"/></w:pPr><w:rPr>${fonts(hf)}<w:b/><w:bCs/><w:color w:val="${color}"/><w:sz w:val="${Math.round(pt * 2)}"/><w:szCs w:val="${Math.round(pt * 2)}"/></w:rPr></w:style>`;
  const h1 = st.gost
    ? heading("Heading1", "heading 1", st.size, 240, 240, "000000", '<w:jc w:val="center"/>')
    : heading("Heading1", "heading 1", st.size * 1.5, 420, 160, st.heading, `<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="4" w:color="${st.line}"/></w:pBdr>`);
  const h2 = st.gost
    ? heading("Heading2", "heading 2", st.size, 240, 120, "000000", '<w:jc w:val="left"/>')
    : heading("Heading2", "heading 2", st.size * 1.22, 300, 100, st.sub);
  const h3 = st.gost
    ? heading("Heading3", "heading 3", st.size, 200, 100, "000000")
    : heading("Heading3", "heading 3", st.size * 1.08, 220, 80, st.heading);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W_NS}><w:docDefaults><w:rPrDefault><w:rPr>${fonts(f)}${st.text && st.text !== "000000" ? `<w:color w:val="${st.text}"/>` : ""}<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/><w:lang w:val="ru-RU" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>${body}</w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:widowControl/></w:pPr></w:style>
<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Subtitle"/><w:qFormat/><w:pPr><w:spacing w:before="${st.gost ? 0 : 240}" w:after="${st.gost ? 240 : 80}" w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="${st.gost ? "center" : "left"}"/></w:pPr><w:rPr>${fonts(hf)}<w:b/><w:bCs/><w:color w:val="${st.heading}"/><w:sz w:val="${Math.round(st.size * (st.gost ? 1.15 : 2.4) * 2)}"/><w:szCs w:val="${Math.round(st.size * (st.gost ? 1.15 : 2.4) * 2)}"/>${st.gost ? "<w:caps/>" : ""}</w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="60"/><w:ind w:firstLine="0"/><w:jc w:val="${st.gost ? "center" : "left"}"/></w:pPr><w:rPr><w:color w:val="${st.muted}"/><w:sz w:val="${Math.round(st.size * (st.gost ? 1 : 1.15) * 2)}"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="TitleRule"><w:name w:val="Title Rule"/><w:basedOn w:val="Normal"/><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="24" w:space="1" w:color="${st.accent}"/></w:pBdr><w:spacing w:before="0" w:after="360" w:line="120" w:lineRule="exact"/><w:ind w:right="${st.gost ? 0 : 7200}" w:firstLine="0"/></w:pPr><w:rPr><w:sz w:val="4"/></w:rPr></w:style>
${h1}${h2}${h3}
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="60"/><w:ind w:left="720" w:firstLine="0"/><w:contextualSpacing/>${st.gost ? '<w:jc w:val="both"/>' : ""}</w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Formula"><w:name w:val="Formula"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="80" w:after="80"/><w:ind w:firstLine="0"/><w:jc w:val="center"/></w:pPr><w:rPr><w:sz w:val="${sz + 2}"/><w:szCs w:val="${sz + 2}"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Answer"><w:name w:val="Answer"/><w:basedOn w:val="Normal"/><w:pPr><w:pBdr><w:top w:val="single" w:sz="8" w:space="6" w:color="${st.boxLine}"/><w:left w:val="single" w:sz="8" w:space="6" w:color="${st.boxLine}"/><w:bottom w:val="single" w:sz="8" w:space="6" w:color="${st.boxLine}"/><w:right w:val="single" w:sz="8" w:space="6" w:color="${st.boxLine}"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="${st.box}"/><w:spacing w:before="200" w:after="200"/><w:ind w:left="144" w:right="144" w:firstLine="0"/></w:pPr><w:rPr><w:b/><w:bCs/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:pPr><w:pBdr><w:left w:val="single" w:sz="24" w:space="10" w:color="${st.gost ? "808080" : st.accent}"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="${st.quote}"/><w:spacing w:before="120" w:after="120"/><w:ind w:left="284" w:right="284" w:firstLine="0"/></w:pPr><w:rPr><w:i/><w:iCs/>${st.gost ? "" : `<w:color w:val="404040"/>`}</w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F3F4F6"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="left"/></w:pPr><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:sz w:val="${Math.max(16, sz - 4)}"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="TableText"><w:name w:val="Table Text"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="50" w:after="50" w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="left"/></w:pPr><w:rPr><w:sz w:val="${st.gost ? sz - 4 : sz - 1}"/><w:szCs w:val="${st.gost ? sz - 4 : sz - 1}"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="AfterTable"><w:name w:val="After Table"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="120" w:line="240" w:lineRule="auto"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="60"/><w:ind w:firstLine="0"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="TOC2"><w:name w:val="toc 2"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="40"/><w:ind w:left="440" w:firstLine="0"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:firstLine="0"/><w:jc w:val="${st.gost ? "center" : "right"}"/><w:spacing w:after="0"/></w:pPr><w:rPr><w:color w:val="${st.gost ? "000000" : "8A8A8A"}"/><w:sz w:val="${st.gost ? sz : 18}"/></w:rPr></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="${st.line}"/><w:left w:val="single" w:sz="4" w:space="0" w:color="${st.line}"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="${st.line}"/><w:right w:val="single" w:sz="4" w:space="0" w:color="${st.line}"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="${st.line}"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="${st.line}"/></w:tblBorders></w:tblPr></w:style>
</w:styles>`;
}

function wNumbering(numbered, st) {
  const lvls = fmt => Array.from({ length: 5 }, (_, l) => {
    const bullet = st.gost ? ["–", "–", "–", "–", "–"][l] : ["●", "○", "■", "●", "○"][l];
    const text = fmt === "bullet" ? bullet : `%${l + 1}.`;
    const rpr = fmt === "bullet"
      ? `<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:color w:val="${st.gost ? "000000" : st.accent}"/><w:sz w:val="${st.gost ? Math.round(st.size * 2) : 16}"/></w:rPr>`
      : `<w:rPr><w:b/><w:color w:val="${st.gost ? "000000" : st.heading}"/></w:rPr>`;
    return `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="${fmt === "bullet" ? "bullet" : "decimal"}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${(st.gost ? 1069 : 720) + l * 360}" w:hanging="360"/></w:pPr>${rpr}</w:lvl>`;
  }).join("");
  const nums = [`<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`];
  numbered.forEach((start, k) => nums.push(
    `<w:num w:numId="${k + 2}"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="${start}"/></w:lvlOverride></w:num>`));
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering ${W_NS}><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${lvls("bullet")}</w:abstractNum><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${lvls("decimal")}</w:abstractNum>${nums.join("")}</w:numbering>`;
}

function wFooter(st) {
  const field = `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr ${W_NS}><w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr>${field}</w:p></w:ftr>`;
}

/* ------------------------------ document ------------------------------ */

export function buildDocx(markdown, opts = {}) {
  const st = makeStyle(opts);
  const src = prepare(markdown);
  const blocks = parseBlocks(src);
  const headings = blocks.filter(b => b.type === "heading" && b.level >= 1 && b.level <= 3);
  const body = [];
  const numbered = [];
  let currentNum = 0, prevType = "";
  let index = 0, titleDone = false;
  // first "# " heading becomes the title; plain lines right after it — subtitle
  const titleBlock = blocks.findIndex(b => b.type !== "blank");
  const hasTitle = titleBlock >= 0 && blocks[titleBlock].type === "heading" && blocks[titleBlock].level === 1;

  for (let i = 0; i < blocks.length; i++) {
    const b = blocks[i];
    if (b.type === "para" && /^\s*@toc\s*$/i.test(b.text)) {
      body.push(wPara(wRuns("Содержание"), { style: "Heading1", align: st.gost ? "center" : undefined }));
      for (const h of headings.slice(hasTitle ? 1 : 0)) {
        const lvl = hasTitle ? Math.max(h.level - 1, 1) : h.level;
        if (lvl > 2) continue;
        body.push(wPara(wRuns(h.text.replace(/\*\*/g, "")), { style: lvl === 1 ? "TOC1" : "TOC2" }));
      }
      body.push(PAGE_BREAK);
      prevType = "toc";
      continue;
    }
    if (b.type === "para" && /^\s*(@pagebreak|\\newpage|\\pagebreak)\s*$/i.test(b.text)) { body.push(PAGE_BREAK); continue; }
    if (b.type !== "number" && b.type !== "blank") currentNum = 0;

    if (hasTitle && i === titleBlock) {
      body.push(wPara(wRuns(b.text.replace(/\*\*/g, "")), { style: "Title" }));
      // subtitle lines
      let j = i + 1;
      while (j < blocks.length && blocks[j].type === "para" && j - i <= 3 && !/^\s*(@|\\)/.test(blocks[j].text)) {
        body.push(wPara(wRuns(blocks[j].text), { style: "Subtitle" }));
        j++;
      }
      if (!st.gost) body.push(wPara("", { style: "TitleRule" }));
      else body.push(wPara("", { style: "AfterTable" }));
      i = j - 1;
      titleDone = true;
      continue;
    }
    switch (b.type) {
      case "heading": {
        const lvl = Math.min(hasTitle ? Math.max(b.level - 1, 1) : b.level, 3);
        const text = b.text.replace(/\*\*/g, "");
        body.push(wPara(wRuns(st.gost && lvl === 1 ? text.toUpperCase() : text), { style: `Heading${lvl}` }));
        break;
      }
      case "para": {
        const plain = b.text.replace(/\*\*/g, "").trim();
        if (/^(ответ|итог|вывод|результат|answer)\s*[:：]/i.test(plain)) {
          body.push(wPara(wRuns(b.text.replace(/^\*\*([^*]+)\*\*/, "$1")), { style: "Answer" }));
        } else if (isFormulaLine(b.text)) {
          body.push(wPara(wRuns(b.text), { style: "Formula" }));
        } else body.push(wPara(wRuns(b.text)));
        break;
      }
      case "bullet": body.push(wPara(wRuns(b.text), { style: "ListParagraph", numId: 1, ilvl: b.level })); break;
      case "number": {
        if (!currentNum) { numbered.push(b.n || 1); currentNum = numbered.length + 1; }
        body.push(wPara(wRuns(b.text), { style: "ListParagraph", numId: currentNum, ilvl: b.level }));
        break;
      }
      case "quote": body.push(wPara(wRuns(b.text), { style: "Quote" })); break;
      case "code":
        for (const l of b.lines.length ? b.lines : [""]) body.push(wPara(`<w:r><w:t xml:space="preserve">${xml(l)}</w:t></w:r>`, { style: "Code" }));
        body.push(wPara("", { style: "AfterTable" }));
        break;
      case "table": if (b.rows.length) body.push(wTable(b.rows, st)); break;
      case "rule": body.push(wPara("", { pPr: `<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="${st.line}"/></w:pBdr>` })); break;
      case "blank": break;
    }
    if (b.type !== "blank") { prevType = b.type; index++; }
  }
  if (!body.length) body.push(wPara(""));
  // GOST margins: left 30, right 15, top/bottom 20 mm; otherwise a balanced layout
  const margins = st.gost
    ? '<w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="709" w:footer="567" w:gutter="0"/>'
    : '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="709" w:footer="567" w:gutter="0"/>';
  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W_NS}><w:body>${body.join("")}<w:sectPr><w:footerReference w:type="default" r:id="rId3"/><w:pgSz w:w="11906" w:h="16838"/>${margins}${st.gost ? '<w:titlePg/>' : ""}</w:sectPr></w:body></w:document>`;
  const now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
  return createZip([
    { name: "[Content_Types].xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>` },
    { name: "_rels/.rels", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>` },
    { name: "docProps/core.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${xml(opts.title || "")}</dc:title><dc:creator>Telegram AI bot</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created></cp:coreProperties>` },
    { name: "word/document.xml", data: document },
    { name: "word/styles.xml", data: wStyles(st) },
    { name: "word/numbering.xml", data: wNumbering(numbered, st) },
    { name: "word/footer1.xml", data: wFooter(st) },
    { name: "word/_rels/document.xml.rels", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${REL}/numbering" Target="numbering.xml"/><Relationship Id="rId3" Type="${REL}/footer" Target="footer1.xml"/></Relationships>` },
  ]);
}
