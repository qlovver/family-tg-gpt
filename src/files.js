// Reading files sent by users: Word, Excel, PowerPoint, PDF, text/code.
// Office formats are parsed with regular expressions (Workers have no DOMParser).

import { openZip } from "./zip.js";

export const MAX_FILE_SIZE = 20_000_000;   // Telegram Bot API download limit
export const MAX_PDF_SIZE = 10_000_000;    // sent to the model as base64
export const MAX_EXTRACT_CHARS = 120_000;  // ~40k tokens of document text

const TEXT_EXT = new Set([
  "txt", "md", "markdown", "csv", "tsv", "json", "xml", "yaml", "yml", "ini", "toml", "log",
  "html", "htm", "css", "js", "mjs", "ts", "tsx", "jsx", "py", "java", "c", "h", "cpp", "hpp",
  "cs", "go", "rs", "rb", "php", "swift", "kt", "sql", "sh", "bat", "ps1", "r", "m", "tex",
  "srt", "vtt", "rtf", "svg",
]);
const IMAGE_EXT = new Set(["jpg", "jpeg", "png", "webp", "gif"]);
const LEGACY = { doc: "docx", xls: "xlsx", ppt: "pptx", odt: "docx", ods: "xlsx", odp: "pptx", pages: "docx", numbers: "xlsx", key: "pptx" };

export function fileExt(name = "") {
  const m = /\.([A-Za-z0-9]+)$/.exec(name);
  return m ? m[1].toLowerCase() : "";
}

/** Classifies a Telegram document before downloading it. */
export function classifyDocument(doc) {
  const name = doc.file_name || "файл";
  const ext = fileExt(name);
  const mime = (doc.mime_type || "").toLowerCase();
  if (doc.file_size && doc.file_size > MAX_FILE_SIZE) return { error: "file_too_large" };
  if (IMAGE_EXT.has(ext) || ["image/jpeg", "image/png", "image/webp", "image/gif"].includes(mime)) return { kind: "image", name, mime: mime || "image/jpeg" };
  if (ext === "heic" || mime === "image/heic") return { error: "heic_unsupported" };
  if (ext === "pdf" || mime === "application/pdf") {
    if (doc.file_size && doc.file_size > MAX_PDF_SIZE) return { error: "pdf_too_large" };
    return { kind: "pdf", name };
  }
  if (ext === "docx" || ext === "xlsx" || ext === "xlsm" || ext === "pptx") return { kind: ext === "xlsm" ? "xlsx" : ext, name };
  if (ext in LEGACY) return { error: "legacy_format", hint: LEGACY[ext], name };
  if (TEXT_EXT.has(ext) || mime.startsWith("text/") || mime === "application/json") return { kind: "text", name };
  return { error: "file_unsupported", name };
}

/* ------------------------------ helpers ------------------------------ */

export function decodeXml(s) {
  return String(s)
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&amp;/g, "&");
}

function cap(text, note) {
  const t = text.replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
  if (t.length <= MAX_EXTRACT_CHARS) return { text: t, truncated: false, note };
  return { text: t.slice(0, MAX_EXTRACT_CHARS) + "\n\n[… файл обрезан: показана только первая часть …]", truncated: true, note };
}

export function decodeText(bytes) {
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  try { return new TextDecoder("utf-8", { fatal: true }).decode(u8).replace(/^﻿/, ""); } catch { /* not utf-8 */ }
  try { return new TextDecoder("windows-1251").decode(u8); } catch { /* unsupported */ }
  return new TextDecoder("utf-8").decode(u8);
}

/* ------------------------------ Word ------------------------------ */

function docxParagraph(p) {
  let out = "";
  const re = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<m:t(?:\s[^>]*)?>([^<]*)<\/m:t>|<w:tab\/>|<w:br\/>|<w:cr\/>/g;
  let m;
  while ((m = re.exec(p))) {
    if (m[1] !== undefined) out += decodeXml(m[1]);
    else if (m[2] !== undefined) out += decodeXml(m[2]);
    else if (m[0].startsWith("<w:tab")) out += "\t";
    else out += "\n";
  }
  const style = /<w:pStyle w:val="([^"]+)"/.exec(p)?.[1] || "";
  const h = /^(?:heading|заголовок)\s*(\d)$/i.exec(style.replace(/[^A-Za-zА-Яа-я0-9]/g, " ").trim()) || /Heading(\d)/i.exec(style);
  if (out.trim() && h) return "#".repeat(Math.min(Number(h[1]), 4)) + " " + out.trim();
  if (out.trim() && /^Title$/i.test(style)) return "# " + out.trim();
  if (out.trim() && (/<w:numPr>/.test(p) || /^List/i.test(style))) {
    const lvl = Number(/<w:ilvl w:val="(\d+)"/.exec(p)?.[1] || 0);
    return "  ".repeat(lvl) + "- " + out.trim();
  }
  return out;
}

function docxTable(tbl) {
  const rows = [];
  for (const tr of tbl.match(/<w:tr[\s>][\s\S]*?<\/w:tr>/g) || []) {
    const cells = (tr.match(/<w:tc[\s>][\s\S]*?<\/w:tc>/g) || []).map(tc =>
      (tc.match(/<w:p[\s>][\s\S]*?<\/w:p>|<w:p\/>/g) || []).map(docxParagraph).join(" ").replace(/\s+/g, " ").trim().replace(/\|/g, "/"));
    rows.push("| " + cells.join(" | ") + " |");
  }
  return rows.join("\n");
}

export async function extractDocx(buffer) {
  const zip = openZip(buffer);
  const xml = await zip.text("word/document.xml");
  if (!xml) throw new Error("bad_docx");
  const body = /<w:body>([\s\S]*)<\/w:body>/.exec(xml)?.[1] || xml;
  const blocks = [];
  const re = /<w:tbl>[\s\S]*?<\/w:tbl>|<w:p[\s>][\s\S]*?<\/w:p>/g;
  let m;
  while ((m = re.exec(body))) blocks.push(m[0].startsWith("<w:tbl") ? docxTable(m[0]) : docxParagraph(m[0]));
  // footnotes / endnotes, if any
  for (const part of ["word/footnotes.xml", "word/endnotes.xml"]) {
    const x = await zip.text(part);
    if (!x) continue;
    const notes = (x.match(/<w:p[\s>][\s\S]*?<\/w:p>/g) || []).map(docxParagraph).filter(t => t.trim());
    if (notes.length) blocks.push("\nСноски:\n" + notes.join("\n"));
  }
  const hasImages = zip.names.some(n => n.startsWith("word/media/"));
  return cap(blocks.join("\n"), hasImages ? "В документе есть картинки — их содержимое не передано, только текст." : "");
}

/* ------------------------------ Excel ------------------------------ */

function colIndex(ref) {
  const letters = /^[A-Z]+/.exec(ref)?.[0] || "A";
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function colName(i) {
  let s = "";
  for (i++; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + ((i - 1) % 26)) + s;
  return s;
}

export async function extractXlsx(buffer, { maxRows = 1000, maxCols = 40 } = {}) {
  const zip = openZip(buffer);
  const wb = await zip.text("xl/workbook.xml");
  if (!wb) throw new Error("bad_xlsx");
  const rels = (await zip.text("xl/_rels/workbook.xml.rels")) || "";
  const relMap = {};
  for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = /Id="([^"]+)"/.exec(m[0])?.[1], target = /Target="([^"]+)"/.exec(m[0])?.[1];
    if (id && target) relMap[id] = target.replace(/^\/?xl\//, "").replace(/^\//, "");
  }
  const shared = [];
  const ss = await zip.text("xl/sharedStrings.xml");
  if (ss) for (const si of ss.match(/<si>[\s\S]*?<\/si>/g) || []) {
    shared.push([...si.matchAll(/<t(?:\s[^>]*)?>([^<]*)<\/t>/g)].map(x => decodeXml(x[1])).join(""));
  }
  const out = [];
  let truncated = false;
  for (const s of wb.matchAll(/<sheet\b[^>]*>/g)) {
    const name = decodeXml(/name="([^"]*)"/.exec(s[0])?.[1] || "Лист");
    const rid = /r:id="([^"]+)"/.exec(s[0])?.[1];
    const path = "xl/" + (relMap[rid] || "");
    const xml = await zip.text(path);
    if (!xml) continue;
    const lines = [`## Лист «${name}»`];
    let rowsSeen = 0;
    for (const row of xml.match(/<row\b[^>]*>[\s\S]*?<\/row>|<row\b[^>]*\/>/g) || []) {
      const rnum = /\br="(\d+)"/.exec(row)?.[1] || String(rowsSeen + 1);
      const cells = [];
      for (const c of row.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
        const attrs = c[1], inner = c[2] || "";
        const ref = /r="([A-Z]+)\d+"/.exec(attrs)?.[1] || colName(cells.length);
        const t = /t="([^"]+)"/.exec(attrs)?.[1];
        let v = /<v>([^<]*)<\/v>/.exec(inner)?.[1];
        const f = /<f[^>]*>([^<]*)<\/f>/.exec(inner)?.[1];
        if (t === "s" && v !== undefined) v = shared[Number(v)] ?? "";
        else if (t === "inlineStr") v = [...inner.matchAll(/<t(?:\s[^>]*)?>([^<]*)<\/t>/g)].map(x => x[1]).join("");
        else if (t === "b") v = v === "1" ? "ИСТИНА" : "ЛОЖЬ";
        v = v === undefined ? "" : decodeXml(v);
        const value = f ? `=${decodeXml(f)}${v !== "" ? ` → ${v}` : ""}` : v;
        if (value === "") continue;
        const idx = colIndex(ref);
        if (idx < maxCols) cells.push(`${ref}${rnum}: ${value.replace(/\n/g, " ")}`);
      }
      if (!cells.length) continue;
      if (++rowsSeen > maxRows) { truncated = true; break; }
      lines.push(cells.join(" | "));
    }
    if (lines.length === 1) lines.push("(пустой лист)");
    out.push(lines.join("\n"));
  }
  const r = cap(out.join("\n\n"), "Ячейки даны как «адрес: значение»; формулы — как «=формула → результат». Даты могут быть показаны числами Excel.");
  if (truncated) r.truncated = true;
  return r;
}

/* ------------------------------ PowerPoint ------------------------------ */

function pptxText(xml) {
  return (xml.match(/<a:p>[\s\S]*?<\/a:p>|<a:p\s[\s\S]*?<\/a:p>/g) || [])
    .map(p => [...p.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map(x => decodeXml(x[1])).join(""))
    .filter(t => t.trim()).join("\n");
}

export async function extractPptx(buffer) {
  const zip = openZip(buffer);
  const slides = zip.names.filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => Number(/(\d+)\.xml$/.exec(a)[1]) - Number(/(\d+)\.xml$/.exec(b)[1]));
  if (!slides.length) throw new Error("bad_pptx");
  const out = [];
  for (let k = 0; k < slides.length; k++) {
    const num = /(\d+)\.xml$/.exec(slides[k])[1];
    const text = pptxText((await zip.text(slides[k])) || "");
    const notes = await zip.text(`ppt/notesSlides/notesSlide${num}.xml`);
    const notesText = notes ? pptxText(notes).replace(/^\d+$/m, "").trim() : "";
    out.push(`## Слайд ${k + 1}\n${text || "(без текста)"}${notesText ? `\nЗаметки докладчика: ${notesText}` : ""}`);
  }
  const hasImages = zip.names.some(n => n.startsWith("ppt/media/"));
  return cap(out.join("\n\n"), hasImages ? "На слайдах есть картинки — передан только текст." : "");
}

/* ------------------------------ entry ------------------------------ */

function stripRtf(s) {
  return s.replace(/\\par[d]?/g, "\n").replace(/\{\\\*[^{}]*\}/g, "").replace(/\\'([0-9a-f]{2})/gi, (_, h) => String.fromCharCode(parseInt(h, 16)))
    .replace(/\\[a-z]+-?\d* ?/gi, "").replace(/[{}]/g, "");
}

/**
 * bytes + classification -> { text, note, truncated } for text-like files.
 */
export async function extractDocument(kind, name, bytes) {
  try {
    if (kind === "docx") return await extractDocx(bytes);
    if (kind === "xlsx") return await extractXlsx(bytes);
    if (kind === "pptx") return await extractPptx(bytes);
    if (kind === "text") {
      let t = decodeText(bytes);
      if (fileExt(name) === "rtf") t = stripRtf(t);
      return cap(t, "");
    }
  } catch (err) {
    if (err?.message === "zip_entry_too_large") throw new Error("file_too_large");
    throw new Error("file_parse_failed");
  }
  throw new Error("file_unsupported");
}

export function base64FromBytes(bytes) {
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  // Native in the Workers runtime: much cheaper on CPU than btoa().
  if (typeof u8.toBase64 === "function") return u8.toBase64();
  let bin = "";
  for (let i = 0; i < u8.length; i += 0x8000) bin += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000));
  return btoa(bin);
}
