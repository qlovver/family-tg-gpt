// Editing a user's own Word / Excel / PowerPoint file while keeping its formatting.
//
// 1. The model sees the file with addresses: [P12] for Word paragraphs,
//    [S2P3] for slide paragraphs, A1-style cells for Excel.
// 2. It answers with an <edit file="…"> block of operations:
//      replace P12: новый текст        after P12: новый абзац     before P3: …
//      delete P15  |  delete P15-P18    find: старое => новое      append: …
//      sheet: Лист1   set B5: 123   set C7: =SUM(C2:C6)   clear B5
// 3. The Worker applies them to the original XML and returns the same file.

import { openZip, createZipCompressed } from "./zip.js";
import { inlineRuns } from "./markdown.js";
import { decodeXml } from "./files.js";
import { convertLatex } from "./format.js";

const enc = s => String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
  .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

/* ------------------------------ paragraph engine (w: and a:) ------------------------------ */

function paraRegex(ns) {
  return new RegExp(`<${ns}:p\\b[^>]*?(?:/>|>[\\s\\S]*?</${ns}:p>)`, "g");
}

export function listParagraphs(xml, ns) {
  const out = [];
  const re = paraRegex(ns);
  let m;
  while ((m = re.exec(xml))) out.push({ start: m.index, end: m.index + m[0].length, xml: m[0], text: paraText(m[0], ns) });
  return out;
}

export function paraText(p, ns) {
  let s = "";
  const re = new RegExp(`<${ns}:t(?:\\s[^>]*)?>([^<]*)</${ns}:t>|<${ns}:tab/>|<${ns}:br/>`, "g");
  let m;
  while ((m = re.exec(p))) s += m[1] !== undefined ? decodeXml(m[1]) : m[0].includes("tab") ? "\t" : "\n";
  return s;
}

function firstRunProps(p, ns) {
  // properties of the FIRST run only (a run without rPr means plain text)
  const first = ns === "w" ? /<w:r\b(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/.exec(p) : /<a:r>([\s\S]*?)<\/a:r>/.exec(p);
  if (first) {
    const m = ns === "w"
      ? /^\s*(<w:rPr>[\s\S]*?<\/w:rPr>)/.exec(first[1])
      : /^\s*(<a:rPr\b[^>]*\/>|<a:rPr\b[^>]*>[\s\S]*?<\/a:rPr>)/.exec(first[1]);
    return m ? m[1] : "";
  }
  if (ns === "a") {
    const e = /<a:endParaRPr\b([^>]*?)(\/>|>([\s\S]*?)<\/a:endParaRPr>)/.exec(p);
    if (e) return e[3] !== undefined ? `<a:rPr${e[1]}>${e[3]}</a:rPr>` : `<a:rPr${e[1]}/>`;
  }
  return "";
}

function paraProps(p, ns) {
  const m = new RegExp(`<${ns}:pPr\\b[^>]*/>|<${ns}:pPr\\b[^>]*>[\\s\\S]*?</${ns}:pPr>`).exec(p);
  return m ? m[0] : "";
}

function styledRunProps(base, ns, run) {
  if (ns === "w") {
    let inner = base.replace(/^<w:rPr>|<\/w:rPr>$/g, "").replace(/<w:(b|bCs|i|iCs|strike)(\s[^>]*)?\/>/g, "");
    if (run.code) inner = inner.replace(/<w:rFonts\b[^>]*\/>/, "") + '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>';
    // keep schema order roughly: put b/i near the start (after rStyle/rFonts)
    const flags = (run.b ? "<w:b/><w:bCs/>" : "") + (run.i ? "<w:i/><w:iCs/>" : "") + (run.s ? "<w:strike/>" : "");
    const pos = inner.search(/<w:(?!rStyle|rFonts)[a-zA-Z]/);
    inner = pos < 0 ? inner + flags : inner.slice(0, pos) + flags + inner.slice(pos);
    return inner ? `<w:rPr>${inner}</w:rPr>` : "";
  }
  let tag = base || '<a:rPr lang="ru-RU" dirty="0"/>';
  tag = tag.replace(/^<a:rPr\b([^>]*?)(\/?)>/, (_, attrs, close) => {
    attrs = attrs.replace(/\s(b|i|strike)="[^"]*"/g, "");
    if (run.b) attrs += ' b="1"';
    if (run.i) attrs += ' i="1"';
    if (run.s) attrs += ' strike="sngStrike"';
    return `<a:rPr${attrs}${close}>`;
  });
  return tag;
}

function buildRuns(text, baseRPr, ns) {
  return inlineRuns(text).map(r => {
    const rpr = styledRunProps(baseRPr, ns, r);
    if (ns === "w") {
      const parts = r.text.split("\t").map(x => `<w:t xml:space="preserve">${enc(x)}</w:t>`).join("<w:tab/>");
      return `<w:r>${rpr}${parts}</w:r>`;
    }
    return `<a:r>${rpr}<a:t>${enc(r.text)}</a:t></a:r>`;
  }).join("");
}

function newParagraph(text, ref, ns, ctx) {
  let pPr = paraProps(ref, ns);
  let t = convertLatex(text);
  let m;
  if (ns === "w") {
    if ((m = /^(#{1,3})\s+(.*)$/.exec(t)) && ctx.headingStyles[m[1].length]) {
      pPr = `<w:pPr><w:pStyle w:val="${ctx.headingStyles[m[1].length]}"/></w:pPr>`;
      t = m[2].replace(/\*\*/g, "");
    } else if ((m = /^[-*•]\s+(.*)$/.exec(t))) {
      if (ctx.listPPr) pPr = ctx.listPPr; else t = "• " + m[1];
      if (ctx.listPPr) t = m[1];
    } else if (/<w:pStyle w:val="[^"]*(Heading|Title|Заголов)/i.test(pPr) && ctx.normalPPr !== undefined && !/^#/.test(text)) {
      pPr = ctx.normalPPr;
    }
    if ((/<w:numPr>/.test(pPr) || /<w:pStyle w:val="[^"]*List/i.test(pPr)) && !/^[-*•]\s/.test(text) && ctx.normalPPr !== undefined && !/^\d+[.)]\s/.test(text)) pPr = ctx.normalPPr;
  } else if ((m = /^[-*•]\s+(.*)$/.exec(t))) t = m[1];
  const rPr = firstRunProps(ref, ns);
  return `<${ns}:p>${pPr}${buildRuns(t, rPr, ns)}</${ns}:p>`;
}

function replaceParagraphText(p, text, ns) {
  const pPr = paraProps(p, ns);
  const rPr = firstRunProps(p, ns);
  const open = new RegExp(`^<${ns}:p\\b[^>]*?>`).exec(p)?.[0].replace(/\/>$/, ">") || `<${ns}:p>`;
  const tail = ns === "a" ? (/<a:endParaRPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:endParaRPr>)/.exec(p)?.[0] || "") : "";
  return `${open}${pPr}${buildRuns(convertLatex(text).replace(/^#{1,3}\s+/, "").replace(/^[-*•]\s+/, ""), rPr, ns)}${tail}</${ns}:p>`;
}

/** Find & replace inside a paragraph, even when the text is split across runs. */
export function findReplaceInParagraph(p, find, repl, ns) {
  if (!find) return { xml: p, count: 0 };
  const re = new RegExp(`(<${ns}:t(?:\\s[^>]*)?>)([^<]*)(</${ns}:t>)`, "g");
  const nodes = [];
  let m;
  while ((m = re.exec(p))) nodes.push({ start: m.index, end: m.index + m[0].length, open: m[1], text: decodeXml(m[2]), close: m[3] });
  if (!nodes.length) return { xml: p, count: 0 };
  const full = nodes.map(n => n.text).join("");
  let idx = full.indexOf(find);
  if (idx < 0) return { xml: p, count: 0 };
  idx = 0;
  const texts = nodes.map(n => n.text);
  const offsets = [];
  let acc = 0;
  for (const tx of texts) { offsets.push(acc); acc += tx.length; }
  let count = 0;
  let searchFrom = 0;
  // apply from the end so earlier offsets stay valid
  const hits = [];
  while ((idx = full.indexOf(find, searchFrom)) >= 0) { hits.push(idx); searchFrom = idx + find.length; }
  for (const h of hits.reverse()) {
    const endPos = h + find.length;
    let first = -1;
    for (let k = 0; k < texts.length; k++) {
      const s = offsets[k], e = offsets[k] + nodes[k].text.length;
      if (e <= h || s >= endPos) continue;
      const localStart = Math.max(0, h - s), localEnd = Math.min(nodes[k].text.length, endPos - s);
      if (first < 0) { first = k; texts[k] = texts[k].slice(0, localStart) + repl + texts[k].slice(localEnd); }
      else texts[k] = texts[k].slice(0, localStart) + texts[k].slice(localEnd);
    }
    count++;
  }
  let out = "", last = 0;
  nodes.forEach((n, k) => {
    out += p.slice(last, n.start);
    const openTag = ns === "w" && !/xml:space=/.test(n.open) ? n.open.replace(/>$/, ' xml:space="preserve">') : n.open;
    out += openTag + enc(texts[k]) + n.close;
    last = n.end;
  });
  out += p.slice(last);
  return { xml: out, count };
}

/* ------------------------------ operations ------------------------------ */

const OP_RE = /^\s*(replace|заменить|after|insert after|вставить после|before|insert before|вставить перед|delete|удалить|find|найти|append|в конец|добавить в конец|sheet|лист|set|записать|clear|очистить|delete slide|удалить слайд)\b\s*/i;

/** Parses the operations text of an <edit> block. */
export function parseOps(text) {
  const ops = [];
  let cur = null;
  for (const raw of String(text).replace(/\r\n?/g, "\n").split("\n")) {
    const line = raw.replace(/\s+$/, "");
    if (!line.trim()) { cur = null; continue; }
    let m;
    if ((m = /^\s*(?:find|найти|заменить текст)\s*:\s*(.+?)\s*(?:=>|→|->)\s*(.*)$/i.exec(line))) {
      ops.push({ op: "find", find: m[1].replace(/^["«]|["»]$/g, ""), repl: m[2].replace(/^["«]|["»]$/g, "") }); cur = null; continue;
    }
    if ((m = /^\s*(?:delete slide|удалить слайд)\s+(\d+)/i.exec(line))) { ops.push({ op: "deleteSlide", n: Number(m[1]) }); cur = null; continue; }
    if ((m = /^\s*(?:delete|удалить)\s+((?:S\d+)?P\d+)(?:\s*[-–]\s*((?:S\d+)?P\d+))?/i.exec(line))) {
      ops.push({ op: "delete", id: m[1].toUpperCase(), to: m[2]?.toUpperCase() }); cur = null; continue;
    }
    if ((m = /^\s*(replace|заменить|after|insert after|вставить после|before|insert before|вставить перед)\s+((?:S\d+)?P\d+)\s*:\s?(.*)$/i.exec(line))) {
      const kind = /^(replace|заменить)$/i.test(m[1]) ? "replace" : /before|перед/i.test(m[1]) ? "before" : "after";
      cur = { op: kind, id: m[2].toUpperCase(), lines: [m[3]] };
      ops.push(cur); continue;
    }
    if ((m = /^\s*(?:append|в конец|добавить в конец)\s*:\s?(.*)$/i.exec(line))) { cur = { op: "append", lines: [m[1]] }; ops.push(cur); continue; }
    if ((m = /^\s*(?:sheet|лист)\s*:\s*(.+)$/i.exec(line))) { ops.push({ op: "sheet", name: m[1].trim().replace(/^[«'"]|[»'"]$/g, "") }); cur = null; continue; }
    if ((m = /^\s*(?:set|записать)\s+(?:'?([^'!]+?)'?!)?\$?([A-Z]{1,3})\$?(\d{1,7})\s*[:=]\s?(.*)$/i.exec(line))) {
      ops.push({ op: "set", sheet: m[1], ref: m[2].toUpperCase() + m[3], value: m[4] }); cur = null; continue;
    }
    if ((m = /^\s*(?:clear|очистить)\s+(?:'?([^'!]+?)'?!)?([A-Z]{1,3})(\d{1,7})(?::([A-Z]{1,3})(\d{1,7}))?/i.exec(line))) {
      ops.push({ op: "clear", sheet: m[1], ref: m[2].toUpperCase() + m[3], to: m[4] ? m[4].toUpperCase() + m[5] : null }); cur = null; continue;
    }
    if (cur && cur.lines) cur.lines.push(line); // continuation = more paragraphs
  }
  for (const o of ops) if (o.lines) o.lines = o.lines.filter((l, i) => i === 0 || l.trim()).map(l => l.trim()).filter((l, i, a) => l || a.length === 1);
  return ops;
}

/** Extracts <edit file="…">…</edit> blocks from an answer. */
export function extractEditBlocks(answer) {
  const edits = [];
  const re = /(?:```[a-z]*\s*\n)?<edit\b([^>]*)>\s*\n?([\s\S]*?)(?:<\/edit>|$)(?:\s*\n```)?/gi;
  const text = String(answer).replace(re, (_, attrs, body) => {
    const file = /file\s*=\s*"([^"]*)"/i.exec(attrs)?.[1] || "";
    edits.push({ file, ops: parseOps(body) });
    return "\n";
  });
  return { text: text.replace(/\n{3,}/g, "\n\n").trim(), edits };
}

/* ------------------------------ Word ------------------------------ */

function headingStylesOf(stylesXml) {
  const map = {};
  for (const m of (stylesXml || "").matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>[\s\S]*?<w:name w:val="([^"]+)"/g)) {
    const n = /^heading (\d)$/i.exec(m[2]);
    if (n && !map[n[1]]) map[n[1]] = m[1];
  }
  return map;
}

// Applies paragraph operations; marks each op with _done / _hits.
function applyParagraphOps(xml, ops, ns, ctx, idOf) {
  const paras = listParagraphs(xml, ns);
  const byId = new Map(paras.map((p, i) => [idOf(i), i]));
  const replaced = paras.map(p => p.xml);
  const before = paras.map(() => []);
  const after = paras.map(() => []);
  const removed = new Set();
  for (const o of ops) {
    if (o.op === "find") {
      for (let i = 0; i < replaced.length; i++) {
        const r = findReplaceInParagraph(replaced[i], o.find, o.repl, ns);
        replaced[i] = r.xml; o._hits = (o._hits || 0) + r.count;
      }
      continue;
    }
    if (!["replace", "after", "before", "delete"].includes(o.op)) continue;
    const i = byId.get(o.id);
    if (i === undefined) continue;
    o._done = true;
    if (o.op === "delete") {
      const j = o.to ? byId.get(o.to) ?? i : i;
      for (let k = i; k <= j; k++) removed.add(k);
      continue;
    }
    const lines = o.lines.length ? o.lines : [""];
    if (o.op === "replace") {
      replaced[i] = replaceParagraphText(replaced[i], lines[0], ns);
      after[i].push(...lines.slice(1).map(l => newParagraph(l, paras[i].xml, ns, ctx)));
    } else if (o.op === "after") after[i].push(...lines.map(l => newParagraph(l, paras[i].xml, ns, ctx)));
    else before[i].push(...lines.map(l => newParagraph(l, paras[i].xml, ns, ctx)));
  }
  let out = "", last = 0;
  paras.forEach((p, i) => {
    out += xml.slice(last, p.start);
    out += before[i].join("") + (removed.has(i) ? "" : replaced[i]) + after[i].join("");
    last = p.end;
  });
  out += xml.slice(last);
  return out;
}

function reportOf(ops) {
  const relevant = ops.filter(o => o.op !== "sheet");
  const ok = o => o._done || o._hits;
  return {
    applied: relevant.filter(ok).length,
    missing: relevant.filter(o => !ok(o)).map(o => o.op === "find" ? `«${o.find.slice(0, 40)}»` : o.id || o.ref || (o.n ? `слайд ${o.n}` : o.op)),
  };
}

export async function docxForEdit(bytes) {
  const zip = openZip(bytes);
  const xml = await zip.text("word/document.xml");
  if (!xml) throw new Error("bad_docx");
  const styles = (await zip.text("word/styles.xml")) || "";
  const headingIds = Object.fromEntries(Object.entries(headingStylesOf(styles)).map(([l, id]) => [id, Number(l)]));
  const paras = listParagraphs(xml, "w");
  // table ranges, to show where a paragraph sits
  const tables = [...xml.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)].map((m, k) => ({ start: m.index, end: m.index + m[0].length, n: k + 1 }));
  const lines = [];
  paras.forEach((p, i) => {
    const text = p.text.replace(/\n/g, " ").trim();
    if (!text) return;
    const style = /<w:pStyle w:val="([^"]+)"/.exec(p.xml)?.[1] || "";
    const lvl = headingIds[style] || (/^Title$/i.test(style) ? 1 : 0);
    const tbl = tables.find(t => p.start > t.start && p.end < t.end);
    const mark = lvl ? "#".repeat(lvl) + " " : /<w:numPr>/.test(p.xml) || /^List/i.test(style) ? "- " : "";
    lines.push(`[P${i + 1}]${tbl ? ` (табл.${tbl.n})` : ""} ${mark}${text}`);
  });
  return lines.join("\n");
}

async function rewriteZip(bytes, changes) {
  const zip = openZip(bytes);
  const files = [];
  for (const name of zip.names) {
    if (changes.remove?.has(name)) continue;
    if (changes.files[name] !== undefined) files.push({ name, data: changes.files[name] });
    else files.push({ name, data: await zip.read(name) });
  }
  for (const [name, data] of Object.entries(changes.files)) if (!zip.names.includes(name)) files.push({ name, data });
  return createZipCompressed(files);
}

export async function applyDocxEdits(bytes, ops) {
  const zip = openZip(bytes);
  const xml = await zip.text("word/document.xml");
  const styles = (await zip.text("word/styles.xml")) || "";
  const paras = listParagraphs(xml, "w");
  const listPara = paras.find(p => (/<w:numPr>/.test(p.xml) && !/<w:ilvl w:val="[1-9]"/.test(p.xml)) || /<w:pStyle w:val="[^"]*(List|Список)[^"]*"/i.test(p.xml) && !/Number|Нумер/i.test(p.xml));
  const normalPara = paras.find(p => p.text.trim().length > 40 && !/<w:pStyle/.test(p.xml) && !/<w:numPr>/.test(p.xml))
    || paras.find(p => p.text.trim() && !/<w:pStyle w:val="[^"]*(Heading|Title)/i.test(p.xml));
  const ctx = {
    headingStyles: headingStylesOf(styles),
    listPPr: listPara ? paraProps(listPara.xml, "w") : "",
    normalPPr: normalPara ? paraProps(normalPara.xml, "w") : "",
  };
  for (const o of ops) { delete o._done; delete o._hits; }
  let out = applyParagraphOps(xml, ops, "w", ctx, i => `P${i + 1}`);
  // append: before the final sectPr
  for (const o of ops.filter(x => x.op === "append")) {
    const ref = normalPara?.xml || "<w:p/>";
    const paras2 = o.lines.map(l => newParagraph(l, ref, "w", ctx)).join("");
    const pos = out.lastIndexOf("<w:sectPr");
    const bodyEnd = out.lastIndexOf("</w:body>");
    const at = pos > 0 && pos > out.lastIndexOf("</w:p>") - 1 ? pos : bodyEnd;
    out = out.slice(0, at) + paras2 + out.slice(at);
    o._done = true;
  }
  return { bytes: await rewriteZip(bytes, { files: { "word/document.xml": out } }), report: reportOf(ops) };
}

/* ------------------------------ PowerPoint ------------------------------ */

async function slideOrder(zip) {
  const pres = (await zip.text("ppt/presentation.xml")) || "";
  const rels = (await zip.text("ppt/_rels/presentation.xml.rels")) || "";
  const map = {};
  for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = /Id="([^"]+)"/.exec(m[0])?.[1], target = /Target="([^"]+)"/.exec(m[0])?.[1];
    if (id && target) map[id] = "ppt/" + target.replace(/^\/?ppt\//, "").replace(/^\.\.\//, "");
  }
  const order = [...pres.matchAll(/<p:sldId\b[^>]*r:id="([^"]+)"/g)].map(m => map[m[1]]).filter(Boolean);
  return order.length ? order : zip.names.filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n)).sort((a, b) => parseInt(a.match(/\d+/)) - parseInt(b.match(/\d+/)));
}

export async function pptxForEdit(bytes) {
  const zip = openZip(bytes);
  const order = await slideOrder(zip);
  const out = [];
  for (let s = 0; s < order.length; s++) {
    const xml = (await zip.text(order[s])) || "";
    const lines = [`## Слайд ${s + 1}`];
    listParagraphs(xml, "a").forEach((p, i) => {
      const t = p.text.replace(/\n/g, " ").trim();
      if (t) lines.push(`[S${s + 1}P${i + 1}] ${t}`);
    });
    out.push(lines.join("\n"));
  }
  return out.join("\n\n");
}

export async function applyPptxEdits(bytes, ops) {
  const zip = openZip(bytes);
  const order = await slideOrder(zip);
  const files = {};
  for (const o of ops) { delete o._done; delete o._hits; }
  for (let s = 0; s < order.length; s++) {
    const prefix = `S${s + 1}P`;
    const mine = ops.filter(o => o.op === "find" || (o.id && o.id.startsWith(prefix) && /^S\d+P\d+$/.test(o.id) && o.id.split("P")[0] === `S${s + 1}`));
    if (!mine.length) continue;
    const xml = (await zip.text(order[s])) || "";
    const out = applyParagraphOps(xml, mine, "a", {}, i => `S${s + 1}P${i + 1}`);
    if (out !== xml) files[order[s]] = out;
  }
  const del = ops.filter(o => o.op === "deleteSlide");
  if (del.length) {
    let pres = (await zip.text("ppt/presentation.xml")) || "";
    const ids = [...pres.matchAll(/<p:sldId\b[^>]*\/>/g)].map(m => m[0]);
    for (const o of del) if (ids[o.n - 1]) { pres = pres.replace(ids[o.n - 1], ""); o._done = true; }
    files["ppt/presentation.xml"] = pres;
  }
  return { bytes: await rewriteZip(bytes, { files }), report: reportOf(ops) };
}

/* ------------------------------ Excel ------------------------------ */

const colIdx = l => [...l].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const splitRef = ref => { const m = /^([A-Z]+)(\d+)$/.exec(ref); return { col: m[1], row: Number(m[2]) }; };

function cellXml(ref, value, style) {
  const s = style ? ` s="${style}"` : "";
  const v = String(value ?? "").trim();
  if (v === "") return `<c r="${ref}"${s}/>`;
  if (v.startsWith("=") && v.length > 1) return `<c r="${ref}"${s}><f>${enc(v.slice(1))}</f></c>`;
  const n = v.replace(/[\s ]/g, "").replace(/^−/, "-");
  if (/^-?\d+(?:[.,]\d+)?$/.test(n) && !/^-?0\d/.test(n)) return `<c r="${ref}"${s}><v>${Number(n.replace(",", "."))}</v></c>`;
  if (/^-?\d+(?:[.,]\d+)?%$/.test(n)) return `<c r="${ref}"${s}><v>${Number(n.slice(0, -1).replace(",", ".")) / 100}</v></c>`;
  return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${enc(v)}</t></is></c>`;
}

export function setCell(sheetXml, ref, value) {
  const { col, row } = splitRef(ref);
  const rowRe = new RegExp(`<row\\b[^>]*\\br="${row}"[^>]*?(?:/>|>[\\s\\S]*?</row>)`);
  const cellRe = new RegExp(`<c\\b[^>]*\\br="${ref}"[^>]*?(?:/>|>[\\s\\S]*?</c>)`);
  let rowM = rowRe.exec(sheetXml);
  if (!rowM) {
    // insert a new row in order
    const rows = [...sheetXml.matchAll(/<row\b[^>]*\br="(\d+)"[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g)];
    const next = rows.find(r => Number(r[1]) > row);
    const newRow = `<row r="${row}">${cellXml(ref, value)}</row>`;
    if (next) return sheetXml.slice(0, next.index) + newRow + sheetXml.slice(next.index);
    if (/<sheetData\s*\/>/.test(sheetXml)) return sheetXml.replace(/<sheetData\s*\/>/, `<sheetData>${newRow}</sheetData>`);
    return sheetXml.replace("</sheetData>", newRow + "</sheetData>");
  }
  let rowXml = rowM[0];
  const cellM = cellRe.exec(rowXml);
  if (cellM) {
    const style = /\ss="(\d+)"/.exec(cellM[0])?.[1];
    rowXml = rowXml.replace(cellM[0], cellXml(ref, value, style));
  } else {
    if (rowXml.endsWith("/>")) rowXml = rowXml.replace(/\/>$/, "></row>");
    const cells = [...rowXml.matchAll(/<c\b[^>]*\br="([A-Z]+)\d+"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g)];
    const after = cells.find(c => colIdx(c[1]) > colIdx(col));
    // inherit the row's style from a neighbour cell for consistent look
    const neighbour = cells.length ? /\ss="(\d+)"/.exec(cells[cells.length - 1][0])?.[1] : undefined;
    const c = cellXml(ref, value, neighbour);
    rowXml = after ? rowXml.replace(after[0], c + after[0]) : rowXml.replace(/<\/row>$/, c + "</row>");
    // keep the "spans" hint from rejecting new columns
    rowXml = rowXml.replace(/\sspans="[^"]*"/, "");
  }
  return sheetXml.slice(0, rowM.index) + rowXml + sheetXml.slice(rowM.index + rowM[0].length);
}

export async function applyXlsxEdits(bytes, ops) {
  const zip = openZip(bytes);
  const wb = (await zip.text("xl/workbook.xml")) || "";
  const rels = (await zip.text("xl/_rels/workbook.xml.rels")) || "";
  const relMap = {};
  for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = /Id="([^"]+)"/.exec(m[0])?.[1], target = /Target="([^"]+)"/.exec(m[0])?.[1];
    if (id && target) relMap[id] = "xl/" + target.replace(/^\/?xl\//, "");
  }
  const sheets = [...wb.matchAll(/<sheet\b[^>]*>/g)].map(m => ({
    name: decodeXml(/name="([^"]*)"/.exec(m[0])?.[1] || ""), path: relMap[/r:id="([^"]+)"/.exec(m[0])?.[1]],
  })).filter(s => s.path);
  const files = {};
  const missingSheets = [];
  const get = async path => files[path] ?? (await zip.text(path)) ?? "";
  const find = name => sheets.find(s => s.name.toLowerCase() === String(name).toLowerCase().trim()) || null;
  let current = sheets[0];
  for (const o of ops) {
    if (o.op === "sheet") { const s = find(o.name); if (s) current = s; else missingSheets.push(`лист «${o.name}»`); continue; }
    if (o.op !== "set" && o.op !== "clear") continue;
    const sheet = o.sheet ? find(o.sheet) : current;
    if (!sheet) { missingSheets.push(`лист «${o.sheet}»`); continue; }
    let x = await get(sheet.path);
    if (o.op === "set") x = setCell(x, o.ref, o.value);
    else {
      const a = splitRef(o.ref), b = splitRef(o.to || o.ref);
      for (let r = a.row; r <= b.row; r++) for (let c = colIdx(a.col); c <= colIdx(b.col); c++) {
        let letters = ""; for (let n = c; n > 0; n = Math.floor((n - 1) / 26)) letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
        x = setCell(x, letters + r, "");
      }
    }
    files[sheet.path] = x;
    o._done = true;
  }
  // formulas must be recalculated: force recalc and drop the stale calc chain
  let wbOut = wb.replace(/<calcPr\b[^>]*\/>/, m => (/fullCalcOnLoad/.test(m) ? m : m.replace("/>", ' fullCalcOnLoad="1"/>')));
  if (!/<calcPr\b/.test(wbOut)) wbOut = wbOut.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>');
  files["xl/workbook.xml"] = wbOut;
  const remove = new Set();
  if (zip.names.includes("xl/calcChain.xml")) {
    remove.add("xl/calcChain.xml");
    files["xl/_rels/workbook.xml.rels"] = rels.replace(/<Relationship\b[^>]*calcChain[^>]*\/>/, "");
    const ct = (await zip.text("[Content_Types].xml")) || "";
    files["[Content_Types].xml"] = ct.replace(/<Override\b[^>]*calcChain[^>]*\/>/, "");
  }
  const report = reportOf(ops.filter(o => o.op === "set" || o.op === "clear"));
  report.missing.push(...missingSheets);
  return { bytes: await rewriteZip(bytes, { files, remove }), report };
}

/* ------------------------------ entry ------------------------------ */

export function editKindOf(name) {
  const ext = /\.([a-z0-9]+)$/i.exec(name || "")?.[1]?.toLowerCase();
  return ext === "docx" ? "docx" : ext === "xlsx" || ext === "xlsm" ? "xlsx" : ext === "pptx" ? "pptx" : null;
}

export async function applyEdits(kind, bytes, ops) {
  if (kind === "docx") return applyDocxEdits(bytes, ops);
  if (kind === "xlsx") return applyXlsxEdits(bytes, ops);
  if (kind === "pptx") return applyPptxEdits(bytes, ops);
  throw new Error("edit_unsupported");
}
