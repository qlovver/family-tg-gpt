// Designed PowerPoint decks (16:9) drawn from shapes: themes + slide layouts.
// The model writes a light Markdown "deck" format; this module does the design:
//
// <file type="pptx" name="Доклад.pptx" theme="ocean">
// # Заголовок доклада            <- first slide: title
// Подзаголовок, автор
// ---
// @layout: section               <- optional; inferred when omitted
// # Часть 1. Теория
// ---
// # Обычный слайд
// - пункт **с акцентом**
// ---
// @layout: cards                 <- 2–6 blocks "## Заголовок" + text
// ...
// </file>
//
// Layouts: title, section, bullets, two-columns, cards, steps, big-number,
// quote, table, ending.

import { createZip } from "./zip.js";
import { xml, prepare, inlineRuns, isTableRow, isSeparator, splitRow } from "./markdown.js";
import * as T from "./pptx-template.js";

/* ------------------------------ themes ------------------------------ */

export const THEMES = {
  ocean:    { dark: "0B2545", primary: "13315C", accent: "2EC4B6", accent2: "8DA9C4", bg: "F6F9FC", card: "FFFFFF", text: "1B2A41", heading: "13315C", muted: "5C6B7A", onDark: "FFFFFF", onDarkMuted: "B9CCE0", line: "D9E2EC", font: "Calibri", headFont: "Calibri" },
  sunset:   { dark: "3A1F33", primary: "B33F62", accent: "F9A03F", accent2: "F4D35E", bg: "FFF8F1", card: "FFFFFF", text: "2D1E2F", heading: "8E2C4E", muted: "7A6670", onDark: "FFFFFF", onDarkMuted: "F2C9B8", line: "F0DCCF", font: "Calibri", headFont: "Calibri" },
  forest:   { dark: "17332A", primary: "2F6B4F", accent: "D4A72C", accent2: "8FBC94", bg: "F5F8F2", card: "FFFFFF", text: "1E2B24", heading: "245640", muted: "5F6F66", onDark: "FFFFFF", onDarkMuted: "C2D8C9", line: "DCE6DA", font: "Calibri", headFont: "Calibri" },
  minimal:  { dark: "111111", primary: "1F1F1F", accent: "E63946", accent2: "A8A8A8", bg: "FFFFFF", card: "F5F5F5", text: "1F1F1F", heading: "111111", muted: "6B6B6B", onDark: "FFFFFF", onDarkMuted: "C9C9C9", line: "E5E5E5", font: "Calibri", headFont: "Calibri" },
  lavender: { dark: "2E1A47", primary: "5A3E85", accent: "F28482", accent2: "B8A1D9", bg: "F8F5FC", card: "FFFFFF", text: "2A2140", heading: "4A2F75", muted: "6E6483", onDark: "FFFFFF", onDarkMuted: "D3C4EA", line: "E6DDF2", font: "Calibri", headFont: "Calibri" },
  night:    { dark: "0B1120", primary: "1E293B", accent: "38BDF8", accent2: "A78BFA", bg: "0F172A", card: "1E293B", text: "E2E8F0", heading: "F8FAFC", muted: "94A3B8", onDark: "F8FAFC", onDarkMuted: "94A3B8", line: "334155", font: "Calibri", headFont: "Calibri", darkContent: true },
  classic:  { dark: "1C2541", primary: "1C2541", accent: "C9A227", accent2: "8C9AB8", bg: "FBF9F4", card: "FFFFFF", text: "1C2541", heading: "1C2541", muted: "6A6F7E", onDark: "FFFFFF", onDarkMuted: "D8CFAF", line: "E7E1D2", font: "Calibri", headFont: "Georgia" },
};
const THEME_ALIASES = {
  blue: "ocean", синяя: "ocean", синий: "ocean", океан: "ocean", business: "ocean", деловая: "ocean",
  warm: "sunset", тёплая: "sunset", теплая: "sunset", закат: "sunset", red: "sunset", orange: "sunset",
  green: "forest", зелёная: "forest", зеленая: "forest", природа: "forest", eco: "forest",
  white: "minimal", минимализм: "minimal", строгая: "minimal", black: "minimal",
  purple: "lavender", фиолетовая: "lavender", лавандовая: "lavender", pink: "lavender", нежная: "lavender",
  dark: "night", тёмная: "night", темная: "night", ночь: "night", tech: "night",
  gold: "classic", классика: "classic", классическая: "classic", history: "classic", история: "classic",
};

export function pickTheme(name) {
  const n = String(name || "").toLowerCase().trim();
  return THEMES[n] ? n : THEME_ALIASES[n] || "ocean";
}

/* ------------------------------ geometry & XML ------------------------------ */

const W = 13.333, H = 7.5;            // inches, 16:9
const E = v => Math.round(v * 914400); // inches -> EMU
const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

class Canvas {
  constructor(theme) { this.t = theme; this.id = 2; this.parts = []; }
  fill(color, alpha) {
    return color ? `<a:solidFill><a:srgbClr val="${color}">${alpha != null ? `<a:alpha val="${Math.round(alpha * 1000)}"/>` : ""}</a:srgbClr></a:solidFill>` : "<a:noFill/>";
  }
  shape(x, y, w, h, { geom = "rect", fill, alpha, line, lineW = 1, adj, text = "", name = "Shape" } = {}) {
    const id = this.id++;
    const av = adj != null ? `<a:avLst><a:gd name="adj" fmla="val ${adj}"/></a:avLst>` : "<a:avLst/>";
    const ln = line ? `<a:ln w="${Math.round(lineW * 12700)}">${this.fill(line)}</a:ln>` : "<a:ln><a:noFill/></a:ln>";
    this.parts.push(`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name} ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${E(x)}" y="${E(y)}"/><a:ext cx="${E(w)}" cy="${E(h)}"/></a:xfrm><a:prstGeom prst="${geom}">${av}</a:prstGeom>${this.fill(fill, alpha)}${ln}</p:spPr>${text || '<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:endParaRPr lang="ru-RU" dirty="0"/></a:p></p:txBody>'}</p:sp>`);
  }
  /**
   * paras: [{ text, size, color, bold, italic, font, align, bullet, level, spaceBefore }]
   */
  text(x, y, w, h, paras, { anchor = "t", inset = 0.05, name = "Text" } = {}) {
    const id = this.id++;
    const body = paras.map(p => para(p, this.t)).join("") || '<a:p><a:endParaRPr lang="ru-RU"/></a:p>';
    const ins = E(inset);
    this.parts.push(`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name} ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${E(x)}" y="${E(y)}"/><a:ext cx="${E(w)}" cy="${E(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="${ins}" tIns="${ins}" rIns="${ins}" bIns="${ins}" rtlCol="0" anchor="${anchor}"><a:normAutofit/></a:bodyPr><a:lstStyle/>${body}</p:txBody></p:sp>`);
  }
  raw(xmlStr) { this.parts.push(xmlStr); }
  nextId() { return this.id++; }
  xml(bg) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld ${NS}><p:cSld><p:bg><p:bgPr>${this.fill(bg)}<a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${this.parts.join("")}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
  }
}

function runs(text, p, theme) {
  const list = inlineRuns(text);
  return list.map(r => {
    const bold = p.bold || r.b;
    // **highlighted** words inside normal text get the accent/heading colour
    const color = r.b && !p.bold ? (p.emphasis || p.color) : p.color;
    const font = r.code ? "Consolas" : p.font;
    return `<a:r><a:rPr lang="ru-RU" sz="${Math.round(p.size * 100)}"${bold ? ' b="1"' : ""}${p.italic || r.i ? ' i="1"' : ""} dirty="0"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:latin typeface="${xml(font)}"/><a:cs typeface="${xml(font)}"/></a:rPr><a:t>${xml(r.text)}</a:t></a:r>`;
  }).join("") || `<a:endParaRPr lang="ru-RU" sz="${Math.round(p.size * 100)}"/>`;
}

function para(p, theme) {
  const lvl = p.level || 0;
  const algn = p.align ? ` algn="${p.align}"` : "";
  const sp = `<a:lnSpc><a:spcPct val="${p.lineSpacing || 108}000"/></a:lnSpc><a:spcBef><a:spcPts val="${Math.round((p.spaceBefore ?? 0) * 100)}"/></a:spcBef>`;
  let bu = "<a:buNone/>", mar = ' marL="0" indent="0"';
  if (p.bullet) {
    const ind = 0.32 + lvl * 0.35;
    mar = ` marL="${E(ind)}" indent="${-E(0.3)}"`;
    bu = `<a:buClr><a:srgbClr val="${p.bulletColor || theme.accent}"/></a:buClr><a:buSzPct val="100000"/><a:buFont typeface="Arial"/><a:buChar char="${lvl ? "–" : "●"}"/>`;
  }
  return `<a:p><a:pPr${mar}${algn}>${sp}${bu}</a:pPr>${runs(p.text, p, theme)}</a:p>`;
}

const plain = s => String(s).replace(/\*\*|`|~~/g, "").replace(/(^|\s)[*_]([^*_]+)[*_](?=\s|$)/g, "$1$2");

/**
 * Picks the largest font size (pt) at which the paragraphs fit the box.
 * Rough model: average Cyrillic glyph ≈ 0.53 em (0.58 bold), line = 1.2 em.
 */
export function fitSize(texts, w, h, max, min, { bold = false, spacing = 1.22, gap = 0.35, font = "" } = {}) {
  const em = (bold ? 0.58 : 0.53) * (/georgia/i.test(font) ? 1.16 : 1);
  for (let s = max; s > min; s -= 1) {
    let lines = 0;
    for (const t of texts) {
      const width = w - 0.12 - (t.indent || 0);
      const perLine = Math.max(4, Math.floor(width / (s / 72 * em * (t.scale || 1))));
      const words = plain(t.text || t).split(/\s+/);
      let n = 1, cur = 0;
      for (const word of words) {
        const len = [...word].length;
        if (cur && cur + 1 + len > perLine) { n++; cur = len; } else cur += (cur ? 1 : 0) + len;
        while (cur > perLine) { n++; cur -= perLine; }
      }
      lines += n * (t.scale || 1);
    }
    const height = (lines * s * spacing + Math.max(0, texts.length - 1) * s * gap) / 72 + 0.1;
    if (height <= h) return s;
  }
  return min;
}

/* ------------------------------ deck parsing ------------------------------ */

const LAYOUTS = new Set(["title", "section", "bullets", "two-columns", "cards", "steps", "big-number", "quote", "table", "ending"]);
const LAYOUT_ALIASES = { columns: "two-columns", "2col": "two-columns", compare: "two-columns", comparison: "two-columns", grid: "cards", timeline: "steps", process: "steps", number: "big-number", stat: "big-number", fact: "big-number", list: "bullets", content: "bullets", end: "ending", thanks: "ending", cover: "title" };

/** Markdown deck -> [{ layout, title, lines, bullets, numbers, sections, rows, quote, author }] */
export function parseDeck(md) {
  const slides = [];
  let cur = null;
  const fresh = () => ({ layout: "", title: "", lines: [], bullets: [], numbers: [], sections: [], rows: [], quote: [], author: "" });
  const start = () => { cur = fresh(); slides.push(cur); };
  const target = () => (cur.sections.length ? cur.sections[cur.sections.length - 1] : cur);
  const lines = String(md).split("\n");
  for (let k = 0; k < lines.length; k++) {
    const line = lines[k];
    if (/^\s*(---+|\*\*\*+|___+)\s*$/.test(line)) { cur = null; continue; }
    let m;
    if ((m = /^\s*@layout\s*:?\s*([\w-]+)/i.exec(line))) {
      if (!cur || cur.title || cur.lines.length || cur.bullets.length) start();
      const l = m[1].toLowerCase();
      cur.layout = LAYOUTS.has(l) ? l : LAYOUT_ALIASES[l] || "";
      continue;
    }
    if (/^\s*@\w+/.test(line)) continue; // unknown directive
    if ((m = /^\s{0,3}#\s+(.+?)\s*#*\s*$/.exec(line))) {
      if (!cur || cur.title) start();
      cur.title = m[1].replace(/\*\*/g, "").trim();
      continue;
    }
    if (!cur) start();
    if ((m = /^\s{0,3}#{2,6}\s+(.+?)\s*#*\s*$/.exec(line))) {
      cur.sections.push({ title: m[1].replace(/\*\*/g, "").trim(), lines: [], bullets: [], numbers: [] });
      continue;
    }
    if (isTableRow(line)) { if (!isSeparator(line)) cur.rows.push(splitRow(line)); continue; }
    if ((m = /^\s*>\s?(.*)$/.exec(line))) { cur.quote.push(m[1]); continue; }
    if (cur.quote.length && (m = /^\s*[—–-]{1,2}\s+(.+)$/.exec(line)) && !cur.bullets.length) { cur.author = m[1]; continue; }
    if ((m = /^(\s*)[-*+•]\s+(.*)$/.exec(line))) { target().bullets.push({ text: m[2], level: Math.min(Math.floor(m[1].replace(/\t/g, "  ").length / 2), 2) }); continue; }
    if ((m = /^\s*(\d{1,2})[.)]\s+(.*)$/.exec(line))) { target().numbers.push({ n: Number(m[1]), text: m[2] }); continue; }
    if (line.trim()) target().lines.push(line.trim());
  }
  return slides.filter(s => s.title || s.lines.length || s.bullets.length || s.sections.length || s.rows.length || s.quote.length || s.numbers.length);
}

function inferLayout(s, k, total) {
  if (s.layout) return s.layout;
  const bodyCount = s.bullets.length + s.numbers.length + s.sections.length + s.rows.length + s.quote.length;
  if (k === 0) return "title";
  if (s.rows.length) return "table";
  if (s.quote.length) return "quote";
  if (/^[≈~<>]?\s*[-+−]?[$€₽]?\s*\d[\d\s.,]*(\s*(%|‰|×|x|млн|млрд|трлн|тыс\.?|руб\.?|₽|\$|€|лет|года?|раза?|км|кг|т|°|чел\.?))*\s*\+?$/i.test(s.title) && bodyCount <= 2) return "big-number";
  if (k === total - 1 && /спасибо|вопрос|благодар|thank/i.test(s.title) && bodyCount <= 2) return "ending";
  if (s.sections.length === 2) return "two-columns";
  if (s.sections.length >= 3) return "cards";
  if (s.numbers.length >= 3 && !s.bullets.length) return "steps";
  if (!bodyCount && s.title && (!s.lines.length || (s.lines.length === 1 && plain(s.lines[0]).length <= 90))) return "section";
  return "bullets";
}

/* ------------------------------ layouts ------------------------------ */

function headerBlock(c, t, s, num) {
  const titleSize = fitSize([s.title], 11.6, 1.0, 34, 22, { bold: true, font: t.headFont });
  c.shape(0.6, 0.62, 0.75, 0.08, { fill: t.accent, name: "Accent" });
  c.text(0.55, 0.75, 11.9, 1.05, [{ text: s.title, size: titleSize, bold: true, color: t.heading, font: t.headFont }], { anchor: "ctr", name: "Title" });
  c.text(11.9, 6.95, 0.9, 0.35, [{ text: String(num), size: 12, color: t.muted, font: t.font, align: "r" }], { name: "Slide number" });
  c.shape(0.6, 7.02, 0.35, 0.05, { fill: t.line, name: "Footer" });
}

const AREA = { x: 0.6, y: 1.95, w: 12.13, h: 4.85 };

function bodyParas(items, t, size, { color = t.text, emphasis = t.heading } = {}) {
  return items.map((it, k) => ({
    text: it.text, level: it.level || 0, bullet: it.bullet !== false, size: it.level ? size * 0.85 : size,
    color, emphasis, font: t.font, spaceBefore: k ? (it.level ? size * 0.25 : size * 0.55) : 0,
  }));
}

function itemsOf(node) {
  if (node._order) return node._order;
  return [
    ...node.lines.map(text => ({ text, bullet: false })),
    ...node.bullets.map(b => ({ text: b.text, level: b.level, bullet: true })),
    ...node.numbers.map(n => ({ text: `${n.n}. ${n.text}`, bullet: false })),
  ];
}

function layoutBullets(c, t, s) {
  let items = itemsOf(s);
  // sub-sections on a bullet slide become bold lines
  for (const sec of s.sections) items.push({ text: `**${sec.title}**`, bullet: false }, ...itemsOf(sec));
  if (!items.length) items = [{ text: " ", bullet: false }];
  const texts = items.map(i => ({ text: i.text, indent: i.bullet ? 0.35 + (i.level || 0) * 0.35 : 0, scale: i.level ? 0.85 : 1 }));
  const short = items.length <= 4 && items.every(i => plain(i.text).length < 90);
  const wBox = short ? 10.5 : AREA.w;
  const size = fitSize(texts, wBox, AREA.h, short ? 30 : 26, 14);
  c.text(AREA.x, AREA.y, wBox, AREA.h, bodyParas(items, t, size), { anchor: short ? "ctr" : "t", name: "Body" });
}

function columnCard(c, t, x, y, w, h, title, items, color) {
  c.shape(x, y, w, h, { geom: "roundRect", adj: 4500, fill: t.card, line: t.darkContent ? null : t.line, name: "Card" });
  c.shape(x, y + 0.25, 0.09, 0.65, { fill: color, name: "Card accent" });
  const headSize = fitSize([title], w - 0.6, 0.75, 26, 16, { bold: true, font: t.headFont });
  c.text(x + 0.3, y + 0.2, w - 0.55, 0.75, [{ text: title, size: headSize, bold: true, color: t.heading, font: t.headFont }], { anchor: "ctr" });
  const texts = items.map(i => ({ text: i.text, indent: i.bullet ? 0.35 : 0, scale: i.level ? 0.85 : 1 }));
  const size = fitSize(texts.length ? texts : [" "], w - 0.6, h - 1.25, 24, 12);
  c.text(x + 0.3, y + 1.05, w - 0.55, h - 1.25, bodyParas(items, t, size), { name: "Card body" });
}

function layoutTwoColumns(c, t, s) {
  let secs = s.sections.slice(0, 2);
  if (secs.length < 2) {
    const all = itemsOf(s);
    const half = Math.ceil(all.length / 2);
    secs = [{ title: "", items: all.slice(0, half) }, { title: "", items: all.slice(half) }];
  } else secs = secs.map(x => ({ title: x.title, items: itemsOf(x) }));
  let y = AREA.y, h = AREA.h;
  const intro = s.sections.length >= 2 ? itemsOf(s).filter(i => !i.bullet) : [];
  if (intro.length) {
    c.text(AREA.x, y - 0.1, AREA.w, 0.6, [{ text: intro.map(i => i.text).join(" "), size: 18, color: t.muted, font: t.font }]);
    y += 0.55; h -= 0.55;
  }
  const w = (AREA.w - 0.35) / 2;
  secs.forEach((sec, k) => columnCard(c, t, AREA.x + k * (w + 0.35), y, w, h, sec.title || " ", sec.items, k ? t.accent2 : t.accent));
}

function layoutCards(c, t, s) {
  const secs = s.sections.slice(0, 6);
  const n = secs.length || 1;
  const rows = n > 4 ? 2 : 1;
  const perRow = Math.ceil(n / rows);
  const gap = 0.3;
  const w = (AREA.w - gap * (perRow - 1)) / perRow;
  const h = (AREA.h - gap * (rows - 1)) / rows;
  const palette = t.darkContent ? [t.accent, t.accent2, t.muted] : [t.accent, t.primary, t.accent2, t.heading];
  secs.forEach((sec, k) => {
    const x = AREA.x + (k % perRow) * (w + gap);
    const y = AREA.y + Math.floor(k / perRow) * (h + gap);
    const col = palette[k % palette.length];
    c.shape(x, y, w, h, { geom: "roundRect", adj: 5000, fill: t.card, line: t.darkContent ? null : t.line, name: "Card" });
    c.shape(x, y, w, 0.1, { fill: col, name: "Card top" });
    const compact = rows === 2;
    const d = compact ? 0.5 : 0.62;
    const badge = `<p:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"/><a:lstStyle/>${para({ text: String(k + 1), size: compact ? 16 : 20, bold: true, color: col === t.primary || col === t.heading ? t.onDark : t.dark, font: t.headFont, align: "ctr" }, t)}</p:txBody>`;
    let top;
    if (compact) {
      // badge and title on one line
      c.shape(x + 0.25, y + 0.3, d, d, { geom: "ellipse", fill: col, name: "Badge", text: badge });
      const hs = fitSize([sec.title], w - d - 0.6, d + 0.1, 20, 12, { bold: true, font: t.headFont });
      c.text(x + 0.25 + d + 0.12, y + 0.25, w - d - 0.55, d + 0.1, [{ text: sec.title, size: hs, bold: true, color: t.heading, font: t.headFont }], { anchor: "ctr" });
      top = 0.3 + d + 0.18;
    } else {
      c.shape(x + 0.3, y + 0.35, d, d, { geom: "ellipse", fill: col, name: "Badge", text: badge });
      const hs = fitSize([sec.title], w - 0.55, 0.85, 26, 13, { bold: true, font: t.headFont });
      const lines = Math.max(1, Math.ceil([...plain(sec.title)].length * hs / 72 * 0.6 / (w - 0.6)));
      c.text(x + 0.25, y + 0.35 + d + 0.12, w - 0.45, 0.85, [{ text: sec.title, size: hs, bold: true, color: t.heading, font: t.headFont }]);
      top = 0.35 + d + 0.12 + Math.min(0.85, lines * hs * 1.25 / 72 + 0.12);
    }
    const items = itemsOf(sec);
    const texts = items.map(i => ({ text: i.text, indent: i.bullet ? 0.35 : 0 }));
    const size = fitSize(texts.length ? texts : [" "], w - 0.5, h - top - 0.2, rows === 2 ? 17 : 21, 11);
    c.text(x + 0.25, y + top, w - 0.45, h - top - 0.2, bodyParas(items, t, size, { color: t.muted }));
  });
}

function splitStep(text) {
  const m = /^(.{2,60}?)(?:\s*[:—–]\s+|\s+-\s+)(.+)$/.exec(plain(text));
  return m ? [m[1], m[2]] : [plain(text), ""];
}

const stepColor = (t, k) => (k % 2 ? t.accent : t.darkContent ? t.accent2 : t.primary);

function layoutSteps(c, t, s) {
  const steps = (s.numbers.length ? s.numbers.map(n => n.text) : s.bullets.map(b => b.text)).slice(0, 8);
  const n = steps.length || 1;
  if (n <= 5) {
    const slot = AREA.w / n;
    const d = 0.85, cy = AREA.y + 0.55;
    if (n > 1) c.shape(AREA.x + slot / 2, cy + d / 2 - 0.03, slot * (n - 1), 0.06, { fill: t.line, name: "Line" });
    steps.forEach((st, k) => {
      const cx = AREA.x + slot * k + slot / 2;
      c.shape(cx - d / 2, cy, d, d, { geom: "ellipse", fill: stepColor(t, k), line: t.bg, lineW: 4, name: "Step",
        text: `<p:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"/><a:lstStyle/>${para({ text: String(k + 1), size: 24, bold: true, color: stepColor(t, k) === t.primary ? t.onDark : t.dark, font: t.headFont, align: "ctr" }, t)}</p:txBody>` });
      const [head, desc] = splitStep(st);
      const w = slot - 0.25;
      const hs = fitSize([head], w, 0.9, 22, 13, { bold: true, font: t.headFont });
      c.text(cx - w / 2, cy + d + 0.25, w, 0.95, [{ text: head, size: hs, bold: true, color: t.heading, font: t.headFont, align: "ctr" }], { anchor: "t" });
      if (desc) {
        const ds = fitSize([desc], w, 2.4, 19, 11);
        c.text(cx - w / 2, cy + d + 1.15, w, 2.5, [{ text: desc, size: ds, color: t.muted, font: t.font, align: "ctr" }]);
      }
    });
  } else {
    const perCol = Math.ceil(n / 2);
    const colW = (AREA.w - 0.4) / 2;
    const rowH = AREA.h / perCol;
    steps.forEach((st, k) => {
      const x = AREA.x + Math.floor(k / perCol) * (colW + 0.4);
      const y = AREA.y + (k % perCol) * rowH;
      const d = Math.min(0.6, rowH - 0.15);
      c.shape(x, y + (rowH - d) / 2, d, d, { geom: "ellipse", fill: stepColor(t, k), name: "Step",
        text: `<p:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"/><a:lstStyle/>${para({ text: String(k + 1), size: 16, bold: true, color: stepColor(t, k) === t.primary ? t.onDark : t.dark, font: t.headFont, align: "ctr" }, t)}</p:txBody>` });
      const size = fitSize([st], colW - d - 0.3, rowH - 0.1, 18, 11);
      c.text(x + d + 0.2, y, colW - d - 0.25, rowH, [{ text: st, size, color: t.text, emphasis: t.heading, font: t.font }], { anchor: "ctr" });
    });
  }
}

function layoutBigNumber(c, t, s) {
  const caption = itemsOf(s);
  const ns = fitSize([s.title], 6.2, 2.6, 130, 48, { bold: true, spacing: 1.0, font: t.headFont });
  c.shape(0.6, 1.2, 0.12, 5.1, { fill: t.accent, name: "Accent" });
  c.text(1.0, 1.2, 6.4, 5.1, [{ text: s.title, size: ns, bold: true, color: t.heading, font: t.headFont, lineSpacing: 90 }], { anchor: "ctr" });
  const texts = caption.map(i => ({ text: i.text }));
  const size = fitSize(texts.length ? texts : [" "], 5.2, 4.8, 30, 14);
  c.text(7.6, 1.3, 5.2, 4.9, bodyParas(caption.map(i => ({ ...i, bullet: false })), t, size, { color: t.text }), { anchor: "ctr" });
  c.shape(11.4, 5.6, 1.4, 1.4, { geom: "ellipse", fill: t.accent, alpha: 15, name: "Deco" });
}

function layoutQuote(c, t, s) {
  const q = s.quote.join(" ").trim() || itemsOf(s).map(i => i.text).join(" ");
  c.shape(0, 0, W, H, { fill: t.dark, name: "Background" });
  c.text(0.8, 0.5, 2, 2, [{ text: "“", size: 160, bold: true, color: t.accent, font: "Georgia" }]);
  const size = fitSize([q], 10.8, 3.6, 36, 18, { spacing: 1.3 });
  c.text(1.3, 2.0, 10.8, 3.7, [{ text: q, size, italic: true, color: t.onDark, emphasis: t.accent, font: t.headFont, lineSpacing: 120 }], { anchor: "ctr" });
  const author = s.author || (s.title && s.quote.length ? s.title : "");
  if (author) {
    c.shape(1.3, 5.95, 0.6, 0.06, { fill: t.accent });
    c.text(2.05, 5.7, 9, 0.6, [{ text: author, size: 22, color: t.onDarkMuted, font: t.font }], { anchor: "ctr" });
  }
}

function layoutTable(c, t, s) {
  const rows = s.rows.slice(0, 14);
  const cols = Math.max(...rows.map(r => r.length));
  let y = AREA.y;
  const intro = itemsOf(s);
  if (intro.length) {
    c.text(AREA.x, y - 0.1, AREA.w, 0.55, [{ text: intro.map(i => i.text).join(" "), size: 17, color: t.muted, font: t.font }]);
    y += 0.5;
  }
  const size = rows.length > 10 || cols > 6 ? 13 : rows.length > 7 || cols > 4 ? 15 : 18;
  const rowH = Math.min(0.68, (6.85 - y) / rows.length);
  // column widths proportional to content length
  const lens = Array.from({ length: cols }, (_, ci) => Math.max(4, ...rows.map(r => plain(r[ci] ?? "").length)));
  const sum = lens.reduce((a, b) => a + b, 0);
  const widths = lens.map(l => Math.max(1.0, AREA.w * l / sum));
  const scale = AREA.w / widths.reduce((a, b) => a + b, 0);
  const grid = widths.map(w => `<a:gridCol w="${E(w * scale)}"/>`).join("");
  const border = col => `<a:lnL w="6350"><a:solidFill><a:srgbClr val="${col}"/></a:solidFill></a:lnL><a:lnR w="6350"><a:solidFill><a:srgbClr val="${col}"/></a:solidFill></a:lnR><a:lnT w="6350"><a:solidFill><a:srgbClr val="${col}"/></a:solidFill></a:lnT><a:lnB w="6350"><a:solidFill><a:srgbClr val="${col}"/></a:solidFill></a:lnB>`;
  const trs = rows.map((r, ri) => {
    const head = ri === 0;
    const fill = head ? (t.darkContent ? t.accent : t.primary) : ri % 2 ? t.card : t.bg;
    const tcs = Array.from({ length: cols }, (_, ci) => {
      const p = { text: r[ci] ?? "", size, bold: head, color: head ? (t.darkContent ? t.dark : t.onDark) : t.text, emphasis: t.heading, font: t.font, align: ci && /^[-−+]?[\d\s.,]+%?$/.test(plain(r[ci] ?? "")) ? "r" : undefined };
      return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${para(p, t)}</a:txBody><a:tcPr marL="${E(0.12)}" marR="${E(0.12)}" marT="${E(0.05)}" marB="${E(0.05)}" anchor="ctr">${border(t.line)}<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill></a:tcPr></a:tc>`;
    }).join("");
    return `<a:tr h="${E(rowH)}">${tcs}</a:tr>`;
  }).join("");
  const id = c.nextId();
  c.raw(`<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${E(AREA.x)}" y="${E(y)}"/><a:ext cx="${E(AREA.w)}" cy="${E(rowH * rows.length)}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid>${grid}</a:tblGrid>${trs}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`);
}

function layoutCover(c, t, s, ending) {
  c.shape(0, 0, W, H, { fill: t.dark, name: "Background" });
  c.shape(8.6, -1.6, 6.6, 6.6, { geom: "ellipse", fill: t.accent, alpha: 16, name: "Deco" });
  c.shape(10.4, 3.6, 4.6, 4.6, { geom: "ellipse", fill: t.accent2, alpha: 14, name: "Deco" });
  c.shape(0, H - 0.12, W, 0.12, { fill: t.accent, name: "Stripe" });
  const sub = itemsOf(s).map(i => i.text);
  const ts = fitSize([s.title], 9.4, 2.4, ending ? 54 : 50, 28, { bold: true, spacing: 1.05, font: t.headFont });
  c.text(0.85, 1.0, 9.6, 2.85, [{ text: s.title, size: ts, bold: true, color: t.onDark, font: t.headFont, lineSpacing: 95 }], { anchor: "b" });
  c.shape(0.9, 4.1, 1.3, 0.09, { fill: t.accent, name: "Accent" });
  if (sub.length) {
    const ss = fitSize(sub, 9.4, 1.9, 24, 13);
    c.text(0.85, 4.35, 9.6, 2.1, sub.map((x, k) => ({ text: x, size: k ? ss * 0.85 : ss, color: t.onDarkMuted, emphasis: t.accent, font: t.font, spaceBefore: k ? 6 : 0 })));
  }
}

function layoutSection(c, t, s, sectionNo) {
  c.shape(0, 0, W, H, { fill: t.darkContent ? t.dark : t.primary, name: "Background" });
  c.shape(0, 0, 0.35, H, { fill: t.accent, name: "Side" });
  c.text(0.95, 1.3, 4, 1.8, [{ text: String(sectionNo).padStart(2, "0"), size: 88, bold: true, color: t.accent, font: t.headFont }], { anchor: "b" });
  const ts = fitSize([s.title], 11, 2.2, 46, 26, { bold: true, font: t.headFont });
  c.text(0.95, 3.2, 11.3, 2.3, [{ text: s.title, size: ts, bold: true, color: t.onDark, font: t.headFont }]);
  const sub = itemsOf(s).map(i => i.text).join(" ");
  if (sub) c.text(0.95, 5.5, 11.3, 1.0, [{ text: sub, size: 20, color: t.onDarkMuted, font: t.font }]);
}

/* ------------------------------ build ------------------------------ */

export function renderSlide(s, k, total, theme, sectionNo) {
  const t = THEMES[theme];
  const layout = inferLayout(s, k, total);
  const c = new Canvas(t);
  let bg = t.bg;
  if (layout === "title" || layout === "ending") { layoutCover(c, t, s, layout === "ending"); bg = t.dark; }
  else if (layout === "section") { layoutSection(c, t, s, sectionNo); bg = t.primary; }
  else if (layout === "quote") { layoutQuote(c, t, s); bg = t.dark; }
  else {
    if (s.title && layout !== "big-number") headerBlock(c, t, s, k + 1);
    else c.text(11.9, 6.95, 0.9, 0.35, [{ text: String(k + 1), size: 12, color: t.muted, font: t.font, align: "r" }]);
    if (layout === "two-columns") layoutTwoColumns(c, t, s);
    else if (layout === "cards") layoutCards(c, t, s);
    else if (layout === "steps") layoutSteps(c, t, s);
    else if (layout === "big-number") layoutBigNumber(c, t, s);
    else if (layout === "table") s.rows.length ? layoutTable(c, t, s) : layoutBullets(c, t, s);
    else layoutBullets(c, t, s);
  }
  return { xml: c.xml(bg), layout };
}

// A bullet slide with too much text is split into "(продолжение)" slides.
export function splitOverfull(slides) {
  const out = [];
  slides.forEach((s, k) => {
    const layout = inferLayout(s, k, slides.length);
    const items = itemsOf(s);
    const chars = items.reduce((n, i) => n + plain(i.text).length, 0);
    if (layout !== "bullets" || s.sections.length || (items.length <= 7 && chars <= 750)) { out.push(s); return; }
    const chunks = [];
    let cur = [], size = 0;
    for (const it of items) {
      const len = plain(it.text).length;
      if (cur.length && (cur.length >= 6 || size + len > 560) && !(it.level)) { chunks.push(cur); cur = []; size = 0; }
      cur.push(it); size += len;
    }
    if (cur.length) chunks.push(cur);
    chunks.forEach((ch, i) => out.push({
      ...s, layout: "bullets", title: i ? `${s.title} (продолжение)` : s.title,
      lines: ch.filter(x => !x.bullet).map(x => x.text),
      bullets: ch.filter(x => x.bullet).map(x => ({ text: x.text, level: x.level || 0 })),
      numbers: [], _order: ch,
    }));
  });
  return out;
}

export function buildPptx(markdown, { theme } = {}) {
  const themeName = pickTheme(theme);
  let slides = splitOverfull(parseDeck(prepare(markdown))).slice(0, 40);
  if (!slides.length) slides = [{ ...parseDeck("# Презентация")[0] }];
  let sectionNo = 0;
  const rendered = slides.map((s, k) => {
    const layout = inferLayout(s, k, slides.length);
    if (layout === "section") sectionNo++;
    return renderSlide(s, k, slides.length, themeName, sectionNo);
  });
  const rels = (target, type) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/${type}" Target="${target}"/></Relationships>`;
  const sldIds = rendered.map((_, k) => `<p:sldId id="${256 + k}" r:id="rId${10 + k}"/>`).join("");
  const presentation = T.PRESENTATION
    .replace("{{SLIDES}}", `<p:sldIdLst>${sldIds}</p:sldIdLst>`)
    .replace(/<p:sldSz[^>]*\/>/, `<p:sldSz cx="${E(W)}" cy="${E(H)}"/>`);
  const files = [
    { name: "[Content_Types].xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/><Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/><Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>${rendered.map((_, k) => `<Override PartName="/ppt/slides/slide${k + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join("")}</Types>` },
    { name: "_rels/.rels", data: rels("ppt/presentation.xml", "officeDocument") },
    { name: "ppt/presentation.xml", data: presentation },
    { name: "ppt/_rels/presentation.xml.rels", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId3" Type="${REL}/presProps" Target="presProps.xml"/><Relationship Id="rId4" Type="${REL}/viewProps" Target="viewProps.xml"/><Relationship Id="rId5" Type="${REL}/theme" Target="theme/theme1.xml"/><Relationship Id="rId6" Type="${REL}/tableStyles" Target="tableStyles.xml"/>${rendered.map((_, k) => `<Relationship Id="rId${10 + k}" Type="${REL}/slide" Target="slides/slide${k + 1}.xml"/>`).join("")}</Relationships>` },
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
  rendered.forEach((r, k) => {
    files.push({ name: `ppt/slides/slide${k + 1}.xml`, data: r.xml });
    // custom-drawn slides sit on the "Title Slide"/"Title and Content" layouts without using their placeholders
    files.push({ name: `ppt/slides/_rels/slide${k + 1}.xml.rels`, data: rels(`../slideLayouts/slideLayout${r.layout === "title" ? 1 : 2}.xml`, "slideLayout") });
  });
  return createZip(files);
}
