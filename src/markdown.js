// Shared Markdown helpers for building Office files.
import { convertLatex } from "./format.js";

export function xml(s) {
  return String(s ?? "")
    // characters not allowed in XML 1.0
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, "")
    .replace(/[\uE000-\uE0FF]/g, "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/** Markdown text with LaTeX converted to Unicode; boxed answers become **bold**. */
export function prepare(md) {
  const parts = String(md).replace(/\r\n?/g, "\n").split(/(```[\s\S]*?(?:```|$))/);
  return parts.map((p, k) => (k % 2 ? p : convertLatex(p)))
    .join("")
    .replace(/[\uE000\uE001]/g, "**").replace(/[\uE002\uE003]/g, "*").replace(/[\uE004\uE005]/g, "~~");
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

export const isTableRow = l => /^\s*\|.*\|\s*$/.test(l);
export const isSeparator = l => /^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$/.test(l);
export const splitRow = l => l.trim().replace(/^\|/, "").replace(/\|$/, "").split("|").map(c => c.trim());

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

