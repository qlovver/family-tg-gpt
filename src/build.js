// Building Word / Excel / PowerPoint files from the model's Markdown.
// The model asks for a file with a block:
//   <file type="docx|xlsx|pptx" name="Имя.docx"> ...markdown... </file>

import { xml, inlineRuns, parseBlocks } from "./markdown.js";
import { buildPptx, parseDeck, THEMES } from "./slides.js";
import { buildDocx } from "./docx.js";
import { buildXlsx, parseSheets } from "./xlsx.js";

export { xml, inlineRuns, parseBlocks, buildPptx, parseDeck, THEMES, buildDocx, buildXlsx, parseSheets };

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
  const text = String(answer).replace(re, (_, attrs_, content) => {
    const attrs = {};
    for (const m of attrs_.matchAll(/([\w-]+)\s*=\s*"([^"]*)"/g)) attrs[m[1].toLowerCase()] = m[2];
    const type = (attrs.type || "").toLowerCase();
    if (!MIME[type]) return _;
    files.push({
      type, name: safeFileName(attrs.name, type), content: content.replace(/\s+$/, ""), attrs,
      font: attrs.font, size: Number(attrs.size) || undefined, theme: attrs.theme, style: attrs.style,
    });
    return `\n[[FILE:${files.length - 1}]]\n`;
  });
  return { text: text.replace(/\n?\[\[FILE:\d+\]\]\n?/g, "\n").replace(/\n{3,}/g, "\n\n").trim(), files };
}

/* ------------------------------ entry ------------------------------ */

export function buildFile(block) {
  const attrs = { ...(block.attrs || {}), font: block.font ?? block.attrs?.font, size: block.size ?? block.attrs?.size, theme: block.theme ?? block.attrs?.theme, style: block.style ?? block.attrs?.style };
  if (block.type === "docx") return buildDocx(block.content, { ...attrs, title: block.name.replace(/\.docx$/i, "") });
  if (block.type === "xlsx") return buildXlsx(block.content, attrs);
  if (block.type === "pptx") return buildPptx(block.content, attrs);
  throw new Error("unknown_file_type");
}
