import { test } from "node:test";
import assert from "node:assert/strict";
import { createWorker, historyKey, fileKey } from "../src/bot.js";
import { buildDocx, buildXlsx, buildPptx, extractFileBlocks, safeFileName, parseSheets, MIME } from "../src/build.js";
import { parseDeck, splitOverfull, pickTheme, THEMES, fitSize } from "../src/slides.js";
import { extractDocx, extractXlsx, extractPptx, classifyDocument, decodeText } from "../src/files.js";
import { createZip, openZip } from "../src/zip.js";
import { makeEnv, makeFetch, messagesReply, textUpdate, documentUpdate, webhookRequest, execCtx, sentTexts } from "./helpers.js";

async function send(env, f, update) {
  return createWorker({ fetchImpl: f.fetch }).fetch(webhookRequest(update), env, execCtx);
}

const SAMPLE_DOC = `# Реферат
## Введение
Текст с **жирным** и формулой $\\frac{1}{2}$ & <теги>.
- пункт
1. первый
2. второй
| A | B |
|---|---|
| 1 | 2 |`;

test("zip round trip (stored) and reading deflated entries", async () => {
  const z = createZip([{ name: "a.txt", data: "привет" }, { name: "dir/b.bin", data: new Uint8Array([1, 2, 3]) }]);
  const r = openZip(z);
  assert.deepEqual(r.names, ["a.txt", "dir/b.bin"]);
  assert.equal(await r.text("a.txt"), "привет");
  // deflated entry produced by CompressionStream, like real Office files
  const raw = new TextEncoder().encode("сжатый текст ".repeat(50));
  const comp = new Uint8Array(await new Response(new Blob([raw]).stream().pipeThrough(new CompressionStream("deflate-raw"))).arrayBuffer());
  const z2 = createZip([{ name: "x", data: comp }]);
  // patch method=8 and sizes in both headers to simulate a deflated entry
  const v = new DataView(z2.buffer);
  v.setUint16(8, 8, true); v.setUint32(22, raw.length, true);
  const cd = z2.length - 22 - 46 - 1;
  v.setUint16(cd + 10, 8, true); v.setUint32(cd + 24, raw.length, true);
  assert.equal(await openZip(z2).text("x"), "сжатый текст ".repeat(50));
});

test("docx: build then extract keeps headings, lists, tables, escaped text", async () => {
  const bytes = buildDocx(SAMPLE_DOC, { font: "Times New Roman", size: 14 });
  const zip = openZip(bytes);
  assert.ok(zip.names.includes("word/document.xml") && zip.names.includes("word/numbering.xml"));
  assert.match(await zip.text("word/styles.xml"), /Times New Roman/);
  const { text } = await extractDocx(bytes);
  assert.match(text, /# Реферат/);
  assert.match(text, /# Введение/);
  assert.match(text, /Текст с жирным и формулой ½ & <теги>\./);
  assert.match(text, /- первый/);
  assert.match(text, /\| A \| B \|\n\| 1 \| 2 \|/);
});

test("xlsx: numbers, percents, formulas, several sheets", async () => {
  const md = "## Расходы\n| Статья | Сумма | Доля |\n|---|---|---|\n| Еда | 15 000 | 30% |\n| Итого | =SUM(B2:B2) | 0,5 |\n## Второй\n| x |\n|---|\n| 1 |";
  const sheets = parseSheets(md);
  assert.deepEqual(sheets.map(s => s.name), ["Расходы", "Второй"]);
  const bytes = buildXlsx(md);
  const sheet1 = await openZip(bytes).text("xl/worksheets/sheet1.xml");
  assert.match(sheet1, /<c r="B2" s="\d+"><v>15000<\/v><\/c>/);
  assert.match(sheet1, /<c r="C2" s="\d+"><v>0.3<\/v><\/c>/);
  assert.match(sheet1, /<f>SUM\(B2:B2\)<\/f>/);
  assert.match(sheet1, /<v>0.5<\/v>/);
  const { text } = await extractXlsx(bytes);
  assert.match(text, /## Лист «Расходы»/);
  assert.match(text, /B3: =SUM\(B2:B2\)/);
  assert.match(text, /## Лист «Второй»/);
});

test("pptx: designed 16:9 deck, layouts inferred, text extractable", async () => {
  const md = "# Модель Солоу\nЭкономика\n---\n# Теория\n---\n# Уравнение\n- k̇ = s√k − δk\n---\n# Факторы\n## Капитал\nтекст\n## Труд\nтекст\n## Технологии\nтекст\n---\n# 73%\nподпись\n---\n> цитата\n— Автор\n---\n# Спасибо за внимание!";
  const bytes = buildPptx(md, { theme: "night" });
  const zip = openZip(bytes);
  assert.ok(zip.names.includes("ppt/slides/slide7.xml"));
  assert.match(await zip.text("ppt/presentation.xml"), /<p:sldSz cx="12191695" cy="6858000"\/>/);
  assert.match(await zip.text("ppt/slides/slide1.xml"), /0B1120/, "night theme colours used");
  assert.match(await zip.text("ppt/slides/slide2.xml"), />01</, "section number");
  assert.match(await zip.text("ppt/slides/slide4.xml"), /Технологии/);
  const { text } = await extractPptx(bytes);
  assert.match(text, /Уравнение[\s\S]*k̇ = s√k − δk/);
  assert.match(text, /73%/);
});

test("deck parsing and layout directives", () => {
  const deck = parseDeck("@layout: cards\n# A\n## x\n1\n## y\n2\n---\n# B\n1. a\n2. b\n3. c\n---\n@layout: compare\n# C\n- p\n- q");
  assert.equal(deck.length, 3);
  assert.equal(deck[0].layout, "cards");
  assert.equal(deck[0].sections.length, 2);
  assert.equal(deck[1].numbers.length, 3);
  assert.equal(deck[2].layout, "two-columns");
  // slides without "---" are split on "# " headings
  assert.equal(parseDeck("# A\n- x\n# B\n- y").length, 2);
});

test("themes: names, Russian aliases, fallback", () => {
  assert.equal(pickTheme("forest"), "forest");
  assert.equal(pickTheme("Тёмная"), "night");
  assert.equal(pickTheme("что-то"), "ocean");
  for (const t of Object.values(THEMES)) for (const k of ["dark", "primary", "accent", "bg", "card", "text"]) assert.match(t[k], /^[0-9A-F]{6}$/);
});

test("overfull bullet slides are split, short text gets a big font", () => {
  const long = "очень длинный пункт ".repeat(8);
  const deck = parseDeck("# T\nsub\n---\n# Много\n" + Array.from({ length: 12 }, () => "- " + long).join("\n"));
  const out = splitOverfull(deck);
  assert.ok(out.length >= 3);
  assert.match(out[2].title, /продолжение/);
  assert.ok(fitSize(["коротко"], 10, 4, 30, 12) === 30);
  assert.ok(fitSize([long.repeat(5)], 10, 1, 30, 12) < 20);
});

test("file blocks are extracted from the answer", () => {
  const { text, files } = extractFileBlocks('Вот файл:\n<file type="xlsx" name="Бюджет">\n| a |\n|---|\n| 1 |\n</file>\nГотово.');
  assert.equal(text, "Вот файл:\n\nГотово.");
  assert.equal(files[0].type, "xlsx");
  assert.equal(files[0].name, "Бюджет.xlsx");
  assert.equal(extractFileBlocks('<file type="exe" name="x">bad</file>').files.length, 0);
  assert.equal(safeFileName('../../etc/pa"ss<>.txt', "docx"), "....etcpass.docx");
  assert.equal(safeFileName("", "pptx"), "Презентация.pptx");
});

test("classifies documents", () => {
  assert.equal(classifyDocument({ file_name: "a.PDF" }).kind, "pdf");
  assert.equal(classifyDocument({ file_name: "a.xlsm" }).kind, "xlsx");
  assert.equal(classifyDocument({ file_name: "main.py" }).kind, "text");
  assert.equal(classifyDocument({ file_name: "a.ppt" }).hint, "pptx");
  assert.equal(classifyDocument({ file_name: "big.zip", file_size: 30_000_000 }).error, "file_too_large");
  assert.equal(classifyDocument({ file_name: "photo.heic" }).error, "heic_unsupported");
  assert.equal(decodeText(new Uint8Array([0xcf, 0xf0, 0xe8, 0xe2, 0xe5, 0xf2])), "Привет"); // windows-1251
});

test("bot: Word file -> text goes to the model and is remembered for follow-ups", async () => {
  const env = makeEnv();
  const f = makeFetch({ files: { doc1: buildDocx(SAMPLE_DOC) } });
  await send(env, f, documentUpdate(111, "doc1", "реферат.docx", MIME.docx, "проверь ошибки"));
  const sent = f.calls.model[0].body.messages.at(-1).content.at(-1).text;
  assert.match(sent, /<document name="реферат.docx">[\s\S]*\[P\d+\] # Введение[\s\S]*<\/document>/);
  assert.match(sent, /проверь ошибки$/);
  assert.ok(env.HISTORY.map.has(fileKey("111")));
  const rec = JSON.parse(env.HISTORY.map.get(historyKey("111")));
  assert.equal(rec.messages[0].content, "[Файл «реферат.docx»] проверь ошибки");

  await send(env, f, textUpdate(111, "а сколько там разделов?"));
  assert.match(f.calls.model[1].body.messages.at(-1).content.at(-1).text, /Ранее в этом диалоге прислан файл «реферат.docx»[\s\S]*Введение/);

  // the other user never sees this file
  await send(env, f, textUpdate(222, "что в файле?"));
  assert.ok(!JSON.stringify(f.calls.model[2].body).includes("Введение"));

  await send(env, f, textUpdate(111, "/reset"));
  assert.ok(!env.HISTORY.map.has(fileKey("111")));
});

test("bot: PDF is attached as a document block", async () => {
  const env = makeEnv();
  const pdf = new TextEncoder().encode("%PDF-1.4 fake");
  const f = makeFetch({ files: { p1: pdf } });
  await send(env, f, documentUpdate(111, "p1", "задание.pdf", "application/pdf"));
  const content = f.calls.model[0].body.messages.at(-1).content;
  assert.equal(content[0].type, "document");
  assert.equal(content[0].source.media_type, "application/pdf");
  assert.equal(Buffer.from(content[0].source.data, "base64").toString(), "%PDF-1.4 fake");
});

test("bot: PDF rejected by the provider gives a clear message", async () => {
  const env = makeEnv();
  const f = makeFetch({ files: { p1: new TextEncoder().encode("%PDF") }, model: () => ({ status: 400, json: { error: { message: "document not supported" } } }) });
  await send(env, f, documentUpdate(111, "p1", "a.pdf", "application/pdf"));
  assert.match(sentTexts(f.calls).at(-1), /не принял этот PDF/);
});

test("bot: model answer with a <file> block is sent as a real document", async () => {
  const env = makeEnv();
  const answer = 'Сделал таблицу.\n<file type="xlsx" name="Бюджет">\n## Март\n| Статья | Сумма |\n|---|---|\n| Еда | 100 |\n| Итого | =SUM(B2:B2) |\n</file>';
  const f = makeFetch({ model: () => ({ json: messagesReply(answer) }) });
  await send(env, f, textUpdate(111, "сделай бюджет в экселе"));
  assert.equal(f.calls.model[0].body.max_tokens, 8000, "file requests get a bigger output budget");
  assert.equal(sentTexts(f.calls).at(-1), "Сделал таблицу.");
  const up = f.calls.uploads[0];
  assert.equal(up.method, "sendDocument");
  assert.equal(up.file.name, "Бюджет.xlsx");
  assert.equal(up.file.type, MIME.xlsx);
  assert.match(await openZip(up.file.bytes).text("xl/worksheets/sheet1.xml"), /SUM\(B2:B2\)/);
  const rec = JSON.parse(env.HISTORY.map.get(historyKey("111")));
  assert.match(rec.messages[1].content, /<file type="xlsx"/, "history remembers the file for later edits");
});

test("bot: /word exports the last answer without calling the model", async () => {
  const env = makeEnv();
  const f = makeFetch({ model: () => ({ json: messagesReply("## Решение\n**Ответ:** √36 = 6") }) });
  await send(env, f, textUpdate(111, "корень из 36"));
  await send(env, f, textUpdate(111, "/word"));
  assert.equal(f.calls.model.length, 1);
  const up = f.calls.uploads.at(-1);
  assert.equal(up.file.name, "Решение.docx");
  const { text } = await extractDocx(up.file.bytes);
  assert.match(text, /Решение[\s\S]*Ответ: √36 = 6/);
});

test("bot: unreadable office file gives a clear message", async () => {
  const env = makeEnv();
  const f = makeFetch({ files: { bad: new Uint8Array([1, 2, 3, 4]) } });
  await send(env, f, documentUpdate(111, "bad", "broken.docx", MIME.docx));
  assert.equal(f.calls.model.length, 0);
  assert.match(sentTexts(f.calls).at(-1), /Не получилось разобрать файл/);
});
