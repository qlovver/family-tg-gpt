import { test } from "node:test";
import assert from "node:assert/strict";
import { parseOps, extractEditBlocks, findReplaceInParagraph, setCell, applyEdits, docxForEdit, pptxForEdit } from "../src/edit.js";
import { buildDocx, buildXlsx, buildPptx, MIME } from "../src/build.js";
import { resolveTheme, parseColor } from "../src/theme.js";
import { openZip, createZip } from "../src/zip.js";
import { extractXlsx, extractDocx } from "../src/files.js";
import { createWorker, historyKey, fileKey, genKey } from "../src/bot.js";
import { makeEnv, makeFetch, messagesReply, textUpdate, documentUpdate, webhookRequest, execCtx, sentTexts } from "./helpers.js";

const send = (env, f, u) => createWorker({ fetchImpl: f.fetch }).fetch(webhookRequest(u), env, execCtx);

test("parseOps understands all commands (English and Russian)", () => {
  const ops = parseOps(`replace P3: новый текст
вставить после P4: абзац один
абзац два
delete P5-P7
find: опечатка => исправление
append: в самый конец
sheet: Продажи
set B5: 123
set 'Лист 2'!C3: =SUM(A1:A2)
clear D2:E3
delete slide 2
заменить S2P1: заголовок`);
  assert.deepEqual(ops.map(o => o.op), ["replace", "after", "delete", "find", "append", "sheet", "set", "set", "clear", "deleteSlide", "replace"]);
  assert.deepEqual(ops[1].lines, ["абзац один", "абзац два"]);
  assert.equal(ops[2].to, "P7");
  assert.equal(ops[7].sheet, "Лист 2");
  assert.equal(ops[8].to, "E3");
  assert.equal(ops[10].id, "S2P1");
});

test("find & replace works across runs with different formatting", () => {
  const p = '<w:p><w:r><w:t>пре</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>вет м</w:t></w:r><w:r><w:t>ир &amp; всё</w:t></w:r></w:p>';
  const r = findReplaceInParagraph(p, "превет мир", "привет, мир", "w");
  assert.equal(r.count, 1);
  assert.equal(r.xml.replace(/<[^>]+>/g, ""), "привет, мир &amp; всё");
  assert.match(r.xml, /<w:b\/>/, "formatting of runs is kept");
});

test("docx edits keep the rest of the document intact", async () => {
  const src = buildDocx("# Работа\n## Введение\nПервый абзац с ошибкой: превет.\n- пункт\n| A | B |\n|---|---|\n| 1 | 2 |\nКонец.");
  const view = await docxForEdit(src);
  const id = /\[(P\d+)\] Первый абзац/.exec(view)[1];
  const cell = /\[(P\d+)\] \(табл\.1\) 2/.exec(view)[1];
  const { bytes, report } = await applyEdits("docx", src, parseOps(`find: превет => привет
after ${id}: ## Новый раздел
Текст **нового** раздела.
replace ${cell}: 42
replace P999: нет такого`));
  assert.equal(report.applied, 3);
  assert.deepEqual(report.missing, ["P999"]);
  const { text } = await extractDocx(bytes);
  assert.match(text, /Первый абзац с ошибкой: привет\./);
  assert.match(text, /# Новый раздел\nТекст нового раздела\./, "inserted heading uses the document's heading style");
  assert.match(text, /\| 1 \| 42 \|/);
  assert.ok(openZip(bytes).names.includes("word/styles.xml"));
});

test("xlsx edits: change cells, add rows, formulas recalculated on open", async () => {
  const src = buildXlsx("## Продажи\n| Товар | Цена | Кол-во | Сумма |\n|---|---|---|---|\n| Чай | 100 | 3 | =B2*C2 |\n| Итого | | | =SUM(D2:D2) |");
  const { bytes, report } = await applyEdits("xlsx", src, parseOps("sheet: Продажи\nset C2: 5\nset A5: Кофе\nset B5: 250\nset D5: =B5*C5\nset D3: =SUM(D2:D2)+D5\nsheet: Нет такого"));
  assert.equal(report.applied, 5);
  assert.match(report.missing[0], /Нет такого/);
  const { text } = await extractXlsx(bytes);
  assert.match(text, /C2: 5/);
  assert.match(text, /A5: Кофе \| B5: 250 \| D5: =B5\*C5/);
  assert.match(await openZip(bytes).text("xl/workbook.xml"), /fullCalcOnLoad="1"/);
});

test("setCell keeps cell order and style, creates rows in order", () => {
  let x = '<sheetData><row r="1"><c r="A1" s="3"><v>1</v></c><c r="C1" s="3"><v>3</v></c></row><row r="4"><c r="A4"><v>4</v></c></row></sheetData>';
  x = setCell(x, "B1", "2");
  x = setCell(x, "A1", "текст");
  x = setCell(x, "A2", "=A1*2");
  assert.match(x, /<c r="A1" s="3" t="inlineStr"><is><t xml:space="preserve">текст<\/t><\/is><\/c><c r="B1" s="3"><v>2<\/v><\/c><c r="C1"/);
  assert.match(x, /<\/row><row r="2"><c r="A2"><f>A1\*2<\/f><\/c><\/row><row r="4">/);
});

test("pptx edits by slide paragraph address", async () => {
  const src = buildPptx("# Отчёт\nподзаголовок\n---\n# Итоги\n- Рост 10%\n- Клиенты: 5");
  const view = await pptxForEdit(src);
  const growth = /\[(S2P\d+)\] Рост 10%/.exec(view)[1];
  const { bytes, report } = await applyEdits("pptx", src, parseOps(`find: 10% => 12%\nafter ${growth}: Новый пункт\nreplace S1P1: Отчёт за год`));
  assert.equal(report.applied, 3);
  const after = await pptxForEdit(bytes);
  assert.match(after, /Отчёт за год/);
  assert.match(after, /Рост 12%\n\[S2P\d+\] Новый пункт/);
});

test("custom colours and fonts from the request", async () => {
  assert.equal(parseColor("красный"), "C62828");
  assert.equal(parseColor("#abc"), "AABBCC");
  const t = resolveTheme({ theme: "ocean", primary: "бордовый", font: "Arial", bg: "чёрный" });
  assert.equal(t.primary, "7F1D1D");
  assert.equal(t.font, "Arial");
  assert.equal(t.darkContent, true);
  const deck = buildPptx("# A\nb\n---\n# C\n- d", { primary: "#2E7D32", accent: "золотой" });
  assert.match(await openZip(deck).text("ppt/slides/slide2.xml"), /C9A227/);
  const doc = buildDocx("# T\n## H\ntext", { heading: "красный", font: "Arial", size: "13", spacing: "1.5", align: "justify" });
  const styles = await openZip(doc).text("word/styles.xml");
  assert.match(styles, /w:ascii="Arial"/);
  assert.match(styles, /<w:color w:val="C62828"\/>/);
  assert.match(styles, /w:line="360"/);
  assert.match(styles, /<w:jc w:val="both"\/>/);
  assert.match(styles, /<w:sz w:val="26"\/>/);
});

test("GOST Word style and page numbers", async () => {
  const doc = buildDocx("# Реферат\nВыполнила: А. И.\n@pagebreak\n@toc\n## Введение\nТекст.", { style: "gost" });
  const z = openZip(doc);
  const styles = await z.text("word/styles.xml");
  assert.match(styles, /Times New Roman/);
  assert.match(styles, /w:line="360"/);
  assert.match(styles, /w:firstLine="709"/);
  assert.match(await z.text("word/footer1.xml"), /PAGE/);
  const text = (await extractDocx(doc)).text;
  assert.match(text, /Содержание\nВведение/);
  assert.ok(!/@pagebreak/.test(text));
});

test("xlsx design: totals row, formats, charts", async () => {
  const bytes = buildXlsx("## Бюджет\n| Статья | Сумма, ₽ | Дата |\n|---|---|---|\n| Еда | 1500,5 | 05.03.2026 |\n| Итого | =SUM(B2:B2) | |\n@chart: pie; title: Структура", { theme: "forest" });
  const z = openZip(bytes);
  assert.ok(z.names.includes("xl/charts/chart1.xml") && z.names.includes("xl/drawings/drawing1.xml"));
  const chart = await z.text("xl/charts/chart1.xml");
  assert.match(chart, /<c:pieChart>/);
  assert.match(chart, /\$A\$2:\$A\$2</, "totals row excluded from the chart");
  const styles = await z.text("xl/styles.xml");
  assert.match(styles, /&quot;₽&quot;/);
  assert.match(styles, /dd\.mm\.yyyy/);
  assert.match(await z.text("xl/worksheets/sheet1.xml"), /<c r="C2" s="\d+"><v>46086<\/v>/);
});

test("bot: user's Word file is edited in place and sent back with the same name", async () => {
  const env = makeEnv();
  const src = buildDocx("# Отчёт\nТекст с ошибкой: превет.");
  let call = 0;
  const f = makeFetch({
    files: { d1: src },
    model: () => ({ json: messagesReply(++call === 1
      ? 'Исправил опечатку.\n<edit file="отчёт.docx">\nfind: превет => привет\n</edit>'
      : 'Дописал.\n<edit file="отчёт.docx">\nappend: Новый абзац.\n</edit>') }),
  });
  await send(env, f, documentUpdate(111, "d1", "отчёт.docx", MIME.docx, "исправь ошибки"));
  const up = f.calls.uploads[0];
  assert.equal(up.file.name, "отчёт.docx");
  assert.match(up.fields.caption, /Правок внесено: 1/);
  assert.match((await extractDocx(up.file.bytes)).text, /Текст с ошибкой: привет\./);
  assert.equal(sentTexts(f.calls).at(-1), "Исправил опечатку.");
  const last = JSON.parse(env.HISTORY.map.get(fileKey("111")));
  assert.equal(last.fileId, "up1", "next edits apply to the edited version");

  await send(env, f, textUpdate(111, "добавь в конец абзац"));
  assert.match(f.calls.model[1].body.messages.at(-1).content.at(-1).text, /привет\./, "model sees the edited text");
  const second = (await extractDocx(f.calls.uploads[1].file.bytes)).text;
  assert.match(second, /привет\.[\s\S]*Новый абзац\./, "second edit builds on the first");
});

test("bot: a file the bot made can be reworked on request", async () => {
  const env = makeEnv();
  let call = 0;
  const f = makeFetch({ model: () => ({ json: messagesReply(++call === 1
    ? 'Готово.\n<file type="pptx" name="Доклад" theme="ocean">\n# Доклад\nавтор\n---\n# Слайд\n- пункт\n</file>'
    : 'Сделал красным.\n<file type="pptx" name="Доклад" theme="ocean" primary="красный">\n# Доклад\nавтор\n---\n# Слайд\n- пункт\n</file>') }) });
  await send(env, f, textUpdate(111, "сделай презентацию"));
  assert.ok(env.HISTORY.map.has(genKey("111")));
  await send(env, f, textUpdate(111, "поменяй цвет на красный"));
  const prompt = f.calls.model[1].body.messages.at(-1).content.at(-1).text;
  assert.match(prompt, /Последний файл, который ты создал: «Доклад.pptx»[\s\S]*<file type="pptx" name="Доклад.pptx" theme="ocean">[\s\S]*# Слайд/);
  assert.equal(f.calls.model[1].body.max_tokens, 8000);
  assert.equal(f.calls.uploads.length, 2);
  assert.match(await openZip(f.calls.uploads[1].file.bytes).text("ppt/slides/slide1.xml"), /C62828|631414/);
  // plain chat does not drag the file source along
  await send(env, f, textUpdate(111, "как дела?"));
  assert.ok(!/Последний файл/.test(f.calls.model[2].body.messages.at(-1).content.at(-1).text));
});
