// Telegram bot logic for Cloudflare Workers.
// Secrets:   TELEGRAM_BOT_TOKEN, TOOKEN_API_KEY, WEBHOOK_SECRET, SETUP_KEY
// Variable:  ALLOWED_USER_IDS="123,456" (a Secret also works)
// KV:        HISTORY (namespace kv-get-started)
// Optional:  MODEL, API_FORMAT (messages|chat|responses), MODEL_STREAM, MAX_OUTPUT_TOKENS, TOOKEN_BASE_URL

import { toTelegramHtml, splitHtml, htmlToPlain, escapeHtml } from "./format.js";
import { askModel, ModelError, modelConfig, listModels, DEFAULT_MODEL, probeCapabilities, probeAuthenticity } from "./model.js";
import { classifyDocument, extractDocument, base64FromBytes, MAX_FILE_SIZE, MAX_PDF_SIZE, MAX_EXTRACT_CHARS } from "./files.js";
import { extractFileBlocks, buildFile, buildDocx, safeFileName, MIME } from "./build.js";
import { extractEditBlocks, applyEdits, docxForEdit, pptxForEdit, editKindOf } from "./edit.js";

export const MAX_IMAGE_SIZE = 2_500_000;      // bytes, after Telegram compression
export const MAX_IMAGE_SIDE = 1600;           // px; Claude downsizes larger images anyway
export const MAX_HISTORY = 12;                // messages (6 question/answer pairs) per user
export const MAX_HISTORY_CHARS = 20_000;      // total history budget, keeps token cost bounded
export const MAX_USER_CHARS = 6_000;
export const MAX_STORED_ANSWER = 6_000;
export const STALE_PENDING_MS = 150_000;      // longer than any possible run of one update
export const MODEL_TIMEOUT_MS = 85_000;       // Telegram wait + 30 s of waitUntil after disconnect
const TG_LIMIT = 3900;                        // < 4096 visible chars per message

export const SYSTEM_PROMPT = `Ты — дружелюбный и внимательный ИИ-помощник в Telegram. Тобой пользуются два человека, у каждого своя отдельная переписка. Отвечай по-русски (если собеседник пишет на другом языке — на его языке), естественно и по делу.

Главная задача — помощь с учёбой: математика, экономика, физика, языки и другие предметы. Но ты универсальный помощник: в обычном разговоре отвечай свободно, без формата «решения задачи».

Когда присылают задание (фото или текст):
1. Кратко перепиши условие своими словами (это нужно, чтобы потом можно было обсуждать задачу дальше).
2. Выпиши, что дано и что нужно найти.
3. Назови формулу или метод и почему он подходит.
4. Реши по шагам, каждый шаг — отдельный абзац с короткой подписью.
5. Проверь вычисления (подстановкой, оценкой или другим способом).
6. Выдели окончательный ответ отдельной строкой: **Ответ: ...**
7. Одним-двумя предложениями объясни смысл результата.
Если в задании несколько пунктов — реши все по порядку, каждый с заголовком.
Если на фото что-то не читается или обрезано — не угадывай: назови конкретный фрагмент и попроси переснять его. Не придумывай недостающие условия; если без допущения не обойтись, явно его оговори.

Оформление (сообщение показывается в Telegram, который НЕ отображает LaTeX):
- НЕ используй LaTeX: никаких $, \\frac, \\sqrt, \\cdot, \\dot, \\boxed и т.п.
- Пиши формулы обычным текстом с Unicode: √k, x², x³, xₜ, k̇, ½, a/b, (a + b)/(c − d), ·, ×, ÷, −, ±, ≤, ≥, ≠, ≈, ∞, π, Δ, α, β, δ, ∑, ∫, →, ⇒, °.
- Каждая формула или вычисление — на отдельной строке. Между шагами — пустая строка.
- Можно использовать простую разметку: **жирный**, *курсив*, заголовки через ##, списки через «- » или «1. », \`код\`. Таблицы — только маленькие.
- Не злоупотребляй жирным: выделяй подписи шагов и итоговый ответ.
- Десятичные числа пиши с запятой (0,5), если собеседник сам не использует точку.

Файлы, которые присылают:
- Содержимое Word, Excel, PowerPoint и текстовых файлов приходит в теге <document>; PDF приходит вложением. Картинки внутри Word/PowerPoint тебе не видны — если без них не понять задание, скажи об этом.
- Отвечай по содержимому файла; ссылайся на разделы, страницы, слайды или ячейки, где это помогает.

Создание файлов (Word, Excel, PowerPoint). Делай файл ТОЛЬКО когда об этом просят («в ворде», «документом», «таблицей в экселе», «презентацию» и т.п.) или когда без файла никак. Файл оформляется блоком:
<file type="docx" name="Название.docx">
...содержимое в Markdown...
</file>
Правила по типам:
- docx: документ оформляется автоматически. Первая строка «# Название» — титул, 1–2 строки под ним — подзаголовок (автор, предмет, дата). Дальше разделы «## », подразделы «### », абзацы, списки «- » и «1. », таблицы Markdown, **жирный**, *курсив*, цитаты/важные замечания строкой «> ». Строка «**Ответ:** …» (или «Итог:», «Вывод:») выделяется рамкой. Формулы — отдельной строкой, они встанут по центру. «@toc» — содержание, «@pagebreak» — новая страница.
  Оформление: для рефератов, докладов, курсовых и всего, где нужны требования вуза/школы, ставь style="gost" (Times New Roman 14, интервал 1,5, поля 30/15/20/20, по ширине, номера страниц): <file type="docx" name="Реферат.docx" style="gost">. Иначе — цветная тема theme (те же, что для презентаций: ocean, sunset, forest, lavender, minimal, classic). Можно задать font="..." size="...". Пиши полный текст документа, а не план.
- xlsx: таблица оформляется автоматически (цветная шапка, полосы, рамки, фильтр, закреплённая шапка). Укажи theme (как у презентаций). Каждый лист начинается строкой «## Название листа», дальше таблица Markdown; первая строка — заголовки. В ячейках — чистые числа без единиц: единицы пиши в заголовке («Сумма, ₽» — тогда числа станут рублями), проценты как 15%, даты как 15.03.2026. Строка, начинающаяся с «Итого»/«Всего», выделяется как итоговая. Формулы начинаются с «=» и пишутся по-английски через запятую: =SUM(B2:B10), =AVERAGE(C2:C5), =B2*C2, =IF(B2>0,"да","нет"); первая строка данных — строка 2.
  Диаграмма: строка после таблицы «@chart: column» (или bar, line, pie); можно уточнить «@chart: line; columns: B,C; title: Динамика». Категории берутся из столбца A, строка «Итого» в диаграмму не попадает. Добавляй диаграмму, когда данные просят визуализации (динамика, структура, сравнение).
- pptx: презентация оформляется автоматически в красивом дизайне. Укажи тему атрибутом theme: ocean (синяя деловая), sunset (тёплая), forest (зелёная), lavender (нежная фиолетовая), minimal (строгая бело-красная), night (тёмная), classic (классика с золотом, для истории и литературы). Выбери тему под содержание или по просьбе пользователя.
  Слайды разделяй строкой «---». Каждый слайд — «# Заголовок», затем содержимое. Тип слайда можно задать строкой «@layout: тип» перед заголовком; если не задан, он определяется сам:
  · первый слайд — титульный: заголовок + строка подзаголовка (автор, предмет);
  · section — разделитель частей: только заголовок;
  · bullets — 3–6 коротких пунктов «- » (можно **выделять** ключевые слова);
  · two-columns — сравнение: два блока «## Название» с пунктами;
  · cards — 3–6 блоков «## Название» + 1–2 предложения (факторы, причины, направления);
  · steps — этапы/алгоритм: нумерованный список «1. Название: пояснение» (до 8 шагов);
  · big-number — ключевая цифра: заголовок — само число («73%», «1,5 млн ₽»), ниже подпись;
  · quote — цитата: строка «> текст» и «— Автор»;
  · table — таблица Markdown;
  · ending — последний слайд «Спасибо за внимание!».
  Чередуй типы слайдов, чтобы презентация была живой. Текст на слайде — коротко, тезисами; подробности — не больше 1–2 предложений. Обычно 8–12 слайдов, если не просили другое. Картинок и фото вставить нельзя.
- Внутри файла тоже никакого LaTeX — только Unicode.
- Вне блока напиши одно-два предложения о том, что в файле. Не дублируй содержимое файла в сообщении.
- PDF создать не можешь — предложи Word (на телефоне его можно сохранить как PDF через «Поделиться → Напечатать»).

Оформление по просьбе пользователя — выполняй его буквально, атрибутами блока <file> (для всех трёх типов):
primary="…" (основной цвет: заголовки, шапки таблиц, фон титульного слайда), accent="…" (акцент: линии, маркеры, выделения), bg="…" (фон слайдов; тёмный фон → светлый текст), text="…" (цвет текста), heading="…" (цвет заголовков), font="…" (шрифт текста), headfont="…" (шрифт заголовков), size="…" (размер текста в Word), spacing="1.5" (межстрочный интервал в Word), align="justify|left|center" (выравнивание в Word), indent="1.25" (абзацный отступ в см). Цвета — HEX (#C62828) или словами (красный, тёмно-синий, бордовый, мятный, золотой…). Просьбу «красные заголовки, шрифт Arial 12» переводи в heading="красный" font="Arial" size="12". Свои атрибуты сочетаются с theme и style="gost".

Доработка созданного тобой файла: если просят что-то поменять в файле, который ты сделал раньше (тебе будет показан его полный исходник), пришли новый полный блок <file> с тем же type и name и учти просьбу (цвета, слайды, строки, текст…). Не пиши «я изменил», не прислав файл.

Правка ПРИСЛАННОГО пользователем Word/Excel/PowerPoint с сохранением его оформления: если просят исправить, дописать, удалить, поменять текст или цифры в их файле — НЕ создавай новый файл, а верни блок правок:
<edit file="имя файла">
replace P12: новый текст абзаца
after P12: новый абзац после P12 (можно «## Заголовок», «- пункт списка», **жирный**; несколько строк — несколько абзацев)
before P3: абзац перед P3
delete P15            (или диапазон: delete P15-P18)
find: старый текст => новый текст      (замена во всём файле: опечатки, имена, числа)
append: абзац в конец документа
</edit>
Для PowerPoint адреса вида S2P3 (слайд 2, абзац 3), те же команды, плюс «delete slide 4». Для Excel: «sheet: Имя листа», затем «set B5: 123», «set D7: =SUM(D2:D6)», «set A8: текст», «clear B5» или «clear B5:D9»; можно «set 'Лист 2'!C3: 10». Новые строки — просто записывай ячейки в пустые строки ниже.
Адреса бери только из присланного документа ([P…], [S…P…], адреса ячеек). Мелкие правки — find/replace; переписать абзац — replace. Вне блока коротко перечисли, что изменил. Если просят полностью переделать или красиво переоформить их файл — тогда создай новый файл блоком <file>, взяв содержание из присланного.

Не представляйся и не называй свою модель, пока об этом не спросят. Не выдумывай факты; если не уверен — скажи об этом.`;

/* ============================== helpers ============================== */

function json(data, status = 200) {
  return new Response(JSON.stringify(data, null, 2), {
    status,
    headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" },
  });
}

function html(body, status = 200) {
  return new Response(body, {
    status,
    headers: {
      "content-type": "text/html; charset=utf-8",
      "cache-control": "no-store",
      "referrer-policy": "no-referrer",
      "x-robots-tag": "noindex",
    },
  });
}

function safeEqual(a, b) {
  a = String(a ?? ""); b = String(b ?? "");
  if (!a || !b) return false;
  let diff = a.length ^ b.length;
  for (let i = 0; i < Math.max(a.length, b.length); i++) diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return diff === 0;
}

export function allowedIds(env) {
  return String(env.ALLOWED_USER_IDS || "").split(/[\s,;]+/).map(s => s.trim()).filter(s => /^\d+$/.test(s));
}

function log(event, data = {}) {
  console.log(JSON.stringify({ event, ...data }));
}

export class TelegramError extends Error {
  constructor(method, status, description) {
    super(`Telegram ${method} failed (${status}): ${description}`);
    this.name = "TelegramError";
    this.method = method;
    this.status = status;
    this.description = description;
  }
}

// TELEGRAM_API_URL is only for local tests against a mock server.
const tgBase = env => String(env.TELEGRAM_API_URL || "https://api.telegram.org").replace(/\/+$/, "");

export function makeTelegram(env, fetchImpl) {
  async function telegram(method, payload) {
    const res = await fetchImpl(`${tgBase(env)}/bot${env.TELEGRAM_BOT_TOKEN}/${method}`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(payload),
    });
    let data = {};
    try { data = await res.json(); } catch { /* ignore */ }
    if (!res.ok || !data.ok) throw new TelegramError(method, res.status, String(data.description || "unknown").slice(0, 200));
    return data.result;
  }
  // multipart/form-data upload (sendDocument)
  telegram.upload = async function upload(method, fields, file) {
    const form = new FormData();
    for (const [k, v] of Object.entries(fields)) form.append(k, typeof v === "string" ? v : JSON.stringify(v));
    form.append(file.field, new Blob([file.bytes], { type: file.mime }), file.name);
    const res = await fetchImpl(`${tgBase(env)}/bot${env.TELEGRAM_BOT_TOKEN}/${method}`, { method: "POST", body: form });
    let data = {};
    try { data = await res.json(); } catch { /* ignore */ }
    if (!res.ok || !data.ok) throw new TelegramError(method, res.status, String(data.description || "unknown").slice(0, 200));
    return data.result;
  };
  return telegram;
}

/**
 * Sends model text as formatted HTML, split into safe chunks.
 * Falls back to plain text for any chunk Telegram refuses to parse.
 */
export async function sendFormatted(telegram, chatId, text, replyTo) {
  const chunks = splitHtml(toTelegramHtml(text), TG_LIMIT);
  if (!chunks.length) chunks.push("…");
  for (let k = 0; k < chunks.length; k++) {
    const base = { chat_id: chatId, link_preview_options: { is_disabled: true } };
    if (k === 0 && replyTo) base.reply_parameters = { message_id: replyTo, allow_sending_without_reply: true };
    try {
      await telegram("sendMessage", { ...base, text: chunks[k], parse_mode: "HTML" });
    } catch (err) {
      if (!(err instanceof TelegramError) || err.status !== 400) throw err;
      log("html_fallback", { description: err.description });
      const plain = htmlToPlain(chunks[k]);
      for (let i = 0; i < plain.length; i += 4000) {
        await telegram("sendMessage", { ...base, text: plain.slice(i, i + 4000) });
      }
    }
  }
  return chunks.length;
}

async function sendHtml(telegram, chatId, htmlText) {
  return telegram("sendMessage", { chat_id: chatId, text: htmlText, parse_mode: "HTML", link_preview_options: { is_disabled: true } });
}

export function imageToBase64(arrayBuffer) {
  return base64FromBytes(new Uint8Array(arrayBuffer));
}

export function detectMime(bytes, fallback = "image/jpeg") {
  const b = new Uint8Array(bytes.slice(0, 12));
  if (b[0] === 0xff && b[1] === 0xd8) return "image/jpeg";
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return "image/png";
  if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return "image/gif";
  if (b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46 && b[8] === 0x57 && b[9] === 0x45) return "image/webp";
  return fallback;
}

// Picks the best photo size: as large as useful, but not wastefully large.
export function pickPhoto(sizes) {
  const list = [...(sizes || [])].filter(p => p && p.file_id)
    .sort((a, b) => (a.width * a.height) - (b.width * b.height));
  if (!list.length) return null;
  const fits = list.filter(p => Math.max(p.width, p.height) <= MAX_IMAGE_SIDE && (!p.file_size || p.file_size <= MAX_IMAGE_SIZE));
  if (fits.length) return fits[fits.length - 1];
  const small = list.filter(p => !p.file_size || p.file_size <= MAX_IMAGE_SIZE);
  return small.length ? small[small.length - 1] : null;
}

const IMAGE_MIMES = ["image/jpeg", "image/png", "image/webp", "image/gif"];

// What a message carries besides text: a photo/image, or a document file.
function attachmentOf(m) {
  if (m.photo?.length) {
    const p = pickPhoto(m.photo);
    if (!p) return { error: "image_too_large" };
    return { image: { fileId: p.file_id, mime: "image/jpeg" } };
  }
  if (m.document) {
    const c = classifyDocument(m.document);
    if (c.error) return { error: c.error, hint: c.hint };
    if (c.kind === "image") {
      if (m.document.file_size && m.document.file_size > MAX_IMAGE_SIZE) return { error: "image_too_large" };
      return { image: { fileId: m.document.file_id, mime: IMAGE_MIMES.includes(c.mime) ? c.mime : "image/jpeg" } };
    }
    return { file: { fileId: m.document.file_id, name: c.name, kind: c.kind } };
  }
  return {};
}

// Extracts what we need to call the model from a Telegram message.
export function extractRequest(msg) {
  const text = (msg.text || msg.caption || "").trim();
  const own = attachmentOf(msg);
  const req = { text: text.slice(0, MAX_USER_CHARS), image: own.image || null, file: own.file || null, error: own.error || null, hint: own.hint, quoted: "" };
  const r = msg.reply_to_message;
  if (r) {
    // Replying to an earlier photo/file with a question re-uses it.
    if (!req.image && !req.file && !req.error) {
      const fromReply = attachmentOf(r);
      if (fromReply.image) req.image = { ...fromReply.image, fromReply: true };
      if (fromReply.file) req.file = { ...fromReply.file, fromReply: true };
    }
    const rt = (r.text || r.caption || "").trim();
    if (rt) req.quoted = rt.slice(0, 1200);
  }
  return req;
}

/* ============================== history ============================== */

export function historyKey(uid) {
  return `history:${uid}`;
}
export function genKey(uid) {
  return `lastgen:${uid}`;
}
const LAST_GEN_TTL = 7 * 24 * 3600; // a file the bot made can be reworked for a week
export function fileKey(uid) {
  return `lastfile:${uid}`;
}
const LAST_FILE_TTL = 6 * 3600; // a sent document stays in context for 6 hours
const MAX_LAST_FILE_CHARS = 60_000;

export async function loadRecord(kv, uid) {
  const raw = await kv.get(historyKey(uid), "json");
  if (Array.isArray(raw)) return { v: 2, messages: raw };          // old format: plain array
  if (raw && typeof raw === "object" && Array.isArray(raw.messages)) return raw;
  return { v: 2, messages: [] };
}

export async function saveRecord(kv, uid, record) {
  await kv.put(historyKey(uid), JSON.stringify(record));
}

export function trimHistory(messages) {
  let out = messages.slice(-MAX_HISTORY);
  const size = arr => arr.reduce((n, m) => n + String(m.content || "").length, 0);
  while (out.length > 2 && size(out) > MAX_HISTORY_CHARS) out = out.slice(2);
  while (out.length && out[0].role !== "user") out = out.slice(1);
  return out;
}

/* ============================== models ============================== */

// Used when Tooken doesn't return its model list. IDs as in Tooken's catalogue.
export const FALLBACK_MODELS = [
  "claude-opus-5-5", "claude-sonnet-5-5", "claude-haiku-5-5", "claude-fable-5-1",
  "gpt-6.1-sol", "gpt-6.1-terra", "gpt-6-luna", "gpt-6-astra",
  "gemini-4-argon", "gemini-3.8-flash", "deepseek-v4-pro", "grok-4.7", "kimi-k3.1-pro", "glm-5.3",
];
const MODELS_KEY = "models:list";

export function prettyModel(id) {
  const s = String(id);
  const brand = { claude: "Claude", gpt: "GPT", gemini: "Gemini", deepseek: "DeepSeek", grok: "Grok", kimi: "Kimi", glm: "GLM", qwen: "Qwen", llama: "Llama", mistral: "Mistral" };
  const parts = s.split("-");
  const head = brand[parts[0].toLowerCase()] || parts[0];
  const rest = parts.slice(1);
  // join version digits: "5","5" -> "5.5"
  const out = [];
  for (const p of rest) {
    const prev = out[out.length - 1];
    if (/^\d{1,2}$/.test(p) && prev && /\d$/.test(prev) && !/\./.test(prev)) out[out.length - 1] += "." + p;
    else out.push(/^\d/.test(p) ? p : p.charAt(0).toUpperCase() + p.slice(1));
  }
  // GPT-6.1 style: brand and version joined with a dash
  if (head === "GPT" && out.length && /^\d/.test(out[0])) return `GPT-${out[0]} ${out.slice(1).join(" ")}`.trim();
  return [head, ...out].join(" ");
}

async function availableModels(ctx) {
  const { kv, env, fetchImpl } = ctx;
  const cached = kv ? await kv.get(MODELS_KEY, "json").catch(() => null) : null;
  if (cached?.length) return cached;
  const live = await listModels(env, { fetchImpl });
  const list = live || FALLBACK_MODELS;
  if (live && kv) await kv.put(MODELS_KEY, JSON.stringify(list), { expirationTtl: 6 * 3600 }).catch(() => {});
  return list;
}

function modelKeyboard(models, current) {
  const sorted = [...models].sort((a, b) => {
    const rank = id => ["claude", "gpt", "gemini", "deepseek", "grok"].findIndex(p => id.toLowerCase().startsWith(p));
    const ra = rank(a), rb = rank(b);
    return (ra < 0 ? 99 : ra) - (rb < 0 ? 99 : rb) || a.localeCompare(b);
  });
  const buttons = sorted.filter(id => `m:${id}`.length <= 64).map(id => ({ text: `${id === current ? "✅ " : ""}${prettyModel(id)}`, callback_data: `m:${id}` }));
  const rows = [];
  for (let i = 0; i < buttons.length; i += 2) rows.push(buttons.slice(i, i + 2));
  return { inline_keyboard: rows };
}

export function currentModel(rec, env) {
  return rec?.model || env.MODEL || DEFAULT_MODEL;
}

// Make sure Telegram sends button presses (older webhook setups only had "message").
async function ensureCallbacks(ctx) {
  try {
    const info = await ctx.telegram("getWebhookInfo", {});
    if (info?.url && !(info.allowed_updates || []).includes("callback_query") && ctx.env.WEBHOOK_SECRET) {
      await ctx.telegram("setWebhook", { url: info.url, secret_token: ctx.env.WEBHOOK_SECRET, max_connections: info.max_connections || 1, allowed_updates: ["message", "callback_query"] });
      log("webhook_updated", { allowed_updates: ["message", "callback_query"] });
    }
  } catch (err) {
    log("webhook_check_failed", { message: String(err?.message || "").slice(0, 120) });
  }
}

async function showModelMenu(ctx, uid, chatId) {
  const [rec, models] = await Promise.all([loadRecord(ctx.kv, uid), availableModels(ctx)]);
  await ensureCallbacks(ctx);
  const cur = currentModel(rec, ctx.env);
  await ctx.telegram("sendMessage", {
    chat_id: chatId, parse_mode: "HTML",
    text: `🧠 <b>Выбор модели</b>\nСейчас: <b>${escapeHtml(prettyModel(cur))}</b> (<code>${escapeHtml(cur)}</code>)\n\nВыбор действует только для тебя. История диалога сохраняется.\n<i>Названия моделей даёт Tooken Club; какая модель отвечает на самом деле, проверить нельзя.</i>`,
    reply_markup: modelKeyboard(models, cur),
  });
}

async function setUserModel(ctx, uid, model) {
  const rec = await loadRecord(ctx.kv, uid);
  rec.model = model;
  await saveRecord(ctx.kv, uid, rec);
}

async function handleCallback(cq, ctx) {
  const { env, telegram, kv } = ctx;
  const uid = String(cq.from?.id || "");
  const answer = text => telegram("answerCallbackQuery", { callback_query_id: cq.id, text, show_alert: false }).catch(() => {});
  if (!allowedIds(env).includes(uid)) { await answer("🔒 Это приватный бот"); return "ok"; }
  const data = String(cq.data || "");
  if (!data.startsWith("m:") || !kv) { await answer(""); return "ok"; }
  const model = data.slice(2);
  const models = await availableModels(ctx);
  if (!models.includes(model)) { await answer("Этой модели больше нет в списке"); return "ok"; }
  await setUserModel(ctx, uid, model);
  await answer(`Модель: ${prettyModel(model)}`);
  if (cq.message) {
    await telegram("editMessageText", {
      chat_id: cq.message.chat.id, message_id: cq.message.message_id, parse_mode: "HTML",
      text: `🧠 <b>Модель выбрана:</b> ${escapeHtml(prettyModel(model))} (<code>${escapeHtml(model)}</code>)\nМожно продолжать диалог. Сменить — /model`,
    }).catch(() => {});
  }
  log("model_selected", { uid, model });
  return "ok";
}

/* ============================== texts ============================== */

const HELP = `<b>Привет! Я ИИ-помощник 🤖</b>

Пиши вопрос текстом или присылай <b>фото задания</b> — разберу и решу по шагам.

<b>Файлы</b>
• Читаю: PDF, Word, Excel, PowerPoint, текст и код, картинки.
• Создаю: Word, Excel, PowerPoint — просто попроси: «оформи в ворде», «сделай таблицу в экселе», «сделай презентацию на 7 слайдов».
• Присланный файл я помню 6 часов — можно задавать вопросы по нему.

<b>Подсказки</b>
• Одно фото — одно задание: так ответ точнее.
• Чтобы спросить про старое фото или файл, ответь (reply) на него.
• Если ответ оборвался — напиши «продолжи».

<b>Команды</b>
/model — выбрать модель ИИ
/agent — сервер-агент (красивые файлы, графики, PDF)
/check — проверить модель (инструменты, зрение, подлинность)
/word — прислать мой последний ответ файлом Word
/reset — очистить историю и забыть файл (только свои)
/retry — повторить последний неудачный запрос
/id — узнать свой Telegram ID`;

export function errorText(code, hint) {
  switch (code) {
    case "timeout":
    case "upstream_timeout":
      return "⏳ Модель думала слишком долго, и я не дождался ответа. Отправь /retry, чтобы попробовать ещё раз, или раздели задание на части.";
    case "no_balance":
      return "💳 На счёте Tooken Club закончились средства. Пополни баланс и отправь /retry.";
    case "rate_limited":
      return "🚦 Слишком много запросов подряд. Подожди минуту и отправь /retry.";
    case "auth_failed":
      return "🔑 Tooken Club не принял API-ключ. Нужно проверить секрет TOOKEN_API_KEY в Cloudflare.";
    case "model_not_found":
      return "❓ Tooken Club не нашёл выбранную модель. Выбери другую: /model";
    case "image_unsupported":
      return "🖼 Эта модель, похоже, не понимает фото или файлы. Выбери другую (например, Claude или GPT): /model";
    case "endpoint_not_found":
    case "bad_request":
      return "⚠️ Сервис ИИ отклонил запрос (неподходящий формат). Отправь /retry; если повторится — нужна проверка настроек бота.";
    case "pdf_rejected":
      return "📄 Сервис ИИ не принял этот PDF. Пришли нужные страницы скриншотами или файл Word.";
    case "context_too_long":
    case "request_too_large":
      return "📚 Запрос получился слишком большим. Отправь /reset или пришли файл поменьше (часть страниц).";
    case "image_too_large":
      return "🖼 Фото слишком большое. Отправь его как обычное фото (не файлом), до 2,5 МБ.";
    case "image_download_failed":
    case "file_download_failed":
      return "📥 Не получилось скачать файл с серверов Telegram. Отправь его ещё раз.";
    case "file_too_large":
      return "📦 Файл слишком большой (Telegram отдаёт ботам файлы до 20 МБ). Пришли файл поменьше или нужную часть.";
    case "pdf_too_large":
      return "📄 PDF больше 10 МБ — такой я не потяну. Пришли нужные страницы отдельным PDF или скриншотами.";
    case "legacy_format":
      return `📎 Это старый формат, его я не читаю. Открой файл и сохрани как .${hint || "docx"} — тогда пришли снова.`;
    case "heic_unsupported":
      return "🖼 Формат HEIC я не читаю. Отправь фото обычным способом (не файлом) — Telegram сам переведёт его в JPG.";
    case "file_parse_failed":
      return "📎 Не получилось разобрать файл — возможно, он повреждён или защищён паролем.";
    case "file_unsupported":
      return "📎 Такой тип файла я не читаю. Подходят: PDF, Word, Excel, PowerPoint, текст/код и картинки.";
    case "no_storage":
      return "🛠 Бот не настроен: нет хранилища HISTORY. Владельцу нужно проверить KV binding в Cloudflare.";
    case "no_api_key":
      return "🛠 Бот не настроен: нет ключа TOOKEN_API_KEY.";
    default:
      return "😕 Не удалось получить ответ от ИИ. Отправь /retry чуть позже. Если ошибка повторяется — проверь баланс Tooken Club.";
  }
}

/* ============================== bot core ============================== */

async function downloadFile(telegram, env, fetchImpl, fileId, maxSize, tooLargeCode) {
  let f;
  try { f = await telegram("getFile", { file_id: fileId }); }
  catch (err) {
    if (err instanceof TelegramError && /too big/i.test(err.description)) throw new ModelError(tooLargeCode);
    throw new ModelError("file_download_failed");
  }
  if (!f?.file_path) throw new ModelError("file_download_failed");
  if (f.file_size && f.file_size > maxSize) throw new ModelError(tooLargeCode);
  const res = await fetchImpl(`${tgBase(env)}/file/bot${env.TELEGRAM_BOT_TOKEN}/${f.file_path}`);
  if (!res.ok) throw new ModelError("file_download_failed");
  const data = await res.arrayBuffer();
  if (data.byteLength > maxSize) throw new ModelError(tooLargeCode);
  return data;
}

async function getImage(telegram, env, fetchImpl, fileId, mimeHint) {
  const data = await downloadFile(telegram, env, fetchImpl, fileId, MAX_IMAGE_SIZE, "image_too_large");
  return { mime: detectMime(data, mimeHint), base64: imageToBase64(data) };
}

// Loads a document: PDFs become a model attachment, others become text.
// Text the model sees. Word and PowerPoint get [P12]/[S2P3] addresses for edits.
async function describeDocument(kind, name, data) {
  if (kind === "docx" || kind === "pptx") {
    let text = kind === "docx" ? await docxForEdit(data) : await pptxForEdit(data);
    const truncated = text.length > MAX_EXTRACT_CHARS;
    if (truncated) text = text.slice(0, MAX_EXTRACT_CHARS) + "\n[… файл обрезан …]";
    const note = kind === "docx"
      ? "Абзацы пронумерованы [P…] — это адреса для правок; «#» — заголовок, «- » — пункт списка. Картинки не показаны, но в файле сохраняются."
      : "Абзацы слайдов пронумерованы [S<слайд>P<абзац>] — это адреса для правок.";
    return { name, kind, text, note, truncated };
  }
  const r = await extractDocument(kind, name, data);
  return { name, kind, text: r.text, note: r.note, truncated: r.truncated };
}

async function loadDocument(ctx, file) {
  const { telegram, env, fetchImpl } = ctx;
  if (file.kind === "pdf") {
    const data = await downloadFile(telegram, env, fetchImpl, file.fileId, MAX_PDF_SIZE, "pdf_too_large");
    return { name: file.name, kind: "pdf", attachment: { mime: "application/pdf", base64: base64FromBytes(data), name: file.name } };
  }
  const data = await downloadFile(telegram, env, fetchImpl, file.fileId, MAX_FILE_SIZE, "file_too_large");
  try {
    return await describeDocument(file.kind, file.name, data);
  } catch (err) {
    throw new ModelError(err.message === "file_too_large" ? "file_too_large" : "file_parse_failed");
  }
}

const EDIT_WORDS = /измени|поменя|исправ|поправ|замени|добав|допиш|убер|удали|сократ|расшир|перепиш|переделай|отредакт|сделай (?:его|её|ее)|в (?:этом|моём|моем) файле|цвет|шрифт|фон|оформ|стил|крупнее|меньше|больше слайд/i;
const FILE_WORDS = /ворд|word|docx|документ|реферат|доклад|эссе|сочинени|отч[её]т|эксел|excel|xlsx|таблиц|презентац|слайд|pptx|powerpoint|файл/i;

function documentPrompt(doc, followUp) {
  const head = followUp
    ? `(Ранее в этом диалоге прислан файл «${doc.name}». Его содержимое ниже — используй его, если вопрос к нему относится.)`
    : `Пользователь прислал файл «${doc.name}».${doc.note ? " " + doc.note : ""}${doc.truncated ? " Файл большой — передана только первая часть." : ""}`;
  if (doc.kind === "pdf") return head + " PDF приложен к сообщению.";
  return `${head}\n<document name="${doc.name.replace(/"/g, "'")}">\n${doc.text}\n</document>`;
}

function userPrompt(req) {
  let text = req.text;
  if (!text && req.image) text = "Реши задание на фото подробно по шагам, проверь вычисления и выдели ответ.";
  if (!text && req.file) text = "Посмотри файл. Если это задание — реши его по шагам; если нет — кратко расскажи, что в нём, и спроси, что с ним сделать.";
  if (req.quoted && !req.image?.fromReply && !req.file?.fromReply) text = `(Это ответ на сообщение: «${req.quoted}»)\n\n${text}`;
  return text;
}

function historyUserEntry(req) {
  let mark = "";
  if (req.image) mark = req.image.fromReply ? "[Вопрос по ранее присланному фото] " : "[Фото задания] ";
  if (req.file) mark = `[Файл «${req.file.name}»] `;
  return (mark + (req.text || (req.image || req.file ? "(без подписи)" : ""))).slice(0, MAX_USER_CHARS);
}

// Shortens file blocks so the history keeps their gist without the full size.
function answerForHistory(answer) {
  return answer.replace(/(<file\b[^>]*>)([\s\S]*?)(<\/file>|$)/gi, (_, open, body, close) =>
    open + (body.length > 2500 ? body.slice(0, 2500) + "\n[… содержимое файла сокращено …]\n" : body) + (close || "</file>"))
    .slice(0, MAX_STORED_ANSWER);
}

export async function sendFile(telegram, chatId, name, mime, bytes, caption) {
  return telegram.upload("sendDocument", {
    chat_id: String(chatId),
    ...(caption ? { caption: caption.slice(0, 1000) } : {}),
  }, { field: "document", name, mime, bytes });
}

/**
 * Sends an answer: text as formatted messages, <file> blocks as new documents,
 * <edit> blocks applied to the user's last file (same formatting kept).
 */
export async function deliver(ctx, uid, chatId, answer, replyTo) {
  const { telegram, kv } = ctx;
  const edited = extractEditBlocks(answer);
  const { text, files } = extractFileBlocks(edited.text);
  const notes = [];
  // edits of the user's own file
  for (const e of edited.edits) {
    const last = kv ? await kv.get(fileKey(uid), "json").catch(() => null) : null;
    const kind = last && editKindOf(last.name);
    if (!last?.fileId || !kind) { notes.push("⚠️ Не нашёл файл для правки — пришли его ещё раз."); continue; }
    try {
      const original = await downloadFile(telegram, ctx.env, ctx.fetchImpl, last.fileId, MAX_FILE_SIZE, "file_too_large");
      const { bytes, report } = await applyEdits(kind, original, e.ops);
      const sent = await sendFile(telegram, chatId, last.name, MIME[kind], bytes,
        `✏️ Правок внесено: ${report.applied}${report.missing.length ? ` · не найдено: ${report.missing.slice(0, 5).join(", ")}` : ""}`);
      log("file_edited", { kind, applied: report.applied, missing: report.missing.length });
      // the edited version becomes the file for the next edits
      const newId = sent?.document?.file_id;
      if (newId) {
        const doc = await describeDocument(kind, last.name, bytes).catch(() => null);
        await kv.put(fileKey(uid), JSON.stringify({ ...last, fileId: newId, text: doc ? doc.text.slice(0, MAX_LAST_FILE_CHARS) : last.text, ts: Date.now() }), { expirationTtl: LAST_FILE_TTL }).catch(() => {});
      }
    } catch (err) {
      if (err instanceof TelegramError && err.status >= 500) throw err;
      log("edit_error", { message: String(err?.message || "").slice(0, 200) });
      notes.push("⚠️ Не получилось применить правки к файлу. Попробуй сформулировать иначе или пришли файл ещё раз.");
    }
  }
  const message = [text, ...notes].filter(Boolean).join("\n\n");
  if (message || (!files.length && !edited.edits.length)) await sendFormatted(telegram, chatId, message || answer, replyTo);
  for (const f of files) {
    let bytes;
    try { bytes = buildFile(f); }
    catch (err) {
      log("build_error", { type: f.type, message: String(err?.message || "").slice(0, 200) });
      await sendFormatted(telegram, chatId, `Не получилось собрать файл «${f.name}». Вот его содержимое текстом:\n\n${f.content}`);
      continue;
    }
    await sendFile(telegram, chatId, f.name, MIME[f.type], bytes);
    log("file_sent", { type: f.type, bytes: bytes.length });
    // keep the full source so the file can be reworked on request
    if (kv) await kv.put(genKey(uid), JSON.stringify({ type: f.type, name: f.name, attrs: f.attrs || {}, content: f.content.slice(0, 60_000), ts: Date.now() }), { expirationTtl: LAST_GEN_TTL }).catch(() => {});
  }
}

/**
 * Runs one model request for a user and delivers the answer.
 * Returns "ok" | "failed" | "undelivered".
 */
async function runRequest(ctx, uid, chatId, updateId, req, replyTo) {
  const { env, kv, telegram, fetchImpl } = ctx;

  let rec = await loadRecord(kv, uid);
  rec.pending = {
    updateId, startedAt: Date.now(), chatId,
    text: req.text, quoted: req.quoted || "", image: req.image || null, file: req.file || null,
  };
  await saveRecord(kv, uid, rec);

  const typing = () => telegram("sendChatAction", { chat_id: chatId, action: req.wantsFile ? "upload_document" : "typing" }).catch(() => {});
  typing();
  const timer = setInterval(typing, 4500);

  let result;
  try {
    let attachment = null;
    let prefix = "";
    if (req.image) attachment = await getImage(telegram, env, fetchImpl, req.image.fileId, req.image.mime);
    if (req.file) {
      const doc = await loadDocument(ctx, req.file);
      if (doc.attachment) attachment = doc.attachment;
      prefix = documentPrompt(doc, false);
      // Remember the document for follow-up questions (one KV write per file).
      const keep = doc.kind === "pdf"
        ? { name: doc.name, kind: "pdf", fileId: req.file.fileId, ts: Date.now() }
        : { name: doc.name, kind: doc.kind, fileId: req.file.fileId, text: doc.text.slice(0, MAX_LAST_FILE_CHARS), note: doc.note, ts: Date.now() };
      await kv.put(fileKey(uid), JSON.stringify(keep), { expirationTtl: LAST_FILE_TTL }).catch(() => {});
    } else if (!req.image) {
      const [last, gen] = await Promise.all([
        kv.get(fileKey(uid), "json").catch(() => null),
        kv.get(genKey(uid), "json").catch(() => null),
      ]);
      const parts = [];
      if (last?.kind === "pdf" && last.fileId) {
        const doc = await loadDocument(ctx, { fileId: last.fileId, name: last.name, kind: "pdf" }).catch(() => null);
        if (doc) { attachment = doc.attachment; parts.push(documentPrompt(doc, true)); }
      } else if (last?.text) {
        parts.push(documentPrompt(last, true));
      }
      // the bot's own last file: only when the message is about changing/making files
      if (gen && (EDIT_WORDS.test(req.text) || FILE_WORDS.test(req.text)) && (!last || (gen.ts || 0) >= (last.ts || 0) || /презентац|слайд|таблиц|эксел|excel/i.test(req.text))) {
        const attrs = Object.entries(gen.attrs || {}).filter(([k]) => k !== "type" && k !== "name").map(([k, v]) => ` ${k}="${String(v).replace(/"/g, "'")}"`).join("");
        parts.push(`(Последний файл, который ты создал: «${gen.name}». Его полный исходник ниже. Если просят изменить ЭТОТ файл — пришли новый полный блок <file> с тем же type и name и с учётом просьбы.)\n<file type="${gen.type}" name="${gen.name}"${attrs}>\n${gen.content}\n</file>`);
        req.wantsFile = true;
      }
      prefix = parts.join("\n\n");
    }
    if (req.file && EDIT_WORDS.test(req.text)) req.wantsFile = true;
    const question = userPrompt(req);
    result = await askModel(env, {
      system: SYSTEM_PROMPT,
      history: rec.messages,
      text: prefix ? `${prefix}\n\n${question}` : question,
      image: attachment,
      model: currentModel(rec, env),
      maxTokens: req.wantsFile ? 8000 : undefined,
    }, { fetchImpl, timeoutMs: ctx.modelTimeoutMs ?? MODEL_TIMEOUT_MS });
  } catch (err) {
    clearInterval(timer);
    let code = err instanceof ModelError ? err.code : "unknown";
    if (code === "bad_request" && req.file?.kind === "pdf") code = "pdf_rejected";
    else if (code === "bad_request" && (req.image || req.file)) code = "image_unsupported";
    if (!(err instanceof ModelError)) log("request_error", { name: err?.name, message: String(err?.message || "").slice(0, 200) });
    await telegram("sendMessage", { chat_id: chatId, text: errorText(code) }).catch(e => log("send_error", { message: e.message }));
    return "failed";
  }
  clearInterval(timer);

  let answer = result.text;
  if (result.truncated) answer += "\n\n*Ответ получился очень длинным и оборвался. Напиши «продолжи».*";

  // Re-read right before writing so a parallel request's result isn't lost.
  rec = await loadRecord(kv, uid);
  rec.messages = trimHistory([
    ...rec.messages,
    { role: "user", content: historyUserEntry(req) },
    { role: "assistant", content: answerForHistory(answer) },
  ]);
  if (rec.pending?.updateId === updateId) delete rec.pending;
  rec.lastUpdateId = updateId;
  delete rec.undelivered;
  await saveRecord(kv, uid, rec);

  try {
    await deliver(ctx, uid, chatId, answer, replyTo);
    return "ok";
  } catch (err) {
    log("deliver_error", { message: err.message });
    // Answer is saved; remember to resend it when Telegram retries the update.
    rec.undelivered = { chatId, text: answer.slice(0, 100_000) };
    await saveRecord(kv, uid, rec).catch(() => {});
    return "undelivered";
  }
}

/** /word: sends the user's last answer as a Word file, without calling the model. */
async function exportLastAnswer(ctx, uid, chatId) {
  const rec = await loadRecord(ctx.kv, uid);
  const last = [...rec.messages].reverse().find(m => m.role === "assistant");
  if (!last) {
    await ctx.telegram("sendMessage", { chat_id: chatId, text: "Пока нечего сохранять: сначала задай вопрос 🙂" });
    return;
  }
  const { text, files } = extractFileBlocks(last.content);
  const docx = files.find(f => f.type === "docx");
  const source = docx ? docx.content : text.replace(/\n\n\*Ответ получился очень длинным[\s\S]*$/, "");
  const firstLine = (source.split("\n").find(l => l.trim()) || "").replace(/^#+\s*|\*\*/g, "").trim();
  const name = safeFileName(docx?.name || firstLine.slice(0, 50) || "Ответ", "docx");
  const bytes = buildDocx(source, { font: docx?.font, size: docx?.size, theme: docx?.theme, style: docx?.style });
  await sendFile(ctx.telegram, chatId, name, MIME.docx, bytes);
}

/* ============================== server agent ============================== */

const AGENT_WORDS = /график|диаграм|схем|pdf|пдф|рисун|картин|визуал|постер|инфограф|плакат|открытк|логотип|чертёж|чертеж|майнд|mind ?map|красив|дизайн|оформи/i;

export function agentEnabled(env, rec) {
  return !!(env.AGENT_URL && env.AGENT_SECRET) && rec?.agent !== false;
}

export function shouldUseAgent(req) {
  const t = req.text || "";
  if (AGENT_WORDS.test(t) || FILE_WORDS.test(t)) return true;
  if (req.file && ["docx", "xlsx", "pptx", "pdf"].includes(req.file.kind) && EDIT_WORDS.test(t)) return true;
  return false;
}

/** Hands a heavy task to the server agent. Returns true when the agent accepted it. */
async function dispatchToAgent(ctx, uid, chatId, msg, req, rec) {
  const { env, kv, fetchImpl } = ctx;
  const files = [];
  const photos = [];
  if (req.file) files.push({ file_id: req.file.fileId, name: req.file.name });
  if (req.image) photos.push({ file_id: req.image.fileId });
  if (!req.file && !req.image && EDIT_WORDS.test(req.text || "")) {
    const last = await kv.get(fileKey(uid), "json").catch(() => null);
    if (last?.fileId) files.push({ file_id: last.fileId, name: last.name });
  }
  const job = {
    chat_id: chatId, uid, reply_to: msg.message_id, model: currentModel(rec, env),
    text: userPrompt(req), history: rec.messages.slice(-8), history_entry: historyUserEntry(req), files, photos,
  };
  try {
    const res = await fetchImpl(`${String(env.AGENT_URL).replace(/\/+$/, "")}/job`, {
      method: "POST",
      headers: { "content-type": "application/json", "x-agent-secret": env.AGENT_SECRET },
      body: JSON.stringify(job),
      signal: AbortSignal.timeout(10_000),
    });
    if (res.status === 202 || res.ok) {
      const data = await res.json().catch(() => ({}));
      log("agent_dispatched", { uid, position: data.position });
      if (data.position > 0) await ctx.telegram("sendMessage", { chat_id: chatId, text: `⏳ Задача в очереди (перед ней ${data.position}). Пришлю результат, как будет готово.` }).catch(() => {});
      return true;
    }
    log("agent_rejected", { status: res.status });
  } catch (err) {
    log("agent_unreachable", { message: String(err?.message || err).slice(0, 120) });
  }
  return false;
}

/** POST /agent/done from the server: store the exchange in the user's history. */
async function agentDone(request, env, kv) {
  if (!env.AGENT_SECRET || !safeEqual(request.headers.get("x-agent-secret"), env.AGENT_SECRET)) return json({ error: "Forbidden" }, 403);
  if (!kv) return json({ error: "no storage" }, 500);
  let body;
  try { body = await request.json(); } catch { return json({ error: "Invalid JSON" }, 400); }
  const uid = String(body.uid || "");
  if (!allowedIds(env).includes(uid)) return json({ error: "unknown user" }, 403);
  const rec = await loadRecord(kv, uid);
  rec.messages = trimHistory([
    ...rec.messages,
    { role: "user", content: String(body.user || "").slice(0, MAX_USER_CHARS) || "(задача для сервера)" },
    { role: "assistant", content: String(body.answer || "Готово.").slice(0, MAX_STORED_ANSWER) },
  ]);
  await saveRecord(kv, uid, rec);
  const f = Array.isArray(body.files) ? body.files.filter(x => x?.file_id && x?.name).pop() : null;
  if (f) await kv.put(fileKey(uid), JSON.stringify({ name: f.name, kind: classifyDocument({ file_name: f.name }).kind || "file", fileId: f.file_id, ts: Date.now() }), { expirationTtl: LAST_FILE_TTL }).catch(() => {});
  return json({ ok: true });
}

async function agentStatus(env, fetchImpl) {
  if (!env.AGENT_URL || !env.AGENT_SECRET) return "не подключён (нет AGENT_URL / AGENT_SECRET в Cloudflare)";
  try {
    const res = await fetchImpl(`${String(env.AGENT_URL).replace(/\/+$/, "")}/health`, { signal: AbortSignal.timeout(6000) });
    const d = await res.json();
    return d.ok ? `работает ✅ (в очереди: ${d.queue}${d.running ? ", сейчас занят задачей" : ""}; выполнено: ${d.done})` : `отвечает с ошибкой (${res.status})`;
  } catch (err) {
    return `не отвечает ❌ (${String(err?.message || err).slice(0, 80)})`;
  }
}

/**
 * Handles one Telegram update. Returns "ok" or "retry" (ask Telegram to redeliver).
 */
export async function handleUpdate(update, ctx) {
  const { env, telegram } = ctx;
  if (update?.callback_query) return handleCallback(update.callback_query, ctx);
  const msg = update?.message;
  if (!msg || msg.chat?.type !== "private" || !msg.from?.id || msg.from.is_bot) return "ok";
  const uid = String(msg.from.id);
  const chatId = msg.chat.id;
  const updateId = update.update_id;
  const rawText = (msg.text || msg.caption || "").trim();
  const command = rawText.startsWith("/") ? rawText.split(/\s+/)[0].split("@")[0].toLowerCase() : "";

  // Anyone may ask for their own ID; it works even if KV or the API key is broken.
  if (command === "/id") {
    await sendHtml(telegram, chatId, `Твой Telegram ID: <code>${uid}</code>`);
    return "ok";
  }
  if (!allowedIds(env).includes(uid)) {
    if (!allowedIds(env).length) log("config_warning", { problem: "ALLOWED_USER_IDS is empty" });
    log("access_denied", { uid });
    await telegram("sendMessage", { chat_id: chatId, text: "🔒 Это приватный бот. Чтобы узнать свой Telegram ID, отправь /id." });
    return "ok";
  }
  if (command === "/start" || command === "/help") {
    await sendHtml(telegram, chatId, HELP);
    return "ok";
  }

  const kv = ctx.kv;
  if (!kv) {
    log("config_error", { problem: "KV binding HISTORY is missing" });
    await telegram("sendMessage", { chat_id: chatId, text: errorText("no_storage") });
    return "ok";
  }

  if (command === "/reset") {
    // Only the caller's own keys, derived from the verified sender ID.
    const prev = await loadRecord(kv, uid);
    if (prev.model) await saveRecord(kv, uid, { v: 2, messages: [], model: prev.model });
    else await kv.delete(historyKey(uid));
    await kv.delete(fileKey(uid));
    await kv.delete(genKey(uid));
    await telegram("sendMessage", { chat_id: chatId, text: "🧹 Твоя история очищена, присланный файл забыт. Начинаем с чистого листа." });
    return "ok";
  }
  if (command === "/model" || command === "/models") {
    const arg = rawText.split(/\s+/)[1];
    if (arg) {
      const models = await availableModels(ctx);
      const pick = models.find(m => m.toLowerCase() === arg.toLowerCase()) || models.find(m => m.toLowerCase().includes(arg.toLowerCase()));
      if (!pick) { await telegram("sendMessage", { chat_id: chatId, text: "Не нашёл такую модель. Список: /model" }); return "ok"; }
      await setUserModel(ctx, uid, pick);
      await sendHtml(telegram, chatId, `🧠 Модель: <b>${escapeHtml(prettyModel(pick))}</b>`);
      return "ok";
    }
    await showModelMenu(ctx, uid, chatId);
    return "ok";
  }
  if (command === "/agent" || command === "/server") {
    const arg = (rawText.split(/\s+/)[1] || "").toLowerCase();
    const rec0 = await loadRecord(kv, uid);
    if (arg === "off" || arg === "выкл") { rec0.agent = false; await saveRecord(kv, uid, rec0); }
    if (arg === "on" || arg === "вкл") { delete rec0.agent; await saveRecord(kv, uid, rec0); }
    const st = await agentStatus(env, ctx.fetchImpl);
    await sendHtml(telegram, chatId, `🖥 <b>Сервер-агент</b>: ${escapeHtml(st)}\nДля тебя: <b>${rec0.agent === false ? "выключен" : "включён"}</b>\n\nАгент сам пишет код, проверяет результат и присылает файлы — для презентаций, документов, таблиц, графиков, PDF.\n/agent on · /agent off`);
    return "ok";
  }
  if (command === "/setup") {
    const origin = ctx.origin || "https://family-tg-gpt.qlovver.workers.dev";
    await sendHtml(telegram, chatId, `⚙️ Настройка открывается в браузере:\n${escapeHtml(origin)}/setup\n\nПроверить модель можно прямо здесь: /check`);
    return "ok";
  }
  if (command === "/check") {
    const rec0 = await loadRecord(kv, uid);
    const model = currentModel(rec0, env);
    await sendHtml(telegram, chatId, `🔎 Проверяю <code>${escapeHtml(model)}</code>… (несколько маленьких запросов, ~20–40 с)`);
    const lines = [`🔎 <b>Проверка модели</b> <code>${escapeHtml(model)}</code>`];
    try {
      const c = await probeCapabilities(env, model, { fetchImpl: ctx.fetchImpl });
      lines.push(`Инструменты: ${escapeHtml(c.tools)}`, `Зрение: ${escapeHtml(c.vision)}`, `Модель в ответе: <code>${escapeHtml(c.response_model || "—")}</code>`);
    } catch (err) { lines.push(`Инструменты/зрение: ошибка ${escapeHtml(String(err?.message || err).slice(0, 80))}`); }
    if (/^claude/i.test(model)) {
      try {
        const a = await probeAuthenticity(env, model, { fetchImpl: ctx.fetchImpl });
        lines.push("", `<b>Подлинность:</b> ${escapeHtml(a.verdict)}`, `ID ответа: <code>${escapeHtml(a.response_id || "—")}</code> — ${escapeHtml(a.id_format || "")}`);
        if (a.signature) lines.push(`Подпись размышлений: ${escapeHtml(a.signature)}`);
        if (a.tampered_signature_rejected) lines.push(`Испорченная подпись отклонена: ${escapeHtml(a.tampered_signature_rejected)}`);
      } catch (err) { lines.push(`Подлинность: ошибка ${escapeHtml(String(err?.message || err).slice(0, 80))}`); }
    } else lines.push("", "Проверка подлинности по подписи есть только для Claude-моделей.");
    await sendHtml(telegram, chatId, lines.join("\n"));
    return "ok";
  }
  if (command === "/word") {
    await exportLastAnswer(ctx, uid, chatId);
    return "ok";
  }

  if (!env.TOOKEN_API_KEY) {
    log("config_error", { problem: "TOOKEN_API_KEY is missing" });
    await telegram("sendMessage", { chat_id: chatId, text: errorText("no_api_key") });
    return "ok";
  }

  // Duplicate / redelivered update handling (no extra KV writes in the normal case).
  const rec = await loadRecord(kv, uid);
  if (typeof updateId === "number") {
    if (rec.lastUpdateId === updateId) {
      if (rec.undelivered) {
        try {
          await deliver(ctx, uid, rec.undelivered.chatId, rec.undelivered.text);
          delete rec.undelivered;
          await saveRecord(kv, uid, rec);
          return "ok";
        } catch { return "retry"; }
      }
      log("duplicate_update", { updateId });
      return "ok";
    }
    if (rec.pending?.updateId === updateId && Date.now() - rec.pending.startedAt < STALE_PENDING_MS) {
      log("duplicate_in_progress", { updateId });
      return "ok";
    }
  }

  if (command === "/retry") {
    const p = rec.pending;
    if (!p) {
      await telegram("sendMessage", { chat_id: chatId, text: "Нечего повторять: последний запрос уже получил ответ." });
      return "ok";
    }
    const req = { text: p.text || "", quoted: p.quoted || "", image: p.image || null, file: p.file || null };
    req.wantsFile = FILE_WORDS.test(req.text);
    const r = await runRequest(ctx, uid, chatId, updateId, req, msg.message_id);
    return r === "undelivered" ? "retry" : "ok";
  }
  if (command) {
    await sendHtml(telegram, chatId, "Команды: /start, /help, /model, /agent, /check, /word, /reset, /retry, /id");
    return "ok";
  }

  const req = extractRequest(msg);
  if (req.error) {
    await telegram("sendMessage", { chat_id: chatId, text: errorText(req.error, req.hint) });
    return "ok";
  }
  if (!req.text && !req.image && !req.file) {
    await telegram("sendMessage", { chat_id: chatId, text: "Пришли текст, фото задания или файл 🙂" });
    return "ok";
  }
  req.wantsFile = FILE_WORDS.test(req.text);
  if (agentEnabled(env, rec) && shouldUseAgent(req)) {
    if (await dispatchToAgent(ctx, uid, chatId, msg, req, rec)) {
      // remember this update so a Telegram retry doesn't start the job twice
      rec.lastUpdateId = updateId;
      await saveRecord(kv, uid, rec);
      return "ok";
    }
    await telegram("sendMessage", { chat_id: chatId, text: "🖥 Сервер сейчас недоступен — сделаю попроще, прямо здесь." }).catch(() => {});
  }
  const r = await runRequest(ctx, uid, chatId, updateId, req, msg.message_id);
  return r === "undelivered" ? "retry" : "ok";
}

/* ============================== setup page ============================== */

function setupPage(content = "") {
  return `<!doctype html><html lang="ru"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Настройка бота</title>
<style>
body{font:16px/1.5 -apple-system,system-ui,sans-serif;max-width:560px;margin:0 auto;padding:16px;background:#f6f6f8;color:#111}
@media (prefers-color-scheme:dark){body{background:#111;color:#eee}input,button{color:#111}}
form{display:grid;gap:12px;margin:16px 0}
input[type=password]{font-size:16px;padding:10px;border-radius:8px;border:1px solid #999}
button{font-size:16px;padding:12px;border-radius:8px;border:0;background:#2a7ae2;color:#fff}
label{display:flex;gap:8px;align-items:center}
pre{white-space:pre-wrap;word-break:break-word;background:rgba(127,127,127,.15);padding:12px;border-radius:8px}
</style></head><body>
<h2>Настройка Telegram-бота</h2>
<form method="post" action="/setup">
<input type="password" name="key" placeholder="SETUP_KEY" autocomplete="current-password" required>
<label><input type="checkbox" name="webhook" checked> Подключить / обновить webhook</label>
<label><input type="checkbox" name="model"> Проверить модель (1 маленький платный запрос)</label>
<label><input type="checkbox" name="caps"> Проверить инструменты и зрение моделей (по 2 маленьких запроса на модель)</label>
<label><input type="checkbox" name="auth"> Проверить подлинность Claude (3 маленьких запроса на каждую Claude-модель)</label>
<input type="text" name="caps_models" placeholder="модели через запятую (пусто — Claude Opus и GPT-6.1 Sol)" style="font-size:16px;padding:10px;border-radius:8px;border:1px solid #999">
<button type="submit">Выполнить</button>
</form>
${content}
</body></html>`;
}

async function runSetup(request, env, ctx) {
  const form = await request.formData().catch(() => null);
  const key = form?.get("key");
  if (!env.SETUP_KEY || !safeEqual(key, env.SETUP_KEY)) {
    log("setup_denied");
    return html(setupPage("<p>❌ Неверный ключ.</p>"), 403);
  }
  const telegram = makeTelegram(env, ctx.fetchImpl);
  const report = {};
  report.config = configChecks(env, ctx.kv);
  const origin = new URL(request.url).origin;

  if (form.get("webhook")) {
    if (!env.WEBHOOK_SECRET || !env.TELEGRAM_BOT_TOKEN) {
      report.webhook = "Не заданы WEBHOOK_SECRET или TELEGRAM_BOT_TOKEN";
    } else {
      try {
        await telegram("setWebhook", {
          url: `${origin}/webhook`,
          secret_token: env.WEBHOOK_SECRET,
          // One update at a time: keeps each person's dialogue in order and
          // prevents parallel writes to the same history.
          max_connections: 1,
          allowed_updates: ["message", "callback_query"],
        });
        await telegram("setMyCommands", { commands: [
          { command: "help", description: "Как пользоваться ботом" },
          { command: "model", description: "Выбрать модель ИИ" },
          { command: "agent", description: "Сервер-агент: статус, вкл/выкл" },
          { command: "check", description: "Проверить модель" },
          { command: "reset", description: "Очистить мою историю" },
          { command: "retry", description: "Повторить неудачный запрос" },
          { command: "id", description: "Мой Telegram ID" },
        ] }).catch(() => {});
        report.webhook = "✅ Webhook подключён";
      } catch (err) {
        report.webhook = `❌ ${err.message}`;
      }
    }
  }
  if (env.TELEGRAM_BOT_TOKEN) {
    try {
      const info = await telegram("getWebhookInfo", {});
      const me = await telegram("getMe", {});
      report.bot = `@${me.username}`;
      report.webhookInfo = {
        url: info.url,
        pending_update_count: info.pending_update_count,
        max_connections: info.max_connections,
        last_error: info.last_error_message
          ? `${info.last_error_message} (${new Date(info.last_error_date * 1000).toISOString()})` : "нет",
      };
    } catch (err) {
      report.bot = `❌ ${err.message}`;
    }
  }
  if (form.get("model")) {
    const started = Date.now();
    try {
      const r = await askModel({ ...env, MAX_OUTPUT_TOKENS: "20" }, {
        system: "Отвечай очень кратко.", history: [], text: "Ответь одним словом: работает",
      }, { fetchImpl: ctx.fetchImpl, timeoutMs: 60_000 });
      report.model = {
        result: "✅ Ответ получен", endpoint: r.format, model_in_response: r.model || "(не указана)",
        answer: r.text.slice(0, 100), ms: Date.now() - started, request_id: r.requestId || "(нет)",
        usage: r.usage || null,
      };
    } catch (err) {
      report.model = { result: `❌ ${err.code || err.message}`, status: err.status, detail: err.detail, request_id: err.requestId };
    }
  }
  if (form.get("caps")) {
    const list = String(form.get("caps_models") || "").split(/[\s,;]+/).filter(Boolean).slice(0, 6);
    const models = list.length ? list : [env.MODEL || DEFAULT_MODEL, "gpt-6.1-sol"];
    report.capabilities = [];
    for (const m of [...new Set(models)]) report.capabilities.push(await probeCapabilities(env, m, { fetchImpl: ctx.fetchImpl }));
  }
  if (form.get("auth")) {
    const list = String(form.get("caps_models") || "").split(/[\s,;]+/).filter(m => /^claude/i.test(m)).slice(0, 4);
    const models = list.length ? list : [/^claude/i.test(env.MODEL || DEFAULT_MODEL) ? env.MODEL || DEFAULT_MODEL : DEFAULT_MODEL];
    report.authenticity = [];
    for (const m of [...new Set(models)]) {
      try { report.authenticity.push(await probeAuthenticity(env, m, { fetchImpl: ctx.fetchImpl })); }
      catch (err) { report.authenticity.push({ model: m, verdict: `❓ ошибка проверки: ${String(err?.message || err).slice(0, 120)}` }); }
    }
  }
  return html(setupPage(`<pre>${escapeHtml(JSON.stringify(report, null, 2))}</pre>`));
}

function configChecks(env, kv) {
  const cfg = modelConfig(env);
  return {
    HISTORY_binding: !!kv,
    TELEGRAM_BOT_TOKEN: !!env.TELEGRAM_BOT_TOKEN,
    TOOKEN_API_KEY: !!env.TOOKEN_API_KEY,
    WEBHOOK_SECRET: !!env.WEBHOOK_SECRET,
    SETUP_KEY: !!env.SETUP_KEY,
    allowed_users: allowedIds(env).length,
    agent: !!(env.AGENT_URL && env.AGENT_SECRET),
    model: cfg.model,
    api_format: cfg.format,
  };
}

/* ============================== worker entry ============================== */

export function createWorker(options = {}) {
  return {
    async fetch(request, env, execCtx) {
      const fetchImpl = options.fetchImpl || ((...a) => fetch(...a));
      const kv = env.HISTORY || null;
      const url = new URL(request.url);
      const ctx = { env, kv, fetchImpl, telegram: makeTelegram(env, fetchImpl), modelTimeoutMs: options.modelTimeoutMs, origin: url.origin };

      if (request.method === "GET" && url.pathname === "/") {
        const checks = configChecks(env, kv);
        const ok = checks.HISTORY_binding && checks.TELEGRAM_BOT_TOKEN && checks.TOOKEN_API_KEY &&
          checks.WEBHOOK_SECRET && checks.allowed_users > 0;
        return json({ status: ok ? "running" : "misconfigured", checks, setup: "/setup" });
      }
      if (url.pathname === "/setup") {
        if (request.method === "POST") return runSetup(request, env, ctx);
        if (request.method === "GET") return html(setupPage(url.searchParams.has("code")
          ? "<p>⚠️ Ключ больше не передаётся в адресе. Введи его в поле выше. Ссылку с ключом лучше удалить из истории браузера.</p>" : ""));
      }
      if (request.method === "POST" && url.pathname === "/agent/done") return agentDone(request, env, kv);
      if (request.method !== "POST" || url.pathname !== "/webhook") return json({ error: "Not found" }, 404);

      if (!env.WEBHOOK_SECRET || !safeEqual(request.headers.get("X-Telegram-Bot-Api-Secret-Token"), env.WEBHOOK_SECRET)) {
        log("webhook_forbidden");
        return json({ error: "Forbidden" }, 403);
      }
      if (!env.TELEGRAM_BOT_TOKEN) {
        log("config_error", { problem: "TELEGRAM_BOT_TOKEN is missing" });
        return json({ error: "Missing TELEGRAM_BOT_TOKEN" }, 500);
      }
      let update;
      try { update = await request.json(); } catch { return json({ error: "Invalid JSON" }, 400); }

      // Await the work (Telegram waits for us), and also register it with
      // waitUntil so it gets up to 30 s more if Telegram drops the connection.
      const task = handleUpdate(update, ctx).catch(err => {
        log("update_crash", { name: err?.name, message: String(err?.message || "").slice(0, 200) });
        return err instanceof TelegramError && err.status >= 500 ? "retry" : "ok";
      });
      execCtx?.waitUntil?.(task);
      const result = await task;
      return result === "retry" ? json({ ok: false, retry: true }, 500) : json({ ok: true });
    },
  };
}

export default createWorker();
