// Telegram bot logic for Cloudflare Workers.
// Secrets:   TELEGRAM_BOT_TOKEN, TOOKEN_API_KEY, WEBHOOK_SECRET, SETUP_KEY
// Variable:  ALLOWED_USER_IDS="123,456" (a Secret also works)
// KV:        HISTORY (namespace kv-get-started)
// Optional:  MODEL, API_FORMAT (messages|chat|responses), MODEL_STREAM, MAX_OUTPUT_TOKENS, TOOKEN_BASE_URL

import { toTelegramHtml, splitHtml, htmlToPlain, escapeHtml } from "./format.js";
import { askModel, ModelError, modelConfig } from "./model.js";

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

export function makeTelegram(env, fetchImpl) {
  return async function telegram(method, payload) {
    const res = await fetchImpl(`https://api.telegram.org/bot${env.TELEGRAM_BOT_TOKEN}/${method}`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(payload),
    });
    let data = {};
    try { data = await res.json(); } catch { /* ignore */ }
    if (!res.ok || !data.ok) throw new TelegramError(method, res.status, String(data.description || "unknown").slice(0, 200));
    return data.result;
  };
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
  const bytes = new Uint8Array(arrayBuffer);
  let bin = "";
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(bin);
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

// Extracts what we need to call the model from a Telegram message.
export function extractRequest(msg) {
  const text = (msg.text || msg.caption || "").trim();
  let image = null, error = null;
  const fromMessage = m => {
    if (m.photo?.length) {
      const p = pickPhoto(m.photo);
      if (!p) return { error: "image_too_large" };
      return { image: { fileId: p.file_id, mime: "image/jpeg" } };
    }
    if (m.document) {
      const mime = m.document.mime_type || "";
      if (!IMAGE_MIMES.includes(mime)) return { error: mime === "application/pdf" ? "pdf_unsupported" : "file_unsupported" };
      if (m.document.file_size && m.document.file_size > MAX_IMAGE_SIZE) return { error: "image_too_large" };
      return { image: { fileId: m.document.file_id, mime } };
    }
    return {};
  };
  ({ image, error } = fromMessage(msg));
  let quoted = "";
  const r = msg.reply_to_message;
  if (r) {
    // Replying to an earlier photo with a question re-uses that photo.
    if (!image && !error) {
      const fromReply = fromMessage(r);
      if (fromReply.image) image = { ...fromReply.image, fromReply: true };
    }
    const rt = (r.text || r.caption || "").trim();
    if (rt) quoted = rt.slice(0, 1200);
  }
  return { text: text.slice(0, MAX_USER_CHARS), image, error, quoted };
}

/* ============================== history ============================== */

export function historyKey(uid) {
  return `history:${uid}`;
}

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

/* ============================== texts ============================== */

const HELP = `<b>Привет! Я ИИ-помощник 🤖</b>

Пиши вопрос текстом или присылай <b>фото задания</b> — разберу и решу по шагам.

<b>Подсказки</b>
• Одно фото — одно задание: так ответ точнее.
• Можно добавить подпись к фото: «реши пункт б».
• Чтобы спросить ещё про старое фото, ответь (reply) на него.
• Если ответ оборвался — напиши «продолжи».

<b>Команды</b>
/reset — очистить историю (только свою)
/retry — повторить последний неудачный запрос
/id — узнать свой Telegram ID`;

export function errorText(code) {
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
      return "❓ Tooken Club не нашёл модель. Нужно проверить название модели и доступ к ней в кабинете Tooken.";
    case "endpoint_not_found":
    case "bad_request":
      return "⚠️ Сервис ИИ отклонил запрос (неподходящий формат). Отправь /retry; если повторится — нужна проверка настроек бота.";
    case "context_too_long":
    case "request_too_large":
      return "📚 Запрос получился слишком большим. Отправь /reset и задай вопрос заново.";
    case "image_too_large":
      return "🖼 Фото слишком большое. Отправь его как обычное фото (не файлом), до 2,5 МБ.";
    case "image_download_failed":
      return "🖼 Не получилось скачать фото с серверов Telegram. Отправь его ещё раз.";
    case "pdf_unsupported":
      return "📄 PDF пока не поддерживается. Сделай скриншот нужной страницы и пришли как фото.";
    case "file_unsupported":
      return "📎 Такой файл я не читаю. Пришли текст или изображение (JPG, PNG, WEBP).";
    case "no_storage":
      return "🛠 Бот не настроен: нет хранилища HISTORY. Владельцу нужно проверить KV binding в Cloudflare.";
    case "no_api_key":
      return "🛠 Бот не настроен: нет ключа TOOKEN_API_KEY.";
    default:
      return "😕 Не удалось получить ответ от ИИ. Отправь /retry чуть позже. Если ошибка повторяется — проверь баланс Tooken Club.";
  }
}

/* ============================== bot core ============================== */

async function getImage(telegram, env, fetchImpl, fileId, mimeHint) {
  const f = await telegram("getFile", { file_id: fileId });
  if (!f?.file_path) throw new ModelError("image_download_failed");
  if (f.file_size && f.file_size > MAX_IMAGE_SIZE) throw new ModelError("image_too_large");
  const res = await fetchImpl(`https://api.telegram.org/file/bot${env.TELEGRAM_BOT_TOKEN}/${f.file_path}`);
  if (!res.ok) throw new ModelError("image_download_failed");
  const data = await res.arrayBuffer();
  if (data.byteLength > MAX_IMAGE_SIZE) throw new ModelError("image_too_large");
  return { mime: detectMime(data, mimeHint), base64: imageToBase64(data) };
}

function userPrompt(req) {
  let text = req.text;
  if (!text && req.image) text = "Реши задание на фото подробно по шагам, проверь вычисления и выдели ответ.";
  if (req.quoted && !req.image?.fromReply) text = `(Это ответ на сообщение: «${req.quoted}»)\n\n${text}`;
  return text;
}

function historyUserEntry(req) {
  const mark = req.image ? (req.image.fromReply ? "[Вопрос по ранее присланному фото] " : "[Фото задания] ") : "";
  return (mark + (req.text || (req.image ? "(без подписи)" : ""))).slice(0, MAX_USER_CHARS);
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
    text: req.text, quoted: req.quoted || "", image: req.image || null,
  };
  await saveRecord(kv, uid, rec);

  const typing = () => telegram("sendChatAction", { chat_id: chatId, action: "typing" }).catch(() => {});
  typing();
  const timer = setInterval(typing, 4500);

  let result;
  try {
    let image = null;
    if (req.image) image = await getImage(telegram, env, fetchImpl, req.image.fileId, req.image.mime);
    result = await askModel(env, {
      system: SYSTEM_PROMPT,
      history: rec.messages,
      text: userPrompt(req),
      image,
    }, { fetchImpl, timeoutMs: ctx.modelTimeoutMs ?? MODEL_TIMEOUT_MS });
  } catch (err) {
    clearInterval(timer);
    const code = err instanceof ModelError ? err.code : "unknown";
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
    { role: "assistant", content: answer.slice(0, MAX_STORED_ANSWER) },
  ]);
  if (rec.pending?.updateId === updateId) delete rec.pending;
  rec.lastUpdateId = updateId;
  delete rec.undelivered;
  await saveRecord(kv, uid, rec);

  try {
    await sendFormatted(telegram, chatId, answer, replyTo);
    return "ok";
  } catch (err) {
    log("deliver_error", { message: err.message });
    // Answer is saved; remember to resend it when Telegram retries the update.
    rec.undelivered = { chatId, text: answer };
    await saveRecord(kv, uid, rec).catch(() => {});
    return "undelivered";
  }
}

/**
 * Handles one Telegram update. Returns "ok" or "retry" (ask Telegram to redeliver).
 */
export async function handleUpdate(update, ctx) {
  const { env, telegram } = ctx;
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
    // Only the caller's own key, derived from the verified sender ID.
    await kv.delete(historyKey(uid));
    await telegram("sendMessage", { chat_id: chatId, text: "🧹 Твоя история очищена. Начинаем с чистого листа." });
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
          await sendFormatted(telegram, rec.undelivered.chatId, rec.undelivered.text);
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
    const req = { text: p.text || "", quoted: p.quoted || "", image: p.image || null };
    const r = await runRequest(ctx, uid, chatId, updateId, req, msg.message_id);
    return r === "undelivered" ? "retry" : "ok";
  }
  if (command) {
    await sendHtml(telegram, chatId, "Команды: /start, /help, /reset, /retry, /id");
    return "ok";
  }

  const req = extractRequest(msg);
  if (req.error) {
    await telegram("sendMessage", { chat_id: chatId, text: errorText(req.error) });
    return "ok";
  }
  if (!req.text && !req.image) {
    await telegram("sendMessage", { chat_id: chatId, text: "Пришли текст или фотографию задания 🙂" });
    return "ok";
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
          allowed_updates: ["message"],
        });
        await telegram("setMyCommands", { commands: [
          { command: "help", description: "Как пользоваться ботом" },
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
      const ctx = { env, kv, fetchImpl, telegram: makeTelegram(env, fetchImpl), modelTimeoutMs: options.modelTimeoutMs };
      const url = new URL(request.url);

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
