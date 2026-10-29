// Telegram + Tooken Club bot on Cloudflare Workers (no Node.js required).
// Secrets: TELEGRAM_BOT_TOKEN, TOOKEN_API_KEY, WEBHOOK_SECRET, SETUP_KEY.
// Text variable: ALLOWED_USER_IDS="123,456". KV binding: HISTORY.

const MODEL = "gpt-6.1-terra"; // Verify access to this model in Tooken Club.
const MAX_IMAGE_SIZE = 2_500_000; // 2.5 MB; use Telegram's compressed Photo option.
const MAX_HISTORY = 10; // Text messages, unique for each Telegram user.
const SYSTEM_PROMPT =
  "Ты внимательный русскоязычный помощник. Для домашних заданий сначала прочитай " +
  "условие, затем решай подробно по шагам, объясняя каждое действие и проверяя ответ. " +
  "Если на фото что-то не видно, попроси прислать более чёткое. " +
  "Не выдумывай факты и не показывай чужие переписки.";

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" },
  });
}

async function telegram(env, method, payload) {
  const res = await fetch(`https://api.telegram.org/bot${env.TELEGRAM_BOT_TOKEN}/${method}`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(payload),
  });
  const data = await res.json();
  if (!res.ok || !data.ok) {
    throw new Error(`Telegram ${method} failed (${res.status})`);
  }
  return data.result;
}

async function reply(env, chatId, message) {
  const text = String(message || "");
  // Telegram messages must be <= 4096 chars, use a little safety margin.
  for (let i = 0; i < text.length; i += 3800) {
    await telegram(env, "sendMessage", { chat_id: chatId, text: text.slice(i, i + 3800) });
  }
}

function imageToBase64(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer);
  const chunks = [];
  for (let i = 0; i < bytes.length; i += 16384) {
    chunks.push(String.fromCharCode(...bytes.subarray(i, i + 16384)));
  }
  return btoa(chunks.join(""));
}

async function getImage(env, fileId) {
  const f = await telegram(env, "getFile", { file_id: fileId });
  if (!f.file_path || (f.file_size && f.file_size > MAX_IMAGE_SIZE)) {
    throw new Error("image_too_large");
  }
  const response = await fetch(
    `https://api.telegram.org/file/bot${env.TELEGRAM_BOT_TOKEN}/${f.file_path}`
  );
  if (!response.ok) throw new Error("image_download_failed");
  const data = await response.arrayBuffer();
  if (data.byteLength > MAX_IMAGE_SIZE) throw new Error("image_too_large");
  return data;
}

function parseAnswer(data) {
  if (typeof data.output_text === "string" && data.output_text.trim()) {
    return data.output_text.trim();
  }
  // Responses API: output[].content[].text
  const parts = [];
  for (const item of data.output || []) {
    for (const c of item.content || []) {
      if (typeof c.text === "string") parts.push(c.text);
    }
  }
  if (parts.length) return parts.join("\n").trim();
  // Fallback if the provider returns a chat-completions-shaped response.
  return data.choices?.[0]?.message?.content?.trim() || "";
}

async function askModel(env, history, prompt, photo) {
  const input = history.map(({ role, content }) => ({ role, content }));
  if (photo) {
    input.push({ role: "user", content: [
      { type: "input_text", text: prompt },
      { type: "input_image", image_url: photo },
    ] });
  } else {
    input.push({ role: "user", content: prompt });
  }
  const response = await fetch("https://tooken.club/v1/responses", {
    method: "POST",
    headers: {
      "authorization": `Bearer ${env.TOOKEN_API_KEY}`,
      "content-type": "application/json",
    },
    body: JSON.stringify({
      model: MODEL,
      instructions: SYSTEM_PROMPT,
      input,
      max_output_tokens: 3500,
    }),
    signal: AbortSignal.timeout(45_000),
  });
  if (!response.ok) {
    console.error("Tooken HTTP status:", response.status);
    throw new Error(response.status === 402 || response.status === 429
      ? "quota_or_rate_limit" : "provider_failed");
  }
  const data = await response.json();
  const answer = parseAnswer(data);
  if (!answer) throw new Error("empty_ai_response");
  return answer;
}

async function processUpdate(update, env) {
  const msg = update.message;
  if (!msg || msg.chat?.type !== "private" || !msg.from?.id) return;
  const uid = String(msg.from.id);
  const chatId = msg.chat.id;
  const text = (msg.text || msg.caption || "").trim();
  const command = text.split(/\s+/)[0].split("@")[0].toLowerCase();

  // Anyone can ask for their own ID, but nothing else without authorization.
  if (command === "/id") {
    await reply(env, chatId, `Твой Telegram ID: ${uid}`);
    return;
  }
  const allowed = (env.ALLOWED_USER_IDS || "").split(",").map(x => x.trim());
  if (!allowed.includes(uid)) {
    await reply(env, chatId, "Доступ закрыт. Чтобы узнать свой Telegram ID, отправь /id.");
    return;
  }
  const historyKey = `history:${uid}`;
  if (command === "/start" || command === "/help") {
    await reply(env, chatId, "Привет! Отправь текст или фото с заданием — помогу. /id — твой ID; /reset — удалить свою историю.");
    return;
  }
  if (command === "/reset") {
    await env.HISTORY.delete(historyKey);
    await reply(env, chatId, "История твоей переписки удалена из хранилища бота.");
    return;
  }
  if (text.startsWith("/")) {
    await reply(env, chatId, "Команды: /start, /id, /reset, /help");
    return;
  }

  let fileId = null;
  let mime = "image/jpeg";
  if (msg.photo?.length) {
    const pic = msg.photo[msg.photo.length - 1];
    if (pic.file_size && pic.file_size > MAX_IMAGE_SIZE) {
      await reply(env, chatId, "Слишком большое фото. Отправь его как обычное сжатое фото до 2,5 МБ.");
      return;
    }
    fileId = pic.file_id;
  } else if (msg.document) {
    mime = msg.document.mime_type || "";
    if (!["image/jpeg", "image/png", "image/webp"].includes(mime)) {
      await reply(env, chatId, "Пока поддерживаются текст и изображения JPG, PNG, WEBP, но не PDF.");
      return;
    }
    if (msg.document.file_size && msg.document.file_size > MAX_IMAGE_SIZE) {
      await reply(env, chatId, "Изображение слишком большое. Пришли его как фото до 2,5 МБ.");
      return;
    }
    fileId = msg.document.file_id;
  }
  if (!text && !fileId) {
    await reply(env, chatId, "Пришли текст или фотографию задания.");
    return;
  }

  try {
    await telegram(env, "sendChatAction", { chat_id: chatId, action: "typing" });
    const prompt = text.slice(0, 8000) || "Подробно реши задания на этом фото по шагам, объясни и проверь ответ.";
    let photo = null;
    if (fileId) {
      const data = await getImage(env, fileId);
      photo = `data:${mime};base64,${imageToBase64(data)}`;
    }
    const history = (await env.HISTORY.get(historyKey, "json")) || [];
    const answer = await askModel(env, history, prompt, photo);
    const next = [...history,
      { role: "user", content: (fileId ? "[Фото] " : "") + prompt },
      { role: "assistant", content: answer.slice(0, 8000) },
    ].slice(-MAX_HISTORY);
    await env.HISTORY.put(historyKey, JSON.stringify(next));
    await reply(env, chatId, answer);
  } catch (error) {
    console.error("Bot processing error:", error?.name || "Error", error?.message || "unknown");
    let userText = "Не удалось получить ответ. Проверь баланс Tooken Club и доступность модели, потом повтори.";
    if (error?.message === "image_too_large") userText = "Фото слишком большое. Отправь как обычное фото до 2,5 МБ.";
    if (error?.name === "TimeoutError") userText = "Модель отвечает слишком долго. Попробуй ещё раз с одной фотографией.";
    await reply(env, chatId, userText);
  }
}

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    if (request.method === "GET" && url.pathname === "/") {
      return json({ status: "running", setup: "open /setup?code=YOUR_SETUP_KEY" });
    }
    if (request.method === "GET" && url.pathname === "/setup") {
      if (!env.SETUP_KEY || url.searchParams.get("code") !== env.SETUP_KEY) {
        return json({ ok: false, error: "Invalid setup code" }, 403);
      }
      if (!env.WEBHOOK_SECRET || !env.TELEGRAM_BOT_TOKEN) {
        return json({ ok: false, error: "Missing required secrets" }, 500);
      }
      try {
        const result = await telegram(env, "setWebhook", {
          url: `${url.origin}/webhook`,
          secret_token: env.WEBHOOK_SECRET,
          max_connections: 1,
          allowed_updates: ["message"],
        });
        return json({ ok: result === true, message: "Webhook connected. Send /id to the bot." });
      } catch (error) {
        console.error("Webhook setup error:", error?.message || "unknown");
        return json({ ok: false, error: "Telegram refused webhook. Check bot token." }, 502);
      }
    }
    if (request.method !== "POST" || url.pathname !== "/webhook") {
      return json({ error: "Not found" }, 404);
    }
    if (!env.WEBHOOK_SECRET || request.headers.get("X-Telegram-Bot-Api-Secret-Token") !== env.WEBHOOK_SECRET) {
      return json({ error: "Forbidden" }, 403);
    }
    if (!env.HISTORY || !env.TOOKEN_API_KEY || !env.TELEGRAM_BOT_TOKEN) {
      return json({ error: "Missing secrets or HISTORY binding" }, 500);
    }
    let update;
    try { update = await request.json(); }
    catch { return json({ error: "Invalid JSON" }, 400); }

    // KV prevents retries from billing twice. Set webhook max_connections=1.
    if (typeof update.update_id === "number") {
      const key = `update:${update.update_id}`;
      if (await env.HISTORY.get(key)) return json({ ok: true, duplicate: true });
      await env.HISTORY.put(key, "1", { expirationTtl: 86400 });
    }
    // Await work: avoids Cloudflare's 30-second limit on detached tasks.
    const task = processUpdate(update, env);
    ctx.waitUntil(task.catch(err => console.error("Background error:", err?.name || "Error")));
    await task;
    return json({ ok: true });
  },
};
