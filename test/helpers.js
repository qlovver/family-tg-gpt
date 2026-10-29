// Test doubles: in-memory KV and a fake fetch for Telegram + Tooken.

export class MemoryKV {
  constructor() { this.map = new Map(); this.writes = 0; }
  async get(key, type) {
    const v = this.map.get(key);
    if (v === undefined) return null;
    return type === "json" ? JSON.parse(v) : v;
  }
  async put(key, value) { this.writes++; this.map.set(key, String(value)); }
  async delete(key) { this.map.delete(key); }
}

export const BOT_TOKEN = "123:TEST_TOKEN";
export const API_KEY = "tc_live_TEST";
export const SECRET = "hook-secret";

export function makeEnv(extra = {}) {
  return {
    TELEGRAM_BOT_TOKEN: BOT_TOKEN,
    TOOKEN_API_KEY: API_KEY,
    WEBHOOK_SECRET: SECRET,
    SETUP_KEY: "setup-key",
    ALLOWED_USER_IDS: "111, 222",
    HISTORY: new MemoryKV(),
    ...extra,
  };
}

/**
 * Fake fetch. `model` is a function (url, body, init) => Response-like
 * description: { status, json } | { status, text, contentType }.
 */
export function makeFetch({ model, telegram } = {}) {
  const calls = { telegram: [], model: [], files: [] };
  const fn = async (url, init = {}) => {
    url = String(url);
    if (url.startsWith("https://api.telegram.org/file/")) {
      calls.files.push(url);
      const jpeg = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 1, 2, 3, 4]);
      return new Response(jpeg, { status: 200 });
    }
    if (url.startsWith("https://api.telegram.org/bot")) {
      const method = url.split("/").pop();
      const body = init.body ? JSON.parse(init.body) : {};
      calls.telegram.push({ method, body });
      const custom = telegram?.(method, body);
      if (custom) return new Response(JSON.stringify(custom.json), { status: custom.status || 200 });
      if (method === "getFile") return Response.json({ ok: true, result: { file_path: "photos/a.jpg", file_size: 8 } });
      return Response.json({ ok: true, result: method === "sendMessage" ? { message_id: 1 } : true });
    }
    if (url.startsWith("https://tooken.club/")) {
      const body = JSON.parse(init.body);
      calls.model.push({ url, body, headers: init.headers });
      const r = model ? await model(url, body, init) : { json: messagesReply("Ответ") };
      if (r.throw) throw r.throw;
      const headers = { "content-type": r.contentType || "application/json", "x-request-id": "req_1" };
      return new Response(r.text ?? JSON.stringify(r.json), { status: r.status || 200, headers });
    }
    throw new Error("unexpected fetch " + url);
  };
  return { fetch: fn, calls };
}

export function messagesReply(text, extra = {}) {
  return {
    id: "msg_1", type: "message", role: "assistant", model: "claude-opus-5-5",
    content: [{ type: "text", text }], stop_reason: "end_turn",
    usage: { input_tokens: 10, output_tokens: 5 }, ...extra,
  };
}

let updateSeq = 1000;
export function textUpdate(uid, text, extra = {}) {
  return {
    update_id: ++updateSeq,
    message: { message_id: updateSeq, from: { id: uid }, chat: { id: uid, type: "private" }, text, ...extra },
  };
}

export function photoUpdate(uid, caption) {
  return {
    update_id: ++updateSeq,
    message: {
      message_id: updateSeq, from: { id: uid }, chat: { id: uid, type: "private" }, caption,
      photo: [
        { file_id: "small", width: 90, height: 120, file_size: 2000 },
        { file_id: "mid", width: 960, height: 1280, file_size: 150000 },
        { file_id: "huge", width: 1920, height: 2560, file_size: 900000 },
      ],
    },
  };
}

export function webhookRequest(update, secret = SECRET) {
  return new Request("https://family-tg-gpt.example.workers.dev/webhook", {
    method: "POST",
    headers: { "content-type": "application/json", "X-Telegram-Bot-Api-Secret-Token": secret },
    body: JSON.stringify(update),
  });
}

export const execCtx = { waitUntil() {} };

export function sentTexts(calls) {
  return calls.telegram.filter(c => c.method === "sendMessage").map(c => c.body.text);
}
