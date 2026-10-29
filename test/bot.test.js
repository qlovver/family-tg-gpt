import { test } from "node:test";
import assert from "node:assert/strict";
import { createWorker, pickPhoto, trimHistory, MAX_HISTORY, historyKey } from "../src/bot.js";
import {
  makeEnv, makeFetch, messagesReply, textUpdate, photoUpdate, webhookRequest, execCtx, sentTexts, MemoryKV,
} from "./helpers.js";

async function send(env, f, update, secret) {
  const worker = createWorker({ fetchImpl: f.fetch });
  return worker.fetch(webhookRequest(update, secret), env, execCtx);
}

test("rejects webhook calls with a wrong secret", async () => {
  const env = makeEnv(); const f = makeFetch();
  const res = await send(env, f, textUpdate(111, "hi"), "wrong");
  assert.equal(res.status, 403);
  assert.equal(f.calls.telegram.length, 0);
});

test("/id works for anyone, even without KV and API key", async () => {
  const env = makeEnv({ HISTORY: undefined, TOOKEN_API_KEY: undefined }); const f = makeFetch();
  const res = await send(env, f, textUpdate(999, "/id"));
  assert.equal(res.status, 200);
  assert.match(sentTexts(f.calls)[0], /<code>999<\/code>/);
});

test("strangers get no model access and touch no storage", async () => {
  const env = makeEnv(); const f = makeFetch();
  await send(env, f, textUpdate(999, "реши задачу"));
  await send(env, f, textUpdate(999, "/reset"));
  assert.equal(f.calls.model.length, 0);
  assert.equal(env.HISTORY.writes, 0);
  assert.match(sentTexts(f.calls)[0], /приватный бот/);
});

test("/start answers with help", async () => {
  const env = makeEnv(); const f = makeFetch();
  await send(env, f, textUpdate(111, "/start"));
  assert.match(sentTexts(f.calls)[0], /Привет/);
  assert.equal(f.calls.telegram.find(c => c.method === "sendMessage").body.parse_mode, "HTML");
});

test("text message: model called, answer formatted as HTML, history saved", async () => {
  const env = makeEnv();
  const f = makeFetch({ model: () => ({ json: messagesReply("### Итог\n**Ответ:** $\\sqrt{36} = 6$") }) });
  const res = await send(env, f, textUpdate(111, "Сколько будет корень из 36?"));
  assert.equal(res.status, 200);
  const msg = f.calls.telegram.filter(c => c.method === "sendMessage").pop();
  assert.equal(msg.body.parse_mode, "HTML");
  assert.equal(msg.body.text, "<b>Итог</b>\n<b>Ответ:</b> √36 = 6");
  const rec = JSON.parse(env.HISTORY.map.get(historyKey("111")));
  assert.equal(rec.messages.length, 2);
  assert.equal(rec.pending, undefined);
  assert.equal(env.HISTORY.writes, 2, "only 2 KV writes per message (free plan has 1000/day)");
});

test("histories are separate per user and /reset clears only the caller", async () => {
  const env = makeEnv();
  const f = makeFetch({ model: (_u, body) => ({ json: messagesReply("echo:" + body.messages.at(-1).content[0].text) }) });
  await send(env, f, textUpdate(111, "секрет первого"));
  await send(env, f, textUpdate(222, "вопрос второго"));
  const second = f.calls.model[1].body.messages;
  assert.equal(second.length, 1, "user 222 sees no history of user 111");
  assert.ok(!JSON.stringify(second).includes("секрет первого"));

  await send(env, f, textUpdate(222, "/reset"));
  assert.equal(env.HISTORY.map.has(historyKey("222")), false);
  assert.equal(env.HISTORY.map.has(historyKey("111")), true);

  await send(env, f, textUpdate(111, "продолжи"));
  assert.equal(f.calls.model[2].body.messages.length, 3, "user 111 keeps context");
});

test("old array-format history is still read", async () => {
  const env = makeEnv();
  env.HISTORY.map.set(historyKey("111"), JSON.stringify([{ role: "user", content: "старый" }, { role: "assistant", content: "ответ" }]));
  const f = makeFetch();
  await send(env, f, textUpdate(111, "новый"));
  assert.equal(f.calls.model[0].body.messages[0].content, "старый");
  const rec = JSON.parse(env.HISTORY.map.get(historyKey("111")));
  assert.equal(rec.messages.length, 4);
});

test("photo: picks a moderate size, sends image block, stores marker in history", async () => {
  const env = makeEnv(); const f = makeFetch();
  await send(env, f, photoUpdate(111, "реши пункт б"));
  assert.equal(f.calls.telegram.find(c => c.method === "getFile").body.file_id, "mid");
  const content = f.calls.model[0].body.messages.at(-1).content;
  assert.equal(content[0].type, "image");
  assert.equal(content[0].source.media_type, "image/jpeg");
  assert.equal(content[1].text, "реши пункт б");
  const rec = JSON.parse(env.HISTORY.map.get(historyKey("111")));
  assert.equal(rec.messages[0].content, "[Фото задания] реши пункт б");
  assert.ok(!JSON.stringify(rec).includes("base64"), "images are not stored in KV");
});

test("photo without caption gets a default solving prompt", async () => {
  const env = makeEnv(); const f = makeFetch();
  await send(env, f, photoUpdate(111));
  assert.match(f.calls.model[0].body.messages.at(-1).content[1].text, /Реши задание на фото/);
});

test("reply to an old photo re-uses that photo", async () => {
  const env = makeEnv(); const f = makeFetch();
  const old = photoUpdate(111).message;
  await send(env, f, textUpdate(111, "а пункт в?", { reply_to_message: old }));
  assert.equal(f.calls.model[0].body.messages.at(-1).content[0].type, "image");
});

test("old .doc files are refused politely without calling the model", async () => {
  const env = makeEnv(); const f = makeFetch();
  await send(env, f, textUpdate(111, undefined, { document: { file_id: "d", file_name: "курсовая.doc", mime_type: "application/msword" } }));
  assert.equal(f.calls.model.length, 0);
  assert.match(sentTexts(f.calls)[0], /сохрани как \.docx/);
});

test("model failure: user gets a clear message and /retry repeats the request", async () => {
  const env = makeEnv();
  let fail = true;
  const f = makeFetch({ model: () => (fail ? { status: 402, json: { error: { code: "insufficient_quota" } } } : { json: messagesReply("решено") }) });
  const res = await send(env, f, textUpdate(111, "задача 1"));
  assert.equal(res.status, 200, "no Telegram redelivery (would bill again)");
  assert.match(sentTexts(f.calls).at(-1), /баланс/);
  const rec = JSON.parse(env.HISTORY.map.get(historyKey("111")));
  assert.equal(rec.pending.text, "задача 1", "request kept for /retry");
  assert.equal(rec.messages.length, 0);

  fail = false;
  await send(env, f, textUpdate(111, "/retry"));
  assert.equal(f.calls.model.at(-1).body.messages.at(-1).content[0].text, "задача 1");
  assert.equal(sentTexts(f.calls).at(-1), "решено");
  const rec2 = JSON.parse(env.HISTORY.map.get(historyKey("111")));
  assert.equal(rec2.pending, undefined);
  assert.equal(rec2.messages.length, 2);
});

test("duplicate delivery of the same update does not call the model twice", async () => {
  const env = makeEnv(); const f = makeFetch();
  const u = textUpdate(111, "один раз");
  await send(env, f, u);
  await send(env, f, u);
  assert.equal(f.calls.model.length, 1);
});

test("redelivery while still processing is ignored", async () => {
  const env = makeEnv();
  let release;
  const f = makeFetch({ model: () => new Promise(r => { release = () => r({ json: messagesReply("готово") }); }) });
  const u = textUpdate(111, "долгая задача");
  const first = send(env, f, u);
  await new Promise(r => setTimeout(r, 20));
  const second = await send(env, f, u);
  assert.equal(second.status, 200);
  release();
  await first;
  assert.equal(f.calls.model.length, 1);
});

test("if Telegram fails to deliver the answer, the update is retried without a new model call", async () => {
  const env = makeEnv();
  let tgDown = true;
  const f = makeFetch({ telegram: method => (method === "sendMessage" && tgDown ? { status: 502, json: { ok: false, description: "Bad Gateway" } } : null) });
  const u = textUpdate(111, "вопрос");
  const res = await send(env, f, u);
  assert.equal(res.status, 500, "ask Telegram to redeliver");
  tgDown = false;
  const res2 = await send(env, f, u);
  assert.equal(res2.status, 200);
  assert.equal(f.calls.model.length, 1);
  assert.equal(sentTexts(f.calls).at(-1), "Ответ");
});

test("HTML rejected by Telegram falls back to plain text", async () => {
  const env = makeEnv();
  const f = makeFetch({
    model: () => ({ json: messagesReply("**жирный** a < b") }),
    telegram: (method, body) => (method === "sendMessage" && body.parse_mode === "HTML" && body.text.includes("жирный")
      ? { status: 400, json: { ok: false, description: "Bad Request: can't parse entities" } } : null),
  });
  await send(env, f, textUpdate(111, "x"));
  const last = f.calls.telegram.filter(c => c.method === "sendMessage").at(-1);
  assert.equal(last.body.parse_mode, undefined);
  assert.equal(last.body.text, "жирный a < b");
});

test("long answers are split into several messages, nothing lost", async () => {
  const env = makeEnv();
  const long = Array.from({ length: 300 }, (_, i) => `**Шаг ${i}.** ${"текст ".repeat(10)}`).join("\n\n");
  const f = makeFetch({ model: () => ({ json: messagesReply(long) }) });
  await send(env, f, textUpdate(111, "длинно"));
  const msgs = f.calls.telegram.filter(c => c.method === "sendMessage");
  assert.ok(msgs.length >= 4);
  for (const m of msgs) assert.ok(m.body.text.replace(/<[^>]+>/g, "").length <= 4096);
  const all = msgs.map(m => m.body.text).join("\n");
  assert.ok(all.includes("Шаг 0.") && all.includes("Шаг 299."));
});

test("missing KV binding gives a clear message instead of a silent 500", async () => {
  const env = makeEnv({ HISTORY: undefined }); const f = makeFetch();
  const res = await send(env, f, textUpdate(111, "привет"));
  assert.equal(res.status, 200);
  assert.match(sentTexts(f.calls)[0], /HISTORY/);
});

test("history is bounded by count and size", () => {
  const msgs = [];
  for (let i = 0; i < 40; i++) msgs.push({ role: i % 2 ? "assistant" : "user", content: "x".repeat(3000) });
  const t = trimHistory(msgs);
  assert.ok(t.length <= MAX_HISTORY);
  assert.equal(t[0].role, "user");
  assert.ok(t.reduce((n, m) => n + m.content.length, 0) <= 20000);
});

test("pickPhoto prefers the largest size up to 1600px", () => {
  assert.equal(pickPhoto(photoUpdate(1).message.photo).file_id, "mid");
  assert.equal(pickPhoto([{ file_id: "only", width: 3000, height: 4000, file_size: 100 }]).file_id, "only");
  assert.equal(pickPhoto([{ file_id: "big", width: 3000, height: 4000, file_size: 9_000_000 }]), null);
});

test("status page reports config without leaking secrets", async () => {
  const env = makeEnv();
  const res = await createWorker({ fetchImpl: makeFetch().fetch }).fetch(new Request("https://x.dev/"), env, execCtx);
  const text = await res.text();
  assert.match(text, /"status": "running"/);
  for (const secret of [env.TELEGRAM_BOT_TOKEN, env.TOOKEN_API_KEY, env.WEBHOOK_SECRET, env.SETUP_KEY, "111"]) {
    assert.ok(!text.includes(secret), "leaked " + secret);
  }
});

test("setup: key only via POST form; sets webhook with secret", async () => {
  const env = makeEnv(); const f = makeFetch({
    telegram: m => (m === "getWebhookInfo" ? { json: { ok: true, result: { url: "https://x.dev/webhook", pending_update_count: 0 } } }
      : m === "getMe" ? { json: { ok: true, result: { username: "family_bot" } } } : null),
  });
  const worker = createWorker({ fetchImpl: f.fetch });
  const get = await worker.fetch(new Request("https://x.dev/setup?code=setup-key"), env, execCtx);
  assert.equal(f.calls.telegram.length, 0, "GET with code in URL does nothing");
  assert.match(await get.text(), /не передаётся в адресе/);

  const bad = await worker.fetch(new Request("https://x.dev/setup", { method: "POST", body: new URLSearchParams({ key: "nope", webhook: "on" }) }), env, execCtx);
  assert.equal(bad.status, 403);

  const ok = await worker.fetch(new Request("https://x.dev/setup", { method: "POST", body: new URLSearchParams({ key: "setup-key", webhook: "on" }) }), env, execCtx);
  assert.equal(ok.status, 200);
  const sw = f.calls.telegram.find(c => c.method === "setWebhook");
  assert.equal(sw.body.url, "https://x.dev/webhook");
  assert.equal(sw.body.secret_token, env.WEBHOOK_SECRET);
  const page = await ok.text();
  assert.ok(!page.includes(env.WEBHOOK_SECRET) && !page.includes(env.TELEGRAM_BOT_TOKEN));
});

test("setup: model check reports endpoint and model name", async () => {
  const env = makeEnv(); const f = makeFetch({ model: () => ({ json: messagesReply("Работает") }) });
  const worker = createWorker({ fetchImpl: f.fetch });
  const res = await worker.fetch(new Request("https://x.dev/setup", { method: "POST", body: new URLSearchParams({ key: "setup-key", model: "on" }) }), env, execCtx);
  const page = await res.text();
  assert.match(page, /"endpoint": "messages"/);
  assert.match(page, /claude-opus-5-5/);
  assert.equal(f.calls.model[0].body.max_tokens, 20, "cheap test request");
});

test("non-private chats and edited messages are ignored", async () => {
  const env = makeEnv(); const f = makeFetch();
  await send(env, f, { update_id: 1, message: { message_id: 1, from: { id: 111 }, chat: { id: -5, type: "group" }, text: "hi" } });
  await send(env, f, { update_id: 2, edited_message: { message_id: 1, from: { id: 111 }, chat: { id: 111, type: "private" }, text: "hi" } });
  assert.equal(f.calls.telegram.length, 0);
});

test("MemoryKV sanity", async () => {
  const kv = new MemoryKV(); await kv.put("a", "1"); assert.equal(await kv.get("a"), "1");
});

test("model choice: /model menu, button press, per-user model in requests", async () => {
  const env = makeEnv();
  const f = makeFetch({
    telegram: m => (m === "getWebhookInfo" ? { json: { ok: true, result: { url: "https://x.dev/webhook", allowed_updates: ["message"], max_connections: 1 } } } : null),
  });
  // Tooken model list
  const orig = f.fetch;
  const fetch2 = async (url, init) => String(url).endsWith("/v1/models")
    ? Response.json({ data: [{ id: "claude-opus-5-5" }, { id: "gpt-6.1-sol" }, { id: "gpt-image-2.5-sunburst" }] })
    : orig(url, init);
  const worker = createWorker({ fetchImpl: fetch2 });
  const go = u => worker.fetch(webhookRequest(u), env, execCtx);

  await go(textUpdate(111, "/model"));
  const menu = f.calls.telegram.find(c => c.method === "sendMessage" && c.body.reply_markup);
  const buttons = menu.body.reply_markup.inline_keyboard.flat();
  assert.deepEqual(buttons.map(b => b.callback_data), ["m:claude-opus-5-5", "m:gpt-6.1-sol"], "image models hidden");
  assert.match(buttons[0].text, /✅ Claude Opus 5.5/);
  assert.ok(f.calls.telegram.some(c => c.method === "setWebhook" && c.body.allowed_updates.includes("callback_query")), "webhook upgraded for buttons");

  await go({ update_id: 777001, callback_query: { id: "cq1", from: { id: 111 }, data: "m:gpt-6.1-sol", message: { message_id: 5, chat: { id: 111 } } } });
  assert.ok(f.calls.telegram.some(c => c.method === "answerCallbackQuery"));
  assert.ok(f.calls.telegram.some(c => c.method === "editMessageText" && /GPT-6.1 Sol/.test(c.body.text)));

  await go(textUpdate(111, "привет"));
  const req = f.calls.model.at(-1);
  assert.equal(req.body.model, "gpt-6.1-sol");
  assert.equal(req.url, "https://tooken.club/v1/chat/completions", "GPT goes to chat completions");

  // the other user keeps the default
  await go(textUpdate(222, "привет"));
  assert.equal(f.calls.model.at(-1).body.model, "claude-opus-5-5");
  assert.equal(f.calls.model.at(-1).url, "https://tooken.club/v1/messages");

  // /reset keeps the choice, strangers can't press buttons
  await go(textUpdate(111, "/reset"));
  await go(textUpdate(111, "ещё раз"));
  assert.equal(f.calls.model.at(-1).body.model, "gpt-6.1-sol");
  await go({ update_id: 777002, callback_query: { id: "cq2", from: { id: 999 }, data: "m:claude-opus-5-5" } });
  assert.equal(JSON.parse(env.HISTORY.map.get(historyKey("111"))).model, "gpt-6.1-sol");

  // /model with an argument
  await go(textUpdate(111, "/model opus"));
  assert.equal(JSON.parse(env.HISTORY.map.get(historyKey("111"))).model, "claude-opus-5-5");
});

test("model list falls back to the built-in catalogue when Tooken doesn't return it", async () => {
  const env = makeEnv(); const f = makeFetch({ model: (url) => ({ status: 404, json: {} }) });
  await send(env, f, textUpdate(111, "/model"));
  const menu = f.calls.telegram.find(c => c.method === "sendMessage" && c.body.reply_markup);
  assert.ok(menu.body.reply_markup.inline_keyboard.flat().length >= 10);
});
