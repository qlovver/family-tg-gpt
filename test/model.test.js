import { test } from "node:test";
import assert from "node:assert/strict";
import { askModel, normalizeHistory, parseAnswer, parseSSE, ModelError } from "../src/model.js";
import { makeFetch, messagesReply, API_KEY } from "./helpers.js";

const env = { TOOKEN_API_KEY: API_KEY };
const payload = (extra = {}) => ({ system: "SYS", history: [], text: "Привет", ...extra });

test("default request is Claude Messages format on /v1/messages", async () => {
  const { fetch, calls } = makeFetch();
  const r = await askModel(env, payload({
    history: [{ role: "user", content: "q1" }, { role: "assistant", content: "a1" }],
    image: { mime: "image/jpeg", base64: "AAAA" },
  }), { fetchImpl: fetch });
  assert.equal(r.text, "Ответ");
  assert.equal(r.format, "messages");
  const c = calls.model[0];
  assert.equal(c.url, "https://tooken.club/v1/messages");
  assert.equal(c.headers.authorization, `Bearer ${API_KEY}`);
  assert.equal(c.headers["anthropic-version"], "2023-06-01");
  assert.equal(c.body.model, "claude-opus-5-5");
  assert.equal(c.body.system, "SYS");
  assert.equal(typeof c.body.max_tokens, "number");
  assert.equal(c.body.instructions, undefined, "no OpenAI Responses fields");
  assert.deepEqual(c.body.messages.slice(0, 2), [{ role: "user", content: "q1" }, { role: "assistant", content: "a1" }]);
  assert.deepEqual(c.body.messages[2], { role: "user", content: [
    { type: "image", source: { type: "base64", media_type: "image/jpeg", data: "AAAA" } },
    { type: "text", text: "Привет" },
  ] });
});

test("history is normalized to valid alternating roles", () => {
  assert.deepEqual(normalizeHistory([
    { role: "assistant", content: "orphan" },
    { role: "user", content: "a" }, { role: "user", content: "b" },
    { role: "assistant", content: "c" }, { role: "system", content: "x" },
    { role: "user", content: "dangling" },
  ]), [{ role: "user", content: "a\n\nb" }, { role: "assistant", content: "c" }]);
});

test("parses Messages, Chat Completions and Responses shapes", () => {
  assert.equal(parseAnswer(messagesReply("A")).text, "A");
  assert.equal(parseAnswer({ content: [{ type: "thinking", thinking: "..." }, { type: "text", text: "B" }] }).text, "B");
  assert.equal(parseAnswer({ choices: [{ message: { content: "C" } }] }).text, "C");
  assert.equal(parseAnswer({ output_text: "D" }).text, "D");
  assert.equal(parseAnswer({ output: [{ content: [{ type: "output_text", text: "E" }] }] }).text, "E");
  assert.equal(parseAnswer(messagesReply("F", { stop_reason: "max_tokens" })).truncated, true);
});

test("parses streamed SSE responses", async () => {
  const sse = [
    'event: message_start', 'data: {"type":"message_start","message":{"model":"claude-opus-5-5","usage":{"input_tokens":7}}}', '',
    'event: content_block_delta', 'data: {"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"При"}}', '',
    'data: {"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"вет"}}', '',
    'data: {"type":"message_delta","delta":{"stop_reason":"end_turn"},"usage":{"output_tokens":3}}', '',
  ].join("\n");
  assert.equal(parseSSE(sse).text, "Привет");
  const { fetch } = makeFetch({ model: () => ({ text: sse, contentType: "text/event-stream" }) });
  const r = await askModel({ ...env, MODEL_STREAM: "1" }, payload(), { fetchImpl: fetch });
  assert.equal(r.text, "Привет");
  assert.equal(r.usage.output_tokens, 3);
});

test("falls back to chat completions only when the endpoint is missing", async () => {
  const { fetch, calls } = makeFetch({
    model: url => url.endsWith("/messages")
      ? { status: 404, json: { error: { message: "Not Found" } } }
      : { json: { choices: [{ message: { content: "через chat" } }] } },
  });
  const r = await askModel(env, payload({ image: { mime: "image/png", base64: "QQ==" } }), { fetchImpl: fetch });
  assert.equal(r.text, "через chat");
  assert.equal(r.format, "chat");
  assert.equal(calls.model.length, 2);
  const chatBody = calls.model[1].body;
  assert.equal(chatBody.messages[0].role, "system");
  assert.equal(chatBody.messages[1].content[1].image_url.url, "data:image/png;base64,QQ==");
});

test("API_FORMAT=responses keeps the old Responses format", async () => {
  const { fetch, calls } = makeFetch({ model: () => ({ json: { output_text: "ok" } }) });
  await askModel({ ...env, API_FORMAT: "responses" }, payload(), { fetchImpl: fetch });
  assert.equal(calls.model[0].url, "https://tooken.club/v1/responses");
  assert.equal(calls.model[0].body.instructions, "SYS");
});

test("errors are classified and not retried when billing could happen", async () => {
  const cases = [
    [401, { error: { code: "invalid_api_key" } }, "auth_failed"],
    [402, { error: { code: "insufficient_quota" } }, "no_balance"],
    [404, { error: { code: "model_not_found", message: "model x" } }, "model_not_found"],
    [400, { error: { code: "context_length_exceeded" } }, "context_too_long"],
    [504, { error: { code: "upstream_timeout" } }, "upstream_timeout"],
    [400, { error: { message: "bad" } }, "bad_request"],
  ];
  for (const [status, json, code] of cases) {
    const { fetch, calls } = makeFetch({ model: () => ({ status, json }) });
    await assert.rejects(askModel(env, payload(), { fetchImpl: fetch }), e => e instanceof ModelError && e.code === code);
    assert.equal(calls.model.length, 1, `${code} must not retry`);
  }
});

test("transient 503 is retried once", async () => {
  let n = 0;
  const { fetch, calls } = makeFetch({ model: () => (++n === 1 ? { status: 503, json: { error: { message: "overloaded" } } } : { json: messagesReply("ok") }) });
  const r = await askModel(env, payload(), { fetchImpl: fetch });
  assert.equal(r.text, "ok");
  assert.equal(calls.model.length, 2);
});

test("timeout is reported as timeout", async () => {
  const fetchImpl = (_url, init) => new Promise((_, reject) => {
    init.signal.addEventListener("abort", () => reject(init.signal.reason));
  });
  await assert.rejects(askModel(env, payload(), { fetchImpl, timeoutMs: 50 }), e => e.code === "timeout");
});

test("empty answer is an error", async () => {
  const { fetch } = makeFetch({ model: () => ({ json: messagesReply("   ") }) });
  await assert.rejects(askModel(env, payload(), { fetchImpl: fetch }), e => e.code === "empty_response");
});

test("capability probe detects tool calls and vision", async () => {
  const { probeCapabilities } = await import("../src/model.js");
  const fetchImpl = async (url, init) => {
    const body = JSON.parse(init.body);
    const isTools = !!body.tools;
    const data = url.endsWith("/messages")
      ? (isTools ? { model: "claude-opus-5-5", id: "msg_1", content: [{ type: "tool_use", name: "get_weather", input: { city: "Москва" } }] }
        : { content: [{ type: "text", text: "Красный круг на белом фоне" }] })
      : (isTools ? { model: "gpt-6.1-sol", choices: [{ message: { content: "Не знаю" } }] }
        : { choices: [{ message: { content: "Белый квадрат" } }] });
    return Response.json(data);
  };
  const c = await probeCapabilities({ TOOKEN_API_KEY: "k" }, "claude-opus-5-5", { fetchImpl });
  assert.match(c.tools, /✅ вызвал get_weather/);
  assert.match(c.vision, /✅/);
  const g = await probeCapabilities({ TOOKEN_API_KEY: "k" }, "gpt-6.1-sol", { fetchImpl });
  assert.match(g.tools, /❌ ответил без инструмента/);
  assert.match(g.vision, /❌ не видит/);
});
