// Client for Tooken Club. For claude-* models Tooken documents the
// Claude Messages API format, so the default is POST {base}/messages.
// "chat" (OpenAI Chat Completions) and "responses" (OpenAI Responses)
// are kept as configurable alternatives via the API_FORMAT variable.

export const DEFAULT_MODEL = "claude-opus-5-5";
export const DEFAULT_BASE_URL = "https://tooken.club/v1";
export const DEFAULT_MAX_TOKENS = 4000;
export const DEFAULT_TIMEOUT_MS = 85_000;

export class ModelError extends Error {
  constructor(code, detail = "", status = 0, requestId = "") {
    super(code);
    this.name = "ModelError";
    this.code = code;
    this.detail = detail;
    this.status = status;
    this.requestId = requestId;
  }
}

// Which API format a model uses at Tooken: Claude -> Messages, GPT-6.1 Terra ->
// Responses (as the catalogue says), everything else -> Chat Completions.
export function formatFor(model, env = {}) {
  const forced = String(env.API_FORMAT || "").toLowerCase();
  if (["messages", "chat", "responses"].includes(forced)) return forced;
  if (/^claude/i.test(model)) return "messages";
  if (/terra/i.test(model)) return "responses";
  return "chat";
}

export function modelConfig(env = {}, modelOverride) {
  const model = String(modelOverride || env.MODEL || DEFAULT_MODEL);
  return {
    baseUrl: String(env.TOOKEN_BASE_URL || DEFAULT_BASE_URL).replace(/\/+$/, ""),
    model,
    format: formatFor(model, env),
    stream: env.MODEL_STREAM === "1" || env.MODEL_STREAM === "true",
    maxTokens: Number(env.MAX_OUTPUT_TOKENS) > 0 ? Number(env.MAX_OUTPUT_TOKENS) : DEFAULT_MAX_TOKENS,
  };
}

/* ---------------------------- request bodies ---------------------------- */

// Claude requires alternating roles starting with "user". Repairs history
// so a broken/old record can never make the request invalid.
export function normalizeHistory(history) {
  const out = [];
  for (const m of Array.isArray(history) ? history : []) {
    if (!m || (m.role !== "user" && m.role !== "assistant")) continue;
    const content = typeof m.content === "string" ? m.content.trim() : "";
    if (!content) continue;
    if (!out.length && m.role !== "user") continue;
    const last = out[out.length - 1];
    if (last && last.role === m.role) last.content += "\n\n" + content;
    else out.push({ role: m.role, content });
  }
  // The new user message is appended next, so history must end with assistant.
  if (out.length && out[out.length - 1].role === "user") out.pop();
  return out;
}

export function buildMessagesBody(cfg, system, history, text, image) {
  const content = [];
  if (image) {
    // PDFs go as a "document" block, pictures as an "image" block.
    content.push(image.mime === "application/pdf"
      ? { type: "document", source: { type: "base64", media_type: "application/pdf", data: image.base64 }, ...(image.name ? { title: image.name } : {}) }
      : { type: "image", source: { type: "base64", media_type: image.mime, data: image.base64 } });
  }
  content.push({ type: "text", text });
  return {
    model: cfg.model,
    max_tokens: cfg.maxTokens,
    system,
    messages: [...normalizeHistory(history), { role: "user", content }],
    ...(cfg.stream ? { stream: true } : {}),
  };
}

export function buildChatBody(cfg, system, history, text, image) {
  const content = image
    ? [{ type: "text", text }, image.mime === "application/pdf"
      ? { type: "file", file: { filename: image.name || "document.pdf", file_data: `data:application/pdf;base64,${image.base64}` } }
      : { type: "image_url", image_url: { url: `data:${image.mime};base64,${image.base64}` } }]
    : text;
  return {
    model: cfg.model,
    max_tokens: cfg.maxTokens,
    messages: [{ role: "system", content: system }, ...normalizeHistory(history), { role: "user", content }],
    ...(cfg.stream ? { stream: true } : {}),
  };
}

export function buildResponsesBody(cfg, system, history, text, image) {
  const content = image
    ? [{ type: "input_text", text }, image.mime === "application/pdf"
      ? { type: "input_file", filename: image.name || "document.pdf", file_data: `data:application/pdf;base64,${image.base64}` }
      : { type: "input_image", image_url: `data:${image.mime};base64,${image.base64}` }]
    : text;
  return {
    model: cfg.model,
    instructions: system,
    input: [...normalizeHistory(history), { role: "user", content }],
    max_output_tokens: cfg.maxTokens,
  };
}

const ENDPOINTS = {
  messages: { path: "/messages", build: buildMessagesBody },
  chat: { path: "/chat/completions", build: buildChatBody },
  responses: { path: "/responses", build: buildResponsesBody },
};

/* ---------------------------- response parsing -------------------------- */

// Accepts Messages, Chat Completions and Responses shapes.
export function parseAnswer(data) {
  if (!data || typeof data !== "object") return { text: "" };
  // Claude Messages: { content: [{type:"text", text}], stop_reason, usage }
  if (Array.isArray(data.content)) {
    const text = data.content.filter(b => b && b.type === "text" && typeof b.text === "string")
      .map(b => b.text).join("").trim();
    return { text, model: data.model, usage: data.usage, truncated: data.stop_reason === "max_tokens" };
  }
  // Chat Completions
  if (Array.isArray(data.choices)) {
    const c = data.choices[0] || {};
    let text = c.message?.content;
    if (Array.isArray(text)) text = text.map(p => p?.text || "").join("");
    return { text: String(text || "").trim(), model: data.model, usage: data.usage, truncated: c.finish_reason === "length" };
  }
  // Responses API
  if (typeof data.output_text === "string" && data.output_text.trim()) {
    return { text: data.output_text.trim(), model: data.model, usage: data.usage, truncated: data.status === "incomplete" };
  }
  const parts = [];
  for (const item of data.output || []) {
    for (const c of item?.content || []) if (typeof c?.text === "string") parts.push(c.text);
  }
  return { text: parts.join("\n").trim(), model: data.model, usage: data.usage, truncated: data.status === "incomplete" };
}

// Parses an SSE body (Messages or Chat Completions streaming).
export function parseSSE(body) {
  let text = "", model, usage = {}, truncated = false;
  for (const line of String(body).split(/\r?\n/)) {
    if (!line.startsWith("data:")) continue;
    const payload = line.slice(5).trim();
    if (!payload || payload === "[DONE]") continue;
    let ev;
    try { ev = JSON.parse(payload); } catch { continue; }
    if (ev.type === "error") throw new ModelError("provider_failed", ev.error?.message || "stream error");
    if (ev.type === "message_start") { model = ev.message?.model; Object.assign(usage, ev.message?.usage); }
    if (ev.type === "content_block_delta" && ev.delta?.type === "text_delta") text += ev.delta.text || "";
    if (ev.type === "message_delta") {
      Object.assign(usage, ev.usage);
      if (ev.delta?.stop_reason === "max_tokens") truncated = true;
    }
    if (Array.isArray(ev.choices)) {
      model = model || ev.model;
      const ch = ev.choices[0] || {};
      if (typeof ch.delta?.content === "string") text += ch.delta.content;
      if (ch.finish_reason === "length") truncated = true;
      if (ev.usage) Object.assign(usage, ev.usage);
    }
  }
  return { text: text.trim(), model, usage, truncated };
}

/* ---------------------------- errors ------------------------------------ */

function classify(status, body) {
  let code = "", message = "";
  try {
    const j = JSON.parse(body);
    code = String(j?.error?.code || j?.error?.type || j?.code || "");
    message = String(j?.error?.message || j?.message || "");
  } catch { message = String(body || ""); }
  const c = code.toLowerCase(), m = message.toLowerCase();
  if (status === 401 || status === 403 || c.includes("invalid_api_key") || c.includes("authentication")) return ["auth_failed", code, message];
  if (status === 402 || c.includes("insufficient_quota") || m.includes("balance")) return ["no_balance", code, message];
  if (status === 429 || c.includes("rate_limit")) return ["rate_limited", code, message];
  if (c.includes("model_not_found") || (status === 404 && m.includes("model"))) return ["model_not_found", code, message];
  if (status === 404 || status === 405) return ["endpoint_not_found", code, message];
  if (c.includes("context_length") || m.includes("context") && m.includes("length")) return ["context_too_long", code, message];
  if (status === 413) return ["request_too_large", code, message];
  if (status === 408 || status === 504 || c.includes("timeout")) return ["upstream_timeout", code, message];
  if (status === 400 || status === 422) return ["bad_request", code, message];
  if (status >= 500 || c.includes("upstream") || c.includes("overloaded")) return ["provider_unavailable", code, message];
  return ["provider_failed", code, message];
}

const RETRYABLE = new Set(["provider_unavailable", "rate_limited"]);

/* ---------------------------- main call --------------------------------- */

async function callOnce(cfg, format, env, payload, fetchImpl, signal) {
  const ep = ENDPOINTS[format];
  const body = ep.build(cfg, payload.system, payload.history, payload.text, payload.image);
  const headers = {
    "content-type": "application/json",
    // Tooken documents Bearer auth; x-api-key/anthropic-version are what
    // Claude Messages clients normally send, harmless if ignored.
    authorization: `Bearer ${env.TOOKEN_API_KEY}`,
  };
  if (format === "messages") {
    headers["x-api-key"] = env.TOOKEN_API_KEY;
    headers["anthropic-version"] = "2023-06-01";
  }
  const res = await fetchImpl(cfg.baseUrl + ep.path, { method: "POST", headers, body: JSON.stringify(body), signal });
  const requestId = res.headers?.get?.("x-request-id") || res.headers?.get?.("request-id") || "";
  const raw = await res.text();
  if (!res.ok) {
    const [code, providerCode, message] = classify(res.status, raw);
    throw new ModelError(code, `${providerCode} ${message}`.trim().slice(0, 300), res.status, requestId);
  }
  const isSSE = (res.headers?.get?.("content-type") || "").includes("text/event-stream") || raw.startsWith("event:") || raw.startsWith("data:");
  let parsed;
  if (isSSE) parsed = parseSSE(raw);
  else {
    let data;
    try { data = JSON.parse(raw); } catch { throw new ModelError("bad_response", "non-JSON response", res.status, requestId); }
    if (data?.error) {
      const [code, providerCode, message] = classify(res.status, raw);
      throw new ModelError(code, `${providerCode} ${message}`.trim().slice(0, 300), res.status, requestId);
    }
    parsed = parseAnswer(data);
  }
  if (!parsed.text) throw new ModelError("empty_response", "", res.status, requestId);
  return { ...parsed, format, requestId };
}

/**
 * payload: { system, history: [{role, content}], text, image?: {mime, base64} }
 * returns: { text, model, usage, truncated, format, requestId }
 */
export async function askModel(env, payload, { fetchImpl = fetch, timeoutMs = DEFAULT_TIMEOUT_MS } = {}) {
  if (!env.TOOKEN_API_KEY) throw new ModelError("auth_failed", "TOOKEN_API_KEY is not set");
  const cfg = modelConfig(env, payload.model);
  if (Number(payload.maxTokens) > 0) cfg.maxTokens = Number(payload.maxTokens);
  const started = Date.now();
  // AbortController + setTimeout (works in Workers and Node; cleared on finish).
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(new DOMException("timeout", "TimeoutError")), timeoutMs);
  const deadline = controller.signal;
  try {
    return await callWithRetries(cfg, env, payload, fetchImpl, deadline, started);
  } finally {
    clearTimeout(timer);
  }
}

async function callWithRetries(cfg, env, payload, fetchImpl, deadline, started) {
  let format = cfg.format;
  let attempt = 0;
  const tried = new Set();
  for (;;) {
    attempt++;
    try {
      const result = await callOnce(cfg, format, env, payload, fetchImpl, deadline);
      console.log(JSON.stringify({
        event: "model_ok", format, model: result.model || cfg.model, ms: Date.now() - started,
        in: result.usage?.input_tokens ?? result.usage?.prompt_tokens,
        out: result.usage?.output_tokens ?? result.usage?.completion_tokens,
        truncated: !!result.truncated, requestId: result.requestId || undefined,
      }));
      return result;
    } catch (err) {
      if (err?.name === "TimeoutError" || err?.name === "AbortError") {
        throw new ModelError("timeout", `after ${Date.now() - started} ms`);
      }
      if (!(err instanceof ModelError)) throw new ModelError("network_error", String(err?.message || err).slice(0, 200));
      console.warn(JSON.stringify({ event: "model_error", format, code: err.code, status: err.status, detail: err.detail, requestId: err.requestId || undefined }));
      // Endpoint missing (nothing billed): try another API format once.
      if (err.code === "endpoint_not_found" && !env.API_FORMAT && !tried.has("fallback")) {
        tried.add("fallback");
        format = format === "chat" ? "responses" : "chat";
        continue;
      }
      if (RETRYABLE.has(err.code) && attempt === 1 && Date.now() - started < 25_000) {
        await new Promise(r => setTimeout(r, 1500));
        continue;
      }
      throw err;
    }
  }
}

/* ---------------------------- model list ---------------------------- */

const NOT_CHAT = /image|dall|whisper|tts|audio|embed|moderation|rerank|transcri|speech|video|sora/i;

/** Asks Tooken for its models (OpenAI-style GET /models). Returns ids or null. */
export async function listModels(env, { fetchImpl = fetch } = {}) {
  const { baseUrl } = modelConfig(env);
  try {
    const res = await fetchImpl(`${baseUrl}/models`, {
      headers: { authorization: `Bearer ${env.TOOKEN_API_KEY}`, "x-api-key": env.TOOKEN_API_KEY || "" },
      signal: AbortSignal.timeout ? AbortSignal.timeout(8000) : undefined,
    });
    if (!res.ok) return null;
    const data = await res.json();
    const items = Array.isArray(data?.data) ? data.data : Array.isArray(data?.models) ? data.models : Array.isArray(data) ? data : [];
    const ids = items.map(m => (typeof m === "string" ? m : m?.id || m?.name)).filter(id => typeof id === "string" && id && !NOT_CHAT.test(id));
    return ids.length ? [...new Set(ids)] : null;
  } catch {
    return null;
  }
}
