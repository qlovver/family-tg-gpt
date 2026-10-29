import { test } from "node:test";
import assert from "node:assert/strict";
import { createWorker, historyKey, fileKey, shouldUseAgent } from "../src/bot.js";
import { makeEnv, makeFetch, textUpdate, documentUpdate, webhookRequest, execCtx, sentTexts } from "./helpers.js";
import { MIME } from "../src/build.js";

function withAgent(f, handler) {
  const calls = [];
  const fetchImpl = async (url, init) => {
    if (String(url).startsWith("https://agent.example")) { calls.push({ url: String(url), init }); return handler(String(url), init); }
    return f.fetch(url, init);
  };
  return { fetchImpl, calls };
}
const env = () => makeEnv({ AGENT_URL: "https://agent.example", AGENT_SECRET: "agent-secret" });

test("heavy tasks go to the server agent with context; simple ones stay in the Worker", async () => {
  const e = env(); const f = makeFetch();
  const a = withAgent(f, () => Response.json({ queued: true, position: 0 }, { status: 202 }));
  const w = createWorker({ fetchImpl: a.fetchImpl });
  await w.fetch(webhookRequest(textUpdate(111, "сколько будет 2+2?")), e, execCtx);
  assert.equal(a.calls.length, 0);
  assert.equal(f.calls.model.length, 1);

  await w.fetch(webhookRequest(textUpdate(111, "сделай красивую презентацию про Солоу")), e, execCtx);
  assert.equal(a.calls.length, 1);
  assert.equal(f.calls.model.length, 1, "no model call in the Worker");
  const job = JSON.parse(a.calls[0].init.body);
  assert.equal(a.calls[0].init.headers["x-agent-secret"], "agent-secret");
  assert.equal(job.model, "claude-opus-5-5");
  assert.equal(job.history.length, 2, "recent history is passed");
  assert.match(job.text, /презентацию/);

  await w.fetch(webhookRequest(documentUpdate(111, "d1", "курсовая.docx", MIME.docx, "исправь оформление")), e, execCtx);
  assert.deepEqual(JSON.parse(a.calls[1].init.body).files, [{ file_id: "d1", name: "курсовая.docx" }]);
});

test("agent unreachable -> the Worker does the task itself", async () => {
  const e = env(); const f = makeFetch();
  const a = withAgent(f, () => { throw new Error("connect timeout"); });
  await createWorker({ fetchImpl: a.fetchImpl }).fetch(webhookRequest(textUpdate(111, "сделай таблицу в экселе")), e, execCtx);
  assert.equal(f.calls.model.length, 1);
  assert.ok(sentTexts(f.calls).some(t => /Сервер сейчас недоступен/.test(t)));
});

test("/agent off disables the agent for that user only", async () => {
  const e = env(); const f = makeFetch();
  const a = withAgent(f, url => url.endsWith("/health") ? Response.json({ ok: true, queue: 0, done: 3 }) : Response.json({ queued: true, position: 0 }, { status: 202 }));
  const w = createWorker({ fetchImpl: a.fetchImpl });
  await w.fetch(webhookRequest(textUpdate(111, "/agent off")), e, execCtx);
  assert.match(sentTexts(f.calls).at(-1), /работает ✅[\s\S]*выключен/);
  await w.fetch(webhookRequest(textUpdate(111, "сделай презентацию")), e, execCtx);
  assert.equal(a.calls.filter(c => c.url.endsWith("/job")).length, 0);
  await w.fetch(webhookRequest(textUpdate(222, "сделай презентацию")), e, execCtx);
  assert.equal(a.calls.filter(c => c.url.endsWith("/job")).length, 1);
});

test("/agent/done stores the exchange and the sent file; wrong secret is rejected", async () => {
  const e = env(); const f = makeFetch();
  const w = createWorker({ fetchImpl: f.fetch });
  const post = (secret, body) => w.fetch(new Request("https://x.dev/agent/done", { method: "POST", headers: { "x-agent-secret": secret, "content-type": "application/json" }, body: JSON.stringify(body) }), e, execCtx);
  assert.equal((await post("nope", { uid: "111" })).status, 403);
  assert.equal((await post("agent-secret", { uid: "999", answer: "x" })).status, 403, "only allowed users");
  const res = await post("agent-secret", { uid: "111", user: "сделай презентацию", answer: "Готово", files: [{ file_id: "F1", name: "Доклад.pptx" }] });
  assert.equal(res.status, 200);
  const rec = JSON.parse(e.HISTORY.map.get(historyKey("111")));
  assert.deepEqual(rec.messages.map(m => m.content), ["сделай презентацию", "Готово"]);
  const last = JSON.parse(e.HISTORY.map.get(fileKey("111")));
  assert.equal(last.fileId, "F1");
  assert.equal(last.kind, "pptx");
});

test("shouldUseAgent heuristics", () => {
  assert.ok(shouldUseAgent({ text: "построй график функции y = x²" }));
  assert.ok(shouldUseAgent({ text: "сделай PDF с расписанием" }));
  assert.ok(!shouldUseAgent({ text: "реши задачу" }));
  assert.ok(!shouldUseAgent({ text: "", image: { fileId: "p" } }), "photo homework stays fast");
  assert.ok(shouldUseAgent({ text: "исправь ошибки", file: { kind: "docx" } }));
});
