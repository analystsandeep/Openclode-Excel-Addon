/**
 * Mock OpenCode v2 service - used only to test the bridge without OpenCode running.
 *   node tools/mock-opencode.mjs   (then set OPENCODE_SERVICE_URL=http://127.0.0.1:28894)
 *
 * Mirrors the v2.1.x shapes the real service returns:
 *   - messages are newest-first and use `type` + `content[]`
 *   - assistant messages carry `time.created` / `time.completed`
 *   - models are listed by /api/model and selected per session
 */

import http from "node:http";

const PORT = Number(process.env.PORT || 28894);
const PASSWORD = process.env.MOCK_PASSWORD || "testpw";
const SESSION = "ses_mock_1";
const DELAY = Number(process.env.MOCK_DELAY || 40);

const ANSWER = [
  "Here is a formula that flags duplicates in column B:\n\n",
  "```excel\n=IF(COUNTIF($B$2:$B$100,B2)>1,\"duplicate\",\"\")\n```\n\n",
  "And the top offenders as a table:\n\n```tsv\nValue\tCount\n",
  "ACME\t4\nGlobex\t3\nInitech\t2\n```\n",
];

// When a prompt contains this marker the mock answers with a structured edit
// plan instead of the canned text answer. Used by tools/test-pane.mjs.
const PLAN_MARKER = "PLAN_REQUEST";
const FALLBACK_MARKER = "FALLBACK_REQUEST";
const EMPTY_MARKER = "EMPTY_REQUEST";
const PLAN_ANSWER = "```json\n" +
JSON.stringify({
  summary: "Add Budget sheet and rewrite the top-left table",
  operations: [
    { op: "addSheet", name: "Budget", description: "Add a Budget tab" },
    { op: "setValues", sheet: "Sheet1", range: "A1:B2", values: [["Item","Total"],["Widgets",300]], description: "Rewrite header and data" },
    { op: "formatRange", sheet: "Sheet1", range: "A1:B1", format: { bold: true, fill: "#1F3864" }, description: "Style header" }
  ]
}, null, 1) +
"\n```\n";

const MODELS = [
  { id: "flint-40b-free", providerID: "opencode", name: "Flint 40B Free", enabled: true, status: "active", cost: [{ input: 0, output: 0, cache: { read: 0, write: 0 } }] },
  { id: "mimo-v2.6-flash-free", providerID: "opencode", name: "MiMo V2.6 Flash Free", enabled: true, status: "active", cost: [{ input: 0, output: 0, cache: { read: 0, write: 0 } }] },
  { id: "mock-smart", providerID: "mock", name: "Mock Smart", enabled: true, status: "active", cost: [{ input: 0, output: 0, cache: { read: 0, write: 0 } }] },
  { id: "paid-dashboard", providerID: "openai", name: "Paid Dashboard", enabled: true, status: "active", cost: [{ input: 1, output: 2, cache: { read: 0.1, write: 1 } }] },
  { id: "disabled-model", providerID: "mock", name: "Disabled Model", enabled: false, status: "active" },
];

let assistantText = "";
let assistantCreated = 0;
let assistantCompleted = 0;
let assistantError = null;
let selectedModel = { id: "flint-40b-free", providerID: "opencode" };
let lastPromptText = "";
let lastPromptFiles = [];
let phase = 0; // 0 = idle, 1 = streaming
const subscribers = new Set();

function broadcast(event) {
  const payload = `data: ${JSON.stringify(event)}\n\n`;
  for (const res of subscribers) {
    try {
      res.write(payload);
    } catch {
      subscribers.delete(res);
    }
  }
}

function authorized(req) {
  const header = req.headers.authorization || "";
  const expected = "Basic " + Buffer.from(`opencode:${PASSWORD}`).toString("base64");
  return header === expected;
}

function send(res, status, body) {
  const payload = JSON.stringify(body);
  res.writeHead(status, { "content-type": "application/json" });
  res.end(payload);
}

function noContent(res) {
  res.writeHead(204);
  res.end();
}

async function readBody(req) {
  const chunks = [];
  for await (const c of req) chunks.push(c);
  return chunks.length ? Buffer.concat(chunks).toString("utf8") : "";
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://localhost:${PORT}`);
  if (!authorized(req)) {
    res.writeHead(401, { "content-type": "application/json" });
    return res.end(JSON.stringify({ error: "unauthorized" }));
  }

  if (url.pathname === "/api/info") return send(res, 200, { version: "2.0.0-mock", pid: process.pid });
  if (url.pathname === "/api/agent") return send(res, 200, { data: [
    { name: "build", description: "Primary agent for completing workbook requests." },
    { name: "explore", description: "Fast agent specialized for exploring codebases. When calling this agent, specify the desired thoroughness level and include all relevant context." },
  ] });
  if (url.pathname === "/api/skill") return send(res, 200, { data: [
    { id: "financial-analysis", name: "Financial analysis", description: "Analyze financial statements, trends, ratios, forecasts, and supporting workbook ranges in detail." },
  ] });
  if (url.pathname === "/api/model/default") {
    const model = MODELS.find((m) => m.id === "flint-40b-free");
    return send(res, 200, { location: {}, data: model });
  }
  if (url.pathname === "/api/model") return send(res, 200, { location: {}, data: MODELS });

  if (url.pathname === "/api/test/state") {
    return send(res, 200, {
      model: selectedModel,
      lastPromptText,
      lastPromptFiles,
      assistantError,
      phase,
    });
  }

  if (url.pathname === "/api/session" && req.method === "POST") {
    await readBody(req);
    return send(res, 200, { data: { id: SESSION, title: "mock session" } });
  }

  if (url.pathname === `/api/session/${SESSION}/model` && req.method === "POST") {
    const body = JSON.parse((await readBody(req)) || "{}");
    const model = body.model;
    if (!model || !model.id || !model.providerID) return send(res, 400, { error: "model required" });
    const known = MODELS.find(
      (m) => m.enabled && m.id === model.id && m.providerID === model.providerID
    );
    if (!known) return send(res, 400, { error: `unknown model ${model.providerID}/${model.id}` });
    selectedModel = { id: known.id, providerID: known.providerID };
    return noContent(res);
  }

  if (url.pathname === `/api/session/${SESSION}/prompt` && req.method === "POST") {
    const body = JSON.parse((await readBody(req)) || "{}");
    if (typeof body.text !== "string") return send(res, 400, { error: "text required" });
    if (body.agents && process.env.MOCK_REJECT_AGENTS === "1") {
      return send(res, 400, { error: "unknown agent" });
    }

    lastPromptText = body.text;
    lastPromptFiles = body.files || [];

    if (body.text.includes(PLAN_MARKER)) {
      assistantText = PLAN_ANSWER;
      assistantCreated = Date.now();
      assistantCompleted = Date.now();
      assistantError = null;
      phase = 0;
      broadcast({ type: "session.idle", properties: { sessionID: SESSION, status: "idle" } });
      return send(res, 200, { data: { admitted: true } });
    }
    assistantText = "";
    assistantCreated = Date.now();
    assistantCompleted = 0;
    assistantError =
      process.env.MOCK_FAIL === "1" || (body.text.includes(FALLBACK_MARKER) && selectedModel.id === "flint-40b-free")
        ? { type: "provider", message: "503 Upstream request failed: Endpoint is unavailable", status: 503 }
        : null;
    if (body.text.includes(FALLBACK_MARKER) && selectedModel.id === "flint-40b-free") {
      assistantError.message = "Upstream request failed: Model is unavailable.";
    }
    phase = 1;

    if (assistantError) {
      broadcast({ type: "session.error", properties: { sessionID: SESSION, error: assistantError } });
      setTimeout(() => {
        assistantCompleted = Date.now();
        phase = 0;
        broadcast({ type: "session.idle", properties: { sessionID: SESSION, status: "idle" } });
      }, 50);
      return send(res, 200, { data: { admitted: true } });
    }

    if (body.text.includes(EMPTY_MARKER) && selectedModel.id === "flint-40b-free") {
      setTimeout(() => {
        assistantCompleted = Date.now();
        phase = 0;
        broadcast({ type: "session.idle", properties: { sessionID: SESSION, status: "idle" } });
      }, 50);
      return send(res, 200, { data: { admitted: true } });
    }

    // Stream the canned answer with deltas, then go idle.
    (async () => {
      for (const chunk of ANSWER) {
        for (const piece of chunk.match(/[\s\S]{1,18}/g)) {
          await new Promise((r) => setTimeout(r, DELAY));
          assistantText += piece;
          broadcast({
            type: "message.part.updated",
            properties: { sessionID: SESSION, delta: piece, part: { type: "text", sessionID: SESSION, text: assistantText } },
          });
        }
      }
      await new Promise((r) => setTimeout(r, 300));
      assistantCompleted = Date.now();
      phase = 0;
      broadcast({ type: "session.idle", properties: { sessionID: SESSION, status: "idle" } });
    })();
    return send(res, 200, { data: { admitted: true } });
  }

  if (url.pathname === `/api/session/${SESSION}/message`) {
    const messages = [];
    const now = assistantCreated || Date.now();
    messages.push({
      id: "msg_idle",
      type: "idle",
      time: { created: assistantCompleted || now },
    });
    if (assistantCreated) {
      messages.push({
        id: "msg_assistant",
        type: "assistant",
        agent: "build",
        model: selectedModel,
        time: {
          created: assistantCreated,
          ...(assistantCompleted ? { completed: assistantCompleted } : {}),
        },
        content: [
          { type: "reasoning", text: "Thinking about the spreadsheet request." },
          { type: "text", text: assistantText },
        ],
        ...(assistantError ? { error: assistantError } : {}),
      });
    }
    messages.push({
      id: "msg_user",
      type: "user",
      time: { created: now - 5 },
      text: lastPromptText || "hello",
    });
    return send(res, 200, { data: messages, cursor: null });
  }

  if (url.pathname === `/api/session/${SESSION}/interrupt`) {
    phase = 0;
    return send(res, 200, { data: true });
  }

  if (url.pathname === "/api/event") {
    res.writeHead(200, {
      "content-type": "text/event-stream",
      "cache-control": "no-cache",
      connection: "keep-alive",
    });
    subscribers.add(res);
    res.write(`data: ${JSON.stringify({ type: "server.connected", properties: {} })}\n\n`);
    req.on("close", () => subscribers.delete(res));
    if (process.env.MOCK_DROP_EVENTS === "1") {
      // Simulate OpenCode v2 closing /api/event mid-turn: no deltas ever arrive
      // over SSE, so the bridge must finish the turn via message polling.
      setTimeout(() => {
        subscribers.delete(res);
        try { res.end(); } catch {}
      }, 100);
    }
    return;
  }

  return send(res, 404, { error: `mock: no route ${url.pathname}` });
});

server.listen(PORT, "127.0.0.1", () => {
  console.log(`mock opencode v2 service on http://127.0.0.1:${PORT} (password: ${PASSWORD})`);
});
