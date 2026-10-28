/**
 * Mock OpenCode v1 service (the pre-2.0 surface: /session/*, /event, /global/health).
 * Used to prove the bridge still works against OpenCode 1.x.
 *   node tools/mock-opencode-v1.mjs
 */

import http from "node:http";

const PORT = Number(process.env.PORT || 4096);
const PASSWORD = process.env.MOCK_PASSWORD || "testpw";
const SESSION = "ses_v1";
const DELAY = Number(process.env.MOCK_DELAY || 40);

const ANSWER = [
  "Sure - here is the summary.\n\n",
  "```tsv\nRegion\tTotal\nNorth\t120\nSouth\t95\n```\n",
];

let assistantText = "";
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
  return (req.headers.authorization || "") === "Basic " + Buffer.from(`opencode:${PASSWORD}`).toString("base64");
}

function send(res, status, body) {
  res.writeHead(status, { "content-type": "application/json" });
  res.end(JSON.stringify(body));
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

  if (url.pathname === "/global/health") return send(res, 200, { healthy: true, version: "1.1.0-mock" });
  if (url.pathname === "/api/info") return send(res, 404, { error: "not found" }); // v1 has no /api/*

  if (url.pathname === "/session" && req.method === "POST") {
    await readBody(req);
    return send(res, 200, { id: SESSION, title: "v1 session" }); // v1 returns the object directly
  }

  if (url.pathname === `/session/${SESSION}/message` && req.method === "POST") {
    const body = JSON.parse((await readBody(req)) || "{}");
    const text = (body.parts || []).map((p) => p.text || "").join("");
    if (!text) return send(res, 400, { error: "parts required" });
    assistantText = "";
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
      broadcast({ type: "session.idle", properties: { sessionID: SESSION, status: "idle" } });
    })();
    return send(res, 200, { info: { id: "m1", role: "assistant", sessionID: SESSION }, parts: [{ type: "text", text: "" }] });
  }

  if (url.pathname === `/session/${SESSION}/message`) {
    return send(res, 200, [
      { info: { id: "m0", role: "user", sessionID: SESSION }, parts: [{ type: "text", text: "hi" }] },
      { info: { id: "m1", role: "assistant", sessionID: SESSION }, parts: [{ type: "text", text: assistantText }] },
    ]);
  }

  if (url.pathname === `/session/${SESSION}/abort`) return send(res, 200, true);

  if (url.pathname === "/event") {
    res.writeHead(200, { "content-type": "text/event-stream", "cache-control": "no-cache", connection: "keep-alive" });
    subscribers.add(res);
    res.write(`data: ${JSON.stringify({ type: "server.connected", properties: {} })}\n\n`);
    req.on("close", () => subscribers.delete(res));
    return;
  }

  return send(res, 404, { error: `v1 mock: no route ${url.pathname}` });
});

server.listen(PORT, "127.0.0.1", () => {
  console.log(`mock opencode v1 service on http://127.0.0.1:${PORT} (password: ${PASSWORD})`);
});
