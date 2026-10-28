/**
 * OpenCode for Excel - local bridge.
 *
 * Serves the task pane over HTTPS (required by Office) on https://localhost:3000
 * and proxies /api/* to the OpenCode background service on 127.0.0.1.
 *
 * Why a bridge: the task pane is a web page, so talking to OpenCode directly would
 * mean CORS + Basic-auth preflights + a hardcoded port. The bridge is same-origin
 * with the pane, discovers the service for you, holds the password, and survives
 * the OpenCode SSE stream being closed mid-turn.
 */

import fs from "node:fs";
import http from "node:http";
import https from "node:https";
import path from "node:path";
import crypto from "node:crypto";
import { fileURLToPath, pathToFileURL } from "node:url";
import { getHttpsServerOptions, ensureCertificatesAreInstalled } from "office-addin-dev-certs";
import { discoverService, authHeaders, ensureWorkspace, workspaceDir } from "./opencode-service.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");
const SRC = path.join(ROOT, "src");
const ASSETS = path.join(ROOT, "assets");
const PORT = Number(process.env.PORT || 3000);

let service = { ok: false, url: null, password: null, info: null, source: null, hint: null };
let workspace = process.cwd();
let attachmentDir = null;
const attachmentUris = new Map();
const MAX_ATTACHMENT_BYTES = 25 * 1024 * 1024;
const MAX_ATTACHMENTS_PER_PROMPT = 20;

/* ------------------------------------------------------------------ upstream */

function ocUrl(pathname) {
  return `${service.url}${pathname}`;
}

async function oc(pathname, { method = "GET", body, headers = {}, timeout = 20000, stream = false } = {}) {
  if (!service.ok) throw new Error("OpenCode service is not connected");
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeout);
  try {
    const res = await fetch(ocUrl(pathname), {
      method,
      headers: {
        ...authHeaders(service.password),
        accept: "application/json",
        ...(body ? { "content-type": "application/json" } : {}),
        ...headers,
      },
      body: body ? JSON.stringify(body) : undefined,
      signal: controller.signal,
    });
    if (stream) return { res, status: res.status };
    const text = await res.text();
    let json = null;
    try {
      json = text ? JSON.parse(text) : null;
    } catch {
      /* non-JSON response */
    }
    return { status: res.status, json, text };
  } finally {
    clearTimeout(timer);
  }
}

/** v2 wraps payloads in { data: ... }; v1 does not. Be tolerant. */
function unwrap(payload) {
  if (!payload) return null;
  if (Array.isArray(payload)) return payload;
  if (payload.data !== undefined) return payload.data;
  return payload;
}

const EVENT_PATHS = ["/api/event", "/event"];
const PROMPT_PATHS = (id) => [`/api/session/${id}/prompt`, `/session/${id}/prompt_async`, `/session/${id}/message`];
const MESSAGE_PATHS = (id) => [`/api/session/${id}/message`, `/api/session/${id}/messages`, `/session/${id}/message`];
const ABORT_PATHS = (id) => [`/api/session/${id}/interrupt`, `/api/session/${id}/abort`, `/session/${id}/abort`];

function messageInfo(message) {
  return message?.info || message || null;
}

function isAssistantMessage(info) {
  return !!info && (info.type === "assistant" || info.role === "assistant");
}

function messageTime(info) {
  return Number(info?.time?.completed || info?.time?.created || info?.time?.streamed || 0);
}

/**
 * v2 assistant messages carry `content: [{ type: "text" | "reasoning" | "tool" }]`.
 * v1 messages carry `parts: [...]`. Only text parts are user-visible: reasoning
 * must never leak into the answer bubble.
 */
function textFromMessage(message) {
  if (!message) return "";
  const info = message.info || message;
  const items = Array.isArray(message.content)
    ? message.content
    : Array.isArray(message.parts)
      ? message.parts
      : Array.isArray(info.content)
        ? info.content
        : Array.isArray(info.parts)
          ? info.parts
          : [];
  return items
    .filter((part) => part && part.type === "text" && typeof part.text === "string")
    .map((part) => part.text)
    .join("\n")
    .trim();
}

/**
 * Newest assistant message produced for the current turn.
 *
 * v2 returns newest-first messages with `time.created`; v1 returns oldest-first
 * messages with no timestamps. Timestamped messages from before this turn are
 * ignored, so the previous answer cannot be mistaken for the new one.
 * Returns null when no message endpoint could be read.
 */
async function latestAssistantState(sessionID, turn) {
  for (const pathname of MESSAGE_PATHS(sessionID)) {
    try {
      const { status, json } = await oc(pathname, { timeout: 10000 });
      if (status !== 200) continue;
      const messages = unwrap(json);
      if (!Array.isArray(messages)) continue;

      const cutoff = (turn?.startedAt || 0) - 1500;
      let chosen = null;
      for (const message of messages) {
        const info = messageInfo(message);
        if (!isAssistantMessage(info)) continue;
        const created = Number(info.time?.created || 0);
        if (turn?.startedAt && created && created < cutoff) continue;
        if (!chosen || messageTime(info) >= messageTime(chosen.info)) {
          chosen = { message, info };
        }
      }
      if (!chosen) return { text: "", error: null, retry: null, completed: false };

      return {
        text: textFromMessage(chosen.message),
        error: chosen.info.error || null,
        retry: chosen.info.retry || null,
        completed: !!chosen.info.time?.completed,
        finish: chosen.info.finish || null,
      };
    } catch {
      /* try next path */
    }
  }
  return null; // unknown
}

/** Switches the model used by subsequent turns on a v2 session. */
async function switchSessionModel(sessionID, model) {
  let lastError = null;
  for (const pathname of [`/api/session/${sessionID}/model`, `/session/${sessionID}/model`]) {
    try {
      const { status, text } = await oc(pathname, {
        method: "POST",
        body: { model: { id: model.id, providerID: model.providerID } },
        timeout: 10000,
      });
      if (status >= 200 && status < 300) return { ok: true };
      lastError = `${pathname} -> HTTP ${status}: ${(text || "").slice(0, 200)}`;
      if (status === 404 || status === 405) continue;
      return { ok: false, error: lastError };
    } catch (err) {
      lastError = err.message;
    }
  }
  return { ok: false, error: lastError || "model switching is not supported" };
}

/* -------------------------------------------------------------- turn tracking */

const turns = new Map(); // sessionID -> state
const TURN_TIMEOUT_MS = 10 * 60 * 1000;
let eventPathIndex = 0; // set from discovery: v2 uses /api/event, v1 uses /event

function getTurn(sessionID) {
  if (!turns.has(sessionID)) {
    turns.set(sessionID, {
      active: false,
      text: "",
      lastChange: 0,
      idleSeen: false,
      startedAt: 0,
      cancelled: false,
      lastError: null,
      lastRetryKey: null,
      listeners: new Set(),
    });
  }
  return turns.get(sessionID);
}

function emit(turn, event) {
  for (const send of turn.listeners) {
    try {
      send(event);
    } catch {
      /* listener gone */
    }
  }
}

/** Surface provider retries (e.g. a transient 503) without ending the turn. */
function emitRetry(turn, retry) {
  if (!retry) return;
  const attempt = retry.attempt || 0;
  const key = `${attempt}:${retry.at || 0}`;
  if (turn.lastRetryKey === key) return;
  turn.lastRetryKey = key;
  const message = retry.error?.message || retry.message || "Retrying";
  turn.lastError = message;
  emit(turn, { type: "retry", attempt, message });
}

/* ------------------------------------------------------------------- routing */

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
};

function sendJson(res, status, payload) {
  const body = JSON.stringify(payload);
  res.writeHead(status, { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" });
  res.end(body);
}

async function readBody(req, maxBytes = 5 * 1024 * 1024) {
  const chunks = [];
  let total = 0;
  for await (const chunk of req) {
    total += chunk.length;
    if (total > maxBytes) {
      const err = new Error(`Request body exceeds the ${Math.round(maxBytes / 1024 / 1024)} MB limit`);
      err.status = 413;
      throw err;
    }
    chunks.push(chunk);
  }
  if (!chunks.length) return {};
  try {
    return JSON.parse(Buffer.concat(chunks).toString("utf8"));
  } catch {
    const err = new Error("Request body is not valid JSON");
    err.status = 400;
    throw err;
  }
}

function safeAttachmentName(value) {
  const cleaned = path.basename(String(value || "attachment"))
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "_")
    .replace(/^\.+/, "")
    .slice(0, 120);
  return cleaned || "attachment";
}

async function handleApi(req, res, url) {
  const pathname = url.pathname;
  const sessionID = url.searchParams.get("sessionID");

  // ---- status -------------------------------------------------------------
  if (pathname === "/api/status" && req.method === "GET") {
    return sendJson(res, 200, {
      connected: service.ok,
      url: service.url,
      version: service.info?.version || null,
      source: service.source,
      hint: service.ok ? null : service.hint,
      workspace,
      port: PORT,
    });
  }

  // ---- audit log (local only, does not need OpenCode) ---------------------
  if (pathname === "/api/audit" && req.method === "GET") {
    var limit = Number(url.searchParams.get("limit") || 200);
    return sendJson(res, 200, { entries: readAuditLog(limit) });
  }
  if (pathname === "/api/audit" && req.method === "POST") {
    var body = await readBody(req);
    appendAuditLog(body);
    return sendJson(res, 200, { ok: true });
  }

  // ---- attachments (stored only in the locked scratch workspace) ----------
  if (pathname === "/api/attachment" && req.method === "POST") {
    const body = await readBody(req, 36 * 1024 * 1024); // 25 MB + base64 overhead
    if (!body.name || typeof body.data !== "string") {
      return sendJson(res, 400, { error: "name and base64 data are required" });
    }
    if (body.data.length % 4 !== 0 || !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(body.data)) {
      return sendJson(res, 400, { error: "Attachment data is not valid base64" });
    }
    if (body.data.length > Math.ceil(MAX_ATTACHMENT_BYTES * 4 / 3) + 8) {
      return sendJson(res, 413, { error: "Attachment exceeds 25 MB" });
    }
    const bytes = Buffer.from(body.data, "base64");
    if (bytes.length > MAX_ATTACHMENT_BYTES) return sendJson(res, 413, { error: "Attachment exceeds 25 MB" });
    const id = crypto.randomUUID();
    const name = safeAttachmentName(body.name);
    const file = path.join(attachmentDir, `${id}-${name}`);
    fs.writeFileSync(file, bytes, { flag: "wx" });
    const uri = pathToFileURL(file).href;
    attachmentUris.set(id, { uri, file, name, type: String(body.type || "application/octet-stream") });
    return sendJson(res, 201, { id, uri, name, type: String(body.type || "application/octet-stream"), size: bytes.length });
  }
  if (pathname === "/api/attachment" && req.method === "DELETE") {
    const id = url.searchParams.get("id");
    const item = id && attachmentUris.get(id);
    if (!item) return sendJson(res, 404, { error: "Attachment not found" });
    try { fs.unlinkSync(item.file); } catch { /* already removed */ }
    attachmentUris.delete(id);
    return sendJson(res, 200, { ok: true });
  }

  if (!service.ok) {
    service = await discoverService();
    if (!service.ok) return sendJson(res, 503, { error: service.hint, connected: false });
  }

  // ---- agents -------------------------------------------------------------
  if (pathname === "/api/agents" && req.method === "GET") {
    for (const p of ["/api/agent", "/agent"]) {
      try {
        const { status, json } = await oc(p, { timeout: 8000 });
        if (status === 200) {
          const list = unwrap(json) || [];
          return sendJson(res, 200, {
            agents: list.map((a) => ({ id: a.name || a.id, description: a.description || "" })),
          });
        }
      } catch {
        /* try next */
      }
    }
    return sendJson(res, 200, { agents: [] });
  }

  // ---- skills -------------------------------------------------------------
  if (pathname === "/api/skills" && req.method === "GET") {
    try {
      const { status, json } = await oc("/api/skill", { timeout: 8000 });
      if (status === 200) {
        const list = unwrap(json) || [];
        return sendJson(res, 200, {
          skills: (Array.isArray(list) ? list : []).map((s) => ({
            id: s.id || s.name,
            name: s.name || s.id,
            description: s.description || "",
          })).filter((s) => s.id),
        });
      }
    } catch { /* skills are optional */ }
    return sendJson(res, 200, { skills: [] });
  }

  // ---- models -------------------------------------------------------------
  if (pathname === "/api/models" && req.method === "GET") {
    for (const p of ["/api/model", "/model"]) {
      try {
        const { status, json } = await oc(p, { timeout: 8000 });
        if (status === 200) {
          const list = unwrap(json);
          const modelSummary = (m) => {
            const costs = Array.isArray(m.cost) ? m.cost : (m.cost ? [m.cost] : []);
            const hasPositiveCost = costs.some((entry) => {
              const cache = entry?.cache || {};
              return Number(entry?.input || 0) > 0 || Number(entry?.output || 0) > 0 ||
                Number(cache.read || 0) > 0 || Number(cache.write || 0) > 0;
            });
            return {
              id: m.id || m.modelID,
              providerID: m.providerID,
              name: m.name || m.id || m.modelID,
              status: m.status || null,
              // Automatic fallback requires both OpenCode's explicit free ID
              // and no positive catalog price. Other providers are never
              // selected automatically, even when their current cost is zero.
              isFree: m.providerID === "opencode" && /-free$/i.test(m.id || m.modelID || "") && !hasPositiveCost,
            };
          };
          const models = (Array.isArray(list) ? list : [])
            .filter((m) => m && (m.id || m.modelID) && m.providerID && m.enabled !== false)
            .map(modelSummary);
          let defaultModel = null;
          try {
            const selected = await oc("/api/model/default", { timeout: 5000 });
            if (selected.status === 200) {
              const rawDefault = unwrap(selected.json);
              if (rawDefault && rawDefault.providerID && (rawDefault.id || rawDefault.modelID)) {
                defaultModel = modelSummary(rawDefault);
              }
            }
          } catch { /* optional endpoint */ }
          return sendJson(res, 200, { models, defaultModel });
        }
      } catch {
        /* try next */
      }
    }
    return sendJson(res, 200, { models: [], defaultModel: null });
  }

  // ---- configured MCP servers and app integrations (status only) ----------
  if (pathname === "/api/connections" && req.method === "GET") {
    let mcp = [];
    let integrations = [];
    try {
      const response = await oc("/api/mcp", { timeout: 8000 });
      const data = unwrap(response.json);
      mcp = Array.isArray(data) ? data : data && typeof data === "object"
        ? Object.keys(data).map((name) => ({ name, ...(data[name] || {}) })) : [];
    } catch { /* optional API */ }
    try {
      const response = await oc("/api/integration", { timeout: 8000 });
      const data = unwrap(response.json);
      integrations = Array.isArray(data) ? data : [];
    } catch { /* optional API */ }
    return sendJson(res, 200, { mcp, integrations });
  }

  // ---- new session --------------------------------------------------------
  if (pathname === "/api/session" && req.method === "POST") {
    const body = await readBody(req);
    let lastError = null;
    // v2: POST /api/session  { location: { directory } }
    for (const attempt of [
      { path: "/api/session", body: { location: { directory: workspace }, ...(body.agent ? { agent: body.agent } : {}) } },
      { path: "/api/session", body: {} },
      { path: "/session", body: {} },
    ]) {
      try {
        const { status, json, text } = await oc(attempt.path, { method: "POST", body: attempt.body, timeout: 15000 });
        if (status === 200 || status === 201) {
          const data = unwrap(json);
          const id = data?.id || data?.sessionID || data?.sessionId;
          if (id) return sendJson(res, 200, { sessionID: id, raw: data });
          lastError = `unexpected response: ${text?.slice(0, 200)}`;
        } else {
          lastError = `${attempt.path} -> HTTP ${status}: ${text?.slice(0, 200)}`;
        }
      } catch (err) {
        lastError = err.message;
      }
    }
    return sendJson(res, 502, { error: `Could not create a session: ${lastError}` });
  }

  // ---- prompt -------------------------------------------------------------
  if (pathname === "/api/prompt" && req.method === "POST") {
    const { sessionID: id, text, agent, skill, model, files = [] } = await readBody(req);
    if (!id || !text) return sendJson(res, 400, { error: "sessionID and text are required" });
    if (model && (!model.id || !model.providerID)) {
      return sendJson(res, 400, { error: "model must include id and providerID" });
    }
    if (!Array.isArray(files) || files.length > MAX_ATTACHMENTS_PER_PROMPT) {
      return sendJson(res, 400, { error: `files must contain at most ${MAX_ATTACHMENTS_PER_PROMPT} attachments` });
    }
    const promptFiles = [];
    for (const file of files) {
      const stored = file && attachmentUris.get(file.id);
      if (!stored || stored.uri !== file.uri) {
        return sendJson(res, 400, { error: `Attachment is missing or no longer available: ${file?.name || "unknown"}` });
      }
      promptFiles.push({ uri: stored.uri, name: stored.name, description: file.description || undefined });
    }

    if (model) {
      const switched = await switchSessionModel(id, model);
      if (!switched.ok) {
        return sendJson(res, 502, {
          error: `Could not switch model to ${model.providerID}/${model.id}: ${switched.error}`,
        });
      }
    }

    const v2Body = {
      text,
      ...(agent ? { agents: [{ name: agent }] } : {}),
      ...(skill ? { skills: [{ id: String(skill) }] } : {}),
      ...(promptFiles.length ? { files: promptFiles } : {}),
    };
    const v1Body = { parts: [{ type: "text", text }], ...(agent ? { agent } : {}) };
    const promptStartedAt = Date.now();
    let admitted = false;
    let lastError = null;

    for (const p of PROMPT_PATHS(id)) {
      try {
        const body = p.startsWith("/api/") ? v2Body : v1Body;
        const { status, text: raw } = await oc(p, { method: "POST", body, timeout: 30000 });
        if (status >= 200 && status < 300) {
          admitted = true;
          break;
        }
        lastError = `${p} -> HTTP ${status}: ${(raw || "").slice(0, 200)}`;
      } catch (err) {
        lastError = err.message;
      }
      // If an agent override was rejected, retry without it.
      if (agent && lastError) {
        try {
          const body = p.startsWith("/api/") ? { text } : { parts: [{ type: "text", text }] };
          const { status } = await oc(p, { method: "POST", body, timeout: 30000 });
          if (status >= 200 && status < 300) {
            admitted = true;
            lastError = `${lastError} (retried without agent "${agent}")`;
            break;
          }
        } catch {
          /* keep original error */
        }
      }
    }

    if (!admitted) return sendJson(res, 502, { error: `OpenCode rejected the prompt: ${lastError}` });

    const turn = getTurn(id);
    turn.active = true;
    turn.text = "";
    turn.idleSeen = false;
    turn.cancelled = false;
    turn.lastError = null;
    turn.lastRetryKey = null;
    turn.startedAt = promptStartedAt;
    turn.lastChange = Date.now();
    return sendJson(res, 200, { ok: true, sessionID: id, note: lastError || null });
  }

  // ---- abort --------------------------------------------------------------
  if (pathname === "/api/abort" && req.method === "POST") {
    const { sessionID: id } = await readBody(req);
    const turn = getTurn(id);
    turn.cancelled = true;
    for (const p of ABORT_PATHS(id)) {
      try {
        const { status } = await oc(p, { method: "POST", body: {}, timeout: 8000 });
        if (status >= 200 && status < 300) {
          turn.active = false;
          emit(turn, { type: "done", text: turn.text, cancelled: true });
          return sendJson(res, 200, { ok: true });
        }
      } catch {
        /* try next */
      }
    }
    turn.active = false;
    emit(turn, { type: "done", text: turn.text, cancelled: true });
    return sendJson(res, 200, { ok: true, note: "stopped locally" });
  }

  // ---- event stream (SSE) -------------------------------------------------
  if (pathname === "/api/events" && req.method === "GET") {
    if (!sessionID) return sendJson(res, 400, { error: "sessionID is required" });
    res.writeHead(200, {
      "content-type": "text/event-stream; charset=utf-8",
      "cache-control": "no-cache, no-transform",
      connection: "keep-alive",
      "x-accel-buffering": "no",
    });
    const send = (event) => {
      try {
        res.write(`data: ${JSON.stringify(event)}\n\n`);
      } catch {
        /* socket closed */
      }
    };
    const turn = getTurn(sessionID);
    turn.listeners.add(send);
    send({ type: "ready", text: turn.text, active: turn.active });

    const keepAlive = setInterval(() => {
      try {
        res.write(`: ping\n\n`);
      } catch {
        /* closed */
      }
    }, 15000);

    const upstream = startUpstreamStream(sessionID, turn);
    const reconcile = startReconcile(sessionID, turn);

    const cleanup = () => {
      clearInterval(keepAlive);
      turn.listeners.delete(send);
      upstream.stop();
      reconcile.stop();
    };
    req.on("close", cleanup);
    res.on("close", cleanup);
    return;
  }

  return sendJson(res, 404, { error: `Unknown endpoint ${pathname}` });
}

/* --------------------------------------------------- upstream SSE + fallback */

function sessionOfEvent(evt) {
  const p = evt?.properties || {};
  return p.sessionID || p.sessionId || p.part?.sessionID || p.info?.sessionID || p.message?.sessionID || null;
}

function startUpstreamStream(sessionID, turn) {
  let stopped = false;
  let attempt = 0;

  const run = async () => {
    while (!stopped && !turn.cancelled) {
      try {
        const path = EVENT_PATHS[eventPathIndex];
        const controller = new AbortController();
        const streamTimer = setTimeout(() => controller.abort(), 5 * 60 * 1000);
        const res = await fetch(ocUrl(path), {
          headers: { ...authHeaders(service.password), accept: "text/event-stream" },
          signal: controller.signal,
        });
        clearTimeout(streamTimer);
        if (res.status === 404) eventPathIndex = (eventPathIndex + 1) % EVENT_PATHS.length;
        if (!res.ok || !res.body) throw new Error(`HTTP ${res.status}`);
        attempt = 0;

        let buffer = "";
        const reader = res.body.getReader();
        for (;;) {
          if (stopped || turn.cancelled) break;
          const { done, value } = await reader.read();
          if (done) break;
          buffer += Buffer.from(value).toString("utf8");
          let index;
          while ((index = buffer.indexOf("\n\n")) !== -1) {
            const frame = buffer.slice(0, index);
            buffer = buffer.slice(index + 2);
            const dataLine = frame.split("\n").find((l) => l.startsWith("data:"));
            if (!dataLine) continue;
            const raw = dataLine.slice(5).trim();
            if (!raw || raw === "[DONE]") continue;
            let evt;
            try {
              evt = JSON.parse(raw);
            } catch {
              continue;
            }
            if (sessionOfEvent(evt) !== sessionID) continue;
            const props = evt.properties || {};
            const type = evt.type || "";

            if (type === "message.part.updated") {
              const part = props.part || {};
              if (props.delta) {
                turn.text += props.delta;
                turn.lastChange = Date.now();
                emit(turn, { type: "delta", delta: props.delta });
              } else if (part.type === "text" && typeof part.text === "string") {
                if (part.text.length >= turn.text.length) {
                  turn.text = part.text;
                  turn.lastChange = Date.now();
                  emit(turn, { type: "text", text: part.text });
                }
              } else if (part.type === "tool") {
                emit(turn, { type: "tool", name: part.tool || part.name || "tool", state: part.state?.status || "" });
              }
            } else if (type === "message.updated") {
              const info = props.info || props.message || {};
              if (isAssistantMessage(info)) {
                if (info.error?.message) {
                  turn.lastError = info.error.message;
                  emit(turn, { type: "error", message: turn.lastError });
                }
                if (info.retry) emitRetry(turn, info.retry);
              }
            } else if (type.includes("retry")) {
              emitRetry(turn, props.retry || props);
            } else if (type === "permission.asked") {
              const perm = props.permission || props;
              emit(turn, { type: "permission", tool: perm?.tool || perm?.pattern || "tool" });
            } else if (type === "session.error" || type === "message.error") {
              const message = props.error?.message || props.message || "OpenCode reported an error";
              turn.lastError = message;
              emit(turn, { type: "error", message });
            } else if (type === "session.idle" || type === "session.status") {
              const status = props.status || (type === "session.idle" ? "idle" : "");
              if (status === "idle") {
                turn.idleSeen = true;
                turn.lastChange = turn.lastChange || Date.now();
              }
            }
          }
        }
      } catch {
        /* stream ended or failed - fall through to reconnect */
      }
      if (stopped || turn.cancelled) break;
      attempt += 1;
      const delay = Math.min(500 * attempt, 4000);
      await new Promise((r) => setTimeout(r, delay));
    }
  };

  run();
  return {
    stop() {
      stopped = true;
    },
  };
}

/**
 * Polls messages so text, errors and completion survive a flaky /api/event stream.
 *
 * Completion rules (in order): the assistant message is marked completed, an
 * explicit session.idle event, the answer text holding steady across several
 * polls, an overall timeout, or a turn that produced nothing at all. Never rely
 * on SSE alone - OpenCode v2 closes /api/event mid-turn, so idle can be missed.
 */
function startReconcile(sessionID, turn) {
  let stopped = false;
  let stablePolls = 0;
  const STABLE_POLLS_TO_FINISH = 8; // ~10s of no growth

  const loop = async () => {
    while (!stopped) {
      if (turn.active && !turn.cancelled) {
        const snapshot = await latestAssistantState(sessionID, turn);
        if (snapshot) {
          if (snapshot.text && snapshot.text !== turn.text) {
            turn.text = snapshot.text;
            turn.lastChange = Date.now();
            stablePolls = 0;
            emit(turn, { type: "text", text: snapshot.text });
          } else if (snapshot.text && snapshot.text === turn.text) {
            stablePolls += 1;
          }

          if (snapshot.error?.message) turn.lastError = snapshot.error.message;
          if (snapshot.retry) emitRetry(turn, snapshot.retry);

          const idleLongEnough =
            turn.idleSeen && Date.now() - turn.lastChange > 1500 && (!!turn.text || !!turn.lastError);
          const settled = !!snapshot.text && stablePolls >= STABLE_POLLS_TO_FINISH && Date.now() - turn.startedAt > 3000;
          const timedOut = Date.now() - turn.startedAt > TURN_TIMEOUT_MS;
          const producedNothing = !snapshot.text && !snapshot.retry && Date.now() - turn.startedAt > 60000;
          const retryingTooLong = !!snapshot.retry && Date.now() - turn.startedAt > 120000;
          if (snapshot.completed || idleLongEnough || settled || timedOut || producedNothing || retryingTooLong) {
            turn.active = false;
            const error =
              turn.lastError ||
              (snapshot.completed && !turn.text ? "The model finished without returning any text." : null);
            emit(turn, { type: "done", text: turn.text, timeout: timedOut, error });
          }
        }
      }
      await new Promise((r) => setTimeout(r, 1200));
    }
  };

  loop();
  return {
    stop() {
      stopped = true;
    },
  };
}

/* ------------------------------------------------------------ static server */

function serveStatic(req, res, url) {
  let pathname = decodeURIComponent(url.pathname);
  if (pathname === "/" || pathname === "") pathname = "/taskpane.html";

  const candidates = {
    "/taskpane.html": path.join(SRC, "taskpane.html"),
    "/taskpane.js": path.join(SRC, "taskpane.js"),
    "/taskpane.css": path.join(SRC, "taskpane.css"),
    "/taskpane-clean.css": path.join(SRC, "taskpane-clean.css"),
    "/lib.js": path.join(SRC, "lib.js"),
    "/plan.js": path.join(SRC, "plan.js"),
    "/format.js": path.join(SRC, "format.js"),
  };

  let file = candidates[pathname];
  if (!file && pathname.startsWith("/assets/")) {
    file = path.join(ASSETS, path.basename(pathname));
  }
  if (!file || !file.startsWith(ROOT) || !fs.existsSync(file)) {
    res.writeHead(404, { "content-type": "text/plain" });
    return res.end("Not found");
  }

  const ext = path.extname(file).toLowerCase();
  res.writeHead(200, {
    "content-type": MIME[ext] || "application/octet-stream",
    "cache-control": "no-store",
  });
  fs.createReadStream(file).pipe(res);
}

/* -------------------------------------------------------------- audit log */

var AUDIT_PATH = null;

function readAuditLog(limit) {
  try {
    if (!AUDIT_PATH || !fs.existsSync(AUDIT_PATH)) return [];
    var lines = fs.readFileSync(AUDIT_PATH, "utf8").split("\n").filter(function (l) { return l.length > 0; });
    var out = [];
    var start = Math.max(0, lines.length - limit);
    for (var i = start; i < lines.length; i++) {
      try { out.push(JSON.parse(lines[i])); } catch { /* skip corrupt line */ }
    }
    return out;
  } catch {
    return [];
  }
}

function appendAuditLog(record) {
  try {
    if (!AUDIT_PATH) return;
    fs.appendFileSync(AUDIT_PATH, JSON.stringify(record) + "\n", "utf8");
  } catch (err) {
    console.warn("[audit] " + err.message);
  }
}

/* ------------------------------------------------------------------- startup */

async function main() {
  workspace = ensureWorkspace();
  AUDIT_PATH = path.join(workspace, "audit.log");
  attachmentDir = path.join(workspace, ".attachments");
  fs.mkdirSync(attachmentDir, { recursive: true });

  const requestHandler = async (req, res) => {
    const url = new URL(req.url, `http://${req.headers.host || "localhost"}`);
    if (url.pathname.startsWith("/api/")) {
      try {
        return await handleApi(req, res, url);
      } catch (err) {
        return sendJson(res, err.status || 500, { error: err.message });
      }
    }
    return serveStatic(req, res, url);
  };

  let server;
  let scheme = "https";
  try {
    await ensureCertificatesAreInstalled();
    const options = await getHttpsServerOptions();
    server = https.createServer(options, requestHandler);
  } catch (err) {
    console.warn(`\n[cert] Could not create HTTPS dev certificates (${err.message}).`);
    console.warn("[cert] Falling back to plain HTTP - Excel will refuse to load the pane over http.");
    console.warn("[cert] Fix by running: npx office-addin-dev-certs install\n");
    scheme = "http";
    server = http.createServer(requestHandler);
  }

  server.listen(PORT, "0.0.0.0", async () => {
    console.log(`
  OpenCode for Excel - local bridge
  ---------------------------------
  Task pane : ${scheme}://localhost:${PORT}/taskpane.html
  Workspace : ${workspace}
`);

    service = await discoverService();
    if (service.eventPath) {
      const index = EVENT_PATHS.indexOf(service.eventPath);
      if (index >= 0) eventPathIndex = index;
    }
    if (service.ok) {
      console.log(`  OpenCode  : connected -> ${service.url} (v${service.info?.version || "?"}) via ${service.source}, events on ${EVENT_PATHS[eventPathIndex]}`);
    } else {
      console.log(`  OpenCode  : NOT CONNECTED`);
      console.log(`             ${service.hint}`);
    }
    console.log("");

    // Re-discover in the background so the pane recovers when OpenCode restarts.
    setInterval(async () => {
      if (service.ok) {
        try {
          const { status } = await oc("/api/info", { timeout: 5000 });
          if (status !== 200) service = { ...service, ok: false, hint: "OpenCode service stopped responding." };
        } catch {
          service = { ...service, ok: false, hint: "OpenCode service stopped responding." };
        }
      } else {
        service = await discoverService();
        if (service.eventPath) {
          const index = EVENT_PATHS.indexOf(service.eventPath);
          if (index >= 0) eventPathIndex = index;
        }
      }
    }, 20000);
  });
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
