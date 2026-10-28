/**
 * Bridge test suite: service discovery, auth recovery, v1 compatibility,
 * resilience (dead SSE stream, rejected agent override), and HTTP hardening.
 *
 *   node tools/test-bridge.mjs
 */
import fs from "node:fs";
import path from "node:path";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";

process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const BRIDGE_PORT = process.env.TEST_BRIDGE_PORT || "3100";
const MOCK_PORT = process.env.TEST_MOCK_PORT || "28895";
const DEFAULT_PORT = process.env.TEST_DEFAULT_PORT || "49375";
const MOCK_URL = `http://127.0.0.1:${MOCK_PORT}`;
const BRIDGE = `https://localhost:${BRIDGE_PORT}`;
const TMP = "/tmp/bridgetest";

let failures = 0;
let checks = 0;
const ok = (name, condition, detail = "") => {
  checks++;
  console.log(`  ${condition ? "PASS" : "FAIL"}  ${name}${condition ? "" : "  -> " + detail}`);
  if (!condition) failures++;
};
const check = (name, actual, expected) => {
  checks++;
  const pass = JSON.stringify(actual) === JSON.stringify(expected);
  console.log(`  ${pass ? "PASS" : "FAIL"}  ${name}${pass ? "" : `  -> expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`}`);
  if (!pass) failures++;
};
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function start(file, env) {
  const proc = spawn(process.execPath, [path.join(ROOT, "tools", file)], {
    env: { ...process.env, ...env },
    stdio: ["ignore", "pipe", "pipe"],
  });
  proc.stderr.on("data", (d) => process.stderr.write(`[${file}] ${d}`));
  return proc;
}

function startBridge(env, home) {
  const proc = spawn(process.execPath, [path.join(ROOT, "bridge", "server.mjs")], {
    env: {
      ...process.env,
      HOME: home,
      XDG_STATE_HOME: path.join(home, ".local", "state"),
      XDG_CONFIG_HOME: path.join(home, ".config"),
      OPENCODE_SERVICE_DEFAULT_PORTS: DEFAULT_PORT,
      PORT: BRIDGE_PORT,
      ...env,
    },
    stdio: ["ignore", "pipe", "pipe"],
  });
  let log = "";
  proc.stdout.on("data", (d) => (log += d.toString()));
  proc.stderr.on("data", (d) => (log += d.toString()));
  return { proc, getLog: () => log };
}

async function waitForStatus(predicate, timeoutMs = 20000) {
  const deadline = Date.now() + timeoutMs;
  let last = null;
  while (Date.now() < deadline) {
    try {
      const res = await fetch(`${BRIDGE}/api/status`);
      last = await res.json();
      if (predicate(last)) return last;
    } catch {
      /* not up yet */
    }
    await sleep(300);
  }
  return last;
}

/** Open an SSE stream and collect events until `done` or timeout. */
async function collectEvents(sessionID, timeoutMs = 30000) {
  const res = await fetch(`${BRIDGE}/api/events?sessionID=${encodeURIComponent(sessionID)}`, {
    headers: { accept: "text/event-stream" },
  });
  const events = [];
  const reader = res.body.getReader();
  let buffer = "";
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    const { value, done } = await Promise.race([
      reader.read(),
      sleep(Math.max(0, deadline - Date.now())).then(() => ({ value: undefined, done: true })),
    ]);
    if (done || !value) break;
    buffer += Buffer.from(value).toString("utf8");
    let i;
    while ((i = buffer.indexOf("\n\n")) !== -1) {
      const frame = buffer.slice(0, i);
      buffer = buffer.slice(i + 2);
      const line = frame.split("\n").find((l) => l.startsWith("data:"));
      if (!line) continue;
      try {
        const evt = JSON.parse(line.slice(5).trim());
        events.push(evt);
        if (evt.type === "done") {
          await reader.cancel();
          return events;
        }
      } catch {
        /* ignore */
      }
    }
  }
  try {
    await reader.cancel();
  } catch {}
  return events;
}

async function newSession() {
  const res = await fetch(`${BRIDGE}/api/session`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: "{}",
  });
  return res.json();
}

async function prompt(sessionID, text, agent, model) {
  const res = await fetch(`${BRIDGE}/api/prompt`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ sessionID, text, agent, model }),
  });
  return res.json();
}

/* ---------------------------------------------------------------- scenarios */

async function scenario(name, fn) {
  console.log(`\n== ${name} ==`);
  const kill = [];
  try {
    await fn(kill);
  } catch (err) {
    checks++;
    failures++;
    console.log(`  FAIL  threw: ${err.message}`);
  } finally {
    for (const p of kill) {
      try {
        p.kill("SIGKILL");
      } catch {}
    }
    await sleep(400);
  }
}

async function main() {
  fs.rmSync(TMP, { recursive: true, force: true });

  /* 1. discovery from the registration file ------------------------------- */
  await scenario("discovery: registration file (no env vars)", async (kill) => {
    const home = path.join(TMP, "reg");
    fs.mkdirSync(path.join(home, ".local", "state", "opencode"), { recursive: true });
    const mock = start("mock-opencode.mjs", { PORT: MOCK_PORT, MOCK_PASSWORD: "fromfile" });
    kill.push(mock);
    await sleep(600);
    fs.writeFileSync(
      path.join(home, ".local", "state", "opencode", "service.json"),
      JSON.stringify({ id: "abc", version: "2.0.0", url: MOCK_URL, pid: 1, password: "fromfile" })
    );
    const bridge = startBridge({}, home);
    kill.push(bridge.proc);
    const status = await waitForStatus((s) => s && s.connected);
    ok("connected", !!status.connected, JSON.stringify(status));
    ok("sourced from registration file", /registration/.test(status.source || ""), status.source);
    ok("password from file was accepted (no 401)", status.version === "2.0.0-mock", JSON.stringify(status));
  });

  /* 2. discovery from the service config file ----------------------------- */
  await scenario("discovery: service config file (port + password only)", async (kill) => {
    const home = path.join(TMP, "cfg");
    fs.mkdirSync(path.join(home, ".config", "opencode"), { recursive: true });
    const mock = start("mock-opencode.mjs", { PORT: MOCK_PORT, MOCK_PASSWORD: "cfgpass" });
    kill.push(mock);
    await sleep(600);
    fs.writeFileSync(
      path.join(home, ".config", "opencode", "service.json"),
      JSON.stringify({ port: Number(MOCK_PORT), hostname: "127.0.0.1", password: "cfgpass" })
    );
    const bridge = startBridge({}, home);
    kill.push(bridge.proc);
    const status = await waitForStatus((s) => s && s.connected);
    ok("connected", !!status.connected, JSON.stringify(status));
    ok("sourced from service config", /service config/.test(status.source || ""), status.source);
  });

  /* 3. discovery from a default port -------------------------------------- */
  await scenario("discovery: default port fallback", async (kill) => {
    const home = path.join(TMP, "default");
    fs.mkdirSync(home, { recursive: true });
    const mock = start("mock-opencode.mjs", { PORT: DEFAULT_PORT, MOCK_PASSWORD: "defpass" });
    kill.push(mock);
    await sleep(600);
    const bridge = startBridge({ OPENCODE_SERVICE_PASSWORD: "defpass" }, home);
    kill.push(bridge.proc);
    const status = await waitForStatus((s) => s && s.connected);
    ok("connected", !!status.connected, JSON.stringify(status));
    ok("sourced from default port", /default port/.test(status.source || ""), status.source);
  });

  /* 4. 401 -> CLI password recovery --------------------------------------- */
  await scenario("auth: recovers password via `opencode service get password`", async (kill) => {
    const home = path.join(TMP, "cli");
    fs.mkdirSync(path.join(home, "bin"), { recursive: true });
    const fakeCli = path.join(home, "bin", "opencode");
    fs.writeFileSync(fakeCli, `#!/bin/sh\nif [ "$1" = "service" ] && [ "$2" = "get" ] && [ "$3" = "password" ]; then echo "clipass"; fi\n`);
    fs.chmodSync(fakeCli, 0o755);
    // Windows shells resolve .cmd shims, not extensionless shell scripts.
    fs.writeFileSync(
      path.join(home, "bin", "opencode.cmd"),
      '@echo off\r\nif "%1"=="service" if "%2"=="get" if "%3"=="password" echo clipass\r\n'
    );

    const mock = start("mock-opencode.mjs", { PORT: MOCK_PORT, MOCK_PASSWORD: "clipass" });
    kill.push(mock);
    await sleep(600);
    const bridge = startBridge(
      { PATH: `${path.join(home, "bin")}${path.delimiter}${process.env.PATH}`, OPENCODE_SERVICE_URL: MOCK_URL },
      home
    );
    kill.push(bridge.proc);
    const status = await waitForStatus((s) => s && s.connected);
    ok("connected after CLI password recovery", !!status.connected, JSON.stringify(status));
    ok("source mentions the CLI", /service get password/.test(status.source || ""), status.source);
  });

  /* 5. no service at all ---------------------------------------------------- */
  await scenario("no service: /api/status reports 503 with a hint", async (kill) => {
    const home = path.join(TMP, "none");
    const fakeBin = path.join(home, "bin");
    fs.mkdirSync(fakeBin, { recursive: true });
    // Shadow the real `opencode` CLI with a silent shim so it cannot point the
    // bridge at the live desktop service during this scenario.
    fs.writeFileSync(path.join(fakeBin, "opencode"), "#!/bin/sh\nexit 0\n");
    fs.writeFileSync(path.join(fakeBin, "opencode.cmd"), "@echo off\r\nexit /b 0\r\n");
    const bridge = startBridge(
      {
        OPENCODE_SERVICE_URL: "http://127.0.0.1:1",
        PATH: `${fakeBin}${path.delimiter}${process.env.PATH}`,
      },
      home
    );
    kill.push(bridge.proc);
    const status = await waitForStatus((s) => s && s.connected === false && !!s.hint);
    ok("reports not connected", status && status.connected === false, JSON.stringify(status));
    ok("includes an actionable hint", /opencode/i.test(status.hint || ""), status.hint);
    const res = await fetch(`${BRIDGE}/api/session`, { method: "POST", headers: { "content-type": "application/json" }, body: "{}" });
    ok("session creation returns 503 when unconnected", res.status === 503, `status ${res.status}`);
  });

  /* 6. OpenCode v1 compatibility ------------------------------------------- */
  await scenario("OpenCode v1: session, prompt, streaming, abort", async (kill) => {
    const home = path.join(TMP, "v1");
    fs.mkdirSync(home, { recursive: true });
    const mock = start("mock-opencode-v1.mjs", { PORT: "4096", MOCK_PASSWORD: "v1pass", MOCK_DELAY: "120" });
    kill.push(mock);
    await sleep(600);
    const bridge = startBridge({ OPENCODE_SERVICE_URL: "http://127.0.0.1:4096", OPENCODE_SERVICE_PASSWORD: "v1pass" }, home);
    kill.push(bridge.proc);
    const status = await waitForStatus((s) => s && s.connected);
    ok("discovered v1 service via /global/health", !!status.connected && status.version === "1.1.0-mock", JSON.stringify(status));

    const session = await newSession();
    ok("v1 session created", !!session.sessionID, JSON.stringify(session));
    const result = await prompt(session.sessionID, "summarise");
    ok("v1 prompt admitted", result.ok === true, JSON.stringify(result));

    const events = await collectEvents(session.sessionID);
    const types = events.map((e) => e.type);
    ok("v1 turn produced deltas", types.filter((t) => t === "delta").length > 0, JSON.stringify(types));
    ok("v1 turn completed", types.includes("done"), JSON.stringify(types));
    const finalText = (events.find((e) => e.type === "done") || {}).text || "";
    ok("v1 answer contains the tsv block", finalText.includes("```tsv") && finalText.includes("North"), finalText.slice(0, 60));

    const abortRes = await fetch(`${BRIDGE}/api/abort`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ sessionID: session.sessionID }),
    });
    ok("v1 abort returns 200", abortRes.status === 200, `status ${abortRes.status}`);
  });

  /* 7. dead SSE stream -> polling fallback ---------------------------------- */
  await scenario("resilience: turn completes even when /api/event keeps dying", async (kill) => {
    const home = path.join(TMP, "dropevents");
    fs.mkdirSync(home, { recursive: true });
    const mock = start("mock-opencode.mjs", { PORT: MOCK_PORT, MOCK_PASSWORD: "secret", MOCK_DROP_EVENTS: "1", MOCK_DELAY: "20" });
    kill.push(mock);
    await sleep(600);
    const bridge = startBridge({ OPENCODE_SERVICE_URL: MOCK_URL, OPENCODE_SERVICE_PASSWORD: "secret" }, home);
    kill.push(bridge.proc);
    await waitForStatus((s) => s && s.connected);

    const session = await newSession();
    await prompt(session.sessionID, "hello");
    const events = await collectEvents(session.sessionID, 40000);
    const types = events.map((e) => e.type);
    ok("turn completes despite the stream dying", types.includes("done"), JSON.stringify(types));
    const done = events.find((e) => e.type === "done");
    ok("full answer recovered by polling", !!done && done.text.includes("Initech"), JSON.stringify(done && done.text.slice(0, 80)));
  });

  /* 8. rejected agent override --------------------------------------------- */
  await scenario("prompt: retries without the agent when the server rejects it", async (kill) => {
    const home = path.join(TMP, "agent");
    fs.mkdirSync(home, { recursive: true });
    const mock = start("mock-opencode.mjs", { PORT: MOCK_PORT, MOCK_PASSWORD: "secret", MOCK_REJECT_AGENTS: "1", MOCK_DELAY: "20" });
    kill.push(mock);
    await sleep(600);
    const bridge = startBridge({ OPENCODE_SERVICE_URL: MOCK_URL, OPENCODE_SERVICE_PASSWORD: "secret" }, home);
    kill.push(bridge.proc);
    await waitForStatus((s) => s && s.connected);

    const session = await newSession();
    const result = await prompt(session.sessionID, "hello", "excel");
    ok("prompt still admitted", result.ok === true, JSON.stringify(result));
    ok("reports the retry without the agent", /without agent/.test(result.note || ""), result.note);

    const events = await collectEvents(session.sessionID, 20000);
    ok("turn completes after the retry", events.map((e) => e.type).includes("done"), JSON.stringify(events.map((e) => e.type)));
  });

  /* 9. HTTP hardening -------------------------------------------------------- */
  await scenario("http hardening: 404s, bad bodies, path traversal", async (kill) => {
    const home = path.join(TMP, "hard");
    fs.mkdirSync(home, { recursive: true });
    const mock = start("mock-opencode.mjs", { PORT: MOCK_PORT, MOCK_PASSWORD: "secret" });
    kill.push(mock);
    await sleep(600);
    const bridge = startBridge({ OPENCODE_SERVICE_URL: MOCK_URL, OPENCODE_SERVICE_PASSWORD: "secret" }, home);
    kill.push(bridge.proc);
    await waitForStatus((s) => s && s.connected);

    const unknown = await fetch(`${BRIDGE}/api/nope`);
    ok("unknown endpoint -> 404", unknown.status === 404, `status ${unknown.status}`);

    const missing = await fetch(`${BRIDGE}/api/prompt`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ sessionID: "x" }) });
    ok("prompt without text -> 400", missing.status === 400, `status ${missing.status}`);

    const badJson = await fetch(`${BRIDGE}/api/prompt`, { method: "POST", headers: { "content-type": "application/json" }, body: "{not json" });
    ok("malformed JSON body -> 400 (no crash)", badJson.status === 400, `status ${badJson.status}`);

    const upload = await fetch(`${BRIDGE}/api/attachment`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ name: "../budget.csv", type: "text/csv", data: Buffer.from("Month,Amount\nJan,10").toString("base64") }),
    });
    const uploaded = await upload.json();
    ok("attachment upload is accepted", upload.status === 201 && uploaded.id && /^file:/.test(uploaded.uri || ""), JSON.stringify(uploaded));
    ok("attachment filename is sanitized", uploaded.name === "budget.csv", uploaded.name);
    const attachmentSession = await newSession();
    const attachmentPrompt = await fetch(`${BRIDGE}/api/prompt`, {
      method: "POST", headers: { "content-type": "application/json" },
      body: JSON.stringify({ sessionID: attachmentSession.sessionID, text: "read this file", files: [{ id: uploaded.id, uri: uploaded.uri, name: uploaded.name }] }),
    });
    const attachmentPromptBody = await attachmentPrompt.json();
    const auth = "Basic " + Buffer.from("opencode:secret").toString("base64");
    const attachmentState = await (await fetch(`${MOCK_URL}/api/test/state`, { headers: { authorization: auth } })).json();
    ok("attachment URI reaches the OpenCode prompt API", attachmentPromptBody.ok === true && attachmentState.lastPromptFiles?.[0]?.uri === uploaded.uri, JSON.stringify(attachmentState.lastPromptFiles));
    const badUpload = await fetch(`${BRIDGE}/api/attachment`, {
      method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ name: "bad.bin", data: "not-base64" })
    });
    ok("invalid attachment data is rejected", badUpload.status === 400, `status ${badUpload.status}`);
    const removeUpload = await fetch(`${BRIDGE}/api/attachment?id=${encodeURIComponent(uploaded.id)}`, { method: "DELETE" });
    ok("attachment can be removed", removeUpload.status === 200, `status ${removeUpload.status}`);

    const traversal = await fetch(`${BRIDGE}/assets/..%2f..%2fpackage.json`);
    ok("path traversal blocked", traversal.status === 404 || traversal.status === 403, `status ${traversal.status}`);
    const outside = await fetch(`${BRIDGE}/assets/../../etc/passwd`);
    ok("no file outside the project is served", outside.status !== 200, `status ${outside.status}`);

    const pane = await fetch(`${BRIDGE}/taskpane.html`);
    ok("pane served as html", pane.status === 200 && /text\/html/.test(pane.headers.get("content-type") || ""), pane.headers.get("content-type"));
    ok("pane loads the formatter script", (await pane.text()).includes('src="format.js"'));
    const css = await fetch(`${BRIDGE}/taskpane-clean.css`);
    ok("clean UI css served with the right mime type", /text\/css/.test(css.headers.get("content-type") || ""), css.headers.get("content-type"));
    const format = await fetch(`${BRIDGE}/format.js`);
    const formatBody = await format.text();
    ok("formatter script is served", format.status === 200 && /javascript/.test(format.headers.get("content-type") || ""), `status ${format.status}, ${format.headers.get("content-type")}`);
    ok("formatter script defines formatSheet", formatBody.includes("OpenCodeExcelFormat") && formatBody.includes("formatSheet"));
    const icon = await fetch(`${BRIDGE}/assets/icon-80.png`);
    ok("icon served", icon.status === 200 && /image\/png/.test(icon.headers.get("content-type") || ""), icon.headers.get("content-type"));

    // still healthy after all that
    const after = await (await fetch(`${BRIDGE}/api/status`)).json();
    ok("bridge still connected after bad requests", after.connected === true, JSON.stringify(after));
  });

  /* 10. two panes share one session ----------------------------------------- */
  await scenario("two clients receive the same stream", async (kill) => {
    const home = path.join(TMP, "two");
    fs.mkdirSync(home, { recursive: true });
    const mock = start("mock-opencode.mjs", { PORT: MOCK_PORT, MOCK_PASSWORD: "secret", MOCK_DELAY: "30" });
    kill.push(mock);
    await sleep(600);
    const bridge = startBridge({ OPENCODE_SERVICE_URL: MOCK_URL, OPENCODE_SERVICE_PASSWORD: "secret" }, home);
    kill.push(bridge.proc);
    await waitForStatus((s) => s && s.connected);

    const session = await newSession();
    const [a, b] = await Promise.all([collectEvents(session.sessionID, 25000), (async () => { await sleep(150); await prompt(session.sessionID, "hello"); return collectEvents(session.sessionID, 25000); })()]);
    ok("client A saw the turn complete", a.map((e) => e.type).includes("done"), JSON.stringify(a.map((e) => e.type)));
    ok("client B saw the turn complete", b.map((e) => e.type).includes("done"), JSON.stringify(b.map((e) => e.type)));
    const textA = (a.find((e) => e.type === "done") || {}).text || "";
    const textB = (b.find((e) => e.type === "done") || {}).text || "";
    ok("both clients got identical text", textA === textB && textA.length > 0, `${textA.length} vs ${textB.length}`);
  });

  /* 11. model list + model switch ------------------------------------------ */
  await scenario("models: list, reject unknown, switch before prompting", async (kill) => {
    const home = path.join(TMP, "models");
    fs.mkdirSync(home, { recursive: true });
    const mock = start("mock-opencode.mjs", { PORT: MOCK_PORT, MOCK_PASSWORD: "secret", MOCK_DELAY: "20" });
    kill.push(mock);
    await sleep(600);
    const bridge = startBridge({ OPENCODE_SERVICE_URL: MOCK_URL, OPENCODE_SERVICE_PASSWORD: "secret" }, home);
    kill.push(bridge.proc);
    await waitForStatus((s) => s && s.connected);

    const models = await (await fetch(`${BRIDGE}/api/models`)).json();
    ok("models endpoint returns enabled models", models.models.length === 4, JSON.stringify(models));
    ok("models carry id + providerID", models.models.every((m) => m.id && m.providerID && m.name), JSON.stringify(models));
    ok("disabled models are hidden", !models.models.some((m) => m.id === "disabled-model"), JSON.stringify(models));
    ok("free OpenCode models are marked for safe fallback", models.models.filter((m) => m.providerID === "opencode").every((m) => m.isFree === true), JSON.stringify(models));
    ok("paid models are never marked free", models.models.find((m) => m.id === "paid-dashboard")?.isFree === false, JSON.stringify(models));
    check("default model is returned in compact form", models.defaultModel && models.defaultModel.id, "flint-40b-free");

    const session = await newSession();
    const bad = await prompt(session.sessionID, "hello", null, { id: "nope", providerID: "mock" });
    ok("unknown model is rejected", /Could not switch model/.test(bad.error || ""), JSON.stringify(bad));

    const good = await prompt(session.sessionID, "hello", null, { id: "mock-smart", providerID: "mock" });
    ok("known model is accepted", good.ok === true, JSON.stringify(good));

    const auth = "Basic " + Buffer.from("opencode:secret").toString("base64");
    const state = await (await fetch(`${MOCK_URL}/api/test/state`, { headers: { authorization: auth } })).json();
    check("session switched to the requested model", state.model, { id: "mock-smart", providerID: "mock" });

    const events = await collectEvents(session.sessionID, 20000);
    const done = events.find((e) => e.type === "done");
    ok("turn after the switch completes", !!done, JSON.stringify(events.map((e) => e.type)));
    ok("switched turn returns the v2 text content", !!done && done.text.includes("Initech"), done && done.text.slice(0, 80));
  });

  /* 12. provider failure surfaced as an error ------------------------------ */
  await scenario("errors: provider failure is surfaced instead of an empty answer", async (kill) => {
    const home = path.join(TMP, "fail");
    fs.mkdirSync(home, { recursive: true });
    const mock = start("mock-opencode.mjs", { PORT: MOCK_PORT, MOCK_PASSWORD: "secret", MOCK_DELAY: "20", MOCK_FAIL: "1" });
    kill.push(mock);
    await sleep(600);
    const bridge = startBridge({ OPENCODE_SERVICE_URL: MOCK_URL, OPENCODE_SERVICE_PASSWORD: "secret" }, home);
    kill.push(bridge.proc);
    await waitForStatus((s) => s && s.connected);

    const session = await newSession();
    const result = await prompt(session.sessionID, "hello");
    ok("prompt is admitted", result.ok === true, JSON.stringify(result));

    const events = await collectEvents(session.sessionID, 20000);
    const done = events.find((e) => e.type === "done");
    ok("turn completes on a provider failure", !!done, JSON.stringify(events.map((e) => e.type)));
    ok("done carries the provider error", !!done && /503 Upstream request failed/.test(done.error || ""), JSON.stringify(done));
    ok("no text is invented", !!done && done.text === "", JSON.stringify(done));
  });

  console.log(`\n${checks - failures}/${checks} checks passed`);
  process.exit(failures ? 1 : 0);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
