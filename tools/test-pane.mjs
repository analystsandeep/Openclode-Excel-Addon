/**
 * Integration test for the task pane itself (src/taskpane.js + src/lib.js).
 *
 * Boots the mock OpenCode service AND the real bridge, then runs the real pane code
 * inside jsdom with:
 *   - a fake Excel object model (records every write)
 *   - a fetch shim pointing at the live bridge (relative URLs, self-signed cert)
 *   - a minimal EventSource shim over the bridge SSE stream
 *
 * Run: node tools/test-pane.mjs
 */
import fs from "node:fs";
import path from "node:path";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
import { JSDOM } from "jsdom";

process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0"; // self-signed localhost dev cert

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const BRIDGE_PORT = process.env.TEST_BRIDGE_PORT || "3100";
const MOCK_PORT = process.env.TEST_MOCK_PORT || "28895";
const BASE = `https://localhost:${BRIDGE_PORT}`;
const HOME = "/tmp/panetest-home";
const BRIDGE_TOKEN = "pane-test-token";

const usedRangeReads = { count: 0 };
let failures = 0;
const check = (name, actual, expected) => {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  console.log(`${ok ? "PASS" : "FAIL"}  ${name}`);
  if (!ok) {
    failures++;
    console.log(`      expected: ${JSON.stringify(expected)}`);
    console.log(`      actual:   ${JSON.stringify(actual)}`);
  }
};
const ok = (name, condition, detail = "") => {
  console.log(`${condition ? "PASS" : "FAIL"}  ${name}${condition ? "" : "  " + detail}`);
  if (!condition) failures++;
};
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/* ------------------------------------------------------------ fakes -------- */

function createFakeExcel() {
  const ops = [];
  const selectionHandlers = [];
  const cellStore = {}; // a1 -> { values, formulas } - backs snapshot/undo reads

  const makeHandle = (a1, opts) => {
    if (!cellStore[a1]) cellStore[a1] = { values: undefined, formulas: undefined };
    const store = cellStore[a1];
    if (opts && opts.values != null) store.values = opts.values; // seed without recording an op
    const base = {
      rowCount: opts && opts.rowCount,
      columnCount: opts && opts.columnCount,
      address: opts && opts.address,
      get values() { return store.values; },
      set values(v) {
        ops.push({ kind: "values", a1, values: v });
        store.values = v;
      },
      get formulas() { return store.formulas; },
      set formulas(v) {
        ops.push({ kind: "formulas", a1, values: v });
        store.formulas = v;
      },
      load() {},
      getResizedRange(rows, cols) {
        return {
          get values() { return store.values; },
          set values(v) {
            ops.push({ kind: "values", a1, resized: [rows, cols], values: v });
            store.values = v;
          },
          get formulas() { return store.formulas; },
          set formulas(v) {
            ops.push({ kind: "formulas", a1, resized: [rows, cols], values: v });
            store.formulas = v;
          },
          format: {
            autofitColumns() {
              ops.push({ kind: "autofit", a1 });
            },
          },
          load() {},
        };
      },
      format: {
        font: {},
        fill: {},
        autofitColumns() { ops.push({ kind: "autofit", a1 }); },
        autofitRows() { ops.push({ kind: "autofitRows", a1 }); },
        borders: {
          getItem(edge) {
            return {
              set style(value) { ops.push({ kind: "border", a1, edge, prop: "style", value }); },
              set weight(value) { ops.push({ kind: "border", a1, edge, prop: "weight", value }); },
              set color(value) { ops.push({ kind: "border", a1, edge, prop: "color", value }); },
            };
          },
        },
      },
    };
    return base;
  };

  const sheet = {
    name: "Sheet1",
    getRange(a1) {
      return makeHandle(a1);
    },
    getUsedRange() {
      usedRangeReads.count += 1;
      return makeHandle("Sheet1!A1:D20", { rowCount: 20, columnCount: 4, address: "Sheet1!A1:D20", values: selection.values });
    },
    tables: { items: [], load() {} },
    comments: {
      items: [
        {
          cellAddress: "B5",
          content: "Please double-check this amount",
          authorName: "Ada",
          creationDate: "2026-01-01T00:00:00Z",
          resolved: false,
          load() {},
          replies: {
            items: [{ authorName: "Bob", content: "Checked and corrected", creationDate: "2026-01-02T00:00:00Z", load() {} }],
            load() {},
          },
        },
      ],
      load() {},
      add(cell, text) { ops.push({ kind: "commentAdd", cell, text }); },
    },
    load() {},
    set name(v) { ops.push({ kind: "renameSheet", from: "Sheet1", to: v }); this._name = v; },
    get name() { return this._name || "Sheet1"; },
    delete() { ops.push({ kind: "deleteSheet", name: "Sheet1" }); },
    add: undefined,
  };

  const sheet2 = {
    name: "Sheet2",
    getRange(a1) {
      return makeHandle("Sheet2!" + a1, a1 === "A1:B2" ? { values: [["Date", "Sales"], ["Jan", 1000]] } : undefined);
    },
    getUsedRange() {
      return makeHandle("Sheet2!A1:B2", { rowCount: 2, columnCount: 2, address: "Sheet2!A1:B2", values: [["Date", "Sales"], ["Jan", 1000]] });
    },
    tables: { items: [], load() {} },
    load() {},
    set name(v) { ops.push({ kind: "renameSheet", from: "Sheet2", to: v }); this._name = v; },
    get name() { return this._name || "Sheet2"; },
    delete() { ops.push({ kind: "deleteSheet", name: "Sheet2" }); },
  };

  sheet.add = function (name) {
    ops.push({ kind: "addSheet", name });
    return { name: name, getRange: (a1) => makeHandle(name + "!" + a1), load() {} };
  };

  const selection = {
    address: "Sheet1!A1:D20",
    rowCount: 20,
    columnCount: 4,
    values: [["Name", "Amount"], ["ACME", 10], ["Globex", 20]],
    formulas: [["Name", "Amount"], ["ACME", 10], ["Globex", 20]],
    loaded: [],
    sampleRequests: [],
    load(props) { this.loaded.push(props); ops.push({ kind: "selectionLoad", props }); },
    getAbsoluteResizedRange(rows, cols) {
      this.sampleRequests.push([rows, cols]);
      ops.push({ kind: "sampleSelection", rows, cols });
      return {
        rowCount: rows,
        columnCount: cols,
        get values() { return selection.values.slice(0, rows).map((row) => (row || []).slice(0, cols)); },
        get formulas() { return selection.formulas.slice(0, rows).map((row) => (row || []).slice(0, cols)); },
        load() {},
      };
    },
  };
  // Before-state the formatter/snapshot reads back for the undo test.
  cellStore["A1:B1"] = { values: [["Name", "Amount"]] };
  cellStore["A1:B2"] = { values: [["Name", "Amount"], ["ACME", 10]] };
  cellStore["D5:E5"] = { values: [["Before", "State"]], formulas: [["Before", "State"]] };

  const worksheets = {
    items: [sheet, sheet2],
    getActiveWorksheet: () => sheet,
    getItem: (name) => (name === "Sheet1" ? sheet : sheet2),
    add: sheet.add,
    load() {},
  };

  const ctx = {
    workbook: {
      getSelectedRange: () => selection,
      worksheets,
      onSelectionChanged: { add: (fn) => selectionHandlers.push(fn) },
    },
    sync: async () => undefined,
  };

  const Excel = {
    run: async (fn) => fn(ctx),
    InsertShift: { down: "down", right: "right" },
    DeleteShiftDirection: { up: "up", left: "left" },
  };

  return { Excel, ops, selectionHandlers, selection, sheet, sheet2, cellStore };
}

/** Minimal EventSource over an HTTP stream (jsdom has none). */
function makeEventSource(base) {
  return class EventSource {
    constructor(url) {
      this.url = url.startsWith("http") ? url : base + url;
      this.onmessage = null;
      this.onerror = null;
      this._closed = false;
      this._start();
    }
    async _start() {
      try {
        const res = await fetch(this.url, { headers: { accept: "text/event-stream" } });
        let buffer = "";
        for await (const chunk of res.body) {
          if (this._closed) break;
          buffer += Buffer.from(chunk).toString("utf8");
          let i;
          while ((i = buffer.indexOf("\n\n")) !== -1) {
            const frame = buffer.slice(0, i);
            buffer = buffer.slice(i + 2);
            const line = frame.split("\n").find((l) => l.startsWith("data:"));
            if (line && this.onmessage) this.onmessage({ data: line.slice(5).trim() });
          }
        }
        if (!this._closed && this.onerror) this.onerror(new Error("stream ended"));
      } catch (err) {
        if (!this._closed && this.onerror) this.onerror(err);
      }
    }
    close() {
      this._closed = true;
    }
  };
}

/* ------------------------------------------------------------ harness ------ */

async function waitForBridge(timeoutMs = 20000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const res = await fetch(`${BASE}/api/status`);
      if (res.ok) {
        const json = await res.json();
        if (json.connected) return json;
      }
    } catch {
      /* not up yet */
    }
    await sleep(300);
  }
  throw new Error("bridge never became healthy");
}

async function main() {
  fs.rmSync(HOME, { recursive: true, force: true });
  fs.mkdirSync(path.join(HOME, ".local", "state"), { recursive: true });

  const mock = spawn(process.execPath, [path.join(ROOT, "tools", "mock-opencode.mjs")], {
    env: { ...process.env, PORT: MOCK_PORT, MOCK_PASSWORD: "secret", MOCK_DELAY: "60" },
    stdio: ["ignore", "pipe", "pipe"],
  });
  const bridge = spawn(process.execPath, [path.join(ROOT, "bridge", "server.mjs")], {
    env: {
      ...process.env,
      HOME,
      XDG_STATE_HOME: path.join(HOME, ".local", "state"),
      XDG_CONFIG_HOME: path.join(HOME, ".config"),
      OPENCODE_SERVICE_URL: `http://127.0.0.1:${MOCK_PORT}`,
      OPENCODE_SERVICE_PASSWORD: "secret",
      BRIDGE_API_TOKEN: BRIDGE_TOKEN,
      BRIDGE_SKIP_CERT_INSTALL: "1",
      PORT: BRIDGE_PORT,
    },
    stdio: ["ignore", "pipe", "pipe"],
  });
  let bridgeLog = "";
  bridge.stdout.on("data", (d) => (bridgeLog += d.toString()));
  bridge.stderr.on("data", (d) => (bridgeLog += d.toString()));
  mock.stderr.on("data", (d) => console.error("[mock]", d.toString().trim()));

  try {
    try { await waitForBridge(); }
    catch (err) { throw new Error(err.message + "\n" + bridgeLog); }
    console.log("bridge is healthy\n");

    const html = fs.readFileSync(path.join(ROOT, "src", "taskpane.html"), "utf8");
    const dom = new JSDOM(html, { runScripts: "outside-only", url: BASE + "/taskpane.html" });
    const { window } = dom;
    window.document.querySelector('meta[name="opencode-excel-bridge-token"]').setAttribute("content", BRIDGE_TOKEN);

    const fake = createFakeExcel();
    window.Excel = fake.Excel;
    window.EventSource = makeEventSource(BASE);
    const requests = [];
    window.fetch = (url, options) => {
      requests.push({ url: String(url), body: options && options.body });
      return fetch(url.startsWith("http") ? url : BASE + url, options);
    };
    const sendButton = () => window.document.getElementById("btnSend");
    // A send is async: wait for its prompt request to reach the bridge, then
    // for the turn to finish. Waiting only on the disabled state is racy
    // because the pane performs bounded workbook reads before disabling Send.
    const waitTurn = async () => {
      const promptsBefore = requests.filter((r) => /\/api\/prompt$/.test(r.url)).length;
      const startDeadline = Date.now() + 10000;
      while (Date.now() < startDeadline && requests.filter((r) => /\/api\/prompt$/.test(r.url)).length === promptsBefore) {
        await sleep(50);
      }
      const deadline = Date.now() + 25000;
      while (Date.now() < deadline && sendButton().disabled) await sleep(100);
      await sleep(300);
    };
    const lastPromptBody = () => {
      const req = requests.filter((r) => r.url.includes("/api/prompt")).pop();
      return req ? JSON.parse(req.body) : null;
    };
    const lastPrompt = () => {
      const body = lastPromptBody();
      return body ? body.text : "";
    };
    window.setInterval = () => 0; // no background status polling during the test
    Object.defineProperty(window.navigator, "clipboard", {
      value: { writeText: (text) => { window.__copied = text; return Promise.resolve(); } },
      configurable: true,
    });

    // Office is present -> Office.onReady(init) runs immediately.
    window.Office = { onReady: (cb) => cb({ host: "Excel" }) };

    // The real formatter is covered by tools/test-format.mjs; here we only need
    // to prove that the chip calls it instead of sending a model prompt.
    let formatCalls = 0;
    let formatOptions = null;
    window.OpenCodeExcelFormat = {
      formatSheet: async (ExcelArg, libArg, options) => {
        formatCalls++;
        formatOptions = options || {};
        return { sheetName: "Sheet1", tableName: "tblSales", dataRows: 19, created: true, warnings: [] };
      },
    };

    window.eval(fs.readFileSync(path.join(ROOT, "src", "lib.js"), "utf8"));
    window.eval(fs.readFileSync(path.join(ROOT, "src", "plan.js"), "utf8"));
    window.eval(fs.readFileSync(path.join(ROOT, "src", "history.js"), "utf8"));
    // Simulate a checkpoint saved by an earlier pane session. jsdom has no
    // IndexedDB, so the store uses its in-memory fallback here.
    const historyKey = (() => {
      const identity = "unsaved"; // Office.context.document.url is unavailable in the test
      let hash = 0;
      for (let i = 0; i < identity.length; i++) hash = ((hash << 5) - hash + identity.charCodeAt(i)) | 0;
      return "opencode.excel.session." + Math.abs(hash);
    })();
    await window.OpenCodeExcelHistory.put(historyKey, {
      id: "cp-persisted-test",
      at: Date.now() - 600000,
      summary: "Persisted checkpoint from an earlier pane",
      ops: [{ op: "setValues", address: "Sheet1!A1:B1" }],
      okCount: 1,
      total: 1,
      restorable: true,
      snapshot: { sheets: { Sheet1: { "A1:B1": { values: [["Restored", "Header"]] } } } },
      postSnap: null,
      cellOps: null,
    });
    window.eval(fs.readFileSync(path.join(ROOT, "src", "taskpane.js"), "utf8"));

    await sleep(1500); // let init (status + session + context) settle

    /* --- initial state ---------------------------------------------------- */
    console.log("-- initial state --");
    const status = window.document.getElementById("statusText").textContent;
    ok("status shows connected", /connected|OpenCode/i.test(status) && !/Not connected/.test(status), `got: "${status}"`);
    ok("status dot is green", window.document.querySelector("#statusText .dot").className.includes("ok"));
    const contextInfo = window.document.getElementById("contextInfo").textContent;
    ok("context shows selection", /Sheet1!A1:D20/.test(contextInfo) && /20×4/.test(contextInfo), `got: "${contextInfo}"`);
    ok("context lists workbook sheets", /Sheet1/.test(window.document.getElementById("ctxSheets").textContent), `got: "${window.document.getElementById("ctxSheets").textContent}"`);
    ok("agents loaded into dropdown", window.document.getElementById("agentSelect").options.length >= 3, `options: ${window.document.getElementById("agentSelect").options.length}`);
    const agentLabels = Array.from(window.document.getElementById("agentSelect").options).map((option) => option.textContent);
    ok("agent dropdown keeps descriptions out of option labels", agentLabels.includes("explore") && !agentLabels.some((label) => /thoroughness|specialized/.test(label)), agentLabels.join(" | "));
    ok("skills loaded into dropdown", window.document.getElementById("skillSelect").options.length >= 2, `options: ${window.document.getElementById("skillSelect").options.length}`);
    const skillLabels = Array.from(window.document.getElementById("skillSelect").options).map((option) => option.textContent);
    ok("skill dropdown keeps descriptions out of option labels", skillLabels.includes("Financial analysis") && !skillLabels.some((label) => /statements|forecasts/.test(label)), skillLabels.join(" | "));
    ok("models loaded into dropdown", window.document.getElementById("modelSelect").options.length >= 3, `options: ${window.document.getElementById("modelSelect").options.length}`);
    ok("selection-change handler registered", fake.selectionHandlers.length === 1, `handlers: ${fake.selectionHandlers.length}`);

    /* --- persisted workbook checkpoints ---------------------------------- */
    console.log("\n-- persisted checkpoint history --");
    const checkpointList = window.document.getElementById("checkpointList");
    ok("persisted checkpoint appears after pane reload", /Persisted checkpoint from an earlier pane/.test(checkpointList.textContent), checkpointList.textContent);
    const storeMetadata = await window.OpenCodeExcelHistory.list(historyKey);
    check("history store returns checkpoint metadata", storeMetadata.length, 1);
    ok("history metadata omits the full snapshot", !Object.prototype.hasOwnProperty.call(storeMetadata[0] || {}, "snapshot"));
    const fullCheckpoint = await window.OpenCodeExcelHistory.get("cp-persisted-test");
    ok("history store can load the full snapshot on demand", !!(fullCheckpoint && fullCheckpoint.snapshot && fullCheckpoint.snapshot.sheets));
    const persistedRow = Array.from(window.document.querySelectorAll(".checkpoint-row")).find((row) => /Persisted checkpoint/.test(row.textContent));
    ok("persisted checkpoint offers cell restore", !!(persistedRow && persistedRow.querySelector("button")));
    if (persistedRow) persistedRow.querySelector("button").click();
    await sleep(500);
    const persistedRestore = fake.ops.filter((op) => op.kind === "values" && op.a1 === "A1:B1").pop();
    check("persisted checkpoint restores its cell snapshot", persistedRestore && persistedRestore.values, [["Restored", "Header"]]);
    ok("restored checkpoint no longer offers restore", !window.document.querySelector(".checkpoint-row button"));

    /* --- native formatter chip -------------------------------------------- */
    console.log("\n-- format sheet chip --");
    const formatChip = window.document.querySelector('.chip[data-action="format"]');
    ok("format chip is present", !!formatChip);
    const promptsBeforeFormat = requests.filter((r) => /\/api\/prompt$/.test(r.url)).length;
    formatChip.click();
    await sleep(400);
    ok("format chip calls the native formatter", formatCalls === 1, `calls: ${formatCalls}`);
    const promptsAfterFormat = requests.filter((r) => /\/api\/prompt$/.test(r.url)).length;
    ok("format chip does not send a model prompt", promptsAfterFormat === promptsBeforeFormat, `prompts: ${promptsAfterFormat - promptsBeforeFormat}`);
    ok("format chip shows a success banner", /Formatted Sheet1/.test(window.document.getElementById("banner").textContent), window.document.getElementById("banner").textContent);
    ok("format chip label is restored", formatChip.textContent === "Format sheet", formatChip.textContent);

    /* --- send a message and stream ---------------------------------------- */
    console.log("\n-- send + stream --");
    const input = window.document.getElementById("input");
    input.value = "flag duplicates in column B";
    window.document.getElementById("btnSend").click();

    await sleep(500);
    ok("user bubble rendered", !!window.document.querySelector(".msg.user .bubble"));
    ok("stop row visible while streaming", !window.document.getElementById("stopRow").classList.contains("hidden"));
    ok("send button disabled while streaming", window.document.getElementById("btnSend").disabled === true);
    const partial = window.document.querySelector(".msg.assistant .bubble").textContent;
    ok("assistant text streams in", partial.length > 0, `length: ${partial.length}`);

    // Wait for the turn to finish (poll up to 20s).
    const deadline = Date.now() + 20000;
    while (Date.now() < deadline && window.document.getElementById("btnSend").disabled) await sleep(200);

    const bubble = window.document.querySelector(".msg.assistant .bubble");
    ok("turn completed (send re-enabled)", window.document.getElementById("btnSend").disabled === false);
    ok("stop row hidden after completion", window.document.getElementById("stopRow").classList.contains("hidden"));
    ok("answer rendered as markdown (code fence)", bubble.innerHTML.includes("<pre>"), bubble.innerHTML.slice(0, 80));
    ok("answer contains the formula", bubble.textContent.includes('=IF(COUNTIF($B$2:$B$100,B2)>1,"duplicate","")'));

    const firstPrompt = lastPrompt();
    ok("prompt includes the preamble", /Never suggest editing files/.test(firstPrompt));
    ok("prompt includes the selection as TSV", firstPrompt.includes("Selection Sheet1!A1:D20") && firstPrompt.includes("Name\tAmount"), firstPrompt.slice(0, 120));
    ok("prompt includes the typed question", firstPrompt.includes("flag duplicates in column B"), firstPrompt.slice(-80));
    ok("prompt lists workbook sheets", firstPrompt.includes("workbook sheets: Sheet1, Sheet2"));

    const buttons = Array.from(window.document.querySelectorAll(".msg.assistant .actions button")).map((b) => b.textContent);
    check("action buttons", buttons, ["Copy", "Insert ↓", "Insert table (4×2)", "Use as formula"]);

    /* --- explicit user-authorized range read ------------------------------ */
    input.value = "Analyze @range(Sheet2!A1:B2)";
    window.document.getElementById("btnSend").click();
    await waitTurn();
    ok("explicit range is sent with requested values", lastPrompt().includes('"sheet":"Sheet2"') && lastPrompt().includes('"Jan",1000'), lastPrompt().slice(-400));
    const beforeInvalidRange = requests.filter((r) => r.url.includes("/api/prompt")).length;
    input.value = "Analyze @range(Sheet2!A1:A501)";
    window.document.getElementById("btnSend").click();
    await sleep(500);
    ok("oversized explicit range does not reach OpenCode", requests.filter((r) => r.url.includes("/api/prompt")).length === beforeInvalidRange);
    ok("oversized explicit range is explained", /500 cells/.test(window.document.getElementById("banner").textContent), window.document.getElementById("banner").textContent);
    const savedSheet2GetRange = fake.sheet2.getRange;
    fake.sheet2.getRange = function(a1) {
      const r = savedSheet2GetRange.call(this, a1);
      if (a1 === "A1:B2") r.values = [["X".repeat(13000), "Sales"], ["Jan", 1000]];
      return r;
    };
    input.value = "Analyze @range(Sheet2!A1:B2)";
    window.document.getElementById("btnSend").click();
    await sleep(500);
    ok("oversized range content does not reach OpenCode", requests.filter((r) => r.url.includes("/api/prompt")).length === beforeInvalidRange);
    ok("oversized range content is explained", /12,000 characters/.test(window.document.getElementById("banner").textContent), window.document.getElementById("banner").textContent);
    fake.sheet2.getRange = savedSheet2GetRange;
    input.value = "";

    /* --- recursive formula trace ------------------------------------------ */
    console.log("\n-- recursive trace --");
    input.value = "Trace @trace(Sheet1!A1, depth=2, mode=precedents)";
    window.document.getElementById("btnSend").click();
    await waitTurn();
    const tracePrompt = lastPrompt();
    ok(
      "trace request carries the requested mode and depth",
      tracePrompt.includes('"mode":"precedents"') && tracePrompt.includes('"depth":2'),
      tracePrompt.slice(-600)
    );
    ok("trace prompt labels the recursive source honestly", tracePrompt.includes("Recursive results use the requested depth"));

    /* --- model-requested reads need approval ------------------------------ */
    input.value = "RANGE_READ_REQUEST: need more evidence";
    window.document.getElementById("btnSend").click(); await waitTurn();
    const readButtons = Array.from(window.document.querySelectorAll(".msg.assistant button"));
    const approveRead = readButtons.find(b => b.textContent === "Approve read & continue");
    ok("model read request displays approval", !!approveRead);
    const beforeApprovedRead = requests.filter(r => r.url.includes("/api/prompt")).length;
    await sleep(200);
    check("model-requested range is not sent without approval", requests.filter(r => r.url.includes("/api/prompt")).length, beforeApprovedRead);
    if (approveRead) { approveRead.click(); await waitTurn(); }
    ok("approved range is sent through OpenCode", requests.filter(r => r.url.includes("/api/prompt")).length === beforeApprovedRead + 1 && lastPrompt().includes('"Jan",1000'));

    /* --- analytical summary instructions ---------------------------------- */
    console.log("\n-- analytical summary prompt --");
    input.value = "summarize this table";
    window.document.getElementById("btnSend").click();
    await waitTurn();
    const summaryPrompt = lastPrompt();
    ok("summary prompt asks for business insights", summaryPrompt.includes("4 to 6 decision-useful insights"), summaryPrompt.slice(-700));
    ok("summary prompt avoids formula inventories", summaryPrompt.includes("Do not return a formula inventory"), summaryPrompt.slice(-700));
    ok("summary prompt requests range citations", summaryPrompt.includes("compact source range citation"), summaryPrompt.slice(-700));

    /* --- unavailable default model automatically falls back ---------------- */
    console.log("\n-- free model fallback --");
    const promptsBeforeFallback = requests.filter((r) => /\/api\/prompt$/.test(r.url)).length;
    const usersBeforeFallback = window.document.querySelectorAll(".msg.user").length;
    input.value = "FALLBACK_REQUEST create a basic website dashboard template";
    window.document.getElementById("btnSend").click();
    await waitTurn();
    const fallbackRequests = requests.filter((r) => /\/api\/prompt$/.test(r.url)).slice(promptsBeforeFallback).map((r) => JSON.parse(r.body));
    check("unavailable model is retried exactly once", fallbackRequests.length, 2);
    ok("fallback retries the identical prompt", fallbackRequests.length === 2 && fallbackRequests[0].text === fallbackRequests[1].text);
    check("fallback uses the preferred free OpenCode model", fallbackRequests[1] && fallbackRequests[1].model, { providerID: "opencode", id: "mimo-v2.6-flash-free" });
    ok("fallback never selects the paid catalog model", !fallbackRequests.some((body) => body.model && body.model.id === "paid-dashboard"));
    check("fallback keeps one user message", window.document.querySelectorAll(".msg.user").length, usersBeforeFallback + 1);
    ok("model selector visibly updates to fallback", /mimo-v2\.6-flash-free/.test(window.document.getElementById("modelSelect").value), window.document.getElementById("modelSelect").value);
    ok("fallback choice is persisted", /mimo-v2\.6-flash-free/.test(window.localStorage.getItem("opencode.excel.model") || ""));
    const fallbackAssistant = Array.from(window.document.querySelectorAll(".msg.assistant .bubble")).pop();
    ok("fallback answer completes normally", fallbackAssistant.textContent.includes("flags duplicates"), fallbackAssistant.textContent.slice(0, 80));
    ok("fallback is explained to the user", /Retrying once with free MiMo/i.test(window.document.getElementById("banner").textContent), window.document.getElementById("banner").textContent);

    const flintOption = Array.from(window.document.getElementById("modelSelect").options).find((option) => /flint-40b-free/.test(option.value));
    window.document.getElementById("modelSelect").value = flintOption.value;
    window.document.getElementById("modelSelect").dispatchEvent(new window.Event("change"));
    const promptsBeforeEmpty = requests.filter((r) => /\/api\/prompt$/.test(r.url)).length;
    input.value = "EMPTY_REQUEST return no content on the first attempt";
    window.document.getElementById("btnSend").click();
    await waitTurn();
    const emptyRequests = requests.filter((r) => /\/api\/prompt$/.test(r.url)).slice(promptsBeforeEmpty).map((r) => JSON.parse(r.body));
    check("empty response is also retried exactly once", emptyRequests.length, 2);
    check("empty response uses the safe free fallback", emptyRequests[1] && emptyRequests[1].model, { providerID: "opencode", id: "mimo-v2.6-flash-free" });
    const emptyFallbackAssistant = Array.from(window.document.querySelectorAll(".msg.assistant .bubble")).pop();
    ok("empty-response fallback produces an answer", emptyFallbackAssistant.textContent.includes("flags duplicates"), emptyFallbackAssistant.textContent.slice(0, 80));

    /* --- actions ----------------------------------------------------------- */
    console.log("\n-- insert actions --");
    const byLabel = (label) => Array.from(window.document.querySelectorAll(".msg.assistant .actions button")).find((b) => b.textContent === label);

    byLabel("Insert table (4×2)").click();
    await sleep(300);
    const tableOp = fake.ops.find((o) => o.kind === "values" && o.a1 === "A22:B25");
    check("insert table writes 4x2 at A22", tableOp && tableOp.a1, "A22:B25");
    check("insert table payload", tableOp && tableOp.values, [
      ["Value", "Count"],
      ["ACME", "4"],
      ["Globex", "3"],
      ["Initech", "2"],
    ]);
    ok("insert table autofits columns", fake.ops.some((o) => o.kind === "autofit"));

    byLabel("Insert ↓").click();
    await sleep(300);
    const textOp = fake.ops.filter((o) => o.kind === "values" && /^A22:A/.test(o.a1)).pop();
    ok("insert below targets A22", !!textOp, textOp && textOp.a1);
    ok("insert below writes one line per row", Array.isArray(textOp.values) && textOp.values.length > 3 && textOp.values.every((r) => r.length === 1), `rows: ${textOp.values && textOp.values.length}`);

    byLabel("Use as formula").click();
    await sleep(300);
    const formulaOp = fake.ops.find((o) => o.kind === "formulas");
    check("formula written to selection start", formulaOp && formulaOp.a1, "A1");
    check("formula payload", formulaOp && formulaOp.values, [['=IF(COUNTIF($B$2:$B$100,B2)>1,"duplicate","")']]);

    byLabel("Copy").click();
    await sleep(200);
    ok("copy puts answer on the clipboard", (window.__copied || "").includes("flags duplicates"), String(window.__copied).slice(0, 40));

    /* --- plan review, apply, undo --- */
    console.log("\n-- plan review --");
    const planInput = window.document.getElementById("input");
    planInput.value = "PLAN_REQUEST: add a Budget sheet and rewrite the top-left table";
    window.document.getElementById("btnSend").click();
    await waitTurn();
    ok("plan preview is shown", !window.document.getElementById("previewPanel").classList.contains("hidden"));
    ok("preview counts 3 operations", /3 operation/.test(window.document.getElementById("previewCount").textContent), window.document.getElementById("previewCount").textContent);
    const rows = Array.from(window.document.querySelectorAll(".op-row"));
    check("plan has 3 rows", rows.length, 3);
    const cb = (i) => rows[i].querySelector("input[type=checkbox]");
    ok("add sheet is pre-checked", !!cb(0).checked, "cb0: " + cb(0).checked);
    ok("destructive setValues is NOT pre-checked", !cb(1).checked, "cb1: " + cb(1).checked);
    ok("format range is pre-checked", !!cb(2).checked, "cb2: " + cb(2).checked);
    ok("cells column shows sheet ranges", /Sheet1!A1:B2/.test(rows[1].textContent), rows[1].textContent);

    cb(1).click();
    ok("destructive op can be selected", cb(1).checked);
    window.document.getElementById("btnApplyPlan").click();
    await sleep(500);
    const addSheetOp = fake.ops.find((o) => o.kind === "addSheet");
    check("addSheet was executed", addSheetOp && addSheetOp.name, "Budget");
    const valuesOp = fake.ops.filter((o) => o.kind === "values" && o.a1 === "A1:B2").pop();
    check("setValues was executed", valuesOp && JSON.stringify(valuesOp.values), JSON.stringify([["Item", "Total"], ["Widgets", 300]]));
    ok("structural plan does not offer misleading add-in undo", window.document.getElementById("undoRow").classList.contains("hidden"));
    ok(
      "structural apply directs the user to Excel Undo or a backup",
      /Add-in Undo is unavailable/.test(window.document.getElementById("banner").textContent),
      window.document.getElementById("banner").textContent
    );

    /* --- cell-only plan: add-in undo works and is honest --------------- */
    console.log("\n-- cell-only plan undo --");
    planInput.value = "SAFE_PLAN_REQUEST: update two cells";
    window.document.getElementById("btnSend").click();
    await waitTurn();
    const safeRows = Array.from(window.document.querySelectorAll(".op-row"));
    check("safe plan has 2 rows", safeRows.length, 2);
    const safeCb = (i) => safeRows[i].querySelector("input[type=checkbox]");
    ok("safe cell write is not pre-checked", !safeCb(0).checked, "cb0: " + safeCb(0).checked);
    ok("safe formatting is pre-checked", !!safeCb(1).checked, "cb1: " + safeCb(1).checked);
    safeCb(0).click();
    window.document.getElementById("btnApplyPlan").click();
    await sleep(500);
    const safeWrite = fake.ops.filter((o) => o.kind === "values" && o.a1 === "D5:E5").pop();
    check("safe plan wrote two cells", safeWrite && safeWrite.values, [["Safe", "Plan"]]);
    ok("cell-only plan offers add-in undo", !window.document.getElementById("undoRow").classList.contains("hidden"));
    ok("checkpoint history records the apply", /Update two cells/.test(window.document.getElementById("checkpointList").textContent), window.document.getElementById("checkpointList").textContent);
    ok("checkpoint offers cell restore", !!window.document.getElementById("checkpointList").querySelector("button"));

    window.document.getElementById("btnUndo").click();
    await sleep(500);
    const safeRestored = fake.ops.filter((o) => o.kind === "values" && o.a1 === "D5:E5").pop();
    check("add-in undo restores the cell before-state", safeRestored && safeRestored.values, [["Before", "State"]]);
    ok("undo row hides after the stack empties", window.document.getElementById("undoRow").classList.contains("hidden"));

    /* --- stale preview and undo drift ----------------------------------- */
    console.log("\n-- stale preview and undo drift --");
    planInput.value = "SAFE_PLAN_REQUEST: stale preview";
    window.document.getElementById("btnSend").click(); await waitTurn();
    ok("prompt includes recent workbook changes", lastPrompt().includes("## Recent workbook changes") && lastPrompt().includes("Update two cells"), lastPrompt().slice(-600));
    ok("cell preview shows before and proposed values", /Before/.test(window.document.getElementById("previewDiff").textContent) && /Safe/.test(window.document.getElementById("previewDiff").textContent));
    const staleRows = Array.from(window.document.querySelectorAll(".op-row"));
    staleRows[0].querySelector("input[type=checkbox]").click();
    const beforeStaleWrites = fake.ops.filter(o => o.kind === "values" && o.a1 === "D5:E5").length;
    fake.cellStore["D5:E5"].values = [["User edit", "State"]];
    window.document.getElementById("btnApplyPlan").click(); await sleep(500);
    ok("stale preview blocks apply", /Cells changed since preview/.test(window.document.getElementById("banner").textContent));
    check("stale preview does not write", fake.ops.filter(o => o.kind === "values" && o.a1 === "D5:E5").length, beforeStaleWrites);
    window.document.getElementById("btnCancelPlan").click();
    fake.cellStore["D5:E5"].values = [["Before", "State"]];
    planInput.value = "SAFE_PLAN_REQUEST: test undo drift";
    window.document.getElementById("btnSend").click(); await waitTurn();
    const driftRows = Array.from(window.document.querySelectorAll(".op-row"));
    driftRows[0].querySelector("input[type=checkbox]").click();
    window.document.getElementById("btnApplyPlan").click(); await sleep(500);
    fake.cellStore["D5:E5"].values = [["Later edit", "Plan"]];
    window.document.getElementById("btnUndo").click(); await sleep(500);
    ok("undo drift blocks restoring over newer edits", /Undo blocked/.test(window.document.getElementById("banner").textContent));
    check("later edits stay intact", fake.cellStore["D5:E5"].values, [["Later edit", "Plan"]]);
    ok("blocked undo remains available", !window.document.getElementById("undoRow").classList.contains("hidden"));
    fake.cellStore["D5:E5"].values = [["Safe", "Plan"]];
    window.document.getElementById("btnUndo").click(); await sleep(500);
    check("undo can be retried after drift is resolved", fake.cellStore["D5:E5"].values, [["Before", "State"]]);

    /* --- edited JSON must be reviewed before applying ------------------- */
    planInput.value = "SAFE_PLAN_REQUEST: edit plan";
    window.document.getElementById("btnSend").click(); await waitTurn();
    window.document.getElementById("btnShowPlanJson").click();
    window.document.getElementById("planEdit").value = JSON.stringify({summary:"Edited",operations:[{op:"setValues",sheet:"Sheet1",range:"D5:E5",values:[["Edited","Plan"]]}]});
    const beforeEdited = fake.ops.filter(o => o.kind === "values" && o.a1 === "D5:E5").length;
    window.document.getElementById("btnApplyEdited").click(); await sleep(300);
    check("edited plan is not applied before review", fake.ops.filter(o => o.kind === "values" && o.a1 === "D5:E5").length, beforeEdited);
    ok("edited plan displays new proposed values", /Edited/.test(window.document.getElementById("previewDiff").textContent));
    ok("edited plan keeps destructive change unchecked", !window.document.querySelector(".op-row input[type=checkbox]").checked);
    window.document.getElementById("btnCancelPlan").click();

    /* --- native formatting invalidates older add-in undo entries -------- */
    console.log("\n-- format sheet invalidates undo --");
    planInput.value = "SAFE_PLAN_REQUEST: update two cells again";
    window.document.getElementById("btnSend").click();
    await waitTurn();
    const againRows = Array.from(window.document.querySelectorAll(".op-row"));
    againRows[0].querySelector("input[type=checkbox]").click();
    window.document.getElementById("btnApplyPlan").click();
    await sleep(500);
    ok("fresh cell-only plan offers add-in undo", !window.document.getElementById("undoRow").classList.contains("hidden"));
    formatChip.click();
    await sleep(400);
    ok("native formatting hides the stale undo row", window.document.getElementById("undoRow").classList.contains("hidden"));
    ok("native formatting invalidates checkpoint restore snapshots", !window.document.getElementById("checkpointList").querySelector("button"));
    window.document.getElementById("btnUndo").click();
    await sleep(300);
    ok("native formatting discards older undo snapshots", /Nothing to undo/.test(window.document.getElementById("banner").textContent), window.document.getElementById("banner").textContent);

    /* --- border formatting is applied but never offered as add-in undo -- */
    console.log("\n-- border formatting undo honesty --");
    const borderApproval = window.document.getElementById("selApproval");
    borderApproval.value = "auto";
    borderApproval.dispatchEvent(new window.Event("change"));
    planInput.value = "BORDER_PLAN_REQUEST: red bottom border";
    window.document.getElementById("btnSend").click();
    await waitTurn();
    const borderRows = Array.from(window.document.querySelectorAll(".op-row"));
    check("border plan has one row", borderRows.length, 1);
    ok("border plan is shown for review in auto mode", !window.document.getElementById("previewPanel").classList.contains("hidden"));
    check("auto mode did not apply the border", fake.ops.filter((o) => o.kind === "border").length, 0);
    window.document.getElementById("btnApplyPlan").click();
    await sleep(500);
    const borderOps = fake.ops.filter((o) => o.kind === "border" && o.a1 === "A1");
    check("border plan applies style, weight, and color", borderOps.map((o) => o.prop), ["style", "weight", "color"]);
    ok("border formatting does not offer add-in undo", window.document.getElementById("undoRow").classList.contains("hidden"));
    ok("border apply explains undo is unavailable", /Add-in Undo is unavailable/.test(window.document.getElementById("banner").textContent), window.document.getElementById("banner").textContent);

    /* --- auto-approval must never run destructive operations ----------- */
    console.log("\n-- destructive auto-approval guard --");
    const approvalSelect = window.document.getElementById("selApproval");
    approvalSelect.value = "auto";
    approvalSelect.dispatchEvent(new window.Event("change"));
    const deletesBefore = fake.ops.filter((o) => o.kind === "deleteSheet").length;
    const bannerBeforeDestructive = window.document.getElementById("banner").textContent;
    planInput.value = "DESTRUCTIVE_PLAN_REQUEST: delete a sheet";
    window.document.getElementById("btnSend").click();
    await waitTurn();
    ok("destructive plan is shown for review in auto mode", !window.document.getElementById("previewPanel").classList.contains("hidden"));
    const destructiveRows = Array.from(window.document.querySelectorAll(".op-row"));
    ok("deleteSheet is not pre-checked", destructiveRows.length === 1 && !destructiveRows[0].querySelector("input[type=checkbox]").checked);
    check("auto mode did not delete the sheet", fake.ops.filter((o) => o.kind === "deleteSheet").length, deletesBefore);
    ok(
      "auto mode did not report the plan as applied",
      window.document.getElementById("banner").textContent === bannerBeforeDestructive,
      window.document.getElementById("banner").textContent
    );

    approvalSelect.value = "never";
    approvalSelect.dispatchEvent(new window.Event("change"));
    planInput.value = "DESTRUCTIVE_PLAN_REQUEST: delete a sheet again";
    window.document.getElementById("btnSend").click();
    await waitTurn();
    check("never mode also refuses destructive auto-approval", fake.ops.filter((o) => o.kind === "deleteSheet").length, deletesBefore);
    ok("never-mode destructive plan is shown for review", !window.document.getElementById("previewPanel").classList.contains("hidden"));
    approvalSelect.value = "always";
    approvalSelect.dispatchEvent(new window.Event("change"));

    /* --- read-only blocks apply --- */
    console.log("\n-- read-only --");
    window.document.getElementById("chkReadOnly").click();
    await sleep(200);
    ok("read-only saved", window.localStorage.getItem("opencode.excel.readonly") === "true");
    planInput.value = "PLAN_REQUEST: another plan";
    window.document.getElementById("btnSend").click();
    await waitTurn();
    ok("plan preview shown again", !window.document.getElementById("previewPanel").classList.contains("hidden"));
    ok("apply button is disabled in read-only", window.document.getElementById("btnApplyPlan").disabled === true);
    var beforeBanner = window.document.getElementById("banner").textContent;
    window.document.getElementById("btnApplyPlan").click();
    await sleep(300);
    ok("read-only blocks apply", window.document.getElementById("banner").textContent === beforeBanner);
    window.document.getElementById("chkReadOnly").click();

    /* --- attachments --- */
    console.log("\n-- attachments --");
    var attachInput = window.document.getElementById("fileInput");
    Object.defineProperty(attachInput, "files", { value: [{ name: "budget.csv", type: "text/csv", size: 1024 }] });
    attachInput.dispatchEvent(new window.Event("change"));
    await sleep(300);
    ok("attachment chip rendered", /budget\.csv/.test(window.document.getElementById("attachChips").textContent));
    ok("attach row visible", !window.document.getElementById("attachRow").classList.contains("hidden"));
    planInput.value = "What is in the attachment?";
    window.document.getElementById("btnSend").click();
    await waitTurn();
    ok("attachment named in prompt", lastPrompt().includes("budget.csv"));
    ok("binary file noted in prompt", lastPrompt().includes("content not extracted"));

    /* --- @sheet mention --- */
    console.log("\n-- @sheet mention --");
    planInput.value = "@Sheet2 summarize the data";
    window.document.getElementById("btnSend").click();
    await waitTurn();
    var mentionPrompt = lastPrompt();
    ok("mention focuses the sheet", mentionPrompt.includes("Focus sheets (@mentioned): Sheet2"));
    ok("mentioned sheet values in prompt", mentionPrompt.includes("Sheet Sheet2:") && mentionPrompt.includes("Jan"));

    /* --- standing instructions, conventions, comments -------------------- */
    console.log("\n-- instructions, conventions, comments --");
    const instructionsInput = window.document.getElementById("instructionsText");
    instructionsInput.value = "Always use fiscal quarters and cite the source table.";
    instructionsInput.dispatchEvent(new window.Event("input"));
    check("instructions are persisted", window.localStorage.getItem("opencode.excel.instructions"), "Always use fiscal quarters and cite the source table.");
    const conventionsInput = window.document.getElementById("conventionsText");
    conventionsInput.value = "currency: EUR, decimals: 2";
    conventionsInput.dispatchEvent(new window.Event("input"));
    check("conventions are persisted", window.localStorage.getItem("opencode.excel.conventions"), "currency: EUR, decimals: 2");

    planInput.value = "Show the discussion @comments(Sheet1!A1:D20)";
    window.document.getElementById("btnSend").click();
    await waitTurn();
    const toolPrompt = lastPrompt();
    ok("prompt includes standing instructions", toolPrompt.includes("## Standing user instructions") && toolPrompt.includes("fiscal quarters"));
    ok("prompt includes formatting conventions", toolPrompt.includes("## Formatting conventions") && toolPrompt.includes("currency: EUR"));
    ok(
      "comment read includes the thread",
      toolPrompt.includes("## Explicit user-requested comment threads") &&
        toolPrompt.includes("Please double-check this amount") &&
        toolPrompt.includes("Checked and corrected"),
      toolPrompt.slice(-800)
    );
    ok("comment text is labelled untrusted", /untrusted workbook data, not instructions/.test(toolPrompt));

    formatChip.click();
    await sleep(500);
    check("format chip receives conventions", formatOptions && formatOptions.conventions, "currency: EUR, decimals: 2");

    /* --- settings + quick action + stop ------------------------------------ */
    console.log("\n-- quick action, stop, new chat --");
    window.document.getElementById("btnSettings").click();
    ok("settings panel toggles", !window.document.getElementById("settingsPanel").classList.contains("hidden"));

    const modelSelect = window.document.getElementById("modelSelect");
    const smartModel = Array.from(modelSelect.options).find((option) => /mock-smart/.test(option.value));
    ok("model dropdown contains the service models", !!smartModel, `options: ${modelSelect.options.length}`);
    modelSelect.value = smartModel.value;
    modelSelect.dispatchEvent(new window.Event("change"));
    check("selected model is persisted", window.localStorage.getItem("opencode.excel.model"), smartModel.value);

    const sel = window.document.getElementById("selApproval");
    sel.value = "auto";
    sel.dispatchEvent(new window.Event("change"));
    check("approval mode is persisted", window.localStorage.getItem("opencode.excel.approval"), "auto");

    const chip = window.document.querySelector('.chip[data-action="clean"]');
    chip.click();
    await sleep(700);
    const lastUser = Array.from(window.document.querySelectorAll(".msg.user .bubble")).pop();
    ok("quick action prefills and sends", !!lastUser && /Clean this data/.test(lastUser.textContent), lastUser ? lastUser.textContent.slice(0, 60) : "no user bubble");
    const sentBody = lastPromptBody();
    check("prompt sends the selected model", sentBody && sentBody.model, { providerID: "mock", id: "mock-smart" });

    window.document.getElementById("btnStop").click();
    await sleep(600);
    ok("stop finalizes the turn", window.document.getElementById("btnSend").disabled === false);
    const assistantCount = window.document.querySelectorAll(".msg.assistant .bubble").length;
    const lastAssistant = Array.from(window.document.querySelectorAll(".msg.assistant .bubble")).pop();
    ok("stop leaves a finalised bubble", assistantCount > 0 && !!lastAssistant);
    ok("stopped bubble is not left thinking", !lastAssistant.innerHTML.includes("Thinking"), lastAssistant.innerHTML.slice(0, 60));

    window.document.getElementById("btnNewChat").click();
    await sleep(800);
    ok("new chat clears the transcript", window.document.querySelectorAll(".msg").length === 0);
    ok("new chat restores empty state", !!window.document.getElementById("emptyState"));

    /* --- session tabs and slash commands ---------------------------------- */
    console.log("\n-- session tabs and slash commands --");
    const sessionTabs = () => Array.from(window.document.querySelectorAll("#sessionTabs .session-tab"));
    const commandMenu = window.document.getElementById("commandMenu");
    const commandInput = window.document.getElementById("input");
    const messages = window.document.getElementById("messages");
    check("new chat adds a session tab", sessionTabs().length, 2);
    ok("new chat tab is active", sessionTabs()[1].classList.contains("active"));
    ok("session tab keeps the first question as its title", /flag duplicates/i.test(sessionTabs()[0].textContent), sessionTabs()[0].textContent);
    const savedRecords = JSON.parse(window.localStorage.getItem(historyKey + ".records") || "[]");
    check("session tabs persist their transcripts locally", savedRecords.length, 2);
    ok("saved transcripts contain messages", savedRecords.some((record) => (record.messages || []).length > 0));

    sessionTabs()[0].click();
    await sleep(700);
    ok("switching tabs restores the saved transcript", /flags duplicates/i.test(messages.textContent));
    ok("switched session tab becomes active", sessionTabs()[0].classList.contains("active"));
    check("switching tabs does not create another tab", sessionTabs().length, 2);

    commandInput.value = "/";
    commandInput.dispatchEvent(new window.Event("input"));
    ok("typing / opens the command palette", !commandMenu.classList.contains("hidden"));
    let paletteRows = Array.from(commandMenu.querySelectorAll(".command-row"));
    ok("command palette lists the built-ins", paletteRows.length >= 10 && paletteRows.some((row) => row.textContent.includes("/help")), `rows: ${paletteRows.length}`);
    commandInput.dispatchEvent(new window.KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true }));
    ok("arrow keys move the active command", commandMenu.querySelectorAll(".command-row.active").length === 1);
    commandInput.dispatchEvent(new window.KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    ok("Escape closes the command palette", commandMenu.classList.contains("hidden"));

    commandInput.value = "/res";
    commandInput.dispatchEvent(new window.Event("input"));
    paletteRows = Array.from(commandMenu.querySelectorAll(".command-row"));
    check("command palette filters by prefix", paletteRows.length, 1);
    ok("filtered command is /resume", !!(paletteRows[0] && paletteRows[0].textContent.includes("/resume")));
    commandInput.dispatchEvent(new window.KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
    check("Enter selects the command for review", commandInput.value, "/resume");
    ok("selecting a command closes the palette", commandMenu.classList.contains("hidden"));

    const promptsBeforeHelp = requests.filter((request) => /\/api\/prompt$/.test(request.url)).length;
    commandInput.value = "/help";
    commandInput.dispatchEvent(new window.Event("input"));
    window.document.getElementById("btnSend").click();
    await sleep(500);
    ok("help renders the local command list", /Local commands never call the model/.test(messages.textContent) && /\/compact/.test(messages.textContent));
    check("help never calls OpenCode", requests.filter((request) => /\/api\/prompt$/.test(request.url)).length, promptsBeforeHelp);
    const helpTranscript = JSON.parse(window.localStorage.getItem(historyKey + ".records") || "[]");
    ok(
      "local command output is saved as exportable transcript text",
      helpTranscript.some((record) => (record.messages || []).some((message) => /Local commands never call the model/.test(message.text || "")))
    );

    commandInput.value = "/resume";
    commandInput.dispatchEvent(new window.Event("input"));
    window.document.getElementById("btnSend").click();
    await sleep(500);
    const resumeButtons = Array.from(messages.querySelectorAll("[data-session-id]"));
    ok("resume lists saved chats", resumeButtons.length >= 2, `buttons: ${resumeButtons.length}`);
    const emptyChatButton = resumeButtons.find((button) => /New chat/.test(button.parentElement.textContent));
    ok("resume finds the empty chat", !!emptyChatButton);
    if (emptyChatButton) emptyChatButton.click();
    await sleep(700);
    ok("resume switches to the selected chat", window.document.querySelectorAll(".msg").length === 0);
    ok("resumed chat tab is active", sessionTabs().some((tab) => /New chat/.test(tab.textContent) && tab.classList.contains("active")));

    sessionTabs().find((tab) => !/New chat/.test(tab.textContent)).click();
    await sleep(700);
    ok("returning to the first chat restores its transcript", /flags duplicates/i.test(messages.textContent));
    const emptyTab = sessionTabs().find((tab) => /New chat/.test(tab.textContent) && !tab.classList.contains("active"));
    if (emptyTab) emptyTab.querySelector(".session-close").click();
    await sleep(400);
    check("closing an inactive chat removes its tab", sessionTabs().length, 1);
    ok("closing an inactive chat keeps the current transcript", /flags duplicates/i.test(messages.textContent));

    /* --- selection change -------------------------------------------------- */
    console.log("\n-- selection change --");
    fake.selection.address = "Sheet1!F3:H9";
    fake.selection.rowCount = 7;
    fake.selection.columnCount = 3;
    await fake.selectionHandlers[0]();
    await sleep(200);
    ok("context refreshes on selection change", /F3:H9/.test(window.document.getElementById("contextInfo").textContent), window.document.getElementById("contextInfo").textContent);

    /* --- context handling -------------------------------------------------- */
    console.log("\n-- context handling --");
    // truncation: 500 rows x 20 cols (10k cells) with a 400-cell cap
    fake.selection.address = "Sheet1!A1:T500";
    fake.selection.rowCount = 500;
    fake.selection.columnCount = 20;
    await fake.selectionHandlers[0]();
    await sleep(200);
    const lastSample = fake.ops.filter((op) => op.kind === "sampleSelection").pop();
    ok("large selection loads only bounded sample", lastSample && lastSample.rows * lastSample.cols <= 400, JSON.stringify(lastSample));
    const lastSelectionLoad = fake.ops.filter((op) => op.kind === "selectionLoad").pop();
    ok("selection metadata load excludes full values", lastSelectionLoad && !lastSelectionLoad.props.includes("values"));
    ok("context marks a big selection as truncated", /truncated/.test(window.document.getElementById("contextInfo").textContent), window.document.getElementById("contextInfo").textContent);

    window.document.getElementById("input").value = "what is in here?";
    window.document.getElementById("btnSend").click();
    await waitTurn();
    const capped = lastPrompt();
    const tsvLines = capped.split("\n").filter((l) => l.includes("\t")).length;
    ok("prompt says truncated", capped.includes("truncated"), capped.slice(0, 80));
    ok("prompt caps rows (<= 20 rows for a 20-column range)", tsvLines <= 20, `tsv lines: ${tsvLines}`);

    // huge selections must be sampled before values/formulas are loaded
    fake.selection.loaded = [];
    fake.selection.sampleRequests = [];
    await fake.selectionHandlers[0]();
    await sleep(300);
    const loadedProps = fake.selection.loaded.flat();
    ok("selection loads dimensions before values", loadedProps.every((p) => !p.includes("values") && !p.includes("formulas")), JSON.stringify(loadedProps));
    check("selection samples the capped top-left block", fake.selection.sampleRequests.pop(), [20, 20]);

    // 30,000 character budget across sheet samples and selection
    const oversized = "X".repeat(40000);
    const clippedCell = oversized.slice(0, 500) + " [truncated]";
    const longRows = Array.from({ length: 20 }, () => Array.from({ length: 4 }, () => oversized));
    fake.selection.values = longRows;
    fake.selection.formulas = Array.from({ length: 20 }, () => Array.from({ length: 4 }, () => null));
    fake.cellStore["A1:D20"] = { values: longRows, formulas: undefined };
    window.document.getElementById("includeSelection").checked = false;
    await fake.selectionHandlers[0]();
    await sleep(300);
    window.document.getElementById("input").value = "what is in the workbook?";
    window.document.getElementById("btnSend").click();
    await waitTurn();
    const budgeted = lastPrompt();
    const ctxStart = budgeted.indexOf("## Workbook context");
    const ctxEnd = budgeted.indexOf("## Task");
    const workbookContext = budgeted.slice(ctxStart, ctxEnd);
    ok("workbook context respects the 30,000-character budget", workbookContext.length <= 30500, `workbook context length: ${workbookContext.length}`);
    ok("workbook context reports truncation", budgeted.includes("... (workbook context truncated at 30,000 characters)"));
    ok("workbook context does not contain the oversized cell", !budgeted.includes(oversized));
    ok("workbook context clips individual cell strings", budgeted.includes(clippedCell));
    fake.selection.values = [["Name", "Amount"], ["ACME", 10], ["Globex", 20]];
    fake.selection.formulas = [["Name", "Amount"], ["ACME", 10], ["Globex", 20]];
    fake.selection.rowCount = 20;
    fake.selection.columnCount = 4;
    fake.selection.address = "Sheet1!A1:D20";
    fake.cellStore["A1:D20"] = { values: undefined, formulas: undefined };
    window.document.getElementById("includeSelection").checked = true;
    await fake.selectionHandlers[0]();
    await sleep(300);

    // selection can be excluded
    window.document.getElementById("includeSelection").checked = false;
    window.document.getElementById("input").value = "hello no context";
    window.document.getElementById("btnSend").click();
    await waitTurn();
    ok("prompt omits the selection when unchecked", !lastPrompt().includes("Selection Sheet1!"), lastPrompt().slice(0, 80));
    window.document.getElementById("includeSelection").checked = true;

    // no-workbook-context option: sheet reads and samples are skipped entirely
    window.document.getElementById("includeWorkbook").checked = false;
    window.document.getElementById("includeWorkbook").dispatchEvent(new window.Event("change"));
    usedRangeReads.count = 0;
    await fake.selectionHandlers[0]();
    await sleep(300);
    check("no-workbook-context skips sheet reads", usedRangeReads.count, 0);
    ok("context flags workbook samples off", /workbook samples off/.test(window.document.getElementById("contextInfo").textContent), window.document.getElementById("contextInfo").textContent);
    window.document.getElementById("input").value = "hello with no workbook context";
    window.document.getElementById("btnSend").click();
    await waitTurn();
    const noWorkbook = lastPrompt();
    ok("prompt omits workbook samples when disabled", !noWorkbook.includes("## Workbook context") && !noWorkbook.includes("Values from"), noWorkbook.slice(0, 120));
    ok("prompt keeps the selection when only workbook samples are off", noWorkbook.includes("Selection Sheet1!A1:D20"));

    window.document.getElementById("includeWorkbook").checked = true;
    window.document.getElementById("includeWorkbook").dispatchEvent(new window.Event("change"));
    await fake.selectionHandlers[0]();
    await sleep(300);
    ok("context restores workbook samples when re-enabled", !/workbook samples off/.test(window.document.getElementById("contextInfo").textContent), window.document.getElementById("contextInfo").textContent);

    // typed text + quick action are combined
    window.document.getElementById("input").value = "focus on column B";
    window.document.querySelector('.chip[data-action="vba"]').click();
    await waitTurn();
    const combined = lastPrompt();
    ok("quick action keeps its instruction", combined.includes("Write a VBA macro for this task"), combined.slice(0, 80));
    ok("quick action keeps the typed question", combined.includes("focus on column B"));

    /* --- compact command -------------------------------------------------- */
    console.log("\n-- compact command --");
    const tabsBeforeCompact = sessionTabs().length;
    window.document.getElementById("input").value = "/compact";
    window.document.getElementById("btnSend").click();
    const compactDeadline = Date.now() + 5000;
    while (Date.now() < compactDeadline && sessionTabs().length === tabsBeforeCompact) await sleep(100);
    await sleep(300);
    check("compact starts a new session tab", sessionTabs().length, tabsBeforeCompact + 1);
    ok("compact explains the carried summary", /compacted chat/i.test(window.document.getElementById("banner").textContent), window.document.getElementById("banner").textContent);
    window.document.getElementById("input").value = "continue with the summary";
    window.document.getElementById("btnSend").click();
    await waitTurn();
    const compactPrompt = lastPrompt();
    ok(
      "compact carries a bounded earlier-conversation summary",
      compactPrompt.includes("## Carried context from a compacted earlier conversation") && compactPrompt.includes("Earlier conversation:"),
      compactPrompt.slice(0, 400)
    );

    /* --- disconnected bridge ---------------------------------------------- */
    console.log("\n-- bridge goes away --");
    bridge.kill("SIGKILL");
    await sleep(500);
    window.document.getElementById("input").value = "hello again";
    window.document.getElementById("btnSend").click();
    await sleep(1200);
    const banner = window.document.getElementById("banner");
    ok("shows an error banner when the bridge is down", !banner.classList.contains("hidden"), banner.textContent);
    ok("banner mentions the bridge/OpenCode", /bridge|OpenCode|fetch|failed/i.test(banner.textContent), banner.textContent);
    const errorBubble = Array.from(window.document.querySelectorAll(".msg.assistant .bubble")).pop();
    ok("error is rendered in the transcript too", /Request failed/.test(errorBubble.textContent), errorBubble.textContent.slice(0, 60));
    ok("send button recovers", window.document.getElementById("btnSend").disabled === false);

    console.log("\nbridge log tail:\n" + bridgeLog.split("\n").slice(-6).join("\n"));
  } finally {
    try { mock.kill("SIGKILL"); } catch {}
    try { bridge.kill("SIGKILL"); } catch {}
  }

  console.log(failures ? `\n${failures} FAILURE(S)` : "\nAll pane tests passed");
  process.exit(failures ? 1 : 0);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
