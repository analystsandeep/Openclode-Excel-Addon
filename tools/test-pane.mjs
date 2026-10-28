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
      return makeHandle("Sheet1!A1:D20", { rowCount: 20, columnCount: 4, address: "Sheet1!A1:D20", values: selection.values });
    },
    tables: { items: [], load() {} },
    load() {},
    set name(v) { ops.push({ kind: "renameSheet", from: "Sheet1", to: v }); this._name = v; },
    get name() { return this._name || "Sheet1"; },
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
    load() {},
  };
  // Before-state the formatter/snapshot reads back for the undo test.
  cellStore["A1:B1"] = { values: [["Name", "Amount"]] };
  cellStore["A1:B2"] = { values: [["Name", "Amount"], ["ACME", 10]] };

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

    const fake = createFakeExcel();
    window.Excel = fake.Excel;
    window.EventSource = makeEventSource(BASE);
    const requests = [];
    window.fetch = (url, options) => {
      requests.push({ url: String(url), body: options && options.body });
      return fetch(url.startsWith("http") ? url : BASE + url, options);
    };
    const sendButton = () => window.document.getElementById("btnSend");
    // A send is async: wait for it to start, then for it to finish.
    const waitTurn = async () => {
      await sleep(400);
      const deadline = Date.now() + 25000;
      while (Date.now() < deadline && sendButton().disabled) await sleep(200);
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
    window.OpenCodeExcelFormat = {
      formatSheet: async () => {
        formatCalls++;
        return { sheetName: "Sheet1", tableName: "tblSales", dataRows: 19, created: true, warnings: [] };
      },
    };

    window.eval(fs.readFileSync(path.join(ROOT, "src", "lib.js"), "utf8"));
    window.eval(fs.readFileSync(path.join(ROOT, "src", "plan.js"), "utf8"));
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
    ok("models loaded into dropdown", window.document.getElementById("modelSelect").options.length >= 3, `options: ${window.document.getElementById("modelSelect").options.length}`);
    ok("selection-change handler registered", fake.selectionHandlers.length === 1, `handlers: ${fake.selectionHandlers.length}`);

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
    ok("undo row is shown", !window.document.getElementById("undoRow").classList.contains("hidden"));

    window.document.getElementById("btnUndo").click();
    await sleep(500);
    const restored = fake.ops.filter((o) => o.kind === "values" && o.a1 === "A1:B2").pop();
    check("undo restores previous values", restored && JSON.stringify(restored.values), JSON.stringify([["Name", "Amount"], ["ACME", 10]]));
    ok("undo remains available for earlier direct inserts", !window.document.getElementById("undoRow").classList.contains("hidden"));

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
    check("stop leaves a finalised bubble (8 turns)", assistantCount, 8);
    ok("stopped bubble is not left thinking", !lastAssistant.innerHTML.includes("Thinking"), lastAssistant.innerHTML.slice(0, 60));

    window.document.getElementById("btnNewChat").click();
    await sleep(800);
    ok("new chat clears the transcript", window.document.querySelectorAll(".msg").length === 0);
    ok("new chat restores empty state", !!window.document.getElementById("emptyState"));

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
    ok("context marks a big selection as truncated", /truncated/.test(window.document.getElementById("contextInfo").textContent), window.document.getElementById("contextInfo").textContent);

    window.document.getElementById("input").value = "what is in here?";
    window.document.getElementById("btnSend").click();
    await waitTurn();
    const capped = lastPrompt();
    const tsvLines = capped.split("\n").filter((l) => l.includes("\t")).length;
    ok("prompt says truncated", capped.includes("truncated"), capped.slice(0, 80));
    ok("prompt caps rows (<= 20 rows for a 20-column range)", tsvLines <= 20, `tsv lines: ${tsvLines}`);

    // selection can be excluded
    window.document.getElementById("includeSelection").checked = false;
    window.document.getElementById("input").value = "hello no context";
    window.document.getElementById("btnSend").click();
    await waitTurn();
    ok("prompt omits the selection when unchecked", !lastPrompt().includes("Selection Sheet1!"), lastPrompt().slice(0, 80));
    window.document.getElementById("includeSelection").checked = true;

    // typed text + quick action are combined
    window.document.getElementById("input").value = "focus on column B";
    window.document.querySelector('.chip[data-action="vba"]').click();
    await waitTurn();
    const combined = lastPrompt();
    ok("quick action keeps its instruction", combined.includes("Write a VBA macro for this task"), combined.slice(0, 80));
    ok("quick action keeps the typed question", combined.includes("focus on column B"));

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
