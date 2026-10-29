/* Regression probes for the bugs documented in BUG-REPORT.md. */

import assert from "node:assert/strict";
import fs from "node:fs";

let passed = 0;
let failed = 0;
async function it(name, fn) {
  try {
    await fn();
    passed++;
    console.log("PASS  " + name);
  } catch (err) {
    failed++;
    console.log("FAIL  " + name + "\n      " + String(err.message).split("\n")[0]);
  }
}

const lib = await import("../src/lib.js").then(() => globalThis.OpenCodeExcel);
const plan = await import("../src/plan.js").then(() => globalThis.OpenCodeExcelPlan);
const fmt = await import("../src/format.js").then(() => globalThis.OpenCodeExcelFormat);
await import("../src/history.js");
const history = globalThis.OpenCodeExcelHistory;

function traceMock({ withApi, apiResult }) {
  const cells = { A1: { v: 1, f: "" }, B1: { v: 2, f: "=A1*2" } };
  function makeRange(address) {
    const a1 = String(address).split("!").pop().replace(/\$/g, "").split(":")[0];
    const cell = cells[a1] || { v: null, f: "" };
    const range = {
      address: "S1!" + a1,
      load() {},
      get values() { return [[cell.v]]; },
      get formulas() { return [[cell.f]]; },
      rowCount: 1,
      columnCount: 1,
      rowIndex: 0,
      columnIndex: 0,
    };
    if (withApi) {
      range.getDirectDependents = () => ({ load() {}, addresses: apiResult });
      range.getDirectPrecedents = () => ({ load() {}, addresses: [] });
    }
    return range;
  }
  const sheet = {
    name: "S1",
    load() {},
    getRange: makeRange,
    getUsedRangeOrNullObject() {
      return {
        isNullObject: false,
        load() {},
        rowCount: 1,
        columnCount: 2,
        rowIndex: 0,
        columnIndex: 0,
        formulas: [["", "=A1*2"]],
      };
    },
  };
  return {
    run: async (fn) => fn({
      workbook: { worksheets: { getItem: () => sheet, load() {}, items: [sheet] } },
      sync: async () => {},
    }),
  };
}

await it("trace fallback works without relation APIs", async () => {
  const out = await lib.traceRelations(traceMock({ withApi: false }), { sheet: "S1", range: "A1:A1" }, "dependents", true, { depth: 1 });
  assert.equal(out.status, "ok");
  assert.ok(!out.tree.error, "node error: " + out.tree.error);
  assert.equal(out.tree.relations[0].address, "S1!B1");
});

await it("empty relation API result falls back to formula scan", async () => {
  const out = await lib.traceRelations(traceMock({ withApi: true, apiResult: [] }), { sheet: "S1", range: "A1:A1" }, "dependents", true, { depth: 1 });
  assert.ok(!out.tree.error, "node error: " + out.tree.error);
  assert.equal(out.tree.relations[0].address, "S1!B1");
});

await it("extractTable skips markdown separator rows", () => {
  const table = lib.extractTable("| Month | Revenue |\n|-------|---------|\n| Jan | 100 |\n| Feb | 200 |");
  assert.deepEqual(table, [["Month", "Revenue"], ["Jan", "100"], ["Feb", "200"]]);
});

await it("extractTable skips aligned markdown separators", () => {
  assert.deepEqual(lib.extractTable("| A | B |\n|:--|--:|\n| 1 | 2 |"), [["A", "B"], ["1", "2"]]);
});

function structuralMockExcel() {
  const ops = [];
  const sheet = {
    getRange(ref) {
      return {
        insert(shift) { ops.push({ kind: "insert", ref, shift }); },
        delete(dir) { ops.push({ kind: "delete", ref, dir }); },
      };
    },
  };
  const Excel = {
    InsertShiftDirection: { up: "Up", down: "Down", left: "Left", right: "Right" },
    DeleteShiftDirection: { up: "Up", left: "Left" },
    run: async (fn) => fn({
      workbook: { worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet } },
      sync: async () => {},
    }),
  };
  return { Excel, ops };
}

await it("row and column inserts use InsertShiftDirection", async () => {
  const { Excel, ops } = structuralMockExcel();
  globalThis.Excel = Excel;
  try {
    const results = await plan.applyOps(Excel, [
      plan.normalizeOp({ op: "insertRows", sheet: "S1", at: 3, count: 2 }),
      plan.normalizeOp({ op: "insertColumns", sheet: "S1", at: "C" }),
    ]);
    assert.equal(results[0].ok, true, results[0].error);
    assert.equal(results[1].ok, true, results[1].error);
    assert.deepEqual(ops.map((o) => o.shift), ["Down", "Right"]);
  } finally {
    delete globalThis.Excel;
  }
});

await it("charts resolve camelCase Office.js enum members", async () => {
  const added = [];
  const chart = { set name(_value) {}, setPosition() {} };
  const sheet = {
    getRange: (ref) => ({ ref }),
    charts: { add(type, source, seriesBy) { added.push({ type, source, seriesBy }); return chart; } },
  };
  const Excel = {
    ChartType: { columnClustered: "ColumnClustered", line: "Line", pie: "Pie" },
    ChartSeriesBy: { auto: "Auto", columns: "Columns", rows: "Rows" },
    run: async (fn) => fn({
      workbook: { worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet } },
      sync: async () => {},
    }),
  };
  globalThis.Excel = Excel;
  try {
    const results = await plan.applyOps(Excel, [plan.normalizeOp({ op: "addChart", sheet: "S1", source: "A1:D10", chartType: "ColumnClustered" })]);
    assert.equal(results[0].ok, true, results[0].error);
    assert.equal(added[0].type, "ColumnClustered");
    assert.equal(added[0].seriesBy, "Columns");
  } finally {
    delete globalThis.Excel;
  }
});

await it("unknown ISO currency codes do not become USD", () => {
  const out = fmt.parseConventions("currency: INR, decimals: 2");
  assert.match(out.currency, /INR/);
  assert.ok(!out.currency.includes("$"), out.currency);
});

await it("known currency codes and defaults remain intact", () => {
  assert.match(fmt.parseConventions("currency: EUR").currency, /€/);
  assert.equal(fmt.parseConventions("").currency, "$#,##0;[Red]-$#,##0");
});

await it("history.get returns null for an IndexedDB miss", async () => {
  const fakeRequest = { result: undefined, onsuccess: null };
  const fakeTx = {
    objectStore: () => ({
      get: () => fakeRequest,
      index: () => ({ getAll: () => ({ result: [] }), openCursor: () => ({ onsuccess: null, result: null }) }),
      put: () => fakeRequest,
      delete: () => fakeRequest,
    }),
    oncomplete: null,
    onerror: null,
    onabort: null,
  };
  const fakeDb = {
    transaction: () => { setTimeout(() => fakeTx.oncomplete && fakeTx.oncomplete(), 0); return fakeTx; },
    objectStoreNames: { contains: () => true },
    createObjectStore: () => fakeTx.objectStore(),
  };
  globalThis.indexedDB = {
    open: () => {
      const request = { onupgradeneeded: null, onsuccess: null, onerror: null, onblocked: null, result: fakeDb };
      setTimeout(() => request.onsuccess && request.onsuccess(), 0);
      return request;
    },
  };
  try {
    assert.equal(await history.get("does-not-exist"), null);
  } finally {
    delete globalThis.indexedDB;
  }
});

await it("address helpers accept absolute and whole-row references", () => {
  assert.equal(lib.belowAddress("Sheet1!$A$1:$D$20"), "A22");
  assert.equal(lib.belowAddress("Sheet1!3:3"), "A5");
  assert.equal(lib.startAddress("Sheet1!$B$3:$D$9"), "B3");
  assert.equal(lib.belowAddress("Sheet1!A1:D20"), "A22");
});

const pane = fs.readFileSync(new URL("../src/taskpane.js", import.meta.url), "utf8");

await it("text-only DOM sinks do not double-escape strings", () => {
  assert.ok(!/banner\("File too large: " \+ escapeHtml/.test(pane));
  assert.ok(!/chip\.textContent = escapeHtml\(a\.name\)/.test(pane));
  assert.ok(!/"Remove " \+ escapeHtml/.test(pane));
  assert.ok(!/return when \+ "  " \+ e\.type \+ "  " \+ escapeHtml/.test(pane));
});

await it("status.version is escaped before innerHTML", () => {
  const setStatus = pane.slice(pane.indexOf("function setStatus"), pane.indexOf("async function ensureSession"));
  assert.ok(/escapeHtml\((?:String\()?status\.version/.test(setStatus) || /textContent = "OpenCode "/.test(setStatus));
});

await it("pane guard outlives the bridge timeout", () => {
  const match = pane.match(/setTimeout\(function \(\) \{\s*if \(state\.streaming\)[\s\S]*?\}, ([\d \*]+)\);/);
  assert.ok(match, "guard timer not found");
  const milliseconds = Function('"use strict";return (' + match[1] + ")")();
  assert.ok(milliseconds >= 10 * 60 * 1000, String(milliseconds));
});

await it("handleEvent renders late done/text events", () => {
  const handle = pane.slice(pane.indexOf("function handleEvent"), pane.indexOf("function errorMessage"));
  assert.match(handle, /if \(!state\.bubble\) \{[\s\S]{0,400}"done"[\s\S]{0,400}renderMarkdown/);
});

await it("stale approved reads release the streaming lock", () => {
  const block = pane.slice(pane.indexOf("var results = await readRequestedRanges(request.tokens)"));
  const earlyReturn = block.slice(0, block.indexOf("state.readRounds++"));
  assert.match(earlyReturn, /state\.streaming = false/);
});

const server = fs.readFileSync(new URL("../bridge/server.mjs", import.meta.url), "utf8");

await it("upstream SSE is aborted when its client disconnects", () => {
  const stopFunctions = server.match(/stop\(\) \{[\s\S]{0,240}?\}/g) || [];
  assert.ok(stopFunctions.some((source) => source.includes("abort")));
});

console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
