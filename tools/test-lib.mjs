/* Unit tests for src/lib.js - run: node tools/test-lib.mjs */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const src = fs.readFileSync(path.join(root, "src", "lib.js"), "utf8");
const planSrc = fs.readFileSync(path.join(root, "src", "plan.js"), "utf8");

// Load each classic script into its own sandbox so window.OpenCodeExcel /
// window.OpenCodeExcelPlan get defined.
function load(src) {
  const sandbox = { module: undefined };
  new Function("globalThis", "window", "module", src)(sandbox, sandbox, undefined);
  return sandbox;
}
const lib = load(src).OpenCodeExcel;
const plan = load(planSrc).OpenCodeExcelPlan;

let failures = 0;
function check(name, actual, expected) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  console.log(`${ok ? "PASS" : "FAIL"}  ${name}`);
  if (!ok) {
    failures++;
    console.log(`      expected: ${JSON.stringify(expected)}`);
    console.log(`      actual:   ${JSON.stringify(actual)}`);
  }
}

// The exact payload the mock OpenCode service streams back.
const answer = [
  "Here is a formula that flags duplicates in column B:\n\n",
  "```excel\n=IF(COUNTIF($B$2:$B$100,B2)>1,\"duplicate\",\"\")\n```\n\n",
  "And the top offenders as a table:\n\n```tsv\nValue\tCount\nACME\t4\nGlobex\t3\nInitech\t2\n```\n",
].join("");

check("extractTable finds the fenced tsv block", lib.extractTable(answer), [
  ["Value", "Count"],
  ["ACME", "4"],
  ["Globex", "3"],
  ["Initech", "2"],
]);
check("singleFormula picks the lone = line", lib.singleFormula(answer), '=IF(COUNTIF($B$2:$B$100,B2)>1,"duplicate","")');
check("singleFormula ignores prose", lib.singleFormula("Use =A1+B1 and then =A2+B2"), null);
check("belowAddress (A1:D20 -> A22)", lib.belowAddress("Sheet1!A1:D20"), "A22");
check("belowAddress (single cell B5 -> B7)", lib.belowAddress("Sheet1!B5"), "B7");
check("startAddress (Sheet1!C7:F9 -> C7)", lib.startAddress("Sheet1!C7:F9"), "C7");
check("colName 0/25/26/701", [lib.colName(0), lib.colName(25), lib.colName(26), lib.colName(701)], ["A", "Z", "AA", "ZZ"]);
check("toA1(0,0) / toA1(9,2)", [lib.toA1(0, 0), lib.toA1(9, 2)], ["A1", "C10"]);

const html = lib.renderMarkdown(answer);
check("renderMarkdown escapes + fences code", html.includes("<pre>") && html.includes("&quot;duplicate&quot;"), true);
check("renderMarkdown has no raw html injection", lib.renderMarkdown("<script>x</script>").includes("&lt;script&gt;"), true);
check("renderMarkdown renders bold", lib.renderMarkdown("**hi**").includes("<strong>hi</strong>"), true);
check("renderMarkdown renders bullets", lib.renderMarkdown("- one\n- two").includes("<ul><li>one</li><li>two</li></ul>"), true);

// --- edge cases -------------------------------------------------------------
check("extractTable: no table in prose", lib.extractTable("Just a sentence, nothing tabular."), null);
check("extractTable: header only is not a table", lib.extractTable("\`\`\`tsv\nA\tB\n\`\`\`"), null);
check("extractTable: pipe table", lib.extractTable("| A | B |\n| 1 | 2 |\n| 3 | 4 |"), [["A", "B"], ["1", "2"], ["3", "4"]]);
check("extractTable: ragged rows are dropped", lib.extractTable("\`\`\`tsv\nA\tB\n1\t2\t3\n4\t5\n\`\`\`"), [["A", "B"], ["4", "5"]]);
check("extractTable: tab table without a fence", lib.extractTable("A\tB\n1\t2\n3\t4"), [["A", "B"], ["1", "2"], ["3", "4"]]);
check("singleFormula: null when several formulas", lib.singleFormula("=A1\n=B1"), null);
check("singleFormula: ignores overly long lines", lib.singleFormula("=" + "A".repeat(400)), null);
check("belowAddress: invalid input", lib.belowAddress("nonsense"), null);
check("colName(702)", lib.colName(702), "AAA");
check("renderMarkdown: CRLF input", lib.renderMarkdown("line1\r\nline2").includes("<p>line1</p>"), true);
check("renderMarkdown: escapes html inside code", lib.renderMarkdown("\`\`\`\n<script>alert(1)</script>\n\`\`\`").includes("&lt;script&gt;"), true);
check("renderMarkdown: code attribute cannot break out", /data-code="[^"]*&quot;/.test(lib.renderMarkdown("\`\`\`\nsay \"hi\"\n\`\`\`")), true);
check("renderMarkdown: empty input", lib.renderMarkdown(""), "");
check("escapeHtml: all three entities", lib.escapeHtml("<a href=\"x\">&</a>"), "&lt;a href=\"x\"&gt;&amp;&lt;/a&gt;");

// --- structured edit plans --------------------------------------------------

const planAnswer = "Here is the plan:\n\n```json\n{\n  \"summary\": \"Add a sales table\",\n  \"assumptions\": [\"Qty x Price\"],\n  \"operations\": [\n    { \"op\": \"addSheet\", \"name\": \"Sales\", \"description\": \"Add a Sales tab\" },\n    { \"op\": \"setValues\", \"sheet\": \"Sales\", \"range\": \"A1:C1\", \"values\": [[\"Month\",\"Units\",\"Amount\"]], \"description\": \"Write headers\" },\n    { \"op\": \"setFormulas\", \"sheet\": \"Sales\", \"range\": \"D2\", \"formula\": \"=B2*C2\", \"description\": \"Total\" }\n  ]\n}\n```\n";

check("parsePlan finds the fenced json plan", plan.parsePlan(planAnswer).ok, true);
const parsed = plan.parsePlan(planAnswer);
check("parsePlan summary", parsed.summary, "Add a sales table");
check("parsePlan operation count", parsed.operations.length, 3);
check("parsePlan cells affected", parsed.cellsAffected, ["Sales!A1:C1", "Sales!D2"]);
check("parsePlan destructive flag", parsed.destructive, true);
check("parsePlan flags setValues as destructive", parsed.operations[1].destructive, true);
check("parsePlan flags addSheet as safe", parsed.operations[0].destructive, false);
check("parsePlan keeps descriptions", parsed.operations[2].description, "Total");

check("parsePlan rejects plain text", plan.parsePlan("Just a sentence.").ok, false);
check("parsePlan rejects bad json", plan.parsePlan("```json\n{ not valid\n```\n").ok, false);
check("parsePlan rejects missing operations", plan.parsePlan("```json\n{ \"summary\": \"x\" }\n```\n").ok, false);

const lines = plan.planLines(parsed.operations);
check("planLines has labels and cells", lines[0].label, "Add a Sales tab");
check("planLines op names", lines.map(function (l) { return l.op; }), ["addSheet", "setValues", "setFormulas"]);
const rowOp = plan.normalizeOp({ op: "insertRows", sheet: "Plan", at: 5, count: 3 });
check("row operation expands at/count", rowOp.range, "5:7");
const colOp = plan.normalizeOp({ op: "deleteColumns", sheet: "Plan", at: "C", count: 2 });
check("column operation expands at/count", colOp.range, "C:D");
const r1c1 = plan.normalizeOp({ op: "setFormulas", sheet: "Plan", range: "E2:E10", formulaR1C1: "=RC[-2]*RC[-1]" });
check("multi-cell R1C1 plan is valid", r1c1.ok, true);
check("plan estimates affected cells", r1c1.cellCount, 9);
check("rangeCellCount handles rectangles", plan.rangeCellCount("Sheet1!B2:D5"), 12);
check("range operation requires an exact sheet", plan.normalizeOp({ op: "clearRange", range: "A1" }).ok, false);
check("expanded operation set is available", ["clearRange", "addTable", "sortRange", "freezePanes"].every((op) => plan.ALLOWED_OPS.includes(op)), true);

// --- safety gates: classification, caps, auto-approval, undo ---------------
const mergeOp = plan.normalizeOp({ op: "mergeRange", sheet: "Data", range: "A1:B2" });
check("merge is destructive", mergeOp.destructive, true);
const sortOp = plan.normalizeOp({ op: "sortRange", sheet: "Data", range: "A1:B2", sort: [] });
check("sort is destructive", sortOp.destructive, true);
check("deleteSheet has an unbounded cell count", plan.normalizeOp({ op: "deleteSheet", name: "Data" }).cellCount, Infinity);
check("deleteTable has an unbounded cell count", plan.normalizeOp({ op: "deleteTable", sheet: "Data", name: "tbl" }).cellCount, Infinity);
check("deleteChart has an unbounded cell count", plan.normalizeOp({ op: "deleteChart", sheet: "Data", name: "Chart 1" }).cellCount, Infinity);

const safeFormat = plan.normalizeOp({ op: "formatRange", sheet: "Data", range: "A1:B2", format: { bold: true } });
const smallValues = plan.normalizeOp({ op: "setValues", sheet: "Data", range: "A1", values: [[1]] });
const namedFormat = plan.normalizeOp({ op: "formatRange", sheet: "Data", range: "MyRange", format: { bold: true } });
check("small safe formatting can auto-apply", plan.canAutoApply([safeFormat], 5), true);
check("formatting above the auto limit is not auto-applied", plan.canAutoApply([safeFormat], 3), false);
check("cell writes always require review", plan.canAutoApply([smallValues], 5000), false);
check("sheet creation is never auto-applied", plan.canAutoApply([plan.normalizeOp({ op: "addSheet", name: "New" })], 5000), false);
check("sheet deletion is never auto-applied", plan.canAutoApply([plan.normalizeOp({ op: "deleteSheet", name: "Data" })], 5000), false);
check("table deletion is never auto-applied", plan.canAutoApply([plan.normalizeOp({ op: "deleteTable", sheet: "Data", name: "tbl" })], 5000), false);
check("chart deletion is never auto-applied", plan.canAutoApply([plan.normalizeOp({ op: "deleteChart", sheet: "Data", name: "Chart 1" })], 5000), false);
check("named ranges have unbounded cell counts", namedFormat.cellCount, Infinity);
check("named ranges are not auto-applied", plan.canAutoApply([namedFormat], 5000), false);

const oversizedPlanText = JSON.stringify({
  summary: "big",
  operations: [{ op: "setValues", sheet: "Data", range: "A1:Z100000", values: [[1]] }],
});
const oversizedPlan = plan.parsePlan(oversizedPlanText);
check("oversized ranges are rejected during parsing", oversizedPlan.ok, false);
check("oversized range error names the cell limit", /5000/.test(oversizedPlan.error || ""), true);

check("cell writes are add-in undo-safe", plan.isUndoSafe(smallValues), true);
check("sheet creation is not add-in undo-safe", plan.isUndoSafe(plan.normalizeOp({ op: "addSheet", name: "New" })), false);
check("clearing is not add-in undo-safe", plan.isUndoSafe(plan.normalizeOp({ op: "clearRange", sheet: "Data", range: "A1" })), false);
check("row insertion is not add-in undo-safe", plan.isUndoSafe(plan.normalizeOp({ op: "insertRows", sheet: "Data", at: 2, count: 1 })), false);

check("whole-column ranges are unbounded", plan.rangeCellCount("Data!A:C"), Infinity);
check("whole-row ranges are unbounded", plan.rangeCellCount("Data!1:3"), Infinity);
check("named ranges are unbounded", plan.rangeCellCount("Data!MyRange"), Infinity);

function makeSheet(name) {
  return {
    name: name,
    deleted: false,
    getRange: function (address) {
      return {
        values: [[name + "!" + address]],
        formulas: [[name + "!" + address]],
        numberFormat: [["General"]],
        format: {
          horizontalAlignment: "General", verticalAlignment: "Bottom",
          wrapText: false, rowHeight: 15, columnWidth: 8.43,
          font: { bold: false, italic: false, size: 11, color: "#000000" },
          fill: { color: "#FFFFFF" },
          load: function () {},
        },
        load: function () {},
      };
    },
    delete: function () { this.deleted = true; },
  };
}

function makeExcel(sheets) {
  return {
    run: async function (fn) {
      return fn({
        workbook: {
          worksheets: {
            getActiveWorksheet: function () { return sheets[0]; },
            getItem: function (name) {
              var found = sheets.filter(function (s) { return s.name === name; })[0];
              if (!found) throw new Error("sheet not found");
              return found;
            },
          },
        },
        sync: async function () {},
      });
    },
  };
}

const existingSheet = makeSheet("Existing");
const dataSheet = makeSheet("Data");
const snapshotExcel = makeExcel([dataSheet, existingSheet]);
const failedCreateOps = [
  plan.normalizeOp({ op: "addSheet", name: "Existing" }),
  plan.normalizeOp({ op: "setValues", sheet: "Data", range: "A1", values: [[42]] }),
];
const snapshot = await plan.snapshotCells(snapshotExcel, failedCreateOps);
check("snapshot excludes sheet creation", Object.keys(snapshot.sheets), ["Data"]);
check("snapshot has no created-sheet list", snapshot.createdSheets, undefined);
await plan.undoSnapshot(snapshotExcel, snapshot);
check("undo never deletes a pre-existing sheet", existingSheet.deleted, false);

const liedOp = {
  op: "setValues", sheet: "Data", range: "A1:F1000", address: "Data!A1:F1000",
  cellCount: 1, destructive: true, ok: true,
};
let capError = null;
try { await plan.snapshotCells(snapshotExcel, [liedOp]); } catch (err) { capError = err; }
check("snapshot re-checks the real range size", !!capError && /safety limit/.test(capError.message), true);

var formulaAssignments = [];
var formulaState = null;
var formulaHandle = {
  getCell: function () {
    return {
      set formulas(value) { formulaAssignments.push(["seed", value]); },
    };
  },
  getRow: function () {
    return {
      fillRight: function () {
        formulaAssignments.push(["fillRight"]);
        formulaState = [["=A2", "=B2", "=C2"]];
      },
    };
  },
  get formulas() { return formulaState; },
  set formulas(value) { formulaAssignments.push(["range", value]); formulaState = value; },
};
var formulaSheet = { getRange: function () { return formulaHandle; } };
var formulaExcel = {
  run: async function (fn) {
    return fn({
      workbook: {
        worksheets: {
          getActiveWorksheet: function () { return formulaSheet; },
          getItem: function () { return formulaSheet; },
        },
      },
      sync: async function () {},
    });
  },
};
const formulaOp = plan.normalizeOp({ op: "setFormulas", sheet: "Sheet1", range: "A1:C1", formula: "=A2" });
const formulaResults = await plan.applyOps(formulaExcel, [formulaOp]);
check("horizontal formula fill succeeds", formulaResults[0].ok, true);
check("fillRight translated formula is retained", formulaState, [["=A2", "=B2", "=C2"]]);
check(
  "fillRight is not overwritten by a matrix",
  formulaAssignments.filter(function (item) { return item[0] === "range"; }).length,
  0
);

// --- sheet formatter helpers ------------------------------------------------
check(
  "cleanHeaders fixes the SaleData headers",
  lib.cleanHeaders(["OrderDate", "Region", "Manager", "SalesMan", "Item", "Units", " Unit_price ", "Sale_amt"]),
  ["Order Date", "Region", "Manager", "Sales Rep", "Item", "Units", "Unit Price", "Sale Amount"]
);
check("cleanHeaders makes duplicate/blank names unique", lib.cleanHeaders(["Name", "name", ""]), ["Name", "Name 2", "Column 3"]);
check(
  "detectColumns finds date/units/amount/region/people/item",
  lib.detectColumns(["Order Date", "Region", "Manager", "Sales Rep", "Item", "Units", "Unit Price", "Sale Amount"]),
  { date: 0, units: 5, amount: 7, region: 1, people: 3, item: 4 }
);
check("structuredRef quotes table names with spaces", lib.structuredRef("My Table", "Sale Amount"), "'My Table'[Sale Amount]");
check("structuredRef keeps simple table names bare", lib.structuredRef("tblSales", "Units"), "tblSales[Units]");
check("uniqueTableName skips taken names", lib.uniqueTableName(["tblSales", "tblSales2"], "tblSales"), "tblSales3");

console.log(failures ? `\n${failures} test(s) failed` : "\nAll tests passed");
process.exit(failures ? 1 : 0);
