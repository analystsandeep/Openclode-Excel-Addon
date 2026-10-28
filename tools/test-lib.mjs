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

// The model prompt uses a cell-shaped range with count. It must still affect
// whole rows/columns, otherwise only one column/row shifts.
const promptRowOp = plan.normalizeOp({ op: "insertRows", sheet: "Sheet1", range: "A5", count: 3 });
check("prompt row example is valid", promptRowOp.ok, true);
check("prompt row example becomes whole rows", promptRowOp.range, "5:7");
check("prompt row example keeps count", promptRowOp.count, 3);
check("prompt row example names whole rows in the preview", promptRowOp.address, "Sheet1!5:7");
const promptColOp = plan.normalizeOp({ op: "insertColumns", sheet: "Sheet1", range: "A5", count: 3 });
check("cell-shaped column edit becomes whole columns", promptColOp.range, "A:C");
check("cell-shaped column edit names whole columns", promptColOp.address, "Sheet1!A:C");
const spannedRows = plan.normalizeOp({ op: "deleteRows", sheet: "Sheet1", range: "A5:C7" });
check("multi-row range becomes whole rows", spannedRows.range, "5:7");
check("multi-row range keeps its span", spannedRows.count, 3);
const spannedCols = plan.normalizeOp({ op: "deleteColumns", sheet: "Sheet1", range: "B2:D9" });
check("multi-column range becomes whole columns", spannedCols.range, "B:D");
const wholeRows = plan.normalizeOp({ op: "insertRows", sheet: "Sheet1", range: "5:7", count: 3 });
check("whole-row range with matching count is accepted", wholeRows.range, "5:7");
const conflictingRows = plan.normalizeOp({ op: "insertRows", sheet: "Sheet1", range: "A5:C7", count: 2 });
check("conflicting row span and count is rejected", conflictingRows.ok, false);
check("conflicting row error explains the mismatch", /spans 3 row/.test(conflictingRows.error || ""), true);
const conflictingCols = plan.normalizeOp({ op: "deleteColumns", sheet: "Sheet1", range: "C3:E9", count: 2 });
check("conflicting column span and count is rejected", conflictingCols.ok, false);
const bothForms = plan.normalizeOp({ op: "insertRows", sheet: "Sheet1", at: 5, range: "A5", count: 3 });
check("at plus range is rejected", bothForms.ok, false);
const namedRows = plan.normalizeOp({ op: "insertRows", sheet: "Sheet1", range: "MyRows" });
check("named row range is rejected", namedRows.ok, false);
const zeroCount = plan.normalizeOp({ op: "insertRows", sheet: "Sheet1", at: 5, count: 0 });
check("zero count is rejected", zeroCount.ok, false);
check("zero count error explains the requirement", /positive integer/.test(zeroCount.error || ""), true);
const r1c1 = plan.normalizeOp({ op: "setFormulas", sheet: "Plan", range: "E2:E10", formulaR1C1: "=RC[-2]*RC[-1]" });
check("multi-cell R1C1 plan is valid", r1c1.ok, true);
check("plan estimates affected cells", r1c1.cellCount, 9);
check("rangeCellCount handles rectangles", plan.rangeCellCount("Sheet1!B2:D5"), 12);
check("range operation requires an exact sheet", plan.normalizeOp({ op: "clearRange", range: "A1" }).ok, false);
check("expanded operation set is available", ["clearRange", "addTable", "sortRange", "freezePanes"].every((op) => plan.ALLOWED_OPS.includes(op)), true);

// Whole-row/column safety: never shift just the anchored cell.
check("A5 plus count normalizes to whole rows", plan.normalizeOp({op:"insertRows",sheet:"S",range:"A5",count:3}).range, "5:7");
check("B2 plus count normalizes to whole columns", plan.normalizeOp({op:"deleteColumns",sheet:"S",range:"B2",count:2}).range, "B:C");
check("single-row cell range honors count", plan.normalizeOp({op:"deleteRows",sheet:"S",range:"A5:C5",count:3}).range, "5:7");
check("conflicting row count rejected", plan.normalizeOp({op:"deleteRows",sheet:"S",range:"5:7",count:2}).ok, false);
check("negative column rejected", plan.normalizeOp({op:"insertColumns",sheet:"S",at:-1}).ok, false);
check("fractional count rejected", plan.normalizeOp({op:"insertRows",sheet:"S",at:5,count:1.5}).ok, false);
check("out of bounds row rejected", plan.normalizeOp({op:"insertRows",sheet:"S",at:1048576,count:2}).ok, false);
check("whole range is destructive and explicit", plan.normalizeOp({op:"insertRows",sheet:"S",range:"A5",count:3,description:"harmless"}).cells, ["S!5:7"]);
let requested = lib.parseRequestedRanges("Analyze @range('Sales Data'!A1:B2)");
check("explicit range accepts quoted sheet", requested, [{sheet:"Sales Data",range:"A1:B2",cellCount:4}]);
function rangeRejected(text) { try { lib.parseRequestedRanges(text); return false; } catch { return true; } }
check("explicit range rejects whole columns", rangeRejected("@range(S!A:A)"), true);
check("explicit range rejects >500 cells", rangeRejected("@range(S!A1:A501)"), true);
check("explicit range rejects worksheet overflow", rangeRejected("@range(S!XFE1)"), true);
check("explicit range rejects 4 requests", rangeRejected("@range(S!A1) @range(S!A2) @range(S!A3) @range(S!A4)"), true);
const structuralCalls = [];
const structuralExcel = { InsertShift: {down:"Down",right:"Right"}, DeleteShiftDirection: {up:"Up",left:"Left"}, run: async function(fn) {
  return fn({workbook:{worksheets:{getItem:()=>({getRange: address => ({insert: direction => structuralCalls.push([address,direction]),delete:direction=>structuralCalls.push([address,direction])})})}},sync:async()=>{}});
}};
await plan.applyOps(structuralExcel,[plan.normalizeOp({op:"insertRows",sheet:"S",range:"A5",count:3}),plan.normalizeOp({op:"deleteColumns",sheet:"S",range:"B2",count:2})]);
check("structural execution uses whole ranges",structuralCalls,[["5:7","Down"],["B:C","Left"]]);

// Chart API argument order must match Excel.ChartCollection.add(type, sourceData, seriesBy).
const chartCalls = [];
const oldGlobalExcel = globalThis.Excel;
globalThis.Excel = {ChartType:{ColumnClustered:"ColumnClustered"}, ChartSeriesBy:{columns:"Columns"}};
// plan is loaded in its own sandbox; expose the Excel enum there via a fresh load.
const chartSandbox = {Excel: globalThis.Excel};
new Function("globalThis", "window", "module", "Excel", planSrc)(chartSandbox, chartSandbox, undefined, globalThis.Excel);
const chartPlan = chartSandbox.OpenCodeExcelPlan;
const sourceRange = { address: "A1:B2" };
const chartExcel = {run: async fn => fn({workbook:{worksheets:{getItem:()=>({getRange:()=>sourceRange,charts:{add:(...args)=>{chartCalls.push(args);return {};}}})}},sync:async()=>{}})};
const chartResult = await chartPlan.applyOps(chartExcel, [chartPlan.normalizeOp({op:"addChart",sheet:"S",source:"A1:B2"})]);
check("chart uses source data as second argument", chartCalls[0] && chartCalls[0][1] === sourceRange, true);
check("chart uses series direction as third argument", chartCalls[0] && chartCalls[0][2], "Columns");
check("chart call succeeds in mock", chartResult[0].ok, true);
globalThis.Excel = oldGlobalExcel;
let activeFallback = false;
const missingSheetExcel = {run: async fn => fn({workbook:{worksheets:{getItem:()=>{throw Error("sheet missing")},getActiveWorksheet:()=>{activeFallback=true;return {getRange:()=>({})}}}},sync:async()=>{}})};
const missingResult = await plan.applyOps(missingSheetExcel,[plan.normalizeOp({op:"setValues",sheet:"Missing",range:"A1",values:[[1]]})]);
check("missing sheet fails safely", missingResult[0].ok, false);
check("missing sheet never falls back to active worksheet", activeFallback, false);

// Agent range requests require strict finite addresses and explicit consent in the UI.
const modelRead = lib.parseAgentReadRequest(JSON.stringify({read_ranges:[{sheet:"Sales",range:"A1:B2"}],reason:"Need data"}));
check("model read request is bounded and normalized",modelRead.ranges,[{sheet:"Sales",range:"A1:B2",cellCount:4}]);
function agentRequestRejected(payload) { try { lib.parseAgentReadRequest(JSON.stringify(payload)); return false; } catch { return true; } }
check("model read rejects whole column",agentRequestRejected({read_ranges:[{sheet:"S",range:"A:A"}]}),true);
check("model read rejects embedded qualified address",agentRequestRejected({read_ranges:[{sheet:"S",range:"T!A1"}]}),true);
check("model read rejects mixed write operations",agentRequestRejected({read_ranges:[{sheet:"S",range:"A1"}],operations:[]}),true);
check("trace parser accepts exact single cell",lib.parseRequestedTraces("@trace('Sales Data'!C5)"),[{sheet:"Sales Data",range:"C5:C5",cellCount:1}]);
check("trace parser rejects range", (()=>{try{lib.parseRequestedTraces("@trace(S!A1:B2)");return false;}catch{return true;}})(),true);
const traceRequest = {sheet:"S",range:"A1:A1"};
let traceMethod = null;
const traceExcel = {run: async fn => fn({workbook:{worksheets:{getItem:()=>({getRange:()=>({load(){},getDirectPrecedents(){traceMethod="precedents";return {load(){},addresses:["S!B1"]}},getDirectDependents(){traceMethod="dependents";return {load(){},addresses:["S!C1"]}}})})}},sync:async()=>{}})};
check("direct precedents use Excel relation API",await lib.traceDirectRelations(traceExcel,traceRequest,"precedents",true),{status:"ok",addresses:["S!B1"],truncated:false});
check("direct dependents use Excel relation API",await lib.traceDirectRelations(traceExcel,traceRequest,"dependents",true),{status:"ok",addresses:["S!C1"],truncated:false});
check("unsupported host is reported honestly",await lib.traceDirectRelations(traceExcel,traceRequest,"dependents",false),{status:"unsupported",addresses:[]});

const noRelationsExcel = {run:async fn=>{let sync=0;return fn({workbook:{worksheets:{getItem:()=>({getRange:()=>({load(){},getDirectPrecedents(){return {load(){}}}})})}},sync:async()=>{if(++sync===2)throw {code:"ItemNotFound"}}})}};
check("missing direct relation reported as none",await lib.traceDirectRelations(noRelationsExcel,traceRequest,"precedents",true),{status:"none",addresses:[]});
const traceMissingSheetExcel = {run:async fn=>fn({workbook:{worksheets:{getItem:()=>({getRange:()=>({load(){}})})}},sync:async()=>{throw {code:"ItemNotFound"}}})};
check("missing sheet is not mistaken for no dependencies",await (async()=>{try{await lib.traceDirectRelations(traceMissingSheetExcel,traceRequest,"precedents",true);return false;}catch{return true;}})(),true);

const verifyOp = plan.normalizeOp({op:"setValues",sheet:"S",range:"A1:B1",values:[[1,2]]});
const verifySnapshot = {sheets:{S:{"A1:B1":{values:[[1,3]],formulas:[[1,3]],valueTypes:[["Double","Error"]]}}}};
const verifyResult = plan.verifyCellResults([verifyOp],verifySnapshot)[0];
check("post-write mismatch is detected",verifyResult.mismatches,1);
check("post-write Excel error is identified by value type",verifyResult.errors,[{cell:"S!B1",error:"3"}]);
check("post-write mismatch is not labelled verified",verifyResult.status,"mismatch");
const literalSnapshot = {sheets:{S:{"A1:B1":{values:[[1,"#REF!"]],formulas:[[1,"#REF!"]],valueTypes:[["Double","String"]]}}}};
check("literal error-like string is not an Excel error",plan.verifyCellResults([verifyOp],literalSnapshot)[0].errors,[]);

check("missing Excel error types prevent verified status",plan.verifyCellResults([verifyOp],{sheets:{S:{"A1:B1":{values:[[1,2]],formulas:[[1,2]]}}}})[0].status,"unverified");
const r1c1Op = plan.normalizeOp({op:"setFormulas",sheet:"S",range:"A1",formula:"=1",formulaR1C1:"=2"});
check("R1C1 takes priority and stays unverified",plan.verifyCellResults([r1c1Op],{sheets:{S:{"A1":{values:[[2]],formulas:[["=2"]],valueTypes:[["Double"]]}}}})[0].status,"unverified");
check("overlapping writes are unverified",plan.verifyCellResults([verifyOp,verifyOp],verifySnapshot)[0].status,"unverified");
check("missing readback cell is unverified",plan.verifyCellResults([verifyOp],{sheets:{S:{"A1:B1":{values:[[1]],formulas:[[1]],valueTypes:[["Double","Double"]]}}}})[0].status,"unverified");
check("search parser requires explicit quoted term",lib.parseWorkbookSearch('find @search("rent")'),"rent");
check("search parser rejects unquoted term",(()=>{try{lib.parseWorkbookSearch('@search(rent)');return false;}catch{return true;}})(),true);
const searchRange = {isNullObject:false,rowCount:2,columnCount:2,rowIndex:3,columnIndex:1,values:[["Rent",22],["other",3]],formulas:[["Rent",22],["other","=RENT()"]],load(){}};
const searchExcel = {run:async fn=>fn({workbook:{worksheets:{items:[{name:"Data",getUsedRangeOrNullObject:()=>searchRange}],load(){}}},sync:async()=>{}})};
const searchResult = await lib.searchWorkbook(searchExcel,"rent");
check("search returns matching addresses and formula",searchResult.matches.map(x=>[x.cell,x.value,x.formula]),[["B4","Rent",""],["C5","3","=RENT()"]]);
const fallbackSearchRange = {rowCount:1,columnCount:1,rowIndex:0,columnIndex:0,values:[["Rent"]],formulas:[["Rent"]],load(){}};
const fallbackSearchExcel = {run:async fn=>fn({workbook:{worksheets:{items:[{name:"Data",getUsedRange:()=>fallbackSearchRange}],load(){}}},sync:async()=>{}})};
check("search supports hosts without nullable used ranges",(await lib.searchWorkbook(fallbackSearchExcel,"rent")).totalMatches,1);
searchRange.rowCount=5001; searchRange.columnCount=1;
check("search refuses oversized used ranges",await (async()=>{try{await lib.searchWorkbook(searchExcel,"rent");return false;}catch{return true;}})(),true);
let backupClosed = false;
const backupFile = {size:3,sliceCount:2,getSliceAsync(i,cb){cb({status:"succeeded",value:{data:i?[3]:[1,2]}})},closeAsync(cb){backupClosed=true;cb({status:"succeeded"})}};
const backupOffice = {FileType:{Compressed:"compressed"},context:{document:{getFileAsync(type,opts,cb){cb({status:"succeeded",value:backupFile})}}}};
const backupChunks = await lib.readWorkbookBackup(backupOffice,10);
check("backup reads all slices in order",backupChunks.map(a=>Array.from(a)),[[1,2],[3]]);
check("backup closes the Office file",backupClosed,true);
backupClosed=false;
check("backup rejects too-large workbook",await (async()=>{try{await lib.readWorkbookBackup(backupOffice,2);return false;}catch{return true;}})(),true);
check("backup closes on failure",backupClosed,true);

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

const singleCellFormat = plan.normalizeOp({ op: "formatRange", sheet: "Data", range: "A1", format: { bold: true, fill: "#FFEEAA" } });
check("single-cell cell-local formatting is add-in undo-safe", plan.isUndoSafe(singleCellFormat), true);
check("multi-cell formatting is not add-in undo-safe", plan.isUndoSafe(safeFormat), false);
check("border formatting is not add-in undo-safe", plan.isUndoSafe(plan.normalizeOp({
  op: "formatRange", sheet: "Data", range: "A1",
  format: { borders: { EdgeBottom: { style: "Continuous", weight: "Thick", color: "#FF0000" } } },
})), false);
check("row-height formatting is not add-in undo-safe", plan.isUndoSafe(plan.normalizeOp({
  op: "formatRange", sheet: "Data", range: "A1", format: { rowHeight: 200 },
})), false);
check("column-width formatting is not add-in undo-safe", plan.isUndoSafe(plan.normalizeOp({
  op: "formatRange", sheet: "Data", range: "A1", format: { columnWidth: 300 },
})), false);
check("autofit formatting is not add-in undo-safe", plan.isUndoSafe(plan.normalizeOp({
  op: "formatRange", sheet: "Data", range: "A1", format: { autofitColumns: true },
})), false);

check("single-cell cell-local formatting can auto-apply", plan.canAutoApply([singleCellFormat], 5), true);
check("autofitRange is not auto-applied", plan.canAutoApply([plan.normalizeOp({ op: "autofitRange", sheet: "Data", range: "A1" })], 5000), false);
check("border formatting is not auto-applied", plan.canAutoApply([plan.normalizeOp({
  op: "formatRange", sheet: "Data", range: "A1",
  format: { borders: { EdgeBottom: { style: "Continuous" } } },
})], 5000), false);
check("row-height formatting is not auto-applied", plan.canAutoApply([plan.normalizeOp({
  op: "formatRange", sheet: "Data", range: "A1", format: { rowHeight: 200 },
})], 5000), false);
check("column-width formatting is not auto-applied", plan.canAutoApply([plan.normalizeOp({
  op: "formatRange", sheet: "Data", range: "A1", format: { columnWidth: 300 },
})], 5000), false);
check("autofit formatting is not auto-applied", plan.canAutoApply([plan.normalizeOp({
  op: "formatRange", sheet: "Data", range: "A1", format: { autofitColumns: true },
})], 5000), false);

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

// Snapshots must capture only the state the operation can change. Restoring
// range-level formatting for a value write can overwrite mixed formatting and
// whole-row/column dimensions outside the reviewed cells.
function makeRecordingSheet() {
  const state = { loads: [] };
  const range = {
    values: [["before"]],
    formulas: [["=1+1"]],
    numberFormat: [["General"]],
    load: function (props) { state.loads.push(props.slice()); },
    format: {
      horizontalAlignment: "Left", verticalAlignment: "Top", wrapText: false,
      font: { bold: true, italic: false, size: 12, color: "#111111" },
      fill: { color: "#EEEEEE" },
      load: function (props) { state.loads.push(props.map(function (p) { return "format." + p; })); },
    },
  };
  range.format.font.load = function (props) { state.loads.push(props.map(function (p) { return "font." + p; })); };
  range.format.fill.load = function (props) { state.loads.push(["fill." + props]); };
  const sheet = { name: "Data", getRange: function () { return range; } };
  const Excel = {
    run: async function (fn) {
      return fn({
        workbook: {
          worksheets: {
            getActiveWorksheet: function () { return sheet; },
            getItem: function () { return sheet; },
          },
        },
        sync: async function () {},
      });
    },
  };
  return { Excel, state, range };
}

const writeSnap = makeRecordingSheet();
const writeSnapshot = await plan.snapshotCells(writeSnap.Excel, [
  plan.normalizeOp({ op: "setValues", sheet: "Data", range: "A1", values: [[1]] }),
]);
check("value-write snapshot loads values, formulas and types", writeSnap.state.loads, [["values", "formulas", "valueTypes"]]);
check("value-write snapshot captures no formatting", Object.keys(writeSnapshot.sheets.Data.A1).sort(), ["formulas", "values"]);
writeSnap.range.values = [["after"]];
writeSnap.range.formulas = [["=2+2"]];
writeSnap.range.format.font.bold = false;
await plan.undoSnapshot(writeSnap.Excel, writeSnapshot);
check("value-write undo restores values", writeSnap.range.values, [["before"]]);
check("value-write undo restores formulas", writeSnap.range.formulas, [["=1+1"]]);
check("value-write undo leaves formatting untouched", writeSnap.range.format.font.bold, false);

const formatSnap = makeRecordingSheet();
const formatSnapshot = await plan.snapshotCells(formatSnap.Excel, [
  plan.normalizeOp({ op: "formatRange", sheet: "Data", range: "A1", format: { bold: false } }),
]);
check("format snapshot loads number format and cell formatting", formatSnap.state.loads, [
  ["numberFormat"],
  ["format.horizontalAlignment", "format.verticalAlignment", "format.wrapText"],
  ["font.bold", "font.italic", "font.size", "font.color"],
  ["fill.color"],
]);
check("format snapshot captures no values", Object.keys(formatSnapshot.sheets.Data.A1).sort(), ["format", "numberFormat"]);
formatSnap.range.numberFormat = [["0.00"]];
formatSnap.range.format.horizontalAlignment = "Center";
formatSnap.range.format.font.bold = false;
formatSnap.range.format.fill.color = "#000000";
await plan.undoSnapshot(formatSnap.Excel, formatSnapshot);
check("format undo restores number format", formatSnap.range.numberFormat, [["General"]]);
check("format undo restores alignment", formatSnap.range.format.horizontalAlignment, "Left");
check("format undo restores font", formatSnap.range.format.font.bold, true);
check("format undo restores fill", formatSnap.range.format.fill.color, "#EEEEEE");

const mixedSnapshot = await plan.snapshotCells(writeSnap.Excel, [safeFormat]);
check("multi-cell formatting produces no undo snapshot", Object.keys(mixedSnapshot.sheets), []);

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

// --- row/column executor uses the normalized whole-axis ranges -------------
var structuralCallsDeferred = [];
var structuralSheet = {
  getRange: function (address) {
    return {
      insert: function (shift) { structuralCallsDeferred.push(["insert", address, shift]); },
      delete: function (shift) { structuralCallsDeferred.push(["delete", address, shift]); },
    };
  },
};
var structuralExcelDeferred = {
  run: async function (fn) {
    return fn({
      workbook: {
        worksheets: {
          getActiveWorksheet: function () { return structuralSheet; },
          getItem: function () { return structuralSheet; },
        },
      },
      sync: async function () {},
    });
  },
  InsertShift: { down: "down", right: "right" },
  DeleteShiftDirection: { up: "up", left: "left" },
};
const rowApply = plan.normalizeOp({ op: "insertRows", sheet: "Sheet1", range: "A5", count: 3 });
await plan.applyOps(structuralExcelDeferred, [rowApply]);
check("row executor targets whole rows", structuralCallsDeferred[0], ["insert", "5:7", "down"]);
const colApply = plan.normalizeOp({ op: "deleteColumns", sheet: "Sheet1", range: "C3", count: 2 });
await plan.applyOps(structuralExcelDeferred, [colApply]);
check("column executor targets whole columns", structuralCallsDeferred[1], ["delete", "C:D", "left"]);

// --- chart creation uses the documented Office.js argument order ------------
var deferredChartCalls = [];
var chartSheet = {
  getRange: function (address) { return { address: address }; },
  charts: {
    add: function (chartType, source, seriesBy) {
      deferredChartCalls.push([chartType, source && source.address, seriesBy]);
      return { name: "", setPosition: function () {} };
    },
  },
};
var deferredChartExcel = {
  run: async function (fn) {
    return fn({
      workbook: { worksheets: { getActiveWorksheet: function () { return chartSheet; }, getItem: function () { return chartSheet; } } },
      sync: async function () { throw new Error("chart add failed at sync"); },
    });
  },
};
var realExcel = globalThis.Excel;
globalThis.Excel = { ChartType: { ColumnClustered: "ColumnClustered" }, ChartSeriesBy: { columns: "Columns" } };
const deferredChartOp = plan.normalizeOp({ op: "addChart", sheet: "Sheet1", chartType: "ColumnClustered", source: "A1:B5", name: "Chart 1" });
const chartResults = await plan.applyOps(deferredChartExcel, [deferredChartOp]);
globalThis.Excel = realExcel;
check("chart add uses the documented argument order", deferredChartCalls[0], ["ColumnClustered", "A1:B5", "Columns"]);
check("chart add is attempted once (no fallback after deferred errors)", deferredChartCalls.length, 1);
check("deferred chart error is reported as a failed operation", chartResults[0].ok, false);
check("deferred chart error keeps the host error", /chart add failed at sync/.test(chartResults[0].error || ""), true);

// Review diffs and conservative drift checks are computed from frozen snapshots.
const reviewOps = [plan.normalizeOp({op:"setValues",sheet:"Data",range:"A1:B1",values:[["new",42]]})];
const oldSnapshot = {sheets:{Data:{"A1:B1":{values:[["old",1]],formulas:[["old",1]]}}}};
const proposedDiff = plan.previewCellChanges(reviewOps,oldSnapshot,50);
check("review diff contains exact first cell", proposedDiff.rows[0], {address:"Data!A1",before:"old",after:"new"});
check("review diff contains exact second cell", proposedDiff.rows[1], {address:"Data!B1",before:"1",after:"42"});
check("unchanged preview has no conflicts", plan.snapshotConflicts(oldSnapshot,JSON.parse(JSON.stringify(oldSnapshot)),reviewOps), []);
const changedSnapshot = JSON.parse(JSON.stringify(oldSnapshot)); changedSnapshot.sheets.Data["A1:B1"].values[0][1] = 5;
check("changed preview detects affected range", plan.snapshotConflicts(oldSnapshot,changedSnapshot,reviewOps), ["Data!A1:B1"]);
check("missing snapshot fails closed", plan.snapshotConflicts(oldSnapshot,{sheets:{}},reviewOps), ["Data!A1:B1"]);
check("visual diff respects limit", plan.previewCellChanges(reviewOps,oldSnapshot,1).truncated, true);
const mergedSnapshot = {sheets:{Data:{"A1:B1":{values:[[1,2]],formulas:[[1,2]],format:{font:{bold:true}},numberFormat:[["General"]]}}}};
const filtered = plan.selectSnapshot(mergedSnapshot,reviewOps);
check("failed-format properties excluded from undo snapshot", filtered.sheets.Data["A1:B1"].format, undefined);

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
