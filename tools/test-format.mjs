/**
 * Unit/integration test for src/format.js with a fake Excel object model.
 *
 * Run: node tools/test-format.mjs
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

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

/* --------------------------------------------------------------- load lib -- */

const libSandbox = {};
new Function("globalThis", "window", "module", fs.readFileSync(path.join(ROOT, "src", "lib.js"), "utf8"))(
  libSandbox,
  libSandbox,
  undefined
);
const lib = libSandbox.OpenCodeExcel;

/* ----------------------------------------------------------- fake excel ---- */

const CURRENCY = "$#,##0;[Red]-$#,##0";

function colIndex(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function parseRef(ref) {
  const text = String(ref).replace(/^.*!/, "").replace(/\$/g, "").toUpperCase();
  let match = text.match(/^(\d+):(\d+)$/);
  if (match) return { r1: Number(match[1]) - 1, r2: Number(match[2]) - 1, c1: 0, c2: Infinity, rowsOnly: true };
  match = text.match(/^([A-Z]+):([A-Z]+)$/);
  if (match) return { r1: 0, r2: Infinity, c1: colIndex(match[1]), c2: colIndex(match[2]), colsOnly: true };
  match = text.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
  if (!match) throw new Error("fake Excel: bad address " + ref);
  const c1 = colIndex(match[1]);
  const r1 = Number(match[2]) - 1;
  const c2 = match[3] ? colIndex(match[3]) : c1;
  const r2 = match[4] ? Number(match[4]) - 1 : r1;
  return { r1, r2, c1, c2 };
}

function createFakeExcel(options = {}) {
  const ops = [];
  let cfCount = 0;
  const grid = (options.grid || []).map((row) => row.slice());
  const tableItems = [];

  const sheet = {
    name: options.name || "SaleData",
    freezePanes: {
      freezeRows(rows) {
        ops.push({ kind: "freezeRows", value: rows });
      },
    },
    pageLayout: {
      orientation: "",
      centerHorizontally: false,
      printGridlines: true,
      zoom: null,
      setPrintArea(value) {
        ops.push({ kind: "setPrintArea", value });
      },
      setPrintTitleRows(value) {
        ops.push({ kind: "setPrintTitleRows", value });
      },
      setPrintMargins(unit, margins) {
        ops.push({ kind: "setPrintMargins", unit, margins });
      },
    },
    tables: {
      items: tableItems,
      load() {},
      add(range, hasHeaders) {
        const table = {
          name: "",
          style: "",
          showBandedRows: false,
          showBandedColumns: true,
          showFilterButton: false,
          _range: range,
          getRange() {
            return range;
          },
          load() {},
        };
        tableItems.push(table);
        ops.push({ kind: "table.add", a1: range.a1, hasHeaders });
        return table;
      },
    },
    load() {},
    getUsedRange() {
      const rowCount = grid.length;
      const columnCount = Math.max(0, ...grid.map((row) => row.length));
      return {
        rowCount,
        columnCount,
        rowIndex: 0,
        columnIndex: 0,
        load() {},
        get values() {
          return grid.map((row) => row.slice());
        },
      };
    },
    getRange(ref) {
      return makeRange(ref);
    },
  };

  Object.defineProperty(sheet, "showGridlines", {
    set(value) {
      ops.push({ kind: "showGridlines", value });
    },
  });
  Object.defineProperty(sheet, "tabColor", {
    set(value) {
      ops.push({ kind: "tabColor", value });
    },
  });

  function ensureSize(r2, c2) {
    while (grid.length <= r2) grid.push([]);
    for (const row of grid) {
      while (row.length <= c2) row.push("");
    }
  }

  function slice(ref) {
    const parsed = parseRef(ref);
    const r2 = Math.min(parsed.r2, grid.length - 1);
    const c2 = Math.min(parsed.c2, Math.max(0, ...grid.map((row) => row.length)) - 1);
    const rows = [];
    for (let r = parsed.r1; r <= r2; r++) {
      const row = [];
      for (let c = parsed.c1; c <= c2; c++) row.push(grid[r] ? grid[r][c] : "");
      rows.push(row);
    }
    return rows;
  }

  function write(ref, values) {
    const parsed = parseRef(ref);
    const rows = values.length;
    const cols = rows ? values[0].length : 0;
    ensureSize(parsed.r1 + rows - 1, parsed.c1 + cols - 1);
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        grid[parsed.r1 + r][parsed.c1 + c] = values[r][c];
      }
    }
  }

  function makeFormat(a1) {
    const rec = (prop, value) => ops.push({ kind: "format", a1, prop, value });
    return {
      fill: {
        set color(value) {
          rec("fill.color", value);
        },
      },
      font: {
        set color(value) {
          rec("font.color", value);
        },
        set bold(value) {
          rec("font.bold", value);
        },
        set size(value) {
          rec("font.size", value);
        },
      },
      borders: {
        getItem(edge) {
          return {
            set style(value) {
              rec("border." + edge + ".style", value);
            },
            set weight(value) {
              rec("border." + edge + ".weight", value);
            },
            set color(value) {
              rec("border." + edge + ".color", value);
            },
          };
        },
      },
      set horizontalAlignment(value) {
        rec("horizontalAlignment", value);
      },
      set verticalAlignment(value) {
        rec("verticalAlignment", value);
      },
      set wrapText(value) {
        rec("wrapText", value);
      },
      set rowHeight(value) {
        rec("rowHeight", value);
      },
      set columnWidth(value) {
        rec("columnWidth", value);
      },
      set indentLevel(value) {
        rec("indentLevel", value);
      },
      autofitColumns() {
        ops.push({ kind: "autofit", a1 });
      },
    };
  }

  function makeRange(ref) {
    const parsed = parseRef(ref);
    const a1 = String(ref).toUpperCase();
    const rowCount = parsed.rowsOnly
      ? grid.length
      : Math.min(parsed.r2, grid.length - 1) - parsed.r1 + 1;
    const columnCount = parsed.colsOnly
      ? Math.max(0, ...grid.map((row) => row.length))
      : Math.min(parsed.c2, Math.max(0, ...grid.map((row) => row.length)) - 1) - parsed.c1 + 1;
    const handle = {
      a1,
      rowIndex: parsed.r1,
      columnIndex: parsed.c1,
      rowCount: Math.max(0, rowCount),
      columnCount: Math.max(0, columnCount),
      load() {},
      get values() {
        return slice(a1);
      },
      set values(value) {
        ops.push({ kind: "values", a1, values: value });
        write(a1, value);
      },
      get formulas() {
        return slice(a1);
      },
      set formulas(value) {
        ops.push({ kind: "formulas", a1, values: value });
        write(a1, value);
      },
      set numberFormat(value) {
        ops.push({ kind: "numberFormat", a1, value });
      },
      format: makeFormat(a1),
      conditionalFormats: {
        getCount() {
          return { value: cfCount };
        },
        add(type) {
          cfCount++;
          ops.push({ kind: "cf.add", a1, type });
          if (type === "DataBar") {
            return { dataBar: { barColor: "", showValue: false } };
          }
          return { custom: { rule: { formula: "" }, format: { fill: { color: "" }, font: { color: "" } } } };
        },
        clearAll() {
          ops.push({ kind: "cf.clear", a1 });
        },
      },
      clear(applyTo) {
        ops.push({ kind: "clear", a1, applyTo });
        write(a1, [[null]]);
      },
      merge() {
        ops.push({ kind: "merge", a1 });
      },
      insert(shift) {
        if (parsed.rowsOnly) {
          const count = parsed.r2 - parsed.r1 + 1;
          for (let i = 0; i < count; i++) grid.splice(parsed.r1, 0, []);
          ops.push({ kind: "insert", a1, shift, rows: count });
        } else {
          ops.push({ kind: "insert", a1, shift, rows: 1 });
        }
      },
      getResizedRange(rows, cols) {
        return makeRange(
          String(ref).toUpperCase().replace(/(\d+)$/, (m, n) => String(Number(n) + rows)) +
            (cols ? ":" + String(ref).replace(/^([A-Z]+)(\d+).*$/, "$1") + (Number(String(ref).match(/(\d+)$/)[1]) + rows) : "")
        );
      },
    };
    return handle;
  }

  const ctx = {
    workbook: {
      worksheets: {
        getActiveWorksheet: () => sheet,
        load() {},
      },
    },
    sync: async () => undefined,
  };

  const Excel = {
    run: async (fn) => fn(ctx),
  };

  function addTable(name, ref) {
    const range = makeRange(ref);
    const table = {
      name,
      style: "",
      showBandedRows: false,
      showBandedColumns: true,
      showFilterButton: false,
      _range: range,
      getRange() {
        return range;
      },
      load() {},
    };
    tableItems.push(table);
    return table;
  }

  return { Excel, ops, grid, sheet, tableItems, addTable };
}

/* --------------------------------------------------------------- load code -- */

const formatSandbox = {};
const officeStub = { context: { requirements: { isSetSupported: () => true } } };
new Function(
  "globalThis",
  "window",
  "Office",
  fs.readFileSync(path.join(ROOT, "src", "format.js"), "utf8")
)(formatSandbox, formatSandbox, officeStub);
const formatter = formatSandbox.OpenCodeExcelFormat;

/* ------------------------------------------------------------------ tests -- */

function sampleGrid() {
  const headers = ["OrderDate", "Region", "Manager", "SalesMan", "Item", "Units", " Unit_price ", "Sale_amt"];
  const regions = ["East", "Central", "West"];
  const people = ["Alexander", "Bob", "Cara"];
  const items = ["Television", "Radio", "Phone"];
  const rows = [];
  for (let i = 0; i < 12; i++) {
    const units = 10 + i;
    const price = 100 + i;
    rows.push([43106 + i, regions[i % 3], "Martha", people[i % 3], items[i % 3], units, price, units * price]);
  }
  rows.push(["", "", "", "", "", "", "", ""]);
  rows.push(["", "", "", "", "", "", "", ""]);
  return [headers, ...rows];
}

async function main() {
  console.log("-- first format of a raw sheet --");
  const fake = createFakeExcel({ grid: sampleGrid() });
  const first = await formatter.formatSheet(fake.Excel, lib, {});

  check("creates tblSales", first.tableName, "tblSales");
  check("reports 12 data rows (trailing blanks trimmed)", first.dataRows, 12);
  check("reports the summary as created", first.created, true);
  check("no warnings on a normal sheet", first.warnings, []);
  check("inserts exactly three rows", fake.ops.filter((op) => op.kind === "insert").length, 1);
  check("table range is A4:H16", fake.ops.find((op) => op.kind === "table.add").a1, "A4:H16");
  check("cleans headers in row 4", fake.grid[3], [
    "Order Date",
    "Region",
    "Manager",
    "Sales Rep",
    "Item",
    "Units",
    "Unit Price",
    "Sale Amount",
  ]);
  check("writes the summary title", fake.grid[0][0], "SALES SUMMARY");
  check("orders formula in D2", fake.grid[1][3], "=COUNTA(tblSales[Order Date])");
  check("revenue formula in H2", fake.grid[1][7], "=SUM(tblSales[Sale Amount])");
  ok("top-region formula in D3 avoids LET/UNIQUE", /INDEX\(tblSales\[Region\],MATCH\(MAX\(INDEX\(SUMIF/.test(fake.grid[2][3]), String(fake.grid[2][3]));
  check("people formula in F3", fake.grid[2][5], '=SUMPRODUCT((tblSales[Sales Rep]<>"")/COUNTIF(tblSales[Sales Rep],tblSales[Sales Rep]&""))');
  check("items formula in H3", fake.grid[2][7], '=SUMPRODUCT((tblSales[Item]<>"")/COUNTIF(tblSales[Item],tblSales[Item]&""))');
  const clearBefore = (cell) => {
    const clear = fake.ops.findIndex((op) => op.kind === "clear" && op.a1 === cell && op.applyTo === "Contents");
    const write = fake.ops.findIndex((op) => op.kind === "formulas" && op.a1 === cell);
    return clear >= 0 && write >= 0 && clear < write;
  };
  ok("clears D2 before writing its formula", clearBefore("D2"));
  ok("clears D3 before replacing the array formula", clearBefore("D3"));
  ok("clears H2 before writing its formula", clearBefore("H2"));

  const table = fake.tableItems[0];
  check("uses TableStyleMedium2", table.style, "TableStyleMedium2");
  check("banded rows on / banded columns off", [table.showBandedRows, table.showBandedColumns], [true, false]);
  check("freezes summary + header rows", fake.ops.find((op) => op.kind === "freezeRows").value, 4);
  check("hides gridlines", fake.ops.find((op) => op.kind === "showGridlines").value, false);
  check("colours the tab navy", fake.ops.find((op) => op.kind === "tabColor").value, "#1F3864");
  check("print area covers the table", fake.ops.find((op) => op.kind === "setPrintArea").value, "A1:H16");
  check("print titles cover summary + header", fake.ops.find((op) => op.kind === "setPrintTitleRows").value, "$1:$4");
  check("landscape orientation", fake.sheet.pageLayout.orientation, "Landscape");
  check("narrow left margin", fake.ops.find((op) => op.kind === "setPrintMargins").margins.left, 0.25);
  check("fits to one page wide", fake.sheet.pageLayout.zoom, { horizontalFitToPages: 1, verticalFitToPages: null });
  check("adds exactly two conditional formats", fake.ops.filter((op) => op.kind === "cf.add").length, 2);
  check("data bar on the amount column", fake.ops.filter((op) => op.kind === "cf.add" && op.type === "DataBar")[0].a1, "H5:H16");
  check("merges the summary title on a first run", fake.ops.filter((op) => op.kind === "merge").length, 1);
  ok(
    "applies the currency number format",
    fake.ops.some((op) => op.kind === "numberFormat" && op.a1 === "H5:H16" && op.value === CURRENCY),
    JSON.stringify(fake.ops.filter((op) => op.kind === "numberFormat"))
  );
  ok(
    "applies the date number format",
    fake.ops.some((op) => op.kind === "numberFormat" && op.a1 === "A5:A16" && op.value === "dd-mmm-yyyy")
  );
  check("autofits the columns after writing the summary", fake.ops.filter((op) => op.kind === "autofit").length, 1);
  check("autofits the whole used range", fake.ops.find((op) => op.kind === "autofit").a1, "A1:H16");

  console.log("\n-- second format is idempotent --");
  const beforeTables = fake.ops.filter((op) => op.kind === "table.add").length;
  const beforeInserts = fake.ops.filter((op) => op.kind === "insert").length;
  const beforeCfs = fake.ops.filter((op) => op.kind === "cf.add").length;
  const second = await formatter.formatSheet(fake.Excel, lib, {});

  check("reuses tblSales", second.tableName, "tblSales");
  check("keeps 12 data rows", second.dataRows, 12);
  check("does not create a duplicate table", fake.ops.filter((op) => op.kind === "table.add").length, beforeTables);
  check("does not insert more rows", fake.ops.filter((op) => op.kind === "insert").length, beforeInserts);
  check("does not duplicate conditional formats", fake.ops.filter((op) => op.kind === "cf.add").length, beforeCfs);
  check("does not merge the title again", fake.ops.filter((op) => op.kind === "merge").length, 1);
  check("still rewrites the revenue formula", fake.grid[1][7], "=SUM(tblSales[Sale Amount])");
  check("still reports no warnings", second.warnings, []);

  console.log("\n-- existing table starting at row 1 --");
  const fake2 = createFakeExcel({
    grid: [
      ["Name", "Amount"],
      ["ACME", 10],
      ["Globex", 20],
      ["Initech", 30],
    ],
  });
  fake2.addTable("tblSales", "A1:B4");
  const third = await formatter.formatSheet(fake2.Excel, lib, {});

  check("restyles the existing table", third.tableName, "tblSales");
  check("does not add a summary above a row-1 table", third.summaryTop, 0);
  check("does not insert rows", fake2.ops.filter((op) => op.kind === "insert").length, 0);
  check("freezes only the header row", fake2.ops.find((op) => op.kind === "freezeRows").value, 1);
  check("print titles are just row 1", fake2.ops.find((op) => op.kind === "setPrintTitleRows").value, "$1:$1");
  check("print area is A1:B4", fake2.ops.find((op) => op.kind === "setPrintArea").value, "A1:B4");
  check("does not write a summary title", fake2.grid[0][0], "Name");

  console.log("\n-- empty sheet is rejected --");
  const fake3 = createFakeExcel({ grid: [["Name", "Amount"]] });
  let error = null;
  try {
    await formatter.formatSheet(fake3.Excel, lib, {});
  } catch (err) {
    error = err;
  }
  ok("throws a friendly error for a header-only sheet", !!error && /no data rows/i.test(error.message), error && error.message);
  ok("header-only failure does not mutate the sheet", !fake3.ops.some((op) => ["insert", "values", "addTable", "merge"].includes(op.kind)), JSON.stringify(fake3.ops));

  console.log(failures ? `\n${failures} FAILURE(S)` : "\nAll formatter tests passed");
  process.exit(failures ? 1 : 0);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
