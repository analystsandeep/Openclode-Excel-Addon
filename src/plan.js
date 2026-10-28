/* OpenCode for Excel - structured edit plans.
 *
 * The assistant is asked to answer modification requests with a JSON plan
 * instead of free text. The pane reviews the plan (cell-level references,
 * destructive-flag, per-operation approval), the user approves individual
 * operations, and this module executes them through Office.js. Because
 * Excel may make Office.js edits available to its native Undo command. This
 * module also keeps bounded before-snapshots for a small, explicit set of
 * cell-only operations. Structural and object operations are never presented
 * as safely reversible by the add-in.
 *
 * Plan schema (version 1):
 * {
 *   "summary": "human readable one-liner",
 *   "assumptions": ["..."],
 *   "operations": [
 *     { "op": "addSheet", "name": "Budget", "description": "..." },
 *     { "op": "setValues", "sheet": "Sheet1", "range": "A1:D1",
 *       "values": [["Date","Region","Units","Amount"]], "description": "..." },
 *     { "op": "setFormulas", "sheet": "Sheet1", "range": "E2:E10",
 *       "formula": "=B2*C2", "description": "..." },
 *     { "op": "insertRows", "sheet": "Sheet1", "at": 5, "count": 3 },
 *     { "op": "formatRange", "sheet": "Sheet1", "range": "A1:D1",
 *       "format": { "bold": true, "fill": "#1F3864", "numberFormat": "$#,##0" } },
 *     { "op": "addChart", "sheet": "Sheet1", "source": "A1:D10",
 *       "chartType": "ColumnClustered", "position": { "left": 400, "top": 100 } }
 *   ]
 * }
 *
 * Exposes window.OpenCodeExcelPlan.
 */

(function (root) {
  "use strict";

  var ALLOWED_OPS = [
    "addSheet", "renameSheet", "deleteSheet",
    "setValues", "setFormulas", "clearRange",
    "insertRows", "deleteRows", "insertColumns", "deleteColumns",
    "formatRange", "autofitRange", "mergeRange", "unmergeRange",
    "addTable", "deleteTable", "sortRange",
    "addChart", "deleteChart", "freezePanes"
  ];

  var CHART_TYPES = {
    ColumnClustered: "ColumnClustered",
    BarClustered: "BarClustered",
    Line: "Line",
    LineMarkers: "LineMarkers",
    Pie: "Pie",
    Doughnut: "Doughnut",
    Area: "Area",
    Scatter: "Scatter"
  };

  var MAX_PREVIEW_CELLS = 5000; // safety cap for snapshot/apply reads

  function columnName(index) {
    var out = "";
    var n = Number(index) + 1;
    while (n > 0) {
      var rem = (n - 1) % 26;
      out = String.fromCharCode(65 + rem) + out;
      n = Math.floor((n - 1) / 26);
    }
    return out;
  }

  function columnIndex(name) {
    var n = 0;
    String(name || "").toUpperCase().split("").forEach(function (ch) {
      n = n * 26 + ch.charCodeAt(0) - 64;
    });
    return n - 1;
  }

  /** Exact cell count for finite A1 ranges; unknown/named/whole ranges are unbounded. */
  function rangeCellCount(address) {
    var ref = addressParts(address).ref.replace(/\$/g, "");
    var m = ref.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
    if (m) {
      var rows = Math.abs(Number(m[4] || m[2]) - Number(m[2])) + 1;
      var cols = Math.abs(columnIndex(m[3] || m[1]) - columnIndex(m[1])) + 1;
      return rows * cols;
    }
    m = ref.match(/^(\d+):(\d+)$/);
    if (m) return Infinity; // whole rows: exact width depends on the used range
    m = ref.match(/^([A-Z]+):([A-Z]+)$/i);
    if (m) return Infinity; // whole columns: exact height depends on the used range
    return Infinity;
  }

  function rangeDimensions(address) {
    var ref = addressParts(address).ref.replace(/\$/g, "");
    var m = ref.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
    if (!m) return { rows: 1, columns: 1 };
    return {
      rows: Math.abs(Number(m[4] || m[2]) - Number(m[2])) + 1,
      columns: Math.abs(columnIndex(m[3] || m[1]) - columnIndex(m[1])) + 1,
    };
  }

  /** Row span of a cell range or whole-row range, using 1-based row numbers. */
  function rowSpan(ref) {
    var clean = String(ref || "").replace(/\$/g, "");
    var m = clean.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
    if (m) {
      var start = Number(m[2]);
      var end = Number(m[4] || m[2]);
      return { start: Math.min(start, end), count: Math.abs(end - start) + 1 };
    }
    m = clean.match(/^(\d+)(?::(\d+))?$/);
    if (m) {
      var a = Number(m[1]);
      var b = Number(m[2] || m[1]);
      return { start: Math.min(a, b), count: Math.abs(b - a) + 1 };
    }
    return null;
  }

  /** Column span of a cell range or whole-column range, using 0-based indexes. */
  function columnSpan(ref) {
    var clean = String(ref || "").replace(/\$/g, "");
    var m = clean.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
    if (m) {
      var start = columnIndex(m[1]);
      var end = columnIndex(m[3] || m[1]);
      return { start: Math.min(start, end), count: Math.abs(end - start) + 1 };
    }
    m = clean.match(/^([A-Z]+)(?::([A-Z]+))?$/i);
    if (m) {
      var a = columnIndex(m[1]);
      var b = columnIndex(m[2] || m[1]);
      return { start: Math.min(a, b), count: Math.abs(b - a) + 1 };
    }
    return null;
  }

  /**
   * Row/column edits must affect whole rows or columns. A cell-shaped range
   * such as A5 with count 3 would otherwise insert only cells in column A.
   * Normalize to an exact axis range and reject ranges/counts that disagree.
   */
  function normalizeStructuralOp(o, op) {
    var isRow = o.op === "insertRows" || o.op === "deleteRows";
    var label = isRow ? "row" : "column";
    var hasRange = !!o.range;
    var hasAt = op.at != null && op.at !== "";
    if (hasRange && hasAt) {
      return { error: o.op + " accepts either at + count or a range, not both." };
    }
    if (!hasRange && !hasAt) {
      return { error: o.op + " requires at + count or a range." };
    }

    var start;
    var spanCount;
    if (hasRange) {
      var span = isRow ? rowSpan(o.range) : columnSpan(o.range);
      if (!span) {
        return { error: o.op + " requires a cell range or a whole-" + label + " range; '" + o.range + "' is not one." };
      }
      start = span.start;
      spanCount = span.count;
    } else if (isRow) {
      start = Number(op.at);
      if (!Number.isInteger(start) || start < 1) {
        return { error: o.op + " requires a positive row number in at." };
      }
      spanCount = 1;
    } else {
      start = typeof op.at === "string" && /^[A-Z]+$/i.test(op.at)
        ? columnIndex(op.at)
        : Number(op.at) - 1;
      if (!Number.isInteger(start) || start < 0) {
        return { error: o.op + " requires a valid column letter or number in at." };
      }
      spanCount = 1;
    }

    var requested = op.count != null ? Number(op.count) : null;
    if (requested != null) {
      if (!Number.isInteger(requested) || requested < 1) {
        return { error: o.op + " requires count to be a positive integer." };
      }
      // A single cell/row/column plus count means "start here and span count".
      // A multi-row/column range must agree with count instead of silently
      // ignoring either value.
      if (spanCount > 1 && requested !== spanCount) {
        return { error: o.op + " range spans " + spanCount + " " + label + "(s) but count is " + requested + "; make them agree or omit count." };
      }
      spanCount = requested;
    }

    var axisLimit = isRow ? 1048576 : 16384;
    if (!Number.isInteger(start) || start < (isRow ? 1 : 0) ||
        start + spanCount - 1 > (isRow ? axisLimit : axisLimit - 1)) {
      return { error: o.op + " exceeds the worksheet " + label + " boundary." };
    }

    if (isRow) {
      return {
        range: start + ":" + (start + spanCount - 1),
        count: spanCount,
        at: start,
      };
    }
    return {
      range: columnName(start) + ":" + columnName(start + spanCount - 1),
      count: spanCount,
      at: columnName(start),
    };
  }

  /* ------------------------------------------------------------- parsing */

  function addressParts(address) {
    // "Sheet1!A1:D20" -> { sheet: "Sheet1", ref: "A1:D20" }
    // "A1:D20"        -> { sheet: null, ref: "A1:D20" }
    var m = String(address || "").match(/^([^!]+)!(.+)$/);
    if (m) return { sheet: m[1], ref: m[2] };
    return { sheet: null, ref: String(address || "") };
  }

  function defaultDescription(o, op) {
    switch (o.op) {
      case "addSheet":   return "Add sheet '" + String(op.name || "New Sheet") + "'.";
      case "renameSheet": return "Rename sheet '" + String(op.from) + "' to '" + String(op.to) + "'.";
      case "deleteSheet": return "Delete sheet '" + String(op.name || "this sheet") + "'.";
      case "setValues":   return "Write values into " + (o.address || "the sheet") + ".";
      case "setFormulas": return "Write formula " + String(o.formula || "") + " into " + (o.address || "the sheet") + ".";
      case "clearRange": return "Clear " + (o.address || "the selected range") + ".";
      case "insertRows":  return "Insert " + (o.count || 1) + " row(s) at row " + (o.at != null ? o.at : "?") + " in " + (o.sheet || "the sheet") + ".";
      case "deleteRows":  return "Delete " + (o.count || 1) + " row(s) at row " + (o.at != null ? o.at : "?") + " in " + (o.sheet || "the sheet") + ".";
      case "insertColumns": return "Insert " + (o.count || 1) + " column(s) at column " + (o.at != null ? o.at : "?") + " in " + (o.sheet || "the sheet") + ".";
      case "deleteColumns": return "Delete " + (o.count || 1) + " column(s) at column " + (o.at != null ? o.at : "?") + " in " + (o.sheet || "the sheet") + ".";
      case "formatRange": return "Format " + (o.address || "the selection") + ".";
      case "autofitRange": return "Autofit " + (o.address || "the selection") + ".";
      case "mergeRange": return "Merge " + (o.address || "the selection") + ".";
      case "unmergeRange": return "Unmerge " + (o.address || "the selection") + ".";
      case "addTable": return "Create table '" + String(o.name || "Table") + "' from " + (o.address || "the data") + ".";
      case "deleteTable": return "Delete table '" + String(o.name || "") + "'.";
      case "sortRange": return "Sort " + (o.address || "the data") + ".";
      case "addChart":    return "Add " + (o.chartType || "chart") + " from " + (o.source || o.address || "the data") + ".";
      case "deleteChart": return "Delete chart '" + String(o.name || "") + "'.";
      case "freezePanes": return "Freeze " + Number(o.rows || 0) + " row(s) and " + Number(o.columns || 0) + " column(s).";
      default:            return "Edit the workbook.";
    }
  }

  function destructiveOp(o) {
    switch (o.op) {
      case "setValues": case "setFormulas": case "clearRange": return true;   // overwrites cells
      case "deleteRows": case "deleteColumns": case "deleteSheet":
      case "deleteTable": case "deleteChart": return true;
      case "insertRows": case "insertColumns": return true; // shifts existing data
      case "mergeRange": case "unmergeRange": case "sortRange": return true;
      default: return false;
    }
  }

  // These formatting properties are not cell-local. Borders are not captured
  // by the snapshot, and row/column dimensions (including autofit) affect cells
  // outside the reviewed range.
  var NON_LOCAL_FORMAT_KEYS = ["borders", "rowHeight", "columnWidth", "autofitColumns", "autofitRows"];

  function formatHasOnlyCellLocalProperties(o) {
    if (!o || o.op !== "formatRange" || !o.format || typeof o.format !== "object") return false;
    for (var i = 0; i < NON_LOCAL_FORMAT_KEYS.length; i++) {
      if (o.format[NON_LOCAL_FORMAT_KEYS[i]] != null) return false;
    }
    return true;
  }

  function finiteCellOperation(o) {
    return !!o && Number.isFinite(o.cellCount) && o.cellCount > 0 && o.cellCount <= MAX_PREVIEW_CELLS;
  }

  // Only cell-only writes and single-cell, cell-local formatting can be
  // restored exactly. A range-level format read can hide mixed cell formatting,
  // and the snapshot deliberately does not capture borders or dimensions.
  function isUndoSafe(o) {
    if (!o || !finiteCellOperation(o)) return false;
    if (o.op === "setValues" || o.op === "setFormulas") return true;
    return o.op === "formatRange" && o.cellCount === 1 && formatHasOnlyCellLocalProperties(o);
  }

  // Automatic approval is limited to formatting that only affects the reviewed
  // cells. autofitRange and dimension changes can affect whole rows/columns, so
  // they always require explicit review.
  function canAutoApply(ops, maxCells) {
    if (!Array.isArray(ops) || !ops.length) return false;
    var limit = Math.min(MAX_PREVIEW_CELLS, Math.max(0, Number(maxCells || 0)));
    var total = 0;
    for (var i = 0; i < ops.length; i++) {
      var o = ops[i];
      if (!o || o.destructive || o.op !== "formatRange" || !finiteCellOperation(o) ||
          !formatHasOnlyCellLocalProperties(o)) return false;
      total += o.cellCount;
      if (!Number.isFinite(total) || total > limit) return false;
    }
    return true;
  }

  function normalizeOp(op) {
    var o = {
      op: null, sheet: null, range: null, address: null,
      at: null, count: null, values: null, formula: null, formulas: null, formulaR1C1: null,
      format: null, chartType: null, source: null, position: null,
      name: null, from: null, to: null, style: null, hasHeaders: true,
      sort: null, applyTo: null, rows: 0, columns: 0,
      description: "", cells: [], cellCount: 0, destructive: false, ok: true, error: null
    };
    if (!op || typeof op !== "object") {
      o.ok = false; o.error = "Operation is not an object.";
      return o;
    }
    o.op = op.op || null;
    if (!o.op || ALLOWED_OPS.indexOf(o.op) < 0) {
      o.ok = false; o.error = "Unknown or missing operation: " + (op.op || "");
      return o;
    }
    var rawRange = op.range || op.address || "";
    var qualifiedRange = op.sheet && rawRange && rawRange.indexOf("!") < 0 ? op.sheet + "!" + rawRange : rawRange;
    var parts = addressParts(qualifiedRange);
    o.sheet = parts.sheet || op.sheet || null;
    o.range = parts.ref || op.range || op.address || null;
    o.address = o.range ? (o.sheet ? o.sheet + "!" + o.range : o.range) : null;
    o.at = op.at != null ? Number(op.at) : null;
    o.count = op.count != null ? Number(op.count) : null;
    o.name = op.name || null;
    o.from = op.from || null;
    o.to = op.to || null;
    o.values = op.values != null ? op.values : null;
    o.formula = op.formula || null;
    o.formulas = op.formulas != null ? op.formulas : null;
    o.formulaR1C1 = op.formulaR1C1 || null;
    o.format = op.format && typeof op.format === "object" ? op.format : null;
    o.chartType = op.chartType && CHART_TYPES[op.chartType] ? op.chartType : null;
    o.source = op.source || null;
    o.position = op.position && typeof op.position === "object" ? op.position : null;
    o.style = op.style || null;
    o.hasHeaders = op.hasHeaders !== false;
    o.sort = Array.isArray(op.sort) ? op.sort : [];
    o.applyTo = op.applyTo || "All";
    o.rows = Math.max(0, Number(op.rows || 0));
    o.columns = Math.max(0, Number(op.columns || 0));

    // Row/column edits must be normalized before execution. Cell-shaped ranges
    // are expanded to whole rows/columns so adjacent data cannot shift apart.
    if (/^(insertRows|deleteRows|insertColumns|deleteColumns)$/.test(o.op)) {
      var structural = normalizeStructuralOp(o, op);
      if (structural.error) {
        o.ok = false;
        o.error = structural.error;
      } else {
        o.range = structural.range;
        o.count = structural.count;
        o.at = structural.at;
      }
    }
    if (o.op === "renameSheet") o.sheet = o.from || o.sheet;
    if (o.op === "deleteSheet") o.sheet = o.name || o.sheet;
    if (o.op === "addChart" && !o.range) o.range = o.source;
    if (o.op === "addTable" && !o.range) o.range = op.source || null;
    o.address = o.range ? (o.sheet ? o.sheet + "!" + o.range : o.range) : null;
    o.description = /^(insertRows|deleteRows|insertColumns|deleteColumns)$/.test(o.op)
      ? defaultDescription(o, op) : (String(op.description || "").trim() || defaultDescription(o, op));
    o.cells = o.address ? [o.address] : [];
    // Operations without a cell range are structural/object operations, not
    // zero-cell edits. Infinity prevents them from satisfying size gates.
    o.cellCount = o.address ? rangeCellCount(o.address) : Infinity;
    o.destructive = destructiveOp(o);

    var needsRange = ["setValues", "setFormulas", "clearRange", "insertRows", "deleteRows",
      "insertColumns", "deleteColumns", "formatRange", "autofitRange", "mergeRange",
      "unmergeRange", "addTable", "sortRange", "addChart"].indexOf(o.op) >= 0;
    if (o.ok && needsRange && (!o.sheet || !o.range)) {
      o.ok = false; o.error = o.op + " requires an exact sheet and range.";
    } else if (o.op === "setValues" && o.values == null) {
      o.ok = false; o.error = "setValues requires values.";
    } else if (o.op === "setFormulas" && o.formula == null && o.formulas == null && o.formulaR1C1 == null) {
      o.ok = false; o.error = "setFormulas requires formula, formulas, or formulaR1C1.";
    } else if ((o.op === "addSheet" || o.op === "deleteSheet" || o.op === "addTable" || o.op === "deleteTable" || o.op === "deleteChart") && !o.name) {
      o.ok = false; o.error = o.op + " requires a name.";
    } else if (o.op === "renameSheet" && (!o.from || !o.to)) {
      o.ok = false; o.error = "renameSheet requires from and to.";
    } else if (Number.isFinite(o.cellCount) && o.cellCount > MAX_PREVIEW_CELLS) {
      o.ok = false; o.error = o.op + " affects " + o.cellCount + " cells; the safety limit is " + MAX_PREVIEW_CELLS + ".";
    }
    return o;
  }

  /** Parse a JSON plan out of an assistant answer. Returns {ok, operations, summary, destructive, cellsAffected} or {ok:false, error}. */
  function parsePlan(text) {
    if (!text) return { ok: false, error: "Empty answer." };
    var body = null;
    var fence = String(text).match(/```(?:json)?\s*\n([\s\S]*?)```/i);
    if (fence) {
      body = fence[1];
    } else {
      var trimmed = String(text).trim();
      if (trimmed.startsWith("{") && trimmed.endsWith("}")) body = trimmed;
    }
    if (body == null) return { ok: false, error: "No JSON plan found in the answer." };
    var raw;
    try {
      raw = JSON.parse(body);
    } catch (err) {
      return { ok: false, error: "The plan was not valid JSON: " + err.message };
    }
    if (!raw || typeof raw !== "object" || !Array.isArray(raw.operations)) {
      return { ok: false, error: "Expected a JSON object with an operations array." };
    }
    var ops = (raw.operations || []).map(normalizeOp);
    var bad = ops.filter(function (o) { return !o.ok; });
    if (bad.length) {
      return { ok: false, error: "Invalid plan step(s): " + bad.map(function (o) { return o.error; }).join("; "), partial: ops };
    }
    return {
      ok: true,
      raw: raw,
      operations: ops,
      summary: String(raw.summary || "").trim() || "Workbook changes",
      assumptions: Array.isArray(raw.assumptions) ? raw.assumptions : [],
      destructive: ops.some(function (o) { return o.destructive; }),
      cellsAffected: uniqueFlat(ops.map(function (o) { return o.cells; })),
      cellCount: ops.reduce(function (sum, o) {
        return sum === Infinity || o.cellCount === Infinity ? Infinity : sum + o.cellCount;
      }, 0),
    };
  }

  function uniqueFlat(lists) {
    var seen = {}, out = [];
    lists.forEach(function (list) {
      (list || []).forEach(function (c) { if (c && !seen[c]) { seen[c] = true; out.push(c); } });
    });
    return out;
  }

  /** Human-readable preview rows for the change-review UI. */
  function planLines(ops) {
    return ops.map(function (o) {
      return { op: o.op, label: o.description, cells: o.cells, destructive: o.destructive, address: o.address };
    });
  }

  /* ------------------------------------------------------------- Excel glue */

  function as2d(values) {
    if (!Array.isArray(values)) return [[values]];
    if (values.length && Array.isArray(values[0])) return values;
    return values.map(function (v) { return [v]; });
  }

  function resolveSheet(ctx, name) {
    if (!name) return ctx.workbook.worksheets.getActiveWorksheet();
    // Fail closed: a missing or invalid sheet must never redirect writes to
    // whichever worksheet happens to be active.
    return ctx.workbook.worksheets.getItem(name);
  }

  function applyFormat(range, fmt) {
    if (!fmt || typeof fmt !== "object") return;
    if (fmt.bold != null) range.format.font.bold = fmt.bold;
    if (fmt.italic != null) range.format.font.italic = fmt.italic;
    if (fmt.size != null) range.format.font.size = Number(fmt.size);
    if (fmt.color) range.format.font.color = String(fmt.color);
    if (fmt.fill) range.format.fill.color = String(fmt.fill);
    if (fmt.numberFormat) range.numberFormat = String(fmt.numberFormat);
    if (fmt.horizontalAlignment) range.format.horizontalAlignment = String(fmt.horizontalAlignment);
    if (fmt.verticalAlignment) range.format.verticalAlignment = String(fmt.verticalAlignment);
    if (fmt.wrapText != null) range.format.wrapText = fmt.wrapText;
    if (fmt.rowHeight != null) range.format.rowHeight = Number(fmt.rowHeight);
    if (fmt.columnWidth != null) range.format.columnWidth = Number(fmt.columnWidth);
    if (fmt.autofitColumns && range.format.autofitColumns) range.format.autofitColumns();
    if (fmt.autofitRows && range.format.autofitRows) range.format.autofitRows();
    if (fmt.borders && range.format.borders && range.format.borders.getItem) {
      Object.keys(fmt.borders).forEach(function (edge) {
        var spec = fmt.borders[edge] || {};
        var border = range.format.borders.getItem(edge);
        if (spec.style) border.style = spec.style;
        if (spec.weight) border.weight = spec.weight;
        if (spec.color) border.color = spec.color;
      });
    }
  }

  function matrix(rows, columns, value) {
    var out = [];
    for (var r = 0; r < rows; r++) {
      var row = [];
      for (var c = 0; c < columns; c++) row.push(value);
      out.push(row);
    }
    return out;
  }

  function writeFormulaRange(range, op) {
    var dims = rangeDimensions(op.range);
    if (op.formulas != null) {
      range.formulas = as2d(op.formulas);
      return;
    }
    if (op.formulaR1C1 != null) {
      range.formulasR1C1 = matrix(dims.rows, dims.columns, op.formulaR1C1);
      return;
    }
    if (dims.rows === 1 && dims.columns === 1) {
      range.formulas = [[op.formula]];
      return;
    }

    // Let Excel translate relative references when possible. For older hosts or
    // test doubles without fillDown/fillRight, fall back to an explicit matrix.
    if (typeof range.getCell === "function") {
      range.getCell(0, 0).formulas = [[op.formula]];
      if (dims.columns > 1 && typeof range.getRow === "function") {
        var firstRow = range.getRow(0);
        if (typeof firstRow.fillRight === "function") {
          firstRow.fillRight();
          if (dims.rows === 1) return;
        }
      }
      if (dims.rows > 1 && typeof range.fillDown === "function") {
        range.fillDown();
        return;
      }
    }
    range.formulas = matrix(dims.rows, dims.columns, op.formula);
  }

  function addChartOp(sheet, op) {
    var typeName = op.chartType || "ColumnClustered";
    var ChartType = (typeof Excel !== "undefined" && Excel.ChartType) || {};
    var chartType = ChartType[typeName];
    if (!chartType) throw new Error("Unsupported chart type '" + typeName + "' (charts need ExcelApi 1.8+)");
    var source = sheet.getRange(op.source || op.range);
    var SeriesBy = (typeof Excel !== "undefined" && Excel.ChartSeriesBy) || {};
    var seriesBy = SeriesBy.columns;
    // Microsoft documents ChartCollection.add(type, sourceData, seriesBy).
    // Do not fall back after a deferred sync error: by then the chart may
    // already exist, and a second add could create a duplicate.
    var chart = sheet.charts.add(chartType, source, seriesBy);
    if (op.position && chart && typeof chart.setPosition === "function") {
      try {
        if (op.position.startCell) chart.setPosition(op.position.startCell, op.position.endCell || undefined);
        if (op.position.left != null) chart.left = Number(op.position.left);
        if (op.position.top != null) chart.top = Number(op.position.top);
        if (op.position.width != null) chart.width = Number(op.position.width);
        if (op.position.height != null) chart.height = Number(op.position.height);
      } catch (err) { /* position is optional */ }
    }
    if (op.name) chart.name = op.name;
    return chart;
  }

  /**
   * State needed to undo one safe operation. Value/formula writes must not
   * restore formatting: a range-level format read can hide mixed cell state and
   * would touch rows/columns outside the reviewed cells. Single-cell formatting
   * captures only the cell-local properties it can restore exactly.
   */
  function snapshotRequirements(op) {
    if (op.op === "setValues" || op.op === "setFormulas") {
      return { values: true, formulas: true, valueTypes: true };
    }
    if (op.op === "formatRange" && op.cellCount === 1 && formatHasOnlyCellLocalProperties(op)) {
      return { format: true };
    }
    return null;
  }

  /** Read bounded before-state for explicitly undo-safe cell operations. */
  async function snapshotCells(Excel, ops) {
    var bySheet = {};
    var safeOps = (ops || []).filter(isUndoSafe);
    var total = safeOps.reduce(function (sum, o) { return sum + o.cellCount; }, 0);
    if (!Number.isFinite(total) || total > MAX_PREVIEW_CELLS) {
      throw new Error("Undo snapshot exceeds the " + MAX_PREVIEW_CELLS + "-cell safety limit.");
    }
    safeOps.forEach(function (o) {
      var count = rangeCellCount(o.address || o.range);
      if (!Number.isFinite(count) || count <= 0 || count > MAX_PREVIEW_CELLS) {
        throw new Error("Undo snapshot requires an exact range within the " + MAX_PREVIEW_CELLS + "-cell safety limit.");
      }
      var needs = snapshotRequirements(o);
      if (!needs) return;
      var s = o.sheet || "__active__";
      if (!bySheet[s]) bySheet[s] = {};
      if (!bySheet[s][o.range]) bySheet[s][o.range] = { values: false, formulas: false, valueTypes: false, format: false };
      var entry = bySheet[s][o.range];
      if (needs.values) entry.values = true;
      if (needs.formulas) entry.formulas = true;
      if (needs.valueTypes) entry.valueTypes = true;
      if (needs.format) entry.format = true;
    });
    var snap = { sheets: {}, at: new Date().toISOString() };
    for (var sheetName in bySheet) {
      var requirements = bySheet[sheetName];
      var addresses = Object.keys(requirements);
      var data = await Excel.run(async function (ctx) {
          var s = (sheetName === "__active__")
            ? ctx.workbook.worksheets.getActiveWorksheet()
            : resolveSheet(ctx, sheetName);
          var items = {};
          addresses.forEach(function (address) {
            var r = s.getRange(address);
            var need = requirements[address];
            var load = [];
            if (need.values) load.push("values");
            if (need.formulas) load.push("formulas");
            if (need.valueTypes) load.push("valueTypes");
            if (need.format) load.push("numberFormat");
            if (load.length) r.load(load);
            if (need.format && r.format && typeof r.format.load === "function") {
              r.format.load(["horizontalAlignment", "verticalAlignment", "wrapText"]);
              if (r.format.font && typeof r.format.font.load === "function") r.format.font.load(["bold", "italic", "size", "color"]);
              if (r.format.fill && typeof r.format.fill.load === "function") r.format.fill.load("color");
            }
            items[address] = r;
          });
          await ctx.sync();
          var out = {};
          addresses.forEach(function (address) {
            var r = items[address];
            var need = requirements[address];
            var entry = {};
            if (need.values) entry.values = r.values;
            if (need.formulas) entry.formulas = r.formulas;
            if (need.valueTypes) entry.valueTypes = r.valueTypes;
            if (need.format) {
              entry.numberFormat = r.numberFormat;
              entry.format = r.format ? {
                horizontalAlignment: r.format.horizontalAlignment,
                verticalAlignment: r.format.verticalAlignment,
                wrapText: r.format.wrapText,
                font: r.format.font ? {
                  bold: r.format.font.bold, italic: r.format.font.italic,
                  size: r.format.font.size, color: r.format.font.color,
                } : null,
                fill: r.format.fill ? { color: r.format.fill.color } : null,
              } : null;
            }
            out[address] = entry;
          });
          return out;
      });
      snap.sheets[sheetName] = data;
    }
    // Freeze the captured data independently of live Office or test objects.
    return JSON.parse(JSON.stringify(snap));
  }

  function selectSnapshot(snapshot, ops) {
    var out = { sheets: {} };
    (ops || []).forEach(function (op) {
      var need = snapshotRequirements(op), source = snapshot && snapshot.sheets && snapshot.sheets[op.sheet] && snapshot.sheets[op.sheet][op.range];
      if (!need || !source) return;
      if (!out.sheets[op.sheet]) out.sheets[op.sheet] = {};
      var entry = out.sheets[op.sheet][op.range] || (out.sheets[op.sheet][op.range] = {});
      if (need.values) entry.values = source.values;
      if (need.formulas) entry.formulas = source.formulas;
      if (need.valueTypes) entry.valueTypes = source.valueTypes;
      if (need.format) { entry.numberFormat = source.numberFormat; entry.format = source.format; }
    });
    return JSON.parse(JSON.stringify(out));
  }

  // Compare only the selected operations' captured properties. Conservative:
  // a recalculated formula value can trigger a re-review rather than risk an
  // overwriting edit. This is a check, not an atomic Excel transaction.
  function snapshotConflicts(expected, actual, ops) {
    var conflicts = [];
    (ops || []).forEach(function (op) {
      var needs = snapshotRequirements(op);
      if (!needs) return;
      var a = expected && expected.sheets && expected.sheets[op.sheet] && expected.sheets[op.sheet][op.range];
      var b = actual && actual.sheets && actual.sheets[op.sheet] && actual.sheets[op.sheet][op.range];
      if (!a || !b) { conflicts.push(op.address); return; }
      var keys = [];
      if (needs.values) keys.push("values");
      if (needs.formulas) keys.push("formulas");
      if (needs.valueTypes) keys.push("valueTypes");
      if (needs.format) keys.push("numberFormat", "format");
      if (keys.some(function (key) { return JSON.stringify(a[key]) !== JSON.stringify(b[key]); })) conflicts.push(op.address);
    });
    return conflicts.filter(function (value, index) { return conflicts.indexOf(value) === index; });
  }

  // Describe proposed changes without claiming that Excel has calculated them.
  // For fills with relative references, the seed is shown rather than invented
  // per-cell formulas. The visual preview is bounded but the guard is not.
  function previewCellChanges(ops, snapshot, limit) {
    var rows = [], truncated = false;
    limit = Math.max(1, Number(limit) || 50);
    (ops || []).forEach(function (op) {
      if (op.op !== "setValues" && op.op !== "setFormulas") return;
      var data = snapshot && snapshot.sheets && snapshot.sheets[op.sheet] && snapshot.sheets[op.sheet][op.range];
      if (!data) return;
      var match = String(op.range).replace(/\$/g, "").match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
      if (!match) return;
      var dims = rangeDimensions(op.range), baseCol = columnIndex(match[1].toUpperCase()), baseRow = Number(match[2]);
      var proposed = op.op === "setValues" ? as2d(op.values) : (op.formulas != null ? as2d(op.formulas) : null);
      for (var r = 0; r < dims.rows; r++) for (var c = 0; c < dims.columns; c++) {
        if (rows.length >= limit) { truncated = true; continue; }
        var beforeFormula = data.formulas && data.formulas[r] && data.formulas[r][c];
        var beforeValue = data.values && data.values[r] && data.values[r][c];
        var before = typeof beforeFormula === "string" && beforeFormula.charAt(0) === "=" ? beforeFormula : beforeValue;
        var after = proposed ? (proposed[r] && proposed[r][c]) : (op.formulaR1C1 || op.formula || "");
        var note = !proposed && dims.rows * dims.columns > 1 ? " (fill seed; Excel translates references)" : "";
        rows.push({ address: op.sheet + "!" + columnName(baseCol + c) + (baseRow + r), before: before == null ? "" : String(before), after: (after == null ? "[missing]" : String(after)) + note });
      }
    });
    return { rows: rows, truncated: truncated };
  }

  // A readback is not automatically verification: compare the exact supported
  // writes and only report Excel error cells when valueTypes identifies them.
  // Translated fills, formatting and overlapping writes remain unverified.
  function verifyCellResults(ops, snapshot) {
    var results = [];
    (ops || []).forEach(function (op, index) {
      if (op.op !== "setValues" && op.op !== "setFormulas") return;
      var data = snapshot && snapshot.sheets && snapshot.sheets[op.sheet] && snapshot.sheets[op.sheet][op.range];
      var result = { address: op.address, status: "unverified", mismatches: 0, errors: [], reason: "" };
      if (!data || !Array.isArray(data.values) || !Array.isArray(data.formulas)) {
        result.reason = "Readback unavailable"; results.push(result); return;
      }
      // Multiple writes on the same sheet might overlap, including different
      // address strings. Avoid wrongly attributing final data to one write.
      if ((ops || []).some(function (other, i) { return i !== index && other.sheet === op.sheet; })) {
        result.reason = "Multiple operations on the same sheet; final state not attributed to one operation";
      }
      var dims = rangeDimensions(op.range);
      var expected = op.op === "setValues" ? as2d(op.values) : (op.formulas != null ? as2d(op.formulas) : (op.formulaR1C1 == null && dims.rows === 1 && dims.columns === 1 && op.formula ? [[op.formula]] : null));
      var foundMissing = false;
      for (var r = 0; r < dims.rows; r++) for (var c = 0; c < dims.columns; c++) {
        var typ = data.valueTypes && data.valueTypes[r] && data.valueTypes[r][c];
        if (typ == null && !result.reason) result.reason = "Excel error-type data unavailable";
        if (String(typ || "").toLowerCase() === "error" && result.errors.length < 20) {
          var parts = addressParts(op.range).ref.replace(/\$/g, "").match(/^([A-Z]+)(\d+)/i);
          var address = parts ? columnName(columnIndex(parts[1].toUpperCase()) + c) + (Number(parts[2]) + r) : op.range;
          result.errors.push({ cell: op.sheet + "!" + address, error: String((data.values[r] || [])[c] || "Excel error") });
        }
        if (expected) {
          var actualRow = (op.op === "setValues" ? data.values : data.formulas)[r];
          if (!Array.isArray(expected[r]) || c >= expected[r].length || !Array.isArray(actualRow) || c >= actualRow.length) { foundMissing = true; continue; }
          var actual = actualRow[c];
          if (actual !== expected[r][c]) result.mismatches++;
        }
      }
      if (!result.reason && foundMissing) result.reason = "Readback dimensions incomplete";
      if (!result.reason && !expected) result.reason = "Relative or R1C1 formula fill requires host-specific verification";
      if (!result.reason) result.status = result.mismatches ? "mismatch" : "verified";
      results.push(result);
    });
    return results;
  }

  /** Execute one operation per Excel.run batch so one failure does not abort the rest. */
  async function applyOps(Excel, ops, { dryRun = false } = {}) {
    var results = [];
    for (var i = 0; i < ops.length; i++) {
      var op = ops[i];
      var result = { op: op.op, address: op.address, ok: false, error: null, cells: op.cells, destructive: op.destructive };
      if (dryRun) {
        result.ok = true; result.note = "preview";
        results.push(result);
        continue;
      }
      try {
        await Excel.run(async function (ctx) {
          var s = resolveSheet(ctx, op.sheet);
          switch (op.op) {
            case "addSheet":
              ctx.workbook.worksheets.add(op.name || "Sheet");
              break;
            case "renameSheet":
              s.name = op.to || op.name;
              break;
            case "deleteSheet":
              s.delete();
              break;
            case "setValues":
              s.getRange(op.range).values = as2d(op.values);
              break;
            case "setFormulas": {
              var rf = s.getRange(op.range);
              writeFormulaRange(rf, op);
              break;
            }
            case "clearRange":
              s.getRange(op.range).clear(op.applyTo || "All");
              break;
            case "insertRows":
              s.getRange(op.range).insert(Excel.InsertShift.down);
              break;
            case "deleteRows":
              s.getRange(op.range).delete(Excel.DeleteShiftDirection.up);
              break;
            case "insertColumns":
              s.getRange(op.range).insert(Excel.InsertShift.right);
              break;
            case "deleteColumns":
              s.getRange(op.range).delete(Excel.DeleteShiftDirection.left);
              break;
            case "formatRange":
              applyFormat(s.getRange(op.range), op.format);
              break;
            case "autofitRange": {
              var autoRange = s.getRange(op.range);
              if (autoRange.format.autofitColumns) autoRange.format.autofitColumns();
              if (autoRange.format.autofitRows) autoRange.format.autofitRows();
              break;
            }
            case "mergeRange":
              s.getRange(op.range).merge(false);
              break;
            case "unmergeRange":
              s.getRange(op.range).unmerge();
              break;
            case "addTable": {
              var table = s.tables.add(s.getRange(op.range), op.hasHeaders !== false);
              if (op.name) table.name = op.name;
              if (op.style) table.style = op.style;
              break;
            }
            case "deleteTable":
              s.tables.getItem(op.name).delete();
              break;
            case "sortRange":
              s.getRange(op.range).sort.apply(op.sort || [], false, op.hasHeaders !== false);
              break;
            case "addChart":
              addChartOp(s, op);
              break;
            case "deleteChart":
              s.charts.getItem(op.name).delete();
              break;
            case "freezePanes":
              if (op.rows > 0) s.freezePanes.freezeRows(op.rows);
              if (op.columns > 0) s.freezePanes.freezeColumns(op.columns);
              if (!op.rows && !op.columns && s.freezePanes.unfreeze) s.freezePanes.unfreeze();
              break;
            default:
              throw new Error("Unsupported operation " + op.op);
          }
          await ctx.sync();
        });
        result.ok = true;
      } catch (err) {
        result.error = err.message || String(err);
      }
      results.push(result);
    }
    return results;
  }

  /** Restore cell state captured by snapshotCells. No structural reversal. */
  async function undoSnapshot(Excel, snap) {
    var restored = 0;
    for (var sheetName in snap.sheets) {
      var sheet = sheetName;
      await Excel.run(async function (ctx) {
        var s = (sheet === "__active__")
          ? ctx.workbook.worksheets.getActiveWorksheet()
          : resolveSheet(ctx, sheet);
        var data = snap.sheets[sheet];
        var wrote = 0;
        for (var address in data) {
          var d = data[address];
          var r = s.getRange(address);
          if (d.values != null) { r.values = d.values; wrote++; }
          if (d.formulas != null) { r.formulas = d.formulas; wrote++; }
          if (d.numberFormat != null) r.numberFormat = d.numberFormat;
          if (d.format && r.format) {
            ["horizontalAlignment", "verticalAlignment", "wrapText"].forEach(function (key) {
              if (d.format[key] != null) r.format[key] = d.format[key];
            });
            if (d.format.font && r.format.font) {
              ["bold", "italic", "size", "color"].forEach(function (key) {
                if (d.format.font[key] != null) r.format.font[key] = d.format.font[key];
              });
            }
            if (d.format.fill && d.format.fill.color != null && r.format.fill) r.format.fill.color = d.format.fill.color;
          }
        }
        await ctx.sync();
        restored += wrote;
      });
    }
    return restored;
  }

  root.OpenCodeExcelPlan = {
    ALLOWED_OPS: ALLOWED_OPS,
    parsePlan: parsePlan,
    normalizeOp: normalizeOp,
    planLines: planLines,
    snapshotCells: snapshotCells,
    selectSnapshot: selectSnapshot,
    snapshotConflicts: snapshotConflicts,
    previewCellChanges: previewCellChanges,
    verifyCellResults: verifyCellResults,
    applyOps: applyOps,
    undoSnapshot: undoSnapshot,
    defaultDescription: defaultDescription,
    rangeCellCount: rangeCellCount,
    isUndoSafe: isUndoSafe,
    canAutoApply: canAutoApply,
    MAX_PREVIEW_CELLS: MAX_PREVIEW_CELLS,
  };
})(typeof globalThis !== "undefined" ? globalThis : this);
