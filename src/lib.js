/* Pure helpers for the task pane (no DOM, no Office) - kept separate so they can be unit tested.
   Exposes window.OpenCodeExcel when loaded as a classic script in the pane. */
(function (root) {
  "use strict";

  function escapeHtml(text) {
    return String(text)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function colName(index /* 0-based */) {
    var name = "";
    var n = index + 1;
    while (n > 0) {
      var rem = (n - 1) % 26;
      name = String.fromCharCode(65 + rem) + name;
      n = Math.floor((n - 1) / 26);
    }
    return name;
  }

  function toA1(row0, col0) {
    return colName(col0) + (row0 + 1);
  }

  /** For attribute values: escapeHtml plus quotes so the attribute cannot be broken out of. */
  function escapeAttr(text) {
    return escapeHtml(text).replace(/"/g, "&quot;").replace(/'/g, "&#39;");
  }

  function inline(text) {
    return escapeHtml(text)
      .replace(/`([^`]+)`/g, "<code>$1</code>")
      .replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>")
      .replace(/(^|\W)\*([^*]+)\*/g, "$1<em>$2</em>");
  }

  /** Tiny markdown renderer: fenced code blocks, headings, lists, paragraphs. No dependencies. */
  function renderMarkdown(text) {
    var out = "";
    var blocks = String(text).split(/```/);
    blocks.forEach(function (block, i) {
      if (i % 2 === 1) {
        var nl = block.indexOf("\n");
        var code = nl === -1 ? "" : block.slice(nl + 1);
        out +=
          '<pre><button class="copy-code" data-code="' +
          escapeAttr(code) +
          '">Copy</button><code>' +
          escapeHtml(code.replace(/\n$/, "")) +
          "</code></pre>";
        return;
      }
      var html = "";
      var list = null;
      function closeList() {
        if (list) {
          html += "</" + list + ">";
          list = null;
        }
      }
      block.split("\n").forEach(function (raw) {
        var line = raw.trimEnd();
        if (!line.trim()) {
          closeList();
          return;
        }
        var heading = line.match(/^(#{1,4})\s+(.*)$/);
        if (heading) {
          closeList();
          html += "<p><strong>" + inline(heading[2]) + "</strong></p>";
          return;
        }
        var bullet = line.match(/^\s*[-*]\s+(.*)$/);
        if (bullet) {
          if (list !== "ul") {
            closeList();
            html += "<ul>";
            list = "ul";
          }
          html += "<li>" + inline(bullet[1]) + "</li>";
          return;
        }
        var ordered = line.match(/^\s*\d+[.)]\s+(.*)$/);
        if (ordered) {
          if (list !== "ol") {
            closeList();
            html += "<ol>";
            list = "ol";
          }
          html += "<li>" + inline(ordered[1]) + "</li>";
          return;
        }
        closeList();
        html += "<p>" + inline(line) + "</p>";
      });
      closeList();
      out += html;
    });
    return out;
  }

  /** Pull the first table out of an answer: fenced tsv/csv block, pipe table, or tabular lines. */
  function extractTable(text) {
    var fenced = String(text).match(/```(?:tsv|csv|table)?\n([\s\S]*?)```/i);
    var candidates = [];
    if (fenced) candidates.push(fenced[1]);
    candidates.push(String(text));

    for (var c = 0; c < candidates.length; c++) {
      var lines = candidates[c]
        .split("\n")
        .filter(function (l) {
          return l.trim().length > 0;
        });
      var rows = [];
      for (var i = 0; i < lines.length; i++) {
        var line = lines[i];
        var cells;
        if (line.indexOf("\t") !== -1) cells = line.split("\t");
        else if (/^\s*\|.*\|\s*$/.test(line)) cells = line.trim().replace(/^\||\|$/g, "").split("|");
        else if (line.indexOf(",") !== -1 && c === 0) cells = line.split(",");
        else continue;
        cells = cells.map(function (cell) {
          return cell.trim();
        });
        if (cells.length < 2) continue;
        if (!rows.length || rows[0].length === cells.length) rows.push(cells);
      }
      if (rows.length >= 2) return rows; // header + at least one data row
    }
    return null;
  }

  /** A single short line starting with "=" is treated as a formula the user can drop into a cell. */
  function singleFormula(text) {
    var lines = String(text)
      .split("\n")
      .map(function (l) {
        return l.trim();
      })
      .filter(Boolean);
    var candidates = lines.filter(function (l) {
      return /^=/.test(l) && l.length < 300;
    });
    return candidates.length === 1 ? candidates[0] : null;
  }

  /** Cell one row below a range address like "Sheet1!A1:D20" (with a blank spacer row). */
  function belowAddress(address) {
    if (!address) return null;
    var ref = String(address).split("!").pop();
    var match = ref.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
    if (!match) return null;
    var col = match[1];
    var lastRow = Number(match[4] || match[2]);
    return col + (lastRow + 2);
  }

  /** Top-left cell of a range address like "Sheet1!A1:D20". */
  function startAddress(address) {
    if (!address) return null;
    var ref = String(address).split("!").pop();
    var match = ref.match(/^([A-Z]+)(\d+)/i);
    return match ? match[1] + match[2] : null;
  }

  /* --------------------------------------------------------- sheet formatter */

  var HEADER_WORDS = {
    amt: "Amount",
    qty: "Quantity",
    qnty: "Quantity",
    num: "Number",
    no: "No.",
    id: "ID",
    sku: "SKU",
    rep: "Rep",
    cust: "Customer",
    desc: "Description",
    dept: "Department",
    mgr: "Manager",
    pct: "Percent",
    yr: "Year",
    mo: "Month",
    tel: "Telephone",
    url: "URL",
  };

  /** Turn a raw header into a readable, table-safe column name. */
  function cleanHeader(raw) {
    var text = String(raw == null ? "" : raw)
      .replace(/_/g, " ")
      .replace(/\s+/g, " ")
      .trim();
    if (!text) return "";
    text = text
      .replace(/sales\s*man/gi, "sales rep")
      .replace(/sales\s*person/gi, "sales rep")
      .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
      .replace(/[\[\]#@]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
    return text
      .split(" ")
      .map(function (word) {
        var lower = word.toLowerCase();
        if (HEADER_WORDS[lower]) return HEADER_WORDS[lower];
        if (word === word.toUpperCase() && word.length <= 4 && /[A-Z]/.test(word)) return word;
        return lower.charAt(0).toUpperCase() + lower.slice(1);
      })
      .join(" ");
  }

  /** Clean every header and make the names unique (Excel tables require that). */
  function cleanHeaders(rawHeaders) {
    var seen = {};
    return (rawHeaders || []).map(function (raw, index) {
      var name = cleanHeader(raw) || "Column " + (index + 1);
      var base = name;
      var suffix = 2;
      while (seen[name.toLowerCase()]) {
        name = base + " " + suffix;
        suffix++;
      }
      seen[name.toLowerCase()] = true;
      return name;
    });
  }

  function bestColumn(headers, tests) {
    var best = -1;
    var bestScore = 0;
    (headers || []).forEach(function (header, index) {
      var text = String(header || "").toLowerCase();
      tests.forEach(function (test) {
        var match = text.match(test[0]);
        if (match && test[1] > bestScore) {
          bestScore = test[1];
          best = index;
        }
      });
    });
    return best;
  }

  /** Guess the important columns from their names. All indexes are 0-based. */
  function detectColumns(headers) {
    return {
      date: bestColumn(headers, [
        [/\b(date|day|month|year|period)\b/, 100],
        [/\b(created|ordered|shipped|invoice)\b/, 60],
      ]),
      units: bestColumn(headers, [
        [/\b(units?|qty|quantity|pieces?|pcs)\b/, 100],
        [/\b(count|number|vol(ume)?)\b/, 60],
      ]),
      amount: bestColumn(headers, [
        [/\b(amount|total|revenue|proceeds)\b/, 100],
        [/\b(sales|turnover|spend|profit)\b/, 90],
        [/\b(cost|value)\b/, 60],
        [/\b(price|rate)\b/, 40],
      ]),
      region: bestColumn(headers, [
        [/\b(region|territory|area|zone|market)\b/, 100],
        [/\b(country|state|province|city)\b/, 80],
      ]),
      people: bestColumn(headers, [
        [/\b(sales\s*rep|salesperson|salesman|rep)\b/, 100],
        [/\b(owner|manager|employee|agent|associate|person|contact)\b/, 70],
        [/\b(customer|client)\b/, 50],
      ]),
      item: bestColumn(headers, [
        [/\b(item|product|sku|model|service)\b/, 100],
        [/\b(category|description|name)\b/, 60],
      ]),
    };
  }

  /** Quote a table name for a structured reference when it contains spaces. */
  function quoteTableName(name) {
    var text = String(name || "");
    return /^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(text) ? text : "'" + text.replace(/'/g, "''") + "'";
  }

  /** Build a structured reference such as tblSales[Sale Amount]. */
  function structuredRef(tableName, columnName) {
    return quoteTableName(tableName) + "[" + String(columnName == null ? "" : columnName) + "]";
  }

  /** Pick tblSales, or tblSales2/3/... when the name is already taken. */
  function uniqueTableName(existingNames, preferred) {
    var used = {};
    (existingNames || []).forEach(function (name) {
      used[String(name).toLowerCase()] = true;
    });
    var base = String(preferred || "tblSales");
    if (!used[base.toLowerCase()]) return base;
    var suffix = 2;
    while (used[(base + suffix).toLowerCase()]) suffix++;
    return base + suffix;
  }

  // Explicit, user-requested bounded reads. Names with spaces use Excel's
  // quoted-sheet syntax, e.g. @range('Sales Data'!A1:B20).
  function parseRequestedRanges(text) {
    var requests = [], re = /@range\(([^()]*)\)/gi, match, total = 0;
    while ((match = re.exec(String(text || "")))) {
      if (requests.length >= 3) throw new Error("At most 3 @range requests are allowed per message.");
      var parsed = match[1].trim().match(/^(?:'((?:[^']|'')+)'|([A-Za-z0-9_ .-]+))!\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?$/);
      if (!parsed) throw new Error("Invalid @range. Use @range(Sheet1!A1:B10) with an exact sheet and finite A1 range.");
      var sheet = parsed[1] != null ? parsed[1].replace(/''/g, "'") : parsed[2];
      var colIndex = function (name) { var n = 0; for (var i = 0; i < name.length; i++) n = n * 26 + name.toUpperCase().charCodeAt(i) - 64; return n; };
      var c1 = colIndex(parsed[3]), r1 = Number(parsed[4]), c2 = colIndex(parsed[5] || parsed[3]), r2 = Number(parsed[6] || parsed[4]);
      if (r1 < 1 || r2 < r1 || r2 > 1048576 || c1 < 1 || c2 < c1 || c2 > 16384) throw new Error("Invalid or out-of-bounds @range.");
      var count = (r2-r1+1)*(c2-c1+1); total += count;
      if (total > 500) throw new Error("Explicit @range reads are limited to 500 cells total per message.");
      requests.push({ sheet: sheet, range: parsed[3].toUpperCase()+r1+":"+(parsed[5] || parsed[3]).toUpperCase()+r2, cellCount: count });
    }
    var expected = (String(text || "").match(/@range\(/gi) || []).length;
    if (expected !== requests.length) throw new Error("Invalid @range syntax.");
    return requests;
  }

  function parseAgentReadRequest(text) {
    var trimmed = String(text || "").trim();
    var fence = trimmed.match(/^```(?:json)?\s*\n([\s\S]*?)\n?```$/i);
    var payload;
    try { payload = JSON.parse(fence ? fence[1] : trimmed); } catch { return null; }
    if (!payload || !Object.prototype.hasOwnProperty.call(payload, "read_ranges")) return null;
    if (!Array.isArray(payload.read_ranges) || payload.read_ranges.length < 1 || payload.read_ranges.length > 3 || payload.operations != null) throw new Error("Invalid model read request.");
    var tokens = payload.read_ranges.map(function (item) {
      if (!item || typeof item.sheet !== "string" || typeof item.range !== "string" || !item.sheet || item.sheet.length > 31 || /[\[\]:*?\\/]/.test(item.sheet) || !/^\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?$/.test(item.range)) throw new Error("Invalid model read request: use an exact sheet and finite A1 range.");
      return "@range('" + item.sheet.replace(/'/g, "''") + "'!" + item.range + ")";
    }).join(" ");
    var requests = parseRequestedRanges(tokens);
    return { tokens: tokens, ranges: requests, reason: typeof payload.reason === "string" ? payload.reason.slice(0, 300) : "" };
  }

  // Only explicit user-requested direct cell relationships are sent to the
  // model. Excel 1.12+ / 1.13+ availability varies by host.
  function parseRequestedTraces(text) {
    var count = (String(text || "").match(/@trace\(/gi) || []).length;
    if (!count) return [];
    if (count > 2) throw new Error("At most 2 @trace requests are allowed per message.");
    var matches = String(text).match(/@trace\([^()]*\)/gi) || [];
    if (matches.length !== count) throw new Error("Invalid @trace syntax.");
    return matches.map(function (token) {
      var item = parseRequestedRanges(token.replace(/^@trace/i, "@range"))[0];
      if (!item || item.cellCount !== 1) throw new Error("@trace requires exactly one cell, e.g. @trace(Sheet1!C5).");
      return item;
    });
  }

  async function traceDirectRelations(Excel, request, kind, supported) {
    if (!supported) return { status: "unsupported", addresses: [] };
    var phase = "cell";
    try {
      return await Excel.run(async function (ctx) {
        var cell = ctx.workbook.worksheets.getItem(request.sheet).getRange(request.range);
        cell.load("address");
        await ctx.sync();
        phase = "relations";
        var fn = kind === "precedents" ? cell.getDirectPrecedents : cell.getDirectDependents;
        if (typeof fn !== "function") return { status: "unsupported", addresses: [] };
        var areas = fn.call(cell);
        areas.load("addresses");
        await ctx.sync();
        var addresses = Array.isArray(areas.addresses) ? areas.addresses : [];
        var clipped = addresses.slice(0, 50).map(function (address) { return String(address).slice(0, 500); });
        return { status: "ok", addresses: clipped, truncated: addresses.length > clipped.length || addresses.some(function (address) { return String(address).length > 500; }) };
      });
    } catch (err) {
      if (phase === "relations" && err && String(err.code || "").toLowerCase() === "itemnotfound") return { status: "none", addresses: [] };
      throw err;
    }
  }

  function parseWorkbookSearch(text) {
    var count = (String(text || "").match(/@search\(/gi) || []).length;
    if (!count) return null;
    var match = String(text).match(/@search\("([^"\r\n]{1,100})"\)/i);
    if (count !== 1 || !match || !match[1].trim()) throw new Error('Use one @search("text") request with a non-empty term of at most 100 characters.');
    return match[1];
  }

  async function searchWorkbook(Excel, term) {
    return Excel.run(async function (ctx) {
      var sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      await ctx.sync();
      if (sheets.items.length > 20) throw new Error("Workbook search is limited to 20 sheets; use explicit @range reads instead.");
      var ranges = sheets.items.map(function (sheet) {
        var nullable = typeof sheet.getUsedRangeOrNullObject === "function";
        var range = nullable ? sheet.getUsedRangeOrNullObject() : sheet.getUsedRange();
        if (nullable) range.load("isNullObject");
        return { sheet: sheet, range: range, nullable: nullable };
      });
      await ctx.sync();
      ranges = ranges.filter(function (entry) { return !entry.nullable || !entry.range.isNullObject; });
      ranges.forEach(function (entry) { entry.range.load(["rowCount", "columnCount", "rowIndex", "columnIndex"]); });
      await ctx.sync();
      var cells = ranges.reduce(function (n, entry) { return n + entry.range.rowCount * entry.range.columnCount; }, 0);
      if (cells > 5000) throw new Error("Workbook used ranges exceed the 5,000-cell search limit; use explicit @range reads instead.");
      ranges.forEach(function (entry) { entry.range.load(["values", "formulas"]); });
      await ctx.sync();
      var matches = [], total = 0, needle = term.toLowerCase();
      function col(n) { var out = ""; for (n++; n > 0; n = Math.floor((n - 1) / 26)) out = String.fromCharCode(65 + (n - 1) % 26) + out; return out; }
      ranges.forEach(function (entry) {
        var range = entry.range;
        for (var r = 0; r < range.rowCount; r++) for (var c = 0; c < range.columnCount; c++) {
          var value = range.values[r][c], formula = range.formulas[r][c];
          var isFormula = typeof formula === "string" && formula.charAt(0) === "=";
          if (String(value == null ? "" : value).toLowerCase().indexOf(needle) < 0 && (!isFormula || formula.toLowerCase().indexOf(needle) < 0)) continue;
          total++;
          if (matches.length < 50) matches.push({sheet: entry.sheet.name, cell: col(range.columnIndex + c) + (range.rowIndex + r + 1), value: String(value == null ? "" : value).slice(0, 100), formula: isFormula ? formula.slice(0, 100) : ""});
        }
      });
      return { term: term, searchedCells: cells, matches: matches, totalMatches: total, truncated: total > matches.length, note: "Literal case-insensitive search of values and formulas in used ranges; matches may have truncated cell text." };
    });
  }

  async function readWorkbookBackup(office, maxBytes) {
    if (!office || !office.context || !office.context.document || typeof office.context.document.getFileAsync !== "function") throw new Error("Workbook export is unavailable in this host.");
    function resultPromise(start) {
      return new Promise(function (resolve, reject) {
        start(function (result) {
          if (result && String(result.status).toLowerCase() === "succeeded") resolve(result.value);
          else reject(new Error(result && result.error && result.error.message || "Office workbook export failed."));
        });
      });
    }
    var file = await resultPromise(function (cb) { office.context.document.getFileAsync(office.FileType.Compressed, { sliceSize: 65536 }, cb); });
    try {
      if (!file || !Number.isFinite(file.size) || file.size <= 0 || file.size > maxBytes || !Number.isInteger(file.sliceCount) || file.sliceCount < 1) throw new Error("Workbook backup is empty or exceeds the 50 MB limit.");
      var chunks = [], received = 0;
      for (var i = 0; i < file.sliceCount; i++) {
        var slice = await resultPromise(function (cb) { file.getSliceAsync(i, cb); });
        var bytes = new Uint8Array(slice.data);
        received += bytes.length;
        if (received > maxBytes) throw new Error("Workbook backup exceeds the 50 MB limit.");
        chunks.push(bytes);
      }
      if (received !== file.size) throw new Error("Incomplete workbook backup; no file was prepared.");
      return chunks;
    } finally {
      if (file && typeof file.closeAsync === "function") file.closeAsync(function () {});
    }
  }

  var api = {
    parseRequestedRanges: parseRequestedRanges,
    parseRequestedTraces: parseRequestedTraces,
    parseAgentReadRequest: parseAgentReadRequest,
    traceDirectRelations: traceDirectRelations,
    readWorkbookBackup: readWorkbookBackup,
    parseWorkbookSearch: parseWorkbookSearch,
    searchWorkbook: searchWorkbook,
    escapeHtml: escapeHtml,
    escapeAttr: escapeAttr,
    colName: colName,
    toA1: toA1,
    renderMarkdown: renderMarkdown,
    extractTable: extractTable,
    singleFormula: singleFormula,
    belowAddress: belowAddress,
    startAddress: startAddress,
    cleanHeader: cleanHeader,
    cleanHeaders: cleanHeaders,
    detectColumns: detectColumns,
    quoteTableName: quoteTableName,
    structuredRef: structuredRef,
    uniqueTableName: uniqueTableName,
  };

  root.OpenCodeExcel = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
