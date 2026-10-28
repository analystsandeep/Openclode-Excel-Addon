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

  var api = {
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
