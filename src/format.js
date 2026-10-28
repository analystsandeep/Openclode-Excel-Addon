/* OpenCode for Excel - native sheet formatter.
 *
 * This module does the formatting inside Excel with Office.js. It is deliberately
 * independent of the DOM and of the model: the pane calls formatSheet() when the
 * user clicks "Format sheet". The same code can therefore be tested with a fake
 * Excel object model.
 *
 * Exposes window.OpenCodeExcelFormat.formatSheet(Excel, lib, options).
 */
(function (root) {
  "use strict";

  var COLORS = {
    navy: "#1F3864",
    blue: "#2F5597",
    lightBlue: "#D9E2F3",
    lightGray: "#F2F2F2",
    white: "#FFFFFF",
    barBlue: "#5B9BD5",
    redFill: "#FFC7CE",
    redText: "#9C0006",
    grayLine: "#BFBFBF",
  };

  var CURRENCY = "$#,##0;[Red]-$#,##0";
  var TABLE_STYLE = "TableStyleMedium2";

  function apiSupported(version) {
    try {
      if (typeof Office === "undefined" || !Office.context || !Office.context.requirements) return true;
      return Office.context.requirements.isSetSupported("ExcelApi", version);
    } catch (err) {
      return true;
    }
  }

  function headerValue(values, rowOffset, colOffset) {
    var row = values && values[rowOffset];
    if (!row) return "";
    var value = row[colOffset];
    return value == null ? "" : String(value).trim();
  }

  /** Remove trailing all-empty rows (formulas count as content). */
  async function trimDataRows(ctx, sheet, firstColLetter, lastColLetter, firstCol, lastCol, headerRow, lastRow) {
    var dataRows = lastRow - headerRow;
    if (dataRows <= 0) return lastRow;
    var cells = dataRows * (lastCol - firstCol + 1);
    if (cells > 20000) return lastRow;

    var range = sheet.getRange(firstColLetter + (headerRow + 1) + ":" + lastColLetter + lastRow);
    range.load("formulas");
    await ctx.sync();
    var rows = range.formulas || [];
    for (var r = rows.length - 1; r >= 0; r--) {
      var row = rows[r] || [];
      var filled = row.some(function (cell) {
        return cell != null && String(cell).trim() !== "";
      });
      if (filled) return headerRow + 1 + r;
    }
    return headerRow;
  }

  function distinctFormula(ref) {
    return '=SUMPRODUCT((' + ref + '<>"")/COUNTIF(' + ref + ',' + ref + '&""))';
  }

  function topRegionFormula(regionRef, amountRef) {
    // INDEX(...,0) forces an array result without requiring Ctrl+Shift+Enter,
    // which matters on Excel 2021 builds where LET/UNIQUE may be unavailable.
    return (
      "=INDEX(" + regionRef + ",MATCH(MAX(INDEX(SUMIF(" + regionRef + "," + regionRef + "," + amountRef + "),0))," +
      "INDEX(SUMIF(" + regionRef + "," + regionRef + "," + amountRef + "),0),0))"
    );
  }

  function summaryPairs(tableName, headers, cols, lib) {
    var ref = function (index) {
      return lib.structuredRef(tableName, headers[index]);
    };
    var hasDate = cols.date >= 0;
    var hasUnits = cols.units >= 0;
    var hasAmount = cols.amount >= 0;
    var hasRegion = cols.region >= 0;

    return [
      {
        label: "Period",
        formula: hasDate ? '=TEXT(MIN(' + ref(cols.date) + '),"mmm yyyy")&" - "&TEXT(MAX(' + ref(cols.date) + '),"mmm yyyy")' : null,
        value: "All data",
        format: null,
        align: "Left",
      },
      {
        label: "Orders",
        formula: "=COUNTA(" + ref(hasDate ? cols.date : 0) + ")",
        value: "—",
        format: "#,##0",
        align: "Right",
      },
      {
        label: "Units",
        formula: hasUnits ? "=SUM(" + ref(cols.units) + ")" : null,
        value: "—",
        format: "#,##0",
        align: "Right",
      },
      {
        label: "Revenue",
        formula: hasAmount ? "=SUM(" + ref(cols.amount) + ")" : null,
        value: "—",
        format: CURRENCY,
        align: "Right",
      },
      {
        label: "Avg order",
        formula: hasAmount ? "=IFERROR(AVERAGE(" + ref(cols.amount) + "),0)" : null,
        value: "—",
        format: CURRENCY,
        align: "Right",
      },
      {
        label: "Top region",
        formula: hasRegion && hasAmount ? topRegionFormula(ref(cols.region), ref(cols.amount)) : null,
        value: "—",
        format: null,
        align: "Left",
      },
      {
        label: "People",
        formula: cols.people >= 0 ? distinctFormula(ref(cols.people)) : null,
        value: "—",
        format: "#,##0",
        align: "Right",
      },
      {
        label: "Items",
        formula: cols.item >= 0 ? distinctFormula(ref(cols.item)) : null,
        value: "—",
        format: "#,##0",
        align: "Right",
      },
    ];
  }

  function writeBorder(range, edge, weight, color) {
    var border = range.format.borders.getItem(edge);
    border.style = "Continuous";
    border.weight = weight;
    border.color = color;
  }

  /**
   * Format the active worksheet: clean headers, create/restyle a table, add a
   * three-row summary strip, set number formats, freeze panes, print setup and
   * conditional formatting. Safe to run repeatedly: an existing tblSales (or the
   * only table on the sheet) is restyled instead of duplicated.
   *
   * @param {object} Excel  The global Excel namespace from Office.js.
   * @param {object} lib    The OpenCodeExcel helpers (colName, detectColumns, ...).
   * @param {object} [options] Optional { onWarning(message) }.
   * @returns {Promise<object>} A small summary of what was formatted.
   */
  async function formatSheet(Excel, lib, options) {
    if (!Excel || !lib) throw new Error("The formatter is not connected to Excel.");
    options = options || {};
    var warnings = [];
    function warn(message) {
      warnings.push(message);
      if (options.onWarning) options.onWarning(message);
    }

    var info = await Excel.run(async function (ctx) {
      var sheet = ctx.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      var used = sheet.getUsedRange();
      used.load(["rowCount", "columnCount", "rowIndex", "columnIndex"]);
      var tables = sheet.tables;
      tables.load("items/name");
      await ctx.sync();

      var sheetName = sheet.name || "Sheet";
      if (used.rowCount < 2) throw new Error("This sheet has no data rows to format.");

      var firstCol = used.columnIndex + 1;
      var lastCol = used.columnIndex + used.columnCount;
      var firstColLetter = lib.colName(firstCol - 1);
      var lastColLetter = lib.colName(lastCol - 1);
      var tableNames = (tables.items || []).map(function (table) {
        return table.name;
      });

      var existing = (tables.items || []).filter(function (table) {
        return String(table.name || "").toLowerCase() === "tblsales";
      })[0];
      if (!existing && tables.items && tables.items.length === 1) existing = tables.items[0];

      var table = null;
      var tableName = null;
      var headerRow = 0;
      var lastRow = 0;
      var summaryTop = 0;
      var created = false;
      var headers = [];

      if (existing) {
        table = existing;
        tableName = table.name;
        var tableRange = table.getRange();
        tableRange.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
        await ctx.sync();

        headerRow = tableRange.rowIndex + 1;
        firstCol = tableRange.columnIndex + 1;
        lastCol = tableRange.columnIndex + tableRange.columnCount;
        lastRow = tableRange.rowIndex + tableRange.rowCount;
        firstColLetter = lib.colName(firstCol - 1);
        lastColLetter = lib.colName(lastCol - 1);

        var headerRange = sheet.getRange(firstColLetter + headerRow + ":" + lastColLetter + headerRow);
        headerRange.load("values");
        await ctx.sync();
        headers = ((headerRange.values && headerRange.values[0]) || []).map(function (value, index) {
          var text = value == null ? "" : String(value).trim();
          return text || "Column " + (index + 1);
        });

        // Only manage the summary when the sheet already has our layout.
        if (headerRow > 3) {
          var titleCell = sheet.getRange(firstColLetter + (headerRow - 3));
          titleCell.load("values");
          await ctx.sync();
          var title = titleCell.values && titleCell.values[0] && titleCell.values[0][0];
          if (String(title || "").trim().toUpperCase() === "SALES SUMMARY") summaryTop = headerRow - 3;
        }
      } else {
        var wantSummary = used.columnCount >= 2;
        var headerStart = used.rowIndex + 1;
        var probeLastRow = used.rowIndex + used.rowCount;

        var probe = sheet.getRange(
          firstColLetter + headerStart + ":" + lastColLetter + Math.min(probeLastRow, headerStart + 4)
        );
        probe.load("values");
        await ctx.sync();
        var firstCell = headerValue(probe.values, 0, 0);
        var hasSummary = wantSummary && firstCell.toUpperCase() === "SALES SUMMARY";

        if (hasSummary) {
          headerRow = 4;
          summaryTop = 1;
          lastRow = probeLastRow;

          var existingHeader = sheet.getRange(firstColLetter + headerRow + ":" + lastColLetter + headerRow);
          existingHeader.load("values");
          await ctx.sync();
          var rawExisting = (existingHeader.values && existingHeader.values[0]) || [];
          if (rawExisting.filter(function (value) { return value != null && String(value).trim() !== ""; }).length < 2) {
            throw new Error("The top of this sheet looks like a summary, but row 4 is not a table header.");
          }
          headers = lib.cleanHeaders(rawExisting);
          existingHeader.values = [headers];
        } else {
          headerRow = headerStart + (wantSummary ? 3 : 0);
          summaryTop = wantSummary ? headerRow - 3 : 0;
          lastRow = probeLastRow + (wantSummary ? 3 : 0);

          if (wantSummary) {
            sheet.getRange(headerStart + ":" + (headerStart + 2)).insert("Down");
            await ctx.sync();
          }

          var headerRangeNew = sheet.getRange(firstColLetter + headerRow + ":" + lastColLetter + headerRow);
          headerRangeNew.load("values");
          await ctx.sync();
          headers = lib.cleanHeaders((headerRangeNew.values && headerRangeNew.values[0]) || []);
          headerRangeNew.values = [headers];
        }

        lastRow = await trimDataRows(
          ctx, sheet, firstColLetter, lastColLetter, firstCol, lastCol, headerRow, lastRow
        );
        if (lastRow <= headerRow) throw new Error("This sheet has a header row but no data rows.");

        tableName = lib.uniqueTableName(tableNames, "tblSales");
        table = sheet.tables.add(
          sheet.getRange(firstColLetter + headerRow + ":" + lastColLetter + lastRow),
          true
        );
        table.name = tableName;
        created = true;
      }

      table.style = TABLE_STYLE;
      table.showBandedRows = true;
      table.showBandedColumns = false;
      table.showFilterButton = true;

      /* header row */
      var header = sheet.getRange(firstColLetter + headerRow + ":" + lastColLetter + headerRow);
      header.format.fill.color = COLORS.navy;
      header.format.font.color = COLORS.white;
      header.format.font.bold = true;
      header.format.font.size = 10;
      header.format.horizontalAlignment = "Center";
      header.format.verticalAlignment = "Center";
      header.format.wrapText = true;
      header.format.rowHeight = 26;
      writeBorder(header, "EdgeBottom", "Thick", COLORS.navy);

      /* data body */
      var body = sheet.getRange(firstColLetter + (headerRow + 1) + ":" + lastColLetter + lastRow);
      body.format.verticalAlignment = "Center";
      body.format.font.size = 10;
      body.format.rowHeight = 18;

      var cols = lib.detectColumns(headers);
      for (var i = 0; i < headers.length; i++) {
        var colLetter = lib.colName(firstCol - 1 + i);
        var columnBody = sheet.getRange(colLetter + (headerRow + 1) + ":" + colLetter + lastRow);

        if (i === cols.date) {
          columnBody.numberFormat = "dd-mmm-yyyy";
          columnBody.format.horizontalAlignment = "Center";
        } else if (i === cols.units) {
          columnBody.numberFormat = "#,##0";
          columnBody.format.horizontalAlignment = "Right";
        } else if (i === cols.amount) {
          columnBody.numberFormat = CURRENCY;
          columnBody.format.horizontalAlignment = "Right";
        } else {
          columnBody.format.horizontalAlignment = "Left";
        }
      }

      /* summary strip */
      var summary = null;
      if (summaryTop > 0) {
        var pairs = summaryPairs(tableName, headers, cols, lib);
        var maxPairs = Math.min(4, Math.floor((lastCol - firstCol + 1) / 2));

        var titleRange = sheet.getRange(firstColLetter + summaryTop + ":" + lastColLetter + summaryTop);
        titleRange.format.fill.color = COLORS.navy;
        titleRange.format.font.color = COLORS.white;
        titleRange.format.font.bold = true;
        titleRange.format.font.size = 14;
        titleRange.format.horizontalAlignment = "Left";
        titleRange.format.verticalAlignment = "Center";
        titleRange.format.indentLevel = 1;
        titleRange.format.rowHeight = 28;
        sheet.getRange(firstColLetter + summaryTop).values = [["SALES SUMMARY"]];

        sheet.getRange(firstColLetter + (summaryTop + 1) + ":" + lastColLetter + (summaryTop + 1)).format.rowHeight = 24;
        sheet.getRange(firstColLetter + (summaryTop + 2) + ":" + lastColLetter + (summaryTop + 2)).format.rowHeight = 22;

        var row2 = sheet.getRange(firstColLetter + (summaryTop + 1) + ":" + lastColLetter + (summaryTop + 1));
        writeBorder(row2, "EdgeBottom", "Medium", COLORS.navy);
        var row3 = sheet.getRange(firstColLetter + (summaryTop + 2) + ":" + lastColLetter + (summaryTop + 2));
        writeBorder(row3, "EdgeBottom", "Thin", COLORS.grayLine);

        for (var p = 0; p < pairs.length; p++) {
          var pair = pairs[p];
          var pairRow;
          var pairCol;
          if (p < maxPairs) {
            pairRow = summaryTop + 1;
            pairCol = firstCol + p * 2;
          } else if (p < maxPairs * 2) {
            pairRow = summaryTop + 2;
            pairCol = firstCol + (p - maxPairs) * 2;
          } else {
            continue;
          }

          var labelLetter = lib.colName(pairCol - 1);
          var valueLetter = lib.colName(pairCol);
          var labelCell = sheet.getRange(labelLetter + pairRow);
          var valueCell = sheet.getRange(valueLetter + pairRow);

          labelCell.values = [[pair.label]];
          labelCell.format.font.bold = true;
          labelCell.format.font.size = 10;
          labelCell.format.horizontalAlignment = "Center";
          labelCell.format.verticalAlignment = "Center";

          if (pairRow === summaryTop + 1) {
            labelCell.format.fill.color = COLORS.blue;
            labelCell.format.font.color = COLORS.white;
            valueCell.format.fill.color = COLORS.lightBlue;
            valueCell.format.font.size = 12;
          } else {
            labelCell.format.fill.color = COLORS.lightBlue;
            labelCell.format.font.color = COLORS.navy;
            valueCell.format.fill.color = COLORS.lightGray;
            valueCell.format.font.size = 11;
          }

          valueCell.format.font.bold = true;
          valueCell.format.verticalAlignment = "Center";
          valueCell.format.horizontalAlignment = pair.align;
          if (pair.format) valueCell.numberFormat = pair.format;
          if (pair.formula) {
            // Clear first: D3 may still hold an old single-cell array formula,
            // and Excel refuses to overwrite part of an array in one step.
            valueCell.clear("Contents");
            valueCell.formulas = [[pair.formula]];
          } else {
            valueCell.values = [[pair.value]];
          }
        }

        summary = {
          row: summaryTop,
          labels: pairs.slice(0, maxPairs * 2).map(function (pair) { return pair.label; }),
        };
      }

      // Let Excel choose the widths: the Office.js columnWidth unit is not the
      // character unit shown in the Excel UI, so fixed numbers come out far too
      // narrow. Autofit also accounts for the summary strip and currency values.
      sheet.getRange(firstColLetter + "1:" + lastColLetter + lastRow).format.autofitColumns();

      return {
        sheetName: sheetName,
        tableName: tableName,
        headerRow: headerRow,
        lastRow: lastRow,
        firstCol: firstCol,
        lastCol: lastCol,
        summaryTop: summaryTop,
        summary: summary,
        created: created,
        headers: headers,
        cols: cols,
      };
    });

    var firstColLetter = lib.colName(info.firstCol - 1);
    var lastColLetter = lib.colName(info.lastCol - 1);

    /* freeze the summary + header, hide gridlines, colour the tab */
    try {
      await Excel.run(async function (ctx) {
        var sheet = ctx.workbook.worksheets.getActiveWorksheet();
        sheet.freezePanes.freezeRows(info.headerRow);
        sheet.showGridlines = false;
        sheet.tabColor = COLORS.navy;
        await ctx.sync();
      });
    } catch (err) {
      warn("view settings: " + err.message);
    }

    /* print setup (ExcelApi 1.9) */
    if (apiSupported("1.9")) {
      var printTop = info.summaryTop || info.headerRow;
      try {
        await Excel.run(async function (ctx) {
          var sheet = ctx.workbook.worksheets.getActiveWorksheet();
          var layout = sheet.pageLayout;
          layout.orientation = "Landscape";
          layout.setPrintArea(firstColLetter + printTop + ":" + lastColLetter + info.lastRow);
          layout.setPrintTitleRows("$" + printTop + ":$" + info.headerRow);
          layout.setPrintMargins("Inches", {
            top: 0.75,
            bottom: 0.75,
            left: 0.25,
            right: 0.25,
            header: 0.3,
            footer: 0.3,
          });
          layout.centerHorizontally = true;
          layout.printGridlines = false;
          await ctx.sync();
        });
      } catch (err) {
        warn("print setup: " + err.message);
      }

      try {
        await Excel.run(async function (ctx) {
          var sheet = ctx.workbook.worksheets.getActiveWorksheet();
          sheet.pageLayout.zoom = { horizontalFitToPages: 1, verticalFitToPages: null };
          await ctx.sync();
        });
      } catch (err) {
        warn("fit-to-width: " + err.message);
      }
    } else {
      warn("print setup needs ExcelApi 1.9");
    }

    /* merge the title across the summary strip on a first run */
    if (info.created && info.summaryTop > 0) {
      try {
        await Excel.run(async function (ctx) {
          var sheet = ctx.workbook.worksheets.getActiveWorksheet();
          var title = sheet.getRange(firstColLetter + info.summaryTop + ":" + lastColLetter + info.summaryTop);
          title.merge();
          await ctx.sync();
        });
      } catch (err) {
        warn("title merge: " + err.message);
      }
    }

    /* conditional formatting: only add it when the body has none */
    try {
      await Excel.run(async function (ctx) {
        var sheet = ctx.workbook.worksheets.getActiveWorksheet();
        var body = sheet.getRange(firstColLetter + (info.headerRow + 1) + ":" + lastColLetter + info.lastRow);
        var count = body.conditionalFormats.getCount();
        await ctx.sync();
        if (count.value > 0) return;

        var amountIndex = info.cols.amount >= 0 ? info.cols.amount : 0;
        var amountLetter = lib.colName(info.firstCol - 1 + amountIndex);
        var amountBody = sheet.getRange(amountLetter + (info.headerRow + 1) + ":" + amountLetter + info.lastRow);

        var bar = amountBody.conditionalFormats.add("DataBar");
        bar.dataBar.barColor = COLORS.barBlue;
        bar.dataBar.showValue = true;

        var custom = body.conditionalFormats.add("Custom");
        custom.custom.rule.formula =
          '=AND($' + amountLetter + (info.headerRow + 1) + '<>"",$' + amountLetter + (info.headerRow + 1) + "<0)";
        custom.custom.format.fill.color = COLORS.redFill;
        custom.custom.format.font.color = COLORS.redText;
        await ctx.sync();
      });
    } catch (err) {
      warn("conditional formatting: " + err.message);
    }

    return {
      sheetName: info.sheetName,
      tableName: info.tableName,
      headerRow: info.headerRow,
      lastRow: info.lastRow,
      dataRows: info.lastRow - info.headerRow,
      summaryTop: info.summaryTop,
      created: info.created,
      warnings: warnings,
    };
  }

  root.OpenCodeExcelFormat = {
    formatSheet: formatSheet,
    colors: COLORS,
    tableStyle: TABLE_STYLE,
  };
})(typeof globalThis !== "undefined" ? globalThis : this);
