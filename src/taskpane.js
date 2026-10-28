/* OpenCode for Excel - task pane logic (agentic edition).
 *
 * Talks only to the local bridge (same origin) using the token the bridge
 * injects into this page. The bridge discovers the OpenCode service and streams
 * events back over SSE.
 *
 * Agentic model: when the user asks to CHANGE the workbook, OpenCode is asked to
 * reply with a JSON edit plan (see src/plan.js). The pane shows a change preview
 * with cell-level references, the user approves individual operations, and they
 * are executed through Office.js. The pane keeps bounded before-snapshots for
 * cell-only operations so it can offer an "Undo last change" button. Structural
 * and object operations are not presented as add-in-undoable; use Excel's own
 * Undo or a backup for those.
 */

(function () {
  "use strict";

  var state = {
    sessionID: null,
    connected: false,
    streaming: false,
    context: null,
    bubble: null, // streaming assistant element
    answer: "",
    source: null,
    lastFinalized: null,
    lastError: null,
    doneTimer: null,
    modelCatalog: [],
    defaultModel: null,
    currentPromptRequest: null,
    fallbackAttempted: false,

    /* agentic state */
    undoStack: [],        /* { snap, ops, at, summary } */
    currentPlan: null,    /* parsed plan awaiting review */
    planJson: null,       /* raw JSON text of the current plan */
    readOnly: false,
    approvalMode: "always", /* always | auto | never */
    attachments: [],      /* { id, uploadID, uri, name, type, size, text } */
    mentionedSheets: [],  /* @SheetName focus from the last prompt */
    wantsPlan: false,     /* nudge the model toward a plan this turn */
  };

  var MODEL_KEY = "opencode.excel.model";
  var SESSION_KEY = "opencode.excel.session";
  var READONLY_KEY = "opencode.excel.readonly";
  var APPROVAL_KEY = "opencode.excel.approval";

  var MAX_ATTACHMENTS = 20;
  var MAX_ATTACHMENT_BYTES = 25 * 1024 * 1024;
  var ATTACH_TEXT_CAP = 6000;
  var WORKBOOK_ROW_CAP = 200;
  var WORKBOOK_COL_CAP = 60;
  var CONTEXT_CHAR_BUDGET = 30000;

  var PLAN_HINT =
    "If this task changes the workbook, reply with a JSON plan only (a ```json fence) using the schema in the " +
    "system instructions. Name the exact sheet and range for every operation so the user can review each change " +
    "cell by cell.";
  var TEXT_HINT =
    "If this task asks to change the workbook, return the reviewable JSON plan described above. Otherwise answer " +
    "in plain text and cite the cells you use (e.g. Sheet1!D12). If the answer is tabular, include a fenced " +
    "```tsv block with a header row.";
  var SUMMARY_HINT =
    "This is a data-summary request. Analyze the row-level records, not just the existing summary cells. " +
    "Start with one short headline covering the scope, period, and main KPI. Then give 4 to 6 decision-useful " +
    "insights: the strongest meaningful category or region breakdown, the leading product or contributor, a " +
    "time trend or period comparison when dates are available, and any material outlier. Compute and check the " +
    "figures from the supplied values; use existing summary cells only to validate totals. Put a compact source " +
    "range citation at the end of each insight, such as (SaleData!A5:H47). Do not return a formula inventory, " +
    "do not explain formulas unless asked, and do not use a Markdown table unless the user explicitly requests one. " +
    "If the available rows are insufficient for a calculation, state that limitation instead of guessing.";

  var PREAMBLE =
    "You are an agentic spreadsheet assistant in the OpenCode for Excel sidebar. The user works in a real " +
    "workbook and sees everything you say. Be concise, practical, and audit-friendly. Work with LIVE Excel " +
    "formulas whenever proposing workbook changes (XLOOKUP, LET, SUMPRODUCT, SUMIFS, ...) instead of pasting " +
    "static values. For analysis requests, report conclusions and calculated results rather than listing formula " +
    "definitions unless the user asks for them. Cite every figure back to the exact cells you used, e.g. " +
    "Sheet1!D12 or tblSales[Amount]. " +
    "Never suggest editing files or running shell commands; you operate on spreadsheet data only.\n" +
    "\nWhen the user asks you to CHANGE the workbook (create sheets, write cells or formulas, insert or delete " +
    "rows or columns, format, add charts), reply with NOTHING except a JSON plan in a ```json fence:\n" +
    "{ \"summary\": \"...\", \"assumptions\": [\"...\"], \"operations\": [\n" +
    "  { \"op\": \"setValues\", \"sheet\": \"Sheet1\", \"range\": \"A1:D1\", \"values\": [[\"Date\",\"Amount\"]], \"description\": \"...\" },\n" +
    "  { \"op\": \"setFormulas\", \"sheet\": \"Sheet1\", \"range\": \"E2:E10\", \"formulaR1C1\": \"=RC[-3]*RC[-2]\", \"description\": \"...\" },\n" +
    "  { \"op\": \"insertRows\", \"sheet\": \"Sheet1\", \"range\": \"A5\", \"count\": 3, \"description\": \"...\" },\n" +
    "  { \"op\": \"addSheet\", \"name\": \"Budget\", \"description\": \"...\" }\n" +
    "] }\n" +
    "Supported ops: setValues (range + values as a 2D array), setFormulas (formula for one cell, formulas as a " +
    "2D array, or formulaR1C1 for a multi-cell fill), clearRange, " +
    "insertRows/deleteRows/insertColumns/deleteColumns (range, or at + count), addSheet (name), " +
    "renameSheet (from + to), deleteSheet (name), formatRange (range + format: {bold, italic, size, color, " +
    "fill, numberFormat, horizontalAlignment, verticalAlignment, wrapText, borders}), autofitRange, mergeRange, " +
    "unmergeRange, addTable/deleteTable, sortRange, freezePanes, addChart/deleteChart (chartType: " +
    "ColumnClustered|BarClustered|Line|LineMarkers|Pie|Doughnut|Area|Scatter, optional position).\n" +
    "Every operation must name the sheet and range it touches.\n" +
    "\nOtherwise answer in plain text; tabular replies should include a fenced ```tsv block with a header row.";

  var $ = function (id) { return document.getElementById(id); };
  var lib = (typeof OpenCodeExcel !== "undefined") ? OpenCodeExcel : {};
  var plan = (typeof OpenCodeExcelPlan !== "undefined") ? OpenCodeExcelPlan : null;
  var formatter = (typeof OpenCodeExcelFormat !== "undefined") ? OpenCodeExcelFormat : null;
  var escapeHtml = lib.escapeHtml || function (t) { return String(t); };
  var renderMarkdown = lib.renderMarkdown || function (t) { return escapeHtml(t); };
  var extractTable = lib.extractTable || function () { return null; };
  var singleFormula = lib.singleFormula || function () { return null; };
  var colName = lib.colName || function (i) { return String.fromCharCode(65 + i); };

  /* ------------------------------------------------------------- utilities */

  function copyText(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(
        function () { banner("Copied to clipboard", "ok"); },
        function () { fallbackCopy(text); }
      );
    } else {
      fallbackCopy(text);
    }
  }

  function fallbackCopy(text) {
    var ta = document.createElement("textarea");
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand("copy"); banner("Copied to clipboard", "ok"); }
    catch (err) { banner("Copy failed - select the text and copy manually", "err"); }
    document.body.removeChild(ta);
  }

  var bannerTimer = null;
  function banner(message, kind) {
    var el = $("banner");
    el.textContent = message;
    el.className = "banner" + (kind ? " " + kind : "");
    clearTimeout(bannerTimer);
    bannerTimer = setTimeout(function () { el.classList.add("hidden"); }, 6000);
  }

  function resizeComposer() {
    var input = $("input");
    if (!input) return;
    input.style.height = "auto";
    input.style.height = Math.min(150, Math.max(48, input.scrollHeight || 48)) + "px";
  }

  /* ---------------------------------------------------------------- bridge */

  function bridgeToken() {
    var meta = document.querySelector('meta[name="opencode-excel-bridge-token"]');
    return meta ? String(meta.getAttribute("content") || "") : "";
  }

  async function api(pathname, options) {
    options = options || {};
    var headers = Object.assign({
      "content-type": "application/json",
      "x-opencode-excel-token": bridgeToken()
    }, options.headers || {});
    var res = await fetch(pathname, Object.assign({}, options, { headers: headers }));
    var json = null;
    try { json = await res.json(); } catch (err) { /* non-JSON */ }
    if (!res.ok) throw new Error((json && json.error) || ("HTTP " + res.status));
    return json;
  }

  async function checkStatus() {
    try {
      var status = await api("/api/status");
      state.connected = !!status.connected;
      setStatus(status);
      return status;
    } catch (err) {
      state.connected = false;
      setStatus({ connected: false, hint: "Bridge not reachable. Is `npm run dev` running?" });
      return null;
    }
  }

  function setStatus(status) {
    var text = $("statusText");
    var dot = '<span class="dot ' + (status.connected ? "ok" : "err") + '"></span>';
    if (status.connected) {
      text.innerHTML = dot + "OpenCode " + (status.version ? "v" + status.version : "connected");
    } else {
      text.innerHTML = dot + "Not connected";
      if (status.hint) banner(status.hint, "err");
    }
  }

  async function ensureSession() {
    if (state.sessionID) { openStream(); return state.sessionID; }
    var data = await api("/api/session", { method: "POST", body: JSON.stringify({}) });
    state.sessionID = data.sessionID;
    try { localStorage.setItem(sessionStorageKey(), state.sessionID); } catch (err) { /* storage can be disabled inside Office */ }
    openStream();
    return data.sessionID;
  }

  function openStream() {
    if (state.source) state.source.close();
    if (!state.sessionID) return;
    state.source = new EventSource("/api/events?sessionID=" + encodeURIComponent(state.sessionID) +
      "&token=" + encodeURIComponent(bridgeToken()));
    state.source.onmessage = function (event) {
      var payload;
      try { payload = JSON.parse(event.data); } catch (err) { return; }
      handleEvent(payload);
    };
    state.source.onerror = function () { /* EventSource reconnects on its own */ };
  }

  /* ------------------------------------------------------------ chat render */

  function addMessage(role, html) {
    $("emptyState").classList.add("hidden");
    var wrap = document.createElement("div");
    wrap.className = "msg " + role;
    var who = document.createElement("div");
    who.className = "who";
    who.textContent = role === "user" ? "You" : "OpenCode";
    var bubble = document.createElement("div");
    bubble.className = "bubble";
    bubble.innerHTML = html;
    wrap.appendChild(who);
    wrap.appendChild(bubble);
    $("messages").appendChild(wrap);
    $("messages").scrollTop = $("messages").scrollHeight;
    return { wrap: wrap, bubble: bubble };
  }

  function beginAssistant() {
    hidePreview();
    hideUndoRow();
    var msg = addMessage("assistant", '<span class="thinking">Thinking…</span>');
    state.bubble = msg.bubble;
    state.answer = "";
    state.lastError = null;
    state.streaming = true;
    $("stopRow").classList.remove("hidden");
    $("activity").textContent = "Working…";
    $("btnSend").disabled = true;
  }

  function handleEvent(evt) {
    if (!state.bubble) return;
    if (evt.type === "delta") {
      state.answer += evt.delta || "";
      state.bubble.textContent = state.answer;
      scroll();
    } else if (evt.type === "text") {
      state.answer = evt.text || "";
      state.bubble.textContent = state.answer;
      scroll();
    } else if (evt.type === "tool") {
      $("activity").textContent = "Using tool: " + (evt.name || "tool");
    } else if (evt.type === "permission") {
      banner("OpenCode asked to use \"" + (evt.tool || "a tool") + "\" - blocked by the workspace permissions.", "err");
    } else if (evt.type === "retry") {
      state.lastError = evt.message || state.lastError;
      $("activity").textContent = "Retrying" + (evt.attempt ? " (attempt " + evt.attempt + ")" : "") + "…";
    } else if (evt.type === "error") {
      state.lastError = evt.message || "OpenCode reported an error";
      if (isRecoverableModelFailure("", state.lastError) && !state.fallbackAttempted) {
        $("activity").textContent = "Model unavailable — finding a free fallback…";
      } else {
        banner(state.lastError, "err");
      }
    } else if (evt.type === "done") {
      var doneText = evt.text != null ? evt.text : state.answer;
      var doneError = evt.error || (!doneText ? state.lastError : null);
      if (startFreeModelFallback(doneText, doneError)) return;
      if (!doneText && doneError) banner(errorMessage(doneError), "err");
      finalize(doneText, doneError);
    }
  }

  function errorMessage(error) {
    if (!error) return "";
    if (typeof error === "string") return error;
    return String(error.message || error.error || error.type || error);
  }

  function isRecoverableModelFailure(text, error) {
    if (text) {
      return /model\s+is\s+unavailable|upstream request failed|no reply was returned/i.test(errorMessage(error));
    }
    var message = errorMessage(error);
    return !message || /model\s+is\s+unavailable|model unavailable|upstream request failed|no reply was returned|produced no (?:text|output|reply)|finished without returning any (?:text|output|reply)/i.test(message);
  }

  function modelKey(model) {
    if (!model) return "";
    return String(model.providerID || "") + "/" + String(model.id || model.modelID || "");
  }

  function modelLabel(model) {
    return String((model && (model.name || model.id || model.modelID)) || "fallback model");
  }

  function chooseFreeFallback(failedModel) {
    var preferred = [
      "mimo-v2.6-flash-free",
      "deepseek-v4.1-flash-free",
      "glm-5.3-flash-free",
      "nemotron-3.5-lightning-free",
      "beam-free"
    ];
    var failedKey = modelKey(failedModel);
    var candidates = state.modelCatalog.filter(function (model) {
      return model && model.providerID === "opencode" && model.isFree === true &&
        modelKey(model) !== failedKey && model.status !== "disabled" && model.status !== "deprecated";
    });
    candidates.sort(function (a, b) {
      var ai = preferred.indexOf(a.id);
      var bi = preferred.indexOf(b.id);
      if (ai < 0) ai = preferred.length;
      if (bi < 0) bi = preferred.length;
      return ai - bi;
    });
    return candidates[0] || null;
  }

  function selectModelInUi(model) {
    var value = JSON.stringify([model.providerID, model.id]);
    var select = $("modelSelect");
    for (var i = 0; i < select.options.length; i++) {
      if (select.options[i].value === value) {
        select.value = value;
        rememberModel();
        return;
      }
    }
  }

  function startFreeModelFallback(text, error) {
    if (!state.streaming || state.fallbackAttempted || !state.currentPromptRequest ||
        !isRecoverableModelFailure(text, error)) return false;

    var failedModel = state.currentPromptRequest.model || state.defaultModel;
    var fallback = chooseFreeFallback(failedModel);
    if (!fallback) return false;

    state.fallbackAttempted = true;
    state.lastError = null;
    state.answer = "";
    state.bubble.innerHTML = '<span class="thinking">Retrying with ' + escapeHtml(modelLabel(fallback)) + '…</span>';
    $("activity").textContent = "Retrying with " + modelLabel(fallback) + "…";
    selectModelInUi(fallback);

    var failedLabel = modelLabel(failedModel);
    banner(failedLabel + " is unavailable. Retrying once with free " + modelLabel(fallback) + ".", "warn");
    audit("model-fallback", {
      from: failedModel || null,
      to: { providerID: fallback.providerID, id: fallback.id },
      error: errorMessage(error) || "empty response"
    });

    state.currentPromptRequest = Object.assign({}, state.currentPromptRequest, {
      model: { providerID: fallback.providerID, id: fallback.id }
    });
    submitPromptRequest(state.currentPromptRequest).catch(function (err) {
      banner("Free fallback failed: " + err.message, "err");
      if (state.bubble) finalize("", "Automatic fallback failed: " + err.message);
    });
    return true;
  }

  function scroll() {
    var m = $("messages");
    m.scrollTop = m.scrollHeight;
  }

  function finalize(text, error) {
    if (!state.bubble) return;
    clearTimeout(state.doneTimer);
    state.answer = text != null ? text : state.answer;
    var answerText = state.answer;
    var errorText = errorMessage(error) || null;
    var msg = state.bubble;

    if (answerText) {
      msg.innerHTML = renderMarkdown(answerText);
      if (errorText) {
        var note = document.createElement("div");
        note.className = "error-text";
        note.textContent = "Note: " + errorText;
        msg.appendChild(note);
      }
    } else {
      msg.innerHTML =
        '<div class="error-text">' +
        escapeHtml(errorText || "No reply was returned. Try again, or choose a different model in Settings.") +
        "</div>";
    }

    var wantsPlanThisTurn = state.wantsPlan;
    state.wantsPlan = false;

    var table = answerText ? extractTable(answerText) : null;
    var formula = answerText ? singleFormula(answerText) : null;

    /* Agentic path: OpenCode proposed a structured edit plan. */
    var parsedPlan = null;
    if (answerText && plan) {
      try { parsedPlan = plan.parsePlan(answerText); } catch (err) { parsedPlan = null; }
    }
    if (parsedPlan && parsedPlan.ok) {
      var note = document.createElement("div");
      note.className = "plan-summary";
      note.textContent = parsedPlan.summary + " — " + parsedPlan.operations.length + " operation(s) to review.";
      msg.appendChild(note);
      state.currentPlan = parsedPlan;
      state.planJson = JSON.stringify(parsedPlan.raw, null, 2);
      renderPlanPreview(parsedPlan, wantsPlanThisTurn);
      var autoLimit = state.approvalMode === "auto" ? 5 : plan.MAX_PREVIEW_CELLS;
      if (!state.readOnly && state.approvalMode !== "always" &&
          plan.canAutoApply(parsedPlan.operations, autoLimit)) {
        setTimeout(function () { applyCore(parsedPlan.operations); }, 0);
      }
    } else if (answerText) {
      var actions = document.createElement("div");
      actions.className = "actions";

      actions.appendChild(makeButton("Copy", function () { copyText(answerText); }));
      if (!state.readOnly) {
        actions.appendChild(makeButton("Insert ↓", function () { insertBelow(answerText); }));
        if (table) actions.appendChild(makeButton("Insert table (" + table.length + "×" + table[0].length + ")", function () { insertTable(table); }));
        if (formula) actions.appendChild(makeButton("Use as formula", function () { insertFormula(formula); }));
      }

      var parent = msg.parentNode;
      parent.appendChild(actions);

      var copyButtons = msg.querySelectorAll(".copy-code");
      for (var i = 0; i < copyButtons.length; i++) {
        copyButtons[i].addEventListener("click", function (e) {
          copyText(e.target.getAttribute("data-code") || "");
        });
      }
    }

    state.lastFinalized = { text: answerText, table: table, formula: formula };
    state.currentPromptRequest = null;
    state.fallbackAttempted = false;
    state.bubble = null;
    state.streaming = false;
    $("stopRow").classList.add("hidden");
    $("btnSend").disabled = false;
    scroll();

    audit("response", { text: answerText, error: errorText });
  }

  function makeButton(label, handler) {
    var btn = document.createElement("button");
    btn.textContent = label;
    btn.addEventListener("click", handler);
    return btn;
  }

  /* ------------------------------------------------------------ plan review */

  function opIcon(op) {
    switch (op) {
      case "addSheet": return "+";
      case "renameSheet": return "↻";
      case "deleteSheet": return "−";
      case "setValues": return "✎";
      case "setFormulas": return "fx";
      case "insertRows": return "⇣";
      case "deleteRows": return "⇡";
      case "insertColumns": return "⇢";
      case "deleteColumns": return "⇠";
      case "formatRange": return "🎨";
      case "clearRange": return "⌫";
      case "autofitRange": return "↔";
      case "mergeRange": case "unmergeRange": return "▦";
      case "addTable": case "deleteTable": return "▤";
      case "sortRange": return "⇅";
      case "addChart": return "📊";
      case "deleteChart": return "⌫";
      case "freezePanes": return "❄";
      default: return "•";
    }
  }

  function renderPlanPreview(parsedPlan, wantsPlan) {
    var list = $("previewList");
    list.innerHTML = "";
    parsedPlan.operations.forEach(function (op, index) {
      var row = document.createElement("div");
      row.className = "op-row";

      var cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = !op.destructive && !state.readOnly;
      if (state.readOnly) cb.disabled = true;
      cb._opIndex = index;

      var icon = document.createElement("span");
      icon.className = "op-icon";
      icon.textContent = opIcon(op.op);

      var desc = document.createElement("div");
      desc.className = "op-desc";
      desc.textContent = op.description || op.label;

      var cells = document.createElement("span");
      cells.className = "op-cells";
      cells.textContent = op.cells.join(", ") || "—";

      var badge = document.createElement("span");
      badge.className = "op-badge " + (op.destructive ? "destructive" : "safe");
      badge.textContent = op.destructive ? "destructive" : "safe";

      row.appendChild(cb);
      row.appendChild(icon);
      row.appendChild(desc);
      row.appendChild(cells);
      row.appendChild(badge);
      list.appendChild(row);
    });

    var nDestructive = parsedPlan.operations.filter(function (o) { return o.destructive; }).length;
    var countText = parsedPlan.cellCount === Infinity ? "structural" : parsedPlan.cellCount + " cell(s)";
    $("previewCount").textContent = parsedPlan.operations.length + " operation(s) · " + countText +
      (nDestructive ? " · " + nDestructive + " destructive" : "");

    var note = $("previewNote");
    if (state.readOnly) {
      note.textContent = "Read-only mode is on — changes cannot be applied.";
    } else if (nDestructive) {
      note.textContent = "Destructive operations overwrite or shift existing data — review them carefully.";
    } else {
      note.textContent = "Check the operations you want to apply, then click Apply checked changes.";
    }
    $("previewPanel").classList.remove("hidden");
    $("btnApplyPlan").disabled = state.readOnly;
    var boxes = list.querySelectorAll('input[type="checkbox"]');
    var checked = Array.prototype.filter.call(boxes, function (box) { return box.checked; }).length;
    $("previewSelectAll").checked = boxes.length > 0 && checked === boxes.length;
    $("previewSelectAll").indeterminate = checked > 0 && checked < boxes.length;

    $("planEdit").value = state.planJson || "";
    $("planJsonText").textContent = state.planJson || "";
    $("planDetails").classList.add("hidden");
    $("btnApplyEdited").classList.add("hidden");
  }

  function hidePreview() {
    $("previewPanel").classList.add("hidden");
    state.currentPlan = null;
    state.planJson = null;
  }

  function collectCheckedOps() {
    var boxes = $("previewList").querySelectorAll("input[type=checkbox]");
    var ops = [];
    for (var i = 0; i < boxes.length; i++) {
      if (boxes[i].checked) ops.push(state.currentPlan.operations[boxes[i]._opIndex]);
    }
    return ops;
  }

  async function applyCore(ops) {
    if (!inExcel()) return banner("Open this pane inside Excel to apply changes.", "warn");
    if (!ops.length) return banner("Select at least one change to apply.", "warn");
    if (state.readOnly) return banner("Read-only mode is on — changes cannot be applied.", "warn");

    var summary = (state.currentPlan && state.currentPlan.summary) || "Workbook changes";
    var snapshotOps = ops.filter(function (op) { return plan.isUndoSafe(op); });
    var snapshotTotal = snapshotOps.reduce(function (sum, op) { return sum + op.cellCount; }, 0);
    var snapshotEligible = snapshotOps.length > 0 && Number.isFinite(snapshotTotal) &&
      snapshotTotal <= plan.MAX_PREVIEW_CELLS &&
      snapshotOps.every(function (op) { return op.cellCount > 0 && op.cellCount <= plan.MAX_PREVIEW_CELLS; });
    var snap = null;
    var snapshotError = null;
    if (snapshotEligible) {
      try { snap = await plan.snapshotCells(Excel, snapshotOps); }
      catch (err) {
        // The reviewed apply is still legitimate; it just cannot offer the
        // add-in's cell-level Undo entry. Keep the reason for the final summary.
        snapshotError = err.message;
      }
    }

    var results;
    try { results = await plan.applyOps(Excel, ops); } catch (err) { return banner("Apply failed: " + err.message, "err"); }

    await refreshContext();

    var okCount = results.filter(function (r) { return r.ok; }).length;
    var failed = results.filter(function (r) { return !r.ok; });

    var successfulOps = ops.filter(function (op, index) { return results[index] && results[index].ok; });
    var addinUndoable = successfulOps.length > 0 && !!snap && successfulOps.every(function (op) {
      return plan.isUndoSafe(op);
    });
    if (addinUndoable) {
      state.undoStack.push({ snap: snap, ops: successfulOps, at: Date.now(), summary: summary });
      if (state.undoStack.length > 8) state.undoStack.shift();
    } else if (successfulOps.length) {
      // A structural/object edit can invalidate every older cell snapshot by
      // shifting or removing its target. Discard stale entries rather than
      // offering an Undo button that may restore the wrong cells.
      state.undoStack = [];
    }

    hidePreview();
    if (addinUndoable) showUndoRow(okCount, results.length);
    else hideUndoRow();
    var notice = "Applied " + okCount + "/" + results.length + " change(s)" +
      (failed.length ? " · " + failed.length + " failed" : "");
    if (successfulOps.length && !addinUndoable) {
      notice += " · Add-in Undo is unavailable for this plan; use Excel Undo or a backup.";
      if (snapshotError) notice += " (snapshot failed: " + snapshotError + ")";
    }
    banner(notice, okCount === results.length && addinUndoable ? "ok" : "warn");
    audit("apply", { summary: summary, operations: ops, results: results });
    return results;
  }

  async function applyPlan() {
    var ops = collectCheckedOps();
    return applyCore(ops);
  }

  async function applyEditedPlan() {
    if (!plan) return banner("The plan module is not available.", "err");
    var parsed = plan.parsePlan($("planEdit").value);
    if (!parsed.ok) return banner(parsed.error, "err");
    state.currentPlan = parsed;
    state.planJson = $("planEdit").value;
    return applyCore(parsed.operations);
  }

  function showUndoRow(okCount, total) {
    $("undoRow").classList.remove("hidden");
    $("undoText").textContent = "Applied " + okCount + "/" + total + " change(s) — undo?";
  }

  function hideUndoRow() {
    $("undoRow").classList.add("hidden");
  }

  async function undoLast() {
    var entry = state.undoStack[state.undoStack.length - 1];
    if (!entry) return banner("Nothing to undo.", "warn");
    if (!inExcel()) return banner("Open this pane inside Excel to undo changes.", "warn");
    try {
      var restored = await plan.undoSnapshot(Excel, entry.snap);
      state.undoStack.pop();
      await refreshContext();
      banner("Undid " + restored + " changed range(s) from \"" + entry.summary + "\".", "ok");
      audit("undo", { summary: entry.summary, operations: entry.ops });
    } catch (err) {
      banner("Undo failed: " + err.message + " You can retry; the undo entry was kept.", "err");
    }
    if (!state.undoStack.length) hideUndoRow();
  }

  /* ------------------------------------------------------------ excel glue */

  function inExcel() {
    return typeof Excel !== "undefined";
  }

  function sessionStorageKey() {
    var identity = "unsaved";
    try {
      identity = (typeof Office !== "undefined" && Office.context && Office.context.document && Office.context.document.url) || identity;
    } catch (err) { /* host may not expose the document URL */ }
    var hash = 0;
    for (var i = 0; i < identity.length; i++) hash = ((hash << 5) - hash + identity.charCodeAt(i)) | 0;
    return SESSION_KEY + "." + Math.abs(hash);
  }

  function resolveMentions(question, context) {
    var mentioned = [];
    var tokens = String(question).match(/@([A-Za-z0-9_\-]+)/g) || [];
    tokens.forEach(function (tok) {
      var name = tok.slice(1).trim();
      (context && context.sheets || []).forEach(function (s) {
        if (s.name.toLowerCase() === name.toLowerCase() && mentioned.indexOf(s.name) < 0) mentioned.push(s.name);
      });
    });
    return mentioned;
  }

  async function readWorkbookContext() {
    if (!inExcel()) {
      state.context = null;
      $("contextInfo").textContent = "Preview mode - open inside Excel to use the selection";
      $("ctxSheets").innerHTML = "";
      return;
    }
    try {
      state.context = await Excel.run(async function (ctx) {
        var range = ctx.workbook.getSelectedRange();
        range.load(["address", "values", "formulas", "rowCount", "columnCount"]);
        var active = ctx.workbook.worksheets.getActiveWorksheet();
        active.load("name");
        var sheets = ctx.workbook.worksheets;
        sheets.load("items/name");
        await ctx.sync();

        var maxCells = Number($("maxCells").value || 400);
        var values = range.values || [];
        var total = range.rowCount * range.columnCount;
        var truncated = total > maxCells;
        var maxRows = Math.max(1, Math.floor(maxCells / Math.max(1, range.columnCount)));
        var slice = values.slice(0, maxRows);
        var tsv = slice
          .map(function (row) {
            return row
              .map(function (cell) {
                var value = cell === null || cell === undefined ? "" : String(cell);
                return value.replace(/[\t\n\r]/g, " ");
              })
              .join("\t");
          })
          .join("\n");

        var formula = null;
        if (range.rowCount === 1 && range.columnCount === 1) {
          formula = range.formulas && range.formulas[0] && range.formulas[0][0];
        }
        if (formula && String(formula).indexOf("=") !== 0) formula = null;

        var c = {
          address: range.address,
          sheet: active.name || "",
          rowCount: range.rowCount,
          columnCount: range.columnCount,
          truncated: truncated,
          tsv: tsv,
          formula: formula,
          sheets: (sheets.items || []).map(function (s) { return { name: s.name }; }),
          sheetDetails: []
        };

        /* Per-sheet dimensions, tables and (bounded) values. */
        var details = [];
        (c.sheets || []).forEach(function (sinfo) {
          var ws = ctx.workbook.worksheets.getItem(sinfo.name);
          var used = typeof ws.getUsedRangeOrNullObject === "function" ? ws.getUsedRangeOrNullObject() : ws.getUsedRange();
          used.load(["address", "rowCount", "columnCount", "rowIndex", "columnIndex", "isNullObject"]);
          var tables = ws.tables;
          tables.load("items/name");
          details.push({ ws: ws, used: used, tables: tables, name: sinfo.name, info: null });
        });
        await ctx.sync();

        var budget = CONTEXT_CHAR_BUDGET;
        var usedChars = 0;
        details.forEach(function (d) {
          var info = { name: d.name, address: "", rows: 0, cols: 0, tables: [], tsv: "", formulas: [], sampleAddress: "" };
          try {
            if (d.used.isNullObject) {
              c.sheetDetails.push(info);
              return;
            }
            info.address = d.used.address || "";
            info.rows = Number(d.used.rowCount || 0);
            info.cols = Number(d.used.columnCount || 0);
            (d.tables.items || []).forEach(function (t) { if (t.name) info.tables.push(t.name); });
            var focused = (d.name === c.sheet) || (state.mentionedSheets.indexOf(d.name) >= 0);
            var includeValues = true;
            if (includeValues && info.rows > 0 && info.cols > 0 && usedChars < budget) {
              var startRow = Number(d.used.rowIndex || 0) + 1;
              var startCol = Number(d.used.columnIndex || 0);
              var rowCap = focused ? WORKBOOK_ROW_CAP : 40;
              var colCap = focused ? WORKBOOK_COL_CAP : 25;
              var sampleRows = Math.min(info.rows, rowCap);
              var sampleCols = Math.min(info.cols, colCap);
              var firstCol = colName(startCol);
              var lastCol = colName(startCol + sampleCols - 1);
              var lastRow = startRow + sampleRows - 1;
              var vr = d.ws.getRange(firstCol + startRow + ":" + lastCol + lastRow);
              vr.load(["values", "formulas"]);
              d.valueRange = vr;
              d.infoRef = info;
              d.sampleStartRow = startRow;
              d.sampleStartCol = startCol;
              info.sampleAddress = firstCol + startRow + ":" + lastCol + lastRow;
            }
          } catch (err) {
            info.error = err.message;
          }
          c.sheetDetails.push(info);
        });
        await ctx.sync();

        details.forEach(function (d) {
          if (!d.valueRange) return;
          var info = d.infoRef;
          var vals = d.valueRange.values || [];
          var lines = [];
          for (var r = 0; r < vals.length; r++) {
            var row = vals[r] || [];
            lines.push(row.map(function (cell) {
              var v = cell == null ? "" : String(cell);
              return v.replace(/\t/g, " ");
            }).join("\t"));
          }
          var block = lines.join("\n");
          if (usedChars + block.length > CONTEXT_CHAR_BUDGET && lines.length > 1) {
            block = lines.slice(0, 1).join("\n") + "\n... (context budget exceeded)";
          }
          info.tsv = block;
          usedChars += block.length;
          var formulas = d.valueRange.formulas || [];
          for (var fr = 0; fr < formulas.length && info.formulas.length < 250; fr++) {
            for (var fc = 0; fc < (formulas[fr] || []).length && info.formulas.length < 250; fc++) {
              var formulaValue = formulas[fr][fc];
              if (typeof formulaValue === "string" && formulaValue.charAt(0) === "=") {
                info.formulas.push(colName(d.sampleStartCol + fc) + (d.sampleStartRow + fr) + "=" + formulaValue);
              }
            }
          }
          usedChars += info.formulas.join("\n").length;
        });

        return c;
      });
    } catch (err) {
      state.context = null;
      $("contextInfo").textContent = "Could not read workbook context: " + err.message;
      $("ctxSheets").innerHTML = "";
      return;
    }
    refreshContextDisplay();
  }

  function refreshContextDisplay() {
    var c = state.context;
    if (!c) {
      $("contextInfo").textContent = "No context";
      $("ctxSheets").innerHTML = "";
      return;
    }
    $("contextInfo").textContent =
      c.sheet + "!" + c.address.split("!").pop() + "  ·  " + c.rowCount + "×" + c.columnCount +
      (c.truncated ? "  ·  truncated" : "") + (c.formula ? "  ·  formula" : "");

    var chips = "";
    (c.sheets || []).forEach(function (s) {
      var marked = state.mentionedSheets.indexOf(s.name) >= 0;
      chips += '<span class="ctx-chip">' + escapeHtml(s.name) + (marked ? " @" : "") + "</span>";
    });
    $("ctxSheets").innerHTML = chips
      ? '<span style="font-size:11px;color:var(--text-soft);margin-right:6px;">Workbook:</span>' + chips
      : "";
  }

  async function refreshContext() {
    await readWorkbookContext();
  }

  function targetA1() {
    return state.context ? lib.belowAddress(state.context.address) : null;
  }

  function startA1() {
    return state.context ? lib.startAddress(state.context.address) : null;
  }

  function rangeFromStart(a1, rows, columns) {
    var match = String(a1 || "").match(/^([A-Z]+)(\d+)$/i);
    if (!match) return a1;
    var startCol = 0;
    String(match[1]).toUpperCase().split("").forEach(function (ch) {
      startCol = startCol * 26 + ch.charCodeAt(0) - 64;
    });
    startCol -= 1;
    var end = colName(startCol + Math.max(1, columns) - 1) + (Number(match[2]) + Math.max(1, rows) - 1);
    return a1 === end ? a1 : a1 + ":" + end;
  }

  async function applyDirectOperation(raw, summary) {
    if (!plan) return banner("The workbook plan module did not load.", "err");
    var op = plan.normalizeOp(raw);
    if (!op.ok) return banner(op.error, "err");
    state.currentPlan = { summary: summary, operations: [op], cellCount: op.cellCount };
    return applyCore([op]);
  }

  async function insertBelow(text) {
    if (state.readOnly) return banner("Read-only mode is on — the workbook cannot be edited.", "warn");
    if (!inExcel()) return banner("Open this pane inside Excel to insert into the grid", "warn");
    var a1 = targetA1();
    if (!a1) return banner("Select a cell first", "warn");
    var rows = String(text).split("\n").map(function (line) { return [line]; });
    return applyDirectOperation({
      op: "setValues", sheet: state.context.sheet, range: rangeFromStart(a1, rows.length, 1),
      values: rows, description: "Insert the answer below the selection."
    }, "Insert answer");
  }

  async function insertTable(rows) {
    if (state.readOnly) return banner("Read-only mode is on — the workbook cannot be edited.", "warn");
    if (!inExcel()) return banner("Open this pane inside Excel to insert a table", "warn");
    var a1 = targetA1();
    if (!a1) return banner("Select a cell first", "warn");
    var range = rangeFromStart(a1, rows.length, rows[0].length);
    var valuesOp = plan.normalizeOp({ op: "setValues", sheet: state.context.sheet, range: range, values: rows, description: "Insert the answer as a table." });
    var fitOp = plan.normalizeOp({ op: "autofitRange", sheet: state.context.sheet, range: range, description: "Autofit the inserted table." });
    state.currentPlan = { summary: "Insert table", operations: [valuesOp], cellCount: valuesOp.cellCount };
    var result = await applyCore([valuesOp]);
    if (result) {
      // Autofit is cosmetic. Run it separately so it cannot make the inserted
      // values appear non-undoable, and do not let a width failure mask a
      // successful insert.
      try {
        var fitResults = await plan.applyOps(Excel, [fitOp]);
        if (fitResults[0] && !fitResults[0].ok) {
          banner("Inserted the table, but autofit failed: " + fitResults[0].error, "warn");
        }
      } catch (err) {
        banner("Inserted the table, but autofit failed: " + err.message, "warn");
      }
    }
    return result;
  }

  async function insertFormula(formula) {
    if (state.readOnly) return banner("Read-only mode is on — the workbook cannot be edited.", "warn");
    if (!inExcel()) return banner("Open this pane inside Excel to write a formula", "warn");
    var a1 = startA1();
    if (!a1) return banner("Select the target cell first", "warn");
    return applyDirectOperation({
      op: "setFormulas", sheet: state.context.sheet, range: a1, formula: formula,
      description: "Write the returned live formula into the selected cell."
    }, "Insert formula");
  }

  /* ------------------------------------------------------------ attachments */

  function isTextFile(file) {
    return /^(text\/|application\/(json|csv|xml))/i.test(file.type) ||
      /\.(csv|tsv|txt|md|json|html|xml|log)$/i.test(file.name);
  }

  function readFileText(file) {
    return new Promise(function (resolve) {
      try {
        var reader = new FileReader();
        reader.onload = function () { resolve(reader.result); };
        reader.onerror = function () { resolve(null); };
        reader.readAsText(file);
      } catch (err) {
        resolve(null);
      }
    });
  }

  function readFileBase64(file) {
    return new Promise(function (resolve) {
      try {
        var reader = new FileReader();
        reader.onload = function () {
          var result = String(reader.result || "");
          resolve(result.indexOf(",") >= 0 ? result.slice(result.indexOf(",") + 1) : result);
        };
        reader.onerror = function () { resolve(null); };
        reader.readAsDataURL(file);
      } catch (err) {
        resolve(null);
      }
    });
  }

  async function attachFiles(fileList) {
    for (var i = 0; i < fileList.length; i++) {
      var file = fileList[i];
      if (state.attachments.length >= MAX_ATTACHMENTS) {
        banner("Attachment limit reached (20 files per conversation).", "warn");
        break;
      }
      if (file.size > MAX_ATTACHMENT_BYTES) {
        banner("File too large: " + escapeHtml(file.name) + " (max 25 MB).", "warn");
        continue;
      }
      var entry = { id: "att_" + Date.now() + "_" + i, uploadID: null, uri: null, name: file.name, type: file.type || "application/octet-stream", size: file.size, text: null };
      if (isTextFile(file)) {
        entry.text = await readFileText(file);
        if (entry.text != null && entry.text.length > ATTACH_TEXT_CAP) {
          entry.text = entry.text.slice(0, ATTACH_TEXT_CAP) + "\n... (file truncated)";
        }
      }
      var base64 = await readFileBase64(file);
      if (base64 != null) {
        try {
          var uploaded = await api("/api/attachment", {
            method: "POST",
            body: JSON.stringify({ name: entry.name, type: entry.type, size: entry.size, data: base64 })
          });
          entry.uploadID = uploaded.id;
          entry.uri = uploaded.uri;
        } catch (err) {
          banner("Could not upload " + entry.name + ": " + err.message, "warn");
        }
      }
      state.attachments.push(entry);
    }
    renderAttachments();
    if (state.attachments.length) banner(state.attachments.length + " file(s) attached.", "ok");
  }

  function renderAttachments() {
    var container = $("attachChips");
    container.innerHTML = "";
    state.attachments.forEach(function (a) {
      var chip = document.createElement("span");
      chip.className = "attach-chip";
      chip.textContent = escapeHtml(a.name) + " (" + Math.round(a.size / 1024) + " KB)";
      var rm = document.createElement("button");
      rm.className = "remove";
      rm.textContent = "✕";
      rm.setAttribute("aria-label", "Remove " + escapeHtml(a.name));
      rm.addEventListener("click", function () { removeAttachment(a.id); });
      chip.appendChild(rm);
      container.appendChild(chip);
    });
    $("attachRow").classList.toggle("hidden", state.attachments.length === 0);
  }

  function removeAttachment(id) {
    var removed = state.attachments.filter(function (a) { return a.id === id; })[0];
    state.attachments = state.attachments.filter(function (a) { return a.id !== id; });
    if (removed && removed.uploadID) {
      api("/api/attachment?id=" + encodeURIComponent(removed.uploadID), { method: "DELETE" }).catch(function () {});
    }
    renderAttachments();
  }

  function clearAttachments() {
    state.attachments.slice().forEach(function (a) { removeAttachment(a.id); });
  }

  /* ---------------------------------------------------------------- sending */

  async function submitPromptRequest(request) {
    await api("/api/prompt", {
      method: "POST",
      body: JSON.stringify(request),
    });
    clearTimeout(state.doneTimer);
    state.doneTimer = setTimeout(function () {
      if (state.streaming) {
        if (startFreeModelFallback(state.answer, state.lastError)) return;
        banner("No completion event arrived - showing what was received so far.", "warn");
        finalize(state.answer);
      }
    }, 180000);
  }

  async function send() {
    var input = $("input");
    var text = input.value.trim();
    if (!text || state.streaming) return;
    if (!state.connected) {
      var status = await checkStatus();
      if (!status || !status.connected) return;
    }
    try {
      if (!state.sessionID) await ensureSession();
      state.mentionedSheets = resolveMentions(text, state.context);
      await refreshContext();
    } catch (err) {
      return banner(err.message, "err");
    }

    var prompt = buildPrompt(text);
    addMessage("user", escapeHtml(text));
    input.value = "";
    resizeComposer();
    state.wantsPlan = false;
    beginAssistant();
    state.fallbackAttempted = false;
    state.currentPromptRequest = {
      sessionID: state.sessionID,
      text: prompt,
      agent: $("agentSelect").value || null,
      skill: $("skillSelect").value || null,
      model: selectedModel(),
      files: state.attachments.filter(function (a) { return a.uri && a.uploadID; }).map(function (a) {
        return { id: a.uploadID, uri: a.uri, name: a.name };
      })
    };
    audit("request", {
      prompt: text,
      selection: state.context ? state.context.address : null,
      sheets: state.mentionedSheets,
      model: selectedModel(),
      attachments: state.attachments.map(function (a) {
        return { name: a.name, size: a.size, type: a.type };
      })
    });

    try {
      await submitPromptRequest(state.currentPromptRequest);
    } catch (err) {
      if (startFreeModelFallback("", err.message)) return;
      banner(err.message || "Request failed", "err");
      if (state.bubble) finalize("**Request failed:** " + err.message);
    }
  }

  function buildPrompt(question) {
    var c = state.context || {};
    var parts = [PREAMBLE];

    parts.push("");
    parts.push("## Workbook context");
    parts.push("Active sheet: " + (c.sheet || "—") + " (workbook sheets: " +
      ((c.sheets && c.sheets.length) ? c.sheets.map(function (s) { return s.name; }).join(", ") : "none") + ")");
    if (state.mentionedSheets.length) {
      parts.push("Focus sheets (@mentioned): " + state.mentionedSheets.join(", "));
    }
    if (c.sheetDetails && c.sheetDetails.length) {
      c.sheetDetails.forEach(function (d) {
            var line = "Sheet " + d.name + ": " + (d.address || "—") + " (" + d.rows + "×" + d.cols + ")";
        if (d.tables && d.tables.length) line += "; tables: " + d.tables.join(", ");
        parts.push(line);
        if (d.tsv) parts.push("  Values from " + (d.sampleAddress || d.address) + " (TSV):\n" + d.tsv);
        if (d.formulas && d.formulas.length) parts.push("  Live formulas (cell=formula):\n" + d.formulas.join("\n"));
      });
    }
    if (c && $("includeSelection").checked && c.tsv) {
      parts.push("Selection " + c.address + " (" + c.rowCount + " rows × " + c.columnCount + " columns" + (c.truncated ? ", truncated" : "") + "), values as TSV:");
      if (c.formula) parts.push("Selected cell formula: " + c.formula);
      parts.push(c.tsv);
    }
    if (state.attachments.length) {
      parts.push("");
      parts.push("## Attached files");
      state.attachments.forEach(function (a) {
        parts.push("### " + String(a.name) + " (" + a.type + ", " + Math.round(a.size / 1024) + " KB)");
        if (a.text != null) parts.push(a.text);
        else if (a.uri) parts.push("[File content attached natively to this OpenCode prompt]");
        else parts.push("[binary file - content not extracted in this host]");
      });
    }
    parts.push("");
    parts.push("## Task");
    parts.push(question);
    parts.push("");
    if (isSummaryRequest(question)) {
      parts.push("## Summary response requirements");
      parts.push(SUMMARY_HINT);
      parts.push("");
    }
    parts.push(state.wantsPlan ? PLAN_HINT : TEXT_HINT);
    return parts.join("\n");
  }

  function isSummaryRequest(question) {
    return /\b(summar(?:y|ize|ise|ized|ised|izing|ising)|overview|key insights?|analy[sz]e (?:this|the) (?:table|data|dataset)|what (?:does|do) (?:this|the) (?:table|data) show)\b/i.test(String(question || ""));
  }

  var QUICK = {
    budget: { text: "Build a complete personal monthly budget in a new sheet with income and expense categories, live formulas, monthly totals, savings rate, professional formatting, and a summary chart.", plan: true },
    debt: { text: "Build a debt payoff planner in a new sheet with balance, interest rate, minimum payment, extra payment, payoff order, live formulas, and a month-by-month payoff summary.", plan: true },
    explain: { text: "Explain the selected cells or formula step by step, in plain language. Cite the cells involved.", plan: false },
    fix: { text: "The selected formula is broken or returns the wrong result. Diagnose it and give the corrected live formula starting with '='.", plan: false },
    clean: { text: "Clean this data: fix headers, trim whitespace, normalise dates and numbers, remove obvious duplicates. Propose the result as a JSON plan of setValues / formatRange operations the user can review before applying.", plan: true },
    table: { text: "Summarise the selected data with a headline and 4 to 6 decision-useful insights covering the main KPI, strongest category or region, leading contributor, time trend, and material outlier. Cite the source range for every insight.", plan: false },
    vba: { text: "Write a VBA macro for this task. Return only the code in a ```vba block plus a one-line usage note.", plan: false },
    pivot: { text: "Recommend the best pivot table layout for this data and give the exact steps plus any formulas needed.", plan: false },
    formula: { text: "Write one live Excel formula that does what I ask. Return only the formula, starting with '='.", plan: false },
    audit: { text: "Audit this workbook for broken links, hard-coded values, inconsistent formulas and structural issues. List each issue with its cell address.", plan: false },
    variance: { text: "Compare actuals versus budget and flag variances above 10%. Return the comparison as a ```tsv block plus short observations.", plan: false },
    scenario: { text: "Build base, upside and downside scenario tabs linked by formulas, plus a summary table comparing key outputs.", plan: true },
    dashboard: { text: "Build a dashboard tab with KPI figures and charts summarising the data across this workbook.", plan: true },
    newsheet: { text: "Add a new sheet with clean headers and formulas for tracking this data.", plan: true },
    research: { text: "Research the web for the information I request, cite every external source with its URL, and return a clean table suitable for insertion into this workbook.", plan: false },
  };

  async function quickAction(action) {
    if (action === "format") {
      await formatSheet();
      return;
    }
    var entry = QUICK[action];
    if (!entry) return;
    $("settingsPanel").classList.add("hidden");
    var input = $("input");
    var typed = input.value.trim();
    input.value = entry.text + (typed ? "\n\n" + typed : "");
    resizeComposer();
    state.wantsPlan = !!entry.plan;
    await send();
  }

  /* Native formatting: the pane writes to the sheet itself, so it works even
     when the model cannot access cells. */
  async function formatSheet() {
    if (state.readOnly) return banner("Read-only mode is on — the workbook cannot be edited.", "warn");
    if (!inExcel()) return banner("Open this pane inside Excel to format a sheet", "warn");
    if (!formatter) return banner("The sheet formatter did not load - reload the add-in.", "err");
    var chip = document.querySelector('.chip[data-action="format"]');
    var label = chip ? chip.textContent : "Format sheet";
    if (chip) {
      chip.disabled = true;
      chip.textContent = "Formatting…";
    }
    try {
      var result = await formatter.formatSheet(Excel, lib, {
        onWarning: function (message) { console.warn("Format sheet:", message); },
      });
      var text =
        "Formatted " + result.sheetName + " · " + result.tableName + " · " + result.dataRows + " data rows";
      if (result.created) text += " · summary added";
      if (result.warnings && result.warnings.length) {
        text += " · " + result.warnings.length + " optional step(s) skipped";
      }
      banner(text, result.warnings && result.warnings.length ? "warn" : "ok");
      audit("format-sheet", result);
      await refreshContext();
    } catch (err) {
      banner("Format failed: " + (err.message || err), "err");
    } finally {
      if (chip) {
        chip.disabled = false;
        chip.textContent = label;
      }
    }
  }

  async function newChat() {
    try {
      var data = await api("/api/session", { method: "POST", body: JSON.stringify({}) });
      state.sessionID = data.sessionID;
      try { localStorage.setItem(sessionStorageKey(), state.sessionID); } catch (err) { /* ignore */ }
      openStream();
      $("messages").innerHTML = suggestionsHtml();
      wireQuickActions($("messages"));
      $("input").value = "";
      clearAttachments();
      state.mentionedSheets = [];
      state.undoStack = [];
      state.currentPlan = null;
      state.planJson = null;
      state.currentPromptRequest = null;
      state.fallbackAttempted = false;
      hidePreview();
      hideUndoRow();
      renderAttachments();
    } catch (err) {
      banner(err.message, "err");
    }
  }

  async function stop() {
    if (!state.sessionID) return;
    try {
      await api("/api/abort", { method: "POST", body: JSON.stringify({ sessionID: state.sessionID }) });
    } catch (err) {
      /* ignore */
    }
    finalize(state.answer, state.answer ? null : (state.lastError || "Stopped before the model replied."));
  }

  /* ---------------------------------------------------------------- audit log */

  async function audit(type, payload) {
    try {
      await api("/api/audit", {
        method: "POST",
        body: JSON.stringify({
          sessionID: state.sessionID,
          type: type,
          at: new Date().toISOString(),
          payload: payload || {}
        }),
      });
    } catch (err) {
      /* audit logging never breaks the flow */
    }
  }

  async function loadAuditLog() {
    try {
      var log = await api("/api/audit");
      renderAuditLog(log);
    } catch (err) {
      $("auditLog").textContent = "Audit log unavailable: " + err.message;
      $("auditLog").classList.remove("hidden");
    }
  }

  function renderAuditLog(log) {
    var el = $("auditLog");
    var entries = (log && log.entries) || [];
    var lines = entries.map(function (e) {
      var when = "";
      try { when = new Date(e.at).toLocaleTimeString(); } catch {} return when + "  " + e.type + "  " + escapeHtml(JSON.stringify(e.payload || {}).slice(0, 200));
    });
    el.textContent = lines.length ? lines.join("\n") : "(no audit entries yet)";
    el.classList.remove("hidden");
  }

  async function loadConnections() {
    var el = $("connectionsList");
    try {
      var data = await api("/api/connections");
      var lines = [];
      (data.mcp || []).forEach(function (item) {
        var name = item.name || item.id || item.server || "MCP server";
        var status = item.status && (item.status.type || item.status.status) || item.type || "configured";
        lines.push("MCP · " + name + " · " + status);
      });
      (data.integrations || []).forEach(function (item) {
        var name = item.name || item.id || "Integration";
        var status = item.status && (item.status.type || item.status.status) || item.connected && "connected" || "configured";
        lines.push("App · " + name + " · " + status);
      });
      el.textContent = lines.length ? lines.join("\n") : "No MCP servers or app integrations are configured for this OpenCode workspace.";
    } catch (err) {
      el.textContent = "Connections unavailable: " + err.message;
    }
    el.classList.remove("hidden");
  }

  /* ---------------------------------------------------------------- settings */

  function loadSettings() {
    try {
      var ro = localStorage.getItem(READONLY_KEY);
      if (ro !== null) state.readOnly = ro === "true";
      var ap = localStorage.getItem(APPROVAL_KEY);
      if (ap === "always" || ap === "auto" || ap === "never") state.approvalMode = ap;
      var sid = localStorage.getItem(sessionStorageKey());
      if (sid) state.sessionID = sid;
    } catch (err) { /* storage can be disabled inside Office */ }
  }

  function saveSetting(key, value) {
    try { localStorage.setItem(key, value); } catch (err) { /* ignore */ }
  }

  async function loadAgents() {
    try {
      var data = await api("/api/agents");
      var select = $("agentSelect");
      select.innerHTML = '<option value="">Default</option>';
      (data.agents || []).forEach(function (agent) {
        var option = document.createElement("option");
        option.value = agent.id;
        // Native select popups in the Excel WebView size themselves from the
        // longest option. Keep labels short so built-in agent instructions do
        // not create a task-pane-wide horizontal scrollbar.
        option.textContent = String(agent.id).slice(0, 60);
        if (agent.description) option.title = String(agent.description).slice(0, 240);
        select.appendChild(option);
      });
    } catch (err) {
      /* agents are optional */
    }
  }

  async function loadSkills() {
    try {
      var data = await api("/api/skills");
      var select = $("skillSelect");
      select.innerHTML = '<option value="">None</option>';
      (data.skills || []).forEach(function (skill) {
        var option = document.createElement("option");
        option.value = skill.id;
        option.textContent = String(skill.name || skill.id).slice(0, 60);
        if (skill.description) option.title = String(skill.description).slice(0, 240);
        select.appendChild(option);
      });
    } catch (err) {
      /* skills are optional */
    }
  }

  function selectedModel() {
    var value = $("modelSelect").value;
    if (!value) return null;
    try {
      var pair = JSON.parse(value);
      if (pair && pair.length === 2 && pair[0] && pair[1]) {
        return { providerID: String(pair[0]), id: String(pair[1]) };
      }
    } catch (err) {
      /* ignore a corrupt saved value */
    }
    return null;
  }

  function rememberModel() {
    try {
      if ($("modelSelect").value) localStorage.setItem(MODEL_KEY, $("modelSelect").value);
      else localStorage.removeItem(MODEL_KEY);
    } catch (err) {
      /* storage can be disabled inside Office */
    }
  }

  async function loadModels() {
    try {
      var data = await api("/api/models");
      state.modelCatalog = Array.isArray(data.models) ? data.models.slice() : [];
      state.defaultModel = data.defaultModel || null;
      var select = $("modelSelect");
      select.innerHTML = '<option value="">Default model</option>';
      if (data.defaultModel) {
        var defaultID = data.defaultModel.name || data.defaultModel.id || data.defaultModel.modelID;
        if (defaultID) select.options[0].textContent = String(defaultID);
      }

      var groups = {};
      var providers = [];
      state.modelCatalog.forEach(function (model) {
        if (!groups[model.providerID]) {
          groups[model.providerID] = [];
          providers.push(model.providerID);
        }
        groups[model.providerID].push(model);
      });

      providers.forEach(function (provider) {
        var group = document.createElement("optgroup");
        group.label = provider;
        groups[provider].forEach(function (model) {
          var option = document.createElement("option");
          option.value = JSON.stringify([model.providerID, model.id]);
          option.textContent = model.name && model.name !== model.id
            ? model.name + " (" + model.id + ")"
            : model.id;
          group.appendChild(option);
        });
        select.appendChild(group);
      });

      var saved = null;
      try { saved = localStorage.getItem(MODEL_KEY); } catch (err) { /* ignore */ }
      if (saved) {
        for (var i = 0; i < select.options.length; i++) {
          if (select.options[i].value === saved) {
            select.value = saved;
            break;
          }
        }
      }
    } catch (err) {
      /* models are optional; the service default still works */
    }
  }

  /* ------------------------------------------------------------------- init */

  function suggestionsHtml() {
    return '<div class="empty" id="emptyState">' +
      '<div class="suggested-title">Suggested</div>' +
      '<button class="chip suggestion" data-action="budget"><span>↳</span> Build a personal budget</button>' +
      '<button class="chip suggestion" data-action="debt"><span>↳</span> Build a debt payoff planner</button>' +
      '<button class="chip suggestion" data-action="dashboard"><span>↳</span> Build a personal finance dashboard</button>' +
      '</div>';
  }

  function wireQuickActions(root) {
    var chips = (root || document).querySelectorAll(".chip");
    for (var i = 0; i < chips.length; i++) {
      if (chips[i]._openCodeWired) continue;
      chips[i]._openCodeWired = true;
      chips[i].addEventListener("click", function (event) {
        quickAction(event.currentTarget.getAttribute("data-action"));
      });
    }
  }

  function wire() {
    $("btnSend").addEventListener("click", send);
    $("btnStop").addEventListener("click", stop);
    $("btnNewChat").addEventListener("click", newChat);
    $("btnRefreshContext").addEventListener("click", refreshContext);
    $("btnSettings").addEventListener("click", function () {
      $("settingsPanel").classList.toggle("hidden");
    });
    $("modelSelect").addEventListener("change", rememberModel);
    $("chkReadOnly").addEventListener("change", function () {
      state.readOnly = $("chkReadOnly").checked;
      saveSetting(READONLY_KEY, state.readOnly);
      if ($("previewPanel") && !$("previewPanel").classList.contains("hidden")) {
        renderPlanPreview(state.currentPlan || { operations: [], summary: "" }, false);
      }
      $("btnApplyPlan").disabled = state.readOnly;
    });
    $("selApproval").addEventListener("change", function () {
      state.approvalMode = $("selApproval").value;
      saveSetting(APPROVAL_KEY, state.approvalMode);
      if ($("previewPanel") && !$("previewPanel").classList.contains("hidden")) {
        renderPlanPreview(state.currentPlan || { operations: [], summary: "" }, false);
      }
    });
    $("input").addEventListener("keydown", function (event) {
      if ((event.ctrlKey || event.metaKey) && event.key === "Enter") {
        event.preventDefault();
        send();
      }
    });
    $("input").addEventListener("input", resizeComposer);
    $("fileInput").addEventListener("change", function (event) {
      attachFiles(event.target.files || []);
      event.target.value = "";
    });
    $("btnAttach").addEventListener("click", function () { $("fileInput").click(); });
    $("btnAttachTop").addEventListener("click", function () { $("fileInput").click(); });
    $("btnClearAttachments").addEventListener("click", function () {
      clearAttachments();
    });
    $("btnApplyPlan").addEventListener("click", applyPlan);
    $("previewSelectAll").addEventListener("change", function () {
      var boxes = $("previewList").querySelectorAll('input[type="checkbox"]');
      for (var i = 0; i < boxes.length; i++) {
        if (!boxes[i].disabled) boxes[i].checked = $("previewSelectAll").checked;
      }
      $("previewSelectAll").indeterminate = false;
    });
    $("btnCancelPlan").addEventListener("click", hidePreview);
    $("btnShowPlanJson").addEventListener("click", function () {
      var details = $("planDetails");
      var hidden = details.classList.contains("hidden");
      details.classList.toggle("hidden");
      if (hidden) {
        $("planJsonText").textContent = state.planJson || "";
        $("planEdit").value = state.planJson || "";
        $("btnApplyEdited").classList.remove("hidden");
      } else {
        $("btnApplyEdited").classList.add("hidden");
      }
    });
    $("btnApplyEdited").addEventListener("click", applyEditedPlan);
    $("btnUndo").addEventListener("click", undoLast);
    $("btnAuditLog").addEventListener("click", function () {
      var log = $("auditLog");
      if (log.classList.contains("hidden")) {
        loadAuditLog();
      } else {
        log.classList.add("hidden");
      }
    });
    $("btnConnections").addEventListener("click", function () {
      var el = $("connectionsList");
      if (el.classList.contains("hidden")) loadConnections();
      else el.classList.add("hidden");
    });
    document.addEventListener("click", function (event) {
      var panel = $("settingsPanel");
      if (!panel.classList.contains("hidden") && !panel.contains(event.target) && !$("btnSettings").contains(event.target)) {
        panel.classList.add("hidden");
      }
    });
    wireQuickActions(document);
  }

  async function init() {
    loadSettings();
    wire();
    $("chkReadOnly").checked = state.readOnly;
    $("selApproval").value = state.approvalMode;
    $("btnApplyPlan").disabled = state.readOnly;
    var status = await checkStatus();
    if (status && status.connected) {
      await ensureSession();
      await Promise.all([loadAgents(), loadModels(), loadSkills()]);
    }
    await refreshContext();

    if (inExcel()) {
      try {
        Excel.run(function (ctx) {
          ctx.workbook.onSelectionChanged.add(refreshContext);
          return ctx.sync();
        });
      } catch (err) {
        /* selection-changed events are optional */
      }
    }
    setInterval(checkStatus, 30000);
  }

  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(function () { init(); });
  } else {
    window.addEventListener("load", init);
  }
})();
