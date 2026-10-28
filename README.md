# OpenCode Excel Add-in

OpenCode for Excel is a local Microsoft Excel add-in for workbook analysis and automation.
It connects an Office.js task pane to the models configured in OpenCode, sends bounded workbook
context, and turns model responses into formulas, tables, dashboards, formatting, and reviewable
cell changes.

The add-in is designed for analysts and spreadsheet users who want an AI assistant in Excel
without sending workbook traffic through a separate hosted add-in service. The local bridge talks
to your existing OpenCode service and keeps workbook writes inside Excel through Office.js.

If you prefer a guided Windows setup, read [`START-HERE.md`](START-HERE.md). It explains the
one-click launcher without requiring terminal commands.

## Key features

- Chat with OpenCode models from a clean Excel task pane.
- Analyze selected cells, formulas, tables, and bounded samples from every worksheet.
- Generate Excel formulas, summaries, variance analysis, scenarios, and dashboard layouts.
- Review workbook edits before applying them, with exact sheet and range references.
- Apply values, formulas, formatting, tables, charts, row and column operations, and sheet changes.
- Restore recent cell writes and single-cell, cell-local formatting with local snapshots; use Excel's own Undo for structural edits.
- Attach documents, images, spreadsheets, CSV files, and text files through OpenCode V2.
- Use OpenCode models, agents, skills, MCP connections, and provider settings.
- Retry unavailable models with a verified free OpenCode fallback without selecting paid providers.
- Run locally with a restricted scratch workspace and no hosted add-in backend.

See [`FEATURE-MATRIX.md`](FEATURE-MATRIX.md) for the exact implemented scope, host-dependent
features, and the work that requires a separate enterprise service.

```
Excel task pane (Office.js)                 Node bridge                 OpenCode
https://localhost:3000/taskpane.html   ->   https://localhost:3000  ->  127.0.0.1:28894
   reads/writes cells via Excel.run         finds the service, holds        your configured
   streams the answer over SSE              the password, survives          providers/models
                                            flaky event streams
```

Tested against **OpenCode v2** (`/api/*` surface) with automatic fallback to
**v1** (`/session/*`). Works on **Excel 2019 / 2021 perpetual and Microsoft 365**
on Windows (Mac desktop works too; Excel *on the web* cannot reach your PC — see
[Limits](#limits)).

---

## Quick start on Windows

1. Install Node.js 18 or later, Microsoft Excel 2019 or later, and OpenCode Desktop.
2. Clone or download this repository.
3. Close every open Excel window.
4. Double-click `Start OpenCode for Excel.bat`.
5. Keep the launcher window open while using the add-in.

To clone the repository:

```powershell
git clone https://github.com/analystsandeep/Openclode-Excel-Addon.git
cd Openclode-Excel-Addon
```

The launcher installs dependencies, prepares the localhost certificate, registers the Office
manifest, creates a compatible starter workbook, starts the local bridge, and opens Excel.

## 1. Requirements

| Requirement | Notes |
|---|---|
| Node.js 18+ | `node -v` — needed only to run the local bridge |
| Excel 2019 / 2021 / Microsoft 365 (desktop) | Manifest validates for "Excel 2019 or later on Windows" |
| Microsoft Edge WebView2 Runtime | Pre-installed on Windows 11 and current Windows 10. If the pane renders blank on the Semi-Annual Enterprise channel, set `HKCU\Software\Microsoft\Office\16.0\WEF\Win32WebView2` (DWORD) to `1` |
| OpenCode desktop app **or** `opencode service start` | v2 keeps a shared background service running on port **28894** (Windows) / 49374 (macOS, Linux) |

## 2. Install and run

The simplest option is to double-click `Start OpenCode for Excel.bat` on Windows. It installs dependencies,
starts the bridge, sideloads the manifest and opens Excel with a small
`OpenCode for Excel.xlsx` starter workbook that activates the Home-tab button and the task
pane. `Remove from Excel.bat` takes it back out, and `Fix Certificates.bat` repairs the
localhost certificate if Excel refuses to load the pane.

From a terminal:

```powershell
cd opencode-excel-addin
npm install
npm run dev
```

You should see:

```
  Task pane : https://localhost:3000/taskpane.html
  Workspace : C:\Users\you\AppData\Local\OpenCodeExcel\workspace

  OpenCode  : connected -> http://127.0.0.1:28894 (v2.x) via registration
```

The first run generates and trusts a localhost development certificate
(`office-addin-dev-certs`). Accept the prompt if Windows asks.

Leave this window open. It is the bridge between Excel and OpenCode.

## 3. Sideload the add-in

In a **second** terminal (keep `npm run dev` running):

```powershell
npm run sideload
```

Then open the bundled `OpenCode for Excel.xlsx` starter workbook (the `.bat` does this
automatically). The **OpenCode** button appears on the **Home** tab and the pane opens on
the right. You can close the starter workbook and keep using the button in your own
workbooks for the rest of the Excel session. To remove the add-in later:
`npm run unsideload`.

> **Office 2021 (perpetual) note:** some volume builds ignore the `Wef\Developer` registry
> entry on its own. Opening a workbook with a `webextension` reference (the bundled starter
> workbook) is what makes Excel register the add-in commands for that session.

### Manual sideload (if the npm step is blocked)

1. Copy `manifest.xml` into a folder, e.g. `C:\OpenCodeAddin`.
2. In Excel, open **File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs**, add `C:\OpenCodeAddin` as the catalog URL, and select *Show in Menu*.
3. Restart Excel, open **Insert > My Add-ins > Shared Folder**, select *OpenCode for Excel*, and choose **Add**.

> If Excel says *"We can't open this add-in from localhost"*, run `npx office-addin-dev-certs install`
> and restart Excel; on locked-down networks also try:
> `CheckNetIsolation LoopbackExempt -a -n="microsoft.win32webviewhost_cw5n1h2txyewy"` from an elevated prompt.

## 4. Using it

* **Ask anything** - type a question; the current selection plus bounded values and live formulas
  from every non-empty tab are sent as context. The whole-workbook context means OpenCode can see
  your tabs, tables, and formula logic, not just the selection. Use `@SheetName` to give one tab a
  larger context budget.
* **Workbook context controls** - **Send workbook samples** in Settings controls the all-sheet
  samples (names, used ranges, values, formulas). It is on by default; turning it off skips every
  sheet read and sends no sheet data. The **selection checkbox** only adds or removes the selected
  cells and does not turn workbook samples off. Both choices are remembered. Workbook and
  selection context are each capped at 30,000 characters; long cells are clipped and the prompt
  notes when data was truncated.
* **Model** - the model selector lists every model your OpenCode app exposes. The choice is remembered
  and applied to the next message; *Default* lets OpenCode pick.
* **Agentic edits (change plans)** - when you ask OpenCode to change the workbook
  (write cells or formulas, insert or delete rows/columns, add or rename sheets, format, add
  charts), it replies with a **change plan** instead of plain text. The plan lists every
  operation with the exact cells/tabs it touches, so you can review before anything is written.
  Uncheck operations you don't want, then click **Apply checked changes**. A before-snapshot is
  kept for cell writes and single-cell, cell-local formatting so you can click **Undo**
  afterwards. Multi-cell formatting, borders, row heights, column widths, autofit, and
  structural/object operations are not add-in-undoable, so use Excel's own Undo or a backup for
  those.
* **Read-only mode** - the three-dot settings menu lets you ask questions and audit without any
  workbook edits being possible (apply buttons are disabled).
* **Approval gates** - use the three-dot settings menu to choose *Ask before every change* (default), *Auto-apply
  safe formatting touching 5 or fewer cells*, or *Auto-apply safe formatting*. Destructive and
  structural operations are always unchecked and never auto-applied. Border, row-height,
  column-width, and autofit formatting always require review because they affect cells outside
  the reviewed range.
* **Quick actions** - *Explain formula*, *Fix formula*, *Write formula*, *Clean this data*,
  *Summarize as table*, *Audit workbook*, *Variance analysis*, *Scenario tabs*, *Dashboard*,
  *New sheet*, *Web research*, *Write VBA*, *Pivot advice*.
* **Attachments** - click **Attach** and add up to 20 files (25 MB each). Files stay in the
  locked local scratch workspace and use OpenCode V2's native attachment API, so capable models
  can inspect PDFs, images, Office files, spreadsheets, and text. Text is also included directly
  in the prompt as a compatibility fallback.
* **Skills and connections** - the three-dot settings menu lists OpenCode Skills/playbooks and configured MCP/app
  connections. Connected tools remain governed by OpenCode's permissions.
* **Format sheet** - styles the active sheet directly with Excel.js, without calling the model:
  cleans headers, creates or restyles `tblSales`, adds a three-row summary strip, applies
  date/quantity/currency formats, freezes the header, hides gridlines, sets landscape print
  area/titles and adds data bars plus a red negative-value rule. Safe to click again — it
  reuses the existing table and does not duplicate the summary or conditional formats.
* **Insert below** - writes the answer one row below your selection, with one line per row.
* **Insert table** - appears when the reply contains a table; parses TSV, CSV, or pipe tables and writes them as a real block with autofit.
* **Use as formula** - appears when the reply is a single line starting with `=`; writes it into the selected cell as a live formula.
* **Copy** - copies the answer or an individual code block.
* **Stop** - aborts the turn (`session.interrupt` or `abort`).
* **New chat** - starts a fresh OpenCode session. History remains in OpenCode.
* Ctrl or Command plus Enter sends.

## 5. Locking the agent down (important)

OpenCode is a *coding* agent: by default it can run shell commands and edit files.
Sessions created by this add-in run in a scratch workspace
(`%LOCALAPPDATA%\OpenCodeExcel\workspace`) that the bridge creates with:

```jsonc
// %LOCALAPPDATA%\OpenCodeExcel\workspace\opencode.jsonc  (auto-created)
{
  "permissions": [
    { "action": "shell", "resource": "*", "effect": "deny" },
    { "action": "edit",  "resource": "*", "effect": "deny" }
  ]
}
```

The bridge listens only on `127.0.0.1`, rejects unexpected `Host` and browser `Origin` headers,
and requires a per-launch token for every API except the health/status probe. The task pane
receives that token from the bridge's own HTML response, so no credential is committed to the
repository.

Add a dedicated agent so the dropdown offers an Excel-tuned, read-only assistant.
Paste this into `%USERPROFILE%\.config\opencode\opencode.jsonc`:

```jsonc
{
  "$schema": "https://opencode.ai/config.json",
  "agents": {
    "excel": {
      "description": "Spreadsheet helper: answers, formulas, and tables only",
      "mode": "subagent",
      "system": "You help inside Excel. Answer from the data given. Never read or write files, never run shell commands.",
      "permissions": [
        { "action": "shell", "resource": "*", "effect": "deny" },
        { "action": "edit", "resource": "*", "effect": "deny" }
      ]
    }
  }
}
```

Then pick **excel** under *Agent* in the three-dot settings menu. Config reference: <https://opencode.ai/v2/docs/config/>.

## 6. Configuration

| Setting | Where | Default |
|---|---|---|
| Bridge port | `PORT` env var | `3000` |
| Bridge bind address | fixed at loopback | `127.0.0.1` |
| Bridge API token | `BRIDGE_API_TOKEN` env var, else generated per launch and injected into the pane | random, 32 bytes |
| Service URL | `OPENCODE_SERVICE_URL` env var, else service registration file | auto-discovered |
| Service password | `OPENCODE_SERVICE_PASSWORD` env var, else registration/config file | auto-discovered |
| Allowed browser origin | checked by the bridge | same-origin `localhost` / `127.0.0.1` only |
| Selected cells sent per request | Three-dot settings menu, *Max selected cells sent* | `400` |
| Workbook samples | Three-dot settings menu, *Send workbook samples* | on |
| Selected cells in prompt | Selection checkbox above the composer | on |
| Read-only mode | Three-dot settings menu, *Read-only mode* | off |
| Approval gates | Three-dot settings menu, *Approval gates* | *Ask before every change* |
| AI model | Composer model selector | OpenCode's session default |

Discovery order: `OPENCODE_SERVICE_URL`, then `%USERPROFILE%\.local\state\opencode\service.json`
(registration written by the running service), then `%USERPROFILE%\.config\opencode\service.json`
(service config: port, hostname, and password), then default ports `28894` and `49374`, and finally
`opencode service status`.
If OpenCode answers `401`, the bridge retries with the password from
`opencode service get password`.

**WSL / remote setups:** start the service reachable from Windows
(`opencode serve --hostname 0.0.0.0 --port 4096` in WSL) and set
`OPENCODE_SERVICE_URL=http://127.0.0.1:4096` plus `OPENCODE_SERVICE_PASSWORD=...`
before `npm run dev`.

## 7. Troubleshooting

| Symptom | Fix |
|---|---|
| Pane says *Not connected* | Is `npm run dev` running? In a terminal run `opencode service status`; start the OpenCode desktop app or `opencode service start` |
| Status dot red, banner shows `401` | `opencode service get password`, then `set OPENCODE_SERVICE_PASSWORD=<value>` and restart the bridge |
| *"We can't open this add-in from localhost"* | `npx office-addin-dev-certs install`, restart Excel (see the loopback command in §3) |
| Yellow *"ADD-IN ERROR… only available during debugging from Visual Studio"* | Save your work and close **every** Excel window, then run `Start OpenCode for Excel.bat` again. The packaged starter workbook is kept in sync with `manifest.xml`, and the launcher also creates a repaired runtime copy whose embedded ID/version exactly matches the manifest; do not reopen an older starter workbook from Excel's Recent list. |
| Ribbon button missing | `npm run unsideload && npm run sideload`, then restart Excel |
| Answer never finishes | Bridge falls back to polling after 45s of silence and finishes with what arrived. Check `%USERPROFILE%\.local\share\opencode\log\opencode.log` |
| *Format sheet* says "The sheet formatter did not load" | Reload the pane (click **OpenCode** on the Home tab) or restart the launcher; the bridge must serve `/format.js` |
| *Format sheet* leaves `#` in the summary | Click **Format sheet** again — it autofits the columns after writing the summary values |
| Reply says `Model is unavailable`, `Upstream request failed`, or no reply was returned | The pane automatically retries the same turn once with an available zero-cost `opencode/*-free` model, updates the visible model selector, and explains the switch in a banner. It never auto-selects models from potentially paid providers. If no safe fallback is available, choose a model manually from the composer. |
| *"OpenCode asked to use a tool"* banner | The agent tried a tool; confirm the deny rules in §5 are in place |
| Bridge log says no service | Set `OPENCODE_SERVICE_URL` to the exact URL shown by `opencode service status` |

## 8. Limits

* **Excel on the web** cannot reach `127.0.0.1` (Local Network Access + mixed content block it). Desktop only, unless you expose the bridge through a tunnel.
* Replies of tables rely on the model returning a ```tsv / ```csv block — the quick actions ask for exactly that.
* `GET /api/event` in OpenCode v2 is volatile (no replay, may close mid-turn); the bridge reconnects and also polls `session/.../message`, so text and completion survive.
* Long selections are truncated to the *Max selected cells sent* setting, and their values are
  loaded only after the cap is applied. Workbook and selection context are each bounded at 30,000
  characters, so use `@SheetName` to give a very large tab a larger sample budget. Turn off
  **Send workbook samples** to send no sheet data at all.
* Native attachment understanding depends on the selected provider/model's supported media types.

## 9. Files

```
Start OpenCode for Excel.bat   One-click launcher (installs, starts, sideloads, opens Excel)
Remove from Excel.bat          Uninstalls the add-in from Excel
Fix Certificates.bat           Reinstalls the localhost dev certificate
manifest.xml            Office add-in manifest (task pane + Home-tab button)
bridge/server.mjs       HTTPS static server + /api proxy + SSE fan-out + polling fallback
bridge/opencode-service.mjs   Service discovery, basic auth, scratch workspace
src/taskpane.html       Sidebar markup
src/taskpane-clean.css  Minimal ChatGPT-style task-pane UI (light + dark)
src/taskpane.js         Chat, streaming, workbook context, change plans, attachments, audit
src/lib.js              Pure helpers (markdown, table parsing, A1 maths, header/column detection)
src/plan.js             Structured edit plans: parse, preview, apply, undo snapshots
src/format.js           Native Excel.js sheet formatter (table, summary, print setup, conditional formatting)
tools/mock-opencode.mjs Fake OpenCode v2 service for offline testing
tools/smoke-test.sh     End-to-end test: mock service -> bridge -> HTTP client
tools/start.mjs         The launcher behind the .bat files (plain-English output)
tools/test-lib.mjs      Unit tests for src/lib.js
tools/test-format.mjs   Formatter tests against a fake Excel object model
tools/test-pane.mjs     Pane integration tests (fake Excel + real bridge + mock OpenCode)
```

### How the agentic edit flow works

OpenCode is a coding agent, so it cannot reach into Excel directly — the task pane does. When you
ask for a workbook change, OpenCode replies with a JSON **change plan**. The executor supports cell
values and formulas (including R1C1 fills), clearing, rows/columns, sheets, rich formatting,
autofit, merge/unmerge, tables, sorting, charts, and freeze panes. Every range operation must name
the exact sheet and range it touches.

1. **Review** — the pane renders the plan as a change preview with cell-level references. Destructive operations (overwrites, deletes, inserts that shift data, merges, sorts) are unchecked by default and are never auto-applied.
2. **Approve** — check the operations you want and click **Apply checked changes**. The pane snapshots bounded before-values for cell-only operations, then applies the operations through Office.js.
3. **Audit** - every prompt, plan and apply is appended to `workspace/audit.log` and can be viewed from the settings menu.
4. **Undo** — **Undo** restores cell values and formulas, plus the cell-local formatting of a single-cell `formatRange`. It does not restore multi-cell formatting, borders, row heights, column widths, autofit, or structural/object changes; use Excel's Ctrl+Z or restore a backup for those. Native **Format sheet** discards older add-in snapshots because it can insert rows and shift cells.

## 10. Verifying without OpenCode

```bash
npm test            # all six suites; runs offline with mock services
npm run test:unit   # markdown, table parsing, A1 maths, escaping, header/column detection
npm run test:format # native formatter: table, summary, print setup, conditional formatting, idempotency
npm run test:bridge # discovery, auth, v1 + v2 protocols, resilience, hardening
npm run test:pane   # the real taskpane.js in jsdom against a fake Excel + real bridge
npm run test:starter # starter workbook manifest compatibility
npm run test:e2e    # curl-driven end-to-end script
```

`test:format` runs the real `src/format.js` against a fake Excel object model and checks the
table range, summary formulas, number formats, freeze/print settings, conditional formatting,
the array-formula clear, autofit, and a second idempotent run. `test:pane` runs the actual pane
code and asserts, among other things, that **Format sheet** calls the native formatter (not the
model), that `Insert table` writes the parsed 4×2 block at A22, and that `Use as formula`
writes the formula into the selected cell. See **TEST-REPORT.md** for what is covered, the bugs
testing caught, and what still needs a real Excel to confirm.

## 11. Where to take it next

* **Custom functions**: `=OPENCODE("explain A1")` via a SharedRuntime (async, so results arrive a turn later).
* **Ribbon actions** for one-click "Explain selection" / "Clean selection".
* **Prompt library**: ship your team's recurring analyses as quick-action chips.
* **Deployment**: publish the pane to a static site and swap `SourceLocation` to HTTPS; keep the bridge as a per-machine tray service (AppSource certification is the hurdle for listing publicly).
