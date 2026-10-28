# Test report

Ran on Node 24 (Windows 11). **263 automated checks across six suites, all passing**, plus
a live end-to-end verification against a real **OpenCode 2.1.4** service and the real task
pane. The automated suites run offline against mock OpenCode services on isolated ports, so
they can run while the add-in is in use.

```bash
npm test          # unit + formatter + health + bridge + pane  (~2 min)
npm run test:e2e  # original curl-driven end-to-end script
```

| Suite | What it exercises | Checks |
|---|---|---|
| `tools/test-lib.mjs` | Pure helpers: markdown rendering, table extraction, A1 maths, escaping, header/column detection, **plan parsing** | **53** |
| `tools/test-format.mjs` | Native formatter against a fake Excel model: table, summary, number formats, freeze/print, conditional formatting, idempotency | **49** |
| `tools/test-health.mjs` | Launcher health check, dev CA, strict-vs-CA validation, 503 handling | **7** |
| `tools/test-starter.mjs` | Runtime starter workbook: embedded add-in ID/version synchronization and valid XML declaration | **4** |
| `tools/test-bridge.mjs` | Real bridge over HTTP/SSE: discovery, auth, v1 + v2 protocols, models, provider errors, resilience, hardening, **attachment and audit endpoints** | **59** |
| `tools/test-pane.mjs` | The actual `src/taskpane.js` in jsdom with a fake Excel object model, a real bridge and a real SSE stream | **91** |

## How the pane is tested without Excel

`tools/test-pane.mjs` boots the mock OpenCode service **and** the real bridge, then loads
`taskpane.html` into jsdom and evaluates the real `lib.js` / `taskpane.js` with shims for
`Office`, `Excel`, `EventSource` and `fetch`. The fake Excel records every `getRange`,
`values`, `formulas` and `autofitColumns` call, so insert-into-sheet behaviour is asserted
exactly:

* insertion lands at **A22** for a selection of `Sheet1!A1:D20` (one row below, spacer row)
* `Insert table` writes the parsed `4×2` block and autofits
* `Insert ↓` writes one line per row
* `Use as formula` writes `[[=IF(COUNTIF($B$2:$B$100,B2)>1,"duplicate","")]]` into **A1**
* `Copy` puts the answer on the clipboard
* context refreshes on `onSelectionChanged`, marks large ranges *truncated*, and caps the
  TSV at `maxCells / columnCount` rows
* `Send selection` off removes the selection from the prompt
* Stop finalises the turn; New chat clears the transcript; a dead bridge raises a banner
  **and** an in-transcript error
* the **model dropdown** is populated from `/api/models`, groups by provider, persists the
  choice in `localStorage`, and sends `{ providerID, id }` with the next prompt
* an unavailable or empty model turn is retried exactly once with the preferred free
  `opencode` fallback; the same prompt is reused, the selector is updated, and a paid model
  present in the mock catalog is never selected automatically
* an empty turn renders the provider's error (or the "no reply was returned" hint) instead
  of `_No output._`
* the **Format sheet** chip calls the native formatter instead of sending a model prompt,
  shows a success banner, and restores its label afterwards
* the **agentic plan flow**: a mock plan answer renders the change preview with per-operation
  checkboxes (destructive ops unchecked by default), applying executes the operations through
  the fake Excel model, **Undo** restores the before-snapshot, and read-only mode disables Apply
* **attachments**: attached files are listed as chips and their names/content flow into the prompt
* **@sheet mentions**: `@Sheet2` focuses that sheet's data in the prompt
* **settings persistence**: model choice, read-only mode and approval gates survive in localStorage

## Format sheet formatter

`tools/test-format.mjs` loads the real `src/format.js` with a fake Excel object model and asserts:

* a raw sheet gets cleaned headers in row 4, `tblSales` at `A4:H16`, and exactly three summary rows inserted
* the summary formulas use structured references and the non-CSE top-region formula
* `D2` / `D3` / `H2` are cleared before formulas are written, so an old single-cell array formula cannot block the update
* number formats: `dd-mmm-yyyy`, `#,##0`, `$#,##0;[Red]-$#,##0`
* `TableStyleMedium2`, banded rows on / columns off, navy header, freeze 4 rows, gridlines off, navy tab
* print area `A1:H16`, print titles `$1:$4`, landscape, narrow margins, fit to one page wide
* two conditional formats (data bar on the amount column + red refund rule), and the title merge only on the first run
* columns are autofit after the summary is written (Office.js `columnWidth` is not in Excel's character unit)
* a second run reuses `tblSales`, does not insert rows, duplicate tables, merge again or add conditional formats
* an existing table starting at row 1 is restyled without a summary above it
* a header-only sheet throws a friendly "no data rows" error

## Bridge scenarios covered

1. Discovery from the **registration file** (`~/.local/state/opencode/service.json`, password accepted, no 401)
2. Discovery from the **service config** (port + hostname + password only)
3. Discovery from a **default port**
4. **401 recovery** through a stubbed `opencode service get password`
5. **No service** → `/api/status` reports 503 with an actionable hint, session creation 503
6. **OpenCode v1** end-to-end: `/global/health`, `/session`, `/session/:id/message`, `/event`, `/abort`
7. **Resilience**: turn completes even when `/api/event` is killed mid-turn, full text recovered by polling
8. **Agent override rejected** → automatic retry without the agent, turn still completes
9. **HTTP hardening**: unknown route 404, missing field 400, malformed JSON 400, path traversal blocked, correct MIME types, still healthy afterwards
10. **Two clients** on one session receive identical text
11. **Models**: `/api/models` returns enabled models with `id` + `providerID` + `name`, hides disabled ones, rejects an unknown model, and switches the session model before prompting
12. **Provider failure**: a real v2 `error` object is surfaced on the `done` event with no invented text
13. **Audit log**: `POST /api/audit` appends JSONL to `workspace/audit.log`; `GET /api/audit` returns the recent entries
14. **Native attachments**: upload, filename sanitization, invalid-base64 rejection, and removal

## Live verification (real OpenCode 2.1.4, real pane)

Verified on the user's machine against `http://127.0.0.1:28894` (v2.1.4):

* `/api/models` returned **57 models**, grouped by provider in the pane
* two real turns, switching model per turn through `POST /api/session/{id}/model`:
  * `opencode/deepseek-v4.1-flash-free` → `LIVE-DEEPSEEK-V4.1-FLASH-FREE`
  * `opencode/mimo-v2.6-flash-free` → `LIVE-MIMO-V2.6-FLASH-FREE`
* each answer came from the newest assistant message's `content[]` **text** parts only;
  `reasoning` parts were present on the second turn and were correctly excluded
* the assistant message reported the model that was requested, proving the switch took effect
* the pane itself (in a browser against the live bridge) showed the model dropdown, the
  selected DeepSeek model, and the real reply `PANE LIVE OK` — no `_No output._`
* the transient `503 Upstream request failed` from `opencode/flint-40b-free` now appears as
  an error in the transcript instead of an empty answer; the pane now retries that class of
  failure once with a safe free `opencode` model
* **Format sheet** was clicked in the live pane on `SaleData`:
  * the bridge first returned **404** for `/format.js` because the static whitelist was missing it; fixed and covered by `test-bridge`
  * the first live run produced 1–2 character columns and `#` in the summary; the formatter now calls `autofitColumns()` instead of setting Office.js `columnWidth`
  * final run: one `tblSales` at `A4:H47`, 43 data rows, summary `B2=Jan 2018 - Dec 2019`, `D2=43`, `F2=2,121`, `H2=$1,305,676`, `B3=$30,365`, `D3=Central`, `F3=11`, `H3=5`
  * `D3` is now a normal formula (no CSE), freeze at row 4, gridlines off, print area `A1:H47`, print titles `$1:$4`, 2 conditional-format rules, navy/white/bold header
  * a repeated click is idempotent: same table, same row count, no duplicate summary or conditional formats

## Bugs found and fixed by testing

| # | Bug | Fix |
|---|---|---|
| 1 | Action buttons closed over `state.answer`, so *Copy* / *Insert* on an older message used the newest answer | Capture the text per message at render time |
| 2 | `data-code="…"` was not quote-escaped, so a code block containing `"` could break out of the attribute | Added `escapeAttr()` and a regression test |
| 3 | `OPENCODE_SERVICE_PASSWORD` was only used for `OPENCODE_SERVICE_URL`, so a service discovered via config file or default port got **401** | Env password now applies to every candidate, including the 401-recovery path |
| 4 | A turn only finished when the bridge *saw* `session.idle` on SSE — with a dropped stream (a known v2 issue) or a missed event the pane hung for 45 s+ | Completion now also fires when the assistant message is marked completed or the answer holds steady across 8 polls, plus explicit timeout and "produced nothing" guards |
| 5 | Clicking a quick-action chip **discarded the chip's instruction** when the user had typed something | Chip instruction + typed question are now combined |
| 6 | The bridge never discovered OpenCode **v1** (it only probed `/api/info`) and wasted its first SSE attempt on a 404 | Added `/global/health` to the health probes and learn the event path (`/api/event` vs `/event`) during discovery |
| 7 | A dead bridge failed silently in the transcript only | Connection errors now raise a banner **and** an in-transcript message |
| 8 | **Every v2 answer rendered `_No output._`** — the parser only read v1 `parts[]`, but v2 returns `content[]` with `type: "text"` / `"reasoning"` | Parse `content[]`, keep only text parts, and select the newest assistant message created after the prompt |
| 9 | An empty v2 turn looked like a successful empty answer, hiding the provider error | Surface `error` / `retry` from the assistant message and SSE; render the message in the pane |
| 10 | No way to choose a model; every session kept OpenCode's default | Added `/api/models`, switch via `POST /api/session/{id}/model` before each prompt, and a persisted ⚙ → *Model* dropdown |
| 11 | The bridge's static whitelist did not include `/format.js`, so the chip appeared but the formatter was a 404 | Added `/format.js` to the static map and a bridge test that serves it |
| 12 | Office.js `Range.format.columnWidth` produced 1–2 character columns, so summary numbers rendered `#` | Replaced fixed widths with `autofitColumns()` after the summary is written |
| 13 | `D3` held a single-cell array formula; writing a normal formula over it could fail with "part of an array" | Clear the summary formula cell (`clear("Contents")`) before setting the new formula |
| 14 | Undo snapshots were started inside an un-awaited async `forEach`, so apply could begin before the snapshot existed | Snapshot sheets sequentially and await every `Excel.run` |
| 15 | Plans advertised `{at,count}` row/column edits but generated no executable range | Normalize them to exact row/column ranges before preview |
| 16 | Worksheet context passed `Sheet!A1:B2` to `Worksheet.getRange`, which expects a local address | Build local ranges from the used-range row/column indexes |
| 17 | Binary attachments were filename-only and the initial file picker was not reachable | Add a visible Attach button and OpenCode V2 native file URI flow |

## Not covered (needs your machine)

* **Excel itself** — real `Excel.run` semantics, WebView2 rendering, and the ribbon button.
  This was verified manually on the target machine (button on the Home tab, pane opens,
  prompt/reply work), and the **Format sheet** path has now been run live (table, summary,
  print setup, conditional formatting, idempotency), but it is not part of `npm test`.
* **Sideloading** on Windows and the `office-addin-dev-certs` trust prompt
* **Model quality** — whether the answers actually produce clean ```tsv blocks for `Insert table`

If discovery fails on your PC, send the bridge's startup output plus `opencode service status`
and `opencode service get password`; the discovery order is documented in README §6.
