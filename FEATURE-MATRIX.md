# OpenCode for Excel — implementation matrix

This document separates features that work in the local add-in from capabilities that require a
hosted product, Microsoft tenant administration, or future Office.js work. It is intentionally
conservative: a feature is marked complete only when the add-in has an executable path for it.

## Available in v1.2

### Workbook agent

- Multi-turn OpenCode sessions with provider/model and Agent selection.
- Workbook-scoped session persistence for saved files.
- Bounded cross-sheet context containing values, table names, and live formulas.
- `@SheetName` focus, selected-range context, and cell/range citations in prompts.
- Explicit user-directed `@range(Sheet!A1:B20)` reads, bounded to three ranges, 500 cells
  and 12,000 characters per message. Model-requested range reads use an optional structured
  protocol and require approval before data is read/sent; at most two requests per turn.
- Explicit bounded `@search("text")` over used-range values/formulas (20 sheets / 5,000 cells /
  50 results) and direct `@trace(Sheet!A1)` relationships on supported Excel hosts; no external
  workbook tracing or complete dependency graph.
- OpenCode Skills/playbooks and visibility into configured MCP/app connections.
- Web-research prompts through the tools enabled in the OpenCode workspace.

### Reviewable workbook changes

- JSON change plans with exact sheets/ranges, assumptions, operation counts, and affected-cell
  estimates.
- Per-operation approval, select-all, and read-only mode. Automatic approval is limited to finite,
  non-destructive, cell-local formatting; border, row-height, column-width, and autofit formatting
  always require review. Destructive, structural, and object operations are never auto-applied.
- Supported cell-only plans get a bounded before/proposed cell preview and a pre-apply
  drift check. Supported undo checks post-apply state before restoring. These are non-atomic
  checks, not protection against every concurrent edit; mixed/structural/unsupported plans and
  automatic formatting do not get the pre-apply guard.
- Values, formulas, R1C1 formula fills, clearing, row/column insertion and deletion, sheet add /
  rename / delete, formatting and borders, autofit, merge/unmerge, tables, sorting, basic charts,
  and freeze panes.
- Bounded before-value/formula snapshots for cell writes and before-formatting snapshots for
  single-cell, cell-local `formatRange`, with an eight-entry undo stack. Multi-cell formatting,
  borders, row heights, column widths, autofit, merge, sort, clear, row/column, sheet, table, and
  chart operations are not add-in-undoable.
- Bounded post-write readback verifies supported exact writes and flags Excel error value types;
  relative fills, multiple same-sheet writes and unsupported plans are explicitly unverified.
- Manual local workbook backup download in plan review where Office supports compressed export,
  capped at 50 MB. Verify the copy and extension; no automatic restore. For structural edits,
  use Excel's own Undo where available or restore a verified backup. Changes are recorded in local
  JSONL audit records.
- Direct answer/table/formula insertion uses the same snapshot and audit path for the cell writes.

### Analysis and inputs

- Formula explanation/fixing, workbook audit, variance, scenario, dashboard, cleaning, and formula
  quick actions.
- Native V2 file attachments: up to 20 files, 25 MB each, stored only in the local scratch
  workspace. Text content is also injected as a fallback.
- Markdown, tables, code copy, live formula insertion, stop/cancel, model errors, and retries.
- Native reusable sheet formatter with tables, summary formulas, print setup, and conditional
  formatting.

## Partial / host-dependent

- **Large workbooks:** all tabs are represented, but samples are bounded. `@SheetName` increases
  the focused tab's budget; retrieval/chunking is still needed for very large models.
- **PDF/image/Office understanding:** files reach OpenCode natively, but extraction depends on the
  chosen provider/model and its media support.
- **Web citations and MCP tools:** available only when the OpenCode workspace has an approved
  search provider or MCP/app connection and its permissions allow use.
- **Undo of structural objects:** the add-in deliberately does not offer Undo for row/column,
  sheet, table, chart, merge, sort, or clear operations because it cannot completely reverse them.
  Use Excel's own Undo or a backup for those changes. Cell-level snapshots are capped at 5,000
  cells, unknown or named ranges are never auto-approved or snapshotted, and native **Format
  sheet** discards older snapshots because it can insert rows and shift cells.
- **Charts:** common chart creation/deletion and positioning are supported; advanced styling and
  specialized chart types are not.
- **VBA and PivotTables:** the assistant can explain or generate instructions/code, but the add-in
  does not execute VBA or create native PivotTables.

## Requires a separate enterprise service

These are not properties a localhost Office add-in can honestly provide by itself:

- Microsoft 365 centralized deployment and group assignment.
- Organization RBAC, SAML SSO, SCIM provisioning, tenant-wide audit logs, usage billing/credits.
- DLP/SIEM exports, compliance APIs, legal holds, regional processing guarantees, customer-managed
  encryption keys, and organization retention policies.
- Marketplace/AppSource certification and globally hosted task-pane/bridge infrastructure.

OpenCode provider, model, MCP, and permission policy still applies locally, but it is not a
replacement for the administration plane above.

## Recommended next phases

1. **Retrieval for large workbooks** — scalable workbook index, larger searches, dependency graph,
   and deterministic formula-audit tools beyond the bounded reads and direct traces.
2. **Complete reversible transactions** — snapshot/restore tables, charts, merged areas, sheet
   visibility, row/column dimensions, conditional formats, and structural edits.
3. **Document ingestion verification** — integration tests against PDF, DOCX, PPTX, XLSX, and image
   attachments across supported model providers.
4. **Native Excel objects** — PivotTables/PivotCharts, named ranges, data validation, slicers,
   Power Query, and Office Scripts.
5. **Production packaging** — signed installer or tray bridge, hosted static assets, telemetry with
   consent, update channel, and AppSource/internal-catalog deployment.
6. **Enterprise control plane** — authentication, tenancy, RBAC, policy, audit export, retention,
   and compliance integrations.
