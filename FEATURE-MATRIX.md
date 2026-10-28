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
- OpenCode Skills/playbooks and visibility into configured MCP/app connections.
- Web-research prompts through the tools enabled in the OpenCode workspace.

### Reviewable workbook changes

- JSON change plans with exact sheets/ranges, assumptions, operation counts, and affected-cell
  estimates.
- Per-operation approval, select-all, read-only mode, small-change auto-approval, and explicit
  never-ask mode.
- Values, formulas, R1C1 formula fills, clearing, row/column insertion and deletion, sheet add /
  rename / delete, formatting and borders, autofit, merge/unmerge, tables, sorting, basic charts,
  and freeze panes.
- Before-value/formula/number-format/basic-format snapshots, an eight-entry undo stack, Excel
  Ctrl+Z compatibility, and local JSONL audit records.
- Direct answer/table/formula insertion now uses the same snapshot, audit, and undo path.

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
- **Undo of structural objects:** cell content and common formatting are snapshotted. Complex
  objects such as deleted charts/tables and every sheet-level setting still rely on Excel's own
  undo behavior and should be reviewed carefully.
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

1. **Retrieval for large workbooks** — workbook index, on-demand range reads, dependency graph, and
   deterministic formula-audit tools instead of relying only on prompt samples.
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
