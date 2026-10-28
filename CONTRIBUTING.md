# Contributing

Contributions that improve Excel compatibility, workbook safety, OpenCode integration, or test
coverage are welcome.

## Development setup

1. Fork and clone the repository.
2. Install Node.js 18 or later.
3. Run `npm install`.
4. Run `npm test` before making changes.
5. Start the local bridge with `npm run dev` when testing the task pane.

## Pull requests

- Keep changes focused on one problem.
- Explain the user-facing behavior and any compatibility tradeoffs.
- Add or update automated tests for behavior changes.
- Do not commit API keys, OpenCode service passwords, certificates, workbook data, or generated
  `node_modules` content.
- Confirm that `npm test` passes.

Workbook operations should remain reviewable and reversible where Office.js allows it. New write
operations must identify exact sheets and ranges and should be covered by plan parsing and pane
integration tests.
