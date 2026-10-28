# Security policy

## Reporting a security issue

Do not open a public issue for suspected credential exposure, unsafe workbook writes, path
traversal, or permission bypasses. Contact the repository owner privately through the contact
options on the GitHub profile at <https://github.com/analystsandeep>.

Include the affected version, reproduction steps, impact, and any relevant logs with credentials
removed. Please allow time for the report to be reviewed before publishing details.

## Local security model

The add-in uses a localhost HTTPS bridge and discovers the locally running OpenCode service.
OpenCode service credentials are read at runtime and are not stored in this repository. Sessions
run in a scratch workspace that denies shell and file-edit tools by default. Workbook changes are
performed through Office.js after review in the task pane.
