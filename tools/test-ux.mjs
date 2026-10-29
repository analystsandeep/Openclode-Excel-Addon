/* Static UX contract checks for the narrow Excel task pane. */

import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const html = fs.readFileSync(path.join(ROOT, "src", "taskpane.html"), "utf8");
const css = fs.readFileSync(path.join(ROOT, "src", "taskpane-clean.css"), "utf8");
const js = fs.readFileSync(path.join(ROOT, "src", "taskpane.js"), "utf8");

let passed = 0;
function check(name, fn) {
  try {
    fn();
    passed++;
    console.log("PASS  " + name);
  } catch (error) {
    console.error("FAIL  " + name + "\n      " + error.message);
    process.exitCode = 1;
  }
}

check("Enter sends while Shift+Enter remains a newline", () => {
  assert.match(js, /event\.key === "Enter" && !event\.shiftKey && !event\.isComposing/);
});

check("composer placeholder fits a narrow pane", () => {
  assert.match(html, /Ask anything — Enter to send · \/ commands · @ranges/);
});

check("banner is a dismissible live flow element before the composer", () => {
  assert.match(html, /id="banner"[^>]*aria-live="polite"[\s\S]*id="btnDismissBanner"/);
  assert.ok(html.indexOf('id="banner"') < html.indexOf('<footer class="composer">'));
  assert.doesNotMatch(css.slice(css.indexOf(".banner {"), css.indexOf(".banner.err")), /position:\s*fixed/);
});

check("common actions have a visible chip rail", () => {
  const rail = html.slice(html.indexOf('<nav class="quick-rail"'), html.indexOf("</nav>", html.indexOf('<nav class="quick-rail"')));
  for (const action of ["format", "explain", "fix", "clean"]) assert.match(rail, new RegExp(`data-action="${action}"`));
});

check("duplicated attachment and session buttons are removed", () => {
  assert.doesNotMatch(html, /btnAttachTop|btnNewSession/);
  assert.match(html, /id="btnTools"[\s\S]*id="toolsMenu"|id="toolsMenu"[\s\S]*id="btnTools"/);
});

check("single-session chrome collapses", () => {
  assert.match(js, /sessionRecords\.length <= 1/);
});

check("narrow flex children and code blocks can shrink or scroll", () => {
  assert.match(css, /\.session-tab[^}]*min-width:\s*0/s);
  assert.match(css, /\.op-row > \* \{ min-width:\s*0/);
  assert.match(css, /\.bubble pre[^}]*overflow-x:\s*auto/s);
  assert.doesNotMatch(css, /\.op-badge\s*\{\s*display:\s*none/);
});

check("offline Send gives feedback and opens connection settings", () => {
  const send = js.slice(js.indexOf("async function send()"), js.indexOf("function buildPrompt"));
  assert.match(send, /Not connected to OpenCode/);
  assert.match(send, /openSettings\("connectionSection"\)/);
});

check("plan JSON is suppressed in favor of the review panel", () => {
  assert.match(js, /review panel[\s\S]{0,220}msg\.innerHTML = ""/);
  assert.match(js, /Apply " \+ checked \+ "\/" \+ boxes\.length \+ " checked"/);
  assert.match(js, /aria-label", "Apply " \+ checked \+ " of "/);
  assert.match(css, /\.preview-panel[^}]*display:\s*flex[^}]*flex-direction:\s*column/s);
  assert.match(css, /\.preview-list[^}]*overflow-y:\s*auto/s);
  assert.match(html, /id="previewMore"[\s\S]*id="btnShowPlanJson"[\s\S]*id="btnPrepareBackup"/);
  assert.match(css, /#app\.reviewing \.quick-rail \{ display:\s*none/);
  assert.match(js, /classList\.add\("reviewing"\)/);
  assert.match(js, /classList\.remove\("reviewing"\)/);
});

check("settings is an accessible keyboard dialog", () => {
  assert.match(html, /id="settingsPanel"[^>]*role="dialog"[^>]*aria-modal="true"/);
  assert.match(js, /settingsFocusable\(\)/);
  assert.match(js, /event\.key === "Escape"/);
  assert.match(css, /\.settings-head[^}]*position:\s*sticky/s);
});

check("composer tools menu is keyboard operable and cannot silently shrink", () => {
  assert.match(html, /id="btnTools"[^>]*aria-controls="toolsMenu"/);
  assert.match(js, /function setToolsMenu/);
  assert.match(js, /toolsMenu"\)\.addEventListener\("keydown"/);
  assert.match(js, /event\.key !== "ArrowDown" && event\.key !== "ArrowUp"/);
  assert.match(css, /\.tools-trigger \{[^}]*flex:\s*0 0 auto/s);
});

check("command palette stays clear of the header in a short pane", () => {
  assert.match(css, /\.command-menu \{[^}]*bottom:\s*calc\(100% \+ 3px\)[^}]*max-height:\s*min\(240px,\s*45vh\)/s);
});

check("empty state names the visible workbook tabs", () => {
  assert.match(js, /state\.context\.sheets/);
  assert.match(js, /escapeHtml\(workbookLine\)/);
});

check("status is persistent and connection details are visible", () => {
  assert.match(html, /id="connectionStatus"/);
  assert.match(html, /id="connectionDetails"/);
  assert.match(js, /headerStatusText/);
});

check("readability, focus, contrast, and reduced motion have explicit rules", () => {
  assert.match(css, /font:\s*14px/);
  assert.match(css, /:focus-visible/);
  assert.match(css, /#input::placeholder \{ color: #757575/);
  assert.match(css, /prefers-reduced-motion/);
  assert.doesNotMatch(css, /font-size:\s*(?:9|10|10\.5|11|11\.5)px/);
});

check("security and optional pane-width enhancements are present", () => {
  assert.match(html, /Content-Security-Policy/);
  assert.match(js, /isSetSupported\("TaskPaneApi", "1\.1"\)/);
  assert.match(js, /Office\.TaskPane\.setWidth\(360\)/);
  assert.equal(fs.existsSync(path.join(ROOT, "src", "taskpane.css")), false);
});

console.log(`\n${passed} UX checks passed`);
