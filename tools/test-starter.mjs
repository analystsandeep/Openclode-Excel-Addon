/* Regression test for the Excel starter workbook's embedded add-in reference. */
import fs from "node:fs";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const manifest = fs.readFileSync(path.join(root, "manifest.xml"), "utf8");
const id = manifest.match(/<Id>\s*([^<]+)\s*<\/Id>/i)?.[1]?.trim();
const version = manifest.match(/<Version>\s*([^<]+)\s*<\/Version>/i)?.[1]?.trim();
let failures = 0;
let checks = 0;

function ok(name, condition, detail = "") {
  checks++;
  console.log(`${condition ? "PASS" : "FAIL"}  ${name}${condition ? "" : `  ${detail}`}`);
  if (!condition) failures++;
}

const prepared = spawnSync(process.execPath, [path.join(root, "tools", "start.mjs"), "--prepare-starter"], {
  cwd: root,
  encoding: "utf8",
});
let result = null;
try { result = JSON.parse((prepared.stdout || "").trim()); } catch {}
ok("launcher provides a starter workbook", prepared.status === 0 && result?.ok && fs.existsSync(result.path), prepared.stderr || prepared.stdout);
if (result?.skipped) console.log(`  (runtime preparation skipped: ${result.why})`);

let xml = "";
if (result?.path) {
  const extracted = spawnSync("tar", ["-xOf", result.path, "xl/webextensions/webextension.xml"], { encoding: "utf8" });
  xml = extracted.stdout || "";
}
ok("XML declaration remains version 1.0", /<\?xml version="1\.0"/.test(xml), xml.slice(0, 80));
ok("embedded reference matches manifest id", new RegExp(`<we:reference[^>]+id="${id}"`, "i").test(xml), xml);
ok("embedded reference matches manifest version", new RegExp(`<we:reference[^>]+version="${version.replace(/\./g, "\\.")}"`, "i").test(xml), xml);

console.log(failures ? `\n${failures}/${checks} starter checks failed` : `\nAll ${checks} starter-workbook checks passed`);
process.exit(failures ? 1 : 0);
