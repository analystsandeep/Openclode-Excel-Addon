/* Regression test for the Excel starter workbook's embedded add-in reference. */
import fs from "node:fs";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { unzipSync } from "fflate";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const packaged = path.join(root, "OpenCode for Excel.xlsx");
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

/** Read one XML entry from an .xlsx ZIP, reporting the exact failure. */
function readZipEntry(zipPath, entryName) {
  let files;
  try {
    files = unzipSync(new Uint8Array(fs.readFileSync(zipPath)));
  } catch (err) {
    throw new Error(`could not read ${path.basename(zipPath)} as a ZIP archive: ${err.message}`);
  }
  const data = files[entryName];
  if (!data || !data.length) throw new Error(`${path.basename(zipPath)} has no readable ${entryName}`);
  return new TextDecoder("utf-8").decode(data);
}

function checkReference(label, xml, error) {
  ok(`${label} has readable webextension XML`, !error && /<\?xml version="1\.0"/.test(xml), error || xml.slice(0, 80));
  ok(`${label} reference matches manifest id`, new RegExp(`<we:reference[^>]+id="${id}"`, "i").test(xml), xml);
  ok(`${label} reference matches manifest version`, new RegExp(`<we:reference[^>]+version="${version.replace(/\./g, "\\.")}"`, "i").test(xml), xml);
}

let packagedXml = "";
let packagedError = "";
try {
  packagedXml = readZipEntry(packaged, "xl/webextensions/webextension.xml");
} catch (err) {
  packagedError = err.message;
}
checkReference("packaged starter workbook", packagedXml, packagedError);

const prepared = spawnSync(process.execPath, [path.join(root, "tools", "start.mjs"), "--prepare-starter"], {
  cwd: root,
  encoding: "utf8",
});
let result = null;
try { result = JSON.parse((prepared.stdout || "").trim()); } catch {}
ok("launcher provides a starter workbook", prepared.status === 0 && result?.ok && fs.existsSync(result.path), prepared.stderr || prepared.stdout);
if (result?.skipped) console.log(`  (runtime preparation skipped: ${result.why})`);

let preparedXml = "";
let preparedError = "";
if (result?.path) {
  try {
    preparedXml = readZipEntry(result.path, "xl/webextensions/webextension.xml");
  } catch (err) {
    preparedError = err.message;
  }
}
checkReference("prepared starter workbook", preparedXml, preparedError);

console.log(failures ? `\n${failures}/${checks} starter checks failed` : `\nAll ${checks} starter-workbook checks passed`);
process.exit(failures ? 1 : 0);
