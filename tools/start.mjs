/**
 * One-click launcher (run by "Start OpenCode for Excel.bat").
 *
 *   1. installs the code dependencies (first run only, a few seconds)
 *   2. starts the local bridge
 *   3. checks that your OpenCode service is answering
 *   4. installs the add-in into Excel (registry -> no extra downloads)
 *   5. opens Excel and stays open - closing this window stops everything
 */

import fs from "node:fs";
import path from "node:path";
import readline from "node:readline";
import { spawn, spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { bridgeStatus } from "./bridge-health.mjs";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const isWin = process.platform === "win32";
const BRIDGE = "https://localhost:3000";
const MANIFEST = path.join(ROOT, "manifest.xml");
const STARTER = path.join(ROOT, "OpenCode for Excel.xlsx");
const WEF_KEY = "HKCU\\Software\\Microsoft\\Office\\16.0\\Wef\\Developer";
const LINE = "------------------------------------------------------------";

const say = (msg = "") => console.log(msg);
const title = (msg) => say(`\n${LINE}\n  ${msg}\n${LINE}`);
const step = (n, msg) => say(`\n[${n}] ${msg}`);
const good = (msg) => say(`    OK   ${msg}`);
const warn = (msg) => say(`    (!)  ${msg}`);
const bad = (msg) => say(`    (X)  ${msg}`);

function ask(question) {
  if (!process.stdin.isTTY) return Promise.resolve("");
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  return new Promise((resolve) =>
    rl.question(question, (answer) => {
      rl.close();
      resolve(answer);
    })
  );
}

/**
 * npm/npx are .cmd shims on Windows, and Node cannot spawn those directly -
 * doing so fails with ENOENT. Running their JavaScript entry point with the
 * same node binary avoids the shell (and its argument-quoting problems).
 */
function npmRunner(name) {
  const cli = path.join(path.dirname(process.execPath), "node_modules", "npm", "bin", `${name}-cli.js`);
  if (fs.existsSync(cli)) return { command: process.execPath, prefix: [cli], options: {} };
  // Unusual Node builds may not ship npm next to node.exe - use the shim.
  return { command: name, prefix: [], options: isWin ? { shell: true } : {} };
}

function run(command, args, options = {}) {
  const result = spawnSync(command, args, { cwd: ROOT, stdio: "inherit", ...options });
  return result.status === 0;
}

/** Run a command and keep its output, so failures can explain themselves. */
function runCapture(command, args, options = {}) {
  const result = spawnSync(command, args, {
    cwd: ROOT,
    stdio: "pipe",
    encoding: "utf8",
    ...options,
  });
  return {
    ok: result.status === 0,
    status: result.status,
    error: result.error,
    out: `${result.stdout || ""}${result.stderr || ""}`.trim(),
  };
}

function excelIsRunning() {
  if (!isWin) return false;
  const result = spawnSync("tasklist", ["/FI", "IMAGENAME eq EXCEL.EXE", "/NH"], {
    stdio: "pipe", encoding: "utf8", windowsHide: true,
  });
  return result.status === 0 && /\bEXCEL\.EXE\b/i.test(result.stdout || "");
}



/* ------------------------------------------------------------- sideload ---- */

function manifestId() {
  try {
    const match = fs.readFileSync(MANIFEST, "utf8").match(/<Id>\s*([^<]+)\s*<\/Id>/i);
    return match ? match[1].trim() : null;
  } catch {
    return null;
  }
}

function manifestVersion() {
  try {
    const match = fs.readFileSync(MANIFEST, "utf8").match(/<Version>\s*([^<]+)\s*<\/Version>/i);
    return match ? match[1].trim() : null;
  } catch {
    return null;
  }
}

/**
 * The embedded webextension reference must use the exact manifest version.
 * Opening a workbook with a stale version makes Excel display the misleading
 * "only available during debugging from Visual Studio" error. Build a fresh
 * runtime copy on every launch so manifest upgrades can never stale the bundled
 * starter workbook (and so an already-open source workbook cannot block repair).
 */
function prepareStarterWorkbook() {
  if (!isWin || !fs.existsSync(STARTER)) return { ok: false, why: "starter workbook is missing" };
  const id = manifestId();
  const version = manifestVersion();
  if (!id || !version) return { ok: false, why: "manifest id or version could not be read" };

  const base = process.env.LOCALAPPDATA || process.env.TEMP || ROOT;
  const dir = path.join(base, "OpenCodeExcel", "launcher");
  try {
    fs.mkdirSync(dir, { recursive: true });
    // Remove old unlocked runtime copies. Excel may still hold one; leave it and
    // use this process-specific filename instead.
    for (const name of fs.readdirSync(dir)) {
      if (!/^OpenCode for Excel-.*\.xlsx$/i.test(name)) continue;
      try { fs.unlinkSync(path.join(dir, name)); } catch { /* currently open in Excel */ }
    }
    const target = path.join(dir, `OpenCode for Excel-${version}-${process.pid}.xlsx`);
    fs.copyFileSync(STARTER, target);

    const script = [
      "Add-Type -AssemblyName System.IO.Compression",
      "Add-Type -AssemblyName System.IO.Compression.FileSystem",
      "$zip=[System.IO.Compression.ZipFile]::Open($env:OC_STARTER_TARGET,[System.IO.Compression.ZipArchiveMode]::Update)",
      "try {",
      "  $entry=$zip.GetEntry('xl/webextensions/webextension.xml')",
      "  if ($null -eq $entry) { throw 'starter workbook has no webextension reference' }",
      "  $reader=New-Object System.IO.StreamReader($entry.Open())",
      "  try { $xml=$reader.ReadToEnd() } finally { $reader.Dispose() }",
      "  $xml=[regex]::Replace($xml,'(<we:reference\\b[^>]*\\bversion=\")[^\"]+(\")','${1}'+$env:OC_MANIFEST_VERSION+'${2}')",
      "  $xml=[regex]::Replace($xml,'(<we:reference\\b[^>]*\\bid=\")[^\"]+(\")','${1}'+$env:OC_MANIFEST_ID+'${2}')",
      "  $entry.Delete()",
      "  $new=$zip.CreateEntry('xl/webextensions/webextension.xml',[System.IO.Compression.CompressionLevel]::Optimal)",
      "  $writer=New-Object System.IO.StreamWriter($new.Open(),(New-Object System.Text.UTF8Encoding($false)))",
      "  try { $writer.Write($xml) } finally { $writer.Dispose() }",
      "} finally { $zip.Dispose() }",
    ].join("; ");
    const result = spawnSync("powershell.exe", ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", script], {
      cwd: ROOT,
      stdio: "pipe",
      encoding: "utf8",
      env: { ...process.env, OC_STARTER_TARGET: target, OC_MANIFEST_ID: id, OC_MANIFEST_VERSION: version },
    });
    if (result.status !== 0) {
      try { fs.unlinkSync(target); } catch {}
      return { ok: false, why: `${result.stderr || result.stdout || "PowerShell could not update it"}`.trim() };
    }
    return { ok: true, path: target, version };
  } catch (err) {
    return { ok: false, why: err.message };
  }
}

/**
 * Register the add-in with Excel the same way office-addin-dev-settings does:
 * one registry value under HKCU\...\Wef\Developer (value name = manifest id).
 * Avoids downloading the 500-package Microsoft tooling.
 */
function sideloadRegistry() {
  const id = manifestId();
  if (!id) return { ok: false, why: "could not read the add-in id from manifest.xml" };

  const write = spawnSync(
    "reg",
    ["add", WEF_KEY, "/v", id, "/t", "REG_SZ", "/d", MANIFEST, "/f"],
    { stdio: "pipe", encoding: "utf8" }
  );
  if (write.status !== 0) {
    return { ok: false, why: (write.stderr || write.stdout || "reg add failed").trim() };
  }

  // Verify it really landed.
  const read = runCapture("reg", ["query", WEF_KEY, "/v", id]);
  if (read.ok && read.out.includes(id)) return { ok: true, id };
  return { ok: false, why: "the registry entry could not be read back" };
}

function sideload() {
  if (!isWin) return { ok: false, why: "not Windows" };
  const direct = sideloadRegistry();
  if (direct.ok) return direct;

  // Fall back to Microsoft's tool (downloads on demand).
  say("        trying Microsoft's sideload tool instead...");
  const npx = npmRunner("npx");
  if (
    run(
      npx.command,
      [...npx.prefix, "--yes", "office-addin-dev-settings", "sideload", "manifest.xml"],
      npx.options
    )
  ) {
    return { ok: true, id: manifestId(), via: "npx" };
  }
  return { ok: false, why: direct.why };
}

function manualSteps() {
  say("        Do this once by hand:");
  say("        1. Copy manifest.xml into a folder, for example C:\\OpenCodeAddin");
  say("        2. Excel > File > Options > Trust Center > Trust Center Settings");
  say("           > Trusted Add-in Catalogs > add C:\\OpenCodeAddin and tick 'Show in Menu'");
  say("        3. Restart Excel > Insert > My Add-ins > Shared Folder > OpenCode for Excel");
}

/* ----------------------------------------------------------------- main ---- */

async function main() {
  title("OpenCode for Excel");
  say("  This window must stay open while you use the add-in in Excel.");
  say("  To stop everything, close this window or press Ctrl + C.\n");

  // 1. Node ------------------------------------------------------------------
  step(1, "Checking Node.js");
  const major = Number(process.versions.node.split(".")[0]);
  if (major < 18) {
    bad(`Node.js ${process.version} is too old. Please install Node.js 20 or newer from https://nodejs.org`);
    await ask("Press Enter to close.");
    process.exit(1);
  }
  good(`Node.js ${process.version}`);

  // 2. Dependencies ---------------------------------------------------------
  step(2, "Preparing the add-in");
  if (fs.existsSync(path.join(ROOT, "node_modules", "office-addin-dev-certs"))) {
    good("already prepared");
  } else {
    say("        Downloading components (a few seconds)...");
    const npm = npmRunner("npm");
    const install = runCapture(
      npm.command,
      [...npm.prefix, "install", "--omit=dev", "--no-audit", "--no-fund"],
      npm.options
    );
    if (!install.ok) {
      bad("The download failed.");
      if (install.error) {
        bad(`Could not run npm: ${install.error.message}`);
      } else if (install.out) {
        // Show the tail of npm's own message - it names the real cause.
        for (const line of install.out.split(/\r?\n/).slice(-12)) say(`        ${line}`);
      }
      say("");
      say("        If the messages above mention the network or a proxy, check your");
      say("        internet connection and run this again.");
      await ask("Press Enter to close.");
      process.exit(1);
    }
    good("done");
  }

  // 3. Is OpenCode around? --------------------------------------------------
  step(3, "Looking for OpenCode on this PC");
  const cli = spawnSync(isWin ? "opencode --version" : "opencode", isWin ? [] : ["--version"], {
    shell: isWin,
    stdio: "ignore",
  });
  if (cli.error) {
    warn("The 'opencode' command is not on your PATH - fine if you use the OpenCode desktop app.");
  } else {
    good("opencode command found");
  }

  // Excel reads the Wef\Developer registration only when the Excel process
  // starts. Registering while an old Excel process is open produces the same
  // misleading "Visual Studio debugging" error as a stale workbook version.
  if (isWin && excelIsRunning()) {
    step("3b", "Waiting for Excel to close");
    warn("Excel is already running and cannot reload the add-in registration.");
    warn("Save your work, then close every Excel window.");
    await ask("        When Excel is fully closed, press Enter to continue: ");
    if (excelIsRunning()) {
      bad("Excel is still running (it may have a hidden window or unsaved workbook).");
      say("        Open Task Manager, end Microsoft Excel only after saving your work,");
      say("        then double-click Start OpenCode for Excel.bat again.");
      await ask("Press Enter to close.");
      process.exit(1);
    }
    good("Excel is closed and ready for a clean add-in load");
  }

  // 4. Start the bridge -----------------------------------------------------
  step(4, "Starting the local bridge on https://localhost:3000");
  const child = spawn(process.execPath, [path.join("bridge", "server.mjs")], {
    cwd: ROOT,
    stdio: ["ignore", "pipe", "pipe"],
  });
  let childOut = "";
  const echo = (data) => {
    const text = data.toString();
    childOut += text;
    process.stdout.write(text);
  };
  child.stdout.on("data", echo);
  child.stderr.on("data", echo);
  child.on("exit", (code) => {
    if (code !== 0 && code !== null) {
      say("");
      bad(`The bridge stopped (exit code ${code}).`);
      if (/EADDRINUSE|address already in use|EACCES/i.test(childOut)) {
        warn("Port 3000 is already in use or blocked - the add-in is probably already");
        warn("running in another window. Use that one, or close it and run this again.");
      }
      ask("Press Enter to close.").then(() => process.exit(code ?? 1));
    }
  });

  let probe = await waitForBridge(90000);
  if (!probe) {
    bad("The bridge did not answer. The messages above should explain why.");
    await ask("Press Enter to close.");
    process.exit(1);
  }
  if (!probe.trusted) {
    warn("Windows does not trust the developer certificate yet, so this check skipped");
    warn("certificate validation. Excel uses its own trust, which is already set up.");
  }
  good("bridge is running");

  // 5. OpenCode connection --------------------------------------------------
  step(5, "Connecting to OpenCode");
  let status = probe.status;
  let attempts = 0;
  while (!status.connected && attempts < 3) {
    bad(status.hint || "OpenCode is not answering yet.");
    say("");
    say("        Open the OpenCode desktop app now, wait a few seconds,");
    say("        then press Enter to check again.");
    const answer = (await ask("        Press Enter to retry (or type q and Enter to continue anyway): ")).trim();
    if (answer.toLowerCase() === "q") break;
    attempts += 1;
    say("        Checking...");
    const next = await waitForConnected(20000);
    status = next ? next.status : status;
  }
  if (status.connected) {
    good(`connected to OpenCode ${status.version ? "v" + status.version : ""} at ${status.url}`);
  } else {
    warn("Still not connected - starting Excel anyway. The sidebar shows 'Not connected'");
    warn("until OpenCode is running; it re-checks by itself every 20 seconds.");
  }

  // 6. Sideload -------------------------------------------------------------
  step(6, "Installing the add-in into Excel");
  const installed = sideload();
  if (installed.ok) {
    good(`installed${installed.via === "npx" ? " (via Microsoft tool)" : ""}`);
  } else {
    warn(`Automatic install did not work (${installed.why}).`);
    manualSteps();
  }

  // 7. Excel ----------------------------------------------------------------
  step(7, "Opening Excel");
  if (isWin) {
    // Excel 2021 on this build ignores the registry entry alone, but it loads the
    // add-in (and its Home-tab button) when a workbook containing a webextension
    // reference is opened. The starter workbook does exactly that.
    const preparedStarter = prepareStarterWorkbook();
    const args = ["/c", "start", "", "excel"];
    if (preparedStarter.ok) args.push(preparedStarter.path);
    spawn(process.env.COMSPEC || "cmd.exe", args, {
      stdio: "ignore",
      windowsHide: true,
    });
    if (preparedStarter.ok) {
      good(`Excel is opening a repaired starter workbook (add-in v${preparedStarter.version})`);
    } else {
      good("Excel is opening");
      warn(`The starter workbook could not be prepared (${preparedStarter.why}).`);
      warn("Use Insert > My Add-ins > Developer Add-ins > OpenCode for Excel.");
    }
  } else {
    warn("This is not Windows - open Excel yourself.");
  }

  title("Ready");
  say('  Excel opens with the "OpenCode for Excel" workbook and the OpenCode panel.');
  say("  The OpenCode button is on the Home tab. You can close that workbook and");
  say("  keep using the button in your own workbooks - it stays until Excel closes.");
  say("  Select some cells, type a question, and press Send.");
  say("");
  say("  Keep this window open. Closing it stops the add-in.\n");

  const stopBridge = () => {
    try {
      child.kill();
    } catch {}
  };
  process.on("exit", stopBridge);
  process.on("SIGINT", () => {
    say("\nStopping...");
    stopBridge();
    process.exit(0);
  });
  process.on("SIGTERM", () => {
    stopBridge();
    process.exit(0);
  });
  process.on("SIGHUP", () => {
    stopBridge();
    process.exit(0);
  });
  process.on("SIGBREAK", () => {
    stopBridge();
    process.exit(0);
  }); // Windows: console window closed

  await new Promise(() => {}); // stay alive; the bridge keeps printing here
}

async function waitForBridge(timeoutMs) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    const probe = await bridgeStatus(`${BRIDGE}/api/status`);
    if (probe && probe.status && probe.status.port) return probe;
    await new Promise((r) => setTimeout(r, 500));
  }
  return null;
}

async function waitForConnected(timeoutMs) {
  const deadline = Date.now() + timeoutMs;
  let last = null;
  while (Date.now() < deadline) {
    const probe = await bridgeStatus(`${BRIDGE}/api/status`);
    if (probe) {
      last = probe;
      if (probe.status.connected) return probe;
    }
    await new Promise((r) => setTimeout(r, 1000));
  }
  return last;
}

if (process.argv.includes("--prepare-starter")) {
  const result = prepareStarterWorkbook();
  console.log(JSON.stringify(result));
  process.exit(result.ok ? 0 : 1);
} else {
  main().catch(async (err) => {
    bad(err.message);
    await ask("Press Enter to close.");
    process.exit(1);
  });
}
