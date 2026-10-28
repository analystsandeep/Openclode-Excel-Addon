/**
 * OpenCode service discovery + auth helpers.
 *
 * OpenCode v2 runs a shared background service that registers itself on disk:
 *   ~/.local/state/opencode/service.json   ->  { id, version, url, pid, password }
 *   ~/.config/opencode/service.json        ->  { port, hostname, password, ... } (service config)
 * On Windows the shared service default port is 28894 (49374 on macOS/Linux).
 *
 * Docs: https://opencode.ai/v2/docs/troubleshooting/   https://opencode.ai/v2/docs/api
 */

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawn } from "node:child_process";

const isWin = process.platform === "win32";
const home = os.homedir();

const stateHome = () => process.env.XDG_STATE_HOME || path.join(home, ".local", "state");
const configHome = () => process.env.XDG_CONFIG_HOME || path.join(home, ".config");

export const REGISTRATION_FILES = [
  path.join(stateHome(), "opencode", "service.json"),
  path.join(home, ".local", "share", "opencode", "service.json"),
  ...(isWin
    ? [
        path.join(process.env.LOCALAPPDATA || "", "opencode", "service.json"),
        path.join(process.env.APPDATA || "", "opencode", "service.json"),
      ]
    : []),
];

export const CONFIG_FILES = [path.join(configHome(), "opencode", "service.json")];

/**
 * Shared service default ports: Windows uses 28894, everything else 49374.
 * OPENCODE_SERVICE_DEFAULT_PORTS can pin the probe list (used by the test suite
 * so a real OpenCode service on the default port cannot leak into tests).
 */
export const DEFAULT_PORTS = process.env.OPENCODE_SERVICE_DEFAULT_PORTS
  ? process.env.OPENCODE_SERVICE_DEFAULT_PORTS.split(",").map(Number).filter(Boolean)
  : isWin
    ? [28894, 49374]
    : [49374, 28894];

function readJson(file) {
  try {
    const raw = fs.readFileSync(file, "utf8").trim();
    if (!raw) return null;
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

/** Best effort: ask the CLI for the service password / url. Never throws. */
function askCli(args, timeoutMs = 8000) {
  return new Promise((resolve) => {
    try {
      // On Windows, pass one command string: spawn(file, args, {shell:true}) is deprecated
      // (DEP0190) and prints a warning in the user's console.
      const child = isWin
        ? spawn(`opencode ${args.join(" ")}`, { shell: true, windowsHide: true, stdio: ["ignore", "pipe", "pipe"] })
        : spawn("opencode", args, { windowsHide: true, stdio: ["ignore", "pipe", "pipe"] });
      let out = "";
      let done = false;
      const finish = (value) => {
        if (done) return;
        done = true;
        resolve(value);
      };
      child.stdout.on("data", (d) => (out += d.toString()));
      child.on("error", () => finish(null));
      child.on("close", () => finish(out.trim() || null));
      setTimeout(() => {
        try {
          child.kill();
        } catch {}
        finish(out.trim() || null);
      }, timeoutMs);
    } catch {
      resolve(null);
    }
  });
}

function normalizeUrl(value) {
  if (!value || typeof value !== "string") return null;
  let url = value.trim().replace(/\/+$/, "");
  if (!/^https?:\/\//i.test(url)) url = `http://${url}`;
  try {
    new URL(url);
    return url;
  } catch {
    return null;
  }
}

/**
 * Collect every plausible (url, password) candidate, best first.
 * Returns { candidates: [{url, password, source}], password, sources }
 */
export async function collectCandidates() {
  const seen = new Set();
  const candidates = [];
  const sources = [];
  // Password supplied by the environment is used for every discovered candidate,
  // not just for OPENCODE_SERVICE_URL.
  const envPassword = process.env.OPENCODE_SERVICE_PASSWORD || process.env.OPENCODE_PASSWORD || null;

  const push = (url, password, source) => {
    const normalized = normalizeUrl(url);
    if (!normalized) return;
    const key = `${normalized}|${password || ""}`;
    if (seen.has(key)) return;
    seen.add(key);
    candidates.push({ url: normalized, password: password || null, source });
  };

  // 1. Environment overrides (handy for WSL / remote servers / testing).
  push(process.env.OPENCODE_SERVICE_URL, envPassword, "env");

  // 2. Registration file written by the running service (has url + password + pid).
  let passwordFromDisk = null;
  for (const file of REGISTRATION_FILES) {
    const reg = readJson(file);
    if (!reg) continue;
    sources.push(`registration: ${file}`);
    if (reg.password) passwordFromDisk = reg.password;
    push(reg.url, reg.password, `registration (${path.basename(path.dirname(file))})`);
  }

  // 3. Service config file (has port/hostname/password).
  for (const file of CONFIG_FILES) {
    const cfg = readJson(file);
    if (!cfg) continue;
    sources.push(`config: ${file}`);
    if (cfg.password) passwordFromDisk = cfg.password;
    const port = cfg.port || (isWin ? 28894 : 49374);
    const host = cfg.hostname || "127.0.0.1";
    push(`http://${host}:${port}`, cfg.password || envPassword, "service config");
  }

  // 4. Well-known default ports.
  const fallbackPassword = passwordFromDisk || envPassword;
  for (const port of DEFAULT_PORTS) {
    push(`http://127.0.0.1:${port}`, fallbackPassword, "default port");
  }

  return { candidates, password: fallbackPassword, sources };
}

/** Health endpoints: v2 is /api/info, v1 is /global/health. */
const HEALTH_PATHS = ["/api/info", "/global/health"];

/** Probe one candidate with optional basic auth. */
async function probe(url, password) {
  const headers = { accept: "application/json" };
  if (password) headers.authorization = `Basic ${Buffer.from(`opencode:${password}`).toString("base64")}`;
  let last = 0;
  for (const path of HEALTH_PATHS) {
    try {
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), 4000);
      const res = await fetch(`${url}${path}`, { headers, signal: controller.signal });
      clearTimeout(timer);
      if (res.status === 200) {
        const info = await res.json().catch(() => ({}));
        return { ok: true, info, status: 200, path };
      }
      last = res.status;
    } catch (err) {
      last = 0;
    }
  }
  return { ok: false, info: null, status: last };
}

/**
 * Find a healthy OpenCode service.
 * Returns { ok, url, password, info, source, hint }
 */
export async function discoverService() {
  const { candidates, password, sources } = await collectCandidates();

  for (const candidate of candidates) {
    const result = await probe(candidate.url, candidate.password);
    if (result.ok) {
      return {
        ok: true,
        url: candidate.url,
        password: candidate.password,
        info: result.info,
        source: candidate.source,
        eventPath: await probeEventPath(candidate.url, candidate.password),
        hint: null,
      };
    }
    // Reachable but auth-rejected: try to recover the password from the CLI.
    if (result.status === 401) {
      const cliPassword = (await askCli(["service", "get", "password"])) || candidate.password || password;
      if (cliPassword) {
        const retry = await probe(candidate.url, cliPassword);
        if (retry.ok) {
          return {
            ok: true,
            url: candidate.url,
            password: cliPassword,
            info: retry.info,
            source: `${candidate.source} + opencode service get password`,
            eventPath: await probeEventPath(candidate.url, cliPassword),
            hint: null,
          };
        }
      }
      return {
        ok: false,
        url: candidate.url,
        password: cliPassword,
        info: null,
        source: candidate.source,
        hint: "OpenCode answered 401 Unauthorized. Run `opencode service get password` and set OPENCODE_SERVICE_PASSWORD, or pin one with `opencode service set password <value>`.",
      };
    }
  }

  // Nothing answered: try the CLI for a url we may have missed.
  const status = await askCli(["service", "status"]);
  const urlFromCli = status ? normalizeUrl((status.match(/https?:\/\/[^\s"']+/) || [])[0]) : null;
  if (urlFromCli) {
    const cliPassword = (await askCli(["service", "get", "password"])) || password;
    const result = await probe(urlFromCli, cliPassword);
    if (result.ok) {
      return {
        ok: true,
        url: urlFromCli,
        password: cliPassword,
        info: result.info,
        source: "opencode service status",
        eventPath: await probeEventPath(urlFromCli, cliPassword),
        hint: null,
      };
    }
  }

  return {
    ok: false,
    url: null,
    password,
    info: null,
    source: null,
    hint:
      "No OpenCode service answered on " +
      candidates.map((c) => c.url).join(", ") +
      ". Start it with the OpenCode desktop app, or run `opencode service start` (v2) / `opencode serve --port 4096` (v1).",
  };
}

/** Which SSE path does this server speak? v2: /api/event, v1: /event. */
export async function probeEventPath(url, password) {
  for (const pathname of ["/api/event", "/event"]) {
    try {
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), 3000);
      const res = await fetch(`${url}${pathname}`, {
        headers: { ...authHeaders(password), accept: "text/event-stream" },
        signal: controller.signal,
      });
      clearTimeout(timer);
      const good = res.ok && res.body;
      try {
        await res.body?.cancel(); // never leave a dangling stream behind
      } catch {
        /* already closed */
      }
      if (good) return pathname;
    } catch {
      /* try the next path */
    }
  }
  return "/api/event";
}

export function authHeaders(password) {
  return password
    ? { authorization: `Basic ${Buffer.from(`opencode:${password}`).toString("base64")}` }
    : {};
}

/** Where the add-in keeps its scratch project (sessions run here, never in your real repo). */
export function workspaceDir() {
  const base = isWin
    ? process.env.LOCALAPPDATA || path.join(home, "AppData", "Local")
    : process.env.XDG_DATA_HOME || path.join(home, ".local", "share");
  return path.join(base, "OpenCodeExcel", "workspace");
}

const RESTRICTED_CONFIG = `{
  // Scratch workspace used by the OpenCode Excel add-in.
  // The agent must never touch your files or run shell commands from Excel.
  "$schema": "https://opencode.ai/config.json",
  "permissions": [
    { "action": "shell", "resource": "*", "effect": "deny" },
    { "action": "edit", "resource": "*", "effect": "deny" }
  ]
}
`;

/** Create the scratch workspace + a locked-down opencode.jsonc so the agent stays read-only. */
export function ensureWorkspace() {
  const dir = workspaceDir();
  try {
    fs.mkdirSync(dir, { recursive: true });
    const configFile = path.join(dir, "opencode.jsonc");
    if (!fs.existsSync(configFile)) fs.writeFileSync(configFile, RESTRICTED_CONFIG, "utf8");
    return dir;
  } catch (err) {
    console.warn(`[workspace] could not prepare ${dir}: ${err.message}`);
    return process.cwd();
  }
}
