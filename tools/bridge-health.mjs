/**
 * Health check for the local bridge over HTTPS.
 *
 * Why this exists: the bridge uses a self-signed developer certificate. On Linux/macOS
 * Node trusts the system store (where office-addin-dev-certs installs its CA), but on
 * Windows Node uses its own bundled CA list and rejects our certificate - so a plain
 * fetch() fails even though the bridge is running fine.
 *
 * Strategy: validate properly using the generated CA when we can find it, and only fall
 * back to "trust anything" for a loopback call to our own process.
 */

import fs from "node:fs";
import https from "node:https";
import os from "node:os";
import path from "node:path";

export const CERT_DIR = path.join(os.homedir(), ".office-addin-dev-certs");
export const CA_FILE_NAMES = ["ca.crt", "rootCA.crt", "localhost.crt"];

/** Path of the CA certificate written by office-addin-dev-certs, or null. */
export function caPath() {
  for (const name of CA_FILE_NAMES) {
    const candidate = path.join(CERT_DIR, name);
    if (fs.existsSync(candidate)) return candidate;
  }
  return null;
}

/**
 * GET JSON from an https URL.
 * @returns the parsed object, or null when the server is unreachable / not ready.
 */
export function getJson(urlString, options = {}) {
  const { ca, rejectUnauthorized = true, timeout = 5000 } = options;
  return new Promise((resolve) => {
    let url;
    try {
      url = new URL(urlString);
    } catch {
      return resolve(null);
    }
    let settled = false;
    const done = (value) => {
      if (!settled) {
        settled = true;
        resolve(value);
      }
    };
    const requestOptions = {
      host: url.hostname,
      port: url.port || 443,
      path: `${url.pathname}${url.search}`,
      method: "GET",
      headers: { accept: "application/json" },
      rejectUnauthorized,
      timeout,
    };
    if (ca && fs.existsSync(ca)) requestOptions.ca = fs.readFileSync(ca);

    const req = https.request(requestOptions, (res) => {
      let body = "";
      res.on("data", (chunk) => (body += chunk));
      res.on("end", () => {
        try {
          done(JSON.parse(body));
        } catch {
          done(null);
        }
      });
    });
    req.on("timeout", () => {
      req.destroy();
      done(null);
    });
    req.on("error", () => done(null));
    req.end();
  });
}

/** Check the bridge: tries proper CA validation first, then loopback-only fallback. */
export async function bridgeStatus(url) {
  const ca = caPath();
  const strict = await getJson(url, { ca, rejectUnauthorized: true });
  if (strict) return { status: strict, trusted: true, usedCa: ca };
  const relaxed = await getJson(url, { rejectUnauthorized: false });
  return relaxed ? { status: relaxed, trusted: false, usedCa: null } : null;
}
