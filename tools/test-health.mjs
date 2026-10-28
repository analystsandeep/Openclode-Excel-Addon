/**
 * Tests tools/bridge-health.mjs - the HTTPS check used by the one-click launcher.
 * Reproduces the Windows case (Node does not trust the dev CA) without Windows.
 *
 *   node tools/test-health.mjs
 */
import fs from "node:fs";
import path from "node:path";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
import { caPath, getJson, bridgeStatus, CERT_DIR } from "./bridge-health.mjs";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const BRIDGE_PORT = process.env.TEST_BRIDGE_PORT || "3100";
const MOCK_PORT = process.env.TEST_MOCK_PORT || "28895";
const BRIDGE = `https://localhost:${BRIDGE_PORT}/api/status`;

let failures = 0;
const ok = (name, condition, detail = "") => {
  console.log(`  ${condition ? "PASS" : "FAIL"}  ${name}${condition ? "" : "  -> " + detail}`);
  if (!condition) failures++;
};
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const mock = spawn(process.execPath, [path.join(ROOT, "tools", "mock-opencode.mjs")], {
  env: { ...process.env, PORT: MOCK_PORT, MOCK_PASSWORD: "secret" },
  stdio: ["ignore", "ignore", "inherit"],
});
const bridge = spawn(process.execPath, [path.join(ROOT, "bridge", "server.mjs")], {
  env: {
    ...process.env,
    OPENCODE_SERVICE_URL: `http://127.0.0.1:${MOCK_PORT}`,
    OPENCODE_SERVICE_PASSWORD: "secret",
    PORT: BRIDGE_PORT,
    BRIDGE_SKIP_CERT_INSTALL: "1",
  },
  stdio: ["ignore", "pipe", "pipe"],
});
let bridgeLog = "";
bridge.stdout.on("data", (d) => (bridgeLog += d.toString()));
bridge.stderr.on("data", (d) => (bridgeLog += d.toString()));
const stop = () => {
  try { bridge.kill("SIGKILL"); } catch {}
  try { mock.kill("SIGKILL"); } catch {}
};

try {
  // wait for the bridge
  let up = false;
  for (let i = 0; i < 60; i++) {
    up = !!(await getJson(BRIDGE, { rejectUnauthorized: false, timeout: 1000 }));
    if (up) break;
    await sleep(500);
  }
  if (!up) {
    const exit = bridge.exitCode != null ? ` (exit ${bridge.exitCode})` : "";
    throw new Error(`bridge did not start${exit}\n${bridgeLog.trim()}`);
  }

  console.log("\n== bridge health check ==");
  const ca = caPath();
  ok("CA certificate file found for office-addin-dev-certs", !!ca && fs.existsSync(ca), String(ca));

  const withCa = await getJson(BRIDGE, { ca, rejectUnauthorized: true });
  ok("validates successfully using the generated CA", !!withCa && withCa.port === Number(BRIDGE_PORT), JSON.stringify(withCa));

  const noCaStrict = await getJson(BRIDGE, { rejectUnauthorized: true });
  console.log(`  note  without an explicit CA, strict validation ${noCaStrict ? "succeeded (CA is in the OS store - Linux/macOS)" : "failed (Windows case)"}`);

  // The leaf certificate is not a CA: using it must fail, proving we really verify.
  const leaf = path.join(CERT_DIR, "localhost.crt");
  const wrongCa = await getJson(BRIDGE, { ca: leaf, rejectUnauthorized: true });
  ok("rejects the server when the wrong CA is used", wrongCa === null, JSON.stringify(wrongCa));

  const relaxed = await getJson(BRIDGE, { rejectUnauthorized: false });
  ok("loopback fallback works when nothing is trusted", !!relaxed && relaxed.connected === true, JSON.stringify(relaxed));

  const probe = await bridgeStatus(BRIDGE);
  ok("bridgeStatus returns the status", !!probe && probe.status.connected === true, JSON.stringify(probe));
  ok("bridgeStatus reports how it validated", typeof probe.trusted === "boolean", JSON.stringify(probe.trusted));

  console.log("\n== bridge stopped ==");
  bridge.kill("SIGKILL");
  await sleep(1200);
  ok("returns null once the bridge is gone", (await bridgeStatus(BRIDGE)) === null);
} catch (err) {
  failures++;
  console.log(`  FAIL  threw: ${err.message}`);
} finally {
  stop();
  await sleep(300);
}

console.log(failures ? `\n${failures} FAILURE(S)` : "\nAll health-check tests passed");
process.exit(failures ? 1 : 0);
