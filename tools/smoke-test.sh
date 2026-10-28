#!/usr/bin/env bash
# End-to-end smoke test: mock OpenCode service -> bridge -> HTTP client.
# Usage: bash tools/smoke-test.sh
set -u
cd "$(dirname "$0")/.."
ROOT="$PWD"

mkdir -p /tmp/mockhome/.local/state /tmp/mockhome/.config

echo "== starting mock OpenCode service on 28894 =="
PORT=28894 MOCK_PASSWORD=secret node tools/mock-opencode.mjs > /tmp/mock.log 2>&1 &
MOCK_PID=$!
sleep 1.5

echo "== starting bridge on 3000 =="
HOME=/tmp/mockhome \
XDG_STATE_HOME=/tmp/mockhome/.local/state \
XDG_CONFIG_HOME=/tmp/mockhome/.config \
OPENCODE_SERVICE_URL=http://127.0.0.1:28894 \
OPENCODE_SERVICE_PASSWORD=secret \
PORT=3000 node bridge/server.mjs > /tmp/bridge.log 2>&1 &
BRIDGE_PID=$!
sleep 3

cleanup() {
  kill "$BRIDGE_PID" 2>/dev/null
  kill "$MOCK_PID" 2>/dev/null
  wait "$BRIDGE_PID" 2>/dev/null
  wait "$MOCK_PID" 2>/dev/null
}
trap cleanup EXIT

echo "== /api/status =="
curl -sk https://localhost:3000/api/status; echo

echo "== /api/agents =="
curl -sk https://localhost:3000/api/agents; echo

echo "== static taskpane =="
curl -sk -o /dev/null -w "taskpane.html -> HTTP %{http_code}\n" https://localhost:3000/taskpane.html
curl -sk -o /dev/null -w "taskpane.js   -> HTTP %{http_code}\n" https://localhost:3000/taskpane.js
curl -sk -o /dev/null -w "icon-32.png   -> HTTP %{http_code}\n" https://localhost:3000/assets/icon-32.png
curl -sk -o /dev/null -w "lib.js        -> HTTP %{http_code}\n" https://localhost:3000/lib.js

echo "== POST /api/session =="
SID=$(curl -sk -X POST https://localhost:3000/api/session -H "content-type: application/json" -d '{}' \
  | python3 -c 'import sys,json;print(json.load(sys.stdin)["sessionID"])')
echo "sessionID=$SID"

echo "== POST /api/prompt =="
curl -sk -X POST https://localhost:3000/api/prompt -H "content-type: application/json" \
  -d "{\"sessionID\":\"$SID\",\"text\":\"flag duplicates in column B\"}"; echo

echo "== GET /api/events (SSE, 25s max) =="
curl -sk -N --max-time 25 "https://localhost:3000/api/events?sessionID=$SID" > /tmp/events.log

echo "-- event summary --"
python3 - "$SID" <<'PY'
import json, sys
types = {}
last_text = ""
done_text = None
for line in open("/tmp/events.log"):
    line = line.strip()
    if not line.startswith("data:"):
        continue
    try:
        evt = json.loads(line[5:].strip())
    except Exception:
        continue
    t = evt.get("type")
    types[t] = types.get(t, 0) + 1
    if t == "text":
        last_text = evt.get("text", "")
    if t == "done":
        done_text = evt.get("text", "")
print("event counts:", types)
final = done_text if done_text is not None else last_text
print("final answer length:", len(final or ""))
print("----- answer -----")
print(final)
PY

echo
echo "== bridge log =="
cat /tmp/bridge.log
