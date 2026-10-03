"""
Local web app for the Reference Data / HC / Appian consolidation.

How to start it (any one of these):
  - Double-click consolidation_app.py
  - Open it in Spyder / VS Code / IDLE and press Run
  - Command Prompt:  python consolidation_app.py

Keep reference_data_consolidation.py in the same folder as this file.

A browser page opens where you choose the three input files and click Run.
Everything stays on this computer: the page talks only to this local
program (127.0.0.1), which runs reference_data_consolidation.py on the
uploaded files and returns the output workbooks.

Needs only the packages the script already uses (pandas, numpy, openpyxl).
"""

import base64
import contextlib
import io
import json
import os
import re
import shutil
import tempfile
import threading
import traceback
import webbrowser

from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

APP_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT_PATH = os.path.join(APP_DIR, "reference_data_consolidation.py")
DEFAULT_OUTPUT_DIR = os.path.join(APP_DIR, "OUTPUT")

HOST = "127.0.0.1"
PORT = 8765

INPUTS = {
    # form field -> script variable
    "reference": "reference_file",
    "hc": "hc_file",
    "appian": "appian_file",
}

run_lock = threading.Lock()


# ============================================================
# WEB PAGE (kept inside this file so only two files are needed)
# ============================================================

PAGE_HTML = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Reference Data Consolidation</title>
<style>
  :root {
    --bg: #f4f5f7;
    --surface: #ffffff;
    --border: #dcdfe4;
    --text: #1d2228;
    --muted: #5f6b7a;
    --brand: #db0011;
    --brand-dark: #a8000d;
    --ok: #1a7f37;
    --ok-bg: #e8f5ec;
    --warn: #9a6700;
    --err: #c62828;
    --err-bg: #fdecec;
    --drop: #fff4f5;
    --log-bg: #1e2329;
    --log-text: #d8dee6;
  }
  @media (prefers-color-scheme: dark) {
    :root {
      --bg: #15181c;
      --surface: #1f2329;
      --border: #343a42;
      --text: #e6e9ed;
      --muted: #9aa5b1;
      --ok: #5cc27a;
      --ok-bg: #1b2c21;
      --warn: #e0b341;
      --err: #ff7b72;
      --err-bg: #3a1f1f;
      --drop: #2a1d1f;
      --log-bg: #0f1215;
    }
  }
  * { box-sizing: border-box; }
  body {
    margin: 0;
    background: var(--bg);
    color: var(--text);
    font: 15px/1.5 "Segoe UI", system-ui, -apple-system, Arial, sans-serif;
  }
  header {
    background: var(--brand);
    color: #fff;
    padding: 18px 24px;
  }
  header h1 { margin: 0; font-size: 20px; font-weight: 600; }
  header p { margin: 4px 0 0; opacity: .9; font-size: 13px; }
  main { max-width: 920px; margin: 0 auto; padding: 24px 16px 48px; }
  section {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 20px;
    margin-bottom: 18px;
  }
  h2 { margin: 0 0 14px; font-size: 16px; font-weight: 600; }
  h2 .step {
    display: inline-flex; align-items: center; justify-content: center;
    width: 22px; height: 22px; margin-right: 8px;
    border-radius: 50%; background: var(--brand); color: #fff;
    font-size: 12px; vertical-align: 1px;
  }
  .files { display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 12px; }
  .drop {
    position: relative;
    border: 2px dashed var(--border);
    border-radius: 8px;
    padding: 14px;
    cursor: pointer;
    transition: border-color .15s, background .15s;
    min-height: 104px;
  }
  .drop:hover, .drop.over { border-color: var(--brand); background: var(--drop); }
  .drop.has-file { border-style: solid; border-color: var(--ok); background: var(--ok-bg); }
  .drop input { position: absolute; inset: 0; opacity: 0; cursor: pointer; }
  .drop .label { font-weight: 600; }
  .drop .hint { color: var(--muted); font-size: 13px; }
  .drop .file { margin-top: 8px; font-size: 13px; word-break: break-all; color: var(--ok); font-weight: 600; }
  .options label { display: flex; gap: 10px; align-items: flex-start; margin-bottom: 10px; cursor: pointer; }
  .options input[type=checkbox] { margin-top: 4px; accent-color: var(--brand); }
  .options small { display: block; color: var(--muted); }
  .field { margin-top: 6px; }
  .field label { display: block; font-weight: 600; margin-bottom: 4px; }
  .field input {
    width: 100%; padding: 8px 10px; font: inherit; font-size: 14px;
    border: 1px solid var(--border); border-radius: 6px;
    background: var(--bg); color: var(--text);
  }
  .run-row { display: flex; align-items: center; gap: 16px; flex-wrap: wrap; }
  button {
    font: inherit; font-weight: 600; cursor: pointer;
    border: none; border-radius: 6px; padding: 10px 22px;
  }
  .primary { background: var(--brand); color: #fff; font-size: 16px; padding: 12px 36px; }
  .primary:hover:not(:disabled) { background: var(--brand-dark); }
  button:disabled { opacity: .45; cursor: not-allowed; }
  .secondary { background: transparent; color: var(--text); border: 1px solid var(--border); padding: 8px 16px; }
  .secondary:hover { border-color: var(--brand); color: var(--brand); }
  #status { color: var(--muted); }
  #status.ok { color: var(--ok); font-weight: 600; }
  #status.err { color: var(--err); font-weight: 600; }
  .spinner {
    display: inline-block; width: 16px; height: 16px; vertical-align: -3px; margin-right: 6px;
    border: 2px solid var(--border); border-top-color: var(--brand); border-radius: 50%;
    animation: spin .8s linear infinite;
  }
  @keyframes spin { to { transform: rotate(360deg); } }
  .outputs { list-style: none; margin: 0; padding: 0; }
  .outputs li {
    display: flex; align-items: center; justify-content: space-between; gap: 12px;
    padding: 10px 0; border-bottom: 1px solid var(--border);
  }
  .outputs li:last-child { border-bottom: none; }
  .outputs .name { word-break: break-all; }
  .outputs a {
    flex-shrink: 0; text-decoration: none; font-weight: 600; font-size: 14px;
    color: var(--brand); border: 1px solid var(--brand); border-radius: 6px; padding: 6px 14px;
  }
  .outputs a:hover { background: var(--brand); color: #fff; }
  .saved { color: var(--muted); font-size: 13px; margin: 10px 0 0; word-break: break-all; }
  pre#log {
    margin: 0; max-height: 360px; overflow: auto;
    background: var(--log-bg); color: var(--log-text);
    border-radius: 6px; padding: 12px 14px;
    font: 12.5px/1.55 Consolas, "Cascadia Mono", Menlo, monospace;
    white-space: pre-wrap; word-break: break-word;
  }
  pre#log .w { color: #f2cc60; }
  pre#log .e { color: #ff8a80; }
  .hidden { display: none; }
  .banner { padding: 10px 14px; border-radius: 6px; background: var(--err-bg); color: var(--err); margin-bottom: 18px; }
</style>
</head>
<body>
<header>
  <h1>Reference Data / Headcount / Appian Consolidation</h1>
  <p>Files are processed on this computer only. Nothing is uploaded to the internet.</p>
</header>

<main>
  <div id="offline" class="banner hidden">
    The app program isn't running. Start <b>consolidation_app.py</b> and reload this page.
  </div>

  <section>
    <h2><span class="step">1</span>Choose input files</h2>
    <div class="files">
      <div class="drop" data-field="reference">
        <input type="file" accept=".xlsx,.xlsm">
        <div class="label">Reference Data Hierarchy</div>
        <div class="hint">Needs Config, APPIAN_CONFIG and New_Migrations sheets</div>
        <div class="file"></div>
      </div>
      <div class="drop" data-field="hc">
        <input type="file" accept=".xlsx,.xlsm">
        <div class="label">HC Current Month</div>
        <div class="hint">Click or drop the headcount file</div>
        <div class="file"></div>
      </div>
      <div class="drop" data-field="appian">
        <input type="file" accept=".xlsx,.xlsm">
        <div class="label">Appian</div>
        <div class="hint">Click or drop the Appian extract</div>
        <div class="file"></div>
      </div>
    </div>
  </section>

  <section class="options">
    <h2><span class="step">2</span>Options</h2>
    <label>
      <input type="checkbox" id="reset_hiring_flg">
      <span>Reset HIRING_FLG every run
        <small>Clears old YES values before re-deriving them</small></span>
    </label>
    <label>
      <input type="checkbox" id="drop_removed">
      <span>Drop removed positions from the Reference output
        <small>They are always listed in the Exception Report either way</small></span>
    </label>
    <div class="field">
      <label for="output_dir">Also save outputs to this folder</label>
      <input type="text" id="output_dir" spellcheck="false">
    </div>
  </section>

  <section>
    <h2><span class="step">3</span>Run</h2>
    <div class="run-row">
      <button id="run" class="primary" disabled>Run</button>
      <span id="status">Choose all three files to enable Run.</span>
    </div>
  </section>

  <section id="results" class="hidden">
    <h2>Output files</h2>
    <ul id="outputs" class="outputs"></ul>
    <p id="saved" class="saved"></p>
  </section>

  <section id="log-section" class="hidden">
    <div class="run-row" style="justify-content: space-between; margin-bottom: 10px;">
      <h2 style="margin: 0;">Run log</h2>
      <button id="copy-log" class="secondary">Copy log</button>
    </div>
    <pre id="log"></pre>
  </section>
</main>

<script>
const files = { reference: null, hc: null, appian: null };
const runBtn = document.getElementById("run");
const statusEl = document.getElementById("status");
let busy = false;

function setStatus(text, cls = "", spinning = false) {
  statusEl.className = cls;
  statusEl.innerHTML = (spinning ? '<span class="spinner"></span>' : "");
  statusEl.appendChild(document.createTextNode(text));
}

function refreshRunButton() {
  const ready = Object.values(files).every(Boolean);
  runBtn.disabled = !ready || busy;
  if (!busy && !ready) setStatus("Choose all three files to enable Run.");
  else if (!busy && ready && !statusEl.classList.contains("ok") && !statusEl.classList.contains("err"))
    setStatus("Ready to run.");
}

function pickFile(drop, file) {
  if (!file) return;
  if (!/\.(xlsx|xlsm)$/i.test(file.name)) {
    alert(`"${file.name}" is not an .xlsx / .xlsm file.`);
    return;
  }
  files[drop.dataset.field] = file;
  drop.classList.add("has-file");
  drop.querySelector(".file").textContent = file.name;
  setStatus("");
  refreshRunButton();
}

document.querySelectorAll(".drop").forEach(drop => {
  const input = drop.querySelector("input");
  input.addEventListener("change", () => pickFile(drop, input.files[0]));
  drop.addEventListener("dragover", e => { e.preventDefault(); drop.classList.add("over"); });
  drop.addEventListener("dragleave", () => drop.classList.remove("over"));
  drop.addEventListener("drop", e => {
    e.preventDefault();
    drop.classList.remove("over");
    pickFile(drop, e.dataTransfer.files[0]);
  });
});

function toBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result.split(",", 2)[1]);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function base64ToBlob(b64) {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

function showLog(text) {
  const log = document.getElementById("log");
  log.innerHTML = "";
  text.split("\n").forEach(line => {
    const span = document.createElement("span");
    if (/^(WARNING|Note:)/.test(line)) span.className = "w";
    if (/Traceback|Error|skipped/i.test(line)) span.className = "e";
    span.textContent = line + "\n";
    log.appendChild(span);
  });
  document.getElementById("log-section").classList.remove("hidden");
}

function showOutputs(outputs, outputDir) {
  const list = document.getElementById("outputs");
  list.querySelectorAll("a").forEach(a => URL.revokeObjectURL(a.href));
  list.innerHTML = "";
  outputs.forEach(o => {
    const li = document.createElement("li");
    const name = document.createElement("span");
    name.className = "name";
    name.textContent = o.name;
    const a = document.createElement("a");
    a.href = URL.createObjectURL(base64ToBlob(o.data));
    a.download = o.name;
    a.textContent = "Download";
    li.append(name, a);
    list.appendChild(li);
  });
  document.getElementById("saved").textContent =
    outputs.length ? `Also saved to: ${outputDir}` : "";
  document.getElementById("results").classList.toggle("hidden", !outputs.length);
}

runBtn.addEventListener("click", async () => {
  busy = true;
  refreshRunButton();
  document.getElementById("results").classList.add("hidden");
  setStatus("Reading files...", "", true);

  try {
    const payload = {
      reset_hiring_flg: document.getElementById("reset_hiring_flg").checked,
      drop_removed: document.getElementById("drop_removed").checked,
      output_dir: document.getElementById("output_dir").value,
    };
    for (const [field, file] of Object.entries(files)) {
      payload[field] = { name: file.name, data: await toBase64(file) };
    }

    setStatus("Running consolidation... this can take a minute for large files.", "", true);
    const started = Date.now();
    const res = await fetch("/run", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const result = await res.json();
    const secs = Math.round((Date.now() - started) / 1000);

    showLog(result.log || "(no output)");
    showOutputs(result.outputs || [], result.output_dir);

    if (result.ok) {
      setStatus(`Completed in ${secs}s - ${result.outputs.length} file(s) created.`, "ok");
    } else {
      setStatus("Run failed - see the run log below.", "err");
    }
  } catch (err) {
    setStatus("Could not reach the app program. Is the app window still open?", "err");
    showLog(String(err));
  } finally {
    busy = false;
    refreshRunButton();
  }
});

document.getElementById("copy-log").addEventListener("click", () => {
  navigator.clipboard.writeText(document.getElementById("log").textContent);
});

fetch("/config")
  .then(r => r.json())
  .then(cfg => { document.getElementById("output_dir").value = cfg.default_output_dir; })
  .catch(() => document.getElementById("offline").classList.remove("hidden"));
</script>
</body>
</html>
"""


def set_assignment(code, name, value):
    # Replace the line "name = ..." in section 1 of the script
    pattern = re.compile(rf"^{re.escape(name)} = .*$", re.MULTILINE)
    if not pattern.search(code):
        raise ValueError(f"Could not find '{name} = ...' in the script")
    return pattern.sub(lambda _: f"{name} = {value}", code, count=1)


def safe_filename(name, fallback):
    name = os.path.basename(name or "")
    name = re.sub(r"[^A-Za-z0-9._ -]", "_", name).strip()
    return name or fallback


def run_consolidation(payload):
    with open(SCRIPT_PATH, encoding="utf-8") as f:
        code = f.read()

    output_dir = (payload.get("output_dir") or "").strip() or DEFAULT_OUTPUT_DIR
    output_dir = os.path.abspath(os.path.expanduser(output_dir))
    os.makedirs(output_dir, exist_ok=True)

    work_dir = tempfile.mkdtemp(prefix="consolidation_")
    run_output_dir = os.path.join(work_dir, "OUTPUT")
    os.makedirs(run_output_dir)

    try:
        for field, var in INPUTS.items():
            item = payload.get(field)
            if not item or not item.get("data"):
                raise ValueError(f"Missing input file: {field}")

            path = os.path.join(
                work_dir, f"{field}_{safe_filename(item.get('name'), field + '.xlsx')}"
            )
            with open(path, "wb") as f:
                f.write(base64.b64decode(item["data"]))

            code = set_assignment(code, var, repr(path))

        code = set_assignment(code, "output_folder", repr(run_output_dir))
        code = set_assignment(
            code, "RESET_HIRING_FLG", str(bool(payload.get("reset_hiring_flg")))
        )
        code = set_assignment(
            code, "DROP_REMOVED_POSITIONS", str(bool(payload.get("drop_removed")))
        )

        log = io.StringIO()
        ok = True
        with contextlib.redirect_stdout(log), contextlib.redirect_stderr(log):
            try:
                exec(compile(code, SCRIPT_PATH, "exec"), {"__name__": "__main__"})
            except Exception:
                ok = False
                traceback.print_exc()

        outputs = []
        for name in sorted(os.listdir(run_output_dir)):
            src = os.path.join(run_output_dir, name)
            dest = os.path.join(output_dir, name)
            shutil.copy2(src, dest)
            with open(src, "rb") as f:
                outputs.append({
                    "name": name,
                    "data": base64.b64encode(f.read()).decode("ascii"),
                })

        return {
            "ok": ok,
            # Show the real save folder, not the temporary run folder
            "log": log.getvalue().replace(run_output_dir, output_dir),
            "outputs": outputs,
            "output_dir": output_dir,
        }
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


class Handler(BaseHTTPRequestHandler):
    def _send(self, status, body, content_type):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _send_json(self, status, obj):
        self._send(status, json.dumps(obj).encode("utf-8"), "application/json")

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            self._send(200, PAGE_HTML.encode("utf-8"), "text/html; charset=utf-8")
        elif self.path == "/config":
            self._send_json(200, {
                "default_output_dir": DEFAULT_OUTPUT_DIR,
                "script_found": os.path.exists(SCRIPT_PATH),
            })
        else:
            self._send(404, b"Not found", "text/plain")

    def do_POST(self):
        if self.path != "/run":
            self._send(404, b"Not found", "text/plain")
            return

        if not run_lock.acquire(blocking=False):
            self._send_json(409, {"ok": False, "log": "A run is already in progress."})
            return

        try:
            length = int(self.headers.get("Content-Length", 0))
            payload = json.loads(self.rfile.read(length))
            result = run_consolidation(payload)
            self._send_json(200, result)
        except Exception:
            self._send_json(200, {
                "ok": False,
                "log": traceback.format_exc(),
                "outputs": [],
            })
        finally:
            run_lock.release()

    def log_message(self, fmt, *args):
        pass  # keep the console quiet


def main():
    if not os.path.exists(SCRIPT_PATH):
        print(f"ERROR: {SCRIPT_PATH} not found - keep it in the same folder as this app.")
        input("Press Enter to close...")
        return

    port = PORT
    while True:
        try:
            server = ThreadingHTTPServer((HOST, port), Handler)
            break
        except OSError:
            port += 1  # port in use, try the next one

    url = f"http://{HOST}:{port}/"
    print("Consolidation app running at", url)
    print("Keep this window open while using the app. Close it to stop.")
    threading.Timer(0.5, lambda: webbrowser.open(url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
