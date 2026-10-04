"""
Local Python engine for consolidation_app.html.

The app itself is consolidation_app.html. It runs the uploaded notebook
in one of two ways:
  1. Local Python (this file) - fast, works offline, uses your installed
     pandas/openpyxl. Used automatically whenever this file is running.
  2. In-browser Python - used when this file is not running. Needs
     internet access to cdn.jsdelivr.net the first time each session.

How to start this file (any one of these - no .bat file needed):
  - Double-click consolidation_app.py
  - Open it in Spyder / VS Code / IDLE and press Run
  - Jupyter: paste this whole file into a notebook cell and run the cell
    (stop it with Kernel -> Interrupt)
  - Command Prompt:  python consolidation_app.py

It opens the app in your browser. Keep consolidation_app.html in the same
folder as this file, or simply open the .html yourself while this runs.

Everything stays on this computer: the page talks only to 127.0.0.1.
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

try:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    APP_DIR = os.getcwd()  # pasted into a Jupyter cell

PAGE_PATH = os.path.join(APP_DIR, "consolidation_app.html")
DEFAULT_OUTPUT_DIR = os.path.join(APP_DIR, "OUTPUT")

HOST = "127.0.0.1"
PORT = 8765  # the page looks for the engine on ports 8765-8769

INPUTS = {
    # form field -> notebook variable
    "reference": "reference_file",
    "hc": "hc_file",
    "appian": "appian_file",
}

run_lock = threading.Lock()

MISSING_PAGE = """<!doctype html><meta charset="utf-8">
<title>Consolidation engine</title>
<body style="font-family:Arial,sans-serif;margin:40px;color:#333">
<h2 style="border-left:6px solid #db0011;padding-left:12px">Consolidation engine is running</h2>
<p>consolidation_app.html was not found next to consolidation_app.py.</p>
<p>Open <b>consolidation_app.html</b> in your browser - it will connect to this engine.</p>
</body>"""


# ============================================================
# RUNNING THE NOTEBOOK
# ============================================================

def set_assignment(code, name, value):
    # Replace the line "name = ..." in section 1 of the notebook/script
    pattern = re.compile(rf"^{re.escape(name)} = .*$", re.MULTILINE)
    if not pattern.search(code):
        raise ValueError(
            f"Could not find the line '{name} = ...' in the uploaded notebook. "
            "Is this the consolidation notebook?"
        )
    return pattern.sub(lambda _: f"{name} = {value}", code, count=1)


def safe_filename(name, fallback):
    name = os.path.basename(name or "")
    name = re.sub(r"[^A-Za-z0-9._ -]", "_", name).strip()
    return name or fallback


def notebook_to_code(raw_bytes):
    # Join the code cells of a .ipynb into one script.
    # Jupyter-only lines (%magic, !shell) are commented out.
    nb = json.loads(raw_bytes.decode("utf-8"))
    cells = []
    for cell in nb.get("cells", []):
        if cell.get("cell_type") != "code":
            continue
        source = cell.get("source", "")
        if isinstance(source, list):
            source = "".join(source)
        lines = [
            "# " + line if line.lstrip().startswith(("%", "!")) else line
            for line in source.splitlines()
        ]
        cells.append("\n".join(lines))
    if not cells:
        raise ValueError("The uploaded notebook has no code cells.")
    return "\n\n".join(cells) + "\n"


def load_logic(item):
    if not item or not item.get("data"):
        raise ValueError("Missing the consolidation notebook (.ipynb).")

    name = item.get("name") or "notebook.ipynb"
    raw = base64.b64decode(item["data"])

    if name.lower().endswith(".ipynb"):
        try:
            return name, notebook_to_code(raw)
        except json.JSONDecodeError:
            raise ValueError(f"'{name}' is not a valid Jupyter notebook.")
    return name, raw.decode("utf-8-sig")


def run_consolidation(payload):
    logic_name, code = load_logic(payload.get("logic"))

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
                logic_path = os.path.join(work_dir, safe_filename(logic_name, "notebook"))
                exec(
                    compile(code, logic_path, "exec"),
                    {"__name__": "__main__", "__file__": logic_path},
                )
            except Exception:
                ok = False
                traceback.print_exc()

        outputs = []
        for name in sorted(os.listdir(run_output_dir)):
            src = os.path.join(run_output_dir, name)
            shutil.copy2(src, os.path.join(output_dir, name))
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


# ============================================================
# LOCAL WEB SERVER
# ============================================================

class Handler(BaseHTTPRequestHandler):
    def _send(self, status, body, content_type):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        # Lets consolidation_app.html talk to this engine when it is
        # opened straight from disk (file://)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def _send_json(self, status, obj):
        self._send(status, json.dumps(obj).encode("utf-8"), "application/json")

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            if os.path.exists(PAGE_PATH):
                with open(PAGE_PATH, "rb") as f:
                    self._send(200, f.read(), "text/html; charset=utf-8")
            else:
                self._send(200, MISSING_PAGE.encode("utf-8"), "text/html; charset=utf-8")
        elif self.path == "/config":
            self._send_json(200, {
                "engine": "consolidation-local",
                "default_output_dir": DEFAULT_OUTPUT_DIR,
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
            self._send_json(200, run_consolidation(payload))
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
    port = PORT
    while True:
        try:
            server = ThreadingHTTPServer((HOST, port), Handler)
            break
        except OSError:
            port += 1  # port in use, try the next one

    url = f"http://{HOST}:{port}/"
    print("Consolidation engine running at", url)
    print("Keep this window open while using the app. Close it to stop.")
    print("(In Jupyter: the cell keeps running - use Kernel -> Interrupt to stop.)")
    threading.Timer(0.5, lambda: webbrowser.open(url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
        print("Consolidation engine stopped.")


if __name__ == "__main__":
    main()
