"""
Local web app for the Reference Data / HC / Appian consolidation.

Run:   python consolidation_app.py      (or double-click Run_Consolidation_App.bat)

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
PAGE_PATH = os.path.join(APP_DIR, "consolidation_app.html")
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
            with open(PAGE_PATH, "rb") as f:
                self._send(200, f.read(), "text/html; charset=utf-8")
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
