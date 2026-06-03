"""
server.py — SheetMind local bridge for the Office.js task-pane add-in.

The Office add-in is a web UI rendered inside Excel; it can't call Python
directly. This tiny zero-dependency (stdlib only) server:
  • serves the task-pane assets (taskpane.html / styles.css / app.js / logos)
  • exposes a JSON API that runs the existing SheetMind engine:
        POST /api/plan       interpret a command + dry-run → preview (uses the LLM)
        POST /api/apply      commit a previously-built plan → new grid (no LLM)
        POST /api/transcribe raw audio bytes → text (Whisper, optional)
        GET  /api/health     model + readiness

The add-in reads the active range with Office.js, posts it here, renders the
returned preview, and on Apply writes the new grid back to the sheet.

Run:
    python server.py            # HTTPS on https://localhost:8765 (for Excel)
    python server.py --http     # plain HTTP (for quick browser preview)
"""
from __future__ import annotations

import io
import os
import sys
import json
import ssl
import socketserver
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import pandas as pd
from dotenv import load_dotenv

HERE = os.path.dirname(os.path.abspath(__file__))
WEB_DIR = os.path.join(HERE, "addin_web")
load_dotenv(os.path.join(HERE, ".env"))

sys.path.insert(0, HERE)
import engine  # noqa: E402
import llm      # noqa: E402

PORT = int(os.environ.get("SHEETMIND_PORT", "8765"))

_MIME = {
    ".html": "text/html; charset=utf-8", ".css": "text/css; charset=utf-8",
    ".js": "application/javascript; charset=utf-8", ".json": "application/json",
    ".png": "image/png", ".svg": "image/svg+xml", ".ico": "image/x-icon",
    ".woff2": "font/woff2", ".map": "application/json", ".xml": "application/xml",
}


# ── data <-> grid helpers ─────────────────────────────────────────────────
class MemoryStore:
    """Engine-compatible store backed by an in-memory DataFrame."""
    def __init__(self, df):
        self._df = df
        self.saved = 0

    @property
    def df(self):
        return self._df

    def _push_history(self):
        pass

    def save(self):
        self.saved += 1


def _df_from_grid(columns, rows):
    cols = [str(c).strip() for c in (columns or [])]
    df = pd.DataFrame(rows or [], columns=cols if cols else None)
    df.columns = [str(c).strip() for c in df.columns]
    for col in df.columns:
        conv = pd.to_numeric(df[col], errors="coerce")
        if conv.notna().sum() > 0 and conv.notna().mean() > 0.5:
            df[col] = conv
    return df


def _cell(v):
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    try:
        import numpy as np
        if isinstance(v, (np.integer,)):
            return int(v)
        if isinstance(v, (np.floating,)):
            f = float(v)
            return int(f) if f == int(f) else round(f, 6)
    except Exception:
        pass
    return v


def _grid_from_df(df):
    return {
        "columns": [str(c) for c in df.columns],
        "rows": [[_cell(v) for v in row] for row in df.itertuples(index=False, name=None)],
    }


def _step_json(s):
    return {
        "op": s.op, "kind": s.kind, "summary": s.summary,
        "affected": int(s.affected or 0),
        "value": _cell(s.value) if s.value is not None else None,
        "table": _grid_from_df(s.table) if s.table is not None else None,
        "error": s.error,
    }


def _diff(before: pd.DataFrame, after: pd.DataFrame, ops):
    """Lightweight per-cell diff for the preview table. Conservative: only marks
    cell-level changes when shapes line up and no reorder/clear happened."""
    out = {"row_status": [], "cell_changes": []}
    risky = any(o.get("op") in ("sort", "clear_rows") for o in ops)
    same_cols = list(before.columns) == list(after.columns)
    nb, na = len(before), len(after)
    for i in range(na):
        if risky or not same_cols:
            out["row_status"].append("normal")
            out["cell_changes"].append([])
        elif i >= nb:
            out["row_status"].append("added")
            out["cell_changes"].append(list(range(len(after.columns))))
        else:
            changed = [j for j in range(len(after.columns))
                       if str(before.iloc[i, j]) != str(after.iloc[i, j])]
            out["row_status"].append("changed" if changed else "normal")
            out["cell_changes"].append(changed)
    out["removed"] = max(0, nb - na)
    return out


# ── API handlers ──────────────────────────────────────────────────────────
def api_health(_):
    model = llm.model_chain()[0] if llm.is_configured() else None
    return {"ok": True, "configured": llm.is_configured(), "model": model}


def api_plan(body):
    cmd = (body.get("command") or "").strip()
    df = _df_from_grid(body.get("columns"), body.get("rows"))
    if not cmd:
        return {"ok": False, "error": "Empty command."}
    try:
        plan = engine.interpret(cmd, df)
    except Exception as e:  # noqa: BLE001
        return {"ok": False, "error": str(e)}

    if not plan.operations:
        return {"ok": True, "needs_confirm": False, "kind": "message",
                "summary": plan.summary or "I couldn't turn that into an action — try rephrasing.",
                "operations": [], "steps": []}

    pv = engine.dry_run(plan, df)
    if not pv.ok:
        return {"ok": False, "error": pv.error, "summary": plan.summary}

    res = {
        "ok": True,
        "needs_confirm": bool(plan.needs_confirm),
        "kind": "mutation" if plan.has_mutation else "query",
        "summary": plan.summary,
        "source": plan.source,
        "operations": plan.operations,
        "steps": [_step_json(s) for s in pv.steps],
    }
    if plan.has_mutation and pv.after is not None:
        res["preview"] = _grid_from_df(pv.after)
        res["diff"] = _diff(pv.before, pv.after, plan.operations)
    return res


def api_apply(body):
    ops = body.get("operations") or []
    df = _df_from_grid(body.get("columns"), body.get("rows"))
    plan = engine.Plan(operations=ops, summary=body.get("summary", ""))
    try:
        store = MemoryStore(df)
        done = engine.commit(plan, store)
    except Exception as e:  # noqa: BLE001
        return {"ok": False, "error": str(e)}
    if not done.ok:
        return {"ok": False, "error": done.error}
    return {"ok": True, "steps": [_step_json(s) for s in done.steps],
            "data": _grid_from_df(store.df)}


def api_load(raw: bytes):
    """Parse an uploaded .xlsx/.xls into a grid (browser preview / no-Excel use)."""
    try:
        xl = pd.ExcelFile(io.BytesIO(raw))
        sheet = xl.sheet_names[0]
        df = xl.parse(sheet)
        df.columns = [str(c).strip() for c in df.columns]
        for col in df.columns:
            conv = pd.to_numeric(df[col], errors="coerce")
            if conv.notna().sum() > 0 and conv.notna().mean() > 0.5:
                df[col] = conv
        g = _grid_from_df(df)
        g.update({"ok": True, "sheet": sheet, "sheets": xl.sheet_names})
        return g
    except Exception as e:  # noqa: BLE001
        return {"ok": False, "error": f"Could not read that file: {e}"}


def api_save(body):
    """Serialise a grid to .xlsx bytes for download. Returns (bytes, filename)."""
    df = _df_from_grid(body.get("columns"), body.get("rows"))
    name = os.path.basename(str(body.get("filename") or "SheetMind.xlsx")).strip()
    if not name.lower().endswith((".xlsx", ".xls")):
        name += ".xlsx"
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Sheet1")
    return buf.getvalue(), name


def api_transcribe(raw: bytes):
    try:
        from core import transcribe
        text = transcribe(raw)
        return {"ok": True, "text": text}
    except Exception as e:  # noqa: BLE001
        return {"ok": False, "error": str(e)}


# ── HTTP handler ──────────────────────────────────────────────────────────
class Handler(BaseHTTPRequestHandler):
    server_version = "SheetMind/2.0"

    def _send(self, code, payload=None, ctype="application/json", raw=None, extra=None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        body = raw if raw is not None else json.dumps(payload, default=str).encode("utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def _serve_static(self, path):
        if path in ("/", ""):
            path = "/taskpane.html"
        safe = os.path.normpath(path).lstrip("/")
        full = os.path.join(WEB_DIR, safe)
        if not full.startswith(WEB_DIR) or not os.path.isfile(full):
            self._send(404, {"error": "not found"})
            return
        ext = os.path.splitext(full)[1].lower()
        with open(full, "rb") as f:
            data = f.read()
        self._send(200, raw=data, ctype=_MIME.get(ext, "application/octet-stream"))

    def do_OPTIONS(self):
        self._send(204, raw=b"")

    def do_HEAD(self):
        self.do_GET()

    def do_GET(self):
        p = self.path.split("?", 1)[0]
        if p == "/api/health":
            self._send(200, api_health(None))
        elif p.startswith("/api/"):
            self._send(404, {"error": "unknown endpoint"})
        else:
            self._serve_static(p)

    def do_POST(self):
        p = self.path.split("?", 1)[0]
        length = int(self.headers.get("Content-Length", 0) or 0)
        body = self.rfile.read(length) if length else b""
        try:
            if p == "/api/transcribe":
                self._send(200, api_transcribe(body))
                return
            if p == "/api/load":
                self._send(200, api_load(body))
                return
            data = json.loads(body or b"{}")
            if p == "/api/plan":
                self._send(200, api_plan(data))
            elif p == "/api/apply":
                self._send(200, api_apply(data))
            elif p == "/api/save":
                blob, name = api_save(data)
                self._send(200, raw=blob,
                           ctype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           extra={"Content-Disposition": f'attachment; filename="{name}"'})
            else:
                self._send(404, {"error": "unknown endpoint"})
        except Exception as e:  # noqa: BLE001
            self._send(500, {"ok": False, "error": str(e)})

    def log_message(self, *args):  # quieter console
        return


# ── cert handling ─────────────────────────────────────────────────────────
def _find_certs():
    home = os.path.expanduser("~")
    # office-addin-dev-certs (trusted by the OS — best for Excel)
    oad = os.path.join(home, ".office-addin-dev-certs")
    crt, key = os.path.join(oad, "localhost.crt"), os.path.join(oad, "localhost.key")
    if os.path.isfile(crt) and os.path.isfile(key):
        return crt, key
    # our own self-signed (fallback — user must trust once)
    cdir = os.path.join(HERE, ".certs")
    crt, key = os.path.join(cdir, "localhost.crt"), os.path.join(cdir, "localhost.key")
    if os.path.isfile(crt) and os.path.isfile(key):
        return crt, key
    return _generate_self_signed(cdir)


def _generate_self_signed(cdir):
    os.makedirs(cdir, exist_ok=True)
    crt, key = os.path.join(cdir, "localhost.crt"), os.path.join(cdir, "localhost.key")
    try:
        import datetime
        from cryptography import x509
        from cryptography.x509.oid import NameOID
        from cryptography.hazmat.primitives import hashes, serialization
        from cryptography.hazmat.primitives.asymmetric import rsa
        k = rsa.generate_private_key(public_exponent=65537, key_size=2048)
        name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "localhost")])
        now = datetime.datetime.utcnow()
        cert = (x509.CertificateBuilder()
                .subject_name(name).issuer_name(name).public_key(k.public_key())
                .serial_number(x509.random_serial_number())
                .not_valid_before(now - datetime.timedelta(days=1))
                .not_valid_after(now + datetime.timedelta(days=825))
                .add_extension(x509.SubjectAlternativeName([x509.DNSName("localhost")]), False)
                .sign(k, hashes.SHA256()))
        with open(key, "wb") as f:
            f.write(k.private_bytes(serialization.Encoding.PEM,
                                    serialization.PrivateFormat.TraditionalOpenSSL,
                                    serialization.NoEncryption()))
        with open(crt, "wb") as f:
            f.write(cert.public_bytes(serialization.Encoding.PEM))
        print(f"  Generated self-signed cert at {cdir} (trust it once for Excel).")
        return crt, key
    except Exception as e:  # noqa: BLE001
        print(f"  Could not generate a cert ({e}). Run with --http for browser preview.")
        return None, None


def main():
    use_http = "--http" in sys.argv
    httpd = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    scheme = "http"
    if not use_http:
        crt, key = _find_certs()
        if crt and key:
            ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
            ctx.load_cert_chain(crt, key)
            httpd.socket = ctx.wrap_socket(httpd.socket, server_side=True)
            scheme = "https"
        else:
            print("  No certs — falling back to HTTP (Excel needs HTTPS).")

    model = llm.model_chain()[0] if llm.is_configured() else "no key (local parser only)"
    print("\n  🧠  SheetMind bridge")
    print(f"     {scheme}://localhost:{PORT}/taskpane.html")
    print(f"     model: {model}")
    print("     Ctrl+C to stop\n")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n  stopped.")


if __name__ == "__main__":
    main()
