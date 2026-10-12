"""
Flask Web Application for the IIT / NIT / IIIT JoSAA Cut-off Book Table Maker.

Flow
-----
  1. Login (josaa2027 / josaa@2027).
  2. Sidebar holds the Excel + Word-template paths (saved into config.json).
  3. "Scan" parses the workbook with JosaaExcelParser and lists every
     IIT / NIT / IIIT institute with its branch count and quotas.
  4. Clicking an institute renders its cut-off table instantly in the main
     pane (same data the book builder uses).
  5. "Build Full Book" runs book_updater.update_book() in a background
     thread, then word_com.fit_all_tables() + word_com.fix_index_pages()
     against MS Word.  Progress is polled from /api/build-status.
  6. The finished .docx is served from /api/download-book.
"""

import json
import hashlib
import io
import os
import re
import sys
import threading
import traceback
from functools import wraps

from flask import (Flask, render_template, request, jsonify, send_file,
                   session, redirect, url_for)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from parser import JosaaExcelParser
from book_updater import update_book
from institute_docx import build_institute_docx

app = Flask(__name__, static_folder="static", template_folder="templates")
app.secret_key = os.environ.get("SECRET_KEY", "josaa-book-maker-2027-secret")
app.config["TEMPLATES_AUTO_RELOAD"] = True   # always re-read templates/index.html

# --------------------------------------------------------------------------- #
#  Credentials
# --------------------------------------------------------------------------- #
LOGIN_USER = "josaa2027"
LOGIN_PASS = "josaa@2027"

# --------------------------------------------------------------------------- #
#  Paths  (all relative to this folder - works on Vercel / Linux / Windows)
# --------------------------------------------------------------------------- #
IS_VERCEL = bool(os.environ.get("VERCEL"))
LOCAL_BOOK = "ENGINNERING FINAL BOOK 2027 - UPDATED IIT NIT IIIT.docx"

# Excel + Word files ship in the bundled data/ folder.
DATA_DIR = os.path.join(BASE_DIR, "data")


def _first_existing(*paths):
    for p in paths:
        if os.path.exists(p):
            return p
    return paths[0]


DEFAULT_EXCEL = _first_existing(
    os.path.join(DATA_DIR, "iit,iiit,nit 2026.xlsx"),
    os.path.join(BASE_DIR, "iit,iiit,nit 2026.xlsx"))
DEFAULT_TEMPLATE = _first_existing(
    os.path.join(DATA_DIR, "ENGINNERING FINAL BOOK sachin sir.docx"),
    os.path.join(BASE_DIR, "ENGINNERING FINAL BOOK sachin sir.docx"))
BUNDLED_BOOK = _first_existing(
    os.path.join(DATA_DIR, LOCAL_BOOK),
    os.path.join(BASE_DIR, LOCAL_BOOK))
DEFAULT_OUTPUT_DIR = "/tmp" if (IS_VERCEL or os.name != "nt") else BASE_DIR

CONFIG_FILE = os.path.join(BASE_DIR, "config.json")


def _load_config():
    cfg = {"excel_path": DEFAULT_EXCEL,
           "template_path": DEFAULT_TEMPLATE,
           "output_dir": DEFAULT_OUTPUT_DIR}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg.update(json.load(f))
        except Exception:
            pass
    return cfg


def _save_config(cfg):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception:
        pass


def _clean(value):
    return str(value or "").strip().strip('"').strip("'")


# --------------------------------------------------------------------------- #
#  Auth decorator
# --------------------------------------------------------------------------- #
def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("logged_in"):
            if request.path.startswith("/api/"):
                return jsonify({"status": "auth", "message": "Login required"}), 401
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return wrapper


# --------------------------------------------------------------------------- #
#  Routes - pages
# --------------------------------------------------------------------------- #
@app.route("/login", methods=["GET", "POST"])
def login():
    error = ""
    if request.method == "POST":
        user = (request.form.get("username") or "").strip()
        pwd = request.form.get("password") or ""
        if user == LOGIN_USER and pwd == LOGIN_PASS:
            session["logged_in"] = True
            session["user"] = user
            nxt = request.args.get("next") or "/"
            if not str(nxt).startswith("/"):
                nxt = "/"
            return redirect(nxt)
        error = "Invalid username or password. Please try again."
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/whoami")
def whoami():
    return jsonify({"logged_in": bool(session.get("logged_in")),
                    "user": session.get("user", "")})


@app.route("/")
@login_required
def index():
    cfg = _load_config()
    return render_template("index.html",
                           excel_path=cfg["excel_path"],
                           template_path=cfg["template_path"],
                           output_dir=cfg["output_dir"],
                           logged_user=session.get("user", ""))


# --------------------------------------------------------------------------- #
#  Routes - config
# --------------------------------------------------------------------------- #
@app.route("/api/config", methods=["GET", "POST"])
@login_required
def api_config():
    cfg = _load_config()
    if request.method == "POST":
        data = request.json or {}
        for key in ("excel_path", "template_path", "output_dir"):
            if key in data:
                val = _clean(data.get(key))
                if val:
                    cfg[key] = val
        _save_config(cfg)
    return jsonify({"status": "success", **cfg})


# --------------------------------------------------------------------------- #
#  Routes - scan / institutes
# --------------------------------------------------------------------------- #
_PARSED = {}          # excel content-hash -> processed_data from parser.process_all()
_PARSED_LOCK = threading.Lock()
_PARSED_CACHE_NAME = "parsed_cache.json"


def _writable_cache_dir():
    return "/tmp" if IS_VERCEL else BASE_DIR


def _excel_key(excel_path):
    """Content hash of the workbook - cache identity that survives renames."""
    try:
        size = os.path.getsize(excel_path)
        h = hashlib.sha1()
        with open(excel_path, "rb") as f:
            for chunk in iter(lambda: f.read(1 << 20), b""):
                h.update(chunk)
        return "%d:%s" % (size, h.hexdigest())
    except OSError:
        return None


def _load_parsed_cache(key):
    seen = set()
    for d in (_writable_cache_dir(), DATA_DIR, BASE_DIR):
        p = os.path.join(d, _PARSED_CACHE_NAME)
        if p in seen:
            continue
        seen.add(p)
        try:
            with open(p, "r", encoding="utf-8") as f:
                blob = json.load(f)
            if blob.get("key") == key and isinstance(blob.get("parsed"), dict):
                return blob["parsed"]
        except Exception:
            continue
    return None


def _save_parsed_cache(key, parsed):
    try:
        d = _writable_cache_dir()
        os.makedirs(d, exist_ok=True)
        with open(os.path.join(d, _PARSED_CACHE_NAME), "w",
                  encoding="utf-8") as f:
            json.dump({"key": key, "parsed": parsed}, f,
                      ensure_ascii=False, separators=(",", ":"))
    except Exception:
        pass


def _get_parsed(excel_path, force=False):
    """Parse (and memoise) the workbook.

    Two cache layers make every reload instant:
      1. in-process memory keyed by the Excel content hash,
      2. a bundled / writable JSON file - survives cold starts, so the very
         first page load after a Vercel cold start is already fast.
    `force=True` (Scan button) re-reads the Excel and rewrites the cache.
    """
    key = _excel_key(excel_path)
    if key and not force:
        with _PARSED_LOCK:
            hit = _PARSED.get(key)
        if hit is not None:
            return hit
        cached = _load_parsed_cache(key)
        if cached is not None:
            with _PARSED_LOCK:
                _PARSED[key] = cached
            return cached
    parser = JosaaExcelParser(excel_path)
    data = parser.process_all()
    if key:
        with _PARSED_LOCK:
            _PARSED[key] = data
        _save_parsed_cache(key, data)
    return data


# --------------------------------------------------------------------------- #
#  "New in 2026" detection - compared against the 2025 branch list
# --------------------------------------------------------------------------- #
BRANCHES_2025_FILE = os.path.join(DATA_DIR, "branches_2025.xlsx")
_BRANCHES_2025 = None          # {norm_institute: {norm_program, ...}}


def _norm_text(s):
    """Loose text key so '&'-vs-'and', dots and spacing never fake a NEW branch."""
    s = str(s or "").replace("\xa0", " ").lower()
    s = s.replace("&", " and ")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _load_branches_2025():
    global _BRANCHES_2025
    if _BRANCHES_2025 is not None:
        return _BRANCHES_2025
    table = {}
    try:
        import pandas as pd
        if os.path.exists(BRANCHES_2025_FILE):
            xl = pd.ExcelFile(BRANCHES_2025_FILE)
            for sheet in xl.sheet_names:
                df = xl.parse(sheet)
                low = {str(c).lower(): c for c in df.columns}
                c_inst = next((c for k, c in low.items() if "inst" in k), None)
                c_prog = next((c for k, c in low.items()
                               if "program" in k or "branch" in k), None)
                if not c_inst or not c_prog:
                    continue
                for _, row in df.iterrows():
                    if pd.isna(row[c_inst]) or pd.isna(row[c_prog]):
                        continue
                    table.setdefault(_norm_text(row[c_inst]), set()).add(
                        _norm_text(row[c_prog]))
    except Exception:
        table = {}
    _BRANCHES_2025 = table
    return table


def _flag_new_branches(inst):
    """Set branch['is_new'] = True for branches absent from the 2025 list."""
    table = _load_branches_2025()
    if not table:
        return inst
    old = table.get(_norm_text(inst.get("institute_name", "")))
    for b in inst.get("branches", []):
        if old is None:                      # institute itself is new in 2026
            b["is_new"] = True
        else:
            b["is_new"] = _norm_text(b.get("program_name", "")) not in old
    return inst


@app.route("/api/scan", methods=["POST"])
@login_required
def api_scan():
    req = request.json or {}
    excel_path = _clean(req.get("excel_path")) or _load_config()["excel_path"]
    if not os.path.exists(excel_path):
        return jsonify({"status": "error",
                        "message": f"File not found: {excel_path}"}), 400
    try:
        data = _get_parsed(excel_path, force=bool(req.get("force")))
        institutes = []
        for kind in ("IIT", "NIT", "IIIT"):
            for inst in data.get(kind, {}).get("institutes", []):
                quotas = sorted({b["quota"] for b in inst["branches"]})
                institutes.append({
                    "kind": kind,
                    "sr_no": inst["sr_no"],
                    "name": inst["institute_name"],
                    "meta": inst.get("meta", {}),
                    "branch_count": len(inst["branches"]),
                    "quotas": quotas,
                })
        return jsonify({"status": "success",
                        "excel_path": excel_path,
                        "total": len(institutes),
                        "counts": {k: len(data.get(k, {}).get("institutes", []))
                                   for k in ("IIT", "NIT", "IIIT")},
                        "institutes": institutes})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e),
                        "trace": traceback.format_exc()}), 500


@app.route("/api/institute", methods=["POST"])
@login_required
def api_institute():
    req = request.json or {}
    excel_path = _clean(req.get("excel_path")) or _load_config()["excel_path"]
    name = str(req.get("name") or "").strip()
    kind = str(req.get("kind") or "").strip().upper()
    if not name or kind not in ("IIT", "NIT", "IIIT"):
        return jsonify({"status": "error", "message": "name + kind required"}), 400
    try:
        data = _get_parsed(excel_path)
        for inst in data.get(kind, {}).get("institutes", []):
            if inst["institute_name"] == name:
                return jsonify({"status": "success",
                                "institute": _flag_new_branches(inst)})
        return jsonify({"status": "error", "message": "Institute not found"}), 404
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


# --------------------------------------------------------------------------- #
#  Routes - single institute Word download (Excel ranks only)
# --------------------------------------------------------------------------- #
@app.route("/api/download-institute", methods=["POST"])
@login_required
def api_download_institute():
    req = request.json or {}
    excel_path = _clean(req.get("excel_path")) or _load_config()["excel_path"]
    name = str(req.get("name") or "").strip()
    kind_req = str(req.get("kind") or "").strip().upper()
    if not name:
        return jsonify({"status": "error", "message": "Institute name required"}), 400
    if not os.path.exists(excel_path):
        return jsonify({"status": "error",
                        "message": f"File not found: {excel_path}"}), 400
    try:
        data = _get_parsed(excel_path)
        kinds = ([kind_req] if kind_req in ("IIT", "NIT", "IIIT")
                 else ["IIT", "NIT", "IIIT"])
        inst = None
        for kind in kinds:
            for it in data.get(kind, {}).get("institutes", []):
                if it["institute_name"] == name:
                    inst = it
                    break
            if inst:
                break
        if inst is None:
            return jsonify({"status": "error",
                            "message": "Institute not found in Excel"}), 404
        blob, filename = build_institute_docx(inst)
        return send_file(io.BytesIO(blob), as_attachment=True,
                         download_name=filename,
                         mimetype="application/vnd.openxmlformats-officedocument"
                                 ".wordprocessingml.document")
    except Exception as e:
        return jsonify({"status": "error", "message": str(e),
                        "trace": traceback.format_exc()}), 500


# --------------------------------------------------------------------------- #
#  Routes - build the full book (background thread + progress polling)
# --------------------------------------------------------------------------- #
BUILD_STATE = {"pct": 0, "msg": "", "done": False, "error": "",
               "report": None, "output": None, "running": False}
_BUILD_LOCK = threading.Lock()


def _set_state(**kw):
    with _BUILD_LOCK:
        BUILD_STATE.update(kw)


def _get_state():
    with _BUILD_LOCK:
        return dict(BUILD_STATE)


def _build_worker(excel_path, template_path, output_path):
    try:
        _set_state(pct=0, msg="Starting book rebuild...", done=False,
                   error="", report=None, output=None)

        def cb(pct, msg):
            _set_state(pct=int(pct), msg=str(msg))

        report = update_book(excel_path, template_path, output_path,
                             progress_cb=cb)
        report.update(_word_steps(output_path, cb))
        _set_state(pct=100, msg="Book built successfully.",
                   done=True, report=report, output=output_path, running=False)
    except Exception as e:
        _set_state(pct=100, msg="ERROR", done=True, running=False,
                   error=f"{e}\n{traceback.format_exc()}")


def _word_steps(output_path, cb):
    """MS Word page-fit + index fix.  Skipped on servers (no MS Word)."""
    _set_state(pct=70, msg="Running Word page-fit + index fix...")
    if os.name != "nt" or IS_VERCEL:
        return {"word_error": "Skipped: MS Word is not available on this server."}
    try:
        import word_com
        fit = word_com.fit_all_tables(output_path,
                                      progress_cb=lambda p, m: cb(70 + p * 0.2, m))
        idx = word_com.fix_index_pages(output_path,
                                       progress_cb=lambda p, m: cb(90 + p * 0.1, m))
        return {"fit": fit,
                "index": {"updated": len(idx.get("updated", [])),
                          "missing": idx.get("missing", [])}}
    except Exception as e:
        return {"word_error": str(e)}


@app.route("/api/build-book", methods=["POST"])
@login_required
def api_build_book():
    st = _get_state()
    if st.get("running"):
        return jsonify({"status": "error",
                        "message": "A build is already running."}), 409

    req = request.json or {}
    cfg = _load_config()
    excel_path = _clean(req.get("excel_path")) or cfg["excel_path"]
    template_path = _clean(req.get("template_path")) or cfg["template_path"]
    output_dir = _clean(req.get("output_dir")) or cfg["output_dir"]

    if not os.path.exists(excel_path):
        return jsonify({"status": "error",
                        "message": f"Excel not found: {excel_path}"}), 400
    if not os.path.exists(template_path):
        return jsonify({"status": "error",
                        "message": f"Template not found: {template_path}"}), 400
    os.makedirs(output_dir, exist_ok=True)

    output_path = os.path.join(
        output_dir,
        "ENGINNERING FINAL BOOK 2027 - UPDATED IIT NIT IIIT.docx")

    _set_state(running=True, done=False, pct=0, msg="Queued...", error="",
               report=None, output=None)

    if IS_VERCEL:
        # Serverless: no shared background state - build synchronously and
        # return everything the UI needs in this single response.
        def cb(pct, msg):
            _set_state(pct=int(pct), msg=str(msg))
        try:
            _set_state(msg="Starting book rebuild...")
            report = update_book(excel_path, template_path, output_path,
                                 progress_cb=cb)
            report.update(_word_steps(output_path, cb))
            _set_state(pct=100, msg="Book built successfully.", done=True,
                       report=report, output=output_path, running=False)
            return jsonify({"status": "success", "done": True,
                            "report": report, "output": output_path})
        except Exception as e:
            _set_state(pct=100, msg="ERROR", done=True, running=False,
                       error=f"{e}\n{traceback.format_exc()}")
            return jsonify({"status": "error", "message": str(e),
                            "trace": traceback.format_exc()}), 500

    t = threading.Thread(target=_build_worker,
                         args=(excel_path, template_path, output_path),
                         daemon=True)
    t.start()
    return jsonify({"status": "success", "message": "Build started.",
                    "output": output_path})


@app.route("/api/build-status")
@login_required
def api_build_status():
    return jsonify({"status": "success", **_get_state()})


@app.route("/api/download-book")
@login_required
def api_download_book():
    st = _get_state()
    path = st.get("output")
    if path and os.path.exists(path):
        return send_file(path, as_attachment=True,
                         download_name=os.path.basename(path))
    cfg = _load_config()
    fallback = os.path.join(cfg["output_dir"],
                            "ENGINNERING FINAL BOOK 2027 - UPDATED IIT NIT IIIT.docx")
    if os.path.exists(fallback):
        return send_file(fallback, as_attachment=True,
                         download_name=os.path.basename(fallback))
    if os.path.exists(BUNDLED_BOOK):
        # Ready-made book shipped with the deployment.
        return send_file(BUNDLED_BOOK, as_attachment=True,
                         download_name=LOCAL_BOOK)
    return jsonify({"status": "error", "message": "No book built yet."}), 404


# --------------------------------------------------------------------------- #
if __name__ == "__main__":
    print("=" * 60)
    print("  JoSAA IIT / NIT / IIIT Cut-off Book Table Maker")
    print("  http://localhost:5000")
    print("=" * 60)
    app.run(host="0.0.0.0", port=5000, debug=False)
