"""Local dashboard: http://127.0.0.1:8765 — upload Excel, Check now, settings, optional schedule."""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import urllib.parse
import webbrowser
from datetime import datetime, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import monitor

ROOT = monitor.ROOT
CONFIG = ROOT / "config.json"
DATA = ROOT / "data"
UPLOADS = DATA / "uploads"
UPLOAD_INDEX = UPLOADS / "index.json"
TEMPLATE = ROOT / "templates" / "land_records_template.xlsx"
MAX_UPLOAD = 10 * 1024 * 1024
SECRET_KEYS = {("email", "password"), ("google_sheet", "token"), (None, "twocaptcha_api_key")}
lock = threading.Lock()
stop_event = threading.Event()
PAGE = (ROOT / "ui.html").read_text(encoding="utf-8")  # read once: page and backend always come from the same version
state = {"running": False, "stopping": False, "progress": None, "run_started": None, "run_keys": None, "last_run": None,
         "last_result": None, "ran_slots": set()}
_rows_cache = {}


# ---------- config ----------
def load_cfg():
    if not CONFIG.exists():
        shutil.copy(ROOT / "config.example.json", CONFIG)
    return json.loads(CONFIG.read_text(encoding="utf-8"))


def save_cfg(cfg):
    CONFIG.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")


def excel_path():
    p = load_cfg().get("excel_path")
    return ROOT / p if p and (ROOT / p).exists() else None


def rows_for(path: Path):
    """Parsed sheet, cached until the file changes (re-reading 1000s of rows every poll is slow)."""
    mt = path.stat().st_mtime_ns
    hit = _rows_cache.get(path)
    if not hit or hit[0] != mt:
        _rows_cache[path] = hit = (mt, list({r["key"]: r for r in monitor.read_rows(path)}.values()))
    return hit[1]


def public_cfg():
    """Config for the settings form; secrets are never sent back, only whether they are set."""
    cfg = load_cfg()
    for section, key in SECRET_KEYS:
        holder = cfg.setdefault(section, {}) if section else cfg
        holder[f"{key}_set"] = bool(holder.get(key))
        holder[key] = ""
    return cfg


def update_cfg(body):
    old, cfg = load_cfg(), load_cfg()
    for section in ("email", "schedule", "google_sheet"):
        if isinstance(body.get(section), dict):
            cfg.setdefault(section, {}).update(body[section])
    for key in ("twocaptcha_api_key", "delay_between_checks_sec", "show_browser", "parallel_browsers"):
        if key in body:
            cfg[key] = body[key]
    for section, key in SECRET_KEYS:  # blank secret field = keep the saved one
        holder = cfg.get(section, {}) if section else cfg
        holder.pop(f"{key}_set", None)
        if not holder.get(key):
            holder[key] = ((old.get(section) or {}) if section else old).get(key) or ""
    for section, key in (("email", "to"), ("schedule", "times")):
        if isinstance(cfg.get(section, {}).get(key), str):
            cfg[section][key] = [t.strip() for t in cfg[section][key].split(",") if t.strip()]
    save_cfg(cfg)


# ---------- runs ----------
def start_run(keys=None):
    """Check every parcel, or only `keys` (a re-check of failed ones, or a single parcel)."""
    if not excel_path() or not lock.acquire(blocking=False):
        return False
    stop_event.clear()
    keys = set(keys) if keys else None

    def job():
        total = len(keys) if keys else len(rows_for(excel_path()))
        state.update(running=True, progress={"done": 0, "total": total, "key": ""}, run_keys=sorted(keys) if keys else None,
                     run_started=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        try:
            state["last_result"] = monitor.run(
                load_cfg(), lambda d, n, k: state.update(progress={"done": d, "total": n, "key": k}), stop_event, keys)
        except Exception as ex:
            traceback.print_exc()
            state["last_result"] = {"errors": [f"{type(ex).__name__}: {ex}"]}
        finally:
            state.update(running=False, progress=None, stopping=False,
                         last_run=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            lock.release()

    threading.Thread(target=job, daemon=True).start()
    return True


def scheduler():
    # ponytail: only fires while this app is running; use launchd/Task Scheduler if it must run unattended.
    while True:
        cfg = load_cfg().get("schedule", {})
        slot = datetime.now().strftime("%Y-%m-%d %H:%M")
        if cfg.get("enabled") and slot[-5:] in cfg.get("times", []) and slot not in state["ran_slots"]:
            state["ran_slots"].add(slot)
            start_run()
        time.sleep(20)


# ---------- state for the dashboard ----------
def version():
    """Cheap fingerprint; the page skips re-rendering when nothing changed."""
    mt = lambda p: p.stat().st_mtime_ns if p and p.exists() else 0
    xl = excel_path()
    return (f"{mt(monitor.DB)}|{xl}|{mt(xl)}|{mt(CONFIG)}|{state['running']}|{state['stopping']}|"
            f"{state['progress']}|{state['last_run']}")


def parcel_state(snap, st):
    if snap and (snap.get("_Flags", "None") != "None" or any(
            str(v).startswith("MISMATCH") for k, v in snap.items() if k.endswith("Check"))):
        return "attention"
    if st and st[2]:
        return "error"
    return "ok" if snap else "pending"


WARNINGS = [("Pending transaction", "Pending transaction"), ("Prohibited", "Prohibited"),
            ("Not digitally signed", "Not digitally signed"), ("NALA", "NALA / non-agricultural")]


def snapshot():
    cfg = load_cfg()
    xl = excel_path()
    em = cfg.get("email") or {}
    out = {k: v for k, v in state.items() if k != "ran_slots"}
    out.update(version=version(), schedule=cfg.get("schedule", {}), excel={"name": xl.name.split("__", 1)[-1]} if xl else None,
               email_ready=bool(em.get("smtp_host") and em.get("password") and em.get("to")),
               sheet_ready=bool((cfg.get("google_sheet") or {}).get("web_app_url")),
               parcels=[], recent=[], unseen=0, daily={}, runs=[], warnings=[])
    if not xl:
        return out
    try:
        rows = rows_for(xl)
    except Exception as ex:
        out["load_error"] = str(ex)
        return out
    con = monitor.db()
    snaps = monitor.latest_all(con)
    status = {r[0]: r for r in con.execute("SELECT key,last_checked,fails,last_error,last_ok FROM status")}
    keys = {r["key"] for r in rows}
    warn_counts = {label: 0 for _, label in WARNINGS} | {"Owner differs": 0, "Extent differs": 0}
    for r in rows:
        snap, st = snaps.get(r["key"]), status.get(r["key"])
        s = snap or {}
        flags = s.get("_Flags", "")
        for needle, label in WARNINGS:
            warn_counts[label] += needle in flags
        warn_counts["Owner differs"] += s.get("_Owner Check", "").startswith("MISMATCH")
        warn_counts["Extent differs"] += s.get("_Extent Check", "").startswith("MISMATCH")
        out["parcels"].append({
            **{k: r.get(k, "") for k in ("key", "district", "mandal", "village", "survey", "owner", "acres", "guntas", "alert_on")},
            "site_owner": s.get(monitor.NAME_LABEL, "").split(" / ")[0], "site_extent": s.get(monitor.EXTENT_LABEL, ""),
            "extent_guntas": monitor.site_extent_guntas(s.get(monitor.EXTENT_LABEL, "")) or 0,
            "land_status": s.get("Land Status", ""), "flags": flags,
            "owner_check": s.get("_Owner Check", ""), "extent_check": s.get("_Extent Check", ""),
            "last_checked": st[1] if st else None, "fails": st[2] if st else 0,
            "last_error": st[3] if st else None, "last_ok": st[4] if st else None,
            "state": parcel_state(snap, st)})
    out["warnings"] = [[k, v] for k, v in warn_counts.items()]
    since = (datetime.now() - timedelta(days=29)).strftime("%Y-%m-%d")
    for d, kind, key in con.execute("SELECT substr(ts,1,10), kind, key FROM events WHERE kind!='BASELINE' AND ts>=?", (since,)):
        if key in keys:
            out["daily"].setdefault(d, {"CHANGE": 0, "ISSUE": 0, "ERROR": 0})[kind] += 1
    out["unseen"] = sum(1 for (k,) in con.execute("SELECT key FROM events WHERE seen=0 AND kind!='BASELINE'") if k in keys)
    out["recent"] = [dict(zip(("id", "ts", "key", "kind", "message", "seen"), e)) for e in con.execute(
        "SELECT rowid,ts,key,kind,message,seen FROM events WHERE kind!='BASELINE' ORDER BY rowid DESC LIMIT 200") if e[2] in keys][:6]
    out["runs"] = [dict(zip(("started", "finished", "total", "ok", "failed", "alerts", "seconds"), r))
                   for r in con.execute("SELECT * FROM runs ORDER BY rowid DESC LIMIT 10")]
    con.close()
    return out


def parcel_detail(key):
    con = monitor.db()
    snap = monitor.latest(con, key) or {}
    events = [dict(zip(("id", "ts", "kind", "message", "seen"), e)) for e in con.execute(
        "SELECT rowid,ts,kind,message,seen FROM events WHERE key=? ORDER BY rowid DESC LIMIT 200", (key,))]
    con.close()
    return {"data": {k: v for k, v in snap.items() if not k.startswith("_")}, "events": events}


def events_page(q):
    """Alerts page: filtered + paged on the server so thousands of alerts stay fast."""
    xl = excel_path()
    keys = {r["key"] for r in rows_for(xl)} if xl else set()
    kind, text, only_new = q.get("kind", [""])[0], q.get("q", [""])[0].lower(), q.get("new", [""])[0] == "1"
    offset, limit = int(q.get("offset", ["0"])[0]), 50
    sql, args = "SELECT rowid,ts,key,kind,message,seen FROM events WHERE kind!='BASELINE'", []
    if kind:
        sql, args = sql + " AND kind=?", args + [kind]
    if only_new:
        sql += " AND seen=0"
    con = monitor.db()
    hits = [e for e in con.execute(sql + " ORDER BY rowid DESC", args)
            if e[2] in keys and (not text or text in e[2].lower() or text in e[4].lower())]
    con.close()
    return {"total": len(hits), "items": [dict(zip(("id", "ts", "key", "kind", "message", "seen"), e))
                                          for e in hits[offset:offset + limit]]}


def mark_seen(body):
    con = monitor.db()
    if body.get("all"):
        con.execute("UPDATE events SET seen=1")
    else:
        con.executemany("UPDATE events SET seen=1 WHERE rowid=?", [(int(i),) for i in body.get("ids", [])])
    con.commit()
    con.close()


# ---------- sheets & storage ----------
def sheets():
    index = json.loads(UPLOAD_INDEX.read_text(encoding="utf-8")) if UPLOAD_INDEX.exists() else []
    active = load_cfg().get("excel_path", "")
    return [{**e, "active": active == f"data/uploads/{e['file']}",
             "size": (UPLOADS / e["file"]).stat().st_size} for e in reversed(index) if (UPLOADS / e["file"]).exists()]


def accept_excel(data: bytes, filename: str):
    if not filename.lower().endswith(".xlsx"):
        return {"ok": False, "error": "Upload an .xlsx file (Excel 2007 or newer). Save .xls or .csv as .xlsx first."}
    if len(data) > MAX_UPLOAD:
        return {"ok": False, "error": "File is larger than 10 MB."}
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        f.write(data)
    tmp = Path(f.name)
    try:
        rows = monitor.read_rows(tmp)
    except Exception as ex:
        tmp.unlink()
        return {"ok": False, "error": f"Could not read this file. Make sure it is a real .xlsx Excel file ({ex})."}
    if not rows:
        tmp.unlink()
        return {"ok": False, "error": "No rows with a Survey No were found in the first sheet."}
    UPLOADS.mkdir(parents=True, exist_ok=True)
    safe = "".join(c for c in Path(filename).name if c not in '\\/:*?"<>|') or "land_records.xlsx"
    dest = UPLOADS / f"{datetime.now().strftime('%Y-%m-%d_%H%M%S')}__{safe}"
    shutil.move(tmp, dest)
    index = json.loads(UPLOAD_INDEX.read_text(encoding="utf-8")) if UPLOAD_INDEX.exists() else []
    index.append({"file": dest.name, "name": safe, "uploaded": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                  "rows": len(rows), "villages": len({(r["mandal"], r["village"]) for r in rows if r["village"]}),
                  "mandals": len({(r["district"], r["mandal"]) for r in rows if r["mandal"]})})
    UPLOAD_INDEX.write_text(json.dumps(index, indent=2, ensure_ascii=False), encoding="utf-8")
    cfg = load_cfg()
    cfg["excel_path"] = f"data/uploads/{dest.name}"
    save_cfg(cfg)
    return {"ok": True, "rows": len(rows), "name": safe, "villages": index[-1]["villages"], "started": start_run()}


def activate_sheet(file):
    if "/" in file or not (UPLOADS / file).exists():
        return {"ok": False, "error": "That file no longer exists."}
    cfg = load_cfg()
    cfg["excel_path"] = f"data/uploads/{file}"
    save_cfg(cfg)
    return {"ok": True}


def delete_sheet(file):
    if "/" in file or load_cfg().get("excel_path") == f"data/uploads/{file}":
        return {"ok": False, "error": "This is the sheet in use. Switch to another sheet first."}
    (UPLOADS / file).unlink(missing_ok=True)
    index = [e for e in json.loads(UPLOAD_INDEX.read_text(encoding="utf-8")) if e["file"] != file]
    UPLOAD_INDEX.write_text(json.dumps(index, indent=2, ensure_ascii=False), encoding="utf-8")
    return {"ok": True}


def storage():
    def size(p):
        if not p.exists():
            return 0
        return sum(f.stat().st_size for f in p.rglob("*") if f.is_file()) if p.is_dir() else p.stat().st_size
    con = monitor.db()
    counts = {t: con.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0] for t in ("snapshots", "events", "runs")}
    con.close()
    return {"root": str(ROOT), "items": [
        {"id": "uploads", "label": "Uploaded sheets", "path": str(UPLOADS), "size": size(UPLOADS),
         "note": f"{len(sheets())} file(s). Every upload is kept with the date and time in its name."},
        {"id": "db", "label": "Check history", "path": str(monitor.DB), "size": size(monitor.DB),
         "note": f"{counts['snapshots']} saved records, {counts['events']} alerts, {counts['runs']} checks."},
        {"id": "report", "label": "Excel report", "path": str(monitor.REPORT), "size": size(monitor.REPORT),
         "note": "Rewritten after every check. Same file as the Report button."},
        {"id": "config", "label": "Settings file", "path": str(CONFIG), "size": size(CONFIG),
         "note": "Email, schedule and other settings, including your email app password."}]}


def open_in_finder(which):
    target = {"uploads": UPLOADS, "db": DATA, "report": monitor.REPORT.parent, "config": ROOT}.get(which)
    if not target:
        return {"ok": False}
    target.mkdir(parents=True, exist_ok=True)
    if sys.platform == "darwin":
        subprocess.Popen(["open", str(target)])
    elif sys.platform == "win32":
        os.startfile(str(target))
    else:
        subprocess.Popen(["xdg-open", str(target)])
    return {"ok": True}


def run_request(body):
    if not excel_path():
        return {"started": False, "error": "Add a land list first."}
    keys = body.get("keys")
    if body.get("failed"):
        current = {r["key"] for r in rows_for(excel_path())}
        con = monitor.db()
        keys = [k for (k,) in con.execute("SELECT key FROM status WHERE fails>0") if k in current]
        con.close()
        if not keys:
            return {"started": False, "error": "No failed parcels to re-check."}
    started = start_run(keys)
    return {"started": started, "count": len(keys) if keys else None,
            **({} if started else {"error": "A check is already running."})}


def sheet_detail(file):
    """Rows of one uploaded sheet, problems found in it, and how it differs from the list in use."""
    path = UPLOADS / file
    if "/" in file or not path.exists():
        return {"error": "That sheet no longer exists."}
    meta = next((e for e in sheets() if e["file"] == file), {"file": file, "name": file})
    rows = monitor.read_rows(path)
    seen, issues = {}, []
    for r in rows:
        if r["key"] in seen:
            issues.append(f"Row {r['row']}: survey {r['survey']} in {r['village']} repeats row {seen[r['key']]}; only the first is checked.")
        seen.setdefault(r["key"], r["row"])
        missing = [k for k in ("district", "mandal", "village") if not r[k]]
        if missing:
            issues.append(f"Row {r['row']}: {', '.join(missing)} is blank, so survey {r['survey']} cannot be found.")
        for k in ("acres", "guntas"):
            v = r.get(k, "")
            try:
                if v != "" and (float(v) < 0 or (k == "guntas" and float(v) >= 40)):
                    issues.append(f"Row {r['row']}: {k} is {v}" + (", but there are only 40 guntas in an acre." if k == "guntas" else "."))
            except ValueError:
                issues.append(f"Row {r['row']}: {k} '{v}' is not a number.")
    diff = None
    active = excel_path()
    if active and active != path:
        now_keys, these = {r["key"] for r in rows_for(active)}, set(seen)
        diff = {"added": sorted(these - now_keys)[:200], "removed": sorted(now_keys - these)[:200],
                "added_count": len(these - now_keys), "removed_count": len(now_keys - these)}
    return {"meta": meta, "rows": rows, "issues": issues[:200], "issue_count": len(issues), "diff": diff}


def stop_run():
    if not state["running"]:
        return {"ok": False, "error": "No check is running."}
    stop_event.set()
    state["stopping"] = True
    return {"ok": True}


def migrate_legacy():
    """Older versions kept the sheet in data/ and the report in output/; move both into the current layout."""
    p = load_cfg().get("excel_path") or ""
    src = ROOT / p
    if p and src.exists() and not p.startswith("data/uploads/"):
        UPLOADS.mkdir(parents=True, exist_ok=True)
        dest = UPLOADS / f"{datetime.fromtimestamp(src.stat().st_mtime).strftime('%Y-%m-%d_%H%M%S')}__{src.name}"
        rows = monitor.read_rows(src)
        shutil.move(src, dest)
        index = json.loads(UPLOAD_INDEX.read_text(encoding="utf-8")) if UPLOAD_INDEX.exists() else []
        index.append({"file": dest.name, "name": src.name, "uploaded": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                      "rows": len(rows), "villages": len({(r["mandal"], r["village"]) for r in rows if r["village"]}),
                      "mandals": len({(r["district"], r["mandal"]) for r in rows if r["mandal"]})})
        UPLOAD_INDEX.write_text(json.dumps(index, indent=2, ensure_ascii=False), encoding="utf-8")
        cfg = load_cfg()
        cfg["excel_path"] = f"data/uploads/{dest.name}"
        save_cfg(cfg)
    old_report = ROOT / "output" / "land_status_report.xlsx"
    if old_report.exists():
        monitor.REPORT.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(old_report, monitor.REPORT)
    if (ROOT / "output").exists() and not any((ROOT / "output").iterdir()):
        (ROOT / "output").rmdir()


def clear_history():
    if state["running"]:
        return {"ok": False, "error": "Wait for the running check to finish."}
    con = monitor.db()
    con.executescript("DELETE FROM snapshots; DELETE FROM events; DELETE FROM status; DELETE FROM runs; VACUUM;")
    con.close()
    return {"ok": True}


# ---------- http ----------
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class Handler(BaseHTTPRequestHandler):
    def _send(self, body, ctype="application/json", code=200, filename=None):
        data = body if isinstance(body, bytes) else (
            body.encode() if isinstance(body, str) else json.dumps(body, ensure_ascii=False).encode())
        self.send_response(code)
        self.send_header("Content-Type", ctype if filename else f"{ctype}; charset=utf-8")
        if filename:
            self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{urllib.parse.quote(filename)}")
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def _file(self, path, name):
        if not path or not path.exists():
            return self._send({"error": "File not found."}, code=404)
        self._send(path.read_bytes(), XLSX, filename=name)

    def do_GET(self):
        url = urllib.parse.urlparse(self.path)
        q = urllib.parse.parse_qs(url.query)
        if url.path == "/":
            return self._send(PAGE, "text/html")
        if url.path == "/api/state":
            return self._send({"same": True} if q.get("v", [""])[0] == version() else snapshot())
        if url.path == "/api/parcel":
            return self._send(parcel_detail(q.get("key", [""])[0]))
        if url.path == "/api/events":
            return self._send(events_page(q))
        if url.path == "/api/config":
            return self._send(public_cfg())
        if url.path == "/api/sheet":
            return self._send(sheet_detail(q.get("file", [""])[0]))
        if url.path == "/api/sheets":
            return self._send(sheets())
        if url.path == "/api/storage":
            return self._send(storage())
        if url.path == "/download/template":
            return self._file(TEMPLATE, "land_records_template.xlsx")
        if url.path == "/download/report":
            return self._file(monitor.REPORT, "land_status_report.xlsx")
        if url.path == "/download/sheet":
            f = q.get("file", [""])[0]
            return self._file(UPLOADS / f if f and "/" not in f else None, f.split("__", 1)[-1])
        self._send({"error": "not found"}, code=404)

    def do_POST(self):
        length = int(self.headers.get("Content-Length") or 0)
        if self.path == "/api/upload":
            if length > MAX_UPLOAD:
                return self._send({"ok": False, "error": "File is larger than 10 MB."})
            name = urllib.parse.unquote(self.headers.get("X-Filename", "upload.xlsx"))
            return self._send(accept_excel(self.rfile.read(length), name))
        body = json.loads(self.rfile.read(length) or b"{}")
        routes = {
            "/api/stop": stop_run,
            "/api/run": lambda: run_request(body),
            "/api/config": lambda: (update_cfg(body), {"ok": True})[1],
            "/api/seen": lambda: (mark_seen(body), {"ok": True})[1],
            "/api/sheets/activate": lambda: activate_sheet(body.get("file", "")),
            "/api/sheets/delete": lambda: delete_sheet(body.get("file", "")),
            "/api/open": lambda: open_in_finder(body.get("which")),
            "/api/clear-history": lambda: clear_history() if body.get("confirm") == "yes" else {"ok": False},
            "/api/test-email": test_email,
        }
        if self.path in routes:
            return self._send(routes[self.path]())
        self._send({"error": "not found"}, code=404)

    def log_message(self, *a):
        pass


def test_email():
    try:
        monitor.send_email(load_cfg(), "Bhu Bharati monitor: test email", "Email alerts are working.")
        return {"ok": True}
    except Exception as ex:
        return {"ok": False, "error": str(ex)}


if __name__ == "__main__":
    migrate_legacy()
    port = load_cfg().get("port", 8765)
    threading.Thread(target=scheduler, daemon=True).start()
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    print(f"Bhu Bharati monitor running at http://127.0.0.1:{port}  (Ctrl+C to stop)")
    webbrowser.open(f"http://127.0.0.1:{port}")
    server.serve_forever()
