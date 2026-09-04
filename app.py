import os
import io
import sys
import csv
import json
import time
import threading
import subprocess

from flask import Flask, request, Response, render_template, jsonify, send_file

'''
**************************************************************************************************************************************************************************************
Control panel for the monthly lab pipeline. Serves an HTML/JS frontend and
runs each script (index.py / clsSplit.py / clsSeguiment.py) as a subprocess,
streaming its stdout to the browser live via Server-Sent Events (SSE).

Deployment: ONE Windows machine (the one with Excel + the Z: drive) runs this
server; everybody else just opens it in a browser and installs nothing. The
Excel COM automation of Pas 2 needs a real, interactive Excel session, so that
machine must stay logged on -- see docs/INSTALLACIO_SERVIDOR.md.

    Panell.bat            -- start + open the browser (day-to-day use)
    Panell.bat servei     -- start without opening a browser (autostart)
    Panell.bat setup      -- only prepare the venv

Author: Marc Eslava
**************************************************************************************************************************************************************************************
'''

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable  # same interpreter that launched the server

# Listen on every interface so the rest of the team can reach the panel over the
# LAN. Set PANELL_HOST=127.0.0.1 to go back to a local-only panel.
HOST = os.environ.get("PANELL_HOST", "0.0.0.0")
PORT = int(os.environ.get("PANELL_PORT", "5099"))

app = Flask(__name__)
# Local dev tool: never let the browser cache static files, and re-read the HTML
# template on every request, so edits to app.js/style.css/index.html take effect
# on a normal reload (a server restart is only needed for app.py itself).
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 0
app.config["TEMPLATES_AUTO_RELOAD"] = True
# Sostre de la pujada del detector: un BIExportGrid gran no passa de pocs MB.
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024


@app.after_request
def _no_cache(resp):
    resp.headers["Cache-Control"] = "no-store"
    return resp

# Only one heavy job at a time: Excel COM automation and the network Excel files
# don't tolerate concurrent runs. A second request while busy is rejected.
# With several people sharing the panel this happens for real, so the running
# job is tracked to tell the second user WHAT is running and since when,
# instead of a bare "busy".
_run_lock = threading.Lock()
_current_job = {}  # {"script", "started", "by"} while a run is in flight

# Whitelist of runnable scripts (never build a path from user input).
SCRIPTS = {
    "index":     "index.py",
    "split":     "clsSplit.py",
    "seguiment": "clsSeguiment.py",
}

# Output CSVs kept up to date by index.py.
CONDITIONS_CSV = os.path.join(BASE_DIR, "output", "product_conditions.csv")  # upsert per Cod Unif
LOST_MONEY_CSV = os.path.join(BASE_DIR, "output", "lost_money.csv")          # per-lab lost-money report
DETECTOR_CSV = os.path.join(BASE_DIR, "output", "detector_refs.csv")         # refs del grid que la BD no coneix

# Network folder that must be reachable on the Z: drive for the pipeline to work.
Z_CHECK_DIR = r"Z:\Compres\INDÚSTRIA FARMACÈUTICA\01. CARPETES LABORATORIS\102. Informes i Rappel"


# Formats one Server-Sent Event data line (UTF-8, accents preserved).
def _sse(obj):
    return "data: " + json.dumps(obj, ensure_ascii=False) + "\n\n"


# Builds the subprocess command from the request params. Returns None if the
# script is not whitelisted. Only known flags are forwarded (no shell, list argv).
def _build_command(p):
    script = p.get("script")
    if script not in SCRIPTS:
        return None
    cmd = [PY, "-u", os.path.join(BASE_DIR, SCRIPTS[script])]
    year, month = p.get("year"), p.get("month")
    period = p.get("period", "")

    if script == "index":
        cmd += ["--rappel", p.get("rappel", "BIFARMA"), "--period", period]
        if year:
            cmd += ["--year", str(year)]
        if month:
            cmd += ["--month", str(month)]
        if p.get("only_para", "") in ("1", "true", "on"):
            cmd += ["--only-para"]
        if p.get("laboratori", "").strip():
            cmd += ["--laboratori", p["laboratori"].strip()]
    elif script == "split":
        # No default for --steps: a request that forwards no step must do
        # NOTHING (clsSplit.py exits with a message), never silently split.
        cmd += ["--period", period, "--steps", p.get("steps", "")]
        if year:
            cmd += ["--year", str(year)]
        if month:
            cmd += ["--month", str(month)]
        # Single-lab split: reads the one-lab master written by Pas 1.
        if p.get("laboratori", "").strip():
            cmd += ["--laboratori", p["laboratori"].strip()]
    elif script == "seguiment":
        cmd += ["--min-year", str(p.get("min_year", "2024"))]
        if year:
            cmd += ["--year", str(year)]
    return cmd


# Serves the single-page control panel.
@app.route("/")
def index():
    return render_template("index.html")


# Quick health check: is the SQL server reachable and is the Z: drive accessible?
# Used by the "Comprovar servidor i Z:" button. Short login_timeout so a dead
# server answers fast instead of hanging the request.
@app.route("/check")
def check():
    out = {}
    t = time.time()
    try:
        import db_config as cfg
        from SQLConnection import SQLConnection
        conn = SQLConnection(
            db_host=cfg.DB_HOST, db_port=getattr(cfg, "DB_PORT", 1433),
            db_database=cfg.DB_NAME, db_username=cfg.DB_USER,
            db_password=cfg.DB_PASS, login_timeout=5, timeout=5,
        )
        with conn as d:
            d.fech_dataframe("SELECT 1")
        out["sql"] = {"ok": True, "detail": f"Connexió correcta ({time.time() - t:.1f}s) · {cfg.DB_HOST}"}
    except Exception as e:
        out["sql"] = {"ok": False, "detail": f"{type(e).__name__}: {str(e).splitlines()[0][:150]}"}

    try:
        if os.path.isdir(Z_CHECK_DIR):
            os.listdir(Z_CHECK_DIR)  # actually reach the share, not just a stale mapping
            out["z"] = {"ok": True, "detail": "Accessible · " + Z_CHECK_DIR}
        else:
            out["z"] = {"ok": False, "detail": "No trobat: " + Z_CHECK_DIR}
    except Exception as e:
        out["z"] = {"ok": False, "detail": f"{type(e).__name__}: {str(e)[:150]}"}
    return jsonify(out)


# Lists the categorized labs (Mapa_Acords 'Laboratori' column, only rows with
# an accord) for the Pas 1 and Pas 2 dropdowns, so the user can generate a
# single-lab report and, as a separate step, its split file.
# Fails soft if the Z: drive / file is unreachable.
MAPA_ACORDS_XLSX = r"Z:\Compres\INDÚSTRIA FARMACÈUTICA\Mapa_Acords.xlsx"


@app.route("/labs")
def labs():
    try:
        import pandas as pd
        mapa = pd.read_excel(MAPA_ACORDS_XLSX, sheet_name="Mapa")
        # Same criterion as index.py DataFilters: drop labs whose Acord says NO.
        with_acord = ~mapa["Acord"].astype(str).str.contains("NO", na=False)
        names = sorted({str(x).strip() for x in mapa.loc[with_acord, "Laboratori"].dropna()
                        if str(x).strip()})
        return jsonify({"ok": True, "labs": names})
    except Exception as e:
        return jsonify({"ok": False, "labs": [],
                        "error": f"{type(e).__name__}: {str(e).splitlines()[0][:150]}"})


# ---------------------------------------------------------------------------
# Editor de productes: PVL/IVA/MARCA/SUBMARCA sobre ClickHouse dim_producte, el
# store gobernat que Pas 1 llegeix (index.py.transformNumValue). Editar aquí és
# el que substitueix el treball manual al YTD Excel i destrava el tancament
# autònom. store_ch s'importa mandrós dins de cada ruta: una incidència de
# ClickHouse no pot impedir que arrenqui el panell.
# ---------------------------------------------------------------------------
@app.route("/editor")
def editor():
    return render_template("editor.html")


@app.route("/api/productes")
def api_productes():
    try:
        import store_ch
        rows = store_ch.cerca_productes(
            q=request.args.get("q", ""),
            only_missing=request.args.get("missing", "") in ("1", "true", "on"),
            limit=request.args.get("limit", 200))
        return jsonify({"ok": True, "rows": rows})
    except Exception as e:
        return jsonify({"ok": False, "error": f"{type(e).__name__}: {str(e).splitlines()[0][:200]}"})


@app.route("/api/productes/desa", methods=["POST"])
def api_productes_desa():
    try:
        import store_ch
        edits = (request.get_json(force=True) or {}).get("edits", [])
        n = store_ch.desa_producte_attr(edits)
        return jsonify({"ok": True, "n": n})
    except Exception as e:
        return jsonify({"ok": False, "error": f"{type(e).__name__}: {str(e).splitlines()[0][:200]}"})


# ---------------------------------------------------------------------------
# Analítica Metabase (Cierre) embeguda al panell. Metabase ja està connectat a
# ClickHouse (DB 'tancament'). Aquí es llisten els dashboards de la col·lecció
# Cierre i s'embeuen via public link. NO s'activa el public sharing des d'aquí:
# és una decisió d'exposició de dades que pren l'usuari (Admin > Settings >
# Public sharing). Si està desactivat, l'endpoint ho diu i la pàgina ho explica.
# ---------------------------------------------------------------------------
# Els secrets (MB_PASS, etc.) surten de l'entorn o del db_config.py (git-ignored),
# MAI hardcodejats. El repo public només porta placeholders.
def _secret(name, default=""):
    v = os.environ.get(name)
    if v:
        return v
    try:
        import db_config
        return getattr(db_config, name, None) or default
    except Exception:
        return default

MB_URL  = os.environ.get("MB_URL", "http://127.0.0.1:3000").rstrip("/")   # app -> Metabase (servidor)
MB_USER = _secret("MB_USER", "admin@example.com")
MB_PASS = _secret("MB_PASS", "")
MB_PORT = os.environ.get("MB_PORT", "3000")   # port de Metabase per al NAVEGADOR (mateix host que el panell)


def _mb_call(path, data=None, session=None, method=None):
    import urllib.request, urllib.error
    h = {"Content-Type": "application/json"}
    if session:
        h["X-Metabase-Session"] = session
    req = urllib.request.Request(
        MB_URL + path,
        data=(json.dumps(data).encode() if data is not None else None),
        headers=h, method=method or ("POST" if data is not None else "GET"))
    try:
        return json.loads(urllib.request.urlopen(req, timeout=20).read().decode() or "null")
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"MB {e.code}: {e.read().decode()[:200]}")


@app.route("/analitica")
def analitica():
    return render_template("analitica.html")


@app.route("/api/mb/cierre")
def api_mb_cierre():
    try:
        # URL de Metabase per al NAVEGADOR: mateix host que el panell, port de MB.
        mb_base = f"http://{request.host.split(':')[0]}:{MB_PORT}"
        S = _mb_call("/api/session", {"username": MB_USER, "password": MB_PASS})["id"]
        ul = _mb_call("/api/user", session=S)
        ul = ul.get("data", ul) if isinstance(ul, dict) else ul
        umap = {u["id"]: (u.get("common_name")
                          or (u.get("first_name", "") + " " + u.get("last_name", "")).strip()
                          or u.get("email", "")) for u in ul}
        cols = _mb_call("/api/collection", session=S)

        def subtree(match):
            root = next((c["id"] for c in cols if match in (c.get("name") or "")
                         and (c.get("location") in ("/", None, "")) and not c.get("personal_owner_id")), None)
            ids = {root} if root is not None else set()
            if root is not None:
                ids |= {c["id"] for c in cols if f"/{root}/" in (c.get("location") or "")}
            return root, ids

        cierre_root, cierre_ids = subtree("Cierre")
        # NOMÉS l'àrea Cierre, agrupat PER PERSONA (creador). Res de Pricing/demos.
        # El setting enable-public-sharing retorna None encara que estigui ON, així que
        # ens basem en el public_uuid: l'usem si existeix o el creem.
        dashboards, sharing_off = [], False
        for d in _mb_call("/api/dashboard", session=S):
            if d.get("archived") or d.get("collection_id") not in cierre_ids:
                continue
            uuid = d.get("public_uuid")
            if not uuid:
                try:
                    uuid = _mb_call(f"/api/dashboard/{d['id']}/public_link", {}, session=S, method="POST").get("uuid")
                except Exception as e:
                    if "sharing" in str(e).lower() or "not enabled" in str(e).lower():
                        sharing_off = True
                    continue
            if not uuid:
                continue
            group = umap.get(d.get("creator_id")) or "General"
            dashboards.append({"id": d["id"], "name": d.get("name", "(sense nom)"),
                               "group": group, "uuid": uuid})
        dashboards.sort(key=lambda x: (x["group"].lower(), x["name"].lower()))
        return jsonify({"ok": True, "sharing": (not sharing_off), "mb_base": mb_base,
                        "cierre_collection": cierre_root, "dashboards": dashboards})
    except Exception as e:
        return jsonify({"ok": False, "error": f"{type(e).__name__}: {str(e).splitlines()[0][:200]}"})


# Reads a CSV into {exists, columns, rows, count, updated} for the data tabs.
def _csv_json(path):
    if not os.path.exists(path):
        return {"exists": False, "columns": [], "rows": [], "count": 0}
    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    cols = rows[0] if rows else []
    data = rows[1:] if len(rows) > 1 else []
    return {
        "exists": True, "columns": cols, "rows": data, "count": len(data),
        "updated": time.strftime("%Y-%m-%d %H:%M", time.localtime(os.path.getmtime(path))),
    }


# Product-conditions CSV as JSON (2nd tab).
@app.route("/conditions")
def conditions():
    return jsonify(_csv_json(CONDITIONS_CSV))


# Per-lab lost-money report as JSON (3rd tab).
@app.route("/lostmoney")
def lostmoney():
    return jsonify(_csv_json(LOST_MONEY_CSV))


# Detector de referencies (4a pestanya).
#   GET  -> l'ultim informe generat, com les altres pestanyes.
#   POST -> puja un BIExportGrid de BIFarma, l'analitza i reescriu l'informe.
# Nomes compara existencia de la descripcio del producte al mestre SQL; el
# perque i les seves limitacions son al capdamunt de detector.py.
@app.route("/detector", methods=["GET", "POST"])
def detector_refs():
    if request.method == "GET":
        return jsonify(_csv_json(DETECTOR_CSV))

    f = request.files.get("export")
    if f is None or not f.filename:
        return jsonify({"ok": False, "error": "No has triat cap fitxer."}), 400
    if not f.filename.lower().endswith((".xlsx", ".xlsm")):
        return jsonify({"ok": False, "error": "Ha de ser un .xlsx exportat de BIFarma."}), 400

    # A un fitxer temporal amb el nom original: detector.py el retorna al resum
    # i pandas necessita un path, no el stream.
    tmp_dir = os.path.join(BASE_DIR, "output", "_upload")
    os.makedirs(tmp_dir, exist_ok=True)
    safe = "".join(c if (c.isalnum() or c in " .-_") else "_" for c in f.filename)
    tmp = os.path.join(tmp_dir, safe)
    f.save(tmp)

    try:
        import detector
        cols, rows, resum = detector.analitza(tmp)
    except Exception as e:
        return jsonify({"ok": False,
                        "error": f"{type(e).__name__}: {str(e).splitlines()[0][:300]}"}), 500
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass  # el temporal no es critic

    os.makedirs(os.path.dirname(DETECTOR_CSV), exist_ok=True)
    with open(DETECTOR_CSV, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(cols)
        w.writerows(rows)

    out = _csv_json(DETECTOR_CSV)
    out["ok"] = True
    out["resum"] = resum
    return jsonify(out)


# Downloads an output CSV, optionally filtered by laboratory (?lab=NAME).
# (Especialitat is decided at generation time via the Pas 1 "Només parafarmàcia"
# option, not here.)
@app.route("/download/<key>")
def download(key):
    conf = {"conditions": CONDITIONS_CSV, "lostmoney": LOST_MONEY_CSV,
            "detector": DETECTOR_CSV}
    if key not in conf:
        return "Baixada desconeguda.", 404
    path = conf[key]
    if not os.path.exists(path):
        if key == "detector":
            return "Encara no hi ha informe. Puja un export de BIFarma al detector.", 404
        return "El CSV encara no existeix. Executa el pas 1 (Generar mestre).", 404

    lab = request.args.get("lab", "").strip()
    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    header = rows[0] if rows else []
    body = rows[1:]
    # 'Laboratorio Categorizado' is the real classification and the column the
    # tables group by; plain 'Laboratorio' is only a fallback for a CSV written
    # by an older run.
    lab_i = next((header.index(c) for c in ("Laboratorio Categorizado", "Laboratorio")
                  if c in header), None)

    if lab and lab_i is not None:
        body = [r for r in body if r[lab_i].strip() == lab]

    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(header)
    writer.writerows(body)
    data = buf.getvalue().encode("utf-8-sig")

    base = os.path.splitext(os.path.basename(path))[0]
    tag = "_" + "".join(c if c.isalnum() else "-" for c in lab) if lab else ""
    return send_file(io.BytesIO(data), mimetype="text/csv",
                     as_attachment=True, download_name=f"{base}{tag}.csv")


# Downloads the Pas-1 master Excel for the current period/lab selection.
# The source file (output/df_bifarma_output{period}{lab_tag}.xlsx) is renamed
# on the fly to PARAFARMACIA_{MM.AAAA}_{YTD|RES}[_LAB].xlsx.
@app.route("/download/master")
def download_master():
    period = request.args.get("period", "")
    lab = request.args.get("laboratori", "").strip()
    year = request.args.get("year", "")
    month = request.args.get("month", "")

    # Rebuild index.py's output filename (same lab-tag sanitization).
    lab_tag = "_" + "".join(c if c.isalnum() else "_" for c in lab)[:40] if lab else ""
    src = os.path.join(BASE_DIR, "output", f"df_bifarma_output{period}{lab_tag}.xlsx")
    if not os.path.exists(src):
        return "El fitxer del pas 1 encara no existeix. Executa el pas 1 (Generar mestre).", 404

    try:
        stamp = f"{int(month):02d}.{int(year)}"
    except (ValueError, TypeError):
        stamp = time.strftime("%m.%Y")
    tag = "YTD" if period == "YTD" else "RES"
    name = f"PARAFARMACIA_{stamp}_{tag}"
    if lab:
        name += "_" + "".join(c if c.isalnum() else "-" for c in lab)
    return send_file(src, as_attachment=True, download_name=f"{name}.xlsx")


# Launches the requested script and streams its output as SSE. EventSource uses
# GET, so params arrive in the query string. Rejects concurrent runs via the lock.
@app.route("/run")
def run():
    # Read everything off `request` HERE: the generator below runs after the
    # request context is gone, so touching `request` inside it would blow up.
    params = request.args.to_dict()
    cmd = _build_command(params)
    who = request.remote_addr or "?"

    def generate():
        if cmd is None:
            yield _sse({"type": "error", "line": "Script no vàlid."})
            yield _sse({"type": "done", "code": -1})
            return
        if not _run_lock.acquire(blocking=False):
            job = dict(_current_job)  # copy: the other thread may clear it mid-read
            if job:
                mins = max(0, int((time.time() - job["started"]) / 60))
                busy = (f"Ja hi ha una execució en marxa: {job['script']}, iniciada des de "
                        f"{job['by']} fa {mins} min. Espera que acabi.")
            else:
                busy = "Ja hi ha una execució en marxa. Espera que acabi."
            yield _sse({"type": "error", "line": busy})
            yield _sse({"type": "done", "code": -1})
            return
        _current_job.update(script=params.get("script", "?"), started=time.time(), by=who)
        proc = None
        handed_off = False
        try:
            yield _sse({"type": "cmd", "line": "$ " + " ".join(cmd)})
            # Force UTF-8 in the child so Catalan/accented prints never crash.
            env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
            proc = subprocess.Popen(
                cmd, cwd=BASE_DIR,
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                bufsize=1, universal_newlines=True,
                encoding="utf-8", errors="replace", env=env,
            )
            for line in proc.stdout:
                yield _sse({"type": "log", "line": line.rstrip("\n")})
            proc.wait()
            yield _sse({"type": "done", "code": proc.returncode})
        except GeneratorExit:
            # Browser tab closed mid-run: the subprocess keeps running, so the
            # lock must stay held until it ends. Hand it to a watcher thread;
            # otherwise a second run could start while Excel/COM is still busy.
            if proc is not None and proc.poll() is None:
                handed_off = True

                def _release_when_done(p):
                    p.wait()
                    _current_job.clear()
                    _run_lock.release()

                threading.Thread(target=_release_when_done, args=(proc,), daemon=True).start()
            raise
        except Exception as e:
            yield _sse({"type": "error", "line": str(e)})
            yield _sse({"type": "done", "code": -1})
        finally:
            # When handed off, the watcher thread owns both the lock and the
            # job state until the subprocess actually ends.
            if not handed_off:
                _current_job.clear()
                _run_lock.release()

    return Response(
        generate(),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


# Prints the addresses the panel is reachable at, so whoever starts it on the
# server can just copy one and share it with the team.
def _banner():
    import socket
    print("=" * 66, flush=True)
    print("  PANELL DE TANCAMENT en marxa", flush=True)
    if HOST in ("0.0.0.0", ""):
        name = socket.gethostname()
        print(f"    en aquest equip:  http://127.0.0.1:{PORT}", flush=True)
        print(f"    des de la xarxa:  http://{name}:{PORT}", flush=True)
        try:
            for ip in sorted({a[4][0] for a in socket.getaddrinfo(name, None, socket.AF_INET)}):
                print(f"                      http://{ip}:{PORT}", flush=True)
        except OSError:
            pass  # no DNS/hostname resolution: the name above is enough
    else:
        print(f"    http://{HOST}:{PORT}", flush=True)
    print("  Tanca aquesta finestra per aturar el panell.", flush=True)
    print("=" * 66, flush=True)


if __name__ == "__main__":
    _banner()
    # threaded=True so the long SSE request doesn't block the page/static files.
    app.run(host=HOST, port=PORT, debug=False, threaded=True)
