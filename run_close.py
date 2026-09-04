# -*- coding: utf-8 -*-
"""
run_close.py -- Orquestador headless del tancament (sin panel Flask).
================================================================================
Encadena el cierre mensual de punta a punta, DESATENDIDO, reutilizando las mismas
CLIs que el panel dispara a botón:

    Pas 1  (master)  ->  python -u index.py    --rappel .. --period .. [--year --month] [--only-para]
    Pas 2  (split)   ->  python -u clsSplit.py --period .. [--year --month] --steps split,fee,enqueue

Con 'enqueue', Pas 2 deja los .xlsx en splitFiles/inbox y el worker Excel-COM
(com_worker.py --watch, el ancla) los refresca async y los devuelve a pasteFiles.
Para refrescar en proceso (sin worker) usa --refresh-mode inprocess.

Es la pieza que convierte "pasos ya headless" en "cierre orquestable": lo dispara
una Tarea programada o un DAG, propaga el codigo de salida (para en el primer
fallo) y deja un log con marca de tiempo en logs/.

NO es 100% autonomo: la curacion manual de MARCA/SUBMARCA (productos nuevos
marcados "New!") sigue necesitando una persona hasta que exista el editor de
gobernanza. Este orquestador avisa de cuantos quedan, pero no los inventa.

USO
    Cierre mensual completo del ultimo mes tancat:
        python run_close.py
    Un mes concreto, solo PARAFARMACIA, refresco async via worker:
        python run_close.py --year 2026 --month 7 --only-para
    Solo el master (Pas 1), sin split:
        python run_close.py --phases master
    Split refrescando en proceso (sin worker):
        python run_close.py --refresh-mode inprocess

Requisitos: acceso a Z: y a la BD BIFarma (como Pas 1). Codigo de salida != 0 si
cualquier fase falla.
"""
import os
import sys
import shutil
import subprocess
from datetime import datetime, date
from dateutil.relativedelta import relativedelta

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable  # el mismo interprete que lanza el orquestador
LOG_DIR = os.path.join(BASE_DIR, "logs")

# Igual que el panel: fuerza UTF-8 en el hijo para que los prints con acentos
# (catalan) nunca revienten por la consola cp1252.
CHILD_ENV = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")


def _run_phase(name, cmd, logf):
    """Ejecuta una fase como subproceso, transmitiendo su stdout a la consola y
    al log en vivo. Devuelve el codigo de salida del hijo."""
    line = f"$ {' '.join(cmd)}"
    _emit(logf, f"\n=== {name} ===")
    _emit(logf, line)
    proc = subprocess.Popen(
        cmd, cwd=BASE_DIR,
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        bufsize=1, universal_newlines=True,
        encoding="utf-8", errors="replace", env=CHILD_ENV,
    )
    for out in proc.stdout:
        _emit(logf, out.rstrip("\n"))
    proc.wait()
    _emit(logf, f"--- {name}: codi de sortida {proc.returncode} ---")
    return proc.returncode


def _emit(logf, text):
    print(text, flush=True)
    logf.write(text + "\n")
    logf.flush()


def _master_cmd(a):
    cmd = [PY, "-u", os.path.join(BASE_DIR, "index.py"),
           "--rappel", a.rappel, "--period", a.period]
    if a.year:
        cmd += ["--year", str(a.year)]
    if a.month:
        cmd += ["--month", str(a.month)]
    if a.only_para:
        cmd += ["--only-para"]
    return cmd


def _split_cmd(a):
    # enqueue = async via worker; inprocess = refresh COM aqui mismo; none = no refresca
    step_refresh = {"enqueue": "enqueue", "inprocess": "refresh", "none": ""}[a.refresh_mode]
    steps = ",".join(s for s in ("split", "fee", step_refresh) if s)
    cmd = [PY, "-u", os.path.join(BASE_DIR, "clsSplit.py"),
           "--period", a.period, "--steps", steps]
    if a.year:
        cmd += ["--year", str(a.year)]
    if a.month:
        cmd += ["--month", str(a.month)]
    return cmd


def _month_stamp(a):
    d = date(a.year, a.month, 1) if (a.year and a.month) else (date.today() - relativedelta(months=1))
    return format(d, "%m.%Y")


def _publish_master(a, logf):
    """Publica el master de Pas 1 (output/df_bifarma_output{period}.xlsx) al master
    de Z: (Parafarmacia MM.AAAA[ YTD].xlsx) que Pas 2 llegeix. Abans es feia a mà
    (descarregar + curar MARCA a l'Excel + desar a Z:); ara la curació MARCA viu a
    ClickHouse (editor + dim_producte) i Pas 1 ja l'incorpora, així que es pot
    automatitzar. Per seguretat NO sobreescriu un master existent tret de
    --overwrite-master. Retorna 0 ok, !=0 error."""
    from clsSplit import MAIN_FILE_DIR
    src = os.path.join(BASE_DIR, "output", f"df_bifarma_output{a.period}.xlsx")
    tmp = " YTD" if a.period == "YTD" else ""
    dst = os.path.join(MAIN_FILE_DIR, f"Parafarmacia {_month_stamp(a)}{tmp}.xlsx")
    _emit(logf, f"\n=== publish (master -> Z:) ===")
    if not os.path.exists(src):
        _emit(logf, f"ERROR publish: no existeix {src}. Pas 1 no ha escrit el master?")
        return 2
    if os.path.exists(dst) and not a.overwrite_master:
        _emit(logf, f"ATURAT publish: el master de Z: ja existeix:\n  {dst}\n"
                    f"  No es sobreescriu (evita trepitjar un tancament real). "
                    f"Usa --overwrite-master per forçar-ho.")
        return 3
    shutil.copy(src, dst)
    _emit(logf, f"publish OK: {os.path.basename(src)} -> {dst}")
    return 0


def main(argv=None):
    import argparse
    ap = argparse.ArgumentParser(description="Orquestador headless del tancament (Pas 1 -> Pas 2).")
    ap.add_argument("--rappel", default="BIFARMA", help='"BIFARMA" o "bifarma" (con BAJAS)')
    ap.add_argument("--period", default="", choices=["", "YTD"], help='"" mensual, "YTD" acumulat')
    ap.add_argument("--year", type=int, default=None, help="Any del mes a processar (opcional)")
    ap.add_argument("--month", type=int, default=None, help="Mes a processar 1-12 (opcional)")
    ap.add_argument("--only-para", action="store_true",
                    help="PARAFARMACIA + especialitat pactada (com el checkbox del panell)")
    ap.add_argument("--phases", default="master,publish,split",
                    help="Coma-separat: master, publish, split (per defecte les tres)")
    ap.add_argument("--refresh-mode", default="enqueue", choices=["enqueue", "inprocess", "none"],
                    help="enqueue=async via worker (defecte); inprocess=refresca COM aqui; none=no refresca")
    ap.add_argument("--overwrite-master", action="store_true",
                    help="publish: sobreescriu el master de Z: encara que ja existeixi")
    a = ap.parse_args(argv)
    phases = [p.strip() for p in a.phases.split(",") if p.strip()]

    os.makedirs(LOG_DIR, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = os.path.join(LOG_DIR, f"close_{stamp}.log")

    with open(log_path, "w", encoding="utf-8") as logf:
        _emit(logf, f"Tancament orquestrat · {stamp} · fases={phases} · refresh={a.refresh_mode}")

        if "master" in phases:
            code = _run_phase("Pas 1 · master (index.py)", _master_cmd(a), logf)
            if code != 0:
                _emit(logf, f"\nATURAT: Pas 1 ha fallat (codi {code}). No es continua.")
                return code

        if "publish" in phases:
            code = _publish_master(a, logf)
            if code != 0:
                _emit(logf, f"\nATURAT: publish ha fallat (codi {code}). No es continua amb el split.")
                return code

        if "split" in phases:
            code = _run_phase("Pas 2 · split (clsSplit.py)", _split_cmd(a), logf)
            if code != 0:
                _emit(logf, f"\nATURAT: Pas 2 ha fallat (codi {code}).")
                return code
            if a.refresh_mode == "enqueue":
                _emit(logf, "\nNota: fitxers encuats a splitFiles/inbox. El worker Excel-COM "
                            "(com_worker.py --watch) els refrescara i els tornara a pasteFiles.")

        _emit(logf, f"\nOK: tancament orquestrat complet. Log: {log_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
