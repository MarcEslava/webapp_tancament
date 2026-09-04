# -*- coding: utf-8 -*-
"""
com_worker.py -- El ancla Excel-COM, aislada.
================================================================================
Refresca ficheros .xlsx/.xlsm vía Excel COM. Es el ÚNICO componente del cierre
que no se puede desacoplar de Windows: necesita Excel instalado y corre en serie.
Todo lo demás (Pas 1, split, histórico, analítica) es headless alrededor de esto.

CONTRATO LIMPIO (por eso se puede aislar):
    entrada  = una carpeta de .xlsx/.xlsm  (o una lista explícita de rutas)
    acción   = abrir · RefreshAll · esperar queries async · ocultar 'Acuerdo book' · guardar
    salida   = los mismos ficheros, recalculados  (+ código de salida != 0 si falló alguno)

No guarda estado en memoria entre ficheros: cada uno se abre por ruta y se cierra.
Extraído tal cual de clsSplit.refresh_workbooks / _refresh_one / _start_excel, para
que exista UNA sola implementación del refresh y sea disparable por el pipeline.

USO
    Un disparo, carpeta por defecto (los 78 split del mes):
        python com_worker.py
    Un disparo, carpeta concreta:
        python com_worker.py --dir "D:\\ruta\\pasteFiles"
    Un disparo, ficheros concretos (modo single-lab):
        python com_worker.py --files a.xlsx b.xlsx
    Servicio (cola): vigila un inbox, refresca lo que caiga y lo mueve a --done:
        python com_worker.py --watch "D:\\inbox" --done "D:\\refrescados" --interval 5

REQUISITOS: Windows + Microsoft Excel + pywin32. NO contenerizable.

Author: Marc Eslava
"""
import os
import sys
import time
import argparse

import win32com.client
try:
    import pythoncom  # inicialización COM explícita (seguro en bucles largos)
except Exception:  # pragma: no cover - solo existe en Windows/pywin32
    pythoncom = None

# 0x800706BA = "RPC server unavailable" -> el proceso de Excel ha muerto.
_RPC_DEAD = -2147023174

# Carpeta por defecto: idéntica a clsSplit (BASE_DIR/splitFiles/pasteFiles),
# para ser drop-in de `python clsSplit.py --steps refresh`.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_DIR = os.path.join(BASE_DIR, "splitFiles", "pasteFiles")

_EXTS = (".xlsx", ".xlsm")


# --- COM primitivos (extraídos de clsSplit, sin cambios de comportamiento) ----

def _start_excel():
    """Instancia oculta de Excel con alertas desactivadas."""
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    return excel


def _refresh_one(excel, full_path):
    """Abre, refresca links/fórmulas, oculta 'Acuerdo book', guarda y cierra."""
    wb = excel.Workbooks.Open(full_path, UpdateLinks=False)
    wb.RefreshAll()
    # RefreshAll() es async para queries/links en segundo plano. Hay que esperar
    # a que terminen antes de guardar: guardar/cerrar a mitad de refresco
    # desestabiliza Excel y lo tira (RPC unavailable).
    excel.CalculateUntilAsyncQueriesDone()
    # Ocultar la hoja de datos es específico de los ficheros del split; se tolera
    # su ausencia para que el worker pueda refrescar cualquier libro.
    try:
        wb.Worksheets("Acuerdo book").Visible = 0
    except Exception:
        pass
    wb.Save()
    wb.Close(SaveChanges=True)


# --- API pública --------------------------------------------------------------

def scan_folder(directory):
    """Rutas absolutas de todos los .xlsx/.xlsm de una carpeta."""
    return [os.path.join(directory, f) for f in os.listdir(directory)
            if f.lower().endswith(_EXTS)]


def refresh_files(paths, log=print):
    """Refresca una lista de ficheros con UNA instancia de Excel, respawneándola
    si el proceso muere para que un crash no arrastre al resto.

    Devuelve (ok, failed): listas de nombres de fichero.
    """
    targets = [os.path.abspath(p) for p in paths]
    if not targets:
        log("No hi ha cap fitxer per refrescar.")
        return [], []

    if pythoncom is not None:
        pythoncom.CoInitialize()
    excel = _start_excel()
    ok, failed = [], []
    try:
        for full_path in targets:
            file_name = os.path.basename(full_path)
            try:
                _refresh_one(excel, full_path)
                log(file_name)
                ok.append(file_name)
            except Exception as e:
                # Excel murió (RPC): respawnear y reintentar este fichero una vez.
                if getattr(e, "hresult", None) == _RPC_DEAD or _RPC_DEAD in getattr(e, "args", ()):
                    log(f"  Excel crashed on {file_name}, restarting and retrying...")
                    try:
                        excel.Quit()
                    except Exception:
                        pass
                    excel = _start_excel()
                    try:
                        _refresh_one(excel, full_path)
                        log(file_name)
                        ok.append(file_name)
                        continue
                    except Exception as e2:
                        e = e2
                log(f"  SKIP {file_name}: {e}")
                failed.append(file_name)
    finally:
        try:
            excel.Quit()
        except Exception:
            pass
        if pythoncom is not None:
            pythoncom.CoUninitialize()

    if failed:
        log(f"\n{len(failed)} file(s) failed to refresh:")
        for f in failed:
            log(f"  {f}")
    return ok, failed


def refresh_folder(directory, log=print):
    """Refresca todos los .xlsx/.xlsm de una carpeta."""
    return refresh_files(scan_folder(directory), log=log)


# --- Modo servicio (cola por carpeta) -----------------------------------------

def watch(inbox, done=None, interval=5, log=print):
    """Vigila `inbox`; refresca cada .xlsx/.xlsm que aparezca y, si se da `done`,
    lo mueve allí. Bucle hasta Ctrl+C. Convierte el ancla en un servicio: el
    pipeline deja ficheros en `inbox`, el worker los procesa. Sin acoplamiento.
    """
    import shutil
    os.makedirs(inbox, exist_ok=True)   # el worker pot arrencar abans que el productor creï la cua
    if done:
        os.makedirs(done, exist_ok=True)
    log(f"[worker] vigilando {inbox}  (intervalo {interval}s)"
        + (f"  ->  {done}" if done else "  (in-place)"))
    try:
        while True:
            batch = scan_folder(inbox)
            if batch:
                log(f"[worker] {len(batch)} fitxer(s) a refrescar")
                ok, failed = refresh_files(batch, log=log)
                if done:
                    for p in batch:
                        name = os.path.basename(p)
                        if name in ok:
                            try:
                                shutil.move(p, os.path.join(done, name))
                            except Exception as e:
                                log(f"  no puc moure {name}: {e}")
            time.sleep(interval)
    except KeyboardInterrupt:
        log("[worker] aturat")


# --- CLI ----------------------------------------------------------------------

def main(argv=None):
    ap = argparse.ArgumentParser(description="Worker Excel-COM: refresca .xlsx/.xlsm (el ancla del cierre).")
    ap.add_argument("--dir", help="carpeta a refrescar (por defecto: splitFiles/pasteFiles)")
    ap.add_argument("--files", nargs="+", help="ficheros concretos a refrescar")
    ap.add_argument("--watch", metavar="INBOX", help="modo servicio: vigila esta carpeta en bucle")
    ap.add_argument("--done", help="con --watch: mueve aquí los ficheros refrescados")
    ap.add_argument("--interval", type=int, default=5, help="con --watch: segundos entre sondeos")
    args = ap.parse_args(argv)

    if args.watch:
        watch(args.watch, done=args.done, interval=args.interval)
        return 0

    if args.files:
        _, failed = refresh_files(args.files)
    else:
        directory = args.dir or DEFAULT_DIR
        if not os.path.isdir(directory):
            print(f"No existe la carpeta: {directory}", file=sys.stderr)
            return 2
        _, failed = refresh_folder(directory)

    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
