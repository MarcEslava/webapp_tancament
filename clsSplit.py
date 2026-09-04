import os
import pandas as pd
import shutil
import datetime
from dateutil.relativedelta import relativedelta
import win32com.client



'''
**************************************************************************************************************************************************************************************
The purpose of this program is the automatization of the process we use to split diferent lab data into a specific excel file, this way we can automate the
process of sending information to diferent providers.

Close all excel files before executing this program

Author: Marc Eslava

**************************************************************************************************************************************************************************************
'''
MAIN_FILE_DIR = r"Z:\Compres\INDÚSTRIA FARMACÈUTICA\01. CARPETES LABORATORIS\102. Informes i Rappel"


# Builds the per-campaign SEGUIMENT file path (source of labs + fee conditions).
# The year comes from the processing month, so it stays valid across campaigns.
def seguiment_file(year):
    return os.path.join(MAIN_FILE_DIR, f"a.SEGUIMENT LABS {year}.xlsx")

# Split-file lab names that don't match their SEGUIMENT label (key = split name upper).
FEE_ALIASES = {
    "OCCITANE": "L'OCCITANE",
    "AVOGEL": "VOGEL",
    "VERMONTPHARMA": "VERMONT PHARMA",
}

# Labs to generate that are not (yet) in the SEGUIMENT file. Appended to the
# SEGUIMENT lab list. They have no fee conditions, so their Fee sheet is flagged.
EXTRA_LABS = [
    "EXELTIS",
    "SENSILIS",
    "ECOCEUTICS",
    "NO ENVIAR",
    "DERMOFARM ORIGINAL",
]


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PASTE_DIR = os.path.join(BASE_DIR, "splitFiles", "pasteFiles")
# Cua del worker Excel-COM (com_worker.py --watch): el pas 'enqueue' hi mou els
# fitxers ja fets (split+fee) i el worker els refresca i els torna a PASTE_DIR.
INBOX_DIR = os.path.join(BASE_DIR, "splitFiles", "inbox")
# Single structural template copied once per lab to build every output file.
TEMPLATE_FILE = os.path.join(BASE_DIR, "splitFiles", "template.xlsx")

ACCUMULATE_CSV = os.path.join(MAIN_FILE_DIR, "output", "accumulated_labs.csv")


PERIODO = ""  # Set to "YTD" for year-to-date files, or "" for regular monthly files
CALCULATE_FEE = False  # Set to True on months where the Fee sheet must be generated


class clsSplit:
    # Computes current and previous month dates.
    # year/month: target processing month. When omitted, defaults to the last
    # complete month (today - 1 month) so the monthly run needs no arguments;
    # a future frontend can pass them explicitly.
    # Lab names are extracted in run() once temporality is known.
    def __init__(self, year=None, month=None) -> None:
        if year and month:
            this_month = datetime.date(year, month, 1)
        else:
            this_month = (datetime.datetime.today() - relativedelta(months=1)).date()
        last_month = this_month - relativedelta(months=1)
        self.this_month = format(this_month, "%m.%Y")
        self.last_month = format(last_month, "%m.%Y")
        # SEGUIMENT file is per-campaign-year; derived from the processing month.
        self.seguiment_file = seguiment_file(this_month.year)
        self.labs = []

    # Main entry point. Copies the template per lab, reads the master Excel
    # of the month from Z: and writes each lab's Acuerdo book.
    # temporality: optional suffix like "YTD" for the file naming.
    # only_lab: restrict the run to that single lab. It is the SAME crossing as
    # the full run against the SAME master -- just one book written -- so a
    # single-lab split never depends on Pas 1 having been run for that lab.
    # Returns the list of files written.
    def run(self, temporality: str = "", only_lab=None) -> list:
        self.tmp = " YTD" if temporality == "YTD" else ""
        self.only_lab = (only_lab or "").strip() or None
        if self.only_lab:
            self.labs = [self.only_lab]
            print(f"Un sol laboratori: {self.only_lab}")
        else:
            # Lab list comes from the SEGUIMENT file (not per-lab template files).
            self.labs = self._read_labs()
            print(f"{len(self.labs)} labs read from SEGUIMENT file")
        # A lab name becomes a file name, so anything Windows can't take in a
        # path (e.g. the "#N/D" placeholder) is dropped here instead of blowing
        # up mid-run inside shutil.copy.
        illegal = set(r'\/:*?"<>|')
        net = [l for l in self.labs if not (set(l) & illegal)]
        if len(net) != len(self.labs):
            fora = [l for l in self.labs if set(l) & illegal]
            print(f"AVIS: {len(fora)} nom(s) de lab no poden ser un nom de fitxer i "
                  f"s'ignoren: {fora}", flush=True)
            self.labs = net
        # Each phase depends on the previous one: abort on the first failure
        # (exit code 1 so the panel/CLI skips the remaining steps) instead of
        # running the next phase against missing/stale state.
        for label, phase in (("copying files", self._copy_files),
                             ("reading files", self._read),
                             ("transforming data", self._transform)):
            try:
                phase()
            except Exception as e:
                print(f"Error {label}: {type(e).__name__}: {e}", flush=True)
                raise SystemExit(1)
        return self._written

    # Reads the list of labs to generate from the SEGUIMENT file:
    # column A of the Parafarmacia sheet plus column A of the EFG sheet.
    # Preserves the file's casing and order; drops blanks and TOTAL GENERAL.
    def _read_labs(self):
        labs = []
        seen = set()

        def add(lab):
            if lab is None or (isinstance(lab, float) and pd.isna(lab)):
                return
            name = str(lab).strip()
            key = name.upper()
            if key in ("", "NAN", "TOTAL GENERAL") or key in seen:
                return
            seen.add(key)
            labs.append(name)

        pf = pd.read_excel(self.seguiment_file, sheet_name="Parafarmacia", header=1, usecols="A")
        for v in pf.iloc[:, 0]:
            add(v)
        efg = pd.read_excel(self.seguiment_file, sheet_name="EFG", header=1, usecols="A")
        for v in efg.iloc[:, 0]:
            add(v)
        # Labs not in SEGUIMENT that must still be generated.
        for lab in EXTRA_LABS:
            add(lab)
        return labs

    # Builds every output file by copying the single template once per lab,
    # naming it with the current month in PASTE_DIR.
    def _copy_files(self):
        if not os.path.isfile(TEMPLATE_FILE):
            raise FileNotFoundError(f"Template not found: {TEMPLATE_FILE}")
        for lab in self.labs:
            dst = os.path.join(PASTE_DIR, f"PARAFARMACIA {self.this_month}{self.tmp} - {lab}.xlsx")
            shutil.copy(TEMPLATE_FILE, dst)
        print(f"{len(self.labs)} files copied to pasteFiles folder from template")

    # Reads the master Parafarmacia file. Mapa_Acords is NOT needed here: the
    # master already carries the categorized lab label in
    # 'Laboratorio Categorizado' (see _replace_accord).
    def _read(self):
        main_file_path = os.path.join(MAIN_FILE_DIR, f"Parafarmacia {self.this_month}{self.tmp}.xlsx")
        print(f"Reading main file: {main_file_path}", flush=True)
        self.main_file = pd.read_excel(main_file_path, index_col="Cod Unif", sheet_name="SI Acord book")
        print("Reading file done")

    def _transform(self):
        self.main_file.drop(self.main_file[self.main_file['MARCA'] == 'NO'].index, inplace=True)
        print("Removing NO MARCA1")
        self.main_file.drop(self.main_file[self.main_file['SUBMARCA'] == 'NO'].index, inplace=True)
        print("Removing NO MARCA2")

        self._accumulated = []
        self._empty = []
        self._written = []
        for lab in self.labs:
            self._written.append(self._replace_accord(lab))
        print("All files processed")
        # Empty books mean either no purchases this month or a name that doesn't
        # match any 'Laboratori' of the Mapa_Acords: worth checking, not silent.
        if self._empty:
            print("", flush=True)
            print(f"AVIS: {len(self._empty)} de {len(self.labs)} informes han quedat BUITS:",
                  flush=True)
            for lab in self._empty:
                print(f"  - {lab}", flush=True)
            print("Comprova que el nom coincideixi amb la columna 'Laboratori' del Mapa_Acords "
                  "i que el lab tingui compres aquest mes.", flush=True)
        # Categories with data that no generated lab claims -> nobody gets them.
        # Only meaningful on a full run: with one lab every other category would
        # be listed as an orphan.
        if getattr(self, "only_lab", None):
            return
        cats = (self.main_file['Laboratorio Categorizado'].dropna().astype(str)
                .str.strip().str.upper().value_counts())
        claimed = {l.strip().upper() for l in self.labs}
        orphans = [(c, k) for c, k in cats.items() if c not in claimed]
        if orphans:
            print("", flush=True)
            print(f"AVIS: {len(orphans)} categories amb dades que no van a cap informe "
                  f"(no son a la llista del SEGUIMENT):", flush=True)
            for c, k in orphans:
                print(f"  - {c}: {k} files", flush=True)

    # The file belonging to one lab for the processing month. Used when the fee
    # or refresh step runs on its own in single-lab mode, so it doesn't touch
    # the other labs' books. A single-lab split writes exactly this one file:
    # brands reclassified out of the queried lab already carry another
    # 'Laboratorio Categorizado' in the master and belong to that lab's own run.
    def lab_files(self, lab, temporality=""):
        self.tmp = " YTD" if temporality == "YTD" else ""
        want = f"PARAFARMACIA {self.this_month}{self.tmp} - {lab}.xlsx".upper()
        return [os.path.join(PASTE_DIR, f) for f in os.listdir(PASTE_DIR)
                if f.upper() == want]

    # Resolves the file set a step works on: an explicit list (single-lab mode)
    # or every workbook in PASTE_DIR (the global run).
    @staticmethod
    def _targets(only=None):
        if only is not None:
            return [os.path.abspath(p) for p in only]
        return [os.path.join(PASTE_DIR, f) for f in os.listdir(PASTE_DIR)
                if f.endswith((".xlsx", ".xlsm"))]

    # Filters the master data to one lab and appends the result as the
    # "Acuerdo book" sheet in that lab's output Excel file.
    # The match is on 'Laboratorio Categorizado', which index.py fills with the
    # Mapa_Acords 'Laboratori' label (addLab maps BIF -> Laboratori). The old
    # code instead looked the lab up in the Mapa 'BIF' column and matched
    # against THAT, so it only worked when Laboratori and BIF happen to be
    # spelled the same and silently wrote an EMPTY book otherwise -- e.g. FAES
    # (BIF 'FAES FARMA'), SANDOZ ('SANDOZ FARMACEUTICA S.A.'), VERMONT PHARMA.
    # Matching the categorized label directly returns the same rows for every
    # lab that already worked, and the missing ones for those that didn't.
    def _replace_accord(self, lab):
        cat = self.main_file['Laboratorio Categorizado'].astype(str).str.strip().str.upper()
        df_book = self.main_file[cat == lab.strip().upper()].drop_duplicates()
        # Printed per lab so an empty report is visible in the panel console
        # instead of only showing up as a 36 KB file nobody opens.
        print(f"  {lab}: {len(df_book)} files" + ("   <-- BUIT" if df_book.empty else ""))
        if df_book.empty:
            self._empty.append(lab)

        output_path = os.path.join(PASTE_DIR, f"PARAFARMACIA {self.this_month}{self.tmp} - {lab}.xlsx")
        with pd.ExcelWriter(output_path, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
            df_book.to_excel(writer, sheet_name="Acuerdo book", index=True, header=True)

        acc = df_book.copy()
        acc["lab"] = lab
        acc["month"] = self.this_month
        self._accumulated.append(acc)
        return output_path

    # Coerces a cell to float, treating None/NaN/blank as 0.
    def _num(self, x):
        return 0.0 if x is None or pd.isna(x) else float(x)

    # Reads each lab's fee conditions from the SEGUIMENT file.
    # Parafarmacia sheet: Etiquetas de fila | Neto | Fijo | Variable | fee mes
    # EFG sheet:          Laboratorio | Id Lab | Neto | Fijo | Variable
    # Returns {LAB_UPPER: (neto_bool, fijo, variable)}. Parafarmacia wins
    # over EFG for labs present in both (the split files are Parafarmacia data).
    def _read_fee_conditions(self):
        cond = {}

        def add(lab, neto, fijo, var):
            if lab is None or (isinstance(lab, float) and pd.isna(lab)):
                return
            key = str(lab).strip().upper()
            if key in ("", "NAN", "TOTAL GENERAL") or key in cond:
                return
            cond[key] = (str(neto).strip().upper().startswith("S"),
                         self._num(fijo), self._num(var))
            # Space-insensitive alias so "VermontPharma" matches "VERMONT PHARMA".
            cond.setdefault(key.replace(" ", ""), cond[key])

        pf = pd.read_excel(self.seguiment_file, sheet_name="Parafarmacia", header=1, usecols="A:E")
        pf.columns = ["lab", "neto", "fijo", "variable", "fee_mes"]
        for _, r in pf.iterrows():
            add(r["lab"], r["neto"], r["fijo"], r["variable"])

        efg = pd.read_excel(self.seguiment_file, sheet_name="EFG", header=1, usecols="A:E")
        efg.columns = ["lab", "idlab", "neto", "fijo", "variable"]
        for _, r in efg.iterrows():
            add(r["lab"], r["neto"], r["fijo"], r["variable"])
        return cond

    # Looks up a split lab's conditions, applying aliases and a
    # space-insensitive fallback. Returns None if not found.
    def _lookup_cond(self, cond, lab):
        key = str(lab).strip().upper()
        key = FEE_ALIASES.get(key, key)
        if key in cond:
            return cond[key]
        return cond.get(key.replace(" ", ""))

    # Builds the per-farmacia fee table for one lab's Acuerdo book df.
    # Base is Compra PUC when Neto=Sí, else Compra PVL. fee = Fijo*base,
    # variable = Variable*base (column shown only when Variable > 0),
    # total = fee + variable. Ends with a TOTAL row (flagged if no conditions).
    def _build_fee_table(self, df, cond):
        found = cond is not None
        neto, fijo, var = cond if found else (False, 0.0, 0.0)
        basecol = "Compra PUC" if neto else "Compra PVL"
        base_label = f"Suma de {basecol}"
        fee_label = f"Fee {fijo * 100:g}%"

        d = df.copy()
        d[basecol] = pd.to_numeric(d[basecol], errors="coerce").fillna(0)
        g = (d.groupby(["Nombre Oficina", "NIF"], dropna=False)[basecol]
             .sum().reset_index()
             .rename(columns={basecol: base_label})
             .sort_values(base_label, ascending=False))

        g[fee_label] = g[base_label] * fijo
        cols = ["Nombre Oficina", "NIF", base_label, fee_label]
        if var > 0:
            var_label = f"Variable {var * 100:g}%"
            g[var_label] = g[base_label] * var
            cols.append(var_label)
            g["Total"] = g[fee_label] + g[var_label]
        else:
            g["Total"] = g[fee_label]
        cols.append("Total")
        g = g[cols]

        totals = {c: (g[c].sum() if g[c].dtype.kind in "fi" else "") for c in cols}
        totals["Nombre Oficina"] = "TOTAL" if found else "TOTAL (CONDICIONS NO TROBADES)"
        totals["NIF"] = ""
        return pd.concat([g, pd.DataFrame([totals])[cols]], ignore_index=True)

    # For every output file in PASTE_DIR, computes the per-farmacia fee
    # from its Acuerdo book sheet and writes/replaces a "Fee" sheet.
    def add_fee_sheets(self, only=None):
        cond = self._read_fee_conditions()
        missing = []
        for path in self._targets(only):
            file_name = os.path.basename(path)
            lab = file_name.rsplit(" - ", 1)[-1].removesuffix(".xlsx")
            try:
                df = pd.read_excel(path, sheet_name="Acuerdo book")
            except Exception as e:
                print(f"  SKIP {file_name}: cannot read Acuerdo book ({e})")
                continue
            c = self._lookup_cond(cond, lab)
            if c is None:
                missing.append(lab)
            fee_df = self._build_fee_table(df, c)
            with pd.ExcelWriter(path, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
                fee_df.to_excel(writer, sheet_name="Fee", index=False)
            print(f"  {file_name} -> Fee ({len(fee_df) - 1} farmacies)"
                  + ("" if c else "  [NO CONDITIONS]"))
        if missing:
            print(f"\n{len(missing)} lab(s) had no conditions in SEGUIMENT: {missing}")

    # Refreshes the output workbooks via the isolated Excel-COM worker.
    # The refresh is the ONLY step tied to Windows+Excel: it lives in
    # com_worker.py (the "ancla") so the panel and the pipeline refresh
    # through exactly the same code path -- no duplicated COM logic to drift.
    # com_worker.refresh_files owns the single-instance reuse + respawn-on-crash
    # (RPC unavailable) and the "no files" message.
    def refresh_workbooks(self, only=None):
        import com_worker
        com_worker.refresh_files(self._targets(only))

    # Producer side of the decoupled refresh: moves the finished (split+fee)
    # workbooks out of PASTE_DIR into INBOX_DIR, where the standalone worker
    # service (com_worker.py --watch) picks them up, refreshes them and returns
    # them to PASTE_DIR (its --done). Use 'enqueue' INSTEAD of 'refresh' to run
    # the refresh asynchronously through the worker; 'refresh' still does it
    # in-process. Nothing downstream changes: files land back in PASTE_DIR.
    def enqueue_for_refresh(self, only=None):
        targets = self._targets(only)
        if not targets:
            print("No hi ha cap fitxer per encuar.")
            return
        os.makedirs(INBOX_DIR, exist_ok=True)
        for full_path in targets:
            file_name = os.path.basename(full_path)
            dest = os.path.join(INBOX_DIR, file_name)
            if os.path.exists(dest):
                os.remove(dest)  # re-run: reemplaca la versio anterior a la cua
            shutil.move(full_path, dest)
            print(f"  -> cua: {file_name}")
        print(f"{len(targets)} fitxer(s) encuats a {INBOX_DIR}. El worker els refrescara.")

    # Removes files from PASTE_DIR that haven't been modified
    # in the last 30 days to keep the output folder clean.
    def deleteOldFiles(self):
        for file_name in os.listdir(PASTE_DIR):
            file_path = os.path.join(PASTE_DIR, file_name)
            file_modified_time = os.path.getmtime(file_path)
            # If the file is older than 30 days, delete it
            if (datetime.datetime.now().timestamp() - file_modified_time) > (30 * 86400):
                os.remove(file_path)
                print(f"Deleted old file: {file_name}")
    # Appends each lab's per-month data into a single CSV file.
    # If data for the current month already exists, it is replaced.
    def noSQLAcumulateData(self):
        if not getattr(self, "_accumulated", None):
            print("No data accumulated -- run() may have failed. Skipping CSV update.")
            return
        new_data = pd.concat(self._accumulated, ignore_index=True)

        if os.path.exists(ACCUMULATE_CSV):
            existing = pd.read_csv(ACCUMULATE_CSV)
            print(existing.columns)
            existing = existing[existing["month"] != self.this_month]
            combined = pd.concat([existing, new_data], ignore_index=True)
        else:
            combined = new_data

        combined.to_csv(ACCUMULATE_CSV, index=False)
        print(f"Accumulated data saved to {ACCUMULATE_CSV}")

if __name__ == "__main__":
    import argparse
    # --steps selects which phases run (comma-separated). Every step is
    # independent: nothing here is ever chained from Pas 1 (index.py).
    #   split      -> run(): copy the template per lab + write its Acuerdo book
    #   fee        -> add_fee_sheets(): per-farmacia Fee sheet
    #   refresh    -> refresh_workbooks(): Excel COM refresh + hide Acuerdo book (in-process)
    #   enqueue    -> enqueue_for_refresh(): mou els fitxers a splitFiles/inbox
    #                 perque el worker (com_worker.py --watch) els refresqui async.
    #                 Alternativa a 'refresh': un dels dos, no els dos alhora.
    #   accumulate -> noSQLAcumulateData(): append to accumulated_labs.csv
    #   cleanup    -> deleteOldFiles(): remove pasteFiles older than 30 days
    # --laboratori restricts the run to one lab. It reads the SAME master of the
    # month from Z: (Parafarmacia {MM.AAAA}[ YTD].xlsx) as the full run and just
    # writes that lab's book, so it does NOT need Pas 1 to have been run for
    # that lab. fee/refresh then touch only that book; 'accumulate' and
    # 'cleanup' stay global.
    parser = argparse.ArgumentParser(description="Split del mestre en fitxers per laboratori.")
    parser.add_argument("--year", type=int, default=None, help="Any del mes a processar (opcional)")
    parser.add_argument("--month", type=int, default=None, help="Mes a processar 1-12 (opcional)")
    parser.add_argument("--period", default=PERIODO, choices=["", "YTD"], help='"" mensual, "YTD" acumulat')
    parser.add_argument("--steps", default="split,refresh",
                        help="Coma-separat: split, fee, refresh, enqueue, accumulate, cleanup")
    parser.add_argument("--laboratori", default=None,
                        help="Un sol laboratori: escriu nomes el seu llibre, del mateix mestre del mes")
    args = parser.parse_args()
    steps = [s.strip() for s in args.steps.split(",") if s.strip()]
    if not steps:
        print("Cap pas seleccionat (--steps buit): no hi ha res a fer.", flush=True)
        raise SystemExit(2)

    splitter = clsSplit(args.year, args.month)
    lab = (args.laboratori or "").strip()

    if lab:
        for glob_step in ("accumulate", "cleanup"):
            if glob_step in steps:
                print(f"El pas '{glob_step}' es global: s'ignora amb un laboratori seleccionat.",
                      flush=True)
        targets = None
        if "split" in steps:
            targets = [t for t in splitter.run(args.period, only_lab=lab) if t]
            # One lab was asked for and its book came out empty: stop with an
            # error instead of refreshing and handing over an empty file.
            if not targets or lab in getattr(splitter, "_empty", []):
                print(f"El llibre de '{lab}' ha quedat buit: no es refresca ni es dona per bo.",
                      flush=True)
                raise SystemExit(1)
        else:
            targets = splitter.lab_files(lab, args.period)
            if not targets:
                print(f"No trobo cap fitxer de split de '{lab}' del mes seleccionat a "
                      f"{PASTE_DIR}", flush=True)
                print("Marca tambe el pas 'Split' per generar-lo.", flush=True)
                raise SystemExit(1)
        if "fee" in steps:
            splitter.add_fee_sheets(only=targets)
        if "refresh" in steps:
            splitter.refresh_workbooks(only=targets)
        if "enqueue" in steps:
            splitter.enqueue_for_refresh(only=targets)
    else:
        if "split" in steps:
            splitter.run(args.period)
        if "fee" in steps:
            splitter.add_fee_sheets()
        if "refresh" in steps:
            splitter.refresh_workbooks()
        if "enqueue" in steps:
            splitter.enqueue_for_refresh()
        if "accumulate" in steps:
            splitter.noSQLAcumulateData()
        if "cleanup" in steps:
            splitter.deleteOldFiles()
