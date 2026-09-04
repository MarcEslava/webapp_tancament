import pandas as pd
import os
import re
import time
from datetime import date
from dateutil.relativedelta import relativedelta

'''
**************************************************************************************************************************************************************************************
The purpose of this program is the automatization of the process we use to clean the data extracted from BIFarma eco where we have all SO and SI from this year,
this procedure is done monthly so we can get it was necessary to make it more doable.


SI - Sell in (Compres)
SO - Sell out (Ventas)

Author: Marc Eslava
 
**************************************************************************************************************************************************************************************
'''

MAPA_ACORDS_PATH = r"Z:\Compres\INDÚSTRIA FARMACÈUTICA\Mapa_Acords.xlsx"

# Seconds to wait between month-chunk SQL queries so the server can serve
# other users between aggregations (see _read_bifarma_sql).
CHUNK_PAUSE_S = 2

# CARBO VITAL S.L. is a distributor, not a lab: one BIF carries several labs'
# brands, so the Mapa_Acords BIF -> Laboratori mapping can't resolve it (addLab
# keeps only the first row per BIF). Route by what the product description says.
# Add a line here when another of its brands gets its own lab; anything not
# matched stays uncategorized and shows up in the Pas 2 orphan warning.
# "bett?er": BIFarma writes BETER's makeup line as "BETTER" (two t), same EAN
# prefix 8412122 as the rest of BETER.
CARBO_VITAL_BRANDS = (
    ("BETER", "bett?er"),
    ("LUXEOL", "luxeol"),
    ("PURESSENTIEL", "puressentiel"),
)

# The report is about PARAFARMACIA, but a couple of especialitat slices have
# always belonged in it: the old manual BIFarma export added them by hand, which
# is what the checkAmox / checkAlmirall guards were watching for. They are part
# of the SQL now (see _esp_queries), so the Pas 2 split files carry them without
# anyone having to remember.
#   (Laboratori del Mapa_Acords, subgrup BIFarma, patro de producte | None = tot el subgrup)
# REIG JOFRE: nomes les amoxicilines, no tota la seva gamma EFG.
# ALMIRALL: tot el subgrup ESPECIALIDAD (ESPEC. CARAS queda fora, com al proces
# manual). Afegeix una linia aqui quan l'especialitat d'un altre laboratori hagi
# d'entrar; el lab s'anomena com a la columna 'Laboratori' del Mapa_Acords, i
# d'alli surten els seus noms BIF i Codi BIF -- mai es repeteixen aqui.
# NOT here on purpose: MENAVEN. checkMenaven's message says it is
# ESPECIALIDAD / Especialidad, and in BIFarma today it is not -- MENARINI
# (E0387) books it under PARAFARMACIA / CONSEJO FARMA, so the plain
# nombreGrupoProducto filter already brings it in (checked 08.2026: 617 rows,
# 2 refs, 11.695,81 EUR YTD). A rule for it would match 0 rows and warn on
# every run, and widening it to the whole MENARINI subgroup would drag 117
# unrelated especialitat refs (843.285 EUR) into a parafarmacia report.
ESPECIALITAT_INCLOSA = (
    ("REIG JOFRE", "EFG",          "AMOXI"),
    ("ALMIRALL",   "ESPECIALIDAD", None),
)

# "Parafarmacia 06.2026 YTD.xlsx" -> groups (month, year). Anchored, so Excel's
# temporary lock files ("~$Parafarmacia ...") and any other variant are ignored.
YTD_FILE_RE = re.compile(r"^Parafarmacia (\d{2})\.(\d{4}) YTD\.xlsx$", re.IGNORECASE)


class clsBiFarmaEco:
    # year/month: target processing month (e.g. 2026, 6). When omitted, defaults
    # to the last complete month (today - 1 month) so the monthly run needs no
    # arguments; a future frontend can pass them explicitly.
    # only_para: when True the SQL keeps PARAFARMACIA plus the especialitat
    # slices of ESPECIALITAT_INCLOSA (Almirall RX, amoxi Reig Jofre). When False
    # nothing is filtered and every product group comes in.
    # laboratori: when set (name from the Mapa_Acords 'Laboratori' column), the
    # SQL is filtered to that lab's BIF names/codes -> a one-lab report
    # (doesn't touch the global CSVs). It does NOT chain Pas 2: every step of
    # the pipeline runs on its own, so the split is asked for explicitly at
    # Pas 2 (clsSplit.py --laboratori) against the master written here.
    def __init__(self, rappel, period, year=None, month=None, only_para=False,
                 laboratori=None) -> None:
        self.rappel = rappel
        self.period = period
        self.only_para = only_para
        self.laboratori = (laboratori or "").strip() or None
        self._dir_ = os.getcwd()
        self._file_ = os.path.dirname(os.path.abspath(__file__))
        # proc_date = the month being processed. Single source of truth for every
        # date derived below (SQL ranges, file names, reference-list year).
        if year and month:
            self.proc_date = date(year, month, 1)
        else:
            self.proc_date = date.today() - relativedelta(months=1)
        self.proc_month = format(self.proc_date, "%m.%Y")  # e.g. "06.2026"
        self.proc_year = self.proc_date.year
        #absolute path to informes i rappel
        self.rappel_path = r"Z:\Compres\INDÚSTRIA FARMACÈUTICA\01. CARPETES LABORATORIS\102. Informes i Rappel"

# ----------FUNCTIONS------------- 
        # try:
        self.getFiles(period)
        # Legacy guards over the pull-everything run: they exit() when a slice
        # is missing, so they only make sense when nothing is filtered at all.
        # The PARAFARMACIA + ESPECIALITAT_INCLOSA path has its own, non-fatal
        # check instead (_report_especialitat).
        if not self.only_para and not self.laboratori:
            self.checkAmox()
            self.checkAlmirall()
            self.checkMenaven()
        self.nompikis()
        self.superestalvi()
        self.DataFilters() #dataFilters
        self.noComunesRename()
        self.cleanData()
        self.collateData()
        self.transformNumValue()
        self.addFormulas()
        self.unifyCamps()  # Farmacia change/unify names
        self.addLab()
        self.splitSensilir()
        self.splitCarboVital()
        self.renameColumns()
        self.detectDuplicateValues()
        self.unionSisfarma()
        # self.newLabs()
        self.toExcel(period)
        # Single-lab runs are ad-hoc reports: don't overwrite the global CSVs.
        if not self.laboratori:
            self._upsert_conditions_csv()  # always keep the product-conditions CSV updated
            self._lost_money_report()      # always refresh the per-lab lost-money report
            self._desa_a_clickhouse(period)  # ClickHouse (tancament): capa analítica única
        else:
            print(f"Informe d'un sol laboratori ({self.laboratori}): no s'actualitzen els CSV globals.", flush=True)
            print(f"Per obtenir-ne el fitxer de split, executa el Pas 2 amb el laboratori "
                  f"'{self.laboratori}' seleccionat.", flush=True)
        # except Exception as e:
        #     print(f"Error during processing: {e}")

    # Builds a SQLConnection from th1e local (git-ignored) db_config.py.
    # Lazy imports so the module still loads if db_config.py is absent.
    # connect_retries: a brief "Adaptive Server is unavailable" (connection
    # refused while the server is momentarily saturated -- e.g. a blocking
    # chain elsewhere exhausting its worker threads) shouldn't sink a whole
    # multi-minute run; retry a few times before giving up.
    def _db(self):
        import db_config as cfg
        from SQLConnection import SQLConnection
        return SQLConnection(
            db_host=cfg.DB_HOST, db_port=getattr(cfg, "DB_PORT", 1433),
            db_database=cfg.DB_NAME, db_username=cfg.DB_USER,
            db_password=cfg.DB_PASS, dialect="mssql", driver="pymssql",
            login_timeout=15,  # fail fast per attempt (don't hang)
            connect_retries=3, retry_delay=15,
        )

    # Builds the SQL condition + bind params that restrict the query to one
    # categorized lab from Mapa_Acords: all its BIF names (pr.deslab LIKE, the
    # same "contains" match DataFilters uses) and Codi BIF codes (pr.codlab).
    def _lab_sql_filter(self):
        cond, params = self._lab_cond(self.laboratori, "lab")
        if cond is None:
            print(f"ERROR: laboratori '{self.laboratori}' no trobat a la columna 'Laboratori' del Mapa_Acords.",
                  flush=True)
            raise SystemExit(2)
        names, codes = self._lab_names_codes(self.laboratori)
        print(f"Filtre de laboratori '{self.laboratori}': {len(names)} noms BIF, {len(codes)} codis.",
              flush=True)
        return cond, params

    # Mapa_Acords lives on Z: and several steps need it (the lab filter, the
    # especialitat rules, df_acords). Read it once per run, not once per caller.
    def _mapa(self):
        if getattr(self, "_mapa_cache", None) is None:
            self._mapa_cache = pd.read_excel(MAPA_ACORDS_PATH, sheet_name="Mapa")
        return self._mapa_cache

    # The BIF names and Codi BIF codes Mapa_Acords lists for one 'Laboratori'.
    # Both lists come back empty when the lab isn't in the Mapa -- the caller
    # decides whether that is fatal.
    def _lab_names_codes(self, laboratori):
        mapa = self._mapa()
        sel = mapa[mapa['Laboratori'].astype(str).str.strip().str.upper()
                   == str(laboratori).strip().upper()]

        # Excel numeric cells arrive as floats (123.0): normalize to "123".
        def clean(v):
            if isinstance(v, float) and v.is_integer():
                return str(int(v))
            return str(v).strip()

        names = [clean(v) for v in sel['BIF'].dropna() if clean(v)]
        codes = [clean(v) for v in sel['Codi BIF'].dropna() if clean(v)]
        return names, codes

    # SQL condition matching one lab, the same hybrid way DataFilters does --
    # BIF name (pr.deslab LIKE) or Codi BIF (pr.codlab) -- with every value bound
    # as a driver parameter, never interpolated. 'prefix' keeps the bind names
    # unique so several labs can share one query. (None, {}) when the lab has
    # neither a BIF name nor a code in the Mapa.
    def _lab_cond(self, laboratori, prefix):
        names, codes = self._lab_names_codes(laboratori)
        conds, params = [], {}
        for i, n in enumerate(names):
            conds.append(f"pr.deslab LIKE %({prefix}_nom{i})s")
            params[f"{prefix}_nom{i}"] = f"%{n}%"
        if codes:
            placeholders = ", ".join(f"%({prefix}_codi{i})s" for i in range(len(codes)))
            conds.append(f"pr.codlab IN ({placeholders})")
            params.update({f"{prefix}_codi{i}": c for i, c in enumerate(codes)})
        if not conds:
            return None, {}
        return "(" + " OR ".join(conds) + ")", params

    # One narrow query per ESPECIALITAT_INCLOSA rule -- (etiqueta, WHERE, binds)
    # -- instead of OR branches bolted onto the PARAFARMACIA scan. Grup and
    # subgrup are attributes of the product's family (vteco_familias joined by
    # pr.idfamilia), so the row sets are disjoint and their union is exactly what
    # the old OR selected. What changes is the plan: each slice is filtered on its
    # own narrow predicate instead of making the server evaluate every
    # 'deslab LIKE %...%' (leading wildcard, no index seek) against the whole
    # month partition of the PARAFARMACIA scan.
    # A rule whose lab is missing from the Mapa_Acords is skipped with a warning
    # -- the rest of the report is still valid, so it must not kill the run.
    # In a one-lab run the other labs' rules are dropped here: the LAB filter
    # would return 0 rows for them anyway, and that is one round trip per month
    # per year for nothing.
    def _esp_queries(self):
        out = []
        for i, (lab, subgrup, patro) in enumerate(ESPECIALITAT_INCLOSA):
            if self.laboratori and self.laboratori.strip().upper() != lab.strip().upper():
                continue
            cond, params = self._lab_cond(lab, f"esp{i}")
            if cond is None:
                print(f"Avis: '{lab}' no es al Mapa_Acords; la seva especialitat "
                      f"'{subgrup}' no entrara a l'informe.", flush=True)
                continue
            trossos = [cond, f"f.nombresubgrupoproducto = %(esp{i}_sub)s"]
            params[f"esp{i}_sub"] = subgrup
            if patro:
                trossos.append(f"pr.desproducto LIKE %(esp{i}_prod)s")
                params[f"esp{i}_prod"] = f"%{patro}%"
            detall = f"productes amb '{patro}'" if patro else "tot el subgrup"
            print(f"Especialitat inclosa: {lab} / {subgrup} ({detall}).", flush=True)
            etiqueta = f"{lab} / {subgrup}" + (f" / '{patro}'" if patro else "")
            out.append((etiqueta, "AND " + " AND ".join(trossos), params))
        return out

    # Non-blocking check of what the especialitat queries actually returned. The
    # SQL can't forget a slice the way the manual export could, but a rule can
    # quietly stop matching (a product renamed in BIFarma, a lab that left the
    # Mapa), and that has to be visible in the log. The counts are the rows each
    # query returned -- there is no second copy of the rule in pandas that could
    # drift out of step with the SQL.
    def _report_especialitat(self, stats):
        for etiqueta, (n_act, n_ant, refs) in stats.items():
            if n_act or n_ant:
                print(f"  Especialitat {etiqueta}: {n_act} files {self.proc_year}, "
                      f"{n_ant} files {self.proc_year - 1}, {len(refs)} referencies.",
                      flush=True)
            else:
                print(f"  AVIS: especialitat {etiqueta}: 0 files als dos anys. Comprova "
                      f"el subgrup a BIFarma o el nom del producte.", flush=True)

    # Queries the BIFarma SQL DB and returns df_bifarma with the legacy Excel
    # column names. period "" -> current month only; "YTD" -> Jan..current month.
    # 'act' = current year, 'ant' = same range one year earlier.
    def _read_bifarma_sql(self, period):
        d = self.proc_date
        curr_yy = d.year
        prev_yy = d.year - 1
        fin_month = str(d.month).zfill(2)

        GROUP_BY = """
            GROUP BY
                pr.codproducto, pr.desproducto, pr.codlab, pr.deslab,
                de.identidad, de.iddelegacion, de.delegacion,
                pr.idproducto, f.nombresubgrupoproducto,
                pr.idsuperfamilia, f.nombresuperfamiliaeco,
                pr.idfamilia, f.nombrefamiliaeco"""

        # WITH (NOLOCK): read without taking shared locks so this reporting
        # query never blocks other users' writes/reads. Dirty reads are fine
        # here -- we aggregate closed (past) months that no longer change.
        BASE_FROM = """
            FROM dbo.bench_dwComprasVentasMesS T1 WITH (NOLOCK)
            INNER JOIN dbo.tme_delegaciones de WITH (NOLOCK) ON T1.idendeS = de.idendeS
            INNER JOIN dbo.tbi_productosERS pr WITH (NOLOCK) ON T1.idendeS = pr.idendeS AND T1.idproducto = pr.idproducto
            INNER JOIN dbo.vteco_familias f    WITH (NOLOCK) ON f.idfamiliaeco = pr.idfamilia"""

        BASE_COLS = """
                pr.codproducto AS CodProducto,
                pr.desproducto AS Producto,
                pr.codlab AS IdLaboratorio,
                pr.deslab AS Laboratorio,
                de.identidad AS IdEntidad,
                de.iddelegacion AS IdDelegacion,
                de.delegacion AS Delegacion,
                pr.idproducto AS IdProducto,
                f.nombresubgrupoproducto AS SubGrupoProducto,
                pr.idsuperfamilia AS IdSuperFamilia,
                f.nombresuperfamiliaeco AS SuperFamilia,
                pr.idfamilia AS IdFamilia,
                f.nombrefamiliaeco AS Familia"""

        ECO_FILTER = "(T1.idendes IN (SELECT idendes FROM tme_delegaciones WITH (NOLOCK) WHERE grupoCompras = 'ECO'))"
        # only_para: the main query keeps PARAFARMACIA and nothing else; the
        # especialitat slices of ESPECIALITAT_INCLOSA (Almirall RX, amoxi Reig
        # Jofre) come from their own queries and are stacked on top in
        # fetch_side. Without only_para nothing is filtered here and every
        # product group comes in, all of especialitat included.
        PARA = "AND f.nombreGrupoProducto = 'PARAFARMACIA'" if self.only_para else ""
        ESP = self._esp_queries() if self.only_para else []
        # Optional: restrict to a single lab (Mapa_Acords 'Laboratori'). Same
        # hybrid criterion as DataFilters -- BIF name (LIKE) or Codi BIF -- with
        # every value bound as a driver parameter, never interpolated. ANDed on
        # top of every query, the especialitat ones included.
        LAB, lab_params = "", {}
        if self.laboratori:
            lab_cond, lab_params = self._lab_sql_filter()
            LAB = f"AND {lab_cond}"

        # 'act' carries the current stock (Estoc); 'ant' does not.
        act_cols = f"""{BASE_COLS},
                    MIN(pr.stockActual) AS Estoc,
                    SUM(ISNULL(T1.cantidad, 0))       AS CantidadAct,
                    SUM(ISNULL(T1.importe, 0))        AS ImporteAct,
                    SUM(ISNULL(T1.cantidadcompra, 0)) AS CantidadCompraAct,
                    SUM(ISNULL(T1.importecompra, 0))  AS ImporteCompraAct"""
        ant_cols = f"""{BASE_COLS},
                    MIN(pr.stockActual) AS Estoc,
                    SUM(ISNULL(T1.cantidad, 0))       AS CantidadAnt,
                    SUM(ISNULL(T1.importe, 0))        AS ImporteAnt,
                    SUM(ISNULL(T1.cantidadcompra, 0)) AS CantidadCompraAnt,
                    SUM(ISNULL(T1.importecompra, 0))  AS ImporteCompraAnt"""

        # Monthly run = one query; YTD = ONE QUERY PER MONTH (Jan..current) with
        # a pause between them. A whole-YTD aggregation in a single query
        # saturated the server for minutes; per-month chunks keep each
        # aggregation small and the monthly partials are recombined in pandas
        # (same result as the big GROUP BY).
        if period == "YTD":
            act_months = [f"{curr_yy}{m:02d}" for m in range(1, d.month + 1)]
            ant_months = [f"{prev_yy}{m:02d}" for m in range(1, d.month + 1)]
        else:
            act_months = [f"{curr_yy}{fin_month}"]
            ant_months = [f"{prev_yy}{fin_month}"]

        # The GROUP BY (dimension) columns, as returned by the SELECT aliases.
        DIM_KEYS = ['CodProducto', 'Producto', 'IdLaboratorio', 'Laboratorio',
                    'IdEntidad', 'IdDelegacion', 'Delegacion', 'IdProducto',
                    'SubGrupoProducto', 'IdSuperFamilia', 'SuperFamilia',
                    'IdFamilia', 'Familia']

        # The PARAFARMACIA side: ONE QUERY PER MONTH, because that aggregation
        # over a whole YTD is what saturated the server. lab_params carries the
        # binds of the optional one-lab filter; no binds at all -> None, the
        # plain (unparameterized) driver path.
        def fetch_months(db, cols, months, agg):
            params = lab_params or None
            parts = []
            for i, ym in enumerate(months):
                print(f"    mes {ym} ({i + 1}/{len(months)})...", flush=True)
                parts.append(db.fech_dataframe(
                    f"SELECT {cols} {BASE_FROM} WHERE T1.anyomes = {ym} AND {ECO_FILTER} {PARA} {LAB} {GROUP_BY}",
                    params))
                if i < len(months) - 1:
                    time.sleep(CHUNK_PAUSE_S)  # let the server breathe between chunks
            df = pd.concat(parts, ignore_index=True)
            if len(parts) > 1:
                # Recombine monthly aggregates: sums add up; Estoc (current
                # stock) is identical every month, so 'min' just picks it.
                df = df.groupby(DIM_KEYS, dropna=False, as_index=False).agg(agg)
            return df

        # One especialitat slice, the whole period in a SINGLE query -- no
        # chunking. The chunking above is there for the size of the PARAFARMACIA
        # aggregation; a slice is one lab and one subgroup, a few hundred rows,
        # so there is nothing to break up. GROUP_BY has no anyomes column, so
        # the server already sums the months and the result needs no
        # recombining in pandas. `where` is the rule's condition and `binds` its
        # parameters, with the lab filter's binds merged in.
        def fetch_range(db, cols, months, where, binds):
            params = {**binds, **lab_params} or None
            return db.fech_dataframe(
                f"SELECT {cols} {BASE_FROM} WHERE T1.anyomes IN ({', '.join(months)}) "
                f"AND {ECO_FILTER} {where} {LAB} {GROUP_BY}",
                params)

        # One side of the comparative (act or ant): the PARAFARMACIA query plus
        # one query per especialitat rule, stacked. `slot` says which side's
        # counter to fill in esp_stats (0 = act, 1 = ant).
        # Today's rules select disjoint rows, so the final groupby changes
        # nothing; it is there so that two rules which ever do overlap get their
        # figures summed once instead of the row appearing twice.
        def fetch_side(db, cols, months, agg, slot):
            df = fetch_months(db, cols, months, agg)
            if ESP:
                print(f"    -> {len(df)} files de PARAFARMACIA.", flush=True)
            stacked = False
            for etiqueta, where, binds in ESP:
                print(f"    especialitat {etiqueta}...", flush=True)
                esp = fetch_range(db, cols, months, where, binds)
                st = esp_stats.setdefault(etiqueta, [0, 0, set()])
                st[slot] = len(esp)
                st[2].update(esp['CodProducto'].astype(str))
                if not esp.empty:
                    df = pd.concat([df, esp], ignore_index=True)
                    stacked = True
            if stacked:
                df = df.groupby(DIM_KEYS, dropna=False, as_index=False).agg(agg)
            return df

        esp_stats = {}

        try:
            with self._db() as db:
                print(f"  SQL any actual {curr_yy} ({len(act_months)} mes(os))...", flush=True)
                act_df = fetch_side(db, act_cols, act_months,
                                    {'Estoc': 'min', 'CantidadAct': 'sum', 'ImporteAct': 'sum',
                                     'CantidadCompraAct': 'sum', 'ImporteCompraAct': 'sum'}, 0)
                print(f"  -> {len(act_df)} files rebudes. SQL any anterior {prev_yy}...", flush=True)
                ant_df = fetch_side(db, ant_cols, ant_months,
                                    {'Estoc': 'min', 'CantidadAnt': 'sum', 'ImporteAnt': 'sum',
                                     'CantidadCompraAnt': 'sum', 'ImporteCompraAnt': 'sum'}, 1)
                print(f"  -> {len(ant_df)} files rebudes. Processant dades...", flush=True)
        except Exception as e:
            # Clean message instead of a huge SQLAlchemy/pymssql traceback when the
            # DB is down/unreachable or drops the connection mid-transfer (10054).
            import db_config as cfg
            print("\n" + "=" * 72, flush=True)
            print("✗ No s'ha pogut llegir de la base de dades.", flush=True)
            print(f"  Servidor: {getattr(cfg, 'DB_HOST', '?')}", flush=True)
            print("  Comprova que el servidor SQL està en marxa i accessible", flush=True)
            print("  (botó 'Comprovar servidor i Z:' del panell) i torna-ho a provar.", flush=True)
            print(f"  Detall tècnic: {type(e).__name__}: {str(e).splitlines()[0][:160]}", flush=True)
            print("=" * 72, flush=True)
            raise SystemExit(2)

        MERGE_KEYS = ['CodProducto', 'IdLaboratorio', 'IdEntidad', 'IdDelegacion', 'IdProducto']
        # OUTER, not left: a product/pharmacy pair with movement LAST year and none
        # this year has no 'act' row, so a left join dropped it -- and with it its
        # whole prior-year figure. Measured on 07: 49% of the 'ant' rows had no
        # 'act' twin, worth ~26% of the units and ~32% of the euros of the 'Ant'
        # columns, which is exactly why the comparatives came out well below the
        # legacy BIFarma Excel export.
        products_df = pd.merge(act_df, ant_df, on=MERGE_KEYS, how='outer', suffixes=('', '_ant'))
        # Descriptive columns (Producto, Laboratorio, Delegacion, families, Estoc)
        # come from both sides: on an 'ant only' row the 'act' side is NaN, so fill
        # it from its '_ant' twin before that twin gets dropped below.
        for col in [c for c in products_df.columns if c.endswith('_ant')]:
            base = col[:-len('_ant')]
            if base in products_df.columns:
                products_df[base] = products_df[base].combine_first(products_df[col])
        # A row present on only one side has no figures on the other -> 0
        # ('act' N/D for products dropped since last year, 'ant' N/D for new ones)
        for col in ['CantidadAct', 'ImporteAct', 'CantidadCompraAct', 'ImporteCompraAct',
                    'CantidadAnt', 'ImporteAnt', 'CantidadCompraAnt', 'ImporteCompraAnt',
                    'Estoc']:
            if col in products_df.columns:
                products_df[col] = products_df[col].fillna(0)
        # Drop the duplicate descriptive columns of the 'ant' side plus the unused IdProducto
        drop_extra = [c for c in products_df.columns if c.endswith('_ant')] + ['IdProducto']
        products_df = products_df.drop(columns=drop_extra, errors='ignore')
        # Rename SQL columns to match the legacy Excel column names used downstream
        products_df.rename(columns={
            'CodProducto':       'Cod Unif',
            'IdLaboratorio':     'Id Lab',
            'IdEntidad':         'Farm',
            'IdDelegacion':      'Of',
            'Delegacion':        'Nombre Oficina',
            'CantidadAct':       'Venta (Ud)\nAct',
            'ImporteAct':        'Venta (€)\nAct',
            'CantidadCompraAct': 'Compra (Ud)\nAct',
            'ImporteCompraAct':  'Compra (€)\nAct',
            'CantidadAnt':       'Venta (Ud)\nAnt',
            'ImporteAnt':        'Venta (€)\nAnt',
            'CantidadCompraAnt': 'Compra (Ud)\nAnt',
            'ImporteCompraAnt':  'Compra (€)\nAnt',
            'Estoc':             'Stock\nactual',
            'SubGrupoProducto':  'SubGrupo',
            'IdSuperFamilia':    'Id\nSupFam',
            'IdFamilia':         'Id\nFamilia',
        }, inplace=True)
        products_df['Cod Nac'] = products_df['Cod Unif']
        if self.only_para:
            self._report_especialitat(esp_stats)
        return products_df

    # Get the diferent files we need to run the program
    def getFiles(self, period):
        print(f"Processant {self.proc_month} (període '{period or 'mensual'}')", flush=True)
        # df_bifarma now comes from the BIFarma SQL DB instead of the Excel exports
        self.df_bifarma = self._read_bifarma_sql(period)
        print("Llegint fitxers de suport de Z: (Acords, Books, YTD)...", flush=True)
        # Acords file to check labs we have a deal with
        self.df_acords = self._mapa().copy()  # cleanData drops rows -> never share the cache
        #  We get 'NIF', 'BOOK' and 'SOLAR' from this file
        self.df_books = pd.read_excel(r"Z:\Compres\INDÚSTRIA FARMACÈUTICA\BBDD_Book INTERN.xlsx", sheet_name=self.rappel)
        
        # Reference list year taken from the processing date. Only feeds
        # newLabs() (disabled), so a missing file must not break the pipeline.
        master_path = (
            r"Z:\Compres\INDÚSTRIA FARMACÈUTICA\01. CARPETES LABORATORIS"
            r"\101. Llistat referencies per laboratori i books"
            rf"\Listado referencias por labo y cluster - {self.proc_year} YTD.xlsx"
        )
        try:
            self.df_products_master = pd.read_excel(master_path, sheet_name="Listado Acuerdos")
        except FileNotFoundError:
            print(f"Avís: no s'ha trobat el llistat de referències ({master_path}). "
                  "df_products_master queda buit (només afecta newLabs, desactivat).")
            self.df_products_master = pd.DataFrame()
        # We get 'PVL', 'IVA' and 'MARCA' from a YTD master (always YTD: it's the
        # richest source of product prices/brands).
        ytd_path = self._find_ytd_source()
        self.df_bifarma_lastMonth = pd.read_excel(ytd_path, sheet_name="SI Acord book")

    # Picks the Parafarmacia YTD master used as the PVL/IVA/MARCA/SUBMARCA source.
    # Preference order:
    #   1. The processing month's own YTD, when it already exists on Z: -- the
    #      monthly run needs it (it carries this month's manual MARCA work).
    #   2. Otherwise the MOST RECENT EARLIER YTD (normally the previous month).
    #      This is the case when generating the YTD itself: 07.2026 YTD can't
    #      read 07.2026 YTD (it's the file being produced), so it falls back to
    #      06.2026 YTD -- or to whatever the latest available month is.
    # Months come from the file name, not the modification date, so a file
    # re-saved later doesn't jump the queue.
    def _find_ytd_source(self):
        same = os.path.join(self.rappel_path, f"Parafarmacia {self.proc_month} YTD.xlsx")
        if os.path.isfile(same):
            print(f"  Font PVL/IVA/MARCA: {os.path.basename(same)}", flush=True)
            return same

        # The master of the month in progress sits in the rappel root; once the
        # month is closed it gets archived into its year folder
        # ("...\2026\Parafarmacia 06.2026 YTD.xlsx"), so scan both. Only one
        # level deep: month folders and "copia seguretat" hold per-lab splits
        # and dated backups, not masters.
        try:
            folders = [self.rappel_path] + [
                os.path.join(self.rappel_path, n) for n in os.listdir(self.rappel_path)
                if n.isdigit() and os.path.isdir(os.path.join(self.rappel_path, n))
            ]
        except OSError as e:
            print(f"ERROR: no s'ha pogut llegir la carpeta de rappel ({self.rappel_path}): {e}", flush=True)
            raise SystemExit(2)

        candidates = []
        for folder in folders:
            try:
                names = os.listdir(folder)
            except OSError:
                continue
            for name in names:
                m = YTD_FILE_RE.match(name)
                if not m:
                    continue
                file_date = date(int(m.group(2)), int(m.group(1)), 1)
                if file_date < self.proc_date:  # strictly earlier: never itself, never the future
                    candidates.append((file_date, os.path.join(folder, name)))

        if not candidates:
            print("\n" + "=" * 72, flush=True)
            print(f"ERROR: no s'ha trobat cap 'Parafarmacia MM.AAAA YTD.xlsx' anterior a {self.proc_month}.", flush=True)
            print(f"  Carpeta: {self.rappel_path}", flush=True)
            print("  Cal el YTD d'un mes anterior per obtenir PVL / IVA / MARCA / SUBMARCA.", flush=True)
            print("=" * 72, flush=True)
            raise SystemExit(2)

        file_date, path = max(candidates)
        print(f"  '{os.path.basename(same)}' no existeix encara -> "
              f"font PVL/IVA/MARCA: {os.path.basename(path)}", flush=True)
        return path
        
    # Creates de filtered file from the original bifarmaeco file we download, the result has the columns we need from the orginal file with only labs we have an acord with all cleaned. 
    def cleanData(self):
        # if self.filt_codUnif.any() != 0:
        #     self.df_bifarma = self.df_bifarma.drop(index = self.df_bifarma[self.filt_codUnif].index)
        self.df_bifarma =  self.df_bifarma.dropna(axis='index', subset=['Cod Unif'])
        # errors='ignore': the SQL source lacks the Excel-only columns (Var, PVP, blanks...)
        self.df_bifarma = self.df_bifarma.drop(columns=['Var', '% Var','Var.2','% Var.2', 'PVP','Pr. Medio\nAnt', 'Pr. Medio\nAct', 'Var.1', '% Var.1', 'Var.3', '% Var.3', 'SubGrupo', 'Id\nSupFam', 'SuperFamilia', 'Id\nFamilia', 'Familia',' ',' .1',' .2',' .3',' .4',' .5',' .6',' .7'], errors='ignore')
        self.df_acords = self.df_acords.drop(index = self.df_acords[self.filt_isAcord].index)
        self.df_bifarma_filtered = self.df_bifarma[self.filt_accords]
        
    #  Creates the diferent filters we use when cleaning data i/o treating it.
    def DataFilters(self):
        if self.df_bifarma['Cod Unif'].astype(str).str.contains('NOCOMUNES', na=False).any():
            self.filt_codUnif = self.df_bifarma['Cod Unif'].astype(str).str.contains('NOCOMUNES', na=False)
        else:
            # No NOCOMUNES rows (e.g. SQL source): empty False mask so noComunesRename is a no-op
            self.filt_codUnif = pd.Series(False, index=self.df_bifarma.index)
        self.filt_isAcord = self.df_acords['Acord'].str.contains('NO', na=False)
        # Match per nom (llegat): manté les files on 'Laboratorio' conté algun BIF de l'acord
        list_acords = self.df_acords['BIF'].tolist()
        filt_nom = self.df_bifarma['Laboratorio'].apply(
            lambda x: any(acord in x for acord in list_acords) if isinstance(x, str) else False
        )
        # Match per codi: manté les files on 'Id Lab' coincideix amb un 'Codi BIF' de l'acord
        codis_bif = set(self.df_acords['Codi BIF'].dropna().astype(str).str.strip())
        filt_codi = self.df_bifarma['Id Lab'].astype(str).str.strip().isin(codis_bif)
        # Híbrid: es queda la fila si coincideix per codi O per nom
        self.filt_accords = filt_codi | filt_nom

    #  We add the columns 'BIFARMA',"NIF","BOOK" and "SOLAR".
    def collateData(self):
        list_books = self.df_books[['BIFARMA',"NIF","BOOK","SOLAR"]]
        self.df_bifarma_books = pd.merge(self.df_bifarma_filtered, list_books,left_on='Nombre Oficina', right_on='BIFARMA', how="left")
        self.df_bifarma_books = self.df_bifarma_books.dropna(axis=0, subset=['BOOK'])
        self.df_bifarma_books = self.df_bifarma_books.drop(columns=['BIFARMA'])
        # print(self.df_bifarma_books['Nombre Oficina'].unique())

    #  SO and SI values come in object format so we transform them to float values to treat them properly.
    def transformNumValue(self):
        columns = ['Venta (Ud)\nAct', 'Venta (Ud)\nAnt',   #this are the columns we need in numeric value
       'Venta (€)\nAct', 'Venta (€)\nAnt', 'Compra (Ud)\nAct',
       'Compra (Ud)\nAnt', 'Compra (€)\nAct', 'Compra (€)\nAnt','Stock\nactual']
        self.df_bifarma_books.loc[:, columns] = self.df_bifarma_books[columns].astype(str).astype(float)
        numeric_columns = self.df_bifarma_books.select_dtypes(include=['number'])
        numeric_columns = numeric_columns.clip(lower=0)
        self.df_bifarma_books.loc[:, numeric_columns.columns] = numeric_columns
        
        
        # Font de PVL/IVA/MARCA/SUBMARCA: ClickHouse dim_producte, el store
        # gobernat i editable des de l'editor web (ReplacingMergeTree: l'última
        # edició mana). Substitueix la lectura del YTD Excel. Fail-soft: si
        # ClickHouse no respon o està buit, es cau al YTD Excel com a xarxa de
        # seguretat, de manera que una incidència de ClickHouse mai trenca un tancament.
        try:
            import store_ch
            self.list_anterior = store_ch.read_producte_attr()
            if self.list_anterior.empty:
                raise RuntimeError("dim_producte buit")
            print(f"PVL/IVA/MARCA/SUBMARCA: {len(self.list_anterior)} productes des de "
                  f"ClickHouse (dim_producte).", flush=True)
        except Exception as e:
            print(f"Avís: no s'ha pogut llegir dim_producte de ClickHouse "
                  f"({type(e).__name__}: {e}). S'usa el YTD Excel com a font de "
                  f"PVL/IVA/MARCA/SUBMARCA.", flush=True)
            self.list_anterior = self.df_bifarma_lastMonth[["Cod Unif", 'PVL', "IVA", "MARCA", "SUBMARCA"]].copy()
        self.list_anterior.drop_duplicates(inplace=True)
        # ---- We need to transform "cod unif" to a numeric value to truncate them ----
        # Non-numeric codes (e.g. 'NOCOMUNES-ESP') become NaN instead of raising.
        self.df_bifarma_books.loc[:,'Cod Unif'] = pd.to_numeric(
            self.df_bifarma_books['Cod Unif'].astype(str).str.strip(), errors='coerce')
        self.list_anterior.loc[:,'Cod Unif'] = pd.to_numeric(
            self.list_anterior['Cod Unif'].astype(str).str.strip(), errors='coerce')
        self.df_bifarma_final = pd.merge(self.df_bifarma_books, self.list_anterior,on="Cod Unif", how="left")
        self.df_bifarma_final = self.df_bifarma_final.drop_duplicates()
        # New products (not in last month's file) have no MARCA -> flag them as new
        self.df_bifarma_final['MARCA'] = self.df_bifarma_final['MARCA'].fillna("New!")
        
    #  we add the formulas for "Compra PUC" and the formula for "Compra PVL"
        #  PUC - The price of 1 product "neta"
        #  PVL - The selling price of the lab of 1 unit. "bruta"
    def addFormulas(self):
        # IVA as a real float (like PVL) so it stores as 0.1 -> shown as 0,1 in the locale, not text
        self.df_bifarma_final['IVA'] = pd.to_numeric(
            self.df_bifarma_final['IVA'].astype(str).str.strip().str.replace('#N/D', '0').str.replace(',', '.'),
            errors='coerce')
        # PVL and SI(Ud)Act as real floats so the cells the Excel formulas reference are numeric.
        # Compra PUC / Compra PVL are written as Excel formulas in toExcel (not precalculated here).
        self.df_bifarma_final['PVL'] = pd.to_numeric(
            self.df_bifarma_final['PVL'].astype(str).str.strip().str.replace('#N/D', '0').str.replace(',', '.'),
            errors='coerce')
        self.df_bifarma_final['Compra (Ud)\nAct'] = pd.to_numeric(
            self.df_bifarma_final['Compra (Ud)\nAct'].astype(str).str.strip().str.replace(',', '.'),
            errors='coerce')
        self.df_bifarma_final['Compra (Ud)\nAct'] = self.df_bifarma_final['Compra (Ud)\nAct'].astype(float)
        
    # We rename the columns so they are more acurate to their real function, on the future we should change them to a more generic name so we can use the pivot tables.
    def renameColumns(self):
        self.df_bifarma_final.rename(columns={'Venta (Ud)\nAct':'SO (Ud)\nAct','Venta (Ud)\nAnt':'SO (Ud)\nAnt','Venta (€)\nAct':'SO (€)\nAct','Venta (€)\nAnt':'SO (€)\nAnt','Compra (Ud)\nAct':'SI (Ud)\nAct','Compra (Ud)\nAnt':'SI (Ud)\nAnt','Compra (€)\nAct':'SI (€)\nAct','Compra (€)\nAnt':'SI (€)\nAnt'}, inplace=True)
    
    def unifyCamps(self):
        # Viñamata has 2 office we unify them on the same name using the NIF
        self.df_bifarma_final.loc[self.df_bifarma_final['NIF'] == '38793542V', 'Nombre Oficina'] = 'FARMACIA VIÑAMATA'
        self.df_bifarma_final.loc[self.df_bifarma_final['NIF'] == '52150363W', 'Nombre Oficina'] = 'FARMACIA VILA'
        
    #  Creates the resulting excel file we will use to create the reports for the labs
    def toExcel(self, period):
        from xlsxwriter.utility import xl_col_to_name
        df = self.df_bifarma_final
        # Formula columns must exist so they land in the canonical order below
        for c in ['Compra PUC', 'Compra PVL']:
            if c not in df.columns:
                df[c] = None
        # Reorder to the canonical structure (see output/df_bifarma_output_ytd_426.xlsx).
        # Missing columns are skipped; any unexpected leftover is appended at the end.
        column_order = [
            'Cod Nac', 'Cod Unif', 'Producto', 'Id Lab', 'Laboratorio', 'Farm', 'Of', 'Nombre Oficina',
            'SO (Ud)\nAct', 'SO (Ud)\nAnt', 'SO (€)\nAct', 'SO (€)\nAnt',
            'SI (Ud)\nAct', 'SI (Ud)\nAnt', 'SI (€)\nAct', 'SI (€)\nAnt',
            'Stock\nactual', 'NIF', 'BOOK', 'SOLAR', 'PVL', 'IVA', 'MARCA', 'SUBMARCA',
            'Compra PUC', 'Compra PVL', 'Laboratorio Categorizado',
        ]
        ordered = [c for c in column_order if c in df.columns]
        extras = [c for c in df.columns if c not in column_order]
        df = df[ordered + extras].reset_index(drop=True)

        # Write Compra PUC / Compra PVL as Excel formulas (not precalculated values),
        # referencing the final cell positions:
        #   Compra PUC = SI (€) Act / (1 + IVA)      Compra PVL = PVL * SI (Ud) Act
        cols = list(df.columns)
        col = lambda name: xl_col_to_name(cols.index(name))
        si_eur, iva, pvl, si_ud = col('SI (€)\nAct'), col('IVA'), col('PVL'), col('SI (Ud)\nAct')
        excel_rows = range(2, len(df) + 2)  # data starts on row 2 (row 1 is the header)
        df['Compra PUC'] = [f"={si_eur}{r}/(1+{iva}{r})" for r in excel_rows]
        df['Compra PVL'] = [f"={pvl}{r}*{si_ud}{r}" for r in excel_rows]

        self.df_bifarma_final = df
        # Single-lab runs go to a distinct file so they don't overwrite the master.
        lab_tag = ""
        if self.laboratori:
            lab_tag = "_" + "".join(c if c.isalnum() else "_" for c in self.laboratori)[:40]
        out_path = self._file_ + rf"\output\df_bifarma_output{period}{lab_tag}.xlsx"
        df.to_excel(out_path, sheet_name="SI Acord book", index=False, header=True,
                    engine='xlsxwriter', na_rep="#N/D")
        self.out_path = out_path
        print(f"done! -> {os.path.basename(out_path)}")

    # Always-on: upserts each product's conditions (PVL/IVA/MARCA/SUBMARCA) into
    # output/product_conditions.csv keyed by Cod Unif, so the frontend can consult
    # them without opening the network YTD Excel. Latest run wins per product.
    def _upsert_conditions_csv(self):
        path = os.path.join(self._file_, "output", "product_conditions.csv")
        # 'Laboratorio Categorizado' first: it is the real classification, the one
        # every calculation and every report is grouped by. The raw BIFarma
        # 'Laboratorio' stays as a second column because it is what has to be
        # mapped in the Mapa_Acords when the category is missing.
        cols = ['Cod Unif', 'Producto', 'Laboratorio Categorizado', 'Laboratorio',
                'PVL', 'IVA', 'MARCA', 'SUBMARCA']
        new = self.df_bifarma_final[[c for c in cols if c in self.df_bifarma_final.columns]].copy()
        if 'Laboratorio Categorizado' in new.columns:
            cat = new['Laboratorio Categorizado']
            new['Laboratorio Categorizado'] = cat.where(cat.notna(), '#N/D')
        new['Cod Unif'] = pd.to_numeric(new['Cod Unif'], errors='coerce').astype('Int64')
        new = new.dropna(subset=['Cod Unif']).drop_duplicates(subset=['Cod Unif'], keep='last')
        new['Actualitzat'] = self.proc_month

        if os.path.exists(path):
            old = pd.read_csv(path)
            old['Cod Unif'] = pd.to_numeric(old['Cod Unif'], errors='coerce').astype('Int64')
            # Upsert: drop old rows for products present in this run, then append new.
            old = old[~old['Cod Unif'].isin(new['Cod Unif'])]
            combined = pd.concat([old, new], ignore_index=True)
        else:
            combined = new
        combined = combined.sort_values('Cod Unif').reset_index(drop=True)
        # Keep the canonical column order even when the CSV on disk predates a
        # column (older rows simply have it empty until their next run).
        order = [c for c in cols + ['Actualitzat'] if c in combined.columns]
        combined = combined[order + [c for c in combined.columns if c not in order]]
        # utf-8-sig so Excel opens the accents correctly.
        combined.to_csv(path, index=False, encoding='utf-8-sig')
        print(f"product_conditions.csv: {len(new)} productes actualitzats ({self.proc_month}), total {len(combined)}")

    # Escriu el master a ClickHouse (base `tancament`), la capa
    # analítica/serving única del grup (DuckDB en retirada). Escriu per HTTP amb la
    # stdlib, sense cap driver. Try/except propi: un problema de ClickHouse mai no pot
    # trencar un tancament. Config per entorn (CLICKHOUSE_URL/USER/PASSWORD) a store_ch.
    def _desa_a_clickhouse(self, period):
        try:
            import store_ch
            r = store_ch.desa_tancament_ch(self.df_bifarma_final, self.proc_date, period)
            fets = ("no (mestre YTD: els fets es guarden només al tancament mensual)"
                    if r["fets"] is None else f"{r['fets']:,} files")
            print(f"ClickHouse: fets {fets} · {r['productes']:,} productes · "
                  f"{r['farmacies']:,} farmàcies · base amb {r['total_fets']:,} fets "
                  f"de {r['mesos']} mes(os)", flush=True)
        except Exception as e:
            print(f"Avís: no s'ha pogut escriure a ClickHouse ({type(e).__name__}: {e}). "
                  "El tancament no en depèn: Excel i CSV s'han generat igualment.", flush=True)

    # Reads the net/gross fee base per lab from the SEGUIMENT file: True = Neto
    # (base = Compra PUC, driven by IVA), False = Bruto (base = Compra PVL, driven
    # by PVL). Parafarmacia wins over EFG on overlap. Missing file -> empty dict.
    def _read_neto_flags(self):
        seg_path = os.path.join(self.rappel_path, f"a.SEGUIMENT LABS {self.proc_year}.xlsx")
        flags = {}

        def add(lab, neto):
            if lab is None or (isinstance(lab, float) and pd.isna(lab)):
                return
            key = str(lab).strip().upper()
            if key in ("", "NAN", "TOTAL GENERAL") or key in flags:
                return
            flags[key] = str(neto).strip().upper().startswith("S")

        try:
            pf = pd.read_excel(seg_path, sheet_name="Parafarmacia", header=1, usecols="A:B")
            pf.columns = ["lab", "neto"]
            for _, r in pf.iterrows():
                add(r["lab"], r["neto"])
            efg = pd.read_excel(seg_path, sheet_name="EFG", header=1, usecols="A:C")
            efg.columns = ["lab", "idlab", "neto"]
            for _, r in efg.iterrows():
                add(r["lab"], r["neto"])
        except Exception as e:
            print(f"Avís: no s'han pogut llegir les condicions Neto del SEGUIMENT ({e}). "
                  "El report de diners perduts assumirà base bruta (PVL).")
        return flags

    # Always-on: per-lab "lost money" report. For each lab, whether the fee base
    # is net (PUC, driven by IVA) or gross (PVL, driven by PVL), computes the share
    # of the lab's total SI (€) whose key field is 0 / N/D / new (MARCA "New!")
    # -> money we can't compute a fee on. Weighted by SI so big products count more.
    def _lost_money_report(self):
        df = self.df_bifarma_final.copy()
        si_col = "SI (€)\nAct"
        if si_col not in df.columns or "PVL" not in df.columns or "IVA" not in df.columns:
            print("Avís: falten columnes per al report de diners perduts; s'omet.")
            return
        df[si_col] = pd.to_numeric(df[si_col], errors="coerce").fillna(0)
        df["PVL"] = pd.to_numeric(df["PVL"], errors="coerce")
        df["IVA"] = pd.to_numeric(df["IVA"], errors="coerce")
        # Group by the REAL classification only. Rows with no category used to
        # fall back to the raw BIFarma name, which split one lab across several
        # rows (ABBOTT vs Abbott) and listed distributors and legal names
        # (CARBO VITAL S.L., HALEON SPAIN S.A.) as if they were labs. They now
        # group under "#N/D", keeping the raw name in its own column so it is
        # still clear what has to be mapped in the Mapa_Acords.
        cat = df.get("Laboratorio Categorizado")
        if cat is None:
            cat = pd.Series(pd.NA, index=df.index)
        cat = cat.astype("string").str.strip()
        df["_lab"] = cat.where(cat.notna() & (cat != ""), "#N/D")
        bif = df["Laboratorio"].astype("string").str.strip().fillna("")
        df["_bif"] = bif.where(df["_lab"] == "#N/D", "")
        marca = df["MARCA"].astype(str).str.upper() if "MARCA" in df.columns else None
        # Count DISTINCT products, not master lines: the master has one line per
        # product AND pharmacy, so a plain row count inflated 'Productes' by the
        # number of pharmacies carrying it (6x to 31x depending on the lab).
        cod = "Cod Unif" if "Cod Unif" in df.columns else None
        flags = self._read_neto_flags()

        rows = []
        for (lab, bif_name), g in df.groupby(["_lab", "_bif"]):
            neto = flags.get(lab.upper())
            if neto is None:
                base_label, field, suffix = "Bruta (PVL)", "PVL", " [sense condició]"
            elif neto:
                base_label, field, suffix = "Neta (PUC)", "IVA", ""
            else:
                base_label, field, suffix = "Bruta (PVL)", "PVL", ""
            vals = g[field]
            is_new = marca.loc[g.index].eq("NEW!") if marca is not None else False
            missing = vals.isna() | (vals == 0) | is_new
            si_total = g[si_col].sum()
            si_lost = g.loc[missing, si_col].sum()
            pct = (si_lost / si_total * 100) if si_total else 0.0
            rows.append({
                "Laboratorio Categorizado": lab,
                "Laboratorio (BIF)": bif_name,
                "Base fee": base_label + suffix,
                "Camp clau": field,
                "SI total (€)": round(float(si_total), 2),
                "SI perdut (€)": round(float(si_lost), 2),
                "% perdut (sobre SI)": round(float(pct), 1),
                "Productes": int(g[cod].nunique()) if cod else int(len(g)),
                "Perduts": int(g.loc[missing, cod].nunique()) if cod else int(missing.sum()),
                "Linies (prod x farmacia)": int(len(g)),
            })
        report = pd.DataFrame(rows).sort_values("% perdut (sobre SI)", ascending=False)
        path = os.path.join(self._file_, "output", "lost_money.csv")
        report.to_csv(path, index=False, encoding="utf-8-sig")
        print(f"lost_money.csv: {len(report)} laboratoris analitzats.")
        # També a ClickHouse (tancament.diners_perduts) perquè Metabase el serveixi.
        # Fail-soft: un problema de ClickHouse mai trenca el tancament.
        try:
            import store_ch
            store_ch.desa_diners_perduts(report, self.proc_date)
            print(f"ClickHouse: diners perduts de {self.proc_month} desats.", flush=True)
        except Exception as e:
            print(f"Avís: no s'ha pogut escriure diners perduts a ClickHouse "
                  f"({type(e).__name__}: {e}).", flush=True)

    def addLab(self):
        # Vectorized: map each Laboratorio to its Acords 'Laboratori' label
        # (first BIF match wins, like the old row loop); no match -> NaN.
        # self.df_acords is already Acord != NO here (cleanData drops those).
        mapping = self.df_acords.drop_duplicates(subset='BIF').set_index('BIF')['Laboratori']
        self.df_bifarma_final['Laboratorio Categorizado'] = self.df_bifarma_final['Laboratorio'].map(mapping)
        self._addLabPerCodi()

    # Fills the categories the lab NAME couldn't resolve using the 'Codi BIF'
    # code. Those are BOT PLUS codes, an industry standard, so they are stable,
    # while BIFarma writes the lab name in many variants ('VEMEDIA' vs
    # 'VEMEDIA PH.', 'BIOFORCE AVOGEL' vs 'BIOFORCE ESPANYA AVOGEL'). DataFilters
    # already keeps a row when EITHER the name or the code matches, so without
    # this a row could enter the master by code and still end up with no
    # category at all -- which is where most of the "#N/D" money came from.
    # Name first, code only to fill the gaps: this can never change a category
    # that already resolved, only add the missing ones.
    # A code shared by two different labs (a distributor: P2868 -> BETER and
    # LUXEOL) cannot decide on its own, so it is skipped; those rows are routed
    # later by product description (see splitCarboVital).
    def _addLabPerCodi(self):
        if 'Id Lab' not in self.df_bifarma_final.columns:
            return
        codis = self.df_acords[['Codi BIF', 'Laboratori']].dropna(subset=['Codi BIF', 'Laboratori']).copy()
        codis['cod'] = codis['Codi BIF'].astype(str).str.strip()
        codis = codis[(codis['cod'] != '') & (codis['cod'].str.lower() != 'nan')]
        if codis.empty:
            return
        per_cod = codis.groupby('cod')['Laboratori'].nunique()
        ambigus = sorted(per_cod[per_cod > 1].index)
        if ambigus:
            print(f"Avís: {len(ambigus)} codi(s) BIF apunten a més d'un laboratori, "
                  f"no s'usen per categoritzar: {ambigus}", flush=True)
        taula = (codis[~codis['cod'].isin(ambigus)]
                 .drop_duplicates(subset='cod').set_index('cod')['Laboratori'])
        buit = self.df_bifarma_final['Laboratorio Categorizado'].isna()
        if not buit.any() or taula.empty:
            return
        rec = self.df_bifarma_final.loc[buit, 'Id Lab'].astype(str).str.strip().map(taula)
        self.df_bifarma_final.loc[buit, 'Laboratorio Categorizado'] = rec
        n = int(rec.notna().sum())
        print(f"addLab: {n} de {int(buit.sum())} files sense categoria resoltes pel "
              f"Codi BIF (BOT PLUS).", flush=True)
        if n:
            for lab, k in rec.dropna().value_counts().items():
                print(f"    {lab}: {k} files", flush=True)
    def checkAmox (self):
        if  self.df_bifarma['Producto'].str.contains('amoxici', case=False, na=False).any():
            print("ok RJ")
        else:
            print('''
**************************************************************************************************************************************************************************************
*                                                                                                                                                                                    *
* Amox RJ -> Grupo Producto: ESPECIALIDAD; SubGrupo: EFG; Producto: amoxi; Laboratorio: Reig Jofre                                                                                   *
*                                                                                                                                                                                    *
**************************************************************************************************************************************************************************************
''')
            raise SystemExit(1)  # aborta amb codi != 0 perque l'orquestrador/panell ho detecti
    def noComunesRename(self):
        if self.filt_codUnif.any():
            unique_labs = self.df_bifarma.loc[self.filt_codUnif, 'Laboratorio'].unique()

            # Generate a unique code for each lab and map it to avoid duplicates
            lab_code_map = {lab: f"999999999{str(i).zfill(4)}" for i, lab in enumerate(unique_labs)}

            # Update the Cod Unif values based on the map
            self.df_bifarma.loc[self.filt_codUnif, 'Cod Unif'] = self.df_bifarma.loc[self.filt_codUnif, 'Laboratorio'].map(lab_code_map)
    def splitSensilir(self):
        # since MARCA is something we create manually we have to remember that this f have some margin error induced by human work
        mask = self.df_bifarma_final['MARCA'].isin(["SENSILIS", "COMODYNES", "AXOVITAL"])
        self.df_bifarma_final.loc[mask, ['Laboratorio Categorizado', 'Laboratorio']] = "SENSILIS"
    # Routes the CARBO VITAL S.L. distributor's rows to the right lab by product
    # description -- see CARBO_VITAL_BRANDS for the table and the why.
    def splitCarboVital(self):
        mask = self.df_bifarma_final['Laboratorio'].astype(str).str.strip().str.upper() == "CARBO VITAL S.L."
        if not mask.any():
            return
        prod = self.df_bifarma_final['Producto'].astype(str)
        for lab, patro in CARBO_VITAL_BRANDS:
            hit = mask & prod.str.contains(patro, case=False, regex=True, na=False)
            # Both columns, like splitSensilir/unionSisfarma: the report has to
            # show the lab it belongs to, not the distributor it came through.
            self.df_bifarma_final.loc[hit, ['Laboratorio Categorizado', 'Laboratorio']] = lab
            print(f"CARBO VITAL S.L.: {int(hit.sum())} files -> {lab}", flush=True)
        # Same table, so the two can never drift apart. The unmatched rows keep
        # Laboratorio = CARBO VITAL S.L. on purpose: with no category assigned,
        # knowing which distributor they came through is exactly what's needed
        # to categorize them in the Mapa_Acords.
        cap = "|".join(patro for _, patro in CARBO_VITAL_BRANDS)
        resta = mask & ~prod.str.contains(cap, case=False, regex=True, na=False)
        self.df_bifarma_final.loc[resta, 'Laboratorio Categorizado'] = None
        print(f"CARBO VITAL S.L.: {int(resta.sum())} files d'altres marques sense "
              f"categoria (sortiran com a #N/D)", flush=True)

    def unionSisfarma(self):
        # since MARCA is something we create manually we have to remember that this f have some margin error induced by human work
        mask = self.df_bifarma_final['Producto'].astype(str).str.contains("elmex", case=False, na=False)
        self.df_bifarma_final.loc[mask, ['Laboratorio Categorizado', 'Laboratorio']] = "SISFARMA"

    def __str__(self):
        return f"Rappel: {self.rappel}, Period: {self.period}"

    def nompikis(self):
        # astype(str) covers both the "contains" and the int == 156119 cases
        mask = self.df_bifarma['Cod Nac'].astype(str).str.contains("156119", na=False)
        self.df_bifarma.loc[mask, 'Laboratorio'] = "ECOCEUTICS"

    def detectDuplicateValues(self):
        # We detect duplicate values in the file (only on columns that exist,
        # since the SQL source has no 'Farm'/'Of')
        dup_cols = [c for c in ["Cod Unif","Producto","Id Lab","Laboratorio","Farm","Of","Nombre Oficina"]
                    if c in self.df_bifarma_final.columns]
        self.duplicated = self.df_bifarma_final[self.df_bifarma_final[dup_cols].duplicated()]
        print(self.duplicated)
        if self.duplicated.empty:
            print("No duplicates found")
        else:
            print("Duplicates found")
    def checkAlmirall(self):
        if  self.df_bifarma['Producto'].str.contains('ESERTIA', na=False).any() or self.df_bifarma['Producto'].str.contains('PARAPRES', na=False).any():
            print("ok Almirall")
        else:
            print('''
**************************************************************************************************************************************************************************************
*                                                                                                                                                                                    *
* RX Almirall -> Grupo Producto: ESPECIALIDAD; SubGrupo: Especialidad; Laboratorio: Almirall                                                                                         *
*                                                                                                                                                                                    *
**************************************************************************************************************************************************************************************
''')
            raise SystemExit(1)  # aborta amb codi != 0 perque l'orquestrador/panell ho detecti
    def checkMenaven(self):
        if  self.df_bifarma['Producto'].str.contains('MENAVEN ', na=False).any() :
            print("ok Menarini")
        else:
            print('''
**************************************************************************************************************************************************************************************
*                                                                                                                                                                                    *
* MV Menarini -> Grupo Producto: ESPECIALIDAD; SubGrupo: Especialidad; Producto: menaven                                                                                             *
*                                                                                                                                                                                    *
**************************************************************************************************************************************************************************************
''')
            raise SystemExit(1)  # aborta amb codi != 0 perque l'orquestrador/panell ho detecti
    # Fills PVL/IVA/MARCA from the reference master (matched by EAN or CN6) for
    # rows whose PVL is missing or 0. Currently disabled in the pipeline.
    def newLabs(self):
        if self.df_products_master.empty:
            return
        master = self.df_products_master
        for x, row_lab in self.df_bifarma_final.iterrows():
            pvl = row_lab['PVL']
            if not (pd.isna(pvl) or pvl == 0):
                continue
            match = master[(master['EAN'] == row_lab['Cod Unif']) | (master['CN6'] == row_lab['Cod Unif'])]
            if match.empty:
                continue
            m = match.iloc[0]
            self.df_bifarma_final.at[x, 'PVL'] = m['PVL']
            self.df_bifarma_final.at[x, 'IVA'] = m['IVA']
            self.df_bifarma_final.at[x, 'MARCA'] = m['MARCA']
            self.df_bifarma_final.at[x, 'SUBMARCA'] = "-"
    def superestalvi(self):
        # We check if the lab is in the superestalvi list and we change the name to "SUPERESTALVI"
        superestalvi = [151329, 153335, 171831, 196432, 196433, 214798, 219996, 219997, 260083, 263665, 300293, 395715, 395756]
        # Numeric coercion so string codes ("151329") also match the int list
        mask = pd.to_numeric(self.df_bifarma['Cod Unif'], errors='coerce').isin(superestalvi)
        self.df_bifarma.loc[mask, 'Laboratorio'] = "SUPERESTALVI"
                 



if __name__ == "__main__":
    import argparse
    # ([bifarma, BIFARMA con BAJAS], [YTD, ""]). year/month optional: default to
    # the last complete month. Frontend example: --rappel BIFARMA --period YTD --year 2026 --month 6
    parser = argparse.ArgumentParser(description="Genera el mestre BIFarma (output/df_bifarma_output).")
    parser.add_argument("--rappel", default="BIFARMA", help='"BIFARMA" o "bifarma" (con BAJAS)')
    parser.add_argument("--period", default="", choices=["", "YTD"], help='"" mensual, "YTD" acumulat')
    parser.add_argument("--year", type=int, default=None, help="Any del mes a processar (opcional)")
    parser.add_argument("--month", type=int, default=None, help="Mes a processar 1-12 (opcional)")
    parser.add_argument("--only-para", action="store_true",
                        help="PARAFARMACIA + l'especialitat pactada (RX Almirall, amoxi Reig Jofre). "
                             "Sense la flag entra tota l'especialitat.")
    parser.add_argument("--laboratori", default=None,
                        help="Nom de laboratori (columna 'Laboratori' del Mapa_Acords): informe d'un sol lab")
    args = parser.parse_args()
    clsBiFarmaEco(args.rappel, args.period, args.year, args.month,
                  only_para=args.only_para, laboratori=args.laboratori)

