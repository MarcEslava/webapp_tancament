import os, json, urllib.request, urllib.parse
import pandas as pd
from store import _split_master   # split de columnas del master (única fuente de verdad)

'''
**************************************************************************************************************************************************************************************
ClickHouse store for the monthly close.

Same model as store.py (fase 0 de docs/PLA_DADES.md) pero aterrizado en ClickHouse,
que es la capa analítica/serving única del grupo (se retira DuckDB). Escribe por HTTP,
así que la máquina del panell no necesita ningún driver.

Motores:
  - tancament.fet_mes      MergeTree particionado por mes. Re-ejecutar un mes es
                           idempotente: se hace DROP PARTITION del mes y se reinserta.
  - tancament.dim_producte ReplacingMergeTree(carregat_el) por cod_unif -> el último
  - tancament.dim_farmacia ReplacingMergeTree(carregat_el) por (farm, of)   write manda.

Igual que store.py: solo ESCRIBE, en paralelo. Se llama dentro de un try/except desde
index.py, así que un fallo aquí nunca puede tumbar un tancament.

Config por entorno (defaults de dev):
  CLICKHOUSE_URL (http://localhost:8123) · CLICKHOUSE_USER · CLICKHOUSE_PASSWORD · CLICKHOUSE_DB_TANCAMENT (tancament)

Author: Marc Eslava
**************************************************************************************************************************************************************************************
'''

# El password surt de l'entorn o del db_config.py (git-ignored), mai hardcodejat.
def _secret(name, default=""):
    v = os.environ.get(name)
    if v:
        return v
    try:
        import db_config
        return getattr(db_config, name, None) or default
    except Exception:
        return default

CH_URL  = os.environ.get("CLICKHOUSE_URL", "http://localhost:8123").rstrip("/")
CH_USER = os.environ.get("CLICKHOUSE_USER", "price_monitor")
CH_PASS = _secret("CLICKHOUSE_PASSWORD", "changeme")
CH_DB   = os.environ.get("CLICKHOUSE_DB_TANCAMENT", "tancament")

_NUMERIC = {"so_ud", "so_eur", "si_ud", "si_eur", "stock", "pvl", "iva",
            "si_total", "si_perdut", "pct_perdut"}
_INT     = {"cod_unif", "productes", "perduts", "linies"}
_DATES   = {"mes", "actualitzat"}


def _ch(sql, body=None):
    h = {"X-ClickHouse-User": CH_USER, "X-ClickHouse-Key": CH_PASS}
    if body is None:                                   # DDL / comandos: SQL en el cuerpo
        req = urllib.request.Request(CH_URL + "/", data=sql.encode("utf-8"), headers=h)
    else:                                              # INSERT: SQL en query param, filas en el cuerpo
        req = urllib.request.Request(CH_URL + "/?" + urllib.parse.urlencode({"query": sql}), data=body, headers=h)
    try:
        return urllib.request.urlopen(req, timeout=180).read().decode()
    except urllib.error.HTTPError as e:
        raise RuntimeError("ClickHouse: " + e.read().decode()[:400])


def _cell(k, v):
    if k in _DATES:
        return None if pd.isna(v) else pd.Timestamp(v).strftime("%Y-%m-%d")
    if k in _NUMERIC:
        return None if pd.isna(v) else float(v)
    if k in _INT:
        return None if pd.isna(v) else int(v)
    # strings: null -> "" (columnas String no-nullable en ClickHouse)
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v)


def _insert(table, df):
    if df is None or df.empty:
        return 0
    lines = []
    for _, r in df.iterrows():
        lines.append(json.dumps({k: _cell(k, v) for k, v in r.items()}, ensure_ascii=False))
    body = "\n".join(lines).encode("utf-8")
    _ch(f"INSERT INTO {CH_DB}.{table} FORMAT JSONEachRow", body=body)
    return len(df)


def _select(sql):
    """SELECT ... -> lista de dicts. Añade FORMAT JSONEachRow si falta."""
    if "FORMAT " not in sql.upper():
        sql += " FORMAT JSONEachRow"
    return [json.loads(l) for l in _ch(sql).splitlines() if l.strip()]


# --- Editor de gobernanza: PVL/IVA/MARCA/SUBMARCA por producte ----------------
# dim_producte ES el store gobernat: ReplacingMergeTree(carregat_el) -> l'última
# escriptura per cod_unif mana (= versionat). Pas 1 llegeix aquests 4 camps d'aquí
# (store_ch.read_producte_attr) en comptes del YTD Excel; l'editor web hi escriu
# overrides (desa_producte_attr) amb carregat_el=now(). Com que el tancament també
# reescriu dim_producte amb el que Pas 1 ha llegit, el bucle és estable: l'edició
# no es perd.

# Columnes de dim_producte (per reinserir la fila sencera: ReplacingMergeTree
# reemplaça la fila, no fa merge de camps). carregat_el s'omet -> DEFAULT now().
_DIMPROD_COLS = ["cod_unif", "producte", "id_lab", "laboratori",
                 "laboratori_categoritzat", "pvl", "iva", "marca", "submarca", "actualitzat"]


def read_producte_attr():
    """Els atributs curats per producte des de dim_producte (FINAL), amb els noms
    de columna del master de Pas 1 (Cod Unif/PVL/IVA/MARCA/SUBMARCA) per ser
    drop-in de list_anterior a index.py.transformNumValue."""
    rows = _select(f"SELECT cod_unif, pvl, iva, marca, submarca "
                   f"FROM {CH_DB}.dim_producte FINAL")
    df = pd.DataFrame(rows, columns=["cod_unif", "pvl", "iva", "marca", "submarca"])
    return df.rename(columns={"cod_unif": "Cod Unif", "pvl": "PVL", "iva": "IVA",
                              "marca": "MARCA", "submarca": "SUBMARCA"})


def _lit(s):
    """Literal de cadena per a ClickHouse (escapa \\ i '). Eina interna a la LAN."""
    return "'" + str(s).replace("\\", "\\\\").replace("'", "\\'") + "'"


def cerca_productes(q="", only_missing=False, limit=200):
    """Cerca a dim_producte (FINAL) per a l'editor. `q` casa amb cod_unif (si és
    numèric), producte o laboratori. `only_missing` filtra els que necessiten
    curació (New!/buit/PVL o IVA a 0). Els New! primer. Retorna llista de dicts."""
    q = (q or "").strip()
    where = []
    if q:
        ql = _lit(f"%{q}%")
        cond = f"(producte ILIKE {ql} OR laboratori_categoritzat ILIKE {ql}"
        if q.isdigit():
            cond += f" OR cod_unif = {int(q)}"
        where.append(cond + ")")
    if only_missing:
        where.append("(marca='New!' OR marca='' OR pvl IS NULL OR pvl=0 OR iva IS NULL OR iva=0)")
    w = ("WHERE " + " AND ".join(where)) if where else ""
    limit = max(1, min(int(limit or 200), 1000))
    return _select(
        f"SELECT cod_unif, producte, laboratori_categoritzat, pvl, iva, marca, submarca "
        f"FROM {CH_DB}.dim_producte FINAL {w} "
        f"ORDER BY (marca='New!') DESC, cod_unif LIMIT {limit}")


def desa_producte_attr(edits):
    """Aplica edicions (override) sobre dim_producte. `edits` = llista de dicts amb
    'cod_unif' i qualsevol de pvl/iva/marca/submarca. Llegeix la fila actual (FINAL)
    de cada producte, hi aplica els canvis i reinsereix (carregat_el=now() -> mana).
    Retorna el nombre de files escrites."""
    by_cod = {int(e["cod_unif"]): e for e in (edits or []) if e.get("cod_unif") is not None}
    if not by_cod:
        return 0
    codes = ",".join(str(c) for c in by_cod)
    cur = {int(r["cod_unif"]): r for r in _select(
        f"SELECT {', '.join(_DIMPROD_COLS)} FROM {CH_DB}.dim_producte FINAL "
        f"WHERE cod_unif IN ({codes})")}
    out = []
    for cod, e in by_cod.items():
        row = cur.get(cod) or {c: None for c in _DIMPROD_COLS}
        row["cod_unif"] = cod
        for f in ("pvl", "iva", "marca", "submarca"):
            if f in e and e[f] is not None:
                row[f] = e[f]
        if not row.get("actualitzat"):
            row["actualitzat"] = "1970-01-01"      # Date no-nullable: mai NULL
        for c in ("producte", "laboratori", "laboratori_categoritzat", "id_lab", "marca", "submarca"):
            if row.get(c) is None:
                row[c] = ""
        row.pop("carregat_el", None)               # DEFAULT now() -> l'última mana
        out.append({c: row.get(c) for c in _DIMPROD_COLS})
    return _insert("dim_producte", pd.DataFrame(out))


# Escribe un mestre de Pas 1 a ClickHouse. `mes` = mes procesado (date, dia 1);
# `period` = "" (mensual) o "YTD". Los YTD son acumulados: refrescan dimensiones
# pero NUNCA escriben hechos (contarían cada mes dos veces).
def desa_tancament_ch(df, mes, period):
    facts, products, pharmacies = _split_master(df, mes)
    n_prod = _insert("dim_producte", products)          # ReplacingMergeTree: el último carregat_el manda
    n_farm = _insert("dim_farmacia", pharmacies)
    n_fets = None
    if period != "YTD" and not facts.empty:
        yyyymm = int(pd.Timestamp(mes).strftime("%Y%m"))
        _ch(f"ALTER TABLE {CH_DB}.fet_mes DROP PARTITION {yyyymm}")   # re-ejecutar el mes es idempotente
        n_fets = _insert("fet_mes", facts)
    total = _ch(f"SELECT count(), uniqExact(mes) FROM {CH_DB}.fet_mes FORMAT TSV").split()
    return {"fets": n_fets, "productes": n_prod, "farmacies": n_farm,
            "total_fets": int(total[0]) if total else 0,
            "mesos": int(total[1]) if len(total) > 1 else 0}


# Report de "diners perduts" (per lab, del mes) -> ClickHouse. index.py el calcula
# amb la lògica exacta (neto/bruto del SEGUIMENT) i aquí es desa perquè Metabase el
# llegeixi tal qual. Idempotent: DROP PARTITION del mes + insert.
_DP_MAP = {
    "Laboratorio Categorizado": "laboratori_categoritzat",
    "Laboratorio (BIF)": "laboratori_bif",
    "Base fee": "base_fee", "Camp clau": "camp_clau",
    "SI total (€)": "si_total", "SI perdut (€)": "si_perdut",
    "% perdut (sobre SI)": "pct_perdut", "Productes": "productes",
    "Perduts": "perduts", "Linies (prod x farmacia)": "linies",
}
_DP_COLS = ["mes", "laboratori_categoritzat", "laboratori_bif", "base_fee", "camp_clau",
            "si_total", "si_perdut", "pct_perdut", "productes", "perduts", "linies"]


def desa_diners_perduts(df, mes):
    if df is None or df.empty:
        return 0
    d = df.rename(columns=_DP_MAP).copy()
    # "#N/D" no és un laboratori real (productes sense categoritzar) -> no puja a la
    # BD ni compta a l'anàlisi. El detall de què cal mapejar queda al CSV/detector.
    if "laboratori_categoritzat" in d.columns:
        d = d[d["laboratori_categoritzat"].astype(str).str.strip() != "#N/D"]
    if d.empty:
        return 0
    d["mes"] = pd.Timestamp(mes).strftime("%Y-%m-%d")
    d = d[[c for c in _DP_COLS if c in d.columns]]
    yyyymm = int(pd.Timestamp(mes).strftime("%Y%m"))
    _ch(f"ALTER TABLE {CH_DB}.diners_perduts DROP PARTITION {yyyymm}")
    return _insert("diners_perduts", d)
