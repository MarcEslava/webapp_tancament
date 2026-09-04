import os
import re

import pandas as pd

'''
**************************************************************************************************************************************************************************************
Detector de referencies que el BIFarma web te i la replica SQL no.

El pipeline llegeix de la base de dades SQL (bench_dwComprasVentasMesS +
tbi_productosERS). El grid del BIFarma web, en canvi, pot tenir referencies que
encara no han arribat a aquella replica: la fila simplement no existeix, i el
pipeline no la pot perdre perque no hi ha res a perdre. El cas que ho va treure
a la llum va ser LUXEOL RIZOS PACK VIAJE (124,99 EUR de compra): 15 de les 16
referencies de LUXEOL quadraven al centim amb l'export, i la 16a no era al
mestre de productes sota cap nom ni dins el rang d'EAN del laboratori.

Aixo NO es pot arreglar en codi -- la referencia s'ha de donar d'alta o
sincronitzar. El que fa aquest modul es que no ho descobreixis quan un
laboratori et diu que li falten diners.

Que compara, i que NO compara
-----------------------------
Compara EXISTENCIA de la descripcio del producte al mestre, no imports. Els
imports de la BD no son comparables amb els de l'export: el grid es una vista
filtrada de la qual no sabem l'abast, aixi que qualsevol quadre d'euros
donaria diferencies que no volen dir res. Els euros que surten a l'informe son
els de l'export, i nomes serveixen per ordenar per impacte.

Per que per nom i no per codi
-----------------------------
La columna "Id" de l'export es tbi_productosERS.idproducto, pero aquell id es
PER FARMACIA: el mateix '000268' es el pack de LUXEOL a la farmacia que ha
exportat i un PRANAROM en una altra. Sense saber de quina farmacia surt cada
fila, l'id sol no decideix res. El nom, en canvi, es el mateix a les dues
bandes quan el producte existeix (verificat: 925 de 988 files d'un export real
casaven la parella (id, nom) exactament).

Com que els noms s'escriuen diferent entre els dos sistemes
("MATTERS ULLERA ARCILLA 3.5" al grid i "MATTERS ARCILLA +1.5" al mestre), una
llista de noms sense coincidencia esta plena de variants que si existeixen. Per
aixo cada fila porta els vens mes propers del mestre i el seu laboratori: amb
aixo es veu d'un cop d'ull si es una referencia nova de veritat o nomes el
mateix producte escrit d'una altra manera. No hi ha veredicte automatic perque
no n'hi ha cap que es pugui sostenir.
**************************************************************************************************************************************************************************************
'''

# "Compra\nImpte\nAc 26" -> any 2026. Al grid les columnes venen amb l'any a dos
# digits i amb salts de linia pel mig.
RE_COMPRA = re.compile(r"compra.*impte.*ac\s*(\d{2})", re.IGNORECASE | re.DOTALL)

MAX_VEINS = 3       # vens del mestre que ensenyem per fila
MAX_FILES = 2000    # sostre de files a l'informe (l'export tipic en te ~250)


def _norm(s):
    return " ".join(str(s).split()).strip().upper()


def _primer_mot(s):
    parts = _norm(s).split()
    return parts[0] if parts else ""


def _db():
    import db_config as cfg
    from SQLConnection import SQLConnection
    return SQLConnection(
        db_host=cfg.DB_HOST, db_port=getattr(cfg, "DB_PORT", 1433),
        db_database=cfg.DB_NAME, db_username=cfg.DB_USER,
        db_password=cfg.DB_PASS, dialect="mssql", driver="pymssql",
        login_timeout=15, connect_retries=3, retry_delay=15,
    )


def llegeix_export(path):
    """El grid de BIFarma -> DataFrame amb prod / id / eur_act / eur_ant.

    dtype=str + keep_default_na=False a proposit: hi ha un producte que es diu
    literalment "NAN" (Nestle) que pandas convertiria en valor buit, i els
    idproducto son cadenes amb zeros al davant ('000110') que com a numero
    perdrien els zeros.
    """
    df = pd.read_excel(path, dtype=str, keep_default_na=False)

    if "Producto" not in df.columns:
        raise ValueError("L'export no te columna 'Producto'. "
                         "Fa la pinta de no ser un BIExportGrid de BIFarma.")

    # Les dues columnes de compra: l'any mes alt es l'actual, l'altre l'anterior.
    anys = {}
    for c in df.columns:
        m = RE_COMPRA.search(str(c))
        if m:
            anys[int(m.group(1))] = c
    ordre = sorted(anys, reverse=True)
    col_act = anys.get(ordre[0]) if ordre else None
    col_ant = anys.get(ordre[1]) if len(ordre) > 1 else None

    def num(col):
        if col is None:
            return 0.0
        return pd.to_numeric(df[col].astype(str).str.replace(",", ".").str.strip(),
                             errors="coerce").fillna(0.0)

    out = pd.DataFrame({
        "prod": df["Producto"].map(_norm),
        "id": df["Id"].astype(str).str.strip() if "Id" in df.columns else "",
        "eur_act": num(col_act),
        "eur_ant": num(col_ant),
    })
    out = out[out["prod"] != ""]
    return out, (ordre[0] if ordre else None), (ordre[1] if len(ordre) > 1 else None)


def analitza(path):
    """Retorna (columns, rows, resum) amb les referencies que el mestre no coneix."""
    exp, any_act, any_ant = llegeix_export(path)

    # Un registre per descripcio: sumem els euros de totes les farmacies.
    per_prod = exp.groupby("prod", as_index=False).agg(
        eur_act=("eur_act", "sum"), eur_ant=("eur_ant", "sum"),
        files=("prod", "size"),
        ids=("id", lambda s: ", ".join(sorted({v for v in s if v})[:6])),
    )
    # Tot el cataleg de noms del mestre en una sola consulta (~86.000 parelles,
    # 13 MB, 3 s). Surt mes barat que trossejar IN/LIKE per lots, i sobretot
    # permet comparar amb els espais interns normalitzats: al mestre hi ha noms
    # com "IROHA MASC  FACIAL HYDROGEL COLLAGEN" amb espai doble, que contra
    # l'export donarien un fals positiu. SQL Server no col.lapsa espais, Python si.
    with _db() as db:
        cat = db.fech_dataframe("""
            SELECT DISTINCT UPPER(LTRIM(RTRIM(pr.desproducto))) AS p,
                   LTRIM(RTRIM(ISNULL(pr.deslab, ''))) AS lab
            FROM dbo.tbi_productosERS pr WITH (NOLOCK)
            WHERE pr.desproducto IS NOT NULL""")

    coneguts, veins = set(), {}
    for p, lab in zip(cat["p"], cat["lab"]):
        nom = _norm(p)
        if not nom:
            continue
        coneguts.add(nom)
        veins.setdefault(_primer_mot(nom), []).append((nom, lab or ""))

    desconeguts = per_prod[~per_prod["prod"].isin(coneguts)].copy()

    def veins_de(nom):
        cands = veins.get(_primer_mot(nom), [])
        # Els mes semblants primer: mes paraules compartides amb el nom buscat.
        toks = set(_norm(nom).split())
        cands = sorted(cands, key=lambda c: -len(toks & set(c[0].split())))
        return cands

    # Noms de columna fixos (l'any va al resum): aixi la capcalera del CSV no
    # canvia d'un mes a l'altre i el frontend pot alinear els numeros.
    col_act, col_ant = "Compra (€) actual", "Compra (€) anterior"
    files = []
    for _, r in desconeguts.iterrows():
        cands = veins_de(r["prod"])
        labs = sorted({lab for _, lab in cands if lab})
        files.append({
            "Producte (export)": r["prod"],
            col_act: round(r["eur_act"], 2),
            col_ant: round(r["eur_ant"], 2),
            "Files a l'export": int(r["files"]),
            "Id producte": r["ids"],
            "Refs del mateix nom al mestre": len(cands),
            "Vens al mestre": " | ".join(p for p, _ in cands[:MAX_VEINS]),
            "Laboratori dels vens": " | ".join(labs[:MAX_VEINS]),
        })

    # Per impacte: primer les que no tenen cap ve al mestre (candidates fortes a
    # ser noves de veritat), i dins de cada grup pels euros de compra.
    files.sort(key=lambda f: (f["Refs del mateix nom al mestre"] > 0, -abs(f[col_act])))
    files = files[:MAX_FILES]

    cols = ["Producte (export)", col_act, col_ant, "Files a l'export", "Id producte",
            "Refs del mateix nom al mestre", "Vens al mestre", "Laboratori dels vens"]
    rows = [[f[c] for c in cols] for f in files]

    resum = {
        "productes_export": len(per_prod),
        "files_export": int(exp.shape[0]),
        "desconeguts": len(desconeguts),
        "sense_cap_ve": sum(1 for f in files if f["Refs del mateix nom al mestre"] == 0),
        "eur_desconeguts": round(float(desconeguts["eur_act"].sum()), 2),
        "fitxer": os.path.basename(path),
        "any_actual": 2000 + any_act if any_act else None,
        "any_anterior": 2000 + any_ant if any_ant else None,
    }
    return cols, rows, resum
