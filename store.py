import pandas as pd

'''
**************************************************************************************************************************************************************************************
Helpers de "split" del master del tancament (Pas 1).

Abans això era el store DuckDB (fase 0 de docs/PLA_DADES.md). **DuckDB s'ha retirat**:
la capa analítica/serving única del grup és ClickHouse (veure store_ch.py). Aquí es
conserva només el split de columnes del master, perquè store_ch el reutilitza i el
mapeig de columnes és únic (una sola font de veritat).

_split_master parteix el master de Pas 1 en (facts, products, pharmacies) amb els noms
de columna del model tancament: fet_mes (mes × producte × farmàcia), dim_producte i
dim_farmacia.

Author: Marc Eslava
**************************************************************************************************************************************************************************************
'''


# Excel and SQL disagree on how they hand over the same id: one says 123, the
# other 123.0 or " 123". Normalize every key to a plain string so the same
# pharmacy/lab doesn't end up stored twice.
def _txt(s):
    def one(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return None
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        t = str(v).strip()
        return t or None
    return s.map(one)


def _num(s):
    return pd.to_numeric(s, errors="coerce")


# Splits the Pas 1 master into the three tables. Returns (facts, products,
# pharmacies) as DataFrames with the column names of the tancament schema.
def _split_master(df, mes):
    cod = _num(df.get("Cod Unif")).astype("Int64")
    keep = cod.notna()

    facts = pd.DataFrame({
        "mes": mes,
        "cod_unif": cod,
        "id_lab": _txt(df.get("Id Lab")),
        "farm": _txt(df.get("Farm")),
        "of": _txt(df.get("Of")),
        "so_ud": _num(df.get("SO (Ud)\nAct")),
        "so_eur": _num(df.get("SO (€)\nAct")),
        "si_ud": _num(df.get("SI (Ud)\nAct")),
        "si_eur": _num(df.get("SI (€)\nAct")),
        "stock": _num(df.get("Stock\nactual")),
    })[keep]

    products = pd.DataFrame({
        "cod_unif": cod,
        "producte": df.get("Producto"),
        "id_lab": _txt(df.get("Id Lab")),
        "laboratori": df.get("Laboratorio"),
        "laboratori_categoritzat": df.get("Laboratorio Categorizado"),
        "pvl": _num(df.get("PVL")),
        "iva": _num(df.get("IVA")),
        "marca": df.get("MARCA"),
        "submarca": df.get("SUBMARCA"),
        "actualitzat": mes,
    })[keep].drop_duplicates(subset=["cod_unif"], keep="last")

    pharmacies = pd.DataFrame({
        "farm": _txt(df.get("Farm")),
        "of": _txt(df.get("Of")),
        "nombre_oficina": df.get("Nombre Oficina"),
        "nif": _txt(df.get("NIF")),
        "book": _txt(df.get("BOOK")),
        "solar": _txt(df.get("SOLAR")),
        "actualitzat": mes,
    })
    pharmacies = (pharmacies.dropna(subset=["farm", "of"])
                            .drop_duplicates(subset=["farm", "of"], keep="last"))
    return facts, products, pharmacies
