# Pla — treure el tancament dels CSV i els Excels intermedis

> Proposta de model de dades i migració per fases.
> **Estat: fase 0 implementada** (escriptura en paral·lel, res no llegeix encara de la base).
> La resta de fases són proposta pendent d'acord.
>
> Resum de la proposta: **DuckDB** com a magatzem (un fitxer, sense servidor), les dades
> mestres com a taules de veritat en comptes de columnes dins d'Excels mensuals, i l'Excel
> només a les vores del procés.

---

## 1. Diagnòstic

Xifres mesurades sobre les dades reals (agost 2026):

| Artefacte | Mida | Problema |
|---|---|---|
| `output/accumulated_labs.csv` (Z:) | **5,5 M files · 1,2 GB** | Es llegeix i reescriu **sencer** cada tancament, per SMB |
| `output/Seguiment_labs.csv` (Z:) | **1,1 GB** | Es reconstrueix rellegint centenars d'Excels |
| Mestre del Pas 1 | **345.667 files × 27 col** (49 MB xlsx) | Excel al límit del que aguanta |
| `output/product_conditions.csv` | 22.100 productes | **Ja és una taula mestra a mitges** |

### Per què ocupen tant

No és el volum: 5,5 M de files són poques. És que **cada fila repeteix tot el context**.
A `accumulated_labs.csv`, cada línia porta el nom del producte, el del laboratori, `NIF`,
`BOOK`, `SOLAR`, `PVL`, `IVA`, `MARCA`, `SUBMARCA`... per a cada combinació de mes, producte i
farmàcia. Normalitzat, això col·lapsa a una fracció.

També hi arrosseguem brutícia: les columnes `Unnamed: 41` i ` .8` (restes de columnes buides
d'Excel) viatgen dins d'1,2 GB de CSV. I com que [clsSplit.py:426](../clsSplit.py#L426) escriu
amb `index=False`, **el `Cod Unif` es perd**: només sobreviu com a `Cod Nac`.

### Els tres problemes de fons

1. **La veritat viu en un Excel que el propi procés genera.** `MARCA`, `SUBMARCA`, `PVL` i `IVA`
   surten del `Parafarmacia {mes} YTD.xlsx`, que és sortida del Pas 1 i alhora entrada del Pas 1
   següent. D'aquí venen la dependència circular que vam pegar al juliol, els productes «New!»
   que reapareixen cada mes i les marques que no quadren amb Zoho.

2. **El YTD és una execució sencera en comptes d'una consulta.** Es processa dues vegades cada
   mes (mensual i acumulat) i es guarden els dos resultats com a fitxers diferents.

3. **Cada fila duplica l'any anterior.** Les columnes `Act` i `Ant` conviuen a la mateixa fila,
   quan «l'any anterior» és exactament la mateixa taula dotze mesos enrere.

---

## 2. Model proposat

### Taula de fets

**`fet_mes`** — gra: **un mes × un producte × una farmàcia**

| Camp | Origen actual |
|---|---|
| `mes` (date) | nom de carpeta / `month` |
| `cod_unif` | `Cod Unif` |
| `id_lab` | `Id Lab` |
| `farm`, `of` | `Farm`, `Of` |
| `so_ud`, `so_eur` | `SO (Ud) Act`, `SO (€) Act` |
| `si_ud`, `si_eur` | `SI (Ud) Act`, `SI (€) Act` |
| `stock` | `Stock actual` |

**El que desapareix i per què:**

- **Les columnes `Ant`** → l'any anterior és `WHERE mes = target - 12 mesos`. Mateixa taula.
- **Les files YTD** → l'acumulat és `SUM(...) WHERE mes BETWEEN gener AND mes_actual`.
- **`Producto`, `Laboratorio`, `NIF`, `BOOK`, `SOLAR`, `PVL`, `IVA`, `MARCA`, `SUBMARCA`** →
  passen a les dimensions; a la taula de fets hi ha només la clau.
- **`Compra PUC` / `Compra PVL`** → són fórmules (`si_eur/(1+iva)` i `pvl*si_ud`), es calculen
  en consulta. Avui es guarden com a text de fórmula que Excel encara no ha calculat, i per això
  [clsSplit.py:180](../clsSplit.py#L180) les ha de recalcular a mà.

### Taules mestres

**`dim_producte`** — clau `cod_unif`. `producte`, `pvl`, `iva`, `marca`, `submarca`,
`laboratori_categoritzat`, `actualitzat`, `origen` (manual / BIFarma / referències).
És `product_conditions.csv` promogut a font de veritat.

**`dim_laboratori`** — clau `laboratori`. `acord`, `neto`, `fijo`, `variable`, `fee_mes`.
Fusiona `Mapa_Acords.xlsx` amb els fulls `Parafarmacia` i `EFG` d'`a.SEGUIMENT LABS`.

**`lab_bif`** — `nom_bif` / `codi_bif` → `laboratori`. És la correspondència que avui es resol
amb el filtre híbrid per nom o codi de [index.py:363](../index.py#L363).

**`dim_farmacia`** — `nombre_oficina` → `nif`, `book`, `solar`. Ve de `BBDD_Book INTERN.xlsx`.
Aquí també hi cauen les unificacions d'oficines de `unifyCamps()` (Viñamata, Vila).

**`regla_reclassificacio`** — `tipus` (marca | producte_conté | cod_unif), `valor`,
`laboratori_desti`. Substitueix el que ara són llistes escrites dins el codi: SENSILIS /
COMODYNES / AXOVITAL, `elmex` → SISFARMA, CN 156119 → ECOCEUTICS i els 13 codis de SUPERESTALVI.
Passen a ser dades editables, no constants d'un `.py`.

---

## 3. Què guanya cada pas

| Pas | Ara | Amb el model nou |
|---|---|---|
| **1 · Mestre** | Llegeix condicions d'un Excel de 49 MB que ell mateix va generar | Llegeix `dim_producte`. Mor la circularitat i `_find_ytd_source()` deixa de tenir sentit |
| **YTD** | Segona execució completa, amb consulta SQL mes a mes | Una consulta d'agregació |
| **2 · Split** | Igual | **Igual** — continua generant els 78 Excels |
| **3 · Històric** | Rellegeix centenars d'Excels i reescriu 1,1 GB | La taula de fets **ja és** l'històric |
| **Diners perduts** | CSV regenerat a cada execució | Consulta sobre `dim_producte` |

---

## 4. Motor: DuckDB

**Per què DuckDB i no ClickHouse:** 5,5 M de files són poques per a ClickHouse, que està pensat
per a milers de milions. Muntar-lo vol dir mantenir un servidor, còpies i usuaris per a un
conjunt de dades que cap a la RAM. DuckDB és **un sol fitxer**, sense servidor, columnar, i corre
dins el mateix procés del panell.

### On viu el fitxer

**Un sol fitxer, al disc local del servidor.** Mai una còpia per equip: DuckDB és d'un sol
escriptor, i N còpies editables són N veritats divergents — precisament el problema que volem
matar. Tampoc a Z:: una base de dades sobre SMB és font de bloquejos i corrupció.

| On | Què hi ha | Qui hi toca |
|---|---|---|
| Disc local del servidor | el fitxer `.duckdb` | **només el procés del panell** |
| Z: | Excels dels labs, fitxers de referència, còpies | les persones |
| PC dels usuaris | res | el navegador |

Els usuaris no obren mai la base: hi arriben pel panell, igual que avui arriben a les pestanyes
de dades sense saber on és `product_conditions.csv`.

**Per analitzar des de fora**, lectura i mai una segona còpia editable: bolcat a Parquet a Z:, o
Metabase apuntant a la base. Escriure, només el panell.

**Decisions operatives:**

- **Un sol escriptor**: DuckDB ho exigeix i el panell ja ho garanteix amb el seu lock d'una
  execució alhora.
- **Còpia diària a Z:** del fitxer i/o bolcat a Parquet per mes. La còpia és barata; el fitxer
  serà molt més petit que els 2,3 GB de CSV actuals.
- **Si el servidor cau, no hi ha tancament** — però això ja passa avui: el Pas 2 necessita
  l'Excel i la Z: d'aquella màquina. El que canvia és que la política de còpies passa a ser
  explícita en comptes de «ja hi ha els fitxers a Z:».

**Si algun dia cal ClickHouse:** quan això sigui una font entre moltes d'un magatzem de tota
l'empresa, amb consultes concurrents de molta gent. Llavors el pipeline hi carrega i ClickHouse
és la capa d'analítica — no perquè el volum ho demani.

---

## 5. Migració per fases

Cap fase trenca el tancament del mes: fins a la fase 4 l'Excel continua sent la font.

| Fase | Què es fa | Com es valida |
|---|---|---|
| **0** ✅ | Crear la base i escriure-hi **en paral·lel** al final del Pas 1. Res no la llegeix encara. | Fet — veure sota |
| **1** | Carregar l'històric: `accumulated_labs.csv` + els Excels de `{any}\{MM.AAAA}\mes\` | Totals per mes i lab **quadren amb el CSV actual** |
| **2** | El Pas 3 llegeix de DuckDB en comptes de rellegir Excels | Comparar amb `Seguiment_labs.csv` |
| **3** | `dim_producte` passa a font de veritat + **editor al panell** per a MARCA/SUBMARCA | Mateixes marques que l'Excel del mes |
| **4** | El Pas 1 deixa de llegir el `Parafarmacia YTD` | Mestre idèntic al d'un tancament real |
| **5** | El YTD passa a ser una consulta | YTD calculat = YTD generat |
| **6** | Analítica a sobre (Metabase, o ClickHouse si es vol compartir) | — |

Les fases 1 i 2 ja donen el benefici gros (adéu als 2,3 GB de CSV) i **no toquen la
categorització**, que és la part delicada.

### Fase 0 — feta

- **`store.py`** crea la base i hi escriu `fet_mes`, `dim_producte` i `dim_farmacia`.
- **`index.py`** hi bolca el mestre al final del Pas 1 (`_desa_a_duckdb`), dins d'un `try/except`:
  si falla, avisa i el tancament continua. Els Excels i CSV es generen exactament igual que abans.
- Ubicació per defecte: `%LOCALAPPDATA%\WebAppTancament\tancament.duckdb` (disc local, fora
  d'OneDrive i de Z:). Es pot moure amb la variable d'entorn `PANELL_DB`.
- **Els mestres YTD no escriuen fets**, només refresquen dimensions: són acumulats i comptarien
  cada mes dues vegades.

**Validat amb un mestre real de 345.666 files:**

| Comprovació | Resultat |
|---|---|
| Sumes `SI (€)`, `SO (€)`, `SI (Ud)` contra l'Excel d'origen | Idèntiques (diferència < 0,000001) |
| Reexecutar el mateix mes | 345.666 files, no duplica |
| Escriure un mestre YTD | No toca els fets; refresca 21.509 productes |
| Mida al disc | **17,8 MB** per 691.332 fets, contra 97,8 MB dels Excels equivalents |

Extrapolat als 5,5 M de files de l'històric: al voltant de **150 MB** de base, contra **1,2 GB**
de CSV.

---

## 6. Què no canvia

- **Els 78 Excels que reben els laboratoris.** Plantilla, taules dinàmiques i refresc per COM:
  igual. Això obliga a mantenir el servidor Windows amb Excel (veure `INSTALLACIO_SERVIDOR.md`).
- **El SQL de BIFarma** com a origen dels fets.
- **Els Excels de referència** (`Mapa_Acords`, `BBDD_Book INTERN`, `a.SEGUIMENT LABS`) durant tota
  la transició: s'importen a les taules mestres, no cal deixar d'editar-los de cop.

---

## 7. Riscos

- **Doble font mentre duri la transició.** Entre les fases 3 i 4, MARCA es pot editar a l'Excel
  **i** a la base. Cal decidir qui mana i des de quin dia.
- **Validació.** Cada fase ha de quadrar totals contra el que hi ha ara abans de continuar. Si no
  quadra, es para: no es migra res «a veure què surt».
- **Còpies de seguretat.** Avui el «backup» és que tot són fitxers a Z:. Amb una base de dades,
  la còpia s'ha de programar explícitament.
- **Històric a migrar.** L'acumulat actual ha perdut el `Cod Unif` (només hi ha `Cod Nac`) i porta
  columnes escombraria. Cal netejar-lo en carregar-lo, i decidir quants anys enrere val la pena.

---

## 8. Decisions que necessito de tu

1. **Quants anys d'històric** cal migrar? (l'acumulat cobreix ~16 mesos; les carpetes van des de 2023)
2. **Qui pot editar** MARCA/SUBMARCA des del panell — tothom qui hi entra, o cal distingir?
3. **Còpies de seguretat**: n'hi ha prou amb una còpia diària a Z:, o hi ha una política d'empresa?
4. **Metabase**: ha de llegir aquestes dades? Això inclina el motor de la fase 6.
5. **Ordre**: comencem per les fases 1–2 (benefici ràpid, risc baix) o vols atacar abans la
   categorització (fase 3), que és l'arrel del problema però toca el cor del Pas 1?

---
*Redactat el 26-08-2026 a partir de les dades reals i del codi actual. Cap canvi implementat.*
