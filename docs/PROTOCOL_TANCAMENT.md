# PROTOCOL DE TANCAMENT — Seguiment mensual laboratoris

> Procés per generar el seguiment mensual/YTD dels laboratoris del grup ECO, des de la
> consulta directa al **SQL de BIFarma** fins als fitxers per laboratori i l'històric acumulat.
>
> Tot s'executa des del **panell web local** (`Panell.bat` → `app.py` → http://127.0.0.1:5099),
> que llança els tres scripts (`index.py`, `clsSplit.py`, `clsSeguiment.py`) i mostra la consola
> en directe. Els scripts també es poden executar per línia de comandes (veure taula d'equivalents).
>
> ⚠️ Els passos marcats amb **[MANUAL]** no els fa cap script.
>
> 📄 Hi ha una versió il·lustrada d'aquest document, amb diagrama, a `docs/PROTOCOL_TANCAMENT.html`.

---

## 0. Visió general (ordre del procés)

```
[MANUAL] Panell.bat → http://127.0.0.1:5099 → "Comprovar servidor i Z:"
        │
        ▼
PAS 1 · index.py (clsBiFarmaEco)          SQL BIFarma → neteja → creuament amb Z:
        │                                 checks especialitats: si fallen → recategoritzar i repetir
        ▼
output/df_bifarma_output[period].xlsx     (+ product_conditions.csv · lost_money.csv)
        │
        ▼
[MANUAL] Parafarmacia {MM.AAAA}[ YTD].xlsx     full "SI Acord book" + taules dinàmiques (a Z:)
        │
        ▼
PAS 2 · clsSplit.py (clsSplit)            split · fee · refresh · accumulate · cleanup
        │
        ▼
splitFiles/pasteFiles/PARAFARMACIA {MM.AAAA}[ YTD] - {LAB}.xlsx
        │
        ▼
[MANUAL] 102. Informes i Rappel\{any}\{MM.AAAA}\mes\     ← col·locar/enviar els fitxers
        │
        ▼
PAS 3 · clsSeguiment.py (clsSeguiment)    → 102. Informes i Rappel\Seguiment_labs.csv
```

**Variables de període (camp «Temporalitat» del panell):**
- Mensual (`period = ""`) → dades del mes.
- YTD (`period = "YTD"`) → acumulat de l'any (gener → mes processat).

Per defecte el període és **l'últim mes tancat** (avui − 1 mes); es pot canviar amb els camps
*Any* i *Mes*, que fan servir els passos 1 i 2.

---

## 1. [MANUAL] Obrir el panell i comprovar l'entorn

El panell viu en **una sola màquina** (la que té Excel i la unitat Z:) i tothom hi entra pel
navegador — **no cal instal·lar res al teu PC**.

1. Obrir **`http://NOM-EQUIP:5099`** (demana l'adreça a qui porti el tancament i desa-la a favorits).
2. Prémer **«Comprovar servidor i Z:»**: valida la connexió SQL i l'accés real a la carpeta rappel.
   Es comprova **des del servidor**, que és qui hi ha de tenir accés.

> Muntar aquesta màquina servidor és feina d'un sol cop: veure **`docs/INSTALLACIO_SERVIDOR.md`**
> (inclou per què no es pot fer amb Docker).

**Si has d'executar-ho al teu propi equip** (perquè el servidor és fora de servei, o per provar):
cal Windows amb Excel i la Z: mapada, i executar **`Panell.bat`**, que comprova que hi ha Python
(si no, l'instal·la amb winget), prepara l'entorn virtual la primera vegada i obre el navegador.
- El venv es crea a `%LOCALAPPDATA%\WebAppTancament\venv` expressament: la carpeta del projecte és
  dins d'OneDrive i no volem sincronitzar-hi milers de fitxers.
- La primera execució (o quan canvia `requirements.txt`) triga uns minuts instal·lant; les següents
  arrenquen directes. `Panell.bat setup` només prepara l'entorn sense arrencar.
- Cal `db_config.py` amb les credencials del SQL (copiar de `db_config.example.py`; està al
  `.gitignore`, no es puja mai).

> El panell només permet **una execució alhora** (l'automatització COM d'Excel i els fitxers de
> xarxa no toleren concurrència). Compartint servidor això passa de debò: si algú altre està
> executant, la consola et diu què hi ha en marxa i des de quina adreça. Si tanques la pestanya a
> mig procés, el subprocés continua al servidor i no en deixa començar un altre fins que acaba.

---

## 2. PAS 1 · `index.py` — Generar mestre (clsBiFarmaEco)

**Ja no cal exportar res a mà de BIFarma Eco.** El script consulta directament la base de dades:
any actual + any anterior, només farmàcies del grup **ECO**, i després es queda amb els labs amb
acord segons `Mapa_Acords.xlsx` (coincidència **per codi BIF o per nom**).

En mode **YTD** la consulta es fa **mes a mes** (gener → mes processat) amb una pausa entremig i es
recombina en pandas: una agregació YTD sencera en una sola query saturava el servidor.

**Fitxers de suport que llegeix de Z::**

| Fitxer | Què n'agafa |
|--------|-------------|
| `Mapa_Acords.xlsx` (full `Mapa`) | labs amb acord · noms i codis BIF · `Laboratorio Categorizado` |
| `BBDD_Book INTERN.xlsx` (full segons *Rappel*) | `NIF`, `BOOK`, `SOLAR` |
| `Parafarmacia MM.AAAA YTD.xlsx` (veure sota quin) | `PVL`, `IVA`, `MARCA`, `SUBMARCA` |

**Quin YTD agafa** (`_find_ytd_source`), per aquest ordre:
1. El **YTD del mes processat**, si ja existeix a l'arrel de la carpeta rappel — és el cas del
   tancament mensual, i així conserva la feina manual de MARCA d'aquell mes.
2. Si no existeix, el **YTD anterior més recent**. És el cas de generar el YTD mateix: el 07.2026 YTD
   no es pot llegir a si mateix, així que agafa el 06.2026 YTD. Busca a l'arrel **i a les carpetes
   d'any** (`...\2026\Parafarmacia 06.2026 YTD.xlsx`), on s'arxiven els mesos tancats, i es queda
   amb el mes més alt **estrictament anterior** al processat (mes del nom del fitxer, no data de
   modificació).
3. Si no en troba cap, s'atura amb un missatge clar.

Els productes que no hi són surten amb **MARCA «New!»**.

**Opcions del panell:**

| Opció | Efecte |
|-------|--------|
| *Rappel* | `BIFARMA` o `bifarma` (con BAJAS) — full del `BBDD_Book INTERN`. |
| *Temporalitat* | Mensual o YTD. |
| *Només parafarmàcia* | Exclou l'especialitat ja a la SQL (`ESPECIALIDAD` / `ESPEC. CARAS`). |
| *Laboratori* | Informe d'un sol lab (veure sota). El split d'aquell lab es demana al **Pas 2**. |

**Checks d'especialitats** — si en falta cap, el script **s'atura**: cal recategoritzar a BIFarma i
tornar a executar.

| Comprovació | Què cal categoritzar |
|-------------|----------------------|
| **Amox RJ** | Grupo Producto: ESPECIALIDAD · SubGrupo: EFG · Producto: amoxi · Laboratorio: **Reig Jofre** |
| **RX Almirall** | Productes ESERTIA / PARAPRES → Laboratorio: **Almirall** |
| **MV Menarini** | Producto: **MENAVEN** → Especialidad Menarini |

Amb *Només parafarmàcia* o amb un laboratori concret **els checks se salten** (no hi hauria
especialitat a trobar).

**Comprovar a la consola:** `ok RJ`, `ok Almirall`, `ok Menarini` i **«No duplicates found»**.

**Sortida:** `output/df_bifarma_output[period].xlsx`, full **`SI Acord book`**
(botó **«⬇ Descarregar resultat»** del panell), més els dos CSV de les pestanyes de dades.

### Informe d'un sol laboratori

Amb *Laboratori* seleccionat (llista de la columna `Laboratori` del `Mapa_Acords`):

- La SQL es filtra als **noms i codis BIF** d'aquell lab.
- L'informe va a un **fitxer a part** (`output/df_bifarma_output[period]_{LAB}.xlsx`) i **no toca els
  CSV globals**.
- **El Pas 1 acaba aquí.** No encadena res: si en vols el fitxer de split, ho demanes explícitament
  al **Pas 2** amb el mateix laboratori seleccionat (veure *Split d'un sol laboratori*).

> **Cada pas s'executa per separat.** Cap pas del panell engega un altre pas pel seu compte: si no
> marques una casella, allò no s'executa.

---

## 3. [MANUAL] Muntar el fitxer mestre `Parafarmacia {MM.AAAA}.xlsx`

Portar les dades del mestre generat a `Parafarmacia {MM.AAAA}[ YTD].xlsx` dins de
`Z:\...\102. Informes i Rappel\`, full **`SI Acord book`**, **mantenint les taules dinàmiques**.
És el fitxer que llegeix el Pas 2 — i el YTD també alimenta els PVL/IVA/MARCA del Pas 1 següent.

Revisar les **marques noves («New!»)** i completar `MARCA`/`SUBMARCA` a mà: la pestanya
**Diners perduts** del panell diu a quins labs cal prioritzar.

---

## 4. PAS 2 · `clsSplit.py` — repartir per laboratori (clsSplit)

La llista de labs surt de **`a.SEGUIMENT LABS {any}.xlsx`** (columna A dels fulls `Parafarmacia` i
`EFG`) més els `EXTRA_LABS` definits al codi. Cada fitxer es crea copiant la **plantilla única**
`splitFiles/template.xlsx` (fulls `Acuerdo book` + taules dinàmiques, **sense** full `Fee`).

**Passos seleccionables al panell:**

| Pas | Què fa |
|-----|--------|
| `split` | Copia la plantilla per lab i hi escriu el full `Acuerdo book` (fora les files amb `MARCA`/`SUBMARCA` = `NO`). |
| `fee` | Full **`Fee`** per farmàcia (veure sota). |
| `refresh` | Obre cada fitxer amb Excel (COM), refresca fórmules, amaga `Acuerdo book` i desa. |
| `accumulate` | Afegeix el mes a `accumulated_labs.csv` (substitueix el mes si ja hi era). |
| `cleanup` | Esborra de `pasteFiles` els fitxers de més de 30 dies. |

Els passos són **independents**: pots executar-los d'un en un i en execucions separades. Si no
n'arriba cap (`--steps` buit), el script no fa res i acaba amb un avís — mai no fa el split «per
defecte».

**Comprovar a la consola.** El pas `split` imprimeix les files de cada lab i, al final, dos avisos:

- **informes BUITS** — labs de la llista del SEGUIMENT sense cap fila. Pot ser que no tinguin compres
  aquest mes, o que el nom no coincideixi amb la columna `Laboratori` del `Mapa_Acords`.
- **categories amb dades que no van a cap informe** — valors de `Laboratorio Categorizado` del mestre
  que cap lab de la llista del SEGUIMENT reclama, així que no es genera cap fitxer per a ells
  (`#N/D` inclòs: és el que veus a la pestanya *Diners perduts*).

El creuament es fa per `Laboratorio Categorizado`, que el Pas 1 ja omple amb l'etiqueta `Laboratori`
del `Mapa_Acords`. **No** es fa cap segona consulta al `Mapa_Acords` des del Pas 2.

⚠️ **Tancar tots els Excel abans del `refresh`.** Si Excel mor (*RPC server unavailable*), el script
en reinicia un de nou i reintenta el fitxer una vegada; els que fallin es llisten al final.

**Sortida:** `splitFiles/pasteFiles/PARAFARMACIA {MM.AAAA}[ YTD] - {LAB}.xlsx`

### Split d'un sol laboratori

Amb el desplegable **Laboratori** del Pas 2 (o `--laboratori "NOM LAB"`):

- Es llegeix **el mateix mestre del mes que el split complet** (`Parafarmacia {MM.AAAA}[ YTD].xlsx`
  de Z:) i s'hi busca el lab per `Laboratorio Categorizado`. És exactament el mateix creuament, només
  que escriu un sol llibre. **No depèn del Pas 1**: no cal haver-lo executat per a aquell lab.
- Escriu exactament un fitxer: `PARAFARMACIA {MM.AAAA}[ YTD] - {LAB}.xlsx`. Les marques
  reclassificades fora del lab (p. ex. **SENSILIS** fora de DERMOFARM) ja porten un altre
  `Laboratorio Categorizado` al mestre, així que surten a l'execució del seu propi lab.
- `fee` i `refresh` només toquen aquest fitxer, no la resta de `pasteFiles`. Es poden executar sense
  el pas `split` si el fitxer del mes ja existeix.
- Si el llibre del lab surt **buit**, el pas acaba amb error i no refresca: no té sentit donar per bo
  ni enviar un fitxer sense dades. Comprova el nom contra la columna `Laboratori` del `Mapa_Acords`,
  i que el mestre del mes ja estigui regenerat si acabes de canviar la categorització.
- `accumulate` i `cleanup` són **globals**: amb un laboratori seleccionat s'ignoren (amb avís).

> El mestre és la font de veritat del Pas 2. Si has canviat el `Mapa_Acords` o la categorització del
> Pas 1, el split no ho veurà fins que el mestre `Parafarmacia {MM.AAAA}[ YTD].xlsx` es torni a
> muntar (Pas 1 + pas manual 3).

---

## 5. [MANUAL] Distribució

Col·locar els fitxers generats a `102. Informes i Rappel\{any}\{MM.AAAA}\mes\` i/o enviar-los a cada
laboratori. Aquesta estructura és la que després llegeix el Pas 3.

---

## 6. PAS 3 · `clsSeguiment.py` — històric acumulat (clsSeguiment)

Recorre les carpetes `{any ≥ any mínim}\{MM.AAAA}\mes\`, llegeix el full `Acuerdo book` de cada
`PARAFARMACIA *.xlsx` (**salta** els YTD i els `(NC)`), hi afegeix les condicions de fee del
SEGUIMENT de la campanya i ho desa tot a un CSV històric (separador `;`).

- **Any mínim:** per defecte 2024. **Any:** campanya del fitxer `a.SEGUIMENT LABS {any}.xlsx`.
- **Sortida:** `Z:\...\102. Informes i Rappel\Seguiment_labs.csv`

---

## Pestanyes de dades del panell

| Pestanya | Font | Què mostra |
|----------|------|------------|
| **Condicions de producte** | `output/product_conditions.csv` | `PVL` / `IVA` / `MARCA` / `SUBMARCA` per `Cod Unif`. S'actualitza (upsert) a cada execució del Pas 1: l'última execució mana per producte. Cerca, filtre per lab i descàrrega CSV filtrable. |
| **Diners perduts** | `output/lost_money.csv` | Per lab: segons si el fee és sobre compra **neta** (PUC → camp clau `IVA`) o **bruta** (PVL → camp clau `PVL`), % del **SI (€)** total amb el camp clau a 0 / N/D / producte nou — diners sobre els quals **no es pot calcular fee**. Ponderat pel SI, així els productes grans pesen més. |

### Per quina columna de laboratori s'agrupa

Les dues pestanyes fan servir **`Laboratorio Categorizado`** — l'etiqueta `Laboratori` del `Mapa_Acords`
que el Pas 1 assigna a cada fila. És la classificació real: és per aquí que s'agrupa, que es calculen
els fees i que el Pas 2 reparteix els fitxers. El nom cru de BIFarma (`Laboratorio`) es conserva com a
segona columna, perquè és el que cal donar d'alta a la columna `BIF` del `Mapa_Acords` quan falta.

Les files **sense categoritzar** surten com a `#N/D`:

- A *Condicions de producte*, a la columna `Laboratorio Categorizado`.
- A *Diners perduts*, agrupades sota `#N/D` amb el nom cru a `Laboratorio (BIF)`. Tenen el 100% del
  SI perdut per definició: sense categoria no van a cap lab, així que no hi ha fee possible.

Abans s'hi feia servir el nom cru quan faltava la categoria, i això partia un mateix lab en diverses
files (`ABBOTT` i `Abbott` per separat) i llistava distribuïdors i raons socials (`CARBO VITAL S.L.`,
`HALEON SPAIN S.A.`) com si fossin laboratoris.

Cap de les dues es regenera en execucions d'un sol laboratori.

---

## Equivalents per línia de comandes

```bash
python index.py --rappel BIFARMA --period YTD --year 2026 --month 6 [--only-para] [--laboratori "NOM LAB"]
python clsSplit.py --period "" --year 2026 --month 6 --steps split,refresh [--laboratori "NOM LAB"]
python clsSeguiment.py --min-year 2024 --year 2026
```

---

## Punts crítics i errors habituals

- **Tancar tots els Excel abans del `refresh`.** El refresc es fa per COM: amb llibres oberts pot
  fallar o desestabilitzar l'execució.
- **Especialitats mal categoritzades aturen el Pas 1.** Si falten amoxi / ESERTIA·PARAPRES / MENAVEN,
  recategoritzar a BIFarma i tornar a executar. (Amb *Només parafarmàcia* o un sol lab no s'apliquen.)
- **`MARCA` / `SUBMARCA` és feina manual.** Ve del `Parafarmacia YTD` del mes processat; els productes
  nous surten com «New!» i sense PVL/IVA no computen fee. Fer servir **Diners perduts** per prioritzar,
  i mantenir els mateixos criteris que el mes anterior perquè l'històric sigui comparable.
- **Marques que no quadren (Zoho ↔ mes passat)** → no hi ha regla fixa: es revisa cada discrepància a
  mà i es decideix cas per cas.
- **Ordre YTD → mensual.** El Pas 1 mensual agafa el `Parafarmacia {MM.AAAA} YTD.xlsx` del mateix mes
  si ja hi és; generant el YTD, cau automàticament al del mes anterior. Igualment convé fer **primer el
  YTD i muntar-lo**, perquè el mensual parteixi de les marques ja revisades d'aquell mes.
- **Prerequisits de Z::** han d'existir `Mapa_Acords.xlsx`, `BBDD_Book INTERN.xlsx`,
  `a.SEGUIMENT LABS {any}.xlsx` i almenys un `Parafarmacia MM.AAAA YTD.xlsx` (del mes processat o
  d'un mes anterior).
  El botó «Comprovar servidor i Z:» valida l'accés a la carpeta.
- **`FileNotFoundError` al Pas 2** → falta `splitFiles/template.xlsx`.
- **Els valors negatius es tallen a 0.** Abonaments i retorns es converteixen en 0 — tenir-ho present
  si els totals no quadren amb BIFarma.
- **Duplicats al Pas 1** → la consola ha de dir «No duplicates found».
- **`db_config.py` mai al git.** Conté les credencials del SQL i està al `.gitignore`.

---
*Actualitzat el 26-08-2026 a partir del codi (`app.py` · `index.py` · `clsSplit.py` · `clsSeguiment.py`).
El flux antic amb exports manuals de BIFarma Eco ja no aplica: el Pas 1 llegeix directament del SQL.*
