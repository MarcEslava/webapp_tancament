# Instal·lació — un sol servidor per a tothom

> El panell viu en **una sola màquina Windows** (la que té Excel i la unitat Z:).
> La resta de gent no instal·la absolutament res: obre una adreça al navegador.

---

## Per què no un contenidor Docker

Es va valorar. **No és viable per al Pas 2**: `clsSplit.py` automatitza **Excel de veritat**
(COM: `Excel.Application` → `RefreshAll`) perquè la plantilla `splitFiles/template.xlsx` porta
**taules dinàmiques**. Sense un Excel real recalculant-les, els fitxers sortirien amb les pivots de
la plantilla: plens per dins però amb els resums equivocats — i **sense donar cap error**.

Excel no existeix per a Linux i les imatges de Windows Server no admeten Office. Docker resoldria
Python, pandas i pymssql, que és justament la part que `Panell.bat` **ja** instal·la sola.

Aquest desplegament resol el mateix problema per una altra via: **si ningú ha d'executar-ho al seu
PC, ningú ha d'instal·lar-hi res.**

*(Els passos 1 i 3 sí que serien portables — no toquen Excel. Si algun dia es substitueixen les
taules dinàmiques per fulls calculats amb pandas, el contenidor tornaria a ser una opció.)*

---

## Qui necessita què

| | Què li cal |
|---|---|
| **Usuaris** | Un navegador. Res més. |
| **Màquina servidor** | Windows · Excel instal·lat · unitat **Z:** mapada · accés al SQL de BIFarma · el projecte en una carpeta local. |

---

## Preparar el servidor (un sol cop)

1. **Triar la màquina.** Ha de poder quedar-se encesa **i amb la sessió iniciada** (motiu més avall).
2. **Copiar-hi el projecte** i configurar les credencials: copiar `db_config.example.py` →
   `db_config.py` i omplir-les. Està al `.gitignore`, no es puja mai.
3. **Comprovar que Z: hi és** i que Excel obre bé des d'aquell usuari.
4. **Clic dret a `InstalarServidor.bat` → «Executar com a administrador».** Fa tres coses:
   - prepara l'entorn virtual i les dependències (`Panell.bat setup`);
   - crea la tasca programada **«Panell Tancament»**, que arrenca el panell a cada inici de sessió
     (amb 1 minut de retard, perquè la xarxa i la Z: estiguin llestes);
   - obre el port **5099** al tallafocs, **només per a la xarxa local** (perfils domini/privat).
5. **Arrencar-lo ara sense reiniciar:**
   ```
   schtasks /Run /TN "Panell Tancament"
   ```

En arrencar, la finestra del panell escriu les adreces bones per repartir:

```
==================================================================
  PANELL DE TANCAMENT en marxa
    en aquest equip:  http://127.0.0.1:5099
    des de la xarxa:  http://ECOPC019:5099
                      http://192.168.1.167:5099
==================================================================
```

### ⚠️ Per què la sessió ha de quedar iniciada

La tasca corre **dins la sessió de l'usuari**, no com a servei de sistema. És deliberat:

- **Excel necessita una sessió interactiva.** Automatitzat des d'un servei o com a SYSTEM,
  Excel falla o es penja: és un problema conegut i sense solució neta.
- **La unitat Z: només existeix dins la sessió** de l'usuari que la té mapada.

Per tant: l'equip encès, la sessió iniciada, i la pantalla es pot bloquejar sense problema
(bloquejar ≠ tancar sessió). Si es tanca la sessió, el panell s'atura fins al proper inici.

---

## Ús diari

**Els usuaris** obren `http://NOM-EQUIP:5099` i treballen com sempre: mateixos passos, mateixa
consola en directe, mateixes pestanyes de dades. Val la pena repartir l'enllaç com a favorit.

**Només una execució alhora.** És un límit real, no una precaució: l'automatització d'Excel i els
fitxers de xarxa no toleren dues execucions simultànies. Si algú prova d'executar mentre hi ha
feina en marxa, la consola li diu **què** s'està executant i **des de quina adreça**:

```
Ja hi ha una execució en marxa: index, iniciada des de 192.168.1.42 fa 7 min. Espera que acabi.
```

Tancar la pestanya del navegador **no atura** el procés: continua fins al final al servidor, i el
panell no deixa començar-ne un altre fins que acaba.

Els fitxers de sortida (`output/`, `splitFiles/pasteFiles/`) són **compartits**: qui executa
sobreescriu el que hi hagi, igual que abans. Val la pena avisar abans de reprocessar un mes.

---

## Seguretat

El panell **no té contrasenya**: qualsevol que arribi al port pot llançar execucions que escriuen a
Z: i llegeixen del SQL. Per això la regla del tallafocs es limita a `remoteip=localsubnet` i als
perfils domini/privat — mai al perfil públic.

- **No l'exposis a internet** ni li facis port forwarding.
- Per tornar-lo a deixar només local en aquell equip: variable d'entorn `PANELL_HOST=127.0.0.1`.
- Per canviar de port: `PANELL_PORT=5050` (recorda actualitzar la regla del tallafocs).

---

## Comprovacions i problemes

| Símptoma | Què mirar |
|----------|-----------|
| Els usuaris no hi arriben | El servidor té la sessió iniciada? Prova primer `http://127.0.0.1:5099` **al servidor**: si va, el problema és tallafocs o xarxa. |
| Va per IP però no pel nom | DNS intern; reparteix directament l'adreça IP. |
| No arrenca sol després de reiniciar | `schtasks /Query /TN "Panell Tancament" /V /FO LIST` — mira l'últim resultat i que l'usuari de la tasca sigui el que té la Z:. |
| Els refrescos d'Excel fallen | Algú té Excel obert al servidor? Cal tancar-lo abans del `refresh`. Comprova que no queden `EXCEL.EXE` penjats al Gestor de tasques. |
| «Comprovar servidor i Z:» falla | Es comprova des del **servidor**, no des del PC de l'usuari: la Z: i el SQL han de ser accessibles allà. |

**Aturar el panell:** tancar la finestra de consola al servidor.
**Reprendre'l:** `schtasks /Run /TN "Panell Tancament"`, o `Panell.bat`.
**Desfer-ho tot:** `InstalarServidor.bat elimina` com a administrador (treu la tasca i la regla).

---

## Nota sobre el servidor web

Flask arrenca amb el seu servidor de desenvolupament i avisa que no és per a producció. Per a un ús
intern d'unes poques persones a la xarxa local és suficient. Si algun dia hi ha més gent o es vol
més robustesa, el pas natural és posar-hi **waitress** (`pip install waitress` →
`waitress-serve --host=0.0.0.0 --port=5099 app:app`), que a Windows funciona sense complicacions.

---
*Veure `docs/PROTOCOL_TANCAMENT.md` per al procés de tancament pas a pas.*
