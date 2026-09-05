# Monthly Close Control Panel

A **web control panel that orchestrates a monthly data-close pipeline** for pharmaceutical lab data. It replaces a fragile, manual process (running scripts by hand, editing gigabyte-sized CSVs, coordinating who runs what) with a single browser-based panel that the whole team can use without installing anything.

> Built to solve a real operational problem: a monthly data pipeline that cleaned, split, and consolidated lab sell-in / sell-out data spread across hundreds of Excel files and multi-gigabyte CSVs.

## The problem

The monthly close involved several heavy steps run by hand:

- Cleaning raw data exported from the source BI system
- Splitting per-lab data into separate files to send to each provider
- Rebuilding a consolidated tracking dataset from hundreds of Excel files

The artefacts had grown painful: a **1.2 GB accumulated CSV (5.5M rows)** re-read and rewritten in full every close over a network share, a tracking file rebuilt from hundreds of Excels, and a master file at the limit of what Excel can hold. The "source of truth" lived inside an Excel file the process generated itself, causing circular dependencies and data that didn't reconcile.

## What this does

**A single-machine web panel** (Flask + HTML/JS) that:

- Runs each pipeline step (`index.py`, `clsSplit.py`, `clsSeguiment.py`) as a subprocess and **streams its output live to the browser** via Server-Sent Events (SSE), so anyone can watch progress in real time.
- Deploys on **one Windows machine** (the one with Excel and the network drive); everyone else just opens a browser and installs nothing.
- **Serialises heavy jobs with a lock**: Excel COM automation and the network Excel files don't tolerate concurrent runs, so a second request while busy is told *what* is running and *since when*, instead of silently colliding.
- Includes a **detector** for references present in the source web grid but missing from the SQL replica, so the pipeline never silently drops rows.

## Architecture highlights

- **Orchestration over a web panel** — instead of a monolith, the panel runs each step as an isolated subprocess and streams stdout. Steps stay independent and debuggable.
- **ClickHouse as the analytics/serving layer** — the monthly facts land in a partitioned `MergeTree` table (re-running a month is idempotent), with product dimensions in a `ReplacingMergeTree` keyed by product code. Written over HTTP, so the panel machine needs no driver.
- **Excel COM automation** — one pipeline step drives a real, interactive Excel session on the host machine (documented deployment constraints in `docs/`).
- **Secrets and business data never in the repo** — credentials live in a git-ignored `db_config.py` (a template is provided); all Excel/CSV business data is git-ignored.

## Data model migration plan

`docs/PLA_DADES.md` documents a phased plan to move the "truth" out of self-generated Excel files and into proper master tables, normalising a 1.2 GB denormalised CSV down to a fraction by not repeating full context on every row. Phase 0 (parallel writes) is implemented; later phases are a documented proposal. This shows the reasoning behind the data model, not just the code.

## Stack

`Python` · `Flask` · `pandas` · `ClickHouse` · `SQL (MSSQL / SQLAlchemy)` · `Excel COM automation (pywin32)` · `Server-Sent Events` · `HTML/JS`

## Layout

```
app.py            # Flask panel: runs steps as subprocesses, streams stdout via SSE
index.py          # Step 1: clean raw sell-in / sell-out data
clsSplit.py       # Step 2: split per-lab data into provider files (Excel COM)
clsSeguiment.py   # Step 3: rebuild the consolidated tracking dataset
detector.py       # Detect references missing from the SQL replica
store.py          # Data model / master column split (phase 0)
store_ch.py       # ClickHouse store: partitioned facts + product dimension
db_config.example.py  # Template for local secrets (real one is git-ignored)
docs/             # Deployment guide, data model migration plan, close protocol
```

## Notes

This is a portfolio-oriented writeup of a real operational tool. All business data, credentials, and provider/product details are excluded from the repository; only the code and its documentation are shared.
