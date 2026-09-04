"use strict";

const $ = (id) => document.getElementById(id);
const consoleEl = $("console");
const badge = $("statusBadge");
const runButtons = () => Array.from(document.querySelectorAll("button.run"));

let activeSource = null;

// Default period = last complete month (today - 1 month).
(function initDefaults() {
  const now = new Date();
  const prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  $("year").value = prev.getFullYear();
  $("month").value = String(prev.getMonth() + 1);
})();

// Populate the Pas-1 and Pas-2 lab dropdowns from Mapa_Acords (fails soft if
// Z: is down). Pas 1 filters the SQL to that lab; Pas 2 splits that lab's
// master -- they are separate steps and each has its own selector.
(async function loadLabs() {
  const sels = [$("laboratori"), $("splitLab")];
  try {
    const d = await (await fetch("/labs")).json();
    if (d.ok && d.labs.length) {
      const esc = (s) => String(s).replace(/[&<>"]/g, (c) =>
        ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
      const html = '<option value="">Tots els laboratoris</option>' +
        d.labs.map((n) => `<option>${esc(n)}</option>`).join("");
      sels.forEach((sel) => (sel.innerHTML = html));
    }
  } catch (e) {
    /* keep the single "Tots els laboratoris" option */
  }
})();

// Download the Pas-1 master for the current period/lab selection. The server
// renames it to PARAFARMACIA_{MM.AAAA}_{YTD|RES}[_LAB].xlsx and answers 404
// (shown in the console) if that run hasn't been generated yet.
$("downloadMaster").addEventListener("click", () => {
  const p = new URLSearchParams({
    period: $("period").value,
    year: $("year").value,
    month: $("month").value,
    laboratori: $("laboratori").value,
  });
  window.location = "/download/master?" + p.toString();
});

// Appends one line to the console, colour-coded by type, and auto-scrolls.
function appendLine(text, type) {
  const span = document.createElement("span");
  span.className = "l-" + (type || "log");
  span.textContent = text + "\n";
  consoleEl.appendChild(span);
  consoleEl.scrollTop = consoleEl.scrollHeight;
}

function setBadge(state, label) {
  badge.className = "badge " + state;
  badge.textContent = label;
}

function setRunning(on) {
  runButtons().forEach((b) => (b.disabled = on));
}

// Gathers the shared period params (year / month / period).
function sharedParams() {
  return {
    year: $("year").value,
    month: $("month").value,
    period: $("period").value,
  };
}

// Builds the query string for a given script.
function paramsFor(script) {
  const p = { script };
  if (script === "index") {
    Object.assign(p, sharedParams(), {
      rappel: $("rappel").value,
      only_para: $("onlyPara").checked ? "1" : "0",
      laboratori: $("laboratori").value,
    });
  } else if (script === "split") {
    const steps = Array.from(document.querySelectorAll("input[name=splitstep]:checked"))
      .map((c) => c.value)
      .join(",");
    // Empty = every lab (the global split); a name = only that lab's file.
    Object.assign(p, sharedParams(), { steps, laboratori: $("splitLab").value });
  } else if (script === "seguiment") {
    p.year = $("year").value;
    p.min_year = $("minYear").value;
  }
  return p;
}

function run(script) {
  if (activeSource) return; // one run at a time

  const params = paramsFor(script);
  if (script === "split" && !params.steps) {
    appendLine("⚠ Selecciona almenys un pas per al split.", "err");
    return;
  }

  setRunning(true);
  setBadge("running", "Executant…");
  appendLine("", "muted");
  appendLine("──────── " + script + " · " + new Date().toLocaleTimeString() + " ────────", "muted");

  const qs = new URLSearchParams(params).toString();
  const es = new EventSource("/run?" + qs);
  activeSource = es;

  es.onmessage = (ev) => {
    let msg;
    try {
      msg = JSON.parse(ev.data);
    } catch {
      return;
    }
    if (msg.type === "done") {
      // Close FIRST so EventSource doesn't auto-reconnect and re-trigger the run.
      es.close();
      activeSource = null;
      setRunning(false);
      if (msg.code === 0) {
        setBadge("ok", "Completat ✓");
        appendLine("✓ Finalitzat correctament (codi 0).", "done");
      } else {
        setBadge("err", "Error");
        appendLine("✗ Finalitzat amb codi " + msg.code + ".", "err");
      }
    } else if (msg.type === "cmd") {
      appendLine(msg.line, "cmd");
    } else if (msg.type === "error") {
      appendLine("ERROR: " + msg.line, "err");
    } else {
      appendLine(msg.line, "log");
    }
  };

  es.onerror = () => {
    // Network drop / server closed unexpectedly. Guard against reconnect loop.
    if (activeSource) {
      es.close();
      activeSource = null;
      setRunning(false);
      setBadge("err", "Connexió perduda");
      appendLine("✗ Connexió amb el servidor interrompuda.", "err");
    }
  };
}

runButtons().forEach((btn) =>
  btn.addEventListener("click", () => run(btn.dataset.script))
);

$("clearBtn").addEventListener("click", () => {
  consoleEl.textContent = "";
});

// ---------------- Tabs ----------------
const tabs = Array.from(document.querySelectorAll(".tab"));
const panels = {
  run: $("tab-run"), detector: $("tab-detector"), analitica: $("tab-analitica"),
};

tabs.forEach((t) =>
  t.addEventListener("click", () => {
    const name = t.dataset.tab;
    tabs.forEach((x) => x.classList.toggle("active", x === t));
    Object.entries(panels).forEach(([k, el]) => { if (el) el.classList.toggle("active", k === name); });
    if (name === "detector") detTable.loadOnce();
    if (name === "analitica") mbInit();
  })
);

// ------------- Analítica: Metabase embegut (com l'Explorar del cockpit) -------------
// Pestanya dins de la SPA: dropdown de dashboards del Cierre + iframe. Canviar de
// pestanya és "tornar enrere" (no és una pàgina a part).
let MB_INIT = false, MB_BASE = "";
function mbFill(setFirst) {
  const sel = $("mbSelect"), fr = $("mbFrame"), cur = sel.value;
  return fetch("/api/mb/cierre").then((r) => r.json()).then((d) => {
    if (!d.ok) { sel.innerHTML = '<option>Error carregant</option>'; return; }
    MB_BASE = d.mb_base;
    const c = $("mbCreate"); if (c) c.href = MB_BASE + (d.cierre_collection ? "/collection/" + d.cierre_collection : "/");
    if (!d.sharing) { sel.innerHTML = '<option>Activa el «public sharing» a Metabase</option>'; return; }
    if (!d.dashboards.length) { sel.innerHTML = '<option>Cap dashboard</option>'; return; }
    // Agrupat per àrea/persona (<optgroup>), com l'Explorar del cockpit de pricing.
    const emb = "#theme=transparent&bordered=false&titled=false";
    const byG = {};
    d.dashboards.forEach((x) => { (byG[x.group] = byG[x.group] || []).push(x); });
    sel.innerHTML = Object.keys(byG).map((g) =>
      `<optgroup label="${g}">` + byG[g].map((x) =>
        `<option value="${MB_BASE}/public/dashboard/${x.uuid}${emb}">${x.name}</option>`).join("")
      + `</optgroup>`).join("");
    const keep = [...sel.options].some((o) => o.value === cur);
    if (keep) { sel.value = cur; }
    else if (setFirst) {
      const def = [...sel.options].find((o) => o.textContent === "Cierre · Insights") || sel.options[0];
      sel.value = def.value; fr.src = def.value;
    }
  }).catch(() => { sel.innerHTML = '<option>Error de xarxa</option>'; });
}
function mbInit() {
  if (MB_INIT) return; MB_INIT = true;
  $("mbSelect").onchange = () => { $("mbFrame").src = $("mbSelect").value; };
  $("mbRefresh").onclick = () => mbFill(false);
  mbFill(true);
}

// ------------- Reusable CSV table -------------
// Renders a server CSV into a table with client-side search + a lab dropdown.
// The lab selection also drives the (server-side) filtered download.
// numCols: right-aligned columns. pctCol: a % column to colour-code.
function csvTable(o) {
  const tableEl = $(o.table);
  const metaEl = $(o.meta);
  const searchEl = $(o.search);
  const labEl = o.lab ? $(o.lab) : null;
  const numCols = o.numCols || [];
  const pctCol = o.pctCol || null;
  // Group/filter by the real classification; plain "Laboratorio" is only a
  // fallback for a CSV written by an older run.
  const labCols = o.labCols || ["Laboratorio Categorizado", "Laboratorio"];
  const labIndex = () => {
    for (const c of labCols) {
      const i = cols.indexOf(c);
      if (i >= 0) return i;
    }
    return -1;
  };
  const MAX = 1500; // cap DOM rows so a huge CSV doesn't freeze the page
  let cols = [], rows = [], loaded = false, updated = "?";

  const esc = (s) =>
    String(s).replace(/[&<>"]/g, (c) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

  function render(rws) {
    tableEl.querySelector("thead").innerHTML =
      "<tr>" + cols.map((c) => `<th>${esc(c)}</th>`).join("") + "</tr>";
    const shown = rws.slice(0, MAX);
    tableEl.querySelector("tbody").innerHTML = shown.map((r) =>
      "<tr>" + r.map((v, i) => {
        let cls = numCols.includes(cols[i]) ? "numcell" : "";
        if (pctCol && cols[i] === pctCol) {
          const n = parseFloat(String(v).replace(",", "."));
          if (n >= 10) cls += " bad";
          else if (n >= 3) cls += " warn-cell";
        }
        cls = cls.trim();
        // Long values are clipped with an ellipsis by the CSS, so keep the full
        // text reachable on hover instead of losing it.
        const t = String(v).length > 40 ? ` title="${esc(v)}"` : "";
        return `<td${cls ? ` class="${cls}"` : ""}${t}>${esc(v)}</td>`;
      }).join("") + "</tr>"
    ).join("");
    const capped = rws.length > MAX;
    metaEl.textContent = `Actualitzat ${updated} · ${rws.length} files` +
      (capped ? ` (mostrant ${MAX})` : "");
  }

  function selected() {
    const q = searchEl.value.trim().toLowerCase();
    const lab = labEl ? labEl.value : "";
    const labI = labIndex();
    return rows.filter((r) => {
      if (lab && labI >= 0 && r[labI] !== lab) return false;
      if (q && !r.some((v) => String(v).toLowerCase().includes(q))) return false;
      return true;
    });
  }
  const apply = () => render(selected());

  function populateLabs() {
    if (!labEl) return;
    const labI = labIndex();
    if (labI < 0) return;
    const cur = labEl.value;
    const labs = Array.from(new Set(rows.map((r) => r[labI]).filter(Boolean))).sort();
    labEl.innerHTML = '<option value="">Tots els labs</option>' +
      labs.map((l) => `<option${l === cur ? " selected" : ""}>${esc(l)}</option>`).join("");
  }

  async function load() {
    metaEl.textContent = "Carregant…";
    try {
      const d = await (await fetch(o.url)).json();
      if (!d.exists) {
        loaded = false; cols = []; rows = [];
        tableEl.querySelector("thead").innerHTML = "";
        tableEl.querySelector("tbody").innerHTML = "";
        metaEl.textContent = o.emptyMsg ||
          "Encara no hi ha dades — executa el pas 1 (Generar mestre).";
        return;
      }
      cols = d.columns; rows = d.rows; updated = d.updated; loaded = true;
      populateLabs();
      apply();
    } catch (e) {
      metaEl.textContent = "Error carregant: " + e;
    }
  }

  function downloadUrl() {
    const p = new URLSearchParams();
    if (labEl && labEl.value) p.set("lab", labEl.value);
    const qs = p.toString();
    return `/download/${o.key}` + (qs ? "?" + qs : "");
  }

  searchEl.addEventListener("input", () => { if (loaded) apply(); });
  if (labEl) labEl.addEventListener("change", () => { if (loaded) apply(); });
  if (o.download) $(o.download).addEventListener("click", () => { window.location = downloadUrl(); });

  return { load, loadOnce: () => { if (!loaded) load(); } };
}

// Refs the BIFarma grid has and the SQL master doesn't (4th tab). No lab
// dropdown: the rows carry the lab of the master's NEIGHBOURS, not their own --
// an unknown reference has no lab in the DB by definition.
const detTable = csvTable({
  url: "/detector", key: "detector",
  table: "detTable", meta: "detMeta", search: "detSearch",
  download: "detDownload",
  numCols: ["Compra (€) actual", "Compra (€) anterior", "Files a l'export",
            "Refs del mateix nom al mestre"],
  emptyMsg: "Encara no hi ha informe — puja un export de BIFarma aquí a dalt.",
});

$("detRefresh").addEventListener("click", detTable.load);

// Uploads the export and rebuilds the report. The analysis reads the whole
// product master (~86k names), so it takes a few seconds: disable the button
// and say so instead of leaving the page looking idle.
$("detRun").addEventListener("click", async () => {
  const input = $("detFile");
  const btn = $("detRun");
  const meta = $("detMeta");
  const file = input.files && input.files[0];
  if (!file) {
    meta.textContent = "Tria un fitxer .xlsx exportat de BIFarma.";
    return;
  }
  const body = new FormData();
  body.append("export", file);
  btn.disabled = true;
  meta.textContent = `Analitzant ${file.name}…`;
  try {
    const d = await (await fetch("/detector", { method: "POST", body })).json();
    if (!d.ok) {
      meta.textContent = "Error: " + (d.error || "no s'ha pogut analitzar el fitxer.");
      return;
    }
    await detTable.load();
    const r = d.resum || {};
    const anys = r.any_actual
      ? ` · compra ${r.any_actual}${r.any_anterior ? " vs " + r.any_anterior : ""}`
      : "";
    meta.textContent =
      `${r.fitxer} · ${r.productes_export} productes a l'export, ` +
      `${r.desconeguts} sense coincidencia al mestre ` +
      `(${r.sense_cap_ve} sense cap vei) · ${r.eur_desconeguts} € de compra${anys}`;
  } catch (e) {
    meta.textContent = "Error de pujada: " + e;
  } finally {
    btn.disabled = false;
  }
});

// ------------- Health check (SQL + Z:) -------------
function setPill(id, ok, text) {
  const el = $(id);
  el.textContent = text;
  el.className = "pill" + (ok === true ? " ok" : ok === false ? " err" : " pending");
}

$("checkBtn").addEventListener("click", async () => {
  const btn = $("checkBtn");
  btn.disabled = true;
  setPill("checkSql", null, "Comprovant…");
  setPill("checkZ", null, "Comprovant…");
  $("checkDetail").textContent = "";
  try {
    const res = await fetch("/check");
    const d = await res.json();
    setPill("checkSql", d.sql.ok, d.sql.ok ? "Operatiu" : "Error");
    setPill("checkZ", d.z.ok, d.z.ok ? "Operatiu" : "Error");
    $("checkDetail").textContent = "SQL — " + d.sql.detail + "\nZ:  — " + d.z.detail;
  } catch (e) {
    setPill("checkSql", false, "Desconegut");
    setPill("checkZ", false, "Desconegut");
    $("checkDetail").textContent = "Error de comprovació: " + e;
  } finally {
    btn.disabled = false;
  }
});
