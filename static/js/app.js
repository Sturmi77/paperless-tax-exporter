/* ─── Paperless Tax Exporter · Frontend ─────────────────────────────── */

const $ = id => document.getElementById(id);

// State
let allTags      = [];
let selectedTags = new Set();
let polling      = null;

// ─── Init ──────────────────────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", () => {
  const fy = $("footer-year");
  if (fy) fy.textContent = new Date().getFullYear();
  buildYearButtons();
  loadTags();
  checkConnection();
  updateOutputPath();

  $("date-from").addEventListener("change", () => { clearActiveYear(); updateInfo(); });
  $("date-to").addEventListener("change",   () => { clearActiveYear(); updateInfo(); });
  $("start-btn").addEventListener("click",  startExport);
  $("new-export-btn").addEventListener("click", resetUI);
});

// ─── Verbindungscheck ──────────────────────────────────────────────────
async function checkConnection() {
  const badge = $("connection-status");
  try {
    const res  = await fetch("/api/health");
    const data = await res.json();
    if (!data.token_configured) {
      badge.textContent = "Token nicht konfiguriert";
      badge.className   = "badge badge-error";
      badge.title       = "PAPERLESS_TOKEN Umgebungsvariable fehlt – bitte .env auf dem NAS anlegen";
    } else if (data.paperless_reachable === false) {
      badge.textContent = "Paperless nicht erreichbar";
      badge.className   = "badge badge-error";
      badge.title       = data.error || "";
    } else {
      badge.textContent = "Paperless verbunden";
      badge.className   = "badge badge-ok";
    }
  } catch {
    badge.textContent = "Paperless nicht erreichbar";
    badge.className   = "badge badge-error";
  }
}

// ─── Output-Pfad anzeigen ──────────────────────────────────────────────────
async function updateOutputPath() {
  try {
    const res  = await fetch("/api/config");
    const data = await res.json();
    if (data.output_dir) {
      $("output-path-text").textContent = data.output_dir + "/{Jahr}/";
    }
  } catch {
    $("output-path-text").textContent = "(konfigurierter Ausgabepfad auf dem NAS)";
  }
}

// ─── Schnellauswahl Kalenderjahre ─────────────────────────────────────
function buildYearButtons() {
  const container = $("quick-years");
  const thisYear  = new Date().getFullYear();
  // Aktuelle + 3 Vorjahre
  for (let y = thisYear; y >= thisYear - 3; y--) {
    const btn = document.createElement("button");
    btn.className   = "year-btn";
    btn.textContent = y;
    btn.dataset.year = y;
    btn.addEventListener("click", () => selectYear(y, btn));
    container.appendChild(btn);
  }
}

function selectYear(year, btn) {
  document.querySelectorAll(".year-btn").forEach(b => b.classList.remove("active"));
  btn.classList.add("active");
  $("date-from").value = `${year}-01-01`;
  $("date-to").value   = `${year}-12-31`;
  updateInfo();
}

function clearActiveYear() {
  document.querySelectorAll(".year-btn").forEach(b => b.classList.remove("active"));
}

// ─── Tags laden ────────────────────────────────────────────────────────
async function loadTags() {
  try {
    const healthRes  = await fetch("/api/health");
    const healthData = await healthRes.json();

    if (!healthData.token_configured) {
      throw new Error("PAPERLESS_TOKEN nicht gesetzt. Bitte .env auf dem NAS anlegen.");
    }

    const res  = await fetch("/api/tags");
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    if (data.error) throw new Error(data.error);

    allTags = data.tags || [];
    renderTags();
    $("tag-loading").classList.add("hidden");
    $("tag-list").classList.remove("hidden");
    $("start-btn").disabled = false;
    updateInfo();
  } catch (err) {
    $("tag-loading").classList.add("hidden");
    $("tag-error").textContent = `Fehler beim Laden der Tags: ${err.message}`;
    $("tag-error").classList.remove("hidden");
    $("start-btn").disabled = false;
  }
}

function renderTags() {
  const container = $("tag-list");
  container.innerHTML = "";

  if (allTags.length === 0) {
    container.innerHTML = '<span style="color:#a0aec0;font-size:.88rem;">Keine Tags in Paperless gefunden.</span>';
    return;
  }

  const sorted = [...allTags].sort((a, b) => a.name.localeCompare(b.name, "de"));

  const select = document.createElement("select");
  select.id      = "tag-select";
  select.multiple = true;
  select.size    = Math.min(sorted.length, 8);
  select.style.cssText = "width:100%;border:1.5px solid #c5d2d4;border-radius:6px;padding:4px;font-size:.93rem;font-family:inherit;background:#fff;color:#444;outline:none;";

  sorted.forEach(tag => {
    const opt = document.createElement("option");
    opt.value       = tag.id;
    opt.textContent = tag.name;
    select.appendChild(opt);
  });

  select.addEventListener("change", () => {
    selectedTags.clear();
    Array.from(select.selectedOptions).forEach(o => selectedTags.add(Number(o.value)));
    updateInfo();
  });

  const hint = document.createElement("div");
  hint.style.cssText = "font-size:.8rem;color:#718096;margin-top:4px;";
  hint.textContent   = "Strg+Klick für Mehrfachauswahl · Keine Auswahl = alle Dokumente";

  container.appendChild(select);
  container.appendChild(hint);
}

// ─── Info-Text ─────────────────────────────────────────────────────────
function updateInfo() {
  const from = $("date-from").value;
  const to   = $("date-to").value;
  const info = $("selected-info");

  if (!from || !to) {
    info.textContent = "Bitte Datumsbereich wählen.";
    return;
  }
  const tagText = selectedTags.size > 0
    ? `${selectedTags.size} Tag(s) ausgewählt`
    : "alle Tags";
  info.textContent = `${formatDate(from)} – ${formatDate(to)} · ${tagText}`;
}

function formatDate(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}.${m}.${y}`;
}

// ─── Export starten ────────────────────────────────────────────────────
async function startExport() {
  const from = $("date-from").value;
  const to   = $("date-to").value;
  if (!from || !to) {
    alert("Bitte Datumsbereich auswählen.");
    return;
  }

  const tagIds   = Array.from(selectedTags);
  const tagNames = tagIds.map(id => {
    const t = allTags.find(t => t.id === id);
    return t ? t.name : String(id);
  });
  const yearLabel = from.slice(0, 4);

  $("config-card").classList.add("hidden");
  $("progress-card").classList.remove("hidden");
  $("download-area").style.display = "none";

  const bar = $("progress-bar");
  bar.className = "progress-bar indeterminate";

  const logBox = $("log-box");
  logBox.innerHTML = "";

  try {
    const res = await fetch("/api/start", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ date_from: from, date_to: to, tag_ids: tagIds, tag_names: tagNames, year_label: yearLabel }),
    });
    if (!res.ok) {
      const err = await res.json();
      logLine(`Fehler: ${err.error}`, "error");
      return;
    }
  } catch (err) {
    logLine(`Verbindungsfehler: ${err.message}`, "error");
    return;
  }

  // Polling starten
  polling = setInterval(pollStatus, 800);
}

// ─── Status pollen ─────────────────────────────────────────────────────
let lastLogCount = 0;

async function pollStatus() {
  try {
    const res  = await fetch("/api/status");
    const data = await res.json();

    // Neue Log-Zeilen
    const lines = data.log || [];
    for (let i = lastLogCount; i < lines.length; i++) {
      const line = lines[i];
      if (line.includes("✗") || line.toLowerCase().includes("fehler")) {
        logLine(line, "error");
      } else if (line.startsWith("✓")) {
        logLine(line, "done");
      } else {
        logLine(line);
      }
    }
    lastLogCount = lines.length;

    if (data.done) {
      clearInterval(polling);
      polling = null;
      lastLogCount = 0;

      const bar = $("progress-bar");
      bar.className = "progress-bar";

      if (data.error) {
        bar.style.width = "100%";
        bar.style.background = "#e53e3e";
      } else {
        bar.style.width = "100%";
        bar.style.background = "#2e7d32";
        $("download-area").style.display = "flex";
        $("progress-card").querySelector(".card-title").textContent = "Export abgeschlossen";
      }
    }
  } catch { /* ignore polling errors */ }
}

function logLine(text, type = "") {
  const logBox = $("log-box");
  const line   = document.createElement("div");
  line.textContent = text;
  if (type === "error") line.classList.add("log-line-error");
  if (type === "done")  line.classList.add("log-line-done");
  logBox.appendChild(line);
  logBox.scrollTop = logBox.scrollHeight;
}

// ─── Reset ─────────────────────────────────────────────────────────────
function resetUI() {
  $("progress-card").classList.add("hidden");
  $("config-card").classList.remove("hidden");
  $("progress-bar").style.width = "0%";
  $("progress-bar").style.background = "var(--primary)";
  $("progress-bar").className = "progress-bar";
  $("log-box").innerHTML = "";
  lastLogCount = 0;
}
