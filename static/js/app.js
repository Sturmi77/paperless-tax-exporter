/* ─── Paperless Tax Exporter · Frontend v2.1 ────────────────────────── */

const $ = id => document.getElementById(id);

let allTags      = [];
let selectedTags = new Set();
let polling      = null;
let lastLogCount = 0;

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

  $("btn-stage0").addEventListener("click",  () => startExport("stage0"));
  $("btn-stage1").addEventListener("click",  () => startExport("stage1"));
  $("btn-both").addEventListener("click",    () => startExport("both"));
  $("btn-stage2").addEventListener("click",  () => startExport("stage2"));
  $("btn-cancel").addEventListener("click",  cancelJob);
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
      badge.title       = "PAPERLESS_TOKEN fehlt – bitte .env auf dem NAS anlegen";
    } else if (data.paperless_reachable === false) {
      badge.textContent = "Paperless nicht erreichbar";
      badge.className   = "badge badge-error";
      badge.title       = data.error || "";
    } else {
      const ollama = data.ollama_available ? " · Ollama ✓" : " · Ollama ✗";
      badge.textContent = "Paperless verbunden" + ollama;
      badge.className   = "badge badge-ok";
      badge.title       = data.ollama_status || "";
    }
  } catch {
    badge.textContent = "Verbindungsfehler";
    badge.className   = "badge badge-error";
  }
}

// ─── Output-Pfad ──────────────────────────────────────────────────────
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
  for (let y = thisYear; y >= thisYear - 3; y--) {
    const btn = document.createElement("button");
    btn.className    = "year-btn";
    btn.textContent  = y;
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
    enableButtons();
    updateInfo();
  } catch (err) {
    $("tag-loading").classList.add("hidden");
    $("tag-error").textContent = `Fehler beim Laden der Tags: ${err.message}`;
    $("tag-error").classList.remove("hidden");
    enableButtons();
  }
}

function enableButtons() {
  $("btn-stage0").disabled = false;
  $("btn-stage1").disabled = false;
  $("btn-both").disabled   = false;
  $("btn-stage2").disabled = false;
}

function disableButtons() {
  $("btn-stage0").disabled = true;
  $("btn-stage1").disabled = true;
  $("btn-both").disabled   = true;
  $("btn-stage2").disabled = true;
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
  select.id       = "tag-select";
  select.multiple = true;
  select.size     = Math.min(sorted.length, 8);
  select.style.cssText = "width:100%;border:1.5px solid #c5d2d4;border-radius:6px;padding:4px;font-size:.93rem;font-family:inherit;background:#fff;color:#444;outline:none;";

  sorted.forEach(tag => {
    const opt       = document.createElement("option");
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

function formatDuration(seconds) {
  const s = Math.round(seconds);
  if (s < 60) return `${s} Sek.`;
  const m = Math.floor(s / 60);
  const r = s % 60;
  if (m < 60) return r > 0 ? `${m} Min. ${r} Sek.` : `${m} Min.`;
  const h  = Math.floor(m / 60);
  const rm = m % 60;
  return rm > 0 ? `${h} Std. ${rm} Min.` : `${h} Std.`;
}

function getDateField() {
  const el = document.querySelector("input[name='date-field']:checked");
  return el ? el.value : "created";
}

// ─── Issue #4: Überschreib-Prüfung ────────────────────────────────────
async function checkExistsAndConfirm(yearLabel, mode) {
  // Nur bei Stufe 1 oder both relevant (nicht stage0, nicht stage2)
  if (mode === "stage2" || mode === "stage0") return true;

  let data;
  try {
    const res = await fetch(`/api/check-exists?year=${encodeURIComponent(yearLabel)}`);
    data = await res.json();
  } catch {
    return true; // Bei Fehler einfach fortfahren
  }

  if (!data.excel_exists && !data.pdfs_exist) return true; // nichts vorhanden

  // Modal befüllen und anzeigen
  let details = "";
  if (data.excel_exists) details += `<li>Excel-Datei: <strong>Rechnungsaufstellung_${yearLabel}.xlsx</strong></li>`;
  if (data.pdfs_exist)   details += `<li>PDF-Ordner: <strong>Belege/</strong> (${data.pdf_count} Dateien)</li>`;

  $("overwrite-details").innerHTML = details;
  $("overwrite-modal").classList.remove("hidden");

  return new Promise(resolve => {
    $("btn-overwrite-confirm").onclick = () => {
      $("overwrite-modal").classList.add("hidden");
      resolve(true);
    };
    $("btn-overwrite-cancel").onclick = () => {
      $("overwrite-modal").classList.add("hidden");
      resolve(false);
    };
  });
}

// ─── Export starten ────────────────────────────────────────────────────
async function startExport(mode) {
  const from = $("date-from").value;
  const to   = $("date-to").value;
  if (!from || !to) {
    alert("Bitte Datumsbereich auswählen.");
    return;
  }

  const tagIds    = Array.from(selectedTags);
  const tagNames  = tagIds.map(id => { const t = allTags.find(t => t.id === id); return t ? t.name : String(id); });
  const yearLabel = from.slice(0, 4);
  const dateField = getDateField();  // Issue #3

  // Issue #4: Überschreib-Prüfung
  const confirmed = await checkExistsAndConfirm(yearLabel, mode);
  if (!confirmed) return;

  const titles = {
    stage0: "Excel wird erstellt…",
    stage1: "Stufe 1: PDFs & Excel wird erstellt…",
    stage2: "Stufe 2: OCR-Analyse läuft…",
    both:   "Stufe 1 + 2: Export & OCR-Analyse läuft…",
  };

  $("config-card").classList.add("hidden");
  $("progress-card").classList.remove("hidden");
  $("download-area").style.display = "none";
  $("ocr-progress").classList.add("hidden");
  $("btn-cancel").classList.add("hidden");

  const titleEl = $("progress-title");
  if (titleEl) titleEl.textContent = titles[mode] || "Export läuft…";

  const bar = $("progress-bar");
  bar.className     = "progress-bar indeterminate";
  bar.style.width   = "";
  bar.style.background = "";

  $("log-box").innerHTML = "";
  lastLogCount = 0;

  try {
    const res = await fetch("/api/start", {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body:    JSON.stringify({
        date_from: from, date_to: to,
        tag_ids: tagIds, tag_names: tagNames,
        year_label: yearLabel, mode,
        date_field: dateField,   // Issue #3
      }),
    });
    if (!res.ok) {
      const err = await res.json();
      logLine(`Fehler: ${err.error}`, "error");
      resetToConfig();
      return;
    }
  } catch (err) {
    logLine(`Verbindungsfehler: ${err.message}`, "error");
    resetToConfig();
    return;
  }

  polling = setInterval(pollStatus, 800);
}

// ─── Issue #2: Job abbrechen ───────────────────────────────────────────
async function cancelJob() {
  $("btn-cancel").disabled = true;
  $("btn-cancel").textContent = "Abbrechen…";
  try {
    await fetch("/api/cancel", { method: "POST" });
  } catch { /* ignore */ }
}

// ─── Status pollen ─────────────────────────────────────────────────────
async function pollStatus() {
  try {
    const res  = await fetch("/api/status");
    const data = await res.json();

    const lines = data.log || [];
    for (let i = lastLogCount; i < lines.length; i++) {
      const line = lines[i];
      if (line.includes("✗") || line.toLowerCase().includes("fehler")) {
        logLine(line, "error");
      } else if (line.startsWith("✓")) {
        logLine(line, "done");
      } else if (line.startsWith("⚠")) {
        logLine(line, "warn");
      } else {
        logLine(line);
      }
    }
    lastLogCount = lines.length;

    // Issue #2: Abbrechen-Button zeigen während OCR läuft
    if (data.stage === "stage2" && !data.done) {
      $("btn-cancel").classList.remove("hidden");
    } else {
      $("btn-cancel").classList.add("hidden");
    }

    // OCR-Fortschritt (Stufe 2)
    if (data.stage === "stage2" && data.ocr_total > 0) {
      $("ocr-progress").classList.remove("hidden");
      $("ocr-counter").textContent  = `${data.ocr_current} / ${data.ocr_total} Dokumente`;
      $("ocr-doc-title").textContent = data.ocr_current_title || "";
      const pct    = Math.round((data.ocr_current / data.ocr_total) * 100);
      $("progress-bar-ocr").style.width = pct + "%";

      if (data.avg_sec_per_doc !== null && data.avg_sec_per_doc !== undefined) {
        $("ocr-avg-time").textContent = `Ø ${formatDuration(data.avg_sec_per_doc)} / Dokument`;
      }
      if (data.eta_seconds !== null && data.eta_seconds !== undefined && data.eta_seconds >= 0) {
        $("ocr-eta").textContent = data.eta_seconds < 10
          ? "Gleich fertig…"
          : `noch ca. ${formatDuration(data.eta_seconds)}`;
      }
    }

    if (data.done) {
      clearInterval(polling);
      polling      = null;
      lastLogCount = 0;
      $("btn-cancel").classList.add("hidden");

      const bar = $("progress-bar");
      bar.className = "progress-bar";

      if (data.error) {
        bar.style.width      = "100%";
        bar.style.background = "#e53e3e";
      } else {
        bar.style.width      = "100%";
        bar.style.background = data.cancelled ? "#e69800" : "#2e7d32";
        $("download-area").style.display = "flex";
        const ct = $("progress-title");
        if (ct) ct.textContent = data.cancelled ? "Job abgebrochen" : "Export abgeschlossen";
      }
    }
  } catch { /* ignore */ }
}

function logLine(text, type = "") {
  const logBox = $("log-box");
  const line   = document.createElement("div");
  line.textContent = text;
  if (type === "error") line.classList.add("log-line-error");
  if (type === "done")  line.classList.add("log-line-done");
  if (type === "warn")  line.classList.add("log-line-warn");
  logBox.appendChild(line);
  logBox.scrollTop = logBox.scrollHeight;
}

// ─── Reset ─────────────────────────────────────────────────────────────
function resetToConfig() {
  $("progress-card").classList.add("hidden");
  $("config-card").classList.remove("hidden");
  enableButtons();
}

function resetUI() {
  $("progress-card").classList.add("hidden");
  $("config-card").classList.remove("hidden");
  $("progress-bar").style.width      = "0%";
  $("progress-bar").style.background = "var(--primary)";
  $("progress-bar").className        = "progress-bar";
  $("log-box").innerHTML             = "";
  $("ocr-progress").classList.add("hidden");
  $("btn-cancel").classList.add("hidden");
  lastLogCount = 0;
  enableButtons();
}
