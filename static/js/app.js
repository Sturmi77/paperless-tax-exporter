/* --- Paperless Tax Exporter · Frontend v2.2 -------------------------- */

const $ = id => document.getElementById(id);

let allTags          = [];
let selectedTags     = new Set();
let doctypeDropdown  = null;  // createChipDropdown Handle für Dokumenttypen
let pollTimer    = null;   // setTimeout-Handle (ersetzt setInterval)
let pollRunning  = false;  // true solange ein fetch läuft → keine Parallelinstanz
let lastLogCount = 0;

// --- Init --------------------------------------------------------------
document.addEventListener("DOMContentLoaded", () => {
  const fy = $("footer-year");
  if (fy) fy.textContent = new Date().getFullYear();

  buildYearButtons();
  loadTags();
  loadDocumentTypes();
  checkConnection();
  updateOutputPath();

  $("date-from").addEventListener("change", () => { clearActiveYear(); updateInfo(); validateSubfolderInput(); });
  $("date-to").addEventListener("change",   () => { clearActiveYear(); updateInfo(); });
  $("subfolder-input").addEventListener("input", validateSubfolderInput);

  $("btn-stage0").addEventListener("click",  () => startExport("stage0"));
  $("btn-stage1").addEventListener("click",  () => startExport("stage1"));
  $("btn-both").addEventListener("click",    () => startExport("both"));
  $("btn-stage2").addEventListener("click",  () => startExport("stage2"));
  $("btn-cancel").addEventListener("click",  cancelJob);
  $("new-export-btn").addEventListener("click", resetUI);

  // Segmented Control für Datumsfeld (Issue #10)
  // A4: ARIA radiogroup-Pattern mit Arrow-Key-Navigation
  updateDateToggleHint();
  const pillBtns = Array.from(document.querySelectorAll('.pill-btn'));
  pillBtns.forEach((btn, idx) => {
    btn.addEventListener('click', () => activatePill(btn));
    btn.addEventListener('keydown', e => {
      let next = null;
      if (e.key === 'ArrowRight' || e.key === 'ArrowDown') {
        e.preventDefault();
        next = pillBtns[(idx + 1) % pillBtns.length];
      } else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') {
        e.preventDefault();
        next = pillBtns[(idx - 1 + pillBtns.length) % pillBtns.length];
      }
      if (next) { activatePill(next); next.focus(); }
    });
  });

  function activatePill(btn) {
    pillBtns.forEach(b => {
      b.classList.remove('pill-active');
      b.setAttribute('aria-pressed', 'false');
      b.setAttribute('tabindex', '-1');
    });
    btn.classList.add('pill-active');
    btn.setAttribute('aria-pressed', 'true');
    btn.setAttribute('tabindex', '0');
    updateDateToggleHint();
  }
  // Nur aktiven Pill im Tab-Pfad (roving tabindex)
  pillBtns.forEach(b => b.setAttribute('tabindex', b.classList.contains('pill-active') ? '0' : '-1'));
});

// --- Verbindungscheck --------------------------------------------------
// SVG-Icons für Connection-Badge (S1: nicht nur Farbe)
const BADGE_ICONS = {
  checking: '<svg class="badge-icon" width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><circle cx="12" cy="12" r="10"/></svg>',
  ok:       '<svg class="badge-icon" width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>',
  error:    '<svg class="badge-icon" width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>',
};

function setBadge(state, text, title = "") {
  const badge = $("connection-status");
  badge.className = `badge badge-${state}`;
  badge.title     = title;
  badge.innerHTML = (BADGE_ICONS[state] || "") + `<span class="badge-label">${text}</span>`;
}

// S4: Globaler API-Fehler-Banner (Paperless nicht erreichbar)
function showGlobalError(msg) {
  let banner = $("global-error-banner");
  if (!banner) {
    banner = document.createElement("div");
    banner.id        = "global-error-banner";
    banner.setAttribute("role", "alert");
    banner.className = "global-error-banner";
    banner.innerHTML = `
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/>
        <line x1="12" y1="16" x2="12.01" y2="16"/>
      </svg>
      <span id="global-error-text"></span>
      <button class="global-error-retry" onclick="checkConnection(); loadTags(); loadDocumentTypes();" title="Erneut versuchen">
        Verbindung wiederherstellen
      </button>`;
    document.querySelector(".main").prepend(banner);
  }
  $("global-error-text").textContent = msg;
  banner.classList.remove("hidden");
}

function hideGlobalError() {
  const banner = $("global-error-banner");
  if (banner) banner.classList.add("hidden");
}

async function checkConnection() {
  setBadge("checking", "Verbinde…");
  try {
    const res  = await fetch("/api/health");
    const data = await res.json();
    if (!data.token_configured) {
      setBadge("error", "Token nicht konfiguriert", "PAPERLESS_TOKEN fehlt – bitte .env auf dem NAS anlegen");
      showGlobalError("PAPERLESS_TOKEN nicht konfiguriert. Bitte .env auf dem NAS anlegen und Container neu starten.");
    } else if (data.paperless_reachable === false) {
      setBadge("error", "Paperless nicht erreichbar", data.error || "");
      showGlobalError("Paperless NGX ist nicht erreichbar. Bitte Verbindung und URL prüfen.");
    } else {
      const ollama = data.ollama_available ? " · Ollama ✓" : " · Ollama ✗";
      setBadge("ok", "Paperless verbunden" + ollama, data.ollama_status || "");
      hideGlobalError();
    }
  } catch {
    setBadge("error", "Verbindungsfehler");
    showGlobalError("Verbindung zur App unterbrochen. Bitte Seite neu laden.");
  }
}

// --- Output-Pfad + Unterordner (Issue #9) ---------------------------------
// Allowlist-Regex (spiegelt Backend-Validierung exakt)
const SUBFOLDER_RE = /^[A-Za-z0-9_\-]{1,50}$/;
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

function getSubfolder() {
  return ($("subfolder-input").value || "").trim();
}

function validateSubfolderInput() {
  const val     = getSubfolder();
  const errEl   = $("subfolder-error");
  const preview = $("subfolder-preview");

  if (!val) {
    errEl.classList.add("hidden");
    preview.classList.add("hidden");
    return true;
  }

  if (!SUBFOLDER_RE.test(val)) {
    errEl.textContent = "Nur Buchstaben, Zahlen, _ und - erlaubt (keine Leerzeichen, Schrägstriche, Sonderzeichen).";
    errEl.classList.remove("hidden");
    preview.classList.add("hidden");
    return false;
  }

  errEl.classList.add("hidden");
  // Pfad-Vorschau: {Jahr}/{Unterordner}/Belege/
  const year = $("date-from").value ? $("date-from").value.slice(0, 4) : "{Jahr}";
  preview.textContent = `→ ${year}/${val}/Belege/`;
  preview.classList.remove("hidden");
  return true;
}

// --- Schnellauswahl Kalenderjahre -------------------------------------
function buildYearButtons() {
  const container = $("quick-years");
  const thisYear  = new Date().getFullYear();
  let firstBtn    = null;
  for (let y = thisYear; y >= thisYear - 3; y--) {
    const btn = document.createElement("button");
    btn.className        = "year-btn";
    btn.textContent      = y;
    btn.dataset.year     = y;
    btn.setAttribute("aria-pressed", "false");
    btn.addEventListener("click", () => selectYear(y, btn));
    container.appendChild(btn);
    if (!firstBtn) firstBtn = btn;
  }
  // Aktuelles Jahr (= erster Button) automatisch vorauswählen (Issue #12)
  if (firstBtn) selectYear(thisYear, firstBtn);
}

function selectYear(year, btn) {
  document.querySelectorAll(".year-btn").forEach(b => {
    b.classList.remove("active");
    b.setAttribute("aria-pressed", "false");
  });
  btn.classList.add("active");
  btn.setAttribute("aria-pressed", "true");
  $("date-from").value = `${year}-01-01`;
  $("date-to").value   = `${year}-12-31`;
  updateInfo();
}

function clearActiveYear() {
  document.querySelectorAll(".year-btn").forEach(b => {
    b.classList.remove("active");
    b.setAttribute("aria-pressed", "false");
  });
}

function getDateField() {
  const active = document.querySelector('.pill-btn.pill-active');
  return active ? active.dataset.value : 'created';
}

// --- Datumsfeld-Hilfetext (Issue #10) --------------------------------------
const DATE_TOGGLE_HINTS = {
  created: 'Filtert nach dem Datum auf der Rechnung ("created" in Paperless)',
  added:   'Filtert nach dem Datum, an dem das Dokument in Paperless eingescannt wurde',
};

function updateDateToggleHint() {
  const hintEl = $('date-toggle-hint');
  if (!hintEl) return;
  const field = getDateField();
  hintEl.textContent = DATE_TOGGLE_HINTS[field] || '';
}

// --- createChipDropdown() Factory ------------------------------------
/**
 * Erstellt ein wiederverwendbares Chip-Dropdown-Widget.
 *
 * config: {
 *   containerId:    ID des äußeren Wrapper-Elements (z.B. 'tag-dropdown')
 *   inputWrapId:    ID des Klick-Bereichs mit Chips + Suchinput
 *   chipsId:        ID des Chip-Containers
 *   searchId:       ID des Suchinput-Elements
 *   panelId:        ID der Dropdown-Liste
 *   loadingId:      ID des Lade-Spinners (optional)
 *   errorId:        ID des Fehler-Elements (optional)
 *   items:          Array von { id, name } Objekten
 *   onSelectionChange: callback(selectedIds: Set) – wird bei jeder Änderung aufgerufen
 *   placeholder:    Placeholder-Text wenn nichts ausgewählt (default: 'Auswählen…')
 * }
 *
 * Gibt zurück: { selectedIds: Set, refresh(items) }
 */
function createChipDropdown(config) {
  const {
    containerId, inputWrapId, chipsId, searchId, panelId,
    loadingId, errorId,
    items = [],
    onSelectionChange = () => {},
    placeholder = 'Auswählen…',
  } = config;

  const container  = $(containerId);
  const inputWrap  = $(inputWrapId);
  const chipsEl    = $(chipsId);
  const searchEl   = $(searchId);
  const panel      = $(panelId);

  const selectedIds = new Set();
  let currentItems  = [...items].sort((a, b) => a.name.localeCompare(b.name, 'de'));

  function renderPanel(filter = '') {
    panel.innerHTML = '';
    const f = filter.toLowerCase();
    const visible = currentItems.filter(item => item.name.toLowerCase().includes(f));
    if (visible.length === 0) {
      panel.innerHTML = '<div class="tag-no-results">Keine Einträge gefunden.</div>';
      return;
    }
    visible.forEach(item => {
      const opt = document.createElement('div');
      opt.className  = 'tag-option' + (selectedIds.has(item.id) ? ' selected' : '');
      opt.dataset.id = item.id;
      opt.innerHTML = `
        <span class="tag-option-check">
          ${selectedIds.has(item.id)
            ? '<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>'
            : ''}
        </span>
        ${item.name}
      `;
      opt.addEventListener('click', () => toggle(item.id));
      panel.appendChild(opt);
    });
  }

  function renderChips() {
    chipsEl.innerHTML = '';
    selectedIds.forEach(id => {
      const item = currentItems.find(t => t.id === id);
      if (!item) return;
      const chip = document.createElement('div');
      chip.className = 'tag-chip';
      chip.innerHTML = `${item.name}<button class="tag-chip-remove" data-id="${id}" title="Entfernen">×</button>`;
      chip.querySelector('.tag-chip-remove').addEventListener('click', e => {
        e.stopPropagation();
        toggle(id);
      });
      chipsEl.appendChild(chip);
    });
    searchEl.placeholder = selectedIds.size === 0 ? placeholder : '';
    onSelectionChange(selectedIds);
  }

  function toggle(id) {
    if (selectedIds.has(id)) selectedIds.delete(id);
    else selectedIds.add(id);
    renderChips();
    renderPanel(searchEl.value);
  }

  function openPanel() {
    renderPanel(searchEl.value);
    panel.classList.remove('hidden');
  }

  function closePanel() {
    panel.classList.add('hidden');
    searchEl.value = '';
  }

  // Events
  inputWrap.addEventListener('click', () => { searchEl.focus(); openPanel(); });
  searchEl.addEventListener('input',  () => renderPanel(searchEl.value));
  searchEl.addEventListener('focus',  () => openPanel());
  document.addEventListener('click',  e => {
    if (!container.contains(e.target)) closePanel();
  });

  // Initialer Render
  if (loadingId) $(loadingId).classList.add('hidden');
  if (container) container.classList.remove('hidden');
  renderChips();

  // Öffentliche API
  return {
    selectedIds,
    /** Ersetzt die Item-Liste (z.B. nach asynchronem Nachladen) */
    refresh(newItems) {
      currentItems = [...newItems].sort((a, b) => a.name.localeCompare(b.name, 'de'));
      renderChips();
      renderPanel(searchEl.value);
    },
  };
}

// --- Dokumenttypen laden (Issue #7) -----------------------------------
async function loadDocumentTypes() {
  try {
    const res  = await fetch('/api/document-types');
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    if (data.error) throw new Error(data.error);

    const types = data.document_types || [];
    if (types.length === 0) {
      // Keine Typen vorhanden – Lade-Spinner ausblenden, kein Fehler
      $('doctype-loading').classList.add('hidden');
      return;
    }

    doctypeDropdown = createChipDropdown({
      containerId:  'doctype-dropdown',
      inputWrapId:  'doctype-input-wrap',
      chipsId:      'doctype-chips',
      searchId:     'doctype-search',
      panelId:      'doctype-panel',
      loadingId:    'doctype-loading',
      items:        types,
      placeholder:  'Typ wählen…',
      onSelectionChange: () => updateInfo(),
    });
  } catch (err) {
    $('doctype-loading').classList.add('hidden');
    const errEl = $('doctype-error');
    errEl.textContent = `Fehler beim Laden der Dokumenttypen: ${err.message}`;
    errEl.classList.remove('hidden');
  }
}

// --- Tags laden --------------------------------------------------------
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
    enableButtons();
    updateInfo();
  } catch (err) {
    $("tag-loading").classList.add("hidden");
    $("tag-error").textContent = `Fehler beim Laden der Tags: ${err.message}`;
    $("tag-error").classList.remove("hidden");
    enableButtons();
  }
}

const EXPORT_BTN_IDS = ["btn-stage0", "btn-stage1", "btn-both", "btn-stage2"];

function enableButtons() {
  // F3: aria-disabled + disabled synchron; Buttons bleiben fokussierbar bis disabled gesetzt
  EXPORT_BTN_IDS.forEach(id => {
    const btn = $(id);
    btn.disabled = false;
    btn.removeAttribute("aria-disabled");
  });
}

function disableButtons() {
  EXPORT_BTN_IDS.forEach(id => {
    const btn = $(id);
    btn.disabled = true;
    btn.setAttribute("aria-disabled", "true");
  });
}

// tagDropdown-Handle für externen Zugriff (selectedTags sync)
let tagDropdown = null;

function renderTags() {
  $('tag-loading').classList.add('hidden');

  if (allTags.length === 0) {
    $('tag-error').textContent = 'Keine Tags in Paperless gefunden.';
    $('tag-error').classList.remove('hidden');
    return;
  }

  // createChipDropdown() Factory verwenden
  tagDropdown = createChipDropdown({
    containerId:  'tag-dropdown',
    inputWrapId:  'tag-input-wrap',
    chipsId:      'tag-chips',
    searchId:     'tag-search',
    panelId:      'tag-panel',
    loadingId:    'tag-loading',
    items:        allTags,
    placeholder:  'Tags wählen…',
    onSelectionChange: (ids) => {
      // selectedTags (globale Variable) synchron halten
      selectedTags = ids;
      updateInfo();
    },
  });

  // selectedTags auf das gleiche Set zeigen lassen
  selectedTags = tagDropdown.selectedIds;
}

// --- Info-Text ---------------------------------------------------------
function updateInfo() {
  const from = $("date-from").value;
  const to   = $("date-to").value;
  const info = $("selected-info");
  if (!from || !to) {
    // F3: erklärender Initialtext warum Buttons inaktiv sind
    info.textContent = "Datumsbereich wählen, um Export zu starten.";
    return;
  }
  const tagText = selectedTags.size > 0
    ? `${selectedTags.size} Tag(s) ausgewählt`
    : "alle Tags";
  const doctypeText = (doctypeDropdown && doctypeDropdown.selectedIds.size > 0)
    ? ` · ${doctypeDropdown.selectedIds.size} Typ(en)`
    : "";
  info.textContent = `${formatDate(from)} – ${formatDate(to)} · ${tagText}${doctypeText}`;
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

// --- Issue #4: Überschreib-Prüfung ------------------------------------
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
  const modal = $("overwrite-modal");
  modal.classList.remove("hidden");

  // Fokus auf ersten Button setzen (A1: Focus Management)
  setTimeout(() => $("btn-overwrite-cancel").focus(), 50);

  return new Promise(resolve => {
    function closeModal(result) {
      modal.classList.add("hidden");
      document.removeEventListener("keydown", escHandler);
      modal.removeEventListener("click", backdropHandler);
      resolve(result);
    }

    // M2: Escape-Key schließt Modal
    function escHandler(e) {
      if (e.key === "Escape") closeModal(false);
    }
    document.addEventListener("keydown", escHandler);

    // M3: Klick auf Backdrop schließt Modal
    function backdropHandler(e) {
      if (e.target === modal) closeModal(false);
    }
    modal.addEventListener("click", backdropHandler);

    $("btn-overwrite-append").onclick  = () => closeModal("append");
    $("btn-overwrite-confirm").onclick = () => closeModal(true);
    $("btn-overwrite-cancel").onclick  = () => closeModal(false);
  });
}

// --- Export starten ----------------------------------------------------
async function startExport(mode) {
  const from = $("date-from").value;
  const to   = $("date-to").value;
  if (!from || !to) {
    alert("Bitte Datumsbereich auswählen.");
    return;
  }

  const tagIds          = Array.from(selectedTags);
  const tagNames        = tagIds.map(id => { const t = allTags.find(t => t.id === id); return t ? t.name : String(id); });
  const doctypeIds      = doctypeDropdown ? Array.from(doctypeDropdown.selectedIds) : [];
  const yearLabel       = from.slice(0, 4);
  const dateField       = getDateField();
  const subfolder       = getSubfolder();

  // Clientseitige Subfolder-Validierung (Issue #9)
  if (subfolder && !SUBFOLDER_RE.test(subfolder)) {
    $("subfolder-error").textContent = "Ungültiger Unterordner – bitte korrigieren.";
    $("subfolder-error").classList.remove("hidden");
    $("subfolder-input").focus();
    return;
  }

  // Issue #4: Überschreib-Prüfung (false=Abbrechen, true=Überschreiben, "append"=Nur neue)
  const confirmed = await checkExistsAndConfirm(yearLabel, mode);
  if (confirmed === false) return;


  const appendMode = confirmed === "append";

  const titles = {
    stage0:  "Nur Excel wird erstellt…",
    stage1:  appendMode ? "Neue Belege werden hinzugefügt…" : "PDFs & Excel wird erstellt…",
    stage2:  "OCR-Analyse läuft…",
    both:    appendMode ? "Neue Belege + OCR-Analyse läuft…" : "Vollständiger Export läuft…",
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
        date_field:  dateField,
        tag_ids: tagIds, tag_names: tagNames,
        document_type_ids: doctypeIds,
        subfolder,
        year_label: yearLabel, mode,
        append_mode: appendMode,
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

  // setTimeout-Loop starten (kein setInterval → keine parallelen Instanzen)
  schedulePoll();
}

// --- Poll-Loop (setTimeout statt setInterval) --------------------------
// Garantiert: nächster Poll startet erst NACH Abschluss des aktuellen fetch.
// Damit kann es keine zwei gleichzeitig laufenden pollStatus()-Instanzen geben,
// die den Button-Zustand wechselseitig überschreiben (Blink-Ursache).
function schedulePoll() {
  pollTimer = setTimeout(pollStatus, 800);
}

function stopPoll() {
  if (pollTimer !== null) {
    clearTimeout(pollTimer);
    pollTimer = null;
  }
  pollRunning = false;
}

async function pollStatus() {
  pollTimer   = null;
  pollRunning = true;

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

    // Abbrechen-Button:
    // - Nur beim ersten Einblenden (hidden→sichtbar) wird der Inhalt gesetzt.
    // - Wenn bereits sichtbar: DOM NICHT anfassen → "Wird abgebrochen…" bleibt.
    // - pollRunning=true schützt auch vor Race Conditions durch setInterval.
    const cancelBtn = $("btn-cancel");
    if (data.cancellable) {
      if (cancelBtn.classList.contains("hidden")) {
        cancelBtn.disabled  = false;
        cancelBtn.innerHTML = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
          <circle cx="12" cy="12" r="10"/>
          <line x1="15" y1="9" x2="9" y2="15"/>
          <line x1="9" y1="9" x2="15" y2="15"/>
        </svg> OCR abbrechen`;
        cancelBtn.classList.remove("hidden");
        // P4: Fokus auf Cancel-Button beim ersten Einblenden
        cancelBtn.focus();
      }
      // bereits sichtbar → nichts tun
    } else {
      cancelBtn.classList.add("hidden");
    }

    // OCR-Fortschritt (Stufe 2)
    if (data.stage === "stage2" && data.ocr_total > 0) {
      $("ocr-progress").classList.remove("hidden");
      $("ocr-counter").textContent  = `${data.ocr_current} / ${data.ocr_total} Dokumente`;
      $("ocr-doc-title").textContent = data.ocr_current_title || "";
      const pct = Math.round((data.ocr_current / data.ocr_total) * 100);
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

    // P3: Phasen-Label in progress-title aktualisieren
    const titleEl = $("progress-title");
    if (titleEl && data.stage && !data.done) {
      const phaseLabels = {
        stage0: "Excel wird erstellt…",
        stage1: "PDFs werden heruntergeladen…",
        stage2: "OCR-Analyse läuft…",
        both:   "Export läuft…",
      };
      const lbl = phaseLabels[data.stage];
      // Nur ändern wenn sich Phase geändert hat (nicht bei jedem Poll-Tick)
      if (lbl && titleEl.dataset.phase !== data.stage) {
        titleEl.textContent   = lbl;
        titleEl.dataset.phase = data.stage;
      }
    }

    if (data.done) {
      stopPoll();
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
        if (ct) {
          ct.textContent   = data.cancelled ? "Job abgebrochen" : "Export abgeschlossen";
          ct.dataset.phase = "done";
        }
      }
      return; // kein schedulePoll() mehr
    }
  } catch { /* ignore */ }

  pollRunning = false;
  schedulePoll(); // nächsten Tick erst hier einplanen → nie parallel
}

// --- Issue #2: Job abbrechen -------------------------------------------
async function cancelJob() {
  const btn = $("btn-cancel");
  btn.disabled  = true;
  btn.innerHTML = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
    <circle cx="12" cy="12" r="10"/>
    <line x1="15" y1="9" x2="9" y2="15"/>
    <line x1="9" y1="9" x2="15" y2="15"/>
  </svg> Wird abgebrochen…`;
  try {
    await fetch("/api/cancel", { method: "POST" });
  } catch { /* ignore */ }
  // Button bleibt disabled – pollStatus() versteckt ihn wenn cancellable=false
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

// --- Reset -------------------------------------------------------------
function resetToConfig() {
  stopPoll();
  $("progress-card").classList.add("hidden");
  $("config-card").classList.remove("hidden");
  enableButtons();
}

function resetUI() {
  stopPoll();
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
