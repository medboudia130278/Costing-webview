/**
 * app.js — Maintenance Costing Control Panel
 * Handles: navigation, card states, log rendering, modals, API calls
 */

"use strict";

/* ═══════════════════════════════════════════════════════════
   SECTION METADATA
   ═══════════════════════════════════════════════════════════ */
const SECTIONS = {
  railway: {
    subtitle:    "RAILWAY INFRASTRUCTURE · MAINTENANCE COSTING",
    title:       "Railway Linear Assets",
    accent:      "#4ea3ff",
    sidebarStripe: "#4ea3ff",
    cardCount:   2,
  },
  apm: {
    subtitle:    "APM INFRASTRUCTURE · MAINTENANCE COSTING",
    title:       "APM Linear Assets",
    accent:      "#7c3aed",
    sidebarStripe: "#7c3aed",
    cardCount:   2,
  },
  shifts: {
    subtitle:    "MAINTENANCE PLANNING & WORKFORCE",
    title:       "Planning & Iterated Shifts",
    accent:      "#FBBF24",
    sidebarStripe: "#F59E0B",
    cardCount:   4,
  },
  assess: {
    subtitle:    "WORKFORCE & LOGISTICS ASSESSMENT",
    title:       "Shift Assessment & Vehicles",
    accent:      "#0891b2",
    sidebarStripe: "#0891b2",
    cardCount:   3,
  },
  benchmark: {
    subtitle:    "CROSS-PROJECT ANALYSIS",
    title:       "Benchmark",
    accent:      "#6366f1",
    sidebarStripe: "#6366f1",
    cardCount:   1,
  },
};

let currentSection = "railway";
let darkMode = false;

/* ═══════════════════════════════════════════════════════════
   COLOR HELPERS
   ═══════════════════════════════════════════════════════════ */
function hexToRgba(hex, alpha) {
  const h = hex.replace("#", "");
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  return `rgba(${r},${g},${b},${alpha})`;
}

function applyCardAccent(card, accent) {
  card.style.borderColor = accent;
  card.style.boxShadow   = `0 1px 2px rgba(0,0,0,.06), 0 3px 0 ${hexToRgba(accent, 0.40)}`;
}

function initCardAccents() {
  Object.entries(SECTIONS).forEach(([key, meta]) => {
    const panel = document.getElementById(`section-${key}`);
    if (!panel) return;
    panel.querySelectorAll(".card").forEach(card => {
      card.dataset.accent = meta.accent;
      applyCardAccent(card, meta.accent);
    });
  });
}

/* ═══════════════════════════════════════════════════════════
   NAVIGATION
   ═══════════════════════════════════════════════════════════ */
function switchSection(key) {
  if (!SECTIONS[key]) return;
  currentSection = key;
  const meta = SECTIONS[key];

  // Update all nav items
  document.querySelectorAll(".nav-item").forEach(el => {
    const isActive = el.dataset.section === key;
    el.classList.toggle("active", isActive);
    const stripe = el.querySelector(".nav-stripe");
    if (stripe) stripe.style.background = isActive ? meta.sidebarStripe : "transparent";
  });

  // Update section header
  document.getElementById("section-subtitle").textContent = meta.subtitle;
  document.getElementById("section-title").textContent    = meta.title;
  document.getElementById("section-accent-bar").style.background = meta.accent;

  // Show/hide panels
  document.querySelectorAll(".section-panel").forEach(el => {
    el.classList.toggle("active", el.id === `section-${key}`);
  });

  // Adapter la hauteur de la zone cartes au nombre de cartes
  const cardsArea = document.querySelector(".sections-stack");
  if (cardsArea) {
    const rows  = Math.ceil(meta.cardCount / 2);
    const cardH = rows * 82 + (rows - 1) * 10 + 40; // rows × hauteur + gaps + padding
    cardsArea.style.height = cardH + "px";
  }
}

/* ═══════════════════════════════════════════════════════════
   CARD STATE
   ═══════════════════════════════════════════════════════════ */
/**
 * setCardState(cardId, state)
 *   cardId : e.g. 'card-railway-linear'
 *   state  : 'ready' | 'running' | 'success' | 'error'
 */
function setCardState(cardId, state) {
  const card = document.getElementById(cardId);
  if (!card) return;

  card.classList.remove("running", "success", "error");

  const badge  = card.querySelector(".card-badge");
  const btn    = card.querySelector(".run-btn");
  const accent = card.dataset.accent || "#4ea3ff";

  switch (state) {
    case "running":
      card.classList.add("running");
      card.style.borderColor = "#16a34a";
      card.style.boxShadow   = `0 1px 2px rgba(0,0,0,.06), 0 3px 0 rgba(34,197,94,0.40)`;
      if (badge) badge.textContent = "RUNNING";
      if (btn)   btn.disabled = true;
      break;

    case "success":
      card.classList.add("success");
      card.style.borderColor = "#22c55e";
      card.style.boxShadow   = `0 1px 2px rgba(0,0,0,.06), 0 3px 0 rgba(34,197,94,0.45)`;
      if (badge) badge.textContent = "DONE";
      if (btn)   btn.disabled = false;
      break;

    case "error":
      card.classList.add("error");
      card.style.borderColor = "#dc2626";
      card.style.boxShadow   = `0 1px 2px rgba(0,0,0,.06), 0 3px 0 rgba(220,38,38,0.45)`;
      if (badge) badge.textContent = "ERROR";
      if (btn)   btn.disabled = false;
      break;

    default: // 'ready'
      applyCardAccent(card, accent);
      if (badge) badge.textContent = "READY";
      if (btn)   btn.disabled = false;
      break;
  }
}

function resetAllCards() {
  document.querySelectorAll(".card").forEach(card => {
    // Only reset cards that are not currently running
    if (!card.classList.contains("running")) {
      setCardState(card.id, "ready");
    }
  });
  appendLog("[i] All idle cards reset to READY.\n");
}

/* ═══════════════════════════════════════════════════════════
   RUN ACTIONS  (called by card RUN buttons)
   ═══════════════════════════════════════════════════════════ */
async function runAction(action) {
  if (!window.pywebview) {
    appendLog("[!] PyWebView API not available yet — please wait.\n");
    return;
  }
  const api = window.pywebview.api;

  switch (action) {
    // ── Railway
    case "railway-linear":  api.run_railway_linear();  break;
    case "railway-ovh":     api.run_railway_ovh();     break;
    // ── APM
    case "apm-linear":      api.run_apm_linear();      break;
    case "apm-ovh":         api.run_apm_ovh();         break;
    // ── Shifts / Planning
    case "maintenance-planning": api.run_maintenance_planning(); break;
    case "maintenance-ovh":      api.run_maintenance_ovh();      break;
    case "iterated-shift":       api.run_iterated_shift();       break;
    case "shift-247":            await runShift247();            break;
    // ── Assessment
    case "shift-assessment": api.run_shift_assessment(); break;
    case "night-shift":      api.run_night_shift();      break;
    case "team-vehicles":    api.run_team_vehicles();    break;
    // ── Benchmark
    case "benchmark":        api.run_benchmark();        break;
    default:
      appendLog(`[!] Unknown action: ${action}\n`);
  }
}

/**
 * Shift 24/7 requires file dialogs before calling the API.
 */
async function runShift247() {
  const api = window.pywebview.api;
  setCardState("card-shift-247", "running");
  appendLog("[→] Shift 24/7: please select the input file…\n");

  const inPath = await api.open_file_dialog();
  if (!inPath) {
    appendLog("[!] Shift 24/7: no input file selected. Aborted.\n");
    setCardState("card-shift-247", "ready");
    return;
  }
  appendLog(`[✓] Input: ${inPath}\n`);
  appendLog("[→] Shift 24/7: please choose the output file…\n");

  const outPath = await api.save_file_dialog();
  if (!outPath) {
    appendLog("[!] Shift 24/7: no output file selected. Aborted.\n");
    setCardState("card-shift-247", "ready");
    return;
  }
  appendLog(`[✓] Output: ${outPath}\n`);

  // The actual run (including state changes) happens inside Python
  // We reset the card first because Python will call setCardState('running') again
  setCardState("card-shift-247", "ready");
  api.run_shift_247(inPath, outPath);
}

/* ═══════════════════════════════════════════════════════════
   LOG TERMINAL
   ═══════════════════════════════════════════════════════════ */
const LOG_OUTPUT = document.getElementById("log-output");

/**
 * appendLog(text)
 * Called from Python via window.evaluate_js("appendLog(`...`)")
 * Parses lines and applies colour spans.
 */
function appendLog(text) {
  if (!text) return;
  const lines = text.split("\n");
  lines.forEach((line, idx) => {
    if (idx === lines.length - 1 && line === "") return; // trailing newline artefact

    const timeStamp = new Date().toLocaleTimeString("en-GB", { hour12: false });
    const timeSpan  = `<span class="log-time">[${timeStamp}] </span>`;

    const low = line.toLowerCase();
    let lineSpan;
    if (/error|exception|traceback|\[✗\]/.test(low)) {
      lineSpan = `<span class="log-error">${escapeHtml(line)}</span>`;
    } else if (/warning|warn/.test(low)) {
      lineSpan = `<span class="log-warn">${escapeHtml(line)}</span>`;
    } else if (/\[✓\]|completed successfully/.test(low)) {
      lineSpan = `<span class="log-success">${escapeHtml(line)}</span>`;
    } else {
      lineSpan = escapeHtml(line);
    }

    LOG_OUTPUT.innerHTML += timeSpan + lineSpan + "\n";
  });

  // Auto-scroll
  const body = LOG_OUTPUT.parentElement;
  body.scrollTop = body.scrollHeight;
}

function escapeHtml(str) {
  if (str == null) return "—";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function clearLog() {
  LOG_OUTPUT.innerHTML = "";
  appendLog("[i] Log cleared.\n");
}

/* ═══════════════════════════════════════════════════════════
   TOOLBAR HANDLERS
   ═══════════════════════════════════════════════════════════ */
async function handleClearLog() {
  if (window.pywebview) {
    await window.pywebview.api.clear_log();
  } else {
    clearLog();
  }
}

async function handleSaveLog() {
  if (!window.pywebview) { appendLog("[!] API not ready.\n"); return; }
  const content = LOG_OUTPUT.innerText;
  if (!content.trim()) { appendLog("[!] Log is empty — nothing to save.\n"); return; }
  const result = await window.pywebview.api.save_log(content);
  if (result && result.ok) {
    appendLog(`[✓] Log saved to: ${result.path}\n`);
  } else if (result && result.error !== "Cancelled") {
    appendLog(`[✗] Could not save log: ${result.error}\n`);
  }
}

async function handleOpenLastError() {
  if (!window.pywebview) { appendLog("[!] API not ready.\n"); return; }
  const result = await window.pywebview.api.open_last_error_log();
  if (result && result.ok) {
    document.getElementById("modal-error-content").textContent = result.content;
    openModal("modal-error-log");
  } else {
    appendLog(`[!] ${result ? result.error : "No error log found."}\n`);
  }
}

/* ═══════════════════════════════════════════════════════════
   MODALS
   ═══════════════════════════════════════════════════════════ */
function openModal(id) {
  const el = document.getElementById(id);
  if (el) el.classList.add("open");
}

function closeModal(id) {
  const el = document.getElementById(id);
  if (el) el.classList.remove("open");
}

function showAbout() {
  openModal("modal-about");
}

async function showLicense() {
  const body = document.getElementById("modal-license-body");

  if (!window.pywebview) {
    body.innerHTML = "<p>PyWebView API not ready.</p>";
    openModal("modal-license");
    return;
  }

  try {
    const info = await window.pywebview.api.get_license_info();

    if (!info || !info.ok) {
      body.innerHTML = `<p class="lic-invalid">${escapeHtml((info && info.error) || "Could not load license.")}</p>`;
    } else {
      const statusClass = info.valid ? "lic-valid" : "lic-invalid";
      const statusIcon  = info.valid ? "✔ Valid"   : "✖ Invalid";
      body.innerHTML = `
        <table class="lic-table">
          <tr><th>Status</th>   <td class="${statusClass}">${statusIcon} — ${escapeHtml(info.status)}</td></tr>
          <tr><th>Client</th>   <td>${escapeHtml(info.name)}</td></tr>
          <tr><th>Issued</th>   <td>${escapeHtml(info.issued)}</td></tr>
          <tr><th>Expires</th>  <td>${escapeHtml(info.expires)}</td></tr>
          <tr><th>Features</th> <td>${escapeHtml(info.features)}</td></tr>
          <tr><th>HWID</th>     <td>${escapeHtml(info.hwid)}</td></tr>
          <tr><th>Path</th>     <td>${escapeHtml(info.path)}</td></tr>
        </table>`;

      // Update sidebar footer name
      if (info.name && info.name !== "N/A") {
        const el = document.getElementById("footer-license-name");
        if (el) el.textContent = info.name;
      }
    }
  } catch (e) {
    body.innerHTML = `<p class="lic-invalid">Error: ${escapeHtml(String(e))}</p>`;
  }

  openModal("modal-license");
}

function showHelp() {
  openModal("modal-help");
}

/* ═══════════════════════════════════════════════════════════
   THEME TOGGLE
   ═══════════════════════════════════════════════════════════ */
function toggleTheme() {
  darkMode = !darkMode;
  document.body.classList.toggle("dark", darkMode);
}

/* ═══════════════════════════════════════════════════════════
   INIT — load license name in sidebar footer
   ═══════════════════════════════════════════════════════════ */
async function initLicenseName() {
  if (!window.pywebview) return;
  try {
    const info = await window.pywebview.api.get_license_info();
    if (info && info.ok && info.name && info.name !== "N/A") {
      const el = document.getElementById("footer-license-name");
      if (el) el.textContent = info.name;
    }
  } catch (_) { /* silent */ }
}

window.addEventListener("pywebviewready", () => {
  initLicenseName();
  appendLog("[✓] Interface ready. Select a section and press RUN.\n");
});

/* Initial section + card accents */
switchSection("railway");
initCardAccents();

/* ═══════════════════════════════════════════════════════════
   SPLITTER — drag to resize cards area / log panel
   ═══════════════════════════════════════════════════════════ */
(function () {
  const splitter     = document.getElementById("splitter");
  const cardsArea    = document.querySelector(".sections-stack");
  if (!splitter || !cardsArea) return;

  let dragging  = false;
  let startY    = 0;
  let startH    = 0;

  splitter.addEventListener("mousedown", e => {
    dragging = true;
    startY   = e.clientY;
    startH   = cardsArea.offsetHeight;
    splitter.classList.add("dragging");
    document.body.style.cursor     = "ns-resize";
    document.body.style.userSelect = "none";
    e.preventDefault();
  });

  document.addEventListener("mousemove", e => {
    if (!dragging) return;
    const delta  = e.clientY - startY;
    const newH   = Math.max(80, Math.min(startH + delta, window.innerHeight - 250));
    cardsArea.style.height = newH + "px";
  });

  document.addEventListener("mouseup", () => {
    if (!dragging) return;
    dragging = false;
    splitter.classList.remove("dragging");
    document.body.style.cursor     = "";
    document.body.style.userSelect = "";
  });
})();
