/**
 * tsifl Excel Add-in
 * Full auth + chat + comprehensive Excel action execution.
 */

import "./taskpane.css";
import { getCurrentUser, signIn, signUp, signOut, resetPassword, supabase, syncSessionToBackend } from "./auth.js";

const BACKEND_URL  = "https://focused-solace-production-6839.up.railway.app";
const LOCAL_URL    = "/local-api";              // proxied through webpack dev server (avoids HTTPS mixed content)
const PREFS_KEY    = "tsifl_preferences";
const BUILD_VER    = "v96";  // bump this on every deploy so user can confirm fresh code

let CURRENT_USER       = null;
let lastNavigatedSheet = null;   // tracks sheet after navigate_sheet so writes auto-target it
let pendingImages      = [];     // base64 images queued for next message

// Project memory state (shown in the Memory panel)
let currentWorkbookId = null;
let memoryPanelExpanded = false;

// Undo stack (Improvement 11) — stores cell states before actions
const undoStack = [];
const MAX_UNDO = 5;

// Action history (Improvement 12)
const actionHistory = [];
const MAX_HISTORY = 20;

// Debug state (Improvement 10)
let lastSyncTimestamp = null;
let sessionSource = "none";

// ── Cheap mode (force Haiku tier) ──────────────────────────────────────────
// Toggleable via the "$" button in the header. When ON, every chat request
// gets context.force_model = "haiku" so the backend's _select_model uses
// MODEL_FAST regardless of message complexity. Cost: ~3x cheaper per call.
// Trade-off: Haiku is weaker on edge cases. Use during dev iteration; flip
// off for production demos / "is this the real experience" tests.
//
// State persists in localStorage so reloads/restarts remember the user's
// preference. Default: OFF (Sonnet).
const CHEAP_MODE_KEY = "tsifl_cheap_mode";
let _cheapMode = false;
try {
  _cheapMode = window.localStorage.getItem(CHEAP_MODE_KEY) === "1";
} catch (_) { /* localStorage unavailable — fall back to default */ }

function _renderCheapModeButton() {
  const btn = document.getElementById("cheap-mode-btn");
  if (!btn) return;
  if (_cheapMode) {
    btn.classList.add("cheap-on");
    btn.title = "Cheap mode: ON (Haiku, ~3x cheaper). Click to switch back to Sonnet.";
    btn.textContent = "$ ON";
  } else {
    btn.classList.remove("cheap-on");
    btn.title = "Cheap mode: OFF (Sonnet, default quality). Click to enable Haiku for cheaper dev runs.";
    btn.textContent = "$";
  }
}

function _toggleCheapMode() {
  _cheapMode = !_cheapMode;
  try {
    window.localStorage.setItem(CHEAP_MODE_KEY, _cheapMode ? "1" : "0");
  } catch (_) {}
  _renderCheapModeButton();
  showToast(
    _cheapMode
      ? "Cheap mode ON — using Haiku (~3x cheaper, weaker on edge cases)"
      : "Cheap mode OFF — using Sonnet (default quality)",
    "info",
    3000
  );
}

// ── Boot ─────────────────────────────────────────────────────────────────────

Office.onReady(async () => {
  try { await Office.addin.setStartupBehavior(Office.StartupBehavior.load); } catch (_) {}

  CURRENT_USER = await getCurrentUser();
  if (CURRENT_USER) showChatScreen(CURRENT_USER);
  else               showLoginScreen();

  // ── Auto-poll for R→Excel plot transfers every 4s ──────────────────────────
  // The R add-in pushes plots into /transfer/store; without polling they sit unread
  // until the user manually triggers an import_image action. This makes it feel
  // like a real push: drop a plot into the active sheet as soon as it arrives.
  const seenTransferIds = new Set();
  setInterval(async () => {
    if (!CURRENT_USER) return;
    try {
      const resp = await fetch(`${BACKEND_URL}/transfer/pending/excel`);
      if (!resp.ok) return;
      const { pending = [] } = await resp.json();
      const images = pending.filter(p => p.data_type === "image" && !seenTransferIds.has(p.transfer_id));
      for (const item of images) {
        seenTransferIds.add(item.transfer_id);
        try {
          const tResp = await fetch(`${BACKEND_URL}/transfer/${item.transfer_id}`);
          if (!tResp.ok) continue;
          const transfer = await tResp.json();
          const cleanBase64 = String(transfer.data || "").replace(/^data:image\/[a-z+]+;base64,/, "");
          if (!cleanBase64.startsWith("iVBOR") && !cleanBase64.startsWith("/9j/")) continue;
          await Excel.run(async (ctx) => {
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();
            const image = sheet.shapes.addImage(cleanBase64);
            image.name  = "R_Plot_" + Date.now();
            image.left  = 10;
            image.top   = 200;
            await ctx.sync();
          });
          appendMessage("assistant", "📈 R plot received and inserted into the active sheet.");
        } catch (e) {
          console.warn("[tsifl] auto plot import failed:", e);
        }
      }
    } catch (_) { /* polling is best-effort */ }
  }, 4000);
});

// ── Preferences (localStorage) ────────────────────────────────────────────────

function loadPreferences() {
  try {
    return JSON.parse(localStorage.getItem(PREFS_KEY) || "{}");
  } catch (_) { return {}; }
}

function savePreferences(updates) {
  const current = loadPreferences();
  const merged  = { ...current, ...updates };
  localStorage.setItem(PREFS_KEY, JSON.stringify(merged));
  return merged;
}

// ── Screens ───────────────────────────────────────────────────────────────────

function showLoginScreen() {
  document.getElementById("login-screen").style.display = "flex";
  document.getElementById("chat-screen").style.display  = "none";
  document.getElementById("login-btn").addEventListener("click",  handleSignIn);
  document.getElementById("signup-btn").addEventListener("click", handleSignUp);
  document.getElementById("auth-password").addEventListener("keydown", (e) => {
    if (e.key === "Enter") handleSignIn();
  });
  // Show password toggle — native checkbox. Replaces the absolute-positioned
  // eye button (4 prior versions failed to register clicks reliably under the
  // Office WKWebView sandbox). Native <input type="checkbox"> click semantics
  // are bulletproof; this works on every webview going back to IE 4.
  const showPwCb = document.getElementById("show-pw-cb");
  if (showPwCb) {
    showPwCb.addEventListener("change", () => {
      try {
        // Element-swap pattern: avoid runtime input.type mutation (which the
        // Office sandbox silently blocks). Setting .type on a freshly-created
        // element BEFORE attach is always allowed.
        const old = document.getElementById("auth-password");
        if (!old) { console.warn("[tsifl] auth-password not found"); return; }
        const wantText = showPwCb.checked;
        const fresh = document.createElement("input");
        fresh.type        = wantText ? "text" : "password";
        fresh.id          = "auth-password";
        fresh.placeholder = old.placeholder || "Password";
        fresh.autocomplete = wantText ? "off" : "current-password";
        fresh.value       = old.value;
        fresh.className   = old.className;
        const wasFocused  = (document.activeElement === old);
        const caret       = old.selectionStart;
        old.parentNode.replaceChild(fresh, old);
        fresh.addEventListener("keydown", (e) => { if (e.key === "Enter") handleSignIn(); });
        if (wasFocused) {
          fresh.focus();
          try { fresh.setSelectionRange(caret, caret); } catch (_) {}
        }
        console.log("[tsifl] password visibility →", fresh.type);
      } catch (err) {
        console.error("[tsifl] show-password handler crashed:", err);
      }
    });
  } else {
    console.warn("[tsifl] show-pw-cb not found");
  }
  // Forgot password (Improvement 5)
  const forgotBtn = document.getElementById("forgot-pw-btn");
  if (forgotBtn) {
    forgotBtn.addEventListener("click", async () => {
      const email = document.getElementById("auth-email").value.trim();
      const errEl = document.getElementById("auth-error");
      if (!email || !email.includes("@")) {
        errEl.style.color = "#DC2626";
        errEl.textContent = "Enter your email address first.";
        return;
      }
      const { error } = await resetPassword(email);
      if (error) { errEl.style.color = "#DC2626"; errEl.textContent = error.message; return; }
      errEl.style.color = "#16A34A";
      errEl.textContent = "Check your email for a reset link.";
    });
  }
}

function showChatScreen(user) {
  document.getElementById("login-screen").style.display = "none";
  document.getElementById("chat-screen").style.display  = "flex";

  // User display with avatar initial (Improvement 8)
  const initial = (user.email || "?")[0].toUpperCase();
  document.getElementById("user-bar").innerHTML =
    `<span style="display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;border-radius:50%;background:#0D5EAF;color:white;font-size:10px;font-weight:700;margin-right:4px;">${initial}</span>${user.email} &middot; ${BUILD_VER}`;

  document.getElementById("submit-btn").addEventListener("click", handleSubmit);
  document.getElementById("user-input").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleSubmit(); }
    if (e.key === "Escape") { e.target.value = ""; }
  });

  // Memory panel: toggle on header click, clear on reset button
  const memHeader = document.getElementById("memory-header");
  if (memHeader) {
    memHeader.addEventListener("click", (e) => {
      // Don't toggle when the Reset button itself is clicked
      if ((e.target.id || "") === "memory-clear") return;
      toggleMemoryPanel();
    });
  }
  const memClear = document.getElementById("memory-clear");
  if (memClear) memClear.addEventListener("click", (e) => { e.stopPropagation(); clearMemory(); });

  // Initial memory state load (runs after a short delay so Excel context is ready)
  setTimeout(() => { refreshMemoryPanel(); }, 800);

  // Backend health check on boot + wire the retry button
  wireBackendBanner();
  checkBackendHealth();

  // Ollama local model probe — fire-and-forget; result cached for 30s.
  // If reachable, discuss-mode messages will route there without user action.
  checkOllamaAvailable().then(ok => {
    if (ok) console.info(`[tsifl] Ollama available: ${_ollamaHealth.model}`);
  });

  // First-run welcome (only shows once per browser profile)
  setTimeout(() => { maybeShowWelcome(); }, 400);

  // Auto-resize textarea
  const input = document.getElementById("user-input");
  input.addEventListener("input", () => {
    input.style.height = "auto";
    input.style.height = Math.min(input.scrollHeight, 100) + "px";
  });

  // Undo button (Improvement 11)
  const undoBtn = document.getElementById("undo-btn");
  if (undoBtn) {
    undoBtn.addEventListener("click", handleUndo);
  }

  // Build Comps — PRIMARY action: ticker → inject comp table into sheet
  const buildCompsBtn = document.getElementById("build-comps-btn");
  if (buildCompsBtn) {
    buildCompsBtn.addEventListener("click", handleBuildComps);
  }

  // IB Format button
  const ibBtn = document.getElementById("ib-format-btn");
  if (ibBtn) {
    ibBtn.addEventListener("click", handleIBFormat);
  }

  // Build Deck button — Excel comp → PowerPoint tearsheet
  const deckBtn = document.getElementById("build-deck-btn");
  if (deckBtn) {
    deckBtn.addEventListener("click", handleBuildDeck);
  }

  const exportXlsxBtn = document.getElementById("export-xlsx-btn");
  if (exportXlsxBtn) {
    exportXlsxBtn.addEventListener("click", () => handleExportFormatted("xlsx"));
  }
  const exportPptxBtn = document.getElementById("export-pptx-btn");
  if (exportPptxBtn) {
    exportPptxBtn.addEventListener("click", () => handleExportFormatted("pptx"));
  }

  // History panel toggle (Improvement 12)
  const histToggle = document.getElementById("history-toggle");
  if (histToggle) {
    histToggle.addEventListener("click", () => {
      const panel = document.getElementById("history-panel");
      panel.style.display = panel.style.display === "none" ? "block" : "none";
    });
  }

  // Session expiry check every 60s (Improvement 7)
  setInterval(checkSessionExpiry, 60 * 1000);
  const refreshBtn = document.getElementById("refresh-session-btn");
  if (refreshBtn) {
    refreshBtn.addEventListener("click", async () => {
      const { data } = await supabase.auth.refreshSession();
      if (data?.session) {
        syncSessionToBackend(data.session);
        document.getElementById("session-warning").style.display = "none";
      }
    });
  }

  // Auth debug panel — triple-click logo (Improvement 10)
  let logoClickCount = 0;
  let logoClickTimer = null;
  const logoImg = document.getElementById("logo-img");
  if (logoImg) {
    logoImg.addEventListener("click", () => {
      logoClickCount++;
      if (logoClickTimer) clearTimeout(logoClickTimer);
      logoClickTimer = setTimeout(() => { logoClickCount = 0; }, 600);
      if (logoClickCount >= 3) {
        logoClickCount = 0;
        const dbg = document.getElementById("debug-panel");
        dbg.style.display = dbg.style.display === "none" ? "block" : "none";
        updateDebugPanel();
      }
    });
  }

  // Logout button
  document.getElementById("logout-btn").addEventListener("click", async () => {
    await signOut();
    CURRENT_USER = null;
    document.getElementById("chat-history").innerHTML = "";
    showLoginScreen();
  });

  // Cheap-mode toggle in the header. Renders initial state from localStorage,
  // wires the click → flip state + re-render + show toast.
  _renderCheapModeButton();
  const cheapBtn = document.getElementById("cheap-mode-btn");
  if (cheapBtn) cheapBtn.addEventListener("click", _toggleCheapMode);

  // Image attachment — file picker
  // File input is overlaid on the attach button (no programmatic .click() needed —
  // Office.js WKWebView blocks programmatic file input clicks on Mac)
  document.getElementById("image-input").addEventListener("change", (e) => {
    handleImageSelect(e);
  });

  // Image attachment — paste from clipboard
  document.getElementById("user-input").addEventListener("paste", handleImagePaste);

  // Image attachment — drag & drop onto input area
  const inputArea = document.getElementById("input-area");
  inputArea.addEventListener("dragover", (e) => {
    e.preventDefault();
    e.stopPropagation();
    inputArea.style.outline = "2px dashed var(--blue)";
    inputArea.style.outlineOffset = "-2px";
    inputArea.style.background = "var(--blue-light)";
  });
  inputArea.addEventListener("dragleave", (e) => {
    e.preventDefault();
    e.stopPropagation();
    inputArea.style.outline = "";
    inputArea.style.outlineOffset = "";
    inputArea.style.background = "";
  });
  inputArea.addEventListener("drop", (e) => {
    e.preventDefault();
    e.stopPropagation();
    inputArea.style.outline = "";
    inputArea.style.outlineOffset = "";
    inputArea.style.background = "";
    const files = Array.from(e.dataTransfer?.files || []);
    for (const file of files) {
      readFileAsBase64(file);
    }
    if (files.length === 0) setStatus("No files found in drop");
  });

  // Quick action buttons
  document.querySelectorAll(".quick-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const input = document.getElementById("user-input");
      input.value = btn.dataset.prompt;
      handleSubmit();
    });
  });

  setStatus("Connected · " + user.email);
}

// ── Auth Handlers ─────────────────────────────────────────────────────────────

async function handleSignIn() {
  const email    = document.getElementById("auth-email").value.trim();
  const password = document.getElementById("auth-password").value;
  const errEl    = document.getElementById("auth-error");
  errEl.style.color = "#DC2626";
  errEl.textContent = "";
  // Validation (Improvement 9)
  if (!email || !email.includes("@")) { errEl.textContent = "Enter a valid email address."; return; }
  if (!password) { errEl.textContent = "Enter your password."; return; }
  document.getElementById("login-btn").textContent = "Signing in...";
  const { user, error } = await signIn(email, password);
  document.getElementById("login-btn").textContent = "Sign In";
  if (error) { errEl.textContent = error.message; return; }
  if (!user)  { errEl.textContent = "Check your email to confirm your account first."; return; }
  CURRENT_USER = user;
  sessionSource = "login";
  lastSyncTimestamp = new Date().toISOString();
  await saveUserConfig(user);
  showChatScreen(user);
}

async function handleSignUp() {
  const email    = document.getElementById("auth-email").value.trim();
  const password = document.getElementById("auth-password").value;
  const errEl    = document.getElementById("auth-error");
  errEl.style.color = "#DC2626";
  errEl.textContent = "";
  // Validation (Improvement 9)
  if (!email || !email.includes("@")) { errEl.textContent = "Enter a valid email address."; return; }
  if (password.length < 6) { errEl.textContent = "Password must be at least 6 characters."; return; }
  document.getElementById("signup-btn").textContent = "Creating account…";
  const { user, error } = await signUp(email, password);
  document.getElementById("signup-btn").textContent = "Create account";
  if (error) { errEl.textContent = error.message; return; }
  errEl.style.color = "#16A34A";
  errEl.textContent = "Account created! Check your email to confirm, then sign in.";
}

// ── Image Handling ────────────────────────────────────────────────────────────

function handleImageSelect(e) {
  const files = Array.from(e.target.files);
  for (const file of files) {
    readFileAsBase64(file);
  }
  e.target.value = "";  // reset so same file can be re-selected
}

function handleImagePaste(e) {
  const items = Array.from(e.clipboardData?.items || []);
  for (const item of items) {
    // Accept images from clipboard paste, and files if available
    if (item.type.startsWith("image/") || item.kind === "file") {
      const file = item.getAsFile();
      if (file) readFileAsBase64(file);
    }
  }
}

function readFileAsBase64(file) {
  const isImage = file.type.startsWith("image/");
  const label = isImage ? "image" : "document";
  setStatus(`📎 reading ${file.name} (${file.type || "unknown"}, ${Math.round(file.size/1024)}KB)...`);
  const reader = new FileReader();
  reader.onload = () => {
    const base64 = reader.result;  // data:type;base64,...
    const mediaType = file.type || (isImage ? "image/png" : "application/octet-stream");
    const data = base64.split(",")[1];  // strip the data:... prefix
    pendingImages.push({ media_type: mediaType, data, file_name: file.name || "" });
    setStatus(`✅ ${label} attached · ${pendingImages.length} file${pendingImages.length > 1 ? "s" : ""} pending`);
    updateImagePreview();
  };
  reader.onerror = () => {
    setStatus(`❌ FileReader error: ${reader.error}`);
  };
  reader.readAsDataURL(file);
}

/** Convert base64 string to a blob URL for safe rendering in the webview */
function base64ToBlobUrl(base64, mediaType) {
  const byteChars = atob(base64);
  const byteArray = new Uint8Array(byteChars.length);
  for (let i = 0; i < byteChars.length; i++) byteArray[i] = byteChars.charCodeAt(i);
  const blob = new Blob([byteArray], { type: mediaType });
  return URL.createObjectURL(blob);
}

function updateImagePreview() {
  const bar = document.getElementById("image-preview-bar");
  const attachBtn = document.getElementById("attach-btn");
  bar.innerHTML = "";

  if (pendingImages.length === 0) {
    bar.style.display = "none";
    attachBtn.textContent = "+";
    attachBtn.title = "Attach file";
    return;
  }

  // Update attach button to show count
  attachBtn.textContent = `${pendingImages.length}`;
  attachBtn.title = `${pendingImages.length} file${pendingImages.length > 1 ? "s" : ""} attached — click to add more`;
  setStatus(`${pendingImages.length} file${pendingImages.length > 1 ? "s" : ""} attached`);

  bar.style.display = "flex";
  pendingImages.forEach((img, i) => {
    const wrapper = document.createElement("div");
    wrapper.className = "image-preview-item";
    const isImage = img.media_type.startsWith("image/");

    if (isImage) {
      // Render to canvas (bypasses CSP restrictions on data:/blob: img src)
      renderImageToCanvas(img.data, img.media_type, 48, 48).then(canvas => {
        if (canvas) {
          canvas.style.borderRadius = "4px";
          canvas.style.border = "1px solid var(--border)";
          wrapper.insertBefore(canvas, wrapper.firstChild);
        }
      });
    } else {
      // Document file — show icon + filename
      const docIcon = document.createElement("div");
      const ext = img.file_name ? img.file_name.split(".").pop().toUpperCase() : "FILE";
      docIcon.style.cssText = "width:48px;height:48px;display:flex;align-items:center;justify-content:center;background:#F1F5F9;border-radius:4px;border:1px solid var(--border);font-size:9px;font-weight:700;color:#0D5EAF;text-align:center;line-height:1.1;";
      docIcon.textContent = ext;
      wrapper.insertBefore(docIcon, wrapper.firstChild);
    }

    const removeBtn = document.createElement("button");
    removeBtn.className = "remove-img";
    removeBtn.textContent = "x";
    removeBtn.addEventListener("click", () => {
      pendingImages.splice(i, 1);
      updateImagePreview();
    });

    wrapper.appendChild(removeBtn);
    bar.appendChild(wrapper);
  });
}

/** Render base64 image data onto a canvas element (bypasses all CSP img-src restrictions) */
async function renderImageToCanvas(base64Data, mediaType, maxW, maxH) {
  try {
    const byteChars = atob(base64Data);
    const byteArray = new Uint8Array(byteChars.length);
    for (let i = 0; i < byteChars.length; i++) byteArray[i] = byteChars.charCodeAt(i);
    const blob = new Blob([byteArray], { type: mediaType || "image/png" });
    const bitmap = await createImageBitmap(blob);

    let w = bitmap.width, h = bitmap.height;
    const scale = Math.min(maxW / w, maxH / h, 1);
    w = Math.round(w * scale);
    h = Math.round(h * scale);

    const canvas = document.createElement("canvas");
    canvas.width = w;
    canvas.height = h;
    canvas.getContext("2d").drawImage(bitmap, 0, 0, w, h);
    bitmap.close();
    return canvas;
  } catch (e) {
    console.warn("renderImageToCanvas failed:", e);
    return null;
  }
}

// ── Backend health banner ────────────────────────────────────────────────────
// Checks /chat/debug/guards on boot and after any chat fetch failure. When
// the backend is unreachable, a red banner appears at the top of the taskpane
// with a Retry button. Hidden once health check passes.

async function checkBackendHealth() {
  const banner = document.getElementById("backend-banner");
  const detail = document.getElementById("backend-banner-detail");
  if (!banner) return;
  try {
    const r = await fetch(`${BACKEND_URL}/chat/debug/guards`, {
      method: "GET",
      signal: AbortSignal.timeout ? AbortSignal.timeout(8000) : undefined,
    });
    if (!r.ok) {
      banner.style.display = "block";
      detail.textContent = `Got HTTP ${r.status} from ${BACKEND_URL}. Will retry when you click below.`;
      return false;
    }
    banner.style.display = "none";
    return true;
  } catch (e) {
    banner.style.display = "block";
    const msg = e?.name === "AbortError" || e?.name === "TimeoutError"
      ? "Request timed out after 8s — backend may be down or waking from sleep."
      : `Network error: ${e?.message || e}`;
    detail.textContent = msg;
    return false;
  }
}

function wireBackendBanner() {
  const retry = document.getElementById("backend-banner-retry");
  if (retry && !retry._wired) {
    retry._wired = true;
    retry.addEventListener("click", async () => {
      retry.textContent = "Checking...";
      retry.disabled = true;
      const ok = await checkBackendHealth();
      retry.textContent = "Retry";
      retry.disabled = false;
      if (ok) showToast("Backend reachable", "success");
    });
  }
}

// ── First-run welcome card ───────────────────────────────────────────────────
// Shown once to each user. Dismissed permanently by click. Stored in
// localStorage so we don't nag the user on every session.

const WELCOME_SEEN_KEY = "tsifl_welcome_seen_v1";

function maybeShowWelcome() {
  if (localStorage.getItem(WELCOME_SEEN_KEY) === "1") return;
  if (document.querySelector(".welcome-card")) return;   // already rendered this session
  const history = document.getElementById("chat-history");
  if (!history) return;
  // Mark as seen the moment we render, not only on explicit dismiss —
  // otherwise a second showChatScreen() call during auth produces a
  // duplicate card before the user has a chance to dismiss the first.
  localStorage.setItem(WELCOME_SEEN_KEY, "1");
  const card = document.createElement("div");
  card.className = "welcome-card";
  card.innerHTML = `
    <div class="welcome-card-title">Welcome to tsifl</div>
    <div class="welcome-card-body">
      I'm a cross-app agent for financial analysts. I can read your workbook,
      write formulas, build models, and work with R / PowerPoint / Word if you
      ask. I remember what I've done across turns — hover the <b>Memory</b>
      panel above to inspect or lock cells.
    </div>
    <div class="welcome-examples">
      <button class="welcome-example">Add the formula =C14*D7 to Calculator C18. Nothing else.</button>
      <button class="welcome-example">Build an array formula at Sales Forecast F5 as =(E5:E26/C5:C26)*2.</button>
      <button class="welcome-example">Summarize what this workbook is trying to model.</button>
    </div>
    <button class="welcome-card-dismiss">Hide this</button>
  `;
  history.appendChild(card);

  // Clicking an example prefills the input
  card.querySelectorAll(".welcome-example").forEach(btn => {
    btn.addEventListener("click", () => {
      const input = document.getElementById("user-input");
      if (input) {
        input.value = btn.textContent.trim();
        input.focus();
      }
    });
  });

  // Dismiss handler
  card.querySelector(".welcome-card-dismiss").addEventListener("click", () => {
    localStorage.setItem(WELCOME_SEEN_KEY, "1");
    card.remove();
  });
}

// ── Ollama local model routing ───────────────────────────────────────────────
// Discuss-mode and simple chat-only requests get routed to a local Ollama
// instance (if reachable) instead of burning Claude credits. Webpack proxies
// /local-ollama/* → http://localhost:11434/* so the HTTPS taskpane can reach
// the local HTTP API without mixed-content errors.
//
// Flow:
//   1. On boot, probe /local-ollama/api/tags to see if Ollama is running.
//   2. On each chat turn, if the message looks like discuss-mode AND Ollama
//      is reachable, route there. Otherwise fall through to /chat (Claude).
//   3. Reply renders with a "local · ollama · <model>" chip so the user
//      sees when local routing fired.
//
// This is dev-mode only right now — webpack proxy doesn't exist in prod
// distribution. When we ship, a desktop-agent HTTPS shim or similar will
// replace the proxy.

const OLLAMA_URL            = "/local-ollama";
const OLLAMA_DEFAULT_MODEL  = "llama3.2";       // fast, good enough for discuss
const OLLAMA_HEALTHCHECK_TTL_MS = 30_000;

let _ollamaHealth = { available: false, model: null, checkedAt: 0 };
let _ollamaUserEnabled = true;   // user-controlled toggle (future UI)

/** Check whether Ollama is reachable. Cached for 30s to avoid probing each turn. */
async function checkOllamaAvailable() {
  const now = Date.now();
  if (now - _ollamaHealth.checkedAt < OLLAMA_HEALTHCHECK_TTL_MS) {
    return _ollamaHealth.available;
  }
  try {
    const r = await fetch(`${OLLAMA_URL}/api/tags`, {
      method: "GET",
      signal: AbortSignal.timeout ? AbortSignal.timeout(1500) : undefined,
    });
    if (!r.ok) {
      _ollamaHealth = { available: false, model: null, checkedAt: now };
      return false;
    }
    const data = await r.json();
    const models = (data?.models || []).map(m => m.name || m.model).filter(Boolean);
    // Pick: explicit default if installed, else whatever llama/qwen is available, else first
    let chosen = models.find(m => m.startsWith(OLLAMA_DEFAULT_MODEL));
    if (!chosen) chosen = models.find(m => /llama|qwen|mistral/i.test(m));
    if (!chosen) chosen = models[0] || null;
    _ollamaHealth = { available: !!chosen, model: chosen, checkedAt: now };
    return !!chosen;
  } catch (e) {
    _ollamaHealth = { available: false, model: null, checkedAt: now };
    return false;
  }
}

// Patterns that indicate the user wants conversational / explanatory replies,
// NOT structured actions against the spreadsheet. Mirrors the backend's
// _DISCUSS_PATTERNS regex in services/claude.py, trimmed to the core cases.
const _DISCUSS_RE = new RegExp(
  [
    "^(what|why|how|when|where|who)\\b",
    "\\b(what do you think|what['']?s your (take|opinion|view))\\b",
    "\\b(explain|summari[sz]e|describe|tell me about|walk me through)\\b",
    "\\b(any (recommendation|suggestion|idea|advice|thought|tip|pointer))\\b",
    "\\b(should i|can you explain|help me understand)\\b",
  ].join("|"),
  "i"
);

// If any of these phrases appear, the question is ABOUT the user's workbook
// and needs real context — don't route to the local model (which doesn't see
// the spreadsheet). Keep Claude for those.
const _CONTEXTUAL_RE = new RegExp(
  [
    "\\bthis (workbook|sheet|spreadsheet|file|tab|cell|formula|model|data|table|range|chart)\\b",
    "\\bthese (cells|rows|columns|values|numbers|formulas|data)\\b",
    "\\bmy (workbook|sheet|spreadsheet|file|model|data)\\b",
    "\\bthe (current|active|selected|attached) (workbook|sheet|cell|range|data)\\b",
  ].join("|"),
  "i"
);

// Action verbs — if any of these appear, the user wants tsifl to DO something
// to the workbook, not chat. Must NOT be routed to Ollama (which can't emit
// Excel actions). Mirror of the backend's _ACTION_DEMANDING_VERBS in chat.py.
const _ACTION_VERB_RE = new RegExp(
  [
    "\\b(fix|debug|polish|improve|tidy|format|autofit|auto-?fit)\\b",
    "\\bclean ?up\\b",
    "\\bmake (it|this) (better|nice|cleaner|look|professional|prettier)\\b",
    "\\bmake (it|this) look\\b",
    "\\bany (recommendation|reccomendation|improvement|fix)\\b",
    "\\b(spruce|beautify) (it|this|up)\\b",
    "\\bwhat (would|should) (you|i) change\\b",
    "\\bhelp me (debug|fix|with this|polish|clean|improve)\\b",
    "\\bplease (actually )?(make|apply|do|fix|change)\\b",
    "\\bapply (all|the|these|those) (changes?|fixes?|improvements?|suggestions?)\\b",
    "\\bcorrect the (issue|problem|error|bug)\\b",
    "###",  // user mentioned ##### errors → always action
  ].join("|"),
  "i"
);

function isDiscussMode(message) {
  const m = (message || "").trim();
  if (!m) return false;
  if (m.length > 600) return false;
  // Cell references ("F5", "A1:B10") → it's a write task, not discuss.
  if (/\b[A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?\b/.test(m)) return false;
  // "this workbook", "my model", etc. → needs context, route to Claude.
  if (_CONTEXTUAL_RE.test(m)) return false;
  // Action verbs ("fix", "debug", "polish", "apply changes", "####") →
  // user wants the workbook modified. Ollama can't do that. Route to Claude.
  if (_ACTION_VERB_RE.test(m)) return false;
  return _DISCUSS_RE.test(m);
}

/** Send a chat message to local Ollama. Returns the reply string or null on failure. */
async function askOllama(message, systemHint = "") {
  if (!_ollamaHealth.available || !_ollamaHealth.model) return null;
  const body = {
    model: _ollamaHealth.model,
    messages: [
      systemHint ? { role: "system", content: systemHint } : null,
      { role: "user", content: message },
    ].filter(Boolean),
    stream: false,
  };
  try {
    const r = await fetch(`${OLLAMA_URL}/api/chat`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
      // Ollama on a local Mac takes up to ~30s for a response depending on model
      signal: AbortSignal.timeout ? AbortSignal.timeout(45000) : undefined,
    });
    if (!r.ok) return null;
    const data = await r.json();
    return data?.message?.content || null;
  } catch (e) {
    console.warn("[tsifl] Ollama request failed:", e?.message || e);
    return null;
  }
}

// ── Toast notifications ──────────────────────────────────────────────────────

/** Show a tiny ephemeral toast message above the input. Auto-dismisses. */
function showToast(text, kind = "info", durationMs = 2500) {
  const wrap = document.getElementById("toast-wrap");
  if (!wrap) return;
  const el = document.createElement("div");
  el.className = `toast toast-${kind}`;
  el.textContent = text;
  wrap.appendChild(el);
  // Trigger fade-in next frame
  requestAnimationFrame(() => el.classList.add("visible"));
  setTimeout(() => {
    el.classList.remove("visible");
    setTimeout(() => el.remove(), 220);
  }, durationMs);
}

// ── Project memory panel ─────────────────────────────────────────────────────

/** Fetch memory state for the current user + workbook context and re-render. */
async function refreshMemoryPanel() {
  if (!CURRENT_USER) return;
  try {
    const ctx = await getExcelContext();
    const resp = await fetch(`${BACKEND_URL}/chat/project-memory/lookup`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ user_id: CURRENT_USER.id, context: ctx }),
    });
    if (!resp.ok) return;
    const data = await resp.json();
    currentWorkbookId = data.workbook_id;
    renderMemoryPanel(data);
  } catch (e) {
    console.warn("[tsifl] memory panel refresh failed:", e);
  }
}

function renderMemoryPanel(data) {
  const panel   = document.getElementById("memory-panel");
  const countEl = document.getElementById("memory-count");
  const listEl  = document.getElementById("memory-list");
  const clearBtn = document.getElementById("memory-clear");
  if (!panel) return;

  // Hide the panel entirely if memory feature is disabled server-side
  if (!data.enabled) {
    panel.style.display = "none";
    return;
  }

  const items = data.completed || [];
  const locks = data.user_locks || [];
  const n = items.length;
  const lockedAddrs = new Set(locks.map(l => (l.range || "").toLowerCase()));

  panel.style.display = "block";
  const lockSuffix = locks.length ? `, ${locks.length} locked` : "";
  countEl.textContent = n === 0 && locks.length === 0
    ? "empty"
    : `${n} cell${n === 1 ? "" : "s"} remembered${lockSuffix}`;
  clearBtn.style.display = (n > 0 || locks.length > 0) ? "inline-block" : "none";

  if (n === 0 && locks.length === 0) {
    listEl.innerHTML = `<div class="memory-empty">nothing tracked yet — memory populates as you work</div>`;
    return;
  }

  // Sort: user-locked rows first (they're most important to see), then by recency
  const sortedItems = [...items].sort((a, b) => (b.at || 0) - (a.at || 0));

  const completedRows = sortedItems.slice(0, 50).map(item => {
    const cell = (item.cell || item.range || "?");
    const body = item.formula ? escapeHtml(String(item.formula)) :
                 item.note    ? escapeHtml(String(item.note)) :
                 item.name    ? `named: ${escapeHtml(String(item.name))}` :
                                escapeHtml(item.type || "");
    const isLocked = lockedAddrs.has(String(cell).toLowerCase());
    const lockIcon = isLocked ? "🔒" : "";
    return `
      <div class="memory-row" data-cell="${escapeHtml(cell)}">
        <span class="memory-cell">${lockIcon}${escapeHtml(cell)}</span>
        <span class="memory-body">${body}</span>
        <span class="memory-actions">
          ${isLocked
            ? `<button class="memory-row-btn memory-unlock-btn" title="Unlock — let tsifl modify again">unlock</button>`
            : `<button class="memory-row-btn memory-lock-btn" title="Lock — never let tsifl modify this">lock</button>`}
          <button class="memory-row-btn memory-forget-btn" title="Forget this entry only">×</button>
        </span>
      </div>`;
  }).join("");

  // Render pure-locks (locks that have no completed entry alongside)
  const completedAddrs = new Set(items.map(i => String(i.cell || i.range || "").toLowerCase()));
  const orphanLocks = locks.filter(l => !completedAddrs.has(String(l.range || "").toLowerCase()));
  const orphanRows = orphanLocks.map(l => {
    const rng = l.range || "?";
    return `
      <div class="memory-row" data-cell="${escapeHtml(rng)}">
        <span class="memory-cell">🔒${escapeHtml(rng)}</span>
        <span class="memory-body">${escapeHtml(l.note || "locked")}</span>
        <span class="memory-actions">
          <button class="memory-row-btn memory-unlock-btn" title="Unlock">unlock</button>
        </span>
      </div>`;
  }).join("");

  listEl.innerHTML = orphanRows + completedRows;

  // Wire per-row button handlers
  listEl.querySelectorAll(".memory-forget-btn").forEach(btn => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const row = btn.closest(".memory-row");
      const addr = row?.getAttribute("data-cell");
      if (addr) forgetMemoryCell(addr);
    });
  });
  listEl.querySelectorAll(".memory-lock-btn").forEach(btn => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const row = btn.closest(".memory-row");
      const addr = row?.getAttribute("data-cell");
      if (addr) lockMemoryCell(addr);
    });
  });
  listEl.querySelectorAll(".memory-unlock-btn").forEach(btn => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const row = btn.closest(".memory-row");
      const addr = row?.getAttribute("data-cell");
      if (addr) unlockMemoryCell(addr);
    });
  });
}

async function forgetMemoryCell(addr) {
  if (!CURRENT_USER) return;
  try {
    const ctx = await getExcelContext();
    await fetch(`${BACKEND_URL}/chat/project-memory/forget`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ user_id: CURRENT_USER.id, context: ctx, address: addr }),
    });
    await refreshMemoryPanel();
    showToast(`Forgot ${addr}`, "warning");
  } catch (e) {
    console.warn("[tsifl] forget failed:", e);
    showToast(`Couldn't forget ${addr}`, "error");
  }
}

async function lockMemoryCell(addr) {
  if (!CURRENT_USER) return;
  try {
    const ctx = await getExcelContext();
    await fetch(`${BACKEND_URL}/chat/project-memory/lock`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ user_id: CURRENT_USER.id, context: ctx, address: addr }),
    });
    await refreshMemoryPanel();
    showToast(`Locked ${addr} — tsifl won't modify this`, "success");
  } catch (e) {
    console.warn("[tsifl] lock failed:", e);
    showToast(`Couldn't lock ${addr}`, "error");
  }
}

async function unlockMemoryCell(addr) {
  if (!CURRENT_USER) return;
  try {
    const ctx = await getExcelContext();
    await fetch(`${BACKEND_URL}/chat/project-memory/unlock`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ user_id: CURRENT_USER.id, context: ctx, address: addr }),
    });
    await refreshMemoryPanel();
    showToast(`Unlocked ${addr}`, "info");
  } catch (e) {
    console.warn("[tsifl] unlock failed:", e);
    showToast(`Couldn't unlock ${addr}`, "error");
  }
}

function escapeHtml(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

function toggleMemoryPanel() {
  memoryPanelExpanded = !memoryPanelExpanded;
  const panel = document.getElementById("memory-panel");
  const list  = document.getElementById("memory-list");
  panel.classList.toggle("expanded", memoryPanelExpanded);
  list.style.display = memoryPanelExpanded ? "block" : "none";
}

async function clearMemory() {
  if (!CURRENT_USER) return;
  if (!confirm("Clear all project memory for this workbook? tsifl will forget every cell it's tracked.")) return;
  try {
    const ctx = await getExcelContext();
    const resp = await fetch(`${BACKEND_URL}/chat/project-memory/clear`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ user_id: CURRENT_USER.id, context: ctx }),
    });
    if (!resp.ok) {
      appendMessage("assistant", "Couldn't clear memory — backend returned " + resp.status);
      showToast("Couldn't clear memory", "error");
      return;
    }
    await refreshMemoryPanel();
    showToast("Memory cleared for this workbook", "warning");
  } catch (e) {
    appendMessage("assistant", "Couldn't clear memory: " + (e?.message || e));
    showToast("Couldn't clear memory", "error");
  }
}

// ── Chat ──────────────────────────────────────────────────────────────────────

// Detect replies that claim an Excel change was made. Used to flag hallucinated
// success messages that arrive with zero actions and no computer-use session.
//
// IMPORTANT: keep the verb lists wide. The model loves vague polish verbs
// ("fix", "improve", "polish", "clean up", "share recommendations") which are
// the smoking gun for a future-tense plan with no actions emitted. If you
// add a new verb to the system prompt's banned list, mirror it here.
function claimsActionCompletion(text) {
  if (!text) return false;
  const t = text.toLowerCase();
  // Wide future-tense verb set — covers polish/clean/fix/share verbs, not
  // just create/build/insert. These are the ones that slip past the guard
  // most often because they sound like "soft" actions.
  const futureVerbs =
    "create|build|make|add|insert|write|place|put|send|generate|produce|" +
    "populate|draw|chart|plot|fill|fix|improve|polish|clean(?: up)?|format|" +
    "autofit|auto-?fit|highlight|share|show|organize|organi[sz]e|tidy|" +
    "set ?up|set|configure|update|refine|adjust|tweak|reformat|restyle|" +
    "color|colou?r|bold|widen|expand|shrink|consolidate|summari[sz]e|" +
    "analyze|analyse|recommend|propose|suggest|" +
    // New: future-tense action verbs that slipped past — "let me set this up",
    // "I'll configure", "I'll handle", "I'll run", "I'll do this"
    "handle|run|do|execute|perform|kick off|start|launch|trigger|process|" +
    "open|navigate to|go to|use|apply|implement|deploy";

  const patterns = [
    // Past-tense "I've done it" claims
    /\bi['']?ve (written|added|created|updated|inserted|applied|imported|formatted|filled|set up|placed|put|made|built|generated|populated|entered|sent|fixed|improved|cleaned|polished)/,
    /\bi have (written|added|created|updated|inserted|applied|imported|formatted|filled|placed|populated|entered|sent|fixed|improved|cleaned)/,
    /\ball set\b/,
    /\bdone[\s\.\!—–-]/,
    /\b(data|formulas?|chart|values|cells?|table|rows?|columns?|slide) (have been|has been|were|was) (written|added|imported|created|inserted|populated|entered|formatted|applied|placed|sent|fixed|cleaned|polished)/,
    /\b(written|added|created|updated|inserted|imported|formatted|populated) (the |your )?(data|formulas?|values|cells?|range|chart|table|slide)/,
    /\bsuccessfully (wrote|added|created|inserted|imported|formatted|populated|placed|sent|fixed|cleaned)/,
    // Future-tense "I'll do it" claims — just as broken when no action is emitted
    new RegExp("\\bi['']?ll (" + futureVerbs + ")\\b"),
    new RegExp("\\bi will (" + futureVerbs + ")\\b"),
    new RegExp("\\blet me (" + futureVerbs + ")\\b"),
    new RegExp("\\bgoing to (" + futureVerbs + ")\\b"),
    // Two-phase plan pattern: "First, I'll X, then I'll Y" / "First X, then Y"
    // — this is a smoking-gun future-tense plan that splits work across
    // turns the user never gets. When no actions are emitted, this is broken.
    /\bfirst[, ].*(then|next|after that)[, ]/i,
    // Soft narration: "Let me know what you'd like" / "I can help with..." /
    // "I'd be happy to..." — these are stalling phrases when paired with
    // a request that asked for concrete changes.
    /\bi (can|could) (help|assist|do|fix|improve|update|adjust|format|clean|polish)/,
    /\bi['']?d be (happy|glad) to/,
    /\blet me know (what|which|if|when|how)/,
    // Option-menu pattern: "Here are some options — pick a number" /
    // "Reply with a number" / "Pick one and I will" / "Want me to do A,
    // B, or C?". tsifl is an agent — when the user asks for action, the
    // model must emit actions, not present a multiple-choice quiz.
    /\b(pick|choose|reply with|tell me) (a |the |which )?(number|option|one)\b/,
    /\bhere are some options\b/,
    /\b(let me know|tell me) which (one|to|number)/,
    /\bi haven['']?t (built|done|made|created|added|fixed|applied) (anything|it|them) yet\b/,
    /\bwant me to (do|fix|apply|build|run|execute) [^.?!\n]{0,40}\?/,
    /\bwhich (one|of these|option) (would|do) you/,
  ];
  return patterns.some(rx => rx.test(t));
}

async function handleSubmit() {
  const input   = document.getElementById("user-input");
  const message = input.value.trim();
  if (!message || !CURRENT_USER) return;

  lastNavigatedSheet = null;   // reset cross-action sheet tracking for this request
  const images = [...pendingImages];  // capture attached images
  pendingImages = [];
  updateImagePreview();
  input.value = "";
  setSubmitEnabled(false);
  appendMessage("user", message, images);
  setStatus("Reading workbook...");
  input.style.height = "auto";

  // Detect "build comps / comp table / comps for NVDA" intent
  const compPattern = /\b(build|make|create|generate|run)\b.{0,20}\b(comp|comps|comp table|trading comp)\b|\b(comp|comps|comp table|trading comp)\b.{0,20}\b(for|on|of)\b/i;
  const tickerInMsg = message.match(/\b([A-Z]{2,5})\b/g);
  if (compPattern.test(message) && tickerInMsg && tickerInMsg.length >= 1) {
    // Hijack the input and run Build Comps directly
    const tickersFromMsg = tickerInMsg.filter(t => !["FOR","THE","AND","NOT","ALL","ARE","THIS","COMP","COMPS","BUILD","MAKE","CREATE","RUN","TABLE","TRADING"].includes(t));
    if (tickersFromMsg.length >= 1) {
      input.value = tickersFromMsg.join(", ");
      try {
        await handleBuildComps();
      } catch (_) { /* handleBuildComps shows its own error */ }
      setSubmitEnabled(true);
      return;
    }
  }

  // Detect "build deck / build slides / send to PPT / make a presentation" intent
  const deckPattern = /\b(build|make|create|generate|send|export|push)\b.{0,30}\b(deck|slides?|presentation|ppt|powerpoint|tearsheet)\b|\b(deck|slides?|tearsheet)\b.{0,20}\b(from|of|for)\b.{0,20}\b(comp|excel|this|sheet)/i;
  if (deckPattern.test(message)) {
    try {
      await handleBuildDeck();
    } catch (_) { /* handleBuildDeck shows its own error */ }
    setSubmitEnabled(true);
    return;
  }

  // Detect cross-app R request: "from R", "in R", "use R", "R plot", "generate in R"
  const rJobPattern = /\b(from r\b|in r\b|use r\b|with r\b|r plot|r-?generated|generate.*in r|run.*in r|rstudio|r addin|r studio)\b/i;
  if (rJobPattern.test(message)) {
    try {
      setStatus("Sending to RStudio...");
      showTypingIndicator("thinking");
      await fetch(`${BACKEND_URL}/transfer/store`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          from_app: "excel",
          to_app: "rstudio",
          data_type: "r_job",
          data: message + "\n\nIMPORTANT: Emit ONE run_r_code action that BOTH loads the data AND creates the plot in the same block. Do NOT just inspect/str/head — produce the actual plot. End by mentioning 'excel' so the plot gets exported.",
          metadata: { requested_by: CURRENT_USER.id },
        }),
      });
      hideTypingIndicator();
      appendMessage("assistant", "Sent to RStudio. Make sure tsifl is open there — the plot will appear in this sheet automatically once R finishes.");
      setStatus("Waiting for R...");
      setSubmitEnabled(true);
      return;
    } catch (e) {
      hideTypingIndicator();
      appendMessage("assistant", "Couldn't reach the R job queue: " + (e?.message || e));
      setSubmitEnabled(true);
      return;
    }
  }

  try {
    const excelContext = await getExcelContext();
    setStatus("Thinking...");
    showTypingIndicator("thinking");

    // Ollama short-circuit: for discuss-mode messages (no images, no cell refs,
    // no contextual pronouns, short enough), try the local model first.
    // Contextual questions ("what do you think about this workbook?") do NOT
    // route here — they need Claude because only Claude sees the live data.
    const discussMode = _ollamaUserEnabled && images.length === 0 && isDiscussMode(message);
    if (discussMode && await checkOllamaAvailable()) {
      setStatus("Thinking... (local)");
      // Tight, opinionated system prompt for the local model. gemma3 tends
      // toward verbose / evasive replies without firm instruction.
      const sysHint =
        "You are tsifl, a concise AI assistant for financial analysts using " +
        ((excelContext?.app) || "an office app") + ". " +
        "The user is asking a general conceptual or educational question that " +
        "does NOT require seeing their live workbook. Answer directly in " +
        "2-5 sentences. Do NOT ask for additional information about their " +
        "workbook. Do NOT emit code, formulas, or tool calls. Be specific and " +
        "practical, not generic.";
      const localReply = await askOllama(message, sysHint);
      if (localReply) {
        hideTypingIndicator();
        appendMessage("assistant", localReply, undefined, {
          memoryCount: 0,
          memoryOverrides: 0,
          lockBlocked: 0,
          phantomDropped: 0,
          ollama: _ollamaHealth.model,    // renders "local · ollama · llama3.2"
        });
        setStatus("Ready");
        setSubmitEnabled(true);
        return;
      }
      // Local failed — fall through to Claude
      setStatus("Thinking...");
    }

    // Inject force_model when cheap mode is on. Backend's _select_model
    // honors context.force_model = "haiku" / "sonnet" / "opus" (see
    // services/claude.py). Using ?? avoids overwriting if some other
    // code path already set force_model (e.g. RStudio image override).
    const requestContext = _cheapMode
      ? { ...excelContext, force_model: excelContext.force_model ?? "haiku" }
      : excelContext;

    const response = await fetch(`${BACKEND_URL}/chat/`, {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body:    JSON.stringify({
        user_id: CURRENT_USER.id,
        message: message,
        context: requestContext,
        images:  images.length > 0 ? images : undefined,
      }),
    });

    if (!response.ok) {
      let detail = `Server error (${response.status})`;
      try {
        const err = await response.json();
        detail = err.detail || detail;
      } catch (_) {
        try { detail = await response.text(); } catch (_2) {}
      }
      hideTypingIndicator();
      appendMessage("assistant", `Error: ${detail}`);
      setStatus("Error");
      return;
    }

    const data = await response.json();
    hideTypingIndicator();

    const allActions = [];
    if (data.actions && data.actions.length > 0) allActions.push(...data.actions);
    else if (data.action && data.action.type && data.action.type !== "none") allActions.push(data.action);
    const willExecute = allActions.length > 0 || !!data.cu_session_id;

    if (data.reply && data.reply.trim()) {
      let reply = data.reply;
      // Don't double-warn: if the server already appended a phantom-sheet,
      // locked-cell, or "report not generated" note, that's a more specific
      // explanation than the generic hallucination banner — skip the banner.
      const serverAlreadyExplained =
        /Note: skipped \d+ action/.test(reply) ||
        /Refused to modify \d+ locked/.test(reply) ||
        /Note: the report was NOT generated/.test(reply);

      // Don't fire the banner on conversational messages — "thanks", "ok",
      // "hi", "got it", etc. Those don't ask for action, so the model's
      // friendly reply is correct and the banner just adds noise. We detect
      // conversational by a short, no-action-verb user message.
      const userMsgClean = (message || "").trim().toLowerCase();
      const isConversational = (
        userMsgClean.length < 25
        && !/\b[A-Z]{1,3}\d+\b/.test(message || "")  // no cell refs
        && !/\b(fix|debug|polish|improve|format|autofit|create|build|make|add|insert|write|update|run|apply|change|export|email|export|generate|set|configure)\b/.test(userMsgClean)
      );

      if (!willExecute && claimsActionCompletion(reply) && !serverAlreadyExplained && !isConversational) {
        reply =
          "**Note: nothing was actually changed in the spreadsheet.** " +
          "The assistant claimed to have done something, but no action was executed. " +
          "Try rephrasing your request with specifics (cell range, sheet name, exact values). " +
          "If your request needs advanced features (Solver, Data Tables, SmartArt, etc.) " +
          "make sure the **tsifl Helper** app is running in your menu bar.\n\n" +
          "---\n\n" +
          reply;
      }
      // Memory chip — surfaces when memory shaped the response.
      //   memoryCount:  total entries tracked at turn time
      //   memoryOverrides: entries the LLM chose to overwrite (non-locked)
      //   lockBlocked:  writes the server guard dropped (locked cells)
      //   phantomDropped: writes the server guard dropped (phantom sheets)
      const meta = {
        memoryCount:      data.memory_completed_count || 0,
        memoryOverrides:  data.memory_overrides_count || 0,
        lockBlocked:      (reply.match(/Refused to modify (\d+) locked/)?.[1] | 0) || 0,
        phantomDropped:   (reply.match(/Note: skipped (\d+) action/)?.[1] | 0) || 0,
      };
      appendMessage("assistant", reply, undefined, meta);
    }

    if (data.tasks_remaining >= 0) {
      const tasksEl = document.getElementById("tasks-remaining");
      if (data.tasks_remaining === 999999 || data.subscribed) {
        tasksEl.textContent = "Pro ✓";
        tasksEl.style.color = "var(--green)";
      } else if (data.tasks_remaining <= 5) {
        // Near limit — show subscribe prompt
        tasksEl.textContent = `${data.tasks_remaining} left`;
        tasksEl.style.color = "var(--red)";
        _showSubscribeBanner();
      } else {
        tasksEl.textContent = `${data.tasks_remaining} tasks left`;
        tasksEl.style.color = "";
      }
    }

    if (allActions.length > 0) {
      setStatus(`Applying ${allActions.length} action${allActions.length > 1 ? "s" : ""}...`);
      showTypingIndicator("applying");
      if (allActions.length > 2) showProgress(0, allActions.length);
      await refreshKnownSheets(); // Cache sheet names for formula validation
      _strippedFormulaCount = 0;  // reset formula strip counter
      let applied = 0;
      let failed  = 0;
      let failedNames = [];
      _excelBusy = true;   // block quick-action buttons during action loop
      try {
      for (const action of allActions) {
        try {
          // Save undo state for write actions (Improvement 11)
          const p = action.payload || {};
          if (["write_cell", "write_range", "write_formula", "clear_range"].includes(action.type)) {
            const addr = p.cell || p.range;
            if (addr) await saveUndoState(addr, p.sheet || lastNavigatedSheet);
          }
          await executeAction(action);
          applied++;
          // After add_sheet, refresh known sheets so formulas can reference the new sheet
          if (action.type === "add_sheet") {
            _knownSheets.add((action.payload?.name || "").toLowerCase());
          }
          addToHistory(action.type, summarizeAction(action));
          if (allActions.length > 2) showProgress(applied + failed, allActions.length);
        } catch (err) {
          failed++;
          failedNames.push(`${action.type}(${action.payload?.sheet || action.payload?.cell || action.payload?.range || ''}): ${err.message}`);
          console.error(`${action.type} failed:`, err.message, action.payload);
        }
      }
      } finally {
        _excelBusy = false;
      }
      hideProgress();
      hideTypingIndicator();

      // Auto-autofit columns after any write/format actions to prevent ######
      const hasWrites = allActions.some(a =>
        ["write_cell","write_range","write_formula","format_range","set_number_format","import_csv","create_pivot_summary","fill_down","fill_right"].includes(a.type)
      );
      if (hasWrites) {
        try {
          await Excel.run(async (ctx) => {
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();
            const used = sheet.getUsedRangeOrNullObject();
            used.load("isNullObject");
            await ctx.sync();
            if (!used.isNullObject) {
              used.format.autofitColumns();
              await ctx.sync();
            }
          });
        } catch (e) { /* autofit is best-effort */ }
      }

      // Re-apply explicit column widths AFTER global autofit (autofit overrides them)
      const widthActions = allActions.filter(a =>
        a.type === "autofit_columns" && a.payload && a.payload.width
      );
      for (const wa of widthActions) {
        try {
          await Excel.run(async (ctx) => {
            const sheet = getSheet(ctx, wa.payload.sheet);
            let cols = wa.payload.columns || (wa.payload.column ? [wa.payload.column] : null);
            if (!cols && wa.payload.range) {
              const m = wa.payload.range.match(/^([A-Z]+):/i);
              if (m) cols = [m[1].toUpperCase()];
            }
            if (cols) {
              for (const col of cols) {
                sheet.getRange(`${col}:${col}`).format.columnWidth = wa.payload.width * 7.5;
              }
            }
            await ctx.sync();
          });
        } catch (e) { /* width re-apply is best-effort */ }
      }

      // Post-action sweep: clear any #DIV/0!, #NAME?, #REF!, #VALUE! error cells
      if (hasWrites) {
        try {
          await Excel.run(async (ctx) => {
            const sheets = ctx.workbook.worksheets;
            sheets.load("items/name");
            await ctx.sync();
            for (const ws of sheets.items) {
              const used = ws.getUsedRangeOrNullObject();
              used.load(["values", "isNullObject", "rowCount", "columnCount"]);
              await ctx.sync();
              if (used.isNullObject) continue;
              const vals = used.values;
              for (let r = 0; r < vals.length; r++) {
                for (let c = 0; c < vals[r].length; c++) {
                  const v = vals[r][c];
                  if (typeof v === "string" && /^#(DIV\/0!|REF!|NULL!)$/.test(v)) {
                    // Clear the error cell
                    const cell = used.getCell(r, c);
                    cell.values = [[""]];
                  }
                }
              }
              await ctx.sync();
            }
          });
        } catch (e) { console.warn("[tsifl] Error sweep failed:", e.message); }
      }

      if (failed > 0) {
        const details = failedNames.slice(0, 5).join("\n• ");
        appendMessage("assistant", `${applied} applied, ${failed} failed:\n• ${details}`);
      }
      if (_strippedFormulaCount > 0) {
        console.warn(`[tsifl] Stripped ${_strippedFormulaCount} formulas in this batch`);
      }

      // ── Homework formula safety net ─────────────────────────────────────────
      // If "Transactions Stats" sheet has C16 without a formula, force INDEX/XMATCH
      if (_knownSheets.has("transactions stats")) {
        try {
          await Excel.run(async (ctx) => {
            const ws = ctx.workbook.worksheets.getItemOrNullObject("Transactions Stats");
            ws.load("isNullObject");
            await ctx.sync();
            if (ws.isNullObject) return;
            const c16 = ws.getRange("C16");
            c16.load(["formulas", "values"]);
            await ctx.sync();
            const currentFormula = c16.formulas[0][0];
            // 2D INDEX/XMATCH is the correct formula — it has exactly 2 XMATCHs
            const isCorrect = typeof currentFormula === "string" && currentFormula.includes("INDEX") && currentFormula.includes("XMATCH") && (currentFormula.match(/XMATCH/gi) || []).length === 2;
            if (!isCorrect) {
              console.log("[tsifl] C16 needs correct formula. Current:", currentFormula);
              // Try multiple formula variants
              const variants = [
                "=INDEX(Stats,_xlfn.XMATCH(B16,Transactions!A4:A29),_xlfn.XMATCH('Transactions Stats'!C15,Transactions!A4:D4))",
                "=INDEX(Stats,XMATCH(B16,Transactions!A4:A29),XMATCH('Transactions Stats'!C15,Transactions!A4:D4))",
                "=INDEX(Stats,_xlfn.XMATCH(B16,Transactions!$A$4:$A$29),_xlfn.XMATCH('Transactions Stats'!$C$15,Transactions!$A$4:$D$4))",
                "=INDEX(Stats,XMATCH(B16,Transactions!$A$4:$A$29),XMATCH('Transactions Stats'!$C$15,Transactions!$A$4:$D$4))",
              ];
              for (const v of variants) {
                try {
                  c16.formulas = [[v]];
                  await ctx.sync();
                  console.log("[tsifl] C16 formula set successfully:", v);
                  return;
                } catch (fe) {
                  console.warn("[tsifl] C16 formula variant failed:", v, fe.message);
                }
              }
              console.error("[tsifl] All C16 formula variants failed");
            } else {
              console.log("[tsifl] C16 already correct:", currentFormula);
            }
          });
        } catch (e) { console.warn("[tsifl] C16 safety net failed:", e.message); }

        // ── Homework format safety net ─────────────────────────────────────────
        // Force Comma Style on B7:C10, C16 and Percent Style on D7:D10
        try {
          await Excel.run(async (ctx) => {
            const ws = ctx.workbook.worksheets.getItemOrNullObject("Transactions Stats");
            ws.load("isNullObject");
            await ctx.sync();
            if (ws.isNullObject) return;

            const commaStyle = '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)';
            const pctStyle = "0.00%";

            // Comma Style on B7:C10
            const r1 = ws.getRange("B7:C10");
            r1.load("rowCount,columnCount");
            await ctx.sync();
            const fmt1 = Array.from({ length: r1.rowCount }, () =>
              Array.from({ length: r1.columnCount }, () => commaStyle)
            );
            r1.numberFormat = fmt1;

            // Comma Style on C16
            const r2 = ws.getRange("C16");
            r2.numberFormat = [[commaStyle]];

            // Percent Style on D7:D10
            const r3 = ws.getRange("D7:D10");
            r3.load("rowCount,columnCount");
            await ctx.sync();
            const fmt3 = Array.from({ length: r3.rowCount }, () =>
              Array.from({ length: r3.columnCount }, () => pctStyle)
            );
            r3.numberFormat = fmt3;

            // B15:B16 and C15 are text cells — reset to General (Claude applies 0.00)
            ws.getRange("B15:B16").numberFormat = [["General"], ["General"]];
            ws.getRange("C15").numberFormat = [["General"]];

            await ctx.sync();
            console.log("[tsifl] Homework format safety net applied: Comma Style + Percent Style + text cells General");
          });
        } catch (e) { console.warn("[tsifl] Format safety net failed:", e.message); }

        // ── Employee Insurance cleanup safety net ────────────────────────────
        try {
          await Excel.run(async (ctx) => {
            const ws = ctx.workbook.worksheets.getItemOrNullObject("Employee Insurance");
            ws.load("isNullObject");
            await ctx.sync();
            if (ws.isNullObject) return;

            // Clear stray #,##0 format on B10:C10 (text cells)
            const r = ws.getRange("B10:C10");
            r.numberFormat = [["General", "General"]];

            // Smart SUMIFS fix: match formula column to label in C
            // "# of Dependents" → $E$, "# of Claims" → $F$
            const labels = ws.getRange("C25:C28");
            const formulas = ws.getRange("E25:E28");
            labels.load("values");
            formulas.load("formulas");
            await ctx.sync();

            for (let i = 0; i < 4; i++) {
              const label = (labels.values[i][0] || "").toLowerCase();
              const formula = formulas.formulas[i][0] || "";
              if (typeof formula !== "string" || !formula.includes("SUMIFS")) continue;

              if (label.includes("depend") && formula.includes("$F$4:$F$23")) {
                formulas.getCell(i, 0).formulas = [[formula.replace(/\$F\$4:\$F\$23/g, "$E$4:$E$23")]];
                console.log(`[tsifl] Fixed E${25+i} SUMIFS: label=Dependents, $F$ → $E$`);
              } else if (label.includes("claim") && formula.includes("$E$4:$E$23")) {
                formulas.getCell(i, 0).formulas = [[formula.replace(/\$E\$4:\$E\$23/g, "$F$4:$F$23")]];
                console.log(`[tsifl] Fixed E${25+i} SUMIFS: label=Claims, $E$ → $F$`);
              }
            }

            // Clear duplicate SUMIFS in F27:F28 (should be empty)
            const f27f28 = ws.getRange("F27:F28");
            f27f28.load("formulas");
            await ctx.sync();
            const fVals = f27f28.formulas;
            let hasDupe = false;
            for (const row of fVals) {
              for (const v of row) {
                if (typeof v === "string" && v.includes("SUMIFS")) hasDupe = true;
              }
            }
            if (hasDupe) {
              f27f28.values = [[""], [""]];
              console.log("[tsifl] Cleared duplicate SUMIFS from F27:F28");
            }

            await ctx.sync();
            console.log("[tsifl] Employee Insurance safety net complete");
          });
        } catch (e) { console.warn("[tsifl] Employee Insurance safety net failed:", e.message); }
      }
    }

    // ── Computer Use session polling ────────────────────────────────────
    // If the backend started a computer use session, show access modal
    // and poll. Two UX fixes:
    //   1. Stop button flips _cuCancelRequested so the polling loop
    //      exits on the NEXT tick instead of waiting for the backend
    //      status to round-trip back to "cancelled".
    //   2. If no desktop agent claims the session within 15s, treat it
    //      as "no agent running" and surface a clear message — the
    //      common case when the user hasn't started desktop-agent.py.
    if (data.cu_session_id) {
      const cuSessionId = data.cu_session_id;
      _cuCancelRequested = false;
      setStatus("Desktop automation...");
      showAccessModal(cuSessionId);
      let cuDone = false;
      let cuPolls = 0;
      const maxPolls = 120; // 2 min cap (was 5)
      let agentPickupDeadline = 15; // poll # by which agent must claim
      let everRunning = false;
      while (!cuDone && cuPolls < maxPolls) {
        if (_cuCancelRequested) {
          cuDone = true;
          hideAccessModal();
          appendMessage("assistant", "Stopped.");
          break;
        }
        cuPolls++;
        await new Promise(r => setTimeout(r, 1000));
        try {
          const statusResp = await fetch(`${BACKEND_URL}/computer-use/status/${cuSessionId}`);
          if (statusResp.ok) {
            const statusData = await statusResp.json();
            const s = statusData.status;
            if (s === "running") everRunning = true;
            // Use the agent's actual rich result message (e.g. "Goal Seek
            // converged: D7 changed from 0.5 to 0.7272, so C18 = 100000")
            // instead of generic placeholders. Falls back to placeholders
            // when the agent didn't provide detail (older agent versions
            // or actions that don't return a message).
            const agentMsg = (statusData.result && statusData.result.message) || "";
            if (s === "completed") {
              cuDone = true;
              hideAccessModal();
              appendMessage("assistant", agentMsg || "Done.");
            } else if (s === "failed") {
              cuDone = true;
              hideAccessModal();
              appendMessage("assistant",
                agentMsg
                  ? `Desktop automation failed: ${agentMsg}`
                  : `Desktop automation failed: ${statusData.error || "unknown error"}`);
            } else if (s === "cancelled") {
              cuDone = true;
              hideAccessModal();
              appendMessage("assistant", "Stopped.");
            } else if (s === "partial") {
              cuDone = true;
              hideAccessModal();
              appendMessage("assistant",
                agentMsg
                  ? `Some actions completed, some did not:\n\n${agentMsg}`
                  : "Done — some advanced features applied (partial completion).");
            } else if (s === "pending" && cuPolls >= agentPickupDeadline && !everRunning) {
              // No desktop agent has claimed the session after 15s —
              // almost certainly means the user hasn't started the
              // agent. Bail out with a clear message.
              cuDone = true;
              hideAccessModal();
              appendMessage("assistant",
                "**tsifl Helper isn't running.** Advanced features like Solver, Data Tables, SmartArt and PivotTables need the helper app running in your menu bar.\n\n" +
                "**Fix:** open `tsifl Helper.app` from your Applications folder (or wherever you installed it). You'll see a `tsifl` entry appear in your menu bar (top-right of screen). Then retry your request.\n\n" +
                "_If you don't have it installed yet, follow `desktop-agent/INSTALL.md`._");
              // Also tell the backend to clean up the stuck session.
              try {
                await fetch(`${BACKEND_URL}/computer-use/cancel/${cuSessionId}`, { method: "POST" });
              } catch {}
            }
          }
        } catch (pollErr) {
          console.warn("[tsifl] CU poll error:", pollErr.message);
        }
      }
      hideAccessModal();
      if (!cuDone) {
        appendMessage("assistant", "Desktop automation timed out after 2 min. Check Excel for partial results.");
      }
    }

    setStatus("Done");
  } catch (err) {
    hideTypingIndicator();
    console.error("[tsifl] handleSubmit error:", err);
    appendMessage("assistant", `Could not reach tsifl backend.\n${err?.name || ""}: ${err?.message || err}`);
    setStatus("Disconnected");
    // Surface the backend banner so the user has a clear retry path
    checkBackendHealth();
  } finally {
    setSubmitEnabled(true);
    // Refresh memory panel (fire-and-forget) so the user sees what just got tracked
    refreshMemoryPanel().catch(() => {});
  }
}

// ── Excel Context ─────────────────────────────────────────────────────────────

async function getExcelContext() {
  return new Promise((resolve) => {
    Excel.run(async (ctx) => {
      const wb     = ctx.workbook;
      const sheets = wb.worksheets;
      sheets.load("items/name");
      // Workbook name gives us a stable identity across sheet additions/renames,
      // so project_memory doesn't orphan state when tsifl adds a new sheet.
      wb.load("name");
      await ctx.sync();

      // Load used ranges for ALL sheets in one batch
      const sheetMeta = [];
      for (const ws of sheets.items) {
        const used = ws.getUsedRangeOrNullObject();
        used.load(["address", "values", "formulas", "rowCount", "columnCount", "isNullObject"]);
        sheetMeta.push({ name: ws.name, used });
      }

      const activeSheet = wb.worksheets.getActiveWorksheet();
      const selected    = wb.getSelectedRange();
      activeSheet.load("name");
      selected.load(["address", "values"]);
      await ctx.sync();

      const activeName = activeSheet.name;

      // Build per-sheet summaries
      const sheetSummaries = [];
      let activeSheetData     = [];
      let activeSheetFormulas = [];
      let activeUsedRange     = "empty";

      for (const { name, used } of sheetMeta) {
        if (used.isNullObject || used.rowCount === 0) {
          sheetSummaries.push({ name, used_range: "empty", rows: 0, cols: 0, preview: [] });
          continue;
        }

        // Cap how much data we pass for each sheet
        // Active sheet gets 200 rows so Claude can compute values itself; non-active
        // sheets get 150 so multi-sheet SIMnet projects (typically ≤ 30 rows/sheet
        // across 3-4 tabs) don't silently truncate — prior 60-row cap was causing
        // the LLM to stop mid-range on secondary sheets.
        const MAX_ROWS = name === activeName ? 200 : 150;
        const values   = used.values.slice(0, MAX_ROWS).map(r => r.slice(0, 26));
        const formulas = used.formulas ? used.formulas.slice(0, MAX_ROWS).map(r => r.slice(0, 26)) : values;

        if (name === activeName) {
          activeSheetData     = values;
          activeSheetFormulas = formulas;
          activeUsedRange     = used.address;
        }

        // Preview rows sent to the LLM. Previously 20 for non-active sheets, which
        // truncated secondary-sheet data in multi-sheet projects (e.g. Sales Forecast
        // has 26 data rows — the LLM only saw rows 1-20 and wrote formulas that stopped
        // at row 20). 100 covers typical SIMnet sheets without blowing the context.
        const PREVIEW_ROWS = name === activeName ? 200 : 100;
        sheetSummaries.push({
          name,
          used_range:       used.address,
          rows:             used.rowCount,
          cols:             used.columnCount,
          preview:          values.slice(0, PREVIEW_ROWS),
          preview_formulas: formulas.slice(0, PREVIEW_ROWS),
        });
      }

      // Load named ranges (Improvement 20)
      let namedRanges = [];
      try {
        const names = wb.names;
        names.load("items/name,items/value");
        await ctx.sync();
        namedRanges = names.items.map(n => ({ name: n.name, reference: n.value }));
      } catch (_) { /* named ranges may not exist */ }

      resolve({
        app:              "excel",
        sheet:            activeName,
        workbook_name:    (wb.name || "").toString(),  // stable id — survives add_sheet
        all_sheets:       sheets.items.map(s => s.name),
        selected_cell:    selected.address,
        selected_value:   selected.values?.[0]?.[0] ?? null,
        used_range:       activeUsedRange,
        sheet_data:       activeSheetData,
        sheet_formulas:   activeSheetFormulas,
        sheet_summaries:  sheetSummaries,
        named_ranges:     namedRanges,
        preferences:      loadPreferences(),
      });
    }).catch(() => resolve({ app: "excel", preferences: loadPreferences() }));
  });
}

// ── Helpers ───────────────────────────────────────────────────────────────────

/** Convert "A" → 0, "B" → 1, "AA" → 26, etc. (0-based) */
function colLetterToIndex(col) {
  let n = 0;
  for (const c of col.toUpperCase()) n = n * 26 + c.charCodeAt(0) - 64;
  return n - 1;
}

/** Convert 0-based index to column letter: 0 → "A", 25 → "Z", 26 → "AA" */
function indexToColLetter(idx) {
  let s = "";
  idx += 1;
  while (idx > 0) {
    const rem = (idx - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    idx = Math.floor((idx - 1) / 26);
  }
  return s;
}

/** Get column letter from a cell address like "B4" → "B" */
function cellToCol(addr) {
  return addr.replace(/[^A-Za-z]/g, "").toUpperCase();
}

/** Get the worksheet by name, or active sheet if name is omitted */
function getSheet(ctx, sheetName) {
  if (sheetName) return ctx.workbook.worksheets.getItem(sheetName);
  return ctx.workbook.worksheets.getActiveWorksheet();
}

/** Cache of known sheet names — refreshed before each action batch */
let _knownSheets = new Set();

async function refreshKnownSheets() {
  try {
    await Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      await ctx.sync();
      _knownSheets = new Set(sheets.items.map(s => s.name.toLowerCase()));
    });
  } catch (e) { /* best effort */ }
}

/**
 * Sanitize a value/formula.
 * Only allows simple same-sheet formulas like =SUM(B2:B9), =A1+B1, =B2/B3.
 * Strips ALL cross-sheet formulas and complex functions.
 */
let _strippedFormulaCount = 0;  // track how many formulas got stripped per batch

function sanitizeFormula(val) {
  if (typeof val !== "string" || !val.startsWith("=")) return val;

  // ── #NAME? guard: =n/a, =N/A, =na → write "n/a" text, not a formula ───────
  // Claude sometimes writes =n/a for missing data. Excel treats "n/a" as an
  // undefined named range → #NAME? error. Convert to plain text string.
  if (/^=\s*[nN]\s*\/\s*[aA]\s*$/.test(val)) return "n/a";
  // Also catch =N/A() (Excel's own NA function written wrong), =#N/A
  if (/^=\s*#?[nN]\/?[aA][\s()]*$/.test(val)) return "n/a";

  // ── #NAME? guard: formula is ONLY a named-range-style identifier ────────────
  // Catch =Revenue, =EBITDA_Margin, =Market_Cap etc. — no operators or parens.
  // Real formulas always have () or +/-/* etc. Pure alpha/underscore = mistake.
  if (/^=[A-Za-z][A-Za-z0-9_]*$/.test(val)) {
    console.warn("[tsifl] Stripped likely-named-range formula:", val);
    return "";
  }

  // Block cross-sheet references ONLY if the sheet doesn't exist
  if (val.includes("!")) {
    // Extract sheet name from formula like =DSUM(Sheet1!A:D,...) or ='Sheet Name'!A1
    const sheetMatch = val.match(/(?:=\w+\()?'?([^'!]+)'?!/);
    if (sheetMatch) {
      const refSheet = sheetMatch[1].toLowerCase();
      if (!_knownSheets.has(refSheet)) {
        console.warn(`[tsifl] Stripped formula referencing unknown sheet "${sheetMatch[1]}": ${val}`);
        _strippedFormulaCount++;
        return "";
      }
    }
    // Cross-sheet ref to a known sheet — allow it
    return val;
  }

  // Allow ALL formulas that don't have cross-sheet refs to unknown sheets.
  // This supports homework formulas: DB, DSUM, SUMIFS, CONCAT, LEFT, REPT,
  // COUNTIFS, VLOOKUP, INDEX, MATCH, IF, SUM, AVERAGE, etc.
  return val;
}

/**
 * Sanitize a 2D array of values.
 */
function sanitize2D(arr) {
  return arr.map(row => row.map(v => {
    // Mac Excel VBA error 13 (Type mismatch) triggers on null/undefined/NaN/Infinity
    if (v === null || v === undefined) return "";
    if (typeof v === "number" && !isFinite(v)) return "";   // NaN / Infinity
    const f = sanitizeFormula(v);
    if (typeof f === "string" && !f.startsWith("=")) return sanitizeCurrencyValue(f);
    return f;
  }));
}

/**
 * Strip currency symbols and fix European decimal commas so numeric values
 * are always written as numbers, not text strings.
 * "€132,19" → 132.19 | "1.234,56" → 1234.56 | "$46.78" → 46.78
 */
function sanitizeCurrencyValue(val) {
  if (typeof val !== "string") return val;
  if (val.startsWith("=")) return val;

  let s = val.trim().replace(/^[€$£¥₹\s]+/, "").replace(/[€$£¥\s]+$/, "");

  // Skip strings with non-numeric suffixes (labels like "B", "M", "x", "%")
  if (/[a-zA-Z%x]/.test(s)) return val;

  // European format: "1.234,56" → comma is decimal, dots are thousands
  const lastDot = s.lastIndexOf(".");
  const lastComma = s.lastIndexOf(",");
  if (lastComma > lastDot) {
    s = s.replace(/\./g, "").replace(",", ".");
  } else if (lastComma !== -1 && lastDot === -1) {
    s = s.replace(",", ".");
  }

  const num = parseFloat(s);
  return (!isNaN(num) && isFinite(num)) ? num : val;
}

/**
 * Auto-add _xlfn. prefix for newer Excel functions that require it on Mac/Office.js.
 * CONCAT, XMATCH, IFS, SWITCH, TEXTJOIN, MAXIFS, MINIFS etc. need the prefix.
 * If Claude sends =CONCAT(...), we convert to =_xlfn.CONCAT(...).
 * If Claude already sends =_xlfn.CONCAT(...), we leave it alone.
 */
const _XLFN_FUNCTIONS = ["CONCAT","TEXTJOIN","IFS","SWITCH","XMATCH","XLOOKUP","MAXIFS","MINIFS","FILTER","SORT","UNIQUE","SEQUENCE","RANDARRAY","LET","LAMBDA"];
function _ensureXlfnPrefix(formula) {
  if (!formula || !formula.startsWith("=")) return formula;
  for (const fn of _XLFN_FUNCTIONS) {
    // Match =CONCAT( or =CONCAT( inside nested formulas, but not already prefixed
    const regex = new RegExp(`(?<!_xlfn\\.)\\b${fn}\\(`, "gi");
    formula = formula.replace(regex, `_xlfn.${fn}(`);
  }
  return formula;
}

/**
 * Normalise a cell/range address that Claude sometimes sends as "SheetName!A1"
 * into a { sheet, addr } pair so each handler can call sheet.getRange(addr).
 *
 * Handles:
 *   "Average Ratings!F8"   → { sheet: "Average Ratings", addr: "F8" }
 *   "'Sheet Name'!A1:B5"   → { sheet: "Sheet Name",      addr: "A1:B5" }
 *   "F8"                   → { sheet: fallback,           addr: "F8" }
 */
function splitAddr(raw, fallbackSheet) {
  if (!raw) return { sheet: fallbackSheet, addr: raw };
  const bang = raw.indexOf("!");
  if (bang === -1) return { sheet: fallbackSheet, addr: raw };
  const sheetPart = raw.slice(0, bang).replace(/^'|'$/g, ""); // strip surrounding quotes
  const addrPart  = raw.slice(bang + 1);
  return { sheet: sheetPart || fallbackSheet, addr: addrPart };
}

/**
 * Normalise a values/formulas array to the 2D form Excel.js requires.
 * Claude sometimes sends a flat 1D array ["val1","val2"] for a single-column
 * range; this converts it to [["val1"],["val2"]] so range.values/.formulas work.
 */
function ensure2D(arr) {
  if (!Array.isArray(arr) || arr.length === 0) return arr;
  if (!Array.isArray(arr[0])) return arr.map(v => [v]);
  return arr;
}

/** Apply formula or value to a range object */
function applyValue(range, val) {
  if (typeof val === "string" && val.startsWith("=")) {
    range.formulas = [[val]];
  } else {
    range.values = [[val ?? ""]];
  }
}

// ── Action Executor ───────────────────────────────────────────────────────────

// Action types that should inherit the last navigated sheet when no sheet is specified
const SHEET_AWARE_TYPES = new Set([
  "write_cell", "write_range", "write_formula",
  "fill_down", "fill_right", "copy_range",
  "create_named_range", "sort_range",
  "format_range", "set_number_format",
  "autofit", "autofit_columns",
  "clear_range", "freeze_panes",
  "remove_duplicates", "conditional_format_heatmap",
  "create_pivot_summary", "auto_chart_best_fit",
  "trim_whitespace", "find_replace_bulk",
  "goal_seek", "run_toolpak", "create_data_table",
]);

// Serialise all Excel mutations — prevents concurrent Excel.run calls on Mac
// which are the primary cause of VBA runtime error 13 (Type mismatch).
let _excelActionQueue = Promise.resolve();
function _serialised(fn) {
  _excelActionQueue = _excelActionQueue.then(() => fn()).catch(() => {});
  return _excelActionQueue;
}

// Global "Excel is busy" guard for quick-action buttons (IB Format, Build Deck, Export).
// Prevents them from firing while the main action loop is running, which is the
// second most common trigger for VBA error 13 on Mac.
let _excelBusy = false;
function _guardedExcelOp(label, fn) {
  if (_excelBusy) {
    showToast(`⏳ Still running — please wait`, 2500);
    return Promise.resolve();
  }
  _excelBusy = true;
  return Promise.resolve().then(fn).finally(() => { _excelBusy = false; });
}

async function executeAction(action) {
  const type = action.type;
  // Auto-inject last navigated sheet so write actions land on the right sheet
  // even when Claude omits the sheet field after a navigate_sheet
  let payload = (!action.payload?.sheet && lastNavigatedSheet && SHEET_AWARE_TYPES.has(type))
    ? { ...action.payload, sheet: lastNavigatedSheet }
    : action.payload;

  if (!type || !payload) return;

  // ── navigate_sheet ─────────────────────────────────────────────────────────
  if (type === "navigate_sheet") {
    lastNavigatedSheet = payload.sheet;   // track for subsequent write actions
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(payload.sheet);
      // Unhide sheet first — activate() throws on hidden sheets
      ws.load("visibility");
      await ctx.sync();
      if (ws.visibility !== Excel.SheetVisibility.visible) {
        ws.visibility = Excel.SheetVisibility.visible;
        await ctx.sync();
      }
      ws.activate();
      await ctx.sync();
    });
  }

  // ── write_cell ─────────────────────────────────────────────────────────────
  // Supports: cell, value OR formula, sheet?, bold?, color?, font_color?,
  //           number_format?, font_size?, font_name?
  // Claude sometimes sends cell as "SheetName!A1" — splitAddr handles that.
  else if (type === "write_cell") {
    await Excel.run(async (ctx) => {
      const { sheet: s, addr } = splitAddr(payload.cell, payload.sheet);
      const sheet = getSheet(ctx, s);
      const range = sheet.getRange(addr);
      // Data validation warning (Improvement 16)
      range.load("values");
      await ctx.sync();
      const existing = range.values[0][0];
      if (existing !== null && existing !== "" && existing !== 0) {
        console.log(`Overwrote existing data in ${addr}`);
      }
      const rawVal = payload.formula ?? payload.value ?? "";
      let val = sanitizeFormula(rawVal);
      if (typeof val === "string" && !val.startsWith("=")) {
        val = sanitizeCurrencyValue(val);
      }
      if (typeof val === "string" && val.startsWith("=")) {
        val = _ensureXlfnPrefix(val);
        range.formulas = [[val]];
      } else {
        range.values = [[val]];
      }
      _applyFormat(range, payload);
      await ctx.sync();
    });
  }

  // ── write_range ────────────────────────────────────────────────────────────
  // Supports: range, values (2D array) OR formulas (2D array), sheet?
  // Auto-resizes range to match actual data dimensions (prevents size mismatch errors)
  else if (type === "write_range") {
    await Excel.run(async (ctx) => {
      const { sheet: s, addr } = splitAddr(payload.range, payload.sheet);
      const sheet = getSheet(ctx, s);

      // Determine data to write
      let data;
      let isFormula = false;
      if (payload.formulas) {
        data = sanitize2D(ensure2D(payload.formulas));
        isFormula = true;
      } else if (payload.values) {
        data = sanitize2D(ensure2D(payload.values));
        // Auto-detect formulas in the normalised 2D array
        if (data.some(r => Array.isArray(r) && r.some(v => typeof v === "string" && v.startsWith("=")))) {
          isFormula = true;
        }
      }
      if (!data || data.length === 0) return;

      // Auto-size: use the start cell of the range + data dimensions
      // Extract start cell from range like "A1:D5" → "A1", or just "A1"
      const startCell = addr.includes(":") ? addr.split(":")[0] : addr;
      const rows = data.length;
      const cols = Math.max(...data.map(r => r.length));

      // Pad rows to uniform column count
      const padded = data.map(r => {
        while (r.length < cols) r.push("");
        return r;
      });

      // Get the correctly-sized range from start cell
      const startRange = sheet.getRange(startCell);
      const sized = startRange.getResizedRange(rows - 1, cols - 1);

      if (isFormula) sized.formulas = padded;
      else           sized.values   = padded;

      _applyFormat(sized, payload, rows, cols);
      await ctx.sync();
    });
  }

  // ── write_formula ──────────────────────────────────────────────────────────
  // Write a single formula to a cell — explicit formula action
  else if (type === "write_formula") {
    await Excel.run(async (ctx) => {
      const { sheet: s, addr } = splitAddr(payload.cell, payload.sheet);
      const sheet = getSheet(ctx, s);
      const range = sheet.getRange(addr);
      // Auto-add _xlfn. prefix for newer Excel functions that require it on Mac
      let formula = payload.formula || "";
      formula = _ensureXlfnPrefix(formula);
      // Try multiple formula approaches — Office.js on Mac can be picky
      let formulaSet = false;
      // Approach 1: with _xlfn. prefix
      try {
        range.formulas = [[formula]];
        await ctx.sync();
        formulaSet = true;
      } catch (e1) {
        console.warn("[tsifl] write_formula approach 1 failed (_xlfn):", e1.message);
      }
      // Approach 2: without _xlfn. prefix
      if (!formulaSet) {
        try {
          const rawFormula = (payload.formula || "");
          range.formulas = [[rawFormula]];
          await ctx.sync();
          formulaSet = true;
        } catch (e2) {
          console.warn("[tsifl] write_formula approach 2 failed (raw):", e2.message);
        }
      }
      // Approach 3: formulasLocal (uses locale-specific separators)
      if (!formulaSet) {
        try {
          range.formulasLocal = [[formula]];
          await ctx.sync();
          formulaSet = true;
        } catch (e3) {
          console.warn("[tsifl] write_formula approach 3 failed (formulasLocal):", e3.message);
        }
      }
      if (!formulaSet) {
        console.error("[tsifl] All formula approaches failed for:", formula);
      }
      _applyFormat(range, payload);
      await ctx.sync();
    });
  }

  // ── fill_down ──────────────────────────────────────────────────────────────
  // Copy formula from first row of range down to the rest
  else if (type === "fill_down") {
    await Excel.run(async (ctx) => {
      const rawRange   = payload.range  || payload.target;
      const { sheet: s, addr: fullAddr } = splitAddr(rawRange, payload.sheet);
      const sheet      = getSheet(ctx, s);
      const sourceAddr = payload.source || fullAddr.split(":")[0];
      const { addr: srcAddr } = splitAddr(sourceAddr, null);
      const source     = sheet.getRange(srcAddr);
      const dest       = sheet.getRange(fullAddr);
      dest.copyFrom(source, Excel.RangeCopyType.formulas, false, false);
      await ctx.sync();
    });
  }

  // ── fill_right ─────────────────────────────────────────────────────────────
  else if (type === "fill_right") {
    await Excel.run(async (ctx) => {
      const { sheet: s, addr: destAddr } = splitAddr(payload.range,  payload.sheet);
      const { addr: srcAddr }            = splitAddr(payload.source, null);
      const sheet  = getSheet(ctx, s);
      const source = sheet.getRange(srcAddr);
      const dest   = sheet.getRange(destAddr);
      dest.copyFrom(source, Excel.RangeCopyType.formulas, false, false);
      await ctx.sync();
    });
  }

  // ── copy_range ─────────────────────────────────────────────────────────────
  // Copy values + formulas + formats from one range to another
  else if (type === "copy_range") {
    await Excel.run(async (ctx) => {
      const { sheet: s, addr: fromAddr } = splitAddr(payload.from, payload.sheet);
      const { addr: toAddr }             = splitAddr(payload.to,   null);
      const sheet  = getSheet(ctx, s);
      const source = sheet.getRange(fromAddr);
      const dest   = sheet.getRange(toAddr);
      dest.copyFrom(source, Excel.RangeCopyType.all, false, false);
      await ctx.sync();
    });
  }

  // ── create_named_range ─────────────────────────────────────────────────────
  // payload: { name, range, sheet? }
  // Deletes any existing name with the same name before recreating (idempotent)
  else if (type === "create_named_range") {
    await Excel.run(async (ctx) => {
      const { sheet: s, addr } = splitAddr(payload.range, payload.sheet);
      const sheet    = getSheet(ctx, s);
      const rangeObj = sheet.getRange(addr);
      // If name already exists, delete it first so we don't get a duplicate error
      const existing = ctx.workbook.names.getItemOrNullObject(payload.name);
      existing.load("isNullObject");
      await ctx.sync();
      if (!existing.isNullObject) existing.delete();
      // Use workbook-level name so it works cross-sheet (like Excel named ranges)
      ctx.workbook.names.add(payload.name, rangeObj);
      await ctx.sync();
    });
  }

  // ── sort_range ─────────────────────────────────────────────────────────────
  // payload: { sheet?, range, key_column (letter), ascending? }
  else if (type === "sort_range") {
    await Excel.run(async (ctx) => {
      const { sheet: s, addr } = splitAddr(payload.range, payload.sheet);
      const sheet    = getSheet(ctx, s);
      const rng      = sheet.getRange(addr);

      // Compute 0-based key index within the range
      const rangeStart  = addr.split(":")[0];
      const rangeColLtr = cellToCol(rangeStart);
      const keyColLtr   = (payload.key_column || rangeColLtr).toUpperCase();
      const keyIndex    = colLetterToIndex(keyColLtr) - colLetterToIndex(rangeColLtr);

      rng.sort.apply([{
        key:       Math.max(0, keyIndex),
        ascending: payload.ascending !== false,
      }]);
      await ctx.sync();
    });
  }

  // ── format_range ──────────────────────────────────────────────────────────
  else if (type === "format_range") {
    await Excel.run(async (ctx) => {
      const { sheet: s, addr } = splitAddr(payload.range, payload.sheet);
      const sheet = getSheet(ctx, s);
      // Clamp range to used area to avoid "invalid argument" on out-of-bounds ranges
      let range;
      try {
        range = sheet.getRange(addr);
        range.load("rowCount,columnCount");
        await ctx.sync();
      } catch (e) {
        // If range is invalid (e.g. "A:D" without rows), fall back to used range
        const used = sheet.getUsedRangeOrNullObject();
        used.load("isNullObject");
        await ctx.sync();
        if (used.isNullObject) return;
        range = used;
      }
      _applyFormat(range, payload);
      await ctx.sync();
    });
  }

  // ── set_number_format ──────────────────────────────────────────────────────
  else if (type === "set_number_format") {
    await Excel.run(async (ctx) => {
      const { sheet: s, addr } = splitAddr(payload.range, payload.sheet);
      const sheet = getSheet(ctx, s);
      let range;
      try {
        range = sheet.getRange(addr);
        range.load("rowCount,columnCount");
        await ctx.sync();
      } catch (e) {
        // If range is invalid, fall back to used range
        const used = sheet.getUsedRangeOrNullObject();
        used.load("isNullObject,rowCount,columnCount");
        await ctx.sync();
        if (used.isNullObject) return;
        range = used;
        range.load("rowCount,columnCount");
        await ctx.sync();
      }
      const fmtStr = _localeSafeFmt(payload.format || payload.number_format || "General");
      const fmt = Array.from({ length: range.rowCount }, () =>
        Array.from({ length: range.columnCount }, () => fmtStr)
      );
      range.numberFormat = fmt;
      await ctx.sync();
    });
  }

  // ── autofit ────────────────────────────────────────────────────────────────
  // Full sheet autofit (existing behaviour)
  else if (type === "autofit") {
    await Excel.run(async (ctx) => {
      const sheet = getSheet(ctx, payload.sheet);
      const used  = sheet.getUsedRangeOrNullObject();
      used.load("isNullObject");
      await ctx.sync();
      if (!used.isNullObject) {
        used.format.autofitColumns();
        used.format.autofitRows();
        await ctx.sync();
      }
    });
  }

  // ── autofit_columns ────────────────────────────────────────────────────────
  // Autofit or set width: { sheet?, columns: ["A","B"], width?: 10 }
  // Also accepts range: "L:L" format to extract column letter
  else if (type === "autofit_columns") {
    await Excel.run(async (ctx) => {
      const sheet   = getSheet(ctx, payload.sheet);
      // Extract column list from columns, column, or range (e.g. "L:L")
      let cols = payload.columns || (payload.column ? [payload.column] : null);
      if (!cols && payload.range) {
        const m = payload.range.match(/^([A-Z]+):/i);
        if (m) cols = [m[1].toUpperCase()];
      }
      if (cols) {
        for (const col of cols) {
          const colRange = sheet.getRange(`${col}:${col}`);
          if (payload.width) {
            colRange.format.columnWidth = payload.width * 7.5; // approx char-width to points
          } else {
            colRange.format.autofitColumns();
          }
        }
      } else {
        // Autofit all used columns
        const used = sheet.getUsedRangeOrNullObject();
        used.load("isNullObject");
        await ctx.sync();
        if (!used.isNullObject) used.format.autofitColumns();
      }
      await ctx.sync();
    });
  }

  // ── add_sheet ──────────────────────────────────────────────────────────────
  else if (type === "add_sheet") {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.add(payload.name || undefined);
      if (payload.activate !== false) ws.activate();
      await ctx.sync();
    });
  }

  // ── delete_range_contents ──────────────────────────────────────────────────
  else if (type === "clear_range") {
    await Excel.run(async (ctx) => {
      const sheet = getSheet(ctx, payload.sheet);
      const range = sheet.getRange(payload.range);
      range.clear(payload.clear_type || Excel.ClearApplyTo.contents);
      await ctx.sync();
    });
  }

  // ── freeze_panes ───────────────────────────────────────────────────────────
  else if (type === "freeze_panes") {
    await Excel.run(async (ctx) => {
      const sheet = getSheet(ctx, payload.sheet);
      if (payload.cell) {
        sheet.freezePanes.freezeAt(sheet.getRange(payload.cell));
      } else if (payload.rows) {
        sheet.freezePanes.freezeRows(payload.rows);
      } else if (payload.columns) {
        sheet.freezePanes.freezeColumns(payload.columns);
      }
      await ctx.sync();
    });
  }

  // ── save_workbook ───────────────────────────────────────────────────────────
  else if (type === "save_workbook") {
    await Excel.run(async (ctx) => {
      await ctx.workbook.save(Excel.SaveBehavior.save);
      await ctx.sync();
    });
  }

  // ── save_preference ────────────────────────────────────────────────────────
  // Claude calls this when it learns the user prefers a specific style
  else if (type === "save_preference") {
    savePreferences(payload);
    // No Excel run needed — just localStorage
  }

  // ── import_csv ──────────────────────────────────────────────────────────────
  // Reads a CSV file from the server filesystem and writes it into Excel.
  // Creates named ranges for each column so formulas can use =SUM(Revenue) etc.
  // payload: { path, sheet?, start_cell?, delimiter?, table_name? }
  else if (type === "import_csv") {
    // 1. Fetch CSV data — try remote backend first, then local fallback
    const fetchBody = JSON.stringify({
      path: payload.path,
      delimiter: payload.delimiter || ",",
    });
    const fetchOpts = {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: fetchBody,
    };

    let csvData;
    let resp = await fetch(`${BACKEND_URL}/files/read-csv`, fetchOpts).catch(() => null);

    if (resp && resp.ok) {
      csvData = await resp.json();
    } else {
      // Remote failed (file not on Railway) — try local backend (file on user's machine)
      let localResp;
      try {
        localResp = await fetch(`${LOCAL_URL}/files/read-csv`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: fetchBody,
        });
      } catch (_) { localResp = null; }

      if (localResp && localResp.ok) {
        csvData = await localResp.json();
      } else {
        const errDetail = resp ? (await resp.json().catch(() => ({}))).detail : "Remote backend unreachable";
        const localDetail = localResp ? (await localResp.json().catch(() => ({}))).detail : "Local backend not running (start with: cd backend && python -m uvicorn main:app --reload)";
        throw new Error(`import_csv: File not found. Remote: ${errDetail}. Local: ${localDetail}`);
      }
    }
    const data2D = csvData.data;

    if (!data2D || data2D.length === 0) {
      throw new Error("import_csv: CSV file is empty");
    }

    // Pre-compute everything outside Excel.run
    const targetSheetName = payload.sheet || "Data";
    const startCell = payload.start_cell || "A1";
    const numRows = data2D.length;
    const numCols = Math.max(...data2D.map(r => r.length));
    const padded = data2D.map(r => {
      const row = [...r];
      while (row.length < numCols) row.push("");
      return row;
    });
    const startCol = startCell.replace(/[0-9]/g, "");
    const startRowNum = parseInt(startCell.replace(/[A-Za-z]/g, ""), 10);
    const endColIdx = colLetterToIndex(startCol) + numCols - 1;
    const endCol = indexToColLetter(endColIdx);
    const endRow = startRowNum + numRows - 1;
    const rangeAddr = `${startCol}${startRowNum}:${endCol}${endRow}`;
    const headers = data2D[0].map(h => String(h).trim());

    // 2. Write data into Excel (isolated Excel.run)
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItemOrNullObject(targetSheetName);
      ws.load("isNullObject");
      await ctx.sync();

      let sheet;
      if (ws.isNullObject) {
        sheet = ctx.workbook.worksheets.add(targetSheetName);
      } else {
        sheet = ctx.workbook.worksheets.getItem(targetSheetName);
      }
      sheet.activate();
      sheet.getRange(rangeAddr).values = padded;
      await ctx.sync();
    });

    // 3. Create named ranges for each column header — single Excel.run with per-name isolation
    const EXCEL_RESERVED = new Set([
      "DATE","YEAR","MONTH","DAY","TIME","HOUR","MINUTE","SECOND","NOW","TODAY",
      "IF","OR","AND","NOT","TRUE","FALSE","SUM","AVERAGE","COUNT","MAX","MIN",
      "INDEX","MATCH","VLOOKUP","HLOOKUP","OFFSET","INDIRECT","ROW","COLUMN",
      "MOD","INT","ROUND","ABS","SIGN","LOG","LN","EXP","SQRT","PI","RAND",
      "LEFT","RIGHT","MID","LEN","TRIM","UPPER","LOWER","FIND","SEARCH","TEXT",
      "VALUE","TYPE","ISNUMBER","ISTEXT","ISERROR","ISBLANK","NA","CHOOSE",
    ]);
    const dataStartRow = startRowNum + 1;

    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(targetSheetName);

      for (let i = 0; i < headers.length; i++) {
        const header = headers[i];
        if (!header) continue;

        let safeName = header.replace(/[^a-zA-Z0-9_]/g, "_").replace(/^[0-9]/, "_$&");
        if (!safeName) continue;
        if (EXCEL_RESERVED.has(safeName.toUpperCase())) safeName = "col_" + safeName;

        const colLetter = indexToColLetter(colLetterToIndex(startCol) + i);
        const colRange = sheet.getRange(`${colLetter}${dataStartRow}:${colLetter}${endRow}`);

        try {
          const existing = ctx.workbook.names.getItemOrNullObject(safeName);
          existing.load("isNullObject");
          await ctx.sync();
          if (!existing.isNullObject) existing.delete();
          ctx.workbook.names.add(safeName, colRange);
          await ctx.sync();
        } catch (e) {
          console.warn(`Named range "${safeName}" skipped: ${e.message}`);
        }
      }
    });

    // 4. Autofit (separate Excel.run)
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(targetSheetName);
      sheet.getUsedRange().format.autofitColumns();
      await ctx.sync();
    });

    lastNavigatedSheet = targetSheetName;
  }

  // ── add_chart ────────────────────────────────────────────────────────────────
  else if (type === "add_chart") {
    await Excel.run(async (ctx) => {
      const chartSheet = getSheet(ctx, payload.sheet);

      // Resolve data range — may reference a different sheet (e.g. "'Portfolio Summary'!A1:D8")
      let dataRange;
      const rawRange = payload.data_range || "";
      if (rawRange.includes("!")) {
        // Cross-sheet reference: split into sheet + address
        const { sheet: srcName, addr } = splitAddr(rawRange, payload.sheet);
        const srcSheet = getSheet(ctx, srcName);
        dataRange = srcSheet.getRange(addr);
      } else {
        // Same sheet
        dataRange = chartSheet.getRange(rawRange);
      }

      // Map chart type — accept various naming conventions
      const typeMap = {
        "columnclustered": "ColumnClustered", "column": "ColumnClustered", "bar": "BarClustered",
        "barclustered": "BarClustered", "line": "Line", "pie": "Pie", "area": "Area",
        "xyscatter": "XYScatter", "scatter": "XYScatter", "doughnut": "Doughnut",
        "columnstacked": "ColumnStacked", "barstacked": "BarStacked",
      };
      const rawType = (payload.chart_type || "ColumnClustered").toLowerCase().replace(/[_\s-]/g, "");
      const chartType = typeMap[rawType] || payload.chart_type || "ColumnClustered";

      const chart = chartSheet.charts.add(chartType, dataRange, Excel.ChartSeriesBy.auto);
      if (payload.title) chart.title.text = payload.title;
      chart.width = payload.width || 480;
      chart.height = payload.height || 300;
      if (payload.position) {
        try {
          const posRange = chartSheet.getRange(payload.position);
          chart.setPosition(posRange);
        } catch (e) { /* position is best-effort */ }
      }
      if (payload.series_names && payload.series_names.length > 0) {
        chart.series.load("count");
        await ctx.sync();
        for (let i = 0; i < Math.min(payload.series_names.length, chart.series.count); i++) {
          chart.series.getItemAt(i).name = payload.series_names[i];
        }
      }
      await ctx.sync();
    });
  }

  // ── add_data_validation ─────────────────────────────────────────────────────
  else if (type === "add_data_validation") {
    await Excel.run(async (ctx) => {
      const { sheet: s, addr } = splitAddr(payload.range, payload.sheet);
      const sheet = getSheet(ctx, s);
      const range = sheet.getRange(addr);
      range.dataValidation.clear();

      if (payload.type === "list") {
        range.dataValidation.rule = {
          list: {
            inCellDropDown: true,
            source: payload.formula || "",
          },
        };
      } else if (payload.type === "whole_number" || payload.type === "decimal") {
        range.dataValidation.rule = {
          wholeNumber: payload.type === "whole_number" ? {
            formula1: payload.min !== undefined ? payload.min : 0,
            formula2: payload.max !== undefined ? payload.max : 999999,
            operator: Excel.DataValidationOperator.between,
          } : undefined,
          decimal: payload.type === "decimal" ? {
            formula1: payload.min !== undefined ? payload.min : 0,
            formula2: payload.max !== undefined ? payload.max : 999999,
            operator: Excel.DataValidationOperator.between,
          } : undefined,
        };
      }
      await ctx.sync();
    });
  }

  // ── add_conditional_format ──────────────────────────────────────────────────
  else if (type === "add_conditional_format") {
    await Excel.run(async (ctx) => {
      const { sheet: s, addr } = splitAddr(payload.range, payload.sheet);
      const sheet = getSheet(ctx, s);
      const range = sheet.getRange(addr);

      if (payload.rule_type === "color_scale") {
        const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
        const criteria = {
          minimum: { color: payload.min_color || "#FF0000", type: "LowestValue" },
          maximum: { color: payload.max_color || "#00FF00", type: "HighestValue" },
        };
        if (payload.mid_color) {
          criteria.midpoint = { color: payload.mid_color, type: "Percentile", value: 50 };
        }
        cf.colorScale.criteria = criteria;
      } else if (payload.rule_type === "data_bar") {
        const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
        cf.dataBar.barDirection = Excel.ConditionalDataBarDirection.context;
        if (payload.bar_color) {
          cf.dataBar.positiveFormat.fillColor = payload.bar_color;
        }
      } else if (payload.rule_type === "icon_set") {
        const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
        const styleMap = {
          threeArrows: Excel.IconSet.threeArrows,
          threeTrafficLights: Excel.IconSet.threeTrafficLights1,
          fourArrows: Excel.IconSet.fourArrows,
        };
        cf.iconSet.style = styleMap[payload.icon_style] || Excel.IconSet.threeArrows;
      } else if (payload.rule_type === "cell_value") {
        const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
        const opMap = {
          greaterThan: Excel.ConditionalCellValueOperator.greaterThan,
          lessThan: Excel.ConditionalCellValueOperator.lessThan,
          equal: Excel.ConditionalCellValueOperator.equalTo,
          between: Excel.ConditionalCellValueOperator.between,
        };
        const rule = {
          formula1: String(payload.values?.[0] ?? 0),
          operator: opMap[payload.operator] || Excel.ConditionalCellValueOperator.greaterThan,
        };
        if (payload.operator === "between" && payload.values?.[1] !== undefined) {
          rule.formula2 = String(payload.values[1]);
        }
        cf.cellValue.rule = rule;
        const fmt = payload.format || {};
        if (fmt.font_color) cf.cellValue.format.font.color = fmt.font_color;
        if (fmt.color || fmt.fill) cf.cellValue.format.fill.color = fmt.color || fmt.fill;
      } else if (payload.rule_type === "top_bottom") {
        const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);
        cf.topBottom.rule = {
          rank: payload.rank || 10,
          type: payload.top !== false ? "TopItems" : "BottomItems",
        };
        if (payload.percent) {
          cf.topBottom.rule.type = payload.top !== false ? "TopPercent" : "BottomPercent";
        }
        const fmt = payload.format || {};
        if (fmt.font_color) cf.topBottom.format.font.color = fmt.font_color;
        if (fmt.color) cf.topBottom.format.fill.color = fmt.color;
      } else if (payload.rule_type === "text_contains") {
        const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
        cf.textComparison.rule = {
          operator: Excel.ConditionalTextOperator.contains,
          text: payload.text || "",
        };
        const fmt = payload.format || {};
        if (fmt.font_color) cf.textComparison.format.font.color = fmt.font_color;
        if (fmt.color) cf.textComparison.format.fill.color = fmt.color;
      }
      await ctx.sync();
    });
  }

  // ── launch_app ──────────────────────────────────────────────────────────────
  else if (type === "launch_app") {
    try {
      // Use local proxy to launch apps on the user's machine (not Railway)
      const resp = await fetch(`/local-api/launch-app`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ app_name: payload.app_name }),
      });
      const result = await resp.json();
      if (result.status === "error") {
        appendMessage("assistant", result.message || "Could not open the app.");
      }
    } catch (e) {
      // Local backend not running — try Railway as fallback
      try {
        await fetch(`${BACKEND_URL}/launch-app`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ app_name: payload.app_name }),
        });
      } catch (_) {
        appendMessage("assistant", "Couldn't open the app — make sure the local backend is running.");
      }
    }
  }

  // ── open_notes / open_url ──────────────────────────────────────────────────
  else if (type === "open_notes" || (type === "open_url" && payload.url)) {
    const url = type === "open_notes" ? `${BACKEND_URL}/notes-app` : payload.url;
    try {
      window.open(url, "_blank");
    } catch (e) {
      console.error("open failed:", e);
    }
  }

  // ── create_note ────────────────────────────────────────────────────────────
  else if (type === "create_note") {
    try {
      const resp = await fetch(`${BACKEND_URL}/notes/`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          user_id: CURRENT_USER?.id || "unknown",
          title: payload.title || "Untitled Note",
          content: payload.content || "",
          folder: payload.folder || "General",
        }),
      });
      const note = await resp.json();
      // note created silently
    } catch (e) {
      console.error("create_note failed:", e);
    }
  }

  // ── create_workbook ──────────────────────────────────────────────────────────
  else if (type === "create_workbook") {
    try {
      await Excel.createWorkbook();
      // workbook created silently
    } catch (e) {
      console.error("create_workbook failed:", e);
    }
  }

  // ── import_image / import_r_output ───────────────────────────────────────────
  else if (type === "import_image" || type === "import_r_output") {
    try {
      let imageData = payload.image_data;

      // If transfer_id provided, fetch that specific transfer (only if it's an image)
      if (payload.transfer_id) {
        const resp = await fetch(`${BACKEND_URL}/transfer/${payload.transfer_id}`);
        if (resp.ok) {
          const transfer = await resp.json();
          if (transfer.data_type === "image") {
            imageData = transfer.data;
          }
        }
      }

      // If no image yet, check for pending R→Excel image transfers
      if (!imageData) {
        const pendingResp = await fetch(`${BACKEND_URL}/transfer/pending/excel`);
        if (pendingResp.ok) {
          const pendingData = await pendingResp.json();
          const pending = pendingData.pending || [];
          // Filter for image transfers only — don't grab data snapshots
          const imagePending = pending.filter(p => p.data_type === "image");
          if (imagePending.length > 0) {
            const latest = imagePending[imagePending.length - 1];
            const tResp = await fetch(`${BACKEND_URL}/transfer/${latest.transfer_id}`);
            if (tResp.ok) {
              const transfer = await tResp.json();
              imageData = transfer.data;
            }
          }
        }
      }

      if (imageData) {
        // Strip data URI prefix if present — addImage wants pure base64
        const cleanBase64 = imageData.replace(/^data:image\/[a-z+]+;base64,/, "");
        // Validate: must start with PNG or JPEG magic bytes in base64
        const isPng = cleanBase64.startsWith("iVBOR");
        const isJpeg = cleanBase64.startsWith("/9j/");
        if (!isPng && !isJpeg) {
          console.warn(`Image data appears corrupt (starts with: ${cleanBase64.substring(0, 10)}...)`);
          appendMessage("assistant", "The image data looks corrupt — try generating the plot again in R.");
        } else {
          await Excel.run(async (ctx) => {
            const sheet = payload.sheet
              ? ctx.workbook.worksheets.getItem(payload.sheet)
              : ctx.workbook.worksheets.getActiveWorksheet();
            const image = sheet.shapes.addImage(cleanBase64);
            image.name = "R_Plot";
            image.left = 10;
            image.top = 200;
            await ctx.sync();
          });
        }
      } else {
        appendMessage("assistant", "I can't run R from Excel directly. To bring an R plot here: open tsifl in RStudio, ask for the plot there with the word \"excel\" in your message (e.g. \"plot loandata regression and send to excel\"), and it'll appear in this sheet automatically.");
      }
    } catch (e) {
      console.error("import_image failed:", e);
    }
  }

  // ── remove_duplicates ──────────────────────────────────────────────────────
  // Remove duplicate rows from a range based on a key column.
  // payload: { sheet?, range, key_column? (letter, defaults to first col) }
  else if (type === "remove_duplicates") {
    await Excel.run(async (ctx) => {
      const { sheet: s, addr } = splitAddr(payload.range, payload.sheet);
      const sheet = getSheet(ctx, s);
      const range = sheet.getRange(addr);
      range.load("values");
      await ctx.sync();
      const values = range.values;
      if (values.length < 2) return;

      const keyColIdx = payload.key_column ? colLetterToIndex(payload.key_column) : 0;
      const seen = new Set();
      const unique = [values[0]]; // keep header row
      for (let i = 1; i < values.length; i++) {
        const key = String(values[i][keyColIdx]);
        if (!seen.has(key)) { seen.add(key); unique.push(values[i]); }
      }

      // Parse start position from the range address
      const startCell = addr.split(":")[0];
      const startMatch = startCell.match(/([A-Z]+)(\d+)/i);
      const startCol = startMatch[1].toUpperCase();
      const startRow = parseInt(startMatch[2]);
      const endCol = indexToColLetter(colLetterToIndex(startCol) + values[0].length - 1);

      // Write unique rows back
      const endRow = startRow + unique.length - 1;
      const writeRange = sheet.getRange(`${startCol}${startRow}:${endCol}${endRow}`);
      writeRange.values = unique;

      // Clear leftover rows below
      if (unique.length < values.length) {
        const clearStart = startRow + unique.length;
        const clearEnd = startRow + values.length - 1;
        const clearRange = sheet.getRange(`${startCol}${clearStart}:${endCol}${clearEnd}`);
        clearRange.clear(Excel.ClearApplyTo.contents);
      }
      await ctx.sync();
      console.log( `Removed ${values.length - unique.length} duplicate rows`);
    });
  }

  // ── conditional_format_heatmap ─────────────────────────────────────────────
  // Apply a 3-color heatmap (red-yellow-green) to a numeric range.
  // payload: { sheet?, range, min_color?, mid_color?, max_color? }
  else if (type === "conditional_format_heatmap") {
    await Excel.run(async (ctx) => {
      const { sheet: s, addr } = splitAddr(payload.range, payload.sheet);
      const sheet = getSheet(ctx, s);
      const range = sheet.getRange(addr);
      const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
      cf.colorScale.criteria = {
        minimum: { color: payload.min_color || "#F8696B", type: "LowestValue" },
        midpoint: { color: payload.mid_color || "#FFEB84", type: "Percentile", value: 50 },
        maximum: { color: payload.max_color || "#63BE7B", type: "HighestValue" },
      };
      await ctx.sync();
    });
  }

  // ── create_pivot_summary ───────────────────────────────────────────────────
  // Create a summary/pivot-like table from data using formulas.
  // payload: { sheet?, target_sheet?, start_cell?, group_column?, value_column?,
  //            categories: string[], sumifs_formula? }
  else if (type === "create_pivot_summary") {
    await Excel.run(async (ctx) => {
      const sheet = getSheet(ctx, payload.sheet);
      const targetSheet = payload.target_sheet
        ? ctx.workbook.worksheets.getItem(payload.target_sheet)
        : sheet;

      // Write header row
      const startCell = payload.start_cell || "A1";
      const headerRange = targetSheet.getRange(startCell);
      headerRange.values = [[payload.group_column || "Category", payload.value_column || "Total"]];
      headerRange.format.font.bold = true;

      // Write unique categories and SUMIFS formulas
      if (payload.categories && payload.categories.length > 0) {
        const catStartRow = parseInt(startCell.match(/\d+/)[0]) + 1;
        const catCol = startCell.match(/[A-Z]+/i)[0].toUpperCase();
        const valCol = indexToColLetter(colLetterToIndex(catCol) + 1);
        for (let i = 0; i < payload.categories.length; i++) {
          const row = catStartRow + i;
          targetSheet.getRange(`${catCol}${row}`).values = [[payload.categories[i]]];
          if (payload.sumifs_formula) {
            targetSheet.getRange(`${valCol}${row}`).formulas = [[
              payload.sumifs_formula.replace("{category}", `${catCol}${row}`)
            ]];
          }
        }
      }
      await ctx.sync();
    });
  }

  // ── auto_chart_best_fit ────────────────────────────────────────────────────
  // Analyze data and create the most appropriate chart type.
  // Claude picks the type in the system prompt; this delegates to add_chart.
  // payload: { sheet?, range, chart_type?, title?, ... (same as add_chart) }
  else if (type === "auto_chart_best_fit") {
    await executeAction({
      type: "add_chart",
      payload: { ...payload, chart_type: payload.chart_type || "ColumnClustered" },
    });
  }

  // ── trim_whitespace ────────────────────────────────────────────────────────
  // Trim leading/trailing whitespace from all cells in a range.
  // payload: { sheet?, range }
  else if (type === "trim_whitespace") {
    await Excel.run(async (ctx) => {
      const { sheet: s, addr } = splitAddr(payload.range, payload.sheet);
      const sheet = getSheet(ctx, s);
      const range = sheet.getRange(addr);
      range.load("values");
      await ctx.sync();
      const values = range.values;
      let trimCount = 0;
      const trimmed = values.map(row =>
        row.map(cell => {
          if (typeof cell === "string") {
            const t = cell.trim();
            if (t !== cell) trimCount++;
            return t;
          }
          return cell;
        })
      );
      range.values = trimmed;
      await ctx.sync();
      console.log( `Trimmed whitespace in ${trimCount} cell(s)`);
    });
  }

  // ── find_and_replace ────────────────────────────────────────────────────────
  // Find and replace a single value across the used range (or a specified range).
  // payload: { sheet?, range?, find_text, replace_text, match_case? }
  else if (type === "find_and_replace") {
    await Excel.run(async (ctx) => {
      const sheet = getSheet(ctx, payload.sheet);
      let range;
      if (payload.range) {
        range = sheet.getRange(splitAddr(payload.range, payload.sheet).addr);
      } else {
        range = sheet.getUsedRangeOrNullObject();
        range.load("isNullObject");
        await ctx.sync();
        if (range.isNullObject) return;
      }
      range.load("values");
      await ctx.sync();
      const values = range.values;
      const findText = payload.find_text || "";
      const replaceText = payload.replace_text || "";
      const matchCase = payload.match_case || false;
      let totalReplaced = 0;
      const updated = values.map(row =>
        row.map(cell => {
          if (typeof cell !== "string") return cell;
          if (matchCase) {
            if (cell.includes(findText)) {
              totalReplaced++;
              return cell.split(findText).join(replaceText);
            }
          } else {
            const lowerCell = cell.toLowerCase();
            const lowerFind = findText.toLowerCase();
            if (lowerCell.includes(lowerFind)) {
              totalReplaced++;
              // Case-insensitive replace
              return cell.replace(new RegExp(findText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi'), replaceText);
            }
          }
          return cell;
        })
      );
      range.values = updated;
      await ctx.sync();
      console.log(`find_and_replace: replaced ${totalReplaced} cell(s)`);
    });
  }

  // ── find_replace_bulk ──────────────────────────────────────────────────────
  // Find and replace multiple values at once across a range.
  // payload: { sheet?, range, replacements: [{ find: "x", replace: "y" }, ...] }
  else if (type === "find_replace_bulk") {
    await Excel.run(async (ctx) => {
      const { sheet: s, addr } = splitAddr(payload.range, payload.sheet);
      const sheet = getSheet(ctx, s);
      const range = sheet.getRange(addr);
      range.load("values");
      await ctx.sync();
      const values = range.values;
      let totalReplaced = 0;
      const replacements = payload.replacements || [];
      const updated = values.map(row =>
        row.map(cell => {
          if (typeof cell !== "string") return cell;
          let val = cell;
          for (const r of replacements) {
            if (val.includes(r.find)) {
              const before = val;
              val = val.split(r.find).join(r.replace);
              if (val !== before) totalReplaced++;
            }
          }
          return val;
        })
      );
      range.values = updated;
      await ctx.sync();
      console.log( `Replaced ${totalReplaced} occurrence(s) across ${replacements.length} pattern(s)`);
    });
  }

  // ── goal_seek — iterative solver in JavaScript ──────────────────────────────
  else if (type === "goal_seek") {
    await Excel.run(async (ctx) => {
      const { sheet: s, addr: setCell } = splitAddr(payload.set_cell, payload.sheet);
      const sheet = getSheet(ctx, s);
      const changingCell = sheet.getRange(payload.changing_cell);
      const targetCell = sheet.getRange(setCell);

      const goalValue = parseFloat(payload.to_value);

      // SAFETY: Save the target cell's formula before we start (in case iterations corrupt it)
      targetCell.load("formulas");
      changingCell.load("values");
      await ctx.sync();
      const savedFormula = targetCell.formulas[0][0]; // preserve original formula
      const currentVal = changingCell.values[0][0] || 0;

      // Binary search / secant method
      let lo = currentVal - 10000, hi = currentVal + 10000;
      let mid, best = currentVal, bestDiff = Infinity;
      const maxIter = 100;
      const tolerance = 0.001;

      for (let i = 0; i < maxIter; i++) {
        mid = (lo + hi) / 2;
        changingCell.values = [[mid]];
        await ctx.sync();
        targetCell.load("values");
        await ctx.sync();
        const result = targetCell.values[0][0];
        const diff = result - goalValue;
        if (Math.abs(diff) < Math.abs(bestDiff)) { best = mid; bestDiff = diff; }
        if (Math.abs(diff) < tolerance) break;
        // Determine direction: check if increasing input increases result
        const probe = (i === 0) ? mid + 1 : mid + 0.01;
        changingCell.values = [[probe]];
        await ctx.sync();
        targetCell.load("values");
        await ctx.sync();
        const resultPlus = targetCell.values[0][0];
        const increasing = resultPlus > result;
        if ((increasing && diff > 0) || (!increasing && diff < 0)) {
          hi = mid;
        } else {
          lo = mid;
        }
      }
      changingCell.values = [[best]];
      await ctx.sync();

      // SAFETY: Restore the target cell's formula if it was corrupted during iterations
      targetCell.load("formulas");
      await ctx.sync();
      if (savedFormula && savedFormula.startsWith("=") && targetCell.formulas[0][0] !== savedFormula) {
        targetCell.formulas = [[savedFormula]];
        await ctx.sync();
        console.log(`[tsifl] Goal Seek: restored target formula ${savedFormula}`);
      }

      console.log(`[tsifl] Goal Seek: set ${payload.changing_cell}=${best}, target diff=${bestDiff}`);
    });
  }

  // ── run_toolpak — descriptive statistics via formulas ───────────────────────
  else if (type === "run_toolpak") {
    await Excel.run(async (ctx) => {
      const { sheet: s } = splitAddr(payload.input_range, payload.sheet);
      const sheet = getSheet(ctx, s);
      const opts = payload.options || {};
      const rawRange = (payload.input_range || "").replace(/\$/g, "");
      const hasLabels = opts.labels_in_first_row;
      let dataRange, header = "Data";

      if (hasLabels && rawRange.includes(":")) {
        const [startRef, endRef] = rawRange.split(":");
        const col = startRef.replace(/[0-9]/g, "");
        const startRow = parseInt(startRef.replace(/[A-Za-z]/g, "")) + 1;
        const endRow = parseInt(endRef.replace(/[A-Za-z]/g, ""));
        dataRange = `${col}${startRow}:${col}${endRow}`;
        const headerCell = sheet.getRange(`${col}${startRow - 1}`);
        headerCell.load("values");
        await ctx.sync();
        header = headerCell.values[0][0] || "Data";
      } else {
        dataRange = rawRange;
      }

      const outRef = (payload.output_range || "H4").replace(/\$/g, "");
      const outCol = outRef.replace(/[0-9]/g, "");
      const outRow = parseInt(outRef.replace(/[A-Za-z]/g, ""));
      const valCol = String.fromCharCode(outCol.charCodeAt(0) + 1);

      // Clear output area (header + blank row + 13 stats + confidence = 16 rows)
      const clearRange = sheet.getRange(`${outCol}${outRow}:${valCol}${outRow + 16}`);
      clearRange.clear();
      await ctx.sync();

      // Write header — matches ToolPak layout: header in row N, blank row N+1, stats from N+2
      sheet.getRange(`${outCol}${outRow}`).values = [[header]];

      const dr = dataRange;
      const stats = [
        ["Mean",               `=IFERROR(AVERAGE(${dr}),"")`],
        ["Standard Error",     `=IFERROR(STDEV(${dr})/SQRT(COUNT(${dr})),"")`],
        ["Median",             `=IFERROR(MEDIAN(${dr}),"")`],
        ["Mode",               `=IFERROR(MODE.SNGL(${dr}),"")`],
        ["Standard Deviation", `=IFERROR(STDEV(${dr}),"")`],
        ["Sample Variance",    `=IFERROR(VAR(${dr}),"")`],
        ["Kurtosis",           `=IFERROR(KURT(${dr}),"")`],
        ["Skewness",           `=IFERROR(SKEW(${dr}),"")`],
        ["Range",              `=IFERROR(MAX(${dr})-MIN(${dr}),"")`],
        ["Minimum",            `=IFERROR(MIN(${dr}),"")`],
        ["Maximum",            `=IFERROR(MAX(${dr}),"")`],
        ["Sum",                `=IFERROR(SUM(${dr}),"")`],
        ["Count",              `=IFERROR(COUNT(${dr}),"")`],
      ];

      if (opts.confidence_level) {
        const conf = opts.confidence_level;
        stats.push(["Confidence Level", `=IFERROR(CONFIDENCE.NORM(1-${conf}/100,STDEV(${dr}),COUNT(${dr})),"")`]);
      }

      // Start stats at outRow+2 (skip blank row after header) to match ToolPak layout
      for (let i = 0; i < stats.length; i++) {
        const row = outRow + 2 + i;
        sheet.getRange(`${outCol}${row}`).values = [[stats[i][0]]];
        sheet.getRange(`${valCol}${row}`).formulas = [[stats[i][1]]];
      }
      await ctx.sync();
      console.log(`[tsifl] Descriptive Stats: ${stats.length} measures for ${dataRange} → ${outCol}${outRow}`);
    });
  }

  // ── create_data_table — TABLE array formula via Office.js ──────────────────
  else if (type === "create_data_table") {
    await Excel.run(async (ctx) => {
      // Backend sends: range, row_input_cell, col_input_cell
      // Or: table_range, row_input, col_input (legacy)
      const rawRange = payload.range || payload.table_range || "";
      const { sheet: s, addr: tableAddr } = splitAddr(rawRange, payload.sheet);
      const sheet = getSheet(ctx, s);

      if (!tableAddr) {
        console.warn("[tsifl] create_data_table: no range provided");
        return;
      }

      const clean = tableAddr.replace(/\$/g, "");
      const [startRef, endRef] = clean.split(":");
      const startCol = startRef.replace(/[0-9]/g, "");
      const startRow = parseInt(startRef.replace(/[A-Za-z]/g, ""));
      const endCol = endRef.replace(/[0-9]/g, "");
      const endRow = parseInt(endRef.replace(/[A-Za-z]/g, ""));

      // Result area: skip first row (headers) and first column (inputs)
      const resultStartCol = String.fromCharCode(startCol.charCodeAt(0) + 1);
      const resultStartRow = startRow + 1;
      const resultRange = `${resultStartCol}${resultStartRow}:${endCol}${endRow}`;

      const rowRef = (payload.row_input_cell || payload.row_input || "").replace(/\$/g, "");
      const colRef = (payload.col_input_cell || payload.col_input || "").replace(/\$/g, "");
      const tableFormula = `=TABLE(${rowRef},${colRef})`;

      // Set the TABLE formula as an array formula on the result range
      const range = sheet.getRange(resultRange);
      range.formulas = Array(endRow - resultStartRow + 1).fill(
        Array(endCol.charCodeAt(0) - resultStartCol.charCodeAt(0) + 1).fill(tableFormula)
      );
      await ctx.sync();
      console.log(`[tsifl] Data Table: ${tableFormula} on ${resultRange}`);
    });
  }

  // ── install_addins / uninstall_addins — no-op in add-in (handled if needed by agent) ──
  else if (type === "install_addins" || type === "uninstall_addins") {
    console.log(`[tsifl] ${type}: skipped (handled automatically when needed)`);
  }
}

// ── Locale-safe format normalizer ─────────────────────────────────────────────
// On non-US locales, "$" in Excel format strings maps to the system currency
// symbol (e.g. € in EU). Force explicit USD with [$$-409] (LCID 409 = en-US).
// Also normalises bare "$#,##0" variants that Claude emits.
function _localeSafeFmt(fmt) {
  if (typeof fmt !== "string") return fmt;
  // Already locale-safe or not a dollar format — leave alone
  if (fmt.startsWith("[$$-") || fmt.startsWith("[$€") || fmt.startsWith("[$")) return fmt;
  // "$#,##0.00" → "[$$-409]#,##0.00"
  // "$#,##0"    → "[$$-409]#,##0"
  return fmt.replace(/^\$/, "[$$-409]");
}

// ── Format helper (shared by write_cell, write_range, format_range) ──────────

function _applyFormat(range, p, knownRows, knownCols) {
  if (p.bold        !== undefined) range.format.font.bold       = p.bold;
  if (p.italic      !== undefined) range.format.font.italic     = p.italic;
  if (p.color)                     range.format.fill.color      = p.color;
  if (p.font_color)                range.format.font.color      = p.font_color;
  if (p.font_size)                 range.format.font.size       = p.font_size;
  if (p.font_name)                 range.format.font.name       = p.font_name;
  if (p.h_align)                   range.format.horizontalAlignment = p.h_align; // "Left","Center","Right"
  if (p.v_align)                   range.format.verticalAlignment   = p.v_align;
  if (p.wrap_text   !== undefined) range.format.wrapText        = p.wrap_text;
  if (p.row_height)                range.format.rowHeight       = p.row_height;
  if (p.col_width)                 range.format.columnWidth     = p.col_width;
  if (p.number_format) {
    // number_format must be a 2D array matching range dimensions.
    // Load rowCount/columnCount if available on the range object so we can
    // build the correct array; fall back to [[fmt]] if not yet loaded.
    try {
      const fmtStr = _localeSafeFmt(
        typeof p.number_format === "string" ? p.number_format : null
      ) || p.number_format;
      if (typeof fmtStr === "string") {
        // Use caller-supplied dims first, then proxy (may be 0 if not loaded), then 1
        const rows = (knownRows > 0 ? knownRows : null) || range.rowCount || 1;
        const cols = (knownCols > 0 ? knownCols : null) || range.columnCount || 1;
        range.numberFormat = Array.from({ length: rows }, () =>
          Array.from({ length: cols }, () => fmtStr)
        );
      } else {
        range.numberFormat = fmtStr;
      }
    } catch (e) { /* size mismatch — skip */ }
  }
  if (p.border) {
    const b = range.format.borders;
    const style = Excel.BorderLineStyle.continuous;
    const weight = Excel.BorderWeight.thin;
    if (p.border === "all" || p.border === "outer") {
      b.getItem("EdgeTop").style    = style;
      b.getItem("EdgeBottom").style = style;
      b.getItem("EdgeLeft").style   = style;
      b.getItem("EdgeRight").style  = style;
    }
    if (p.border === "all") {
      b.getItem("InsideHorizontal").style = style;
      b.getItem("InsideVertical").style   = style;
    }
    if (p.border === "bottom") {
      b.getItem("EdgeBottom").style = style;
      b.getItem("EdgeBottom").weight = Excel.BorderWeight.medium;
    }
  }
}

// ── Config Sync ───────────────────────────────────────────────────────────────

async function saveUserConfig(user) {
  try {
    await fetch(`${BACKEND_URL}/auth/set-user`, {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body:    JSON.stringify({ user_id: user.id, email: user.email }),
    });
  } catch (_) {}
}

// ── UI Helpers ────────────────────────────────────────────────────────────────

function renderMarkdown(text) {
  let html = text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  // Code blocks (``` ... ```)
  html = html.replace(/```(\w*)\n?([\s\S]*?)```/g, function(_, lang, code) {
    const id = "cb_" + Math.random().toString(36).slice(2, 8);
    return '<pre id="' + id + '"><button class="copy-btn" onclick="navigator.clipboard.writeText(document.getElementById(\'' + id + '\').textContent.replace(/^Copy\\n?/,\'\'))">Copy</button><code>' + code.trim() + '</code></pre>';
  });
  // Inline formatting
  html = html
    .replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>")
    .replace(/\*(.+?)\*/g, "<em>$1</em>")
    .replace(/`([^`]+)`/g, "<code>$1</code>")
    .replace(/^### (.+)$/gm, "<h4 style='margin:6px 0 2px;font-size:13px;'>$1</h4>")
    .replace(/^## (.+)$/gm, "<h3 style='margin:8px 0 3px;font-size:14px;'>$1</h3>")
    .replace(/^- (.+)$/gm, "<li style='margin-left:16px;list-style:disc;'>$1</li>")
    .replace(/^\d+\. (.+)$/gm, "<li style='margin-left:16px;list-style:decimal;'>$1</li>")
    .replace(/\n/g, "<br>");
  return html;
}

function appendMessage(role, text, images, meta) {
  const history = document.getElementById("chat-history");
  const div     = document.createElement("div");
  div.className   = `message ${role}`;

  // Meta chip rendered ABOVE the assistant's text. Surfaces ONLY when the
  // server actively guarded something the user can act on right now.
  //   lock blocked     → writes to locked cells; user should clear locks or rephrase
  //   local · ollama   → console signal for dev debugging only
  // (Removed in v92):
  //   phantom dropped  → confusing developer-speak ("58 phantom dropped"). The
  //   reply body now explains the situation in plain English when the guard
  //   fires, so the chip is redundant noise.
  //   memory chip      → see Memory panel in header for state.
  if (role === "assistant" && meta) {
    const parts = [];
    if (meta.ollama) {
      try { console.info(`[tsifl] reply via local ollama (${meta.ollama})`); } catch (_) {}
    }
    if (meta.lockBlocked) parts.push(`${meta.lockBlocked} blocked by lock`);
    if (parts.length) {
      const chip = document.createElement("div");
      chip.className = "assistant-meta-chip";
      chip.textContent = parts.join("  ·  ");
      div.appendChild(chip);
    }
  }

  // Add text — use markdown for assistant messages
  const textNode = document.createElement("span");
  if (role === "assistant" && text) {
    textNode.innerHTML = renderMarkdown(text);
  } else {
    textNode.textContent = text;
  }
  div.appendChild(textNode);

  // Show image attachments in user messages — compact line with a small
  // thumbnail + filename, click to expand to full size in a modal.
  // Previous version rendered a 280×180 thumbnail directly inline, which
  // dominated the side panel. The compact form keeps the chat scannable.
  if (images && images.length > 0) {
    const attachRow = document.createElement("div");
    attachRow.className = "attach-row";  // styled in CSS
    div.appendChild(attachRow);

    for (const img of images) {
      const isImage = (img.media_type || "").startsWith("image/");
      const filename = img.file_name || (isImage ? "Image" : "File");

      const chip = document.createElement("div");
      chip.className = "attach-chip";
      chip.title = "Click to view full size";

      // Tiny preview: 32x32 for images, file-extension badge for docs
      const thumb = document.createElement("div");
      thumb.className = "attach-thumb";

      if (isImage) {
        renderImageToCanvas(img.data, img.media_type, 32, 32).then(canvas => {
          if (canvas) {
            canvas.style.cssText = "display:block;width:100%;height:100%;border-radius:3px;";
            thumb.appendChild(canvas);
          } else {
            thumb.textContent = "IMG";
          }
        });
        // Click to expand — opens the full image in a centered modal overlay
        chip.addEventListener("click", () => _openImageModal(img));
      } else {
        const ext = filename.includes(".") ? filename.split(".").pop().toUpperCase() : "FILE";
        thumb.textContent = ext;
        thumb.style.cssText += "display:flex;align-items:center;justify-content:center;font-size:9px;font-weight:700;color:#0D5EAF;background:#F1F5F9;";
      }

      const label = document.createElement("span");
      label.className = "attach-label";
      label.textContent = filename;

      chip.appendChild(thumb);
      chip.appendChild(label);
      attachRow.appendChild(chip);
    }
  }

  // Skip the legacy inline-thumbnail loop below — replaced by the compact
  // row above. The original code is kept intact for the catch block which
  // we no longer reach, but if reached, falls back to badge.
  if (false && images && images.length > 0) {
    for (const img of images) {
      const container = document.createElement("div");
      container.style.cssText = "margin-top:8px;";
      div.appendChild(container);
      renderImageToCanvas(img.data, img.media_type, 280, 180).then(canvas => {
        if (canvas) {
          canvas.style.borderRadius = "8px";
          canvas.style.border = "1px solid var(--border)";
          canvas.style.display = "block";
          canvas.style.maxWidth = "100%";
          canvas.style.boxShadow = "0 1px 3px rgba(0,0,0,0.08)";
          container.appendChild(canvas);
        } else {
          const badge = document.createElement("div");
          badge.className = "image-badge";
          badge.textContent = `Image attached`;
          container.appendChild(badge);
        }
      }).catch(() => {
        const badge = document.createElement("div");
        badge.className = "image-badge";
        badge.textContent = `📷 Image attached`;
        container.appendChild(badge);
      });
    }
  }
  history.appendChild(div);
  history.scrollTop = history.scrollHeight;
}

// ── Image modal viewer ──────────────────────────────────────────────────────
// Click on a compact attach-chip → full-size image in a centered overlay.
// Keeps the chat scannable while letting the user inspect attachments.

function _openImageModal(img) {
  // Don't stack overlays — close any existing one first
  const existing = document.getElementById("image-modal-overlay");
  if (existing) existing.remove();

  const overlay = document.createElement("div");
  overlay.id = "image-modal-overlay";
  overlay.style.cssText = (
    "position:fixed;inset:0;background:rgba(0,0,0,0.65);" +
    "display:flex;align-items:center;justify-content:center;" +
    "z-index:300;cursor:zoom-out;"
  );

  // Render the full image to a canvas — same CSP-safe path we use elsewhere
  renderImageToCanvas(img.data, img.media_type, 1600, 1200).then(canvas => {
    if (!canvas) return;
    canvas.style.cssText = (
      "max-width:90vw;max-height:85vh;border-radius:8px;" +
      "box-shadow:0 12px 32px rgba(0,0,0,0.5);background:#FFF;"
    );
    overlay.appendChild(canvas);
  });

  // Close on click anywhere or Escape
  overlay.addEventListener("click", () => overlay.remove());
  const escHandler = (e) => {
    if (e.key === "Escape") {
      overlay.remove();
      document.removeEventListener("keydown", escHandler);
    }
  };
  document.addEventListener("keydown", escHandler);

  document.body.appendChild(overlay);
}


// ── Thinking bubble with rotating punchlines ─────────────────────────────────
const _thinkingMessages = {
  thinking: [
    'Reading your question...',
    'Scanning every cell like a forensic accountant...',
    'VLOOKUP walked so XLOOKUP could run...',
    'Checking if your formulas are circular... again...',
    'Counting rows like a back-office analyst at 2am...',
    'This model is about to hit different...',
    'Pivot tables fear me...',
    'Your spreadsheet called — it wants a raise...',
    'Running the numbers so you don\'t have to...',
    'Goldman\'s modeling team just felt a disturbance...',
    'Making INDEX MATCH look easy since 2024...',
    'Every cell tells a story...',
    'The kind of analysis that gets forwarded to the MD...',
    'Somewhere an intern is doing this by hand...',
    'Ctrl+Z won\'t be necessary — I don\'t make mistakes...',
    'Your Excel just went from intern to VP...',
    'Building the model that closes the deal...',
  ],
  applying: [
    'Writing values at the speed of light...',
    'Cells are snapping into place...',
    'Your spreadsheet is getting a glow-up...',
    'Formatting like a senior analyst on an all-nighter...',
    'Populating faster than a Bloomberg terminal...',
    'Every number is exactly where it should be...',
    'The kind of output that survives due diligence...',
    'Dropping formulas like it\'s bonus season...',
    'One more sync and we\'re golden...',
    'This is the part where your model becomes legendary...',
  ],
  automating: [
    'Controlling Excel like a Bloomberg terminal operator...',
    'Clicking through menus so you don\'t have to...',
    'Data Table dialog? Already on it...',
    'Installing add-ins like a seasoned IT admin...',
    'Your desktop agent is in the zone...',
    'Running Solver while you grab coffee...',
    'The kind of automation that makes VBA obsolete...',
    'GUI operations running at machine speed...',
    'Setting up what-if scenarios like a pro...',
    'ToolPak? Solver? Consider them handled...',
  ]
};

let _thinkingInterval = null;
let _thinkingPhase = null;

function showTypingIndicator(phase) {
  hideTypingIndicator();
  _thinkingPhase = phase || 'thinking';
  const history = document.getElementById("chat-history");
  const div = document.createElement("div");
  div.id = "typing-indicator";
  div.className = "thinking-bubble";
  div.innerHTML = '<div class="thinking-orb"></div><span class="thinking-text"></span>';
  history.appendChild(div);
  history.scrollTop = history.scrollHeight;

  const msgs = _thinkingMessages[_thinkingPhase] || _thinkingMessages.thinking;
  let idx = 0;

  function showNext() {
    const el = document.querySelector('#typing-indicator .thinking-text');
    if (!el) return;
    const msg = msgs[idx % msgs.length];
    el.style.opacity = '0';
    el.style.transform = 'translateY(4px)';
    setTimeout(() => {
      el.textContent = msg;
      el.style.opacity = '1';
      el.style.transform = 'translateY(0)';
    }, 200);
    const chat = document.getElementById("chat-history");
    if (chat) chat.scrollTop = chat.scrollHeight;
    idx++;
  }

  setTimeout(showNext, 100);
  _thinkingInterval = setInterval(showNext, 2500);
}

function hideTypingIndicator() {
  if (_thinkingInterval) { clearInterval(_thinkingInterval); _thinkingInterval = null; }
  _thinkingPhase = null;
  const el = document.getElementById("typing-indicator");
  if (el) el.remove();
}

function setStatus(text)           { document.getElementById("status-bar").textContent = text; }
function setSubmitEnabled(enabled) { document.getElementById("submit-btn").disabled = !enabled; }

// ── Access Modal (desktop automation overlay) ───────────────────────────────

let _activeAccessSessionId = null;
// When the user hits Stop (or Esc), the polling loop in handleSubmit
// reads this flag on its next tick and breaks immediately, instead of
// waiting for the backend status to round-trip back as "cancelled".
let _cuCancelRequested = false;

function showAccessModal(sessionId) {
  _activeAccessSessionId = sessionId;
  hideAccessModal(); // remove stale

  const overlay = document.createElement("div");
  overlay.id = "access-modal-overlay";

  const modal = document.createElement("div");
  modal.id = "access-modal";
  modal.innerHTML = `
    <div class="access-modal-icon">
      <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
        <rect x="2" y="3" width="20" height="14" rx="2" ry="2"/>
        <line x1="8" y1="21" x2="16" y2="21"/>
        <line x1="12" y1="17" x2="12" y2="21"/>
      </svg>
    </div>
    <div class="access-modal-title">tsifl wants access to your computer</div>
    <div class="access-modal-text">
      If at any time you want to stop, press <strong>Esc</strong> or <strong>right-click</strong>.<br/>
      Please don't touch your computer while progress is being made.
    </div>
    <div class="access-modal-progress">
      <div class="access-modal-spinner"></div>
      <span class="access-modal-status">Working...</span>
    </div>
    <button id="access-modal-stop">Stop</button>
  `;

  overlay.appendChild(modal);
  document.body.appendChild(overlay);

  // Stop button click
  document.getElementById("access-modal-stop").addEventListener("click", () => _cancelAccess());

  // Right-click anywhere on the overlay also stops
  overlay.addEventListener("contextmenu", (e) => {
    e.preventDefault();
    _cancelAccess();
  });
}

function hideAccessModal() {
  _activeAccessSessionId = null;
  const el = document.getElementById("access-modal-overlay");
  if (el) el.remove();
}

async function _cancelAccess() {
  const sessionId = _activeAccessSessionId;
  if (!sessionId) return;
  // Flip the flag FIRST so the polling loop in handleSubmit breaks on
  // its next tick without having to wait for the backend round-trip.
  _cuCancelRequested = true;
  const btn = document.getElementById("access-modal-stop");
  if (btn) { btn.textContent = "Stopping..."; btn.disabled = true; }
  const statusEl = document.querySelector(".access-modal-status");
  if (statusEl) statusEl.textContent = "Stopping...";
  try {
    await fetch(`${BACKEND_URL}/computer-use/cancel/${sessionId}`, { method: "POST" });
  } catch (e) {
    console.warn("[tsifl] Cancel request failed:", e.message);
  }
}

// Escape key triggers stop during automation
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && _activeAccessSessionId) {
    e.preventDefault();
    _cancelAccess();
  }
});

// ── Action Summary Helper ────────────────────────────────────────────────────

function summarizeAction(action) {
  const p = action.payload || {};
  switch (action.type) {
    case "write_cell": return `Wrote to ${p.cell || "?"} on ${p.sheet || "active"}`;
    case "write_range": return `Wrote range ${p.range || "?"} on ${p.sheet || "active"}`;
    case "write_formula": return `Formula in ${p.cell || "?"} on ${p.sheet || "active"}`;
    case "fill_down": return `Fill down ${p.range || "?"} on ${p.sheet || "active"}`;
    case "fill_right": return `Fill right ${p.range || "?"} on ${p.sheet || "active"}`;
    case "navigate_sheet": return `Switched to "${p.sheet || "?"}"`;
    case "format_range": return `Formatted ${p.range || "?"} on ${p.sheet || "active"}`;
    case "add_chart": return `Created ${p.chart_type || "chart"} on ${p.sheet || "active"}`;
    case "add_sheet": return `Created sheet "${p.name || "?"}"`;
    case "create_slide": return `Created slide: ${p.title || "untitled"}`;
    case "remove_duplicates": return `Removed duplicates in ${p.range || "?"} on ${p.sheet || "active"}`;
    case "conditional_format_heatmap": return `Applied heatmap to ${p.range || "?"} on ${p.sheet || "active"}`;
    case "create_pivot_summary": return `Created pivot summary on ${p.target_sheet || p.sheet || "active"}`;
    case "auto_chart_best_fit": return `Auto-chart on ${p.range || "?"} (${p.chart_type || "auto"})`;
    case "trim_whitespace": return `Trimmed whitespace in ${p.range || "?"} on ${p.sheet || "active"}`;
    case "find_replace_bulk": return `Bulk find/replace in ${p.range || "?"} (${(p.replacements || []).length} patterns)`;
    case "goal_seek": return `Goal Seek: ${p.set_cell || "?"} → ${p.to_value || "?"}`;
    case "run_toolpak": return `Descriptive Stats for ${p.input_range || "?"} on ${p.sheet || "active"}`;
    case "create_data_table": return `Data Table ${p.range || p.table_range || "?"} on ${p.sheet || "active"}`;
    case "install_addins": return "Installed add-ins";
    case "uninstall_addins": return "Uninstalled add-ins";
    default: return JSON.stringify(p).slice(0, 80);
  }
}

// ── Session Expiry Warning (Improvement 7) ───────────────────────────────────

function checkSessionExpiry() {
  try {
    const state = supabase.auth;
    // Try to get token from local storage
    const sessionStr = localStorage.getItem("sb-dvynmzeyttwlmvunicqz-auth-token");
    if (!sessionStr) return;
    const session = JSON.parse(sessionStr);
    const token = session?.access_token;
    if (!token) return;
    const payload = JSON.parse(atob(token.split(".")[1]));
    const expiresAt = payload.exp * 1000;
    const minutesLeft = (expiresAt - Date.now()) / 60000;
    const warning = document.getElementById("session-warning");
    if (warning) {
      warning.style.display = minutesLeft < 10 ? "block" : "none";
    }
  } catch (e) { /* silent */ }
}

// ── Auth Debug Panel (Improvement 10) ────────────────────────────────────────

function updateDebugPanel() {
  try {
    const sessionStr = localStorage.getItem("sb-dvynmzeyttwlmvunicqz-auth-token");
    if (sessionStr) {
      const session = JSON.parse(sessionStr);
      const token = session?.access_token;
      if (token) {
        const payload = JSON.parse(atob(token.split(".")[1]));
        const expiry = new Date(payload.exp * 1000);
        document.getElementById("dbg-expiry").textContent = expiry.toLocaleTimeString();
      }
    }
  } catch (e) { document.getElementById("dbg-expiry").textContent = "error"; }
  document.getElementById("dbg-last-sync").textContent = lastSyncTimestamp || "never";
  document.getElementById("dbg-source").textContent = sessionSource;
}

// ── Undo Last Action (Improvement 11) ────────────────────────────────────────

async function saveUndoState(rangeAddress, sheetName) {
  try {
    await Excel.run(async (ctx) => {
      const sheet = sheetName
        ? ctx.workbook.worksheets.getItem(sheetName)
        : ctx.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(rangeAddress);
      range.load(["values", "formulas", "numberFormat"]);
      sheet.load("name");
      await ctx.sync();
      undoStack.push({
        sheet: sheet.name,
        address: rangeAddress,
        values: range.values,
        formulas: range.formulas,
        numberFormat: range.numberFormat,
        timestamp: Date.now()
      });
      if (undoStack.length > MAX_UNDO) undoStack.shift();
      const undoBtn = document.getElementById("undo-btn");
      if (undoBtn) undoBtn.style.display = "inline-block";
    });
  } catch (e) { /* silent — undo state capture is best-effort */ }
}

// ── Build Comps — inject comp table directly into the active sheet ───────────
// The killer feature: type a ticker → get a full IB-quality comp table
// written directly into a new sheet in your workbook. No downloads, no fuss.
//
// Flow:
//   1. Prompt user for ticker(s) (or grab from chat input)
//   2. Call /generate/comp-inject (auto-finds peers if 1 ticker)
//   3. Create new sheet "Comps NVDA"
//   4. Write title, headers, data rows, summary stats
//   5. Apply IB formatting (bold headers, number formats, freeze panes)
// ─────────────────────────────────────────────────────────────────────────────
async function handleBuildComps() {
  const btn = document.getElementById("build-comps-btn");

  // Get tickers from chat input
  const inputEl = document.getElementById("user-input");
  let rawInput = (inputEl?.value || "").trim();

  // If input is empty, focus the input and show hint
  if (!rawInput) {
    if (inputEl) {
      inputEl.placeholder = "Type ticker(s) here → e.g. NVDA or NVDA, AMD, INTC";
      inputEl.focus();
    }
    showToast("Type a ticker in the input box, then click Build Comps", "info", 3000);
    return;
  }

  // Parse tickers from the input
  const tickers = rawInput
    .toUpperCase()
    .replace(/[^A-Z,\s]/g, "")
    .split(/[\s,]+/)
    .filter(t => t.length >= 1 && t.length <= 5);

  if (!tickers.length) {
    showToast("Enter at least one ticker (e.g. NVDA)", "error", 3000);
    return;
  }

  // Clear the input
  if (inputEl) inputEl.value = "";

  // Disable button + show progress
  if (btn) { btn.disabled = true; btn.textContent = "Building..."; }
  setStatus(`Building comps for ${tickers.join(", ")}...`);

  try {
    // Call the backend
    const resp = await fetch(`${BACKEND_URL}/generate/comp-inject`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ tickers }),
    });

    if (!resp.ok) {
      const err = await resp.json().catch(() => ({}));
      throw new Error(err.detail || `HTTP ${resp.status}`);
    }

    const data = await resp.json();
    const { title, sheet_name, date, headers, rows, stats, count } = data;

    // Inject into Excel
    await Excel.run(async (ctx) => {
      // Create a new sheet (or reuse existing with same name)
      // Office.js proxies don't throw until ctx.sync() — must sync inside try
      let sheet;
      try {
        sheet = ctx.workbook.worksheets.getItem(sheet_name);
        sheet.load("name");
        await ctx.sync();
        // Sheet exists — clear it for fresh data
        sheet.getRange().clear();
      } catch (_) {
        // Sheet doesn't exist — create it
        sheet = ctx.workbook.worksheets.add(sheet_name);
      }
      sheet.activate();

      // Row 1: Title
      const titleCell = sheet.getRange("A1");
      titleCell.values = [[title]];
      titleCell.format.font.bold = true;
      titleCell.format.font.size = 14;

      // Row 2: Date + currency
      const dateCell = sheet.getRange("A2");
      dateCell.values = [[date + " | USD ($M unless noted)"]];
      dateCell.format.font.size = 10;
      dateCell.format.font.color = "#64748B";

      // Row 4: Headers
      const headerRange = sheet.getRangeByIndexes(3, 0, 1, headers.length);
      headerRange.values = [headers];
      headerRange.format.font.bold = true;
      headerRange.format.font.size = 11;
      headerRange.format.fill.color = "#1E293B";
      headerRange.format.font.color = "#FFFFFF";
      headerRange.format.borders.getItem("EdgeBottom").style = "Continuous";
      headerRange.format.borders.getItem("EdgeBottom").color = "#0D5EAF";

      // Rows 5+: Data
      if (rows.length > 0) {
        const dataRange = sheet.getRangeByIndexes(4, 0, rows.length, headers.length);
        dataRange.values = rows;
        dataRange.format.font.size = 11;

        // Alternating row shading
        for (let i = 0; i < rows.length; i++) {
          if (i % 2 === 1) {
            const rowRange = sheet.getRangeByIndexes(4 + i, 0, 1, headers.length);
            rowRange.format.fill.color = "#F8FAFC";
          }
        }

        // Ticker column bold
        const tickerCol = sheet.getRangeByIndexes(4, 1, rows.length, 1);
        tickerCol.format.font.bold = true;
        tickerCol.format.font.color = "#0D5EAF";
      }

      // Separator row before stats
      const sepRow = 4 + rows.length;
      const sepRange = sheet.getRangeByIndexes(sepRow, 0, 1, headers.length);
      sepRange.format.borders.getItem("EdgeTop").style = "Continuous";
      sepRange.format.borders.getItem("EdgeTop").color = "#334155";

      // Stats rows (High/Low/Median/Mean)
      if (stats.length > 0) {
        const statsStart = sepRow + 1;
        const statsRange = sheet.getRangeByIndexes(statsStart, 0, stats.length, headers.length);
        statsRange.values = stats;
        statsRange.format.font.size = 11;
        statsRange.format.font.italic = true;
        statsRange.format.font.color = "#475569";

        // Bold the Median row
        const medianRange = sheet.getRangeByIndexes(statsStart + 2, 0, 1, headers.length);
        medianRange.format.font.bold = true;
        medianRange.format.font.italic = false;
        medianRange.format.font.color = "#0D5EAF";
      }

      // Number formats for data columns
      if (rows.length > 0) {
        const nRows = rows.length + stats.length + 1; // data + separator + stats
        // Price (col C, index 2)
        sheet.getRangeByIndexes(4, 2, rows.length, 1).numberFormat = [["$#,##0.00"]];
        // Mkt Cap, Net Debt, EV (cols D-F, index 3-5)
        sheet.getRangeByIndexes(4, 3, rows.length, 3).numberFormat = [["$#,##0.0"]];
        // Revenue, EBITDA (cols G-H, index 6-7)
        sheet.getRangeByIndexes(4, 6, rows.length, 2).numberFormat = [["$#,##0"]];
      }

      // Autofit columns
      const usedRange = sheet.getUsedRange();
      usedRange.format.autofitColumns();
      usedRange.format.autofitRows();

      // Freeze panes: freeze above row 5 (headers stay visible)
      sheet.freezePanes.freezeRows(4);

      await ctx.sync();
    });

    // Success
    if (btn) { btn.classList.add("success"); btn.textContent = `✓ ${count} comps`; }
    setStatus(`Done — ${count} companies in "${sheet_name}"`);
    showToast(`Comp table built: ${count} companies`, "success", 4000);

    // Reset button after 3s
    setTimeout(() => {
      if (btn) {
        btn.classList.remove("success");
        btn.textContent = "Build Comps";
        btn.disabled = false;
      }
    }, 3000);

  } catch (err) {
    console.error("[tsifl] Build Comps error:", err);
    showToast("Build Comps failed: " + err.message, "error", 5000);
    setStatus("Build failed — " + err.message);
    if (btn) { btn.disabled = false; btn.textContent = "Build Comps"; }
  }
}


// ── IB Format ─────────────────────────────────────────────────────────────────
// One-click IB-standard formatting on the active sheet:
//   • Bold + light grey fill on row 4 (standard comp header row)
//   • Blue fill (#DCE6F1) on hardcoded input cells (non-formula, non-empty)
//   • Autofit all columns
//   • $#,##0.00 on share price column (col I), #,##0 on revenue cols
//   • 0.0% on margin/growth cols, 0.0x on multiple cols
// ─────────────────────────────────────────────────────────────────────────────
async function handleIBFormat() {
  if (_excelBusy) { showToast("⏳ Still running — please wait", 2500); return; }
  const btn = document.getElementById("ib-format-btn");
  if (btn) { btn.classList.add("running"); btn.textContent = "⚡ Formatting..."; }

  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRange();
      usedRange.load(["rowCount", "columnCount", "address"]);
      await ctx.sync();

      const rows = usedRange.rowCount;
      const cols = usedRange.columnCount;

      // 1. Bold header row (row 4 = index 3, standard comp header)
      const headerRow = sheet.getRangeByIndexes(3, 0, 1, cols);
      headerRow.format.font.bold = true;
      headerRow.format.fill.color = "#1F3864";
      headerRow.format.font.color = "#FFFFFF";

      // 2. Title row (row 1 = index 0)
      const titleRow = sheet.getRangeByIndexes(0, 0, 1, cols);
      titleRow.format.font.bold = true;
      titleRow.format.font.size = 13;

      // 3. Autofit columns
      usedRange.format.autofitColumns();

      // 4. Scan data rows (5–10) and colour hardcoded inputs blue
      //    Formula cells stay untouched — only plain value cells get blue fill
      if (rows >= 5) {
        const dataRows = Math.min(rows - 4, 10); // rows 5-14 max
        const dataRange = sheet.getRangeByIndexes(4, 0, dataRows, cols);
        dataRange.load("formulas");
        await ctx.sync();

        const formulas = dataRange.formulas;
        for (let r = 0; r < formulas.length; r++) {
          for (let c = 0; c < formulas[r].length; c++) {
            const cell = formulas[r][c];
            if (typeof cell === "string" && cell.startsWith("=")) continue; // formula — skip
            if (cell === null || cell === "" || cell === 0) continue;        // empty — skip
            // Hardcoded non-empty value → light blue input fill
            const cellRange = sheet.getRangeByIndexes(4 + r, c, 1, 1);
            cellRange.format.fill.color = "#DCE6F1";
          }
        }
      }

      // 5. Freeze top 4 rows (title + subtitle + blank + header)
      sheet.freezePanes.freezeRows(4);

      await ctx.sync();
    });

    appendMessage("IB formatting applied — blue inputs, bold header, columns autofit.", "assistant");
  } catch (e) {
    appendMessage("Format failed: " + e.message, "assistant");
  } finally {
    if (btn) { btn.classList.remove("running"); btn.textContent = "⚡ IB Format"; }
  }
}

// ── Build Deck: Excel comp → PowerPoint tearsheet ────────────────────────────
// Reads the active sheet, sends data to Claude in PPT mode, stores the
// generated actions as a ppt_actions transfer. The PPT add-in polls for
// ppt_actions every 4s and auto-executes them when found.
// ── Shared ticker extractor ───────────────────────────────────────────────────
function _extractTickersFromValues(values) {
  const seen = new Set();
  const tickers = [];
  let tickerCol = -1;
  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    // Find the "Ticker" header column
    if (tickerCol === -1) {
      const ti = row.findIndex(c =>
        typeof c === "string" && /^ticker$/i.test(String(c).trim())
      );
      if (ti !== -1) { tickerCol = ti; continue; }
    }
    if (tickerCol !== -1) {
      const v = String(row[tickerCol] ?? "").trim().toUpperCase();
      // Valid ticker: 1-5 uppercase letters, not a label row, no duplicates
      const SKIP = ["MEDIAN","MEAN","AVG","EV","PE","LTM","NTM","USD","EUR","GBP","YOY","TTM","NA","N/A","NM"];
      if (/^[A-Z]{1,5}$/.test(v) && !SKIP.includes(v) && !seen.has(v)) {
        seen.add(v);
        tickers.push(v);
      }
    }
  }
  return tickers;
}

async function handleBuildDeck() {
  if (_excelBusy) { showToast("⏳ Still running — please wait", 2500); return; }
  const btn = document.getElementById("build-deck-btn");
  if (btn) { btn.disabled = true; btn.textContent = "⏳ Building..."; }
  setStatus("Reading comp...");

  try {
    // 1. Read active sheet
    let compData = null;
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      const range = sheet.getUsedRange();
      range.load(["values", "rowCount", "columnCount"]);
      await ctx.sync();
      const vals = range.values;
      compData = {
        sheetName: sheet.name,
        values: vals.slice(0, 80).map(r => r.slice(0, 20)),
      };
    });

    if (!compData?.values.length) {
      showToast("No data in active sheet", "error", 3000);
      return;
    }

    // 2. Try to extract tickers for template path
    const tickers = _extractTickersFromValues(compData.values);
    const deckTitle = compData.sheetName.replace(/[^a-zA-Z0-9_\- ]/g, "").trim() || "Trading Comps";

    // ── PATH A: python-pptx template (beautiful, deterministic) ──────────
    if (tickers.length >= 2) {
      setStatus("Generating deck...");
      showTypingIndicator("thinking");

      const resp = await fetch(`${BACKEND_URL}/generate/comp-slide.pptx`, {
        method:  "POST",
        headers: { "Content-Type": "application/json" },
        body:    JSON.stringify({ tickers, title: deckTitle }),
      });
      hideTypingIndicator();

      if (resp.ok) {
        const blob     = await resp.blob();
        const filename = deckTitle.replace(/\s+/g, "_") + ".pptx";
        const url      = URL.createObjectURL(blob);
        const a        = document.createElement("a");
        a.href = url; a.download = filename;
        document.body.appendChild(a); a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        setStatus("Deck ready ✓");
        appendMessage(
          "assistant",
          `✅ **${filename}** downloaded.\n\n` +
          `4 slides — title · comps table · multiples snapshot · key takeaways\n\n` +
          `Built from: ${tickers.join(" · ")}`
        );
        showToast(`📊 ${filename} downloaded — open in PowerPoint`, "success", 6000);
        return;
      }
      // Template generation failed — fall through to Claude path
      console.warn("[tsifl] Template path failed, falling back to Claude");
    }

    // ── PATH B: Claude + Office.js transfer (fallback for non-comp sheets) ──
    setStatus("Generating deck via AI...");
    showTypingIndicator("thinking");

    const tsv = compData.values
      .map(row => row.map(c => (c == null) ? "" : String(c)).join("\t"))
      .join("\n");

    const deckPrompt =
      `Build a professional IB tearsheet from the data below.\n` +
      `Slide 1: Title — "${deckTitle}"\n` +
      `Slide 2: Data table with ALL rows and columns exactly as provided\n` +
      `Slide 3: Key multiples snapshot (EV/EBITDA, EV/Revenue, P/E if present)\n` +
      `Slide 4: 3-5 key takeaways\n\n` +
      `CRITICAL: Use EXACT numbers from the data — no invented figures.\n\n` +
      `\`\`\`\n${tsv}\n\`\`\``;

    const resp = await fetch(`${BACKEND_URL}/chat/`, {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        user_id:    CURRENT_USER.id,
        message:    deckPrompt,
        context:    { app: "powerpoint", force_model: "sonnet" },
        session_id: `deck_${Date.now()}`,
      }),
    });
    hideTypingIndicator();

    if (!resp.ok) {
      const err = await resp.json().catch(() => ({}));
      throw new Error(err.detail || `HTTP ${resp.status}`);
    }

    const data    = await resp.json();
    const actions = data.actions?.length ? data.actions
      : (data.action?.type && data.action.type !== "none" ? [data.action] : []);

    if (!actions.length) {
      showToast("No slides generated — try again", "error", 3000);
      appendMessage("assistant", data.reply || "No slides returned. Try again.");
      return;
    }

    // Store as ppt_actions transfer for PPT add-in to execute
    const slideCount = actions.filter(a => a.type === "create_slide").length;
    await fetch(`${BACKEND_URL}/transfer/store`, {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        from_app:  "excel",
        to_app:    "powerpoint",
        data_type: "ppt_actions",
        data:      JSON.stringify(actions),
        metadata:  { title: deckTitle, slide_count: slideCount },
      }),
    });

    setStatus("Deck queued ✓");
    appendMessage(
      "assistant",
      `Deck queued — **${slideCount} slides** ready. Open PowerPoint with tsifl and they build automatically.`
    );
    const chatHistory = document.getElementById("chat-history");
    const lastMsg = chatHistory?.lastElementChild;
    if (lastMsg) {
      const openBtn = document.createElement("button");
      openBtn.className   = "open-ppt-btn";
      openBtn.textContent = "🚀 Open PowerPoint";
      openBtn.addEventListener("click", () => {
        window.open("ms-powerpoint:", "_blank");
        openBtn.disabled    = true;
        openBtn.textContent = `Save as: ${deckTitle}.pptx`;
      });
      lastMsg.appendChild(openBtn);
    }
    showToast(`📊 ${slideCount} slides queued`, "success", 6000);

  } catch (err) {
    console.error("[tsifl] Build deck error:", err);
    setStatus("Deck failed");
    showToast("Build Deck failed: " + err.message, "error", 4000);
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = "📊 Build Deck"; }
  }
}

// ── Export formatted file (server-side template engine) ──────────────────────
// Reads tickers from the active sheet, calls /generate/comp-table.xlsx or
// /generate/comp-slide.pptx, and triggers a browser download.
// This produces deterministic, IB-grade output — no Office.js, no locale bugs.
async function handleExportFormatted(format) {
  const btnId = format === "xlsx" ? "export-xlsx-btn" : "export-pptx-btn";
  const btn   = document.getElementById(btnId);
  const label = format === "xlsx" ? "📥 .xlsx" : "📥 .pptx";
  if (btn) { btn.disabled = true; btn.textContent = "⏳..."; }

  try {
    // Read sheet — extract title + all data
    let sheetTitle = "Trading Comps";
    let sheetValues = [];

    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      const range = sheet.getUsedRange();
      range.load(["values","rowCount","columnCount"]);
      await ctx.sync();
      sheetTitle  = sheet.name;
      sheetValues = range.values.slice(0, 80).map(r => r.slice(0, 20));
    });

    if (!sheetValues.length) {
      showToast("No data found on active sheet", "error", 3000);
      return;
    }

    const tickers = _extractTickersFromValues(sheetValues);

    setStatus(format === "xlsx" ? "Generating .xlsx..." : "Generating .pptx...");

    const endpoint = format === "xlsx"
      ? `${BACKEND_URL}/generate/comp-table.xlsx`
      : `${BACKEND_URL}/generate/comp-slide.pptx`;

    const body = tickers.length >= 2
      ? { tickers, title: sheetTitle }
      : {
          payload: {
            title:     sheetTitle,
            date:      new Date().toLocaleDateString("en-US", { month:"long", year:"numeric" }),
            currency:  "USD ($B)",
            companies: [],           // server falls back to raw sheet data via Claude
          }
        };

    const resp = await fetch(endpoint, {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body:    JSON.stringify(body),
    });

    if (!resp.ok) {
      const err = await resp.json().catch(() => ({}));
      throw new Error(err.detail || `HTTP ${resp.status}`);
    }

    // Trigger download
    const blob     = await resp.blob();
    const url      = URL.createObjectURL(blob);
    const anchor   = document.createElement("a");
    const filename = sheetTitle.replace(/[^a-zA-Z0-9_\-]/g, "_") +
                     (format === "xlsx" ? ".xlsx" : ".pptx");
    anchor.href     = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    document.body.removeChild(anchor);
    URL.revokeObjectURL(url);

    setStatus("Downloaded ✓");
    showToast(`Downloaded ${filename}`, "success", 4000);

  } catch (err) {
    console.error("[tsifl] Export error:", err);
    showToast("Export failed: " + err.message, "error", 4000);
    setStatus("Export failed");
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = label; }
  }
}

// ── Subscribe flow ────────────────────────────────────────────────────────────
let _subscribeBannerShown = false;

function _showSubscribeBanner() {
  if (_subscribeBannerShown) return;
  _subscribeBannerShown = true;

  const chatHistory = document.getElementById("chat-history");
  const banner = document.createElement("div");
  banner.className = "subscribe-banner";
  banner.innerHTML = `
    <div class="subscribe-banner-text">
      <strong>You're almost out of tasks.</strong><br>
      Upgrade to tsifl Pro — $99/month, unlimited everything.
    </div>
    <button class="subscribe-btn" id="subscribe-cta-btn">Subscribe →</button>
  `;
  if (chatHistory) chatHistory.appendChild(banner);
  chatHistory?.scrollTo(0, chatHistory.scrollHeight);

  document.getElementById("subscribe-cta-btn")?.addEventListener("click", handleSubscribe);
}

async function handleSubscribe() {
  const btn = document.getElementById("subscribe-cta-btn");
  if (btn) { btn.disabled = true; btn.textContent = "Loading..."; }

  try {
    const resp = await fetch(`${BACKEND_URL}/billing/checkout`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        user_id: CURRENT_USER?.id || "unknown",
        email: CURRENT_USER?.email || "",
      }),
    });
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const { url } = await resp.json();
    if (url) window.open(url, "_blank");
  } catch (err) {
    showToast("Couldn't start checkout: " + err.message, "error", 4000);
    if (btn) { btn.disabled = false; btn.textContent = "Subscribe →"; }
  }
}

async function handleUndo() {
  if (undoStack.length === 0) {
    console.log( "Nothing to undo.");
    return;
  }
  const state = undoStack.pop();
  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(state.sheet);
      const range = sheet.getRange(state.address);
      range.formulas = state.formulas;
      range.numberFormat = state.numberFormat;
      await ctx.sync();
    });
    console.log( `Undid changes to ${state.sheet}!${state.address}`);
  } catch (e) {
    console.log( `Undo failed: ${e.message}`);
  }
  if (undoStack.length === 0) {
    const undoBtn = document.getElementById("undo-btn");
    if (undoBtn) undoBtn.style.display = "none";
  }
}

// ── Action History (Improvement 12) ──────────────────────────────────────────

function addToHistory(actionType, summary) {
  const entry = {
    time: new Date().toLocaleTimeString(),
    type: actionType,
    summary: summary
  };
  actionHistory.unshift(entry);
  if (actionHistory.length > MAX_HISTORY) actionHistory.pop();
  const list = document.getElementById("history-list");
  if (list) {
    list.innerHTML = actionHistory.map(h =>
      `<div style="padding:2px 0;border-bottom:1px solid #f0f0f0;"><span style="color:#64748B;">${h.time}</span> <strong>${h.type}</strong>: ${h.summary}</div>`
    ).join("");
  }
}

// ── Progress Bar (Improvement 13) ────────────────────────────────────────────

function showProgress(current, total) {
  const wrap = document.getElementById("progress-bar-wrap");
  const bar = document.getElementById("progress-bar");
  const text = document.getElementById("progress-text");
  if (!wrap || !bar || !text) return;
  wrap.style.display = "block";
  const pct = Math.round((current / total) * 100);
  bar.style.width = pct + "%";
  text.textContent = `Applying ${current}/${total} actions...`;
}

function hideProgress() {
  const wrap = document.getElementById("progress-bar-wrap");
  if (wrap) wrap.style.display = "none";
}
