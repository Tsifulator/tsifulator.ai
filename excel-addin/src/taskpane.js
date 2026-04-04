/**
 * tsifl Excel Add-in
 * Full auth + chat + comprehensive Excel action execution.
 */

import "./taskpane.css";
import { getCurrentUser, signIn, signUp, signOut, resetPassword, supabase, syncSessionToBackend } from "./auth.js";

const BACKEND_URL  = "https://focused-solace-production-6839.up.railway.app";
const LOCAL_URL    = "/local-api";              // proxied through webpack dev server (avoids HTTPS mixed content)
const PREFS_KEY    = "tsifl_preferences";
const BUILD_VER    = "v45";  // bump this on every deploy so user can confirm fresh code

let CURRENT_USER       = null;
let lastNavigatedSheet = null;   // tracks sheet after navigate_sheet so writes auto-target it
let pendingImages      = [];     // base64 images queued for next message

// Undo stack (Improvement 11) — stores cell states before actions
const undoStack = [];
const MAX_UNDO = 5;

// Action history (Improvement 12)
const actionHistory = [];
const MAX_HISTORY = 20;

// Debug state (Improvement 10)
let lastSyncTimestamp = null;
let sessionSource = "none";

// ── Boot ─────────────────────────────────────────────────────────────────────

Office.onReady(async () => {
  try { await Office.addin.setStartupBehavior(Office.StartupBehavior.load); } catch (_) {}

  CURRENT_USER = await getCurrentUser();
  if (CURRENT_USER) showChatScreen(CURRENT_USER);
  else               showLoginScreen();
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
  // Show/hide password toggle (Improvement 9)
  const togglePw = document.getElementById("toggle-pw-btn");
  if (togglePw) {
    togglePw.addEventListener("click", () => {
      const pwInput = document.getElementById("auth-password");
      pwInput.type = pwInput.type === "password" ? "text" : "password";
    });
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
  document.getElementById("signup-btn").textContent = "Creating account...";
  const { user, error } = await signUp(email, password);
  document.getElementById("signup-btn").textContent = "Create Account";
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

// ── Chat ──────────────────────────────────────────────────────────────────────

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

  try {
    const excelContext = await getExcelContext();
    setStatus("Thinking...");
    showTypingIndicator("thinking");

    const response = await fetch(`${BACKEND_URL}/chat/`, {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body:    JSON.stringify({
        user_id: CURRENT_USER.id,
        message: message,
        context: excelContext,
        images:  images.length > 0 ? images : undefined,
      }),
    });

    if (!response.ok) {
      const err = await response.json();
      appendMessage("assistant", `⚠️ ${err.detail}`);
      setStatus("Error");
      return;
    }

    const data = await response.json();
    hideTypingIndicator();
    appendMessage("assistant", data.reply);

    if (data.tasks_remaining >= 0) {
      document.getElementById("tasks-remaining").textContent =
        `${data.tasks_remaining} tasks left`;
    }

    const allActions = [];
    if (data.actions && data.actions.length > 0) allActions.push(...data.actions);
    else if (data.action && data.action.type && data.action.type !== "none") allActions.push(data.action);

    if (allActions.length > 0) {
      setStatus(`Applying ${allActions.length} action${allActions.length > 1 ? "s" : ""}...`);
      showTypingIndicator("applying");
      if (allActions.length > 2) showProgress(0, allActions.length);
      await refreshKnownSheets(); // Cache sheet names for formula validation
      _strippedFormulaCount = 0;  // reset formula strip counter
      let applied = 0;
      let failed  = 0;
      let failedNames = [];
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

            await ctx.sync();
            console.log("[tsifl] Homework format safety net applied: Comma Style + Percent Style");
          });
        } catch (e) { console.warn("[tsifl] Format safety net failed:", e.message); }
      }
    }

    setStatus("Done");
  } catch (err) {
    hideTypingIndicator();
    appendMessage("assistant", "Could not reach tsifl backend.");
    setStatus("Disconnected");
  } finally {
    setSubmitEnabled(true);
  }
}

// ── Excel Context ─────────────────────────────────────────────────────────────

async function getExcelContext() {
  return new Promise((resolve) => {
    Excel.run(async (ctx) => {
      const wb     = ctx.workbook;
      const sheets = wb.worksheets;
      sheets.load("items/name");
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
        // Active sheet gets 200 rows so Claude can compute values itself; others get 60
        const MAX_ROWS = name === activeName ? 200 : 60;
        const values   = used.values.slice(0, MAX_ROWS).map(r => r.slice(0, 26));
        const formulas = used.formulas ? used.formulas.slice(0, MAX_ROWS).map(r => r.slice(0, 26)) : values;

        if (name === activeName) {
          activeSheetData     = values;
          activeSheetFormulas = formulas;
          activeUsedRange     = used.address;
        }

        // Non-active sheets: 20-row preview; active sheet gets full 200-row data
        const PREVIEW_ROWS = name === activeName ? 200 : 20;
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
  return arr.map(row => row.map(v => sanitizeFormula(v)));
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
]);

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

      _applyFormat(sized, payload);
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
      const fmtStr = payload.format || payload.number_format || "General";
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
  // Autofit specific columns: { sheet?, columns: ["A","B","F"] } or { column: "F" }
  else if (type === "autofit_columns") {
    await Excel.run(async (ctx) => {
      const sheet   = getSheet(ctx, payload.sheet);
      const cols    = payload.columns || (payload.column ? [payload.column] : null);
      if (cols) {
        for (const col of cols) {
          const colRange = sheet.getRange(`${col}:${col}`);
          colRange.format.autofitColumns();
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
        appendMessage("assistant", "No R plot found. Generate a plot in RStudio first, then try again.");
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
}

// ── Format helper (shared by write_cell, write_range, format_range) ──────────

function _applyFormat(range, p) {
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
    // number_format needs to be a 2D array matching the range dimensions.
    // Best-effort: try setting it; if it fails (size mismatch), skip silently.
    try {
      if (typeof p.number_format === "string") {
        range.numberFormat = [[p.number_format]];
      } else {
        range.numberFormat = p.number_format;
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

function appendMessage(role, text, images) {
  const history = document.getElementById("chat-history");
  const div     = document.createElement("div");
  div.className   = `message ${role}`;

  // Add text — use markdown for assistant messages
  const textNode = document.createElement("span");
  if (role === "assistant" && text) {
    textNode.innerHTML = renderMarkdown(text);
  } else {
    textNode.textContent = text;
  }
  div.appendChild(textNode);

  // Show image thumbnails in user messages (like Claude's inline preview)
  if (images && images.length > 0) {
    for (const img of images) {
      const container = document.createElement("div");
      container.style.cssText = "margin-top:8px;";
      div.appendChild(container);

      // Render to canvas (bypasses CSP restrictions on img src in Office.js webview)
      renderImageToCanvas(img.data, img.media_type, 280, 180).then(canvas => {
        if (canvas) {
          canvas.style.borderRadius = "8px";
          canvas.style.border = "1px solid var(--border)";
          canvas.style.display = "block";
          canvas.style.maxWidth = "100%";
          canvas.style.boxShadow = "0 1px 3px rgba(0,0,0,0.08)";
          container.appendChild(canvas);
        } else {
          // Fallback: show badge if canvas fails
          const badge = document.createElement("div");
          badge.className = "image-badge";
          badge.textContent = `📷 Image attached`;
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
