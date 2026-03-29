/**
 * tsifl Excel Add-in
 * Full auth + chat + comprehensive Excel action execution.
 */

import "./taskpane.css";
import { getCurrentUser, signIn, signUp, signOut } from "./auth.js";

const BACKEND_URL  = "https://focused-solace-production-6839.up.railway.app";
const LOCAL_URL    = "/local-api";              // proxied through webpack dev server (avoids HTTPS mixed content)
const PREFS_KEY    = "tsifl_preferences";
const BUILD_VER    = "v40";  // bump this on every deploy so user can confirm fresh code

let CURRENT_USER       = null;
let lastNavigatedSheet = null;   // tracks sheet after navigate_sheet so writes auto-target it
let pendingImages      = [];     // base64 images queued for next message

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
}

function showChatScreen(user) {
  document.getElementById("login-screen").style.display = "none";
  document.getElementById("chat-screen").style.display  = "flex";
  document.getElementById("user-bar").textContent = `${user.email} · ${BUILD_VER}`;

  document.getElementById("submit-btn").addEventListener("click", handleSubmit);
  document.getElementById("user-input").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleSubmit(); }
  });

  // Image attachment — file picker
  document.getElementById("attach-btn").addEventListener("click", () => {
    document.getElementById("image-input").click();
  });
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
    let count = 0;
    for (const file of files) {
      if (!file.type.startsWith("image/")) continue;
      readImageAsBase64(file);
      count++;
    }
    if (count === 0) setStatus("No image files found in drop");
  });
  document.getElementById("logout-btn").addEventListener("click", async () => {
    await signOut();
    CURRENT_USER = null;
    document.getElementById("chat-history").innerHTML = "";
    showLoginScreen();
  });

  setStatus("Connected · " + user.email);
}

// ── Auth Handlers ─────────────────────────────────────────────────────────────

async function handleSignIn() {
  const email    = document.getElementById("auth-email").value.trim();
  const password = document.getElementById("auth-password").value;
  const errEl    = document.getElementById("auth-error");
  errEl.textContent = "";
  document.getElementById("login-btn").textContent = "Signing in...";
  const { user, error } = await signIn(email, password);
  document.getElementById("login-btn").textContent = "Sign In";
  if (error) { errEl.textContent = error.message; return; }
  if (!user)  { errEl.textContent = "Check your email to confirm your account first."; return; }
  CURRENT_USER = user;
  await saveUserConfig(user);
  showChatScreen(user);
}

async function handleSignUp() {
  const email    = document.getElementById("auth-email").value.trim();
  const password = document.getElementById("auth-password").value;
  const errEl    = document.getElementById("auth-error");
  errEl.textContent = "";
  document.getElementById("signup-btn").textContent = "Creating account...";
  const { user, error } = await signUp(email, password);
  document.getElementById("signup-btn").textContent = "Create Account";
  if (error) { errEl.textContent = error.message; return; }
  errEl.style.color = "#2ecc71";
  errEl.textContent = "Account created! Check your email to confirm, then sign in.";
}

// ── Image Handling ────────────────────────────────────────────────────────────

function handleImageSelect(e) {
  const files = Array.from(e.target.files);
  for (const file of files) {
    if (!file.type.startsWith("image/")) continue;
    readImageAsBase64(file);
  }
  e.target.value = "";  // reset so same file can be re-selected
}

function handleImagePaste(e) {
  const items = Array.from(e.clipboardData?.items || []);
  for (const item of items) {
    if (!item.type.startsWith("image/")) continue;
    const file = item.getAsFile();
    if (file) readImageAsBase64(file);
  }
}

function readImageAsBase64(file) {
  setStatus(`📎 reading ${file.name} (${file.type}, ${Math.round(file.size/1024)}KB)...`);
  const reader = new FileReader();
  reader.onload = () => {
    const base64 = reader.result;  // data:image/png;base64,...
    const mediaType = file.type || "image/png";
    const data = base64.split(",")[1];  // strip the data:... prefix
    pendingImages.push({ media_type: mediaType, data });
    setStatus(`✅ image captured · ${data.length} chars base64 · ${pendingImages.length} pending`);
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
    attachBtn.title = "Attach image";
    return;
  }

  // Update attach button to show count
  attachBtn.textContent = `${pendingImages.length}`;
  attachBtn.title = `${pendingImages.length} image${pendingImages.length > 1 ? "s" : ""} attached — click to add more`;
  setStatus(`${pendingImages.length} image${pendingImages.length > 1 ? "s" : ""} attached`);

  bar.style.display = "flex";
  pendingImages.forEach((img, i) => {
    const wrapper = document.createElement("div");
    wrapper.className = "image-preview-item";

    // Render to canvas (bypasses CSP restrictions on data:/blob: img src)
    renderImageToCanvas(img.data, img.media_type, 48, 48).then(canvas => {
      if (canvas) {
        canvas.style.borderRadius = "4px";
        canvas.style.border = "1px solid var(--border)";
        wrapper.insertBefore(canvas, wrapper.firstChild);
      }
    });

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

  try {
    const excelContext = await getExcelContext();
    setStatus("Thinking...");

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
    appendMessage("assistant", data.reply);

    if (data.tasks_remaining >= 0) {
      document.getElementById("tasks-remaining").textContent =
        `${data.tasks_remaining} tasks left`;
    }

    const allActions = [];
    if (data.actions && data.actions.length > 0) allActions.push(...data.actions);
    else if (data.action && data.action.type && data.action.type !== "none") allActions.push(data.action);

    // Always show how many actions Claude returned — helps diagnose "no changes" issues
    appendMessage("action", `📋 ${BUILD_VER} · Claude returned ${allActions.length} action${allActions.length !== 1 ? "s" : ""}` +
      (allActions.length > 0 ? `: ${[...new Set(allActions.map(a => a.type))].join(", ")}` : " — text-only reply"));

    if (allActions.length > 0) {
      setStatus(`Applying ${allActions.length} action${allActions.length > 1 ? "s" : ""}...`);
      let applied = 0;
      let failed  = 0;
      for (const action of allActions) {
        try {
          await executeAction(action);
          applied++;
        } catch (err) {
          failed++;
          appendMessage("action", `⚠️ ${action.type} → ${err.message} | payload: ${JSON.stringify(action.payload || {}).slice(0, 120)}`);
        }
      }
      appendMessage("action", `✅ ${applied} applied${failed > 0 ? ` · ⚠️ ${failed} failed` : ""}`);
    }

    setStatus("Done");
  } catch (err) {
    appendMessage("assistant", "⚠️ Could not reach tsifl backend.");
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
        const values   = used.values.slice(0, 60).map(r => r.slice(0, 26));
        const formulas = used.formulas ? used.formulas.slice(0, 60).map(r => r.slice(0, 26)) : values;

        if (name === activeName) {
          activeSheetData     = values;
          activeSheetFormulas = formulas;
          activeUsedRange     = used.address;
        }

        // Non-active sheets: 20-row preview with formulas so Claude can see empty cells
        // (active sheet already gets full 60-row data above)
        const PREVIEW_ROWS = name === activeName ? 60 : 20;
        sheetSummaries.push({
          name,
          used_range:       used.address,
          rows:             used.rowCount,
          cols:             used.columnCount,
          preview:          values.slice(0, PREVIEW_ROWS),
          preview_formulas: formulas.slice(0, PREVIEW_ROWS),
        });
      }

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
      const val   = payload.formula ?? payload.value ?? "";
      if (typeof val === "string" && val.startsWith("=")) {
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
  // Note: ensure2D() handles flat 1D arrays that Claude sometimes sends
  else if (type === "write_range") {
    await Excel.run(async (ctx) => {
      const { sheet: s, addr } = splitAddr(payload.range, payload.sheet);
      const sheet = getSheet(ctx, s);
      const range = sheet.getRange(addr);
      if (payload.formulas) {
        range.formulas = ensure2D(payload.formulas);
      } else if (payload.values) {
        const vals = ensure2D(payload.values);
        // Auto-detect formulas in the normalised 2D array
        const hasFormulas = vals.some(r => Array.isArray(r) && r.some(v => typeof v === "string" && v.startsWith("=")));
        if (hasFormulas) range.formulas = vals;
        else              range.values  = vals;
      }
      _applyFormat(range, payload);
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
      range.formulas = [[payload.formula]];
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
      const range = sheet.getRange(addr);
      _applyFormat(range, payload);
      await ctx.sync();
    });
  }

  // ── set_number_format ──────────────────────────────────────────────────────
  else if (type === "set_number_format") {
    await Excel.run(async (ctx) => {
      const { sheet: s, addr } = splitAddr(payload.range, payload.sheet);
      const sheet = getSheet(ctx, s);
      const range = sheet.getRange(addr);
      // Load dimensions first — numberFormat must be a 2D array matching range size
      range.load("rowCount,columnCount");
      await ctx.sync();
      const fmt = Array.from({ length: range.rowCount }, () =>
        Array.from({ length: range.columnCount }, () => payload.format)
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

  // ── launch_app ──────────────────────────────────────────────────────────────
  else if (type === "launch_app") {
    try {
      const resp = await fetch(`${BACKEND_URL}/launch-app`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ app_name: payload.app_name }),
      });
      const result = await resp.json();
      appendMessage("action", `launch_app: ${result.message || "Requested"}`);
    } catch (e) {
      appendMessage("action", `launch_app: ${e.message}`);
    }
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
    // Accept single string or 2D array
    range.numberFormat = typeof p.number_format === "string"
      ? [[p.number_format]]
      : p.number_format;
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

function appendMessage(role, text, images) {
  const history = document.getElementById("chat-history");
  const div     = document.createElement("div");
  div.className   = `message ${role}`;

  // Add text as a text node
  const textNode = document.createElement("span");
  textNode.textContent = text;
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

function setStatus(text)           { document.getElementById("status-bar").textContent = text; }
function setSubmitEnabled(enabled) { document.getElementById("submit-btn").disabled = !enabled; }
