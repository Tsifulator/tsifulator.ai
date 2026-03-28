/**
 * tsifl Excel Add-in
 * Full auth + chat + comprehensive Excel action execution.
 */

import "./taskpane.css";
import { getCurrentUser, signIn, signUp, signOut } from "./auth.js";

const BACKEND_URL = "https://focused-solace-production-6839.up.railway.app";
const PREFS_KEY   = "tsifl_preferences";

let CURRENT_USER       = null;
let lastNavigatedSheet = null;   // tracks sheet after navigate_sheet so writes auto-target it

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
  document.getElementById("user-bar").textContent = user.email;

  document.getElementById("submit-btn").addEventListener("click", handleSubmit);
  document.getElementById("user-input").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleSubmit(); }
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

// ── Chat ──────────────────────────────────────────────────────────────────────

async function handleSubmit() {
  const input   = document.getElementById("user-input");
  const message = input.value.trim();
  if (!message || !CURRENT_USER) return;

  lastNavigatedSheet = null;   // reset cross-action sheet tracking for this request
  input.value = "";
  setSubmitEnabled(false);
  appendMessage("user", message);
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

    if (allActions.length > 0) {
      setStatus(`Applying ${allActions.length} action${allActions.length > 1 ? "s" : ""}...`);
      let applied = 0;
      for (const action of allActions) {
        try {
          await executeAction(action);
          applied++;
        } catch (err) {
          appendMessage("action", `⚠️ ${action.type} failed: ${err.message}`);
        }
      }
      if (applied > 0) {
        appendMessage("action", `✅ ${applied} action${applied > 1 ? "s" : ""} applied`);
      }
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

        // For non-active sheets: only first 8 rows as structural preview
        sheetSummaries.push({
          name,
          used_range: used.address,
          rows:       used.rowCount,
          cols:       used.columnCount,
          preview:    values.slice(0, 8),
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

/** Get column letter from a cell address like "B4" → "B" */
function cellToCol(addr) {
  return addr.replace(/[^A-Za-z]/g, "").toUpperCase();
}

/** Get the worksheet by name, or active sheet if name is omitted */
function getSheet(ctx, sheetName) {
  if (sheetName) return ctx.workbook.worksheets.getItem(sheetName);
  return ctx.workbook.worksheets.getActiveWorksheet();
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
      ws.activate();
      await ctx.sync();
    });
  }

  // ── write_cell ─────────────────────────────────────────────────────────────
  // Supports: cell, value OR formula, sheet?, bold?, color?, font_color?,
  //           number_format?, font_size?, font_name?
  else if (type === "write_cell") {
    await Excel.run(async (ctx) => {
      const sheet = getSheet(ctx, payload.sheet);
      const range = sheet.getRange(payload.cell);
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
  else if (type === "write_range") {
    await Excel.run(async (ctx) => {
      const sheet = getSheet(ctx, payload.sheet);
      const range = sheet.getRange(payload.range);
      if (payload.formulas) {
        range.formulas = payload.formulas;
      } else if (payload.values) {
        // Auto-detect formulas in values array
        const hasFormulas = payload.values.some(r => r.some(v => typeof v === "string" && v.startsWith("=")));
        if (hasFormulas) range.formulas = payload.values;
        else              range.values  = payload.values;
      }
      _applyFormat(range, payload);
      await ctx.sync();
    });
  }

  // ── write_formula ──────────────────────────────────────────────────────────
  // Write a single formula to a cell — explicit formula action
  else if (type === "write_formula") {
    await Excel.run(async (ctx) => {
      const sheet = getSheet(ctx, payload.sheet);
      const range = sheet.getRange(payload.cell);
      range.formulas = [[payload.formula]];
      _applyFormat(range, payload);
      await ctx.sync();
    });
  }

  // ── fill_down ──────────────────────────────────────────────────────────────
  // Copy formula from first row of range down to the rest
  // payload: { sheet?, source: "D5", target: "D5:D40" }
  //      OR  { sheet?, range: "D5:D40" }  (source = first row)
  else if (type === "fill_down") {
    await Excel.run(async (ctx) => {
      const sheet      = getSheet(ctx, payload.sheet);
      const fullRange  = payload.range  || payload.target;
      const sourceAddr = payload.source || fullRange.split(":")[0];
      const source     = sheet.getRange(sourceAddr);
      const dest       = sheet.getRange(fullRange);
      // copyFrom handles relative reference adjustment automatically
      dest.copyFrom(source, Excel.RangeCopyType.formulas, false, false);
      await ctx.sync();
    });
  }

  // ── fill_right ─────────────────────────────────────────────────────────────
  else if (type === "fill_right") {
    await Excel.run(async (ctx) => {
      const sheet  = getSheet(ctx, payload.sheet);
      const source = sheet.getRange(payload.source);
      const dest   = sheet.getRange(payload.range);
      dest.copyFrom(source, Excel.RangeCopyType.formulas, false, false);
      await ctx.sync();
    });
  }

  // ── copy_range ─────────────────────────────────────────────────────────────
  // Copy values + formulas + formats from one range to another
  else if (type === "copy_range") {
    await Excel.run(async (ctx) => {
      const sheet  = getSheet(ctx, payload.sheet);
      const source = sheet.getRange(payload.from);
      const dest   = sheet.getRange(payload.to);
      dest.copyFrom(source, Excel.RangeCopyType.all, false, false);
      await ctx.sync();
    });
  }

  // ── create_named_range ─────────────────────────────────────────────────────
  // payload: { name, range, sheet? }
  // Deletes any existing name with the same name before recreating (idempotent)
  else if (type === "create_named_range") {
    await Excel.run(async (ctx) => {
      const sheet    = getSheet(ctx, payload.sheet);
      const rangeObj = sheet.getRange(payload.range);
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
      const sheet    = getSheet(ctx, payload.sheet);
      const rng      = sheet.getRange(payload.range);

      // Compute 0-based key index within the range
      const rangeStart  = payload.range.split(":")[0];            // e.g. "A4"
      const rangeColLtr = cellToCol(rangeStart);                   // e.g. "A"
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
  // Enhanced: supports sheet?, font_size?, font_name?, number_format (single or array)
  else if (type === "format_range") {
    await Excel.run(async (ctx) => {
      const sheet = getSheet(ctx, payload.sheet);
      const range = sheet.getRange(payload.range);
      _applyFormat(range, payload);
      await ctx.sync();
    });
  }

  // ── set_number_format ──────────────────────────────────────────────────────
  else if (type === "set_number_format") {
    await Excel.run(async (ctx) => {
      const sheet = getSheet(ctx, payload.sheet);
      const range = sheet.getRange(payload.range);
      range.numberFormat = [[payload.format]];
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

  // ── save_preference ────────────────────────────────────────────────────────
  // Claude calls this when it learns the user prefers a specific style
  else if (type === "save_preference") {
    savePreferences(payload);
    // No Excel run needed — just localStorage
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

function appendMessage(role, text) {
  const history = document.getElementById("chat-history");
  const div     = document.createElement("div");
  div.className   = `message ${role}`;
  div.textContent = text;
  history.appendChild(div);
  history.scrollTop = history.scrollHeight;
}

function setStatus(text)           { document.getElementById("status-bar").textContent = text; }
function setSubmitEnabled(enabled) { document.getElementById("submit-btn").disabled = !enabled; }
