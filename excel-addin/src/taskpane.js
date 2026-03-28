/**
 * Tsifulator.ai Excel Add-in
 * Full auth + chat + Excel action execution.
 */

import "./taskpane.css";
import { getCurrentUser, signIn, signUp, signOut } from "./auth.js";

const BACKEND_URL = "https://focused-solace-production-6839.up.railway.app";

let CURRENT_USER = null;

// ── Boot ─────────────────────────────────────────────────────────────────────

Office.onReady(async () => {
  // Pin the taskpane so it reopens automatically every time Excel starts
  try {
    await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
  } catch (_) {}

  CURRENT_USER = await getCurrentUser();

  if (CURRENT_USER) {
    showChatScreen(CURRENT_USER);
  } else {
    showLoginScreen();
  }
});

// ── Screens ──────────────────────────────────────────────────────────────────

function showLoginScreen() {
  document.getElementById("login-screen").style.display = "flex";
  document.getElementById("chat-screen").style.display  = "none";

  document.getElementById("login-btn").addEventListener("click",  handleSignIn);
  document.getElementById("signup-btn").addEventListener("click", handleSignUp);

  // Allow Enter key on password field
  document.getElementById("auth-password").addEventListener("keydown", (e) => {
    if (e.key === "Enter") handleSignIn();
  });
}

function showChatScreen(user) {
  document.getElementById("login-screen").style.display = "none";
  document.getElementById("chat-screen").style.display  = "flex";

  // Show user email in bar
  document.getElementById("user-bar").textContent = user.email;

  // Wire up chat
  document.getElementById("submit-btn").addEventListener("click", handleSubmit);
  document.getElementById("user-input").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleSubmit(); }
  });

  // Logout
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

  if (error) {
    errEl.textContent = error.message;
    return;
  }

  if (!user) {
    errEl.textContent = "Please check your email and confirm your account first, then sign in.";
    return;
  }

  CURRENT_USER = user;

  // Save user ID for RStudio addin to pick up
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

  if (error) {
    errEl.textContent = error.message;
    return;
  }

  errEl.style.color = "#2ecc71";
  errEl.textContent = "Account created! Check your email to confirm, then sign in.";
}

// ── Chat ──────────────────────────────────────────────────────────────────────

async function handleSubmit() {
  const input   = document.getElementById("user-input");
  const message = input.value.trim();
  if (!message || !CURRENT_USER) return;

  input.value = "";
  setSubmitEnabled(false);
  appendMessage("user", message);
  setStatus("Reading spreadsheet...");

  try {
    const excelContext = await getExcelContext();
    setStatus("Thinking...");

    const response = await fetch(`${BACKEND_URL}/chat/`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        user_id:    CURRENT_USER.id,   // Real Supabase user UUID
        message:    message,
        context:    excelContext,
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

    // Execute actions — collect from either array or single action field
    const allActions = [];
    if (data.actions && data.actions.length > 0) {
      allActions.push(...data.actions);
    } else if (data.action && data.action.type && data.action.type !== "none") {
      allActions.push(data.action);
    }

    if (allActions.length > 0) {
      setStatus(`Applying ${allActions.length} change${allActions.length > 1 ? "s" : ""}...`);
      for (const action of allActions) {
        try {
          await executeAction(action);
        } catch (actionErr) {
          appendMessage("action", `⚠️ ${action.type} failed: ${actionErr.message}`);
        }
      }
      appendMessage("action", `✅ ${allActions.length} change${allActions.length > 1 ? "s" : ""} applied to sheet`);
    }

    setStatus("Done");
  } catch (err) {
    appendMessage("assistant", "⚠️ Could not reach Tsifulator backend. Is the server running?");
    setStatus("Disconnected");
  } finally {
    setSubmitEnabled(true);
  }
}

// ── Excel Context ─────────────────────────────────────────────────────────────

async function getExcelContext() {
  return new Promise((resolve) => {
    Excel.run(async (ctx) => {
      const sheet    = ctx.workbook.worksheets.getActiveWorksheet();
      const selected = ctx.workbook.getSelectedRange();
      const used     = sheet.getUsedRangeOrNullObject();

      sheet.load("name");
      selected.load(["address", "values"]);
      used.load(["address", "values", "rowCount", "columnCount"]);
      await ctx.sync();

      let sheetData = [], usedAddress = "empty";
      if (!used.isNullObject && used.rowCount > 0) {
        sheetData   = used.values;
        usedAddress = used.address;
      }

      resolve({
        app:            "excel",
        sheet:          sheet.name,
        selected_cell:  selected.address,
        selected_value: selected.values?.[0]?.[0] ?? null,
        used_range:     usedAddress,
        sheet_data:     sheetData,
      });
    }).catch(() => resolve({ app: "excel" }));
  });
}

// ── Excel Actions ─────────────────────────────────────────────────────────────

async function executeAction(action) {
  const { type, payload } = action;
  if (!type || !payload) return;

  if (type === "write_cell") {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(payload.cell);
      range.values = [[payload.value ?? ""]];
      if (payload.bold)       range.format.font.bold    = true;
      if (payload.color)      range.format.fill.color   = payload.color;
      if (payload.font_color) range.format.font.color   = payload.font_color;
      await ctx.sync();
    });
  }

  else if (type === "write_range") {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(payload.range);
      range.values = payload.values;
      if (payload.bold)       range.format.font.bold    = true;
      if (payload.color)      range.format.fill.color   = payload.color;
      if (payload.font_color) range.format.font.color   = payload.font_color;
      await ctx.sync();
    });
  }

  else if (type === "format_range") {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(payload.range);
      if (payload.bold !== undefined) range.format.font.bold  = payload.bold;
      if (payload.color)              range.format.fill.color = payload.color;
      if (payload.font_color)         range.format.font.color = payload.font_color;
      if (payload.number_format)      range.numberFormat      = [[payload.number_format]];
      await ctx.sync();
    });
  }

  else if (type === "autofit") {
    await Excel.run(async (ctx) => {
      const sheet    = ctx.workbook.worksheets.getActiveWorksheet();
      const used     = sheet.getUsedRangeOrNullObject();
      used.load("isNullObject");
      await ctx.sync();
      if (!used.isNullObject) {
        used.format.autofitColumns();
        used.format.autofitRows();
        await ctx.sync();
      }
    });
  }
}

// ── Config Sync (shares user ID with RStudio addin) ──────────────────────────

async function saveUserConfig(user) {
  /**
   * Writes user ID to backend so RStudio addin can use the same identity.
   * This is what gives R and Excel the same shared memory.
   */
  try {
    await fetch(`${BACKEND_URL}/auth/set-user`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ user_id: user.id, email: user.email }),
    });
  } catch (_) {}  // Non-critical — R addin falls back gracefully
}

// ── UI Helpers ────────────────────────────────────────────────────────────────

function appendMessage(role, text) {
  const history = document.getElementById("chat-history");
  const div     = document.createElement("div");
  div.className = `message ${role}`;
  div.textContent = text;
  history.appendChild(div);
  history.scrollTop = history.scrollHeight;
}

function setStatus(text)          { document.getElementById("status-bar").textContent = text; }
function setSubmitEnabled(enabled) { document.getElementById("submit-btn").disabled = !enabled; }
