/**
 * tsifl — Side Panel Script
 *
 * Complete code paths:
 *
 * AUTH: checkAuth() → chrome.storage.local → Supabase refresh → showChat()
 *       If local fails → restoreFromBackend() → GET /auth/get-session → Supabase refresh → showChat()
 *       If both fail → showLogin()
 *
 * CHAT: handleSubmit() → getContext() → POST /chat/ → appendMessage() → executeAction()
 *
 * ACTIONS: executeAction() → chrome.runtime.sendMessage({action:"execute_browser_action"})
 *          → background.js → chrome.tabs.create / chrome.tabs.update
 */

const BACKEND_URL = "https://focused-solace-production-6839.up.railway.app";
const SUPABASE_URL = "https://dvynmzeyttwlmvunicqz.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImR2eW5temV5dHR3bG12dW5pY3F6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQ2NTIwMTIsImV4cCI6MjA5MDIyODAxMn0.9j_f-2f1VswxWfqiuXy4bPnUi1qLk9nAeTDlodUBUZw";
const NOTES_URL = `${BACKEND_URL}/notes-app`;

let currentUser = null;
let pendingImages = [];
let sessionId = `browser-${Date.now()}`;

// ══════════════════════════════════════════════════════════════════════════
// AUTH
// ══════════════════════════════════════════════════════════════════════════

async function supabaseAuth(endpoint, body) {
  const resp = await fetch(`${SUPABASE_URL}/auth/v1/${endpoint}`, {
    method: "POST",
    headers: { "Content-Type": "application/json", "apikey": SUPABASE_ANON_KEY },
    body: JSON.stringify(body),
  });
  return resp.json();
}

async function syncSessionToBackend(session) {
  if (!session?.access_token) return;
  try {
    await fetch(`${BACKEND_URL}/auth/set-session`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        access_token: session.access_token,
        refresh_token: session.refresh_token,
        user_id: session.user?.id || "",
        email: session.user?.email || "",
      }),
    });
  } catch (e) {
    console.warn("tsifl: sync to backend failed:", e);
  }
}

async function restoreFromBackend() {
  try {
    const resp = await fetch(`${BACKEND_URL}/auth/get-session`);
    const data = await resp.json();
    if (!data.session || !data.session.refresh_token) return false;

    const result = await supabaseAuth("token?grant_type=refresh_token", {
      refresh_token: data.session.refresh_token,
    });

    if (result.access_token && result.user) {
      const email = result.user.email || data.session.email;
      chrome.storage.local.set({ tsifl_session: result, tsifl_email: email });
      syncSessionToBackend(result);
      showChat({ id: result.user.id || data.session.user_id, email });
      return true;
    }
  } catch (e) {
    console.warn("tsifl: restore from backend failed:", e);
  }
  return false;
}

async function checkAuth() {
  // Step 1: Try backend session FIRST — always has the latest token
  // (if user logged in via another add-in, backend has the freshest refresh_token)
  const restored = await restoreFromBackend();
  if (restored) return;

  // Step 2: Fall back to chrome.storage.local
  let stored = null;
  try {
    stored = await new Promise((resolve) => {
      chrome.storage.local.get("tsifl_session", (data) => resolve(data.tsifl_session || null));
    });
  } catch (e) {
    console.warn("tsifl: chrome.storage.local.get failed:", e);
  }

  if (stored && stored.refresh_token) {
    try {
      const result = await supabaseAuth("token?grant_type=refresh_token", {
        refresh_token: stored.refresh_token,
      });
      if (result.access_token && result.user) {
        const email = result.user.email || stored.user?.email;
        chrome.storage.local.set({ tsifl_session: result, tsifl_email: email });
        syncSessionToBackend(result);
        showChat({ id: result.user.id || stored.user?.id, email });
        return;
      }
    } catch (e) {
      console.warn("tsifl: local session refresh failed:", e);
    }
  }

  // Step 3: If we have a stored access_token that might still be valid, try using it directly
  // (Supabase access tokens are valid for 1 hour even if refresh fails)
  if (stored && stored.access_token) {
    try {
      // Decode JWT to check expiry
      const payload = JSON.parse(atob(stored.access_token.split(".")[1]));
      const expiresAt = payload.exp * 1000;
      if (Date.now() < expiresAt) {
        // Token still valid — use it
        const email = stored.user?.email || stored.email || "";
        const uid = stored.user?.id || stored.user_id || "";
        if (email) {
          showChat({ id: uid, email });
          return;
        }
      }
    } catch (e) {
      console.warn("tsifl: JWT decode failed:", e);
    }
  }

  // All failed — show login
  // Pre-fill email if we remember it from a previous session
  try {
    const savedEmail = await new Promise((resolve) => {
      chrome.storage.local.get("tsifl_email", (data) => resolve(data.tsifl_email || ""));
    });
    if (savedEmail) {
      document.getElementById("tsifl-auth-email").value = savedEmail;
      document.getElementById("tsifl-auth-password").focus();
      document.getElementById("tsifl-auth-error").textContent = "Session expired. Enter your password to continue.";
      document.getElementById("tsifl-auth-error").style.color = "#64748B";
    }
  } catch (e) {}
  showLogin();
}

// Proactive token refresh every 45 minutes — keeps session alive
// (matches Excel add-in behavior so tokens don't expire during use)
setInterval(async () => {
  if (!currentUser) return;
  try {
    const stored = await new Promise((resolve) => {
      chrome.storage.local.get("tsifl_session", (data) => resolve(data.tsifl_session || null));
    });
    if (stored && stored.refresh_token) {
      const result = await supabaseAuth("token?grant_type=refresh_token", {
        refresh_token: stored.refresh_token,
      });
      if (result.access_token && result.user) {
        chrome.storage.local.set({ tsifl_session: result, tsifl_email: result.user.email });
        syncSessionToBackend(result);
      }
    }
  } catch (e) {
    console.warn("tsifl: proactive refresh failed:", e);
  }
}, 45 * 60 * 1000);

async function handleSignIn() {
  const email = document.getElementById("tsifl-auth-email").value.trim();
  const password = document.getElementById("tsifl-auth-password").value;
  const errEl = document.getElementById("tsifl-auth-error");
  errEl.textContent = "";
  errEl.style.color = "#DC2626";
  // Validation (Improvement 9)
  if (!email || !email.includes("@")) { errEl.textContent = "Enter a valid email address."; return; }
  if (!password) { errEl.textContent = "Enter your password."; return; }

  try {
    const result = await supabaseAuth("token?grant_type=password", { email, password });
    if (result.error || result.error_description) {
      errEl.textContent = result.error_description || result.error || "Sign in failed";
      return;
    }
    chrome.storage.local.set({ tsifl_session: result, tsifl_email: result.user.email });
    await syncSessionToBackend(result);
    showChat({ id: result.user.id, email: result.user.email });
  } catch (e) {
    errEl.textContent = "Network error — check your connection.";
  }
}

async function handleSignUp() {
  const email = document.getElementById("tsifl-auth-email").value.trim();
  const password = document.getElementById("tsifl-auth-password").value;
  const errEl = document.getElementById("tsifl-auth-error");
  errEl.textContent = "";
  errEl.style.color = "#DC2626";
  if (!email || !email.includes("@")) { errEl.textContent = "Enter a valid email address."; return; }
  if (password.length < 6) { errEl.textContent = "Password must be at least 6 characters."; return; }

  try {
    const result = await supabaseAuth("signup", { email, password });
    if (result.error || result.error_description) {
      errEl.textContent = result.error_description || result.error || "Sign up failed";
      return;
    }
    errEl.style.color = "#16A34A";
    errEl.textContent = "Check your email to confirm, then sign in.";
  } catch (e) {
    errEl.textContent = "Network error — check your connection.";
  }
}

// Forgot password (Improvement 5)
async function handleForgotPassword() {
  const email = document.getElementById("tsifl-auth-email").value.trim();
  const errEl = document.getElementById("tsifl-auth-error");
  if (!email || !email.includes("@")) { errEl.style.color = "#DC2626"; errEl.textContent = "Enter your email first."; return; }
  try {
    const resp = await fetch(`${SUPABASE_URL}/auth/v1/recover`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "apikey": SUPABASE_ANON_KEY },
      body: JSON.stringify({ email }),
    });
    if (resp.ok) { errEl.style.color = "#16A34A"; errEl.textContent = "Check your email for a reset link."; }
    else { errEl.style.color = "#DC2626"; errEl.textContent = "Could not send reset email."; }
  } catch (e) { errEl.style.color = "#DC2626"; errEl.textContent = "Network error."; }
}

function showLogin() {
  document.getElementById("tsifl-login").style.display = "flex";
  document.getElementById("tsifl-chat-area").style.display = "none";
}

function showChat(user) {
  currentUser = user;
  document.getElementById("tsifl-login").style.display = "none";
  document.getElementById("tsifl-chat-area").style.display = "flex";
  // User display with avatar initial (Improvement 8)
  const initial = (user.email || "?")[0].toUpperCase();
  document.getElementById("tsifl-user-bar").innerHTML =
    `<span style="display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;border-radius:50%;background:#0D5EAF;color:white;font-size:10px;font-weight:700;margin-right:4px;">${initial}</span>${user.email || "Signed in"}`;
  // Update tab context display (Improvement 73)
  updateTabContext();
}

// ══════════════════════════════════════════════════════════════════════════
// CONTEXT CAPTURE
// ══════════════════════════════════════════════════════════════════════════

function getContext() {
  return new Promise((resolve) => {
    const timeout = setTimeout(() => resolve({ app: "browser" }), 5000);
    try {
      chrome.runtime.sendMessage({ action: "get_context" }, (response) => {
        clearTimeout(timeout);
        if (chrome.runtime.lastError) {
          resolve({ app: "browser" });
        } else {
          resolve(response?.context || { app: "browser" });
        }
      });
    } catch (e) {
      clearTimeout(timeout);
      resolve({ app: "browser" });
    }
  });
}

function getPageText() {
  return new Promise((resolve) => {
    const timeout = setTimeout(() => resolve(""), 5000);
    try {
      chrome.runtime.sendMessage({ action: "get_page_text" }, (response) => {
        clearTimeout(timeout);
        if (chrome.runtime.lastError) {
          resolve("");
        } else {
          resolve(response?.text || "");
        }
      });
    } catch (e) {
      clearTimeout(timeout);
      resolve("");
    }
  });
}

// ══════════════════════════════════════════════════════════════════════════
// IMAGES
// ══════════════════════════════════════════════════════════════════════════

function addFile(file) {
  const reader = new FileReader();
  reader.onload = () => {
    pendingImages.push({
      media_type: file.type || (file.name && file.name.match(/\.(png|jpg|jpeg|gif|webp)$/i) ? "image/png" : "application/octet-stream"),
      data: reader.result.split(",")[1],
      preview: file.type.startsWith("image/") ? reader.result : null,
      file_name: file.name || "",
    });
    updateImagePreview();
  };
  reader.readAsDataURL(file);
}

function addImage(file) { addFile(file); }

function updateImagePreview() {
  const bar = document.getElementById("tsifl-image-preview-bar");
  if (!pendingImages.length) { bar.style.display = "none"; bar.innerHTML = ""; return; }
  bar.style.display = "flex";
  bar.innerHTML = pendingImages.map((img, i) => {
    if (img.preview && img.media_type.startsWith("image/")) {
      return `<div class="tsifl-image-preview-item">
        <img src="${img.preview}"/>
        <button class="tsifl-remove-img" data-i="${i}">\u00d7</button>
      </div>`;
    } else {
      const ext = img.file_name ? img.file_name.split(".").pop().toUpperCase() : "FILE";
      return `<div class="tsifl-image-preview-item">
        <div style="width:40px;height:40px;display:flex;align-items:center;justify-content:center;background:#F1F5F9;border-radius:4px;border:1px solid #E2E8F0;font-size:9px;font-weight:700;color:#0D5EAF;">${ext}</div>
        <button class="tsifl-remove-img" data-i="${i}">\u00d7</button>
      </div>`;
    }
  }).join("");
  bar.querySelectorAll(".tsifl-remove-img").forEach(b =>
    b.onclick = () => { pendingImages.splice(+b.dataset.i, 1); updateImagePreview(); }
  );
}

// ══════════════════════════════════════════════════════════════════════════
// CHAT
// ══════════════════════════════════════════════════════════════════════════

function isSummarizationRequest(msg) {
  const lower = msg.toLowerCase();
  const triggers = [
    "summarize", "summary", "summarise", "main points", "key points",
    "key takeaways", "tldr", "tl;dr", "what does this page say",
    "what is this page about", "what is this article about",
    "explain this page", "explain this article", "break down this article",
    "overview of this", "give me the gist", "what's this about", "digest this",
  ];
  return triggers.some(t => lower.includes(t));
}

async function handleSubmit() {
  const input = document.getElementById("tsifl-user-input");
  const msg = input.value.trim();
  if (!msg && !pendingImages.length) return;
  if (!currentUser) { appendMessage("assistant", "Please sign in first."); return; }

  input.value = "";
  setSubmitEnabled(false);
  setStatus("Thinking...");

  appendMessage("user", msg, pendingImages.length);

  const images = pendingImages.map(i => ({ media_type: i.media_type, data: i.data }));
  pendingImages = [];
  updateImagePreview();

  try {
    const context = await getContext();

    // Update context display
    updateContextDisplay(context);
    updateContextActions(context);

    // If summarization request, capture full page text
    if (isSummarizationRequest(msg)) {
      setStatus("Reading page...");
      const pageText = await getPageText();
      if (pageText) context.full_page_text = pageText;
    }

    // Update UI badges
    const siteLabels = {
      gmail: "Gmail", google_sheets: "Sheets", google_docs: "Docs",
      google_slides: "Slides", browser: "Browser",
    };
    const siteName = siteLabels[context.app] || "Browser";
    const badge = document.getElementById("tsifl-site-badge");
    if (badge) badge.textContent = siteName;
    const userBar = document.getElementById("tsifl-user-bar");
    if (userBar && currentUser) userBar.textContent = `${currentUser.email} \u00b7 ${siteName}`;

    setStatus("Thinking...");
    showTypingIndicator();

    // Send to backend with timeout
    const controller = new AbortController();
    const fetchTimeout = setTimeout(() => controller.abort(), 90000);

    const resp = await fetch(`${BACKEND_URL}/chat/`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        user_id: currentUser.id,
        message: msg,
        context,
        session_id: sessionId,
        images,
      }),
      signal: controller.signal,
    });
    clearTimeout(fetchTimeout);

    if (!resp.ok) {
      hideTypingIndicator();
      const err = await resp.json().catch(() => ({ detail: resp.statusText }));
      appendMessage("assistant", `Error: ${err.detail || "Request failed"}`);
      setSubmitEnabled(true);
      setStatus("Ready");
      return;
    }

    const result = await resp.json();
    hideTypingIndicator();

    // Show the reply
    if (result.reply) {
      appendMessage("assistant", result.reply);
    }

    // Update tasks count
    if (result.tasks_remaining >= 0) {
      const tasksEl = document.getElementById("tsifl-tasks-remaining");
      if (tasksEl) tasksEl.textContent = `${result.tasks_remaining} tasks left`;
    }

    // Execute ALL actions — handle both single and multiple action formats
    const actions = [];
    if (result.actions && result.actions.length > 0) {
      actions.push(...result.actions);
    } else if (result.action && result.action.type) {
      actions.push(result.action);
    }

    for (const action of actions) {
      await executeAction(action);
    }

  } catch (e) {
    hideTypingIndicator();
    if (e.name === "AbortError") {
      appendMessage("assistant", "Request timed out. Please try again.");
    } else {
      appendMessage("assistant", `Error: ${e.message}`);
    }
  }

  setSubmitEnabled(true);
  setStatus("Ready");
}

// ══════════════════════════════════════════════════════════════════════════
// ACTION EXECUTION
// ══════════════════════════════════════════════════════════════════════════

async function executeAction(action) {
  if (!action || !action.type) return;
  const { type, payload } = action;
  if (!payload) return;

  // Gmail backend-proxied actions
  const gmailActions = ["draft_email", "send_email", "reply_email", "search_emails"];
  if (gmailActions.includes(type)) {
    try {
      const endpoint = type === "search_emails" ? "/gmail/search"
        : type === "draft_email" ? "/gmail/draft"
        : "/gmail/send";
      const resp = await fetch(`${BACKEND_URL}${endpoint}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const result = await resp.json();
      // action executed silently
    } catch (e) {
      console.error(`${type} failed:`, e);
    }
    return;
  }

  // Google Workspace actions — execute DIRECTLY on the tab via chrome.scripting.
  // This bypasses the background service worker entirely, avoiding the
  // "message port closed before response" bug with chrome.runtime.sendMessage.
  const workspaceActions = [
    "format_text", "insert_text", "insert_paragraph", "insert_table",
    "find_and_replace", "insert_page_break", "insert_header", "insert_footer",
    "write_cell", "write_range", "format_range", "add_sheet", "navigate_sheet",
    "sort_range", "add_chart", "clear_range", "set_number_format", "freeze_panes",
    "autofit", "create_slide", "add_text_box", "add_shape", "delete_slide",
    "set_slide_background", "apply_style",
  ];

  if (workspaceActions.includes(type)) {
    try {
      // Get active tab
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      if (!tab?.id) throw new Error("No active tab");

      // Execute workspace action INLINE on the tab.
      // All logic is self-contained — no dependency on content script.
      const results = await chrome.scripting.executeScript({
        target: { tabId: tab.id },
        world: "ISOLATED",
        func: async (actionType, actionPayload) => {
          try {
          const sleep = (ms) => new Promise(r => setTimeout(r, ms));

          function selectText(term) {
            window.getSelection()?.removeAllRanges();
            return window.find(term, false, false, true);
          }

          function clickToolbar(label) {
            for (const attr of ["aria-label", "data-tooltip"]) {
              for (const variant of [label, label.toLowerCase(), label.charAt(0).toUpperCase() + label.slice(1).toLowerCase()]) {
                const btn = document.querySelector(`[${attr}*="${variant}"]`);
                if (btn) { btn.click(); return true; }
              }
            }
            return false;
          }

          const site = (() => {
            const path = location.pathname;
            if (location.hostname !== "docs.google.com") return "other";
            if (path.startsWith("/document")) return "google_docs";
            if (path.startsWith("/spreadsheets")) return "google_sheets";
            if (path.startsWith("/presentation")) return "google_slides";
            return "other";
          })();

          // ── Google Docs ──────────────────────────────────────────
          // Google Docs uses a canvas-based renderer — window.find() doesn't work.
          // We use the native Find toolbar (Ctrl+F / Edit menu) to find & select text,
          // then click toolbar buttons to format.
          if (site === "google_docs") {

            // Helper: open the Find toolbar and search for a term.
            if (actionType === "format_text") {
              const term = actionPayload.range_description || "";
              if (!term) return { success: false, message: "No text specified" };

              // Check if Find & Replace dialog is ALREADY open (from a previous run)
              let findInput = null;
              let findBarOpened = false;

              // Scan ALL inputs — find one with aria-label/placeholder "Find" that's not a toolbar element
              const allInputs = document.querySelectorAll('input');
              for (const inp of allInputs) {
                const r = inp.getBoundingClientRect();
                if (r.width < 50 || r.height < 10) continue;
                const lbl = (inp.getAttribute('aria-label') || '').trim();
                const ph = (inp.placeholder || '').trim();
                const cls = (inp.className || '').toLowerCase();
                // Match Find dialog inputs
                if (lbl === 'Find' || lbl === 'Find in document' || ph === 'Find' || ph === 'Find in document') {
                  // Exclude toolbar elements
                  if (cls.includes('docs-title') || cls.includes('omnibox') || cls.includes('toolbar-combo')) continue;
                  findInput = inp;
                  findBarOpened = true;
                  break;
                }
              }

              // If no existing dialog, open one
              if (!findInput) {
              // Snapshot all current inputs/editables BEFORE opening Find (with visibility)
              const visibilityBefore = new Map();
              document.querySelectorAll('input, textarea, [contenteditable="true"]').forEach(el => {
                const r = el.getBoundingClientRect();
                visibilityBefore.set(el, r.width > 0 && r.height > 0);
              });

              // Step 1: Open Find bar via Ctrl+H (Find & Replace) keyboard shortcut
              // Synthetic keyboard — may not work on Google Docs (isTrusted=false) but worth trying
              document.body.dispatchEvent(new KeyboardEvent("keydown", {
                key: "h", code: "KeyH", keyCode: 72,
                ctrlKey: true, metaKey: false,
                bubbles: true, cancelable: true
              }));
              await sleep(800);

              // Check if a dialog opened
              let findInput = null;
              let findBarOpened = false;

              // Strategy A: Look for Find & Replace dialog by class
              const frDialog = document.querySelector('.docs-findandreplacedialog');
              if (frDialog) {
                const dInputs = frDialog.querySelectorAll('input');
                if (dInputs.length > 0) {
                  findInput = dInputs[0];
                  findBarOpened = true;
                }
              }

              // Strategy B: Look for NEW or newly-visible inputs
              if (!findInput) {
                const inputsAfter = document.querySelectorAll('input, textarea, [contenteditable="true"]');
                for (const el of inputsAfter) {
                  const rect = el.getBoundingClientRect();
                  if (rect.width <= 0 || rect.height <= 0) continue;
                  const wasBefore = visibilityBefore.get(el);
                  if (wasBefore === undefined || wasBefore === false) {
                    findInput = el;
                    findBarOpened = true;
                    break;
                  }
                }
              }

              // Strategy C: If Ctrl+H didn't work, try clicking Find button in toolbar
              if (!findInput) {
                // Try many selectors — Google Docs toolbar uses various attribute patterns
                const findBtn = document.querySelector('[data-tooltip="Find and replace"]') ||
                               document.querySelector('[aria-label="Find and replace"]') ||
                               document.querySelector('[data-tooltip="Find"]') ||
                               document.querySelector('[aria-label="Search the menus (Alt+/)"]')?.closest('.docs-icon-search') ||
                               document.querySelector('.docs-icon-search')?.closest('div[role="button"]') ||
                               document.querySelector('.docs-icon-img-container[style*="find"]')?.closest('div[role="button"]') ||
                               // The magnifying glass on the left toolbar
                               document.querySelector('[aria-label*="Search" i][role="button"]') ||
                               document.querySelector('div.goog-toolbar-button[aria-label*="ind"]');
                if (findBtn) {
                  // Record which inputs are currently visible
                  const visibleBefore = new Map();
                  document.querySelectorAll('input, textarea, [contenteditable="true"]').forEach(el => {
                    const r = el.getBoundingClientRect();
                    visibleBefore.set(el, r.width > 0 && r.height > 0);
                  });
                  findBtn.click();
                  await sleep(800);
                  // Find NEW inputs or inputs that became visible
                  const inputsAfterClick = document.querySelectorAll('input, textarea, [contenteditable="true"]');
                  for (const el of inputsAfterClick) {
                    const rect = el.getBoundingClientRect();
                    if (rect.width <= 0 || rect.height <= 0) continue;
                    const wasBefore = visibleBefore.get(el);
                    // Either brand new element OR was hidden and now visible
                    if (wasBefore === undefined || wasBefore === false) {
                      findInput = el;
                      findBarOpened = true;
                      break;
                    }
                  }
                }
              }

              // Strategy D: If still nothing, try Edit menu → Find and replace
              if (!findInput) {
                const menuBtnsAll = document.querySelectorAll('.menu-button');
                let editMenuBtn = null;
                for (const btn of menuBtnsAll) {
                  if (btn.textContent.trim() === "Edit") { editMenuBtn = btn; break; }
                }
                if (editMenuBtn) {
                  const visBeforeMenu = new Map();
                  document.querySelectorAll('input, textarea, [contenteditable="true"]').forEach(el => {
                    const r = el.getBoundingClientRect();
                    visBeforeMenu.set(el, r.width > 0 && r.height > 0);
                  });
                  // Open Edit menu
                  editMenuBtn.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
                  await sleep(50);
                  editMenuBtn.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true }));
                  editMenuBtn.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
                  await sleep(600);

                  // Find the "Find and replace" menu item
                  const menuItems = document.querySelectorAll('.goog-menuitem');
                  let frItem = null;
                  for (const el of menuItems) {
                    if (el.textContent.trim().toLowerCase().includes("find and replace")) {
                      frItem = el;
                      break;
                    }
                  }

                  if (frItem) {
                    // Get the inner content element — Closure Library renders content inside
                    const target = frItem.querySelector('.goog-menuitem-content') || frItem;
                    const rect = target.getBoundingClientRect();
                    const cx = rect.left + rect.width / 2;
                    const cy = rect.top + rect.height / 2;
                    const evtBase = {
                      bubbles: true, cancelable: true, view: window,
                      clientX: cx, clientY: cy, screenX: cx, screenY: cy,
                      button: 0, buttons: 1
                    };

                    // Full realistic mouse event sequence with coordinates
                    target.dispatchEvent(new PointerEvent('pointerover', evtBase));
                    target.dispatchEvent(new MouseEvent('mouseover', evtBase));
                    target.dispatchEvent(new PointerEvent('pointerenter', { ...evtBase, bubbles: false }));
                    target.dispatchEvent(new MouseEvent('mouseenter', { ...evtBase, bubbles: false }));
                    await sleep(100);
                    target.dispatchEvent(new PointerEvent('pointerdown', evtBase));
                    target.dispatchEvent(new MouseEvent('mousedown', evtBase));
                    await sleep(50);
                    target.dispatchEvent(new PointerEvent('pointerup', evtBase));
                    target.dispatchEvent(new MouseEvent('mouseup', evtBase));
                    target.dispatchEvent(new MouseEvent('click', evtBase));
                    await sleep(800);

                    // If that didn't work, try clicking the outer element too
                    const frDialog2 = document.querySelector('.docs-findandreplacedialog');
                    if (!frDialog2) {
                      frItem.dispatchEvent(new PointerEvent('pointerdown', evtBase));
                      frItem.dispatchEvent(new MouseEvent('mousedown', evtBase));
                      await sleep(50);
                      frItem.dispatchEvent(new PointerEvent('pointerup', evtBase));
                      frItem.dispatchEvent(new MouseEvent('mouseup', evtBase));
                      frItem.dispatchEvent(new MouseEvent('click', evtBase));
                      await sleep(800);
                    }
                  }

                  // Detect new/newly-visible inputs
                  const inputsAfterMenu = document.querySelectorAll('input, textarea, [contenteditable="true"]');
                  for (const el of inputsAfterMenu) {
                    const rect = el.getBoundingClientRect();
                    if (rect.width <= 0 || rect.height <= 0) continue;
                    const wasBefore = visBeforeMenu.get(el);
                    if (wasBefore === undefined || wasBefore === false) {
                      findInput = el;
                      findBarOpened = true;
                      break;
                    }
                  }
                }
              }

              // Strategy E: Last resort — scan ALL visible inputs for anything find-related
              if (!findInput) {
                const allInputs = Array.from(document.querySelectorAll('input, textarea, [contenteditable="true"]'));
                const diagInfo = allInputs.filter(i => {
                  const r = i.getBoundingClientRect();
                  return r.width > 0 && r.height > 0;
                }).map(i => {
                  const tag = i.tagName;
                  const cls = (i.className || '').substring(0, 30);
                  const lbl = i.getAttribute('aria-label') || i.type || '';
                  const par = (i.parentElement?.className || '').substring(0, 30);
                  return `${tag}.${cls}|${lbl}|parent:${par}`;
                });
                return { success: false, message: `Could not open Find bar. Visible inputs: ${diagInfo.slice(0, 8).join(" /// ")}` };
              }
              } // end: if (!findInput) — already-open dialog check

              // Final safety check
              if (!findInput) {
                // Dump ALL visible inputs for debugging
                const debugInputs = Array.from(document.querySelectorAll('input')).filter(i => {
                  const r = i.getBoundingClientRect();
                  return r.width > 20 && r.height > 5;
                }).map(i => {
                  return `aria="${i.getAttribute('aria-label')}" ph="${i.placeholder}" cls="${(i.className||'').substring(0,25)}"`;
                });
                return { success: false, message: `findInput null after all strategies. Inputs: ${debugInputs.slice(0,6).join(' | ')}` };
              }

              // Step 3: Fill search text and click Next to find & select it
              findInput.focus();
              const nativeSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")?.set;
              if (nativeSetter) nativeSetter.call(findInput, term);
              else findInput.value = term;
              findInput.dispatchEvent(new Event("input", { bubbles: true }));
              findInput.dispatchEvent(new Event("change", { bubbles: true }));
              await sleep(500);

              // Click "Next" / "Find" button in the Find & Replace dialog
              // (Enter key doesn't work because isTrusted=false)
              let nextClicked = false;
              const nextBtn = document.querySelector('[aria-label="Next"]') ||
                             document.querySelector('[aria-label="Find next"]') ||
                             document.querySelector('.docs-findandreplacedialog button[name="next"]') ||
                             document.querySelector('.docs-findandreplacedialog [data-tooltip*="Next"]');
              if (nextBtn) {
                nextBtn.click();
                nextClicked = true;
                await sleep(500);
              }
              // Fallback: look for any button in the dialog with "next" or arrow text
              if (!nextClicked) {
                const dialogBtns = document.querySelectorAll('.docs-findandreplacedialog button, [role="dialog"] button');
                for (const btn of dialogBtns) {
                  const lbl = (btn.getAttribute('aria-label') || btn.textContent || '').toLowerCase();
                  if (lbl.includes('next') || lbl.includes('find') || lbl === '▼' || lbl === '↓') {
                    btn.click();
                    nextClicked = true;
                    await sleep(500);
                    break;
                  }
                }
              }
              // Last fallback: try Enter anyway
              if (!nextClicked) {
                findInput.dispatchEvent(new KeyboardEvent("keydown", {
                  key: "Enter", code: "Enter", keyCode: 13, bubbles: true, cancelable: true
                }));
                await sleep(500);
              }

              // Step 7: Apply formatting WHILE dialog is still open
              // (Google Docs deselects text when dialog closes)
              await sleep(200);
              let formatted = false;
              const formatLog = [];
              formatLog.push("nextBtn:" + (nextClicked ? "clicked" : "NOT FOUND"));

              if (actionPayload.bold) {
                const ok = clickToolbar("Bold");
                formatLog.push("bold:" + ok);
                formatted = formatted || ok;
                await sleep(150);
              }
              if (actionPayload.italic) {
                const ok = clickToolbar("Italic");
                formatLog.push("italic:" + ok);
                formatted = formatted || ok;
                await sleep(150);
              }
              if (actionPayload.underline) {
                const ok = clickToolbar("Underline");
                formatLog.push("underline:" + ok);
                formatted = formatted || ok;
                await sleep(150);
              }

              if (actionPayload.highlight_color) {
                const colorName = actionPayload.highlight_color.toLowerCase();
                const hlBtn = document.querySelector('[aria-label*="Highlight color"]') ||
                             document.querySelector('[data-tooltip*="Highlight color"]') ||
                             document.querySelector('[aria-label*="highlight" i]');
                formatLog.push("hlBtn:" + (hlBtn ? hlBtn.getAttribute("aria-label") : "NOT FOUND"));

                if (hlBtn) {
                  // Snapshot elements before clicking to detect the color popup
                  const elsBefore = new Set(document.querySelectorAll('*'));

                  // In MD3 Google Docs, the highlight button may have a dropdown arrow
                  // as a sibling or child element. Try multiple approaches to open the picker.
                  const arrow = hlBtn.querySelector('.goog-toolbar-menu-button-dropdown') ||
                               hlBtn.querySelector('[class*="dropdown"]') ||
                               hlBtn.querySelector('svg')?.parentElement;

                  // Approach 1: Click the dropdown arrow with coordinates
                  const clickTarget = arrow || hlBtn;
                  const ctRect = clickTarget.getBoundingClientRect();
                  // Click near the RIGHT edge of the button (where dropdown arrows are)
                  const cx = arrow ? (ctRect.left + ctRect.width / 2) : (ctRect.right - 5);
                  const cy = ctRect.top + ctRect.height / 2;
                  const evtOpts = {
                    bubbles: true, cancelable: true, view: window,
                    clientX: cx, clientY: cy, screenX: cx, screenY: cy,
                    button: 0, buttons: 1
                  };
                  clickTarget.dispatchEvent(new PointerEvent('pointerdown', evtOpts));
                  clickTarget.dispatchEvent(new MouseEvent('mousedown', evtOpts));
                  await sleep(50);
                  clickTarget.dispatchEvent(new PointerEvent('pointerup', evtOpts));
                  clickTarget.dispatchEvent(new MouseEvent('mouseup', evtOpts));
                  clickTarget.dispatchEvent(new MouseEvent('click', evtOpts));
                  await sleep(700);

                  // Find NEW elements that appeared (the color popup)
                  let colorPopupEls = [];
                  const allNow = document.querySelectorAll('[style*="background"], [data-color], [role="option"], [role="listbox"] *, [class*="color"] [style]');
                  for (const el of allNow) {
                    if (!elsBefore.has(el)) colorPopupEls.push(el);
                  }
                  formatLog.push("newPopupEls:" + colorPopupEls.length);

                  // Also scan ALL elements with background-color style (popup might reuse existing nodes)
                  if (colorPopupEls.length === 0) {
                    // Look for any popup/overlay that appeared
                    const popups = document.querySelectorAll('[role="listbox"], [role="menu"], [class*="popup"], [class*="picker"], [class*="palette"]');
                    for (const popup of popups) {
                      const r = popup.getBoundingClientRect();
                      if (r.width > 0 && r.height > 0) {
                        colorPopupEls = Array.from(popup.querySelectorAll('*'));
                        formatLog.push("popup found:" + popup.className.substring(0, 30));
                        break;
                      }
                    }
                  }

                  // Color hex map for matching
                  const hexMap = {
                    yellow: ['#ffff00', '#fff200', '#ffd600', '#ffff00', 'rgb(255, 255, 0)', 'rgb(255, 242, 0)'],
                    green: ['#00ff00', '#00c853', '#00e676', 'rgb(0, 255, 0)'],
                    cyan: ['#00ffff', '#00bcd4', 'rgb(0, 255, 255)'],
                    blue: ['#0000ff', '#2962ff', 'rgb(0, 0, 255)'],
                    red: ['#ff0000', '#d50000', 'rgb(255, 0, 0)'],
                    orange: ['#ff9900', '#ff6d00', '#ff9900', 'rgb(255, 153, 0)', 'rgb(255, 109, 0)'],
                    pink: ['#ff00ff', '#f50057', '#e91e63', 'rgb(255, 0, 255)'],
                    magenta: ['#ff00ff', 'rgb(255, 0, 255)'],
                  };
                  const targetHexes = hexMap[colorName] || [colorName];

                  let clicked = false;

                  // Pass 1: Check new popup elements by background-color
                  for (const el of colorPopupEls) {
                    const bg = window.getComputedStyle(el).backgroundColor;
                    if (!bg || bg === 'rgba(0, 0, 0, 0)' || bg === 'transparent') continue;
                    const bgLower = bg.toLowerCase();
                    const matches = targetHexes.some(h => bgLower.includes(h)) || bgLower.includes(colorName);
                    if (matches) {
                      el.click(); clicked = true;
                      formatLog.push("popup color clicked:" + bg);
                      break;
                    }
                  }

                  // Pass 2: Check by aria-label, title, data attributes
                  if (!clicked) {
                    const labeled = document.querySelectorAll('[data-color], [aria-label*="color" i], [title]');
                    for (const el of labeled) {
                      const lbl = (el.getAttribute("aria-label") || el.getAttribute("title") || "").toLowerCase();
                      const dc = (el.getAttribute("data-color") || "").toLowerCase();
                      if (lbl.includes(colorName) || dc.includes(colorName)) {
                        el.click(); clicked = true;
                        formatLog.push("label color clicked:" + (lbl || dc).substring(0, 30));
                        break;
                      }
                      for (const hex of targetHexes) {
                        if (dc.includes(hex)) {
                          el.click(); clicked = true;
                          formatLog.push("hex color clicked:" + dc);
                          break;
                        }
                      }
                      if (clicked) break;
                    }
                  }

                  // Pass 3: Scan ALL visible small elements with bg color (brute force color picker detection)
                  if (!clicked) {
                    const allEls = document.querySelectorAll('div, span, td, button');
                    const colorEls = [];
                    for (const el of allEls) {
                      const r = el.getBoundingClientRect();
                      // Color cells are typically small squares (10-30px)
                      if (r.width < 8 || r.width > 50 || r.height < 8 || r.height > 50) continue;
                      if (Math.abs(r.width - r.height) > 10) continue; // roughly square
                      const bg = window.getComputedStyle(el).backgroundColor;
                      if (!bg || bg === 'rgba(0, 0, 0, 0)' || bg === 'transparent' || bg === 'rgb(255, 255, 255)') continue;
                      colorEls.push({ el, bg });
                    }
                    formatLog.push("squareColorEls:" + colorEls.length);

                    for (const { el, bg } of colorEls) {
                      const bgLower = bg.toLowerCase();
                      const matches = targetHexes.some(h => bgLower.includes(h)) || bgLower.includes(colorName);
                      if (matches) {
                        el.click(); clicked = true;
                        formatLog.push("square color clicked:" + bg);
                        break;
                      }
                    }

                    // If no exact match but we have color cells, pick yellow as default
                    if (!clicked && colorEls.length > 0) {
                      // Try to find ANY yellow-ish cell as default highlight
                      for (const { el, bg } of colorEls) {
                        if (bg.includes('255, 255, 0') || bg.includes('255, 242, 0') || bg.includes('255, 214, 0')) {
                          el.click(); clicked = true;
                          formatLog.push("default yellow:" + bg);
                          break;
                        }
                      }
                    }
                  }

                  if (!clicked) formatLog.push("no color applied");
                  formatted = clicked;
                }
              }

              if (actionPayload.font_color) {
                const colorBtn = document.querySelector('[aria-label*="Text color"]') ||
                                document.querySelector('[data-tooltip*="Text color"]');
                if (colorBtn) {
                  const arrow = colorBtn.querySelector('[class*="dropdown"]') || colorBtn;
                  arrow.click();
                  await sleep(500);
                  const cells = document.querySelectorAll('[data-color], [aria-label], [title]');
                  for (const cell of cells) {
                    const lbl = (cell.getAttribute("aria-label") || cell.getAttribute("title") || "").toLowerCase();
                    if (lbl.includes(actionPayload.font_color.toLowerCase())) { cell.click(); formatted = true; break; }
                  }
                }
              }

              // Step 8: Close any open find bar or dialog
              const closeBtnFR = document.querySelector('.docs-findandreplacedialog [aria-label="Close"]') ||
                                document.querySelector('.docs-findandreplacedialog-close') ||
                                document.querySelector('.docs-findinput-container [aria-label="Close"]') ||
                                document.querySelector('[aria-label="Close"][class*="Gm3Wiz"]') ||
                                // MD3: close button near the Find dialog
                                (() => {
                                  const closeBtns = document.querySelectorAll('[aria-label="Close"], button[aria-label="Close"]');
                                  for (const b of closeBtns) {
                                    const r = b.getBoundingClientRect();
                                    if (r.width > 0 && r.height > 0 && r.top < 200) return b;
                                  }
                                  return null;
                                })();
              if (closeBtnFR) {
                closeBtnFR.click();
                await sleep(200);
              }
              // Also close the find bar if it's open (press Escape on the find input)
              if (findInput) {
                findInput.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
                await sleep(200);
              }

              return formatted
                ? { success: true, message: `Formatted "${term}" [${formatLog.join(", ")}]` }
                : { success: false, message: `Found "${term}" but toolbar buttons not found. Log: ${formatLog.join(", ")}` };
            }

            if (actionType === "find_and_replace") {
              // Click Edit menu → Find and replace
              const menuBtns2 = document.querySelectorAll('.menu-button.goog-control');
              let editMenu2 = null;
              for (const btn of menuBtns2) { if (btn.textContent.trim() === "Edit") { editMenu2 = btn; break; } }
              if (!editMenu2) return { success: false, message: "Could not find Edit menu" };
              editMenu2.click();
              await sleep(400);
              const menuItems = document.querySelectorAll('[role="menuitem"], .goog-menuitem');
              const frItem = Array.from(menuItems).find(el => el.textContent.toLowerCase().includes("find and replace"));
              if (!frItem) return { success: false, message: "Could not find Find & Replace menu item" };
              frItem.click();
              await sleep(500);
              const inputs = Array.from(document.querySelectorAll('input[type="text"]')).filter(i => i.offsetParent !== null);
              if (inputs.length >= 2) {
                inputs[0].focus(); inputs[0].value = actionPayload.find_text || "";
                inputs[0].dispatchEvent(new Event("input", { bubbles: true }));
                inputs[1].focus(); inputs[1].value = actionPayload.replace_text || "";
                inputs[1].dispatchEvent(new Event("input", { bubbles: true }));
                await sleep(200);
                const raBtn = Array.from(document.querySelectorAll("button")).find(b => b.textContent.toLowerCase().includes("replace all"));
                if (raBtn) { raBtn.click(); await sleep(300); }
                const closeBtn = document.querySelector('[aria-label="Close"]');
                if (closeBtn) closeBtn.click();
                return { success: true, message: "Find and replace completed" };
              }
              return { success: false, message: "Could not find dialog inputs" };
            }

            if (actionType === "insert_text") {
              document.execCommand("insertText", false, actionPayload.text || "");
              return { success: true, message: "Text inserted" };
            }

            return { success: false, message: `Google Docs action "${actionType}" not yet supported.` };
          }

          // ── Google Sheets ────────────────────────────────────────
          if (site === "google_sheets") {
            if (actionType === "write_cell") {
              const cell = actionPayload.cell || "A1";
              const value = actionPayload.formula || actionPayload.value || "";
              const nameBox = document.querySelector("#t-name-box input") ||
                             document.querySelector('[aria-label="Name Box"]');
              if (nameBox) {
                nameBox.click(); await sleep(100);
                nameBox.focus(); nameBox.value = cell;
                nameBox.dispatchEvent(new Event("input", { bubbles: true }));
                nameBox.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", keyCode: 13, bubbles: true }));
                await sleep(300);
                const editor = document.querySelector(".cell-input") || document.activeElement;
                if (editor) {
                  editor.focus();
                  document.execCommand("selectAll", false, null);
                  document.execCommand("insertText", false, value.toString());
                  editor.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", keyCode: 13, bubbles: true }));
                  return { success: true, message: `Wrote "${value}" to ${cell}` };
                }
              }
              return { success: false, message: "Could not find Name Box" };
            }

            if (actionType === "navigate_sheet") {
              const tabs = document.querySelectorAll(".docs-sheet-tab");
              for (const tab of tabs) {
                const name = tab.querySelector(".docs-sheet-tab-name");
                if (name && name.textContent.trim() === (actionPayload.sheet || "")) {
                  tab.click();
                  return { success: true, message: `Navigated to ${actionPayload.sheet}` };
                }
              }
              return { success: false, message: `Sheet "${actionPayload.sheet}" not found` };
            }

            return { success: false, message: `Google Sheets action "${actionType}" not yet supported.` };
          }

          return { success: false, message: `Not on a supported Google Workspace page (detected: ${site})` };
          } catch (err) {
            return { success: false, message: "SCRIPT ERROR: " + (err.message || err) + " | stack: " + (err.stack || "").substring(0, 200) };
          }
        },
        args: [type, payload],
      });

      const result = results?.[0]?.result;
      console.log("workspace result:", result);
      // ALWAYS show result during development so we can see what happened
      if (result?.message) {
        appendMessage("assistant", result.message);
      } else if (!result) {
        appendMessage("assistant", "No result from workspace action.");
      }
    } catch (e) {
      console.error(`workspace ${type} failed:`, e);
      appendMessage("assistant", `Could not execute ${type}: ${e.message || e}`);
    }
    return;
  }

  // Browser/DOM actions — route through background.js
  const browserActions = [
    "open_url", "open_url_current_tab", "search_web",
    "navigate_back", "navigate_forward",
    "scroll_to", "click_element", "fill_input", "extract_text",
  ];

  if (browserActions.includes(type)) {
    try {
      const result = await sendToBackground("execute_browser_action", type, payload);
      // action executed silently
    } catch (e) {
      console.error(`${type} failed:`, e);
    }
    return;
  }

  // Launch local app
  if (type === "launch_app") {
    try {
      const resp = await fetch(`${BACKEND_URL}/launch-app`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const result = await resp.json();
      // action executed silently
    } catch (e) {
      console.error("launch_app failed:", e);
    }
    return;
  }

  // Open notes
  if (type === "open_notes") {
    try {
      await sendToBackground("execute_browser_action", "open_url", { url: NOTES_URL });
      // action executed silently
    } catch (e) {
      console.error("open_notes failed:", e);
    }
    return;
  }

  // Create a note
  if (type === "create_note") {
    try {
      const resp = await fetch(`${BACKEND_URL}/notes/`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          user_id: currentUser.id,
          title: payload.title || "Untitled Note",
          content: payload.content || "",
          folder: payload.folder || "General",
        }),
      });
      const note = await resp.json();
      // action executed silently
    } catch (e) {
      console.error("create_note failed:", e);
    }
    return;
  }

  // Fallback — execute silently
  console.log(`Action executed: ${type}`);
}

/**
 * Send message to background.js and wait for response.
 * This is the critical path for ALL browser actions.
 */
function sendToBackground(action, type, payload) {
  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      reject(new Error("Background script did not respond within 10s"));
    }, 10000);

    try {
      chrome.runtime.sendMessage(
        { action, type, payload },
        (response) => {
          clearTimeout(timeout);
          if (chrome.runtime.lastError) {
            reject(new Error(chrome.runtime.lastError.message));
          } else {
            resolve(response || { success: true, message: "Done" });
          }
        }
      );
    } catch (e) {
      clearTimeout(timeout);
      reject(e);
    }
  });
}

// ══════════════════════════════════════════════════════════════════════════
// UI HELPERS
// ══════════════════════════════════════════════════════════════════════════

function appendMessage(role, text, imageCount) {
  const h = document.getElementById("tsifl-chat-history");
  const d = document.createElement("div");
  d.className = `tsifl-msg ${role}`;

  if (role === "assistant" && text) {
    d.innerHTML = renderMarkdown(text);
  } else {
    d.textContent = text || "";
  }

  if (imageCount > 0) {
    const b = document.createElement("div");
    b.className = "tsifl-image-badge";
    b.textContent = `${imageCount} image${imageCount > 1 ? "s" : ""} attached`;
    d.appendChild(b);
  }
  h.appendChild(d);
  h.scrollTop = h.scrollHeight;
}

function renderMarkdown(text) {
  let html = text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  html = html.replace(/```(\w*)\n?([\s\S]*?)```/g, function(_, lang, code) {
    const id = "cb_" + Math.random().toString(36).slice(2, 8);
    return '<pre id="' + id + '"><button class="copy-btn" onclick="navigator.clipboard.writeText(document.getElementById(\'' + id + '\').textContent.replace(/^Copy\\n?/,\'\'))">Copy</button><code>' + code.trim() + '</code></pre>';
  });
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

const _thinkingMessages = [
  'Reading your question...',
  'Thinking about this...',
  'Processing that thought...',
  'Crafting the perfect response...',
  'Hold on, almost there...',
  'Consulting my inner Shakespeare...',
  'This one requires some brainpower...',
  'Typing faster than you can read...',
  'If I had hands I would be rubbing them together...',
  'Brew yourself a coffee, this is gonna be good...',
  'Give me a sec, genius takes time...',
  'Loading witty response...',
  'Warming up the neural networks...',
  'My therapist said I should take on more challenges...',
  'Running calculations at the speed of thought...',
  'Almost there, just dotting my i\'s...',
  'The answer is forming... like a beautiful butterfly...',
  'Consulting the ancient scrolls of knowledge...',
  'McKinsey would charge you $50k for this...',
];

let _thinkingInterval = null;

function showTypingIndicator() {
  const history = document.getElementById("tsifl-chat-history");
  const div = document.createElement("div");
  div.id = "typing-indicator";
  div.className = "thinking-bubble";
  div.innerHTML = '<div class="thinking-orb"></div><div class="thinking-text"></div>';
  history.appendChild(div);
  history.scrollTop = history.scrollHeight;

  const textEl = div.querySelector(".thinking-text");
  let idx = 0;
  textEl.textContent = _thinkingMessages[0];
  _thinkingInterval = setInterval(() => {
    idx = (idx + 1) % _thinkingMessages.length;
    textEl.style.opacity = "0";
    setTimeout(() => {
      textEl.textContent = _thinkingMessages[idx];
      textEl.style.opacity = "1";
    }, 200);
  }, 3000);
}

function hideTypingIndicator() {
  if (_thinkingInterval) { clearInterval(_thinkingInterval); _thinkingInterval = null; }
  const el = document.getElementById("typing-indicator");
  if (el) el.remove();
}

function setStatus(text) {
  const bar = document.getElementById("tsifl-status-bar");
  if (bar) {
    bar.textContent = text;
    bar.className = text.includes("Thinking") || text.includes("Reading") ? "thinking" : "";
  }
}

function setSubmitEnabled(enabled) {
  const btn = document.getElementById("tsifl-submit-btn");
  if (btn) btn.disabled = !enabled;
}

// ══════════════════════════════════════════════════════════════════════════
// CONTEXT DISPLAY & CONTEXT ACTIONS
// ══════════════════════════════════════════════════════════════════════════

function updateContextDisplay(context) {
  const el = document.getElementById("tsifl-context-display");
  if (!el) return;

  let summary = "";

  if (context.app === "gmail") {
    const count = context.message_count || (context.messages?.length) || 0;
    const sender = context.messages?.[0]?.sender_email || context.messages?.[0]?.sender_name || "";
    const subject = context.thread_subject || "";
    if (count > 0 && sender) {
      summary = `Reading: Gmail \u2014 ${count} message${count !== 1 ? "s" : ""} in thread from ${sender}`;
    } else if (subject) {
      summary = `Reading: Gmail \u2014 "${subject}"`;
    } else if (context.is_composing) {
      summary = "Reading: Gmail \u2014 Composing new email";
    } else {
      summary = "Reading: Gmail";
    }
  } else if (context.product) {
    const p = context.product;
    const details = [p.name];
    if (p.price) details.push(p.price);
    if (p.rating) details.push(`${p.rating} stars`);
    const source = (p.source || "").charAt(0).toUpperCase() + (p.source || "").slice(1);
    summary = `Analyzing: ${source || "Product"} \u2014 ${details.join(", ")}`;
  } else if (context.app === "google_sheets") {
    summary = `Analyzing: Google Sheets \u2014 ${context.sheet_title || "Spreadsheet"}`;
  } else if (context.app === "google_docs") {
    summary = `Editing: Google Docs \u2014 ${context.doc_title || "Document"}`;
  } else if (context.app === "google_slides") {
    summary = `Viewing: Google Slides \u2014 ${context.slide_count || 0} slides`;
  } else {
    // Generic browser
    const pageTitle = context.title || "";
    const host = (() => {
      try { return new URL(context.url || "").hostname.replace("www.", ""); } catch (e) { return ""; }
    })();
    const pageType = context.page_type || "";
    const typeLabels = {
      wikipedia: "Wikipedia", article: "Article", financial: "Financial",
      github: "GitHub", stackoverflow: "Stack Overflow", search_results: "Search Results",
      linkedin_profile: "LinkedIn Profile", linkedin_job: "LinkedIn Job",
    };
    const label = typeLabels[pageType] || (host ? host.split(".")[0].charAt(0).toUpperCase() + host.split(".")[0].slice(1) : "Page");
    const titleSnippet = pageTitle.length > 50 ? pageTitle.slice(0, 50) + "..." : pageTitle;
    summary = `Browsing: ${label}${titleSnippet ? " \u2014 " + titleSnippet : ""}`;
  }

  if (summary) {
    el.textContent = summary;
    el.style.display = "block";
  } else {
    el.style.display = "none";
  }

  // Also update tab context
  const tabCtx = document.getElementById("tsifl-tab-context");
  const titleEl = document.getElementById("tsifl-tab-title");
  if (tabCtx && titleEl) {
    tabCtx.style.display = "block";
    titleEl.textContent = summary || "On: " + (context.title || "").slice(0, 60);
  }
}

function updateContextActions(context) {
  const container = document.getElementById("tsifl-context-actions");
  if (!container) return;

  const actions = [];

  if (context.app === "gmail") {
    actions.push({ label: "Draft reply", prompt: "Draft a professional reply to this email thread" });
    actions.push({ label: "Summarize thread", prompt: "Summarize this email thread with key points and decisions" });
    actions.push({ label: "Extract action items", prompt: "Extract all action items and deadlines from this email thread" });
    actions.push({ label: "Find emails", prompt: "Help me find specific emails — what should I search for?" });
  }

  if (context.app === "google_sheets") {
    actions.push({ label: "Analyze data", prompt: "Analyze the data in this spreadsheet and provide key insights" });
    actions.push({ label: "Create chart", prompt: "Suggest the best chart type for this data and how to create it" });
    actions.push({ label: "Write formula", prompt: "Help me write a formula for this spreadsheet" });
  }

  if (context.tables && context.tables.length > 0) {
    actions.push({ label: "Extract table data", prompt: "Extract and organize the table data on this page into a clean format" });
  }

  if (context.app === "google_docs") {
    actions.push({ label: "Improve writing", prompt: "Review and suggest improvements to this document" });
    actions.push({ label: "Summarize doc", prompt: "Summarize this document with key points" });
  }

  if (context.product) {
    actions.push({ label: "Compare prices", prompt: "Help me compare this product with alternatives and find the best deal" });
    actions.push({ label: "Product summary", prompt: "Give me a quick summary of this product: pros, cons, and value" });
  }

  if (actions.length === 0) {
    container.style.display = "none";
    return;
  }

  container.style.display = "flex";
  container.innerHTML = actions.map(a =>
    `<button class="tsifl-quick-btn tsifl-context-btn" data-prompt="${a.prompt.replace(/"/g, '&quot;')}">${a.label}</button>`
  ).join("");

  // Wire up click handlers
  container.querySelectorAll(".tsifl-context-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const prompt = btn.getAttribute("data-prompt");
      if (prompt) {
        document.getElementById("tsifl-user-input").value = prompt;
        handleSubmit();
      }
    });
  });
}

// ══════════════════════════════════════════════════════════════════════════
// EVENT WIRING
// ══════════════════════════════════════════════════════════════════════════

document.getElementById("tsifl-login-btn").onclick = handleSignIn;
document.getElementById("tsifl-signup-btn").onclick = handleSignUp;
document.getElementById("tsifl-submit-btn").onclick = handleSubmit;
document.getElementById("tsifl-attach-btn").onclick = () => document.getElementById("tsifl-image-input").click();

// Password toggle (Improvement 9)
const togglePw = document.getElementById("tsifl-toggle-pw");
if (togglePw) {
  togglePw.onclick = () => {
    const pw = document.getElementById("tsifl-auth-password");
    pw.type = pw.type === "password" ? "text" : "password";
  };
}

// Forgot password (Improvement 5)
const forgotPwBtn = document.getElementById("tsifl-forgot-pw");
if (forgotPwBtn) forgotPwBtn.onclick = handleForgotPassword;

document.getElementById("tsifl-auth-password").addEventListener("keydown", (e) => {
  if (e.key === "Enter") handleSignIn();
});

document.getElementById("tsifl-user-input").addEventListener("keydown", (e) => {
  if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleSubmit(); }
});

document.getElementById("tsifl-image-input").onchange = (e) => {
  for (const file of e.target.files) addFile(file);
  e.target.value = "";
};

// Paste images & files
document.getElementById("tsifl-user-input").addEventListener("paste", (e) => {
  for (const item of (e.clipboardData || {}).items || []) {
    if (item.type.startsWith("image/") || item.kind === "file") {
      const file = item.getAsFile();
      if (file) addFile(file);
    }
  }
});

// Sign out on double-click user bar
document.getElementById("tsifl-user-bar").addEventListener("dblclick", () => {
  chrome.storage.local.remove("tsifl_session");
  fetch(`${BACKEND_URL}/auth/clear-session`, { method: "POST" }).catch(() => {});
  currentUser = null;
  document.getElementById("tsifl-chat-history").innerHTML = "";
  showLogin();
});

// ══════════════════════════════════════════════════════════════════════════
// TAB CONTEXT (Improvement 73)
// ══════════════════════════════════════════════════════════════════════════

async function updateTabContext() {
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (tab) {
      const ctxEl = document.getElementById("tsifl-tab-context");
      const titleEl = document.getElementById("tsifl-tab-title");
      const urlEl = document.getElementById("tsifl-tab-url");
      if (ctxEl && titleEl && urlEl) {
        ctxEl.style.display = "block";
        titleEl.textContent = "On: " + (tab.title || "").slice(0, 60);
        urlEl.textContent = (tab.url || "").slice(0, 80);
      }
    }
    // Pre-load context actions on panel open
    const context = await getContext();
    updateContextDisplay(context);
    updateContextActions(context);
  } catch (e) { /* silent */ }
}

// ══════════════════════════════════════════════════════════════════════════
// INIT
// ══════════════════════════════════════════════════════════════════════════
checkAuth();
