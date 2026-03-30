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
  // Step 1: Check chrome.storage.local for saved session
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

  // Step 2: Try to restore from backend (logged in via another add-in)
  const restored = await restoreFromBackend();
  if (!restored) {
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
}

async function handleSignIn() {
  const email = document.getElementById("tsifl-auth-email").value.trim();
  const password = document.getElementById("tsifl-auth-password").value;
  const errEl = document.getElementById("tsifl-auth-error");
  errEl.textContent = "";
  errEl.style.color = "#DC2626";
  if (!email || !password) { errEl.textContent = "Enter email and password."; return; }

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
  if (!email || !password) { errEl.textContent = "Enter email and password."; return; }
  if (password.length < 6) { errEl.textContent = "Password must be 6+ characters."; return; }

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

function showLogin() {
  document.getElementById("tsifl-login").style.display = "flex";
  document.getElementById("tsifl-chat-area").style.display = "none";
}

function showChat(user) {
  currentUser = user;
  document.getElementById("tsifl-login").style.display = "none";
  document.getElementById("tsifl-chat-area").style.display = "flex";
  document.getElementById("tsifl-user-bar").textContent = user.email || "Signed in";
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

function addImage(file) {
  const reader = new FileReader();
  reader.onload = () => {
    pendingImages.push({
      media_type: file.type || "image/png",
      data: reader.result.split(",")[1],
      preview: reader.result,
    });
    updateImagePreview();
  };
  reader.readAsDataURL(file);
}

function updateImagePreview() {
  const bar = document.getElementById("tsifl-image-preview-bar");
  if (!pendingImages.length) { bar.style.display = "none"; bar.innerHTML = ""; return; }
  bar.style.display = "flex";
  bar.innerHTML = pendingImages.map((img, i) =>
    `<div class="tsifl-image-preview-item">
      <img src="${img.preview}"/>
      <button class="tsifl-remove-img" data-i="${i}">\u00d7</button>
    </div>`
  ).join("");
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
      const err = await resp.json().catch(() => ({ detail: resp.statusText }));
      appendMessage("assistant", `Error: ${err.detail || "Request failed"}`);
      setSubmitEnabled(true);
      setStatus("Ready");
      return;
    }

    const result = await resp.json();

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
      appendMessage("action", `${type}: ${result.status || "Done"}`);
    } catch (e) {
      appendMessage("action", `${type}: Failed \u2014 ${e.message}`);
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
      appendMessage("action", `${type}: ${result?.message || "Done"}`);
    } catch (e) {
      appendMessage("action", `${type}: Failed \u2014 ${e.message}`);
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
      appendMessage("action", `launch_app: ${result.message || result.status || "Requested"}`);
    } catch (e) {
      appendMessage("action", `launch_app: ${e.message}`);
    }
    return;
  }

  // Open notes
  if (type === "open_notes") {
    try {
      await sendToBackground("execute_browser_action", "open_url", { url: NOTES_URL });
      appendMessage("action", "Opened Notes");
    } catch (e) {
      appendMessage("action", `open_notes: ${e.message}`);
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
      appendMessage("action", `Note created: "${note.title || "Untitled"}"`);
    } catch (e) {
      appendMessage("action", `create_note: ${e.message}`);
    }
    return;
  }

  // Fallback — show action as text
  appendMessage("action", `${type}: Done`);
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
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>")
    .replace(/\*(.+?)\*/g, "<em>$1</em>")
    .replace(/`([^`]+)`/g, '<code style="background:#F1F5F9;padding:1px 4px;border-radius:3px;font-size:12px;">$1</code>')
    .replace(/\n/g, "<br>");
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
// EVENT WIRING
// ══════════════════════════════════════════════════════════════════════════

document.getElementById("tsifl-login-btn").onclick = handleSignIn;
document.getElementById("tsifl-signup-btn").onclick = handleSignUp;
document.getElementById("tsifl-submit-btn").onclick = handleSubmit;
document.getElementById("tsifl-attach-btn").onclick = () => document.getElementById("tsifl-image-input").click();

// Notes button
const notesBtn = document.getElementById("tsifl-notes-btn");
if (notesBtn) {
  notesBtn.onclick = () => {
    sendToBackground("execute_browser_action", "open_url", { url: NOTES_URL }).catch(() => {
      window.open(NOTES_URL, "_blank");
    });
  };
}

document.getElementById("tsifl-auth-password").addEventListener("keydown", (e) => {
  if (e.key === "Enter") handleSignIn();
});

document.getElementById("tsifl-user-input").addEventListener("keydown", (e) => {
  if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleSubmit(); }
});

document.getElementById("tsifl-image-input").onchange = (e) => {
  for (const file of e.target.files) addImage(file);
  e.target.value = "";
};

// Paste images
document.getElementById("tsifl-user-input").addEventListener("paste", (e) => {
  for (const item of (e.clipboardData || {}).items || []) {
    if (item.type.startsWith("image/")) {
      const file = item.getAsFile();
      if (file) addImage(file);
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
// INIT
// ══════════════════════════════════════════════════════════════════════════
checkAuth();
