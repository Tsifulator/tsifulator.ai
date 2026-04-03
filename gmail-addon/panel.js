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
