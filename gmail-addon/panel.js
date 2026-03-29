/**
 * tsifl — Side Panel Script
 * Handles auth, chat, images, and action execution.
 * All browser actions are routed through background.js for reliability.
 */

const BACKEND_URL = "https://focused-solace-production-6839.up.railway.app";
const SUPABASE_URL = "https://dvynmzeyttwlmvunicqz.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImR2eW5temV5dHR3bG12dW5pY3F6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQ2NTIwMTIsImV4cCI6MjA5MDIyODAxMn0.9j_f-2f1VswxWfqiuXy4bPnUi1qLk9nAeTDlodUBUZw";

let currentUser = null;
let pendingImages = [];
let sessionId = `browser-${Date.now()}`;

// ── Auth ─────────────────────────────────────────────────────────────────

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
  } catch (e) {}
}

async function restoreFromBackend() {
  try {
    const resp = await fetch(`${BACKEND_URL}/auth/get-session`);
    const data = await resp.json();
    if (!data.session?.access_token) return false;
    const result = await supabaseAuth("token?grant_type=refresh_token", {
      refresh_token: data.session.refresh_token,
    });
    if (result.access_token) {
      chrome.storage.local.set({ tsifl_session: result });
      syncSessionToBackend(result);
      showChat({ id: result.user?.id || data.session.user_id, email: result.user?.email || data.session.email });
      return true;
    }
  } catch (e) {}
  return false;
}

async function checkAuth() {
  const stored = await new Promise(r => chrome.storage.local.get("tsifl_session", d => r(d.tsifl_session)));
  if (stored) {
    try {
      const result = await supabaseAuth("token?grant_type=refresh_token", {
        refresh_token: stored.refresh_token,
      });
      if (result.access_token) {
        chrome.storage.local.set({ tsifl_session: result });
        syncSessionToBackend(result);
        showChat({ id: result.user?.id || stored.user?.id, email: result.user?.email || stored.user?.email });
        return;
      }
    } catch (e) {}
  }
  const restored = await restoreFromBackend();
  if (!restored) showLogin();
}

async function handleSignIn() {
  const email = document.getElementById("tsifl-auth-email").value.trim();
  const password = document.getElementById("tsifl-auth-password").value;
  const errEl = document.getElementById("tsifl-auth-error");
  errEl.textContent = "";
  errEl.style.color = "#DC2626";
  if (!email || !password) { errEl.textContent = "Enter email and password."; return; }

  const result = await supabaseAuth("token?grant_type=password", { email, password });
  if (result.error || result.error_description) {
    errEl.textContent = result.error_description || "Sign in failed";
    return;
  }
  chrome.storage.local.set({ tsifl_session: result });
  syncSessionToBackend(result);
  showChat({ id: result.user.id, email: result.user.email });
}

async function handleSignUp() {
  const email = document.getElementById("tsifl-auth-email").value.trim();
  const password = document.getElementById("tsifl-auth-password").value;
  const errEl = document.getElementById("tsifl-auth-error");
  errEl.textContent = "";
  errEl.style.color = "#DC2626";
  if (!email || !password) { errEl.textContent = "Enter email and password."; return; }
  if (password.length < 6) { errEl.textContent = "Password must be 6+ characters."; return; }

  const result = await supabaseAuth("signup", { email, password });
  if (result.error || result.error_description) {
    errEl.textContent = result.error_description || "Sign up failed";
    return;
  }
  errEl.style.color = "#16A34A";
  errEl.textContent = "Check your email to confirm, then sign in.";
}

function showLogin() {
  document.getElementById("tsifl-login").style.display = "flex";
  document.getElementById("tsifl-chat-area").style.display = "none";
}

function showChat(user) {
  currentUser = user;
  document.getElementById("tsifl-login").style.display = "none";
  document.getElementById("tsifl-chat-area").style.display = "flex";
  document.getElementById("tsifl-user-bar").textContent = user.email;
}

// ── Context ──────────────────────────────────────────────────────────────

function getContext() {
  return new Promise((resolve) => {
    const timeout = setTimeout(() => resolve({ app: "browser" }), 5000);
    chrome.runtime.sendMessage({ action: "get_context" }, (response) => {
      clearTimeout(timeout);
      if (chrome.runtime.lastError) {
        resolve({ app: "browser" });
      } else {
        resolve(response?.context || { app: "browser" });
      }
    });
  });
}

// Get full page text for summarization
function getPageText() {
  return new Promise((resolve) => {
    const timeout = setTimeout(() => resolve(""), 5000);
    chrome.runtime.sendMessage({ action: "get_page_text" }, (response) => {
      clearTimeout(timeout);
      if (chrome.runtime.lastError) {
        resolve("");
      } else {
        resolve(response?.text || "");
      }
    });
  });
}

// ── Images ───────────────────────────────────────────────────────────────

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

// ── Chat ─────────────────────────────────────────────────────────────────

// Detect if user is asking for a page summary
function isSummarizationRequest(msg) {
  const lower = msg.toLowerCase();
  const triggers = [
    "summarize", "summary", "summarise", "main points", "key points", "key takeaways",
    "tldr", "tl;dr", "what does this page say", "what is this page about", "what is this article about",
    "explain this page", "explain this article", "break down this article", "overview of this",
    "give me the gist", "what's this about", "digest this",
  ];
  return triggers.some(t => lower.includes(t));
}

async function handleSubmit() {
  const input = document.getElementById("tsifl-user-input");
  const msg = input.value.trim();
  if (!msg && !pendingImages.length) return;

  input.value = "";
  setSubmitEnabled(false);
  setStatus("Thinking...");

  appendMessage("user", msg, pendingImages.length);

  const images = pendingImages.map(i => ({ media_type: i.media_type, data: i.data }));
  pendingImages = [];
  updateImagePreview();

  try {
    const context = await getContext();

    // If this looks like a summarization request, capture full page text
    if (isSummarizationRequest(msg)) {
      const pageText = await getPageText();
      if (pageText) {
        context.full_page_text = pageText;
      }
    }

    // Update site badge
    const badge = document.getElementById("tsifl-site-badge");
    const siteLabels = { gmail: "Gmail", google_sheets: "Sheets", google_docs: "Docs", google_slides: "Slides", browser: "Browser" };
    if (badge) badge.textContent = siteLabels[context.app] || "Browser";
    const userBar = document.getElementById("tsifl-user-bar");
    if (userBar && currentUser) userBar.textContent = `${currentUser.email} \u00b7 ${siteLabels[context.app] || "Browser"}`;

    const chatBody = JSON.stringify({
      user_id: currentUser.id,
      message: msg,
      context,
      session_id: sessionId,
      images,
    });

    // Fetch with timeout and single retry
    let resp;
    for (let attempt = 0; attempt < 2; attempt++) {
      try {
        const controller = new AbortController();
        const timeout = setTimeout(() => controller.abort(), 90000); // 90s timeout
        resp = await fetch(`${BACKEND_URL}/chat/`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: chatBody,
          signal: controller.signal,
        });
        clearTimeout(timeout);
        if (resp.ok) break;
      } catch (e) {
        if (attempt === 0 && e.name === "AbortError") {
          setStatus("Retrying...");
          continue;
        }
        throw e;
      }
    }

    if (!resp || !resp.ok) {
      const err = resp ? await resp.json().catch(() => ({ detail: resp.statusText })) : { detail: "Request timed out" };
      appendMessage("assistant", `Error: ${err.detail || "Request failed"}`);
      setSubmitEnabled(true);
      setStatus("Ready");
      return;
    }

    const result = await resp.json();
    appendMessage("assistant", result.reply);

    if (result.tasks_remaining >= 0) {
      document.getElementById("tsifl-tasks-remaining").textContent = `${result.tasks_remaining} tasks left`;
    }

    // Execute actions
    const actions = result.actions?.length ? result.actions : (result.action?.type ? [result.action] : []);
    for (const action of actions) {
      await executeAction(action);
    }
  } catch (e) {
    appendMessage("assistant", `Error: ${e.message}`);
  }

  setSubmitEnabled(true);
  setStatus("Ready");
}

// ── Action Execution ─────────────────────────────────────────────────────

async function executeAction(action) {
  const { type, payload } = action;
  if (!type || !payload) return;

  // Gmail backend-proxied actions
  if (["draft_email", "send_email", "reply_email", "search_emails"].includes(type)) {
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

  // All browser/DOM actions — route through background service worker for reliability
  const browserActions = [
    "open_url", "open_url_current_tab", "search_web",
    "navigate_back", "navigate_forward",
    "scroll_to", "click_element", "fill_input", "extract_text",
  ];

  if (browserActions.includes(type)) {
    try {
      const result = await new Promise((resolve, reject) => {
        chrome.runtime.sendMessage(
          { action: "execute_browser_action", type, payload },
          (response) => {
            if (chrome.runtime.lastError) {
              reject(new Error(chrome.runtime.lastError.message));
            } else {
              resolve(response);
            }
          }
        );
      });
      appendMessage("action", `${type}: ${result?.message || "Done"}`);
    } catch (e) {
      appendMessage("action", `${type}: Failed \u2014 ${e.message}`);
    }
    return;
  }

  // launch_app — request backend to open a local app
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

  // Fallback for unknown action types
  appendMessage("action", `${type}: Done`);
}

// ── UI Helpers ───────────────────────────────────────────────────────────

function appendMessage(role, text, imageCount) {
  const h = document.getElementById("tsifl-chat-history");
  const d = document.createElement("div");
  d.className = `tsifl-msg ${role}`;

  // Render markdown-like formatting for assistant messages
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
  // Basic markdown rendering — bold, italic, code, links
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
  bar.textContent = text;
  bar.className = text.includes("Thinking") || text.includes("Retrying") ? "thinking" : "";
}

function setSubmitEnabled(enabled) {
  document.getElementById("tsifl-submit-btn").disabled = !enabled;
}

// ── Event Wiring ─────────────────────────────────────────────────────────

document.getElementById("tsifl-login-btn").onclick = handleSignIn;
document.getElementById("tsifl-signup-btn").onclick = handleSignUp;
document.getElementById("tsifl-submit-btn").onclick = handleSubmit;
document.getElementById("tsifl-attach-btn").onclick = () => document.getElementById("tsifl-image-input").click();

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

// ── Init ─────────────────────────────────────────────────────────────────
checkAuth();
