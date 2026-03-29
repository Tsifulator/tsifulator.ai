/**
 * tsifl — Gmail Chrome Extension Content Script
 * Injects a tsifl sidebar into Gmail with chat UI, auth, and email context reading.
 */

(function () {
  "use strict";

  const BACKEND_URL = "https://focused-solace-production-6839.up.railway.app";
  const SUPABASE_URL = "https://dvynmzeyttwlmvunicqz.supabase.co";
  const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImR2eW5temV5dHR3bG12dW5pY3F6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQ2NTIwMTIsImV4cCI6MjA5MDIyODAxMn0.9j_f-2f1VswxWfqiuXy4bPnUi1qLk9nAeTDlodUBUZw";

  let currentUser = null;
  let pendingImages = [];
  let sidebarVisible = false;

  // ── Sidebar HTML ────────────────────────────────────────────────────────────

  function createSidebar() {
    if (document.getElementById("tsifl-sidebar")) return;

    const sidebar = document.createElement("div");
    sidebar.id = "tsifl-sidebar";
    sidebar.className = "hidden";
    sidebar.innerHTML = `
      <div id="tsifl-header">
        <span style="font-weight:700;color:#0D5EAF;font-size:15px;">tsifl</span>
        <div id="tsifl-header-right">
          <span id="tsifl-tasks-remaining"></span>
          <button id="tsifl-close-btn" title="Close">×</button>
        </div>
      </div>

      <!-- Login -->
      <div id="tsifl-login">
        <div style="font-weight:700;color:#0D5EAF;font-size:18px;margin-bottom:4px;">tsifl</div>
        <div id="tsifl-login-tagline">Agentic Sandbox for Financial Analysts</div>
        <div id="tsifl-login-form">
          <input type="email" id="tsifl-auth-email" placeholder="Email" />
          <input type="password" id="tsifl-auth-password" placeholder="Password" />
          <button id="tsifl-login-btn">Sign In</button>
          <button id="tsifl-signup-btn">Create Account</button>
          <div id="tsifl-auth-error"></div>
        </div>
      </div>

      <!-- Chat (hidden until login) -->
      <div id="tsifl-chat-area" style="display:none;flex:1;display:none;flex-direction:column;overflow:hidden;">
        <div id="tsifl-user-bar"></div>
        <div id="tsifl-chat-history"></div>
        <div id="tsifl-input-area">
          <div id="tsifl-image-preview-bar"></div>
          <textarea id="tsifl-user-input" placeholder="Draft emails, summarize threads, extract action items..." rows="3"></textarea>
          <div id="tsifl-input-actions">
            <input type="file" id="tsifl-image-input" accept="image/*" multiple style="display:none;" />
            <button id="tsifl-attach-btn" title="Attach image">+</button>
            <button id="tsifl-submit-btn">Send</button>
          </div>
        </div>
        <div id="tsifl-status-bar">Ready</div>
      </div>
    `;

    document.body.appendChild(sidebar);
    wireEvents();
    checkAuth();
  }

  // ── Auth (using Supabase REST API directly, no SDK needed) ────────────────

  async function supabaseAuth(endpoint, body) {
    const resp = await fetch(`${SUPABASE_URL}/auth/v1/${endpoint}`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "apikey": SUPABASE_ANON_KEY,
      },
      body: JSON.stringify(body),
    });
    return resp.json();
  }

  async function checkAuth() {
    const stored = localStorage.getItem("tsifl_gmail_session");
    if (stored) {
      try {
        const session = JSON.parse(stored);
        // Refresh token
        const result = await supabaseAuth("token?grant_type=refresh_token", {
          refresh_token: session.refresh_token,
        });
        if (result.access_token) {
          localStorage.setItem("tsifl_gmail_session", JSON.stringify(result));
          showChat({ id: result.user?.id || session.user?.id, email: result.user?.email || session.user?.email });
          return;
        }
      } catch (e) { /* continue to login */ }
    }
    showLogin();
  }

  async function handleSignIn() {
    const email = document.getElementById("tsifl-auth-email").value.trim();
    const password = document.getElementById("tsifl-auth-password").value;
    const errEl = document.getElementById("tsifl-auth-error");
    errEl.textContent = "";

    if (!email || !password) { errEl.textContent = "Enter email and password."; return; }

    const result = await supabaseAuth("token?grant_type=password", { email, password });
    if (result.error || result.error_description) {
      errEl.textContent = result.error_description || result.msg || "Sign in failed";
      return;
    }
    localStorage.setItem("tsifl_gmail_session", JSON.stringify(result));
    showChat({ id: result.user.id, email: result.user.email });
  }

  async function handleSignUp() {
    const email = document.getElementById("tsifl-auth-email").value.trim();
    const password = document.getElementById("tsifl-auth-password").value;
    const errEl = document.getElementById("tsifl-auth-error");
    errEl.textContent = "";

    if (!email || !password) { errEl.textContent = "Enter email and password."; return; }
    if (password.length < 6) { errEl.textContent = "Password must be 6+ characters."; return; }

    const result = await supabaseAuth("signup", { email, password });
    if (result.error || result.error_description) {
      errEl.textContent = result.error_description || result.msg || "Sign up failed";
      return;
    }
    errEl.style.color = "#16A34A";
    errEl.textContent = "Check your email to confirm, then sign in.";
  }

  function showLogin() {
    document.getElementById("tsifl-login").style.display = "flex";
    const chatArea = document.getElementById("tsifl-chat-area");
    chatArea.style.display = "none";
  }

  function showChat(user) {
    currentUser = user;
    document.getElementById("tsifl-login").style.display = "none";
    const chatArea = document.getElementById("tsifl-chat-area");
    chatArea.style.display = "flex";
    document.getElementById("tsifl-user-bar").textContent = user.email;
  }

  // ── Gmail Context Reading ─────────────────────────────────────────────────

  function getGmailContext() {
    const context = {
      app: "gmail",
      email: currentUser?.email || "",
      recent_emails: [],
      current_thread: null,
    };

    try {
      // Try to read the current open email
      const subjectEl = document.querySelector('h2[data-thread-perm-id]') ||
                        document.querySelector('.hP') ||
                        document.querySelector('[role="main"] h2');
      if (subjectEl) {
        const subject = subjectEl.textContent.trim();

        // Read message bodies in the thread
        const messageEls = document.querySelectorAll('.gs .ii.gt div[dir="ltr"], .gs .ii.gt div[data-message-id]');
        const messages = [];

        // Read sender info
        const senderEls = document.querySelectorAll('.gD, [email]');
        const senders = [];
        senderEls.forEach((el, i) => {
          if (i < 10) senders.push(el.getAttribute("email") || el.textContent.trim());
        });

        messageEls.forEach((el, i) => {
          if (i < 10) {
            messages.push({
              from: senders[i] || "",
              snippet: el.textContent.trim().substring(0, 300),
            });
          }
        });

        // Fallback: read the entire visible thread text
        if (messages.length === 0) {
          const threadBody = document.querySelector('[role="main"] .nH .adn');
          if (threadBody) {
            messages.push({
              from: senders[0] || "",
              snippet: threadBody.textContent.trim().substring(0, 500),
            });
          }
        }

        context.current_thread = { subject, messages };
      }

      // Read inbox snippets if no thread is open
      if (!context.current_thread) {
        const rows = document.querySelectorAll('tr.zA');
        rows.forEach((row, i) => {
          if (i >= 5) return;
          const from = row.querySelector('.yX .yW .bA4 span, .yX .yW span[email]');
          const subj = row.querySelector('.y6 span:first-child, .bog span');
          context.recent_emails.push({
            from: from ? from.textContent.trim() : "",
            subject: subj ? subj.textContent.trim() : "",
          });
        });
      }
    } catch (e) {
      console.error("tsifl: Gmail context error", e);
    }

    return context;
  }

  // ── Image Handling ────────────────────────────────────────────────────────

  function addImage(file) {
    const reader = new FileReader();
    reader.onload = () => {
      const base64 = reader.result.split(",")[1];
      pendingImages.push({
        media_type: file.type || "image/png",
        data: base64,
        preview: reader.result,
      });
      updateImagePreview();
    };
    reader.readAsDataURL(file);
  }

  function updateImagePreview() {
    const bar = document.getElementById("tsifl-image-preview-bar");
    if (pendingImages.length === 0) { bar.style.display = "none"; bar.innerHTML = ""; return; }
    bar.style.display = "flex";
    bar.innerHTML = pendingImages.map((img, i) => `
      <div class="tsifl-image-preview-item">
        <img src="${img.preview}" />
        <button class="tsifl-remove-img" data-idx="${i}">×</button>
      </div>
    `).join("");
    bar.querySelectorAll(".tsifl-remove-img").forEach(btn => {
      btn.onclick = () => { pendingImages.splice(+btn.dataset.idx, 1); updateImagePreview(); };
    });
  }

  // ── Chat Submit ───────────────────────────────────────────────────────────

  async function handleSubmit() {
    const input = document.getElementById("tsifl-user-input");
    const message = input.value.trim();
    if (!message && pendingImages.length === 0) return;

    input.value = "";
    setSubmitEnabled(false);
    setStatus("Thinking...");

    const imageCount = pendingImages.length;
    appendMessage("user", message, imageCount);

    const images = pendingImages.map(img => ({ media_type: img.media_type, data: img.data }));
    pendingImages = [];
    updateImagePreview();

    try {
      const context = getGmailContext();

      const resp = await fetch(`${BACKEND_URL}/chat/`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          user_id: currentUser.id,
          message,
          context,
          session_id: "gmail-" + Date.now(),
          images,
        }),
      });

      if (!resp.ok) {
        const err = await resp.json().catch(() => ({ detail: resp.statusText }));
        throw new Error(err.detail || "Request failed");
      }

      const result = await resp.json();

      if (result.tasks_remaining >= 0) {
        document.getElementById("tsifl-tasks-remaining").textContent = `${result.tasks_remaining} tasks left`;
      }

      appendMessage("assistant", result.reply);

      // Execute actions
      const actions = result.actions?.length ? result.actions : (result.action?.type ? [result.action] : []);
      if (actions.length > 0) {
        appendMessage("action", `Executing ${actions.length} action(s): ${actions.map(a => a.type).join(", ")}`);
        for (const action of actions) {
          await executeAction(action);
        }
      }

      setStatus("Ready");
    } catch (e) {
      appendMessage("assistant", `Error: ${e.message}`);
      setStatus("Error — try again");
    }

    setSubmitEnabled(true);
  }

  // ── Action Executor ───────────────────────────────────────────────────────

  async function executeAction(action) {
    const { type, payload } = action;
    if (!payload) return;

    switch (type) {
      case "draft_email":
      case "send_email": {
        const endpoint = type === "draft_email" ? "/gmail/draft" : "/gmail/send";
        try {
          const resp = await fetch(`${BACKEND_URL}${endpoint}`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              to: payload.to,
              subject: payload.subject,
              body: payload.body,
              reply_to_id: payload.reply_to_id || "",
            }),
          });
          const result = await resp.json();
          appendMessage("action", `Email ${type === "draft_email" ? "drafted" : "sent"}: ${result.status || result.id || "ok"}`);
        } catch (e) {
          appendMessage("action", `Failed to ${type}: ${e.message}`);
        }
        break;
      }

      case "reply_email": {
        try {
          const resp = await fetch(`${BACKEND_URL}/gmail/send`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              to: payload.to || "",
              subject: payload.subject || "",
              body: payload.body,
              reply_to_id: payload.thread_id || "",
            }),
          });
          const result = await resp.json();
          appendMessage("action", `Reply sent: ${result.status || "ok"}`);
        } catch (e) {
          appendMessage("action", `Failed to reply: ${e.message}`);
        }
        break;
      }

      case "search_emails": {
        try {
          const resp = await fetch(`${BACKEND_URL}/gmail/search`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ query: payload.query, max_results: 10 }),
          });
          const result = await resp.json();
          const emails = result.emails || [];
          if (emails.length > 0) {
            const summary = emails.map(e => `${e.from}: ${e.subject}`).join("\n");
            appendMessage("action", `Found ${emails.length} emails:\n${summary}`);
          } else {
            appendMessage("action", "No emails found matching query.");
          }
        } catch (e) {
          appendMessage("action", `Search failed: ${e.message}`);
        }
        break;
      }

      case "summarize_thread":
      case "extract_action_items":
        // These are handled by Claude's text response — the action just signals intent
        appendMessage("action", `${type}: See reply above.`);
        break;

      default:
        console.warn("tsifl: Unknown action type:", type);
    }
  }

  // ── UI Helpers ────────────────────────────────────────────────────────────

  function appendMessage(role, text, imageCount) {
    const history = document.getElementById("tsifl-chat-history");
    const div = document.createElement("div");
    div.className = `tsifl-msg ${role}`;
    div.textContent = text || "";

    if (imageCount && imageCount > 0) {
      const badge = document.createElement("div");
      badge.className = "tsifl-image-badge";
      badge.textContent = `${imageCount} image${imageCount > 1 ? "s" : ""} attached`;
      div.appendChild(badge);
    }

    history.appendChild(div);
    history.scrollTop = history.scrollHeight;
  }

  function setStatus(text) {
    const el = document.getElementById("tsifl-status-bar");
    if (el) el.textContent = text;
  }

  function setSubmitEnabled(enabled) {
    const btn = document.getElementById("tsifl-submit-btn");
    if (btn) btn.disabled = !enabled;
  }

  // ── Event Wiring ──────────────────────────────────────────────────────────

  function wireEvents() {
    document.getElementById("tsifl-close-btn").onclick = toggleSidebar;
    document.getElementById("tsifl-login-btn").onclick = handleSignIn;
    document.getElementById("tsifl-signup-btn").onclick = handleSignUp;
    document.getElementById("tsifl-submit-btn").onclick = handleSubmit;

    document.getElementById("tsifl-user-input").addEventListener("keydown", (e) => {
      if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleSubmit(); }
    });

    // Image handling
    const imageInput = document.getElementById("tsifl-image-input");
    document.getElementById("tsifl-attach-btn").onclick = () => imageInput.click();
    imageInput.onchange = (e) => {
      for (const file of e.target.files) addImage(file);
      imageInput.value = "";
    };

    // Drag & drop on input area
    const inputArea = document.getElementById("tsifl-input-area");
    inputArea.addEventListener("dragover", (e) => { e.preventDefault(); });
    inputArea.addEventListener("drop", (e) => {
      e.preventDefault();
      for (const file of e.dataTransfer.files) {
        if (file.type.startsWith("image/")) addImage(file);
      }
    });

    // Paste
    document.getElementById("tsifl-user-input").addEventListener("paste", (e) => {
      for (const item of (e.clipboardData || {}).items || []) {
        if (item.type.startsWith("image/")) addImage(item.getAsFile());
      }
    });
  }

  // ── Toggle Sidebar ────────────────────────────────────────────────────────

  function toggleSidebar() {
    const sidebar = document.getElementById("tsifl-sidebar");
    if (!sidebar) { createSidebar(); return; }

    sidebarVisible = !sidebarVisible;
    sidebar.classList.toggle("hidden", !sidebarVisible);
  }

  // ── Listen for extension icon clicks ──────────────────────────────────────

  chrome.runtime.onMessage.addListener((msg) => {
    if (msg.action === "toggle_sidebar") {
      if (!document.getElementById("tsifl-sidebar")) createSidebar();
      toggleSidebar();
    }
  });

  // ── Auto-create sidebar (hidden) on page load ────────────────────────────
  createSidebar();

})();
