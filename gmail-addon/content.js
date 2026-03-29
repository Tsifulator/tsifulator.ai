/**
 * tsifl — Browser-Wide Floating Sidebar
 * Works on ANY webpage with context-aware AI assistance.
 * Gmail: email context + email actions
 * Google Sheets/Docs/Slides: document context + editing actions
 * Any other page: page content/selection + general Q&A
 */

(function () {
  "use strict";

  const BACKEND_URL = "https://focused-solace-production-6839.up.railway.app";
  const SUPABASE_URL = "https://dvynmzeyttwlmvunicqz.supabase.co";
  const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImR2eW5temV5dHR3bG12dW5pY3F6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQ2NTIwMTIsImV4cCI6MjA5MDIyODAxMn0.9j_f-2f1VswxWfqiuXy4bPnUi1qLk9nAeTDlodUBUZw";

  let currentUser = null;
  let pendingImages = [];
  let sidebarVisible = false;
  let sidebarMinimized = false;

  // ── Site Detection ──────────────────────────────────────────────────────────

  function detectSite() {
    const url = window.location.href;
    if (url.includes("mail.google.com")) return "gmail";
    if (url.includes("docs.google.com/spreadsheets")) return "google_sheets";
    if (url.includes("docs.google.com/document")) return "google_docs";
    if (url.includes("docs.google.com/presentation")) return "google_slides";
    return "browser";
  }

  function getSiteLabel() {
    const site = detectSite();
    const labels = {
      gmail: "Gmail",
      google_sheets: "Google Sheets",
      google_docs: "Google Docs",
      google_slides: "Google Slides",
      browser: "Browser",
    };
    return labels[site] || "Browser";
  }

  function getPlaceholder() {
    const site = detectSite();
    const placeholders = {
      gmail: "Draft emails, summarize threads, extract action items...",
      google_sheets: "Analyze data, create formulas, format cells...",
      google_docs: "Draft text, format documents, summarize content...",
      google_slides: "Create slides, add content, design presentations...",
      browser: "Ask about this page, summarize content, extract data...",
    };
    return placeholders[site] || placeholders.browser;
  }

  // ── Sidebar HTML ────────────────────────────────────────────────────────────

  function createSidebar() {
    if (document.getElementById("tsifl-sidebar")) return;

    const sidebar = document.createElement("div");
    sidebar.id = "tsifl-sidebar";
    sidebar.className = "hidden";
    sidebar.innerHTML = `
      <div id="tsifl-resize-handle"></div>
      <div id="tsifl-header">
        <div id="tsifl-drag-handle">
          <span style="font-weight:700;color:#0D5EAF;font-size:15px;">tsifl</span>
          <span id="tsifl-site-badge">${getSiteLabel()}</span>
        </div>
        <div id="tsifl-header-right">
          <span id="tsifl-tasks-remaining"></span>
          <button id="tsifl-minimize-btn" title="Minimize">−</button>
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

      <!-- Chat -->
      <div id="tsifl-chat-area">
        <div id="tsifl-user-bar"></div>
        <div id="tsifl-chat-history"></div>
        <div id="tsifl-input-area">
          <div id="tsifl-image-preview-bar"></div>
          <textarea id="tsifl-user-input" placeholder="${getPlaceholder()}" rows="3"></textarea>
          <div id="tsifl-input-actions">
            <input type="file" id="tsifl-image-input" accept="image/*" multiple style="display:none;" />
            <button id="tsifl-attach-btn" title="Attach image">+</button>
            <button id="tsifl-submit-btn">Send</button>
          </div>
        </div>
        <div id="tsifl-status-bar">Ready</div>
      </div>
    `;

    // Minimized floating button
    const minBtn = document.createElement("div");
    minBtn.id = "tsifl-fab";
    minBtn.className = "hidden";
    minBtn.innerHTML = `<span style="font-weight:700;color:white;font-size:12px;">t</span>`;
    minBtn.onclick = () => {
      sidebarMinimized = false;
      minBtn.classList.add("hidden");
      sidebar.classList.remove("hidden");
      sidebarVisible = true;
    };

    document.body.appendChild(sidebar);
    document.body.appendChild(minBtn);
    wireEvents();
    initDrag();
    initResize();
    checkAuth();
  }

  // ── Auth (Supabase REST API) ────────────────────────────────────────────────

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
    const stored = localStorage.getItem("tsifl_session");
    if (stored) {
      try {
        const session = JSON.parse(stored);
        const result = await supabaseAuth("token?grant_type=refresh_token", {
          refresh_token: session.refresh_token,
        });
        if (result.access_token) {
          localStorage.setItem("tsifl_session", JSON.stringify(result));
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
    localStorage.setItem("tsifl_session", JSON.stringify(result));
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
    document.getElementById("tsifl-chat-area").style.display = "none";
  }

  function showChat(user) {
    currentUser = user;
    document.getElementById("tsifl-login").style.display = "none";
    document.getElementById("tsifl-chat-area").style.display = "flex";
    document.getElementById("tsifl-user-bar").textContent = `${user.email} · ${getSiteLabel()}`;
  }

  // ── Context Capture ─────────────────────────────────────────────────────────

  function getContext() {
    const site = detectSite();
    switch (site) {
      case "gmail": return getGmailContext();
      case "google_sheets": return getGoogleSheetsContext();
      case "google_docs": return getGoogleDocsContext();
      case "google_slides": return getGoogleSlidesContext();
      default: return getBrowserContext();
    }
  }

  function getGmailContext() {
    const context = { app: "gmail", email: currentUser?.email || "", recent_emails: [], current_thread: null };

    try {
      const subjectEl = document.querySelector('h2[data-thread-perm-id]') ||
                        document.querySelector('.hP') ||
                        document.querySelector('[role="main"] h2');
      if (subjectEl) {
        const subject = subjectEl.textContent.trim();
        const messageEls = document.querySelectorAll('.gs .ii.gt div[dir="ltr"], .gs .ii.gt div[data-message-id]');
        const senderEls = document.querySelectorAll('.gD, [email]');
        const senders = [];
        senderEls.forEach((el, i) => {
          if (i < 10) senders.push(el.getAttribute("email") || el.textContent.trim());
        });
        const messages = [];
        messageEls.forEach((el, i) => {
          if (i < 10) messages.push({ from: senders[i] || "", snippet: el.textContent.trim().substring(0, 300) });
        });
        if (messages.length === 0) {
          const threadBody = document.querySelector('[role="main"] .nH .adn');
          if (threadBody) messages.push({ from: senders[0] || "", snippet: threadBody.textContent.trim().substring(0, 500) });
        }
        context.current_thread = { subject, messages };
      }

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
    } catch (e) { console.error("tsifl: Gmail context error", e); }

    return context;
  }

  function getGoogleSheetsContext() {
    const context = { app: "google_sheets", url: window.location.href, title: "", selection: "", visible_data: [] };
    try {
      const titleEl = document.querySelector('.docs-title-input') || document.querySelector('[data-tooltip="Rename"]');
      context.title = titleEl ? titleEl.value || titleEl.textContent.trim() : document.title;

      // Read selected text
      const sel = window.getSelection();
      context.selection = sel ? sel.toString().substring(0, 500) : "";

      // Read sheet tab names
      const tabs = document.querySelectorAll('.docs-sheet-tab .docs-sheet-tab-name');
      context.sheet_tabs = [];
      tabs.forEach(t => context.sheet_tabs.push(t.textContent.trim()));

      // Read active cell info from the name box
      const nameBox = document.querySelector('#t-name-box .jfk-textinput, .waffle-name-box input');
      context.active_cell = nameBox ? nameBox.value : "";

      // Read formula bar
      const formulaBar = document.querySelector('.cell-input, #t-formula-bar-input .jfk-textinput');
      context.formula = formulaBar ? formulaBar.textContent || formulaBar.value || "" : "";

      // Read visible cell data (from DOM)
      const cells = document.querySelectorAll('.cell-input');
      context.visible_data = [];
    } catch (e) { console.error("tsifl: Sheets context error", e); }
    return context;
  }

  function getGoogleDocsContext() {
    const context = { app: "google_docs", url: window.location.href, title: "", selection: "", content: "" };
    try {
      const titleEl = document.querySelector('.docs-title-input');
      context.title = titleEl ? titleEl.value || titleEl.textContent.trim() : document.title;

      const sel = window.getSelection();
      context.selection = sel ? sel.toString().substring(0, 1000) : "";

      // Read document content from the editor
      const editor = document.querySelector('.kix-appview-editor');
      if (editor) {
        context.content = editor.textContent.substring(0, 3000);
      }
    } catch (e) { console.error("tsifl: Docs context error", e); }
    return context;
  }

  function getGoogleSlidesContext() {
    const context = { app: "google_slides", url: window.location.href, title: "", selection: "", slide_count: 0 };
    try {
      const titleEl = document.querySelector('.docs-title-input');
      context.title = titleEl ? titleEl.value || titleEl.textContent.trim() : document.title;

      const sel = window.getSelection();
      context.selection = sel ? sel.toString().substring(0, 500) : "";

      // Count slides from filmstrip
      const slides = document.querySelectorAll('.punch-filmstrip-thumbnail');
      context.slide_count = slides.length;

      // Read current slide text
      const currentSlide = document.querySelector('.punch-viewer-svgpage-svgcontainer');
      if (currentSlide) {
        const texts = currentSlide.querySelectorAll('text, tspan');
        context.current_slide_text = [];
        texts.forEach((t, i) => {
          if (i < 20 && t.textContent.trim()) context.current_slide_text.push(t.textContent.trim());
        });
      }
    } catch (e) { console.error("tsifl: Slides context error", e); }
    return context;
  }

  function getBrowserContext() {
    const context = { app: "browser", url: window.location.href, title: document.title, selection: "", page_text: "" };
    try {
      const sel = window.getSelection();
      context.selection = sel ? sel.toString().substring(0, 2000) : "";

      // Get main content
      const main = document.querySelector('main, article, [role="main"], .content, #content');
      if (main) {
        context.page_text = main.textContent.substring(0, 3000);
      } else {
        context.page_text = document.body.textContent.substring(0, 3000);
      }

      // Get meta description
      const meta = document.querySelector('meta[name="description"]');
      context.meta_description = meta ? meta.getAttribute("content") : "";
    } catch (e) { console.error("tsifl: Browser context error", e); }
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
    if (!bar) return;
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
      const context = getContext();

      const resp = await fetch(`${BACKEND_URL}/chat/`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          user_id: currentUser.id,
          message,
          context,
          session_id: `${detectSite()}-${Date.now()}`,
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
              to: payload.to, subject: payload.subject, body: payload.body,
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
            body: JSON.stringify({ to: payload.to || "", subject: payload.subject || "", body: payload.body, reply_to_id: payload.thread_id || "" }),
          });
          const result = await resp.json();
          appendMessage("action", `Reply sent: ${result.status || "ok"}`);
        } catch (e) { appendMessage("action", `Failed to reply: ${e.message}`); }
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
        } catch (e) { appendMessage("action", `Search failed: ${e.message}`); }
        break;
      }

      case "summarize_thread":
      case "extract_action_items":
        appendMessage("action", `${type}: See reply above.`);
        break;

      default:
        // For browser/google workspace — these are text-only responses
        appendMessage("action", `${type}: ${JSON.stringify(payload).substring(0, 200)}`);
    }
  }

  // ── UI Helpers ────────────────────────────────────────────────────────────

  function appendMessage(role, text, imageCount) {
    const history = document.getElementById("tsifl-chat-history");
    if (!history) return;
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

  // ── Drag & Drop to move sidebar ──────────────────────────────────────────

  function initDrag() {
    const handle = document.getElementById("tsifl-drag-handle");
    const sidebar = document.getElementById("tsifl-sidebar");
    if (!handle || !sidebar) return;

    let isDragging = false;
    let startX, startY, startLeft, startTop;

    handle.addEventListener("mousedown", (e) => {
      isDragging = true;
      startX = e.clientX;
      startY = e.clientY;
      const rect = sidebar.getBoundingClientRect();
      startLeft = rect.left;
      startTop = rect.top;
      sidebar.style.transition = "none";
      e.preventDefault();
    });

    document.addEventListener("mousemove", (e) => {
      if (!isDragging) return;
      const dx = e.clientX - startX;
      const dy = e.clientY - startY;
      sidebar.style.left = `${startLeft + dx}px`;
      sidebar.style.top = `${startTop + dy}px`;
      sidebar.style.right = "auto";
    });

    document.addEventListener("mouseup", () => {
      if (isDragging) {
        isDragging = false;
        sidebar.style.transition = "";
      }
    });
  }

  // ── Resize handle ────────────────────────────────────────────────────────

  function initResize() {
    const handle = document.getElementById("tsifl-resize-handle");
    const sidebar = document.getElementById("tsifl-sidebar");
    if (!handle || !sidebar) return;

    let isResizing = false;
    let startX, startWidth;

    handle.addEventListener("mousedown", (e) => {
      isResizing = true;
      startX = e.clientX;
      startWidth = sidebar.offsetWidth;
      sidebar.style.transition = "none";
      e.preventDefault();
    });

    document.addEventListener("mousemove", (e) => {
      if (!isResizing) return;
      const dx = startX - e.clientX;
      const newWidth = Math.max(280, Math.min(600, startWidth + dx));
      sidebar.style.width = `${newWidth}px`;
    });

    document.addEventListener("mouseup", () => {
      if (isResizing) {
        isResizing = false;
        sidebar.style.transition = "";
      }
    });
  }

  // ── Event Wiring ──────────────────────────────────────────────────────────

  function wireEvents() {
    document.getElementById("tsifl-close-btn").onclick = () => {
      sidebarVisible = false;
      document.getElementById("tsifl-sidebar").classList.add("hidden");
      document.getElementById("tsifl-fab").classList.add("hidden");
    };

    document.getElementById("tsifl-minimize-btn").onclick = () => {
      sidebarMinimized = true;
      sidebarVisible = false;
      document.getElementById("tsifl-sidebar").classList.add("hidden");
      document.getElementById("tsifl-fab").classList.remove("hidden");
    };

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
    inputArea.addEventListener("dragover", (e) => {
      e.preventDefault();
      inputArea.style.background = "#EBF3FB";
    });
    inputArea.addEventListener("dragleave", () => { inputArea.style.background = ""; });
    inputArea.addEventListener("drop", (e) => {
      e.preventDefault();
      inputArea.style.background = "";
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

    // Logout
    const logoutHandler = () => {
      localStorage.removeItem("tsifl_session");
      currentUser = null;
      document.getElementById("tsifl-chat-history").innerHTML = "";
      showLogin();
    };
    // Add logout on double-click of user bar
    document.getElementById("tsifl-user-bar").addEventListener("dblclick", logoutHandler);
  }

  // ── Toggle Sidebar ────────────────────────────────────────────────────────

  function toggleSidebar() {
    const sidebar = document.getElementById("tsifl-sidebar");
    const fab = document.getElementById("tsifl-fab");
    if (!sidebar) { createSidebar(); return; }

    if (sidebarMinimized) {
      sidebarMinimized = false;
      fab.classList.add("hidden");
      sidebar.classList.remove("hidden");
      sidebarVisible = true;
    } else {
      sidebarVisible = !sidebarVisible;
      sidebar.classList.toggle("hidden", !sidebarVisible);
      if (!sidebarVisible && fab) fab.classList.add("hidden");
    }

    // Update site badge and placeholder
    const badge = document.getElementById("tsifl-site-badge");
    if (badge) badge.textContent = getSiteLabel();
    const input = document.getElementById("tsifl-user-input");
    if (input) input.placeholder = getPlaceholder();
    const userBar = document.getElementById("tsifl-user-bar");
    if (userBar && currentUser) userBar.textContent = `${currentUser.email} · ${getSiteLabel()}`;
  }

  // ── Listen for messages ──────────────────────────────────────────────────

  chrome.runtime.onMessage.addListener((msg) => {
    if (msg.action === "toggle_sidebar") {
      if (!document.getElementById("tsifl-sidebar")) createSidebar();
      toggleSidebar();
    }
  });

  // ── Auto-create sidebar (hidden) ────────────────────────────────────────
  createSidebar();

})();
