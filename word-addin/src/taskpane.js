/**
 * tsifl — Word Add-in Taskpane
 * Agentic AI Sandbox for Financial Analysts
 * Mirrors the Excel add-in architecture for Word.
 */

import "./taskpane.css";
import { supabase, getCurrentUser, signIn, signUp, signOut, resetPassword, syncSessionToBackend } from "./auth.js";

// ── Config ──────────────────────────────────────────────────────────────────
const BACKEND_URL = "https://focused-solace-production-6839.up.railway.app";
const BUILD_VERSION = "v3";

let CURRENT_USER = null;
let pendingImages = [];

// ── Office Ready ────────────────────────────────────────────────────────────

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Word) return;

  console.log(`tsifl Word add-in ${BUILD_VERSION} loaded`);

  const user = await getCurrentUser();
  if (user) {
    showChatScreen(user);
  } else {
    showLoginScreen();
  }
});

// ── Auth Screens ────────────────────────────────────────────────────────────

function showLoginScreen() {
  document.getElementById("login-screen").style.display = "flex";
  document.getElementById("chat-screen").style.display = "none";

  document.getElementById("login-btn").onclick = handleSignIn;
  document.getElementById("signup-btn").onclick = handleSignUp;
  // Show/hide password toggle (Improvement 9)
  const togglePw = document.getElementById("toggle-pw-btn");
  if (togglePw) {
    togglePw.onclick = () => {
      const pwInput = document.getElementById("auth-password");
      pwInput.type = pwInput.type === "password" ? "text" : "password";
    };
  }
  // Forgot password (Improvement 5)
  const forgotBtn = document.getElementById("forgot-pw-btn");
  if (forgotBtn) {
    forgotBtn.onclick = async () => {
      const email = document.getElementById("auth-email").value.trim();
      const errEl = document.getElementById("auth-error");
      if (!email || !email.includes("@")) { errEl.style.color = "#DC2626"; errEl.textContent = "Enter your email first."; return; }
      const { error } = await resetPassword(email);
      if (error) { errEl.style.color = "#DC2626"; errEl.textContent = error.message; return; }
      errEl.style.color = "#16A34A"; errEl.textContent = "Check your email for a reset link.";
    };
  }
}

async function handleSignIn() {
  const email = document.getElementById("auth-email").value.trim();
  const password = document.getElementById("auth-password").value;
  const errEl = document.getElementById("auth-error");
  errEl.style.color = "#DC2626";
  errEl.textContent = "";

  if (!email || !email.includes("@")) { errEl.textContent = "Enter a valid email address."; return; }
  if (!password) { errEl.textContent = "Enter your password."; return; }

  const { user, error } = await signIn(email, password);
  if (error) { errEl.textContent = error.message; return; }
  if (user) showChatScreen(user);
}

async function handleSignUp() {
  const email = document.getElementById("auth-email").value.trim();
  const password = document.getElementById("auth-password").value;
  const errEl = document.getElementById("auth-error");
  errEl.style.color = "#DC2626";
  errEl.textContent = "";

  if (!email || !email.includes("@")) { errEl.textContent = "Enter a valid email address."; return; }
  if (password.length < 6) { errEl.textContent = "Password must be at least 6 characters."; return; }

  const { user, error } = await signUp(email, password);
  if (error) { errEl.textContent = error.message; return; }
  errEl.style.color = "#16A34A";
  errEl.textContent = "Check your email to confirm, then sign in.";
}

function showChatScreen(user) {
  CURRENT_USER = user;
  document.getElementById("login-screen").style.display = "none";
  document.getElementById("chat-screen").style.display = "flex";

  // User display with avatar initial (Improvement 8)
  const initial = (user.email || "?")[0].toUpperCase();
  document.getElementById("user-bar").innerHTML =
    `<span style="display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;border-radius:50%;background:#0D5EAF;color:white;font-size:10px;font-weight:700;margin-right:4px;">${initial}</span>${user.email} &middot; ${BUILD_VERSION}`;

  saveUserConfig(user);

  document.getElementById("submit-btn").onclick = handleSubmit;
  document.getElementById("logout-btn").onclick = async () => {
    await signOut();
    CURRENT_USER = null;
    showLoginScreen();
  };

  document.getElementById("user-input").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleSubmit(); }
  });

  // Image handling
  const imageInput = document.getElementById("image-input");
  document.getElementById("attach-btn").onclick = () => imageInput.click();
  imageInput.onchange = (e) => {
    for (const file of e.target.files) addImage(file);
    imageInput.value = "";
  };

  const inputArea = document.getElementById("input-area");
  inputArea.addEventListener("dragover", (e) => { e.preventDefault(); inputArea.style.background = "var(--blue-light)"; });
  inputArea.addEventListener("dragleave", () => { inputArea.style.background = ""; });
  inputArea.addEventListener("drop", (e) => {
    e.preventDefault();
    inputArea.style.background = "";
    for (const file of e.dataTransfer.files) {
      addFile(file);
    }
  });

  document.getElementById("user-input").addEventListener("paste", (e) => {
    for (const item of (e.clipboardData || {}).items || []) {
      if (item.type.startsWith("image/") || item.kind === "file") {
        const f = item.getAsFile();
        if (f) addFile(f);
      }
    }
  });

  // Notes button
  document.getElementById("notes-btn").onclick = () => {
    window.open(`${BACKEND_URL}/notes-app`, "_blank");
  };

  // Templates dropdown (Improvement 30)
  const templatesBtn = document.getElementById("templates-btn");
  const templatesDd = document.getElementById("templates-dropdown");
  if (templatesBtn && templatesDd) {
    templatesBtn.onclick = (e) => {
      e.stopPropagation();
      templatesDd.style.display = templatesDd.style.display === "none" ? "block" : "none";
    };
    document.querySelectorAll(".template-item").forEach(item => {
      item.onclick = () => {
        document.getElementById("user-input").value = `Create a ${item.dataset.template} document with professional formatting`;
        templatesDd.style.display = "none";
        handleSubmit();
      };
    });
    document.addEventListener("click", () => { templatesDd.style.display = "none"; });
  }

  // Document outline toggle (Improvement 26)
  const outlineToggle = document.getElementById("outline-toggle");
  if (outlineToggle) {
    outlineToggle.onclick = () => {
      const panel = document.getElementById("outline-panel");
      panel.style.display = panel.style.display === "none" ? "block" : "none";
      if (panel.style.display === "block") loadDocumentOutline();
    };
  }

  // Quick action buttons
  document.querySelectorAll(".quick-btn").forEach(btn => {
    btn.onclick = () => {
      document.getElementById("user-input").value = btn.dataset.prompt;
      handleSubmit();
    };
  });

  // Auto-resize textarea
  document.getElementById("user-input").addEventListener("input", (e) => {
    e.target.style.height = "auto";
    e.target.style.height = Math.min(e.target.scrollHeight, 100) + "px";
  });

  // Escape to clear input
  document.getElementById("user-input").addEventListener("keydown", (e) => {
    if (e.key === "Escape") { e.target.value = ""; e.target.style.height = "auto"; }
  });

  setStatus("Ready");
  updateWordCount();
}

// Document outline (Improvement 26)
async function loadDocumentOutline() {
  try {
    await Word.run(async (ctx) => {
      const paragraphs = ctx.document.body.paragraphs;
      paragraphs.load("items/text,items/style");
      await ctx.sync();
      const headings = [];
      for (const p of paragraphs.items) {
        if (p.style && p.style.match(/^Heading\s?[1-3]$/i)) {
          const level = parseInt(p.style.replace(/\D/g, "")) || 1;
          headings.push({ text: p.text, level });
        }
      }
      const list = document.getElementById("outline-list");
      if (headings.length === 0) {
        list.innerHTML = '<div style="color:#94A3B8;font-style:italic;">No headings found</div>';
      } else {
        list.innerHTML = headings.map(h =>
          `<div style="padding:2px 0;padding-left:${(h.level - 1) * 12}px;cursor:pointer;color:#0D5EAF;" class="outline-item">${h.text}</div>`
        ).join("");
      }
    });
  } catch (e) {
    document.getElementById("outline-list").innerHTML = `<div style="color:#DC2626;">${e.message}</div>`;
  }
}

// Word count update (Improvement 27)
async function updateWordCount() {
  try {
    await Word.run(async (ctx) => {
      const body = ctx.document.body;
      body.load("text");
      await ctx.sync();
      const text = body.text || "";
      const words = text.trim().split(/\s+/).filter(w => w.length > 0).length;
      const chars = text.length;
      const readTime = Math.max(1, Math.ceil(words / 250));
      const bar = document.getElementById("word-count-bar");
      if (bar) bar.textContent = `${words} words \u00b7 ${chars} characters \u00b7 ~${readTime} min read`;
    });
  } catch (e) { /* silent */ }
}

async function saveUserConfig(user) {
  try {
    await fetch(`${BACKEND_URL}/auth/set-user`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ user_id: user.id }),
    });
  } catch (e) { /* ignore */ }
}

// ── Image Handling ──────────────────────────────────────────────────────────

function addFile(file) {
  const reader = new FileReader();
  reader.onload = () => {
    const base64 = reader.result.split(",")[1];
    pendingImages.push({
      media_type: file.type || (file.name && file.name.match(/\.(png|jpg|jpeg|gif|webp)$/i) ? "image/png" : "application/octet-stream"),
      data: base64,
      preview: reader.result,
      file_name: file.name || "",
    });
    updateImagePreview();
  };
  reader.readAsDataURL(file);
}

// Backward compat alias
function addImage(file) { addFile(file); }

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

  attachBtn.textContent = `${pendingImages.length}`;
  attachBtn.title = `${pendingImages.length} file${pendingImages.length > 1 ? "s" : ""} attached — click to add more`;

  bar.style.display = "flex";
  pendingImages.forEach((img, i) => {
    const wrapper = document.createElement("div");
    wrapper.className = "image-preview-item";
    const isImage = img.media_type.startsWith("image/");

    if (isImage) {
      renderImageToCanvas(img.data, img.media_type, 48, 48).then(canvas => {
        if (canvas) {
          canvas.style.borderRadius = "4px";
          canvas.style.border = "1px solid var(--border)";
          wrapper.insertBefore(canvas, wrapper.firstChild);
        }
      });
    } else {
      const docIcon = document.createElement("div");
      const ext = img.file_name ? img.file_name.split(".").pop().toUpperCase() : "FILE";
      docIcon.style.cssText = "width:48px;height:48px;display:flex;align-items:center;justify-content:center;background:#F1F5F9;border-radius:4px;border:1px solid var(--border);font-size:9px;font-weight:700;color:#0D5EAF;text-align:center;";
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

// ── Word Context Capture ────────────────────────────────────────────────────

async function getWordContext() {
  try {
    return await Word.run(async (wordContext) => {
      const body = wordContext.document.body;
      body.load("text");

      const paragraphs = body.paragraphs;
      paragraphs.load("items");
      await wordContext.sync();

      // Load each paragraph's details
      const paraData = [];
      const limit = Math.min(paragraphs.items.length, 50);
      for (let i = 0; i < limit; i++) {
        const p = paragraphs.items[i];
        p.load("text,style,alignment");
      }
      await wordContext.sync();

      for (let i = 0; i < limit; i++) {
        const p = paragraphs.items[i];
        paraData.push({
          text: p.text,
          style: p.style,
          alignment: p.alignment,
        });
      }

      // Get tables
      const tables = body.tables;
      tables.load("items");
      await wordContext.sync();

      const tableData = [];
      for (let i = 0; i < Math.min(tables.items.length, 5); i++) {
        const t = tables.items[i];
        t.load("rowCount,values");
        await wordContext.sync();
        tableData.push({
          rows: t.rowCount,
          columns: t.values && t.values[0] ? t.values[0].length : 0,
        });
      }

      // Get selection
      const selection = wordContext.document.getSelection();
      selection.load("text");
      await wordContext.sync();

      return {
        app: "word",
        total_paragraphs: paragraphs.items.length,
        paragraphs: paraData,
        tables: tableData,
        selection: selection.text || "",
      };
    });
  } catch (e) {
    console.error("Context capture error:", e);
    return { app: "word", total_paragraphs: 0, paragraphs: [], tables: [] };
  }
}

// ── Chat Submit ─────────────────────────────────────────────────────────────

async function handleSubmit() {
  const input = document.getElementById("user-input");
  const message = input.value.trim();
  if (!message && pendingImages.length === 0) return;

  input.value = "";
  setSubmitEnabled(false);
  setStatus("Thinking...");
  showTypingIndicator();

  const imageCount = pendingImages.length;
  appendMessage("user", message, imageCount);

  const images = pendingImages.map(img => ({ media_type: img.media_type, data: img.data }));
  pendingImages = [];
  updateImagePreview();

  try {
    const context = await getWordContext();

    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 120000);
    const resp = await fetch(`${BACKEND_URL}/chat/`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      signal: controller.signal,
      body: JSON.stringify({
        user_id: CURRENT_USER.id,
        message,
        context,
        session_id: "word-" + Date.now(),
        images,
      }),
    });
    clearTimeout(timeout);

    if (!resp.ok) {
      const err = await resp.json().catch(() => ({ detail: resp.statusText }));
      throw new Error(err.detail || "Request failed");
    }

    const result = await resp.json();

    if (result.tasks_remaining >= 0) {
      document.getElementById("tasks-remaining").textContent = `${result.tasks_remaining} tasks left`;
    }

    hideTypingIndicator();
    appendMessage("assistant", result.reply);

    const actions = result.actions?.length ? result.actions : (result.action?.type ? [result.action] : []);
    if (actions.length > 0) {
      let success = 0, failed = 0;
      for (const action of actions) {
        try {
          await executeAction(action);
          success++;
        } catch (e) {
          console.error("Action failed:", action.type, e);
          failed++;
        }
      }
      if (failed > 0) {
        appendMessage("action", `${success} succeeded, ${failed} failed`);
      }
    }

    setStatus("Ready");
  } catch (e) {
    hideTypingIndicator();
    appendMessage("assistant", `Error: ${e.message}`);
    setStatus("Error — try again");
  }

  setSubmitEnabled(true);
}

// ── Action Executor ─────────────────────────────────────────────────────────

async function executeAction(action) {
  const { type, payload } = action;
  if (!payload) return;

  switch (type) {
    case "insert_text":
      await Word.run(async (ctx) => {
        const body = ctx.document.body;
        const position = payload.position || "end";
        if (position === "replace_selection") {
          const sel = ctx.document.getSelection();
          sel.insertText(payload.text, "Replace");
        } else if (position === "start") {
          body.insertText(payload.text, "Start");
        } else if (position === "after_selection") {
          const sel = ctx.document.getSelection();
          sel.insertText(payload.text, "After");
        } else {
          body.insertText(payload.text, "End");
        }
        await ctx.sync();
      });
      break;

    case "insert_paragraph":
      await Word.run(async (ctx) => {
        const body = ctx.document.body;
        const para = body.insertParagraph(payload.text || "", "End");

        if (payload.style) {
          para.style = payload.style;
        }
        if (payload.alignment) {
          const alignMap = { left: "Left", center: "Centered", right: "Right", justify: "Justified" };
          para.alignment = alignMap[payload.alignment.toLowerCase()] || payload.alignment;
        }
        if (payload.spacing_after !== undefined) {
          para.spaceAfter = payload.spacing_after;
        }
        if (payload.spacing_before !== undefined) {
          para.spaceBefore = payload.spacing_before;
        }

        await ctx.sync();
      });
      break;

    case "insert_table":
      await Word.run(async (ctx) => {
        const body = ctx.document.body;
        const rows = payload.rows || (payload.data ? payload.data.length : 2);
        const cols = payload.columns || (payload.data && payload.data[0] ? payload.data[0].length : 2);

        // Build 2D string array
        const values = [];
        for (let r = 0; r < rows; r++) {
          const row = [];
          for (let c = 0; c < cols; c++) {
            row.push(payload.data && payload.data[r] && payload.data[r][c] != null
              ? String(payload.data[r][c]) : "");
          }
          values.push(row);
        }

        const table = body.insertTable(rows, cols, "End", values);
        if (payload.style) {
          table.style = payload.style;
        }
        await ctx.sync();
      });
      break;

    case "insert_image":
      await Word.run(async (ctx) => {
        if (payload.image_data) {
          const body = ctx.document.body;
          const inlinePic = body.insertInlinePictureFromBase64(payload.image_data, "End");
          if (payload.width) inlinePic.width = payload.width;
          if (payload.height) inlinePic.height = payload.height;
          await ctx.sync();
        }
      });
      break;

    case "format_text":
      await Word.run(async (ctx) => {
        // Use search to find the range described
        const searchResults = ctx.document.body.search(payload.range_description || "", { matchCase: true });
        searchResults.load("items");
        await ctx.sync();

        if (searchResults.items.length > 0) {
          const range = searchResults.items[0];
          const font = range.font;
          if (payload.bold !== undefined) font.bold = payload.bold;
          if (payload.italic !== undefined) font.italic = payload.italic;
          if (payload.underline !== undefined) font.underline = payload.underline ? "Single" : "None";
          if (payload.font_size) font.size = payload.font_size;
          if (payload.font_color) font.color = payload.font_color;
          if (payload.font_name) font.name = payload.font_name;
          if (payload.highlight_color) font.highlightColor = payload.highlight_color;
          await ctx.sync();
        }
      });
      break;

    case "insert_header":
      await Word.run(async (ctx) => {
        const sections = ctx.document.sections;
        sections.load("items");
        await ctx.sync();

        if (sections.items.length > 0) {
          const headerType = payload.type || "primary";
          const headerMap = { primary: "Primary", firstPage: "FirstPage", evenPages: "EvenPages" };
          const header = sections.items[0].getHeader(headerMap[headerType] || "Primary");
          header.insertParagraph(payload.text || "", "End");
          await ctx.sync();
        }
      });
      break;

    case "insert_footer":
      await Word.run(async (ctx) => {
        const sections = ctx.document.sections;
        sections.load("items");
        await ctx.sync();

        if (sections.items.length > 0) {
          const footerType = payload.type || "primary";
          const footerMap = { primary: "Primary", firstPage: "FirstPage", evenPages: "EvenPages" };
          const footer = sections.items[0].getFooter(footerMap[footerType] || "Primary");
          footer.insertParagraph(payload.text || "", "End");
          await ctx.sync();
        }
      });
      break;

    case "insert_page_break":
      await Word.run(async (ctx) => {
        ctx.document.body.insertBreak("Page", "End");
        await ctx.sync();
      });
      break;

    case "insert_section_break":
      await Word.run(async (ctx) => {
        const typeMap = {
          continuous: "Continuous",
          nextPage: "Next",
          evenPage: "EvenPage",
          oddPage: "OddPage",
        };
        const breakType = typeMap[payload.type || "nextPage"] || "Next";
        ctx.document.body.insertBreak(breakType === "Continuous" ? "SectionContinuous" : "SectionNext", "End");
        await ctx.sync();
      });
      break;

    case "apply_style":
      await Word.run(async (ctx) => {
        const searchResults = ctx.document.body.search(payload.range_description || "", { matchCase: true });
        searchResults.load("items");
        await ctx.sync();

        if (searchResults.items.length > 0) {
          searchResults.items[0].style = payload.style_name;
          await ctx.sync();
        }
      });
      break;

    case "find_and_replace":
      await Word.run(async (ctx) => {
        const searchResults = ctx.document.body.search(payload.find_text, {
          matchCase: payload.match_case || false,
        });
        searchResults.load("items");
        await ctx.sync();

        for (const result of searchResults.items) {
          result.insertText(payload.replace_text, "Replace");
        }
        await ctx.sync();
      });
      break;

    case "insert_table_of_contents":
      await Word.run(async (ctx) => {
        const body = ctx.document.body;
        // Insert TOC heading
        const tocHeading = body.insertParagraph("Table of Contents", "Start");
        tocHeading.style = "Heading 1";
        // Insert placeholder paragraphs for TOC entries
        const tocNote = body.insertParagraph("(Update this field in Word: References > Update Table)", "After");
        tocNote.style = "Normal";
        tocNote.font.italic = true;
        tocNote.font.color = "#64748B";
        tocNote.font.size = 9;
        // Insert page break after TOC
        body.insertBreak("Page", "After");
        await ctx.sync();
      });
      break;

    case "add_comment":
      await Word.run(async (ctx) => {
        const searchResults = ctx.document.body.search(payload.range_description || "", { matchCase: true });
        searchResults.load("items");
        await ctx.sync();

        if (searchResults.items.length > 0) {
          searchResults.items[0].insertComment(payload.comment_text || "");
          await ctx.sync();
        }
      });
      break;

    case "set_page_margins":
      await Word.run(async (ctx) => {
        const sections = ctx.document.sections;
        sections.load("items");
        await ctx.sync();

        if (sections.items.length > 0) {
          const section = sections.items[0];
          section.load("headerDistance,footerDistance");
          await ctx.sync();
          // Margins in points (1 inch = 72 points)
          if (payload.top !== undefined) section.topMargin = payload.top;
          if (payload.bottom !== undefined) section.bottomMargin = payload.bottom;
          if (payload.left !== undefined) section.leftMargin = payload.left;
          if (payload.right !== undefined) section.rightMargin = payload.right;
          await ctx.sync();
        }
      });
      break;

    case "launch_app":
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
      break;

    case "open_notes":
    case "open_url":
      try {
        const url = type === "open_notes" ? `${BACKEND_URL}/notes-app` : (payload.url || "");
        if (url) window.open(url, "_blank");
        appendMessage("action", `Opened: ${url}`);
      } catch (e) {
        appendMessage("action", `open: ${e.message}`);
      }
      break;

    case "create_document":
      try {
        // Note: Word.createDocument() requires specific APIs
        appendMessage("action", "create_document: Please create a new document in Word first");
      } catch (e) {
        appendMessage("action", `create_document: ${e.message}`);
      }
      break;

    case "create_note":
      try {
        const resp = await fetch(`${BACKEND_URL}/notes/`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ user_id: CURRENT_USER?.id || "unknown", title: payload.title || "Untitled", content: payload.content || "", folder: "General" }),
        });
        const note = await resp.json();
        appendMessage("action", `Note created: "${note.title}"`);
      } catch (e) { appendMessage("action", `create_note: ${e.message}`); }
      break;

    default:
      console.warn("Unknown action type:", type);
  }
}

// ── UI Helpers ──────────────────────────────────────────────────────────────

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

function appendMessage(role, text, imageCount) {
  const history = document.getElementById("chat-history");
  const div = document.createElement("div");
  div.className = `message ${role}`;
  if (role === "assistant" && text) {
    div.innerHTML = renderMarkdown(text);
  } else {
    div.textContent = text || "";
  }

  if (imageCount && imageCount > 0) {
    const badge = document.createElement("div");
    badge.className = "image-badge";
    badge.textContent = `${imageCount} image${imageCount > 1 ? "s" : ""} attached`;
    div.appendChild(badge);
  }

  history.appendChild(div);
  history.scrollTop = history.scrollHeight;
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
  'My circuits are tingling...',
  'Give me a sec, genius takes time...',
  'Loading witty response...',
  'Brew yourself a coffee, this is gonna be good...',
  'My keyboard is on fire right now...',
  'Almost done, don\'t go anywhere...',
];

let _thinkingInterval = null;

function showTypingIndicator() {
  hideTypingIndicator();
  const history = document.getElementById("chat-history");
  const div = document.createElement("div");
  div.id = "typing-indicator";
  div.className = "thinking-bubble";
  div.innerHTML = '<div class="thinking-orb"></div><span class="thinking-text"></span>';
  history.appendChild(div);
  history.scrollTop = history.scrollHeight;

  let idx = 0;
  function showNext() {
    const el = document.querySelector('#typing-indicator .thinking-text');
    if (!el) return;
    const msg = _thinkingMessages[idx % _thinkingMessages.length];
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
  const el = document.getElementById("typing-indicator");
  if (el) el.remove();
}

function setStatus(text) {
  document.getElementById("status-bar").textContent = text;
}

function setSubmitEnabled(enabled) {
  document.getElementById("submit-btn").disabled = !enabled;
}
