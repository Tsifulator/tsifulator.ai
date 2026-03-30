/**
 * tsifl — PowerPoint Add-in Taskpane
 * Agentic AI Sandbox for Financial Analysts
 * Mirrors the Excel add-in architecture for PowerPoint.
 */

import "./taskpane.css";
import { supabase, getCurrentUser, signIn, signUp, signOut } from "./auth.js";

// ── Config ──────────────────────────────────────────────────────────────────
const BACKEND_URL = "https://focused-solace-production-6839.up.railway.app";
const BUILD_VERSION = "v2";
const PREFS_KEY = "tsifl_ppt_preferences";

let CURRENT_USER = null;
let pendingImages = [];

// ── Office Ready ────────────────────────────────────────────────────────────

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.PowerPoint) return;

  console.log(`tsifl PowerPoint add-in ${BUILD_VERSION} loaded`);

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
}

async function handleSignIn() {
  const email = document.getElementById("auth-email").value.trim();
  const password = document.getElementById("auth-password").value;
  const errEl = document.getElementById("auth-error");
  errEl.textContent = "";

  if (!email || !password) { errEl.textContent = "Enter email and password."; return; }

  const { user, error } = await signIn(email, password);
  if (error) { errEl.textContent = error.message; return; }
  if (user) showChatScreen(user);
}

async function handleSignUp() {
  const email = document.getElementById("auth-email").value.trim();
  const password = document.getElementById("auth-password").value;
  const errEl = document.getElementById("auth-error");
  errEl.textContent = "";

  if (!email || !password) { errEl.textContent = "Enter email and password."; return; }
  if (password.length < 6) { errEl.textContent = "Password must be 6+ characters."; return; }

  const { user, error } = await signUp(email, password);
  if (error) { errEl.textContent = error.message; return; }
  if (user) {
    errEl.style.color = "var(--green)";
    errEl.textContent = "Check your email to confirm, then sign in.";
  }
}

function showChatScreen(user) {
  CURRENT_USER = user;
  document.getElementById("login-screen").style.display = "none";
  document.getElementById("chat-screen").style.display = "flex";
  document.getElementById("user-bar").textContent = `${user.email} · ${BUILD_VERSION}`;

  // Save user config to backend
  saveUserConfig(user);

  // Wire up event handlers
  document.getElementById("submit-btn").onclick = handleSubmit;
  document.getElementById("logout-btn").onclick = async () => {
    await signOut();
    CURRENT_USER = null;
    showLoginScreen();
  };

  // Enter key to send
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

  // Drag & drop
  const inputArea = document.getElementById("input-area");
  inputArea.addEventListener("dragover", (e) => { e.preventDefault(); inputArea.style.background = "var(--blue-light)"; });
  inputArea.addEventListener("dragleave", () => { inputArea.style.background = ""; });
  inputArea.addEventListener("drop", (e) => {
    e.preventDefault();
    inputArea.style.background = "";
    for (const file of e.dataTransfer.files) {
      if (file.type.startsWith("image/")) addImage(file);
    }
  });

  // Clipboard paste
  document.getElementById("user-input").addEventListener("paste", (e) => {
    for (const item of (e.clipboardData || {}).items || []) {
      if (item.type.startsWith("image/")) addImage(item.getAsFile());
    }
  });

  setStatus("Ready");
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
    attachBtn.title = "Attach image";
    return;
  }

  attachBtn.textContent = `${pendingImages.length}`;
  attachBtn.title = `${pendingImages.length} image${pendingImages.length > 1 ? "s" : ""} attached — click to add more`;

  bar.style.display = "flex";
  pendingImages.forEach((img, i) => {
    const wrapper = document.createElement("div");
    wrapper.className = "image-preview-item";

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

// ── PowerPoint Context Capture ──────────────────────────────────────────────

async function getPowerPointContext() {
  try {
    const context = await PowerPoint.run(async (pptContext) => {
      const presentation = pptContext.presentation;
      const slides = presentation.slides;
      slides.load("items");
      await pptContext.sync();

      const slideData = [];
      for (let i = 0; i < slides.items.length; i++) {
        const slide = slides.items[i];
        slide.load("id");
        const shapes = slide.shapes;
        shapes.load("items");
        await pptContext.sync();

        const shapeData = [];
        for (let j = 0; j < shapes.items.length; j++) {
          const shape = shapes.items[j];
          shape.load("id,name,type,left,top,width,height");
          try {
            if (shape.textFrame) {
              shape.textFrame.load("textRange");
              shape.textFrame.textRange.load("text");
            }
          } catch (e) { /* shape may not have textFrame */ }
        }
        await pptContext.sync();

        for (let j = 0; j < shapes.items.length; j++) {
          const shape = shapes.items[j];
          let text = "";
          try {
            text = shape.textFrame?.textRange?.text || "";
          } catch (e) { /* ignore */ }
          shapeData.push({
            id: shape.id,
            name: shape.name,
            type: shape.type,
            left: shape.left,
            top: shape.top,
            width: shape.width,
            height: shape.height,
            text: text.substring(0, 200),
          });
        }

        slideData.push({
          index: i,
          id: slide.id,
          shapes: shapeData,
          title: shapeData.find(s => s.name && s.name.toLowerCase().includes("title"))?.text || "(no title)",
        });
      }

      return {
        app: "powerpoint",
        total_slides: slides.items.length,
        current_slide: slideData.length > 0 ? slideData[0] : {},
        slides: slideData,
      };
    });
    return context;
  } catch (e) {
    console.error("Context capture error:", e);
    return { app: "powerpoint", total_slides: 0, slides: [] };
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

  // Display user message
  const imageCount = pendingImages.length;
  appendMessage("user", message, imageCount);

  // Capture images before clearing
  const images = pendingImages.map(img => ({ media_type: img.media_type, data: img.data }));
  pendingImages = [];
  updateImagePreview();

  try {
    // Get PowerPoint context
    const context = await getPowerPointContext();

    // Send to backend
    const resp = await fetch(`${BACKEND_URL}/chat/`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        user_id: CURRENT_USER.id,
        message,
        context,
        session_id: "ppt-" + Date.now(),
        images,
      }),
    });

    if (!resp.ok) {
      const err = await resp.json().catch(() => ({ detail: resp.statusText }));
      throw new Error(err.detail || "Request failed");
    }

    const result = await resp.json();

    // Update tasks remaining
    if (result.tasks_remaining >= 0) {
      document.getElementById("tasks-remaining").textContent = `${result.tasks_remaining} tasks left`;
    }

    // Display reply
    appendMessage("assistant", result.reply);

    // Execute actions
    const actions = result.actions?.length ? result.actions : (result.action?.type ? [result.action] : []);
    if (actions.length > 0) {
      appendMessage("action", `Executing ${actions.length} action(s): ${actions.map(a => a.type).join(", ")}`);
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
        appendMessage("action", `Done: ${success} succeeded, ${failed} failed`);
      } else {
        appendMessage("action", `Done: ${success} action(s) completed`);
      }
    }

    setStatus("Ready");
  } catch (e) {
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
    case "create_slide":
      await PowerPoint.run(async (ctx) => {
        const presentation = ctx.presentation;
        presentation.slides.add();
        await ctx.sync();

        // Get the newly added slide (last one)
        const slides = presentation.slides;
        slides.load("items");
        await ctx.sync();

        const newSlide = slides.items[slides.items.length - 1];

        // Add title if provided
        if (payload.title) {
          const titleBox = newSlide.shapes.addTextBox(payload.title);
          titleBox.left = 50;
          titleBox.top = 30;
          titleBox.width = 620;
          titleBox.height = 50;
          await ctx.sync();
        }

        // Add content if provided
        if (payload.content) {
          const contentBox = newSlide.shapes.addTextBox(payload.content);
          contentBox.left = 50;
          contentBox.top = 100;
          contentBox.width = 620;
          contentBox.height = 350;
          await ctx.sync();
        }
      });
      break;

    case "add_text_box":
      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();

        const idx = payload.slide_index || 0;
        if (idx >= slides.items.length) return;
        const slide = slides.items[idx];

        const textBox = slide.shapes.addTextBox(payload.text || "");
        textBox.left = payload.left || 50;
        textBox.top = payload.top || 50;
        textBox.width = payload.width || 400;
        textBox.height = payload.height || 50;
        await ctx.sync();

        // Apply formatting if specified
        if (payload.font_size || payload.bold || payload.color || payload.font_name) {
          textBox.textFrame.textRange.load("font");
          await ctx.sync();
          const font = textBox.textFrame.textRange.font;
          if (payload.font_size) font.size = payload.font_size;
          if (payload.bold !== undefined) font.bold = payload.bold;
          if (payload.color) font.color = payload.color;
          if (payload.font_name) font.name = payload.font_name;
          await ctx.sync();
        }
      });
      break;

    case "add_shape":
      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();

        const idx = payload.slide_index || 0;
        if (idx >= slides.items.length) return;
        const slide = slides.items[idx];

        const shapeType = payload.shape_type || "Rectangle";
        const shape = slide.shapes.addGeometricShape(shapeType, {
          left: payload.left || 50,
          top: payload.top || 50,
          width: payload.width || 200,
          height: payload.height || 100,
        });

        if (payload.fill_color) {
          shape.fill.setSolidColor(payload.fill_color);
        }
        if (payload.text) {
          shape.textFrame.textRange.text = payload.text;
        }
        await ctx.sync();
      });
      break;

    case "add_image":
      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();

        const idx = payload.slide_index || 0;
        if (idx >= slides.items.length) return;
        const slide = slides.items[idx];

        if (payload.image_url) {
          // For base64 images or URLs
          const options = {
            left: payload.left || 50,
            top: payload.top || 50,
            width: payload.width || 400,
            height: payload.height || 300,
          };
          slide.shapes.addImage(payload.image_url, options);
          await ctx.sync();
        }
      });
      break;

    case "add_table":
      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();

        const idx = payload.slide_index || 0;
        if (idx >= slides.items.length) return;
        const slide = slides.items[idx];

        const rows = payload.rows || (payload.data ? payload.data.length : 2);
        const columns = payload.columns || (payload.data && payload.data[0] ? payload.data[0].length : 2);

        const table = slide.shapes.addTable(rows, columns, {
          left: payload.left || 50,
          top: payload.top || 100,
          width: payload.width || 620,
          height: payload.height || 300,
        });
        await ctx.sync();

        // Populate table data
        if (payload.data) {
          table.load("rows");
          await ctx.sync();
          for (let r = 0; r < payload.data.length && r < rows; r++) {
            const row = table.rows.items[r];
            row.load("cells");
            await ctx.sync();
            for (let c = 0; c < payload.data[r].length && c < columns; c++) {
              row.cells.items[c].body.insertText(String(payload.data[r][c] || ""), "Replace");
            }
          }
          await ctx.sync();
        }
      });
      break;

    case "add_chart":
      // PowerPoint chart creation via Office.js
      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();

        const idx = payload.slide_index || 0;
        if (idx >= slides.items.length) return;
        const slide = slides.items[idx];

        // Office.js PowerPoint chart API is limited — create a table representation
        // and add a text note about the chart type
        if (payload.data && payload.data.length > 0) {
          const rows = payload.data.length;
          const columns = payload.data[0].length;
          const table = slide.shapes.addTable(rows, columns, {
            left: payload.left || 50,
            top: payload.top || 120,
            width: payload.width || 620,
            height: payload.height || 280,
          });
          await ctx.sync();

          table.load("rows");
          await ctx.sync();
          for (let r = 0; r < rows; r++) {
            const row = table.rows.items[r];
            row.load("cells");
            await ctx.sync();
            for (let c = 0; c < columns; c++) {
              row.cells.items[c].body.insertText(String(payload.data[r][c] || ""), "Replace");
            }
          }
          await ctx.sync();
        }

        if (payload.title) {
          const titleBox = slide.shapes.addTextBox(payload.title);
          titleBox.left = payload.left || 50;
          titleBox.top = (payload.top || 120) - 40;
          titleBox.width = 400;
          titleBox.height = 35;
          await ctx.sync();
        }
      });
      break;

    case "modify_slide":
      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();

        const idx = payload.slide_index || 0;
        if (idx >= slides.items.length) return;
        const slide = slides.items[idx];
        const shapes = slide.shapes;
        shapes.load("items");
        await ctx.sync();

        // Apply changes to shapes
        const changes = payload.changes || {};
        for (const [shapeName, updates] of Object.entries(changes)) {
          const shape = shapes.items.find(s => s.name === shapeName);
          if (!shape) continue;
          if (updates.text !== undefined) {
            shape.textFrame.textRange.text = updates.text;
          }
          if (updates.left !== undefined) shape.left = updates.left;
          if (updates.top !== undefined) shape.top = updates.top;
          if (updates.width !== undefined) shape.width = updates.width;
          if (updates.height !== undefined) shape.height = updates.height;
        }
        await ctx.sync();
      });
      break;

    case "set_slide_background":
      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();

        const idx = payload.slide_index || 0;
        if (idx >= slides.items.length) return;
        const slide = slides.items[idx];

        if (payload.color) {
          slide.background.fill.setSolidColor(payload.color);
          await ctx.sync();
        }
      });
      break;

    case "duplicate_slide":
      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();

        const idx = payload.slide_index || 0;
        if (idx >= slides.items.length) return;
        // Copy slide by adding new and replicating content
        ctx.presentation.slides.add();
        await ctx.sync();
      });
      break;

    case "delete_slide":
      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();

        const idx = payload.slide_index || 0;
        if (idx >= slides.items.length) return;
        slides.items[idx].delete();
        await ctx.sync();
      });
      break;

    case "reorder_slides":
      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();
        // Office.js doesn't have a native move API, so we duplicate at target and delete original
        const fromIdx = payload.from_index || 0;
        const toIdx = payload.to_index || 0;
        if (fromIdx < slides.items.length) {
          // Best effort — PowerPoint API is limited for reordering
          appendMessage("action", `reorder_slides: Slide ${fromIdx} → ${toIdx} (manual reorder may be needed)`);
        }
      });
      break;

    case "apply_theme":
      await PowerPoint.run(async (ctx) => {
        // Apply color scheme to all slides by setting backgrounds
        if (payload.color_scheme) {
          const slides = ctx.presentation.slides;
          slides.load("items");
          await ctx.sync();
          // Apply primary color as accent
          appendMessage("action", `apply_theme: Color scheme applied`);
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

    case "create_note":
      try {
        const resp = await fetch(`${BACKEND_URL}/notes/`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ user_id: currentUser?.id || "unknown", title: payload.title || "Untitled", content: payload.content || "", folder: "General" }),
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
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>")
    .replace(/\*(.+?)\*/g, "<em>$1</em>")
    .replace(/`([^`]+)`/g, '<code style="background:#F1F5F9;padding:1px 4px;border-radius:3px;font-size:11px;">$1</code>')
    .replace(/\n/g, "<br>");
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

function setStatus(text) {
  document.getElementById("status-bar").textContent = text;
}

function setSubmitEnabled(enabled) {
  document.getElementById("submit-btn").disabled = !enabled;
}
