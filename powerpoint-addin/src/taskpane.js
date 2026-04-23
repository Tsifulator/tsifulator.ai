/**
 * tsifl — PowerPoint Add-in Taskpane
 * Agentic AI Sandbox for Financial Analysts
 * Mirrors the Excel add-in architecture for PowerPoint.
 */

import "./taskpane.css";
import { supabase, getCurrentUser, signIn, signUp, signOut, resetPassword, syncSessionToBackend } from "./auth.js";

// ── Config ──────────────────────────────────────────────────────────────────
const BACKEND_URL = "https://focused-solace-production-6839.up.railway.app";
const BUILD_VERSION = "v4";
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

  // ── Auto-poll for cross-app plot transfers every 4s ───────────────────────
  // Mirrors the Excel add-in's polling pattern. When an R add-in (or server-
  // side plot service) produces an image destined for PowerPoint via
  // /transfer/store with to_app="powerpoint", we pick it up here and insert
  // it into the current slide. Enables the "Excel data → R chart → PPT slide"
  // flow without the user manually triggering an import_image action.
  const seenTransferIds = new Set();
  setInterval(async () => {
    if (!CURRENT_USER) return;
    try {
      const resp = await fetch(`${BACKEND_URL}/transfer/pending/powerpoint`);
      if (!resp.ok) return;
      const { pending = [] } = await resp.json();
      const images = pending.filter(
        p => p.data_type === "image" && !seenTransferIds.has(p.transfer_id)
      );
      for (const item of images) {
        seenTransferIds.add(item.transfer_id);
        try {
          const tResp = await fetch(`${BACKEND_URL}/transfer/${item.transfer_id}`);
          if (!tResp.ok) continue;
          const transfer = await tResp.json();
          const cleanBase64 = String(transfer.data || "")
            .replace(/^data:image\/[a-z+]+;base64,/, "");
          if (!cleanBase64.startsWith("iVBOR") && !cleanBase64.startsWith("/9j/")) continue;
          await insertImageIntoCurrentSlide(cleanBase64, {
            title: (transfer.metadata || {}).title || "Chart",
            from_app: transfer.from_app || "external",
          });
          appendMessage(
            "assistant",
            `📊 Chart from ${transfer.from_app || "another app"} inserted into current slide.`
          );
        } catch (e) {
          console.warn("[tsifl] PowerPoint auto-import failed:", e);
        }
      }
    } catch (_) { /* polling is best-effort */ }
  }, 4000);
});

/** Insert a base64 image into the currently-active slide (or slide 0 if none active). */
async function insertImageIntoCurrentSlide(base64, opts = {}) {
  await PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items");
    await ctx.sync();
    if (!slides.items.length) return;

    // PowerPoint JS API doesn't cleanly expose "current slide" selection, so
    // drop the image on the last slide (where the user was likely editing)
    // unless opts.slide_index is provided.
    const idx = opts.slide_index != null
      ? Math.min(opts.slide_index, slides.items.length - 1)
      : slides.items.length - 1;
    const slide = slides.items[idx];
    const image = slide.shapes.addImage(base64);
    image.left   = opts.left   ?? 50;
    image.top    = opts.top    ?? 100;
    image.width  = opts.width  ?? 600;
    image.height = opts.height ?? 400;
    if (opts.title) image.name = String(opts.title).substring(0, 60);
    await ctx.sync();
  });
}

// ── Auth Screens ────────────────────────────────────────────────────────────

function showLoginScreen() {
  document.getElementById("login-screen").style.display = "flex";
  document.getElementById("chat-screen").style.display = "none";

  document.getElementById("login-btn").onclick = handleSignIn;
  document.getElementById("signup-btn").onclick = handleSignUp;
  // Show/hide password (Improvement 9)
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

  // Image handling — file input overlaid on attach button (no programmatic .click() —
  // Office.js WKWebView blocks programmatic file input clicks on Mac)
  const imageInput = document.getElementById("image-input");
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
      addFile(file);
    }
  });

  // Clipboard paste
  document.getElementById("user-input").addEventListener("paste", (e) => {
    for (const item of (e.clipboardData || {}).items || []) {
      if (item.type.startsWith("image/") || item.kind === "file") {
        const f = item.getAsFile();
        if (f) addFile(f);
      }
    }
  });

  // Notes button (optional — may not exist in HTML)
  const notesBtn = document.getElementById("notes-btn");
  if (notesBtn) {
    notesBtn.onclick = () => {
      window.open(BACKEND_URL + "/notes-app", "_blank");
    };
  }

  // Templates dropdown (Improvement 45)
  const templatesBtn = document.getElementById("templates-btn");
  const templatesDd = document.getElementById("templates-dropdown");
  if (templatesBtn && templatesDd) {
    templatesBtn.onclick = (e) => {
      e.stopPropagation();
      templatesDd.style.display = templatesDd.style.display === "none" ? "block" : "none";
    };
    document.querySelectorAll(".template-item").forEach(item => {
      item.onclick = () => {
        document.getElementById("user-input").value = `Create a ${item.dataset.template} presentation with professional formatting and sample content`;
        templatesDd.style.display = "none";
        handleSubmit();
      };
    });
    document.addEventListener("click", () => { templatesDd.style.display = "none"; });
  }

  // Quick action buttons
  document.querySelectorAll(".quick-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const input = document.getElementById("user-input");
      input.value = btn.dataset.prompt;
      handleSubmit();
    });
  });

  // Auto-resize textarea
  const userInput = document.getElementById("user-input");
  userInput.addEventListener("input", () => {
    userInput.style.height = "auto";
    userInput.style.height = Math.min(userInput.scrollHeight, 100) + "px";
  });

  // Escape to clear input
  userInput.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      userInput.value = "";
      userInput.style.height = "auto";
    }
  });

  // Update slide count (Improvement 36)
  updateSlideCount();
  setStatus("Ready");
}

// Update slide count in header (Improvement 36)
async function updateSlideCount() {
  try {
    await PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load("items");
      await ctx.sync();
      const count = slides.items.length;
      const el = document.getElementById("slide-count");
      if (el) el.textContent = `${count} slide${count !== 1 ? "s" : ""}`;
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
  showTypingIndicator();

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
    hideTypingIndicator();
    appendMessage("assistant", result.reply);

    // Execute actions
    const actions = result.actions?.length ? result.actions : (result.action?.type ? [result.action] : []);
    if (actions.length > 0) {
      setStatus(`Applying ${actions.length} action${actions.length > 1 ? "s" : ""}...`);
      showTypingIndicator("applying");
      if (actions.length > 2) showProgress(0, actions.length);
      let success = 0, failed = 0;
      for (const action of actions) {
        try {
          await executeAction(action);
          success++;
          if (actions.length > 2) showProgress(success + failed, actions.length);
        } catch (e) {
          console.error("Action failed:", action.type, e);
          failed++;
        }
      }
      hideProgress();
      hideTypingIndicator();
      if (failed > 0) {
        appendMessage("assistant", `${success} actions applied, ${failed} failed. Try rephrasing your request.`);
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
    case "create_slide":
      await PowerPoint.run(async (ctx) => {
        const presentation = ctx.presentation;
        presentation.slides.add();
        await ctx.sync();

        const slides = presentation.slides;
        slides.load("items");
        await ctx.sync();

        const newSlide = slides.items[slides.items.length - 1];
        const layout = (payload.layout || "").toLowerCase();
        const isTitleSlide = layout.includes("title slide") || layout.includes("section");
        const bgColor = payload.background_color || (isTitleSlide ? "#0D5EAF" : "#F8FAFC");

        // Set slide background via native API
        try {
          newSlide.background.fill.setSolidColor(bgColor);
          await ctx.sync();
        } catch (e) {
          console.warn("[tsifl] create_slide: background API failed, skipping");
        }

        // Add accent bar — thicker for visual impact
        if (!isTitleSlide) {
          const accentBar = newSlide.shapes.addGeometricShape("Rectangle", {
            left: 0, top: 0, width: 720, height: 8,
          });
          accentBar.fill.setSolidColor(payload.accent_color || "#0D5EAF");
          await ctx.sync();
        }

        // Title text colors based on background
        const defaultTitleColor = isTitleSlide ? "#FFFFFF" : "#1E293B";
        const defaultContentColor = isTitleSlide ? "#E2E8F0" : "#334155";
        const defaultSubColor = isTitleSlide ? "#CBD5E1" : "#64748B";

        // Add title
        if (payload.title) {
          const titleBox = newSlide.shapes.addTextBox(payload.title);
          if (isTitleSlide) {
            titleBox.left = 50; titleBox.top = 150; titleBox.width = 620; titleBox.height = 80;
          } else {
            titleBox.left = 50; titleBox.top = 20; titleBox.width = 620; titleBox.height = 55;
          }
          await ctx.sync();
          titleBox.textFrame.textRange.load("font");
          await ctx.sync();
          const titleFont = titleBox.textFrame.textRange.font;
          titleFont.size = isTitleSlide ? 36 : 26;
          titleFont.bold = true;
          titleFont.color = payload.title_color || defaultTitleColor;
          titleFont.name = "Calibri";
          await ctx.sync();
        }

        // Add subtitle (title slides only)
        if (payload.subtitle && isTitleSlide) {
          const subBox = newSlide.shapes.addTextBox(payload.subtitle);
          subBox.left = 50; subBox.top = 240; subBox.width = 620; subBox.height = 40;
          await ctx.sync();
          subBox.textFrame.textRange.load("font");
          await ctx.sync();
          subBox.textFrame.textRange.font.size = 18;
          subBox.textFrame.textRange.font.color = payload.subtitle_color || defaultSubColor;
          subBox.textFrame.textRange.font.name = "Calibri";
          await ctx.sync();
        }

        // Add bottom line for title slides
        if (isTitleSlide) {
          const bottomLine = newSlide.shapes.addGeometricShape("Rectangle", {
            left: 50, top: 290, width: 100, height: 3,
          });
          bottomLine.fill.setSolidColor("#FFFFFF");
          await ctx.sync();
        }

        // Add content
        if (payload.content) {
          const contentBox = newSlide.shapes.addTextBox(payload.content);
          contentBox.left = 50;
          contentBox.top = isTitleSlide ? 310 : 85;
          contentBox.width = 620;
          contentBox.height = isTitleSlide ? 150 : 380;
          await ctx.sync();
          contentBox.textFrame.textRange.load("font");
          await ctx.sync();
          const contentFont = contentBox.textFrame.textRange.font;
          contentFont.size = payload.font_size || 18;
          contentFont.color = payload.content_color || defaultContentColor;
          contentFont.name = "Calibri";
          await ctx.sync();
        }
      });
      break;

    case "add_text_box":
      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();

        const idx = payload.slide_index != null ? payload.slide_index : 0;
        if (idx >= slides.items.length) {
          console.warn(`[tsifl] add_text_box: slide_index ${idx} out of range (${slides.items.length} slides)`);
          return;
        }
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

        const idx = payload.slide_index != null ? payload.slide_index : 0;
        if (idx >= slides.items.length) {
          console.warn(`[tsifl] add_shape: slide_index ${idx} out of range (${slides.items.length} slides)`);
          return;
        }
        const slide = slides.items[idx];

        // Map common shape names to valid Office.js GeometricShapeType values
        const shapeMap = {
          "rectangle": "Rectangle",
          "roundedrectangle": "RoundedRectangle",
          "oval": "Oval",
          "ellipse": "Oval",
          "circle": "Oval",
          "triangle": "Triangle",
          "arrow": "RightArrow",
          "rightarrow": "RightArrow",
          "callout": "Callout1",
          "diamond": "Diamond",
          "pentagon": "Pentagon",
          "hexagon": "Hexagon",
          "star": "Star5",
        };
        const rawType = (payload.shape_type || "Rectangle").replace(/\s+/g, "");
        const shapeType = shapeMap[rawType.toLowerCase()] || rawType;

        try {
          const shape = slide.shapes.addGeometricShape(shapeType, {
            left: payload.left || 50,
            top: payload.top || 50,
            width: payload.width || 200,
            height: payload.height || 100,
          });

          if (payload.fill_color) {
            shape.fill.setSolidColor(payload.fill_color);
          }
          if (payload.line_color) {
            shape.lineFormat.color = payload.line_color;
          }
          if (payload.text) {
            shape.textFrame.textRange.text = payload.text;
            // Style text in shapes
            shape.textFrame.textRange.load("font");
            await ctx.sync();
            const font = shape.textFrame.textRange.font;
            font.size = payload.font_size || 14;
            font.bold = payload.bold !== undefined ? payload.bold : true;
            font.color = payload.text_color || (payload.fill_color ? "#FFFFFF" : "#1E293B");
            font.name = "Calibri";
          }
          await ctx.sync();
        } catch (shapeErr) {
          console.error(`[tsifl] add_shape failed (type="${shapeType}"):`, shapeErr.message);
          // Fallback to Rectangle if the shape type isn't supported
          if (shapeType !== "Rectangle") {
            const shape = slide.shapes.addGeometricShape("Rectangle", {
              left: payload.left || 50,
              top: payload.top || 50,
              width: payload.width || 200,
              height: payload.height || 100,
            });
            if (payload.fill_color) shape.fill.setSolidColor(payload.fill_color);
            if (payload.text) shape.textFrame.textRange.text = payload.text;
            await ctx.sync();
          }
        }
      });
      break;

    case "add_image":
    case "import_image": {
      // Resolve the image bytes. Three sources supported, in this order:
      //   1. payload.image_data — explicit base64 (from server-side plot_service)
      //   2. payload.transfer_id — fetch from /transfer/<id>, used by cross-app
      //   3. payload.image_url   — URL (http/https or data: URI)
      // If NONE is provided, try /transfer/pending/powerpoint (one-shot claim).
      let base64 = null;
      let urlForAddImage = null;
      let titleFromTransfer = null;

      if (payload.image_data) {
        base64 = String(payload.image_data).replace(/^data:image\/[a-z+]+;base64,/, "");
      } else if (payload.transfer_id) {
        try {
          const tResp = await fetch(`${BACKEND_URL}/transfer/${payload.transfer_id}`);
          if (tResp.ok) {
            const transfer = await tResp.json();
            base64 = String(transfer.data || "").replace(/^data:image\/[a-z+]+;base64,/, "");
            titleFromTransfer = (transfer.metadata || {}).title || null;
          }
        } catch (e) {
          console.warn("[tsifl] transfer fetch failed:", e);
        }
      } else if (payload.image_url) {
        urlForAddImage = payload.image_url;
      } else {
        // No explicit source — claim the next pending image for PowerPoint
        try {
          const pResp = await fetch(`${BACKEND_URL}/transfer/pending/powerpoint`);
          if (pResp.ok) {
            const { pending = [] } = await pResp.json();
            const first = pending.find(p => p.data_type === "image");
            if (first) {
              const tResp = await fetch(`${BACKEND_URL}/transfer/${first.transfer_id}`);
              if (tResp.ok) {
                const transfer = await tResp.json();
                base64 = String(transfer.data || "").replace(/^data:image\/[a-z+]+;base64,/, "");
                titleFromTransfer = (transfer.metadata || {}).title || null;
              }
            }
          }
        } catch (e) {
          console.warn("[tsifl] pending-transfer lookup failed:", e);
        }
      }

      if (!base64 && !urlForAddImage) {
        console.warn("[tsifl] add_image/import_image: no image source available");
        break;
      }

      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();
        if (!slides.items.length) return;

        const idx = payload.slide_index != null
          ? Math.min(payload.slide_index, slides.items.length - 1)
          : slides.items.length - 1;  // default = last slide (where analyst is likely working)
        const slide = slides.items[idx];
        const src = base64 || urlForAddImage;
        const image = slide.shapes.addImage(src);
        image.left   = payload.left   ?? 50;
        image.top    = payload.top    ?? 100;
        image.width  = payload.width  ?? 600;
        image.height = payload.height ?? 400;
        const title = payload.title || titleFromTransfer;
        if (title) image.name = String(title).substring(0, 60);
        await ctx.sync();
      });
      break;
    }

    case "add_table":
      await PowerPoint.run(async (ctx) => {
        try {
          const slides = ctx.presentation.slides;
          slides.load("items");
          await ctx.sync();

          const idx = payload.slide_index != null ? payload.slide_index : 0;
          if (idx >= slides.items.length) {
            console.warn(`[tsifl] add_table: slide_index ${idx} out of range`);
            return;
          }
          const slide = slides.items[idx];

          // Ensure data is a valid 2D array of strings
          let data = payload.data || [];
          if (!Array.isArray(data) || data.length === 0) {
            console.warn("[tsifl] add_table: no data provided");
            return;
          }
          const maxCols = Math.max(...data.map(r => (Array.isArray(r) ? r.length : 0)));
          data = data.map(r => {
            if (!Array.isArray(r)) return Array(maxCols).fill("");
            while (r.length < maxCols) r.push("");
            return r.map(v => (v != null ? String(v) : ""));
          });

          // Cap table size for readability
          const rows = Math.min(data.length, 15);
          const columns = Math.min(maxCols, 8);
          const values = data.slice(0, rows).map(r => r.slice(0, columns));

          // Build per-cell styling: header row blue+white, alternating body rows
          const specificCellProperties = [];
          for (let r = 0; r < rows; r++) {
            const rowProps = [];
            for (let c = 0; c < columns; c++) {
              if (r === 0 && payload.header_row !== false) {
                rowProps.push({
                  fill: { color: "#0D5EAF" },
                  font: { bold: true, color: "#FFFFFF", name: "Calibri", size: 11 },
                });
              } else if (r % 2 === 1) {
                rowProps.push({
                  fill: { color: "#FFFFFF" },
                  font: { bold: false, color: "#1E293B", name: "Calibri", size: 10 },
                });
              } else {
                rowProps.push({
                  fill: { color: "#F1F5F9" },
                  font: { bold: false, color: "#1E293B", name: "Calibri", size: 10 },
                });
              }
            }
            specificCellProperties.push(rowProps);
          }

          // Create table with values AND styling in one call (correct PowerPoint JS API)
          slide.shapes.addTable(rows, columns, {
            left: payload.left || 50,
            top: payload.top || 90,
            width: payload.width || 620,
            height: payload.height || Math.min(rows * 30 + 20, 380),
            values: values,
            specificCellProperties: specificCellProperties,
          });
          await ctx.sync();
        } catch (tableErr) {
          console.error("[tsifl] add_table failed:", tableErr.message);
          // Fallback: try without styling (in case API version doesn't support specificCellProperties)
          try {
            const slides2 = ctx.presentation.slides;
            slides2.load("items");
            await ctx.sync();
            const idx2 = payload.slide_index != null ? payload.slide_index : 0;
            if (idx2 < slides2.items.length) {
              const slide2 = slides2.items[idx2];
              let data2 = payload.data || [];
              const maxCols2 = Math.max(...data2.map(r => (Array.isArray(r) ? r.length : 0)));
              data2 = data2.map(r => {
                if (!Array.isArray(r)) return Array(maxCols2).fill("");
                while (r.length < maxCols2) r.push("");
                return r.map(v => (v != null ? String(v) : ""));
              });
              const rows2 = Math.min(data2.length, 15);
              const cols2 = Math.min(maxCols2, 8);
              const vals2 = data2.slice(0, rows2).map(r => r.slice(0, cols2));

              slide2.shapes.addTable(rows2, cols2, {
                left: payload.left || 50,
                top: payload.top || 90,
                width: payload.width || 620,
                height: payload.height || Math.min(rows2 * 30 + 20, 380),
                values: vals2,
              });
              await ctx.sync();
            }
          } catch (fallbackErr) {
            console.error("[tsifl] add_table fallback also failed:", fallbackErr.message);
          }
        }
      });
      break;

    case "add_chart":
      // PowerPoint chart creation via Office.js — uses table with values passed at creation
      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();

        const idx = payload.slide_index || 0;
        if (idx >= slides.items.length) return;
        const slide = slides.items[idx];

        if (payload.data && payload.data.length > 0) {
          const rows = Math.min(payload.data.length, 15);
          const columns = Math.min(payload.data[0].length, 8);
          const values = payload.data.slice(0, rows).map(r =>
            r.slice(0, columns).map(v => (v != null ? String(v) : ""))
          );

          // Build styling: first row as header
          const specificCellProperties = [];
          for (let r = 0; r < rows; r++) {
            const rowProps = [];
            for (let c = 0; c < columns; c++) {
              if (r === 0) {
                rowProps.push({
                  fill: { color: "#0D5EAF" },
                  font: { bold: true, color: "#FFFFFF", name: "Calibri", size: 10 },
                });
              } else {
                rowProps.push({
                  fill: { color: r % 2 === 1 ? "#FFFFFF" : "#F1F5F9" },
                  font: { bold: false, color: "#1E293B", name: "Calibri", size: 9 },
                });
              }
            }
            specificCellProperties.push(rowProps);
          }

          slide.shapes.addTable(rows, columns, {
            left: payload.left || 50,
            top: payload.top || 120,
            width: payload.width || 620,
            height: payload.height || 280,
            values: values,
            specificCellProperties: specificCellProperties,
          });
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

        const idx = payload.slide_index != null ? payload.slide_index : 0;
        if (idx >= slides.items.length) {
          console.warn(`[tsifl] set_slide_background: slide_index ${idx} out of range (${slides.items.length} slides)`);
          return;
        }
        const slide = slides.items[idx];

        if (payload.color) {
          // The create_slide handler already applies accent bars and backgrounds.
          // For additional background color: add a rectangle FIRST, then move other shapes on top.
          // But this is risky — instead, just skip if the API isn't available.
          try {
            slide.background.fill.setSolidColor(payload.color);
            await ctx.sync();
          } catch (bgErr) {
            console.warn(`[tsifl] set_slide_background: native API failed (${bgErr.message}). Skipping — create_slide already handles styling.`);
            // Don't add a covering rectangle — it hides content.
            // The accent bar from create_slide provides visual distinction.
          }
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
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();

        const fontName = payload.font_scheme || payload.font_name || "Times New Roman";

        for (let s = 0; s < slides.items.length; s++) {
          const slide = slides.items[s];
          const shapes = slide.shapes;
          shapes.load("items");
          await ctx.sync();

          for (let i = 0; i < shapes.items.length; i++) {
            const shape = shapes.items[i];
            try {
              // Check if shape has a textFrame (text boxes, titles, etc.)
              const tf = shape.textFrame;
              if (!tf) continue;
              const textRange = tf.textRange;
              textRange.load("text");
              await ctx.sync();

              // Apply font to entire text range
              textRange.font.name = fontName;
              if (payload.font_size) textRange.font.size = payload.font_size;
              if (payload.font_color) textRange.font.color = payload.font_color;
            } catch (e) {
              // Shape doesn't have text — skip (tables, images, etc.)
            }
          }
        }
        await ctx.sync();
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

// ── Thinking bubble with rotating punchlines ─────────────────────────────────
const _thinkingMessages = {
  thinking: [
    'Reading your request...',
    'Staring at a blank slide so you don\'t have to...',
    'Channeling my inner McKinsey associate...',
    'Picking the perfect shade of corporate blue...',
    'Aligning boxes with pixel-perfect precision...',
    'Wondering if this needs a waterfall chart...',
    'Making sure no slide says "click to add title"...',
    'Designing slides that would survive a board meeting...',
    'Every bullet point is a work of art...',
    'Goldman would bill $200/slide for this...',
    'Bezos banned PowerPoint — let\'s prove him wrong...',
    'Your MD is going to think you made this yourself...',
    'Turning data into something a VP can actually read...',
    'One does not simply make an ugly deck...',
    'If this deck doesn\'t get funded, nothing will...',
    'Adding the kind of detail that gets you promoted...',
    'This is giving "closing the deal" energy...',
  ],
  applying: [
    'Laying down slides like a card dealer in Monte Carlo...',
    'Your deck is getting the glow-up it deserves...',
    'Placing shapes with surgical precision...',
    'Tables are populating... beautifully...',
    'Every font choice is a statement...',
    'Morgan Stanley\'s design team is taking notes...',
    'Building the kind of deck that ends meetings early...',
    'Slides materializing in 3... 2... 1...',
    'KPI cards locking into position...',
    'This is the part where your deck becomes legendary...',
  ]
};

let _thinkingInterval = null;
let _thinkingPhase = null;

function showTypingIndicator(phase) {
  hideTypingIndicator();
  _thinkingPhase = phase || 'thinking';
  const history = document.getElementById("chat-history");
  const div = document.createElement("div");
  div.id = "typing-indicator";
  div.className = "thinking-bubble";
  div.innerHTML = '<div class="thinking-orb"></div><span class="thinking-text"></span>';
  history.appendChild(div);
  history.scrollTop = history.scrollHeight;

  const msgs = _thinkingMessages[_thinkingPhase] || _thinkingMessages.thinking;
  let idx = 0;

  function showNext() {
    const el = document.querySelector('#typing-indicator .thinking-text');
    if (!el) return;
    const msg = msgs[idx % msgs.length];
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

function showProgress(current, total) {
  const wrap = document.getElementById("progress-bar-wrap");
  const bar = document.getElementById("progress-bar");
  const text = document.getElementById("progress-text");
  if (!wrap || !bar) return;
  wrap.style.display = "block";
  const pct = Math.round((current / total) * 100);
  bar.style.width = pct + "%";
  if (text) text.textContent = `${current} / ${total}`;
}

function hideProgress() {
  const wrap = document.getElementById("progress-bar-wrap");
  if (wrap) wrap.style.display = "none";
  const bar = document.getElementById("progress-bar");
  if (bar) bar.style.width = "0%";
}
