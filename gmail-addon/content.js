/**
 * tsifl — Content Script
 * Captures page context, executes DOM actions, and provides full page text for summarization.
 * Auto-injected on all http/https pages via manifest content_scripts.
 */

if (!window.__tsifl_content_loaded) {
  window.__tsifl_content_loaded = true;

  function detectSite() {
    const host = location.hostname;
    const path = location.pathname;
    if (host === "mail.google.com") return "gmail";
    if (host === "docs.google.com" && path.startsWith("/spreadsheets")) return "google_sheets";
    if (host === "docs.google.com" && path.startsWith("/document")) return "google_docs";
    if (host === "docs.google.com" && path.startsWith("/presentation")) return "google_slides";
    return "browser";
  }

  function getContext() {
    const site = detectSite();
    const base = { app: site, url: location.href, title: document.title };
    const selection = window.getSelection()?.toString()?.trim() || "";
    if (selection) base.selection = selection;

    switch (site) {
      case "gmail": return { ...base, ...getGmailContext() };
      case "google_sheets": return { ...base, ...getGoogleSheetsContext() };
      case "google_docs": return { ...base, ...getGoogleDocsContext() };
      case "google_slides": return { ...base, ...getGoogleSlidesContext() };
      default: return { ...base, ...getBrowserContext() };
    }
  }

  function getGmailContext() {
    const ctx = {};
    const subject = document.querySelector('h2[data-thread-perm-id]')?.textContent?.trim();
    if (subject) ctx.thread_subject = subject;
    const msgs = document.querySelectorAll('[data-message-id]');
    if (msgs.length > 0) {
      ctx.messages = Array.from(msgs).slice(-5).map(m => {
        const sender = m.querySelector('[email]')?.getAttribute('email') || "";
        const text = m.querySelector('[data-message-id] .a3s')?.textContent?.trim()?.slice(0, 500) || "";
        return { sender, snippet: text };
      });
    }
    return ctx;
  }

  function getGoogleSheetsContext() {
    const ctx = {};
    const title = document.querySelector('.docs-title-input')?.value || document.title;
    ctx.sheet_title = title;
    const cellInput = document.querySelector('#t-formula-bar-input .cell-input');
    if (cellInput) ctx.formula_bar = cellInput.textContent?.trim()?.slice(0, 500);
    const tabs = document.querySelectorAll('.docs-sheet-tab .docs-sheet-tab-name');
    if (tabs.length) ctx.sheet_tabs = Array.from(tabs).map(t => t.textContent?.trim());
    return ctx;
  }

  function getGoogleDocsContext() {
    const ctx = {};
    const title = document.querySelector('.docs-title-input')?.value || document.title;
    ctx.doc_title = title;
    const editor = document.querySelector('.kix-appview-editor');
    if (editor) ctx.doc_content = editor.textContent?.trim()?.slice(0, 3000);
    return ctx;
  }

  function getGoogleSlidesContext() {
    const ctx = {};
    const slides = document.querySelectorAll('.punch-filmstrip-thumbnail');
    ctx.slide_count = slides.length || 0;
    const current = document.querySelector('.punch-viewer-svgpage-svgcontainer');
    if (current) ctx.current_slide_text = current.textContent?.trim()?.slice(0, 2000);
    return ctx;
  }

  function getBrowserContext() {
    const ctx = {};
    const meta = document.querySelector('meta[name="description"]')?.content;
    if (meta) ctx.meta_description = meta;
    // Get structured page content for context
    const main = document.querySelector('main, article, [role="main"]');
    ctx.page_text = (main || document.body).textContent?.trim()?.replace(/\s+/g, ' ')?.slice(0, 3000);
    return ctx;
  }

  // Full page text extraction for summarization
  function getFullPageText() {
    // Try to get the most meaningful content
    const selectors = [
      'article',
      'main',
      '[role="main"]',
      '.post-content',
      '.article-content',
      '.entry-content',
      '.content',
      '#content',
      '.story-body',
      '.article-body',
    ];

    let contentEl = null;
    for (const sel of selectors) {
      contentEl = document.querySelector(sel);
      if (contentEl && contentEl.textContent.trim().length > 200) break;
    }

    if (!contentEl) contentEl = document.body;

    // Extract text with some structure preserved
    const text = extractStructuredText(contentEl);
    return text.slice(0, 15000); // Up to 15K chars for summarization
  }

  function extractStructuredText(el) {
    const parts = [];
    const blocks = el.querySelectorAll('h1, h2, h3, h4, h5, h6, p, li, blockquote, pre, td, th, figcaption');

    if (blocks.length > 0) {
      for (const block of blocks) {
        const tag = block.tagName.toLowerCase();
        const text = block.textContent?.trim();
        if (!text) continue;

        if (tag.startsWith('h')) {
          parts.push(`\n## ${text}\n`);
        } else if (tag === 'li') {
          parts.push(`- ${text}`);
        } else if (tag === 'blockquote') {
          parts.push(`> ${text}`);
        } else {
          parts.push(text);
        }
      }
    }

    // Fallback to raw text if structured extraction failed
    if (parts.length < 3) {
      return el.textContent?.trim()?.replace(/\s+/g, ' ') || "";
    }

    return parts.join('\n');
  }

  // Handle DOM actions
  function handleDomAction(type, payload) {
    switch (type) {
      case "scroll_to": {
        if (payload.selector) {
          const el = document.querySelector(payload.selector);
          if (el) { el.scrollIntoView({ behavior: "smooth", block: "center" }); return { success: true, message: "Scrolled" }; }
          return { success: false, message: `Element not found: ${payload.selector}` };
        }
        window.scrollTo({ top: payload.y || 0, behavior: "smooth" });
        return { success: true, message: "Scrolled" };
      }
      case "click_element": {
        const el = document.querySelector(payload.selector);
        if (el) { el.click(); return { success: true, message: `Clicked ${payload.selector}` }; }
        return { success: false, message: `Element not found: ${payload.selector}` };
      }
      case "fill_input": {
        const input = document.querySelector(payload.selector);
        if (input) {
          input.focus();
          input.value = payload.value;
          input.dispatchEvent(new Event("input", { bubbles: true }));
          input.dispatchEvent(new Event("change", { bubbles: true }));
          return { success: true, message: `Filled ${payload.selector}` };
        }
        return { success: false, message: `Input not found: ${payload.selector}` };
      }
      case "extract_text": {
        const target = payload.selector ? document.querySelector(payload.selector) : document.body;
        const text = target?.textContent?.trim()?.slice(0, 5000) || "";
        return { success: true, message: text, text };
      }
      default:
        return { success: false, message: `Unknown DOM action: ${type}` };
    }
  }

  // Listen for messages
  chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
    if (msg.action === "capture_context") {
      sendResponse(getContext());
      return;
    }
    if (msg.action === "execute_dom_action") {
      sendResponse(handleDomAction(msg.type, msg.payload));
      return;
    }
    if (msg.action === "get_full_page_text") {
      sendResponse({ text: getFullPageText() });
      return;
    }
  });
}
