/**
 * tsifl — Content Script
 * Captures page context, executes DOM actions, and provides full page text for summarization.
 * Auto-injected on all http/https pages via manifest content_scripts.
 */

// Versioned guard — allows new code to replace old on extension updates.
// Uses `var` (not const) so re-injection doesn't throw SyntaxError.
var TSIFL_CONTENT_VERSION = 3;
if (window.__tsifl_content_version >= TSIFL_CONTENT_VERSION) {
  // Already running this version or newer — skip
} else {
  window.__tsifl_content_version = TSIFL_CONTENT_VERSION;

  // Remove previous message listener so we don't stack duplicates
  if (window.__tsifl_msg_handler) {
    try { chrome.runtime.onMessage.removeListener(window.__tsifl_msg_handler); } catch (e) {}
  }

  function detectSite() {
    const host = location.hostname;
    const path = location.pathname;
    if (host === "mail.google.com") return "gmail";
    if (host === "docs.google.com" && path.startsWith("/spreadsheets")) return "google_sheets";
    if (host === "docs.google.com" && path.startsWith("/document")) return "google_docs";
    if (host === "docs.google.com" && path.startsWith("/presentation")) return "google_slides";
    // Enhanced page type detection (Improvement 75)
    if (host.includes("amazon.") || host.includes("amzn.")) return "browser";
    if (host.includes("linkedin.com")) return "browser";
    if (host.includes("stackoverflow.com") || host.includes("stackexchange.com")) return "browser";
    if (host.includes("wikipedia.org")) return "browser";
    if (host.includes("github.com")) return "browser";
    return "browser";
  }

  // Detect page type for better context (Improvement 75)
  function detectPageType() {
    const host = location.hostname;
    if (host.includes("amazon.") || host.includes("amzn.")) return "product";
    if (host.includes("linkedin.com") && location.pathname.includes("/in/")) return "linkedin_profile";
    if (host.includes("linkedin.com") && location.pathname.includes("/jobs/")) return "linkedin_job";
    if (host.includes("stackoverflow.com") && location.pathname.includes("/questions/")) return "stackoverflow";
    if (host.includes("wikipedia.org")) return "wikipedia";
    if (host.includes("github.com")) return "github";
    if (host.includes("bloomberg.com") || host.includes("reuters.com") || host.includes("sec.gov")) return "financial";
    if (document.querySelector("article") || document.querySelector('[role="article"]')) return "article";
    if (document.querySelectorAll("[data-testid='search-result'], .g, .search-result").length > 2) return "search_results";
    return "general";
  }

  function getContext() {
    const site = detectSite();
    const base = {
      app: site,
      url: location.href,
      title: document.title,
      page_type: detectPageType()
    };
    const selection = window.getSelection()?.toString()?.trim() || "";
    if (selection) base.selection = selection;
    // Meta description
    const metaDesc = document.querySelector('meta[name="description"]')?.content;
    if (metaDesc) base.meta_description = metaDesc;

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

    // Extract thread subject — try multiple selectors for reliability
    const subjectSelectors = [
      'h2[data-thread-perm-id]',
      'h2.hP',
      '.ha h2',
      'h2[data-legacy-thread-id]',
      '.nH .hP',
      '[role="main"] h2',
    ];
    for (const sel of subjectSelectors) {
      const el = document.querySelector(sel);
      if (el?.textContent?.trim()) {
        ctx.thread_subject = el.textContent.trim();
        break;
      }
    }

    // Detect compose mode
    const composeWindows = document.querySelectorAll('.AD, .nH .aoP, div[role="dialog"] .aO7, .dw .editable');
    ctx.is_composing = composeWindows.length > 0;

    // Count unread emails in inbox if visible
    const unreadBadge = document.querySelector('.aim .bsU') || document.querySelector('.aio .bsU');
    if (unreadBadge?.textContent?.trim()) {
      const count = parseInt(unreadBadge.textContent.trim().replace(/[^0-9]/g, ''), 10);
      if (!isNaN(count)) ctx.unread_count = count;
    }

    // Extract messages with full body, sender name/email, and date
    const msgs = document.querySelectorAll('[data-message-id]');
    if (msgs.length > 0) {
      ctx.message_count = msgs.length;
      ctx.messages = Array.from(msgs).slice(-5).map(m => {
        const emailAttr = m.querySelector('[email]')?.getAttribute('email') || "";
        const senderName = m.querySelector('[email]')?.getAttribute('name')
          || m.querySelector('.gD')?.textContent?.trim()
          || m.querySelector('[email]')?.textContent?.trim()
          || "";
        const bodyEl = m.querySelector('.a3s.aiL') || m.querySelector('.a3s') || m.querySelector('[data-message-id] .a3s');
        const body = bodyEl?.textContent?.trim()?.slice(0, 3000) || "";
        // Extract date/time from message header
        const dateEl = m.querySelector('.g3') || m.querySelector('.gH .g3') || m.querySelector('[title]');
        const date = dateEl?.getAttribute('title') || dateEl?.textContent?.trim() || "";
        return { sender_name: senderName, sender_email: emailAttr, snippet: body, date };
      });
    }

    return ctx;
  }

  function getGoogleSheetsContext() {
    const ctx = {};
    const title = document.querySelector('.docs-title-input')?.value || document.title;
    ctx.sheet_title = title;
    const idMatch = location.pathname.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (idMatch) ctx.document_id = idMatch[1];
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
    const idMatch = location.pathname.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (idMatch) ctx.document_id = idMatch[1];
    const editor = document.querySelector('.kix-appview-editor');
    if (editor) ctx.doc_content = editor.textContent?.trim()?.slice(0, 3000);
    return ctx;
  }

  function getGoogleSlidesContext() {
    const ctx = {};
    const idMatch = location.pathname.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (idMatch) ctx.document_id = idMatch[1];
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

    // Extract tables as structured data (headers + first 10 rows)
    const tables = document.querySelectorAll('table');
    if (tables.length > 0) {
      ctx.tables = [];
      for (let t = 0; t < Math.min(tables.length, 3); t++) {
        const table = tables[t];
        const headers = Array.from(table.querySelectorAll('th')).map(th => th.textContent?.trim() || "");
        const rows = [];
        const trs = table.querySelectorAll('tbody tr, tr');
        for (let r = 0; r < Math.min(trs.length, 10); r++) {
          const cells = Array.from(trs[r].querySelectorAll('td')).map(td => td.textContent?.trim()?.slice(0, 200) || "");
          if (cells.length > 0) rows.push(cells);
        }
        if (headers.length > 0 || rows.length > 0) {
          ctx.tables.push({ headers, rows, row_count: trs.length });
        }
      }
      if (ctx.tables.length === 0) delete ctx.tables;
    }

    // Extract form fields for form-filling assistance
    const inputs = document.querySelectorAll('input:not([type="hidden"]), select, textarea');
    if (inputs.length > 0 && inputs.length < 50) {
      ctx.form_fields = Array.from(inputs).slice(0, 20).map(el => ({
        tag: el.tagName.toLowerCase(),
        type: el.type || "",
        name: el.name || el.id || el.getAttribute('aria-label') || "",
        value: el.value?.slice(0, 100) || "",
        placeholder: el.placeholder || "",
      })).filter(f => f.name || f.placeholder);
    }

    // Extract links with text for navigation help
    const links = document.querySelectorAll('a[href]');
    if (links.length > 0) {
      const seen = new Set();
      ctx.links = [];
      for (const link of links) {
        const text = link.textContent?.trim()?.slice(0, 80);
        const href = link.href;
        if (text && href && !seen.has(href) && !href.startsWith('javascript:') && text.length > 2) {
          seen.add(href);
          ctx.links.push({ text, href });
          if (ctx.links.length >= 20) break;
        }
      }
      if (ctx.links.length === 0) delete ctx.links;
    }

    // Extract structured data (JSON-LD, Open Graph, meta tags)
    const structuredData = {};
    const jsonLd = document.querySelector('script[type="application/ld+json"]');
    if (jsonLd) {
      try {
        const parsed = JSON.parse(jsonLd.textContent);
        structuredData.json_ld = typeof parsed === 'object' ? parsed : null;
      } catch (e) { /* skip malformed JSON-LD */ }
    }
    const ogTags = document.querySelectorAll('meta[property^="og:"]');
    if (ogTags.length > 0) {
      structuredData.open_graph = {};
      ogTags.forEach(tag => {
        const prop = tag.getAttribute('property').replace('og:', '');
        structuredData.open_graph[prop] = tag.content?.slice(0, 200) || "";
      });
    }
    if (Object.keys(structuredData).length > 0) ctx.structured_data = structuredData;

    // Google Sheets: try to capture visible cell data
    if (location.hostname === 'docs.google.com' && location.pathname.startsWith('/spreadsheets')) {
      const cells = document.querySelectorAll('.cell-input, .softmerge-inner');
      if (cells.length > 0) {
        ctx.sheet_cells = Array.from(cells).slice(0, 50).map(c => c.textContent?.trim()).filter(Boolean);
      }
    }

    // Google Docs: capture more document content
    if (location.hostname === 'docs.google.com' && location.pathname.startsWith('/document')) {
      const editor = document.querySelector('.kix-appview-editor');
      if (editor) {
        ctx.doc_content = editor.textContent?.trim()?.slice(0, 5000);
      }
    }

    // Product context for e-commerce sites
    const product = getProductContext();
    if (product) ctx.product = product;

    return ctx;
  }

  function getProductContext() {
    const host = location.hostname;

    // Amazon
    if (host.includes('amazon.') || host.includes('amzn.')) {
      const name = document.querySelector('#productTitle, #title')?.textContent?.trim();
      if (!name) return null;
      const priceEl = document.querySelector('.a-price .a-offscreen, #priceblock_ourprice, #priceblock_dealprice, .a-price-whole');
      const price = priceEl?.textContent?.trim() || "";
      const ratingEl = document.querySelector('#acrPopover .a-icon-alt, [data-asin] .a-icon-alt');
      const rating = ratingEl?.textContent?.trim()?.match(/[\d.]+/)?.[0] || "";
      const reviewsEl = document.querySelector('#acrCustomerReviewText');
      const reviews = reviewsEl?.textContent?.trim()?.match(/[\d,]+/)?.[0] || "";
      return { name, price, rating, reviews_count: reviews, source: "amazon" };
    }

    // eBay
    if (host.includes('ebay.')) {
      const name = document.querySelector('h1.x-item-title__mainTitle span, h1[itemprop="name"]')?.textContent?.trim();
      if (!name) return null;
      const price = document.querySelector('.x-price-primary span, [itemprop="price"]')?.textContent?.trim() || "";
      return { name, price, source: "ebay" };
    }

    // Generic product page (schema.org or common patterns)
    const schemaName = document.querySelector('[itemprop="name"]')?.textContent?.trim();
    const schemaPrice = document.querySelector('[itemprop="price"]')?.textContent?.trim()
      || document.querySelector('[itemprop="price"]')?.getAttribute('content');
    const schemaRating = document.querySelector('[itemprop="ratingValue"]')?.textContent?.trim()
      || document.querySelector('[itemprop="ratingValue"]')?.getAttribute('content');
    const schemaReviews = document.querySelector('[itemprop="reviewCount"]')?.textContent?.trim()
      || document.querySelector('[itemprop="reviewCount"]')?.getAttribute('content');

    if (schemaName && schemaPrice) {
      return {
        name: schemaName.slice(0, 200),
        price: schemaPrice,
        rating: schemaRating || "",
        reviews_count: schemaReviews || "",
        source: "generic",
      };
    }

    return null;
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

  // ── Google Workspace Action Handlers ────────────────────────────────────
  // DOM automation for Google Docs/Sheets/Slides editing.
  // These simulate user interactions since direct API access isn't available
  // from a content script.

  function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

  function getGDocsEventTarget() {
    // Google Docs listens for key events on an iframe's body
    const iframe = document.querySelector('.docs-texteventtarget-iframe');
    if (iframe && iframe.contentDocument) return iframe.contentDocument.body;
    return null;
  }

  function dispatchKey(target, key, opts = {}) {
    const base = { key, bubbles: true, cancelable: true, ...opts };
    target.dispatchEvent(new KeyboardEvent('keydown', base));
    target.dispatchEvent(new KeyboardEvent('keyup', base));
  }

  // ── Google Docs/Sheets Helpers ──────────────────────────────────────────
  // Google Docs uses a proprietary canvas editor that rejects synthetic
  // keyboard events (isTrusted === false). We use window.find() for text
  // selection and real DOM clicks on toolbar buttons for formatting.

  /** Fill a native input field reliably (set value + dispatch events) */
  function fillInput(input, value) {
    const nativeSetter = Object.getOwnPropertyDescriptor(
      window.HTMLInputElement.prototype, 'value'
    )?.set;
    if (nativeSetter) nativeSetter.call(input, value);
    else input.value = value;
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
  }

  /** Click a toolbar button by aria-label substring. Returns true if found. */
  function clickToolbarBtn(labelSubstring) {
    // Try multiple selector patterns used by Google Workspace apps
    const variations = [labelSubstring, labelSubstring.toLowerCase(), labelSubstring.charAt(0).toUpperCase() + labelSubstring.slice(1)];
    for (const label of variations) {
      for (const attr of ['aria-label', 'data-tooltip']) {
        const btn = document.querySelector(`[${attr}*="${label}"]`);
        if (btn) {
          console.log(`[tsifl] clicking toolbar btn: ${attr}="${btn.getAttribute(attr)}"`);
          btn.click();
          return true;
        }
      }
    }
    return false;
  }

  /**
   * Select text in Google Docs using window.find().
   * Google Docs renders text as real DOM nodes inside .kix-lineview spans,
   * so window.find() can locate and select them. The browser selection is
   * then recognized by Docs when toolbar buttons are clicked.
   */
  function selectTextInPage(term) {
    // Clear any existing selection first
    window.getSelection()?.removeAllRanges();
    // window.find(string, caseSensitive, backwards, wrapAround)
    const found = window.find(term, false, false, true);
    console.log(`[tsifl] window.find("${term}") = ${found}`);
    return found;
  }

  async function handleWorkspaceAction(type, payload) {
    const site = detectSite();
    console.log("[tsifl] handleWorkspaceAction:", site, type, JSON.stringify(payload));

    // ── Google Docs Actions ─────────────────────────────────────────────
    if (site === "google_docs") {
      switch (type) {
        case "find_and_replace": {
          // Use Ctrl+H shortcut — we dispatch it on the document to trigger
          // the native Google Docs Find & Replace dialog
          // Since dispatchKey may be blocked, try clicking Edit menu instead
          const editMenu = document.getElementById('docs-edit-menu');
          if (editMenu) {
            editMenu.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
            editMenu.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
            editMenu.click();
            await sleep(400);
            // Find "Find and replace" menu item
            const menuItems = document.querySelectorAll('[role="menuitem"], .goog-menuitem');
            const frItem = Array.from(menuItems).find(el =>
              el.textContent.trim().toLowerCase().includes('find and replace')
            );
            if (frItem) {
              frItem.click();
              await sleep(500);
              // Fill dialog
              const inputs = document.querySelectorAll('input[type="text"]');
              const dialogInputs = Array.from(inputs).filter(i =>
                i.closest('.docs-findandreplacedialog') || i.closest('[role="dialog"]')
              );
              if (dialogInputs.length >= 2) {
                fillInput(dialogInputs[0], payload.find_text || '');
                await sleep(100);
                fillInput(dialogInputs[1], payload.replace_text || '');
                await sleep(200);
                const replaceAllBtn = Array.from(document.querySelectorAll('button')).find(b =>
                  b.textContent.trim().toLowerCase().includes('replace all')
                );
                if (replaceAllBtn) {
                  replaceAllBtn.click();
                  await sleep(300);
                  // Close dialog
                  const closeBtn = document.querySelector('[aria-label="Close"] , .modal-dialog-title-close');
                  if (closeBtn) closeBtn.click();
                  return { success: true, message: "Find and replace completed" };
                }
              }
            }
          }
          return { success: false, message: "Could not open Find & Replace in Google Docs" };
        }

        case "format_text": {
          const term = payload.range_description || '';
          if (!term) return { success: false, message: "No text specified" };

          console.log("[tsifl] format_text: selecting text with window.find:", term);

          // Step 1: Select text using window.find()
          // Google Docs renders text as real DOM nodes in .kix-lineview spans.
          // window.find() creates a browser Selection that Docs recognizes
          // when toolbar buttons are subsequently clicked.
          const found = selectTextInPage(term);
          if (!found) {
            return { success: false, message: `Could not find "${term}" in the document.` };
          }
          await sleep(200);

          // Step 2: Apply formatting via toolbar button clicks
          let formatted = false;

          if (payload.bold) {
            if (clickToolbarBtn('Bold')) formatted = true;
            await sleep(150);
          }
          if (payload.italic) {
            if (clickToolbarBtn('Italic')) formatted = true;
            await sleep(150);
          }
          if (payload.underline) {
            if (clickToolbarBtn('Underline')) formatted = true;
            await sleep(150);
          }

          if (payload.highlight_color) {
            const colorName = payload.highlight_color.toLowerCase();
            // Click highlight color button (need to open dropdown first)
            const hlBtn = document.querySelector('[aria-label*="Highlight color"]') ||
                         document.querySelector('[data-tooltip*="Highlight color"]') ||
                         document.querySelector('[aria-label*="highlight"]');
            console.log("[tsifl] highlight btn:", hlBtn?.getAttribute('aria-label') || 'NOT FOUND');
            if (hlBtn) {
              // Click the dropdown arrow to open color picker
              const arrow = hlBtn.querySelector('[class*="dropdown"]') || hlBtn;
              arrow.click();
              await sleep(400);
              // Search all visible elements for color match
              const allColorCells = document.querySelectorAll(
                '[data-color], [aria-label*="color" i], [style*="background"]'
              );
              let colorClicked = false;
              for (const cell of allColorCells) {
                const label = (cell.getAttribute('aria-label') || cell.getAttribute('title') || cell.getAttribute('data-tooltip') || '').toLowerCase();
                if (label.includes(colorName)) {
                  console.log("[tsifl] clicking color:", label);
                  cell.click();
                  colorClicked = true;
                  break;
                }
              }
              if (!colorClicked) {
                // Just click the main highlight button (applies last used color)
                hlBtn.click();
              }
              formatted = true;
              await sleep(150);
            }
          }

          if (payload.font_color) {
            const colorBtn = document.querySelector('[aria-label*="Text color"]') ||
                            document.querySelector('[data-tooltip*="Text color"]');
            if (colorBtn) {
              const arrow = colorBtn.querySelector('[class*="dropdown"]') || colorBtn;
              arrow.click();
              await sleep(400);
              const colorName = payload.font_color.toLowerCase();
              const cells = document.querySelectorAll('[data-color], [aria-label*="color" i]');
              for (const cell of cells) {
                const label = (cell.getAttribute('aria-label') || cell.getAttribute('title') || '').toLowerCase();
                if (label.includes(colorName)) { cell.click(); formatted = true; break; }
              }
              if (!formatted) { colorBtn.click(); formatted = true; }
              await sleep(150);
            }
          }

          if (formatted) {
            return { success: true, message: `Formatted "${term}"` };
          }
          // Text was found/selected but no toolbar buttons matched
          return { success: true, message: `Found and selected "${term}". Toolbar buttons not found — you can format manually.` };
        }

        case "insert_text": {
          const target = getGDocsEventTarget();
          if (target) {
            target.focus();
            document.execCommand('insertText', false, payload.text || '');
            return { success: true, message: "Text inserted" };
          }
          return { success: false, message: "Could not find Docs editor" };
        }

        default:
          return { success: false, message: `Google Docs action "${type}" is not yet supported via Chrome extension.` };
      }
    }

    // ── Google Sheets Actions ───────────────────────────────────────────
    if (site === "google_sheets") {
      switch (type) {
        case "write_cell": {
          const cell = payload.cell || 'A1';
          const value = payload.formula || payload.value || '';

          // Navigate to cell using the Name Box
          const nameBox = document.querySelector('#t-name-box input') ||
                         document.querySelector('.waffle-name-box input') ||
                         document.querySelector('[aria-label="Name Box"]');
          if (nameBox) {
            nameBox.click();
            await sleep(100);
            nameBox.focus();
            nameBox.value = cell;
            nameBox.dispatchEvent(new Event('input', { bubbles: true }));
            nameBox.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', keyCode: 13, bubbles: true }));
            await sleep(300);

            // Type the value into the cell
            const editor = document.querySelector('.cell-input') ||
                          document.querySelector('[contenteditable="true"]') ||
                          document.activeElement;
            if (editor) {
              editor.focus();
              // Clear existing content
              document.execCommand('selectAll', false, null);
              document.execCommand('insertText', false, value.toString());
              // Press Enter to commit
              editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', keyCode: 13, bubbles: true }));
              await sleep(200);
              return { success: true, message: `Wrote "${value}" to ${cell}` };
            }
            return { success: false, message: "Could not find cell editor" };
          }
          return { success: false, message: "Could not find Name Box" };
        }

        case "navigate_sheet": {
          const sheetName = payload.sheet || '';
          const tabs = document.querySelectorAll('.docs-sheet-tab');
          for (const tab of tabs) {
            const name = tab.querySelector('.docs-sheet-tab-name');
            if (name && name.textContent.trim() === sheetName) {
              tab.click();
              return { success: true, message: `Navigated to ${sheetName}` };
            }
          }
          return { success: false, message: `Sheet "${sheetName}" not found` };
        }

        case "find_and_replace": {
          // Open Find & Replace via Edit menu (Sheets uses same menu structure)
          let opened = false;
          const editMenu = document.getElementById('docs-edit-menu');
          if (editMenu) {
            editMenu.click();
            await sleep(300);
            const items = document.querySelectorAll('.goog-menuitem-content');
            const frItem = Array.from(items).find(el =>
              el.textContent.trim().toLowerCase().includes('find and replace')
            );
            if (frItem) {
              (frItem.closest('.goog-menuitem') || frItem).click();
              await sleep(400);
              opened = true;
            }
          }
          if (!opened) return { success: false, message: "Could not open Find & Replace" };

          const inputs = document.querySelectorAll('[aria-label="Find"], [aria-label="Replace with"]');
          if (inputs.length >= 2) {
            fillInput(inputs[0], payload.find_text || '');
            await sleep(100);
            fillInput(inputs[1], payload.replace_text || '');
            await sleep(200);
            const replaceAll = Array.from(document.querySelectorAll('button')).find(b =>
              b.textContent.toLowerCase().includes('replace all')
            );
            if (replaceAll) { replaceAll.click(); await sleep(300); }
            return { success: true, message: "Find and replace completed" };
          }
          return { success: false, message: "Could not find dialog inputs" };
        }

        default:
          return { success: false, message: `Unsupported Google Sheets action: ${type}. For full editing, install the Google Workspace add-on.` };
      }
    }

    // ── Google Slides Actions ───────────────────────────────────────────
    if (site === "google_slides") {
      return { success: false, message: `Google Slides editing via Chrome extension coming soon. Install the Google Workspace add-on for full functionality.` };
    }

    return { success: false, message: `Not on a Google Workspace app` };
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

  // Expose workspace handler on window so background.js can call it
  // directly via chrome.scripting.executeScript (bypasses message port issues)
  window.__tsifl_workspace_action = handleWorkspaceAction;

  // ── Message Listener (stored on window so it can be replaced) ──────────
  window.__tsifl_msg_handler = function(msg, sender, sendResponse) {
    if (msg.action === "capture_context") {
      sendResponse(getContext());
      return;
    }
    if (msg.action === "execute_dom_action") {
      sendResponse(handleDomAction(msg.type, msg.payload));
      return;
    }
    if (msg.action === "execute_workspace_action") {
      console.log("[tsifl] workspace action:", msg.type, msg.payload);
      handleWorkspaceAction(msg.type, msg.payload)
        .then(result => {
          console.log("[tsifl] workspace result:", result);
          sendResponse(result);
        })
        .catch(e => {
          console.error("[tsifl] workspace error:", e);
          sendResponse({ success: false, message: e.message });
        });
      return true; // keep channel open for async
    }
    if (msg.action === "get_full_page_text") {
      sendResponse({ text: getFullPageText() });
      return;
    }
  };
  chrome.runtime.onMessage.addListener(window.__tsifl_msg_handler);

  console.log("[tsifl] content script v" + TSIFL_CONTENT_VERSION + " loaded on", detectSite());
}
