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
