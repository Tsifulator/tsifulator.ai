/**
 * tsifl — Background Service Worker (MV3)
 *
 * Responsibilities:
 * 1. Open side panel on toolbar icon click
 * 2. Handle keyboard shortcut (Cmd+Shift+E)
 * 3. Relay context capture between panel and content scripts
 * 4. Execute browser actions (open tabs, search, navigate, DOM actions)
 */

// ── Side Panel Behavior ─────────────────────────────────────────────────
// Must run EVERY time the service worker starts (it can restart at any time in MV3)

try {
  chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true });
} catch (e) {
  console.warn("tsifl: setPanelBehavior failed on startup:", e);
}

// Also set on install/update for safety
chrome.runtime.onInstalled.addListener(() => {
  try {
    chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true });
  } catch (e) {
    console.warn("tsifl: setPanelBehavior failed on install:", e);
  }
});

// ── Keyboard Shortcut ───────────────────────────────────────────────────
// Cmd+Shift+E (Mac) / Ctrl+Shift+E (Win/Linux)
// NOTE: Cmd+Shift+T was Chrome's "reopen closed tab" — conflicted and never fired.

chrome.commands.onCommand.addListener(async (command) => {
  if (command === "toggle-sidebar") {
    try {
      const win = await chrome.windows.getCurrent();
      if (win?.id) {
        await chrome.sidePanel.open({ windowId: win.id });
      }
    } catch (e) {
      console.warn("tsifl: keyboard shortcut open failed:", e);
    }
  }
});

// ── Message Router ──────────────────────────────────────────────────────
// Panel.js and content.js communicate through here.

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.action === "get_context") {
    handleGetContext()
      .then(ctx => sendResponse({ context: ctx }))
      .catch(() => sendResponse({ context: { app: "browser", url: "", title: "" } }));
    return true; // Keep channel open for async response
  }

  if (msg.action === "execute_browser_action") {
    handleBrowserAction(msg.type, msg.payload)
      .then(result => sendResponse(result))
      .catch(e => sendResponse({ success: false, message: e.message }));
    return true;
  }

  if (msg.action === "get_page_text") {
    handleGetPageText()
      .then(result => sendResponse(result))
      .catch(() => sendResponse({ text: "" }));
    return true;
  }
});

// ── Context Capture ─────────────────────────────────────────────────────

async function handleGetContext() {
  const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
  const tab = tabs[0];

  // Can't inject into chrome:// or extension pages
  if (!tab?.id || tab.url?.startsWith("chrome://") || tab.url?.startsWith("chrome-extension://") || tab.url?.startsWith("about:")) {
    return { app: "browser", url: tab?.url || "", title: tab?.title || "" };
  }

  // Ensure content script is injected (may already be via manifest)
  await safeInjectContentScript(tab.id);

  try {
    const response = await sendToTab(tab.id, { action: "capture_context" });
    if (response && response.app) {
      return response;
    }
  } catch (e) {
    // Content script not available — return basic tab info
  }

  return { app: "browser", url: tab.url || "", title: tab.title || "" };
}

async function handleGetPageText() {
  const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
  const tab = tabs[0];
  if (!tab?.id || tab.url?.startsWith("chrome://")) return { text: "" };

  await safeInjectContentScript(tab.id);

  try {
    const response = await sendToTab(tab.id, { action: "get_full_page_text" });
    return { text: response?.text || "" };
  } catch (e) {
    return { text: "" };
  }
}

// ── Browser Action Execution ────────────────────────────────────────────

async function handleBrowserAction(type, payload) {
  if (!type) return { success: false, message: "No action type" };
  if (!payload) payload = {};

  switch (type) {
    case "open_url": {
      const url = payload.url;
      if (!url) return { success: false, message: "No URL provided" };
      await chrome.tabs.create({ url, active: true });
      return { success: true, message: `Opened ${url}` };
    }

    case "open_url_current_tab": {
      const url = payload.url;
      if (!url) return { success: false, message: "No URL provided" };
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      if (tab?.id) {
        await chrome.tabs.update(tab.id, { url });
      } else {
        await chrome.tabs.create({ url, active: true });
      }
      return { success: true, message: `Navigated to ${url}` };
    }

    case "search_web": {
      const query = payload.query;
      if (!query) return { success: false, message: "No search query" };
      const q = encodeURIComponent(query);
      await chrome.tabs.create({ url: `https://www.google.com/search?q=${q}`, active: true });
      return { success: true, message: `Searched: ${query}` };
    }

    case "navigate_back": {
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      if (tab?.id) await chrome.tabs.goBack(tab.id);
      return { success: true, message: "Went back" };
    }

    case "navigate_forward": {
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      if (tab?.id) await chrome.tabs.goForward(tab.id);
      return { success: true, message: "Went forward" };
    }

    // DOM actions — relay to content script
    case "scroll_to":
    case "click_element":
    case "fill_input":
    case "extract_text": {
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      if (!tab?.id) return { success: false, message: "No active tab" };
      await safeInjectContentScript(tab.id);
      try {
        const response = await sendToTab(tab.id, { action: "execute_dom_action", type, payload });
        return response || { success: false, message: "No response from content script" };
      } catch (e) {
        return { success: false, message: `DOM action failed: ${e.message}` };
      }
    }

    default:
      return { success: false, message: `Unknown browser action: ${type}` };
  }
}

// ── Helpers ──────────────────────────────────────────────────────────────

async function safeInjectContentScript(tabId) {
  try {
    await chrome.scripting.executeScript({
      target: { tabId },
      files: ["content.js"],
    });
  } catch (e) {
    // Already injected, restricted page, or no permission
  }
}

function sendToTab(tabId, message) {
  return new Promise((resolve, reject) => {
    try {
      chrome.tabs.sendMessage(tabId, message, (response) => {
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message));
        } else {
          resolve(response);
        }
      });
    } catch (e) {
      reject(e);
    }
  });
}
