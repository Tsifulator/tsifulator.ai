/**
 * tsifl — Background Service Worker
 * Coordinates between the Side Panel and content scripts.
 * All browser actions and context capture flow through here for reliability.
 */

// ── Side Panel Setup ────────────────────────────────────────────────────

// Ensure side panel opens when toolbar icon is clicked
chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true }).catch(() => {});

// Keyboard shortcut — Cmd+Shift+T / Ctrl+Shift+T
chrome.commands.onCommand.addListener(async (command) => {
  if (command === "toggle-sidebar") {
    try {
      const win = await chrome.windows.getCurrent();
      await chrome.sidePanel.open({ windowId: win.id });
    } catch (e) {
      console.warn("tsifl: Could not open side panel:", e.message);
    }
  }
});

// Re-register panel behavior on install/update (service worker can restart)
chrome.runtime.onInstalled.addListener(() => {
  chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true }).catch(() => {});
});

// ── Message Handler ─────────────────────────────────────────────────────

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.action === "get_context") {
    handleGetContext(msg).then(sendResponse).catch(() => {
      sendResponse({ context: { app: "browser", url: "", title: "" } });
    });
    return true;
  }

  if (msg.action === "execute_browser_action") {
    handleBrowserAction(msg.type, msg.payload).then(sendResponse).catch((e) => {
      sendResponse({ success: false, message: e.message });
    });
    return true;
  }

  if (msg.action === "get_page_text") {
    handleGetPageText().then(sendResponse).catch(() => {
      sendResponse({ text: "" });
    });
    return true;
  }
});

// ── Context Capture ─────────────────────────────────────────────────────

async function handleGetContext(msg) {
  const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
  const tab = tabs[0];
  if (!tab?.id || tab.url?.startsWith("chrome://") || tab.url?.startsWith("chrome-extension://")) {
    return { context: { app: "browser", url: tab?.url || "", title: tab?.title || "" } };
  }

  // Ensure content script is injected
  await injectContentScript(tab.id);

  try {
    const response = await sendTabMessage(tab.id, { action: "capture_context" });
    return { context: response || { app: "browser", url: tab.url, title: tab.title } };
  } catch {
    return { context: { app: "browser", url: tab.url, title: tab.title } };
  }
}

async function handleGetPageText() {
  const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
  const tab = tabs[0];
  if (!tab?.id || tab.url?.startsWith("chrome://")) return { text: "" };

  await injectContentScript(tab.id);

  try {
    const response = await sendTabMessage(tab.id, { action: "get_full_page_text" });
    return { text: response?.text || "" };
  } catch {
    return { text: "" };
  }
}

// ── Browser Action Execution ────────────────────────────────────────────

async function handleBrowserAction(type, payload) {
  switch (type) {
    case "open_url": {
      await chrome.tabs.create({ url: payload.url, active: true });
      return { success: true, message: `Opened ${payload.url}` };
    }
    case "open_url_current_tab": {
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      if (tab) await chrome.tabs.update(tab.id, { url: payload.url });
      return { success: true, message: `Navigated to ${payload.url}` };
    }
    case "search_web": {
      const q = encodeURIComponent(payload.query);
      await chrome.tabs.create({ url: `https://www.google.com/search?q=${q}`, active: true });
      return { success: true, message: `Searched: ${payload.query}` };
    }
    case "navigate_back": {
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      if (tab) await chrome.tabs.goBack(tab.id);
      return { success: true, message: "Went back" };
    }
    case "navigate_forward": {
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      if (tab) await chrome.tabs.goForward(tab.id);
      return { success: true, message: "Went forward" };
    }
    case "scroll_to":
    case "click_element":
    case "fill_input":
    case "extract_text": {
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      if (!tab?.id) return { success: false, message: "No active tab" };
      await injectContentScript(tab.id);
      const response = await sendTabMessage(tab.id, { action: "execute_dom_action", type, payload });
      return response || { success: false, message: "No response from page" };
    }
    default:
      return { success: false, message: `Unknown action: ${type}` };
  }
}

// ── Helpers ──────────────────────────────────────────────────────────────

async function injectContentScript(tabId) {
  try {
    await chrome.scripting.executeScript({
      target: { tabId },
      files: ["content.js"],
    });
  } catch (e) {
    // May already be injected or page is restricted
  }
}

function sendTabMessage(tabId, message) {
  return new Promise((resolve, reject) => {
    chrome.tabs.sendMessage(tabId, message, (response) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
      } else {
        resolve(response);
      }
    });
  });
}
