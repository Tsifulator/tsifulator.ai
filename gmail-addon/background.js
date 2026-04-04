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

  // Set up recurring token refresh alarm (survives service worker idle)
  chrome.alarms.create("tsifl-token-refresh", { periodInMinutes: 30 });
});

// Also create alarm on service worker start (in case install event was missed)
chrome.alarms.get("tsifl-token-refresh", (alarm) => {
  if (!alarm) {
    chrome.alarms.create("tsifl-token-refresh", { periodInMinutes: 30 });
  }
});

// ── Token Refresh Alarm ──────────────────────────────────────────────────
// Keeps the session alive even when the side panel is closed.
// MV3 service workers go idle, killing setInterval — alarms persist.

const SUPABASE_URL = "https://dvynmzeyttwlmvunicqz.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImR2eW5temV5dHR3bG12dW5pY3F6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQ2NTIwMTIsImV4cCI6MjA5MDIyODAxMn0.9j_f-2f1VswxWfqiuXy4bPnUi1qLk9nAeTDlodUBUZw";
const BACKEND_URL = "https://focused-solace-production-6839.up.railway.app";

chrome.alarms.onAlarm.addListener(async (alarm) => {
  if (alarm.name !== "tsifl-token-refresh") return;

  try {
    // Try backend session first (most up-to-date)
    let refreshToken = null;
    try {
      const resp = await fetch(`${BACKEND_URL}/auth/get-session`);
      const data = await resp.json();
      if (data.session?.refresh_token) {
        refreshToken = data.session.refresh_token;
      }
    } catch (e) {}

    // Fall back to chrome.storage.local
    if (!refreshToken) {
      const stored = await chrome.storage.local.get("tsifl_session");
      if (stored.tsifl_session?.refresh_token) {
        refreshToken = stored.tsifl_session.refresh_token;
      }
    }

    if (!refreshToken) return; // Not logged in

    // Refresh the token via Supabase REST API
    const resp = await fetch(`${SUPABASE_URL}/auth/v1/token?grant_type=refresh_token`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "apikey": SUPABASE_ANON_KEY },
      body: JSON.stringify({ refresh_token: refreshToken }),
    });
    const result = await resp.json();

    if (result.access_token && result.user) {
      // Save fresh tokens locally
      chrome.storage.local.set({
        tsifl_session: result,
        tsifl_email: result.user.email,
      });
      // Sync to backend
      await fetch(`${BACKEND_URL}/auth/set-session`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          access_token: result.access_token,
          refresh_token: result.refresh_token,
          user_id: result.user?.id || "",
          email: result.user?.email || "",
        }),
      });
    }
  } catch (e) {
    console.warn("tsifl: background token refresh failed:", e);
  }
});

// ── Keyboard Shortcut ───────────────────────────────────────────────────
// Cmd+Shift+U (Mac) / Ctrl+Shift+U (Win/Linux)
// Uses _execute_action command which Chrome handles automatically —
// it triggers the extension action button click, opening the side panel.

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

  if (msg.action === "execute_workspace_action") {
    handleWorkspaceAction(msg.type, msg.payload)
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

// ── Google Workspace Action Execution ───────────────────────────────

async function handleWorkspaceAction(type, payload) {
  console.log("[tsifl bg] workspace action:", type, payload);
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  console.log("[tsifl bg] active tab:", tab?.id, tab?.url?.slice(0, 60));
  if (!tab?.id) return { success: false, message: "No active tab" };

  // Ensure content script is loaded
  await safeInjectContentScript(tab.id);

  // Use chrome.scripting.executeScript to call the workspace handler DIRECTLY.
  // This bypasses chrome.runtime.sendMessage which suffers from
  // "message port closed before response" errors when multiple listeners exist.
  try {
    const results = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: async (actionType, actionPayload) => {
        if (typeof window.__tsifl_workspace_action === "function") {
          return await window.__tsifl_workspace_action(actionType, actionPayload);
        }
        return { success: false, message: "Content script not ready — try refreshing the page." };
      },
      args: [type, payload],
    });
    const result = results?.[0]?.result;
    console.log("[tsifl bg] executeScript result:", result);
    return result || { success: false, message: "No result from content script" };
  } catch (e) {
    console.error("[tsifl bg] executeScript failed:", e);
    return { success: false, message: `Workspace action failed: ${e.message}` };
  }
}

// ── Helpers ──────────────────────────────────────────────────────────────

async function safeInjectContentScript(tabId) {
  try {
    await chrome.scripting.executeScript({
      target: { tabId },
      files: ["content.js"],
    });
    console.log("[tsifl bg] content.js injected on tab", tabId);
  } catch (e) {
    console.warn("[tsifl bg] inject error (may be OK):", e.message);
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
