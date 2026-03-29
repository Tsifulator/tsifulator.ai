/**
 * tsifl — Background Service Worker
 * Coordinates between the Side Panel and content scripts.
 */

// Open side panel when toolbar icon is clicked
chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true });

// Keyboard shortcut
chrome.commands.onCommand.addListener((command, tab) => {
  if (command === "toggle-sidebar") {
    chrome.sidePanel.open({ windowId: tab.windowId });
  }
});

// Message handler — relay between panel and content scripts
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.action === "get_context") {
    chrome.tabs.query({ active: true, currentWindow: true }, async (tabs) => {
      const tab = tabs[0];
      if (!tab?.id || tab.url?.startsWith("chrome://")) {
        sendResponse({ context: { app: "browser", url: tab?.url || "", title: tab?.title || "" } });
        return;
      }
      try {
        await chrome.scripting.executeScript({
          target: { tabId: tab.id },
          files: ["content.js"],
        });
      } catch (e) { /* may already be injected or restricted page */ }

      try {
        chrome.tabs.sendMessage(tab.id, { action: "capture_context" }, (response) => {
          if (chrome.runtime.lastError || !response) {
            sendResponse({ context: { app: "browser", url: tab.url, title: tab.title } });
          } else {
            sendResponse({ context: response });
          }
        });
      } catch (e) {
        sendResponse({ context: { app: "browser", url: tab.url, title: tab.title } });
      }
    });
    return true;
  }

  if (msg.action === "execute_browser_action") {
    handleBrowserAction(msg, sendResponse);
    return true;
  }
});

async function handleBrowserAction(msg, sendResponse) {
  const { type, payload } = msg;
  try {
    switch (type) {
      case "open_url": {
        await chrome.tabs.create({ url: payload.url });
        sendResponse({ success: true, message: `Opened ${payload.url}` });
        break;
      }
      case "open_url_current_tab": {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        if (tab) await chrome.tabs.update(tab.id, { url: payload.url });
        sendResponse({ success: true, message: `Navigated to ${payload.url}` });
        break;
      }
      case "search_web": {
        const q = encodeURIComponent(payload.query);
        await chrome.tabs.create({ url: `https://www.google.com/search?q=${q}` });
        sendResponse({ success: true, message: `Searched: ${payload.query}` });
        break;
      }
      case "navigate_back": {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        if (tab) await chrome.tabs.goBack(tab.id);
        sendResponse({ success: true, message: "Went back" });
        break;
      }
      case "navigate_forward": {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        if (tab) await chrome.tabs.goForward(tab.id);
        sendResponse({ success: true, message: "Went forward" });
        break;
      }
      case "scroll_to":
      case "click_element":
      case "fill_input":
      case "extract_text": {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        if (tab) {
          chrome.tabs.sendMessage(tab.id, { action: "execute_dom_action", type, payload }, (response) => {
            sendResponse(response || { success: false, message: "No response from page" });
          });
        } else {
          sendResponse({ success: false, message: "No active tab" });
        }
        break;
      }
      default:
        sendResponse({ success: false, message: `Unknown action: ${type}` });
    }
  } catch (e) {
    sendResponse({ success: false, message: e.message });
  }
}
