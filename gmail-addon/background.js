/**
 * tsifl — Background Service Worker
 * Handles extension icon click and keyboard shortcut to toggle sidebar on any page.
 */

chrome.action.onClicked.addListener((tab) => {
  chrome.tabs.sendMessage(tab.id, { action: "toggle_sidebar" });
});

chrome.commands.onCommand.addListener((command, tab) => {
  if (command === "toggle-sidebar") {
    chrome.tabs.sendMessage(tab.id, { action: "toggle_sidebar" });
  }
});
