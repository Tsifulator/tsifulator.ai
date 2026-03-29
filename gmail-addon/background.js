/**
 * tsifl — Gmail Chrome Extension Background Service Worker
 * Handles extension icon click to toggle the sidebar.
 */

chrome.action.onClicked.addListener((tab) => {
  if (tab.url && tab.url.includes("mail.google.com")) {
    chrome.tabs.sendMessage(tab.id, { action: "toggle_sidebar" });
  }
});
