// Background Script (bg.js)

// Global state
let ACCESSPANEL_GLOBAL_OPEN = false;

// URL validation helpers
function isValidWfmSessionUrl(url) {
  if (!url) return false;
  if (!/mykronos\.com/i.test(url)) return false;
  if (/mykronos\.com\/authn\//i.test(url)) return false;
  if (/:\/\/adp-developer\.mykronos\.com\//i.test(url)) return false;
  return true;
}

function getOrigin(u) {
  try {
    const url = new URL(u);
    return url.origin;
  } catch {
    return null;
  }
}

// Context management functions
async function setLinkedContext(tab) {
  if (!tab?.id || !isValidWfmSessionUrl(tab.url)) return;

  const payload = {
    hermesLinkedTabId: tab.id,
    hermesLinkedUrl: tab.url,
    hermesLinkedOrigin: getOrigin(tab.url),
    hermesLinkedTitle: tab.title || "",
    hermesLinkedStatus: "ok",
  };
  await chrome.storage.session.set(payload);
}

async function clearLinkedContext(reason = "closed") {
  const payload = {
    hermesLinkedTabId: null,
    hermesLinkedUrl: null,
    hermesLinkedOrigin: null,
    hermesLinkedTitle: "",
    hermesLinkedStatus: reason,
  };
  await chrome.storage.session.set(payload);
}

// Panel management functions
async function setSidePanelEnabledForAll(enabled) {
  try {
    const tabs = await chrome.tabs.query({});
    const ops = tabs.map((t) =>
      chrome.sidePanel
        .setOptions({
          tabId: t.id,
          path: "accesspanel.html",
          enabled,
        })
        .catch(() => {})
    );
    await Promise.all(ops);
  } catch (error) {
    console.error("Failed to set panel state:", error?.message || error);
  }
}

// Event handler functions
async function handleStartup() {
  try {
    // NOTE: the storage key is "accesspanelGlobalOpen" (lowercase p)
    const { accesspanelGlobalOpen } = await chrome.storage.session.get(
      "accesspanelGlobalOpen"
    );
    ACCESSPANEL_GLOBAL_OPEN = !!accesspanelGlobalOpen;
  } catch (error) {
    console.error("Startup state retrieval failed:", error?.message || error);
  }
}

async function handleToolbarClick(tab) {
  if (!tab?.id) return;

  if (ACCESSPANEL_GLOBAL_OPEN) {
    setSidePanelEnabledForAll(false).catch((e) =>
      console.error("Failed to disable panels:", e?.message || e)
    );
    ACCESSPANEL_GLOBAL_OPEN = false;
    chrome.storage.session
      .set({ accesspanelGlobalOpen: false })
      .catch((e) => console.error("Failed to update storage:", e?.message || e));
    return;
  }

  // Enable panel for current tab
  chrome.sidePanel
    .setOptions({
      tabId: tab.id,
      path: "accesspanel.html",
      enabled: true,
    })
    .catch((e) => console.error("Failed to enable panel:", e?.message || e));

  // Open the panel
  chrome.sidePanel
    .open({ tabId: tab.id })
    .catch((e) => console.error("Failed to open panel:", e?.message || e));

  // Enable for all tabs (so new tabs also have it enabled)
  setSidePanelEnabledForAll(true).catch((e) =>
    console.error("Failed to enable panels globally:", e?.message || e)
  );

  // Update state
  ACCESSPANEL_GLOBAL_OPEN = true;
  chrome.storage.session
    .set({ accesspanelGlobalOpen: true })
    .catch((e) => console.error("Failed to update storage:", e?.message || e));

  // Handle linking
  try {
    if (isValidWfmSessionUrl(tab.url)) {
      await setLinkedContext(tab);
    } else if (/mykronos\.com\/authn\//i.test(tab?.url || "")) {
      await chrome.storage.session.set({ hermesLinkedStatus: "stale" });
    }
  } catch (error) {
    console.error("Failed to handle session linking:", error?.message || error);
  }
}

async function handleTabUpdate(tabId, changeInfo, tab) {
  // Sync panel state
  try {
    const isOpen = ACCESSPANEL_GLOBAL_OPEN;
    await chrome.sidePanel.setOptions({
      tabId,
      path: "accesspanel.html",
      enabled: isOpen,
    });
  } catch (error) {
    console.error("Failed to update panel state:", error?.message || error);
  }

  // Handle linked tab updates
  try {
    const { hermesLinkedTabId } = await chrome.storage.session.get(
      "hermesLinkedTabId"
    );
    if (!hermesLinkedTabId || tabId !== hermesLinkedTabId) return;

    if (changeInfo.url) {
      if (isValidWfmSessionUrl(changeInfo.url)) {
        await setLinkedContext({
          id: tabId,
          url: changeInfo.url,
          title: tab?.title,
        });
      } else {
        await chrome.storage.session.set({
          hermesLinkedStatus: "stale",
          hermesLinkedUrl: changeInfo.url || null,
        });
      }
    }
    if (changeInfo.title) {
      await chrome.storage.session.set({ hermesLinkedTitle: changeInfo.title });
    }
  } catch (error) {
    console.error("Failed to handle tab update:", error?.message || error);
  }
}

async function handleTabRemoved(tabId) {
  try {
    const { hermesLinkedTabId } = await chrome.storage.session.get(
      "hermesLinkedTabId"
    );
    if (hermesLinkedTabId && tabId === hermesLinkedTabId) {
      await clearLinkedContext("closed");
    }
  } catch (error) {
    console.error("Failed to handle tab removal:", error?.message || error);
  }
}

// Event Listeners
chrome.runtime.onStartup?.addListener(handleStartup);
chrome.action.onClicked.addListener(handleToolbarClick);
chrome.tabs.onUpdated.addListener(handleTabUpdate);
chrome.tabs.onRemoved.addListener(handleTabRemoved);