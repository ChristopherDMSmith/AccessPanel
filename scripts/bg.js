// Background Script (bg.js)

// Global state
let ACCESSPANEL_GLOBAL_OPEN = false;

// Retrieve global open state from storage
async function getGlobalOpen() {
  try {
    const { accesspanelGlobalOpen } = await chrome.storage.session.get(
      "accesspanelGlobalOpen"
    );
    return !!accesspanelGlobalOpen;
  } catch {
    return false;
  }
}

// URL validation helpers
function isValidWfmSessionUrl(url) {
  if (!url) return false;
  if (!/mykronos\.com/i.test(url)) return false;
  if (/mykronos\.com\/authn\//i.test(url)) return false;
  if (/:\/\/adp-developer\.mykronos\.com\//i.test(url)) return false;
  return true;
}

// Origin check helpers
function getOrigin(u) {
  try {
    return new URL(u).origin;
  } catch {
    return null;
  }
}

// Check if URL belongs to WFM origins
function isWfmOrigin(u) {
  const origin = getOrigin(u);
  if (!origin) return false;
  return /:\/\/[^/]*mykronos\.com$/i.test(origin);
}

// Check if URL belongs to TMS origins
function isTmsOrigin(u) {
  const origin = getOrigin(u);
  if (!origin) return false;
  return /:\/\/(adpvantage|testadpvantage)\.adp\.com$/i.test(origin);
}

// Clear linked context
async function clearLinkedContext(reason = "unknown") {
  try {
    await chrome.storage.session.set({
      hermesLinkedStatus: "stale",
      hermesLinkedTabId: null,
      hermesLinkedUrl: null,
      hermesLinkedTitle: null,
      hermesLinkedReason: reason,
      hermesLinkedAt: Date.now(),
    });
  } catch (e) {
    console.error("Failed to clear linked context:", e?.message || e);
  }
}

// Set linked context (WFM tab is the source of truth)
async function setLinkedContext(tab) {
  if (!tab?.id) return false;

  try {
    await chrome.storage.session.set({
      hermesLinkedStatus: "linked",
      hermesLinkedTabId: tab.id,
      hermesLinkedUrl: tab.url || null,
      hermesLinkedTitle: tab.title || null,
      hermesLinkedAt: Date.now(),
    });
    return true;
  } catch (e) {
    console.error("Failed to set linked context:", e?.message || e);
    return false;
  }
}

// Enable/disable side panel for all tabs
async function setSidePanelEnabledForAll(enabled) {
  const tabs = await chrome.tabs.query({});
  await Promise.allSettled(
    tabs.map((t) =>
      chrome.sidePanel.setOptions({
        tabId: t.id,
        path: "accesspanel.html",
        enabled,
      })
    )
  );
}

// Event handler functions
async function handleStartup() {
  try {
    ACCESSPANEL_GLOBAL_OPEN = await getGlobalOpen();
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

  // Enable for all tabs (so it stays available)
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
  // Sync panel state (read persisted state; do not trust in-memory globals)
  try {
    const isOpen = await getGlobalOpen();
    ACCESSPANEL_GLOBAL_OPEN = isOpen;

    await chrome.sidePanel.setOptions({
      tabId,
      path: "accesspanel.html",
      enabled: isOpen,
    });
  } catch (error) {
    console.error("Failed to update panel state:", error?.message || error);
  }

  // Handle linking updates (WFM/TMS)
  try {
    const { hermesLinkedTabId, hermesLinkedUrl } =
      await chrome.storage.session.get(["hermesLinkedTabId", "hermesLinkedUrl"]);

    // If the linked tab is gone or becomes invalid, clear it
    if (hermesLinkedTabId && hermesLinkedTabId === tabId) {
      const currentUrl = changeInfo.url || tab?.url || null;

      if (!isValidWfmSessionUrl(currentUrl)) {
        await clearLinkedContext("invalid_wfm_url");
        return;
      }

      // Keep linked metadata fresh
      if (changeInfo.url) {
        await chrome.storage.session.set({
          hermesLinkedStatus: "linked",
          hermesLinkedUrl: changeInfo.url,
          hermesLinkedAt: Date.now(),
          title: tab?.title,
        });
      }
      if (changeInfo.title) {
        await chrome.storage.session.set({
          hermesLinkedTitle: changeInfo.title,
        });
      }
      return;
    }

    // If WFM URL changes in any tab, consider it as a candidate for linking
    if (changeInfo.url) {
      const url = changeInfo.url;

      if (isValidWfmSessionUrl(url)) {
        // Prefer WFM tab as link source
        const candidateTab = tab?.id ? tab : await chrome.tabs.get(tabId);
        await setLinkedContext(candidateTab);
      } else if (hermesLinkedUrl && tabId === hermesLinkedTabId) {
        // Linked tab changed away from WFM
        await clearLinkedContext("linked_tab_navigated_away");
      } else {
        // If it looks like a related origin but not valid, mark stale
        if (isWfmOrigin(url) || isTmsOrigin(url)) {
          await chrome.storage.session.set({
            hermesLinkedStatus: "stale",
            hermesLinkedUrl: url || null,
          });
        }
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
chrome.runtime.onInstalled?.addListener(handleStartup);
handleStartup().catch(() => {});
chrome.action.onClicked.addListener(handleToolbarClick);
chrome.tabs.onUpdated.addListener(handleTabUpdate);
chrome.tabs.onRemoved.addListener(handleTabRemoved);