// Background Script (bg.js)

// Global state
let ACCESSPANEL_GLOBAL_OPEN = false;

// Track per-tab side panel configuration so we don’t spam setOptions.
// Edge can close the side panel if setOptions is called repeatedly during navigation.
const tabPanelState = new Map(); // tabId -> { enabled: boolean, path: string }

// Edge timing resilience: reopen retries per tab
const openRetryTimers = new Map(); // tabId -> { interval, timeout }

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

// ---- Side panel helpers ---- //

async function ensureSidePanelOptions(tabId, enabled) {
  if (!tabId) return;

  const next = { enabled: !!enabled, path: "accesspanel.html" };
  const prev = tabPanelState.get(tabId);

  // Only set when something changed.
  if (prev && prev.enabled === next.enabled && prev.path === next.path) return;

  try {
    await chrome.sidePanel.setOptions({
      tabId,
      path: next.path,
      enabled: next.enabled,
    });
    tabPanelState.set(tabId, next);
  } catch {
    // ignore per-tab failures (tabs can disappear mid-flight)
  }
}

async function configureAllTabs(enabled) {
  try {
    const tabs = await chrome.tabs.query({});
    await Promise.all(
      tabs.map((t) => ensureSidePanelOptions(t.id, enabled).catch(() => {}))
    );
  } catch (e) {
    console.error("Failed to configure side panel tabs:", e?.message || e);
  }
}

function clearOpenRetries(tabId) {
  const timers = openRetryTimers.get(tabId);
  if (!timers) return;
  try {
    if (timers.interval) clearInterval(timers.interval);
    if (timers.timeout) clearTimeout(timers.timeout);
  } catch {}
  openRetryTimers.delete(tabId);
}

// Try to reopen for a short period; Edge can reject open() during rapid tab transitions.
function scheduleReopenRetries(tabId, windowId = null) {
  if (!ACCESSPANEL_GLOBAL_OPEN || !tabId) return;

  // avoid stacking retries
  clearOpenRetries(tabId);

  const start = Date.now();
  const MAX_MS = 2000; // keep short; enough to survive tab-close focus handoff
  const EVERY_MS = 150;

  const attempt = async () => {
    if (!ACCESSPANEL_GLOBAL_OPEN) return;

    try {
      // Ensure per-tab options are correct before opening
      await ensureSidePanelOptions(tabId, true);

      // Only try to open if the tab still exists and is active in that window.
      // (Side panel is tied to active tab.)
      const [activeTab] = await chrome.tabs.query(
        windowId
          ? { active: true, windowId }
          : { active: true, currentWindow: true }
      );

      if (!activeTab?.id) return;

      // If the active tab changed, move the retry target to the active tab.
      const targetId = activeTab.id;

      await ensureSidePanelOptions(targetId, true);
      await chrome.sidePanel.open({ tabId: targetId });

      // success: stop retrying
      clearOpenRetries(tabId);
    } catch {
      // ignore; we’ll retry until timeout
    }

    if (Date.now() - start > MAX_MS) {
      clearOpenRetries(tabId);
    }
  };

  // kick immediately, then repeat
  attempt().catch(() => {});
  const interval = setInterval(() => {
    attempt().catch(() => {});
  }, EVERY_MS);

  const timeout = setTimeout(() => {
    clearOpenRetries(tabId);
  }, MAX_MS + 250);

  openRetryTimers.set(tabId, { interval, timeout });
}

// Re-open panel on the active tab (helps when Edge closes it after new tab / navigation)
async function reopenPanelIfNeeded(tabId) {
  if (!ACCESSPANEL_GLOBAL_OPEN || !tabId) return;
  try {
    await chrome.sidePanel.open({ tabId });
  } catch {
    // If Edge refuses (timing), retry briefly.
    scheduleReopenRetries(tabId);
  }
}

// ---- Event handlers ---- //

async function handleStartup() {
  try {
    const { accesspanelGlobalOpen } = await chrome.storage.session.get(
      "accesspanelGlobalOpen"
    );
    ACCESSPANEL_GLOBAL_OPEN = !!accesspanelGlobalOpen;

    // Apply remembered state to existing tabs once on startup
    await configureAllTabs(ACCESSPANEL_GLOBAL_OPEN);
  } catch (e) {
    console.error("Startup state retrieval failed:", e?.message || e);
  }
}

async function handleInstalled() {
  // Ensure action click opens panel (supported in Chromium; harmless if not)
  try {
    await chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true });
  } catch {
    // ignore
  }

  // Sync initial tab enablement (Edge can be picky right after install/update)
  try {
    await configureAllTabs(ACCESSPANEL_GLOBAL_OPEN);
  } catch {
    // ignore
  }
}

async function handleToolbarClick(tab) {
  if (!tab?.id) return;

  if (ACCESSPANEL_GLOBAL_OPEN) {
    // Turn OFF globally
    ACCESSPANEL_GLOBAL_OPEN = false;
    chrome.storage.session
      .set({ accesspanelGlobalOpen: false })
      .catch(() => {});
    await configureAllTabs(false);
    return;
  }

  // Turn ON globally
  ACCESSPANEL_GLOBAL_OPEN = true;
  chrome.storage.session.set({ accesspanelGlobalOpen: true }).catch(() => {});

  // Enable current tab first, then open panel
  await ensureSidePanelOptions(tab.id, true);
  await reopenPanelIfNeeded(tab.id);

  // Enable the rest so new tabs also have it available
  configureAllTabs(true).catch(() => {});

  // Handle linking
  try {
    if (isValidWfmSessionUrl(tab.url)) {
      await setLinkedContext(tab);
    } else if (/mykronos\.com\/authn\//i.test(tab?.url || "")) {
      await chrome.storage.session.set({ hermesLinkedStatus: "stale" });
    }
  } catch (e) {
    console.error("Failed to handle session linking:", e?.message || e);
  }
}

// New tab created: set enabled/disabled once (do NOT rely on onUpdated spam)
async function handleTabCreated(tab) {
  if (!tab?.id) return;
  await ensureSidePanelOptions(tab.id, ACCESSPANEL_GLOBAL_OPEN);
}

// When user switches tabs, keep panel “sticky” by reopening if global open
async function handleTabActivated(activeInfo) {
  const tabId = activeInfo?.tabId;
  const windowId = activeInfo?.windowId;
  if (!tabId) return;

  await ensureSidePanelOptions(tabId, ACCESSPANEL_GLOBAL_OPEN);

  // Try to keep it sticky on activation
  if (ACCESSPANEL_GLOBAL_OPEN) {
    // activation is a common moment where Edge might have just closed the panel
    // so use the retry version to be resilient.
    scheduleReopenRetries(tabId, windowId);
  }
}

async function handleTabUpdate(tabId, changeInfo, tab) {
  // IMPORTANT: we no longer call setOptions on every update.
  // That behavior can cause Edge to close the side panel during navigation/new tabs.

  // Handle linked tab updates only
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
  } catch (e) {
    console.error("Failed to handle tab update:", e?.message || e);
  }
}

async function handleTabRemoved(tabId, removeInfo) {
  // Clean cached panel state for this tab
  try {
    tabPanelState.delete(tabId);
  } catch {}
  clearOpenRetries(tabId);

  // Handle linked tab closure
  try {
    const { hermesLinkedTabId } = await chrome.storage.session.get(
      "hermesLinkedTabId"
    );
    if (hermesLinkedTabId && tabId === hermesLinkedTabId) {
      await clearLinkedContext("closed");
    }
  } catch (e) {
    console.error("Failed to handle tab removal:", e?.message || e);
  }

  // Edge resilience: if a tab is closed, Edge may close the panel during the focus handoff.
  // Reopen on the newly-active tab in that same window.
  try {
    if (!ACCESSPANEL_GLOBAL_OPEN) return;
    if (removeInfo?.isWindowClosing) return;

    // Let Edge finish selecting the fallback tab first
    setTimeout(async () => {
      try {
        const [activeTab] = await chrome.tabs.query({
          active: true,
          windowId: removeInfo?.windowId,
        });

        if (!activeTab?.id) return;

        await ensureSidePanelOptions(activeTab.id, true);
        scheduleReopenRetries(activeTab.id, removeInfo?.windowId);
      } catch {
        // ignore
      }
    }, 50);
  } catch {
    // ignore
  }
}

// ---- Event listeners ---- //
chrome.runtime.onStartup?.addListener(handleStartup);
chrome.runtime.onInstalled?.addListener(handleInstalled);

chrome.action.onClicked.addListener(handleToolbarClick);

chrome.tabs.onCreated.addListener(handleTabCreated);
chrome.tabs.onActivated.addListener(handleTabActivated);

chrome.tabs.onUpdated.addListener(handleTabUpdate);
chrome.tabs.onRemoved.addListener(handleTabRemoved);
