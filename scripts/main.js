// main.js: Core functionality for AccessPanel extension UI

// ===== GLOBALS ===== //
const LOGGING_ENABLED = false;
const storageKey = "clientdata";
const sessionStorageKey = "clientdata_session";
const SESSION_ONLY_CLIENT_FIELDS = new Set([
  "clientsecret",
  "accesstoken",
  "refreshtoken",
  "accessTokenSource",
  "effectivedatetime",
  "expirationdatetime",
  "refreshExpirationDateTime",
]);
let accessTokenTimerInterval = null;
let refreshTokenTimerInterval = null;
let lastRequestDetails = null;

// ===== UTILITIES ===== //
// Event handler attachment utility
function on(id, event, handler, options = {}) {
  const el = document.getElementById(id);
  if (!el) return;

  // guard to prevent double-wiring
  if (options.onceKey) {
    const key = `wired_${options.onceKey}`;
    if (el.dataset[key] === "1") return;
    el.dataset[key] = "1";
  }

  el.addEventListener(event, handler, options.listenerOptions);
}

// Event handler attachment utility for async functions
function onAsync(id, event, handler, options = {}) {
  on(
    id,
    event,
    (e) => {
      try {
        Promise.resolve(handler(e)).catch((err) =>
          console.error(`${id} handler failed:`, err)
        );
      } catch (err) {
        console.error(`${id} handler failed:`, err);
      }
    },
    options
  );
}

// Download utility for export functions
function downloadFile(filename, content, mimeType) {
  downloadFileWithContext(document, URL, filename, content, mimeType);
}

// Download utility that can run in either the panel window or a popup window
function downloadFileWithContext(doc, urlApi, filename, content, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = urlApi.createObjectURL(blob);

  const a = doc.createElement("a");
  a.href = url;
  a.download = filename;
  a.rel = "noopener";

  doc.body.appendChild(a);
  a.click();
  doc.body.removeChild(a);

  setTimeout(() => urlApi.revokeObjectURL(url), 0);
}

// Get current theme colors for popup windows
function getPopupThemeColors() {
  const styles = getComputedStyle(document.documentElement);

  return {
    primary:
      styles.getPropertyValue("--primary-color").trim() || "#f5f5f5",
    secondary:
      styles.getPropertyValue("--secondary-color").trim() || "#0059B3",
    accent:
      styles.getPropertyValue("--accent-color").trim() || "#00AEEF",
    highlight:
      styles.getPropertyValue("--highlight-color").trim() || "#007ACC",
    buttonText:
      styles.getPropertyValue("--buttontext-color").trim() || "#FFFFFF",
  };
}

// Open and initialize a popup window without document.write()
function createPopupWindow(title, features, cssText = "") {
  const popup = window.open("", "_blank", features);
  if (!popup) return null;

  const doc = popup.document;

  doc.title = title;

  // Clear the default blank document safely
  doc.head.replaceChildren();
  doc.body.replaceChildren();

  const titleElement = doc.createElement("title");
  titleElement.textContent = title;
  doc.head.appendChild(titleElement);

  if (cssText) {
    const style = doc.createElement("style");
    style.textContent = cssText;
    doc.head.appendChild(style);
  }

  return popup;
}

// Format local date as YYYY-MM-DD with optional day offset.
function formatLocalYMD(daysOffset = 0, base = new Date()) {
  // Start from local midnight to avoid DST/clock noise
  const d = new Date(base.getFullYear(), base.getMonth(), base.getDate());
  d.setDate(d.getDate() + daysOffset);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

// Toggle visibility of a given input field
function toggleFieldVisibility(inputId, iconId) {
  const input = document.getElementById(inputId);
  const icon = document.getElementById(iconId);
  if (!input || !icon) return;

  const nowPassword = input.type !== "password";
  input.type = nowPassword ? "password" : "text";
  icon.src = nowPassword ? "icons/eyeopen.png" : "icons/eyeclosed.png";
  icon.alt = nowPassword ? "Show" : "Hide";
}

// Ensure a given input field is masked (type="password")
function ensureMasked(inputId) {
  const el = document.getElementById(inputId);
  if (!el) return;
  if (el.type !== "password") el.type = "password";
}

// Determine whether the current extension context is incognito
async function isIncognitoContext() {
  if (chrome.extension?.inIncognitoContext) return true;

  try {
    const tabs = await chrome.tabs.query({
      active: true,
      lastFocusedWindow: true,
    });
    return !!tabs?.[0]?.incognito;
  } catch {
    return false;
  }
}

// Helper to coorelate panel URL from template or path
function resolvePanelUrl(templateOrPath, ssoClientUrl) {
  if (!templateOrPath || typeof templateOrPath !== "string") return "";

  const raw = templateOrPath.trim();
  if (!raw) return "";

  // Absolute URL: allows links from accesspanel.json to full URLs
  if (/^https?:\/\//i.test(raw)) return raw;

  // Backward compatibility: old templated format
  const templatePrefix = "https://*.mykronos.com/";
  if (raw.startsWith(templatePrefix)) {
    console.warn(
      "[AccessPanel] Deprecated URL template in config. Use a relative path instead:",
      raw
    );
    return ssoClientUrl + raw.slice(templatePrefix.length);
  }

  // Relative: normalize slashes and join
  return ssoClientUrl + raw.replace(/^\/+/, "");
}
// ===================== //

// ===== TITLE BAR ACTIONS ===== //
// Reload app
function reloadApp() {
  window.location.reload();
}

// Reload App Initializer
function initTitleBarActions() {
  const reloadBtn = document.getElementById("reload-app");
  if (reloadBtn) {
    reloadBtn.addEventListener("click", reloadApp);
  }
}
// ============================= //

// ===== MENU BAR | ADMIN FUNCTIONS ===== //
// Menu initializer
function initMenus() {
  const bar = document.querySelector(".menu-bar");
  if (!bar) return;

  const dropdowns = Array.from(bar.querySelectorAll(".dropdown"));
  const buttons = dropdowns.map((dd) => dd.querySelector(".menu-btn"));

  const closeAll = (except = null) => {
    dropdowns.forEach((d) => {
      if (d !== except) {
        d.classList.remove("open");
        const b = d.querySelector(".menu-btn");
        if (b) b.setAttribute("aria-expanded", "false");
      }
    });
  };

  // toggle each menu on button click; close the rest
  buttons.forEach((btn, i) => {
    if (!btn) return;
    const dd = dropdowns[i];
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const willOpen = !dd.classList.contains("open");
      closeAll();
      dd.classList.toggle("open", willOpen);
      btn.setAttribute("aria-expanded", String(willOpen));
    });
  });

  // close menu when pointer leaves the defined area
  dropdowns.forEach((dd) => {
    dd.addEventListener("mouseleave", () => {
      dd.classList.remove("open");
      const b = dd.querySelector(".menu-btn");
      if (b) b.setAttribute("aria-expanded", "false");
    });
  });

  // click outside closes all
  document.addEventListener("click", () => closeAll());

  // esc closes all
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") closeAll();
  });
}

// Click to open menu item
function initClickMenus() {
  const bar = document.querySelector(".menu-bar");
  if (!bar) return;

  // All dropdowns EXCEPT the theme picker (it has its own open state already)
  const dropdowns = Array.from(
    bar.querySelectorAll(".dropdown:not(.theme-picker)")
  );

  // Clicking a menu button toggles that one; closes others
  dropdowns.forEach((dd) => {
    const btn = dd.querySelector(".menu-btn");
    if (!btn) return;
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const isOpen = dd.classList.contains("open");
      dropdowns.forEach((d) => d.classList.remove("open"));
      if (!isOpen) dd.classList.add("open");
    });
  });

  // Click outside closes all
  document.addEventListener("click", () => {
    dropdowns.forEach((d) => d.classList.remove("open"));
  });

  // ESC closes all
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape")
      dropdowns.forEach((d) => d.classList.remove("open"));
  });
}

// Clear all data button
async function clearAllData() {
  if (!confirm("Are you sure you want to clear ALL stored tenant data?"))
    return;

  await Promise.all([
    chrome.storage.local.remove([
      storageKey,
      "hermes_myapis",
      "hermes_preferences",
    ]),
    chrome.storage.session.remove(sessionStorageKey),
  ]);

  // Stop timers and reset UI
  const accessTokenTimerBox = document.getElementById("timer");
  const refreshTokenTimerBox = document.getElementById("refresh-timer");
  stopAccessTokenTimer(accessTokenTimerBox);
  stopRefreshTokenTimer(refreshTokenTimerBox);

  await populateClientID();
  await populateAccessToken();
  await populateClientSecret();
  await populateTenantId();
  await populateRefreshToken();
  await restoreTokenTimers();
}

// Clear client data button
async function clearClientData() {
  const clienturl = await getClientUrl();
  if (!clienturl) {
    alert("No valid Tenant detected.");
    return;
  }

  if (!confirm(`Are you sure you want to clear data for: ${clienturl}?`))
    return;

  const data = await loadClientData();
  if (data[clienturl]) {
    delete data[clienturl];
    await saveClientData(data);
    alert(`Tenant data cleared for: ${clienturl}`);

    // stop timers and reset ui
    const accessTokenTimerBox = document.getElementById("timer");
    const refreshTokenTimerBox = document.getElementById("refresh-timer");
    stopAccessTokenTimer(accessTokenTimerBox);
    stopRefreshTokenTimer(refreshTokenTimerBox);

    await populateClientID();
    await populateAccessToken();
    await populateClientSecret();
    await populateTenantId();
    await populateRefreshToken();
    await restoreTokenTimers();
  } else {
    alert("No data found for this Tenant.");
  }
}
// ====================================== //

// ===== MENU BAR | LINKS FUNCTIONS ===== //
// Developer portal
async function linksDeveloperPortal() {
  try {
    const panelMeta = await fetch("accesspanel.json").then((res) => res.json());
    const developerPortalTemplate = panelMeta?.details?.urls?.developerPortal;

    if (!developerPortalTemplate) {
      console.error("Developer Portal URL not found in accesspanel.json.");
      return;
    }

    const developerPortalURL = resolvePanelUrl(developerPortalTemplate, "");

    if (!developerPortalURL) {
      console.error("Developer Portal URL could not be resolved.");
      return;
    }

    const incognito = await isIncognitoContext();
    if (incognito) {
      chrome.tabs.create({ url: developerPortalURL, active: true });
    } else {
      openURLNormally(developerPortalURL);
    }
  } catch (error) {
    console.error("Failed to load Developer Portal URL:", error);
  }
}

// Boomi button
async function linksBoomi() {
  if (!(await isValidSession())) {
    alert("Requires a valid ADP Workforce Manager session.");
    return;
  }

  const clienturl = await getClientUrl();
  if (!clienturl) return;

  let boomiTemplate = null;
  try {
    const cfg = await fetch("accesspanel.json").then((res) => res.json());
    boomiTemplate = cfg?.details?.urls?.boomiPortal || null;
  } catch (e) {
    console.error("Failed to load AccessPanel config:", e);
  }

  if (!boomiTemplate) {
    alert("Boomi link is not configured.");
    return;
  }

  const ssoClientUrl = createSsoUrl(clienturl);
  const boomiURL = resolvePanelUrl(boomiTemplate, ssoClientUrl);

  if (!boomiURL) {
    alert("Boomi link could not be resolved.");
    return;
  }

  const incognito = await isIncognitoContext();
  if (incognito) {
    chrome.tabs.create({ url: boomiURL, active: true });
  } else {
    openURLNormally(boomiURL);
  }
}

// Install Integrations button
async function linksInstallIntegrations() {
  if (!(await isValidSession())) {
    alert("Requires a valid ADP Workforce Manager session.");
    return;
  }

  const clienturl = await getClientUrl();
  if (!clienturl) return;

  let installTemplate = null;
  try {
    const cfg = await fetch("accesspanel.json").then((res) => res.json());
    installTemplate = cfg?.details?.urls?.installIntegrations || null;
  } catch (e) {
    console.error("Failed to load AccessPanel config:", e);
  }

  if (!installTemplate) {
    alert("Install Integrations link is not configured.");
    return;
  }

  const ssoClientUrl = createSsoUrl(clienturl);
  const installIntegrationsURL = resolvePanelUrl(installTemplate, ssoClientUrl);

  if (!installIntegrationsURL) {
    alert("Install Integrations link could not be resolved.");
    return;
  }

  const incognito = await isIncognitoContext();
  if (incognito) {
    chrome.tabs.create({ url: installIntegrationsURL, active: true });
  } else {
    openURLNormally(installIntegrationsURL);
  }
}
// ====================================== //

// ===== MENU BAR | THEMES FUNCTIONS ===== //
// Load themes from themes.json
async function loadThemes() {
  try {
    const response = await fetch("themes/themes.json");
    if (!response.ok)
      throw new Error(
        `Failed to fetch themes. HTTP status: ${response.status}`
      );
    const themesData = await response.json();
    return themesData.themes;
  } catch (error) {
    console.error("Error loading themes:", error);
    return {};
  }
}

// Populate themes dropdown
async function populateThemeDropdown() {
  const themes = await loadThemes();
  const dropdown = document.getElementById("theme-selector");
  if (!dropdown) {
    console.error("Theme dropdown element not found in DOM.");
    return;
  }

  for (const themeKey in themes) {
    const theme = themes[themeKey];
    const option = document.createElement("option");
    option.value = themeKey;
    option.textContent = theme.name;
    dropdown.appendChild(option);
  }
}

// Apply the selected theme
async function applyTheme(themeKey) {
  const themes = await loadThemes();
  const selectedTheme = themes[themeKey];

  if (!selectedTheme) {
    console.warn(`Theme "${themeKey}" not found.`);
    return;
  }

  const root = document.documentElement;
  const colors = selectedTheme.colors;

  // update color variables
  for (const [key, value] of Object.entries(colors)) {
    root.style.setProperty(`--${key}`, value);
  }

  // update font variables
  const fonts = selectedTheme.fonts;
  root.style.setProperty("--font-family", fonts["font-family-primary"]);
  root.style.setProperty("--title-font", fonts["title-font-primary"]);

  // save the selected theme in local storage
  chrome.storage.local.set({ selectedTheme: themeKey });
}

// Theme selection
function themeSelection(event) {
  const selectedTheme = event.target.value;
  applyTheme(selectedTheme);
}

// Restore the last selected theme on load
async function restoreSelectedTheme() {
  chrome.storage.local.get("selectedTheme", async (result) => {
    const themeKey = result.selectedTheme || "accessPanel";
    await applyTheme(themeKey);

    const dropdown = document.getElementById("theme-selector");
    if (dropdown) dropdown.value = themeKey;
  });
}

// Build themes menu from select
function buildThemeMenuFromSelect() {
  const select = document.getElementById("theme-selector");
  const menu = document.getElementById("theme-menu");
  if (!select || !menu) return;

  menu.replaceChildren();
  [...select.options].forEach((opt) => {
    if (!opt.value) return;
    const btn = document.createElement("button");
    btn.className = "theme-item";
    btn.type = "button";
    btn.dataset.theme = opt.value;
    btn.textContent = opt.textContent;
    menu.appendChild(btn);
  });
}

// Delegate clicks in the themes menu
function wireThemeMenuClicks() {
  const menu = document.getElementById("theme-menu");
  if (!menu) return;

  menu.addEventListener("click", (e) => {
    const item = e.target.closest(".theme-item");
    if (!item) return;

    const select = document.getElementById("theme-selector");
    if (select) {
      select.value = item.dataset.theme;
      select.dispatchEvent(new Event("change", { bubbles: true }));
    }

    // Close the dropdown using the unified controller
    const dd = document.getElementById("themes-dropdown");
    const btn = document.getElementById("theme-menu-btn");
    if (dd) dd.classList.remove("open");
    if (btn) btn.setAttribute("aria-expanded", "false");
  });
}
// ======================================= //

// ===== MENU BAR | HELP FUNCTIONS ===== //
// About button
async function helpAbout() {
  try {
    const cfg = await fetch("accesspanel.json").then((res) => res.json());

    const name = cfg?.name ?? "AccessPanel";
    const description = (cfg?.details?.description ?? "")
      .replace(/\s+/g, " ")
      .trim();
    const version = cfg?.details?.version ?? "";
    const releaseDate = cfg?.details?.release_date ?? "";
    const author = cfg?.details?.author ?? "";
    const descMax = 120;
    const descShort =
      description.length > descMax
        ? description.slice(0, descMax - 1) + "…"
        : description;

    const message = [
      `Name: ${name}`,
      `Description: ${descShort}`,
      `Version: ${version}`,
      `Release Date: ${releaseDate}`,
      `Author: ${author}`,
    ].join("\n");

    alert(message);
  } catch (error) {
    console.error("Failed to load About information:", error);
    alert("Failed to load About information.");
  }
}

// Support: open GitHub Issues
async function helpSupport() {
  try {
    const cfg = await fetch("accesspanel.json").then((res) => res.json());
    const issuesTemplate = cfg?.details?.urls?.reportIssues;

    if (!issuesTemplate) {
      alert("Support link is not configured.");
      return;
    }

    const ok = confirm(
      "Open AccessPanel support (GitHub Issues) in a new tab?"
    );
    if (!ok) return;

    const issuesUrl = resolvePanelUrl(issuesTemplate, "");
    if (!issuesUrl) {
      alert("Support link could not be resolved.");
      return;
    }

    // Consistent behavior with the rest of the app
    openURLNormally(issuesUrl);
  } catch (error) {
    console.error("Failed to open support link:", error);
    alert("Failed to open support link.");
  }
}


// ===== CLIENT DATA STORAGE FUNCTIONS ===== //
// Load client data from persistent + session storage
async function loadClientData() {
  const [localResult, sessionResult] = await Promise.all([
    chrome.storage.local.get(storageKey),
    chrome.storage.session.get(sessionStorageKey),
  ]);

  const localData = localResult[storageKey] || {};
  const sessionData = sessionResult[sessionStorageKey] || {};

  const mergedData = {};
  const tenantKeys = new Set([
    ...Object.keys(localData),
    ...Object.keys(sessionData),
  ]);

  for (const tenantKey of tenantKeys) {
    mergedData[tenantKey] = {
      ...(localData[tenantKey] || {}),
      ...(sessionData[tenantKey] || {}),
    };
  }

  return mergedData;
}

// Save client data, separating persistent configuration
// from session-only authentication data
async function saveClientData(data) {
  const localData = {};
  const sessionData = {};

  for (const [tenantKey, clientData] of Object.entries(data || {})) {
    if (!clientData || typeof clientData !== "object") continue;

    const persistentClientData = {};
    const sessionClientData = {};

    for (const [field, value] of Object.entries(clientData)) {
      if (SESSION_ONLY_CLIENT_FIELDS.has(field)) {
        sessionClientData[field] = value;
      } else {
        persistentClientData[field] = value;
      }
    }

    if (Object.keys(persistentClientData).length > 0) {
      localData[tenantKey] = persistentClientData;
    }

    if (Object.keys(sessionClientData).length > 0) {
      sessionData[tenantKey] = sessionClientData;
    }
  }

  await Promise.all([
    chrome.storage.local.set({
      [storageKey]: localData,
    }),
    chrome.storage.session.set({
      [sessionStorageKey]: sessionData,
    }),
  ]);
}
// =================================== //

// ===== MAIN UI HELPERS ===== //
// Button success text temporary
function setButtonTempText(
  btn,
  okText,
  ms = 2000,
  originalText = btn.textContent
) {
  if (!btn) return;

  const isIcony =
    btn.classList.contains("icon-btn") || btn.querySelector("img");

  if (isIcony) {
    const origTitle = btn.getAttribute("title") || "";
    btn.setAttribute("title", okText);
    btn.classList.add("flash-ok");
    btn.disabled = true;

    setTimeout(() => {
      btn.classList.remove("flash-ok");
      btn.setAttribute("title", origTitle);
      btn.disabled = false;
    }, ms);
  } else {
    btn.textContent = okText;
    btn.disabled = true;
    setTimeout(() => {
      btn.textContent = originalText;
      btn.disabled = false;
    }, ms);
  }
}

// Button fail text temporary
function setButtonFailText(
  btn,
  failText,
  ms = 2000,
  originalText = btn.textContent
) {
  if (!btn) return;

  const isIcony =
    btn.classList.contains("icon-btn") || btn.querySelector("img");

  if (isIcony) {
    const origTitle = btn.getAttribute("title") || "";
    btn.setAttribute("title", failText);
    btn.classList.add("flash-fail");

    setTimeout(() => {
      btn.classList.remove("flash-fail");
      btn.setAttribute("title", origTitle);
    }, ms);
  } else {
    btn.textContent = failText;
    btn.disabled = false;
    setTimeout(() => {
      btn.textContent = originalText;
    }, ms);
  }
}

// Button hourglass animation
function startLoadingAnimation(button) {
  const hourglassFrames = ["⏳", "⌛"];
  let frameIndex = 0;
  let rotationAngle = 0;

  // store original text for later restoration
  const originalText = button.textContent;

  button.replaceChildren();

  const waitingText = document.createTextNode("Waiting... ");
  const hourglassSpan = document.createElement("span");

  hourglassSpan.className = "hourglass";
  hourglassSpan.textContent = hourglassFrames[frameIndex];
  hourglassSpan.style.display = "inline-block";

  button.append(waitingText, hourglassSpan);
  button.disabled = true;

  return {
    interval: setInterval(() => {
      frameIndex = (frameIndex + 1) % hourglassFrames.length;
      rotationAngle += 30;
      hourglassSpan.textContent = hourglassFrames[frameIndex];
      hourglassSpan.style.transform = `rotate(${rotationAngle}deg)`;
    }, 100),
    originalText, // return this to be used later
  };
}

// Restore token timers
async function restoreTokenTimers() {
  const clienturl = await getClientUrl();
  if (!clienturl) {
    const accessTokenTimerBox = document.getElementById("timer");
    const refreshTokenTimerBox = document.getElementById("refresh-timer");

    // reset timers in the UI
    stopAccessTokenTimer(accessTokenTimerBox);
    stopRefreshTokenTimer(refreshTokenTimerBox);
    return;
  }

  const data = await loadClientData();
  const clientData = data[clienturl] || {};
  const currentDateTime = new Date();

  // Restore access token status/timer
  const accessTokenTimerBox = document.getElementById("timer");

  if (clientData.accesstoken) {
    if (clientData.accessTokenSource === "manual") {
      stopAccessTokenTimer(accessTokenTimerBox);
      accessTokenTimerBox.textContent = "Manual";
    } else {
      const expirationTime = new Date(clientData.expirationdatetime);

      if (
        clientData.expirationdatetime &&
        Number.isFinite(expirationTime.getTime()) &&
        currentDateTime < expirationTime
      ) {
        const remainingSeconds = Math.floor(
          (expirationTime - currentDateTime) / 1000
        );

        startAccessTokenTimer(remainingSeconds, accessTokenTimerBox);
      } else {
        accessTokenTimerBox.textContent = "--:--";
      }
    }
  } else {
    accessTokenTimerBox.textContent = "--:--";
  }

  // restore refresh token timer
  const refreshTokenTimerBox = document.getElementById("refresh-timer");
  if (clientData.refreshtoken) {
    const refreshExpirationTime = new Date(
      clientData.refreshExpirationDateTime
    );
    if (currentDateTime < refreshExpirationTime) {
      const remainingSeconds = Math.floor(
        (refreshExpirationTime - currentDateTime) / 1000
      );
      startRefreshTokenTimer(remainingSeconds, refreshTokenTimerBox);
    } else {
      refreshTokenTimerBox.textContent = "--:--";
    }
  } else {
    refreshTokenTimerBox.textContent = "--:--";
  }
}
// =========================== //

// ===== CLIENT URL/ID FIELDS AND BUTTONS ===== //
// Toggle Tenant Information section
function toggleTenantSection() {
  const toggleButton = document.getElementById("toggle-tenant-section");
  const content = document.getElementById("tenant-section-content");

  if (!toggleButton || !content) return;

  const expanded = !content.classList.contains("expanded");
  content.classList.toggle("expanded", expanded);

  toggleButton.textContent = expanded
    ? "▲ Hide Tenant Information ▲"
    : "▼ Show Tenant Information ▼";

  chrome.storage.local.set({ tenantSectionExpanded: expanded });
}

// Restore Tenant Information section state
function restoreTenantSection() {
  chrome.storage.local.get("tenantSectionExpanded", (result) => {
    const isExpanded = !!result.tenantSectionExpanded;
    const toggleButton = document.getElementById("toggle-tenant-section");
    const content = document.getElementById("tenant-section-content");

    if (!toggleButton || !content) return;

    content.classList.toggle("expanded", isExpanded);

    toggleButton.textContent = isExpanded
      ? "▲ Hide Tenant Information ▲"
      : "▼ Show Tenant Information ▼";
  });
}

// Generate BIRT Properties Button
async function generateBirtPropertiesClick() {
  const btn = document.getElementById("generate-birt-file");
  if (!btn) return;

  const originalLabel = btn.textContent || "Generate BIRT Properties";

  const setLabel = (text, autoResetMs = null) => {
    btn.textContent = text;

    if (autoResetMs) {
      setTimeout(() => {
        btn.textContent = originalLabel;
      }, autoResetMs);
    }
  };

  try {
    setLabel("Generating...");

    // --- 1. Validate required fields ---
    const clientId =
      document.getElementById("client-id")?.value.trim() || "";
    const clientSecret =
      document.getElementById("client-secret")?.value.trim() || "";
    const tenantId =
      document.getElementById("tenant-id")?.value.trim() || "";

    if (!clientId || !clientSecret || !tenantId) {
      alert(
        [
          !clientId ? "- Client ID is required." : "",
          !clientSecret ? "- Client Secret is required." : "",
          !tenantId ? "- Tenant ID is required." : "",
        ]
          .filter(Boolean)
          .join("\n")
      );

      setLabel("Error", 1500);
      return;
    }

    // --- 2. Get current client URL and vanity host ---
    const clientUrl = await getClientUrl();

    if (!clientUrl) {
      alert(
        "No valid client URL detected. Make sure AccessPanel is linked to a tenant."
      );
      setLabel("Error", 1500);
      return;
    }

    let vanityHost = "";

    try {
      const urlObj = new URL(clientUrl);
      vanityHost = urlObj.hostname || "";

      if (!vanityHost || !vanityHost.includes("mykronos.com")) {
        throw new Error("Not a mykronos.com vanity URL");
      }
    } catch (error) {
      console.error("Failed to parse vanity URL from Tenant URL.", error);

      alert(
        "Unable to parse a valid vanity hostname from the Tenant URL.\n" +
          "Expected something like https://<tenant>.mykronos.com"
      );

      setLabel("Error", 1500);
      return;
    }

    // --- 3. Request a fresh token ---
    const tokenRequested = await fetchToken();

    if (!tokenRequested) {
      setLabel("Error", 1500);
      return;
    }

    // --- 4. Wait for access + refresh tokens ---
    const waitForTokens = async (
      clientUrlKey,
      maxWaitMs = 12000,
      intervalMs = 250
    ) => {
      const start = Date.now();

      while (Date.now() - start < maxWaitMs) {
        const data = await loadClientData();
        const client = data[clientUrlKey] || {};

        if (client.accesstoken && client.refreshtoken) {
          return client;
        }

        await new Promise((resolve) => setTimeout(resolve, intervalMs));
      }

      const data = await loadClientData();
      return data[clientUrlKey] || {};
    };

    const thisClient = await waitForTokens(clientUrl);

    const accessToken = thisClient.accesstoken || "";
    const refreshToken = thisClient.refreshtoken || "";

    if (!accessToken || !refreshToken) {
      alert(
        "Could not retrieve access/refresh token after requesting one.\n\n" +
          "Tip: Ensure you are actively logged into WFM in this browser context, then try again."
      );

      setLabel("Error", 1500);
      return;
    }

    const editDate = thisClient.editdatetime || new Date().toISOString();

    // --- 5. Build BIRT properties file contents ---
    const propsLines = [
      "report.api.execute.for.external.client=true",
      "report.api.gateway.access.token.appkey=",
      `volume_name=${tenantId}`,
      `report.api.access.token.qparam.client.id=${clientId}`,
      `access.token=${accessToken}`,
      `report.api.access.token.qparam.client.secret=${clientSecret}`,
      `refresh.token=${refreshToken}`,
      "report.api.gateway.access.token.authchain=OAuthLdapService",
      `report.api.env.vanity.url=${vanityHost}`,
      "report.api.gateway.access.token.is.ssl.enable=true",
    ];

    const propsText = propsLines.join("\n");

    window.lastBirtPropertiesText = propsText;
    window.lastBirtPropertiesMeta = {
      clientUrl,
      tenantId,
      editDate,
    };

    // --- 6. Open BIRT Properties popup ---
    const colors = getPopupThemeColors();

    const popupCss = `
      body {
        font-family: Arial, sans-serif;
        margin: 0;
        padding: 16px;
        background-color: ${colors.primary};
        line-height: 1.5;
      }

      h1 {
        font-size: 1.4rem;
        font-weight: bold;
        margin: 0 0 0.25rem 0;
        color: ${colors.accent};
      }

      .meta {
        font-size: 0.85rem;
        font-weight: bold;
        color: ${colors.accent};
        margin-bottom: 12px;
      }

      .btn-row {
        margin-bottom: 10px;
      }

      button {
        font-family: inherit;
        font-size: 0.9rem;
        padding: 6px 12px;
        margin-right: 8px;
        border-radius: 4px;
        border: 1px solid ${colors.secondary};
        background-color: ${colors.accent};
        color: ${colors.buttonText};
        cursor: pointer;
      }

      button:hover {
        background-color: ${colors.highlight};
      }

      pre {
        background: ${colors.buttonText};
        border: 2px solid ${colors.accent};
        border-radius: 6px;
        padding: 10px;
        white-space: pre;
        overflow-x: auto;
        font-family: ui-monospace, SFMono-Regular, Menlo, Monaco,
          Consolas, "Liberation Mono", "Courier New", monospace;
        font-size: 0.9rem;
        color: ${colors.accent};
      }
    `;

    const popup = createPopupWindow(
      "BIRT Properties",
      "width=800,height=600,scrollbars=yes,resizable=yes",
      popupCss
    );

    if (!popup) {
      alert(
        "Unable to open BIRT popup window. Please allow popups for this extension."
      );
      setLabel("Error", 1500);
      return;
    }

    const doc = popup.document;

    const heading = doc.createElement("h1");
    heading.textContent = "BIRT Properties";

    const meta = doc.createElement("div");
    meta.className = "meta";

    const tenantLine = doc.createElement("div");
    tenantLine.textContent = `Tenant: ${tenantId}`;

    const urlLine = doc.createElement("div");
    urlLine.textContent = `Client URL: ${clientUrl}`;

    const editLine = doc.createElement("div");
    editLine.textContent = `Last Edit: ${editDate}`;

    meta.append(tenantLine, urlLine, editLine);

    const buttonRow = doc.createElement("div");
    buttonRow.className = "btn-row";

    const copyBtn = doc.createElement("button");
    copyBtn.type = "button";
    copyBtn.textContent = "Copy To Clipboard";

    const downloadBtn = doc.createElement("button");
    downloadBtn.type = "button";
    downloadBtn.textContent = "Download To File";

    buttonRow.append(copyBtn, downloadBtn);

    const pre = doc.createElement("pre");
    pre.textContent = propsText;

    doc.body.append(heading, meta, buttonRow, pre);

    // --- 7. Copy BIRT properties ---
    copyBtn.addEventListener("click", async () => {
      try {
        popup.focus();

        const text = pre.textContent || "";

        if (!text.trim()) {
          copyBtn.textContent = "No Content";

          setTimeout(() => {
            copyBtn.textContent = "Copy To Clipboard";
          }, 1500);

          return;
        }

        await popup.navigator.clipboard.writeText(text);

        copyBtn.textContent = "Copied!";

        setTimeout(() => {
          copyBtn.textContent = "Copy To Clipboard";
        }, 1500);
      } catch (error) {
        console.error("Copy BIRT properties failed:", error);

        copyBtn.textContent = "Copy Failed";

        setTimeout(() => {
          copyBtn.textContent = "Copy To Clipboard";
        }, 1500);
      }
    });

    // --- 8. Download BIRT properties ---
    downloadBtn.addEventListener("click", async () => {
      try {
        const text = pre.textContent || "";
        const defaultFileName = "custom_reportplugin.properties";

        if (!text.trim()) {
          downloadBtn.textContent = "No Content";

          setTimeout(() => {
            downloadBtn.textContent = "Download To File";
          }, 1500);

          return;
        }

        if (popup.showSaveFilePicker) {
          try {
            const fileHandle = await popup.showSaveFilePicker({
              suggestedName: defaultFileName,
              types: [
                {
                  description: "Properties Files",
                  accept: {
                    "text/plain": [".properties"],
                  },
                },
              ],
            });

            const writable = await fileHandle.createWritable();
            await writable.write(text);
            await writable.close();

            return;
          } catch (error) {
            if (error?.name === "AbortError") {
              return;
            }

            // Fall through to standard download
          }
        }

        downloadFileWithContext(
          popup.document,
          popup.URL,
          defaultFileName,
          text,
          "text/plain"
        );
      } catch (error) {
        console.error("Download BIRT properties failed:", error);
        alert("Failed to download the BIRT properties file.");
      }
    });

    setLabel("Generated", 1500);
  } catch (error) {
    console.error("Generate BIRT Properties failed:", error);
    alert("Failed to generate BIRT properties.");
    setLabel("Error", 1500);
  }
}

// Populate the API access client URL field
async function populateClientUrlField() {
  try {
    const input = document.getElementById("client-url");
    if (!input) return;

    const base = await getClientUrl();
    input.value = base ? toApiUrl(base) : "";
  } catch (e) {
    console.error("populateClientUrlField failed:", e);
  }
}

// Refresh URL button
async function refreshClientUrlClick() {
  const btn = document.getElementById("refresh-client-url");
  const originalText = btn?.textContent || "Refresh";

  try {
    await populateClientUrlField();

    const val = document.getElementById("client-url")?.value || "";
    if (val) {
      setButtonTempText(btn, "URL Refreshed", 2000, originalText);
    } else {
      setButtonFailText(btn, "No URL Detected", 2000, originalText);
    }
  } catch (e) {
    console.error(e);
    setButtonFailText(btn, "Refresh Failed", 2000, originalText);
  }
}

// Copy URL button
async function copyClientUrlClick() {
  const btn = document.getElementById("copy-client-url");
  const originalText = btn?.textContent || "Copy";

  try {
    const val = document.getElementById("client-url")?.value || "";
    if (!val) {
      setButtonFailText(btn, "No URL to Copy", 2000, originalText);
      return;
    }

    await navigator.clipboard.writeText(val);

    setButtonTempText(btn, "URL Copied", 2000, originalText);
  } catch (e) {
    console.error("Copy failed:", e);
    setButtonFailText(btn, "Copy Failed", 2000, originalText);
  }
}

// Populate API Access Client ID Field
async function populateClientID() {
  const clienturl = await getClientUrl();
  const clientIDBox = document.getElementById("client-id");

  if (!clienturl) {
    clientIDBox.value = "";
    clientIDBox.placeholder = "Requires WFMgr Login";
    clientIDBox.readOnly = true;
    return;
  }

  const data = await loadClientData();
  if (data[clienturl]?.clientid) {
    clientIDBox.value = data[clienturl].clientid;
    clientIDBox.placeholder = "";
  } else {
    clientIDBox.value = "";
    clientIDBox.placeholder = "Enter Client ID";
  }
}

// Save API access client ID button
async function saveClientIDClick() {
  const button = document.getElementById("save-client-id");
  const originalText = button?.textContent || "Save";

  try {
    if (!(await isValidSession())) {
      alert("Requires a valid ADP Workforce Manager session.");
      return;
    }

    const clienturl = await getClientUrl();
    if (!clienturl) {
      setButtonFailText(button, "No Client URL", 2000, originalText);
      return;
    }

    const clientid = document.getElementById("client-id")?.value?.trim() || "";
    if (!clientid) {
      setButtonFailText(button, "Client ID Empty", 2000, originalText);
      return;
    }

    const data = await loadClientData();
    data[clienturl] = {
      ...(data[clienturl] || {}),
      clientid: clientid,
      tokenurl: `${clienturl}accessToken?clientId=${clientid}`,
      apiurl: `${clienturl}api`,
      editdatetime: new Date().toISOString(),
    };

    await saveClientData(data);
    setButtonTempText(button, "Client ID Saved", 2000, originalText);
  } catch (error) {
    console.error("Failed to save Client ID:", error);
    setButtonFailText(button, "Save Failed", 2000, originalText);
  }
}

// Copy client ID button
async function copyClientIdClick() {
  const btn = document.getElementById("copy-client-id");
  const originalText = btn?.textContent || "Copy";

  try {
    const val = document.getElementById("client-id")?.value?.trim() || "";
    if (!val) {
      setButtonFailText(btn, "No Client ID", 2000, originalText);
      return;
    }

    await navigator.clipboard.writeText(val);

    setButtonTempText(btn, "Client ID Copied", 2000, originalText);
  } catch (e) {
    console.error("Copy Client ID failed:", e);
    setButtonFailText(btn, "Copy Failed", 2000, originalText);
  }
}

// Populate client secret box
async function populateClientSecret() {
  const clienturl = await getClientUrl();
  const clientSecretBox = document.getElementById("client-secret");
  if (!clienturl) {
    clientSecretBox.value = "";
    clientSecretBox.placeholder = "Requires WFMgr Login";
    clientSecretBox.readOnly = true;
    return;
  }

  const data = await loadClientData();
  if (data[clienturl]?.clientsecret) {
    clientSecretBox.value = data[clienturl].clientsecret;
    clientSecretBox.placeholder = "";
  } else {
    clientSecretBox.value = "";
    clientSecretBox.placeholder = "Enter Client Secret";
  }
}

// Toggle client secret visibility
function toggleClientSecretVisibility() {
  const clientSecretBox = document.getElementById("client-secret");
  const toggleIcon = document.getElementById("toggle-client-secret");

  if (clientSecretBox.type === "password") {
    clientSecretBox.type = "text";
    toggleIcon.src = "icons/eyeclosed.png";
  } else {
    clientSecretBox.type = "password";
    toggleIcon.src = "icons/eyeopen.png";
  }
}

// Save client secret button
async function saveClientSecretClick() {
  const button = document.getElementById("save-client-secret");
  const originalText = button?.textContent || "Save";

  try {
    if (!(await isValidSession())) {
      alert("Requires a valid ADP Workforce Manager session.");
      return;
    }

    const clienturl = await getClientUrl();
    if (!clienturl) {
      setButtonFailText(button, "No Client URL", 2000, originalText);
      return;
    }

    const clientsecret =
      document.getElementById("client-secret")?.value?.trim() || "";
    if (!clientsecret) {
      setButtonFailText(button, "Secret Empty", 2000, originalText);
      return;
    }

    const data = await loadClientData();
    data[clienturl] = {
      ...(data[clienturl] || {}),
      clientsecret: clientsecret,
      editdatetime: new Date().toISOString(),
    };

    await saveClientData(data);
    setButtonTempText(button, "Secret Saved", 2000, originalText);
  } catch (error) {
    console.error("Failed to save Client Secret:", error);
    setButtonFailText(button, "Save Failed", 2000, originalText);
  }
}

// Copy client secret button
async function copyClientSecretClick() {
  const btn = document.getElementById("copy-client-secret");
  const originalText = btn?.textContent || "Copy";

  try {
    const val = document.getElementById("client-secret")?.value?.trim() || "";
    if (!val) {
      setButtonFailText(btn, "No Client Secret", 2000, originalText);
      return;
    }

    await navigator.clipboard.writeText(val);

    setButtonTempText(btn, "Client Secret Copied", 2000, originalText);
  } catch (e) {
    console.error("Copy Client Secret failed:", e);
    setButtonFailText(btn, "Copy Failed", 2000, originalText);
  }
}

// Populate tenant ID field
async function populateTenantId() {
  const tenantIdInput = document.getElementById("tenant-id");
  if (!tenantIdInput) return;

  const clienturl = await getClientUrl();

  // No WFM session or no client URL yet
  if (!clienturl) {
    tenantIdInput.value = "";
    tenantIdInput.placeholder = "Requires WFMgr Login";
    tenantIdInput.readOnly = true;
    return;
  }

  const data = await loadClientData();
  const clientData = data[clienturl];

  tenantIdInput.readOnly = false;

  if (clientData && clientData.tenantid) {
    tenantIdInput.value = clientData.tenantid;
    tenantIdInput.placeholder = "";
  } else {
    tenantIdInput.value = "";
    tenantIdInput.placeholder = "Enter Tenant ID";
  }
}

// Save Tenant ID Button
async function saveTenantIdClick() {
  const button = document.getElementById("save-tenant-id");
  const originalText = button?.textContent || "Save";

  try {
    if (!(await isValidSession())) {
      alert("Requires a valid ADP Workforce Manager session.");
      return;
    }

    const clienturl = await getClientUrl();
    if (!clienturl) {
      setButtonFailText(button, "No Client URL", 2000, originalText);
      return;
    }

    const tenantIdInput = document.getElementById("tenant-id");
    if (!tenantIdInput) {
      console.error("tenant-id input not found.");
      setButtonFailText(button, "Field Missing", 2000, originalText);
      return;
    }

    const tenantId = tenantIdInput.value.trim();
    if (!tenantId) {
      setButtonFailText(button, "Tenant ID Empty", 2000, originalText);
      return;
    }

    const data = await loadClientData();
    data[clienturl] = {
      ...(data[clienturl] || {}),
      tenantid: tenantId,
      editdatetime: new Date().toISOString(),
    };

    await saveClientData(data);
    setButtonTempText(button, "Tenant ID Saved", 2000, originalText);
  } catch (error) {
    console.error("Failed to save Tenant ID:", error);
    setButtonFailText(button, "Save Failed", 2000, originalText);
  }
}

// Copy tenant ID button
async function copyTenantIdClick() {
  const btn = document.getElementById("copy-tenant-id");
  const originalText = btn?.textContent || "Copy";

  try {
    const val = document.getElementById("tenant-id")?.value?.trim() || "";
    if (!val) {
      setButtonFailText(btn, "No Tenant ID", 2000, originalText);
      return;
    }

    await navigator.clipboard.writeText(val);

    setButtonTempText(btn, "Tenant ID Copied", 2000, originalText);
  } catch (e) {
    console.error("Copy Tenant ID failed:", e);
    setButtonFailText(btn, "Copy Failed", 2000, originalText);
  }
}
// ==================================== //

// ===== API TOKEN UI FIELDS AND BUTTONS ===== //
// Access token section collapsed
function toggleAccessSection() {
  const toggleButton = document.getElementById("toggle-access-section");
  const content = document.getElementById("access-section-content");

  if (!toggleButton || !content) return;

  const expanded = !content.classList.contains("expanded");
  content.classList.toggle("expanded", expanded);

  toggleButton.textContent = expanded
    ? "▲ Hide API Token Options ▲"
    : "▼ Show API Token Options ▼";

  chrome.storage.local.set({ accessSectionExpanded: expanded });
}


function restoreAccessSection() {
  chrome.storage.local.get("accessSectionExpanded", (result) => {
    const isExpanded = !!result.accessSectionExpanded;
    const toggleButton = document.getElementById("toggle-access-section");
    const content = document.getElementById("access-section-content");

    if (!toggleButton || !content) return;

    content.classList.toggle("expanded", isExpanded);

    toggleButton.textContent = isExpanded
      ? "▲ Hide API Token Options ▲"
      : "▼ Show API Token Options ▼";
  });
}

// Populate access token field and restore its status
async function populateAccessToken() {
  const clienturl = await getClientUrl();
  const accessTokenBox = document.getElementById("access-token");
  const timerBox = document.getElementById("timer");

  if (!accessTokenBox || !timerBox) return;

  if (!clienturl) {
    accessTokenBox.value = "Requires WFMgr Login";
    timerBox.textContent = "--:--";
    return;
  }

  const data = await loadClientData();
  const clientData = data[clienturl] || {};

  if (!clientData.accesstoken) {
    accessTokenBox.value = "Get New Access Token";
    timerBox.textContent = "--:--";
    return;
  }

  accessTokenBox.value = clientData.accesstoken;

  // Manual tokens have no known expiration time.
  if (clientData.accessTokenSource === "manual") {
    stopAccessTokenTimer(timerBox);
    timerBox.textContent = "Manual";
    return;
  }

  const expirationTime = new Date(clientData.expirationdatetime);
  const currentDateTime = new Date();

  if (
    !clientData.expirationdatetime ||
    !Number.isFinite(expirationTime.getTime()) ||
    currentDateTime > expirationTime
  ) {
    accessTokenBox.value = "Access Token Expired";
    timerBox.textContent = "--:--";
    return;
  }

  const remainingSeconds = Math.floor(
    (expirationTime - currentDateTime) / 1000
  );

  startAccessTokenTimer(remainingSeconds, timerBox);
}

// Get access token button (returns true if token retrieval was initiated)
async function fetchToken() {
  const clienturl = await getClientUrl();

  if (!clienturl) {
    alert(
      "Client URL is required. Please refresh or set the Client URL first."
    );
    return false;
  }

  if (!(await isValidSession())) {
    alert("Requires a valid ADP Workforce Manager session.");
    return false;
  }

  const clientID =
    document.getElementById("client-id")?.value?.trim() || "";

  if (!clientID) {
    alert("Please enter a Client ID first.");
    return false;
  }

  // Access token retrieval is intentionally unsupported in incognito.
  // The previous incognito implementation required script injection into
  // a temporary token page, which AccessPanel no longer uses.
  if (await isIncognitoContext()) {
    alert(
      "Access token retrieval is not supported in an Incognito window.\n\n" +
        "Please use AccessPanel from a normal browser session, or use your " +
        "standard manual API access method while working in Incognito."
    );
    return false;
  }

  const tokenurl = `${clienturl}accessToken?clientId=${clientID}`;

  return await fetchTokenDirectly(tokenurl, clienturl, clientID);
}

// Get access token normal mode (used by fetchToken())
async function fetchTokenDirectly(tokenurl, clienturl, clientID) {
  try {
    const response = await fetch(tokenurl, {
      method: "GET",
      credentials: "include",
      headers: { Accept: "application/json" },
    });

    if (!response.ok) {
      alert(`Failed to fetch token. HTTP status: ${response.status}`);
      return false;
    }

    const result = await response.json();
    await processTokenResponse(result, clienturl, clientID, tokenurl);
    return true;
  } catch (error) {
    // This is a real runtime/network failure, so keeping console.error is appropriate
    console.error("Error fetching token:", error);
    alert(`Failed to fetch token: ${error.message || error}`);
    return false;
  }
}

// Process Token (used by fetchTokenDirectly())
async function processTokenResponse(result, baseClientUrl, clientID, tokenurl) {
  const button = document.getElementById("get-token");

  try {
    // Defensive parsing + validation (avoid truthy traps)
    const accessToken = result?.accessToken;
    const refreshToken = result?.refreshToken;
    const expiresInSeconds = Number(result?.expiresInSeconds);

    if (
      !accessToken ||
      !refreshToken ||
      !Number.isFinite(expiresInSeconds) ||
      expiresInSeconds <= 0
    ) {
      console.warn(
        "Token response is missing required fields or has invalid expiry."
      );
      alert("Failed to fetch token: Invalid response.");
      setButtonFailText(button, "Token Failed!");
      return;
    }

    const now = new Date();

    // Access token expiry comes from API response
    const accessExp = new Date(now.getTime() + expiresInSeconds * 1000);

    // Refresh token expiry is derived (API does not provide it)
    const refreshExp = new Date(now.getTime() + 8 * 60 * 60 * 1000); // 8 hours

    // baseClientUrl is already normalized (e.g., https://foo-nossosomething/ )
    const data = await loadClientData();

    data[baseClientUrl] = {
      ...(data[baseClientUrl] || {}),
      clientid: clientID,
      tokenurl, // keep full /accessToken?clientId=... for later use
      apiurl: data[baseClientUrl]?.apiurl || `${baseClientUrl}api`,
      accesstoken: accessToken,
      accessTokenSource: "automatic",
      refreshtoken: refreshToken,

      effectivedatetime: now.toISOString(),
      expirationdatetime: accessExp.toISOString(),
      refreshExpirationDateTime: refreshExp.toISOString(),
      editdatetime: now.toISOString(),
    };

    await saveClientData(data);

    // UI updates
    populateAccessToken();
    populateRefreshToken();
    restoreTokenTimers();
    setButtonTempText(button, "Token Retrieved!");
  } catch (error) {
    console.error("Failed to process token response:", error);
    setButtonFailText(button, "Token Failed!");
  }
}

// Apply a manually entered access token to the current browser session
async function applyManualAccessTokenClick() {
  const button = document.getElementById("apply-token");
  const originalText = button?.textContent || "Apply Token";

  const accessTokenBox = document.getElementById("access-token");
  const timerBox = document.getElementById("timer");

  const token = accessTokenBox?.value?.trim() || "";

  const invalidValues = new Set([
    "",
    "Get Token",
    "Get New Access Token",
    "Access Token Expired",
    "Requires WFMgr Login",
  ]);

  if (invalidValues.has(token)) {
    setButtonFailText(button, "No Token", 2000, originalText);
    return;
  }

  const clienturl = await getClientUrl();

  if (!clienturl) {
    alert(
      "Client URL is required. Please refresh or set the Client URL first."
    );
    setButtonFailText(button, "No Client URL", 2000, originalText);
    return;
  }

  try {
    const data = await loadClientData();

    const clientData = {
      ...(data[clienturl] || {}),
      accesstoken: token,
      accessTokenSource: "manual",
      editdatetime: new Date().toISOString(),
    };

    // A manually supplied token has no expiration or refresh metadata
    // that AccessPanel can reliably determine.
    delete clientData.effectivedatetime;
    delete clientData.expirationdatetime;
    delete clientData.refreshtoken;
    delete clientData.refreshExpirationDateTime;

    data[clienturl] = clientData;

    await saveClientData(data);

    stopAccessTokenTimer(timerBox);
    stopRefreshTokenTimer(document.getElementById("refresh-timer"));

    await chrome.storage.session.remove([
      "accessTokenTimer",
      "refreshTokenTimer",
    ]);

    timerBox.textContent = "Manual";

    await populateRefreshToken();

    setButtonTempText(button, "Token Applied", 2000, originalText);
  } catch (error) {
    console.error("Failed to apply manual Access Token:", error);
    setButtonFailText(button, "Apply Failed", 2000, originalText);
  }
}

// Start timer for access token
function startAccessTokenTimer(seconds, timerBox) {
  if (accessTokenTimerInterval) {
    clearInterval(accessTokenTimerInterval);
    accessTokenTimerInterval = null;
  }

  let remainingTime = seconds;

  const updateTimer = () => {
    if (remainingTime <= 0) {
      clearInterval(accessTokenTimerInterval);
      accessTokenTimerInterval = null;
      timerBox.textContent = "--:--";

      // clear the remaining time in storage
      chrome.storage.session.remove("accessTokenTimer");
    } else {
      const minutes = Math.floor(remainingTime / 60);
      const seconds = remainingTime % 60;
      timerBox.textContent = `${String(minutes).padStart(2, "0")}:${String(
        seconds
      ).padStart(2, "0")}`;
      remainingTime--;

      // save the remaining time to storage
      chrome.storage.session.set({
        accessTokenTimer: remainingTime,
      });
    }
  };

  // update the timer immediately and then every second
  updateTimer();
  accessTokenTimerInterval = setInterval(updateTimer, 1000);
}

// Stop access token timer
function stopAccessTokenTimer(timerBox) {
  if (accessTokenTimerInterval) {
    clearInterval(accessTokenTimerInterval);
    accessTokenTimerInterval = null;
    timerBox.textContent = "--:--"; // reset the timer box
  }
}

// Stop all token timers
function stopAllTokenTimers() {
  // access-token timer
  const atEl = document.getElementById("timer");
  if (atEl && typeof stopAccessTokenTimer === "function") {
    stopAccessTokenTimer(atEl);
  }
  // refresh-token timer (if you have one)
  const rtEl = document.getElementById("refresh-timer");
  if (rtEl && typeof stopRefreshTokenTimer === "function") {
    stopRefreshTokenTimer(rtEl);
  }
}

// Resume all token timers from storage
async function resumeTokenTimersFromStorage() {
  // repull the current tokens and timers from storage, then restart timers
  await populateAccessToken();
  await populateRefreshToken();
  await restoreTokenTimers();
}

// Copy access token button
function copyAccessToken() {
  const button = document.getElementById("copy-token");
  const originalText = button?.textContent || "Copy";

  const accessTokenBox = document.getElementById("access-token");
  const accessToken = accessTokenBox?.value;

  // validate access token before copying
  if (
    !accessToken ||
    accessToken === "Get Token" ||
    accessToken === "Get New Access Token" ||
    accessToken === "Access Token Expired"
  ) {
    // User clicked copy but there is no valid token
    setButtonFailText(button, "No Token", 2000, originalText);
    return;
  }

  navigator.clipboard
    .writeText(accessToken)
    .then(() => {
      setButtonTempText(button, "Copied!", 2000, originalText);
    })
    .catch((error) => {
      console.error("Failed to copy Access Token:", error);
      setButtonFailText(button, "Copy Failed", 2000, originalText);
    });
}

// Populate refresh token box
async function populateRefreshToken() {
  const clienturl = await getClientUrl();
  const refreshTokenBox = document.getElementById("refresh-token");
  const refreshTimerBox = document.getElementById("refresh-timer");

  if (!clienturl) {
    refreshTokenBox.value = "Requires WFMgr Login";
    refreshTimerBox.textContent = "--:--";
    return;
  }

  const data = await loadClientData();
  const currentDateTime = new Date();

  if (data[clienturl]?.refreshtoken) {
    const refreshExpirationTime = new Date(
      data[clienturl].refreshExpirationDateTime
    );

    if (currentDateTime > refreshExpirationTime) {
      refreshTokenBox.value = "Refresh Token Expired";
      refreshTimerBox.textContent = "--:--";
    } else {
      refreshTokenBox.value = data[clienturl].refreshtoken;

      // calculate remaining time and start the timer
      const remainingSeconds = Math.floor(
        (refreshExpirationTime - currentDateTime) / 1000
      );
      startRefreshTokenTimer(remainingSeconds, refreshTimerBox);
    }
  } else {
    refreshTokenBox.value = "Get New Access Token";
    refreshTimerBox.textContent = "--:--";
  }
}

// Refresh access token using refresh token button
async function refreshAccessToken() {
  const button = document.getElementById("refresh-access-token");

  try {
    const clienturl = await getClientUrl();
    if (!clienturl || !(await isValidSession())) {
      alert("Requires a valid ADP Workforce Manager session.");
      return;
    }

    const data = await loadClientData();
    const client = data[clienturl] || {};
    const { refreshtoken, clientid, clientsecret } = client;

    // validate refresh token
    if (
      !refreshtoken ||
      refreshtoken === "Refresh Token Expired" ||
      new Date() > new Date(client.refreshExpirationDateTime)
    ) {
      alert(
        "No valid Refresh Token found. Please retrieve an Access Token first."
      );
      setButtonFailText(button, "No Valid Token!");
      return;
    }

    // validate client secret
    if (!clientsecret || clientsecret === "Enter Client Secret") {
      alert("Client Secret is required to refresh the Access Token.");
      setButtonFailText(button, "Missing Secret!");
      return;
    }

    const apiurl = `${clienturl}api/authentication/access_token`;

    // make POST request
    const response = await fetch(apiurl, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: new URLSearchParams({
        refresh_token: refreshtoken,
        client_id: clientid,
        client_secret: clientsecret,
        grant_type: "refresh_token",
        auth_chain: "OAuthLdapService",
      }),
    });

    if (!response.ok) {
      throw new Error(
        `Failed to refresh access token. HTTP status: ${response.status}`
      );
    }

    // parse response
    const result = await response.json();

    const { access_token, expires_in } = result;
    if (!access_token || !expires_in) {
      throw new Error(
        "Response is missing required fields: 'access_token' or 'expires_in'."
      );
    }

    // calculate expiration time
    const currentDateTime = new Date();
    const accessTokenExpirationDateTime = new Date(
      currentDateTime.getTime() + expires_in * 1000
    );

    // update client data storage
    data[clienturl] = {
      ...client,
      accesstoken: access_token,
      expirationdatetime: accessTokenExpirationDateTime.toISOString(),
      editdatetime: currentDateTime.toISOString(),
    };

    await saveClientData(data);

    // update the UI
    populateAccessToken();
    restoreTokenTimers();

    setButtonTempText(button, "Token Refreshed!");
  } catch (error) {
    console.error("Error refreshing access token:", error.message);
    alert(`Failed to refresh access token: ${error.message}`);
    setButtonFailText(button, "Refresh Failed!");
  }
}

// Start timer for refresh token
function startRefreshTokenTimer(seconds, timerBox) {
  if (refreshTokenTimerInterval) {
    clearInterval(refreshTokenTimerInterval);
    refreshTokenTimerInterval = null;
  }

  let remainingTime = seconds;

  const updateTimer = () => {
    if (remainingTime <= 0) {
      clearInterval(refreshTokenTimerInterval);
      refreshTokenTimerInterval = null;
      timerBox.textContent = "--:--";
    } else {
      const minutes = Math.floor(remainingTime / 60);
      const seconds = remainingTime % 60;
      timerBox.textContent = `${String(minutes).padStart(2, "0")}:${String(
        seconds
      ).padStart(2, "0")}`;
      remainingTime--;

      // save the remaining time to storage
      chrome.storage.session.set({
        refreshTokenTimer: remainingTime,
      });
    }
  };

  updateTimer();
  refreshTokenTimerInterval = setInterval(updateTimer, 1000);
}

// Stop refresh token timer
function stopRefreshTokenTimer(timerBox) {
  if (refreshTokenTimerInterval) {
    clearInterval(refreshTokenTimerInterval);
    refreshTokenTimerInterval = null;
    timerBox.textContent = "--:--";
  }
}

// Copy refresh token button
function copyRefreshToken() {
  const button = document.getElementById("copy-refresh-token");
  const originalText = button?.textContent || "Copy";

  const refreshTokenBox = document.getElementById("refresh-token");
  const refreshToken = refreshTokenBox?.value;

  // validate refresh token before copying
  if (
    !refreshToken ||
    refreshToken === "Refresh Token" ||
    refreshToken === "Get New Access Token" ||
    refreshToken === "Refresh Token Expired"
  ) {
    setButtonFailText(button, "No Token", 2000, originalText);
    return;
  }

  // copy refresh token to clipboard
  navigator.clipboard
    .writeText(refreshToken)
    .then(() => {
      setButtonTempText(button, "Copied!", 2000, originalText);
    })
    .catch((error) => {
      console.error("Failed to copy Refresh Token:", error);
      setButtonFailText(button, "Copy Failed", 2000, originalText);
    });
}

// Purge expired tokens from storage (access + refresh) without touching other tenant metadata
async function purgeExpiredTokensInStorage() {
  const data = await loadClientData();
  const now = new Date();

  let changed = false;

  for (const [baseUrl, client] of Object.entries(data || {})) {
    if (!client || typeof client !== "object") continue;

    // Access token cleanup (expirationdatetime)
    if (client.accesstoken && client.expirationdatetime) {
      const exp = new Date(client.expirationdatetime);
      if (!isNaN(exp) && now > exp) {
        delete client.accesstoken;
        delete client.expirationdatetime;
        changed = true;
      }
    }

    // Refresh token cleanup (refreshExpirationDateTime)
    if (client.refreshtoken && client.refreshExpirationDateTime) {
      const rexp = new Date(client.refreshExpirationDateTime);
      if (!isNaN(rexp) && now > rexp) {
        delete client.refreshtoken;
        delete client.refreshExpirationDateTime;
        changed = true;
      }
    }

    data[baseUrl] = client;
  }

  if (changed) {
    await saveClientData(data);
  }
}
// =============================================== //

// ===== BODY JSON FONT CONTROLS For API Library ===== //
const ADHOC_BODY_FONT_KEY = "hermes_adhoc_body_font_px";

// Guard body font size within reasonable limits
function clampBodyFont(px) {
  return Math.max(10, Math.min(28, px)); // 10–28px range
}

// Get body textarea element
function getBodyTextarea() {
  return document.getElementById("adhoc-body");
}

// Get body font size label element
function getBodyFontLabel() {
  return document.getElementById("body-font-size-label");
}

// Apply font size to body textarea and label, and store in localStorage
function applyBodyFontSize(px) {
  const ta = getBodyTextarea();
  const label = getBodyFontLabel();
  if (!ta || !label) return;

  const v = clampBodyFont(px);
  ta.style.fontSize = `${v}px`;
  label.textContent = `${v}px`;
  try {
    localStorage.setItem(ADHOC_BODY_FONT_KEY, String(v));
  } catch (e) {
    // ignore
  }
}

// Initialize body font size from localStorage
function initBodyFontSizeFromStorage() {
  const ta = getBodyTextarea();
  if (!ta) return;
  let px = 13;
  try {
    const stored = localStorage.getItem(ADHOC_BODY_FONT_KEY);
    if (stored) px = parseInt(stored, 10) || 13;
  } catch (e) {
    // ignore
  }
  applyBodyFontSize(px);
}

// Adjust body font size by delta
function adjustBodyFontSize(delta) {
  const ta = getBodyTextarea();
  if (!ta) return;
  const current = parseInt(getComputedStyle(ta).fontSize, 10) || 13;
  applyBodyFontSize(current + delta);
}

// Click handlers increase (logic is one line each)
function onBodyFontIncreaseClick() {
  adjustBodyFontSize(+1);
}

// Click handlers decrease (logic is one line each)
function onBodyFontDecreaseClick() {
  adjustBodyFontSize(-1);
}

// Ensure the A− / A+ toolbar exists next to the "Full JSON Body" header. `headerEl` should be the <div class="parameter-header"> for the body.
function ensureBodyFontControls(headerEl) {
  const ta = getBodyTextarea();
  if (!ta || !headerEl) return;

  // already built?
  if (document.getElementById("body-font-toolbar")) {
    initBodyFontSizeFromStorage();
    return;
  }

  const bar = document.createElement("div");
  bar.id = "body-font-toolbar";
  bar.className = "body-font-toolbar";

  const dec = document.createElement("button");
  dec.type = "button";
  dec.id = "body-font-decrease";
  dec.className = "btn3";
  dec.textContent = "A−";

  const size = document.createElement("span");
  size.id = "body-font-size-label";
  size.style.margin = "0 .5rem";

  const inc = document.createElement("button");
  inc.type = "button";
  inc.id = "body-font-increase";
  inc.className = "btn3";
  inc.textContent = "A+";

  //pPut toolbar inside the header, on the right
  headerEl.appendChild(bar);
  bar.append(dec, size, inc);

  // wire listeners (no inline logic)
  dec.addEventListener("click", onBodyFontDecreaseClick);
  inc.addEventListener("click", onBodyFontIncreaseClick);

  // apply initial size from storage
  initBodyFontSizeFromStorage();
}
// =============================================== //

// ===== API LIBRARY FUNCTIONS ===== //
// Toggle API Library section
function toggleApiLibrary() {
  const toggleButton = document.getElementById("toggle-api-library");
  const content = document.getElementById("api-library-content");

  if (!toggleButton || !content) return;

  const expanded = !content.classList.contains("expanded");
  content.classList.toggle("expanded", expanded);

  toggleButton.textContent = expanded
    ? "▲ Hide API Library ▲"
    : "▼ Show API Library ▼";

  chrome.storage.local.set({ apiLibraryExpanded: expanded });
}

// Restore API Library section state
function restoreApiLibrary() {
  chrome.storage.local.get("apiLibraryExpanded", (result) => {
    const isExpanded = !!result.apiLibraryExpanded;
    const toggleButton = document.getElementById("toggle-api-library");
    const content = document.getElementById("api-library-content");

    if (!toggleButton || !content) return;

    content.classList.toggle("expanded", isExpanded);

    toggleButton.textContent = isExpanded
      ? "▲ Hide API Library ▲"
      : "▼ Show API Library ▼";
  });
}

// Clear DevLink
function clearDevLinkBanner() {
  const a = document.getElementById("api-devlink");
  if (!a) return;
  a.hidden = true;
  a.removeAttribute("href");
  a.textContent = "";
}

// Populate API Library dropdown
async function populateApiDropdown() {
  const sel = document.getElementById("api-selector");
  if (!sel) return;

  sel.replaceChildren();
  clearDevLinkBanner();

  // Placeholder
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Select API...";
  placeholder.disabled = true;
  placeholder.selected = true;
  sel.appendChild(placeholder);

  // Ad-Hoc requests
  const adhocGet = document.createElement("option");
  adhocGet.value = "adHocGet";
  adhocGet.textContent = "Ad-Hoc GET";
  sel.appendChild(adhocGet);

  const adhocPost = document.createElement("option");
  adhocPost.value = "adHocPost";
  adhocPost.textContent = "Ad-Hoc POST";
  sel.appendChild(adhocPost);

  const separator = document.createElement("option");
  separator.textContent = "────────";
  separator.disabled = true;
  sel.appendChild(separator);

  // Public API Library
  const response = await fetch("apilibrary/apilibrary.json");
  if (!response.ok) {
    throw new Error(`Failed to fetch API library. HTTP ${response.status}`);
  }

  const data = await response.json();
  const library = data.apiLibrary || {};

  for (const key in library) {
    if (
      key.startsWith("_") ||
      key === "adHocGet" ||
      key === "adHocPost"
    ) {
      continue;
    }

    const option = document.createElement("option");
    option.value = key;
    option.textContent = library[key].name;
    sel.appendChild(option);
  }
}

// Load API public library from apilibrary.json
async function loadApiLibrary() {
  try {
    const response = await fetch("apilibrary/apilibrary.json");
    if (!response.ok)
      throw new Error(`Failed to load API Library: ${response.status}`);
    const apiLibraryData = await response.json();
    return apiLibraryData.apiLibrary;
  } catch (error) {
    console.error("Error loading API Library:", error);
    return {};
  }
}

// Clear existing API parameters
function clearParameters() {
  const queryContainer = document.getElementById(
    "query-parameters-container"
  );
  const bodyContainer = document.getElementById(
    "body-parameters-container"
  );
  const pathContainer = document.getElementById(
    "path-parameters-container"
  );

  queryContainer?.replaceChildren();
  bodyContainer?.replaceChildren();
  pathContainer?.replaceChildren();
}

// Show or clear the Developer Portal link for the selected public API
function renderDevLinkBanner(selectedApiKey, apiObjOrNull) {
  const a = document.getElementById("api-devlink");
  if (!a) return;

  // hide for ad-hoc and for missing devlink data
  if (
    !apiObjOrNull ||
    selectedApiKey === "adHocGet" ||
    selectedApiKey === "adHocPost" ||
    !apiObjOrNull.devLink ||
    !apiObjOrNull.devLink.url ||
    !apiObjOrNull.devLink.urlText
  ) {
    a.hidden = true;
    a.removeAttribute("href");
    a.textContent = "";
    return;
  }

  // show populated dev link
  a.href = apiObjOrNull.devLink.url;
  a.textContent = apiObjOrNull.devLink.urlText;
  a.hidden = false;
}

// Date-time helper
function pad2(n) {
  return String(n).padStart(2, "0");
}

// Format local date-time as YYYY-MM-DDTHH:mm
function formatLocalYMDHM(date = new Date()) {
  const y = date.getFullYear();
  const m = pad2(date.getMonth() + 1);
  const d = pad2(date.getDate());
  const hh = pad2(date.getHours());
  const mm = pad2(date.getMinutes());
  return `${y}-${m}-${d}T${hh}:${mm}`;
}

// Parse relative offsets like "-1" (days), "+3h" (hours), "-90m" (minutes) */
function applyRelativeDateTimeOffset(base, spec) {
  // spec examples: "-1", "+2", "+3h", "-90m"
  const s = String(spec).trim();
  const m = s.match(/^([+-]?\d+)([dhm])?$/i);
  if (!m) return base;

  const val = parseInt(m[1], 10);
  const unit = (m[2] || "d").toLowerCase(); // default to days

  const dt = new Date(base.getTime());
  if (unit === "d") dt.setDate(dt.getDate() + val);
  else if (unit === "h") dt.setHours(dt.getHours() + val);
  else if (unit === "m") dt.setMinutes(dt.getMinutes() + val);
  return dt;
}

// Render path parameters
function renderPathParamRow(param) {
  const wrap = document.createElement("div");
  wrap.className = "query-param-wrapper";

  const label = document.createElement("label");
  label.textContent = `${param.name}:`;
  label.setAttribute("for", `path-${param.name}`);
  wrap.appendChild(label);

  const inputType = (param.type || "text").toLowerCase();

  const input = document.createElement("input");
  input.type = inputType === "integer" ? "number" : "text";
  input.id = `path-${param.name}`;
  input.className = "query-param-input";
  input.dataset.name = param.name;

  if (param.defaultValue !== undefined && param.defaultValue !== "") {
    input.value = String(param.defaultValue);
  } else {
    input.placeholder = param.description || "Enter value";
  }

  wrap.appendChild(input);
  return wrap;
}

// Populate the path parameters area
async function populatePathParameters(selectedApiKey) {
  const host = document.getElementById("path-parameters-container");
  if (!host) return;

  host.replaceChildren();

  // Ad-Hoc requests do not use path parameter UI
  if (
    selectedApiKey === "adHocGet" ||
    selectedApiKey === "adHocPost"
  ) {
    return;
  }

  const apiLibrary = await loadApiLibrary();
  const api = apiLibrary[selectedApiKey];
  if (!api) return;

  const list = api.pathParameters || [];
  if (!list.length) return;

  const header = document.createElement("div");
  header.className = "parameter-header";
  header.textContent = "Path Parameters";
  host.appendChild(header);

  if (api.pathParametersHelp) {
    const help = document.createElement("p");
    help.className = "parameter-help-text";
    help.textContent = api.pathParametersHelp;
    host.appendChild(help);
  }

  list.forEach((p) => host.appendChild(renderPathParamRow(p)));
}

// Replace {name} tokens in a url template with values from the ui
function buildUrlWithPathParams(urlTemplate, apiDef) {
  let url = String(urlTemplate || "");
  const list = apiDef?.pathParameters || [];
  if (!list.length) return url;

  // collect values
  const values = {};
  list.forEach((p) => {
    const el = document.getElementById(`path-${p.name}`);
    values[p.name] = (el?.value ?? "").trim();
  });

  // replace tokens {name}
  url = url.replace(/\{([^}]+)\}/g, (match, name) => {
    const v = values[name];
    // if missing → leave token for now; the guard below will catch it
    return v !== undefined && v !== "" ? encodeURIComponent(v) : match;
  });

  // guard: if any placeholders remain, fail gracefully
  if (/\{[^}]+\}/.test(url)) {
    throw new Error("One or more path parameters are missing.");
  }

  return url;
}

// Populate query parameters
async function populateQueryParameters(selectedApiKey) {
  try {
    const apiLibrary = await loadApiLibrary(); // load the api library
    const selectedApi = apiLibrary[selectedApiKey] || null;

    renderDevLinkBanner?.(selectedApiKey, selectedApi);

    if (!selectedApi) {
      console.error("Selected API not found in the library.");
      return;
    }

    const queryContainer = document.getElementById(
      "query-parameters-container"
    );
    queryContainer.replaceChildren();

    if (selectedApiKey === "adHocGet" || selectedApiKey === "adHocPost") {
      const queryHeader = document.createElement("div");
      queryHeader.className = "parameter-header";
      queryHeader.textContent = "Endpoint URL with Query Parameters";
      queryContainer.appendChild(queryHeader);

      const endpointInput = document.createElement("input");
      endpointInput.type = "text";
      endpointInput.id = "adhoc-endpoint";
      endpointInput.classList.add("query-param-input");
      endpointInput.placeholder = "/v1/endpoint?queryParam=value";
      queryContainer.appendChild(endpointInput);
      return;
    }

    const params = selectedApi.queryParameters || [];
    if (!params.length) {
      return;
    }

    // header
    const queryHeader = document.createElement("div");
    queryHeader.className = "parameter-header";
    queryHeader.textContent = "Query Parameters";
    queryContainer.appendChild(queryHeader);

    // optional help text
    if (selectedApi.queryParametersHelp) {
      const queryHelpText = document.createElement("p");
      queryHelpText.className = "parameter-help-text";
      queryHelpText.textContent = selectedApi.queryParametersHelp;
      queryContainer.appendChild(queryHelpText);
    }

    params.forEach((param) => {
      const wrap = document.createElement("div");
      wrap.classList.add("query-param-wrapper");

      const label = document.createElement("label");
      label.textContent = `${param.name}:`;
      label.setAttribute("for", `query-${param.name}`);
      wrap.appendChild(label);

      const type = (param.type || "").toLowerCase();

      // --- SELECT (supports {label,value} or string options) ---
      if (type === "select") {
        const sel = document.createElement("select");
        sel.id = `query-${param.name}`;
        sel.classList.add("query-param-input");

        // placeholder
        const ph = document.createElement("option");
        ph.value = "";
        ph.textContent = param.description || "";
        ph.disabled = true;
        ph.selected = true;
        ph.hidden = true;
        sel.appendChild(ph);

        // label/value options
        const opts = normalizeOptions(param.options);
        opts.forEach(({ label: lbl, value }) => {
          const o = document.createElement("option");
          o.value = String(value); // API value
          o.textContent = lbl; // UI label
          sel.appendChild(o);
        });

        // placeholder styling behavior
        sel.classList.add("placeholder");
        sel.addEventListener("change", () => {
          if (sel.value === "") sel.classList.add("placeholder");
          else sel.classList.remove("placeholder");
        });

        if (param.defaultValue !== undefined && param.defaultValue !== "") {
          sel.value = String(param.defaultValue);
          if (sel.value !== "") sel.classList.remove("placeholder");
        }

        wrap.appendChild(sel);
        queryContainer.appendChild(wrap);
        return;
      }

      // --- BOOLEAN (right-aligned select, values "true"/"false") ---
      if (type === "boolean") {
        const sel = document.createElement("select");
        sel.id = `query-${param.name}`;
        sel.classList.add("query-param-input");

        const ph = document.createElement("option");
        ph.value = "";
        ph.textContent = param.description || "";
        ph.disabled = true;
        ph.selected = true;
        ph.hidden = true;
        sel.appendChild(ph);

        // use provided options if present, else [true,false]
        const boolOpts =
          param.options && param.options.length
            ? param.options.map((v) => String(v).toLowerCase() === "true")
            : [true, false];

        boolOpts.forEach((val) => {
          const o = document.createElement("option");
          o.value = val ? "true" : "false"; // API value
          o.textContent = val ? "True" : "False";
          sel.appendChild(o);
        });

        sel.classList.add("placeholder");
        sel.addEventListener("change", () => {
          if (sel.value === "") sel.classList.add("placeholder");
          else sel.classList.remove("placeholder");
        });

        if (
          typeof param.defaultValue !== "undefined" &&
          param.defaultValue !== ""
        ) {
          const dv =
            typeof param.defaultValue === "boolean"
              ? param.defaultValue
                ? "true"
                : "false"
              : String(param.defaultValue).toLowerCase();
          if (dv === "true" || dv === "false") {
            sel.value = dv;
            sel.classList.remove("placeholder");
          }
        }

        wrap.appendChild(sel);
        queryContainer.appendChild(wrap);
        return;
      }

      // --- DATE (supports relative offsets like "-1") ---
      if (type === "date") {
        const input = document.createElement("input");
        input.type = "date";
        input.id = `query-${param.name}`;
        input.classList.add("query-param-input");

        if (
          typeof param.defaultValue === "string" &&
          /^[+-]?\d+$/.test(param.defaultValue)
        ) {
          const daysOffset = parseInt(param.defaultValue, 10);
          input.value = formatLocalYMD(daysOffset); // local date (no UTC drift)
        } else if (param.defaultValue) {
          input.value = param.defaultValue;
        } else {
          // leave empty; placeholder text can come from CSS if desired
        }

        wrap.appendChild(input);
        queryContainer.appendChild(wrap);
        return;
      }

      // --- DATETIME (local, with optional relative defaults) ---
      if (type === "datetime") {
        const input = document.createElement("input");
        input.type = "datetime-local";
        input.id = `query-${param.name}`;
        input.classList.add("query-param-input");

        const dv = param.defaultValue;
        if (typeof dv === "string" && dv) {
          // relative offsets like "-1", "+3h", "-90m"
          if (/^[+-]?\d+(?:[dhm])?$/i.test(dv)) {
            const dt = applyRelativeDateTimeOffset(new Date(), dv);
            input.value = formatLocalYMDHM(dt);
          } else if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(dv)) {
            // direct ISO-like string
            input.value = dv;
          } else {
            input.placeholder = param.description || "YYYY-MM-DDTHH:mm";
          }
        } else {
          input.placeholder = param.description || "YYYY-MM-DDTHH:mm";
        }

        wrap.appendChild(input);
        queryContainer.appendChild(wrap);
        return;
      }

      // --- DEFAULT: plain text ---
      const input = document.createElement("input");
      input.type = "text";
      input.id = `query-${param.name}`;
      input.classList.add("query-param-input");

      if (param.defaultValue !== undefined && param.defaultValue !== "") {
        input.value = param.defaultValue;
      } else {
        input.placeholder = param.description || "Enter value";
      }

      wrap.appendChild(input);
      queryContainer.appendChild(wrap);
    });
  } catch (error) {
    console.error("Error populating query parameters:", error);
  }
}

// Populate body parameters
async function populateBodyParameters(selectedApiKey) {
  try {
    const apiLibrary = await loadApiLibrary();
    const selectedApi = apiLibrary[selectedApiKey];

    const bodyParamContainer = document.getElementById(
      "body-parameters-container"
    );
    if (!bodyParamContainer) {
      console.warn("Body Parameters container not found.");
      return;
    }

    bodyParamContainer.replaceChildren();

    // Ad-hoc POST: full JSON textarea
    if (selectedApiKey === "adHocPost") {
      const bodyHeader = document.createElement("div");
      bodyHeader.className = "parameter-header";
      bodyHeader.textContent = "Full JSON Body";
      bodyParamContainer.appendChild(bodyHeader);

      const textarea = document.createElement("textarea");
      textarea.id = "adhoc-body";
      textarea.className = "json-textarea";
      textarea.placeholder = "Enter full JSON body here...";
      bodyParamContainer.appendChild(textarea);
      ensureBodyFontControls(bodyHeader);
      return;
    }

    if (selectedApi?.method === "GET") return;
    if (!selectedApi || !selectedApi.bodyParameters) return;

    // header
    const bodyHeader = document.createElement("div");
    bodyHeader.className = "parameter-header";
    bodyHeader.textContent = "Body Parameters";
    bodyParamContainer.appendChild(bodyHeader);

    // optional help text
    if (selectedApi.bodyParametersHelp) {
      const bodyHelpText = document.createElement("p");
      bodyHelpText.className = "parameter-help-text";
      bodyHelpText.textContent = selectedApi.bodyParametersHelp;
      bodyParamContainer.appendChild(bodyHelpText);
    }

    // build each parameter row
    selectedApi.bodyParameters.forEach((param) => {
      const paramWrapper = document.createElement("div");
      paramWrapper.className = "body-param-wrapper";

      // label
      let labelText = param.name;
      if (param.type === "multi-text" && param.validation?.maxEntered) {
        labelText += ` (max = ${param.validation.maxEntered})`;
      }
      const label = document.createElement("label");
      label.htmlFor = `body-param-${param.name}`;
      label.textContent = labelText;
      label.className = "body-param-label";
      paramWrapper.appendChild(label);

      // branch per type
      if (param.type === "multi-select") {
        // CHECKBOX LIST (supports label/value via normalizeOptions)
        const multiSelectContainer = document.createElement("div");
        multiSelectContainer.className = "multi-select-container";

        const opts = normalizeOptions(param.options);
        opts.forEach(({ label, value }) => {
          const checkboxWrapper = document.createElement("div");
          checkboxWrapper.className = "checkbox-wrapper";

          const checkbox = document.createElement("input");
          checkbox.type = "checkbox";
          checkbox.value = String(value); // API value
          checkbox.dataset.path = param.path;
          checkbox.dataset.type = param.type;
          checkbox.id = `body-param-${param.name}-${value}`;

          // defaultValue can be array or single
          const def = param.defaultValue;
          if (Array.isArray(def) && def.map(String).includes(String(value))) {
            checkbox.checked = true;
          } else if (
            typeof def !== "undefined" &&
            String(def) === String(value)
          ) {
            checkbox.checked = true;
          }

          const checkboxLabel = document.createElement("label");
          checkboxLabel.htmlFor = checkbox.id;
          checkboxLabel.textContent = label; // UI label

          checkboxWrapper.appendChild(checkbox);
          checkboxWrapper.appendChild(checkboxLabel);
          multiSelectContainer.appendChild(checkboxWrapper);
        });

        paramWrapper.appendChild(multiSelectContainer);
      } else if (param.type === "multi-text") {
        // STACKED MULTI-TEXT (unchanged)
        paramWrapper.style.display = "block";
        const multiTextContainer = document.createElement("div");
        multiTextContainer.className = "multi-text-container";

        const addButton = document.createElement("button");
        addButton.type = "button";
        addButton.className = "btn btn-add-item";
        addButton.textContent = "Add Entry";
        addButton.addEventListener("click", () => {
          const textInput = document.createElement("input");
          textInput.type = "text";
          textInput.className = "body-param-input";
          textInput.dataset.path = param.path;
          textInput.dataset.type = param.type;
          textInput.placeholder = param.description || "Enter value";
          multiTextContainer.appendChild(textInput);
        });
        multiTextContainer.appendChild(addButton);

        const defaultTextInput = document.createElement("input");
        defaultTextInput.type = "text";
        defaultTextInput.className = "body-param-input";
        defaultTextInput.dataset.path = param.path;
        defaultTextInput.dataset.type = param.type;
        defaultTextInput.placeholder = param.description || "Enter value";
        multiTextContainer.appendChild(defaultTextInput);

        paramWrapper.appendChild(multiTextContainer);
      } else if (param.type === "select") {
        // SINGLE SELECT with label/value support
        paramWrapper.classList.add("body-select-wrapper");

        const dropdown = document.createElement("select");
        dropdown.className = "body-param-input body-select-input";
        dropdown.dataset.path = param.path;
        dropdown.dataset.type = param.type;
        dropdown.id = `body-param-${param.name}`;

        // Placeholder
        const placeholderOption = document.createElement("option");
        placeholderOption.value = "";
        placeholderOption.textContent =
          param.description || "";
        placeholderOption.disabled = true;
        placeholderOption.selected = true;
        placeholderOption.hidden = true;
        dropdown.appendChild(placeholderOption);

        // Options via normalizeOptions
        const opts = normalizeOptions(param.options);
        opts.forEach(({ label, value }) => {
          const optEl = document.createElement("option");
          optEl.value = String(value); // API value
          optEl.textContent = label; // UI label
          dropdown.appendChild(optEl);
        });

        // Default selection (value)
        if (
          typeof param.defaultValue !== "undefined" &&
          param.defaultValue !== ""
        ) {
          dropdown.value = String(param.defaultValue);
        }

        paramWrapper.appendChild(dropdown);
      } else if (param.type === "boolean") {
        // BOOLEAN SELECT (true/false options, aligned right)
        paramWrapper.classList.add("body-boolean-wrapper");

        const dropdown = document.createElement("select");
        dropdown.className = "body-param-input body-boolean-input";
        dropdown.dataset.path = param.path;
        dropdown.dataset.type = param.type;
        dropdown.id = `body-param-${param.name}`;

        // Placeholder
        const placeholderOption = document.createElement("option");
        placeholderOption.value = "";
        placeholderOption.textContent =
          param.description || "";
        placeholderOption.disabled = true;
        placeholderOption.selected = true;
        placeholderOption.hidden = true;
        dropdown.appendChild(placeholderOption);

        // Render true/false; respect custom options if provided
        const boolOptions =
          param.options && param.options.length
            ? param.options.map((v) => String(v).toLowerCase() === "true")
            : [true, false];

        boolOptions.forEach((val) => {
          const opt = document.createElement("option");
          opt.value = val ? "true" : "false"; // API value (string)
          opt.textContent = val ? "True" : "False"; // UI label
          dropdown.appendChild(opt);
        });

        // Default value: accept "true"/"false" or boolean true/false
        if (
          typeof param.defaultValue !== "undefined" &&
          param.defaultValue !== ""
        ) {
          const dv =
            typeof param.defaultValue === "boolean"
              ? param.defaultValue
                ? "true"
                : "false"
              : String(param.defaultValue).toLowerCase();
          if (dv === "true" || dv === "false") dropdown.value = dv;
        }

        paramWrapper.appendChild(dropdown);
      } else if (param.type === "date") {
        // DATE (local offset logic preserved)
        const input = document.createElement("input");
        input.type = "date";
        input.id = `body-param-${param.name}`;
        input.className = "body-param-input";
        input.dataset.path = param.path;
        input.dataset.type = param.type;

        if (param.defaultValue === "") {
          input.placeholder = param.description || "mm/dd/yyyy";
          input.classList.add("placeholder-style");
        } else if (
          typeof param.defaultValue === "string" &&
          /^[+-]?\d+$/.test(param.defaultValue)
        ) {
          const daysOffset = parseInt(param.defaultValue, 10);
          input.value = formatLocalYMD(daysOffset); // local date
        } else if (param.defaultValue) {
          input.value = param.defaultValue;
        }

        paramWrapper.appendChild(input);
      } else if (param.type === "datetime") {
        // DATETIME-LOCAL (same behavior as date, plus time)
        const input = document.createElement("input");
        input.type = "datetime-local";
        input.id = `body-param-${param.name}`;
        input.className = "body-param-input";
        input.dataset.path = param.path;
        input.dataset.type = param.type;

        const dv = param.defaultValue;
        if (dv === "") {
          input.placeholder = param.description || "YYYY-MM-DDTHH:mm";
          input.classList.add("placeholder-style");
        } else if (typeof dv === "string" && dv) {
          if (/^[+-]?\d+(?:[dhm])?$/i.test(dv)) {
            const dt = applyRelativeDateTimeOffset(new Date(), dv);
            input.value = formatLocalYMDHM(dt);
          } else if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(dv)) {
            input.value = dv;
          } else {
            input.placeholder = param.description || "YYYY-MM-DDTHH:mm";
            input.classList.add("placeholder-style");
          }
        }

        paramWrapper.appendChild(input);
      } else if (param.type === "integer") {
        // INTEGER (aligned like text)
        paramWrapper.classList.add("body-int-wrapper");

        const input = document.createElement("input");
        input.type = "number";
        input.step = "1";
        input.inputMode = "numeric";
        input.pattern = "\\d*";
        input.id = `body-param-${param.name}`;
        input.className = "body-param-input body-int-input";
        input.dataset.path = param.path;
        input.dataset.type = param.type;

        if (Number.isInteger(param.defaultValue)) {
          input.value = String(param.defaultValue);
        } else if (
          typeof param.defaultValue === "string" &&
          /^\d+$/.test(param.defaultValue)
        ) {
          input.value = param.defaultValue;
        } else {
          input.placeholder = param.description || "Enter integer";
        }

        paramWrapper.appendChild(input);
      } else {
        // PLAIN TEXT (aligned right like query)
        paramWrapper.classList.add("body-text-wrapper");

        const input = document.createElement("input");
        input.type = "text";
        input.id = `body-param-${param.name}`;
        input.className = "body-param-input body-text-input";
        input.dataset.path = param.path;
        input.dataset.type = param.type;

        if (param.defaultValue !== undefined && param.defaultValue !== "") {
          input.value = param.defaultValue;
        } else {
          input.placeholder = param.description || "Enter value";
        }

        paramWrapper.appendChild(input);
      }

      bodyParamContainer.appendChild(paramWrapper);
    });
  } catch (error) {
    console.error("Error populating Body Parameters:", error);
  }
}

// Parameter select label value pairs
function normalizeOptions(options) {
  return (options || []).map((opt) => {
    if (typeof opt === "string") return { label: opt, value: opt };
    const label =
      opt && typeof opt.label === "string"
        ? opt.label
        : String(opt?.value ?? "");
    const value =
      opt && typeof opt.value !== "undefined" ? String(opt.value) : label;
    return { label, value };
  });
}

// Stylize ad-hoc APIs
function applyDynamicStyles() {
  // get dynamically generated elements
  const endpointInput = document.getElementById("adhoc-endpoint");
  const bodyTextarea = document.getElementById("adhoc-body");

  // add classes if necessary
  if (endpointInput) {
    endpointInput.classList.add("query-param-input");
  }

  if (bodyTextarea) {
    bodyTextarea.classList.add("json-textarea");
  }
}

// Map user inputs to request profile
function mapUserInputsToRequestProfile(profile, inputs) {
  if (!profile || !inputs) return;

  // helper: set value at dotted path, creating objects as needed
  const setAtPath = (obj, path, val) => {
    if (!path) return;
    const parts = String(path).split(".");
    let cur = obj;
    parts.forEach((p, i) => {
      if (i === parts.length - 1) {
        cur[p] = val;
      } else {
        cur[p] = cur[p] ?? {};
        cur = cur[p];
      }
    });
  };

  // collect multi-text values by path (we’ll set after we sweep)
  const multiTextBuckets = new Map();

  // first pass: handle everything except multi-select checkbox aggregation
  inputs.forEach((el) => {
    const path = el.dataset.path;
    const type = (el.dataset.type || "").toLowerCase();
    if (!path) return;

    // normalize basic value
    const raw = (el.value ?? "").toString().trim();

    if (type === "multi-text") {
      if (!multiTextBuckets.has(path)) multiTextBuckets.set(path, []);
      if (raw !== "") multiTextBuckets.get(path).push(raw);
      return;
    }

    if (type === "multi-select") {
      // handled after this loop (we need all checkboxes)
      return;
    }

    if (type === "boolean") {
      // accept "true"/"false" or select choice; skip if placeholder/empty
      if (raw === "true" || raw === "false") {
        setAtPath(profile, path, raw === "true");
        return;
      }
      if (el.tagName === "SELECT" && el.value !== "") {
        setAtPath(profile, path, el.value === "true");
      }
      return;
    }

    if (type === "integer") {
      if (raw === "") return; // skip empty
      if (!/^-?\d+$/.test(raw)) {
        throw new Error(`"${raw}" is not a valid integer for ${path}`);
      }
      setAtPath(profile, path, parseInt(raw, 10));
      return;
    }

    if (type === "date") {
      // skip empty or placeholder
      const isPlaceholder =
        /^mm\/dd\/yyyy$/i.test(raw) ||
        (typeof el.placeholder === "string" && raw === el.placeholder);
      if (raw === "" || isPlaceholder) return;
      // <input type="date"> gives YYYY-MM-DD; keep as-is
      setAtPath(profile, path, raw);
      return;
    }

    if (type === "datetime") {
      // <input type="datetime-local"> yields 'YYYY-MM-DDTHH:mm'
      const isPlaceholder =
        typeof el.placeholder === "string" && raw === el.placeholder;
      if (raw === "" || isPlaceholder) return;
      // safety check
      if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(raw)) {
        throw new Error(
          `"${raw}" is not a valid datetime (expected YYYY-MM-DDTHH:mm) for ${path}`
        );
      }
      setAtPath(profile, path, raw);
      return;
    }

    if (type === "select") {
      // regular select (e.g., symbolic period): skip if placeholder/empty
      if (raw === "") return;
      setAtPath(profile, path, raw);
      return;
    }

    // default: plain text
    if (raw === "") return;
    setAtPath(profile, path, raw);
  });

  // apply multi-text arrays (only if any non-empty values)
  for (const [path, arr] of multiTextBuckets.entries()) {
    if (arr.length) setAtPath(profile, path, arr);
  }

  // aggregate multi-select checkboxes by path (checked only)
  const byPath = {};
  inputs.forEach((el) => {
    if ((el.dataset.type || "").toLowerCase() !== "multi-select") return;
    const path = el.dataset.path;
    if (!path) return;
    byPath[path] ||= [];
    if (el.checked) byPath[path].push(el.value);
  });
  Object.keys(byPath).forEach((path) => {
    const vals = byPath[path];
    if (vals.length) setAtPath(profile, path, vals);
  });
}

// Clean built request profile of empty/null fields
function pruneRequestBody(node) {
  const isEmptyish = (v) => v === "" || v === null || typeof v === "undefined";

  if (Array.isArray(node)) {
    for (let i = node.length - 1; i >= 0; i--) {
      const v = node[i];
      if (v && typeof v === "object") {
        if (pruneRequestBody(v)) node.splice(i, 1);
      } else if (isEmptyish(v) || (typeof v === "string" && v.trim() === "")) {
        node.splice(i, 1);
      }
    }
    return node.length === 0;
  }

  if (node && typeof node === "object") {
    for (const k of Object.keys(node)) {
      const v = node[k];
      if (Array.isArray(v)) {
        if (pruneRequestBody(v)) delete node[k];
      } else if (v && typeof v === "object") {
        if (pruneRequestBody(v)) delete node[k];
      } else if (isEmptyish(v) || (typeof v === "string" && v.trim() === "")) {
        delete node[k];
      }
    }
    return Object.keys(node).length === 0;
  }

  return false;
}

// Check if token is expired or near expiry
function isTokenExpiredOrNear(expirationIso, bufferSeconds = 10) {
  if (!expirationIso) return true;
  const exp = new Date(expirationIso).getTime();
  if (!Number.isFinite(exp)) return true;
  const now = Date.now();
  return exp - now <= bufferSeconds * 1000;
}

// Ensure API context is ready: client URL, client ID, access token
async function ensureApiReadyContext() {
  // 1) client URL
  const clienturl = await getClientUrl();
  if (!clienturl) {
    alert("Client URL is required. Please set/refresh the Client URL first.");
    return { ok: false };
  }

  // load current data snapshot
  const data = await loadClientData();
  const clientData = data[clienturl] || {};

  // 2) client ID (needed for token acquisition)
  const clientId =
    clientData.clientid ||
    document.getElementById("client-id")?.value?.trim() ||
    "";

  if (!clientId) {
    alert(
      "Client ID is required to request an access token. Please enter and save a Client ID."
    );
    return { ok: false };
  }

  // 3) access token (refresh if missing/expired/near-expiry)
  const tokenMissing = !clientData.accesstoken;
  const tokenStale =
    clientData.accessTokenSource !== "manual" &&
    isTokenExpiredOrNear(clientData.expirationdatetime, 10);

  if (tokenMissing || tokenStale) {
    if (!clientData.clientid && clientId) {
      // persist minimal fields to avoid a later token fetch failure
      data[clienturl] = {
        ...(data[clienturl] || {}),
        clientid: clientId,
        tokenurl: `${clienturl}accessToken?clientId=${clientId}`,
        apiurl: `${clienturl}api`,
        editdatetime: new Date().toISOString(),
      };
      await saveClientData(data);
    }

    try {
      await fetchToken();

      // wait for token to appear; treat failure as a user-facing issue
      const updatedClientData = await waitForUpdatedToken(clienturl, 5, 1000);
      return { ok: true, clienturl, clientData: updatedClientData };
    } catch (e) {
      alert(
        "Unable to retrieve a valid access token. Please verify Client ID/Secret/Tenant settings and try again."
      );
      return { ok: false };
    }
  }
  // token already valid
  return { ok: true, clienturl, clientData };
}

// Wait for new access token if needed
async function waitForUpdatedToken(clienturl, maxRetries = 5, delayMs = 1000) {
  for (let i = 0; i < maxRetries; i++) {
    await new Promise((resolve) => setTimeout(resolve, delayMs));
    const updatedData = await loadClientData();
    const updatedClientData = updatedData[clienturl] || {};

    if (updatedClientData.accesstoken) {
      return updatedClientData;
    }
  }

  return null; // <-- no throw
}

// Clear API response for a new request
function clearApiResponse() {
  const responseSection = document.getElementById("response-section");

  if (responseSection) {
    const pre = document.createElement("pre");
    pre.textContent = "Awaiting API Response...";
    responseSection.replaceChildren(pre);
  }

  window.lastApiResponseObject = null;

  updateResponseDependentButtons(false);
}

// Handle API Library selection
async function handleApiSelection(selectedKey) {
  updateRequestDependentButtons(false);
  updateResponseDependentButtons(false);
  clearApiResponse();

  // Request details must always correspond to a newly sent request
  lastRequestDetails = null;

  if (!selectedKey) return;

  clearParameters();
  clearDevLinkBanner();

  await populatePathParameters(selectedKey);
  await populateQueryParameters(selectedKey);
  await populateBodyParameters(selectedKey);

  applyDynamicStyles();
}

// Reset parameters button
async function onResetParamsClick() {
  try {
    await resetCurrentApiParameters();
  } catch (e) {
    console.error("Reset Parameters failed:", e);
    alert("Unable to reset parameters. See console for details.");
  }
}

// Re-render currently selected API's parameter UI from defaults
async function resetCurrentApiParameters() {
  // 1) determine selected api
  const apiSel = document.getElementById("api-selector");
  const selectedApiKey = apiSel?.value || "";
  if (!selectedApiKey || selectedApiKey === "Select API...") {
    alert("Select an API first to reset its parameters.");
    return;
  }

  // 2) clear current param ui
  clearParameters();

  // 3) repopulate from library defaults
  await populatePathParameters(selectedApiKey);
  await populateQueryParameters(selectedApiKey);
  await populateBodyParameters(selectedApiKey);

  // 4) re-apply any dynamic styles
  if (typeof applyDynamicStyles === "function") {
    applyDynamicStyles();
  }

  // 5) if this is ad-hoc, ensure blank fields
  if (selectedApiKey === "adHocGet") {
    const ep = document.getElementById("adhoc-endpoint");
    if (ep) ep.value = "";
  }
  if (selectedApiKey === "adHocPost") {
    const ep = document.getElementById("adhoc-endpoint");
    const tb = document.getElementById("adhoc-body");
    if (ep) ep.value = "";
    if (tb) tb.value = "";
  }
}

// Execute API call with multi-call support (such as for paginated requests)
async function executeApiCall() {
  const button = document.getElementById("execute-api");
  const originalText = button?.textContent || "Execute";
  let animation;

  const stopAnimation = () => {
    if (animation?.interval) clearInterval(animation.interval);
  };

  try {
    clearApiResponse();

    // start loading animation
    animation = startLoadingAnimation(button);

    // Session check first
    if (!(await isValidSession())) {
      alert("Requires a valid ADP Workforce Manager session.");
      stopAnimation();
      setButtonFailText(
        button,
        "Invalid Session",
        2000,
        animation?.originalText || originalText
      );
      return;
    }

    // Ensure client URL exists
    const clienturl = await getClientUrl();
    if (!clienturl) {
      alert(
        "Client URL is required. Please refresh or set the Client URL first."
      );
      stopAnimation();
      setButtonFailText(
        button,
        "Missing Client URL",
        2000,
        animation?.originalText || originalText
      );
      return;
    }

    // Ensure API selection exists
    const apiDropdown = document.getElementById("api-selector");
    const selectedApiKey = apiDropdown?.value;
    if (!selectedApiKey || selectedApiKey === "Select API...") {
      alert("Please select an API to execute.");
      stopAnimation();
      setButtonFailText(
        button,
        "No API Selected",
        2000,
        animation?.originalText || originalText
      );
      return;
    }

    // Load current tenant data
    let data = await loadClientData();
    let clientData = data[clienturl] || {};

    // Ensure Client ID exists
    // Prefer saved, fallback to field
    const clientId =
      clientData.clientid ||
      document.getElementById("client-id")?.value?.trim() ||
      "";

    if (!clientId) {
      alert("Client ID is required. Please enter and save a Client ID.");
      stopAnimation();
      setButtonFailText(
        button,
        "Missing Client ID",
        2000,
        animation?.originalText || originalText
      );
      return;
    }

    // Token check
    const tokenMissing = !clientData.accesstoken;
    const tokenNearOrExpired =
      clientData.accessTokenSource !== "manual" &&
      isTokenExpiredOrNear(clientData.expirationdatetime, 10);

    if (tokenMissing || tokenNearOrExpired) {
      // Ensure minimal fields exist in storage so token retrieval doesn’t fail
      if (!clientData.clientid) {
        data[clienturl] = {
          ...(data[clienturl] || {}),
          clientid: clientId,
          tokenurl: `${clienturl}accessToken?clientId=${clientId}`,
          apiurl: `${clienturl}api`,
          editdatetime: new Date().toISOString(),
        };
        await saveClientData(data);

        // reload snapshot
        data = await loadClientData();
        clientData = data[clienturl] || {};
      }

      // Initiate token retrieval (returns false if missing prerequisites)
      const initiated = await fetchToken();
      if (!initiated) {
        // fetchToken already alerted the user; just exit cleanly
        stopAnimation();
        setButtonFailText(
          button,
          "Token Needed",
          2000,
          animation?.originalText || originalText
        );
        return;
      }

      const updatedClientData = await waitForUpdatedToken(clienturl, 5, 1000);
      if (!updatedClientData?.accesstoken) {
        alert(
          "Unable to retrieve a valid access token. Please verify your settings and try again."
        );
        stopAnimation();
        setButtonFailText(
          button,
          "Token Failed",
          2000,
          animation?.originalText || originalText
        );
        return;
      }

      clientData = updatedClientData;
    }

    const accessToken = clientData.accesstoken;

    // Real headers used for the actual fetch
    const requestHeaders = {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    };

    // Redacted headers used ONLY for lastRequestDetails display/storage
    const redactedHeaders = {
      ...requestHeaders,
      Authorization: "Bearer <AccessToken>",
    };

    // Shared request variables
    let fullUrl = "";
    let requestBody = null;
    let requestMethod = "GET";

    // Load selected API definition
    const apiLibrary = await loadApiLibrary();
    const selectedApi = apiLibrary[selectedApiKey];
    if (!selectedApi) {
      alert("Selected API not found in the library.");
      stopAnimation();
      setButtonFailText(
        button,
        "API Missing",
        2000,
        animation?.originalText || originalText
      );
      return;
    }

    requestMethod = (selectedApi.method || "GET").toUpperCase();

    // replace {param} tokens
    const pathUrl = buildUrlWithPathParams(selectedApi.url, selectedApi);
    fullUrl = clientData.apiurl + pathUrl;

    // handle query parameters for standard get requests
    if (requestMethod === "GET") {
      const queryParams = new URLSearchParams();
      const queryInputs = document.querySelectorAll(
        "#query-parameters-container .query-param-input"
      );
      queryInputs.forEach((input) => {
        const v = input.value.trim();
        if (v !== "" && v !== input.placeholder) {
          queryParams.append(input.id.replace("query-", ""), v);
        }
      });
      if (queryParams.toString()) fullUrl += "?" + queryParams.toString();
    }

    // handle ad-hoc requests from library keys (adHocGet/adHocPost)
    if (selectedApiKey === "adHocGet" || selectedApiKey === "adHocPost") {
      const endpointInput = document.getElementById("adhoc-endpoint");
      if (!endpointInput || !endpointInput.value.trim()) {
        alert("Please provide an endpoint URL.");
        stopAnimation();
        setButtonFailText(
          button,
          "Missing Endpoint",
          2000,
          animation?.originalText || originalText
        );
        return;
      }
      fullUrl = clientData.apiurl + endpointInput.value.trim();
    }

    // handle pre-request logic if needed
    if (selectedApi.preRequest) {
      const preRequestApi = apiLibrary[selectedApi.preRequest.apiKey];
      const preRequestUrl = clientData.apiurl + preRequestApi.url;

      const preResponse = await fetch(preRequestUrl, {
        method: "GET",
        headers: { Authorization: `Bearer ${accessToken}` },
      });

      if (!preResponse.ok) {
        const errorText = await preResponse.text();
        displayApiResponse({ error: errorText }, selectedApiKey);
        stopAnimation();
        setButtonFailText(
          button,
          "Failed!",
          2000,
          animation?.originalText || originalText
        );
        return;
      }

      const preResult = await preResponse.json();

      const {
        field,
        match,
        mapTo,
        ["data-path"]: dataPath,
      } = selectedApi.preRequest.responseFilter;

      let mappedValues = preResult
        .filter((item) => item[field] === match)
        .map((item) => item[mapTo]);

      const maxLimit =
        selectedApi.bodyParameters.find((p) => p.name === "qualifiers")
          ?.validation?.maxEntered || 1000;

      if (mappedValues.length > maxLimit) {
        alert(
          `Only the first ${maxLimit} entries will be used due to API limitations.`
        );
        mappedValues = mappedValues.slice(0, maxLimit);
      }

      requestBody = {};
      const pathParts = dataPath.split(".");
      let currentLevel = requestBody;
      pathParts.forEach((part, index) => {
        if (index === pathParts.length - 1) currentLevel[part] = mappedValues;
        else {
          currentLevel[part] = currentLevel[part] || {};
          currentLevel = currentLevel[part];
        }
      });
    } else {
      // handle request body for regular post apis
      if (requestMethod === "POST") {
        if (selectedApiKey === "adHocPost") {
          const bodyInput = document.getElementById("adhoc-body");
          if (!bodyInput || !bodyInput.value.trim()) {
            alert("Please provide a JSON body.");
            stopAnimation();
            setButtonFailText(
              button,
              "Missing Body",
              2000,
              animation?.originalText || originalText
            );
            return;
          }
          try {
            requestBody = JSON.parse(bodyInput.value.trim());
          } catch {
            alert("Invalid JSON body. Please correct it.");
            stopAnimation();
            setButtonFailText(
              button,
              "Bad JSON",
              2000,
              animation?.originalText || originalText
            );
            return;
          }
        } else if (selectedApi.requestProfile) {
          const profileTemplate = JSON.parse(
            JSON.stringify(selectedApi.requestProfile)
          );
          const bodyParamsContainer = document.getElementById(
            "body-parameters-container"
          );
          const paramInputs = Array.from(
            bodyParamsContainer.querySelectorAll("[data-path]")
          );
          mapUserInputsToRequestProfile(profileTemplate, paramInputs);
          pruneRequestBody(profileTemplate);
          requestBody = profileTemplate;
        }
      }
    }

    // Save request details for the request details button (redacted)
    lastRequestDetails = {
      method: requestMethod,
      url: fullUrl,
      headers: redactedHeaders,
      body: requestBody ? JSON.stringify(requestBody, null, 2) : null,
    };

    updateRequestDependentButtons(true);

    const response = await fetch(fullUrl, {
      method: requestMethod,
      headers: requestHeaders,
      body: lastRequestDetails.body,
    });

    const responseText = await response.text();
    let result;
    try {
      result = JSON.parse(responseText);
    } catch {
      result = { error: responseText };
    }

    displayApiResponse(result, selectedApiKey);

    stopAnimation();
    const baseText = animation?.originalText || originalText;
    if (response.ok) setButtonTempText(button, "Success!", 2000, baseText);
    else setButtonFailText(button, "Failed!", 2000, baseText);
  } catch (error) {
    console.error("Error executing API call:", error);
    alert(`API call failed: ${error.message}`);
    displayApiResponse({ error: error.message }, "Error");

    stopAnimation();
    setButtonFailText(
      button,
      "Failed!",
      2000,
      animation?.originalText || originalText
    );
  }
}

// Display API response (default = raw view)
async function displayApiResponse(response, apiKey) {
  const responseSection = document.getElementById("response-section");
  if (!responseSection) return;

  // Cache response for popout and raw/tree toggle
  window.lastApiResponseObject = response;

  // Preserve/create popout button
  let popoutButton = document.getElementById("popout-response");

  if (!popoutButton) {
    popoutButton = document.createElement("button");
    popoutButton.id = "popout-response";
    popoutButton.className = "btn3";

    const label = document.createTextNode("Popout Response ");

    const icon = document.createElement("img");
    icon.src = "icons/external-link.png";
    icon.alt = "";
    icon.className = "btn-icon";

    popoutButton.append(label, icon);
    popoutButton.addEventListener("click", popoutResponse);

    responseSection.prepend(popoutButton);
  }

  // Remove any CSV export button from the previous response.
  const existingExport = document.getElementById("export-api-csv");
  if (existingExport) {
    existingExport.remove();
  }

  // Load selected public API definition for optional CSV export.
  const apiLibrary = await loadApiLibrary();
  const selectedApi = apiLibrary[apiKey];

  const isNonLibraryKey =
    typeof apiKey === "string" &&
    (apiKey === "Error" ||
      apiKey === "adHocGet" ||
      apiKey === "adHocPost");

  if (!selectedApi && !isNonLibraryKey) {
    console.warn("API Key Not Found in Library:", apiKey);
  }

  // Add CSV export only when the selected library API supports it.
  if (selectedApi?.exportMap) {
    const exportCsvButton = document.createElement("button");
    exportCsvButton.id = "export-api-csv";
    exportCsvButton.className = "btn3";

    const label = document.createTextNode("Export CSV ");

    const icon = document.createElement("img");
    icon.src = "icons/export-csv.png";
    icon.alt = "";
    icon.className = "btn-icon";

    exportCsvButton.append(label, icon);

    exportCsvButton.addEventListener("click", () => {
      exportApiResponseToCSV(response, selectedApi.exportMap, apiKey);
    });

    responseSection.appendChild(exportCsvButton);
  }

  // Ensure raw/tree toggle exists.
  ensureViewToggle();

  // Clear the previous rendered response only.
  responseSection
    .querySelectorAll(".json-tree, pre")
    .forEach((node) => node.remove());

  // Raw view is the default.
  const pre = document.createElement("pre");
  pre.textContent = JSON.stringify(response, null, 2);
  responseSection.appendChild(pre);

  // Reset toggle state to raw.
  const toggle = document.getElementById("toggle-view");
  if (toggle) {
    toggle.dataset.mode = "raw";
    toggle.textContent = "Tree View";
  }

  // Enable response-dependent controls.
  updateResponseDependentButtons(true);

  const downloadButton = document.getElementById("download-response");
  if (downloadButton) {
    downloadButton.disabled = false;
    downloadButton.onclick = () => downloadApiResponse(response, apiKey);
  }
}

// JSON tree view renderer
function renderJsonTree(data, rootEl, { collapsedDepth = 1 } = {}) {
  rootEl.replaceChildren();

  const doc = rootEl.ownerDocument || document;
  const el = buildNode(data, undefined, 0, collapsedDepth, doc);

  rootEl.appendChild(el);
}

// JSON tree node count helper
function containerBadge(value) {
  if (Array.isArray(value)) return `[${value.length}]`;
  return `{${Object.keys(value).length}}`;
}

// JSON tree view node builder (with counts)
function buildNode(value, key, depth, collapsedDepth, doc = document) {
  const isObjLike = (v) => v && typeof v === "object" && v !== null;

  if (isObjLike(value)) {
    const details = doc.createElement("details");
    details.open = depth < collapsedDepth;

    const summary = doc.createElement("summary");
    const badge = containerBadge(value);

    // Notepad++-style labels: "key {2}"  /  "values [0]"  /  "{3}" or "[1]" for root nodes
    summary.textContent = key != null ? `${key} ${badge}` : badge;
    details.appendChild(summary);

    if (Array.isArray(value)) {
      for (let i = 0; i < value.length; i++) {
        details.appendChild(buildNode(value[i], i, depth + 1, collapsedDepth, doc));
      }
    } else {
      const keys = Object.keys(value);
      for (const k of keys) {
        details.appendChild(buildNode(value[k], k, depth + 1, collapsedDepth, doc));
      }
    }
    return details;
  } else {
    // leaf
    const row = doc.createElement("div");
    row.className = "json-leaf";
    row.textContent =
      key != null ? `${key}: ${formatScalar(value)}` : formatScalar(value);
    return row;
  }
}

// Pretty-print leaf values for the JSON tree
function formatScalar(v) {
  if (typeof v === "string") return `"${v}"`;
  if (v === null) return "null";
  return String(v);
}

// API response raw / tree view toggle
function ensureViewToggle() {
  const section = document.getElementById("response-section");
  let btn = document.getElementById("toggle-view");
  if (btn) return;

  btn = document.createElement("button");
  btn.id = "toggle-view";
  btn.className = "btn3";
  btn.dataset.mode = "raw"; // default mode
  btn.textContent = "Tree View";
  section.prepend(btn);

  btn.onclick = () => {
    const mode = btn.dataset.mode;
    // clear current render
    [...section.querySelectorAll(".json-tree, pre")].forEach((n) => n.remove());

    if (mode === "raw") {
      // switch to tree
      const tree = document.createElement("div");
      tree.className = "json-tree";
      section.appendChild(tree);
      renderJsonTree(window.lastApiResponseObject ?? {}, tree, {
        collapsedDepth: 1,
      });
      btn.dataset.mode = "tree";
      btn.textContent = "Raw View";
    } else {
      // switch back to raw
      const pre = document.createElement("pre");
      pre.textContent = JSON.stringify(
        window.lastApiResponseObject ?? {},
        null,
        2
      );
      section.appendChild(pre);
      btn.dataset.mode = "raw";
      btn.textContent = "Tree View";
    }
  };
}

// Enable/disable buttons that depend on having a last sent request
function updateRequestDependentButtons(enabled) {
  const btn = document.getElementById("view-request-details");
  if (btn) btn.disabled = !enabled;
}

// Enable/disable buttons that depend on having a response in the UI
function updateResponseDependentButtons(hasResponse) {
  const downloadButton = document.getElementById("download-response");
  const copyButton = document.getElementById("copy-api-response");

  if (downloadButton) downloadButton.disabled = !hasResponse;
  if (copyButton) copyButton.disabled = !hasResponse;
}

// Download API response button
async function downloadApiResponse(response, apiName) {
  const sanitizedApiName = String(apiName || "")
    .replace(/[^a-zA-Z0-9_-]/g, "_")
    .slice(0, 120);
  const defaultFileName = `${sanitizedApiName || "api_response"}.json`;

  const content = JSON.stringify(response, null, 2);

  // Size warning (bytes)
  const bytes = new TextEncoder().encode(content).length;
  const mb = bytes / (1024 * 1024);

  const WARN_MB = 5;
  const STRONG_WARN_MB = 20;

  if (mb >= STRONG_WARN_MB) {
    const ok = confirm(
      `This response is about ${mb.toFixed(
        1
      )} MB. Saving large files may be slow and could impact browser performance.\n\n` +
        `Tip: For repeated analysis or sharing, consider exporting the request and running it in Bruno/Postman.\n\n` +
        `Continue saving?`
    );
    if (!ok) return;
  } else if (mb >= WARN_MB) {
    const ok = confirm(
      `This response is about ${mb.toFixed(
        1
      )} MB. Saving may take a moment.\n\nContinue saving?`
    );
    if (!ok) return;
  }

  if (window.showSaveFilePicker) {
    try {
      const fileHandle = await window.showSaveFilePicker({
        suggestedName: defaultFileName,
        types: [
          {
            description: "JSON Files",
            accept: { "application/json": [".json"] },
          },
        ],
      });

      const writableStream = await fileHandle.createWritable();
      await writableStream.write(content);
      await writableStream.close();
      return;
    } catch (error) {
      if (error && error.name === "AbortError") {
        return; // user cancelled – silent exit
      }
      // Non-fatal: fall back to download helper
    }
  }

  try {
    downloadFile(defaultFileName, content, "application/json");
  } catch (error) {
    console.error("Failed to download API response:", error);
    alert("Failed to save the file.");
  }
}

// Copy response button
function copyApiResponse() {
  const button = document.getElementById("copy-api-response");
  const responseSection = document.getElementById("response-section");

  // find the first <pre> or <code> block that contains the json response
  const jsonElement = responseSection?.querySelector("pre, code");

  if (jsonElement) {
    const responseContent = jsonElement.innerText.trim();

    if (responseContent) {
      navigator.clipboard
        .writeText(responseContent)
        .then(() => {
          setButtonTempText(button, "Copied!");
        })
        .catch((err) => {
          console.error("Failed to copy API response:", err);
          setButtonFailText(button, "Copy Failed!");
        });
    } else {
      setButtonFailText(button, "No JSON!");
    }
  } else {
    setButtonFailText(button, "No Response!");
  }
}

// Request details button
function showRequestDetails() {
  const colors = getPopupThemeColors();

  // No request has been sent yet
  if (!lastRequestDetails) {
    const popupCss = `
      body {
        font-family: Arial, sans-serif;
        padding: 20px;
        margin: 0;
        background-color: ${colors.primary};
        color: ${colors.accent};
        line-height: 1.5;
      }

      h1 {
        color: ${colors.accent};
        margin: 0 0 0.5rem 0;
        font-size: 1.3rem;
        font-weight: bold;
      }

      p {
        margin: 0;
        font-size: 0.8rem;
      }
    `;

    const popup = createPopupWindow(
      "No Request Details",
      "width=400,height=300,scrollbars=yes,resizable=yes",
      popupCss
    );

    if (!popup) return;

    const doc = popup.document;

    const heading = doc.createElement("h1");
    heading.textContent = "No request details available.";

    const message = doc.createElement("p");
    message.textContent = "Send a request first to view request details.";

    doc.body.append(heading, message);
    return;
  }

  // Request details popup
  const popupCss = `
    body {
      font-family: Arial, sans-serif;
      padding: 16px;
      margin: 0;
      line-height: 1.6;
      background-color: ${colors.primary};
      color: ${colors.accent};
    }

    h1 {
      color: ${colors.accent};
      font-size: 1.4rem;
      font-weight: bold;
      margin: 0 0 0.75rem 0;
    }

    h2 {
      color: ${colors.accent};
      font-size: 0.9rem;
      margin: 1rem 0 0.35rem 0;
    }

    .meta-line {
      margin: 0.15rem 0;
      font-size: 0.8rem;
    }

    .meta-label {
      font-weight: bold;
    }

    pre {
      background: ${colors.buttonText};
      border: 2px solid ${colors.accent};
      border-radius: 6px;
      padding: 10px;
      overflow-x: auto;
      white-space: pre;
      font-family: ui-monospace, SFMono-Regular, Menlo, Monaco,
        Consolas, "Liberation Mono", "Courier New", monospace;
      font-size: 0.7rem;
      color: ${colors.accent};
    }
  `;

  const popup = createPopupWindow(
    "Request Details",
    "width=1000,height=600,scrollbars=yes,resizable=yes",
    popupCss
  );

  if (!popup) return;

  const doc = popup.document;

  const heading = doc.createElement("h1");
  heading.textContent = "Request Details";

  // Method
  const methodLine = doc.createElement("p");
  methodLine.className = "meta-line";

  const methodLabel = doc.createElement("span");
  methodLabel.className = "meta-label";
  methodLabel.textContent = "Method: ";

  const methodValue = doc.createElement("span");
  methodValue.textContent = String(lastRequestDetails.method || "");

  methodLine.append(methodLabel, methodValue);

  // URL
  const urlLine = doc.createElement("p");
  urlLine.className = "meta-line";

  const urlLabel = doc.createElement("span");
  urlLabel.className = "meta-label";
  urlLabel.textContent = "URL: ";

  const urlValue = doc.createElement("span");
  urlValue.textContent = String(lastRequestDetails.url || "");

  urlLine.append(urlLabel, urlValue);

  // Headers
  const headersHeading = doc.createElement("h2");
  headersHeading.textContent = "Headers";

  const headersPre = doc.createElement("pre");
  headersPre.textContent = JSON.stringify(
    lastRequestDetails.headers || {},
    null,
    2
  );

  // Body
  const bodyHeading = doc.createElement("h2");
  bodyHeading.textContent = "Body";

  const bodyPre = doc.createElement("pre");
  bodyPre.textContent =
    lastRequestDetails.body != null && lastRequestDetails.body !== ""
      ? String(lastRequestDetails.body)
      : "No body";

  doc.body.append(
    heading,
    methodLine,
    urlLine,
    headersHeading,
    headersPre,
    bodyHeading,
    bodyPre
  );
}

// Popout API response
function popoutResponse() {
  const data = window.lastApiResponseObject;
  const colors = getPopupThemeColors();

  // No response available
  if (!data) {
    const popupCss = `
      body {
        font-family: Arial, sans-serif;
        padding: 20px;
        margin: 0;
        background-color: ${colors.primary};
        color: ${colors.accent};
        line-height: 1.5;
      }

      h1 {
        color: ${colors.accent};
        margin: 0 0 0.5rem 0;
        font-size: 1.3rem;
        font-weight: bold;
      }

      p {
        margin: 0;
        font-size: 0.95rem;
      }
    `;

    const popup = createPopupWindow(
      "No Response",
      "width=400,height=300,scrollbars=yes,resizable=yes",
      popupCss
    );

    if (!popup) return;

    const doc = popup.document;

    const heading = doc.createElement("h1");
    heading.textContent = "No response available.";

    const message = doc.createElement("p");
    message.textContent =
      "Please send an API request to generate a response.";

    doc.body.append(heading, message);
    return;
  }

  const mode =
    document.getElementById("toggle-view")?.dataset.mode || "raw";

  // Raw response popup
  if (mode === "raw") {
    const popupCss = `
      body {
        font-family: Arial, sans-serif;
        padding: 16px;
        margin: 0;
        line-height: 1.6;
        background-color: ${colors.primary};
        color: ${colors.accent};
      }

      h1 {
        color: ${colors.accent};
        font-size: 1.4rem;
        font-weight: bold;
        margin: 0 0 0.75rem 0;
      }

      pre {
        background: ${colors.buttonText};
        border: 2px solid ${colors.accent};
        border-radius: 6px;
        padding: 10px;
        overflow-x: auto;
        white-space: pre;
        font-family: ui-monospace, SFMono-Regular, Menlo, Monaco,
          Consolas, "Liberation Mono", "Courier New", monospace;
        font-size: 0.7rem;
        color: ${colors.accent};
      }
    `;

    const popup = createPopupWindow(
      "API Response",
      "width=800,height=900,scrollbars=yes,resizable=yes",
      popupCss
    );

    if (!popup) return;

    const doc = popup.document;

    const heading = doc.createElement("h1");
    heading.textContent = "API Response";

    const pre = doc.createElement("pre");
    pre.textContent = JSON.stringify(data, null, 2);

    doc.body.append(heading, pre);
    return;
  }

  // Tree response popup
  const popupCss = `
    body {
      font-family: Arial, sans-serif;
      padding: 16px;
      margin: 0;
      line-height: 1.6;
      background-color: ${colors.primary};
      color: ${colors.accent};
    }

    h1 {
      color: ${colors.accent};
      margin: 0 0 0.5rem 0;
      font-size: 1.4rem;
      font-weight: bold;
    }

    .json-tree {
      font: 12px/1.4 ui-monospace, SFMono-Regular, Menlo,
        Consolas, monospace;
      background: ${colors.buttonText};
      border: 2px solid ${colors.accent};
      border-radius: 6px;
      padding: 10px;
      color: ${colors.accent};
    }

    .json-tree details {
      margin-left: 0.75rem;
    }

    .json-tree summary {
      cursor: pointer;
      outline: none;
    }

    .json-tree .json-leaf {
      margin-left: 1.5rem;
      white-space: pre-wrap;
    }
  `;

  const popup = createPopupWindow(
    "API Response (Tree)",
    "width=900,height=1000,scrollbars=yes,resizable=yes",
    popupCss
  );

  if (!popup) return;

  const doc = popup.document;

  const heading = doc.createElement("h1");
  heading.textContent = "API Response (Tree)";

  const tree = doc.createElement("div");
  tree.className = "json-tree";

  renderJsonTree(data, tree, {
    collapsedDepth: 1,
  });

  doc.body.append(heading, tree);
}

// Export JSON response to CSV button
async function exportApiResponseToCSV(response, apiKey) {
  if (!response || (Array.isArray(response) && response.length === 0)) {
    alert("No data available to export.");
    return;
  }

  // extract array if the response is an object with a nested array
  let extractedArray = response;
  if (!Array.isArray(response)) {
    for (const key in response) {
      if (Array.isArray(response[key])) {
        extractedArray = response[key];
        break;
      }
    }
  }

  if (!Array.isArray(extractedArray) || extractedArray.length === 0) {
    alert("No valid array data found for export.");
    return;
  }

  // load api library
  const apiLibrary = await loadApiLibrary();

  // set the file name based on the api key
  const safeApiName = apiKey ? apiKey : "api-response"; // ensure safe fallback

  let expandedHeaders = new Set();
  let expandedData = [];

  // **recursive function to flatten objects**
  function flattenObject(obj, parentKey = "") {
    let flatRow = {};
    let arrayFields = {};

    Object.entries(obj).forEach(([key, value]) => {
      const newKey = parentKey ? `${parentKey}.${key}` : key;

      if (Array.isArray(value)) {
        arrayFields[newKey] = value;
      } else if (typeof value === "object" && value !== null) {
        // **recursively flatten nested objects**
        const nestedFlat = flattenObject(value, newKey);
        Object.assign(flatRow, nestedFlat.flatRow);
        Object.assign(arrayFields, nestedFlat.arrayFields);
      } else {
        flatRow[newKey] = value;
        expandedHeaders.add(newKey);
      }
    });

    return { flatRow, arrayFields };
  }

  // flatten each object in the array
  extractedArray.forEach((item) => {
    const { flatRow, arrayFields } = flattenObject(item);
    const maxRows = Math.max(
      ...Object.values(arrayFields).map((arr) => arr.length),
      1
    );

    for (let i = 0; i < maxRows; i++) {
      let rowCopy = { ...flatRow };

      Object.entries(arrayFields).forEach(([field, values]) => {
        if (typeof values[i] === "object" && values[i] !== null) {
          Object.entries(values[i]).forEach(([subKey, subValue]) => {
            let subField = `${field}.${subKey}`;
            rowCopy[subField] = subValue;
            expandedHeaders.add(subField);
          });
        } else {
          rowCopy[field] = values[i] !== undefined ? values[i] : "";
          expandedHeaders.add(field);
        }
      });

      expandedData.push(rowCopy);
    }
  });

  // remove empty columns
  expandedHeaders = Array.from(expandedHeaders);
  const columnsWithData = expandedHeaders.filter((header) =>
    expandedData.some((row) => row[header] !== "" && row[header] !== undefined)
  );

  //const csvRows = [columnsWithData.join(",")];
  const csvRows = [`"${columnsWithData.join('","')}"`];

  expandedData.forEach((row) => {
    const rowData = columnsWithData.map(
      (header) => `"${row[header] !== undefined ? row[header] : ""}"`
    );
    csvRows.push(rowData.join(","));
  });

  const csvContent = csvRows.join("\n");
  const blob = new Blob([csvContent], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  // ensure file always saves with the same name (overwrite existing file)
  link.download = `${safeApiName}-export.csv`;
  link.href = url;

  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}
// ================================= //


// ===== SESSION FUNCTIONS ===== //
// Remove -nosso from client URL
function createSsoUrl(clientUrl) {
  return clientUrl.replace("-nosso.", ".");
}

// Construct API URL
function toApiUrl(url) {
  if (!url) return "";
  try {
    const u = new URL(url);
    // strip query/hash; normalize path
    let path = u.pathname || "/";
    if (!path.endsWith("/")) path += "/";
    // ensure exactly “…/api” (no trailing slash)
    if (path.endsWith("/api/") || path.endsWith("/api")) {
      path = "/api";
    } else {
      path = path + "api";
    }
    return u.origin + path;
  } catch {
    // fallback if url constructor fails
    let s = (url.split(/[?#]/)[0] || "").replace(/\/+$/, "");
    return s + "/api";
  }
}

// Open URL in normal mode (and generally everywhere) using the Tabs API
function openURLNormally(url) {
  try {
    chrome.tabs.create({ url, active: true });
  } catch (e) {
    // fallback: anchor click if tabs API is unavailable for some reason
    const a = document.createElement("a");
    a.href = url;
    a.target = "_blank";
    a.rel = "noopener noreferrer";
    (document.body || document.documentElement).appendChild(a);
    a.click();
    a.remove();
  }
}

// Get the base URL (Vanity URL) from active tab and inject
function getVanityUrl(tabUrl) {
  let url = new URL(tabUrl);
  let hostname = url.hostname;

  // handle the sso url adjustments
  if (hostname.includes(".mykronos.com") && !hostname.includes("-nosso")) {
    if (hostname.includes(".prd.mykronos.com")) {
      hostname = hostname.replace(
        ".prd.mykronos.com",
        "-nosso.prd.mykronos.com"
      );
    } else if (hostname.includes(".npr.mykronos.com")) {
      hostname = hostname.replace(
        ".npr.mykronos.com",
        "-nosso.npr.mykronos.com"
      );
    }
  }

  return `${url.protocol}//${hostname}/`;
}

// Validate Session Based On Active Tab URL
async function isValidSession() {
  const clientUrl = await getClientUrl();
  return clientUrl !== null; // if getClientUrl() resolves null, the session is invalid
}

// Validate current webpage is a valid ADP WFMgr session
function validateWebPage(url) {
  // first check if we're even on mykronos.com
  if (!url.includes("mykronos.com")) {
    return { valid: false, message: "Invalid Domain" };
  }

  // define invalid URL patterns
  const invalidPatterns = [
    {
      pattern: "mykronos.com/authn/",
      message: "Invalid Login - Authentication Required",
    },
    {
      pattern: "mykronos.com/wfd/unauthorized",
      message: "Invalid Login - Unauthorized Access",
    },
    {
      pattern: /:\/\/adp-developer\.mykronos\.com\//i,
      message: "Developer Portal not supported for API session",
    },
  ];

  // check against invalid patterns
  for (const { pattern, message } of invalidPatterns) {
    if (
      typeof pattern === "string" ? url.includes(pattern) : pattern.test(url)
    ) {
      return { valid: false, message };
    }
  }

  // if no invalid patterns matched, the URL is valid
  return { valid: true, message: "Valid" };
}

// Retrieve current client URL, preferring the linked WFM tab
async function getClientUrl() {
  // 1) Prefer the linked tab's origin if available
  try {
    if (window.HermesLink && typeof HermesLink.getBaseUrl === "function") {
      const linkedBase = await HermesLink.getBaseUrl();

      if (linkedBase) {
        const validation = validateWebPage(linkedBase);
        if (validation?.valid) {
          const vanityUrl = getVanityUrl(linkedBase);
          return vanityUrl || null;
        }
        // fall through to active-tab mode
      }
    }
  } catch (e) {
    // Non-fatal: fall back silently
  }

  // 2) Fallback: use the active tab
  return new Promise((resolve) => {
    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
      if (chrome.runtime.lastError) {
        console.warn("Failed to query active tab:", chrome.runtime.lastError);
        resolve(null);
        return;
      }

      const tabUrl = tabs?.[0]?.url;
      if (!tabUrl) {
        resolve(null);
        return;
      }

      const validation = validateWebPage(tabUrl);
      if (!validation?.valid) {
        resolve(null);
        return;
      }

      const vanityUrl = getVanityUrl(tabUrl);
      resolve(vanityUrl || null);
    });
  });
}

// Overlay Return To Linked Tab Button
async function handleReturnToLinkedTab(button) {
  if (!button) return;

  try {
    button.disabled = true;
    await HermesLink.goToLinkedTab();

    // After switching to the linked tab, reload the panel so all
    // fields reinitialize for the active tenant/session.
    window.location.reload();
  } catch (error) {
    console.warn("Failed to return to linked tab:", error);
    alert("Unable to return to linked tab: " + error.message);
  } finally {
    button.disabled = false;
  }
}

// Overlay link this tab instead button
async function handleRelinkToCurrentTab(button) {
  if (!button) return;

  try {
    button.disabled = true;

    const [currentTab] = await chrome.tabs.query({
      active: true,
      currentWindow: true,
    });

    if (!currentTab) {
      throw new Error("No active tab found");
    }

    const validation = validateWebPage(currentTab.url);
    if (!validation.valid) {
      throw new Error(validation.message);
    }

    await HermesLink.relinkToCurrentTab(currentTab);

    // after relinking to this tab, reload the panel so it
    // re-reads storage & URL and shows the correct tenant data.
    window.location.reload();
  } catch (error) {
    alert("Unable to link this tab: " + error.message);
  } finally {
    button.disabled = false;
  }
}

// Hermes link check connection button
async function checkHermesConnectionClick() {
  const btn = document.getElementById("hermes-check-connection");
  const original = btn?.textContent || "Check Connection";

  try {
    if (btn) btn.textContent = "Checking...";
    await HermesLink.checkState();
    if (btn) {
      btn.textContent = "Checked";
      setTimeout(() => (btn.textContent = original), 1200);
    }
  } catch (e) {
    if (btn) {
      btn.textContent = "Failed";
      setTimeout(() => (btn.textContent = original), 1500);
    }
    console.error("HermesLink check failed:", e);
  }
}

window.HermesLink = (function () {
  const ENFORCE_ACTIVE_TAB_OVERLAY = true;
  const PING_INTERVAL = 60 * 1000; // 1 minute polling

  const SESSION_KEYS = {
    TAB_ID: "hermesLinkedTabId",
    WINDOW_ID: "hermesLinkedWindowId",
    URL: "hermesLinkedUrl",
    ORIGIN: "hermesLinkedOrigin",
    TITLE: "hermesLinkedTitle",
    STATUS: "hermesLinkedStatus",
    LAST_VALIDATION: "hermesLastValidation",
    VALIDATION_MESSAGE: "hermesValidationMessage",
  };

  const STATUS_MESSAGES = {
    OK: {
      banner: "Active Tab: ",
      hint: "Session active in this tab",
      overlay: null,
    },

    STALE: {
      banner: "Session Needs Attention: ",
      hint: "Your session may have expired. Please refresh the page.",
      overlay:
        "Session may have expired. Return to WFM to refresh your session.",
    },

    INVALID: {
      banner: "Invalid Session: ",
      hint: "Please return to a valid WFM page.",
      overlay: "Invalid WFM session. Return to a valid WFM page to continue.",
    },

    WRONG_TAB: {
      banner: "Not Active Tab: ",
      hint: "Check Connection or link this tab.",
      overlay:
        "Session not validated in this tab. Click Check Connection or Link This Tab Instead.",
    },

    // TMS is allowed for manual AccessPanel use.
    // AccessPanel does not inspect or validate TMS page content.
    TMS_OK: {
      banner: "Active (TMS): ",
      hint: "AccessPanel available for manual use in Tenant Management System.",
      overlay: null,
    },

    // Developer portal is allowed for manual use.
    DEV_OK: {
      banner: "Developer Portal: ",
      hint: "AccessPanel UI enabled for manual copy/paste from the developer portal.",
      overlay: null,
    },
  };

  // Determine whether the active tab is an allowed TMS page.
  const isTmsUrl = (url) => {
    try {
      const u = new URL(url);
      const origin = u.origin.toLowerCase();

      return (
        origin.startsWith("https://adpvantage.adp.com") ||
        origin.startsWith("https://testadpvantage.adp.com")
      );
    } catch {
      return false;
    }
  };

  // Determine whether the active tab is an allowed developer portal page.
  const isDevPortalUrl = (url) => {
    try {
      const u = new URL(url);
      const origin = u.origin.toLowerCase();

      return (
        origin.startsWith("https://adp-developer.mykronos.com") ||
        origin.startsWith("https://sso-hlp02.gss-kcfn.mykronos.com")
      );
    } catch {
      return false;
    }
  };

  // Determine whether the active tab is a Boomi platform page.
  const isBoomiUrl = (url) => {
    try {
      const u = new URL(url);
      const host = u.hostname.toLowerCase();

      return (
        host === "platform.boomi.com" ||
        host.endsWith(".platform.boomi.com")
      );
    } catch {
      return false;
    }
  };

  // Normalize a WFM vanity URL for same-tenant comparison.
  // Compare protocol + hostname only and ignore path/trailing slash.
  const normalizeVanity = (raw) => {
    if (!raw) return null;

    let value = String(raw).trim();
    if (!value) return null;

    if (!/^https?:\/\//i.test(value)) {
      value = "https://" + value;
    }

    try {
      const normalizedUrl = getVanityUrl(value);
      const parsed = new URL(normalizedUrl);

      return `${parsed.protocol}//${parsed.hostname}`.toLowerCase();
    } catch {
      return null;
    }
  };

  const state = {
    isInitialized: false,
    checkingState: false,
  };

  // Get active tab.
  const getActiveTab = async () => {
    try {
      const [tab] = await chrome.tabs.query({
        active: true,
        currentWindow: true,
      });

      return tab || null;
    } catch (error) {
      console.error("Failed to get active tab:", error);
      return null;
    }
  };

  // Get linked-session state.
  const getLinkedState = async () => {
    try {
      return await chrome.storage.session.get(Object.values(SESSION_KEYS));
    } catch (error) {
      console.error("Failed to get linked state:", error);
      return {};
    }
  };

  // Update linked-session state.
  const updateLinkedState = async (newState) => {
    try {
      const timestamp = new Date().toISOString();

      await chrome.storage.session.set({
        ...newState,
        hermesLastValidation: timestamp,
      });
    } catch (error) {
      console.error("Failed to update linked state:", error);
    }
  };

  // Update AccessPanel UI based on the active browser context.
  const updateUI = async (validationResult = null) => {
    const {
      hermesLinkedTabId,
      hermesLinkedStatus,
      hermesValidationMessage,
      hermesLinkedOrigin,
    } = await getLinkedState();

    const currentTab = await getActiveTab();
    const isLinkedTab = currentTab?.id === hermesLinkedTabId;

    const isTms = currentTab ? isTmsUrl(currentTab.url) : false;
    const isDevPortal = currentTab
      ? isDevPortalUrl(currentTab.url)
      : false;
    const isBoomi = currentTab ? isBoomiUrl(currentTab.url) : false;

    // Validate the active tab as a normal WFM tab.
    const currentTabValidation = currentTab
      ? validateWebPage(currentTab.url)
      : { valid: false };

    // Determine whether the current WFM tab belongs to the same tenant
    // as the linked WFM session.
    let sameTenantAsLinked = false;
    let sessionVanity = null;

    if (hermesLinkedOrigin) {
      sessionVanity = normalizeVanity(hermesLinkedOrigin);
    }

    if (currentTab && currentTabValidation.valid && sessionVanity) {
      try {
        const currentVanity = normalizeVanity(currentTab.url);

        sameTenantAsLinked =
          !!currentVanity && currentVanity === sessionVanity;
      } catch {
        sameTenantAsLinked = false;
      }
    }

    // Determine current UI state.
    let currentStatus = "OK";

    if (isDevPortal && !isLinkedTab) {
      // Developer portal is allowed for manual use.
      currentStatus = "DEV_OK";
    } else if (isTms) {
      // TMS is allowed for manual AccessPanel use.
      // No TMS DOM data is read or inspected.
      currentStatus = "TMS_OK";
    } else {
      // Normal WFM-driven behavior.
      if (!isLinkedTab && sameTenantAsLinked) {
        // Different tab, but same tenant as linked session.
        currentStatus = "OK";
      } else if (!isLinkedTab) {
        currentStatus = "WRONG_TAB";
      } else if (hermesLinkedStatus === "stale") {
        currentStatus = "STALE";
      } else if (validationResult && !validationResult.valid) {
        currentStatus = "INVALID";
      }
    }

    const statusConfig = STATUS_MESSAGES[currentStatus];

    // ----- Overlay -----
    const overlay = document.getElementById("hermes-overlay");

    if (overlay) {
      const overlayMessage = document.querySelector(".overlay-content p");
      const relinkButton = document.getElementById("hermes-relink-tab");

      if (overlayMessage && statusConfig.overlay) {
        overlayMessage.textContent = statusConfig.overlay;
      }

      // Relink is only valid for normal WFM tabs.
      if (relinkButton) {
        relinkButton.style.display = currentTabValidation.valid
          ? "inline-block"
          : "none";
      }

      // OK, TMS_OK and DEV_OK never show the overlay.
      // If ENFORCE_ACTIVE_TAB_OVERLAY is false, WRONG_TAB is also soft.
      const isSoftWrongTab =
        currentStatus === "WRONG_TAB" && !ENFORCE_ACTIVE_TAB_OVERLAY;

      const overlayVisible =
        !["OK", "TMS_OK", "DEV_OK"].includes(currentStatus) &&
        !isBoomi &&
        !isSoftWrongTab;

      overlay.classList.toggle("visible", overlayVisible);
    }

    // ----- Banner -----
    const banner = document.getElementById("hermes-link-banner");

    if (banner) {
      const status = document.getElementById("hermes-link-status");
      const target = document.getElementById("hermes-link-target");
      const hint = document.getElementById("hermes-link-hint");

      if (status) {
        status.textContent = statusConfig.banner;
      }

      if (target) {
        target.textContent = currentTab?.title || "";
      }

      if (hint) {
        hint.textContent =
          hermesValidationMessage || statusConfig.hint;
      }
    }

    // Tab is considered active when:
    //   - we're on the linked or same WFM tenant (OK),
    //   - we're on TMS for manual AccessPanel use (TMS_OK),
    //   - we're on an allowed developer portal (DEV_OK).
    //
    // If ENFORCE_ACTIVE_TAB_OVERLAY is false, WRONG_TAB becomes a soft
    // state: banner/hints remain visible but the UI stays usable.
    const isSoftWrongTabForUi =
      currentStatus === "WRONG_TAB" && !ENFORCE_ACTIVE_TAB_OVERLAY;

    const shouldDisableUi =
      !["OK", "TMS_OK", "DEV_OK"].includes(currentStatus) &&
      !isBoomi &&
      !isSoftWrongTabForUi;

    document.body.classList.toggle("tab-inactive", shouldDisableUi);
  };

  // Validate linked WFM session.
  const validateSession = async () => {
    const { hermesLinkedUrl, hermesLinkedTabId } = await getLinkedState();

    if (!hermesLinkedUrl || !hermesLinkedTabId) {
      return {
        ok: false,
        code: "nolink",
        message: "No linked session found",
      };
    }

    try {
      const tab = await chrome.tabs
        .get(hermesLinkedTabId)
        .catch(() => null);

      if (!tab) {
        await updateLinkedState({
          hermesLinkedStatus: "stale",
          hermesValidationMessage: "Linked tab was closed",
        });

        return {
          ok: false,
          code: "closed",
          message: "Linked tab was closed",
        };
      }

      const validation = validateWebPage(tab.url);

      if (!validation.valid) {
        await updateLinkedState({
          hermesLinkedStatus: "stale",
          hermesValidationMessage: validation.message,
        });

        return {
          ok: false,
          code: "invalid",
          validation,
        };
      }

      // Keep HermesLink synchronized with the WFM URL actually loaded
      // in the linked tab. This handles switching tenants in the same tab.
      let newOrigin = null;

      try {
        newOrigin = new URL(tab.url).origin;
      } catch {
        const stored = await getLinkedState();
        newOrigin = stored.hermesLinkedOrigin || null;
      }

      await updateLinkedState({
        hermesLinkedUrl: tab.url,
        hermesLinkedOrigin: newOrigin,
        hermesLinkedTitle: tab.title || "",
        hermesLinkedStatus: "ok",
        hermesValidationMessage: "Session active",
      });

      return {
        ok: true,
        code: "ok",
        validation,
      };
    } catch {
      await updateLinkedState({
        hermesLinkedStatus: "stale",
        hermesValidationMessage: "Unable to verify session",
      });

      return {
        ok: false,
        code: "error",
        message: "Session check failed",
      };
    }
  };

  const core = {
    async validateAndUpdateState() {
      if (state.checkingState) return;

      state.checkingState = true;

      try {
        const validationResult = await validateSession();

        await updateUI(validationResult.validation);

        return validationResult;
      } catch (error) {
        console.error("State check failed:", error);
      } finally {
        state.checkingState = false;
      }
    },

    async initialize() {
      if (state.isInitialized) return;

      await this.validateAndUpdateState();

      setInterval(() => {
        this.validateAndUpdateState().catch((error) =>
          console.error("Periodic check failed:", error)
        );
      }, PING_INTERVAL);

      state.isInitialized = true;
    },
  };

  // Initialize linked-session management.
  core
    .initialize()
    .catch((error) =>
      console.error("Failed to initialize HermesLink:", error)
    );

  // Public API
  return {
    checkState: () => core.validateAndUpdateState(),

    relinkToCurrentTab: async (tab) => {
      if (!tab?.url) {
        throw new Error("No active tab");
      }

      const validation = validateWebPage(tab.url);

      if (!validation.valid) {
        throw new Error(validation.message);
      }

      await updateLinkedState({
        [SESSION_KEYS.TAB_ID]: tab.id,
        [SESSION_KEYS.WINDOW_ID]: tab.windowId,
        [SESSION_KEYS.URL]: tab.url,
        [SESSION_KEYS.ORIGIN]: new URL(tab.url).origin,
        [SESSION_KEYS.TITLE]: tab.title || "",
        [SESSION_KEYS.STATUS]: "ok",
        hermesValidationMessage: "Successfully linked to current tab",
      });

      await core.validateAndUpdateState();
    },

    getBaseUrl: async () => {
      const {
        hermesLinkedOrigin,
        hermesLinkedStatus,
      } = await getLinkedState();

      return hermesLinkedStatus === "ok"
        ? hermesLinkedOrigin
        : null;
    },
  };
})();

// ============================= //

// ===== EVENT LISTENERS ===== //
document.addEventListener("DOMContentLoaded", async () => {
  await purgeExpiredTokensInStorage();

  // Initial UI population
  await populateClientUrlField();
  await populateClientID();
  await populateAccessToken();
  await populateClientSecret();
  await populateTenantId();
  await populateRefreshToken();
  await restoreTokenTimers();
  await populateThemeDropdown();
  await restoreSelectedTheme();
  buildThemeMenuFromSelect();
  wireThemeMenuClicks();
  initMenus();
  await populateApiDropdown();

  // title bar / reload UI
  on("reload-app", "click", reloadApp);

  // collapsible sections
  on("toggle-tenant-section", "click", toggleTenantSection);
  on("toggle-access-section", "click", toggleAccessSection);
  on("toggle-api-library", "click", toggleApiLibrary);

  // restore persisted states on load
  restoreTenantSection();
  restoreAccessSection();
  restoreApiLibrary();

  // admin menu + settings overlay
  onAsync("clear-all-data", "click", clearAllData);
  onAsync("clear-client-data", "click", clearClientData);

  // links menu
  onAsync("links-developer-portal", "click", linksDeveloperPortal);
  onAsync("links-boomi", "click", linksBoomi);
  onAsync("links-install-integrations", "click", linksInstallIntegrations);

  // theme menu
  on("theme-selector", "change", themeSelection);

  // help menu
  onAsync("help-about", "click", helpAbout);
  onAsync("help-support", "click", helpSupport);

  // mask these fields by default
  ensureMasked("client-url");
  ensureMasked("client-id");
  ensureMasked("tenant-id");
  ensureMasked("access-token");
  ensureMasked("refresh-token");

  // tenant information section
  onAsync("generate-birt-file", "click", generateBirtPropertiesClick);
  on(
    "toggle-client-url",
    "click",
    () => toggleFieldVisibility("client-url", "toggle-client-url"),
    { onceKey: "reveal" }
  );
  onAsync("refresh-client-url", "click", refreshClientUrlClick);
  onAsync("copy-client-url", "click", copyClientUrlClick);
  on(
    "toggle-client-id",
    "click",
    () => toggleFieldVisibility("client-id", "toggle-client-id"),
    { onceKey: "reveal" }
  );
  onAsync("save-client-id", "click", saveClientIDClick);
  onAsync("copy-client-id", "click", copyClientIdClick);
  on("toggle-client-secret", "click", toggleClientSecretVisibility);
  onAsync("copy-client-secret", "click", copyClientSecretClick);
  onAsync("save-client-secret", "click", saveClientSecretClick);
  on(
    "toggle-tenant-id",
    "click",
    () => toggleFieldVisibility("tenant-id", "toggle-tenant-id"),
    { onceKey: "reveal" }
  );
  onAsync("save-tenant-id", "click", saveTenantIdClick);
  onAsync("copy-tenant-id", "click", copyTenantIdClick);

  // api tokens
  on(
    "toggle-access-token",
    "click",
    () => toggleFieldVisibility("access-token", "toggle-access-token"),
    { onceKey: "reveal" }
  );
  onAsync("get-token", "click", fetchToken);
  onAsync("apply-token", "click", applyManualAccessTokenClick);
  on("copy-token", "click", copyAccessToken);
  on(
    "toggle-refresh-token",
    "click",
    () => toggleFieldVisibility("refresh-token", "toggle-refresh-token"),
    { onceKey: "reveal" }
  );
  onAsync("refresh-access-token", "click", refreshAccessToken);
  on("copy-refresh-token", "click", copyRefreshToken);

  // api selector (guard against duplicate wiring)
  const apiSel = document.getElementById("api-selector");
  if (apiSel && !apiSel.dataset.changeListenerAttached) {
    apiSel.addEventListener("change", (e) => {
      void handleApiSelection(e.target.value);
    });
    apiSel.dataset.changeListenerAttached = "1";
  }

  // api buttons
  onAsync("execute-api", "click", executeApiCall);
  onAsync("reset-params", "click", onResetParamsClick);
  onAsync("copy-api-response", "click", copyApiResponse);
  on("view-request-details", "click", showRequestDetails);

  updateRequestDependentButtons(false);
  updateResponseDependentButtons(false);

  // popout response
  on("popout-response", "click", popoutResponse);

  onAsync("hermes-check-connection", "click", checkHermesConnectionClick);
  const relinkButton = document.getElementById("hermes-relink-tab");
  if (relinkButton) {
    relinkButton.addEventListener("click", () =>
      handleRelinkToCurrentTab(relinkButton)
    );
  }

  // visibility change handler
  document.addEventListener("visibilitychange", () => {
    if (document.hidden) {
      stopAllTokenTimers();
    } else {
      purgeExpiredTokensInStorage()
        .then(() => {
          HermesLink.checkState().catch((e) =>
            console.error("Visibility check failed:", e)
          );
          resumeTokenTimersFromStorage().catch((e) =>
            console.error("Timer resume failed:", e)
          );
        })
        .catch((e) => console.error("Token purge failed:", e));
    }
  });

  // focus handler
  window.addEventListener("focus", () => {
    purgeExpiredTokensInStorage()
      .then(() => {
        HermesLink.checkState().catch((e) =>
          console.error("Focus check failed:", e)
        );
        resumeTokenTimersFromStorage().catch((e) =>
          console.error("Timer resume failed:", e)
        );
      })
      .catch((e) => console.error("Token purge failed:", e));
  });
});