// main.js: Core functionality for AccessPanel extension UI

// ===== GLOBALS ===== //
const LOGGING_ENABLED = false;
const storageKey = "clientdata";
const PREFS_KEY = "hermes_preferences";
let accessTokenTimerInterval = null;
let refreshTokenTimerInterval = null;
let lastRequestDetails = null;
let HERMES_MYAPIS_MODE = false;

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
      handler(e).catch((err) => console.error(`${id} handler failed:`, err));
    },
    options
  );
}

// Dedicated event handler so listener wiring stays simplified
function adminSettingsClick(e) {
  e.preventDefault();
  openSettingsOverlay();
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

// Determine whether the current context is incognito
/*function isIncognitoContext() {
  return new Promise((resolve) => {
    try {
      chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
        if (chrome.runtime.lastError) {
          // fallback to window context
          chrome.windows.getCurrent({ populate: false }, (win) => {
            if (chrome.runtime.lastError || !win) return resolve(false);
            return resolve(!!win.incognito);
          });
          return;
        }

        const tab = tabs && tabs[0];
        if (tab && typeof tab.incognito === "boolean") {
          return resolve(tab.incognito);
        }

        // fallback if tab is missing or incognito flag unavailable
        chrome.windows.getCurrent({ populate: false }, (win) => {
          if (chrome.runtime.lastError || !win) return resolve(false);
          return resolve(!!win.incognito);
        });
      });
    } catch (e) {
      // defensive fallback
      chrome.windows.getCurrent({ populate: false }, (win) => {
        if (chrome.runtime.lastError || !win) return resolve(false);
        return resolve(!!win.incognito);
      });
    }
  });
}*/

// Determine whether the current extension context is incognito
function isIncognitoContext() {
  try {
    return Promise.resolve(!!chrome.extension?.inIncognitoContext);
  } catch {
    return Promise.resolve(false);
  }
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
// Load preferences from storage
function loadPreferences() {
  return new Promise((resolve) => {
    chrome.storage.local.get([PREFS_KEY], (res) => {
      const prefs = res[PREFS_KEY] || {};
      EXPORT_API_URL_VAR = prefs.apiUrlVar || "apiUrl";
      EXPORT_ACCESS_TOKEN_VAR = prefs.accessTokenVar || "accessToken";
      resolve();
    });
  });
}

// Save preferences to storage
function savePreferencesToStorage(apiUrlVar, accessTokenVar) {
  const prefs = {
    apiUrlVar: apiUrlVar || "apiUrl",
    accessTokenVar: accessTokenVar || "accessToken",
  };
  return new Promise((resolve) => {
    chrome.storage.local.set({ [PREFS_KEY]: prefs }, resolve);
  });
}

// Open settings overlay and populate form
function openSettingsOverlay() {
  const overlay = document.getElementById("settings-overlay");
  const apiUrlInput = document.getElementById("settings-api-url-var");
  const tokenInput = document.getElementById("settings-access-token-var");

  if (!overlay || !apiUrlInput || !tokenInput) return;

  apiUrlInput.value = EXPORT_API_URL_VAR || "apiUrl";
  tokenInput.value = EXPORT_ACCESS_TOKEN_VAR || "accessToken";

  overlay.hidden = false;
}

// Close settings overlay
function closeSettingsOverlay() {
  const overlay = document.getElementById("settings-overlay");
  if (overlay) overlay.hidden = true;
}

// Restore default settings in form
function restoreSettingsDefaultsInForm() {
  const apiUrlInput = document.getElementById("settings-api-url-var");
  const tokenInput = document.getElementById("settings-access-token-var");
  if (apiUrlInput) apiUrlInput.value = "apiUrl";
  if (tokenInput) tokenInput.value = "accessToken";
}

// Save settings from form to storage
async function saveSettingsFromForm() {
  const apiUrlInput = document.getElementById("settings-api-url-var");
  const tokenInput = document.getElementById("settings-access-token-var");
  if (!apiUrlInput || !tokenInput) return;

  let apiVar = apiUrlInput.value.trim() || "apiUrl";
  let tokenVar = tokenInput.value.trim() || "accessToken";

  // strip {{ }} if user pasted them
  apiVar = apiVar.replace(/^\{\{|\}\}$/g, "");
  tokenVar = tokenVar.replace(/^\{\{|\}\}$/g, "");

  EXPORT_API_URL_VAR = apiVar;
  EXPORT_ACCESS_TOKEN_VAR = tokenVar;

  await savePreferencesToStorage(apiVar, tokenVar);
  closeSettingsOverlay();
}

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

  await new Promise((resolve) => {
    chrome.storage.local.remove(storageKey, () => {
      resolve();
    });
  });

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

async function linksBoomi() {
  if (!(await isValidSession())) {
    alert("Requires a valid ADP WorkForce Manager session.");
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

  const ssoClientUrl = createSsoUrl(clienturl); // expected to end with '/'

  // If config is a template like "https://*.mykronos.com/ihub#..."
  // replace the host+slash with the real ssoClientUrl.
  let boomiURL = "";
  const templatePrefix = "https://*.mykronos.com/";
  if (boomiTemplate.startsWith(templatePrefix)) {
    boomiURL = ssoClientUrl + boomiTemplate.slice(templatePrefix.length);
  } else if (/^https?:\/\//i.test(boomiTemplate)) {
    // Full URL but not templated; use as-is
    boomiURL = boomiTemplate;
  } else {
    // Relative path (with or without leading slash)
    boomiURL = ssoClientUrl + boomiTemplate.replace(/^\/+/, "");
  }

  const incognito = await isIncognitoContext();
  if (incognito) {
    chrome.tabs.create({ url: boomiURL, active: true });
  } else {
    openURLNormally(boomiURL);
  }
}

// Install integrations button
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

  const ssoClientUrl = createSsoUrl(clienturl); // expected to end with '/'

  let installIntegrationsURL = "";
  const templatePrefix = "https://*.mykronos.com/";
  if (installTemplate.startsWith(templatePrefix)) {
    installIntegrationsURL = ssoClientUrl + installTemplate.slice(templatePrefix.length);
  } else if (/^https?:\/\//i.test(installTemplate)) {
    installIntegrationsURL = installTemplate;
  } else {
    installIntegrationsURL = ssoClientUrl + installTemplate.replace(/^\/+/, "");
  }

  const incognito = await isIncognitoContext();
  if (incognito) {
    chrome.tabs.create({ url: installIntegrationsURL, active: true });
  } else {
    openURLNormally(installIntegrationsURL);
  }
}

// Developer portal
async function linksDeveloperPortal() {
  try {
    const hermesData = await fetch("accesspanel.json").then((res) =>
      res.json()
    );
    const developerPortalURL = hermesData.details.urls.developerPortal;

    if (!developerPortalURL) {
      console.error("Developer Portal URL not found in accesspanel.json.");
      return;
    }

    const incognito = await isIncognitoContext();
    if (incognito) {
      chrome.tabs.create({ url: developerPortalURL, active: true });
    } else {
      window.open(developerPortalURL, "_blank");
    }
  } catch (error) {
    console.error("Failed to load Developer Portal URL:", error);
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
    const themeKey = result.selectedTheme || "WorkForce MGR"; // default theme
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

  menu.innerHTML = "";
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
    const hermesData = await fetch("accesspanel.json").then((res) =>
      res.json()
    );
    const aboutMessage = `
            Name: ${hermesData.name}
            Description: ${hermesData.details.description}
            Version: ${hermesData.details.version}
            Release Date: ${hermesData.details.release_date}
            Author: ${hermesData.details.author}`;

    // clean up the message to remove tabs
    const cleansedAboutMessage = aboutMessage.replace(/\t/g, "");
    alert(cleansedAboutMessage);
  } catch (error) {
    console.error("Failed to load About information:", error);
  }
}

// Support: open GitHub Issues (from accesspanel.json)
async function helpSupport() {
  try {
    const cfg = await fetch("accesspanel.json").then((res) => res.json());
    const issuesUrl = cfg?.details?.urls?.reportIssues;

    if (!issuesUrl) {
      alert("Support link is not configured.");
      return;
    }

    const ok = confirm("Open AccessPanel support (GitHub Issues) in a new tab?");
    if (!ok) return;

    // Open in a new tab from the extension UI context
    window.open(issuesUrl, "_blank", "noopener,noreferrer");
  } catch (error) {
    console.error("Failed to open support link:", error);
    alert("Failed to open support link.");
  }
}


// ===== LOCAL STORAGE FUNCTIONS ===== //
// Load client data from local storage
async function loadClientData() {
  return new Promise((resolve) => {
    chrome.storage.local.get([storageKey], (result) => {
      resolve(result[storageKey] || {});
    });
  });
}

// Save client data to local storage
async function saveClientData(data) {
  return new Promise((resolve) => {
    chrome.storage.local.set(
      {
        [storageKey]: data,
      },
      () => resolve()
    );
  });
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

  button.innerHTML = `Waiting... <span class="hourglass">${hourglassFrames[frameIndex]}</span>`;
  button.disabled = true;

  const hourglassSpan = button.querySelector(".hourglass");
  hourglassSpan.style.display = "inline-block";

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

  // restore access token timer
  const accessTokenTimerBox = document.getElementById("timer");
  if (clientData.accesstoken) {
    const expirationTime = new Date(clientData.expirationdatetime);
    if (currentDateTime < expirationTime) {
      const remainingSeconds = Math.floor(
        (expirationTime - currentDateTime) / 1000
      );
      startAccessTokenTimer(remainingSeconds, accessTokenTimerBox);
    } else {
      accessTokenTimerBox.textContent = "--:--";
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
// Tenant section collapsed
function toggleTenantSection() {
  const toggleButton = document.getElementById("toggle-tenant-section");
  const content = document.getElementById("tenant-section-content");
  const wrapper = content?.parentElement;
  if (!toggleButton || !content || !wrapper) return;

  const expanded = !content.classList.contains("expanded");
  content.classList.toggle("expanded", expanded);

  if (expanded) {
    wrapper.style.height = `${
      content.scrollHeight + toggleButton.offsetHeight
    }px`;
    toggleButton.textContent = "▲ Hide Tenant Information ▲";
  } else {
    wrapper.style.height = `${toggleButton.offsetHeight + 15}px`;
    toggleButton.textContent = "▼ Show Tenant Information ▼";
  }

  chrome.storage.local.set({ tenantSectionExpanded: expanded });
}

// Tenant section expanded
function restoreTenantSection() {
  chrome.storage.local.get("tenantSectionExpanded", (result) => {
    const isExpanded = !!result.tenantSectionExpanded;
    const toggleButton = document.getElementById("toggle-tenant-section");
    const content = document.getElementById("tenant-section-content");
    const wrapper = content?.parentElement;
    if (!toggleButton || !content || !wrapper) return;

    content.classList.toggle("expanded", isExpanded);
    if (isExpanded) {
      wrapper.style.height = `${
        content.scrollHeight + toggleButton.offsetHeight
      }px`;
      toggleButton.textContent = "▲ Hide Tenant Information ▲";
    } else {
      wrapper.style.height = `${toggleButton.offsetHeight + 15}px`;
      toggleButton.textContent = "▼ Show Tenant Information ▼";
    }
  });
}

// Pull tenant from TMS button
async function pullApiFromTmsClick() {
  const btn = document.getElementById("tms-pull-api");
  if (!btn) return;

  try {
    // ensure AccessPanel UI is in an active state (not wrong tab / mismatched TMS tenant)
    if (document.body.classList.contains("tab-inactive")) {
      alert(
        "Requires an active Tenant Management System (TMS) tenant matching your current AccessPanel link."
      );
      setButtonFailText(btn, "Invalid Session");
      return;
    }

    const [tab] = await chrome.tabs.query({
      active: true,
      currentWindow: true,
    });
    if (!tab || !tab.url) {
      alert("Unable to detect active browser tab.");
      setButtonFailText(btn, "No Active Tab");
      return;
    }

    // only allow on TMS vantage URLs
    let isTms = false;
    try {
      const u = new URL(tab.url);
      const origin = u.origin.toLowerCase();
      isTms =
        origin.startsWith("https://adpvantage.adp.com") ||
        origin.startsWith("https://testadpvantage.adp.com");
    } catch (e) {
      isTms = false;
    }

    if (!isTms) {
      alert(
        "Pull From TMS is only available when viewing the Tenant Management System (TMS) tenant popup."
      );
      setButtonFailText(btn, "Not on TMS");
      return;
    }

    // scrape values from the TMS popup in the active tab
    const execResults = await chrome.scripting.executeScript({
      target: { tabId: tab.id, allFrames: true },
      func: async () => {
        const found = {
          idInput: null,
          secretInput: null,
          revealSwitch: null,
          tenantAssertionInput: null,
        };

        const walk = (root) => {
          if (!root) return;
          const walker = document.createTreeWalker(
            root,
            NodeFilter.SHOW_ELEMENT,
            null
          );
          let node = walker.currentNode;
          while (node) {
            if (node.tagName === "SDF-INPUT") {
              if (
                !found.idInput &&
                node.id === "apiAccessClientID" &&
                node.shadowRoot
              ) {
                const idInput = node.shadowRoot.querySelector("input#input");
                if (idInput) found.idInput = idInput;
              } else if (
                !found.secretInput &&
                node.id === "apiAccessClientSecret" &&
                node.shadowRoot
              ) {
                const secretInput =
                  node.shadowRoot.querySelector("input#input");
                if (secretInput) found.secretInput = secretInput;
              }

              // look for the "Tenant SAML Assertion URL" field by its label text
              if (!found.tenantAssertionInput && node.shadowRoot) {
                const labelEl = node.shadowRoot.querySelector(
                  ".sdf-form-control-wrapper--label, label[part='label']"
                );
                if (labelEl && labelEl.textContent) {
                  const labelText = labelEl.textContent.trim();
                  if (labelText.startsWith("Tenant SAML Assertion URL")) {
                    const input = node.shadowRoot.querySelector("input#input");
                    if (input) {
                      found.tenantAssertionInput = input;
                    }
                  }
                }
              }
            } else if (
              !found.revealSwitch &&
              node.tagName === "SDF-SWITCH" &&
              node.id === "isRevealAPIAccessClientSecret"
            ) {
              found.revealSwitch = node;
            }

            if (node.shadowRoot) {
              walk(node.shadowRoot);
            }

            node = walker.nextNode();
          }
        };

        const readSecretState = (input) => {
          if (!input) {
            return { value: null, visible: false, disabled: false, type: "" };
          }
          const value = (input.value || "").trim();
          const disabled = !!input.disabled;
          const type = input.type || "";
          const visible = !disabled && type === "text" && !!value;
          return { value: value || null, visible, disabled, type };
        };

        walk(document);

        const clientId = found.idInput
          ? (found.idInput.value || "").trim()
          : null;

        // derive tenant ID from "Tenant SAML Assertion URL"
        let tenantId = null;
        if (found.tenantAssertionInput) {
          const rawUrl = (found.tenantAssertionInput.value || "").trim();
          if (rawUrl) {
            try {
              const url = new URL(rawUrl);
              const parts = url.pathname.split("/").filter(Boolean);
              const idx = parts.indexOf("authn");
              if (idx >= 0 && idx + 1 < parts.length) {
                tenantId = parts[idx + 1];
              }
            } catch (e) {
              console.error("Failed to parse Tenant SAML Assertion URL:", e);
            }
          }
        }

        let secretState = readSecretState(found.secretInput);
        let secretToggled = false;

        // if secret isn't visible yet but we have the Reveal toggle,
        // briefly toggle it on to read the value, then toggle it back off.
        if (!secretState.visible && found.revealSwitch) {
          const ariaChecked = found.revealSwitch.getAttribute("aria-checked");
          if (ariaChecked === "false") {
            // toggle reveal ON
            found.revealSwitch.click();
            secretToggled = true;

            // give the UI a moment to update
            await new Promise((resolve) => setTimeout(resolve, 200));
            secretState = readSecretState(found.secretInput);

            // toggle reveal back OFF so the UI returns to its original state
            found.revealSwitch.click();
          }
        }

        return {
          clientId,
          clientSecret: secretState.value,
          secretVisible: secretState.visible,
          secretToggled,
          tenantId,
        };
      },
    });

    // Pick the first frame that actually returned something
    let data = null;
    if (Array.isArray(execResults)) {
      for (const r of execResults) {
        if (r && r.result && (r.result.clientId || r.result.clientSecret)) {
          data = r.result;
          break;
        }
      }
    }

    if (!data) {
      setButtonFailText(btn, "No Data");
      alert("Unable to locate the API Access fields in the TMS popup.");
      return;
    }

    let anySuccess = false;
    const messages = [];

    // --- Handle Client ID ---
    if (data.clientId) {
      const idInput = document.getElementById("client-id");
      if (idInput) {
        idInput.value = data.clientId;
        try {
          await saveClientIDClick(); // reuse your existing save flow
          anySuccess = true;
          messages.push("Client ID pulled from TMS and saved in AccessPanel.");
        } catch (e) {
          console.error("Failed to save Client ID after TMS pull:", e);
          messages.push(
            "Client ID was pulled from TMS, but saving in AccessPanel failed. Check the console logs."
          );
        }
      } else {
        messages.push(
          "Client ID was found in TMS, but the AccessPanel Client ID field is not available."
        );
      }
    } else {
      messages.push("Client ID was not found in the TMS popup or is empty.");
    }

    // --- Handle Client Secret ---
    if (data.clientSecret && data.secretVisible) {
      const secretInput = document.getElementById("client-secret");
      if (secretInput) {
        secretInput.value = data.clientSecret;
        try {
          await saveClientSecretClick();
          anySuccess = true;
          messages.push(
            "Client Secret pulled from TMS and saved in AccessPanel."
          );
        } catch (e) {
          console.error("Failed to save Client Secret after TMS pull:", e);
          messages.push(
            "Client Secret was pulled from TMS, but saving in AccessPanel failed. Check the console logs."
          );
        }
      } else {
        messages.push(
          "Client Secret was found in TMS, but the AccessPanel Client Secret field is not available."
        );
      }
    } else if (data.secretToggled && !data.clientSecret) {
      messages.push(
        "Tried to reveal the Client Secret in TMS, but it still appeared empty. You may need to reveal it manually and try again."
      );
    } else if (!data.clientSecret) {
      messages.push(
        "Client Secret was not found in the TMS popup or is empty."
      );
    } else if (!data.secretVisible) {
      messages.push(
        "Client Secret appears to be masked in TMS and could not be read. Please use the Reveal toggle in TMS and try again."
      );
    }

    // --- Handle Tenant ID ---
    if (data.tenantId) {
      const tenantIdInput = document.getElementById("tenant-id");
      if (tenantIdInput) {
        tenantIdInput.value = data.tenantId;
        try {
          await saveTenantIdClick(); // reuse your existing save flow
          anySuccess = true;
          messages.push("Tenant ID pulled from TMS and saved in AccessPanel.");
        } catch (e) {
          console.error("Failed to save Tenant ID after TMS pull:", e);
          messages.push(
            "Tenant ID was pulled from TMS, but saving in AccessPanel failed. Check the console logs."
          );
        }
      } else {
        messages.push(
          "Tenant ID was found in TMS, but the AccessPanel Tenant ID field is not available."
        );
      }
    } else {
      messages.push(
        "Tenant ID was not found in the TMS popup or could not be derived from the SAML Assertion URL."
      );
    }

    if (anySuccess) {
      setButtonTempText(btn, "Pulled From TMS");
    } else {
      setButtonFailText(btn, "No Data");
    }

    alert(messages.join("\n"));
  } catch (e) {
    console.error("Pull From TMS failed:", e);
    setButtonFailText(btn, "Error");
    alert("Pull From TMS failed. Check the console logs for more details.");
  }
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
    // quick visual feedback that the button was clicked
    setLabel("Generating...");

    // --- 1. basic field checks (Client ID / Secret / Tenant ID) ---
    const clientIdInput = document.getElementById("client-id");
    const clientSecretInput = document.getElementById("client-secret");
    const tenantIdInput = document.getElementById("tenant-id");

    const clientId = clientIdInput?.value.trim() || "";
    const clientSecret = clientSecretInput?.value.trim() || "";
    const tenantId = tenantIdInput?.value.trim() || "";

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

    // --- 2. get the current client URL and vanity host ---
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
    } catch (e) {
      console.error("Failed to parse vanity URL from Tenant URL.", e);
      alert(
        "Unable to parse a valid vanity hostname from the Tenant URL.\n" +
          "Expected something like https://<tenant>.mykronos.com"
      );
      setLabel("Error", 1500);
      return;
    }

    // --- 3. attempt to fetch a fresh access token ---
    // (this is async; we just start it and then poll storage)
    void fetchToken();

    // --- 4. wait for tokens to appear in storage ---
    const waitForTokens = async (clientUrl, maxWaitMs = 5000, intervalMs = 250) => {
      const start = Date.now();

      while (Date.now() - start < maxWaitMs) {
        const data = await loadClientData();
        const client = data[clientUrl] || {};

        if (client.accesstoken && client.refreshtoken) {
          return client;
        }

        await new Promise((resolve) => setTimeout(resolve, intervalMs));
      }

      // final check before giving up
      const data = await loadClientData();
      return data[clientUrl] || {};
    };

    const thisClient = await waitForTokens(clientUrl);
    const accessToken = thisClient.accesstoken || "";
    const refreshToken = thisClient.refreshtoken || "";

    if (!accessToken || !refreshToken) {
      alert("Could not retrieve access/refresh token after requesting one.");
      setLabel("Error", 1500);
      return;
    }

    const editDate = thisClient.editdatetime || new Date().toISOString();

    // --- 5. build properties file contents ---
    const propsLines = [
      "report.api.execute.for.external.client=true",
      "report.api.gateway.access.token.appkey=", // intentionally blank
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

    // keep a reference for potential future use if needed
    window.lastBirtPropertiesText = propsText;
    window.lastBirtPropertiesMeta = {
      clientUrl,
      tenantId,
      editDate,
    };

    // --- 6. open popup window with BIRT properties + buttons ---
    const escapedPropsText = propsText
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");

    const rootStyles = getComputedStyle(document.documentElement);
    const primary = rootStyles.getPropertyValue("--primary-color").trim();
    const secondary = rootStyles.getPropertyValue("--secondary-color").trim();
    const accent = rootStyles.getPropertyValue("--accent-color").trim();
    const highlight = rootStyles.getPropertyValue("--highlight-color").trim();
    const textOnBtn = rootStyles.getPropertyValue("--buttontext-color").trim();

    const popupHtml = `
      <html>
        <head>
          <title>BIRT Properties</title>
          <style>
            body {
              font-family: Arial, sans-serif;
              margin: 0;
              padding: 16px;
              background-color: ${primary};
              line-height: 1.5;
            }
            h1 {
              font-size: 1.4rem;
              font-weight: bold;
              margin-bottom: 0.25rem;
              color: ${accent};
            }
            .meta {
              font-size: 0.85rem;
              font-weight: bold;
              color: ${accent};
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
              border: 1px solid ${secondary};
              background-color: ${accent};
              color: ${textOnBtn};
              cursor: pointer;
            }
            button:hover {
              background-color: ${highlight};
            }
            pre {
              background: ${textOnBtn};
              border: 2px solid ${accent};
              radius: 6px;
              padding: 10px;
              white-space: pre;
              overflow-x: auto;
              font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace;
              font-size: 0.9rem;
              color: ${accent};
            }
          </style>
        </head>
        <body>
          <h1>BIRT Properties</h1>
          <div class="meta">
            Tenant: ${tenantId}<br/>
            Client URL: ${clientUrl}<br/>
            Last Edit: ${editDate}
          </div>
          <div class="btn-row">
            <button id="copy-birt">Copy To Clipboard</button>
            <button id="download-birt">Download To File</button>
          </div>
          <pre id="birt-text">${escapedPropsText}</pre>
        </body>
      </html>
    `;

    const w = window.open(
      "",
      "_blank",
      "width=800,height=600,scrollbars=yes,resizable=yes"
    );
    if (!w) {
      alert(
        "Unable to open BIRT popup window. Please allow popups for this extension."
      );
      setLabel("Error", 1500);
      return;
    }

    w.document.write(popupHtml);
    w.document.close();

    // wire copy + download in the child window (no inline script)
    const setupChildHandlers = () => {
      const copyBtn = w.document.getElementById("copy-birt");
      const dlBtn = w.document.getElementById("download-birt");
      const pre = w.document.getElementById("birt-text");

      if (!copyBtn || !dlBtn || !pre) return;

      copyBtn.addEventListener("click", async () => {
        try {
          // make sure the POPUP is the focused document
          try {
            w.focus();
          } catch (_) {}

          const text = pre.innerText || pre.textContent || "";
          if (!text.trim()) {
            copyBtn.textContent = "No Content";
            setTimeout(() => (copyBtn.textContent = "Copy To Clipboard"), 1500);
            return;
          }

          // prefer clipboard API from the POPUP window
          if (w.navigator?.clipboard?.writeText) {
            await w.navigator.clipboard.writeText(text);
          } else {
            // fallback: execCommand copy (must be created/selected in POPUP document)
            const ta = w.document.createElement("textarea");
            ta.value = text;
            ta.style.position = "fixed";
            ta.style.left = "-9999px";
            w.document.body.appendChild(ta);
            ta.focus();
            ta.select();
            w.document.execCommand("copy");
            w.document.body.removeChild(ta);
          }

          copyBtn.textContent = "Copied!";
          setTimeout(() => (copyBtn.textContent = "Copy To Clipboard"), 1500);
        } catch (err) {
          console.error("Copy BIRT properties failed:", err);
          copyBtn.textContent = "Copy Failed";
          setTimeout(() => (copyBtn.textContent = "Copy To Clipboard"), 1500);
        }
      });

      dlBtn.addEventListener("click", async () => {
        try {
          const text = pre.innerText || pre.textContent || "";
          const defaultFileName = "custom_reportplugin.properties";

          if (!text.trim()) {
            dlBtn.textContent = "No Content";
            setTimeout(() => (dlBtn.textContent = "Download To File"), 1500);
            return;
          }

          if (w.showSaveFilePicker) {
            try {
              const fileHandle = await w.showSaveFilePicker({
                suggestedName: defaultFileName,
                types: [
                  {
                    description: "Properties Files",
                    accept: { "text/plain": [".properties"] },
                  },
                ],
              });
              const writable = await fileHandle.createWritable();
              await writable.write(text);
              await writable.close();
            } catch (err) {
              if (err && err.name === "AbortError") {
                return; // user cancelled – silent exit
              }
              // fallback to blob download in popup context
              downloadFileWithContext(
                w.document,
                w.URL,
                defaultFileName,
                text,
                "text/plain"
              );
            }
          } else {
            // fallback: blob download in popup context
            downloadFileWithContext(
              w.document,
              w.URL,
              defaultFileName,
              text,
              "text/plain"
            );
          }
        } catch (err) {
          console.error("Download BIRT properties failed:", err);
          alert("Failed to download the BIRT properties file.");
        }
      });
    };

    if (w.document.readyState === "complete") {
      setupChildHandlers();
    } else {
      w.addEventListener("load", setupChildHandlers);
    }

    // success: brief positive feedback, then reset
    setLabel("Generated", 1500);
  } catch (e) {
    console.error("Generate BIRT Properties failed:", e);
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
  try {
    await populateClientUrlField();
    const val = (document.getElementById("client-url") || {}).value || "";
    if (val) {
      setButtonTempText(btn, "URL Refreshed");
    } else {
      setButtonFailText(btn, "No URL Detected");
    }
  } catch (e) {
    console.error(e);
    setButtonFailText(btn, "Refresh Failed");
  }
}

// Copy URL button
async function copyClientUrlClick() {
  const btn = document.getElementById("copy-client-url");
  try {
    const val = (document.getElementById("client-url") || {}).value || "";
    if (!val) {
      setButtonFailText(btn, "No URL to Copy");
      return;
    }

    // use clipboard api; fallback if needed
    if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(val);
    } else {
      // fallback approach
      const ta = document.createElement("textarea");
      ta.value = val;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand("copy");
      document.body.removeChild(ta);
    }

    setButtonTempText(btn, "URL Copied");
  } catch (e) {
    console.error("Copy failed:", e);
    setButtonFailText(btn, "Copy Failed");
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

  try {
    if (!(await isValidSession())) {
      alert("Requires a valid ADP Workforce Manager session.");
      return;
    }

    const clienturl = await getClientUrl();
    if (!clienturl) {
      return;
    }

    const clientid = document.getElementById("client-id").value.trim();
    if (!clientid) {
      alert("Client ID cannot be empty!");
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
    setButtonTempText(button, "Client ID Saved!");
  } catch (error) {
    console.error("Failed to save Client ID:", error);
    setButtonFailText(button, "Save Failed!");
  }
}

// Copy client ID button
async function copyClientIdClick() {
  const btn = document.getElementById("copy-client-id");
  try {
    const input = document.getElementById("client-id");
    const val = input && input.value ? input.value.trim() : "";

    if (!val) {
      setButtonFailText(btn, "No Client ID");
      return;
    }

    if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(val);
    } else {
      const ta = document.createElement("textarea");
      ta.value = val;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand("copy");
      document.body.removeChild(ta);
    }

    setButtonTempText(btn, "Client ID Copied");
  } catch (e) {
    console.error("Copy Client ID failed:", e);
    setButtonFailText(btn, "Copy Failed");
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

  try {
    if (!(await isValidSession())) {
      alert("Requires a valid ADP Workforce Manager session.");
      return;
    }

    const clienturl = await getClientUrl();
    if (!clienturl) {
      return;
    }

    const clientsecret = document.getElementById("client-secret").value.trim();
    if (!clientsecret) {
      alert("Client Secret cannot be empty!");
      return;
    }

    const data = await loadClientData();
    data[clienturl] = {
      ...(data[clienturl] || {}),
      clientsecret: clientsecret,
      editdatetime: new Date().toISOString(),
    };

    await saveClientData(data);
    setButtonTempText(button, "Client Secret Saved!");
  } catch (error) {
    console.error("Failed to save Client Secret:", error);
    setButtonFailText(button, "Save Failed!");
  }
}

// Copy client secret button
async function copyClientSecretClick() {
  const btn = document.getElementById("copy-client-secret");
  try {
    const input = document.getElementById("client-secret");
    const val = input && input.value ? input.value.trim() : "";

    if (!val) {
      setButtonFailText(btn, "No Client Secret");
      return;
    }

    if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(val);
    } else {
      const ta = document.createElement("textarea");
      ta.value = val;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand("copy");
      document.body.removeChild(ta);
    }

    setButtonTempText(btn, "Client Secret Copied");
  } catch (e) {
    console.error("Copy Client Secret failed:", e);
    setButtonFailText(btn, "Copy Failed");
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

  try {
    if (!(await isValidSession())) {
      alert("Requires a valid ADP Workforce Manager session.");
      return;
    }

    const clienturl = await getClientUrl();
    if (!clienturl) {
      return;
    }

    const tenantIdInput = document.getElementById("tenant-id");
    if (!tenantIdInput) {
      console.error("tenant-id input not found.");
      return;
    }

    const tenantId = tenantIdInput.value.trim();
    if (!tenantId) {
      alert("Tenant ID cannot be empty.");
      return;
    }

    const data = await loadClientData();
    data[clienturl] = {
      ...(data[clienturl] || {}),
      tenantid: tenantId,
      editdatetime: new Date().toISOString(),
    };

    await saveClientData(data);
    setButtonTempText(button, "Tenant ID Saved!");
  } catch (error) {
    console.error("Failed to save Tenant ID:", error);
    setButtonFailText(button, "Save Failed!");
  }
}

// Copy tenant ID button
async function copyTenantIdClick() {
  const btn = document.getElementById("copy-tenant-id");
  try {
    const input = document.getElementById("tenant-id");
    const val = input && input.value ? input.value.trim() : "";

    if (!val) {
      setButtonFailText(btn, "No Tenant ID");
      return;
    }

    if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(val);
    } else {
      const ta = document.createElement("textarea");
      ta.value = val;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand("copy");
      document.body.removeChild(ta);
    }

    setButtonTempText(btn, "Tenant ID Copied");
  } catch (e) {
    console.error("Copy Tenant ID failed:", e);
    setButtonFailText(btn, "Copy Failed");
  }
}
// ==================================== //


// ===== API TOKEN UI FIELDS AND BUTTONS ===== //
// Access token section collapsed
function toggleAccessSection() {
  const toggleButton = document.getElementById("toggle-access-section");
  const content = document.getElementById("access-section-content");
  const wrapper = content?.parentElement;
  if (!toggleButton || !content || !wrapper) return;

  const expanded = !content.classList.contains("expanded");
  content.classList.toggle("expanded", expanded);

  if (expanded) {
    wrapper.style.height = `${
      content.scrollHeight + toggleButton.offsetHeight
    }px`;
    toggleButton.textContent = "▲ Hide API Token Options ▲";
  } else {
    wrapper.style.height = `${toggleButton.offsetHeight + 15}px`;
    toggleButton.textContent = "▼ Show API Token Options ▼";
  }

  chrome.storage.local.set({ accessSectionExpanded: expanded });
}

// Access token section expanded
function restoreAccessSection() {
  chrome.storage.local.get("accessSectionExpanded", (result) => {
    const isExpanded = !!result.accessSectionExpanded;
    const toggleButton = document.getElementById("toggle-access-section");
    const content = document.getElementById("access-section-content");
    const wrapper = content?.parentElement;
    if (!toggleButton || !content || !wrapper) return;

    content.classList.toggle("expanded", isExpanded);
    if (isExpanded) {
      wrapper.style.height = `${
        content.scrollHeight + toggleButton.offsetHeight
      }px`;
      toggleButton.textContent = "▲ Hide API Token Options ▲";
    } else {
      wrapper.style.height = `${toggleButton.offsetHeight + 15}px`;
      toggleButton.textContent = "▼ Show API Token Options ▼";
    }
  });
}

// Populate access token field and start timer if token is valid
async function populateAccessToken() {
  const clienturl = await getClientUrl();
  const accessTokenBox = document.getElementById("access-token");
  const timerBox = document.getElementById("timer");

  if (!clienturl) {
    accessTokenBox.value = "Requires WFMgr Login";
    timerBox.textContent = "--:--";
    return;
  }

  const data = await loadClientData();
  const currentDateTime = new Date();

  if (data[clienturl]?.accesstoken) {
    const expirationTime = new Date(data[clienturl].expirationdatetime);

    if (currentDateTime > expirationTime) {
      accessTokenBox.value = "Access Token Expired";
      timerBox.textContent = "--:--";
    } else {
      accessTokenBox.value = data[clienturl].accesstoken;

      // calculate remaining time and start the timer
      const remainingSeconds = Math.floor(
        (expirationTime - currentDateTime) / 1000
      );
      startAccessTokenTimer(remainingSeconds, timerBox);
    }
  } else {
    accessTokenBox.value = "Get New Access Token";
    timerBox.textContent = "--:--";
  }
}

// Get access token button
async function fetchToken() {
  const clienturl = await getClientUrl();
  if (!clienturl || !(await isValidSession())) {
    alert("Requires a valid ADP Workforce Manager session.");
    return;
  }

  const clientID = document.getElementById("client-id").value.trim();
  if (!clientID) {
    alert("Please enter a Client ID first.");
    return;
  }

  const tokenurl = `${clienturl}accessToken?clientId=${clientID}`;

  const incognito = await isIncognitoContext();
  if (incognito) {
    retrieveTokenViaNewTab(tokenurl);
  } else {
    fetchTokenDirectly(tokenurl, clienturl, clientID);
  }
}

// Get access token normal mode (used by fetchToken())
async function fetchTokenDirectly(tokenurl, clienturl, clientID) {
  try {
    const response = await fetch(tokenurl, {
      method: "GET",
      credentials: "include",
      headers: {
        Accept: "application/json",
      },
    });

    if (!response.ok) {
      throw new Error(`Failed to fetch token. HTTP status: ${response.status}`);
    }

    const result = await response.json();
    processTokenResponse(result, clienturl, clientID, tokenurl);
  } catch (error) {
    console.error("Error fetching token:", error);
    alert(`Failed to fetch token: ${error.message || error}`);
  }
}

// Get access token incognito mode (used by fetchToken())
async function retrieveTokenViaNewTab(tokenurl) {
  chrome.tabs.create({ url: tokenurl, active: false }, async (tab) => {
    if (chrome.runtime.lastError || !tab?.id) {
      console.error("Failed to create token tab:", chrome.runtime.lastError);
      alert("Failed to open token tab.");
      return;
    }

    const tabId = tab.id;

    const cleanupAndClose = () => {
      try {
        chrome.tabs.onUpdated.removeListener(onUpdated);
      } catch (e) {
        // ignore
      }

      chrome.tabs.remove(tabId, () => {
        const err = chrome.runtime.lastError?.message || "";
        if (err && !err.toLowerCase().includes("no tab with id")) {
          console.warn("Error closing tab:", chrome.runtime.lastError);
        }
      });
    };

    const onUpdated = async (updatedTabId, changeInfo) => {
      if (updatedTabId !== tabId) return;
      if (changeInfo.status !== "complete") return;

      // stop listening as soon as the tab is complete
      chrome.tabs.onUpdated.removeListener(onUpdated);

      chrome.scripting.executeScript(
        {
          target: { tabId },
          function: scrapeTokenFromPage,
        },
        async (injectionResults) => {
          if (chrome.runtime.lastError) {
            alert("Failed to retrieve token (script injection failed).");
            cleanupAndClose();
            return;
          }

          const payload = injectionResults?.[0]?.result;

          if (!payload?.ok) {
            alert("Failed to retrieve token from the page.");
            cleanupAndClose();
            return;
          }

          // payload.text is the raw <pre> contents; parse in extension context
          let tokenJson;
          try {
            tokenJson = JSON.parse(payload.text);
          } catch (e) {
            console.error("Token page did not return valid JSON.");
            alert("Token response was not valid JSON.");
            cleanupAndClose();
            return;
          }

          // retrieve the existing client ID from storage
          const baseClientUrl = new URL(tokenurl).origin + "/";
          const storedData = await loadClientData();
          const existingClientID =
            storedData?.[baseClientUrl]?.clientid || "unknown-client";

          processTokenResponse(
            tokenJson,
            baseClientUrl,
            existingClientID,
            tokenurl
          );

          cleanupAndClose();
        }
      );
    };

    // Listen for tab load completion instead of using a fixed timeout
    chrome.tabs.onUpdated.addListener(onUpdated);
  });
}

// Scrape token from browser tab (used by retrieveTokenViaNewTab())
function scrapeTokenFromPage() {
  try {
    const preElement = document.querySelector("pre");
    if (!preElement) {
      return { ok: false, error: "Token <pre> element not found." };
    }

    const text = preElement.innerText?.trim();
    if (!text) {
      return { ok: false, error: "Token page <pre> was empty." };
    }

    // Return raw text; parse JSON in the extension context
    return { ok: true, text };
  } catch (error) {
    return { ok: false, error: String(error) };
  }
}

// Process Token (used by fetchTokenDirectly() and retrieveTokenViaNewTab())
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
      chrome.storage.local.remove("accessTokenTimer");
    } else {
      const minutes = Math.floor(remainingTime / 60);
      const seconds = remainingTime % 60;
      timerBox.textContent = `${String(minutes).padStart(2, "0")}:${String(
        seconds
      ).padStart(2, "0")}`;
      remainingTime--;

      // save the remaining time to storage
      chrome.storage.local.set({
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

  // clear timer from storage
  //chrome.storage.local.remove("accessTokenTimer");
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
  const accessTokenBox = document.getElementById("access-token");
  const accessToken = accessTokenBox?.value;

  // validate access token before copying
  if (
    !accessToken ||
    accessToken === "Get Token" ||
    accessToken === "Get New Access Token" ||
    accessToken === "Access Token Expired"
  ) {
    return;
  }

  // copy token to clipboard
  navigator.clipboard
    .writeText(accessToken)
    .then(() => {
      // visual feedback: change button text
      const button = document.getElementById("copy-token");
      const originalText = button.textContent;

      button.textContent = "Copied!";
      button.disabled = true; // disable temporarily

      setTimeout(() => {
        button.textContent = originalText;
        button.disabled = false; // re-enable
      }, 2000);
    })
    .catch((error) => {
      console.error("Failed to copy Access Token:", error);
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

    // update local storage
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
      chrome.storage.local.set({
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

  // clear timer from storage
  //chrome.storage.local.remove("refreshTokenTimer");
}

// Copy refresh token button
function copyRefreshToken() {
  const button = document.getElementById("copy-refresh-token");
  const refreshTokenBox = document.getElementById("refresh-token");
  const refreshToken = refreshTokenBox?.value;

  // validate refresh token before copying
  if (
    !refreshToken ||
    refreshToken === "Refresh Token" ||
    refreshToken === "Get New Access Token" ||
    refreshToken === "Refresh Token Expired"
  ) {
    setButtonFailText(button, "No Token!");
    return;
  }

  // copy refresh token to clipboard
  navigator.clipboard
    .writeText(refreshToken)
    .then(() => {
      setButtonTempText(button, "Copied!");
    })
    .catch((error) => {
      console.error("Failed to copy Refresh Token:", error);
      setButtonFailText(button, "Copy Failed!");
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
const MYAPIS_KEY = "hermes_myapis"; // [{id,name,method,endpoint,body,createdAt,updatedAt}]
let EXPORT_API_URL_VAR = "apiUrl";
let EXPORT_ACCESS_TOKEN_VAR = "accessToken";

// Toggle API library visibility section collapsed / expanded
function toggleApiLibrary() {
  const toggleButton = document.getElementById("toggle-api-library");
  const content = document.getElementById("api-library-content");
  const wrapper = content.parentElement;

  // toggle expanded/collapsed state
  const isExpanded = content.classList.toggle("expanded");

  // dynamically calculate the height
  if (isExpanded) {
    wrapper.style.height = `${
      content.scrollHeight + toggleButton.offsetHeight
    }px`;
    toggleButton.textContent = "▲ Hide API Library ▲";
  } else {
    wrapper.style.height = `${toggleButton.offsetHeight + 15}px`;
    toggleButton.textContent = "▼ Show API Library ▼";
  }

  // persist the state in local storage
  chrome.storage.local.set({ apiLibraryExpanded: isExpanded });
}

// Restore API library visibility on load
function restoreApiLibrary() {
  chrome.storage.local.get("apiLibraryExpanded", (result) => {
    const isExpanded = result.apiLibraryExpanded || false;
    const toggleButton = document.getElementById("toggle-api-library");
    const content = document.getElementById("api-library-content");
    const wrapper = content.parentElement;

    // set initial state based on stored value
    if (isExpanded) {
      content.classList.add("expanded");
      wrapper.style.height = `${
        content.scrollHeight + toggleButton.offsetHeight
      }px`;
      toggleButton.textContent = "▲ Hide API Library ▲";
    } else {
      content.classList.remove("expanded");
      wrapper.style.height = `${toggleButton.offsetHeight + 15}px`;
      toggleButton.textContent = "▼ Show API Library ▼";
    }
  });
}

// Load saved My API's
async function getSavedMyApis() {
  return new Promise((resolve) => {
    chrome.storage.local.get([MYAPIS_KEY], (res) =>
      resolve(res[MYAPIS_KEY] || [])
    );
  });
}

// Load saved My API entry by selector value "myapi:<id>"
async function loadSavedEntryByKey(selectedKey) {
  if (!selectedKey || !selectedKey.startsWith("myapi:")) return null;
  const id = selectedKey.slice(6);
  const { hermes_myapis } = await new Promise((resolve) =>
    chrome.storage.local.get(["hermes_myapis"], resolve)
  );
  const list = hermes_myapis || [];
  return list.find((x) => x.id === id) || null;
}

// Save My API entry to user local storage
async function setSavedMyApis(list) {
  return new Promise((resolve) =>
    chrome.storage.local.set({ [MYAPIS_KEY]: list }, resolve)
  );
}

// Generate unique key for My API entry
function genId() {
  return "a" + Math.random().toString(36).slice(2) + Date.now().toString(36);
}

// Clear DevLink
function clearDevLinkBanner() {
  const a = document.getElementById("api-devlink");
  if (!a) return;
  a.hidden = true;
  a.removeAttribute("href");
  a.textContent = "";
}

// Populate public API library
async function populateApiDropdownPublic() {
  const sel = document.getElementById("api-selector");
  if (!sel) return;
  HERMES_MYAPIS_MODE = false;
  sel.innerHTML = "";
  clearDevLinkBanner();

  // switch-to-my-api
  const toMy = document.createElement("option");
  toMy.value = "__VIEW_MY_APIS__";
  toMy.textContent = "View My APIs…";
  sel.appendChild(toMy);

  // placeholder for public api drop down
  sel.insertAdjacentHTML(
    "beforeend",
    '<option value="" disabled selected>Select Public API...</option>'
  );

  // load and list public library
  const resp = await fetch("apilibrary/apilibrary.json");
  if (!resp.ok)
    throw new Error(`Failed to fetch API library. HTTP ${resp.status}`);
  const data = await resp.json();
  const lib = data.apiLibrary || {};

  for (const key in lib) {
    if (key.startsWith("_")) continue;
    const opt = document.createElement("option");
    opt.value = key;
    opt.textContent = lib[key].name;
    sel.appendChild(opt);
  }
}

// Populate My API List Dropdown
async function populateApiDropdownMyApis() {
  const sel = document.getElementById("api-selector");
  if (!sel) return;
  HERMES_MYAPIS_MODE = true;
  sel.innerHTML = "";
  clearDevLinkBanner();

  // switch-to-public
  const toPub = document.createElement("option");
  toPub.value = "__VIEW_PUBLIC_APIS__";
  toPub.textContent = "View Public APIs…";
  sel.appendChild(toPub);

  // manage My APIs
  const manage = document.createElement("option");
  manage.value = "__MANAGE_MY_APIS__";
  manage.textContent = "Manage My APIs…";
  sel.appendChild(manage);

  // placeholder
  sel.insertAdjacentHTML(
    "beforeend",
    '<option value="" disabled selected>Select My API…</option>'
  );

  // ---- Ad-Hoc entries available directly in My APIs mode ----
  const adhocGet = document.createElement("option");
  adhocGet.value = "adHocGet";
  adhocGet.textContent = "Ad-Hoc GET";
  sel.appendChild(adhocGet);

  const adhocPost = document.createElement("option");
  adhocPost.value = "adHocPost";
  adhocPost.textContent = "Ad-Hoc POST";
  sel.appendChild(adhocPost);

  const sep = document.createElement("option");
  sep.value = "__SEP__";
  sep.textContent = "────────";
  sep.disabled = true;
  sel.appendChild(sep);

  // load saved list
  const items = await getSavedMyApis();

  if (!items.length) {
    const empty = document.createElement("option");
    empty.value = "__EMPTY__";
    empty.textContent = "(No saved APIs yet)";
    empty.disabled = true;
    sel.appendChild(empty);
    return;
  }

  // append saved entries
  items.forEach((item) => {
    const opt = document.createElement("option");
    opt.value = `myapi:${item.id}`;
    opt.textContent = `${item.name} (${item.method})`;
    opt.title = `${item.method} ${item.endpoint}`;
    sel.appendChild(opt);
  });
}

// Wrapper to handle public vs My API's
async function populateApiDropdown() {
  if (HERMES_MYAPIS_MODE) return populateApiDropdownMyApis();
  return populateApiDropdownPublic();
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

// Clear Existing Parameters Button
function clearParameters() {
  const queryContainer = document.getElementById("query-parameters-container");
  const bodyContainer = document.getElementById("body-parameters-container");
  const pathContainer = document.getElementById("path-parameters-container");

  if (queryContainer) {
    queryContainer.innerHTML = ""; // clear query parameters
  }
  if (bodyContainer) {
    bodyContainer.innerHTML = ""; // clear body parameters
  }
  if (pathContainer) {
    pathContainer.innerHTML = ""; // clear path parameters
  }
}

// Show / clear the Dev Portal Link Banner for a given API object (public API's only)
function renderDevLinkBanner(selectedApiKey, apiObjOrNull) {
  const a = document.getElementById("api-devlink");
  if (!a) return;

  // hide for my apis & ad-hoc and for missing devlink data
  if (
    !apiObjOrNull ||
    selectedApiKey.startsWith("myapi:") ||
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

  host.innerHTML = "";

  // ad-hoc modes don’t use path param ui
  if (
    selectedApiKey === "adHocGet" ||
    selectedApiKey === "adHocPost" ||
    selectedApiKey?.startsWith("myapi:")
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
    queryContainer.innerHTML = ""; // clear existing parameters

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
        ph.textContent = param.description || /*"Select an option"*/ "";
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
        ph.textContent = param.description || /*"Select an option"*/ "";
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

    bodyParamContainer.innerHTML = "";

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
          param.description || /*"Select an option"*/ "";
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
          param.description || /*"Select an option"*/ "";
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

// Wait for new access token if needed
async function waitForUpdatedToken(clienturl, maxRetries = 5, delayMs = 1000) {
  for (let i = 0; i < maxRetries; i++) {
    await new Promise((resolve) => setTimeout(resolve, delayMs)); // wait for storage update
    let updatedData = await loadClientData();
    let updatedClientData = updatedData[clienturl] || {};

    if (updatedClientData.accesstoken) {
      return updatedClientData;
    }

    // console.info(`Retry ${i + 1}/${maxRetries}: Token not available yet...`);
  }

  throw new Error(
    "Failed to retrieve updated access token after multiple attempts."
  );
}

// Clear response UI for new data V1
function clearApiResponse() {
  const responseSection = document.getElementById("response-section");
  if (responseSection) {
    responseSection.innerHTML = "<pre>Awaiting API Response...</pre>";
  }

  // clear cached response object
  window.lastApiResponseObject = null;

  // disable response-dependent buttons (Download / Copy)
  if (typeof updateResponseDependentButtons === "function") {
    updateResponseDependentButtons(false);
  }
}

// API entry selector
async function handleApiSelection(selectedKey) {
  // whenever a new API is selected, reset buttons and last request
  if (typeof updateRequestDependentButtons === "function") {
    updateRequestDependentButtons(false);
  }
  if (typeof updateResponseDependentButtons === "function") {
    updateResponseDependentButtons(false);
  }
  if (typeof clearApiResponse === "function") {
    clearApiResponse();
  } else {
    const responseSection = document.getElementById("response-section");
    if (responseSection) {
      responseSection.innerHTML = "<pre>Awaiting API Response...</pre>";
    }
  }
  // clear the lastRequestDetails so view/save/export always correspond to a *fresh* request for this selection.
  lastRequestDetails = null;

  if (selectedKey === "__VIEW_MY_APIS__") {
    await populateApiDropdownMyApis();
    clearParameters?.();
    clearDevLinkBanner();
    applyDynamicStyles?.();
    return;
  }
  if (selectedKey === "__VIEW_PUBLIC_APIS__") {
    await populateApiDropdownPublic();
    clearDevLinkBanner();
    clearParameters?.();
    applyDynamicStyles?.();
    return;
  }
  // open my apis manager overlay
  if (selectedKey === "__MANAGE_MY_APIS__") {
    await openMyApisManager();
    // reset selector back to placeholder so it doesn't look like a real API
    const sel = document.getElementById("api-selector");
    if (sel) sel.value = "";
    return;
  }
  if (!selectedKey || selectedKey === "__EMPTY__") return;

  clearParameters?.();

  // saved item?
  if (HERMES_MYAPIS_MODE && selectedKey.startsWith("myapi:")) {
    const id = selectedKey.slice(6);
    const list = await getSavedMyApis();
    const item = list.find((x) => x.id === id);
    // render like Ad-Hoc
    await renderSavedAsAdHoc(item);
    applyDynamicStyles?.();
    return;
  }

  // public library item
  await populatePathParameters(selectedKey);
  await populateQueryParameters(selectedKey);
  await populateBodyParameters(selectedKey);
  applyDynamicStyles?.();
}

// Helper to render saved My API into ad-hoc UI
async function renderSavedAsAdHoc(item) {
  const queryContainer = document.getElementById("query-parameters-container");
  const bodyContainer = document.getElementById("body-parameters-container");
  if (!queryContainer || !bodyContainer) return;

  // endpoint field (Ad-Hoc style)
  queryContainer.innerHTML = "";
  const qh = document.createElement("div");
  qh.className = "parameter-header";
  qh.textContent = "Endpoint URL with Query Parameters";
  queryContainer.appendChild(qh);

  const ep = document.createElement("input");
  ep.type = "text";
  ep.id = "adhoc-endpoint";
  ep.classList.add("query-param-input");
  ep.placeholder = "/v1/endpoint?queryParam=value";
  ep.value = item?.endpoint || "";
  queryContainer.appendChild(ep);

  // body field for post methods only
  bodyContainer.innerHTML = "";
  if (String(item?.method).toUpperCase() === "POST") {
    const bh = document.createElement("div");
    bh.className = "parameter-header";
    bh.textContent = "Full JSON Body";
    bodyContainer.appendChild(bh);

    const ta = document.createElement("textarea");
    ta.id = "adhoc-body";
    ta.className = "json-textarea";
    ta.placeholder = "Enter full JSON body here...";
    ta.value = item?.body || "";
    bodyContainer.appendChild(ta);
    ensureBodyFontControls(bh);
  }
}

// Save request button
async function onSaveRequestClick() {
  try {
    const sel = document.getElementById("api-selector");
    const selectedKey = sel?.value || "";

    let method = "GET";
    let endpoint = "";
    let bodyStr = "";

    if (HERMES_MYAPIS_MODE && selectedKey.startsWith("myapi:")) {
      // editing an existing saved item → read current inputs
      const id = selectedKey.slice(6);
      const list = await getSavedMyApis();
      const item = list.find((x) => x.id === id);
      if (!item) return;

      method = item.method.toUpperCase();
      endpoint = (
        document.getElementById("adhoc-endpoint")?.value || ""
      ).trim();
      if (method === "POST") {
        bodyStr = (document.getElementById("adhoc-body")?.value || "").trim();
      }

      const name = prompt(
        "Rename and save: (120 Character Limit)",
        item.name
      )?.trim();
      if (!name) return;
      const MAX_NAME = 120;
      if (name.length > MAX_NAME) {
        alert(
          `Name is too long (${name.length}). Please keep it under ${MAX_NAME} characters.`
        );
        return;
      }
      item.name = name;
      item.method = method;
      item.endpoint = endpoint;
      item.body = bodyStr;
      item.updatedAt = new Date().toISOString();
      await setSavedMyApis(list);

      setButtonTempText?.(document.getElementById("save-request"), "Saved!");
      return;
    }

    // ad-hoc get/post → build from fields
    if (selectedKey === "adHocGet" || selectedKey === "adHocPost") {
      method = selectedKey === "adHocPost" ? "POST" : "GET";
      endpoint = (
        document.getElementById("adhoc-endpoint")?.value || ""
      ).trim();
      if (!endpoint) {
        alert("Please enter an endpoint to save.");
        return;
      }
      if (method === "POST") {
        bodyStr = (document.getElementById("adhoc-body")?.value || "").trim();
      }
    } else {
      // public library item → build url with query params and body from your current ui
      const apiLib = await loadApiLibrary();
      const api = apiLib[selectedKey];
      if (!api) {
        alert("Selected API not found.");
        return;
      }

      method = (api.method || "GET").toUpperCase();
      endpoint = api.url;
      // Replace path params
      try {
        endpoint = buildUrlWithPathParams(api.url, api);
      } catch (e) {
        alert(e.message || "Missing path parameters.");
        return;
      }

      if (method === "GET") {
        const params = new URLSearchParams();
        document
          .querySelectorAll("#query-parameters-container .query-param-input")
          .forEach((inp) => {
            const v = (inp.value || "").trim();
            if (!v) return;
            const k = (inp.id || "").replace(/^query-/, "") || inp.name || "";
            if (k) params.append(k, v);
          });
        const qs = params.toString();
        if (qs) endpoint += "?" + qs;
      } else if (method === "POST" && api.requestProfile) {
        const tmpl = JSON.parse(JSON.stringify(api.requestProfile));
        const bodyHost = document.getElementById("body-parameters-container");
        const inputs = Array.from(bodyHost.querySelectorAll("[data-path]"));
        mapUserInputsToRequestProfile(tmpl, inputs);
        bodyStr = JSON.stringify(tmpl, null, 2);
      }
    }

    const defaultName = `${method} ${endpoint.split("?")[0] || ""}`;
    const name = prompt(
      "Save request as: (120 Character Limit)",
      defaultName
    )?.trim();
    if (!name) return;

    const MAX_NAME = 120;
    if (name.length > MAX_NAME) {
      alert(
        `Name is too long (${name.length}). Please keep it under ${MAX_NAME} characters.`
      );
      return;
    }

    const list = await getSavedMyApis();
    const entry = {
      id: genId(),
      name,
      method,
      endpoint,
      body: bodyStr,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
    list.push(entry);
    await setSavedMyApis(list);

    // switch to My APIs mode, repopulate, select the new one, and render
    await populateApiDropdownMyApis();
    const newKey = `myapi:${entry.id}`;
    const dd = document.getElementById("api-selector");
    dd.value = newKey;
    await handleApiSelection(newKey);
    applyDynamicStyles?.();

    setButtonTempText?.(document.getElementById("save-request"), "Saved!");
  } catch (e) {
    console.error(e);
    setButtonFailText?.(document.getElementById("save-request"), "Save Failed");
  }
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
  if (typeof clearParameters === "function") {
    clearParameters();
  } else {
    const q = document.getElementById("query-parameters-container");
    const b = document.getElementById("body-parameters-container");
    const p = document.getElementById("path-parameters-container");
    if (q) q.innerHTML = "";
    if (b) b.innerHTML = "";
    if (p) p.innerHTML = "";
  }

  // (optional) clear the last response panel
  /*if (typeof clearApiResponse === "function") {
    clearApiResponse();
  }*/

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
  let animation;

  try {
    clearApiResponse();

    // start loading animation and store both interval and original text
    animation = startLoadingAnimation(button);

    if (!(await isValidSession())) {
      alert("Requires a valid ADP Workforce Manager session.");
      throw new Error("Invalid session");
    }

    const apiDropdown = document.getElementById("api-selector");
    const selectedApiKey = apiDropdown?.value;
    if (!selectedApiKey || selectedApiKey === "Select API...") {
      alert("Please select an API to execute.");
      throw new Error("No API selected");
    }

    // detect saved my api entry (myapi:<id>)
    let savedEntry = null;
    if (selectedApiKey.startsWith("myapi:")) {
      const id = selectedApiKey.slice(6);
      const { hermes_myapis } = await new Promise((resolve) =>
        chrome.storage.local.get(["hermes_myapis"], resolve)
      );
      const list = hermes_myapis || [];
      savedEntry = list.find((x) => x.id === id) || null;
      if (!savedEntry) {
        alert("Saved request not found.");
        throw new Error("Saved request not found");
      }
    }

    const clienturl = await getClientUrl();
    let data = await loadClientData();
    let clientData = data[clienturl] || {};

    if (
      !clientData.accesstoken ||
      new Date(clientData.expirationdatetime) < new Date()
    ) {
      await fetchToken();
      clientData = await waitForUpdatedToken(clienturl);
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

    // shared variables
    let fullUrl = "";
    let requestBody = null;
    let requestMethod = "GET";

    // branch A: saved My API (ad-hoc style)
    if (savedEntry) {
      requestMethod = String(savedEntry.method || "GET").toUpperCase();

      // prefer current ui values (if user modified), fallback to saved values
      const epUi = document.getElementById("adhoc-endpoint")?.value?.trim();
      const bodyUi = document.getElementById("adhoc-body")?.value;

      const endpoint = (epUi || savedEntry.endpoint || "").trim();
      if (!endpoint) {
        alert("Please provide an endpoint URL.");
        throw new Error("Empty saved endpoint");
      }
      fullUrl = clientData.apiurl + endpoint;

      if (requestMethod === "POST") {
        const raw = (
          typeof bodyUi === "string" ? bodyUi : savedEntry.body || ""
        ).trim();
        if (!raw) {
          alert("Please provide a JSON body.");
          throw new Error("Empty JSON body (saved)");
        }
        try {
          requestBody = JSON.parse(raw);
          pruneRequestBody(requestBody);
        } catch {
          alert("Invalid JSON body. Please correct it.");
          throw new Error("Invalid JSON format (saved)");
        }
      }

      // Save request details for UI (redacted)
      lastRequestDetails = {
        method: requestMethod,
        url: fullUrl,
        headers: redactedHeaders,
        body: requestBody ? JSON.stringify(requestBody, null, 2) : null,
      };

      updateRequestDependentButtons(true);

      const response = await fetch(fullUrl, {
        method: requestMethod,
        headers: requestHeaders, // REAL token used here
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

      if (animation?.interval) clearInterval(animation.interval);
      if (response.ok)
        setButtonTempText(button, "Success!", 2000, animation.originalText);
      else setButtonFailText(button, "Failed!", 2000, animation.originalText);

      return; // done with saved branch
    }

    // branch B: public library path
    const apiLibrary = await loadApiLibrary();
    const selectedApi = apiLibrary[selectedApiKey];
    if (!selectedApi) {
      alert("Selected API not found in the library.");
      throw new Error("API not found");
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
        throw new Error("Empty ad-hoc endpoint");
      }
      fullUrl = clientData.apiurl + endpointInput.value.trim();
    }

    // handle pre-request logic if needed
    if (selectedApi.preRequest) {
      const preRequestApi = apiLibrary[selectedApi.preRequest.apiKey];
      const preRequestUrl = clientData.apiurl + preRequestApi.url;

      const preResponse = await fetch(preRequestUrl, {
        method: "GET",
        headers: { Authorization: `Bearer ${accessToken}` }, // real token required
      });

      if (!preResponse.ok) {
        const errorText = await preResponse.text();
        displayApiResponse({ error: errorText }, selectedApiKey);
        throw new Error(
          `Pre-request failed. HTTP status: ${preResponse.status}`
        );
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

      // dynamically insert mapped values into requestBody using the correct datapath
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
            throw new Error("Empty JSON body");
          }
          try {
            requestBody = JSON.parse(bodyInput.value.trim());
          } catch {
            alert("Invalid JSON body. Please correct it.");
            throw new Error("Invalid JSON format");
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
      headers: requestHeaders, // REAL token used here
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

    if (animation?.interval) clearInterval(animation.interval);
    if (response.ok)
      setButtonTempText(button, "Success!", 2000, animation.originalText);
    else setButtonFailText(button, "Failed!", 2000, animation.originalText);
  } catch (error) {
    console.error("Error executing API call:", error);
    alert(`API call failed: ${error.message}`);
    displayApiResponse({ error: error.message }, "Error");

    if (animation?.interval) clearInterval(animation.interval);
    setButtonFailText(button, "Failed!", 2000, animation.originalText);
  }
}

// Display API response (default = raw view) V1
async function displayApiResponse(response, apiKey) {
  const responseSection = document.getElementById("response-section");
  window.lastApiResponseObject = response; // stash for popout/toggle

  // preserve/create popout button
  let popoutButton = document.getElementById("popout-response");
  if (!popoutButton) {
    popoutButton = document.createElement("button");
    popoutButton.id = "popout-response";
    popoutButton.className = "btn3";
    popoutButton.innerHTML = `
      Popout Response 
      <img src="icons/external-link.png" alt="Popout" class="btn-icon">
    `;
    popoutButton.addEventListener("click", popoutResponse);
    responseSection.prepend(popoutButton);
  }

  // load api library and set up export csv button if needed
  const apiLibrary = await loadApiLibrary();
  const selectedApi = apiLibrary[apiKey];

  if (!selectedApi) {
    console.warn("API Key Not Found in Library:", apiKey);
  }

  if (selectedApi?.exportMap) {
    let exportCsvButton = document.getElementById("export-api-csv");
    if (!exportCsvButton) {
      exportCsvButton = document.createElement("button");
      exportCsvButton.id = "export-api-csv";
      exportCsvButton.className = "btn3";
      exportCsvButton.innerHTML = `
        Export CSV
        <img src="icons/export-csv.png" alt="CSV" class="btn-icon">
      `;
      exportCsvButton.addEventListener("click", () => {
        exportApiResponseToCSV(response, selectedApi.exportMap, apiKey);
      });
      responseSection.appendChild(exportCsvButton);
    }
  }

  // ensure tree/raw toggle exists (default = raw)
  ensureViewToggle();

  // clear prior render
  [...responseSection.querySelectorAll(".json-tree, pre")].forEach((n) =>
    n.remove()
  );

  // raw by default
  const pre = document.createElement("pre");
  pre.textContent = JSON.stringify(response, null, 2);
  responseSection.appendChild(pre);

  // set toggle button state to raw
  const toggle = document.getElementById("toggle-view");
  if (toggle) {
    toggle.dataset.mode = "raw";
    toggle.textContent = "Tree View";
  }

  // enable buttons
  updateResponseDependentButtons(true);
  const downloadButton = document.getElementById("download-response");
  if (downloadButton) {
    downloadButton.disabled = false;
    downloadButton.onclick = () => downloadApiResponse(response, apiKey);
  }
}

// JSON tree view renderer
function renderJsonTree(data, rootEl, { collapsedDepth = 1 } = {}) {
  rootEl.innerHTML = "";
  const el = buildNode(data, undefined, 0, collapsedDepth);
  rootEl.appendChild(el);
}

// JSON tree node count helper
function containerBadge(value) {
  if (Array.isArray(value)) return `[${value.length}]`;
  return `{${Object.keys(value).length}}`;
}

// JSON tree view node builder (with counts)
function buildNode(value, key, depth, collapsedDepth) {
  const isObjLike = (v) => v && typeof v === "object" && v !== null;

  if (isObjLike(value)) {
    const details = document.createElement("details");
    details.open = depth < collapsedDepth;

    const summary = document.createElement("summary");
    const badge = containerBadge(value);

    // Notepad++-style labels: "key {2}"  /  "values [0]"  /  "{3}" or "[1]" for root nodes
    summary.textContent = key != null ? `${key} ${badge}` : badge;
    details.appendChild(summary);

    if (Array.isArray(value)) {
      for (let i = 0; i < value.length; i++) {
        details.appendChild(buildNode(value[i], i, depth + 1, collapsedDepth));
      }
    } else {
      const keys = Object.keys(value);
      for (const k of keys) {
        details.appendChild(buildNode(value[k], k, depth + 1, collapsedDepth));
      }
    }
    return details;
  } else {
    // leaf
    const row = document.createElement("div");
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
  const ids = [
    "view-request-details",
    "save-request-definition",
    "export-bruno-request",
    "save-request",
  ];

  ids.forEach((id) => {
    const btn = document.getElementById(id);
    if (btn) {
      btn.disabled = !enabled;
    }
  });
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

// Save request definition as collection (Postman v2.1+/Bruno)
async function saveRequestDefinition() {
  try {
    if (!lastRequestDetails) {
      alert("No request details available. Send a request first.");
      return;
    }

    const { method, url, headers = {}, body } = lastRequestDetails;

    // ----- 1) build a sensible default name -----
    const apiSel = document.getElementById("api-selector");
    const selectedKey = apiSel?.value || "";
    let defaultName = "";

    // parse the URL early so we can reuse it
    let urlObj;
    try {
      urlObj = new URL(url);
    } catch {
      urlObj = { href: url, pathname: url, origin: "" };
    }
    const pathName = urlObj.pathname || "/";

    try {
      const apiLibrary = await loadApiLibrary();

      // public API: use library entry's "name"
      if (selectedKey && apiLibrary[selectedKey]) {
        defaultName = apiLibrary[selectedKey].name || "";
      }
      // My API: use saved My API name
      else if (selectedKey.startsWith("myapi:")) {
        const myApis = await getSavedMyApis();
        const id = selectedKey.slice(6);
        const item = myApis.find((x) => x.id === id);
        if (item?.name) defaultName = item.name;
      }
    } catch {
      // silent fallback to method + path
    }

    // fallback default name using method + path
    if (!defaultName) {
      defaultName = `${method} ${pathName}`;
    }

    // ----- 2) ask user for a friendly name -----
    const MAX_NAME = 120;
    const requestDisplayName =
      prompt("Save request definition as: (120 Character Limit)", defaultName)?.trim() ||
      "";

    if (!requestDisplayName) return; // user cancelled or left blank

    if (requestDisplayName.length > MAX_NAME) {
      alert(
        `Name is too long (${requestDisplayName.length}). Please keep it under ${MAX_NAME} characters.`
      );
      return;
    }

    // ----- 3) build headers: keep all, but mask Authorization -----
    const headerArray = Object.entries(headers).map(([key, value]) => {
      let v = String(value);
      if (key.toLowerCase() === "authorization") {
        v = `Bearer {{${EXPORT_ACCESS_TOKEN_VAR}}}`;
      }
      return { key, value: v };
    });

    // ----- 4) transform URL to use {{apiUrl}} where possible -----
    let exportedUrl = urlObj.href || url;
    try {
      const full = urlObj.href || url;
      const marker = "/api/";
      const idx = full.toLowerCase().indexOf(marker);
      if (idx !== -1) {
        const afterApi = full.substring(idx + marker.length - 1); // keep leading '/'
        exportedUrl = `{{${EXPORT_API_URL_VAR}}}` + afterApi;
      }
    } catch {
      // silent: keep full URL
    }

    // ----- 5) create a safe filename from the chosen name -----
    const rawSlug = requestDisplayName.replace(/\s+/g, "-");
    const safeSlug = rawSlug
      .replace(/[^a-zA-Z0-9_-]+/g, "-")
      .replace(/^-+|-+$/g, "")
      .toLowerCase();

    const fileName =
      (safeSlug || "api-accesspanel-request") + ".postman_collection.json";

    // ----- 6) build Postman v2.1 collection with 1 request -----
    const collection = {
      info: {
        name: `AccessPanel – ${requestDisplayName}`,
        schema:
          "https://schema.getpostman.com/json/collection/v2.1.0/collection.json",
      },
      item: [
        {
          name: requestDisplayName,
          request: {
            method,
            header: headerArray,
            url: exportedUrl,
            auth: { type: "noauth" },
            ...(body
              ? {
                  body: {
                    mode: "raw",
                    raw: body,
                    options: { raw: { language: "json" } },
                  },
                }
              : {}),
          },
        },
      ],
    };

    const jsonText = JSON.stringify(collection, null, 2);

    // ----- 7) save file -----
    if (window.showSaveFilePicker) {
      try {
        const handle = await window.showSaveFilePicker({
          suggestedName: fileName,
          types: [
            {
              description: "Postman Collection",
              accept: { "application/json": [".json"] },
            },
          ],
        });

        const writable = await handle.createWritable();
        await writable.write(jsonText);
        await writable.close();
        return;
      } catch (e) {
        if (e && e.name === "AbortError") {
          return; // user cancelled – silent exit
        }
        // Non-fatal: fall back to download helper
      }
    }

    // Fallback for browsers without File System Access API (or if picker failed)
    downloadFile(fileName, jsonText, "application/json");
  } catch (err) {
    console.error("Failed to save request definition:", err);
    alert("Failed to save request definition.");
  }
}

// Export request (Bruno .bru) ===== //
async function saveBrunoRequest() {
  try {
    if (!lastRequestDetails) {
      alert("No request details available. Send a request first.");
      return;
    }

    const { method, url, headers = {}, body } = lastRequestDetails;

    // ----- 1) build a sensible default name ----- //
    const apiSel = document.getElementById("api-selector");
    const selectedKey = apiSel?.value || "";
    let defaultName = "";

    // parse URL early so we can reuse it
    let urlObj;
    try {
      urlObj = new URL(url);
    } catch {
      urlObj = { href: url, pathname: url, origin: "" };
    }
    const pathName = urlObj.pathname || "/";

    try {
      const apiLibrary = await loadApiLibrary();

      // public API: use library entry's "name"
      if (selectedKey && apiLibrary[selectedKey]) {
        defaultName = apiLibrary[selectedKey].name || "";
      }
      // My API: use saved My API name
      else if (selectedKey.startsWith("myapi:")) {
        const myApis = await getSavedMyApis();
        const id = selectedKey.slice(6);
        const item = myApis.find((x) => x.id === id);
        if (item?.name) defaultName = item.name;
      }
    } catch {
      // silent: fallback to method + path
    }

    if (!defaultName) {
      defaultName = `${method} ${pathName}`;
    }

    const MAX_NAME = 120;
    const requestDisplayName =
      prompt("Save Bruno request as: (120 Character Limit)", defaultName)?.trim() ||
      "";

    if (!requestDisplayName) return; // user cancelled or empty
    if (requestDisplayName.length > MAX_NAME) {
      alert(
        `Name is too long (${requestDisplayName.length}). Please keep it under ${MAX_NAME} characters.`
      );
      return;
    }

    // ----- 2) transform URL to use {{apiUrl}} if possible ----- //
    let exportedUrl = urlObj.href || url;
    try {
      const full = urlObj.href || url;
      const marker = "/api/";
      const idx = full.toLowerCase().indexOf(marker);
      if (idx !== -1) {
        const afterApi = full.substring(idx + marker.length - 1); // keep leading '/'
        exportedUrl = `{{${EXPORT_API_URL_VAR}}}` + afterApi;
      }
    } catch {
      // silent: keep full URL
    }

    // ----- 3) build headers block (Authorization masked) ----- //
    const headerLines = Object.entries(headers)
      .map(([key, value]) => {
        let v = String(value);
        if (key.toLowerCase() === "authorization") {
          v = `Bearer {{${EXPORT_ACCESS_TOKEN_VAR}}}`;
        }
        return `  ${key}: ${v}`;
      })
      .join("\n");

    // ----- 4) prepare body + method block ----- //
    const verb = (method || "GET").toLowerCase();
    const hasBody = typeof body === "string" && body.trim() !== "";

    let prettyBody = "";
    if (hasBody) {
      try {
        const parsed = JSON.parse(body);
        prettyBody = JSON.stringify(parsed, null, 2);
      } catch {
        prettyBody = body;
      }
    }

    const indentedBody = hasBody
      ? prettyBody
          .split("\n")
          .map((line) => "  " + line)
          .join("\n")
      : "";

    // ----- 5) build .bru file content ----- //
    let bru = "";

    bru += "meta {\n";
    bru += `  name: ${requestDisplayName}\n`;
    bru += "  type: http\n";
    bru += "  seq: 1\n";
    bru += "}\n\n";

    bru += `${verb} {\n`;
    bru += `  url: ${exportedUrl}\n`;
    bru += `  body: ${hasBody ? "json" : "none"}\n`;
    bru += "  auth: none\n";
    bru += "}\n\n";

    bru += "headers {\n";
    if (headerLines) bru += headerLines + "\n";
    bru += "}\n\n";

    if (hasBody) {
      bru += "body:json {\n";
      bru += `${indentedBody}\n`;
      bru += "}\n\n";
    }

    bru += "settings {\n";
    bru += "  encodeUrl: true\n";
    bru += "}\n";

    // ----- 6) build filename from name ----- //
    const rawSlug = requestDisplayName.replace(/\s+/g, "-");
    const safeSlug = rawSlug
      .replace(/[^a-zA-Z0-9_-]+/g, "-")
      .replace(/^-+|-+$/g, "")
      .toLowerCase();

    const fileName = (safeSlug || "api-accesspanel-request") + ".bru";

    // ----- 7) save .bru file ----- //
    if (window.showSaveFilePicker) {
      try {
        const handle = await window.showSaveFilePicker({
          suggestedName: fileName,
          types: [
            {
              description: "Bruno Request (.bru)",
              accept: { "text/plain": [".bru"] },
            },
          ],
        });

        const writable = await handle.createWritable();
        await writable.write(bru);
        await writable.close();
        return;
      } catch (e) {
        if (e && e.name === "AbortError") {
          // user canceled – silent exit
          return;
        }
        // Non-fatal: fall back to download helper
      }
    }

    // Fallback (or no File System Access API support)
    downloadFile(fileName, bru, "text/plain");
  } catch (err) {
    console.error("Failed to export Bruno request:", err);
    alert("Failed to export Bruno request.");
  }
}

// ===== Export environment (Postman/Bruno) ===== //
async function exportEnvironmentDefinition() {
  const btn = document.getElementById("export-env");

  // Use the same logic as populateClientUrlField / the UI, not lastRequestDetails
  let base = null;
  try {
    base = await getClientUrl();
  } catch (e) {
    console.error("exportEnvironmentDefinition: getClientUrl failed:", e);
  }

  if (!base) {
    alert(
      "Unable to determine API URL for the current tenant. Make sure you are on a valid ADP WFM page."
    );
    return;
  }

  const apiUrl = toApiUrl(base); // e.g. https://<tenant>.mykronos.com/api

  // Parse apiUrl to get hostname and build ID / defaults
  let urlObj;
  try {
    urlObj = new URL(apiUrl);
  } catch {
    alert("Could not determine environment host from the API URL.");
    return;
  }

  // Prompt for environment name (default to hostname)
  const defaultEnvName = urlObj.hostname || "API Environment";
  const nameInput = window.prompt("Enter a name for this environment:", defaultEnvName);
  if (nameInput === null) return; // user cancelled
  const envName = nameInput.trim() || defaultEnvName;

  // Build environment ID from hostname + timestamp,
  // then replace '.', '/', ':' with '-'
  const marker = ".mykronos.com";
  let hostBase = urlObj.hostname || "environment";
  if (hostBase.includes(marker)) {
    hostBase = hostBase.split(marker)[0];
  }

  const nowIso = new Date().toISOString(); // 2025-11-26T14:28:22.923Z
  const rawId = `${hostBase}${nowIso}`;
  const envId = rawId.replace(/[./:]/g, "-");

  // Variable names from Preferences (fall back to defaults if not set)
  const apiUrlVarName = EXPORT_API_URL_VAR || "apiUrl";
  const accessTokenVarName = EXPORT_ACCESS_TOKEN_VAR || "accessToken";

  // Load panel version from accesspanel.json for _postman_exported_using
  let exportedUsing = "AccessPanel/unknown";
  try {
    const panelMeta = await fetch("accesspanel.json").then((res) => res.json());
    const version = panelMeta?.details?.version || "unknown";
    exportedUsing = `AccessPanel/${version}`;
  } catch {
    // silent fallback to AccessPanel/unknown
  }

  const environment = {
    id: envId,
    name: envName,
    values: [
      {
        key: apiUrlVarName,
        value: apiUrl,
        type: "default",
        enabled: true,
      },
      {
        key: accessTokenVarName,
        value: "", // never export the real token
        type: "default",
        enabled: true,
      },
    ],
    color: null,
    _postman_variable_scope: "global",
    _postman_exported_at: nowIso,
    _postman_exported_using: exportedUsing,
  };

  const safeName = envName.replace(/[^a-z0-9_\- ]/gi, "_");
  const fileName = `${safeName || "AccessPanel_Env"}.postman_environment.json`;
  const content = JSON.stringify(environment, null, 2);

  try {
    if (window.showSaveFilePicker) {
      const handle = await window.showSaveFilePicker({
        suggestedName: fileName,
        types: [
          {
            description: "JSON Files",
            accept: { "application/json": [".json"] },
          },
        ],
      });

      const writable = await handle.createWritable();
      await writable.write(content);
      await writable.close();
    } else {
      // Fallback for browsers that don't support showSaveFilePicker
      downloadFile(fileName, content, "application/json");
    }

    if (btn) {
      const originalTitle = btn.title;
      btn.title = "Environment exported!";
      setTimeout(() => {
        btn.title = originalTitle;
      }, 2000);
    }
  } catch (error) {
    if (error && error.name === "AbortError") {
      // user cancelled the save dialog – silent exit
      return;
    }
    console.error("Failed to export environment:", error);
    if (btn) {
      const originalTitle = btn.title;
      btn.title = "Export failed";
      setTimeout(() => {
        btn.title = originalTitle;
      }, 2000);
    }
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

// Helper for RAW Popout To Avoid Breaking HTML
function escapeHtml(s) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// Request details button
function showRequestDetails() {
  // pull theme colors from the current document root
  const rootStyles = getComputedStyle(document.documentElement);
  const primary =
    rootStyles.getPropertyValue("--primary-color").trim() || "#f5f5f5";
  const secondary =
    rootStyles.getPropertyValue("--secondary-color").trim() || "#0059B3";
  const accent =
    rootStyles.getPropertyValue("--accent-color").trim() || "#00AEEF";
  const highlight =
    rootStyles.getPropertyValue("--highlight-color").trim() || "#007ACC";
  const textOnBtn =
    rootStyles.getPropertyValue("--buttontext-color").trim() || "#FFFFFF";

  if (!lastRequestDetails) {
    const noDetailsHtml = `
      <html>
        <head>
          <title>No Request Details</title>
          <style>
            body {
              font-family: Arial, sans-serif;
              padding: 20px;
              margin: 0;
              background-color: ${primary};
              color: ${accent};
              line-height: 1.5;
            }
            h1 {
              color: ${accent};
              margin-bottom: 0.5rem;
              font-size: 1.3rem;
              font-weight: bold;
            }
            p {
              margin: 0;
              font-size: 0.8rem;
            }
          </style>
        </head>
        <body>
          <h1>No request details available.</h1>
          <p>Send a request first to view request details.</p>
        </body>
      </html>
    `;

    const noDetailsWindow = window.open(
      "",
      "_blank",
      "width=400,height=300,scrollbars=yes,resizable=yes"
    );
    if (!noDetailsWindow) return;
    noDetailsWindow.document.write(noDetailsHtml);
    noDetailsWindow.document.close();
    return;
  }

  const safeMethod = escapeHtml(String(lastRequestDetails.method || ""));
  const safeUrl = escapeHtml(String(lastRequestDetails.url || ""));
  const safeHeaders = escapeHtml(
    JSON.stringify(lastRequestDetails.headers || {}, null, 2)
  );
  const safeBody = escapeHtml(
    lastRequestDetails.body != null && lastRequestDetails.body !== ""
      ? String(lastRequestDetails.body)
      : "No body"
  );

  const requestDetailsHtml = `
    <html>
      <head>
        <title>Request Details</title>
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 16px;
            margin: 0;
            line-height: 1.6;
            background-color: ${primary};
            color: ${accent};
          }
          h1 {
            color: ${accent};
            font-size: 1.4rem;
            font-weight: bold;
            margin: 0 0 0.75rem 0;
          }
          h2 {
            color: ${accent};
            font-size: .9rem;
            margin: 1rem 0 0.35rem 0;
          }
          .meta-line {
            margin: 0.15rem 0;
            font-size: 0.8rem;
          }
          .meta-line strong {
            font-weight: bold;
          }
          pre {
            background: ${textOnBtn};
            border: 2px solid ${accent};
            border-radius: 6px;
            padding: 10px;
            overflow-x: auto;
            white-space: pre;
            font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace;
            font-size: 0.7rem;
            color: ${accent};
          }
        </style>
      </head>
      <body>
        <h1>Request Details</h1>
        <p class="meta-line"><strong>Method:</strong> ${safeMethod}</p>
        <p class="meta-line"><strong>URL:</strong> ${safeUrl}</p>

        <h2>Headers</h2>
        <pre>${safeHeaders}</pre>

        <h2>Body</h2>
        <pre>${safeBody}</pre>
      </body>
    </html>
  `;

  const detailsWindow = window.open(
    "",
    "_blank",
    "width=1000,height=600,scrollbars=yes,resizable=yes"
  );
  if (!detailsWindow) return;
  detailsWindow.document.write(requestDetailsHtml);
  detailsWindow.document.close();
}

// Popout response button (API Response)
function popoutResponse() {
  const data = window.lastApiResponseObject;

  // pull theme colors from the current document root
  const rootStyles = getComputedStyle(document.documentElement);
  const primary =
    rootStyles.getPropertyValue("--primary-color").trim() || "#f5f5f5";
  const secondary =
    rootStyles.getPropertyValue("--secondary-color").trim() || "#0059B3";
  const accent =
    rootStyles.getPropertyValue("--accent-color").trim() || "#00AEEF";
  const highlight =
    rootStyles.getPropertyValue("--highlight-color").trim() || "#007ACC";
  const textOnBtn =
    rootStyles.getPropertyValue("--buttontext-color").trim() || "#FFFFFF";

  if (!data) {
    // themed "no response" popup
    const noResponseHtml = `
      <html>
        <head>
          <title>No Response</title>
          <style>
            body {
              font-family: Arial, sans-serif;
              padding: 20px;
              margin: 0;
              background-color: ${primary};
              color: ${accent};
              line-height: 1.5;
            }
            h1 {
              color: ${accent};
              margin-bottom: 0.5rem;
              font-size: 1.3rem;
              font-weight: bold;
            }
            p {
              margin: 0;
              font-size: 0.95rem;
            }
          </style>
        </head>
        <body>
          <h1>No response available.</h1>
          <p>Please send an API request to generate a response.</p>
        </body>
      </html>`;
    const w = window.open(
      "",
      "_blank",
      "width=400,height=300,scrollbars=yes,resizable=yes"
    );
    if (!w) return;
    w.document.write(noResponseHtml);
    w.document.close();
    return;
  }

  const mode = document.getElementById("toggle-view")?.dataset.mode || "raw";

  if (mode === "raw") {
    // RAW popout – themed like BIRT popup
    const responseHtml = `
      <html>
        <head>
          <title>API Response</title>
          <style>
            body {
              font-family: Arial, sans-serif;
              padding: 16px;
              margin: 0;
              line-height: 1.6;
              background-color: ${primary};
              color: ${accent};
            }
            h1 {
              color: ${accent};
              font-size: 1.4rem;
              font-weight: bold;
              margin: 0 0 0.75rem 0;
            }
            pre {
              background: ${textOnBtn};
              border: 2px solid ${accent};
              border-radius: 6px;
              padding: 10px;
              overflow-x: auto;
              white-space: pre;
              font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace;
              font-size: 0.7rem;
              color: ${accent};
            }
          </style>
        </head>
        <body>
          <h1>API Response</h1>
          <pre>${escapeHtml(JSON.stringify(data, null, 2))}</pre>
        </body>
      </html>`;
    const w = window.open(
      "",
      "_blank",
      "width=800,height=900,scrollbars=yes,resizable=yes"
    );
    if (!w) return;
    w.document.write(responseHtml);
    w.document.close();
  } else {
    // TREE popout (no inline script: pre-render HTML in the parent)
    const container = document.createElement("div");
    container.className = "json-tree";

    // reuse existing renderer to build DOM in this temp container
    renderJsonTree(data, container, { collapsedDepth: 1 });

    // serialize the built tree to static HTML
    const treeHtml = container.outerHTML;

    const responseHtml = `
      <html>
        <head>
          <title>API Response (Tree)</title>
          <style>
            body {
              font-family: Arial, sans-serif;
              padding: 16px;
              margin: 0;
              line-height: 1.6;
              background-color: ${primary};
              color: ${accent};
            }
            h1 {
              color: ${accent};
              margin: 0 0 0.5rem 0;
              font-size: 1.4rem;
              font-weight: bold;
            }
            .json-tree {
              font: 12px/1.4 ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
              background: ${textOnBtn};
              border: 2px solid ${accent};
              border-radius: 6px;
              padding: 10px;
              color: ${accent};
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
          </style>
        </head>
        <body>
          <h1>API Response (Tree)</h1>
          ${treeHtml}
        </body>
      </html>`;
    const w = window.open(
      "",
      "_blank",
      "width=900,height=1000,scrollbars=yes,resizable=yes"
    );
    if (!w) return;
    w.document.write(responseHtml);
    w.document.close();
  }
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


// ===== MANAGE MY API'S OVERLAY FUNCTIONS ===== //
let MYAPIS_MANAGER_WIRED = false;

// Date time stamp
function formatStamp(iso) {
  if (!iso) return "";
  try {
    return new Date(iso).toLocaleString();
  } catch {
    return iso;
  }
}

// Dedupe My API
function dedupeById(list) {
  const seen = new Set();
  const out = [];
  for (const x of list || []) {
    if (!x?.id) continue;
    if (seen.has(x.id)) continue;
    seen.add(x.id);
    out.push(x);
  }
  return out;
}

// Manage My API's button
function toggleManagerButtons() {
  const listEl = document.getElementById("myapis-list");
  const anyChecked = !!listEl?.querySelector(
    'input[type="checkbox"].row-check:checked'
  );
  const delSelBtn = document.getElementById("myapis-delete-selected");
  if (delSelBtn) delSelBtn.disabled = !anyChecked;
}

// Selectable options for Manage My API's with dynamic event listeners
function wireMyApisManagerOnce() {
  // debug: if the DOM somehow has duplicate controls, warn once.
  const dbgIds = [
    "myapis-list",
    "myapis-select-all",
    "myapis-delete-selected",
    "myapis-delete-all",
    "myapis-cancel",
  ];
  const dups = dbgIds
    .map((id) => [id, document.querySelectorAll(`#${CSS.escape(id)}`).length])
    .filter(([_, n]) => n > 1);
  if (dups.length) {
    console.warn("My APIs manager: duplicate controls in DOM:", dups);
  }

  const listHost = document.getElementById("myapis-list");
  const selAll = document.getElementById("myapis-select-all");
  const delSelBtn = document.getElementById("myapis-delete-selected");
  const delAllBtn = document.getElementById("myapis-delete-all");
  const cancelBtn = document.getElementById("myapis-cancel");

  // ----- delegated checkbox change on list -----
  if (listHost && !listHost.dataset.wired) {
    listHost.addEventListener("change", (e) => {
      const t = e.target;
      if (!(t instanceof HTMLInputElement)) return;
      if (!t.classList.contains("row-check")) return;
      toggleManagerButtons();
    });
    listHost.dataset.wired = "1";
  }

  // ----- select All -----
  if (selAll && !selAll.dataset.wired) {
    selAll.addEventListener("change", () => {
      listHost?.querySelectorAll("input.row-check").forEach((cb) => {
        cb.checked = selAll.checked;
      });
      toggleManagerButtons();
    });
    selAll.dataset.wired = "1";
  }

  // ----- delete selected -----
  if (delSelBtn && !delSelBtn.dataset.wired) {
    delSelBtn.addEventListener("click", onDeleteSelectedClick);
    delSelBtn.dataset.wired = "1";
  }

  // ----- delete all -----
  if (delAllBtn && !delAllBtn.dataset.wired) {
    delAllBtn.addEventListener("click", onDeleteAllClick);
    delAllBtn.dataset.wired = "1";
  }

  // ----- close -----
  if (cancelBtn && !cancelBtn.dataset.wired) {
    cancelBtn.addEventListener("click", () => closeMyApisManager());
    cancelBtn.dataset.wired = "1";
  }

  MYAPIS_MANAGER_WIRED = true;
}

// Render My APIs Manager overlay
async function renderMyApisManager() {
  const list = dedupeById(await getSavedMyApis()); // normalize duplicates
  const host = document.getElementById("myapis-list");
  if (!host) return;
  host.innerHTML = "";

  const selAll = document.getElementById("myapis-select-all");
  const delSelBtn = document.getElementById("myapis-delete-selected");

  if (!list.length) {
    host.innerHTML = `<div class="myapis-row"><div></div><div class="meta"><div class="title">(No saved APIs)</div></div></div>`;
    // ensure toolbar is reset in the empty case
    if (selAll) selAll.checked = false;
    delSelBtn?.setAttribute("disabled", "true");
    return;
  }

  list.forEach((item) => {
    const row = document.createElement("div");
    row.className = "myapis-row";

    const left = document.createElement("div");
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.className = "row-check";
    cb.dataset.id = item.id;
    left.appendChild(cb);

    const meta = document.createElement("div");
    meta.className = "meta";

    const title = document.createElement("div");
    title.className = "title";
    title.textContent = `${item.name} (${String(item.method).toUpperCase()})`;

    const subtitle = document.createElement("div");
    subtitle.className = "subtitle";
    subtitle.textContent = item.endpoint || "";

    const stamp = document.createElement("div");
    stamp.className = "stamp";
    stamp.textContent = `Updated: ${formatStamp(
      item.updatedAt || item.createdAt
    )}`;

    meta.appendChild(title);
    meta.appendChild(subtitle);
    meta.appendChild(stamp);

    row.appendChild(left);
    row.appendChild(meta);
    host.appendChild(row);
  });

  // reset toolbar state after render
  if (selAll) selAll.checked = false;
  toggleManagerButtons();
}

// Open My APIs Manager overlay
async function openMyApisManager() {
  wireMyApisManagerOnce(); // attach exactly once
  await renderMyApisManager();

  const overlay = document.getElementById("manage-myapis-overlay");
  if (overlay) overlay.hidden = false;

  // reset controls each open
  const selAll = document.getElementById("myapis-select-all");
  if (selAll) selAll.checked = false;
  document
    .getElementById("myapis-delete-selected")
    ?.setAttribute("disabled", "true");
}

// Close My APIs Manager overlay
function closeMyApisManager() {
  const overlay = document.getElementById("manage-myapis-overlay");
  if (overlay) overlay.hidden = true;

  const selAll = document.getElementById("myapis-select-all");
  if (selAll) selAll.checked = false;
  document
    .getElementById("myapis-delete-selected")
    ?.setAttribute("disabled", "true");
}

// If current selection was among deleted entries, populate first remaining entry or placeholder
function clearSelectionIfDeleted(deletedIds) {
  const sel = document.getElementById("api-selector");
  if (!sel) return;
  const val = sel.value || "";
  if (!val.startsWith("myapi:")) return;
  const curId = val.slice(6);
  if (!deletedIds.includes(curId)) return;

  // fallback: placeholder or first remaining item if any
  sel.value = "";
  clearParameters?.();
  clearApiResponse?.();
  applyDynamicStyles?.();
}

// Delete selected My API's
async function onDeleteSelectedClick() {
  const ids = Array.from(
    document.querySelectorAll("#myapis-list input.row-check:checked")
  ).map((cb) => cb.dataset.id);
  if (!ids.length) return;

  if (
    !confirm(
      `Delete ${ids.length} selected saved API(s)? This cannot be undone.`
    )
  )
    return;

  const curList = await getSavedMyApis();
  const next = dedupeById(curList).filter((x) => !ids.includes(x.id));
  await setSavedMyApis(next);

  await populateApiDropdownMyApis();
  clearSelectionIfDeleted(ids);
  await renderMyApisManager();
}

// Delete all My API's
async function onDeleteAllClick() {
  const curList = dedupeById(await getSavedMyApis());
  if (!curList.length) return;

  if (
    !confirm(
      `Delete ALL (${curList.length}) saved APIs? This cannot be undone.`
    )
  )
    return;

  await setSavedMyApis([]);
  await populateApiDropdownMyApis();
  clearSelectionIfDeleted(curList.map((x) => x.id));
  await renderMyApisManager();
}
// ============================================= //


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

// HermesLink: Enhanced Tab Management And Session Tracking
/*window.HermesLink = (function () {
  const ENFORCE_ACTIVE_TAB_OVERLAY = true; // set to true to overlay on non-linked tabs always
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
      hint: "Return to linked tab to use AccessPanel features.",
      overlay: "AccessPanel is active in another tab. Click below to return.",
    },
    // TMS-specific states
    TMS_OK: {
      banner: "Active (TMS): ",
      hint: "Using Tenant Management System for the linked WFM tenant.",
      overlay: null,
    },
    TMS_NO_VANITY: {
      banner: "TMS: Vanity URL Required",
      hint: 'Enter or expose the "Vanity Non-SSO URL" for this tenant to use AccessPanel here.',
      overlay:
        'AccessPanel can only run on TMS when the "Vanity Non-SSO URL" is filled in and visible for this tenant.',
    },
    TMS_MISMATCH: {
      banner: "TMS Tenant Mismatch",
      hint: 'The "Vanity Non-SSO URL" in TMS does not match the active HerAccessPanelmes WFM session.',
      overlay:
        "This TMS tenant does not match your active AccessPanel session. Open the matching tenant or relink AccessPanel in WFM.",
    },
    // developer portal – always allow manual use
    DEV_OK: {
      banner: "Developer Portal: ",
      hint: "AccessPanel UI enabled for manual copy/paste from the developer portal.",
      overlay: null,
    },
  };

  // ---- TMS helpers ---- //

  // 1) ss this a TMS URL we should even consider?
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

  // 1b) Is this a developer portal URL where AccessPanel should stay enabled?
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

  const isBoomiUrl = (url) => {
    try {
      const u = new URL(url);
      const host = u.hostname.toLowerCase();
      // matches platform.boomi.com and any subdomain like foo.platform.boomi.com
      return (
        host === "platform.boomi.com" || host.endsWith(".platform.boomi.com")
      );
    } catch {
      return false;
    }
  };

  // 2) normalize any raw vanity string into the same shape as getVanityUrl()
  //    and compare only by protocol + host (ignore path and trailing slash).
  const normalizeVanity = (raw) => {
    if (!raw) return null;
    let value = String(raw).trim();
    if (!value) return null;

    // if user didn’t include the scheme, assume https
    if (!/^https?:\/\//i.test(value)) {
      value = "https://" + value;
    }

    try {
      // reuse existing logic that knows how to flip to -nosso, environment, etc.
      const normalizedUrl = getVanityUrl(value); // e.g. "https://24hf-nosso.prd.mykronos.com/"

      const parsed = new URL(normalizedUrl);
      // compare scheme + hostname only; ignore trailing slash or path
      return `${parsed.protocol}//${parsed.hostname}`.toLowerCase();
    } catch (e) {
      return null;
    }
  };

  // 3) read the TMS "Vanity Non-SSO URL" from the page in the active tab (search all frames)
  const getTmsVanityForTab = async (tabId) => {
    try {
      const results = await chrome.scripting.executeScript({
        target: { tabId, allFrames: true },
        func: () => {
          // recursively search document + all shadow roots for the "Vanity Non-SSO URL" field
          const findVanityInput = (root) => {
            if (!root) return null;

            const isVanityField = (sdfInput) => {
              if (!sdfInput || sdfInput.tagName !== "SDF-INPUT") return false;

              // 1) direct ID match
              if (sdfInput.id === "vanityNonSSOURL") return true;

              // 2) label text match inside its shadow root
              if (!sdfInput.shadowRoot) return false;

              const labelEl = sdfInput.shadowRoot.querySelector(
                ".sdf-form-control-wrapper--label, label[part='label']"
              );
              if (!labelEl || !labelEl.textContent) return false;

              const labelText = labelEl.textContent.trim();
              return labelText.startsWith("Vanity Non-SSO URL");
            };

            // direct root query first
            const direct = root.querySelector
              ? root.querySelector("sdf-input#vanityNonSSOURL")
              : null;
            if (direct && direct.shadowRoot && isVanityField(direct)) {
              const input = direct.shadowRoot.querySelector("input#input");
              if (input) return input;
            }

            // walk all elements, dive into shadow roots
            const walker = document.createTreeWalker(
              root,
              NodeFilter.SHOW_ELEMENT,
              null
            );

            let node = walker.currentNode;
            while (node) {
              if (node.tagName === "SDF-INPUT" && isVanityField(node)) {
                if (node.shadowRoot) {
                  const input = node.shadowRoot.querySelector("input#input");
                  if (input) return input;
                }
              }

              if (node.shadowRoot) {
                const fromShadow = findVanityInput(node.shadowRoot);
                if (fromShadow) return fromShadow;
              }

              node = walker.nextNode();
            }

            return null;
          };

          try {
            const input = findVanityInput(document);
            return input ? input.value.trim() : null;
          } catch (e) {
            console.error("TMS vanity search failed:", e);
            return null;
          }
        },
      });

      // results = [{frameId, result}, ...]; pick first non-null result
      if (Array.isArray(results)) {
        for (const r of results) {
          if (r && r.result) {
            return String(r.result).trim() || null;
          }
        }
      }

      return null;
    } catch (e) {
      return null;
    }
  };

  // 4) core TMS check: does the TMS tenant match the linked WFM session?
  const checkTmsTenantMatchesSession = async (
    currentTab,
    hermesLinkedOrigin
  ) => {
    if (!currentTab?.url || !isTmsUrl(currentTab.url)) {
      return { ok: false, reason: "not_tms" };
    }

    if (!hermesLinkedOrigin) {
      // No linked WFM session – TMS cannot be used yet
      return { ok: false, reason: "no_link" };
    }

    const sessionVanity = normalizeVanity(hermesLinkedOrigin);
    if (!sessionVanity) {
      return { ok: false, reason: "no_link" };
    }

    const tmsVanityRaw = await getTmsVanityForTab(currentTab.id);
    if (!tmsVanityRaw) {
      return { ok: false, reason: "no_tms_vanity" };
    }

    const tmsVanity = normalizeVanity(tmsVanityRaw);
    if (!tmsVanity) {
      return { ok: false, reason: "bad_tms_vanity" };
    }

    const match = tmsVanity === sessionVanity;

    if (!match) {
      return { ok: false, reason: "vanity_mismatch" };
    }

    return { ok: true, reason: "match" };
  };

  // state management
  const state = {
    isInitialized: false,
    checkingState: false,
  };

  // helper functions
  const getActiveTab = async () => {
    try {
      const [tab] = await chrome.tabs.query({
        active: true,
        currentWindow: true,
      });
      return tab || null;
    } catch (e) {
      console.error("Failed to get active tab:", e);
      return null;
    }
  };

  const getLinkedState = async () => {
    try {
      return await chrome.storage.session.get(Object.values(SESSION_KEYS));
    } catch (e) {
      console.error("Failed to get linked state:", e);
      return {};
    }
  };

  const updateLinkedState = async (newState) => {
    try {
      const timestamp = new Date().toISOString();
      await chrome.storage.session.set({
        ...newState,
        hermesLastValidation: timestamp,
      });
    } catch (e) {
      console.error("Failed to update linked state:", e);
    }
  };

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
    const isDevPortal = currentTab ? isDevPortalUrl(currentTab.url) : false;
    const isBoomi = currentTab ? isBoomiUrl(currentTab.url) : false;

    // reuse this for WFM-specific behavior (e.g., relink button)
    const currentTabValidation = currentTab
      ? validateWebPage(currentTab.url)
      : { valid: false };

    // figure out if the current WFM tab is the same tenant as the linked one.
    // if so, we can treat it as effectively "active" even if it's not the exact linked tab.
    let sameTenantAsLinked = false;
    let sessionVanity = null;

    if (hermesLinkedOrigin) {
      sessionVanity = normalizeVanity(hermesLinkedOrigin);
    }

    if (currentTab && currentTabValidation.valid && sessionVanity) {
      try {
        const currentVanity = normalizeVanity(currentTab.url);
        sameTenantAsLinked = !!currentVanity && currentVanity === sessionVanity;
      } catch (e) {
        sameTenantAsLinked = false;
      }
    }

    // if on TMS (and NOT on the linked WFM tab), check the tenant/vanity match
    let tmsCheck = null;
    if (!isLinkedTab && isTms) {
      try {
        tmsCheck = await checkTmsTenantMatchesSession(
          currentTab,
          hermesLinkedOrigin
        );
      } catch (e) {
        tmsCheck = { ok: false, reason: "error" };
      }
    }

    // determine current status
    let currentStatus = "OK";

    if (isDevPortal && !isLinkedTab) {
      // developer portal is always allowed for manual use – no overlay, no disable.
      currentStatus = "DEV_OK";
    } else if (isTms && tmsCheck) {
      if (tmsCheck.ok) {
        // TMS tenant matches the linked WFM session → allow full UI
        currentStatus = "TMS_OK";
      } else if (
        tmsCheck.reason === "no_tms_vanity" ||
        tmsCheck.reason === "bad_tms_vanity"
      ) {
        currentStatus = "TMS_NO_VANITY";
      } else if (tmsCheck.reason === "vanity_mismatch") {
        currentStatus = "TMS_MISMATCH";
      } else if (tmsCheck.reason === "no_link") {
        // no linked WFM session yet – behave like a normal wrong-tab state
        currentStatus = "WRONG_TAB";
      } else {
        currentStatus = "WRONG_TAB";
      }
    } else {
      // normal WFM-driven logic (WFM/mykronos tabs)
      if (!isLinkedTab && sameTenantAsLinked) {
        // different tab, but same tenant as the linked session → treat as active
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

      // relink button only makes sense on valid WFM tabs
      if (relinkButton) {
        if (currentTabValidation.valid) {
          relinkButton.style.display = "inline-block";
        } else {
          relinkButton.style.display = "none";
        }
      }

      // for OK, TMS_OK, DEV_OK, we never show the overlay.
      // if ENFORCE_ACTIVE_TAB_OVERLAY is false, we also avoid overlay for WRONG_TAB
      // so the UI stays usable even when you’re not on the linked tab.
      const isSoftWrongTab =
        currentStatus === "WRONG_TAB" && !ENFORCE_ACTIVE_TAB_OVERLAY;

      const overlayVisible =
        !["OK", "TMS_OK", "DEV_OK"].includes(currentStatus) &&
        !isBoomi &&
        !isSoftWrongTab;

      overlay.classList.toggle("visible", overlayVisible);
    }

    // ----- banner -----
    const banner = document.getElementById("hermes-link-banner");
    if (banner) {
      const status = document.getElementById("hermes-link-status");
      const target = document.getElementById("hermes-link-target");
      const hint = document.getElementById("hermes-link-hint");

      if (status) status.textContent = statusConfig.banner;
      if (target) target.textContent = currentTab?.title || "";
      if (hint) hint.textContent = hermesValidationMessage || statusConfig.hint;
    }

    // tab is considered "active" when:
    //   - we’re on the linked WFM tab (OK),
    //   - we’re on a matching TMS tenant (TMS_OK),
    //   - we’re on an allowed developer portal (DEV_OK).
    // if ENFORCE_ACTIVE_TAB_OVERLAY is false, WRONG_TAB is treated as a soft state:
    // banner + hints still show, but we don’t disable the UI.
    const isSoftWrongTabForUi =
      currentStatus === "WRONG_TAB" && !ENFORCE_ACTIVE_TAB_OVERLAY;

    const shouldDisableUi =
      !["OK", "TMS_OK", "DEV_OK"].includes(currentStatus) &&
      !isBoomi &&
      !isSoftWrongTabForUi;

    document.body.classList.toggle("tab-inactive", shouldDisableUi);
  };

  // session validation
  const validateSession = async () => {
    const { hermesLinkedUrl, hermesLinkedTabId } = await getLinkedState();

    if (!hermesLinkedUrl || !hermesLinkedTabId) {
      return { ok: false, code: "nolink", message: "No linked session found" };
    }

    try {
      const tab = await chrome.tabs.get(hermesLinkedTabId).catch(() => null);
      if (!tab) {
        await updateLinkedState({
          hermesLinkedStatus: "stale",
          hermesValidationMessage: "Linked tab was closed",
        });
        return { ok: false, code: "closed", message: "Linked tab was closed" };
      }

      const validation = validateWebPage(tab.url);
      if (!validation.valid) {
        await updateLinkedState({
          hermesLinkedStatus: "stale",
          hermesValidationMessage: validation.message,
        });
        return { ok: false, code: "invalid", validation };
      }

      // keep HermesLink in sync with whatever WFM URL is
      // actually loaded in that tab (handles “switched tenants
      // in the same tab”).
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

      return { ok: true, code: "ok", validation };
    } catch (e) {
      await updateLinkedState({
        hermesLinkedStatus: "stale",
        hermesValidationMessage: "Unable to verify session",
      });
      return { ok: false, code: "error", message: "Session check failed" };
    }
  };

  // core functionality
  const core = {
    async validateAndUpdateState() {
      if (state.checkingState) return;
      state.checkingState = true;

      try {
        const validationResult = await validateSession();
        await updateUI(validationResult.validation);
        return validationResult;
      } catch (e) {
        console.error("State check failed:", e);
      } finally {
        state.checkingState = false;
      }
    },

    async switchToLinkedTab() {
      try {
        const { hermesLinkedTabId } = await getLinkedState();
        if (!hermesLinkedTabId) {
          throw new Error("No linked tab found");
        }

        const tab = await chrome.tabs.get(hermesLinkedTabId).catch(() => null);
        if (!tab) {
          throw new Error("Linked tab no longer exists");
        }

        // get current window state
        const currentWindow = await chrome.windows.get(tab.windowId);

        // switch to window while preserving its state
        await chrome.windows.update(tab.windowId, {
          focused: true,
          // only pass state if it's not 'normal' to preserve maximized/fullscreen
          ...(currentWindow.state !== "normal" && {
            state: currentWindow.state,
          }),
        });

        // small delay before activating tab
        await new Promise((resolve) => setTimeout(resolve, 100));

        // activate the tab
        await chrome.tabs.update(hermesLinkedTabId, { active: true });
        await new Promise((resolve) => setTimeout(resolve, 250));
        await this.validateAndUpdateState();
      } catch (error) {
        throw error;
      }
    },

    async initialize() {
      if (state.isInitialized) return;

      // initial state check
      await this.validateAndUpdateState();

      // set up periodic check
      setInterval(() => {
        this.validateAndUpdateState().catch((e) =>
          console.error("Periodic check failed:", e)
        );
      }, PING_INTERVAL);

      state.isInitialized = true;
    },
  };

  // initialize core
  core
    .initialize()
    .catch((e) => console.error("Failed to initialize HermesLink:", e));

  // public api
  return {
    checkState: () => core.validateAndUpdateState(),
    goToLinkedTab: () => core.switchToLinkedTab(),
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
      const { hermesLinkedOrigin, hermesLinkedStatus } = await getLinkedState();
      return hermesLinkedStatus === "ok" ? hermesLinkedOrigin : null;
    },
  };
})();*/

window.HermesLink = (function () {
  const ENFORCE_ACTIVE_TAB_OVERLAY = true; // set to true to overlay on non-linked tabs always
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
  overlay: "Session not validated in this tab. Click Check Connection or Link This Tab Instead.",
},
    // TMS-specific states
    TMS_OK: {
      banner: "Active (TMS): ",
      hint: "Using Tenant Management System for the linked WFM tenant.",
      overlay: null,
    },
    TMS_NO_VANITY: {
      banner: "TMS: Vanity URL Required",
      hint: 'Enter or expose the "Vanity Non-SSO URL" for this tenant to use AccessPanel here.',
      overlay:
        'AccessPanel can only run on TMS when the "Vanity Non-SSO URL" is filled in and visible for this tenant.',
    },
    TMS_MISMATCH: {
      banner: "TMS Tenant Mismatch",
      hint: 'The "Vanity Non-SSO URL" in TMS does not match the active HerAccessPanelmes WFM session.',
      overlay:
        "This TMS tenant does not match your active AccessPanel session. Open the matching tenant or relink AccessPanel in WFM.",
    },
    // developer portal – always allow manual use
    DEV_OK: {
      banner: "Developer Portal: ",
      hint: "AccessPanel UI enabled for manual copy/paste from the developer portal.",
      overlay: null,
    },
  };

  // ---- TMS helpers ---- //

  // 1) ss this a TMS URL we should even consider?
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

  // 1b) Is this a developer portal URL where AccessPanel should stay enabled?
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

  const isBoomiUrl = (url) => {
    try {
      const u = new URL(url);
      const host = u.hostname.toLowerCase();
      // matches platform.boomi.com and any subdomain like foo.platform.boomi.com
      return (
        host === "platform.boomi.com" || host.endsWith(".platform.boomi.com")
      );
    } catch {
      return false;
    }
  };

  // 2) normalize any raw vanity string into the same shape as getVanityUrl()
  //    and compare only by protocol + host (ignore path and trailing slash).
  const normalizeVanity = (raw) => {
    if (!raw) return null;
    let value = String(raw).trim();
    if (!value) return null;

    // if user didn’t include the scheme, assume https
    if (!/^https?:\/\//i.test(value)) {
      value = "https://" + value;
    }

    try {
      // reuse existing logic that knows how to flip to -nosso, environment, etc.
      const normalizedUrl = getVanityUrl(value); // e.g. "https://24hf-nosso.prd.mykronos.com/"

      const parsed = new URL(normalizedUrl);
      // compare scheme + hostname only; ignore trailing slash or path
      return `${parsed.protocol}//${parsed.hostname}`.toLowerCase();
    } catch (e) {
      return null;
    }
  };

  // 3) read the TMS "Vanity Non-SSO URL" from the page in the active tab (search all frames)
  const getTmsVanityForTab = async (tabId) => {
    try {
      const results = await chrome.scripting.executeScript({
        target: { tabId, allFrames: true },
        func: () => {
          // recursively search document + all shadow roots for the "Vanity Non-SSO URL" field
          const findVanityInput = (root) => {
            if (!root) return null;

            const isVanityField = (sdfInput) => {
              if (!sdfInput || sdfInput.tagName !== "SDF-INPUT") return false;

              // 1) direct ID match
              if (sdfInput.id === "vanityNonSSOURL") return true;

              // 2) label text match inside its shadow root
              if (!sdfInput.shadowRoot) return false;

              const labelEl = sdfInput.shadowRoot.querySelector(
                ".sdf-form-control-wrapper--label, label[part='label']"
              );
              if (!labelEl || !labelEl.textContent) return false;

              const labelText = labelEl.textContent.trim();
              return labelText.startsWith("Vanity Non-SSO URL");
            };

            // direct root query first
            const direct = root.querySelector
              ? root.querySelector("sdf-input#vanityNonSSOURL")
              : null;
            if (direct && direct.shadowRoot && isVanityField(direct)) {
              const input = direct.shadowRoot.querySelector("input#input");
              if (input) return input;
            }

            // walk all elements, dive into shadow roots
            const walker = document.createTreeWalker(
              root,
              NodeFilter.SHOW_ELEMENT,
              null
            );

            let node = walker.currentNode;
            while (node) {
              if (node.tagName === "SDF-INPUT" && isVanityField(node)) {
                if (node.shadowRoot) {
                  const input = node.shadowRoot.querySelector("input#input");
                  if (input) return input;
                }
              }

              if (node.shadowRoot) {
                const fromShadow = findVanityInput(node.shadowRoot);
                if (fromShadow) return fromShadow;
              }

              node = walker.nextNode();
            }

            return null;
          };

          try {
            const input = findVanityInput(document);
            return input ? input.value.trim() : null;
          } catch (e) {
            console.error("TMS vanity search failed:", e);
            return null;
          }
        },
      });

      // results = [{frameId, result}, ...]; pick first non-null result
      if (Array.isArray(results)) {
        for (const r of results) {
          if (r && r.result) {
            return String(r.result).trim() || null;
          }
        }
      }

      return null;
    } catch (e) {
      return null;
    }
  };

  // 4) core TMS check: does the TMS tenant match the linked WFM session?
  const checkTmsTenantMatchesSession = async (
    currentTab,
    hermesLinkedOrigin
  ) => {
    if (!currentTab?.url || !isTmsUrl(currentTab.url)) {
      return { ok: false, reason: "not_tms" };
    }

    if (!hermesLinkedOrigin) {
      // No linked WFM session – TMS cannot be used yet
      return { ok: false, reason: "no_link" };
    }

    const sessionVanity = normalizeVanity(hermesLinkedOrigin);
    if (!sessionVanity) {
      return { ok: false, reason: "no_link" };
    }

    const tmsVanityRaw = await getTmsVanityForTab(currentTab.id);
    if (!tmsVanityRaw) {
      return { ok: false, reason: "no_tms_vanity" };
    }

    const tmsVanity = normalizeVanity(tmsVanityRaw);
    if (!tmsVanity) {
      return { ok: false, reason: "bad_tms_vanity" };
    }

    const match = tmsVanity === sessionVanity;

    if (!match) {
      return { ok: false, reason: "vanity_mismatch" };
    }

    return { ok: true, reason: "match" };
  };

  // state management
  const state = {
    isInitialized: false,
    checkingState: false,
  };

  // helper functions
  const getActiveTab = async () => {
    try {
      const [tab] = await chrome.tabs.query({
        active: true,
        currentWindow: true,
      });
      return tab || null;
    } catch (e) {
      console.error("Failed to get active tab:", e);
      return null;
    }
  };

  const getLinkedState = async () => {
    try {
      return await chrome.storage.session.get(Object.values(SESSION_KEYS));
    } catch (e) {
      console.error("Failed to get linked state:", e);
      return {};
    }
  };

  const updateLinkedState = async (newState) => {
    try {
      const timestamp = new Date().toISOString();
      await chrome.storage.session.set({
        ...newState,
        hermesLastValidation: timestamp,
      });
    } catch (e) {
      console.error("Failed to update linked state:", e);
    }
  };

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
    const isDevPortal = currentTab ? isDevPortalUrl(currentTab.url) : false;
    const isBoomi = currentTab ? isBoomiUrl(currentTab.url) : false;

    // reuse this for WFM-specific behavior (e.g., relink button)
    const currentTabValidation = currentTab
      ? validateWebPage(currentTab.url)
      : { valid: false };

    // figure out if the current WFM tab is the same tenant as the linked one.
    // if so, we can treat it as effectively "active" even if it's not the exact linked tab.
    let sameTenantAsLinked = false;
    let sessionVanity = null;

    if (hermesLinkedOrigin) {
      sessionVanity = normalizeVanity(hermesLinkedOrigin);
    }

    if (currentTab && currentTabValidation.valid && sessionVanity) {
      try {
        const currentVanity = normalizeVanity(currentTab.url);
        sameTenantAsLinked = !!currentVanity && currentVanity === sessionVanity;
      } catch (e) {
        sameTenantAsLinked = false;
      }
    }

    // if on TMS (and NOT on the linked WFM tab), check the tenant/vanity match
    let tmsCheck = null;
    if (!isLinkedTab && isTms) {
      try {
        tmsCheck = await checkTmsTenantMatchesSession(
          currentTab,
          hermesLinkedOrigin
        );
      } catch (e) {
        tmsCheck = { ok: false, reason: "error" };
      }
    }

    // determine current status
    let currentStatus = "OK";

    if (isDevPortal && !isLinkedTab) {
      // developer portal is always allowed for manual use – no overlay, no disable.
      currentStatus = "DEV_OK";
    } else if (isTms && tmsCheck) {
      if (tmsCheck.ok) {
        // TMS tenant matches the linked WFM session → allow full UI
        currentStatus = "TMS_OK";
      } else if (
        tmsCheck.reason === "no_tms_vanity" ||
        tmsCheck.reason === "bad_tms_vanity"
      ) {
        currentStatus = "TMS_NO_VANITY";
      } else if (tmsCheck.reason === "vanity_mismatch") {
        currentStatus = "TMS_MISMATCH";
      } else if (tmsCheck.reason === "no_link") {
        // no linked WFM session yet – behave like a normal wrong-tab state
        currentStatus = "WRONG_TAB";
      } else {
        currentStatus = "WRONG_TAB";
      }
    } else {
      // normal WFM-driven logic (WFM/mykronos tabs)
      if (!isLinkedTab && sameTenantAsLinked) {
        // different tab, but same tenant as the linked session → treat as active
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

      // relink button only makes sense on valid WFM tabs
      if (relinkButton) {
        if (currentTabValidation.valid) {
          relinkButton.style.display = "inline-block";
        } else {
          relinkButton.style.display = "none";
        }
      }

      // for OK, TMS_OK, DEV_OK, we never show the overlay.
      // if ENFORCE_ACTIVE_TAB_OVERLAY is false, we also avoid overlay for WRONG_TAB
      // so the UI stays usable even when you’re not on the linked tab.
      const isSoftWrongTab =
        currentStatus === "WRONG_TAB" && !ENFORCE_ACTIVE_TAB_OVERLAY;

      const overlayVisible =
        !["OK", "TMS_OK", "DEV_OK"].includes(currentStatus) &&
        !isBoomi &&
        !isSoftWrongTab;

      overlay.classList.toggle("visible", overlayVisible);
    }

    // ----- banner -----
    const banner = document.getElementById("hermes-link-banner");
    if (banner) {
      const status = document.getElementById("hermes-link-status");
      const target = document.getElementById("hermes-link-target");
      const hint = document.getElementById("hermes-link-hint");

      if (status) status.textContent = statusConfig.banner;
      if (target) target.textContent = currentTab?.title || "";
      if (hint) hint.textContent = hermesValidationMessage || statusConfig.hint;
    }

    // tab is considered "active" when:
    //   - we’re on the linked WFM tab (OK),
    //   - we’re on a matching TMS tenant (TMS_OK),
    //   - we’re on an allowed developer portal (DEV_OK).
    // if ENFORCE_ACTIVE_TAB_OVERLAY is false, WRONG_TAB is treated as a soft state:
    // banner + hints still show, but we don’t disable the UI.
    const isSoftWrongTabForUi =
      currentStatus === "WRONG_TAB" && !ENFORCE_ACTIVE_TAB_OVERLAY;

    const shouldDisableUi =
      !["OK", "TMS_OK", "DEV_OK"].includes(currentStatus) &&
      !isBoomi &&
      !isSoftWrongTabForUi;

    document.body.classList.toggle("tab-inactive", shouldDisableUi);
  };

  // session validation
  const validateSession = async () => {
    const { hermesLinkedUrl, hermesLinkedTabId } = await getLinkedState();

    if (!hermesLinkedUrl || !hermesLinkedTabId) {
      return { ok: false, code: "nolink", message: "No linked session found" };
    }

    try {
      const tab = await chrome.tabs.get(hermesLinkedTabId).catch(() => null);
      if (!tab) {
        await updateLinkedState({
          hermesLinkedStatus: "stale",
          hermesValidationMessage: "Linked tab was closed",
        });
        return { ok: false, code: "closed", message: "Linked tab was closed" };
      }

      const validation = validateWebPage(tab.url);
      if (!validation.valid) {
        await updateLinkedState({
          hermesLinkedStatus: "stale",
          hermesValidationMessage: validation.message,
        });
        return { ok: false, code: "invalid", validation };
      }

      // keep HermesLink in sync with whatever WFM URL is
      // actually loaded in that tab (handles “switched tenants
      // in the same tab”).
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

      return { ok: true, code: "ok", validation };
    } catch (e) {
      await updateLinkedState({
        hermesLinkedStatus: "stale",
        hermesValidationMessage: "Unable to verify session",
      });
      return { ok: false, code: "error", message: "Session check failed" };
    }
  };

  // core functionality
  const core = {
    async validateAndUpdateState() {
      if (state.checkingState) return;
      state.checkingState = true;

      try {
        const validationResult = await validateSession();
        await updateUI(validationResult.validation);
        return validationResult;
      } catch (e) {
        console.error("State check failed:", e);
      } finally {
        state.checkingState = false;
      }
    },

    async initialize() {
      if (state.isInitialized) return;

      // initial state check
      await this.validateAndUpdateState();

      // set up periodic check
      setInterval(() => {
        this.validateAndUpdateState().catch((e) =>
          console.error("Periodic check failed:", e)
        );
      }, PING_INTERVAL);

      state.isInitialized = true;
    },
  };

  // initialize core
  core
    .initialize()
    .catch((e) => console.error("Failed to initialize HermesLink:", e));

  // public api
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
    const { hermesLinkedOrigin, hermesLinkedStatus } = await getLinkedState();
    return hermesLinkedStatus === "ok" ? hermesLinkedOrigin : null;
  },
};

})();

// ============================= //

// ===== EVENT LISTENERS ===== //
document.addEventListener("DOMContentLoaded", async () => {
  await loadPreferences();
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
  on("admin-settings", "click", adminSettingsClick);
  onAsync("clear-all-data", "click", clearAllData);
  onAsync("clear-client-data", "click", clearClientData);
  on("settings-restore-defaults", "click", restoreSettingsDefaultsInForm);
  on("settings-cancel", "click", closeSettingsOverlay);
  onAsync("settings-save-exit", "click", async () => {
    await saveSettingsFromForm();
  });

  // links menu
  onAsync("links-boomi", "click", linksBoomi);
  onAsync("links-install-integrations", "click", linksInstallIntegrations);
  onAsync("links-developer-portal", "click", linksDeveloperPortal);

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
  onAsync("tms-pull-api", "click", pullApiFromTmsClick);
  onAsync("generate-birt-file", "click", generateBirtPropertiesClick);
  on("toggle-client-url", "click", () => toggleFieldVisibility("client-url", "toggle-client-url"), { onceKey: "reveal" });
  onAsync("refresh-client-url", "click", refreshClientUrlClick);
  onAsync("copy-client-url", "click", copyClientUrlClick);
  on("toggle-client-id", "click", () => toggleFieldVisibility("client-id", "toggle-client-id"), { onceKey: "reveal" });
  onAsync("save-client-id", "click", saveClientIDClick);
  onAsync("copy-client-id", "click", copyClientIdClick);
  on("toggle-client-secret", "click", toggleClientSecretVisibility);
  onAsync("copy-client-secret", "click", copyClientSecretClick);
  onAsync("save-client-secret", "click", saveClientSecretClick);
  on("toggle-tenant-id", "click", () => toggleFieldVisibility("tenant-id", "toggle-tenant-id"), { onceKey: "reveal" });
  onAsync("save-tenant-id", "click", saveTenantIdClick);
  onAsync("copy-tenant-id", "click", copyTenantIdClick);

  // api tokens
  on("toggle-access-token", "click", () => toggleFieldVisibility("access-token", "toggle-access-token"), { onceKey: "reveal" });
  onAsync("get-token", "click", fetchToken);
  onAsync("copy-token", "click", copyAccessToken);
  on("toggle-refresh-token", "click", () => toggleFieldVisibility("refresh-token", "toggle-refresh-token"), { onceKey: "reveal" });
  onAsync("refresh-access-token", "click", refreshAccessToken);
  onAsync("copy-refresh-token", "click", copyRefreshToken);

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
  onAsync("save-request", "click", onSaveRequestClick);
  onAsync("save-request-definition", "click", saveRequestDefinition);
  onAsync("export-bruno-request", "click", saveBrunoRequest);
  onAsync("export-env", "click", exportEnvironmentDefinition);

  updateRequestDependentButtons(false);
  updateResponseDependentButtons(false);

  // popout response
  on("popout-response", "click", popoutResponse);

  // wire escape my api ui
  document.addEventListener("keydown", (e) => {
    const overlay = document.getElementById("manage-myapis-overlay");
    if (e.key === "Escape" && overlay && !overlay.hidden) {
      closeMyApisManager();
    }
  });

  // hermeslink button handlers (these pass the button element, so keep inline)
  /*const returnButton = document.getElementById("hermes-return-to-tab");
  if (returnButton) {
    returnButton.addEventListener("click", () =>
      handleReturnToLinkedTab(returnButton)
    );
  }*/

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
          HermesLink.checkState().catch((e) => console.error("Visibility check failed:", e));
          resumeTokenTimersFromStorage().catch((e) => console.error("Timer resume failed:", e));
        })
        .catch((e) => console.error("Token purge failed:", e));
    }
  });


  // focus handler
  window.addEventListener("focus", () => {
    purgeExpiredTokensInStorage()
      .then(() => {
        HermesLink.checkState().catch((e) => console.error("Focus check failed:", e));
        resumeTokenTimersFromStorage().catch((e) => console.error("Timer resume failed:", e));
      })
      .catch((e) => console.error("Token purge failed:", e));
  });
});