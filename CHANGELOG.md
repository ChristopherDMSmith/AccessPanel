# Changelog

All notable changes to AccessPanel are documented in this file.

---

## 2.0.0

### Security

- Moved Client Secret, Access Token, Refresh Token, and token expiration data from persistent local storage to session-scoped storage.
- Added migration cleanup for sensitive data stored by previous AccessPanel versions.
- Removed legacy `document.execCommand("copy")` clipboard fallbacks.
- Removed webpage script injection and the associated `scripting` browser permission.
- Removed TMS credential and page-content scraping.
- Removed the previous Incognito token retrieval mechanism that required script injection.
- Added an explicit Content Security Policy.
- Restricted Workforce Management host access to HTTPS.
- Removed obsolete host permissions.
- Replaced `document.write()` popup construction with DOM-based rendering.
- Removed remaining `innerHTML` usage from extension JavaScript.
- Standardized dynamic UI rendering around safer DOM APIs including `textContent` and `replaceChildren()`.
- Reduced persistence of API request data by removing saved custom API definitions.

### Changed

- Sensitive authentication information is now retained only for the current browser session.
- TMS remains a recognized AccessPanel context but AccessPanel no longer extracts credentials or tenant information from TMS pages.
- Automatic Access Token retrieval is no longer supported in Incognito mode.
- Simplified collapsible UI section state handling.
- Simplified API Library controls and request-dependent actions.

### Added

- Added manual Access Token entry using the **Apply Token** action.
- Manual Access Tokens are stored only for the current browser session.
- Manual Access Tokens can be used with the API Library in Incognito mode.
- Added explicit tracking of manually supplied versus automatically retrieved Access Tokens.
- Manual Access Tokens display **Manual** instead of an expiration countdown.

### Removed

- Removed saved **My APIs** functionality and persistent custom API request definitions.
- Removed Postman collection export.
- Removed Bruno request export.
- Removed environment export.
- Removed Postman/Bruno export preferences and related settings.
- Removed **Pull from TMS** functionality.
- Removed automatic Incognito token-page scraping.

### Fixed

- Corrected the Developer Portal link for **Timekeeping - Retrieve Punches**.
- Corrected initial collapsed-section rendering behavior.
- Preserved API response, tree view, popup, CSV export, BIRT, theme, and parameter-rendering behavior following DOM security hardening.

---

## 1.0.1

- Fix bg.js to preserve open state and prevent unintended closure.
- Other minor adjustmenmts.

---

## 1.0.0

- Initial publish release.