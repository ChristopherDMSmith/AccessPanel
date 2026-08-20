# AccessPanel

AccessPanel is a browser extension that provides a unified side panel for managing tenant context, API access tokens, and API requests within multi-tenant Workforce Management (WFM) environments.

It is designed for integration engineers, technical administrators, report developers, and API-focused support teams who need fast, consistent access to tenant-specific tooling directly from the browser.

AccessPanel is a Chromium-based browser extension officially distributed through the Chrome Web Store and Microsoft Edge Add-ons.

---

## Key Features

- Tenant-aware context management
- Automatic API Access Token retrieval
- Manual session Access Token support
- Built-in API Library
- Ad-Hoc GET and POST requests
- API response viewing, copying, downloading, and JSON tree navigation
- CSV export support for configured API Library requests
- BIRT properties generation
- Browser side panel workflow for in-context use
- Session-scoped storage of sensitive authentication data

---

## Installation

### Official Browser Stores

AccessPanel is officially distributed through:

- Chrome Web Store
- Microsoft Edge Add-ons

Install AccessPanel from the appropriate browser store, then:

1. Open AccessPanel from the browser toolbar.
2. Navigate to a supported Workforce Management tenant.
3. Configure tenant information as needed.
4. Retrieve an Access Token or manually apply an existing token.
5. Use the API Library or other AccessPanel utilities as needed.

### Development Installation

For development and testing:

1. Open Manage Extensions in Chrome or Edge.
2. Enable Developer mode.
3. Select **Load unpacked**.
4. Select the AccessPanel project folder.

---

## Intended Audience

AccessPanel is intended for technical users working with Workforce Management environments, including:

- Integration engineers
- Technical administrators
- Report developers
- API-focused support and operations teams

It is not intended for general consumer use.

---

## Authentication

AccessPanel supports two Access Token workflows.

### Automatic Access Token

In a normal browser session, AccessPanel can request an Access Token directly from the active Workforce Management tenant using the configured Client ID and authenticated browser session.

Automatically retrieved tokens include expiration information and display a countdown timer.

### Manual Access Token

An existing Access Token can be entered manually and applied to the current browser session.

Manual tokens can be used by the API Library and are particularly useful when AccessPanel is running in Incognito mode, where automatic token retrieval is intentionally unsupported.

Because AccessPanel does not know the expiration time of a manually supplied token, the token status displays **Manual** rather than an expiration countdown.

---

## Privacy & Data Handling

AccessPanel does not collect analytics, telemetry, or usage information and does not transmit user data to the developer.

Non-sensitive tenant configuration and extension preferences may be retained locally in browser extension storage.

Sensitive authentication data—including Client Secrets, Access Tokens, Refresh Tokens, and token expiration information—is maintained using session-scoped extension storage rather than persistent local storage.

AccessPanel does not persist user-entered API request bodies, query parameter values, or custom API request definitions between browser sessions.

See `PRIVACY.md` for the complete privacy and data-handling policy.

---

## Security

AccessPanel follows a least-privilege approach to browser permissions and extension functionality.

Security controls include:

- HTTPS-only Workforce Management host access
- Restrictive Content Security Policy
- Session-only sensitive credential storage
- No remote executable code
- No webpage script injection
- No credential or DOM scraping
- DOM-safe dynamic content rendering
- Restricted browser and host permissions

---

## Support & Feedback

Bug reports, questions, and feature requests can be submitted through GitHub Issues.

---

## Versioning

Current version: **2.0.0**

AccessPanel uses semantic versioning:

- **MAJOR** versions may contain significant behavioral or compatibility changes.
- **MINOR** versions add backward-compatible functionality.
- **PATCH** versions contain backward-compatible fixes.

---

## License

AccessPanel is proprietary software.

The source code is provided for transparency only. No permission is granted to modify, redistribute, sublicense, sell, or create derivative works without explicit written permission from the copyright holder.

See `LICENSE` for complete terms.