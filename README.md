# CompTracker (FROI / Workers’ Comp Tracker) — Google Apps Script Web App

This repository mirrors a Google Apps Script (GAS) web app used to manage Workers’ Comp / FROI case data and related admin workflows.

Primary purpose of this repo:
- Speed up troubleshooting and code review by keeping a GitHub-readable copy of the `.gs` and `.html` files
- Enable line-specific reviews (permalinks/PRs) and “surgical” diffs

---

## Architecture Overview

### Server-side (Apps Script: `.gs`)
Apps Script backend code:
- Reads/writes Google Sheets data
- Enforces access control (domain/admin)
- Builds HTML templates and returns evaluated pages
- Serves JSON/data to the UI via `google.script.run`

### Client-side (HTML: `.html`)
UI code:
- Renders views and modals
- Handles clicks and UI state (read-only vs edit)
- Calls backend functions via `google.script.run`

---

## Entry / Routing (Web App)

Routing is handled in `Code.gs` via `doGet(e)` using `?page=`.

### Routes
- `?page=providers` → `Providers.html` (public directory)
- `?page=notice` → `Notice.html` (public, secured by signature)
  - Signature verification: `verifyNoticeSig_()` (server-side)
  - Notice payload: `getNoticeData(rid, row)` (EmailService.gs or related)
- `?page=reports` → `Reports.html` (admin reports dashboard)
- default (no `page`) → `Index.html` (FROI form)

### Helpers
- `include(filename)` for HTML partials/includes
- `getScriptUrl()` returns the deployed web app URL

---

## Configuration (Server-side)

### Domain Allowlist
- `ALLOWED_DOMAIN` is used to limit access (example: `@portlandmaine.gov`)
- Utility helpers:
  - `normalizeEmail_(email)`
  - `isAllowedDomain_(email)`

### Spreadsheet + Sheet Names
The app uses a Google Spreadsheet as the primary datastore.

Important:
- Do not commit real Spreadsheet IDs to a public repo.
- In GitHub, use placeholders and keep the real ID only in the Apps Script project (or use Script Properties).

Recommended public repo pattern:
- `SPREADSHEET_ID = 'REDACTED_IN_PUBLIC_REPO'`

#### Sheet bootstrap helper
- `ensureSheet_(name, headers)` creates the sheet if missing and ensures Row 1 headers exist.

---

## Repository File Map (Current)

### Core / Routing
- `Code.gs`
  - `doGet(e)` routing
  - shared utilities (`ensureSheet_`, parsing helpers)
  - `include(filename)` / `getScriptUrl()`

### Backend services (server-side `.gs`)
- `AuthService.gs`
  - Authentication and authorization helpers (domain/admin access)

- `AdminCases.gs`
  - Admin case list, case retrieval, case update endpoints

- `AdminSettings.gs`
  - Admin configuration/settings endpoints

- `FormService.gs`
  - FROI form read/write logic, field mapping, save/update operations

- `DashboardService.gs`
  - Dashboard data aggregation and queries

- `EmailService.gs`
  - Email workflows + notice-related payload generation (`getNoticeData` or similar)

- `Messaging.gs`
  - Case notes / comm log / messaging functions

- `Migration.gs`
  - Migration helpers / data normalization routines

- `Rebuildspreadsheetstructure.gs`
  - Spreadsheet structure rebuild/repair utilities

### Main views (UI `.html`)
- `Index.html`
  - Default FROI form entry page

- `View_AdminDashboard.html`
  - Admin dashboard view (case list, navigation into modals/details)

- `Providers.html`
  - Provider directory view (`?page=providers`)

- `Notice.html`
  - Signed notice view (`?page=notice`)

- `Reports.html`
  - Admin reports dashboard (`?page=reports`)

- `Settings.html`
  - Settings view (admin/user settings UI)

- `SettingsPartial.html`
  - Settings UI partial included by other views

### Modals / shared UI components
- `Modal_FroiForm.html`
  - FROI form modal (Incident Record + other sections)

- `Modal_CaseDetail.html`
  - Case detail modal

- `Modal_Dashboard.html`
  - Dashboard modal

- `Modal_Messaging.html`
  - Messaging/notes modal

- `Modal_Settings.html`
  - Settings modal

- `Buttonhandlers.html`
  - Shared click handlers / UI helpers (often where save/edit actions are wired)

---

## Troubleshooting Map (Where to look first)

### UI / layout / missing fields / wrong labels
Start in:
- `Index.html` (page shell / entry form UI)
- `Modal_FroiForm.html` (form content + fields)
- `Buttonhandlers.html` (if the behavior is click-driven: edit/save/cancel)

### Buttons don’t work / spinner forever / nothing happens
- `Buttonhandlers.html` (event wiring)
- The calling HTML (modal/view) for the exact button
- Search for `google.script.run.<functionName>` and confirm the backend function exists in:
  - `FormService.gs`, `AdminCases.gs`, `DashboardService.gs`, etc.

### Save doesn’t persist / saves to wrong column / values shifted
- `FormService.gs` (write/update logic)
- Any header-based mapping logic (Row 1 header assumptions)
- `ensureSheet_()` behavior (headers and structure)

### Admin dashboard problems (list doesn’t load / details won’t open)
- `View_AdminDashboard.html` (admin dashboard UI)
- `AdminCases.gs` (admin list + detail payload functions)

### Notice link / signature errors (`?page=notice`)
- `Code.gs` route: `page=notice`
- `verifyNoticeSig_()` (server-side)
- Notice builder: `EmailService.gs` (or wherever `getNoticeData()` is defined)

### Login / “not authorized” / domain issues
- `ALLOWED_DOMAIN` + `isAllowedDomain_()` (in `Code.gs`)
- `AuthService.gs` (admin checks and access logic)

### Settings issues
- `Settings.html`, `SettingsPartial.html`, `Modal_Settings.html`
- `AdminSettings.gs` (settings read/write)

---

## How to report an issue (fastest format)

Open a GitHub Issue per bug and include:

1) Where you were in the UI  
   Example: `Admin Dashboard → Open Case → FROI Form → Incident Record`

2) Steps to reproduce (3–6 bullets)

3) Expected vs actual

4) Errors (copy/paste)
- Apps Script execution error text
- Browser console error text (if relevant)

5) Code pointers (ideal, not required)
- GitHub permalink to the file + line range (`#Lx-Ly`)
- If you’re unsure, just name the UI area; reviewers can trace calls.

---

## Review workflow (recommended)

### Option A: Issues + direct commits
- One issue per bug
- Commit message references the issue
- Keep changes small and targeted

### Option B: Branch + Pull Request (best)
- Branch example: `fix/issue-12-incident-record-save`
- PR description includes repro steps + expected behavior
- Review changes as a diff (fast, safe)

---

## “Surgical change” rules (protect production behavior)

When making fixes, prefer the smallest safe change:
- Do not rename functions referenced by `google.script.run`, triggers, menus, or endpoints.
- Do not refactor unrelated code “for cleanliness.”
- Do not change sheet names/IDs/column order unless explicitly required.
- Preserve existing behavior except for the requested fix.
- If changing field mappings, ensure backward compatibility with existing sheet headers/rows.

---

## Security / secrets (public repo hygiene)

Do not commit:
- Real Spreadsheet IDs (if sensitive)
- OAuth tokens
- API keys
- Personal contact data, medical info, or case details

Recommended pattern for secrets:
- Use `PropertiesService.getScriptProperties()` in Apps Script
- Keep placeholders in this repo (`REDACTED_IN_PUBLIC_REPO`)

---

## Common call flow (how most features work)

1) UI action in `.html` (button click)
2) JS handler runs (often in `Buttonhandlers.html` or the same modal/view)
3) Calls backend via `google.script.run.someServerFunction(...)`
4) `.gs` function reads/writes Sheets (`FormService.gs`, `AdminCases.gs`, etc.)
5) Backend returns payload
6) UI updates DOM/state

Tip: when debugging, search for the `google.script.run` function name first.
