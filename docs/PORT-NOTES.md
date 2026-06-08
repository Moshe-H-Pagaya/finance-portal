# Port notes: Apps Script → Cloud Run

Faithful 1:1 port. The card data, sheet, JSON schema, admin allowlist
semantics, and UI layout are all identical to the original. Only the
hosting + auth + transport layer changed.

## What's the same

- Cards live in **cell A1** of the `Cards` tab in the same spreadsheet,
  JSON-stringified array. No schema migration needed.
- All three HTML pages render the same DOM and use the same CSS as the
  Apps Script originals.
- Admin emails are still an allowlist; same semantics, just moved from a
  hardcoded array in `Code.gs` to the `ADMIN_EMAILS` env var.
- `id` slugification for new cards is unchanged.
- Drag-and-drop reorder, edit, delete, search, category tabs — all
  behave identically.

## What changed

| Concern | Apps Script | Cloud Run |
|---|---|---|
| Identity | `Session.getActiveUser().getEmail()` | `X-Goog-Authenticated-User-Email` header from IAP, with the `accounts.google.com:` prefix stripped |
| Routing | `doGet(e)` looking at `e.parameter.page` | Express routes: `/`, `/admin`, `/admin-form` |
| Server calls | `google.script.run.withSuccessHandler(...).foo(...)` | `fetch('/api/foo')` returning `{ok, ...}` |
| Sheet access | `SpreadsheetApp.openById()` | Google Sheets API v4 via the runtime service account |
| Sheet permission | The deploying user's Google session | Sheet shared with `finance-portal-run@<project>.iam.gserviceaccount.com` as Editor |
| Admin "Access Denied" | Inline HTML returned from `doGet` | `public/access-denied.html` served at 403 |
| Template tags | `<?= cardId ?>` in `admin-form.html` | URL query string parsed by client JS |
| Iframe options | `setXFrameOptionsMode(ALLOWALL)` | Not needed — we're not iframed |
| Web app URL | `ScriptApp.getService().getUrl()` | Always relative (`/`, `/admin`, etc.) |

## Server-side function mapping

| Apps Script | HTTP endpoint | Notes |
|---|---|---|
| `getCards()` | `GET /api/cards` | Returns `{ok, cards}` |
| `updateCard(json)` | `POST /api/cards` | Body is the card object directly |
| `deleteCard(id)` | `DELETE /api/cards/:id` | |
| `reorderCards(idsJson)` | `POST /api/cards/reorder` | Body: `{ids: string[]}` |
| `getWebAppUrl()` | (removed) | Not needed — everything is relative |
| `isAdmin_()` | `requireAdmin` middleware | Reads IAP header, checks env allowlist |
| `setup() / setSheetId()` | (removed) | Sheet ID is set via `CARDS_SHEET_ID` env var |

## Identity header details

Cloud IAP injects two headers on every authenticated request:

- `X-Goog-Authenticated-User-Email: accounts.google.com:user@pagaya.com`
- `X-Goog-Authenticated-User-Id: accounts.google.com:<numeric-id>`

`server.ts` reads the email header, strips the `accounts.google.com:`
prefix, lowercases, and uses that as the user identity for admin checks.
The Cloud Run service is configured with
`--ingress=internal-and-cloud-load-balancing`, so requests from
anywhere except the LB are dropped — meaning we can fully trust these
headers without additional JWT verification. If we ever change ingress
settings, we'd need to additionally verify the `X-Goog-IAP-JWT-Assertion`
JWT signature.

## What was deliberately not ported

- `setup()` and `setSheetId()` helpers — no longer needed; sheet ID is
  set declaratively via env var.
- The favicon URL still points at `https://pagaya.com/favicon.ico`
  (same as original). If we want to self-host, drop a `favicon.png` in
  `public/` and update the `<link>` in all three HTML files.
