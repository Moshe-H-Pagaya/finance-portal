# Finance AI & Automation Portal

Pagaya's Finance team portal for automation tools, AI assistants, and
reporting workflows. This is the Cloud Run port of the original Google
Apps Script web app — same UI, same Google Sheet as the data source,
but served from a container behind the GCP HTTPS Load Balancer with
Cloud IAP (federated to Okta) handling auth.

| Layer | Old (Apps Script) | New (this repo) |
|---|---|---|
| Hosting | `script.google.com` | Cloud Run in `finance-ai-497313` |
| URL | `script.google.com/a/macros/pagaya.com/s/...` | `portal.finance.pagaya.com` |
| Auth | Google login (Workspace domain) | Cloud IAP → Okta |
| Server-side code | `Code.gs` | `server.ts` (Express) |
| Data store | Cell A1 of `Cards` tab | **unchanged** — same cell, same sheet |
| Client calls | `google.script.run.xxx()` | `fetch('/api/xxx')` |
| Admin gate | Hardcoded `ADMIN_EMAILS` array | `ADMIN_EMAILS` env var |
| Identity source | `Session.getActiveUser().getEmail()` | `X-Goog-Authenticated-User-Email` (injected by IAP) |

## Quick start (local dev)

```bash
npm install
cp .env.example .env.local
# Edit .env.local:
#   CARDS_SHEET_ID=<your sheet id>
#   AUTH_MODE=dev
#   DEV_USER_EMAIL=moshe.halfon@pagaya.com   # to test the admin pages
#   GOOGLE_APPLICATION_CREDENTIALS=/abs/path/to/service-account-key.json

npm run dev
# → http://localhost:3000           portal
# → http://localhost:3000/admin     admin (requires DEV_USER_EMAIL in ADMIN_EMAILS)
```

The service account key file must have access to the Cards spreadsheet
(share the Sheet with the SA email as Editor).

## Project layout

```
finance-portal/
├── Dockerfile               # multi-stage node:22-alpine build
├── server.ts                # Express app — API routes + static serving
├── package.json
├── tsconfig.json
├── .env.example
├── public/                  # static HTML, served as-is
│   ├── portal.html          # the main portal (everyone behind IAP)
│   ├── admin.html           # admin table view (admins only)
│   ├── admin-form.html      # create/edit/delete one card
│   └── access-denied.html   # shown to non-admins hitting /admin
├── .github/workflows/
│   └── deploy.yml           # build → Artifact Registry → gcloud run deploy
└── docs/
    ├── DEPLOY.md            # one-time GCP setup + ongoing deploy guide
    └── PORT-NOTES.md        # what changed vs. the Apps Script original
```

## API

All endpoints require a valid IAP session (handled by the LB in front of
the service). Identity is read from `X-Goog-Authenticated-User-Email`.

| Method | Path | Auth | Replaces |
|---|---|---|---|
| `GET`  | `/api/me`              | any IAP user | (new) |
| `GET`  | `/api/cards`           | any IAP user | `getCards()` |
| `POST` | `/api/cards`           | admin only   | `updateCard(json)` |
| `DELETE` | `/api/cards/:id`     | admin only   | `deleteCard(id)` |
| `POST` | `/api/cards/reorder`   | admin only   | `reorderCards(ids)` |

## Data model

Cards are stored as a JSON array, JSON-stringified into cell `A1` of the
tab named `Cards` (or whatever `CARDS_SHEET_TAB` points to). One card:

```json
{
  "id": "credit-card-expenses-analysis",
  "name": "Credit Card Expenses Analysis",
  "category": "Accounts Payable",
  "description": "<p>Analyses…</p>",
  "type": "python_script",
  "badge": "Python Script",
  "url": "https://drive.google.com/drive/folders/...",
  "sheet_url": "https://docs.google.com/spreadsheets/d/...",
  "chat_url": "",
  "slack_url": "",
  "tools": ["slack", "google_sheet"],
  "business_owner": ["jane.doe@pagaya.com"]
}
```

## Deploy

See [`docs/DEPLOY.md`](docs/DEPLOY.md) for the one-time GCP setup
(project, Artifact Registry, Cloud Run, LB backend, IAP) and ongoing
deploy via GitHub Actions.
