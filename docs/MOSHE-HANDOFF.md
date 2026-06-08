# Finance Portal — Handoff for Moshe

This is the 1-pager for the day-to-day owner. The portal is **live and
working** — you don't need to touch GCP unless something breaks. Most
changes (adding/editing cards) are done from the admin UI in the browser
and take effect instantly.

---

## Live URLs

| What                | Where                                                    |
| ------------------- | -------------------------------------------------------- |
| Portal (users)      | <https://finance-portal.gcp.pagaya.com/>                 |
| Admin console       | <https://finance-portal.gcp.pagaya.com/admin>            |
| Source repo         | (your GitHub repo URL — set after transfer)              |
| GCP project         | `finance-ai-497313` (region `us-east1`)                  |
| Cloud Run service   | `finance-portal`                                         |
| Cards storage       | `gs://finance-portal-cards/cards.json` (auto-versioned)  |

---

## How users access the portal

Anyone on the Pagaya office network, Cato VPN, or Zscaler can open
<https://finance-portal.gcp.pagaya.com/>. There is **no Google
SSO** — IP allowlist only (Cloud Armor policy `pagaya-local-policy`).

If a colleague reports `403 Forbidden` it means their IP is not in the
allowlist. Ask DevOps to add their office/VPN egress IP, or have them
connect to Cato.

---

## How to edit cards (the daily workflow)

1. Open <https://finance-portal.gcp.pagaya.com/admin>
2. Browser prompts for username + password. These were set by Amit and
   shared with you 1:1 — **do not commit them anywhere**.
3. You can:
   - **Add a card** — click "New Card", fill in the form, save.
   - **Edit a card** — click the card, change fields, save.
   - **Delete a card** — click delete on the card.
   - **Reorder cards** — drag and drop (saves on drop).
4. Changes are **instant** — refresh the portal in another tab to see them.

Behind the scenes every save writes a fresh `cards.json` into the GCS
bucket. Object Versioning is on, so every edit is automatically retained
as a separate generation — we can roll back without any code change.

### Rotating the admin password

When personnel changes, rotate the password from Cloud Shell:

```bash
NEW_PW="$(openssl rand -base64 24 | tr -d '/+=' | head -c 32)"
echo "new admin password: $NEW_PW"
gcloud run services update finance-portal --region=us-east1 \
  --update-env-vars="ADMIN_PASSWORD=$NEW_PW"
```

Cloud Run rolls out a new revision automatically. Share the new password
1:1 with whoever needs admin access.

---

## How code changes get deployed

If you ever need to change anything beyond cards (e.g. portal layout,
new field on the form, security tweak), the workflow is:

1. Clone the repo locally.
2. Edit the file. The interesting ones are:
   - `public/portal.html` — the user-facing page (HTML + CSS + the JS
     that renders cards).
   - `public/admin.html`, `public/admin-form.html` — the admin UI.
   - `server.ts` — the backend Express server (cards API, auth gate).
3. Test locally if you want:
   ```bash
   npm install
   npm run build
   ADMIN_USERNAME=admin ADMIN_PASSWORD=test123 npm start
   # open http://localhost:3000
   ```
4. Commit + push to `main`. **GitHub Actions auto-deploys to Cloud Run**
   on every push to `main` (see `.github/workflows/deploy.yml`).
5. Watch the deploy:
   ```bash
   gh run watch
   ```
6. Verify at <https://finance-portal.gcp.pagaya.com/>.

If you don't feel like editing code yourself, open a GitHub issue or
ping Amit / DevOps.

---

## Rolling back a bad card edit

GCS Object Versioning keeps every prior `cards.json`. To see history
and restore an older version:

```bash
# List every generation of cards.json (newest first).
gcloud storage ls -a gs://finance-portal-cards/cards.json

# Inspect a specific generation.
gcloud storage cat gs://finance-portal-cards/cards.json#<GENERATION_NUMBER>

# Restore by copying that generation over the live object.
gcloud storage cp \
  gs://finance-portal-cards/cards.json#<GENERATION_NUMBER> \
  gs://finance-portal-cards/cards.json
```

Refresh the portal — old cards are back.

---

## Rolling back a bad code deploy

Every Cloud Run deploy creates a numbered revision. To roll back:

```bash
# List recent revisions.
gcloud run revisions list --service=finance-portal --region=us-east1 --limit=5

# Send 100% of traffic to a known-good revision.
gcloud run services update-traffic finance-portal --region=us-east1 \
  --to-revisions=<REVISION_NAME>=100
```

---

## Troubleshooting

| Symptom                                | What it usually means                                                            | Where to look                                                              |
| -------------------------------------- | -------------------------------------------------------------------------------- | -------------------------------------------------------------------------- |
| User gets `403 Forbidden` (Google FE) | IP is not in `pagaya-local-policy` allowlist                                     | Ask DevOps to add their egress IP                                          |
| User gets `403 Forbidden` (small page) | Same — Cloud Armor block                                                          | Same                                                                       |
| Admin UI loads but save returns 500    | Service account lost `roles/storage.objectAdmin` on the bucket, or env var lost  | `gcloud storage buckets get-iam-policy gs://finance-portal-cards`          |
| `/api/cards` empty `[]`                | Bucket object was deleted; will auto-reseed from bundled `public/cards.json` next request | `gcloud storage ls gs://finance-portal-cards/`                             |
| Portal totally down                    | Cloud Run service crashed                                                         | Cloud Run console → revisions → logs, look for `[startup]` lines           |
| GitHub Actions deploy failed           | WIF token expired, or quota issue                                                 | `gh run view --log <RUN_ID>`                                               |

For any Cloud Run logs:

```bash
gcloud logging read \
  'resource.type=cloud_run_revision AND resource.labels.service_name=finance-portal' \
  --limit=30 --freshness=15m --order=desc \
  --format='value(timestamp,severity,textPayload)' \
  --project=finance-ai-497313
```

---

## Contacts

- **DevOps / GCP infra** — Amit Zack (`amit.zack@pagaya.com`)
- **Original Apps Script author** — Moshe (you)
- **Cloud Armor / IP allowlist** — DevOps team via the usual Jira
  queue. Reference Jira ticket `IS-49023` for the migration context.

---

## What this repo actually is (60-second tour)

- `server.ts` — Node.js + Express. Serves `/`, `/admin`, and the
  `/api/cards` REST endpoints. Talks to GCS via `@google-cloud/storage`.
- `public/portal.html` — the home page. Renders the cards grid.
- `public/admin.html` + `public/admin-form.html` — the admin console.
- `public/cards.json` — the **initial** card data that gets seeded into
  the GCS bucket on first deploy. After that, the live data lives in
  the bucket; this file is only a fallback.
- `Dockerfile` — how Cloud Build packages it.
- `.github/workflows/deploy.yml` — the auto-deploy pipeline (Workload
  Identity Federation → Cloud Build → Cloud Run).
- `docs/DEPLOY.md` — full GCP setup runbook (only needed if rebuilding
  from scratch).
- `docs/PORT-NOTES.md` — historical notes on the migration from
  Apps Script.
