# Deploy guide — Finance Portal

This is the one-time GCP infrastructure setup plus the ongoing deploy
workflow. Mirrors the pattern used by `mft-uploads-report-webapp`.

## One-time setup

### 1. GCP project

```bash
export PROJECT_ID=finance-ai-497313
export REGION=us-east1     # org policy on finance-ai-497313 restricts to us-east1 / us-west1
export SERVICE=finance-portal

# Project already exists — skip create/link. Just point gcloud at it.
gcloud config set project $PROJECT_ID

gcloud services enable \
  run.googleapis.com \
  artifactregistry.googleapis.com \
  iap.googleapis.com \
  compute.googleapis.com \
  sheets.googleapis.com \
  iamcredentials.googleapis.com
```

### 2. Artifact Registry

```bash
gcloud artifacts repositories create $SERVICE \
  --repository-format=docker \
  --location=$REGION \
  --description="Finance Portal container images"
```

### 3. Runtime service account

This SA runs the Cloud Run service and authenticates to Google Sheets.

```bash
gcloud iam service-accounts create finance-portal-run \
  --display-name="Finance Portal Cloud Run runtime"

export RUN_SA=finance-portal-run@$PROJECT_ID.iam.gserviceaccount.com
echo $RUN_SA
```

Share the Cards spreadsheet with `$RUN_SA` as **Editor** in Google Sheets.

### 4. GitHub Actions deploy SA (separate from the runtime SA)

```bash
gcloud iam service-accounts create gha-deployer \
  --display-name="GitHub Actions deployer"

export GHA_SA=gha-deployer@$PROJECT_ID.iam.gserviceaccount.com

gcloud projects add-iam-policy-binding $PROJECT_ID \
  --member="serviceAccount:$GHA_SA" \
  --role="roles/run.admin"

gcloud projects add-iam-policy-binding $PROJECT_ID \
  --member="serviceAccount:$GHA_SA" \
  --role="roles/artifactregistry.writer"

# Allow GHA SA to actAs the runtime SA when deploying
gcloud iam service-accounts add-iam-policy-binding $RUN_SA \
  --member="serviceAccount:$GHA_SA" \
  --role="roles/iam.serviceAccountUser"
```

Set up Workload Identity Federation for GitHub Actions (preferred over
SA keys) — same setup used for `mft-uploads-report-webapp`. Document
the WIF provider/pool here once configured:

```
WIF_PROVIDER=projects/<NUM>/locations/global/workloadIdentityPools/github/providers/github
WIF_SERVICE_ACCOUNT=$GHA_SA
```

### 5. First Cloud Run deploy (manual, just to bootstrap)

After pushing an initial image to Artifact Registry:

```bash
gcloud run deploy $SERVICE \
  --image=$REGION-docker.pkg.dev/$PROJECT_ID/$SERVICE/app:bootstrap \
  --region=$REGION \
  --service-account=$RUN_SA \
  --no-allow-unauthenticated \
  --ingress=internal-and-cloud-load-balancing \
  --set-env-vars="CARDS_SHEET_ID=<sheet id>,CARDS_SHEET_TAB=Cards,SHEET_PUBLIC=true,ADMIN_EMAILS=,ADMIN_USERNAME=admin,ADMIN_PASSWORD=<strong-secret>,AUTH_MODE=iap"
```

`--ingress=internal-and-cloud-load-balancing` locks the `*.run.app` URL
so only LB traffic can hit the container. Critical.

### 6. Wire into the existing HTTPS Load Balancer

The LB lives in the central LB project. Run these from that project.

```bash
export LB_PROJECT=<lb-project-id>

# Serverless NEG (cross-project — NEG in LB project, points at Cloud Run
# in the app project)
gcloud compute network-endpoint-groups create finance-portal-neg \
  --project=$LB_PROJECT \
  --region=$REGION \
  --network-endpoint-type=serverless \
  --cloud-run-service=$SERVICE \
  --cloud-run-url-mask="/" \
  --cloud-run-service-project=$PROJECT_ID

# Backend service
gcloud compute backend-services create finance-portal-bes \
  --project=$LB_PROJECT \
  --global \
  --load-balancing-scheme=EXTERNAL_MANAGED

gcloud compute backend-services add-backend finance-portal-bes \
  --project=$LB_PROJECT \
  --global \
  --network-endpoint-group=finance-portal-neg \
  --network-endpoint-group-region=$REGION

# URL map host rule (edit the existing url-map for the LB)
# Either via console or:
gcloud compute url-maps add-path-matcher <url-map-name> \
  --project=$LB_PROJECT \
  --path-matcher-name=finance-portal-matcher \
  --new-hosts=portal.finance.pagaya.com \
  --default-service=finance-portal-bes
```

### 7. Enable IAP

```bash
gcloud compute backend-services update finance-portal-bes \
  --project=$LB_PROJECT \
  --global \
  --iap=enabled

# Grant @pagaya.com access at the backend level
gcloud iap web add-iam-policy-binding \
  --project=$LB_PROJECT \
  --resource-type=backend-services \
  --service=finance-portal-bes \
  --member='domain:pagaya.com' \
  --role='roles/iap.httpsResourceAccessor'
```

The IAP integration with Okta should already exist (Workforce Identity
Federation) from the `mft-uploads-report-webapp` setup. No changes
needed there.

### 8. DNS

Add an A record on `pagaya.com` pointing `portal.finance` at the LB's
external IPv4 IP. If the LB has IPv6, add AAAA too.

```bash
# On Cloud DNS:
gcloud dns record-sets create portal.finance.pagaya.com. \
  --project=<dns-project-id> \
  --zone=pagaya-com \
  --type=A --ttl=300 \
  --rrdatas=<LB_IPV4>
```

Confirm the LB's TLS cert covers `portal.finance.pagaya.com` (either
via a `*.pagaya.com` wildcard or by adding it to a managed cert).

### 9. Verify

- `dig portal.finance.pagaya.com` → LB IP.
- `curl -I https://portal.finance.pagaya.com` → 302 to IAP.
- In a `@pagaya.com` browser → see the portal.
- In incognito with personal Gmail → IAP "no access" page.
- As `moshe.halfon@pagaya.com` → `/admin` works, can edit cards.

## Ongoing deploys

After the one-time setup is done, every push to `main` triggers
`.github/workflows/deploy.yml`:

1. Build the Docker image.
2. Tag and push to Artifact Registry.
3. `gcloud run deploy` with the new image.

Cloud Run does zero-downtime traffic shifting. If a deploy goes bad,
`gcloud run services update-traffic $SERVICE --to-revisions <prev>=100`.

## Environment variables

Set these on the Cloud Run service (not in the repo):

| Var | Required | Value |
|---|---|---|
| `CARDS_SHEET_ID` | yes | Spreadsheet ID from the sheet URL |
| `CARDS_SHEET_TAB` | no | Defaults to `Cards` |
| `SHEET_PUBLIC` | yes* | `true` when the sheet is shared as **Anyone with the link, Viewer** (needed if org policy blocks sharing with the runtime SA). `false` when the sheet is shared with `finance-portal-run@…` as Editor. |
| `ADMIN_USERNAME` | yes | Basic Auth username for `/admin` and admin-write APIs |
| `ADMIN_PASSWORD` | yes | Basic Auth password (rotate on personnel changes) |
| `ADMIN_AUTH_REALM` | no | Browser login dialog title; defaults to `Finance Portal Admin` |
| `ADMIN_EMAILS` | no | Comma-separated email allowlist; only enforced when Cloud IAP injects identity headers. Leave empty for Basic-Auth-only deployments. |
| `AUTH_MODE` | yes | `iap` in production (even without IAP enabled — just means no fake dev user), `dev` for local |
| `DEV_USER_EMAIL` | no | Local dev only |
| `PORT` | no | Cloud Run injects this automatically |

\* Required in practice for Pagaya orgs that block sharing Sheets with
`@*.iam.gserviceaccount.com` addresses. When `SHEET_PUBLIC=true`, admin
writes via `/admin` are disabled — edit cell A1 in the Sheet directly.

### Wire up live card data (after bootstrap)

Once the Cards sheet exists and cell A1 holds the JSON array:

```bash
gcloud run services update finance-portal \
  --project=finance-ai-497313 \
  --region=us-east1 \
  --update-env-vars="CARDS_SHEET_ID=11rvPaVslzd7aQtuqArbpwauTeKui7I8QL4vvzDTTP5I,CARDS_SHEET_TAB=Cards,SHEET_PUBLIC=true"
```

If the sheet is shared with the runtime SA instead, omit `SHEET_PUBLIC` or
set it to `false`.

Do **not** set `GOOGLE_APPLICATION_CREDENTIALS` on Cloud Run — ADC picks
up the runtime SA automatically.

## Rollback

```bash
gcloud run revisions list --service=$SERVICE --region=$REGION
gcloud run services update-traffic $SERVICE \
  --region=$REGION \
  --to-revisions=<previous-revision>=100
```

Or fully fall back to the Apps Script URL — leave that deployment live
for at least 2 weeks post-cutover.
