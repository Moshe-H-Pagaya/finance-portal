// =============================================================================
// Finance AI & Automation Portal — Express server
// =============================================================================
// 1:1 port of the original Google Apps Script Code.gs, plus:
//   • Identity comes from the X-Goog-Authenticated-User-Email header injected
//     by Cloud IAP (replaces Session.getActiveUser().getEmail()).
//   • Data still lives in cell A1 of the Cards tab in the Google Sheet,
//     accessed via the Sheets API with the runtime service account.
//   • Admin allowlist moved from a hardcoded array to env (ADMIN_EMAILS).
// =============================================================================

import { execFile } from "node:child_process";
import dns from "node:dns";
import net from "node:net";
import express, { type Request, type Response, type NextFunction } from "express";
import path from "node:path";
import fs from "node:fs";
import { google, type sheets_v4 } from "googleapis";
import { Storage } from "@google-cloud/storage";
import { VertexAI } from "@google-cloud/vertexai";

// ── Config ──────────────────────────────────────────────────────────────────
const PORT = Number(process.env.PORT) || 3000;
const SHEET_ID = process.env.CARDS_SHEET_ID || "";
const SHEET_TAB = process.env.CARDS_SHEET_TAB || "Cards";
const ADMIN_EMAILS = (process.env.ADMIN_EMAILS || "")
  .split(",")
  .map((s) => s.trim().toLowerCase())
  .filter(Boolean);
const AUTH_MODE = (process.env.AUTH_MODE || "iap").toLowerCase();
const DEV_USER_EMAIL = (process.env.DEV_USER_EMAIL || "").toLowerCase();

// Shared-secret admin login used by the HTTP Basic Auth gate on /admin*.
// Cloud Armor already restricts who can reach the LB; this adds a
// username/password prompt on top so even an allowed-network user has to
// prove they're the admin before reaching the management UI.
const ADMIN_USERNAME = process.env.ADMIN_USERNAME || "";
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || "";
const ADMIN_AUTH_REALM = process.env.ADMIN_AUTH_REALM || "Finance Portal Admin";

// When SHEET_PUBLIC=true, the sheet is shared as "Anyone with the link, Viewer"
// and we read it via the public gviz endpoint. No SA auth needed. Writes are
// disabled in this mode (the admin would edit the sheet directly in Google
// Sheets UI). When false (default), we authenticate to the Sheets API as the
// Cloud Run runtime SA — the sheet must be shared with that SA as Editor.
const SHEET_PUBLIC = (process.env.SHEET_PUBLIC || "").toLowerCase() === "true";

// SOX Agent Cloud Run Job — triggered via POST /api/sox-run (admin only)
const SOX_JOB_NAME = process.env.SOX_JOB_NAME || "sox-agent";
const SOX_JOB_REGION = process.env.SOX_JOB_REGION || "us-east1";
const SOX_JOB_PROJECT = process.env.SOX_JOB_PROJECT || "finance-ai-497313";

// GCS-backed storage for cards. When CARDS_GCS_BUCKET is set, the Express
// app reads and writes a single JSON object in that bucket. Runtime SA needs
// roles/storage.objectAdmin on the bucket. This is the path used when the
// portal must allow /admin writes but org policy blocks sharing Sheets with
// the SA email — the same pattern as uploads-report-app.
const CARDS_GCS_BUCKET = process.env.CARDS_GCS_BUCKET || "";
const CARDS_GCS_OBJECT = process.env.CARDS_GCS_OBJECT || "cards.json";

// Decide where card data comes from, in priority order:
//   1. CARDS_GCS_BUCKET set → GCS object (read + write supported)
//   2. CARDS_SHEET_ID set + SHEET_PUBLIC=true → public gviz endpoint (read only)
//   3. CARDS_SHEET_ID set + SHEET_PUBLIC=false → authenticated Sheets API (SA)
//   4. None of the above → bundled file <static>/cards.json (read only)
if (CARDS_GCS_BUCKET) {
  console.log(
    `[startup] Cards storage: GCS gs://${CARDS_GCS_BUCKET}/${CARDS_GCS_OBJECT}`,
  );
} else if (SHEET_ID) {
  console.log(
    `[startup] Cards storage: Google Sheet ${SHEET_ID} (public=${SHEET_PUBLIC})`,
  );
} else {
  console.warn(
    "[startup] Cards storage: bundled cards.json (read-only; set CARDS_GCS_BUCKET for admin writes).",
  );
}
if (!ADMIN_USERNAME || !ADMIN_PASSWORD) {
  console.warn(
    "[startup] ADMIN_USERNAME / ADMIN_PASSWORD not set — /admin will refuse ALL requests.",
  );
}

// ── Google Sheets client (lazy, singleton) ──────────────────────────────────
let sheetsClient: sheets_v4.Sheets | null = null;
async function sheets(): Promise<sheets_v4.Sheets> {
  if (sheetsClient) return sheetsClient;
  const auth = new google.auth.GoogleAuth({
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  sheetsClient = google.sheets({ version: "v4", auth: await auth.getClient() as any });
  return sheetsClient;
}

// ── Google Cloud Storage client (lazy, singleton) ──────────────────────────
let gcsClient: Storage | null = null;
function gcs(): Storage {
  if (gcsClient) return gcsClient;
  gcsClient = new Storage();
  return gcsClient;
}

// ── Storage helpers (A1 of <SHEET_TAB> holds a JSON-encoded array) ──────────
interface Card {
  id: string;
  name?: string;
  category?: string;
  description?: string;
  type?: string;
  badge?: string;
  url?: string;
  sheet_url?: string;
  chat_url?: string;
  slack_url?: string;
  tools?: string[];
  business_owner?: string[];
}

async function readCards(): Promise<Card[]> {
  // Priority 1: GCS object (read+write).
  if (CARDS_GCS_BUCKET) {
    return readCardsFromGcs();
  }
  // Priority 2: Google Sheet.
  if (SHEET_ID) {
    const raw = SHEET_PUBLIC
      ? await readCellPublic(SHEET_ID, SHEET_TAB, "A1")
      : await readCellPrivate(SHEET_ID, SHEET_TAB, "A1");
    if (!raw) return [];
    try {
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? parsed : [];
    } catch (err) {
      console.error("[readCards] Cell A1 is not valid JSON:", err);
      return [];
    }
  }
  // Priority 3: bundled-file mode.
  return readCardsFromFile();
}

// GCS read. On first call (or when the object is missing) we seed the bucket
// with whatever is in the bundled cards.json. This makes the first deploy
// "just work" — admins can immediately edit, and we don't need a separate
// seeding step.
async function readCardsFromGcs(): Promise<Card[]> {
  const file = gcs().bucket(CARDS_GCS_BUCKET).file(CARDS_GCS_OBJECT);
  try {
    const [buf] = await file.download();
    const parsed = JSON.parse(buf.toString("utf8"));
    return Array.isArray(parsed) ? parsed : [];
  } catch (err: any) {
    if (err?.code === 404) {
      // First-run seed from the bundled cards.json.
      const seed = readCardsFromFile();
      console.log(
        `[readCardsFromGcs] gs://${CARDS_GCS_BUCKET}/${CARDS_GCS_OBJECT} missing; seeding from bundled file (${seed.length} cards)`,
      );
      try {
        await writeCardsToGcs(seed);
      } catch (seedErr) {
        console.error("[readCardsFromGcs] seed write failed:", seedErr);
      }
      return seed;
    }
    console.error("[readCardsFromGcs] download failed:", err);
    throw err;
  }
}

// Resolves cards.json from the same directory the static assets live in.
// Walks the same candidate list as resolveStaticDir() so dev (./public) and
// prod (./dist/public) both work, without depending on STATIC_DIR being
// defined yet at this point in the file.
function resolveCardsFile(): string {
  const candidates = [
    path.resolve(process.cwd(), "dist", "public", "cards.json"),
    path.resolve(process.cwd(), "public", "cards.json"),
  ];
  for (const c of candidates) if (fs.existsSync(c)) return c;
  return candidates[candidates.length - 1];
}

function readCardsFromFile(): Card[] {
  const file = resolveCardsFile();
  try {
    const raw = fs.readFileSync(file, "utf8");
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (err) {
    console.error(`[readCardsFromFile] failed to read ${file}:`, err);
    return [];
  }
}

// Private path: authenticated read via the Sheets v4 API as the runtime SA.
// Requires the sheet to be shared with the SA.
async function readCellPrivate(
  spreadsheetId: string,
  tab: string,
  cell: string,
): Promise<string> {
  const api = await sheets();
  const res = await api.spreadsheets.values.get({
    spreadsheetId,
    range: `${tab}!${cell}`,
  });
  const v = res.data.values?.[0]?.[0];
  return typeof v === "string" ? v : "";
}

// Public path: anonymous read via the gviz endpoint. Works when the sheet is
// shared as "Anyone with the link". Returns the raw cell content as a string.
// gviz returns a wrapped JSON like:
//   /*O_o*/ google.visualization.Query.setResponse({...,"table":{"rows":[{"c":[{"v":"..."}]}]}});
async function readCellPublic(
  spreadsheetId: string,
  tab: string,
  cell: string,
): Promise<string> {
  const url =
    `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq` +
    `?sheet=${encodeURIComponent(tab)}&range=${encodeURIComponent(cell)}` +
    `&tqx=out:json`;
  const res = await fetch(url, { redirect: "follow" });
  if (!res.ok) {
    throw new Error(`gviz request failed: HTTP ${res.status}`);
  }
  const text = await res.text();
  const m = text.match(/google\.visualization\.Query\.setResponse\((.*)\);?\s*$/);
  if (!m) throw new Error("gviz response not in expected shape");
  const payload = JSON.parse(m[1]);
  const v = payload?.table?.rows?.[0]?.c?.[0]?.v;
  if (typeof v === "string") return v;
  // gviz may return structured cell values as objects rather than strings.
  if (v != null) return JSON.stringify(v);
  return "";
}

async function writeCards(cards: Card[]): Promise<void> {
  if (CARDS_GCS_BUCKET) {
    await writeCardsToGcs(cards);
    return;
  }
  if (!SHEET_ID) {
    throw new Error(
      "Write disabled: no persistent storage configured. " +
        "Set CARDS_GCS_BUCKET to enable admin writes, or edit public/cards.json " +
        "in the source repo and redeploy.",
    );
  }
  if (SHEET_PUBLIC) {
    // In public-read mode we can't write back via the public gviz endpoint.
    // Admins edit the sheet directly in Google Sheets UI. Surface a clear
    // 4xx-ish error instead of a confusing API auth failure.
    throw new Error(
      "Write disabled: sheet is configured as public-read (SHEET_PUBLIC=true). " +
        "Edit cards directly in the Google Sheet, or set CARDS_GCS_BUCKET / share " +
        "the sheet with the runtime SA as Editor.",
    );
  }
  const api = await sheets();
  await api.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_TAB}!A1`,
    valueInputOption: "RAW",
    requestBody: { values: [[JSON.stringify(cards)]] },
  });
}

async function writeCardsToGcs(cards: Card[]): Promise<void> {
  const file = gcs().bucket(CARDS_GCS_BUCKET).file(CARDS_GCS_OBJECT);
  await file.save(JSON.stringify(cards, null, 2), {
    contentType: "application/json",
    // GCS Object Versioning gives us history; we don't need to manage
    // generations explicitly here.
    resumable: false,
  });
}

// ── Identity & admin gate ───────────────────────────────────────────────────
function getUserEmail(req: Request): string {
  if (AUTH_MODE === "dev") return DEV_USER_EMAIL;
  // Cloud IAP injects this header on every authenticated request.
  // Format: "accounts.google.com:user@pagaya.com" — we strip the prefix.
  const raw = req.header("X-Goog-Authenticated-User-Email") || "";
  const idx = raw.indexOf(":");
  return (idx >= 0 ? raw.slice(idx + 1) : raw).toLowerCase();
}

function isAdmin(email: string): boolean {
  return !!email && ADMIN_EMAILS.includes(email);
}

function requireAdmin(req: Request, res: Response, next: NextFunction): void {
  const email = getUserEmail(req);
  if (!isAdmin(email)) {
    res.status(403).json({ ok: false, error: "Not authorised" });
    return;
  }
  next();
}

// Constant-time string compare to avoid leaking the password via timing.
function safeEqual(a: string, b: string): boolean {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function parseBasicAuth(header: string | undefined): { user: string; pass: string } | null {
  if (!header || !header.toLowerCase().startsWith("basic ")) return null;
  const decoded = Buffer.from(header.slice(6).trim(), "base64").toString("utf8");
  const idx = decoded.indexOf(":");
  if (idx < 0) return null;
  return { user: decoded.slice(0, idx), pass: decoded.slice(idx + 1) };
}

// HTTP Basic Auth gate for /admin pages and admin-write APIs.
// When credentials are missing/wrong, returns 401 with WWW-Authenticate so
// the browser shows its native login dialog. With credentials present and
// matching ADMIN_USERNAME/ADMIN_PASSWORD, the request passes through.
function requireBasicAuth(req: Request, res: Response, next: NextFunction): void {
  if (!ADMIN_USERNAME || !ADMIN_PASSWORD) {
    res.status(503).json({ ok: false, error: "Admin auth not configured" });
    return;
  }
  const creds = parseBasicAuth(req.header("authorization"));
  if (
    creds &&
    safeEqual(creds.user, ADMIN_USERNAME) &&
    safeEqual(creds.pass, ADMIN_PASSWORD)
  ) {
    next();
    return;
  }
  res.set("WWW-Authenticate", `Basic realm="${ADMIN_AUTH_REALM}", charset="UTF-8"`);
  res.status(401).send("Authentication required");
}

// Combined guard for admin write API endpoints: Basic Auth + optional email
// allowlist check (matches the two-layer check already on the admin pages).
// When ADMIN_EMAILS is empty the email check is skipped so existing deploys
// with no allowlist configured continue to work exactly as before.
function requireAdminAccess(req: Request, res: Response, next: NextFunction): void {
  requireBasicAuth(req, res, () => {
    if (ADMIN_EMAILS.length === 0) { next(); return; }
    const email = getUserEmail(req);
    if (!isAdmin(email)) {
      res.status(403).json({ ok: false, error: "Not authorised" });
      return;
    }
    next();
  });
}

// ── SSRF-guarded URL fetch (for the admin "Polish" feature) ─────────────────
// Optionally pulls the App URL's page content as extra context for the LLM.
// Since this is a server-side request to an admin-supplied URL, we guard
// against SSRF: only http/https, and every resolved IP (across redirects)
// must be public — blocks loopback, private ranges, link-local and the GCP
// metadata IP (169.254.169.254).

function ipIsPrivate(ip: string): boolean {
  const type = net.isIP(ip);
  if (type === 4) {
    const [a, b] = ip.split(".").map((n) => parseInt(n, 10));
    if (a === 0 || a === 10 || a === 127) return true;
    if (a === 169 && b === 254) return true;            // link-local + metadata
    if (a === 172 && b >= 16 && b <= 31) return true;   // 172.16/12
    if (a === 192 && b === 168) return true;            // 192.168/16
    if (a === 100 && b >= 64 && b <= 127) return true;  // CGNAT 100.64/10
    return false;
  }
  if (type === 6) {
    const lower = ip.toLowerCase();
    if (lower === "::1" || lower === "::") return true;            // loopback
    if (lower.startsWith("fe80")) return true;                    // link-local
    if (lower.startsWith("fc") || lower.startsWith("fd")) return true; // unique-local
    const mapped = lower.match(/::ffff:(\d+\.\d+\.\d+\.\d+)$/);    // IPv4-mapped
    if (mapped) return ipIsPrivate(mapped[1]);
    return false;
  }
  return true; // not a valid IP literal → treat as unsafe
}

async function hostResolvesToPublicIp(host: string): Promise<boolean> {
  if (net.isIP(host)) return !ipIsPrivate(host);
  let addrs: dns.LookupAddress[];
  try {
    addrs = await dns.promises.lookup(host, { all: true });
  } catch {
    return false;
  }
  return addrs.length > 0 && addrs.every((a) => !ipIsPrivate(a.address));
}

function htmlToText(html: string): string {
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/\s+/g, " ")
    .trim();
}

async function fetchUrlSafely(
  rawUrl: string,
): Promise<{ ok: true; text: string } | { ok: false }> {
  const MAX_HOPS = 3;
  const MAX_BYTES = 1_500_000;
  const MAX_TEXT_CHARS = 16_000;
  const TIMEOUT_MS = 8_000;
  let current = rawUrl;

  for (let hop = 0; hop <= MAX_HOPS; hop++) {
    let url: URL;
    try {
      url = new URL(current);
    } catch {
      return { ok: false };
    }
    if (url.protocol !== "http:" && url.protocol !== "https:") return { ok: false };
    if (!(await hostResolvesToPublicIp(url.hostname))) return { ok: false };

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), TIMEOUT_MS);
    let res: Awaited<ReturnType<typeof fetch>>;
    try {
      res = await fetch(url.toString(), {
        method: "GET",
        redirect: "manual",
        signal: controller.signal,
        headers: {
          "User-Agent": "FinancePortal-Polish/1.0",
          Accept: "text/html,text/plain,*/*",
        },
      });
    } catch {
      clearTimeout(timer);
      return { ok: false };
    }
    clearTimeout(timer);

    // Re-validate each redirect hop's host manually.
    if (res.status >= 300 && res.status < 400) {
      const loc = res.headers.get("location");
      if (!loc || hop === MAX_HOPS) return { ok: false };
      try {
        current = new URL(loc, url).toString();
      } catch {
        return { ok: false };
      }
      continue;
    }

    if (!res.ok) return { ok: false };
    const ctype = (res.headers.get("content-type") || "").toLowerCase();
    if (!ctype.startsWith("text/")) return { ok: false };

    const reader = res.body?.getReader();
    if (!reader) {
      const t = await res.text();
      return { ok: true, text: htmlToText(t).slice(0, MAX_TEXT_CHARS) };
    }
    let received = 0;
    const chunks: Buffer[] = [];
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      if (value) {
        received += value.length;
        if (received > MAX_BYTES) {
          try { await reader.cancel(); } catch { /* ignore */ }
          break;
        }
        chunks.push(Buffer.from(value));
      }
    }
    const text = Buffer.concat(chunks).toString("utf8");
    return { ok: true, text: htmlToText(text).slice(0, MAX_TEXT_CHARS) };
  }
  return { ok: false };
}

// Server-side allowlist sanitizer for the polished description HTML. Mirrors
// the DOMPurify allowlist used on the portal: keep a small set of formatting
// tags, drop every attribute. The client sanitizes again (defense in depth).
function sanitizeDescriptionHtml(html: string): string {
  const ALLOWED = new Set([
    "b", "strong", "i", "em", "code", "ul", "ol", "li", "br", "p",
  ]);
  let out = html
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<style[\s\S]*?<\/style>/gi, "");
  out = out.replace(/<\/?([a-zA-Z0-9]+)(?:\s[^>]*)?>/g, (match, tag: string) => {
    const t = String(tag).toLowerCase();
    if (!ALLOWED.has(t)) return "";
    if (match.startsWith("</")) return `</${t}>`;
    if (t === "br") return "<br>";
    return `<${t}>`;
  });
  return out.trim();
}

// ── App ─────────────────────────────────────────────────────────────────────
const app = express();
app.set("trust proxy", true);
app.use(express.json({ limit: "1mb" }));

// Tiny request log for Cloud Logging
app.use((req, _res, next) => {
  if (req.path.startsWith("/api/")) {
    console.log(`[${new Date().toISOString()}] ${req.method} ${req.path}`);
  }
  next();
});

// ── API: cards ──────────────────────────────────────────────────────────────
// GET /api/cards — everyone behind IAP can read (matches portal.html)
// Falls back to the bundled cards.json when no sheet is configured.
app.get("/api/cards", async (_req, res) => {
  try {
    res.json({ ok: true, cards: await readCards() });
  } catch (err: any) {
    console.error("[GET /api/cards]", err);
    res.status(500).json({ ok: false, error: err?.message || String(err) });
  }
});

// GET /api/me — for the admin UI to know who's logged in / whether they're admin
app.get("/api/me", (req, res) => {
  const email = getUserEmail(req);
  res.json({ ok: true, email, admin: isAdmin(email) });
});

// POST /api/cards — upsert (admin only). Matches updateCard() in Code.gs.
// Gated by HTTP Basic Auth so the admin UI's fetch() calls re-use the
// browser's stored credentials.
app.post("/api/cards", requireAdminAccess, async (req, res) => {
  try {
    const card = req.body as Card;
    if (!card || typeof card !== "object" || !card.id) {
      res.status(400).json({ ok: false, error: "Card must include an id" });
      return;
    }
    const cards = await readCards();
    const idx = cards.findIndex((c) => c.id === card.id);
    if (idx >= 0) cards[idx] = card;
    else cards.push(card);
    await writeCards(cards);
    res.json({ ok: true });
  } catch (err: any) {
    console.error("[POST /api/cards]", err);
    res.status(500).json({ ok: false, error: err?.message || String(err) });
  }
});

// DELETE /api/cards/:id — admin only. Matches deleteCard() in Code.gs.
app.delete("/api/cards/:id", requireAdminAccess, async (req, res) => {
  try {
    const id = req.params.id;
    const cards = (await readCards()).filter((c) => c.id !== id);
    await writeCards(cards);
    res.json({ ok: true });
  } catch (err: any) {
    console.error("[DELETE /api/cards/:id]", err);
    res.status(500).json({ ok: false, error: err?.message || String(err) });
  }
});

// POST /api/cards/reorder — admin only. Matches reorderCards() in Code.gs.
app.post("/api/cards/reorder", requireAdminAccess, async (req, res) => {
  try {
    const ids = req.body?.ids;
    if (!Array.isArray(ids)) {
      res.status(400).json({ ok: false, error: "Body must be { ids: string[] }" });
      return;
    }
    const cards = await readCards();
    const byId = new Map(cards.map((c) => [c.id, c] as const));
    const reordered = ids
      .map((id: unknown) => (typeof id === "string" ? byId.get(id) : undefined))
      .filter((c): c is Card => !!c);
    // Tack on any cards whose ids weren't in the reorder list to be safe
    for (const c of cards) if (!ids.includes(c.id)) reordered.push(c);
    await writeCards(reordered);
    res.json({ ok: true });
  } catch (err: any) {
    console.error("[POST /api/cards/reorder]", err);
    res.status(500).json({ ok: false, error: err?.message || String(err) });
  }
});

// ── API: SOX single-control analysis (called by Workato) ────────────────────
// POST /api/sox-analyze — admin-only endpoint.
// Body (JSON):
//   {
//     controlId:        string          // e.g. "SOX-042"
//     controlName:      string          // short display name
//     controlObjective: string          // what the control verifies
//     evidence: [                       // zero or more evidence files
//       { name: string, mimeType: string, data: string }  // data = base64
//     ]
//   }
// Response:
//   { ok: true, verdict: "Passed"|"Failed", gaps: string[], summary: string }

const GEMINI_PROJECT  = process.env.SOX_JOB_PROJECT  || "finance-ai-497313";
const GEMINI_LOCATION = process.env.SOX_JOB_REGION   || "us-east1";
const GEMINI_MODEL    = process.env.GEMINI_MODEL      || "gemini-2.0-flash-001";

// Dedicated model for the admin "Polish description" feature. Kept separate
// from GEMINI_MODEL so changing it never affects SOX analysis behavior.
const POLISH_MODEL = process.env.POLISH_MODEL || "gemini-2.5-pro";

const SOX_SYSTEM_INSTRUCTION = `You are a SOX (Sarbanes-Oxley) compliance auditor.
Your job is to review evidence files for an internal control and determine whether
it is operating effectively. Be precise and concise. Avoid vague statements.
Respond ONLY with a valid JSON object — no markdown, no prose outside the JSON.`;

const SOX_PROMPT_TEMPLATE = (
  controlId: string,
  controlName: string,
  controlObjective: string,
) => `CONTROL DETAILS
Control ID:        ${controlId}
Control Name:      ${controlName}
Control Objective: ${controlObjective || "(not specified)"}

INSTRUCTIONS
1. Review all attached evidence files carefully.
2. Determine whether the evidence shows the control is operating effectively.
3. If the control is failing, list specific, actionable gaps.

RESPONSE FORMAT (JSON only):
{
  "verdict": "Passed" or "Failed",
  "gaps": ["gap 1", "gap 2"],
  "summary": "One or two sentence summary"
}

If no evidence files were provided, respond with:
{"verdict":"Failed","gaps":["No evidence files provided"],"summary":"Unable to test — no evidence uploaded"}`;

app.post("/api/sox-analyze", requireAdminAccess, async (req, res) => {
  try {
    const { controlId = "", controlName = "", controlObjective = "", evidence = [] } =
      req.body as {
        controlId?: string;
        controlName?: string;
        controlObjective?: string;
        evidence?: Array<{ name: string; mimeType: string; data: string }>;
      };

    if (!controlId) {
      res.status(400).json({ ok: false, error: "controlId is required" });
      return;
    }

    const vertexai = new VertexAI({ project: GEMINI_PROJECT, location: GEMINI_LOCATION });
    const model = vertexai.getGenerativeModel({
      model: GEMINI_MODEL,
      systemInstruction: SOX_SYSTEM_INSTRUCTION,
    });

    const parts: import("@google-cloud/vertexai").Part[] = [
      { text: SOX_PROMPT_TEMPLATE(controlId, controlName, controlObjective) },
    ];

    for (const ef of evidence) {
      if (!ef.data) continue;
      parts.push({
        inlineData: {
          mimeType: ef.mimeType || "application/octet-stream",
          data: ef.data, // already base64 from Workato
        },
      });
    }

    const result = await model.generateContent({ contents: [{ role: "user", parts }] });
    let raw = result.response.candidates?.[0]?.content?.parts?.[0]?.text?.trim() ?? "";

    // Strip markdown code fences if model wraps output
    const fenceMatch = raw.match(/```(?:json)?\s*([\s\S]*?)```/);
    if (fenceMatch) raw = fenceMatch[1].trim();

    let parsed: { verdict: string; gaps: string[]; summary: string };
    try {
      parsed = JSON.parse(raw);
    } catch {
      console.error("[POST /api/sox-analyze] Gemini returned non-JSON:", raw.slice(0, 300));
      parsed = {
        verdict: "Failed",
        gaps: ["Agent error: could not parse model response"],
        summary: raw.slice(0, 200),
      };
    }

    res.json({ ok: true, ...parsed });
  } catch (err: any) {
    console.error("[POST /api/sox-analyze]", err);
    res.status(500).json({ ok: false, error: err?.message || String(err) });
  }
});

// ── API: Polish description (admin only) ────────────────────────────────────
// POST /api/polish-description — improves a card description with Gemini.
// Body (JSON):
//   { description: string (HTML), appUrl?: string }
// When appUrl is provided and reachable (public URL only), its page text is
// fetched and passed as reference context. Returns sanitized HTML.
//   { ok: true, polished: string, usedAppUrl: boolean }

const POLISH_SYSTEM_INSTRUCTION = `You are an expert technical writer for an internal "Finance AI & Automation Portal".
You rewrite short descriptions of finance automations and apps so they are clear, professional, concise and easy to scan.
Rules:
- Improve grammar, clarity, tone and structure. Keep it factual.
- Do NOT invent features, integrations, owners, or capabilities that are not present in the input.
- Prefer a short lead sentence, optionally followed by a compact bullet list of key points.
- Output ONLY HTML using these tags: <b>, <strong>, <i>, <em>, <code>, <ul>, <ol>, <li>, <br>, <p>.
- Do not include markdown, code fences, headings, links, images, styles, scripts, or any attributes.`;

app.post("/api/polish-description", requireAdminAccess, async (req, res) => {
  try {
    const { description = "", appUrl = "" } = req.body as {
      description?: string;
      appUrl?: string;
    };

    const plainDescription = htmlToText(String(description));
    if (!plainDescription) {
      res.status(400).json({ ok: false, error: "description is required" });
      return;
    }

    // Optionally fetch the App URL content for extra context.
    let referenceText = "";
    let usedAppUrl = false;
    const trimmedUrl = String(appUrl).trim();
    if (trimmedUrl && trimmedUrl !== "#") {
      const fetched = await fetchUrlSafely(trimmedUrl);
      if (fetched.ok && fetched.text) {
        referenceText = fetched.text;
        usedAppUrl = true;
      }
    }

    const promptParts = [
      "Rewrite and improve the following automation/app description for the portal.",
      "",
      "CURRENT DESCRIPTION (HTML):",
      String(description),
    ];
    if (referenceText) {
      promptParts.push(
        "",
        "REFERENCE CONTEXT (extracted text from the linked app page — use only to clarify, never to invent):",
        referenceText,
      );
    }
    promptParts.push(
      "",
      "Return the improved description as HTML only, following all the rules.",
    );

    const vertexai = new VertexAI({ project: GEMINI_PROJECT, location: GEMINI_LOCATION });
    const model = vertexai.getGenerativeModel({
      model: POLISH_MODEL,
      systemInstruction: POLISH_SYSTEM_INSTRUCTION,
    });

    const result = await model.generateContent({
      contents: [{ role: "user", parts: [{ text: promptParts.join("\n") }] }],
    });
    let raw = result.response.candidates?.[0]?.content?.parts?.[0]?.text?.trim() ?? "";

    // Strip markdown code fences if the model wraps output.
    const fenceMatch = raw.match(/```(?:html)?\s*([\s\S]*?)```/);
    if (fenceMatch) raw = fenceMatch[1].trim();

    const polished = sanitizeDescriptionHtml(raw);
    if (!polished) {
      res.status(502).json({ ok: false, error: "Model returned no usable content" });
      return;
    }

    res.json({ ok: true, polished, usedAppUrl });
  } catch (err: any) {
    console.error("[POST /api/polish-description]", err);
    res.status(500).json({ ok: false, error: err?.message || String(err) });
  }
});

// ── API: SOX agent trigger ───────────────────────────────────────────────────
// POST /api/sox-run — triggers the sox-agent Cloud Run Job (admin only).
// Returns immediately with the execution ID; the job runs asynchronously.
app.post("/api/sox-run", requireAdminAccess, async (req, res) => {
  if (!SOX_JOB_NAME) {
    res.status(503).json({ ok: false, error: "SOX_JOB_NAME not configured" });
    return;
  }

  // Build the gcloud command. We override env vars from the request body if
  // provided, so the admin can target a different sheet without redeploying.
  const overrides: string[] = [];
  const body = req.body as Record<string, string> | undefined;
  if (body?.spreadsheetId) overrides.push(`SOX_SPREADSHEET_ID=${body.spreadsheetId}`);
  if (body?.driveFolderId) overrides.push(`SOX_DRIVE_FOLDER_ID=${body.driveFolderId}`);
  if (body?.dryRun === "true") overrides.push("DRY_RUN=true");

  const args = [
    "run", "jobs", "execute", SOX_JOB_NAME,
    `--project=${SOX_JOB_PROJECT}`,
    `--region=${SOX_JOB_REGION}`,
    "--async",
    "--format=json",
  ];
  if (overrides.length) args.push(`--update-env-vars=${overrides.join(",")}`);

  execFile("gcloud", args, { timeout: 30_000 }, (err, stdout, stderr) => {
    if (err) {
      console.error("[POST /api/sox-run] gcloud error:", stderr || err.message);
      res.status(500).json({ ok: false, error: stderr || err.message });
      return;
    }
    try {
      const out = JSON.parse(stdout);
      const executionId = out?.metadata?.name || out?.name || stdout.trim();
      res.json({ ok: true, executionId });
    } catch {
      res.json({ ok: true, executionId: stdout.trim() });
    }
  });
});

// ── Static + page routes ────────────────────────────────────────────────────
// Resolve the static dir at runtime so it works both in dev (./public) and
// in the built bundle (./dist/public next to dist/server.cjs).
function resolveStaticDir(): string {
  const candidates = [
    path.resolve(process.cwd(), "dist", "public"), // production: node dist/server.cjs from repo root
    path.resolve(process.cwd(), "public"),         // dev: tsx server.ts from repo root
  ];
  for (const c of candidates) {
    if (fs.existsSync(path.join(c, "portal.html"))) return c;
  }
  return candidates[candidates.length - 1];
}
const STATIC_DIR = resolveStaticDir();

// Healthcheck (Cloud Run uses this implicitly)
app.get("/healthz", (_req, res) => res.json({ ok: true }));

// Admin gate runs BEFORE the static middleware. Covers both the pretty paths
// (/admin, /admin-form) and the raw filenames (/admin.html, /admin-form.html)
// so the gate can't be bypassed by guessing the file name.
//
// Two layers:
//   1. requireBasicAuth — browser-native username/password prompt. Anyone
//      who can't supply ADMIN_USERNAME/ADMIN_PASSWORD gets 401 and the
//      dialog re-prompts.
//   2. Optional email allowlist (ADMIN_EMAILS) — only kicks in when an
//      identity header is present (Cloud IAP mode). With Basic Auth alone
//      we just trust the shared secret.
const ADMIN_PATHS = [
  "/admin", "/admin/", "/admin.html",
  "/admin-form", "/admin-form/", "/admin-form.html",
];
app.get(ADMIN_PATHS, requireBasicAuth, (req, res, next) => {
  if (ADMIN_EMAILS.length === 0) {
    next();
    return;
  }
  const email = getUserEmail(req);
  if (!isAdmin(email)) {
    res.status(403).sendFile(path.join(STATIC_DIR, "access-denied.html"));
    return;
  }
  next();
});

// Explicit page routes for the pretty paths.
app.get(["/admin", "/admin/", "/admin.html"], (_req, res) =>
  res.sendFile(path.join(STATIC_DIR, "admin.html")),
);
app.get(["/admin-form", "/admin-form/", "/admin-form.html"], (_req, res) =>
  res.sendFile(path.join(STATIC_DIR, "admin-form.html")),
);

// Root → portal (also expose /portal.html for symmetry)
app.get(["/", "/portal.html"], (_req, res) =>
  res.sendFile(path.join(STATIC_DIR, "portal.html")),
);

// Static assets last — only serves files we haven't already routed above.
// `extensions` is intentionally NOT set.
app.use(express.static(STATIC_DIR));

// ── Boot ────────────────────────────────────────────────────────────────────
app.listen(PORT, "0.0.0.0", () => {
  console.log(
    `[startup] finance-portal listening on :${PORT} ` +
      `(auth=${AUTH_MODE}, sheet=${SHEET_ID ? "configured" : "MISSING"}, ` +
      `sheetPublic=${SHEET_PUBLIC}, admins=${ADMIN_EMAILS.length})`,
  );
});
