/**
 * LRS Analytics Dashboard — single-file build.
 *
 * The entire application (SQLite schema, access-code auth, spreadsheet
 * parser, import/records/analytics logic, Express routes, and the frontend
 * page itself) lives in this one file. Run with `node app.js` (or `npm
 * start`). Data still lives outside this file, in ./data/ (the SQLite
 * database, the access code, and files mid-upload) — that's runtime state,
 * not source code, so it isn't and shouldn't be embedded here.
 *
 * The frontend (originally public/index.html) is embedded below as a
 * base64 string and served verbatim on GET / — base64 avoids the need to
 * escape the backticks/${...} template-literal syntax the page's own inline
 * <script> uses extensively, which would otherwise conflict with wrapping
 * it in a JS template literal here.
 */

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { DatabaseSync } = require('node:sqlite'); // built into Node 22.5+/24 — no native build step required
const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const { parse: parseCsvSync } = require('csv-parse/sync');

/* ============================================================
   Paths / data directory — override with the DATA_DIR env var to point
   this at a mounted persistent volume (e.g. on Railway/Render/Fly.io,
   where the app's own directory is rebuilt fresh on every deploy and
   anything not on a volume is lost). Defaults to ./data for local use.
   ============================================================ */
const DATA_DIR = process.env.DATA_DIR ? path.resolve(process.env.DATA_DIR) : path.join(__dirname, 'data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
const UPLOAD_DIR = path.join(DATA_DIR, 'uploads');
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });

/* ============================================================
   Database + schema
   ============================================================ */
const DB_PATH = path.join(DATA_DIR, 'lrs.db');
const db = new DatabaseSync(DB_PATH);
db.exec('PRAGMA journal_mode = WAL');
db.exec('PRAGMA foreign_keys = ON');

db.exec(`
-- One row per active login: the whole app sits behind a single shared access
-- code (see getAccessCode/verifyAccessCode below), not per-user accounts, so
-- there is no users table and nothing here is attributed to a person.
CREATE TABLE IF NOT EXISTS access_sessions (
  token TEXT PRIMARY KEY,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  expires_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS uploads (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  filename TEXT NOT NULL,
  uploaded_at TEXT NOT NULL DEFAULT (datetime('now')),
  status TEXT NOT NULL,               -- success | partial | failed | cancelled
  total_rows INTEGER NOT NULL DEFAULT 0,
  imported_rows INTEGER NOT NULL DEFAULT 0,
  updated_rows INTEGER NOT NULL DEFAULT 0,
  skipped_rows INTEGER NOT NULL DEFAULT 0,
  error_count INTEGER NOT NULL DEFAULT 0,
  weeks_affected TEXT,                 -- JSON array of week_start dates
  action_summary TEXT,                 -- JSON: { hash: 'skip'|'update'|'create' } per duplicate encountered
  sheet_summary TEXT,                  -- JSON: [{ name, format, totalRows, validRows, errorRows }]
  notes TEXT
);

CREATE TABLE IF NOT EXISTS upload_errors (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  upload_id INTEGER NOT NULL REFERENCES uploads(id) ON DELETE CASCADE,
  row_number INTEGER,
  severity TEXT NOT NULL DEFAULT 'error', -- error | warning | skipped
  message TEXT NOT NULL,
  raw_snippet TEXT
);

CREATE TABLE IF NOT EXISTS posts (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  upload_id INTEGER REFERENCES uploads(id) ON DELETE SET NULL,
  -- Full-record fingerprint: every identifier field AND every metric for
  -- every platform on the row (see buildFullHash below). Two rows are only
  -- ever the "same" record if EVERY compared field matches exactly — a
  -- changed Reach/Views/etc. value produces a different hash, so a
  -- week-over-week analytics update is a new row, never silently dropped.
  source_row_hash TEXT NOT NULL UNIQUE,
  campaign_type TEXT,
  caption TEXT,
  content_type TEXT,
  publish_date TEXT NOT NULL,
  posting_time TEXT,
  week_start TEXT NOT NULL,
  month TEXT NOT NULL,               -- YYYY-MM
  quarter TEXT NOT NULL,             -- YYYY-Q#
  year INTEGER NOT NULL,
  platforms_raw TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_posts_week ON posts(week_start);
CREATE INDEX IF NOT EXISTS idx_posts_month ON posts(month);
CREATE INDEX IF NOT EXISTS idx_posts_quarter ON posts(quarter);
CREATE INDEX IF NOT EXISTS idx_posts_year ON posts(year);
CREATE INDEX IF NOT EXISTS idx_posts_date ON posts(publish_date);

CREATE TABLE IF NOT EXISTS post_metrics (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  post_id INTEGER NOT NULL REFERENCES posts(id) ON DELETE CASCADE,
  platform TEXT NOT NULL,
  posting_link TEXT,
  views INTEGER,
  reach INTEGER,
  impressions INTEGER,
  engagement INTEGER,
  clicks INTEGER,
  followers_gained INTEGER,
  watch_time_seconds REAL,
  shares INTEGER,
  comments INTEGER,
  saves INTEGER,
  UNIQUE(post_id, platform)
);

CREATE INDEX IF NOT EXISTS idx_metrics_platform ON post_metrics(platform);
CREATE INDEX IF NOT EXISTS idx_metrics_post ON post_metrics(post_id);

-- Verbatim copy of every non-blank source row, one row per sheet row, regardless
-- of whether it could be turned into a dashboard post. This is the "nothing from
-- the file is ever lost" guarantee — post_id is filled in when the row did become
-- (or update) a post; it stays NULL for rows that were unparseable or came from an
-- unrecognized sheet layout, but the original cell values are always kept.
CREATE TABLE IF NOT EXISTS raw_rows (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  upload_id INTEGER REFERENCES uploads(id) ON DELETE CASCADE,
  sheet_name TEXT,
  row_number INTEGER,
  post_id INTEGER REFERENCES posts(id) ON DELETE SET NULL,
  headers_json TEXT,     -- JSON array of column headers for this sheet, if any were detected
  raw_json TEXT NOT NULL, -- JSON array of every cell value in the row, original order, unmodified
  created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_raw_rows_upload ON raw_rows(upload_id);
`);

/* ============================================================
   Platform / metric config — single source of truth for platform
   metadata and header matching.
   ============================================================ */

// Colors are the validated 8-slot categorical palette (see dataviz skill /
// references/palette.md), assigned in FIXED order — never cycled, never
// swapped for a platform's real-world brand hue, since several brand colors
// (e.g. TikTok black, Threads gray) fail the chroma/lightness gate and two
// platforms' blues (Facebook, LinkedIn) fail CVD separation against each
// other. Identity comes from the label/legend, not from mimicking brand hue.
// Slots 7 (violet) and 8 (red) are held in reserve for future platforms.
const PLATFORMS = [
  { id: 'facebook', label: 'Facebook', color: '#2a78d6', darkColor: '#3987e5', groupAliases: ['facebook', 'fb'] },
  { id: 'instagram', label: 'Instagram', color: '#008300', darkColor: '#008300', groupAliases: ['instagram', 'ig'] },
  { id: 'tiktok', label: 'TikTok', color: '#e87ba4', darkColor: '#d55181', groupAliases: ['tiktok', 'tik tok'] },
  { id: 'linkedin', label: 'LinkedIn', color: '#eda100', darkColor: '#c98500', groupAliases: ['linked in', 'linkedin'] },
  { id: 'threads', label: 'Threads', color: '#1baf7a', darkColor: '#199e70', groupAliases: ['threads'] },
  { id: 'youtube', label: 'YouTube', color: '#eb6834', darkColor: '#d95926', groupAliases: ['youtube', 'yt'] },
];

const PLATFORM_IDS = PLATFORMS.map((p) => p.id);

// Canonical metrics tracked across the system. `agg` controls how weekly /
// period rollups combine per-post values. Extend this list to support a new
// metric everywhere (DB column, aggregation, API, and chart) in one place.
const CANONICAL_METRICS = [
  { key: 'views', label: 'Views', agg: 'sum' },
  { key: 'reach', label: 'Reach', agg: 'sum' },
  { key: 'impressions', label: 'Impressions', agg: 'sum' },
  { key: 'engagement', label: 'Engagement', agg: 'sum' },
  { key: 'clicks', label: 'Clicks', agg: 'sum' },
  { key: 'followers_gained', label: 'Followers Gained', agg: 'sum' },
  { key: 'watch_time_seconds', label: 'Watch Time', agg: 'sum', unit: 'seconds' },
  { key: 'shares', label: 'Shares', agg: 'sum' },
  { key: 'comments', label: 'Comments', agg: 'sum' },
  { key: 'saves', label: 'Saves', agg: 'sum' },
];

const CANONICAL_METRIC_KEYS = CANONICAL_METRICS.map((m) => m.key);

// Maps a normalized (lowercased, punctuation-stripped) source header label to
// a canonical metric key. Add new synonyms here when a platform export uses
// different wording for the same underlying metric.
const METRIC_SYNONYMS = {
  'views': 'views',
  'reach': 'reach',
  'impressions': 'impressions',
  'engagement': 'engagement',
  'engagements': 'engagement',
  'interactions': 'engagement',
  'clicks': 'clicks',
  'followers gained': 'followers_gained',
  'subscribers': 'followers_gained',
  'follower growth': 'followers_gained',
  'ave watch time': 'watch_time_raw', // needs unit-aware parsing (e.g. "2.96s")
  'average watch time': 'watch_time_raw',
  'watch time hours': 'watch_time_hours', // numeric hours
  'watch time': 'watch_time_raw',
  'duration': 'duration', // video length, informational only
  'shares': 'shares',
  'comments': 'comments',
  'saves': 'saves',
  'posting link': 'posting_link',
  'posting links': 'posting_link',
  'link': 'posting_link',
};

const IDENTIFIER_COLUMN_SYNONYMS = {
  'ads organic': 'campaign_type',
  'ads': 'campaign_type',
  'campaign': 'campaign_type',
  'post': 'caption',
  'caption': 'caption',
  'format': 'content_type',
  'content type': 'content_type',
  'publish date': 'publish_date',
  'date': 'publish_date',
  'posting time': 'posting_time',
  'time': 'posting_time',
  'platforms': 'platforms_raw',
  'platform': 'platform_single', // used by the "simple" per-row format
};

// Curated metric columns shown in the Data Records table when a specific
// platform is selected — a small, glanceable subset matching what that
// platform's block in the source sheet actually captures (see
// METRIC_SYNONYMS above). This only controls what the summary table shows;
// the full record — every imported field — is always available via the
// View/Edit popup regardless of this list.
const PLATFORM_RECORD_COLUMNS = {
  facebook: [
    { key: 'views', label: 'Views' },
    { key: 'reach', label: 'Reach' },
    { key: 'engagement', label: 'Engagement' },
    { key: 'posting_link', label: 'Link' },
  ],
  instagram: [
    { key: 'views', label: 'Views' },
    { key: 'reach', label: 'Reach' },
    { key: 'engagement', label: 'Interactions' },
    { key: 'posting_link', label: 'Link' },
  ],
  tiktok: [
    { key: 'views', label: 'Views' },
    { key: 'engagement', label: 'Engagements' },
    { key: 'followers_gained', label: 'Followers Gained' },
    { key: 'watch_time_seconds', label: 'Avg. Watch Time' },
    { key: 'posting_link', label: 'Link' },
  ],
  linkedin: [
    { key: 'impressions', label: 'Impressions' },
    { key: 'reach', label: 'Reach' },
    { key: 'posting_link', label: 'Link' },
  ],
  threads: [
    { key: 'views', label: 'Views' },
    { key: 'engagement', label: 'Interactions' },
    { key: 'posting_link', label: 'Link' },
  ],
  youtube: [
    { key: 'views', label: 'Views' },
    { key: 'watch_time_seconds', label: 'Watch Time' },
    { key: 'followers_gained', label: 'Subscribers' },
    { key: 'impressions', label: 'Impressions' },
    { key: 'posting_link', label: 'Link' },
  ],
};

function normalizeHeaderLabel(label) {
  return String(label || '')
    .toLowerCase()
    .replace(/[().]/g, '')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function findPlatformByGroupLabel(label) {
  const norm = normalizeHeaderLabel(label);
  if (!norm) return null;
  return PLATFORMS.find((p) => p.groupAliases.some((alias) => norm === normalizeHeaderLabel(alias)));
}

/* ============================================================
   Access code — the whole app sits behind one shared code, no accounts.
   ============================================================ */
const CODE_FILE = path.join(DATA_DIR, 'access-code.txt');
const DEFAULT_CODE = 'LRS2026';

function ensureCodeFile() {
  if (!fs.existsSync(CODE_FILE)) fs.writeFileSync(CODE_FILE, DEFAULT_CODE, 'utf8');
}

function getAccessCode() {
  ensureCodeFile();
  return fs.readFileSync(CODE_FILE, 'utf8').trim();
}

/** Fixed-length digest comparison so a wrong guess's length/content can't be timed out. */
function verifyAccessCode(candidate) {
  const actual = getAccessCode();
  const a = crypto.createHash('sha256').update(String(candidate || '')).digest();
  const b = crypto.createHash('sha256').update(actual).digest();
  return crypto.timingSafeEqual(a, b);
}

/* ============================================================
   Small utils
   ============================================================ */
function generateToken() {
  return crypto.randomBytes(32).toString('hex');
}

/** Stable fingerprint for a source row, used for dedup / merge matching. */
function stableHash(parts) {
  const normalized = parts
    .map((p) => String(p ?? '').trim().toLowerCase())
    .join('||');
  return crypto.createHash('sha256').update(normalized).digest('hex');
}

/**
 * Date / duration parsing helpers tolerant of the messy formats found in
 * real-world social media export spreadsheets (e.g. "January 1, 2026",
 * "3/26/2026", "2.96s", "1h:0m:44s").
 */
const MONTHS = [
  'january', 'february', 'march', 'april', 'may', 'june',
  'july', 'august', 'september', 'october', 'november', 'december',
];

function pad2(n) {
  return String(n).padStart(2, '0');
}

/** Parses a flexible date string/number into an ISO "YYYY-MM-DD" string, or null. */
function parseFlexibleDate(value) {
  if (value === null || value === undefined || value === '') return null;

  // Excel serial date numbers (xlsx can hand these back for date-formatted cells).
  if (typeof value === 'number' && Number.isFinite(value)) {
    const epoch = new Date(Date.UTC(1899, 11, 30));
    const d = new Date(epoch.getTime() + value * 86400000);
    if (!Number.isNaN(d.getTime())) {
      return `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
    }
    return null;
  }

  const raw = String(value).trim();
  if (!raw) return null;

  // "January 1, 2026" / "Jan 1 2026"
  const monthNameMatch = raw.match(/^([A-Za-z]+)\.?\s+(\d{1,2}),?\s+(\d{4})$/);
  if (monthNameMatch) {
    const monthIdx = MONTHS.findIndex((m) => m.startsWith(monthNameMatch[1].toLowerCase()));
    if (monthIdx >= 0) {
      return `${monthNameMatch[3]}-${pad2(monthIdx + 1)}-${pad2(Number(monthNameMatch[2]))}`;
    }
  }

  // "3/26/2026" or "3/26/26" (US month/day/year, matching the source sheet)
  const slashMatch = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (slashMatch) {
    let [, m, d, y] = slashMatch;
    if (y.length === 2) y = `20${y}`;
    return `${y}-${pad2(Number(m))}-${pad2(Number(d))}`;
  }

  // Already ISO
  const isoMatch = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) return raw;

  // Last resort: let the JS Date parser try.
  const parsed = new Date(raw);
  if (!Number.isNaN(parsed.getTime())) {
    return `${parsed.getFullYear()}-${pad2(parsed.getMonth() + 1)}-${pad2(parsed.getDate())}`;
  }

  return null;
}

/** Monday-start ISO week for a given "YYYY-MM-DD" date string. */
function isoWeekStart(isoDate) {
  const [y, m, d] = isoDate.split('-').map(Number);
  const date = new Date(Date.UTC(y, m - 1, d));
  const day = date.getUTCDay(); // 0 = Sunday
  const diffToMonday = day === 0 ? 6 : day - 1;
  date.setUTCDate(date.getUTCDate() - diffToMonday);
  return `${date.getUTCFullYear()}-${pad2(date.getUTCMonth() + 1)}-${pad2(date.getUTCDate())}`;
}

function quarterOf(month) {
  return Math.floor((month - 1) / 3) + 1;
}

/** Returns { year, month, quarter, weekStart } derived from an ISO date string. */
function deriveDateParts(isoDate) {
  const [y, m] = isoDate.split('-').map(Number);
  return {
    year: y,
    month: m,
    quarter: quarterOf(m),
    weekStart: isoWeekStart(isoDate),
  };
}

/**
 * Parses assorted watch-time representations into seconds.
 * Accepts: "2.96s", "9.02s", "1h:0m:44s", plain seconds numbers.
 */
function parseWatchTimeSeconds(value) {
  if (value === null || value === undefined || value === '') return null;
  const raw = String(value).trim();
  if (!raw || raw === '0' || /^0+\.?0*s?$/i.test(raw)) return raw === '0' ? 0 : 0;

  const hms = raw.match(/^(\d+)h:?(\d+)m:?(\d+)s$/i);
  if (hms) {
    const [, h, m, s] = hms.map(Number);
    return h * 3600 + m * 60 + s;
  }

  const secOnly = raw.match(/^([\d.]+)\s*s$/i);
  if (secOnly) return Number(secOnly[1]);

  const num = Number(raw.replace(/[^0-9.]/g, ''));
  return Number.isFinite(num) ? num : null;
}

function hoursToSeconds(value) {
  const num = Number(String(value).replace(/[^0-9.]/g, ''));
  return Number.isFinite(num) ? num * 3600 : null;
}

/**
 * Tolerant numeric coercion for messy analytics export cells: thousands
 * separators ("1,714"), shorthand suffixes ("6.4k"), and placeholder text
 * ("Reshare only", "no insights to show", "#REF!") which should quietly
 * become "no data" rather than blow up the import.
 */
function parseMetricNumber(value) {
  if (value === null || value === undefined) return null;
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;

  const raw = String(value).trim();
  if (!raw) return null;

  const suffixMatch = raw.match(/^([\d,.]+)\s*([km])$/i);
  if (suffixMatch) {
    const base = Number(suffixMatch[1].replace(/,/g, ''));
    if (!Number.isFinite(base)) return null;
    const mult = suffixMatch[2].toLowerCase() === 'k' ? 1_000 : 1_000_000;
    return base * mult;
  }

  const cleaned = raw.replace(/,/g, '');
  if (!/^-?\d+(\.\d+)?$/.test(cleaned)) return null; // placeholder text, "#REF!", etc.
  const num = Number(cleaned);
  return Number.isFinite(num) ? num : null;
}

/* ============================================================
   Access sessions (login/logout/session check)
   ============================================================ */
const SESSION_DAYS = 30;

function createSession() {
  const token = generateToken();
  const expiresAt = new Date(Date.now() + SESSION_DAYS * 86400000).toISOString();
  db.prepare('INSERT INTO access_sessions (token, expires_at) VALUES (?, ?)').run(token, expiresAt);
  return { token, expiresAt };
}

function login(code) {
  if (!verifyAccessCode(code)) throw new Error('Incorrect access code.');
  return createSession();
}

function logout(token) {
  if (token) db.prepare('DELETE FROM access_sessions WHERE token = ?').run(token);
}

function isValidSession(token) {
  if (!token) return false;
  const row = db.prepare('SELECT * FROM access_sessions WHERE token = ?').get(token);
  if (!row) return false;
  if (new Date(row.expires_at) < new Date()) {
    db.prepare('DELETE FROM access_sessions WHERE token = ?').run(token);
    return false;
  }
  return true;
}

/* ============================================================
   Auth middleware (cookie parsing + requireAuth guard)
   ============================================================ */
const SESSION_COOKIE = 'lrs_session';

function parseCookies(header) {
  const out = {};
  if (!header) return out;
  header.split(';').forEach((part) => {
    const idx = part.indexOf('=');
    if (idx === -1) return;
    const key = part.slice(0, idx).trim();
    const val = part.slice(idx + 1).trim();
    if (key) out[key] = decodeURIComponent(val);
  });
  return out;
}

/** Reads the session cookie (if any) and attaches req.authenticated / req.sessionToken. Never blocks. */
function attachSession(req, res, next) {
  req.cookies = parseCookies(req.headers.cookie);
  const token = req.cookies[SESSION_COOKIE];
  req.sessionToken = token || null;
  req.authenticated = token ? isValidSession(token) : false;
  next();
}

/** Blocks the request unless attachSession found a valid session. */
function requireAuth(req, res, next) {
  if (!req.authenticated) return res.status(401).json({ error: 'Access code required.' });
  next();
}

function setSessionCookie(res, token) {
  const maxAgeSeconds = 30 * 24 * 60 * 60;
  res.setHeader('Set-Cookie', `${SESSION_COOKIE}=${token}; HttpOnly; Path=/; Max-Age=${maxAgeSeconds}; SameSite=Lax`);
}

function clearSessionCookie(res) {
  res.setHeader('Set-Cookie', `${SESSION_COOKIE}=; HttpOnly; Path=/; Max-Age=0; SameSite=Lax`);
}

/* ============================================================
   Parser — CSV/Excel -> normalized rows (format auto-detect)
   ============================================================ */

/** Flattens an exceljs cell value (rich text, formula result, hyperlink, Date) into a plain scalar. */
function flattenCellValue(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return value; // parseFlexibleDate handles numbers, not Dates — convert below
  if (typeof value === 'object') {
    if (Array.isArray(value.richText)) return value.richText.map((t) => t.text).join('');
    if (value.text !== undefined) return value.text; // hyperlink { text, hyperlink }
    if (value.result !== undefined) return value.result; // formula cell
    if (value.error !== undefined) return `#${value.error}`;
    return '';
  }
  return value;
}

/**
 * Reads a .csv or .xlsx file into one or more sheets of array-of-arrays cell
 * data. A CSV always yields exactly one sheet; an .xlsx workbook yields one
 * entry per worksheet tab, in workbook order, so multi-tab exports are never
 * silently reduced to just the first tab.
 */
async function readSheets(filePath, originalName) {
  const ext = path.extname(originalName || filePath).toLowerCase();
  if (ext === '.csv' || ext === '.txt') {
    const content = fs.readFileSync(filePath, 'utf8');
    const rows = parseCsvSync(content, {
      relax_column_count: true,
      skip_empty_lines: false,
      bom: true,
    });
    return [{ sheetName: 'Sheet1', rows }];
  }

  // .xlsx / .xls — parsed with exceljs (avoids the unpatched prototype-pollution /
  // ReDoS advisories in the classic `xlsx` package when handling untrusted uploads).
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  return workbook.worksheets.map((worksheet) => {
    const rows = [];
    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const values = row.values; // 1-indexed sparse array; values[0] is always undefined
      const arr = [];
      for (let i = 1; i < values.length; i += 1) {
        const flat = flattenCellValue(values[i]);
        arr[i - 1] = flat instanceof Date
          ? `${flat.getFullYear()}-${String(flat.getMonth() + 1).padStart(2, '0')}-${String(flat.getDate()).padStart(2, '0')}`
          : flat;
      }
      rows[rowNumber - 1] = arr;
    });
    // Fill any fully-skipped row indices (exceljs skips truly empty rows even with includeEmpty in some edge cases).
    for (let i = 0; i < rows.length; i += 1) if (!rows[i]) rows[i] = [];
    return { sheetName: worksheet.name || 'Sheet1', rows };
  });
}

function isRowBlank(row) {
  return !row || row.every((cell) => cell === undefined || cell === null || String(cell).trim() === '');
}

/** Detects whether this is the "agenda tracker" wide format or a simple long-format table. */
function detectFormat(row0, row1) {
  const groupHits = (row0 || []).filter((cell) => findPlatformByGroupLabel(cell)).length;
  if (groupHits >= 2) return 'agenda';

  const headerRow = row0 || [];
  const hasPlatformColumn = headerRow.some((cell) => normalizeHeaderLabel(cell) === 'platform');
  if (hasPlatformColumn) return 'simple';

  return null;
}

/** Builds a per-column plan for the agenda (wide, multi-platform-block) format. */
function buildAgendaColumnPlan(groupRow, headerRow) {
  const plan = [];
  let currentPlatform = null;
  const width = Math.max(groupRow.length, headerRow.length);

  for (let i = 0; i < width; i += 1) {
    const groupCell = groupRow[i];
    if (groupCell !== undefined && String(groupCell).trim() !== '') {
      const matched = findPlatformByGroupLabel(groupCell);
      currentPlatform = matched ? matched.id : null;
    }

    const label = normalizeHeaderLabel(headerRow[i]);
    if (!label) {
      plan.push({ col: i, kind: 'ignore' });
      continue;
    }

    if (!currentPlatform) {
      const field = IDENTIFIER_COLUMN_SYNONYMS[label];
      plan.push(field ? { col: i, kind: 'identifier', field } : { col: i, kind: 'ignore' });
      continue;
    }

    const metric = METRIC_SYNONYMS[label];
    plan.push(metric ? { col: i, kind: 'metric', platform: currentPlatform, metric } : { col: i, kind: 'ignore' });
  }

  return plan;
}

/**
 * Builds a human-readable, per-column-unique header label for the agenda
 * format by qualifying each metric label with its platform group (e.g.
 * "FACEBOOK — Views", "INSTAGRAM — Views"). The source sheet reuses the same
 * label ("Views", "Reach", "Posting link", ...) under every platform block,
 * so the bare header row alone is not a safe display/edit column key.
 */
function buildQualifiedHeaders(groupRow, headerRow) {
  const headers = [];
  let currentGroupLabel = '';
  const width = Math.max(groupRow.length, headerRow.length);

  for (let i = 0; i < width; i += 1) {
    const groupCell = groupRow[i];
    if (groupCell !== undefined && String(groupCell).trim() !== '') {
      currentGroupLabel = String(groupCell).trim();
    }
    const label = headerRow[i] !== undefined && headerRow[i] !== null ? String(headerRow[i]).trim() : '';
    if (!label) {
      headers.push('');
      continue;
    }
    headers.push(currentGroupLabel ? `${currentGroupLabel} — ${label}` : label);
  }
  return headers;
}

function applyWatchTime(metricBag, key, rawValue) {
  if (key === 'watch_time_raw') {
    const seconds = parseWatchTimeSeconds(rawValue);
    if (seconds !== null) metricBag.watch_time_seconds = seconds;
  } else if (key === 'watch_time_hours') {
    const seconds = hoursToSeconds(rawValue);
    if (seconds !== null) metricBag.watch_time_seconds = seconds;
  }
}

const NON_METRIC_KEYS = new Set(['posting_link', 'watch_time_raw', 'watch_time_hours', 'duration']);

function emptyMetricBag() {
  return {
    views: null,
    reach: null,
    impressions: null,
    engagement: null,
    clicks: null,
    followers_gained: null,
    watch_time_seconds: null,
    shares: null,
    comments: null,
    saves: null,
    posting_link: null,
  };
}

function metricBagHasData(bag) {
  return Object.entries(bag).some(([k, v]) => k !== 'posting_link' && v !== null && v !== undefined);
}

// Fixed key order (matches emptyMetricBag()) so the fingerprint below is stable
// regardless of the order values were parsed in.
const METRIC_FINGERPRINT_KEYS = [
  'views', 'reach', 'impressions', 'engagement', 'clicks', 'followers_gained',
  'watch_time_seconds', 'shares', 'comments', 'saves', 'posting_link',
];

/**
 * A row is only ever a true duplicate of another if EVERY imported field
 * matches exactly — not just the date/caption/platform identity, but every
 * metric (Reach, Views, Engagement, etc.) for every platform on the row too.
 * This fingerprint covers the metrics half; combined with the identity
 * fields it becomes the row's full hash (see buildFullHash below). A post
 * whose numbers changed week over week — the normal case for a recurring
 * analytics export — is therefore NOT a duplicate and is preserved as its
 * own record rather than silently skipped or merged away.
 */
function fingerprintPlatforms(platforms) {
  return [...platforms]
    .sort((a, b) => a.platform.localeCompare(b.platform))
    .map((p) => `${p.platform}:${METRIC_FINGERPRINT_KEYS.map((k) => p[k]).join(',')}`)
    .join('|');
}

/**
 * Full-record fingerprint: identity fields (date/time/caption/content
 * type/platform list/campaign) PLUS every metric for every platform. Used as
 * `posts.source_row_hash` — the sole basis for "is this row a duplicate."
 */
function buildFullHash(identityHash, campaignType, platforms) {
  return stableHash([identityHash, campaignType, fingerprintPlatforms(platforms)]);
}

/**
 * Parses a single agenda-format row against a column plan. Returns
 * `{ post, error }` where exactly one is non-null, or both null for a row
 * with no identifiers and no platform data (treated as blank). This is the
 * shared core used both for bulk import and for re-interpreting a single
 * row after a manual edit on the Data Records page.
 */
function parseOneAgendaRow(row, plan, rowNumber) {
  const identifiers = {};
  const platformBags = {};

  plan.forEach((col) => {
    const raw = row[col.col];
    if (raw === undefined || raw === null || String(raw).trim() === '') return;

    if (col.kind === 'identifier') {
      identifiers[col.field] = raw;
    } else if (col.kind === 'metric') {
      const bag = (platformBags[col.platform] = platformBags[col.platform] || emptyMetricBag());
      if (col.metric === 'posting_link') {
        bag.posting_link = String(raw).trim();
      } else if (NON_METRIC_KEYS.has(col.metric)) {
        applyWatchTime(bag, col.metric, raw);
      } else {
        bag[col.metric] = parseMetricNumber(raw);
      }
    }
  });

  const hasAnyPlatformData = Object.values(platformBags).some(metricBagHasData);
  const hasAnyIdentifier = Object.keys(identifiers).length > 0;
  if (!hasAnyIdentifier && !hasAnyPlatformData) return { post: null, error: null };

  const isoDate = parseFlexibleDate(identifiers.publish_date);
  if (!isoDate) {
    return {
      post: null,
      error: {
        rowNumber,
        severity: 'skipped',
        message: identifiers.publish_date
          ? `Unrecognized date format: "${identifiers.publish_date}" — row skipped.`
          : 'Missing publish date — row skipped.',
        rawSnippet: (identifiers.caption || '').slice(0, 80),
      },
    };
  }

  const { year, month, quarter, weekStart } = deriveDateParts(isoDate);
  const campaignTypeRaw = identifiers.campaign_type ? String(identifiers.campaign_type).trim() : '';
  const campaignType = campaignTypeRaw && campaignTypeRaw.toLowerCase() !== 'choose' ? campaignTypeRaw : null;

  const platforms = Object.entries(platformBags)
    .filter(([, bag]) => metricBagHasData(bag))
    .map(([platform, bag]) => ({ platform, ...bag }));

  const identityHash = stableHash([
    isoDate,
    identifiers.posting_time,
    identifiers.caption,
    identifiers.content_type,
    identifiers.platforms_raw,
  ]);
  const hash = buildFullHash(identityHash, campaignType, platforms);

  const post = {
    rowNumber,
    hash,
    identityHash,
    campaignType,
    caption: identifiers.caption ? String(identifiers.caption).trim() : null,
    contentType: identifiers.content_type ? String(identifiers.content_type).trim() : null,
    publishDate: isoDate,
    postingTime: identifiers.posting_time ? String(identifiers.posting_time).trim() : null,
    weekStart,
    month: `${year}-${String(month).padStart(2, '0')}`,
    quarter: `${year}-Q${quarter}`,
    year,
    platformsRaw: identifiers.platforms_raw ? String(identifiers.platforms_raw).trim() : null,
    platforms,
  };
  return { post, error: null };
}

function parseAgendaRows(dataRows, plan, startRowNumber) {
  const posts = [];
  const errors = [];
  const rawRows = [];

  dataRows.forEach((row, idx) => {
    const rowNumber = startRowNumber + idx;
    if (isRowBlank(row)) return; // fully blank row — nothing to preserve, not stored
    rawRows.push({ rowNumber, raw: row });

    const { post, error } = parseOneAgendaRow(row, plan, rowNumber);
    if (error) {
      errors.push(error);
      return;
    }
    if (post) {
      rawRows[rawRows.length - 1].hash = post.hash;
      posts.push(post);
    }
  });

  return { posts, errors, rawRows };
}

function buildSimpleColumnPlan(headerRow) {
  return headerRow.map((cell, col) => {
    const label = normalizeHeaderLabel(cell);
    if (!label) return { col, kind: 'ignore' };
    if (IDENTIFIER_COLUMN_SYNONYMS[label]) return { col, kind: 'identifier', field: IDENTIFIER_COLUMN_SYNONYMS[label] };
    if (METRIC_SYNONYMS[label]) return { col, kind: 'metric', metric: METRIC_SYNONYMS[label] };
    return { col, kind: 'ignore' };
  });
}

/** Parses a single simple-format (one-platform-per-row) row against a column plan. */
function parseOneSimpleRow(row, plan, rowNumber) {
  const identifiers = {};
  const bag = emptyMetricBag();

  plan.forEach((col) => {
    const raw = row[col.col];
    if (raw === undefined || raw === null || String(raw).trim() === '') return;
    if (col.kind === 'identifier') {
      identifiers[col.field] = raw;
    } else if (col.kind === 'metric') {
      if (col.metric === 'posting_link') bag.posting_link = String(raw).trim();
      else if (NON_METRIC_KEYS.has(col.metric)) applyWatchTime(bag, col.metric, raw);
      else bag[col.metric] = parseMetricNumber(raw);
    }
  });

  if (!identifiers.platform_single) {
    return { post: null, error: { rowNumber, severity: 'skipped', message: 'Missing "platform" value — row skipped.' } };
  }
  const platformMatch = findPlatformByGroupLabel(identifiers.platform_single);
  if (!platformMatch) {
    return {
      post: null,
      error: { rowNumber, severity: 'skipped', message: `Unrecognized platform "${identifiers.platform_single}" — row skipped.` },
    };
  }

  const isoDate = parseFlexibleDate(identifiers.publish_date);
  if (!isoDate) {
    return {
      post: null,
      error: {
        rowNumber,
        severity: 'skipped',
        message: identifiers.publish_date
          ? `Unrecognized date format: "${identifiers.publish_date}" — row skipped.`
          : 'Missing publish date — row skipped.',
      },
    };
  }

  const { year, month, quarter, weekStart } = deriveDateParts(isoDate);
  const campaignTypeRaw = identifiers.campaign_type ? String(identifiers.campaign_type).trim() : '';
  const campaignType = campaignTypeRaw && campaignTypeRaw.toLowerCase() !== 'choose' ? campaignTypeRaw : null;

  const identityHash = stableHash([isoDate, identifiers.posting_time, identifiers.caption, identifiers.content_type, platformMatch.id]);
  const platforms = metricBagHasData(bag) ? [{ platform: platformMatch.id, ...bag }] : [];
  const hash = buildFullHash(identityHash, campaignType, platforms);

  const post = {
    rowNumber,
    hash,
    identityHash,
    campaignType,
    caption: identifiers.caption ? String(identifiers.caption).trim() : null,
    contentType: identifiers.content_type ? String(identifiers.content_type).trim() : null,
    publishDate: isoDate,
    postingTime: identifiers.posting_time ? String(identifiers.posting_time).trim() : null,
    weekStart,
    month: `${year}-${String(month).padStart(2, '0')}`,
    quarter: `${year}-Q${quarter}`,
    year,
    platformsRaw: platformMatch.id,
    platforms,
  };
  return { post, error: null };
}

function parseSimpleRows(dataRows, plan, startRowNumber) {
  const posts = [];
  const errors = [];
  const rawRows = [];

  dataRows.forEach((row, idx) => {
    const rowNumber = startRowNumber + idx;
    if (isRowBlank(row)) return;
    rawRows.push({ rowNumber, raw: row });

    const { post, error } = parseOneSimpleRow(row, plan, rowNumber);
    if (error) {
      errors.push(error);
      return;
    }
    if (post) {
      rawRows[rawRows.length - 1].hash = post.hash;
      posts.push(post);
    }
  });

  return { posts, errors, rawRows };
}

/** Every non-blank row of a sheet, captured verbatim, with no attempt to parse it. */
function captureRawOnly(rows, startRowNumber) {
  const rawRows = [];
  rows.forEach((row, idx) => {
    if (isRowBlank(row)) return;
    rawRows.push({ rowNumber: startRowNumber + idx, raw: row });
  });
  return rawRows;
}

/**
 * Parses a single sheet's rows, auto-detecting the agenda (wide, grouped
 * headers) layout vs. a simple one-row-per-platform table. A sheet whose
 * layout isn't recognized is never dropped — every non-blank row is still
 * captured verbatim (rawRows) and flagged with a warning, just without
 * becoming dashboard posts. Also returns the column `plan` (null for
 * unrecognized sheets) so a single row can be re-interpreted later, e.g.
 * after a manual edit on the Data Records page.
 */
function parseSheetRows(rows, sheetName) {
  if (!rows || rows.length < 2) {
    return {
      sheetName,
      format: null,
      posts: [],
      errors: [],
      rawRows: captureRawOnly(rows || [], 1),
      headers: null,
      plan: null,
    };
  }

  const format = detectFormat(rows[0], rows[1]);
  if (format === 'agenda') {
    const plan = buildAgendaColumnPlan(rows[0], rows[1]);
    const headers = buildQualifiedHeaders(rows[0], rows[1]);
    const { posts, errors, rawRows } = parseAgendaRows(rows.slice(2), plan, 3);
    return { sheetName, format, posts, errors, rawRows, headers, plan };
  }
  if (format === 'simple') {
    const plan = buildSimpleColumnPlan(rows[0]);
    const headers = rows[0].map((h) => (h !== undefined && h !== null ? String(h).trim() : ''));
    const { posts, errors, rawRows } = parseSimpleRows(rows.slice(1), plan, 2);
    return { sheetName, format, posts, errors, rawRows, headers, plan };
  }

  const rawRows = captureRawOnly(rows, 1);
  const errors = rawRows.length
    ? [
        {
          rowNumber: null,
          severity: 'warning',
          message: `Sheet "${sheetName}": layout not recognized (no platform-group header or "platform" column found). ${rawRows.length} row(s) were saved as raw data only and are not included in dashboard totals.`,
        },
      ]
    : [];
  const headers = (rows[0] || []).map((h) => (h !== undefined && h !== null ? String(h).trim() : ''));
  return { sheetName, format: null, posts: [], errors, rawRows, headers, plan: null };
}

/**
 * Parses an uploaded analytics file into normalized post + platform-metric
 * records, across every sheet in the workbook (a CSV always has exactly one).
 * Returns one result per sheet; nothing from any sheet is skipped — rows that
 * can't become a dashboard post are still returned in that sheet's rawRows.
 */
async function parseAnalyticsFile(filePath, originalName) {
  const sheets = await readSheets(filePath, originalName);
  if (!sheets.length) {
    throw new Error('File appears to be empty.');
  }

  const results = sheets.map(({ sheetName, rows }) => parseSheetRows(rows, sheetName));

  const anyRecognized = results.some((r) => r.format !== null);
  const anyRows = results.some((r) => r.rawRows.length > 0);
  if (!anyRecognized && !anyRows) {
    throw new Error(
      'Unrecognized file layout. Expected either the LRS agenda tracker (grouped platform headers) or a simple table with a "platform" column.'
    );
  }

  return { sheets: results };
}

/* ============================================================
   Import service — dedup, per-record Skip/Update/Create, transactional writes
   ============================================================ */

/** Splits parsed posts into unique-by-hash posts and duplicate-within-file notices (literal copy/paste repeats in the same file). */
function dedupeWithinFile(posts) {
  const seen = new Map();
  const duplicates = [];
  for (const post of posts) {
    if (seen.has(post.hash)) {
      duplicates.push({
        rowNumber: post.rowNumber,
        sheetName: post.sheetName,
        severity: 'skipped',
        message: `Duplicate of row ${seen.get(post.hash)} within this file — skipped.`,
        rawSnippet: (post.caption || '').slice(0, 80),
      });
    } else {
      seen.set(post.hash, post.rowNumber);
    }
  }
  const unique = [...new Map(posts.map((p) => [p.hash, p])).values()];
  return { unique, duplicates };
}

/**
 * Flattens the per-sheet parse results (parseAnalyticsFile returns one entry
 * per worksheet) into single posts/errors/rawRows lists, tagging every entry
 * with which sheet it came from, plus a compact per-sheet summary for the
 * upload record.
 */
function flattenSheets(sheets) {
  const posts = [];
  const errors = [];
  const rawRows = [];
  const sheetSummary = [];

  for (const s of sheets) {
    posts.push(...s.posts.map((p) => ({ ...p, sheetName: s.sheetName })));
    errors.push(...s.errors.map((e) => ({ ...e, sheetName: s.sheetName })));
    rawRows.push(...s.rawRows.map((r) => ({ ...r, sheetName: s.sheetName, headers: s.headers, plan: s.plan, format: s.format })));
    sheetSummary.push({
      name: s.sheetName,
      format: s.format,
      totalRows: s.rawRows.length,
      validRows: s.posts.length,
      errorRows: s.errors.filter((e) => e.severity !== 'warning').length,
    });
  }

  return { posts, errors, rawRows, sheetSummary };
}

function prefixWithSheet(issue, multiSheet) {
  if (!multiSheet || !issue.sheetName) return issue;
  return { ...issue, message: `[Sheet: ${issue.sheetName}] ${issue.message}` };
}

// Shared across import + records editing — same SQL, no need for two prepared copies.
const findPostBySourceHashStmt = db.prepare(`
  SELECT id, publish_date, caption, campaign_type, content_type, updated_at FROM posts WHERE source_row_hash = ?
`);
const sharedDeleteMetricsForPostStmt = db.prepare('DELETE FROM post_metrics WHERE post_id = ?');
const sharedInsertMetricStmt = db.prepare(`
  INSERT INTO post_metrics (post_id, platform, posting_link, views, reach, impressions, engagement,
                             clicks, followers_gained, watch_time_seconds, shares, comments, saves)
  VALUES (@postId, @platform, @posting_link, @views, @reach, @impressions, @engagement,
          @clicks, @followers_gained, @watch_time_seconds, @shares, @comments, @saves)
`);

/**
 * Splits already-deduped posts into brand-new records vs. ones matching an
 * existing DB post. A match requires every compared field — identifiers AND
 * every platform's metrics — to be identical (see buildFullHash above); a
 * row that shares the same date/caption/platform as an existing post but has
 * even one different metric value is NOT a match and is classified as a new
 * record, so week-over-week analytics changes are never lost.
 */
function classifyAgainstExisting(posts) {
  const newPosts = [];
  const duplicates = [];
  for (const post of posts) {
    const existing = findPostBySourceHashStmt.get(post.hash);
    if (existing) duplicates.push({ post, existing });
    else newPosts.push(post);
  }
  return { newPosts, duplicates };
}

/** Dry-run: parses the file and reports what would happen, without writing anything. */
async function previewImport(filePath, originalName) {
  const { sheets } = await parseAnalyticsFile(filePath, originalName);
  const { posts, errors, rawRows, sheetSummary } = flattenSheets(sheets);
  const { unique, duplicates: dupWithinFile } = dedupeWithinFile(posts);
  const multiSheet = sheets.length > 1;
  const { newPosts, duplicates } = classifyAgainstExisting(unique);

  const nonWarningErrors = errors.filter((e) => e.severity !== 'warning');

  return {
    sheets: sheetSummary,
    totalDataRows: rawRows.length,
    validRows: unique.length,
    duplicateRowsInFile: dupWithinFile.length,
    errorRows: nonWarningErrors.length,
    newRecordsCount: newPosts.length,
    duplicates: duplicates.map(({ post, existing }) => ({
      hash: post.hash,
      rowNumber: post.rowNumber,
      sheetName: post.sheetName,
      publishDate: post.publishDate,
      caption: (post.caption || '').slice(0, 100),
      contentType: post.contentType,
      platforms: post.platforms.map((p) => p.platform),
      existing: {
        postId: existing.id,
        publishDate: existing.publish_date,
        caption: (existing.caption || '').slice(0, 100),
        campaignType: existing.campaign_type,
        contentType: existing.content_type,
        updatedAt: existing.updated_at,
      },
    })),
    issues: [...errors, ...dupWithinFile]
      .map((i) => prefixWithSheet(i, multiSheet))
      .sort((a, b) => (a.rowNumber || 0) - (b.rowNumber || 0)),
    sample: unique.slice(0, 5).map((p) => ({
      publishDate: p.publishDate,
      caption: (p.caption || '').slice(0, 100),
      contentType: p.contentType,
      campaignType: p.campaignType,
      platforms: p.platforms.map((pl) => pl.platform),
      sheetName: p.sheetName,
    })),
  };
}

const importInsertPostStmt = db.prepare(`
  INSERT INTO posts (upload_id, source_row_hash, campaign_type, caption, content_type, publish_date,
                      posting_time, week_start, month, quarter, year, platforms_raw, updated_at)
  VALUES (@uploadId, @hash, @campaignType, @caption, @contentType, @publishDate,
          @postingTime, @weekStart, @month, @quarter, @year, @platformsRaw, datetime('now'))
`);

const importUpdatePostStmt = db.prepare(`
  UPDATE posts SET upload_id=@uploadId, campaign_type=@campaignType, caption=@caption,
    content_type=@contentType, publish_date=@publishDate, posting_time=@postingTime,
    week_start=@weekStart, month=@month, quarter=@quarter, year=@year,
    platforms_raw=@platformsRaw, updated_at=datetime('now')
  WHERE id=@id
`);

const insertRawRowStmt = db.prepare(`
  INSERT INTO raw_rows (upload_id, sheet_name, row_number, post_id, headers_json, raw_json)
  VALUES (@uploadId, @sheetName, @rowNumber, @postId, @headersJson, @rawJson)
`);

/** Writes (insert or update) a post + its platform metrics. Returns the post id. */
function writePostAndMetrics(post, uploadId, { isUpdate, existingId }) {
  const params = {
    uploadId,
    campaignType: post.campaignType,
    caption: post.caption,
    contentType: post.contentType,
    publishDate: post.publishDate,
    postingTime: post.postingTime,
    weekStart: post.weekStart,
    month: post.month,
    quarter: post.quarter,
    year: post.year,
    platformsRaw: post.platformsRaw,
  };

  let postId;
  if (isUpdate) {
    importUpdatePostStmt.run({ ...params, id: existingId });
    sharedDeleteMetricsForPostStmt.run(existingId);
    postId = existingId;
  } else {
    const info = importInsertPostStmt.run({ ...params, hash: post.hash });
    postId = Number(info.lastInsertRowid);
  }

  for (const m of post.platforms) {
    sharedInsertMetricStmt.run({
      postId,
      platform: m.platform,
      posting_link: m.posting_link ?? null,
      views: m.views ?? null,
      reach: m.reach ?? null,
      impressions: m.impressions ?? null,
      engagement: m.engagement ?? null,
      clicks: m.clicks ?? null,
      followers_gained: m.followers_gained ?? null,
      watch_time_seconds: m.watch_time_seconds ?? null,
      shares: m.shares ?? null,
      comments: m.comments ?? null,
      saves: m.saves ?? null,
    });
  }

  return postId;
}

const insertUploadStmt = db.prepare(`
  INSERT INTO uploads (filename, status, total_rows, imported_rows, updated_rows, skipped_rows,
                        error_count, weeks_affected, action_summary, sheet_summary, notes)
  VALUES (@filename, @status, @totalRows, @importedRows, @updatedRows, @skippedRows,
          @errorCount, @weeksAffected, @actionSummary, @sheetSummary, @notes)
`);

const insertUploadErrorStmt = db.prepare(`
  INSERT INTO upload_errors (upload_id, row_number, severity, message, raw_snippet)
  VALUES (@uploadId, @rowNumber, @severity, @message, @rawSnippet)
`);

/**
 * Commits an upload: re-parses the file (every sheet), appends every
 * brand-new record automatically, and resolves each record that's an EXACT
 * duplicate of an existing post — every identifier and every metric
 * matching, not just date/caption/platform — according to `duplicateActions`
 * (keyed by hash) or `defaultDuplicateAction` for any duplicate not given an
 * explicit choice. A row with the same date/caption/platform as an existing
 * post but different metrics is not a duplicate at all; it's appended as its
 * own new record automatically, same as any other brand-new row:
 *   - 'skip'   — leave the existing record untouched
 *   - 'update' — overwrite the existing record with this upload's values
 *   - 'create' — keep the existing record AND add this row as a new, separate record
 *
 * Every non-blank source row — across every sheet, whatever was decided — is
 * also preserved verbatim in `raw_rows`, linked to whichever post it ended
 * up mapped to.
 */
async function commitImport(filePath, originalName, { defaultDuplicateAction = 'skip', duplicateActions = {}, notes = null } = {}) {
  const { sheets } = await parseAnalyticsFile(filePath, originalName);
  const { posts, errors, rawRows, sheetSummary } = flattenSheets(sheets);
  const { unique, duplicates: dupWithinFile } = dedupeWithinFile(posts);
  const multiSheet = sheets.length > 1;
  const issues = [...errors, ...dupWithinFile].map((i) => prefixWithSheet(i, multiSheet));

  let importedRows = 0; // brand-new records appended
  let updatedRows = 0; // duplicates resolved as "update existing"
  let skippedRows = 0; // duplicates resolved as "skip"
  const weeksAffected = new Set();
  const actionSummary = {};
  const hashToWrittenPostId = new Map();

  const run = () => {
    for (const post of unique) {
      const existing = findPostBySourceHashStmt.get(post.hash);

      if (!existing) {
        const id = writePostAndMetrics(post, null, { isUpdate: false });
        hashToWrittenPostId.set(post.hash, id);
        importedRows += 1;
        weeksAffected.add(post.weekStart);
        continue;
      }

      const action = duplicateActions[post.hash] || defaultDuplicateAction;
      actionSummary[post.hash] = action;

      if (action === 'update') {
        writePostAndMetrics(post, null, { isUpdate: true, existingId: existing.id });
        hashToWrittenPostId.set(post.hash, existing.id);
        updatedRows += 1;
        weeksAffected.add(post.weekStart);
      } else if (action === 'create') {
        const disambiguatedHash = `${post.hash}#dup-${Date.now()}-${Math.floor(Math.random() * 1e6)}`;
        const id = writePostAndMetrics({ ...post, hash: disambiguatedHash }, null, { isUpdate: false });
        hashToWrittenPostId.set(post.hash, id);
        importedRows += 1;
        weeksAffected.add(post.weekStart);
      } else {
        // 'skip' (default): leave the existing record exactly as it is.
        hashToWrittenPostId.set(post.hash, existing.id);
        skippedRows += 1;
        issues.push({
          rowNumber: post.rowNumber,
          severity: 'skipped',
          message: `Exact duplicate of existing record #${existing.id} — every field (including all metrics) matches, so nothing changed — skipped.`,
        });
      }
    }

    const errorCount = issues.filter((i) => i.severity === 'error').length;
    const status = importedRows + updatedRows === 0 ? 'failed' : issues.length > 0 ? 'partial' : 'success';
    const weeksArray = [...weeksAffected].sort();

    const uploadInfo = insertUploadStmt.run({
      filename: originalName,
      status,
      totalRows: rawRows.length,
      importedRows,
      updatedRows,
      skippedRows,
      errorCount,
      weeksAffected: JSON.stringify(weeksArray),
      actionSummary: JSON.stringify(actionSummary),
      sheetSummary: JSON.stringify(sheetSummary),
      notes: notes || null,
    });
    const uploadId = Number(uploadInfo.lastInsertRowid);

    // Backfill upload_id on newly written/updated posts (best-effort attribution to this batch).
    for (const weekStart of weeksArray) {
      db.prepare("UPDATE posts SET upload_id = ? WHERE week_start = ? AND upload_id IS NULL").run(uploadId, weekStart);
    }

    for (const issue of issues) {
      insertUploadErrorStmt.run({
        uploadId,
        rowNumber: issue.rowNumber ?? null,
        severity: issue.severity || 'error',
        message: issue.message,
        rawSnippet: issue.rawSnippet || null,
      });
    }

    // Preserve every non-blank source row verbatim, linked to whichever post it
    // ended up mapped to (brand-new, updated-existing, skip-target, or a fresh
    // "create" duplicate) — resolved via hashToWrittenPostId built up above.
    for (const r of rawRows) {
      insertRawRowStmt.run({
        uploadId,
        sheetName: r.sheetName,
        rowNumber: r.rowNumber,
        postId: r.hash ? (hashToWrittenPostId.get(r.hash) ?? null) : null,
        headersJson: JSON.stringify({ headers: r.headers || null, plan: r.plan || null, format: r.format || null }),
        rawJson: JSON.stringify(r.raw),
      });
    }

    return uploadId;
  };

  let uploadId;
  db.exec('BEGIN');
  try {
    uploadId = run();
    db.exec('COMMIT');
  } catch (err) {
    db.exec('ROLLBACK');
    throw err;
  }

  return {
    uploadId,
    sheets: sheetSummary,
    importedRows,
    updatedRows,
    skippedRows,
    errorCount: issues.length,
    weeksAffected: [...weeksAffected].sort(),
    actionSummary,
  };
}

/* ============================================================
   Records service — Data Records read/edit/delete (raw mirror +
   the platform-grouped CRM table), kept in sync on every save.
   ============================================================ */

function parseMeta(headersJson) {
  if (!headersJson) return { headers: null, plan: null, format: null };
  try {
    return JSON.parse(headersJson);
  } catch {
    return { headers: null, plan: null, format: null };
  }
}

function rowToRecord(row) {
  const meta = parseMeta(row.headers_json);
  return {
    id: row.id,
    uploadId: row.upload_id,
    sheetName: row.sheet_name,
    rowNumber: row.row_number,
    postId: row.post_id,
    format: meta.format,
    headers: meta.headers,
    values: JSON.parse(row.raw_json),
  };
}

/** Builds an optional filter, joining to `posts`/`post_metrics` only when a filter needs it. */
function buildFilteredQuery({ dateFrom, dateTo, platform, campaignType, contentType }) {
  const clauses = [];
  const params = [];
  const needsJoin = Boolean(
    dateFrom || dateTo || (platform && platform !== 'all') || (campaignType && campaignType !== 'all') || (contentType && contentType !== 'all')
  );

  if (needsJoin) {
    if (dateFrom) {
      clauses.push('p.publish_date >= ?');
      params.push(dateFrom);
    }
    if (dateTo) {
      clauses.push('p.publish_date <= ?');
      params.push(dateTo);
    }
    if (campaignType && campaignType !== 'all') {
      clauses.push('p.campaign_type = ?');
      params.push(campaignType);
    }
    if (contentType && contentType !== 'all') {
      clauses.push('p.content_type LIKE ?');
      params.push(`%${contentType}%`);
    }
    if (platform && platform !== 'all') {
      clauses.push('rr.post_id IN (SELECT post_id FROM post_metrics WHERE platform = ?)');
      params.push(platform);
    }
  }

  const from = needsJoin ? 'raw_rows rr JOIN posts p ON p.id = rr.post_id' : 'raw_rows rr';
  const where = clauses.length ? `WHERE ${clauses.join(' AND ')}` : '';
  return { from, where, params };
}

function listRecords({ dateFrom, dateTo, platform, campaignType, contentType, page = 1, pageSize = 25 }) {
  const { from, where, params } = buildFilteredQuery({ dateFrom, dateTo, platform, campaignType, contentType });
  const total = db.prepare(`SELECT COUNT(*) AS cnt FROM ${from} ${where}`).get(...params).cnt;
  const safePage = Math.max(1, page);
  const safePageSize = Math.min(200, Math.max(1, pageSize));
  const offset = (safePage - 1) * safePageSize;

  const rows = db
    .prepare(`SELECT rr.* FROM ${from} ${where} ORDER BY rr.sheet_name ASC, rr.row_number ASC LIMIT ? OFFSET ?`)
    .all(...params, safePageSize, offset);

  return { rows: rows.map(rowToRecord), total, page: safePage, pageSize: safePageSize };
}

function getRecord(id) {
  const row = db.prepare('SELECT * FROM raw_rows WHERE id = ?').get(id);
  return row ? rowToRecord(row) : null;
}

const updatePostFieldsStmt = db.prepare(`
  UPDATE posts SET source_row_hash=@hash, campaign_type=@campaignType, caption=@caption, content_type=@contentType,
    publish_date=@publishDate, posting_time=@postingTime, week_start=@weekStart, month=@month,
    quarter=@quarter, year=@year, platforms_raw=@platformsRaw, updated_at=datetime('now')
  WHERE id=@id
`);
const recordsInsertPostStmt = db.prepare(`
  INSERT INTO posts (source_row_hash, campaign_type, caption, content_type, publish_date, posting_time,
                      week_start, month, quarter, year, platforms_raw, updated_at)
  VALUES (@hash, @campaignType, @caption, @contentType, @publishDate, @postingTime,
          @weekStart, @month, @quarter, @year, @platformsRaw, datetime('now'))
`);
const updateRawJsonStmt = db.prepare('UPDATE raw_rows SET raw_json = ? WHERE id = ?');
const updateRawPostIdStmt = db.prepare('UPDATE raw_rows SET post_id = ? WHERE id = ?');

/** Writes (insert or update) the post + its platform metrics implied by a re-parsed row. */
function writePostFromParsedRow(post, existingPostId) {
  let targetId = existingPostId || null;
  if (!targetId) {
    const found = findPostBySourceHashStmt.get(post.hash);
    if (found) targetId = found.id;
  }

  const params = {
    hash: post.hash,
    campaignType: post.campaignType,
    caption: post.caption,
    contentType: post.contentType,
    publishDate: post.publishDate,
    postingTime: post.postingTime,
    weekStart: post.weekStart,
    month: post.month,
    quarter: post.quarter,
    year: post.year,
    platformsRaw: post.platformsRaw,
  };

  if (targetId) {
    updatePostFieldsStmt.run({ ...params, id: targetId });
    sharedDeleteMetricsForPostStmt.run(targetId);
  } else {
    const info = recordsInsertPostStmt.run(params);
    targetId = Number(info.lastInsertRowid);
  }

  for (const m of post.platforms) {
    sharedInsertMetricStmt.run({
      postId: targetId,
      platform: m.platform,
      posting_link: m.posting_link ?? null,
      views: m.views ?? null,
      reach: m.reach ?? null,
      impressions: m.impressions ?? null,
      engagement: m.engagement ?? null,
      clicks: m.clicks ?? null,
      followers_gained: m.followers_gained ?? null,
      watch_time_seconds: m.watch_time_seconds ?? null,
      shares: m.shares ?? null,
      comments: m.comments ?? null,
      saves: m.saves ?? null,
    });
  }

  return targetId;
}

/**
 * Saves an edit to one spreadsheet row. `values` is the full array of cell
 * values for that row, in the same column order it was imported with. The
 * literal values are always persisted verbatim (so the record keeps
 * mirroring exactly what's "in the sheet" per the user's edit); if the row's
 * sheet had a recognized layout, it's also re-run through that layout's
 * parser so the dashboard/comparisons/reports reflect the edit immediately.
 */
function updateRecord(id, values) {
  const row = db.prepare('SELECT * FROM raw_rows WHERE id = ?').get(id);
  if (!row) throw new Error('Record not found.');
  if (!Array.isArray(values)) throw new Error('Invalid payload: expected an array of column values.');

  const meta = parseMeta(row.headers_json);

  if (!meta.plan || !meta.format) {
    // Row came from a sheet layout we don't parse (or never had one) — just
    // persist the literal edit; there is no dashboard record to keep in sync.
    updateRawJsonStmt.run(JSON.stringify(values), id);
    return getRecord(id);
  }

  const parsed = meta.format === 'agenda'
    ? parseOneAgendaRow(values, meta.plan, row.row_number)
    : parseOneSimpleRow(values, meta.plan, row.row_number);

  if (parsed.error) {
    throw new Error(parsed.error.message.replace(/\s*—\s*row skipped\.?$/i, '.'));
  }
  if (!parsed.post) {
    throw new Error('This row has no identifying information left (date, caption, or platform data). Fill in at least the publish date to save.');
  }

  db.exec('BEGIN');
  try {
    updateRawJsonStmt.run(JSON.stringify(values), id);
    const postId = writePostFromParsedRow(parsed.post, row.post_id);
    if (postId !== row.post_id) updateRawPostIdStmt.run(postId, id);
    db.exec('COMMIT');
  } catch (err) {
    db.exec('ROLLBACK');
    throw err;
  }

  return getRecord(id);
}

/**
 * The CRM-style Data Records table (grouped by platform) — a curated,
 * glanceable view backed by `posts`/`post_metrics` (the same normalized data
 * every dashboard/comparison/report reads), NOT the raw dynamic-column
 * mirror above. Every field of a record remains available regardless of
 * this curation via View/Edit, which reads the full raw_rows mirror.
 */
function buildPostWhereClause({ dateFrom, dateTo, campaignType, contentType, search }) {
  const clauses = [];
  const params = [];
  if (dateFrom) {
    clauses.push('p.publish_date >= ?');
    params.push(dateFrom);
  }
  if (dateTo) {
    clauses.push('p.publish_date <= ?');
    params.push(dateTo);
  }
  if (campaignType && campaignType !== 'all') {
    clauses.push('p.campaign_type = ?');
    params.push(campaignType);
  }
  if (contentType && contentType !== 'all') {
    clauses.push('p.content_type LIKE ?');
    params.push(`%${contentType}%`);
  }
  if (search && search.trim()) {
    clauses.push('(p.caption LIKE ? OR p.platforms_raw LIKE ? OR p.campaign_type LIKE ? OR p.content_type LIKE ?)');
    const like = `%${search.trim()}%`;
    params.push(like, like, like, like);
  }
  return { clauses, params };
}

function paginate(page, pageSize) {
  const safePage = Math.max(1, Number(page) || 1);
  const safePageSize = Math.min(200, Math.max(1, Number(pageSize) || 25));
  return { safePage, safePageSize, offset: (safePage - 1) * safePageSize };
}

function rowStatus(createdAt, updatedAt) {
  return createdAt === updatedAt ? 'original' : 'edited';
}

/** "All Platforms" view: one row per post, common fields only, spanning every platform it touched. */
function listRecordsSummary(filters, page, pageSize) {
  const { clauses, params } = buildPostWhereClause(filters);
  const where = clauses.length ? `WHERE ${clauses.join(' AND ')}` : '';
  const total = db.prepare(`SELECT COUNT(*) AS cnt FROM posts p ${where}`).get(...params).cnt;
  const { safePage, safePageSize, offset } = paginate(page, pageSize);

  const rows = db
    .prepare(`
      SELECT p.id AS post_id, p.publish_date, p.caption, p.campaign_type, p.content_type,
             p.platforms_raw, p.created_at, p.updated_at,
             (SELECT MAX(rr.id) FROM raw_rows rr WHERE rr.post_id = p.id) AS raw_row_id,
             (SELECT GROUP_CONCAT(DISTINCT pm.platform) FROM post_metrics pm WHERE pm.post_id = p.id) AS platform_ids
      FROM posts p
      ${where}
      ORDER BY p.publish_date DESC, p.id DESC
      LIMIT ? OFFSET ?
    `)
    .all(...params, safePageSize, offset);

  return {
    platform: 'all',
    rows: rows.map((r) => ({
      postId: r.post_id,
      rawRowId: r.raw_row_id,
      publishDate: r.publish_date,
      caption: r.caption,
      campaignType: r.campaign_type,
      contentType: r.content_type,
      platformIds: r.platform_ids ? r.platform_ids.split(',') : [],
      status: rowStatus(r.created_at, r.updated_at),
      updatedAt: r.updated_at,
    })),
    total,
    page: safePage,
    pageSize: safePageSize,
  };
}

/** One platform's view: one row per post that has metrics on that platform, with that platform's curated columns. */
function listRecordsByPlatform(platform, filters, page, pageSize) {
  const { clauses, params } = buildPostWhereClause(filters);
  clauses.push('pm.platform = ?');
  params.push(platform);
  const where = `WHERE ${clauses.join(' AND ')}`;
  const total = db
    .prepare(`SELECT COUNT(*) AS cnt FROM posts p JOIN post_metrics pm ON pm.post_id = p.id ${where}`)
    .get(...params).cnt;
  const { safePage, safePageSize, offset } = paginate(page, pageSize);

  const rows = db
    .prepare(`
      SELECT p.id AS post_id, p.publish_date, p.caption, p.campaign_type, p.content_type,
             p.created_at, p.updated_at,
             pm.views, pm.reach, pm.impressions, pm.engagement, pm.clicks, pm.followers_gained,
             pm.watch_time_seconds, pm.shares, pm.comments, pm.saves, pm.posting_link,
             (SELECT MAX(rr.id) FROM raw_rows rr WHERE rr.post_id = p.id) AS raw_row_id
      FROM posts p JOIN post_metrics pm ON pm.post_id = p.id
      ${where}
      ORDER BY p.publish_date DESC, p.id DESC
      LIMIT ? OFFSET ?
    `)
    .all(...params, safePageSize, offset);

  return {
    platform,
    columns: PLATFORM_RECORD_COLUMNS[platform] || [],
    rows: rows.map((r) => ({
      postId: r.post_id,
      rawRowId: r.raw_row_id,
      publishDate: r.publish_date,
      caption: r.caption,
      campaignType: r.campaign_type,
      contentType: r.content_type,
      status: rowStatus(r.created_at, r.updated_at),
      updatedAt: r.updated_at,
      metrics: {
        views: r.views,
        reach: r.reach,
        impressions: r.impressions,
        engagement: r.engagement,
        clicks: r.clicks,
        followers_gained: r.followers_gained,
        watch_time_seconds: r.watch_time_seconds,
        shares: r.shares,
        comments: r.comments,
        saves: r.saves,
        posting_link: r.posting_link,
      },
    })),
    total,
    page: safePage,
    pageSize: safePageSize,
  };
}

function listRecordsTable({ platform, dateFrom, dateTo, campaignType, contentType, search, page, pageSize }) {
  const filters = { dateFrom, dateTo, campaignType, contentType, search };
  if (!platform || platform === 'all') return listRecordsSummary(filters, page, pageSize);
  return listRecordsByPlatform(platform, filters, page, pageSize);
}

const deletePostStmt = db.prepare('DELETE FROM posts WHERE id = ?');
const deletePlatformMetricStmt = db.prepare('DELETE FROM post_metrics WHERE post_id = ? AND platform = ?');
const countPlatformsForPostStmt = db.prepare('SELECT COUNT(*) AS cnt FROM post_metrics WHERE post_id = ?');

/** Deletes an entire record (every platform). raw_rows referencing it are kept, just detached (post_id -> NULL). */
function deletePost(postId) {
  const info = deletePostStmt.run(postId);
  if (info.changes === 0) throw new Error('Record not found.');
}

/** Deletes just one platform's metrics from a record; if that was its only platform, the whole record goes too. */
function deletePlatformFromPost(postId, platform) {
  const info = deletePlatformMetricStmt.run(postId, platform);
  if (info.changes === 0) throw new Error('That platform is not on this record.');
  const remaining = countPlatformsForPostStmt.get(postId).cnt;
  if (remaining === 0) deletePostStmt.run(postId);
}

/* ============================================================
   Analytics service — all aggregation/comparison queries
   ============================================================ */
const SUM_METRICS_SQL = CANONICAL_METRIC_KEYS.map((k) => `SUM(pm.${k}) AS ${k}`).join(', ');
const METRIC_LABELS = Object.fromEntries(CANONICAL_METRICS.map((m) => [m.key, m.label]));

/**
 * "Engagement Rate" is the one derived (non-imported) metric the dashboard
 * offers — engagement / (reach, falling back to impressions) * 100. It's
 * computed on the fly, never stored, and only ever offered where the
 * underlying columns actually have data (see platformMetricOptions).
 */
const ALL_METRIC_KEYS = [...CANONICAL_METRIC_KEYS, 'engagement_rate'];
const ENGAGEMENT_RATE_EXPR = '(pm.engagement * 100.0 / NULLIF(COALESCE(NULLIF(pm.reach, 0), pm.impressions), 0))';

function metricExpr(key) {
  return key === 'engagement_rate' ? ENGAGEMENT_RATE_EXPR : `pm.${key}`;
}

function metricLabel(key) {
  return key === 'engagement_rate' ? 'Engagement Rate' : METRIC_LABELS[key] || key;
}

function round1(n) {
  return Math.round(n * 10) / 10;
}

/** Adds a computed `engagement_rate` field to an already-summed totals/breakdown/trend row. */
function withEngagementRate(row) {
  const denom = row.reach || row.impressions;
  row.engagement_rate = denom && row.engagement ? round1((row.engagement / denom) * 100) : (row.engagement === 0 && denom ? 0 : null);
  return row;
}

function toIso(d) {
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}
function addDays(iso, days) {
  const [y, m, d] = iso.split('-').map(Number);
  const dt = new Date(y, m - 1, d + days);
  return toIso(dt);
}
function daysBetween(fromIso, toIso_) {
  const [y1, m1, d1] = fromIso.split('-').map(Number);
  const [y2, m2, d2] = toIso_.split('-').map(Number);
  return Math.round((new Date(y2, m2 - 1, d2) - new Date(y1, m1 - 1, d1)) / 86400000) + 1;
}

function monthRange(year, month) {
  const from = `${year}-${pad2(month)}-01`;
  const lastDay = new Date(year, month, 0).getDate();
  return { from, to: `${year}-${pad2(month)}-${pad2(lastDay)}` };
}
function quarterRange(year, quarter) {
  const startMonth = (quarter - 1) * 3 + 1;
  const { from } = monthRange(year, startMonth);
  const { to } = monthRange(year, startMonth + 2);
  return { from, to };
}
function ytdRange(year) {
  const today = new Date();
  const cappedTo = today.getFullYear() === year ? toIso(today) : `${year}-12-31`;
  return { from: `${year}-01-01`, to: year < today.getFullYear() ? `${year}-12-31` : cappedTo };
}
function previousEqualRange(from, to) {
  const len = daysBetween(from, to);
  const prevTo = addDays(from, -1);
  const prevFrom = addDays(prevTo, -(len - 1));
  return { from: prevFrom, to: prevTo };
}
function samePeriodLastYear(from, to) {
  const shift = (iso) => {
    const [y, m, d] = iso.split('-').map(Number);
    return `${y - 1}-${pad2(m)}-${pad2(d)}`;
  };
  return { from: shift(from), to: shift(to) };
}

/** Builds a WHERE clause + params array from the shared filter set used across the API. */
function buildFilter(filters = {}) {
  const clauses = [];
  const params = [];
  if (filters.dateFrom) {
    clauses.push('p.publish_date >= ?');
    params.push(filters.dateFrom);
  }
  if (filters.dateTo) {
    clauses.push('p.publish_date <= ?');
    params.push(filters.dateTo);
  }
  if (filters.platform && filters.platform !== 'all') {
    clauses.push('pm.platform = ?');
    params.push(filters.platform);
  }
  if (filters.campaignType && filters.campaignType !== 'all') {
    clauses.push('p.campaign_type = ?');
    params.push(filters.campaignType);
  }
  if (filters.contentType && filters.contentType !== 'all') {
    clauses.push('p.content_type LIKE ?');
    params.push(`%${filters.contentType}%`);
  }
  return { where: clauses.length ? `WHERE ${clauses.join(' AND ')}` : '', params };
}

function aggregateTotals(filters) {
  const { where, params } = buildFilter(filters);
  const sql = `
    SELECT ${SUM_METRICS_SQL}, COUNT(DISTINCT p.id) AS post_count
    FROM post_metrics pm JOIN posts p ON p.id = pm.post_id
    ${where}`;
  const row = db.prepare(sql).get(...params);
  const result = { post_count: row.post_count || 0 };
  for (const key of CANONICAL_METRIC_KEYS) result[key] = row[key] || 0;
  return withEngagementRate(result);
}

function platformBreakdown(filters) {
  const { where, params } = buildFilter(filters);
  const sql = `
    SELECT pm.platform, ${SUM_METRICS_SQL}, COUNT(DISTINCT p.id) AS post_count
    FROM post_metrics pm JOIN posts p ON p.id = pm.post_id
    ${where}
    GROUP BY pm.platform`;
  const rows = db.prepare(sql).all(...params);
  const byId = Object.fromEntries(rows.map((r) => [r.platform, r]));
  return PLATFORM_IDS.filter((id) => byId[id]).map((id) => {
    const r = byId[id];
    const meta = PLATFORMS.find((p) => p.id === id);
    const out = { platform: id, label: meta.label, color: meta.color, post_count: r.post_count };
    for (const key of CANONICAL_METRIC_KEYS) out[key] = r[key] || 0;
    return withEngagementRate(out);
  });
}

function pctChange(cur, prev) {
  if (!prev) return cur > 0 ? null : 0;
  return Math.round(((cur - prev) / prev) * 1000) / 10;
}

function withGrowth(current, previous) {
  const growth = {};
  for (const key of CANONICAL_METRIC_KEYS) growth[key] = pctChange(current[key] || 0, previous[key] || 0);
  growth.post_count = pctChange(current.post_count || 0, previous.post_count || 0);
  return growth;
}

/** Compares two arbitrary date ranges (same filters otherwise), returning totals, platform breakdown, and % growth. */
function compareRanges({ rangeA, rangeB, filters = {} }) {
  const fA = { ...filters, dateFrom: rangeA.from, dateTo: rangeA.to };
  const fB = { ...filters, dateFrom: rangeB.from, dateTo: rangeB.to };
  const totalsA = aggregateTotals(fA);
  const totalsB = aggregateTotals(fB);
  return {
    rangeA: { ...rangeA, totals: totalsA, platforms: platformBreakdown(fA) },
    rangeB: { ...rangeB, totals: totalsB, platforms: platformBreakdown(fB) },
    growth: withGrowth(totalsA, totalsB),
  };
}

function kpiSummary({ dateFrom, dateTo, filters = {} }) {
  const current = { from: dateFrom, to: dateTo };
  const previous = previousEqualRange(dateFrom, dateTo);
  const cmp = compareRanges({ rangeA: current, rangeB: previous, filters });
  return {
    dateRange: current,
    previousRange: previous,
    totals: cmp.rangeA.totals,
    growthVsPrevious: cmp.growth,
    platforms: cmp.rangeA.platforms,
  };
}

function weeklyTrend({ dateFrom, dateTo, filters = {} }) {
  const { where, params } = buildFilter({ ...filters, dateFrom, dateTo });
  const sql = `
    SELECT p.week_start AS period, ${SUM_METRICS_SQL}, COUNT(DISTINCT p.id) AS post_count
    FROM post_metrics pm JOIN posts p ON p.id = pm.post_id
    ${where}
    GROUP BY p.week_start
    ORDER BY p.week_start ASC`;
  return db.prepare(sql).all(...params).map(withEngagementRate);
}

function topPosts({ dateFrom, dateTo, filters = {}, sortBy = 'engagement', limit = 10 }) {
  const sortKey = ALL_METRIC_KEYS.includes(sortBy) ? sortBy : 'engagement';
  const expr = metricExpr(sortKey);
  const { where, params } = buildFilter({ ...filters, dateFrom, dateTo });
  const notNull = `${expr} IS NOT NULL`;
  const fullWhere = where ? `${where} AND ${notNull}` : `WHERE ${notNull}`;
  const sql = `
    SELECT p.id AS post_id, p.publish_date, p.content_type, p.campaign_type, p.caption,
           pm.platform, pm.posting_link, ${CANONICAL_METRIC_KEYS.map((k) => `pm.${k}`).join(', ')},
           ${expr} AS metric_value,
           (SELECT MAX(rr.id) FROM raw_rows rr WHERE rr.post_id = p.id) AS raw_row_id
    FROM post_metrics pm JOIN posts p ON p.id = pm.post_id
    ${fullWhere}
    ORDER BY ${expr} DESC NULLS LAST
    LIMIT ?`;
  return db.prepare(sql).all(...params, limit).map((row) => {
    row.raw_row_id = row.raw_row_id ?? null;
    return withEngagementRate(row);
  });
}

/** Breakdown of the ten canonical metrics (+ engagement_rate) grouped by campaign type (Ads/Organic/etc). */
function campaignBreakdown(filters) {
  const { where, params } = buildFilter(filters);
  const clause = where ? `${where} AND p.campaign_type IS NOT NULL AND p.campaign_type != ''` : "WHERE p.campaign_type IS NOT NULL AND p.campaign_type != ''";
  const sql = `
    SELECT p.campaign_type AS campaign_type, ${SUM_METRICS_SQL}, COUNT(DISTINCT p.id) AS post_count
    FROM post_metrics pm JOIN posts p ON p.id = pm.post_id
    ${clause}
    GROUP BY p.campaign_type
    ORDER BY p.campaign_type ASC`;
  return db.prepare(sql).all(...params).map(withEngagementRate);
}

/** Breakdown grouped by the post's raw content-type string (e.g. "REEL, STORY" is its own bucket, not split). */
function contentTypeBreakdown(filters) {
  const { where, params } = buildFilter(filters);
  const clause = where ? `${where} AND p.content_type IS NOT NULL AND p.content_type != ''` : "WHERE p.content_type IS NOT NULL AND p.content_type != ''";
  const sql = `
    SELECT p.content_type AS content_type, ${SUM_METRICS_SQL}, COUNT(DISTINCT p.id) AS post_count
    FROM post_metrics pm JOIN posts p ON p.id = pm.post_id
    ${clause}
    GROUP BY p.content_type
    ORDER BY p.content_type ASC`;
  return db.prepare(sql).all(...params).map(withEngagementRate);
}

/** The six single-metric KPI stats for the Dashboard's metric-focused KPI cards. */
function metricSummary({ dateFrom, dateTo, filters = {}, metric }) {
  const key = ALL_METRIC_KEYS.includes(metric) ? metric : 'engagement';
  const expr = metricExpr(key);
  const { where, params } = buildFilter({ ...filters, dateFrom, dateTo });
  const notNull = `${expr} IS NOT NULL`;
  const fullWhere = where ? `${where} AND ${notNull}` : `WHERE ${notNull}`;

  const agg = db
    .prepare(`
      SELECT MAX(${expr}) AS highest, MIN(${expr}) AS lowest, AVG(${expr}) AS average,
             SUM(${expr}) AS total, COUNT(*) AS post_count
      FROM post_metrics pm JOIN posts p ON p.id = pm.post_id
      ${fullWhere}
    `)
    .get(...params);

  const pickPost = (direction) =>
    db
      .prepare(`
        SELECT p.id AS post_id, p.publish_date, p.caption, p.campaign_type, p.content_type, pm.platform,
               ${expr} AS value, (SELECT MAX(rr.id) FROM raw_rows rr WHERE rr.post_id = p.id) AS raw_row_id
        FROM post_metrics pm JOIN posts p ON p.id = pm.post_id
        ${fullWhere}
        ORDER BY ${expr} ${direction} NULLS LAST
        LIMIT 1
      `)
      .get(...params) || null;

  return {
    metric: key,
    label: metricLabel(key),
    unit: key === 'watch_time_seconds' ? 'duration' : key === 'engagement_rate' ? 'percent' : 'number',
    total: agg.total || 0,
    average: agg.average || 0,
    highest: agg.highest ?? null,
    lowest: agg.lowest ?? null,
    postCount: agg.post_count || 0,
    bestPost: pickPost('DESC'),
    worstPost: pickPost('ASC'),
  };
}

/** Which metrics have real (non-null) data for this platform — drives the Dashboard's metric dropdown. Never hardcoded per platform; always read from what's actually in the database. */
function platformMetricOptions(platform) {
  const clauses = [];
  const params = [];
  if (platform && platform !== 'all') {
    clauses.push('platform = ?');
    params.push(platform);
  }
  const where = clauses.length ? `WHERE ${clauses.join(' AND ')}` : '';
  const presenceSql = CANONICAL_METRIC_KEYS.map((k) => `MAX(${k} IS NOT NULL) AS has_${k}`).join(', ');
  const row = db.prepare(`SELECT ${presenceSql} FROM post_metrics ${where}`).get(...params);

  const options = CANONICAL_METRIC_KEYS
    .filter((k) => row[`has_${k}`])
    .map((k) => ({ key: k, label: metricLabel(k), unit: k === 'watch_time_seconds' ? 'duration' : 'number' }));

  if (row.has_engagement && (row.has_reach || row.has_impressions)) {
    options.push({ key: 'engagement_rate', label: 'Engagement Rate', unit: 'percent' });
  }
  return options;
}

function monthlyReport({ year, month, filters = {} }) {
  const range = monthRange(year, month);
  const vsPrevious = compareRanges({ rangeA: range, rangeB: previousEqualRange(range.from, range.to), filters });
  const vsLastYear = compareRanges({ rangeA: range, rangeB: samePeriodLastYear(range.from, range.to), filters });
  return { range, totals: vsPrevious.rangeA.totals, platforms: vsPrevious.rangeA.platforms, vsPreviousPeriod: vsPrevious, vsLastYear };
}

function quarterlyReport({ year, quarter, filters = {} }) {
  const range = quarterRange(year, quarter);
  const vsPrevious = compareRanges({ rangeA: range, rangeB: previousEqualRange(range.from, range.to), filters });
  const vsLastYear = compareRanges({ rangeA: range, rangeB: samePeriodLastYear(range.from, range.to), filters });
  return { range, totals: vsPrevious.rangeA.totals, platforms: vsPrevious.rangeA.platforms, vsPreviousPeriod: vsPrevious, vsLastYear };
}

function ytdReport({ year, filters = {} }) {
  const range = ytdRange(year);
  const vsLastYear = compareRanges({ rangeA: range, rangeB: samePeriodLastYear(range.from, range.to), filters });
  return { range, totals: vsLastYear.rangeA.totals, platforms: vsLastYear.rangeA.platforms, vsLastYear };
}

function filterOptions() {
  const campaignTypes = db
    .prepare("SELECT DISTINCT campaign_type FROM posts WHERE campaign_type IS NOT NULL AND campaign_type != '' ORDER BY campaign_type")
    .all()
    .map((r) => r.campaign_type);

  const rawContentTypes = db
    .prepare("SELECT DISTINCT content_type FROM posts WHERE content_type IS NOT NULL AND content_type != ''")
    .all()
    .map((r) => r.content_type);
  const contentTypeSet = new Set();
  rawContentTypes.forEach((val) => val.split(',').forEach((v) => contentTypeSet.add(v.trim())));

  const dateRow = db.prepare('SELECT MIN(publish_date) AS min, MAX(publish_date) AS max FROM posts').get();
  const platformsInUse = db.prepare('SELECT DISTINCT platform FROM post_metrics').all().map((r) => r.platform);

  return {
    platforms: PLATFORMS.filter((p) => platformsInUse.includes(p.id)),
    campaignTypes,
    contentTypes: [...contentTypeSet].sort(),
    dateRange: { min: dateRow.min, max: dateRow.max },
  };
}

/* ============================================================
   Express app + routes
   ============================================================ */
const app = express();
const PORT = process.env.PORT || 4000;

app.use(attachSession);

function filtersFromQuery(q) {
  return {
    platform: q.platform || 'all',
    campaignType: q.campaignType || 'all',
    contentType: q.contentType || 'all',
  };
}

// ---- /api/auth ----
const authRouter = express.Router();

authRouter.post('/login', express.json(), (req, res) => {
  try {
    const { token } = login((req.body || {}).code);
    setSessionCookie(res, token);
    res.json({ ok: true });
  } catch (err) {
    res.status(401).json({ error: err.message });
  }
});

authRouter.post('/logout', (req, res) => {
  logout(req.sessionToken);
  clearSessionCookie(res);
  res.json({ ok: true });
});

authRouter.get('/me', (req, res) => {
  res.json({ authenticated: !!req.authenticated });
});

app.use('/api/auth', authRouter);

// ---- /api/uploads ----
const uploadStorage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOAD_DIR),
  filename: (req, file, cb) => cb(null, `${Date.now()}-${file.originalname.replace(/[^a-zA-Z0-9._-]/g, '_')}`),
});
const upload = multer({
  storage: uploadStorage,
  limits: { fileSize: 25 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ok = /\.(csv|xlsx|xls|txt)$/i.test(file.originalname);
    cb(ok ? null : new Error('Only .csv, .xlsx, .xls files are accepted.'), ok);
  },
});

const uploadRouter = express.Router();

uploadRouter.post('/preview', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file received.' });
  try {
    const preview = await previewImport(req.file.path, req.file.originalname);
    res.json({
      ...preview,
      filePath: req.file.path,
      originalName: req.file.originalname,
    });
  } catch (err) {
    fs.unlink(req.file.path, () => {});
    res.status(422).json({ error: err.message });
  }
});

uploadRouter.post('/commit', express.json(), async (req, res) => {
  const { filePath, originalName, defaultDuplicateAction, duplicateActions, notes } = req.body || {};
  if (!filePath || !originalName) return res.status(400).json({ error: 'Missing filePath/originalName. Re-run preview first.' });
  if (!fs.existsSync(filePath)) return res.status(410).json({ error: 'Upload expired or file was removed. Please re-upload.' });

  try {
    const result = await commitImport(filePath, originalName, {
      defaultDuplicateAction: defaultDuplicateAction || 'skip',
      duplicateActions: duplicateActions || {},
      notes: notes || null,
    });
    res.json(result);
  } catch (err) {
    res.status(422).json({ error: err.message });
  } finally {
    fs.unlink(filePath, () => {});
  }
});

uploadRouter.get('/history', (req, res) => {
  const limit = Math.min(Number(req.query.limit) || 50, 200);
  const rows = db
    .prepare('SELECT * FROM uploads ORDER BY uploaded_at DESC, id DESC LIMIT ?')
    .all(limit)
    .map((r) => ({
      ...r,
      weeks_affected: JSON.parse(r.weeks_affected || '[]'),
      action_summary: JSON.parse(r.action_summary || '{}'),
      sheet_summary: JSON.parse(r.sheet_summary || '[]'),
    }));
  res.json(rows);
});

uploadRouter.get('/:id/errors', (req, res) => {
  const rows = db.prepare('SELECT * FROM upload_errors WHERE upload_id = ? ORDER BY id ASC').all(req.params.id);
  res.json(rows);
});

uploadRouter.get('/:id/raw-rows', (req, res) => {
  const limit = Math.min(Number(req.query.limit) || 500, 2000);
  const rows = db
    .prepare('SELECT * FROM raw_rows WHERE upload_id = ? ORDER BY sheet_name ASC, row_number ASC LIMIT ?')
    .all(req.params.id, limit)
    .map((r) => {
      const meta = r.headers_json ? JSON.parse(r.headers_json) : {};
      return {
        ...r,
        headers: meta.headers || null,
        format: meta.format || null,
        raw: JSON.parse(r.raw_json),
        headers_json: undefined,
        raw_json: undefined,
      };
    });
  const total = db.prepare('SELECT COUNT(*) AS cnt FROM raw_rows WHERE upload_id = ?').get(req.params.id).cnt;
  res.json({ rows, total, limit });
});

app.use('/api/uploads', requireAuth, uploadRouter);

// ---- /api/records ----
const recordsRouter = express.Router();

recordsRouter.get('/', (req, res) => {
  const { dateFrom, dateTo, platform, campaignType, contentType, page, pageSize } = req.query;
  const result = listRecords({
    dateFrom,
    dateTo,
    platform,
    campaignType,
    contentType,
    page: Number(page) || 1,
    pageSize: Number(pageSize) || 25,
  });
  res.json(result);
});

recordsRouter.get('/table', (req, res) => {
  const { dateFrom, dateTo, platform, campaignType, contentType, search, page, pageSize } = req.query;
  const result = listRecordsTable({
    dateFrom,
    dateTo,
    platform,
    campaignType,
    contentType,
    search,
    page: Number(page) || 1,
    pageSize: Number(pageSize) || 25,
  });
  res.json(result);
});

recordsRouter.delete('/post/:postId', (req, res) => {
  try {
    deletePost(Number(req.params.postId));
    res.json({ ok: true });
  } catch (err) {
    res.status(404).json({ error: err.message });
  }
});

recordsRouter.delete('/post/:postId/platform/:platform', (req, res) => {
  try {
    deletePlatformFromPost(Number(req.params.postId), req.params.platform);
    res.json({ ok: true });
  } catch (err) {
    res.status(404).json({ error: err.message });
  }
});

recordsRouter.get('/:id', (req, res) => {
  const record = getRecord(req.params.id);
  if (!record) return res.status(404).json({ error: 'Record not found.' });
  res.json(record);
});

recordsRouter.put('/:id', express.json({ limit: '5mb' }), (req, res) => {
  const { values } = req.body || {};
  try {
    const record = updateRecord(req.params.id, values);
    res.json(record);
  } catch (err) {
    res.status(422).json({ error: err.message });
  }
});

app.use('/api/records', requireAuth, recordsRouter);

// ---- /api/analytics ----
const analyticsRouter = express.Router();

analyticsRouter.get('/filter-options', (req, res) => {
  res.json(filterOptions());
});

analyticsRouter.get('/kpis', (req, res) => {
  const { dateFrom, dateTo } = req.query;
  if (!dateFrom || !dateTo) return res.status(400).json({ error: 'dateFrom and dateTo are required.' });
  res.json(kpiSummary({ dateFrom, dateTo, filters: filtersFromQuery(req.query) }));
});

analyticsRouter.get('/platform-breakdown', (req, res) => {
  const { dateFrom, dateTo } = req.query;
  res.json(platformBreakdown({ ...filtersFromQuery(req.query), dateFrom, dateTo }));
});

analyticsRouter.get('/campaign-breakdown', (req, res) => {
  const { dateFrom, dateTo } = req.query;
  res.json(campaignBreakdown({ ...filtersFromQuery(req.query), dateFrom, dateTo }));
});

analyticsRouter.get('/content-type-breakdown', (req, res) => {
  const { dateFrom, dateTo } = req.query;
  res.json(contentTypeBreakdown({ ...filtersFromQuery(req.query), dateFrom, dateTo }));
});

analyticsRouter.get('/metric-options', (req, res) => {
  res.json({ options: platformMetricOptions(req.query.platform || 'all') });
});

analyticsRouter.get('/metric-summary', (req, res) => {
  const { dateFrom, dateTo, metric } = req.query;
  if (!dateFrom || !dateTo) return res.status(400).json({ error: 'dateFrom and dateTo are required.' });
  res.json(metricSummary({ dateFrom, dateTo, filters: filtersFromQuery(req.query), metric }));
});

analyticsRouter.get('/trend', (req, res) => {
  const { dateFrom, dateTo } = req.query;
  if (!dateFrom || !dateTo) return res.status(400).json({ error: 'dateFrom and dateTo are required.' });
  res.json(weeklyTrend({ dateFrom, dateTo, filters: filtersFromQuery(req.query) }));
});

analyticsRouter.get('/top-posts', (req, res) => {
  const { dateFrom, dateTo, sortBy, limit } = req.query;
  res.json(
    topPosts({
      dateFrom,
      dateTo,
      filters: filtersFromQuery(req.query),
      sortBy: sortBy || 'engagement',
      limit: Math.min(Number(limit) || 10, 100),
    })
  );
});

analyticsRouter.get('/compare', (req, res) => {
  const { fromA, toA, fromB, toB } = req.query;
  if (!fromA || !toA || !fromB || !toB) return res.status(400).json({ error: 'fromA, toA, fromB, toB are required.' });
  res.json(
    compareRanges({
      rangeA: { from: fromA, to: toA },
      rangeB: { from: fromB, to: toB },
      filters: filtersFromQuery(req.query),
    })
  );
});

analyticsRouter.get('/monthly', (req, res) => {
  const year = Number(req.query.year);
  const month = Number(req.query.month);
  if (!year || !month) return res.status(400).json({ error: 'year and month are required.' });
  res.json(monthlyReport({ year, month, filters: filtersFromQuery(req.query) }));
});

analyticsRouter.get('/quarterly', (req, res) => {
  const year = Number(req.query.year);
  const quarter = Number(req.query.quarter);
  if (!year || !quarter) return res.status(400).json({ error: 'year and quarter are required.' });
  res.json(quarterlyReport({ year, quarter, filters: filtersFromQuery(req.query) }));
});

analyticsRouter.get('/ytd', (req, res) => {
  const year = Number(req.query.year);
  if (!year) return res.status(400).json({ error: 'year is required.' });
  res.json(ytdReport({ year, filters: filtersFromQuery(req.query) }));
});

app.use('/api/analytics', requireAuth, analyticsRouter);

// ---- /api/backup — download the live database, or restore a previous one ----
const BACKUP_DIR = path.join(DATA_DIR, 'backup-uploads');
if (!fs.existsSync(BACKUP_DIR)) fs.mkdirSync(BACKUP_DIR, { recursive: true });

const backupUpload = multer({
  dest: BACKUP_DIR,
  limits: { fileSize: 200 * 1024 * 1024 },
});

const backupRouter = express.Router();

backupRouter.get('/export', (req, res) => {
  try {
    // Flushes the write-ahead log into the main file so the download is a
    // complete, self-contained snapshot (WAL mode keeps recent writes in a
    // separate -wal file that a plain copy of lrs.db alone could miss).
    db.exec('PRAGMA wal_checkpoint(FULL)');
  } catch (err) {
    return res.status(500).json({ error: `Could not prepare backup: ${err.message}` });
  }
  const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-');
  res.download(DB_PATH, `lrs-backup-${stamp}.db`, (err) => {
    if (err && !res.headersSent) res.status(500).json({ error: 'Backup download failed.' });
  });
});

// A SQLite file always starts with this exact 16-byte magic header.
const SQLITE_MAGIC = 'SQLite format 3 ';

backupRouter.post('/restore', backupUpload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file received.' });
  const cleanup = () => fs.unlink(req.file.path, () => {});

  try {
    const header = Buffer.alloc(16);
    const fd = fs.openSync(req.file.path, 'r');
    fs.readSync(fd, header, 0, 16, 0);
    fs.closeSync(fd);
    if (header.toString('utf8') !== SQLITE_MAGIC) {
      cleanup();
      return res.status(422).json({ error: 'That file is not a valid SQLite database backup.' });
    }

    // The live `db` handle (and every prepared statement built from it, all
    // created once at startup) stays bound to the connection that was open
    // when the process started — swapping the file underneath it wouldn't
    // be picked up. So: close it, replace the file, drop any stale WAL/SHM
    // sidecars from the *old* database, and exit — the process manager (or
    // you, running `npm start` again) brings it back up reading the
    // restored file fresh.
    db.close();
    fs.copyFileSync(req.file.path, DB_PATH);
    cleanup();
    for (const suffix of ['-wal', '-shm']) {
      const sidecar = `${DB_PATH}${suffix}`;
      if (fs.existsSync(sidecar)) fs.unlinkSync(sidecar);
    }

    res.json({ ok: true, message: 'Backup restored. The server is restarting to load it — refresh in a few seconds.' });
    setTimeout(() => process.exit(0), 300);
  } catch (err) {
    cleanup();
    res.status(500).json({ error: `Restore failed: ${err.message}` });
  }
});

app.use('/api/backup', requireAuth, backupRouter);

// ---- Frontend (embedded) ----
// Base64-encoded so the page's own inline <script> — which uses backticks and
// ${...} template literals extensively — never conflicts with the literal
// wrapping it here.
const INDEX_HTML_BASE64 = 'PCFkb2N0eXBlIGh0bWw+CjxodG1sIGxhbmc9ImVuIj4KPGhlYWQ+CjxtZXRhIGNoYXJzZXQ9IlVURi04IiAvPgo8bWV0YSBuYW1lPSJ2aWV3cG9ydCIgY29udGVudD0id2lkdGg9ZGV2aWNlLXdpZHRoLCBpbml0aWFsLXNjYWxlPTEuMCIgLz4KPHRpdGxlPkxSUyBBbmFseXRpY3MgRGFzaGJvYXJkPC90aXRsZT4KPGxpbmsgcmVsPSJwcmVjb25uZWN0IiBocmVmPSJodHRwczovL2ZvbnRzLmdvb2dsZWFwaXMuY29tIiAvPgo8bGluayByZWw9InByZWNvbm5lY3QiIGhyZWY9Imh0dHBzOi8vZm9udHMuZ3N0YXRpYy5jb20iIGNyb3Nzb3JpZ2luIC8+CjxsaW5rIGhyZWY9Imh0dHBzOi8vZm9udHMuZ29vZ2xlYXBpcy5jb20vY3NzMj9mYW1pbHk9SW50ZXI6d2dodEA0MDA7NTAwOzYwMDs3MDA7ODAwJmRpc3BsYXk9c3dhcCIgcmVsPSJzdHlsZXNoZWV0IiAvPgo8c2NyaXB0IHNyYz0iaHR0cHM6Ly9jZG4uanNkZWxpdnIubmV0L25wbS9jaGFydC5qc0A0LjQuNC9kaXN0L2NoYXJ0LnVtZC5taW4uanMiPjwvc2NyaXB0Pgo8c2NyaXB0IHNyYz0iaHR0cHM6Ly9jZG4uanNkZWxpdnIubmV0L25wbS9jaGFydGpzLXBsdWdpbi1kYXRhbGFiZWxzQDIvZGlzdC9jaGFydGpzLXBsdWdpbi1kYXRhbGFiZWxzLm1pbi5qcyI+PC9zY3JpcHQ+CjxzY3JpcHQgc3JjPSJodHRwczovL2Nkbi5qc2RlbGl2ci5uZXQvbnBtL2x1Y2lkZUAwLjQ2Mi4wL2Rpc3QvdW1kL2x1Y2lkZS5taW4uanMiPjwvc2NyaXB0Pgo8c3R5bGU+Ci8qIC0tLS0tLS0tLS0gRGVzaWduIHRva2VuczogZ2xhc3Ntb3JwaGlzbSBzdXJmYWNlcyArIHZhbGlkYXRlZCBjYXRlZ29yaWNhbC9zdGF0dXMgcGFsZXR0ZSAtLS0tLS0tLS0tICovCjpyb290IHsKICBjb2xvci1zY2hlbWU6IGxpZ2h0OwogIC0tZm9udC1zYW5zOiAnSW50ZXInLCAtYXBwbGUtc3lzdGVtLCBCbGlua01hY1N5c3RlbUZvbnQsICdTRiBQcm8gRGlzcGxheScsICdTZWdvZSBVSScsIFJvYm90bywgc2Fucy1zZXJpZjsKCiAgLS1wYWdlLXBsYW5lOiBsaW5lYXItZ3JhZGllbnQoMTgwZGVnLCAjZjdmOGZiIDAlLCAjZWNlZWYzIDEwMCUpOwogIC0tcGFnZS1wbGFuZS1zb2xpZDogI2VjZWVmMzsKICAtLXN1cmZhY2UtMTogcmdiYSgyNTUsIDI1NSwgMjU1LCAwLjY4KTsgLyogZ2xhc3M6IGNhcmRzLCBLUEkgdGlsZXMsIHRvcGJhciwgZmlsdGVyIGJhciAqLwogIC0tc3VyZmFjZS0yOiByZ2JhKDI1NSwgMjU1LCAyNTUsIDAuNTUpOyAvKiBnbGFzczogaW5wdXRzLCBuZXN0ZWQgcm93cywgcGlsbHMgKi8KICAtLXN1cmZhY2Utc29saWQ6ICNmZmZmZmY7CiAgLS1nbGFzcy1ibHVyOiBibHVyKDIwcHgpOwogIC0tYm9yZGVyOiByZ2JhKDE1LCAxNywgMjEsIDAuMDgpOwogIC0tdGV4dC1wcmltYXJ5OiAjMGYxMTE1OwogIC0tdGV4dC1zZWNvbmRhcnk6ICM1NjViNjY7CiAgLS10ZXh0LW11dGVkOiAjOGE4ZjlhOwogIC0tZ3JpZGxpbmU6IHJnYmEoMTUsIDE3LCAyMSwgMC4wOCk7CiAgLS1iYXNlbGluZTogcmdiYSgxNSwgMTcsIDIxLCAwLjIpOwogIC0tc3VjY2Vzcy10ZXh0OiAjMWE3ZDNjOwoKICAtLXN0YXR1cy1nb29kOiAjMWE5YzRhOwogIC0tc3RhdHVzLXdhcm5pbmc6ICNlMDhhMWY7CiAgLS1zdGF0dXMtc2VyaW91czogI2VjODM1YTsKICAtLXN0YXR1cy1jcml0aWNhbDogI2QxNDAzZjsKCiAgLS1zZXJpZXMtMTogIzJhNzhkNjsgLyogTFJTIGJsdWUg4oCUIGJyYW5kICsgZmFjZWJvb2sgKi8KICAtLXNlcmllcy0yOiAjMDA4MzAwOyAvKiBpbnN0YWdyYW0gKi8KICAtLXNlcmllcy0zOiAjZTg3YmE0OyAvKiB0aWt0b2sgKi8KICAtLXNlcmllcy00OiAjZWRhMTAwOyAvKiBsaW5rZWRpbiAqLwogIC0tc2VyaWVzLTU6ICMxYmFmN2E7IC8qIHRocmVhZHMgKi8KICAtLXNlcmllcy02OiAjZWI2ODM0OyAvKiB5b3V0dWJlICovCiAgLS1zZXJpZXMtNzogIzRhM2FhNzsgLyogcmVzZXJ2ZWQgKi8KICAtLXNlcmllcy04OiAjZTM0OTQ4OyAvKiByZXNlcnZlZCAqLwoKICAtLXJhZGl1cy1zbTogMTBweDsKICAtLXJhZGl1cy1tZDogMTRweDsKICAtLXJhZGl1cy1sZzogMThweDsKCiAgLS1zaGFkb3ctY2FyZDogMCAxcHggMnB4IHJnYmEoMTUsMTcsMjEsMC4wMyksIDAgOHB4IDI0cHggLTEwcHggcmdiYSgxNSwxNywyMSwwLjE0KTsKICAtLXNoYWRvdy1ob3ZlcjogMCA2cHggMTJweCAtMnB4IHJnYmEoMTUsMTcsMjEsMC4wOCksIDAgMThweCA0MHB4IC0xNHB4IHJnYmEoMTUsMTcsMjEsMC4yKTsKICAtLXNoYWRvdy1tb2RhbDogMCAyNHB4IDY0cHggLTEycHggcmdiYSgxNSwxNywyMSwwLjM1KTsKICAtLWVhc2U6IGN1YmljLWJlemllcigwLjQsIDAsIDAuMiwgMSk7Cn0KCkBtZWRpYSAocHJlZmVycy1jb2xvci1zY2hlbWU6IGRhcmspIHsKICA6cm9vdDp3aGVyZSg6bm90KFtkYXRhLXRoZW1lPSJsaWdodCJdKSkgewogICAgY29sb3Itc2NoZW1lOiBkYXJrOwogICAgLS1wYWdlLXBsYW5lOiBsaW5lYXItZ3JhZGllbnQoMTgwZGVnLCAjMGIwYzBmIDAlLCAjMTcxODFkIDEwMCUpOwogICAgLS1wYWdlLXBsYW5lLXNvbGlkOiAjMGIwYzBmOwogICAgLS1zdXJmYWNlLTE6IHJnYmEoMzIsIDM0LCA0MCwgMC42Mik7CiAgICAtLXN1cmZhY2UtMjogcmdiYSgyNTUsIDI1NSwgMjU1LCAwLjA2KTsKICAgIC0tc3VyZmFjZS1zb2xpZDogIzFjMWQyMzsKICAgIC0tYm9yZGVyOiByZ2JhKDI1NSwgMjU1LCAyNTUsIDAuMDkpOwogICAgLS10ZXh0LXByaW1hcnk6ICNmNGY1Zjc7CiAgICAtLXRleHQtc2Vjb25kYXJ5OiAjYjhiYmM0OwogICAgLS10ZXh0LW11dGVkOiAjODI4NjhmOwogICAgLS1ncmlkbGluZTogcmdiYSgyNTUsIDI1NSwgMjU1LCAwLjA4KTsKICAgIC0tYmFzZWxpbmU6IHJnYmEoMjU1LCAyNTUsIDI1NSwgMC4yKTsKICAgIC0tc3VjY2Vzcy10ZXh0OiAjMzRjNzZmOwoKICAgIC0tc3RhdHVzLWdvb2Q6ICMyZmI4NjI7CiAgICAtLXN0YXR1cy13YXJuaW5nOiAjZjBhMTNhOwogICAgLS1zdGF0dXMtc2VyaW91czogI2VjODM1YTsKICAgIC0tc3RhdHVzLWNyaXRpY2FsOiAjZTA2MDVmOwoKICAgIC0tc2VyaWVzLTE6ICMzOTg3ZTU7CiAgICAtLXNlcmllcy0yOiAjMDA4MzAwOwogICAgLS1zZXJpZXMtMzogI2Q1NTE4MTsKICAgIC0tc2VyaWVzLTQ6ICNjOTg1MDA7CiAgICAtLXNlcmllcy01OiAjMTk5ZTcwOwogICAgLS1zZXJpZXMtNjogI2Q5NTkyNjsKICAgIC0tc2VyaWVzLTc6ICM5MDg1ZTk7CiAgICAtLXNlcmllcy04OiAjZTY2NzY3OwoKICAgIC0tc2hhZG93LWNhcmQ6IDAgMXB4IDJweCByZ2JhKDAsMCwwLDAuMiksIDAgOHB4IDI0cHggLTEwcHggcmdiYSgwLDAsMCwwLjUpOwogICAgLS1zaGFkb3ctaG92ZXI6IDAgNnB4IDEycHggLTJweCByZ2JhKDAsMCwwLDAuMyksIDAgMThweCA0MHB4IC0xNHB4IHJnYmEoMCwwLDAsMC42KTsKICAgIC0tc2hhZG93LW1vZGFsOiAwIDI0cHggNjRweCAtMTJweCByZ2JhKDAsMCwwLDAuNyk7CiAgfQp9Cjpyb290W2RhdGEtdGhlbWU9ImRhcmsiXSB7CiAgY29sb3Itc2NoZW1lOiBkYXJrOwogIC0tcGFnZS1wbGFuZTogbGluZWFyLWdyYWRpZW50KDE4MGRlZywgIzBiMGMwZiAwJSwgIzE3MTgxZCAxMDAlKTsKICAtLXBhZ2UtcGxhbmUtc29saWQ6ICMwYjBjMGY7CiAgLS1zdXJmYWNlLTE6IHJnYmEoMzIsIDM0LCA0MCwgMC42Mik7CiAgLS1zdXJmYWNlLTI6IHJnYmEoMjU1LCAyNTUsIDI1NSwgMC4wNik7CiAgLS1zdXJmYWNlLXNvbGlkOiAjMWMxZDIzOwogIC0tYm9yZGVyOiByZ2JhKDI1NSwgMjU1LCAyNTUsIDAuMDkpOwogIC0tdGV4dC1wcmltYXJ5OiAjZjRmNWY3OwogIC0tdGV4dC1zZWNvbmRhcnk6ICNiOGJiYzQ7CiAgLS10ZXh0LW11dGVkOiAjODI4NjhmOwogIC0tZ3JpZGxpbmU6IHJnYmEoMjU1LCAyNTUsIDI1NSwgMC4wOCk7CiAgLS1iYXNlbGluZTogcmdiYSgyNTUsIDI1NSwgMjU1LCAwLjIpOwogIC0tc3VjY2Vzcy10ZXh0OiAjMzRjNzZmOwoKICAtLXN0YXR1cy1nb29kOiAjMmZiODYyOwogIC0tc3RhdHVzLXdhcm5pbmc6ICNmMGExM2E7CiAgLS1zdGF0dXMtc2VyaW91czogI2VjODM1YTsKICAtLXN0YXR1cy1jcml0aWNhbDogI2UwNjA1ZjsKCiAgLS1zZXJpZXMtMTogIzM5ODdlNTsKICAtLXNlcmllcy0yOiAjMDA4MzAwOwogIC0tc2VyaWVzLTM6ICNkNTUxODE7CiAgLS1zZXJpZXMtNDogI2M5ODUwMDsKICAtLXNlcmllcy01OiAjMTk5ZTcwOwogIC0tc2VyaWVzLTY6ICNkOTU5MjY7CiAgLS1zZXJpZXMtNzogIzkwODVlOTsKICAtLXNlcmllcy04OiAjZTY2NzY3OwoKICAtLXNoYWRvdy1jYXJkOiAwIDFweCAycHggcmdiYSgwLDAsMCwwLjIpLCAwIDhweCAyNHB4IC0xMHB4IHJnYmEoMCwwLDAsMC41KTsKICAtLXNoYWRvdy1ob3ZlcjogMCA2cHggMTJweCAtMnB4IHJnYmEoMCwwLDAsMC4zKSwgMCAxOHB4IDQwcHggLTE0cHggcmdiYSgwLDAsMCwwLjYpOwogIC0tc2hhZG93LW1vZGFsOiAwIDI0cHggNjRweCAtMTJweCByZ2JhKDAsMCwwLDAuNyk7Cn0KCiogeyBib3gtc2l6aW5nOiBib3JkZXItYm94OyB9Cmh0bWwsIGJvZHkgeyBoZWlnaHQ6IDEwMCU7IH0KYm9keSB7CiAgbWFyZ2luOiAwOwogIGZvbnQtZmFtaWx5OiB2YXIoLS1mb250LXNhbnMpOwogIGJhY2tncm91bmQ6IHZhcigtLXBhZ2UtcGxhbmUpOwogIGJhY2tncm91bmQtYXR0YWNobWVudDogZml4ZWQ7CiAgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7CiAgLXdlYmtpdC1mb250LXNtb290aGluZzogYW50aWFsaWFzZWQ7CiAgLW1vei1vc3gtZm9udC1zbW9vdGhpbmc6IGdyYXlzY2FsZTsKfQpidXR0b24sIHNlbGVjdCwgaW5wdXQsIHRleHRhcmVhIHsgZm9udC1mYW1pbHk6IGluaGVyaXQ7IH0KaDEsIGgyLCBoMywgaDQgeyBmb250LXdlaWdodDogNzAwOyBsZXR0ZXItc3BhY2luZzogLTAuMDFlbTsgfQoKOjpzZWxlY3Rpb24geyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDMwJSwgdHJhbnNwYXJlbnQpOyB9CgovKiBDdXN0b20gc2Nyb2xsYmFyIOKAlCB0aGluLCB1bm9idHJ1c2l2ZSwgZml0cyB0aGUgZ2xhc3MgYWVzdGhldGljICovCjo6LXdlYmtpdC1zY3JvbGxiYXIgeyB3aWR0aDogMTBweDsgaGVpZ2h0OiAxMHB4OyB9Cjo6LXdlYmtpdC1zY3JvbGxiYXItdHJhY2sgeyBiYWNrZ3JvdW5kOiB0cmFuc3BhcmVudDsgfQo6Oi13ZWJraXQtc2Nyb2xsYmFyLXRodW1iIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXRleHQtbXV0ZWQpIDQwJSwgdHJhbnNwYXJlbnQpOyBib3JkZXItcmFkaXVzOiAyMHB4OyBib3JkZXI6IDJweCBzb2xpZCB0cmFuc3BhcmVudDsgYmFja2dyb3VuZC1jbGlwOiBwYWRkaW5nLWJveDsgfQo6Oi13ZWJraXQtc2Nyb2xsYmFyLXRodW1iOmhvdmVyIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXRleHQtbXV0ZWQpIDYwJSwgdHJhbnNwYXJlbnQpOyBiYWNrZ3JvdW5kLWNsaXA6IHBhZGRpbmctYm94OyB9CgouYXBwLXNoZWxsIHsgbWluLWhlaWdodDogMTAwJTsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgfQoKLyogLS0tLS0tLS0tLSBUb3BiYXIgLS0tLS0tLS0tLSAqLwoudG9wYmFyIHsKICBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDI0cHg7CiAgcGFkZGluZzogMTJweCAyMHB4OwogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOwogIHBvc2l0aW9uOiBzdGlja3k7IHRvcDogMDsgei1pbmRleDogMjA7Cn0KLnRvcGJhci1icmFuZCB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBiYXNlbGluZTsgZ2FwOiA4cHg7IHdoaXRlLXNwYWNlOiBub3dyYXA7IH0KLmJyYW5kLW1hcmsgewogIGZvbnQtd2VpZ2h0OiA3MDA7IGxldHRlci1zcGFjaW5nOiAwLjAyZW07CiAgYmFja2dyb3VuZDogdmFyKC0tc2VyaWVzLTEpOyBjb2xvcjogI2ZmZjsKICBwYWRkaW5nOiA0cHggOXB4OyBib3JkZXItcmFkaXVzOiA4cHg7IGZvbnQtc2l6ZTogMTNweDsKICBib3gtc2hhZG93OiAwIDRweCAxNHB4IC00cHggY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSA2MCUsIHRyYW5zcGFyZW50KTsKfQouYnJhbmQtdGl0bGUgeyBmb250LXdlaWdodDogNjAwOyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgbGV0dGVyLXNwYWNpbmc6IC0wLjAxZW07IH0KCi50YWJzIHsgZGlzcGxheTogZmxleDsgZ2FwOiAycHg7IGZsZXg6IDE7IG92ZXJmbG93LXg6IGF1dG87IHBvc2l0aW9uOiByZWxhdGl2ZTsgfQoudGFiLWJ0biB7CiAgZGlzcGxheTogaW5saW5lLWZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogN3B4OwogIGJvcmRlcjogbm9uZTsgYmFja2dyb3VuZDogdHJhbnNwYXJlbnQ7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7CiAgcGFkZGluZzogOXB4IDE2cHg7IGJvcmRlci1yYWRpdXM6IDEwcHg7IGN1cnNvcjogcG9pbnRlcjsgZm9udC1zaXplOiAxNHB4OyBmb250LXdlaWdodDogNTAwOwogIHdoaXRlLXNwYWNlOiBub3dyYXA7IHBvc2l0aW9uOiByZWxhdGl2ZTsKICB0cmFuc2l0aW9uOiBjb2xvciAxODBtcyB2YXIoLS1lYXNlKSwgYmFja2dyb3VuZCAxODBtcyB2YXIoLS1lYXNlKTsKfQoudGFiLWJ0biBzdmcgeyBmbGV4LXNocmluazogMDsgb3BhY2l0eTogMC44OyB9Ci50YWItYnRuLmlzLWFjdGl2ZSBzdmcgeyBvcGFjaXR5OiAxOyBjb2xvcjogdmFyKC0tc2VyaWVzLTEpOyB9Ci50YWItYnRuOmhvdmVyIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSA4JSwgdHJhbnNwYXJlbnQpOyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgfQoudGFiLWJ0bi5pcy1hY3RpdmUgeyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgZm9udC13ZWlnaHQ6IDYwMDsgfQoudGFiLWJ0bi5pcy1hY3RpdmU6OmFmdGVyIHsKICBjb250ZW50OiAnJzsgcG9zaXRpb246IGFic29sdXRlOyBsZWZ0OiAxMnB4OyByaWdodDogMTJweDsgYm90dG9tOiAtMXB4OyBoZWlnaHQ6IDJweDsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zZXJpZXMtMSk7IGJvcmRlci1yYWRpdXM6IDJweCAycHggMCAwOwogIGFuaW1hdGlvbjogdGFiSW5kaWNhdG9ySW4gMjIwbXMgdmFyKC0tZWFzZSk7Cn0KQGtleWZyYW1lcyB0YWJJbmRpY2F0b3JJbiB7IGZyb20geyBvcGFjaXR5OiAwOyB0cmFuc2Zvcm06IHNjYWxlWCgwLjQpOyB9IHRvIHsgb3BhY2l0eTogMTsgdHJhbnNmb3JtOiBzY2FsZVgoMSk7IH0gfQoKLnRoZW1lLXRvZ2dsZSB7CiAgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0yKTsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHdpZHRoOiAzNnB4OyBoZWlnaHQ6IDM2cHg7IGJvcmRlci1yYWRpdXM6IDEwcHg7IGN1cnNvcjogcG9pbnRlcjsgZm9udC1zaXplOiAxNXB4OwogIGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogY2VudGVyOwogIHRyYW5zaXRpb246IHRyYW5zZm9ybSAxODBtcyB2YXIoLS1lYXNlKSwgYm94LXNoYWRvdyAxODBtcyB2YXIoLS1lYXNlKSwgYmFja2dyb3VuZCAxODBtcyB2YXIoLS1lYXNlKTsKfQoudGhlbWUtdG9nZ2xlOmhvdmVyIHsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKC0xcHgpOyBib3gtc2hhZG93OiB2YXIoLS1zaGFkb3ctaG92ZXIpOyB9Ci50aGVtZS10b2dnbGU6YWN0aXZlIHsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDApIHNjYWxlKDAuOTQpOyB9Ci50b3BiYXItdXNlciB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogMTBweDsgZm9udC1zaXplOiAxM3B4OyB9CgovKiAtLS0tLS0tLS0tIEF1dGggc2NyZWVuIC0tLS0tLS0tLS0gKi8KLmF1dGgtc2NyZWVuIHsKICBtaW4taGVpZ2h0OiAxMDB2aDsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsganVzdGlmeS1jb250ZW50OiBjZW50ZXI7CiAgYmFja2dyb3VuZDogdmFyKC0tcGFnZS1wbGFuZSk7IHBhZGRpbmc6IDIwcHg7Cn0KLmF1dGgtY2FyZCB7CiAgd2lkdGg6IDEwMCU7IG1heC13aWR0aDogNDAwcHg7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1sZyk7IHBhZGRpbmc6IDMycHg7IGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdy1tb2RhbCk7CiAgYW5pbWF0aW9uOiBtb2RhbFBhbmVsSW4gMjYwbXMgdmFyKC0tZWFzZSk7Cn0KLmF1dGgtYnJhbmQgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogYmFzZWxpbmU7IGdhcDogOHB4OyBtYXJnaW4tYm90dG9tOiAyMnB4OyB9Ci5hdXRoLWJyYW5kIC5icmFuZC10aXRsZSB7IGZvbnQtd2VpZ2h0OiA3MDA7IGZvbnQtc2l6ZTogMTdweDsgfQouYXV0aC1mb3JtIHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiAxNHB4OyBtYXJnaW4tdG9wOiAxNnB4OyB9Ci5hdXRoLWZvcm0gLmZvcm0tZmllbGQgaW5wdXQgeyB3aWR0aDogMTAwJTsgfQouYXV0aC1lcnJvciB7IGNvbG9yOiB2YXIoLS1zdGF0dXMtY3JpdGljYWwpOyBmb250LXNpemU6IDEycHg7IG1pbi1oZWlnaHQ6IDE2cHg7IH0KCi8qIC0tLS0tLS0tLS0gRmlsdGVyIGJhciAtLS0tLS0tLS0tICovCi5maWx0ZXItYmFyIHsKICBkaXNwbGF5OiBmbGV4OyBmbGV4LXdyYXA6IHdyYXA7IGFsaWduLWl0ZW1zOiBlbmQ7IGdhcDogMTZweDsKICBwYWRkaW5nOiAxNHB4IDIwcHg7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOwogIHBvc2l0aW9uOiBzdGlja3k7IHRvcDogNTdweDsgei1pbmRleDogMTk7Cn0KLmZpbHRlci1maWVsZCB7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGdhcDogNXB4OyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IH0KLmZpbHRlci1maWVsZCBsYWJlbCB7IGZvbnQtd2VpZ2h0OiA2MDA7IH0KLmZpbHRlci1wcmVzZXRzIHsgZmxleC1kaXJlY3Rpb246IHJvdzsgZ2FwOiA2cHg7IH0KLmZpbHRlci1wcmVzZXRzIGJ1dHRvbiB7CiAgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0yKTsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgYm9yZGVyLXJhZGl1czogMjBweDsgcGFkZGluZzogN3B4IDEzcHg7IGZvbnQtc2l6ZTogMTJweDsgZm9udC13ZWlnaHQ6IDUwMDsgY3Vyc29yOiBwb2ludGVyOwogIHRyYW5zaXRpb246IGNvbG9yIDE4MG1zIHZhcigtLWVhc2UpLCBiYWNrZ3JvdW5kIDE4MG1zIHZhcigtLWVhc2UpLCB0cmFuc2Zvcm0gMTUwbXMgdmFyKC0tZWFzZSk7Cn0KLmZpbHRlci1wcmVzZXRzIGJ1dHRvbjpob3ZlciB7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoLTFweCk7IH0KLmZpbHRlci1wcmVzZXRzIGJ1dHRvbjphY3RpdmUgeyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMCkgc2NhbGUoMC45Nik7IH0KLmZpbHRlci1wcmVzZXRzIGJ1dHRvbi5pcy1hY3RpdmUgeyBiYWNrZ3JvdW5kOiB2YXIoLS1zZXJpZXMtMSk7IGNvbG9yOiAjZmZmOyBib3JkZXItY29sb3I6IHRyYW5zcGFyZW50OyBib3gtc2hhZG93OiAwIDRweCAxNHB4IC01cHggY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSA2MCUsIHRyYW5zcGFyZW50KTsgfQoKLyogLS0tLS0tLS0tLSBWaWV3IGFyZWEgLS0tLS0tLS0tLSAqLwoudmlldy1hcmVhIHsgZmxleDogMTsgcGFkZGluZzogMjRweDsgbWF4LXdpZHRoOiAxNDAwcHg7IHdpZHRoOiAxMDAlOyBtYXJnaW46IDAgYXV0bzsgfQoudmlldyB7IGRpc3BsYXk6IG5vbmU7IH0KLnZpZXcuaXMtYWN0aXZlIHsgZGlzcGxheTogYmxvY2s7IGFuaW1hdGlvbjogdmlld0ZhZGVJbiAyNjBtcyB2YXIoLS1lYXNlKTsgfQpAa2V5ZnJhbWVzIHZpZXdGYWRlSW4gewogIGZyb20geyBvcGFjaXR5OiAwOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoNnB4KTsgfQogIHRvIHsgb3BhY2l0eTogMTsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDApOyB9Cn0KCi5zZWN0aW9uLXRpdGxlIHsgZm9udC1zaXplOiAxNnB4OyBmb250LXdlaWdodDogNzAwOyBtYXJnaW46IDMycHggMCAxNHB4OyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgbGV0dGVyLXNwYWNpbmc6IC0wLjAxZW07IH0KLnNlY3Rpb24tdGl0bGU6Zmlyc3QtY2hpbGQgeyBtYXJnaW4tdG9wOiAwOyB9CgovKiAtLS0tLS0tLS0tIElucHV0cyDigJQgb25lIHNoYXJlZCBnbGFzcyB0cmVhdG1lbnQgZm9yIGV2ZXJ5IHRleHQgaW5wdXQsIHNlbGVjdCwgYW5kIGRhdGUgcGlja2VyIC0tLS0tLS0tLS0gKi8KLmZpbHRlci1maWVsZCBzZWxlY3QsIC5maWx0ZXItZmllbGQgaW5wdXRbdHlwZT0iZGF0ZSJdLAouZm9ybS1maWVsZCBpbnB1dCwgLmZvcm0tZmllbGQgc2VsZWN0LCAuZm9ybS1maWVsZCB0ZXh0YXJlYSwKLmRhc2hib2FyZC1jb250cm9scyBzZWxlY3QsIC5yZWNvcmRzLXNlYXJjaCBpbnB1dCwKLmZpZWxkLWlubGluZSBzZWxlY3QsIC5maWVsZC1pbmxpbmUgaW5wdXQsCi5jb25mbGljdC1yb3cgc2VsZWN0LCAuY2FyZC1oZWFkZXIgc2VsZWN0IHsKICBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtc20pOwogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMik7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyBmb250LXNpemU6IDEzcHg7CiAgcGFkZGluZzogOHB4IDEycHg7IG1pbi13aWR0aDogMTQwcHg7CiAgdHJhbnNpdGlvbjogYm9yZGVyLWNvbG9yIDE2MG1zIHZhcigtLWVhc2UpLCBib3gtc2hhZG93IDE2MG1zIHZhcigtLWVhc2UpOwp9Ci5maWx0ZXItZmllbGQgc2VsZWN0OmhvdmVyLCAuZmlsdGVyLWZpZWxkIGlucHV0W3R5cGU9ImRhdGUiXTpob3ZlciwKLmZvcm0tZmllbGQgaW5wdXQ6aG92ZXIsIC5mb3JtLWZpZWxkIHNlbGVjdDpob3ZlciwKLmRhc2hib2FyZC1jb250cm9scyBzZWxlY3Q6aG92ZXIsIC5yZWNvcmRzLXNlYXJjaCBpbnB1dDpob3ZlciwKLmZpZWxkLWlubGluZSBzZWxlY3Q6aG92ZXIsIC5maWVsZC1pbmxpbmUgaW5wdXQ6aG92ZXIsCi5jb25mbGljdC1yb3cgc2VsZWN0OmhvdmVyLCAuY2FyZC1oZWFkZXIgc2VsZWN0OmhvdmVyIHsKICBib3JkZXItY29sb3I6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zZXJpZXMtMSkgMzUlLCB2YXIoLS1ib3JkZXIpKTsKfQouZmlsdGVyLWZpZWxkIHNlbGVjdDpmb2N1cywgLmZpbHRlci1maWVsZCBpbnB1dFt0eXBlPSJkYXRlIl06Zm9jdXMsCi5mb3JtLWZpZWxkIGlucHV0OmZvY3VzLCAuZm9ybS1maWVsZCBzZWxlY3Q6Zm9jdXMsIC5mb3JtLWZpZWxkIHRleHRhcmVhOmZvY3VzLAouZGFzaGJvYXJkLWNvbnRyb2xzIHNlbGVjdDpmb2N1cywgLnJlY29yZHMtc2VhcmNoIGlucHV0OmZvY3VzLAouZmllbGQtaW5saW5lIHNlbGVjdDpmb2N1cywgLmZpZWxkLWlubGluZSBpbnB1dDpmb2N1cywKLmNvbmZsaWN0LXJvdyBzZWxlY3Q6Zm9jdXMsIC5jYXJkLWhlYWRlciBzZWxlY3Q6Zm9jdXMsCi5hdXRoLWZvcm0gaW5wdXQ6Zm9jdXMgewogIG91dGxpbmU6IG5vbmU7IGJvcmRlci1jb2xvcjogdmFyKC0tc2VyaWVzLTEpOwogIGJveC1zaGFkb3c6IDAgMCAwIDNweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDE4JSwgdHJhbnNwYXJlbnQpOwp9CgovKiAtLS0tLS0tLS0tIFN0YXQgdGlsZXMgLS0tLS0tLS0tLSAqLwouc3RhdC1ncmlkIHsKICBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdChhdXRvLWZpdCwgbWlubWF4KDE4MHB4LCAxZnIpKTsgZ2FwOiAxNHB4Owp9Ci5zdGF0LXRpbGUgewogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1tZCk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHBhZGRpbmc6IDE2cHggMThweDsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWNhcmQpOwogIHRyYW5zaXRpb246IHRyYW5zZm9ybSAyMDBtcyB2YXIoLS1lYXNlKSwgYm94LXNoYWRvdyAyMDBtcyB2YXIoLS1lYXNlKTsKICBhbmltYXRpb246IGNhcmRJbiAzMjBtcyB2YXIoLS1lYXNlKSBiYWNrd2FyZHM7Cn0KLnN0YXQtdGlsZTpob3ZlciB7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgtM3B4KTsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWhvdmVyKTsgfQouc3RhdC1sYWJlbCB7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgZm9udC13ZWlnaHQ6IDYwMDsgfQouc3RhdC12YWx1ZSB7IGZvbnQtc2l6ZTogMjdweDsgZm9udC13ZWlnaHQ6IDcwMDsgbWFyZ2luLXRvcDogNXB4OyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgbGV0dGVyLXNwYWNpbmc6IC0wLjAyZW07IH0KLnN0YXQtZGVsdGEgeyBmb250LXNpemU6IDEycHg7IG1hcmdpbi10b3A6IDdweDsgZm9udC13ZWlnaHQ6IDYwMDsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA0cHg7IH0KLnN0YXQtZGVsdGEudXAgeyBjb2xvcjogdmFyKC0tc3VjY2Vzcy10ZXh0KTsgfQouc3RhdC1kZWx0YS5kb3duIHsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IH0KLnN0YXQtZGVsdGEuZmxhdCB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgfQouc3RhdC1kZWx0YS51cDo6YmVmb3JlIHsgY29udGVudDogJ+KGkSc7IH0KLnN0YXQtZGVsdGEuZG93bjo6YmVmb3JlIHsgY29udGVudDogJ+KGkyc7IH0KCkBrZXlmcmFtZXMgY2FyZEluIHsKICBmcm9tIHsgb3BhY2l0eTogMDsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDEwcHgpOyB9CiAgdG8geyBvcGFjaXR5OiAxOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMCk7IH0KfQoKLyogLS0tLS0tLS0tLSBDYXJkcyAvIGNoYXJ0cyAtLS0tLS0tLS0tICovCi5jYXJkLWdyaWQgeyBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IDJmciAxZnI7IGdhcDogMTZweDsgYWxpZ24taXRlbXM6IHN0YXJ0OyB9Ci5jYXJkLWdyaWQuZXZlbiB7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogMWZyIDFmcjsgfQpAbWVkaWEgKG1heC13aWR0aDogOTAwcHgpIHsgLmNhcmQtZ3JpZCwgLmNhcmQtZ3JpZC5ldmVuIHsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiAxZnI7IH0gfQouY2FyZCB7CiAgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLWxnKTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgcGFkZGluZzogMThweDsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWNhcmQpOwogIHRyYW5zaXRpb246IGJveC1zaGFkb3cgMjIwbXMgdmFyKC0tZWFzZSksIHRyYW5zZm9ybSAyMjBtcyB2YXIoLS1lYXNlKTsKICBhbmltYXRpb246IGNhcmRJbiAzMjBtcyB2YXIoLS1lYXNlKSBiYWNrd2FyZHM7Cn0KLmNhcmQ6aG92ZXIgeyBib3gtc2hhZG93OiB2YXIoLS1zaGFkb3ctaG92ZXIpOyB9Ci5jYXJkLWhlYWRlciB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogc3BhY2UtYmV0d2VlbjsgZ2FwOiA4cHg7IG1hcmdpbi1ib3R0b206IDE0cHg7IH0KLmNhcmQtaGVhZGVyIGgzIHsgZm9udC1zaXplOiAxNHB4OyBtYXJnaW46IDA7IGZvbnQtd2VpZ2h0OiA3MDA7IGxldHRlci1zcGFjaW5nOiAtMC4wMDVlbTsgfQouY2FyZC1oZWFkZXIgc2VsZWN0IHsgZm9udC1zaXplOiAxMnB4OyBwYWRkaW5nOiA2cHggMTBweDsgbWluLXdpZHRoOiAwOyB9Ci5jaGFydC13cmFwIHsgcG9zaXRpb246IHJlbGF0aXZlOyBoZWlnaHQ6IDI4MHB4OyB9Ci5jaGFydC13cmFwLnRhbGwgeyBoZWlnaHQ6IDM0MHB4OyB9CgoubGVnZW5kLXJvdyB7IGRpc3BsYXk6IGZsZXg7IGZsZXgtd3JhcDogd3JhcDsgZ2FwOiAxMnB4OyBtYXJnaW4tdG9wOiAxMHB4OyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IH0KLmxlZ2VuZC1pdGVtIHsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA2cHg7IH0KLmxlZ2VuZC1zd2F0Y2ggeyB3aWR0aDogMTBweDsgaGVpZ2h0OiAxMHB4OyBib3JkZXItcmFkaXVzOiAzcHg7IGRpc3BsYXk6IGlubGluZS1ibG9jazsgfQoubGVnZW5kLWxpbmUgeyB3aWR0aDogMTRweDsgaGVpZ2h0OiAycHg7IGJvcmRlci1yYWRpdXM6IDJweDsgZGlzcGxheTogaW5saW5lLWJsb2NrOyB9CgovKiAtLS0tLS0tLS0tIFRhYmxlcyDigJQgcHJlbWl1bSBkYXRhYmFzZSBmZWVsLCBub3QgYSBzcHJlYWRzaGVldCAtLS0tLS0tLS0tICovCi50YWJsZS1zY3JvbGwgewogIG92ZXJmbG93LXg6IGF1dG87IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1tZCk7CiAgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0yKTsKfQouZGF0YS10YWJsZSB7IHdpZHRoOiAxMDAlOyBib3JkZXItY29sbGFwc2U6IHNlcGFyYXRlOyBib3JkZXItc3BhY2luZzogMDsgZm9udC1zaXplOiAxM3B4OyB9Ci5kYXRhLXRhYmxlIHRoLCAuZGF0YS10YWJsZSB0ZCB7IHRleHQtYWxpZ246IGxlZnQ7IHBhZGRpbmc6IDExcHggMTRweDsgYm9yZGVyLWJvdHRvbTogMXB4IHNvbGlkIHZhcigtLWdyaWRsaW5lKTsgfQouZGF0YS10YWJsZSB0aGVhZCB0aCB7CiAgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgZm9udC13ZWlnaHQ6IDYwMDsgZm9udC1zaXplOiAxMXB4OyB0ZXh0LXRyYW5zZm9ybTogdXBwZXJjYXNlOyBsZXR0ZXItc3BhY2luZzogMC4wNGVtOwogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBwb3NpdGlvbjogc3RpY2t5OyB0b3A6IDA7IHotaW5kZXg6IDE7Cn0KLmRhdGEtdGFibGUgdGJvZHkgdHI6bnRoLWNoaWxkKGV2ZW4pIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXRleHQtbXV0ZWQpIDQlLCB0cmFuc3BhcmVudCk7IH0KLmRhdGEtdGFibGUgdGQubnVtIHsgZm9udC12YXJpYW50LW51bWVyaWM6IHRhYnVsYXItbnVtczsgdGV4dC1hbGlnbjogcmlnaHQ7IH0KLmRhdGEtdGFibGUgdGgubnVtIHsgdGV4dC1hbGlnbjogcmlnaHQ7IH0KLmRhdGEtdGFibGUgdGJvZHkgdHIgeyB0cmFuc2l0aW9uOiBiYWNrZ3JvdW5kIDE1MG1zIHZhcigtLWVhc2UpOyB9Ci5kYXRhLXRhYmxlIHRib2R5IHRyOmhvdmVyIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSA3JSwgdHJhbnNwYXJlbnQpOyB9Ci5kYXRhLXRhYmxlIHRib2R5IHRyOmxhc3QtY2hpbGQgdGQgeyBib3JkZXItYm90dG9tOiBub25lOyB9Ci5wbGF0Zm9ybS1waWxsIHsKICBkaXNwbGF5OiBpbmxpbmUtZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA2cHg7IGZvbnQtc2l6ZTogMTJweDsgZm9udC13ZWlnaHQ6IDYwMDsKICBwYWRkaW5nOiA0cHggMTBweDsgYm9yZGVyLXJhZGl1czogMjBweDsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsKfQoucGxhdGZvcm0tZG90IHsgd2lkdGg6IDhweDsgaGVpZ2h0OiA4cHg7IGJvcmRlci1yYWRpdXM6IDUwJTsgfQoKLyogLS0tLS0tLS0tLSBCdXR0b25zIOKAlCBuZXZlciBmbGF0OiBzb2Z0IHNoYWRvdywgaG92ZXIgbGlmdCwgcHJlc3Mgc2NhbGUgLS0tLS0tLS0tLSAqLwouYnRuIHsKICBkaXNwbGF5OiBpbmxpbmUtZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsganVzdGlmeS1jb250ZW50OiBjZW50ZXI7IGdhcDogNnB4OwogIGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMik7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOwogIGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBwYWRkaW5nOiA5cHggMTdweDsgYm9yZGVyLXJhZGl1czogMTFweDsgY3Vyc29yOiBwb2ludGVyOyBmb250LXNpemU6IDEzcHg7IGZvbnQtd2VpZ2h0OiA2MDA7CiAgYm94LXNoYWRvdzogMCAxcHggMnB4IHJnYmEoMTUsMTcsMjEsMC4wNCk7CiAgdHJhbnNpdGlvbjogdHJhbnNmb3JtIDE1MG1zIHZhcigtLWVhc2UpLCBib3gtc2hhZG93IDE1MG1zIHZhcigtLWVhc2UpLCBmaWx0ZXIgMTUwbXMgdmFyKC0tZWFzZSksIGJhY2tncm91bmQgMTUwbXMgdmFyKC0tZWFzZSk7Cn0KLmJ0biBzdmcgeyBmbGV4LXNocmluazogMDsgfQouYnRuOmhvdmVyIHsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKC0xcHgpOyBib3gtc2hhZG93OiB2YXIoLS1zaGFkb3ctaG92ZXIpOyBmaWx0ZXI6IGJyaWdodG5lc3MoMS4wMik7IH0KLmJ0bjphY3RpdmUgeyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMCkgc2NhbGUoMC45Nik7IGJveC1zaGFkb3c6IDAgMXB4IDJweCByZ2JhKDE1LDE3LDIxLDAuMDYpOyB9Ci5idG4ucHJpbWFyeSB7CiAgYmFja2dyb3VuZDogdmFyKC0tc2VyaWVzLTEpOyBjb2xvcjogI2ZmZjsgYm9yZGVyLWNvbG9yOiB0cmFuc3BhcmVudDsKICBib3gtc2hhZG93OiAwIDRweCAxNHB4IC01cHggY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSA2NSUsIHRyYW5zcGFyZW50KTsKfQouYnRuLnByaW1hcnk6aG92ZXIgeyBmaWx0ZXI6IGJyaWdodG5lc3MoMS4wNyk7IGJveC1zaGFkb3c6IDAgOHB4IDIycHggLTZweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDcwJSwgdHJhbnNwYXJlbnQpOyB9Ci5idG4uZGFuZ2VyIHsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdGF0dXMtY3JpdGljYWwpOyBjb2xvcjogI2ZmZjsgYm9yZGVyLWNvbG9yOiB0cmFuc3BhcmVudDsKICBib3gtc2hhZG93OiAwIDRweCAxNHB4IC01cHggY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXN0YXR1cy1jcml0aWNhbCkgNTUlLCB0cmFuc3BhcmVudCk7Cn0KLmJ0bi5kYW5nZXI6aG92ZXIgeyBmaWx0ZXI6IGJyaWdodG5lc3MoMS4wNik7IGJveC1zaGFkb3c6IDAgOHB4IDIycHggLTZweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc3RhdHVzLWNyaXRpY2FsKSA2MCUsIHRyYW5zcGFyZW50KTsgfQouYnRuLnN1Y2Nlc3MgewogIGJhY2tncm91bmQ6IHZhcigtLXN0YXR1cy1nb29kKTsgY29sb3I6ICNmZmY7IGJvcmRlci1jb2xvcjogdHJhbnNwYXJlbnQ7CiAgYm94LXNoYWRvdzogMCA0cHggMTRweCAtNXB4IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zdGF0dXMtZ29vZCkgNTUlLCB0cmFuc3BhcmVudCk7Cn0KLmJ0bi5zdWNjZXNzOmhvdmVyIHsgZmlsdGVyOiBicmlnaHRuZXNzKDEuMDYpOyBib3gtc2hhZG93OiAwIDhweCAyMnB4IC02cHggY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXN0YXR1cy1nb29kKSA2MCUsIHRyYW5zcGFyZW50KTsgfQouYnRuOmRpc2FibGVkIHsgb3BhY2l0eTogMC40NTsgY3Vyc29yOiBub3QtYWxsb3dlZDsgdHJhbnNmb3JtOiBub25lOyBib3gtc2hhZG93OiBub25lOyBmaWx0ZXI6IG5vbmU7IH0KLmJ0bi1yb3cgeyBkaXNwbGF5OiBmbGV4OyBnYXA6IDhweDsgZmxleC13cmFwOiB3cmFwOyB9CgovKiAtLS0tLS0tLS0tIFVwbG9hZCAtLS0tLS0tLS0tICovCi5kcm9wem9uZSB7CiAgYm9yZGVyOiAycHggZGFzaGVkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1sZyk7IHBhZGRpbmc6IDQwcHggMjBweDsKICB0ZXh0LWFsaWduOiBjZW50ZXI7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBjdXJzb3I6IHBvaW50ZXI7IHRyYW5zaXRpb246IGJvcmRlci1jb2xvciAyMDBtcyB2YXIoLS1lYXNlKSwgYmFja2dyb3VuZCAyMDBtcyB2YXIoLS1lYXNlKSwgdHJhbnNmb3JtIDIwMG1zIHZhcigtLWVhc2UpOwp9Ci5kcm9wem9uZTpob3ZlciB7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgtMXB4KTsgfQouZHJvcHpvbmUuaXMtZHJhZyB7IGJvcmRlci1jb2xvcjogdmFyKC0tc2VyaWVzLTEpOyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDYlLCB2YXIoLS1zdXJmYWNlLTIpKTsgdHJhbnNmb3JtOiBzY2FsZSgxLjAwNSk7IH0KLmRyb3B6b25lIGgzIHsgbWFyZ2luOiAwIDAgNnB4OyBmb250LXNpemU6IDE1cHg7IH0KLmRyb3B6b25lIHAgeyBtYXJnaW46IDA7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IGZvbnQtc2l6ZTogMTNweDsgfQouZHJvcHpvbmUgaW5wdXRbdHlwZT0iZmlsZSJdIHsgZGlzcGxheTogbm9uZTsgfQoKLmNvbmZsaWN0LWxpc3QgeyBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBnYXA6IDhweDsgbWFyZ2luOiAxMnB4IDA7IH0KLmNvbmZsaWN0LXJvdyB7CiAgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsganVzdGlmeS1jb250ZW50OiBzcGFjZS1iZXR3ZWVuOyBnYXA6IDEycHg7CiAgcGFkZGluZzogMTFweCAxNHB4OyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtc20pOyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTIpOwogIHRyYW5zaXRpb246IGJveC1zaGFkb3cgMTgwbXMgdmFyKC0tZWFzZSk7Cn0KLmNvbmZsaWN0LXJvdzpob3ZlciB7IGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdy1jYXJkKTsgfQouY29uZmxpY3Qtcm93IC53ZWVrLWxhYmVsIHsgZm9udC13ZWlnaHQ6IDYwMDsgZm9udC1zaXplOiAxM3B4OyB9Ci5jb25mbGljdC1yb3cgLndlZWstbWV0YSB7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgfQouY29uZmxpY3Qtcm93IHNlbGVjdCB7IG1pbi13aWR0aDogMDsgfQoKLmJhZGdlIHsgZGlzcGxheTogaW5saW5lLWJsb2NrOyBwYWRkaW5nOiAzcHggMTBweDsgYm9yZGVyLXJhZGl1czogMjBweDsgZm9udC1zaXplOiAxMXB4OyBmb250LXdlaWdodDogNzAwOyB9Ci5iYWRnZS5zdWNjZXNzIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXN0YXR1cy1nb29kKSAxOCUsIHRyYW5zcGFyZW50KTsgY29sb3I6IHZhcigtLXN0YXR1cy1nb29kKTsgfQouYmFkZ2UucGFydGlhbCB7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zdGF0dXMtd2FybmluZykgMjUlLCB0cmFuc3BhcmVudCk7IGNvbG9yOiAjOGE2MzAwOyB9Ci5iYWRnZS5mYWlsZWQgeyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc3RhdHVzLWNyaXRpY2FsKSAxOCUsIHRyYW5zcGFyZW50KTsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IH0KLmJhZGdlLmVycm9yLXNldiB7IGNvbG9yOiB2YXIoLS1zdGF0dXMtY3JpdGljYWwpOyB9Ci5iYWRnZS53YXJuaW5nLXNldiB7IGNvbG9yOiAjOGE2MzAwOyB9Ci5iYWRnZS5za2lwLXNldiB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgfQoKLmlzc3Vlcy1saXN0IHsgbWF4LWhlaWdodDogMjIwcHg7IG92ZXJmbG93LXk6IGF1dG87IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1zbSk7IH0KLmlzc3VlLXJvdyB7IHBhZGRpbmc6IDlweCAxNHB4OyBib3JkZXItYm90dG9tOiAxcHggc29saWQgdmFyKC0tZ3JpZGxpbmUpOyBmb250LXNpemU6IDEycHg7IH0KLmlzc3VlLXJvdzpsYXN0LWNoaWxkIHsgYm9yZGVyLWJvdHRvbTogbm9uZTsgfQouaXNzdWUtcm93IC5yb3ctbm8geyBjb2xvcjogdmFyKC0tdGV4dC1tdXRlZCk7IG1hcmdpbi1yaWdodDogNnB4OyB9CgovKiAtLS0tLS0tLS0tIFRvYXN0IC0tLS0tLS0tLS0gKi8KLnRvYXN0LXJvb3QgeyBwb3NpdGlvbjogZml4ZWQ7IGJvdHRvbTogMjBweDsgcmlnaHQ6IDIwcHg7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGdhcDogOHB4OyB6LWluZGV4OiAxMDA7IH0KLnRvYXN0IHsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTEpOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtc20pOwogIGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBwYWRkaW5nOiAxMnB4IDE2cHg7IGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdy1tb2RhbCk7IGZvbnQtc2l6ZTogMTNweDsgbWF4LXdpZHRoOiAzNDBweDsKICBhbmltYXRpb246IHRvYXN0LWluIDIyMG1zIHZhcigtLWVhc2UpOwp9Ci50b2FzdC5zdWNjZXNzIHsgYm9yZGVyLWxlZnQ6IDNweCBzb2xpZCB2YXIoLS1zdGF0dXMtZ29vZCk7IH0KLnRvYXN0LmVycm9yIHsgYm9yZGVyLWxlZnQ6IDNweCBzb2xpZCB2YXIoLS1zdGF0dXMtY3JpdGljYWwpOyB9CkBrZXlmcmFtZXMgdG9hc3QtaW4geyBmcm9tIHsgb3BhY2l0eTogMDsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDEwcHgpIHNjYWxlKDAuOTgpOyB9IHRvIHsgb3BhY2l0eTogMTsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDApIHNjYWxlKDEpOyB9IH0KCi8qIC0tLS0tLS0tLS0gTWlzYyAtLS0tLS0tLS0tICovCi5tdXRlZCB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgfQouZW1wdHktc3RhdGUgewogIHBhZGRpbmc6IDU2cHggMjRweDsgdGV4dC1hbGlnbjogY2VudGVyOyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOwogIGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogMTJweDsKICBhbmltYXRpb246IGNhcmRJbiAyNjBtcyB2YXIoLS1lYXNlKTsKfQouZW1wdHktc3RhdGUgLmVtcHR5LWljb24gewogIHdpZHRoOiA1MnB4OyBoZWlnaHQ6IDUycHg7IGJvcmRlci1yYWRpdXM6IDE2cHg7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogY2VudGVyOwogIGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zZXJpZXMtMSkgMTAlLCB0cmFuc3BhcmVudCk7IGNvbG9yOiB2YXIoLS1zZXJpZXMtMSk7Cn0KLmVtcHR5LXN0YXRlIC5lbXB0eS10aXRsZSB7IGZvbnQtc2l6ZTogMTRweDsgZm9udC13ZWlnaHQ6IDYwMDsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IH0KLmVtcHR5LXN0YXRlIC5lbXB0eS1tZXNzYWdlIHsgZm9udC1zaXplOiAxM3B4OyBtYXgtd2lkdGg6IDM2MHB4OyB9Ci5zcGlubmVyIHsgd2lkdGg6IDE2cHg7IGhlaWdodDogMTZweDsgYm9yZGVyLXJhZGl1czogNTAlOyBib3JkZXI6IDJweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItdG9wLWNvbG9yOiB2YXIoLS1zZXJpZXMtMSk7IGFuaW1hdGlvbjogc3BpbiAuNnMgbGluZWFyIGluZmluaXRlOyBkaXNwbGF5OiBpbmxpbmUtYmxvY2s7IH0KQGtleWZyYW1lcyBzcGluIHsgdG8geyB0cmFuc2Zvcm06IHJvdGF0ZSgzNjBkZWcpOyB9IH0KLmxvYWRpbmctcm93IHsgcGFkZGluZzogNDBweCAyMHB4OyB0ZXh0LWFsaWduOiBjZW50ZXI7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IH0KCi8qIFNrZWxldG9uIGxvYWRlcnMg4oCUIHNoaW1tZXJpbmcgcGxhY2Vob2xkZXJzIHNob3duIHdoaWxlIGEgc2VjdGlvbidzIGRhdGEgaXMgaW4gZmxpZ2h0ICovCi5za2VsZXRvbiB7CiAgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLXNtKTsKICBiYWNrZ3JvdW5kOiBsaW5lYXItZ3JhZGllbnQoMTAwZGVnLCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tdGV4dC1tdXRlZCkgMTIlLCB0cmFuc3BhcmVudCkgMzAlLCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tdGV4dC1tdXRlZCkgMjIlLCB0cmFuc3BhcmVudCkgNTAlLCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tdGV4dC1tdXRlZCkgMTIlLCB0cmFuc3BhcmVudCkgNzAlKTsKICBiYWNrZ3JvdW5kLXNpemU6IDIwMCUgMTAwJTsKICBhbmltYXRpb246IHNrZWxldG9uU2hpbW1lciAxLjRzIGVhc2UtaW4tb3V0IGluZmluaXRlOwp9CkBrZXlmcmFtZXMgc2tlbGV0b25TaGltbWVyIHsgZnJvbSB7IGJhY2tncm91bmQtcG9zaXRpb246IDE1MCUgMDsgfSB0byB7IGJhY2tncm91bmQtcG9zaXRpb246IC01MCUgMDsgfSB9Ci5za2VsZXRvbi1zdGF0LWdyaWQgeyBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdChhdXRvLWZpdCwgbWlubWF4KDE4MHB4LCAxZnIpKTsgZ2FwOiAxNHB4OyB9Ci5za2VsZXRvbi10aWxlIHsgaGVpZ2h0OiA4NHB4OyB9Ci5za2VsZXRvbi1jaGFydCB7IGhlaWdodDogMjgwcHg7IHdpZHRoOiAxMDAlOyB9Ci5za2VsZXRvbi1yb3cgeyBoZWlnaHQ6IDQwcHg7IG1hcmdpbi1ib3R0b206IDhweDsgfQoKLnR3by1jb2wgeyBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IDFmciAxZnI7IGdhcDogMTZweDsgfQpAbWVkaWEgKG1heC13aWR0aDogOTAwcHgpIHsgLnR3by1jb2wgeyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IDFmcjsgfSB9CgoubW9kZS10YWJzIHsgZGlzcGxheTogZmxleDsgZ2FwOiA2cHg7IGZsZXgtd3JhcDogd3JhcDsgbWFyZ2luLWJvdHRvbTogMTZweDsgfQoubW9kZS10YWJzIGJ1dHRvbiB7CiAgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgcGFkZGluZzogN3B4IDE0cHg7IGJvcmRlci1yYWRpdXM6IDIwcHg7IGZvbnQtc2l6ZTogMTJweDsgZm9udC13ZWlnaHQ6IDYwMDsgY3Vyc29yOiBwb2ludGVyOwogIHRyYW5zaXRpb246IGNvbG9yIDE4MG1zIHZhcigtLWVhc2UpLCBiYWNrZ3JvdW5kIDE4MG1zIHZhcigtLWVhc2UpLCB0cmFuc2Zvcm0gMTUwbXMgdmFyKC0tZWFzZSk7Cn0KLm1vZGUtdGFicyBidXR0b246aG92ZXIgeyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoLTFweCk7IH0KLm1vZGUtdGFicyBidXR0b24uaXMtYWN0aXZlIHsgYmFja2dyb3VuZDogdmFyKC0tc2VyaWVzLTEpOyBjb2xvcjogI2ZmZjsgYm9yZGVyLWNvbG9yOiB0cmFuc3BhcmVudDsgYm94LXNoYWRvdzogMCA0cHggMTRweCAtNXB4IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zZXJpZXMtMSkgNjAlLCB0cmFuc3BhcmVudCk7IH0KCi5maWVsZC1pbmxpbmUgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDhweDsgZm9udC1zaXplOiAxMnB4OyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyB9Ci5maWVsZC1pbmxpbmUgc2VsZWN0LCAuZmllbGQtaW5saW5lIGlucHV0IHsgbWluLXdpZHRoOiAwOyBwYWRkaW5nOiA2cHggMTBweDsgfQoKLyogLS0tLS0tLS0tLSBQYWdpbmF0aW9uIC0tLS0tLS0tLS0gKi8KLnBhZ2luYXRpb24tcm93IHsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiAxMnB4OyBtYXJnaW4tdG9wOiAxNHB4OyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IH0KLnBhZ2luYXRpb24tcm93IC5idG4geyBwYWRkaW5nOiA2cHggMTJweDsgfQoKLyogLS0tLS0tLS0tLSBEYXNoYm9hcmQgY29udHJvbHMgLyBtZXRyaWMtZm9jdXNlZCBLUElzIC0tLS0tLS0tLS0gKi8KLmRhc2hib2FyZC1jb250cm9scyB7CiAgZGlzcGxheTogZmxleDsgZmxleC13cmFwOiB3cmFwOyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDEwcHg7IG1hcmdpbi1ib3R0b206IDE4cHg7Cn0KLmRhc2hib2FyZC1jb250cm9scyBsYWJlbCB7IGZvbnQtc2l6ZTogMTJweDsgZm9udC13ZWlnaHQ6IDYwMDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgbWFyZ2luLXJpZ2h0OiA2cHg7IH0KLmRhc2hib2FyZC1jb250cm9scyBzZWxlY3QgeyBmb250LXdlaWdodDogNjAwOyB9Ci5zdGF0LXRpbGUucG9zdC10aWxlIC5zdGF0LXZhbHVlIHsgZm9udC1zaXplOiAxNXB4OyBsaW5lLWhlaWdodDogMS4zNTsgbWFyZ2luLXRvcDogNnB4OyB9Ci5zdGF0LXRpbGUucG9zdC10aWxlIC5wb3N0LW1ldGEgeyBmb250LXNpemU6IDExcHg7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgbWFyZ2luLXRvcDogNHB4OyB9Ci5zdGF0LXRpbGUucG9zdC10aWxlIC5wb3N0LW1ldHJpYy12YWx1ZSB7IGZvbnQtc2l6ZTogMjBweDsgZm9udC13ZWlnaHQ6IDcwMDsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IG1hcmdpbi10b3A6IDZweDsgfQoKLyogLS0tLS0tLS0tLSBEYXRhIFJlY29yZHMgKHBsYXRmb3JtLWdyb3VwZWQpIC0tLS0tLS0tLS0gKi8KLnJlY29yZHMtdG9vbGJhciB7CiAgZGlzcGxheTogZmxleDsgZmxleC13cmFwOiB3cmFwOyBhbGlnbi1pdGVtczogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6IHNwYWNlLWJldHdlZW47CiAgZ2FwOiAxMnB4OyBtYXJnaW4tYm90dG9tOiAxNHB4Owp9Ci5wbGF0Zm9ybS1maWx0ZXItcGlsbHMgeyBkaXNwbGF5OiBmbGV4OyBmbGV4LXdyYXA6IHdyYXA7IGdhcDogNnB4OyB9Ci5wbGF0Zm9ybS1maWx0ZXItcGlsbHMgYnV0dG9uIHsKICBkaXNwbGF5OiBpbmxpbmUtZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA2cHg7CiAgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgcGFkZGluZzogN3B4IDE0cHg7IGJvcmRlci1yYWRpdXM6IDIwcHg7IGZvbnQtc2l6ZTogMTJweDsgZm9udC13ZWlnaHQ6IDYwMDsgY3Vyc29yOiBwb2ludGVyOwogIHRyYW5zaXRpb246IGNvbG9yIDE4MG1zIHZhcigtLWVhc2UpLCBiYWNrZ3JvdW5kIDE4MG1zIHZhcigtLWVhc2UpLCB0cmFuc2Zvcm0gMTUwbXMgdmFyKC0tZWFzZSksIGJveC1zaGFkb3cgMTgwbXMgdmFyKC0tZWFzZSk7Cn0KLnBsYXRmb3JtLWZpbHRlci1waWxscyBidXR0b246aG92ZXIgeyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKC0xcHgpOyB9Ci5wbGF0Zm9ybS1maWx0ZXItcGlsbHMgYnV0dG9uOmFjdGl2ZSB7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgwKSBzY2FsZSgwLjk2KTsgfQoucGxhdGZvcm0tZmlsdGVyLXBpbGxzIGJ1dHRvbi5pcy1hY3RpdmUgeyBiYWNrZ3JvdW5kOiB2YXIoLS1zZXJpZXMtMSk7IGNvbG9yOiAjZmZmOyBib3JkZXItY29sb3I6IHRyYW5zcGFyZW50OyBib3gtc2hhZG93OiAwIDRweCAxNHB4IC01cHggY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSA2MCUsIHRyYW5zcGFyZW50KTsgfQoucGxhdGZvcm0tZmlsdGVyLXBpbGxzIGJ1dHRvbi5pcy1hY3RpdmUgLnBsYXRmb3JtLWRvdCB7IGJveC1zaGFkb3c6IDAgMCAwIDJweCByZ2JhKDI1NSwyNTUsMjU1LDAuNSk7IH0KLnJlY29yZHMtc2VhcmNoIGlucHV0IHsgYm9yZGVyLXJhZGl1czogMjBweDsgbWluLXdpZHRoOiAyMjBweDsgfQouc3RhdHVzLXBpbGwgeyBkaXNwbGF5OiBpbmxpbmUtYmxvY2s7IHBhZGRpbmc6IDNweCAxMHB4OyBib3JkZXItcmFkaXVzOiAyMHB4OyBmb250LXNpemU6IDExcHg7IGZvbnQtd2VpZ2h0OiA3MDA7IH0KLnN0YXR1cy1waWxsLm9yaWdpbmFsIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXRleHQtbXV0ZWQpIDE1JSwgdHJhbnNwYXJlbnQpOyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyB9Ci5zdGF0dXMtcGlsbC5lZGl0ZWQgeyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc3RhdHVzLXdhcm5pbmcpIDIyJSwgdHJhbnNwYXJlbnQpOyBjb2xvcjogIzhhNjMwMDsgfQoucm93LWFjdGlvbnMgeyBkaXNwbGF5OiBmbGV4OyBnYXA6IDZweDsgZmxleC13cmFwOiBub3dyYXA7IH0KLnJvdy1hY3Rpb25zIC5idG4geyBwYWRkaW5nOiA1cHggMTBweDsgZm9udC1zaXplOiAxMnB4OyB9Ci5saW5rLWNlbGwgYSB7IGNvbG9yOiB2YXIoLS1zZXJpZXMtMSk7IHRleHQtZGVjb3JhdGlvbjogbm9uZTsgZm9udC13ZWlnaHQ6IDYwMDsgZm9udC1zaXplOiAxMnB4OyB9Ci5saW5rLWNlbGwgYTpob3ZlciB7IHRleHQtZGVjb3JhdGlvbjogdW5kZXJsaW5lOyB9Ci5yZWNvcmQtc2VjdGlvbiB7CiAgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLXNtKTsgcGFkZGluZzogMTZweDsgbWFyZ2luLWJvdHRvbTogMTRweDsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTIpOyBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7Cn0KLnJlY29yZC1zZWN0aW9uIGg0IHsgbWFyZ2luOiAwIDAgMTJweDsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNzAwOyBsZXR0ZXItc3BhY2luZzogMC4wM2VtOyB0ZXh0LXRyYW5zZm9ybTogdXBwZXJjYXNlOyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyB9Ci5yZWNvcmQtc2VjdGlvbiAuZm9ybS1ncmlkIHsgbWFyZ2luLWJvdHRvbTogMDsgfQoucmVjb3JkLXNlY3Rpb24gLnZpZXctZmllbGQgeyBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBnYXA6IDJweDsgZm9udC1zaXplOiAxM3B4OyB9Ci5yZWNvcmQtc2VjdGlvbiAudmlldy1maWVsZCAudmlldy1sYWJlbCB7IGZvbnQtc2l6ZTogMTFweDsgZm9udC13ZWlnaHQ6IDYwMDsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyB9Ci5yZWNvcmQtc2VjdGlvbiAudmlldy1maWVsZCAudmlldy12YWx1ZSB7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyB3b3JkLWJyZWFrOiBicmVhay13b3JkOyB9CkBtZWRpYSAobWF4LXdpZHRoOiA2NDBweCkgewogIC5yZWNvcmRzLXRvb2xiYXIgeyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBhbGlnbi1pdGVtczogc3RyZXRjaDsgfQogIC5yZWNvcmRzLXNlYXJjaCBpbnB1dCB7IHdpZHRoOiAxMDAlOyB9Cn0KCi8qIC0tLS0tLS0tLS0gTW9kYWwgKHJlY29yZCBlZGl0b3IpIC0tLS0tLS0tLS0gKi8KLm1vZGFsLW92ZXJsYXkgewogIHBvc2l0aW9uOiBmaXhlZDsgaW5zZXQ6IDA7IGJhY2tncm91bmQ6IHJnYmEoMTAsMTEsMTMsMC41KTsKICBiYWNrZHJvcC1maWx0ZXI6IGJsdXIoNnB4KTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IGJsdXIoNnB4KTsKICBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogZmxleC1zdGFydDsganVzdGlmeS1jb250ZW50OiBjZW50ZXI7CiAgcGFkZGluZzogNDBweCAxNnB4OyBvdmVyZmxvdy15OiBhdXRvOyB6LWluZGV4OiAyMDA7CiAgYW5pbWF0aW9uOiBvdmVybGF5SW4gMjAwbXMgdmFyKC0tZWFzZSk7Cn0KQGtleWZyYW1lcyBvdmVybGF5SW4geyBmcm9tIHsgb3BhY2l0eTogMDsgfSB0byB7IG9wYWNpdHk6IDE7IH0gfQpAa2V5ZnJhbWVzIG1vZGFsUGFuZWxJbiB7CiAgZnJvbSB7IG9wYWNpdHk6IDA7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgxNHB4KSBzY2FsZSgwLjk3KTsgfQogIHRvIHsgb3BhY2l0eTogMTsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDApIHNjYWxlKDEpOyB9Cn0KLm1vZGFsLXBhbmVsIHsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTEpOyBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtbGcpOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOwogIGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBwYWRkaW5nOiAyNHB4OyB3aWR0aDogMTAwJTsgbWF4LXdpZHRoOiA3MjBweDsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LW1vZGFsKTsKICBtYXgtaGVpZ2h0OiBjYWxjKDEwMHZoIC0gODBweCk7IG92ZXJmbG93LXk6IGF1dG87CiAgYW5pbWF0aW9uOiBtb2RhbFBhbmVsSW4gMjQwbXMgdmFyKC0tZWFzZSk7Cn0KLm1vZGFsLXBhbmVsLndpZGUgeyBtYXgtd2lkdGg6IDExMDBweDsgfQoubW9kYWwtcGFuZWwgaDIgeyBtYXJnaW46IDAgMCA0cHg7IGZvbnQtc2l6ZTogMTdweDsgbGV0dGVyLXNwYWNpbmc6IC0wLjAxZW07IH0KLm1vZGFsLXBhbmVsIC5tb2RhbC1zdWIgeyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyBmb250LXNpemU6IDEycHg7IG1hcmdpbjogMCAwIDE4cHg7IH0KLmZvcm0tZ3JpZCB7IGRpc3BsYXk6IGdyaWQ7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogcmVwZWF0KGF1dG8tZml0LCBtaW5tYXgoMjAwcHgsIDFmcikpOyBnYXA6IDEycHg7IG1hcmdpbi1ib3R0b206IDE2cHg7IH0KLmZvcm0tZ3JpZC5mdWxsIHsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiAxZnI7IH0KQG1lZGlhIChtYXgtd2lkdGg6IDY0MHB4KSB7IC5mb3JtLWdyaWQgeyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IDFmcjsgfSB9Ci5mb3JtLWZpZWxkIHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiA1cHg7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgfQouZm9ybS1maWVsZCBsYWJlbCB7IGZvbnQtd2VpZ2h0OiA2MDA7IH0KLmZvcm0tZmllbGQgdGV4dGFyZWEgeyByZXNpemU6IHZlcnRpY2FsOyBtaW4taGVpZ2h0OiA2MHB4OyB9CgoucGxhdGZvcm0tZWRpdC1yb3cgewogIGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1zbSk7IHBhZGRpbmc6IDE0cHg7IG1hcmdpbi1ib3R0b206IDEwcHg7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMik7Cn0KLnBsYXRmb3JtLWVkaXQtcm93IC5wbGF0Zm9ybS1lZGl0LWhlYWQgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6IHNwYWNlLWJldHdlZW47IGdhcDogOHB4OyBtYXJnaW4tYm90dG9tOiAxMHB4OyB9Ci5wbGF0Zm9ybS1lZGl0LXJvdyAubWV0cmljcy1ncmlkIHsgZGlzcGxheTogZ3JpZDsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiByZXBlYXQoYXV0by1maXQsIG1pbm1heCgxMjBweCwgMWZyKSk7IGdhcDogOHB4OyB9Ci5yZW1vdmUtcGxhdGZvcm0tYnRuIHsgYm9yZGVyOiBub25lOyBiYWNrZ3JvdW5kOiB0cmFuc3BhcmVudDsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IGN1cnNvcjogcG9pbnRlcjsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNjAwOyB0cmFuc2l0aW9uOiBvcGFjaXR5IDE1MG1zIHZhcigtLWVhc2UpOyB9Ci5yZW1vdmUtcGxhdGZvcm0tYnRuOmhvdmVyIHsgb3BhY2l0eTogMC43OyB9Ci5tb2RhbC1hY3Rpb25zIHsgZGlzcGxheTogZmxleDsganVzdGlmeS1jb250ZW50OiBzcGFjZS1iZXR3ZWVuOyBhbGlnbi1pdGVtczogY2VudGVyOyBtYXJnaW4tdG9wOiAxOHB4OyBnYXA6IDhweDsgZmxleC13cmFwOiB3cmFwOyB9CgovKiAtLS0tLS0tLS0tIFJlc3BvbnNpdmUgdGlnaHRlbmluZyAtLS0tLS0tLS0tICovCkBtZWRpYSAobWF4LXdpZHRoOiA3MjBweCkgewogIC50b3BiYXIgeyBnYXA6IDEycHg7IHBhZGRpbmc6IDEwcHggMTRweDsgZmxleC13cmFwOiB3cmFwOyB9CiAgLnRvcGJhci1icmFuZCB7IG9yZGVyOiAxOyB9CiAgLnRvcGJhci11c2VyIHsgb3JkZXI6IDI7IG1hcmdpbi1sZWZ0OiBhdXRvOyB9CiAgLnRhYnMgeyBvcmRlcjogMzsgd2lkdGg6IDEwMCU7IH0KICAudmlldy1hcmVhIHsgcGFkZGluZzogMTRweDsgfQogIC5maWx0ZXItYmFyIHsgdG9wOiBhdXRvOyBwb3NpdGlvbjogc3RhdGljOyBwYWRkaW5nOiAxMnB4IDE0cHg7IH0KICAuc3RhdC1ncmlkIHsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiByZXBlYXQoYXV0by1maXQsIG1pbm1heCgxNDBweCwgMWZyKSk7IH0KfQo8L3N0eWxlPgo8L2hlYWQ+Cjxib2R5Pgo8ZGl2IGNsYXNzPSJhdXRoLXNjcmVlbiIgaWQ9ImF1dGhTY3JlZW4iPgogIDxkaXYgY2xhc3M9ImF1dGgtY2FyZCI+CiAgICA8ZGl2IGNsYXNzPSJhdXRoLWJyYW5kIj4KICAgICAgPHNwYW4gY2xhc3M9ImJyYW5kLW1hcmsiPkxSUzwvc3Bhbj4KICAgICAgPHNwYW4gY2xhc3M9ImJyYW5kLXRpdGxlIj5Tb2NpYWwgTWVkaWEgQW5hbHl0aWNzPC9zcGFuPgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJhdXRoLWZvcm0iPgogICAgICA8ZGl2IGNsYXNzPSJmb3JtLWZpZWxkIj4KICAgICAgICA8bGFiZWwgZm9yPSJhdXRoQ29kZSI+QWNjZXNzIGNvZGU8L2xhYmVsPgogICAgICAgIDxpbnB1dCB0eXBlPSJwYXNzd29yZCIgaWQ9ImF1dGhDb2RlIiBhdXRvY29tcGxldGU9Im9mZiIgLz4KICAgICAgPC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImF1dGgtZXJyb3IiIGlkPSJhdXRoRXJyb3IiPjwvZGl2PgogICAgICA8YnV0dG9uIGNsYXNzPSJidG4gcHJpbWFyeSIgaWQ9ImF1dGhTdWJtaXRCdG4iIHR5cGU9ImJ1dHRvbiI+PGkgZGF0YS1sdWNpZGU9ImFycm93LXJpZ2h0IiBzdHlsZT0id2lkdGg6MTRweDtoZWlnaHQ6MTRweDsiPjwvaT4gRW50ZXI8L2J1dHRvbj4KICAgIDwvZGl2PgogIDwvZGl2Pgo8L2Rpdj4KCjxkaXYgY2xhc3M9ImFwcC1zaGVsbCIgaWQ9ImFwcFNoZWxsIiBzdHlsZT0iZGlzcGxheTpub25lOyI+CiAgPGhlYWRlciBjbGFzcz0idG9wYmFyIj4KICAgIDxkaXYgY2xhc3M9InRvcGJhci1icmFuZCI+CiAgICAgIDxzcGFuIGNsYXNzPSJicmFuZC1tYXJrIj5MUlM8L3NwYW4+CiAgICAgIDxzcGFuIGNsYXNzPSJicmFuZC10aXRsZSI+U29jaWFsIE1lZGlhIEFuYWx5dGljczwvc3Bhbj4KICAgIDwvZGl2PgogICAgPG5hdiBjbGFzcz0idGFicyIgcm9sZT0idGFibGlzdCIgYXJpYS1sYWJlbD0iU2VjdGlvbnMiPgogICAgICA8YnV0dG9uIGNsYXNzPSJ0YWItYnRuIGlzLWFjdGl2ZSIgZGF0YS10YWI9ImRhc2hib2FyZCIgcm9sZT0idGFiIiBhcmlhLXNlbGVjdGVkPSJ0cnVlIj48aSBkYXRhLWx1Y2lkZT0ibGF5b3V0LWRhc2hib2FyZCIgc3R5bGU9IndpZHRoOjE0cHg7aGVpZ2h0OjE0cHg7Ij48L2k+IERhc2hib2FyZDwvYnV0dG9uPgogICAgICA8YnV0dG9uIGNsYXNzPSJ0YWItYnRuIiBkYXRhLXRhYj0icmVjb3JkcyIgcm9sZT0idGFiIiBhcmlhLXNlbGVjdGVkPSJmYWxzZSI+PGkgZGF0YS1sdWNpZGU9ImRhdGFiYXNlIiBzdHlsZT0id2lkdGg6MTRweDtoZWlnaHQ6MTRweDsiPjwvaT4gRGF0YSBSZWNvcmRzPC9idXR0b24+CiAgICAgIDxidXR0b24gY2xhc3M9InRhYi1idG4iIGRhdGEtdGFiPSJjb21wYXJpc29uIiByb2xlPSJ0YWIiIGFyaWEtc2VsZWN0ZWQ9ImZhbHNlIj48aSBkYXRhLWx1Y2lkZT0iZ2l0LWNvbXBhcmUiIHN0eWxlPSJ3aWR0aDoxNHB4O2hlaWdodDoxNHB4OyI+PC9pPiBDb21wYXJpc29uczwvYnV0dG9uPgogICAgICA8YnV0dG9uIGNsYXNzPSJ0YWItYnRuIiBkYXRhLXRhYj0idXBsb2FkIiByb2xlPSJ0YWIiIGFyaWEtc2VsZWN0ZWQ9ImZhbHNlIj48aSBkYXRhLWx1Y2lkZT0idXBsb2FkLWNsb3VkIiBzdHlsZT0id2lkdGg6MTRweDtoZWlnaHQ6MTRweDsiPjwvaT4gVXBsb2FkIERhdGE8L2J1dHRvbj4KICAgICAgPGJ1dHRvbiBjbGFzcz0idGFiLWJ0biIgZGF0YS10YWI9Imhpc3RvcnkiIHJvbGU9InRhYiIgYXJpYS1zZWxlY3RlZD0iZmFsc2UiPjxpIGRhdGEtbHVjaWRlPSJoaXN0b3J5IiBzdHlsZT0id2lkdGg6MTRweDtoZWlnaHQ6MTRweDsiPjwvaT4gVXBsb2FkIEhpc3Rvcnk8L2J1dHRvbj4KICAgIDwvbmF2PgogICAgPGRpdiBjbGFzcz0idG9wYmFyLXVzZXIiPgogICAgICA8YnV0dG9uIGNsYXNzPSJidG4iIGlkPSJsb2dvdXRCdG4iIHR5cGU9ImJ1dHRvbiI+PGkgZGF0YS1sdWNpZGU9ImxvY2siIHN0eWxlPSJ3aWR0aDoxNHB4O2hlaWdodDoxNHB4OyI+PC9pPiBMb2NrPC9idXR0b24+CiAgICAgIDxidXR0b24gY2xhc3M9InRoZW1lLXRvZ2dsZSIgaWQ9InRoZW1lVG9nZ2xlIiB0eXBlPSJidXR0b24iIGFyaWEtbGFiZWw9IlRvZ2dsZSBkYXJrIG1vZGUiPgogICAgICAgIDxzcGFuIGlkPSJ0aGVtZVRvZ2dsZUljb24iPjxpIGRhdGEtbHVjaWRlPSJtb29uIiBzdHlsZT0id2lkdGg6MTZweDtoZWlnaHQ6MTZweDsiPjwvaT48L3NwYW4+CiAgICAgIDwvYnV0dG9uPgogICAgPC9kaXY+CiAgPC9oZWFkZXI+CgogIDxzZWN0aW9uIGNsYXNzPSJmaWx0ZXItYmFyIiBpZD0iZmlsdGVyQmFyIj4KICAgIDxkaXYgY2xhc3M9ImZpbHRlci1maWVsZCI+CiAgICAgIDxsYWJlbCBmb3I9ImZpbHRlckRhdGVGcm9tIj5Gcm9tPC9sYWJlbD4KICAgICAgPGlucHV0IHR5cGU9ImRhdGUiIGlkPSJmaWx0ZXJEYXRlRnJvbSIgLz4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0iZmlsdGVyLWZpZWxkIj4KICAgICAgPGxhYmVsIGZvcj0iZmlsdGVyRGF0ZVRvIj5UbzwvbGFiZWw+CiAgICAgIDxpbnB1dCB0eXBlPSJkYXRlIiBpZD0iZmlsdGVyRGF0ZVRvIiAvPgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJmaWx0ZXItZmllbGQgZmlsdGVyLXByZXNldHMiIGlkPSJmaWx0ZXJQcmVzZXRzIj4KICAgICAgPGJ1dHRvbiB0eXBlPSJidXR0b24iIGRhdGEtcHJlc2V0PSI3Ij5MYXN0IDcgZGF5czwvYnV0dG9uPgogICAgICA8YnV0dG9uIHR5cGU9ImJ1dHRvbiIgZGF0YS1wcmVzZXQ9IjMwIj5MYXN0IDMwIGRheXM8L2J1dHRvbj4KICAgICAgPGJ1dHRvbiB0eXBlPSJidXR0b24iIGRhdGEtcHJlc2V0PSI5MCI+TGFzdCA5MCBkYXlzPC9idXR0b24+CiAgICAgIDxidXR0b24gdHlwZT0iYnV0dG9uIiBkYXRhLXByZXNldD0iYWxsIj5BbGwgdGltZTwvYnV0dG9uPgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJmaWx0ZXItZmllbGQiPgogICAgICA8bGFiZWwgZm9yPSJmaWx0ZXJQbGF0Zm9ybSI+UGxhdGZvcm08L2xhYmVsPgogICAgICA8c2VsZWN0IGlkPSJmaWx0ZXJQbGF0Zm9ybSI+PG9wdGlvbiB2YWx1ZT0iYWxsIj5BbGwgcGxhdGZvcm1zPC9vcHRpb24+PC9zZWxlY3Q+CiAgICA8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImZpbHRlci1maWVsZCI+CiAgICAgIDxsYWJlbCBmb3I9ImZpbHRlckNhbXBhaWduIj5DYW1wYWlnbjwvbGFiZWw+CiAgICAgIDxzZWxlY3QgaWQ9ImZpbHRlckNhbXBhaWduIj48b3B0aW9uIHZhbHVlPSJhbGwiPkFsbCBjYW1wYWlnbnM8L29wdGlvbj48L3NlbGVjdD4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0iZmlsdGVyLWZpZWxkIj4KICAgICAgPGxhYmVsIGZvcj0iZmlsdGVyQ29udGVudFR5cGUiPkNvbnRlbnQgdHlwZTwvbGFiZWw+CiAgICAgIDxzZWxlY3QgaWQ9ImZpbHRlckNvbnRlbnRUeXBlIj48b3B0aW9uIHZhbHVlPSJhbGwiPkFsbCBjb250ZW50IHR5cGVzPC9vcHRpb24+PC9zZWxlY3Q+CiAgICA8L2Rpdj4KICA8L3NlY3Rpb24+CgogIDxtYWluIGNsYXNzPSJ2aWV3LWFyZWEiPgogICAgPHNlY3Rpb24gaWQ9InZpZXctZGFzaGJvYXJkIiBjbGFzcz0idmlldyBpcy1hY3RpdmUiPjwvc2VjdGlvbj4KICAgIDxzZWN0aW9uIGlkPSJ2aWV3LXJlY29yZHMiIGNsYXNzPSJ2aWV3Ij48L3NlY3Rpb24+CiAgICA8c2VjdGlvbiBpZD0idmlldy1jb21wYXJpc29uIiBjbGFzcz0idmlldyI+PC9zZWN0aW9uPgogICAgPHNlY3Rpb24gaWQ9InZpZXctdXBsb2FkIiBjbGFzcz0idmlldyI+PC9zZWN0aW9uPgogICAgPHNlY3Rpb24gaWQ9InZpZXctaGlzdG9yeSIgY2xhc3M9InZpZXciPjwvc2VjdGlvbj4KICA8L21haW4+CjwvZGl2PgoKPGRpdiBpZD0idG9hc3RSb290IiBjbGFzcz0idG9hc3Qtcm9vdCIgYXJpYS1saXZlPSJwb2xpdGUiPjwvZGl2PgoKPHNjcmlwdD4KLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIEFwaSDigJQgdGhpbiBmZXRjaCB3cmFwcGVycyBhcm91bmQgdGhlIFJFU1QgQVBJLgogICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KY29uc3QgQXBpID0gKCgpID0+IHsKICBhc3luYyBmdW5jdGlvbiByZXF1ZXN0KHBhdGgsIG9wdGlvbnMpIHsKICAgIGNvbnN0IHJlcyA9IGF3YWl0IGZldGNoKHBhdGgsIG9wdGlvbnMpOwogICAgbGV0IGJvZHk7CiAgICB0cnkgewogICAgICBib2R5ID0gYXdhaXQgcmVzLmpzb24oKTsKICAgIH0gY2F0Y2ggewogICAgICBib2R5ID0gbnVsbDsKICAgIH0KICAgIGlmIChyZXMuc3RhdHVzID09PSA0MDEgJiYgIXBhdGguc3RhcnRzV2l0aCgnL2FwaS9hdXRoLycpKSB7CiAgICAgIHdpbmRvdy5kaXNwYXRjaEV2ZW50KG5ldyBDdXN0b21FdmVudCgnbHJzOnNpZ25lZC1vdXQnKSk7CiAgICB9CiAgICBpZiAoIXJlcy5vaykgewogICAgICBjb25zdCBtZXNzYWdlID0gKGJvZHkgJiYgYm9keS5lcnJvcikgfHwgYFJlcXVlc3QgZmFpbGVkICgke3Jlcy5zdGF0dXN9KWA7CiAgICAgIHRocm93IG5ldyBFcnJvcihtZXNzYWdlKTsKICAgIH0KICAgIHJldHVybiBib2R5OwogIH0KCiAgZnVuY3Rpb24gcXMocGFyYW1zKSB7CiAgICBjb25zdCB1c3AgPSBuZXcgVVJMU2VhcmNoUGFyYW1zKCk7CiAgICBPYmplY3QuZW50cmllcyhwYXJhbXMgfHwge30pLmZvckVhY2goKFtrLCB2XSkgPT4gewogICAgICBpZiAodiAhPT0gdW5kZWZpbmVkICYmIHYgIT09IG51bGwgJiYgdiAhPT0gJycpIHVzcC5zZXQoaywgdik7CiAgICB9KTsKICAgIGNvbnN0IHMgPSB1c3AudG9TdHJpbmcoKTsKICAgIHJldHVybiBzID8gYD8ke3N9YCA6ICcnOwogIH0KCiAgcmV0dXJuIHsKICAgIGF1dGhNZTogKCkgPT4gcmVxdWVzdCgnL2FwaS9hdXRoL21lJyksCiAgICBhdXRoTG9naW46IChjb2RlKSA9PgogICAgICByZXF1ZXN0KCcvYXBpL2F1dGgvbG9naW4nLCB7IG1ldGhvZDogJ1BPU1QnLCBoZWFkZXJzOiB7ICdDb250ZW50LVR5cGUnOiAnYXBwbGljYXRpb24vanNvbicgfSwgYm9keTogSlNPTi5zdHJpbmdpZnkoeyBjb2RlIH0pIH0pLAogICAgYXV0aExvZ291dDogKCkgPT4gcmVxdWVzdCgnL2FwaS9hdXRoL2xvZ291dCcsIHsgbWV0aG9kOiAnUE9TVCcgfSksCgogICAgZmlsdGVyT3B0aW9uczogKCkgPT4gcmVxdWVzdCgnL2FwaS9hbmFseXRpY3MvZmlsdGVyLW9wdGlvbnMnKSwKICAgIGtwaXM6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL2twaXMke3FzKHBhcmFtcyl9YCksCiAgICBwbGF0Zm9ybUJyZWFrZG93bjogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9hbmFseXRpY3MvcGxhdGZvcm0tYnJlYWtkb3duJHtxcyhwYXJhbXMpfWApLAogICAgY2FtcGFpZ25CcmVha2Rvd246IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL2NhbXBhaWduLWJyZWFrZG93biR7cXMocGFyYW1zKX1gKSwKICAgIGNvbnRlbnRUeXBlQnJlYWtkb3duOiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9jb250ZW50LXR5cGUtYnJlYWtkb3duJHtxcyhwYXJhbXMpfWApLAogICAgbWV0cmljT3B0aW9uczogKHBsYXRmb3JtKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9tZXRyaWMtb3B0aW9ucyR7cXMoeyBwbGF0Zm9ybSB9KX1gKSwKICAgIG1ldHJpY1N1bW1hcnk6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL21ldHJpYy1zdW1tYXJ5JHtxcyhwYXJhbXMpfWApLAogICAgdHJlbmQ6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL3RyZW5kJHtxcyhwYXJhbXMpfWApLAogICAgdG9wUG9zdHM6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL3RvcC1wb3N0cyR7cXMocGFyYW1zKX1gKSwKICAgIGNvbXBhcmU6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL2NvbXBhcmUke3FzKHBhcmFtcyl9YCksCiAgICBtb250aGx5OiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9tb250aGx5JHtxcyhwYXJhbXMpfWApLAogICAgcXVhcnRlcmx5OiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9xdWFydGVybHkke3FzKHBhcmFtcyl9YCksCiAgICB5dGQ6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL3l0ZCR7cXMocGFyYW1zKX1gKSwKCiAgICBwcmV2aWV3VXBsb2FkOiAoZmlsZSkgPT4gewogICAgICBjb25zdCBmb3JtID0gbmV3IEZvcm1EYXRhKCk7CiAgICAgIGZvcm0uYXBwZW5kKCdmaWxlJywgZmlsZSk7CiAgICAgIHJldHVybiByZXF1ZXN0KCcvYXBpL3VwbG9hZHMvcHJldmlldycsIHsgbWV0aG9kOiAnUE9TVCcsIGJvZHk6IGZvcm0gfSk7CiAgICB9LAogICAgY29tbWl0VXBsb2FkOiAocGF5bG9hZCkgPT4KICAgICAgcmVxdWVzdCgnL2FwaS91cGxvYWRzL2NvbW1pdCcsIHsKICAgICAgICBtZXRob2Q6ICdQT1NUJywKICAgICAgICBoZWFkZXJzOiB7ICdDb250ZW50LVR5cGUnOiAnYXBwbGljYXRpb24vanNvbicgfSwKICAgICAgICBib2R5OiBKU09OLnN0cmluZ2lmeShwYXlsb2FkKSwKICAgICAgfSksCiAgICB1cGxvYWRIaXN0b3J5OiAoKSA9PiByZXF1ZXN0KCcvYXBpL3VwbG9hZHMvaGlzdG9yeScpLAogICAgdXBsb2FkRXJyb3JzOiAoaWQpID0+IHJlcXVlc3QoYC9hcGkvdXBsb2Fkcy8ke2lkfS9lcnJvcnNgKSwKICAgIHVwbG9hZFJhd1Jvd3M6IChpZCkgPT4gcmVxdWVzdChgL2FwaS91cGxvYWRzLyR7aWR9L3Jhdy1yb3dzYCksCgogICAgbGlzdFJlY29yZHM6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvcmVjb3JkcyR7cXMocGFyYW1zKX1gKSwKICAgIHJlY29yZHNUYWJsZTogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9yZWNvcmRzL3RhYmxlJHtxcyhwYXJhbXMpfWApLAogICAgZ2V0UmVjb3JkOiAoaWQpID0+IHJlcXVlc3QoYC9hcGkvcmVjb3Jkcy8ke2lkfWApLAogICAgdXBkYXRlUmVjb3JkOiAoaWQsIHZhbHVlcykgPT4KICAgICAgcmVxdWVzdChgL2FwaS9yZWNvcmRzLyR7aWR9YCwgewogICAgICAgIG1ldGhvZDogJ1BVVCcsCiAgICAgICAgaGVhZGVyczogeyAnQ29udGVudC1UeXBlJzogJ2FwcGxpY2F0aW9uL2pzb24nIH0sCiAgICAgICAgYm9keTogSlNPTi5zdHJpbmdpZnkoeyB2YWx1ZXMgfSksCiAgICAgIH0pLAogICAgZGVsZXRlUmVjb3JkUG9zdDogKHBvc3RJZCkgPT4gcmVxdWVzdChgL2FwaS9yZWNvcmRzL3Bvc3QvJHtwb3N0SWR9YCwgeyBtZXRob2Q6ICdERUxFVEUnIH0pLAogICAgZGVsZXRlUmVjb3JkUGxhdGZvcm06IChwb3N0SWQsIHBsYXRmb3JtKSA9PgogICAgICByZXF1ZXN0KGAvYXBpL3JlY29yZHMvcG9zdC8ke3Bvc3RJZH0vcGxhdGZvcm0vJHtwbGF0Zm9ybX1gLCB7IG1ldGhvZDogJ0RFTEVURScgfSksCgogICAgcmVzdG9yZUJhY2t1cDogKGZvcm0pID0+IHJlcXVlc3QoJy9hcGkvYmFja3VwL3Jlc3RvcmUnLCB7IG1ldGhvZDogJ1BPU1QnLCBib2R5OiBmb3JtIH0pLAogIH07Cn0pKCk7CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgU3RhdGUgLyBGb3JtYXQgLyBUb2FzdCDigJQgc2hhcmVkIGFwcCBzdGF0ZSArIHNtYWxsIHV0aWxpdGllcy4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmNvbnN0IFN0YXRlID0gKCgpID0+IHsKICBjb25zdCB0b2RheSA9IG5ldyBEYXRlKCk7CiAgY29uc3QgaXNvID0gKGQpID0+IGQudG9JU09TdHJpbmcoKS5zbGljZSgwLCAxMCk7CiAgY29uc3QgdGhpcnR5RGF5c0FnbyA9IG5ldyBEYXRlKHRvZGF5KTsKICB0aGlydHlEYXlzQWdvLnNldERhdGUodGhpcnR5RGF5c0Fnby5nZXREYXRlKCkgLSAyOSk7CgogIGNvbnN0IGZpbHRlcnMgPSB7CiAgICBkYXRlRnJvbTogaXNvKHRoaXJ0eURheXNBZ28pLAogICAgZGF0ZVRvOiBpc28odG9kYXkpLAogICAgcGxhdGZvcm06ICdhbGwnLAogICAgY2FtcGFpZ25UeXBlOiAnYWxsJywKICAgIGNvbnRlbnRUeXBlOiAnYWxsJywKICB9OwoKICBjb25zdCBsaXN0ZW5lcnMgPSBbXTsKCiAgcmV0dXJuIHsKICAgIGdldEZpbHRlcnM6ICgpID0+ICh7IC4uLmZpbHRlcnMgfSksCiAgICBzZXRGaWx0ZXJzKHBhcnRpYWwpIHsKICAgICAgT2JqZWN0LmFzc2lnbihmaWx0ZXJzLCBwYXJ0aWFsKTsKICAgICAgbGlzdGVuZXJzLmZvckVhY2goKGZuKSA9PiBmbih0aGlzLmdldEZpbHRlcnMoKSkpOwogICAgfSwKICAgIG9uQ2hhbmdlKGZuKSB7CiAgICAgIGxpc3RlbmVycy5wdXNoKGZuKTsKICAgIH0sCiAgfTsKfSkoKTsKCmNvbnN0IEZvcm1hdCA9IHsKICBudW1iZXIobikgewogICAgaWYgKG4gPT09IG51bGwgfHwgbiA9PT0gdW5kZWZpbmVkKSByZXR1cm4gJ+KAlCc7CiAgICByZXR1cm4gTWF0aC5yb3VuZChuKS50b0xvY2FsZVN0cmluZygnZW4tVVMnKTsKICB9LAogIGNvbXBhY3QobikgewogICAgaWYgKG4gPT09IG51bGwgfHwgbiA9PT0gdW5kZWZpbmVkKSByZXR1cm4gJ+KAlCc7CiAgICBjb25zdCBhYnMgPSBNYXRoLmFicyhuKTsKICAgIGlmIChhYnMgPj0gMV8wMDBfMDAwKSByZXR1cm4gYCR7KG4gLyAxXzAwMF8wMDApLnRvRml4ZWQoMSkucmVwbGFjZSgvXC4wJC8sICcnKX1NYDsKICAgIGlmIChhYnMgPj0gMV8wMDApIHJldHVybiBgJHsobiAvIDFfMDAwKS50b0ZpeGVkKDEpLnJlcGxhY2UoL1wuMCQvLCAnJyl9S2A7CiAgICByZXR1cm4gYCR7TWF0aC5yb3VuZChuKX1gOwogIH0sCiAgLyoqIERhc2hib2FyZC13aWRlICJwcm9mZXNzaW9uYWwiIG51bWJlciBmb3JtYXQ6IHBsYWluIHVuZGVyIDEsMDAwOyBjb21tYS1ncm91cGVkCiAgICAgIHVwIHRvIDEwLDAwMDsgYWJicmV2aWF0ZWQgKEsvTSkgYmV5b25kIHRoYXQg4oCUIGUuZy4gODUwLCAxLDI1MCwgMTIuNUssIDE1NkssIDEuMjVNLiAqLwogIHNtYXJ0KG4pIHsKICAgIGlmIChuID09PSBudWxsIHx8IG4gPT09IHVuZGVmaW5lZCkgcmV0dXJuICfigJQnOwogICAgY29uc3QgYWJzID0gTWF0aC5hYnMobik7CiAgICBpZiAoYWJzIDwgMTAwMCkgcmV0dXJuIGAke01hdGgucm91bmQobil9YDsKICAgIGlmIChhYnMgPCAxMDAwMCkgcmV0dXJuIE1hdGgucm91bmQobikudG9Mb2NhbGVTdHJpbmcoJ2VuLVVTJyk7CiAgICBpZiAoYWJzIDwgMV8wMDBfMDAwKSByZXR1cm4gYCR7KG4gLyAxMDAwKS50b0ZpeGVkKDEpLnJlcGxhY2UoL1wuMCQvLCAnJyl9S2A7CiAgICByZXR1cm4gYCR7KG4gLyAxXzAwMF8wMDApLnRvRml4ZWQoMikucmVwbGFjZSgvXC4/MCskLywgJycpfU1gOwogIH0sCiAgcGVyY2VudChuKSB7CiAgICBpZiAobiA9PT0gbnVsbCB8fCBuID09PSB1bmRlZmluZWQpIHJldHVybiAn4oCUJzsKICAgIHJldHVybiBgJHtOdW1iZXIobikudG9GaXhlZCgxKS5yZXBsYWNlKC9cLjAkLywgJycpfSVgOwogIH0sCiAgcGN0KG4pIHsKICAgIGlmIChuID09PSBudWxsIHx8IG4gPT09IHVuZGVmaW5lZCkgcmV0dXJuICfigJQnOwogICAgY29uc3Qgc2lnbiA9IG4gPiAwID8gJysnIDogJyc7CiAgICByZXR1cm4gYCR7c2lnbn0ke24udG9GaXhlZCgxKX0lYDsKICB9LAogIGRhdGUoaXNvXykgewogICAgaWYgKCFpc29fKSByZXR1cm4gJ+KAlCc7CiAgICBjb25zdCBbeSwgbSwgZF0gPSBpc29fLnNwbGl0KCctJykubWFwKE51bWJlcik7CiAgICByZXR1cm4gbmV3IERhdGUoeSwgbSAtIDEsIGQpLnRvTG9jYWxlRGF0ZVN0cmluZygnZW4tVVMnLCB7IG1vbnRoOiAnc2hvcnQnLCBkYXk6ICdudW1lcmljJywgeWVhcjogJ251bWVyaWMnIH0pOwogIH0sCiAgZHVyYXRpb24oc2Vjb25kcykgewogICAgaWYgKHNlY29uZHMgPT09IG51bGwgfHwgc2Vjb25kcyA9PT0gdW5kZWZpbmVkKSByZXR1cm4gJ+KAlCc7CiAgICBjb25zdCBzID0gTWF0aC5yb3VuZChzZWNvbmRzKTsKICAgIGlmIChzIDwgNjApIHJldHVybiBgJHtzfXNgOwogICAgaWYgKHMgPCAzNjAwKSByZXR1cm4gYCR7TWF0aC5mbG9vcihzIC8gNjApfW0gJHtzICUgNjB9c2A7CiAgICBjb25zdCBoID0gTWF0aC5mbG9vcihzIC8gMzYwMCk7CiAgICBjb25zdCBtID0gTWF0aC5yb3VuZCgocyAlIDM2MDApIC8gNjApOwogICAgcmV0dXJuIGAke2h9aCAke219bWA7CiAgfSwKICBkZWx0YUNsYXNzKG4pIHsKICAgIGlmIChuID09PSBudWxsIHx8IG4gPT09IHVuZGVmaW5lZCkgcmV0dXJuICdmbGF0JzsKICAgIGlmIChuID4gMC41KSByZXR1cm4gJ3VwJzsKICAgIGlmIChuIDwgLTAuNSkgcmV0dXJuICdkb3duJzsKICAgIHJldHVybiAnZmxhdCc7CiAgfSwKfTsKCmNvbnN0IFRvYXN0ID0gewogIHNob3cobWVzc2FnZSwgdHlwZSA9ICdzdWNjZXNzJykgewogICAgY29uc3Qgcm9vdCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd0b2FzdFJvb3QnKTsKICAgIGNvbnN0IGVsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBlbC5jbGFzc05hbWUgPSBgdG9hc3QgJHt0eXBlfWA7CiAgICBlbC50ZXh0Q29udGVudCA9IG1lc3NhZ2U7CiAgICByb290LmFwcGVuZENoaWxkKGVsKTsKICAgIHNldFRpbWVvdXQoKCkgPT4gZWwucmVtb3ZlKCksIDUwMDApOwogIH0sCn07CgovKiogU2FmZWx5IGJ1aWxkcyBET00gdGV4dCBub2RlcyBmb3IgdW50cnVzdGVkIHN0cmluZ3MgKGNhcHRpb25zLCBmaWxlbmFtZXMsIHBsYXRmb3JtIGxhYmVscyBmcm9tIGRhdGEpLiAqLwpmdW5jdGlvbiB0ZXh0RWwodGFnLCB0ZXh0LCBjbGFzc05hbWUpIHsKICBjb25zdCBlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQodGFnKTsKICBpZiAoY2xhc3NOYW1lKSBlbC5jbGFzc05hbWUgPSBjbGFzc05hbWU7CiAgZWwuYXBwZW5kQ2hpbGQoZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUodGV4dCA/PyAnJykpOwogIHJldHVybiBlbDsKfQoKLyoqIEEgcHJlbWl1bSBlbXB0eSBzdGF0ZTogaWNvbiArIGV4cGxhbmF0aW9uICsgb3B0aW9uYWwgYWN0aW9uLCBpbnN0ZWFkIG9mIGEgYmxhbmsgYXJlYS4KICAgIEljb25zIHJlbmRlciB2aWEgdGhlIHBhZ2Utd2lkZSBNdXRhdGlvbk9ic2VydmVyIHRoYXQgY2FsbHMgbHVjaWRlLmNyZWF0ZUljb25zKCkgKHNlZSBib290c3RyYXApLiAqLwpmdW5jdGlvbiBlbXB0eVN0YXRlKHsgaWNvbiA9ICdpbmJveCcsIHRpdGxlLCBtZXNzYWdlLCBhY3Rpb25MYWJlbCwgb25BY3Rpb24gfSkgewogIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICB3cmFwLmNsYXNzTmFtZSA9ICdlbXB0eS1zdGF0ZSc7CiAgY29uc3QgaWNvbldyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICBpY29uV3JhcC5jbGFzc05hbWUgPSAnZW1wdHktaWNvbic7CiAgaWNvbldyYXAuaW5uZXJIVE1MID0gYDxpIGRhdGEtbHVjaWRlPSIke2ljb259IiBzdHlsZT0id2lkdGg6MjJweDtoZWlnaHQ6MjJweDsiPjwvaT5gOwogIHdyYXAuYXBwZW5kQ2hpbGQoaWNvbldyYXApOwogIGlmICh0aXRsZSkgd3JhcC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIHRpdGxlLCAnZW1wdHktdGl0bGUnKSk7CiAgaWYgKG1lc3NhZ2UpIHdyYXAuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBtZXNzYWdlLCAnZW1wdHktbWVzc2FnZScpKTsKICBpZiAoYWN0aW9uTGFiZWwgJiYgb25BY3Rpb24pIHsKICAgIGNvbnN0IGJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgYnRuLmNsYXNzTmFtZSA9ICdidG4gcHJpbWFyeSc7CiAgICBidG4udGV4dENvbnRlbnQgPSBhY3Rpb25MYWJlbDsKICAgIGJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIG9uQWN0aW9uKTsKICAgIHdyYXAuYXBwZW5kQ2hpbGQoYnRuKTsKICB9CiAgcmV0dXJuIHdyYXA7Cn0KCi8qKiBBIDxidXR0b24+IHdpdGggYSBzbWFsbCBsZWFkaW5nIEx1Y2lkZSBpY29uIGJlZm9yZSBpdHMgbGFiZWwgKGxhYmVsIGlzIGFsd2F5cyBhIHN0YXRpYywgZGV2ZWxvcGVyLXN1cHBsaWVkIHN0cmluZyBhdCBjYWxsIHNpdGVzLCBuZXZlciB1c2VyIGRhdGEg4oCUIGluc2VydGVkIHZpYSBjcmVhdGVUZXh0Tm9kZSByZWdhcmRsZXNzKS4gKi8KZnVuY3Rpb24gaWNvbkJ0bihjbGFzc05hbWUsIGljb25OYW1lLCBsYWJlbCkgewogIGNvbnN0IGJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogIGJ0bi5jbGFzc05hbWUgPSBjbGFzc05hbWU7CiAgY29uc3QgaWNvbiA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2knKTsKICBpY29uLnNldEF0dHJpYnV0ZSgnZGF0YS1sdWNpZGUnLCBpY29uTmFtZSk7CiAgaWNvbi5zdHlsZS53aWR0aCA9ICcxM3B4JzsKICBpY29uLnN0eWxlLmhlaWdodCA9ICcxM3B4JzsKICBidG4uYXBwZW5kQ2hpbGQoaWNvbik7CiAgYnRuLmFwcGVuZENoaWxkKGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKGAgJHtsYWJlbH1gKSk7CiAgcmV0dXJuIGJ0bjsKfQoKLyoqIFNoaW1tZXJpbmcgcGxhY2Vob2xkZXJzIHNob3duIHRoZSBpbnN0YW50IGEgc2VjdGlvbiBzdGFydHMgbG9hZGluZywgc3dhcHBlZCBmb3IgcmVhbAogICAgY29udGVudCAob3IgYW4gZW1wdHkgc3RhdGUpIG9uY2UgdGhlIGZldGNoIHJlc29sdmVzIOKAlCBubyBibGFuayBhcmVhcyB3aGlsZSB3YWl0aW5nLiAqLwpmdW5jdGlvbiBza2VsZXRvblN0YXRHcmlkKGNvdW50ID0gNikgewogIGNvbnN0IGdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICBncmlkLmNsYXNzTmFtZSA9ICdza2VsZXRvbi1zdGF0LWdyaWQnOwogIGZvciAobGV0IGkgPSAwOyBpIDwgY291bnQ7IGkgKz0gMSkgewogICAgY29uc3QgdGlsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGlsZS5jbGFzc05hbWUgPSAnc2tlbGV0b24gc2tlbGV0b24tdGlsZSc7CiAgICBncmlkLmFwcGVuZENoaWxkKHRpbGUpOwogIH0KICByZXR1cm4gZ3JpZDsKfQpmdW5jdGlvbiBza2VsZXRvbkNoYXJ0KCkgewogIGNvbnN0IGRpdiA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogIGRpdi5jbGFzc05hbWUgPSAnc2tlbGV0b24gc2tlbGV0b24tY2hhcnQnOwogIHJldHVybiBkaXY7Cn0KZnVuY3Rpb24gc2tlbGV0b25Sb3dzKGNvdW50ID0gNikgewogIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICBmb3IgKGxldCBpID0gMDsgaSA8IGNvdW50OyBpICs9IDEpIHsKICAgIGNvbnN0IHJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcm93LmNsYXNzTmFtZSA9ICdza2VsZXRvbiBza2VsZXRvbi1yb3cnOwogICAgd3JhcC5hcHBlbmRDaGlsZChyb3cpOwogIH0KICByZXR1cm4gd3JhcDsKfQoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIENoYXJ0cyDigJQgQ2hhcnQuanMgYnVpbGRlcnMgKHZhbGlkYXRlZCBjYXRlZ29yaWNhbCBwYWxldHRlLAogICBoYWlybGluZSByZWNlc3NpdmUgZ3JpZGxpbmVzLCBzaW5nbGUgYXhpcywgbGVnZW5kIGFsd2F5cwogICBwcmVzZW50IGZvciAyKyBzZXJpZXMsIGluZGV4LW1vZGUgdG9vbHRpcHMpLgogICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KaWYgKHdpbmRvdy5DaGFydERhdGFMYWJlbHMpIENoYXJ0LnJlZ2lzdGVyKHdpbmRvdy5DaGFydERhdGFMYWJlbHMpOwoKY29uc3QgQ2hhcnRzID0gKCgpID0+IHsKICBjb25zdCByZWdpc3RyeSA9IG5ldyBNYXAoKTsgLy8gY2FudmFzSWQgLT4gQ2hhcnQgaW5zdGFuY2UsIHNvIHJlLXJlbmRlcnMgZGVzdHJveSB0aGUgb2xkIG9uZSBmaXJzdAogIGNvbnN0IE1BWF9MQUJFTEVEX0lURU1TID0gMjA7IC8vIGJleW9uZCB0aGlzLCBwZXItaXRlbSB2YWx1ZSBsYWJlbHMgd291bGQgb3ZlcmxhcCDigJQgcmVseSBvbiB0b29sdGlwcyBpbnN0ZWFkCgogIGZ1bmN0aW9uIGNzc1ZhcihuYW1lKSB7CiAgICByZXR1cm4gZ2V0Q29tcHV0ZWRTdHlsZShkb2N1bWVudC5kb2N1bWVudEVsZW1lbnQpLmdldFByb3BlcnR5VmFsdWUobmFtZSkudHJpbSgpOwogIH0KCiAgY29uc3QgU0VSSUVTX1ZBUlMgPSBbJy0tc2VyaWVzLTEnLCAnLS1zZXJpZXMtMicsICctLXNlcmllcy0zJywgJy0tc2VyaWVzLTQnLCAnLS1zZXJpZXMtNScsICctLXNlcmllcy02JywgJy0tc2VyaWVzLTcnLCAnLS1zZXJpZXMtOCddOwogIGZ1bmN0aW9uIHNlcmllc0NvbG9yKGluZGV4KSB7CiAgICByZXR1cm4gY3NzVmFyKFNFUklFU19WQVJTW2luZGV4ICUgU0VSSUVTX1ZBUlMubGVuZ3RoXSk7CiAgfQoKICBmdW5jdGlvbiBiYXNlR3JpZCgpIHsKICAgIHJldHVybiB7CiAgICAgIGNvbG9yOiBjc3NWYXIoJy0tZ3JpZGxpbmUnKSwKICAgICAgZHJhd1RpY2tzOiBmYWxzZSwKICAgIH07CiAgfQogIGZ1bmN0aW9uIGJhc2VUaWNrcygpIHsKICAgIHJldHVybiB7IGNvbG9yOiBjc3NWYXIoJy0tdGV4dC1tdXRlZCcpLCBmb250OiB7IHNpemU6IDExIH0gfTsKICB9CiAgZnVuY3Rpb24gYmFzZVRvb2x0aXAoKSB7CiAgICByZXR1cm4gewogICAgICBiYWNrZ3JvdW5kQ29sb3I6IGNzc1ZhcignLS1zdXJmYWNlLTEnKSwKICAgICAgdGl0bGVDb2xvcjogY3NzVmFyKCctLXRleHQtcHJpbWFyeScpLAogICAgICBib2R5Q29sb3I6IGNzc1ZhcignLS10ZXh0LXNlY29uZGFyeScpLAogICAgICBib3JkZXJDb2xvcjogY3NzVmFyKCctLWJvcmRlcicpLAogICAgICBib3JkZXJXaWR0aDogMSwKICAgICAgY29ybmVyUmFkaXVzOiAxMCwKICAgICAgcGFkZGluZzogMTIsCiAgICAgIGJveFBhZGRpbmc6IDQsCiAgICAgIHRpdGxlRm9udDogeyBzaXplOiAxMiwgd2VpZ2h0OiAnNzAwJyB9LAogICAgICBib2R5Rm9udDogeyBzaXplOiAxMiB9LAogICAgfTsKICB9CiAgZnVuY3Rpb24gbGFiZWxDb2xvcigpIHsKICAgIHJldHVybiBjc3NWYXIoJy0tdGV4dC1wcmltYXJ5Jyk7CiAgfQogIC8qKiBTbmFwcHksIHN1YnRsZSBtb3Rpb24g4oCUIGluIHRoZSAxNTAtMzAwbXMgcmFuZ2UgdGhlIHJlZGVzaWduIGNhbGxzIGZvciwgbmV2ZXIgYm91bmN5LiAqLwogIGZ1bmN0aW9uIGJhc2VBbmltYXRpb24oKSB7CiAgICByZXR1cm4geyBkdXJhdGlvbjogMjgwLCBlYXNpbmc6ICdlYXNlT3V0UXVhcnQnIH07CiAgfQoKICBmdW5jdGlvbiBkZXN0cm95KGNhbnZhc0lkKSB7CiAgICBpZiAocmVnaXN0cnkuaGFzKGNhbnZhc0lkKSkgewogICAgICByZWdpc3RyeS5nZXQoY2FudmFzSWQpLmRlc3Ryb3koKTsKICAgICAgcmVnaXN0cnkuZGVsZXRlKGNhbnZhc0lkKTsKICAgIH0KICB9CgogIC8qKiBNdWx0aS1zZXJpZXMgbGluZSBjaGFydCAoZS5nLiB3ZWVrbHkgdHJlbmQgcGVyIHBsYXRmb3JtKS4gT25lIHNlcmllcyBuZWVkcyBubyBsZWdlbmQgYm94LgogICAgICBQZXItcG9pbnQgdmFsdWUgbGFiZWxzIGFyZSBzaG93biBvbmx5IGZvciBhIHNpbmdsZSBzZXJpZXMg4oCUIHdpdGggc2V2ZXJhbCBzZXJpZXMgb3ZlcmxhaWQsCiAgICAgIGxhYmVsaW5nIGV2ZXJ5IHBvaW50IHdvdWxkIG92ZXJsYXAsIHNvIHRob3NlIHJlbHkgb24gdGhlIChzdGlsbC1wcmVzZW50KSBob3ZlciB0b29sdGlwLiAqLwogIGZ1bmN0aW9uIHRyZW5kQ2hhcnQoY2FudmFzSWQsIHsgbGFiZWxzLCBzZXJpZXMsIGZvcm1hdFZhbHVlIH0pIHsKICAgIGRlc3Ryb3koY2FudmFzSWQpOwogICAgY29uc3QgY3R4ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoY2FudmFzSWQpOwogICAgaWYgKCFjdHgpIHJldHVybiBudWxsOwogICAgY29uc3QgZm10ID0gZm9ybWF0VmFsdWUgfHwgKCh2KSA9PiBGb3JtYXQuc21hcnQodikpOwogICAgY29uc3Qgc2hvd0xhYmVscyA9IHNlcmllcy5sZW5ndGggPT09IDEgJiYgbGFiZWxzLmxlbmd0aCA8PSBNQVhfTEFCRUxFRF9JVEVNUzsKCiAgICBjb25zdCBkYXRhc2V0cyA9IHNlcmllcy5tYXAoKHMsIGkpID0+ICh7CiAgICAgIGxhYmVsOiBzLmxhYmVsLAogICAgICBkYXRhOiBzLmRhdGEsCiAgICAgIGJvcmRlckNvbG9yOiBzLmNvbG9yIHx8IHNlcmllc0NvbG9yKGkpLAogICAgICBiYWNrZ3JvdW5kQ29sb3I6IHMuY29sb3IgfHwgc2VyaWVzQ29sb3IoaSksCiAgICAgIGJvcmRlcldpZHRoOiAyLAogICAgICBwb2ludFJhZGl1czogc2hvd0xhYmVscyA/IDMgOiAwLAogICAgICBwb2ludEhvdmVyUmFkaXVzOiA0LAogICAgICBwb2ludEhpdFJhZGl1czogMTIsCiAgICAgIHRlbnNpb246IDAuMjUsCiAgICAgIGZpbGw6IGZhbHNlLAogICAgfSkpOwoKICAgIGNvbnN0IGNoYXJ0ID0gbmV3IENoYXJ0KGN0eCwgewogICAgICB0eXBlOiAnbGluZScsCiAgICAgIGRhdGE6IHsgbGFiZWxzLCBkYXRhc2V0cyB9LAogICAgICBvcHRpb25zOiB7CiAgICAgICAgcmVzcG9uc2l2ZTogdHJ1ZSwKICAgICAgICBtYWludGFpbkFzcGVjdFJhdGlvOiBmYWxzZSwKICAgICAgICBpbnRlcmFjdGlvbjogeyBtb2RlOiAnaW5kZXgnLCBpbnRlcnNlY3Q6IGZhbHNlIH0sCiAgICAgICAgbGF5b3V0OiB7IHBhZGRpbmc6IHsgdG9wOiBzaG93TGFiZWxzID8gMjAgOiA4IH0gfSwKICAgICAgICBhbmltYXRpb246IGJhc2VBbmltYXRpb24oKSwKICAgICAgICBwbHVnaW5zOiB7CiAgICAgICAgICBsZWdlbmQ6IHsKICAgICAgICAgICAgZGlzcGxheTogc2VyaWVzLmxlbmd0aCA+IDEsCiAgICAgICAgICAgIHBvc2l0aW9uOiAnYm90dG9tJywKICAgICAgICAgICAgbGFiZWxzOiB7IGNvbG9yOiBjc3NWYXIoJy0tdGV4dC1zZWNvbmRhcnknKSwgdXNlUG9pbnRTdHlsZTogdHJ1ZSwgcG9pbnRTdHlsZTogJ2xpbmUnLCBib3hXaWR0aDogMTYsIHBhZGRpbmc6IDE2LCBmb250OiB7IHNpemU6IDExIH0gfSwKICAgICAgICAgIH0sCiAgICAgICAgICB0b29sdGlwOiB7IC4uLmJhc2VUb29sdGlwKCksIHVzZVBvaW50U3R5bGU6IHRydWUgfSwKICAgICAgICAgIGRhdGFsYWJlbHM6IHNob3dMYWJlbHMKICAgICAgICAgICAgPyB7IGFsaWduOiAndG9wJywgYW5jaG9yOiAnZW5kJywgY29sb3I6IGxhYmVsQ29sb3IoKSwgZm9udDogeyBzaXplOiAxMSwgd2VpZ2h0OiAnNjAwJyB9LCBmb3JtYXR0ZXI6ICh2KSA9PiBmbXQodikgfQogICAgICAgICAgICA6IHsgZGlzcGxheTogZmFsc2UgfSwKICAgICAgICB9LAogICAgICAgIHNjYWxlczogewogICAgICAgICAgeDogeyBncmlkOiB7IGRpc3BsYXk6IGZhbHNlIH0sIHRpY2tzOiBiYXNlVGlja3MoKSB9LAogICAgICAgICAgeTogeyBncmlkOiBiYXNlR3JpZCgpLCB0aWNrczogYmFzZVRpY2tzKCksIGJvcmRlcjogeyBkaXNwbGF5OiBmYWxzZSB9LCBiZWdpbkF0WmVybzogdHJ1ZSB9LAogICAgICAgIH0sCiAgICAgIH0sCiAgICB9KTsKICAgIHJlZ2lzdHJ5LnNldChjYW52YXNJZCwgY2hhcnQpOwogICAgcmV0dXJuIGNoYXJ0OwogIH0KCiAgLyoqIFNpbmdsZS1tZXRyaWMgYmFyIGNoYXJ0IGFjcm9zcyBwbGF0Zm9ybXMgKGlkZW50aXR5IGVuY29kaW5nIOKAlCBlYWNoIGJhciBJUyBhIHBsYXRmb3JtKS4gKi8KICBmdW5jdGlvbiBwbGF0Zm9ybUJhckNoYXJ0KGNhbnZhc0lkLCB7IGxhYmVscywgZGF0YSwgY29sb3JzLCBmb3JtYXRWYWx1ZSB9KSB7CiAgICBkZXN0cm95KGNhbnZhc0lkKTsKICAgIGNvbnN0IGN0eCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGNhbnZhc0lkKTsKICAgIGlmICghY3R4KSByZXR1cm4gbnVsbDsKICAgIGNvbnN0IGZtdCA9IGZvcm1hdFZhbHVlIHx8ICgodikgPT4gRm9ybWF0LnNtYXJ0KHYpKTsKICAgIGNvbnN0IHNob3dMYWJlbHMgPSBsYWJlbHMubGVuZ3RoIDw9IE1BWF9MQUJFTEVEX0lURU1TOwoKICAgIGNvbnN0IGNoYXJ0ID0gbmV3IENoYXJ0KGN0eCwgewogICAgICB0eXBlOiAnYmFyJywKICAgICAgZGF0YTogewogICAgICAgIGxhYmVscywKICAgICAgICBkYXRhc2V0czogWwogICAgICAgICAgewogICAgICAgICAgICBkYXRhLAogICAgICAgICAgICBiYWNrZ3JvdW5kQ29sb3I6IGNvbG9ycywKICAgICAgICAgICAgYm9yZGVyUmFkaXVzOiA0LAogICAgICAgICAgICBtYXhCYXJUaGlja25lc3M6IDI4LAogICAgICAgICAgICBib3JkZXJTa2lwcGVkOiAnYm90dG9tJywKICAgICAgICAgIH0sCiAgICAgICAgXSwKICAgICAgfSwKICAgICAgb3B0aW9uczogewogICAgICAgIHJlc3BvbnNpdmU6IHRydWUsCiAgICAgICAgbWFpbnRhaW5Bc3BlY3RSYXRpbzogZmFsc2UsCiAgICAgICAgbGF5b3V0OiB7IHBhZGRpbmc6IHsgdG9wOiBzaG93TGFiZWxzID8gMjAgOiA4IH0gfSwKICAgICAgICBhbmltYXRpb246IGJhc2VBbmltYXRpb24oKSwKICAgICAgICBwbHVnaW5zOiB7CiAgICAgICAgICBsZWdlbmQ6IHsgZGlzcGxheTogZmFsc2UgfSwKICAgICAgICAgIHRvb2x0aXA6IGJhc2VUb29sdGlwKCksCiAgICAgICAgICBkYXRhbGFiZWxzOiBzaG93TGFiZWxzCiAgICAgICAgICAgID8geyBhbGlnbjogJ2VuZCcsIGFuY2hvcjogJ2VuZCcsIGNvbG9yOiBsYWJlbENvbG9yKCksIGZvbnQ6IHsgc2l6ZTogMTEsIHdlaWdodDogJzYwMCcgfSwgZm9ybWF0dGVyOiAodikgPT4gZm10KHYpIH0KICAgICAgICAgICAgOiB7IGRpc3BsYXk6IGZhbHNlIH0sCiAgICAgICAgfSwKICAgICAgICBzY2FsZXM6IHsKICAgICAgICAgIHg6IHsgZ3JpZDogeyBkaXNwbGF5OiBmYWxzZSB9LCB0aWNrczogYmFzZVRpY2tzKCkgfSwKICAgICAgICAgIHk6IHsgZ3JpZDogYmFzZUdyaWQoKSwgdGlja3M6IGJhc2VUaWNrcygpLCBib3JkZXI6IHsgZGlzcGxheTogZmFsc2UgfSwgYmVnaW5BdFplcm86IHRydWUgfSwKICAgICAgICB9LAogICAgICB9LAogICAgfSk7CiAgICByZWdpc3RyeS5zZXQoY2FudmFzSWQsIGNoYXJ0KTsKICAgIHJldHVybiBjaGFydDsKICB9CgogIC8qKiBHcm91cGVkIGJhciBjaGFydCBjb21wYXJpbmcgdHdvIHRpbWUgcGVyaW9kcyAoY29sb3IgPSBwZXJpb2QsIHgtYXhpcyA9IGNhdGVnb3J5KS4gKi8KICBmdW5jdGlvbiBjb21wYXJpc29uQmFyQ2hhcnQoY2FudmFzSWQsIHsgbGFiZWxzLCBjdXJyZW50RGF0YSwgcHJldmlvdXNEYXRhLCBjdXJyZW50TGFiZWwsIHByZXZpb3VzTGFiZWwsIGZvcm1hdFZhbHVlIH0pIHsKICAgIGRlc3Ryb3koY2FudmFzSWQpOwogICAgY29uc3QgY3R4ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoY2FudmFzSWQpOwogICAgaWYgKCFjdHgpIHJldHVybiBudWxsOwogICAgY29uc3QgZm10ID0gZm9ybWF0VmFsdWUgfHwgKCh2KSA9PiBGb3JtYXQuc21hcnQodikpOwogICAgY29uc3Qgc2hvd0xhYmVscyA9IGxhYmVscy5sZW5ndGggPD0gTUFYX0xBQkVMRURfSVRFTVMgLyAyOwoKICAgIGNvbnN0IGNoYXJ0ID0gbmV3IENoYXJ0KGN0eCwgewogICAgICB0eXBlOiAnYmFyJywKICAgICAgZGF0YTogewogICAgICAgIGxhYmVscywKICAgICAgICBkYXRhc2V0czogWwogICAgICAgICAgeyBsYWJlbDogcHJldmlvdXNMYWJlbCwgZGF0YTogcHJldmlvdXNEYXRhLCBiYWNrZ3JvdW5kQ29sb3I6IGNzc1ZhcignLS10ZXh0LW11dGVkJyksIGJvcmRlclJhZGl1czogNCwgbWF4QmFyVGhpY2tuZXNzOiAyMiB9LAogICAgICAgICAgeyBsYWJlbDogY3VycmVudExhYmVsLCBkYXRhOiBjdXJyZW50RGF0YSwgYmFja2dyb3VuZENvbG9yOiBjc3NWYXIoJy0tc2VyaWVzLTEnKSwgYm9yZGVyUmFkaXVzOiA0LCBtYXhCYXJUaGlja25lc3M6IDIyIH0sCiAgICAgICAgXSwKICAgICAgfSwKICAgICAgb3B0aW9uczogewogICAgICAgIHJlc3BvbnNpdmU6IHRydWUsCiAgICAgICAgbWFpbnRhaW5Bc3BlY3RSYXRpbzogZmFsc2UsCiAgICAgICAgbGF5b3V0OiB7IHBhZGRpbmc6IHsgdG9wOiBzaG93TGFiZWxzID8gMjAgOiA4IH0gfSwKICAgICAgICBhbmltYXRpb246IGJhc2VBbmltYXRpb24oKSwKICAgICAgICBwbHVnaW5zOiB7CiAgICAgICAgICBsZWdlbmQ6IHsgZGlzcGxheTogdHJ1ZSwgcG9zaXRpb246ICdib3R0b20nLCBsYWJlbHM6IHsgY29sb3I6IGNzc1ZhcignLS10ZXh0LXNlY29uZGFyeScpLCBib3hXaWR0aDogMTIsIHBhZGRpbmc6IDE2LCBmb250OiB7IHNpemU6IDExIH0gfSB9LAogICAgICAgICAgdG9vbHRpcDogYmFzZVRvb2x0aXAoKSwKICAgICAgICAgIGRhdGFsYWJlbHM6IHNob3dMYWJlbHMKICAgICAgICAgICAgPyB7IGFsaWduOiAnZW5kJywgYW5jaG9yOiAnZW5kJywgY29sb3I6IGxhYmVsQ29sb3IoKSwgZm9udDogeyBzaXplOiAxMCwgd2VpZ2h0OiAnNjAwJyB9LCBmb3JtYXR0ZXI6ICh2KSA9PiBmbXQodikgfQogICAgICAgICAgICA6IHsgZGlzcGxheTogZmFsc2UgfSwKICAgICAgICB9LAogICAgICAgIHNjYWxlczogewogICAgICAgICAgeDogeyBncmlkOiB7IGRpc3BsYXk6IGZhbHNlIH0sIHRpY2tzOiBiYXNlVGlja3MoKSB9LAogICAgICAgICAgeTogeyBncmlkOiBiYXNlR3JpZCgpLCB0aWNrczogYmFzZVRpY2tzKCksIGJvcmRlcjogeyBkaXNwbGF5OiBmYWxzZSB9LCBiZWdpbkF0WmVybzogdHJ1ZSB9LAogICAgICAgIH0sCiAgICAgIH0sCiAgICB9KTsKICAgIHJlZ2lzdHJ5LnNldChjYW52YXNJZCwgY2hhcnQpOwogICAgcmV0dXJuIGNoYXJ0OwogIH0KCiAgLyoqIFBpZSBjaGFydCAoYSBoYW5kZnVsIG9mIGNhdGVnb3JpZXMgb25seSDigJQgZS5nLiBDYW1wYWlnbiBQZXJmb3JtYW5jZSdzIEFkcy9PcmdhbmljIHNwbGl0KS4KICAgICAgU2xpY2UgbGFiZWxzIHNob3cgYm90aCBzaGFyZS1vZi13aG9sZSBhbmQgdGhlIGFjdHVhbCB2YWx1ZSwgcGVyIHRoZSAibm8gaG92ZXIgcmVxdWlyZWQiIGdvYWwuICovCiAgZnVuY3Rpb24gcGllQ2hhcnQoY2FudmFzSWQsIHsgbGFiZWxzLCBkYXRhLCBjb2xvcnMsIGZvcm1hdFZhbHVlIH0pIHsKICAgIGRlc3Ryb3koY2FudmFzSWQpOwogICAgY29uc3QgY3R4ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoY2FudmFzSWQpOwogICAgaWYgKCFjdHgpIHJldHVybiBudWxsOwogICAgY29uc3QgZm10ID0gZm9ybWF0VmFsdWUgfHwgKCh2KSA9PiBGb3JtYXQuc21hcnQodikpOwogICAgY29uc3QgdG90YWwgPSBkYXRhLnJlZHVjZSgoc3VtLCB2KSA9PiBzdW0gKyAodiB8fCAwKSwgMCk7CgogICAgY29uc3QgY2hhcnQgPSBuZXcgQ2hhcnQoY3R4LCB7CiAgICAgIHR5cGU6ICdwaWUnLAogICAgICBkYXRhOiB7CiAgICAgICAgbGFiZWxzLAogICAgICAgIGRhdGFzZXRzOiBbeyBkYXRhLCBiYWNrZ3JvdW5kQ29sb3I6IGNvbG9ycywgYm9yZGVyQ29sb3I6IGNzc1ZhcignLS1zdXJmYWNlLTEnKSwgYm9yZGVyV2lkdGg6IDIgfV0sCiAgICAgIH0sCiAgICAgIG9wdGlvbnM6IHsKICAgICAgICByZXNwb25zaXZlOiB0cnVlLAogICAgICAgIG1haW50YWluQXNwZWN0UmF0aW86IGZhbHNlLAogICAgICAgIGFuaW1hdGlvbjogYmFzZUFuaW1hdGlvbigpLAogICAgICAgIHBsdWdpbnM6IHsKICAgICAgICAgIGxlZ2VuZDogeyBkaXNwbGF5OiB0cnVlLCBwb3NpdGlvbjogJ2JvdHRvbScsIGxhYmVsczogeyBjb2xvcjogY3NzVmFyKCctLXRleHQtc2Vjb25kYXJ5JyksIGJveFdpZHRoOiAxMiwgcGFkZGluZzogMTYsIGZvbnQ6IHsgc2l6ZTogMTEgfSB9IH0sCiAgICAgICAgICB0b29sdGlwOiBiYXNlVG9vbHRpcCgpLAogICAgICAgICAgZGF0YWxhYmVsczogewogICAgICAgICAgICBjb2xvcjogJyNmZmYnLAogICAgICAgICAgICBmb250OiB7IHNpemU6IDEyLCB3ZWlnaHQ6ICc3MDAnIH0sCiAgICAgICAgICAgIGZvcm1hdHRlcjogKHYpID0+IHsKICAgICAgICAgICAgICBjb25zdCBwY3QgPSB0b3RhbCA/IE1hdGgucm91bmQoKHYgLyB0b3RhbCkgKiAxMDAwKSAvIDEwIDogMDsKICAgICAgICAgICAgICByZXR1cm4gYCR7cGN0fSVcbiR7Zm10KHYpfWA7CiAgICAgICAgICAgIH0sCiAgICAgICAgICB9LAogICAgICAgIH0sCiAgICAgIH0sCiAgICB9KTsKICAgIHJlZ2lzdHJ5LnNldChjYW52YXNJZCwgY2hhcnQpOwogICAgcmV0dXJuIGNoYXJ0OwogIH0KCiAgZnVuY3Rpb24gZGVzdHJveUFsbCgpIHsKICAgIFsuLi5yZWdpc3RyeS5rZXlzKCldLmZvckVhY2goZGVzdHJveSk7CiAgfQoKICByZXR1cm4geyB0cmVuZENoYXJ0LCBwbGF0Zm9ybUJhckNoYXJ0LCBjb21wYXJpc29uQmFyQ2hhcnQsIHBpZUNoYXJ0LCBzZXJpZXNDb2xvciwgZGVzdHJveSwgZGVzdHJveUFsbCB9Owp9KSgpOwoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIERhc2hib2FyZCB0YWI6IGEgbWV0cmljLWZvY3VzZWQgcHJlbWl1bSBCSSBkYXNoYm9hcmQuIEEgc2luZ2xlCiAgIE1ldHJpYyBzZWxlY3RvciAoZHluYW1pY2FsbHkgcG9wdWxhdGVkIGZyb20gd2hhdGV2ZXIgdGhlCiAgIHNlbGVjdGVkIHBsYXRmb3JtJ3MgZGF0YSBhY3R1YWxseSBoYXMg4oCUIG5ldmVyIGhhcmRjb2RlZCkgZHJpdmVzCiAgIHRoZSBLUEkgY2FyZHMsIHdlZWtseSB0cmVuZCwgcGxhdGZvcm0vY2FtcGFpZ24vY29udGVudC10eXBlCiAgIGJyZWFrZG93bnMsIGFuZCB0aGUgVG9wIFBlcmZvcm1pbmcgUG9zdHMgcmFua2luZyB0b2dldGhlcjsKICAgUGxhdGZvcm0vZGF0ZS9jYW1wYWlnbi9jb250ZW50LXR5cGUgZmlsdGVyaW5nIGNvbWVzIGZyb20gdGhlCiAgIHNoYXJlZCBmaWx0ZXIgYmFyLiBFdmVyeSBjaGFydCBzaG93cyBpdHMgdmFsdWVzIGRpcmVjdGx5ICh2aWEKICAgY2hhcnRqcy1wbHVnaW4tZGF0YWxhYmVscykgc28gbm90aGluZyByZXF1aXJlcyBhIGhvdmVyIHRvIHJlYWQuCiAgID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwpjb25zdCBEYXNoYm9hcmQgPSAoKCkgPT4gewogIGxldCByb290OwogIGxldCBtZXRyaWMgPSAndmlld3MnOwogIGxldCBtZXRyaWNPcHRpb25zID0gW107CgogIGZ1bmN0aW9uIG9wdGlvbkZvcihrZXkpIHsKICAgIHJldHVybiBtZXRyaWNPcHRpb25zLmZpbmQoKG0pID0+IG0ua2V5ID09PSBrZXkpOwogIH0KICBmdW5jdGlvbiBtZXRyaWNMYWJlbChrZXkpIHsKICAgIGNvbnN0IG9wdCA9IG9wdGlvbkZvcihrZXkpOwogICAgcmV0dXJuIG9wdCA/IG9wdC5sYWJlbCA6IGtleTsKICB9CiAgZnVuY3Rpb24gbWV0cmljVW5pdChrZXkpIHsKICAgIGNvbnN0IG9wdCA9IG9wdGlvbkZvcihrZXkpOwogICAgcmV0dXJuIG9wdCA/IG9wdC51bml0IDogJ251bWJlcic7CiAgfQogIGZ1bmN0aW9uIGZvcm1hdE1ldHJpY1ZhbHVlKGtleSwgdmFsdWUpIHsKICAgIGNvbnN0IHVuaXQgPSBtZXRyaWNVbml0KGtleSk7CiAgICBpZiAodmFsdWUgPT09IG51bGwgfHwgdmFsdWUgPT09IHVuZGVmaW5lZCkgcmV0dXJuICfigJQnOwogICAgaWYgKHVuaXQgPT09ICdkdXJhdGlvbicpIHJldHVybiBGb3JtYXQuZHVyYXRpb24odmFsdWUpOwogICAgaWYgKHVuaXQgPT09ICdwZXJjZW50JykgcmV0dXJuIEZvcm1hdC5wZXJjZW50KHZhbHVlKTsKICAgIHJldHVybiBGb3JtYXQuc21hcnQodmFsdWUpOwogIH0KCiAgZnVuY3Rpb24gc2hlbGwoKSB7CiAgICByb290LmlubmVySFRNTCA9ICcnOwoKICAgIGNvbnN0IGNvbnRyb2xzID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjb250cm9scy5jbGFzc05hbWUgPSAnZGFzaGJvYXJkLWNvbnRyb2xzJzsKICAgIGNvbnN0IGxhYmVsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnbGFiZWwnKTsKICAgIGxhYmVsLnRleHRDb250ZW50ID0gJ01ldHJpYyc7CiAgICBsYWJlbC5zZXRBdHRyaWJ1dGUoJ2ZvcicsICdkYXNoYm9hcmRNZXRyaWNTZWxlY3QnKTsKICAgIGNvbnN0IG1ldHJpY1NlbGVjdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NlbGVjdCcpOwogICAgbWV0cmljU2VsZWN0LmlkID0gJ2Rhc2hib2FyZE1ldHJpY1NlbGVjdCc7CiAgICBtZXRyaWNPcHRpb25zLmZvckVhY2goKG0pID0+IHsKICAgICAgY29uc3Qgb3B0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7CiAgICAgIG9wdC52YWx1ZSA9IG0ua2V5OwogICAgICBvcHQudGV4dENvbnRlbnQgPSBtLmxhYmVsOwogICAgICBpZiAobS5rZXkgPT09IG1ldHJpYykgb3B0LnNlbGVjdGVkID0gdHJ1ZTsKICAgICAgbWV0cmljU2VsZWN0LmFwcGVuZENoaWxkKG9wdCk7CiAgICB9KTsKICAgIG1ldHJpY1NlbGVjdC5hZGRFdmVudExpc3RlbmVyKCdjaGFuZ2UnLCAoKSA9PiB7CiAgICAgIG1ldHJpYyA9IG1ldHJpY1NlbGVjdC52YWx1ZTsKICAgICAgcmVmcmVzaEZvck1ldHJpYygpOwogICAgfSk7CiAgICBjb250cm9scy5hcHBlbmQobGFiZWwsIG1ldHJpY1NlbGVjdCk7CiAgICByb290LmFwcGVuZENoaWxkKGNvbnRyb2xzKTsKCiAgICBjb25zdCBrcGlUaXRsZSA9IHRleHRFbCgnZGl2JywgJ0tleSBwZXJmb3JtYW5jZSBpbmRpY2F0b3JzJywgJ3NlY3Rpb24tdGl0bGUnKTsKICAgIGNvbnN0IGtwaUdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGtwaUdyaWQuY2xhc3NOYW1lID0gJ3N0YXQtZ3JpZCc7CiAgICBrcGlHcmlkLmlkID0gJ2twaUdyaWQnOwogICAgcm9vdC5hcHBlbmQoa3BpVGl0bGUsIGtwaUdyaWQpOwoKICAgIGNvbnN0IGNoYXJ0c1RpdGxlID0gdGV4dEVsKCdkaXYnLCAnVHJlbmQgJiBwZXJmb3JtYW5jZSBicmVha2Rvd24nLCAnc2VjdGlvbi10aXRsZScpOwogICAgcm9vdC5hcHBlbmQoY2hhcnRzVGl0bGUpOwoKICAgIGNvbnN0IHRyZW5kQ2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdHJlbmRDYXJkLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGNvbnN0IHRyZW5kSGVhZGVyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0cmVuZEhlYWRlci5jbGFzc05hbWUgPSAnY2FyZC1oZWFkZXInOwogICAgdHJlbmRIZWFkZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMycsICdXZWVrbHkgcGVyZm9ybWFuY2UnKSk7CiAgICB0cmVuZEhlYWRlci5maXJzdENoaWxkLmlkID0gJ3RyZW5kQ2FyZFRpdGxlJzsKICAgIHRyZW5kQ2FyZC5hcHBlbmRDaGlsZCh0cmVuZEhlYWRlcik7CiAgICBjb25zdCB0cmVuZENoYXJ0V3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdHJlbmRDaGFydFdyYXAuY2xhc3NOYW1lID0gJ2NoYXJ0LXdyYXAgdGFsbCc7CiAgICB0cmVuZENoYXJ0V3JhcC5pZCA9ICd0cmVuZENoYXJ0V3JhcCc7CiAgICB0cmVuZENoYXJ0V3JhcC5pbm5lckhUTUwgPSAnPGNhbnZhcyBpZD0idHJlbmRDYW52YXMiPjwvY2FudmFzPic7CiAgICB0cmVuZENhcmQuYXBwZW5kQ2hpbGQodHJlbmRDaGFydFdyYXApOwogICAgcm9vdC5hcHBlbmRDaGlsZCh0cmVuZENhcmQpOwoKICAgIGNvbnN0IGdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGdyaWQuY2xhc3NOYW1lID0gJ2NhcmQtZ3JpZCBldmVuJzsKICAgIGdyaWQuc3R5bGUubWFyZ2luVG9wID0gJzE2cHgnOwoKICAgIGNvbnN0IGJyZWFrZG93bkNhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGJyZWFrZG93bkNhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgY29uc3QgYnJlYWtkb3duSGVhZGVyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBicmVha2Rvd25IZWFkZXIuY2xhc3NOYW1lID0gJ2NhcmQtaGVhZGVyJzsKICAgIGJyZWFrZG93bkhlYWRlci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2gzJywgJycpKTsKICAgIGJyZWFrZG93bkhlYWRlci5maXJzdENoaWxkLmlkID0gJ2JyZWFrZG93bkNhcmRUaXRsZSc7CiAgICBicmVha2Rvd25DYXJkLmFwcGVuZENoaWxkKGJyZWFrZG93bkhlYWRlcik7CiAgICBjb25zdCBicmVha2Rvd25XcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBicmVha2Rvd25XcmFwLmNsYXNzTmFtZSA9ICdjaGFydC13cmFwJzsKICAgIGJyZWFrZG93bldyYXAuaWQgPSAnYnJlYWtkb3duQ2hhcnRXcmFwJzsKICAgIGJyZWFrZG93bldyYXAuaW5uZXJIVE1MID0gJzxjYW52YXMgaWQ9ImJyZWFrZG93bkNhbnZhcyI+PC9jYW52YXM+JzsKICAgIGJyZWFrZG93bkNhcmQuYXBwZW5kQ2hpbGQoYnJlYWtkb3duV3JhcCk7CgogICAgY29uc3QgY29udGVudFR5cGVDYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjb250ZW50VHlwZUNhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgY29uc3QgY29udGVudFR5cGVIZWFkZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNvbnRlbnRUeXBlSGVhZGVyLmNsYXNzTmFtZSA9ICdjYXJkLWhlYWRlcic7CiAgICBjb250ZW50VHlwZUhlYWRlci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2gzJywgJycpKTsKICAgIGNvbnRlbnRUeXBlSGVhZGVyLmZpcnN0Q2hpbGQuaWQgPSAnY29udGVudFR5cGVDYXJkVGl0bGUnOwogICAgY29udGVudFR5cGVDYXJkLmFwcGVuZENoaWxkKGNvbnRlbnRUeXBlSGVhZGVyKTsKICAgIGNvbnN0IGNvbnRlbnRUeXBlV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY29udGVudFR5cGVXcmFwLmNsYXNzTmFtZSA9ICdjaGFydC13cmFwJzsKICAgIGNvbnRlbnRUeXBlV3JhcC5pZCA9ICdjb250ZW50VHlwZUNoYXJ0V3JhcCc7CiAgICBjb250ZW50VHlwZVdyYXAuaW5uZXJIVE1MID0gJzxjYW52YXMgaWQ9ImNvbnRlbnRUeXBlQ2FudmFzIj48L2NhbnZhcz4nOwogICAgY29udGVudFR5cGVDYXJkLmFwcGVuZENoaWxkKGNvbnRlbnRUeXBlV3JhcCk7CgogICAgZ3JpZC5hcHBlbmQoYnJlYWtkb3duQ2FyZCwgY29udGVudFR5cGVDYXJkKTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQoZ3JpZCk7CgogICAgY29uc3QgdG9wVGl0bGUgPSB0ZXh0RWwoJ2RpdicsICdUb3AtcGVyZm9ybWluZyBwb3N0cycsICdzZWN0aW9uLXRpdGxlJyk7CiAgICBjb25zdCB0b3BDYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0b3BDYXJkLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGNvbnN0IHRvcEhlYWRlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdG9wSGVhZGVyLmNsYXNzTmFtZSA9ICdjYXJkLWhlYWRlcic7CiAgICB0b3BIZWFkZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMycsICdSYW5rZWQgYnkgc2VsZWN0ZWQgbWV0cmljJykpOwogICAgdG9wQ2FyZC5hcHBlbmRDaGlsZCh0b3BIZWFkZXIpOwogICAgY29uc3QgdGFibGVXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0YWJsZVdyYXAuY2xhc3NOYW1lID0gJ3RhYmxlLXNjcm9sbCc7CiAgICB0YWJsZVdyYXAuaWQgPSAndG9wUG9zdHNUYWJsZSc7CiAgICB0b3BDYXJkLmFwcGVuZENoaWxkKHRhYmxlV3JhcCk7CiAgICByb290LmFwcGVuZCh0b3BUaXRsZSwgdG9wQ2FyZCk7CiAgfQoKICBmdW5jdGlvbiBwb3N0VGlsZShsYWJlbCwgcG9zdCkgewogICAgY29uc3QgdGlsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGlsZS5jbGFzc05hbWUgPSAnc3RhdC10aWxlIHBvc3QtdGlsZSc7CiAgICB0aWxlLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgbGFiZWwsICdzdGF0LWxhYmVsJykpOwogICAgaWYgKCFwb3N0KSB7CiAgICAgIHRpbGUuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCAnTm8gZGF0YSB5ZXQnLCAnc3RhdC12YWx1ZScpKTsKICAgICAgcmV0dXJuIHRpbGU7CiAgICB9CiAgICBjb25zdCBwbGF0Zm9ybU9wdGlvbnMgPSAod2luZG93Ll9fZmlsdGVyT3B0aW9uc0NhY2hlIHx8IHsgcGxhdGZvcm1zOiBbXSB9KS5wbGF0Zm9ybXM7CiAgICBjb25zdCBtZXRhID0gcGxhdGZvcm1PcHRpb25zLmZpbmQoKHApID0+IHAuaWQgPT09IHBvc3QucGxhdGZvcm0pIHx8IHsgbGFiZWw6IHBvc3QucGxhdGZvcm0gfTsKICAgIGNvbnN0IGNhcHRpb24gPSBwb3N0LmNhcHRpb24gfHwgJyhubyBjYXB0aW9uKSc7CiAgICBjb25zdCB2YWx1ZUVsID0gdGV4dEVsKCdkaXYnLCBjYXB0aW9uLmxlbmd0aCA+IDcwID8gYCR7Y2FwdGlvbi5zbGljZSgwLCA3MCl94oCmYCA6IGNhcHRpb24sICdzdGF0LXZhbHVlJyk7CiAgICB2YWx1ZUVsLnRpdGxlID0gY2FwdGlvbjsKICAgIHRpbGUuYXBwZW5kQ2hpbGQodmFsdWVFbCk7CiAgICB0aWxlLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgYCR7bWV0YS5sYWJlbH0gwrcgJHtGb3JtYXQuZGF0ZShwb3N0LnB1Ymxpc2hfZGF0ZSl9YCwgJ3Bvc3QtbWV0YScpKTsKICAgIHRpbGUuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBmb3JtYXRNZXRyaWNWYWx1ZShtZXRyaWMsIHBvc3QudmFsdWUpLCAncG9zdC1tZXRyaWMtdmFsdWUnKSk7CiAgICByZXR1cm4gdGlsZTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlcktwaXMoc3VtbWFyeSkgewogICAgY29uc3QgZ3JpZCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdrcGlHcmlkJyk7CiAgICBpZiAoIWdyaWQpIHJldHVybjsKICAgIGdyaWQuaW5uZXJIVE1MID0gJyc7CgogICAgY29uc3Qgc3RhdFRpbGUgPSAobGFiZWwsIHZhbHVlKSA9PiB7CiAgICAgIGNvbnN0IHRpbGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgdGlsZS5jbGFzc05hbWUgPSAnc3RhdC10aWxlJzsKICAgICAgdGlsZS5hcHBlbmQodGV4dEVsKCdkaXYnLCBsYWJlbCwgJ3N0YXQtbGFiZWwnKSwgdGV4dEVsKCdkaXYnLCB2YWx1ZSwgJ3N0YXQtdmFsdWUnKSk7CiAgICAgIHJldHVybiB0aWxlOwogICAgfTsKCiAgICBncmlkLmFwcGVuZENoaWxkKHN0YXRUaWxlKCdIaWdoZXN0IFZhbHVlJywgZm9ybWF0TWV0cmljVmFsdWUobWV0cmljLCBzdW1tYXJ5LmhpZ2hlc3QpKSk7CiAgICBncmlkLmFwcGVuZENoaWxkKHN0YXRUaWxlKCdBdmVyYWdlIFZhbHVlJywgZm9ybWF0TWV0cmljVmFsdWUobWV0cmljLCBzdW1tYXJ5LmF2ZXJhZ2UpKSk7CiAgICBpZiAoc3VtbWFyeS51bml0ICE9PSAncGVyY2VudCcpIHsKICAgICAgZ3JpZC5hcHBlbmRDaGlsZChzdGF0VGlsZSgnVG90YWwgVmFsdWUnLCBmb3JtYXRNZXRyaWNWYWx1ZShtZXRyaWMsIHN1bW1hcnkudG90YWwpKSk7CiAgICB9CiAgICBncmlkLmFwcGVuZENoaWxkKHN0YXRUaWxlKCdOdW1iZXIgb2YgUG9zdHMnLCBGb3JtYXQubnVtYmVyKHN1bW1hcnkucG9zdENvdW50KSkpOwogICAgZ3JpZC5hcHBlbmRDaGlsZChwb3N0VGlsZSgnQmVzdCBQZXJmb3JtaW5nIFBvc3QnLCBzdW1tYXJ5LmJlc3RQb3N0KSk7CiAgICBncmlkLmFwcGVuZENoaWxkKHBvc3RUaWxlKCdMb3dlc3QgUGVyZm9ybWluZyBQb3N0Jywgc3VtbWFyeS53b3JzdFBvc3QpKTsKICB9CgogIC8qKiBTd2FwcyBhIGNoYXJ0IGNhcmQncyBjYW52YXMgZm9yIGFuIGVtcHR5LXN0YXRlIG1lc3NhZ2UsIG9yIHJlc3RvcmVzIHRoZSBjYW52YXMg4oCUIHNpbmNlCiAgICAgIHJlLXJlbmRlcmluZyBhIENoYXJ0LmpzIGluc3RhbmNlIG5lZWRzIGEgbGl2ZSA8Y2FudmFzPiwgbm90IHdoYXRldmVyIHRoZSBsYXN0IHJlbmRlciBsZWZ0IHRoZXJlLiAqLwogIGZ1bmN0aW9uIGNoYXJ0T3JFbXB0eSh3cmFwSWQsIGNhbnZhc0lkLCBoYXNEYXRhLCBlbXB0eU1lc3NhZ2UsIHJlbmRlckZuKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQod3JhcElkKTsKICAgIGlmICghd3JhcCkgcmV0dXJuOwogICAgQ2hhcnRzLmRlc3Ryb3koY2FudmFzSWQpOwogICAgaWYgKCFoYXNEYXRhKSB7CiAgICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQoZW1wdHlTdGF0ZSh7IGljb246ICdiYXItY2hhcnQtMycsIG1lc3NhZ2U6IGVtcHR5TWVzc2FnZSB9KSk7CiAgICAgIHJldHVybjsKICAgIH0KICAgIHdyYXAuaW5uZXJIVE1MID0gYDxjYW52YXMgaWQ9IiR7Y2FudmFzSWR9Ij48L2NhbnZhcz5gOwogICAgcmVuZGVyRm4oKTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIHJlbmRlclRyZW5kKGZpbHRlcnMpIHsKICAgIGNvbnN0IHBsYXRmb3JtT3B0aW9ucyA9ICh3aW5kb3cuX19maWx0ZXJPcHRpb25zQ2FjaGUgfHwgeyBwbGF0Zm9ybXM6IFtdIH0pLnBsYXRmb3JtczsKICAgIGNvbnN0IG1MYWJlbCA9IG1ldHJpY0xhYmVsKG1ldHJpYyk7CiAgICBjb25zdCBwbGF0Zm9ybXNUb0ZldGNoID0gZmlsdGVycy5wbGF0Zm9ybSA9PT0gJ2FsbCcgPyBwbGF0Zm9ybU9wdGlvbnMubWFwKChwKSA9PiBwLmlkKSA6IFtmaWx0ZXJzLnBsYXRmb3JtXTsKICAgIGNvbnN0IHRyZW5kUmVzcG9uc2VzID0gYXdhaXQgUHJvbWlzZS5hbGwoCiAgICAgIHBsYXRmb3Jtc1RvRmV0Y2gubWFwKChwKSA9PgogICAgICAgIEFwaS50cmVuZCh7IGRhdGVGcm9tOiBmaWx0ZXJzLmRhdGVGcm9tLCBkYXRlVG86IGZpbHRlcnMuZGF0ZVRvLCBwbGF0Zm9ybTogcCwgY2FtcGFpZ25UeXBlOiBmaWx0ZXJzLmNhbXBhaWduVHlwZSwgY29udGVudFR5cGU6IGZpbHRlcnMuY29udGVudFR5cGUgfSkKICAgICAgKQogICAgKTsKICAgIGNvbnN0IHdlZWtTZXQgPSBuZXcgU2V0KCk7CiAgICB0cmVuZFJlc3BvbnNlcy5mb3JFYWNoKChyb3dzKSA9PiByb3dzLmZvckVhY2goKHIpID0+IHdlZWtTZXQuYWRkKHIucGVyaW9kKSkpOwogICAgY29uc3Qgd2Vla3MgPSBbLi4ud2Vla1NldF0uc29ydCgpOwogICAgY29uc3Qgc2VyaWVzID0gcGxhdGZvcm1zVG9GZXRjaC5tYXAoKHAsIGkpID0+IHsKICAgICAgY29uc3QgbWV0YSA9IHBsYXRmb3JtT3B0aW9ucy5maW5kKChwbCkgPT4gcGwuaWQgPT09IHApIHx8IHsgbGFiZWw6IHAgfTsKICAgICAgY29uc3QgYnlXZWVrID0gT2JqZWN0LmZyb21FbnRyaWVzKHRyZW5kUmVzcG9uc2VzW2ldLm1hcCgocikgPT4gW3IucGVyaW9kLCByW21ldHJpY11dKSk7CiAgICAgIHJldHVybiB7IGxhYmVsOiBtZXRhLmxhYmVsLCBjb2xvcjogbWV0YS5jb2xvciwgZGF0YTogd2Vla3MubWFwKCh3KSA9PiAoYnlXZWVrW3ddID09PSB1bmRlZmluZWQgPyBudWxsIDogYnlXZWVrW3ddKSkgfTsKICAgIH0pOwoKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd0cmVuZENhcmRUaXRsZScpLnRleHRDb250ZW50ID0KICAgICAgZmlsdGVycy5wbGF0Zm9ybSA9PT0gJ2FsbCcgPyBgV2Vla2x5ICR7bUxhYmVsfSBieSBQbGF0Zm9ybWAgOiBgJHttTGFiZWx9IFRyZW5kYDsKCiAgICBjaGFydE9yRW1wdHkoJ3RyZW5kQ2hhcnRXcmFwJywgJ3RyZW5kQ2FudmFzJywgd2Vla3MubGVuZ3RoID4gMCwgJ05vIGRhdGEgaW4gdGhpcyByYW5nZSB5ZXQuJywgKCkgPT4gewogICAgICBDaGFydHMudHJlbmRDaGFydCgndHJlbmRDYW52YXMnLCB7IGxhYmVsczogd2Vla3MubWFwKEZvcm1hdC5kYXRlKSwgc2VyaWVzLCBmb3JtYXRWYWx1ZTogKHYpID0+IGZvcm1hdE1ldHJpY1ZhbHVlKG1ldHJpYywgdikgfSk7CiAgICB9KTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIHJlbmRlckJyZWFrZG93bihmaWx0ZXJzKSB7CiAgICBjb25zdCBtTGFiZWwgPSBtZXRyaWNMYWJlbChtZXRyaWMpOwogICAgY29uc3QgdGl0bGVFbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdicmVha2Rvd25DYXJkVGl0bGUnKTsKCiAgICBpZiAoZmlsdGVycy5wbGF0Zm9ybSA9PT0gJ2FsbCcpIHsKICAgICAgdGl0bGVFbC50ZXh0Q29udGVudCA9IGBQbGF0Zm9ybSBDb21wYXJpc29uIOKAlCAke21MYWJlbH1gOwogICAgICBjb25zdCBicmVha2Rvd24gPSBhd2FpdCBBcGkucGxhdGZvcm1CcmVha2Rvd24oZmlsdGVycyk7CiAgICAgIGNvbnN0IHNvcnRlZCA9IGJyZWFrZG93bi5maWx0ZXIoKHApID0+IHBbbWV0cmljXSAhPT0gbnVsbCAmJiBwW21ldHJpY10gIT09IHVuZGVmaW5lZCkuc29ydCgoYSwgYikgPT4gYlttZXRyaWNdIC0gYVttZXRyaWNdKTsKICAgICAgY2hhcnRPckVtcHR5KCdicmVha2Rvd25DaGFydFdyYXAnLCAnYnJlYWtkb3duQ2FudmFzJywgc29ydGVkLmxlbmd0aCA+IDAsICdObyBkYXRhIGluIHRoaXMgcmFuZ2UgeWV0LicsICgpID0+IHsKICAgICAgICBDaGFydHMucGxhdGZvcm1CYXJDaGFydCgnYnJlYWtkb3duQ2FudmFzJywgewogICAgICAgICAgbGFiZWxzOiBzb3J0ZWQubWFwKChwKSA9PiBwLmxhYmVsKSwKICAgICAgICAgIGRhdGE6IHNvcnRlZC5tYXAoKHApID0+IHBbbWV0cmljXSksCiAgICAgICAgICBjb2xvcnM6IHNvcnRlZC5tYXAoKHApID0+IHAuY29sb3IpLAogICAgICAgICAgZm9ybWF0VmFsdWU6ICh2KSA9PiBmb3JtYXRNZXRyaWNWYWx1ZShtZXRyaWMsIHYpLAogICAgICAgIH0pOwogICAgICB9KTsKICAgIH0gZWxzZSB7CiAgICAgIHRpdGxlRWwudGV4dENvbnRlbnQgPSBgQ2FtcGFpZ24gUGVyZm9ybWFuY2Ug4oCUICR7bUxhYmVsfWA7CiAgICAgIGNvbnN0IGNhbXBhaWducyA9IGF3YWl0IEFwaS5jYW1wYWlnbkJyZWFrZG93bihmaWx0ZXJzKTsKICAgICAgY29uc3Qgd2l0aFZhbHVlID0gY2FtcGFpZ25zLmZpbHRlcigoYykgPT4gY1ttZXRyaWNdICE9PSBudWxsICYmIGNbbWV0cmljXSAhPT0gdW5kZWZpbmVkICYmIGNbbWV0cmljXSA+IDApOwogICAgICBjaGFydE9yRW1wdHkoJ2JyZWFrZG93bkNoYXJ0V3JhcCcsICdicmVha2Rvd25DYW52YXMnLCB3aXRoVmFsdWUubGVuZ3RoID4gMCwgJ05vIGNhbXBhaWduIGRhdGEgaW4gdGhpcyByYW5nZSB5ZXQuJywgKCkgPT4gewogICAgICAgIENoYXJ0cy5waWVDaGFydCgnYnJlYWtkb3duQ2FudmFzJywgewogICAgICAgICAgbGFiZWxzOiB3aXRoVmFsdWUubWFwKChjKSA9PiBjLmNhbXBhaWduX3R5cGUpLAogICAgICAgICAgZGF0YTogd2l0aFZhbHVlLm1hcCgoYykgPT4gY1ttZXRyaWNdKSwKICAgICAgICAgIGNvbG9yczogd2l0aFZhbHVlLm1hcCgoXywgaSkgPT4gQ2hhcnRzLnNlcmllc0NvbG9yKGkpKSwKICAgICAgICAgIGZvcm1hdFZhbHVlOiAodikgPT4gZm9ybWF0TWV0cmljVmFsdWUobWV0cmljLCB2KSwKICAgICAgICB9KTsKICAgICAgfSk7CiAgICB9CiAgfQoKICBhc3luYyBmdW5jdGlvbiByZW5kZXJDb250ZW50VHlwZUJyZWFrZG93bihmaWx0ZXJzKSB7CiAgICBjb25zdCBtTGFiZWwgPSBtZXRyaWNMYWJlbChtZXRyaWMpOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2NvbnRlbnRUeXBlQ2FyZFRpdGxlJykudGV4dENvbnRlbnQgPSBgQ29udGVudCBUeXBlIFBlcmZvcm1hbmNlIOKAlCAke21MYWJlbH1gOwogICAgY29uc3Qgcm93cyA9IGF3YWl0IEFwaS5jb250ZW50VHlwZUJyZWFrZG93bihmaWx0ZXJzKTsKICAgIGNvbnN0IHNvcnRlZCA9IHJvd3MuZmlsdGVyKChjKSA9PiBjW21ldHJpY10gIT09IG51bGwgJiYgY1ttZXRyaWNdICE9PSB1bmRlZmluZWQpLnNvcnQoKGEsIGIpID0+IGJbbWV0cmljXSAtIGFbbWV0cmljXSk7CiAgICBjaGFydE9yRW1wdHkoJ2NvbnRlbnRUeXBlQ2hhcnRXcmFwJywgJ2NvbnRlbnRUeXBlQ2FudmFzJywgc29ydGVkLmxlbmd0aCA+IDAsICdObyBkYXRhIGluIHRoaXMgcmFuZ2UgeWV0LicsICgpID0+IHsKICAgICAgQ2hhcnRzLnBsYXRmb3JtQmFyQ2hhcnQoJ2NvbnRlbnRUeXBlQ2FudmFzJywgewogICAgICAgIGxhYmVsczogc29ydGVkLm1hcCgoYykgPT4gYy5jb250ZW50X3R5cGUpLAogICAgICAgIGRhdGE6IHNvcnRlZC5tYXAoKGMpID0+IGNbbWV0cmljXSksCiAgICAgICAgY29sb3JzOiBzb3J0ZWQubWFwKChfLCBpKSA9PiBDaGFydHMuc2VyaWVzQ29sb3IoaSkpLAogICAgICAgIGZvcm1hdFZhbHVlOiAodikgPT4gZm9ybWF0TWV0cmljVmFsdWUobWV0cmljLCB2KSwKICAgICAgfSk7CiAgICB9KTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIHJlbmRlclRvcFBvc3RzKGZpbHRlcnMpIHsKICAgIGNvbnN0IHBvc3RzID0gYXdhaXQgQXBpLnRvcFBvc3RzKHsgLi4uZmlsdGVycywgc29ydEJ5OiBtZXRyaWMsIGxpbWl0OiAxMCB9KTsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndG9wUG9zdHNUYWJsZScpOwogICAgaWYgKCF3cmFwKSByZXR1cm47CiAgICBpZiAoIXBvc3RzLmxlbmd0aCkgewogICAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgICB3cmFwLmFwcGVuZENoaWxkKGVtcHR5U3RhdGUoewogICAgICAgIGljb246ICd0cm9waHknLAogICAgICAgIHRpdGxlOiAnTm8gcG9zdHMgaW4gdGhpcyByYW5nZSB5ZXQnLAogICAgICAgIG1lc3NhZ2U6ICdVcGxvYWQgYSB3ZWVrbHkgZXhwb3J0LCBvciB3aWRlbiB0aGUgZGF0ZSByYW5nZSwgdG8gc2VlIHRvcCBwZXJmb3JtZXJzIGhlcmUuJywKICAgICAgICBhY3Rpb25MYWJlbDogJ1VwbG9hZCBkYXRhJywKICAgICAgICBvbkFjdGlvbjogKCkgPT4gZG9jdW1lbnQucXVlcnlTZWxlY3RvcignLnRhYi1idG5bZGF0YS10YWI9InVwbG9hZCJdJyk/LmNsaWNrKCksCiAgICAgIH0pKTsKICAgICAgcmV0dXJuOwogICAgfQogICAgY29uc3QgcGxhdGZvcm1PcHRpb25zID0gKHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZSB8fCB7IHBsYXRmb3JtczogW10gfSkucGxhdGZvcm1zOwogICAgY29uc3Qgc2hvd0VuZ2FnZW1lbnRSYXRlID0gbWV0cmljICE9PSAnZW5nYWdlbWVudF9yYXRlJyAmJiBwb3N0cy5zb21lKChwKSA9PiBwLmVuZ2FnZW1lbnRfcmF0ZSAhPT0gbnVsbCAmJiBwLmVuZ2FnZW1lbnRfcmF0ZSAhPT0gdW5kZWZpbmVkKTsKCiAgICBjb25zdCB0YWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICB0YWJsZS5jbGFzc05hbWUgPSAnZGF0YS10YWJsZSc7CiAgICBjb25zdCB0aGVhZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RoZWFkJyk7CiAgICBjb25zdCBoZWFkVHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgaGVhZFRyLmFwcGVuZCgKICAgICAgdGV4dEVsKCd0aCcsICdSYW5rJyksCiAgICAgIHRleHRFbCgndGgnLCAnRGF0ZScpLAogICAgICB0ZXh0RWwoJ3RoJywgJ1BsYXRmb3JtJyksCiAgICAgIHRleHRFbCgndGgnLCAnQ2FtcGFpZ24nKSwKICAgICAgdGV4dEVsKCd0aCcsICdDb250ZW50IFR5cGUnKSwKICAgICAgdGV4dEVsKCd0aCcsICdDYXB0aW9uJyksCiAgICAgIHRleHRFbCgndGgnLCBtZXRyaWNMYWJlbChtZXRyaWMpLCAnbnVtJykKICAgICk7CiAgICBpZiAoc2hvd0VuZ2FnZW1lbnRSYXRlKSBoZWFkVHIuYXBwZW5kQ2hpbGQodGV4dEVsKCd0aCcsICdFbmdhZ2VtZW50IFJhdGUnLCAnbnVtJykpOwogICAgaGVhZFRyLmFwcGVuZENoaWxkKHRleHRFbCgndGgnLCAnJykpOwogICAgdGhlYWQuYXBwZW5kQ2hpbGQoaGVhZFRyKTsKCiAgICBjb25zdCB0Ym9keSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3Rib2R5Jyk7CiAgICBwb3N0cy5mb3JFYWNoKChwLCBpKSA9PiB7CiAgICAgIGNvbnN0IHRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgY29uc3QgbWV0YSA9IHBsYXRmb3JtT3B0aW9ucy5maW5kKChwbCkgPT4gcGwuaWQgPT09IHAucGxhdGZvcm0pIHx8IHsgbGFiZWw6IHAucGxhdGZvcm0sIGNvbG9yOiAnIzk5OScgfTsKICAgICAgY29uc3QgcGxhdGZvcm1UZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIGNvbnN0IHBpbGwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICAgIHBpbGwuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLXBpbGwnOwogICAgICBjb25zdCBkb3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICAgIGRvdC5jbGFzc05hbWUgPSAncGxhdGZvcm0tZG90JzsKICAgICAgZG90LnN0eWxlLmJhY2tncm91bmQgPSBtZXRhLmNvbG9yOwogICAgICBwaWxsLmFwcGVuZChkb3QsIGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKG1ldGEubGFiZWwpKTsKICAgICAgcGxhdGZvcm1UZC5hcHBlbmRDaGlsZChwaWxsKTsKCiAgICAgIGNvbnN0IGNhcHRpb24gPSBwLmNhcHRpb24gfHwgJyhubyBjYXB0aW9uKSc7CiAgICAgIGNvbnN0IGNhcHRpb25UZCA9IHRleHRFbCgndGQnLCBjYXB0aW9uLmxlbmd0aCA+IDYwID8gYCR7Y2FwdGlvbi5zbGljZSgwLCA2MCl94oCmYCA6IGNhcHRpb24pOwogICAgICBjYXB0aW9uVGQudGl0bGUgPSBjYXB0aW9uOwoKICAgICAgdHIuYXBwZW5kKAogICAgICAgIHRleHRFbCgndGQnLCBgIyR7aSArIDF9YCksCiAgICAgICAgdGV4dEVsKCd0ZCcsIEZvcm1hdC5kYXRlKHAucHVibGlzaF9kYXRlKSksCiAgICAgICAgcGxhdGZvcm1UZCwKICAgICAgICB0ZXh0RWwoJ3RkJywgcC5jYW1wYWlnbl90eXBlIHx8ICfigJQnKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgcC5jb250ZW50X3R5cGUgfHwgJ+KAlCcpLAogICAgICAgIGNhcHRpb25UZCwKICAgICAgICB0ZXh0RWwoJ3RkJywgZm9ybWF0TWV0cmljVmFsdWUobWV0cmljLCBwLm1ldHJpY192YWx1ZSksICdudW0nKQogICAgICApOwogICAgICBpZiAoc2hvd0VuZ2FnZW1lbnRSYXRlKSB7CiAgICAgICAgdHIuYXBwZW5kQ2hpbGQodGV4dEVsKCd0ZCcsIHAuZW5nYWdlbWVudF9yYXRlID09PSBudWxsIHx8IHAuZW5nYWdlbWVudF9yYXRlID09PSB1bmRlZmluZWQgPyAn4oCUJyA6IEZvcm1hdC5wZXJjZW50KHAuZW5nYWdlbWVudF9yYXRlKSwgJ251bScpKTsKICAgICAgfQoKICAgICAgY29uc3QgYWN0aW9uVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBjb25zdCB2aWV3QnRuID0gaWNvbkJ0bignYnRuJywgJ2V5ZScsICdWaWV3IERldGFpbHMnKTsKICAgICAgdmlld0J0bi5kaXNhYmxlZCA9ICFwLnJhd19yb3dfaWQ7CiAgICAgIHZpZXdCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBSZWNvcmRzLm9wZW5WaWV3KHAucmF3X3Jvd19pZCkpOwogICAgICBhY3Rpb25UZC5hcHBlbmRDaGlsZCh2aWV3QnRuKTsKICAgICAgdHIuYXBwZW5kQ2hpbGQoYWN0aW9uVGQpOwoKICAgICAgdGJvZHkuYXBwZW5kQ2hpbGQodHIpOwogICAgfSk7CiAgICB0YWJsZS5hcHBlbmQodGhlYWQsIHRib2R5KTsKICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CiAgICB3cmFwLmFwcGVuZENoaWxkKHRhYmxlKTsKICB9CgogIC8qKiBNZXRyaWMgKG9yIGFueSBmaWx0ZXIpIGNoYW5nZWQgYnV0IHRoZSBwbGF0Zm9ybSDigJQgYW5kIHRoZXJlZm9yZSB0aGUgYXZhaWxhYmxlIG1ldHJpYyBsaXN0IOKAlCBkaWRuJ3Q6IG5vIG5lZWQgdG8gcmUtZmV0Y2ggbWV0cmljLW9wdGlvbnMgb3IgcmVidWlsZCB0aGUgc2hlbGwsIGp1c3QgcmVmcmVzaCB0aGUgZGF0YS4gKi8KICBhc3luYyBmdW5jdGlvbiByZWZyZXNoRm9yTWV0cmljKCkgewogICAgY29uc3QgZmlsdGVycyA9IFN0YXRlLmdldEZpbHRlcnMoKTsKICAgIGNvbnN0IHN1bW1hcnkgPSBhd2FpdCBBcGkubWV0cmljU3VtbWFyeSh7IC4uLmZpbHRlcnMsIG1ldHJpYyB9KTsKICAgIHJlbmRlcktwaXMoc3VtbWFyeSk7CiAgICBhd2FpdCBQcm9taXNlLmFsbChbcmVuZGVyVHJlbmQoZmlsdGVycyksIHJlbmRlckJyZWFrZG93bihmaWx0ZXJzKSwgcmVuZGVyQ29udGVudFR5cGVCcmVha2Rvd24oZmlsdGVycyksIHJlbmRlclRvcFBvc3RzKGZpbHRlcnMpXSk7CiAgfQoKICBmdW5jdGlvbiBzaG93U2tlbGV0b25zKCkgewogICAgY29uc3Qga3BpR3JpZCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdrcGlHcmlkJyk7CiAgICBpZiAoa3BpR3JpZCkgeyBrcGlHcmlkLmlubmVySFRNTCA9ICcnOyBrcGlHcmlkLmFwcGVuZENoaWxkKHNrZWxldG9uU3RhdEdyaWQoNikpOyB9CiAgICBbJ3RyZW5kQ2hhcnRXcmFwJywgJ2JyZWFrZG93bkNoYXJ0V3JhcCcsICdjb250ZW50VHlwZUNoYXJ0V3JhcCddLmZvckVhY2goKGlkKSA9PiB7CiAgICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZChpZCk7CiAgICAgIGlmICh3cmFwKSB7IHdyYXAuaW5uZXJIVE1MID0gJyc7IHdyYXAuYXBwZW5kQ2hpbGQoc2tlbGV0b25DaGFydCgpKTsgfQogICAgfSk7CiAgICBjb25zdCB0b3BQb3N0cyA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd0b3BQb3N0c1RhYmxlJyk7CiAgICBpZiAodG9wUG9zdHMpIHsgdG9wUG9zdHMuaW5uZXJIVE1MID0gJyc7IHRvcFBvc3RzLmFwcGVuZENoaWxkKHNrZWxldG9uUm93cyg2KSk7IH0KICB9CgogIGFzeW5jIGZ1bmN0aW9uIHJlbmRlcigpIHsKICAgIHJvb3QgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndmlldy1kYXNoYm9hcmQnKTsKICAgIGNvbnN0IGZpbHRlcnMgPSBTdGF0ZS5nZXRGaWx0ZXJzKCk7CiAgICBjb25zdCB7IG9wdGlvbnMgfSA9IGF3YWl0IEFwaS5tZXRyaWNPcHRpb25zKGZpbHRlcnMucGxhdGZvcm0pOwogICAgbWV0cmljT3B0aW9ucyA9IG9wdGlvbnM7CiAgICBpZiAoIW1ldHJpY09wdGlvbnMuc29tZSgobSkgPT4gbS5rZXkgPT09IG1ldHJpYykpIHsKICAgICAgbWV0cmljID0gbWV0cmljT3B0aW9ucy5sZW5ndGggPyBtZXRyaWNPcHRpb25zWzBdLmtleSA6ICd2aWV3cyc7CiAgICB9CiAgICBzaGVsbCgpOwogICAgc2hvd1NrZWxldG9ucygpOwogICAgYXdhaXQgcmVmcmVzaEZvck1ldHJpYygpOwogIH0KCiAgcmV0dXJuIHsgcmVuZGVyIH07Cn0pKCk7CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgRGF0YSBSZWNvcmRzIHRhYjogYSBDUk0tc3R5bGUsIHBsYXRmb3JtLWdyb3VwZWQgYnJvd3NlciBiYWNrZWQKICAgYnkgcG9zdHMvcG9zdF9tZXRyaWNzICh0aGUgc2FtZSBub3JtYWxpemVkIGRhdGEgdGhlIGRhc2hib2FyZCwKICAgY29tcGFyaXNvbnMsIGFuZCByZXBvcnRzIHJlYWQpIOKAlCAiQWxsIFBsYXRmb3JtcyIgc2hvd3MgYSBjb21tb24KICAgY3Jvc3MtcGxhdGZvcm0gc3VtbWFyeSwgYSBzcGVjaWZpYyBwbGF0Zm9ybSBzaG93cyBvbmx5IHRoYXQKICAgcGxhdGZvcm0ncyBjdXJhdGVkIG1ldHJpY3MuIEV2ZXJ5IGZpZWxkIG9mIGEgcmVjb3JkIChleGFjdGx5IGFzCiAgIGltcG9ydGVkKSBpcyBhbHdheXMgcmVhY2hhYmxlIHZpYSBWaWV3L0VkaXQgcmVnYXJkbGVzcyBvZiB0aGUKICAgdGFibGUncyBjdXJhdGlvbiwgd2hpY2ggcmVhZHMgdGhlIHJhd19yb3dzIG1pcnJvciBhbmQsIG9uIHNhdmUsCiAgIHJlLXN5bmNzIHBvc3RzL3Bvc3RfbWV0cmljcyBzbyBldmVyeSB2aWV3IHJlZmxlY3RzIHRoZSBjaGFuZ2UKICAgaW1tZWRpYXRlbHkuCiAgID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwpjb25zdCBSZWNvcmRzID0gKCgpID0+IHsKICBsZXQgcm9vdDsKICBsZXQgcGFnZSA9IDE7CiAgY29uc3QgcGFnZVNpemUgPSAyNTsKICBsZXQgc2VhcmNoVmFsdWUgPSAnJzsKICBsZXQgc2VhcmNoRGVib3VuY2UgPSBudWxsOwogIGxldCBtb2RhbFN0YXRlID0gbnVsbDsgLy8geyByZWNvcmQsIHZhbHVlczogWy4uLl0gfSDigJQgRWRpdCBtb2RhbCBvbmx5CgogIGZ1bmN0aW9uIHBsYXRmb3JtTWV0YSgpIHsKICAgIHJldHVybiAod2luZG93Ll9fZmlsdGVyT3B0aW9uc0NhY2hlIHx8IHsgcGxhdGZvcm1zOiBbXSB9KS5wbGF0Zm9ybXM7CiAgfQoKICBmdW5jdGlvbiBwbGF0Zm9ybUxhYmVsKGlkKSB7CiAgICBjb25zdCBtID0gcGxhdGZvcm1NZXRhKCkuZmluZCgocCkgPT4gcC5pZCA9PT0gaWQpOwogICAgcmV0dXJuIG0gPyBtLmxhYmVsIDogaWQ7CiAgfQoKICBmdW5jdGlvbiBzaGVsbCgpIHsKICAgIHJvb3QuaW5uZXJIVE1MID0gJyc7CiAgICByb290LmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgJ0RhdGEgUmVjb3JkcycsICdzZWN0aW9uLXRpdGxlJykpOwogICAgcm9vdC5hcHBlbmRDaGlsZCh0ZXh0RWwoCiAgICAgICdkaXYnLAogICAgICAnQnJvd3NlIGJ5IHBsYXRmb3JtIHRvIHNlZSBvbmx5IGl0cyBtZXRyaWNzLCBvciBzdGF5IG9uIEFsbCBQbGF0Zm9ybXMgZm9yIGEgY3Jvc3MtcGxhdGZvcm0gc3VtbWFyeS4gRXZlcnkgcmVjb3JkIHN0YXlzIGZ1bGx5IGVkaXRhYmxlIOKAlCBWaWV3IG9yIEVkaXQgYWx3YXlzIG9wZW5zIGV2ZXJ5IGZpZWxkIGltcG9ydGVkIGZyb20gdGhlIHNwcmVhZHNoZWV0LCBub3QganVzdCB3aGF04oCZcyBpbiB0aGUgdGFibGUuJywKICAgICAgJ211dGVkJwogICAgKSk7CgogICAgY29uc3QgdG9vbGJhciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdG9vbGJhci5jbGFzc05hbWUgPSAncmVjb3Jkcy10b29sYmFyJzsKICAgIGNvbnN0IHBpbGxzID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBwaWxscy5jbGFzc05hbWUgPSAncGxhdGZvcm0tZmlsdGVyLXBpbGxzJzsKICAgIHBpbGxzLmlkID0gJ3JlY29yZHNQbGF0Zm9ybVBpbGxzJzsKICAgIGNvbnN0IHNlYXJjaCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgc2VhcmNoLmNsYXNzTmFtZSA9ICdyZWNvcmRzLXNlYXJjaCc7CiAgICBjb25zdCBzZWFyY2hJbnB1dCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7CiAgICBzZWFyY2hJbnB1dC50eXBlID0gJ3NlYXJjaCc7CiAgICBzZWFyY2hJbnB1dC5wbGFjZWhvbGRlciA9ICdTZWFyY2ggY2FwdGlvbnMsIGNhbXBhaWducywgY29udGVudCB0eXBl4oCmJzsKICAgIHNlYXJjaElucHV0LmlkID0gJ3JlY29yZHNTZWFyY2hJbnB1dCc7CiAgICBzZWFyY2hJbnB1dC52YWx1ZSA9IHNlYXJjaFZhbHVlOwogICAgc2VhcmNoSW5wdXQuYWRkRXZlbnRMaXN0ZW5lcignaW5wdXQnLCAoKSA9PiB7CiAgICAgIGNsZWFyVGltZW91dChzZWFyY2hEZWJvdW5jZSk7CiAgICAgIHNlYXJjaERlYm91bmNlID0gc2V0VGltZW91dCgoKSA9PiB7CiAgICAgICAgc2VhcmNoVmFsdWUgPSBzZWFyY2hJbnB1dC52YWx1ZTsKICAgICAgICBwYWdlID0gMTsKICAgICAgICBsb2FkKCk7CiAgICAgIH0sIDMwMCk7CiAgICB9KTsKICAgIHNlYXJjaC5hcHBlbmRDaGlsZChzZWFyY2hJbnB1dCk7CiAgICB0b29sYmFyLmFwcGVuZChwaWxscywgc2VhcmNoKTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQodG9vbGJhcik7CgogICAgY29uc3QgY2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY2FyZC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICBjb25zdCB0YWJsZVdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRhYmxlV3JhcC5jbGFzc05hbWUgPSAndGFibGUtc2Nyb2xsJzsKICAgIHRhYmxlV3JhcC5pZCA9ICdyZWNvcmRzVGFibGVXcmFwJzsKICAgIGNhcmQuYXBwZW5kQ2hpbGQodGFibGVXcmFwKTsKICAgIGNvbnN0IHBhZ2VyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBwYWdlci5jbGFzc05hbWUgPSAncGFnaW5hdGlvbi1yb3cnOwogICAgcGFnZXIuaWQgPSAncmVjb3Jkc1BhZ2VyJzsKICAgIGNhcmQuYXBwZW5kQ2hpbGQocGFnZXIpOwogICAgcm9vdC5hcHBlbmRDaGlsZChjYXJkKTsKCiAgICByZW5kZXJQaWxscygpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyUGlsbHMoKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3JlY29yZHNQbGF0Zm9ybVBpbGxzJyk7CiAgICBpZiAoIXdyYXApIHJldHVybjsKICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CiAgICBjb25zdCBjdXJyZW50ID0gU3RhdGUuZ2V0RmlsdGVycygpLnBsYXRmb3JtIHx8ICdhbGwnOwogICAgY29uc3Qgb3B0aW9ucyA9IFt7IGlkOiAnYWxsJywgbGFiZWw6ICdBbGwgUGxhdGZvcm1zJywgY29sb3I6IG51bGwgfSwgLi4ucGxhdGZvcm1NZXRhKCldOwogICAgb3B0aW9ucy5mb3JFYWNoKChvcHQpID0+IHsKICAgICAgY29uc3QgYnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICAgIGJ0bi50eXBlID0gJ2J1dHRvbic7CiAgICAgIGJ0bi5jbGFzc0xpc3QudG9nZ2xlKCdpcy1hY3RpdmUnLCBjdXJyZW50ID09PSBvcHQuaWQpOwogICAgICBpZiAob3B0LmNvbG9yKSB7CiAgICAgICAgY29uc3QgZG90ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICAgIGRvdC5jbGFzc05hbWUgPSAncGxhdGZvcm0tZG90JzsKICAgICAgICBkb3Quc3R5bGUuYmFja2dyb3VuZCA9IG9wdC5jb2xvcjsKICAgICAgICBidG4uYXBwZW5kQ2hpbGQoZG90KTsKICAgICAgfQogICAgICBidG4uYXBwZW5kQ2hpbGQoZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUob3B0LmxhYmVsKSk7CiAgICAgIGJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsKICAgICAgICBpZiAoY3VycmVudCA9PT0gb3B0LmlkKSByZXR1cm47CiAgICAgICAgY29uc3QgZmlsdGVyU2VsZWN0ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZpbHRlclBsYXRmb3JtJyk7CiAgICAgICAgaWYgKGZpbHRlclNlbGVjdCkgZmlsdGVyU2VsZWN0LnZhbHVlID0gb3B0LmlkOwogICAgICAgIHBhZ2UgPSAxOwogICAgICAgIFN0YXRlLnNldEZpbHRlcnMoeyBwbGF0Zm9ybTogb3B0LmlkIH0pOwogICAgICB9KTsKICAgICAgd3JhcC5hcHBlbmRDaGlsZChidG4pOwogICAgfSk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiBsb2FkKCkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdyZWNvcmRzVGFibGVXcmFwJyk7CiAgICBpZiAod3JhcCkgeyB3cmFwLmlubmVySFRNTCA9ICcnOyB3cmFwLmFwcGVuZENoaWxkKHNrZWxldG9uUm93cyg4KSk7IH0KICAgIGNvbnN0IGZpbHRlcnMgPSBTdGF0ZS5nZXRGaWx0ZXJzKCk7CiAgICBjb25zdCByZXN1bHQgPSBhd2FpdCBBcGkucmVjb3Jkc1RhYmxlKHsgLi4uZmlsdGVycywgc2VhcmNoOiBzZWFyY2hWYWx1ZSwgcGFnZSwgcGFnZVNpemUgfSk7CiAgICByZW5kZXJUYWJsZShyZXN1bHQpOwogICAgcmVuZGVyUGFnZXIocmVzdWx0KTsKICB9CgogIGZ1bmN0aW9uIGNvbHVtbkxhYmVsc0ZvcihyZWNvcmQpIHsKICAgIHJldHVybiByZWNvcmQuaGVhZGVycyAmJiByZWNvcmQuaGVhZGVycy5sZW5ndGgKICAgICAgPyByZWNvcmQuaGVhZGVycy5tYXAoKGgpID0+IChoICYmIGgudHJpbSgpID8gaCA6ICcodW5sYWJlbGVkIGNvbHVtbiknKSkKICAgICAgOiByZWNvcmQudmFsdWVzLm1hcCgoXywgaSkgPT4gYENvbHVtbiAke2kgKyAxfWApOwogIH0KCiAgLyoqIEdyb3VwcyBhIHJhdyByZWNvcmQncyBmaWVsZHMgYnkgdGhlIHF1YWxpZmllZCBoZWFkZXIncyBwbGF0Zm9ybS1ncm91cCBwcmVmaXgKICAgICAgKGUuZy4gIkZBQ0VCT09LIOKAlCBWaWV3cyIpLCBzbyB0aGUgVmlldy9FZGl0IHBvcHVwIHJlYWRzIGFzIHNlY3Rpb25zIGluc3RlYWQKICAgICAgb2Ygb25lIGxvbmcgZmxhdCBsaXN0IOKAlCBmYWxscyBiYWNrIHRvIGEgc2luZ2xlICJEZXRhaWxzIiBzZWN0aW9uIGZvcgogICAgICBpZGVudGlmaWVyIGNvbHVtbnMgYW5kIGZvciB0aGUgc2ltcGxlIChvbmUtcGxhdGZvcm0tcGVyLXJvdykgZm9ybWF0LiAqLwogIGZ1bmN0aW9uIGdyb3VwRmllbGRSb3dzKGxhYmVscywgdmFsdWVzKSB7CiAgICBjb25zdCBncm91cHMgPSBbXTsKICAgIGNvbnN0IGluZGV4ID0gbmV3IE1hcCgpOwogICAgbGFiZWxzLmZvckVhY2goKGxhYmVsLCBpZHgpID0+IHsKICAgICAgY29uc3Qgc2VwSWR4ID0gbGFiZWwuaW5kZXhPZignIOKAlCAnKTsKICAgICAgY29uc3QgZ3JvdXBOYW1lID0gc2VwSWR4ID49IDAgPyBsYWJlbC5zbGljZSgwLCBzZXBJZHgpIDogJ0RldGFpbHMnOwogICAgICBjb25zdCBmaWVsZExhYmVsID0gc2VwSWR4ID49IDAgPyBsYWJlbC5zbGljZShzZXBJZHggKyAzKSA6IGxhYmVsOwogICAgICBpZiAoIWluZGV4Lmhhcyhncm91cE5hbWUpKSB7CiAgICAgICAgaW5kZXguc2V0KGdyb3VwTmFtZSwgeyBncm91cDogZ3JvdXBOYW1lLCBmaWVsZHM6IFtdIH0pOwogICAgICAgIGdyb3Vwcy5wdXNoKGluZGV4LmdldChncm91cE5hbWUpKTsKICAgICAgfQogICAgICBpbmRleC5nZXQoZ3JvdXBOYW1lKS5maWVsZHMucHVzaCh7IGlkeCwgbGFiZWw6IGZpZWxkTGFiZWwgfHwgYENvbHVtbiAke2lkeCArIDF9YCwgdmFsdWU6IHZhbHVlc1tpZHhdIH0pOwogICAgfSk7CiAgICByZXR1cm4gZ3JvdXBzOwogIH0KCiAgZnVuY3Rpb24gcGxhdGZvcm1CYWRnZXMoaWRzKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB3cmFwLnN0eWxlLmRpc3BsYXkgPSAnZmxleCc7CiAgICB3cmFwLnN0eWxlLmZsZXhXcmFwID0gJ3dyYXAnOwogICAgd3JhcC5zdHlsZS5nYXAgPSAnNHB4JzsKICAgIGlmICghaWRzLmxlbmd0aCkgcmV0dXJuIHRleHRFbCgnc3BhbicsICfigJQnLCAnbXV0ZWQnKTsKICAgIGNvbnN0IG1ldGEgPSBwbGF0Zm9ybU1ldGEoKTsKICAgIGlkcy5mb3JFYWNoKChpZCkgPT4gewogICAgICBjb25zdCBtID0gbWV0YS5maW5kKChwKSA9PiBwLmlkID09PSBpZCkgfHwgeyBsYWJlbDogaWQsIGNvbG9yOiAnIzk5OScgfTsKICAgICAgY29uc3QgcGlsbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgICAgcGlsbC5jbGFzc05hbWUgPSAncGxhdGZvcm0tcGlsbCc7CiAgICAgIGNvbnN0IGRvdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgICAgZG90LmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1kb3QnOwogICAgICBkb3Quc3R5bGUuYmFja2dyb3VuZCA9IG0uY29sb3I7CiAgICAgIHBpbGwuYXBwZW5kKGRvdCwgZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUobS5sYWJlbCkpOwogICAgICB3cmFwLmFwcGVuZENoaWxkKHBpbGwpOwogICAgfSk7CiAgICByZXR1cm4gd3JhcDsKICB9CgogIGZ1bmN0aW9uIHN0YXR1c1BpbGwoc3RhdHVzKSB7CiAgICBjb25zdCBzcGFuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgc3Bhbi5jbGFzc05hbWUgPSBgc3RhdHVzLXBpbGwgJHtzdGF0dXN9YDsKICAgIHNwYW4udGV4dENvbnRlbnQgPSBzdGF0dXMgPT09ICdlZGl0ZWQnID8gJ0VkaXRlZCcgOiAnT3JpZ2luYWwnOwogICAgcmV0dXJuIHNwYW47CiAgfQoKICBmdW5jdGlvbiBtZXRyaWNDZWxsKGtleSwgdmFsdWUpIHsKICAgIGlmIChrZXkgPT09ICdwb3N0aW5nX2xpbmsnKSB7CiAgICAgIGNvbnN0IHRkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgdGQuY2xhc3NOYW1lID0gJ2xpbmstY2VsbCc7CiAgICAgIGlmICh2YWx1ZSkgewogICAgICAgIGNvbnN0IGEgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdhJyk7CiAgICAgICAgYS5ocmVmID0gdmFsdWU7CiAgICAgICAgYS50YXJnZXQgPSAnX2JsYW5rJzsKICAgICAgICBhLnJlbCA9ICdub29wZW5lciBub3JlZmVycmVyJzsKICAgICAgICBhLnRleHRDb250ZW50ID0gJ09wZW4g4oaXJzsKICAgICAgICB0ZC5hcHBlbmRDaGlsZChhKTsKICAgICAgfSBlbHNlIHsKICAgICAgICB0ZC5hcHBlbmRDaGlsZChkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZSgn4oCUJykpOwogICAgICB9CiAgICAgIHJldHVybiB0ZDsKICAgIH0KICAgIGNvbnN0IGRpc3BsYXkgPSBrZXkgPT09ICd3YXRjaF90aW1lX3NlY29uZHMnID8gRm9ybWF0LmR1cmF0aW9uKHZhbHVlKSA6IEZvcm1hdC5udW1iZXIodmFsdWUpOwogICAgcmV0dXJuIHRleHRFbCgndGQnLCBkaXNwbGF5LCAnbnVtJyk7CiAgfQoKICBmdW5jdGlvbiBhY3Rpb25CdXR0b25zKHJvdywgcGxhdGZvcm0pIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHdyYXAuY2xhc3NOYW1lID0gJ3Jvdy1hY3Rpb25zJzsKICAgIGNvbnN0IHZpZXdCdG4gPSBpY29uQnRuKCdidG4nLCAnZXllJywgJ1ZpZXcnKTsKICAgIHZpZXdCdG4uZGlzYWJsZWQgPSAhcm93LnJhd1Jvd0lkOwogICAgdmlld0J0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IG9wZW5WaWV3KHJvdy5yYXdSb3dJZCkpOwogICAgY29uc3QgZWRpdEJ0biA9IGljb25CdG4oJ2J0bicsICdwZW5jaWwnLCAnRWRpdCcpOwogICAgZWRpdEJ0bi5kaXNhYmxlZCA9ICFyb3cucmF3Um93SWQ7CiAgICBlZGl0QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gb3BlbkVkaXRvcihyb3cucmF3Um93SWQpKTsKICAgIGNvbnN0IGRlbGV0ZUJ0biA9IGljb25CdG4oJ2J0biBkYW5nZXInLCAndHJhc2gtMicsICdEZWxldGUnKTsKICAgIGRlbGV0ZUJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IGhhbmRsZURlbGV0ZShyb3csIHBsYXRmb3JtKSk7CiAgICB3cmFwLmFwcGVuZCh2aWV3QnRuLCBlZGl0QnRuLCBkZWxldGVCdG4pOwogICAgcmV0dXJuIHdyYXA7CiAgfQoKICBmdW5jdGlvbiBjYXB0aW9uQ2VsbChjYXB0aW9uKSB7CiAgICBjb25zdCB0ZXh0ID0gY2FwdGlvbiB8fCAnKG5vIGNhcHRpb24pJzsKICAgIHJldHVybiB0ZXh0RWwoJ3RkJywgdGV4dC5sZW5ndGggPiA3MCA/IGAke3RleHQuc2xpY2UoMCwgNzApfeKApmAgOiB0ZXh0KTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlclN1bW1hcnlUYWJsZShyZXN1bHQpIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncmVjb3Jkc1RhYmxlV3JhcCcpOwogICAgaWYgKCFyZXN1bHQucm93cy5sZW5ndGgpIHsKICAgICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgICAgd3JhcC5hcHBlbmRDaGlsZChlbXB0eVN0YXRlKHsKICAgICAgICBpY29uOiAnZGF0YWJhc2UnLAogICAgICAgIHRpdGxlOiAnTm8gcmVjb3JkcyBtYXRjaCB0aGVzZSBmaWx0ZXJzIHlldCcsCiAgICAgICAgbWVzc2FnZTogJ1VwbG9hZCBhIHdlZWtseSBleHBvcnQsIG9yIHdpZGVuIHRoZSBkYXRlIHJhbmdlLCB0byBzZWUgcmVjb3JkcyBoZXJlLicsCiAgICAgICAgYWN0aW9uTGFiZWw6ICdVcGxvYWQgZGF0YScsCiAgICAgICAgb25BY3Rpb246ICgpID0+IGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3IoJy50YWItYnRuW2RhdGEtdGFiPSJ1cGxvYWQiXScpPy5jbGljaygpLAogICAgICB9KSk7CiAgICAgIHJldHVybjsKICAgIH0KICAgIGNvbnN0IHRhYmxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGFibGUnKTsKICAgIHRhYmxlLmNsYXNzTmFtZSA9ICdkYXRhLXRhYmxlJzsKICAgIGNvbnN0IHRoZWFkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGhlYWQnKTsKICAgIHRoZWFkLmlubmVySFRNTCA9ICc8dHI+PHRoPkRhdGU8L3RoPjx0aD5QbGF0Zm9ybXM8L3RoPjx0aD5DYXB0aW9uPC90aD48dGg+Q2FtcGFpZ248L3RoPjx0aD5Db250ZW50IFR5cGU8L3RoPjx0aD5TdGF0dXM8L3RoPjx0aD5MYXN0IFVwZGF0ZWQ8L3RoPjx0aD5BY3Rpb25zPC90aD48L3RyPic7CiAgICBjb25zdCB0Ym9keSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3Rib2R5Jyk7CiAgICByZXN1bHQucm93cy5mb3JFYWNoKChyKSA9PiB7CiAgICAgIGNvbnN0IHRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgY29uc3QgcGxhdGZvcm1zVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBwbGF0Zm9ybXNUZC5hcHBlbmRDaGlsZChwbGF0Zm9ybUJhZGdlcyhyLnBsYXRmb3JtSWRzKSk7CiAgICAgIGNvbnN0IHN0YXR1c1RkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgc3RhdHVzVGQuYXBwZW5kQ2hpbGQoc3RhdHVzUGlsbChyLnN0YXR1cykpOwogICAgICBjb25zdCBhY3Rpb25zVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBhY3Rpb25zVGQuYXBwZW5kQ2hpbGQoYWN0aW9uQnV0dG9ucyhyLCAnYWxsJykpOwogICAgICB0ci5hcHBlbmQoCiAgICAgICAgdGV4dEVsKCd0ZCcsIEZvcm1hdC5kYXRlKHIucHVibGlzaERhdGUpKSwKICAgICAgICBwbGF0Zm9ybXNUZCwKICAgICAgICBjYXB0aW9uQ2VsbChyLmNhcHRpb24pLAogICAgICAgIHRleHRFbCgndGQnLCByLmNhbXBhaWduVHlwZSB8fCAn4oCUJyksCiAgICAgICAgdGV4dEVsKCd0ZCcsIHIuY29udGVudFR5cGUgfHwgJ+KAlCcpLAogICAgICAgIHN0YXR1c1RkLAogICAgICAgIHRleHRFbCgndGQnLCByLnVwZGF0ZWRBdCksCiAgICAgICAgYWN0aW9uc1RkCiAgICAgICk7CiAgICAgIHRib2R5LmFwcGVuZENoaWxkKHRyKTsKICAgIH0pOwogICAgdGFibGUuYXBwZW5kKHRoZWFkLCB0Ym9keSk7CiAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgd3JhcC5hcHBlbmRDaGlsZCh0YWJsZSk7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJQbGF0Zm9ybVRhYmxlKHJlc3VsdCkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdyZWNvcmRzVGFibGVXcmFwJyk7CiAgICBpZiAoIXJlc3VsdC5yb3dzLmxlbmd0aCkgewogICAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgICB3cmFwLmFwcGVuZENoaWxkKGVtcHR5U3RhdGUoewogICAgICAgIGljb246ICdkYXRhYmFzZScsCiAgICAgICAgdGl0bGU6IGBObyAke3BsYXRmb3JtTGFiZWwocmVzdWx0LnBsYXRmb3JtKX0gcmVjb3JkcyBtYXRjaCB0aGVzZSBmaWx0ZXJzIHlldGAsCiAgICAgICAgbWVzc2FnZTogJ1RyeSBhIGRpZmZlcmVudCBwbGF0Zm9ybSwgb3Igd2lkZW4gdGhlIGRhdGUgcmFuZ2UuJywKICAgICAgfSkpOwogICAgICByZXR1cm47CiAgICB9CiAgICBjb25zdCB0YWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICB0YWJsZS5jbGFzc05hbWUgPSAnZGF0YS10YWJsZSc7CiAgICBjb25zdCB0aGVhZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RoZWFkJyk7CiAgICBjb25zdCBoZWFkVHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgaGVhZFRyLmFwcGVuZCh0ZXh0RWwoJ3RoJywgJ0RhdGUnKSwgdGV4dEVsKCd0aCcsICdDYXB0aW9uJyksIHRleHRFbCgndGgnLCAnQ2FtcGFpZ24nKSwgdGV4dEVsKCd0aCcsICdDb250ZW50IFR5cGUnKSk7CiAgICByZXN1bHQuY29sdW1ucy5mb3JFYWNoKChjKSA9PiBoZWFkVHIuYXBwZW5kQ2hpbGQodGV4dEVsKCd0aCcsIGMubGFiZWwsIGMua2V5ID09PSAncG9zdGluZ19saW5rJyA/ICcnIDogJ251bScpKSk7CiAgICBoZWFkVHIuYXBwZW5kKHRleHRFbCgndGgnLCAnU3RhdHVzJyksIHRleHRFbCgndGgnLCAnQWN0aW9ucycpKTsKICAgIHRoZWFkLmFwcGVuZENoaWxkKGhlYWRUcik7CiAgICBjb25zdCB0Ym9keSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3Rib2R5Jyk7CiAgICByZXN1bHQucm93cy5mb3JFYWNoKChyKSA9PiB7CiAgICAgIGNvbnN0IHRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgdHIuYXBwZW5kKHRleHRFbCgndGQnLCBGb3JtYXQuZGF0ZShyLnB1Ymxpc2hEYXRlKSksIGNhcHRpb25DZWxsKHIuY2FwdGlvbiksIHRleHRFbCgndGQnLCByLmNhbXBhaWduVHlwZSB8fCAn4oCUJyksIHRleHRFbCgndGQnLCByLmNvbnRlbnRUeXBlIHx8ICfigJQnKSk7CiAgICAgIHJlc3VsdC5jb2x1bW5zLmZvckVhY2goKGMpID0+IHRyLmFwcGVuZENoaWxkKG1ldHJpY0NlbGwoYy5rZXksIHIubWV0cmljc1tjLmtleV0pKSk7CiAgICAgIGNvbnN0IHN0YXR1c1RkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgc3RhdHVzVGQuYXBwZW5kQ2hpbGQoc3RhdHVzUGlsbChyLnN0YXR1cykpOwogICAgICB0ci5hcHBlbmRDaGlsZChzdGF0dXNUZCk7CiAgICAgIGNvbnN0IGFjdGlvbnNUZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIGFjdGlvbnNUZC5hcHBlbmRDaGlsZChhY3Rpb25CdXR0b25zKHIsIHJlc3VsdC5wbGF0Zm9ybSkpOwogICAgICB0ci5hcHBlbmRDaGlsZChhY3Rpb25zVGQpOwogICAgICB0Ym9keS5hcHBlbmRDaGlsZCh0cik7CiAgICB9KTsKICAgIHRhYmxlLmFwcGVuZCh0aGVhZCwgdGJvZHkpOwogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGFibGUpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyVGFibGUocmVzdWx0KSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3JlY29yZHNUYWJsZVdyYXAnKTsKICAgIGlmICghd3JhcCkgcmV0dXJuOwogICAgaWYgKHJlc3VsdC5wbGF0Zm9ybSA9PT0gJ2FsbCcpIHJlbmRlclN1bW1hcnlUYWJsZShyZXN1bHQpOwogICAgZWxzZSByZW5kZXJQbGF0Zm9ybVRhYmxlKHJlc3VsdCk7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJQYWdlcihyZXN1bHQpIHsKICAgIGNvbnN0IHBhZ2VyID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3JlY29yZHNQYWdlcicpOwogICAgaWYgKCFwYWdlcikgcmV0dXJuOwogICAgcGFnZXIuaW5uZXJIVE1MID0gJyc7CiAgICBjb25zdCB0b3RhbFBhZ2VzID0gTWF0aC5tYXgoMSwgTWF0aC5jZWlsKHJlc3VsdC50b3RhbCAvIHJlc3VsdC5wYWdlU2l6ZSkpOwogICAgY29uc3QgcHJldkJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgcHJldkJ0bi5jbGFzc05hbWUgPSAnYnRuJzsKICAgIHByZXZCdG4udGV4dENvbnRlbnQgPSAnUHJldmlvdXMnOwogICAgcHJldkJ0bi5kaXNhYmxlZCA9IHJlc3VsdC5wYWdlIDw9IDE7CiAgICBwcmV2QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4geyBwYWdlIC09IDE7IGxvYWQoKTsgfSk7CiAgICBjb25zdCBuZXh0QnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBuZXh0QnRuLmNsYXNzTmFtZSA9ICdidG4nOwogICAgbmV4dEJ0bi50ZXh0Q29udGVudCA9ICdOZXh0JzsKICAgIG5leHRCdG4uZGlzYWJsZWQgPSByZXN1bHQucGFnZSA+PSB0b3RhbFBhZ2VzOwogICAgbmV4dEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsgcGFnZSArPSAxOyBsb2FkKCk7IH0pOwogICAgcGFnZXIuYXBwZW5kKHByZXZCdG4sIHRleHRFbCgnc3BhbicsIGBQYWdlICR7cmVzdWx0LnBhZ2V9IG9mICR7dG90YWxQYWdlc30g4oCUICR7cmVzdWx0LnRvdGFsfSByZWNvcmQocylgKSwgbmV4dEJ0bik7CiAgfQoKICBhc3luYyBmdW5jdGlvbiBoYW5kbGVEZWxldGUocm93LCBwbGF0Zm9ybSkgewogICAgY29uc3QgY2FwdGlvbiA9IChyb3cuY2FwdGlvbiB8fCAnKG5vIGNhcHRpb24pJykuc2xpY2UoMCwgNjApOwogICAgY29uc3QgbWVzc2FnZSA9IHBsYXRmb3JtID09PSAnYWxsJwogICAgICA/IGBEZWxldGUgdGhpcyBlbnRpcmUgcmVjb3JkIOKAlCAiJHtjYXB0aW9ufSIg4oCUIGFjcm9zcyBldmVyeSBwbGF0Zm9ybT8gSXRzIG9yaWdpbmFsIGltcG9ydCBzdGF5cyBpbiBVcGxvYWQgSGlzdG9yeSwgYnV0IGl0IHdpbGwgZGlzYXBwZWFyIGZyb20gdGhlIGRhc2hib2FyZCwgY29tcGFyaXNvbnMsIGFuZCByZXBvcnRzLmAKICAgICAgOiBgUmVtb3ZlIHRoaXMgcmVjb3JkJ3MgJHtwbGF0Zm9ybUxhYmVsKHBsYXRmb3JtKX0gZGF0YSDigJQgIiR7Y2FwdGlvbn0iPyBJZiB0aGlzIGlzIGl0cyBvbmx5IHBsYXRmb3JtLCB0aGUgd2hvbGUgcmVjb3JkIHdpbGwgYmUgcmVtb3ZlZCBmcm9tIHRoZSBkYXNoYm9hcmQuYDsKICAgIGlmICghd2luZG93LmNvbmZpcm0obWVzc2FnZSkpIHJldHVybjsKICAgIHRyeSB7CiAgICAgIGlmIChwbGF0Zm9ybSA9PT0gJ2FsbCcpIGF3YWl0IEFwaS5kZWxldGVSZWNvcmRQb3N0KHJvdy5wb3N0SWQpOwogICAgICBlbHNlIGF3YWl0IEFwaS5kZWxldGVSZWNvcmRQbGF0Zm9ybShyb3cucG9zdElkLCBwbGF0Zm9ybSk7CiAgICAgIFRvYXN0LnNob3coJ1JlY29yZCBkZWxldGVkLicsICdzdWNjZXNzJyk7CiAgICAgIGF3YWl0IGxvYWQoKTsKICAgICAgd2luZG93LmRpc3BhdGNoRXZlbnQobmV3IEN1c3RvbUV2ZW50KCdscnM6ZGF0YS11cGRhdGVkJykpOwogICAgfSBjYXRjaCAoZXJyKSB7CiAgICAgIFRvYXN0LnNob3coZXJyLm1lc3NhZ2UsICdlcnJvcicpOwogICAgfQogIH0KCiAgZnVuY3Rpb24gY2xvc2VNb2RhbCgpIHsKICAgIGNvbnN0IG92ZXJsYXkgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncmVjb3JkTW9kYWxPdmVybGF5Jyk7CiAgICBpZiAob3ZlcmxheSkgb3ZlcmxheS5yZW1vdmUoKTsKICAgIG1vZGFsU3RhdGUgPSBudWxsOwogIH0KCiAgZnVuY3Rpb24gbW9kYWxTaGVsbCh0aXRsZVRleHQpIHsKICAgIGNsb3NlTW9kYWwoKTsKICAgIGNvbnN0IG92ZXJsYXkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIG92ZXJsYXkuY2xhc3NOYW1lID0gJ21vZGFsLW92ZXJsYXknOwogICAgb3ZlcmxheS5pZCA9ICdyZWNvcmRNb2RhbE92ZXJsYXknOwogICAgb3ZlcmxheS5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIChlKSA9PiB7IGlmIChlLnRhcmdldCA9PT0gb3ZlcmxheSkgY2xvc2VNb2RhbCgpOyB9KTsKICAgIGNvbnN0IHBhbmVsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBwYW5lbC5jbGFzc05hbWUgPSAnbW9kYWwtcGFuZWwgd2lkZSc7CiAgICBwYW5lbC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2gyJywgdGl0bGVUZXh0KSk7CiAgICBvdmVybGF5LmFwcGVuZENoaWxkKHBhbmVsKTsKICAgIHJldHVybiB7IG92ZXJsYXksIHBhbmVsIH07CiAgfQoKICBmdW5jdGlvbiByZWNvcmRTdWJ0aXRsZShyKSB7CiAgICByZXR1cm4gYFNoZWV0ICIke3Iuc2hlZXROYW1lfSIsIHJvdyAke3Iucm93TnVtYmVyfSR7ci5wb3N0SWQgPyBgIOKAlCBsaW5rZWQgdG8gZGFzaGJvYXJkIHBvc3QgIyR7ci5wb3N0SWR9YCA6ICcg4oCUIG5vdCBwYXJ0IG9mIHRoZSBkYXNoYm9hcmQgKGUuZy4gbmVlZHMgYSB2YWxpZCBkYXRlKSd9YDsKICB9CgogIC8vIC0tLS0tLS0tLS0gVmlldyBwb3B1cDogcmVhZC1vbmx5LCBldmVyeSBmaWVsZCwgZ3JvdXBlZCBpbnRvIHNlY3Rpb25zIC0tLS0tLS0tLS0KICBhc3luYyBmdW5jdGlvbiBvcGVuVmlldyhpZCkgewogICAgY29uc3QgcmVjb3JkID0gYXdhaXQgQXBpLmdldFJlY29yZChpZCk7CiAgICBjb25zdCB7IG92ZXJsYXksIHBhbmVsIH0gPSBtb2RhbFNoZWxsKCdSZWNvcmQgZGV0YWlscycpOwogICAgcGFuZWwuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCByZWNvcmRTdWJ0aXRsZShyZWNvcmQpLCAnbW9kYWwtc3ViJykpOwoKICAgIGNvbnN0IGdyb3VwcyA9IGdyb3VwRmllbGRSb3dzKGNvbHVtbkxhYmVsc0ZvcihyZWNvcmQpLCByZWNvcmQudmFsdWVzKTsKICAgIGdyb3Vwcy5mb3JFYWNoKChnKSA9PiB7CiAgICAgIGNvbnN0IHNlY3Rpb24gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgc2VjdGlvbi5jbGFzc05hbWUgPSAncmVjb3JkLXNlY3Rpb24nOwogICAgICBzZWN0aW9uLmFwcGVuZENoaWxkKHRleHRFbCgnaDQnLCBnLmdyb3VwKSk7CiAgICAgIGNvbnN0IGdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgZ3JpZC5jbGFzc05hbWUgPSAnZm9ybS1ncmlkJzsKICAgICAgZy5maWVsZHMuZm9yRWFjaCgoZikgPT4gewogICAgICAgIGNvbnN0IGZpZWxkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgICAgZmllbGQuY2xhc3NOYW1lID0gJ3ZpZXctZmllbGQnOwogICAgICAgIGZpZWxkLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgZi5sYWJlbCwgJ3ZpZXctbGFiZWwnKSk7CiAgICAgICAgY29uc3QgdmFsID0gZi52YWx1ZSA9PT0gdW5kZWZpbmVkIHx8IGYudmFsdWUgPT09IG51bGwgfHwgZi52YWx1ZSA9PT0gJycgPyAn4oCUJyA6IFN0cmluZyhmLnZhbHVlKTsKICAgICAgICBmaWVsZC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIHZhbCwgJ3ZpZXctdmFsdWUnKSk7CiAgICAgICAgZ3JpZC5hcHBlbmRDaGlsZChmaWVsZCk7CiAgICAgIH0pOwogICAgICBzZWN0aW9uLmFwcGVuZENoaWxkKGdyaWQpOwogICAgICBwYW5lbC5hcHBlbmRDaGlsZChzZWN0aW9uKTsKICAgIH0pOwoKICAgIGNvbnN0IGFjdGlvbnMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGFjdGlvbnMuY2xhc3NOYW1lID0gJ21vZGFsLWFjdGlvbnMnOwogICAgY29uc3QgYnRuUm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBidG5Sb3cuY2xhc3NOYW1lID0gJ2J0bi1yb3cnOwogICAgY29uc3QgY2xvc2VCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIGNsb3NlQnRuLmNsYXNzTmFtZSA9ICdidG4nOwogICAgY2xvc2VCdG4udGV4dENvbnRlbnQgPSAnQ2xvc2UnOwogICAgY2xvc2VCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBjbG9zZU1vZGFsKTsKICAgIGNvbnN0IGVkaXRCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIGVkaXRCdG4uY2xhc3NOYW1lID0gJ2J0biBwcmltYXJ5JzsKICAgIGVkaXRCdG4udGV4dENvbnRlbnQgPSAnRWRpdCB0aGlzIHJlY29yZCc7CiAgICBlZGl0QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gb3BlbkVkaXRvcihyZWNvcmQuaWQpKTsKICAgIGJ0blJvdy5hcHBlbmQoY2xvc2VCdG4sIGVkaXRCdG4pOwogICAgYWN0aW9ucy5hcHBlbmRDaGlsZChidG5Sb3cpOwogICAgcGFuZWwuYXBwZW5kQ2hpbGQoYWN0aW9ucyk7CgogICAgZG9jdW1lbnQuYm9keS5hcHBlbmRDaGlsZChvdmVybGF5KTsKICB9CgogIC8vIC0tLS0tLS0tLS0gRWRpdCBwb3B1cDogZXZlcnkgZmllbGQsIGdyb3VwZWQgaW50byBzZWN0aW9ucywgYWxsIGVkaXRhYmxlIC0tLS0tLS0tLS0KICBhc3luYyBmdW5jdGlvbiBvcGVuRWRpdG9yKGlkKSB7CiAgICBjb25zdCByZWNvcmQgPSBhd2FpdCBBcGkuZ2V0UmVjb3JkKGlkKTsKICAgIG1vZGFsU3RhdGUgPSB7IHJlY29yZCwgdmFsdWVzOiBbLi4ucmVjb3JkLnZhbHVlc10gfTsKICAgIHJlbmRlckVkaXRNb2RhbCgpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyRWRpdE1vZGFsKCkgewogICAgY29uc3QgciA9IG1vZGFsU3RhdGUucmVjb3JkOwogICAgY29uc3QgeyBvdmVybGF5LCBwYW5lbCB9ID0gbW9kYWxTaGVsbCgnRWRpdCByZWNvcmQnKTsKICAgIHBhbmVsLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgcmVjb3JkU3VidGl0bGUociksICdtb2RhbC1zdWInKSk7CgogICAgY29uc3QgZ3JvdXBzID0gZ3JvdXBGaWVsZFJvd3MoY29sdW1uTGFiZWxzRm9yKHIpLCBtb2RhbFN0YXRlLnZhbHVlcyk7CiAgICBncm91cHMuZm9yRWFjaCgoZykgPT4gewogICAgICBjb25zdCBzZWN0aW9uID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIHNlY3Rpb24uY2xhc3NOYW1lID0gJ3JlY29yZC1zZWN0aW9uJzsKICAgICAgc2VjdGlvbi5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2g0JywgZy5ncm91cCkpOwogICAgICBjb25zdCBncmlkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIGdyaWQuY2xhc3NOYW1lID0gJ2Zvcm0tZ3JpZCc7CiAgICAgIGcuZmllbGRzLmZvckVhY2goKGYpID0+IHsKICAgICAgICBjb25zdCBmaWVsZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICAgIGZpZWxkLmNsYXNzTmFtZSA9ICdmb3JtLWZpZWxkJzsKICAgICAgICBmaWVsZC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2xhYmVsJywgZi5sYWJlbCkpOwogICAgICAgIGNvbnN0IHN0clZhbCA9IGYudmFsdWUgPT09IHVuZGVmaW5lZCB8fCBmLnZhbHVlID09PSBudWxsID8gJycgOiBTdHJpbmcoZi52YWx1ZSk7CiAgICAgICAgY29uc3QgaXNMb25nID0gc3RyVmFsLmxlbmd0aCA+IDgwOwogICAgICAgIGNvbnN0IGlucHV0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudChpc0xvbmcgPyAndGV4dGFyZWEnIDogJ2lucHV0Jyk7CiAgICAgICAgaWYgKCFpc0xvbmcpIGlucHV0LnR5cGUgPSAndGV4dCc7CiAgICAgICAgZWxzZSBmaWVsZC5zdHlsZS5ncmlkQ29sdW1uID0gJzEgLyAtMSc7CiAgICAgICAgaW5wdXQudmFsdWUgPSBzdHJWYWw7CiAgICAgICAgaW5wdXQuYWRkRXZlbnRMaXN0ZW5lcignaW5wdXQnLCAoKSA9PiB7IG1vZGFsU3RhdGUudmFsdWVzW2YuaWR4XSA9IGlucHV0LnZhbHVlOyB9KTsKICAgICAgICBmaWVsZC5hcHBlbmRDaGlsZChpbnB1dCk7CiAgICAgICAgZ3JpZC5hcHBlbmRDaGlsZChmaWVsZCk7CiAgICAgIH0pOwogICAgICBzZWN0aW9uLmFwcGVuZENoaWxkKGdyaWQpOwogICAgICBwYW5lbC5hcHBlbmRDaGlsZChzZWN0aW9uKTsKICAgIH0pOwoKICAgIGNvbnN0IGFjdGlvbnMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGFjdGlvbnMuY2xhc3NOYW1lID0gJ21vZGFsLWFjdGlvbnMnOwogICAgY29uc3QgZXJyb3JNc2cgPSB0ZXh0RWwoJ3NwYW4nLCAnJywgJ211dGVkJyk7CiAgICBlcnJvck1zZy5pZCA9ICdtb2RhbEVycm9yTXNnJzsKICAgIGNvbnN0IGJ0blJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYnRuUm93LmNsYXNzTmFtZSA9ICdidG4tcm93JzsKICAgIGNvbnN0IGNhbmNlbEJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgY2FuY2VsQnRuLmNsYXNzTmFtZSA9ICdidG4nOwogICAgY2FuY2VsQnRuLnRleHRDb250ZW50ID0gJ0NhbmNlbCc7CiAgICBjYW5jZWxCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBjbG9zZU1vZGFsKTsKICAgIGNvbnN0IHNhdmVCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIHNhdmVCdG4uY2xhc3NOYW1lID0gJ2J0biBwcmltYXJ5JzsKICAgIHNhdmVCdG4udGV4dENvbnRlbnQgPSAnU2F2ZSBjaGFuZ2VzJzsKICAgIHNhdmVCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBzYXZlRWRpdChzYXZlQnRuKSk7CiAgICBidG5Sb3cuYXBwZW5kKGNhbmNlbEJ0biwgc2F2ZUJ0bik7CiAgICBhY3Rpb25zLmFwcGVuZChlcnJvck1zZywgYnRuUm93KTsKICAgIHBhbmVsLmFwcGVuZENoaWxkKGFjdGlvbnMpOwoKICAgIGRvY3VtZW50LmJvZHkuYXBwZW5kQ2hpbGQob3ZlcmxheSk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiBzYXZlRWRpdChidG4pIHsKICAgIGNvbnN0IGVycm9yRWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbW9kYWxFcnJvck1zZycpOwogICAgZXJyb3JFbC50ZXh0Q29udGVudCA9ICcnOwogICAgYnRuLmRpc2FibGVkID0gdHJ1ZTsKICAgIGJ0bi50ZXh0Q29udGVudCA9ICdTYXZpbmfigKYnOwogICAgdHJ5IHsKICAgICAgYXdhaXQgQXBpLnVwZGF0ZVJlY29yZChtb2RhbFN0YXRlLnJlY29yZC5pZCwgbW9kYWxTdGF0ZS52YWx1ZXMpOwogICAgICBUb2FzdC5zaG93KCdSZWNvcmQgdXBkYXRlZC4nLCAnc3VjY2VzcycpOwogICAgICBjbG9zZU1vZGFsKCk7CiAgICAgIGF3YWl0IGxvYWQoKTsKICAgICAgd2luZG93LmRpc3BhdGNoRXZlbnQobmV3IEN1c3RvbUV2ZW50KCdscnM6ZGF0YS11cGRhdGVkJykpOwogICAgfSBjYXRjaCAoZXJyKSB7CiAgICAgIGVycm9yRWwudGV4dENvbnRlbnQgPSBlcnIubWVzc2FnZTsKICAgICAgZXJyb3JFbC5zdHlsZS5jb2xvciA9ICd2YXIoLS1zdGF0dXMtY3JpdGljYWwpJzsKICAgICAgYnRuLmRpc2FibGVkID0gZmFsc2U7CiAgICAgIGJ0bi50ZXh0Q29udGVudCA9ICdTYXZlIGNoYW5nZXMnOwogICAgfQogIH0KCiAgYXN5bmMgZnVuY3Rpb24gcmVuZGVyKCkgewogICAgcm9vdCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd2aWV3LXJlY29yZHMnKTsKICAgIHBhZ2UgPSAxOwogICAgc2hlbGwoKTsKICAgIGF3YWl0IGxvYWQoKTsKICB9CgogIHJldHVybiB7IHJlbmRlciwgcmVsb2FkOiBsb2FkLCBvcGVuVmlldyB9Owp9KSgpOwoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIENvbXBhcmlzb25zIHRhYjogd2Vlay12cy13ZWVrLCBjdXN0b20gcmFuZ2UsIG1vbnRobHksCiAgIHF1YXJ0ZXJseSwgWVRELgogICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KY29uc3QgQ29tcGFyaXNvbiA9ICgoKSA9PiB7CiAgY29uc3QgTU9ERVMgPSBbCiAgICB7IGtleTogJ3dlZWsnLCBsYWJlbDogJ1dlZWsgdnMgV2VlaycgfSwKICAgIHsga2V5OiAnY3VzdG9tJywgbGFiZWw6ICdDdXN0b20gUmFuZ2UnIH0sCiAgICB7IGtleTogJ21vbnRoJywgbGFiZWw6ICdNb250aGx5JyB9LAogICAgeyBrZXk6ICdxdWFydGVyJywgbGFiZWw6ICdRdWFydGVybHknIH0sCiAgICB7IGtleTogJ3l0ZCcsIGxhYmVsOiAnWWVhciB0byBEYXRlJyB9LAogIF07CiAgY29uc3QgTUVUUklDX1JPV1MgPSBbCiAgICB7IGtleTogJ3ZpZXdzJywgbGFiZWw6ICdWaWV3cycgfSwKICAgIHsga2V5OiAncmVhY2gnLCBsYWJlbDogJ1JlYWNoJyB9LAogICAgeyBrZXk6ICdpbXByZXNzaW9ucycsIGxhYmVsOiAnSW1wcmVzc2lvbnMnIH0sCiAgICB7IGtleTogJ2VuZ2FnZW1lbnQnLCBsYWJlbDogJ0VuZ2FnZW1lbnQnIH0sCiAgICB7IGtleTogJ2NsaWNrcycsIGxhYmVsOiAnQ2xpY2tzJyB9LAogICAgeyBrZXk6ICdmb2xsb3dlcnNfZ2FpbmVkJywgbGFiZWw6ICdGb2xsb3dlcnMgR2FpbmVkJyB9LAogICAgeyBrZXk6ICd3YXRjaF90aW1lX3NlY29uZHMnLCBsYWJlbDogJ1dhdGNoIFRpbWUnIH0sCiAgICB7IGtleTogJ3NoYXJlcycsIGxhYmVsOiAnU2hhcmVzJyB9LAogICAgeyBrZXk6ICdjb21tZW50cycsIGxhYmVsOiAnQ29tbWVudHMnIH0sCiAgICB7IGtleTogJ3NhdmVzJywgbGFiZWw6ICdTYXZlcycgfSwKICBdOwoKICBsZXQgbW9kZSA9ICd3ZWVrJzsKICBsZXQgcm9vdDsKICBsZXQgY2hhcnRNZXRyaWMgPSAnZW5nYWdlbWVudCc7CgogIGZ1bmN0aW9uIG1vbmRheU9mKGRhdGVTdHIpIHsKICAgIGNvbnN0IGQgPSBuZXcgRGF0ZShkYXRlU3RyKTsKICAgIGNvbnN0IGRheSA9IGQuZ2V0RGF5KCk7CiAgICBjb25zdCBkaWZmID0gZGF5ID09PSAwID8gNiA6IGRheSAtIDE7CiAgICBkLnNldERhdGUoZC5nZXREYXRlKCkgLSBkaWZmKTsKICAgIHJldHVybiBkLnRvSVNPU3RyaW5nKCkuc2xpY2UoMCwgMTApOwogIH0KICBmdW5jdGlvbiBhZGREYXlzKGRhdGVTdHIsIG4pIHsKICAgIGNvbnN0IGQgPSBuZXcgRGF0ZShkYXRlU3RyKTsKICAgIGQuc2V0RGF0ZShkLmdldERhdGUoKSArIG4pOwogICAgcmV0dXJuIGQudG9JU09TdHJpbmcoKS5zbGljZSgwLCAxMCk7CiAgfQoKICBmdW5jdGlvbiBzaGVsbCgpIHsKICAgIHJvb3QuaW5uZXJIVE1MID0gJyc7CgogICAgY29uc3QgdGFicyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGFicy5jbGFzc05hbWUgPSAnbW9kZS10YWJzJzsKICAgIE1PREVTLmZvckVhY2goKG0pID0+IHsKICAgICAgY29uc3QgYnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICAgIGJ0bi50ZXh0Q29udGVudCA9IG0ubGFiZWw7CiAgICAgIGJ0bi50eXBlID0gJ2J1dHRvbic7CiAgICAgIGlmIChtLmtleSA9PT0gbW9kZSkgYnRuLmNsYXNzTGlzdC5hZGQoJ2lzLWFjdGl2ZScpOwogICAgICBidG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7CiAgICAgICAgbW9kZSA9IG0ua2V5OwogICAgICAgIHNoZWxsKCk7CiAgICAgIH0pOwogICAgICB0YWJzLmFwcGVuZENoaWxkKGJ0bik7CiAgICB9KTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQodGFicyk7CgogICAgY29uc3QgY29udHJvbHMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNvbnRyb2xzLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGNvbnRyb2xzLmlkID0gJ2NvbXBhcmlzb25Db250cm9scyc7CiAgICByb290LmFwcGVuZENoaWxkKGNvbnRyb2xzKTsKCiAgICBjb25zdCByZXN1bHRzID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICByZXN1bHRzLmlkID0gJ2NvbXBhcmlzb25SZXN1bHRzJzsKICAgIHJvb3QuYXBwZW5kQ2hpbGQocmVzdWx0cyk7CgogICAgcmVuZGVyQ29udHJvbHMoKTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlckNvbnRyb2xzKCkgewogICAgY29uc3QgY29udHJvbHMgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnY29tcGFyaXNvbkNvbnRyb2xzJyk7CiAgICBjb250cm9scy5pbm5lckhUTUwgPSAnJzsKICAgIGNvbnN0IHJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcm93LmNsYXNzTmFtZSA9ICdidG4tcm93JzsKICAgIHJvdy5zdHlsZS5hbGlnbkl0ZW1zID0gJ2VuZCc7CgogICAgY29uc3QgdG9kYXkgPSBuZXcgRGF0ZSgpLnRvSVNPU3RyaW5nKCkuc2xpY2UoMCwgMTApOwogICAgY29uc3QgdGhpc1llYXIgPSBuZXcgRGF0ZSgpLmdldEZ1bGxZZWFyKCk7CgogICAgaWYgKG1vZGUgPT09ICd3ZWVrJykgewogICAgICBjb25zdCB3QSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7IHdBLnR5cGUgPSAnZGF0ZSc7IHdBLnZhbHVlID0gbW9uZGF5T2YodG9kYXkpOwogICAgICBjb25zdCB3QiA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7IHdCLnR5cGUgPSAnZGF0ZSc7IHdCLnZhbHVlID0gbW9uZGF5T2YoYWRkRGF5cyh0b2RheSwgLTcpKTsKICAgICAgcm93LmFwcGVuZChsYWJlbGVkKCdXZWVrIEEgKGFueSBkYXkgaW4gd2VlayknLCB3QSksIGxhYmVsZWQoJ1dlZWsgQiAoYW55IGRheSBpbiB3ZWVrKScsIHdCKSwgcnVuQnRuKCgpID0+IHsKICAgICAgICBjb25zdCByYW5nZUEgPSB7IGZyb206IG1vbmRheU9mKHdBLnZhbHVlKSwgdG86IGFkZERheXMobW9uZGF5T2Yod0EudmFsdWUpLCA2KSB9OwogICAgICAgIGNvbnN0IHJhbmdlQiA9IHsgZnJvbTogbW9uZGF5T2Yod0IudmFsdWUpLCB0bzogYWRkRGF5cyhtb25kYXlPZih3Qi52YWx1ZSksIDYpIH07CiAgICAgICAgcnVuQ29tcGFyZShyYW5nZUEsIHJhbmdlQiwgYFdlZWsgb2YgJHtGb3JtYXQuZGF0ZShyYW5nZUEuZnJvbSl9YCwgYFdlZWsgb2YgJHtGb3JtYXQuZGF0ZShyYW5nZUIuZnJvbSl9YCk7CiAgICAgIH0pKTsKICAgIH0gZWxzZSBpZiAobW9kZSA9PT0gJ2N1c3RvbScpIHsKICAgICAgY29uc3QgZkEgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOyBmQS50eXBlID0gJ2RhdGUnOyBmQS52YWx1ZSA9IGFkZERheXModG9kYXksIC0xMyk7CiAgICAgIGNvbnN0IHRBID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsgdEEudHlwZSA9ICdkYXRlJzsgdEEudmFsdWUgPSB0b2RheTsKICAgICAgY29uc3QgZkIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOyBmQi50eXBlID0gJ2RhdGUnOyBmQi52YWx1ZSA9IGFkZERheXModG9kYXksIC0yNyk7CiAgICAgIGNvbnN0IHRCID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsgdEIudHlwZSA9ICdkYXRlJzsgdEIudmFsdWUgPSBhZGREYXlzKHRvZGF5LCAtMTQpOwogICAgICByb3cuYXBwZW5kKAogICAgICAgIGxhYmVsZWQoJ1JhbmdlIEEgZnJvbScsIGZBKSwgbGFiZWxlZCgndG8nLCB0QSksCiAgICAgICAgbGFiZWxlZCgnUmFuZ2UgQiBmcm9tJywgZkIpLCBsYWJlbGVkKCd0bycsIHRCKSwKICAgICAgICBydW5CdG4oKCkgPT4gcnVuQ29tcGFyZSh7IGZyb206IGZBLnZhbHVlLCB0bzogdEEudmFsdWUgfSwgeyBmcm9tOiBmQi52YWx1ZSwgdG86IHRCLnZhbHVlIH0sICdSYW5nZSBBJywgJ1JhbmdlIEInKSkKICAgICAgKTsKICAgIH0gZWxzZSBpZiAobW9kZSA9PT0gJ21vbnRoJykgewogICAgICBjb25zdCB5ID0geWVhclNlbGVjdCh0aGlzWWVhcik7IGNvbnN0IG0gPSBtb250aFNlbGVjdChuZXcgRGF0ZSgpLmdldE1vbnRoKCkgKyAxKTsKICAgICAgY29uc3QgdG9nZ2xlID0gcGVyaW9kVG9nZ2xlKCk7CiAgICAgIHJvdy5hcHBlbmQobGFiZWxlZCgnWWVhcicsIHkpLCBsYWJlbGVkKCdNb250aCcsIG0pLCB0b2dnbGUuZWwsIHJ1bkJ0bihhc3luYyAoKSA9PiB7CiAgICAgICAgY29uc3QgcmVwb3J0ID0gYXdhaXQgQXBpLm1vbnRobHkoeyB5ZWFyOiB5LnZhbHVlLCBtb250aDogbS52YWx1ZSwgLi4uU3RhdGUuZ2V0RmlsdGVycygpIH0pOwogICAgICAgIHJlbmRlclBlcmlvZFJlcG9ydChyZXBvcnQsIHRvZ2dsZS5nZXQoKSk7CiAgICAgIH0pKTsKICAgIH0gZWxzZSBpZiAobW9kZSA9PT0gJ3F1YXJ0ZXInKSB7CiAgICAgIGNvbnN0IHkgPSB5ZWFyU2VsZWN0KHRoaXNZZWFyKTsgY29uc3QgcSA9IHF1YXJ0ZXJTZWxlY3QoKTsKICAgICAgY29uc3QgdG9nZ2xlID0gcGVyaW9kVG9nZ2xlKCk7CiAgICAgIHJvdy5hcHBlbmQobGFiZWxlZCgnWWVhcicsIHkpLCBsYWJlbGVkKCdRdWFydGVyJywgcSksIHRvZ2dsZS5lbCwgcnVuQnRuKGFzeW5jICgpID0+IHsKICAgICAgICBjb25zdCByZXBvcnQgPSBhd2FpdCBBcGkucXVhcnRlcmx5KHsgeWVhcjogeS52YWx1ZSwgcXVhcnRlcjogcS52YWx1ZSwgLi4uU3RhdGUuZ2V0RmlsdGVycygpIH0pOwogICAgICAgIHJlbmRlclBlcmlvZFJlcG9ydChyZXBvcnQsIHRvZ2dsZS5nZXQoKSk7CiAgICAgIH0pKTsKICAgIH0gZWxzZSBpZiAobW9kZSA9PT0gJ3l0ZCcpIHsKICAgICAgY29uc3QgeSA9IHllYXJTZWxlY3QodGhpc1llYXIpOwogICAgICByb3cuYXBwZW5kKGxhYmVsZWQoJ1llYXInLCB5KSwgcnVuQnRuKGFzeW5jICgpID0+IHsKICAgICAgICBjb25zdCByZXBvcnQgPSBhd2FpdCBBcGkueXRkKHsgeWVhcjogeS52YWx1ZSwgLi4uU3RhdGUuZ2V0RmlsdGVycygpIH0pOwogICAgICAgIHJlbmRlclBlcmlvZFJlcG9ydChyZXBvcnQsICd2c0xhc3RZZWFyJyk7CiAgICAgIH0pKTsKICAgIH0KCiAgICBjb250cm9scy5hcHBlbmRDaGlsZChyb3cpOwogIH0KCiAgZnVuY3Rpb24gbGFiZWxlZChsYWJlbCwgZWwpIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHdyYXAuY2xhc3NOYW1lID0gJ2ZpZWxkLWlubGluZSc7CiAgICB3cmFwLmFwcGVuZCh0ZXh0RWwoJ2xhYmVsJywgbGFiZWwpLCBlbCk7CiAgICByZXR1cm4gd3JhcDsKICB9CiAgZnVuY3Rpb24gcnVuQnRuKG9uQ2xpY2spIHsKICAgIGNvbnN0IGJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgYnRuLmNsYXNzTmFtZSA9ICdidG4gcHJpbWFyeSc7CiAgICBidG4udGV4dENvbnRlbnQgPSAnQ29tcGFyZSc7CiAgICBidG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBvbkNsaWNrKCkpOwogICAgcmV0dXJuIGJ0bjsKICB9CiAgZnVuY3Rpb24geWVhclNlbGVjdChkZWZhdWx0WWVhcikgewogICAgY29uc3Qgc2VsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc2VsZWN0Jyk7CiAgICBmb3IgKGxldCB5ID0gZGVmYXVsdFllYXIgLSAzOyB5IDw9IGRlZmF1bHRZZWFyICsgMTsgeSArPSAxKSB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOyBvcHQudmFsdWUgPSB5OyBvcHQudGV4dENvbnRlbnQgPSB5OwogICAgICBpZiAoeSA9PT0gZGVmYXVsdFllYXIpIG9wdC5zZWxlY3RlZCA9IHRydWU7CiAgICAgIHNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfQogICAgcmV0dXJuIHNlbDsKICB9CiAgZnVuY3Rpb24gbW9udGhTZWxlY3QoZGVmYXVsdE1vbnRoKSB7CiAgICBjb25zdCBuYW1lcyA9IFsnSmFudWFyeScsJ0ZlYnJ1YXJ5JywnTWFyY2gnLCdBcHJpbCcsJ01heScsJ0p1bmUnLCdKdWx5JywnQXVndXN0JywnU2VwdGVtYmVyJywnT2N0b2JlcicsJ05vdmVtYmVyJywnRGVjZW1iZXInXTsKICAgIGNvbnN0IHNlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NlbGVjdCcpOwogICAgbmFtZXMuZm9yRWFjaCgobiwgaSkgPT4gewogICAgICBjb25zdCBvcHQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsgb3B0LnZhbHVlID0gaSArIDE7IG9wdC50ZXh0Q29udGVudCA9IG47CiAgICAgIGlmIChpICsgMSA9PT0gZGVmYXVsdE1vbnRoKSBvcHQuc2VsZWN0ZWQgPSB0cnVlOwogICAgICBzZWwuYXBwZW5kQ2hpbGQob3B0KTsKICAgIH0pOwogICAgcmV0dXJuIHNlbDsKICB9CiAgZnVuY3Rpb24gcXVhcnRlclNlbGVjdCgpIHsKICAgIGNvbnN0IGN1cnJlbnRRID0gTWF0aC5mbG9vcihuZXcgRGF0ZSgpLmdldE1vbnRoKCkgLyAzKSArIDE7CiAgICBjb25zdCBzZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzZWxlY3QnKTsKICAgIFsxLCAyLCAzLCA0XS5mb3JFYWNoKChxKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOyBvcHQudmFsdWUgPSBxOyBvcHQudGV4dENvbnRlbnQgPSBgUSR7cX1gOwogICAgICBpZiAocSA9PT0gY3VycmVudFEpIG9wdC5zZWxlY3RlZCA9IHRydWU7CiAgICAgIHNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgICByZXR1cm4gc2VsOwogIH0KICBmdW5jdGlvbiBwZXJpb2RUb2dnbGUoKSB7CiAgICBjb25zdCBzZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzZWxlY3QnKTsKICAgIFtbJ3ZzUHJldmlvdXNQZXJpb2QnLCAndnMgUHJldmlvdXMgUGVyaW9kJ10sIFsndnNMYXN0WWVhcicsICd2cyBTYW1lIFBlcmlvZCBMYXN0IFllYXInXV0uZm9yRWFjaCgoW3YsIGxdKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOyBvcHQudmFsdWUgPSB2OyBvcHQudGV4dENvbnRlbnQgPSBsOwogICAgICBzZWwuYXBwZW5kQ2hpbGQob3B0KTsKICAgIH0pOwogICAgcmV0dXJuIHsgZWw6IGxhYmVsZWQoJ0NvbXBhcmUnLCBzZWwpLCBnZXQ6ICgpID0+IHNlbC52YWx1ZSB9OwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gcnVuQ29tcGFyZShyYW5nZUEsIHJhbmdlQiwgbGFiZWxBLCBsYWJlbEIpIHsKICAgIGNvbnN0IGZpbHRlcnMgPSBTdGF0ZS5nZXRGaWx0ZXJzKCk7CiAgICBjb25zdCByZXN1bHQgPSBhd2FpdCBBcGkuY29tcGFyZSh7CiAgICAgIGZyb21BOiByYW5nZUEuZnJvbSwgdG9BOiByYW5nZUEudG8sIGZyb21COiByYW5nZUIuZnJvbSwgdG9COiByYW5nZUIudG8sCiAgICAgIHBsYXRmb3JtOiBmaWx0ZXJzLnBsYXRmb3JtLCBjYW1wYWlnblR5cGU6IGZpbHRlcnMuY2FtcGFpZ25UeXBlLCBjb250ZW50VHlwZTogZmlsdGVycy5jb250ZW50VHlwZSwKICAgIH0pOwogICAgcmVuZGVyQ29tcGFyZVJlc3VsdChyZXN1bHQsIGxhYmVsQSwgbGFiZWxCKTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlclBlcmlvZFJlcG9ydChyZXBvcnQsIHdoaWNoKSB7CiAgICBjb25zdCBjbXAgPSByZXBvcnRbd2hpY2hdOwogICAgY29uc3QgbGFiZWxBID0gJ0N1cnJlbnQgcGVyaW9kJzsKICAgIGNvbnN0IGxhYmVsQiA9IHdoaWNoID09PSAndnNMYXN0WWVhcicgPyAnU2FtZSBwZXJpb2QgbGFzdCB5ZWFyJyA6ICdQcmV2aW91cyBwZXJpb2QnOwogICAgcmVuZGVyQ29tcGFyZVJlc3VsdChjbXAsIGxhYmVsQSwgbGFiZWxCLCByZXBvcnQucmFuZ2UpOwogIH0KCiAgZnVuY3Rpb24gc3RhdFRpbGUobGFiZWwsIGN1cnJlbnQsIHByZXZpb3VzLCBncm93dGgsIGlzRHVyYXRpb24pIHsKICAgIGNvbnN0IHRpbGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRpbGUuY2xhc3NOYW1lID0gJ3N0YXQtdGlsZSc7CiAgICBjb25zdCBjdXJEaXNwbGF5ID0gaXNEdXJhdGlvbiA/IEZvcm1hdC5kdXJhdGlvbihjdXJyZW50KSA6IEZvcm1hdC5jb21wYWN0KGN1cnJlbnQpOwogICAgY29uc3QgcHJldkRpc3BsYXkgPSBpc0R1cmF0aW9uID8gRm9ybWF0LmR1cmF0aW9uKHByZXZpb3VzKSA6IEZvcm1hdC5jb21wYWN0KHByZXZpb3VzKTsKICAgIHRpbGUuYXBwZW5kKAogICAgICB0ZXh0RWwoJ2RpdicsIGxhYmVsLCAnc3RhdC1sYWJlbCcpLAogICAgICB0ZXh0RWwoJ2RpdicsIGN1ckRpc3BsYXksICdzdGF0LXZhbHVlJyksCiAgICAgIHRleHRFbCgnZGl2JywgYCR7Rm9ybWF0LnBjdChncm93dGgpfSDCtyB3YXMgJHtwcmV2RGlzcGxheX1gLCBgc3RhdC1kZWx0YSAke0Zvcm1hdC5kZWx0YUNsYXNzKGdyb3d0aCl9YCkKICAgICk7CiAgICByZXR1cm4gdGlsZTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlckNvbXBhcmVSZXN1bHQocmVzdWx0LCBsYWJlbEEsIGxhYmVsQiwgaGVhZGxpbmUpIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnY29tcGFyaXNvblJlc3VsdHMnKTsKICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CgogICAgY29uc3QgdGl0bGUgPSB0ZXh0RWwoJ2RpdicsIGhlYWRsaW5lCiAgICAgID8gYCR7Rm9ybWF0LmRhdGUocmVzdWx0LnJhbmdlQS5mcm9tKX0g4oCTICR7Rm9ybWF0LmRhdGUocmVzdWx0LnJhbmdlQS50byl9YAogICAgICA6IGAke2xhYmVsQX06ICR7Rm9ybWF0LmRhdGUocmVzdWx0LnJhbmdlQS5mcm9tKX0g4oCTICR7Rm9ybWF0LmRhdGUocmVzdWx0LnJhbmdlQS50byl9ICB2cyAgJHtsYWJlbEJ9OiAke0Zvcm1hdC5kYXRlKHJlc3VsdC5yYW5nZUIuZnJvbSl9IOKAkyAke0Zvcm1hdC5kYXRlKHJlc3VsdC5yYW5nZUIudG8pfWAsCiAgICAgICdzZWN0aW9uLXRpdGxlJyk7CiAgICB3cmFwLmFwcGVuZENoaWxkKHRpdGxlKTsKCiAgICBjb25zdCBncmlkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBncmlkLmNsYXNzTmFtZSA9ICdzdGF0LWdyaWQnOwogICAgZ3JpZC5hcHBlbmRDaGlsZChzdGF0VGlsZSgnUG9zdHMnLCByZXN1bHQucmFuZ2VBLnRvdGFscy5wb3N0X2NvdW50LCByZXN1bHQucmFuZ2VCLnRvdGFscy5wb3N0X2NvdW50LCByZXN1bHQuZ3Jvd3RoLnBvc3RfY291bnQsIGZhbHNlKSk7CiAgICBNRVRSSUNfUk9XUy5mb3JFYWNoKChtKSA9PiB7CiAgICAgIGdyaWQuYXBwZW5kQ2hpbGQoc3RhdFRpbGUobS5sYWJlbCwgcmVzdWx0LnJhbmdlQS50b3RhbHNbbS5rZXldLCByZXN1bHQucmFuZ2VCLnRvdGFsc1ttLmtleV0sIHJlc3VsdC5ncm93dGhbbS5rZXldLCBtLmtleSA9PT0gJ3dhdGNoX3RpbWVfc2Vjb25kcycpKTsKICAgIH0pOwogICAgd3JhcC5hcHBlbmRDaGlsZChncmlkKTsKCiAgICBjb25zdCBjaGFydFRpdGxlID0gdGV4dEVsKCdkaXYnLCBgJHtsYWJlbEF9IHZzICR7bGFiZWxCfSBieSBwbGF0Zm9ybWAsICdzZWN0aW9uLXRpdGxlJyk7CiAgICBjb25zdCBjYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjYXJkLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGNvbnN0IGhlYWRlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgaGVhZGVyLmNsYXNzTmFtZSA9ICdjYXJkLWhlYWRlcic7CiAgICBoZWFkZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMycsICdQbGF0Zm9ybSBjb21wYXJpc29uJykpOwogICAgY29uc3QgbWV0cmljU2VsZWN0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc2VsZWN0Jyk7CiAgICBNRVRSSUNfUk9XUy5mb3JFYWNoKChtKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOyBvcHQudmFsdWUgPSBtLmtleTsgb3B0LnRleHRDb250ZW50ID0gbS5sYWJlbDsKICAgICAgaWYgKG0ua2V5ID09PSBjaGFydE1ldHJpYykgb3B0LnNlbGVjdGVkID0gdHJ1ZTsKICAgICAgbWV0cmljU2VsZWN0LmFwcGVuZENoaWxkKG9wdCk7CiAgICB9KTsKICAgIG1ldHJpY1NlbGVjdC5hZGRFdmVudExpc3RlbmVyKCdjaGFuZ2UnLCAoKSA9PiB7CiAgICAgIGNoYXJ0TWV0cmljID0gbWV0cmljU2VsZWN0LnZhbHVlOwogICAgICBkcmF3Q29tcGFyaXNvbkNoYXJ0KHJlc3VsdCwgbGFiZWxBLCBsYWJlbEIpOwogICAgfSk7CiAgICBoZWFkZXIuYXBwZW5kQ2hpbGQobWV0cmljU2VsZWN0KTsKICAgIGNvbnN0IGNoYXJ0V3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY2hhcnRXcmFwLmNsYXNzTmFtZSA9ICdjaGFydC13cmFwIHRhbGwnOwogICAgY2hhcnRXcmFwLmlubmVySFRNTCA9ICc8Y2FudmFzIGlkPSJjb21wYXJpc29uQ2FudmFzIj48L2NhbnZhcz4nOwogICAgY2FyZC5hcHBlbmQoaGVhZGVyLCBjaGFydFdyYXApOwogICAgd3JhcC5hcHBlbmQoY2hhcnRUaXRsZSwgY2FyZCk7CgogICAgZHJhd0NvbXBhcmlzb25DaGFydChyZXN1bHQsIGxhYmVsQSwgbGFiZWxCKTsKICB9CgogIGZ1bmN0aW9uIGRyYXdDb21wYXJpc29uQ2hhcnQocmVzdWx0LCBsYWJlbEEsIGxhYmVsQikgewogICAgY29uc3QgcGxhdGZvcm1JZHMgPSBuZXcgU2V0KFsuLi5yZXN1bHQucmFuZ2VBLnBsYXRmb3JtcywgLi4ucmVzdWx0LnJhbmdlQi5wbGF0Zm9ybXNdLm1hcCgocCkgPT4gcC5wbGF0Zm9ybSkpOwogICAgY29uc3QgcGxhdGZvcm1PcHRpb25zID0gKHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZSB8fCB7IHBsYXRmb3JtczogW10gfSkucGxhdGZvcm1zOwogICAgY29uc3QgbGFiZWxzID0gWy4uLnBsYXRmb3JtSWRzXS5tYXAoKGlkKSA9PiAocGxhdGZvcm1PcHRpb25zLmZpbmQoKHApID0+IHAuaWQgPT09IGlkKSB8fCB7IGxhYmVsOiBpZCB9KS5sYWJlbCk7CiAgICBjb25zdCBieUlkQSA9IE9iamVjdC5mcm9tRW50cmllcyhyZXN1bHQucmFuZ2VBLnBsYXRmb3Jtcy5tYXAoKHApID0+IFtwLnBsYXRmb3JtLCBwXSkpOwogICAgY29uc3QgYnlJZEIgPSBPYmplY3QuZnJvbUVudHJpZXMocmVzdWx0LnJhbmdlQi5wbGF0Zm9ybXMubWFwKChwKSA9PiBbcC5wbGF0Zm9ybSwgcF0pKTsKICAgIGNvbnN0IGN1cnJlbnREYXRhID0gWy4uLnBsYXRmb3JtSWRzXS5tYXAoKGlkKSA9PiAoYnlJZEFbaWRdIHx8IHt9KVtjaGFydE1ldHJpY10gfHwgMCk7CiAgICBjb25zdCBwcmV2aW91c0RhdGEgPSBbLi4ucGxhdGZvcm1JZHNdLm1hcCgoaWQpID0+IChieUlkQltpZF0gfHwge30pW2NoYXJ0TWV0cmljXSB8fCAwKTsKICAgIENoYXJ0cy5jb21wYXJpc29uQmFyQ2hhcnQoJ2NvbXBhcmlzb25DYW52YXMnLCB7IGxhYmVscywgY3VycmVudERhdGEsIHByZXZpb3VzRGF0YSwgY3VycmVudExhYmVsOiBsYWJlbEEsIHByZXZpb3VzTGFiZWw6IGxhYmVsQiB9KTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlcigpIHsKICAgIHJvb3QgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndmlldy1jb21wYXJpc29uJyk7CiAgICBzaGVsbCgpOwogIH0KCiAgcmV0dXJuIHsgcmVuZGVyIH07Cn0pKCk7CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgVXBsb2FkIHRhYjogZHJhZy1kcm9wLCB2YWxpZGF0aW9uIHByZXZpZXcsIHBlci13ZWVrIGNvbmZsaWN0CiAgIHJlc29sdXRpb24sIGNvbW1pdCDigJQgcGx1cyB0aGUgVXBsb2FkIEhpc3RvcnkgdGFiLgogICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KY29uc3QgVXBsb2FkID0gKCgpID0+IHsKICBsZXQgcm9vdDsKICBsZXQgY3VycmVudFByZXZpZXcgPSBudWxsOyAvLyB7IGZpbGVQYXRoLCBvcmlnaW5hbE5hbWUsIGR1cGxpY2F0ZXMsIGlzc3Vlcywgc2FtcGxlLCAuLi4gfQogIGNvbnN0IGR1cGxpY2F0ZUFjdGlvbk92ZXJyaWRlcyA9IHt9OwoKICBmdW5jdGlvbiBzaGVsbCgpIHsKICAgIHJvb3QuaW5uZXJIVE1MID0gJyc7CgogICAgY29uc3QgaW50cm8gPSB0ZXh0RWwoJ2RpdicsICdVcGxvYWQgYSB3ZWVrbHkgZXhwb3J0JywgJ3NlY3Rpb24tdGl0bGUnKTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQoaW50cm8pOwoKICAgIGNvbnN0IGRyb3B6b25lID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBkcm9wem9uZS5jbGFzc05hbWUgPSAnZHJvcHpvbmUnOwogICAgZHJvcHpvbmUuaWQgPSAnZHJvcHpvbmUnOwogICAgZHJvcHpvbmUuaW5uZXJIVE1MID0gYAogICAgICA8ZGl2IGNsYXNzPSJlbXB0eS1pY29uIiBzdHlsZT0ibWFyZ2luOiAwIGF1dG8gMTRweDsiPjxpIGRhdGEtbHVjaWRlPSJ1cGxvYWQtY2xvdWQiIHN0eWxlPSJ3aWR0aDoyMnB4O2hlaWdodDoyMnB4OyI+PC9pPjwvZGl2PgogICAgICA8aDM+RHJhZyAmYW1wOyBkcm9wIHlvdXIgLmNzdiBvciAueGxzeCBmaWxlIGhlcmU8L2gzPgogICAgICA8cD5vciBjbGljayB0byBicm93c2Ug4oCUIGZpbGVzIGFyZSB2YWxpZGF0ZWQgYmVmb3JlIGFueXRoaW5nIGlzIHNhdmVkPC9wPgogICAgICA8aW5wdXQgdHlwZT0iZmlsZSIgaWQ9ImZpbGVJbnB1dCIgYWNjZXB0PSIuY3N2LC54bHN4LC54bHMiIC8+CiAgICBgOwogICAgcm9vdC5hcHBlbmRDaGlsZChkcm9wem9uZSk7CgogICAgY29uc3QgcHJldmlld0FyZWEgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHByZXZpZXdBcmVhLmlkID0gJ3ByZXZpZXdBcmVhJzsKICAgIHJvb3QuYXBwZW5kQ2hpbGQocHJldmlld0FyZWEpOwoKICAgIHdpcmVEcm9wem9uZShkcm9wem9uZSk7CiAgfQoKICBmdW5jdGlvbiB3aXJlRHJvcHpvbmUoZHJvcHpvbmUpIHsKICAgIGNvbnN0IGlucHV0ID0gZHJvcHpvbmUucXVlcnlTZWxlY3RvcignI2ZpbGVJbnB1dCcpOwogICAgZHJvcHpvbmUuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBpbnB1dC5jbGljaygpKTsKICAgIGlucHV0LmFkZEV2ZW50TGlzdGVuZXIoJ2NoYW5nZScsICgpID0+IHsKICAgICAgaWYgKGlucHV0LmZpbGVzWzBdKSBoYW5kbGVGaWxlKGlucHV0LmZpbGVzWzBdKTsKICAgIH0pOwogICAgWydkcmFnZW50ZXInLCAnZHJhZ292ZXInXS5mb3JFYWNoKChldnQpID0+CiAgICAgIGRyb3B6b25lLmFkZEV2ZW50TGlzdGVuZXIoZXZ0LCAoZSkgPT4geyBlLnByZXZlbnREZWZhdWx0KCk7IGRyb3B6b25lLmNsYXNzTGlzdC5hZGQoJ2lzLWRyYWcnKTsgfSkKICAgICk7CiAgICBbJ2RyYWdsZWF2ZScsICdkcm9wJ10uZm9yRWFjaCgoZXZ0KSA9PgogICAgICBkcm9wem9uZS5hZGRFdmVudExpc3RlbmVyKGV2dCwgKGUpID0+IHsgZS5wcmV2ZW50RGVmYXVsdCgpOyBkcm9wem9uZS5jbGFzc0xpc3QucmVtb3ZlKCdpcy1kcmFnJyk7IH0pCiAgICApOwogICAgZHJvcHpvbmUuYWRkRXZlbnRMaXN0ZW5lcignZHJvcCcsIChlKSA9PiB7CiAgICAgIGNvbnN0IGZpbGUgPSBlLmRhdGFUcmFuc2Zlci5maWxlc1swXTsKICAgICAgaWYgKGZpbGUpIGhhbmRsZUZpbGUoZmlsZSk7CiAgICB9KTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIGhhbmRsZUZpbGUoZmlsZSkgewogICAgY29uc3QgYXJlYSA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdwcmV2aWV3QXJlYScpOwogICAgYXJlYS5pbm5lckhUTUwgPSAnJzsKICAgIGFyZWEuYXBwZW5kQ2hpbGQocm93V2l0aFNwaW5uZXIoJ1ZhbGlkYXRpbmcgZmlsZeKApicpKTsKICAgIE9iamVjdC5rZXlzKGR1cGxpY2F0ZUFjdGlvbk92ZXJyaWRlcykuZm9yRWFjaCgoaykgPT4gZGVsZXRlIGR1cGxpY2F0ZUFjdGlvbk92ZXJyaWRlc1trXSk7CiAgICB0cnkgewogICAgICBjdXJyZW50UHJldmlldyA9IGF3YWl0IEFwaS5wcmV2aWV3VXBsb2FkKGZpbGUpOwogICAgICByZW5kZXJQcmV2aWV3KCk7CiAgICB9IGNhdGNoIChlcnIpIHsKICAgICAgYXJlYS5pbm5lckhUTUwgPSAnJzsKICAgICAgYXJlYS5hcHBlbmRDaGlsZChlcnJvckJhbm5lcihlcnIubWVzc2FnZSkpOwogICAgfQogIH0KCiAgZnVuY3Rpb24gcm93V2l0aFNwaW5uZXIodGV4dCkgewogICAgY29uc3QgZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGVsLmNsYXNzTmFtZSA9ICdsb2FkaW5nLXJvdyc7CiAgICBjb25zdCBzcGlubmVyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgc3Bpbm5lci5jbGFzc05hbWUgPSAnc3Bpbm5lcic7CiAgICBlbC5hcHBlbmQoc3Bpbm5lciwgZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUoYCAke3RleHR9YCkpOwogICAgcmV0dXJuIGVsOwogIH0KICBmdW5jdGlvbiBlcnJvckJhbm5lcihtZXNzYWdlKSB7CiAgICBjb25zdCBlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZWwuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgZWwuc3R5bGUuYm9yZGVyTGVmdCA9ICczcHggc29saWQgdmFyKC0tc3RhdHVzLWNyaXRpY2FsKSc7CiAgICBlbC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIGBDb3VsZCBub3QgcmVhZCB0aGlzIGZpbGU6ICR7bWVzc2FnZX1gLCAnbXV0ZWQnKSk7CiAgICByZXR1cm4gZWw7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJQcmV2aWV3KCkgewogICAgY29uc3QgYXJlYSA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdwcmV2aWV3QXJlYScpOwogICAgYXJlYS5pbm5lckhUTUwgPSAnJzsKICAgIGNvbnN0IHAgPSBjdXJyZW50UHJldmlldzsKCiAgICBjb25zdCBzdW1tYXJ5VGl0bGUgPSB0ZXh0RWwoJ2RpdicsICdWYWxpZGF0aW9uIHN1bW1hcnknLCAnc2VjdGlvbi10aXRsZScpOwogICAgY29uc3Qgc3VtbWFyeUdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHN1bW1hcnlHcmlkLmNsYXNzTmFtZSA9ICdzdGF0LWdyaWQnOwogICAgc3VtbWFyeUdyaWQuYXBwZW5kKAogICAgICBzdGF0VGlsZSgnRmlsZScsIHAub3JpZ2luYWxOYW1lKSwKICAgICAgc3RhdFRpbGUoJ1NoZWV0cyBmb3VuZCcsIHAuc2hlZXRzLmxlbmd0aCksCiAgICAgIHN0YXRUaWxlKCdUb3RhbCByb3dzIChhbGwgc2hlZXRzKScsIHAudG90YWxEYXRhUm93cyksCiAgICAgIHN0YXRUaWxlKCdOZXcgcmVjb3JkcycsIHAubmV3UmVjb3Jkc0NvdW50KSwKICAgICAgc3RhdFRpbGUoJ0V4YWN0IGR1cGxpY2F0ZXMgZm91bmQnLCBwLmR1cGxpY2F0ZXMubGVuZ3RoKSwKICAgICAgc3RhdFRpbGUoJ0R1cGxpY2F0ZSByb3dzIGluIGZpbGUnLCBwLmR1cGxpY2F0ZVJvd3NJbkZpbGUpLAogICAgICBzdGF0VGlsZSgnUm93cyB3aXRoIGVycm9ycycsIHAuZXJyb3JSb3dzKQogICAgKTsKICAgIGFyZWEuYXBwZW5kKHN1bW1hcnlUaXRsZSwgc3VtbWFyeUdyaWQpOwoKICAgIGlmIChwLnNoZWV0cy5sZW5ndGgpIHsKICAgICAgY29uc3Qgc2hlZXRzVGl0bGUgPSB0ZXh0RWwoJ2RpdicsICdTaGVldCBicmVha2Rvd24nLCAnc2VjdGlvbi10aXRsZScpOwogICAgICBjb25zdCBzaGVldHNUYWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICAgIHNoZWV0c1RhYmxlLmNsYXNzTmFtZSA9ICdkYXRhLXRhYmxlJzsKICAgICAgc2hlZXRzVGFibGUuaW5uZXJIVE1MID0gJzx0aGVhZD48dHI+PHRoPlNoZWV0PC90aD48dGg+TGF5b3V0IGRldGVjdGVkPC90aD48dGggY2xhc3M9Im51bSI+Um93czwvdGg+PHRoIGNsYXNzPSJudW0iPlZhbGlkPC90aD48dGggY2xhc3M9Im51bSI+RXJyb3JzPC90aD48L3RyPjwvdGhlYWQ+JzsKICAgICAgY29uc3Qgc2hlZXRzQm9keSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3Rib2R5Jyk7CiAgICAgIHAuc2hlZXRzLmZvckVhY2goKHMpID0+IHsKICAgICAgICBjb25zdCB0ciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAgICAgY29uc3QgbGF5b3V0TGFiZWwgPSBzLmZvcm1hdCA9PT0gJ2FnZW5kYScgPyAnTFJTIGFnZW5kYSB0cmFja2VyJyA6IHMuZm9ybWF0ID09PSAnc2ltcGxlJyA/ICdTaW1wbGUgcGxhdGZvcm0gdGFibGUnIDogJ05vdCByZWNvZ25pemVkIOKAlCBzYXZlZCBhcyByYXcgZGF0YSBvbmx5JzsKICAgICAgICB0ci5hcHBlbmQoCiAgICAgICAgICB0ZXh0RWwoJ3RkJywgcy5uYW1lKSwKICAgICAgICAgIHRleHRFbCgndGQnLCBsYXlvdXRMYWJlbCksCiAgICAgICAgICB0ZXh0RWwoJ3RkJywgU3RyaW5nKHMudG90YWxSb3dzKSwgJ251bScpLAogICAgICAgICAgdGV4dEVsKCd0ZCcsIFN0cmluZyhzLnZhbGlkUm93cyksICdudW0nKSwKICAgICAgICAgIHRleHRFbCgndGQnLCBTdHJpbmcocy5lcnJvclJvd3MpLCAnbnVtJykKICAgICAgICApOwogICAgICAgIHNoZWV0c0JvZHkuYXBwZW5kQ2hpbGQodHIpOwogICAgICB9KTsKICAgICAgc2hlZXRzVGFibGUuYXBwZW5kQ2hpbGQoc2hlZXRzQm9keSk7CiAgICAgIGNvbnN0IHNoZWV0c1dyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgc2hlZXRzV3JhcC5jbGFzc05hbWUgPSAndGFibGUtc2Nyb2xsJzsKICAgICAgc2hlZXRzV3JhcC5hcHBlbmRDaGlsZChzaGVldHNUYWJsZSk7CiAgICAgIGFyZWEuYXBwZW5kKHNoZWV0c1RpdGxlLCBzaGVldHNXcmFwKTsKICAgIH0KCiAgICBpZiAocC5kdXBsaWNhdGVzLmxlbmd0aCkgewogICAgICBjb25zdCBkdXBUaXRsZSA9IHRleHRFbCgnZGl2JywgYEV4YWN0IGR1cGxpY2F0ZXMgZm91bmQgKCR7cC5kdXBsaWNhdGVzLmxlbmd0aH0pYCwgJ3NlY3Rpb24tdGl0bGUnKTsKICAgICAgYXJlYS5hcHBlbmRDaGlsZChkdXBUaXRsZSk7CiAgICAgIGFyZWEuYXBwZW5kQ2hpbGQodGV4dEVsKAogICAgICAgICdkaXYnLAogICAgICAgICdFYWNoIG9mIHRoZXNlIHJvd3MgaXMgYnl0ZS1mb3ItYnl0ZSBpZGVudGljYWwgdG8gYW4gYWxyZWFkeS1zYXZlZCByZWNvcmQg4oCUIGV2ZXJ5IGZpZWxkIG1hdGNoZXMsIGluY2x1ZGluZyBldmVyeSBtZXRyaWMsIG5vdCBqdXN0IHRoZSBkYXRlL2NhcHRpb24vcGxhdGZvcm0uIENob29zZSB3aGF0IHRvIGRvIHdpdGggZWFjaCDigJQgb3Igc2V0IGEgZGVmYXVsdCBmb3IgYWxsIG9mIHRoZW0uIChBIHJvdyB0aGF0IHNoYXJlcyB0aGUgc2FtZSBkYXRlL2NhcHRpb24vcGxhdGZvcm0gYnV0IGhhcyBkaWZmZXJlbnQgbnVtYmVycyBpcyBub3Qgc2hvd24gaGVyZSDigJQgaXTigJlzIGltcG9ydGVkIGF1dG9tYXRpY2FsbHkgYXMgaXRzIG93biBuZXcgcmVjb3JkLCBzaW5jZSBpdHMgYW5hbHl0aWNzIGNoYW5nZWQuKScsCiAgICAgICAgJ211dGVkJwogICAgICApKTsKICAgICAgY29uc3QgZGVmYXVsdFJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBkZWZhdWx0Um93LmNsYXNzTmFtZSA9ICdmaWVsZC1pbmxpbmUnOwogICAgICBkZWZhdWx0Um93LnN0eWxlLm1hcmdpbiA9ICcxMHB4IDAnOwogICAgICBjb25zdCBkZWZhdWx0U2VsZWN0ID0gYWN0aW9uU2VsZWN0KCdza2lwJyk7CiAgICAgIGRlZmF1bHRTZWxlY3QuaWQgPSAnZGVmYXVsdER1cGxpY2F0ZUFjdGlvblNlbGVjdCc7CiAgICAgIGRlZmF1bHRTZWxlY3QuYWRkRXZlbnRMaXN0ZW5lcignY2hhbmdlJywgKCkgPT4gewogICAgICAgIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJy5jb25mbGljdC1yb3cgc2VsZWN0W2RhdGEtaGFzaF0nKS5mb3JFYWNoKChzZWwpID0+IHsKICAgICAgICAgIGlmICghZHVwbGljYXRlQWN0aW9uT3ZlcnJpZGVzW3NlbC5kYXRhc2V0Lmhhc2hdKSBzZWwudmFsdWUgPSBkZWZhdWx0U2VsZWN0LnZhbHVlOwogICAgICAgIH0pOwogICAgICB9KTsKICAgICAgZGVmYXVsdFJvdy5hcHBlbmQodGV4dEVsKCdsYWJlbCcsICdEZWZhdWx0IGFjdGlvbiBmb3IgYWxsIG1hdGNoZXMnKSwgZGVmYXVsdFNlbGVjdCk7CiAgICAgIGFyZWEuYXBwZW5kQ2hpbGQoZGVmYXVsdFJvdyk7CgogICAgICBjb25zdCBsaXN0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIGxpc3QuY2xhc3NOYW1lID0gJ2NvbmZsaWN0LWxpc3QnOwogICAgICBwLmR1cGxpY2F0ZXMuZm9yRWFjaCgoZCkgPT4gewogICAgICAgIGNvbnN0IHJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICAgIHJvdy5jbGFzc05hbWUgPSAnY29uZmxpY3Qtcm93JzsKICAgICAgICBjb25zdCBsZWZ0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgICAgbGVmdC5hcHBlbmQoCiAgICAgICAgICB0ZXh0RWwoJ2RpdicsIGAke0Zvcm1hdC5kYXRlKGQucHVibGlzaERhdGUpfSDigJQgJHsoZC5jYXB0aW9uIHx8ICcobm8gY2FwdGlvbiknKS5zbGljZSgwLCA3MCl9YCwgJ3dlZWstbGFiZWwnKSwKICAgICAgICAgIHRleHRFbCgnZGl2JywgYEV4YWN0IG1hdGNoIG9mIGV4aXN0aW5nIHJlY29yZCAjJHtkLmV4aXN0aW5nLnBvc3RJZH0gKGxhc3QgdXBkYXRlZCAke2QuZXhpc3RpbmcudXBkYXRlZEF0fSlgLCAnd2Vlay1tZXRhJykKICAgICAgICApOwogICAgICAgIHJvdy5hcHBlbmRDaGlsZChsZWZ0KTsKICAgICAgICBjb25zdCBzZWwgPSBhY3Rpb25TZWxlY3QoJ3NraXAnKTsKICAgICAgICBzZWwuZGF0YXNldC5oYXNoID0gZC5oYXNoOwogICAgICAgIHNlbC5hZGRFdmVudExpc3RlbmVyKCdjaGFuZ2UnLCAoKSA9PiB7IGR1cGxpY2F0ZUFjdGlvbk92ZXJyaWRlc1tkLmhhc2hdID0gc2VsLnZhbHVlOyB9KTsKICAgICAgICByb3cuYXBwZW5kQ2hpbGQoc2VsKTsKICAgICAgICBsaXN0LmFwcGVuZENoaWxkKHJvdyk7CiAgICAgIH0pOwogICAgICBhcmVhLmFwcGVuZENoaWxkKGxpc3QpOwogICAgfQoKICAgIGNvbnN0IG5vdGVzRmllbGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIG5vdGVzRmllbGQuY2xhc3NOYW1lID0gJ2Zvcm0tZmllbGQnOwogICAgbm90ZXNGaWVsZC5zdHlsZS5tYXJnaW4gPSAnMTJweCAwJzsKICAgIG5vdGVzRmllbGQuYXBwZW5kQ2hpbGQodGV4dEVsKCdsYWJlbCcsICdVcGxvYWQgbm90ZXMgKG9wdGlvbmFsKScpKTsKICAgIGNvbnN0IG5vdGVzSW5wdXQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOwogICAgbm90ZXNJbnB1dC50eXBlID0gJ3RleHQnOwogICAgbm90ZXNJbnB1dC5pZCA9ICd1cGxvYWROb3Rlc0lucHV0JzsKICAgIG5vdGVzSW5wdXQucGxhY2Vob2xkZXIgPSAnZS5nLiAiV2VlayAzIGV4cG9ydCwgaW5jbHVkZXMgY29ycmVjdGVkIFRpa1RvayBudW1iZXJzIic7CiAgICBub3Rlc0ZpZWxkLmFwcGVuZENoaWxkKG5vdGVzSW5wdXQpOwogICAgYXJlYS5hcHBlbmRDaGlsZChub3Rlc0ZpZWxkKTsKCiAgICBpZiAocC5pc3N1ZXMubGVuZ3RoKSB7CiAgICAgIGNvbnN0IGlzc3Vlc1RpdGxlID0gdGV4dEVsKCdkaXYnLCBgUm93cyBza2lwcGVkIG9yIGZsYWdnZWQgKCR7cC5pc3N1ZXMubGVuZ3RofSlgLCAnc2VjdGlvbi10aXRsZScpOwogICAgICBjb25zdCBpc3N1ZXNDYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIGlzc3Vlc0NhcmQuY2xhc3NOYW1lID0gJ2lzc3Vlcy1saXN0JzsKICAgICAgcC5pc3N1ZXMuZm9yRWFjaCgoaXNzdWUpID0+IHsKICAgICAgICBjb25zdCByb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgICByb3cuY2xhc3NOYW1lID0gJ2lzc3VlLXJvdyc7CiAgICAgICAgaWYgKGlzc3VlLnJvd051bWJlcikgcm93LmFwcGVuZENoaWxkKHRleHRFbCgnc3BhbicsIGBSb3cgJHtpc3N1ZS5yb3dOdW1iZXJ9YCwgJ3Jvdy1ubycpKTsKICAgICAgICByb3cuYXBwZW5kQ2hpbGQoZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUoaXNzdWUubWVzc2FnZSkpOwogICAgICAgIGlzc3Vlc0NhcmQuYXBwZW5kQ2hpbGQocm93KTsKICAgICAgfSk7CiAgICAgIGFyZWEuYXBwZW5kKGlzc3Vlc1RpdGxlLCBpc3N1ZXNDYXJkKTsKICAgIH0KCiAgICBpZiAocC5zYW1wbGUubGVuZ3RoKSB7CiAgICAgIGNvbnN0IHNhbXBsZVRpdGxlID0gdGV4dEVsKCdkaXYnLCAnU2FtcGxlIG9mIHBhcnNlZCByb3dzJywgJ3NlY3Rpb24tdGl0bGUnKTsKICAgICAgY29uc3QgdGFibGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0YWJsZScpOwogICAgICB0YWJsZS5jbGFzc05hbWUgPSAnZGF0YS10YWJsZSc7CiAgICAgIHRhYmxlLmlubmVySFRNTCA9ICc8dGhlYWQ+PHRyPjx0aD5EYXRlPC90aD48dGg+Q2FwdGlvbjwvdGg+PHRoPlR5cGU8L3RoPjx0aD5DYW1wYWlnbjwvdGg+PHRoPlBsYXRmb3JtczwvdGg+PC90cj48L3RoZWFkPic7CiAgICAgIGNvbnN0IHRib2R5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGJvZHknKTsKICAgICAgcC5zYW1wbGUuZm9yRWFjaCgocykgPT4gewogICAgICAgIGNvbnN0IHRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgICB0ci5hcHBlbmQoCiAgICAgICAgICB0ZXh0RWwoJ3RkJywgRm9ybWF0LmRhdGUocy5wdWJsaXNoRGF0ZSkpLAogICAgICAgICAgdGV4dEVsKCd0ZCcsIHMuY2FwdGlvbiB8fCAn4oCUJyksCiAgICAgICAgICB0ZXh0RWwoJ3RkJywgcy5jb250ZW50VHlwZSB8fCAn4oCUJyksCiAgICAgICAgICB0ZXh0RWwoJ3RkJywgcy5jYW1wYWlnblR5cGUgfHwgJ1Vuc3BlY2lmaWVkJyksCiAgICAgICAgICB0ZXh0RWwoJ3RkJywgcy5wbGF0Zm9ybXMuam9pbignLCAnKSkKICAgICAgICApOwogICAgICAgIHRib2R5LmFwcGVuZENoaWxkKHRyKTsKICAgICAgfSk7CiAgICAgIHRhYmxlLmFwcGVuZENoaWxkKHRib2R5KTsKICAgICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICB3cmFwLmNsYXNzTmFtZSA9ICd0YWJsZS1zY3JvbGwnOwogICAgICB3cmFwLmFwcGVuZENoaWxkKHRhYmxlKTsKICAgICAgYXJlYS5hcHBlbmQoc2FtcGxlVGl0bGUsIHdyYXApOwogICAgfQoKICAgIGNvbnN0IGFjdGlvbnMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGFjdGlvbnMuY2xhc3NOYW1lID0gJ2J0bi1yb3cnOwogICAgYWN0aW9ucy5zdHlsZS5tYXJnaW5Ub3AgPSAnMTZweCc7CiAgICBjb25zdCBjb21taXRCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIGNvbW1pdEJ0bi5jbGFzc05hbWUgPSAnYnRuIHByaW1hcnknOwogICAgY29tbWl0QnRuLnRleHRDb250ZW50ID0gcC52YWxpZFJvd3MgPiAwID8gYEltcG9ydCAke3AudmFsaWRSb3dzfSByb3cocylgIDogJ05vdGhpbmcgdG8gaW1wb3J0JzsKICAgIGNvbW1pdEJ0bi5kaXNhYmxlZCA9IHAudmFsaWRSb3dzID09PSAwOwogICAgY29tbWl0QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gY29tbWl0KGNvbW1pdEJ0bikpOwogICAgY29uc3QgY2FuY2VsQnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBjYW5jZWxCdG4uY2xhc3NOYW1lID0gJ2J0bic7CiAgICBjYW5jZWxCdG4udGV4dENvbnRlbnQgPSAnQ2FuY2VsJzsKICAgIGNhbmNlbEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsgY3VycmVudFByZXZpZXcgPSBudWxsOyBzaGVsbCgpOyB9KTsKICAgIGFjdGlvbnMuYXBwZW5kKGNvbW1pdEJ0biwgY2FuY2VsQnRuKTsKICAgIGFyZWEuYXBwZW5kQ2hpbGQoYWN0aW9ucyk7CiAgfQoKICBmdW5jdGlvbiBzdGF0VGlsZShsYWJlbCwgdmFsdWUpIHsKICAgIGNvbnN0IHRpbGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRpbGUuY2xhc3NOYW1lID0gJ3N0YXQtdGlsZSc7CiAgICB0aWxlLmFwcGVuZCh0ZXh0RWwoJ2RpdicsIGxhYmVsLCAnc3RhdC1sYWJlbCcpLCB0ZXh0RWwoJ2RpdicsIFN0cmluZyh2YWx1ZSksICdzdGF0LXZhbHVlJykpOwogICAgcmV0dXJuIHRpbGU7CiAgfQogIGZ1bmN0aW9uIGFjdGlvblNlbGVjdChkZWZhdWx0VmFsKSB7CiAgICBjb25zdCBzZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzZWxlY3QnKTsKICAgIFtbJ3NraXAnLCAnU2tpcCAoa2VlcCBleGlzdGluZyByZWNvcmQgdW5jaGFuZ2VkKSddLCBbJ3VwZGF0ZScsICdVcGRhdGUgZXhpc3RpbmcgcmVjb3JkJ10sIFsnY3JlYXRlJywgJ0NyZWF0ZSBhcyBhIG5ldywgc2VwYXJhdGUgcmVjb3JkJ11dLmZvckVhY2goKFt2LCBsXSkgPT4gewogICAgICBjb25zdCBvcHQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsgb3B0LnZhbHVlID0gdjsgb3B0LnRleHRDb250ZW50ID0gbDsKICAgICAgaWYgKHYgPT09IGRlZmF1bHRWYWwpIG9wdC5zZWxlY3RlZCA9IHRydWU7CiAgICAgIHNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgICByZXR1cm4gc2VsOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gY29tbWl0KGJ0bikgewogICAgYnRuLmRpc2FibGVkID0gdHJ1ZTsKICAgIGJ0bi50ZXh0Q29udGVudCA9ICdJbXBvcnRpbmfigKYnOwogICAgY29uc3QgZGVmYXVsdER1cGxpY2F0ZUFjdGlvbiA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdkZWZhdWx0RHVwbGljYXRlQWN0aW9uU2VsZWN0Jyk/LnZhbHVlIHx8ICdza2lwJzsKICAgIGNvbnN0IG5vdGVzID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3VwbG9hZE5vdGVzSW5wdXQnKT8udmFsdWUgfHwgbnVsbDsKICAgIHRyeSB7CiAgICAgIGNvbnN0IHJlc3VsdCA9IGF3YWl0IEFwaS5jb21taXRVcGxvYWQoewogICAgICAgIGZpbGVQYXRoOiBjdXJyZW50UHJldmlldy5maWxlUGF0aCwKICAgICAgICBvcmlnaW5hbE5hbWU6IGN1cnJlbnRQcmV2aWV3Lm9yaWdpbmFsTmFtZSwKICAgICAgICBkZWZhdWx0RHVwbGljYXRlQWN0aW9uLAogICAgICAgIGR1cGxpY2F0ZUFjdGlvbnM6IGR1cGxpY2F0ZUFjdGlvbk92ZXJyaWRlcywKICAgICAgICBub3RlcywKICAgICAgfSk7CiAgICAgIFRvYXN0LnNob3coCiAgICAgICAgYEltcG9ydGVkOiAke3Jlc3VsdC5pbXBvcnRlZFJvd3N9IG5ldywgJHtyZXN1bHQudXBkYXRlZFJvd3N9IHVwZGF0ZWQsICR7cmVzdWx0LnNraXBwZWRSb3dzfSBza2lwcGVkLmAsCiAgICAgICAgcmVzdWx0LmVycm9yQ291bnQgPiAwID8gJ2Vycm9yJyA6ICdzdWNjZXNzJwogICAgICApOwogICAgICBjdXJyZW50UHJldmlldyA9IG51bGw7CiAgICAgIHNoZWxsKCk7CiAgICAgIHdpbmRvdy5kaXNwYXRjaEV2ZW50KG5ldyBDdXN0b21FdmVudCgnbHJzOmRhdGEtdXBkYXRlZCcpKTsKICAgIH0gY2F0Y2ggKGVycikgewogICAgICBUb2FzdC5zaG93KGVyci5tZXNzYWdlLCAnZXJyb3InKTsKICAgICAgYnRuLmRpc2FibGVkID0gZmFsc2U7CiAgICAgIGJ0bi50ZXh0Q29udGVudCA9ICdSZXRyeSBpbXBvcnQnOwogICAgfQogIH0KCiAgZnVuY3Rpb24gcmVuZGVyKCkgewogICAgcm9vdCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd2aWV3LXVwbG9hZCcpOwogICAgc2hlbGwoKTsKICB9CgogIHJldHVybiB7IHJlbmRlciB9Owp9KSgpOwoKY29uc3QgSGlzdG9yeSA9ICgoKSA9PiB7CiAgbGV0IHJvb3Q7CgogIGZ1bmN0aW9uIGJhZGdlQ2xhc3Moc3RhdHVzKSB7CiAgICBpZiAoc3RhdHVzID09PSAnc3VjY2VzcycpIHJldHVybiAnc3VjY2Vzcyc7CiAgICBpZiAoc3RhdHVzID09PSAncGFydGlhbCcpIHJldHVybiAncGFydGlhbCc7CiAgICByZXR1cm4gJ2ZhaWxlZCc7CiAgfQoKICBmdW5jdGlvbiBidWlsZEJhY2t1cENhcmQoKSB7CiAgICBjb25zdCBjYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjYXJkLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGNhcmQuc3R5bGUubWFyZ2luQm90dG9tID0gJzIwcHgnOwogICAgY29uc3QgaGVhZGVyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBoZWFkZXIuY2xhc3NOYW1lID0gJ2NhcmQtaGVhZGVyJzsKICAgIGhlYWRlci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2gzJywgJ0JhY2t1cCAmIFJlc3RvcmUnKSk7CiAgICBjYXJkLmFwcGVuZENoaWxkKGhlYWRlcik7CiAgICBjYXJkLmFwcGVuZENoaWxkKHRleHRFbCgKICAgICAgJ2RpdicsCiAgICAgICdEb3dubG9hZCBhIGZ1bGwgc25hcHNob3Qgb2YgdGhlIGRhdGFiYXNlIGFueSB0aW1lLiBSZXN0b3JpbmcgcmVwbGFjZXMgQUxMIGN1cnJlbnQgZGF0YSB3aXRoIHRoZSB1cGxvYWRlZCBiYWNrdXAgYW5kIHJlc3RhcnRzIHRoZSBzZXJ2ZXIg4oCUIHRoaXMgY2Fubm90IGJlIHVuZG9uZS4nLAogICAgICAnbXV0ZWQnCiAgICApKTsKCiAgICBjb25zdCBhY3Rpb25zID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBhY3Rpb25zLmNsYXNzTmFtZSA9ICdidG4tcm93JzsKICAgIGFjdGlvbnMuc3R5bGUubWFyZ2luVG9wID0gJzE0cHgnOwoKICAgIGNvbnN0IGRvd25sb2FkQnRuID0gaWNvbkJ0bignYnRuIHByaW1hcnknLCAnZG93bmxvYWQnLCAnRG93bmxvYWQgQmFja3VwJyk7CiAgICBkb3dubG9hZEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsgd2luZG93LmxvY2F0aW9uLmhyZWYgPSAnL2FwaS9iYWNrdXAvZXhwb3J0JzsgfSk7CgogICAgY29uc3QgcmVzdG9yZUlucHV0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsKICAgIHJlc3RvcmVJbnB1dC50eXBlID0gJ2ZpbGUnOwogICAgcmVzdG9yZUlucHV0LmFjY2VwdCA9ICcuZGInOwogICAgcmVzdG9yZUlucHV0LnN0eWxlLmRpc3BsYXkgPSAnbm9uZSc7CgogICAgY29uc3QgcmVzdG9yZUJ0biA9IGljb25CdG4oJ2J0biBkYW5nZXInLCAndXBsb2FkJywgJ1Jlc3RvcmUgZnJvbSBCYWNrdXAnKTsKICAgIHJlc3RvcmVCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiByZXN0b3JlSW5wdXQuY2xpY2soKSk7CgogICAgcmVzdG9yZUlucHV0LmFkZEV2ZW50TGlzdGVuZXIoJ2NoYW5nZScsIGFzeW5jICgpID0+IHsKICAgICAgY29uc3QgZmlsZSA9IHJlc3RvcmVJbnB1dC5maWxlc1swXTsKICAgICAgaWYgKCFmaWxlKSByZXR1cm47CiAgICAgIGNvbnN0IHN1cmUgPSB3aW5kb3cuY29uZmlybSgKICAgICAgICAnUmVzdG9yaW5nIHdpbGwgUkVQTEFDRSBhbGwgY3VycmVudCBkYXRhIHdpdGggdGhpcyBiYWNrdXAgZmlsZSBhbmQgcmVzdGFydCB0aGUgc2VydmVyLiBUaGlzIGNhbm5vdCBiZSB1bmRvbmUuIENvbnRpbnVlPycKICAgICAgKTsKICAgICAgaWYgKCFzdXJlKSB7CiAgICAgICAgcmVzdG9yZUlucHV0LnZhbHVlID0gJyc7CiAgICAgICAgcmV0dXJuOwogICAgICB9CiAgICAgIHJlc3RvcmVCdG4uZGlzYWJsZWQgPSB0cnVlOwogICAgICB0cnkgewogICAgICAgIGNvbnN0IGZvcm0gPSBuZXcgRm9ybURhdGEoKTsKICAgICAgICBmb3JtLmFwcGVuZCgnZmlsZScsIGZpbGUpOwogICAgICAgIGNvbnN0IHJlc3VsdCA9IGF3YWl0IEFwaS5yZXN0b3JlQmFja3VwKGZvcm0pOwogICAgICAgIFRvYXN0LnNob3cocmVzdWx0Lm1lc3NhZ2UgfHwgJ0JhY2t1cCByZXN0b3JlZC4gVGhlIHNlcnZlciBpcyByZXN0YXJ0aW5nLicsICdzdWNjZXNzJyk7CiAgICAgIH0gY2F0Y2ggKGVycikgewogICAgICAgIFRvYXN0LnNob3coZXJyLm1lc3NhZ2UsICdlcnJvcicpOwogICAgICAgIHJlc3RvcmVCdG4uZGlzYWJsZWQgPSBmYWxzZTsKICAgICAgfSBmaW5hbGx5IHsKICAgICAgICByZXN0b3JlSW5wdXQudmFsdWUgPSAnJzsKICAgICAgfQogICAgfSk7CgogICAgYWN0aW9ucy5hcHBlbmQoZG93bmxvYWRCdG4sIHJlc3RvcmVCdG4sIHJlc3RvcmVJbnB1dCk7CiAgICBjYXJkLmFwcGVuZENoaWxkKGFjdGlvbnMpOwogICAgcmV0dXJuIGNhcmQ7CiAgfQoKICBhc3luYyBmdW5jdGlvbiByZW5kZXIoKSB7CiAgICByb290ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3ZpZXctaGlzdG9yeScpOwogICAgcm9vdC5pbm5lckhUTUwgPSAnJzsKICAgIHJvb3QuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCAnVXBsb2FkIGhpc3RvcnknLCAnc2VjdGlvbi10aXRsZScpKTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQoYnVpbGRCYWNrdXBDYXJkKCkpOwoKICAgIGNvbnN0IHVwbG9hZHMgPSBhd2FpdCBBcGkudXBsb2FkSGlzdG9yeSgpOwogICAgaWYgKCF1cGxvYWRzLmxlbmd0aCkgewogICAgICByb290LmFwcGVuZENoaWxkKGVtcHR5U3RhdGUoewogICAgICAgIGljb246ICd1cGxvYWQtY2xvdWQnLAogICAgICAgIHRpdGxlOiAnTm8gdXBsb2FkcyB5ZXQnLAogICAgICAgIG1lc3NhZ2U6ICdJbXBvcnQgeW91ciBmaXJzdCB3ZWVrbHkgZXhwb3J0IHRvIHN0YXJ0IHNlZWluZyBkYXRhIGFjcm9zcyB0aGUgYXBwLicsCiAgICAgICAgYWN0aW9uTGFiZWw6ICdVcGxvYWQgZGF0YScsCiAgICAgICAgb25BY3Rpb246ICgpID0+IGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3IoJy50YWItYnRuW2RhdGEtdGFiPSJ1cGxvYWQiXScpPy5jbGljaygpLAogICAgICB9KSk7CiAgICAgIHJldHVybjsKICAgIH0KCiAgICBjb25zdCB0YWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICB0YWJsZS5jbGFzc05hbWUgPSAnZGF0YS10YWJsZSc7CiAgICB0YWJsZS5pbm5lckhUTUwgPSAnPHRoZWFkPjx0cj48dGg+RmlsZTwvdGg+PHRoPlVwbG9hZGVkPC90aD48dGg+U3RhdHVzPC90aD48dGggY2xhc3M9Im51bSI+SW1wb3J0ZWQ8L3RoPjx0aCBjbGFzcz0ibnVtIj5VcGRhdGVkPC90aD48dGggY2xhc3M9Im51bSI+U2tpcHBlZDwvdGg+PHRoIGNsYXNzPSJudW0iPkVycm9yczwvdGg+PHRoPldlZWtzPC90aD48dGg+Tm90ZXM8L3RoPjwvdHI+PC90aGVhZD4nOwogICAgY29uc3QgdGJvZHkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0Ym9keScpOwogICAgdXBsb2Fkcy5mb3JFYWNoKCh1KSA9PiB7CiAgICAgIGNvbnN0IHRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgdHIuc3R5bGUuY3Vyc29yID0gJ3BvaW50ZXInOwogICAgICBjb25zdCBiYWRnZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgICAgYmFkZ2UuY2xhc3NOYW1lID0gYGJhZGdlICR7YmFkZ2VDbGFzcyh1LnN0YXR1cyl9YDsKICAgICAgYmFkZ2UudGV4dENvbnRlbnQgPSB1LnN0YXR1czsKICAgICAgY29uc3Qgc3RhdHVzVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBzdGF0dXNUZC5hcHBlbmRDaGlsZChiYWRnZSk7CiAgICAgIHRyLmFwcGVuZCgKICAgICAgICB0ZXh0RWwoJ3RkJywgdS5maWxlbmFtZSksCiAgICAgICAgdGV4dEVsKCd0ZCcsIHUudXBsb2FkZWRfYXQpLAogICAgICAgIHN0YXR1c1RkLAogICAgICAgIHRleHRFbCgndGQnLCBTdHJpbmcodS5pbXBvcnRlZF9yb3dzKSwgJ251bScpLAogICAgICAgIHRleHRFbCgndGQnLCBTdHJpbmcodS51cGRhdGVkX3Jvd3MpLCAnbnVtJyksCiAgICAgICAgdGV4dEVsKCd0ZCcsIFN0cmluZyh1LnNraXBwZWRfcm93cyksICdudW0nKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgU3RyaW5nKHUuZXJyb3JfY291bnQpLCAnbnVtJyksCiAgICAgICAgdGV4dEVsKCd0ZCcsIHUud2Vla3NfYWZmZWN0ZWQubWFwKCh3KSA9PiBGb3JtYXQuZGF0ZSh3KSkuam9pbignLCAnKSB8fCAn4oCUJyksCiAgICAgICAgdGV4dEVsKCd0ZCcsIHUubm90ZXMgfHwgJ+KAlCcpCiAgICAgICk7CiAgICAgIHRyLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gdG9nZ2xlRXJyb3JzKHUuaWQsIHRyKSk7CiAgICAgIHRib2R5LmFwcGVuZENoaWxkKHRyKTsKICAgIH0pOwogICAgdGFibGUuYXBwZW5kQ2hpbGQodGJvZHkpOwogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgd3JhcC5jbGFzc05hbWUgPSAndGFibGUtc2Nyb2xsJzsKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGFibGUpOwogICAgcm9vdC5hcHBlbmRDaGlsZCh3cmFwKTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIHRvZ2dsZUVycm9ycyh1cGxvYWRJZCwgdHIpIHsKICAgIGNvbnN0IGV4aXN0aW5nID0gdHIubmV4dEVsZW1lbnRTaWJsaW5nOwogICAgaWYgKGV4aXN0aW5nICYmIGV4aXN0aW5nLmNsYXNzTGlzdC5jb250YWlucygnZXJyb3ItbG9nLXJvdycpKSB7CiAgICAgIGV4aXN0aW5nLnJlbW92ZSgpOwogICAgICByZXR1cm47CiAgICB9CiAgICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcuZXJyb3ItbG9nLXJvdycpLmZvckVhY2goKGVsKSA9PiBlbC5yZW1vdmUoKSk7CiAgICBjb25zdCBlcnJvcnMgPSBhd2FpdCBBcGkudXBsb2FkRXJyb3JzKHVwbG9hZElkKTsKICAgIGNvbnN0IGxvZ1JvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICBsb2dSb3cuY2xhc3NOYW1lID0gJ2Vycm9yLWxvZy1yb3cnOwogICAgY29uc3QgdGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgdGQuY29sU3BhbiA9IDk7CiAgICBpZiAoIWVycm9ycy5sZW5ndGgpIHsKICAgICAgdGQuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCAnTm8gaXNzdWVzIGxvZ2dlZCBmb3IgdGhpcyB1cGxvYWQuJywgJ211dGVkJykpOwogICAgfSBlbHNlIHsKICAgICAgY29uc3QgbGlzdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBsaXN0LmNsYXNzTmFtZSA9ICdpc3N1ZXMtbGlzdCc7CiAgICAgIGVycm9ycy5mb3JFYWNoKChlKSA9PiB7CiAgICAgICAgY29uc3Qgcm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgICAgcm93LmNsYXNzTmFtZSA9ICdpc3N1ZS1yb3cnOwogICAgICAgIGNvbnN0IGJhZGdlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICAgIGJhZGdlLmNsYXNzTmFtZSA9IGBiYWRnZSAke2Uuc2V2ZXJpdHl9LXNldmA7CiAgICAgICAgYmFkZ2UudGV4dENvbnRlbnQgPSBlLnNldmVyaXR5OwogICAgICAgIHJvdy5hcHBlbmQoYmFkZ2UsIGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKGAgJHtlLnJvd19udW1iZXIgPyBgUm93ICR7ZS5yb3dfbnVtYmVyfTogYCA6ICcnfSR7ZS5tZXNzYWdlfWApKTsKICAgICAgICBsaXN0LmFwcGVuZENoaWxkKHJvdyk7CiAgICAgIH0pOwogICAgICB0ZC5hcHBlbmRDaGlsZChsaXN0KTsKICAgIH0KCiAgICBjb25zdCByYXdCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIHJhd0J0bi5jbGFzc05hbWUgPSAnYnRuJzsKICAgIHJhd0J0bi5zdHlsZS5tYXJnaW5Ub3AgPSAnMTBweCc7CiAgICByYXdCdG4udGV4dENvbnRlbnQgPSAnVmlldyBldmVyeSByYXcgc291cmNlIHJvdyBmcm9tIHRoaXMgdXBsb2FkJzsKICAgIHJhd0J0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IGxvYWRSYXdSb3dzKHVwbG9hZElkLCByYXdCdG4pKTsKICAgIHRkLmFwcGVuZENoaWxkKHJhd0J0bik7CiAgICBjb25zdCByYXdXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICByYXdXcmFwLmlkID0gYHJhd1dyYXAtJHt1cGxvYWRJZH1gOwogICAgdGQuYXBwZW5kQ2hpbGQocmF3V3JhcCk7CgogICAgbG9nUm93LmFwcGVuZENoaWxkKHRkKTsKICAgIHRyLmFmdGVyKGxvZ1Jvdyk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiBsb2FkUmF3Um93cyh1cGxvYWRJZCwgYnRuKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoYHJhd1dyYXAtJHt1cGxvYWRJZH1gKTsKICAgIGlmICghd3JhcCkgcmV0dXJuOwogICAgaWYgKHdyYXAuZGF0YXNldC5sb2FkZWQpIHsKICAgICAgd3JhcC5zdHlsZS5kaXNwbGF5ID0gd3JhcC5zdHlsZS5kaXNwbGF5ID09PSAnbm9uZScgPyAnYmxvY2snIDogJ25vbmUnOwogICAgICByZXR1cm47CiAgICB9CiAgICBidG4udGV4dENvbnRlbnQgPSAnTG9hZGluZ+KApic7CiAgICBjb25zdCB7IHJvd3MsIHRvdGFsIH0gPSBhd2FpdCBBcGkudXBsb2FkUmF3Um93cyh1cGxvYWRJZCk7CiAgICB3cmFwLmRhdGFzZXQubG9hZGVkID0gJzEnOwogICAgYnRuLnRleHRDb250ZW50ID0gYFNob3dpbmcgJHtyb3dzLmxlbmd0aH0gb2YgJHt0b3RhbH0gcmF3IHJvdyhzKWA7CgogICAgY29uc3QgYnlTaGVldCA9IG5ldyBNYXAoKTsKICAgIHJvd3MuZm9yRWFjaCgocikgPT4gewogICAgICBpZiAoIWJ5U2hlZXQuaGFzKHIuc2hlZXRfbmFtZSkpIGJ5U2hlZXQuc2V0KHIuc2hlZXRfbmFtZSwgW10pOwogICAgICBieVNoZWV0LmdldChyLnNoZWV0X25hbWUpLnB1c2gocik7CiAgICB9KTsKCiAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgd3JhcC5zdHlsZS5tYXJnaW5Ub3AgPSAnMTBweCc7CiAgICBieVNoZWV0LmZvckVhY2goKHNoZWV0Um93cywgc2hlZXROYW1lKSA9PiB7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBgU2hlZXQ6ICR7c2hlZXROYW1lfSAoJHtzaGVldFJvd3MubGVuZ3RofSByb3cocykpYCwgJ3N0YXQtbGFiZWwnKSk7CiAgICAgIGNvbnN0IGhlYWRlcnMgPSBzaGVldFJvd3NbMF0uaGVhZGVyczsKICAgICAgY29uc3QgdGFibGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0YWJsZScpOwogICAgICB0YWJsZS5jbGFzc05hbWUgPSAnZGF0YS10YWJsZSc7CiAgICAgIGNvbnN0IHRoZWFkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgdGhlYWQuYXBwZW5kKHRleHRFbCgndGgnLCAnUm93ICMnKSwgdGV4dEVsKCd0aCcsICdMaW5rZWQgdG8gcG9zdCcpKTsKICAgICAgY29uc3QgY29sQ291bnQgPSBoZWFkZXJzID8gaGVhZGVycy5sZW5ndGggOiBNYXRoLm1heCguLi5zaGVldFJvd3MubWFwKChyKSA9PiByLnJhdy5sZW5ndGgpKTsKICAgICAgZm9yIChsZXQgaSA9IDA7IGkgPCBjb2xDb3VudDsgaSArPSAxKSB0aGVhZC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3RoJywgaGVhZGVycyAmJiBoZWFkZXJzW2ldID8gU3RyaW5nKGhlYWRlcnNbaV0pIDogYENvbCAke2kgKyAxfWApKTsKICAgICAgY29uc3QgdGhlYWRXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGhlYWQnKTsKICAgICAgdGhlYWRXcmFwLmFwcGVuZENoaWxkKHRoZWFkKTsKICAgICAgY29uc3QgdGJvZHkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0Ym9keScpOwogICAgICBzaGVldFJvd3MuZm9yRWFjaCgocikgPT4gewogICAgICAgIGNvbnN0IHRyMiA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAgICAgdHIyLmFwcGVuZCh0ZXh0RWwoJ3RkJywgU3RyaW5nKHIucm93X251bWJlcikpLCB0ZXh0RWwoJ3RkJywgci5wb3N0X2lkID8gYCMke3IucG9zdF9pZH1gIDogJ+KAlCcpKTsKICAgICAgICBmb3IgKGxldCBpID0gMDsgaSA8IGNvbENvdW50OyBpICs9IDEpIHsKICAgICAgICAgIGNvbnN0IHZhbCA9IHIucmF3W2ldOwogICAgICAgICAgdHIyLmFwcGVuZENoaWxkKHRleHRFbCgndGQnLCB2YWwgPT09IHVuZGVmaW5lZCB8fCB2YWwgPT09IG51bGwgPyAnJyA6IFN0cmluZyh2YWwpLnNsaWNlKDAsIDYwKSkpOwogICAgICAgIH0KICAgICAgICB0Ym9keS5hcHBlbmRDaGlsZCh0cjIpOwogICAgICB9KTsKICAgICAgdGFibGUuYXBwZW5kKHRoZWFkV3JhcCwgdGJvZHkpOwogICAgICBjb25zdCBzY3JvbGxXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIHNjcm9sbFdyYXAuY2xhc3NOYW1lID0gJ3RhYmxlLXNjcm9sbCc7CiAgICAgIHNjcm9sbFdyYXAuc3R5bGUubWFyZ2luQm90dG9tID0gJzE2cHgnOwogICAgICBzY3JvbGxXcmFwLmFwcGVuZENoaWxkKHRhYmxlKTsKICAgICAgd3JhcC5hcHBlbmRDaGlsZChzY3JvbGxXcmFwKTsKICAgIH0pOwogIH0KCiAgcmV0dXJuIHsgcmVuZGVyIH07Cn0pKCk7CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgQXBwIGJvb3RzdHJhcDogdGFiIHJvdXRpbmcsIGZpbHRlciBiYXIgd2lyaW5nLCB0aGVtZSB0b2dnbGUuCiAgID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwooKCkgPT4gewogIGNvbnN0IFZJRVdTID0gewogICAgZGFzaGJvYXJkOiBEYXNoYm9hcmQsCiAgICByZWNvcmRzOiBSZWNvcmRzLAogICAgY29tcGFyaXNvbjogQ29tcGFyaXNvbiwKICAgIHVwbG9hZDogVXBsb2FkLAogICAgaGlzdG9yeTogSGlzdG9yeSwKICB9OwoKICBsZXQgYWN0aXZlVGFiID0gJ2Rhc2hib2FyZCc7CgogIGZ1bmN0aW9uIHN3aXRjaFRhYih0YWIpIHsKICAgIGFjdGl2ZVRhYiA9IHRhYjsKICAgIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJy50YWItYnRuJykuZm9yRWFjaCgoYnRuKSA9PiB7CiAgICAgIGNvbnN0IGlzQWN0aXZlID0gYnRuLmRhdGFzZXQudGFiID09PSB0YWI7CiAgICAgIGJ0bi5jbGFzc0xpc3QudG9nZ2xlKCdpcy1hY3RpdmUnLCBpc0FjdGl2ZSk7CiAgICAgIGJ0bi5zZXRBdHRyaWJ1dGUoJ2FyaWEtc2VsZWN0ZWQnLCBTdHJpbmcoaXNBY3RpdmUpKTsKICAgIH0pOwogICAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnLnZpZXcnKS5mb3JFYWNoKCh2aWV3KSA9PiB7CiAgICAgIHZpZXcuY2xhc3NMaXN0LnRvZ2dsZSgnaXMtYWN0aXZlJywgdmlldy5pZCA9PT0gYHZpZXctJHt0YWJ9YCk7CiAgICB9KTsKICAgIC8vIEZpbHRlcnMgYXBwbHkgdG8gRGFzaGJvYXJkIGFuZCBEYXRhIFJlY29yZHMgKENvbXBhcmlzb25zIGhhcyBpdHMgb3duIHJhbmdlIGNvbnRyb2xzKS4KICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJCYXInKS5zdHlsZS5kaXNwbGF5ID0gKHRhYiA9PT0gJ2Rhc2hib2FyZCcgfHwgdGFiID09PSAncmVjb3JkcycpID8gJ2ZsZXgnIDogJ25vbmUnOwogICAgcmVuZGVyQWN0aXZlVmlldygpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyQWN0aXZlVmlldygpIHsKICAgIGNvbnN0IHZpZXcgPSBWSUVXU1thY3RpdmVUYWJdOwogICAgaWYgKHZpZXcgJiYgdmlldy5yZW5kZXIpIHZpZXcucmVuZGVyKCk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiBsb2FkRmlsdGVyT3B0aW9ucygpIHsKICAgIGNvbnN0IG9wdGlvbnMgPSBhd2FpdCBBcGkuZmlsdGVyT3B0aW9ucygpOwogICAgd2luZG93Ll9fZmlsdGVyT3B0aW9uc0NhY2hlID0gb3B0aW9uczsKCiAgICBjb25zdCBwbGF0Zm9ybVNlbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJQbGF0Zm9ybScpOwogICAgcGxhdGZvcm1TZWwubGVuZ3RoID0gMTsKICAgIG9wdGlvbnMucGxhdGZvcm1zLmZvckVhY2goKHApID0+IHsKICAgICAgY29uc3Qgb3B0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7CiAgICAgIG9wdC52YWx1ZSA9IHAuaWQ7CiAgICAgIG9wdC50ZXh0Q29udGVudCA9IHAubGFiZWw7CiAgICAgIHBsYXRmb3JtU2VsLmFwcGVuZENoaWxkKG9wdCk7CiAgICB9KTsKCiAgICBjb25zdCBjYW1wYWlnblNlbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJDYW1wYWlnbicpOwogICAgY2FtcGFpZ25TZWwubGVuZ3RoID0gMTsKICAgIG9wdGlvbnMuY2FtcGFpZ25UeXBlcy5mb3JFYWNoKChjKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOwogICAgICBvcHQudmFsdWUgPSBjOwogICAgICBvcHQudGV4dENvbnRlbnQgPSBjOwogICAgICBjYW1wYWlnblNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CgogICAgY29uc3QgY29udGVudFNlbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJDb250ZW50VHlwZScpOwogICAgY29udGVudFNlbC5sZW5ndGggPSAxOwogICAgb3B0aW9ucy5jb250ZW50VHlwZXMuZm9yRWFjaCgoYykgPT4gewogICAgICBjb25zdCBvcHQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsKICAgICAgb3B0LnZhbHVlID0gYzsKICAgICAgb3B0LnRleHRDb250ZW50ID0gYzsKICAgICAgY29udGVudFNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgfQoKICBmdW5jdGlvbiB3aXJlRmlsdGVyQmFyKCkgewogICAgY29uc3QgZGF0ZUZyb20gPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmlsdGVyRGF0ZUZyb20nKTsKICAgIGNvbnN0IGRhdGVUbyA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJEYXRlVG8nKTsKICAgIGNvbnN0IHBsYXRmb3JtID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZpbHRlclBsYXRmb3JtJyk7CiAgICBjb25zdCBjYW1wYWlnbiA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJDYW1wYWlnbicpOwogICAgY29uc3QgY29udGVudFR5cGUgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmlsdGVyQ29udGVudFR5cGUnKTsKICAgIGNvbnN0IGYgPSBTdGF0ZS5nZXRGaWx0ZXJzKCk7CiAgICBkYXRlRnJvbS52YWx1ZSA9IGYuZGF0ZUZyb207CiAgICBkYXRlVG8udmFsdWUgPSBmLmRhdGVUbzsKCiAgICBmdW5jdGlvbiBhcHBseSgpIHsKICAgICAgU3RhdGUuc2V0RmlsdGVycyh7CiAgICAgICAgZGF0ZUZyb206IGRhdGVGcm9tLnZhbHVlLAogICAgICAgIGRhdGVUbzogZGF0ZVRvLnZhbHVlLAogICAgICAgIHBsYXRmb3JtOiBwbGF0Zm9ybS52YWx1ZSwKICAgICAgICBjYW1wYWlnblR5cGU6IGNhbXBhaWduLnZhbHVlLAogICAgICAgIGNvbnRlbnRUeXBlOiBjb250ZW50VHlwZS52YWx1ZSwKICAgICAgfSk7CiAgICB9CiAgICBbZGF0ZUZyb20sIGRhdGVUbywgcGxhdGZvcm0sIGNhbXBhaWduLCBjb250ZW50VHlwZV0uZm9yRWFjaCgoZWwpID0+IGVsLmFkZEV2ZW50TGlzdGVuZXIoJ2NoYW5nZScsIGFwcGx5KSk7CgogICAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnI2ZpbHRlclByZXNldHMgYnV0dG9uJykuZm9yRWFjaCgoYnRuKSA9PiB7CiAgICAgIGJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsKICAgICAgICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcjZmlsdGVyUHJlc2V0cyBidXR0b24nKS5mb3JFYWNoKChiKSA9PiBiLmNsYXNzTGlzdC5yZW1vdmUoJ2lzLWFjdGl2ZScpKTsKICAgICAgICBidG4uY2xhc3NMaXN0LmFkZCgnaXMtYWN0aXZlJyk7CiAgICAgICAgY29uc3QgcHJlc2V0ID0gYnRuLmRhdGFzZXQucHJlc2V0OwogICAgICAgIGNvbnN0IHRvZGF5ID0gbmV3IERhdGUoKTsKICAgICAgICBjb25zdCB0byA9IHRvZGF5LnRvSVNPU3RyaW5nKCkuc2xpY2UoMCwgMTApOwogICAgICAgIGxldCBmcm9tOwogICAgICAgIGlmIChwcmVzZXQgPT09ICdhbGwnKSB7CiAgICAgICAgICBjb25zdCBtaW4gPSAod2luZG93Ll9fZmlsdGVyT3B0aW9uc0NhY2hlICYmIHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZS5kYXRlUmFuZ2UubWluKSB8fCB0bzsKICAgICAgICAgIGZyb20gPSBtaW47CiAgICAgICAgfSBlbHNlIHsKICAgICAgICAgIGNvbnN0IGQgPSBuZXcgRGF0ZSh0b2RheSk7CiAgICAgICAgICBkLnNldERhdGUoZC5nZXREYXRlKCkgLSAoTnVtYmVyKHByZXNldCkgLSAxKSk7CiAgICAgICAgICBmcm9tID0gZC50b0lTT1N0cmluZygpLnNsaWNlKDAsIDEwKTsKICAgICAgICB9CiAgICAgICAgZGF0ZUZyb20udmFsdWUgPSBmcm9tOwogICAgICAgIGRhdGVUby52YWx1ZSA9IHRvOwogICAgICAgIGFwcGx5KCk7CiAgICAgIH0pOwogICAgfSk7CgogICAgU3RhdGUub25DaGFuZ2UoKCkgPT4gewogICAgICBpZiAoYWN0aXZlVGFiID09PSAnZGFzaGJvYXJkJykgRGFzaGJvYXJkLnJlbmRlcigpOwogICAgICBpZiAoYWN0aXZlVGFiID09PSAncmVjb3JkcycpIFJlY29yZHMucmVuZGVyKCk7CiAgICB9KTsKICB9CgogIGZ1bmN0aW9uIHdpcmVUYWJzKCkgewogICAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnLnRhYi1idG4nKS5mb3JFYWNoKChidG4pID0+IHsKICAgICAgYnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gc3dpdGNoVGFiKGJ0bi5kYXRhc2V0LnRhYikpOwogICAgfSk7CiAgfQoKICBmdW5jdGlvbiB3aXJlVGhlbWUoKSB7CiAgICBjb25zdCB0b2dnbGUgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndGhlbWVUb2dnbGUnKTsKICAgIGNvbnN0IGljb25TbG90ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3RoZW1lVG9nZ2xlSWNvbicpOwogICAgY29uc3Qgc2V0SWNvbiA9IChuYW1lKSA9PiB7IGljb25TbG90LmlubmVySFRNTCA9IGA8aSBkYXRhLWx1Y2lkZT0iJHtuYW1lfSIgc3R5bGU9IndpZHRoOjE2cHg7aGVpZ2h0OjE2cHg7Ij48L2k+YDsgfTsKICAgIGNvbnN0IHN0b3JlZCA9IGxvY2FsU3RvcmFnZS5nZXRJdGVtKCdscnMtdGhlbWUnKTsKICAgIGlmIChzdG9yZWQpIHsKICAgICAgZG9jdW1lbnQuZG9jdW1lbnRFbGVtZW50LnNldEF0dHJpYnV0ZSgnZGF0YS10aGVtZScsIHN0b3JlZCk7CiAgICAgIHNldEljb24oc3RvcmVkID09PSAnZGFyaycgPyAnc3VuJyA6ICdtb29uJyk7CiAgICB9CiAgICB0b2dnbGUuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7CiAgICAgIGNvbnN0IHByZWZlcnNEYXJrID0gd2luZG93Lm1hdGNoTWVkaWEoJyhwcmVmZXJzLWNvbG9yLXNjaGVtZTogZGFyayknKS5tYXRjaGVzOwogICAgICBjb25zdCBjdXJyZW50ID0gZG9jdW1lbnQuZG9jdW1lbnRFbGVtZW50LmdldEF0dHJpYnV0ZSgnZGF0YS10aGVtZScpIHx8IChwcmVmZXJzRGFyayA/ICdkYXJrJyA6ICdsaWdodCcpOwogICAgICBjb25zdCBuZXh0ID0gY3VycmVudCA9PT0gJ2RhcmsnID8gJ2xpZ2h0JyA6ICdkYXJrJzsKICAgICAgZG9jdW1lbnQuZG9jdW1lbnRFbGVtZW50LnNldEF0dHJpYnV0ZSgnZGF0YS10aGVtZScsIG5leHQpOwogICAgICBsb2NhbFN0b3JhZ2Uuc2V0SXRlbSgnbHJzLXRoZW1lJywgbmV4dCk7CiAgICAgIHNldEljb24obmV4dCA9PT0gJ2RhcmsnID8gJ3N1bicgOiAnbW9vbicpOwogICAgICBDaGFydHMuZGVzdHJveUFsbCgpOwogICAgICByZW5kZXJBY3RpdmVWaWV3KCk7CiAgICB9KTsKICB9CgogIHdpbmRvdy5hZGRFdmVudExpc3RlbmVyKCdscnM6ZGF0YS11cGRhdGVkJywgYXN5bmMgKCkgPT4gewogICAgYXdhaXQgbG9hZEZpbHRlck9wdGlvbnMoKTsKICAgIHJlbmRlckFjdGl2ZVZpZXcoKTsKICB9KTsKCiAgLy8gLS0tLS0tLS0tLSBBdXRoIHNjcmVlbiAtLS0tLS0tLS0tCiAgbGV0IGFwcEluaXRpYWxpemVkID0gZmFsc2U7CgogIGZ1bmN0aW9uIHNob3dBdXRoU2NyZWVuKCkgewogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2F1dGhTY3JlZW4nKS5zdHlsZS5kaXNwbGF5ID0gJ2ZsZXgnOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2FwcFNoZWxsJykuc3R5bGUuZGlzcGxheSA9ICdub25lJzsKICAgIGNvbnN0IGNvZGVJbnB1dCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhdXRoQ29kZScpOwogICAgY29kZUlucHV0LnZhbHVlID0gJyc7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXV0aEVycm9yJykudGV4dENvbnRlbnQgPSAnJzsKICAgIGNvZGVJbnB1dC5mb2N1cygpOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gc2hvd0FwcCgpIHsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhdXRoU2NyZWVuJykuc3R5bGUuZGlzcGxheSA9ICdub25lJzsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhcHBTaGVsbCcpLnN0eWxlLmRpc3BsYXkgPSAnJzsKICAgIGlmICghYXBwSW5pdGlhbGl6ZWQpIHsKICAgICAgYXBwSW5pdGlhbGl6ZWQgPSB0cnVlOwogICAgICB3aXJlVGFicygpOwogICAgICB3aXJlVGhlbWUoKTsKICAgICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2xvZ291dEJ0bicpLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgYXN5bmMgKCkgPT4gewogICAgICAgIGF3YWl0IEFwaS5hdXRoTG9nb3V0KCk7CiAgICAgICAgYXBwSW5pdGlhbGl6ZWQgPSBmYWxzZTsKICAgICAgICBzaG93QXV0aFNjcmVlbigpOwogICAgICB9KTsKICAgICAgYXdhaXQgbG9hZEZpbHRlck9wdGlvbnMoKTsKICAgICAgd2lyZUZpbHRlckJhcigpOwogICAgICBzd2l0Y2hUYWIoJ2Rhc2hib2FyZCcpOwogICAgfSBlbHNlIHsKICAgICAgYXdhaXQgbG9hZEZpbHRlck9wdGlvbnMoKTsKICAgICAgcmVuZGVyQWN0aXZlVmlldygpOwogICAgfQogIH0KCiAgYXN5bmMgZnVuY3Rpb24gc3VibWl0QXV0aCgpIHsKICAgIGNvbnN0IGVycm9yRWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXV0aEVycm9yJyk7CiAgICBjb25zdCBidG4gPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXV0aFN1Ym1pdEJ0bicpOwogICAgY29uc3QgY29kZUlucHV0ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2F1dGhDb2RlJyk7CiAgICBlcnJvckVsLnRleHRDb250ZW50ID0gJyc7CiAgICBidG4uZGlzYWJsZWQgPSB0cnVlOwogICAgYnRuLnRleHRDb250ZW50ID0gJ0NoZWNraW5n4oCmJzsKICAgIHRyeSB7CiAgICAgIGF3YWl0IEFwaS5hdXRoTG9naW4oY29kZUlucHV0LnZhbHVlKTsKICAgICAgYXdhaXQgc2hvd0FwcCgpOwogICAgfSBjYXRjaCAoZXJyKSB7CiAgICAgIGVycm9yRWwudGV4dENvbnRlbnQgPSBlcnIubWVzc2FnZTsKICAgIH0gZmluYWxseSB7CiAgICAgIGJ0bi5kaXNhYmxlZCA9IGZhbHNlOwogICAgICBidG4uaW5uZXJIVE1MID0gJzxpIGRhdGEtbHVjaWRlPSJhcnJvdy1yaWdodCIgc3R5bGU9IndpZHRoOjE0cHg7aGVpZ2h0OjE0cHg7Ij48L2k+IEVudGVyJzsKICAgIH0KICB9CgogIGZ1bmN0aW9uIHdpcmVBdXRoRm9ybSgpIHsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhdXRoU3VibWl0QnRuJykuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBzdWJtaXRBdXRoKTsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhdXRoQ29kZScpLmFkZEV2ZW50TGlzdGVuZXIoJ2tleWRvd24nLCAoZSkgPT4geyBpZiAoZS5rZXkgPT09ICdFbnRlcicpIHN1Ym1pdEF1dGgoKTsgfSk7CiAgfQoKICB3aW5kb3cuYWRkRXZlbnRMaXN0ZW5lcignbHJzOnNpZ25lZC1vdXQnLCAoKSA9PiB7CiAgICBhcHBJbml0aWFsaXplZCA9IGZhbHNlOwogICAgc2hvd0F1dGhTY3JlZW4oKTsKICB9KTsKCiAgYXN5bmMgZnVuY3Rpb24gaW5pdCgpIHsKICAgIHdpcmVBdXRoRm9ybSgpOwogICAgY29uc3QgeyBhdXRoZW50aWNhdGVkIH0gPSBhd2FpdCBBcGkuYXV0aE1lKCk7CiAgICBpZiAoYXV0aGVudGljYXRlZCkgYXdhaXQgc2hvd0FwcCgpOwogICAgZWxzZSBzaG93QXV0aFNjcmVlbigpOwogIH0KCiAgLy8gSWNvbnMgYXJlIHBsYWNlZCBhcyA8aSBkYXRhLWx1Y2lkZT0iLi4uIj4gcGxhY2Vob2xkZXJzIHRocm91Z2hvdXQgdGhlIGR5bmFtaWNhbGx5CiAgLy8gcmVuZGVyZWQgVUk7IEx1Y2lkZSByZXBsYWNlcyBlYWNoIHdpdGggYW4gaW5saW5lIFNWRy4gUmF0aGVyIHRoYW4gcmVtZW1iZXJpbmcgdG8gY2FsbAogIC8vIHRoaXMgYWZ0ZXIgZXZlcnkgc2luZ2xlIHJlbmRlciwgb25lIG9ic2VydmVyIGNhdGNoZXMgZXZlcnkgRE9NIGNoYW5nZSB0aGF0IGNvdWxkIGhhdmUKICAvLyBpbnRyb2R1Y2VkIGEgbmV3IHBsYWNlaG9sZGVyLgogIGlmICh3aW5kb3cubHVjaWRlKSB7CiAgICB3aW5kb3cubHVjaWRlLmNyZWF0ZUljb25zKCk7CiAgICAvLyBjcmVhdGVJY29ucygpIHJlcGxhY2VzIDxpIGRhdGEtbHVjaWRlPiBwbGFjZWhvbGRlcnMgd2l0aCA8c3ZnPiDigJQgaXRzZWxmIGEgRE9NCiAgICAvLyBtdXRhdGlvbi4gV2l0aG91dCBkaXNjb25uZWN0aW5nIGZpcnN0LCB0aGF0IHdyaXRlIHJlLXRyaWdnZXJzIHRoaXMgc2FtZSBvYnNlcnZlcgogICAgLy8gZm9yZXZlciAoYW4gaW5maW5pdGUgbXV0YXRlL29ic2VydmUgbG9vcCB0aGF0IHBlZ3MgdGhlIENQVSBhbmQgY3Jhc2hlcyB0aGUgdGFiKS4KICAgIC8vIERpc2Nvbm5lY3RpbmcgYmVmb3JlIGVhY2ggcGFzcyBhbmQgcmVjb25uZWN0aW5nIGFmdGVyLCBwbHVzIGJhdGNoaW5nIGJ1cnN0cyBvZgogICAgLy8gbXV0YXRpb25zIGludG8gYSBzaW5nbGUgbWljcm90YXNrLCBicmVha3MgdGhlIGN5Y2xlLgogICAgbGV0IGljb25zU2NoZWR1bGVkID0gZmFsc2U7CiAgICBjb25zdCBpY29uT2JzZXJ2ZXIgPSBuZXcgTXV0YXRpb25PYnNlcnZlcigoKSA9PiB7CiAgICAgIGlmIChpY29uc1NjaGVkdWxlZCkgcmV0dXJuOwogICAgICBpY29uc1NjaGVkdWxlZCA9IHRydWU7CiAgICAgIHF1ZXVlTWljcm90YXNrKCgpID0+IHsKICAgICAgICBpY29uc1NjaGVkdWxlZCA9IGZhbHNlOwogICAgICAgIGljb25PYnNlcnZlci5kaXNjb25uZWN0KCk7CiAgICAgICAgd2luZG93Lmx1Y2lkZS5jcmVhdGVJY29ucygpOwogICAgICAgIGljb25PYnNlcnZlci5vYnNlcnZlKGRvY3VtZW50LmJvZHksIHsgY2hpbGRMaXN0OiB0cnVlLCBzdWJ0cmVlOiB0cnVlIH0pOwogICAgICB9KTsKICAgIH0pOwogICAgaWNvbk9ic2VydmVyLm9ic2VydmUoZG9jdW1lbnQuYm9keSwgeyBjaGlsZExpc3Q6IHRydWUsIHN1YnRyZWU6IHRydWUgfSk7CiAgfQoKICBpbml0KCk7Cn0pKCk7Cjwvc2NyaXB0Pgo8L2JvZHk+CjwvaHRtbD4K';
const INDEX_HTML = Buffer.from(INDEX_HTML_BASE64, 'base64').toString('utf8');

app.get('/', (req, res) => {
  res.type('html').send(INDEX_HTML);
});

// eslint-disable-next-line no-unused-vars
app.use((err, req, res, next) => {
  console.error(err);
  res.status(500).json({ error: err.message || 'Unexpected server error.' });
});

app.listen(PORT, () => {
  console.log(`LRS Analytics Dashboard running at http://localhost:${PORT}`);
  console.log(`Access code: "${getAccessCode()}" (change it any time by editing ${CODE_FILE})`);
});
