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
   Paths / data directory
   ============================================================ */
const DATA_DIR = path.join(__dirname, 'data');
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

// ---- Frontend (embedded) ----
// Base64-encoded so the page's own inline <script> — which uses backticks and
// ${...} template literals extensively — never conflicts with the literal
// wrapping it here.
const INDEX_HTML_BASE64 = 'PCFkb2N0eXBlIGh0bWw+CjxodG1sIGxhbmc9ImVuIj4KPGhlYWQ+CjxtZXRhIGNoYXJzZXQ9IlVURi04IiAvPgo8bWV0YSBuYW1lPSJ2aWV3cG9ydCIgY29udGVudD0id2lkdGg9ZGV2aWNlLXdpZHRoLCBpbml0aWFsLXNjYWxlPTEuMCIgLz4KPHRpdGxlPkxSUyBBbmFseXRpY3MgRGFzaGJvYXJkPC90aXRsZT4KPGxpbmsgcmVsPSJwcmVjb25uZWN0IiBocmVmPSJodHRwczovL2ZvbnRzLmdvb2dsZWFwaXMuY29tIiAvPgo8bGluayByZWw9InByZWNvbm5lY3QiIGhyZWY9Imh0dHBzOi8vZm9udHMuZ3N0YXRpYy5jb20iIGNyb3Nzb3JpZ2luIC8+CjxsaW5rIGhyZWY9Imh0dHBzOi8vZm9udHMuZ29vZ2xlYXBpcy5jb20vY3NzMj9mYW1pbHk9SW50ZXI6d2dodEA0MDA7NTAwOzYwMDs3MDA7ODAwJmRpc3BsYXk9c3dhcCIgcmVsPSJzdHlsZXNoZWV0IiAvPgo8c2NyaXB0IHNyYz0iaHR0cHM6Ly9jZG4uanNkZWxpdnIubmV0L25wbS9jaGFydC5qc0A0LjQuNC9kaXN0L2NoYXJ0LnVtZC5taW4uanMiPjwvc2NyaXB0Pgo8c2NyaXB0IHNyYz0iaHR0cHM6Ly9jZG4uanNkZWxpdnIubmV0L25wbS9jaGFydGpzLXBsdWdpbi1kYXRhbGFiZWxzQDIvZGlzdC9jaGFydGpzLXBsdWdpbi1kYXRhbGFiZWxzLm1pbi5qcyI+PC9zY3JpcHQ+CjxzY3JpcHQgc3JjPSJodHRwczovL2Nkbi5qc2RlbGl2ci5uZXQvbnBtL2x1Y2lkZUAwLjQ2Mi4wL2Rpc3QvdW1kL2x1Y2lkZS5taW4uanMiPjwvc2NyaXB0Pgo8c3R5bGU+Ci8qIC0tLS0tLS0tLS0gRGVzaWduIHRva2VuczogZ2xhc3Ntb3JwaGlzbSBzdXJmYWNlcyArIHZhbGlkYXRlZCBjYXRlZ29yaWNhbC9zdGF0dXMgcGFsZXR0ZSAtLS0tLS0tLS0tICovCjpyb290IHsKICBjb2xvci1zY2hlbWU6IGxpZ2h0OwogIC0tZm9udC1zYW5zOiAnSW50ZXInLCAtYXBwbGUtc3lzdGVtLCBCbGlua01hY1N5c3RlbUZvbnQsICdTRiBQcm8gRGlzcGxheScsICdTZWdvZSBVSScsIFJvYm90bywgc2Fucy1zZXJpZjsKCiAgLS1wYWdlLXBsYW5lOiBsaW5lYXItZ3JhZGllbnQoMTgwZGVnLCAjZjdmOGZiIDAlLCAjZWNlZWYzIDEwMCUpOwogIC0tcGFnZS1wbGFuZS1zb2xpZDogI2VjZWVmMzsKICAtLXN1cmZhY2UtMTogcmdiYSgyNTUsIDI1NSwgMjU1LCAwLjY4KTsgLyogZ2xhc3M6IGNhcmRzLCBLUEkgdGlsZXMsIHRvcGJhciwgZmlsdGVyIGJhciAqLwogIC0tc3VyZmFjZS0yOiByZ2JhKDI1NSwgMjU1LCAyNTUsIDAuNTUpOyAvKiBnbGFzczogaW5wdXRzLCBuZXN0ZWQgcm93cywgcGlsbHMgKi8KICAtLXN1cmZhY2Utc29saWQ6ICNmZmZmZmY7CiAgLS1nbGFzcy1ibHVyOiBibHVyKDIwcHgpOwogIC0tYm9yZGVyOiByZ2JhKDE1LCAxNywgMjEsIDAuMDgpOwogIC0tdGV4dC1wcmltYXJ5OiAjMGYxMTE1OwogIC0tdGV4dC1zZWNvbmRhcnk6ICM1NjViNjY7CiAgLS10ZXh0LW11dGVkOiAjOGE4ZjlhOwogIC0tZ3JpZGxpbmU6IHJnYmEoMTUsIDE3LCAyMSwgMC4wOCk7CiAgLS1iYXNlbGluZTogcmdiYSgxNSwgMTcsIDIxLCAwLjIpOwogIC0tc3VjY2Vzcy10ZXh0OiAjMWE3ZDNjOwoKICAtLXN0YXR1cy1nb29kOiAjMWE5YzRhOwogIC0tc3RhdHVzLXdhcm5pbmc6ICNlMDhhMWY7CiAgLS1zdGF0dXMtc2VyaW91czogI2VjODM1YTsKICAtLXN0YXR1cy1jcml0aWNhbDogI2QxNDAzZjsKCiAgLS1zZXJpZXMtMTogIzJhNzhkNjsgLyogTFJTIGJsdWUg4oCUIGJyYW5kICsgZmFjZWJvb2sgKi8KICAtLXNlcmllcy0yOiAjMDA4MzAwOyAvKiBpbnN0YWdyYW0gKi8KICAtLXNlcmllcy0zOiAjZTg3YmE0OyAvKiB0aWt0b2sgKi8KICAtLXNlcmllcy00OiAjZWRhMTAwOyAvKiBsaW5rZWRpbiAqLwogIC0tc2VyaWVzLTU6ICMxYmFmN2E7IC8qIHRocmVhZHMgKi8KICAtLXNlcmllcy02OiAjZWI2ODM0OyAvKiB5b3V0dWJlICovCiAgLS1zZXJpZXMtNzogIzRhM2FhNzsgLyogcmVzZXJ2ZWQgKi8KICAtLXNlcmllcy04OiAjZTM0OTQ4OyAvKiByZXNlcnZlZCAqLwoKICAtLXJhZGl1cy1zbTogMTBweDsKICAtLXJhZGl1cy1tZDogMTRweDsKICAtLXJhZGl1cy1sZzogMThweDsKCiAgLS1zaGFkb3ctY2FyZDogMCAxcHggMnB4IHJnYmEoMTUsMTcsMjEsMC4wMyksIDAgOHB4IDI0cHggLTEwcHggcmdiYSgxNSwxNywyMSwwLjE0KTsKICAtLXNoYWRvdy1ob3ZlcjogMCA2cHggMTJweCAtMnB4IHJnYmEoMTUsMTcsMjEsMC4wOCksIDAgMThweCA0MHB4IC0xNHB4IHJnYmEoMTUsMTcsMjEsMC4yKTsKICAtLXNoYWRvdy1tb2RhbDogMCAyNHB4IDY0cHggLTEycHggcmdiYSgxNSwxNywyMSwwLjM1KTsKICAtLWVhc2U6IGN1YmljLWJlemllcigwLjQsIDAsIDAuMiwgMSk7Cn0KCkBtZWRpYSAocHJlZmVycy1jb2xvci1zY2hlbWU6IGRhcmspIHsKICA6cm9vdDp3aGVyZSg6bm90KFtkYXRhLXRoZW1lPSJsaWdodCJdKSkgewogICAgY29sb3Itc2NoZW1lOiBkYXJrOwogICAgLS1wYWdlLXBsYW5lOiBsaW5lYXItZ3JhZGllbnQoMTgwZGVnLCAjMGIwYzBmIDAlLCAjMTcxODFkIDEwMCUpOwogICAgLS1wYWdlLXBsYW5lLXNvbGlkOiAjMGIwYzBmOwogICAgLS1zdXJmYWNlLTE6IHJnYmEoMzIsIDM0LCA0MCwgMC42Mik7CiAgICAtLXN1cmZhY2UtMjogcmdiYSgyNTUsIDI1NSwgMjU1LCAwLjA2KTsKICAgIC0tc3VyZmFjZS1zb2xpZDogIzFjMWQyMzsKICAgIC0tYm9yZGVyOiByZ2JhKDI1NSwgMjU1LCAyNTUsIDAuMDkpOwogICAgLS10ZXh0LXByaW1hcnk6ICNmNGY1Zjc7CiAgICAtLXRleHQtc2Vjb25kYXJ5OiAjYjhiYmM0OwogICAgLS10ZXh0LW11dGVkOiAjODI4NjhmOwogICAgLS1ncmlkbGluZTogcmdiYSgyNTUsIDI1NSwgMjU1LCAwLjA4KTsKICAgIC0tYmFzZWxpbmU6IHJnYmEoMjU1LCAyNTUsIDI1NSwgMC4yKTsKICAgIC0tc3VjY2Vzcy10ZXh0OiAjMzRjNzZmOwoKICAgIC0tc3RhdHVzLWdvb2Q6ICMyZmI4NjI7CiAgICAtLXN0YXR1cy13YXJuaW5nOiAjZjBhMTNhOwogICAgLS1zdGF0dXMtc2VyaW91czogI2VjODM1YTsKICAgIC0tc3RhdHVzLWNyaXRpY2FsOiAjZTA2MDVmOwoKICAgIC0tc2VyaWVzLTE6ICMzOTg3ZTU7CiAgICAtLXNlcmllcy0yOiAjMDA4MzAwOwogICAgLS1zZXJpZXMtMzogI2Q1NTE4MTsKICAgIC0tc2VyaWVzLTQ6ICNjOTg1MDA7CiAgICAtLXNlcmllcy01OiAjMTk5ZTcwOwogICAgLS1zZXJpZXMtNjogI2Q5NTkyNjsKICAgIC0tc2VyaWVzLTc6ICM5MDg1ZTk7CiAgICAtLXNlcmllcy04OiAjZTY2NzY3OwoKICAgIC0tc2hhZG93LWNhcmQ6IDAgMXB4IDJweCByZ2JhKDAsMCwwLDAuMiksIDAgOHB4IDI0cHggLTEwcHggcmdiYSgwLDAsMCwwLjUpOwogICAgLS1zaGFkb3ctaG92ZXI6IDAgNnB4IDEycHggLTJweCByZ2JhKDAsMCwwLDAuMyksIDAgMThweCA0MHB4IC0xNHB4IHJnYmEoMCwwLDAsMC42KTsKICAgIC0tc2hhZG93LW1vZGFsOiAwIDI0cHggNjRweCAtMTJweCByZ2JhKDAsMCwwLDAuNyk7CiAgfQp9Cjpyb290W2RhdGEtdGhlbWU9ImRhcmsiXSB7CiAgY29sb3Itc2NoZW1lOiBkYXJrOwogIC0tcGFnZS1wbGFuZTogbGluZWFyLWdyYWRpZW50KDE4MGRlZywgIzBiMGMwZiAwJSwgIzE3MTgxZCAxMDAlKTsKICAtLXBhZ2UtcGxhbmUtc29saWQ6ICMwYjBjMGY7CiAgLS1zdXJmYWNlLTE6IHJnYmEoMzIsIDM0LCA0MCwgMC42Mik7CiAgLS1zdXJmYWNlLTI6IHJnYmEoMjU1LCAyNTUsIDI1NSwgMC4wNik7CiAgLS1zdXJmYWNlLXNvbGlkOiAjMWMxZDIzOwogIC0tYm9yZGVyOiByZ2JhKDI1NSwgMjU1LCAyNTUsIDAuMDkpOwogIC0tdGV4dC1wcmltYXJ5OiAjZjRmNWY3OwogIC0tdGV4dC1zZWNvbmRhcnk6ICNiOGJiYzQ7CiAgLS10ZXh0LW11dGVkOiAjODI4NjhmOwogIC0tZ3JpZGxpbmU6IHJnYmEoMjU1LCAyNTUsIDI1NSwgMC4wOCk7CiAgLS1iYXNlbGluZTogcmdiYSgyNTUsIDI1NSwgMjU1LCAwLjIpOwogIC0tc3VjY2Vzcy10ZXh0OiAjMzRjNzZmOwoKICAtLXN0YXR1cy1nb29kOiAjMmZiODYyOwogIC0tc3RhdHVzLXdhcm5pbmc6ICNmMGExM2E7CiAgLS1zdGF0dXMtc2VyaW91czogI2VjODM1YTsKICAtLXN0YXR1cy1jcml0aWNhbDogI2UwNjA1ZjsKCiAgLS1zZXJpZXMtMTogIzM5ODdlNTsKICAtLXNlcmllcy0yOiAjMDA4MzAwOwogIC0tc2VyaWVzLTM6ICNkNTUxODE7CiAgLS1zZXJpZXMtNDogI2M5ODUwMDsKICAtLXNlcmllcy01OiAjMTk5ZTcwOwogIC0tc2VyaWVzLTY6ICNkOTU5MjY7CiAgLS1zZXJpZXMtNzogIzkwODVlOTsKICAtLXNlcmllcy04OiAjZTY2NzY3OwoKICAtLXNoYWRvdy1jYXJkOiAwIDFweCAycHggcmdiYSgwLDAsMCwwLjIpLCAwIDhweCAyNHB4IC0xMHB4IHJnYmEoMCwwLDAsMC41KTsKICAtLXNoYWRvdy1ob3ZlcjogMCA2cHggMTJweCAtMnB4IHJnYmEoMCwwLDAsMC4zKSwgMCAxOHB4IDQwcHggLTE0cHggcmdiYSgwLDAsMCwwLjYpOwogIC0tc2hhZG93LW1vZGFsOiAwIDI0cHggNjRweCAtMTJweCByZ2JhKDAsMCwwLDAuNyk7Cn0KCiogeyBib3gtc2l6aW5nOiBib3JkZXItYm94OyB9Cmh0bWwsIGJvZHkgeyBoZWlnaHQ6IDEwMCU7IH0KYm9keSB7CiAgbWFyZ2luOiAwOwogIGZvbnQtZmFtaWx5OiB2YXIoLS1mb250LXNhbnMpOwogIGJhY2tncm91bmQ6IHZhcigtLXBhZ2UtcGxhbmUpOwogIGJhY2tncm91bmQtYXR0YWNobWVudDogZml4ZWQ7CiAgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7CiAgLXdlYmtpdC1mb250LXNtb290aGluZzogYW50aWFsaWFzZWQ7CiAgLW1vei1vc3gtZm9udC1zbW9vdGhpbmc6IGdyYXlzY2FsZTsKfQpidXR0b24sIHNlbGVjdCwgaW5wdXQsIHRleHRhcmVhIHsgZm9udC1mYW1pbHk6IGluaGVyaXQ7IH0KaDEsIGgyLCBoMywgaDQgeyBmb250LXdlaWdodDogNzAwOyBsZXR0ZXItc3BhY2luZzogLTAuMDFlbTsgfQoKOjpzZWxlY3Rpb24geyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDMwJSwgdHJhbnNwYXJlbnQpOyB9CgovKiBDdXN0b20gc2Nyb2xsYmFyIOKAlCB0aGluLCB1bm9idHJ1c2l2ZSwgZml0cyB0aGUgZ2xhc3MgYWVzdGhldGljICovCjo6LXdlYmtpdC1zY3JvbGxiYXIgeyB3aWR0aDogMTBweDsgaGVpZ2h0OiAxMHB4OyB9Cjo6LXdlYmtpdC1zY3JvbGxiYXItdHJhY2sgeyBiYWNrZ3JvdW5kOiB0cmFuc3BhcmVudDsgfQo6Oi13ZWJraXQtc2Nyb2xsYmFyLXRodW1iIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXRleHQtbXV0ZWQpIDQwJSwgdHJhbnNwYXJlbnQpOyBib3JkZXItcmFkaXVzOiAyMHB4OyBib3JkZXI6IDJweCBzb2xpZCB0cmFuc3BhcmVudDsgYmFja2dyb3VuZC1jbGlwOiBwYWRkaW5nLWJveDsgfQo6Oi13ZWJraXQtc2Nyb2xsYmFyLXRodW1iOmhvdmVyIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXRleHQtbXV0ZWQpIDYwJSwgdHJhbnNwYXJlbnQpOyBiYWNrZ3JvdW5kLWNsaXA6IHBhZGRpbmctYm94OyB9CgouYXBwLXNoZWxsIHsgbWluLWhlaWdodDogMTAwJTsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgfQoKLyogLS0tLS0tLS0tLSBUb3BiYXIgLS0tLS0tLS0tLSAqLwoudG9wYmFyIHsKICBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDI0cHg7CiAgcGFkZGluZzogMTJweCAyMHB4OwogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOwogIHBvc2l0aW9uOiBzdGlja3k7IHRvcDogMDsgei1pbmRleDogMjA7Cn0KLnRvcGJhci1icmFuZCB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBiYXNlbGluZTsgZ2FwOiA4cHg7IHdoaXRlLXNwYWNlOiBub3dyYXA7IH0KLmJyYW5kLW1hcmsgewogIGZvbnQtd2VpZ2h0OiA3MDA7IGxldHRlci1zcGFjaW5nOiAwLjAyZW07CiAgYmFja2dyb3VuZDogdmFyKC0tc2VyaWVzLTEpOyBjb2xvcjogI2ZmZjsKICBwYWRkaW5nOiA0cHggOXB4OyBib3JkZXItcmFkaXVzOiA4cHg7IGZvbnQtc2l6ZTogMTNweDsKICBib3gtc2hhZG93OiAwIDRweCAxNHB4IC00cHggY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSA2MCUsIHRyYW5zcGFyZW50KTsKfQouYnJhbmQtdGl0bGUgeyBmb250LXdlaWdodDogNjAwOyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgbGV0dGVyLXNwYWNpbmc6IC0wLjAxZW07IH0KCi50YWJzIHsgZGlzcGxheTogZmxleDsgZ2FwOiAycHg7IGZsZXg6IDE7IG92ZXJmbG93LXg6IGF1dG87IHBvc2l0aW9uOiByZWxhdGl2ZTsgfQoudGFiLWJ0biB7CiAgZGlzcGxheTogaW5saW5lLWZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogN3B4OwogIGJvcmRlcjogbm9uZTsgYmFja2dyb3VuZDogdHJhbnNwYXJlbnQ7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7CiAgcGFkZGluZzogOXB4IDE2cHg7IGJvcmRlci1yYWRpdXM6IDEwcHg7IGN1cnNvcjogcG9pbnRlcjsgZm9udC1zaXplOiAxNHB4OyBmb250LXdlaWdodDogNTAwOwogIHdoaXRlLXNwYWNlOiBub3dyYXA7IHBvc2l0aW9uOiByZWxhdGl2ZTsKICB0cmFuc2l0aW9uOiBjb2xvciAxODBtcyB2YXIoLS1lYXNlKSwgYmFja2dyb3VuZCAxODBtcyB2YXIoLS1lYXNlKTsKfQoudGFiLWJ0biBzdmcgeyBmbGV4LXNocmluazogMDsgb3BhY2l0eTogMC44OyB9Ci50YWItYnRuLmlzLWFjdGl2ZSBzdmcgeyBvcGFjaXR5OiAxOyBjb2xvcjogdmFyKC0tc2VyaWVzLTEpOyB9Ci50YWItYnRuOmhvdmVyIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSA4JSwgdHJhbnNwYXJlbnQpOyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgfQoudGFiLWJ0bi5pcy1hY3RpdmUgeyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgZm9udC13ZWlnaHQ6IDYwMDsgfQoudGFiLWJ0bi5pcy1hY3RpdmU6OmFmdGVyIHsKICBjb250ZW50OiAnJzsgcG9zaXRpb246IGFic29sdXRlOyBsZWZ0OiAxMnB4OyByaWdodDogMTJweDsgYm90dG9tOiAtMXB4OyBoZWlnaHQ6IDJweDsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zZXJpZXMtMSk7IGJvcmRlci1yYWRpdXM6IDJweCAycHggMCAwOwogIGFuaW1hdGlvbjogdGFiSW5kaWNhdG9ySW4gMjIwbXMgdmFyKC0tZWFzZSk7Cn0KQGtleWZyYW1lcyB0YWJJbmRpY2F0b3JJbiB7IGZyb20geyBvcGFjaXR5OiAwOyB0cmFuc2Zvcm06IHNjYWxlWCgwLjQpOyB9IHRvIHsgb3BhY2l0eTogMTsgdHJhbnNmb3JtOiBzY2FsZVgoMSk7IH0gfQoKLnRoZW1lLXRvZ2dsZSB7CiAgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0yKTsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHdpZHRoOiAzNnB4OyBoZWlnaHQ6IDM2cHg7IGJvcmRlci1yYWRpdXM6IDEwcHg7IGN1cnNvcjogcG9pbnRlcjsgZm9udC1zaXplOiAxNXB4OwogIGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogY2VudGVyOwogIHRyYW5zaXRpb246IHRyYW5zZm9ybSAxODBtcyB2YXIoLS1lYXNlKSwgYm94LXNoYWRvdyAxODBtcyB2YXIoLS1lYXNlKSwgYmFja2dyb3VuZCAxODBtcyB2YXIoLS1lYXNlKTsKfQoudGhlbWUtdG9nZ2xlOmhvdmVyIHsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKC0xcHgpOyBib3gtc2hhZG93OiB2YXIoLS1zaGFkb3ctaG92ZXIpOyB9Ci50aGVtZS10b2dnbGU6YWN0aXZlIHsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDApIHNjYWxlKDAuOTQpOyB9Ci50b3BiYXItdXNlciB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogMTBweDsgZm9udC1zaXplOiAxM3B4OyB9CgovKiAtLS0tLS0tLS0tIEF1dGggc2NyZWVuIC0tLS0tLS0tLS0gKi8KLmF1dGgtc2NyZWVuIHsKICBtaW4taGVpZ2h0OiAxMDB2aDsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsganVzdGlmeS1jb250ZW50OiBjZW50ZXI7CiAgYmFja2dyb3VuZDogdmFyKC0tcGFnZS1wbGFuZSk7IHBhZGRpbmc6IDIwcHg7Cn0KLmF1dGgtY2FyZCB7CiAgd2lkdGg6IDEwMCU7IG1heC13aWR0aDogNDAwcHg7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1sZyk7IHBhZGRpbmc6IDMycHg7IGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdy1tb2RhbCk7CiAgYW5pbWF0aW9uOiBtb2RhbFBhbmVsSW4gMjYwbXMgdmFyKC0tZWFzZSk7Cn0KLmF1dGgtYnJhbmQgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogYmFzZWxpbmU7IGdhcDogOHB4OyBtYXJnaW4tYm90dG9tOiAyMnB4OyB9Ci5hdXRoLWJyYW5kIC5icmFuZC10aXRsZSB7IGZvbnQtd2VpZ2h0OiA3MDA7IGZvbnQtc2l6ZTogMTdweDsgfQouYXV0aC1mb3JtIHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiAxNHB4OyBtYXJnaW4tdG9wOiAxNnB4OyB9Ci5hdXRoLWZvcm0gLmZvcm0tZmllbGQgaW5wdXQgeyB3aWR0aDogMTAwJTsgfQouYXV0aC1lcnJvciB7IGNvbG9yOiB2YXIoLS1zdGF0dXMtY3JpdGljYWwpOyBmb250LXNpemU6IDEycHg7IG1pbi1oZWlnaHQ6IDE2cHg7IH0KCi8qIC0tLS0tLS0tLS0gRmlsdGVyIGJhciAtLS0tLS0tLS0tICovCi5maWx0ZXItYmFyIHsKICBkaXNwbGF5OiBmbGV4OyBmbGV4LXdyYXA6IHdyYXA7IGFsaWduLWl0ZW1zOiBlbmQ7IGdhcDogMTZweDsKICBwYWRkaW5nOiAxNHB4IDIwcHg7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOwogIHBvc2l0aW9uOiBzdGlja3k7IHRvcDogNTdweDsgei1pbmRleDogMTk7Cn0KLmZpbHRlci1maWVsZCB7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGdhcDogNXB4OyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IH0KLmZpbHRlci1maWVsZCBsYWJlbCB7IGZvbnQtd2VpZ2h0OiA2MDA7IH0KLmZpbHRlci1wcmVzZXRzIHsgZmxleC1kaXJlY3Rpb246IHJvdzsgZ2FwOiA2cHg7IH0KLmZpbHRlci1wcmVzZXRzIGJ1dHRvbiB7CiAgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0yKTsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgYm9yZGVyLXJhZGl1czogMjBweDsgcGFkZGluZzogN3B4IDEzcHg7IGZvbnQtc2l6ZTogMTJweDsgZm9udC13ZWlnaHQ6IDUwMDsgY3Vyc29yOiBwb2ludGVyOwogIHRyYW5zaXRpb246IGNvbG9yIDE4MG1zIHZhcigtLWVhc2UpLCBiYWNrZ3JvdW5kIDE4MG1zIHZhcigtLWVhc2UpLCB0cmFuc2Zvcm0gMTUwbXMgdmFyKC0tZWFzZSk7Cn0KLmZpbHRlci1wcmVzZXRzIGJ1dHRvbjpob3ZlciB7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoLTFweCk7IH0KLmZpbHRlci1wcmVzZXRzIGJ1dHRvbjphY3RpdmUgeyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMCkgc2NhbGUoMC45Nik7IH0KLmZpbHRlci1wcmVzZXRzIGJ1dHRvbi5pcy1hY3RpdmUgeyBiYWNrZ3JvdW5kOiB2YXIoLS1zZXJpZXMtMSk7IGNvbG9yOiAjZmZmOyBib3JkZXItY29sb3I6IHRyYW5zcGFyZW50OyBib3gtc2hhZG93OiAwIDRweCAxNHB4IC01cHggY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSA2MCUsIHRyYW5zcGFyZW50KTsgfQoKLyogLS0tLS0tLS0tLSBWaWV3IGFyZWEgLS0tLS0tLS0tLSAqLwoudmlldy1hcmVhIHsgZmxleDogMTsgcGFkZGluZzogMjRweDsgbWF4LXdpZHRoOiAxNDAwcHg7IHdpZHRoOiAxMDAlOyBtYXJnaW46IDAgYXV0bzsgfQoudmlldyB7IGRpc3BsYXk6IG5vbmU7IH0KLnZpZXcuaXMtYWN0aXZlIHsgZGlzcGxheTogYmxvY2s7IGFuaW1hdGlvbjogdmlld0ZhZGVJbiAyNjBtcyB2YXIoLS1lYXNlKTsgfQpAa2V5ZnJhbWVzIHZpZXdGYWRlSW4gewogIGZyb20geyBvcGFjaXR5OiAwOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoNnB4KTsgfQogIHRvIHsgb3BhY2l0eTogMTsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDApOyB9Cn0KCi5zZWN0aW9uLXRpdGxlIHsgZm9udC1zaXplOiAxNnB4OyBmb250LXdlaWdodDogNzAwOyBtYXJnaW46IDMycHggMCAxNHB4OyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgbGV0dGVyLXNwYWNpbmc6IC0wLjAxZW07IH0KLnNlY3Rpb24tdGl0bGU6Zmlyc3QtY2hpbGQgeyBtYXJnaW4tdG9wOiAwOyB9CgovKiAtLS0tLS0tLS0tIElucHV0cyDigJQgb25lIHNoYXJlZCBnbGFzcyB0cmVhdG1lbnQgZm9yIGV2ZXJ5IHRleHQgaW5wdXQsIHNlbGVjdCwgYW5kIGRhdGUgcGlja2VyIC0tLS0tLS0tLS0gKi8KLmZpbHRlci1maWVsZCBzZWxlY3QsIC5maWx0ZXItZmllbGQgaW5wdXRbdHlwZT0iZGF0ZSJdLAouZm9ybS1maWVsZCBpbnB1dCwgLmZvcm0tZmllbGQgc2VsZWN0LCAuZm9ybS1maWVsZCB0ZXh0YXJlYSwKLmRhc2hib2FyZC1jb250cm9scyBzZWxlY3QsIC5yZWNvcmRzLXNlYXJjaCBpbnB1dCwKLmZpZWxkLWlubGluZSBzZWxlY3QsIC5maWVsZC1pbmxpbmUgaW5wdXQsCi5jb25mbGljdC1yb3cgc2VsZWN0LCAuY2FyZC1oZWFkZXIgc2VsZWN0IHsKICBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtc20pOwogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMik7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyBmb250LXNpemU6IDEzcHg7CiAgcGFkZGluZzogOHB4IDEycHg7IG1pbi13aWR0aDogMTQwcHg7CiAgdHJhbnNpdGlvbjogYm9yZGVyLWNvbG9yIDE2MG1zIHZhcigtLWVhc2UpLCBib3gtc2hhZG93IDE2MG1zIHZhcigtLWVhc2UpOwp9Ci5maWx0ZXItZmllbGQgc2VsZWN0OmhvdmVyLCAuZmlsdGVyLWZpZWxkIGlucHV0W3R5cGU9ImRhdGUiXTpob3ZlciwKLmZvcm0tZmllbGQgaW5wdXQ6aG92ZXIsIC5mb3JtLWZpZWxkIHNlbGVjdDpob3ZlciwKLmRhc2hib2FyZC1jb250cm9scyBzZWxlY3Q6aG92ZXIsIC5yZWNvcmRzLXNlYXJjaCBpbnB1dDpob3ZlciwKLmZpZWxkLWlubGluZSBzZWxlY3Q6aG92ZXIsIC5maWVsZC1pbmxpbmUgaW5wdXQ6aG92ZXIsCi5jb25mbGljdC1yb3cgc2VsZWN0OmhvdmVyLCAuY2FyZC1oZWFkZXIgc2VsZWN0OmhvdmVyIHsKICBib3JkZXItY29sb3I6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zZXJpZXMtMSkgMzUlLCB2YXIoLS1ib3JkZXIpKTsKfQouZmlsdGVyLWZpZWxkIHNlbGVjdDpmb2N1cywgLmZpbHRlci1maWVsZCBpbnB1dFt0eXBlPSJkYXRlIl06Zm9jdXMsCi5mb3JtLWZpZWxkIGlucHV0OmZvY3VzLCAuZm9ybS1maWVsZCBzZWxlY3Q6Zm9jdXMsIC5mb3JtLWZpZWxkIHRleHRhcmVhOmZvY3VzLAouZGFzaGJvYXJkLWNvbnRyb2xzIHNlbGVjdDpmb2N1cywgLnJlY29yZHMtc2VhcmNoIGlucHV0OmZvY3VzLAouZmllbGQtaW5saW5lIHNlbGVjdDpmb2N1cywgLmZpZWxkLWlubGluZSBpbnB1dDpmb2N1cywKLmNvbmZsaWN0LXJvdyBzZWxlY3Q6Zm9jdXMsIC5jYXJkLWhlYWRlciBzZWxlY3Q6Zm9jdXMsCi5hdXRoLWZvcm0gaW5wdXQ6Zm9jdXMgewogIG91dGxpbmU6IG5vbmU7IGJvcmRlci1jb2xvcjogdmFyKC0tc2VyaWVzLTEpOwogIGJveC1zaGFkb3c6IDAgMCAwIDNweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDE4JSwgdHJhbnNwYXJlbnQpOwp9CgovKiAtLS0tLS0tLS0tIFN0YXQgdGlsZXMgLS0tLS0tLS0tLSAqLwouc3RhdC1ncmlkIHsKICBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdChhdXRvLWZpdCwgbWlubWF4KDE4MHB4LCAxZnIpKTsgZ2FwOiAxNHB4Owp9Ci5zdGF0LXRpbGUgewogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1tZCk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHBhZGRpbmc6IDE2cHggMThweDsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWNhcmQpOwogIHRyYW5zaXRpb246IHRyYW5zZm9ybSAyMDBtcyB2YXIoLS1lYXNlKSwgYm94LXNoYWRvdyAyMDBtcyB2YXIoLS1lYXNlKTsKICBhbmltYXRpb246IGNhcmRJbiAzMjBtcyB2YXIoLS1lYXNlKSBiYWNrd2FyZHM7Cn0KLnN0YXQtdGlsZTpob3ZlciB7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgtM3B4KTsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWhvdmVyKTsgfQouc3RhdC1sYWJlbCB7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgZm9udC13ZWlnaHQ6IDYwMDsgfQouc3RhdC12YWx1ZSB7IGZvbnQtc2l6ZTogMjdweDsgZm9udC13ZWlnaHQ6IDcwMDsgbWFyZ2luLXRvcDogNXB4OyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgbGV0dGVyLXNwYWNpbmc6IC0wLjAyZW07IH0KLnN0YXQtZGVsdGEgeyBmb250LXNpemU6IDEycHg7IG1hcmdpbi10b3A6IDdweDsgZm9udC13ZWlnaHQ6IDYwMDsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA0cHg7IH0KLnN0YXQtZGVsdGEudXAgeyBjb2xvcjogdmFyKC0tc3VjY2Vzcy10ZXh0KTsgfQouc3RhdC1kZWx0YS5kb3duIHsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IH0KLnN0YXQtZGVsdGEuZmxhdCB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgfQouc3RhdC1kZWx0YS51cDo6YmVmb3JlIHsgY29udGVudDogJ+KGkSc7IH0KLnN0YXQtZGVsdGEuZG93bjo6YmVmb3JlIHsgY29udGVudDogJ+KGkyc7IH0KCkBrZXlmcmFtZXMgY2FyZEluIHsKICBmcm9tIHsgb3BhY2l0eTogMDsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDEwcHgpOyB9CiAgdG8geyBvcGFjaXR5OiAxOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMCk7IH0KfQoKLyogLS0tLS0tLS0tLSBDYXJkcyAvIGNoYXJ0cyAtLS0tLS0tLS0tICovCi5jYXJkLWdyaWQgeyBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IDJmciAxZnI7IGdhcDogMTZweDsgYWxpZ24taXRlbXM6IHN0YXJ0OyB9Ci5jYXJkLWdyaWQuZXZlbiB7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogMWZyIDFmcjsgfQpAbWVkaWEgKG1heC13aWR0aDogOTAwcHgpIHsgLmNhcmQtZ3JpZCwgLmNhcmQtZ3JpZC5ldmVuIHsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiAxZnI7IH0gfQouY2FyZCB7CiAgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLWxnKTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgcGFkZGluZzogMThweDsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWNhcmQpOwogIHRyYW5zaXRpb246IGJveC1zaGFkb3cgMjIwbXMgdmFyKC0tZWFzZSksIHRyYW5zZm9ybSAyMjBtcyB2YXIoLS1lYXNlKTsKICBhbmltYXRpb246IGNhcmRJbiAzMjBtcyB2YXIoLS1lYXNlKSBiYWNrd2FyZHM7Cn0KLmNhcmQ6aG92ZXIgeyBib3gtc2hhZG93OiB2YXIoLS1zaGFkb3ctaG92ZXIpOyB9Ci5jYXJkLWhlYWRlciB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogc3BhY2UtYmV0d2VlbjsgZ2FwOiA4cHg7IG1hcmdpbi1ib3R0b206IDE0cHg7IH0KLmNhcmQtaGVhZGVyIGgzIHsgZm9udC1zaXplOiAxNHB4OyBtYXJnaW46IDA7IGZvbnQtd2VpZ2h0OiA3MDA7IGxldHRlci1zcGFjaW5nOiAtMC4wMDVlbTsgfQouY2FyZC1oZWFkZXIgc2VsZWN0IHsgZm9udC1zaXplOiAxMnB4OyBwYWRkaW5nOiA2cHggMTBweDsgbWluLXdpZHRoOiAwOyB9Ci5jaGFydC13cmFwIHsgcG9zaXRpb246IHJlbGF0aXZlOyBoZWlnaHQ6IDI4MHB4OyB9Ci5jaGFydC13cmFwLnRhbGwgeyBoZWlnaHQ6IDM0MHB4OyB9CgoubGVnZW5kLXJvdyB7IGRpc3BsYXk6IGZsZXg7IGZsZXgtd3JhcDogd3JhcDsgZ2FwOiAxMnB4OyBtYXJnaW4tdG9wOiAxMHB4OyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IH0KLmxlZ2VuZC1pdGVtIHsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA2cHg7IH0KLmxlZ2VuZC1zd2F0Y2ggeyB3aWR0aDogMTBweDsgaGVpZ2h0OiAxMHB4OyBib3JkZXItcmFkaXVzOiAzcHg7IGRpc3BsYXk6IGlubGluZS1ibG9jazsgfQoubGVnZW5kLWxpbmUgeyB3aWR0aDogMTRweDsgaGVpZ2h0OiAycHg7IGJvcmRlci1yYWRpdXM6IDJweDsgZGlzcGxheTogaW5saW5lLWJsb2NrOyB9CgovKiAtLS0tLS0tLS0tIFRhYmxlcyDigJQgcHJlbWl1bSBkYXRhYmFzZSBmZWVsLCBub3QgYSBzcHJlYWRzaGVldCAtLS0tLS0tLS0tICovCi50YWJsZS1zY3JvbGwgewogIG92ZXJmbG93LXg6IGF1dG87IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1tZCk7CiAgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0yKTsKfQouZGF0YS10YWJsZSB7IHdpZHRoOiAxMDAlOyBib3JkZXItY29sbGFwc2U6IHNlcGFyYXRlOyBib3JkZXItc3BhY2luZzogMDsgZm9udC1zaXplOiAxM3B4OyB9Ci5kYXRhLXRhYmxlIHRoLCAuZGF0YS10YWJsZSB0ZCB7IHRleHQtYWxpZ246IGxlZnQ7IHBhZGRpbmc6IDExcHggMTRweDsgYm9yZGVyLWJvdHRvbTogMXB4IHNvbGlkIHZhcigtLWdyaWRsaW5lKTsgfQouZGF0YS10YWJsZSB0aGVhZCB0aCB7CiAgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgZm9udC13ZWlnaHQ6IDYwMDsgZm9udC1zaXplOiAxMXB4OyB0ZXh0LXRyYW5zZm9ybTogdXBwZXJjYXNlOyBsZXR0ZXItc3BhY2luZzogMC4wNGVtOwogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBwb3NpdGlvbjogc3RpY2t5OyB0b3A6IDA7IHotaW5kZXg6IDE7Cn0KLmRhdGEtdGFibGUgdGJvZHkgdHI6bnRoLWNoaWxkKGV2ZW4pIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXRleHQtbXV0ZWQpIDQlLCB0cmFuc3BhcmVudCk7IH0KLmRhdGEtdGFibGUgdGQubnVtIHsgZm9udC12YXJpYW50LW51bWVyaWM6IHRhYnVsYXItbnVtczsgdGV4dC1hbGlnbjogcmlnaHQ7IH0KLmRhdGEtdGFibGUgdGgubnVtIHsgdGV4dC1hbGlnbjogcmlnaHQ7IH0KLmRhdGEtdGFibGUgdGJvZHkgdHIgeyB0cmFuc2l0aW9uOiBiYWNrZ3JvdW5kIDE1MG1zIHZhcigtLWVhc2UpOyB9Ci5kYXRhLXRhYmxlIHRib2R5IHRyOmhvdmVyIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSA3JSwgdHJhbnNwYXJlbnQpOyB9Ci5kYXRhLXRhYmxlIHRib2R5IHRyOmxhc3QtY2hpbGQgdGQgeyBib3JkZXItYm90dG9tOiBub25lOyB9Ci5wbGF0Zm9ybS1waWxsIHsKICBkaXNwbGF5OiBpbmxpbmUtZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA2cHg7IGZvbnQtc2l6ZTogMTJweDsgZm9udC13ZWlnaHQ6IDYwMDsKICBwYWRkaW5nOiA0cHggMTBweDsgYm9yZGVyLXJhZGl1czogMjBweDsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsKfQoucGxhdGZvcm0tZG90IHsgd2lkdGg6IDhweDsgaGVpZ2h0OiA4cHg7IGJvcmRlci1yYWRpdXM6IDUwJTsgfQoKLyogLS0tLS0tLS0tLSBCdXR0b25zIOKAlCBuZXZlciBmbGF0OiBzb2Z0IHNoYWRvdywgaG92ZXIgbGlmdCwgcHJlc3Mgc2NhbGUgLS0tLS0tLS0tLSAqLwouYnRuIHsKICBkaXNwbGF5OiBpbmxpbmUtZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsganVzdGlmeS1jb250ZW50OiBjZW50ZXI7IGdhcDogNnB4OwogIGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMik7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOwogIGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBwYWRkaW5nOiA5cHggMTdweDsgYm9yZGVyLXJhZGl1czogMTFweDsgY3Vyc29yOiBwb2ludGVyOyBmb250LXNpemU6IDEzcHg7IGZvbnQtd2VpZ2h0OiA2MDA7CiAgYm94LXNoYWRvdzogMCAxcHggMnB4IHJnYmEoMTUsMTcsMjEsMC4wNCk7CiAgdHJhbnNpdGlvbjogdHJhbnNmb3JtIDE1MG1zIHZhcigtLWVhc2UpLCBib3gtc2hhZG93IDE1MG1zIHZhcigtLWVhc2UpLCBmaWx0ZXIgMTUwbXMgdmFyKC0tZWFzZSksIGJhY2tncm91bmQgMTUwbXMgdmFyKC0tZWFzZSk7Cn0KLmJ0biBzdmcgeyBmbGV4LXNocmluazogMDsgfQouYnRuOmhvdmVyIHsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKC0xcHgpOyBib3gtc2hhZG93OiB2YXIoLS1zaGFkb3ctaG92ZXIpOyBmaWx0ZXI6IGJyaWdodG5lc3MoMS4wMik7IH0KLmJ0bjphY3RpdmUgeyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMCkgc2NhbGUoMC45Nik7IGJveC1zaGFkb3c6IDAgMXB4IDJweCByZ2JhKDE1LDE3LDIxLDAuMDYpOyB9Ci5idG4ucHJpbWFyeSB7CiAgYmFja2dyb3VuZDogdmFyKC0tc2VyaWVzLTEpOyBjb2xvcjogI2ZmZjsgYm9yZGVyLWNvbG9yOiB0cmFuc3BhcmVudDsKICBib3gtc2hhZG93OiAwIDRweCAxNHB4IC01cHggY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSA2NSUsIHRyYW5zcGFyZW50KTsKfQouYnRuLnByaW1hcnk6aG92ZXIgeyBmaWx0ZXI6IGJyaWdodG5lc3MoMS4wNyk7IGJveC1zaGFkb3c6IDAgOHB4IDIycHggLTZweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDcwJSwgdHJhbnNwYXJlbnQpOyB9Ci5idG4uZGFuZ2VyIHsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdGF0dXMtY3JpdGljYWwpOyBjb2xvcjogI2ZmZjsgYm9yZGVyLWNvbG9yOiB0cmFuc3BhcmVudDsKICBib3gtc2hhZG93OiAwIDRweCAxNHB4IC01cHggY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXN0YXR1cy1jcml0aWNhbCkgNTUlLCB0cmFuc3BhcmVudCk7Cn0KLmJ0bi5kYW5nZXI6aG92ZXIgeyBmaWx0ZXI6IGJyaWdodG5lc3MoMS4wNik7IGJveC1zaGFkb3c6IDAgOHB4IDIycHggLTZweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc3RhdHVzLWNyaXRpY2FsKSA2MCUsIHRyYW5zcGFyZW50KTsgfQouYnRuLnN1Y2Nlc3MgewogIGJhY2tncm91bmQ6IHZhcigtLXN0YXR1cy1nb29kKTsgY29sb3I6ICNmZmY7IGJvcmRlci1jb2xvcjogdHJhbnNwYXJlbnQ7CiAgYm94LXNoYWRvdzogMCA0cHggMTRweCAtNXB4IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zdGF0dXMtZ29vZCkgNTUlLCB0cmFuc3BhcmVudCk7Cn0KLmJ0bi5zdWNjZXNzOmhvdmVyIHsgZmlsdGVyOiBicmlnaHRuZXNzKDEuMDYpOyBib3gtc2hhZG93OiAwIDhweCAyMnB4IC02cHggY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXN0YXR1cy1nb29kKSA2MCUsIHRyYW5zcGFyZW50KTsgfQouYnRuOmRpc2FibGVkIHsgb3BhY2l0eTogMC40NTsgY3Vyc29yOiBub3QtYWxsb3dlZDsgdHJhbnNmb3JtOiBub25lOyBib3gtc2hhZG93OiBub25lOyBmaWx0ZXI6IG5vbmU7IH0KLmJ0bi1yb3cgeyBkaXNwbGF5OiBmbGV4OyBnYXA6IDhweDsgZmxleC13cmFwOiB3cmFwOyB9CgovKiAtLS0tLS0tLS0tIFVwbG9hZCAtLS0tLS0tLS0tICovCi5kcm9wem9uZSB7CiAgYm9yZGVyOiAycHggZGFzaGVkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1sZyk7IHBhZGRpbmc6IDQwcHggMjBweDsKICB0ZXh0LWFsaWduOiBjZW50ZXI7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBjdXJzb3I6IHBvaW50ZXI7IHRyYW5zaXRpb246IGJvcmRlci1jb2xvciAyMDBtcyB2YXIoLS1lYXNlKSwgYmFja2dyb3VuZCAyMDBtcyB2YXIoLS1lYXNlKSwgdHJhbnNmb3JtIDIwMG1zIHZhcigtLWVhc2UpOwp9Ci5kcm9wem9uZTpob3ZlciB7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgtMXB4KTsgfQouZHJvcHpvbmUuaXMtZHJhZyB7IGJvcmRlci1jb2xvcjogdmFyKC0tc2VyaWVzLTEpOyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDYlLCB2YXIoLS1zdXJmYWNlLTIpKTsgdHJhbnNmb3JtOiBzY2FsZSgxLjAwNSk7IH0KLmRyb3B6b25lIGgzIHsgbWFyZ2luOiAwIDAgNnB4OyBmb250LXNpemU6IDE1cHg7IH0KLmRyb3B6b25lIHAgeyBtYXJnaW46IDA7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IGZvbnQtc2l6ZTogMTNweDsgfQouZHJvcHpvbmUgaW5wdXRbdHlwZT0iZmlsZSJdIHsgZGlzcGxheTogbm9uZTsgfQoKLmNvbmZsaWN0LWxpc3QgeyBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBnYXA6IDhweDsgbWFyZ2luOiAxMnB4IDA7IH0KLmNvbmZsaWN0LXJvdyB7CiAgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsganVzdGlmeS1jb250ZW50OiBzcGFjZS1iZXR3ZWVuOyBnYXA6IDEycHg7CiAgcGFkZGluZzogMTFweCAxNHB4OyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtc20pOyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTIpOwogIHRyYW5zaXRpb246IGJveC1zaGFkb3cgMTgwbXMgdmFyKC0tZWFzZSk7Cn0KLmNvbmZsaWN0LXJvdzpob3ZlciB7IGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdy1jYXJkKTsgfQouY29uZmxpY3Qtcm93IC53ZWVrLWxhYmVsIHsgZm9udC13ZWlnaHQ6IDYwMDsgZm9udC1zaXplOiAxM3B4OyB9Ci5jb25mbGljdC1yb3cgLndlZWstbWV0YSB7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgfQouY29uZmxpY3Qtcm93IHNlbGVjdCB7IG1pbi13aWR0aDogMDsgfQoKLmJhZGdlIHsgZGlzcGxheTogaW5saW5lLWJsb2NrOyBwYWRkaW5nOiAzcHggMTBweDsgYm9yZGVyLXJhZGl1czogMjBweDsgZm9udC1zaXplOiAxMXB4OyBmb250LXdlaWdodDogNzAwOyB9Ci5iYWRnZS5zdWNjZXNzIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXN0YXR1cy1nb29kKSAxOCUsIHRyYW5zcGFyZW50KTsgY29sb3I6IHZhcigtLXN0YXR1cy1nb29kKTsgfQouYmFkZ2UucGFydGlhbCB7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zdGF0dXMtd2FybmluZykgMjUlLCB0cmFuc3BhcmVudCk7IGNvbG9yOiAjOGE2MzAwOyB9Ci5iYWRnZS5mYWlsZWQgeyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc3RhdHVzLWNyaXRpY2FsKSAxOCUsIHRyYW5zcGFyZW50KTsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IH0KLmJhZGdlLmVycm9yLXNldiB7IGNvbG9yOiB2YXIoLS1zdGF0dXMtY3JpdGljYWwpOyB9Ci5iYWRnZS53YXJuaW5nLXNldiB7IGNvbG9yOiAjOGE2MzAwOyB9Ci5iYWRnZS5za2lwLXNldiB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgfQoKLmlzc3Vlcy1saXN0IHsgbWF4LWhlaWdodDogMjIwcHg7IG92ZXJmbG93LXk6IGF1dG87IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1zbSk7IH0KLmlzc3VlLXJvdyB7IHBhZGRpbmc6IDlweCAxNHB4OyBib3JkZXItYm90dG9tOiAxcHggc29saWQgdmFyKC0tZ3JpZGxpbmUpOyBmb250LXNpemU6IDEycHg7IH0KLmlzc3VlLXJvdzpsYXN0LWNoaWxkIHsgYm9yZGVyLWJvdHRvbTogbm9uZTsgfQouaXNzdWUtcm93IC5yb3ctbm8geyBjb2xvcjogdmFyKC0tdGV4dC1tdXRlZCk7IG1hcmdpbi1yaWdodDogNnB4OyB9CgovKiAtLS0tLS0tLS0tIFRvYXN0IC0tLS0tLS0tLS0gKi8KLnRvYXN0LXJvb3QgeyBwb3NpdGlvbjogZml4ZWQ7IGJvdHRvbTogMjBweDsgcmlnaHQ6IDIwcHg7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGdhcDogOHB4OyB6LWluZGV4OiAxMDA7IH0KLnRvYXN0IHsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTEpOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtc20pOwogIGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBwYWRkaW5nOiAxMnB4IDE2cHg7IGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdy1tb2RhbCk7IGZvbnQtc2l6ZTogMTNweDsgbWF4LXdpZHRoOiAzNDBweDsKICBhbmltYXRpb246IHRvYXN0LWluIDIyMG1zIHZhcigtLWVhc2UpOwp9Ci50b2FzdC5zdWNjZXNzIHsgYm9yZGVyLWxlZnQ6IDNweCBzb2xpZCB2YXIoLS1zdGF0dXMtZ29vZCk7IH0KLnRvYXN0LmVycm9yIHsgYm9yZGVyLWxlZnQ6IDNweCBzb2xpZCB2YXIoLS1zdGF0dXMtY3JpdGljYWwpOyB9CkBrZXlmcmFtZXMgdG9hc3QtaW4geyBmcm9tIHsgb3BhY2l0eTogMDsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDEwcHgpIHNjYWxlKDAuOTgpOyB9IHRvIHsgb3BhY2l0eTogMTsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDApIHNjYWxlKDEpOyB9IH0KCi8qIC0tLS0tLS0tLS0gTWlzYyAtLS0tLS0tLS0tICovCi5tdXRlZCB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgfQouZW1wdHktc3RhdGUgewogIHBhZGRpbmc6IDU2cHggMjRweDsgdGV4dC1hbGlnbjogY2VudGVyOyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOwogIGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogMTJweDsKICBhbmltYXRpb246IGNhcmRJbiAyNjBtcyB2YXIoLS1lYXNlKTsKfQouZW1wdHktc3RhdGUgLmVtcHR5LWljb24gewogIHdpZHRoOiA1MnB4OyBoZWlnaHQ6IDUycHg7IGJvcmRlci1yYWRpdXM6IDE2cHg7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogY2VudGVyOwogIGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zZXJpZXMtMSkgMTAlLCB0cmFuc3BhcmVudCk7IGNvbG9yOiB2YXIoLS1zZXJpZXMtMSk7Cn0KLmVtcHR5LXN0YXRlIC5lbXB0eS10aXRsZSB7IGZvbnQtc2l6ZTogMTRweDsgZm9udC13ZWlnaHQ6IDYwMDsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IH0KLmVtcHR5LXN0YXRlIC5lbXB0eS1tZXNzYWdlIHsgZm9udC1zaXplOiAxM3B4OyBtYXgtd2lkdGg6IDM2MHB4OyB9Ci5zcGlubmVyIHsgd2lkdGg6IDE2cHg7IGhlaWdodDogMTZweDsgYm9yZGVyLXJhZGl1czogNTAlOyBib3JkZXI6IDJweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItdG9wLWNvbG9yOiB2YXIoLS1zZXJpZXMtMSk7IGFuaW1hdGlvbjogc3BpbiAuNnMgbGluZWFyIGluZmluaXRlOyBkaXNwbGF5OiBpbmxpbmUtYmxvY2s7IH0KQGtleWZyYW1lcyBzcGluIHsgdG8geyB0cmFuc2Zvcm06IHJvdGF0ZSgzNjBkZWcpOyB9IH0KLmxvYWRpbmctcm93IHsgcGFkZGluZzogNDBweCAyMHB4OyB0ZXh0LWFsaWduOiBjZW50ZXI7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IH0KCi8qIFNrZWxldG9uIGxvYWRlcnMg4oCUIHNoaW1tZXJpbmcgcGxhY2Vob2xkZXJzIHNob3duIHdoaWxlIGEgc2VjdGlvbidzIGRhdGEgaXMgaW4gZmxpZ2h0ICovCi5za2VsZXRvbiB7CiAgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLXNtKTsKICBiYWNrZ3JvdW5kOiBsaW5lYXItZ3JhZGllbnQoMTAwZGVnLCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tdGV4dC1tdXRlZCkgMTIlLCB0cmFuc3BhcmVudCkgMzAlLCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tdGV4dC1tdXRlZCkgMjIlLCB0cmFuc3BhcmVudCkgNTAlLCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tdGV4dC1tdXRlZCkgMTIlLCB0cmFuc3BhcmVudCkgNzAlKTsKICBiYWNrZ3JvdW5kLXNpemU6IDIwMCUgMTAwJTsKICBhbmltYXRpb246IHNrZWxldG9uU2hpbW1lciAxLjRzIGVhc2UtaW4tb3V0IGluZmluaXRlOwp9CkBrZXlmcmFtZXMgc2tlbGV0b25TaGltbWVyIHsgZnJvbSB7IGJhY2tncm91bmQtcG9zaXRpb246IDE1MCUgMDsgfSB0byB7IGJhY2tncm91bmQtcG9zaXRpb246IC01MCUgMDsgfSB9Ci5za2VsZXRvbi1zdGF0LWdyaWQgeyBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdChhdXRvLWZpdCwgbWlubWF4KDE4MHB4LCAxZnIpKTsgZ2FwOiAxNHB4OyB9Ci5za2VsZXRvbi10aWxlIHsgaGVpZ2h0OiA4NHB4OyB9Ci5za2VsZXRvbi1jaGFydCB7IGhlaWdodDogMjgwcHg7IHdpZHRoOiAxMDAlOyB9Ci5za2VsZXRvbi1yb3cgeyBoZWlnaHQ6IDQwcHg7IG1hcmdpbi1ib3R0b206IDhweDsgfQoKLnR3by1jb2wgeyBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IDFmciAxZnI7IGdhcDogMTZweDsgfQpAbWVkaWEgKG1heC13aWR0aDogOTAwcHgpIHsgLnR3by1jb2wgeyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IDFmcjsgfSB9CgoubW9kZS10YWJzIHsgZGlzcGxheTogZmxleDsgZ2FwOiA2cHg7IGZsZXgtd3JhcDogd3JhcDsgbWFyZ2luLWJvdHRvbTogMTZweDsgfQoubW9kZS10YWJzIGJ1dHRvbiB7CiAgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgcGFkZGluZzogN3B4IDE0cHg7IGJvcmRlci1yYWRpdXM6IDIwcHg7IGZvbnQtc2l6ZTogMTJweDsgZm9udC13ZWlnaHQ6IDYwMDsgY3Vyc29yOiBwb2ludGVyOwogIHRyYW5zaXRpb246IGNvbG9yIDE4MG1zIHZhcigtLWVhc2UpLCBiYWNrZ3JvdW5kIDE4MG1zIHZhcigtLWVhc2UpLCB0cmFuc2Zvcm0gMTUwbXMgdmFyKC0tZWFzZSk7Cn0KLm1vZGUtdGFicyBidXR0b246aG92ZXIgeyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoLTFweCk7IH0KLm1vZGUtdGFicyBidXR0b24uaXMtYWN0aXZlIHsgYmFja2dyb3VuZDogdmFyKC0tc2VyaWVzLTEpOyBjb2xvcjogI2ZmZjsgYm9yZGVyLWNvbG9yOiB0cmFuc3BhcmVudDsgYm94LXNoYWRvdzogMCA0cHggMTRweCAtNXB4IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zZXJpZXMtMSkgNjAlLCB0cmFuc3BhcmVudCk7IH0KCi5maWVsZC1pbmxpbmUgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDhweDsgZm9udC1zaXplOiAxMnB4OyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyB9Ci5maWVsZC1pbmxpbmUgc2VsZWN0LCAuZmllbGQtaW5saW5lIGlucHV0IHsgbWluLXdpZHRoOiAwOyBwYWRkaW5nOiA2cHggMTBweDsgfQoKLyogLS0tLS0tLS0tLSBQYWdpbmF0aW9uIC0tLS0tLS0tLS0gKi8KLnBhZ2luYXRpb24tcm93IHsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiAxMnB4OyBtYXJnaW4tdG9wOiAxNHB4OyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IH0KLnBhZ2luYXRpb24tcm93IC5idG4geyBwYWRkaW5nOiA2cHggMTJweDsgfQoKLyogLS0tLS0tLS0tLSBEYXNoYm9hcmQgY29udHJvbHMgLyBtZXRyaWMtZm9jdXNlZCBLUElzIC0tLS0tLS0tLS0gKi8KLmRhc2hib2FyZC1jb250cm9scyB7CiAgZGlzcGxheTogZmxleDsgZmxleC13cmFwOiB3cmFwOyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDEwcHg7IG1hcmdpbi1ib3R0b206IDE4cHg7Cn0KLmRhc2hib2FyZC1jb250cm9scyBsYWJlbCB7IGZvbnQtc2l6ZTogMTJweDsgZm9udC13ZWlnaHQ6IDYwMDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgbWFyZ2luLXJpZ2h0OiA2cHg7IH0KLmRhc2hib2FyZC1jb250cm9scyBzZWxlY3QgeyBmb250LXdlaWdodDogNjAwOyB9Ci5zdGF0LXRpbGUucG9zdC10aWxlIC5zdGF0LXZhbHVlIHsgZm9udC1zaXplOiAxNXB4OyBsaW5lLWhlaWdodDogMS4zNTsgbWFyZ2luLXRvcDogNnB4OyB9Ci5zdGF0LXRpbGUucG9zdC10aWxlIC5wb3N0LW1ldGEgeyBmb250LXNpemU6IDExcHg7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgbWFyZ2luLXRvcDogNHB4OyB9Ci5zdGF0LXRpbGUucG9zdC10aWxlIC5wb3N0LW1ldHJpYy12YWx1ZSB7IGZvbnQtc2l6ZTogMjBweDsgZm9udC13ZWlnaHQ6IDcwMDsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IG1hcmdpbi10b3A6IDZweDsgfQoKLyogLS0tLS0tLS0tLSBEYXRhIFJlY29yZHMgKHBsYXRmb3JtLWdyb3VwZWQpIC0tLS0tLS0tLS0gKi8KLnJlY29yZHMtdG9vbGJhciB7CiAgZGlzcGxheTogZmxleDsgZmxleC13cmFwOiB3cmFwOyBhbGlnbi1pdGVtczogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6IHNwYWNlLWJldHdlZW47CiAgZ2FwOiAxMnB4OyBtYXJnaW4tYm90dG9tOiAxNHB4Owp9Ci5wbGF0Zm9ybS1maWx0ZXItcGlsbHMgeyBkaXNwbGF5OiBmbGV4OyBmbGV4LXdyYXA6IHdyYXA7IGdhcDogNnB4OyB9Ci5wbGF0Zm9ybS1maWx0ZXItcGlsbHMgYnV0dG9uIHsKICBkaXNwbGF5OiBpbmxpbmUtZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA2cHg7CiAgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgcGFkZGluZzogN3B4IDE0cHg7IGJvcmRlci1yYWRpdXM6IDIwcHg7IGZvbnQtc2l6ZTogMTJweDsgZm9udC13ZWlnaHQ6IDYwMDsgY3Vyc29yOiBwb2ludGVyOwogIHRyYW5zaXRpb246IGNvbG9yIDE4MG1zIHZhcigtLWVhc2UpLCBiYWNrZ3JvdW5kIDE4MG1zIHZhcigtLWVhc2UpLCB0cmFuc2Zvcm0gMTUwbXMgdmFyKC0tZWFzZSksIGJveC1zaGFkb3cgMTgwbXMgdmFyKC0tZWFzZSk7Cn0KLnBsYXRmb3JtLWZpbHRlci1waWxscyBidXR0b246aG92ZXIgeyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKC0xcHgpOyB9Ci5wbGF0Zm9ybS1maWx0ZXItcGlsbHMgYnV0dG9uOmFjdGl2ZSB7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgwKSBzY2FsZSgwLjk2KTsgfQoucGxhdGZvcm0tZmlsdGVyLXBpbGxzIGJ1dHRvbi5pcy1hY3RpdmUgeyBiYWNrZ3JvdW5kOiB2YXIoLS1zZXJpZXMtMSk7IGNvbG9yOiAjZmZmOyBib3JkZXItY29sb3I6IHRyYW5zcGFyZW50OyBib3gtc2hhZG93OiAwIDRweCAxNHB4IC01cHggY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSA2MCUsIHRyYW5zcGFyZW50KTsgfQoucGxhdGZvcm0tZmlsdGVyLXBpbGxzIGJ1dHRvbi5pcy1hY3RpdmUgLnBsYXRmb3JtLWRvdCB7IGJveC1zaGFkb3c6IDAgMCAwIDJweCByZ2JhKDI1NSwyNTUsMjU1LDAuNSk7IH0KLnJlY29yZHMtc2VhcmNoIGlucHV0IHsgYm9yZGVyLXJhZGl1czogMjBweDsgbWluLXdpZHRoOiAyMjBweDsgfQouc3RhdHVzLXBpbGwgeyBkaXNwbGF5OiBpbmxpbmUtYmxvY2s7IHBhZGRpbmc6IDNweCAxMHB4OyBib3JkZXItcmFkaXVzOiAyMHB4OyBmb250LXNpemU6IDExcHg7IGZvbnQtd2VpZ2h0OiA3MDA7IH0KLnN0YXR1cy1waWxsLm9yaWdpbmFsIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXRleHQtbXV0ZWQpIDE1JSwgdHJhbnNwYXJlbnQpOyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyB9Ci5zdGF0dXMtcGlsbC5lZGl0ZWQgeyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc3RhdHVzLXdhcm5pbmcpIDIyJSwgdHJhbnNwYXJlbnQpOyBjb2xvcjogIzhhNjMwMDsgfQoucm93LWFjdGlvbnMgeyBkaXNwbGF5OiBmbGV4OyBnYXA6IDZweDsgZmxleC13cmFwOiBub3dyYXA7IH0KLnJvdy1hY3Rpb25zIC5idG4geyBwYWRkaW5nOiA1cHggMTBweDsgZm9udC1zaXplOiAxMnB4OyB9Ci5saW5rLWNlbGwgYSB7IGNvbG9yOiB2YXIoLS1zZXJpZXMtMSk7IHRleHQtZGVjb3JhdGlvbjogbm9uZTsgZm9udC13ZWlnaHQ6IDYwMDsgZm9udC1zaXplOiAxMnB4OyB9Ci5saW5rLWNlbGwgYTpob3ZlciB7IHRleHQtZGVjb3JhdGlvbjogdW5kZXJsaW5lOyB9Ci5yZWNvcmQtc2VjdGlvbiB7CiAgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLXNtKTsgcGFkZGluZzogMTZweDsgbWFyZ2luLWJvdHRvbTogMTRweDsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTIpOyBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7Cn0KLnJlY29yZC1zZWN0aW9uIGg0IHsgbWFyZ2luOiAwIDAgMTJweDsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNzAwOyBsZXR0ZXItc3BhY2luZzogMC4wM2VtOyB0ZXh0LXRyYW5zZm9ybTogdXBwZXJjYXNlOyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyB9Ci5yZWNvcmQtc2VjdGlvbiAuZm9ybS1ncmlkIHsgbWFyZ2luLWJvdHRvbTogMDsgfQoucmVjb3JkLXNlY3Rpb24gLnZpZXctZmllbGQgeyBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBnYXA6IDJweDsgZm9udC1zaXplOiAxM3B4OyB9Ci5yZWNvcmQtc2VjdGlvbiAudmlldy1maWVsZCAudmlldy1sYWJlbCB7IGZvbnQtc2l6ZTogMTFweDsgZm9udC13ZWlnaHQ6IDYwMDsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyB9Ci5yZWNvcmQtc2VjdGlvbiAudmlldy1maWVsZCAudmlldy12YWx1ZSB7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyB3b3JkLWJyZWFrOiBicmVhay13b3JkOyB9CkBtZWRpYSAobWF4LXdpZHRoOiA2NDBweCkgewogIC5yZWNvcmRzLXRvb2xiYXIgeyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBhbGlnbi1pdGVtczogc3RyZXRjaDsgfQogIC5yZWNvcmRzLXNlYXJjaCBpbnB1dCB7IHdpZHRoOiAxMDAlOyB9Cn0KCi8qIC0tLS0tLS0tLS0gTW9kYWwgKHJlY29yZCBlZGl0b3IpIC0tLS0tLS0tLS0gKi8KLm1vZGFsLW92ZXJsYXkgewogIHBvc2l0aW9uOiBmaXhlZDsgaW5zZXQ6IDA7IGJhY2tncm91bmQ6IHJnYmEoMTAsMTEsMTMsMC41KTsKICBiYWNrZHJvcC1maWx0ZXI6IGJsdXIoNnB4KTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IGJsdXIoNnB4KTsKICBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogZmxleC1zdGFydDsganVzdGlmeS1jb250ZW50OiBjZW50ZXI7CiAgcGFkZGluZzogNDBweCAxNnB4OyBvdmVyZmxvdy15OiBhdXRvOyB6LWluZGV4OiAyMDA7CiAgYW5pbWF0aW9uOiBvdmVybGF5SW4gMjAwbXMgdmFyKC0tZWFzZSk7Cn0KQGtleWZyYW1lcyBvdmVybGF5SW4geyBmcm9tIHsgb3BhY2l0eTogMDsgfSB0byB7IG9wYWNpdHk6IDE7IH0gfQpAa2V5ZnJhbWVzIG1vZGFsUGFuZWxJbiB7CiAgZnJvbSB7IG9wYWNpdHk6IDA7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgxNHB4KSBzY2FsZSgwLjk3KTsgfQogIHRvIHsgb3BhY2l0eTogMTsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDApIHNjYWxlKDEpOyB9Cn0KLm1vZGFsLXBhbmVsIHsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTEpOyBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtbGcpOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOwogIGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBwYWRkaW5nOiAyNHB4OyB3aWR0aDogMTAwJTsgbWF4LXdpZHRoOiA3MjBweDsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LW1vZGFsKTsKICBtYXgtaGVpZ2h0OiBjYWxjKDEwMHZoIC0gODBweCk7IG92ZXJmbG93LXk6IGF1dG87CiAgYW5pbWF0aW9uOiBtb2RhbFBhbmVsSW4gMjQwbXMgdmFyKC0tZWFzZSk7Cn0KLm1vZGFsLXBhbmVsLndpZGUgeyBtYXgtd2lkdGg6IDExMDBweDsgfQoubW9kYWwtcGFuZWwgaDIgeyBtYXJnaW46IDAgMCA0cHg7IGZvbnQtc2l6ZTogMTdweDsgbGV0dGVyLXNwYWNpbmc6IC0wLjAxZW07IH0KLm1vZGFsLXBhbmVsIC5tb2RhbC1zdWIgeyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyBmb250LXNpemU6IDEycHg7IG1hcmdpbjogMCAwIDE4cHg7IH0KLmZvcm0tZ3JpZCB7IGRpc3BsYXk6IGdyaWQ7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogcmVwZWF0KGF1dG8tZml0LCBtaW5tYXgoMjAwcHgsIDFmcikpOyBnYXA6IDEycHg7IG1hcmdpbi1ib3R0b206IDE2cHg7IH0KLmZvcm0tZ3JpZC5mdWxsIHsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiAxZnI7IH0KQG1lZGlhIChtYXgtd2lkdGg6IDY0MHB4KSB7IC5mb3JtLWdyaWQgeyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IDFmcjsgfSB9Ci5mb3JtLWZpZWxkIHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiA1cHg7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgfQouZm9ybS1maWVsZCBsYWJlbCB7IGZvbnQtd2VpZ2h0OiA2MDA7IH0KLmZvcm0tZmllbGQgdGV4dGFyZWEgeyByZXNpemU6IHZlcnRpY2FsOyBtaW4taGVpZ2h0OiA2MHB4OyB9CgoucGxhdGZvcm0tZWRpdC1yb3cgewogIGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1zbSk7IHBhZGRpbmc6IDE0cHg7IG1hcmdpbi1ib3R0b206IDEwcHg7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMik7Cn0KLnBsYXRmb3JtLWVkaXQtcm93IC5wbGF0Zm9ybS1lZGl0LWhlYWQgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6IHNwYWNlLWJldHdlZW47IGdhcDogOHB4OyBtYXJnaW4tYm90dG9tOiAxMHB4OyB9Ci5wbGF0Zm9ybS1lZGl0LXJvdyAubWV0cmljcy1ncmlkIHsgZGlzcGxheTogZ3JpZDsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiByZXBlYXQoYXV0by1maXQsIG1pbm1heCgxMjBweCwgMWZyKSk7IGdhcDogOHB4OyB9Ci5yZW1vdmUtcGxhdGZvcm0tYnRuIHsgYm9yZGVyOiBub25lOyBiYWNrZ3JvdW5kOiB0cmFuc3BhcmVudDsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IGN1cnNvcjogcG9pbnRlcjsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNjAwOyB0cmFuc2l0aW9uOiBvcGFjaXR5IDE1MG1zIHZhcigtLWVhc2UpOyB9Ci5yZW1vdmUtcGxhdGZvcm0tYnRuOmhvdmVyIHsgb3BhY2l0eTogMC43OyB9Ci5tb2RhbC1hY3Rpb25zIHsgZGlzcGxheTogZmxleDsganVzdGlmeS1jb250ZW50OiBzcGFjZS1iZXR3ZWVuOyBhbGlnbi1pdGVtczogY2VudGVyOyBtYXJnaW4tdG9wOiAxOHB4OyBnYXA6IDhweDsgZmxleC13cmFwOiB3cmFwOyB9CgovKiAtLS0tLS0tLS0tIFJlc3BvbnNpdmUgdGlnaHRlbmluZyAtLS0tLS0tLS0tICovCkBtZWRpYSAobWF4LXdpZHRoOiA3MjBweCkgewogIC50b3BiYXIgeyBnYXA6IDEycHg7IHBhZGRpbmc6IDEwcHggMTRweDsgZmxleC13cmFwOiB3cmFwOyB9CiAgLnRvcGJhci1icmFuZCB7IG9yZGVyOiAxOyB9CiAgLnRvcGJhci11c2VyIHsgb3JkZXI6IDI7IG1hcmdpbi1sZWZ0OiBhdXRvOyB9CiAgLnRhYnMgeyBvcmRlcjogMzsgd2lkdGg6IDEwMCU7IH0KICAudmlldy1hcmVhIHsgcGFkZGluZzogMTRweDsgfQogIC5maWx0ZXItYmFyIHsgdG9wOiBhdXRvOyBwb3NpdGlvbjogc3RhdGljOyBwYWRkaW5nOiAxMnB4IDE0cHg7IH0KICAuc3RhdC1ncmlkIHsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiByZXBlYXQoYXV0by1maXQsIG1pbm1heCgxNDBweCwgMWZyKSk7IH0KfQo8L3N0eWxlPgo8L2hlYWQ+Cjxib2R5Pgo8ZGl2IGNsYXNzPSJhdXRoLXNjcmVlbiIgaWQ9ImF1dGhTY3JlZW4iPgogIDxkaXYgY2xhc3M9ImF1dGgtY2FyZCI+CiAgICA8ZGl2IGNsYXNzPSJhdXRoLWJyYW5kIj4KICAgICAgPHNwYW4gY2xhc3M9ImJyYW5kLW1hcmsiPkxSUzwvc3Bhbj4KICAgICAgPHNwYW4gY2xhc3M9ImJyYW5kLXRpdGxlIj5Tb2NpYWwgTWVkaWEgQW5hbHl0aWNzPC9zcGFuPgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJhdXRoLWZvcm0iPgogICAgICA8ZGl2IGNsYXNzPSJmb3JtLWZpZWxkIj4KICAgICAgICA8bGFiZWwgZm9yPSJhdXRoQ29kZSI+QWNjZXNzIGNvZGU8L2xhYmVsPgogICAgICAgIDxpbnB1dCB0eXBlPSJwYXNzd29yZCIgaWQ9ImF1dGhDb2RlIiBhdXRvY29tcGxldGU9Im9mZiIgLz4KICAgICAgPC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImF1dGgtZXJyb3IiIGlkPSJhdXRoRXJyb3IiPjwvZGl2PgogICAgICA8YnV0dG9uIGNsYXNzPSJidG4gcHJpbWFyeSIgaWQ9ImF1dGhTdWJtaXRCdG4iIHR5cGU9ImJ1dHRvbiI+PGkgZGF0YS1sdWNpZGU9ImFycm93LXJpZ2h0IiBzdHlsZT0id2lkdGg6MTRweDtoZWlnaHQ6MTRweDsiPjwvaT4gRW50ZXI8L2J1dHRvbj4KICAgIDwvZGl2PgogIDwvZGl2Pgo8L2Rpdj4KCjxkaXYgY2xhc3M9ImFwcC1zaGVsbCIgaWQ9ImFwcFNoZWxsIiBzdHlsZT0iZGlzcGxheTpub25lOyI+CiAgPGhlYWRlciBjbGFzcz0idG9wYmFyIj4KICAgIDxkaXYgY2xhc3M9InRvcGJhci1icmFuZCI+CiAgICAgIDxzcGFuIGNsYXNzPSJicmFuZC1tYXJrIj5MUlM8L3NwYW4+CiAgICAgIDxzcGFuIGNsYXNzPSJicmFuZC10aXRsZSI+U29jaWFsIE1lZGlhIEFuYWx5dGljczwvc3Bhbj4KICAgIDwvZGl2PgogICAgPG5hdiBjbGFzcz0idGFicyIgcm9sZT0idGFibGlzdCIgYXJpYS1sYWJlbD0iU2VjdGlvbnMiPgogICAgICA8YnV0dG9uIGNsYXNzPSJ0YWItYnRuIGlzLWFjdGl2ZSIgZGF0YS10YWI9ImRhc2hib2FyZCIgcm9sZT0idGFiIiBhcmlhLXNlbGVjdGVkPSJ0cnVlIj48aSBkYXRhLWx1Y2lkZT0ibGF5b3V0LWRhc2hib2FyZCIgc3R5bGU9IndpZHRoOjE0cHg7aGVpZ2h0OjE0cHg7Ij48L2k+IERhc2hib2FyZDwvYnV0dG9uPgogICAgICA8YnV0dG9uIGNsYXNzPSJ0YWItYnRuIiBkYXRhLXRhYj0icmVjb3JkcyIgcm9sZT0idGFiIiBhcmlhLXNlbGVjdGVkPSJmYWxzZSI+PGkgZGF0YS1sdWNpZGU9ImRhdGFiYXNlIiBzdHlsZT0id2lkdGg6MTRweDtoZWlnaHQ6MTRweDsiPjwvaT4gRGF0YSBSZWNvcmRzPC9idXR0b24+CiAgICAgIDxidXR0b24gY2xhc3M9InRhYi1idG4iIGRhdGEtdGFiPSJjb21wYXJpc29uIiByb2xlPSJ0YWIiIGFyaWEtc2VsZWN0ZWQ9ImZhbHNlIj48aSBkYXRhLWx1Y2lkZT0iZ2l0LWNvbXBhcmUiIHN0eWxlPSJ3aWR0aDoxNHB4O2hlaWdodDoxNHB4OyI+PC9pPiBDb21wYXJpc29uczwvYnV0dG9uPgogICAgICA8YnV0dG9uIGNsYXNzPSJ0YWItYnRuIiBkYXRhLXRhYj0idXBsb2FkIiByb2xlPSJ0YWIiIGFyaWEtc2VsZWN0ZWQ9ImZhbHNlIj48aSBkYXRhLWx1Y2lkZT0idXBsb2FkLWNsb3VkIiBzdHlsZT0id2lkdGg6MTRweDtoZWlnaHQ6MTRweDsiPjwvaT4gVXBsb2FkIERhdGE8L2J1dHRvbj4KICAgICAgPGJ1dHRvbiBjbGFzcz0idGFiLWJ0biIgZGF0YS10YWI9Imhpc3RvcnkiIHJvbGU9InRhYiIgYXJpYS1zZWxlY3RlZD0iZmFsc2UiPjxpIGRhdGEtbHVjaWRlPSJoaXN0b3J5IiBzdHlsZT0id2lkdGg6MTRweDtoZWlnaHQ6MTRweDsiPjwvaT4gVXBsb2FkIEhpc3Rvcnk8L2J1dHRvbj4KICAgIDwvbmF2PgogICAgPGRpdiBjbGFzcz0idG9wYmFyLXVzZXIiPgogICAgICA8YnV0dG9uIGNsYXNzPSJidG4iIGlkPSJsb2dvdXRCdG4iIHR5cGU9ImJ1dHRvbiI+PGkgZGF0YS1sdWNpZGU9ImxvY2siIHN0eWxlPSJ3aWR0aDoxNHB4O2hlaWdodDoxNHB4OyI+PC9pPiBMb2NrPC9idXR0b24+CiAgICAgIDxidXR0b24gY2xhc3M9InRoZW1lLXRvZ2dsZSIgaWQ9InRoZW1lVG9nZ2xlIiB0eXBlPSJidXR0b24iIGFyaWEtbGFiZWw9IlRvZ2dsZSBkYXJrIG1vZGUiPgogICAgICAgIDxzcGFuIGlkPSJ0aGVtZVRvZ2dsZUljb24iPjxpIGRhdGEtbHVjaWRlPSJtb29uIiBzdHlsZT0id2lkdGg6MTZweDtoZWlnaHQ6MTZweDsiPjwvaT48L3NwYW4+CiAgICAgIDwvYnV0dG9uPgogICAgPC9kaXY+CiAgPC9oZWFkZXI+CgogIDxzZWN0aW9uIGNsYXNzPSJmaWx0ZXItYmFyIiBpZD0iZmlsdGVyQmFyIj4KICAgIDxkaXYgY2xhc3M9ImZpbHRlci1maWVsZCI+CiAgICAgIDxsYWJlbCBmb3I9ImZpbHRlckRhdGVGcm9tIj5Gcm9tPC9sYWJlbD4KICAgICAgPGlucHV0IHR5cGU9ImRhdGUiIGlkPSJmaWx0ZXJEYXRlRnJvbSIgLz4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0iZmlsdGVyLWZpZWxkIj4KICAgICAgPGxhYmVsIGZvcj0iZmlsdGVyRGF0ZVRvIj5UbzwvbGFiZWw+CiAgICAgIDxpbnB1dCB0eXBlPSJkYXRlIiBpZD0iZmlsdGVyRGF0ZVRvIiAvPgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJmaWx0ZXItZmllbGQgZmlsdGVyLXByZXNldHMiIGlkPSJmaWx0ZXJQcmVzZXRzIj4KICAgICAgPGJ1dHRvbiB0eXBlPSJidXR0b24iIGRhdGEtcHJlc2V0PSI3Ij5MYXN0IDcgZGF5czwvYnV0dG9uPgogICAgICA8YnV0dG9uIHR5cGU9ImJ1dHRvbiIgZGF0YS1wcmVzZXQ9IjMwIj5MYXN0IDMwIGRheXM8L2J1dHRvbj4KICAgICAgPGJ1dHRvbiB0eXBlPSJidXR0b24iIGRhdGEtcHJlc2V0PSI5MCI+TGFzdCA5MCBkYXlzPC9idXR0b24+CiAgICAgIDxidXR0b24gdHlwZT0iYnV0dG9uIiBkYXRhLXByZXNldD0iYWxsIj5BbGwgdGltZTwvYnV0dG9uPgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJmaWx0ZXItZmllbGQiPgogICAgICA8bGFiZWwgZm9yPSJmaWx0ZXJQbGF0Zm9ybSI+UGxhdGZvcm08L2xhYmVsPgogICAgICA8c2VsZWN0IGlkPSJmaWx0ZXJQbGF0Zm9ybSI+PG9wdGlvbiB2YWx1ZT0iYWxsIj5BbGwgcGxhdGZvcm1zPC9vcHRpb24+PC9zZWxlY3Q+CiAgICA8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImZpbHRlci1maWVsZCI+CiAgICAgIDxsYWJlbCBmb3I9ImZpbHRlckNhbXBhaWduIj5DYW1wYWlnbjwvbGFiZWw+CiAgICAgIDxzZWxlY3QgaWQ9ImZpbHRlckNhbXBhaWduIj48b3B0aW9uIHZhbHVlPSJhbGwiPkFsbCBjYW1wYWlnbnM8L29wdGlvbj48L3NlbGVjdD4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0iZmlsdGVyLWZpZWxkIj4KICAgICAgPGxhYmVsIGZvcj0iZmlsdGVyQ29udGVudFR5cGUiPkNvbnRlbnQgdHlwZTwvbGFiZWw+CiAgICAgIDxzZWxlY3QgaWQ9ImZpbHRlckNvbnRlbnRUeXBlIj48b3B0aW9uIHZhbHVlPSJhbGwiPkFsbCBjb250ZW50IHR5cGVzPC9vcHRpb24+PC9zZWxlY3Q+CiAgICA8L2Rpdj4KICA8L3NlY3Rpb24+CgogIDxtYWluIGNsYXNzPSJ2aWV3LWFyZWEiPgogICAgPHNlY3Rpb24gaWQ9InZpZXctZGFzaGJvYXJkIiBjbGFzcz0idmlldyBpcy1hY3RpdmUiPjwvc2VjdGlvbj4KICAgIDxzZWN0aW9uIGlkPSJ2aWV3LXJlY29yZHMiIGNsYXNzPSJ2aWV3Ij48L3NlY3Rpb24+CiAgICA8c2VjdGlvbiBpZD0idmlldy1jb21wYXJpc29uIiBjbGFzcz0idmlldyI+PC9zZWN0aW9uPgogICAgPHNlY3Rpb24gaWQ9InZpZXctdXBsb2FkIiBjbGFzcz0idmlldyI+PC9zZWN0aW9uPgogICAgPHNlY3Rpb24gaWQ9InZpZXctaGlzdG9yeSIgY2xhc3M9InZpZXciPjwvc2VjdGlvbj4KICA8L21haW4+CjwvZGl2PgoKPGRpdiBpZD0idG9hc3RSb290IiBjbGFzcz0idG9hc3Qtcm9vdCIgYXJpYS1saXZlPSJwb2xpdGUiPjwvZGl2PgoKPHNjcmlwdD4KLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIEFwaSDigJQgdGhpbiBmZXRjaCB3cmFwcGVycyBhcm91bmQgdGhlIFJFU1QgQVBJLgogICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KY29uc3QgQXBpID0gKCgpID0+IHsKICBhc3luYyBmdW5jdGlvbiByZXF1ZXN0KHBhdGgsIG9wdGlvbnMpIHsKICAgIGNvbnN0IHJlcyA9IGF3YWl0IGZldGNoKHBhdGgsIG9wdGlvbnMpOwogICAgbGV0IGJvZHk7CiAgICB0cnkgewogICAgICBib2R5ID0gYXdhaXQgcmVzLmpzb24oKTsKICAgIH0gY2F0Y2ggewogICAgICBib2R5ID0gbnVsbDsKICAgIH0KICAgIGlmIChyZXMuc3RhdHVzID09PSA0MDEgJiYgIXBhdGguc3RhcnRzV2l0aCgnL2FwaS9hdXRoLycpKSB7CiAgICAgIHdpbmRvdy5kaXNwYXRjaEV2ZW50KG5ldyBDdXN0b21FdmVudCgnbHJzOnNpZ25lZC1vdXQnKSk7CiAgICB9CiAgICBpZiAoIXJlcy5vaykgewogICAgICBjb25zdCBtZXNzYWdlID0gKGJvZHkgJiYgYm9keS5lcnJvcikgfHwgYFJlcXVlc3QgZmFpbGVkICgke3Jlcy5zdGF0dXN9KWA7CiAgICAgIHRocm93IG5ldyBFcnJvcihtZXNzYWdlKTsKICAgIH0KICAgIHJldHVybiBib2R5OwogIH0KCiAgZnVuY3Rpb24gcXMocGFyYW1zKSB7CiAgICBjb25zdCB1c3AgPSBuZXcgVVJMU2VhcmNoUGFyYW1zKCk7CiAgICBPYmplY3QuZW50cmllcyhwYXJhbXMgfHwge30pLmZvckVhY2goKFtrLCB2XSkgPT4gewogICAgICBpZiAodiAhPT0gdW5kZWZpbmVkICYmIHYgIT09IG51bGwgJiYgdiAhPT0gJycpIHVzcC5zZXQoaywgdik7CiAgICB9KTsKICAgIGNvbnN0IHMgPSB1c3AudG9TdHJpbmcoKTsKICAgIHJldHVybiBzID8gYD8ke3N9YCA6ICcnOwogIH0KCiAgcmV0dXJuIHsKICAgIGF1dGhNZTogKCkgPT4gcmVxdWVzdCgnL2FwaS9hdXRoL21lJyksCiAgICBhdXRoTG9naW46IChjb2RlKSA9PgogICAgICByZXF1ZXN0KCcvYXBpL2F1dGgvbG9naW4nLCB7IG1ldGhvZDogJ1BPU1QnLCBoZWFkZXJzOiB7ICdDb250ZW50LVR5cGUnOiAnYXBwbGljYXRpb24vanNvbicgfSwgYm9keTogSlNPTi5zdHJpbmdpZnkoeyBjb2RlIH0pIH0pLAogICAgYXV0aExvZ291dDogKCkgPT4gcmVxdWVzdCgnL2FwaS9hdXRoL2xvZ291dCcsIHsgbWV0aG9kOiAnUE9TVCcgfSksCgogICAgZmlsdGVyT3B0aW9uczogKCkgPT4gcmVxdWVzdCgnL2FwaS9hbmFseXRpY3MvZmlsdGVyLW9wdGlvbnMnKSwKICAgIGtwaXM6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL2twaXMke3FzKHBhcmFtcyl9YCksCiAgICBwbGF0Zm9ybUJyZWFrZG93bjogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9hbmFseXRpY3MvcGxhdGZvcm0tYnJlYWtkb3duJHtxcyhwYXJhbXMpfWApLAogICAgY2FtcGFpZ25CcmVha2Rvd246IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL2NhbXBhaWduLWJyZWFrZG93biR7cXMocGFyYW1zKX1gKSwKICAgIGNvbnRlbnRUeXBlQnJlYWtkb3duOiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9jb250ZW50LXR5cGUtYnJlYWtkb3duJHtxcyhwYXJhbXMpfWApLAogICAgbWV0cmljT3B0aW9uczogKHBsYXRmb3JtKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9tZXRyaWMtb3B0aW9ucyR7cXMoeyBwbGF0Zm9ybSB9KX1gKSwKICAgIG1ldHJpY1N1bW1hcnk6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL21ldHJpYy1zdW1tYXJ5JHtxcyhwYXJhbXMpfWApLAogICAgdHJlbmQ6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL3RyZW5kJHtxcyhwYXJhbXMpfWApLAogICAgdG9wUG9zdHM6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL3RvcC1wb3N0cyR7cXMocGFyYW1zKX1gKSwKICAgIGNvbXBhcmU6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL2NvbXBhcmUke3FzKHBhcmFtcyl9YCksCiAgICBtb250aGx5OiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9tb250aGx5JHtxcyhwYXJhbXMpfWApLAogICAgcXVhcnRlcmx5OiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9xdWFydGVybHkke3FzKHBhcmFtcyl9YCksCiAgICB5dGQ6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL3l0ZCR7cXMocGFyYW1zKX1gKSwKCiAgICBwcmV2aWV3VXBsb2FkOiAoZmlsZSkgPT4gewogICAgICBjb25zdCBmb3JtID0gbmV3IEZvcm1EYXRhKCk7CiAgICAgIGZvcm0uYXBwZW5kKCdmaWxlJywgZmlsZSk7CiAgICAgIHJldHVybiByZXF1ZXN0KCcvYXBpL3VwbG9hZHMvcHJldmlldycsIHsgbWV0aG9kOiAnUE9TVCcsIGJvZHk6IGZvcm0gfSk7CiAgICB9LAogICAgY29tbWl0VXBsb2FkOiAocGF5bG9hZCkgPT4KICAgICAgcmVxdWVzdCgnL2FwaS91cGxvYWRzL2NvbW1pdCcsIHsKICAgICAgICBtZXRob2Q6ICdQT1NUJywKICAgICAgICBoZWFkZXJzOiB7ICdDb250ZW50LVR5cGUnOiAnYXBwbGljYXRpb24vanNvbicgfSwKICAgICAgICBib2R5OiBKU09OLnN0cmluZ2lmeShwYXlsb2FkKSwKICAgICAgfSksCiAgICB1cGxvYWRIaXN0b3J5OiAoKSA9PiByZXF1ZXN0KCcvYXBpL3VwbG9hZHMvaGlzdG9yeScpLAogICAgdXBsb2FkRXJyb3JzOiAoaWQpID0+IHJlcXVlc3QoYC9hcGkvdXBsb2Fkcy8ke2lkfS9lcnJvcnNgKSwKICAgIHVwbG9hZFJhd1Jvd3M6IChpZCkgPT4gcmVxdWVzdChgL2FwaS91cGxvYWRzLyR7aWR9L3Jhdy1yb3dzYCksCgogICAgbGlzdFJlY29yZHM6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvcmVjb3JkcyR7cXMocGFyYW1zKX1gKSwKICAgIHJlY29yZHNUYWJsZTogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9yZWNvcmRzL3RhYmxlJHtxcyhwYXJhbXMpfWApLAogICAgZ2V0UmVjb3JkOiAoaWQpID0+IHJlcXVlc3QoYC9hcGkvcmVjb3Jkcy8ke2lkfWApLAogICAgdXBkYXRlUmVjb3JkOiAoaWQsIHZhbHVlcykgPT4KICAgICAgcmVxdWVzdChgL2FwaS9yZWNvcmRzLyR7aWR9YCwgewogICAgICAgIG1ldGhvZDogJ1BVVCcsCiAgICAgICAgaGVhZGVyczogeyAnQ29udGVudC1UeXBlJzogJ2FwcGxpY2F0aW9uL2pzb24nIH0sCiAgICAgICAgYm9keTogSlNPTi5zdHJpbmdpZnkoeyB2YWx1ZXMgfSksCiAgICAgIH0pLAogICAgZGVsZXRlUmVjb3JkUG9zdDogKHBvc3RJZCkgPT4gcmVxdWVzdChgL2FwaS9yZWNvcmRzL3Bvc3QvJHtwb3N0SWR9YCwgeyBtZXRob2Q6ICdERUxFVEUnIH0pLAogICAgZGVsZXRlUmVjb3JkUGxhdGZvcm06IChwb3N0SWQsIHBsYXRmb3JtKSA9PgogICAgICByZXF1ZXN0KGAvYXBpL3JlY29yZHMvcG9zdC8ke3Bvc3RJZH0vcGxhdGZvcm0vJHtwbGF0Zm9ybX1gLCB7IG1ldGhvZDogJ0RFTEVURScgfSksCiAgfTsKfSkoKTsKCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBTdGF0ZSAvIEZvcm1hdCAvIFRvYXN0IOKAlCBzaGFyZWQgYXBwIHN0YXRlICsgc21hbGwgdXRpbGl0aWVzLgogICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KY29uc3QgU3RhdGUgPSAoKCkgPT4gewogIGNvbnN0IHRvZGF5ID0gbmV3IERhdGUoKTsKICBjb25zdCBpc28gPSAoZCkgPT4gZC50b0lTT1N0cmluZygpLnNsaWNlKDAsIDEwKTsKICBjb25zdCB0aGlydHlEYXlzQWdvID0gbmV3IERhdGUodG9kYXkpOwogIHRoaXJ0eURheXNBZ28uc2V0RGF0ZSh0aGlydHlEYXlzQWdvLmdldERhdGUoKSAtIDI5KTsKCiAgY29uc3QgZmlsdGVycyA9IHsKICAgIGRhdGVGcm9tOiBpc28odGhpcnR5RGF5c0FnbyksCiAgICBkYXRlVG86IGlzbyh0b2RheSksCiAgICBwbGF0Zm9ybTogJ2FsbCcsCiAgICBjYW1wYWlnblR5cGU6ICdhbGwnLAogICAgY29udGVudFR5cGU6ICdhbGwnLAogIH07CgogIGNvbnN0IGxpc3RlbmVycyA9IFtdOwoKICByZXR1cm4gewogICAgZ2V0RmlsdGVyczogKCkgPT4gKHsgLi4uZmlsdGVycyB9KSwKICAgIHNldEZpbHRlcnMocGFydGlhbCkgewogICAgICBPYmplY3QuYXNzaWduKGZpbHRlcnMsIHBhcnRpYWwpOwogICAgICBsaXN0ZW5lcnMuZm9yRWFjaCgoZm4pID0+IGZuKHRoaXMuZ2V0RmlsdGVycygpKSk7CiAgICB9LAogICAgb25DaGFuZ2UoZm4pIHsKICAgICAgbGlzdGVuZXJzLnB1c2goZm4pOwogICAgfSwKICB9Owp9KSgpOwoKY29uc3QgRm9ybWF0ID0gewogIG51bWJlcihuKSB7CiAgICBpZiAobiA9PT0gbnVsbCB8fCBuID09PSB1bmRlZmluZWQpIHJldHVybiAn4oCUJzsKICAgIHJldHVybiBNYXRoLnJvdW5kKG4pLnRvTG9jYWxlU3RyaW5nKCdlbi1VUycpOwogIH0sCiAgY29tcGFjdChuKSB7CiAgICBpZiAobiA9PT0gbnVsbCB8fCBuID09PSB1bmRlZmluZWQpIHJldHVybiAn4oCUJzsKICAgIGNvbnN0IGFicyA9IE1hdGguYWJzKG4pOwogICAgaWYgKGFicyA+PSAxXzAwMF8wMDApIHJldHVybiBgJHsobiAvIDFfMDAwXzAwMCkudG9GaXhlZCgxKS5yZXBsYWNlKC9cLjAkLywgJycpfU1gOwogICAgaWYgKGFicyA+PSAxXzAwMCkgcmV0dXJuIGAkeyhuIC8gMV8wMDApLnRvRml4ZWQoMSkucmVwbGFjZSgvXC4wJC8sICcnKX1LYDsKICAgIHJldHVybiBgJHtNYXRoLnJvdW5kKG4pfWA7CiAgfSwKICAvKiogRGFzaGJvYXJkLXdpZGUgInByb2Zlc3Npb25hbCIgbnVtYmVyIGZvcm1hdDogcGxhaW4gdW5kZXIgMSwwMDA7IGNvbW1hLWdyb3VwZWQKICAgICAgdXAgdG8gMTAsMDAwOyBhYmJyZXZpYXRlZCAoSy9NKSBiZXlvbmQgdGhhdCDigJQgZS5nLiA4NTAsIDEsMjUwLCAxMi41SywgMTU2SywgMS4yNU0uICovCiAgc21hcnQobikgewogICAgaWYgKG4gPT09IG51bGwgfHwgbiA9PT0gdW5kZWZpbmVkKSByZXR1cm4gJ+KAlCc7CiAgICBjb25zdCBhYnMgPSBNYXRoLmFicyhuKTsKICAgIGlmIChhYnMgPCAxMDAwKSByZXR1cm4gYCR7TWF0aC5yb3VuZChuKX1gOwogICAgaWYgKGFicyA8IDEwMDAwKSByZXR1cm4gTWF0aC5yb3VuZChuKS50b0xvY2FsZVN0cmluZygnZW4tVVMnKTsKICAgIGlmIChhYnMgPCAxXzAwMF8wMDApIHJldHVybiBgJHsobiAvIDEwMDApLnRvRml4ZWQoMSkucmVwbGFjZSgvXC4wJC8sICcnKX1LYDsKICAgIHJldHVybiBgJHsobiAvIDFfMDAwXzAwMCkudG9GaXhlZCgyKS5yZXBsYWNlKC9cLj8wKyQvLCAnJyl9TWA7CiAgfSwKICBwZXJjZW50KG4pIHsKICAgIGlmIChuID09PSBudWxsIHx8IG4gPT09IHVuZGVmaW5lZCkgcmV0dXJuICfigJQnOwogICAgcmV0dXJuIGAke051bWJlcihuKS50b0ZpeGVkKDEpLnJlcGxhY2UoL1wuMCQvLCAnJyl9JWA7CiAgfSwKICBwY3QobikgewogICAgaWYgKG4gPT09IG51bGwgfHwgbiA9PT0gdW5kZWZpbmVkKSByZXR1cm4gJ+KAlCc7CiAgICBjb25zdCBzaWduID0gbiA+IDAgPyAnKycgOiAnJzsKICAgIHJldHVybiBgJHtzaWdufSR7bi50b0ZpeGVkKDEpfSVgOwogIH0sCiAgZGF0ZShpc29fKSB7CiAgICBpZiAoIWlzb18pIHJldHVybiAn4oCUJzsKICAgIGNvbnN0IFt5LCBtLCBkXSA9IGlzb18uc3BsaXQoJy0nKS5tYXAoTnVtYmVyKTsKICAgIHJldHVybiBuZXcgRGF0ZSh5LCBtIC0gMSwgZCkudG9Mb2NhbGVEYXRlU3RyaW5nKCdlbi1VUycsIHsgbW9udGg6ICdzaG9ydCcsIGRheTogJ251bWVyaWMnLCB5ZWFyOiAnbnVtZXJpYycgfSk7CiAgfSwKICBkdXJhdGlvbihzZWNvbmRzKSB7CiAgICBpZiAoc2Vjb25kcyA9PT0gbnVsbCB8fCBzZWNvbmRzID09PSB1bmRlZmluZWQpIHJldHVybiAn4oCUJzsKICAgIGNvbnN0IHMgPSBNYXRoLnJvdW5kKHNlY29uZHMpOwogICAgaWYgKHMgPCA2MCkgcmV0dXJuIGAke3N9c2A7CiAgICBpZiAocyA8IDM2MDApIHJldHVybiBgJHtNYXRoLmZsb29yKHMgLyA2MCl9bSAke3MgJSA2MH1zYDsKICAgIGNvbnN0IGggPSBNYXRoLmZsb29yKHMgLyAzNjAwKTsKICAgIGNvbnN0IG0gPSBNYXRoLnJvdW5kKChzICUgMzYwMCkgLyA2MCk7CiAgICByZXR1cm4gYCR7aH1oICR7bX1tYDsKICB9LAogIGRlbHRhQ2xhc3MobikgewogICAgaWYgKG4gPT09IG51bGwgfHwgbiA9PT0gdW5kZWZpbmVkKSByZXR1cm4gJ2ZsYXQnOwogICAgaWYgKG4gPiAwLjUpIHJldHVybiAndXAnOwogICAgaWYgKG4gPCAtMC41KSByZXR1cm4gJ2Rvd24nOwogICAgcmV0dXJuICdmbGF0JzsKICB9LAp9OwoKY29uc3QgVG9hc3QgPSB7CiAgc2hvdyhtZXNzYWdlLCB0eXBlID0gJ3N1Y2Nlc3MnKSB7CiAgICBjb25zdCByb290ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3RvYXN0Um9vdCcpOwogICAgY29uc3QgZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGVsLmNsYXNzTmFtZSA9IGB0b2FzdCAke3R5cGV9YDsKICAgIGVsLnRleHRDb250ZW50ID0gbWVzc2FnZTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQoZWwpOwogICAgc2V0VGltZW91dCgoKSA9PiBlbC5yZW1vdmUoKSwgNTAwMCk7CiAgfSwKfTsKCi8qKiBTYWZlbHkgYnVpbGRzIERPTSB0ZXh0IG5vZGVzIGZvciB1bnRydXN0ZWQgc3RyaW5ncyAoY2FwdGlvbnMsIGZpbGVuYW1lcywgcGxhdGZvcm0gbGFiZWxzIGZyb20gZGF0YSkuICovCmZ1bmN0aW9uIHRleHRFbCh0YWcsIHRleHQsIGNsYXNzTmFtZSkgewogIGNvbnN0IGVsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCh0YWcpOwogIGlmIChjbGFzc05hbWUpIGVsLmNsYXNzTmFtZSA9IGNsYXNzTmFtZTsKICBlbC5hcHBlbmRDaGlsZChkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZSh0ZXh0ID8/ICcnKSk7CiAgcmV0dXJuIGVsOwp9CgovKiogQSBwcmVtaXVtIGVtcHR5IHN0YXRlOiBpY29uICsgZXhwbGFuYXRpb24gKyBvcHRpb25hbCBhY3Rpb24sIGluc3RlYWQgb2YgYSBibGFuayBhcmVhLgogICAgSWNvbnMgcmVuZGVyIHZpYSB0aGUgcGFnZS13aWRlIE11dGF0aW9uT2JzZXJ2ZXIgdGhhdCBjYWxscyBsdWNpZGUuY3JlYXRlSWNvbnMoKSAoc2VlIGJvb3RzdHJhcCkuICovCmZ1bmN0aW9uIGVtcHR5U3RhdGUoeyBpY29uID0gJ2luYm94JywgdGl0bGUsIG1lc3NhZ2UsIGFjdGlvbkxhYmVsLCBvbkFjdGlvbiB9KSB7CiAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogIHdyYXAuY2xhc3NOYW1lID0gJ2VtcHR5LXN0YXRlJzsKICBjb25zdCBpY29uV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogIGljb25XcmFwLmNsYXNzTmFtZSA9ICdlbXB0eS1pY29uJzsKICBpY29uV3JhcC5pbm5lckhUTUwgPSBgPGkgZGF0YS1sdWNpZGU9IiR7aWNvbn0iIHN0eWxlPSJ3aWR0aDoyMnB4O2hlaWdodDoyMnB4OyI+PC9pPmA7CiAgd3JhcC5hcHBlbmRDaGlsZChpY29uV3JhcCk7CiAgaWYgKHRpdGxlKSB3cmFwLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgdGl0bGUsICdlbXB0eS10aXRsZScpKTsKICBpZiAobWVzc2FnZSkgd3JhcC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIG1lc3NhZ2UsICdlbXB0eS1tZXNzYWdlJykpOwogIGlmIChhY3Rpb25MYWJlbCAmJiBvbkFjdGlvbikgewogICAgY29uc3QgYnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBidG4uY2xhc3NOYW1lID0gJ2J0biBwcmltYXJ5JzsKICAgIGJ0bi50ZXh0Q29udGVudCA9IGFjdGlvbkxhYmVsOwogICAgYnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgb25BY3Rpb24pOwogICAgd3JhcC5hcHBlbmRDaGlsZChidG4pOwogIH0KICByZXR1cm4gd3JhcDsKfQoKLyoqIEEgPGJ1dHRvbj4gd2l0aCBhIHNtYWxsIGxlYWRpbmcgTHVjaWRlIGljb24gYmVmb3JlIGl0cyBsYWJlbCAobGFiZWwgaXMgYWx3YXlzIGEgc3RhdGljLCBkZXZlbG9wZXItc3VwcGxpZWQgc3RyaW5nIGF0IGNhbGwgc2l0ZXMsIG5ldmVyIHVzZXIgZGF0YSDigJQgaW5zZXJ0ZWQgdmlhIGNyZWF0ZVRleHROb2RlIHJlZ2FyZGxlc3MpLiAqLwpmdW5jdGlvbiBpY29uQnRuKGNsYXNzTmFtZSwgaWNvbk5hbWUsIGxhYmVsKSB7CiAgY29uc3QgYnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgYnRuLmNsYXNzTmFtZSA9IGNsYXNzTmFtZTsKICBjb25zdCBpY29uID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaScpOwogIGljb24uc2V0QXR0cmlidXRlKCdkYXRhLWx1Y2lkZScsIGljb25OYW1lKTsKICBpY29uLnN0eWxlLndpZHRoID0gJzEzcHgnOwogIGljb24uc3R5bGUuaGVpZ2h0ID0gJzEzcHgnOwogIGJ0bi5hcHBlbmRDaGlsZChpY29uKTsKICBidG4uYXBwZW5kQ2hpbGQoZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUoYCAke2xhYmVsfWApKTsKICByZXR1cm4gYnRuOwp9CgovKiogU2hpbW1lcmluZyBwbGFjZWhvbGRlcnMgc2hvd24gdGhlIGluc3RhbnQgYSBzZWN0aW9uIHN0YXJ0cyBsb2FkaW5nLCBzd2FwcGVkIGZvciByZWFsCiAgICBjb250ZW50IChvciBhbiBlbXB0eSBzdGF0ZSkgb25jZSB0aGUgZmV0Y2ggcmVzb2x2ZXMg4oCUIG5vIGJsYW5rIGFyZWFzIHdoaWxlIHdhaXRpbmcuICovCmZ1bmN0aW9uIHNrZWxldG9uU3RhdEdyaWQoY291bnQgPSA2KSB7CiAgY29uc3QgZ3JpZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogIGdyaWQuY2xhc3NOYW1lID0gJ3NrZWxldG9uLXN0YXQtZ3JpZCc7CiAgZm9yIChsZXQgaSA9IDA7IGkgPCBjb3VudDsgaSArPSAxKSB7CiAgICBjb25zdCB0aWxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0aWxlLmNsYXNzTmFtZSA9ICdza2VsZXRvbiBza2VsZXRvbi10aWxlJzsKICAgIGdyaWQuYXBwZW5kQ2hpbGQodGlsZSk7CiAgfQogIHJldHVybiBncmlkOwp9CmZ1bmN0aW9uIHNrZWxldG9uQ2hhcnQoKSB7CiAgY29uc3QgZGl2ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgZGl2LmNsYXNzTmFtZSA9ICdza2VsZXRvbiBza2VsZXRvbi1jaGFydCc7CiAgcmV0dXJuIGRpdjsKfQpmdW5jdGlvbiBza2VsZXRvblJvd3MoY291bnQgPSA2KSB7CiAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogIGZvciAobGV0IGkgPSAwOyBpIDwgY291bnQ7IGkgKz0gMSkgewogICAgY29uc3Qgcm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICByb3cuY2xhc3NOYW1lID0gJ3NrZWxldG9uIHNrZWxldG9uLXJvdyc7CiAgICB3cmFwLmFwcGVuZENoaWxkKHJvdyk7CiAgfQogIHJldHVybiB3cmFwOwp9CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgQ2hhcnRzIOKAlCBDaGFydC5qcyBidWlsZGVycyAodmFsaWRhdGVkIGNhdGVnb3JpY2FsIHBhbGV0dGUsCiAgIGhhaXJsaW5lIHJlY2Vzc2l2ZSBncmlkbGluZXMsIHNpbmdsZSBheGlzLCBsZWdlbmQgYWx3YXlzCiAgIHByZXNlbnQgZm9yIDIrIHNlcmllcywgaW5kZXgtbW9kZSB0b29sdGlwcykuCiAgID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwppZiAod2luZG93LkNoYXJ0RGF0YUxhYmVscykgQ2hhcnQucmVnaXN0ZXIod2luZG93LkNoYXJ0RGF0YUxhYmVscyk7Cgpjb25zdCBDaGFydHMgPSAoKCkgPT4gewogIGNvbnN0IHJlZ2lzdHJ5ID0gbmV3IE1hcCgpOyAvLyBjYW52YXNJZCAtPiBDaGFydCBpbnN0YW5jZSwgc28gcmUtcmVuZGVycyBkZXN0cm95IHRoZSBvbGQgb25lIGZpcnN0CiAgY29uc3QgTUFYX0xBQkVMRURfSVRFTVMgPSAyMDsgLy8gYmV5b25kIHRoaXMsIHBlci1pdGVtIHZhbHVlIGxhYmVscyB3b3VsZCBvdmVybGFwIOKAlCByZWx5IG9uIHRvb2x0aXBzIGluc3RlYWQKCiAgZnVuY3Rpb24gY3NzVmFyKG5hbWUpIHsKICAgIHJldHVybiBnZXRDb21wdXRlZFN0eWxlKGRvY3VtZW50LmRvY3VtZW50RWxlbWVudCkuZ2V0UHJvcGVydHlWYWx1ZShuYW1lKS50cmltKCk7CiAgfQoKICBjb25zdCBTRVJJRVNfVkFSUyA9IFsnLS1zZXJpZXMtMScsICctLXNlcmllcy0yJywgJy0tc2VyaWVzLTMnLCAnLS1zZXJpZXMtNCcsICctLXNlcmllcy01JywgJy0tc2VyaWVzLTYnLCAnLS1zZXJpZXMtNycsICctLXNlcmllcy04J107CiAgZnVuY3Rpb24gc2VyaWVzQ29sb3IoaW5kZXgpIHsKICAgIHJldHVybiBjc3NWYXIoU0VSSUVTX1ZBUlNbaW5kZXggJSBTRVJJRVNfVkFSUy5sZW5ndGhdKTsKICB9CgogIGZ1bmN0aW9uIGJhc2VHcmlkKCkgewogICAgcmV0dXJuIHsKICAgICAgY29sb3I6IGNzc1ZhcignLS1ncmlkbGluZScpLAogICAgICBkcmF3VGlja3M6IGZhbHNlLAogICAgfTsKICB9CiAgZnVuY3Rpb24gYmFzZVRpY2tzKCkgewogICAgcmV0dXJuIHsgY29sb3I6IGNzc1ZhcignLS10ZXh0LW11dGVkJyksIGZvbnQ6IHsgc2l6ZTogMTEgfSB9OwogIH0KICBmdW5jdGlvbiBiYXNlVG9vbHRpcCgpIHsKICAgIHJldHVybiB7CiAgICAgIGJhY2tncm91bmRDb2xvcjogY3NzVmFyKCctLXN1cmZhY2UtMScpLAogICAgICB0aXRsZUNvbG9yOiBjc3NWYXIoJy0tdGV4dC1wcmltYXJ5JyksCiAgICAgIGJvZHlDb2xvcjogY3NzVmFyKCctLXRleHQtc2Vjb25kYXJ5JyksCiAgICAgIGJvcmRlckNvbG9yOiBjc3NWYXIoJy0tYm9yZGVyJyksCiAgICAgIGJvcmRlcldpZHRoOiAxLAogICAgICBjb3JuZXJSYWRpdXM6IDEwLAogICAgICBwYWRkaW5nOiAxMiwKICAgICAgYm94UGFkZGluZzogNCwKICAgICAgdGl0bGVGb250OiB7IHNpemU6IDEyLCB3ZWlnaHQ6ICc3MDAnIH0sCiAgICAgIGJvZHlGb250OiB7IHNpemU6IDEyIH0sCiAgICB9OwogIH0KICBmdW5jdGlvbiBsYWJlbENvbG9yKCkgewogICAgcmV0dXJuIGNzc1ZhcignLS10ZXh0LXByaW1hcnknKTsKICB9CiAgLyoqIFNuYXBweSwgc3VidGxlIG1vdGlvbiDigJQgaW4gdGhlIDE1MC0zMDBtcyByYW5nZSB0aGUgcmVkZXNpZ24gY2FsbHMgZm9yLCBuZXZlciBib3VuY3kuICovCiAgZnVuY3Rpb24gYmFzZUFuaW1hdGlvbigpIHsKICAgIHJldHVybiB7IGR1cmF0aW9uOiAyODAsIGVhc2luZzogJ2Vhc2VPdXRRdWFydCcgfTsKICB9CgogIGZ1bmN0aW9uIGRlc3Ryb3koY2FudmFzSWQpIHsKICAgIGlmIChyZWdpc3RyeS5oYXMoY2FudmFzSWQpKSB7CiAgICAgIHJlZ2lzdHJ5LmdldChjYW52YXNJZCkuZGVzdHJveSgpOwogICAgICByZWdpc3RyeS5kZWxldGUoY2FudmFzSWQpOwogICAgfQogIH0KCiAgLyoqIE11bHRpLXNlcmllcyBsaW5lIGNoYXJ0IChlLmcuIHdlZWtseSB0cmVuZCBwZXIgcGxhdGZvcm0pLiBPbmUgc2VyaWVzIG5lZWRzIG5vIGxlZ2VuZCBib3guCiAgICAgIFBlci1wb2ludCB2YWx1ZSBsYWJlbHMgYXJlIHNob3duIG9ubHkgZm9yIGEgc2luZ2xlIHNlcmllcyDigJQgd2l0aCBzZXZlcmFsIHNlcmllcyBvdmVybGFpZCwKICAgICAgbGFiZWxpbmcgZXZlcnkgcG9pbnQgd291bGQgb3ZlcmxhcCwgc28gdGhvc2UgcmVseSBvbiB0aGUgKHN0aWxsLXByZXNlbnQpIGhvdmVyIHRvb2x0aXAuICovCiAgZnVuY3Rpb24gdHJlbmRDaGFydChjYW52YXNJZCwgeyBsYWJlbHMsIHNlcmllcywgZm9ybWF0VmFsdWUgfSkgewogICAgZGVzdHJveShjYW52YXNJZCk7CiAgICBjb25zdCBjdHggPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZChjYW52YXNJZCk7CiAgICBpZiAoIWN0eCkgcmV0dXJuIG51bGw7CiAgICBjb25zdCBmbXQgPSBmb3JtYXRWYWx1ZSB8fCAoKHYpID0+IEZvcm1hdC5zbWFydCh2KSk7CiAgICBjb25zdCBzaG93TGFiZWxzID0gc2VyaWVzLmxlbmd0aCA9PT0gMSAmJiBsYWJlbHMubGVuZ3RoIDw9IE1BWF9MQUJFTEVEX0lURU1TOwoKICAgIGNvbnN0IGRhdGFzZXRzID0gc2VyaWVzLm1hcCgocywgaSkgPT4gKHsKICAgICAgbGFiZWw6IHMubGFiZWwsCiAgICAgIGRhdGE6IHMuZGF0YSwKICAgICAgYm9yZGVyQ29sb3I6IHMuY29sb3IgfHwgc2VyaWVzQ29sb3IoaSksCiAgICAgIGJhY2tncm91bmRDb2xvcjogcy5jb2xvciB8fCBzZXJpZXNDb2xvcihpKSwKICAgICAgYm9yZGVyV2lkdGg6IDIsCiAgICAgIHBvaW50UmFkaXVzOiBzaG93TGFiZWxzID8gMyA6IDAsCiAgICAgIHBvaW50SG92ZXJSYWRpdXM6IDQsCiAgICAgIHBvaW50SGl0UmFkaXVzOiAxMiwKICAgICAgdGVuc2lvbjogMC4yNSwKICAgICAgZmlsbDogZmFsc2UsCiAgICB9KSk7CgogICAgY29uc3QgY2hhcnQgPSBuZXcgQ2hhcnQoY3R4LCB7CiAgICAgIHR5cGU6ICdsaW5lJywKICAgICAgZGF0YTogeyBsYWJlbHMsIGRhdGFzZXRzIH0sCiAgICAgIG9wdGlvbnM6IHsKICAgICAgICByZXNwb25zaXZlOiB0cnVlLAogICAgICAgIG1haW50YWluQXNwZWN0UmF0aW86IGZhbHNlLAogICAgICAgIGludGVyYWN0aW9uOiB7IG1vZGU6ICdpbmRleCcsIGludGVyc2VjdDogZmFsc2UgfSwKICAgICAgICBsYXlvdXQ6IHsgcGFkZGluZzogeyB0b3A6IHNob3dMYWJlbHMgPyAyMCA6IDggfSB9LAogICAgICAgIGFuaW1hdGlvbjogYmFzZUFuaW1hdGlvbigpLAogICAgICAgIHBsdWdpbnM6IHsKICAgICAgICAgIGxlZ2VuZDogewogICAgICAgICAgICBkaXNwbGF5OiBzZXJpZXMubGVuZ3RoID4gMSwKICAgICAgICAgICAgcG9zaXRpb246ICdib3R0b20nLAogICAgICAgICAgICBsYWJlbHM6IHsgY29sb3I6IGNzc1ZhcignLS10ZXh0LXNlY29uZGFyeScpLCB1c2VQb2ludFN0eWxlOiB0cnVlLCBwb2ludFN0eWxlOiAnbGluZScsIGJveFdpZHRoOiAxNiwgcGFkZGluZzogMTYsIGZvbnQ6IHsgc2l6ZTogMTEgfSB9LAogICAgICAgICAgfSwKICAgICAgICAgIHRvb2x0aXA6IHsgLi4uYmFzZVRvb2x0aXAoKSwgdXNlUG9pbnRTdHlsZTogdHJ1ZSB9LAogICAgICAgICAgZGF0YWxhYmVsczogc2hvd0xhYmVscwogICAgICAgICAgICA/IHsgYWxpZ246ICd0b3AnLCBhbmNob3I6ICdlbmQnLCBjb2xvcjogbGFiZWxDb2xvcigpLCBmb250OiB7IHNpemU6IDExLCB3ZWlnaHQ6ICc2MDAnIH0sIGZvcm1hdHRlcjogKHYpID0+IGZtdCh2KSB9CiAgICAgICAgICAgIDogeyBkaXNwbGF5OiBmYWxzZSB9LAogICAgICAgIH0sCiAgICAgICAgc2NhbGVzOiB7CiAgICAgICAgICB4OiB7IGdyaWQ6IHsgZGlzcGxheTogZmFsc2UgfSwgdGlja3M6IGJhc2VUaWNrcygpIH0sCiAgICAgICAgICB5OiB7IGdyaWQ6IGJhc2VHcmlkKCksIHRpY2tzOiBiYXNlVGlja3MoKSwgYm9yZGVyOiB7IGRpc3BsYXk6IGZhbHNlIH0sIGJlZ2luQXRaZXJvOiB0cnVlIH0sCiAgICAgICAgfSwKICAgICAgfSwKICAgIH0pOwogICAgcmVnaXN0cnkuc2V0KGNhbnZhc0lkLCBjaGFydCk7CiAgICByZXR1cm4gY2hhcnQ7CiAgfQoKICAvKiogU2luZ2xlLW1ldHJpYyBiYXIgY2hhcnQgYWNyb3NzIHBsYXRmb3JtcyAoaWRlbnRpdHkgZW5jb2Rpbmcg4oCUIGVhY2ggYmFyIElTIGEgcGxhdGZvcm0pLiAqLwogIGZ1bmN0aW9uIHBsYXRmb3JtQmFyQ2hhcnQoY2FudmFzSWQsIHsgbGFiZWxzLCBkYXRhLCBjb2xvcnMsIGZvcm1hdFZhbHVlIH0pIHsKICAgIGRlc3Ryb3koY2FudmFzSWQpOwogICAgY29uc3QgY3R4ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoY2FudmFzSWQpOwogICAgaWYgKCFjdHgpIHJldHVybiBudWxsOwogICAgY29uc3QgZm10ID0gZm9ybWF0VmFsdWUgfHwgKCh2KSA9PiBGb3JtYXQuc21hcnQodikpOwogICAgY29uc3Qgc2hvd0xhYmVscyA9IGxhYmVscy5sZW5ndGggPD0gTUFYX0xBQkVMRURfSVRFTVM7CgogICAgY29uc3QgY2hhcnQgPSBuZXcgQ2hhcnQoY3R4LCB7CiAgICAgIHR5cGU6ICdiYXInLAogICAgICBkYXRhOiB7CiAgICAgICAgbGFiZWxzLAogICAgICAgIGRhdGFzZXRzOiBbCiAgICAgICAgICB7CiAgICAgICAgICAgIGRhdGEsCiAgICAgICAgICAgIGJhY2tncm91bmRDb2xvcjogY29sb3JzLAogICAgICAgICAgICBib3JkZXJSYWRpdXM6IDQsCiAgICAgICAgICAgIG1heEJhclRoaWNrbmVzczogMjgsCiAgICAgICAgICAgIGJvcmRlclNraXBwZWQ6ICdib3R0b20nLAogICAgICAgICAgfSwKICAgICAgICBdLAogICAgICB9LAogICAgICBvcHRpb25zOiB7CiAgICAgICAgcmVzcG9uc2l2ZTogdHJ1ZSwKICAgICAgICBtYWludGFpbkFzcGVjdFJhdGlvOiBmYWxzZSwKICAgICAgICBsYXlvdXQ6IHsgcGFkZGluZzogeyB0b3A6IHNob3dMYWJlbHMgPyAyMCA6IDggfSB9LAogICAgICAgIGFuaW1hdGlvbjogYmFzZUFuaW1hdGlvbigpLAogICAgICAgIHBsdWdpbnM6IHsKICAgICAgICAgIGxlZ2VuZDogeyBkaXNwbGF5OiBmYWxzZSB9LAogICAgICAgICAgdG9vbHRpcDogYmFzZVRvb2x0aXAoKSwKICAgICAgICAgIGRhdGFsYWJlbHM6IHNob3dMYWJlbHMKICAgICAgICAgICAgPyB7IGFsaWduOiAnZW5kJywgYW5jaG9yOiAnZW5kJywgY29sb3I6IGxhYmVsQ29sb3IoKSwgZm9udDogeyBzaXplOiAxMSwgd2VpZ2h0OiAnNjAwJyB9LCBmb3JtYXR0ZXI6ICh2KSA9PiBmbXQodikgfQogICAgICAgICAgICA6IHsgZGlzcGxheTogZmFsc2UgfSwKICAgICAgICB9LAogICAgICAgIHNjYWxlczogewogICAgICAgICAgeDogeyBncmlkOiB7IGRpc3BsYXk6IGZhbHNlIH0sIHRpY2tzOiBiYXNlVGlja3MoKSB9LAogICAgICAgICAgeTogeyBncmlkOiBiYXNlR3JpZCgpLCB0aWNrczogYmFzZVRpY2tzKCksIGJvcmRlcjogeyBkaXNwbGF5OiBmYWxzZSB9LCBiZWdpbkF0WmVybzogdHJ1ZSB9LAogICAgICAgIH0sCiAgICAgIH0sCiAgICB9KTsKICAgIHJlZ2lzdHJ5LnNldChjYW52YXNJZCwgY2hhcnQpOwogICAgcmV0dXJuIGNoYXJ0OwogIH0KCiAgLyoqIEdyb3VwZWQgYmFyIGNoYXJ0IGNvbXBhcmluZyB0d28gdGltZSBwZXJpb2RzIChjb2xvciA9IHBlcmlvZCwgeC1heGlzID0gY2F0ZWdvcnkpLiAqLwogIGZ1bmN0aW9uIGNvbXBhcmlzb25CYXJDaGFydChjYW52YXNJZCwgeyBsYWJlbHMsIGN1cnJlbnREYXRhLCBwcmV2aW91c0RhdGEsIGN1cnJlbnRMYWJlbCwgcHJldmlvdXNMYWJlbCwgZm9ybWF0VmFsdWUgfSkgewogICAgZGVzdHJveShjYW52YXNJZCk7CiAgICBjb25zdCBjdHggPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZChjYW52YXNJZCk7CiAgICBpZiAoIWN0eCkgcmV0dXJuIG51bGw7CiAgICBjb25zdCBmbXQgPSBmb3JtYXRWYWx1ZSB8fCAoKHYpID0+IEZvcm1hdC5zbWFydCh2KSk7CiAgICBjb25zdCBzaG93TGFiZWxzID0gbGFiZWxzLmxlbmd0aCA8PSBNQVhfTEFCRUxFRF9JVEVNUyAvIDI7CgogICAgY29uc3QgY2hhcnQgPSBuZXcgQ2hhcnQoY3R4LCB7CiAgICAgIHR5cGU6ICdiYXInLAogICAgICBkYXRhOiB7CiAgICAgICAgbGFiZWxzLAogICAgICAgIGRhdGFzZXRzOiBbCiAgICAgICAgICB7IGxhYmVsOiBwcmV2aW91c0xhYmVsLCBkYXRhOiBwcmV2aW91c0RhdGEsIGJhY2tncm91bmRDb2xvcjogY3NzVmFyKCctLXRleHQtbXV0ZWQnKSwgYm9yZGVyUmFkaXVzOiA0LCBtYXhCYXJUaGlja25lc3M6IDIyIH0sCiAgICAgICAgICB7IGxhYmVsOiBjdXJyZW50TGFiZWwsIGRhdGE6IGN1cnJlbnREYXRhLCBiYWNrZ3JvdW5kQ29sb3I6IGNzc1ZhcignLS1zZXJpZXMtMScpLCBib3JkZXJSYWRpdXM6IDQsIG1heEJhclRoaWNrbmVzczogMjIgfSwKICAgICAgICBdLAogICAgICB9LAogICAgICBvcHRpb25zOiB7CiAgICAgICAgcmVzcG9uc2l2ZTogdHJ1ZSwKICAgICAgICBtYWludGFpbkFzcGVjdFJhdGlvOiBmYWxzZSwKICAgICAgICBsYXlvdXQ6IHsgcGFkZGluZzogeyB0b3A6IHNob3dMYWJlbHMgPyAyMCA6IDggfSB9LAogICAgICAgIGFuaW1hdGlvbjogYmFzZUFuaW1hdGlvbigpLAogICAgICAgIHBsdWdpbnM6IHsKICAgICAgICAgIGxlZ2VuZDogeyBkaXNwbGF5OiB0cnVlLCBwb3NpdGlvbjogJ2JvdHRvbScsIGxhYmVsczogeyBjb2xvcjogY3NzVmFyKCctLXRleHQtc2Vjb25kYXJ5JyksIGJveFdpZHRoOiAxMiwgcGFkZGluZzogMTYsIGZvbnQ6IHsgc2l6ZTogMTEgfSB9IH0sCiAgICAgICAgICB0b29sdGlwOiBiYXNlVG9vbHRpcCgpLAogICAgICAgICAgZGF0YWxhYmVsczogc2hvd0xhYmVscwogICAgICAgICAgICA/IHsgYWxpZ246ICdlbmQnLCBhbmNob3I6ICdlbmQnLCBjb2xvcjogbGFiZWxDb2xvcigpLCBmb250OiB7IHNpemU6IDEwLCB3ZWlnaHQ6ICc2MDAnIH0sIGZvcm1hdHRlcjogKHYpID0+IGZtdCh2KSB9CiAgICAgICAgICAgIDogeyBkaXNwbGF5OiBmYWxzZSB9LAogICAgICAgIH0sCiAgICAgICAgc2NhbGVzOiB7CiAgICAgICAgICB4OiB7IGdyaWQ6IHsgZGlzcGxheTogZmFsc2UgfSwgdGlja3M6IGJhc2VUaWNrcygpIH0sCiAgICAgICAgICB5OiB7IGdyaWQ6IGJhc2VHcmlkKCksIHRpY2tzOiBiYXNlVGlja3MoKSwgYm9yZGVyOiB7IGRpc3BsYXk6IGZhbHNlIH0sIGJlZ2luQXRaZXJvOiB0cnVlIH0sCiAgICAgICAgfSwKICAgICAgfSwKICAgIH0pOwogICAgcmVnaXN0cnkuc2V0KGNhbnZhc0lkLCBjaGFydCk7CiAgICByZXR1cm4gY2hhcnQ7CiAgfQoKICAvKiogUGllIGNoYXJ0IChhIGhhbmRmdWwgb2YgY2F0ZWdvcmllcyBvbmx5IOKAlCBlLmcuIENhbXBhaWduIFBlcmZvcm1hbmNlJ3MgQWRzL09yZ2FuaWMgc3BsaXQpLgogICAgICBTbGljZSBsYWJlbHMgc2hvdyBib3RoIHNoYXJlLW9mLXdob2xlIGFuZCB0aGUgYWN0dWFsIHZhbHVlLCBwZXIgdGhlICJubyBob3ZlciByZXF1aXJlZCIgZ29hbC4gKi8KICBmdW5jdGlvbiBwaWVDaGFydChjYW52YXNJZCwgeyBsYWJlbHMsIGRhdGEsIGNvbG9ycywgZm9ybWF0VmFsdWUgfSkgewogICAgZGVzdHJveShjYW52YXNJZCk7CiAgICBjb25zdCBjdHggPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZChjYW52YXNJZCk7CiAgICBpZiAoIWN0eCkgcmV0dXJuIG51bGw7CiAgICBjb25zdCBmbXQgPSBmb3JtYXRWYWx1ZSB8fCAoKHYpID0+IEZvcm1hdC5zbWFydCh2KSk7CiAgICBjb25zdCB0b3RhbCA9IGRhdGEucmVkdWNlKChzdW0sIHYpID0+IHN1bSArICh2IHx8IDApLCAwKTsKCiAgICBjb25zdCBjaGFydCA9IG5ldyBDaGFydChjdHgsIHsKICAgICAgdHlwZTogJ3BpZScsCiAgICAgIGRhdGE6IHsKICAgICAgICBsYWJlbHMsCiAgICAgICAgZGF0YXNldHM6IFt7IGRhdGEsIGJhY2tncm91bmRDb2xvcjogY29sb3JzLCBib3JkZXJDb2xvcjogY3NzVmFyKCctLXN1cmZhY2UtMScpLCBib3JkZXJXaWR0aDogMiB9XSwKICAgICAgfSwKICAgICAgb3B0aW9uczogewogICAgICAgIHJlc3BvbnNpdmU6IHRydWUsCiAgICAgICAgbWFpbnRhaW5Bc3BlY3RSYXRpbzogZmFsc2UsCiAgICAgICAgYW5pbWF0aW9uOiBiYXNlQW5pbWF0aW9uKCksCiAgICAgICAgcGx1Z2luczogewogICAgICAgICAgbGVnZW5kOiB7IGRpc3BsYXk6IHRydWUsIHBvc2l0aW9uOiAnYm90dG9tJywgbGFiZWxzOiB7IGNvbG9yOiBjc3NWYXIoJy0tdGV4dC1zZWNvbmRhcnknKSwgYm94V2lkdGg6IDEyLCBwYWRkaW5nOiAxNiwgZm9udDogeyBzaXplOiAxMSB9IH0gfSwKICAgICAgICAgIHRvb2x0aXA6IGJhc2VUb29sdGlwKCksCiAgICAgICAgICBkYXRhbGFiZWxzOiB7CiAgICAgICAgICAgIGNvbG9yOiAnI2ZmZicsCiAgICAgICAgICAgIGZvbnQ6IHsgc2l6ZTogMTIsIHdlaWdodDogJzcwMCcgfSwKICAgICAgICAgICAgZm9ybWF0dGVyOiAodikgPT4gewogICAgICAgICAgICAgIGNvbnN0IHBjdCA9IHRvdGFsID8gTWF0aC5yb3VuZCgodiAvIHRvdGFsKSAqIDEwMDApIC8gMTAgOiAwOwogICAgICAgICAgICAgIHJldHVybiBgJHtwY3R9JVxuJHtmbXQodil9YDsKICAgICAgICAgICAgfSwKICAgICAgICAgIH0sCiAgICAgICAgfSwKICAgICAgfSwKICAgIH0pOwogICAgcmVnaXN0cnkuc2V0KGNhbnZhc0lkLCBjaGFydCk7CiAgICByZXR1cm4gY2hhcnQ7CiAgfQoKICBmdW5jdGlvbiBkZXN0cm95QWxsKCkgewogICAgWy4uLnJlZ2lzdHJ5LmtleXMoKV0uZm9yRWFjaChkZXN0cm95KTsKICB9CgogIHJldHVybiB7IHRyZW5kQ2hhcnQsIHBsYXRmb3JtQmFyQ2hhcnQsIGNvbXBhcmlzb25CYXJDaGFydCwgcGllQ2hhcnQsIHNlcmllc0NvbG9yLCBkZXN0cm95LCBkZXN0cm95QWxsIH07Cn0pKCk7CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgRGFzaGJvYXJkIHRhYjogYSBtZXRyaWMtZm9jdXNlZCBwcmVtaXVtIEJJIGRhc2hib2FyZC4gQSBzaW5nbGUKICAgTWV0cmljIHNlbGVjdG9yIChkeW5hbWljYWxseSBwb3B1bGF0ZWQgZnJvbSB3aGF0ZXZlciB0aGUKICAgc2VsZWN0ZWQgcGxhdGZvcm0ncyBkYXRhIGFjdHVhbGx5IGhhcyDigJQgbmV2ZXIgaGFyZGNvZGVkKSBkcml2ZXMKICAgdGhlIEtQSSBjYXJkcywgd2Vla2x5IHRyZW5kLCBwbGF0Zm9ybS9jYW1wYWlnbi9jb250ZW50LXR5cGUKICAgYnJlYWtkb3ducywgYW5kIHRoZSBUb3AgUGVyZm9ybWluZyBQb3N0cyByYW5raW5nIHRvZ2V0aGVyOwogICBQbGF0Zm9ybS9kYXRlL2NhbXBhaWduL2NvbnRlbnQtdHlwZSBmaWx0ZXJpbmcgY29tZXMgZnJvbSB0aGUKICAgc2hhcmVkIGZpbHRlciBiYXIuIEV2ZXJ5IGNoYXJ0IHNob3dzIGl0cyB2YWx1ZXMgZGlyZWN0bHkgKHZpYQogICBjaGFydGpzLXBsdWdpbi1kYXRhbGFiZWxzKSBzbyBub3RoaW5nIHJlcXVpcmVzIGEgaG92ZXIgdG8gcmVhZC4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmNvbnN0IERhc2hib2FyZCA9ICgoKSA9PiB7CiAgbGV0IHJvb3Q7CiAgbGV0IG1ldHJpYyA9ICd2aWV3cyc7CiAgbGV0IG1ldHJpY09wdGlvbnMgPSBbXTsKCiAgZnVuY3Rpb24gb3B0aW9uRm9yKGtleSkgewogICAgcmV0dXJuIG1ldHJpY09wdGlvbnMuZmluZCgobSkgPT4gbS5rZXkgPT09IGtleSk7CiAgfQogIGZ1bmN0aW9uIG1ldHJpY0xhYmVsKGtleSkgewogICAgY29uc3Qgb3B0ID0gb3B0aW9uRm9yKGtleSk7CiAgICByZXR1cm4gb3B0ID8gb3B0LmxhYmVsIDoga2V5OwogIH0KICBmdW5jdGlvbiBtZXRyaWNVbml0KGtleSkgewogICAgY29uc3Qgb3B0ID0gb3B0aW9uRm9yKGtleSk7CiAgICByZXR1cm4gb3B0ID8gb3B0LnVuaXQgOiAnbnVtYmVyJzsKICB9CiAgZnVuY3Rpb24gZm9ybWF0TWV0cmljVmFsdWUoa2V5LCB2YWx1ZSkgewogICAgY29uc3QgdW5pdCA9IG1ldHJpY1VuaXQoa2V5KTsKICAgIGlmICh2YWx1ZSA9PT0gbnVsbCB8fCB2YWx1ZSA9PT0gdW5kZWZpbmVkKSByZXR1cm4gJ+KAlCc7CiAgICBpZiAodW5pdCA9PT0gJ2R1cmF0aW9uJykgcmV0dXJuIEZvcm1hdC5kdXJhdGlvbih2YWx1ZSk7CiAgICBpZiAodW5pdCA9PT0gJ3BlcmNlbnQnKSByZXR1cm4gRm9ybWF0LnBlcmNlbnQodmFsdWUpOwogICAgcmV0dXJuIEZvcm1hdC5zbWFydCh2YWx1ZSk7CiAgfQoKICBmdW5jdGlvbiBzaGVsbCgpIHsKICAgIHJvb3QuaW5uZXJIVE1MID0gJyc7CgogICAgY29uc3QgY29udHJvbHMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNvbnRyb2xzLmNsYXNzTmFtZSA9ICdkYXNoYm9hcmQtY29udHJvbHMnOwogICAgY29uc3QgbGFiZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdsYWJlbCcpOwogICAgbGFiZWwudGV4dENvbnRlbnQgPSAnTWV0cmljJzsKICAgIGxhYmVsLnNldEF0dHJpYnV0ZSgnZm9yJywgJ2Rhc2hib2FyZE1ldHJpY1NlbGVjdCcpOwogICAgY29uc3QgbWV0cmljU2VsZWN0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc2VsZWN0Jyk7CiAgICBtZXRyaWNTZWxlY3QuaWQgPSAnZGFzaGJvYXJkTWV0cmljU2VsZWN0JzsKICAgIG1ldHJpY09wdGlvbnMuZm9yRWFjaCgobSkgPT4gewogICAgICBjb25zdCBvcHQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsKICAgICAgb3B0LnZhbHVlID0gbS5rZXk7CiAgICAgIG9wdC50ZXh0Q29udGVudCA9IG0ubGFiZWw7CiAgICAgIGlmIChtLmtleSA9PT0gbWV0cmljKSBvcHQuc2VsZWN0ZWQgPSB0cnVlOwogICAgICBtZXRyaWNTZWxlY3QuYXBwZW5kQ2hpbGQob3B0KTsKICAgIH0pOwogICAgbWV0cmljU2VsZWN0LmFkZEV2ZW50TGlzdGVuZXIoJ2NoYW5nZScsICgpID0+IHsKICAgICAgbWV0cmljID0gbWV0cmljU2VsZWN0LnZhbHVlOwogICAgICByZWZyZXNoRm9yTWV0cmljKCk7CiAgICB9KTsKICAgIGNvbnRyb2xzLmFwcGVuZChsYWJlbCwgbWV0cmljU2VsZWN0KTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQoY29udHJvbHMpOwoKICAgIGNvbnN0IGtwaVRpdGxlID0gdGV4dEVsKCdkaXYnLCAnS2V5IHBlcmZvcm1hbmNlIGluZGljYXRvcnMnLCAnc2VjdGlvbi10aXRsZScpOwogICAgY29uc3Qga3BpR3JpZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAga3BpR3JpZC5jbGFzc05hbWUgPSAnc3RhdC1ncmlkJzsKICAgIGtwaUdyaWQuaWQgPSAna3BpR3JpZCc7CiAgICByb290LmFwcGVuZChrcGlUaXRsZSwga3BpR3JpZCk7CgogICAgY29uc3QgY2hhcnRzVGl0bGUgPSB0ZXh0RWwoJ2RpdicsICdUcmVuZCAmIHBlcmZvcm1hbmNlIGJyZWFrZG93bicsICdzZWN0aW9uLXRpdGxlJyk7CiAgICByb290LmFwcGVuZChjaGFydHNUaXRsZSk7CgogICAgY29uc3QgdHJlbmRDYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0cmVuZENhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgY29uc3QgdHJlbmRIZWFkZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRyZW5kSGVhZGVyLmNsYXNzTmFtZSA9ICdjYXJkLWhlYWRlcic7CiAgICB0cmVuZEhlYWRlci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2gzJywgJ1dlZWtseSBwZXJmb3JtYW5jZScpKTsKICAgIHRyZW5kSGVhZGVyLmZpcnN0Q2hpbGQuaWQgPSAndHJlbmRDYXJkVGl0bGUnOwogICAgdHJlbmRDYXJkLmFwcGVuZENoaWxkKHRyZW5kSGVhZGVyKTsKICAgIGNvbnN0IHRyZW5kQ2hhcnRXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0cmVuZENoYXJ0V3JhcC5jbGFzc05hbWUgPSAnY2hhcnQtd3JhcCB0YWxsJzsKICAgIHRyZW5kQ2hhcnRXcmFwLmlkID0gJ3RyZW5kQ2hhcnRXcmFwJzsKICAgIHRyZW5kQ2hhcnRXcmFwLmlubmVySFRNTCA9ICc8Y2FudmFzIGlkPSJ0cmVuZENhbnZhcyI+PC9jYW52YXM+JzsKICAgIHRyZW5kQ2FyZC5hcHBlbmRDaGlsZCh0cmVuZENoYXJ0V3JhcCk7CiAgICByb290LmFwcGVuZENoaWxkKHRyZW5kQ2FyZCk7CgogICAgY29uc3QgZ3JpZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZ3JpZC5jbGFzc05hbWUgPSAnY2FyZC1ncmlkIGV2ZW4nOwogICAgZ3JpZC5zdHlsZS5tYXJnaW5Ub3AgPSAnMTZweCc7CgogICAgY29uc3QgYnJlYWtkb3duQ2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYnJlYWtkb3duQ2FyZC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICBjb25zdCBicmVha2Rvd25IZWFkZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGJyZWFrZG93bkhlYWRlci5jbGFzc05hbWUgPSAnY2FyZC1oZWFkZXInOwogICAgYnJlYWtkb3duSGVhZGVyLmFwcGVuZENoaWxkKHRleHRFbCgnaDMnLCAnJykpOwogICAgYnJlYWtkb3duSGVhZGVyLmZpcnN0Q2hpbGQuaWQgPSAnYnJlYWtkb3duQ2FyZFRpdGxlJzsKICAgIGJyZWFrZG93bkNhcmQuYXBwZW5kQ2hpbGQoYnJlYWtkb3duSGVhZGVyKTsKICAgIGNvbnN0IGJyZWFrZG93bldyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGJyZWFrZG93bldyYXAuY2xhc3NOYW1lID0gJ2NoYXJ0LXdyYXAnOwogICAgYnJlYWtkb3duV3JhcC5pZCA9ICdicmVha2Rvd25DaGFydFdyYXAnOwogICAgYnJlYWtkb3duV3JhcC5pbm5lckhUTUwgPSAnPGNhbnZhcyBpZD0iYnJlYWtkb3duQ2FudmFzIj48L2NhbnZhcz4nOwogICAgYnJlYWtkb3duQ2FyZC5hcHBlbmRDaGlsZChicmVha2Rvd25XcmFwKTsKCiAgICBjb25zdCBjb250ZW50VHlwZUNhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNvbnRlbnRUeXBlQ2FyZC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICBjb25zdCBjb250ZW50VHlwZUhlYWRlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY29udGVudFR5cGVIZWFkZXIuY2xhc3NOYW1lID0gJ2NhcmQtaGVhZGVyJzsKICAgIGNvbnRlbnRUeXBlSGVhZGVyLmFwcGVuZENoaWxkKHRleHRFbCgnaDMnLCAnJykpOwogICAgY29udGVudFR5cGVIZWFkZXIuZmlyc3RDaGlsZC5pZCA9ICdjb250ZW50VHlwZUNhcmRUaXRsZSc7CiAgICBjb250ZW50VHlwZUNhcmQuYXBwZW5kQ2hpbGQoY29udGVudFR5cGVIZWFkZXIpOwogICAgY29uc3QgY29udGVudFR5cGVXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjb250ZW50VHlwZVdyYXAuY2xhc3NOYW1lID0gJ2NoYXJ0LXdyYXAnOwogICAgY29udGVudFR5cGVXcmFwLmlkID0gJ2NvbnRlbnRUeXBlQ2hhcnRXcmFwJzsKICAgIGNvbnRlbnRUeXBlV3JhcC5pbm5lckhUTUwgPSAnPGNhbnZhcyBpZD0iY29udGVudFR5cGVDYW52YXMiPjwvY2FudmFzPic7CiAgICBjb250ZW50VHlwZUNhcmQuYXBwZW5kQ2hpbGQoY29udGVudFR5cGVXcmFwKTsKCiAgICBncmlkLmFwcGVuZChicmVha2Rvd25DYXJkLCBjb250ZW50VHlwZUNhcmQpOwogICAgcm9vdC5hcHBlbmRDaGlsZChncmlkKTsKCiAgICBjb25zdCB0b3BUaXRsZSA9IHRleHRFbCgnZGl2JywgJ1RvcC1wZXJmb3JtaW5nIHBvc3RzJywgJ3NlY3Rpb24tdGl0bGUnKTsKICAgIGNvbnN0IHRvcENhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRvcENhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgY29uc3QgdG9wSGVhZGVyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0b3BIZWFkZXIuY2xhc3NOYW1lID0gJ2NhcmQtaGVhZGVyJzsKICAgIHRvcEhlYWRlci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2gzJywgJ1JhbmtlZCBieSBzZWxlY3RlZCBtZXRyaWMnKSk7CiAgICB0b3BDYXJkLmFwcGVuZENoaWxkKHRvcEhlYWRlcik7CiAgICBjb25zdCB0YWJsZVdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRhYmxlV3JhcC5jbGFzc05hbWUgPSAndGFibGUtc2Nyb2xsJzsKICAgIHRhYmxlV3JhcC5pZCA9ICd0b3BQb3N0c1RhYmxlJzsKICAgIHRvcENhcmQuYXBwZW5kQ2hpbGQodGFibGVXcmFwKTsKICAgIHJvb3QuYXBwZW5kKHRvcFRpdGxlLCB0b3BDYXJkKTsKICB9CgogIGZ1bmN0aW9uIHBvc3RUaWxlKGxhYmVsLCBwb3N0KSB7CiAgICBjb25zdCB0aWxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0aWxlLmNsYXNzTmFtZSA9ICdzdGF0LXRpbGUgcG9zdC10aWxlJzsKICAgIHRpbGUuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBsYWJlbCwgJ3N0YXQtbGFiZWwnKSk7CiAgICBpZiAoIXBvc3QpIHsKICAgICAgdGlsZS5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsICdObyBkYXRhIHlldCcsICdzdGF0LXZhbHVlJykpOwogICAgICByZXR1cm4gdGlsZTsKICAgIH0KICAgIGNvbnN0IHBsYXRmb3JtT3B0aW9ucyA9ICh3aW5kb3cuX19maWx0ZXJPcHRpb25zQ2FjaGUgfHwgeyBwbGF0Zm9ybXM6IFtdIH0pLnBsYXRmb3JtczsKICAgIGNvbnN0IG1ldGEgPSBwbGF0Zm9ybU9wdGlvbnMuZmluZCgocCkgPT4gcC5pZCA9PT0gcG9zdC5wbGF0Zm9ybSkgfHwgeyBsYWJlbDogcG9zdC5wbGF0Zm9ybSB9OwogICAgY29uc3QgY2FwdGlvbiA9IHBvc3QuY2FwdGlvbiB8fCAnKG5vIGNhcHRpb24pJzsKICAgIGNvbnN0IHZhbHVlRWwgPSB0ZXh0RWwoJ2RpdicsIGNhcHRpb24ubGVuZ3RoID4gNzAgPyBgJHtjYXB0aW9uLnNsaWNlKDAsIDcwKX3igKZgIDogY2FwdGlvbiwgJ3N0YXQtdmFsdWUnKTsKICAgIHZhbHVlRWwudGl0bGUgPSBjYXB0aW9uOwogICAgdGlsZS5hcHBlbmRDaGlsZCh2YWx1ZUVsKTsKICAgIHRpbGUuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBgJHttZXRhLmxhYmVsfSDCtyAke0Zvcm1hdC5kYXRlKHBvc3QucHVibGlzaF9kYXRlKX1gLCAncG9zdC1tZXRhJykpOwogICAgdGlsZS5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIGZvcm1hdE1ldHJpY1ZhbHVlKG1ldHJpYywgcG9zdC52YWx1ZSksICdwb3N0LW1ldHJpYy12YWx1ZScpKTsKICAgIHJldHVybiB0aWxlOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyS3BpcyhzdW1tYXJ5KSB7CiAgICBjb25zdCBncmlkID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2twaUdyaWQnKTsKICAgIGlmICghZ3JpZCkgcmV0dXJuOwogICAgZ3JpZC5pbm5lckhUTUwgPSAnJzsKCiAgICBjb25zdCBzdGF0VGlsZSA9IChsYWJlbCwgdmFsdWUpID0+IHsKICAgICAgY29uc3QgdGlsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICB0aWxlLmNsYXNzTmFtZSA9ICdzdGF0LXRpbGUnOwogICAgICB0aWxlLmFwcGVuZCh0ZXh0RWwoJ2RpdicsIGxhYmVsLCAnc3RhdC1sYWJlbCcpLCB0ZXh0RWwoJ2RpdicsIHZhbHVlLCAnc3RhdC12YWx1ZScpKTsKICAgICAgcmV0dXJuIHRpbGU7CiAgICB9OwoKICAgIGdyaWQuYXBwZW5kQ2hpbGQoc3RhdFRpbGUoJ0hpZ2hlc3QgVmFsdWUnLCBmb3JtYXRNZXRyaWNWYWx1ZShtZXRyaWMsIHN1bW1hcnkuaGlnaGVzdCkpKTsKICAgIGdyaWQuYXBwZW5kQ2hpbGQoc3RhdFRpbGUoJ0F2ZXJhZ2UgVmFsdWUnLCBmb3JtYXRNZXRyaWNWYWx1ZShtZXRyaWMsIHN1bW1hcnkuYXZlcmFnZSkpKTsKICAgIGlmIChzdW1tYXJ5LnVuaXQgIT09ICdwZXJjZW50JykgewogICAgICBncmlkLmFwcGVuZENoaWxkKHN0YXRUaWxlKCdUb3RhbCBWYWx1ZScsIGZvcm1hdE1ldHJpY1ZhbHVlKG1ldHJpYywgc3VtbWFyeS50b3RhbCkpKTsKICAgIH0KICAgIGdyaWQuYXBwZW5kQ2hpbGQoc3RhdFRpbGUoJ051bWJlciBvZiBQb3N0cycsIEZvcm1hdC5udW1iZXIoc3VtbWFyeS5wb3N0Q291bnQpKSk7CiAgICBncmlkLmFwcGVuZENoaWxkKHBvc3RUaWxlKCdCZXN0IFBlcmZvcm1pbmcgUG9zdCcsIHN1bW1hcnkuYmVzdFBvc3QpKTsKICAgIGdyaWQuYXBwZW5kQ2hpbGQocG9zdFRpbGUoJ0xvd2VzdCBQZXJmb3JtaW5nIFBvc3QnLCBzdW1tYXJ5LndvcnN0UG9zdCkpOwogIH0KCiAgLyoqIFN3YXBzIGEgY2hhcnQgY2FyZCdzIGNhbnZhcyBmb3IgYW4gZW1wdHktc3RhdGUgbWVzc2FnZSwgb3IgcmVzdG9yZXMgdGhlIGNhbnZhcyDigJQgc2luY2UKICAgICAgcmUtcmVuZGVyaW5nIGEgQ2hhcnQuanMgaW5zdGFuY2UgbmVlZHMgYSBsaXZlIDxjYW52YXM+LCBub3Qgd2hhdGV2ZXIgdGhlIGxhc3QgcmVuZGVyIGxlZnQgdGhlcmUuICovCiAgZnVuY3Rpb24gY2hhcnRPckVtcHR5KHdyYXBJZCwgY2FudmFzSWQsIGhhc0RhdGEsIGVtcHR5TWVzc2FnZSwgcmVuZGVyRm4pIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCh3cmFwSWQpOwogICAgaWYgKCF3cmFwKSByZXR1cm47CiAgICBDaGFydHMuZGVzdHJveShjYW52YXNJZCk7CiAgICBpZiAoIWhhc0RhdGEpIHsKICAgICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgICAgd3JhcC5hcHBlbmRDaGlsZChlbXB0eVN0YXRlKHsgaWNvbjogJ2Jhci1jaGFydC0zJywgbWVzc2FnZTogZW1wdHlNZXNzYWdlIH0pKTsKICAgICAgcmV0dXJuOwogICAgfQogICAgd3JhcC5pbm5lckhUTUwgPSBgPGNhbnZhcyBpZD0iJHtjYW52YXNJZH0iPjwvY2FudmFzPmA7CiAgICByZW5kZXJGbigpOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gcmVuZGVyVHJlbmQoZmlsdGVycykgewogICAgY29uc3QgcGxhdGZvcm1PcHRpb25zID0gKHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZSB8fCB7IHBsYXRmb3JtczogW10gfSkucGxhdGZvcm1zOwogICAgY29uc3QgbUxhYmVsID0gbWV0cmljTGFiZWwobWV0cmljKTsKICAgIGNvbnN0IHBsYXRmb3Jtc1RvRmV0Y2ggPSBmaWx0ZXJzLnBsYXRmb3JtID09PSAnYWxsJyA/IHBsYXRmb3JtT3B0aW9ucy5tYXAoKHApID0+IHAuaWQpIDogW2ZpbHRlcnMucGxhdGZvcm1dOwogICAgY29uc3QgdHJlbmRSZXNwb25zZXMgPSBhd2FpdCBQcm9taXNlLmFsbCgKICAgICAgcGxhdGZvcm1zVG9GZXRjaC5tYXAoKHApID0+CiAgICAgICAgQXBpLnRyZW5kKHsgZGF0ZUZyb206IGZpbHRlcnMuZGF0ZUZyb20sIGRhdGVUbzogZmlsdGVycy5kYXRlVG8sIHBsYXRmb3JtOiBwLCBjYW1wYWlnblR5cGU6IGZpbHRlcnMuY2FtcGFpZ25UeXBlLCBjb250ZW50VHlwZTogZmlsdGVycy5jb250ZW50VHlwZSB9KQogICAgICApCiAgICApOwogICAgY29uc3Qgd2Vla1NldCA9IG5ldyBTZXQoKTsKICAgIHRyZW5kUmVzcG9uc2VzLmZvckVhY2goKHJvd3MpID0+IHJvd3MuZm9yRWFjaCgocikgPT4gd2Vla1NldC5hZGQoci5wZXJpb2QpKSk7CiAgICBjb25zdCB3ZWVrcyA9IFsuLi53ZWVrU2V0XS5zb3J0KCk7CiAgICBjb25zdCBzZXJpZXMgPSBwbGF0Zm9ybXNUb0ZldGNoLm1hcCgocCwgaSkgPT4gewogICAgICBjb25zdCBtZXRhID0gcGxhdGZvcm1PcHRpb25zLmZpbmQoKHBsKSA9PiBwbC5pZCA9PT0gcCkgfHwgeyBsYWJlbDogcCB9OwogICAgICBjb25zdCBieVdlZWsgPSBPYmplY3QuZnJvbUVudHJpZXModHJlbmRSZXNwb25zZXNbaV0ubWFwKChyKSA9PiBbci5wZXJpb2QsIHJbbWV0cmljXV0pKTsKICAgICAgcmV0dXJuIHsgbGFiZWw6IG1ldGEubGFiZWwsIGNvbG9yOiBtZXRhLmNvbG9yLCBkYXRhOiB3ZWVrcy5tYXAoKHcpID0+IChieVdlZWtbd10gPT09IHVuZGVmaW5lZCA/IG51bGwgOiBieVdlZWtbd10pKSB9OwogICAgfSk7CgogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3RyZW5kQ2FyZFRpdGxlJykudGV4dENvbnRlbnQgPQogICAgICBmaWx0ZXJzLnBsYXRmb3JtID09PSAnYWxsJyA/IGBXZWVrbHkgJHttTGFiZWx9IGJ5IFBsYXRmb3JtYCA6IGAke21MYWJlbH0gVHJlbmRgOwoKICAgIGNoYXJ0T3JFbXB0eSgndHJlbmRDaGFydFdyYXAnLCAndHJlbmRDYW52YXMnLCB3ZWVrcy5sZW5ndGggPiAwLCAnTm8gZGF0YSBpbiB0aGlzIHJhbmdlIHlldC4nLCAoKSA9PiB7CiAgICAgIENoYXJ0cy50cmVuZENoYXJ0KCd0cmVuZENhbnZhcycsIHsgbGFiZWxzOiB3ZWVrcy5tYXAoRm9ybWF0LmRhdGUpLCBzZXJpZXMsIGZvcm1hdFZhbHVlOiAodikgPT4gZm9ybWF0TWV0cmljVmFsdWUobWV0cmljLCB2KSB9KTsKICAgIH0pOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gcmVuZGVyQnJlYWtkb3duKGZpbHRlcnMpIHsKICAgIGNvbnN0IG1MYWJlbCA9IG1ldHJpY0xhYmVsKG1ldHJpYyk7CiAgICBjb25zdCB0aXRsZUVsID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2JyZWFrZG93bkNhcmRUaXRsZScpOwoKICAgIGlmIChmaWx0ZXJzLnBsYXRmb3JtID09PSAnYWxsJykgewogICAgICB0aXRsZUVsLnRleHRDb250ZW50ID0gYFBsYXRmb3JtIENvbXBhcmlzb24g4oCUICR7bUxhYmVsfWA7CiAgICAgIGNvbnN0IGJyZWFrZG93biA9IGF3YWl0IEFwaS5wbGF0Zm9ybUJyZWFrZG93bihmaWx0ZXJzKTsKICAgICAgY29uc3Qgc29ydGVkID0gYnJlYWtkb3duLmZpbHRlcigocCkgPT4gcFttZXRyaWNdICE9PSBudWxsICYmIHBbbWV0cmljXSAhPT0gdW5kZWZpbmVkKS5zb3J0KChhLCBiKSA9PiBiW21ldHJpY10gLSBhW21ldHJpY10pOwogICAgICBjaGFydE9yRW1wdHkoJ2JyZWFrZG93bkNoYXJ0V3JhcCcsICdicmVha2Rvd25DYW52YXMnLCBzb3J0ZWQubGVuZ3RoID4gMCwgJ05vIGRhdGEgaW4gdGhpcyByYW5nZSB5ZXQuJywgKCkgPT4gewogICAgICAgIENoYXJ0cy5wbGF0Zm9ybUJhckNoYXJ0KCdicmVha2Rvd25DYW52YXMnLCB7CiAgICAgICAgICBsYWJlbHM6IHNvcnRlZC5tYXAoKHApID0+IHAubGFiZWwpLAogICAgICAgICAgZGF0YTogc29ydGVkLm1hcCgocCkgPT4gcFttZXRyaWNdKSwKICAgICAgICAgIGNvbG9yczogc29ydGVkLm1hcCgocCkgPT4gcC5jb2xvciksCiAgICAgICAgICBmb3JtYXRWYWx1ZTogKHYpID0+IGZvcm1hdE1ldHJpY1ZhbHVlKG1ldHJpYywgdiksCiAgICAgICAgfSk7CiAgICAgIH0pOwogICAgfSBlbHNlIHsKICAgICAgdGl0bGVFbC50ZXh0Q29udGVudCA9IGBDYW1wYWlnbiBQZXJmb3JtYW5jZSDigJQgJHttTGFiZWx9YDsKICAgICAgY29uc3QgY2FtcGFpZ25zID0gYXdhaXQgQXBpLmNhbXBhaWduQnJlYWtkb3duKGZpbHRlcnMpOwogICAgICBjb25zdCB3aXRoVmFsdWUgPSBjYW1wYWlnbnMuZmlsdGVyKChjKSA9PiBjW21ldHJpY10gIT09IG51bGwgJiYgY1ttZXRyaWNdICE9PSB1bmRlZmluZWQgJiYgY1ttZXRyaWNdID4gMCk7CiAgICAgIGNoYXJ0T3JFbXB0eSgnYnJlYWtkb3duQ2hhcnRXcmFwJywgJ2JyZWFrZG93bkNhbnZhcycsIHdpdGhWYWx1ZS5sZW5ndGggPiAwLCAnTm8gY2FtcGFpZ24gZGF0YSBpbiB0aGlzIHJhbmdlIHlldC4nLCAoKSA9PiB7CiAgICAgICAgQ2hhcnRzLnBpZUNoYXJ0KCdicmVha2Rvd25DYW52YXMnLCB7CiAgICAgICAgICBsYWJlbHM6IHdpdGhWYWx1ZS5tYXAoKGMpID0+IGMuY2FtcGFpZ25fdHlwZSksCiAgICAgICAgICBkYXRhOiB3aXRoVmFsdWUubWFwKChjKSA9PiBjW21ldHJpY10pLAogICAgICAgICAgY29sb3JzOiB3aXRoVmFsdWUubWFwKChfLCBpKSA9PiBDaGFydHMuc2VyaWVzQ29sb3IoaSkpLAogICAgICAgICAgZm9ybWF0VmFsdWU6ICh2KSA9PiBmb3JtYXRNZXRyaWNWYWx1ZShtZXRyaWMsIHYpLAogICAgICAgIH0pOwogICAgICB9KTsKICAgIH0KICB9CgogIGFzeW5jIGZ1bmN0aW9uIHJlbmRlckNvbnRlbnRUeXBlQnJlYWtkb3duKGZpbHRlcnMpIHsKICAgIGNvbnN0IG1MYWJlbCA9IG1ldHJpY0xhYmVsKG1ldHJpYyk7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnY29udGVudFR5cGVDYXJkVGl0bGUnKS50ZXh0Q29udGVudCA9IGBDb250ZW50IFR5cGUgUGVyZm9ybWFuY2Ug4oCUICR7bUxhYmVsfWA7CiAgICBjb25zdCByb3dzID0gYXdhaXQgQXBpLmNvbnRlbnRUeXBlQnJlYWtkb3duKGZpbHRlcnMpOwogICAgY29uc3Qgc29ydGVkID0gcm93cy5maWx0ZXIoKGMpID0+IGNbbWV0cmljXSAhPT0gbnVsbCAmJiBjW21ldHJpY10gIT09IHVuZGVmaW5lZCkuc29ydCgoYSwgYikgPT4gYlttZXRyaWNdIC0gYVttZXRyaWNdKTsKICAgIGNoYXJ0T3JFbXB0eSgnY29udGVudFR5cGVDaGFydFdyYXAnLCAnY29udGVudFR5cGVDYW52YXMnLCBzb3J0ZWQubGVuZ3RoID4gMCwgJ05vIGRhdGEgaW4gdGhpcyByYW5nZSB5ZXQuJywgKCkgPT4gewogICAgICBDaGFydHMucGxhdGZvcm1CYXJDaGFydCgnY29udGVudFR5cGVDYW52YXMnLCB7CiAgICAgICAgbGFiZWxzOiBzb3J0ZWQubWFwKChjKSA9PiBjLmNvbnRlbnRfdHlwZSksCiAgICAgICAgZGF0YTogc29ydGVkLm1hcCgoYykgPT4gY1ttZXRyaWNdKSwKICAgICAgICBjb2xvcnM6IHNvcnRlZC5tYXAoKF8sIGkpID0+IENoYXJ0cy5zZXJpZXNDb2xvcihpKSksCiAgICAgICAgZm9ybWF0VmFsdWU6ICh2KSA9PiBmb3JtYXRNZXRyaWNWYWx1ZShtZXRyaWMsIHYpLAogICAgICB9KTsKICAgIH0pOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gcmVuZGVyVG9wUG9zdHMoZmlsdGVycykgewogICAgY29uc3QgcG9zdHMgPSBhd2FpdCBBcGkudG9wUG9zdHMoeyAuLi5maWx0ZXJzLCBzb3J0Qnk6IG1ldHJpYywgbGltaXQ6IDEwIH0pOwogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd0b3BQb3N0c1RhYmxlJyk7CiAgICBpZiAoIXdyYXApIHJldHVybjsKICAgIGlmICghcG9zdHMubGVuZ3RoKSB7CiAgICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQoZW1wdHlTdGF0ZSh7CiAgICAgICAgaWNvbjogJ3Ryb3BoeScsCiAgICAgICAgdGl0bGU6ICdObyBwb3N0cyBpbiB0aGlzIHJhbmdlIHlldCcsCiAgICAgICAgbWVzc2FnZTogJ1VwbG9hZCBhIHdlZWtseSBleHBvcnQsIG9yIHdpZGVuIHRoZSBkYXRlIHJhbmdlLCB0byBzZWUgdG9wIHBlcmZvcm1lcnMgaGVyZS4nLAogICAgICAgIGFjdGlvbkxhYmVsOiAnVXBsb2FkIGRhdGEnLAogICAgICAgIG9uQWN0aW9uOiAoKSA9PiBkb2N1bWVudC5xdWVyeVNlbGVjdG9yKCcudGFiLWJ0bltkYXRhLXRhYj0idXBsb2FkIl0nKT8uY2xpY2soKSwKICAgICAgfSkpOwogICAgICByZXR1cm47CiAgICB9CiAgICBjb25zdCBwbGF0Zm9ybU9wdGlvbnMgPSAod2luZG93Ll9fZmlsdGVyT3B0aW9uc0NhY2hlIHx8IHsgcGxhdGZvcm1zOiBbXSB9KS5wbGF0Zm9ybXM7CiAgICBjb25zdCBzaG93RW5nYWdlbWVudFJhdGUgPSBtZXRyaWMgIT09ICdlbmdhZ2VtZW50X3JhdGUnICYmIHBvc3RzLnNvbWUoKHApID0+IHAuZW5nYWdlbWVudF9yYXRlICE9PSBudWxsICYmIHAuZW5nYWdlbWVudF9yYXRlICE9PSB1bmRlZmluZWQpOwoKICAgIGNvbnN0IHRhYmxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGFibGUnKTsKICAgIHRhYmxlLmNsYXNzTmFtZSA9ICdkYXRhLXRhYmxlJzsKICAgIGNvbnN0IHRoZWFkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGhlYWQnKTsKICAgIGNvbnN0IGhlYWRUciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICBoZWFkVHIuYXBwZW5kKAogICAgICB0ZXh0RWwoJ3RoJywgJ1JhbmsnKSwKICAgICAgdGV4dEVsKCd0aCcsICdEYXRlJyksCiAgICAgIHRleHRFbCgndGgnLCAnUGxhdGZvcm0nKSwKICAgICAgdGV4dEVsKCd0aCcsICdDYW1wYWlnbicpLAogICAgICB0ZXh0RWwoJ3RoJywgJ0NvbnRlbnQgVHlwZScpLAogICAgICB0ZXh0RWwoJ3RoJywgJ0NhcHRpb24nKSwKICAgICAgdGV4dEVsKCd0aCcsIG1ldHJpY0xhYmVsKG1ldHJpYyksICdudW0nKQogICAgKTsKICAgIGlmIChzaG93RW5nYWdlbWVudFJhdGUpIGhlYWRUci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3RoJywgJ0VuZ2FnZW1lbnQgUmF0ZScsICdudW0nKSk7CiAgICBoZWFkVHIuYXBwZW5kQ2hpbGQodGV4dEVsKCd0aCcsICcnKSk7CiAgICB0aGVhZC5hcHBlbmRDaGlsZChoZWFkVHIpOwoKICAgIGNvbnN0IHRib2R5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGJvZHknKTsKICAgIHBvc3RzLmZvckVhY2goKHAsIGkpID0+IHsKICAgICAgY29uc3QgdHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgICBjb25zdCBtZXRhID0gcGxhdGZvcm1PcHRpb25zLmZpbmQoKHBsKSA9PiBwbC5pZCA9PT0gcC5wbGF0Zm9ybSkgfHwgeyBsYWJlbDogcC5wbGF0Zm9ybSwgY29sb3I6ICcjOTk5JyB9OwogICAgICBjb25zdCBwbGF0Zm9ybVRkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgY29uc3QgcGlsbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgICAgcGlsbC5jbGFzc05hbWUgPSAncGxhdGZvcm0tcGlsbCc7CiAgICAgIGNvbnN0IGRvdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgICAgZG90LmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1kb3QnOwogICAgICBkb3Quc3R5bGUuYmFja2dyb3VuZCA9IG1ldGEuY29sb3I7CiAgICAgIHBpbGwuYXBwZW5kKGRvdCwgZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUobWV0YS5sYWJlbCkpOwogICAgICBwbGF0Zm9ybVRkLmFwcGVuZENoaWxkKHBpbGwpOwoKICAgICAgY29uc3QgY2FwdGlvbiA9IHAuY2FwdGlvbiB8fCAnKG5vIGNhcHRpb24pJzsKICAgICAgY29uc3QgY2FwdGlvblRkID0gdGV4dEVsKCd0ZCcsIGNhcHRpb24ubGVuZ3RoID4gNjAgPyBgJHtjYXB0aW9uLnNsaWNlKDAsIDYwKX3igKZgIDogY2FwdGlvbik7CiAgICAgIGNhcHRpb25UZC50aXRsZSA9IGNhcHRpb247CgogICAgICB0ci5hcHBlbmQoCiAgICAgICAgdGV4dEVsKCd0ZCcsIGAjJHtpICsgMX1gKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgRm9ybWF0LmRhdGUocC5wdWJsaXNoX2RhdGUpKSwKICAgICAgICBwbGF0Zm9ybVRkLAogICAgICAgIHRleHRFbCgndGQnLCBwLmNhbXBhaWduX3R5cGUgfHwgJ+KAlCcpLAogICAgICAgIHRleHRFbCgndGQnLCBwLmNvbnRlbnRfdHlwZSB8fCAn4oCUJyksCiAgICAgICAgY2FwdGlvblRkLAogICAgICAgIHRleHRFbCgndGQnLCBmb3JtYXRNZXRyaWNWYWx1ZShtZXRyaWMsIHAubWV0cmljX3ZhbHVlKSwgJ251bScpCiAgICAgICk7CiAgICAgIGlmIChzaG93RW5nYWdlbWVudFJhdGUpIHsKICAgICAgICB0ci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3RkJywgcC5lbmdhZ2VtZW50X3JhdGUgPT09IG51bGwgfHwgcC5lbmdhZ2VtZW50X3JhdGUgPT09IHVuZGVmaW5lZCA/ICfigJQnIDogRm9ybWF0LnBlcmNlbnQocC5lbmdhZ2VtZW50X3JhdGUpLCAnbnVtJykpOwogICAgICB9CgogICAgICBjb25zdCBhY3Rpb25UZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIGNvbnN0IHZpZXdCdG4gPSBpY29uQnRuKCdidG4nLCAnZXllJywgJ1ZpZXcgRGV0YWlscycpOwogICAgICB2aWV3QnRuLmRpc2FibGVkID0gIXAucmF3X3Jvd19pZDsKICAgICAgdmlld0J0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IFJlY29yZHMub3BlblZpZXcocC5yYXdfcm93X2lkKSk7CiAgICAgIGFjdGlvblRkLmFwcGVuZENoaWxkKHZpZXdCdG4pOwogICAgICB0ci5hcHBlbmRDaGlsZChhY3Rpb25UZCk7CgogICAgICB0Ym9keS5hcHBlbmRDaGlsZCh0cik7CiAgICB9KTsKICAgIHRhYmxlLmFwcGVuZCh0aGVhZCwgdGJvZHkpOwogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGFibGUpOwogIH0KCiAgLyoqIE1ldHJpYyAob3IgYW55IGZpbHRlcikgY2hhbmdlZCBidXQgdGhlIHBsYXRmb3JtIOKAlCBhbmQgdGhlcmVmb3JlIHRoZSBhdmFpbGFibGUgbWV0cmljIGxpc3Qg4oCUIGRpZG4ndDogbm8gbmVlZCB0byByZS1mZXRjaCBtZXRyaWMtb3B0aW9ucyBvciByZWJ1aWxkIHRoZSBzaGVsbCwganVzdCByZWZyZXNoIHRoZSBkYXRhLiAqLwogIGFzeW5jIGZ1bmN0aW9uIHJlZnJlc2hGb3JNZXRyaWMoKSB7CiAgICBjb25zdCBmaWx0ZXJzID0gU3RhdGUuZ2V0RmlsdGVycygpOwogICAgY29uc3Qgc3VtbWFyeSA9IGF3YWl0IEFwaS5tZXRyaWNTdW1tYXJ5KHsgLi4uZmlsdGVycywgbWV0cmljIH0pOwogICAgcmVuZGVyS3BpcyhzdW1tYXJ5KTsKICAgIGF3YWl0IFByb21pc2UuYWxsKFtyZW5kZXJUcmVuZChmaWx0ZXJzKSwgcmVuZGVyQnJlYWtkb3duKGZpbHRlcnMpLCByZW5kZXJDb250ZW50VHlwZUJyZWFrZG93bihmaWx0ZXJzKSwgcmVuZGVyVG9wUG9zdHMoZmlsdGVycyldKTsKICB9CgogIGZ1bmN0aW9uIHNob3dTa2VsZXRvbnMoKSB7CiAgICBjb25zdCBrcGlHcmlkID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2twaUdyaWQnKTsKICAgIGlmIChrcGlHcmlkKSB7IGtwaUdyaWQuaW5uZXJIVE1MID0gJyc7IGtwaUdyaWQuYXBwZW5kQ2hpbGQoc2tlbGV0b25TdGF0R3JpZCg2KSk7IH0KICAgIFsndHJlbmRDaGFydFdyYXAnLCAnYnJlYWtkb3duQ2hhcnRXcmFwJywgJ2NvbnRlbnRUeXBlQ2hhcnRXcmFwJ10uZm9yRWFjaCgoaWQpID0+IHsKICAgICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGlkKTsKICAgICAgaWYgKHdyYXApIHsgd3JhcC5pbm5lckhUTUwgPSAnJzsgd3JhcC5hcHBlbmRDaGlsZChza2VsZXRvbkNoYXJ0KCkpOyB9CiAgICB9KTsKICAgIGNvbnN0IHRvcFBvc3RzID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3RvcFBvc3RzVGFibGUnKTsKICAgIGlmICh0b3BQb3N0cykgeyB0b3BQb3N0cy5pbm5lckhUTUwgPSAnJzsgdG9wUG9zdHMuYXBwZW5kQ2hpbGQoc2tlbGV0b25Sb3dzKDYpKTsgfQogIH0KCiAgYXN5bmMgZnVuY3Rpb24gcmVuZGVyKCkgewogICAgcm9vdCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd2aWV3LWRhc2hib2FyZCcpOwogICAgY29uc3QgZmlsdGVycyA9IFN0YXRlLmdldEZpbHRlcnMoKTsKICAgIGNvbnN0IHsgb3B0aW9ucyB9ID0gYXdhaXQgQXBpLm1ldHJpY09wdGlvbnMoZmlsdGVycy5wbGF0Zm9ybSk7CiAgICBtZXRyaWNPcHRpb25zID0gb3B0aW9uczsKICAgIGlmICghbWV0cmljT3B0aW9ucy5zb21lKChtKSA9PiBtLmtleSA9PT0gbWV0cmljKSkgewogICAgICBtZXRyaWMgPSBtZXRyaWNPcHRpb25zLmxlbmd0aCA/IG1ldHJpY09wdGlvbnNbMF0ua2V5IDogJ3ZpZXdzJzsKICAgIH0KICAgIHNoZWxsKCk7CiAgICBzaG93U2tlbGV0b25zKCk7CiAgICBhd2FpdCByZWZyZXNoRm9yTWV0cmljKCk7CiAgfQoKICByZXR1cm4geyByZW5kZXIgfTsKfSkoKTsKCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBEYXRhIFJlY29yZHMgdGFiOiBhIENSTS1zdHlsZSwgcGxhdGZvcm0tZ3JvdXBlZCBicm93c2VyIGJhY2tlZAogICBieSBwb3N0cy9wb3N0X21ldHJpY3MgKHRoZSBzYW1lIG5vcm1hbGl6ZWQgZGF0YSB0aGUgZGFzaGJvYXJkLAogICBjb21wYXJpc29ucywgYW5kIHJlcG9ydHMgcmVhZCkg4oCUICJBbGwgUGxhdGZvcm1zIiBzaG93cyBhIGNvbW1vbgogICBjcm9zcy1wbGF0Zm9ybSBzdW1tYXJ5LCBhIHNwZWNpZmljIHBsYXRmb3JtIHNob3dzIG9ubHkgdGhhdAogICBwbGF0Zm9ybSdzIGN1cmF0ZWQgbWV0cmljcy4gRXZlcnkgZmllbGQgb2YgYSByZWNvcmQgKGV4YWN0bHkgYXMKICAgaW1wb3J0ZWQpIGlzIGFsd2F5cyByZWFjaGFibGUgdmlhIFZpZXcvRWRpdCByZWdhcmRsZXNzIG9mIHRoZQogICB0YWJsZSdzIGN1cmF0aW9uLCB3aGljaCByZWFkcyB0aGUgcmF3X3Jvd3MgbWlycm9yIGFuZCwgb24gc2F2ZSwKICAgcmUtc3luY3MgcG9zdHMvcG9zdF9tZXRyaWNzIHNvIGV2ZXJ5IHZpZXcgcmVmbGVjdHMgdGhlIGNoYW5nZQogICBpbW1lZGlhdGVseS4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmNvbnN0IFJlY29yZHMgPSAoKCkgPT4gewogIGxldCByb290OwogIGxldCBwYWdlID0gMTsKICBjb25zdCBwYWdlU2l6ZSA9IDI1OwogIGxldCBzZWFyY2hWYWx1ZSA9ICcnOwogIGxldCBzZWFyY2hEZWJvdW5jZSA9IG51bGw7CiAgbGV0IG1vZGFsU3RhdGUgPSBudWxsOyAvLyB7IHJlY29yZCwgdmFsdWVzOiBbLi4uXSB9IOKAlCBFZGl0IG1vZGFsIG9ubHkKCiAgZnVuY3Rpb24gcGxhdGZvcm1NZXRhKCkgewogICAgcmV0dXJuICh3aW5kb3cuX19maWx0ZXJPcHRpb25zQ2FjaGUgfHwgeyBwbGF0Zm9ybXM6IFtdIH0pLnBsYXRmb3JtczsKICB9CgogIGZ1bmN0aW9uIHBsYXRmb3JtTGFiZWwoaWQpIHsKICAgIGNvbnN0IG0gPSBwbGF0Zm9ybU1ldGEoKS5maW5kKChwKSA9PiBwLmlkID09PSBpZCk7CiAgICByZXR1cm4gbSA/IG0ubGFiZWwgOiBpZDsKICB9CgogIGZ1bmN0aW9uIHNoZWxsKCkgewogICAgcm9vdC5pbm5lckhUTUwgPSAnJzsKICAgIHJvb3QuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCAnRGF0YSBSZWNvcmRzJywgJ3NlY3Rpb24tdGl0bGUnKSk7CiAgICByb290LmFwcGVuZENoaWxkKHRleHRFbCgKICAgICAgJ2RpdicsCiAgICAgICdCcm93c2UgYnkgcGxhdGZvcm0gdG8gc2VlIG9ubHkgaXRzIG1ldHJpY3MsIG9yIHN0YXkgb24gQWxsIFBsYXRmb3JtcyBmb3IgYSBjcm9zcy1wbGF0Zm9ybSBzdW1tYXJ5LiBFdmVyeSByZWNvcmQgc3RheXMgZnVsbHkgZWRpdGFibGUg4oCUIFZpZXcgb3IgRWRpdCBhbHdheXMgb3BlbnMgZXZlcnkgZmllbGQgaW1wb3J0ZWQgZnJvbSB0aGUgc3ByZWFkc2hlZXQsIG5vdCBqdXN0IHdoYXTigJlzIGluIHRoZSB0YWJsZS4nLAogICAgICAnbXV0ZWQnCiAgICApKTsKCiAgICBjb25zdCB0b29sYmFyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0b29sYmFyLmNsYXNzTmFtZSA9ICdyZWNvcmRzLXRvb2xiYXInOwogICAgY29uc3QgcGlsbHMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHBpbGxzLmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1maWx0ZXItcGlsbHMnOwogICAgcGlsbHMuaWQgPSAncmVjb3Jkc1BsYXRmb3JtUGlsbHMnOwogICAgY29uc3Qgc2VhcmNoID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBzZWFyY2guY2xhc3NOYW1lID0gJ3JlY29yZHMtc2VhcmNoJzsKICAgIGNvbnN0IHNlYXJjaElucHV0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsKICAgIHNlYXJjaElucHV0LnR5cGUgPSAnc2VhcmNoJzsKICAgIHNlYXJjaElucHV0LnBsYWNlaG9sZGVyID0gJ1NlYXJjaCBjYXB0aW9ucywgY2FtcGFpZ25zLCBjb250ZW50IHR5cGXigKYnOwogICAgc2VhcmNoSW5wdXQuaWQgPSAncmVjb3Jkc1NlYXJjaElucHV0JzsKICAgIHNlYXJjaElucHV0LnZhbHVlID0gc2VhcmNoVmFsdWU7CiAgICBzZWFyY2hJbnB1dC5hZGRFdmVudExpc3RlbmVyKCdpbnB1dCcsICgpID0+IHsKICAgICAgY2xlYXJUaW1lb3V0KHNlYXJjaERlYm91bmNlKTsKICAgICAgc2VhcmNoRGVib3VuY2UgPSBzZXRUaW1lb3V0KCgpID0+IHsKICAgICAgICBzZWFyY2hWYWx1ZSA9IHNlYXJjaElucHV0LnZhbHVlOwogICAgICAgIHBhZ2UgPSAxOwogICAgICAgIGxvYWQoKTsKICAgICAgfSwgMzAwKTsKICAgIH0pOwogICAgc2VhcmNoLmFwcGVuZENoaWxkKHNlYXJjaElucHV0KTsKICAgIHRvb2xiYXIuYXBwZW5kKHBpbGxzLCBzZWFyY2gpOwogICAgcm9vdC5hcHBlbmRDaGlsZCh0b29sYmFyKTsKCiAgICBjb25zdCBjYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjYXJkLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGNvbnN0IHRhYmxlV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGFibGVXcmFwLmNsYXNzTmFtZSA9ICd0YWJsZS1zY3JvbGwnOwogICAgdGFibGVXcmFwLmlkID0gJ3JlY29yZHNUYWJsZVdyYXAnOwogICAgY2FyZC5hcHBlbmRDaGlsZCh0YWJsZVdyYXApOwogICAgY29uc3QgcGFnZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHBhZ2VyLmNsYXNzTmFtZSA9ICdwYWdpbmF0aW9uLXJvdyc7CiAgICBwYWdlci5pZCA9ICdyZWNvcmRzUGFnZXInOwogICAgY2FyZC5hcHBlbmRDaGlsZChwYWdlcik7CiAgICByb290LmFwcGVuZENoaWxkKGNhcmQpOwoKICAgIHJlbmRlclBpbGxzKCk7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJQaWxscygpIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncmVjb3Jkc1BsYXRmb3JtUGlsbHMnKTsKICAgIGlmICghd3JhcCkgcmV0dXJuOwogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgIGNvbnN0IGN1cnJlbnQgPSBTdGF0ZS5nZXRGaWx0ZXJzKCkucGxhdGZvcm0gfHwgJ2FsbCc7CiAgICBjb25zdCBvcHRpb25zID0gW3sgaWQ6ICdhbGwnLCBsYWJlbDogJ0FsbCBQbGF0Zm9ybXMnLCBjb2xvcjogbnVsbCB9LCAuLi5wbGF0Zm9ybU1ldGEoKV07CiAgICBvcHRpb25zLmZvckVhY2goKG9wdCkgPT4gewogICAgICBjb25zdCBidG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgICAgYnRuLnR5cGUgPSAnYnV0dG9uJzsKICAgICAgYnRuLmNsYXNzTGlzdC50b2dnbGUoJ2lzLWFjdGl2ZScsIGN1cnJlbnQgPT09IG9wdC5pZCk7CiAgICAgIGlmIChvcHQuY29sb3IpIHsKICAgICAgICBjb25zdCBkb3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICAgICAgZG90LmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1kb3QnOwogICAgICAgIGRvdC5zdHlsZS5iYWNrZ3JvdW5kID0gb3B0LmNvbG9yOwogICAgICAgIGJ0bi5hcHBlbmRDaGlsZChkb3QpOwogICAgICB9CiAgICAgIGJ0bi5hcHBlbmRDaGlsZChkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShvcHQubGFiZWwpKTsKICAgICAgYnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gewogICAgICAgIGlmIChjdXJyZW50ID09PSBvcHQuaWQpIHJldHVybjsKICAgICAgICBjb25zdCBmaWx0ZXJTZWxlY3QgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmlsdGVyUGxhdGZvcm0nKTsKICAgICAgICBpZiAoZmlsdGVyU2VsZWN0KSBmaWx0ZXJTZWxlY3QudmFsdWUgPSBvcHQuaWQ7CiAgICAgICAgcGFnZSA9IDE7CiAgICAgICAgU3RhdGUuc2V0RmlsdGVycyh7IHBsYXRmb3JtOiBvcHQuaWQgfSk7CiAgICAgIH0pOwogICAgICB3cmFwLmFwcGVuZENoaWxkKGJ0bik7CiAgICB9KTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIGxvYWQoKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3JlY29yZHNUYWJsZVdyYXAnKTsKICAgIGlmICh3cmFwKSB7IHdyYXAuaW5uZXJIVE1MID0gJyc7IHdyYXAuYXBwZW5kQ2hpbGQoc2tlbGV0b25Sb3dzKDgpKTsgfQogICAgY29uc3QgZmlsdGVycyA9IFN0YXRlLmdldEZpbHRlcnMoKTsKICAgIGNvbnN0IHJlc3VsdCA9IGF3YWl0IEFwaS5yZWNvcmRzVGFibGUoeyAuLi5maWx0ZXJzLCBzZWFyY2g6IHNlYXJjaFZhbHVlLCBwYWdlLCBwYWdlU2l6ZSB9KTsKICAgIHJlbmRlclRhYmxlKHJlc3VsdCk7CiAgICByZW5kZXJQYWdlcihyZXN1bHQpOwogIH0KCiAgZnVuY3Rpb24gY29sdW1uTGFiZWxzRm9yKHJlY29yZCkgewogICAgcmV0dXJuIHJlY29yZC5oZWFkZXJzICYmIHJlY29yZC5oZWFkZXJzLmxlbmd0aAogICAgICA/IHJlY29yZC5oZWFkZXJzLm1hcCgoaCkgPT4gKGggJiYgaC50cmltKCkgPyBoIDogJyh1bmxhYmVsZWQgY29sdW1uKScpKQogICAgICA6IHJlY29yZC52YWx1ZXMubWFwKChfLCBpKSA9PiBgQ29sdW1uICR7aSArIDF9YCk7CiAgfQoKICAvKiogR3JvdXBzIGEgcmF3IHJlY29yZCdzIGZpZWxkcyBieSB0aGUgcXVhbGlmaWVkIGhlYWRlcidzIHBsYXRmb3JtLWdyb3VwIHByZWZpeAogICAgICAoZS5nLiAiRkFDRUJPT0sg4oCUIFZpZXdzIiksIHNvIHRoZSBWaWV3L0VkaXQgcG9wdXAgcmVhZHMgYXMgc2VjdGlvbnMgaW5zdGVhZAogICAgICBvZiBvbmUgbG9uZyBmbGF0IGxpc3Qg4oCUIGZhbGxzIGJhY2sgdG8gYSBzaW5nbGUgIkRldGFpbHMiIHNlY3Rpb24gZm9yCiAgICAgIGlkZW50aWZpZXIgY29sdW1ucyBhbmQgZm9yIHRoZSBzaW1wbGUgKG9uZS1wbGF0Zm9ybS1wZXItcm93KSBmb3JtYXQuICovCiAgZnVuY3Rpb24gZ3JvdXBGaWVsZFJvd3MobGFiZWxzLCB2YWx1ZXMpIHsKICAgIGNvbnN0IGdyb3VwcyA9IFtdOwogICAgY29uc3QgaW5kZXggPSBuZXcgTWFwKCk7CiAgICBsYWJlbHMuZm9yRWFjaCgobGFiZWwsIGlkeCkgPT4gewogICAgICBjb25zdCBzZXBJZHggPSBsYWJlbC5pbmRleE9mKCcg4oCUICcpOwogICAgICBjb25zdCBncm91cE5hbWUgPSBzZXBJZHggPj0gMCA/IGxhYmVsLnNsaWNlKDAsIHNlcElkeCkgOiAnRGV0YWlscyc7CiAgICAgIGNvbnN0IGZpZWxkTGFiZWwgPSBzZXBJZHggPj0gMCA/IGxhYmVsLnNsaWNlKHNlcElkeCArIDMpIDogbGFiZWw7CiAgICAgIGlmICghaW5kZXguaGFzKGdyb3VwTmFtZSkpIHsKICAgICAgICBpbmRleC5zZXQoZ3JvdXBOYW1lLCB7IGdyb3VwOiBncm91cE5hbWUsIGZpZWxkczogW10gfSk7CiAgICAgICAgZ3JvdXBzLnB1c2goaW5kZXguZ2V0KGdyb3VwTmFtZSkpOwogICAgICB9CiAgICAgIGluZGV4LmdldChncm91cE5hbWUpLmZpZWxkcy5wdXNoKHsgaWR4LCBsYWJlbDogZmllbGRMYWJlbCB8fCBgQ29sdW1uICR7aWR4ICsgMX1gLCB2YWx1ZTogdmFsdWVzW2lkeF0gfSk7CiAgICB9KTsKICAgIHJldHVybiBncm91cHM7CiAgfQoKICBmdW5jdGlvbiBwbGF0Zm9ybUJhZGdlcyhpZHMpIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHdyYXAuc3R5bGUuZGlzcGxheSA9ICdmbGV4JzsKICAgIHdyYXAuc3R5bGUuZmxleFdyYXAgPSAnd3JhcCc7CiAgICB3cmFwLnN0eWxlLmdhcCA9ICc0cHgnOwogICAgaWYgKCFpZHMubGVuZ3RoKSByZXR1cm4gdGV4dEVsKCdzcGFuJywgJ+KAlCcsICdtdXRlZCcpOwogICAgY29uc3QgbWV0YSA9IHBsYXRmb3JtTWV0YSgpOwogICAgaWRzLmZvckVhY2goKGlkKSA9PiB7CiAgICAgIGNvbnN0IG0gPSBtZXRhLmZpbmQoKHApID0+IHAuaWQgPT09IGlkKSB8fCB7IGxhYmVsOiBpZCwgY29sb3I6ICcjOTk5JyB9OwogICAgICBjb25zdCBwaWxsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICBwaWxsLmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1waWxsJzsKICAgICAgY29uc3QgZG90ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICBkb3QuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLWRvdCc7CiAgICAgIGRvdC5zdHlsZS5iYWNrZ3JvdW5kID0gbS5jb2xvcjsKICAgICAgcGlsbC5hcHBlbmQoZG90LCBkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShtLmxhYmVsKSk7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQocGlsbCk7CiAgICB9KTsKICAgIHJldHVybiB3cmFwOwogIH0KCiAgZnVuY3Rpb24gc3RhdHVzUGlsbChzdGF0dXMpIHsKICAgIGNvbnN0IHNwYW4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICBzcGFuLmNsYXNzTmFtZSA9IGBzdGF0dXMtcGlsbCAke3N0YXR1c31gOwogICAgc3Bhbi50ZXh0Q29udGVudCA9IHN0YXR1cyA9PT0gJ2VkaXRlZCcgPyAnRWRpdGVkJyA6ICdPcmlnaW5hbCc7CiAgICByZXR1cm4gc3BhbjsKICB9CgogIGZ1bmN0aW9uIG1ldHJpY0NlbGwoa2V5LCB2YWx1ZSkgewogICAgaWYgKGtleSA9PT0gJ3Bvc3RpbmdfbGluaycpIHsKICAgICAgY29uc3QgdGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICB0ZC5jbGFzc05hbWUgPSAnbGluay1jZWxsJzsKICAgICAgaWYgKHZhbHVlKSB7CiAgICAgICAgY29uc3QgYSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2EnKTsKICAgICAgICBhLmhyZWYgPSB2YWx1ZTsKICAgICAgICBhLnRhcmdldCA9ICdfYmxhbmsnOwogICAgICAgIGEucmVsID0gJ25vb3BlbmVyIG5vcmVmZXJyZXInOwogICAgICAgIGEudGV4dENvbnRlbnQgPSAnT3BlbiDihpcnOwogICAgICAgIHRkLmFwcGVuZENoaWxkKGEpOwogICAgICB9IGVsc2UgewogICAgICAgIHRkLmFwcGVuZENoaWxkKGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKCfigJQnKSk7CiAgICAgIH0KICAgICAgcmV0dXJuIHRkOwogICAgfQogICAgY29uc3QgZGlzcGxheSA9IGtleSA9PT0gJ3dhdGNoX3RpbWVfc2Vjb25kcycgPyBGb3JtYXQuZHVyYXRpb24odmFsdWUpIDogRm9ybWF0Lm51bWJlcih2YWx1ZSk7CiAgICByZXR1cm4gdGV4dEVsKCd0ZCcsIGRpc3BsYXksICdudW0nKTsKICB9CgogIGZ1bmN0aW9uIGFjdGlvbkJ1dHRvbnMocm93LCBwbGF0Zm9ybSkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgd3JhcC5jbGFzc05hbWUgPSAncm93LWFjdGlvbnMnOwogICAgY29uc3Qgdmlld0J0biA9IGljb25CdG4oJ2J0bicsICdleWUnLCAnVmlldycpOwogICAgdmlld0J0bi5kaXNhYmxlZCA9ICFyb3cucmF3Um93SWQ7CiAgICB2aWV3QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gb3BlblZpZXcocm93LnJhd1Jvd0lkKSk7CiAgICBjb25zdCBlZGl0QnRuID0gaWNvbkJ0bignYnRuJywgJ3BlbmNpbCcsICdFZGl0Jyk7CiAgICBlZGl0QnRuLmRpc2FibGVkID0gIXJvdy5yYXdSb3dJZDsKICAgIGVkaXRCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBvcGVuRWRpdG9yKHJvdy5yYXdSb3dJZCkpOwogICAgY29uc3QgZGVsZXRlQnRuID0gaWNvbkJ0bignYnRuIGRhbmdlcicsICd0cmFzaC0yJywgJ0RlbGV0ZScpOwogICAgZGVsZXRlQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gaGFuZGxlRGVsZXRlKHJvdywgcGxhdGZvcm0pKTsKICAgIHdyYXAuYXBwZW5kKHZpZXdCdG4sIGVkaXRCdG4sIGRlbGV0ZUJ0bik7CiAgICByZXR1cm4gd3JhcDsKICB9CgogIGZ1bmN0aW9uIGNhcHRpb25DZWxsKGNhcHRpb24pIHsKICAgIGNvbnN0IHRleHQgPSBjYXB0aW9uIHx8ICcobm8gY2FwdGlvbiknOwogICAgcmV0dXJuIHRleHRFbCgndGQnLCB0ZXh0Lmxlbmd0aCA+IDcwID8gYCR7dGV4dC5zbGljZSgwLCA3MCl94oCmYCA6IHRleHQpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyU3VtbWFyeVRhYmxlKHJlc3VsdCkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdyZWNvcmRzVGFibGVXcmFwJyk7CiAgICBpZiAoIXJlc3VsdC5yb3dzLmxlbmd0aCkgewogICAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgICB3cmFwLmFwcGVuZENoaWxkKGVtcHR5U3RhdGUoewogICAgICAgIGljb246ICdkYXRhYmFzZScsCiAgICAgICAgdGl0bGU6ICdObyByZWNvcmRzIG1hdGNoIHRoZXNlIGZpbHRlcnMgeWV0JywKICAgICAgICBtZXNzYWdlOiAnVXBsb2FkIGEgd2Vla2x5IGV4cG9ydCwgb3Igd2lkZW4gdGhlIGRhdGUgcmFuZ2UsIHRvIHNlZSByZWNvcmRzIGhlcmUuJywKICAgICAgICBhY3Rpb25MYWJlbDogJ1VwbG9hZCBkYXRhJywKICAgICAgICBvbkFjdGlvbjogKCkgPT4gZG9jdW1lbnQucXVlcnlTZWxlY3RvcignLnRhYi1idG5bZGF0YS10YWI9InVwbG9hZCJdJyk/LmNsaWNrKCksCiAgICAgIH0pKTsKICAgICAgcmV0dXJuOwogICAgfQogICAgY29uc3QgdGFibGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0YWJsZScpOwogICAgdGFibGUuY2xhc3NOYW1lID0gJ2RhdGEtdGFibGUnOwogICAgY29uc3QgdGhlYWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0aGVhZCcpOwogICAgdGhlYWQuaW5uZXJIVE1MID0gJzx0cj48dGg+RGF0ZTwvdGg+PHRoPlBsYXRmb3JtczwvdGg+PHRoPkNhcHRpb248L3RoPjx0aD5DYW1wYWlnbjwvdGg+PHRoPkNvbnRlbnQgVHlwZTwvdGg+PHRoPlN0YXR1czwvdGg+PHRoPkxhc3QgVXBkYXRlZDwvdGg+PHRoPkFjdGlvbnM8L3RoPjwvdHI+JzsKICAgIGNvbnN0IHRib2R5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGJvZHknKTsKICAgIHJlc3VsdC5yb3dzLmZvckVhY2goKHIpID0+IHsKICAgICAgY29uc3QgdHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgICBjb25zdCBwbGF0Zm9ybXNUZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIHBsYXRmb3Jtc1RkLmFwcGVuZENoaWxkKHBsYXRmb3JtQmFkZ2VzKHIucGxhdGZvcm1JZHMpKTsKICAgICAgY29uc3Qgc3RhdHVzVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBzdGF0dXNUZC5hcHBlbmRDaGlsZChzdGF0dXNQaWxsKHIuc3RhdHVzKSk7CiAgICAgIGNvbnN0IGFjdGlvbnNUZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIGFjdGlvbnNUZC5hcHBlbmRDaGlsZChhY3Rpb25CdXR0b25zKHIsICdhbGwnKSk7CiAgICAgIHRyLmFwcGVuZCgKICAgICAgICB0ZXh0RWwoJ3RkJywgRm9ybWF0LmRhdGUoci5wdWJsaXNoRGF0ZSkpLAogICAgICAgIHBsYXRmb3Jtc1RkLAogICAgICAgIGNhcHRpb25DZWxsKHIuY2FwdGlvbiksCiAgICAgICAgdGV4dEVsKCd0ZCcsIHIuY2FtcGFpZ25UeXBlIHx8ICfigJQnKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgci5jb250ZW50VHlwZSB8fCAn4oCUJyksCiAgICAgICAgc3RhdHVzVGQsCiAgICAgICAgdGV4dEVsKCd0ZCcsIHIudXBkYXRlZEF0KSwKICAgICAgICBhY3Rpb25zVGQKICAgICAgKTsKICAgICAgdGJvZHkuYXBwZW5kQ2hpbGQodHIpOwogICAgfSk7CiAgICB0YWJsZS5hcHBlbmQodGhlYWQsIHRib2R5KTsKICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CiAgICB3cmFwLmFwcGVuZENoaWxkKHRhYmxlKTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlclBsYXRmb3JtVGFibGUocmVzdWx0KSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3JlY29yZHNUYWJsZVdyYXAnKTsKICAgIGlmICghcmVzdWx0LnJvd3MubGVuZ3RoKSB7CiAgICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQoZW1wdHlTdGF0ZSh7CiAgICAgICAgaWNvbjogJ2RhdGFiYXNlJywKICAgICAgICB0aXRsZTogYE5vICR7cGxhdGZvcm1MYWJlbChyZXN1bHQucGxhdGZvcm0pfSByZWNvcmRzIG1hdGNoIHRoZXNlIGZpbHRlcnMgeWV0YCwKICAgICAgICBtZXNzYWdlOiAnVHJ5IGEgZGlmZmVyZW50IHBsYXRmb3JtLCBvciB3aWRlbiB0aGUgZGF0ZSByYW5nZS4nLAogICAgICB9KSk7CiAgICAgIHJldHVybjsKICAgIH0KICAgIGNvbnN0IHRhYmxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGFibGUnKTsKICAgIHRhYmxlLmNsYXNzTmFtZSA9ICdkYXRhLXRhYmxlJzsKICAgIGNvbnN0IHRoZWFkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGhlYWQnKTsKICAgIGNvbnN0IGhlYWRUciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICBoZWFkVHIuYXBwZW5kKHRleHRFbCgndGgnLCAnRGF0ZScpLCB0ZXh0RWwoJ3RoJywgJ0NhcHRpb24nKSwgdGV4dEVsKCd0aCcsICdDYW1wYWlnbicpLCB0ZXh0RWwoJ3RoJywgJ0NvbnRlbnQgVHlwZScpKTsKICAgIHJlc3VsdC5jb2x1bW5zLmZvckVhY2goKGMpID0+IGhlYWRUci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3RoJywgYy5sYWJlbCwgYy5rZXkgPT09ICdwb3N0aW5nX2xpbmsnID8gJycgOiAnbnVtJykpKTsKICAgIGhlYWRUci5hcHBlbmQodGV4dEVsKCd0aCcsICdTdGF0dXMnKSwgdGV4dEVsKCd0aCcsICdBY3Rpb25zJykpOwogICAgdGhlYWQuYXBwZW5kQ2hpbGQoaGVhZFRyKTsKICAgIGNvbnN0IHRib2R5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGJvZHknKTsKICAgIHJlc3VsdC5yb3dzLmZvckVhY2goKHIpID0+IHsKICAgICAgY29uc3QgdHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgICB0ci5hcHBlbmQodGV4dEVsKCd0ZCcsIEZvcm1hdC5kYXRlKHIucHVibGlzaERhdGUpKSwgY2FwdGlvbkNlbGwoci5jYXB0aW9uKSwgdGV4dEVsKCd0ZCcsIHIuY2FtcGFpZ25UeXBlIHx8ICfigJQnKSwgdGV4dEVsKCd0ZCcsIHIuY29udGVudFR5cGUgfHwgJ+KAlCcpKTsKICAgICAgcmVzdWx0LmNvbHVtbnMuZm9yRWFjaCgoYykgPT4gdHIuYXBwZW5kQ2hpbGQobWV0cmljQ2VsbChjLmtleSwgci5tZXRyaWNzW2Mua2V5XSkpKTsKICAgICAgY29uc3Qgc3RhdHVzVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBzdGF0dXNUZC5hcHBlbmRDaGlsZChzdGF0dXNQaWxsKHIuc3RhdHVzKSk7CiAgICAgIHRyLmFwcGVuZENoaWxkKHN0YXR1c1RkKTsKICAgICAgY29uc3QgYWN0aW9uc1RkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgYWN0aW9uc1RkLmFwcGVuZENoaWxkKGFjdGlvbkJ1dHRvbnMociwgcmVzdWx0LnBsYXRmb3JtKSk7CiAgICAgIHRyLmFwcGVuZENoaWxkKGFjdGlvbnNUZCk7CiAgICAgIHRib2R5LmFwcGVuZENoaWxkKHRyKTsKICAgIH0pOwogICAgdGFibGUuYXBwZW5kKHRoZWFkLCB0Ym9keSk7CiAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgd3JhcC5hcHBlbmRDaGlsZCh0YWJsZSk7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJUYWJsZShyZXN1bHQpIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncmVjb3Jkc1RhYmxlV3JhcCcpOwogICAgaWYgKCF3cmFwKSByZXR1cm47CiAgICBpZiAocmVzdWx0LnBsYXRmb3JtID09PSAnYWxsJykgcmVuZGVyU3VtbWFyeVRhYmxlKHJlc3VsdCk7CiAgICBlbHNlIHJlbmRlclBsYXRmb3JtVGFibGUocmVzdWx0KTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlclBhZ2VyKHJlc3VsdCkgewogICAgY29uc3QgcGFnZXIgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncmVjb3Jkc1BhZ2VyJyk7CiAgICBpZiAoIXBhZ2VyKSByZXR1cm47CiAgICBwYWdlci5pbm5lckhUTUwgPSAnJzsKICAgIGNvbnN0IHRvdGFsUGFnZXMgPSBNYXRoLm1heCgxLCBNYXRoLmNlaWwocmVzdWx0LnRvdGFsIC8gcmVzdWx0LnBhZ2VTaXplKSk7CiAgICBjb25zdCBwcmV2QnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBwcmV2QnRuLmNsYXNzTmFtZSA9ICdidG4nOwogICAgcHJldkJ0bi50ZXh0Q29udGVudCA9ICdQcmV2aW91cyc7CiAgICBwcmV2QnRuLmRpc2FibGVkID0gcmVzdWx0LnBhZ2UgPD0gMTsKICAgIHByZXZCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7IHBhZ2UgLT0gMTsgbG9hZCgpOyB9KTsKICAgIGNvbnN0IG5leHRCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIG5leHRCdG4uY2xhc3NOYW1lID0gJ2J0bic7CiAgICBuZXh0QnRuLnRleHRDb250ZW50ID0gJ05leHQnOwogICAgbmV4dEJ0bi5kaXNhYmxlZCA9IHJlc3VsdC5wYWdlID49IHRvdGFsUGFnZXM7CiAgICBuZXh0QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4geyBwYWdlICs9IDE7IGxvYWQoKTsgfSk7CiAgICBwYWdlci5hcHBlbmQocHJldkJ0biwgdGV4dEVsKCdzcGFuJywgYFBhZ2UgJHtyZXN1bHQucGFnZX0gb2YgJHt0b3RhbFBhZ2VzfSDigJQgJHtyZXN1bHQudG90YWx9IHJlY29yZChzKWApLCBuZXh0QnRuKTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIGhhbmRsZURlbGV0ZShyb3csIHBsYXRmb3JtKSB7CiAgICBjb25zdCBjYXB0aW9uID0gKHJvdy5jYXB0aW9uIHx8ICcobm8gY2FwdGlvbiknKS5zbGljZSgwLCA2MCk7CiAgICBjb25zdCBtZXNzYWdlID0gcGxhdGZvcm0gPT09ICdhbGwnCiAgICAgID8gYERlbGV0ZSB0aGlzIGVudGlyZSByZWNvcmQg4oCUICIke2NhcHRpb259IiDigJQgYWNyb3NzIGV2ZXJ5IHBsYXRmb3JtPyBJdHMgb3JpZ2luYWwgaW1wb3J0IHN0YXlzIGluIFVwbG9hZCBIaXN0b3J5LCBidXQgaXQgd2lsbCBkaXNhcHBlYXIgZnJvbSB0aGUgZGFzaGJvYXJkLCBjb21wYXJpc29ucywgYW5kIHJlcG9ydHMuYAogICAgICA6IGBSZW1vdmUgdGhpcyByZWNvcmQncyAke3BsYXRmb3JtTGFiZWwocGxhdGZvcm0pfSBkYXRhIOKAlCAiJHtjYXB0aW9ufSI/IElmIHRoaXMgaXMgaXRzIG9ubHkgcGxhdGZvcm0sIHRoZSB3aG9sZSByZWNvcmQgd2lsbCBiZSByZW1vdmVkIGZyb20gdGhlIGRhc2hib2FyZC5gOwogICAgaWYgKCF3aW5kb3cuY29uZmlybShtZXNzYWdlKSkgcmV0dXJuOwogICAgdHJ5IHsKICAgICAgaWYgKHBsYXRmb3JtID09PSAnYWxsJykgYXdhaXQgQXBpLmRlbGV0ZVJlY29yZFBvc3Qocm93LnBvc3RJZCk7CiAgICAgIGVsc2UgYXdhaXQgQXBpLmRlbGV0ZVJlY29yZFBsYXRmb3JtKHJvdy5wb3N0SWQsIHBsYXRmb3JtKTsKICAgICAgVG9hc3Quc2hvdygnUmVjb3JkIGRlbGV0ZWQuJywgJ3N1Y2Nlc3MnKTsKICAgICAgYXdhaXQgbG9hZCgpOwogICAgICB3aW5kb3cuZGlzcGF0Y2hFdmVudChuZXcgQ3VzdG9tRXZlbnQoJ2xyczpkYXRhLXVwZGF0ZWQnKSk7CiAgICB9IGNhdGNoIChlcnIpIHsKICAgICAgVG9hc3Quc2hvdyhlcnIubWVzc2FnZSwgJ2Vycm9yJyk7CiAgICB9CiAgfQoKICBmdW5jdGlvbiBjbG9zZU1vZGFsKCkgewogICAgY29uc3Qgb3ZlcmxheSA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdyZWNvcmRNb2RhbE92ZXJsYXknKTsKICAgIGlmIChvdmVybGF5KSBvdmVybGF5LnJlbW92ZSgpOwogICAgbW9kYWxTdGF0ZSA9IG51bGw7CiAgfQoKICBmdW5jdGlvbiBtb2RhbFNoZWxsKHRpdGxlVGV4dCkgewogICAgY2xvc2VNb2RhbCgpOwogICAgY29uc3Qgb3ZlcmxheSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgb3ZlcmxheS5jbGFzc05hbWUgPSAnbW9kYWwtb3ZlcmxheSc7CiAgICBvdmVybGF5LmlkID0gJ3JlY29yZE1vZGFsT3ZlcmxheSc7CiAgICBvdmVybGF5LmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKGUpID0+IHsgaWYgKGUudGFyZ2V0ID09PSBvdmVybGF5KSBjbG9zZU1vZGFsKCk7IH0pOwogICAgY29uc3QgcGFuZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHBhbmVsLmNsYXNzTmFtZSA9ICdtb2RhbC1wYW5lbCB3aWRlJzsKICAgIHBhbmVsLmFwcGVuZENoaWxkKHRleHRFbCgnaDInLCB0aXRsZVRleHQpKTsKICAgIG92ZXJsYXkuYXBwZW5kQ2hpbGQocGFuZWwpOwogICAgcmV0dXJuIHsgb3ZlcmxheSwgcGFuZWwgfTsKICB9CgogIGZ1bmN0aW9uIHJlY29yZFN1YnRpdGxlKHIpIHsKICAgIHJldHVybiBgU2hlZXQgIiR7ci5zaGVldE5hbWV9Iiwgcm93ICR7ci5yb3dOdW1iZXJ9JHtyLnBvc3RJZCA/IGAg4oCUIGxpbmtlZCB0byBkYXNoYm9hcmQgcG9zdCAjJHtyLnBvc3RJZH1gIDogJyDigJQgbm90IHBhcnQgb2YgdGhlIGRhc2hib2FyZCAoZS5nLiBuZWVkcyBhIHZhbGlkIGRhdGUpJ31gOwogIH0KCiAgLy8gLS0tLS0tLS0tLSBWaWV3IHBvcHVwOiByZWFkLW9ubHksIGV2ZXJ5IGZpZWxkLCBncm91cGVkIGludG8gc2VjdGlvbnMgLS0tLS0tLS0tLQogIGFzeW5jIGZ1bmN0aW9uIG9wZW5WaWV3KGlkKSB7CiAgICBjb25zdCByZWNvcmQgPSBhd2FpdCBBcGkuZ2V0UmVjb3JkKGlkKTsKICAgIGNvbnN0IHsgb3ZlcmxheSwgcGFuZWwgfSA9IG1vZGFsU2hlbGwoJ1JlY29yZCBkZXRhaWxzJyk7CiAgICBwYW5lbC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIHJlY29yZFN1YnRpdGxlKHJlY29yZCksICdtb2RhbC1zdWInKSk7CgogICAgY29uc3QgZ3JvdXBzID0gZ3JvdXBGaWVsZFJvd3MoY29sdW1uTGFiZWxzRm9yKHJlY29yZCksIHJlY29yZC52YWx1ZXMpOwogICAgZ3JvdXBzLmZvckVhY2goKGcpID0+IHsKICAgICAgY29uc3Qgc2VjdGlvbiA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBzZWN0aW9uLmNsYXNzTmFtZSA9ICdyZWNvcmQtc2VjdGlvbic7CiAgICAgIHNlY3Rpb24uYXBwZW5kQ2hpbGQodGV4dEVsKCdoNCcsIGcuZ3JvdXApKTsKICAgICAgY29uc3QgZ3JpZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBncmlkLmNsYXNzTmFtZSA9ICdmb3JtLWdyaWQnOwogICAgICBnLmZpZWxkcy5mb3JFYWNoKChmKSA9PiB7CiAgICAgICAgY29uc3QgZmllbGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgICBmaWVsZC5jbGFzc05hbWUgPSAndmlldy1maWVsZCc7CiAgICAgICAgZmllbGQuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBmLmxhYmVsLCAndmlldy1sYWJlbCcpKTsKICAgICAgICBjb25zdCB2YWwgPSBmLnZhbHVlID09PSB1bmRlZmluZWQgfHwgZi52YWx1ZSA9PT0gbnVsbCB8fCBmLnZhbHVlID09PSAnJyA/ICfigJQnIDogU3RyaW5nKGYudmFsdWUpOwogICAgICAgIGZpZWxkLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgdmFsLCAndmlldy12YWx1ZScpKTsKICAgICAgICBncmlkLmFwcGVuZENoaWxkKGZpZWxkKTsKICAgICAgfSk7CiAgICAgIHNlY3Rpb24uYXBwZW5kQ2hpbGQoZ3JpZCk7CiAgICAgIHBhbmVsLmFwcGVuZENoaWxkKHNlY3Rpb24pOwogICAgfSk7CgogICAgY29uc3QgYWN0aW9ucyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYWN0aW9ucy5jbGFzc05hbWUgPSAnbW9kYWwtYWN0aW9ucyc7CiAgICBjb25zdCBidG5Sb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGJ0blJvdy5jbGFzc05hbWUgPSAnYnRuLXJvdyc7CiAgICBjb25zdCBjbG9zZUJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgY2xvc2VCdG4uY2xhc3NOYW1lID0gJ2J0bic7CiAgICBjbG9zZUJ0bi50ZXh0Q29udGVudCA9ICdDbG9zZSc7CiAgICBjbG9zZUJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIGNsb3NlTW9kYWwpOwogICAgY29uc3QgZWRpdEJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgZWRpdEJ0bi5jbGFzc05hbWUgPSAnYnRuIHByaW1hcnknOwogICAgZWRpdEJ0bi50ZXh0Q29udGVudCA9ICdFZGl0IHRoaXMgcmVjb3JkJzsKICAgIGVkaXRCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBvcGVuRWRpdG9yKHJlY29yZC5pZCkpOwogICAgYnRuUm93LmFwcGVuZChjbG9zZUJ0biwgZWRpdEJ0bik7CiAgICBhY3Rpb25zLmFwcGVuZENoaWxkKGJ0blJvdyk7CiAgICBwYW5lbC5hcHBlbmRDaGlsZChhY3Rpb25zKTsKCiAgICBkb2N1bWVudC5ib2R5LmFwcGVuZENoaWxkKG92ZXJsYXkpOwogIH0KCiAgLy8gLS0tLS0tLS0tLSBFZGl0IHBvcHVwOiBldmVyeSBmaWVsZCwgZ3JvdXBlZCBpbnRvIHNlY3Rpb25zLCBhbGwgZWRpdGFibGUgLS0tLS0tLS0tLQogIGFzeW5jIGZ1bmN0aW9uIG9wZW5FZGl0b3IoaWQpIHsKICAgIGNvbnN0IHJlY29yZCA9IGF3YWl0IEFwaS5nZXRSZWNvcmQoaWQpOwogICAgbW9kYWxTdGF0ZSA9IHsgcmVjb3JkLCB2YWx1ZXM6IFsuLi5yZWNvcmQudmFsdWVzXSB9OwogICAgcmVuZGVyRWRpdE1vZGFsKCk7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJFZGl0TW9kYWwoKSB7CiAgICBjb25zdCByID0gbW9kYWxTdGF0ZS5yZWNvcmQ7CiAgICBjb25zdCB7IG92ZXJsYXksIHBhbmVsIH0gPSBtb2RhbFNoZWxsKCdFZGl0IHJlY29yZCcpOwogICAgcGFuZWwuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCByZWNvcmRTdWJ0aXRsZShyKSwgJ21vZGFsLXN1YicpKTsKCiAgICBjb25zdCBncm91cHMgPSBncm91cEZpZWxkUm93cyhjb2x1bW5MYWJlbHNGb3IociksIG1vZGFsU3RhdGUudmFsdWVzKTsKICAgIGdyb3Vwcy5mb3JFYWNoKChnKSA9PiB7CiAgICAgIGNvbnN0IHNlY3Rpb24gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgc2VjdGlvbi5jbGFzc05hbWUgPSAncmVjb3JkLXNlY3Rpb24nOwogICAgICBzZWN0aW9uLmFwcGVuZENoaWxkKHRleHRFbCgnaDQnLCBnLmdyb3VwKSk7CiAgICAgIGNvbnN0IGdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgZ3JpZC5jbGFzc05hbWUgPSAnZm9ybS1ncmlkJzsKICAgICAgZy5maWVsZHMuZm9yRWFjaCgoZikgPT4gewogICAgICAgIGNvbnN0IGZpZWxkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgICAgZmllbGQuY2xhc3NOYW1lID0gJ2Zvcm0tZmllbGQnOwogICAgICAgIGZpZWxkLmFwcGVuZENoaWxkKHRleHRFbCgnbGFiZWwnLCBmLmxhYmVsKSk7CiAgICAgICAgY29uc3Qgc3RyVmFsID0gZi52YWx1ZSA9PT0gdW5kZWZpbmVkIHx8IGYudmFsdWUgPT09IG51bGwgPyAnJyA6IFN0cmluZyhmLnZhbHVlKTsKICAgICAgICBjb25zdCBpc0xvbmcgPSBzdHJWYWwubGVuZ3RoID4gODA7CiAgICAgICAgY29uc3QgaW5wdXQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KGlzTG9uZyA/ICd0ZXh0YXJlYScgOiAnaW5wdXQnKTsKICAgICAgICBpZiAoIWlzTG9uZykgaW5wdXQudHlwZSA9ICd0ZXh0JzsKICAgICAgICBlbHNlIGZpZWxkLnN0eWxlLmdyaWRDb2x1bW4gPSAnMSAvIC0xJzsKICAgICAgICBpbnB1dC52YWx1ZSA9IHN0clZhbDsKICAgICAgICBpbnB1dC5hZGRFdmVudExpc3RlbmVyKCdpbnB1dCcsICgpID0+IHsgbW9kYWxTdGF0ZS52YWx1ZXNbZi5pZHhdID0gaW5wdXQudmFsdWU7IH0pOwogICAgICAgIGZpZWxkLmFwcGVuZENoaWxkKGlucHV0KTsKICAgICAgICBncmlkLmFwcGVuZENoaWxkKGZpZWxkKTsKICAgICAgfSk7CiAgICAgIHNlY3Rpb24uYXBwZW5kQ2hpbGQoZ3JpZCk7CiAgICAgIHBhbmVsLmFwcGVuZENoaWxkKHNlY3Rpb24pOwogICAgfSk7CgogICAgY29uc3QgYWN0aW9ucyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYWN0aW9ucy5jbGFzc05hbWUgPSAnbW9kYWwtYWN0aW9ucyc7CiAgICBjb25zdCBlcnJvck1zZyA9IHRleHRFbCgnc3BhbicsICcnLCAnbXV0ZWQnKTsKICAgIGVycm9yTXNnLmlkID0gJ21vZGFsRXJyb3JNc2cnOwogICAgY29uc3QgYnRuUm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBidG5Sb3cuY2xhc3NOYW1lID0gJ2J0bi1yb3cnOwogICAgY29uc3QgY2FuY2VsQnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBjYW5jZWxCdG4uY2xhc3NOYW1lID0gJ2J0bic7CiAgICBjYW5jZWxCdG4udGV4dENvbnRlbnQgPSAnQ2FuY2VsJzsKICAgIGNhbmNlbEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIGNsb3NlTW9kYWwpOwogICAgY29uc3Qgc2F2ZUJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgc2F2ZUJ0bi5jbGFzc05hbWUgPSAnYnRuIHByaW1hcnknOwogICAgc2F2ZUJ0bi50ZXh0Q29udGVudCA9ICdTYXZlIGNoYW5nZXMnOwogICAgc2F2ZUJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHNhdmVFZGl0KHNhdmVCdG4pKTsKICAgIGJ0blJvdy5hcHBlbmQoY2FuY2VsQnRuLCBzYXZlQnRuKTsKICAgIGFjdGlvbnMuYXBwZW5kKGVycm9yTXNnLCBidG5Sb3cpOwogICAgcGFuZWwuYXBwZW5kQ2hpbGQoYWN0aW9ucyk7CgogICAgZG9jdW1lbnQuYm9keS5hcHBlbmRDaGlsZChvdmVybGF5KTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIHNhdmVFZGl0KGJ0bikgewogICAgY29uc3QgZXJyb3JFbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdtb2RhbEVycm9yTXNnJyk7CiAgICBlcnJvckVsLnRleHRDb250ZW50ID0gJyc7CiAgICBidG4uZGlzYWJsZWQgPSB0cnVlOwogICAgYnRuLnRleHRDb250ZW50ID0gJ1NhdmluZ+KApic7CiAgICB0cnkgewogICAgICBhd2FpdCBBcGkudXBkYXRlUmVjb3JkKG1vZGFsU3RhdGUucmVjb3JkLmlkLCBtb2RhbFN0YXRlLnZhbHVlcyk7CiAgICAgIFRvYXN0LnNob3coJ1JlY29yZCB1cGRhdGVkLicsICdzdWNjZXNzJyk7CiAgICAgIGNsb3NlTW9kYWwoKTsKICAgICAgYXdhaXQgbG9hZCgpOwogICAgICB3aW5kb3cuZGlzcGF0Y2hFdmVudChuZXcgQ3VzdG9tRXZlbnQoJ2xyczpkYXRhLXVwZGF0ZWQnKSk7CiAgICB9IGNhdGNoIChlcnIpIHsKICAgICAgZXJyb3JFbC50ZXh0Q29udGVudCA9IGVyci5tZXNzYWdlOwogICAgICBlcnJvckVsLnN0eWxlLmNvbG9yID0gJ3ZhcigtLXN0YXR1cy1jcml0aWNhbCknOwogICAgICBidG4uZGlzYWJsZWQgPSBmYWxzZTsKICAgICAgYnRuLnRleHRDb250ZW50ID0gJ1NhdmUgY2hhbmdlcyc7CiAgICB9CiAgfQoKICBhc3luYyBmdW5jdGlvbiByZW5kZXIoKSB7CiAgICByb290ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3ZpZXctcmVjb3JkcycpOwogICAgcGFnZSA9IDE7CiAgICBzaGVsbCgpOwogICAgYXdhaXQgbG9hZCgpOwogIH0KCiAgcmV0dXJuIHsgcmVuZGVyLCByZWxvYWQ6IGxvYWQsIG9wZW5WaWV3IH07Cn0pKCk7CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgQ29tcGFyaXNvbnMgdGFiOiB3ZWVrLXZzLXdlZWssIGN1c3RvbSByYW5nZSwgbW9udGhseSwKICAgcXVhcnRlcmx5LCBZVEQuCiAgID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwpjb25zdCBDb21wYXJpc29uID0gKCgpID0+IHsKICBjb25zdCBNT0RFUyA9IFsKICAgIHsga2V5OiAnd2VlaycsIGxhYmVsOiAnV2VlayB2cyBXZWVrJyB9LAogICAgeyBrZXk6ICdjdXN0b20nLCBsYWJlbDogJ0N1c3RvbSBSYW5nZScgfSwKICAgIHsga2V5OiAnbW9udGgnLCBsYWJlbDogJ01vbnRobHknIH0sCiAgICB7IGtleTogJ3F1YXJ0ZXInLCBsYWJlbDogJ1F1YXJ0ZXJseScgfSwKICAgIHsga2V5OiAneXRkJywgbGFiZWw6ICdZZWFyIHRvIERhdGUnIH0sCiAgXTsKICBjb25zdCBNRVRSSUNfUk9XUyA9IFsKICAgIHsga2V5OiAndmlld3MnLCBsYWJlbDogJ1ZpZXdzJyB9LAogICAgeyBrZXk6ICdyZWFjaCcsIGxhYmVsOiAnUmVhY2gnIH0sCiAgICB7IGtleTogJ2ltcHJlc3Npb25zJywgbGFiZWw6ICdJbXByZXNzaW9ucycgfSwKICAgIHsga2V5OiAnZW5nYWdlbWVudCcsIGxhYmVsOiAnRW5nYWdlbWVudCcgfSwKICAgIHsga2V5OiAnY2xpY2tzJywgbGFiZWw6ICdDbGlja3MnIH0sCiAgICB7IGtleTogJ2ZvbGxvd2Vyc19nYWluZWQnLCBsYWJlbDogJ0ZvbGxvd2VycyBHYWluZWQnIH0sCiAgICB7IGtleTogJ3dhdGNoX3RpbWVfc2Vjb25kcycsIGxhYmVsOiAnV2F0Y2ggVGltZScgfSwKICAgIHsga2V5OiAnc2hhcmVzJywgbGFiZWw6ICdTaGFyZXMnIH0sCiAgICB7IGtleTogJ2NvbW1lbnRzJywgbGFiZWw6ICdDb21tZW50cycgfSwKICAgIHsga2V5OiAnc2F2ZXMnLCBsYWJlbDogJ1NhdmVzJyB9LAogIF07CgogIGxldCBtb2RlID0gJ3dlZWsnOwogIGxldCByb290OwogIGxldCBjaGFydE1ldHJpYyA9ICdlbmdhZ2VtZW50JzsKCiAgZnVuY3Rpb24gbW9uZGF5T2YoZGF0ZVN0cikgewogICAgY29uc3QgZCA9IG5ldyBEYXRlKGRhdGVTdHIpOwogICAgY29uc3QgZGF5ID0gZC5nZXREYXkoKTsKICAgIGNvbnN0IGRpZmYgPSBkYXkgPT09IDAgPyA2IDogZGF5IC0gMTsKICAgIGQuc2V0RGF0ZShkLmdldERhdGUoKSAtIGRpZmYpOwogICAgcmV0dXJuIGQudG9JU09TdHJpbmcoKS5zbGljZSgwLCAxMCk7CiAgfQogIGZ1bmN0aW9uIGFkZERheXMoZGF0ZVN0ciwgbikgewogICAgY29uc3QgZCA9IG5ldyBEYXRlKGRhdGVTdHIpOwogICAgZC5zZXREYXRlKGQuZ2V0RGF0ZSgpICsgbik7CiAgICByZXR1cm4gZC50b0lTT1N0cmluZygpLnNsaWNlKDAsIDEwKTsKICB9CgogIGZ1bmN0aW9uIHNoZWxsKCkgewogICAgcm9vdC5pbm5lckhUTUwgPSAnJzsKCiAgICBjb25zdCB0YWJzID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0YWJzLmNsYXNzTmFtZSA9ICdtb2RlLXRhYnMnOwogICAgTU9ERVMuZm9yRWFjaCgobSkgPT4gewogICAgICBjb25zdCBidG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgICAgYnRuLnRleHRDb250ZW50ID0gbS5sYWJlbDsKICAgICAgYnRuLnR5cGUgPSAnYnV0dG9uJzsKICAgICAgaWYgKG0ua2V5ID09PSBtb2RlKSBidG4uY2xhc3NMaXN0LmFkZCgnaXMtYWN0aXZlJyk7CiAgICAgIGJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsKICAgICAgICBtb2RlID0gbS5rZXk7CiAgICAgICAgc2hlbGwoKTsKICAgICAgfSk7CiAgICAgIHRhYnMuYXBwZW5kQ2hpbGQoYnRuKTsKICAgIH0pOwogICAgcm9vdC5hcHBlbmRDaGlsZCh0YWJzKTsKCiAgICBjb25zdCBjb250cm9scyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY29udHJvbHMuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgY29udHJvbHMuaWQgPSAnY29tcGFyaXNvbkNvbnRyb2xzJzsKICAgIHJvb3QuYXBwZW5kQ2hpbGQoY29udHJvbHMpOwoKICAgIGNvbnN0IHJlc3VsdHMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHJlc3VsdHMuaWQgPSAnY29tcGFyaXNvblJlc3VsdHMnOwogICAgcm9vdC5hcHBlbmRDaGlsZChyZXN1bHRzKTsKCiAgICByZW5kZXJDb250cm9scygpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyQ29udHJvbHMoKSB7CiAgICBjb25zdCBjb250cm9scyA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdjb21wYXJpc29uQ29udHJvbHMnKTsKICAgIGNvbnRyb2xzLmlubmVySFRNTCA9ICcnOwogICAgY29uc3Qgcm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICByb3cuY2xhc3NOYW1lID0gJ2J0bi1yb3cnOwogICAgcm93LnN0eWxlLmFsaWduSXRlbXMgPSAnZW5kJzsKCiAgICBjb25zdCB0b2RheSA9IG5ldyBEYXRlKCkudG9JU09TdHJpbmcoKS5zbGljZSgwLCAxMCk7CiAgICBjb25zdCB0aGlzWWVhciA9IG5ldyBEYXRlKCkuZ2V0RnVsbFllYXIoKTsKCiAgICBpZiAobW9kZSA9PT0gJ3dlZWsnKSB7CiAgICAgIGNvbnN0IHdBID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsgd0EudHlwZSA9ICdkYXRlJzsgd0EudmFsdWUgPSBtb25kYXlPZih0b2RheSk7CiAgICAgIGNvbnN0IHdCID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsgd0IudHlwZSA9ICdkYXRlJzsgd0IudmFsdWUgPSBtb25kYXlPZihhZGREYXlzKHRvZGF5LCAtNykpOwogICAgICByb3cuYXBwZW5kKGxhYmVsZWQoJ1dlZWsgQSAoYW55IGRheSBpbiB3ZWVrKScsIHdBKSwgbGFiZWxlZCgnV2VlayBCIChhbnkgZGF5IGluIHdlZWspJywgd0IpLCBydW5CdG4oKCkgPT4gewogICAgICAgIGNvbnN0IHJhbmdlQSA9IHsgZnJvbTogbW9uZGF5T2Yod0EudmFsdWUpLCB0bzogYWRkRGF5cyhtb25kYXlPZih3QS52YWx1ZSksIDYpIH07CiAgICAgICAgY29uc3QgcmFuZ2VCID0geyBmcm9tOiBtb25kYXlPZih3Qi52YWx1ZSksIHRvOiBhZGREYXlzKG1vbmRheU9mKHdCLnZhbHVlKSwgNikgfTsKICAgICAgICBydW5Db21wYXJlKHJhbmdlQSwgcmFuZ2VCLCBgV2VlayBvZiAke0Zvcm1hdC5kYXRlKHJhbmdlQS5mcm9tKX1gLCBgV2VlayBvZiAke0Zvcm1hdC5kYXRlKHJhbmdlQi5mcm9tKX1gKTsKICAgICAgfSkpOwogICAgfSBlbHNlIGlmIChtb2RlID09PSAnY3VzdG9tJykgewogICAgICBjb25zdCBmQSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7IGZBLnR5cGUgPSAnZGF0ZSc7IGZBLnZhbHVlID0gYWRkRGF5cyh0b2RheSwgLTEzKTsKICAgICAgY29uc3QgdEEgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOyB0QS50eXBlID0gJ2RhdGUnOyB0QS52YWx1ZSA9IHRvZGF5OwogICAgICBjb25zdCBmQiA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7IGZCLnR5cGUgPSAnZGF0ZSc7IGZCLnZhbHVlID0gYWRkRGF5cyh0b2RheSwgLTI3KTsKICAgICAgY29uc3QgdEIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOyB0Qi50eXBlID0gJ2RhdGUnOyB0Qi52YWx1ZSA9IGFkZERheXModG9kYXksIC0xNCk7CiAgICAgIHJvdy5hcHBlbmQoCiAgICAgICAgbGFiZWxlZCgnUmFuZ2UgQSBmcm9tJywgZkEpLCBsYWJlbGVkKCd0bycsIHRBKSwKICAgICAgICBsYWJlbGVkKCdSYW5nZSBCIGZyb20nLCBmQiksIGxhYmVsZWQoJ3RvJywgdEIpLAogICAgICAgIHJ1bkJ0bigoKSA9PiBydW5Db21wYXJlKHsgZnJvbTogZkEudmFsdWUsIHRvOiB0QS52YWx1ZSB9LCB7IGZyb206IGZCLnZhbHVlLCB0bzogdEIudmFsdWUgfSwgJ1JhbmdlIEEnLCAnUmFuZ2UgQicpKQogICAgICApOwogICAgfSBlbHNlIGlmIChtb2RlID09PSAnbW9udGgnKSB7CiAgICAgIGNvbnN0IHkgPSB5ZWFyU2VsZWN0KHRoaXNZZWFyKTsgY29uc3QgbSA9IG1vbnRoU2VsZWN0KG5ldyBEYXRlKCkuZ2V0TW9udGgoKSArIDEpOwogICAgICBjb25zdCB0b2dnbGUgPSBwZXJpb2RUb2dnbGUoKTsKICAgICAgcm93LmFwcGVuZChsYWJlbGVkKCdZZWFyJywgeSksIGxhYmVsZWQoJ01vbnRoJywgbSksIHRvZ2dsZS5lbCwgcnVuQnRuKGFzeW5jICgpID0+IHsKICAgICAgICBjb25zdCByZXBvcnQgPSBhd2FpdCBBcGkubW9udGhseSh7IHllYXI6IHkudmFsdWUsIG1vbnRoOiBtLnZhbHVlLCAuLi5TdGF0ZS5nZXRGaWx0ZXJzKCkgfSk7CiAgICAgICAgcmVuZGVyUGVyaW9kUmVwb3J0KHJlcG9ydCwgdG9nZ2xlLmdldCgpKTsKICAgICAgfSkpOwogICAgfSBlbHNlIGlmIChtb2RlID09PSAncXVhcnRlcicpIHsKICAgICAgY29uc3QgeSA9IHllYXJTZWxlY3QodGhpc1llYXIpOyBjb25zdCBxID0gcXVhcnRlclNlbGVjdCgpOwogICAgICBjb25zdCB0b2dnbGUgPSBwZXJpb2RUb2dnbGUoKTsKICAgICAgcm93LmFwcGVuZChsYWJlbGVkKCdZZWFyJywgeSksIGxhYmVsZWQoJ1F1YXJ0ZXInLCBxKSwgdG9nZ2xlLmVsLCBydW5CdG4oYXN5bmMgKCkgPT4gewogICAgICAgIGNvbnN0IHJlcG9ydCA9IGF3YWl0IEFwaS5xdWFydGVybHkoeyB5ZWFyOiB5LnZhbHVlLCBxdWFydGVyOiBxLnZhbHVlLCAuLi5TdGF0ZS5nZXRGaWx0ZXJzKCkgfSk7CiAgICAgICAgcmVuZGVyUGVyaW9kUmVwb3J0KHJlcG9ydCwgdG9nZ2xlLmdldCgpKTsKICAgICAgfSkpOwogICAgfSBlbHNlIGlmIChtb2RlID09PSAneXRkJykgewogICAgICBjb25zdCB5ID0geWVhclNlbGVjdCh0aGlzWWVhcik7CiAgICAgIHJvdy5hcHBlbmQobGFiZWxlZCgnWWVhcicsIHkpLCBydW5CdG4oYXN5bmMgKCkgPT4gewogICAgICAgIGNvbnN0IHJlcG9ydCA9IGF3YWl0IEFwaS55dGQoeyB5ZWFyOiB5LnZhbHVlLCAuLi5TdGF0ZS5nZXRGaWx0ZXJzKCkgfSk7CiAgICAgICAgcmVuZGVyUGVyaW9kUmVwb3J0KHJlcG9ydCwgJ3ZzTGFzdFllYXInKTsKICAgICAgfSkpOwogICAgfQoKICAgIGNvbnRyb2xzLmFwcGVuZENoaWxkKHJvdyk7CiAgfQoKICBmdW5jdGlvbiBsYWJlbGVkKGxhYmVsLCBlbCkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgd3JhcC5jbGFzc05hbWUgPSAnZmllbGQtaW5saW5lJzsKICAgIHdyYXAuYXBwZW5kKHRleHRFbCgnbGFiZWwnLCBsYWJlbCksIGVsKTsKICAgIHJldHVybiB3cmFwOwogIH0KICBmdW5jdGlvbiBydW5CdG4ob25DbGljaykgewogICAgY29uc3QgYnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBidG4uY2xhc3NOYW1lID0gJ2J0biBwcmltYXJ5JzsKICAgIGJ0bi50ZXh0Q29udGVudCA9ICdDb21wYXJlJzsKICAgIGJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IG9uQ2xpY2soKSk7CiAgICByZXR1cm4gYnRuOwogIH0KICBmdW5jdGlvbiB5ZWFyU2VsZWN0KGRlZmF1bHRZZWFyKSB7CiAgICBjb25zdCBzZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzZWxlY3QnKTsKICAgIGZvciAobGV0IHkgPSBkZWZhdWx0WWVhciAtIDM7IHkgPD0gZGVmYXVsdFllYXIgKyAxOyB5ICs9IDEpIHsKICAgICAgY29uc3Qgb3B0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7IG9wdC52YWx1ZSA9IHk7IG9wdC50ZXh0Q29udGVudCA9IHk7CiAgICAgIGlmICh5ID09PSBkZWZhdWx0WWVhcikgb3B0LnNlbGVjdGVkID0gdHJ1ZTsKICAgICAgc2VsLmFwcGVuZENoaWxkKG9wdCk7CiAgICB9CiAgICByZXR1cm4gc2VsOwogIH0KICBmdW5jdGlvbiBtb250aFNlbGVjdChkZWZhdWx0TW9udGgpIHsKICAgIGNvbnN0IG5hbWVzID0gWydKYW51YXJ5JywnRmVicnVhcnknLCdNYXJjaCcsJ0FwcmlsJywnTWF5JywnSnVuZScsJ0p1bHknLCdBdWd1c3QnLCdTZXB0ZW1iZXInLCdPY3RvYmVyJywnTm92ZW1iZXInLCdEZWNlbWJlciddOwogICAgY29uc3Qgc2VsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc2VsZWN0Jyk7CiAgICBuYW1lcy5mb3JFYWNoKChuLCBpKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOyBvcHQudmFsdWUgPSBpICsgMTsgb3B0LnRleHRDb250ZW50ID0gbjsKICAgICAgaWYgKGkgKyAxID09PSBkZWZhdWx0TW9udGgpIG9wdC5zZWxlY3RlZCA9IHRydWU7CiAgICAgIHNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgICByZXR1cm4gc2VsOwogIH0KICBmdW5jdGlvbiBxdWFydGVyU2VsZWN0KCkgewogICAgY29uc3QgY3VycmVudFEgPSBNYXRoLmZsb29yKG5ldyBEYXRlKCkuZ2V0TW9udGgoKSAvIDMpICsgMTsKICAgIGNvbnN0IHNlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NlbGVjdCcpOwogICAgWzEsIDIsIDMsIDRdLmZvckVhY2goKHEpID0+IHsKICAgICAgY29uc3Qgb3B0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7IG9wdC52YWx1ZSA9IHE7IG9wdC50ZXh0Q29udGVudCA9IGBRJHtxfWA7CiAgICAgIGlmIChxID09PSBjdXJyZW50USkgb3B0LnNlbGVjdGVkID0gdHJ1ZTsKICAgICAgc2VsLmFwcGVuZENoaWxkKG9wdCk7CiAgICB9KTsKICAgIHJldHVybiBzZWw7CiAgfQogIGZ1bmN0aW9uIHBlcmlvZFRvZ2dsZSgpIHsKICAgIGNvbnN0IHNlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NlbGVjdCcpOwogICAgW1sndnNQcmV2aW91c1BlcmlvZCcsICd2cyBQcmV2aW91cyBQZXJpb2QnXSwgWyd2c0xhc3RZZWFyJywgJ3ZzIFNhbWUgUGVyaW9kIExhc3QgWWVhciddXS5mb3JFYWNoKChbdiwgbF0pID0+IHsKICAgICAgY29uc3Qgb3B0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7IG9wdC52YWx1ZSA9IHY7IG9wdC50ZXh0Q29udGVudCA9IGw7CiAgICAgIHNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgICByZXR1cm4geyBlbDogbGFiZWxlZCgnQ29tcGFyZScsIHNlbCksIGdldDogKCkgPT4gc2VsLnZhbHVlIH07CiAgfQoKICBhc3luYyBmdW5jdGlvbiBydW5Db21wYXJlKHJhbmdlQSwgcmFuZ2VCLCBsYWJlbEEsIGxhYmVsQikgewogICAgY29uc3QgZmlsdGVycyA9IFN0YXRlLmdldEZpbHRlcnMoKTsKICAgIGNvbnN0IHJlc3VsdCA9IGF3YWl0IEFwaS5jb21wYXJlKHsKICAgICAgZnJvbUE6IHJhbmdlQS5mcm9tLCB0b0E6IHJhbmdlQS50bywgZnJvbUI6IHJhbmdlQi5mcm9tLCB0b0I6IHJhbmdlQi50bywKICAgICAgcGxhdGZvcm06IGZpbHRlcnMucGxhdGZvcm0sIGNhbXBhaWduVHlwZTogZmlsdGVycy5jYW1wYWlnblR5cGUsIGNvbnRlbnRUeXBlOiBmaWx0ZXJzLmNvbnRlbnRUeXBlLAogICAgfSk7CiAgICByZW5kZXJDb21wYXJlUmVzdWx0KHJlc3VsdCwgbGFiZWxBLCBsYWJlbEIpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyUGVyaW9kUmVwb3J0KHJlcG9ydCwgd2hpY2gpIHsKICAgIGNvbnN0IGNtcCA9IHJlcG9ydFt3aGljaF07CiAgICBjb25zdCBsYWJlbEEgPSAnQ3VycmVudCBwZXJpb2QnOwogICAgY29uc3QgbGFiZWxCID0gd2hpY2ggPT09ICd2c0xhc3RZZWFyJyA/ICdTYW1lIHBlcmlvZCBsYXN0IHllYXInIDogJ1ByZXZpb3VzIHBlcmlvZCc7CiAgICByZW5kZXJDb21wYXJlUmVzdWx0KGNtcCwgbGFiZWxBLCBsYWJlbEIsIHJlcG9ydC5yYW5nZSk7CiAgfQoKICBmdW5jdGlvbiBzdGF0VGlsZShsYWJlbCwgY3VycmVudCwgcHJldmlvdXMsIGdyb3d0aCwgaXNEdXJhdGlvbikgewogICAgY29uc3QgdGlsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGlsZS5jbGFzc05hbWUgPSAnc3RhdC10aWxlJzsKICAgIGNvbnN0IGN1ckRpc3BsYXkgPSBpc0R1cmF0aW9uID8gRm9ybWF0LmR1cmF0aW9uKGN1cnJlbnQpIDogRm9ybWF0LmNvbXBhY3QoY3VycmVudCk7CiAgICBjb25zdCBwcmV2RGlzcGxheSA9IGlzRHVyYXRpb24gPyBGb3JtYXQuZHVyYXRpb24ocHJldmlvdXMpIDogRm9ybWF0LmNvbXBhY3QocHJldmlvdXMpOwogICAgdGlsZS5hcHBlbmQoCiAgICAgIHRleHRFbCgnZGl2JywgbGFiZWwsICdzdGF0LWxhYmVsJyksCiAgICAgIHRleHRFbCgnZGl2JywgY3VyRGlzcGxheSwgJ3N0YXQtdmFsdWUnKSwKICAgICAgdGV4dEVsKCdkaXYnLCBgJHtGb3JtYXQucGN0KGdyb3d0aCl9IMK3IHdhcyAke3ByZXZEaXNwbGF5fWAsIGBzdGF0LWRlbHRhICR7Rm9ybWF0LmRlbHRhQ2xhc3MoZ3Jvd3RoKX1gKQogICAgKTsKICAgIHJldHVybiB0aWxlOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyQ29tcGFyZVJlc3VsdChyZXN1bHQsIGxhYmVsQSwgbGFiZWxCLCBoZWFkbGluZSkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdjb21wYXJpc29uUmVzdWx0cycpOwogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKCiAgICBjb25zdCB0aXRsZSA9IHRleHRFbCgnZGl2JywgaGVhZGxpbmUKICAgICAgPyBgJHtGb3JtYXQuZGF0ZShyZXN1bHQucmFuZ2VBLmZyb20pfSDigJMgJHtGb3JtYXQuZGF0ZShyZXN1bHQucmFuZ2VBLnRvKX1gCiAgICAgIDogYCR7bGFiZWxBfTogJHtGb3JtYXQuZGF0ZShyZXN1bHQucmFuZ2VBLmZyb20pfSDigJMgJHtGb3JtYXQuZGF0ZShyZXN1bHQucmFuZ2VBLnRvKX0gIHZzICAke2xhYmVsQn06ICR7Rm9ybWF0LmRhdGUocmVzdWx0LnJhbmdlQi5mcm9tKX0g4oCTICR7Rm9ybWF0LmRhdGUocmVzdWx0LnJhbmdlQi50byl9YCwKICAgICAgJ3NlY3Rpb24tdGl0bGUnKTsKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGl0bGUpOwoKICAgIGNvbnN0IGdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGdyaWQuY2xhc3NOYW1lID0gJ3N0YXQtZ3JpZCc7CiAgICBncmlkLmFwcGVuZENoaWxkKHN0YXRUaWxlKCdQb3N0cycsIHJlc3VsdC5yYW5nZUEudG90YWxzLnBvc3RfY291bnQsIHJlc3VsdC5yYW5nZUIudG90YWxzLnBvc3RfY291bnQsIHJlc3VsdC5ncm93dGgucG9zdF9jb3VudCwgZmFsc2UpKTsKICAgIE1FVFJJQ19ST1dTLmZvckVhY2goKG0pID0+IHsKICAgICAgZ3JpZC5hcHBlbmRDaGlsZChzdGF0VGlsZShtLmxhYmVsLCByZXN1bHQucmFuZ2VBLnRvdGFsc1ttLmtleV0sIHJlc3VsdC5yYW5nZUIudG90YWxzW20ua2V5XSwgcmVzdWx0Lmdyb3d0aFttLmtleV0sIG0ua2V5ID09PSAnd2F0Y2hfdGltZV9zZWNvbmRzJykpOwogICAgfSk7CiAgICB3cmFwLmFwcGVuZENoaWxkKGdyaWQpOwoKICAgIGNvbnN0IGNoYXJ0VGl0bGUgPSB0ZXh0RWwoJ2RpdicsIGAke2xhYmVsQX0gdnMgJHtsYWJlbEJ9IGJ5IHBsYXRmb3JtYCwgJ3NlY3Rpb24tdGl0bGUnKTsKICAgIGNvbnN0IGNhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgY29uc3QgaGVhZGVyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBoZWFkZXIuY2xhc3NOYW1lID0gJ2NhcmQtaGVhZGVyJzsKICAgIGhlYWRlci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2gzJywgJ1BsYXRmb3JtIGNvbXBhcmlzb24nKSk7CiAgICBjb25zdCBtZXRyaWNTZWxlY3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzZWxlY3QnKTsKICAgIE1FVFJJQ19ST1dTLmZvckVhY2goKG0pID0+IHsKICAgICAgY29uc3Qgb3B0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7IG9wdC52YWx1ZSA9IG0ua2V5OyBvcHQudGV4dENvbnRlbnQgPSBtLmxhYmVsOwogICAgICBpZiAobS5rZXkgPT09IGNoYXJ0TWV0cmljKSBvcHQuc2VsZWN0ZWQgPSB0cnVlOwogICAgICBtZXRyaWNTZWxlY3QuYXBwZW5kQ2hpbGQob3B0KTsKICAgIH0pOwogICAgbWV0cmljU2VsZWN0LmFkZEV2ZW50TGlzdGVuZXIoJ2NoYW5nZScsICgpID0+IHsKICAgICAgY2hhcnRNZXRyaWMgPSBtZXRyaWNTZWxlY3QudmFsdWU7CiAgICAgIGRyYXdDb21wYXJpc29uQ2hhcnQocmVzdWx0LCBsYWJlbEEsIGxhYmVsQik7CiAgICB9KTsKICAgIGhlYWRlci5hcHBlbmRDaGlsZChtZXRyaWNTZWxlY3QpOwogICAgY29uc3QgY2hhcnRXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjaGFydFdyYXAuY2xhc3NOYW1lID0gJ2NoYXJ0LXdyYXAgdGFsbCc7CiAgICBjaGFydFdyYXAuaW5uZXJIVE1MID0gJzxjYW52YXMgaWQ9ImNvbXBhcmlzb25DYW52YXMiPjwvY2FudmFzPic7CiAgICBjYXJkLmFwcGVuZChoZWFkZXIsIGNoYXJ0V3JhcCk7CiAgICB3cmFwLmFwcGVuZChjaGFydFRpdGxlLCBjYXJkKTsKCiAgICBkcmF3Q29tcGFyaXNvbkNoYXJ0KHJlc3VsdCwgbGFiZWxBLCBsYWJlbEIpOwogIH0KCiAgZnVuY3Rpb24gZHJhd0NvbXBhcmlzb25DaGFydChyZXN1bHQsIGxhYmVsQSwgbGFiZWxCKSB7CiAgICBjb25zdCBwbGF0Zm9ybUlkcyA9IG5ldyBTZXQoWy4uLnJlc3VsdC5yYW5nZUEucGxhdGZvcm1zLCAuLi5yZXN1bHQucmFuZ2VCLnBsYXRmb3Jtc10ubWFwKChwKSA9PiBwLnBsYXRmb3JtKSk7CiAgICBjb25zdCBwbGF0Zm9ybU9wdGlvbnMgPSAod2luZG93Ll9fZmlsdGVyT3B0aW9uc0NhY2hlIHx8IHsgcGxhdGZvcm1zOiBbXSB9KS5wbGF0Zm9ybXM7CiAgICBjb25zdCBsYWJlbHMgPSBbLi4ucGxhdGZvcm1JZHNdLm1hcCgoaWQpID0+IChwbGF0Zm9ybU9wdGlvbnMuZmluZCgocCkgPT4gcC5pZCA9PT0gaWQpIHx8IHsgbGFiZWw6IGlkIH0pLmxhYmVsKTsKICAgIGNvbnN0IGJ5SWRBID0gT2JqZWN0LmZyb21FbnRyaWVzKHJlc3VsdC5yYW5nZUEucGxhdGZvcm1zLm1hcCgocCkgPT4gW3AucGxhdGZvcm0sIHBdKSk7CiAgICBjb25zdCBieUlkQiA9IE9iamVjdC5mcm9tRW50cmllcyhyZXN1bHQucmFuZ2VCLnBsYXRmb3Jtcy5tYXAoKHApID0+IFtwLnBsYXRmb3JtLCBwXSkpOwogICAgY29uc3QgY3VycmVudERhdGEgPSBbLi4ucGxhdGZvcm1JZHNdLm1hcCgoaWQpID0+IChieUlkQVtpZF0gfHwge30pW2NoYXJ0TWV0cmljXSB8fCAwKTsKICAgIGNvbnN0IHByZXZpb3VzRGF0YSA9IFsuLi5wbGF0Zm9ybUlkc10ubWFwKChpZCkgPT4gKGJ5SWRCW2lkXSB8fCB7fSlbY2hhcnRNZXRyaWNdIHx8IDApOwogICAgQ2hhcnRzLmNvbXBhcmlzb25CYXJDaGFydCgnY29tcGFyaXNvbkNhbnZhcycsIHsgbGFiZWxzLCBjdXJyZW50RGF0YSwgcHJldmlvdXNEYXRhLCBjdXJyZW50TGFiZWw6IGxhYmVsQSwgcHJldmlvdXNMYWJlbDogbGFiZWxCIH0pOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyKCkgewogICAgcm9vdCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd2aWV3LWNvbXBhcmlzb24nKTsKICAgIHNoZWxsKCk7CiAgfQoKICByZXR1cm4geyByZW5kZXIgfTsKfSkoKTsKCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBVcGxvYWQgdGFiOiBkcmFnLWRyb3AsIHZhbGlkYXRpb24gcHJldmlldywgcGVyLXdlZWsgY29uZmxpY3QKICAgcmVzb2x1dGlvbiwgY29tbWl0IOKAlCBwbHVzIHRoZSBVcGxvYWQgSGlzdG9yeSB0YWIuCiAgID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwpjb25zdCBVcGxvYWQgPSAoKCkgPT4gewogIGxldCByb290OwogIGxldCBjdXJyZW50UHJldmlldyA9IG51bGw7IC8vIHsgZmlsZVBhdGgsIG9yaWdpbmFsTmFtZSwgZHVwbGljYXRlcywgaXNzdWVzLCBzYW1wbGUsIC4uLiB9CiAgY29uc3QgZHVwbGljYXRlQWN0aW9uT3ZlcnJpZGVzID0ge307CgogIGZ1bmN0aW9uIHNoZWxsKCkgewogICAgcm9vdC5pbm5lckhUTUwgPSAnJzsKCiAgICBjb25zdCBpbnRybyA9IHRleHRFbCgnZGl2JywgJ1VwbG9hZCBhIHdlZWtseSBleHBvcnQnLCAnc2VjdGlvbi10aXRsZScpOwogICAgcm9vdC5hcHBlbmRDaGlsZChpbnRybyk7CgogICAgY29uc3QgZHJvcHpvbmUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGRyb3B6b25lLmNsYXNzTmFtZSA9ICdkcm9wem9uZSc7CiAgICBkcm9wem9uZS5pZCA9ICdkcm9wem9uZSc7CiAgICBkcm9wem9uZS5pbm5lckhUTUwgPSBgCiAgICAgIDxkaXYgY2xhc3M9ImVtcHR5LWljb24iIHN0eWxlPSJtYXJnaW46IDAgYXV0byAxNHB4OyI+PGkgZGF0YS1sdWNpZGU9InVwbG9hZC1jbG91ZCIgc3R5bGU9IndpZHRoOjIycHg7aGVpZ2h0OjIycHg7Ij48L2k+PC9kaXY+CiAgICAgIDxoMz5EcmFnICZhbXA7IGRyb3AgeW91ciAuY3N2IG9yIC54bHN4IGZpbGUgaGVyZTwvaDM+CiAgICAgIDxwPm9yIGNsaWNrIHRvIGJyb3dzZSDigJQgZmlsZXMgYXJlIHZhbGlkYXRlZCBiZWZvcmUgYW55dGhpbmcgaXMgc2F2ZWQ8L3A+CiAgICAgIDxpbnB1dCB0eXBlPSJmaWxlIiBpZD0iZmlsZUlucHV0IiBhY2NlcHQ9Ii5jc3YsLnhsc3gsLnhscyIgLz4KICAgIGA7CiAgICByb290LmFwcGVuZENoaWxkKGRyb3B6b25lKTsKCiAgICBjb25zdCBwcmV2aWV3QXJlYSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcHJldmlld0FyZWEuaWQgPSAncHJldmlld0FyZWEnOwogICAgcm9vdC5hcHBlbmRDaGlsZChwcmV2aWV3QXJlYSk7CgogICAgd2lyZURyb3B6b25lKGRyb3B6b25lKTsKICB9CgogIGZ1bmN0aW9uIHdpcmVEcm9wem9uZShkcm9wem9uZSkgewogICAgY29uc3QgaW5wdXQgPSBkcm9wem9uZS5xdWVyeVNlbGVjdG9yKCcjZmlsZUlucHV0Jyk7CiAgICBkcm9wem9uZS5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IGlucHV0LmNsaWNrKCkpOwogICAgaW5wdXQuYWRkRXZlbnRMaXN0ZW5lcignY2hhbmdlJywgKCkgPT4gewogICAgICBpZiAoaW5wdXQuZmlsZXNbMF0pIGhhbmRsZUZpbGUoaW5wdXQuZmlsZXNbMF0pOwogICAgfSk7CiAgICBbJ2RyYWdlbnRlcicsICdkcmFnb3ZlciddLmZvckVhY2goKGV2dCkgPT4KICAgICAgZHJvcHpvbmUuYWRkRXZlbnRMaXN0ZW5lcihldnQsIChlKSA9PiB7IGUucHJldmVudERlZmF1bHQoKTsgZHJvcHpvbmUuY2xhc3NMaXN0LmFkZCgnaXMtZHJhZycpOyB9KQogICAgKTsKICAgIFsnZHJhZ2xlYXZlJywgJ2Ryb3AnXS5mb3JFYWNoKChldnQpID0+CiAgICAgIGRyb3B6b25lLmFkZEV2ZW50TGlzdGVuZXIoZXZ0LCAoZSkgPT4geyBlLnByZXZlbnREZWZhdWx0KCk7IGRyb3B6b25lLmNsYXNzTGlzdC5yZW1vdmUoJ2lzLWRyYWcnKTsgfSkKICAgICk7CiAgICBkcm9wem9uZS5hZGRFdmVudExpc3RlbmVyKCdkcm9wJywgKGUpID0+IHsKICAgICAgY29uc3QgZmlsZSA9IGUuZGF0YVRyYW5zZmVyLmZpbGVzWzBdOwogICAgICBpZiAoZmlsZSkgaGFuZGxlRmlsZShmaWxlKTsKICAgIH0pOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gaGFuZGxlRmlsZShmaWxlKSB7CiAgICBjb25zdCBhcmVhID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3ByZXZpZXdBcmVhJyk7CiAgICBhcmVhLmlubmVySFRNTCA9ICcnOwogICAgYXJlYS5hcHBlbmRDaGlsZChyb3dXaXRoU3Bpbm5lcignVmFsaWRhdGluZyBmaWxl4oCmJykpOwogICAgT2JqZWN0LmtleXMoZHVwbGljYXRlQWN0aW9uT3ZlcnJpZGVzKS5mb3JFYWNoKChrKSA9PiBkZWxldGUgZHVwbGljYXRlQWN0aW9uT3ZlcnJpZGVzW2tdKTsKICAgIHRyeSB7CiAgICAgIGN1cnJlbnRQcmV2aWV3ID0gYXdhaXQgQXBpLnByZXZpZXdVcGxvYWQoZmlsZSk7CiAgICAgIHJlbmRlclByZXZpZXcoKTsKICAgIH0gY2F0Y2ggKGVycikgewogICAgICBhcmVhLmlubmVySFRNTCA9ICcnOwogICAgICBhcmVhLmFwcGVuZENoaWxkKGVycm9yQmFubmVyKGVyci5tZXNzYWdlKSk7CiAgICB9CiAgfQoKICBmdW5jdGlvbiByb3dXaXRoU3Bpbm5lcih0ZXh0KSB7CiAgICBjb25zdCBlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZWwuY2xhc3NOYW1lID0gJ2xvYWRpbmctcm93JzsKICAgIGNvbnN0IHNwaW5uZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICBzcGlubmVyLmNsYXNzTmFtZSA9ICdzcGlubmVyJzsKICAgIGVsLmFwcGVuZChzcGlubmVyLCBkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShgICR7dGV4dH1gKSk7CiAgICByZXR1cm4gZWw7CiAgfQogIGZ1bmN0aW9uIGVycm9yQmFubmVyKG1lc3NhZ2UpIHsKICAgIGNvbnN0IGVsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBlbC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICBlbC5zdHlsZS5ib3JkZXJMZWZ0ID0gJzNweCBzb2xpZCB2YXIoLS1zdGF0dXMtY3JpdGljYWwpJzsKICAgIGVsLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgYENvdWxkIG5vdCByZWFkIHRoaXMgZmlsZTogJHttZXNzYWdlfWAsICdtdXRlZCcpKTsKICAgIHJldHVybiBlbDsKICB9CgogIGZ1bmN0aW9uIHJlbmRlclByZXZpZXcoKSB7CiAgICBjb25zdCBhcmVhID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3ByZXZpZXdBcmVhJyk7CiAgICBhcmVhLmlubmVySFRNTCA9ICcnOwogICAgY29uc3QgcCA9IGN1cnJlbnRQcmV2aWV3OwoKICAgIGNvbnN0IHN1bW1hcnlUaXRsZSA9IHRleHRFbCgnZGl2JywgJ1ZhbGlkYXRpb24gc3VtbWFyeScsICdzZWN0aW9uLXRpdGxlJyk7CiAgICBjb25zdCBzdW1tYXJ5R3JpZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgc3VtbWFyeUdyaWQuY2xhc3NOYW1lID0gJ3N0YXQtZ3JpZCc7CiAgICBzdW1tYXJ5R3JpZC5hcHBlbmQoCiAgICAgIHN0YXRUaWxlKCdGaWxlJywgcC5vcmlnaW5hbE5hbWUpLAogICAgICBzdGF0VGlsZSgnU2hlZXRzIGZvdW5kJywgcC5zaGVldHMubGVuZ3RoKSwKICAgICAgc3RhdFRpbGUoJ1RvdGFsIHJvd3MgKGFsbCBzaGVldHMpJywgcC50b3RhbERhdGFSb3dzKSwKICAgICAgc3RhdFRpbGUoJ05ldyByZWNvcmRzJywgcC5uZXdSZWNvcmRzQ291bnQpLAogICAgICBzdGF0VGlsZSgnRXhhY3QgZHVwbGljYXRlcyBmb3VuZCcsIHAuZHVwbGljYXRlcy5sZW5ndGgpLAogICAgICBzdGF0VGlsZSgnRHVwbGljYXRlIHJvd3MgaW4gZmlsZScsIHAuZHVwbGljYXRlUm93c0luRmlsZSksCiAgICAgIHN0YXRUaWxlKCdSb3dzIHdpdGggZXJyb3JzJywgcC5lcnJvclJvd3MpCiAgICApOwogICAgYXJlYS5hcHBlbmQoc3VtbWFyeVRpdGxlLCBzdW1tYXJ5R3JpZCk7CgogICAgaWYgKHAuc2hlZXRzLmxlbmd0aCkgewogICAgICBjb25zdCBzaGVldHNUaXRsZSA9IHRleHRFbCgnZGl2JywgJ1NoZWV0IGJyZWFrZG93bicsICdzZWN0aW9uLXRpdGxlJyk7CiAgICAgIGNvbnN0IHNoZWV0c1RhYmxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGFibGUnKTsKICAgICAgc2hlZXRzVGFibGUuY2xhc3NOYW1lID0gJ2RhdGEtdGFibGUnOwogICAgICBzaGVldHNUYWJsZS5pbm5lckhUTUwgPSAnPHRoZWFkPjx0cj48dGg+U2hlZXQ8L3RoPjx0aD5MYXlvdXQgZGV0ZWN0ZWQ8L3RoPjx0aCBjbGFzcz0ibnVtIj5Sb3dzPC90aD48dGggY2xhc3M9Im51bSI+VmFsaWQ8L3RoPjx0aCBjbGFzcz0ibnVtIj5FcnJvcnM8L3RoPjwvdHI+PC90aGVhZD4nOwogICAgICBjb25zdCBzaGVldHNCb2R5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGJvZHknKTsKICAgICAgcC5zaGVldHMuZm9yRWFjaCgocykgPT4gewogICAgICAgIGNvbnN0IHRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgICBjb25zdCBsYXlvdXRMYWJlbCA9IHMuZm9ybWF0ID09PSAnYWdlbmRhJyA/ICdMUlMgYWdlbmRhIHRyYWNrZXInIDogcy5mb3JtYXQgPT09ICdzaW1wbGUnID8gJ1NpbXBsZSBwbGF0Zm9ybSB0YWJsZScgOiAnTm90IHJlY29nbml6ZWQg4oCUIHNhdmVkIGFzIHJhdyBkYXRhIG9ubHknOwogICAgICAgIHRyLmFwcGVuZCgKICAgICAgICAgIHRleHRFbCgndGQnLCBzLm5hbWUpLAogICAgICAgICAgdGV4dEVsKCd0ZCcsIGxheW91dExhYmVsKSwKICAgICAgICAgIHRleHRFbCgndGQnLCBTdHJpbmcocy50b3RhbFJvd3MpLCAnbnVtJyksCiAgICAgICAgICB0ZXh0RWwoJ3RkJywgU3RyaW5nKHMudmFsaWRSb3dzKSwgJ251bScpLAogICAgICAgICAgdGV4dEVsKCd0ZCcsIFN0cmluZyhzLmVycm9yUm93cyksICdudW0nKQogICAgICAgICk7CiAgICAgICAgc2hlZXRzQm9keS5hcHBlbmRDaGlsZCh0cik7CiAgICAgIH0pOwogICAgICBzaGVldHNUYWJsZS5hcHBlbmRDaGlsZChzaGVldHNCb2R5KTsKICAgICAgY29uc3Qgc2hlZXRzV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBzaGVldHNXcmFwLmNsYXNzTmFtZSA9ICd0YWJsZS1zY3JvbGwnOwogICAgICBzaGVldHNXcmFwLmFwcGVuZENoaWxkKHNoZWV0c1RhYmxlKTsKICAgICAgYXJlYS5hcHBlbmQoc2hlZXRzVGl0bGUsIHNoZWV0c1dyYXApOwogICAgfQoKICAgIGlmIChwLmR1cGxpY2F0ZXMubGVuZ3RoKSB7CiAgICAgIGNvbnN0IGR1cFRpdGxlID0gdGV4dEVsKCdkaXYnLCBgRXhhY3QgZHVwbGljYXRlcyBmb3VuZCAoJHtwLmR1cGxpY2F0ZXMubGVuZ3RofSlgLCAnc2VjdGlvbi10aXRsZScpOwogICAgICBhcmVhLmFwcGVuZENoaWxkKGR1cFRpdGxlKTsKICAgICAgYXJlYS5hcHBlbmRDaGlsZCh0ZXh0RWwoCiAgICAgICAgJ2RpdicsCiAgICAgICAgJ0VhY2ggb2YgdGhlc2Ugcm93cyBpcyBieXRlLWZvci1ieXRlIGlkZW50aWNhbCB0byBhbiBhbHJlYWR5LXNhdmVkIHJlY29yZCDigJQgZXZlcnkgZmllbGQgbWF0Y2hlcywgaW5jbHVkaW5nIGV2ZXJ5IG1ldHJpYywgbm90IGp1c3QgdGhlIGRhdGUvY2FwdGlvbi9wbGF0Zm9ybS4gQ2hvb3NlIHdoYXQgdG8gZG8gd2l0aCBlYWNoIOKAlCBvciBzZXQgYSBkZWZhdWx0IGZvciBhbGwgb2YgdGhlbS4gKEEgcm93IHRoYXQgc2hhcmVzIHRoZSBzYW1lIGRhdGUvY2FwdGlvbi9wbGF0Zm9ybSBidXQgaGFzIGRpZmZlcmVudCBudW1iZXJzIGlzIG5vdCBzaG93biBoZXJlIOKAlCBpdOKAmXMgaW1wb3J0ZWQgYXV0b21hdGljYWxseSBhcyBpdHMgb3duIG5ldyByZWNvcmQsIHNpbmNlIGl0cyBhbmFseXRpY3MgY2hhbmdlZC4pJywKICAgICAgICAnbXV0ZWQnCiAgICAgICkpOwogICAgICBjb25zdCBkZWZhdWx0Um93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIGRlZmF1bHRSb3cuY2xhc3NOYW1lID0gJ2ZpZWxkLWlubGluZSc7CiAgICAgIGRlZmF1bHRSb3cuc3R5bGUubWFyZ2luID0gJzEwcHggMCc7CiAgICAgIGNvbnN0IGRlZmF1bHRTZWxlY3QgPSBhY3Rpb25TZWxlY3QoJ3NraXAnKTsKICAgICAgZGVmYXVsdFNlbGVjdC5pZCA9ICdkZWZhdWx0RHVwbGljYXRlQWN0aW9uU2VsZWN0JzsKICAgICAgZGVmYXVsdFNlbGVjdC5hZGRFdmVudExpc3RlbmVyKCdjaGFuZ2UnLCAoKSA9PiB7CiAgICAgICAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnLmNvbmZsaWN0LXJvdyBzZWxlY3RbZGF0YS1oYXNoXScpLmZvckVhY2goKHNlbCkgPT4gewogICAgICAgICAgaWYgKCFkdXBsaWNhdGVBY3Rpb25PdmVycmlkZXNbc2VsLmRhdGFzZXQuaGFzaF0pIHNlbC52YWx1ZSA9IGRlZmF1bHRTZWxlY3QudmFsdWU7CiAgICAgICAgfSk7CiAgICAgIH0pOwogICAgICBkZWZhdWx0Um93LmFwcGVuZCh0ZXh0RWwoJ2xhYmVsJywgJ0RlZmF1bHQgYWN0aW9uIGZvciBhbGwgbWF0Y2hlcycpLCBkZWZhdWx0U2VsZWN0KTsKICAgICAgYXJlYS5hcHBlbmRDaGlsZChkZWZhdWx0Um93KTsKCiAgICAgIGNvbnN0IGxpc3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgbGlzdC5jbGFzc05hbWUgPSAnY29uZmxpY3QtbGlzdCc7CiAgICAgIHAuZHVwbGljYXRlcy5mb3JFYWNoKChkKSA9PiB7CiAgICAgICAgY29uc3Qgcm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgICAgcm93LmNsYXNzTmFtZSA9ICdjb25mbGljdC1yb3cnOwogICAgICAgIGNvbnN0IGxlZnQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgICBsZWZ0LmFwcGVuZCgKICAgICAgICAgIHRleHRFbCgnZGl2JywgYCR7Rm9ybWF0LmRhdGUoZC5wdWJsaXNoRGF0ZSl9IOKAlCAkeyhkLmNhcHRpb24gfHwgJyhubyBjYXB0aW9uKScpLnNsaWNlKDAsIDcwKX1gLCAnd2Vlay1sYWJlbCcpLAogICAgICAgICAgdGV4dEVsKCdkaXYnLCBgRXhhY3QgbWF0Y2ggb2YgZXhpc3RpbmcgcmVjb3JkICMke2QuZXhpc3RpbmcucG9zdElkfSAobGFzdCB1cGRhdGVkICR7ZC5leGlzdGluZy51cGRhdGVkQXR9KWAsICd3ZWVrLW1ldGEnKQogICAgICAgICk7CiAgICAgICAgcm93LmFwcGVuZENoaWxkKGxlZnQpOwogICAgICAgIGNvbnN0IHNlbCA9IGFjdGlvblNlbGVjdCgnc2tpcCcpOwogICAgICAgIHNlbC5kYXRhc2V0Lmhhc2ggPSBkLmhhc2g7CiAgICAgICAgc2VsLmFkZEV2ZW50TGlzdGVuZXIoJ2NoYW5nZScsICgpID0+IHsgZHVwbGljYXRlQWN0aW9uT3ZlcnJpZGVzW2QuaGFzaF0gPSBzZWwudmFsdWU7IH0pOwogICAgICAgIHJvdy5hcHBlbmRDaGlsZChzZWwpOwogICAgICAgIGxpc3QuYXBwZW5kQ2hpbGQocm93KTsKICAgICAgfSk7CiAgICAgIGFyZWEuYXBwZW5kQ2hpbGQobGlzdCk7CiAgICB9CgogICAgY29uc3Qgbm90ZXNGaWVsZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgbm90ZXNGaWVsZC5jbGFzc05hbWUgPSAnZm9ybS1maWVsZCc7CiAgICBub3Rlc0ZpZWxkLnN0eWxlLm1hcmdpbiA9ICcxMnB4IDAnOwogICAgbm90ZXNGaWVsZC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2xhYmVsJywgJ1VwbG9hZCBub3RlcyAob3B0aW9uYWwpJykpOwogICAgY29uc3Qgbm90ZXNJbnB1dCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7CiAgICBub3Rlc0lucHV0LnR5cGUgPSAndGV4dCc7CiAgICBub3Rlc0lucHV0LmlkID0gJ3VwbG9hZE5vdGVzSW5wdXQnOwogICAgbm90ZXNJbnB1dC5wbGFjZWhvbGRlciA9ICdlLmcuICJXZWVrIDMgZXhwb3J0LCBpbmNsdWRlcyBjb3JyZWN0ZWQgVGlrVG9rIG51bWJlcnMiJzsKICAgIG5vdGVzRmllbGQuYXBwZW5kQ2hpbGQobm90ZXNJbnB1dCk7CiAgICBhcmVhLmFwcGVuZENoaWxkKG5vdGVzRmllbGQpOwoKICAgIGlmIChwLmlzc3Vlcy5sZW5ndGgpIHsKICAgICAgY29uc3QgaXNzdWVzVGl0bGUgPSB0ZXh0RWwoJ2RpdicsIGBSb3dzIHNraXBwZWQgb3IgZmxhZ2dlZCAoJHtwLmlzc3Vlcy5sZW5ndGh9KWAsICdzZWN0aW9uLXRpdGxlJyk7CiAgICAgIGNvbnN0IGlzc3Vlc0NhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgaXNzdWVzQ2FyZC5jbGFzc05hbWUgPSAnaXNzdWVzLWxpc3QnOwogICAgICBwLmlzc3Vlcy5mb3JFYWNoKChpc3N1ZSkgPT4gewogICAgICAgIGNvbnN0IHJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICAgIHJvdy5jbGFzc05hbWUgPSAnaXNzdWUtcm93JzsKICAgICAgICBpZiAoaXNzdWUucm93TnVtYmVyKSByb3cuYXBwZW5kQ2hpbGQodGV4dEVsKCdzcGFuJywgYFJvdyAke2lzc3VlLnJvd051bWJlcn1gLCAncm93LW5vJykpOwogICAgICAgIHJvdy5hcHBlbmRDaGlsZChkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShpc3N1ZS5tZXNzYWdlKSk7CiAgICAgICAgaXNzdWVzQ2FyZC5hcHBlbmRDaGlsZChyb3cpOwogICAgICB9KTsKICAgICAgYXJlYS5hcHBlbmQoaXNzdWVzVGl0bGUsIGlzc3Vlc0NhcmQpOwogICAgfQoKICAgIGlmIChwLnNhbXBsZS5sZW5ndGgpIHsKICAgICAgY29uc3Qgc2FtcGxlVGl0bGUgPSB0ZXh0RWwoJ2RpdicsICdTYW1wbGUgb2YgcGFyc2VkIHJvd3MnLCAnc2VjdGlvbi10aXRsZScpOwogICAgICBjb25zdCB0YWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICAgIHRhYmxlLmNsYXNzTmFtZSA9ICdkYXRhLXRhYmxlJzsKICAgICAgdGFibGUuaW5uZXJIVE1MID0gJzx0aGVhZD48dHI+PHRoPkRhdGU8L3RoPjx0aD5DYXB0aW9uPC90aD48dGg+VHlwZTwvdGg+PHRoPkNhbXBhaWduPC90aD48dGg+UGxhdGZvcm1zPC90aD48L3RyPjwvdGhlYWQ+JzsKICAgICAgY29uc3QgdGJvZHkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0Ym9keScpOwogICAgICBwLnNhbXBsZS5mb3JFYWNoKChzKSA9PiB7CiAgICAgICAgY29uc3QgdHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgICAgIHRyLmFwcGVuZCgKICAgICAgICAgIHRleHRFbCgndGQnLCBGb3JtYXQuZGF0ZShzLnB1Ymxpc2hEYXRlKSksCiAgICAgICAgICB0ZXh0RWwoJ3RkJywgcy5jYXB0aW9uIHx8ICfigJQnKSwKICAgICAgICAgIHRleHRFbCgndGQnLCBzLmNvbnRlbnRUeXBlIHx8ICfigJQnKSwKICAgICAgICAgIHRleHRFbCgndGQnLCBzLmNhbXBhaWduVHlwZSB8fCAnVW5zcGVjaWZpZWQnKSwKICAgICAgICAgIHRleHRFbCgndGQnLCBzLnBsYXRmb3Jtcy5qb2luKCcsICcpKQogICAgICAgICk7CiAgICAgICAgdGJvZHkuYXBwZW5kQ2hpbGQodHIpOwogICAgICB9KTsKICAgICAgdGFibGUuYXBwZW5kQ2hpbGQodGJvZHkpOwogICAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIHdyYXAuY2xhc3NOYW1lID0gJ3RhYmxlLXNjcm9sbCc7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQodGFibGUpOwogICAgICBhcmVhLmFwcGVuZChzYW1wbGVUaXRsZSwgd3JhcCk7CiAgICB9CgogICAgY29uc3QgYWN0aW9ucyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYWN0aW9ucy5jbGFzc05hbWUgPSAnYnRuLXJvdyc7CiAgICBhY3Rpb25zLnN0eWxlLm1hcmdpblRvcCA9ICcxNnB4JzsKICAgIGNvbnN0IGNvbW1pdEJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgY29tbWl0QnRuLmNsYXNzTmFtZSA9ICdidG4gcHJpbWFyeSc7CiAgICBjb21taXRCdG4udGV4dENvbnRlbnQgPSBwLnZhbGlkUm93cyA+IDAgPyBgSW1wb3J0ICR7cC52YWxpZFJvd3N9IHJvdyhzKWAgOiAnTm90aGluZyB0byBpbXBvcnQnOwogICAgY29tbWl0QnRuLmRpc2FibGVkID0gcC52YWxpZFJvd3MgPT09IDA7CiAgICBjb21taXRCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBjb21taXQoY29tbWl0QnRuKSk7CiAgICBjb25zdCBjYW5jZWxCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIGNhbmNlbEJ0bi5jbGFzc05hbWUgPSAnYnRuJzsKICAgIGNhbmNlbEJ0bi50ZXh0Q29udGVudCA9ICdDYW5jZWwnOwogICAgY2FuY2VsQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4geyBjdXJyZW50UHJldmlldyA9IG51bGw7IHNoZWxsKCk7IH0pOwogICAgYWN0aW9ucy5hcHBlbmQoY29tbWl0QnRuLCBjYW5jZWxCdG4pOwogICAgYXJlYS5hcHBlbmRDaGlsZChhY3Rpb25zKTsKICB9CgogIGZ1bmN0aW9uIHN0YXRUaWxlKGxhYmVsLCB2YWx1ZSkgewogICAgY29uc3QgdGlsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGlsZS5jbGFzc05hbWUgPSAnc3RhdC10aWxlJzsKICAgIHRpbGUuYXBwZW5kKHRleHRFbCgnZGl2JywgbGFiZWwsICdzdGF0LWxhYmVsJyksIHRleHRFbCgnZGl2JywgU3RyaW5nKHZhbHVlKSwgJ3N0YXQtdmFsdWUnKSk7CiAgICByZXR1cm4gdGlsZTsKICB9CiAgZnVuY3Rpb24gYWN0aW9uU2VsZWN0KGRlZmF1bHRWYWwpIHsKICAgIGNvbnN0IHNlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NlbGVjdCcpOwogICAgW1snc2tpcCcsICdTa2lwIChrZWVwIGV4aXN0aW5nIHJlY29yZCB1bmNoYW5nZWQpJ10sIFsndXBkYXRlJywgJ1VwZGF0ZSBleGlzdGluZyByZWNvcmQnXSwgWydjcmVhdGUnLCAnQ3JlYXRlIGFzIGEgbmV3LCBzZXBhcmF0ZSByZWNvcmQnXV0uZm9yRWFjaCgoW3YsIGxdKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOyBvcHQudmFsdWUgPSB2OyBvcHQudGV4dENvbnRlbnQgPSBsOwogICAgICBpZiAodiA9PT0gZGVmYXVsdFZhbCkgb3B0LnNlbGVjdGVkID0gdHJ1ZTsKICAgICAgc2VsLmFwcGVuZENoaWxkKG9wdCk7CiAgICB9KTsKICAgIHJldHVybiBzZWw7CiAgfQoKICBhc3luYyBmdW5jdGlvbiBjb21taXQoYnRuKSB7CiAgICBidG4uZGlzYWJsZWQgPSB0cnVlOwogICAgYnRuLnRleHRDb250ZW50ID0gJ0ltcG9ydGluZ+KApic7CiAgICBjb25zdCBkZWZhdWx0RHVwbGljYXRlQWN0aW9uID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2RlZmF1bHREdXBsaWNhdGVBY3Rpb25TZWxlY3QnKT8udmFsdWUgfHwgJ3NraXAnOwogICAgY29uc3Qgbm90ZXMgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndXBsb2FkTm90ZXNJbnB1dCcpPy52YWx1ZSB8fCBudWxsOwogICAgdHJ5IHsKICAgICAgY29uc3QgcmVzdWx0ID0gYXdhaXQgQXBpLmNvbW1pdFVwbG9hZCh7CiAgICAgICAgZmlsZVBhdGg6IGN1cnJlbnRQcmV2aWV3LmZpbGVQYXRoLAogICAgICAgIG9yaWdpbmFsTmFtZTogY3VycmVudFByZXZpZXcub3JpZ2luYWxOYW1lLAogICAgICAgIGRlZmF1bHREdXBsaWNhdGVBY3Rpb24sCiAgICAgICAgZHVwbGljYXRlQWN0aW9uczogZHVwbGljYXRlQWN0aW9uT3ZlcnJpZGVzLAogICAgICAgIG5vdGVzLAogICAgICB9KTsKICAgICAgVG9hc3Quc2hvdygKICAgICAgICBgSW1wb3J0ZWQ6ICR7cmVzdWx0LmltcG9ydGVkUm93c30gbmV3LCAke3Jlc3VsdC51cGRhdGVkUm93c30gdXBkYXRlZCwgJHtyZXN1bHQuc2tpcHBlZFJvd3N9IHNraXBwZWQuYCwKICAgICAgICByZXN1bHQuZXJyb3JDb3VudCA+IDAgPyAnZXJyb3InIDogJ3N1Y2Nlc3MnCiAgICAgICk7CiAgICAgIGN1cnJlbnRQcmV2aWV3ID0gbnVsbDsKICAgICAgc2hlbGwoKTsKICAgICAgd2luZG93LmRpc3BhdGNoRXZlbnQobmV3IEN1c3RvbUV2ZW50KCdscnM6ZGF0YS11cGRhdGVkJykpOwogICAgfSBjYXRjaCAoZXJyKSB7CiAgICAgIFRvYXN0LnNob3coZXJyLm1lc3NhZ2UsICdlcnJvcicpOwogICAgICBidG4uZGlzYWJsZWQgPSBmYWxzZTsKICAgICAgYnRuLnRleHRDb250ZW50ID0gJ1JldHJ5IGltcG9ydCc7CiAgICB9CiAgfQoKICBmdW5jdGlvbiByZW5kZXIoKSB7CiAgICByb290ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3ZpZXctdXBsb2FkJyk7CiAgICBzaGVsbCgpOwogIH0KCiAgcmV0dXJuIHsgcmVuZGVyIH07Cn0pKCk7Cgpjb25zdCBIaXN0b3J5ID0gKCgpID0+IHsKICBsZXQgcm9vdDsKCiAgZnVuY3Rpb24gYmFkZ2VDbGFzcyhzdGF0dXMpIHsKICAgIGlmIChzdGF0dXMgPT09ICdzdWNjZXNzJykgcmV0dXJuICdzdWNjZXNzJzsKICAgIGlmIChzdGF0dXMgPT09ICdwYXJ0aWFsJykgcmV0dXJuICdwYXJ0aWFsJzsKICAgIHJldHVybiAnZmFpbGVkJzsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIHJlbmRlcigpIHsKICAgIHJvb3QgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndmlldy1oaXN0b3J5Jyk7CiAgICByb290LmlubmVySFRNTCA9ICcnOwogICAgcm9vdC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsICdVcGxvYWQgaGlzdG9yeScsICdzZWN0aW9uLXRpdGxlJykpOwoKICAgIGNvbnN0IHVwbG9hZHMgPSBhd2FpdCBBcGkudXBsb2FkSGlzdG9yeSgpOwogICAgaWYgKCF1cGxvYWRzLmxlbmd0aCkgewogICAgICByb290LmFwcGVuZENoaWxkKGVtcHR5U3RhdGUoewogICAgICAgIGljb246ICd1cGxvYWQtY2xvdWQnLAogICAgICAgIHRpdGxlOiAnTm8gdXBsb2FkcyB5ZXQnLAogICAgICAgIG1lc3NhZ2U6ICdJbXBvcnQgeW91ciBmaXJzdCB3ZWVrbHkgZXhwb3J0IHRvIHN0YXJ0IHNlZWluZyBkYXRhIGFjcm9zcyB0aGUgYXBwLicsCiAgICAgICAgYWN0aW9uTGFiZWw6ICdVcGxvYWQgZGF0YScsCiAgICAgICAgb25BY3Rpb246ICgpID0+IGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3IoJy50YWItYnRuW2RhdGEtdGFiPSJ1cGxvYWQiXScpPy5jbGljaygpLAogICAgICB9KSk7CiAgICAgIHJldHVybjsKICAgIH0KCiAgICBjb25zdCB0YWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICB0YWJsZS5jbGFzc05hbWUgPSAnZGF0YS10YWJsZSc7CiAgICB0YWJsZS5pbm5lckhUTUwgPSAnPHRoZWFkPjx0cj48dGg+RmlsZTwvdGg+PHRoPlVwbG9hZGVkPC90aD48dGg+U3RhdHVzPC90aD48dGggY2xhc3M9Im51bSI+SW1wb3J0ZWQ8L3RoPjx0aCBjbGFzcz0ibnVtIj5VcGRhdGVkPC90aD48dGggY2xhc3M9Im51bSI+U2tpcHBlZDwvdGg+PHRoIGNsYXNzPSJudW0iPkVycm9yczwvdGg+PHRoPldlZWtzPC90aD48dGg+Tm90ZXM8L3RoPjwvdHI+PC90aGVhZD4nOwogICAgY29uc3QgdGJvZHkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0Ym9keScpOwogICAgdXBsb2Fkcy5mb3JFYWNoKCh1KSA9PiB7CiAgICAgIGNvbnN0IHRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgdHIuc3R5bGUuY3Vyc29yID0gJ3BvaW50ZXInOwogICAgICBjb25zdCBiYWRnZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgICAgYmFkZ2UuY2xhc3NOYW1lID0gYGJhZGdlICR7YmFkZ2VDbGFzcyh1LnN0YXR1cyl9YDsKICAgICAgYmFkZ2UudGV4dENvbnRlbnQgPSB1LnN0YXR1czsKICAgICAgY29uc3Qgc3RhdHVzVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBzdGF0dXNUZC5hcHBlbmRDaGlsZChiYWRnZSk7CiAgICAgIHRyLmFwcGVuZCgKICAgICAgICB0ZXh0RWwoJ3RkJywgdS5maWxlbmFtZSksCiAgICAgICAgdGV4dEVsKCd0ZCcsIHUudXBsb2FkZWRfYXQpLAogICAgICAgIHN0YXR1c1RkLAogICAgICAgIHRleHRFbCgndGQnLCBTdHJpbmcodS5pbXBvcnRlZF9yb3dzKSwgJ251bScpLAogICAgICAgIHRleHRFbCgndGQnLCBTdHJpbmcodS51cGRhdGVkX3Jvd3MpLCAnbnVtJyksCiAgICAgICAgdGV4dEVsKCd0ZCcsIFN0cmluZyh1LnNraXBwZWRfcm93cyksICdudW0nKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgU3RyaW5nKHUuZXJyb3JfY291bnQpLCAnbnVtJyksCiAgICAgICAgdGV4dEVsKCd0ZCcsIHUud2Vla3NfYWZmZWN0ZWQubWFwKCh3KSA9PiBGb3JtYXQuZGF0ZSh3KSkuam9pbignLCAnKSB8fCAn4oCUJyksCiAgICAgICAgdGV4dEVsKCd0ZCcsIHUubm90ZXMgfHwgJ+KAlCcpCiAgICAgICk7CiAgICAgIHRyLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gdG9nZ2xlRXJyb3JzKHUuaWQsIHRyKSk7CiAgICAgIHRib2R5LmFwcGVuZENoaWxkKHRyKTsKICAgIH0pOwogICAgdGFibGUuYXBwZW5kQ2hpbGQodGJvZHkpOwogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgd3JhcC5jbGFzc05hbWUgPSAndGFibGUtc2Nyb2xsJzsKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGFibGUpOwogICAgcm9vdC5hcHBlbmRDaGlsZCh3cmFwKTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIHRvZ2dsZUVycm9ycyh1cGxvYWRJZCwgdHIpIHsKICAgIGNvbnN0IGV4aXN0aW5nID0gdHIubmV4dEVsZW1lbnRTaWJsaW5nOwogICAgaWYgKGV4aXN0aW5nICYmIGV4aXN0aW5nLmNsYXNzTGlzdC5jb250YWlucygnZXJyb3ItbG9nLXJvdycpKSB7CiAgICAgIGV4aXN0aW5nLnJlbW92ZSgpOwogICAgICByZXR1cm47CiAgICB9CiAgICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcuZXJyb3ItbG9nLXJvdycpLmZvckVhY2goKGVsKSA9PiBlbC5yZW1vdmUoKSk7CiAgICBjb25zdCBlcnJvcnMgPSBhd2FpdCBBcGkudXBsb2FkRXJyb3JzKHVwbG9hZElkKTsKICAgIGNvbnN0IGxvZ1JvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICBsb2dSb3cuY2xhc3NOYW1lID0gJ2Vycm9yLWxvZy1yb3cnOwogICAgY29uc3QgdGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgdGQuY29sU3BhbiA9IDk7CiAgICBpZiAoIWVycm9ycy5sZW5ndGgpIHsKICAgICAgdGQuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCAnTm8gaXNzdWVzIGxvZ2dlZCBmb3IgdGhpcyB1cGxvYWQuJywgJ211dGVkJykpOwogICAgfSBlbHNlIHsKICAgICAgY29uc3QgbGlzdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBsaXN0LmNsYXNzTmFtZSA9ICdpc3N1ZXMtbGlzdCc7CiAgICAgIGVycm9ycy5mb3JFYWNoKChlKSA9PiB7CiAgICAgICAgY29uc3Qgcm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgICAgcm93LmNsYXNzTmFtZSA9ICdpc3N1ZS1yb3cnOwogICAgICAgIGNvbnN0IGJhZGdlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICAgIGJhZGdlLmNsYXNzTmFtZSA9IGBiYWRnZSAke2Uuc2V2ZXJpdHl9LXNldmA7CiAgICAgICAgYmFkZ2UudGV4dENvbnRlbnQgPSBlLnNldmVyaXR5OwogICAgICAgIHJvdy5hcHBlbmQoYmFkZ2UsIGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKGAgJHtlLnJvd19udW1iZXIgPyBgUm93ICR7ZS5yb3dfbnVtYmVyfTogYCA6ICcnfSR7ZS5tZXNzYWdlfWApKTsKICAgICAgICBsaXN0LmFwcGVuZENoaWxkKHJvdyk7CiAgICAgIH0pOwogICAgICB0ZC5hcHBlbmRDaGlsZChsaXN0KTsKICAgIH0KCiAgICBjb25zdCByYXdCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIHJhd0J0bi5jbGFzc05hbWUgPSAnYnRuJzsKICAgIHJhd0J0bi5zdHlsZS5tYXJnaW5Ub3AgPSAnMTBweCc7CiAgICByYXdCdG4udGV4dENvbnRlbnQgPSAnVmlldyBldmVyeSByYXcgc291cmNlIHJvdyBmcm9tIHRoaXMgdXBsb2FkJzsKICAgIHJhd0J0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IGxvYWRSYXdSb3dzKHVwbG9hZElkLCByYXdCdG4pKTsKICAgIHRkLmFwcGVuZENoaWxkKHJhd0J0bik7CiAgICBjb25zdCByYXdXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICByYXdXcmFwLmlkID0gYHJhd1dyYXAtJHt1cGxvYWRJZH1gOwogICAgdGQuYXBwZW5kQ2hpbGQocmF3V3JhcCk7CgogICAgbG9nUm93LmFwcGVuZENoaWxkKHRkKTsKICAgIHRyLmFmdGVyKGxvZ1Jvdyk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiBsb2FkUmF3Um93cyh1cGxvYWRJZCwgYnRuKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoYHJhd1dyYXAtJHt1cGxvYWRJZH1gKTsKICAgIGlmICghd3JhcCkgcmV0dXJuOwogICAgaWYgKHdyYXAuZGF0YXNldC5sb2FkZWQpIHsKICAgICAgd3JhcC5zdHlsZS5kaXNwbGF5ID0gd3JhcC5zdHlsZS5kaXNwbGF5ID09PSAnbm9uZScgPyAnYmxvY2snIDogJ25vbmUnOwogICAgICByZXR1cm47CiAgICB9CiAgICBidG4udGV4dENvbnRlbnQgPSAnTG9hZGluZ+KApic7CiAgICBjb25zdCB7IHJvd3MsIHRvdGFsIH0gPSBhd2FpdCBBcGkudXBsb2FkUmF3Um93cyh1cGxvYWRJZCk7CiAgICB3cmFwLmRhdGFzZXQubG9hZGVkID0gJzEnOwogICAgYnRuLnRleHRDb250ZW50ID0gYFNob3dpbmcgJHtyb3dzLmxlbmd0aH0gb2YgJHt0b3RhbH0gcmF3IHJvdyhzKWA7CgogICAgY29uc3QgYnlTaGVldCA9IG5ldyBNYXAoKTsKICAgIHJvd3MuZm9yRWFjaCgocikgPT4gewogICAgICBpZiAoIWJ5U2hlZXQuaGFzKHIuc2hlZXRfbmFtZSkpIGJ5U2hlZXQuc2V0KHIuc2hlZXRfbmFtZSwgW10pOwogICAgICBieVNoZWV0LmdldChyLnNoZWV0X25hbWUpLnB1c2gocik7CiAgICB9KTsKCiAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgd3JhcC5zdHlsZS5tYXJnaW5Ub3AgPSAnMTBweCc7CiAgICBieVNoZWV0LmZvckVhY2goKHNoZWV0Um93cywgc2hlZXROYW1lKSA9PiB7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBgU2hlZXQ6ICR7c2hlZXROYW1lfSAoJHtzaGVldFJvd3MubGVuZ3RofSByb3cocykpYCwgJ3N0YXQtbGFiZWwnKSk7CiAgICAgIGNvbnN0IGhlYWRlcnMgPSBzaGVldFJvd3NbMF0uaGVhZGVyczsKICAgICAgY29uc3QgdGFibGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0YWJsZScpOwogICAgICB0YWJsZS5jbGFzc05hbWUgPSAnZGF0YS10YWJsZSc7CiAgICAgIGNvbnN0IHRoZWFkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgdGhlYWQuYXBwZW5kKHRleHRFbCgndGgnLCAnUm93ICMnKSwgdGV4dEVsKCd0aCcsICdMaW5rZWQgdG8gcG9zdCcpKTsKICAgICAgY29uc3QgY29sQ291bnQgPSBoZWFkZXJzID8gaGVhZGVycy5sZW5ndGggOiBNYXRoLm1heCguLi5zaGVldFJvd3MubWFwKChyKSA9PiByLnJhdy5sZW5ndGgpKTsKICAgICAgZm9yIChsZXQgaSA9IDA7IGkgPCBjb2xDb3VudDsgaSArPSAxKSB0aGVhZC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3RoJywgaGVhZGVycyAmJiBoZWFkZXJzW2ldID8gU3RyaW5nKGhlYWRlcnNbaV0pIDogYENvbCAke2kgKyAxfWApKTsKICAgICAgY29uc3QgdGhlYWRXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGhlYWQnKTsKICAgICAgdGhlYWRXcmFwLmFwcGVuZENoaWxkKHRoZWFkKTsKICAgICAgY29uc3QgdGJvZHkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0Ym9keScpOwogICAgICBzaGVldFJvd3MuZm9yRWFjaCgocikgPT4gewogICAgICAgIGNvbnN0IHRyMiA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAgICAgdHIyLmFwcGVuZCh0ZXh0RWwoJ3RkJywgU3RyaW5nKHIucm93X251bWJlcikpLCB0ZXh0RWwoJ3RkJywgci5wb3N0X2lkID8gYCMke3IucG9zdF9pZH1gIDogJ+KAlCcpKTsKICAgICAgICBmb3IgKGxldCBpID0gMDsgaSA8IGNvbENvdW50OyBpICs9IDEpIHsKICAgICAgICAgIGNvbnN0IHZhbCA9IHIucmF3W2ldOwogICAgICAgICAgdHIyLmFwcGVuZENoaWxkKHRleHRFbCgndGQnLCB2YWwgPT09IHVuZGVmaW5lZCB8fCB2YWwgPT09IG51bGwgPyAnJyA6IFN0cmluZyh2YWwpLnNsaWNlKDAsIDYwKSkpOwogICAgICAgIH0KICAgICAgICB0Ym9keS5hcHBlbmRDaGlsZCh0cjIpOwogICAgICB9KTsKICAgICAgdGFibGUuYXBwZW5kKHRoZWFkV3JhcCwgdGJvZHkpOwogICAgICBjb25zdCBzY3JvbGxXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIHNjcm9sbFdyYXAuY2xhc3NOYW1lID0gJ3RhYmxlLXNjcm9sbCc7CiAgICAgIHNjcm9sbFdyYXAuc3R5bGUubWFyZ2luQm90dG9tID0gJzE2cHgnOwogICAgICBzY3JvbGxXcmFwLmFwcGVuZENoaWxkKHRhYmxlKTsKICAgICAgd3JhcC5hcHBlbmRDaGlsZChzY3JvbGxXcmFwKTsKICAgIH0pOwogIH0KCiAgcmV0dXJuIHsgcmVuZGVyIH07Cn0pKCk7CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgQXBwIGJvb3RzdHJhcDogdGFiIHJvdXRpbmcsIGZpbHRlciBiYXIgd2lyaW5nLCB0aGVtZSB0b2dnbGUuCiAgID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwooKCkgPT4gewogIGNvbnN0IFZJRVdTID0gewogICAgZGFzaGJvYXJkOiBEYXNoYm9hcmQsCiAgICByZWNvcmRzOiBSZWNvcmRzLAogICAgY29tcGFyaXNvbjogQ29tcGFyaXNvbiwKICAgIHVwbG9hZDogVXBsb2FkLAogICAgaGlzdG9yeTogSGlzdG9yeSwKICB9OwoKICBsZXQgYWN0aXZlVGFiID0gJ2Rhc2hib2FyZCc7CgogIGZ1bmN0aW9uIHN3aXRjaFRhYih0YWIpIHsKICAgIGFjdGl2ZVRhYiA9IHRhYjsKICAgIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJy50YWItYnRuJykuZm9yRWFjaCgoYnRuKSA9PiB7CiAgICAgIGNvbnN0IGlzQWN0aXZlID0gYnRuLmRhdGFzZXQudGFiID09PSB0YWI7CiAgICAgIGJ0bi5jbGFzc0xpc3QudG9nZ2xlKCdpcy1hY3RpdmUnLCBpc0FjdGl2ZSk7CiAgICAgIGJ0bi5zZXRBdHRyaWJ1dGUoJ2FyaWEtc2VsZWN0ZWQnLCBTdHJpbmcoaXNBY3RpdmUpKTsKICAgIH0pOwogICAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnLnZpZXcnKS5mb3JFYWNoKCh2aWV3KSA9PiB7CiAgICAgIHZpZXcuY2xhc3NMaXN0LnRvZ2dsZSgnaXMtYWN0aXZlJywgdmlldy5pZCA9PT0gYHZpZXctJHt0YWJ9YCk7CiAgICB9KTsKICAgIC8vIEZpbHRlcnMgYXBwbHkgdG8gRGFzaGJvYXJkIGFuZCBEYXRhIFJlY29yZHMgKENvbXBhcmlzb25zIGhhcyBpdHMgb3duIHJhbmdlIGNvbnRyb2xzKS4KICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJCYXInKS5zdHlsZS5kaXNwbGF5ID0gKHRhYiA9PT0gJ2Rhc2hib2FyZCcgfHwgdGFiID09PSAncmVjb3JkcycpID8gJ2ZsZXgnIDogJ25vbmUnOwogICAgcmVuZGVyQWN0aXZlVmlldygpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyQWN0aXZlVmlldygpIHsKICAgIGNvbnN0IHZpZXcgPSBWSUVXU1thY3RpdmVUYWJdOwogICAgaWYgKHZpZXcgJiYgdmlldy5yZW5kZXIpIHZpZXcucmVuZGVyKCk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiBsb2FkRmlsdGVyT3B0aW9ucygpIHsKICAgIGNvbnN0IG9wdGlvbnMgPSBhd2FpdCBBcGkuZmlsdGVyT3B0aW9ucygpOwogICAgd2luZG93Ll9fZmlsdGVyT3B0aW9uc0NhY2hlID0gb3B0aW9uczsKCiAgICBjb25zdCBwbGF0Zm9ybVNlbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJQbGF0Zm9ybScpOwogICAgcGxhdGZvcm1TZWwubGVuZ3RoID0gMTsKICAgIG9wdGlvbnMucGxhdGZvcm1zLmZvckVhY2goKHApID0+IHsKICAgICAgY29uc3Qgb3B0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7CiAgICAgIG9wdC52YWx1ZSA9IHAuaWQ7CiAgICAgIG9wdC50ZXh0Q29udGVudCA9IHAubGFiZWw7CiAgICAgIHBsYXRmb3JtU2VsLmFwcGVuZENoaWxkKG9wdCk7CiAgICB9KTsKCiAgICBjb25zdCBjYW1wYWlnblNlbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJDYW1wYWlnbicpOwogICAgY2FtcGFpZ25TZWwubGVuZ3RoID0gMTsKICAgIG9wdGlvbnMuY2FtcGFpZ25UeXBlcy5mb3JFYWNoKChjKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOwogICAgICBvcHQudmFsdWUgPSBjOwogICAgICBvcHQudGV4dENvbnRlbnQgPSBjOwogICAgICBjYW1wYWlnblNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CgogICAgY29uc3QgY29udGVudFNlbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJDb250ZW50VHlwZScpOwogICAgY29udGVudFNlbC5sZW5ndGggPSAxOwogICAgb3B0aW9ucy5jb250ZW50VHlwZXMuZm9yRWFjaCgoYykgPT4gewogICAgICBjb25zdCBvcHQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsKICAgICAgb3B0LnZhbHVlID0gYzsKICAgICAgb3B0LnRleHRDb250ZW50ID0gYzsKICAgICAgY29udGVudFNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgfQoKICBmdW5jdGlvbiB3aXJlRmlsdGVyQmFyKCkgewogICAgY29uc3QgZGF0ZUZyb20gPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmlsdGVyRGF0ZUZyb20nKTsKICAgIGNvbnN0IGRhdGVUbyA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJEYXRlVG8nKTsKICAgIGNvbnN0IHBsYXRmb3JtID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZpbHRlclBsYXRmb3JtJyk7CiAgICBjb25zdCBjYW1wYWlnbiA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJDYW1wYWlnbicpOwogICAgY29uc3QgY29udGVudFR5cGUgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmlsdGVyQ29udGVudFR5cGUnKTsKICAgIGNvbnN0IGYgPSBTdGF0ZS5nZXRGaWx0ZXJzKCk7CiAgICBkYXRlRnJvbS52YWx1ZSA9IGYuZGF0ZUZyb207CiAgICBkYXRlVG8udmFsdWUgPSBmLmRhdGVUbzsKCiAgICBmdW5jdGlvbiBhcHBseSgpIHsKICAgICAgU3RhdGUuc2V0RmlsdGVycyh7CiAgICAgICAgZGF0ZUZyb206IGRhdGVGcm9tLnZhbHVlLAogICAgICAgIGRhdGVUbzogZGF0ZVRvLnZhbHVlLAogICAgICAgIHBsYXRmb3JtOiBwbGF0Zm9ybS52YWx1ZSwKICAgICAgICBjYW1wYWlnblR5cGU6IGNhbXBhaWduLnZhbHVlLAogICAgICAgIGNvbnRlbnRUeXBlOiBjb250ZW50VHlwZS52YWx1ZSwKICAgICAgfSk7CiAgICB9CiAgICBbZGF0ZUZyb20sIGRhdGVUbywgcGxhdGZvcm0sIGNhbXBhaWduLCBjb250ZW50VHlwZV0uZm9yRWFjaCgoZWwpID0+IGVsLmFkZEV2ZW50TGlzdGVuZXIoJ2NoYW5nZScsIGFwcGx5KSk7CgogICAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnI2ZpbHRlclByZXNldHMgYnV0dG9uJykuZm9yRWFjaCgoYnRuKSA9PiB7CiAgICAgIGJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsKICAgICAgICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcjZmlsdGVyUHJlc2V0cyBidXR0b24nKS5mb3JFYWNoKChiKSA9PiBiLmNsYXNzTGlzdC5yZW1vdmUoJ2lzLWFjdGl2ZScpKTsKICAgICAgICBidG4uY2xhc3NMaXN0LmFkZCgnaXMtYWN0aXZlJyk7CiAgICAgICAgY29uc3QgcHJlc2V0ID0gYnRuLmRhdGFzZXQucHJlc2V0OwogICAgICAgIGNvbnN0IHRvZGF5ID0gbmV3IERhdGUoKTsKICAgICAgICBjb25zdCB0byA9IHRvZGF5LnRvSVNPU3RyaW5nKCkuc2xpY2UoMCwgMTApOwogICAgICAgIGxldCBmcm9tOwogICAgICAgIGlmIChwcmVzZXQgPT09ICdhbGwnKSB7CiAgICAgICAgICBjb25zdCBtaW4gPSAod2luZG93Ll9fZmlsdGVyT3B0aW9uc0NhY2hlICYmIHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZS5kYXRlUmFuZ2UubWluKSB8fCB0bzsKICAgICAgICAgIGZyb20gPSBtaW47CiAgICAgICAgfSBlbHNlIHsKICAgICAgICAgIGNvbnN0IGQgPSBuZXcgRGF0ZSh0b2RheSk7CiAgICAgICAgICBkLnNldERhdGUoZC5nZXREYXRlKCkgLSAoTnVtYmVyKHByZXNldCkgLSAxKSk7CiAgICAgICAgICBmcm9tID0gZC50b0lTT1N0cmluZygpLnNsaWNlKDAsIDEwKTsKICAgICAgICB9CiAgICAgICAgZGF0ZUZyb20udmFsdWUgPSBmcm9tOwogICAgICAgIGRhdGVUby52YWx1ZSA9IHRvOwogICAgICAgIGFwcGx5KCk7CiAgICAgIH0pOwogICAgfSk7CgogICAgU3RhdGUub25DaGFuZ2UoKCkgPT4gewogICAgICBpZiAoYWN0aXZlVGFiID09PSAnZGFzaGJvYXJkJykgRGFzaGJvYXJkLnJlbmRlcigpOwogICAgICBpZiAoYWN0aXZlVGFiID09PSAncmVjb3JkcycpIFJlY29yZHMucmVuZGVyKCk7CiAgICB9KTsKICB9CgogIGZ1bmN0aW9uIHdpcmVUYWJzKCkgewogICAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnLnRhYi1idG4nKS5mb3JFYWNoKChidG4pID0+IHsKICAgICAgYnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gc3dpdGNoVGFiKGJ0bi5kYXRhc2V0LnRhYikpOwogICAgfSk7CiAgfQoKICBmdW5jdGlvbiB3aXJlVGhlbWUoKSB7CiAgICBjb25zdCB0b2dnbGUgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndGhlbWVUb2dnbGUnKTsKICAgIGNvbnN0IGljb25TbG90ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3RoZW1lVG9nZ2xlSWNvbicpOwogICAgY29uc3Qgc2V0SWNvbiA9IChuYW1lKSA9PiB7IGljb25TbG90LmlubmVySFRNTCA9IGA8aSBkYXRhLWx1Y2lkZT0iJHtuYW1lfSIgc3R5bGU9IndpZHRoOjE2cHg7aGVpZ2h0OjE2cHg7Ij48L2k+YDsgfTsKICAgIGNvbnN0IHN0b3JlZCA9IGxvY2FsU3RvcmFnZS5nZXRJdGVtKCdscnMtdGhlbWUnKTsKICAgIGlmIChzdG9yZWQpIHsKICAgICAgZG9jdW1lbnQuZG9jdW1lbnRFbGVtZW50LnNldEF0dHJpYnV0ZSgnZGF0YS10aGVtZScsIHN0b3JlZCk7CiAgICAgIHNldEljb24oc3RvcmVkID09PSAnZGFyaycgPyAnc3VuJyA6ICdtb29uJyk7CiAgICB9CiAgICB0b2dnbGUuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7CiAgICAgIGNvbnN0IHByZWZlcnNEYXJrID0gd2luZG93Lm1hdGNoTWVkaWEoJyhwcmVmZXJzLWNvbG9yLXNjaGVtZTogZGFyayknKS5tYXRjaGVzOwogICAgICBjb25zdCBjdXJyZW50ID0gZG9jdW1lbnQuZG9jdW1lbnRFbGVtZW50LmdldEF0dHJpYnV0ZSgnZGF0YS10aGVtZScpIHx8IChwcmVmZXJzRGFyayA/ICdkYXJrJyA6ICdsaWdodCcpOwogICAgICBjb25zdCBuZXh0ID0gY3VycmVudCA9PT0gJ2RhcmsnID8gJ2xpZ2h0JyA6ICdkYXJrJzsKICAgICAgZG9jdW1lbnQuZG9jdW1lbnRFbGVtZW50LnNldEF0dHJpYnV0ZSgnZGF0YS10aGVtZScsIG5leHQpOwogICAgICBsb2NhbFN0b3JhZ2Uuc2V0SXRlbSgnbHJzLXRoZW1lJywgbmV4dCk7CiAgICAgIHNldEljb24obmV4dCA9PT0gJ2RhcmsnID8gJ3N1bicgOiAnbW9vbicpOwogICAgICBDaGFydHMuZGVzdHJveUFsbCgpOwogICAgICByZW5kZXJBY3RpdmVWaWV3KCk7CiAgICB9KTsKICB9CgogIHdpbmRvdy5hZGRFdmVudExpc3RlbmVyKCdscnM6ZGF0YS11cGRhdGVkJywgYXN5bmMgKCkgPT4gewogICAgYXdhaXQgbG9hZEZpbHRlck9wdGlvbnMoKTsKICAgIHJlbmRlckFjdGl2ZVZpZXcoKTsKICB9KTsKCiAgLy8gLS0tLS0tLS0tLSBBdXRoIHNjcmVlbiAtLS0tLS0tLS0tCiAgbGV0IGFwcEluaXRpYWxpemVkID0gZmFsc2U7CgogIGZ1bmN0aW9uIHNob3dBdXRoU2NyZWVuKCkgewogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2F1dGhTY3JlZW4nKS5zdHlsZS5kaXNwbGF5ID0gJ2ZsZXgnOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2FwcFNoZWxsJykuc3R5bGUuZGlzcGxheSA9ICdub25lJzsKICAgIGNvbnN0IGNvZGVJbnB1dCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhdXRoQ29kZScpOwogICAgY29kZUlucHV0LnZhbHVlID0gJyc7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXV0aEVycm9yJykudGV4dENvbnRlbnQgPSAnJzsKICAgIGNvZGVJbnB1dC5mb2N1cygpOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gc2hvd0FwcCgpIHsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhdXRoU2NyZWVuJykuc3R5bGUuZGlzcGxheSA9ICdub25lJzsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhcHBTaGVsbCcpLnN0eWxlLmRpc3BsYXkgPSAnJzsKICAgIGlmICghYXBwSW5pdGlhbGl6ZWQpIHsKICAgICAgYXBwSW5pdGlhbGl6ZWQgPSB0cnVlOwogICAgICB3aXJlVGFicygpOwogICAgICB3aXJlVGhlbWUoKTsKICAgICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2xvZ291dEJ0bicpLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgYXN5bmMgKCkgPT4gewogICAgICAgIGF3YWl0IEFwaS5hdXRoTG9nb3V0KCk7CiAgICAgICAgYXBwSW5pdGlhbGl6ZWQgPSBmYWxzZTsKICAgICAgICBzaG93QXV0aFNjcmVlbigpOwogICAgICB9KTsKICAgICAgYXdhaXQgbG9hZEZpbHRlck9wdGlvbnMoKTsKICAgICAgd2lyZUZpbHRlckJhcigpOwogICAgICBzd2l0Y2hUYWIoJ2Rhc2hib2FyZCcpOwogICAgfSBlbHNlIHsKICAgICAgYXdhaXQgbG9hZEZpbHRlck9wdGlvbnMoKTsKICAgICAgcmVuZGVyQWN0aXZlVmlldygpOwogICAgfQogIH0KCiAgYXN5bmMgZnVuY3Rpb24gc3VibWl0QXV0aCgpIHsKICAgIGNvbnN0IGVycm9yRWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXV0aEVycm9yJyk7CiAgICBjb25zdCBidG4gPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXV0aFN1Ym1pdEJ0bicpOwogICAgY29uc3QgY29kZUlucHV0ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2F1dGhDb2RlJyk7CiAgICBlcnJvckVsLnRleHRDb250ZW50ID0gJyc7CiAgICBidG4uZGlzYWJsZWQgPSB0cnVlOwogICAgYnRuLnRleHRDb250ZW50ID0gJ0NoZWNraW5n4oCmJzsKICAgIHRyeSB7CiAgICAgIGF3YWl0IEFwaS5hdXRoTG9naW4oY29kZUlucHV0LnZhbHVlKTsKICAgICAgYXdhaXQgc2hvd0FwcCgpOwogICAgfSBjYXRjaCAoZXJyKSB7CiAgICAgIGVycm9yRWwudGV4dENvbnRlbnQgPSBlcnIubWVzc2FnZTsKICAgIH0gZmluYWxseSB7CiAgICAgIGJ0bi5kaXNhYmxlZCA9IGZhbHNlOwogICAgICBidG4uaW5uZXJIVE1MID0gJzxpIGRhdGEtbHVjaWRlPSJhcnJvdy1yaWdodCIgc3R5bGU9IndpZHRoOjE0cHg7aGVpZ2h0OjE0cHg7Ij48L2k+IEVudGVyJzsKICAgIH0KICB9CgogIGZ1bmN0aW9uIHdpcmVBdXRoRm9ybSgpIHsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhdXRoU3VibWl0QnRuJykuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBzdWJtaXRBdXRoKTsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhdXRoQ29kZScpLmFkZEV2ZW50TGlzdGVuZXIoJ2tleWRvd24nLCAoZSkgPT4geyBpZiAoZS5rZXkgPT09ICdFbnRlcicpIHN1Ym1pdEF1dGgoKTsgfSk7CiAgfQoKICB3aW5kb3cuYWRkRXZlbnRMaXN0ZW5lcignbHJzOnNpZ25lZC1vdXQnLCAoKSA9PiB7CiAgICBhcHBJbml0aWFsaXplZCA9IGZhbHNlOwogICAgc2hvd0F1dGhTY3JlZW4oKTsKICB9KTsKCiAgYXN5bmMgZnVuY3Rpb24gaW5pdCgpIHsKICAgIHdpcmVBdXRoRm9ybSgpOwogICAgY29uc3QgeyBhdXRoZW50aWNhdGVkIH0gPSBhd2FpdCBBcGkuYXV0aE1lKCk7CiAgICBpZiAoYXV0aGVudGljYXRlZCkgYXdhaXQgc2hvd0FwcCgpOwogICAgZWxzZSBzaG93QXV0aFNjcmVlbigpOwogIH0KCiAgLy8gSWNvbnMgYXJlIHBsYWNlZCBhcyA8aSBkYXRhLWx1Y2lkZT0iLi4uIj4gcGxhY2Vob2xkZXJzIHRocm91Z2hvdXQgdGhlIGR5bmFtaWNhbGx5CiAgLy8gcmVuZGVyZWQgVUk7IEx1Y2lkZSByZXBsYWNlcyBlYWNoIHdpdGggYW4gaW5saW5lIFNWRy4gUmF0aGVyIHRoYW4gcmVtZW1iZXJpbmcgdG8gY2FsbAogIC8vIHRoaXMgYWZ0ZXIgZXZlcnkgc2luZ2xlIHJlbmRlciwgb25lIG9ic2VydmVyIGNhdGNoZXMgZXZlcnkgRE9NIGNoYW5nZSB0aGF0IGNvdWxkIGhhdmUKICAvLyBpbnRyb2R1Y2VkIGEgbmV3IHBsYWNlaG9sZGVyLgogIGlmICh3aW5kb3cubHVjaWRlKSB7CiAgICB3aW5kb3cubHVjaWRlLmNyZWF0ZUljb25zKCk7CiAgICAvLyBjcmVhdGVJY29ucygpIHJlcGxhY2VzIDxpIGRhdGEtbHVjaWRlPiBwbGFjZWhvbGRlcnMgd2l0aCA8c3ZnPiDigJQgaXRzZWxmIGEgRE9NCiAgICAvLyBtdXRhdGlvbi4gV2l0aG91dCBkaXNjb25uZWN0aW5nIGZpcnN0LCB0aGF0IHdyaXRlIHJlLXRyaWdnZXJzIHRoaXMgc2FtZSBvYnNlcnZlcgogICAgLy8gZm9yZXZlciAoYW4gaW5maW5pdGUgbXV0YXRlL29ic2VydmUgbG9vcCB0aGF0IHBlZ3MgdGhlIENQVSBhbmQgY3Jhc2hlcyB0aGUgdGFiKS4KICAgIC8vIERpc2Nvbm5lY3RpbmcgYmVmb3JlIGVhY2ggcGFzcyBhbmQgcmVjb25uZWN0aW5nIGFmdGVyLCBwbHVzIGJhdGNoaW5nIGJ1cnN0cyBvZgogICAgLy8gbXV0YXRpb25zIGludG8gYSBzaW5nbGUgbWljcm90YXNrLCBicmVha3MgdGhlIGN5Y2xlLgogICAgbGV0IGljb25zU2NoZWR1bGVkID0gZmFsc2U7CiAgICBjb25zdCBpY29uT2JzZXJ2ZXIgPSBuZXcgTXV0YXRpb25PYnNlcnZlcigoKSA9PiB7CiAgICAgIGlmIChpY29uc1NjaGVkdWxlZCkgcmV0dXJuOwogICAgICBpY29uc1NjaGVkdWxlZCA9IHRydWU7CiAgICAgIHF1ZXVlTWljcm90YXNrKCgpID0+IHsKICAgICAgICBpY29uc1NjaGVkdWxlZCA9IGZhbHNlOwogICAgICAgIGljb25PYnNlcnZlci5kaXNjb25uZWN0KCk7CiAgICAgICAgd2luZG93Lmx1Y2lkZS5jcmVhdGVJY29ucygpOwogICAgICAgIGljb25PYnNlcnZlci5vYnNlcnZlKGRvY3VtZW50LmJvZHksIHsgY2hpbGRMaXN0OiB0cnVlLCBzdWJ0cmVlOiB0cnVlIH0pOwogICAgICB9KTsKICAgIH0pOwogICAgaWNvbk9ic2VydmVyLm9ic2VydmUoZG9jdW1lbnQuYm9keSwgeyBjaGlsZExpc3Q6IHRydWUsIHN1YnRyZWU6IHRydWUgfSk7CiAgfQoKICBpbml0KCk7Cn0pKCk7Cjwvc2NyaXB0Pgo8L2JvZHk+CjwvaHRtbD4K';
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
