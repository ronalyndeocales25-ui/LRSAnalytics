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

-- Manually-entered weekly follower totals per platform — entirely separate
-- from spreadsheet uploads (posts/post_metrics/raw_rows above). One row per
-- platform per week; re-entering the same platform+week updates it in
-- place rather than creating a duplicate.
CREATE TABLE IF NOT EXISTS followers_history (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  platform TEXT NOT NULL,
  entry_date TEXT NOT NULL,
  followers_count INTEGER NOT NULL,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  UNIQUE(platform, entry_date)
);
CREATE INDEX IF NOT EXISTS idx_followers_platform ON followers_history(platform);
CREATE INDEX IF NOT EXISTS idx_followers_date ON followers_history(entry_date);
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
  { id: 'x', label: 'X', color: '#4a3aa7', darkColor: '#9085e9', groupAliases: ['x', 'twitter', 'x twitter', 'x (twitter)'] },
  { id: 'fb_group', label: 'FB Group', color: '#e34948', darkColor: '#e66767', groupAliases: ['fb group', 'facebook group', 'fbgroup', 'fb groups'] },
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
  'reactions': 'engagement', // FB Group's own block reports "Reactions" rather than "Engagement"
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
  'share': 'shares', // singular variant seen in FB Group's own block
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
  x: [
    { key: 'impressions', label: 'Impressions' },
    { key: 'engagement', label: 'Engagements' },
    { key: 'clicks', label: 'Clicks' },
    { key: 'posting_link', label: 'Link' },
  ],
  fb_group: [
    { key: 'engagement', label: 'Reactions' },
    { key: 'comments', label: 'Comments' },
    { key: 'shares', label: 'Shares' },
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

// A posting link alone counts as data: it's real evidence the post exists on
// that platform even on weeks the platform's own analytics come back blank
// (e.g. a sheet cell reading "Analytics unavailable"), so the post still
// gets attributed there instead of showing no platform at all.
function metricBagHasData(bag) {
  return Object.values(bag).some((v) => v !== null && v !== undefined);
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
 * Facebook and FB Group deliberately share one spreadsheet column block —
 * the source sheet only has a "FACEBOOK" header, never a separate "FB
 * GROUP" one, and per the business rule that stays true even after adding
 * FB Group as its own platform. Instead, each row's free-text "Platforms"
 * identifier decides which of the two the Facebook block's numbers belong
 * to: "Facebook" alone -> facebook only (the pre-existing default, so any
 * row that doesn't mention FB Group is completely unaffected), "FB Group"
 * alone -> fb_group only, both mentioned together (either order) -> both,
 * reusing the exact same metric values for each rather than asking for a
 * second set of columns. Every other platform's column block is untouched.
 */
function splitFacebookGroupPlatformBags(platformBags, platformsRaw) {
  const bag = platformBags.facebook;
  if (!bag || !metricBagHasData(bag)) return;

  // If the sheet also has its own dedicated FB GROUP column block with real
  // data for this row (not every sheet does — some only have a FACEBOOK
  // block and rely entirely on the Platforms text, which is what the rest
  // of this function handles), that block's own numbers are authoritative
  // and must never be overwritten with a copy of Facebook's.
  if (platformBags.fb_group && metricBagHasData(platformBags.fb_group)) return;

  const tokens = String(platformsRaw || '').split(',').map((s) => s.trim()).filter(Boolean);
  let mentionsFacebook = false;
  let mentionsGroup = false;
  tokens.forEach((token) => {
    const match = findPlatformByGroupLabel(token);
    if (!match) return;
    if (match.id === 'facebook') mentionsFacebook = true;
    if (match.id === 'fb_group') mentionsGroup = true;
  });

  if (mentionsGroup && mentionsFacebook) {
    platformBags.fb_group = { ...bag };
  } else if (mentionsGroup && !mentionsFacebook) {
    platformBags.fb_group = bag;
    delete platformBags.facebook;
  }
  // Facebook mentioned alone, or neither explicitly recognized in the
  // Platforms text — leaves the bag under 'facebook' (today's behavior).
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

  splitFacebookGroupPlatformBags(platformBags, identifiers.platforms_raw);

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
    // Every row that will actually be inserted as a brand-new record (i.e. not a
    // dupe-within-file and not an exact match of something already saved) — the
    // preview's counterpart to `duplicates` above, so uploaders can see exactly
    // what's new before committing, not just a generic sample.
    newRecords: newPosts.map((p) => ({
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

/**
 * Every row matching the Data Records table's current search/filters, with
 * no LIMIT/OFFSET — the export counterpart to listRecordsTable() above,
 * sharing the same buildPostWhereClause() filter logic so the two can never
 * drift out of sync with each other.
 */
function exportRecordsRows({ platform, dateFrom, dateTo, campaignType, contentType, search }) {
  const { clauses, params } = buildPostWhereClause({ dateFrom, dateTo, campaignType, contentType, search });

  if (!platform || platform === 'all') {
    const where = clauses.length ? `WHERE ${clauses.join(' AND ')}` : '';
    const rows = db.prepare(`
      SELECT p.publish_date, p.caption, p.campaign_type, p.content_type, p.created_at, p.updated_at,
             (SELECT GROUP_CONCAT(DISTINCT pm.platform) FROM post_metrics pm WHERE pm.post_id = p.id) AS platform_ids
      FROM posts p
      ${where}
      ORDER BY p.publish_date DESC, p.id DESC
    `).all(...params);

    const columns = [
      { key: 'publish_date', label: 'Date' },
      { key: 'platform_ids', label: 'Platforms' },
      { key: 'caption', label: 'Caption' },
      { key: 'campaign_type', label: 'Campaign' },
      { key: 'content_type', label: 'Content Type' },
      { key: 'status', label: 'Status' },
      { key: 'updated_at', label: 'Last Updated' },
    ];
    const outRows = rows.map((r) => ({
      publish_date: r.publish_date,
      platform_ids: r.platform_ids || '',
      caption: r.caption || '',
      campaign_type: r.campaign_type || '',
      content_type: r.content_type || '',
      status: rowStatus(r.created_at, r.updated_at),
      updated_at: r.updated_at,
    }));
    return { columns, rows: outRows };
  }

  const where = `WHERE ${[...clauses, 'pm.platform = ?'].join(' AND ')}`;
  const curated = PLATFORM_RECORD_COLUMNS[platform] || [];
  const rows = db.prepare(`
    SELECT p.publish_date, p.caption, p.campaign_type, p.content_type, p.created_at, p.updated_at,
           pm.views, pm.reach, pm.impressions, pm.engagement, pm.clicks, pm.followers_gained,
           pm.watch_time_seconds, pm.shares, pm.comments, pm.saves, pm.posting_link
    FROM posts p JOIN post_metrics pm ON pm.post_id = p.id
    ${where}
    ORDER BY p.publish_date DESC, p.id DESC
  `).all(...params, platform);

  const columns = [
    { key: 'publish_date', label: 'Date' },
    { key: 'caption', label: 'Caption' },
    { key: 'campaign_type', label: 'Campaign' },
    { key: 'content_type', label: 'Content Type' },
    ...curated.filter((c) => c.key !== 'posting_link'),
    { key: 'posting_link', label: 'Link' },
    { key: 'status', label: 'Status' },
    { key: 'updated_at', label: 'Last Updated' },
  ];
  const outRows = rows.map((r) => {
    const base = {
      publish_date: r.publish_date,
      caption: r.caption || '',
      campaign_type: r.campaign_type || '',
      content_type: r.content_type || '',
      status: rowStatus(r.created_at, r.updated_at),
      updated_at: r.updated_at,
      posting_link: r.posting_link || '',
    };
    curated.forEach((c) => { if (c.key !== 'posting_link') base[c.key] = r[c.key]; });
    return base;
  });
  return { columns, rows: outRows };
}

/** Quoted-comma-join CSV — no library needed for something this simple. Shared by every server-side export (client-side tables port an equivalent version of this same logic locally instead of round-tripping). */
function toCSV(rows, columns) {
  const escape = (v) => {
    const s = v === null || v === undefined ? '' : String(v);
    return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const lines = [columns.map((c) => escape(c.label)).join(',')];
  rows.forEach((row) => lines.push(columns.map((c) => escape(row[c.key])).join(',')));
  return lines.join('\r\n');
}

function sendCSV(res, filename, rows, columns) {
  res.setHeader('Content-Type', 'text/csv; charset=utf-8');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.send(toCSV(rows, columns));
}

/** Reuses exceljs — already a dependency for reading uploaded spreadsheets — to write .xlsx for export too, so no new npm package is needed either way. */
async function sendXLSX(res, filename, rows, columns, sheetName) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet(sheetName || 'Export');
  sheet.columns = columns.map((c) => ({ header: c.label, key: c.key, width: Math.max(12, String(c.label).length + 2) }));
  rows.forEach((row) => sheet.addRow(row));
  sheet.getRow(1).font = { bold: true };
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  await workbook.xlsx.write(res);
  res.end();
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

/* ------------------------------------------------------------
   Duplicate cleanup

   Re-uploading a weekly sheet whose numbers have moved imports
   each changed row as a brand-new record (the full-record
   fingerprint only skips a row when every metric is identical
   too), so Data Records ends up showing the same post more than
   once. This finds those groups and, on confirm, keeps the most
   recently imported copy and deletes the rest. Deletion reuses
   deletePostStmt, so raw_rows are retained (detached) exactly
   like deletePost() and every original import stays recoverable
   from Upload History.

   Two records are treated as the same post when they share a
   publish date and a normalized caption (captionKey below):
   Unicode NFKC folding so styled/"fancy" letters match plain
   ones, curly quotes/dashes/ellipsis folded to ASCII, zero-width
   and emoji-variation marks stripped, and all whitespace
   collapsed -- so an edited row whose newlines or smart quotes
   shifted still matches its original, while ordinary punctuation
   and emoji are kept so genuinely different same-day captions do
   not collide. An older copy is only deleted when it came from a
   different upload than the survivor AND the survivor already
   covers all of its platforms, so neither an in-file sibling nor
   a row with unique platform data is ever dropped. Blank captions
   are skipped.
   ------------------------------------------------------------ */
function captionKey(s) {
  return (s || '')
    .normalize('NFKC')
    .toLowerCase()
    .replace(/[\u200B-\u200D\u2060\uFEFF\uFE0E\uFE0F]/g, '') // zero-width chars + emoji variation selectors
    .replace(/[\u2018\u2019\u02BC\u2032`\u00B4]/g, "'") // apostrophe variants
    .replace(/[\u201C\u201D\u2033]/g, '"')                   // double-quote variants
    .replace(/[\u2013\u2014\u2015]/g, '-')                   // en/em dashes
    .replace(/\u2026/g, '...')                               // ellipsis
    .replace(/\s+/g, ' ')                                    // collapse newlines / runs of spaces
    .trim();
}

function findDuplicateRecordGroups() {
  const rows = db
    .prepare(`
      SELECT p.id AS post_id, p.upload_id, p.publish_date, p.caption, p.created_at, p.updated_at,
             (SELECT GROUP_CONCAT(platform) FROM (
                SELECT DISTINCT platform FROM post_metrics WHERE post_id = p.id ORDER BY platform
              )) AS platform_ids
      FROM posts p
    `)
    .all();

  const groups = new Map();
  for (const r of rows) {
    const ck = captionKey(r.caption);
    if (!ck) continue;
    const key = r.publish_date + ' | ' + ck;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(r);
  }

  const asRef = (m) => ({
    postId: m.post_id,
    uploadId: m.upload_id,
    platformIds: m.platform_ids ? m.platform_ids.split(',') : [],
    importedAt: m.created_at,
    status: rowStatus(m.created_at, m.updated_at),
  });

  const out = [];
  for (const members of groups.values()) {
    if (members.length < 2) continue;
    members.sort((a, b) =>
      a.created_at === b.created_at ? b.post_id - a.post_id : (a.created_at < b.created_at ? 1 : -1)
    );
    const keep = members[0];
    const keepSet = new Set(keep.platform_ids ? keep.platform_ids.split(',') : []);
    // Drop an older copy only when BOTH hold:
    //  - it came from a different upload than the survivor (the signature of a
    //    re-import; two same-key rows inside one imported file are left alone),
    //  - the survivor already covers all of its platforms, so the cleanup never
    //    loses a platform's worth of data.
    const remove = members.slice(1).filter((m) => {
      if (m.upload_id === keep.upload_id) return false;
      const mSet = m.platform_ids ? m.platform_ids.split(',') : [];
      return mSet.every((pf) => keepSet.has(pf));
    });
    if (!remove.length) continue;
    out.push({
      publishDate: keep.publish_date,
      caption: keep.caption,
      keep: asRef(keep),
      remove: remove.map(asRef),
    });
  }
  // Most recent posts first, matching the Data Records table order.
  out.sort((a, b) => (a.publishDate < b.publishDate ? 1 : a.publishDate > b.publishDate ? -1 : 0));
  return out;
}

/** Read-only summary of what a cleanup would remove — powers the confirm prompt. */
function previewDuplicateCleanup() {
  const groups = findDuplicateRecordGroups();
  return {
    groupCount: groups.length,
    removeCount: groups.reduce((n, g) => n + g.remove.length, 0),
    editedInRemoveCount: groups.reduce(
      (n, g) => n + g.remove.filter((r) => r.status === 'edited').length,
      0
    ),
    groups,
  };
}

/** Deletes every older copy in each duplicate group, in one transaction. */
function runDuplicateCleanup() {
  const groups = findDuplicateRecordGroups();
  const ids = groups.flatMap((g) => g.remove.map((r) => r.postId));
  db.exec('BEGIN');
  try {
    for (const id of ids) deletePostStmt.run(id);
    db.exec('COMMIT');
  } catch (err) {
    db.exec('ROLLBACK');
    throw err;
  }
  return { groupCount: groups.length, removedCount: ids.length };
}

/**
 * Deletes every uploaded post record and all upload history in one
 * transaction — posts, post_metrics, raw_rows, upload_errors, uploads — so the
 * data can be rebuilt from a single fresh import. Deliberately leaves
 * followers_history (manually entered, separate system), the access code, and
 * login sessions alone. Sequence counters are reset so the next import starts
 * at id 1.
 */
function wipeUploadedData() {
  const tables = ['post_metrics', 'raw_rows', 'upload_errors', 'posts', 'uploads'];
  const counts = {};
  db.exec('BEGIN');
  try {
    for (const t of tables) {
      counts[t] = db.prepare(`SELECT COUNT(*) AS c FROM ${t}`).get().c;
      db.exec(`DELETE FROM ${t}`);
    }
    db.exec(`DELETE FROM sqlite_sequence WHERE name IN ('post_metrics','raw_rows','upload_errors','posts','uploads')`);
    db.exec('COMMIT');
  } catch (err) {
    db.exec('ROLLBACK');
    throw err;
  }
  return { ok: true, deleted: counts };
}

/* ============================================================
   Analytics service — all aggregation/comparison queries
   ============================================================ */
const SUM_METRICS_SQL = CANONICAL_METRIC_KEYS.map((k) => `SUM(pm.${k}) AS ${k}`).join(', ');
const METRIC_LABELS = Object.fromEntries(CANONICAL_METRICS.map((m) => [m.key, m.label]));

function metricExpr(key) {
  return `pm.${key}`;
}

function metricLabel(key) {
  return METRIC_LABELS[key] || key;
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
  return result;
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
    return out;
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
  return db.prepare(sql).all(...params);
}

function topPosts({ dateFrom, dateTo, filters = {}, sortBy = 'engagement', limit = 10 }) {
  const sortKey = CANONICAL_METRIC_KEYS.includes(sortBy) ? sortBy : 'engagement';
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
    return row;
  });
}

/** Breakdown of the ten canonical metrics grouped by campaign type (Ads/Organic/etc). */
function campaignBreakdown(filters) {
  const { where, params } = buildFilter(filters);
  const clause = where ? `${where} AND p.campaign_type IS NOT NULL AND p.campaign_type != ''` : "WHERE p.campaign_type IS NOT NULL AND p.campaign_type != ''";
  const sql = `
    SELECT p.campaign_type AS campaign_type, ${SUM_METRICS_SQL}, COUNT(DISTINCT p.id) AS post_count
    FROM post_metrics pm JOIN posts p ON p.id = pm.post_id
    ${clause}
    GROUP BY p.campaign_type
    ORDER BY p.campaign_type ASC`;
  return db.prepare(sql).all(...params);
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
  return db.prepare(sql).all(...params);
}

/**
 * The KPI stats for the Dashboard's metric-focused KPI cards: Highest,
 * Average, Total, Number of Posts, and Best Performing Post(s) — the last
 * one is every post tied at the highest value, not just whichever the
 * database happened to return first (capped at 20 so a degenerate
 * all-zero metric doesn't blow up the card). There's no Lowest Performing
 * Post anymore — it was removed as a card.
 */
function metricSummary({ dateFrom, dateTo, filters = {}, metric }) {
  const key = CANONICAL_METRIC_KEYS.includes(metric) ? metric : 'engagement';
  const expr = metricExpr(key);
  const { where, params } = buildFilter({ ...filters, dateFrom, dateTo });
  const notNull = `${expr} IS NOT NULL`;
  const fullWhere = where ? `${where} AND ${notNull}` : `WHERE ${notNull}`;

  const agg = db
    .prepare(`
      SELECT MAX(${expr}) AS highest, AVG(${expr}) AS average,
             SUM(${expr}) AS total, COUNT(*) AS post_count
      FROM post_metrics pm JOIN posts p ON p.id = pm.post_id
      ${fullWhere}
    `)
    .get(...params);

  const bestPosts = agg.highest === null
    ? []
    : db
        .prepare(`
          SELECT p.id AS post_id, p.publish_date, p.caption, p.campaign_type, p.content_type, pm.platform,
                 ${expr} AS value, (SELECT MAX(rr.id) FROM raw_rows rr WHERE rr.post_id = p.id) AS raw_row_id
          FROM post_metrics pm JOIN posts p ON p.id = pm.post_id
          ${fullWhere} AND ${expr} = ?
          ORDER BY p.publish_date DESC
          LIMIT 20
        `)
        .all(...params, agg.highest);

  return {
    metric: key,
    label: metricLabel(key),
    unit: key === 'watch_time_seconds' ? 'duration' : 'number',
    total: agg.total || 0,
    average: agg.average || 0,
    highest: agg.highest ?? null,
    postCount: agg.post_count || 0,
    bestPosts,
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

  return CANONICAL_METRIC_KEYS
    .filter((k) => row[`has_${k}`])
    .map((k) => ({ key: k, label: metricLabel(k), unit: k === 'watch_time_seconds' ? 'duration' : 'number' }));
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
    allPlatforms: PLATFORMS, // every supported platform, regardless of upload history — e.g. for the Followers form, which isn't upload-driven
    campaignTypes,
    contentTypes: [...contentTypeSet].sort(),
    dateRange: { min: dateRow.min, max: dateRow.max },
  };
}

/* ============================================================
   Followers service — manual weekly follower-count entry per platform.
   Entirely independent of the upload/posts/post_metrics pipeline above:
   nothing here is touched by, or touches, a spreadsheet import.
   ============================================================ */
const upsertFollowersStmt = db.prepare(`
  INSERT INTO followers_history (platform, entry_date, followers_count, updated_at)
  VALUES (@platform, @entryDate, @followersCount, datetime('now'))
  ON CONFLICT(platform, entry_date) DO UPDATE SET
    followers_count = excluded.followers_count,
    updated_at = datetime('now')
`);
const updateFollowersByIdStmt = db.prepare(`
  UPDATE followers_history SET platform=@platform, entry_date=@entryDate, followers_count=@followersCount, updated_at=datetime('now')
  WHERE id=@id
`);
const deleteFollowersStmt = db.prepare('DELETE FROM followers_history WHERE id = ?');

function validateFollowersInput({ platform, entryDate, followersCount }) {
  if (!PLATFORM_IDS.includes(platform)) throw new Error(`Unknown platform "${platform}".`);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(entryDate || ''))) throw new Error('A valid date is required.');
  const count = Number(followersCount);
  if (!Number.isFinite(count) || count < 0 || !Number.isInteger(count)) {
    throw new Error('Followers count must be a whole number, 0 or greater.');
  }
  return { platform, entryDate, followersCount: count };
}

/** All entries, optionally filtered by platform/date range, newest first. */
function listFollowers({ platform, dateFrom, dateTo } = {}) {
  const clauses = [];
  const params = [];
  if (platform && platform !== 'all') {
    clauses.push('platform = ?');
    params.push(platform);
  }
  if (dateFrom) {
    clauses.push('entry_date >= ?');
    params.push(dateFrom);
  }
  if (dateTo) {
    clauses.push('entry_date <= ?');
    params.push(dateTo);
  }
  const where = clauses.length ? `WHERE ${clauses.join(' AND ')}` : '';
  return db
    .prepare(`SELECT * FROM followers_history ${where} ORDER BY entry_date DESC, platform ASC`)
    .all(...params);
}

function createOrUpdateFollowers(input) {
  const clean = validateFollowersInput(input);
  upsertFollowersStmt.run(clean);
  return db.prepare('SELECT * FROM followers_history WHERE platform = ? AND entry_date = ?').get(clean.platform, clean.entryDate);
}

function updateFollowersEntry(id, input) {
  const existing = db.prepare('SELECT id FROM followers_history WHERE id = ?').get(id);
  if (!existing) throw new Error('Entry not found.');
  const clean = validateFollowersInput(input);
  try {
    updateFollowersByIdStmt.run({ ...clean, id });
  } catch (err) {
    throw new Error('Another entry already exists for that platform and date.');
  }
  return db.prepare('SELECT * FROM followers_history WHERE id = ?').get(id);
}

function deleteFollowersEntry(id) {
  const info = deleteFollowersStmt.run(id);
  if (info.changes === 0) throw new Error('Entry not found.');
}

/**
 * Week-over-week growth per platform, sorted chronologically: each entry's
 * change (delta) and percent change vs. the previous entry for that same
 * platform. Powers the Follower Growth chart/comparison — always computed
 * from whatever manual entries exist, for any date range.
 */
function followersGrowth({ platform, dateFrom, dateTo } = {}) {
  const clauses = [];
  const params = [];
  if (platform && platform !== 'all') {
    clauses.push('platform = ?');
    params.push(platform);
  }
  if (dateFrom) {
    clauses.push('entry_date >= ?');
    params.push(dateFrom);
  }
  if (dateTo) {
    clauses.push('entry_date <= ?');
    params.push(dateTo);
  }
  const where = clauses.length ? `WHERE ${clauses.join(' AND ')}` : '';
  const rows = db
    .prepare(`SELECT * FROM followers_history ${where} ORDER BY platform ASC, entry_date ASC`)
    .all(...params);

  const byPlatform = {};
  rows.forEach((r) => {
    (byPlatform[r.platform] = byPlatform[r.platform] || []).push(r);
  });

  const result = {};
  for (const [plat, entries] of Object.entries(byPlatform)) {
    result[plat] = entries.map((entry, i) => {
      const prev = entries[i - 1];
      const change = prev ? entry.followers_count - prev.followers_count : null;
      const changePct = prev && prev.followers_count ? Math.round((change / prev.followers_count) * 1000) / 10 : null;
      return {
        id: entry.id,
        entryDate: entry.entry_date,
        followersCount: entry.followers_count,
        change,
        changePct,
      };
    });
  }
  return result;
}

/** Every followers_history row, grouped by platform, oldest-to-newest — the unfiltered base that followersKpis() reduces over. Unlike followersGrowth() above (which is deliberately range-filterable for the growth chart), this needs the true full history to find the actual latest/previous entry regardless of whatever date range the Dashboard filter bar has selected. */
function allFollowersByPlatform() {
  const rows = db.prepare('SELECT * FROM followers_history ORDER BY platform ASC, entry_date ASC').all();
  const byPlatform = {};
  rows.forEach((r) => {
    (byPlatform[r.platform] = byPlatform[r.platform] || []).push(r);
  });
  return byPlatform;
}

/**
 * Powers the Dashboard's three Followers KPI cards. All three are computed
 * per-platform first, then summed across whichever platform(s) are in scope
 * (either the single selected platform, or every platform when the filter is
 * "all") — so a platform with no prior entry contributes 0 to the aggregate
 * change instead of corrupting the total, and the whole thing degrades to a
 * single platform's own numbers when one is selected.
 *
 * - currentFollowers / followersChange / followersChangePct: always based on
 *   each platform's true latest and second-to-last entry, ignoring
 *   dateFrom/dateTo entirely — a point-in-time "where do we stand right now"
 *   figure shouldn't go stale just because the dashboard's date range
 *   happens to be scrolled to an earlier period.
 * - newFollowers: the one figure that IS range-scoped, per the spec ("New
 *   Followers... within the selected date range"). For each platform, the
 *   baseline is the last known entry strictly before dateFrom (so a single
 *   entry landing inside a short range, e.g. "last 7 days" catching one
 *   weekly entry, still produces a real delta against the prior week) —
 *   falling back to the first in-range entry only when there's no earlier
 *   entry at all. Returns null when nothing anywhere is computable, which
 *   the frontend renders as "No follower update".
 */
function followersKpis({ platform, dateFrom, dateTo } = {}) {
  const byPlatform = allFollowersByPlatform();
  const platformIds = platform && platform !== 'all' ? [platform] : Object.keys(byPlatform);

  let currentFollowers = 0;
  let previousFollowers = 0;
  let hasAnyData = false;
  let newFollowers = 0;
  let hasAnyGain = false;

  for (const plat of platformIds) {
    const entries = byPlatform[plat];
    if (!entries || !entries.length) continue;
    hasAnyData = true;
    const last = entries[entries.length - 1];
    const prev = entries.length > 1 ? entries[entries.length - 2] : null;
    currentFollowers += last.followers_count;
    previousFollowers += prev ? prev.followers_count : last.followers_count;

    if (dateFrom && dateTo) {
      const inRange = entries.filter((e) => e.entry_date >= dateFrom && e.entry_date <= dateTo);
      if (inRange.length) {
        const rangeLast = inRange[inRange.length - 1];
        const before = entries.filter((e) => e.entry_date < dateFrom);
        const baseline = before.length ? before[before.length - 1] : inRange[0];
        if (baseline !== rangeLast) {
          newFollowers += rangeLast.followers_count - baseline.followers_count;
          hasAnyGain = true;
        }
      }
    }
  }

  const change = hasAnyData ? currentFollowers - previousFollowers : null;
  const changePct = hasAnyData && previousFollowers ? Math.round((change / previousFollowers) * 1000) / 10 : null;

  return {
    currentFollowers: hasAnyData ? currentFollowers : null,
    followersChange: hasAnyData ? change : null,
    followersChangePct: changePct,
    newFollowers: hasAnyGain ? newFollowers : null,
  };
}

/**
 * Cross-platform comparison report — the Comparisons page's headline report.
 * Deliberately independent of the shared filter bar (platform/campaign/content
 * type): it always covers every platform that has *any* data (uploaded posts
 * and/or manually-entered follower history) and needs no date range to
 * produce a result — an explicit dateFrom/dateTo only narrows it and unlocks
 * the vs-previous-period growth figures.
 */
function platformComparisonReport({ dateFrom, dateTo } = {}) {
  const explicitRange = Boolean(dateFrom && dateTo);
  const dateRow = db.prepare('SELECT MIN(publish_date) AS min, MAX(publish_date) AS max FROM posts').get();
  const effectiveFrom = dateFrom || dateRow.min || null;
  const effectiveTo = dateTo || dateRow.max || null;
  const hasAnyPostData = Boolean(effectiveFrom && effectiveTo);

  const currentByPlatform = Object.fromEntries(
    platformBreakdown(hasAnyPostData ? { dateFrom: effectiveFrom, dateTo: effectiveTo } : {}).map((r) => [r.platform, r])
  );

  let previousByPlatform = {};
  if (explicitRange) {
    const prevRange = previousEqualRange(effectiveFrom, effectiveTo);
    previousByPlatform = Object.fromEntries(
      platformBreakdown({ dateFrom: prevRange.from, dateTo: prevRange.to }).map((r) => [r.platform, r])
    );
  }

  const followerData = followersGrowth(explicitRange ? { dateFrom: effectiveFrom, dateTo: effectiveTo } : {});
  const activeIds = PLATFORM_IDS.filter((id) => currentByPlatform[id] || followerData[id]);

  const platforms = activeIds.map((id) => {
    const meta = PLATFORMS.find((p) => p.id === id);
    const row = currentByPlatform[id] || {};
    const prev = previousByPlatform[id] || {};
    const totals = {};
    const growth = {};
    for (const key of CANONICAL_METRIC_KEYS) {
      totals[key] = row[key] || 0;
      growth[key] = explicitRange ? pctChange(row[key] || 0, prev[key] || 0) : null;
    }
    const entries = followerData[id] || [];
    const first = entries[0] || null;
    const last = entries.length > 1 ? entries[entries.length - 1] : null;
    const followerChange = first && last ? last.followersCount - first.followersCount : null;
    const followerChangePct = followerChange !== null && first.followersCount
      ? Math.round((followerChange / first.followersCount) * 1000) / 10
      : null;

    return {
      platform: id,
      label: meta.label,
      color: meta.color,
      postCount: row.post_count || 0,
      hasPostData: Boolean(row.post_count),
      totals,
      growth,
      followers: {
        latest: entries.length ? entries[entries.length - 1].followersCount : null,
        change: followerChange,
        changePct: followerChangePct,
      },
    };
  });

  const RANK_METRICS = ['reach', 'engagement', 'impressions'];
  RANK_METRICS.forEach((key) => {
    const ranked = platforms.filter((p) => p.hasPostData).sort((a, b) => b.totals[key] - a.totals[key]);
    ranked.forEach((p, i) => {
      p.ranks = p.ranks || {};
      p.ranks[key] = i + 1;
    });
  });
  platforms.forEach((p) => {
    if (!p.hasPostData) { p.ranks = { ...(p.ranks || {}), composite: null }; return; }
    const vals = RANK_METRICS.map((k) => p.ranks[k]);
    p.ranks.composite = vals.reduce((a, b) => a + b, 0) / vals.length;
  });

  const rankable = platforms.filter((p) => p.hasPostData).sort((a, b) => a.ranks.composite - b.ranks.composite);
  rankable.forEach((p, i) => { p.overallRank = i + 1; });
  platforms.forEach((p) => { if (!p.hasPostData) p.overallRank = null; });
  platforms.sort((a, b) => (a.overallRank || 999) - (b.overallRank || 999));

  const bestPlatform = rankable[0] || null;
  const worstPlatform = rankable.length > 1 ? rankable[rankable.length - 1] : null;

  return {
    range: { from: effectiveFrom, to: effectiveTo, isExplicit: explicitRange },
    platforms,
    bestPlatform: bestPlatform ? bestPlatform.platform : null,
    worstPlatform: worstPlatform ? worstPlatform.platform : null,
    insights: buildComparisonInsights(platforms, bestPlatform, worstPlatform, explicitRange),
  };
}

function formatInt(n) {
  return Math.round(n).toLocaleString('en-US');
}

/** Plain-language takeaways generated from the numbers already computed above — a rule-based summary, not an external AI call. */
function buildComparisonInsights(platforms, bestPlatform, worstPlatform, explicitRange) {
  if (!platforms.length) {
    return ['No platform data yet — upload posts or add Followers Data Record entries to see insights here.'];
  }
  const lines = [];

  if (bestPlatform) {
    lines.push(`${bestPlatform.label} is the top-performing platform overall, ranking best on average across reach, engagement, and impressions.`);
  }
  if (worstPlatform) {
    lines.push(`${worstPlatform.label} is currently the lowest-ranked platform among those with post data — worth a closer look at its content strategy.`);
  }

  if (explicitRange) {
    const byEngagementGrowth = platforms
      .filter((p) => p.growth.engagement !== null)
      .sort((a, b) => b.growth.engagement - a.growth.engagement);
    if (byEngagementGrowth.length && byEngagementGrowth[0].growth.engagement > 0) {
      const top = byEngagementGrowth[0];
      lines.push(`${top.label} had the strongest engagement growth, up ${top.growth.engagement}% versus the previous equivalent period.`);
    }
    const declining = byEngagementGrowth.filter((p) => p.growth.engagement !== null && p.growth.engagement < 0);
    if (declining.length) {
      const worstGrowth = declining[declining.length - 1];
      lines.push(`${worstGrowth.label} saw engagement drop ${Math.abs(worstGrowth.growth.engagement)}% compared to the previous period.`);
    }
  }

  const byFollowerGrowth = platforms
    .filter((p) => p.followers.change !== null)
    .sort((a, b) => b.followers.change - a.followers.change);
  if (byFollowerGrowth.length && byFollowerGrowth[0].followers.change > 0) {
    const top = byFollowerGrowth[0];
    const pctPart = top.followers.changePct !== null ? `, ${top.followers.changePct > 0 ? '+' : ''}${top.followers.changePct}%` : '';
    lines.push(`${top.label} gained the most followers over this period (+${formatInt(top.followers.change)}${pctPart}).`);
  }

  const totalReach = platforms.reduce((sum, p) => sum + (p.totals.reach || 0), 0);
  const totalEngagement = platforms.reduce((sum, p) => sum + (p.totals.engagement || 0), 0);
  lines.push(`Across all platforms combined: ${formatInt(totalReach)} total reach and ${formatInt(totalEngagement)} total engagement.`);

  return lines;
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

recordsRouter.get('/export', async (req, res) => {
  const { dateFrom, dateTo, platform, campaignType, contentType, search, format } = req.query;
  const { columns, rows } = exportRecordsRows({ dateFrom, dateTo, platform, campaignType, contentType, search });
  const filename = `data-records-${new Date().toISOString().slice(0, 10)}`;
  try {
    if (format === 'xlsx') {
      await sendXLSX(res, `${filename}.xlsx`, rows, columns, 'Data Records');
    } else {
      sendCSV(res, `${filename}.csv`, rows, columns);
    }
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Duplicate cleanup — declared before the '/:id' catch-all GET below so
// '/duplicates' isn't swallowed by it.
recordsRouter.get('/duplicates', (req, res) => {
  res.json(previewDuplicateCleanup());
});

recordsRouter.post('/duplicates/resolve', express.json(), (req, res) => {
  try {
    res.json(runDuplicateCleanup());
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Wipe every uploaded post record + its upload history, for starting the data
// over from a single clean import. Requires an explicit { confirm: 'DELETE' }
// body. Followers Data Record, the access code, and sessions are untouched.
recordsRouter.post('/wipe', express.json(), (req, res) => {
  if (!req.body || req.body.confirm !== 'DELETE') {
    return res.status(400).json({ error: 'Confirmation phrase required.' });
  }
  try {
    res.json(wipeUploadedData());
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
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

analyticsRouter.get('/platform-report', (req, res) => {
  const { dateFrom, dateTo } = req.query;
  res.json(platformComparisonReport({ dateFrom, dateTo }));
});

app.use('/api/analytics', requireAuth, analyticsRouter);

// ---- /api/followers — manual weekly follower-count entry, independent of uploads ----
const followersRouter = express.Router();

followersRouter.get('/', (req, res) => {
  res.json(listFollowers(req.query));
});

followersRouter.get('/growth', (req, res) => {
  res.json(followersGrowth(req.query));
});

followersRouter.get('/kpis', (req, res) => {
  res.json(followersKpis(req.query));
});

followersRouter.post('/', express.json(), (req, res) => {
  try {
    const body = req.body || {};
    const entry = createOrUpdateFollowers({
      platform: body.platform,
      entryDate: body.entryDate,
      followersCount: body.followersCount,
    });
    res.json(entry);
  } catch (err) {
    res.status(422).json({ error: err.message });
  }
});

followersRouter.put('/:id', express.json(), (req, res) => {
  try {
    const body = req.body || {};
    const entry = updateFollowersEntry(Number(req.params.id), {
      platform: body.platform,
      entryDate: body.entryDate,
      followersCount: body.followersCount,
    });
    res.json(entry);
  } catch (err) {
    res.status(422).json({ error: err.message });
  }
});

followersRouter.delete('/:id', (req, res) => {
  try {
    deleteFollowersEntry(Number(req.params.id));
    res.json({ ok: true });
  } catch (err) {
    res.status(404).json({ error: err.message });
  }
});

app.use('/api/followers', requireAuth, followersRouter);

// ---- /api/export — generic CSV/XLSX export for tables that load their full dataset
// client-side (Followers Data, Upload History) rather than through a server-paginated
// endpoint. The client already has the exact rows on screen (post-search, post-sort);
// this just turns them into a downloadable file — CSV needs no server round trip at
// all (handled entirely client-side), XLSX comes through here so exceljs (already a
// dependency) can generate a real .xlsx without adding any client-side library. ----
const exportRouter = express.Router();

exportRouter.post('/', express.json({ limit: '5mb' }), async (req, res) => {
  const { rows, columns, format, filename, sheetName } = req.body || {};
  if (!Array.isArray(rows) || !Array.isArray(columns)) {
    return res.status(400).json({ error: 'rows and columns arrays are required.' });
  }
  const safeName = String(filename || 'export').replace(/[^a-zA-Z0-9._-]/g, '_');
  try {
    if (format === 'xlsx') {
      await sendXLSX(res, `${safeName}.xlsx`, rows, columns, sheetName);
    } else {
      sendCSV(res, `${safeName}.csv`, rows, columns);
    }
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.use('/api/export', requireAuth, exportRouter);

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
const INDEX_HTML_BASE64 = 'PCFkb2N0eXBlIGh0bWw+CjxodG1sIGxhbmc9ImVuIj4KPGhlYWQ+CjxtZXRhIGNoYXJzZXQ9IlVURi04IiAvPgo8bWV0YSBuYW1lPSJ2aWV3cG9ydCIgY29udGVudD0id2lkdGg9ZGV2aWNlLXdpZHRoLCBpbml0aWFsLXNjYWxlPTEuMCIgLz4KPHRpdGxlPkxSUyBBbmFseXRpY3MgRGFzaGJvYXJkPC90aXRsZT4KPGxpbmsgcmVsPSJpY29uIiB0eXBlPSJpbWFnZS9wbmciIGlkPSJmYXZpY29uTGluayIgLz4KPGxpbmsgcmVsPSJwcmVjb25uZWN0IiBocmVmPSJodHRwczovL2ZvbnRzLmdvb2dsZWFwaXMuY29tIiAvPgo8bGluayByZWw9InByZWNvbm5lY3QiIGhyZWY9Imh0dHBzOi8vZm9udHMuZ3N0YXRpYy5jb20iIGNyb3Nzb3JpZ2luIC8+CjxsaW5rIGhyZWY9Imh0dHBzOi8vZm9udHMuZ29vZ2xlYXBpcy5jb20vY3NzMj9mYW1pbHk9SW50ZXI6d2dodEA0MDA7NTAwOzYwMDs3MDA7ODAwJmRpc3BsYXk9c3dhcCIgcmVsPSJzdHlsZXNoZWV0IiAvPgo8c2NyaXB0IHNyYz0iaHR0cHM6Ly9jZG4uanNkZWxpdnIubmV0L25wbS9jaGFydC5qc0A0LjQuNC9kaXN0L2NoYXJ0LnVtZC5taW4uanMiPjwvc2NyaXB0Pgo8c2NyaXB0IHNyYz0iaHR0cHM6Ly9jZG4uanNkZWxpdnIubmV0L25wbS9jaGFydGpzLXBsdWdpbi1kYXRhbGFiZWxzQDIvZGlzdC9jaGFydGpzLXBsdWdpbi1kYXRhbGFiZWxzLm1pbi5qcyI+PC9zY3JpcHQ+CjxzY3JpcHQgc3JjPSJodHRwczovL2Nkbi5qc2RlbGl2ci5uZXQvbnBtL2x1Y2lkZUAwLjQ2Mi4wL2Rpc3QvdW1kL2x1Y2lkZS5taW4uanMiPjwvc2NyaXB0Pgo8c3R5bGU+Ci8qIC0tLS0tLS0tLS0gRGVzaWduIHRva2VuczogZGFyayBuYXZ5ICsgZ29sZCBicmFuZGVkIHRoZW1lIChzaW5nbGUsIHBlcm1hbmVudCDigJQgbm8gbGlnaHQgdmFyaWFudCkgLS0tLS0tLS0tLSAqLwo6cm9vdCB7CiAgY29sb3Itc2NoZW1lOiBkYXJrOwogIC0tZm9udC1zYW5zOiAnSW50ZXInLCAtYXBwbGUtc3lzdGVtLCBCbGlua01hY1N5c3RlbUZvbnQsICdTRiBQcm8gRGlzcGxheScsICdTZWdvZSBVSScsIFJvYm90bywgc2Fucy1zZXJpZjsKCiAgLS1wYWdlLXBsYW5lOiBsaW5lYXItZ3JhZGllbnQoMTgwZGVnLCAjMGEwZjFjIDAlLCAjMGQxNDI0IDEwMCUpOwogIC0tcGFnZS1wbGFuZS1zb2xpZDogIzBhMGYxYzsKICAtLXNpZGViYXItYmc6ICMwYjEyMjA7CiAgLS1zdXJmYWNlLTE6IHJnYmEoMjMsIDMxLCA1MSwgMC42Mik7IC8qIGdsYXNzOiBjYXJkcywgS1BJIHRpbGVzLCBmaWx0ZXIgYmFyICovCiAgLS1zdXJmYWNlLTI6IHJnYmEoMjU1LCAyNTUsIDI1NSwgMC4wNik7IC8qIGdsYXNzOiBpbnB1dHMsIG5lc3RlZCByb3dzLCBwaWxscyAqLwogIC0tc3VyZmFjZS1zb2xpZDogIzEzMWIyZTsKICAtLWdsYXNzLWJsdXI6IGJsdXIoMjBweCk7CiAgLS1ib3JkZXI6IHJnYmEoMjU1LCAyNTUsIDI1NSwgMC4wOSk7CiAgLS10ZXh0LXByaW1hcnk6ICNmNGY1Zjc7CiAgLS10ZXh0LXNlY29uZGFyeTogI2I4YmJjNDsKICAtLXRleHQtbXV0ZWQ6ICM4Mjg2OGY7CiAgLS1ncmlkbGluZTogcmdiYSgyNTUsIDI1NSwgMjU1LCAwLjA4KTsKICAtLWJhc2VsaW5lOiByZ2JhKDI1NSwgMjU1LCAyNTUsIDAuMik7CiAgLS1zdWNjZXNzLXRleHQ6ICMzNGM3NmY7CgogIC0tc3RhdHVzLWdvb2Q6ICMyZmI4NjI7CiAgLS1zdGF0dXMtd2FybmluZzogI2YwYTEzYTsKICAtLXN0YXR1cy1zZXJpb3VzOiAjZWM4MzVhOwogIC0tc3RhdHVzLWNyaXRpY2FsOiAjZTA2MDVmOwoKICAtLWFjY2VudC1nb2xkOiAjZjJiMzBlOyAvKiBMUlMgYnJhbmQgZ29sZCDigJQgYWN0aXZlIG5hdiBpdGVtLCBwcmltYXJ5IGludGVyYWN0aXZlIGFjY2VudCAqLwoKICAtLXNlcmllcy0xOiAjMzk4N2U1OyAvKiBmYWNlYm9vayAqLwogIC0tc2VyaWVzLTI6ICMwMDgzMDA7IC8qIGluc3RhZ3JhbSAqLwogIC0tc2VyaWVzLTM6ICNkNTUxODE7IC8qIHRpa3RvayAqLwogIC0tc2VyaWVzLTQ6ICNjOTg1MDA7IC8qIGxpbmtlZGluICovCiAgLS1zZXJpZXMtNTogIzE5OWU3MDsgLyogdGhyZWFkcyAqLwogIC0tc2VyaWVzLTY6ICNkOTU5MjY7IC8qIHlvdXR1YmUgKi8KICAtLXNlcmllcy03OiAjOTA4NWU5OyAvKiByZXNlcnZlZCAqLwogIC0tc2VyaWVzLTg6ICNlNjY3Njc7IC8qIHJlc2VydmVkICovCgogIC0tcmFkaXVzLXNtOiAxMHB4OwogIC0tcmFkaXVzLW1kOiAxNHB4OwogIC0tcmFkaXVzLWxnOiAxOHB4OwoKICAtLXNoYWRvdy1jYXJkOiAwIDFweCAycHggcmdiYSgwLDAsMCwwLjIpLCAwIDhweCAyNHB4IC0xMHB4IHJnYmEoMCwwLDAsMC41KTsKICAtLXNoYWRvdy1ob3ZlcjogMCA2cHggMTJweCAtMnB4IHJnYmEoMCwwLDAsMC4zKSwgMCAxOHB4IDQwcHggLTE0cHggcmdiYSgwLDAsMCwwLjYpOwogIC0tc2hhZG93LW1vZGFsOiAwIDI0cHggNjRweCAtMTJweCByZ2JhKDAsMCwwLDAuNyk7CiAgLS1lYXNlOiBjdWJpYy1iZXppZXIoMC40LCAwLCAwLjIsIDEpOwp9CgoqIHsgYm94LXNpemluZzogYm9yZGVyLWJveDsgfQpodG1sLCBib2R5IHsgaGVpZ2h0OiAxMDAlOyB9CmJvZHkgewogIG1hcmdpbjogMDsKICBmb250LWZhbWlseTogdmFyKC0tZm9udC1zYW5zKTsKICBiYWNrZ3JvdW5kOiB2YXIoLS1wYWdlLXBsYW5lKTsKICBiYWNrZ3JvdW5kLWF0dGFjaG1lbnQ6IGZpeGVkOwogIGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOwogIC13ZWJraXQtZm9udC1zbW9vdGhpbmc6IGFudGlhbGlhc2VkOwogIC1tb3otb3N4LWZvbnQtc21vb3RoaW5nOiBncmF5c2NhbGU7Cn0KYnV0dG9uLCBzZWxlY3QsIGlucHV0LCB0ZXh0YXJlYSB7IGZvbnQtZmFtaWx5OiBpbmhlcml0OyB9CmgxLCBoMiwgaDMsIGg0IHsgZm9udC13ZWlnaHQ6IDcwMDsgbGV0dGVyLXNwYWNpbmc6IC0wLjAxZW07IH0KCjo6c2VsZWN0aW9uIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSAzMCUsIHRyYW5zcGFyZW50KTsgfQoKLyogQ3VzdG9tIHNjcm9sbGJhciDigJQgdGhpbiwgdW5vYnRydXNpdmUsIGZpdHMgdGhlIGdsYXNzIGFlc3RoZXRpYyAqLwo6Oi13ZWJraXQtc2Nyb2xsYmFyIHsgd2lkdGg6IDEwcHg7IGhlaWdodDogMTBweDsgfQo6Oi13ZWJraXQtc2Nyb2xsYmFyLXRyYWNrIHsgYmFja2dyb3VuZDogdHJhbnNwYXJlbnQ7IH0KOjotd2Via2l0LXNjcm9sbGJhci10aHVtYiB7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS10ZXh0LW11dGVkKSA0MCUsIHRyYW5zcGFyZW50KTsgYm9yZGVyLXJhZGl1czogMjBweDsgYm9yZGVyOiAycHggc29saWQgdHJhbnNwYXJlbnQ7IGJhY2tncm91bmQtY2xpcDogcGFkZGluZy1ib3g7IH0KOjotd2Via2l0LXNjcm9sbGJhci10aHVtYjpob3ZlciB7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS10ZXh0LW11dGVkKSA2MCUsIHRyYW5zcGFyZW50KTsgYmFja2dyb3VuZC1jbGlwOiBwYWRkaW5nLWJveDsgfQoKLmFwcC1zaGVsbCB7IGhlaWdodDogMTAwdmg7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiByb3c7IG92ZXJmbG93OiBoaWRkZW47IH0KCi8qIC0tLS0tLS0tLS0gU2lkZWJhciAtLS0tLS0tLS0tICovCi5zaWRlYmFyIHsKICBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBnYXA6IDRweDsKICB3aWR0aDogMjQwcHg7IGZsZXg6IDAgMCBhdXRvOwogIHBhZGRpbmc6IDIwcHggMTRweDsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zaWRlYmFyLWJnKTsKICBib3JkZXItcmlnaHQ6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOwogIHotaW5kZXg6IDIwOwogIG92ZXJmbG93LXk6IGF1dG87Cn0KLnNpZGViYXItYnJhbmQgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDhweDsgd2hpdGUtc3BhY2U6IG5vd3JhcDsgcGFkZGluZzogNHB4IDEwcHggMjBweDsgfQouYnJhbmQtbG9nbyB7IGhlaWdodDogMjhweDsgd2lkdGg6IGF1dG87IGRpc3BsYXk6IGJsb2NrOyBmbGV4LXNocmluazogMDsgb2JqZWN0LWZpdDogY29udGFpbjsgfQouYnJhbmQtdGl0bGUgeyBmb250LXdlaWdodDogNjAwOyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgbGV0dGVyLXNwYWNpbmc6IC0wLjAxZW07IH0KCi50YWJzIHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiAycHg7IGZsZXg6IDE7IHBvc2l0aW9uOiByZWxhdGl2ZTsgfQoudGFiLWJ0biB7CiAgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiAxMHB4OwogIGJvcmRlcjogbm9uZTsgYmFja2dyb3VuZDogdHJhbnNwYXJlbnQ7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7CiAgcGFkZGluZzogMTBweCAxMnB4OyBib3JkZXItcmFkaXVzOiAxMHB4OyBjdXJzb3I6IHBvaW50ZXI7IGZvbnQtc2l6ZTogMTRweDsgZm9udC13ZWlnaHQ6IDUwMDsKICB3aGl0ZS1zcGFjZTogbm93cmFwOyBwb3NpdGlvbjogcmVsYXRpdmU7IHRleHQtYWxpZ246IGxlZnQ7IHdpZHRoOiAxMDAlOwogIHRyYW5zaXRpb246IGNvbG9yIDE4MG1zIHZhcigtLWVhc2UpLCBiYWNrZ3JvdW5kIDE4MG1zIHZhcigtLWVhc2UpOwp9Ci50YWItYnRuIHN2ZyB7IGZsZXgtc2hyaW5rOiAwOyBvcGFjaXR5OiAwLjg7IH0KLnRhYi1idG4uaXMtYWN0aXZlIHN2ZyB7IG9wYWNpdHk6IDE7IGNvbG9yOiB2YXIoLS1zZXJpZXMtMSk7IH0KLnRhYi1idG46aG92ZXIgeyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDglLCB0cmFuc3BhcmVudCk7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyB9Ci50YWItYnRuLmlzLWFjdGl2ZSB7CiAgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IGZvbnQtd2VpZ2h0OiA2MDA7CiAgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSAxNiUsIHRyYW5zcGFyZW50KTsKICBhbmltYXRpb246IHRhYkluZGljYXRvckluIDIyMG1zIHZhcigtLWVhc2UpOwp9CkBrZXlmcmFtZXMgdGFiSW5kaWNhdG9ySW4geyBmcm9tIHsgb3BhY2l0eTogMDsgdHJhbnNmb3JtOiB0cmFuc2xhdGVYKC00cHgpOyB9IHRvIHsgb3BhY2l0eTogMTsgdHJhbnNmb3JtOiB0cmFuc2xhdGVYKDApOyB9IH0KLnNpZGViYXItZm9vdGVyIHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiAxMHB4OyBwYWRkaW5nLXRvcDogMTRweDsgbWFyZ2luLXRvcDogMTRweDsgYm9yZGVyLXRvcDogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IH0KCi5zaWRlYmFyLXVzZXIgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDEwcHg7IGZvbnQtc2l6ZTogMTNweDsgfQoKLyogLS0tLS0tLS0tLSBBdXRoIHNjcmVlbiAtLS0tLS0tLS0tICovCi5hdXRoLXNjcmVlbiB7CiAgbWluLWhlaWdodDogMTAwdmg7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogY2VudGVyOwogIGJhY2tncm91bmQ6IHZhcigtLXBhZ2UtcGxhbmUpOyBwYWRkaW5nOiAyMHB4Owp9Ci5hdXRoLWNhcmQgewogIHdpZHRoOiAxMDAlOyBtYXgtd2lkdGg6IDQwMHB4OyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTEpOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOwogIGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtbGcpOyBwYWRkaW5nOiAzMnB4OyBib3gtc2hhZG93OiB2YXIoLS1zaGFkb3ctbW9kYWwpOwogIGFuaW1hdGlvbjogbW9kYWxQYW5lbEluIDI2MG1zIHZhcigtLWVhc2UpOwp9Ci5hdXRoLWJyYW5kIHsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA4cHg7IG1hcmdpbi1ib3R0b206IDIycHg7IH0KLmF1dGgtYnJhbmQgLmJyYW5kLXRpdGxlIHsgZm9udC13ZWlnaHQ6IDcwMDsgZm9udC1zaXplOiAxN3B4OyB9Ci5hdXRoLWJyYW5kIC5icmFuZC1sb2dvIHsgaGVpZ2h0OiAzNnB4OyB9Ci5hdXRoLWZvcm0geyBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBnYXA6IDE0cHg7IG1hcmdpbi10b3A6IDE2cHg7IH0KLmF1dGgtZm9ybSAuZm9ybS1maWVsZCBpbnB1dCB7IHdpZHRoOiAxMDAlOyB9Ci5hdXRoLWVycm9yIHsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IGZvbnQtc2l6ZTogMTJweDsgbWluLWhlaWdodDogMTZweDsgfQoKLyogLS0tLS0tLS0tLSBGaWx0ZXIgYmFyIC0tLS0tLS0tLS0gKi8KLmZpbHRlci1iYXIgewogIGRpc3BsYXk6IGZsZXg7IGZsZXgtd3JhcDogd3JhcDsgYWxpZ24taXRlbXM6IGVuZDsgZ2FwOiAxNnB4OwogIHBhZGRpbmc6IDE0cHggMjBweDsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgYm9yZGVyLWJvdHRvbTogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7CiAgcG9zaXRpb246IHN0aWNreTsgdG9wOiAwOyB6LWluZGV4OiAxOTsKfQouZmlsdGVyLWZpZWxkIHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiA1cHg7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgfQouZmlsdGVyLWZpZWxkIGxhYmVsIHsgZm9udC13ZWlnaHQ6IDYwMDsgfQouZmlsdGVyLXByZXNldHMgeyBmbGV4LWRpcmVjdGlvbjogcm93OyBnYXA6IDZweDsgfQouZmlsdGVyLXByZXNldHMgYnV0dG9uIHsKICBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTIpOyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOwogIGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBib3JkZXItcmFkaXVzOiAyMHB4OyBwYWRkaW5nOiA3cHggMTNweDsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNTAwOyBjdXJzb3I6IHBvaW50ZXI7CiAgdHJhbnNpdGlvbjogY29sb3IgMTgwbXMgdmFyKC0tZWFzZSksIGJhY2tncm91bmQgMTgwbXMgdmFyKC0tZWFzZSksIHRyYW5zZm9ybSAxNTBtcyB2YXIoLS1lYXNlKTsKfQouZmlsdGVyLXByZXNldHMgYnV0dG9uOmhvdmVyIHsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgtMXB4KTsgfQouZmlsdGVyLXByZXNldHMgYnV0dG9uOmFjdGl2ZSB7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgwKSBzY2FsZSgwLjk2KTsgfQouZmlsdGVyLXByZXNldHMgYnV0dG9uLmlzLWFjdGl2ZSB7IGJhY2tncm91bmQ6IHZhcigtLXNlcmllcy0xKTsgY29sb3I6ICNmZmY7IGJvcmRlci1jb2xvcjogdHJhbnNwYXJlbnQ7IGJveC1zaGFkb3c6IDAgNHB4IDE0cHggLTVweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDYwJSwgdHJhbnNwYXJlbnQpOyB9CgovKiAtLS0tLS0tLS0tIE1haW4gY29sdW1uIChzaXRzIGJlc2lkZSB0aGUgc2lkZWJhcjsgc2Nyb2xscyBpbmRlcGVuZGVudGx5IHNvIHRoZSBzaWRlYmFyIHN0YXlzIGZ1bGx5IHZpc2libGUpIC0tLS0tLS0tLS0gKi8KLm1haW4tY29sIHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZmxleDogMSAxIGF1dG87IG1pbi13aWR0aDogMDsgaGVpZ2h0OiAxMDAlOyBvdmVyZmxvdy15OiBhdXRvOyB9CgovKiAtLS0tLS0tLS0tIFZpZXcgYXJlYSAtLS0tLS0tLS0tICovCi52aWV3LWFyZWEgeyBmbGV4OiAxOyBwYWRkaW5nOiAyNHB4OyBtYXgtd2lkdGg6IDE4MDBweDsgd2lkdGg6IDEwMCU7IG1hcmdpbjogMCBhdXRvOyB9Ci52aWV3IHsgZGlzcGxheTogbm9uZTsgfQoudmlldy5pcy1hY3RpdmUgeyBkaXNwbGF5OiBibG9jazsgYW5pbWF0aW9uOiB2aWV3RmFkZUluIDI2MG1zIHZhcigtLWVhc2UpOyB9CkBrZXlmcmFtZXMgdmlld0ZhZGVJbiB7CiAgZnJvbSB7IG9wYWNpdHk6IDA7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSg2cHgpOyB9CiAgdG8geyBvcGFjaXR5OiAxOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMCk7IH0KfQoKLnNlY3Rpb24tdGl0bGUgeyBmb250LXNpemU6IDE2cHg7IGZvbnQtd2VpZ2h0OiA3MDA7IG1hcmdpbjogMzJweCAwIDE0cHg7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyBsZXR0ZXItc3BhY2luZzogLTAuMDFlbTsgfQouc2VjdGlvbi10aXRsZTpmaXJzdC1jaGlsZCB7IG1hcmdpbi10b3A6IDA7IH0KCi8qIC0tLS0tLS0tLS0gSW5wdXRzIOKAlCBvbmUgc2hhcmVkIGdsYXNzIHRyZWF0bWVudCBmb3IgZXZlcnkgdGV4dCBpbnB1dCwgc2VsZWN0LCBhbmQgZGF0ZSBwaWNrZXIgLS0tLS0tLS0tLSAqLwouZmlsdGVyLWZpZWxkIHNlbGVjdCwgLmZpbHRlci1maWVsZCBpbnB1dFt0eXBlPSJkYXRlIl0sCi5mb3JtLWZpZWxkIGlucHV0LCAuZm9ybS1maWVsZCBzZWxlY3QsIC5mb3JtLWZpZWxkIHRleHRhcmVhLAouZGFzaGJvYXJkLWNvbnRyb2xzIHNlbGVjdCwgLnJlY29yZHMtc2VhcmNoIGlucHV0LAouZmllbGQtaW5saW5lIHNlbGVjdCwgLmZpZWxkLWlubGluZSBpbnB1dCwKLmNvbmZsaWN0LXJvdyBzZWxlY3QsIC5jYXJkLWhlYWRlciBzZWxlY3QgewogIGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1zbSk7CiAgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0yKTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IGZvbnQtc2l6ZTogMTNweDsKICBwYWRkaW5nOiA4cHggMTJweDsgbWluLXdpZHRoOiAxNDBweDsKICB0cmFuc2l0aW9uOiBib3JkZXItY29sb3IgMTYwbXMgdmFyKC0tZWFzZSksIGJveC1zaGFkb3cgMTYwbXMgdmFyKC0tZWFzZSk7Cn0KLmZpbHRlci1maWVsZCBzZWxlY3Q6aG92ZXIsIC5maWx0ZXItZmllbGQgaW5wdXRbdHlwZT0iZGF0ZSJdOmhvdmVyLAouZm9ybS1maWVsZCBpbnB1dDpob3ZlciwgLmZvcm0tZmllbGQgc2VsZWN0OmhvdmVyLAouZGFzaGJvYXJkLWNvbnRyb2xzIHNlbGVjdDpob3ZlciwgLnJlY29yZHMtc2VhcmNoIGlucHV0OmhvdmVyLAouZmllbGQtaW5saW5lIHNlbGVjdDpob3ZlciwgLmZpZWxkLWlubGluZSBpbnB1dDpob3ZlciwKLmNvbmZsaWN0LXJvdyBzZWxlY3Q6aG92ZXIsIC5jYXJkLWhlYWRlciBzZWxlY3Q6aG92ZXIgewogIGJvcmRlci1jb2xvcjogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSAzNSUsIHZhcigtLWJvcmRlcikpOwp9Ci8qIEEgPHNlbGVjdD4ncyBvd24gYmFja2dyb3VuZCBpcyBhIHRyYW5zbHVjZW50IGdsYXNzIHRpbnQgbWVhbnQgdG8gYmxlbmQgd2l0aAogICB0aGUgcGFnZSBiZWhpbmQgaXQg4oCUIGJ1dCBpdHMgZHJvcGRvd24gcG9wdXAgcmVuZGVycyBvbiBhbiBpc29sYXRlZCBvcGFxdWUKICAgY2FudmFzLCBzbyB0aGF0IHNhbWUgdHJhbnNsdWNlbnQgdmFsdWUgc2hvd3MgdXAgdGhlcmUgYXMgcGxhaW4gd2hpdGUKICAgaW5zdGVhZCBvZiBkYXJrLiBFdmVyeSA8b3B0aW9uPiwgaW4gZXZlcnkgc2VsZWN0IGluIHRoZSBhcHAsIG5lZWRzIGFuCiAgIGV4cGxpY2l0IHNvbGlkIGRhcmsgYmFja2dyb3VuZC90ZXh0IGNvbG9yIHNvIHRoZSBwb3B1cCBtYXRjaGVzIHRoZSB0aGVtZS4gKi8Kb3B0aW9uIHsgYmFja2dyb3VuZC1jb2xvcjogdmFyKC0tc3VyZmFjZS1zb2xpZCk7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyB9CgouZmlsdGVyLWZpZWxkIHNlbGVjdDpmb2N1cywgLmZpbHRlci1maWVsZCBpbnB1dFt0eXBlPSJkYXRlIl06Zm9jdXMsCi5mb3JtLWZpZWxkIGlucHV0OmZvY3VzLCAuZm9ybS1maWVsZCBzZWxlY3Q6Zm9jdXMsIC5mb3JtLWZpZWxkIHRleHRhcmVhOmZvY3VzLAouZGFzaGJvYXJkLWNvbnRyb2xzIHNlbGVjdDpmb2N1cywgLnJlY29yZHMtc2VhcmNoIGlucHV0OmZvY3VzLAouZmllbGQtaW5saW5lIHNlbGVjdDpmb2N1cywgLmZpZWxkLWlubGluZSBpbnB1dDpmb2N1cywKLmNvbmZsaWN0LXJvdyBzZWxlY3Q6Zm9jdXMsIC5jYXJkLWhlYWRlciBzZWxlY3Q6Zm9jdXMsCi5hdXRoLWZvcm0gaW5wdXQ6Zm9jdXMgewogIG91dGxpbmU6IG5vbmU7IGJvcmRlci1jb2xvcjogdmFyKC0tc2VyaWVzLTEpOwogIGJveC1zaGFkb3c6IDAgMCAwIDNweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDE4JSwgdHJhbnNwYXJlbnQpOwp9CgovKiAtLS0tLS0tLS0tIFN0YXQgdGlsZXMgLS0tLS0tLS0tLSAqLwouc3RhdC1ncmlkIHsKICBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdChhdXRvLWZpdCwgbWlubWF4KDE4MHB4LCAxZnIpKTsgZ2FwOiAxNHB4Owp9Ci5zdGF0LXRpbGUgewogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1tZCk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHBhZGRpbmc6IDE2cHggMThweDsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWNhcmQpOwogIHRyYW5zaXRpb246IHRyYW5zZm9ybSAyMDBtcyB2YXIoLS1lYXNlKSwgYm94LXNoYWRvdyAyMDBtcyB2YXIoLS1lYXNlKTsKICBhbmltYXRpb246IGNhcmRJbiAzMjBtcyB2YXIoLS1lYXNlKSBiYWNrd2FyZHM7Cn0KLnN0YXQtdGlsZTpob3ZlciB7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgtM3B4KTsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWhvdmVyKTsgfQouc3RhdC1sYWJlbCB7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgZm9udC13ZWlnaHQ6IDYwMDsgfQouc3RhdC12YWx1ZSB7IGZvbnQtc2l6ZTogMjdweDsgZm9udC13ZWlnaHQ6IDcwMDsgbWFyZ2luLXRvcDogNXB4OyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgbGV0dGVyLXNwYWNpbmc6IC0wLjAyZW07IH0KLnN0YXQtZGVsdGEgeyBmb250LXNpemU6IDEycHg7IG1hcmdpbi10b3A6IDdweDsgZm9udC13ZWlnaHQ6IDYwMDsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA0cHg7IH0KLnN0YXQtZGVsdGEudXAgeyBjb2xvcjogdmFyKC0tc3VjY2Vzcy10ZXh0KTsgfQouc3RhdC1kZWx0YS5kb3duIHsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IH0KLnN0YXQtZGVsdGEuZmxhdCB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgfQouc3RhdC1kZWx0YS51cDo6YmVmb3JlIHsgY29udGVudDogJ+KGkSc7IH0KLnN0YXQtZGVsdGEuZG93bjo6YmVmb3JlIHsgY29udGVudDogJ+KGkyc7IH0KCi8qIC0tLS0tLS0tLS0gRGFzaGJvYXJkIEtQSSBncmlkIOKAlCBjb21wYWN0LCBzaW5nbGUtcm93LW9uLWRlc2t0b3AgbGF5b3V0LgogICBTY29wZWQgdG8gI2twaUdyaWQgc3BlY2lmaWNhbGx5IChub3QgdGhlIHNoYXJlZCAuc3RhdC1ncmlkLy5zdGF0LXRpbGUKICAgY2xhc3Nlcywgd2hpY2ggQ29tcGFyaXNvbnMgYW5kIHRoZSBVcGxvYWQgcHJldmlldyBzdW1tYXJ5IGFsc28gdXNlKSBzbwogICB0aGlzIGNvbXBhY3RpbmcgZG9lc24ndCBhZmZlY3QgdGhvc2Ugb3RoZXIgc3RhdC10aWxlIGdyaWRzLiAxMCBncmlkIHVuaXRzCiAgIHRvdGFsOiA3IHN0YW5kYXJkIEtQSSB0aWxlcyBhdCAxIHVuaXQgZWFjaCArIEJlc3QgUGVyZm9ybWluZyBQb3N0IGF0IDMKICAgdW5pdHMgKGEgM3gtd2lkZSBsYW5kc2NhcGUgY2FyZCksIGFsbCBzaGFyaW5nIG9uZSBmaXhlZCByb3cgaGVpZ2h0LiAtLS0tLS0tLS0tICovCiNrcGlHcmlkIHsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiByZXBlYXQoMTAsIG1pbm1heCgwLCAxZnIpKTsgZ2FwOiAxMnB4OyB9CiNrcGlHcmlkIC5zdGF0LXRpbGUgewogIGhlaWdodDogMTMycHg7IGJveC1zaXppbmc6IGJvcmRlci1ib3g7IG92ZXJmbG93OiBoaWRkZW47CiAgcGFkZGluZzogMTZweCAxOHB4OwogIGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGp1c3RpZnktY29udGVudDogY2VudGVyOwp9CiNrcGlHcmlkIC5zdGF0LWxhYmVsIHsgZm9udC1zaXplOiAxMnB4OyBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBhbGlnbi1pdGVtczogZmxleC1zdGFydDsgZ2FwOiA4cHg7IH0KI2twaUdyaWQgLnN0YXQtdmFsdWUgeyBmb250LXNpemU6IDMycHg7IG1hcmdpbi10b3A6IDhweDsgbGluZS1oZWlnaHQ6IDEuMTsgfQoja3BpR3JpZCAuc3RhdC1kZWx0YSB7IGZvbnQtc2l6ZTogMTNweDsgbWFyZ2luLXRvcDogOHB4OyB9Cgouc3RhdC1pY29uIHsKICB3aWR0aDogMjhweDsgaGVpZ2h0OiAyOHB4OyBmbGV4OiAwIDAgYXV0bzsgYm9yZGVyLXJhZGl1czogNTAlOwogIGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogY2VudGVyOwogIGNvbG9yOiAjZmZmOwp9Ci5zdGF0LWljb24udjEgeyBiYWNrZ3JvdW5kOiB2YXIoLS1zZXJpZXMtMSk7IH0KLnN0YXQtaWNvbi52MiB7IGJhY2tncm91bmQ6IHZhcigtLXNlcmllcy0yKTsgfQouc3RhdC1pY29uLnYzIHsgYmFja2dyb3VuZDogdmFyKC0tc2VyaWVzLTMpOyB9Ci5zdGF0LWljb24udjQgeyBiYWNrZ3JvdW5kOiB2YXIoLS1zZXJpZXMtNCk7IH0KLnN0YXQtaWNvbi52NSB7IGJhY2tncm91bmQ6IHZhcigtLXNlcmllcy01KTsgfQouc3RhdC1pY29uLnY2IHsgYmFja2dyb3VuZDogdmFyKC0tc2VyaWVzLTYpOyB9Ci5zdGF0LWljb24uZ29sZCB7IGJhY2tncm91bmQ6IHZhcigtLWFjY2VudC1nb2xkKTsgfQpAbWVkaWEgKG1heC13aWR0aDogOTAwcHgpIHsgI2twaUdyaWQgeyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdCg1LCBtaW5tYXgoMCwgMWZyKSk7IH0gfQpAbWVkaWEgKG1heC13aWR0aDogNjQwcHgpIHsgI2twaUdyaWQgeyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdCgyLCBtaW5tYXgoMCwgMWZyKSk7IH0gfQoKLmluc2lnaHRzLWxpc3QgeyBsaXN0LXN0eWxlOiBub25lOyBtYXJnaW46IDA7IHBhZGRpbmc6IDA7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGdhcDogMTBweDsgfQouaW5zaWdodHMtbGlzdCBsaSB7CiAgZm9udC1zaXplOiAxM3B4OyBsaW5lLWhlaWdodDogMS41OyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyBwYWRkaW5nLWxlZnQ6IDE4cHg7IHBvc2l0aW9uOiByZWxhdGl2ZTsKfQouaW5zaWdodHMtbGlzdCBsaTo6YmVmb3JlIHsKICBjb250ZW50OiAn4pymJzsgcG9zaXRpb246IGFic29sdXRlOyBsZWZ0OiAwOyBjb2xvcjogdmFyKC0tc2VyaWVzLTEpOyBmb250LXNpemU6IDExcHg7IHRvcDogMnB4Owp9CgpAa2V5ZnJhbWVzIGNhcmRJbiB7CiAgZnJvbSB7IG9wYWNpdHk6IDA7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgxMHB4KTsgfQogIHRvIHsgb3BhY2l0eTogMTsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDApOyB9Cn0KCi8qIC0tLS0tLS0tLS0gQ2FyZHMgLyBjaGFydHMgLS0tLS0tLS0tLSAqLwouY2FyZC1ncmlkIHsgZGlzcGxheTogZ3JpZDsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiAyZnIgMWZyOyBnYXA6IDE2cHg7IGFsaWduLWl0ZW1zOiBzdGFydDsgfQouY2FyZC1ncmlkLmV2ZW4geyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IDFmciAxZnI7IH0KQG1lZGlhIChtYXgtd2lkdGg6IDkwMHB4KSB7IC5jYXJkLWdyaWQsIC5jYXJkLWdyaWQuZXZlbiB7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogMWZyOyB9IH0KLmNhcmQgewogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1sZyk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHBhZGRpbmc6IDE4cHg7IGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdy1jYXJkKTsKICB0cmFuc2l0aW9uOiBib3gtc2hhZG93IDIyMG1zIHZhcigtLWVhc2UpLCB0cmFuc2Zvcm0gMjIwbXMgdmFyKC0tZWFzZSk7CiAgYW5pbWF0aW9uOiBjYXJkSW4gMzIwbXMgdmFyKC0tZWFzZSkgYmFja3dhcmRzOwp9Ci5jYXJkOmhvdmVyIHsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWhvdmVyKTsgfQouY2FyZC1oZWFkZXIgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6IHNwYWNlLWJldHdlZW47IGdhcDogOHB4OyBtYXJnaW4tYm90dG9tOiAxNHB4OyB9Ci5jYXJkLWhlYWRlciBoMyB7IGZvbnQtc2l6ZTogMTRweDsgbWFyZ2luOiAwOyBmb250LXdlaWdodDogNzAwOyBsZXR0ZXItc3BhY2luZzogLTAuMDA1ZW07IH0KLmNhcmQtaGVhZGVyIHNlbGVjdCB7IGZvbnQtc2l6ZTogMTJweDsgcGFkZGluZzogNnB4IDEwcHg7IG1pbi13aWR0aDogMDsgfQouY2hhcnQtd3JhcCB7IHBvc2l0aW9uOiByZWxhdGl2ZTsgaGVpZ2h0OiAyODBweDsgfQouY2hhcnQtd3JhcC50YWxsIHsgaGVpZ2h0OiAzNDBweDsgfQoKLmxlZ2VuZC1yb3cgeyBkaXNwbGF5OiBmbGV4OyBmbGV4LXdyYXA6IHdyYXA7IGdhcDogMTJweDsgbWFyZ2luLXRvcDogMTBweDsgZm9udC1zaXplOiAxMnB4OyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyB9Ci5sZWdlbmQtaXRlbSB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogNnB4OyB9Ci5sZWdlbmQtc3dhdGNoIHsgd2lkdGg6IDEwcHg7IGhlaWdodDogMTBweDsgYm9yZGVyLXJhZGl1czogM3B4OyBkaXNwbGF5OiBpbmxpbmUtYmxvY2s7IH0KLmxlZ2VuZC1saW5lIHsgd2lkdGg6IDE0cHg7IGhlaWdodDogMnB4OyBib3JkZXItcmFkaXVzOiAycHg7IGRpc3BsYXk6IGlubGluZS1ibG9jazsgfQoKLyogLS0tLS0tLS0tLSBUYWJsZXMg4oCUIHByZW1pdW0gZGF0YWJhc2UgZmVlbCwgbm90IGEgc3ByZWFkc2hlZXQgLS0tLS0tLS0tLSAqLwoudGFibGUtc2Nyb2xsIHsKICBvdmVyZmxvdy14OiBhdXRvOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtbWQpOwogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMik7Cn0KLmRhdGEtdGFibGUgeyB3aWR0aDogMTAwJTsgYm9yZGVyLWNvbGxhcHNlOiBzZXBhcmF0ZTsgYm9yZGVyLXNwYWNpbmc6IDA7IGZvbnQtc2l6ZTogMTNweDsgfQouZGF0YS10YWJsZSB0aCwgLmRhdGEtdGFibGUgdGQgeyB0ZXh0LWFsaWduOiBsZWZ0OyBwYWRkaW5nOiAxMXB4IDE0cHg7IGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1ncmlkbGluZSk7IHdoaXRlLXNwYWNlOiBub3dyYXA7IH0KLmRhdGEtdGFibGUgdGQud3JhcCB7IHdoaXRlLXNwYWNlOiBub3JtYWw7IH0KLmRhdGEtdGFibGUgdGhlYWQgdGggewogIGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IGZvbnQtd2VpZ2h0OiA2MDA7IGZvbnQtc2l6ZTogMTFweDsgdGV4dC10cmFuc2Zvcm06IHVwcGVyY2FzZTsgbGV0dGVyLXNwYWNpbmc6IDAuMDRlbTsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTEpOyBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgcG9zaXRpb246IHN0aWNreTsgdG9wOiAwOyB6LWluZGV4OiAxOwp9Ci5kYXRhLXRhYmxlIHRoZWFkIHRoLnNvcnRhYmxlLXRoIHsgY3Vyc29yOiBwb2ludGVyOyB1c2VyLXNlbGVjdDogbm9uZTsgdHJhbnNpdGlvbjogY29sb3IgMTUwbXMgdmFyKC0tZWFzZSk7IH0KLmRhdGEtdGFibGUgdGhlYWQgdGguc29ydGFibGUtdGg6aG92ZXIgeyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgfQouZGF0YS10YWJsZSB0aGVhZCB0aCAuc29ydC1hcnJvdyB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgZm9udC1zaXplOiAxMHB4OyBtYXJnaW4tbGVmdDogMnB4OyB9Ci5kYXRhLXRhYmxlIHRoZWFkIHRoLnNvcnRhYmxlLXRoOmhvdmVyIC5zb3J0LWFycm93IHsgY29sb3I6IHZhcigtLXNlcmllcy0xKTsgfQouZGF0YS10YWJsZSB0Ym9keSB0cjpudGgtY2hpbGQoZXZlbikgeyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tdGV4dC1tdXRlZCkgNCUsIHRyYW5zcGFyZW50KTsgfQouZGF0YS10YWJsZSB0ZC5udW0geyBmb250LXZhcmlhbnQtbnVtZXJpYzogdGFidWxhci1udW1zOyB0ZXh0LWFsaWduOiByaWdodDsgfQouZGF0YS10YWJsZSB0aC5udW0geyB0ZXh0LWFsaWduOiByaWdodDsgfQouZGF0YS10YWJsZSB0Ym9keSB0ciB7IHRyYW5zaXRpb246IGJhY2tncm91bmQgMTUwbXMgdmFyKC0tZWFzZSk7IH0KLmRhdGEtdGFibGUgdGJvZHkgdHI6aG92ZXIgeyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDclLCB0cmFuc3BhcmVudCk7IH0KLmRhdGEtdGFibGUgdGJvZHkgdHI6bGFzdC1jaGlsZCB0ZCB7IGJvcmRlci1ib3R0b206IG5vbmU7IH0KLnBsYXRmb3JtLXBpbGwgewogIGRpc3BsYXk6IGlubGluZS1mbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDZweDsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNjAwOwogIHBhZGRpbmc6IDRweCAxMHB4OyBib3JkZXItcmFkaXVzOiAyMHB4OyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTEpOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOwp9Ci5wbGF0Zm9ybS1kb3QgeyB3aWR0aDogOHB4OyBoZWlnaHQ6IDhweDsgYm9yZGVyLXJhZGl1czogNTAlOyB9CgovKiAtLS0tLS0tLS0tIEJ1dHRvbnMg4oCUIG5ldmVyIGZsYXQ6IHNvZnQgc2hhZG93LCBob3ZlciBsaWZ0LCBwcmVzcyBzY2FsZSAtLS0tLS0tLS0tICovCi5idG4gewogIGRpc3BsYXk6IGlubGluZS1mbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6IGNlbnRlcjsgZ2FwOiA2cHg7CiAgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0yKTsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHBhZGRpbmc6IDlweCAxN3B4OyBib3JkZXItcmFkaXVzOiAxMXB4OyBjdXJzb3I6IHBvaW50ZXI7IGZvbnQtc2l6ZTogMTNweDsgZm9udC13ZWlnaHQ6IDYwMDsKICBib3gtc2hhZG93OiAwIDFweCAycHggcmdiYSgxNSwxNywyMSwwLjA0KTsKICB0cmFuc2l0aW9uOiB0cmFuc2Zvcm0gMTUwbXMgdmFyKC0tZWFzZSksIGJveC1zaGFkb3cgMTUwbXMgdmFyKC0tZWFzZSksIGZpbHRlciAxNTBtcyB2YXIoLS1lYXNlKSwgYmFja2dyb3VuZCAxNTBtcyB2YXIoLS1lYXNlKTsKfQouYnRuIHN2ZyB7IGZsZXgtc2hyaW5rOiAwOyB9Ci5idG46aG92ZXIgeyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoLTFweCk7IGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdy1ob3Zlcik7IGZpbHRlcjogYnJpZ2h0bmVzcygxLjAyKTsgfQouYnRuOmFjdGl2ZSB7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgwKSBzY2FsZSgwLjk2KTsgYm94LXNoYWRvdzogMCAxcHggMnB4IHJnYmEoMTUsMTcsMjEsMC4wNik7IH0KLmJ0bi5wcmltYXJ5IHsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zZXJpZXMtMSk7IGNvbG9yOiAjZmZmOyBib3JkZXItY29sb3I6IHRyYW5zcGFyZW50OwogIGJveC1zaGFkb3c6IDAgNHB4IDE0cHggLTVweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDY1JSwgdHJhbnNwYXJlbnQpOwp9Ci5idG4ucHJpbWFyeTpob3ZlciB7IGZpbHRlcjogYnJpZ2h0bmVzcygxLjA3KTsgYm94LXNoYWRvdzogMCA4cHggMjJweCAtNnB4IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zZXJpZXMtMSkgNzAlLCB0cmFuc3BhcmVudCk7IH0KLmJ0bi5kYW5nZXIgewogIGJhY2tncm91bmQ6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IGNvbG9yOiAjZmZmOyBib3JkZXItY29sb3I6IHRyYW5zcGFyZW50OwogIGJveC1zaGFkb3c6IDAgNHB4IDE0cHggLTVweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc3RhdHVzLWNyaXRpY2FsKSA1NSUsIHRyYW5zcGFyZW50KTsKfQouYnRuLmRhbmdlcjpob3ZlciB7IGZpbHRlcjogYnJpZ2h0bmVzcygxLjA2KTsgYm94LXNoYWRvdzogMCA4cHggMjJweCAtNnB4IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zdGF0dXMtY3JpdGljYWwpIDYwJSwgdHJhbnNwYXJlbnQpOyB9Ci5idG4uc3VjY2VzcyB7CiAgYmFja2dyb3VuZDogdmFyKC0tc3RhdHVzLWdvb2QpOyBjb2xvcjogI2ZmZjsgYm9yZGVyLWNvbG9yOiB0cmFuc3BhcmVudDsKICBib3gtc2hhZG93OiAwIDRweCAxNHB4IC01cHggY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXN0YXR1cy1nb29kKSA1NSUsIHRyYW5zcGFyZW50KTsKfQouYnRuLnN1Y2Nlc3M6aG92ZXIgeyBmaWx0ZXI6IGJyaWdodG5lc3MoMS4wNik7IGJveC1zaGFkb3c6IDAgOHB4IDIycHggLTZweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc3RhdHVzLWdvb2QpIDYwJSwgdHJhbnNwYXJlbnQpOyB9Ci5idG46ZGlzYWJsZWQgeyBvcGFjaXR5OiAwLjQ1OyBjdXJzb3I6IG5vdC1hbGxvd2VkOyB0cmFuc2Zvcm06IG5vbmU7IGJveC1zaGFkb3c6IG5vbmU7IGZpbHRlcjogbm9uZTsgfQouYnRuLXJvdyB7IGRpc3BsYXk6IGZsZXg7IGdhcDogOHB4OyBmbGV4LXdyYXA6IHdyYXA7IH0KCi8qIC0tLS0tLS0tLS0gVXBsb2FkIC0tLS0tLS0tLS0gKi8KLmRyb3B6b25lIHsKICBib3JkZXI6IDJweCBkYXNoZWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLWxnKTsgcGFkZGluZzogNDBweCAyMHB4OwogIHRleHQtYWxpZ246IGNlbnRlcjsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIGN1cnNvcjogcG9pbnRlcjsgdHJhbnNpdGlvbjogYm9yZGVyLWNvbG9yIDIwMG1zIHZhcigtLWVhc2UpLCBiYWNrZ3JvdW5kIDIwMG1zIHZhcigtLWVhc2UpLCB0cmFuc2Zvcm0gMjAwbXMgdmFyKC0tZWFzZSk7Cn0KLmRyb3B6b25lOmhvdmVyIHsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKC0xcHgpOyB9Ci5kcm9wem9uZS5pcy1kcmFnIHsgYm9yZGVyLWNvbG9yOiB2YXIoLS1zZXJpZXMtMSk7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zZXJpZXMtMSkgNiUsIHZhcigtLXN1cmZhY2UtMikpOyB0cmFuc2Zvcm06IHNjYWxlKDEuMDA1KTsgfQouZHJvcHpvbmUgaDMgeyBtYXJnaW46IDAgMCA2cHg7IGZvbnQtc2l6ZTogMTVweDsgfQouZHJvcHpvbmUgcCB7IG1hcmdpbjogMDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgZm9udC1zaXplOiAxM3B4OyB9Ci5kcm9wem9uZSBpbnB1dFt0eXBlPSJmaWxlIl0geyBkaXNwbGF5OiBub25lOyB9CgouY29uZmxpY3QtbGlzdCB7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGdhcDogOHB4OyBtYXJnaW46IDEycHggMDsgfQouY29uZmxpY3Qtcm93IHsKICBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6IHNwYWNlLWJldHdlZW47IGdhcDogMTJweDsKICBwYWRkaW5nOiAxMXB4IDE0cHg7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1zbSk7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMik7CiAgdHJhbnNpdGlvbjogYm94LXNoYWRvdyAxODBtcyB2YXIoLS1lYXNlKTsKfQouY29uZmxpY3Qtcm93OmhvdmVyIHsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWNhcmQpOyB9Ci5jb25mbGljdC1yb3cgLndlZWstbGFiZWwgeyBmb250LXdlaWdodDogNjAwOyBmb250LXNpemU6IDEzcHg7IH0KLmNvbmZsaWN0LXJvdyAud2Vlay1tZXRhIHsgZm9udC1zaXplOiAxMnB4OyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyB9Ci5jb25mbGljdC1yb3cgc2VsZWN0IHsgbWluLXdpZHRoOiAwOyB9CgouYmFkZ2UgeyBkaXNwbGF5OiBpbmxpbmUtYmxvY2s7IHBhZGRpbmc6IDNweCAxMHB4OyBib3JkZXItcmFkaXVzOiAyMHB4OyBmb250LXNpemU6IDExcHg7IGZvbnQtd2VpZ2h0OiA3MDA7IH0KLmJhZGdlLnN1Y2Nlc3MgeyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc3RhdHVzLWdvb2QpIDE4JSwgdHJhbnNwYXJlbnQpOyBjb2xvcjogdmFyKC0tc3RhdHVzLWdvb2QpOyB9Ci5iYWRnZS5wYXJ0aWFsIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXN0YXR1cy13YXJuaW5nKSAyNSUsIHRyYW5zcGFyZW50KTsgY29sb3I6ICM4YTYzMDA7IH0KLmJhZGdlLmZhaWxlZCB7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zdGF0dXMtY3JpdGljYWwpIDE4JSwgdHJhbnNwYXJlbnQpOyBjb2xvcjogdmFyKC0tc3RhdHVzLWNyaXRpY2FsKTsgfQouYmFkZ2UuZXJyb3Itc2V2IHsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IH0KLmJhZGdlLndhcm5pbmctc2V2IHsgY29sb3I6ICM4YTYzMDA7IH0KLmJhZGdlLnNraXAtc2V2IHsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyB9CgouaXNzdWVzLWxpc3QgeyBtYXgtaGVpZ2h0OiAyMjBweDsgb3ZlcmZsb3cteTogYXV0bzsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLXNtKTsgfQouaXNzdWUtcm93IHsgcGFkZGluZzogOXB4IDE0cHg7IGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1ncmlkbGluZSk7IGZvbnQtc2l6ZTogMTJweDsgfQouaXNzdWUtcm93Omxhc3QtY2hpbGQgeyBib3JkZXItYm90dG9tOiBub25lOyB9Ci5pc3N1ZS1yb3cgLnJvdy1ubyB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgbWFyZ2luLXJpZ2h0OiA2cHg7IH0KCi8qIC0tLS0tLS0tLS0gVG9hc3QgLS0tLS0tLS0tLSAqLwoudG9hc3Qtcm9vdCB7IHBvc2l0aW9uOiBmaXhlZDsgYm90dG9tOiAyMHB4OyByaWdodDogMjBweDsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiA4cHg7IHotaW5kZXg6IDEwMDsgfQoudG9hc3QgewogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1zbSk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHBhZGRpbmc6IDEycHggMTZweDsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LW1vZGFsKTsgZm9udC1zaXplOiAxM3B4OyBtYXgtd2lkdGg6IDM0MHB4OwogIGFuaW1hdGlvbjogdG9hc3QtaW4gMjIwbXMgdmFyKC0tZWFzZSk7Cn0KLnRvYXN0LnN1Y2Nlc3MgeyBib3JkZXItbGVmdDogM3B4IHNvbGlkIHZhcigtLXN0YXR1cy1nb29kKTsgfQoudG9hc3QuZXJyb3IgeyBib3JkZXItbGVmdDogM3B4IHNvbGlkIHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IH0KQGtleWZyYW1lcyB0b2FzdC1pbiB7IGZyb20geyBvcGFjaXR5OiAwOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMTBweCkgc2NhbGUoMC45OCk7IH0gdG8geyBvcGFjaXR5OiAxOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMCkgc2NhbGUoMSk7IH0gfQoKLyogLS0tLS0tLS0tLSBNaXNjIC0tLS0tLS0tLS0gKi8KLm11dGVkIHsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyB9Ci5lbXB0eS1zdGF0ZSB7CiAgcGFkZGluZzogNTZweCAyNHB4OyB0ZXh0LWFsaWduOiBjZW50ZXI7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7CiAgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiAxMnB4OwogIGFuaW1hdGlvbjogY2FyZEluIDI2MG1zIHZhcigtLWVhc2UpOwp9Ci5lbXB0eS1zdGF0ZSAuZW1wdHktaWNvbiB7CiAgd2lkdGg6IDUycHg7IGhlaWdodDogNTJweDsgYm9yZGVyLXJhZGl1czogMTZweDsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsganVzdGlmeS1jb250ZW50OiBjZW50ZXI7CiAgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSAxMCUsIHRyYW5zcGFyZW50KTsgY29sb3I6IHZhcigtLXNlcmllcy0xKTsKfQouZW1wdHktc3RhdGUgLmVtcHR5LXRpdGxlIHsgZm9udC1zaXplOiAxNHB4OyBmb250LXdlaWdodDogNjAwOyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgfQouZW1wdHktc3RhdGUgLmVtcHR5LW1lc3NhZ2UgeyBmb250LXNpemU6IDEzcHg7IG1heC13aWR0aDogMzYwcHg7IH0KLnNwaW5uZXIgeyB3aWR0aDogMTZweDsgaGVpZ2h0OiAxNnB4OyBib3JkZXItcmFkaXVzOiA1MCU7IGJvcmRlcjogMnB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci10b3AtY29sb3I6IHZhcigtLXNlcmllcy0xKTsgYW5pbWF0aW9uOiBzcGluIC42cyBsaW5lYXIgaW5maW5pdGU7IGRpc3BsYXk6IGlubGluZS1ibG9jazsgfQpAa2V5ZnJhbWVzIHNwaW4geyB0byB7IHRyYW5zZm9ybTogcm90YXRlKDM2MGRlZyk7IH0gfQoubG9hZGluZy1yb3cgeyBwYWRkaW5nOiA0MHB4IDIwcHg7IHRleHQtYWxpZ246IGNlbnRlcjsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgfQoKLyogU2tlbGV0b24gbG9hZGVycyDigJQgc2hpbW1lcmluZyBwbGFjZWhvbGRlcnMgc2hvd24gd2hpbGUgYSBzZWN0aW9uJ3MgZGF0YSBpcyBpbiBmbGlnaHQgKi8KLnNrZWxldG9uIHsKICBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtc20pOwogIGJhY2tncm91bmQ6IGxpbmVhci1ncmFkaWVudCgxMDBkZWcsIGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS10ZXh0LW11dGVkKSAxMiUsIHRyYW5zcGFyZW50KSAzMCUsIGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS10ZXh0LW11dGVkKSAyMiUsIHRyYW5zcGFyZW50KSA1MCUsIGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS10ZXh0LW11dGVkKSAxMiUsIHRyYW5zcGFyZW50KSA3MCUpOwogIGJhY2tncm91bmQtc2l6ZTogMjAwJSAxMDAlOwogIGFuaW1hdGlvbjogc2tlbGV0b25TaGltbWVyIDEuNHMgZWFzZS1pbi1vdXQgaW5maW5pdGU7Cn0KQGtleWZyYW1lcyBza2VsZXRvblNoaW1tZXIgeyBmcm9tIHsgYmFja2dyb3VuZC1wb3NpdGlvbjogMTUwJSAwOyB9IHRvIHsgYmFja2dyb3VuZC1wb3NpdGlvbjogLTUwJSAwOyB9IH0KLnNrZWxldG9uLXN0YXQtZ3JpZCB7IGRpc3BsYXk6IGdyaWQ7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogcmVwZWF0KGF1dG8tZml0LCBtaW5tYXgoMTgwcHgsIDFmcikpOyBnYXA6IDE0cHg7IH0KLnNrZWxldG9uLXRpbGUgeyBoZWlnaHQ6IDg0cHg7IH0KLnNrZWxldG9uLWNoYXJ0IHsgaGVpZ2h0OiAyODBweDsgd2lkdGg6IDEwMCU7IH0KLnNrZWxldG9uLXJvdyB7IGhlaWdodDogNDBweDsgbWFyZ2luLWJvdHRvbTogOHB4OyB9CgovKiBBbmltYXRlZCBob3Jpem9udGFsIGNvbXBhcmlzb24gYmFyIOKAlCBhIGxhYmVsZWQgcm93IHdpdGggYSB0cmFjayB0aGF0IGZpbGxzIGluIG9uIGluc2VydGlvbiAqLwouYmFyLXJvdyB7IGRpc3BsYXk6IGdyaWQ7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogbWlubWF4KDkwcHgsIDE0MHB4KSAxZnIgYXV0bzsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiAxMHB4OyBwYWRkaW5nOiA1cHggMDsgfQouYmFyLWxhYmVsIHsgZm9udC1zaXplOiAxMnB4OyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyBmb250LXdlaWdodDogNjAwOyB9Ci5iYXItdHJhY2sgeyBoZWlnaHQ6IDhweDsgYm9yZGVyLXJhZGl1czogNXB4OyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tdGV4dC1tdXRlZCkgMTQlLCB0cmFuc3BhcmVudCk7IG92ZXJmbG93OiBoaWRkZW47IH0KLmJhci1maWxsIHsgaGVpZ2h0OiAxMDAlOyB3aWR0aDogMCU7IGJvcmRlci1yYWRpdXM6IDVweDsgdHJhbnNpdGlvbjogd2lkdGggNzAwbXMgY3ViaWMtYmV6aWVyKDAuMTYsIDEsIDAuMywgMSk7IH0KLmJhci12YWx1ZSB7IGZvbnQtc2l6ZTogMTJweDsgZm9udC13ZWlnaHQ6IDcwMDsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IGZvbnQtdmFyaWFudC1udW1lcmljOiB0YWJ1bGFyLW51bXM7IHRleHQtYWxpZ246IHJpZ2h0OyBtaW4td2lkdGg6IDU2cHg7IH0KCkBtZWRpYSAocHJlZmVycy1yZWR1Y2VkLW1vdGlvbjogcmVkdWNlKSB7CiAgLmJhci1maWxsIHsgdHJhbnNpdGlvbi1kdXJhdGlvbjogMW1zOyB9CiAgLnNrZWxldG9uIHsgYW5pbWF0aW9uLWR1cmF0aW9uOiAxbXM7IH0KICAuY2FyZCwgLnN0YXQtdGlsZSB7IGFuaW1hdGlvbi1kdXJhdGlvbjogMW1zOyB9Cn0KCi50d28tY29sIHsgZGlzcGxheTogZ3JpZDsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiAxZnIgMWZyOyBnYXA6IDE2cHg7IH0KQG1lZGlhIChtYXgtd2lkdGg6IDkwMHB4KSB7IC50d28tY29sIHsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiAxZnI7IH0gfQoKLm1vZGUtdGFicyB7IGRpc3BsYXk6IGZsZXg7IGdhcDogNnB4OyBmbGV4LXdyYXA6IHdyYXA7IG1hcmdpbi1ib3R0b206IDE2cHg7IH0KLm1vZGUtdGFicyBidXR0b24gewogIGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHBhZGRpbmc6IDdweCAxNHB4OyBib3JkZXItcmFkaXVzOiAyMHB4OyBmb250LXNpemU6IDEycHg7IGZvbnQtd2VpZ2h0OiA2MDA7IGN1cnNvcjogcG9pbnRlcjsKICB0cmFuc2l0aW9uOiBjb2xvciAxODBtcyB2YXIoLS1lYXNlKSwgYmFja2dyb3VuZCAxODBtcyB2YXIoLS1lYXNlKSwgdHJhbnNmb3JtIDE1MG1zIHZhcigtLWVhc2UpOwp9Ci5tb2RlLXRhYnMgYnV0dG9uOmhvdmVyIHsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKC0xcHgpOyB9Ci5tb2RlLXRhYnMgYnV0dG9uLmlzLWFjdGl2ZSB7IGJhY2tncm91bmQ6IHZhcigtLXNlcmllcy0xKTsgY29sb3I6ICNmZmY7IGJvcmRlci1jb2xvcjogdHJhbnNwYXJlbnQ7IGJveC1zaGFkb3c6IDAgNHB4IDE0cHggLTVweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDYwJSwgdHJhbnNwYXJlbnQpOyB9CgouZmllbGQtaW5saW5lIHsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA4cHg7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgfQouZmllbGQtaW5saW5lIHNlbGVjdCwgLmZpZWxkLWlubGluZSBpbnB1dCB7IG1pbi13aWR0aDogMDsgcGFkZGluZzogNnB4IDEwcHg7IH0KCi8qIC0tLS0tLS0tLS0gUGxhdGZvcm0gUGVyZm9ybWFuY2UgQ29tcGFyaXNvbiBjYXJkcyAtLS0tLS0tLS0tICovCi5wY2Mtc2VjdGlvbiB7IG1hcmdpbi10b3A6IDI0cHg7IH0KLnBjYy1jb250cm9scyB7IGRpc3BsYXk6IGZsZXg7IGZsZXgtd3JhcDogd3JhcDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiAxNnB4OyBtYXJnaW4tYm90dG9tOiAxNnB4OyB9Ci5wbGF0Zm9ybS1jb21wYXJlLWdyaWQgewogIGRpc3BsYXk6IGdyaWQ7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogcmVwZWF0KDIsIDFmcik7IGdhcDogMTZweDsKfQpAbWVkaWEgKG1heC13aWR0aDogOTAwcHgpIHsgLnBsYXRmb3JtLWNvbXBhcmUtZ3JpZCB7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogMWZyOyB9IH0KLnBsYXRmb3JtLWNvbXBhcmUtY2FyZCB7CiAgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLWxnKTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgcGFkZGluZzogMThweDsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWNhcmQpOwogIHRyYW5zaXRpb246IGJveC1zaGFkb3cgMjIwbXMgdmFyKC0tZWFzZSksIHRyYW5zZm9ybSAyMjBtcyB2YXIoLS1lYXNlKTsKICBhbmltYXRpb246IGNhcmRJbiAzMjBtcyB2YXIoLS1lYXNlKSBiYWNrd2FyZHM7Cn0KLnBsYXRmb3JtLWNvbXBhcmUtY2FyZDpob3ZlciB7IGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdy1ob3Zlcik7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgtMnB4KTsgfQoucGNjLWhlYWRlciB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogc3BhY2UtYmV0d2VlbjsgZ2FwOiAxMHB4OyB9Ci5wY2MtaGVhZGVyLW5hbWUgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDhweDsgfQoucGNjLW5hbWUgeyBmb250LXNpemU6IDE1cHg7IGZvbnQtd2VpZ2h0OiA3MDA7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyB9Ci5wY2MtYmFkZ2UgeyBmb250LXNpemU6IDEzcHg7IGZvbnQtd2VpZ2h0OiA3MDA7IHBhZGRpbmc6IDRweCAxMHB4OyBib3JkZXItcmFkaXVzOiAyMHB4OyB9Ci5wY2MtYmFkZ2UudXAgeyBjb2xvcjogdmFyKC0tc3VjY2Vzcy10ZXh0KTsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXN0YXR1cy1nb29kKSAxNCUsIHRyYW5zcGFyZW50KTsgfQoucGNjLWJhZGdlLmRvd24geyBjb2xvcjogdmFyKC0tc3RhdHVzLWNyaXRpY2FsKTsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXN0YXR1cy1jcml0aWNhbCkgMTIlLCB0cmFuc3BhcmVudCk7IH0KLnBjYy1iYWRnZS5mbGF0IHsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tdGV4dC1tdXRlZCkgMTIlLCB0cmFuc3BhcmVudCk7IH0KLnBjYy1jYXB0aW9uIHsgZm9udC1zaXplOiAxMnB4OyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyBtYXJnaW4tdG9wOiA2cHg7IH0KLnBjYy1tZXRyaWNzIHsgbWFyZ2luLXRvcDogMTZweDsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiAxNHB4OyB9Ci5wY2MtbWV0cmljLXJvdyB7IHBhZGRpbmctdG9wOiAxMnB4OyBib3JkZXItdG9wOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgfQoucGNjLW1ldHJpYy1yb3c6Zmlyc3QtY2hpbGQgeyBwYWRkaW5nLXRvcDogMDsgYm9yZGVyLXRvcDogbm9uZTsgfQoucGNjLW1ldHJpYy1oZWFkZXIgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6IHNwYWNlLWJldHdlZW47IGdhcDogOHB4OyBtYXJnaW4tYm90dG9tOiA2cHg7IH0KLnBjYy1tZXRyaWMtbGFiZWwgeyBmb250LXNpemU6IDEycHg7IGZvbnQtd2VpZ2h0OiA3MDA7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyB9Ci5wY2MtbWV0cmljLWRpZmYgeyBmb250LXNpemU6IDEycHg7IGZvbnQtd2VpZ2h0OiA3MDA7IH0KLnBjYy1tZXRyaWMtZGlmZi51cCB7IGNvbG9yOiB2YXIoLS1zdWNjZXNzLXRleHQpOyB9Ci5wY2MtbWV0cmljLWRpZmYuZG93biB7IGNvbG9yOiB2YXIoLS1zdGF0dXMtY3JpdGljYWwpOyB9Ci5wY2MtbWV0cmljLWRpZmYuZmxhdCB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgfQoucGNjLWZvb3RlciB7IG1hcmdpbi10b3A6IDE2cHg7IHBhZGRpbmctdG9wOiAxNHB4OyBib3JkZXItdG9wOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgfQoucGNjLWZvb3Rlci1sYWJlbCB7IGZvbnQtc2l6ZTogMTFweDsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyBmb250LXdlaWdodDogNjAwOyB0ZXh0LXRyYW5zZm9ybTogdXBwZXJjYXNlOyBsZXR0ZXItc3BhY2luZzogMC4wM2VtOyB9Ci5wY2MtZm9vdGVyLXZhbHVlIHsgZm9udC1zaXplOiAxNXB4OyBmb250LXdlaWdodDogNzAwOyBtYXJnaW4tdG9wOiA0cHg7IH0KLnBjYy1mb290ZXItdmFsdWUudXAgeyBjb2xvcjogdmFyKC0tc3VjY2Vzcy10ZXh0KTsgfQoucGNjLWZvb3Rlci12YWx1ZS5kb3duIHsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IH0KLnBjYy1mb290ZXItdmFsdWUuZmxhdCB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgfQoucGNjLWZvb3Rlci1kZXRhaWwgeyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IG1hcmdpbi10b3A6IDRweDsgfQoucGNjLXZpZXctbGluayB7CiAgZGlzcGxheTogaW5saW5lLWZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogNHB4OyBtYXJnaW4tdG9wOiAxNHB4OwogIGJhY2tncm91bmQ6IG5vbmU7IGJvcmRlcjogbm9uZTsgY29sb3I6IHZhcigtLXNlcmllcy0xKTsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNzAwOyBjdXJzb3I6IHBvaW50ZXI7IHBhZGRpbmc6IDA7CiAgdHJhbnNpdGlvbjogb3BhY2l0eSAxNTBtcyB2YXIoLS1lYXNlKTsKfQoucGNjLXZpZXctbGluazpob3ZlciB7IG9wYWNpdHk6IDAuNzU7IHRleHQtZGVjb3JhdGlvbjogdW5kZXJsaW5lOyB9CgovKiAtLS0tLS0tLS0tIFBhZ2luYXRpb24gLS0tLS0tLS0tLSAqLwoucGFnaW5hdGlvbi1yb3cgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDEycHg7IG1hcmdpbi10b3A6IDE0cHg7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgfQoucGFnaW5hdGlvbi1yb3cgLmJ0biB7IHBhZGRpbmc6IDZweCAxMnB4OyB9Ci5leHBvcnQtYnV0dG9ucyB7IGRpc3BsYXk6IGZsZXg7IGdhcDogOHB4OyBmbGV4LXdyYXA6IHdyYXA7IG1hcmdpbi1ib3R0b206IDEycHg7IH0KLmV4cG9ydC1idXR0b25zIC5idG4geyBwYWRkaW5nOiA3cHggMTNweDsgZm9udC1zaXplOiAxMnB4OyB9CgovKiAtLS0tLS0tLS0tIERhc2hib2FyZCBjb250cm9scyAvIG1ldHJpYy1mb2N1c2VkIEtQSXMgLS0tLS0tLS0tLSAqLwouZGFzaGJvYXJkLWNvbnRyb2xzIHsKICBkaXNwbGF5OiBmbGV4OyBmbGV4LXdyYXA6IHdyYXA7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogMTBweDsgbWFyZ2luLWJvdHRvbTogMThweDsKfQouZGFzaGJvYXJkLWNvbnRyb2xzIGxhYmVsIHsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNjAwOyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyBtYXJnaW4tcmlnaHQ6IDZweDsgfQouZGFzaGJvYXJkLWNvbnRyb2xzIHNlbGVjdCB7IGZvbnQtd2VpZ2h0OiA2MDA7IH0KLyogQmVzdCBQZXJmb3JtaW5nIFBvc3Qg4oCUIGEgZmVhdHVyZWQgbGFuZHNjYXBlIGNhcmQgc3Bhbm5pbmcgMyBLUEktdGlsZS13aWR0aHMKICAgKGEgc3RhbmRhcmQgdGlsZSBpcyAxIHVuaXQ7ICNrcGlHcmlkIGhhcyAxMCB1bml0cyB0b3RhbCksIHNhbWUgZml4ZWQKICAgaGVpZ2h0IGFzIHRoZSByZXN0IG9mICNrcGlHcmlkOiBjYXB0aW9uL3BsYXRmb3JtL2RhdGUgc2l0IG9uIHRoZSBsZWZ0LAogICB3aXRoIHRoZSBzZWxlY3RlZCBtZXRyaWMgKGxhcmdlKSBhbmQgQ3VycmVudCBGb2xsb3dlcnMgKHNtYWxsZXIsIGJlbG93IGEKICAgZGl2aWRlcikgc3RhY2tlZCBpbiBhIG5hcnJvd2VyIGNvbHVtbiBvbiB0aGUgcmlnaHQuICovCiNrcGlHcmlkIC5wb3N0LXRpbGUgewogIGdyaWQtY29sdW1uOiBzcGFuIDM7CiAgZmxleC1kaXJlY3Rpb246IHJvdzsKICBhbGlnbi1pdGVtczogc3RyZXRjaDsKICBqdXN0aWZ5LWNvbnRlbnQ6IGZsZXgtc3RhcnQ7CiAgZ2FwOiAyMHB4OwogIHBhZGRpbmc6IDE0cHggMjBweDsKfQoja3BpR3JpZCAucG9zdC10aWxlLW1haW4geyBmbGV4OiAxIDEgYXV0bzsgbWluLXdpZHRoOiAwOyBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBqdXN0aWZ5LWNvbnRlbnQ6IGNlbnRlcjsgZ2FwOiA2cHg7IH0KI2twaUdyaWQgLnBvc3QtdGlsZS1jYXB0aW9uIHsKICBmb250LXNpemU6IDEzcHg7IGZvbnQtd2VpZ2h0OiA2MDA7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOwogIGxpbmUtaGVpZ2h0OiAxLjQ7CiAgZGlzcGxheTogLXdlYmtpdC1ib3g7IC13ZWJraXQtbGluZS1jbGFtcDogMzsgLXdlYmtpdC1ib3gtb3JpZW50OiB2ZXJ0aWNhbDsgb3ZlcmZsb3c6IGhpZGRlbjsKfQoja3BpR3JpZCAucG9zdC10aWxlLWNhcHRpb24ubXV0ZWQgeyBjb2xvcjogdmFyKC0tdGV4dC1tdXRlZCk7IGZvbnQtd2VpZ2h0OiA1MDA7IC13ZWJraXQtbGluZS1jbGFtcDogMTsgfQoja3BpR3JpZCAucG9zdC10aWxlLW1ldGEgeyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA2cHg7IHdoaXRlLXNwYWNlOiBub3dyYXA7IG92ZXJmbG93OiBoaWRkZW47IHRleHQtb3ZlcmZsb3c6IGVsbGlwc2lzOyB9CiNrcGlHcmlkIC5wb3N0LXRpbGUtZGl2aWRlciB7IGZsZXg6IDAgMCBhdXRvOyB3aWR0aDogMXB4OyBhbGlnbi1zZWxmOiBzdHJldGNoOyBiYWNrZ3JvdW5kOiB2YXIoLS1ib3JkZXIpOyB9CiNrcGlHcmlkIC5wb3N0LXRpbGUtbWV0cmljcyB7CiAgZmxleDogMCAwIGF1dG87IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGp1c3RpZnktY29udGVudDogY2VudGVyOyBnYXA6IDhweDsgbWluLXdpZHRoOiAxMjBweDsKfQoja3BpR3JpZCAucG9zdC10aWxlLW1ldHJpYy1ibG9jayB7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGdhcDogM3B4OyB9CiNrcGlHcmlkIC5wb3N0LXRpbGUtbWV0cmljLWJsb2NrLnNlY29uZGFyeSB7IHBhZGRpbmctdG9wOiA4cHg7IGJvcmRlci10b3A6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyB9CiNrcGlHcmlkIC5wb3N0LXRpbGUtbWV0cmljLWxhYmVsIHsgZm9udC1zaXplOiAxMXB4OyBjb2xvcjogdmFyKC0tdGV4dC1tdXRlZCk7IGZvbnQtd2VpZ2h0OiA2MDA7IHRleHQtdHJhbnNmb3JtOiB1cHBlcmNhc2U7IGxldHRlci1zcGFjaW5nOiAwLjAzZW07IH0KI2twaUdyaWQgLnBvc3QtdGlsZS1tZXRyaWMtdmFsdWUgeyBmb250LXNpemU6IDIycHg7IGZvbnQtd2VpZ2h0OiA3MDA7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyBmb250LXZhcmlhbnQtbnVtZXJpYzogdGFidWxhci1udW1zOyB9CiNrcGlHcmlkIC5wb3N0LXRpbGUtbWV0cmljLWJsb2NrLnNlY29uZGFyeSAucG9zdC10aWxlLW1ldHJpYy12YWx1ZSB7IGZvbnQtc2l6ZTogMTVweDsgfQpAbWVkaWEgKG1heC13aWR0aDogNjQwcHgpIHsgI2twaUdyaWQgLnBvc3QtdGlsZSB7IGdyaWQtY29sdW1uOiBzcGFuIDI7IH0gfQoKLnN0YXQtdmFsdWUtbXV0ZWQgeyBmb250LXNpemU6IDE1cHggIWltcG9ydGFudDsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyBmb250LXdlaWdodDogNjAwOyB9Ci5jYXB0aW9uLWxpbmsgeyBjb2xvcjogdmFyKC0tc2VyaWVzLTEpOyB0ZXh0LWRlY29yYXRpb246IG5vbmU7IH0KLmNhcHRpb24tbGluazpob3ZlciB7IHRleHQtZGVjb3JhdGlvbjogdW5kZXJsaW5lOyB9CgovKiAtLS0tLS0tLS0tIERhdGEgUmVjb3JkcyAocGxhdGZvcm0tZ3JvdXBlZCkgLS0tLS0tLS0tLSAqLwoucmVjb3Jkcy10b29sYmFyIHsKICBkaXNwbGF5OiBmbGV4OyBmbGV4LXdyYXA6IHdyYXA7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogc3BhY2UtYmV0d2VlbjsKICBnYXA6IDEycHg7IG1hcmdpbi1ib3R0b206IDE0cHg7Cn0KLnBsYXRmb3JtLWZpbHRlci1waWxscyB7IGRpc3BsYXk6IGZsZXg7IGZsZXgtd3JhcDogd3JhcDsgZ2FwOiA2cHg7IH0KLnBsYXRmb3JtLWZpbHRlci1waWxscyBidXR0b24gewogIGRpc3BsYXk6IGlubGluZS1mbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDZweDsKICBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTEpOyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOwogIGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBwYWRkaW5nOiA3cHggMTRweDsgYm9yZGVyLXJhZGl1czogMjBweDsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNjAwOyBjdXJzb3I6IHBvaW50ZXI7CiAgdHJhbnNpdGlvbjogY29sb3IgMTgwbXMgdmFyKC0tZWFzZSksIGJhY2tncm91bmQgMTgwbXMgdmFyKC0tZWFzZSksIHRyYW5zZm9ybSAxNTBtcyB2YXIoLS1lYXNlKSwgYm94LXNoYWRvdyAxODBtcyB2YXIoLS1lYXNlKTsKfQoucGxhdGZvcm0tZmlsdGVyLXBpbGxzIGJ1dHRvbjpob3ZlciB7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoLTFweCk7IH0KLnBsYXRmb3JtLWZpbHRlci1waWxscyBidXR0b246YWN0aXZlIHsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDApIHNjYWxlKDAuOTYpOyB9Ci5wbGF0Zm9ybS1maWx0ZXItcGlsbHMgYnV0dG9uLmlzLWFjdGl2ZSB7IGJhY2tncm91bmQ6IHZhcigtLXNlcmllcy0xKTsgY29sb3I6ICNmZmY7IGJvcmRlci1jb2xvcjogdHJhbnNwYXJlbnQ7IGJveC1zaGFkb3c6IDAgNHB4IDE0cHggLTVweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDYwJSwgdHJhbnNwYXJlbnQpOyB9Ci5wbGF0Zm9ybS1maWx0ZXItcGlsbHMgYnV0dG9uLmlzLWFjdGl2ZSAucGxhdGZvcm0tZG90IHsgYm94LXNoYWRvdzogMCAwIDAgMnB4IHJnYmEoMjU1LDI1NSwyNTUsMC41KTsgfQoucmVjb3Jkcy1zZWFyY2ggaW5wdXQgeyBib3JkZXItcmFkaXVzOiAyMHB4OyBtaW4td2lkdGg6IDIyMHB4OyB9Ci5zdGF0dXMtcGlsbCB7IGRpc3BsYXk6IGlubGluZS1ibG9jazsgcGFkZGluZzogM3B4IDEwcHg7IGJvcmRlci1yYWRpdXM6IDIwcHg7IGZvbnQtc2l6ZTogMTFweDsgZm9udC13ZWlnaHQ6IDcwMDsgfQouc3RhdHVzLXBpbGwub3JpZ2luYWwgeyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tdGV4dC1tdXRlZCkgMTUlLCB0cmFuc3BhcmVudCk7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IH0KLnN0YXR1cy1waWxsLmVkaXRlZCB7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zdGF0dXMtd2FybmluZykgMjIlLCB0cmFuc3BhcmVudCk7IGNvbG9yOiAjOGE2MzAwOyB9Ci5yb3ctYWN0aW9ucyB7IGRpc3BsYXk6IGZsZXg7IGdhcDogNnB4OyBmbGV4LXdyYXA6IG5vd3JhcDsgfQoucm93LWFjdGlvbnMgLmJ0biB7IHBhZGRpbmc6IDVweCAxMHB4OyBmb250LXNpemU6IDEycHg7IH0KLmxpbmstY2VsbCBhIHsgY29sb3I6IHZhcigtLXNlcmllcy0xKTsgdGV4dC1kZWNvcmF0aW9uOiBub25lOyBmb250LXdlaWdodDogNjAwOyBmb250LXNpemU6IDEycHg7IH0KLmxpbmstY2VsbCBhOmhvdmVyIHsgdGV4dC1kZWNvcmF0aW9uOiB1bmRlcmxpbmU7IH0KLnJlY29yZC1zZWN0aW9uIHsKICBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtc20pOyBwYWRkaW5nOiAxNnB4OyBtYXJnaW4tYm90dG9tOiAxNHB4OwogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMik7IGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKfQoucmVjb3JkLXNlY3Rpb24gaDQgeyBtYXJnaW46IDAgMCAxMnB4OyBmb250LXNpemU6IDEycHg7IGZvbnQtd2VpZ2h0OiA3MDA7IGxldHRlci1zcGFjaW5nOiAwLjAzZW07IHRleHQtdHJhbnNmb3JtOiB1cHBlcmNhc2U7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IH0KLnJlY29yZC1zZWN0aW9uIC5mb3JtLWdyaWQgeyBtYXJnaW4tYm90dG9tOiAwOyB9Ci5yZWNvcmQtc2VjdGlvbiAudmlldy1maWVsZCB7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGdhcDogMnB4OyBmb250LXNpemU6IDEzcHg7IH0KLnJlY29yZC1zZWN0aW9uIC52aWV3LWZpZWxkIC52aWV3LWxhYmVsIHsgZm9udC1zaXplOiAxMXB4OyBmb250LXdlaWdodDogNjAwOyBjb2xvcjogdmFyKC0tdGV4dC1tdXRlZCk7IH0KLnJlY29yZC1zZWN0aW9uIC52aWV3LWZpZWxkIC52aWV3LXZhbHVlIHsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IHdvcmQtYnJlYWs6IGJyZWFrLXdvcmQ7IH0KQG1lZGlhIChtYXgtd2lkdGg6IDY0MHB4KSB7CiAgLnJlY29yZHMtdG9vbGJhciB7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGFsaWduLWl0ZW1zOiBzdHJldGNoOyB9CiAgLnJlY29yZHMtc2VhcmNoIGlucHV0IHsgd2lkdGg6IDEwMCU7IH0KfQoKLyogLS0tLS0tLS0tLSBNb2RhbCAocmVjb3JkIGVkaXRvcikgLS0tLS0tLS0tLSAqLwoubW9kYWwtb3ZlcmxheSB7CiAgcG9zaXRpb246IGZpeGVkOyBpbnNldDogMDsgYmFja2dyb3VuZDogcmdiYSgxMCwxMSwxMywwLjUpOwogIGJhY2tkcm9wLWZpbHRlcjogYmx1cig2cHgpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogYmx1cig2cHgpOwogIGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBmbGV4LXN0YXJ0OyBqdXN0aWZ5LWNvbnRlbnQ6IGNlbnRlcjsKICBwYWRkaW5nOiA0MHB4IDE2cHg7IG92ZXJmbG93LXk6IGF1dG87IHotaW5kZXg6IDIwMDsKICBhbmltYXRpb246IG92ZXJsYXlJbiAyMDBtcyB2YXIoLS1lYXNlKTsKfQpAa2V5ZnJhbWVzIG92ZXJsYXlJbiB7IGZyb20geyBvcGFjaXR5OiAwOyB9IHRvIHsgb3BhY2l0eTogMTsgfSB9CkBrZXlmcmFtZXMgbW9kYWxQYW5lbEluIHsKICBmcm9tIHsgb3BhY2l0eTogMDsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDE0cHgpIHNjYWxlKDAuOTcpOyB9CiAgdG8geyBvcGFjaXR5OiAxOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMCkgc2NhbGUoMSk7IH0KfQoubW9kYWwtcGFuZWwgewogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1sZyk7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHBhZGRpbmc6IDI0cHg7IHdpZHRoOiAxMDAlOyBtYXgtd2lkdGg6IDcyMHB4OyBib3gtc2hhZG93OiB2YXIoLS1zaGFkb3ctbW9kYWwpOwogIG1heC1oZWlnaHQ6IGNhbGMoMTAwdmggLSA4MHB4KTsgb3ZlcmZsb3cteTogYXV0bzsKICBhbmltYXRpb246IG1vZGFsUGFuZWxJbiAyNDBtcyB2YXIoLS1lYXNlKTsKfQoubW9kYWwtcGFuZWwud2lkZSB7IG1heC13aWR0aDogMTEwMHB4OyB9Ci5tb2RhbC1wYW5lbCBoMiB7IG1hcmdpbjogMCAwIDRweDsgZm9udC1zaXplOiAxN3B4OyBsZXR0ZXItc3BhY2luZzogLTAuMDFlbTsgfQoubW9kYWwtcGFuZWwgLm1vZGFsLXN1YiB7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IGZvbnQtc2l6ZTogMTJweDsgbWFyZ2luOiAwIDAgMThweDsgfQouZm9ybS1ncmlkIHsgZGlzcGxheTogZ3JpZDsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiByZXBlYXQoYXV0by1maXQsIG1pbm1heCgyMDBweCwgMWZyKSk7IGdhcDogMTJweDsgbWFyZ2luLWJvdHRvbTogMTZweDsgfQouZm9ybS1ncmlkLmZ1bGwgeyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IDFmcjsgfQpAbWVkaWEgKG1heC13aWR0aDogNjQwcHgpIHsgLmZvcm0tZ3JpZCB7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogMWZyOyB9IH0KLmZvcm0tZmllbGQgeyBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBnYXA6IDVweDsgZm9udC1zaXplOiAxMnB4OyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyB9Ci5mb3JtLWZpZWxkIGxhYmVsIHsgZm9udC13ZWlnaHQ6IDYwMDsgfQouZm9ybS1maWVsZCB0ZXh0YXJlYSB7IHJlc2l6ZTogdmVydGljYWw7IG1pbi1oZWlnaHQ6IDYwcHg7IH0KCi5wbGF0Zm9ybS1lZGl0LXJvdyB7CiAgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLXNtKTsgcGFkZGluZzogMTRweDsgbWFyZ2luLWJvdHRvbTogMTBweDsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0yKTsKfQoucGxhdGZvcm0tZWRpdC1yb3cgLnBsYXRmb3JtLWVkaXQtaGVhZCB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogc3BhY2UtYmV0d2VlbjsgZ2FwOiA4cHg7IG1hcmdpbi1ib3R0b206IDEwcHg7IH0KLnBsYXRmb3JtLWVkaXQtcm93IC5tZXRyaWNzLWdyaWQgeyBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdChhdXRvLWZpdCwgbWlubWF4KDEyMHB4LCAxZnIpKTsgZ2FwOiA4cHg7IH0KLnJlbW92ZS1wbGF0Zm9ybS1idG4geyBib3JkZXI6IG5vbmU7IGJhY2tncm91bmQ6IHRyYW5zcGFyZW50OyBjb2xvcjogdmFyKC0tc3RhdHVzLWNyaXRpY2FsKTsgY3Vyc29yOiBwb2ludGVyOyBmb250LXNpemU6IDEycHg7IGZvbnQtd2VpZ2h0OiA2MDA7IHRyYW5zaXRpb246IG9wYWNpdHkgMTUwbXMgdmFyKC0tZWFzZSk7IH0KLnJlbW92ZS1wbGF0Zm9ybS1idG46aG92ZXIgeyBvcGFjaXR5OiAwLjc7IH0KLm1vZGFsLWFjdGlvbnMgeyBkaXNwbGF5OiBmbGV4OyBqdXN0aWZ5LWNvbnRlbnQ6IHNwYWNlLWJldHdlZW47IGFsaWduLWl0ZW1zOiBjZW50ZXI7IG1hcmdpbi10b3A6IDE4cHg7IGdhcDogOHB4OyBmbGV4LXdyYXA6IHdyYXA7IH0KCi8qIC0tLS0tLS0tLS0gUmVzcG9uc2l2ZSB0aWdodGVuaW5nIC0tLS0tLS0tLS0gKi8KQG1lZGlhIChtYXgtd2lkdGg6IDcyMHB4KSB7CiAgLmFwcC1zaGVsbCB7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IH0KICAuc2lkZWJhciB7IHdpZHRoOiAxMDAlOyBoZWlnaHQ6IGF1dG87IHBvc2l0aW9uOiBzdGF0aWM7IGZsZXgtZGlyZWN0aW9uOiByb3c7IGZsZXgtd3JhcDogd3JhcDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA4cHg7IHBhZGRpbmc6IDEwcHggMTRweDsgfQogIC5zaWRlYmFyLWJyYW5kIHsgcGFkZGluZzogMDsgbWFyZ2luLXJpZ2h0OiBhdXRvOyB9CiAgLnRhYnMgeyBmbGV4LWRpcmVjdGlvbjogcm93OyB3aWR0aDogMTAwJTsgb3ZlcmZsb3cteDogYXV0bzsgb3JkZXI6IDM7IH0KICAuc2lkZWJhci1mb290ZXIgeyBmbGV4LWRpcmVjdGlvbjogcm93OyBib3JkZXItdG9wOiBub25lOyBtYXJnaW4tdG9wOiAwOyBwYWRkaW5nLXRvcDogMDsgfQogIC52aWV3LWFyZWEgeyBwYWRkaW5nOiAxNHB4OyB9CiAgLmZpbHRlci1iYXIgeyB0b3A6IGF1dG87IHBvc2l0aW9uOiBzdGF0aWM7IHBhZGRpbmc6IDEycHggMTRweDsgfQogIC5zdGF0LWdyaWQgeyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdChhdXRvLWZpdCwgbWlubWF4KDE0MHB4LCAxZnIpKTsgfQogIC5icmFuZC1sb2dvIHsgaGVpZ2h0OiAyMnB4OyB9Cn0KPC9zdHlsZT4KPC9oZWFkPgo8Ym9keT4KPGRpdiBjbGFzcz0iYXV0aC1zY3JlZW4iIGlkPSJhdXRoU2NyZWVuIj4KICA8ZGl2IGNsYXNzPSJhdXRoLWNhcmQiPgogICAgPGRpdiBjbGFzcz0iYXV0aC1icmFuZCI+CiAgICAgIDxpbWcgY2xhc3M9ImJyYW5kLWxvZ28iIGFsdD0iTGlnb24tUmF6b24gU29sdXRpb25zIGxvZ28iIC8+CiAgICAgIDxzcGFuIGNsYXNzPSJicmFuZC10aXRsZSI+U29jaWFsIE1lZGlhIEFuYWx5dGljczwvc3Bhbj4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0iYXV0aC1mb3JtIj4KICAgICAgPGRpdiBjbGFzcz0iZm9ybS1maWVsZCI+CiAgICAgICAgPGxhYmVsIGZvcj0iYXV0aENvZGUiPkFjY2VzcyBjb2RlPC9sYWJlbD4KICAgICAgICA8aW5wdXQgdHlwZT0icGFzc3dvcmQiIGlkPSJhdXRoQ29kZSIgYXV0b2NvbXBsZXRlPSJvZmYiIC8+CiAgICAgIDwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJhdXRoLWVycm9yIiBpZD0iYXV0aEVycm9yIj48L2Rpdj4KICAgICAgPGJ1dHRvbiBjbGFzcz0iYnRuIHByaW1hcnkiIGlkPSJhdXRoU3VibWl0QnRuIiB0eXBlPSJidXR0b24iPjxpIGRhdGEtbHVjaWRlPSJhcnJvdy1yaWdodCIgc3R5bGU9IndpZHRoOjE0cHg7aGVpZ2h0OjE0cHg7Ij48L2k+IEVudGVyPC9idXR0b24+CiAgICA8L2Rpdj4KICA8L2Rpdj4KPC9kaXY+Cgo8ZGl2IGNsYXNzPSJhcHAtc2hlbGwiIGlkPSJhcHBTaGVsbCIgc3R5bGU9ImRpc3BsYXk6bm9uZTsiPgogIDxhc2lkZSBjbGFzcz0ic2lkZWJhciI+CiAgICA8ZGl2IGNsYXNzPSJzaWRlYmFyLWJyYW5kIj4KICAgICAgPGltZyBjbGFzcz0iYnJhbmQtbG9nbyIgYWx0PSJMaWdvbi1SYXpvbiBTb2x1dGlvbnMgbG9nbyIgLz4KICAgICAgPHNwYW4gY2xhc3M9ImJyYW5kLXRpdGxlIj5Tb2NpYWwgTWVkaWEgQW5hbHl0aWNzPC9zcGFuPgogICAgPC9kaXY+CiAgICA8bmF2IGNsYXNzPSJ0YWJzIiByb2xlPSJ0YWJsaXN0IiBhcmlhLWxhYmVsPSJTZWN0aW9ucyI+CiAgICAgIDxidXR0b24gY2xhc3M9InRhYi1idG4gaXMtYWN0aXZlIiBkYXRhLXRhYj0iZGFzaGJvYXJkIiByb2xlPSJ0YWIiIGFyaWEtc2VsZWN0ZWQ9InRydWUiPjxpIGRhdGEtbHVjaWRlPSJsYXlvdXQtZGFzaGJvYXJkIiBzdHlsZT0id2lkdGg6MTRweDtoZWlnaHQ6MTRweDsiPjwvaT4gRGFzaGJvYXJkPC9idXR0b24+CiAgICAgIDxidXR0b24gY2xhc3M9InRhYi1idG4iIGRhdGEtdGFiPSJyZWNvcmRzIiByb2xlPSJ0YWIiIGFyaWEtc2VsZWN0ZWQ9ImZhbHNlIj48aSBkYXRhLWx1Y2lkZT0iZGF0YWJhc2UiIHN0eWxlPSJ3aWR0aDoxNHB4O2hlaWdodDoxNHB4OyI+PC9pPiBEYXRhIFJlY29yZHM8L2J1dHRvbj4KICAgICAgPGJ1dHRvbiBjbGFzcz0idGFiLWJ0biIgZGF0YS10YWI9ImZvbGxvd2VycyIgcm9sZT0idGFiIiBhcmlhLXNlbGVjdGVkPSJmYWxzZSI+PGkgZGF0YS1sdWNpZGU9InVzZXJzIiBzdHlsZT0id2lkdGg6MTRweDtoZWlnaHQ6MTRweDsiPjwvaT4gRm9sbG93ZXJzIERhdGE8L2J1dHRvbj4KICAgICAgPGJ1dHRvbiBjbGFzcz0idGFiLWJ0biIgZGF0YS10YWI9ImNvbXBhcmlzb24iIHJvbGU9InRhYiIgYXJpYS1zZWxlY3RlZD0iZmFsc2UiPjxpIGRhdGEtbHVjaWRlPSJnaXQtY29tcGFyZSIgc3R5bGU9IndpZHRoOjE0cHg7aGVpZ2h0OjE0cHg7Ij48L2k+IENvbXBhcmlzb25zPC9idXR0b24+CiAgICAgIDxidXR0b24gY2xhc3M9InRhYi1idG4iIGRhdGEtdGFiPSJ1cGxvYWQiIHJvbGU9InRhYiIgYXJpYS1zZWxlY3RlZD0iZmFsc2UiPjxpIGRhdGEtbHVjaWRlPSJ1cGxvYWQtY2xvdWQiIHN0eWxlPSJ3aWR0aDoxNHB4O2hlaWdodDoxNHB4OyI+PC9pPiBVcGxvYWQgRGF0YTwvYnV0dG9uPgogICAgICA8YnV0dG9uIGNsYXNzPSJ0YWItYnRuIiBkYXRhLXRhYj0iaGlzdG9yeSIgcm9sZT0idGFiIiBhcmlhLXNlbGVjdGVkPSJmYWxzZSI+PGkgZGF0YS1sdWNpZGU9Imhpc3RvcnkiIHN0eWxlPSJ3aWR0aDoxNHB4O2hlaWdodDoxNHB4OyI+PC9pPiBVcGxvYWQgSGlzdG9yeTwvYnV0dG9uPgogICAgPC9uYXY+CiAgICA8ZGl2IGNsYXNzPSJzaWRlYmFyLWZvb3RlciI+CiAgICAgIDxidXR0b24gY2xhc3M9ImJ0biIgaWQ9ImxvZ291dEJ0biIgdHlwZT0iYnV0dG9uIj48aSBkYXRhLWx1Y2lkZT0ibG9jayIgc3R5bGU9IndpZHRoOjE0cHg7aGVpZ2h0OjE0cHg7Ij48L2k+IExvY2s8L2J1dHRvbj4KICAgIDwvZGl2PgogIDwvYXNpZGU+CgogIDxkaXYgY2xhc3M9Im1haW4tY29sIj4KICA8c2VjdGlvbiBjbGFzcz0iZmlsdGVyLWJhciIgaWQ9ImZpbHRlckJhciI+CiAgICA8ZGl2IGNsYXNzPSJmaWx0ZXItZmllbGQiPgogICAgICA8bGFiZWwgZm9yPSJmaWx0ZXJEYXRlRnJvbSI+RnJvbTwvbGFiZWw+CiAgICAgIDxpbnB1dCB0eXBlPSJkYXRlIiBpZD0iZmlsdGVyRGF0ZUZyb20iIC8+CiAgICA8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImZpbHRlci1maWVsZCI+CiAgICAgIDxsYWJlbCBmb3I9ImZpbHRlckRhdGVUbyI+VG88L2xhYmVsPgogICAgICA8aW5wdXQgdHlwZT0iZGF0ZSIgaWQ9ImZpbHRlckRhdGVUbyIgLz4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0iZmlsdGVyLWZpZWxkIGZpbHRlci1wcmVzZXRzIiBpZD0iZmlsdGVyUHJlc2V0cyI+CiAgICAgIDxidXR0b24gdHlwZT0iYnV0dG9uIiBkYXRhLXByZXNldD0iNyI+TGFzdCA3IGRheXM8L2J1dHRvbj4KICAgICAgPGJ1dHRvbiB0eXBlPSJidXR0b24iIGRhdGEtcHJlc2V0PSIzMCI+TGFzdCAzMCBkYXlzPC9idXR0b24+CiAgICAgIDxidXR0b24gdHlwZT0iYnV0dG9uIiBkYXRhLXByZXNldD0iOTAiPkxhc3QgOTAgZGF5czwvYnV0dG9uPgogICAgICA8YnV0dG9uIHR5cGU9ImJ1dHRvbiIgZGF0YS1wcmVzZXQ9ImFsbCI+QWxsIHRpbWU8L2J1dHRvbj4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0iZmlsdGVyLWZpZWxkIj4KICAgICAgPGxhYmVsIGZvcj0iZmlsdGVyUGxhdGZvcm0iPlBsYXRmb3JtPC9sYWJlbD4KICAgICAgPHNlbGVjdCBpZD0iZmlsdGVyUGxhdGZvcm0iPjxvcHRpb24gdmFsdWU9ImFsbCI+QWxsIHBsYXRmb3Jtczwvb3B0aW9uPjwvc2VsZWN0PgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJmaWx0ZXItZmllbGQiPgogICAgICA8bGFiZWwgZm9yPSJmaWx0ZXJDYW1wYWlnbiI+Q2FtcGFpZ248L2xhYmVsPgogICAgICA8c2VsZWN0IGlkPSJmaWx0ZXJDYW1wYWlnbiI+PG9wdGlvbiB2YWx1ZT0iYWxsIj5BbGwgY2FtcGFpZ25zPC9vcHRpb24+PC9zZWxlY3Q+CiAgICA8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImZpbHRlci1maWVsZCI+CiAgICAgIDxsYWJlbCBmb3I9ImZpbHRlckNvbnRlbnRUeXBlIj5Db250ZW50IHR5cGU8L2xhYmVsPgogICAgICA8c2VsZWN0IGlkPSJmaWx0ZXJDb250ZW50VHlwZSI+PG9wdGlvbiB2YWx1ZT0iYWxsIj5BbGwgY29udGVudCB0eXBlczwvb3B0aW9uPjwvc2VsZWN0PgogICAgPC9kaXY+CiAgPC9zZWN0aW9uPgoKICA8bWFpbiBjbGFzcz0idmlldy1hcmVhIj4KICAgIDxzZWN0aW9uIGlkPSJ2aWV3LWRhc2hib2FyZCIgY2xhc3M9InZpZXcgaXMtYWN0aXZlIj48L3NlY3Rpb24+CiAgICA8c2VjdGlvbiBpZD0idmlldy1yZWNvcmRzIiBjbGFzcz0idmlldyI+PC9zZWN0aW9uPgogICAgPHNlY3Rpb24gaWQ9InZpZXctZm9sbG93ZXJzIiBjbGFzcz0idmlldyI+PC9zZWN0aW9uPgogICAgPHNlY3Rpb24gaWQ9InZpZXctY29tcGFyaXNvbiIgY2xhc3M9InZpZXciPjwvc2VjdGlvbj4KICAgIDxzZWN0aW9uIGlkPSJ2aWV3LXVwbG9hZCIgY2xhc3M9InZpZXciPjwvc2VjdGlvbj4KICAgIDxzZWN0aW9uIGlkPSJ2aWV3LWhpc3RvcnkiIGNsYXNzPSJ2aWV3Ij48L3NlY3Rpb24+CiAgPC9tYWluPgogIDwvZGl2Pgo8L2Rpdj4KCjxkaXYgaWQ9InRvYXN0Um9vdCIgY2xhc3M9InRvYXN0LXJvb3QiIGFyaWEtbGl2ZT0icG9saXRlIj48L2Rpdj4KCjxzY3JpcHQ+Ci8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBCcmFuZCBsb2dvIOKAlCBlbWJlZGRlZCBvbmNlIGhlcmUgYW5kIHdpcmVkIG9udG8gZXZlcnkgLmJyYW5kLWxvZ28KICAgPGltZz4gYW5kIHRoZSBmYXZpY29uIDxsaW5rPiBhdCBib290c3RyYXAsIHNvIHRoZSBiYXNlNjQgcGF5bG9hZAogICBhcHBlYXJzIGV4YWN0bHkgb25jZSBpbiB0aGlzIGZpbGUgaW5zdGVhZCBvZiBvbmNlIHBlciB1c2FnZSBzaXRlLgogICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KY29uc3QgTE9HT19EQVRBX1VSSSA9ICdkYXRhOmltYWdlL3BuZztiYXNlNjQsaVZCT1J3MEtHZ29BQUFBTlNVaEVVZ0FBQ0ZvQUFBZHpDQVlBQUFCbmI4bzNBQUFBQ1hCSVdYTUFBQzRqQUFBdUl3RjRwVDkyQUFBZ0FFbEVRVlI0bk96ZFRXN2JXTG9HNE9QZ3p1TmFnVjByTUdzRlVVMDhUUzRJR0pwRk80aDNRSEVIemdyYU5STUlHSlZNT2JuT0NvcFpRVHM3U0ZhZ2l4TXhzVlg1bzIzOUhKTFBBM2pZeU5GM0pIV2orZXA3RDViTFpRQUFBQUFBQUFBQTROZWVtQkVBQUFBQUFBQUFRRGVDRmdBQUFBQUFBQUFBSFFsYUFBQUFBQUFBQUFCMEpHZ0JBQUFBQUFBQUFOQ1JvQVVBQUFBQUFBQUFRRWVDRmdBQUFBQUFBQUFBSFFsYUFBQUFBQUFBQUFCMEpHZ0JBQUFBQUFBQUFOQ1JvQVVBQUFBQUFBQUFRRWVDRmdBQUFBQUFBQUFBSFFsYUFBQUFBQUFBQUFCMEpHZ0JBQUFBQUFBQUFOQ1JvQVVBQUFBQUFBQUFRRWVDRmdBQUFBQUFBQUFBSFFsYUFBQUFBQUFBQUFCMEpHZ0JBQUFBQUFBQUFOQ1JvQVVBQUFBQUFBQUFRRWVDRmdBQUFBQUFBQUFBSFFsYUFBQUFBQUFBQUFCMEpHZ0JBQUFBQUFBQUFOQ1JvQVVBQUFBQUFBQUFRRWVDRmdBQUFBQUFBQUFBSFFsYUFBQUFBQUFBQUFCMEpHZ0JBQUFBQUFBQUFOQ1JvQVVBQUFBQUFBQUFRRWVDRmdBQUFBQUFBQUFBSFFsYUFBQUFBQUFBQUFCMEpHZ0JBQUFBQUFBQUFOQ1JvQVVBQUFBQUFBQUFRRWVDRmdBQUFBQUFBQUFBSFFsYUFBQUFBQUFBQUFCMEpHZ0JBQUFBQUFBQUFOQ1JvQVVBQUFBQUFBQUFRRWVDRmdBQUFBQUFBQUFBSFFsYUFBQUFBQUFBQUFCMEpHZ0JBQUFBQUFBQUFOQ1JvQVVBQUFBQUFBQUFRRWVDRmdBQUFBQUFBQUFBSFFsYUFBQUFBQUFBQUFCMEpHZ0JBQUFBQUFBQUFOQ1JvQVVBQUFBQUFBQUFRRWVDRmdBQUFBQUFBQUFBSFFsYUFBQUFBQUFBQUFCMEpHZ0JBQUFBQUFBQUFOQ1JvQVVBQUFBQUFBQUFRRWVDRmdBQUFBQUFBQUFBSFFsYUFBQUFBQUFBQUFCMEpHZ0JBQUFBQUFBQUFOQ1JvQVVBQUFBQUFBQUFRRWVDRmdBQUFBQUFBQUFBSFFsYUFBQUFBQUFBQUFCMEpHZ0JBQUFBQUFBQUFOQ1JvQVVBQUFBQUFBQUFRRWVDRmdBQUFBQUFBQUFBSFFsYUFBQUFBQUFBQUFCMEpHZ0JBQUFBQUFBQUFOQ1JvQVVBQUFBQUFBQUFRRWVDRmdBQUFBQUFBQUFBSFFsYUFBQUFBQUFBQUFCMEpHZ0JBQUFBQUFBQUFOQ1JvQVVBQUFBQUFBQUFRRWVDRmdBQUFBQUFBQUFBSFFsYUFBQUFBQUFBQUFCMEpHZ0JBQUFBQUFBQUFOQ1JvQVVBQUFBQUFBQUFRRWVDRmdBQUFBQUFBQUFBSFFsYUFBQUFBQUFBQUFCMEpHZ0JBQUFBQUFBQUFOQ1JvQVVBQUFBQUFBQUFRRWVDRmdBQUFBQUFBQUFBSFFsYUFBQUFBQUFBQUFCMEpHZ0JBQUFBQUFBQUFOQ1JvQVVBQUFBQUFBQUFRRWVDRmdBQUFBQUFBQUFBSFFsYUFBQUFBQUFBQUFCMEpHZ0JBQUFBQUFBQUFOQ1JvQVVBQUFBQUFBQUFRRWVDRmdBQUFBQUFBQUFBSFFsYUFBQUFBQUFBQUFCMEpHZ0JBQUFBQUFBQUFOQ1JvQVVBQUFBQUFBQUFRRWVDRmdBQUFBQUFBQUFBSFFsYUFBQUFBQUFBQUFCMEpHZ0JBQUFBQUFBQUFOQ1JvQVVBQUFBQUFBQUFRRWVDRmdBQUFBQUFBQUFBSFFsYUFBQUFBQUFBQUFCMEpHZ0JBQUFBQUFBQUFOQ1JvQVVBQUFBQUFBQUFRRWVDRmdBQUFBQUFBQUFBSFFsYUFBQUFBQUQwVUY3VldWN1ZGKzRPQUFCMjYzL01Hd0FBQUFDZ2x5NWRHd0FBN0o2Z0JRQUFBQUJBeitSVlBROGhuTGczQUFEWVBkVWhBQUFBQUFBOWtsZjFjUWpoL011Sjg2cWV1RDhBQU5nZFFRc0FBQUFBZ0g2SmxTRlA3NXc0YzM4QUFMQTdnaFlBQUFBQUFEMlJWM1hjWlBIc1g2ZTEwUUlBQUhaSTBBSUFBQUFBb0FmeXFqNE1JY3kvYzFJYkxRQUFZSWNFTFFBQUFBQUErdUhmbFNGZkhMVWhEQUFBWUFjRUxRQUFBQUFBRXBkWDlZc1F3dk9mbkZKOUNBQUE3SWlnQlFBQUFBQkF3dHB0RlplL09LSDZFQUFBMkJGQkN3QUFBQUNBdE0xL1VCbHlsNDBXQUFDd0k0SVdBQUFBQUFDSnlxczZCaWhlZFRpZGpSWUFBTEFqZ2hZQUFBQUFBQW5xV0JueXhkTzhxb1V0QUFCZ0J3UXRBQUFBQUFEU2RCNUNPTHJIeVFRdEFBQmdCd1F0QUFBQUFBQVMwMjZuS081NXFvbDdCQUNBN1JPMEFBQUFBQUJJVDlmS2tMdHN0QUFBZ0IwUXRBQUFBQUFBU0VoZTFmTVF3c2tEVHZTUS93d0FBSEJQZ2hZQUFBQUFBSW5JcS9vNGhIRCswTlBrVmEwK0JBQUF0a3pRQWdBQUFBQWdIYkV5NU9ralRpTm9BUUFBV3lab0FRQUFBQUNRZ0x5cTR5YUxaNDg4U2VZdUFRQmd1d1F0QUFBQUFBRDJMSy9xd3hEQ2ZBT25FTFFBQUlBdEU3UUFBQUFBQU5pL3gxYUdmSEdVVi9XeCt3UUFnTzBSdEFBQUFBQUEyS084cWwrRUVKNXY4QVMyV2dBQXdCWUpXZ0FBQUFBQTdFbGJHWEs1NFg5ZDBBSUFBTFpJMEFJQUFBQUFZSC9tRzZvTXVXdmlQZ0VBWUhzRUxRQUFBQUFBOWlDdjZoaUllTFdGZi9tWit3UUFnTzBSdEFBQUFBQUEyTEV0VllaOGxWZTEraEFBQU5nU1FRc0FBQUFBZ04wN0R5RWNiZkZmVlI4Q0FBQmJJbWdCQUFBQUFMQkQ3YmFKWXN2L29vMFdBQUN3SllJV0FBQUFBQUM3dGJYS2tEc0VMUUFBWUVzRUxRQUFBQUNBcmNrVzVTeGJsQjc2dC9LcW5vY1FUbmJ3VDUza1ZYMjRnMzhIQUFCRzUzOWNPUUFBQUFDd2FURmdFVUtJb1lLYlpscE1EUGh6eU9KNEI1VWhkOFdBeS9VTy96MEFBQmdGUVFzQUFBQUFZR1B1QkN5T1FnaWZRZ2d6MC8xcUY1VWhkMDBFTFFBQVlQTUVMUUFBQUFDQVI4c1c1WXNRd2tVYnNQaGkza3lMRzlQOXZNM2lQSVR3Yk1mL3JNb1dBQURZQWtFTEFBQUFBT0RCc2tVNWFUZFkvRHRFOEs2WkZoY20rN1V5Wkw2SGYxcGxDd0FBYk1FVFF3VUFBQUFBN2lzR0xMSkZHV3NwL3U4SG14ck9EZldyR0RoNXVvZC85MmtiOGdBQUFEYklSZ3NBQUFBQW9MT2ZiTEM0cTJ5bVJXT3FuN2RaeEVxVjUzczhRcXdQVWQ4Q0FBQWJKR2dCQUFBQUFQeFN4NEJGOUw2WkZ2dW95VWhPWHRXSElZVExQWjhyM3R1YkJNY0RBQUM5SldnQkFBQUFBUHhRdGlpUDI0REZ5NDVUbXBubVYvdXFETGtyMi9PL0R3QUFneU5vQVFBQUFBQjg0d0VCaStpMXlwQ1Z2S29uOTV6ZHR2eHFBd2tBQUhCUEI4dmwwc3dBQUFBQWdNOGVHTENJUHNUL2VETXRQbzU5a20xbFNBeWNIQ1Z3bk9qUHE3UFQ2d1RPQVFBQWcyQ2pCUUFBQUFEd21JREZGek1oaTYvbUNZVXNRbHNmSW1nQkFBQWJJbWdCQUFBQUFDT1dMY3E0ZmVFOGhGQThZZ3F4TXNTRC9OVTJpeGhxZUpYQVVlN0swamtLQUFEMG42QUZBQUFBQUl6UW5ZQkYvSHY2aUFsOGFEYzRzSEtaNEJ3bUNad0JBQUFHNDJDNVhMcE5BQUFBQUJpSkRRWXN2dmpmWmxxODhmNzV2TTFpL3NqTklOdjAyOVhacVdvWEFBRFlBQnN0QUFBQUFHQUV0aEN3aU40S1dhemtWWDJjY01naXRQVWg2bDBBQUdBRG5oZ2lBQUFBQUF4YnRpaG5JWVNiTmdpd3FaREZweERDekZ2bnF4UXJRKzVTSHdJQUFCdGlvd1VBQUFBQURGUWJzSWgxRmtkYmVJV3pabHFvb2xodHM0aGJRcDRsY0pTZkViUUFBSUFORWJRQUFBQUFnSUhaY3NBaWVxY3laS1d0REptbmNKWmZ5SkkrSFFBQTlJanFFQUFBQUFBWWlCaXd5QlpsckFqNXp4WkRGaXBEMWwxc3NJNWxtNTYyb1JBQUFPQ1JiTFFBQUFBQWdKN0xGbVdzaGJqY1lyamlybmt6TFc2OFp6NXZzM2dSUW5pZXdGRzYrdkkrQVFBQUhrSFFBZ0FBQUFCNnFnMVl4TnFLWnp0NkJiRXk1TUw3NVhQSTRyQ0hvUVgxSVFBQXNBR0NGZ0FBQUFEUU0zc0lXQVNWSWQvb1MyWElYWk4wamdJQUFQMGxhQUVBQUFBQVBiR25nTVVYRnlwRFZ2S3FqdmZ3TW9XejNOTkpyMDRMQUFDSk9sZ3VsKzRHQUFBQUFCS1dMY3JqdHFaaUh3R0w2SDB6TGRSTzNGYUdOQ0dFb3dTTzh4Qi9YcDJkWHZmdjJBQUFrQTRiTFFBQUFBQWdVVzNBWXA3QTlnU1ZJYmZtUFE1WlJERXdJMmdCQUFDUElHZ0JBQUFBQUlsSktHQVJsYzIwYUJJNHg5NjFsU0d2ZXY0eTRtdTRTT0FjQUFEUVc0SVdBQUFBQUpDSXhBSVcwWWRtV3N3VE9FY3FoaEJRVUFFREFBQ1BKR2dCQUFBQUFIdVdMY3JEOWlGK0tnR0xMMVNHdFBLcWpvR1RreVFPOHpoSGVWVWZYcDJkZnV6eml3QUFnSDBTdEFBQUFBQ0FQV2tERnVmdDM5UEU3dUYxTXkydUV6akgzdVZWSGJkQUZBTjZTYkUrNUUwQzV3QUFnRjRTdEFBQUFBQ0FIVXM4WUJGOWFDdE1XQmxDWmNoZG1hQUZBQUE4bktBRkFBQUFBT3hJRHdJV1g4eWFhYUZhWXJYTkl0N1Zzd1NPc2ttVDRid1VBQURZUFVFTEFBQUFBTml5SGdVc29yOVVocXprVlgwODBNMGVRd3VPQUFEQVRqMHhiZ0FBQUFEWW5teFJ6a0lJVFFpaDZFSEk0bE1iQm1IbHNnZDM5aUI1VldjOVBEWUFBQ1RCUmdzQUFBQUEySUkyWUJHM0lSejFhTDRxUTFwNVZiOFkrT2FIckEwQUFRQUE5eVJvQVFBQUFBQWIxTk9BUmZTMm1SWnZFampIM3VWVmZkaHVzeGl5eVFoZUl3QUFiSVdnQlFBQUFBQnNRSThERnFHdERKa2xjSTVVRExZeTVBN1ZJUUFBOEVDQ0ZnQUFBQUR3Q05taW5MUUJpejdYVE14VmhxemtWUjN2ODNrS1o5bXlrMEcvT2dBQTJLS0Q1WEpwdmdBQUFBQndUd01KV0VUdm1ta3hTZUFjZTlkV2hqUTkzVXJ5RUg5ZW5aMWU5Ky9ZQUFDd1h6WmFBQUFBQU1BOURDaGdFVlNHZktPdjFTOFBGZC9MZ2hZQUFIQlBnaFlBQUFBQTBNSEFBaFpmeE1xUW16U09zbDl0WmNpcmtiM3NMSUV6QUFCQTd3aGFBQUFBQU1CUFpJdnlPSVJ3RVVKNFByQTV4Y3FRaXdUT2tZb3h6a0xRQWdBQUh1Q0pvUUVBQUFEQXQyTEFJbHVVbHlHRS93NHdaQkdkSjNDR0pPUlZIVGVWbkl6d3BSL2xWWDJjd0RrQUFLQlhiTFFBQUFBQWdEdmFEUmJ4d2Z2TEFjK2xiS1pGazhBNTlpNnY2cmpWb1JqeENPTHJWeDhEQUFEM0lHZ0JBQUFBQU9NSldFVHZtMmt4VCtBY3FSaDdmY29raFBBbWdYTUFBRUJ2Q0ZvQUFBQUFNR3Jab2p4c0F4YXZSaktIV1FKblNFSmUxYkUrNWRuSXg1QWxjQVlBQU9nVlFRc0FBQUFBUnFrTldKeTNmMDlITW9QWEtrTlc4cXIrc3NGazdNWWVOQUVBZ0hzVHRBQUFBQUJnVkVZYXNJZytDQmFzdVJ6Wi9mOVFYdFhaMWRtcEFBNEFBSFFrYUFFQUFBREFLSXc0WVBIRnJKa1dIOU00eW43bFZmM0NKb2Mxa3hDQ29BVUFBSFFrYUFFQUFBREE0R1dMOHJ6ZDVqRFdEUWF4TXVRNmdYUHNYVjdWaCswMkMyNWxaZ0VBQU4wSldnQUFBQUF3V05taW5MVUJpNk1SMzdMS2tIVXFRNzRsYUFFQUFQZHdzRnd1elFzQUFBQ0FRUkd3V1BPL3piUjRrOUI1OXFhdERQbDdwQy8vVjM2N09qdFZMUU1BQUIzWWFBRUFBQURBWUFoWWZPT3RrTVZLV3hseWtjSlpFaFczV3FpWEFRQ0FEZ1F0QUFBQUFPaTliRkcrYUIraUMxamMraFJDbUtWeW1BUUk0UHpjUk5BQ0FBQzZFYlFBQUFBQW9MZXlSVGxwSDZBL2M0dmZtRFhUUWhYRWFwdEZmSis4U3VBb0tjdkdQZ0FBQU9oSzBBSUFBQUNBM2hHdytLVjNLa1BXWENaMGxsUk54ajRBQUFEbzZvbEpBUUFBQU5BWE1XQ1JMY3BZYi9CL1FoWS9wRExranJ5cVZZWjA4elN2NnVNK0hCUUFBUGJOUmdzQUFBQUFrcGN0eWxocmNDRmMwY204bVJZM1BUam4xdVZWSGQ4M3hjQmY1aVpOYlA4QUFJQmZFN1FBQUFBQUlGblpvanh1SzBKZXVxVk9ZbVhJUlEvT3VTdENBL2VUOWVtd0FBQ3dMNElXQUFBQUFDUkh3T0pCVkliY2tWZjFlUWpoSkprRDlZT2dCUUFBZENCb0FRQUFBRUF5QkN3ZTVVSmx5RXBlMVYvZVI5eVBhaDRBQU9qZ1lMbGNtaE1BQUFBQWV5Vmc4V2p2bTJsaEcwRXJyK3Byb1lFSCsvUHE3UFM2cDJjSEFJQ2RzTkVDQUFBQWdMM0pGdVZoQ09HOC9YdnFKaDVNWlVncnIrcVprTVdqeE1DT29BVUFBUHlFb0FVQUFBQUFPeWRnc1ZGbE15MmFBYjJlQjh1ck9yNnZMbnA2L0ZUWWpBSUFBTDhnYUFFQUFBREF6Z2hZYk55SFpsck1CL2FhSHVQUysrclJKajAvUHdBQWJKMmdCUUFBQUFCYkoyQ3hOU3BEV25sVnZ3Z2hQRS9pTVAxMkZEZURYSjJkZmh6N0lBQUE0RWVlbUF3QUFBQUEyNVF0eWhnR3VBa2hGRUlXRy9XNm1SYlhBM285RDZZeVpPUFVod0FBd0UvWWFBRUFBQURBVnJRQmkxaHJjV1RDRy9laG5TMHIzbWViRmV0RGhIZ0FBT0FIQkMwQUFBQUEyQ2dCaTUyWU5kTkN0Y05xbTBVTUJieEs0Q2hETWhuN0FBQUE0R2NFTFFBQUFBRFlDQUdMbmZsTFpjaWF5NFRPTWhTcVF3QUE0Q2VlR0E0QUFBQUFqNUV0eWttMktKc1F3bitFTExidVV3amhmT0N2c2JPOHFnVjd0dU5wWHRYQ0ZnQUE4QU0yV2dBQUFBRHdJREZnMFc2d2VHYUNPNk15cE5VR0FZb2tEak5NY2I3TjJJY0FBQURmSTJnQkFBQUF3TDBJV096TjIyWmF2Qm5wYS84ZWxTSGJaYU1GQUFEOGdLQUZBQUFBQUowSVdPeFZyQXlaamZqMXI4bXJPdGFubkNSMHBDR2FqSDBBQUFEd0k0SVdBQUFBQVB4VXRpaVAyKzBCQWhiN00xY1pzcEpYOVhFYitHRzdCRmtBQU9BSG5oZ01BQUFBQU44VEF4Ylpvb3dCaS84S1dlelZ1MlphWEl6NDlmOWJmRTgrVGV0SXc1Ulh0YTBXQUFEd0hUWmFBQUFBQUxDbTNXQVJOd2E4TkptOVV4bHlSMXNaSXZTek8xa0k0WG9zTHhZQUFMb1N0QUFBQUFEZ013R0xKTVhLa0p1eER5R3NRaGFIS2tOMkxtNjBzRTBGQUFEK1JkQUNBQUFBWU9TeVJmbmxBZmFyc2M4aU1TcEQxcWtNMmIxc2JDOFlBQUM2T0ZndWx3WUZBQUFBTUVKdHdPSzgvZk1BT3oxL05OT2lHZnNRd21xYnhZc1F3dDhKSEdXTWZyczZPLzA0OWlFQUFNQmRObG9BQUFBQWpJeUFSUytVUWhZcmJXWElaUXBuR2FsWUgvSm03RU1BQUlDN0JDMEFBQUFBUmtMQW9qZmVOOU5pUHZZaDNESDNmdDJyVE5BQ0FBRFdDVm9BQUFBQWpFQzJLT2NDRnIweEcvc0F2c2lyT201VGVKWEdhVVpyTXZZQkFBREF2d2xhQUFBQUFBeFl0aWhuN1VhQUkvZmNDNjlWaHF5b0RFbkdzN0VQQUFBQS91MkppUUFBQUFBTVR3eFlaSXZ5Sm9Ud0h5R0wzdmpRaG1KWU9mZmVUVU5lMWRuWVp3QUFBSGZaYUFFQUFBQXdJRFpZOU5xc21SWWZ4ejZFY1B0Z3YwamdLS3pFKzdCcEJRQUFXb0lXQUFBQUFBTWdZTkY3c1RMa2V1eER1RU5sU0ZvbTdnUUFBRzRKV2dBQUFBRDBXTFlvSjIzQTRwbDc3QzJWSVhma1ZSMW5jWkxNZ1FqdFJnc0FBS0FsYUFFQUFBRFFRd0lXZzNLdU1tUWxyK3JqT0k4VXpzSWF3UmNBQUxqamlXRUFBQUFBOUVjTVdHU0xNbFpNL0orUXhTQzhiYWJGbTdFUDRZNVlUL0UwbWRQd1ZWN1ZFOU1BQUlBVkd5MEFBQUFBZXNBR2kwSDZGRUtZalgwSVgrUlZmZTc5bmJUNEhYUTk5aUVBQUVBUXRBQUFBQUJJVzdZb2o5dUF4VXRYTlRnemxTRXJlVlVmdHU5ejBwVzVHd0FBV0JHMEFBQUFBRWlRZ01YZ3ZWTVpza1psU1BwVWh3QUFRT3VKUVFBQUFBQ2tJd1lzc2tVWkh6ci9WOGhpc0ZTRzNKRlg5WXNRd3ZOa0RzU1BQTTJyK3RoMEFBREFSZ3NBQUFDQUpOaGdNU3J6WmxyY2pIMEk0Yll5NURLQm85Qk5yQS94M2dVQVlQUUVMUUFBQUFEMktGdVU4VUh6ZVFpaGNBK2pFQ3RETHNZK2hEdm1La042SmRhSHFMd0JBR0QwQkMwQUFBQUE5dUJPd09MY2crYlJVQmx5UjE3VjhhSDlxMlFPUkJlWktRRUFnS0FGQUFBQXdFNEpXSXphaGNxUUZaVWh2ZlZzN0FNQUFJQWdhQUVBQUFDd0d3SVdvL2UrbVJienNRL2hqdmc1T0VybU5IU1dWM1YyZFhiYW1CZ0FBR1AyeE8wREFBQUFiRmUyS0dOZFJOeGtVQWhaakpiS2tGWjhVTjkrRnVpbmlYc0RBR0RzYkxRQUFBQUEySkkyWURIM3kvM1JLNXRwWVFQQUxaVWgvWmFOZlFBQUFDQm9BUUFBQUxCaEFoYmM4VUZseUsyOHF1TXNUbEk1RHc4aWFBRUF3T2dkTEpmTHNjOEFBQUFBWUNNRUxQaU9QNXRwY1cwd24wTVd4eUdFUm4zT0lQeDJkWGI2Y2V4REFBQmd2R3kwQUFBQUFIaWtiRkZPMmpvRUFRdnVlaTFrc2VaU3lHSXc0bFlMNzIwQUFFWkwwQUlBQUFEZ2dkcUFSZHhnOGN3TStaY1A3WHVEMVRhTGM1K1RRWmtJV2dBQU1HYUNGZ0FBQUFEM0pHQkJCN05tV3FoV3VLME1FVG9abHNuWUJ3QUF3TGdKV2dBQUFBQjBKR0JCUjMrcERGbHpvVEprY0xLeER3QUFnSEU3V0M2WFk1OEJBQUFBd0U5bGl6SnJIeFlMV1BBcm4wSUl4N1pack9SVi9TS0U4SGNLWjJIamZyODZPNzB4VmdBQXhzaEdDd0FBQUlBZnlCYmxsOHFEbDJaRVJ5cERXbmxWSDRZUUxwTTRETnN3Y2I4QUFJeVZvQVVBQUFEQXZ3aFk4RUJ2bTJueHh2QytVaGt5Yk9wREFBQVlMVUVMQUFBQWdKYUFCWThRSzBObUJyaVNWL1hFNTJqd0JDMEFBQmd0UVFzQUFBQmc5TEpGZWRqKyt0NkRZUjVxcmpKa1JXWElhRHdiK3dBQUFCaXZKKzRlQUFBQUdLc1lzTWdXWmR4Z2NTTmt3U084YTZiRmhRRitGVDlUUjRtY2hTMXFONWNBQU1EbzJHZ0JBQUFBakU2N3dlSzgvWHZxSGNBanFBeTVJNi9xV0NmeEtwa0RzVzN4dnE5TkdRQ0FzUkcwQUFBQUFFWkR3SUl0aUpVaE53YjdsY3FRY2NuR1BnQUFBTVpKMEFJQUFBQVlQQUVMdGtSbHlCMTVWY2ZLa0pOa0RzUXVxQTRCQUdDVW5yaDJBQUFBWU1peVJSbHJIWm9RUWlGa3dZYWRHK2hLWHRYSDdXZU1jVG5LcS9yUW5RTUFNRFkyV2dBQUFBQ0QxQVlzNGkvc2o5d3dXMUEyMDZJeDJLOVVob3hYM0dyeFp1eERBQUJnWEFRdEFBQUFnRUVSc0dBSDNqZlRZbTdRSzNsVng4MGV6MUk0QzN1UkNWb0FBREEyZ2hZQUFBREFJQWhZc0VNencxNXBLME9FVHNadE12WUJBQUF3UG9JV0FBQUFRSzlsaTNMU1B1ajFpM3AyNGJYS2tEVVhJWVNuQ1oySDNjdk1IQUNBc1RsWUxwY3VIUUFBQU9nZEFRdjI0RU44NnpYVDRxUGhmOTVtOFNLRThIY0NSMkgvL3JnNk94VkFBZ0JnTkd5MEFBQUFBSHBGd0lJOW1nbFpyT1JWZlJoQ3VFemhMQ1FoYnJVUXRBQUFZRFFFTFFBQUFJQmVFTEJnejJKbHlMVkwrRXBsQ0hlcER3RUFZRlFFTFFBQUFJQ2taWXZ5dUgybys5eE5zU2NmMnBBUHEyMFdNZlQwMGl5NFkySVlBQUNNaWFBRkFBQUFrS1EyWURIM1FKY0VuS3NNV1ZFWndnK2NHQXdBQUdQeXhHMERBQUFBS1lrQmkyeFJ4Z2U1L3hXeUlBRnZtMm54eGtWOEZjTlBSNG1jaFlTMG0wNEFBR0FVYkxRQUFBQUFrbUNEQlFuNkZFS1l1WmlWdktxekVNS3JGTTVDa3VMNzQ5clZBQUF3Qm9JV0FBQUF3RjVsaS9Ld0RWaDRnRXRxWmlwRDFxZ000V2ZpUm9zTEV3SUFZQXdFTFFBQUFJQzlhQU1XNSszZlU3ZEFZdDZwRExtVlYzVU1RNTJrY2g2U2xMa1dBQURHNG1DNVhMcHNBQUFBWUdjRUxPaUJXQm1TTmRQaXhtVjlyUXo1SjRHamtMN2ZyODVPZlc0QUFCZzhHeTBBQUFDQW5SQ3dvRWZtUWhacjFFSFFWUXpsK093QUFEQjRUMXd4QUFBQXNHM1pvanh2SDc0VlFoWWtMbGFHQ0JhMDhxcU9uOTFuU1J5R1BsQWZBZ0RBS05ob0FRQUFBR3hOdGlobmNUdEFDT0hJbE9tQldCa3ljMUVyZVZVZnQ1OWY2R3BpVWdBQWpJR2dCUUFBQUxCeEFoYjAxSVhLa0RXWE50QndUN2FmQUFBd0NnZkw1ZEpOQXdBQUFCc2hZRUdQdlcrbWhkcURWbDdWTDBJSWZ5ZHhHUHJtajZ1ejA4YXRBUUF3WkRaYUFBQUFBSStXTGNyNFVQWkN3SUllVXhuU3lxdjZzTjFtQVE4UkEwdUNGZ0FBREpxZ0JRQUFBUEJnMmFLY3RCc3NySXVuejhwbVduZ3dmRXRsQ0k4eEVkUUJBR0RvQkMwQUFBQ0FleE93WUVBK05OTmk3a0pYOHFxT24rM25LWnlGM2xMQkF3REE0QjBzbDB1M0RBQUFBSFFpWU1FQS9kbE1pMnNYKzdVeXBGRUJ4QWI4ZG5WMit0RWdBUUFZS2hzdEFBQUFnRi9LRm1YOGhmS0ZnQVVEODFySVlzMWN5SUlOaWYrZDRiTUZBTUJnQ1ZvQUFBQUFQNVF0eXVQMjRldExVMkpnUHJUdmJXNHJRMTZaQlJzeUViUUFBR0RJQkMwQUFBQ0Fid2hZTUFLelpscW9Ocmgxa2NwQkdJVE1OUUlBTUdTQ0ZnQUFBTUJYQWhhTXhGOHFRMjdsVlIwLzh5ZXBuSWRCbUxoR0FBQ0c3R0M1WExwZ0FBQUFHRGtCQzBia1V3amgyRGFMbGJ5cTQrYUJmMUk0QzRQeis5WFo2WTFyQlFCZ2lHeTBBQUFBZ0JITEZ1VmhDT0c4L1h2cXZjQUlxQXhacHpLRWJZa2hIa0VMQUFBR1NkQUNBQUFBUmtqQWdwRjYyMHlMTnk1L0phL3ErUGwvbHNKWkdLUllIK0x6QmdEQUlBbGFBQUFBd0lnSVdEQmlzVEprNWcyd2tsZjFsN29nMkpiTVpBRUFHQ3BCQ3dBQUFCZ0JBUXNJYzVVaGF5NTlGN0JsdHFVQUFEQllCOHZsMHUwQ0FBREFnR1dMTXY2Sy84SkRWVWJzWFRNdEp0NEFLM2xWdndnaC9KM0NXUmk4UDY3T1RodlhEQURBME5ob0FRQUFBQVBWQml4aU5jQ1JPMmJFVkliY2tWZjFZYnZOQW5ZaEJwd0VMUUFBR0J4QkN3QUFBQmdZQVF0WUV5dERib3prSzVVaDdGSm0yZ0FBREpHZ0JRQUFBQXlFZ0FWOEkxYUdYQmpMU2w3VmNidkE4eFRPd21pbzdBRUFZSkFPbHN1bG13VUFBSUFleXhabGZKQVZIeWFmdUVkWTgwY3pMZFFXM0ZhR05JSlk3TUZ2VjJlbkh3MGVBSUFoc2RFQ0FBQUFlcW9OV01RTkZzL2NJWHlqRkxKWVk5c04reExyUTY1Tkh3Q0FJUkcwQUFBQWdKNFJzSUJmZXQ5TWk3a3hyYlNWSWE5U09BdWpOQkcwQUFCZ2FBUXRBQUFBb0NjRUxLQ3ptVkd0dVV6b0xJelB4SjBEQURBMGdoWUFBQUNRdUd4UkhyY1BTZ1VzNE5kZXF3eTVsVmUxeWhEMkxYTURBQUFNemNGeXVYU3BBQUFBa0tBMllCRWZrcjUwUDlESmgvalJhYWJGUitQNkhMS0lEN2ovU2VBbzhQdlYyZW5ONktjQUFNQmcyR2dCQUFBQWlSR3dnQWViQ1Ztc1VSbENLaWJlandBQURJbWdCUUFBQUNSQ3dBSWVKVmFHWEJ2aFNsN1Y1eUdFa3hUT0F1cERBQUFZR2tFTEFBQUEyTE5zVVI2MkFZdFg3Z0llNUVQN0dXSVZzamcyRHhJamFBRUF3S0FJV2dBQUFNQ2V0QUdMOC9idnFYdUFCenRYR2JMbTBuY0tpWG5tUWdBQUdKS0Q1WExwUWdFQUFHQ0hCQ3hnbzk0MjArS0ZrYTdrVlQwTElmd25oYlBBdi94NWRYYXEzZ2NBZ0VHdzBRSUFBQUIyUk1BQ051NVRDR0ZtckN0NVZjZnZtSXNVemdMZkVldERCQzBBQUJpRUo2NFJBQUFBdGk5YmxQTVF3azBJb1JDeWdJMlpxUXhab3pLRWxFM2NEZ0FBUTJHakJRQUFBR3hSdGlqanIrMWp5T0xJbkdHajNqWFQ0bzJScnVSVkhldFRucWR3RnZpQnpHQUFBQmlLZytWeTZUSUJBQUJnd3dRc1lLdGlaVWpXVElzYlkvNWFHZEw0dnFFSGZyczZPN1dGQmdDQTNyUFJBZ0FBQURaSXdBSjJZaTVrc2NaM0RuMFI2ME5zb2dFQW9QY0VMUUFBQUdBREJDeGdaMkpseUlWeHIrUlZIUjljdjByaExOQkJKbWdCQU1BUUNGb0FBQURBSTJTTGN0SUdMSjZaSTJ4ZHJBeVpHZk9heTRUT0FyOHlNU0VBQUlaQTBBSUFBQUFlUU1BQzl1SkNaY2l0dktwdDBhRnZNamNHQU1BUUhDeVhTeGNKQUFBQUhRbFl3TjY4YjZhRmg3U3R2S3JqTFA1SjRqQndQMzljblowMlpnWUFRSi9aYUFFQUFBQWRDRmpBM3FrTVdhY3loTDZLSVNGQkN3QUFlazNRQWdBQUFINGlXNVRIYmNEaXBUbkIzcFROdFBCZ3RwVlg5WGtJNFNTSnc4RDkyVXdEQUVEdkNWb0FBQURBZHdoWVFESStOTk5pN2pwVzhxcis4dDBFZlRWeGN3QUE5SjJnQlFBQUFOd2hZQUhKVVJteUxsYUdQRTNwUUhCUHRyRUFBTkI3Qjh2bDBpMENBQUF3ZWdJV2tLVFh6YlE0ZHpVcmVWWEgwTWwvVWpnTFBOS2ZWMmVuMTRZSUFFQmYyV2dCQUFEQXFHV0w4akNFRUIva0ZtT2ZCU1RtZzRxTVczbFZ4KytxaTFUT0E0OFU2ME1FTFFBQTZDMUJDd0FBQU1NOWhzc0FBQ0FBU1VSQlZFYnBUc0RpM0JwK1NOS3NtUllmWGMxWEtrTVlrc3h0QWdEUVo0SVdBQUFBaklxQUJmVENYODIwOEd2M1ZsN1ZMMElJejVNNERHeUdvQVVBQUwxMnNGd3UzU0FBQUFDREoyQUJ2ZkVwaEhCc204VktXeGx5NDN1TEFmcjk2dXoweHNVQ0FOQkhObG9BQUFBd2VObWluSVVRTGp5b2hGNVFHYkp1N3J1TGdjcmFFQkVBQVBTT29BVUFBQUNEMVFZczRrUEtJN2NNdmZDMm1SWnZYTlZLWHRXVEVNS3JGTTRDV3hDREZqN3ZBQUQwa3FBRkFBQUFneU5nQWIwVUswTm1ybTZsclF5NVRPRXNzQ1VUZ3dVQW9LOEVMUUFBQUJnTUFRdm90Ym5La0RYbnZzc1l1R2N1R0FDQXZqcFlMcGN1RHdBQWdGN0xGdVdrL2VXM2g1TFFUKythYWVIWDdhMjhxbU9sd2o5SkhBYTI2NCtyczlQR2pBRUE2QnNiTFFBQUFPaXRObUF4OTZ0WTZEV1ZJZDlTR2NKWXhGQ1JvQVVBQUwwamFBRUFBRUR2Q0ZqQW9NVEtrQnRYdXBKWGRmeHVPMG5oTExBREU4RWlBQUQ2U05BQ0FBQ0EzaEN3Z01HSmxTRVhyblVscityakVNSjVDbWVCSGNrTUdnQ0FQaEswQUFBQUlIblpvb3dQWWk0RUxHQndoQXJXeFYvMlAwM3BRTEJsSjNsVkgxNmRuWDQwYUFBQStrVFFBZ0FBZ0dSbGkvSzQzV0R4MGkzQjRKVE50R2hjNjBwZTFlZkNaSXhVREZOZXUzd0FBUHBFMEFJQUFJRGtDRmpBNEwxdnBzWGNOYS9FWC9TMzMza3dSaE5CQ3dBQStrYlFBZ0FBZ0dRSVdNQm96RnoxR3BVaGpGbm05Z0VBNkJ0QkN3QUFBUFl1VzVUeDE5d1hBaFl3Q3E5Vmh0ektxL3BGQ09GNUt1ZUJQWmdZT2dBQWZYT3dYQzVkR2dBQUFIdlJCaXpPMnorLzVvYmgreEEvK3MyMCtPaXV2MWFHM1BqK2cvRDcxZG5walRFQUFOQVhObG9BQUFDd2N3SVdNRm96SVlzMWM5K0I4Rm5XaG80QUFLQVhCQzBBQUFEWUdRRUxHTFZZR1hJOTlpRjhrVmQxckV0NGxjWnBZTy9pNStHTmF3QUFvQzhFTFFBQUFOZzZBUXNZdlEvdDlnWnVLME11elFLK3lvd0NBSUErRWJRQUFBQmdxN0pGT1dzZnNCNlpOSXpXdWNxUU5lZStFMkhOTStNQUFLQlBEcGJMcFFzREFBQmc0d1FzZ05iYlpscThNSXlWdktyakwvZi9TZUVza0pnL3I4NU8xUXNCQU5BTE5sb0FBQUN3VVFJV3dCMmZRZ2d6QTFtak1nUytMNGFRQkMwQUFPZ0ZRUXNBQUFBMlFzQUMrSTZaeXBCYmVWWEg3OGlUVk00RGljbGNDQUFBZlNGb0FRQUF3S05raTNMU0JpejBxd04zdld1bXhSc1RXY21yK2ppRWNKN0NXU0JSRXhjREFFQmZDRm9BQUFEd0lBSVd3RStvRFBsV3JBeDVtdHFoSUNGSGVWVWZYcDJkMm9JREFFRHlCQzBBQUFDNEZ3RUxvSU41TXkxdURHb2xyK3B6MzVuUVNhd1B1VFlxQUFCU0oyZ0JBQUJBSndJV1FFZXhNdVRDc0ZiaUwvVGI3MDdnMXlhQ0ZnQUE5SUdnQlFBQUFEK1ZMY3JqRUVKOGFQcmNwSUJmVUJueUxaVWgwTjNFckFBQTZBTkJDd0FBQUw2ckRWakVYMkcvTkNHZ293dVZJYmZ5cW40aHBBYjNraGtYQUFCOWNMQmNMbDBVQUFBQVh3bFlBQS8wdnBrV0hwSzIyc3FRRzlzczRONSt2em83RmRnQ0FDQnBObG9BQUFEd21ZQUY4RWdxUTlaZENGbkFnMHpheWgwQUFFaVdvQVVBQU1ESVpZdnlzQTFZdkJyN0xJQUhLNXRwMFJqZlNsN1ZFNkUxZURDYmNRQUFTSjZnQlFBQXdFaTFBWXZ6OXMrdnJvR0grdEJNaTducHJiU1ZJWDZORHc4M01Uc0FBRkluYUFFQUFEQXlBaGJBaHFrTVdSZERKMGNwSFFoNjVzU0ZBUUNRdW9QbGN1bVNBQUFBUmtEQUF0aUMxODIwT0RmWWxieXFZK1hCUHltY0JYcnV6NnV6MDJ1WENBQkFxcDY0R1FBQWdPSExGbVY4RUhvVFFpaUVMSUFOK2RCdWIrQ1d5aERZak13Y0FRQkltZW9RQUFDQUFjc1c1Y3dhZTJCTFpzMjArR2k0SzNsVnoxVWV3TVpNUWdnWHhna0FRS29FTFFBQUFBWkl3QUxZc3IrYWFXR3RmeXV2NnVOMll4Q3dHVFphQUFDUU5FRUxBQUNBQVJHd0FIYmdVd2poM0tEWHFBeUJ6VHJLcS9ydzZ1elUxaHdBQUpJa2FBRUFBREFBMmFKODBhN1lGckFBdGsxbHlCMTVWY2ZReWJOa0RnVERFZXREM3JoUEFBQlNKR2dCQUFEUVk5bWluTFFiTER6a0EzYmhiVE10UFBoc3RaVWg4eVFPQThPVENWb0FBSkFxUVFzQUFJQWVFckFBOWlCV2hzd01mazNjSlBRMG9mUEFrRXpjSmdBQXFSSzBBQUFBNkJFQkMyQ1A1aXBEYnVWVkhTdWJucWR5SGhpZ3pLVUNBSkNxZytWeTZYSUFBQUFTbHkzS3JQM2x0SUFGc0Evdm1tbmgxK1d0dktvUFF3ZzN0bG5BMXYxeGRYYmFHRE1BQUtteDBRSUFBQ0JoMmFJOGJqZFl2SFJQd0o2b0RQbVd5aERZalJnMEZiUUFBQ0E1Z2hZQUFBQUpFckFBRWhJclEyNWN5RXBlMVJQZnpiQXo4Zk4yYWR3QUFLUkcwQUlBQUNBaEFoWkFZbUpseUlWTFdXa3JRenowaGQzSnpCb0FnQlFKV2dBQUFDUkF3QUpJMUxtTFdSTy9wNDhTT2c4TTNZa2JCZ0FnUlFmTDVkTEZBQUFBN0VtMktBL2JCNW5uK3Y2QnhKVE50Smk3bEpXOHF1TXY2LzlKNFN3d01uOWVuWjFldTNRQUFGSmlvd1VBQU1BZUNGZ0FpWHN2WlBFTmxTR3dINU1RZ3FBRkFBQkpFYlFBQUFEWUlRRUxvQ2RtTHVwV1h0VnpGUWF3TjVuUkF3Q1FHa0VMQUFDQUhSQ3dBSHJrZFRNdEdoZTJrbGYxY1FpaFNPRXNNRktDRmdBQUpPZUpLd0VBQU5pdWJGSEdYNGJmdEEvcWhDeUFsSDBJSWFnTVdhY3lCUGJycUEwOEFRQkFNbXkwQUFBQTJKSTJZQkVmV0I2Wk1kQVRzMlphZkhSWkszbFZ4eTFFejFJNEM0eGMxb1pXQVFBZ0NZSVdBQUFBR3laZ0FmUlVyQXk1ZG5rcjdTL29iZmVBTk1TZ3hSdDNBUUJBS2dRdEFBQUFOa1RBQXVneGxTSGZ1bFQzQk1tWXVBb0FBRklpYUFFQUFQQkkyYUtNLytmL1JRamh4Q3lCbmpwWEdYSXJyK29YS2tNZ0tUNlBBQUFrNVdDNVhMb1JBQUNBQjJnREZuUC81ei9RYzIrYmFmSENKYTdrVlgwWVFyaXh6UUtTODhmVjJXbmpXZ0FBU0lHTkZnQUFBUGNrWUFFTXlLY1F3c3lGcmxFWkFtbUsvL3RMMEFJQWdDUUlXZ0FBQUhRa1lBRU0wRXhseUsyOHF1UDMvUE5VemdPc3lZd0RBSUJVQ0ZvQUFBRDhRcllvajl0Zk9BdFlBRVB5cnBrV2I5em9TbHNaY3BuQ1dZRHZFclFBQUNBWlQxd0ZBQURBOThXQVJiWW80ME8zL3dwWkFBT2pNdVJiY1dQUlVXcUhBcjQ2YVFOUkFBQ3dkelphQUFBQS9FdTd3U0krY0h0cE5zQkF6WnRwY2VOeVY5cktrRmNwbkFYNHFialY0dHFJQUFEWU4wRUxBQUNBbG9BRk1CS3hNdVRDWmE4eEQraUhpYUFGQUFBcEVMUUFBQUJHTDF1VWgyM0F3cStaZ2FGVEdmSXZlVlhINy8rVHBBNEYvRWhtTWdBQXBFRFFBZ0FBR0swMllISGUvajMxVGdCRzRFSmx5SzI4cXVORDJ5S1Y4d0MvTkRFaUFBQlNjTEJjTGwwRUFBQXdLZ0lXd0VpOWI2YUZYNFBma1ZkMXJDQjRsc3lCZ0M1K3Z6bzdGUmdEQUdDdmJMUUFBQUJHUThBQ0dEbVZJWGZrVlgwdVpBRzlGTGRhWExvNkFBRDJTZEFDQUFBWWhXeFJ6Z1VzZ0JFcm0yblJlQU9zNUZWOUhFS1lwM0FXNE41czVnRUFZTzhFTFFBQWdFSExGdVdzZlpoMjVLYUJrZnJRVEF1aGduV1hnbmZRVzRJV0FBRHNuYUFGQUFBd1NBSVdBRitwRExranIrb1hLa09nMTN4K0FRRFl1NFBsY3VrV0FBQ0F3UkN3QUZqenVwa1c1MGF5a2xmMVlRamh4allMNkwwL3I4NU9yMTBqQUFEN1lxTUZBQUF3Q0FJV0FOLzQwSDR2Y2t0bENBeERyQThSdEFBQVlHOEVMUUFBZ0Y3TEZ1V2tmWkJvalRUQXVsa3pMVDZheVVwZTFmRy9MNTZuY0JiZzBUSWpCQUJnbndRdEFBQ0FYaEt3QVBpcHY1cHA0ZGZlcmJZeTVES0p3d0NiTURGRkFBRDJTZEFDQUFEb0ZRRUxnRi82RkVJNE42WTFxcVZnV0k1aWdPcnE3TlRXSGdBQTlrTFFBZ0FBNkFVQkM0RE9WSWJjMFZhR3ZFcm1RTUNteFBvUW0zc0FBTmdMUVFzQUFDQnAyYUk4YmdNV0w5MFV3Qys5YmFiRkcyTmFjNUhRV1lETm1RaGE4UC9zM1U5dTNOYldMK3p0NFBTbE93THBqTUE4STFDbHc2NzlnUUNobnZXT0lMb2pVTlVJampLQzErNFZDQkRYNmJMelNpTzQ1UkZjYVFiV0NQUmhKL1NKNC9pUC9sVFY1dDU4SHNETkFPUmFGWkZWL0hFdEFJQlVCQzBBQUlCSkVyQUFlTFM0TXVSTTJmN1VkRU84anJ5Y3l2RUFXN1ZRVGdBQVVoRzBBQUFBSmtYQUF1REpsbGFHL0tucGhyaFc0R0lxeHdOc1hhV2tBQUNrOHVMKy9sN3hBUUNBNUFRc0FKN2xlbk42NGUzdXp6VGRzREhOQW9yM3I3NnROOW9NQU1DK21XZ0JBQUFrVmExWGh5R0VjMjhkQXp5WmxTRmZhTHJoWE1nQ1ppRk90UkMwQUFCZzd3UXRBQUNBSkQ0TFdNUi9CN29BOEdSeFpjaU44djJoNllaUEU1S0E4bGtmQWdCQUVvSVdBQURBWGdsWUFHeFZYQmx5cWFSLzhkYjFCV2JEeWlRQUFKSVF0QUFBQVBaQ3dBSmdKODZWOVU5Tk44UVZLaWRUT1I1ZzU2d0lBZ0FnaVorVUhRQUEyTFZxdllvUEF1Tlkrd3NoQzRDdFdXMU9MemJLK1llbUcyS2d6M1FQbUptbUcweTFBQUJnNzB5MEFBQUFkcVphcjg3R1BmbEhxZ3l3VlI4MnB4ZExKZjBMSzBOZ25xb1F3cFhlQXdDd1Q0SVdBQURBMWdsWUFPemNtUkwvcWVtRzF5R0VWMU01SG1DdkZxYlpBQUN3YjRJV0FBREExZ2hZQU96RnIxYUcvTW5LRUppOWF1NEZBQUJnL3dRdEFBQ0FaNnZXcThVNHNsM0FBbUMzYnNkQUczOFM4SU41TzRxQnE3NnRQODY5RUFBQTdJK2dCUUFBOEdSandDSSs0RHBSUllDOU9OdWNYbmlZT0dxNklWNkhmcG5Fd1FBcHhiOEY3M1VBQUlCOUViUUFBQUFlVGNBQ0lJbTRNdVJLNmYvaTdZU09CVWluRXJRQUFHQ2ZCQzBBQUlBSEU3QUFTTWJLa0M4MDNXQmxDUERKUWlVQUFOZ25RUXNBQU9DSHF2VXF2aVY0S1dBQmtNeTVsU0YvYXJvaFhwY3Vwbkk4UUhMdVVRRUEyS3VmbEJzQUFQaVdhcjA2cnRhck9KYjkvL29CR3lDWjN6YW5GMGJpLzVXVkljQmZqQUVzQUFEWUN4TXRBQUNBdjRrQmkzRkUvUnZWQVVqcUxvUndwZ1YvYXJyaFBJVHdjaXJIQTB4R0RGcHN0QU1BZ0gwUXRBQUFBUDVEd0FKZ2NzNnNEUGxUMHcyZnJsTUFYMXFZZGdNQXdMNElXZ0FBQURGZ2NSaEN1QlN3QUppVWF5dEQvaVkrUkQyWTJERUIwMkIxQ0FBQWV5Tm9BUUFBTXpZR0xNN0hmeDVjQVV5SGxTRmZhTG9oMXVOa1VnY0ZUSW1WUWdBQTdNMkwrL3Q3MVFZQWdKa1JzQUNZdlArOU9iMjQxS1kvTk4wUXIxczNybG5BRC96Y3QvV1ZJZ0VBc0dzbVdnQUF3SXdJV0FCazRWckk0bStzREFFZVloRkNFTFFBQUdEbkJDMEFBR0FHQkN3QXNtRmx5QmVhYm5nZFFuZzFxWU1DcHFyU0dRQUE5a0hRQWdBQUNsZXRWL0dCM1RLRWNLVFhBSk4zdVRtOXVOR21QNHdyUTB6M0FCNUswQUlBZ0wxNGNYOS9yOUlBQUZBZ0FRdUE3SHpZbkY1NFNQaVpwaHRpeU9LWHlSd1FrSU4vOW0wdHNBWUF3RTZaYUFFQUFJVVJzQURJbHBVaG4ybTZZU0ZrQVR4QkRLd0pXZ0FBc0ZPQ0ZnQUFVQWdCQzRDc3JUYW5GeHN0L01PNE11VHRGSTRGeUU0TWFiM1hOZ0FBZGtuUUFnQUFNbGV0VjR0eGYvMUx2UVRJMHUzbTlHS3BkWDl4TGpnSVBKRVZUQUFBN0p5Z0JRQUFaR29NV01RSGN5ZDZDSkExSzBNKzAzUkRmRWg2TVprREFuTGozaGdBZ0ozN1NZa0JBQ0F2TVdCUnJWZFhJWVQvOFVNeVFQWiszWnhlWEdualgxZ1pBanpMR05nQ0FJQ2RNZEVDQUFBeVlZSUZRSEZ1eDcvcmpKcHVXRnFGQld4QnZHL2VLQ1FBQUxzaWFBRUFBQk5YclZmSElZVExFTUlydlFJb3l0bm05T0tqbHY2aDZZWjR2VHVmd3JFQTJUUFJBZ0NBblJLMEFBQ0FpUm9ERnZITjNqZDZCRkNjZDFhRy9FMWNHWEl3c1dNQzhpUm9BUURBVHIyNHY3OVhZUUFBbUJBQkM0RGkzWVVRamsyeitGUFREWEdTeGIrbmNqeEFFZjVYMzliK3pnSUFzQk1tV2dBQXdFUUlXQURNd3EyVklYL1ZkTVBoZVAwRDJLWTQxY0xrSUFBQWRrTFFBZ0FBRXF2V3EwOFBtSDdSQzRCaXhZREZjbk42OFZhTC84YktFR0FYRm9JV0FBRHNpcUFGQUFBa01nWXN6c2QvSGpBQmxFbkE0anVhYm5nZFFuZzEyUU1FY2xicEhnQUF1eUpvQVFBQWV5WmdBVEFMQWhZL01LNE1VUjlnVnhZcUN3REFyZ2hhQUFEQW5naFlBTXlDZ01YRExWMFBnUjA2YUxyaHVHL3JHMFVHQUdEYkJDMEFBR0FQcXZYcTNBTWxnS0lKV0R4QzB3M3hUZk5mc2psZ0lGY0xrM01BQU5nRlFRc0FBTmloYXIwNkd3TVdSK29NVUNRQmkwZXlNZ1RZbzBxeEFRRFlCVUVMQUFEWUFRRUxnT0lKV0R6ZHVlc2pzQ2VDRmdBQTdJU2dCUUFBYkpHQUJVRHhCQ3llb2VtRytORHpJdHNUQUhKem9tTUFBT3lDb0FVQUFHeEJ0VjY5RGlGY0NsZ0FGRXZBWWp2VUQ5aXJwaHNXZlZ0ZnFUb0FBTnNrYUFFQUFNOVFyVmVMY1lLRnQrVUF5aVJnc1NWTk44U1ZJUytMT0JrZ0ozR1NqcUFGQUFCYkpXZ0JBQUJQSUdBQlVEd0JpeTFxdXVGd3ZHNEM3RnVsNGdBQWJKdWdCUUFBUElLQUJVRHhCQ3gySTY3WE9panh4SURKVzJnUkFBRGI5dUwrL2w1UkFRRGdCd1FzQUlyM0lZWUJCQ3kycittRytEYjUveTN0dklDcy9LKytyVDlxR1FBQTIyS2lCUUFBZkVlMVhoMlBBWXMzNmdSUXBPdHhnb1g5L2J0eldlcUpBZG1Jb2VuMzJnVUF3TFlJV2dBQXdGY0lXQUFVVDhCaUQ1cHVlRzBhRkRBQmxhQUZBQURiSkdnQkFBQ2ZFYkFBS0o2QXhYNlpaZ0ZNd1VJWEFBRFlKa0VMQUFBUXNBQ1lBd0dMUFd1NklWNVhqMloxMHNCVVZUb0RBTUEydmJpL3YxZFFBQUJtcTFxdkRrTUk1K08vQTU4RWdPSUlXQ1RRZEVNTU1HNWNXNEVKK1ZmZjFoc05BUUJnRzB5MEFBQmdsZ1FzQUlvbllKSFcwdlVWbUpocURJQUJBTUN6Q1ZvQUFEQXJBaFlBeFJPd1NLenBob1ZWWE1BRVdSOENBTURXQ0ZvQUFEQUxBaFlBeFJPd21JN2wzQXNBVE5KQ1d3QUEyQlpCQ3dBQWlsZXRWMmNoaEVzQkM0QWlDVmhNU05NTjhacDdNdmM2QUpQMFVsc0FBTmlXRi9mMzk0b0pBRUNSeG9CRmZLdjJTSWNCaWlOZ01URk5OOFRwVVJ2WFhXRENmdTdiMm5VREFJQm5NOUVDQUlEaUNGZ0FGRTNBWXJyT1hYdUJpYXRDQ0s0ZkFBQThtNkFGQUFERkVMQUFLSnFBeFlRMTNYQWNRcmlZZXgyQXlWdU1Ld1VCQU9CWkJDMEFBTWhldFY1OStzSFUzbVdBOGdoWTVNR0RTeUFIbFM0QkFMQU5MKzd2N3hVU0FJQXNqUUdMT01IaVJBY0JpaU5na1ltbUcrTDErSC9tWGdjZ0cvL3MyL3BHdXdBQWVBNFRMUUFBeUk2QUJVRFJCQ3p5WTVvRmtKTTQxVUxRQWdDQVp4RzBBQUFnR3dJV0FFVVRzTWhRMHcxblZuY0JtWWxCaS9lYUJnREFjd2hhQUFBd2VkVjZkUnhDZUN0Z0FWQWtBWXRNTmQxd2FKb0ZrS0dGcGdFQThGeUNGZ0FBVE5ZWXNJZ1RMTjdvRWtCeEJDenlGNi9SQjNNdkFwQWQ0VzBBQUo3dHhmMzl2U29DQURBcEFoWUFSUk93S0VEVERmRmEvZi9tWGdjZ1cvL3EyM3FqZlFBQVBKV0pGZ0FBVElhQUJVRFJCQ3pLOG5idUJRQ3lWb1VRQkMwQUFIZ3lRUXNBQUpLcjFxdFBPOTRGTEFES0kyQlJtS1liRmtidkE1bGJDSXdCQVBBY2doWUFBQ1F6Qml6T3gzOTJ2QU9VUmNDaVhCNU9Bcm1yZEJBQWdPY1F0QUFBWU84RUxBQ0tKbUJSc0tZYjRvcXZvN25YQWNqZVN5MEVBT0E1WHR6ZjN5c2dBQUI3SVdBQlVEUUJpOEkxM1JDdjR6ZXU0VUFoZnU3YjJqVUxBSUFuTWRFQ0FJQzlxTmFycFlBRlFKRUVMT2JqMG5VY0tNZ2loT0RhQlFEQWt3aGFBQUN3VTlWNmRSWWZ3Qmt6RGxBY0FZc1phYnFoQ2lHOG1Yc2RnS0pVMmdrQXdGTUpXZ0FBc0JNQ0ZnREZFckNZcDh1NUZ3QW96a0pMQVFCNHFoZjM5L2VLQndEQTFnaFlBQlJMd0dLbW1tNTRIVUw0UDNPdkExQ2tmL1p0ZmFPMUFBQThsb2tXQUFCc2hZQUZRTEYraTlNTUJDem1xZW1HUTlNc2dJTEY5U0dDRmdBQVBKcWdCUUFBejFLdFY0c3hZSEdpa2dCRmVUZE9zUEFBYXQ3T2hTaUJnc1h2TXU4MUdBQ0F4eEswQUFEZ1NRUXNBSW9sWU1Idm1tNDRIb01XQUtXcWRCWUFnS2NRdEFBQTRGRUVMQUNLSldEQmwrTDEva0JWZ0lMNVRnTUF3Sk84dUwrL1Z6a0FBSDVJd0FLZ1dBSVcvRTNURGZHNi96OHFBOHpBdi9xMjNtZzBBQUNQWWFJRkFBRGZWYTFYeDJQQTRvMUtBUlJGd0lMdldhb09NQk14V0Nab0FRREFvd2hhQUFEd1ZRSVdBTVVTc09DN21tNDRNOEVLbUpGS3N3RUFlQ3hCQ3dBQS9rTEFBcUJZQWhiOFVOTU5oeUdFUzVVQ1prVFFBZ0NBUnhPMEFBRGdkd0lXQU1VU3NPQXh6a01JQnlvR3pNakxHRExyMi9xanBnTUE4RkNDRmdBQU0xZXRWNGZqUTVXTHVkY0NvREFDRmp4SzB3M0g3Z2VBbVlwVExhNDBId0NBaHhLMEFBQ1lxYzhDRnQ1Y0JTaUxnQVZQWldVSU1GY0xRUXNBQUI1RDBBSUFZR1lFTEFDS0pXREJrelhkRUI4eXZsSkJZS1lXR2c4QXdHTUlXZ0FBeklTQUJVQ3hCQ3pZQnRNc2dEbXJkQjhBZ01kNGNYOS9yMkFBQUlXcjFxc1lybGdLV0FBVVJjQ0NyV2k2SWQ0bi9GczFnWm43WjkvV3Jxa0FBRHlJaVJZQUFBV3IxcXV6TVdCeHBNOEF4UkN3WUd1YWJqZ2M3eFVBNWk2dUQzazc5eUlBQVBBd2doWUFBQVVTc0FBb2tvQUZ1MkRpRmNBZnJBOEJBT0RCQkMwQUFBb2lZQUZRSkFFTGRxTHBodU1Rd2krcUMvQTdRUXNBQUI1TTBBSUFvQURWZXZWcHpLMkFCVUE1QkN6WU5TUHlBZjUwb2hZQUFEelVpL3Y3ZThVQ0FNalVHTEJZK2xFUW9DZ0NGdXhjMHcydlF3ai9SNlVCL3VMbnZxMnZsQVFBZ0I4eDBRSUFJRU1DRmdCRkVyQmdueTVWRytCdjR2b1FRUXNBQUg1STBBSUFJQ01DRmdCRkVyQmdyNXB1V0ZvM0J2QlZsYklBQVBBUWdoWUFBQm1vMXF0cWZQTlV3QUtnSEFJVzdGM1REWWNoaEhPVkIvaXFoYklBQVBBUWdoWUFBQk5XclZmSDR3U0xOL29FVUF3QkMxS0t3YzBESFFENHFxTVlTT3ZiK3FQeUFBRHdQWUlXQUFBVEpHQUJVQ1FCQzVKcXVxRnlid0h3UTNHcXhYdGxBZ0RnZXdRdEFBQW1STUFDb0VnQ0ZrekZwVTRBL0ZBbGFBRUF3SThJV2dBQVRFQzFYaDJPRHo4RUxBREtJV0RCWkRUZGNCWkNPTkVSZ0I5YUtCRUFBRDhpYUFFQWtOQVlzRGdmLzltWERsQUdBUXNtcGVtR3czRmlGZ0EvVnFrUkFBQS9JbWdCQUpDQWdBVkFrUVFzbUtwNHYzR2tPd0FQY3RCMFE5VzM5VWE1QUFENEZrRUxBSUE5RXJBQUtKS0FCWlBWZE1QeGVOOEJ3TVBGcVJhQ0ZnQUFmSk9nQlFEQUhnaFlBQlRuTG9Ud05vUndLV0RCeEMzZGV3QThtdlVoQUFCOGw2QUZBTUNPVmV2VjJmaVF3OGh1Z1B6RmdNWGxHTEQ0cUo5TVdkTU5peERDRzAwQ2VMU0ZrZ0VBOEQyQ0ZnQUFPeUpnQVZBVUFRdHlkS2xyQUUveVV0a0FBUGllRi9mMzl3b0VBTEJGQWhZQVJSR3dJRXROTjhUN2tmL1dQWUFuKzdsdjZ5dmxBd0RnYTB5MEFBRFlFZ0VMZ0tJSVdKQ3RwaHNPVGJNQWVMYTRQa1RRQWdDQXJ4SzBBQUI0cG1xOVdvd1BNNHlYQmNpZmdBVWxPQThoSE9na3dMTlV5Z2NBd0xjSVdnQUFQTkVZc0lnVExFN1VFQ0I3QWhZVW9lbUc0eERDaFc0Q1BKdWdCUUFBM3lSb0FRRHdTQUlXQUVVUnNLQTBiM1VVWUN1T1luaXRiK3NiNVFRQTRFdUNGZ0FBRHlSZ0FWQVVBUXVLMDNURHduMEt3RmJGcVJhQ0ZnQUEvSTJnQlFEQUQxVHIxZkg0TU82VldnRmtUOENDa3BsbUFiQmRNV2p4WGswQkFQaVNvQVVBd0RlTUFZczR3ZUtOR2dGa1Q4Q0NvalhkY0I3SDNPc3l3Rll0bEJNQWdLOFJ0QUFBK0lLQUJVQlJCQ3dvWHRNTmgrTzlDd0RiWlIwVEFBQmZKV2dCQURBU3NBQW9pb0FGY3hMdlh3NTBIR0Q3bW02bytyYmVLQzBBQUo4VHRBQUFacTlhcno2OUJmckwzR3NCVUFBQkMyWWxQZ0IwRHdPd1UvSHZyS0FGQUFCL0lXZ0JBTXpXR0xBNEgvOTVDeFFnYndJV3pOV2x6bE80Ni9GQnQvdDFVbG1FRU42cVBnQUFueE8wQUFCbVI4QUNvQ2dDRnN4VzB3MnZRd2duUGdFVTduejhPKyt6VGlxVnlnTUE4Q1ZCQ3dCZ05nUXNBSW9pWUFHbVdWQytYL3UyM2pUZGNDVm9RVUl2bTI0NDdOdmEvUVlBQVA4aGFBRUF6RUsxWHNWd3hWTEFBaUI3QWhid3h6U0xlRjl6cEJZVTdHNjhmNDlpME9KQ3MwbW9HaitIQUFEd08wRUxBS0JvMVhwMU52NUE2MEVFUU40RUxHQVUzNndlSjNSQnlaYWZUUkRZNkRTSkxRUXRBQUQ0bktBRkFGQWtBUXVBWWdoWXdOOWRtdEpGNFc3N3R2N1BhcHdZdUdpNjRVTmM0YUR4SkZJcFBBQUFueE8wQUFDS0ltQUJVQXdCQy9pS3BodmlXOVZ2MUliQ25YM2w5RGFDRmlTMFVId0FBRDRuYUFFQUZLRmFyMTZQRCtRRUxBRHlKbUFCMzdkVUh3cjNXOS9XWDF2UnNCRXlJcUdEcGh1Tys3YSswUVFBQUlLZ0JRQ1F1MnE5V293UEhFNDBFeUJyQWhid0EwMDNuTG5uWVFiT3YzR0tYd3Rmd0Q3RjlTR0NGZ0FBL0U3UUFnRElrb0FGUURFRUxPQUJtbTQ0Tk0yQ0dWaDlhMkpBMzlhYnBodDhCa2dwZmdkOXJ3TUFBQVJCQ3dBZ053SVdBTVVRc0lESE9iY2lqY0o5dWk1OHo3WHZBU1JVS1Q0QUFKOElXZ0FBV1JDd0FDaUdnQVU4VXRNTng5OVpwd0NsT08vYitrZlhoU3ZmQjBqSVp3OEFnUDhRdEFBQUpxMWFyNDdIZ01VYm5RTEltb0FGUEYzOGYrZEEvU2pZaDc2dDN6N2c5RFkrQktUVWRFTVYxOWhvQWdBQWdoWUF3Q1FKV0FBVVE4QUNucUhwaGpqVjY1VWFVcmlIVG16eGdKdlVGajZIQUFBRVFRc0FZR29FTEFDS0lXQUIyM0dwamhUdVhkL1dWdzg1eGI2dGI1cHV1QTBoSFBsUWtFaWw4QUFBQkVFTEFHQXFCQ3dBaW5FNy9qMS9MMkFCejlOMHcxa0k0YVV5VXJDNzhacnhHQnRCQ3hKYUtENEFBRUhRQWdCSXJWcXZEc2RSd2VkMmp3Tms3ZmVBeGViMDRpRTc5b0VmYUxyaDBEUUxadUF5VHFsNDVHbGVXYWREUWtmeDczUGYxc0trQUFBekoyZ0JBQ1FoWUFGUURBRUwySTJsZXlRS2QvdkVNTkhHQjRQRXFqSHdBd0RBakFsYUFBQjdKV0FCVUF3QkM5aVJwaHZpU3JWZjFKZkNMWjh5RmFCdjY2dW1HM3cyU0draGFBRUFnS0FGQUxBWEFoWUF4UkN3Z04zei94ZWx1KzdiK2ptZjh3OGhoSmMrSlNTeVVIZ0FBQVF0QUlDZHE5YXJzM0Vzc0lBRlFMNEVMR0FQbW02SUQvQk8xSnJDblQvejlLNEVMVWlvVW53QUFBUXRBSUNkR1FNV2NiLzRrU29EWkV2QUF2YkwvMnVVN2wzZjFwdG5udU56LzN0NGpvTzQ0cWx2Nnh0VkJBQ1lMMEVMQUdEckJDd0FpaUJnQVh2V2RNTzUreWNLZDdlRmFSWmhuR2dCS1MwRTR3QUE1azNRQWdEWUdnRUxnQ0lJV0VBQ1RUY2NqdmRSVUxMTHZxMC9QdmY4NGlTQnBodnVyQ1lrSWV0REFBQm1UdEFDQUhpMmFyMktiL05jMnBNTWtEVUJDMGpyMGtOakNuZmJ0L1UydzBSeGZjaUpEdzJKQ0ZvQUFNeWNvQVVBOEdSandHTHBCMDZBckFsWVFHSk5OOFFIZG0vMGdjSnRZMlhJNTY1OER5RWhuejBBZ0prVHRBQUFIazNBQXFBSUFoWXdIWmQ2UWVHdSs3Wit2K1ZUakVHTEN4OGNVbW02WWRHMzlaVUdBQURNazZBRkFQQmdBaFlBUlJDd2dBbHB1dUcxZXl0bTRHd0hwN2p4d1NHeGFnejhBQUF3UTRJV0FNQVBWZXZWY1FqaHJZY0FBRmtUc0lCcE1zMkMwdjNhdC9YTnRzK3hiK3VQVFRkOENDRzg5QWtpa1lXLzRRQUE4eVZvQVFCODB4aXdXTm9aRHBBMUFRdVlxS1liNG4zV2tmNVFzTHZ4KzhTdWJBUXRTS2hTZkFDQStSSzBBQUQrUnNBQ29BZ0NGakJoVFRmRSs2MXpQYUp3eXpoNVlvZW5lT1U3Q3drZE5kMXdjcXNraFFBQUlBQkpSRUZVdU9QUE9BQUFFeVZvQVFEOGg0QUZRQkVFTENBUDhaN3JRSzhvMkllK3JYZTlWbUhqQTBSaWNYM0llMDBBQUpnZlFRc0FJQVlzRHNmZHNnSVdBUGtTc0lCTU5OMndjTi9GRE94OFlrdmYxcHVtRzN5V1NLa1N0QUFBbUNkQkN3Q1lzVEZnY1Q3KzgwWWxRSjRFTENBL1N6MmpjTC8xYlgyMXAxTzhEaUdjK0VDUnlFTGhBUURtU2RBQ0FHWkl3QUtnQ0FJV2tLR21HODQ4RkdZR2RqN040ak5YL3A4aW9VcnhBUURtU2RBQ0FHWkV3QUtnQ0FJV2tLbW1HdzVOczJBR1ZuMWIzK3p4TkRjK1ZDUjAwSFJERmRmWWFBSUF3THdJV2dEQVRGVHIxVkxBQWlCckFoYVF2M2d2ZHFTUEZPd3VoSEM1NTlQYjE0b1MrSlpLNEFjQVlINWUzTi9mYXpzQUZLeGFyODdHTnlmOXFBK1FKd0VMS0VEVERjY2hoUCtubHhUdXYvcTIzdnYxcXVtR0c5OTNTT2hkMzlabkdnQUFNQzhtV2dCQW9RUXNBTEluWUFGbDJmZGIvckJ2MXlsQ0ZxT043ejBrVkNrK0FNRDhDRm9BUUdFRUxBQ3lKMkFCaFdtNllSRkNlS1d2Rkc2WjhQU3UvRDlHUWk4Vkh3QmdmZ1F0QUtBUUFoWUEyUk93Z0hLWlprSHA0dXFFcTRUbnVQRUpJNlVZcUV2OC93QUFBSHNtYUFFQW1hdldxOFVZc0RqUlM0QXNDVmhBd1pwdU9QZTJNNFc3U3p6TklzUUgzRTAzK0p5UjBtS2NyQUlBd0V3SVdnQkFwZ1FzQUxJbllBR0ZhN3JoTVBVRGFOaUR5NzZ0YnlaUTZHdmZqVWlvVW53QWdIa1J0QUNBekFoWUFHUlB3QUxtSTk2ekhlZzNCYnVkMEdxY2plOUlKQ1JvQVFBd000SVdBSkFKQVF1QTdGMlBBUXRqcFdFR21tNDREaUg4b3RjVTdyeHY2NDhUT2NYTkJJNkIrVHFLZi9jbk10MEZBSUE5RUxRQWdJbXIxcXZqTVdEeFJxOEFzaVJnQWZOa2FnMmx1KzdiK3YyRXp0RjFsdFRpVkF0QkN3Q0FtUkMwQUlDSkVyQUF5SjZBQmN4VTB3MExVOGlZZ2ZNcG5XS2NKTkIwdzUxMVBTUVVneFpUQ2g4QkFMQkRnaFlBTURFQ0ZnRFpFN0FBVExPZ2RPLzZ0cDdpcW81NDdYMDFnZU5nbmhiNkRnQXdINElXQURBUkFoWUEyUk93QU9JMGkzZy9kNlFTRk94dWF0TXNQck1SdENBaGs0d0FBR1pFMEFJQUVxdldxOFB4aDhvTHZRRElrb0FGOEx1bUd3NG4vQUFhdHVXeWIrdVBFNjNtbGU5VnBOUjBRelhSYVM4QUFHeVpvQVVBSlBKWndPTGNIbUdBTEFsWUFGKzZkRjlINFc3N3RsNU85UlQ3dHI1cXVtRUNSOEtNTGNiSktnQUFGRTdRQWdEMlRNQUNJSHNDRnNEZnhMZVlyWUJqQnM0eU9NVVBJWVNYRXpnTzVxblNkd0NBZVJDMEFJQTlFYkFBeUo2QUJmQTlsNnBENGE3anhJZ01UbkVqYUVGQ2doWUFBRFB4azBZRHdPNVY2MVVNVjl5TSs0S0ZMQUR5RWdNV1AyOU9MeFpDRnNEWE5OM3dPb1J3b2pnVUxvZHBGcEZyTlNtOWJMcmhVQWNBQU1wbm9nVUE3RkMxWHNVZkkrTU80eU4xQnNpT0NSYkFENDBQMUV5em9IUy85bTE5azhrNWJpWndETXhiSmZBREFGQStRUXNBMkFFQkM0Q3NDVmdBajNIdW5vL0MzWTNmYmJMUXQvV202WVk3a3dSSmFDRm9BUUJRUGtFTEFOZ2lBUXVBckFsWUFJL1NkTVB4R0xTQWtwMzNiZjB4cy9QYldPZERRcFhpQXdDVVQ5QUNBTGFnV3EvaUd5dHZCU3dBc2lSZ0FUelYwbHZ6Rk81RDM5WnZNenpGSzBFTEVsb29QZ0JBK1FRdEFPQVp4b0RGMG85NEFGa1NzQUNlck9tR2VCLzRSZ1VwWEs0VFd6WVRPQWJtNnlCT1BPcmIrc1puQUFDZ1hJSVdBUEFFQWhZQVdST3dBTFpocVlvVTdyZStyWE85VnJyR2sxcGNIeUpvQVFCUU1FRUxBSGdFQVF1QXJBbFlBRnZSZE1PWiswRm1JTmRwRnFGdjY0OU5OOXhhN1VoQzhiZUQ5eG9BQUZBdVFRc0FlSUJxdllwdm8xejZRUjBnU3dJV3dOWTAzWEE0M2hkQ3lWWUZyRDI0c3Q2SGhDckZCd0FvbTZBRkFIeEh0VjRkanhNcy9FQUhrQjhCQzJBWDRsditCeXBMd1c0TENSTnRmSThqSVM5cEFBQVVUdEFDQUw1Q3dBSWdhd0lXd0U0MDNSRHZFUzlVbDhJdDQrcU5BazV4TTRGallNYWFibGowYmUxK0ZBQ2dVSUlXQVBBWkFRdUFyQWxZQUx0bVpRaWx1KzdiK20wSjV4Z2ZjRGZkTUlFalljYXFjWVVOQUFBRkVyUUFnRDhDRnA5MmJRdFlBT1JId0FMWXVmaG1jZ2pobFVwVHVHVmhwM2R0aFFNSlZZb1BBRkF1UVFzQVptME1XSnpidFEyUUpRRUxZSitLZU1zZnZ1TmRnV3NPTm9JV0pMUlFmQUNBY2dsYUFEQkxBaFlBV1JPd0FQYXE2WVo0ejNpazZoVHNyc0JwRm1GYzIvRExCSTZEZVRwcXV1R3diK3VQK2c4QVVCNUJDd0JtUmNBQ0lHc0NGc0RleFlka2hUNkFoczlkOW0xOVUyQkZOaE00QnVhdEdnTS9BQUFVUnRBQ2dGa1FzQURJbW9BRmtOTFMvU09GdSszYnVzZ3dVUXlQTk4xdzUvOWhFbG9JV2dBQWxFblFBb0RpVmV2VjJmZ0R1WEhQQUhrUnNBQ1NhcnJoMk5vQlp1Qzg4Rk9NOXhHdkpuQWN6Tk5DM3dFQXlpUm9BVUN4QkN3QXNpVmdBVXpGVzUyZ2NOZDlXNzh2L0J3M2doWWtWQ2srQUVDWkJDMEFLSTZBQlVDMmZvczc0Z1VzZ0Nsb3V1RjFDT0ZFTXloYzZkTXN3ampSNG1JQ3g4RThIY1RwU0hHTmpmNERBSlJGMEFLQVlnaFlBR1RyM1RqQndnL1F3SlJjNmdhRmU5ZTM5YWIwayt6YitxcnBoZ2tjQ1RPMk1DRUpBS0E4Z2hZQVpLOWFyeGJqRCtFdmRSTWdLd0lXd0NRMTNTQzhTK251WmpMTjRwTVB2aStTa1BVaEFBQUZFclFBSUZ0andHSnBwRE5BZGdRc2dNbHF1dUZ3WmcrZ21hZGwzOVlmWjNUbVY0SVdKTFJRZkFDQThnaGFBSkFkQVF1QWJBbFlBRG1JazlJT2RJcUMzZlp0UGJmVk9NV3ZTR0hTaEh3QUFBb2thQUZBTmdRc0FMSWxZQUZrb2VtR09ONzlqVzVSdUxNWk5salFncVNhYmxqMGJYMmxDd0FBNVJDMEFHRHlxdlhxZUh5ejhKVnVBV1JGd0FMSXpkemU4bWQrcnVmNHNMZHY2MDNURFhlbTFaQlFOYTZ3QVFDZ0VJSVdBRXpXR0xCWWVxc1FJRHNDRmtCMm1tNDRNem1OR1pqak5JdFBOdjRmSjZHRk1COEFRRmtFTFFDWUhBRUxnR3dKV0FCWmFycmhjTHovaEpMOTJyZjFuSy9SVjRJV0pGUXBQZ0JBV1FRdEFKZ01BUXVBYkFsWUFMazdEeUVjNlNJRnV4TW0rajFvY1RHQjQyQ2VqbUtvcjIvcmovb1BBRkFHUVFzQWtxdldxMDl2RVA2aUd3QlpFYkFBc3RkMHcvRVl0SUNTblh2QSsvdnFFRWdwcmc5NXJ3TUFBR1VRdEFBZ21URmdjVDcrTzlBSmdHd0lXQUFsdVhRdlN1RSs5RzM5ZHU1TmprR1RwaHR1VGE4aG9VclFBZ0NnSElJV0FPeWRnQVZBdGdRc2dLSTAzUkRmTG42bHF4VE94SlkvWFZsVlNVSUx4UWNBS0llZ0JRQjdJMkFCa0MwQkM2QlVsenBMNFg3cjIvcEtrLzlqSTJoQlFwWGlBd0NVUTlBQ2dMMm8xcXVsZ0FWQWRnUXNnR0kxM1hBV1FuaXB3eFRzempTTHZ4RTZJYVdEcGh1cXZxMDN1Z0FBa0Q5QkN3QjJxbHF2NGcvWVMzdHdBYklpWUFFVXJlbUdROU1zbUlITHZxMWR5ejhUSDNBMzNUQ1o0MkdXcW5HeUNnQUFtUk8wQUdBbkJDd0FzaVJnQWN5RlNXdVU3bGFZNkp1dVF3Z25FejAyeXJjSUlielZad0NBL0FsYUFMQlZBaFlBV1JLd0FHYWo2WWJqRU1LRmpsTzRaZC9XSHpYNXF6YUNGaVJVS1Q0QVFCa0VMUURZaW1xOWVqMitNU1ZnQVpBUEFRdGdqcnhKVE9tdSs3YjJPZisycXhEQ0wxTTlPSXIzVW9zQkFNb2dhQUhBczFUcjFXS2NZT0dOSUlCOENGZ0FzOVIwdzhKOUt6T3cxT1R2Mmt6NDJKaUJlQzNxMi9wS3J3RUE4aVpvQWNDVENGZ0FaRW5BQXBnN2IvbFR1bmNlNEg1ZjM5WTNUVGZjbXNaSVFvdHhzZ29BQUJrVHRBRGdVUVFzQUxJa1lBSE1YdE1ONXg2c1VyaTdFTUs1SmovSXh0OERFcW9VSHdBZ2Y0SVdBRHlJZ0FWQWxnUXNBUDRJV1J4YXA4QU1YUFp0L1ZHakh5UUdMVjVsY0p5VVNkQUNBS0FBZ2hZQWZGZTFYaDJQUDBxL1VTbUFiQWhZQVB4VnZKODlVQk1LZHR1M3RURFJ3OFcxRFJlNUhDekZPV3E2NFRpdXNkRmFBSUI4Q1ZvQThGVUNGZ0JaRXJBQStFTFREZkhONFYvVWhjSlpHZklJZlZ0Zk5kMlF6ZkZTcEhodGNzOE9BSkF4UVFzQS9rTEFBaUJMQWhZQTMzYXBOaFR1dW0vcjk1cjhhQjlDQ0M4ek8yYktFZGV6K3Y4V0FDQmpnaFlBL0U3QUFpQkxBaFlBMzlGMHcrc1F3b2thVVRqVExKN21TdENDaENyRkJ3REltNkFGd014VjY5WGgrTVBjdWIzVkFOa1FzQUI0R05Nc0tOMnZmVnR2ZFBsSjFJMlVoQUFCQURJbmFBRXdVd0lXQUZrU3NBQjRvS1liNHJTMkkvV2lZSGZqVkVLZTVrcmRTS25waGtwUUNnQWdYNElXQURNallBR1FuYnR4ZjdPQUJjQUROZDF3YUowQ003RHMyL3FqUmo5TjM5WTNUVGZjK1Y1TVFndVRWUUFBOGlWb0FUQVRBaFlBMmJrYlI5NWZiazR2UEVRQmVKeEw5N3dVN3JadmE2dHhubTlqaFFNSlZZb1BBSkF2UVF1QUdhaldxek0vTmdOa1E4QUM0Qm1hYm9odkNMOVJRd3AzcHNGYmNTVm9RVUtDRmdBQUdSTzBBQ2pZR0xDd214b2dEd0lXQU51eFZFY0tkOTIzOVpVbWIwV3M0MFVCNTBHZVhzWlZWMVlBQVFEa1NkQUNvRUFDRmdCWkViQUEySkttRzg2OG5jNE1tR2F4UFp0U1RvUnNWV1BnQndDQXpBaGFBQlJFd0FJZ0t3SVdBRnNVM3dvMnpZSVpXUFZ0ZmFQUjJ4RW5DVFRkOENGT0ZpamhmTWpTUXRBQ0FDQlBnaFlBQmFqV3EvakYvSzJBQlVBV0JDd0FkdVBjL1RDRiszUVB3WFp0QkMxSXFGSjhBSUE4Q1ZvQVpHd01XQ3lOUndiSWdvQUZ3STQwM1hBY1FyaFFYd3AzSGljd2FQTFd4YURGbThMT2lYd3M5QW9BSUUrQ0ZnQVpFckFBeUlxQUJjRHVlY3VmMG4zbzIvcXRMdStFdFEya2RCRERnbFlDQVFEa1I5QUNJQ01DRmdCWkViQUEySU9tRytJOThpdTFwbkRuR3J3YmZWdHZtbTRvOGRUSXg2ZDFzQUFBWkVUUUFpQUQxWHAxUEg3cEZyQUFtRDRCQzREOU1zMkMwcjNyMjlyVWhkMjY5bjJiaENyRkJ3RElqNkFGd0lTTkFZdWxmYkVBV1JDd0FOaXpwaHZPUWdndjFaMkMzWTNmQ2RtdEswRUxFaEswQUFESWtLQUZ3QVFKV0FCa1JjQUNJSUdtR3c1TnMyQUdMdnUydnRIb25kc1VmbjVNbTVBUEFFQ0dCQzBBSmtUQUFpQXJBaFlBYWNYNzVnTTlvR0Mzd2tSN0kyaEJVazAzTEt3SUFnRElpNkFGd0FSVTY5V250L0VFTEFDbVQ4QUNJTEdtRzJKQStSZDlvSERMdnEzZGEreEJuQnJTZEVNTXRod1ZmN0pNVlRXdXNBRUFJQk9DRmdBSmpRR0w4L0dmdC9FQXBrM0FBbUE2M3VvRmhidnUyOXJuZkw4MmdoWWtWQ2srQUVCZUJDMEFFaG5YaEd3RUxBQW1UOEFDWUVMaWVIWDc3Sm1CYzAzZXV6aE40TlhNenBucFdPZ0ZBRUJlZnRJdmdEUTJweGMzSVlRYjVRZVlyQml3V0lVUWpqZW5GMHNoQzRESjhKWS9wWHZYdC9WR2wvZE96VW5wcU9tR1F4MEFBTWlIb0FWQVduNGtCcGdlQVF1QWlXcTY0ZHhvZndwM1o1cEZHbjFiWDgzeHZKa1U2ME1BQURKaWRRaEFXdTlEQ1AvV0E0QkpzQ0lFWU1MR04zMlhla1RoTHZ1MmRoK1N6clhWUkNTMEdGZllBQUNRQVJNdEFCSWExNGQ4MEFPQXBFeXdBTWhERE1NZDZCVUZ1KzNiV3Bnb0xldERTR21oK2dBQStURFJBaUM5dDZaYUFDUmhnZ1ZBSnBwdWlPUFUzK2dYaGJNeUpEMUJDMUt5T2dRQUlDTW1XZ0NrOTE0UEFQYktCQXVBL0Z6cUdZVzc3dHZhZDhQMHJHMGdwWU14V0FnQVFBWUVMUUFTc3o0RVlHOEVMQUF5MUhURDZ4RENpZDVSdURNTlRxOXY2NXZ4bmhGU0ViUUFBTWlFMVNFQTAyQjlDTUR1V0JFQ2tEZlRMQ2pkcitNRGZxWWhUclY0cFJja0ltZ0JBSkFKUVF1QWFYZ3ZhQUd3ZFFJV0FKbHJ1bUVaUWpqU1J3b1c3MWVXR2p3cEcwRUxFbG9vUGdCQUhxd09BWmlBY1gzSXRWNEFiSVVWSVFBRmFMcmhPSVJ3cnBjVWJ0bTN0WHVWYWJtYWV3Rkk2cVh5QXdEa1FkQUNZRHJlNmdYQXN3aFlBSlFsdnVWL29LY1U3TFp2YTZ0eHBtY3o5d0tRVnRNTnBsb0FBR1JBMEFKZ090N3JCY0NUQ0ZnQUZHWjh5UFJHWHluY21RWlB6emhoNU1QYzYwQlNsZklEQUV5Zm9BWEFSSXdQQm4vVEQ0QUh1dzBoL0c4QkM0QWlMYldWd3YzV3Q3VVZGZE5scWdVcG1XZ0JBSkNCZjJnU3dLVEVxUmF2dEFUZ3UyTEFJZ1lyckZ3Q0tGRFREZkV0L3hPOXBYRG5HanhwVjZicWtKQ0pGZ0FBR1REUkFtQmFyQThCK0xZWXNQaXZ6ZW5Gc1pBRlFKbWFiamdNSVZ4cUw0VmI5VzE5bzhtVFpxSUZLUjJOMTBNQUFDWk0wQUpnUXF3UEFmZ3FBUXVBK1loditSL29Od1c3RXlhYXZyNnRCUzFJemZvUUFJQ0pzem9FWUhxc0R3SDRneFVoQURQU2RNTnhDT0ZDenluY2VkL1dIelU1QzlmV0dKRlFaZW9wQU1DMENWb0FURS84SXYzZitnTE1tSUFGd0R4NXk1L1NYZmR0N2Y0bUgxZUNGaVJrb2dVQXdNUlpIUUl3TWRhSEFETm1SUWpBVERYZHNERFZqUmxZYW5KV3JBOGhKU0VmQUlDSkU3UUFtQ1lQR0lFNUViQUF3RFFMU3ZldWIrc3JYYzZLZnBGVTB3MlZEZ0FBVEplZ0JjQUViVTR2NHZxUU83MEJDaWRnQVVCOGtIUWVRbmlwRWhUc3pqU0wvUFJ0L1hHOFg0VlVCQzBBQUNic0g1b0RNRmt4YlBGR2U0QUN4UitzbDhJVkFEVGRjT2dCTkROdzJiZjFqVVpuS2E0UE9acDdFVWhtWWVJcEFNQjBDVm9BVEplZ0JWQWFBUXNBdmhSREZnZXFRc0Z1cmNiSldsd2Y4bXJ1UlNBWkV5MEFBQ2JNNmhDQWliSStCQ2lJRlNFQS9FM1REY2NoaEY5VWhzS2RqeXNveU5ORzMwaklXaTBBZ0FrVHRBQ1l0dmY2QTJSTXdBS0E3M0Z0b0hUWGZWdjdUcGV4dnEydjVsNEQwbXE2WWFFRkFBRFRaSFVJd0xSWkh3TGt5SW9RQUw1cmZIQjBva29VN2x5RGkzRHQ3eFVKTGNZVk5nQUFUSXlKRmdBVFpuMElrQmtUTEFCNEtOY0pTdmV1YjJ0cko4cWdqNlJVcVQ0QXdEU1phQUV3ZmFaYUFGTm5nZ1VBRDlaMHd6S0VjS1JpRk96T05JdWl4R2tDdjh5OUNDUWphQUVBTUZFbVdnQk1ud2VYd0ZTWllBSEFvelRkY09nQk5ETncyYmYxUjQwdWhva1dwSFRVZE1PeERnQUFUSStnQmNERWJVNHZyc2FIbVFCVElXQUJ3Rk5kaGhBT1ZJK0MzZlp0dmRUZ2N2UnRmV09sSjRtWmFnRUFNRUdDRmdCNWVLOVB3QVFJV0FEd1pFMDNWRmJpTVFObm1seWtxN2tYZ0tRV3lnOEFNRDMvMEJPQUxMeTFFeFpJS0FZc2xzSVZBRHpUcFFKU3VPdStyVDJRTDFOY0gvSnE3a1VnR1JNdEFBQW15RVFMZ0F4c1RpODIxb2NBQ1poZ0FjQldOTjN3T29Sd29wb1V6alNMY2duUWtKTHJKd0RBQkFsYUFPVEQraEJnWHdRc0FOaWFwaHNPVGJOZ0JuN3QyL3BHbzh0a1VnbXBqZXUzQUFDWUVLdERBUEpoZlFpd2ExYUVBTEFMNXlHRUk1V2xZSGZ4SGtxRGkvY2hoUEJ5N2tVZ21jVzR3Z1lBZ0lrdzBRSWdFOWFIQUR0a2dnVUFPOUYwdy9FWXRJQ1NMZnUyL3FqRHhmT1FtNVJNdEFBQW1CZ1RMUUR5OHQ1VUMyQ0xUTEFBWU5maVcvNEhxa3pCUHZSdGJUWE9QTVQxSVcvbVhnU1NFYlFBQUpnWUV5MEE4dUlIUEdBYlRMQUFZT2VhYmxoNEtNa01tTmd5SHlaYWtOTExwaHNPZFFBQVlEb0VMUUF5c2ptOXVCbjN3Z0k4aFlBRkFQc2tKRXpwZnV2YitrcVg1NkZ2NnhpMHVKdDdIVWpLVkFzQWdBbXhPZ1FnUC9IaDZMLzFEWGdFSzBJQTJLdW1HODdpMjdlcVR1Rk1zNWlmR0xZNG1Yc1JTR1l4cnJBQkFHQUNCQzBBOHZOZTBBSjRJQUVMQVBadUhHMXVtZ1dsVy9WdGZhUExzM01sYUVGQ0M4VUhBSmdPcTBNQU1tTjlDUEFBMXlHRS84K0tFQUFTaVcvNUh5ZytCYnNUSnBxdHpkd0xRRkpXaHdBQVRJaWdCVUNlUERnRnZpWUdMSDdlbkY0c05xY1g3MVVJZ0gxcnV1RTRoSENoOEJUdXZHL3JqNW84UzlZMmtOTEJlSjBGQUdBQ0JDMEE4dVFCS3ZDNXp3TVdmdndGSUNWditWTzY2NzZ0QmQ5bmFnelkzTTY5RGlSbGZRZ0F3RVFJV2dCa3lQb1FZQ1JnQWNCa05OMFFILzY4MGhFS3Q5VGcyWFBmVFVyV2h3QUFUTVEvTkFJZ1cvRXRxbjlySDh4U0RGZ3NoU3NBbUJoditWTzZkMzFidS85aUUwSjRNL3Nxa0lxZ0JRREFSQWhhQU9UcnZhQUZ6STZBQlFDVDFIVERlUWpoU0hjbzJKMXBGb3pjaTVQU2llb0RBRXpEaS92N2U2MEF5RlMxWHIwM25obG1RY0FDZ01scXV1RXdoQkJYMngzb0VnVmI5VzB0YU1Idm1tN3dneW9wL1d5NkRnQkFlaVphQU9STjBBTEtKbUFCUUE2V1FoWVU3bGJJZ2k5Y215eEFRcFhKS2dBQTZRbGFBT1F0QmkzK1d3K2hPQUlXQUdTaDZZYmpFTUl2dWtYaHpqV1lMMndFTFVpb1Vud0FnUFIrMGdPQWZHMU9MejZHRUg3VFFpaEdERmo4dkRtOVdBaFpBSkNKdHhwRjRhNzd0bjZ2eVh6QnZUb3BMVlFmQUNBOVFRdUEvUG5SRC9JbllBRkFkcHB1ZU8yTmJtYkFOQXUrWnFNcUpIVFVkTU9oQmdBQXBDVm9BWkEvUVF2SWw0QUZBRG03MUQwSzk2NXZhdy9VK1p1K3JXOUNDTGNxUTBLbVdnQUFKQ1pvQVpBNTYwTWdTd0lXQUdTdDZZWmxmS05XRnluWW5Xa1cvSUFRRGlsVnFnOEFrTlkvMUIrZ0NIR3F4U3V0aE1tTEFZdWxjQVVBT1J2SGxYc0FUZW1XZlZ0LzFHVytZK043T0FtWmFBRUFrSmlnQlVBWll0RGl2L1VTSmt2QUFvQ1N4SlVoQnpwS3dXNzd0cllhaHgrSjkvWVhxa1FpSmxvQUFDUm1kUWhBQWNiMUllLzBFaWJIaWhBQWl0SjBRM3lEOW8ydVVyZ3pEZVpIK3JaMmYwOUtCMDAzQ0ZzQUFDUmtvZ1ZBT2Q3NzBSc213d1FMQUVxMTFGa0tkKzBCT28vd0lZVHdVc0ZJcEJwWDJBQUFrSUNKRmdDRjJKeGV4S0RGblg1Q1VpWllBRkNzcGh2aVcvNG5Pa3poVExQZ01kenprNUtKRmdBQUNabG9BVkFXVXkwZ0RSTXNBQ2hhMHcySHBsa3dBNy8yYlgyajBUeUNhUUtrdEZCOUFJQjBCQzBBeWlKb0Fmc2xZQUhBWEp5SEVJNTBtNExkQ1JQeEJJSVdwR1J0RFFCQVFpL3U3Ky9WSDZBZzFYcjFNWVJ3b0tld1V3SVdBTXhHMHczSDQ4TkU5NWlVN0wvNnRuNnJ3enhXMHcyK2c1UFN6MzFiKzE0S0FKQ0FpUllBNVRIVkFuWkh3QUtBT2JyMEVKSENmUkN5NEJsaUVPMUVBVW1rQ2lINGZnb0FrTUJQaWc1UW5QZGFDbHNYQXhZL2IwNHZGa0lXQU14SjB3MXgvL3NyVGFkdzV4ck1NL2grUUVvTDFRY0FTRVBRQXFBd205T0w5K04rWWVENUJDd0FtTHZMdVJlQTR2MW03RDdQNVBORFNwWHFBd0NrSVdnQlVDWlRMZUI1QkN3QW1MMm1HODVDQ0MvblhnZUtaNW9GejdWUlFSSTZhcnJoV0FNQUFQWlAwQUtnVE40OGhLY1JzQUNBUDBJV2grNHBtWUZWMzlZM0dzMXo5RzM5TVlSd3E0Z2taS29GQUVBQ2doWUFCZHFjWG16ODBBT1BJbUFCQUg4VjMvSS9VQk1LZGl0TXhCYjVEa0ZLZ2hZQUFBbjhROUVCaWhYWGgveWl2ZkJkTVdDeEZLNEFnRCtOSThndmxJVENMY2RKQkxBTjhXV0hOeXBKSWd1RkJ3RFlQMEVMZ0hLOUZiU0FieEt3QUlCdmU2czJGTzY2YjJ1ZmM3Yko5d3BTT2xGOUFJRDllM0YvZjYvc0FJV3ExcXU0Yi9oSWYrRS8zbzBCQzd1NEFlQXJtbTZJYjhYK2o5cFF1Si83dHZaZ25LMXF1c0dQcktUMHI3NnROem9BQUxBL0psb0FsTTM2RVBpRGdBVUFQSXkzL0NuZE95RUxkdVRhWkFFU3FzWVZOZ0FBN01sUENnMVFORCtVTTNjeFlQSFB6ZW5GbVpBRkFIeGYwdzNucHFGUnVMc1F3cmttc3lNQ1BLUzBVSDBBZ1AwU3RBQW8yT2IwSXI3TmNLdkh6SkNBQlFBOFF0TU5oM0g2azVwUnVNdStyVDlxTWp0aW1nQXBWYW9QQUxCZlZvY0FsTS82RU9iRWloQUFlSm9Zc2poUU93cDIyN2UxTUJHN0pHaEJTaTlWSHdCZ3YweTBBQ2pmcFI0ekF5WllBTUFUTmQxUUNlWXlBMWFHc0ZOOVc5K1lLRWxLVFRkWUh3SUFzRWVDRmdDRkd4ODZmOUJuQ2lWZ0FRRFBKNWhMNmE3N3RuNnZ5K3lCcVJha0pHZ0JBTEJIVm9jQXpNUGJFTUsvOVpxQ1dCRUNBRnZRZE1QckVNS0pXbEk0MHl6WWw2c1F3aXZWSnBGSzRRRUE5a2ZRQW1BZTNndGFVQWdCQ3dEWUx0TXNLTjJ2ZlZ1Yk1zQysrS3lSa29rV0FBQjc5T0wrL2w2OUFXYWdXcS9pRHo0djlacE1DVmdBd0pZMTNiQU1JVnlvS3dXN0N5RWM5MjM5VVpQWmw2WWIvTmhLU3YvczI5cjNaZ0NBUFREUkFtQStyQThoUndJV0FMQURUVGNjVzZmQURDeUZMRWpnZzVjY1NDaXVEL0g5R1FCZ0R3UXRBT2JEK2hCeUltQUJBTHNWcDFrY3FERUZ1KzNiMm1vY1VyZ1N0Q0NoeGZqN0R3QUFPL2FUQWdQTXcvakErb04yTTNFeFlQSFB6ZW5GbVpBRkFPeEcwdzN4SWN3YjVhVndaeHBNSWh1Rko2Rks4UUVBOXNORUM0QjVzVDZFcVRMQkFnRDJaNm5XRk82NmIrc3JUU1lSbnoxU09sRjlBSUQ5TU5FQ1lGNk1qMlJxVExBQWdEMXF1dUhNUXhobXdEUUxrdW5iT242dnVkTUJVbW02d1ZRTEFJQTlFTFFBbUpIeFFmWnZlczVFZkJDd0FJRDlhYnJoMERRTFptQTFQdWlHbEt3UElhV0Y2Z01BN0o2Z0JjRDhtR3JCVkx5czFxdGozUUNBdlRrUElSd3BOd1dMVXdRdU5aZ0pzRDZFbEV5MEFBRFlBMEVMZ1BrUnRHQktYdXNHQU94ZTB3MHgzSGloMUJUdXZHL3JqNXJNQkFoYWtKS0pGZ0FBZXlCb0FUQXptOU9MajlhSE1DSDJad1BBZm5qTG45Sjk2TnY2clM0ekVWYUhrTkxSdUM0TUFJQWRFclFBbUNkVExaZ0s2ME1BWU1lYWJvaHZ0cjVTWndwM3JzRk14VGhaNVlPR2tKRDFJUUFBT3lab0FUQlBnaFpNaWZVaEFMQmJwbGxRdXQvNnRyYXFnYWt4MVlLVXJBOEJBTmd4UVF1QUdiSStoSW54OWlFQTdFalREWEZOMTB2MXBXQjM3aWVaS0VFTFVoSzBBQURZTVVFTGdQa3kxWUtwT0tyV0syTk5BV0RMeHYzc3BsbFF1c3Urclc5MG1Ra3laWVdVZk1jR0FOZ3hRUXVBK1hvL3Z2MEZVM0NtQ3dDd2Rjc1F3b0d5VXJCYllTS21xbTlyRXkxSTZhRHBobU1kQUFEWUhVRUxnSmthMTRlWWFzRlV2TllKQU5pZThlSEtMMHBLNFpaOVczL1VaQ2JzV25OSXlQb1FBSUFkRXJRQW1EZEJDNmJDK2hBQTJLNjM2a25ocnZ1MjlqbG42cXdQSVNYZnNRRUFka2pRQW1ER05xY1gxb2N3SmRhSEFNQVdOTjBRMzJBOVVVc0t0OVJnTW1COUNDa0pXZ0FBN0pDZ0JRQ21XakFWMW9jQXdIWjR5NS9TdmV2YjJxUUFjdUJ6U2twQ2x3QUFPeVJvQVlDZ0JWTVIxNGZZSVFzQXo5QjB3M204cHFvaEJZc1QrYzQxbUJ6MGJmMHhoSENyV2FReVRya0NBR0FIQkMwQVpzNzZFQ2JHK2hBQWVLS21HdzZ0VTJBR0xzZUgxNUFMNjBOSXlmb1FBSUFkRWJRQUlKaHF3WVJZSHdJQVQzY1pRamhRUHdwMjI3ZTFNQkc1c1Q2RWxFeTBBQURZRVVFTEFJS2dCUk55VUsxWHdoWUE4RWhOTjhRM1Z0K29HNFd6TW9RY21XaEJTaVphQUFEc2lLQUZBTmFITURXQ0ZnRHdlSmRxUnVHdSs3WVdFQ2M3ZlZ1YmFFRktSK05xTVFBQXRrelFBb0JQM3FvRUV5Rm9BUUNQMEhSRHZIYWVxQm1GTzlOZ01uYXRlU1JrZlFnQXdBNElXZ0R3aWFBRlUyRjlDQUE4MFBpV3Fta1dsTzdYdnExdmRKbU1XUjlDU3RhSEFBRHNnS0FGQUwvYm5GN0VIMzV1VllPSkVMUUFnSWM1ajJQQjFZcUN4UldIU3cwbWM0SVdwR1NpQlFEQURnaGFBUEE1TzQrWmlqZlZlbVdQTEFCOFI5TU54MlBRQWtxMjdOdjZvdzZUdVNzTkpDRVRMUUFBZGtEUUFvRFBXUi9DbEpocUFRRGZGOS95UDFBakNuYmJ0N1hWT0dSdlhIMXpwNU1rY3RCMGc3QUZBTUNXQ1ZvQThCL1doekF4Z2hZQThBMU5OOFF4NEcvVWg4S2RhVEFGTWRXQ2xBUXRBQUMyVE5BQ2dDOVpIOEpVdkxJK0JBQythYWswRk82M3ZxMDltS1lrRzkwa0lVRUxBSUF0RTdTb082VTBBQUFnQUVsRVFWUUE0RXZXaHpBbHBsb0F3QmVhYm9oditaK29DNFU3MTJBS0l6aEVTZ3ZWQndEWUxrRUxBUDdpLzJmdmJtNGl5YkkyQUljSGlRZmdBWGdBcTloV1NTR0YySUVIblI2UWVKRHRRYkZESVlWVXRZMVY0VUhpQVhoQWV2QXBadUtiNlo2dUgzNHk4NTY0OTNsTU9JZWlnSHp2ZWRXSEVJeWdCUUQ4UmRNTjQ3V250Wm1RdWR1K3JaOHNtY3k0YUVGS3A2WVBBTEJiZ2hZQS9JaXJGa1F4MW9lYzJBWUEvTWY0eW45aEhHUnNLMHhFanZxMmZxbXE2dEZ5U2FYcEJsY3RBQUIyU05BQ2dCOFJ0Q0FTVnkwQTROOGZrSXpod3h1eklIUEw2UU5weUpHckZxUWthQUVBc0VPQ0ZnRDh3K2J5NXNsTEd3SzV0Z3dBK0Jldi9NbmRZOS9XUXQvazdMdnRrdENaNFFNQTdJNmdCUUEvNHcrY1JIR3FQZ1NBMGszbnZqK1ZQZ2V5dDdSaU11ZWlCU2tKV2dBQTdKQ2dCUUEvODlWa0NFUjlDQUNsYzgyQzNOMzFiZTIxUDFucjIzb01XbXh0bVVTT3B4b3lBQUIyUU5BQ2dCOVNIMEl3NmtNQUtGYlREZU1yLzFOZkFXUnMvT0I1WmNFVXdsVUxVbkxWQWdCZ1J3UXRBUGdWOVNGRW9UNEVnQ0kxM1hEa0EyZ0tzTzdiK3NtaUtZVExMYVFrYUFFQXNDT0NGZ0Q4aXZvUUluSFZBb0FTalNHTGhjMlRzV2ZWT0JUR1JRdFN1akI5QUlEZEVMUUE0S2ZVaHhDTW9BVUFSWmw2MVArd2RUSzM2dHY2eFpJcGlJc1dwSFJ1K2dBQXV5Rm9BY0R2cUE4aGl1T3orMXRuVGdFb2laL0R5TjFEMzlhK3ppbktGQ3g2dG5WU2FickI3OVVBQURzZ2FBSEE3L2pESjVHNGFnRkFFWnB1dVBEcWxBSXNMWmxDdVdwQlNvSVdBQUE3SUdnQndDOXRMbS9HMXpiZlRJa2dQbHNFQUlVUWRpVjNkMzFiYjJ5WlF2bmFKNlVMMHdjQStEaEJDd0JlNDZzcEVZVDZFQUN5MTNURGF2dy96NmJKMk5ZMUN3b25hRUZLZnFjR0FOZ0JRUXNBWGtQUWdralVod0NRcmFZYmpud0FUUUhXZlZ1L1dEU2w2dHRhZFFncG5VNC9id0FBOEFHQ0ZnRDhsdm9RZ2xFZkFrRE8xbFZWTFd5WWpEMzNiYjJ5WUtnZWpJQ0VYTFVBQVBnZ1FRc0FYc3RWQzZJWTYwT0VMUURJVHRNTjQ0Y2VWelpMNWx3bmczOVRIMEpLRjZZUEFQQXhnaFlBdkphZ0JaRUlXZ0NRbzdXdGtya0hsUW53SC80dGtKS0xGZ0FBSHlSb0FjQ3JxQThoR0VFTEFMTFNkTVA0eXYvY1ZzbWNheGJ3WHk1YWtKS0xGZ0FBSHlSb0FjQmJ1R3BCRkF2MUlRRGtvdW1HbzZxcVZoWks1djdzMi9ySmt1SGZwbjhQVytNZ2tVWFREU2VHRHdEd2ZvSVdBTHpGVjM4SUloQkJDd0J5c2F5cTZ0ZzJ5ZGhXbUFoK1NIMElLYWtQQVFENEFFRUxBRjV0cWc5eDFZSW9CQzBBbUwzcE5lblNKc25jcW0vckYwdUdmMUFmUWtycVF3QUFQa0RRQW9DM0VyUWdDdlVoQU9SZ2ZPVy9zRWt5OXRpMzlkcUM0WWRjdENBbEZ5MEFBRDVBMEFLQU45bGMzcWdQSVpKcjJ3QmdycHB1R0YrU1hsa2dtWE94Qlg2aWIydEJDMUk2TjMwQWdQY1R0QURnUFZ5MUlJcFBaL2UzUjdZQndFeDU1VS91dnZrZ0dYN3IwWWhJcGVrR1Z5MEFBTjVKMEFLQTl4QzBJQkwxSVFETVR0TU40MVdtVTVzamM2NVp3Tzl0eklpRUxnd2ZBT0I5QkMwQWVEUDFJUVFqYUFIQXJEVGRjT1NhQlFXNDdkdjZ5YUxodDF4OUlTVVhMUUFBM2tuUUFvRDNjdFdDS05TSEFEQTM0eXYvaGEyUnNhMHdFYnlhaXhhazVLSUZBTUE3Q1ZvQThGNkNGa1RpcWdVQXM5QjB3MGxWVlRlMlJlYVdmVnUvV0RMOFh0L1dHeGNqU2VoNHVyUUZBTUFiQ1ZvQThDNVRmY2l6NlJIRXRVVUFNQk5lK1pPN2g3NnR2OWd5dkltckZxU2tQZ1FBNEIwRUxRRDRDRmN0aU9MODdQNzJ4RFlBaUt6cGh2RTg5eWRMSW5NckM0WTMrMjVrSktRK0JBRGdIUVF0QVBnSUw5V0lSSDBJQU5INTJZbmMzZlZ0N1FOamVEdi9ia2hKMEFJQTRCMEVMUUI0dDgzbHpVWjlDSUdvRHdFZ3JLWWJsbU1QdWcyUnNhMXJGdkJ1cWtOSVNYVUlBTUE3Q0ZvQThGSHFRNGppVkgwSUFCRTEzWERrQTJnS3NPN2Irc21pNGUzNnRuN3hpSUdFRmswMytGMGFBT0NOQkMwQStDZ25zSWxFZlFnQUVZMGhpNFhOa0xIeEErSzFCY09IcUE4aEpmVWhBQUJ2SkdnQndJZW9EeUVZOVNFQWhOSjB3M2lPK3c5YklYUEw2VVUrOEg3cVEwaEpmUWdBd0JzSldnQ3dDK3BEaUVKOUNBRFJlT1ZQN2g3NnR2YjdBSHljaXhhazVLSUZBTUFiQ1ZvQXNBdnFRNGhrYVJzQVJOQjB3MWhwZFc0WlpNN1BYckFEZlZ1N2FFRktwNllQQVBBMmdoWUFmSmo2RUlMNWJDRUFCT0dhQmJtNzgrRXc3TlNEY1pKSzB3MnVXZ0FBdklHZ0JRQzc0b01Fb2pnK3U3L1ZMd3RBVWswM3JNYi9rMnlCakcxZHM0Q2RFMXdpSmI5SEF3QzhnYUFGQUx1aWw1bElybTBEZ0ZTYWJqanlBVFFGV1BWdC9XTFJzRlBmalpPRVhMUUFBSGdEUVFzQWRtSnplZk5VVmRXamFSS0UraEFBVWhvdmZTMXNnSXc5OTIzdG9oM3Nub3NXcE9TaUJRREFHd2hhQUxCTFgweVRJTlNIQUpERTFHOStaZnBrenZVdzJJTytyY2NIRE05bVN5TEgwMVV1QUFCZVFkQUNnRjFTSDBJa1BnQUFJSVdWcVpPNWg3NnQxUnZBL3JocVFVcnFRd0FBWGtuUUFvQ2RVUjlDTUlJV0FCeFUwdzNqL3ozbnBrN20vSXdGK3lWb1FVb3VRd0lBdkpLZ0JRQzdwajZFS0Jabjk3ZWZiUU9BUTVoT2JidG1RZTcrbktvTmdQMXhNWWFVWExRQUFIZ2xRUXNBZGsxOUNKRUlXZ0J3S011eDI5eTB5ZGhXbUFqMlR6VVBpYmxvQVFEd1NvSVdBT3lVK2hDQ0ViUUFZTythYmppWmdoYVFzMlhmMWk4MkRBZmhkMnBTV1RUZElHd0JBUEFLZ2hZQTdNUGFWQWxDZlFnQWh6RCs3TE13YVRMMjJMZTFpa0E0SEZjdFNFblFBZ0RnRlFRdEFOZ0g5U0ZFSW1nQndONDAzVEIybVg4eVlUTG5ZZ3NjMXNhOFNlakM4QUVBZmsvUUFvQ2QyMXplakNlRnY1a3NRWHcrdTc4OXNnd0E5c1FsTDNMM3JXOXJyK3Zoc1B5Ykl5VVhMUUFBWGtIUUFvQjljZFdDS0JhdVdnQ3dEMDAzWEZkVmRXcTRaTTQxQ3ppd3ZxMmZxcXJhbWp1SitOa0dBT0FWQkMwQTJCZEJDeUlSdEFCZ3A1cHVPSExOZ2dMY1RoLzRBb2VuUG9Sa3BtbzBBQUIrUWRBQ2dMMVFIMEl3bjlTSEFMQmp5K2xxRXVUcVdaZ0lrbElmUWtxQ0ZnQUF2eUZvQWNBK3VXcEJKSzVhQUxBVFRUZWNWRlYxWTVwa2J0VzM5WXNsUXpLQ0ZxUjBadm9BQUw4bWFBSEFQZ2xhRUltZ0JRQzc4c1VreWR4RDM5YSt6aUV0MVNHa0pHZ0JBUEFiZ2hZQTdJMzZFSUpSSHdMQWgwMmQ1ZWNtU2VaV0ZneHBUUmRsSHEyQlJJNm5DMTRBQVB5RW9BVUErK2FxQlpHNGFnSEFSM25sVCs3dStyWldXUUF4dUdwQlNxNWFBQUQ4Z3FBRkFIdTF1YndaUDR6WW1qSkJMQzBDZ1BkcXVtSDhmK1RZQU1uWTFqVUxDRVhRZ3BRRUxRQUFma0hRQW9CRGNOV0NLRTdQN20rZFB3WGd6WnB1T1BJQk5BVlk5MjM5Wk5FUWh1c3lwSFJoK2dBQVB5ZG9BY0FoQ0ZvUWlmb1FBTjVqWFZYVnd1VEkySFBmMXNKRUVFamYxaTVha05LNTZRTUEvSnlnQlFCN3Q3bTgrYW8raEVDdUxRT0F0Mmk2WVR5ZGZXVm9aRTdGR3NUMFlDK2tNdjBNQkFEQUR3aGFBSEFvcmxvUWhmb1FBTjVxYldKazdxRnZheit2UTB6cVEwaEpmUWdBd0U4SVdnQndLUDV3U3lUcVF3QjRsYVliUGp1ZFRRRmNzNEM0MUllUWtvc1dBQUEvSVdnQndFR29EeUVZOVNFQXZKWnJGdVR1cm05ckgrUkNYUDU5a3BLZ0JRREFUd2hhQUhCSXJsb1F4VmdmNGc5R0FQeFMwdzJycXFxT1RZbU1iVjJ6Z05qNnRuNnFxdXJabWtqa3RPbUdJOE1IQVBnblFRc0FEa25RZ2toY3RRRGdwNXB1T1BFQk5BVlk5VzM5WXRFUW5xc1dwT1NSQWdEQUR3aGFBSEF3VTMySWx6aEU4ZGttQVBpRjhackZ3b0RJMkhQZjFxcHhZQjYrMnhNSlhSZytBTUEvQ1ZvQWNHaXVXaERGc2ZvUUFINms2WWJ4QTRVcnd5RnpybnZCZkxob1FVcCtid1lBK0FGQkN3QU83WXVKRTRnUEdBRDRrWldwa0xtSHZxMjlrSWVaOE8rVnhGeTBBQUQ0QVVFTEFBNXFjM216VVI5Q0lPcERBUGlicGh2R0VONjVxWkE1WVZPWW4wYzdJNUZGMHcwbmhnOEE4SGVDRmdDa29ENkVLTWI2RUs5ekFQaVhwaHVPWExPZ0FMZDlXejlaTk15T3F4YWtwRDRFQU9CL0NGb0FrSUw2RUNMeG9oT0EvN2NjUTNpbVFjYTJWVld0TFJobWFXTnRKT1NCQWdEQS94QzBBT0RnMUljUWpQb1FBS3JwSlBhTlNaQzVaZC9XTDVZTXMrU2lCU201YUFFQThEOEVMUUJJUlgwSVVTek83bStGTFFEd3lwL2NQZlp0N2JJY3pOUlUrYk8xUHhJNU4zZ0FnTDhUdEFBZ0ZYL2tKUkpCQzRDQ05kMHduc1ArNUd1QXpDMHRHR2JQVlF1U21YNWVBZ0JnSW1nQlFCSlRmY2lqNlJPRW9BVkEyVnl6SUhmZityYjJBUzNNMzhZT1NVaDlDQURBWHdoYUFKQ1NxeFpFb1Q0RW9GQk5OMXhYVlhWcS8yUnM2NW9GWkVOZ2lwUUVMUUFBL2tMUUFvQ1V2cG8rZ1FoYUFCU202WVlqMXl3b3dMcHY2eWVMaGl5NGFFRktxa01BQVA1QzBBS0FaRGFYTjAvcVF3ams2dXorOXNoQ0FJcXlHcThhV1RrWmV4WW1nbnowYmYzaWQyZ1NPcDVDcWdBQXhhc0VMUUFJUUgwSWtiaHFBVkNJcGh0T3FxcjZ3NzdKM0dyNllCYkloNnNXcEtRK0JBQmdJbWdCUUdycVE0aEUwQUtnSE1LZTVPNmhiMnRmNTVDZjczWktRdXBEQUFBbWdoWUFKS1UraEdBK3FROEJ5Ri9URGVPSEJPZFdUZVpXRmd4WmN0R0NsQVF0QUFBbWdoWUFST0NsSFpHNGFnR1FQejk3a0x1N3ZxMjllb2NNOVcwdGFFRktxa01BQUNhQ0ZnQkVvRDZFU0FRdEFETFdkTVA0eXYvWWpzbll0cXFxcFFWRDFoNnNsMFFXVFRlY0dENEFnS0FGQUFGTTlTSCtVRVFVNmtNQU10VjB3NUVQb0NuQXVtL3JGNHVHckxsWVEwcnFRd0NBNGxXQ0ZnQUU0b1Eza1Z6YkJrQ1cxdU5MVEtzbFk4OTlXNjhzR0xLblBvU1UxSWNBQU1XckJDMEFDRVI5Q0pFSVdnQmtwdW1HOFVPQkszc2xjeTYyUUJsY3RDQWxGeTBBZ09KVmdoWUFSTEc1dkJuUEczK3pFSUk0UGJ1LzFUc0xrSmUxZlpLNWg3NnRoWmVoQUZNOTBMTmRrOGlwd1FNQUNGb0FFSXMvREJQSlo5c0F5RVBURGVQMzlIUHJKSE91V1VCWjFJZVFUTk1OcmxvQUFNVVR0QUFnRWtFTElsRWZBcENCcGh1T1hMT2dBSC8yYmUxRFZ5aUwraEJTT2pOOUFLQjBnaFlBaEtFK2hHRFVod0RrWVh6bGYyeVhaR3hiVmRYS2dxRTR3bFdrNUtJRkFGQThRUXNBb25IVmdralVod0RNV05NTkorb1VLTUNxYitzWGk0YXk5RzN0b2dVcHVXZ0JBQlJQMEFLQWFBUXRpTVNIY3dEek5yN3lYOWdoR1h2dTIxbzFEcFRyd2U1SjVIaXFad01BS0phZ0JRQ2hxQThobU9Peisxc3ZkUUJtcU9tRzhhVDFsZDJSdVdzTGhxS3BEeUVsOVNFQVFORUVMUUNJNkl1dEVJZ1BNQURtYVdWdlpPNmI2Z0FvbnFBRktYbVVBQUFVVGRBQ2dIQTJsemRqZmNqV1pnamlzMFVBekV2VERXTkk3dHpheUp5S00wRFlpcFJjdEFBQWlpWm9BVUJVWDIyR0lOU0hBTXpJMUJlK3RqTXlkOXUzOVpNbFE5bW03d01lS1pDS1VDc0FVRFJCQ3dDaUVyUWdFdlVoQVBNeHZ2SmYyQmNaMndvVEFYL2hxZ1hKTk4zZ1VRSUFVQ3hCQ3dCQ1VoOUNNT3BEQUdhZzZZYVRxcXB1N0lyTUxmdTJmckZrWUxJeENCSVN0QUFBaWlWb0FVQmtybG9ReFZnZkltd0JFSjlYL3VUdXNXL3JMN1lNL0lXTEZxUjBZZm9BUUtrRUxRQ0lUTkNDU0FRdEFBSnJ1bUg4US84bk95SnpTd3NHL3FwdmEwRUxVbkxSQWdBb2xxQUZBR0dwRHlFWVFRdUEyRnl6SUhkM1BsQUZmdUxSWUVqazFPQUJnRklKV2dBUW5hc1dSTEZRSHdJUVU5TU5TMy9vSjNOaitIaGx5Y0JQYkF5R1ZLYXJZZ0FBeFJHMEFDQTZIZFJFSW1nQkVFelREVWMrZ0tZQTY3NnRueXdhK0FuWGJraEowQUlBS0pLZ0JRQ2hiUzV2eGo4WVBkc1NRUWhhQU1Remhpd1c5a0xHbmxYakFML2hvZ1VwblprK0FGQWlRUXNBNWtCOUNGR29Ed0VJcE9tR2s2cXEvckFUTXJmcTIvckZrb0dmNmR0Nk0xVU1RUXFDRmdCQWtRUXRBSmdEOVNGRWNtMGJBR0g0R1lIY1BmUnQ3ZXNjZUExWExVamxlQXEvQWdBVVJkQUNnUEEybHpjYjlTRUU4dW5zL3ZiSVFnRFNhcnBodkRCMGJnMWtibW5Cd0N0OU55Z1NjdFVDQUNpT29BVUFjNkUraEVqVWh3Q2t0N1lETW5jMzFRRUF2SWJ2RjZSMFlmb0FRR2tFTFFDWUN5ZVRpVVRRQWlDaHBodFc0NWxxT3lCalc5Y3NnRGR5MFlLVVhMUUFBSW9qYUFIQUxLZ1BJUmoxSVFDSk5OMXc1QU5vQ3JEdTIvckZvb0hYbXI1bitKMlpWTlM1QVFERkViUUFZRTdVaHhDSnF4WUFhWXlWSVF1ekoyUFBmVnV2TEJoNEIxY3RTS2JwQmxjdEFJQ2lDRm9BTUNmcVE0aEUwQUxnd0tZLzRGK1pPNWx6c1FWNHI0M0prZENGNFFNQUpSRzBBR0EycHZxUVJ4c2ppTEUrNU1ReUFBNXFiZHhrN3FGdmExZmNnUGNTdENBbEZ5MEFnS0lJV2dBd042NWFFSW1yRmdBSDBuVER0ZjV2Q25CdHljQjc5VzJ0T29TVUJDMEFnS0lJV2dBd04xNzRFWWtQUXdBT29PbUdvNnFxVm1aTjV2N3MyL3JKa29FUGVqQkFFam1kZm1ZREFDaUNvQVVBczdLNXZIbFNIMElncCtwREFBNWlXVlhWc1ZHVHNhMHdFYkFqNmtOSXlWVUxBS0FZZ2hZQXpKSDZFQ0pSSHdLd1IwMDNuRXhCQzhqWnFtL3JGeHNHZGtCOUNDbGRtRDRBVUFwQkN3RG1TSDBJa2FnUEFkaXY4Wlgvd296SjJHUGYxbXNMQm5iRVJRdFNjdEVDQUNpR29BVUFzNk0raEdEVWh3RHNTZE1ONDZ2SUsvTWxjeTYyQUR2VHQvWDQrL0t6aVpLSWl4WUFRREVFTFFDWUsvVWhST0lERW9EOThNcWYzSDNyMjlxWmYyRFhYTFVnbGNWVSt3WUFrRDFCQ3dEbVNuMElrWHkyRFlEZGFycGhyR1k2TlZZeUo2d0o3SU9nQlNtNWFnRUFGRUhRQW9CWm11cER2dGtlUVJ5ZjNkL3FvZ1hZa2FZYmpseXpvQUMzMDRsL2dGMXpLWWVVL0c0TUFCUkIwQUtBT1hQVmdraXViUU5nWjhaWC9ndmpKR05iWVNKZ1gxUVNrWmlnQlFCUUJFRUxBT1pNMElKSTFJY0E3TURVNjMxamxtUnUyYmYxaXlVRGUvUm91Q1J5YnZBQVFBa0VMUUNZcmMzbHpZdjZFQUpSSHdLd0cxL01rY3c5OUczdDZ4ellOMWN0U0ticGhndlRCd0J5SjJnQndOeTVha0VrNmtNQVBtRDZvN3hYa09SdVpjUEFBV3dNbVlROFFnQUFzaWRvQWNEY0NWb1FpYUFGd01kNDVVL3U3dnEyOXNvY09BUkJDMUlTdEFBQXNpZG9BY0NzcVE4aG1NWFovZTFuU3dGNHU2WWJsbU1OazlHUnNhMXJGc0NoOUcyOW1iN3ZRQXFxUXdDQTdBbGFBSkFEVnkySVJOQUM0STJhYmpqeUFUUUZXUGR0L1dUUndBRzVha0VxeDlQUGR3QUEyUkswQUNBSGdoWkVJbWdCOEhaanlHSmhibVRzZVF4YVdEQndZS3FLU0VsOUNBQ1FOVUVMQUdadnFnKzVzMG1DVUI4QzhBWk5ONHgvaFAvRHpNamNzbS9yRjBzR0RrelFncFRVaHdBQVdSTzBBQ0FYcmxvUWlhQUZ3T3Q1NVUvdUh2cTI5ck1xa0lMcUVGSVN0QUFBc2lab0FVQVdOcGMzNHgrdnQ3WkpFSUlXQUsvUWRNUDQvZkxjck1qYzBvS0JGS1pMT3MrR1R5S3FRd0NBckFsYUFKQVRMd1dKWXF3UHViWU5nTjl5ellMYzNmVnQ3VVU1a0pMNkVGSlpUQlZ4QUFCWkVyUUFJQ2VDRmtUaXFnWEFMelRkc0txcTZ0aU15TmpXTlFzZ0FHRXZVaEswQUFDeUpXZ0JRRGJVaHhETXA3UDcyeU5MQWZpbnBodU9mQUJOQWRiVDJYNkFsRnkwSUNWQkN3QWdXNElXQU9UR1ZRc2ljZFVDNE1mR3lwQ0YyWkN4NTc2dFZ4WU1wS2EraU1RdUxBQUF5SldnQlFDNUViUWdFa0VMZ1AvUmRNUDRCL2NyY3lGejF4WU1CUEpnR1NSeWF2QUFRSzRFTFFESWl2b1FnbEVmQXZCUFh2bVR1NGUrclozcUJ5SngxWUprcHBBdEFFQjJCQzBBeUpHckZrVGlxZ1hBcE9tRzhaWC91WG1RT2Rjc2dHaUV2MGpwelBRQmdCd0pXZ0NRbzdXdEVzalNNZ0QrRmJJNGNzMkNBdnpadC9XVFJRUEJ1R2hCU2k1YUFBQlpFclFBSUR1Ynk1dnhqMGpQTmtzUXAyZjN0eWVXQWZDdjRObXhNWkN4clRBUkVORVVBUE03TXFtNGFBRUFaRW5RQW9CY3FROGhFdlVoUU5HYWJqaHg0WWNDclBxMmZyRm9JQ2hYTFVqbGVQcFpFQUFnSzRJV0FPVHFpODBTaUs1Mm9IUmpyZGVpOUNHUXRjZStyZFhYQVpGOXR4MFNjdFVDQU1pT29BVUFXVklmUWpEcVE0QmlOZDB3OW5KLzhoVkE1bHhzQWFKejBZS1VCQzBBZ093SVdnQ1FNL1VoUktJK0JDaVZWLzdrN2x2ZjFsNktBNkg1UGtWaUZ4WUFBT1JHMEFLQW5La1BJUkwxSVVCeG1tNFl2L2VkMmp5WmM4MENtSXRIbXlLUmM0TUhBSElqYUFGQXR0U0hFTXhZSCtKY0tsQ01waHVPWExPZ0FMZDlXejlaTkRBVHJscVFUTk1OZmg4R0FMSWlhQUZBN3RTSEVJbXJGa0JKVmxWVkxXeWNqRDBMRXdFenM3RXdFaEswQUFDeUltZ0JRTzc4OFp0SVB0c0dVSUttRzA2cXF2ckRzc25jcW0vckYwc0dac1JGQzFLNk1IMEFJQ2VDRmdCa2JYTjU4NlNIbGtDTzFZY0FoZmhpMFdUdW9XOXJYK2ZBckV4VlIxdGJJeEcvQ3dNQVdSRzBBS0FFL2doT0pPcERnS3cxM1RDK1ZqeTNaVEszc21CZ3B0U0hrTXFweVFNQU9SRzBBS0FFWDIyWlFOU0hBTGtUY0NSM2QzMWJPNzhQekpYdlh5UXpCWElCQUxJZ2FBRkE5dFNIRUl6NkVDQmJUVGNzeCs5ek5rekd0cTVaQURNbmFFRktnaFlBUURZRUxRQW9oZGUxUkxLMERTQTNUVGNjK1FDYUFxejd0bjZ5YUdER1ZJZVFra2NIQUVBMkJDMEFLSVg2RUNKUkh3TGthRjFWMWNKbXlkaHozOWJDUk1DczlXMzk0dUlqQ2Jsb0FRQmtROUFDZ0NLb0R5R1l4ZG45cmJBRmtJMm1HOGJYaVZjMlN1WmNwQUp5NGFvRnFTeWFiamd4ZlFBZ0I0SVdBSlJFZlFpUkNGb0FPVm5iSnBsNzZOdmFoVFFnRjRJV3BLUStCQURJZ3FBRkFDWHh4M0VpRWJRQXN0QjB3L2o5N053MnlaeHJGa0JPdnRzbUNha1BBUUN5SUdnQlFER20rcEJ2Tms0UTZrT0FYTGhtUWU3dStyYjIraHZJaHU5cEpPYWlCUUNRQlVFTEFFcmpxZ1dSQ0ZvQXM5WjB3NnFxcW1OYkpHTmIxeXlBVEQxWUxJbTRoQVlBWkVIUUFvRFNDRm9ReWRYWi9lMlJqUUJ6MUhURGlRK2dLY0NxYitzWGl3WXlwRDZFWkpwdWNOVUNBSmc5UVFzQWlySzV2SGxSSDBJd3Jsb0FjelZlczFqWUhobDc3dHRhTlE2UUsvVWhwSFJoK2dEQTNBbGFBRkFpVnkySVJOQUNtSjJtRzhZL2psL1pISm03dG1BZ1l5NWFrSktMRmdEQTdBbGFBRkFpUVFzaSthUStCSmlobGFXUnVZZStyWDBJQ1dScnFrVjZ0bUVTRWJRQUFHWlAwQUtBNHFnUElTQlhMWURaYUxwaGZPVi9ibU5rempVTG9BVHFRMGpsdE9rR0R3NEFnRmtUdEFDZ1ZLNWFFSW1nQlRBTDB4L0VYYk1nZDMvMmJmMWt5MEFCWE80aEpWY3RBSUJaRTdRQW9GU0NGa1NpUGdTWWkyVlZWY2UyUmNhMndrUkFRVnkwSUtVTDB3Y0E1a3pRQW9BaVRmVWhkN1pQSUU2VUE2RTEzWEJTVmRXTkxaRzVaZC9XTDVZTWxLQnZheGN0U0VuUUFnQ1lOVUVMQUVybXFnV1JDRm9BMGExdGlNdzk5bTM5eFpLQndqeFlPSW1vRGdFQVprM1FBb0JpYlM1dnZrN25vU0dDMDdQNzJ4T2JBQ0pxdW1GOGNmakpjc2pjMG9LQkFxa1BJWlhGZERFTkFHQ1dCQzBBS0oyckZrVHkyVGFBb0Z5eklIZmZuTkFIQ2lWb1FVcnFRd0NBMlJLMEFLQjBnaFpFb2o0RUNLZnBodkdWLzZuTmtMR3RheFpBd1lUTVNFbDlDQUF3VzRJV0FCUk5mUWpCcUE4QlFtbTY0YWlxcXBXdGtMbDEzOVpQbGd5VWFQcis1M2RpVWhHMEFBQm1TOUFDQUZ5MUlCYjFJVUFrWThoaVlTTms3RmsxRG9DckZpUnpidlFBd0Z3SldnQ0FvQVd4T0YwT2hOQjB3M2hoNXcvYklIT3J2cTFmTEJrbzNLYjBBWkJPMHcwWHhnOEF6SkdnQlFERlV4OUNNTWRuOTdmT3B3SVJmTEVGTXZmUXQ3V3Zjd0FYTFVqTDc3OEF3Q3dKV2dEQXYva2pPNUZjMndhUTB2U3kwQ2xuY3JleVlZQi9jZEdDbEFRdEFJQlpFclFBZ0g4VHRDQ1N6N1lCSk9iL1JYSjMxN2UxRjl3QVZWVk5GVXFQWmtFaXFrTUFnRmtTdEFDQWY5ZUhqQzk0bnMyQ0lOU0hBTWswM1RDKzhqKzJBVEkyVnNZdExSamdiMXkxSUpYanBodU9UQjhBbUJ0QkN3RDRyNjltUVNEcVE0Q0RtLzdJN1FOb2NyZWVYbThEOEYrdS9KQ1NxeFlBd093SVdnREFmem1UVGlUcVE0QVUxbFZWTFV5ZWpEMzNiYjJ5WUlCL2NOR0NsRngwQkFCbVI5QUNBQ2JxUXdobXJBL3hxZ2M0bUtZYnhqOXdYNWs0bVhPeEJlQUgrcllXdENBbHYvc0NBTE1qYUFFQWY2YytoRWpVaHdDSHREWnRNdmZRdDdXZjlRQis3c0ZzU01SRkN3QmdkZ1F0QU9EdjFJY1FpZm9RNENDYWJoaS8zNXliTnBsenpRTGcxNzZiRDRrc3B1dHFBQUN6SVdnQkFIK2hQb1JnRm1mM3Q4SVd3RjQxM1hEa21nVUYrTk5aZklEZjhuMlNsQVF0QUlCWkViUUFnSDl5MVlKSUJDMkFmUnRmK1IrYk1obmJWbFcxc21DQTMzTFJncFFFTFFDQVdSRzBBSUIvRXJRZ0VrRUxZRythYmpoUnAwQUJWbjFidjFnMHdLOU4zeXRkZUNTVkM1TUhBT1pFMEFJQS9zZm04dWFwcXFwSGN5RUk5U0hBUG8ydi9CY21UTWFlKzdaV2pRUHdlcTVha01xcHlRTUFjeUpvQVFBLzVxb0ZrVnpiQnJCclRUZU1yd2F2REpiTStUOFU0RzAyNWtVcTA4K25BQUN6SUdnQkFELzIxVndJNU5QWi9lMlJoUUE3dGpKUU12ZlF0N1dYMlFCdkkyaEJTbWVtRHdETWhhQUZBUHlBK2hBQ1VoOEM3RXpURGVNci8zTVRKWE91V1FDOGtZQWFpYmxvQVFETWhxQUZBUHljK2hBaUViUUFkcUxwaHZGQ3p0bzB5ZHh0MzlaUGxnendMZy9HUmlJdVdnQUFzeUZvQVFBL3B6NkVTTlNIQUx1eXJLcHFZWnBrYkN0TUJQQWg2a05JNWJqcGhoUFRCd0RtUU5BQ0FINUNmUWdCdVdvQmZNajBoK3NiVXlSenk3NnRYeXdaNE4zVWg1Q1NxeFlBd0N3SVdnREFyNmtQSVJKQkMrQ2p2UEluZDQ5OVcvdjVEZUJqWExRZ0pVRUxBR0FXQkMwQTROZjhvWjVJeHZvUVoxU0JkMm02NFdMOFBtSjZaRzVwd1FBZjA3ZjEwMVREQkNsY21Eb0FNQWVDRmdEd0M1dkxtL0hzOURjekloQlhMWUQzRWg0a2QzZDlXenQzRDdBYnZwK1N5cm5KQXdCeklHZ0JBTC8zMVl3STVOb3lnTGRxdW1GODVYOXNjR1JzZkhtOXNtQ0FuVkVmUWpKTk42Z1BBUURDRTdRQWdOOFR0Q0NTVS9VaHdGczAzWERrQTJnS3NKNU8zUU93R3k1YWtKS2dCUUFRbnFBRkFQeUcraEFDVWg4Q3ZNVVlzbGlZR0JsN0hvTVdGZ3l3TzZxWVNPekNBZ0NBNkFRdEFPQjFYTFVnRXZVaHdLczAzVEJld1BuRHRNamNxbS9yRjBzRzJMbEhJeVVSRnkwQWdQQUVMUURnZFFRdGlFUjlDUEJhWDB5S3pEMzBiZTNySEdBL051WktJcWRUL1IwQVFGaUNGZ0R3Q3VwRENNaFZDK0NYbW00WWE0Yk9UWW5NTFMwWVlHL1VoNUNTcXhZQVFHaUNGZ0R3ZXE1YUVJbWdCZkE3YXhNaWMzZDlXM3R0RGJBL3ZzZVMwb1hwQXdDUkNWb0F3T3VOUVl1dGVSSEU4ZG45clJjK3dBODEzYkFhdjArWURobmJ1bVlCc0Y5VG1NM3Z3S1RpOTEwQUlEUkJDd0I0cGFrK3hGVUxJbkhWQXZpSHFjL2FCOURrYnQyMzlZc3RBK3lkcXhhazRxSUZBQkNhb0FVQXZJMmdCWkY4dGczZ0I4YktrSVhCa0xIbnZxMVhGZ3h3RU4rTm1VUVdUVGVjR0Q0QUVKV2dCUUM4d2VieVJuMElrYWdQQWY2bTZZYnhlOEtWcVpBNUYxc0FEc2RGQzFMeSt5NEFFSmFnQlFDOG5hc1dSS0krQlBpcnRXbVF1V0l3cm40QUFDQUFTVVJCVkllK3JmMHNCbkE0TGxxUWt2b1FBQ0FzUVFzQWVEdC8zQ2NTUVF2Z1g1cHVHTDhmbkpzR21mUC9Ic0FCOVczOU1sWTJtVG1KdUdnQkFJUWxhQUVBYjZRK2hHQVdaL2UzbnkwRnl0WjB3MUZWVmF2UzUwRDIvdXpiK3NtYUFRN09WUXRTRVNJR0FNSVN0QUNBOTNIVmdrZ0VMWUJsVlZYSHhVK0JuRzJGaVFDUzJSZzlxVFRkNEtvRkFCQ1NvQVVBdkkrZ0JaRUlXa0RCbW00NG1ZSVdrTFBWZEw0ZWdNTnowWUtVTGt3ZkFJaEkwQUlBM2tGOUNNR29ENEd5amEvOEY2VVBnYXc5OTIyOXRtS0FOUHEyZHRHQ2xGeTBBQUJDRXJRQWdQZjdZbllFSW1nQkJXcTZZWHpoZDJYM1pPN2FnZ0dTZTdBQ0VuSFJBZ0FJU2RBQ0FONVAwSUpJQkMyZ1RGNzVrN3R2ZlZzN1dRK1FucXNXcEhMY2RNT1I2UU1BMFFoYUFNQTdiUzV2eGo4MFBac2ZRWXoxSVY3OFFrR2FiaGovelovYU9abGJXakJBQ0VKdnBLUStCQUFJUjlBQ0FEN21xL2tSaUtzV1VJanBWWjlyRnVUdXRtL3JKMXNHQ01GRkMxSlNId0lBaENOb0FRQWZvejZFU0Q2ZDNkODZxUXBsR0YvNUwreWFqRzJGaVFEaW1JSnZManFTaXFBRkFCQ09vQVVBZklENkVBSnkxUUl5MTNURFNWVlZOL1pNNXBaOVc3OVlNa0FvcmxxUWl1b1FBQ0FjUVFzQStEajFJVVFpYUFINWMwMkozRDMyYmUzckhDQWVRUXRTV1V4aFl3Q0FNQVF0QU9EamZCQkFKT3BESUdOTk40eG5rOC90bU13dExSZ2dwTy9XUWtMcVF3Q0FVQVF0QU9DRDFJY1FrS3NXa0MvaFBuSjMxN2UxRC9JQUF2TDltY1RVaHdBQW9RaGFBTUJ1ck0yUlFLNHRBL0xUZE1QNHl2L1lhc25ZdHFxcWxRVURoUFpvUFNRaWFBRUFoQ0pvQVFDNzhkVWNDZVQ4N1A1V2Z5MWtwT21HSXg5QVU0QjEzOVpQRmcwUW1xc1dwS0krRHdBSVJkQUNBSFpnYzNuejVHVVB3YWdQZ2J5TUlZdUZuWkt4WnhmQ0FHWmhZMDJrMG5URGhlRURBRkVJV2dEQTd1ak5KeEwxSVpDSnBodkdNOGwvMkNlWlcvWnQvV0xKQU9FSldwQ1MraEFBSUF4QkN3RFlIZlVoUkhLcVBnU3k0WlUvdVh2bzI5clBVUUF6MExmMUdMVFkyaFdKdUdnQkFJUWhhQUVBTzZJK2hJRFVoOERNTmQzd1dSODFCVmhhTXNDc3VHcEJLaTVhQUFCaENGb0F3RzZwRHlFUzlTRXdmNjVaa0x1NzZYVTBBUFB4M2E1STVManBoaVBEQndBaUVMUUFnTjF5OXBwSXh2b1FMMzVncHBwdVdJMS9UTFkvTXJaMXpRSmdsZ1F0U0VsOUNBQVFncUFGQU95UStoQUNjdFVDWm1oNnFlY0RhSEszN3R2NnhaWUJac2NsSWxMeW1BQUFDRUhRQWdCMno1bDNJdmxzR3pCTDQvOGxDNnNqWTg5OVc2OHNHR0IrcHBDY0J3YWs0cUlGQUJDQ29BVUE3Sjc2RUNJNVZoOEM4OUowdy9qSDR5dHJJM011TGdITW02c1dwT0wzV3dBZ0JFRUxBTml4emVYTitMcm5tN2tTaUErellGNjg4aWQzRDMxYjYvY0htRGRCQzFKWk5OMGdiQUVBSkNkb0FRRDc0YW9Ga2FnUGdabG91bUVNUnAzYkY1a1RBQVNZUDRFNVVoSzBBQUNTRTdRQWdQMFF0Q0FTOVNFd0EwMDNITGxtUVFIKzdOdjZ5YUlCNXExdmF4Y3RTTW52dHdCQWNvSVdBTEFINmtNSWFHa3BFTjc0Ny9UWW1zallWcGdJSUNzUDFra2lGd1lQQUtRbWFBRUErK09xQlpHb0Q0SEFtbTQ0cWFycXhvN0kzS3B2NnhkTEJzaUcraEJTT1RWNUFDQTFRUXNBMkI5QkN5SlpuTjNmQ2x0QVhHdTdJWE9QZlZ2N09nZklpL29Ra21tNndWVUxBQ0FwUVFzQTJCUDFJUVFrYUFFQlRYOGsvbVEzWkU2RkZVQitCQzFJU2RBQ0FFaEswQUlBOXN0VkN5SVJ0SUNZdlBJbmQ5LzZ0blplSGlBemZWcy9WVlgxYks4a2NtYndBRUJLZ2hZQXNFZWJ5NXN2VlZWdHpaZ2cxSWRBTUUwM1hPdVlwZ0N1V1FEa3kxVUxVaEcwQUFDU0VyUUFnUDF6MVlKSUJDMGdpS1liamx5em9BQzMwNHRuQVBMa1loR3BIRGZkY0dMNkFFQXFnaFlBc0grQ0ZrVHkrZXorOXNoR0lJVFZlR25HS3NqWVZwZ0lJSHN1V3BDU3F4WUFRREtDRmdDd1o1dkxtNi9xUXdoazRhb0ZwRGU5dnZ2REtzamNzbS9yRjBzR3lGZmYxaTVha0pLZ0JRQ1FqS0FGQUJ5R3F4WkVJbWdCNlgyeEF6TDMwTGUxcjNPQU1qemFNNGxjR0R3QWtJcWdCUUFjaHFBRmtYeFNId0xwTk4wdy9rSDQzQXJJM01xQ0FZcmhxZ1dwK0prYUFFaEcwQUlBRGtCOUNBRzVhZ0hwZU9WUDd1NmNrZ2NveXNhNlNhWHBCdlVoQUVBU2doWUFjRGl1V2hDSm9BVWswSFREc3FxcVk3TW5ZMXZYTEFDS0kxeEhTdXBEQUlBa0JDMEE0SEFFTFloRWZRZ2NXTk1OUno2QXBnRHJ2cTJmTEJxZ0hOUDNmUmNjU2NWRkN3QWdDVUVMQURpUXFUN2syYndKeEZVTE9LeDFWVlVMTXlkanozMWJDeE1CbEVsOUNLa0lXZ0FBU1FoYUFNQmh1V3BCSkV2YmdNT1l1cU92akp2TStYOEZvRnpxUTBqbGRMb2NCd0J3VUlJV0FIQllYOHliUUU3UDdtOVBMQVFPWW0zTVpPNmhiMnVCVW9CeUNWcVFrcXNXQU1EQkNWb0F3QUZ0TG04MjZrTUlSbjBJN0ZuVERlTy9zM056Sm5PdVdRQ1VUWFVJS1YyWVBnQndhSUlXQUhCNFhuc1N5YlZ0d042NVprSHU3dnEyOWdFYlFNSDZ0bjZwcXVyUjF3Q0p1R2dCQUJ5Y29BVUFISjc2RUNKUkh3SjcxSFREcXFxcVl6TW1ZMXZYTEFDWUNOMlJpb3NXQU1EQkNWb0F3SUdwRHlFZzlTR3dCMDAzblBnQW1nS3NwbGZNQVBDOStBbVF5bUw2MlJzQTRHQUVMUUFnRGZVaFJPS0RZTmlQOFpyRndtekoySFBmMXFweEFQaC9MbHFRa3ZvUUFPQ2dCQzBBSUEzMUlVUnlmSFovNjQ5U3NFTk5ONHpuaTYvTWxNeGRXekFBLzY5dmEwRUxVbElmQWdBY2xLQUZBQ1NnUG9TQWZGZ0d1N1V5VHpMMzBMZTFFL0VBL0s4SEV5RVJqd2NBZ0lNU3RBQ0FkSnphSnBMUHRnRzcwWFRER0Z3Nk4wNHlKNkFId0k4STRaR0tuNzhCZ0lNU3RBQ0FkTDZhUFlHb0Q0RWRhTHJoU0pDT0F2elp0L1dUUlFQd0ErcERTR2FxN3dNQU9BaEJDd0JJWkhONU0zNUE4V2orQk9KMU1uemNzcXFxaFRtU3NhMXFIQUIrd1VVTFV2SjRBQUE0R0VFTEFFanJpL2tUaVBvUStJQ21HMDZxcXJveFF6SzM3TnY2eFpJQitKSHAvNGhud3lFUlFRc0E0R0FFTFFBZ0xmVWhSRExXaHppMUN1K25Nb1RjUGZadExTUUt3TytvRHlFVnY4OENBQWNqYUFFQUNha1BJU0QxSWZBT1V4LzBKN01qYzBzTEJ1QVYxSWVReW5IVERVZW1Ed0FjZ3FBRkFLVG5aU2lScUErQjkzSE5ndHg5Njl2YUIyY0F2SWFMRnFTa1BnUUFPQWhCQ3dCSVQzMElrU3pPN20rRkxlQU5tbTRZWC9tZm1obVpjODBDZ0ZjUnpDTXg5U0VBd0VFSVdnQkFZbE45eUlNOUVJaWdCYnpTZEpwNFpWNWs3clp2NnlkTEJ1QU4vSTVMS29JV0FNQkJDRm9BUUF6cVE0aEUwQUplYnd4WkxNeUxqRDJyeGdIZ0hkU0hrSXJxRUFEZ0lBUXRBQ0FHOVNGRW9qNEVYcUhwaHBPcXF2NHdLekszNnR2NnhaSUJlQ05CQzFKWlREK25Bd0RzbGFBRkFBU3d1YndaUDhENFpoY0VJbWdCditjYUVibDc2TnZhMXprQTcvSGQxRWhJZlFnQXNIZUNGZ0FRaDZzV1JISjFkbjk3WkNQd1kwMDNqSCs4UFRjZU1yZXlZQURlbzIvcnA2cXF0b1pISXVwREFJQzlFN1FBZ0RnRUxZakdWUXY0T2EvOHlkMWQzOVplSXdQd0VmNGZJUlVYTFFDQXZSTzBBSUFnMUljUWtLQUYvRURURGVNci8yT3pJV1BqQytTbEJRUHdRUnNESkpGVGd3Y0E5azNRQWdCaWNkV0NTRDZwRDRHL2E3cmh5QWZRRkdEZHQvV0xSUVB3UVM1YWtNeFU5UWNBc0RlQ0ZnQVFpNkFGMGJocUFYKzNycXBxWVNaazdMbHY2NVVGQS9CUktxaEk3TXdDQUlCOUVyUUFnRURVaHhDUW9BVk1tbTRZLzFoN1pSNWt6c1VXQUhicDBUUkp4RVVMQUdDdkJDMEFJSjR2ZGtJZ1kzM0lpWVhBdjZ5Tmdjdzk5RzN0dWhZQXU3UXhUUkp4MFFJQTJDdEJDd0FJWm5ONU0zN0FzYlVYQW5IVmd1STEzVEQrT3pndmZRNWt6elVMQUhaTmZRaXBIRGZkY0dUNkFNQytDRm9BUUV4ZWt4TEp0VzFRc3VrUHRLNVprTHMvKzdiMjZoaUFYZk4vQ3ltcER3RUE5a2JRQWdCaUVyUWdrbFAxSVJSdWZPVi9YUG9ReU5wNFNXdGx4UURzMmhUaWM3R1JWTlNIQUFCN0kyZ0JBQUdwRHlFZzlTRVVxZW1HRTNVS0ZHRFZ0L1dMUlFPd0o2NWFrSXFMRmdEQTNnaGFBRUJjcmxvUWlmb1FTalcrOGwvWVBobDc3dHRhTlE0QSsvVGRkRW5FUlFzQVlHOEVMUUFnTGtFTElsRWZRbkdhYmhoZndGM1pQSmtUcEFOZzMxeTBJSlZGMHczQ0ZnREFYZ2hhQUVCUTZrTUl5SWR4bE1ZcmYzTDMwTGUxVjhZQTdKdi9hMGhKMEFJQTJBdEJDd0NJelZVTEloRzBvQmhOTjR4Zjc2YzJUdVo4WHdkZzcvcTJmaG1ycWt5YVJDNE1IZ0RZQjBFTEFJanRpLzBReVBIWi9hM1hRR1N2NllZajF5d293RzNmMWs4V0RjQ0J1R3BCS242SEJRRDJRdEFDQUFMYlhONTg5L0tIWUx4K3BnVExzYy9acHNuWVZwZ0lnQVBiR0RpSnVGSUhBT3lGb0FVQXhLYytoRWcrMndZNWE3cmhwS3FxRzBzbWM4dnBqRHNBSElxZ0JjazAzYUErQkFEWU9VRUxBSWhQZlFpUnFBOGhkMTc1azd2SHZxMzliQUhBUWZWdHJUcUVsQVF0QUlDZEU3UUFnT0EybHpjYjlTRUVvejZFTEUwdjNUN1pMcGxiV2pBQWlUd1lQSWw0TEFBQTdKeWdCUURNZy9vUUlsRWZRcTY4OGlkMzM3d29CaUFoOVNHa0ltZ0JBT3ljb0FVQXpJTVAvNGhrckE4UnRpQXJUVGVNci95UGJaV01iVjJ6QUNBeFlUOVNPVzY2NGNUMEFZQmRFclFBZ0JsUUgwSkFnaFprbyttR282cXFWalpLNXRaOVd6OVpNZ0FKdVdoQlNxNWFBQUE3SldnQkFQT2hQb1JJQkMzSXlSaXlXTmdvR1J2RG1tc0xCaUNsS2ZDM3RRUVNFYlFBQUhaSzBBSUE1a045Q0pFczFJZVFnK21FOEIrV1NlWldmVnUvV0RJQUFhZ1BJWlVMa3djQWRrblFBZ0JtWXFvUGViUXZBaEcwSUFkQ2JPVHVvVzlyWCtjQVJLRStoRlRPVFI0QTJDVkJDd0NZRngrVUVJbWdCYlBXZE1ObmYzQ2xBQ3RMQmlBUUZ5MUlwdWtHOVNFQXdNNElXZ0RBdkh5MUx3SVo2ME91TFlRWlcxc2VtYnZyMjlvSFdnQ0U0ZjhsRWxNZkFnRHNqS0FGQU16STV2TG1TWDBJd2JocXdTdzEzVEMrOGorMlBUSzJyYXBxYWNFQUJPUjNXbEp4MFFJQTJCbEJDd0NZSC9VaFJQTHA3UDcyeUVhWWs2WWJqbndBVFFIV2ZWdS9XRFFBQWJscVFTcUNGZ0RBemdoYUFNRDhxQThoR2xjdG1KdXhNbVJoYTJUc3VXL3JsUVVERU5UR1lramtkQXBkQXdCOG1LQUZBTXlNK2hBQ0VyUmdOcHB1R0YreFhka1ltWE94QllESUJDMUl5VlVMQUdBbkJDMEFZSjdVaHhDSitoRG1aRzFiWk82aGIydlhyd0FJcTIvck1XaXh0U0VTdVRCNEFHQVhCQzBBWUo1OGdFSTBybG9RWHRNTjExVlZuZHNVbWJ1MllBQm13RlVMVW5IUkFnRFlDVUVMQUppaHFUN2t3ZTRJeEFkN2hEWjFNYTlzaWN6OTJiZjFreVVETUFQZkxZbEVYTFFBQUhaQzBBSUE1a3Q5Q0pHY245M2ZudGdJZ1MycnFqcTJJREsyRlNZQ1lFWUVMVWhsMFhTRDMxMEJnQThUdEFDQStWSWZRalRxUXdocCtrUHEwbmJJM0twdjZ4ZExCbUFtVkllUWtxc1dBTUNIQ1ZvQXdFeHRMbS9HRDFPKzJSK0JxQThocXZYNGNzMTJ5Tmh6MzlackN3WmdMcVp3NExPRmtjaVp3UU1BSHlWb0FRRHo1cW9Ga1p5cUR5R2FwaHZHMTJxZkxJYk1DYm9CTUVmcVEwaEYwQUlBK0RCQkN3Q1lOMEVMb2xFZlFqUmUrWk83YjMxYis2QUtnRGxTSDBJcTV5WVBBSHlVb0FVQXpKajZFQUx5cXBvd21tNFl2eDVQYllUTUxTMFlnSmtTRkNTWjZmSWRBTUM3Q1ZvQXdQeTVha0VrNmtNSW9lbUdJOWNzS01CdDM5WlBGZzNBSFBWdDdhSUZLYWtQQVFBK1JOQUNBT1pQMElKb3ZLNG1ndkhyY0dFVFpHd3JUQVJBQmg0c2tVUUVMUUNBRHhHMEFJQ1ptK3BEN3V5UlFENWJCaWsxM1RCZVZibXhCREszN052NnhaSUJtRGxYTFVoRmRRZ0E4Q0dDRmdDUUIxY3RpT1Q0N1A3VzZ5QlMrbUw2Wk82eGIydGY1d0RrNExzdGtzanhWRGNJQVBBdWdoWUFrSUhONWMzWDZZUTRSSEZ0RTZUUWRNUDRNdTNjOE1tY2lpWUFjdUdpQlNsNUlBQUF2SnVnQlFEa3cxVUxJbEVmUWlwZStaTzd1NzZ0dmY0RklBdDlXejlWVmZWc215U2lQZ1FBZURkQkN3REloNkFGa2FnUDRlQ2FiaGhmK1IrYlBCa2JyMWV0TEJpQXpMaHFRU3FDRmdEQXV3bGFBRUFtMUljUWtQb1FEbWJxVi9ZQk5MbGJUeTkvQVNBbmdoYWs0bkVBQVBCdWdoWUFrQmRYTFloRTBJSkRHa01XQ3hNblkrTlo5YlVGQTVBaGxWaWtzbWk2UWRnQ0FIZ1hRUXNBeUl1Z0JaRXN6dTV2UDlzSSt6YjljZlFQZ3laenE3NnRYeXdaZ056MGJTMW9RVXFDRmdEQXV3aGFBRUJHMUljUWtLQUZoK0NWUDdsNzZOdjZpeTBEa0xGSHl5VVJRUXNBNEYwRUxRQWdQNjVhRUltZ0JYdlZkTVA0TlhadXltUnVhY0VBWk01VkMxSzVNSGtBNEQwRUxRQWdQMTUyRTRuNkVQYk45enh5ZDllMzljYVdBY2ljLyt0STVkVGtBWUQzRUxRQWdNeHNMbS9HUDFBOTJ5dUJDRnF3RjAwM3JLcXFPalpkTXJaMXpRS0FRcmhvUVRKTk43aHFBUUM4bWFBRkFPUkpmUWlSZkQ2N3Z6MnlFWGFwNllZVEgwQlRnSFhmMWk4V0RVRHUrclorbWdLR2tNS1pxUU1BYnlWb0FRQjUrbUt2QkxKdzFZSTlXRTFmVzVDcjU3NnRWN1lMUUVIVWg1Q0tpeFlBd0pzSldnQkFodFNIRUpDZ0JUc3puZmE5TWxFeWQyM0JBQlJHZlFpcHVHZ0JBTHlab0FVQTVFdDlDSkY4VWgvQ0RubmxUKzRlK3JiMllSTUFwZkYvSDZrY045M2c5MVVBNEUwRUxRQWdYK3BEaU1aVkN6NnM2WWJ4bGYrNVNaSTUxeXdBS0pIcUVGSlNId0lBdkltZ0JRQmtTbjBJQVFsYThDSFRLelBYTE1qZG4zMWJQOWt5QUtYcDIvcWxxcXBIaXljUjlTRUF3SnNJV2dCQTN0U0hFSW42RUQ1cU9aNzFOVVV5dGhVbUFxQndybHFRaW9zV0FNQ2JDRm9BUU43Vzlrc3dybHJ3TGswM25GUlZkV042Wkc0MXZlWUZnRklKV3BDS2VrSUE0RTBFTFFBZ1k1dkxteWVuVndsbWFTRzhrK0FZdVh2czI5clhPUUNsKzE3NkFFaW42UWIxSVFEQXF3bGFBRUQrdnRneGdaeWUzZCtlV0FodjBYVERlTWIzazZHUk9VRTBBSXJYdDdXTEZxUWthQUVBdkpxZ0JRRGs3NnNkRTR6NkVON0tLMzl5OTYxdmF5OTRBZURmSHN5QlJDNE1IZ0I0TFVFTEFNaWMraEFDdXJZVVhxdnBodkhyNWRUQXlKeHJGZ0R3WDhLSHBPS2lCUUR3YW9JV0FGQUc5U0ZFb2o2RVYybTY0Y2cxQ3dwdzI3ZjFrMFVEd0grb0R5RVZBVzhBNE5VRUxRQ2dET3BEaUVaOUNLK3hxcXBxWVZKa2JDdE1CQUQvSUdoQk1rMDNxQThCQUY1RjBBSUFDcUEraElEVWgvQkxUVGVNVjAvK01DVXl0K3piK3NXU0FlQy9wa3RQejBaQ0lvSVdBTUNyQ0ZvQVFEblVoeERKV0IraS81WmY4VDJMM0QzMGJlM3JIQUIrekZVTFV2RjdLZ0R3S29JV0FGQU85U0ZFNDZvRlB6U2Q2ejAzSFRLM3NtQUErS252UmtNaWdoWUF3S3NJV2dCQUlhYjZrRy8yVFNDZkxZT2Y4TXFmM04zMWJlMERKQUQ0T1JjdFNPVjRxakVFQVBnbFFRc0FLSXVyRmtSeXJENkUvOVYwdzNMODJqQVlNcloxelFJQWZrMGdrY1Q4bmdvQS9KYWdCUUNVUmRDQ2FOU0g4QjlOTnh6NUFKb0NyUHUyZnJKb0FQaXRCeU1pa1F1REJ3QitSOUFDQUFxeXVieDVVUjlDTU9wRCtLdDFWVlVMRXlGano5UFhPUUR3ZStwRFNNVkZDd0RndHdRdEFLQThybG9ReVZnZjRyVVE0eldMOFkrWlZ5WkI1cFo5Vzc5WU1nQzhpcUFGcVp5YlBBRHdPNElXQUZBZVFRdWlVUjlDNVpVL0JYam8yOXIvd1FEd2V0L05pbFNtSURnQXdFOEpXZ0JBWWRTSEVKRDZrTUkxM2ZEWnF6RUtzTFJrQUhpOXZxMmZxcXJhR2htSnVMd0lBUHlTb0FVQWxNbUxXaUpabk4zZkNsc1VxdW1HSTljc0tNQmQzOWJPbndQQTI3bHFRU291V2dBQXZ5Um9BUUJsRXJRZ0drR0xjbzJ2L0k5TEh3SloyN3BtQVFEdkpxaElLb0lXQU1BdkNWb0FRSUdtK3BBN3V5Y1FRWXNDTmQxdzRnTm9DckRxMi9yRm9nSGdYVnkwSUpYVDZmb2VBTUFQQ1ZvQVFMbGN0U0FTOVNGbFdvMjdMMzBJWk8yNWIydlZPQUR3Zmk1YWtKS3JGZ0RBVHdsYUFFQ2hOcGMzWDZkejVoQ0ZvRVZCbW02NHFLcnFxdlE1a0wxckt3YUE5NXV1UWowYUlZbGNHRHdBOERPQ0ZnQlFObGN0aU9UcTdQN1dhZFp5ckVvZkFObDc2TnZhdVhNQStEaFhMVWpGUlFzQTRLY0VMUUNnYklJVy9COTc5NU1UU1pMdEM5ank2YzdocmdCNkJVbXZBSHJpMCtUSkpWZk1NbW9GRmJXQ0NGWnd5UlUwekZJdXVSNU1mZExKQ3RwWlFjRUtMcXdnbnJ6YnM1cXF5aitRUklTYm0zMmZoSHJTVWtXY0U4a2Z0NStkRXh0VExUSlExbTEveS84NDl6cVFQTk1zQUdBekJCY1ppNGtXQU1CWENWb0FRTWFzRHlGQ2doYUpLK3UybjFweW5uc2RTTjZIcGlydXRCa0FOc0pFQzhheVY5YnRvZW9EQUY4aWFBRUFtR3BCVE41Wkg1SzhSZi9BTXZjaWtMUkhxM0VBWUhPYXFoQzBZRXltV2dBQVh5Um9BUUFJV2hBYlV5MFNOZHdHVytaZUI1SzNhS3JpUVpzQllLTnVsSk9SSENrOEFQQWxnaFlBa0RuclE0aVFvRVc2ckF3aGRiZE5WVnpvTWdCczNDY2xaU1NDRmdEQUZ3bGFBQURCVkFzaVkzMUlnc3E2N1VmdXZzdTlEaVJ2b2NVQXNCWFdoekNXWTVVSEFMNUUwQUlBQ0c2WkU2RzVwaVRIOXhsU2Q5MVVoZHUyQUxBZGZzWXltaUUwRGdEd080SVdBRUMvUHFTL0hYU3ZFa1JFMENJaFpkMzJ0L3pmNWw0SGttZWFCUUJzU1ZNVkQvNW1aVVRXaHdBQWZ5Sm9BUUI4Wm4wSU1YbDc5UEhzVUVlbXI2emJmZzNNS3ZjNmtMeXpwaXJ1dEJrQXRzcjZFTVlpYUFFQS9JbWdCUUR3MllWS0VKbFREVWxDSDdMWXk3MElKTzNlYWh3QTJBbnJReGlMMVNFQXdKOElXZ0FBLzJKOUNCR3lQbVRpeXJydHA1TDhuSHNkU041cUdHY09BR3lYaVJhTTVXQ1kxQWNBOEJ0QkN3RGdLZXREaUluMUlkTm5VZzZwdTJtcXd1Y2NBSGFncVFvVExSaVRxUllBd084SVdnQUFUemtzSWphbVdreFVXYmY5ZzhqajNPdEE4bFphREFBN2RhUGNqT1JJNFFHQXB3UXRBSURmV0I5Q2hBUXRwa3R3aTlSZHVsa0xBRHRuZlFoak1kRUNBUGdkUVFzQTRJOGNqaEtUTzkyWW5ySnUrMXYrQjduWGdhUTltbVlCQUtNUWNtUXNKbG9BQUw4amFBRUEvSkdnQmJIb0R6SlBkV05heXJyZER5RXNjcThEeVR0dnFrSVFEQUIyejBRTHhySlgxcTJ3QlFEd0cwRUxBT0IzdXRteVB6aTZWUlVpY05yTmxnOGFNVG5uL1VQSTNJdEEwdTZicWpETkFnQkdNQVFkSDlXZWtRaGFBQUMvRWJRQUFMN0VWQXZHZHRiTmxzWUNUOHh3dyt0OTduVWdlU2EyQU1DNC9KM0FXQVF0QUlEZkNGb0FBRjl5cFNxTTZLYWJMZDBXbjZiejNBdEE4bTZhcXZBekVnREdaWDBJWXpsUmVRRGdNMEVMQU9CUHJBOWhSUDBZNEZNTm1KNnlidWNoaE9QYzYwRHlUTE1BZ1BHWmFNRlkzcW84QVBDWm9BVUE4RFhXaHpDRzAyNjJmRkQ1YVNucmRqK0VZQW9KcWJ0c3FzSU5XZ0FZV1ZNVmdoYU1wcXhiVXkwQWdIOFJ0QUFBdnNab2RIYnRySnN0UFRTZHB2NlcvMEh1UlNCcGo2WlpBRUJVVEdCa0xFY3FEd0FFUVFzQTRHdXNEMkhIYnJyWjBrU0VDU3JyOXRBQk5CbFlOVlZoMmc0QXhNT1VLY1ppb2dVQThDK0NGZ0RBdDFnZndpNzBOOFZQVlhxeStvRE1YdTVGSUduM1RWV2NhekVBUk1Va1BNWmlvZ1VBOEMrQ0ZnREF0d2hhc0F1bjNXenBwdmdFRGZ1SjMrZGVCNUkzMTJJQWlJNkpGb3psWUpqcUJ3QmtUdEFDQVBpcTRmRDdXb1hZb3JOdXRuUWJiYnJjOGlkMU4wMVYrQjRGQUpGcHFxSWJKdVBCR0V5MUFBQUVMUUNBNzdwU0lyYmtwcHN0VjRvN1RXWGQ5cmY4MytaZUI1Sm5tZ1VBeE10VUM4WWlhQUVBQ0ZvQUFOOGxhTUUyOUxmUFRsVjJtc3E2M1RmTmdneWNOVlZ4cDlFQUVDMVRweGpMaWNvREFJSVdBTUEzV1IvQ2xwd09ueTJtYVJGQzJOTTdFdllvVEFRQTBUUFJnckVjcXp3QUlHZ0JBRHlIcVJaczBvZHV0blQ3YktMS3VqME1JU3h6cndQSld6UlZJUXdHQUhIek53V2pLZXZXK2hBQXlKeWdCUUR3SElJV2JNcHROMXN1VkhQUzNQSW5kYmROVlZ6b01nREViUWhGM21zVEl4RzBBSURNQ1ZvQUFOOWxmUWdiMG8vaVAxWE02U3JydHQ5Ri9DNzNPcEE4WVRBQW1BNVRMUmpMaWNvRFFONEVMUUNBNXpMVmd0ZWFkN1BsblNwT21sditwTzY2cVFvSE5nQXdIWjFlTVJJVExRQWdjNElXQU1CelhRMFRDZUJIZk9obVMyR2RDU3ZydHIvbGY1QjdIVWphbzJrV0FEQTVBcEtNNWEzS0EwRGVCQzBBZ0djWjFvYzRLT2RIM0hhenBjUExDU3ZyZGorRXNNcTlEaVR2dktrS1UzY0FZRUthcWpEUmd0RU1xeFVCZ0V3SldnQUFMeUZvd1V2MU44UlBWVzN5K3BERlh1NUZJR24zZmRCQ2l3RmdrbTYwalpFSVdnQkF4Z1F0QUlCbkcxWS9XQi9DUzh5NzJkSU44UWtyNi9Zd2hQQno3blVnZWF1bUtoNjBHUUFteVZRTHhuS2s4Z0NRTDBFTEFPQ2xUTFhndVQ0TTRSeW03VUwvU054TlV4VSs1d0F3WFovMGpwR1lhQUVBR1JPMEFBQmV5c0U1ejNIYnpaWUxsWnEyc203N3RTL0h1ZGVCNUsyMEdBQW16VVFMeHJJM1RBQUVBRElrYUFFQXZJajFJVHhELy9rNFZhZ2tuT2RlQUpKMzJWU0ZXN0FBTUdGTlZmU3JDdS8xa0pGWUh3SUFtUkswQUFCK2hLa1dmTXU4bXkzdlZHamF5cnJ0Yi9rZjVGNEhrdGFId2t6ZUFZQTBtR3JCV0t3UEFZQk1DVm9BQUQ5QzBJS3YrVEJNUFdIQ3lycmRkd0JOQnM2YnFualFhQUJJZ3FBRll6SFJBZ0F5SldnQkFMeVk5U0Y4eFcwM1d6cWNUME8vTW1Rdjl5S1F0UHVtS2xaYURBREpzQXFNc1J5clBBRGtTZEFDQVBoUkZ5ckhFMzN3NWxSQnBxK3MyMzcwN2Z2YzYwRHloTUlBSUNGTlZRaGFNSnF5YmsyMUFJQU1DVm9BQUQ5SzBJS241dDFzZWFjaVNYRExuOVRkTkZWaHhSRUFwT2RXVHhuSmljSURRSDRFTFFDQUg5TE5sdjBPM0h2Vkk0VHdZVmdudzhTVmRUczMrcFlNbUdZQkFHa3kxWUt4bUdnQkFCa1N0QUFBWHNQaE9yZmRiT25RTWdGbDNlNmJaa0VHUGpSVjBXazBBQ1RKejNqR0ltZ0JBQmtTdEFBQVhzUDZrTHc5aGhCT2N5OUNRdnJBekVIdVJTQnBqOEpFQUpBMFFRdkc4bllJcmdNQUdSRzBBQUIrbVBVaDJadDNzK1ZkN2tWSVFWbTNoOVlwa0lGVlV4VVBHZzBBYVJxbVZqMXFMeU14MVFJQU1pTm9BUUM4bHZVaGVmclF6Wlo2bjQ3ekVNSmU3a1VnYWZkTlZaeHJNUUFrejFRTHhuS2k4Z0NRRjBFTEFPQzFyQS9KejIwM1c1cCtrSWl5YnZzSGd1OXlyd1BKbTJzeEFHVGhrell6RWtFTEFNaU1vQVVBOENyV2gyU25IOFY3bW5zUkV1T1dQNm03YnFyQ29Rc0E1TUhQZk1aaWRRZ0FaRWJRQWdEWUJGTXQ4akh2WnN1NzNJdVFpckp1KzF2K2IzT3ZBOGt6Z1FjQThtRjFDR1BaSyt2MlVQVUJJQitDRmdEQUpnaGE1T0ZETjF0ZTVWNkVWSlIxdTIrYUJSazRhNnBDT0F3QU10RlV4WU9KaTR6SStoQUF5SWlnQlFEd2FzT0VnMXVWVE5wdE4xdTZGWjZXdnA5N3VSZUJwRDBLRXdGQWxxd1BZU3pXaHdCQVJnUXRBSUJOTWRVaVhmMWg1V251UlVqSk1OSjJtWHNkU041aXVOVUtBT1RGK2hER0ltZ0JBQmtSdEFBQU5zVktpWFROaDZrbHBFTXdpdFRkTmxYaGN3NEFlVExSZ3JFY3F6d0E1RVBRQWdEWUNPdERrdldobXkyRmFCSlMxdTJKQjRCa3dLb2pBTWhVVXhVbXUxU1U3UUFBSUFCSlJFRlVXakNhNGU4dEFDQURnaFlBd0NhNVBaeVcyMjYyZEZpWkh2OU9TZDFsVXhWdXNnSkEzbTV5THdDanNUNEVBREloYUFFQWJKTEpCK2w0RENHYzVsNkUxSlIxMndkbkRuS3ZBMG5ydjNldHRCZ0FzaWQweVZoTXRBQ0FUQWhhQUFBYlkzMUlVdVpEUDBsRVdiZjdEcURKd0hsVEZiNTNBUURXaHpBV0V5MEFJQk9DRmdEQXBwMnI2T1I5NkdaTDAwblMwLy9iM011OUNDVHQzczhnQUdBZ2FNRllEb2FRT3dDUU9FRUxBR0RUSE5CUDIyMDNXeTV5TDBKcXlycnRiMVc5ejcwT0pHL1ZWTVdETmdNQXc0U3IrK3dMd1Zpc0R3R0FEQWhhQUFBYjFjMlcvU0hYdGFwTzBtTUk0VFQzSWlUS0xYOVNkOU5VeFlVdUF3QlBtR3JCV0t3UEFZQU1DRm9BQU50Z3FzVTB6YnZaOGk3M0lxU21yTnMrUEhPY2V4MUlua2s4QU1BZmZWSVJSbUtpQlFCa1FOQUNBTmdHUVl2cCtkRE5sdnFXSnRNc1NOMWxVeFZ1ckFJQWYrVDNBOFppb2dVQVpFRFFBZ0RZT090REp1ZTJteTNkQms5UVdiZXJFTUpCN25VZ2FZK21XUUFBWDlKVWhZa1dqR1d2ckZ0aEN3QkluS0FGQUxBdHBpTk1RMzlJZVpwN0VWSlUxdTJoQTJneWNONVV4WU5HQXdCZmNhc3dqRVRRQWdBU0oyZ0JBR3lMb01VMHpMdlo4aTczSWlTcW4yYXhsM3NSU05wOVV4VXJMUVlBdnNGVUM4WWlhQUVBaVJPMEFBQzJ3dnFRU2ZqUXpaWUNNUWtxNi9Za2hQQSs5enFRUEJOYkFJRHY2VlNJa1p3b1BBQ2tUZEFDQU5nbWgvanh1dTFtUzRlVTZYTExuOVRkTkZYaFp3d0E4RDJDRm96bHJjb0RRTm9FTFFDQXJlbG15NHNRd3FNS1I2ZnZ5V251UlVoVldiZnpFTUp4N25VZ2VYTXRCZ0MrcEt6Yi9mNTM0ckp1KzFEbVB4V0pzUXlUQmdHQVJQMlh4Z0lBVzNabGhVRjA1dDFzZVpkN0VWTFVQMVEyellJTWZHaXF3dmN3QU9BM3crL0JwOFBYTzVVaEVuM1E0cE5tQUVDYUJDMEFnRzBUdElqTGgyNjJORzQvWGYwNm1JUGNpMERTSG9XSkFJQWdYTUUwSE9rVEFLVHJ6WHE5MWw0QVlLdU9QcDQ5aEJEMlZIbDB0OTFzNlVGUG9zcTZQUXdoL0pwN0hVamVMMDFWbkdzekFPUkp1SUtKdVcrcTRsRFRBQ0JOLzBkZkFZQWRNRUZoZlAwdDhIbnVSVWljdzJkU2R5dGtBUUI1S3V2MnRLemIvdS9LL3cwaC9GM0lnb2s0R0FMeEFFQ0NyQTRCQUhiQitwRHhMYnJac3N1OUNLa3E2L2JFdzJZeXNOQmtBTWhISDY1NE1yM0NoRVNtcXA4cWVhZDdBSkFlUVFzQVlPdTYyZkxxNk9QWm80ZGpvN25zWnN1TFROOTdMdHp5SjNYWFRWVjgwbVVBU0p0d0JRazZNdVVUQU5Ja2FBRUE3SXFwRnVPNGRRczhiV1hkOWl0aDN1WmVCNUxuK3hnQUpFcTRnc1NkYURBQXBFblFBZ0RZRlVHTDNldW5pTXk3MmZJaHR6ZWVpN0p1OTAyeklBTm5UVlVZdHd3QUNSR3VJQ1BIbWcwQWFYcXpYcSsxRmdEWWlhT1BaLzFCMllGcTc4eFBWb2FrcmF6YlBtVHhjKzUxSUdsOVlPeXdxUXFCTVFDWU9PRUtNdmJYcGlvNkh3QUFTSXVKRmdEQUxsMDVGTjZaU3lHTHRKVjFlK2pmRXhsWUNGa0F3SFFKVjhDL0hJVVFCQzBBSURHQ0ZnREFMbDA0R042SjIvNXdNb1AzbVR0QkdsSjMwMVNGenprQVRJeHdCZnpKaWIvZkFDQTlnaFlBd001MHMyVjM5UEhzM3ZxUXJlckg3TSs3MmRJTjhJU1ZkWHRpMXk4WldHa3lBRXhEV2JmOWpmMzVFSzd3OXg3ODNwRjZBRUI2QkMwQWdGMnpQbVM3Rm4yZ0plVTN5TCs0RFVYcUxwdXErS1RMQUJBdjRRcDR0cmRsM2U1YmlRY0FhUkcwQUFCMnpmcVE3Ym5zWmtzSDhJa3I2M2JsUVRhSmV6VE5BZ0RpSkZ3QlA2ei90eU5JREFBSkViUUFBSGJLK3BDdHVlMm5XU1Q2M2hqMHQ2RDBtUXljTjFWeHA5RUFFQWZoQ3RpSUUwRUxBRWlMb0FVQU1BYnJRemFydi8wOTcyWkxZMGpUZHg1QzJNdTlDQ1R0ZnZpY0F3QWpFcTZBalR0U1VnQklpNkFGQURBRzYwTTJhOUZQQ2tucERmRm53OFB1OTBwRDRoWjJWd1BBT0lRcllLdE9sQmNBMHZKbXZWNXJLUUN3YzBjZnorNDh2TnVJeTI2Mm5DZndQdmlPc203N01iUEg2a1RDYnBxcThBQWFBSFpJdUFKMjZpOVc1QUZBT2t5MEFBREcwbytHL3gvVmY1WGIvdmIzaEY4L3oxVFc3YW1RQlJudy9Rd0FkcUNzMjhNaFhERVhyb0NkNm9OTmdoWUFrQWhCQ3dCZ0xGZUNGcS95MkQ4WTdXWkxJL1lUVjlidC9oQk1ncFJkTmxWaEJSSUFiTWtRcmpnZHdoVnYxUmxHY1RJOEN3RUFFaUJvQVFDTW9wc3Q3NDQrbnQxNnlQZkRGdDFzNlZBeUR3czNEVW5jbzJrV0FMQjV3aFVRblNNdEFZQjBDRm9BQUdPNk1OWGloMXgycytYRkJGODNMelE4SEhjQVRlck9tNm93blFjQU5rQzRBcUptSFNRQUpFVFFBZ0FZay9VaEwzZnI0RDBycXhEQ1h1NUZJR24zVFZXc3RCZ0FmcHh3QlV4SFdiZEhWdVlCUUJvRUxRQ0EwVmdmOG1MOWVQMTVOMXU2K1oyQnNtNzcvYjN2YzY4RHladHJNUUM4bkhBRlRGYi9kNTZnQlFBa1FOQUNBQmliOVNIUHQraG1TdzlrOHVHV1A2bTdhYXJpa3k0RHdQTUlWMEFTanJRUkFOSWdhQUVBak0zNmtPZTU3R2JMaXltOFVGNnZyTnU1L2Ixa3dEUUxBUGdPNFFwSXpvbVdBa0FhM3F6WGE2MEVBRVoxOVBHczg5RHdtL3IxS2lkV2h1U2hyTnY5RU1KZENHRXY5MXFRdEE5TlZTeTBHQUQrYlBoOThIVDRlcWRFa0p6L2JxckMzL2NBTUhFbVdnQUFNVGdQSWZ4ZEo3N29zYis5Sm1TUmxZV1FCWWw3dEJvSEFINVB1QUt5MHE4UHNVSVBBQ1pPMEFJQWlNR1ZvTVZYTGJyWnNvdjB0YkZodzJqb3BicVN1SlViZkFBZ1hBRVpPeEcwQUlEcEU3UUFBRWJYVDJzNCtuaDI3ZUhpbjF4MnMrVkZaSytKN1RwWFh4SjMyMVNGenprQTJSS3VBSWFnQlFBd2NZSVdBRUFzcmp4by9KM2JZWVVFbVNqcjlzUy9BVExnK3hvQTJSR3VBUDdnU0VFQVlQb0VMUUNBV0ZnZjhoK1BJWVI1UCtramxoZkVUcmpsVCtxdW02b3dJaG1BTEFoWEFOK3cxNitOYktyaVRwRUFZTG9FTFFDQUtGZ2Y4anVMYnJic0lubzliRmxadC8wdC83ZnFUT0pNc3dBZ2FjSVZ3QXYwRXcydENnV0FDUk8wQUFCaVluMUlDSmZkYk9saFMwYUdCL0tyM090QThzN2MyQU1nVldYZDlzR0t1YjlsZ0Jld1BnUUFKazdRQWdDSVNlN3JRMjdkK001U0g3TFl5NzBJSk8zZWFod0FVak9FS3o1LytWME9lQ2xCQ3dDWXVEZnI5Vm9QQVlCb0hIMDh5M1dxeFdNL090VEtrTHowZTNsRENML21YZ2VTOTFOVEZTYjFBREI1d2hYQUpqVlY4VVpCQVdDNlRMUUFBR0p6a1duUVlpRmtrU1dIejZUdVJzZ0NnQ2tUcmdDMnBhemJrNllxUGlrd0FFeVRvQVVBRUpWdXRydzYrbmoybU5sRHpNdHV0blFRbVpuaG9mMXg3blVnZVNzdEJtQnFoQ3VBSGVuWGh3aGFBTUJFQ1ZvQUFESHExNGU4ejZRenQvMDBpd2hlQjd0M3J1WWs3dElOUFFDbVFyZ0NHTUdKdndzQllMb0VMUUNBR09VU3RPZ25kOHk3MmZJaGd0ZkNEcFYxMjkveVAxQnpFdlpvbWdVQXNST3VBRVoycEFFQU1GMXYxdXUxOWdFQTBUbjZlUGFRd2NQT242d015VTladC9zaGhEc1A4MG5jV1ZNVmdoWUFSS2VzMi80RytWeTRBb2pFZnpkVjRmSUZBRXlRaVJZQVFLeFNuMnB4S1dTUnJYTVA5VW5jdlpBRkFERXA2L2JvU2JqQ1ZERWdKaWZEOHc4QVlHSUVMUUNBV0tVY3RMZ05JU3dpZUIzczJQQ1FQNGUxT09UTjl6Y0FSaWRjQVV6RWthQUZBRXlUb0FVQUVLVnV0cnc2K25qMm1PRE4vLzQ5emJ2WjBtalFQSjNuWGdDU2Q5TlVoUWZGQUl4Q3VBS1lvQk5OQTRCcEVyUUFBR0tXNGxTTFJUZGJkaEc4RG5hc3JOditvZit4dXBNNDB5d0EyQ25oQ21EaWpqUVFBS2JwelhxOTFqb0FJRXBISDgvNm14My9TS2c3bDkxc09ZL2dkYkJqWmQzdTk0TmFQUHduY1pkTlZmZ2VCOERXQ1ZjQWlmbHJVeFV1WkFEQXhKaG9BUUJFcTVzdFB4MTlQTHRQNU9IcHJadmVXVnM0QkNCeGo3N0hBYkJOd2hWQXdvNkdZRDRBTUNHQ0ZnQkE3UHIxSVQ5UHZFdjlBZVM4bXkwZkluZ3Q3RmhadDRjT29NbkFxcWtLMytNQTJDamhDaUFUL1RUUEM4MEdnR2tSdEFBQVluZVJRTkJpMGMyV2JxZmtheFZDMk11OUNDVHR2cW1LY3kwR1lCT0drT3FwaVdCQVJvNDBHd0NtUjlBQ0FJaGFIMUNZK1BxUTYyNjJkRE1sVTJYZDlqZVQzdWRlQjVJMzEySUFYdU5KdUtML21mSldNWUhNK0w0SEFCTWthQUVBVE1GVTE0ZmNPNERNbmx2K3BPNm1xWXBQdWd6QVN3bFhBUHhISDlMM2V6VUFUSXVnQlFBd0JWTmRIM0xhelpZUEVid09SbERXclVNRGNpQk1Cc0N6Q1ZjQWZGVS9EVkhRQWdBbVJOQUNBSWplUk5lSC9OSy83Z2hlQnlNbzYzYmZOQXN5OEtHcGlqdU5CdUJiaENzQW51VkltUUJnV2dRdEFJQ3BtTkw2a090dXRuVElucmRGQ0dFdjl5S1F0TWNRd2txTEFmZ1M0UXFBRnhPMEFJQ0pFYlFBQUtaaUt1dEQ3bzNTejl0d3NMRE12UTRrYjlGVWhkVklBUHhHdUFMZ1ZRNzY3Nk1teGdIQWRBaGFBQUNUTUt3UHVaM0FROXZUYnJaMCtKZzMwMHhJM1cxVEZSZTZETUN3TG0wdVhBR3dFZjFVQzBFTEFKZ0lRUXNBWUVyNmc3My9pZmoxL3RJSFFpSjRIWXlrck51VEVNSTc5U2R4Q3cwR3lOY1FyamdkdnZ6ZUE3QTVSOFBhVkFCZ0FnUXRBSUFwdVlvNGFISGR6WlltR2VDV1A2bTdicXJpa3k0RDVFVzRBbUFuVHBRWkFLWkQwQUlBbUl4dXRyeUxkSDNJL1RBdW1ZeVZkZHZmOGovd0dTQmhqNlpaQU9SRHVBSmc1NDZWSEFDbVE5QUNBSmlhR05lSG5IYXo1VU1FcjRPUkRBY1JLL1VuY2VkTlZkZ1pEWkF3NFFxQWNaVjFlOVJVaFpXa0FEQUJnaFlBd05URXRqN2tsMjYyOUJDRVBtU3hsMzBWU0ZrL3VjZDZKSUFFQ1ZjQVJLVmZIK0laQXdCTWdLQUZBREFwa2EwUHVlNW1Td2VQbWV0dkhJVVFmczY5RGlSdjFWU0Z5VDBBaVJDdUFJaldrZFlBd0RRSVdnQUFVeFREK3BEK2R2ZDhtdVZqdzRSdFNOMU5VeFVYdWd3d2JjSVZBSk1nYUFFQUUvRi9OQW9BbUtDckNGN3lhVGRidXQyZHViSnUrNE9LNDl6clFQSldXZ3d3WGYzdksyWGQ5b0c1L3cwaC9GM0lBaUJxYjRkZ0hBQVFPUk10QUlESkdkYUgzSXg0d1AxTE4xdmFtVW93ellJTVhEWlY4VW1qQWFabENJTisvdHJUUG9CSjZhZGErQjBjQUNJbmFBRUFUTlhGU0VHTDYyNjJkTGhPZjREUjMvSS9VQWtTOWhoQ1dHZ3d3RFFJVndBazQwVFFBZ0RpSjJnQkFFelYxVEQ2ZUpmdVF3aHpueGlHVWE0T29FbmRlVk1WVmlRQlJFeTRBaUJKUjlvS0FQRVR0QUFBSnFtYkxSK09QcDVkNzNqSDlHbi8zL1dKWVZnWjRqQ0RsTjAzVmJIU1lZRDRDRmNBSk85RWl3RWdmb0lXQU1DVVhlMHdhUEZMTjF0MlBpMlVkZHMvOUhxZmZTRkluWWt0QUJFUnJnREl5bDVadDRkTlZkeHBPd0RFUzlBQ0FKaXlYYTBQdWU1bXkzT2ZGQVp1K1pPNm02WXFyblFaWUZ6Q0ZRQlo2OWVIQ0ZvQVFNUUVMUUNBeWRyUitwRDdFTUxjcDRUdzd3T1Avck53ckJna3pqUUxnSkVNazdQbXdoVUEyVHNaTHBjQUFKRVN0QUFBcG03YjYwTk8rMENIVHdsbDNlNmJaa0VHUGpSVllVMFN3QTZWZFh2MEpGeHhvUFlBREJNdEFJQ0lDVm9BQUZPM3pmVWh2M1N6cFFOSFBsczQvQ0J4ajhKRUFMc2hYQUhBZDVpa0NBQ1JlN05lci9VSUFKaTBvNDluMjVocWNkM05scWMrR1lSL0g0WWM5dHRxalBBbWNiODBWWEd1eVFEYklWd0J3QXY5cmFtS1Q0b0dBSEV5MFFJQVNNSEZob01XOThORGNQanNYTWlDeE4wTFdRQnNubkFGQUsvUS93d1J0QUNBU0FsYUFBQ1QxODJXVjBjZnp4NDNlQkIrMnMyV0R6NFpoSDhma0p4c1lXSUt4RWE0REdCRGhDc0EySkFqaFFTQWVBbGFBQUNwNk5lSHZOL0FlL21sbXkwN253cWVjTXVmMU4wWVNRendPc0lWQUd6QmlhSUNRTHdFTFFDQVZHd2lhSEhkelpZTzFmbE5XYmY5Z2NsYkZTRnhwbGtBL0FEaENnQzI3S0NzMi8ybUtremNCSUFJQ1ZvQUFFbll3UHFRZTRlTlBOVS8wRExOZ2d5Y05WVnhwOUVBejFQVzdlRVFyRmdJVndDd0EzMm96L1E1QUlpUW9BVUFrSkxYVExVNDdXWkx0MFI0YXZXSzRBNU13YU13RWNEM1BRbFhtSFFGd0s2ZENGb0FRSndFTFFDQWxQeG8wT0tYYnJic2ZCTDRiRGhRK1ZsQlNOekNHR0tBTHhPdUFDQVNKeG9CQUhFU3RBQUFrdkdENjBPdXU5blNqVzcrNkVKRlNOeHRVeFUrNXdCUENGY0FFS0VqVFFHQU9BbGFBQUNwZWNsVWkvdmhRVHI4cHF6Yi9zYlFzWXFRdUlVR0F3aFhBQkM5dmY1blZWTVZkMW9GQUhFUnRBQUFVblArZ3FERmFUZGJHcHZQSDdubFQrb3VtNnF3NXhuSWxuQUZBQk56NHU5VUFJaVBvQVVBa0pSdXR1eU9QcDcxa3lvT3Z2TytmdW4vdjdyUFUyWGRMcDd4MllFcDY5Y3JyWFFReUkxd0JRQVRabjBJQUVSSTBBSUFTRkcvUHVUbmI3eXY2MjYyUE5kNW5pcnJkdDhCTkJrNE4zWVl5TVh3cy8xMFdKY2tYQUhBVkozb0hBREVSOUFDQUVqUnhUZUNGdmZEVFViNG96NThzNmNxSk94KytKd0RKT3RKdUtML2VxZlRBQ1JBV0JBQUl2Um12VjdyQ3dDUW5LT1BaM2RmV1FIeFZ5dEQrS095YnZ0UnJQOVVHQkwzVTFNVmRqc0R5Ukd1QUNBRGYydXE0cE5HQTBBOFRMUUFBRkwxcGZVaHZ3aFo4QlZ1K1pPNkd5RUxJQ1hDRlFCa3ByOGNJR2dCQUJFUnRBQUFVdlhIOVNIWDNXenBNSjAvS2V1MlA2QTVWaGtTdDlCZ1lPcUVLd0RJMklrTEFnQVFGMEVMQUNCSi9lU0tvNDluOThQNmtQNS81enJOVjNoWVJlb3VtNm93elFlWUpPRUtBUGlYSTJVQWdMZ0lXZ0FBS2Z1OFB1UzBteTBmZEpvL0t1dDJOWVJ4SUZXUHBsa0FVeU5jQVFCL2N0RC9mR3lxd3JNTkFJaUVvQVVBa0xKK1VrSFhUN2ZRWmY2b3JOdERCOUJrNE56RFdHQUtoQ3NBNEx0T2hnc2xBRUFFQkMwQWdHUjFzK1ZkQ09GQ2gvbUtmcHJGbnVLUXNQdW1LbFlhRE1Tc3JOdlRKd0VMUDVjQjRPdU9CQzBBSUI2Q0ZnQUFaS2VzMi80bTBIdWRKM0VtdGdCUkVxNEFnQjl5b213QUVBOUJDd0FBY3VTV1A2bTdhYXJDYlRjZ0dzSVZBUEJxUjBvSUFQRjRzMTZ2dFFNQWdHeVVkVHNQSWZ4ZHgwbmNYNXFxdU5Oa1lFekNGUUN3Y1g5dHFxSlRWZ0FZbjRrV0FBQmtvNnpiZmRNc3lNQUhJUXRnTE1JVkFMQlYvVlFMUVFzQWlJQ2dCUUFBT1ZtRUVBNTBuSVE5Q2hNQnV5WmNBUUE3Y3hKQ3VGQnVBQmlmb0FVQUFGa282L1l3aExEVWJSSzNhcXJpUVpPQmJST3VBSUJSSENrN0FNUkIwQUlBZ0Z5YzZ6U0p1Mitxd3VjYzJKcXlibytHNlZEQ0ZRQXdqcmZxRGdCeEVMUUFBQ0I1WmQzMjQxWGY2VFNKbTJzd3NHbER1R0kraEN1czN3S0FjZDMzUDV1YnF1ajBBUURHSldnQkFFQU8zUEluZGRkTlZYelNaV0FUaENzQUlDcjNJWVNyRU1LRmdBVUF4RVBRQWdDQXBKVjF1ekJlbFF3c05CbDREZUVLQUlpS2NBVUFSRTdRQWdDQVpKVjF1eDlDV09rd2lUdHJxdUpPazRHWEVxNEFnS2dJVndEQWhBaGFBQUNRc2o1a3NhZkRKT3pSYWh6Z0pZUXJBQ0FxajAvQ0ZWWUJBc0NFdkZtdjEvb0ZBRUJ5eXJvOURDSDhxck1rN3FlbUtpNDBHZmdXNFFvQWlNcm5jTVZWVXhWWFdnTUEwMlNpQlFBQXFYTDRUT3B1aFN5QXJ4R3VBSUNvQ0ZjQVFHSUVMUUFBU0U1WnR5Y2hoR09kSlhFTERRYWVHcVk1blE0Qmk3ZUtBd0NqRXE0QWdJUUpXZ0FBa0NLMy9FbmRwUjNPUUJDdUFJRFlDRmNBUUNZRUxRQUFTRXBadHlzajBrbGMvL0IycGNtUUwrRUtBSWlLY0FVQVpFalFBZ0NBWkpSMXUyK2RBaGs0YjZyaVRxTWhMOElWQUJBVjRRb0F5SnlnQlFBQUtUa1BJZXpwS0FtN0h6N25RQWFFS3dBZ090Zjlxa3JoQ2dEZ3pYcTl6cjRJQUFCTVgxbTNSeUdFZjJvbGlmdS9IdXBDMm9RckFDQTYxMCttVnp4b0R3QVFUTFFBQUNBaGJ2bVR1aHNoQzBpVGNBVUFSRWU0QWdENEprRUxBQUFtcjZ6Yi9uRHFXQ2RKM0VLRElSMWwzZTRQNFlxRmNBVUFSRUc0QWdCNE5rRUxBQUFtYlRpb01zMkMxRjAyVmRIcE1remJrM0JGLy9WT093RmdkTUlWQU1BUEViUUFBR0RxK3B2QUI3cEl3aDVOczREcEVxNEFnT2dJVndBQXJ5Wm9BUURBWkEwNzdSMUFrN3B6RDRCaFdvUXJBQ0E2d2hVQXdFWUpXZ0FBTUdXckVNS2VEcEt3KzZZcVZob004Uk91QUlEb0NGY0FBRnNqYUFFQXdDU1ZkWHNTUW5pdmV5UnVyc0VRTCtFS0FJak9iUWpoWWdoWDNHa1BBTEF0Z2hZQUFFeVZXLzZrN3FhcGlrKzZESEVScmdDQTZBaFhBQUE3SjJnQkFNRGtsSFhiMy9JLzFqa1NaNW9GUkdUNDJTTmNBUUJ4RUs0QUFFWWxhQUVBd0tRTU40blBkWTNFZmZEQUdNWlgxdTNwaytrVmUxb0NBS01TcmdBQW9pRm9BUURBMUN3Y2RwRzRSNnR4WUR6Q0ZRQVFGZUVLQUNCS2I5YnJ0YzRBQURBSlpkMGVoaEIrMVMwUzkwdFRGYWEyd0E0SlZ3QkFWSVFyQUlEb21XZ0JBTUNVT0h3bWRiZENGckFid2hVQUVCWGhDZ0JnVWdRdEFBQ1loTEp1VDBJSTczU0x4QzAwR0xaSHVBSUFvbkxmQnl2NlFMMXdCUUF3TllJV0FBQk14WVZPa2JqcnBpbythVEpzbG5BRkFFVGxjN2ppb3FtS1Rtc0FnS2tTdEFBQUlIcGwzZmEzL0E5MGlzU1paZ0ViSWx3QkFGRVJyZ0FBa2lOb0FRQkExTXE2M1E4aHJIU0p4SjBabHd5dlU5YnRVUWhoUG53SlZ3REF1SVFyQUlDa0NWb0FBQkM3bFFNekV2Zlk3NlhXWkhpNUorR0tVNU9QQUdCMHdoVUFRRGJlck5kcjNRWUFJRXBsM1I2R0VIN1ZIUkwzVTFNVkY1b016eU5jQVFCUkVhNEFBTEprb2dVQUFERnorRXpxYm9RczRQdUVLd0FnS3NJVkFFRDJCQzBBQUloU1diZjlZZHF4N3BDNGxRYkRsd2xYQUVCVUhvY2d2SEFGQUpDOUlHZ0JBRURFempXSHhGMDJWZkZKaytFL2hDc0FJQ3FQdytTS3E2WXFyclFHQU9BL0JDMEFBSWhPV2JjckIyd2s3dEUwQy9nMzRRb0FpSXB3QlFEQU13aGFBQUFRbGJKdTkwTUlDMTBoY2VkTlZkeHBNcmtxNi9adytGNHZYQUVBNHhPdUFBQjRJVUVMQUFCaTA2OE0yZE1WRW5iZlZJVnBGbVJuQ0ZlY0R0TXIzdm9FQU1Db2hDc0FBRjVCMEFJQWdHZ000K1BmNndpSk03R0ZiQWhYQUVCVWhDc0FBRFpFMEFJQWdKaWM2d2FKdS9GUW05UUpWd0JBVklRckFBQzJRTkFDQUlBb2xIWGJIOGdkNndhSk04MkNKQWxYQUVCVWhDc0FBTFpNMEFJQWdOR1ZkYnNmUWxqcEJJbTdiS3FpMDJSU0lWd0JBTkc1ZmhLd2VOQWVBSUR0RWJRQUFDQUcvUzMvQTUwZ1lZK21XWkFDNFFvQWlJNXdCUURBQ0FRdEFBQVkxWEJvNXdDYTFLMDgrR2FxaENzQUlEckNGUUFBSXhPMEFBQmdiUDNLa0QxZElHSDNUVldjYXpCVE1xeDAraHl1T05ZOEFCaWRjQVVBUUVRRUxRQUFHRTFadHljaGhQYzZRT0xtR3N3VVBBbFg5Ri92TkEwQVJpZGNBUUFRS1VFTEFBREc1SlkvcWJ0cHF1S1RMaE1yNFFvQWlJNXdCUURBQkFoYUFBQXdpckp1N2ZvbkI2WlpFQjNoQ2dDSWpuQUZBTURFQ0ZvQUFMQnp3eUdmYVJhazdrTlRGWGU2VEF5RUt3QWdPcmZEMzBUQ0ZRQUFFeVJvQVFEQUdCWWhoRDJWSjJHUElZU1ZCak1tNFFvQWlFNGZycmdZd2hVQ3VRQUFFL1ptdlY3ckh3QUFPMVBXN1dFSTRWY1ZKM0UvTlZWeG9jbnNtbkFGQUVSSHVBSUFJRUVtV2dBQXNHc09uMG5kclpBRnUxYlc3Vnk0QWdDaUlWd0JBSkE0UVFzQUFIYW1yTnVURU1LeGlwTzRoUWF6QzJYZG5qNlpYbUVkRXdDTVM3Z0NBQ0FqZ2hZQUFPeVNXLzZrN3JxcGlrKzZ6TFlJVndCQVZJUXJBQUF5SldnQkFNQk9sSFhiMy9JL1VHMFNaNW9GR3lkY0FRQlJFYTRBQUVEUUFnQ0E3U3ZyZGorRXNGSnFFbmZtWVR1Yklsd0JBRkc1RHlHY0MxY0FBUENab0FVQUFMdXdjbEJJNGo0L2ZJY2ZKbHdCQUZIcGY3Kzc2cWRYTkZYUmFRMEFBRSs5V2EvWENnSUF3TmFVZFhzVVF2aW5DcE80bjVxcXVOQmtYa3E0QWdDaUlsd0JBTUN6bUdnQkFNQzJ1ZVZQNm02RUxIZ0o0UW9BaUlwd0JRQUFMeVpvQVFEQTFneUhpY2NxVE9KV0dzejNETk45NXNPWGNBVUFqRXU0QWdDQVZ4RzBBQUJnbTB5eklIV1hUVlY4MG1XKzVFbTRvZytkSFNnU0FJeEt1QUlBZ0kwUnRBQUFZQ3ZLdWwwNVdDUnhqNlpaOEVmQ0ZRQVFGZUVLQUFDMlF0QUNBSUNOSyt0MlA0U3dVRmtTZDk1VXhaMG1JMXdCQUZFUnJnQUFZT3NFTFFBQTJJWitaY2lleXBLdys2WXFUTFBJbUhBRkFFVGw4VW00d2xvM0FBQzJUdEFDQUlDTkt1djJKSVR3WGxWSm5Ja3RHUkt1QUlDb2ZBNVhYRFZWY2FVMUFBRHNrcUFGQUFDYjVwWS9xYnZ4TUQ4ZndoVUFFQlhoQ2dBQW9pQm9BUURBeHBSMTJ4OUdIcXNvaVRQTkluRmwzUjRPZlJhdUFJRHhDVmNBQUJBZFFRc0FBRGFpck50OTB5ekl3SWVtS2pxTlRzOFFyamdkcGxlOHpiMGVBREF5NFFvQUFLSW1hQUVBd0tZczNQd21jWS9DUkdrUnJnQ0FxQWhYQUFBd0dXL1c2N1Z1QVFEd0tzTmg1YStxU09KK2FhcmlYSk9uVGJnQ0FLSWlYQUVBd0NTWmFBRUF3Q1k0ZkNaMTkwSVcweVZjQVFEUnVRNGhYQWhYQUFBd1ZZSVdBQUM4U2xtM0p5R0VkNnBJNHVZYVBDM0NGUUFRbmVzbjB5c2V0QWNBZ0NrVHRBQUE0TFhjOGlkMU4wMVZmTkxsK0FsWEFFQjBoQ3NBQUVpU29BVUFBRCtzckZ1SG1lVEFOSXVJbFhXNy95UmNjWng3UFFBZ0FzSVZBQUFrVDlBQ0FJQWZNaHh1bW1aQjZzNmFxcmpUNWJnOENWZWNXbDBFQUZFUXJnQUFJQ3VDRmdBQS9LaFZDR0ZQOVVqWW96QlJQSVFyQUNBNndoVUFBR1RyelhxOTFuMEFBRjZrck52REVNS3Zxa2JpZm1xcTRrS1R4eU5jQVFEUkVhNEFBQ0I3d1VRTEFBQitrTU5uVW5jclpERU80UW9BaUk1d0JRQUEvSUdnQlFBQUwxTFc3VWtJNFZqVlNOeENnM2RIdUFJQW9uTTdoS3Y3Y01XZDlnQUF3TzhKV2dBQThGSnUrWk82NjZZcVB1bnk5cFYxT3hldUFJQm9DRmNBQU1BekNWb0FBUEJzWmQzMnQvd1BWSXlFUFpwbXNWMWwzWjQrbVY2eGwvSjdCWUFKRUs0QUFJQWZJR2dCQU1DekRLUDlWNnBGNHM0ZE1teWVjQVVBUkVXNEFnQUFYa25RQWdDQTV6cDNRRXJpN29mUE9Sc2dYQUVBVVJHdUFBQ0FEUkswQUFEZ3U4cTZQUW9odkZjcEVyZHFxdUpCazMrY2NBVUFSRVc0QWdBQXRrVFFBZ0NBNTNETG45VGRORlZ4b2Nzdkoxd0JBRkVScmdBQWdCMFF0QUFBNEp1R1E5UmpWU0p4S3cxK1B1RUtBSWhLdi83c3FnOUhDMWNBQU1CdUNGb0FBUEE5cGxtUXVzdW1LajdwOHJjTks0VG13NWR3QlFDTTYzTzQ0cUtwaWs0dkFBQmd0d1F0QUFENHFySnUrMXYrQnlwRXdoNURDQXNOL3JJbjRZcFQzd3NBWUhUQ0ZRQUFFQWxCQ3dBQXZxaXMyME1IMEdTZ0g3SDlvTkgvSVZ3QkFGRVJyZ0FBZ0FnSldnQUE4RFVyNndGSTNIMVRGU3RORnE0QWdNZ0lWd0FBUU9RRUxRQUErSk95Yms5Q0NPOVZoc1JsUGJGRnVBSUFvaUpjQVFBQUV5Sm9BUURBbDdqbFQrcHVtcXE0eXEzTHdoVUFFQlhoQ2dBQW1DaEJDd0FBZnFlczIvNFE5bGhWU053OGx3YVhkWHM0VE84UXJnQ0E4VDMyd1FyaENnQUFtRFpCQ3dBQWZsUFc3WDRJNFZ4RlNOeUhwaXJ1VW42TFE3amlkQWlVdkkzZ0pRRkF6aDZIeVJWWE9VN1VBZ0NBRkFsYUFBRHdWSC9yZlU5RlNOaGpxcXR4aENzQUlDckNGUUFBa0xBMzYvVmFmd0VBK0h4SSs2dEtrTGhmbXFwSVptcUxjQVVBUkVXNEFnQUFNbUdpQlFBQW4xa1pRdXJ1VXdoWkNGY0FRRlNFS3dBQUlFT0NGZ0FBOUFlM0p5R0VkeXFXMUlWb0FBQWdBRWxFUVZSQjR1WlRmWHZDRlFBUUZlRUtBQURJbktBRkFBREJOQXN5Y04xVXhhY3B2VTNoQ2dDSWluQUZBQUR3RzBFTEFJRE1sWFc3Y0loTEJoWlRlSXRsM2U0UDRZcFRVMllBWUhUQ0ZRQUF3QmNKV2dBQVpHdzQxRjM1REpDNHM2WXE3bUo5aThJVkFCQ2Q2eUZjY2FFMUFBREFsd2hhQUFEa3JROVo3T1ZlQkpMMkdPTnFIT0VLQUlqTzlaUHBGUS9hQXdBQWZNdWI5WHF0UUFBQUdTcnI5akNFOEt2ZWs3aWZZcm1OS2x3QkFORVJyZ0FBQUg2SWlSWUFBUGt5Q3BuVTNZNGRzaEN1QUlEb0NGY0FBQUN2Sm1nQkFKQ2hzbTVQUWdqSGVrL2lGbU84UGVFS0FJaU9jQVVBQUxCUmdoWUFBSGt5ellMVVhUWlY4V21YNzdHczJ6NVlNUmV1QUlBb0NGY0FBQUJiSTJnQkFKQ1pzbTVYSVlRRGZTZGhqeUdFMVM3ZTNoQ3UrUHkxNTBNRkFLTVNyZ0FBQUhaQzBBSUFJQ1BEU29OUjFpbkFEcDAzVlhHM3JmK2NjQVVBUk9XMi85a3ZYQUVBQU95U29BVUFRRjdPSFF5VHVQdmhjNzVSd2hVQUVKWGJZUlhlMVRiRGxRQUFBRjhqYUFFQWtJbXlibzlDQ08vMW04U3RObldiVmJnQ0FLSWlYQUVBQUVSRDBBSUFJQjhiditVUGtibHBxdUxpTlM5SnVBSUFvaUpjQVFBQVJFblFBZ0FnQThQaDhiRmVrN2pGajd3OTRRb0FpSXB3QlFBQUVEMUJDd0NBeEpWMXUyK2FCUm00YktxaWUrN2JIRmJweklkd3hZRVBDQUNNU3JnQ0FBQ1lGRUVMQUlEMExSd2trN2pINTB5ekVLNEFnS2dJVndBQUFKTWxhQUVBa0xDeWJnOS9kSjBDVE1oNVV4VVBYM3E1d2hVQUVCWGhDZ0FBSUFtQ0ZnQUFhVnVGRVBiMG1JVGROMVd4ZXZyMmhDc0FJQ3IzZmJDaUQxaThaTTBYQUFCQXpBUXRBQUFTVmRidFNRamh2ZjZTdUQ1UUlWd0JBSEVScmdBQUFKTDJacjFlNnpBQVFJTEt1dTBmYXIvVld4TDIrUkJIdUFJQXhpZGNBUUFBWkVQUUFnQWdRV1hkOWpmNy82NjNBQUJza1hBRkFBQ1FKVUVMQUlERWxIVzdIMEs0Q3lIczZTMEFBQnNtWEFFQUFHVHZ2M0l2QUFCQWdoWkNGZ0FBYkpCd0JRQUF3Qk1tV2dBQUpLU3MyOE1Rd3E5NkNnREFLd2xYQUFBQWZJV0pGZ0FBYVRuWFR3QUFmdERqazNERkowVUVBQUQ0TWhNdEFBQVNVZGJ0U1FqaEgvb0pBTUFMZkE1WFhEVlZjYVZ3QUFBQTMyZWlCUUJBT2k3MEVnQ0FaeEN1QUFBQWVBVkJDd0NBQkpSMXV3Z2hIT2dsQUFCZklWd0JBQUN3SVZhSEFBQk1YRm0zK3lHRXV4RENubDRDQVBDRWNBVUFBTUFXbUdnQkFEQjlLeUVMQUFBR3doVUFBQUJiWnFJRkFNQ0VsWFY3R0VMNFZROEJBTEltWEFFQUFMQkRKbG9BQUV6YmhmNEJBR1RyVXJnQ0FBQmc5d1F0QUFBbXFxemIweERDc2Y0QkFHVGwrc24waWdldEJ3QUEyRDFCQ3dDQTZUclhPd0NBTEFoWEFBQUFSRVRRQWdCZ2dzcTZYWVVRRHZRT0FDQlp3aFVBQUFDUmVyTmVyL1VHQUdCQ3lycmREeUhjaFJEMjlBMEFJQ25DRlFBQUFCTmdvZ1VBd1BTY0Mxa0FBQ1JEdUFJQUFHQmlUTFFBQUppUXNtNlBRZ2ovMURNQWdFa1RyZ0FBQUpnd0V5MEFBS2JsWEw4QUFDWkp1QUlBQUNBUmdoWUFBQk5SMXUwOGhIQ3NYd0FBazNFYlFyZ1l3aFYzMmdZQUFKQUdxME1BQUNhZ3JOdjlFRUlYUWpqUUx3Q0FxQWxYQUFBQUpNNUVDd0NBYVZnSVdRQUFSRXU0QWdBQUlDTW1XZ0FBUks2czI4Tmhtc1dlWGdFQVJFTzRBZ0FBSUZNbVdnQUF4TzljeUFJQUlBckNGUUFBQUpob0FRQVFzN0p1VDBJSS85QWtBSURSQ0ZjQUFBRHdPeVphQUFERTdWeC9BQUIyVHJnQ0FBQ0FyeEswQUFDSVZGbTM4eERDVy8wQkFOaUoreUZjY1NGY0FRQUF3TGRZSFFJQUVLR3lidmREQ1AwRC9qMzlBUURZbWo1Y2NUV0VLenBsQmdBQTREbE10QUFBaU5OQ3lBSUFZQ3VFS3dBQUFIZ1ZFeTBBQUNKVDF1MWhDT0ZYZlFFQTJCamhDZ0FBQURiR1JBc0FnUGhjNkFrQXdLc0pWd0FBQUxBVmdoWUFBQkVwNi9Za2hIQ3NKd0FBUDBTNEFnQUFnSzBUdEFBQWlJdHBGZ0FBTHlOY0FRQUF3RTRKV2dBQVJLS3MyMFVJNFVBL0FBQys2M0VJcUFwWEFBQUFzSE52MXV1MXFnTUFqS3lzMi8wUXdsMElZVTh2QUFDKzZIR1lYSEhWVk1XVkVnRUFBREFXRXkwQUFPS3dFcklBQVBnVDRRb0FBQUNpWTZJRkFNREl5cm85Q2lIOFV4OUkzRmtJNFRDRWNDcFVCTUIzQ0ZjQUFBQVFOUk10QUFER2Q2NEhKTzZtcVlyVjU3ZFkxdTNwRUxnUXVnRGdTL3FmRFE4aGhFK3FBd0FBUUl4TXRBQUFHTkZ3NFB6LzlJREUvYldwaXU2UGI3R3MyLzBuZ1l0M1BnUUEvTUY5Q0dIZVZJWEFCUUFBQUZFUnRBQUFHRkZadDNjaGhBTTlJR0dYVFZYTXYvZjJoQzRBK0lZUElZUlZVeFVQaWdRQUFFQU1CQzBBQUVaUzFtMi9TbUdwL2lTczM3Ri8rTktEc1NlaGkwVUk0YTBQQ0FBaGhOdGh1c1dmSmlRQkFBREFyZ2xhQUFDTW9LemJ3eEJDTit3Z2gxU2ROVld4ZXMxN0cvNnQ5S0dMdWRBRkFKdjQyUUlBQUFDdkpXZ0JBRENDc200dlFnanYxWjZFM1RkVmNiakp0eWQwQWNEQWRBc0FBQUJHSldnQkFMQmpaZDJlaEJEK29lNGs3bTlOVlh6YTFsc3M2L1pvQ0Z6MHdZc0RIeWFBN1BUcnFWWk5WWnhyUFFBQUFMc21hQUVBc0dObDNmYUh6OGZxVHNKdW1xbzQyZFhiRTdvQXlOck5NTjNpTHZkQ0FBQUFzRHVDRmdBQU8xVFdiWDhZL0hjMUozRi9HZXZBYXdoZExJYlF4WjRQR2tBV0hvZXd4WlYyQXdBQXNBdUNGZ0FBTzFMVzdYNElvWFBqbnNSOWFLcGlFY05iTE92MmRBaGNDRjBBNU9GNkNGdzg2RGNBQUFEYkpHZ0JBTEFqWmQydVFnaEw5U1poL1kzaXd4Z1B1SVF1QUxKeFA0UXRQbWs1QUFBQTJ5Sm9BUUN3QTJYZEhvWVFmbFZyRXZkTFV4WG5NYi9GWWJMTTU4REZ1d2hlRWdEYmNkWlV4VXB0QVFBQTJBWkJDd0NBSFNqcjlzcWhMb203YmFyaWFFcHZVZWdDSUhtMy9mZjRwaXJ1dEJvQUFJQk5FclFBQU5peXNtNVBRZ2ovVUdjUzk3Y3BqMmwvRXJxWWh4Q09JM2hKQUd4R3Y5WnEwVlRGaFhvQ0FBQ3dLWUlXQUFCYlZ0WnRGMEo0cTg0azdMcXBpdE5VM3Q2dzZ1ZHo2TUsvWFlBMFhQZmYxNXVxZU5CUEFBQUFYa3ZRQWdCZ2k4cTY3UTlxLzY3R0pPNHZxWTVsRjdvQVNNcjlzRXFrMDFZQUFBQmVROUFDQUdCTGhsVUUvZUh6bmhxVHNMT21LbFk1TkhnSVhTeUc0TVZCQkM4SmdCK1R6Yzh1QUFBQXRrUFFBZ0JnUzhxNlBROGgvS3krSkt6ZmUzK1k0eGoyc202UGhpa1hRaGNBMDNRelRMZXdTZ1FBQUlBWEU3UUFBTmlDNGViN3IycEw0bjVxcXVJaTl5WS9DVjNNVGJBQm1KVEhJV3p4U2RzQUFBQjRDVUVMQUlBdEtPdTJmMkIvckxZazdLYXBpaE1OL3IyeWJrK0hLUmVuUWhjQWsyR1ZDQUFBQUM4aWFBRUFzR0ZsM2ZhSHovOVFWeEwzTnplQXYwM29BbUJTckJJQkFBRGcyUVF0QUFBMnJLemJ1eERDZ2JxU3NNdW1LdVlhL0h4bDNjNkh3TVc3cWJ4bWdBeFpKUUlBQU1DekNGb0FBR3hRV2JlTEVNTC9xQ2tKNncraGpwcXF1TlBrbHl2cmR2L0psQXVoQzRBNFdTVUNBQURBTndsYUFBQnN5SENBZW1kRkFJbHorTFFoVDBJWC9iU0w0eVRlRkVBNnJCSUJBQURncXdRdGdQL1AzdDBrMVhHa2JRTk9POTQ1ZkNzQXIwQjRCYUJKVG9XakluSXF0QUxqRmZob0JZMVdZSmhXUk1ZclRYTmlXRUhEQ2d3N0VDdmdpL0tiZHN0dVNTNStEcWNxNjdvaXZJQitiaHFPVHQ3NUpBQlBwT3ZMYVFqaHRYblNzSnVjNHE2QW4xN1hsOTFQU2hjdld2dmZCekJUbmhJQkFBRGdzeFF0QUFDZVFOZVh2UkRDdjgyU3h2MlFVM3d2NVBWU3VnQ1luSjl5aWlkaUFRQUE0QStLRmdBQVQ2RHJ5N25WL3pUdUlxZDRJT1RuVlVzWHg3VjRzYk9rLyswQUUvTmhLTUI1U2dRQUFJQ2dhQUVBOEhoZFg0WUQwUDgxU2hyM2ZVN3hVc2liVXpmbkhDbGRBR3pNVFgxS3hOOURBQUNBaFZPMEFBQjRoSzR2MnlHRVM0ZWVOTzRzcDNnazVPbjRwSFF4L0xlMTlIa0FQS1BiWWROUVR2SFUwQUVBQUpaTDBRSUE0Qkc2dnF4Q0NEK2JJUTBiRHBSMnJVcWZycnBWNTQvL2xDNEFuc2U3bk9LeFdRTUFBQ3lUb2dVQXdBTjFmZG10Mnl3Y2JOS3luM0tLSnhLZUI2VUxnR2QxRlVJNFVFWUVBQUJZSGtVTEFJQUg2dm95ckl4K2JYNDA3Q2FudUN2Z2VlcjZjbFFMRjYrV1BndUFOYnF0Wll0TFF3WUFBRmdPUlFzQWdBZm8rbklRUXZqVjdHamN5NXppdVpEbnJldkw5aWRiTHBRdUFOYmpUVTd4MUd3QkFBQ1dRZEVDQU9BQnVyNE1oOC83WmtmRExuS0tCd0p1eXllbGl5Ty93d0NlM0ZsTzhjaFlBUUFBMnFkb0FRQndUM1VkL3kvbVJ1Tyt5eWxlQzdsZFhWOTJQeWxkdkZqNlBBQ2V5RlY5U3VTamdRSUFBTFJMMFFJQTRCN3FiZkRoOEhuTDNHall1NXppc1lDWFEra0M0RW5kMXJMRnBiRUNBQUMwU2RFQ0FPQWV1cjZzUWdnL214a05HdzZIZHQzRVhhNWF1aml1eFl1ZHBjOEQ0SUdHdjZmSE9jVlRBd1FBQUdpUG9nVUF3RWoxOFBFMzg2SnhieHdLOFlldUwzdDF5NFhTQmNERDJCSUZBQURRSUVVTEFJQ1J1cjY4RHlHOE1pOGFkcFZUM0JNd24vTko2ZUxJODBrQTkvSmgrTjFwV3hRQUFFQTdGQzBBQUVibytuSVFRdmpWckdqY3k1eml1WkQ1SjExZkR1dVdpME9sQzRCUnJvYmZtVG5GYStNQ0FBQ1lQMFVMQUlBUnVyNWNoaEJlbUJVTis1QlRQQlF3OTZWMEFURGFiUWpoSUtkNGFXUUFBQUR6cG1nQkFQQVB1cjRNNzJyL3k1eG8ySER3cytlV0xZL1Y5ZVdvRmk0OHN3VHdaVzl5aXFmbUF3QUFNRitLRmdBQVg5SDFaVHVFY08yV05vMTdtMU5jQ1ptblVuOTNIaXBkQUh6UnU1emlzZkVBQUFETWs2SUZBTUJYZEgwNUNTSDhhRVkwN0tadXMvZ29aTlpCNlFMZ2l6NkVFSTc4RFFZQUFKZ2ZSUXNBZ0MvbytySWJRdmpOZkdpYzllVThtL3A3ZFNoY0RFK012REI1Z0hBVlFqaFF0Z0FBQUpnWFJRc0FnQy9vK25JZVF0ZzNIeHAya1ZNOEVEQ2JvSFFCOEtmYldyYTROQklBQUlCNVVMUUFBUGlNcmkvRDRmT3Zaa1BqWHVZVXo0WE1wdFhTeFZIOWIwY2d3QUlOWll0RGY1Y0JBQURtUWRFQ0FPQXp1cjVjTyt5amNXYzV4U01oTXpWZFgvWnE0ZUxRNzJGZ2dUenBCUUFBTUFPS0ZnQUFmOVAxWlJWQytObGNhTmh3YTNiWGUvQk1uZElGc0ZEdmNvckh3Z2NBQUpndVJRc0FnRTkwZmRrT0lRemJMTGJNaFlhOXpTbXVCTXljZEgwNXJJV0xRNytqZ1FXd2VRb0FBR0RDRkMwQUFEN1I5V1ZZMWZ6YVRHallUVTV4VjhETW1kSUZzQkJYSVlRREc2Z0FBQUNtUjlFQ0FLQ3FLK3IvYlI0MDdvZWM0bnNoMDRwYXVoaHVmYjhTS3RBZ1pRc0FBSUFKVXJRQUFLaTZ2cHlIRVBiTmc0WmQ1QlFQQkV5TDZ0TlBmMnk1VUxvQVduSmJ5eGFYVWdVQUFKZ0dSUXNBZ1A4N29CdHVRLzlpRmpUdWU0YzBMSUhTQmRBZ1pRc0FBSUFKVWJRQUFCYXZIc2dOWDFydkxIMFdOTzFkVHZGWXhDeE4xNWZkV3JnWUNuVXYvQUFBTXphVUxZNXppcWRDQkFBQTJDeEZDd0JnOGJxK3JFSUlQeTk5RGpSdE9Kalo5YjQ3UzZkMEFUVGlqYklGQUFEQVppbGFBQUNMVmcvZGhtMFdXMHVmQlUzN0thZDRJbUw0ai9yNy82aitaNk1STURmK3RnTUFBR3lRb2dVQXNHaGRYNGJiZ0srWFBnZWFkcE5UM0JVeGZGblhsNzFhdURoVXVnQm01Q3luZUNRd0FBQ0E1NmRvQVFBc1Z0ZVhneERDcjM0Q2FOekxuT0s1a0dFY3BRdGdacFF0QUFBQU5rRFJBZ0JZcks0dmw5N29wM0VYT2NVREljUERkSDA1cklXTFEwOU1BUk4yRmtJNHppbCtGQklBQU1EelVMUUFBQmFwNjh0dzgrOFg2ZE80NzNLSzEwS0d4NnQvTjRiQ3hTdmpCQ2JvS29Sd29Hd0JBQUR3UEJRdEFJREY2ZnF5SFVLNGRqdVp4cjNOS2E2RURFK3IvZzA1cXYvWmlnUk1pYklGQUFEQU0xRzBBQUFXcCt2TGNQajhzK1JwMkcwSVlkZEJDNnhYMTVlOVQwb1h5bnZBRkNoYkFBQUFQQU5GQ3dCZ1VicSs3SVlRZnBNNmpYdVRVendWTWp3ZlQ0c0FFNkpzQVFBQXNHYUtGZ0RBb25SOWVlOFFqTVpkNVJUM2hBeWJVUXQ5ZjJ5NTJCRURzQ0czdFd4eEtRQUFBSUNucDJnQkFDeEcxNWVERU1LdkVxZHhMM09LNTBLR3pldjZjbGdMRndwK3dDWW9Xd0FBQUt5Sm9nVUFzQmhkWDY3ZExxWnhaem5GSXlIRHROaHlBV3lRc2dVQUFNQWFLRm9BQUl2UTllVTRoUEF2YWRPdzRTQmxMNmQ0TFdTWXJxNHZmeFF1OXNVRVBCTmxDd0FBZ0NlbWFBRUFOSy9yeTNZSVlUaDgzcEkyRFh1YlUxd0pHT2FoYnJrNHJxVUxmNStBZFZPMkFBQUFlRUxmR2lZQXNBQXJoMWcwN2lhRWNDSmttSTloKzB4T2NTaGFESVdMTi9YL3h3RHJNbndXUHUvNnNtZkNBQUFBajJlakJRRFF0SHBqK0RjcDA3ZzNPY1ZUSWNPOGRYMDVxRnN1WG9rU1dCT2JMUUFBQUo2QW9nVUEwTFN1TCtmZXdhZHhGem5GQXlGRE96d3JBcXlac2dVQUFNQWpLVm9BQU0zcStuSVlRdmhmQ2RPNDd4MlVRSnU2dm15SEVBN3JFMWc3WWdhZWtMSUZBQURBSXloYUFBRE42dnB5N1dDS3hwM2xGSStFRE8ycjVjRmpXNXFBSjZSc0FRQUE4RURmR2h3QTBLS3VMMjcvMHJyYmV1Z0tMRUJPOFgxOUp1ajdvV1FsYytBSkRFOFRuWGQ5MlROTUFBQ0ErN0hSQWdCb1RsMjFmdTFkZXhyM05xZTRFaklzVTllWDNSRENVUzFjK1hzSFBJYk5GZ0FBQVBla2FBRUFOS2ZyeTJrSTRiVmthZGhOVG5GWHdFQXRGLzVSdUxESkNYZ29aUXNBQUlCN1VMUUFBSnJTOVdWWXEvNnJWR25jRDhNekFrSUdQdFgxNVkvQ3hRdURBUjVBMlFJQUFHQWtSUXNBb0NsZFg4NURDUHRTcFdFWE9jVURBUU5mVWt1SEszOFBnUWRRdGdBQUFCaEIwUUlBYUVhOXlmdUxSR25jZHpuRmF5RUQvNlFXTG80OHB3WGMwMUMyMlBONUF3QUE0TXNVTFFDQUp0UTM2aSs5VDAvajN1VVVqNFVNM0VmWGw5MjY0VUxoQWhqcnFtNjIrR2hpQUFBQS8rMWJNd0VBR25Hc1pFSGpidXRCS2NDOURMZlNjNHJEWm92dlFnaG45ZmNKd05lOENDR2Mxekl6QUFBQWYyT2pCUUF3ZS9XbTdyRE5Za3VhTk95bm5PS0pnSUhIcWdlbngvVS9menVCcjdIWkFnQUE0RE1VTFFDQTJldjY4ajZFOEVxU05Pd3FwN2duWU9BcEtWd0FJeWxiQUFBQS9JMm5Rd0NBV2V2NmNxQmt3UUljQ3hsNGFzT2hhVTV4ZUpKbzJBejExcE1pd0JjTXo0allxZ1VBQVBBSkd5MEFnRm5yK25KWnYveUZWbjNJS1I1S0YxZzNHeTZBZjNDV1V6d3lKQUFBQUVVTEFHREd1cjRNWC9UK0lrTWE5MTFPOFZySXdIT3BoWXVoNERWc3U5Z3hlT0FUeWhZQUFNRGlCVVVMQUdDdTZpSFF0UnUzTk81dFhlc1BzQkcxMUtod0FYenFwNXlpcDBRQUFJQkYrM2JwQXdBQVpzdGFjMXAzNnoxMFlOTnlpcWM1eGQyaCtGVi9Md0g4cTVhd0FBQUFGc3RHQ3dCZ2RycStEQWMrdjBtT3hyMFpEamlGREV4RjNTWjFyT3dJVkQ2ckFBQUFpNlZvQVFETVR0ZVg4eERDdnVSbzJFVk84VURBd0JRcFhBRFZzT1htSUtkNGFTQUFBTURTS0ZvQUFMUFM5V1U0ZlA1VmFqVHVaVTd4WE1qQWxDbGNBTW9XQUFEQVVpbGFBQUN6MHZYbE9vU3dJelVhZHBaVDlPNDVNQnNLRjdCNE55R0V2WnppeDZVUEFnQUFXSTV2WlEwQXpFWFhsMk1sQ3hvMzNBcGRDUm1ZaytGd05hYzQvTzdhRzhwaXdvUEZHVDZmbjlmU0ZRQUF3Q0lvV2dBQXMxQy91SFVBVGV0T2NvclhVZ2JtYVBqOVZUZnlmS2R3QVl2eklvUndLbllBQUdBcEZDMEFnTGxZV1VkTzQ0YTEyeWRDQnVidWI0V0xDNEhDWXJ6cStxSnNBUUFBTE1JM2QzZDNrZ1lBSnEzcnk3Q0svTjlTb25FLzVCVGZDeGxvVGRlWGcxcVkzQmN1TE1KUE9VWGxVUUFBb0dtS0ZnREE1SFY5T1hjNFErTXVjb29IUWdaYTF2WGxzRzd1MlJFME5POU5UdEYyQ3dBQW9GbUtGZ0RBcE5WRG1mK1ZFbzM3UHFkNEtXUmdDYnErSE5VTkZ3b1gwSzdiRU1LQnp6Y0FBRUNydnBVc0FEQngxZzdUdWpPSEVNQ1MxRnZ1dzdOZ2IrdGhMTkNlclJEQ2VkZVhYZGtDQUFBdFVyUUFBQ2FyNjR2YnJyUnVPR0E4bGpLd05EbkZqem5GNGUvOGNBaDc1Z2NBbWpTVUxkNTNmZGtXTHdBQTBCcEZDd0Jna3VydE53ZlF0RzQxSERaS0dWaXFXcmdZbmhMNUxvUnc0UWNCbXZNaWhIQXFWZ0FBb0RXS0ZnREFWSzNxTFRobzFVMU8wZE00QVA5WHVMak9LUjZFRUY2R0VLN01CSnJ5cXV1THp6d0FBRUJURkMwQWdNbnArakljdEx5V0RJMDdFakRBWCtVVXozT0tleUdFTjBNaHpYaWdHVDkyZmZIWkJ3QUFhSWFpQlFBd1JTdXAwTGlMNFRCUnlBQ2ZsMU1jbmhvWUNoZHZRd2kzeGdSTitLWHJ5NTRvQVFDQUZueHpkM2NuU0FCZ011cE50MThrUXVPK0c5YmtDeG5nbjNWOTJhMGxUTnV1WVA2RzR0UnVUdkdqTEFFQWdEbXowUUlBbUl5dUw5dTJXYkFBNzVRc0FNWWJmbWZtRkljaTVzdGhJNURSd2F4dGhSQnM5UUlBQUdaUDBRSUFtSkxqRU1LT1JHallyVElSd01NTVR5N2xGQTlDQ0c4OEp3S3o5cUxyeTZrSUFRQ0FPZk4wQ0FBd0NYVXQrRy9Tb0hGdmNvb09GZ0FlcVc3QkdncWFQNXNsek5aUE9jVVQ4UUVBQUhPa2FBRUFURUxYbC9jaGhGZlNvR0ZYT2NVOUFRTThuVnJVSEFwcys4WUtzL1J5MkZZak9nQUFZRzRVTFFDQWpldjZNcXdCLzFVU05NNUJBc0NhMU04U3A1NGdnOWtabmdIYXpTbCtGQjBBQURBbjMwb0xBSmdBSzROcDNRY2xDNEQxR1g3SDVoU0g3Ulp2NjhFdE1BOWJJUVNma1FBQWdObFJ0QUFBTnFycnkxRUk0WVVVYU55eGdBSFdMNmU0Q2lFTXp6UjlNRzZZalJkZFgwN0ZCUUFBekltblF3Q0FqZW42c2gxQ3VLNDMyYUJWYit2Qkh3RFB5SE1pTUR0dmNvb0tGd0FBd0N6WWFBRUFiTkpLeVlMRzNYZ2FCMkF6L3ZhY0NEQjlKMTFmOXVRRUFBRE1nWTBXQU1CR2RIMFpEajUrTTMwYTUyWW13QVRVengzRDcrTjllY0NrWFlVUURuS0tIOFVFQUFCTW1ZMFdBTUNtT0h5bWRSZEtGZ0RUa0ZPOHppa09UNG44VUxjTkFkUDB3cjhUQUFDQU9WQzBBQUNlWFgwejNZMVNXcmVTTU1DMDVCVGZoeENHcHduZWlRWW02MVhYbDJQeEFBQUFVK2JwRUFEZzJYVjl1UTRoN0pnOERUdkxLUjRKR0dDNnVyN3MxWnZ6TDhRRWsvUjlUdkZTTkFBQXdCVFphQUVBUEt1dUx5c2xDeHAzYTVzRndQUU5CN2c1eGFGczhWUDkzUTFNeS91dUw5c3lBUUFBcGtqUkFnQjROdldMVW11QWFkMUpUdkZheWdEemtGTThxYytKWElnTUptV25icDBCQUFDWUhFK0hBQURQcHV2TDhFWHBheE9uWVRjNXhWMEJBOHhUMTVmRGVyQzdKVUtZakRjNVJZVUxBQUJnVW15MEFBQ2VSWDBIWGNtQzF0bllBakJqT2NYM0lZU2hNSGNtUjVpTWsvcHZDUUFBZ01tdzBRSUFlQlpkWDg1RENQdW1UY011Y29vSEFnWm9ROWVYZzdyZFlrZWtzSEZYSVlTRG5PSkhVUUFBQUZOZ293VUFzSFoxRGJlU0JhMnp6UUtnSVRuRm9TUTYzS0ovSjFmWXVCY2hoSlVZQUFDQXFiRFJBZ0JZcTY0djJ5R0VTN2RCYWR5N25LS2lCVUNqYkxlQXlmaWhQdkVEQUFDd1VUWmFBQURyZHV4UWdzYmR1bUVKMExaUHRsdThGVFZzMUdrdGNnTUFBR3lVb2dVQXNEWmRYM1k5cDhBQ3JMd1hEdEMrNFhkOVRuRW8xbjBmUXJnU09XekVWZ2pCUmdzQUFHRGpGQzBBZ0hWYTFTOURvVlUzT2NVVDZRSXNSMDd4TXFkb3V3VnN6bjdYRjJWdUFBQmdveFF0QUlDMXFHK1p2elpkR25ja1lJQmxzdDBDTm1yVjlXVlBCQUFBd0tZb1dnQUE2N0l5V1JwM1VkL3NCMkNoaHUwV0lZU2hYUHJPendBOHEyRnIzcW1SQXdBQW02Sm9BUUE4dWE0dnd5My9mWk9sY2JaWkFEQ1VMVDdtRklkbkRGNE9UMHFaQ0R5YkYxMWZsTHNCQUlDTitPYnU3czdrQVlBbjAvVmxPNFJ3WFcrWlFhdmUxcFh4QVBDbitqbG8rUHZ3bzZuQXMvbSticGNCQUFCNE5qWmFBQUJQN1ZqSmdzYmRoaEJPaEF6QTMzMnkzZUtIK3ZjQ1dEOVBpQUFBQU05TzBRSUFlREpkWDNaRENEK2JLSTA3SGc3U2hBekFsK1FVMzRjUWhzOUZId3dKMW01NFFrUUpGZ0FBZUZhZURnRUFua3pYbCtGUTRaV0owckNybk9LZWdBRVlxK3ZMY1gxT3hNWXZXSytYT2NWek13WUFBSjZEalJZQXdKUG8rbktnWk1FQ0hBc1pnUHZJS1E0MzdZZVMzcFhCd1ZxZGRuM1pObUlBQU9BNUtGb0FBRS9GdWw1YTk4RXRTUUFlSXFkNFhUY2l2VFZBV0pzZHBWZ0FBT0M1S0ZvQUFJOVdWMksvTUVrYWR1dUxld0FlSzZjNFBDSHlNb1J3WTVpd0ZqOTNmZkhNR3dBQXNIYUtGZ0RBbzlUMXZDdFRwSEVudzIxa0lRUHdXSFU3MG5BUS9NRXdZUzFPalJVQUFGZzNSUXNBNExHR2tzV1dLZEt3RzAvakFQQ1Vjb29mYzRxSElZUTNkV3NTOEhSZWRIMVJCQWNBQU5icW03dTdPeE1HQUI2azY4dHVDT0UzMDZOeGIzS0tia1lDc0JiMW1ZTlR6N0RCay9zK3AzaHByQUFBd0RyWWFBRUFQSWJEWjFwM29XUUJ3RHJWZytDREVNSTdnNFluWlNNWkFBQ3dOb29XQU1DRGRIMFoxbDN2bXg2Tk94WXdBT3RXbnhJWi91Yjg0Q2tSZURMN1hWOThsZ01BQU5aQzBRSUFlQ2czeEdqZG1YWFRBRHlubk9MN0VNTHdsTWlWd2NPVFdIVjkyVFpLQUFEZ3FTbGFBQUQzMXZWbEZVTFlNVGthZG11YkJRQ2JrRk84emludWVVb0Vuc1NXNXc0QkFJQjFVTFFBQU82bDNnaHpBRTNyVG9ZMTdsSUdZRk04SlFKUDVsWFhsd1BqQkFBQW5wS2lCUUJ3WHlmMVpoaTA2aWFudUpJdUFKdFdueEk1OEpRSVBOcXBKMFFBQUlDbnBHZ0JBSXpXOVdWWVkvM2F4R2ljalMwQVRFWk84YktXTGM2a0FnKzI0ek1lQUFEd2xMNjV1N3N6VUFCZ2xLNHY1eUdFZmRPaVlSYzVSYXVsQVppa3JpOUhJWVJmcEFNUDlsMU84ZHI0QUFDQXg3TFJBZ0FZcFg2eHIyUkI2NDRrRE1CVTVSUlBRd2pmRDg5Y0NRa2U1TlRZQUFDQXA2Qm9BUUQ4by9xZThjcWthTnc3Tnh3Qm1McjZsTWp3bk51RnNPRGU5cnUrSEJvYkFBRHdXSW9XQU1BWXgvVmRZMmpWclRJUkFIT1JVL3hZbjdwNkp6UzR0NU5hSkFjQUFIZ3dSUXNBNEt1NnZ1eldvZ1cwYkRVY1dra1lnRG5KS1E2ZjBkN1V3aUF3em81LzN3QUFBSStsYUFFQS9KUGhsditXS2RHd201emlpWUFCbUtPYzRta0lZZGh1Y1NOQUdPMjRGc29CQUFBZVJORUNBUGlpcmkvRGwvYXZUWWpHSFFrWWdEbkxLVjZHRVBaQ0NCZUNoRkcyUEJzSEFBQThocUlGQVBBMWJ2blR1Zzg1eFhNcEF6QjN3eE5ZT2NXaEpQdE9tRERLNjFvc0J3QUF1RGRGQ3dEZ3M3cStETGY4WDVnT2pmTStOd0JOeVNrT2Y5dmVTQlZHc2RVQ0FBQjRFRVVMQU9DL2RIM1p0czJDQlhpYlU3d1dOQUN0eVNtZWhoQytEeUhjQ2hlK2FyOFd6QUVBQU81RjBRSUErSnpqK200eHRPcFdtUWlBbHVVVUwwTUlleUdFSzBIRFY2MXEwUndBQUdBMFJRc0E0Qys2dnV5R0VINDJGUnAzUEx4bEwyUUFXbFkzTngyRUVENElHcjVveDNOeUFBREFmWDF6ZDNkbmFBREFuN3ErdkE4aHZESVJHbmFWVTl3VE1BQkwwdlZsMk9UMG85RGhzNFp0Wjd1S3VBQUF3RmcyV2dBQWYrcjZjcUJrd1FLNHNRakE0dVFVaDc5L2J5UVBuN1hsV1RrQUFPQStGQzBBZ0UrZG1nYU5POHNwbmdzWmdDWEtLUTZmOVY3VzIvdkFYNzJ1enlnQ0FBRDhJMFVMQU9CM1hWK082L3ZFMEtyaFVHa2xYUUNXckJZT0Q1UXQ0TE44VmdRQUFFWlJ0QUFBaHBMRnRpOFZXWUNUbk9LMW9BRll1cHppWlFoaHVMbC90ZlJad04rOHJzOHBBZ0FBZkpXaUJRQVFhc2xpeXlSbzJCWk5IdFVBQUJiVVNVUkJWSTEzdHdIZ1AzS0tIK3RtaXd0amdiOVFRQWNBQVA3Uk4zZDNkNllFQUF2VzlXVXZoUEJ2UHdNMDdvZWM0bnNoQThCLzYvcHlPdHprTnhyNDA4djZ6QTRBQU1CbjJXZ0JBTGpsVCtzdWxDd0E0TXR5aWtjaGhIZEdCSCt5MVFJQUFQZ3FSUXNBV0xDdUw0Y2hoSDAvQXpUdVdNQUE4SFU1eGVIdjVSdGpndC90MTM4ckFRQUFmSmFpQlFBc20yMFd0TzRzcDNncFpRRDRaem5GVTJVTCtKTi9Ld0VBQUYra2FBRUFDOVgxWlZpSHV5Ti9HblpybXdVQTNFOHRXM3hmLzQ3Q2t1MTBmVG55RXdBQUFIeU9vZ1VBTEZEWGwyMEgwQ3pBU1U3eG82QUI0SDdxTnFnRFpRc0lLeU1BQUFBK1I5RUNBSlpwV0lPN0pYc2FkcE5UOU1VNEFEeVFzZ1g4emxZTEFBRGdzeFF0QUdCaHVyNE1YNWkvbGp1Tjg0VTRBRHhTTFZ2c2hSQ3V6SklGVTk0RkFBRCtpNklGQUN5UEx3cHAzVVZPOFZ6S0FQQjRPY1hydXRsQzJZS2xzdFVDQUFENEw0b1dBTEFnOVF2Q2ZablRPRitFQThBVHlpbCtWTFpnNFpUVkFRQ0F2MUMwQUlDRjZQcXk3UXRDRnVCZHZYa0xBRHdoWlFzV2J0aHFjYmowSVFBQUFQK2hhQUVBeTNFOGZFRW9ieHAycTB3RUFPdWpiTUhDSFM5OUFBQUF3SDhvV2dEQUFuUjkyZlhGSUF1d3FnZEFBTUNhS0Z1d1lQdGRYdzc4QUFBQUFFSFJBZ0FXNHlTRXNDVnVHbmFWVXp3Uk1BQ3MzeWRsaXcvR3pjTFluZ1lBQVB4TzBRSUFHbGR2WGIyU000MnpzUVVBbnRGUXRzZ3BIb1lRenN5ZEJiSFZBZ0FBK0oyaUJRQzB6eTEvV3ZjaHAzZ3VaUUI0ZmpuRkkyVUxGdVpJNEFBQWdLSUZBRFNzNjh2d0plQUxHZE00Mnl3QVlJTnEyZUpDQml6RTY2NHZ1OElHQUlCbFU3UUFnRVoxZmRtMnpZSUZlSnRUdkJZMEFHemM4SXpJbFJoWWlKV2dBUUJnMlJRdEFLQmR3eTMvTGZuU3NCdGxJZ0NZaHB6aXh4RENnYklGQy9HNkZ0c0JBSUNGVXJRQWdBYlZWYlkveTViR3JlcWhEZ0F3QWNvV0xJem42d0FBWU1FVUxRQ2dUYWR5cFhFWE9VVS81d0F3TWNvV0xJaWlCUUFBTEppaUJRQTBwdXZMOE1YMnZseHBuSGV4QVdDaWxDMVlpSzJ1TDBmQ0JnQ0FaVkswQUlEMnVPVlA2ODV5aXVkU0JvRHArcVJzY1NNbUdtYXJCUUFBTEpTaUJRQTBwT3ZMOEVYZmpreHAySzF0RmdBd0Q3VnNjVmovZmtPTFh0U05nZ0FBd01Jb1dnQkFJN3ErYkR1QVpnRk9jb3JYZ2dhQWVjZ3BYdGJORnNvV3RNcnpJUUFBc0VDS0ZnRFFqcFBobldCNTByQ2JuS0l5RVFETWpMSUZqWHZkOVdWWHlBQUFzQ3lLRmdEUWdLNHZlOE1YZkxLa2NkN0FCb0NacW1VTGY4dHBsYTBXQUFDd01Jb1dBTkNHRXpuU3VJdWM0bnNoQThCODVSUlBRd2h2UkVpRGxJZ0FBR0JoRkMwQVlPYTZ2aHlHRVBibFNPTjhlUTBBRGFobGk3ZXlwREZiWFY5c3RRQUFnQVZSdEFDQStiUE5ndGFkMVhYakFFQURjb3FyNGUrN0xHbU1vZ1VBQUN5SW9nVUF6RmpYbCtGTDZoMFowckJiMnl3QW9EMDV4ZUZRK2tLME5HUy82OHV1UUFFQVlCa1VMUUJncHVxWGVBNmdhZDBxcC9oUnlnRFFwT0VKdkN2UjBoRC9QZ01BZ0lWUXRBQ0ErUnEyV1d6Smo0YmQ1QlE5alFNQWphcGx5c082d1FwYTRQa1FBQUJZQ0VVTEFKaWhyaThISVlUWHNxTnh2cWdHZ01ibEZLOURDQWZLRmpSaXErdUx6N0FBQUxBQWloWUFNRThydWRHNGk1eml1WkFCb0gwNXhVdFBMdENRUTJFQ0FFRDdGQzBBWUdicURhbDl1ZEU0TndFQllFRnlpcWNoaExjeXB3R3Z1cjdzQ2hJQUFOcW1hQUVBTTlMMVpkczJDeGJnWFYwakRnQXNTRTV4K0p6N1FlWTB3RllMQUFCb25LSUZBTXpMc0ZKNVIyWTA3RmFaQ0FBV2JkaHFkYlgwSVRCN25zSUJBSURHZlhOM2R5ZGpBSmlCdW43Mk4xblJ1RGQxZFRnQXNGRDFjKzlsQ0dITHp3QXo5bjFPOFZLQUFBRFFKaHN0QUdBK1RtUkY0NjZVTEFDQStvU1lweGVZTzFzdEFBQ2dZWW9XQURBRFhWOE9RZ2l2WkVYamZCa05BUHd1cDNnZVF2akpOSmd4WlNFQUFHaVlvZ1VBeklOdEZyVHVRejFRQVFENFhVNXgrQXo4d1RTWXFhMnVMOG9XQUFEUUtFVUxBSmk0cmkvRExmOFhjcUpodDdaWkFBQmZjQlJDdURFY1prclJBZ0FBR3ZYTjNkMmRiQUZnb3JxK2JJY1FoamVxdDJSRXc5N21GRmNDQmdBK3ArdkxYZ2poM0dkaVp1ci81UlEvQ2c4QUFOcGlvd1VBVE52S0Y4bzA3c2JUT0FEQTErUVVMMjIvWXNac3RRQUFnQVlwV2dEQVJIVjkyUTBoL0NnZkdyZHl3dzhBK0NjNXhkTVF3cGxCTVVPS0ZnQUEwQ0JGQ3dDWXJsUFowTGlMZW1nQ0FEREdzTlhpeXFTWW1WZjFTVWdBQUtBaGloWUFNRUZkWHc1Q0NQdXlvWEVyQVFNQVk5VXRXRWNHeGd6WmFnRUFBSTFSdEFDQWFYTExuOWFkNVJUUHBRd0EzRWRPOFRLRThKT2hNVE9LRmdBQTBCaEZDd0NZbUs0dnd5My9IYm5Rc051Nitoc0E0TjV5aWlmREUyUW14NHg0UGdRQUFCcWphQUVBRTFLL2ZITUFUZXRPNnVwdkFJQ0hPcXpsVFpnTFd5MEFBS0FoaWhZQU1DM0Q3Ynd0bWRDd201emlTc0FBd0dQVTB1YVJJVElqaWhZQUFOQVFSUXNBbUlpdUwzc2hoTmZ5b0hFMnRnQUFUeUtuK0Q2RWNHYWF6TVFyUVFFQVFEc1VMUUJnT2s1a1FlTXU2b0VJQU1CVEdVcWNONmJKSEhSOXNkVUNBQUFhb1dnQkFCTlF2M0RibHdXTnM4MENBSGhTbmhCaFpoUXRBQUNnRVlvV0FMQmhYVisyYmJOZ0FkN2xGQzhGRFFBOHRaemkrZkJadzJDWkFVVUxBQUJvaEtJRkFHemVjTXQvUnc0MDdEYUVzQkl3QUxCR0swK0lNQU5iWFYvMkJBVUFBUE9uYUFFQUc5VDFaZGR6Q2l6QXFxNzFCZ0JZQzArSU1DTzJXZ0FBUUFNVUxRQmdzNGFiZDFzeW9HRTNPVVZQNHdBQWExZWZFRGt6YVNaTzBRSUFBQnFnYUFFQUc5TDE1U0NFOE5yOGFaeWJwUURBY3pxdXo1YkJWTDJvbXcwQkFJQVpVN1FBZ00xWm1UMk4rMUJ2bGdJQVBJdjZoSWluK1ppNkF3a0JBTUM4S1ZvQXdBWjBmUmx1K2UrYlBZMXp5QUVBUEx1YzRta0k0Y0xrbVRCRkN3QUFtRGxGQ3dCNFpsMWZ0a01JSitaTzQ5N21GSytGREFCc2lPZkxtTEpENlFBQXdMd3BXZ0RBOHh0dStXK1pPdzI3VlNZQ0FEYXBGajdmQ29HSjJ1cjZzaWNjQUFDWUwwVUxBSGhHWFY5MlF3Zy9tem1OTzY3dm93TUFiTkpRL0x5UkFCUGwrUkFBQUpneFJRc0FlRjV1K2RPNnEvb3VPZ0RBUnRYaTU3RVVtQ2pQaHdBQXdJd3BXZ0RBTStuNk10eFllbVhlTk01aEJnQXdHVG5GOXlHRUM0a3dRZnRDQVFDQStWSzBBSURuNDVZL3JUdkxLWjVMR1FDWUdFVlFKcW1XOFFFQWdCbFN0QUNBWjlEMVpmaHlkOGVzYWRodENHRWxZQUJnYW5LS2wwTWhWREJNa0tJRkFBRE1sS0lGQUt4WjE1ZHRCOUFzd0VsTzhWclFBTUJFcldveEZLWkUwUUlBQUdaSzBRSUExbS80VW5mTG5Hbll6VkMwRURBQU1GVzFFT3J6Q2xPekx4RUFBSmduUlFzQVdLT3VMN3NoaEIvTm1NYXRjb29maFF3QVROeUpyUlpNVGRjWFd5MEFBR0NHRkMwQVlMMU96WmZHWGVRVS9ad0RBSk5YaTZHMldqQTFpaFlBQURCRGloWUFzQ1pkWHc2dGdtVUJqb1VNQU14RlRuRlZuejJEcVZDMEFBQ0FHVkswQUlEMWNWdU8xcDNsRkMrbERBRE16RXBnVEloeVBnQUF6SkNpQlFDc1FkZVg0Y3ZiSGJPbFliZTJXUUFBYzFTZlBiUFZnc25vK3JJbkRRQUFtQmRGQ3dCNFlsMWZ0aDFBc3dBbjlaMXpBSUE1c3RXQ0tmRjhDQUFBekl5aUJRQTh2ZUhKa0MxenBXRTM5WDF6QUlCWnN0V0NpYkhSQWdBQVprYlJBZ0NlVUYzNSt0cE1hZHlSZ0FHQUJpaU9NaFUyV2dBQXdNd29XZ0RBMHpveFR4cDNrVk04RnpJQU1IZDFxOFd0SUptQW5mb0VKUUFBTUJPS0ZnRHdSTHErRExmODk4MlR4dGxtQVFDMFJGR2FxZkI4Q0FBQXpJaWlCUUE4Z1hyN3lPcGhXdmN1cDNndFpRQ2dJU2UyV2pBUm5nOEJBSUFaVWJRQWdLZHhQS3g3TlVzYWRxdE1CQUMwSnFmNE1ZUndLbGdtd0VZTEFBQ1lFVVVMQUhpa3JpKzd0V2dCTFZ2Vmd3Z0FnTlo0UG9RcFVMUUFBSUFaVWJRQWdNY2Jidmx2bVNNTnU4b3BPb0FBQUpwVW4wWTdreTRidGxPZnBBUUFBR1pBMFFJQUhxSHJ5L0NPN21zenBIRTJ0Z0FBcmZOOENGTmdxd1VBQU15RW9nVUFQSTViL3JUdVEwN3hYTW9BUU12cTU1MHJJYk5oQndJQUFJQjVVTFFBZ0FmcStuSVVRbmhoZmpUT05nc0FZQ21VcU5tMFhRa0FBTUE4S0ZvQXdBUFV0M045RVV2cjN0WTN5d0VBbHVCOUNPRlcwbXlRcDBNQUFHQW1GQzBBNEdHR1cvNWJaa2ZEYnBXSkFJQWx5U2wrckdVTDJCUWJFd0VBWUNZVUxRRGducnErRE90Y2Z6WTNHbmRjRHhzQUFKWkUwWlNONnZwaXF3VUFBTXlBb2dVQTNOK3BtZEc0aTV5aW4zTUFZSEZ5aXBjaGhCdkpzMEc3aGc4QUFOT25hQUVBOTlEMTVTQ0VzRzltTkc0bFlBQmd3V3kxWUpOc3RBQUFnQmxRdEFDQSszSExuOWFkNVJUUHBRd0FMTmg3NGJOQmloWUFBREFEaWhZQU1GTFhsK01Rd281NTBiQmIyeXdBZ0tYTEtWNkhFRDRzZlE1c2pLZERBQUJnQmhRdEFHQ0VyaS9iRHFCWmdKTjZzQUFBc0hTMldyQXBMMHdlQUFDbVQ5RUNBTVlaU2haYlprWERicnhIRGdEd0owVUxOcWJyaTYwV0FBQXdjWW9XQVBBUHVyNE1iK1QrYUU0MDdqaW4rRkhJQUFDL1B4L3kwZk1oYkpDaUJRQUFUSnlpQlFEOE03ZjhhZDFGVHRHdFRRQ0F2L0w1aUUwNU1Ia0FBSmcyUlFzQStJcXVMNGNoaEgwem9uSEhBZ1lBK0MrS0Ztekt0c2tEQU1DMEtWb0F3TmZaWmtIcnpuS0tsMUlHQVBpcitueklsYkd3QVh1R0RnQUEwNlpvQVFCZjBQVmxGVUxZTVI4YWRtdWJCUURBVjUwYUR4dGdvd1VBQUV6Yy93Z0lBTDdvT29UdzFuaG8yR1c5cVFrQXdPZTlkK2dOQUFEQTMzMXpkM2RuS0FBQUFBQUFBQUFBSTNnNkJBQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBQUFBQUFBQUFBRVpTdEFBQUFBQUFBQUFBR0VuUkFnQUFBQUFBQUFCZ0pFVUxBQUFBQUFBQUFJQ1JGQzBBK1AvdDJvRUFBQUFBdzZEN1UxOWhBTVVSQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQVVHdzdFYnp4RFZhTzlwTUFBQUFBU1VWT1JLNUNZSUk9JzsKZnVuY3Rpb24gYXBwbHlCcmFuZGluZygpIHsKICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcuYnJhbmQtbG9nbycpLmZvckVhY2goKGltZykgPT4geyBpbWcuc3JjID0gTE9HT19EQVRBX1VSSTsgfSk7CiAgY29uc3QgZmF2aWNvbiA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmYXZpY29uTGluaycpOwogIGlmIChmYXZpY29uKSBmYXZpY29uLmhyZWYgPSBMT0dPX0RBVEFfVVJJOwp9CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgQXBpIOKAlCB0aGluIGZldGNoIHdyYXBwZXJzIGFyb3VuZCB0aGUgUkVTVCBBUEkuCiAgID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwpjb25zdCBBcGkgPSAoKCkgPT4gewogIGFzeW5jIGZ1bmN0aW9uIHJlcXVlc3QocGF0aCwgb3B0aW9ucykgewogICAgY29uc3QgcmVzID0gYXdhaXQgZmV0Y2gocGF0aCwgb3B0aW9ucyk7CiAgICBsZXQgYm9keTsKICAgIHRyeSB7CiAgICAgIGJvZHkgPSBhd2FpdCByZXMuanNvbigpOwogICAgfSBjYXRjaCB7CiAgICAgIGJvZHkgPSBudWxsOwogICAgfQogICAgaWYgKHJlcy5zdGF0dXMgPT09IDQwMSAmJiAhcGF0aC5zdGFydHNXaXRoKCcvYXBpL2F1dGgvJykpIHsKICAgICAgd2luZG93LmRpc3BhdGNoRXZlbnQobmV3IEN1c3RvbUV2ZW50KCdscnM6c2lnbmVkLW91dCcpKTsKICAgIH0KICAgIGlmICghcmVzLm9rKSB7CiAgICAgIGNvbnN0IG1lc3NhZ2UgPSAoYm9keSAmJiBib2R5LmVycm9yKSB8fCBgUmVxdWVzdCBmYWlsZWQgKCR7cmVzLnN0YXR1c30pYDsKICAgICAgdGhyb3cgbmV3IEVycm9yKG1lc3NhZ2UpOwogICAgfQogICAgcmV0dXJuIGJvZHk7CiAgfQoKICBmdW5jdGlvbiBxcyhwYXJhbXMpIHsKICAgIGNvbnN0IHVzcCA9IG5ldyBVUkxTZWFyY2hQYXJhbXMoKTsKICAgIE9iamVjdC5lbnRyaWVzKHBhcmFtcyB8fCB7fSkuZm9yRWFjaCgoW2ssIHZdKSA9PiB7CiAgICAgIGlmICh2ICE9PSB1bmRlZmluZWQgJiYgdiAhPT0gbnVsbCAmJiB2ICE9PSAnJykgdXNwLnNldChrLCB2KTsKICAgIH0pOwogICAgY29uc3QgcyA9IHVzcC50b1N0cmluZygpOwogICAgcmV0dXJuIHMgPyBgPyR7c31gIDogJyc7CiAgfQoKICByZXR1cm4gewogICAgYXV0aE1lOiAoKSA9PiByZXF1ZXN0KCcvYXBpL2F1dGgvbWUnKSwKICAgIGF1dGhMb2dpbjogKGNvZGUpID0+CiAgICAgIHJlcXVlc3QoJy9hcGkvYXV0aC9sb2dpbicsIHsgbWV0aG9kOiAnUE9TVCcsIGhlYWRlcnM6IHsgJ0NvbnRlbnQtVHlwZSc6ICdhcHBsaWNhdGlvbi9qc29uJyB9LCBib2R5OiBKU09OLnN0cmluZ2lmeSh7IGNvZGUgfSkgfSksCiAgICBhdXRoTG9nb3V0OiAoKSA9PiByZXF1ZXN0KCcvYXBpL2F1dGgvbG9nb3V0JywgeyBtZXRob2Q6ICdQT1NUJyB9KSwKCiAgICBmaWx0ZXJPcHRpb25zOiAoKSA9PiByZXF1ZXN0KCcvYXBpL2FuYWx5dGljcy9maWx0ZXItb3B0aW9ucycpLAogICAga3BpczogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9hbmFseXRpY3Mva3BpcyR7cXMocGFyYW1zKX1gKSwKICAgIHBsYXRmb3JtQnJlYWtkb3duOiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9wbGF0Zm9ybS1icmVha2Rvd24ke3FzKHBhcmFtcyl9YCksCiAgICBjYW1wYWlnbkJyZWFrZG93bjogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9hbmFseXRpY3MvY2FtcGFpZ24tYnJlYWtkb3duJHtxcyhwYXJhbXMpfWApLAogICAgY29udGVudFR5cGVCcmVha2Rvd246IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL2NvbnRlbnQtdHlwZS1icmVha2Rvd24ke3FzKHBhcmFtcyl9YCksCiAgICBtZXRyaWNPcHRpb25zOiAocGxhdGZvcm0pID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL21ldHJpYy1vcHRpb25zJHtxcyh7IHBsYXRmb3JtIH0pfWApLAogICAgbWV0cmljU3VtbWFyeTogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9hbmFseXRpY3MvbWV0cmljLXN1bW1hcnkke3FzKHBhcmFtcyl9YCksCiAgICB0cmVuZDogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9hbmFseXRpY3MvdHJlbmQke3FzKHBhcmFtcyl9YCksCiAgICB0b3BQb3N0czogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9hbmFseXRpY3MvdG9wLXBvc3RzJHtxcyhwYXJhbXMpfWApLAogICAgY29tcGFyZTogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9hbmFseXRpY3MvY29tcGFyZSR7cXMocGFyYW1zKX1gKSwKICAgIG1vbnRobHk6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL21vbnRobHkke3FzKHBhcmFtcyl9YCksCiAgICBxdWFydGVybHk6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL3F1YXJ0ZXJseSR7cXMocGFyYW1zKX1gKSwKICAgIHl0ZDogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9hbmFseXRpY3MveXRkJHtxcyhwYXJhbXMpfWApLAogICAgcGxhdGZvcm1SZXBvcnQ6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL3BsYXRmb3JtLXJlcG9ydCR7cXMocGFyYW1zKX1gKSwKCiAgICBwcmV2aWV3VXBsb2FkOiAoZmlsZSkgPT4gewogICAgICBjb25zdCBmb3JtID0gbmV3IEZvcm1EYXRhKCk7CiAgICAgIGZvcm0uYXBwZW5kKCdmaWxlJywgZmlsZSk7CiAgICAgIHJldHVybiByZXF1ZXN0KCcvYXBpL3VwbG9hZHMvcHJldmlldycsIHsgbWV0aG9kOiAnUE9TVCcsIGJvZHk6IGZvcm0gfSk7CiAgICB9LAogICAgY29tbWl0VXBsb2FkOiAocGF5bG9hZCkgPT4KICAgICAgcmVxdWVzdCgnL2FwaS91cGxvYWRzL2NvbW1pdCcsIHsKICAgICAgICBtZXRob2Q6ICdQT1NUJywKICAgICAgICBoZWFkZXJzOiB7ICdDb250ZW50LVR5cGUnOiAnYXBwbGljYXRpb24vanNvbicgfSwKICAgICAgICBib2R5OiBKU09OLnN0cmluZ2lmeShwYXlsb2FkKSwKICAgICAgfSksCiAgICB1cGxvYWRIaXN0b3J5OiAoKSA9PiByZXF1ZXN0KCcvYXBpL3VwbG9hZHMvaGlzdG9yeScpLAogICAgdXBsb2FkRXJyb3JzOiAoaWQpID0+IHJlcXVlc3QoYC9hcGkvdXBsb2Fkcy8ke2lkfS9lcnJvcnNgKSwKICAgIHVwbG9hZFJhd1Jvd3M6IChpZCkgPT4gcmVxdWVzdChgL2FwaS91cGxvYWRzLyR7aWR9L3Jhdy1yb3dzYCksCgogICAgbGlzdFJlY29yZHM6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvcmVjb3JkcyR7cXMocGFyYW1zKX1gKSwKICAgIHJlY29yZHNUYWJsZTogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9yZWNvcmRzL3RhYmxlJHtxcyhwYXJhbXMpfWApLAogICAgZ2V0UmVjb3JkOiAoaWQpID0+IHJlcXVlc3QoYC9hcGkvcmVjb3Jkcy8ke2lkfWApLAogICAgdXBkYXRlUmVjb3JkOiAoaWQsIHZhbHVlcykgPT4KICAgICAgcmVxdWVzdChgL2FwaS9yZWNvcmRzLyR7aWR9YCwgewogICAgICAgIG1ldGhvZDogJ1BVVCcsCiAgICAgICAgaGVhZGVyczogeyAnQ29udGVudC1UeXBlJzogJ2FwcGxpY2F0aW9uL2pzb24nIH0sCiAgICAgICAgYm9keTogSlNPTi5zdHJpbmdpZnkoeyB2YWx1ZXMgfSksCiAgICAgIH0pLAogICAgZGVsZXRlUmVjb3JkUG9zdDogKHBvc3RJZCkgPT4gcmVxdWVzdChgL2FwaS9yZWNvcmRzL3Bvc3QvJHtwb3N0SWR9YCwgeyBtZXRob2Q6ICdERUxFVEUnIH0pLAogICAgZGVsZXRlUmVjb3JkUGxhdGZvcm06IChwb3N0SWQsIHBsYXRmb3JtKSA9PgogICAgICByZXF1ZXN0KGAvYXBpL3JlY29yZHMvcG9zdC8ke3Bvc3RJZH0vcGxhdGZvcm0vJHtwbGF0Zm9ybX1gLCB7IG1ldGhvZDogJ0RFTEVURScgfSksCiAgICBkdXBsaWNhdGVSZWNvcmRzUHJldmlldzogKCkgPT4gcmVxdWVzdCgnL2FwaS9yZWNvcmRzL2R1cGxpY2F0ZXMnKSwKICAgIHJlc29sdmVEdXBsaWNhdGVSZWNvcmRzOiAoKSA9PiByZXF1ZXN0KCcvYXBpL3JlY29yZHMvZHVwbGljYXRlcy9yZXNvbHZlJywgeyBtZXRob2Q6ICdQT1NUJyB9KSwKICAgIHdpcGVVcGxvYWRlZFJlY29yZHM6ICgpID0+CiAgICAgIHJlcXVlc3QoJy9hcGkvcmVjb3Jkcy93aXBlJywgewogICAgICAgIG1ldGhvZDogJ1BPU1QnLAogICAgICAgIGhlYWRlcnM6IHsgJ0NvbnRlbnQtVHlwZSc6ICdhcHBsaWNhdGlvbi9qc29uJyB9LAogICAgICAgIGJvZHk6IEpTT04uc3RyaW5naWZ5KHsgY29uZmlybTogJ0RFTEVURScgfSksCiAgICAgIH0pLAoKICAgIHJlc3RvcmVCYWNrdXA6IChmb3JtKSA9PiByZXF1ZXN0KCcvYXBpL2JhY2t1cC9yZXN0b3JlJywgeyBtZXRob2Q6ICdQT1NUJywgYm9keTogZm9ybSB9KSwKCiAgICBsaXN0Rm9sbG93ZXJzOiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2ZvbGxvd2VycyR7cXMocGFyYW1zKX1gKSwKICAgIGZvbGxvd2Vyc0dyb3d0aDogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9mb2xsb3dlcnMvZ3Jvd3RoJHtxcyhwYXJhbXMpfWApLAogICAgZm9sbG93ZXJzS3BpczogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9mb2xsb3dlcnMva3BpcyR7cXMocGFyYW1zKX1gKSwKICAgIHNhdmVGb2xsb3dlcnM6IChwYXlsb2FkKSA9PgogICAgICByZXF1ZXN0KCcvYXBpL2ZvbGxvd2VycycsIHsgbWV0aG9kOiAnUE9TVCcsIGhlYWRlcnM6IHsgJ0NvbnRlbnQtVHlwZSc6ICdhcHBsaWNhdGlvbi9qc29uJyB9LCBib2R5OiBKU09OLnN0cmluZ2lmeShwYXlsb2FkKSB9KSwKICAgIHVwZGF0ZUZvbGxvd2VyczogKGlkLCBwYXlsb2FkKSA9PgogICAgICByZXF1ZXN0KGAvYXBpL2ZvbGxvd2Vycy8ke2lkfWAsIHsgbWV0aG9kOiAnUFVUJywgaGVhZGVyczogeyAnQ29udGVudC1UeXBlJzogJ2FwcGxpY2F0aW9uL2pzb24nIH0sIGJvZHk6IEpTT04uc3RyaW5naWZ5KHBheWxvYWQpIH0pLAogICAgZGVsZXRlRm9sbG93ZXJzOiAoaWQpID0+IHJlcXVlc3QoYC9hcGkvZm9sbG93ZXJzLyR7aWR9YCwgeyBtZXRob2Q6ICdERUxFVEUnIH0pLAogIH07Cn0pKCk7CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgU3RhdGUgLyBGb3JtYXQgLyBUb2FzdCDigJQgc2hhcmVkIGFwcCBzdGF0ZSArIHNtYWxsIHV0aWxpdGllcy4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmNvbnN0IFN0YXRlID0gKCgpID0+IHsKICBjb25zdCB0b2RheSA9IG5ldyBEYXRlKCk7CiAgY29uc3QgaXNvID0gKGQpID0+IGQudG9JU09TdHJpbmcoKS5zbGljZSgwLCAxMCk7CiAgY29uc3QgdGhpcnR5RGF5c0FnbyA9IG5ldyBEYXRlKHRvZGF5KTsKICB0aGlydHlEYXlzQWdvLnNldERhdGUodGhpcnR5RGF5c0Fnby5nZXREYXRlKCkgLSAyOSk7CgogIGNvbnN0IGZpbHRlcnMgPSB7CiAgICBkYXRlRnJvbTogaXNvKHRoaXJ0eURheXNBZ28pLAogICAgZGF0ZVRvOiBpc28odG9kYXkpLAogICAgcGxhdGZvcm06ICdhbGwnLAogICAgY2FtcGFpZ25UeXBlOiAnYWxsJywKICAgIGNvbnRlbnRUeXBlOiAnYWxsJywKICB9OwoKICBjb25zdCBsaXN0ZW5lcnMgPSBbXTsKCiAgcmV0dXJuIHsKICAgIGdldEZpbHRlcnM6ICgpID0+ICh7IC4uLmZpbHRlcnMgfSksCiAgICBzZXRGaWx0ZXJzKHBhcnRpYWwpIHsKICAgICAgT2JqZWN0LmFzc2lnbihmaWx0ZXJzLCBwYXJ0aWFsKTsKICAgICAgbGlzdGVuZXJzLmZvckVhY2goKGZuKSA9PiBmbih0aGlzLmdldEZpbHRlcnMoKSkpOwogICAgfSwKICAgIG9uQ2hhbmdlKGZuKSB7CiAgICAgIGxpc3RlbmVycy5wdXNoKGZuKTsKICAgIH0sCiAgfTsKfSkoKTsKCmNvbnN0IEZvcm1hdCA9IHsKICBudW1iZXIobikgewogICAgaWYgKG4gPT09IG51bGwgfHwgbiA9PT0gdW5kZWZpbmVkKSByZXR1cm4gJ+KAlCc7CiAgICByZXR1cm4gTWF0aC5yb3VuZChuKS50b0xvY2FsZVN0cmluZygnZW4tVVMnKTsKICB9LAogIGNvbXBhY3QobikgewogICAgaWYgKG4gPT09IG51bGwgfHwgbiA9PT0gdW5kZWZpbmVkKSByZXR1cm4gJ+KAlCc7CiAgICBjb25zdCBhYnMgPSBNYXRoLmFicyhuKTsKICAgIGlmIChhYnMgPj0gMV8wMDBfMDAwKSByZXR1cm4gYCR7KG4gLyAxXzAwMF8wMDApLnRvRml4ZWQoMSkucmVwbGFjZSgvXC4wJC8sICcnKX1NYDsKICAgIGlmIChhYnMgPj0gMV8wMDApIHJldHVybiBgJHsobiAvIDFfMDAwKS50b0ZpeGVkKDEpLnJlcGxhY2UoL1wuMCQvLCAnJyl9S2A7CiAgICByZXR1cm4gYCR7TWF0aC5yb3VuZChuKX1gOwogIH0sCiAgLyoqIERhc2hib2FyZC13aWRlICJwcm9mZXNzaW9uYWwiIG51bWJlciBmb3JtYXQ6IHBsYWluIHVuZGVyIDEsMDAwOyBjb21tYS1ncm91cGVkCiAgICAgIHVwIHRvIDEwLDAwMDsgYWJicmV2aWF0ZWQgKEsvTSkgYmV5b25kIHRoYXQg4oCUIGUuZy4gODUwLCAxLDI1MCwgMTIuNUssIDE1NkssIDEuMjVNLiAqLwogIHNtYXJ0KG4pIHsKICAgIGlmIChuID09PSBudWxsIHx8IG4gPT09IHVuZGVmaW5lZCkgcmV0dXJuICfigJQnOwogICAgY29uc3QgYWJzID0gTWF0aC5hYnMobik7CiAgICBpZiAoYWJzIDwgMTAwMCkgcmV0dXJuIGAke01hdGgucm91bmQobil9YDsKICAgIGlmIChhYnMgPCAxMDAwMCkgcmV0dXJuIE1hdGgucm91bmQobikudG9Mb2NhbGVTdHJpbmcoJ2VuLVVTJyk7CiAgICBpZiAoYWJzIDwgMV8wMDBfMDAwKSByZXR1cm4gYCR7KG4gLyAxMDAwKS50b0ZpeGVkKDEpLnJlcGxhY2UoL1wuMCQvLCAnJyl9S2A7CiAgICByZXR1cm4gYCR7KG4gLyAxXzAwMF8wMDApLnRvRml4ZWQoMikucmVwbGFjZSgvXC4/MCskLywgJycpfU1gOwogIH0sCiAgcGVyY2VudChuKSB7CiAgICBpZiAobiA9PT0gbnVsbCB8fCBuID09PSB1bmRlZmluZWQpIHJldHVybiAn4oCUJzsKICAgIHJldHVybiBgJHtOdW1iZXIobikudG9GaXhlZCgxKS5yZXBsYWNlKC9cLjAkLywgJycpfSVgOwogIH0sCiAgcGN0KG4pIHsKICAgIGlmIChuID09PSBudWxsIHx8IG4gPT09IHVuZGVmaW5lZCkgcmV0dXJuICfigJQnOwogICAgY29uc3Qgc2lnbiA9IG4gPiAwID8gJysnIDogJyc7CiAgICByZXR1cm4gYCR7c2lnbn0ke24udG9GaXhlZCgxKX0lYDsKICB9LAogIGRhdGUoaXNvXykgewogICAgaWYgKCFpc29fKSByZXR1cm4gJ+KAlCc7CiAgICBjb25zdCBbeSwgbSwgZF0gPSBpc29fLnNwbGl0KCctJykubWFwKE51bWJlcik7CiAgICByZXR1cm4gbmV3IERhdGUoeSwgbSAtIDEsIGQpLnRvTG9jYWxlRGF0ZVN0cmluZygnZW4tVVMnLCB7IG1vbnRoOiAnc2hvcnQnLCBkYXk6ICdudW1lcmljJywgeWVhcjogJ251bWVyaWMnIH0pOwogIH0sCiAgZHVyYXRpb24oc2Vjb25kcykgewogICAgaWYgKHNlY29uZHMgPT09IG51bGwgfHwgc2Vjb25kcyA9PT0gdW5kZWZpbmVkKSByZXR1cm4gJ+KAlCc7CiAgICBjb25zdCBzID0gTWF0aC5yb3VuZChzZWNvbmRzKTsKICAgIGlmIChzIDwgNjApIHJldHVybiBgJHtzfXNgOwogICAgaWYgKHMgPCAzNjAwKSByZXR1cm4gYCR7TWF0aC5mbG9vcihzIC8gNjApfW0gJHtzICUgNjB9c2A7CiAgICBjb25zdCBoID0gTWF0aC5mbG9vcihzIC8gMzYwMCk7CiAgICBjb25zdCBtID0gTWF0aC5yb3VuZCgocyAlIDM2MDApIC8gNjApOwogICAgcmV0dXJuIGAke2h9aCAke219bWA7CiAgfSwKICBkZWx0YUNsYXNzKG4pIHsKICAgIGlmIChuID09PSBudWxsIHx8IG4gPT09IHVuZGVmaW5lZCkgcmV0dXJuICdmbGF0JzsKICAgIGlmIChuID4gMC41KSByZXR1cm4gJ3VwJzsKICAgIGlmIChuIDwgLTAuNSkgcmV0dXJuICdkb3duJzsKICAgIHJldHVybiAnZmxhdCc7CiAgfSwKfTsKCmNvbnN0IFRvYXN0ID0gewogIHNob3cobWVzc2FnZSwgdHlwZSA9ICdzdWNjZXNzJykgewogICAgY29uc3Qgcm9vdCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd0b2FzdFJvb3QnKTsKICAgIGNvbnN0IGVsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBlbC5jbGFzc05hbWUgPSBgdG9hc3QgJHt0eXBlfWA7CiAgICBlbC50ZXh0Q29udGVudCA9IG1lc3NhZ2U7CiAgICByb290LmFwcGVuZENoaWxkKGVsKTsKICAgIHNldFRpbWVvdXQoKCkgPT4gZWwucmVtb3ZlKCksIDUwMDApOwogIH0sCn07CgovKiogU2FmZWx5IGJ1aWxkcyBET00gdGV4dCBub2RlcyBmb3IgdW50cnVzdGVkIHN0cmluZ3MgKGNhcHRpb25zLCBmaWxlbmFtZXMsIHBsYXRmb3JtIGxhYmVscyBmcm9tIGRhdGEpLiAqLwpmdW5jdGlvbiB0ZXh0RWwodGFnLCB0ZXh0LCBjbGFzc05hbWUpIHsKICBjb25zdCBlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQodGFnKTsKICBpZiAoY2xhc3NOYW1lKSBlbC5jbGFzc05hbWUgPSBjbGFzc05hbWU7CiAgZWwuYXBwZW5kQ2hpbGQoZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUodGV4dCA/PyAnJykpOwogIHJldHVybiBlbDsKfQoKLyoqIEEgcHJlbWl1bSBlbXB0eSBzdGF0ZTogaWNvbiArIGV4cGxhbmF0aW9uICsgb3B0aW9uYWwgYWN0aW9uLCBpbnN0ZWFkIG9mIGEgYmxhbmsgYXJlYS4KICAgIEljb25zIHJlbmRlciB2aWEgdGhlIHBhZ2Utd2lkZSBNdXRhdGlvbk9ic2VydmVyIHRoYXQgY2FsbHMgbHVjaWRlLmNyZWF0ZUljb25zKCkgKHNlZSBib290c3RyYXApLiAqLwpmdW5jdGlvbiBlbXB0eVN0YXRlKHsgaWNvbiA9ICdpbmJveCcsIHRpdGxlLCBtZXNzYWdlLCBhY3Rpb25MYWJlbCwgb25BY3Rpb24gfSkgewogIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICB3cmFwLmNsYXNzTmFtZSA9ICdlbXB0eS1zdGF0ZSc7CiAgY29uc3QgaWNvbldyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICBpY29uV3JhcC5jbGFzc05hbWUgPSAnZW1wdHktaWNvbic7CiAgaWNvbldyYXAuaW5uZXJIVE1MID0gYDxpIGRhdGEtbHVjaWRlPSIke2ljb259IiBzdHlsZT0id2lkdGg6MjJweDtoZWlnaHQ6MjJweDsiPjwvaT5gOwogIHdyYXAuYXBwZW5kQ2hpbGQoaWNvbldyYXApOwogIGlmICh0aXRsZSkgd3JhcC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIHRpdGxlLCAnZW1wdHktdGl0bGUnKSk7CiAgaWYgKG1lc3NhZ2UpIHdyYXAuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBtZXNzYWdlLCAnZW1wdHktbWVzc2FnZScpKTsKICBpZiAoYWN0aW9uTGFiZWwgJiYgb25BY3Rpb24pIHsKICAgIGNvbnN0IGJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgYnRuLmNsYXNzTmFtZSA9ICdidG4gcHJpbWFyeSc7CiAgICBidG4udGV4dENvbnRlbnQgPSBhY3Rpb25MYWJlbDsKICAgIGJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIG9uQWN0aW9uKTsKICAgIHdyYXAuYXBwZW5kQ2hpbGQoYnRuKTsKICB9CiAgcmV0dXJuIHdyYXA7Cn0KCi8qKiBBIDxidXR0b24+IHdpdGggYSBzbWFsbCBsZWFkaW5nIEx1Y2lkZSBpY29uIGJlZm9yZSBpdHMgbGFiZWwgKGxhYmVsIGlzIGFsd2F5cyBhIHN0YXRpYywgZGV2ZWxvcGVyLXN1cHBsaWVkIHN0cmluZyBhdCBjYWxsIHNpdGVzLCBuZXZlciB1c2VyIGRhdGEg4oCUIGluc2VydGVkIHZpYSBjcmVhdGVUZXh0Tm9kZSByZWdhcmRsZXNzKS4gKi8KZnVuY3Rpb24gaWNvbkJ0bihjbGFzc05hbWUsIGljb25OYW1lLCBsYWJlbCkgewogIGNvbnN0IGJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogIGJ0bi5jbGFzc05hbWUgPSBjbGFzc05hbWU7CiAgY29uc3QgaWNvbiA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2knKTsKICBpY29uLnNldEF0dHJpYnV0ZSgnZGF0YS1sdWNpZGUnLCBpY29uTmFtZSk7CiAgaWNvbi5zdHlsZS53aWR0aCA9ICcxM3B4JzsKICBpY29uLnN0eWxlLmhlaWdodCA9ICcxM3B4JzsKICBidG4uYXBwZW5kQ2hpbGQoaWNvbik7CiAgYnRuLmFwcGVuZENoaWxkKGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKGAgJHtsYWJlbH1gKSk7CiAgcmV0dXJuIGJ0bjsKfQoKLyoqIFNoaW1tZXJpbmcgcGxhY2Vob2xkZXJzIHNob3duIHRoZSBpbnN0YW50IGEgc2VjdGlvbiBzdGFydHMgbG9hZGluZywgc3dhcHBlZCBmb3IgcmVhbAogICAgY29udGVudCAob3IgYW4gZW1wdHkgc3RhdGUpIG9uY2UgdGhlIGZldGNoIHJlc29sdmVzIOKAlCBubyBibGFuayBhcmVhcyB3aGlsZSB3YWl0aW5nLiAqLwpmdW5jdGlvbiBza2VsZXRvblN0YXRHcmlkKGNvdW50ID0gNikgewogIGNvbnN0IGdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICBncmlkLmNsYXNzTmFtZSA9ICdza2VsZXRvbi1zdGF0LWdyaWQnOwogIGZvciAobGV0IGkgPSAwOyBpIDwgY291bnQ7IGkgKz0gMSkgewogICAgY29uc3QgdGlsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGlsZS5jbGFzc05hbWUgPSAnc2tlbGV0b24gc2tlbGV0b24tdGlsZSc7CiAgICBncmlkLmFwcGVuZENoaWxkKHRpbGUpOwogIH0KICByZXR1cm4gZ3JpZDsKfQpmdW5jdGlvbiBza2VsZXRvbkNoYXJ0KCkgewogIGNvbnN0IGRpdiA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogIGRpdi5jbGFzc05hbWUgPSAnc2tlbGV0b24gc2tlbGV0b24tY2hhcnQnOwogIHJldHVybiBkaXY7Cn0KZnVuY3Rpb24gc2tlbGV0b25Sb3dzKGNvdW50ID0gNikgewogIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICBmb3IgKGxldCBpID0gMDsgaSA8IGNvdW50OyBpICs9IDEpIHsKICAgIGNvbnN0IHJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcm93LmNsYXNzTmFtZSA9ICdza2VsZXRvbiBza2VsZXRvbi1yb3cnOwogICAgd3JhcC5hcHBlbmRDaGlsZChyb3cpOwogIH0KICByZXR1cm4gd3JhcDsKfQoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIFNoYXJlZCBhbmltYXRpb24gcHJpbWl0aXZlcyDigJQgYSBjb3VudC11cCBmb3IgS1BJIG51bWJlcnMgYW5kIGEKICAgQ1NTIHdpZHRoLXRyYW5zaXRpb24gYmFyLCBib3RoIHJldXNlZCBhY3Jvc3MgdGhlIERhc2hib2FyZCBhbmQKICAgQ29tcGFyaXNvbnMgcGFnZXMuIEJvdGggcmVzcGVjdCBwcmVmZXJzLXJlZHVjZWQtbW90aW9uIChndWFyZGVkCiAgIGluIENTUywgc2VlIC5iYXItZmlsbCAvIHRoZSBhbmltYXRlQ291bnQgZHVyYXRpb24gY2hlY2sgYmVsb3cpLgogICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KY29uc3QgUFJFRkVSU19SRURVQ0VEX01PVElPTiA9IHdpbmRvdy5tYXRjaE1lZGlhICYmIHdpbmRvdy5tYXRjaE1lZGlhKCcocHJlZmVycy1yZWR1Y2VkLW1vdGlvbjogcmVkdWNlKScpLm1hdGNoZXM7CgovKiogU2hyaW5rcyBgZWxgJ3MgZm9udCBzaXplIGp1c3QgZW5vdWdoIGZvciBpdHMgY3VycmVudCB0ZXh0IHRvIGZpdCBpdHMgb3duIHdpZHRoIOKAlCBhIEtQSSB0aWxlJ3MgYm94IGlzIGEgZml4ZWQgc2l6ZSwgYnV0IHRoZSB2YWx1ZSBpbnNpZGUgaXQgaXNuJ3QgKGEgZm9sbG93ZXIgY291bnQgY2FuIGJlICIwIiBvciAiMSwwNDgsNTc2IiksIHNvIGEgc2luZ2xlIGZpeGVkIGZvbnQtc2l6ZSB3aWxsIGV2ZW50dWFsbHkgb3ZlcmZsb3cuIFJlc2V0cyB0byB0aGUgQ1NTLWRlZmluZWQgc2l6ZSBmaXJzdCwgdGhlbiBzdGVwcyBkb3duIGJ5IDFweCBhdCBhIHRpbWUgdW50aWwgaXQgZml0cyBvciBoaXRzIGBtaW5TaXplYC4gKi8KZnVuY3Rpb24gZml0U3RhdFZhbHVlKGVsLCBtaW5TaXplID0gMTgpIHsKICBpZiAoIWVsKSByZXR1cm47CiAgZWwuc3R5bGUuZm9udFNpemUgPSAnJzsKICBjb25zdCBtYXhXaWR0aCA9IGVsLmNsaWVudFdpZHRoOwogIGlmICghbWF4V2lkdGgpIHJldHVybjsKICBsZXQgc2l6ZSA9IHBhcnNlRmxvYXQoZ2V0Q29tcHV0ZWRTdHlsZShlbCkuZm9udFNpemUpOwogIHdoaWxlIChlbC5zY3JvbGxXaWR0aCA+IG1heFdpZHRoICYmIHNpemUgPiBtaW5TaXplKSB7CiAgICBzaXplIC09IDE7CiAgICBlbC5zdHlsZS5mb250U2l6ZSA9IGAke3NpemV9cHhgOwogIH0KfQoKLyoqIEFuaW1hdGVzIGEgbnVtYmVyIGZyb20gYGZyb21gIHRvIGB0b2AgaW5zaWRlIGBlbGAgb3ZlciBgZHVyYXRpb25gbXMsIGZvcm1hdHRpbmcgZWFjaCBmcmFtZSB3aXRoIGBmb3JtYXRgIChkZWZhdWx0cyB0byBhIHBsYWluIHJvdW5kZWQgaW50ZWdlcikuIFNraXBzIHN0cmFpZ2h0IHRvIHRoZSBmaW5hbCB2YWx1ZSB1bmRlciBwcmVmZXJzLXJlZHVjZWQtbW90aW9uLiBTaHJpbmtzIHRoZSBmb250IHRvIGZpdCBvbmNlIHRoZSBmaW5hbCB2YWx1ZSBsYW5kcywgc2luY2UgdGhlIGFuaW1hdGVkIGRpZ2l0cyBjYW4gYmUgYSBkaWZmZXJlbnQgd2lkdGggdGhhbiB0aGUgc2V0dGxlZCB2YWx1ZSDigJQgZGVmZXJyZWQgYSBmcmFtZSBiZWNhdXNlIGBlbGAgaXMgdHlwaWNhbGx5IHN0aWxsIGRldGFjaGVkIGZyb20gdGhlIGRvY3VtZW50IChtaWQtY29uc3RydWN0aW9uIGJ5IGl0cyBjYWxsZXIpIHdoZW4gYW5pbWF0ZUNvdW50IGlzIGZpcnN0IGludm9rZWQsIGFuZCBjbGllbnRXaWR0aCByZWFkcyAwIHVudGlsIGl0J3MgYWN0dWFsbHkgYXR0YWNoZWQgYW5kIGxhaWQgb3V0LiAqLwpmdW5jdGlvbiBhbmltYXRlQ291bnQoZWwsIGZyb20sIHRvLCBkdXJhdGlvbiA9IDkwMCwgZm9ybWF0KSB7CiAgaWYgKCFlbCkgcmV0dXJuOwogIGNvbnN0IGZtdCA9IGZvcm1hdCB8fCAoKHYpID0+IE1hdGgucm91bmQodikudG9Mb2NhbGVTdHJpbmcoJ2VuLVVTJykpOwogIGlmIChQUkVGRVJTX1JFRFVDRURfTU9USU9OIHx8IGZyb20gPT09IHRvIHx8ICFOdW1iZXIuaXNGaW5pdGUoZnJvbSkgfHwgIU51bWJlci5pc0Zpbml0ZSh0bykpIHsKICAgIGVsLnRleHRDb250ZW50ID0gZm10KHRvKTsKICAgIHJlcXVlc3RBbmltYXRpb25GcmFtZSgoKSA9PiBmaXRTdGF0VmFsdWUoZWwpKTsKICAgIHJldHVybjsKICB9CiAgY29uc3Qgc3RhcnQgPSBwZXJmb3JtYW5jZS5ub3coKTsKICBmdW5jdGlvbiB0aWNrKG5vdykgewogICAgY29uc3QgZWxhcHNlZCA9IG5vdyAtIHN0YXJ0OwogICAgY29uc3QgcHJvZ3Jlc3MgPSBNYXRoLm1pbigxLCBlbGFwc2VkIC8gZHVyYXRpb24pOwogICAgY29uc3QgZWFzZWQgPSAxIC0gTWF0aC5wb3coMSAtIHByb2dyZXNzLCAzKTsgLy8gZWFzZU91dEN1YmljCiAgICBlbC50ZXh0Q29udGVudCA9IGZtdChmcm9tICsgKHRvIC0gZnJvbSkgKiBlYXNlZCk7CiAgICBpZiAocHJvZ3Jlc3MgPCAxKSB7CiAgICAgIHJlcXVlc3RBbmltYXRpb25GcmFtZSh0aWNrKTsKICAgIH0gZWxzZSB7CiAgICAgIGZpdFN0YXRWYWx1ZShlbCk7CiAgICB9CiAgfQogIHJlcXVlc3RBbmltYXRpb25GcmFtZSh0aWNrKTsKfQoKLyoqIEEgbGFiZWxlZCBob3Jpem9udGFsIGJhciB0aGF0IGFuaW1hdGVzIGl0cyB3aWR0aCBpbiBvbiBpbnNlcnRpb24g4oCUIHVzZWQgZm9yIHRoZSBDb21wYXJpc29ucyBwYWdlJ3MgcGFpcmVkIFJhbmdlIEEvQiBiYXJzLiBgdmFsdWVgL2BtYXhgIGRyaXZlIHRoZSBmaWxsIHBlcmNlbnRhZ2U7IGBjb2xvclZhcmAgaXMgYSBDU1MgY3VzdG9tIHByb3BlcnR5IG5hbWUgKGUuZy4gJy0tc2VyaWVzLTEnKS4gKi8KZnVuY3Rpb24gYnVpbGRCYXIoeyBsYWJlbCwgdmFsdWUsIG1heCwgY29sb3JWYXIsIGZvcm1hdFZhbHVlIH0pIHsKICBjb25zdCByb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICByb3cuY2xhc3NOYW1lID0gJ2Jhci1yb3cnOwogIGNvbnN0IGxhYmVsRWwgPSB0ZXh0RWwoJ2RpdicsIGxhYmVsLCAnYmFyLWxhYmVsJyk7CiAgY29uc3QgdHJhY2sgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICB0cmFjay5jbGFzc05hbWUgPSAnYmFyLXRyYWNrJzsKICBjb25zdCBmaWxsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgZmlsbC5jbGFzc05hbWUgPSAnYmFyLWZpbGwnOwogIGZpbGwuc3R5bGUuYmFja2dyb3VuZCA9IGNvbG9yVmFyID8gYHZhcigke2NvbG9yVmFyfSlgIDogJ3ZhcigtLXNlcmllcy0xKSc7CiAgdHJhY2suYXBwZW5kQ2hpbGQoZmlsbCk7CiAgY29uc3QgdmFsdWVFbCA9IHRleHRFbCgnZGl2JywgZm9ybWF0VmFsdWUgPyBmb3JtYXRWYWx1ZSh2YWx1ZSkgOiBTdHJpbmcodmFsdWUpLCAnYmFyLXZhbHVlJyk7CiAgcm93LmFwcGVuZChsYWJlbEVsLCB0cmFjaywgdmFsdWVFbCk7CiAgY29uc3QgcGN0ID0gbWF4ID4gMCA/IE1hdGgubWluKDEwMCwgTWF0aC5yb3VuZCgodmFsdWUgLyBtYXgpICogMTAwMCkgLyAxMCkgOiAwOwogIHJlcXVlc3RBbmltYXRpb25GcmFtZSgoKSA9PiB7IGZpbGwuc3R5bGUud2lkdGggPSBgJHtwY3R9JWA7IH0pOwogIHJldHVybiByb3c7Cn0KCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBTaGFyZWQgdGFibGUgdG9vbGJhciBwaWVjZXMg4oCUIHNlYXJjaCBib3gsIGNsaWVudC1zaWRlIHBhZ2VyLAogICBhbmQgQ1NWL1hMU1ggZXhwb3J0IOKAlCByZXVzZWQgYnkgdGhlIEZvbGxvd2VycyBEYXRhIGFuZCBVcGxvYWQKICAgSGlzdG9yeSB0YWJzIChib3RoIGxvYWQgdGhlaXIgZnVsbCBkYXRhc2V0IG9uY2UgYW5kIHNlYXJjaC8KICAgc29ydC9wYWdpbmF0ZSBpdCBpbiB0aGUgYnJvd3NlciwgdW5saWtlIERhdGEgUmVjb3JkcyB3aGljaCBpcwogICBzZXJ2ZXItcGFnaW5hdGVkKS4gQ1NWIG5lZWRzIG5vIHNlcnZlciByb3VuZCB0cmlwIGF0IGFsbDsgWExTWAogICBnb2VzIHRocm91Z2ggUE9TVCAvYXBpL2V4cG9ydCBzbyBleGNlbGpzIChhbHJlYWR5IGEgZGVwZW5kZW5jeSkKICAgY2FuIGdlbmVyYXRlIGEgcmVhbCAueGxzeCB3aXRob3V0IGFkZGluZyBhIGNsaWVudC1zaWRlIGxpYnJhcnkuCiAgID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwpmdW5jdGlvbiBidWlsZFNlYXJjaEJveCh7IHBsYWNlaG9sZGVyLCB2YWx1ZSwgb25DaGFuZ2UgfSkgewogIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICB3cmFwLmNsYXNzTmFtZSA9ICdyZWNvcmRzLXNlYXJjaCc7CiAgY29uc3QgaW5wdXQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOwogIGlucHV0LnR5cGUgPSAnc2VhcmNoJzsKICBpbnB1dC5wbGFjZWhvbGRlciA9IHBsYWNlaG9sZGVyOwogIGlucHV0LnZhbHVlID0gdmFsdWUgfHwgJyc7CiAgbGV0IGRlYm91bmNlID0gbnVsbDsKICBpbnB1dC5hZGRFdmVudExpc3RlbmVyKCdpbnB1dCcsICgpID0+IHsKICAgIGNsZWFyVGltZW91dChkZWJvdW5jZSk7CiAgICBkZWJvdW5jZSA9IHNldFRpbWVvdXQoKCkgPT4gb25DaGFuZ2UoaW5wdXQudmFsdWUpLCAzMDApOwogIH0pOwogIHdyYXAuYXBwZW5kQ2hpbGQoaW5wdXQpOwogIHJldHVybiB3cmFwOwp9CgovKiogU2xpY2VzIGFuIGFscmVhZHktbG9hZGVkLCBhbHJlYWR5LWZpbHRlcmVkL3NvcnRlZCBhcnJheSBmb3IgY2xpZW50LXNpZGUgcGFnaW5hdGlvbiDigJQgdGhlIGNvdW50ZXJwYXJ0IHRvIHRoZSBzZXJ2ZXItc2lkZSBwYWdpbmF0ZSgpIGluIGFwcC5qcywgZm9yIHRhYmxlcyB0aGF0IGRvbid0IGhhdmUgYSBwYWdpbmF0ZWQgZW5kcG9pbnQuICovCmZ1bmN0aW9uIHBhZ2luYXRlQ2xpZW50U2lkZShyb3dzLCBwYWdlLCBwYWdlU2l6ZSkgewogIGNvbnN0IHRvdGFsUGFnZXMgPSBNYXRoLm1heCgxLCBNYXRoLmNlaWwocm93cy5sZW5ndGggLyBwYWdlU2l6ZSkpOwogIGNvbnN0IHNhZmVQYWdlID0gTWF0aC5taW4oTWF0aC5tYXgoMSwgcGFnZSksIHRvdGFsUGFnZXMpOwogIGNvbnN0IHN0YXJ0ID0gKHNhZmVQYWdlIC0gMSkgKiBwYWdlU2l6ZTsKICByZXR1cm4geyBwYWdlUm93czogcm93cy5zbGljZShzdGFydCwgc3RhcnQgKyBwYWdlU2l6ZSksIHRvdGFsUGFnZXMsIHNhZmVQYWdlLCB0b3RhbDogcm93cy5sZW5ndGggfTsKfQoKZnVuY3Rpb24gYnVpbGRQYWdlcih7IHBhZ2UsIHRvdGFsUGFnZXMsIHRvdGFsLCBvblByZXYsIG9uTmV4dCB9KSB7CiAgY29uc3QgcGFnZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICBwYWdlci5jbGFzc05hbWUgPSAncGFnaW5hdGlvbi1yb3cnOwogIGNvbnN0IHByZXZCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICBwcmV2QnRuLmNsYXNzTmFtZSA9ICdidG4nOyBwcmV2QnRuLnR5cGUgPSAnYnV0dG9uJzsgcHJldkJ0bi50ZXh0Q29udGVudCA9ICdQcmV2aW91cyc7CiAgcHJldkJ0bi5kaXNhYmxlZCA9IHBhZ2UgPD0gMTsKICBwcmV2QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgb25QcmV2KTsKICBjb25zdCBuZXh0QnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgbmV4dEJ0bi5jbGFzc05hbWUgPSAnYnRuJzsgbmV4dEJ0bi50eXBlID0gJ2J1dHRvbic7IG5leHRCdG4udGV4dENvbnRlbnQgPSAnTmV4dCc7CiAgbmV4dEJ0bi5kaXNhYmxlZCA9IHBhZ2UgPj0gdG90YWxQYWdlczsKICBuZXh0QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgb25OZXh0KTsKICBwYWdlci5hcHBlbmQocHJldkJ0biwgdGV4dEVsKCdzcGFuJywgYFBhZ2UgJHtwYWdlfSBvZiAke3RvdGFsUGFnZXN9IOKAlCAke3RvdGFsfSByZWNvcmQocylgKSwgbmV4dEJ0bik7CiAgcmV0dXJuIHBhZ2VyOwp9CgpmdW5jdGlvbiBkb3dubG9hZEJsb2IoZmlsZW5hbWUsIG1pbWVUeXBlLCBjb250ZW50KSB7CiAgY29uc3QgYmxvYiA9IG5ldyBCbG9iKFtjb250ZW50XSwgeyB0eXBlOiBtaW1lVHlwZSB9KTsKICBjb25zdCB1cmwgPSBVUkwuY3JlYXRlT2JqZWN0VVJMKGJsb2IpOwogIGNvbnN0IGEgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdhJyk7CiAgYS5ocmVmID0gdXJsOwogIGEuZG93bmxvYWQgPSBmaWxlbmFtZTsKICBkb2N1bWVudC5ib2R5LmFwcGVuZENoaWxkKGEpOwogIGEuY2xpY2soKTsKICBhLnJlbW92ZSgpOwogIFVSTC5yZXZva2VPYmplY3RVUkwodXJsKTsKfQoKLyoqIFF1b3RlZC1jb21tYS1qb2luIENTViDigJQgdGhlIGNsaWVudC1zaWRlIG1pcnJvciBvZiBhcHAuanMncyB0b0NTVigpLCBmb3IgdGFibGVzIHdob3NlIGZ1bGwgZGF0YXNldCBpcyBhbHJlYWR5IGxvYWRlZCBpbiB0aGUgYnJvd3Nlci4gKi8KZnVuY3Rpb24gdG9DU1ZDbGllbnRTaWRlKHJvd3MsIGNvbHVtbnMpIHsKICBjb25zdCBlc2NhcGUgPSAodikgPT4gewogICAgY29uc3QgcyA9IHYgPT09IG51bGwgfHwgdiA9PT0gdW5kZWZpbmVkID8gJycgOiBTdHJpbmcodik7CiAgICByZXR1cm4gL1siLFxyXG5dLy50ZXN0KHMpID8gYCIke3MucmVwbGFjZSgvIi9nLCAnIiInKX0iYCA6IHM7CiAgfTsKICBjb25zdCBsaW5lcyA9IFtjb2x1bW5zLm1hcCgoYykgPT4gZXNjYXBlKGMubGFiZWwpKS5qb2luKCcsJyldOwogIHJvd3MuZm9yRWFjaCgocm93KSA9PiBsaW5lcy5wdXNoKGNvbHVtbnMubWFwKChjKSA9PiBlc2NhcGUocm93W2Mua2V5XSkpLmpvaW4oJywnKSkpOwogIHJldHVybiBsaW5lcy5qb2luKCdcclxuJyk7Cn0KCi8qKiBBIHNtYWxsICJFeHBvcnQgQ1NWIC8gRXhwb3J0IEV4Y2VsIiBidXR0b24gcGFpci4gYGdldFJvd3NBbmRDb2x1bW5zKClgIGlzIGNhbGxlZCBhdCBjbGljayB0aW1lIHNvIGl0IGFsd2F5cyBleHBvcnRzIHdoYXRldmVyJ3MgY3VycmVudGx5IGZpbHRlcmVkL3NvcnRlZCwgbmV2ZXIgYSBzdGFsZSBzbmFwc2hvdC4gKi8KZnVuY3Rpb24gYnVpbGRFeHBvcnRCdXR0b25zKHsgZ2V0Um93c0FuZENvbHVtbnMsIGZpbGVuYW1lQmFzZSwgc2hlZXROYW1lIH0pIHsKICBjb25zdCB3cmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgd3JhcC5jbGFzc05hbWUgPSAnZXhwb3J0LWJ1dHRvbnMnOwogIGNvbnN0IGNzdkJ0biA9IGljb25CdG4oJ2J0bicsICdmaWxlLWRvd24nLCAnRXhwb3J0IENTVicpOwogIGNzdkJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsKICAgIGNvbnN0IHsgcm93cywgY29sdW1ucyB9ID0gZ2V0Um93c0FuZENvbHVtbnMoKTsKICAgIGlmICghcm93cy5sZW5ndGgpIHsgVG9hc3Quc2hvdygnTm90aGluZyB0byBleHBvcnQuJywgJ2Vycm9yJyk7IHJldHVybjsgfQogICAgZG93bmxvYWRCbG9iKGAke2ZpbGVuYW1lQmFzZX0uY3N2YCwgJ3RleHQvY3N2O2NoYXJzZXQ9dXRmLTgnLCB0b0NTVkNsaWVudFNpZGUocm93cywgY29sdW1ucykpOwogIH0pOwogIGNvbnN0IHhsc3hCdG4gPSBpY29uQnRuKCdidG4nLCAnZmlsZS1zcHJlYWRzaGVldCcsICdFeHBvcnQgRXhjZWwnKTsKICB4bHN4QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgYXN5bmMgKCkgPT4gewogICAgY29uc3QgeyByb3dzLCBjb2x1bW5zIH0gPSBnZXRSb3dzQW5kQ29sdW1ucygpOwogICAgaWYgKCFyb3dzLmxlbmd0aCkgeyBUb2FzdC5zaG93KCdOb3RoaW5nIHRvIGV4cG9ydC4nLCAnZXJyb3InKTsgcmV0dXJuOyB9CiAgICB0cnkgewogICAgICBjb25zdCByZXMgPSBhd2FpdCBmZXRjaCgnL2FwaS9leHBvcnQnLCB7CiAgICAgICAgbWV0aG9kOiAnUE9TVCcsCiAgICAgICAgaGVhZGVyczogeyAnQ29udGVudC1UeXBlJzogJ2FwcGxpY2F0aW9uL2pzb24nIH0sCiAgICAgICAgYm9keTogSlNPTi5zdHJpbmdpZnkoeyByb3dzLCBjb2x1bW5zLCBmb3JtYXQ6ICd4bHN4JywgZmlsZW5hbWU6IGZpbGVuYW1lQmFzZSwgc2hlZXROYW1lOiBzaGVldE5hbWUgfHwgZmlsZW5hbWVCYXNlIH0pLAogICAgICB9KTsKICAgICAgaWYgKCFyZXMub2spIHRocm93IG5ldyBFcnJvcignRXhwb3J0IGZhaWxlZC4nKTsKICAgICAgY29uc3QgYmxvYiA9IGF3YWl0IHJlcy5ibG9iKCk7CiAgICAgIGNvbnN0IHVybCA9IFVSTC5jcmVhdGVPYmplY3RVUkwoYmxvYik7CiAgICAgIGNvbnN0IGEgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdhJyk7CiAgICAgIGEuaHJlZiA9IHVybDsKICAgICAgYS5kb3dubG9hZCA9IGAke2ZpbGVuYW1lQmFzZX0ueGxzeGA7CiAgICAgIGRvY3VtZW50LmJvZHkuYXBwZW5kQ2hpbGQoYSk7CiAgICAgIGEuY2xpY2soKTsKICAgICAgYS5yZW1vdmUoKTsKICAgICAgVVJMLnJldm9rZU9iamVjdFVSTCh1cmwpOwogICAgfSBjYXRjaCAoZXJyKSB7CiAgICAgIFRvYXN0LnNob3coZXJyLm1lc3NhZ2UgfHwgJ0V4cG9ydCBmYWlsZWQuJywgJ2Vycm9yJyk7CiAgICB9CiAgfSk7CiAgd3JhcC5hcHBlbmQoY3N2QnRuLCB4bHN4QnRuKTsKICByZXR1cm4gd3JhcDsKfQoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIENoYXJ0cyDigJQgQ2hhcnQuanMgYnVpbGRlcnMgKHZhbGlkYXRlZCBjYXRlZ29yaWNhbCBwYWxldHRlLAogICBoYWlybGluZSByZWNlc3NpdmUgZ3JpZGxpbmVzLCBzaW5nbGUgYXhpcywgbGVnZW5kIGFsd2F5cwogICBwcmVzZW50IGZvciAyKyBzZXJpZXMsIGluZGV4LW1vZGUgdG9vbHRpcHMpLgogICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KaWYgKHdpbmRvdy5DaGFydERhdGFMYWJlbHMpIENoYXJ0LnJlZ2lzdGVyKHdpbmRvdy5DaGFydERhdGFMYWJlbHMpOwoKY29uc3QgQ2hhcnRzID0gKCgpID0+IHsKICBjb25zdCByZWdpc3RyeSA9IG5ldyBNYXAoKTsgLy8gY2FudmFzSWQgLT4gQ2hhcnQgaW5zdGFuY2UsIHNvIHJlLXJlbmRlcnMgZGVzdHJveSB0aGUgb2xkIG9uZSBmaXJzdAogIGNvbnN0IE1BWF9MQUJFTEVEX0lURU1TID0gMjA7IC8vIGJleW9uZCB0aGlzLCBwZXItaXRlbSB2YWx1ZSBsYWJlbHMgd291bGQgb3ZlcmxhcCDigJQgcmVseSBvbiB0b29sdGlwcyBpbnN0ZWFkCgogIGZ1bmN0aW9uIGNzc1ZhcihuYW1lKSB7CiAgICByZXR1cm4gZ2V0Q29tcHV0ZWRTdHlsZShkb2N1bWVudC5kb2N1bWVudEVsZW1lbnQpLmdldFByb3BlcnR5VmFsdWUobmFtZSkudHJpbSgpOwogIH0KCiAgY29uc3QgU0VSSUVTX1ZBUlMgPSBbJy0tc2VyaWVzLTEnLCAnLS1zZXJpZXMtMicsICctLXNlcmllcy0zJywgJy0tc2VyaWVzLTQnLCAnLS1zZXJpZXMtNScsICctLXNlcmllcy02JywgJy0tc2VyaWVzLTcnLCAnLS1zZXJpZXMtOCddOwogIGZ1bmN0aW9uIHNlcmllc0NvbG9yKGluZGV4KSB7CiAgICByZXR1cm4gY3NzVmFyKFNFUklFU19WQVJTW2luZGV4ICUgU0VSSUVTX1ZBUlMubGVuZ3RoXSk7CiAgfQoKICBmdW5jdGlvbiBiYXNlR3JpZCgpIHsKICAgIHJldHVybiB7CiAgICAgIGNvbG9yOiBjc3NWYXIoJy0tZ3JpZGxpbmUnKSwKICAgICAgZHJhd1RpY2tzOiBmYWxzZSwKICAgIH07CiAgfQogIGZ1bmN0aW9uIGJhc2VUaWNrcygpIHsKICAgIHJldHVybiB7IGNvbG9yOiBjc3NWYXIoJy0tdGV4dC1tdXRlZCcpLCBmb250OiB7IHNpemU6IDExIH0gfTsKICB9CiAgZnVuY3Rpb24gYmFzZVRvb2x0aXAoKSB7CiAgICByZXR1cm4gewogICAgICBiYWNrZ3JvdW5kQ29sb3I6IGNzc1ZhcignLS1zdXJmYWNlLTEnKSwKICAgICAgdGl0bGVDb2xvcjogY3NzVmFyKCctLXRleHQtcHJpbWFyeScpLAogICAgICBib2R5Q29sb3I6IGNzc1ZhcignLS10ZXh0LXNlY29uZGFyeScpLAogICAgICBib3JkZXJDb2xvcjogY3NzVmFyKCctLWJvcmRlcicpLAogICAgICBib3JkZXJXaWR0aDogMSwKICAgICAgY29ybmVyUmFkaXVzOiAxMCwKICAgICAgcGFkZGluZzogMTIsCiAgICAgIGJveFBhZGRpbmc6IDQsCiAgICAgIHRpdGxlRm9udDogeyBzaXplOiAxMiwgd2VpZ2h0OiAnNzAwJyB9LAogICAgICBib2R5Rm9udDogeyBzaXplOiAxMiB9LAogICAgfTsKICB9CiAgZnVuY3Rpb24gbGFiZWxDb2xvcigpIHsKICAgIHJldHVybiBjc3NWYXIoJy0tdGV4dC1wcmltYXJ5Jyk7CiAgfQogIC8qKiBTbmFwcHksIHN1YnRsZSBtb3Rpb24g4oCUIGluIHRoZSAxNTAtMzAwbXMgcmFuZ2UgdGhlIHJlZGVzaWduIGNhbGxzIGZvciwgbmV2ZXIgYm91bmN5LiAqLwogIGZ1bmN0aW9uIGJhc2VBbmltYXRpb24oKSB7CiAgICByZXR1cm4geyBkdXJhdGlvbjogMjgwLCBlYXNpbmc6ICdlYXNlT3V0UXVhcnQnIH07CiAgfQoKICBmdW5jdGlvbiBkZXN0cm95KGNhbnZhc0lkKSB7CiAgICBpZiAocmVnaXN0cnkuaGFzKGNhbnZhc0lkKSkgewogICAgICByZWdpc3RyeS5nZXQoY2FudmFzSWQpLmRlc3Ryb3koKTsKICAgICAgcmVnaXN0cnkuZGVsZXRlKGNhbnZhc0lkKTsKICAgIH0KICB9CgogIC8qKiBNdWx0aS1zZXJpZXMgbGluZSBjaGFydCAoZS5nLiB3ZWVrbHkgdHJlbmQgcGVyIHBsYXRmb3JtKS4gT25lIHNlcmllcyBuZWVkcyBubyBsZWdlbmQgYm94LgogICAgICBQZXItcG9pbnQgdmFsdWUgbGFiZWxzIGFyZSBzaG93biBvbmx5IGZvciBhIHNpbmdsZSBzZXJpZXMg4oCUIHdpdGggc2V2ZXJhbCBzZXJpZXMgb3ZlcmxhaWQsCiAgICAgIGxhYmVsaW5nIGV2ZXJ5IHBvaW50IHdvdWxkIG92ZXJsYXAsIHNvIHRob3NlIHJlbHkgb24gdGhlIChzdGlsbC1wcmVzZW50KSBob3ZlciB0b29sdGlwLiAqLwogIGZ1bmN0aW9uIHRyZW5kQ2hhcnQoY2FudmFzSWQsIHsgbGFiZWxzLCBzZXJpZXMsIGZvcm1hdFZhbHVlIH0pIHsKICAgIGRlc3Ryb3koY2FudmFzSWQpOwogICAgY29uc3QgY3R4ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoY2FudmFzSWQpOwogICAgaWYgKCFjdHgpIHJldHVybiBudWxsOwogICAgY29uc3QgZm10ID0gZm9ybWF0VmFsdWUgfHwgKCh2KSA9PiBGb3JtYXQuc21hcnQodikpOwogICAgY29uc3Qgc2hvd0xhYmVscyA9IHNlcmllcy5sZW5ndGggPT09IDEgJiYgbGFiZWxzLmxlbmd0aCA8PSBNQVhfTEFCRUxFRF9JVEVNUzsKCiAgICBjb25zdCBkYXRhc2V0cyA9IHNlcmllcy5tYXAoKHMsIGkpID0+ICh7CiAgICAgIGxhYmVsOiBzLmxhYmVsLAogICAgICBkYXRhOiBzLmRhdGEsCiAgICAgIGJvcmRlckNvbG9yOiBzLmNvbG9yIHx8IHNlcmllc0NvbG9yKGkpLAogICAgICBiYWNrZ3JvdW5kQ29sb3I6IHMuY29sb3IgfHwgc2VyaWVzQ29sb3IoaSksCiAgICAgIGJvcmRlcldpZHRoOiAyLAogICAgICBwb2ludFJhZGl1czogc2hvd0xhYmVscyA/IDMgOiAwLAogICAgICBwb2ludEhvdmVyUmFkaXVzOiA0LAogICAgICBwb2ludEhpdFJhZGl1czogMTIsCiAgICAgIHRlbnNpb246IDAuMjUsCiAgICAgIGZpbGw6IGZhbHNlLAogICAgfSkpOwoKICAgIGNvbnN0IGNoYXJ0ID0gbmV3IENoYXJ0KGN0eCwgewogICAgICB0eXBlOiAnbGluZScsCiAgICAgIGRhdGE6IHsgbGFiZWxzLCBkYXRhc2V0cyB9LAogICAgICBvcHRpb25zOiB7CiAgICAgICAgcmVzcG9uc2l2ZTogdHJ1ZSwKICAgICAgICBtYWludGFpbkFzcGVjdFJhdGlvOiBmYWxzZSwKICAgICAgICBpbnRlcmFjdGlvbjogeyBtb2RlOiAnaW5kZXgnLCBpbnRlcnNlY3Q6IGZhbHNlIH0sCiAgICAgICAgbGF5b3V0OiB7IHBhZGRpbmc6IHsgdG9wOiBzaG93TGFiZWxzID8gMjAgOiA4IH0gfSwKICAgICAgICBhbmltYXRpb246IGJhc2VBbmltYXRpb24oKSwKICAgICAgICBwbHVnaW5zOiB7CiAgICAgICAgICBsZWdlbmQ6IHsKICAgICAgICAgICAgZGlzcGxheTogc2VyaWVzLmxlbmd0aCA+IDEsCiAgICAgICAgICAgIHBvc2l0aW9uOiAnYm90dG9tJywKICAgICAgICAgICAgbGFiZWxzOiB7IGNvbG9yOiBjc3NWYXIoJy0tdGV4dC1zZWNvbmRhcnknKSwgdXNlUG9pbnRTdHlsZTogdHJ1ZSwgcG9pbnRTdHlsZTogJ2xpbmUnLCBib3hXaWR0aDogMTYsIHBhZGRpbmc6IDE2LCBmb250OiB7IHNpemU6IDExIH0gfSwKICAgICAgICAgIH0sCiAgICAgICAgICB0b29sdGlwOiB7IC4uLmJhc2VUb29sdGlwKCksIHVzZVBvaW50U3R5bGU6IHRydWUgfSwKICAgICAgICAgIGRhdGFsYWJlbHM6IHNob3dMYWJlbHMKICAgICAgICAgICAgPyB7IGFsaWduOiAndG9wJywgYW5jaG9yOiAnZW5kJywgY29sb3I6IGxhYmVsQ29sb3IoKSwgZm9udDogeyBzaXplOiAxMSwgd2VpZ2h0OiAnNjAwJyB9LCBmb3JtYXR0ZXI6ICh2KSA9PiBmbXQodikgfQogICAgICAgICAgICA6IHsgZGlzcGxheTogZmFsc2UgfSwKICAgICAgICB9LAogICAgICAgIHNjYWxlczogewogICAgICAgICAgeDogeyBncmlkOiB7IGRpc3BsYXk6IGZhbHNlIH0sIHRpY2tzOiBiYXNlVGlja3MoKSB9LAogICAgICAgICAgeTogeyBncmlkOiBiYXNlR3JpZCgpLCB0aWNrczogYmFzZVRpY2tzKCksIGJvcmRlcjogeyBkaXNwbGF5OiBmYWxzZSB9LCBiZWdpbkF0WmVybzogdHJ1ZSB9LAogICAgICAgIH0sCiAgICAgIH0sCiAgICB9KTsKICAgIHJlZ2lzdHJ5LnNldChjYW52YXNJZCwgY2hhcnQpOwogICAgcmV0dXJuIGNoYXJ0OwogIH0KCiAgLyoqIFNpbmdsZS1tZXRyaWMgYmFyIGNoYXJ0IGFjcm9zcyBwbGF0Zm9ybXMgKGlkZW50aXR5IGVuY29kaW5nIOKAlCBlYWNoIGJhciBJUyBhIHBsYXRmb3JtKS4gKi8KICBmdW5jdGlvbiBwbGF0Zm9ybUJhckNoYXJ0KGNhbnZhc0lkLCB7IGxhYmVscywgZGF0YSwgY29sb3JzLCBmb3JtYXRWYWx1ZSB9KSB7CiAgICBkZXN0cm95KGNhbnZhc0lkKTsKICAgIGNvbnN0IGN0eCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGNhbnZhc0lkKTsKICAgIGlmICghY3R4KSByZXR1cm4gbnVsbDsKICAgIGNvbnN0IGZtdCA9IGZvcm1hdFZhbHVlIHx8ICgodikgPT4gRm9ybWF0LnNtYXJ0KHYpKTsKICAgIGNvbnN0IHNob3dMYWJlbHMgPSBsYWJlbHMubGVuZ3RoIDw9IE1BWF9MQUJFTEVEX0lURU1TOwoKICAgIGNvbnN0IGNoYXJ0ID0gbmV3IENoYXJ0KGN0eCwgewogICAgICB0eXBlOiAnYmFyJywKICAgICAgZGF0YTogewogICAgICAgIGxhYmVscywKICAgICAgICBkYXRhc2V0czogWwogICAgICAgICAgewogICAgICAgICAgICBkYXRhLAogICAgICAgICAgICBiYWNrZ3JvdW5kQ29sb3I6IGNvbG9ycywKICAgICAgICAgICAgYm9yZGVyUmFkaXVzOiA0LAogICAgICAgICAgICBtYXhCYXJUaGlja25lc3M6IDI4LAogICAgICAgICAgICBib3JkZXJTa2lwcGVkOiAnYm90dG9tJywKICAgICAgICAgIH0sCiAgICAgICAgXSwKICAgICAgfSwKICAgICAgb3B0aW9uczogewogICAgICAgIHJlc3BvbnNpdmU6IHRydWUsCiAgICAgICAgbWFpbnRhaW5Bc3BlY3RSYXRpbzogZmFsc2UsCiAgICAgICAgbGF5b3V0OiB7IHBhZGRpbmc6IHsgdG9wOiBzaG93TGFiZWxzID8gMjAgOiA4IH0gfSwKICAgICAgICBhbmltYXRpb246IGJhc2VBbmltYXRpb24oKSwKICAgICAgICBwbHVnaW5zOiB7CiAgICAgICAgICBsZWdlbmQ6IHsgZGlzcGxheTogZmFsc2UgfSwKICAgICAgICAgIHRvb2x0aXA6IGJhc2VUb29sdGlwKCksCiAgICAgICAgICBkYXRhbGFiZWxzOiBzaG93TGFiZWxzCiAgICAgICAgICAgID8geyBhbGlnbjogJ2VuZCcsIGFuY2hvcjogJ2VuZCcsIGNvbG9yOiBsYWJlbENvbG9yKCksIGZvbnQ6IHsgc2l6ZTogMTEsIHdlaWdodDogJzYwMCcgfSwgZm9ybWF0dGVyOiAodikgPT4gZm10KHYpIH0KICAgICAgICAgICAgOiB7IGRpc3BsYXk6IGZhbHNlIH0sCiAgICAgICAgfSwKICAgICAgICBzY2FsZXM6IHsKICAgICAgICAgIHg6IHsgZ3JpZDogeyBkaXNwbGF5OiBmYWxzZSB9LCB0aWNrczogYmFzZVRpY2tzKCkgfSwKICAgICAgICAgIHk6IHsgZ3JpZDogYmFzZUdyaWQoKSwgdGlja3M6IGJhc2VUaWNrcygpLCBib3JkZXI6IHsgZGlzcGxheTogZmFsc2UgfSwgYmVnaW5BdFplcm86IHRydWUgfSwKICAgICAgICB9LAogICAgICB9LAogICAgfSk7CiAgICByZWdpc3RyeS5zZXQoY2FudmFzSWQsIGNoYXJ0KTsKICAgIHJldHVybiBjaGFydDsKICB9CgogIC8qKiBQaWUgY2hhcnQgKGEgaGFuZGZ1bCBvZiBjYXRlZ29yaWVzIG9ubHkg4oCUIGUuZy4gQ2FtcGFpZ24gUGVyZm9ybWFuY2UncyBBZHMvT3JnYW5pYyBzcGxpdCkuCiAgICAgIFNsaWNlIGxhYmVscyBzaG93IGJvdGggc2hhcmUtb2Ytd2hvbGUgYW5kIHRoZSBhY3R1YWwgdmFsdWUsIHBlciB0aGUgIm5vIGhvdmVyIHJlcXVpcmVkIiBnb2FsLiAqLwogIGZ1bmN0aW9uIHBpZUNoYXJ0KGNhbnZhc0lkLCB7IGxhYmVscywgZGF0YSwgY29sb3JzLCBmb3JtYXRWYWx1ZSB9KSB7CiAgICBkZXN0cm95KGNhbnZhc0lkKTsKICAgIGNvbnN0IGN0eCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGNhbnZhc0lkKTsKICAgIGlmICghY3R4KSByZXR1cm4gbnVsbDsKICAgIGNvbnN0IGZtdCA9IGZvcm1hdFZhbHVlIHx8ICgodikgPT4gRm9ybWF0LnNtYXJ0KHYpKTsKICAgIGNvbnN0IHRvdGFsID0gZGF0YS5yZWR1Y2UoKHN1bSwgdikgPT4gc3VtICsgKHYgfHwgMCksIDApOwoKICAgIGNvbnN0IGNoYXJ0ID0gbmV3IENoYXJ0KGN0eCwgewogICAgICB0eXBlOiAncGllJywKICAgICAgZGF0YTogewogICAgICAgIGxhYmVscywKICAgICAgICBkYXRhc2V0czogW3sgZGF0YSwgYmFja2dyb3VuZENvbG9yOiBjb2xvcnMsIGJvcmRlckNvbG9yOiBjc3NWYXIoJy0tc3VyZmFjZS0xJyksIGJvcmRlcldpZHRoOiAyIH1dLAogICAgICB9LAogICAgICBvcHRpb25zOiB7CiAgICAgICAgcmVzcG9uc2l2ZTogdHJ1ZSwKICAgICAgICBtYWludGFpbkFzcGVjdFJhdGlvOiBmYWxzZSwKICAgICAgICBhbmltYXRpb246IGJhc2VBbmltYXRpb24oKSwKICAgICAgICBwbHVnaW5zOiB7CiAgICAgICAgICBsZWdlbmQ6IHsgZGlzcGxheTogdHJ1ZSwgcG9zaXRpb246ICdib3R0b20nLCBsYWJlbHM6IHsgY29sb3I6IGNzc1ZhcignLS10ZXh0LXNlY29uZGFyeScpLCBib3hXaWR0aDogMTIsIHBhZGRpbmc6IDE2LCBmb250OiB7IHNpemU6IDExIH0gfSB9LAogICAgICAgICAgdG9vbHRpcDogYmFzZVRvb2x0aXAoKSwKICAgICAgICAgIGRhdGFsYWJlbHM6IHsKICAgICAgICAgICAgY29sb3I6ICcjZmZmJywKICAgICAgICAgICAgZm9udDogeyBzaXplOiAxMiwgd2VpZ2h0OiAnNzAwJyB9LAogICAgICAgICAgICBmb3JtYXR0ZXI6ICh2KSA9PiB7CiAgICAgICAgICAgICAgY29uc3QgcGN0ID0gdG90YWwgPyBNYXRoLnJvdW5kKCh2IC8gdG90YWwpICogMTAwMCkgLyAxMCA6IDA7CiAgICAgICAgICAgICAgcmV0dXJuIGAke3BjdH0lXG4ke2ZtdCh2KX1gOwogICAgICAgICAgICB9LAogICAgICAgICAgfSwKICAgICAgICB9LAogICAgICB9LAogICAgfSk7CiAgICByZWdpc3RyeS5zZXQoY2FudmFzSWQsIGNoYXJ0KTsKICAgIHJldHVybiBjaGFydDsKICB9CgogIGZ1bmN0aW9uIGRlc3Ryb3lBbGwoKSB7CiAgICBbLi4ucmVnaXN0cnkua2V5cygpXS5mb3JFYWNoKGRlc3Ryb3kpOwogIH0KCiAgcmV0dXJuIHsgdHJlbmRDaGFydCwgcGxhdGZvcm1CYXJDaGFydCwgcGllQ2hhcnQsIHNlcmllc0NvbG9yLCBkZXN0cm95LCBkZXN0cm95QWxsIH07Cn0pKCk7CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgRGFzaGJvYXJkIHRhYjogYSBtZXRyaWMtZm9jdXNlZCBwcmVtaXVtIEJJIGRhc2hib2FyZC4gQSBzaW5nbGUKICAgTWV0cmljIHNlbGVjdG9yIChkeW5hbWljYWxseSBwb3B1bGF0ZWQgZnJvbSB3aGF0ZXZlciB0aGUKICAgc2VsZWN0ZWQgcGxhdGZvcm0ncyBkYXRhIGFjdHVhbGx5IGhhcyDigJQgbmV2ZXIgaGFyZGNvZGVkKSBkcml2ZXMKICAgdGhlIEtQSSBjYXJkcywgd2Vla2x5IHRyZW5kLCBwbGF0Zm9ybS9jYW1wYWlnbi9jb250ZW50LXR5cGUKICAgYnJlYWtkb3ducywgYW5kIHRoZSBUb3AgUGVyZm9ybWluZyBQb3N0cyByYW5raW5nIHRvZ2V0aGVyOwogICBQbGF0Zm9ybS9kYXRlL2NhbXBhaWduL2NvbnRlbnQtdHlwZSBmaWx0ZXJpbmcgY29tZXMgZnJvbSB0aGUKICAgc2hhcmVkIGZpbHRlciBiYXIuIEV2ZXJ5IGNoYXJ0IHNob3dzIGl0cyB2YWx1ZXMgZGlyZWN0bHkgKHZpYQogICBjaGFydGpzLXBsdWdpbi1kYXRhbGFiZWxzKSBzbyBub3RoaW5nIHJlcXVpcmVzIGEgaG92ZXIgdG8gcmVhZC4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmNvbnN0IERhc2hib2FyZCA9ICgoKSA9PiB7CiAgbGV0IHJvb3Q7CiAgbGV0IG1ldHJpYyA9ICd2aWV3cyc7CiAgbGV0IG1ldHJpY09wdGlvbnMgPSBbXTsKCiAgZnVuY3Rpb24gb3B0aW9uRm9yKGtleSkgewogICAgcmV0dXJuIG1ldHJpY09wdGlvbnMuZmluZCgobSkgPT4gbS5rZXkgPT09IGtleSk7CiAgfQogIGZ1bmN0aW9uIG1ldHJpY0xhYmVsKGtleSkgewogICAgY29uc3Qgb3B0ID0gb3B0aW9uRm9yKGtleSk7CiAgICByZXR1cm4gb3B0ID8gb3B0LmxhYmVsIDoga2V5OwogIH0KICBmdW5jdGlvbiBtZXRyaWNVbml0KGtleSkgewogICAgY29uc3Qgb3B0ID0gb3B0aW9uRm9yKGtleSk7CiAgICByZXR1cm4gb3B0ID8gb3B0LnVuaXQgOiAnbnVtYmVyJzsKICB9CiAgZnVuY3Rpb24gZm9ybWF0TWV0cmljVmFsdWUoa2V5LCB2YWx1ZSkgewogICAgY29uc3QgdW5pdCA9IG1ldHJpY1VuaXQoa2V5KTsKICAgIGlmICh2YWx1ZSA9PT0gbnVsbCB8fCB2YWx1ZSA9PT0gdW5kZWZpbmVkKSByZXR1cm4gJ+KAlCc7CiAgICBpZiAodW5pdCA9PT0gJ2R1cmF0aW9uJykgcmV0dXJuIEZvcm1hdC5kdXJhdGlvbih2YWx1ZSk7CiAgICByZXR1cm4gRm9ybWF0LnNtYXJ0KHZhbHVlKTsKICB9CgogIGZ1bmN0aW9uIHNoZWxsKCkgewogICAgcm9vdC5pbm5lckhUTUwgPSAnJzsKCiAgICBjb25zdCBjb250cm9scyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY29udHJvbHMuY2xhc3NOYW1lID0gJ2Rhc2hib2FyZC1jb250cm9scyc7CiAgICBjb25zdCBsYWJlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2xhYmVsJyk7CiAgICBsYWJlbC50ZXh0Q29udGVudCA9ICdNZXRyaWMnOwogICAgbGFiZWwuc2V0QXR0cmlidXRlKCdmb3InLCAnZGFzaGJvYXJkTWV0cmljU2VsZWN0Jyk7CiAgICBjb25zdCBtZXRyaWNTZWxlY3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzZWxlY3QnKTsKICAgIG1ldHJpY1NlbGVjdC5pZCA9ICdkYXNoYm9hcmRNZXRyaWNTZWxlY3QnOwogICAgbWV0cmljT3B0aW9ucy5mb3JFYWNoKChtKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOwogICAgICBvcHQudmFsdWUgPSBtLmtleTsKICAgICAgb3B0LnRleHRDb250ZW50ID0gbS5sYWJlbDsKICAgICAgaWYgKG0ua2V5ID09PSBtZXRyaWMpIG9wdC5zZWxlY3RlZCA9IHRydWU7CiAgICAgIG1ldHJpY1NlbGVjdC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgICBtZXRyaWNTZWxlY3QuYWRkRXZlbnRMaXN0ZW5lcignY2hhbmdlJywgKCkgPT4gewogICAgICBtZXRyaWMgPSBtZXRyaWNTZWxlY3QudmFsdWU7CiAgICAgIHJlZnJlc2hGb3JNZXRyaWMoKTsKICAgIH0pOwogICAgY29udHJvbHMuYXBwZW5kKGxhYmVsLCBtZXRyaWNTZWxlY3QpOwogICAgcm9vdC5hcHBlbmRDaGlsZChjb250cm9scyk7CgogICAgY29uc3Qga3BpVGl0bGUgPSB0ZXh0RWwoJ2RpdicsICdLZXkgcGVyZm9ybWFuY2UgaW5kaWNhdG9ycycsICdzZWN0aW9uLXRpdGxlJyk7CiAgICBjb25zdCBrcGlHcmlkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBrcGlHcmlkLmNsYXNzTmFtZSA9ICdzdGF0LWdyaWQnOwogICAga3BpR3JpZC5pZCA9ICdrcGlHcmlkJzsKICAgIHJvb3QuYXBwZW5kKGtwaVRpdGxlLCBrcGlHcmlkKTsKCiAgICBjb25zdCBjaGFydHNUaXRsZSA9IHRleHRFbCgnZGl2JywgJ1RyZW5kICYgcGVyZm9ybWFuY2UgYnJlYWtkb3duJywgJ3NlY3Rpb24tdGl0bGUnKTsKICAgIHJvb3QuYXBwZW5kKGNoYXJ0c1RpdGxlKTsKCiAgICBjb25zdCB0cmVuZENhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRyZW5kQ2FyZC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICBjb25zdCB0cmVuZEhlYWRlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdHJlbmRIZWFkZXIuY2xhc3NOYW1lID0gJ2NhcmQtaGVhZGVyJzsKICAgIHRyZW5kSGVhZGVyLmFwcGVuZENoaWxkKHRleHRFbCgnaDMnLCAnV2Vla2x5IHBlcmZvcm1hbmNlJykpOwogICAgdHJlbmRIZWFkZXIuZmlyc3RDaGlsZC5pZCA9ICd0cmVuZENhcmRUaXRsZSc7CiAgICB0cmVuZENhcmQuYXBwZW5kQ2hpbGQodHJlbmRIZWFkZXIpOwogICAgY29uc3QgdHJlbmRDaGFydFdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRyZW5kQ2hhcnRXcmFwLmNsYXNzTmFtZSA9ICdjaGFydC13cmFwIHRhbGwnOwogICAgdHJlbmRDaGFydFdyYXAuaWQgPSAndHJlbmRDaGFydFdyYXAnOwogICAgdHJlbmRDaGFydFdyYXAuaW5uZXJIVE1MID0gJzxjYW52YXMgaWQ9InRyZW5kQ2FudmFzIj48L2NhbnZhcz4nOwogICAgdHJlbmRDYXJkLmFwcGVuZENoaWxkKHRyZW5kQ2hhcnRXcmFwKTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQodHJlbmRDYXJkKTsKCiAgICBjb25zdCBncmlkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBncmlkLmNsYXNzTmFtZSA9ICdjYXJkLWdyaWQgZXZlbic7CiAgICBncmlkLnN0eWxlLm1hcmdpblRvcCA9ICcxNnB4JzsKCiAgICBjb25zdCBicmVha2Rvd25DYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBicmVha2Rvd25DYXJkLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGNvbnN0IGJyZWFrZG93bkhlYWRlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYnJlYWtkb3duSGVhZGVyLmNsYXNzTmFtZSA9ICdjYXJkLWhlYWRlcic7CiAgICBicmVha2Rvd25IZWFkZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMycsICcnKSk7CiAgICBicmVha2Rvd25IZWFkZXIuZmlyc3RDaGlsZC5pZCA9ICdicmVha2Rvd25DYXJkVGl0bGUnOwogICAgYnJlYWtkb3duQ2FyZC5hcHBlbmRDaGlsZChicmVha2Rvd25IZWFkZXIpOwogICAgY29uc3QgYnJlYWtkb3duV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYnJlYWtkb3duV3JhcC5jbGFzc05hbWUgPSAnY2hhcnQtd3JhcCc7CiAgICBicmVha2Rvd25XcmFwLmlkID0gJ2JyZWFrZG93bkNoYXJ0V3JhcCc7CiAgICBicmVha2Rvd25XcmFwLmlubmVySFRNTCA9ICc8Y2FudmFzIGlkPSJicmVha2Rvd25DYW52YXMiPjwvY2FudmFzPic7CiAgICBicmVha2Rvd25DYXJkLmFwcGVuZENoaWxkKGJyZWFrZG93bldyYXApOwoKICAgIGNvbnN0IGNvbnRlbnRUeXBlQ2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY29udGVudFR5cGVDYXJkLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGNvbnN0IGNvbnRlbnRUeXBlSGVhZGVyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjb250ZW50VHlwZUhlYWRlci5jbGFzc05hbWUgPSAnY2FyZC1oZWFkZXInOwogICAgY29udGVudFR5cGVIZWFkZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMycsICcnKSk7CiAgICBjb250ZW50VHlwZUhlYWRlci5maXJzdENoaWxkLmlkID0gJ2NvbnRlbnRUeXBlQ2FyZFRpdGxlJzsKICAgIGNvbnRlbnRUeXBlQ2FyZC5hcHBlbmRDaGlsZChjb250ZW50VHlwZUhlYWRlcik7CiAgICBjb25zdCBjb250ZW50VHlwZVdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNvbnRlbnRUeXBlV3JhcC5jbGFzc05hbWUgPSAnY2hhcnQtd3JhcCc7CiAgICBjb250ZW50VHlwZVdyYXAuaWQgPSAnY29udGVudFR5cGVDaGFydFdyYXAnOwogICAgY29udGVudFR5cGVXcmFwLmlubmVySFRNTCA9ICc8Y2FudmFzIGlkPSJjb250ZW50VHlwZUNhbnZhcyI+PC9jYW52YXM+JzsKICAgIGNvbnRlbnRUeXBlQ2FyZC5hcHBlbmRDaGlsZChjb250ZW50VHlwZVdyYXApOwoKICAgIGdyaWQuYXBwZW5kKGJyZWFrZG93bkNhcmQsIGNvbnRlbnRUeXBlQ2FyZCk7CiAgICByb290LmFwcGVuZENoaWxkKGdyaWQpOwoKICAgIGNvbnN0IHRvcFRpdGxlID0gdGV4dEVsKCdkaXYnLCAnVG9wLXBlcmZvcm1pbmcgcG9zdHMnLCAnc2VjdGlvbi10aXRsZScpOwogICAgY29uc3QgdG9wQ2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdG9wQ2FyZC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICBjb25zdCB0b3BIZWFkZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRvcEhlYWRlci5jbGFzc05hbWUgPSAnY2FyZC1oZWFkZXInOwogICAgdG9wSGVhZGVyLmFwcGVuZENoaWxkKHRleHRFbCgnaDMnLCAnUmFua2VkIGJ5IHNlbGVjdGVkIG1ldHJpYycpKTsKICAgIHRvcENhcmQuYXBwZW5kQ2hpbGQodG9wSGVhZGVyKTsKICAgIGNvbnN0IHRhYmxlV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGFibGVXcmFwLmNsYXNzTmFtZSA9ICd0YWJsZS1zY3JvbGwnOwogICAgdGFibGVXcmFwLmlkID0gJ3RvcFBvc3RzVGFibGUnOwogICAgdG9wQ2FyZC5hcHBlbmRDaGlsZCh0YWJsZVdyYXApOwogICAgcm9vdC5hcHBlbmQodG9wVGl0bGUsIHRvcENhcmQpOwogIH0KCiAgLyoqIEEgc21hbGwgbGFiZWwrdmFsdWUgcGFpciB1c2VkIGluc2lkZSB0aGUgQmVzdCBQZXJmb3JtaW5nIFBvc3QgY2FyZCdzIG1ldHJpY3MgY29sdW1uLiBgdmFyaWFudGAgKCdwcmltYXJ5Jy8nc2Vjb25kYXJ5JykgY29udHJvbHMgc2l6ZSBhbmQgd2hldGhlciBhIGRpdmlkZXIgcnVsZSBzaXRzIGFib3ZlIGl0LiAqLwogIGZ1bmN0aW9uIG1ldHJpY0Jsb2NrKGxhYmVsLCB2YWx1ZSwgdmFyaWFudCkgewogICAgY29uc3QgYmxvY2sgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGJsb2NrLmNsYXNzTmFtZSA9IHZhcmlhbnQgPT09ICdzZWNvbmRhcnknID8gJ3Bvc3QtdGlsZS1tZXRyaWMtYmxvY2sgc2Vjb25kYXJ5JyA6ICdwb3N0LXRpbGUtbWV0cmljLWJsb2NrJzsKICAgIGNvbnN0IHZhbHVlRWwgPSB0ZXh0RWwoJ2RpdicsIHZhbHVlLCAncG9zdC10aWxlLW1ldHJpYy12YWx1ZScpOwogICAgYmxvY2suYXBwZW5kKHRleHRFbCgnZGl2JywgbGFiZWwsICdwb3N0LXRpbGUtbWV0cmljLWxhYmVsJyksIHZhbHVlRWwpOwogICAgcmVxdWVzdEFuaW1hdGlvbkZyYW1lKCgpID0+IGZpdFN0YXRWYWx1ZSh2YWx1ZUVsLCAxMikpOwogICAgcmV0dXJuIGJsb2NrOwogIH0KCiAgLyoqIEEgZmVhdHVyZWQgbGFuZHNjYXBlIGNhcmQgKDMgS1BJLXRpbGUtd2lkdGhzLCBzYW1lIGZpeGVkIGhlaWdodCBhcyB0aGUgcmVzdCBvZiB0aGUKICAgICAgcm93KTogdGhlIHRvcC10aWVkIHBvc3QncyBjYXB0aW9uICh3cmFwcyB1cCB0byAzIGxpbmVzKSB3aXRoIGEgcGxhdGZvcm0tY29sb3IgZG90CiAgICAgICsgcGxhdGZvcm0gbmFtZSArIGRhdGUgb24gdGhlIGxlZnQgKHdoZW4gdGhlcmUncyBtb3JlIHRoYW4gb25lIHRpZSwgdGhlIGV4dHJhIGNvdW50CiAgICAgIGlzIGZvbGRlZCBpbnRvIHRoYXQgc2FtZSBtZXRhIGxpbmUgcmF0aGVyIHRoYW4gbGlzdGluZyBldmVyeSB0aWVkIHBvc3QsIHNvIHRoZSBjYXJkCiAgICAgIG5ldmVyIGhhcyB0byBncm93IHRhbGxlciB0aGFuIGl0cyBuZWlnaGJvcnMpOyB0aGUgc2VsZWN0ZWQgbWV0cmljIChsYXJnZSkgYW5kCiAgICAgIEN1cnJlbnQgRm9sbG93ZXJzIChzbWFsbGVyLCBiZWxvdyBhIGRpdmlkZXIpIHN0YWNrZWQgaW4gYSBuYXJyb3cgY29sdW1uIG9uIHRoZSByaWdodC4gKi8KICBmdW5jdGlvbiBiZXN0UG9zdHNUaWxlKGxhYmVsLCBwb3N0cywgY3VycmVudEZvbGxvd2VycykgewogICAgY29uc3QgdGlsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGlsZS5jbGFzc05hbWUgPSAnc3RhdC10aWxlIHBvc3QtdGlsZSc7CgogICAgY29uc3QgbWFpbiA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgbWFpbi5jbGFzc05hbWUgPSAncG9zdC10aWxlLW1haW4nOwogICAgbWFpbi5hcHBlbmRDaGlsZChzdGF0TGFiZWxFbChsYWJlbCwgJ3Ryb3BoeScsICdnb2xkJykpOwoKICAgIGNvbnN0IGhhc1Bvc3RzID0gQm9vbGVhbihwb3N0cyAmJiBwb3N0cy5sZW5ndGgpOwogICAgaWYgKGhhc1Bvc3RzKSB7CiAgICAgIGNvbnN0IHBsYXRmb3JtT3B0aW9ucyA9ICh3aW5kb3cuX19maWx0ZXJPcHRpb25zQ2FjaGUgfHwgeyBwbGF0Zm9ybXM6IFtdIH0pLnBsYXRmb3JtczsKICAgICAgY29uc3QgcHJpbWFyeSA9IHBvc3RzWzBdOwogICAgICBjb25zdCBwbGF0TWV0YSA9IHBsYXRmb3JtT3B0aW9ucy5maW5kKChwKSA9PiBwLmlkID09PSBwcmltYXJ5LnBsYXRmb3JtKSB8fCB7IGxhYmVsOiBwcmltYXJ5LnBsYXRmb3JtLCBjb2xvcjogJ3ZhcigtLXNlcmllcy0xKScgfTsKICAgICAgY29uc3QgY2FwdGlvbiA9IHByaW1hcnkuY2FwdGlvbiB8fCAnKG5vIGNhcHRpb24pJzsKICAgICAgY29uc3QgY2FwdGlvbkVsID0gdGV4dEVsKCdkaXYnLCBjYXB0aW9uLCAncG9zdC10aWxlLWNhcHRpb24nKTsKICAgICAgY2FwdGlvbkVsLnRpdGxlID0gY2FwdGlvbjsKICAgICAgbWFpbi5hcHBlbmRDaGlsZChjYXB0aW9uRWwpOwogICAgICBjb25zdCB0aWVkTm90ZSA9IHBvc3RzLmxlbmd0aCA+IDEgPyBgIMK3ICske3Bvc3RzLmxlbmd0aCAtIDF9IG1vcmUgdGllZGAgOiAnJzsKICAgICAgY29uc3QgbWV0YUxpbmUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgbWV0YUxpbmUuY2xhc3NOYW1lID0gJ3Bvc3QtdGlsZS1tZXRhJzsKICAgICAgY29uc3QgZG90ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICBkb3QuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLWRvdCc7CiAgICAgIGRvdC5zdHlsZS5iYWNrZ3JvdW5kID0gcGxhdE1ldGEuY29sb3I7CiAgICAgIG1ldGFMaW5lLmFwcGVuZChkb3QsIGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKGAke3BsYXRNZXRhLmxhYmVsfSDCtyAke0Zvcm1hdC5kYXRlKHByaW1hcnkucHVibGlzaF9kYXRlKX0ke3RpZWROb3RlfWApKTsKICAgICAgbWFpbi5hcHBlbmRDaGlsZChtZXRhTGluZSk7CiAgICB9IGVsc2UgewogICAgICBtYWluLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgJ05vIGRhdGEgeWV0JywgJ3Bvc3QtdGlsZS1jYXB0aW9uIG11dGVkJykpOwogICAgfQogICAgdGlsZS5hcHBlbmRDaGlsZChtYWluKTsKCiAgICBjb25zdCBoYXNGb2xsb3dlcnMgPSBjdXJyZW50Rm9sbG93ZXJzICE9PSBudWxsICYmIGN1cnJlbnRGb2xsb3dlcnMgIT09IHVuZGVmaW5lZDsKICAgIGlmIChoYXNQb3N0cyB8fCBoYXNGb2xsb3dlcnMpIHsKICAgICAgdGlsZS5hcHBlbmRDaGlsZChPYmplY3QuYXNzaWduKGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpLCB7IGNsYXNzTmFtZTogJ3Bvc3QtdGlsZS1kaXZpZGVyJyB9KSk7CiAgICAgIGNvbnN0IG1ldHJpY3MgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgbWV0cmljcy5jbGFzc05hbWUgPSAncG9zdC10aWxlLW1ldHJpY3MnOwogICAgICBpZiAoaGFzUG9zdHMpIG1ldHJpY3MuYXBwZW5kQ2hpbGQobWV0cmljQmxvY2sobWV0cmljTGFiZWwobWV0cmljKSwgZm9ybWF0TWV0cmljVmFsdWUobWV0cmljLCBwb3N0c1swXS52YWx1ZSksICdwcmltYXJ5JykpOwogICAgICBpZiAoaGFzRm9sbG93ZXJzKSBtZXRyaWNzLmFwcGVuZENoaWxkKG1ldHJpY0Jsb2NrKCdDdXJyZW50IEZvbGxvd2VycycsIEZvcm1hdC5udW1iZXIoY3VycmVudEZvbGxvd2VycyksICdzZWNvbmRhcnknKSk7CiAgICAgIHRpbGUuYXBwZW5kQ2hpbGQobWV0cmljcyk7CiAgICB9CiAgICByZXR1cm4gdGlsZTsKICB9CgogIC8qKiBJY29uLWJhZGdlICsgdGV4dCBsYWJlbCByb3csIHNoYXJlZCBieSBldmVyeSBLUEkgdGlsZSBiZWxvdyAobWF0Y2hlcyB0aGUgcmVmZXJlbmNlIGRhc2hib2FyZCdzIGNvbG9yZWQgcGVyLWNhcmQgaWNvbnMpLiBgdmFyaWFudGAgcGlja3MgdGhlIGJhZGdlIGNvbG9yOiB2MS12NiBtYXAgdG8gdGhlIHBsYXRmb3JtIHNlcmllcyBwYWxldHRlLCAnZ29sZCcgaXMgcmVzZXJ2ZWQgZm9yIHRoZSBicmFuZC1hY2NlbnQgdGlsZS4gKi8KICBmdW5jdGlvbiBzdGF0TGFiZWxFbCh0ZXh0LCBpY29uLCB2YXJpYW50KSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB3cmFwLmNsYXNzTmFtZSA9ICdzdGF0LWxhYmVsJzsKICAgIGNvbnN0IGJhZGdlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgYmFkZ2UuY2xhc3NOYW1lID0gYHN0YXQtaWNvbiAke3ZhcmlhbnR9YDsKICAgIGJhZGdlLmlubmVySFRNTCA9IGA8aSBkYXRhLWx1Y2lkZT0iJHtpY29ufSIgc3R5bGU9IndpZHRoOjE2cHg7aGVpZ2h0OjE2cHg7Ij48L2k+YDsKICAgIHdyYXAuYXBwZW5kKGJhZGdlLCBkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZSh0ZXh0KSk7CiAgICByZXR1cm4gd3JhcDsKICB9CgogIGZ1bmN0aW9uIHN0YXRUaWxlKGxhYmVsLCB2YWx1ZSwgZm9ybWF0Rm4sIGljb24sIHZhcmlhbnQpIHsKICAgIGNvbnN0IHRpbGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRpbGUuY2xhc3NOYW1lID0gJ3N0YXQtdGlsZSc7CiAgICBjb25zdCB2YWx1ZUVsID0gdGV4dEVsKCdkaXYnLCAnJywgJ3N0YXQtdmFsdWUnKTsKICAgIHRpbGUuYXBwZW5kKHN0YXRMYWJlbEVsKGxhYmVsLCBpY29uLCB2YXJpYW50KSwgdmFsdWVFbCk7CiAgICBjb25zdCBmbXQgPSBmb3JtYXRGbiB8fCAoKHYpID0+IEZvcm1hdC5udW1iZXIodikpOwogICAgaWYgKHR5cGVvZiB2YWx1ZSA9PT0gJ251bWJlcicgJiYgTnVtYmVyLmlzRmluaXRlKHZhbHVlKSkgewogICAgICBhbmltYXRlQ291bnQodmFsdWVFbCwgMCwgdmFsdWUsIDkwMCwgZm10KTsKICAgIH0gZWxzZSB7CiAgICAgIHZhbHVlRWwudGV4dENvbnRlbnQgPSBmbXQodmFsdWUpOwogICAgICByZXF1ZXN0QW5pbWF0aW9uRnJhbWUoKCkgPT4gZml0U3RhdFZhbHVlKHZhbHVlRWwpKTsKICAgIH0KICAgIHJldHVybiB0aWxlOwogIH0KCiAgLyoqICJGb2xsb3dlcnMgR3Jvd3RoIiB0aWxlOiBhbiBhYnNvbHV0ZS1kaWZmZXJlbmNlIHN0YXQtdmFsdWUgcGx1cyBhIHBlcmNlbnRhZ2UgZGVsdGEgbGluZSAoYXJyb3cgKyBjb2xvciBkcml2ZW4gYnkgRm9ybWF0LmRlbHRhQ2xhc3MsIHNhbWUgY29udmVudGlvbiBhcyB0aGUgQ29tcGFyaXNvbnMgcGFnZSdzIHN0YXQgdGlsZXMpLiAqLwogIGZ1bmN0aW9uIGZvbGxvd2Vyc0dyb3d0aFRpbGUoY2hhbmdlLCBjaGFuZ2VQY3QpIHsKICAgIGNvbnN0IHRpbGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRpbGUuY2xhc3NOYW1lID0gJ3N0YXQtdGlsZSc7CiAgICBjb25zdCB2YWx1ZUVsID0gdGV4dEVsKCdkaXYnLCAnJywgJ3N0YXQtdmFsdWUnKTsKICAgIHRpbGUuYXBwZW5kKHN0YXRMYWJlbEVsKCdGb2xsb3dlcnMgR3Jvd3RoJywgJ3RyZW5kaW5nLXVwJywgJ3YzJyksIHZhbHVlRWwpOwogICAgaWYgKGNoYW5nZSA9PT0gbnVsbCB8fCBjaGFuZ2UgPT09IHVuZGVmaW5lZCkgewogICAgICB2YWx1ZUVsLnRleHRDb250ZW50ID0gJ+KAlCc7CiAgICAgIHJlcXVlc3RBbmltYXRpb25GcmFtZSgoKSA9PiBmaXRTdGF0VmFsdWUodmFsdWVFbCkpOwogICAgfSBlbHNlIHsKICAgICAgYW5pbWF0ZUNvdW50KHZhbHVlRWwsIDAsIGNoYW5nZSwgOTAwLCAodikgPT4gYCR7diA+IDAgPyAnKycgOiAnJ30ke0Zvcm1hdC5udW1iZXIoTWF0aC5yb3VuZCh2KSl9YCk7CiAgICB9CiAgICBjb25zdCBkZWx0YVRleHQgPSBjaGFuZ2VQY3QgPT09IG51bGwgfHwgY2hhbmdlUGN0ID09PSB1bmRlZmluZWQgPyAn4oCUJyA6IEZvcm1hdC5wY3QoY2hhbmdlUGN0KTsKICAgIHRpbGUuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBkZWx0YVRleHQsIGBzdGF0LWRlbHRhICR7Rm9ybWF0LmRlbHRhQ2xhc3MoY2hhbmdlUGN0KX1gKSk7CiAgICByZXR1cm4gdGlsZTsKICB9CgogIC8qKiAiTmV3IEZvbGxvd2VycyIgdGlsZTogZm9sbG93ZXJzIGdhaW5lZCB3aXRoaW4gdGhlIGN1cnJlbnRseSBzZWxlY3RlZCBkYXRlIHJhbmdlIOKAlCBzaG93cyAiTm8gZm9sbG93ZXIgdXBkYXRlIiByYXRoZXIgdGhhbiAwIHdoZW4gbm90aGluZyBpcyBjb21wdXRhYmxlIGZvciB0aGUgcmFuZ2UgKHBlciBzcGVjKSwgd2hpY2ggaXMgZGlmZmVyZW50IGZyb20gYSBnZW51aW5lIHplcm8uICovCiAgZnVuY3Rpb24gbmV3Rm9sbG93ZXJzVGlsZShuZXdGb2xsb3dlcnMpIHsKICAgIGNvbnN0IHRpbGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRpbGUuY2xhc3NOYW1lID0gJ3N0YXQtdGlsZSc7CiAgICBjb25zdCB2YWx1ZUVsID0gdGV4dEVsKCdkaXYnLCAnJywgJ3N0YXQtdmFsdWUnKTsKICAgIHRpbGUuYXBwZW5kKHN0YXRMYWJlbEVsKCdOZXcgRm9sbG93ZXJzJywgJ3VzZXItcGx1cycsICd2MScpLCB2YWx1ZUVsKTsKICAgIGlmIChuZXdGb2xsb3dlcnMgPT09IG51bGwgfHwgbmV3Rm9sbG93ZXJzID09PSB1bmRlZmluZWQpIHsKICAgICAgdmFsdWVFbC50ZXh0Q29udGVudCA9ICdObyBmb2xsb3dlciB1cGRhdGUnOwogICAgICB2YWx1ZUVsLmNsYXNzTGlzdC5hZGQoJ3N0YXQtdmFsdWUtbXV0ZWQnKTsKICAgICAgcmVxdWVzdEFuaW1hdGlvbkZyYW1lKCgpID0+IGZpdFN0YXRWYWx1ZSh2YWx1ZUVsKSk7CiAgICB9IGVsc2UgewogICAgICBhbmltYXRlQ291bnQodmFsdWVFbCwgMCwgbmV3Rm9sbG93ZXJzLCA5MDAsICh2KSA9PiBgJHt2ID4gMCA/ICcrJyA6ICcnfSR7Rm9ybWF0Lm51bWJlcihNYXRoLnJvdW5kKHYpKX1gKTsKICAgIH0KICAgIHJldHVybiB0aWxlOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyS3BpcyhzdW1tYXJ5LCBmb2xsb3dlcnMpIHsKICAgIGNvbnN0IGdyaWQgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgna3BpR3JpZCcpOwogICAgaWYgKCFncmlkKSByZXR1cm47CiAgICBncmlkLmlubmVySFRNTCA9ICcnOwoKICAgIGdyaWQuYXBwZW5kQ2hpbGQoc3RhdFRpbGUoJ0hpZ2hlc3QgVmFsdWUnLCBzdW1tYXJ5LmhpZ2hlc3QsICh2KSA9PiBmb3JtYXRNZXRyaWNWYWx1ZShtZXRyaWMsIHYpLCAndHJlbmRpbmctdXAnLCAndjEnKSk7CiAgICBncmlkLmFwcGVuZENoaWxkKHN0YXRUaWxlKCdBdmVyYWdlIFZhbHVlJywgc3VtbWFyeS5hdmVyYWdlLCAodikgPT4gZm9ybWF0TWV0cmljVmFsdWUobWV0cmljLCB2KSwgJ2Jhci1jaGFydC0yJywgJ3Y0JykpOwogICAgZ3JpZC5hcHBlbmRDaGlsZChzdGF0VGlsZSgnVG90YWwgVmFsdWUnLCBzdW1tYXJ5LnRvdGFsLCAodikgPT4gZm9ybWF0TWV0cmljVmFsdWUobWV0cmljLCB2KSwgJ2xheWVycycsICd2NScpKTsKICAgIGdyaWQuYXBwZW5kQ2hpbGQoc3RhdFRpbGUoJ051bWJlciBvZiBQb3N0cycsIHN1bW1hcnkucG9zdENvdW50LCAodikgPT4gRm9ybWF0Lm51bWJlcih2KSwgJ2ZpbGUtdGV4dCcsICd2NicpKTsKICAgIGdyaWQuYXBwZW5kQ2hpbGQoc3RhdFRpbGUoJ0N1cnJlbnQgRm9sbG93ZXJzJywgZm9sbG93ZXJzLmN1cnJlbnRGb2xsb3dlcnMsICh2KSA9PiAodiA9PT0gbnVsbCB8fCB2ID09PSB1bmRlZmluZWQgPyAn4oCUJyA6IEZvcm1hdC5udW1iZXIodikpLCAndXNlcnMnLCAndjInKSk7CiAgICBncmlkLmFwcGVuZENoaWxkKGZvbGxvd2Vyc0dyb3d0aFRpbGUoZm9sbG93ZXJzLmZvbGxvd2Vyc0NoYW5nZSwgZm9sbG93ZXJzLmZvbGxvd2Vyc0NoYW5nZVBjdCkpOwogICAgZ3JpZC5hcHBlbmRDaGlsZChuZXdGb2xsb3dlcnNUaWxlKGZvbGxvd2Vycy5uZXdGb2xsb3dlcnMpKTsKICAgIGdyaWQuYXBwZW5kQ2hpbGQoYmVzdFBvc3RzVGlsZSgnQmVzdCBQZXJmb3JtaW5nIFBvc3QnLCBzdW1tYXJ5LmJlc3RQb3N0cywgZm9sbG93ZXJzLmN1cnJlbnRGb2xsb3dlcnMpKTsKICB9CgoKICAvKiogU3dhcHMgYSBjaGFydCBjYXJkJ3MgY2FudmFzIGZvciBhbiBlbXB0eS1zdGF0ZSBtZXNzYWdlLCBvciByZXN0b3JlcyB0aGUgY2FudmFzIOKAlCBzaW5jZQogICAgICByZS1yZW5kZXJpbmcgYSBDaGFydC5qcyBpbnN0YW5jZSBuZWVkcyBhIGxpdmUgPGNhbnZhcz4sIG5vdCB3aGF0ZXZlciB0aGUgbGFzdCByZW5kZXIgbGVmdCB0aGVyZS4gKi8KICBmdW5jdGlvbiBjaGFydE9yRW1wdHkod3JhcElkLCBjYW52YXNJZCwgaGFzRGF0YSwgZW1wdHlNZXNzYWdlLCByZW5kZXJGbikgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKHdyYXBJZCk7CiAgICBpZiAoIXdyYXApIHJldHVybjsKICAgIENoYXJ0cy5kZXN0cm95KGNhbnZhc0lkKTsKICAgIGlmICghaGFzRGF0YSkgewogICAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgICB3cmFwLmFwcGVuZENoaWxkKGVtcHR5U3RhdGUoeyBpY29uOiAnYmFyLWNoYXJ0LTMnLCBtZXNzYWdlOiBlbXB0eU1lc3NhZ2UgfSkpOwogICAgICByZXR1cm47CiAgICB9CiAgICB3cmFwLmlubmVySFRNTCA9IGA8Y2FudmFzIGlkPSIke2NhbnZhc0lkfSI+PC9jYW52YXM+YDsKICAgIHJlbmRlckZuKCk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiByZW5kZXJUcmVuZChmaWx0ZXJzKSB7CiAgICBjb25zdCBwbGF0Zm9ybU9wdGlvbnMgPSAod2luZG93Ll9fZmlsdGVyT3B0aW9uc0NhY2hlIHx8IHsgcGxhdGZvcm1zOiBbXSB9KS5wbGF0Zm9ybXM7CiAgICBjb25zdCBtTGFiZWwgPSBtZXRyaWNMYWJlbChtZXRyaWMpOwogICAgY29uc3QgcGxhdGZvcm1zVG9GZXRjaCA9IGZpbHRlcnMucGxhdGZvcm0gPT09ICdhbGwnID8gcGxhdGZvcm1PcHRpb25zLm1hcCgocCkgPT4gcC5pZCkgOiBbZmlsdGVycy5wbGF0Zm9ybV07CiAgICBjb25zdCB0cmVuZFJlc3BvbnNlcyA9IGF3YWl0IFByb21pc2UuYWxsKAogICAgICBwbGF0Zm9ybXNUb0ZldGNoLm1hcCgocCkgPT4KICAgICAgICBBcGkudHJlbmQoeyBkYXRlRnJvbTogZmlsdGVycy5kYXRlRnJvbSwgZGF0ZVRvOiBmaWx0ZXJzLmRhdGVUbywgcGxhdGZvcm06IHAsIGNhbXBhaWduVHlwZTogZmlsdGVycy5jYW1wYWlnblR5cGUsIGNvbnRlbnRUeXBlOiBmaWx0ZXJzLmNvbnRlbnRUeXBlIH0pCiAgICAgICkKICAgICk7CiAgICBjb25zdCB3ZWVrU2V0ID0gbmV3IFNldCgpOwogICAgdHJlbmRSZXNwb25zZXMuZm9yRWFjaCgocm93cykgPT4gcm93cy5mb3JFYWNoKChyKSA9PiB3ZWVrU2V0LmFkZChyLnBlcmlvZCkpKTsKICAgIGNvbnN0IHdlZWtzID0gWy4uLndlZWtTZXRdLnNvcnQoKTsKICAgIGNvbnN0IHNlcmllcyA9IHBsYXRmb3Jtc1RvRmV0Y2gubWFwKChwLCBpKSA9PiB7CiAgICAgIGNvbnN0IG1ldGEgPSBwbGF0Zm9ybU9wdGlvbnMuZmluZCgocGwpID0+IHBsLmlkID09PSBwKSB8fCB7IGxhYmVsOiBwIH07CiAgICAgIGNvbnN0IGJ5V2VlayA9IE9iamVjdC5mcm9tRW50cmllcyh0cmVuZFJlc3BvbnNlc1tpXS5tYXAoKHIpID0+IFtyLnBlcmlvZCwgclttZXRyaWNdXSkpOwogICAgICByZXR1cm4geyBsYWJlbDogbWV0YS5sYWJlbCwgY29sb3I6IG1ldGEuY29sb3IsIGRhdGE6IHdlZWtzLm1hcCgodykgPT4gKGJ5V2Vla1t3XSA9PT0gdW5kZWZpbmVkID8gbnVsbCA6IGJ5V2Vla1t3XSkpIH07CiAgICB9KTsKCiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndHJlbmRDYXJkVGl0bGUnKS50ZXh0Q29udGVudCA9CiAgICAgIGZpbHRlcnMucGxhdGZvcm0gPT09ICdhbGwnID8gYFdlZWtseSAke21MYWJlbH0gYnkgUGxhdGZvcm1gIDogYCR7bUxhYmVsfSBUcmVuZGA7CgogICAgY2hhcnRPckVtcHR5KCd0cmVuZENoYXJ0V3JhcCcsICd0cmVuZENhbnZhcycsIHdlZWtzLmxlbmd0aCA+IDAsICdObyBkYXRhIGluIHRoaXMgcmFuZ2UgeWV0LicsICgpID0+IHsKICAgICAgQ2hhcnRzLnRyZW5kQ2hhcnQoJ3RyZW5kQ2FudmFzJywgeyBsYWJlbHM6IHdlZWtzLm1hcChGb3JtYXQuZGF0ZSksIHNlcmllcywgZm9ybWF0VmFsdWU6ICh2KSA9PiBmb3JtYXRNZXRyaWNWYWx1ZShtZXRyaWMsIHYpIH0pOwogICAgfSk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiByZW5kZXJCcmVha2Rvd24oZmlsdGVycykgewogICAgY29uc3QgbUxhYmVsID0gbWV0cmljTGFiZWwobWV0cmljKTsKICAgIGNvbnN0IHRpdGxlRWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYnJlYWtkb3duQ2FyZFRpdGxlJyk7CgogICAgaWYgKGZpbHRlcnMucGxhdGZvcm0gPT09ICdhbGwnKSB7CiAgICAgIHRpdGxlRWwudGV4dENvbnRlbnQgPSBgUGxhdGZvcm0gQ29tcGFyaXNvbiDigJQgJHttTGFiZWx9YDsKICAgICAgY29uc3QgYnJlYWtkb3duID0gYXdhaXQgQXBpLnBsYXRmb3JtQnJlYWtkb3duKGZpbHRlcnMpOwogICAgICBjb25zdCBzb3J0ZWQgPSBicmVha2Rvd24uZmlsdGVyKChwKSA9PiBwW21ldHJpY10gIT09IG51bGwgJiYgcFttZXRyaWNdICE9PSB1bmRlZmluZWQpLnNvcnQoKGEsIGIpID0+IGJbbWV0cmljXSAtIGFbbWV0cmljXSk7CiAgICAgIGNoYXJ0T3JFbXB0eSgnYnJlYWtkb3duQ2hhcnRXcmFwJywgJ2JyZWFrZG93bkNhbnZhcycsIHNvcnRlZC5sZW5ndGggPiAwLCAnTm8gZGF0YSBpbiB0aGlzIHJhbmdlIHlldC4nLCAoKSA9PiB7CiAgICAgICAgQ2hhcnRzLnBsYXRmb3JtQmFyQ2hhcnQoJ2JyZWFrZG93bkNhbnZhcycsIHsKICAgICAgICAgIGxhYmVsczogc29ydGVkLm1hcCgocCkgPT4gcC5sYWJlbCksCiAgICAgICAgICBkYXRhOiBzb3J0ZWQubWFwKChwKSA9PiBwW21ldHJpY10pLAogICAgICAgICAgY29sb3JzOiBzb3J0ZWQubWFwKChwKSA9PiBwLmNvbG9yKSwKICAgICAgICAgIGZvcm1hdFZhbHVlOiAodikgPT4gZm9ybWF0TWV0cmljVmFsdWUobWV0cmljLCB2KSwKICAgICAgICB9KTsKICAgICAgfSk7CiAgICB9IGVsc2UgewogICAgICB0aXRsZUVsLnRleHRDb250ZW50ID0gYENhbXBhaWduIFBlcmZvcm1hbmNlIOKAlCAke21MYWJlbH1gOwogICAgICBjb25zdCBjYW1wYWlnbnMgPSBhd2FpdCBBcGkuY2FtcGFpZ25CcmVha2Rvd24oZmlsdGVycyk7CiAgICAgIGNvbnN0IHdpdGhWYWx1ZSA9IGNhbXBhaWducy5maWx0ZXIoKGMpID0+IGNbbWV0cmljXSAhPT0gbnVsbCAmJiBjW21ldHJpY10gIT09IHVuZGVmaW5lZCAmJiBjW21ldHJpY10gPiAwKTsKICAgICAgY2hhcnRPckVtcHR5KCdicmVha2Rvd25DaGFydFdyYXAnLCAnYnJlYWtkb3duQ2FudmFzJywgd2l0aFZhbHVlLmxlbmd0aCA+IDAsICdObyBjYW1wYWlnbiBkYXRhIGluIHRoaXMgcmFuZ2UgeWV0LicsICgpID0+IHsKICAgICAgICBDaGFydHMucGllQ2hhcnQoJ2JyZWFrZG93bkNhbnZhcycsIHsKICAgICAgICAgIGxhYmVsczogd2l0aFZhbHVlLm1hcCgoYykgPT4gYy5jYW1wYWlnbl90eXBlKSwKICAgICAgICAgIGRhdGE6IHdpdGhWYWx1ZS5tYXAoKGMpID0+IGNbbWV0cmljXSksCiAgICAgICAgICBjb2xvcnM6IHdpdGhWYWx1ZS5tYXAoKF8sIGkpID0+IENoYXJ0cy5zZXJpZXNDb2xvcihpKSksCiAgICAgICAgICBmb3JtYXRWYWx1ZTogKHYpID0+IGZvcm1hdE1ldHJpY1ZhbHVlKG1ldHJpYywgdiksCiAgICAgICAgfSk7CiAgICAgIH0pOwogICAgfQogIH0KCiAgYXN5bmMgZnVuY3Rpb24gcmVuZGVyQ29udGVudFR5cGVCcmVha2Rvd24oZmlsdGVycykgewogICAgY29uc3QgbUxhYmVsID0gbWV0cmljTGFiZWwobWV0cmljKTsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdjb250ZW50VHlwZUNhcmRUaXRsZScpLnRleHRDb250ZW50ID0gYENvbnRlbnQgVHlwZSBQZXJmb3JtYW5jZSDigJQgJHttTGFiZWx9YDsKICAgIGNvbnN0IHJvd3MgPSBhd2FpdCBBcGkuY29udGVudFR5cGVCcmVha2Rvd24oZmlsdGVycyk7CiAgICBjb25zdCBzb3J0ZWQgPSByb3dzLmZpbHRlcigoYykgPT4gY1ttZXRyaWNdICE9PSBudWxsICYmIGNbbWV0cmljXSAhPT0gdW5kZWZpbmVkKS5zb3J0KChhLCBiKSA9PiBiW21ldHJpY10gLSBhW21ldHJpY10pOwogICAgY2hhcnRPckVtcHR5KCdjb250ZW50VHlwZUNoYXJ0V3JhcCcsICdjb250ZW50VHlwZUNhbnZhcycsIHNvcnRlZC5sZW5ndGggPiAwLCAnTm8gZGF0YSBpbiB0aGlzIHJhbmdlIHlldC4nLCAoKSA9PiB7CiAgICAgIENoYXJ0cy5wbGF0Zm9ybUJhckNoYXJ0KCdjb250ZW50VHlwZUNhbnZhcycsIHsKICAgICAgICBsYWJlbHM6IHNvcnRlZC5tYXAoKGMpID0+IGMuY29udGVudF90eXBlKSwKICAgICAgICBkYXRhOiBzb3J0ZWQubWFwKChjKSA9PiBjW21ldHJpY10pLAogICAgICAgIGNvbG9yczogc29ydGVkLm1hcCgoXywgaSkgPT4gQ2hhcnRzLnNlcmllc0NvbG9yKGkpKSwKICAgICAgICBmb3JtYXRWYWx1ZTogKHYpID0+IGZvcm1hdE1ldHJpY1ZhbHVlKG1ldHJpYywgdiksCiAgICAgIH0pOwogICAgfSk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiByZW5kZXJUb3BQb3N0cyhmaWx0ZXJzKSB7CiAgICBjb25zdCBwb3N0cyA9IGF3YWl0IEFwaS50b3BQb3N0cyh7IC4uLmZpbHRlcnMsIHNvcnRCeTogbWV0cmljLCBsaW1pdDogMTAgfSk7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3RvcFBvc3RzVGFibGUnKTsKICAgIGlmICghd3JhcCkgcmV0dXJuOwogICAgaWYgKCFwb3N0cy5sZW5ndGgpIHsKICAgICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgICAgd3JhcC5hcHBlbmRDaGlsZChlbXB0eVN0YXRlKHsKICAgICAgICBpY29uOiAndHJvcGh5JywKICAgICAgICB0aXRsZTogJ05vIHBvc3RzIGluIHRoaXMgcmFuZ2UgeWV0JywKICAgICAgICBtZXNzYWdlOiAnVXBsb2FkIGEgd2Vla2x5IGV4cG9ydCwgb3Igd2lkZW4gdGhlIGRhdGUgcmFuZ2UsIHRvIHNlZSB0b3AgcGVyZm9ybWVycyBoZXJlLicsCiAgICAgICAgYWN0aW9uTGFiZWw6ICdVcGxvYWQgZGF0YScsCiAgICAgICAgb25BY3Rpb246ICgpID0+IGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3IoJy50YWItYnRuW2RhdGEtdGFiPSJ1cGxvYWQiXScpPy5jbGljaygpLAogICAgICB9KSk7CiAgICAgIHJldHVybjsKICAgIH0KICAgIGNvbnN0IHBsYXRmb3JtT3B0aW9ucyA9ICh3aW5kb3cuX19maWx0ZXJPcHRpb25zQ2FjaGUgfHwgeyBwbGF0Zm9ybXM6IFtdIH0pLnBsYXRmb3JtczsKCiAgICBjb25zdCB0YWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICB0YWJsZS5jbGFzc05hbWUgPSAnZGF0YS10YWJsZSc7CiAgICBjb25zdCB0aGVhZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RoZWFkJyk7CiAgICBjb25zdCBoZWFkVHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgaGVhZFRyLmFwcGVuZCgKICAgICAgdGV4dEVsKCd0aCcsICdSYW5rJyksCiAgICAgIHRleHRFbCgndGgnLCAnRGF0ZScpLAogICAgICB0ZXh0RWwoJ3RoJywgJ1BsYXRmb3JtJyksCiAgICAgIHRleHRFbCgndGgnLCAnQ2FtcGFpZ24nKSwKICAgICAgdGV4dEVsKCd0aCcsICdDb250ZW50IFR5cGUnKSwKICAgICAgdGV4dEVsKCd0aCcsICdDYXB0aW9uJyksCiAgICAgIHRleHRFbCgndGgnLCBtZXRyaWNMYWJlbChtZXRyaWMpLCAnbnVtJykKICAgICk7CiAgICBoZWFkVHIuYXBwZW5kQ2hpbGQodGV4dEVsKCd0aCcsICcnKSk7CiAgICB0aGVhZC5hcHBlbmRDaGlsZChoZWFkVHIpOwoKICAgIGNvbnN0IHRib2R5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGJvZHknKTsKICAgIHBvc3RzLmZvckVhY2goKHAsIGkpID0+IHsKICAgICAgY29uc3QgdHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgICBjb25zdCBtZXRhID0gcGxhdGZvcm1PcHRpb25zLmZpbmQoKHBsKSA9PiBwbC5pZCA9PT0gcC5wbGF0Zm9ybSkgfHwgeyBsYWJlbDogcC5wbGF0Zm9ybSwgY29sb3I6ICcjOTk5JyB9OwogICAgICBjb25zdCBwbGF0Zm9ybVRkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgY29uc3QgcGlsbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgICAgcGlsbC5jbGFzc05hbWUgPSAncGxhdGZvcm0tcGlsbCc7CiAgICAgIGNvbnN0IGRvdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgICAgZG90LmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1kb3QnOwogICAgICBkb3Quc3R5bGUuYmFja2dyb3VuZCA9IG1ldGEuY29sb3I7CiAgICAgIHBpbGwuYXBwZW5kKGRvdCwgZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUobWV0YS5sYWJlbCkpOwogICAgICBwbGF0Zm9ybVRkLmFwcGVuZENoaWxkKHBpbGwpOwoKICAgICAgY29uc3QgY2FwdGlvbiA9IHAuY2FwdGlvbiB8fCAnKG5vIGNhcHRpb24pJzsKICAgICAgY29uc3QgdHJ1bmNhdGVkID0gY2FwdGlvbi5sZW5ndGggPiA2MCA/IGAke2NhcHRpb24uc2xpY2UoMCwgNjApfeKApmAgOiBjYXB0aW9uOwogICAgICBjb25zdCBjYXB0aW9uVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBpZiAocC5wb3N0aW5nX2xpbmspIHsKICAgICAgICBjb25zdCBsaW5rID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYScpOwogICAgICAgIGxpbmsuY2xhc3NOYW1lID0gJ2NhcHRpb24tbGluayc7CiAgICAgICAgbGluay5ocmVmID0gcC5wb3N0aW5nX2xpbms7CiAgICAgICAgbGluay50YXJnZXQgPSAnX2JsYW5rJzsKICAgICAgICBsaW5rLnJlbCA9ICdub29wZW5lciBub3JlZmVycmVyJzsKICAgICAgICBsaW5rLnRpdGxlID0gY2FwdGlvbjsKICAgICAgICBsaW5rLmFwcGVuZENoaWxkKGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKHRydW5jYXRlZCkpOwogICAgICAgIGNhcHRpb25UZC5hcHBlbmRDaGlsZChsaW5rKTsKICAgICAgfSBlbHNlIHsKICAgICAgICBjYXB0aW9uVGQuYXBwZW5kQ2hpbGQoZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUodHJ1bmNhdGVkKSk7CiAgICAgICAgY2FwdGlvblRkLnRpdGxlID0gY2FwdGlvbjsKICAgICAgfQoKICAgICAgdHIuYXBwZW5kKAogICAgICAgIHRleHRFbCgndGQnLCBgIyR7aSArIDF9YCksCiAgICAgICAgdGV4dEVsKCd0ZCcsIEZvcm1hdC5kYXRlKHAucHVibGlzaF9kYXRlKSksCiAgICAgICAgcGxhdGZvcm1UZCwKICAgICAgICB0ZXh0RWwoJ3RkJywgcC5jYW1wYWlnbl90eXBlIHx8ICfigJQnKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgcC5jb250ZW50X3R5cGUgfHwgJ+KAlCcpLAogICAgICAgIGNhcHRpb25UZCwKICAgICAgICB0ZXh0RWwoJ3RkJywgZm9ybWF0TWV0cmljVmFsdWUobWV0cmljLCBwLm1ldHJpY192YWx1ZSksICdudW0nKQogICAgICApOwoKICAgICAgY29uc3QgYWN0aW9uVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBjb25zdCB2aWV3QnRuID0gaWNvbkJ0bignYnRuJywgJ2V5ZScsICdWaWV3IERldGFpbHMnKTsKICAgICAgdmlld0J0bi5kaXNhYmxlZCA9ICFwLnJhd19yb3dfaWQ7CiAgICAgIHZpZXdCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBSZWNvcmRzLm9wZW5WaWV3KHAucmF3X3Jvd19pZCkpOwogICAgICBhY3Rpb25UZC5hcHBlbmRDaGlsZCh2aWV3QnRuKTsKICAgICAgdHIuYXBwZW5kQ2hpbGQoYWN0aW9uVGQpOwoKICAgICAgdGJvZHkuYXBwZW5kQ2hpbGQodHIpOwogICAgfSk7CiAgICB0YWJsZS5hcHBlbmQodGhlYWQsIHRib2R5KTsKICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CiAgICB3cmFwLmFwcGVuZENoaWxkKHRhYmxlKTsKICB9CgogIC8qKiBNZXRyaWMgKG9yIGFueSBmaWx0ZXIpIGNoYW5nZWQgYnV0IHRoZSBwbGF0Zm9ybSDigJQgYW5kIHRoZXJlZm9yZSB0aGUgYXZhaWxhYmxlIG1ldHJpYyBsaXN0IOKAlCBkaWRuJ3Q6IG5vIG5lZWQgdG8gcmUtZmV0Y2ggbWV0cmljLW9wdGlvbnMgb3IgcmVidWlsZCB0aGUgc2hlbGwsIGp1c3QgcmVmcmVzaCB0aGUgZGF0YS4gKi8KICBhc3luYyBmdW5jdGlvbiByZWZyZXNoRm9yTWV0cmljKCkgewogICAgY29uc3QgZmlsdGVycyA9IFN0YXRlLmdldEZpbHRlcnMoKTsKICAgIGNvbnN0IFtzdW1tYXJ5LCBmb2xsb3dlcnNdID0gYXdhaXQgUHJvbWlzZS5hbGwoWwogICAgICBBcGkubWV0cmljU3VtbWFyeSh7IC4uLmZpbHRlcnMsIG1ldHJpYyB9KSwKICAgICAgQXBpLmZvbGxvd2Vyc0twaXMoZmlsdGVycyksCiAgICBdKTsKICAgIHJlbmRlcktwaXMoc3VtbWFyeSwgZm9sbG93ZXJzKTsKICAgIGF3YWl0IFByb21pc2UuYWxsKFsKICAgICAgcmVuZGVyVHJlbmQoZmlsdGVycyksIHJlbmRlckJyZWFrZG93bihmaWx0ZXJzKSwgcmVuZGVyQ29udGVudFR5cGVCcmVha2Rvd24oZmlsdGVycyksIHJlbmRlclRvcFBvc3RzKGZpbHRlcnMpLAogICAgXSk7CiAgfQoKICBmdW5jdGlvbiBzaG93U2tlbGV0b25zKCkgewogICAgY29uc3Qga3BpR3JpZCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdrcGlHcmlkJyk7CiAgICBpZiAoa3BpR3JpZCkgeyBrcGlHcmlkLmlubmVySFRNTCA9ICcnOyBrcGlHcmlkLmFwcGVuZENoaWxkKHNrZWxldG9uU3RhdEdyaWQoOCkpOyB9CiAgICBbJ3RyZW5kQ2hhcnRXcmFwJywgJ2JyZWFrZG93bkNoYXJ0V3JhcCcsICdjb250ZW50VHlwZUNoYXJ0V3JhcCddLmZvckVhY2goKGlkKSA9PiB7CiAgICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZChpZCk7CiAgICAgIGlmICh3cmFwKSB7IHdyYXAuaW5uZXJIVE1MID0gJyc7IHdyYXAuYXBwZW5kQ2hpbGQoc2tlbGV0b25DaGFydCgpKTsgfQogICAgfSk7CiAgICBjb25zdCB0b3BQb3N0cyA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd0b3BQb3N0c1RhYmxlJyk7CiAgICBpZiAodG9wUG9zdHMpIHsgdG9wUG9zdHMuaW5uZXJIVE1MID0gJyc7IHRvcFBvc3RzLmFwcGVuZENoaWxkKHNrZWxldG9uUm93cyg2KSk7IH0KICB9CgogIGFzeW5jIGZ1bmN0aW9uIHJlbmRlcigpIHsKICAgIHJvb3QgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndmlldy1kYXNoYm9hcmQnKTsKICAgIGNvbnN0IGZpbHRlcnMgPSBTdGF0ZS5nZXRGaWx0ZXJzKCk7CiAgICBjb25zdCB7IG9wdGlvbnMgfSA9IGF3YWl0IEFwaS5tZXRyaWNPcHRpb25zKGZpbHRlcnMucGxhdGZvcm0pOwogICAgbWV0cmljT3B0aW9ucyA9IG9wdGlvbnM7CiAgICBpZiAoIW1ldHJpY09wdGlvbnMuc29tZSgobSkgPT4gbS5rZXkgPT09IG1ldHJpYykpIHsKICAgICAgbWV0cmljID0gbWV0cmljT3B0aW9ucy5sZW5ndGggPyBtZXRyaWNPcHRpb25zWzBdLmtleSA6ICd2aWV3cyc7CiAgICB9CiAgICBzaGVsbCgpOwogICAgc2hvd1NrZWxldG9ucygpOwogICAgYXdhaXQgcmVmcmVzaEZvck1ldHJpYygpOwogIH0KCiAgcmV0dXJuIHsgcmVuZGVyIH07Cn0pKCk7CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgRGF0YSBSZWNvcmRzIHRhYjogYSBDUk0tc3R5bGUsIHBsYXRmb3JtLWdyb3VwZWQgYnJvd3NlciBiYWNrZWQKICAgYnkgcG9zdHMvcG9zdF9tZXRyaWNzICh0aGUgc2FtZSBub3JtYWxpemVkIGRhdGEgdGhlIGRhc2hib2FyZCwKICAgY29tcGFyaXNvbnMsIGFuZCByZXBvcnRzIHJlYWQpIOKAlCAiQWxsIFBsYXRmb3JtcyIgc2hvd3MgYSBjb21tb24KICAgY3Jvc3MtcGxhdGZvcm0gc3VtbWFyeSwgYSBzcGVjaWZpYyBwbGF0Zm9ybSBzaG93cyBvbmx5IHRoYXQKICAgcGxhdGZvcm0ncyBjdXJhdGVkIG1ldHJpY3MuIEV2ZXJ5IGZpZWxkIG9mIGEgcmVjb3JkIChleGFjdGx5IGFzCiAgIGltcG9ydGVkKSBpcyBhbHdheXMgcmVhY2hhYmxlIHZpYSBWaWV3L0VkaXQgcmVnYXJkbGVzcyBvZiB0aGUKICAgdGFibGUncyBjdXJhdGlvbiwgd2hpY2ggcmVhZHMgdGhlIHJhd19yb3dzIG1pcnJvciBhbmQsIG9uIHNhdmUsCiAgIHJlLXN5bmNzIHBvc3RzL3Bvc3RfbWV0cmljcyBzbyBldmVyeSB2aWV3IHJlZmxlY3RzIHRoZSBjaGFuZ2UKICAgaW1tZWRpYXRlbHkuCiAgID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwpjb25zdCBSZWNvcmRzID0gKCgpID0+IHsKICBsZXQgcm9vdDsKICBsZXQgcGFnZSA9IDE7CiAgY29uc3QgcGFnZVNpemUgPSAyNTsKICBsZXQgc2VhcmNoVmFsdWUgPSAnJzsKICBsZXQgc2VhcmNoRGVib3VuY2UgPSBudWxsOwogIGxldCBtb2RhbFN0YXRlID0gbnVsbDsgLy8geyByZWNvcmQsIHZhbHVlczogWy4uLl0gfSDigJQgRWRpdCBtb2RhbCBvbmx5CiAgbGV0IGN1cnJlbnRSZXN1bHQgPSBudWxsOyAvLyBsYXN0LWxvYWRlZCBwYWdlLCBrZXB0IHNvIHNvcnRpbmcgY2FuIHJlLXJlbmRlciB3aXRob3V0IGEgbmV0d29yayByb3VuZC10cmlwCiAgbGV0IHNvcnRTdGF0ZSA9IHsga2V5OiBudWxsLCBkaXI6ICdhc2MnLCB0eXBlOiAnc3RyaW5nJyB9OwoKICAvKiogU29ydHMgYSBjb3B5IG9mIGByb3dzYCBieSBhIChwb3NzaWJseSBkb3R0ZWQsIGUuZy4gIm1ldHJpY3MucmVhY2giKSBrZXkgcGF0aC4gTnVsbHMgYWx3YXlzIHNvcnQgbGFzdCByZWdhcmRsZXNzIG9mIGRpcmVjdGlvbi4gKi8KICBmdW5jdGlvbiBzb3J0Um93cyhyb3dzLCBrZXksIGRpciwgdHlwZSkgewogICAgY29uc3QgZmFjdG9yID0gZGlyID09PSAnYXNjJyA/IDEgOiAtMTsKICAgIGNvbnN0IHJlYWQgPSAocm93KSA9PiBrZXkuc3BsaXQoJy4nKS5yZWR1Y2UoKG8sIGspID0+IChvID09PSBudWxsIHx8IG8gPT09IHVuZGVmaW5lZCA/IHVuZGVmaW5lZCA6IG9ba10pLCByb3cpOwogICAgcmV0dXJuIFsuLi5yb3dzXS5zb3J0KChhLCBiKSA9PiB7CiAgICAgIGNvbnN0IGF2ID0gcmVhZChhKTsKICAgICAgY29uc3QgYnYgPSByZWFkKGIpOwogICAgICBjb25zdCBhTWlzc2luZyA9IGF2ID09PSBudWxsIHx8IGF2ID09PSB1bmRlZmluZWQgfHwgYXYgPT09ICcnOwogICAgICBjb25zdCBiTWlzc2luZyA9IGJ2ID09PSBudWxsIHx8IGJ2ID09PSB1bmRlZmluZWQgfHwgYnYgPT09ICcnOwogICAgICBpZiAoYU1pc3NpbmcgJiYgYk1pc3NpbmcpIHJldHVybiAwOwogICAgICBpZiAoYU1pc3NpbmcpIHJldHVybiAxOwogICAgICBpZiAoYk1pc3NpbmcpIHJldHVybiAtMTsKICAgICAgaWYgKHR5cGUgPT09ICdudW1iZXInKSByZXR1cm4gKGF2IC0gYnYpICogZmFjdG9yOwogICAgICByZXR1cm4gU3RyaW5nKGF2KS5sb2NhbGVDb21wYXJlKFN0cmluZyhidikpICogZmFjdG9yOwogICAgfSk7CiAgfQoKICAvKiogQSA8dGg+IHRoYXQgdG9nZ2xlcyBhc2NlbmRpbmcvZGVzY2VuZGluZyBvbiBjbGljayBhbmQgc2hvd3MgYW4gYXJyb3cgb24gd2hpY2hldmVyIGNvbHVtbiBpcyBhY3RpdmUg4oCUIHNvcnRzIHRoZSBhbHJlYWR5LWxvYWRlZCBwYWdlIGluc3RhbnRseSwgbm8gcmVsb2FkLiAqLwogIGZ1bmN0aW9uIHNvcnRhYmxlSGVhZGVyKGxhYmVsLCBrZXksIHR5cGUpIHsKICAgIGNvbnN0IHRoID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGgnKTsKICAgIGlmICh0eXBlID09PSAnbnVtYmVyJykgdGguY2xhc3NOYW1lID0gJ251bSc7CiAgICB0aC5jbGFzc0xpc3QuYWRkKCdzb3J0YWJsZS10aCcpOwogICAgY29uc3QgaXNBY3RpdmUgPSBzb3J0U3RhdGUua2V5ID09PSBrZXk7CiAgICB0aC5hcHBlbmRDaGlsZChkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShsYWJlbCkpOwogICAgdGguYXBwZW5kQ2hpbGQodGV4dEVsKCdzcGFuJywgaXNBY3RpdmUgPyAoc29ydFN0YXRlLmRpciA9PT0gJ2FzYycgPyAnIOKGkScgOiAnIOKGkycpIDogJyDihpUnLCAnc29ydC1hcnJvdycpKTsKICAgIHRoLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gewogICAgICBzb3J0U3RhdGUgPSB7IGtleSwgZGlyOiBzb3J0U3RhdGUua2V5ID09PSBrZXkgJiYgc29ydFN0YXRlLmRpciA9PT0gJ2FzYycgPyAnZGVzYycgOiAnYXNjJywgdHlwZSB9OwogICAgICBpZiAoY3VycmVudFJlc3VsdCkgcmVuZGVyVGFibGUoY3VycmVudFJlc3VsdCk7CiAgICB9KTsKICAgIHJldHVybiB0aDsKICB9CgogIGZ1bmN0aW9uIHBsYXRmb3JtTWV0YSgpIHsKICAgIHJldHVybiAod2luZG93Ll9fZmlsdGVyT3B0aW9uc0NhY2hlIHx8IHsgcGxhdGZvcm1zOiBbXSB9KS5wbGF0Zm9ybXM7CiAgfQoKICBmdW5jdGlvbiBwbGF0Zm9ybUxhYmVsKGlkKSB7CiAgICBjb25zdCBtID0gcGxhdGZvcm1NZXRhKCkuZmluZCgocCkgPT4gcC5pZCA9PT0gaWQpOwogICAgcmV0dXJuIG0gPyBtLmxhYmVsIDogaWQ7CiAgfQoKICBmdW5jdGlvbiBzaGVsbCgpIHsKICAgIHJvb3QuaW5uZXJIVE1MID0gJyc7CiAgICByb290LmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgJ0RhdGEgUmVjb3JkcycsICdzZWN0aW9uLXRpdGxlJykpOwogICAgcm9vdC5hcHBlbmRDaGlsZCh0ZXh0RWwoCiAgICAgICdkaXYnLAogICAgICAnQnJvd3NlIGJ5IHBsYXRmb3JtIHRvIHNlZSBvbmx5IGl0cyBtZXRyaWNzLCBvciBzdGF5IG9uIEFsbCBQbGF0Zm9ybXMgZm9yIGEgY3Jvc3MtcGxhdGZvcm0gc3VtbWFyeS4gRXZlcnkgcmVjb3JkIHN0YXlzIGZ1bGx5IGVkaXRhYmxlIOKAlCBWaWV3IG9yIEVkaXQgYWx3YXlzIG9wZW5zIGV2ZXJ5IGZpZWxkIGltcG9ydGVkIGZyb20gdGhlIHNwcmVhZHNoZWV0LCBub3QganVzdCB3aGF04oCZcyBpbiB0aGUgdGFibGUuJywKICAgICAgJ211dGVkJwogICAgKSk7CgogICAgY29uc3QgdG9vbGJhciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdG9vbGJhci5jbGFzc05hbWUgPSAncmVjb3Jkcy10b29sYmFyJzsKICAgIGNvbnN0IHBpbGxzID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBwaWxscy5jbGFzc05hbWUgPSAncGxhdGZvcm0tZmlsdGVyLXBpbGxzJzsKICAgIHBpbGxzLmlkID0gJ3JlY29yZHNQbGF0Zm9ybVBpbGxzJzsKICAgIGNvbnN0IHNlYXJjaCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgc2VhcmNoLmNsYXNzTmFtZSA9ICdyZWNvcmRzLXNlYXJjaCc7CiAgICBjb25zdCBzZWFyY2hJbnB1dCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7CiAgICBzZWFyY2hJbnB1dC50eXBlID0gJ3NlYXJjaCc7CiAgICBzZWFyY2hJbnB1dC5wbGFjZWhvbGRlciA9ICdTZWFyY2ggY2FwdGlvbnMsIGNhbXBhaWducywgY29udGVudCB0eXBl4oCmJzsKICAgIHNlYXJjaElucHV0LmlkID0gJ3JlY29yZHNTZWFyY2hJbnB1dCc7CiAgICBzZWFyY2hJbnB1dC52YWx1ZSA9IHNlYXJjaFZhbHVlOwogICAgc2VhcmNoSW5wdXQuYWRkRXZlbnRMaXN0ZW5lcignaW5wdXQnLCAoKSA9PiB7CiAgICAgIGNsZWFyVGltZW91dChzZWFyY2hEZWJvdW5jZSk7CiAgICAgIHNlYXJjaERlYm91bmNlID0gc2V0VGltZW91dCgoKSA9PiB7CiAgICAgICAgc2VhcmNoVmFsdWUgPSBzZWFyY2hJbnB1dC52YWx1ZTsKICAgICAgICBwYWdlID0gMTsKICAgICAgICBsb2FkKCk7CiAgICAgIH0sIDMwMCk7CiAgICB9KTsKICAgIHNlYXJjaC5hcHBlbmRDaGlsZChzZWFyY2hJbnB1dCk7CiAgICB0b29sYmFyLmFwcGVuZChwaWxscywgc2VhcmNoKTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQodG9vbGJhcik7CgogICAgY29uc3QgZXhwb3J0Um93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBleHBvcnRSb3cuY2xhc3NOYW1lID0gJ2V4cG9ydC1idXR0b25zJzsKICAgIGNvbnN0IGV4cG9ydENzdkJ0biA9IGljb25CdG4oJ2J0bicsICdmaWxlLWRvd24nLCAnRXhwb3J0IENTVicpOwogICAgZXhwb3J0Q3N2QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gdHJpZ2dlclJlY29yZHNFeHBvcnQoJ2NzdicpKTsKICAgIGNvbnN0IGV4cG9ydFhsc3hCdG4gPSBpY29uQnRuKCdidG4nLCAnZmlsZS1zcHJlYWRzaGVldCcsICdFeHBvcnQgRXhjZWwnKTsKICAgIGV4cG9ydFhsc3hCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB0cmlnZ2VyUmVjb3Jkc0V4cG9ydCgneGxzeCcpKTsKICAgIGNvbnN0IGRlZHVwZUJ0biA9IGljb25CdG4oJ2J0bicsICdjb3B5LXgnLCAnUmVtb3ZlIGR1cGxpY2F0ZXMnKTsKICAgIGRlZHVwZUJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IGhhbmRsZURlZHVwZShkZWR1cGVCdG4pKTsKICAgIGV4cG9ydFJvdy5hcHBlbmQoZXhwb3J0Q3N2QnRuLCBleHBvcnRYbHN4QnRuLCBkZWR1cGVCdG4pOwogICAgcm9vdC5hcHBlbmRDaGlsZChleHBvcnRSb3cpOwoKICAgIGNvbnN0IGNhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgY29uc3QgdGFibGVXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0YWJsZVdyYXAuY2xhc3NOYW1lID0gJ3RhYmxlLXNjcm9sbCc7CiAgICB0YWJsZVdyYXAuaWQgPSAncmVjb3Jkc1RhYmxlV3JhcCc7CiAgICBjYXJkLmFwcGVuZENoaWxkKHRhYmxlV3JhcCk7CiAgICBjb25zdCBwYWdlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcGFnZXIuY2xhc3NOYW1lID0gJ3BhZ2luYXRpb24tcm93JzsKICAgIHBhZ2VyLmlkID0gJ3JlY29yZHNQYWdlcic7CiAgICBjYXJkLmFwcGVuZENoaWxkKHBhZ2VyKTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQoY2FyZCk7CgogICAgcmVuZGVyUGlsbHMoKTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlclBpbGxzKCkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdyZWNvcmRzUGxhdGZvcm1QaWxscycpOwogICAgaWYgKCF3cmFwKSByZXR1cm47CiAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgY29uc3QgY3VycmVudCA9IFN0YXRlLmdldEZpbHRlcnMoKS5wbGF0Zm9ybSB8fCAnYWxsJzsKICAgIGNvbnN0IG9wdGlvbnMgPSBbeyBpZDogJ2FsbCcsIGxhYmVsOiAnQWxsIFBsYXRmb3JtcycsIGNvbG9yOiBudWxsIH0sIC4uLnBsYXRmb3JtTWV0YSgpXTsKICAgIG9wdGlvbnMuZm9yRWFjaCgob3B0KSA9PiB7CiAgICAgIGNvbnN0IGJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgICBidG4udHlwZSA9ICdidXR0b24nOwogICAgICBidG4uY2xhc3NMaXN0LnRvZ2dsZSgnaXMtYWN0aXZlJywgY3VycmVudCA9PT0gb3B0LmlkKTsKICAgICAgaWYgKG9wdC5jb2xvcikgewogICAgICAgIGNvbnN0IGRvdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgICAgICBkb3QuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLWRvdCc7CiAgICAgICAgZG90LnN0eWxlLmJhY2tncm91bmQgPSBvcHQuY29sb3I7CiAgICAgICAgYnRuLmFwcGVuZENoaWxkKGRvdCk7CiAgICAgIH0KICAgICAgYnRuLmFwcGVuZENoaWxkKGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKG9wdC5sYWJlbCkpOwogICAgICBidG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7CiAgICAgICAgaWYgKGN1cnJlbnQgPT09IG9wdC5pZCkgcmV0dXJuOwogICAgICAgIGNvbnN0IGZpbHRlclNlbGVjdCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJQbGF0Zm9ybScpOwogICAgICAgIGlmIChmaWx0ZXJTZWxlY3QpIGZpbHRlclNlbGVjdC52YWx1ZSA9IG9wdC5pZDsKICAgICAgICBwYWdlID0gMTsKICAgICAgICBTdGF0ZS5zZXRGaWx0ZXJzKHsgcGxhdGZvcm06IG9wdC5pZCB9KTsKICAgICAgfSk7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQoYnRuKTsKICAgIH0pOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gbG9hZCgpIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncmVjb3Jkc1RhYmxlV3JhcCcpOwogICAgaWYgKHdyYXApIHsgd3JhcC5pbm5lckhUTUwgPSAnJzsgd3JhcC5hcHBlbmRDaGlsZChza2VsZXRvblJvd3MoOCkpOyB9CiAgICBjb25zdCBmaWx0ZXJzID0gU3RhdGUuZ2V0RmlsdGVycygpOwogICAgY29uc3QgcmVzdWx0ID0gYXdhaXQgQXBpLnJlY29yZHNUYWJsZSh7IC4uLmZpbHRlcnMsIHNlYXJjaDogc2VhcmNoVmFsdWUsIHBhZ2UsIHBhZ2VTaXplIH0pOwogICAgcmVuZGVyVGFibGUocmVzdWx0KTsKICAgIHJlbmRlclBhZ2VyKHJlc3VsdCk7CiAgfQoKICAvKiogUmVjb3JkcyBpcyBzZXJ2ZXItcGFnaW5hdGVkL3NlYXJjaGVkLCBzbyBpdHMgZXhwb3J0IGlzIGEgZGlyZWN0IG5hdmlnYXRpb24gdG8gYSBiYWNrZW5kIHJvdXRlIHRoYXQgcmV1c2VzIHRoZSBleGFjdCBzYW1lIGZpbHRlci1idWlsZGluZyB0aGUgbGlzdCBlbmRwb2ludCBkb2VzIOKAlCBleHBvcnRzIHRoZSBmdWxsIG1hdGNoaW5nIGRhdGFzZXQsIG5vdCBqdXN0IHRoZSBjdXJyZW50IHBhZ2UuICovCiAgZnVuY3Rpb24gdHJpZ2dlclJlY29yZHNFeHBvcnQoZm9ybWF0KSB7CiAgICBjb25zdCBmaWx0ZXJzID0gU3RhdGUuZ2V0RmlsdGVycygpOwogICAgY29uc3QgcGFyYW1zID0gbmV3IFVSTFNlYXJjaFBhcmFtcyh7IC4uLmZpbHRlcnMsIHNlYXJjaDogc2VhcmNoVmFsdWUsIGZvcm1hdCB9KTsKICAgIHdpbmRvdy5sb2NhdGlvbi5ocmVmID0gYC9hcGkvcmVjb3Jkcy9leHBvcnQ/JHtwYXJhbXMudG9TdHJpbmcoKX1gOwogIH0KCiAgZnVuY3Rpb24gY29sdW1uTGFiZWxzRm9yKHJlY29yZCkgewogICAgcmV0dXJuIHJlY29yZC5oZWFkZXJzICYmIHJlY29yZC5oZWFkZXJzLmxlbmd0aAogICAgICA/IHJlY29yZC5oZWFkZXJzLm1hcCgoaCkgPT4gKGggJiYgaC50cmltKCkgPyBoIDogJyh1bmxhYmVsZWQgY29sdW1uKScpKQogICAgICA6IHJlY29yZC52YWx1ZXMubWFwKChfLCBpKSA9PiBgQ29sdW1uICR7aSArIDF9YCk7CiAgfQoKICAvKiogR3JvdXBzIGEgcmF3IHJlY29yZCdzIGZpZWxkcyBieSB0aGUgcXVhbGlmaWVkIGhlYWRlcidzIHBsYXRmb3JtLWdyb3VwIHByZWZpeAogICAgICAoZS5nLiAiRkFDRUJPT0sg4oCUIFZpZXdzIiksIHNvIHRoZSBWaWV3L0VkaXQgcG9wdXAgcmVhZHMgYXMgc2VjdGlvbnMgaW5zdGVhZAogICAgICBvZiBvbmUgbG9uZyBmbGF0IGxpc3Qg4oCUIGZhbGxzIGJhY2sgdG8gYSBzaW5nbGUgIkRldGFpbHMiIHNlY3Rpb24gZm9yCiAgICAgIGlkZW50aWZpZXIgY29sdW1ucyBhbmQgZm9yIHRoZSBzaW1wbGUgKG9uZS1wbGF0Zm9ybS1wZXItcm93KSBmb3JtYXQuICovCiAgZnVuY3Rpb24gZ3JvdXBGaWVsZFJvd3MobGFiZWxzLCB2YWx1ZXMpIHsKICAgIGNvbnN0IGdyb3VwcyA9IFtdOwogICAgY29uc3QgaW5kZXggPSBuZXcgTWFwKCk7CiAgICBsYWJlbHMuZm9yRWFjaCgobGFiZWwsIGlkeCkgPT4gewogICAgICBjb25zdCBzZXBJZHggPSBsYWJlbC5pbmRleE9mKCcg4oCUICcpOwogICAgICBjb25zdCBncm91cE5hbWUgPSBzZXBJZHggPj0gMCA/IGxhYmVsLnNsaWNlKDAsIHNlcElkeCkgOiAnRGV0YWlscyc7CiAgICAgIGNvbnN0IGZpZWxkTGFiZWwgPSBzZXBJZHggPj0gMCA/IGxhYmVsLnNsaWNlKHNlcElkeCArIDMpIDogbGFiZWw7CiAgICAgIGlmICghaW5kZXguaGFzKGdyb3VwTmFtZSkpIHsKICAgICAgICBpbmRleC5zZXQoZ3JvdXBOYW1lLCB7IGdyb3VwOiBncm91cE5hbWUsIGZpZWxkczogW10gfSk7CiAgICAgICAgZ3JvdXBzLnB1c2goaW5kZXguZ2V0KGdyb3VwTmFtZSkpOwogICAgICB9CiAgICAgIGluZGV4LmdldChncm91cE5hbWUpLmZpZWxkcy5wdXNoKHsgaWR4LCBsYWJlbDogZmllbGRMYWJlbCB8fCBgQ29sdW1uICR7aWR4ICsgMX1gLCB2YWx1ZTogdmFsdWVzW2lkeF0gfSk7CiAgICB9KTsKICAgIHJldHVybiBncm91cHM7CiAgfQoKICBmdW5jdGlvbiBwbGF0Zm9ybUJhZGdlcyhpZHMpIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHdyYXAuc3R5bGUuZGlzcGxheSA9ICdmbGV4JzsKICAgIHdyYXAuc3R5bGUuZmxleFdyYXAgPSAnd3JhcCc7CiAgICB3cmFwLnN0eWxlLmdhcCA9ICc0cHgnOwogICAgaWYgKCFpZHMubGVuZ3RoKSByZXR1cm4gdGV4dEVsKCdzcGFuJywgJ+KAlCcsICdtdXRlZCcpOwogICAgY29uc3QgbWV0YSA9IHBsYXRmb3JtTWV0YSgpOwogICAgaWRzLmZvckVhY2goKGlkKSA9PiB7CiAgICAgIGNvbnN0IG0gPSBtZXRhLmZpbmQoKHApID0+IHAuaWQgPT09IGlkKSB8fCB7IGxhYmVsOiBpZCwgY29sb3I6ICcjOTk5JyB9OwogICAgICBjb25zdCBwaWxsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICBwaWxsLmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1waWxsJzsKICAgICAgY29uc3QgZG90ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICBkb3QuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLWRvdCc7CiAgICAgIGRvdC5zdHlsZS5iYWNrZ3JvdW5kID0gbS5jb2xvcjsKICAgICAgcGlsbC5hcHBlbmQoZG90LCBkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShtLmxhYmVsKSk7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQocGlsbCk7CiAgICB9KTsKICAgIHJldHVybiB3cmFwOwogIH0KCiAgZnVuY3Rpb24gc3RhdHVzUGlsbChzdGF0dXMpIHsKICAgIGNvbnN0IHNwYW4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICBzcGFuLmNsYXNzTmFtZSA9IGBzdGF0dXMtcGlsbCAke3N0YXR1c31gOwogICAgc3Bhbi50ZXh0Q29udGVudCA9IHN0YXR1cyA9PT0gJ2VkaXRlZCcgPyAnRWRpdGVkJyA6ICdPcmlnaW5hbCc7CiAgICByZXR1cm4gc3BhbjsKICB9CgogIGZ1bmN0aW9uIG1ldHJpY0NlbGwoa2V5LCB2YWx1ZSkgewogICAgaWYgKGtleSA9PT0gJ3Bvc3RpbmdfbGluaycpIHsKICAgICAgY29uc3QgdGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICB0ZC5jbGFzc05hbWUgPSAnbGluay1jZWxsJzsKICAgICAgaWYgKHZhbHVlKSB7CiAgICAgICAgY29uc3QgYSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2EnKTsKICAgICAgICBhLmhyZWYgPSB2YWx1ZTsKICAgICAgICBhLnRhcmdldCA9ICdfYmxhbmsnOwogICAgICAgIGEucmVsID0gJ25vb3BlbmVyIG5vcmVmZXJyZXInOwogICAgICAgIGEudGV4dENvbnRlbnQgPSAnT3BlbiDihpcnOwogICAgICAgIHRkLmFwcGVuZENoaWxkKGEpOwogICAgICB9IGVsc2UgewogICAgICAgIHRkLmFwcGVuZENoaWxkKGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKCfigJQnKSk7CiAgICAgIH0KICAgICAgcmV0dXJuIHRkOwogICAgfQogICAgY29uc3QgZGlzcGxheSA9IGtleSA9PT0gJ3dhdGNoX3RpbWVfc2Vjb25kcycgPyBGb3JtYXQuZHVyYXRpb24odmFsdWUpIDogRm9ybWF0Lm51bWJlcih2YWx1ZSk7CiAgICByZXR1cm4gdGV4dEVsKCd0ZCcsIGRpc3BsYXksICdudW0nKTsKICB9CgogIGZ1bmN0aW9uIGFjdGlvbkJ1dHRvbnMocm93LCBwbGF0Zm9ybSkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgd3JhcC5jbGFzc05hbWUgPSAncm93LWFjdGlvbnMnOwogICAgY29uc3Qgdmlld0J0biA9IGljb25CdG4oJ2J0bicsICdleWUnLCAnVmlldycpOwogICAgdmlld0J0bi5kaXNhYmxlZCA9ICFyb3cucmF3Um93SWQ7CiAgICB2aWV3QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gb3BlblZpZXcocm93LnJhd1Jvd0lkKSk7CiAgICBjb25zdCBlZGl0QnRuID0gaWNvbkJ0bignYnRuJywgJ3BlbmNpbCcsICdFZGl0Jyk7CiAgICBlZGl0QnRuLmRpc2FibGVkID0gIXJvdy5yYXdSb3dJZDsKICAgIGVkaXRCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBvcGVuRWRpdG9yKHJvdy5yYXdSb3dJZCkpOwogICAgY29uc3QgZGVsZXRlQnRuID0gaWNvbkJ0bignYnRuIGRhbmdlcicsICd0cmFzaC0yJywgJ0RlbGV0ZScpOwogICAgZGVsZXRlQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gaGFuZGxlRGVsZXRlKHJvdywgcGxhdGZvcm0pKTsKICAgIHdyYXAuYXBwZW5kKHZpZXdCdG4sIGVkaXRCdG4sIGRlbGV0ZUJ0bik7CiAgICByZXR1cm4gd3JhcDsKICB9CgogIGZ1bmN0aW9uIGNhcHRpb25DZWxsKGNhcHRpb24pIHsKICAgIGNvbnN0IHRleHQgPSBjYXB0aW9uIHx8ICcobm8gY2FwdGlvbiknOwogICAgcmV0dXJuIHRleHRFbCgndGQnLCB0ZXh0Lmxlbmd0aCA+IDcwID8gYCR7dGV4dC5zbGljZSgwLCA3MCl94oCmYCA6IHRleHQpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyU3VtbWFyeVRhYmxlKHJlc3VsdCkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdyZWNvcmRzVGFibGVXcmFwJyk7CiAgICBpZiAoIXJlc3VsdC5yb3dzLmxlbmd0aCkgewogICAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgICB3cmFwLmFwcGVuZENoaWxkKGVtcHR5U3RhdGUoewogICAgICAgIGljb246ICdkYXRhYmFzZScsCiAgICAgICAgdGl0bGU6ICdObyByZWNvcmRzIG1hdGNoIHRoZXNlIGZpbHRlcnMgeWV0JywKICAgICAgICBtZXNzYWdlOiAnVXBsb2FkIGEgd2Vla2x5IGV4cG9ydCwgb3Igd2lkZW4gdGhlIGRhdGUgcmFuZ2UsIHRvIHNlZSByZWNvcmRzIGhlcmUuJywKICAgICAgICBhY3Rpb25MYWJlbDogJ1VwbG9hZCBkYXRhJywKICAgICAgICBvbkFjdGlvbjogKCkgPT4gZG9jdW1lbnQucXVlcnlTZWxlY3RvcignLnRhYi1idG5bZGF0YS10YWI9InVwbG9hZCJdJyk/LmNsaWNrKCksCiAgICAgIH0pKTsKICAgICAgcmV0dXJuOwogICAgfQogICAgY29uc3QgdGFibGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0YWJsZScpOwogICAgdGFibGUuY2xhc3NOYW1lID0gJ2RhdGEtdGFibGUnOwogICAgY29uc3QgdGhlYWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0aGVhZCcpOwogICAgY29uc3QgaGVhZFRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgIGhlYWRUci5hcHBlbmQoCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdEYXRlJywgJ3B1Ymxpc2hEYXRlJywgJ3N0cmluZycpLAogICAgICBzb3J0YWJsZUhlYWRlcignUGxhdGZvcm1zJywgJ3BsYXRmb3JtSWRzLjAnLCAnc3RyaW5nJyksCiAgICAgIHRleHRFbCgndGgnLCAnQ2FwdGlvbicpLAogICAgICB0ZXh0RWwoJ3RoJywgJ0NhbXBhaWduJyksCiAgICAgIHRleHRFbCgndGgnLCAnQ29udGVudCBUeXBlJyksCiAgICAgIHRleHRFbCgndGgnLCAnU3RhdHVzJyksCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdMYXN0IFVwZGF0ZWQnLCAndXBkYXRlZEF0JywgJ3N0cmluZycpLAogICAgICB0ZXh0RWwoJ3RoJywgJ0FjdGlvbnMnKQogICAgKTsKICAgIHRoZWFkLmFwcGVuZENoaWxkKGhlYWRUcik7CiAgICBjb25zdCB0Ym9keSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3Rib2R5Jyk7CiAgICBjb25zdCByb3dzID0gc29ydFN0YXRlLmtleSA/IHNvcnRSb3dzKHJlc3VsdC5yb3dzLCBzb3J0U3RhdGUua2V5LCBzb3J0U3RhdGUuZGlyLCBzb3J0U3RhdGUudHlwZSkgOiByZXN1bHQucm93czsKICAgIHJvd3MuZm9yRWFjaCgocikgPT4gewogICAgICBjb25zdCB0ciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAgIGNvbnN0IHBsYXRmb3Jtc1RkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgcGxhdGZvcm1zVGQuYXBwZW5kQ2hpbGQocGxhdGZvcm1CYWRnZXMoci5wbGF0Zm9ybUlkcykpOwogICAgICBjb25zdCBzdGF0dXNUZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIHN0YXR1c1RkLmFwcGVuZENoaWxkKHN0YXR1c1BpbGwoci5zdGF0dXMpKTsKICAgICAgY29uc3QgYWN0aW9uc1RkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgYWN0aW9uc1RkLmFwcGVuZENoaWxkKGFjdGlvbkJ1dHRvbnMociwgJ2FsbCcpKTsKICAgICAgdHIuYXBwZW5kKAogICAgICAgIHRleHRFbCgndGQnLCBGb3JtYXQuZGF0ZShyLnB1Ymxpc2hEYXRlKSksCiAgICAgICAgcGxhdGZvcm1zVGQsCiAgICAgICAgY2FwdGlvbkNlbGwoci5jYXB0aW9uKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgci5jYW1wYWlnblR5cGUgfHwgJ+KAlCcpLAogICAgICAgIHRleHRFbCgndGQnLCByLmNvbnRlbnRUeXBlIHx8ICfigJQnKSwKICAgICAgICBzdGF0dXNUZCwKICAgICAgICB0ZXh0RWwoJ3RkJywgci51cGRhdGVkQXQpLAogICAgICAgIGFjdGlvbnNUZAogICAgICApOwogICAgICB0Ym9keS5hcHBlbmRDaGlsZCh0cik7CiAgICB9KTsKICAgIHRhYmxlLmFwcGVuZCh0aGVhZCwgdGJvZHkpOwogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGFibGUpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyUGxhdGZvcm1UYWJsZShyZXN1bHQpIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncmVjb3Jkc1RhYmxlV3JhcCcpOwogICAgaWYgKCFyZXN1bHQucm93cy5sZW5ndGgpIHsKICAgICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgICAgd3JhcC5hcHBlbmRDaGlsZChlbXB0eVN0YXRlKHsKICAgICAgICBpY29uOiAnZGF0YWJhc2UnLAogICAgICAgIHRpdGxlOiBgTm8gJHtwbGF0Zm9ybUxhYmVsKHJlc3VsdC5wbGF0Zm9ybSl9IHJlY29yZHMgbWF0Y2ggdGhlc2UgZmlsdGVycyB5ZXRgLAogICAgICAgIG1lc3NhZ2U6ICdUcnkgYSBkaWZmZXJlbnQgcGxhdGZvcm0sIG9yIHdpZGVuIHRoZSBkYXRlIHJhbmdlLicsCiAgICAgIH0pKTsKICAgICAgcmV0dXJuOwogICAgfQogICAgY29uc3QgdGFibGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0YWJsZScpOwogICAgdGFibGUuY2xhc3NOYW1lID0gJ2RhdGEtdGFibGUnOwogICAgY29uc3QgdGhlYWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0aGVhZCcpOwogICAgY29uc3QgaGVhZFRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgIGhlYWRUci5hcHBlbmQoc29ydGFibGVIZWFkZXIoJ0RhdGUnLCAncHVibGlzaERhdGUnLCAnc3RyaW5nJyksIHRleHRFbCgndGgnLCAnQ2FwdGlvbicpLCB0ZXh0RWwoJ3RoJywgJ0NhbXBhaWduJyksIHRleHRFbCgndGgnLCAnQ29udGVudCBUeXBlJykpOwogICAgcmVzdWx0LmNvbHVtbnMuZm9yRWFjaCgoYykgPT4gewogICAgICBpZiAoYy5rZXkgPT09ICdwb3N0aW5nX2xpbmsnKSB7CiAgICAgICAgaGVhZFRyLmFwcGVuZENoaWxkKHRleHRFbCgndGgnLCBjLmxhYmVsKSk7CiAgICAgIH0gZWxzZSB7CiAgICAgICAgaGVhZFRyLmFwcGVuZENoaWxkKHNvcnRhYmxlSGVhZGVyKGMubGFiZWwsIGBtZXRyaWNzLiR7Yy5rZXl9YCwgJ251bWJlcicpKTsKICAgICAgfQogICAgfSk7CiAgICBoZWFkVHIuYXBwZW5kKHRleHRFbCgndGgnLCAnU3RhdHVzJyksIHRleHRFbCgndGgnLCAnQWN0aW9ucycpKTsKICAgIHRoZWFkLmFwcGVuZENoaWxkKGhlYWRUcik7CiAgICBjb25zdCB0Ym9keSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3Rib2R5Jyk7CiAgICBjb25zdCByb3dzID0gc29ydFN0YXRlLmtleSA/IHNvcnRSb3dzKHJlc3VsdC5yb3dzLCBzb3J0U3RhdGUua2V5LCBzb3J0U3RhdGUuZGlyLCBzb3J0U3RhdGUudHlwZSkgOiByZXN1bHQucm93czsKICAgIHJvd3MuZm9yRWFjaCgocikgPT4gewogICAgICBjb25zdCB0ciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAgIHRyLmFwcGVuZCh0ZXh0RWwoJ3RkJywgRm9ybWF0LmRhdGUoci5wdWJsaXNoRGF0ZSkpLCBjYXB0aW9uQ2VsbChyLmNhcHRpb24pLCB0ZXh0RWwoJ3RkJywgci5jYW1wYWlnblR5cGUgfHwgJ+KAlCcpLCB0ZXh0RWwoJ3RkJywgci5jb250ZW50VHlwZSB8fCAn4oCUJykpOwogICAgICByZXN1bHQuY29sdW1ucy5mb3JFYWNoKChjKSA9PiB0ci5hcHBlbmRDaGlsZChtZXRyaWNDZWxsKGMua2V5LCByLm1ldHJpY3NbYy5rZXldKSkpOwogICAgICBjb25zdCBzdGF0dXNUZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIHN0YXR1c1RkLmFwcGVuZENoaWxkKHN0YXR1c1BpbGwoci5zdGF0dXMpKTsKICAgICAgdHIuYXBwZW5kQ2hpbGQoc3RhdHVzVGQpOwogICAgICBjb25zdCBhY3Rpb25zVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBhY3Rpb25zVGQuYXBwZW5kQ2hpbGQoYWN0aW9uQnV0dG9ucyhyLCByZXN1bHQucGxhdGZvcm0pKTsKICAgICAgdHIuYXBwZW5kQ2hpbGQoYWN0aW9uc1RkKTsKICAgICAgdGJvZHkuYXBwZW5kQ2hpbGQodHIpOwogICAgfSk7CiAgICB0YWJsZS5hcHBlbmQodGhlYWQsIHRib2R5KTsKICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CiAgICB3cmFwLmFwcGVuZENoaWxkKHRhYmxlKTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlclRhYmxlKHJlc3VsdCkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdyZWNvcmRzVGFibGVXcmFwJyk7CiAgICBpZiAoIXdyYXApIHJldHVybjsKICAgIGN1cnJlbnRSZXN1bHQgPSByZXN1bHQ7CiAgICBpZiAocmVzdWx0LnBsYXRmb3JtID09PSAnYWxsJykgcmVuZGVyU3VtbWFyeVRhYmxlKHJlc3VsdCk7CiAgICBlbHNlIHJlbmRlclBsYXRmb3JtVGFibGUocmVzdWx0KTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlclBhZ2VyKHJlc3VsdCkgewogICAgY29uc3QgcGFnZXIgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncmVjb3Jkc1BhZ2VyJyk7CiAgICBpZiAoIXBhZ2VyKSByZXR1cm47CiAgICBwYWdlci5pbm5lckhUTUwgPSAnJzsKICAgIGNvbnN0IHRvdGFsUGFnZXMgPSBNYXRoLm1heCgxLCBNYXRoLmNlaWwocmVzdWx0LnRvdGFsIC8gcmVzdWx0LnBhZ2VTaXplKSk7CiAgICBjb25zdCBwcmV2QnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBwcmV2QnRuLmNsYXNzTmFtZSA9ICdidG4nOwogICAgcHJldkJ0bi50ZXh0Q29udGVudCA9ICdQcmV2aW91cyc7CiAgICBwcmV2QnRuLmRpc2FibGVkID0gcmVzdWx0LnBhZ2UgPD0gMTsKICAgIHByZXZCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7IHBhZ2UgLT0gMTsgbG9hZCgpOyB9KTsKICAgIGNvbnN0IG5leHRCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIG5leHRCdG4uY2xhc3NOYW1lID0gJ2J0bic7CiAgICBuZXh0QnRuLnRleHRDb250ZW50ID0gJ05leHQnOwogICAgbmV4dEJ0bi5kaXNhYmxlZCA9IHJlc3VsdC5wYWdlID49IHRvdGFsUGFnZXM7CiAgICBuZXh0QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4geyBwYWdlICs9IDE7IGxvYWQoKTsgfSk7CiAgICBwYWdlci5hcHBlbmQocHJldkJ0biwgdGV4dEVsKCdzcGFuJywgYFBhZ2UgJHtyZXN1bHQucGFnZX0gb2YgJHt0b3RhbFBhZ2VzfSDigJQgJHtyZXN1bHQudG90YWx9IHJlY29yZChzKWApLCBuZXh0QnRuKTsKICB9CgogIC8qKiAiUmVtb3ZlIGR1cGxpY2F0ZXMiOiBjb2xsYXBzZSByZWNvcmRzIHRoYXQgc2hhcmUgYSBwdWJsaXNoIGRhdGUgKyBjYXB0aW9uCiAgICAgICh0aGUgY29waWVzIGxlZnQgYmVoaW5kIHdoZW4gYSB3ZWVrbHkgc2hlZXQgaXMgcmUtdXBsb2FkZWQgYW5kIGEgcG9zdCdzCiAgICAgIG51bWJlcnMgY2hhbmdlZCksIGtlZXBpbmcgdGhlIG1vc3QgcmVjZW50bHkgaW1wb3J0ZWQgb25lLiBQcmV2aWV3cyB0aGUKICAgICAgY291bnQgZmlyc3QsIHRoZW4gZGVsZXRlcyBvbiBjb25maXJtIOKAlCB0aGUgb3JpZ2luYWxzIHN0YXkgaW4gVXBsb2FkCiAgICAgIEhpc3RvcnkncyByYXctcm93IHZpZXdlciBlaXRoZXIgd2F5LiAqLwogIGFzeW5jIGZ1bmN0aW9uIGhhbmRsZURlZHVwZShidG4pIHsKICAgIGlmIChidG4gJiYgYnRuLmRpc2FibGVkKSByZXR1cm47CiAgICB0cnkgewogICAgICBpZiAoYnRuKSBidG4uZGlzYWJsZWQgPSB0cnVlOwogICAgICBjb25zdCBwcmV2aWV3ID0gYXdhaXQgQXBpLmR1cGxpY2F0ZVJlY29yZHNQcmV2aWV3KCk7CiAgICAgIGlmICghcHJldmlldy5yZW1vdmVDb3VudCkgewogICAgICAgIFRvYXN0LnNob3coJ05vIGR1cGxpY2F0ZSByZWNvcmRzIGZvdW5kLicsICdzdWNjZXNzJyk7CiAgICAgICAgcmV0dXJuOwogICAgICB9CiAgICAgIGNvbnN0IGVkaXRlZE5vdGUgPSBwcmV2aWV3LmVkaXRlZEluUmVtb3ZlQ291bnQKICAgICAgICA/IGBcblxuJHtwcmV2aWV3LmVkaXRlZEluUmVtb3ZlQ291bnR9IG9mIHRoZSBjb3BpZXMgdG8gYmUgcmVtb3ZlZCB3ZXJlIGhhbmQtZWRpdGVkIGFmdGVyIGltcG9ydCDigJQgdGhlIG1vc3QgcmVjZW50IGltcG9ydCBpcyBzdGlsbCB0aGUgb25lIGtlcHQuYAogICAgICAgIDogJyc7CiAgICAgIGNvbnN0IG1lc3NhZ2UgPQogICAgICAgIGBGb3VuZCAke3ByZXZpZXcucmVtb3ZlQ291bnR9IGR1cGxpY2F0ZSByZWNvcmQocykgYWNyb3NzICR7cHJldmlldy5ncm91cENvdW50fSBwb3N0KHMpIGAgKwogICAgICAgIGAoc2FtZSBwdWJsaXNoIGRhdGUgYW5kIGNhcHRpb24pLiBSZW1vdmUgdGhlIG9sZGVyIGNvcGllcyBhbmQga2VlcCB0aGUgbW9zdCByZWNlbnRseSB1cGxvYWRlZCBvbmUgZm9yIGVhY2g/YCArCiAgICAgICAgYFxuXG5UaGlzIHVwZGF0ZXMgdGhlIGRhc2hib2FyZCwgY29tcGFyaXNvbnMsIGFuZCByZXBvcnRzLiBFdmVyeSBvcmlnaW5hbCBpbXBvcnQgc3RheXMgaW4gVXBsb2FkIEhpc3RvcnkuYCArCiAgICAgICAgZWRpdGVkTm90ZTsKICAgICAgaWYgKCF3aW5kb3cuY29uZmlybShtZXNzYWdlKSkgcmV0dXJuOwogICAgICBjb25zdCByZXN1bHQgPSBhd2FpdCBBcGkucmVzb2x2ZUR1cGxpY2F0ZVJlY29yZHMoKTsKICAgICAgVG9hc3Quc2hvdyhgUmVtb3ZlZCAke3Jlc3VsdC5yZW1vdmVkQ291bnR9IGR1cGxpY2F0ZSByZWNvcmQocykuYCwgJ3N1Y2Nlc3MnKTsKICAgICAgcGFnZSA9IDE7CiAgICAgIGF3YWl0IGxvYWQoKTsKICAgICAgd2luZG93LmRpc3BhdGNoRXZlbnQobmV3IEN1c3RvbUV2ZW50KCdscnM6ZGF0YS11cGRhdGVkJykpOwogICAgfSBjYXRjaCAoZXJyKSB7CiAgICAgIFRvYXN0LnNob3coZXJyLm1lc3NhZ2UgfHwgJ0R1cGxpY2F0ZSBjbGVhbnVwIGZhaWxlZC4nLCAnZXJyb3InKTsKICAgIH0gZmluYWxseSB7CiAgICAgIGlmIChidG4pIGJ0bi5kaXNhYmxlZCA9IGZhbHNlOwogICAgfQogIH0KCiAgYXN5bmMgZnVuY3Rpb24gaGFuZGxlRGVsZXRlKHJvdywgcGxhdGZvcm0pIHsKICAgIGNvbnN0IGNhcHRpb24gPSAocm93LmNhcHRpb24gfHwgJyhubyBjYXB0aW9uKScpLnNsaWNlKDAsIDYwKTsKICAgIGNvbnN0IG1lc3NhZ2UgPSBwbGF0Zm9ybSA9PT0gJ2FsbCcKICAgICAgPyBgRGVsZXRlIHRoaXMgZW50aXJlIHJlY29yZCDigJQgIiR7Y2FwdGlvbn0iIOKAlCBhY3Jvc3MgZXZlcnkgcGxhdGZvcm0/IEl0cyBvcmlnaW5hbCBpbXBvcnQgc3RheXMgaW4gVXBsb2FkIEhpc3RvcnksIGJ1dCBpdCB3aWxsIGRpc2FwcGVhciBmcm9tIHRoZSBkYXNoYm9hcmQsIGNvbXBhcmlzb25zLCBhbmQgcmVwb3J0cy5gCiAgICAgIDogYFJlbW92ZSB0aGlzIHJlY29yZCdzICR7cGxhdGZvcm1MYWJlbChwbGF0Zm9ybSl9IGRhdGEg4oCUICIke2NhcHRpb259Ij8gSWYgdGhpcyBpcyBpdHMgb25seSBwbGF0Zm9ybSwgdGhlIHdob2xlIHJlY29yZCB3aWxsIGJlIHJlbW92ZWQgZnJvbSB0aGUgZGFzaGJvYXJkLmA7CiAgICBpZiAoIXdpbmRvdy5jb25maXJtKG1lc3NhZ2UpKSByZXR1cm47CiAgICB0cnkgewogICAgICBpZiAocGxhdGZvcm0gPT09ICdhbGwnKSBhd2FpdCBBcGkuZGVsZXRlUmVjb3JkUG9zdChyb3cucG9zdElkKTsKICAgICAgZWxzZSBhd2FpdCBBcGkuZGVsZXRlUmVjb3JkUGxhdGZvcm0ocm93LnBvc3RJZCwgcGxhdGZvcm0pOwogICAgICBUb2FzdC5zaG93KCdSZWNvcmQgZGVsZXRlZC4nLCAnc3VjY2VzcycpOwogICAgICBhd2FpdCBsb2FkKCk7CiAgICAgIHdpbmRvdy5kaXNwYXRjaEV2ZW50KG5ldyBDdXN0b21FdmVudCgnbHJzOmRhdGEtdXBkYXRlZCcpKTsKICAgIH0gY2F0Y2ggKGVycikgewogICAgICBUb2FzdC5zaG93KGVyci5tZXNzYWdlLCAnZXJyb3InKTsKICAgIH0KICB9CgogIGZ1bmN0aW9uIHJlbW92ZUV4aXN0aW5nT3ZlcmxheSgpIHsKICAgIGNvbnN0IG92ZXJsYXkgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncmVjb3JkTW9kYWxPdmVybGF5Jyk7CiAgICBpZiAob3ZlcmxheSkgb3ZlcmxheS5yZW1vdmUoKTsKICB9CgogIGZ1bmN0aW9uIGNsb3NlTW9kYWwoKSB7CiAgICByZW1vdmVFeGlzdGluZ092ZXJsYXkoKTsKICAgIG1vZGFsU3RhdGUgPSBudWxsOwogIH0KCiAgLy8gT25seSBjbGVhcnMgdGhlIHN0YWxlIERPTSBub2RlIOKAlCBOT1QgbW9kYWxTdGF0ZS4gcmVuZGVyRWRpdE1vZGFsIHJlYWRzCiAgLy8gbW9kYWxTdGF0ZSByaWdodCBhZnRlciBjYWxsaW5nIHRoaXMgdG8gYnVpbGQgdGhlIGZvcm07IGlmIHRoaXMgY2FsbGVkCiAgLy8gdGhlIHJlYWwgY2xvc2VNb2RhbCgpIChhcyBpdCB1c2VkIHRvKSwgdGhhdCByZXNldCBtb2RhbFN0YXRlIHRvIG51bGwgb3V0CiAgLy8gZnJvbSB1bmRlciBpdCBiZWZvcmUgdGhlIHJlYWQsIHdoaWNoIGlzIGV4YWN0bHkgd2h5IEVkaXQgd2FzIGJyb2tlbi4KICBmdW5jdGlvbiBtb2RhbFNoZWxsKHRpdGxlVGV4dCkgewogICAgcmVtb3ZlRXhpc3RpbmdPdmVybGF5KCk7CiAgICBjb25zdCBvdmVybGF5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBvdmVybGF5LmNsYXNzTmFtZSA9ICdtb2RhbC1vdmVybGF5JzsKICAgIG92ZXJsYXkuaWQgPSAncmVjb3JkTW9kYWxPdmVybGF5JzsKICAgIG92ZXJsYXkuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoZSkgPT4geyBpZiAoZS50YXJnZXQgPT09IG92ZXJsYXkpIGNsb3NlTW9kYWwoKTsgfSk7CiAgICBjb25zdCBwYW5lbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcGFuZWwuY2xhc3NOYW1lID0gJ21vZGFsLXBhbmVsIHdpZGUnOwogICAgcGFuZWwuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMicsIHRpdGxlVGV4dCkpOwogICAgb3ZlcmxheS5hcHBlbmRDaGlsZChwYW5lbCk7CiAgICByZXR1cm4geyBvdmVybGF5LCBwYW5lbCB9OwogIH0KCiAgZnVuY3Rpb24gcmVjb3JkU3VidGl0bGUocikgewogICAgcmV0dXJuIGBTaGVldCAiJHtyLnNoZWV0TmFtZX0iLCByb3cgJHtyLnJvd051bWJlcn0ke3IucG9zdElkID8gYCDigJQgbGlua2VkIHRvIGRhc2hib2FyZCBwb3N0ICMke3IucG9zdElkfWAgOiAnIOKAlCBub3QgcGFydCBvZiB0aGUgZGFzaGJvYXJkIChlLmcuIG5lZWRzIGEgdmFsaWQgZGF0ZSknfWA7CiAgfQoKICAvLyAtLS0tLS0tLS0tIFZpZXcgcG9wdXA6IHJlYWQtb25seSwgZXZlcnkgZmllbGQsIGdyb3VwZWQgaW50byBzZWN0aW9ucyAtLS0tLS0tLS0tCiAgYXN5bmMgZnVuY3Rpb24gb3BlblZpZXcoaWQpIHsKICAgIGNvbnN0IHJlY29yZCA9IGF3YWl0IEFwaS5nZXRSZWNvcmQoaWQpOwogICAgY29uc3QgeyBvdmVybGF5LCBwYW5lbCB9ID0gbW9kYWxTaGVsbCgnUmVjb3JkIGRldGFpbHMnKTsKICAgIHBhbmVsLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgcmVjb3JkU3VidGl0bGUocmVjb3JkKSwgJ21vZGFsLXN1YicpKTsKCiAgICBjb25zdCBncm91cHMgPSBncm91cEZpZWxkUm93cyhjb2x1bW5MYWJlbHNGb3IocmVjb3JkKSwgcmVjb3JkLnZhbHVlcyk7CiAgICBncm91cHMuZm9yRWFjaCgoZykgPT4gewogICAgICBjb25zdCBzZWN0aW9uID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIHNlY3Rpb24uY2xhc3NOYW1lID0gJ3JlY29yZC1zZWN0aW9uJzsKICAgICAgc2VjdGlvbi5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2g0JywgZy5ncm91cCkpOwogICAgICBjb25zdCBncmlkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIGdyaWQuY2xhc3NOYW1lID0gJ2Zvcm0tZ3JpZCc7CiAgICAgIGcuZmllbGRzLmZvckVhY2goKGYpID0+IHsKICAgICAgICBjb25zdCBmaWVsZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICAgIGZpZWxkLmNsYXNzTmFtZSA9ICd2aWV3LWZpZWxkJzsKICAgICAgICBmaWVsZC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIGYubGFiZWwsICd2aWV3LWxhYmVsJykpOwogICAgICAgIGNvbnN0IHZhbCA9IGYudmFsdWUgPT09IHVuZGVmaW5lZCB8fCBmLnZhbHVlID09PSBudWxsIHx8IGYudmFsdWUgPT09ICcnID8gJ+KAlCcgOiBTdHJpbmcoZi52YWx1ZSk7CiAgICAgICAgZmllbGQuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCB2YWwsICd2aWV3LXZhbHVlJykpOwogICAgICAgIGdyaWQuYXBwZW5kQ2hpbGQoZmllbGQpOwogICAgICB9KTsKICAgICAgc2VjdGlvbi5hcHBlbmRDaGlsZChncmlkKTsKICAgICAgcGFuZWwuYXBwZW5kQ2hpbGQoc2VjdGlvbik7CiAgICB9KTsKCiAgICBjb25zdCBhY3Rpb25zID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBhY3Rpb25zLmNsYXNzTmFtZSA9ICdtb2RhbC1hY3Rpb25zJzsKICAgIGNvbnN0IGJ0blJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYnRuUm93LmNsYXNzTmFtZSA9ICdidG4tcm93JzsKICAgIGNvbnN0IGNsb3NlQnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBjbG9zZUJ0bi5jbGFzc05hbWUgPSAnYnRuJzsKICAgIGNsb3NlQnRuLnRleHRDb250ZW50ID0gJ0Nsb3NlJzsKICAgIGNsb3NlQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgY2xvc2VNb2RhbCk7CiAgICBjb25zdCBlZGl0QnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBlZGl0QnRuLmNsYXNzTmFtZSA9ICdidG4gcHJpbWFyeSc7CiAgICBlZGl0QnRuLnRleHRDb250ZW50ID0gJ0VkaXQgdGhpcyByZWNvcmQnOwogICAgZWRpdEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IG9wZW5FZGl0b3IocmVjb3JkLmlkKSk7CiAgICBidG5Sb3cuYXBwZW5kKGNsb3NlQnRuLCBlZGl0QnRuKTsKICAgIGFjdGlvbnMuYXBwZW5kQ2hpbGQoYnRuUm93KTsKICAgIHBhbmVsLmFwcGVuZENoaWxkKGFjdGlvbnMpOwoKICAgIGRvY3VtZW50LmJvZHkuYXBwZW5kQ2hpbGQob3ZlcmxheSk7CiAgfQoKICAvLyAtLS0tLS0tLS0tIEVkaXQgcG9wdXA6IGV2ZXJ5IGZpZWxkLCBncm91cGVkIGludG8gc2VjdGlvbnMsIGFsbCBlZGl0YWJsZSAtLS0tLS0tLS0tCiAgYXN5bmMgZnVuY3Rpb24gb3BlbkVkaXRvcihpZCkgewogICAgY29uc3QgcmVjb3JkID0gYXdhaXQgQXBpLmdldFJlY29yZChpZCk7CiAgICBtb2RhbFN0YXRlID0geyByZWNvcmQsIHZhbHVlczogWy4uLnJlY29yZC52YWx1ZXNdIH07CiAgICByZW5kZXJFZGl0TW9kYWwoKTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlckVkaXRNb2RhbCgpIHsKICAgIGNvbnN0IHIgPSBtb2RhbFN0YXRlLnJlY29yZDsKICAgIGNvbnN0IHsgb3ZlcmxheSwgcGFuZWwgfSA9IG1vZGFsU2hlbGwoJ0VkaXQgcmVjb3JkJyk7CiAgICBwYW5lbC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIHJlY29yZFN1YnRpdGxlKHIpLCAnbW9kYWwtc3ViJykpOwoKICAgIGNvbnN0IGdyb3VwcyA9IGdyb3VwRmllbGRSb3dzKGNvbHVtbkxhYmVsc0ZvcihyKSwgbW9kYWxTdGF0ZS52YWx1ZXMpOwogICAgZ3JvdXBzLmZvckVhY2goKGcpID0+IHsKICAgICAgY29uc3Qgc2VjdGlvbiA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBzZWN0aW9uLmNsYXNzTmFtZSA9ICdyZWNvcmQtc2VjdGlvbic7CiAgICAgIHNlY3Rpb24uYXBwZW5kQ2hpbGQodGV4dEVsKCdoNCcsIGcuZ3JvdXApKTsKICAgICAgY29uc3QgZ3JpZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBncmlkLmNsYXNzTmFtZSA9ICdmb3JtLWdyaWQnOwogICAgICBnLmZpZWxkcy5mb3JFYWNoKChmKSA9PiB7CiAgICAgICAgY29uc3QgZmllbGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgICBmaWVsZC5jbGFzc05hbWUgPSAnZm9ybS1maWVsZCc7CiAgICAgICAgZmllbGQuYXBwZW5kQ2hpbGQodGV4dEVsKCdsYWJlbCcsIGYubGFiZWwpKTsKICAgICAgICBjb25zdCBzdHJWYWwgPSBmLnZhbHVlID09PSB1bmRlZmluZWQgfHwgZi52YWx1ZSA9PT0gbnVsbCA/ICcnIDogU3RyaW5nKGYudmFsdWUpOwogICAgICAgIGNvbnN0IGlzTG9uZyA9IHN0clZhbC5sZW5ndGggPiA4MDsKICAgICAgICBjb25zdCBpbnB1dCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoaXNMb25nID8gJ3RleHRhcmVhJyA6ICdpbnB1dCcpOwogICAgICAgIGlmICghaXNMb25nKSBpbnB1dC50eXBlID0gJ3RleHQnOwogICAgICAgIGVsc2UgZmllbGQuc3R5bGUuZ3JpZENvbHVtbiA9ICcxIC8gLTEnOwogICAgICAgIGlucHV0LnZhbHVlID0gc3RyVmFsOwogICAgICAgIGlucHV0LmFkZEV2ZW50TGlzdGVuZXIoJ2lucHV0JywgKCkgPT4geyBtb2RhbFN0YXRlLnZhbHVlc1tmLmlkeF0gPSBpbnB1dC52YWx1ZTsgfSk7CiAgICAgICAgZmllbGQuYXBwZW5kQ2hpbGQoaW5wdXQpOwogICAgICAgIGdyaWQuYXBwZW5kQ2hpbGQoZmllbGQpOwogICAgICB9KTsKICAgICAgc2VjdGlvbi5hcHBlbmRDaGlsZChncmlkKTsKICAgICAgcGFuZWwuYXBwZW5kQ2hpbGQoc2VjdGlvbik7CiAgICB9KTsKCiAgICBjb25zdCBhY3Rpb25zID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBhY3Rpb25zLmNsYXNzTmFtZSA9ICdtb2RhbC1hY3Rpb25zJzsKICAgIGNvbnN0IGVycm9yTXNnID0gdGV4dEVsKCdzcGFuJywgJycsICdtdXRlZCcpOwogICAgZXJyb3JNc2cuaWQgPSAnbW9kYWxFcnJvck1zZyc7CiAgICBjb25zdCBidG5Sb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGJ0blJvdy5jbGFzc05hbWUgPSAnYnRuLXJvdyc7CiAgICBjb25zdCBjYW5jZWxCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIGNhbmNlbEJ0bi5jbGFzc05hbWUgPSAnYnRuJzsKICAgIGNhbmNlbEJ0bi50ZXh0Q29udGVudCA9ICdDYW5jZWwnOwogICAgY2FuY2VsQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgY2xvc2VNb2RhbCk7CiAgICBjb25zdCBzYXZlQnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBzYXZlQnRuLmNsYXNzTmFtZSA9ICdidG4gcHJpbWFyeSc7CiAgICBzYXZlQnRuLnRleHRDb250ZW50ID0gJ1NhdmUgY2hhbmdlcyc7CiAgICBzYXZlQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gc2F2ZUVkaXQoc2F2ZUJ0bikpOwogICAgYnRuUm93LmFwcGVuZChjYW5jZWxCdG4sIHNhdmVCdG4pOwogICAgYWN0aW9ucy5hcHBlbmQoZXJyb3JNc2csIGJ0blJvdyk7CiAgICBwYW5lbC5hcHBlbmRDaGlsZChhY3Rpb25zKTsKCiAgICBkb2N1bWVudC5ib2R5LmFwcGVuZENoaWxkKG92ZXJsYXkpOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gc2F2ZUVkaXQoYnRuKSB7CiAgICBjb25zdCBlcnJvckVsID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ21vZGFsRXJyb3JNc2cnKTsKICAgIGVycm9yRWwudGV4dENvbnRlbnQgPSAnJzsKICAgIGJ0bi5kaXNhYmxlZCA9IHRydWU7CiAgICBidG4udGV4dENvbnRlbnQgPSAnU2F2aW5n4oCmJzsKICAgIHRyeSB7CiAgICAgIGF3YWl0IEFwaS51cGRhdGVSZWNvcmQobW9kYWxTdGF0ZS5yZWNvcmQuaWQsIG1vZGFsU3RhdGUudmFsdWVzKTsKICAgICAgVG9hc3Quc2hvdygnUmVjb3JkIHVwZGF0ZWQuJywgJ3N1Y2Nlc3MnKTsKICAgICAgY2xvc2VNb2RhbCgpOwogICAgICBhd2FpdCBsb2FkKCk7CiAgICAgIHdpbmRvdy5kaXNwYXRjaEV2ZW50KG5ldyBDdXN0b21FdmVudCgnbHJzOmRhdGEtdXBkYXRlZCcpKTsKICAgIH0gY2F0Y2ggKGVycikgewogICAgICBlcnJvckVsLnRleHRDb250ZW50ID0gZXJyLm1lc3NhZ2U7CiAgICAgIGVycm9yRWwuc3R5bGUuY29sb3IgPSAndmFyKC0tc3RhdHVzLWNyaXRpY2FsKSc7CiAgICAgIGJ0bi5kaXNhYmxlZCA9IGZhbHNlOwogICAgICBidG4udGV4dENvbnRlbnQgPSAnU2F2ZSBjaGFuZ2VzJzsKICAgIH0KICB9CgogIGFzeW5jIGZ1bmN0aW9uIHJlbmRlcigpIHsKICAgIHJvb3QgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndmlldy1yZWNvcmRzJyk7CiAgICBwYWdlID0gMTsKICAgIHNoZWxsKCk7CiAgICBhd2FpdCBsb2FkKCk7CiAgfQoKICByZXR1cm4geyByZW5kZXIsIHJlbG9hZDogbG9hZCwgb3BlblZpZXcgfTsKfSkoKTsKCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBDb21wYXJpc29ucyB0YWI6IHdlZWstdnMtd2VlaywgY3VzdG9tIHJhbmdlLCBtb250aGx5LAogICBxdWFydGVybHksIFlURC4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmNvbnN0IENvbXBhcmlzb24gPSAoKCkgPT4gewogIGNvbnN0IE1PREVTID0gWwogICAgeyBrZXk6ICdwbGF0Zm9ybXMnLCBsYWJlbDogJ0FsbCBQbGF0Zm9ybXMnIH0sCiAgICB7IGtleTogJ3dlZWsnLCBsYWJlbDogJ1dlZWsgdnMgV2VlaycgfSwKICAgIHsga2V5OiAnY3VzdG9tJywgbGFiZWw6ICdDdXN0b20gUmFuZ2UnIH0sCiAgICB7IGtleTogJ21vbnRoJywgbGFiZWw6ICdNb250aGx5JyB9LAogICAgeyBrZXk6ICdxdWFydGVyJywgbGFiZWw6ICdRdWFydGVybHknIH0sCiAgICB7IGtleTogJ3l0ZCcsIGxhYmVsOiAnWWVhciB0byBEYXRlJyB9LAogIF07CiAgY29uc3QgTUVUUklDX1JPV1MgPSBbCiAgICB7IGtleTogJ3ZpZXdzJywgbGFiZWw6ICdWaWV3cycgfSwKICAgIHsga2V5OiAncmVhY2gnLCBsYWJlbDogJ1JlYWNoJyB9LAogICAgeyBrZXk6ICdpbXByZXNzaW9ucycsIGxhYmVsOiAnSW1wcmVzc2lvbnMnIH0sCiAgICB7IGtleTogJ2VuZ2FnZW1lbnQnLCBsYWJlbDogJ0VuZ2FnZW1lbnQnIH0sCiAgICB7IGtleTogJ2NsaWNrcycsIGxhYmVsOiAnQ2xpY2tzJyB9LAogICAgeyBrZXk6ICdmb2xsb3dlcnNfZ2FpbmVkJywgbGFiZWw6ICdGb2xsb3dlcnMgR2FpbmVkJyB9LAogICAgeyBrZXk6ICd3YXRjaF90aW1lX3NlY29uZHMnLCBsYWJlbDogJ1dhdGNoIFRpbWUnIH0sCiAgICB7IGtleTogJ3NoYXJlcycsIGxhYmVsOiAnU2hhcmVzJyB9LAogICAgeyBrZXk6ICdjb21tZW50cycsIGxhYmVsOiAnQ29tbWVudHMnIH0sCiAgICB7IGtleTogJ3NhdmVzJywgbGFiZWw6ICdTYXZlcycgfSwKICBdOwoKICBsZXQgbW9kZSA9ICdwbGF0Zm9ybXMnOwogIGxldCByb290OwogIGxldCBwbGF0Zm9ybUNoYXJ0TWV0cmljID0gJ2VuZ2FnZW1lbnQnOwogIGxldCBjYXJkU29ydE1vZGUgPSAnb3ZlcmFsbCc7CiAgbGV0IGNhcmRQbGF0Zm9ybUZpbHRlciA9ICdhbGwnOwogIC8vIENvbXBhcmlzb25zIGtlZXBzIGl0cyBvd24gUGxhdGZvcm0gZmlsdGVyLCBpbmRlcGVuZGVudCBvZiB0aGUgRGFzaGJvYXJkJ3Mg4oCUIHRoZQogIC8vIHNoYXJlZCBmaWx0ZXIgYmFyIGlzIGhpZGRlbiBvbiB0aGlzIHRhYiAoc2VlIHN3aXRjaFRhYiksIGJ1dCBpdHMgc3RhdGUgcGVyc2lzdHMKICAvLyBpbiBtZW1vcnksIHNvIHdpdGhvdXQgdGhpcyBldmVyeSBtb2RlIGhlcmUgd291bGQgc2lsZW50bHkga2VlcCB3aGF0ZXZlciBwbGF0Zm9ybQogIC8vIHdhcyBsYXN0IHBpY2tlZCBvbiBEYXNoYm9hcmQgd2l0aCBubyB2aXNpYmxlIGNvbnRyb2wgdG8gc2VlIG9yIGNoYW5nZSBpdC4KICBsZXQgY29tcGFyaXNvblBsYXRmb3JtID0gJ2FsbCc7CgogIGZ1bmN0aW9uIG1vbmRheU9mKGRhdGVTdHIpIHsKICAgIGNvbnN0IGQgPSBuZXcgRGF0ZShkYXRlU3RyKTsKICAgIGNvbnN0IGRheSA9IGQuZ2V0RGF5KCk7CiAgICBjb25zdCBkaWZmID0gZGF5ID09PSAwID8gNiA6IGRheSAtIDE7CiAgICBkLnNldERhdGUoZC5nZXREYXRlKCkgLSBkaWZmKTsKICAgIHJldHVybiBkLnRvSVNPU3RyaW5nKCkuc2xpY2UoMCwgMTApOwogIH0KICBmdW5jdGlvbiBhZGREYXlzKGRhdGVTdHIsIG4pIHsKICAgIGNvbnN0IGQgPSBuZXcgRGF0ZShkYXRlU3RyKTsKICAgIGQuc2V0RGF0ZShkLmdldERhdGUoKSArIG4pOwogICAgcmV0dXJuIGQudG9JU09TdHJpbmcoKS5zbGljZSgwLCAxMCk7CiAgfQoKICBmdW5jdGlvbiBzaGVsbCgpIHsKICAgIHJvb3QuaW5uZXJIVE1MID0gJyc7CgogICAgY29uc3QgdGFicyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGFicy5jbGFzc05hbWUgPSAnbW9kZS10YWJzJzsKICAgIE1PREVTLmZvckVhY2goKG0pID0+IHsKICAgICAgY29uc3QgYnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICAgIGJ0bi50ZXh0Q29udGVudCA9IG0ubGFiZWw7CiAgICAgIGJ0bi50eXBlID0gJ2J1dHRvbic7CiAgICAgIGlmIChtLmtleSA9PT0gbW9kZSkgYnRuLmNsYXNzTGlzdC5hZGQoJ2lzLWFjdGl2ZScpOwogICAgICBidG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7CiAgICAgICAgbW9kZSA9IG0ua2V5OwogICAgICAgIHNoZWxsKCk7CiAgICAgIH0pOwogICAgICB0YWJzLmFwcGVuZENoaWxkKGJ0bik7CiAgICB9KTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQodGFicyk7CgogICAgY29uc3QgY29udHJvbHMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNvbnRyb2xzLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGNvbnRyb2xzLmlkID0gJ2NvbXBhcmlzb25Db250cm9scyc7CiAgICByb290LmFwcGVuZENoaWxkKGNvbnRyb2xzKTsKCiAgICBjb25zdCByZXN1bHRzID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICByZXN1bHRzLmlkID0gJ2NvbXBhcmlzb25SZXN1bHRzJzsKICAgIHJvb3QuYXBwZW5kQ2hpbGQocmVzdWx0cyk7CgogICAgcmVuZGVyQ29udHJvbHMoKTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlckNvbnRyb2xzKCkgewogICAgY29uc3QgY29udHJvbHMgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnY29tcGFyaXNvbkNvbnRyb2xzJyk7CiAgICBjb250cm9scy5pbm5lckhUTUwgPSAnJzsKICAgIGNvbnN0IHJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcm93LmNsYXNzTmFtZSA9ICdidG4tcm93JzsKICAgIHJvdy5zdHlsZS5hbGlnbkl0ZW1zID0gJ2VuZCc7CgogICAgY29uc3QgdG9kYXkgPSBuZXcgRGF0ZSgpLnRvSVNPU3RyaW5nKCkuc2xpY2UoMCwgMTApOwogICAgY29uc3QgdGhpc1llYXIgPSBuZXcgRGF0ZSgpLmdldEZ1bGxZZWFyKCk7CgogICAgaWYgKG1vZGUgPT09ICdwbGF0Zm9ybXMnKSB7CiAgICAgIGNvbnN0IGZGcm9tID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsgZkZyb20udHlwZSA9ICdkYXRlJzsgZkZyb20uaWQgPSAncGxhdGZvcm1SZXBvcnRGcm9tJzsKICAgICAgY29uc3QgZlRvID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsgZlRvLnR5cGUgPSAnZGF0ZSc7IGZUby5pZCA9ICdwbGF0Zm9ybVJlcG9ydFRvJzsKICAgICAgY29uc3QgYXBwbHlCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgICAgYXBwbHlCdG4uY2xhc3NOYW1lID0gJ2J0biBwcmltYXJ5JzsKICAgICAgYXBwbHlCdG4udHlwZSA9ICdidXR0b24nOwogICAgICBhcHBseUJ0bi50ZXh0Q29udGVudCA9ICdBcHBseSBSYW5nZSc7CiAgICAgIGFwcGx5QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gbG9hZFBsYXRmb3JtUmVwb3J0KHsgZGF0ZUZyb206IGZGcm9tLnZhbHVlLCBkYXRlVG86IGZUby52YWx1ZSB9KSk7CiAgICAgIGNvbnN0IGNsZWFyQnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICAgIGNsZWFyQnRuLmNsYXNzTmFtZSA9ICdidG4nOwogICAgICBjbGVhckJ0bi50eXBlID0gJ2J1dHRvbic7CiAgICAgIGNsZWFyQnRuLnRleHRDb250ZW50ID0gJ0FsbCBUaW1lJzsKICAgICAgY2xlYXJCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7IGZGcm9tLnZhbHVlID0gJyc7IGZUby52YWx1ZSA9ICcnOyBsb2FkUGxhdGZvcm1SZXBvcnQoe30pOyB9KTsKICAgICAgcm93LmFwcGVuZCgKICAgICAgICBsYWJlbGVkKCdGcm9tIChvcHRpb25hbCknLCBmRnJvbSksCiAgICAgICAgbGFiZWxlZCgnVG8gKG9wdGlvbmFsKScsIGZUbyksCiAgICAgICAgYXBwbHlCdG4sCiAgICAgICAgY2xlYXJCdG4KICAgICAgKTsKICAgICAgY29udHJvbHMuYXBwZW5kQ2hpbGQocm93KTsKICAgICAgbG9hZFBsYXRmb3JtUmVwb3J0KHt9KTsKICAgICAgcmV0dXJuOwogICAgfSBlbHNlIGlmIChtb2RlID09PSAnd2VlaycpIHsKICAgICAgLy8gV2VlayBBIGlzIGFsd2F5cyB0aGUgUHJldmlvdXMgcGVyaW9kLCBXZWVrIEIgaXMgYWx3YXlzIHRoZSBDdXJyZW50L21vc3QgcmVjZW50CiAgICAgIC8vIHBlcmlvZCDigJQgcnVuQ29tcGFyZSgpJ3MgZmlyc3QgYXJndW1lbnQgaXMgdGhlICJjdXJyZW50IiBzbG90IGV2ZXJ5IG90aGVyIG1vZGUKICAgICAgLy8gaW4gdGhpcyB0YWIgZmVlZHMgaXQgKHBlcmNlbnRDaGFuZ2UgPSAoY3VycmVudCAtIHByZXZpb3VzKSAvIHByZXZpb3VzKSwgc28gQiBnb2VzCiAgICAgIC8vIGluIGZpcnN0IGFuZCBBIHNlY29uZCwgcmVnYXJkbGVzcyBvZiB3aGljaCBjYWxlbmRhciB3ZWVrIGlzIGVhcmxpZXIgb3IgbGF0ZXIuCiAgICAgIGNvbnN0IHdBID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsgd0EudHlwZSA9ICdkYXRlJzsgd0EudmFsdWUgPSBtb25kYXlPZihhZGREYXlzKHRvZGF5LCAtNykpOwogICAgICBjb25zdCB3QiA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7IHdCLnR5cGUgPSAnZGF0ZSc7IHdCLnZhbHVlID0gbW9uZGF5T2YodG9kYXkpOwogICAgICByb3cuYXBwZW5kKGxhYmVsZWQoJ1dlZWsgQSAoUHJldmlvdXMpJywgd0EpLCBsYWJlbGVkKCdXZWVrIEIgKEN1cnJlbnQpJywgd0IpLCBwbGF0Zm9ybUZpbHRlckZpZWxkKCksIHJ1bkJ0bigoKSA9PiB7CiAgICAgICAgY29uc3QgcmFuZ2VBID0geyBmcm9tOiBtb25kYXlPZih3QS52YWx1ZSksIHRvOiBhZGREYXlzKG1vbmRheU9mKHdBLnZhbHVlKSwgNikgfTsKICAgICAgICBjb25zdCByYW5nZUIgPSB7IGZyb206IG1vbmRheU9mKHdCLnZhbHVlKSwgdG86IGFkZERheXMobW9uZGF5T2Yod0IudmFsdWUpLCA2KSB9OwogICAgICAgIHJ1bkNvbXBhcmUocmFuZ2VCLCByYW5nZUEsIGBXZWVrIG9mICR7Rm9ybWF0LmRhdGUocmFuZ2VCLmZyb20pfWAsIGBXZWVrIG9mICR7Rm9ybWF0LmRhdGUocmFuZ2VBLmZyb20pfWApOwogICAgICB9KSk7CiAgICB9IGVsc2UgaWYgKG1vZGUgPT09ICdjdXN0b20nKSB7CiAgICAgIGNvbnN0IGZBID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsgZkEudHlwZSA9ICdkYXRlJzsgZkEudmFsdWUgPSBhZGREYXlzKHRvZGF5LCAtMTMpOwogICAgICBjb25zdCB0QSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7IHRBLnR5cGUgPSAnZGF0ZSc7IHRBLnZhbHVlID0gdG9kYXk7CiAgICAgIGNvbnN0IGZCID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsgZkIudHlwZSA9ICdkYXRlJzsgZkIudmFsdWUgPSBhZGREYXlzKHRvZGF5LCAtMjcpOwogICAgICBjb25zdCB0QiA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7IHRCLnR5cGUgPSAnZGF0ZSc7IHRCLnZhbHVlID0gYWRkRGF5cyh0b2RheSwgLTE0KTsKICAgICAgcm93LmFwcGVuZCgKICAgICAgICBsYWJlbGVkKCdSYW5nZSBBIGZyb20nLCBmQSksIGxhYmVsZWQoJ3RvJywgdEEpLAogICAgICAgIGxhYmVsZWQoJ1JhbmdlIEIgZnJvbScsIGZCKSwgbGFiZWxlZCgndG8nLCB0QiksCiAgICAgICAgcGxhdGZvcm1GaWx0ZXJGaWVsZCgpLAogICAgICAgIHJ1bkJ0bigoKSA9PiBydW5Db21wYXJlKHsgZnJvbTogZkEudmFsdWUsIHRvOiB0QS52YWx1ZSB9LCB7IGZyb206IGZCLnZhbHVlLCB0bzogdEIudmFsdWUgfSwgJ1JhbmdlIEEnLCAnUmFuZ2UgQicpKQogICAgICApOwogICAgfSBlbHNlIGlmIChtb2RlID09PSAnbW9udGgnKSB7CiAgICAgIGNvbnN0IHkgPSB5ZWFyU2VsZWN0KHRoaXNZZWFyKTsgY29uc3QgbSA9IG1vbnRoU2VsZWN0KG5ldyBEYXRlKCkuZ2V0TW9udGgoKSArIDEpOwogICAgICBjb25zdCB0b2dnbGUgPSBwZXJpb2RUb2dnbGUoKTsKICAgICAgcm93LmFwcGVuZChsYWJlbGVkKCdZZWFyJywgeSksIGxhYmVsZWQoJ01vbnRoJywgbSksIHRvZ2dsZS5lbCwgcGxhdGZvcm1GaWx0ZXJGaWVsZCgpLCBydW5CdG4oYXN5bmMgKCkgPT4gewogICAgICAgIGNvbnN0IHJlcG9ydCA9IGF3YWl0IEFwaS5tb250aGx5KHsgeWVhcjogeS52YWx1ZSwgbW9udGg6IG0udmFsdWUsIC4uLlN0YXRlLmdldEZpbHRlcnMoKSwgcGxhdGZvcm06IGNvbXBhcmlzb25QbGF0Zm9ybSB9KTsKICAgICAgICByZW5kZXJQZXJpb2RSZXBvcnQocmVwb3J0LCB0b2dnbGUuZ2V0KCkpOwogICAgICB9KSk7CiAgICB9IGVsc2UgaWYgKG1vZGUgPT09ICdxdWFydGVyJykgewogICAgICBjb25zdCB5ID0geWVhclNlbGVjdCh0aGlzWWVhcik7IGNvbnN0IHEgPSBxdWFydGVyU2VsZWN0KCk7CiAgICAgIGNvbnN0IHRvZ2dsZSA9IHBlcmlvZFRvZ2dsZSgpOwogICAgICByb3cuYXBwZW5kKGxhYmVsZWQoJ1llYXInLCB5KSwgbGFiZWxlZCgnUXVhcnRlcicsIHEpLCB0b2dnbGUuZWwsIHBsYXRmb3JtRmlsdGVyRmllbGQoKSwgcnVuQnRuKGFzeW5jICgpID0+IHsKICAgICAgICBjb25zdCByZXBvcnQgPSBhd2FpdCBBcGkucXVhcnRlcmx5KHsgeWVhcjogeS52YWx1ZSwgcXVhcnRlcjogcS52YWx1ZSwgLi4uU3RhdGUuZ2V0RmlsdGVycygpLCBwbGF0Zm9ybTogY29tcGFyaXNvblBsYXRmb3JtIH0pOwogICAgICAgIHJlbmRlclBlcmlvZFJlcG9ydChyZXBvcnQsIHRvZ2dsZS5nZXQoKSk7CiAgICAgIH0pKTsKICAgIH0gZWxzZSBpZiAobW9kZSA9PT0gJ3l0ZCcpIHsKICAgICAgY29uc3QgeSA9IHllYXJTZWxlY3QodGhpc1llYXIpOwogICAgICByb3cuYXBwZW5kKGxhYmVsZWQoJ1llYXInLCB5KSwgcGxhdGZvcm1GaWx0ZXJGaWVsZCgpLCBydW5CdG4oYXN5bmMgKCkgPT4gewogICAgICAgIGNvbnN0IHJlcG9ydCA9IGF3YWl0IEFwaS55dGQoeyB5ZWFyOiB5LnZhbHVlLCAuLi5TdGF0ZS5nZXRGaWx0ZXJzKCksIHBsYXRmb3JtOiBjb21wYXJpc29uUGxhdGZvcm0gfSk7CiAgICAgICAgcmVuZGVyUGVyaW9kUmVwb3J0KHJlcG9ydCwgJ3ZzTGFzdFllYXInKTsKICAgICAgfSkpOwogICAgfQoKICAgIGNvbnRyb2xzLmFwcGVuZENoaWxkKHJvdyk7CiAgfQoKICAvKiogVGhlIENvbXBhcmlzb25zIHRhYidzIG93biBQbGF0Zm9ybSBmaWx0ZXIg4oCUIHJldXNlcyB0aGUgc2FtZSBwbGF0Zm9ybSBsaXN0IHRoZQogICAgICBEYXNoYm9hcmQncyBmaWx0ZXIgYmFyIHNob3dzLCBidXQgd3JpdGVzIHRvIGNvbXBhcmlzb25QbGF0Zm9ybSwgbm90IFN0YXRlLCBzbwogICAgICB0aGUgdHdvIHN0YXkgZnVsbHkgaW5kZXBlbmRlbnQuICovCiAgZnVuY3Rpb24gcGxhdGZvcm1GaWx0ZXJGaWVsZCgpIHsKICAgIGNvbnN0IHNlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NlbGVjdCcpOwogICAgY29uc3QgYWxsT3B0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7CiAgICBhbGxPcHQudmFsdWUgPSAnYWxsJzsKICAgIGFsbE9wdC50ZXh0Q29udGVudCA9ICdBbGwgcGxhdGZvcm1zJzsKICAgIHNlbC5hcHBlbmRDaGlsZChhbGxPcHQpOwogICAgY29uc3Qgb3B0aW9ucyA9ICh3aW5kb3cuX19maWx0ZXJPcHRpb25zQ2FjaGUgfHwgeyBwbGF0Zm9ybXM6IFtdIH0pLnBsYXRmb3JtcyB8fCBbXTsKICAgIG9wdGlvbnMuZm9yRWFjaCgocCkgPT4gewogICAgICBjb25zdCBvcHQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsKICAgICAgb3B0LnZhbHVlID0gcC5pZDsKICAgICAgb3B0LnRleHRDb250ZW50ID0gcC5sYWJlbDsKICAgICAgc2VsLmFwcGVuZENoaWxkKG9wdCk7CiAgICB9KTsKICAgIHNlbC52YWx1ZSA9IGNvbXBhcmlzb25QbGF0Zm9ybTsKICAgIHNlbC5hZGRFdmVudExpc3RlbmVyKCdjaGFuZ2UnLCAoKSA9PiB7IGNvbXBhcmlzb25QbGF0Zm9ybSA9IHNlbC52YWx1ZTsgfSk7CiAgICByZXR1cm4gbGFiZWxlZCgnUGxhdGZvcm0nLCBzZWwpOwogIH0KCiAgZnVuY3Rpb24gbGFiZWxlZChsYWJlbCwgZWwpIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHdyYXAuY2xhc3NOYW1lID0gJ2ZpZWxkLWlubGluZSc7CiAgICB3cmFwLmFwcGVuZCh0ZXh0RWwoJ2xhYmVsJywgbGFiZWwpLCBlbCk7CiAgICByZXR1cm4gd3JhcDsKICB9CiAgZnVuY3Rpb24gcnVuQnRuKG9uQ2xpY2spIHsKICAgIGNvbnN0IGJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgYnRuLmNsYXNzTmFtZSA9ICdidG4gcHJpbWFyeSc7CiAgICBidG4udGV4dENvbnRlbnQgPSAnQ29tcGFyZSc7CiAgICBidG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBvbkNsaWNrKCkpOwogICAgcmV0dXJuIGJ0bjsKICB9CiAgZnVuY3Rpb24geWVhclNlbGVjdChkZWZhdWx0WWVhcikgewogICAgY29uc3Qgc2VsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc2VsZWN0Jyk7CiAgICBmb3IgKGxldCB5ID0gZGVmYXVsdFllYXIgLSAzOyB5IDw9IGRlZmF1bHRZZWFyICsgMTsgeSArPSAxKSB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOyBvcHQudmFsdWUgPSB5OyBvcHQudGV4dENvbnRlbnQgPSB5OwogICAgICBpZiAoeSA9PT0gZGVmYXVsdFllYXIpIG9wdC5zZWxlY3RlZCA9IHRydWU7CiAgICAgIHNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfQogICAgcmV0dXJuIHNlbDsKICB9CiAgZnVuY3Rpb24gbW9udGhTZWxlY3QoZGVmYXVsdE1vbnRoKSB7CiAgICBjb25zdCBuYW1lcyA9IFsnSmFudWFyeScsJ0ZlYnJ1YXJ5JywnTWFyY2gnLCdBcHJpbCcsJ01heScsJ0p1bmUnLCdKdWx5JywnQXVndXN0JywnU2VwdGVtYmVyJywnT2N0b2JlcicsJ05vdmVtYmVyJywnRGVjZW1iZXInXTsKICAgIGNvbnN0IHNlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NlbGVjdCcpOwogICAgbmFtZXMuZm9yRWFjaCgobiwgaSkgPT4gewogICAgICBjb25zdCBvcHQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsgb3B0LnZhbHVlID0gaSArIDE7IG9wdC50ZXh0Q29udGVudCA9IG47CiAgICAgIGlmIChpICsgMSA9PT0gZGVmYXVsdE1vbnRoKSBvcHQuc2VsZWN0ZWQgPSB0cnVlOwogICAgICBzZWwuYXBwZW5kQ2hpbGQob3B0KTsKICAgIH0pOwogICAgcmV0dXJuIHNlbDsKICB9CiAgZnVuY3Rpb24gcXVhcnRlclNlbGVjdCgpIHsKICAgIGNvbnN0IGN1cnJlbnRRID0gTWF0aC5mbG9vcihuZXcgRGF0ZSgpLmdldE1vbnRoKCkgLyAzKSArIDE7CiAgICBjb25zdCBzZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzZWxlY3QnKTsKICAgIFsxLCAyLCAzLCA0XS5mb3JFYWNoKChxKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOyBvcHQudmFsdWUgPSBxOyBvcHQudGV4dENvbnRlbnQgPSBgUSR7cX1gOwogICAgICBpZiAocSA9PT0gY3VycmVudFEpIG9wdC5zZWxlY3RlZCA9IHRydWU7CiAgICAgIHNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgICByZXR1cm4gc2VsOwogIH0KICBmdW5jdGlvbiBwZXJpb2RUb2dnbGUoKSB7CiAgICBjb25zdCBzZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzZWxlY3QnKTsKICAgIFtbJ3ZzUHJldmlvdXNQZXJpb2QnLCAndnMgUHJldmlvdXMgUGVyaW9kJ10sIFsndnNMYXN0WWVhcicsICd2cyBTYW1lIFBlcmlvZCBMYXN0IFllYXInXV0uZm9yRWFjaCgoW3YsIGxdKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOyBvcHQudmFsdWUgPSB2OyBvcHQudGV4dENvbnRlbnQgPSBsOwogICAgICBzZWwuYXBwZW5kQ2hpbGQob3B0KTsKICAgIH0pOwogICAgcmV0dXJuIHsgZWw6IGxhYmVsZWQoJ0NvbXBhcmUnLCBzZWwpLCBnZXQ6ICgpID0+IHNlbC52YWx1ZSB9OwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gcnVuQ29tcGFyZShyYW5nZUEsIHJhbmdlQiwgbGFiZWxBLCBsYWJlbEIpIHsKICAgIGNvbnN0IGZpbHRlcnMgPSBTdGF0ZS5nZXRGaWx0ZXJzKCk7CiAgICBjb25zdCByZXN1bHQgPSBhd2FpdCBBcGkuY29tcGFyZSh7CiAgICAgIGZyb21BOiByYW5nZUEuZnJvbSwgdG9BOiByYW5nZUEudG8sIGZyb21COiByYW5nZUIuZnJvbSwgdG9COiByYW5nZUIudG8sCiAgICAgIHBsYXRmb3JtOiBjb21wYXJpc29uUGxhdGZvcm0sIGNhbXBhaWduVHlwZTogZmlsdGVycy5jYW1wYWlnblR5cGUsIGNvbnRlbnRUeXBlOiBmaWx0ZXJzLmNvbnRlbnRUeXBlLAogICAgfSk7CiAgICByZW5kZXJDb21wYXJlUmVzdWx0KHJlc3VsdCwgbGFiZWxBLCBsYWJlbEIpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyUGVyaW9kUmVwb3J0KHJlcG9ydCwgd2hpY2gpIHsKICAgIGNvbnN0IGNtcCA9IHJlcG9ydFt3aGljaF07CiAgICBjb25zdCBsYWJlbEEgPSAnQ3VycmVudCBwZXJpb2QnOwogICAgY29uc3QgbGFiZWxCID0gd2hpY2ggPT09ICd2c0xhc3RZZWFyJyA/ICdTYW1lIHBlcmlvZCBsYXN0IHllYXInIDogJ1ByZXZpb3VzIHBlcmlvZCc7CiAgICByZW5kZXJDb21wYXJlUmVzdWx0KGNtcCwgbGFiZWxBLCBsYWJlbEIsIHJlcG9ydC5yYW5nZSk7CiAgfQoKICBmdW5jdGlvbiBzdGF0VGlsZShsYWJlbCwgY3VycmVudCwgcHJldmlvdXMsIGdyb3d0aCwgaXNEdXJhdGlvbikgewogICAgY29uc3QgdGlsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGlsZS5jbGFzc05hbWUgPSAnc3RhdC10aWxlJzsKICAgIGNvbnN0IGN1ckRpc3BsYXkgPSBpc0R1cmF0aW9uID8gRm9ybWF0LmR1cmF0aW9uKGN1cnJlbnQpIDogRm9ybWF0LmNvbXBhY3QoY3VycmVudCk7CiAgICBjb25zdCBwcmV2RGlzcGxheSA9IGlzRHVyYXRpb24gPyBGb3JtYXQuZHVyYXRpb24ocHJldmlvdXMpIDogRm9ybWF0LmNvbXBhY3QocHJldmlvdXMpOwogICAgdGlsZS5hcHBlbmQoCiAgICAgIHRleHRFbCgnZGl2JywgbGFiZWwsICdzdGF0LWxhYmVsJyksCiAgICAgIHRleHRFbCgnZGl2JywgY3VyRGlzcGxheSwgJ3N0YXQtdmFsdWUnKSwKICAgICAgdGV4dEVsKCdkaXYnLCBgJHtGb3JtYXQucGN0KGdyb3d0aCl9IMK3IHdhcyAke3ByZXZEaXNwbGF5fWAsIGBzdGF0LWRlbHRhICR7Rm9ybWF0LmRlbHRhQ2xhc3MoZ3Jvd3RoKX1gKQogICAgKTsKICAgIHJldHVybiB0aWxlOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyQ29tcGFyZVJlc3VsdChyZXN1bHQsIGxhYmVsQSwgbGFiZWxCLCBoZWFkbGluZSkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdjb21wYXJpc29uUmVzdWx0cycpOwogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKCiAgICBjb25zdCB0aXRsZSA9IHRleHRFbCgnZGl2JywgaGVhZGxpbmUKICAgICAgPyBgJHtGb3JtYXQuZGF0ZShyZXN1bHQucmFuZ2VBLmZyb20pfSDigJMgJHtGb3JtYXQuZGF0ZShyZXN1bHQucmFuZ2VBLnRvKX1gCiAgICAgIDogYCR7bGFiZWxBfTogJHtGb3JtYXQuZGF0ZShyZXN1bHQucmFuZ2VBLmZyb20pfSDigJMgJHtGb3JtYXQuZGF0ZShyZXN1bHQucmFuZ2VBLnRvKX0gIHZzICAke2xhYmVsQn06ICR7Rm9ybWF0LmRhdGUocmVzdWx0LnJhbmdlQi5mcm9tKX0g4oCTICR7Rm9ybWF0LmRhdGUocmVzdWx0LnJhbmdlQi50byl9YCwKICAgICAgJ3NlY3Rpb24tdGl0bGUnKTsKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGl0bGUpOwoKICAgIGNvbnN0IGdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGdyaWQuY2xhc3NOYW1lID0gJ3N0YXQtZ3JpZCc7CiAgICBncmlkLmFwcGVuZENoaWxkKHN0YXRUaWxlKCdQb3N0cycsIHJlc3VsdC5yYW5nZUEudG90YWxzLnBvc3RfY291bnQsIHJlc3VsdC5yYW5nZUIudG90YWxzLnBvc3RfY291bnQsIHJlc3VsdC5ncm93dGgucG9zdF9jb3VudCwgZmFsc2UpKTsKICAgIE1FVFJJQ19ST1dTLmZvckVhY2goKG0pID0+IHsKICAgICAgZ3JpZC5hcHBlbmRDaGlsZChzdGF0VGlsZShtLmxhYmVsLCByZXN1bHQucmFuZ2VBLnRvdGFsc1ttLmtleV0sIHJlc3VsdC5yYW5nZUIudG90YWxzW20ua2V5XSwgcmVzdWx0Lmdyb3d0aFttLmtleV0sIG0ua2V5ID09PSAnd2F0Y2hfdGltZV9zZWNvbmRzJykpOwogICAgfSk7CiAgICB3cmFwLmFwcGVuZENoaWxkKGdyaWQpOwoKICAgIHJlbmRlclBsYXRmb3JtQ29tcGFyaXNvbkNhcmRzKHdyYXAsIHJlc3VsdCwgbGFiZWxBLCBsYWJlbEIpOwogIH0KCiAgLyoqCiAgICogIkFsbCBQbGF0Zm9ybXMiIHJlcG9ydCDigJQgdGhlIGhlYWRsaW5lIENvbXBhcmlzb25zIHZpZXcuIFVubGlrZSB0aGUKICAgKiB3ZWVrL2N1c3RvbS9tb250aC9xdWFydGVyL3l0ZCB0b29scyBhYm92ZSwgdGhpcyBpZ25vcmVzIHRoZSBzaGFyZWQKICAgKiBwbGF0Zm9ybS9jYW1wYWlnbi9jb250ZW50LXR5cGUgZmlsdGVyIGJhciBlbnRpcmVseSBhbmQgbmVlZHMgbm8gZGF0ZQogICAqIHJhbmdlOiBpdCBhbHdheXMgY292ZXJzIGV2ZXJ5IHBsYXRmb3JtIHdpdGggYW55IGRhdGEgKHVwbG9hZGVkIHBvc3RzCiAgICogYW5kL29yIG1hbnVhbGx5LWVudGVyZWQgRm9sbG93ZXJzIERhdGEgUmVjb3JkIGhpc3RvcnkpLgogICAqLwogIGFzeW5jIGZ1bmN0aW9uIGxvYWRQbGF0Zm9ybVJlcG9ydChwYXJhbXMpIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnY29tcGFyaXNvblJlc3VsdHMnKTsKICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CiAgICB3cmFwLmFwcGVuZENoaWxkKHNrZWxldG9uU3RhdEdyaWQoMikpOwogICAgd3JhcC5hcHBlbmRDaGlsZChza2VsZXRvbkNoYXJ0KCkpOwogICAgY29uc3QgaGFzRXhwbGljaXRSYW5nZSA9IHBhcmFtcyAmJiBwYXJhbXMuZGF0ZUZyb20gJiYgcGFyYW1zLmRhdGVUbzsKICAgIGNvbnN0IHJlcG9ydCA9IGF3YWl0IEFwaS5wbGF0Zm9ybVJlcG9ydChoYXNFeHBsaWNpdFJhbmdlID8gcGFyYW1zIDoge30pOwogICAgcmVuZGVyUGxhdGZvcm1SZXBvcnQocmVwb3J0KTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlclBsYXRmb3JtUmVwb3J0KHJlcG9ydCkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdjb21wYXJpc29uUmVzdWx0cycpOwogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKCiAgICBpZiAoIXJlcG9ydC5wbGF0Zm9ybXMubGVuZ3RoKSB7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQoZW1wdHlTdGF0ZSh7CiAgICAgICAgaWNvbjogJ2dpdC1jb21wYXJlJywKICAgICAgICB0aXRsZTogJ05vIHBsYXRmb3JtIGRhdGEgeWV0JywKICAgICAgICBtZXNzYWdlOiAnVXBsb2FkIHBvc3RzIG9yIGFkZCBGb2xsb3dlcnMgRGF0YSBSZWNvcmQgZW50cmllcyB0byBzZWUgYSBjcm9zcy1wbGF0Zm9ybSBjb21wYXJpc29uIGhlcmUuJywKICAgICAgfSkpOwogICAgICByZXR1cm47CiAgICB9CgogICAgY29uc3QgcmFuZ2VMYWJlbCA9IHJlcG9ydC5yYW5nZS5pc0V4cGxpY2l0CiAgICAgID8gYCR7Rm9ybWF0LmRhdGUocmVwb3J0LnJhbmdlLmZyb20pfSDigJMgJHtGb3JtYXQuZGF0ZShyZXBvcnQucmFuZ2UudG8pfWAKICAgICAgOiBgQWxsIHRpbWUgKCR7Rm9ybWF0LmRhdGUocmVwb3J0LnJhbmdlLmZyb20pfSDigJMgJHtGb3JtYXQuZGF0ZShyZXBvcnQucmFuZ2UudG8pfSlgOwogICAgd3JhcC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIGBQbGF0Zm9ybSBDb21wYXJpc29uIFJlcG9ydCDigJQgJHtyYW5nZUxhYmVsfWAsICdzZWN0aW9uLXRpdGxlJykpOwoKICAgIGNvbnN0IGJlc3RQID0gcmVwb3J0LnBsYXRmb3Jtcy5maW5kKChwKSA9PiBwLnBsYXRmb3JtID09PSByZXBvcnQuYmVzdFBsYXRmb3JtKTsKICAgIGNvbnN0IHdvcnN0UCA9IHJlcG9ydC5wbGF0Zm9ybXMuZmluZCgocCkgPT4gcC5wbGF0Zm9ybSA9PT0gcmVwb3J0LndvcnN0UGxhdGZvcm0pOwogICAgaWYgKGJlc3RQIHx8IHdvcnN0UCkgewogICAgICBjb25zdCBoaWdobGlnaHRHcmlkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIGhpZ2hsaWdodEdyaWQuY2xhc3NOYW1lID0gJ3N0YXQtZ3JpZCc7CiAgICAgIGlmIChiZXN0UCkgewogICAgICAgIGNvbnN0IHRpbGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgICB0aWxlLmNsYXNzTmFtZSA9ICdzdGF0LXRpbGUnOwogICAgICAgIHRpbGUuYXBwZW5kKAogICAgICAgICAgdGV4dEVsKCdkaXYnLCAnQmVzdC1QZXJmb3JtaW5nIFBsYXRmb3JtJywgJ3N0YXQtbGFiZWwnKSwKICAgICAgICAgIHRleHRFbCgnZGl2JywgYmVzdFAubGFiZWwsICdzdGF0LXZhbHVlJyksCiAgICAgICAgICB0ZXh0RWwoJ2RpdicsIGBSZWFjaCAke0Zvcm1hdC5zbWFydChiZXN0UC50b3RhbHMucmVhY2gpfSDCtyBFbmdhZ2VtZW50ICR7Rm9ybWF0LnNtYXJ0KGJlc3RQLnRvdGFscy5lbmdhZ2VtZW50KX1gLCAncG9zdC1tZXRhJykKICAgICAgICApOwogICAgICAgIGhpZ2hsaWdodEdyaWQuYXBwZW5kQ2hpbGQodGlsZSk7CiAgICAgIH0KICAgICAgaWYgKHdvcnN0UCkgewogICAgICAgIGNvbnN0IHRpbGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgICB0aWxlLmNsYXNzTmFtZSA9ICdzdGF0LXRpbGUnOwogICAgICAgIHRpbGUuYXBwZW5kKAogICAgICAgICAgdGV4dEVsKCdkaXYnLCAnTG93ZXN0LVBlcmZvcm1pbmcgUGxhdGZvcm0nLCAnc3RhdC1sYWJlbCcpLAogICAgICAgICAgdGV4dEVsKCdkaXYnLCB3b3JzdFAubGFiZWwsICdzdGF0LXZhbHVlJyksCiAgICAgICAgICB0ZXh0RWwoJ2RpdicsIGBSZWFjaCAke0Zvcm1hdC5zbWFydCh3b3JzdFAudG90YWxzLnJlYWNoKX0gwrcgRW5nYWdlbWVudCAke0Zvcm1hdC5zbWFydCh3b3JzdFAudG90YWxzLmVuZ2FnZW1lbnQpfWAsICdwb3N0LW1ldGEnKQogICAgICAgICk7CiAgICAgICAgaGlnaGxpZ2h0R3JpZC5hcHBlbmRDaGlsZCh0aWxlKTsKICAgICAgfQogICAgICB3cmFwLmFwcGVuZENoaWxkKGhpZ2hsaWdodEdyaWQpOwogICAgfQoKICAgIGNvbnN0IHRhYmxlQ2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGFibGVDYXJkLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIHRhYmxlQ2FyZC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2gzJywgJ1BsYXRmb3JtIFJhbmtpbmcnKSk7CiAgICBjb25zdCB0YWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICB0YWJsZS5jbGFzc05hbWUgPSAnZGF0YS10YWJsZSc7CiAgICBjb25zdCB0aGVhZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RoZWFkJyk7CiAgICBjb25zdCBoZWFkUm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgIFsnUmFuaycsICdQbGF0Zm9ybScsICdQb3N0cycsICdSZWFjaCcsICdFbmdhZ2VtZW50JywgJ0ltcHJlc3Npb25zJywgJ0ZvbGxvd2VyIEdyb3d0aCddLmZvckVhY2goKGxhYmVsLCBpKSA9PiB7CiAgICAgIGNvbnN0IHRoID0gdGV4dEVsKCd0aCcsIGxhYmVsKTsKICAgICAgaWYgKGkgPj0gMikgdGguY2xhc3NMaXN0LmFkZCgnbnVtJyk7CiAgICAgIGhlYWRSb3cuYXBwZW5kQ2hpbGQodGgpOwogICAgfSk7CiAgICB0aGVhZC5hcHBlbmRDaGlsZChoZWFkUm93KTsKICAgIHRhYmxlLmFwcGVuZENoaWxkKHRoZWFkKTsKICAgIGNvbnN0IHRib2R5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGJvZHknKTsKICAgIHJlcG9ydC5wbGF0Zm9ybXMuZm9yRWFjaCgocCkgPT4gewogICAgICBjb25zdCB0ciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAgIHRyLmFwcGVuZENoaWxkKHRleHRFbCgndGQnLCBwLm92ZXJhbGxSYW5rID8gYCMke3Aub3ZlcmFsbFJhbmt9YCA6ICfigJQnKSk7CiAgICAgIGNvbnN0IHBsYXRUZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIGNvbnN0IHBpbGwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7IHBpbGwuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLXBpbGwnOwogICAgICBjb25zdCBkb3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7IGRvdC5jbGFzc05hbWUgPSAncGxhdGZvcm0tZG90JzsgZG90LnN0eWxlLmJhY2tncm91bmQgPSBwLmNvbG9yOwogICAgICBwaWxsLmFwcGVuZChkb3QsIGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKHAubGFiZWwpKTsKICAgICAgcGxhdFRkLmFwcGVuZENoaWxkKHBpbGwpOwogICAgICB0ci5hcHBlbmRDaGlsZChwbGF0VGQpOwogICAgICB0ci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3RkJywgRm9ybWF0Lm51bWJlcihwLnBvc3RDb3VudCksICdudW0nKSk7CiAgICAgIHRyLmFwcGVuZENoaWxkKHRleHRFbCgndGQnLCBGb3JtYXQuc21hcnQocC50b3RhbHMucmVhY2gpLCAnbnVtJykpOwogICAgICB0ci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3RkJywgRm9ybWF0LnNtYXJ0KHAudG90YWxzLmVuZ2FnZW1lbnQpLCAnbnVtJykpOwogICAgICB0ci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3RkJywgRm9ybWF0LnNtYXJ0KHAudG90YWxzLmltcHJlc3Npb25zKSwgJ251bScpKTsKICAgICAgY29uc3QgZm9sbG93ZXJUZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIGZvbGxvd2VyVGQuY2xhc3NOYW1lID0gJ251bSc7CiAgICAgIGlmIChwLmZvbGxvd2Vycy5jaGFuZ2UgPT09IG51bGwpIHsKICAgICAgICBmb2xsb3dlclRkLmFwcGVuZENoaWxkKGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKHAuZm9sbG93ZXJzLmxhdGVzdCAhPT0gbnVsbCA/IEZvcm1hdC5udW1iZXIocC5mb2xsb3dlcnMubGF0ZXN0KSA6ICfigJQnKSk7CiAgICAgIH0gZWxzZSB7CiAgICAgICAgY29uc3QgZm9sbG93ZXJXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICAgIGZvbGxvd2VyV3JhcC5zdHlsZS5kaXNwbGF5ID0gJ2lubGluZS1mbGV4JzsKICAgICAgICBmb2xsb3dlcldyYXAuc3R5bGUuYWxpZ25JdGVtcyA9ICdjZW50ZXInOwogICAgICAgIGZvbGxvd2VyV3JhcC5zdHlsZS5nYXAgPSAnNnB4JzsKICAgICAgICBmb2xsb3dlcldyYXAuYXBwZW5kQ2hpbGQodGV4dEVsKCdzcGFuJywgYCR7cC5mb2xsb3dlcnMuY2hhbmdlID4gMCA/ICcrJyA6ICcnfSR7Rm9ybWF0Lm51bWJlcihwLmZvbGxvd2Vycy5jaGFuZ2UpfWAsIGBzdGF0LWRlbHRhICR7Rm9ybWF0LmRlbHRhQ2xhc3MocC5mb2xsb3dlcnMuY2hhbmdlKX1gKSk7CiAgICAgICAgaWYgKHAuZm9sbG93ZXJzLmNoYW5nZVBjdCAhPT0gbnVsbCkgZm9sbG93ZXJXcmFwLmFwcGVuZENoaWxkKHRleHRFbCgnc3BhbicsIGAoJHtGb3JtYXQucGN0KHAuZm9sbG93ZXJzLmNoYW5nZVBjdCl9KWAsICdwb3N0LW1ldGEnKSk7CiAgICAgICAgZm9sbG93ZXJUZC5hcHBlbmRDaGlsZChmb2xsb3dlcldyYXApOwogICAgICB9CiAgICAgIHRyLmFwcGVuZENoaWxkKGZvbGxvd2VyVGQpOwogICAgICB0Ym9keS5hcHBlbmRDaGlsZCh0cik7CiAgICB9KTsKICAgIHRhYmxlLmFwcGVuZENoaWxkKHRib2R5KTsKICAgIGNvbnN0IHRhYmxlU2Nyb2xsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0YWJsZVNjcm9sbC5jbGFzc05hbWUgPSAndGFibGUtc2Nyb2xsJzsKICAgIHRhYmxlU2Nyb2xsLmFwcGVuZENoaWxkKHRhYmxlKTsKICAgIHRhYmxlQ2FyZC5hcHBlbmRDaGlsZCh0YWJsZVNjcm9sbCk7CiAgICB3cmFwLmFwcGVuZENoaWxkKHRhYmxlQ2FyZCk7CgogICAgY29uc3QgY2hhcnRDYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjaGFydENhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgY29uc3QgY2hhcnRIZWFkZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNoYXJ0SGVhZGVyLmNsYXNzTmFtZSA9ICdjYXJkLWhlYWRlcic7CiAgICBjaGFydEhlYWRlci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2gzJywgJ01ldHJpYyBDb21wYXJpc29uJykpOwogICAgY29uc3QgbWV0cmljU2VsZWN0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc2VsZWN0Jyk7CiAgICBNRVRSSUNfUk9XUy5mb3JFYWNoKChtKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOyBvcHQudmFsdWUgPSBtLmtleTsgb3B0LnRleHRDb250ZW50ID0gbS5sYWJlbDsKICAgICAgaWYgKG0ua2V5ID09PSBwbGF0Zm9ybUNoYXJ0TWV0cmljKSBvcHQuc2VsZWN0ZWQgPSB0cnVlOwogICAgICBtZXRyaWNTZWxlY3QuYXBwZW5kQ2hpbGQob3B0KTsKICAgIH0pOwogICAgbWV0cmljU2VsZWN0LmFkZEV2ZW50TGlzdGVuZXIoJ2NoYW5nZScsICgpID0+IHsKICAgICAgcGxhdGZvcm1DaGFydE1ldHJpYyA9IG1ldHJpY1NlbGVjdC52YWx1ZTsKICAgICAgZHJhd1BsYXRmb3JtUmVwb3J0Q2hhcnQocmVwb3J0KTsKICAgIH0pOwogICAgY2hhcnRIZWFkZXIuYXBwZW5kQ2hpbGQobWV0cmljU2VsZWN0KTsKICAgIGNvbnN0IGNoYXJ0V3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY2hhcnRXcmFwLmNsYXNzTmFtZSA9ICdjaGFydC13cmFwIHRhbGwnOwogICAgY2hhcnRXcmFwLmlubmVySFRNTCA9ICc8Y2FudmFzIGlkPSJwbGF0Zm9ybVJlcG9ydENhbnZhcyI+PC9jYW52YXM+JzsKICAgIGNoYXJ0Q2FyZC5hcHBlbmQoY2hhcnRIZWFkZXIsIGNoYXJ0V3JhcCk7CiAgICB3cmFwLmFwcGVuZENoaWxkKGNoYXJ0Q2FyZCk7CiAgICBkcmF3UGxhdGZvcm1SZXBvcnRDaGFydChyZXBvcnQpOwoKICAgIGNvbnN0IHdpdGhGb2xsb3dlcnMgPSByZXBvcnQucGxhdGZvcm1zLmZpbHRlcigocCkgPT4gcC5mb2xsb3dlcnMubGF0ZXN0ICE9PSBudWxsKTsKICAgIGlmICh3aXRoRm9sbG93ZXJzLmxlbmd0aCkgewogICAgICBjb25zdCBmb2xsb3dlckNhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgZm9sbG93ZXJDYXJkLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgICAgZm9sbG93ZXJDYXJkLmFwcGVuZENoaWxkKHRleHRFbCgnaDMnLCAnRm9sbG93ZXIgR3Jvd3RoIGJ5IFBsYXRmb3JtJykpOwogICAgICBjb25zdCBmQ2hhcnRXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIGZDaGFydFdyYXAuY2xhc3NOYW1lID0gJ2NoYXJ0LXdyYXAgdGFsbCc7CiAgICAgIGZDaGFydFdyYXAuaW5uZXJIVE1MID0gJzxjYW52YXMgaWQ9InBsYXRmb3JtRm9sbG93ZXJDYW52YXMiPjwvY2FudmFzPic7CiAgICAgIGZvbGxvd2VyQ2FyZC5hcHBlbmRDaGlsZChmQ2hhcnRXcmFwKTsKICAgICAgd3JhcC5hcHBlbmRDaGlsZChmb2xsb3dlckNhcmQpOwogICAgICBDaGFydHMucGxhdGZvcm1CYXJDaGFydCgncGxhdGZvcm1Gb2xsb3dlckNhbnZhcycsIHsKICAgICAgICBsYWJlbHM6IHdpdGhGb2xsb3dlcnMubWFwKChwKSA9PiBwLmxhYmVsKSwKICAgICAgICBkYXRhOiB3aXRoRm9sbG93ZXJzLm1hcCgocCkgPT4gcC5mb2xsb3dlcnMubGF0ZXN0IHx8IDApLAogICAgICAgIGNvbG9yczogd2l0aEZvbGxvd2Vycy5tYXAoKHApID0+IHAuY29sb3IpLAogICAgICAgIGZvcm1hdFZhbHVlOiAodikgPT4gRm9ybWF0LnNtYXJ0KHYpLAogICAgICB9KTsKICAgIH0KCiAgICBjb25zdCBpbnNpZ2h0c0NhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGluc2lnaHRzQ2FyZC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICBpbnNpZ2h0c0NhcmQuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMycsICdJbnNpZ2h0cyAmIFN1bW1hcnknKSk7CiAgICBjb25zdCBsaXN0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndWwnKTsKICAgIGxpc3QuY2xhc3NOYW1lID0gJ2luc2lnaHRzLWxpc3QnOwogICAgcmVwb3J0Lmluc2lnaHRzLmZvckVhY2goKGxpbmUpID0+IGxpc3QuYXBwZW5kQ2hpbGQodGV4dEVsKCdsaScsIGxpbmUpKSk7CiAgICBpbnNpZ2h0c0NhcmQuYXBwZW5kQ2hpbGQobGlzdCk7CiAgICB3cmFwLmFwcGVuZENoaWxkKGluc2lnaHRzQ2FyZCk7CiAgfQoKICBmdW5jdGlvbiBkcmF3UGxhdGZvcm1SZXBvcnRDaGFydChyZXBvcnQpIHsKICAgIENoYXJ0cy5wbGF0Zm9ybUJhckNoYXJ0KCdwbGF0Zm9ybVJlcG9ydENhbnZhcycsIHsKICAgICAgbGFiZWxzOiByZXBvcnQucGxhdGZvcm1zLm1hcCgocCkgPT4gcC5sYWJlbCksCiAgICAgIGRhdGE6IHJlcG9ydC5wbGF0Zm9ybXMubWFwKChwKSA9PiBwLnRvdGFsc1twbGF0Zm9ybUNoYXJ0TWV0cmljXSB8fCAwKSwKICAgICAgY29sb3JzOiByZXBvcnQucGxhdGZvcm1zLm1hcCgocCkgPT4gcC5jb2xvciksCiAgICAgIGZvcm1hdFZhbHVlOiAodikgPT4gKHBsYXRmb3JtQ2hhcnRNZXRyaWMgPT09ICd3YXRjaF90aW1lX3NlY29uZHMnID8gRm9ybWF0LmR1cmF0aW9uKHYpIDogRm9ybWF0LnNtYXJ0KHYpKSwKICAgIH0pOwogIH0KCiAgLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICAgIFBsYXRmb3JtIFBlcmZvcm1hbmNlIENvbXBhcmlzb24g4oCUIHJlcGxhY2VzIHRoZSBvbGQgZ3JvdXBlZAogICAgICJSYW5nZSBBIHZzIFJhbmdlIEIgYnkgUGxhdGZvcm0iIGNoYXJ0LiBPbmUgY2FyZCBwZXIgcGxhdGZvcm0KICAgICB3aXRoIGFueSBkYXRhIGluIGVpdGhlciByYW5nZSwgYnVpbHQgZW50aXJlbHkgZnJvbSB0aGUgc2FtZQogICAgIGNvbXBhcmVSYW5nZXMoKSByZXNwb25zZSB0aGUgc3RhdC10aWxlIGdyaWQgYWJvdmUgYWxyZWFkeQogICAgIHVzZXMgKHJlc3VsdC5yYW5nZUEucGxhdGZvcm1zIC8gcmVzdWx0LnJhbmdlQi5wbGF0Zm9ybXMpIOKAlCBubwogICAgIGV4dHJhIGZldGNoLgogICAgID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KICBjb25zdCBBTExfQ0FSRF9NRVRSSUNTID0gW3sga2V5OiAncG9zdF9jb3VudCcsIGxhYmVsOiAnUG9zdHMnIH0sIC4uLk1FVFJJQ19ST1dTXTsKICBjb25zdCBDQVJEX1NPUlRfTU9ERVMgPSBbCiAgICB7IGtleTogJ292ZXJhbGwnLCBsYWJlbDogJ092ZXJhbGwgUGVyZm9ybWFuY2UnIH0sCiAgICB7IGtleTogJ2dyb3d0aCcsIGxhYmVsOiAnSGlnaGVzdCBHcm93dGgnIH0sCiAgICB7IGtleTogJ2VuZ2FnZW1lbnQnLCBsYWJlbDogJ0hpZ2hlc3QgRW5nYWdlbWVudCcgfSwKICAgIHsga2V5OiAnZm9sbG93ZXJzJywgbGFiZWw6ICdNb3N0IEZvbGxvd2VycycgfSwKICAgIHsga2V5OiAncG9zdHMnLCBsYWJlbDogJ01vc3QgUG9zdHMnIH0sCiAgICB7IGtleTogJ2FscGhhJywgbGFiZWw6ICdBbHBoYWJldGljYWwnIH0sCiAgXTsKCiAgLyoqIFBlci1tZXRyaWMge2EsIGIsIGRpZmYsIHBjdERpZmZ9IGFjcm9zcyBib3RoIHJhbmdlcyBmb3Igb25lIHBsYXRmb3JtLCBza2lwcGluZyBhbnkgbWV0cmljIHRoYXQncyB6ZXJvIGluIGJvdGgg4oCUIGEgcGxhdGZvcm0ncyBjYXJkIHNob3VsZCBvbmx5IGV2ZXIgc2hvdyBtZXRyaWNzIGl0IGFjdHVhbGx5IGhhcy4gKi8KICBmdW5jdGlvbiBjb21wdXRlQ2FyZE1ldHJpY3MocGxhdGZvcm1BLCBwbGF0Zm9ybUIpIHsKICAgIGNvbnN0IG1ldHJpY3MgPSBbXTsKICAgIEFMTF9DQVJEX01FVFJJQ1MuZm9yRWFjaCgoeyBrZXksIGxhYmVsIH0pID0+IHsKICAgICAgY29uc3QgYSA9IChwbGF0Zm9ybUEgJiYgcGxhdGZvcm1BW2tleV0pIHx8IDA7CiAgICAgIGNvbnN0IGIgPSAocGxhdGZvcm1CICYmIHBsYXRmb3JtQltrZXldKSB8fCAwOwogICAgICBpZiAoYSA9PT0gMCAmJiBiID09PSAwKSByZXR1cm47CiAgICAgIGNvbnN0IGRpZmYgPSBhIC0gYjsKICAgICAgY29uc3QgcGN0RGlmZiA9IGIgPyBNYXRoLnJvdW5kKChkaWZmIC8gYikgKiAxMDAwKSAvIDEwIDogKGEgPiAwID8gbnVsbCA6IDApOwogICAgICBtZXRyaWNzLnB1c2goeyBrZXksIGxhYmVsLCBhLCBiLCBkaWZmLCBwY3REaWZmLCBpc0R1cmF0aW9uOiBrZXkgPT09ICd3YXRjaF90aW1lX3NlY29uZHMnIH0pOwogICAgfSk7CiAgICByZXR1cm4gbWV0cmljczsKICB9CgogIC8qKiBBIHNpbmdsZSAiaG93IGRpZCB0aGlzIHBsYXRmb3JtIGRvIG92ZXJhbGwiIG51bWJlcjogdGhlIGF2ZXJhZ2UgJSBjaGFuZ2UgYWNyb3NzIGV2ZXJ5IG1ldHJpYyB0aGF0IGhhcyBhIGNvbXB1dGFibGUgcGVyY2VudGFnZSAoYSBtZXRyaWMgZ29pbmcgZnJvbSAwIHRvIHNvbWV0aGluZyBoYXMgbm8gcGVyY2VudGFnZSDigJQgIm5ldyIsIG5vdCBjb3VudGVkIGVpdGhlciB3YXkpLiAqLwogIGZ1bmN0aW9uIG92ZXJhbGxQY3RDaGFuZ2UobWV0cmljcykgewogICAgY29uc3Qgd2l0aFBjdCA9IG1ldHJpY3MuZmlsdGVyKChtKSA9PiBtLnBjdERpZmYgIT09IG51bGwpOwogICAgaWYgKCF3aXRoUGN0Lmxlbmd0aCkgcmV0dXJuIG51bGw7CiAgICByZXR1cm4gTWF0aC5yb3VuZCgod2l0aFBjdC5yZWR1Y2UoKHN1bSwgbSkgPT4gc3VtICsgbS5wY3REaWZmLCAwKSAvIHdpdGhQY3QubGVuZ3RoKSAqIDEwKSAvIDEwOwogIH0KCiAgZnVuY3Rpb24gYmVzdFdlYWtlc3RNZXRyaWMobWV0cmljcykgewogICAgY29uc3Qgd2l0aFBjdCA9IG1ldHJpY3MuZmlsdGVyKChtKSA9PiBtLnBjdERpZmYgIT09IG51bGwpOwogICAgaWYgKCF3aXRoUGN0Lmxlbmd0aCkgcmV0dXJuIHsgYmVzdDogbnVsbCwgd2Vha2VzdDogbnVsbCB9OwogICAgY29uc3QgYmVzdCA9IHdpdGhQY3QucmVkdWNlKChhLCBiKSA9PiAoYi5wY3REaWZmID4gYS5wY3REaWZmID8gYiA6IGEpKTsKICAgIGNvbnN0IHdlYWtlc3QgPSB3aXRoUGN0LnJlZHVjZSgoYSwgYikgPT4gKGIucGN0RGlmZiA8IGEucGN0RGlmZiA/IGIgOiBhKSk7CiAgICByZXR1cm4geyBiZXN0LCB3ZWFrZXN0IH07CiAgfQoKICBmdW5jdGlvbiB0cmVuZERpcmVjdGlvbihwY3QpIHsKICAgIGlmIChwY3QgPT09IG51bGwgfHwgcGN0ID09PSB1bmRlZmluZWQpIHJldHVybiAnZmxhdCc7CiAgICBpZiAocGN0ID4gMC41KSByZXR1cm4gJ3VwJzsKICAgIGlmIChwY3QgPCAtMC41KSByZXR1cm4gJ2Rvd24nOwogICAgcmV0dXJuICdmbGF0JzsKICB9CgogIGZ1bmN0aW9uIGJ1aWxkUGxhdGZvcm1DYXJkcyhyZXN1bHQpIHsKICAgIGNvbnN0IHBsYXRmb3JtT3B0aW9ucyA9ICh3aW5kb3cuX19maWx0ZXJPcHRpb25zQ2FjaGUgfHwgeyBhbGxQbGF0Zm9ybXM6IFtdIH0pLmFsbFBsYXRmb3JtczsKICAgIGNvbnN0IGlkcyA9IFsuLi5uZXcgU2V0KFsuLi5yZXN1bHQucmFuZ2VBLnBsYXRmb3JtcywgLi4ucmVzdWx0LnJhbmdlQi5wbGF0Zm9ybXNdLm1hcCgocCkgPT4gcC5wbGF0Zm9ybSkpXTsKICAgIGNvbnN0IGJ5SWRBID0gT2JqZWN0LmZyb21FbnRyaWVzKHJlc3VsdC5yYW5nZUEucGxhdGZvcm1zLm1hcCgocCkgPT4gW3AucGxhdGZvcm0sIHBdKSk7CiAgICBjb25zdCBieUlkQiA9IE9iamVjdC5mcm9tRW50cmllcyhyZXN1bHQucmFuZ2VCLnBsYXRmb3Jtcy5tYXAoKHApID0+IFtwLnBsYXRmb3JtLCBwXSkpOwoKICAgIHJldHVybiBpZHMKICAgICAgLm1hcCgoaWQpID0+IHsKICAgICAgICBjb25zdCBtZXRhID0gcGxhdGZvcm1PcHRpb25zLmZpbmQoKHApID0+IHAuaWQgPT09IGlkKSB8fCB7IGlkLCBsYWJlbDogaWQsIGNvbG9yOiAndmFyKC0tc2VyaWVzLTEpJyB9OwogICAgICAgIGNvbnN0IGEgPSBieUlkQVtpZF0gfHwgbnVsbDsKICAgICAgICBjb25zdCBiID0gYnlJZEJbaWRdIHx8IG51bGw7CiAgICAgICAgY29uc3QgbWV0cmljcyA9IGNvbXB1dGVDYXJkTWV0cmljcyhhLCBiKTsKICAgICAgICBjb25zdCB7IGJlc3QsIHdlYWtlc3QgfSA9IGJlc3RXZWFrZXN0TWV0cmljKG1ldHJpY3MpOwogICAgICAgIHJldHVybiB7CiAgICAgICAgICBwbGF0Zm9ybTogaWQsCiAgICAgICAgICBsYWJlbDogbWV0YS5sYWJlbCwKICAgICAgICAgIGNvbG9yOiBtZXRhLmNvbG9yLAogICAgICAgICAgbWV0cmljcywKICAgICAgICAgIG92ZXJhbGw6IG92ZXJhbGxQY3RDaGFuZ2UobWV0cmljcyksCiAgICAgICAgICBiZXN0LAogICAgICAgICAgd2Vha2VzdCwKICAgICAgICAgIGZvbGxvd2Vyc0dhaW5lZDogKGEgPyBhLmZvbGxvd2Vyc19nYWluZWQgfHwgMCA6IDApICsgKGIgPyBiLmZvbGxvd2Vyc19nYWluZWQgfHwgMCA6IDApLAogICAgICAgICAgcG9zdHM6IChhID8gYS5wb3N0X2NvdW50IHx8IDAgOiAwKSArIChiID8gYi5wb3N0X2NvdW50IHx8IDAgOiAwKSwKICAgICAgICAgIGVuZ2FnZW1lbnRUb3RhbDogKGEgPyBhLmVuZ2FnZW1lbnQgfHwgMCA6IDApICsgKGIgPyBiLmVuZ2FnZW1lbnQgfHwgMCA6IDApLAogICAgICAgIH07CiAgICAgIH0pCiAgICAgIC5maWx0ZXIoKGNhcmQpID0+IGNhcmQubWV0cmljcy5sZW5ndGggPiAwKTsKICB9CgogIGZ1bmN0aW9uIHNvcnRDYXJkcyhjYXJkcywgc29ydE1vZGUpIHsKICAgIGNvbnN0IGFyciA9IFsuLi5jYXJkc107CiAgICBpZiAoc29ydE1vZGUgPT09ICdlbmdhZ2VtZW50JykgcmV0dXJuIGFyci5zb3J0KCh4LCB5KSA9PiB5LmVuZ2FnZW1lbnRUb3RhbCAtIHguZW5nYWdlbWVudFRvdGFsKTsKICAgIGlmIChzb3J0TW9kZSA9PT0gJ2ZvbGxvd2VycycpIHJldHVybiBhcnIuc29ydCgoeCwgeSkgPT4geS5mb2xsb3dlcnNHYWluZWQgLSB4LmZvbGxvd2Vyc0dhaW5lZCk7CiAgICBpZiAoc29ydE1vZGUgPT09ICdwb3N0cycpIHJldHVybiBhcnIuc29ydCgoeCwgeSkgPT4geS5wb3N0cyAtIHgucG9zdHMpOwogICAgaWYgKHNvcnRNb2RlID09PSAnYWxwaGEnKSByZXR1cm4gYXJyLnNvcnQoKHgsIHkpID0+IHgubGFiZWwubG9jYWxlQ29tcGFyZSh5LmxhYmVsKSk7CiAgICAvLyAnb3ZlcmFsbCcgYW5kICdncm93dGgnIGJvdGggcmFuayBieSB0aGUgc2FtZSBjb21wb3NpdGUgJSBjaGFuZ2Ug4oCUIHRoZSB0d28gbGFiZWxzCiAgICAvLyByZWFkIGRpZmZlcmVudGx5IG9uIHRoZSBzYW1lIHVuZGVybHlpbmcgbnVtYmVyLCBwZXIgdGhlIHJlcXVlc3RlZCBvcHRpb24gbGlzdC4KICAgIHJldHVybiBhcnIuc29ydCgoeCwgeSkgPT4gKHkub3ZlcmFsbCA/PyAtSW5maW5pdHkpIC0gKHgub3ZlcmFsbCA/PyAtSW5maW5pdHkpKTsKICB9CgogIGZ1bmN0aW9uIGJ1aWxkTWV0cmljUm93KG0pIHsKICAgIGNvbnN0IHJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcm93LmNsYXNzTmFtZSA9ICdwY2MtbWV0cmljLXJvdyc7CiAgICBjb25zdCBoZWFkZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGhlYWRlci5jbGFzc05hbWUgPSAncGNjLW1ldHJpYy1oZWFkZXInOwogICAgY29uc3QgZm10ID0gKHYpID0+IChtLmlzRHVyYXRpb24gPyBGb3JtYXQuZHVyYXRpb24odikgOiBGb3JtYXQuc21hcnQodikpOwogICAgY29uc3QgZGlmZlRleHQgPSBtLnBjdERpZmYgPT09IG51bGwKICAgICAgPyBgJHttLmRpZmYgPiAwID8gJysnIDogJyd9JHtmbXQobS5kaWZmKX0gKG5ldylgCiAgICAgIDogYCR7bS5kaWZmID4gMCA/ICcrJyA6ICcnfSR7Zm10KG0uZGlmZil9ICgke0Zvcm1hdC5wY3QobS5wY3REaWZmKX0pYDsKICAgIGhlYWRlci5hcHBlbmQoCiAgICAgIHRleHRFbCgnc3BhbicsIG0ubGFiZWwsICdwY2MtbWV0cmljLWxhYmVsJyksCiAgICAgIHRleHRFbCgnc3BhbicsIGRpZmZUZXh0LCBgcGNjLW1ldHJpYy1kaWZmICR7Rm9ybWF0LmRlbHRhQ2xhc3MobS5wY3REaWZmKX1gKQogICAgKTsKICAgIHJvdy5hcHBlbmRDaGlsZChoZWFkZXIpOwogICAgY29uc3QgbWF4ID0gTWF0aC5tYXgobS5hLCBtLmIsIDEpOwogICAgLy8gbS5hIGlzIGFsd2F5cyB0aGUgY3VycmVudCBwZXJpb2QgYW5kIG0uYiBhbHdheXMgdGhlIHByZXZpb3VzIHBlcmlvZCAoc2VlCiAgICAvLyBjb21wdXRlQ2FyZE1ldHJpY3MpIHJlZ2FyZGxlc3Mgb2YgY29tcGFyaXNvbiBtb2RlLCBzbyB0aGVzZSBsYWJlbHMgY2FuIGJlCiAgICAvLyBoYXJkY29kZWQgcmF0aGVyIHRoYW4gbmVlZGluZyB0aGUgbW9kZS1zcGVjaWZpYyBsYWJlbEEvbGFiZWxCIHRleHQg4oCUIHVubGlrZQogICAgLy8gdGhlIGdlbmVyaWMgIlJhbmdlIEEiLyJSYW5nZSBCIiB3b3JkaW5nIHRoaXMgcmVwbGFjZWQsIHdoaWNoIHJlYWQgYXMKICAgIC8vIGFyYml0cmFyeSBsZXR0ZXJzIHdpdGggbm8gaW5kaWNhdGlvbiBvZiB3aGljaCBzaWRlIHdhcyBtb3JlIHJlY2VudC4KICAgIHJvdy5hcHBlbmRDaGlsZChidWlsZEJhcih7IGxhYmVsOiAnQ3VycmVudCcsIHZhbHVlOiBtLmEsIG1heCwgY29sb3JWYXI6ICctLXNlcmllcy0xJywgZm9ybWF0VmFsdWU6IGZtdCB9KSk7CiAgICByb3cuYXBwZW5kQ2hpbGQoYnVpbGRCYXIoeyBsYWJlbDogJ1ByZXZpb3VzJywgdmFsdWU6IG0uYiwgbWF4LCBjb2xvclZhcjogJy0tdGV4dC1tdXRlZCcsIGZvcm1hdFZhbHVlOiBmbXQgfSkpOwogICAgcmV0dXJuIHJvdzsKICB9CgogIGZ1bmN0aW9uIGJ1aWxkQ2FyZEZvb3RlcihjYXJkKSB7CiAgICBjb25zdCBmb290ZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGZvb3Rlci5jbGFzc05hbWUgPSAncGNjLWZvb3Rlcic7CiAgICBjb25zdCBkaXIgPSB0cmVuZERpcmVjdGlvbihjYXJkLm92ZXJhbGwpOwogICAgY29uc3QgcmVzdWx0VGV4dCA9IGNhcmQub3ZlcmFsbCA9PT0gbnVsbAogICAgICA/ICdOb3QgZW5vdWdoIGRhdGEgdG8gY29tcGFyZScKICAgICAgOiBgJHtkaXIgPT09ICd1cCcgPyAnSW1wcm92ZWQnIDogZGlyID09PSAnZG93bicgPyAnRGVjbGluZWQnIDogJ05vIHNpZ25pZmljYW50IGNoYW5nZSd9JHtkaXIgIT09ICdmbGF0JyA/IGAgYnkgJHtNYXRoLmFicyhjYXJkLm92ZXJhbGwpfSVgIDogJyd9YDsKICAgIGZvb3Rlci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsICdPdmVyYWxsIFJlc3VsdCcsICdwY2MtZm9vdGVyLWxhYmVsJykpOwogICAgZm9vdGVyLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgcmVzdWx0VGV4dCwgYHBjYy1mb290ZXItdmFsdWUgJHtkaXJ9YCkpOwogICAgaWYgKGNhcmQuYmVzdCkgZm9vdGVyLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgYEJlc3QgTWV0cmljOiAke2NhcmQuYmVzdC5sYWJlbH0gKCR7Rm9ybWF0LnBjdChjYXJkLmJlc3QucGN0RGlmZil9KWAsICdwY2MtZm9vdGVyLWRldGFpbCcpKTsKICAgIGlmIChjYXJkLndlYWtlc3QgJiYgY2FyZC53ZWFrZXN0ICE9PSBjYXJkLmJlc3QpIHsKICAgICAgZm9vdGVyLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgYFdlYWtlc3QgTWV0cmljOiAke2NhcmQud2Vha2VzdC5sYWJlbH0gKCR7Rm9ybWF0LnBjdChjYXJkLndlYWtlc3QucGN0RGlmZil9KWAsICdwY2MtZm9vdGVyLWRldGFpbCcpKTsKICAgIH0KICAgIHJldHVybiBmb290ZXI7CiAgfQoKICAvKiogU2VsZi1jb250YWluZWQgbW9kYWwgZm9yICJWaWV3IEZ1bGwgQ29tcGFyaXNvbiIg4oCUIGEgc2VwYXJhdGUgb3ZlcmxheSBpZCBmcm9tIHRoZSBEYXRhIFJlY29yZHMgRWRpdCBtb2RhbCAoUmVjb3Jkcy5tb2RhbFNoZWxsIGlzIGEgcHJpdmF0ZSBjbG9zdXJlIG9mIHRoYXQgbW9kdWxlLCBub3Qgc2hhcmVkIHN0YXRlKSwgc2FtZSB2aXN1YWwgbGFuZ3VhZ2UgKC5tb2RhbC1vdmVybGF5IC8gLm1vZGFsLXBhbmVsKSBzbyBpdCBsb29rcyBpZGVudGljYWwuICovCiAgZnVuY3Rpb24gY2xvc2VDYXJkTW9kYWwoKSB7CiAgICBjb25zdCBvdmVybGF5ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2NvbXBhcmlzb25Nb2RhbE92ZXJsYXknKTsKICAgIGlmIChvdmVybGF5KSBvdmVybGF5LnJlbW92ZSgpOwogIH0KCiAgZnVuY3Rpb24gb3BlbkNhcmRNb2RhbChjYXJkLCBsYWJlbEEsIGxhYmVsQikgewogICAgY2xvc2VDYXJkTW9kYWwoKTsKICAgIGNvbnN0IG92ZXJsYXkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIG92ZXJsYXkuY2xhc3NOYW1lID0gJ21vZGFsLW92ZXJsYXknOwogICAgb3ZlcmxheS5pZCA9ICdjb21wYXJpc29uTW9kYWxPdmVybGF5JzsKICAgIG92ZXJsYXkuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoZSkgPT4geyBpZiAoZS50YXJnZXQgPT09IG92ZXJsYXkpIGNsb3NlQ2FyZE1vZGFsKCk7IH0pOwogICAgY29uc3QgcGFuZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHBhbmVsLmNsYXNzTmFtZSA9ICdtb2RhbC1wYW5lbCB3aWRlJzsKICAgIHBhbmVsLmFwcGVuZENoaWxkKHRleHRFbCgnaDInLCBgJHtjYXJkLmxhYmVsfSDigJQgRnVsbCBDb21wYXJpc29uYCkpOwogICAgcGFuZWwuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBgJHtsYWJlbEF9IHZzICR7bGFiZWxCfWAsICdtb2RhbC1zdWInKSk7CgogICAgY29uc3QgdGFibGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0YWJsZScpOwogICAgdGFibGUuY2xhc3NOYW1lID0gJ2RhdGEtdGFibGUnOwogICAgY29uc3QgdGhlYWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0aGVhZCcpOwogICAgY29uc3QgaGVhZFJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAvLyBsYWJlbEEvbGFiZWxCIGFyZSB0aGUgc2FtZSBtb2RlLXNwZWNpZmljIHRleHQgc2hvd24gaW4gdGhlIG1vZGFsJ3Mgc3VidGl0bGUKICAgIC8vIGFib3ZlIChlLmcuICJXZWVrIG9mIEphbiAxMywgMjAyNiIgLyAiV2VlayBvZiBEZWMgMzAsIDIwMjUiLCBvciAiQ3VycmVudAogICAgLy8gcGVyaW9kIiAvICJQcmV2aW91cyBwZXJpb2QiKSDigJQgcmV1c2VkIGhlcmUgaW5zdGVhZCBvZiBnZW5lcmljICJSYW5nZSBBIi8KICAgIC8vICJSYW5nZSBCIiBzbyB0aGUgY29sdW1uIGhlYWRlcnMgYWx3YXlzIHNheSB3aGF0IHBlcmlvZCB0aGV5IGFjdHVhbGx5IGhvbGQuCiAgICBjb25zdCBudW1UaCA9ICh0ZXh0KSA9PiB7IGNvbnN0IHRoID0gdGV4dEVsKCd0aCcsIHRleHQpOyB0aC5jbGFzc0xpc3QuYWRkKCdudW0nKTsgcmV0dXJuIHRoOyB9OwogICAgaGVhZFJvdy5hcHBlbmQodGV4dEVsKCd0aCcsICdNZXRyaWMnKSwgbnVtVGgobGFiZWxBKSwgbnVtVGgobGFiZWxCKSwgbnVtVGgoJ0RpZmZlcmVuY2UnKSwgbnVtVGgoJyUgRGlmZmVyZW5jZScpLCB0ZXh0RWwoJ3RoJywgJ1RyZW5kJykpOwogICAgdGhlYWQuYXBwZW5kQ2hpbGQoaGVhZFJvdyk7CiAgICBjb25zdCB0Ym9keSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3Rib2R5Jyk7CiAgICBjYXJkLm1ldHJpY3MuZm9yRWFjaCgobSkgPT4gewogICAgICBjb25zdCBmbXQgPSAodikgPT4gKG0uaXNEdXJhdGlvbiA/IEZvcm1hdC5kdXJhdGlvbih2KSA6IEZvcm1hdC5zbWFydCh2KSk7CiAgICAgIGNvbnN0IHRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgY29uc3QgdHJlbmRFbCA9IHRleHRFbCgnc3BhbicsIHRyZW5kRGlyZWN0aW9uKG0ucGN0RGlmZikgPT09ICd1cCcgPyAn4payJyA6IHRyZW5kRGlyZWN0aW9uKG0ucGN0RGlmZikgPT09ICdkb3duJyA/ICfilrwnIDogJ+KAlCcsIGBzdGF0LWRlbHRhICR7Rm9ybWF0LmRlbHRhQ2xhc3MobS5wY3REaWZmKX1gKTsKICAgICAgY29uc3QgdHJlbmRUZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIHRyZW5kVGQuYXBwZW5kQ2hpbGQodHJlbmRFbCk7CiAgICAgIHRyLmFwcGVuZCgKICAgICAgICB0ZXh0RWwoJ3RkJywgbS5sYWJlbCksCiAgICAgICAgdGV4dEVsKCd0ZCcsIGZtdChtLmEpLCAnbnVtJyksCiAgICAgICAgdGV4dEVsKCd0ZCcsIGZtdChtLmIpLCAnbnVtJyksCiAgICAgICAgdGV4dEVsKCd0ZCcsIGAke20uZGlmZiA+IDAgPyAnKycgOiAnJ30ke2ZtdChtLmRpZmYpfWAsICdudW0nKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgbS5wY3REaWZmID09PSBudWxsID8gJ25ldycgOiBGb3JtYXQucGN0KG0ucGN0RGlmZiksICdudW0nKSwKICAgICAgICB0cmVuZFRkCiAgICAgICk7CiAgICAgIHRib2R5LmFwcGVuZENoaWxkKHRyKTsKICAgIH0pOwogICAgdGFibGUuYXBwZW5kKHRoZWFkLCB0Ym9keSk7CiAgICBjb25zdCB0YWJsZVNjcm9sbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGFibGVTY3JvbGwuY2xhc3NOYW1lID0gJ3RhYmxlLXNjcm9sbCc7CiAgICB0YWJsZVNjcm9sbC5hcHBlbmRDaGlsZCh0YWJsZSk7CiAgICBwYW5lbC5hcHBlbmRDaGlsZCh0YWJsZVNjcm9sbCk7CgogICAgY29uc3QgYWN0aW9ucyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYWN0aW9ucy5jbGFzc05hbWUgPSAnbW9kYWwtYWN0aW9ucyc7CiAgICBjb25zdCBjbG9zZUJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgY2xvc2VCdG4uY2xhc3NOYW1lID0gJ2J0bic7CiAgICBjbG9zZUJ0bi50eXBlID0gJ2J1dHRvbic7CiAgICBjbG9zZUJ0bi50ZXh0Q29udGVudCA9ICdDbG9zZSc7CiAgICBjbG9zZUJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIGNsb3NlQ2FyZE1vZGFsKTsKICAgIGFjdGlvbnMuYXBwZW5kQ2hpbGQoY2xvc2VCdG4pOwogICAgcGFuZWwuYXBwZW5kQ2hpbGQoYWN0aW9ucyk7CgogICAgb3ZlcmxheS5hcHBlbmRDaGlsZChwYW5lbCk7CiAgICBkb2N1bWVudC5ib2R5LmFwcGVuZENoaWxkKG92ZXJsYXkpOwogIH0KCiAgZnVuY3Rpb24gYnVpbGRQbGF0Zm9ybUNhcmQoY2FyZCwgbGFiZWxBLCBsYWJlbEIpIHsKICAgIGNvbnN0IGVsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBlbC5jbGFzc05hbWUgPSAncGxhdGZvcm0tY29tcGFyZS1jYXJkJzsKCiAgICBjb25zdCBoZWFkZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGhlYWRlci5jbGFzc05hbWUgPSAncGNjLWhlYWRlcic7CiAgICBjb25zdCBuYW1lV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgbmFtZVdyYXAuY2xhc3NOYW1lID0gJ3BjYy1oZWFkZXItbmFtZSc7CiAgICBjb25zdCBkb3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICBkb3QuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLWRvdCc7CiAgICBkb3Quc3R5bGUuYmFja2dyb3VuZCA9IGNhcmQuY29sb3I7CiAgICBuYW1lV3JhcC5hcHBlbmQoZG90LCB0ZXh0RWwoJ3NwYW4nLCBjYXJkLmxhYmVsLCAncGNjLW5hbWUnKSk7CiAgICBoZWFkZXIuYXBwZW5kQ2hpbGQobmFtZVdyYXApOwogICAgY29uc3QgZGlyID0gdHJlbmREaXJlY3Rpb24oY2FyZC5vdmVyYWxsKTsKICAgIGhlYWRlci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIGNhcmQub3ZlcmFsbCA9PT0gbnVsbCA/ICfigJQnIDogYCR7ZGlyID09PSAndXAnID8gJ+KWsicgOiBkaXIgPT09ICdkb3duJyA/ICfilrwnIDogJ+KAlCd9ICR7Rm9ybWF0LnBjdChjYXJkLm92ZXJhbGwpfWAsIGBwY2MtYmFkZ2UgJHtkaXJ9YCkpOwogICAgZWwuYXBwZW5kQ2hpbGQoaGVhZGVyKTsKCiAgICBjb25zdCBjYXB0aW9uID0gY2FyZC5vdmVyYWxsID09PSBudWxsCiAgICAgID8gJ05vdCBlbm91Z2ggZGF0YSB0byBjb21wYXJlIHlldCcKICAgICAgOiBkaXIgPT09ICd1cCcgPyAnSW1wcm92ZWQgY29tcGFyZWQgdG8gcHJldmlvdXMgcGVyaW9kJwogICAgICA6IGRpciA9PT0gJ2Rvd24nID8gJ0xvd2VyIHRoYW4gcHJldmlvdXMgcGVyaW9kJwogICAgICA6ICdBYm91dCB0aGUgc2FtZSBhcyB0aGUgcHJldmlvdXMgcGVyaW9kJzsKICAgIGVsLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgY2FwdGlvbiwgJ3BjYy1jYXB0aW9uJykpOwoKICAgIGNvbnN0IG1ldHJpY3NXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBtZXRyaWNzV3JhcC5jbGFzc05hbWUgPSAncGNjLW1ldHJpY3MnOwogICAgY2FyZC5tZXRyaWNzLmZvckVhY2goKG0pID0+IG1ldHJpY3NXcmFwLmFwcGVuZENoaWxkKGJ1aWxkTWV0cmljUm93KG0pKSk7CiAgICBlbC5hcHBlbmRDaGlsZChtZXRyaWNzV3JhcCk7CgogICAgZWwuYXBwZW5kQ2hpbGQoYnVpbGRDYXJkRm9vdGVyKGNhcmQpKTsKCiAgICBjb25zdCB2aWV3TGluayA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgdmlld0xpbmsudHlwZSA9ICdidXR0b24nOwogICAgdmlld0xpbmsuY2xhc3NOYW1lID0gJ3BjYy12aWV3LWxpbmsnOwogICAgdmlld0xpbmsudGV4dENvbnRlbnQgPSAnVmlldyBGdWxsIENvbXBhcmlzb24g4oaSJzsKICAgIHZpZXdMaW5rLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gb3BlbkNhcmRNb2RhbChjYXJkLCBsYWJlbEEsIGxhYmVsQikpOwogICAgZWwuYXBwZW5kQ2hpbGQodmlld0xpbmspOwoKICAgIHJldHVybiBlbDsKICB9CgogIGZ1bmN0aW9uIHJlbmRlclBsYXRmb3JtQ29tcGFyaXNvbkNhcmRzKHdyYXAsIHJlc3VsdCwgbGFiZWxBLCBsYWJlbEIpIHsKICAgIGNvbnN0IGFsbENhcmRzID0gYnVpbGRQbGF0Zm9ybUNhcmRzKHJlc3VsdCk7CgogICAgY29uc3Qgc2VjdGlvbiA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgc2VjdGlvbi5jbGFzc05hbWUgPSAncGNjLXNlY3Rpb24nOwogICAgc2VjdGlvbi5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsICdQbGF0Zm9ybSBQZXJmb3JtYW5jZSBDb21wYXJpc29uJywgJ3NlY3Rpb24tdGl0bGUnKSk7CgogICAgaWYgKCFhbGxDYXJkcy5sZW5ndGgpIHsKICAgICAgc2VjdGlvbi5hcHBlbmRDaGlsZChlbXB0eVN0YXRlKHsKICAgICAgICBpY29uOiAnZ2l0LWNvbXBhcmUnLAogICAgICAgIHRpdGxlOiAnTm8gZGF0YSBhdmFpbGFibGUgZm9yIHRoZSBzZWxlY3RlZCBkYXRlIHJhbmdlcy4nLAogICAgICAgIG1lc3NhZ2U6ICdUcnkgYSB3aWRlciByYW5nZSwgb3IgY2hlY2sgdGhhdCBwb3N0cyBleGlzdCBmb3IgYXQgbGVhc3Qgb25lIHBsYXRmb3JtIGluIFJhbmdlIEEgb3IgUmFuZ2UgQi4nLAogICAgICB9KSk7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQoc2VjdGlvbik7CiAgICAgIHJldHVybjsKICAgIH0KCiAgICBjb25zdCBjb250cm9scyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY29udHJvbHMuY2xhc3NOYW1lID0gJ3BjYy1jb250cm9scyc7CgogICAgY29uc3Qgc29ydFNlbGVjdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NlbGVjdCcpOwogICAgQ0FSRF9TT1JUX01PREVTLmZvckVhY2goKG0pID0+IHsKICAgICAgY29uc3Qgb3B0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7IG9wdC52YWx1ZSA9IG0ua2V5OyBvcHQudGV4dENvbnRlbnQgPSBtLmxhYmVsOwogICAgICBpZiAobS5rZXkgPT09IGNhcmRTb3J0TW9kZSkgb3B0LnNlbGVjdGVkID0gdHJ1ZTsKICAgICAgc29ydFNlbGVjdC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgICBzb3J0U2VsZWN0LmFkZEV2ZW50TGlzdGVuZXIoJ2NoYW5nZScsICgpID0+IHsgY2FyZFNvcnRNb2RlID0gc29ydFNlbGVjdC52YWx1ZTsgcmVuZGVyQ2FyZEdyaWQoKTsgfSk7CiAgICBjb250cm9scy5hcHBlbmRDaGlsZChsYWJlbGVkKCdTb3J0IEJ5Jywgc29ydFNlbGVjdCkpOwoKICAgIGNvbnN0IGZpbHRlclBpbGxzID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBmaWx0ZXJQaWxscy5jbGFzc05hbWUgPSAncGxhdGZvcm0tZmlsdGVyLXBpbGxzJzsKICAgIGNvbnN0IGFsbEJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgYWxsQnRuLnR5cGUgPSAnYnV0dG9uJzsKICAgIGFsbEJ0bi5kYXRhc2V0LmZpbHRlciA9ICdhbGwnOwogICAgYWxsQnRuLnRleHRDb250ZW50ID0gJ0FsbCBQbGF0Zm9ybXMnOwogICAgYWxsQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4geyBjYXJkUGxhdGZvcm1GaWx0ZXIgPSAnYWxsJzsgcmVuZGVyQ2FyZEdyaWQoKTsgfSk7CiAgICBmaWx0ZXJQaWxscy5hcHBlbmRDaGlsZChhbGxCdG4pOwogICAgYWxsQ2FyZHMuZm9yRWFjaCgoY2FyZCkgPT4gewogICAgICBjb25zdCBidG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgICAgYnRuLnR5cGUgPSAnYnV0dG9uJzsKICAgICAgYnRuLmRhdGFzZXQuZmlsdGVyID0gY2FyZC5wbGF0Zm9ybTsKICAgICAgY29uc3QgZG90ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICBkb3QuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLWRvdCc7CiAgICAgIGRvdC5zdHlsZS5iYWNrZ3JvdW5kID0gY2FyZC5jb2xvcjsKICAgICAgYnRuLmFwcGVuZChkb3QsIGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKGNhcmQubGFiZWwpKTsKICAgICAgYnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4geyBjYXJkUGxhdGZvcm1GaWx0ZXIgPSBjYXJkLnBsYXRmb3JtOyByZW5kZXJDYXJkR3JpZCgpOyB9KTsKICAgICAgZmlsdGVyUGlsbHMuYXBwZW5kQ2hpbGQoYnRuKTsKICAgIH0pOwogICAgY29udHJvbHMuYXBwZW5kQ2hpbGQoZmlsdGVyUGlsbHMpOwogICAgc2VjdGlvbi5hcHBlbmRDaGlsZChjb250cm9scyk7CgogICAgY29uc3QgZ3JpZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZ3JpZC5jbGFzc05hbWUgPSAncGxhdGZvcm0tY29tcGFyZS1ncmlkJzsKICAgIGdyaWQuaWQgPSAncGxhdGZvcm1Db21wYXJlR3JpZCc7CiAgICBzZWN0aW9uLmFwcGVuZENoaWxkKGdyaWQpOwogICAgd3JhcC5hcHBlbmRDaGlsZChzZWN0aW9uKTsKCiAgICBmdW5jdGlvbiByZW5kZXJDYXJkR3JpZCgpIHsKICAgICAgY29uc3QgZ3JpZEVsID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3BsYXRmb3JtQ29tcGFyZUdyaWQnKTsKICAgICAgaWYgKCFncmlkRWwpIHJldHVybjsKICAgICAgZ3JpZEVsLmlubmVySFRNTCA9ICcnOwogICAgICBjb25zdCB2aXNpYmxlID0gY2FyZFBsYXRmb3JtRmlsdGVyID09PSAnYWxsJyA/IGFsbENhcmRzIDogYWxsQ2FyZHMuZmlsdGVyKChjKSA9PiBjLnBsYXRmb3JtID09PSBjYXJkUGxhdGZvcm1GaWx0ZXIpOwogICAgICBjb25zdCBzb3J0ZWQgPSBzb3J0Q2FyZHModmlzaWJsZSwgY2FyZFNvcnRNb2RlKTsKICAgICAgaWYgKCFzb3J0ZWQubGVuZ3RoKSB7CiAgICAgICAgZ3JpZEVsLmFwcGVuZENoaWxkKGVtcHR5U3RhdGUoeyBpY29uOiAnZ2l0LWNvbXBhcmUnLCBtZXNzYWdlOiAnTm8gZGF0YSBmb3IgdGhpcyBwbGF0Zm9ybSBpbiB0aGUgc2VsZWN0ZWQgZGF0ZSByYW5nZXMuJyB9KSk7CiAgICAgIH0gZWxzZSB7CiAgICAgICAgc29ydGVkLmZvckVhY2goKGNhcmQpID0+IGdyaWRFbC5hcHBlbmRDaGlsZChidWlsZFBsYXRmb3JtQ2FyZChjYXJkLCBsYWJlbEEsIGxhYmVsQikpKTsKICAgICAgfQogICAgICBmaWx0ZXJQaWxscy5xdWVyeVNlbGVjdG9yQWxsKCdidXR0b24nKS5mb3JFYWNoKChidG4pID0+IHsKICAgICAgICBidG4uY2xhc3NMaXN0LnRvZ2dsZSgnaXMtYWN0aXZlJywgYnRuLmRhdGFzZXQuZmlsdGVyID09PSBjYXJkUGxhdGZvcm1GaWx0ZXIpOwogICAgICB9KTsKICAgIH0KICAgIHJlbmRlckNhcmRHcmlkKCk7CiAgfQoKICBmdW5jdGlvbiByZW5kZXIoKSB7CiAgICByb290ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3ZpZXctY29tcGFyaXNvbicpOwogICAgc2hlbGwoKTsKICB9CgogIHJldHVybiB7IHJlbmRlciB9Owp9KSgpOwoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIFVwbG9hZCB0YWI6IGRyYWctZHJvcCwgdmFsaWRhdGlvbiBwcmV2aWV3LCBwZXItd2VlayBjb25mbGljdAogICByZXNvbHV0aW9uLCBjb21taXQg4oCUIHBsdXMgdGhlIFVwbG9hZCBIaXN0b3J5IHRhYi4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmNvbnN0IFVwbG9hZCA9ICgoKSA9PiB7CiAgbGV0IHJvb3Q7CiAgbGV0IGN1cnJlbnRQcmV2aWV3ID0gbnVsbDsgLy8geyBmaWxlUGF0aCwgb3JpZ2luYWxOYW1lLCBkdXBsaWNhdGVzLCBpc3N1ZXMsIHNhbXBsZSwgLi4uIH0KICBjb25zdCBkdXBsaWNhdGVBY3Rpb25PdmVycmlkZXMgPSB7fTsKCiAgZnVuY3Rpb24gc2hlbGwoKSB7CiAgICByb290LmlubmVySFRNTCA9ICcnOwoKICAgIGNvbnN0IGludHJvID0gdGV4dEVsKCdkaXYnLCAnVXBsb2FkIGEgd2Vla2x5IGV4cG9ydCcsICdzZWN0aW9uLXRpdGxlJyk7CiAgICByb290LmFwcGVuZENoaWxkKGludHJvKTsKCiAgICBjb25zdCBkcm9wem9uZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZHJvcHpvbmUuY2xhc3NOYW1lID0gJ2Ryb3B6b25lJzsKICAgIGRyb3B6b25lLmlkID0gJ2Ryb3B6b25lJzsKICAgIGRyb3B6b25lLmlubmVySFRNTCA9IGAKICAgICAgPGRpdiBjbGFzcz0iZW1wdHktaWNvbiIgc3R5bGU9Im1hcmdpbjogMCBhdXRvIDE0cHg7Ij48aSBkYXRhLWx1Y2lkZT0idXBsb2FkLWNsb3VkIiBzdHlsZT0id2lkdGg6MjJweDtoZWlnaHQ6MjJweDsiPjwvaT48L2Rpdj4KICAgICAgPGgzPkRyYWcgJmFtcDsgZHJvcCB5b3VyIC5jc3Ygb3IgLnhsc3ggZmlsZSBoZXJlPC9oMz4KICAgICAgPHA+b3IgY2xpY2sgdG8gYnJvd3NlIOKAlCBmaWxlcyBhcmUgdmFsaWRhdGVkIGJlZm9yZSBhbnl0aGluZyBpcyBzYXZlZDwvcD4KICAgICAgPGlucHV0IHR5cGU9ImZpbGUiIGlkPSJmaWxlSW5wdXQiIGFjY2VwdD0iLmNzdiwueGxzeCwueGxzIiAvPgogICAgYDsKICAgIHJvb3QuYXBwZW5kQ2hpbGQoZHJvcHpvbmUpOwoKICAgIGNvbnN0IHByZXZpZXdBcmVhID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBwcmV2aWV3QXJlYS5pZCA9ICdwcmV2aWV3QXJlYSc7CiAgICByb290LmFwcGVuZENoaWxkKHByZXZpZXdBcmVhKTsKCiAgICB3aXJlRHJvcHpvbmUoZHJvcHpvbmUpOwogIH0KCiAgZnVuY3Rpb24gd2lyZURyb3B6b25lKGRyb3B6b25lKSB7CiAgICBjb25zdCBpbnB1dCA9IGRyb3B6b25lLnF1ZXJ5U2VsZWN0b3IoJyNmaWxlSW5wdXQnKTsKICAgIGRyb3B6b25lLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gaW5wdXQuY2xpY2soKSk7CiAgICBpbnB1dC5hZGRFdmVudExpc3RlbmVyKCdjaGFuZ2UnLCAoKSA9PiB7CiAgICAgIGlmIChpbnB1dC5maWxlc1swXSkgaGFuZGxlRmlsZShpbnB1dC5maWxlc1swXSk7CiAgICB9KTsKICAgIFsnZHJhZ2VudGVyJywgJ2RyYWdvdmVyJ10uZm9yRWFjaCgoZXZ0KSA9PgogICAgICBkcm9wem9uZS5hZGRFdmVudExpc3RlbmVyKGV2dCwgKGUpID0+IHsgZS5wcmV2ZW50RGVmYXVsdCgpOyBkcm9wem9uZS5jbGFzc0xpc3QuYWRkKCdpcy1kcmFnJyk7IH0pCiAgICApOwogICAgWydkcmFnbGVhdmUnLCAnZHJvcCddLmZvckVhY2goKGV2dCkgPT4KICAgICAgZHJvcHpvbmUuYWRkRXZlbnRMaXN0ZW5lcihldnQsIChlKSA9PiB7IGUucHJldmVudERlZmF1bHQoKTsgZHJvcHpvbmUuY2xhc3NMaXN0LnJlbW92ZSgnaXMtZHJhZycpOyB9KQogICAgKTsKICAgIGRyb3B6b25lLmFkZEV2ZW50TGlzdGVuZXIoJ2Ryb3AnLCAoZSkgPT4gewogICAgICBjb25zdCBmaWxlID0gZS5kYXRhVHJhbnNmZXIuZmlsZXNbMF07CiAgICAgIGlmIChmaWxlKSBoYW5kbGVGaWxlKGZpbGUpOwogICAgfSk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiBoYW5kbGVGaWxlKGZpbGUpIHsKICAgIGNvbnN0IGFyZWEgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncHJldmlld0FyZWEnKTsKICAgIGFyZWEuaW5uZXJIVE1MID0gJyc7CiAgICBhcmVhLmFwcGVuZENoaWxkKHJvd1dpdGhTcGlubmVyKCdWYWxpZGF0aW5nIGZpbGXigKYnKSk7CiAgICBPYmplY3Qua2V5cyhkdXBsaWNhdGVBY3Rpb25PdmVycmlkZXMpLmZvckVhY2goKGspID0+IGRlbGV0ZSBkdXBsaWNhdGVBY3Rpb25PdmVycmlkZXNba10pOwogICAgdHJ5IHsKICAgICAgY3VycmVudFByZXZpZXcgPSBhd2FpdCBBcGkucHJldmlld1VwbG9hZChmaWxlKTsKICAgICAgcmVuZGVyUHJldmlldygpOwogICAgfSBjYXRjaCAoZXJyKSB7CiAgICAgIGFyZWEuaW5uZXJIVE1MID0gJyc7CiAgICAgIGFyZWEuYXBwZW5kQ2hpbGQoZXJyb3JCYW5uZXIoZXJyLm1lc3NhZ2UpKTsKICAgIH0KICB9CgogIGZ1bmN0aW9uIHJvd1dpdGhTcGlubmVyKHRleHQpIHsKICAgIGNvbnN0IGVsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBlbC5jbGFzc05hbWUgPSAnbG9hZGluZy1yb3cnOwogICAgY29uc3Qgc3Bpbm5lciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgIHNwaW5uZXIuY2xhc3NOYW1lID0gJ3NwaW5uZXInOwogICAgZWwuYXBwZW5kKHNwaW5uZXIsIGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKGAgJHt0ZXh0fWApKTsKICAgIHJldHVybiBlbDsKICB9CiAgZnVuY3Rpb24gZXJyb3JCYW5uZXIobWVzc2FnZSkgewogICAgY29uc3QgZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGVsLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGVsLnN0eWxlLmJvcmRlckxlZnQgPSAnM3B4IHNvbGlkIHZhcigtLXN0YXR1cy1jcml0aWNhbCknOwogICAgZWwuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBgQ291bGQgbm90IHJlYWQgdGhpcyBmaWxlOiAke21lc3NhZ2V9YCwgJ211dGVkJykpOwogICAgcmV0dXJuIGVsOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyUHJldmlldygpIHsKICAgIGNvbnN0IGFyZWEgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncHJldmlld0FyZWEnKTsKICAgIGFyZWEuaW5uZXJIVE1MID0gJyc7CiAgICBjb25zdCBwID0gY3VycmVudFByZXZpZXc7CgogICAgY29uc3Qgc3VtbWFyeVRpdGxlID0gdGV4dEVsKCdkaXYnLCAnVmFsaWRhdGlvbiBzdW1tYXJ5JywgJ3NlY3Rpb24tdGl0bGUnKTsKICAgIGNvbnN0IHN1bW1hcnlHcmlkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBzdW1tYXJ5R3JpZC5jbGFzc05hbWUgPSAnc3RhdC1ncmlkJzsKICAgIHN1bW1hcnlHcmlkLmFwcGVuZCgKICAgICAgc3RhdFRpbGUoJ0ZpbGUnLCBwLm9yaWdpbmFsTmFtZSksCiAgICAgIHN0YXRUaWxlKCdTaGVldHMgZm91bmQnLCBwLnNoZWV0cy5sZW5ndGgpLAogICAgICBzdGF0VGlsZSgnVG90YWwgcm93cyAoYWxsIHNoZWV0cyknLCBwLnRvdGFsRGF0YVJvd3MpLAogICAgICBzdGF0VGlsZSgnTmV3IHJlY29yZHMnLCBwLm5ld1JlY29yZHNDb3VudCksCiAgICAgIHN0YXRUaWxlKCdFeGFjdCBkdXBsaWNhdGVzIGZvdW5kJywgcC5kdXBsaWNhdGVzLmxlbmd0aCksCiAgICAgIHN0YXRUaWxlKCdEdXBsaWNhdGUgcm93cyBpbiBmaWxlJywgcC5kdXBsaWNhdGVSb3dzSW5GaWxlKSwKICAgICAgc3RhdFRpbGUoJ1Jvd3Mgd2l0aCBlcnJvcnMnLCBwLmVycm9yUm93cykKICAgICk7CiAgICBhcmVhLmFwcGVuZChzdW1tYXJ5VGl0bGUsIHN1bW1hcnlHcmlkKTsKCiAgICBpZiAocC5zaGVldHMubGVuZ3RoKSB7CiAgICAgIGNvbnN0IHNoZWV0c1RpdGxlID0gdGV4dEVsKCdkaXYnLCAnU2hlZXQgYnJlYWtkb3duJywgJ3NlY3Rpb24tdGl0bGUnKTsKICAgICAgY29uc3Qgc2hlZXRzVGFibGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0YWJsZScpOwogICAgICBzaGVldHNUYWJsZS5jbGFzc05hbWUgPSAnZGF0YS10YWJsZSc7CiAgICAgIHNoZWV0c1RhYmxlLmlubmVySFRNTCA9ICc8dGhlYWQ+PHRyPjx0aD5TaGVldDwvdGg+PHRoPkxheW91dCBkZXRlY3RlZDwvdGg+PHRoIGNsYXNzPSJudW0iPlJvd3M8L3RoPjx0aCBjbGFzcz0ibnVtIj5WYWxpZDwvdGg+PHRoIGNsYXNzPSJudW0iPkVycm9yczwvdGg+PC90cj48L3RoZWFkPic7CiAgICAgIGNvbnN0IHNoZWV0c0JvZHkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0Ym9keScpOwogICAgICBwLnNoZWV0cy5mb3JFYWNoKChzKSA9PiB7CiAgICAgICAgY29uc3QgdHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgICAgIGNvbnN0IGxheW91dExhYmVsID0gcy5mb3JtYXQgPT09ICdhZ2VuZGEnID8gJ0xSUyBhZ2VuZGEgdHJhY2tlcicgOiBzLmZvcm1hdCA9PT0gJ3NpbXBsZScgPyAnU2ltcGxlIHBsYXRmb3JtIHRhYmxlJyA6ICdOb3QgcmVjb2duaXplZCDigJQgc2F2ZWQgYXMgcmF3IGRhdGEgb25seSc7CiAgICAgICAgdHIuYXBwZW5kKAogICAgICAgICAgdGV4dEVsKCd0ZCcsIHMubmFtZSksCiAgICAgICAgICB0ZXh0RWwoJ3RkJywgbGF5b3V0TGFiZWwpLAogICAgICAgICAgdGV4dEVsKCd0ZCcsIFN0cmluZyhzLnRvdGFsUm93cyksICdudW0nKSwKICAgICAgICAgIHRleHRFbCgndGQnLCBTdHJpbmcocy52YWxpZFJvd3MpLCAnbnVtJyksCiAgICAgICAgICB0ZXh0RWwoJ3RkJywgU3RyaW5nKHMuZXJyb3JSb3dzKSwgJ251bScpCiAgICAgICAgKTsKICAgICAgICBzaGVldHNCb2R5LmFwcGVuZENoaWxkKHRyKTsKICAgICAgfSk7CiAgICAgIHNoZWV0c1RhYmxlLmFwcGVuZENoaWxkKHNoZWV0c0JvZHkpOwogICAgICBjb25zdCBzaGVldHNXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIHNoZWV0c1dyYXAuY2xhc3NOYW1lID0gJ3RhYmxlLXNjcm9sbCc7CiAgICAgIHNoZWV0c1dyYXAuYXBwZW5kQ2hpbGQoc2hlZXRzVGFibGUpOwogICAgICBhcmVhLmFwcGVuZChzaGVldHNUaXRsZSwgc2hlZXRzV3JhcCk7CiAgICB9CgogICAgaWYgKHAuZHVwbGljYXRlcy5sZW5ndGgpIHsKICAgICAgY29uc3QgZHVwVGl0bGUgPSB0ZXh0RWwoJ2RpdicsIGBFeGFjdCBkdXBsaWNhdGVzIGZvdW5kICgke3AuZHVwbGljYXRlcy5sZW5ndGh9KWAsICdzZWN0aW9uLXRpdGxlJyk7CiAgICAgIGFyZWEuYXBwZW5kQ2hpbGQoZHVwVGl0bGUpOwogICAgICBhcmVhLmFwcGVuZENoaWxkKHRleHRFbCgKICAgICAgICAnZGl2JywKICAgICAgICAnRWFjaCBvZiB0aGVzZSByb3dzIGlzIGJ5dGUtZm9yLWJ5dGUgaWRlbnRpY2FsIHRvIGFuIGFscmVhZHktc2F2ZWQgcmVjb3JkIOKAlCBldmVyeSBmaWVsZCBtYXRjaGVzLCBpbmNsdWRpbmcgZXZlcnkgbWV0cmljLCBub3QganVzdCB0aGUgZGF0ZS9jYXB0aW9uL3BsYXRmb3JtLiBDaG9vc2Ugd2hhdCB0byBkbyB3aXRoIGVhY2gg4oCUIG9yIHNldCBhIGRlZmF1bHQgZm9yIGFsbCBvZiB0aGVtLiAoQSByb3cgdGhhdCBzaGFyZXMgdGhlIHNhbWUgZGF0ZS9jYXB0aW9uL3BsYXRmb3JtIGJ1dCBoYXMgZGlmZmVyZW50IG51bWJlcnMgaXMgbm90IHNob3duIGhlcmUg4oCUIGl04oCZcyBpbXBvcnRlZCBhdXRvbWF0aWNhbGx5IGFzIGl0cyBvd24gbmV3IHJlY29yZCwgc2luY2UgaXRzIGFuYWx5dGljcyBjaGFuZ2VkLiknLAogICAgICAgICdtdXRlZCcKICAgICAgKSk7CiAgICAgIGNvbnN0IGRlZmF1bHRSb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgZGVmYXVsdFJvdy5jbGFzc05hbWUgPSAnZmllbGQtaW5saW5lJzsKICAgICAgZGVmYXVsdFJvdy5zdHlsZS5tYXJnaW4gPSAnMTBweCAwJzsKICAgICAgY29uc3QgZGVmYXVsdFNlbGVjdCA9IGFjdGlvblNlbGVjdCgnc2tpcCcpOwogICAgICBkZWZhdWx0U2VsZWN0LmlkID0gJ2RlZmF1bHREdXBsaWNhdGVBY3Rpb25TZWxlY3QnOwogICAgICBkZWZhdWx0U2VsZWN0LmFkZEV2ZW50TGlzdGVuZXIoJ2NoYW5nZScsICgpID0+IHsKICAgICAgICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcuY29uZmxpY3Qtcm93IHNlbGVjdFtkYXRhLWhhc2hdJykuZm9yRWFjaCgoc2VsKSA9PiB7CiAgICAgICAgICBpZiAoIWR1cGxpY2F0ZUFjdGlvbk92ZXJyaWRlc1tzZWwuZGF0YXNldC5oYXNoXSkgc2VsLnZhbHVlID0gZGVmYXVsdFNlbGVjdC52YWx1ZTsKICAgICAgICB9KTsKICAgICAgfSk7CiAgICAgIGRlZmF1bHRSb3cuYXBwZW5kKHRleHRFbCgnbGFiZWwnLCAnRGVmYXVsdCBhY3Rpb24gZm9yIGFsbCBtYXRjaGVzJyksIGRlZmF1bHRTZWxlY3QpOwogICAgICBhcmVhLmFwcGVuZENoaWxkKGRlZmF1bHRSb3cpOwoKICAgICAgY29uc3QgbGlzdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBsaXN0LmNsYXNzTmFtZSA9ICdjb25mbGljdC1saXN0JzsKICAgICAgcC5kdXBsaWNhdGVzLmZvckVhY2goKGQpID0+IHsKICAgICAgICBjb25zdCByb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgICByb3cuY2xhc3NOYW1lID0gJ2NvbmZsaWN0LXJvdyc7CiAgICAgICAgY29uc3QgbGVmdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICAgIGxlZnQuYXBwZW5kKAogICAgICAgICAgdGV4dEVsKCdkaXYnLCBgJHtGb3JtYXQuZGF0ZShkLnB1Ymxpc2hEYXRlKX0g4oCUICR7KGQuY2FwdGlvbiB8fCAnKG5vIGNhcHRpb24pJykuc2xpY2UoMCwgNzApfWAsICd3ZWVrLWxhYmVsJyksCiAgICAgICAgICB0ZXh0RWwoJ2RpdicsIGBFeGFjdCBtYXRjaCBvZiBleGlzdGluZyByZWNvcmQgIyR7ZC5leGlzdGluZy5wb3N0SWR9IChsYXN0IHVwZGF0ZWQgJHtkLmV4aXN0aW5nLnVwZGF0ZWRBdH0pYCwgJ3dlZWstbWV0YScpCiAgICAgICAgKTsKICAgICAgICByb3cuYXBwZW5kQ2hpbGQobGVmdCk7CiAgICAgICAgY29uc3Qgc2VsID0gYWN0aW9uU2VsZWN0KCdza2lwJyk7CiAgICAgICAgc2VsLmRhdGFzZXQuaGFzaCA9IGQuaGFzaDsKICAgICAgICBzZWwuYWRkRXZlbnRMaXN0ZW5lcignY2hhbmdlJywgKCkgPT4geyBkdXBsaWNhdGVBY3Rpb25PdmVycmlkZXNbZC5oYXNoXSA9IHNlbC52YWx1ZTsgfSk7CiAgICAgICAgcm93LmFwcGVuZENoaWxkKHNlbCk7CiAgICAgICAgbGlzdC5hcHBlbmRDaGlsZChyb3cpOwogICAgICB9KTsKICAgICAgYXJlYS5hcHBlbmRDaGlsZChsaXN0KTsKICAgIH0KCiAgICBjb25zdCBub3Rlc0ZpZWxkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBub3Rlc0ZpZWxkLmNsYXNzTmFtZSA9ICdmb3JtLWZpZWxkJzsKICAgIG5vdGVzRmllbGQuc3R5bGUubWFyZ2luID0gJzEycHggMCc7CiAgICBub3Rlc0ZpZWxkLmFwcGVuZENoaWxkKHRleHRFbCgnbGFiZWwnLCAnVXBsb2FkIG5vdGVzIChvcHRpb25hbCknKSk7CiAgICBjb25zdCBub3Rlc0lucHV0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsKICAgIG5vdGVzSW5wdXQudHlwZSA9ICd0ZXh0JzsKICAgIG5vdGVzSW5wdXQuaWQgPSAndXBsb2FkTm90ZXNJbnB1dCc7CiAgICBub3Rlc0lucHV0LnBsYWNlaG9sZGVyID0gJ2UuZy4gIldlZWsgMyBleHBvcnQsIGluY2x1ZGVzIGNvcnJlY3RlZCBUaWtUb2sgbnVtYmVycyInOwogICAgbm90ZXNGaWVsZC5hcHBlbmRDaGlsZChub3Rlc0lucHV0KTsKICAgIGFyZWEuYXBwZW5kQ2hpbGQobm90ZXNGaWVsZCk7CgogICAgaWYgKHAuaXNzdWVzLmxlbmd0aCkgewogICAgICBjb25zdCBpc3N1ZXNUaXRsZSA9IHRleHRFbCgnZGl2JywgYFJvd3Mgc2tpcHBlZCBvciBmbGFnZ2VkICgke3AuaXNzdWVzLmxlbmd0aH0pYCwgJ3NlY3Rpb24tdGl0bGUnKTsKICAgICAgY29uc3QgaXNzdWVzQ2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBpc3N1ZXNDYXJkLmNsYXNzTmFtZSA9ICdpc3N1ZXMtbGlzdCc7CiAgICAgIHAuaXNzdWVzLmZvckVhY2goKGlzc3VlKSA9PiB7CiAgICAgICAgY29uc3Qgcm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgICAgcm93LmNsYXNzTmFtZSA9ICdpc3N1ZS1yb3cnOwogICAgICAgIGlmIChpc3N1ZS5yb3dOdW1iZXIpIHJvdy5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3NwYW4nLCBgUm93ICR7aXNzdWUucm93TnVtYmVyfWAsICdyb3ctbm8nKSk7CiAgICAgICAgcm93LmFwcGVuZENoaWxkKGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKGlzc3VlLm1lc3NhZ2UpKTsKICAgICAgICBpc3N1ZXNDYXJkLmFwcGVuZENoaWxkKHJvdyk7CiAgICAgIH0pOwogICAgICBhcmVhLmFwcGVuZChpc3N1ZXNUaXRsZSwgaXNzdWVzQ2FyZCk7CiAgICB9CgogICAgaWYgKHAubmV3UmVjb3Jkcy5sZW5ndGgpIHsKICAgICAgY29uc3QgbmV3VGl0bGUgPSB0ZXh0RWwoJ2RpdicsIGBOZXcgcmVjb3JkcyB0byBpbXBvcnQgKCR7cC5uZXdSZWNvcmRzLmxlbmd0aH0pYCwgJ3NlY3Rpb24tdGl0bGUnKTsKICAgICAgYXJlYS5hcHBlbmRDaGlsZChuZXdUaXRsZSk7CiAgICAgIGFyZWEuYXBwZW5kQ2hpbGQodGV4dEVsKAogICAgICAgICdkaXYnLAogICAgICAgICdUaGVzZSByb3dzIGRvbuKAmXQgbWF0Y2ggYW55dGhpbmcgYWxyZWFkeSBzYXZlZCwgc28gdGhleeKAmWxsIGJlIGltcG9ydGVkIGF1dG9tYXRpY2FsbHkg4oCUIG5vIGRlY2lzaW9uIG5lZWRlZCwgdW5saWtlIHRoZSBleGFjdC1kdXBsaWNhdGUgbWF0Y2hlcyBhYm92ZS4nLAogICAgICAgICdtdXRlZCcKICAgICAgKSk7CiAgICAgIGNvbnN0IHRhYmxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGFibGUnKTsKICAgICAgdGFibGUuY2xhc3NOYW1lID0gJ2RhdGEtdGFibGUnOwogICAgICB0YWJsZS5pbm5lckhUTUwgPSAnPHRoZWFkPjx0cj48dGg+RGF0ZTwvdGg+PHRoPkNhcHRpb248L3RoPjx0aD5UeXBlPC90aD48dGg+Q2FtcGFpZ248L3RoPjx0aD5QbGF0Zm9ybXM8L3RoPjwvdHI+PC90aGVhZD4nOwogICAgICBjb25zdCB0Ym9keSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3Rib2R5Jyk7CiAgICAgIHAubmV3UmVjb3Jkcy5mb3JFYWNoKChzKSA9PiB7CiAgICAgICAgY29uc3QgdHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgICAgIHRyLmFwcGVuZCgKICAgICAgICAgIHRleHRFbCgndGQnLCBGb3JtYXQuZGF0ZShzLnB1Ymxpc2hEYXRlKSksCiAgICAgICAgICB0ZXh0RWwoJ3RkJywgcy5jYXB0aW9uIHx8ICfigJQnKSwKICAgICAgICAgIHRleHRFbCgndGQnLCBzLmNvbnRlbnRUeXBlIHx8ICfigJQnKSwKICAgICAgICAgIHRleHRFbCgndGQnLCBzLmNhbXBhaWduVHlwZSB8fCAnVW5zcGVjaWZpZWQnKSwKICAgICAgICAgIHRleHRFbCgndGQnLCBzLnBsYXRmb3Jtcy5qb2luKCcsICcpKQogICAgICAgICk7CiAgICAgICAgdGJvZHkuYXBwZW5kQ2hpbGQodHIpOwogICAgICB9KTsKICAgICAgdGFibGUuYXBwZW5kQ2hpbGQodGJvZHkpOwogICAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIHdyYXAuY2xhc3NOYW1lID0gJ3RhYmxlLXNjcm9sbCc7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQodGFibGUpOwogICAgICBhcmVhLmFwcGVuZCh3cmFwKTsKICAgIH0KCiAgICBjb25zdCBhY3Rpb25zID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBhY3Rpb25zLmNsYXNzTmFtZSA9ICdidG4tcm93JzsKICAgIGFjdGlvbnMuc3R5bGUubWFyZ2luVG9wID0gJzE2cHgnOwogICAgY29uc3QgY29tbWl0QnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBjb21taXRCdG4uY2xhc3NOYW1lID0gJ2J0biBwcmltYXJ5JzsKICAgIGNvbW1pdEJ0bi50ZXh0Q29udGVudCA9IHAudmFsaWRSb3dzID4gMCA/IGBJbXBvcnQgJHtwLnZhbGlkUm93c30gcm93KHMpYCA6ICdOb3RoaW5nIHRvIGltcG9ydCc7CiAgICBjb21taXRCdG4uZGlzYWJsZWQgPSBwLnZhbGlkUm93cyA9PT0gMDsKICAgIGNvbW1pdEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IGNvbW1pdChjb21taXRCdG4pKTsKICAgIGNvbnN0IGNhbmNlbEJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgY2FuY2VsQnRuLmNsYXNzTmFtZSA9ICdidG4nOwogICAgY2FuY2VsQnRuLnRleHRDb250ZW50ID0gJ0NhbmNlbCc7CiAgICBjYW5jZWxCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7IGN1cnJlbnRQcmV2aWV3ID0gbnVsbDsgc2hlbGwoKTsgfSk7CiAgICBhY3Rpb25zLmFwcGVuZChjb21taXRCdG4sIGNhbmNlbEJ0bik7CiAgICBhcmVhLmFwcGVuZENoaWxkKGFjdGlvbnMpOwogIH0KCiAgZnVuY3Rpb24gc3RhdFRpbGUobGFiZWwsIHZhbHVlKSB7CiAgICBjb25zdCB0aWxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0aWxlLmNsYXNzTmFtZSA9ICdzdGF0LXRpbGUnOwogICAgdGlsZS5hcHBlbmQodGV4dEVsKCdkaXYnLCBsYWJlbCwgJ3N0YXQtbGFiZWwnKSwgdGV4dEVsKCdkaXYnLCBTdHJpbmcodmFsdWUpLCAnc3RhdC12YWx1ZScpKTsKICAgIHJldHVybiB0aWxlOwogIH0KICBmdW5jdGlvbiBhY3Rpb25TZWxlY3QoZGVmYXVsdFZhbCkgewogICAgY29uc3Qgc2VsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc2VsZWN0Jyk7CiAgICBbWydza2lwJywgJ1NraXAgKGtlZXAgZXhpc3RpbmcgcmVjb3JkIHVuY2hhbmdlZCknXSwgWyd1cGRhdGUnLCAnVXBkYXRlIGV4aXN0aW5nIHJlY29yZCddLCBbJ2NyZWF0ZScsICdDcmVhdGUgYXMgYSBuZXcsIHNlcGFyYXRlIHJlY29yZCddXS5mb3JFYWNoKChbdiwgbF0pID0+IHsKICAgICAgY29uc3Qgb3B0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7IG9wdC52YWx1ZSA9IHY7IG9wdC50ZXh0Q29udGVudCA9IGw7CiAgICAgIGlmICh2ID09PSBkZWZhdWx0VmFsKSBvcHQuc2VsZWN0ZWQgPSB0cnVlOwogICAgICBzZWwuYXBwZW5kQ2hpbGQob3B0KTsKICAgIH0pOwogICAgcmV0dXJuIHNlbDsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIGNvbW1pdChidG4pIHsKICAgIGJ0bi5kaXNhYmxlZCA9IHRydWU7CiAgICBidG4udGV4dENvbnRlbnQgPSAnSW1wb3J0aW5n4oCmJzsKICAgIGNvbnN0IGRlZmF1bHREdXBsaWNhdGVBY3Rpb24gPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZGVmYXVsdER1cGxpY2F0ZUFjdGlvblNlbGVjdCcpPy52YWx1ZSB8fCAnc2tpcCc7CiAgICBjb25zdCBub3RlcyA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd1cGxvYWROb3Rlc0lucHV0Jyk/LnZhbHVlIHx8IG51bGw7CiAgICB0cnkgewogICAgICBjb25zdCByZXN1bHQgPSBhd2FpdCBBcGkuY29tbWl0VXBsb2FkKHsKICAgICAgICBmaWxlUGF0aDogY3VycmVudFByZXZpZXcuZmlsZVBhdGgsCiAgICAgICAgb3JpZ2luYWxOYW1lOiBjdXJyZW50UHJldmlldy5vcmlnaW5hbE5hbWUsCiAgICAgICAgZGVmYXVsdER1cGxpY2F0ZUFjdGlvbiwKICAgICAgICBkdXBsaWNhdGVBY3Rpb25zOiBkdXBsaWNhdGVBY3Rpb25PdmVycmlkZXMsCiAgICAgICAgbm90ZXMsCiAgICAgIH0pOwogICAgICBUb2FzdC5zaG93KAogICAgICAgIGBJbXBvcnRlZDogJHtyZXN1bHQuaW1wb3J0ZWRSb3dzfSBuZXcsICR7cmVzdWx0LnVwZGF0ZWRSb3dzfSB1cGRhdGVkLCAke3Jlc3VsdC5za2lwcGVkUm93c30gc2tpcHBlZC5gLAogICAgICAgIHJlc3VsdC5lcnJvckNvdW50ID4gMCA/ICdlcnJvcicgOiAnc3VjY2VzcycKICAgICAgKTsKICAgICAgY3VycmVudFByZXZpZXcgPSBudWxsOwogICAgICBzaGVsbCgpOwogICAgICB3aW5kb3cuZGlzcGF0Y2hFdmVudChuZXcgQ3VzdG9tRXZlbnQoJ2xyczpkYXRhLXVwZGF0ZWQnKSk7CiAgICB9IGNhdGNoIChlcnIpIHsKICAgICAgVG9hc3Quc2hvdyhlcnIubWVzc2FnZSwgJ2Vycm9yJyk7CiAgICAgIGJ0bi5kaXNhYmxlZCA9IGZhbHNlOwogICAgICBidG4udGV4dENvbnRlbnQgPSAnUmV0cnkgaW1wb3J0JzsKICAgIH0KICB9CgogIGZ1bmN0aW9uIHJlbmRlcigpIHsKICAgIHJvb3QgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndmlldy11cGxvYWQnKTsKICAgIHNoZWxsKCk7CiAgfQoKICByZXR1cm4geyByZW5kZXIgfTsKfSkoKTsKCmNvbnN0IEhpc3RvcnkgPSAoKCkgPT4gewogIGxldCByb290OwogIGxldCBjdXJyZW50VXBsb2FkcyA9IFtdOwogIGxldCBzZWFyY2hWYWx1ZSA9ICcnOwogIGxldCBwYWdlID0gMTsKICBjb25zdCBwYWdlU2l6ZSA9IDE1OwogIGxldCBzb3J0U3RhdGUgPSB7IGtleTogJ3VwbG9hZGVkX2F0JywgZGlyOiAnZGVzYycsIHR5cGU6ICdzdHJpbmcnIH07CiAgY29uc3QgRVhQT1JUX0NPTFVNTlMgPSBbCiAgICB7IGtleTogJ2ZpbGVuYW1lJywgbGFiZWw6ICdGaWxlJyB9LAogICAgeyBrZXk6ICd1cGxvYWRlZF9hdCcsIGxhYmVsOiAnVXBsb2FkZWQnIH0sCiAgICB7IGtleTogJ3N0YXR1cycsIGxhYmVsOiAnU3RhdHVzJyB9LAogICAgeyBrZXk6ICdpbXBvcnRlZF9yb3dzJywgbGFiZWw6ICdJbXBvcnRlZCcgfSwKICAgIHsga2V5OiAndXBkYXRlZF9yb3dzJywgbGFiZWw6ICdVcGRhdGVkJyB9LAogICAgeyBrZXk6ICdza2lwcGVkX3Jvd3MnLCBsYWJlbDogJ1NraXBwZWQnIH0sCiAgICB7IGtleTogJ2Vycm9yX2NvdW50JywgbGFiZWw6ICdFcnJvcnMnIH0sCiAgICB7IGtleTogJ3dlZWtzJywgbGFiZWw6ICdXZWVrcycgfSwKICAgIHsga2V5OiAnbm90ZXMnLCBsYWJlbDogJ05vdGVzJyB9LAogIF07CgogIGZ1bmN0aW9uIGJhZGdlQ2xhc3Moc3RhdHVzKSB7CiAgICBpZiAoc3RhdHVzID09PSAnc3VjY2VzcycpIHJldHVybiAnc3VjY2Vzcyc7CiAgICBpZiAoc3RhdHVzID09PSAncGFydGlhbCcpIHJldHVybiAncGFydGlhbCc7CiAgICByZXR1cm4gJ2ZhaWxlZCc7CiAgfQoKICBmdW5jdGlvbiBzb3J0YWJsZUhlYWRlcihsYWJlbCwga2V5LCB0eXBlKSB7CiAgICBjb25zdCB0aCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RoJyk7CiAgICBpZiAodHlwZSA9PT0gJ251bWJlcicpIHRoLmNsYXNzTmFtZSA9ICdudW0nOwogICAgdGguY2xhc3NMaXN0LmFkZCgnc29ydGFibGUtdGgnKTsKICAgIGNvbnN0IGlzQWN0aXZlID0gc29ydFN0YXRlLmtleSA9PT0ga2V5OwogICAgdGguYXBwZW5kQ2hpbGQoZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUobGFiZWwpKTsKICAgIHRoLmFwcGVuZENoaWxkKHRleHRFbCgnc3BhbicsIGlzQWN0aXZlID8gKHNvcnRTdGF0ZS5kaXIgPT09ICdhc2MnID8gJyDihpEnIDogJyDihpMnKSA6ICcg4oaVJywgJ3NvcnQtYXJyb3cnKSk7CiAgICB0aC5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsKICAgICAgc29ydFN0YXRlID0geyBrZXksIGRpcjogc29ydFN0YXRlLmtleSA9PT0ga2V5ICYmIHNvcnRTdGF0ZS5kaXIgPT09ICdhc2MnID8gJ2Rlc2MnIDogJ2FzYycsIHR5cGUgfTsKICAgICAgcmVuZGVyTGlzdCgpOwogICAgfSk7CiAgICByZXR1cm4gdGg7CiAgfQoKICBmdW5jdGlvbiBmaWx0ZXJlZFVwbG9hZHMoKSB7CiAgICBjb25zdCBxID0gc2VhcmNoVmFsdWUudHJpbSgpLnRvTG93ZXJDYXNlKCk7CiAgICBpZiAoIXEpIHJldHVybiBjdXJyZW50VXBsb2FkczsKICAgIHJldHVybiBjdXJyZW50VXBsb2Fkcy5maWx0ZXIoKHUpID0+ICgKICAgICAgdS5maWxlbmFtZS50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpCiAgICAgIHx8ICh1Lm5vdGVzIHx8ICcnKS50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpCiAgICAgIHx8IHUuc3RhdHVzLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkKICAgICkpOwogIH0KCiAgZnVuY3Rpb24gc29ydGVkVXBsb2FkcygpIHsKICAgIGNvbnN0IHsga2V5LCBkaXIsIHR5cGUgfSA9IHNvcnRTdGF0ZTsKICAgIGNvbnN0IGZhY3RvciA9IGRpciA9PT0gJ2FzYycgPyAxIDogLTE7CiAgICByZXR1cm4gWy4uLmZpbHRlcmVkVXBsb2FkcygpXS5zb3J0KChhLCBiKSA9PiB7CiAgICAgIGNvbnN0IGF2ID0gYVtrZXldOwogICAgICBjb25zdCBidiA9IGJba2V5XTsKICAgICAgaWYgKGF2ID09PSBudWxsIHx8IGF2ID09PSB1bmRlZmluZWQpIHJldHVybiAxOwogICAgICBpZiAoYnYgPT09IG51bGwgfHwgYnYgPT09IHVuZGVmaW5lZCkgcmV0dXJuIC0xOwogICAgICBpZiAodHlwZSA9PT0gJ251bWJlcicpIHJldHVybiAoYXYgLSBidikgKiBmYWN0b3I7CiAgICAgIHJldHVybiBTdHJpbmcoYXYpLmxvY2FsZUNvbXBhcmUoU3RyaW5nKGJ2KSkgKiBmYWN0b3I7CiAgICB9KTsKICB9CgogIGZ1bmN0aW9uIGV4cG9ydFJvd3MoKSB7CiAgICByZXR1cm4gc29ydGVkVXBsb2FkcygpLm1hcCgodSkgPT4gKHsKICAgICAgZmlsZW5hbWU6IHUuZmlsZW5hbWUsCiAgICAgIHVwbG9hZGVkX2F0OiB1LnVwbG9hZGVkX2F0LAogICAgICBzdGF0dXM6IHUuc3RhdHVzLAogICAgICBpbXBvcnRlZF9yb3dzOiB1LmltcG9ydGVkX3Jvd3MsCiAgICAgIHVwZGF0ZWRfcm93czogdS51cGRhdGVkX3Jvd3MsCiAgICAgIHNraXBwZWRfcm93czogdS5za2lwcGVkX3Jvd3MsCiAgICAgIGVycm9yX2NvdW50OiB1LmVycm9yX2NvdW50LAogICAgICB3ZWVrczogdS53ZWVrc19hZmZlY3RlZC5tYXAoKHcpID0+IEZvcm1hdC5kYXRlKHcpKS5qb2luKCcsICcpLAogICAgICBub3RlczogdS5ub3RlcyB8fCAnJywKICAgIH0pKTsKICB9CgogIGZ1bmN0aW9uIGJ1aWxkQmFja3VwQ2FyZCgpIHsKICAgIGNvbnN0IGNhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgY2FyZC5zdHlsZS5tYXJnaW5Cb3R0b20gPSAnMjBweCc7CiAgICBjb25zdCBoZWFkZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGhlYWRlci5jbGFzc05hbWUgPSAnY2FyZC1oZWFkZXInOwogICAgaGVhZGVyLmFwcGVuZENoaWxkKHRleHRFbCgnaDMnLCAnQmFja3VwICYgUmVzdG9yZScpKTsKICAgIGNhcmQuYXBwZW5kQ2hpbGQoaGVhZGVyKTsKICAgIGNhcmQuYXBwZW5kQ2hpbGQodGV4dEVsKAogICAgICAnZGl2JywKICAgICAgJ0Rvd25sb2FkIGEgZnVsbCBzbmFwc2hvdCBvZiB0aGUgZGF0YWJhc2UgYW55IHRpbWUuIFJlc3RvcmluZyByZXBsYWNlcyBBTEwgY3VycmVudCBkYXRhIHdpdGggdGhlIHVwbG9hZGVkIGJhY2t1cCBhbmQgcmVzdGFydHMgdGhlIHNlcnZlci4gIkRlbGV0ZSBhbGwgdXBsb2FkZWQgcmVjb3JkcyIgY2xlYXJzIGV2ZXJ5IHJlY29yZCBhbmQgdGhlIHVwbG9hZCBoaXN0b3J5IChidXQgbm90IEZvbGxvd2VycyBEYXRhIFJlY29yZCkgc28geW91IGNhbiByZWJ1aWxkIGZyb20gYSBmcmVzaCBpbXBvcnQuIEJvdGggYWN0aW9ucyBjYW5ub3QgYmUgdW5kb25lIOKAlCBkb3dubG9hZCBhIGJhY2t1cCBmaXJzdC4nLAogICAgICAnbXV0ZWQnCiAgICApKTsKCiAgICBjb25zdCBhY3Rpb25zID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBhY3Rpb25zLmNsYXNzTmFtZSA9ICdidG4tcm93JzsKICAgIGFjdGlvbnMuc3R5bGUubWFyZ2luVG9wID0gJzE0cHgnOwoKICAgIGNvbnN0IGRvd25sb2FkQnRuID0gaWNvbkJ0bignYnRuIHByaW1hcnknLCAnZG93bmxvYWQnLCAnRG93bmxvYWQgQmFja3VwJyk7CiAgICBkb3dubG9hZEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsgd2luZG93LmxvY2F0aW9uLmhyZWYgPSAnL2FwaS9iYWNrdXAvZXhwb3J0JzsgfSk7CgogICAgY29uc3QgcmVzdG9yZUlucHV0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsKICAgIHJlc3RvcmVJbnB1dC50eXBlID0gJ2ZpbGUnOwogICAgcmVzdG9yZUlucHV0LmFjY2VwdCA9ICcuZGInOwogICAgcmVzdG9yZUlucHV0LnN0eWxlLmRpc3BsYXkgPSAnbm9uZSc7CgogICAgY29uc3QgcmVzdG9yZUJ0biA9IGljb25CdG4oJ2J0biBkYW5nZXInLCAndXBsb2FkJywgJ1Jlc3RvcmUgZnJvbSBCYWNrdXAnKTsKICAgIHJlc3RvcmVCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiByZXN0b3JlSW5wdXQuY2xpY2soKSk7CgogICAgcmVzdG9yZUlucHV0LmFkZEV2ZW50TGlzdGVuZXIoJ2NoYW5nZScsIGFzeW5jICgpID0+IHsKICAgICAgY29uc3QgZmlsZSA9IHJlc3RvcmVJbnB1dC5maWxlc1swXTsKICAgICAgaWYgKCFmaWxlKSByZXR1cm47CiAgICAgIGNvbnN0IHN1cmUgPSB3aW5kb3cuY29uZmlybSgKICAgICAgICAnUmVzdG9yaW5nIHdpbGwgUkVQTEFDRSBhbGwgY3VycmVudCBkYXRhIHdpdGggdGhpcyBiYWNrdXAgZmlsZSBhbmQgcmVzdGFydCB0aGUgc2VydmVyLiBUaGlzIGNhbm5vdCBiZSB1bmRvbmUuIENvbnRpbnVlPycKICAgICAgKTsKICAgICAgaWYgKCFzdXJlKSB7CiAgICAgICAgcmVzdG9yZUlucHV0LnZhbHVlID0gJyc7CiAgICAgICAgcmV0dXJuOwogICAgICB9CiAgICAgIHJlc3RvcmVCdG4uZGlzYWJsZWQgPSB0cnVlOwogICAgICB0cnkgewogICAgICAgIGNvbnN0IGZvcm0gPSBuZXcgRm9ybURhdGEoKTsKICAgICAgICBmb3JtLmFwcGVuZCgnZmlsZScsIGZpbGUpOwogICAgICAgIGNvbnN0IHJlc3VsdCA9IGF3YWl0IEFwaS5yZXN0b3JlQmFja3VwKGZvcm0pOwogICAgICAgIFRvYXN0LnNob3cocmVzdWx0Lm1lc3NhZ2UgfHwgJ0JhY2t1cCByZXN0b3JlZC4gVGhlIHNlcnZlciBpcyByZXN0YXJ0aW5nLicsICdzdWNjZXNzJyk7CiAgICAgIH0gY2F0Y2ggKGVycikgewogICAgICAgIFRvYXN0LnNob3coZXJyLm1lc3NhZ2UsICdlcnJvcicpOwogICAgICAgIHJlc3RvcmVCdG4uZGlzYWJsZWQgPSBmYWxzZTsKICAgICAgfSBmaW5hbGx5IHsKICAgICAgICByZXN0b3JlSW5wdXQudmFsdWUgPSAnJzsKICAgICAgfQogICAgfSk7CgogICAgY29uc3Qgd2lwZUJ0biA9IGljb25CdG4oJ2J0biBkYW5nZXInLCAndHJhc2gtMicsICdEZWxldGUgYWxsIHVwbG9hZGVkIHJlY29yZHMnKTsKICAgIHdpcGVCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBhc3luYyAoKSA9PiB7CiAgICAgIGNvbnN0IHR5cGVkID0gd2luZG93LnByb21wdCgKICAgICAgICAnVGhpcyBwZXJtYW5lbnRseSBkZWxldGVzIEVWRVJZIHVwbG9hZGVkIHJlY29yZCBhbmQgYWxsIHVwbG9hZCBoaXN0b3J5LCBzbyB5b3UgY2FuIHJlYnVpbGQgdGhlIGRhdGEgZnJvbSBhIGZyZXNoIGltcG9ydC4gRm9sbG93ZXJzIERhdGEgUmVjb3JkIGlzIG5vdCBhZmZlY3RlZC4gVGhpcyBjYW5ub3QgYmUgdW5kb25lLlxuXG5UeXBlIERFTEVURSB0byBjb25maXJtOicKICAgICAgKTsKICAgICAgaWYgKHR5cGVkICE9PSAnREVMRVRFJykgewogICAgICAgIGlmICh0eXBlZCAhPT0gbnVsbCkgVG9hc3Quc2hvdygnQ29uZmlybWF0aW9uIHBocmFzZSBkaWQgbm90IG1hdGNoIOKAlCBub3RoaW5nIHdhcyBkZWxldGVkLicsICdlcnJvcicpOwogICAgICAgIHJldHVybjsKICAgICAgfQogICAgICB3aXBlQnRuLmRpc2FibGVkID0gdHJ1ZTsKICAgICAgdHJ5IHsKICAgICAgICBjb25zdCByZXN1bHQgPSBhd2FpdCBBcGkud2lwZVVwbG9hZGVkUmVjb3JkcygpOwogICAgICAgIGNvbnN0IG4gPSByZXN1bHQuZGVsZXRlZCA/IHJlc3VsdC5kZWxldGVkLnBvc3RzIDogMDsKICAgICAgICBUb2FzdC5zaG93KGBEZWxldGVkICR7bn0gcmVjb3JkKHMpLiBVcGxvYWQgYSBmaWxlIHRvIGxvYWQgZnJlc2ggZGF0YS5gLCAnc3VjY2VzcycpOwogICAgICAgIHdpbmRvdy5kaXNwYXRjaEV2ZW50KG5ldyBDdXN0b21FdmVudCgnbHJzOmRhdGEtdXBkYXRlZCcpKTsKICAgICAgICByZW5kZXIoKTsKICAgICAgfSBjYXRjaCAoZXJyKSB7CiAgICAgICAgVG9hc3Quc2hvdyhlcnIubWVzc2FnZSB8fCAnRGVsZXRlIGZhaWxlZC4nLCAnZXJyb3InKTsKICAgICAgfSBmaW5hbGx5IHsKICAgICAgICB3aXBlQnRuLmRpc2FibGVkID0gZmFsc2U7CiAgICAgIH0KICAgIH0pOwoKICAgIGFjdGlvbnMuYXBwZW5kKGRvd25sb2FkQnRuLCByZXN0b3JlQnRuLCByZXN0b3JlSW5wdXQsIHdpcGVCdG4pOwogICAgY2FyZC5hcHBlbmRDaGlsZChhY3Rpb25zKTsKICAgIHJldHVybiBjYXJkOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gcmVuZGVyKCkgewogICAgcm9vdCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd2aWV3LWhpc3RvcnknKTsKICAgIHJvb3QuaW5uZXJIVE1MID0gJyc7CiAgICByb290LmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgJ1VwbG9hZCBoaXN0b3J5JywgJ3NlY3Rpb24tdGl0bGUnKSk7CiAgICByb290LmFwcGVuZENoaWxkKGJ1aWxkQmFja3VwQ2FyZCgpKTsKCiAgICBjdXJyZW50VXBsb2FkcyA9IGF3YWl0IEFwaS51cGxvYWRIaXN0b3J5KCk7CiAgICBpZiAoIWN1cnJlbnRVcGxvYWRzLmxlbmd0aCkgewogICAgICByb290LmFwcGVuZENoaWxkKGVtcHR5U3RhdGUoewogICAgICAgIGljb246ICd1cGxvYWQtY2xvdWQnLAogICAgICAgIHRpdGxlOiAnTm8gdXBsb2FkcyB5ZXQnLAogICAgICAgIG1lc3NhZ2U6ICdJbXBvcnQgeW91ciBmaXJzdCB3ZWVrbHkgZXhwb3J0IHRvIHN0YXJ0IHNlZWluZyBkYXRhIGFjcm9zcyB0aGUgYXBwLicsCiAgICAgICAgYWN0aW9uTGFiZWw6ICdVcGxvYWQgZGF0YScsCiAgICAgICAgb25BY3Rpb246ICgpID0+IGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3IoJy50YWItYnRuW2RhdGEtdGFiPSJ1cGxvYWQiXScpPy5jbGljaygpLAogICAgICB9KSk7CiAgICAgIHJldHVybjsKICAgIH0KCiAgICBjb25zdCB0b29sYmFyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0b29sYmFyLmNsYXNzTmFtZSA9ICdyZWNvcmRzLXRvb2xiYXInOwogICAgdG9vbGJhci5hcHBlbmRDaGlsZChidWlsZFNlYXJjaEJveCh7CiAgICAgIHBsYWNlaG9sZGVyOiAnU2VhcmNoIGZpbGVuYW1lLCBzdGF0dXMsIG9yIG5vdGVz4oCmJywKICAgICAgdmFsdWU6IHNlYXJjaFZhbHVlLAogICAgICBvbkNoYW5nZTogKHYpID0+IHsgc2VhcmNoVmFsdWUgPSB2OyBwYWdlID0gMTsgcmVuZGVyTGlzdCgpOyB9LAogICAgfSkpOwogICAgcm9vdC5hcHBlbmRDaGlsZCh0b29sYmFyKTsKCiAgICByb290LmFwcGVuZENoaWxkKGJ1aWxkRXhwb3J0QnV0dG9ucyh7CiAgICAgIGdldFJvd3NBbmRDb2x1bW5zOiAoKSA9PiAoeyByb3dzOiBleHBvcnRSb3dzKCksIGNvbHVtbnM6IEVYUE9SVF9DT0xVTU5TIH0pLAogICAgICBmaWxlbmFtZUJhc2U6ICd1cGxvYWQtaGlzdG9yeScsCiAgICAgIHNoZWV0TmFtZTogJ1VwbG9hZCBIaXN0b3J5JywKICAgIH0pKTsKCiAgICBjb25zdCBjYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjYXJkLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHdyYXAuY2xhc3NOYW1lID0gJ3RhYmxlLXNjcm9sbCc7CiAgICB3cmFwLmlkID0gJ2hpc3RvcnlUYWJsZVdyYXAnOwogICAgY2FyZC5hcHBlbmRDaGlsZCh3cmFwKTsKICAgIGNvbnN0IHBhZ2VyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBwYWdlci5jbGFzc05hbWUgPSAncGFnaW5hdGlvbi1yb3cnOwogICAgcGFnZXIuaWQgPSAnaGlzdG9yeVBhZ2VyJzsKICAgIGNhcmQuYXBwZW5kQ2hpbGQocGFnZXIpOwogICAgcm9vdC5hcHBlbmRDaGlsZChjYXJkKTsKCiAgICByZW5kZXJMaXN0KCk7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJMaXN0KCkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdoaXN0b3J5VGFibGVXcmFwJyk7CiAgICBjb25zdCBwYWdlckVsID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2hpc3RvcnlQYWdlcicpOwogICAgaWYgKCF3cmFwKSByZXR1cm47CiAgICBjb25zdCBhbGxTb3J0ZWQgPSBzb3J0ZWRVcGxvYWRzKCk7CiAgICBpZiAoIWFsbFNvcnRlZC5sZW5ndGgpIHsKICAgICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgICAgd3JhcC5hcHBlbmRDaGlsZChlbXB0eVN0YXRlKHsgaWNvbjogJ3VwbG9hZC1jbG91ZCcsIG1lc3NhZ2U6ICdObyB1cGxvYWRzIG1hdGNoIHlvdXIgc2VhcmNoLicgfSkpOwogICAgICBpZiAocGFnZXJFbCkgcGFnZXJFbC5pbm5lckhUTUwgPSAnJzsKICAgICAgcmV0dXJuOwogICAgfQogICAgY29uc3QgeyBwYWdlUm93cywgdG90YWxQYWdlcywgc2FmZVBhZ2UsIHRvdGFsIH0gPSBwYWdpbmF0ZUNsaWVudFNpZGUoYWxsU29ydGVkLCBwYWdlLCBwYWdlU2l6ZSk7CiAgICBwYWdlID0gc2FmZVBhZ2U7CgogICAgY29uc3QgdGFibGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0YWJsZScpOwogICAgdGFibGUuY2xhc3NOYW1lID0gJ2RhdGEtdGFibGUnOwogICAgY29uc3QgdGhlYWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0aGVhZCcpOwogICAgY29uc3QgaGVhZFRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgIGhlYWRUci5hcHBlbmQoCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdGaWxlJywgJ2ZpbGVuYW1lJywgJ3N0cmluZycpLAogICAgICBzb3J0YWJsZUhlYWRlcignVXBsb2FkZWQnLCAndXBsb2FkZWRfYXQnLCAnc3RyaW5nJyksCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdTdGF0dXMnLCAnc3RhdHVzJywgJ3N0cmluZycpLAogICAgICBzb3J0YWJsZUhlYWRlcignSW1wb3J0ZWQnLCAnaW1wb3J0ZWRfcm93cycsICdudW1iZXInKSwKICAgICAgc29ydGFibGVIZWFkZXIoJ1VwZGF0ZWQnLCAndXBkYXRlZF9yb3dzJywgJ251bWJlcicpLAogICAgICBzb3J0YWJsZUhlYWRlcignU2tpcHBlZCcsICdza2lwcGVkX3Jvd3MnLCAnbnVtYmVyJyksCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdFcnJvcnMnLCAnZXJyb3JfY291bnQnLCAnbnVtYmVyJyksCiAgICAgIHRleHRFbCgndGgnLCAnV2Vla3MnKSwKICAgICAgdGV4dEVsKCd0aCcsICdOb3RlcycpCiAgICApOwogICAgdGhlYWQuYXBwZW5kQ2hpbGQoaGVhZFRyKTsKICAgIGNvbnN0IHRib2R5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGJvZHknKTsKICAgIHBhZ2VSb3dzLmZvckVhY2goKHUpID0+IHsKICAgICAgY29uc3QgdHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgICB0ci5zdHlsZS5jdXJzb3IgPSAncG9pbnRlcic7CiAgICAgIGNvbnN0IGJhZGdlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICBiYWRnZS5jbGFzc05hbWUgPSBgYmFkZ2UgJHtiYWRnZUNsYXNzKHUuc3RhdHVzKX1gOwogICAgICBiYWRnZS50ZXh0Q29udGVudCA9IHUuc3RhdHVzOwogICAgICBjb25zdCBzdGF0dXNUZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIHN0YXR1c1RkLmFwcGVuZENoaWxkKGJhZGdlKTsKICAgICAgdHIuYXBwZW5kKAogICAgICAgIHRleHRFbCgndGQnLCB1LmZpbGVuYW1lKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgdS51cGxvYWRlZF9hdCksCiAgICAgICAgc3RhdHVzVGQsCiAgICAgICAgdGV4dEVsKCd0ZCcsIFN0cmluZyh1LmltcG9ydGVkX3Jvd3MpLCAnbnVtJyksCiAgICAgICAgdGV4dEVsKCd0ZCcsIFN0cmluZyh1LnVwZGF0ZWRfcm93cyksICdudW0nKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgU3RyaW5nKHUuc2tpcHBlZF9yb3dzKSwgJ251bScpLAogICAgICAgIHRleHRFbCgndGQnLCBTdHJpbmcodS5lcnJvcl9jb3VudCksICdudW0nKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgdS53ZWVrc19hZmZlY3RlZC5tYXAoKHcpID0+IEZvcm1hdC5kYXRlKHcpKS5qb2luKCcsICcpIHx8ICfigJQnKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgdS5ub3RlcyB8fCAn4oCUJykKICAgICAgKTsKICAgICAgdHIuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB0b2dnbGVFcnJvcnModS5pZCwgdHIpKTsKICAgICAgdGJvZHkuYXBwZW5kQ2hpbGQodHIpOwogICAgfSk7CiAgICB0YWJsZS5hcHBlbmQodGhlYWQsIHRib2R5KTsKICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CiAgICB3cmFwLmFwcGVuZENoaWxkKHRhYmxlKTsKCiAgICBpZiAocGFnZXJFbCkgewogICAgICBwYWdlckVsLmlubmVySFRNTCA9ICcnOwogICAgICBwYWdlckVsLmFwcGVuZENoaWxkKGJ1aWxkUGFnZXIoewogICAgICAgIHBhZ2U6IHNhZmVQYWdlLAogICAgICAgIHRvdGFsUGFnZXMsCiAgICAgICAgdG90YWwsCiAgICAgICAgb25QcmV2OiAoKSA9PiB7IHBhZ2UgLT0gMTsgcmVuZGVyTGlzdCgpOyB9LAogICAgICAgIG9uTmV4dDogKCkgPT4geyBwYWdlICs9IDE7IHJlbmRlckxpc3QoKTsgfSwKICAgICAgfSkpOwogICAgfQogIH0KCiAgYXN5bmMgZnVuY3Rpb24gdG9nZ2xlRXJyb3JzKHVwbG9hZElkLCB0cikgewogICAgY29uc3QgZXhpc3RpbmcgPSB0ci5uZXh0RWxlbWVudFNpYmxpbmc7CiAgICBpZiAoZXhpc3RpbmcgJiYgZXhpc3RpbmcuY2xhc3NMaXN0LmNvbnRhaW5zKCdlcnJvci1sb2ctcm93JykpIHsKICAgICAgZXhpc3RpbmcucmVtb3ZlKCk7CiAgICAgIHJldHVybjsKICAgIH0KICAgIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJy5lcnJvci1sb2ctcm93JykuZm9yRWFjaCgoZWwpID0+IGVsLnJlbW92ZSgpKTsKICAgIGNvbnN0IGVycm9ycyA9IGF3YWl0IEFwaS51cGxvYWRFcnJvcnModXBsb2FkSWQpOwogICAgY29uc3QgbG9nUm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgIGxvZ1Jvdy5jbGFzc05hbWUgPSAnZXJyb3ItbG9nLXJvdyc7CiAgICBjb25zdCB0ZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICB0ZC5jb2xTcGFuID0gOTsKICAgIGlmICghZXJyb3JzLmxlbmd0aCkgewogICAgICB0ZC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsICdObyBpc3N1ZXMgbG9nZ2VkIGZvciB0aGlzIHVwbG9hZC4nLCAnbXV0ZWQnKSk7CiAgICB9IGVsc2UgewogICAgICBjb25zdCBsaXN0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIGxpc3QuY2xhc3NOYW1lID0gJ2lzc3Vlcy1saXN0JzsKICAgICAgZXJyb3JzLmZvckVhY2goKGUpID0+IHsKICAgICAgICBjb25zdCByb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgICByb3cuY2xhc3NOYW1lID0gJ2lzc3VlLXJvdyc7CiAgICAgICAgY29uc3QgYmFkZ2UgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICAgICAgYmFkZ2UuY2xhc3NOYW1lID0gYGJhZGdlICR7ZS5zZXZlcml0eX0tc2V2YDsKICAgICAgICBiYWRnZS50ZXh0Q29udGVudCA9IGUuc2V2ZXJpdHk7CiAgICAgICAgcm93LmFwcGVuZChiYWRnZSwgZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUoYCAke2Uucm93X251bWJlciA/IGBSb3cgJHtlLnJvd19udW1iZXJ9OiBgIDogJyd9JHtlLm1lc3NhZ2V9YCkpOwogICAgICAgIGxpc3QuYXBwZW5kQ2hpbGQocm93KTsKICAgICAgfSk7CiAgICAgIHRkLmFwcGVuZENoaWxkKGxpc3QpOwogICAgfQoKICAgIGNvbnN0IHJhd0J0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgcmF3QnRuLmNsYXNzTmFtZSA9ICdidG4nOwogICAgcmF3QnRuLnN0eWxlLm1hcmdpblRvcCA9ICcxMHB4JzsKICAgIHJhd0J0bi50ZXh0Q29udGVudCA9ICdWaWV3IGV2ZXJ5IHJhdyBzb3VyY2Ugcm93IGZyb20gdGhpcyB1cGxvYWQnOwogICAgcmF3QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gbG9hZFJhd1Jvd3ModXBsb2FkSWQsIHJhd0J0bikpOwogICAgdGQuYXBwZW5kQ2hpbGQocmF3QnRuKTsKICAgIGNvbnN0IHJhd1dyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHJhd1dyYXAuaWQgPSBgcmF3V3JhcC0ke3VwbG9hZElkfWA7CiAgICB0ZC5hcHBlbmRDaGlsZChyYXdXcmFwKTsKCiAgICBsb2dSb3cuYXBwZW5kQ2hpbGQodGQpOwogICAgdHIuYWZ0ZXIobG9nUm93KTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIGxvYWRSYXdSb3dzKHVwbG9hZElkLCBidG4pIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZChgcmF3V3JhcC0ke3VwbG9hZElkfWApOwogICAgaWYgKCF3cmFwKSByZXR1cm47CiAgICBpZiAod3JhcC5kYXRhc2V0LmxvYWRlZCkgewogICAgICB3cmFwLnN0eWxlLmRpc3BsYXkgPSB3cmFwLnN0eWxlLmRpc3BsYXkgPT09ICdub25lJyA/ICdibG9jaycgOiAnbm9uZSc7CiAgICAgIHJldHVybjsKICAgIH0KICAgIGJ0bi50ZXh0Q29udGVudCA9ICdMb2FkaW5n4oCmJzsKICAgIGNvbnN0IHsgcm93cywgdG90YWwgfSA9IGF3YWl0IEFwaS51cGxvYWRSYXdSb3dzKHVwbG9hZElkKTsKICAgIHdyYXAuZGF0YXNldC5sb2FkZWQgPSAnMSc7CiAgICBidG4udGV4dENvbnRlbnQgPSBgU2hvd2luZyAke3Jvd3MubGVuZ3RofSBvZiAke3RvdGFsfSByYXcgcm93KHMpYDsKCiAgICBjb25zdCBieVNoZWV0ID0gbmV3IE1hcCgpOwogICAgcm93cy5mb3JFYWNoKChyKSA9PiB7CiAgICAgIGlmICghYnlTaGVldC5oYXMoci5zaGVldF9uYW1lKSkgYnlTaGVldC5zZXQoci5zaGVldF9uYW1lLCBbXSk7CiAgICAgIGJ5U2hlZXQuZ2V0KHIuc2hlZXRfbmFtZSkucHVzaChyKTsKICAgIH0pOwoKICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CiAgICB3cmFwLnN0eWxlLm1hcmdpblRvcCA9ICcxMHB4JzsKICAgIGJ5U2hlZXQuZm9yRWFjaCgoc2hlZXRSb3dzLCBzaGVldE5hbWUpID0+IHsKICAgICAgd3JhcC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIGBTaGVldDogJHtzaGVldE5hbWV9ICgke3NoZWV0Um93cy5sZW5ndGh9IHJvdyhzKSlgLCAnc3RhdC1sYWJlbCcpKTsKICAgICAgY29uc3QgaGVhZGVycyA9IHNoZWV0Um93c1swXS5oZWFkZXJzOwogICAgICBjb25zdCB0YWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICAgIHRhYmxlLmNsYXNzTmFtZSA9ICdkYXRhLXRhYmxlJzsKICAgICAgY29uc3QgdGhlYWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgICB0aGVhZC5hcHBlbmQodGV4dEVsKCd0aCcsICdSb3cgIycpLCB0ZXh0RWwoJ3RoJywgJ0xpbmtlZCB0byBwb3N0JykpOwogICAgICBjb25zdCBjb2xDb3VudCA9IGhlYWRlcnMgPyBoZWFkZXJzLmxlbmd0aCA6IE1hdGgubWF4KC4uLnNoZWV0Um93cy5tYXAoKHIpID0+IHIucmF3Lmxlbmd0aCkpOwogICAgICBmb3IgKGxldCBpID0gMDsgaSA8IGNvbENvdW50OyBpICs9IDEpIHRoZWFkLmFwcGVuZENoaWxkKHRleHRFbCgndGgnLCBoZWFkZXJzICYmIGhlYWRlcnNbaV0gPyBTdHJpbmcoaGVhZGVyc1tpXSkgOiBgQ29sICR7aSArIDF9YCkpOwogICAgICBjb25zdCB0aGVhZFdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0aGVhZCcpOwogICAgICB0aGVhZFdyYXAuYXBwZW5kQ2hpbGQodGhlYWQpOwogICAgICBjb25zdCB0Ym9keSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3Rib2R5Jyk7CiAgICAgIHNoZWV0Um93cy5mb3JFYWNoKChyKSA9PiB7CiAgICAgICAgY29uc3QgdHIyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgICB0cjIuYXBwZW5kKHRleHRFbCgndGQnLCBTdHJpbmcoci5yb3dfbnVtYmVyKSksIHRleHRFbCgndGQnLCByLnBvc3RfaWQgPyBgIyR7ci5wb3N0X2lkfWAgOiAn4oCUJykpOwogICAgICAgIGZvciAobGV0IGkgPSAwOyBpIDwgY29sQ291bnQ7IGkgKz0gMSkgewogICAgICAgICAgY29uc3QgdmFsID0gci5yYXdbaV07CiAgICAgICAgICB0cjIuYXBwZW5kQ2hpbGQodGV4dEVsKCd0ZCcsIHZhbCA9PT0gdW5kZWZpbmVkIHx8IHZhbCA9PT0gbnVsbCA/ICcnIDogU3RyaW5nKHZhbCkuc2xpY2UoMCwgNjApKSk7CiAgICAgICAgfQogICAgICAgIHRib2R5LmFwcGVuZENoaWxkKHRyMik7CiAgICAgIH0pOwogICAgICB0YWJsZS5hcHBlbmQodGhlYWRXcmFwLCB0Ym9keSk7CiAgICAgIGNvbnN0IHNjcm9sbFdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgc2Nyb2xsV3JhcC5jbGFzc05hbWUgPSAndGFibGUtc2Nyb2xsJzsKICAgICAgc2Nyb2xsV3JhcC5zdHlsZS5tYXJnaW5Cb3R0b20gPSAnMTZweCc7CiAgICAgIHNjcm9sbFdyYXAuYXBwZW5kQ2hpbGQodGFibGUpOwogICAgICB3cmFwLmFwcGVuZENoaWxkKHNjcm9sbFdyYXApOwogICAgfSk7CiAgfQoKICByZXR1cm4geyByZW5kZXIgfTsKfSkoKTsKCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBGb2xsb3dlcnMgRGF0YSB0YWI6IG1hbnVhbCB3ZWVrbHkgZm9sbG93ZXItY291bnQgZW50cnkgcGVyCiAgIHBsYXRmb3JtIOKAlCBlbnRpcmVseSBpbmRlcGVuZGVudCBvZiBzcHJlYWRzaGVldCB1cGxvYWRzIChpdHMgb3duCiAgIHRhYmxlLCBpdHMgb3duIEFQSSwgbmV2ZXIgdG91Y2hlZCBieSB0aGUgaW1wb3J0IHBpcGVsaW5lKS4gUG93ZXJzCiAgIEZvbGxvd2VyIEdyb3d0aCBjaGFydHMvY29tcGFyaXNvbnMgZWxzZXdoZXJlIGluIHRoZSBhcHAuCiAgID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwpjb25zdCBGb2xsb3dlcnMgPSAoKCkgPT4gewogIGxldCByb290OwogIGxldCBlZGl0aW5nSWQgPSBudWxsOyAvLyBub24tbnVsbCB3aGlsZSB0aGUgZm9ybSBpcyBlZGl0aW5nIGFuIGV4aXN0aW5nIGVudHJ5IHJhdGhlciB0aGFuIGFkZGluZyBhIG5ldyBvbmUKICBsZXQgc29ydFN0YXRlID0geyBrZXk6ICdlbnRyeV9kYXRlJywgZGlyOiAnZGVzYycsIHR5cGU6ICdzdHJpbmcnIH07CiAgbGV0IGN1cnJlbnRSb3dzID0gW107CiAgbGV0IHNlYXJjaFZhbHVlID0gJyc7CiAgbGV0IHBhZ2UgPSAxOwogIGNvbnN0IHBhZ2VTaXplID0gMTA7CiAgY29uc3QgRVhQT1JUX0NPTFVNTlMgPSBbCiAgICB7IGtleTogJ3BsYXRmb3JtX2xhYmVsJywgbGFiZWw6ICdQbGF0Zm9ybScgfSwKICAgIHsga2V5OiAnZW50cnlfZGF0ZScsIGxhYmVsOiAnV2VlayAvIERhdGUnIH0sCiAgICB7IGtleTogJ2ZvbGxvd2Vyc19jb3VudCcsIGxhYmVsOiAnRm9sbG93ZXJzIENvdW50JyB9LAogICAgeyBrZXk6ICd1cGRhdGVkX2F0JywgbGFiZWw6ICdMYXN0IFVwZGF0ZWQnIH0sCiAgXTsKCiAgZnVuY3Rpb24gYWxsUGxhdGZvcm1zKCkgewogICAgcmV0dXJuICh3aW5kb3cuX19maWx0ZXJPcHRpb25zQ2FjaGUgfHwgeyBhbGxQbGF0Zm9ybXM6IFtdIH0pLmFsbFBsYXRmb3JtcyB8fCBbXTsKICB9CgogIGZ1bmN0aW9uIHBsYXRmb3JtTWV0YUZvcihpZCkgewogICAgcmV0dXJuIGFsbFBsYXRmb3JtcygpLmZpbmQoKHApID0+IHAuaWQgPT09IGlkKSB8fCB7IGxhYmVsOiBpZCwgY29sb3I6ICcjOTk5JyB9OwogIH0KCiAgZnVuY3Rpb24gc2hlbGwoKSB7CiAgICByb290LmlubmVySFRNTCA9ICcnOwogICAgcm9vdC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsICdGb2xsb3dlcnMgRGF0YSBSZWNvcmQnLCAnc2VjdGlvbi10aXRsZScpKTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQodGV4dEVsKAogICAgICAnZGl2JywKICAgICAgJ01hbnVhbGx5IGxvZyBlYWNoIHBsYXRmb3Jt4oCZcyB0b3RhbCBmb2xsb3dlciBjb3VudCBvbmNlIGEgd2Vlay4gVGhpcyBpcyBpbmRlcGVuZGVudCBvZiBzcHJlYWRzaGVldCB1cGxvYWRzIOKAlCBpdCBwb3dlcnMgRm9sbG93ZXIgR3Jvd3RoIGNoYXJ0cyBhbmQgY29tcGFyaXNvbnMgZWxzZXdoZXJlIGluIHRoZSBhcHAuJywKICAgICAgJ211dGVkJwogICAgKSk7CgogICAgY29uc3QgZm9ybUNhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGZvcm1DYXJkLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGZvcm1DYXJkLnN0eWxlLm1hcmdpbkJvdHRvbSA9ICcyMHB4JzsKICAgIGZvcm1DYXJkLmlkID0gJ2ZvbGxvd2Vyc0Zvcm1DYXJkJzsKICAgIHJvb3QuYXBwZW5kQ2hpbGQoZm9ybUNhcmQpOwoKICAgIGNvbnN0IHRvb2xiYXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRvb2xiYXIuY2xhc3NOYW1lID0gJ3JlY29yZHMtdG9vbGJhcic7CiAgICB0b29sYmFyLmFwcGVuZENoaWxkKGJ1aWxkU2VhcmNoQm94KHsKICAgICAgcGxhY2Vob2xkZXI6ICdTZWFyY2ggcGxhdGZvcm0gb3IgZGF0ZeKApicsCiAgICAgIHZhbHVlOiBzZWFyY2hWYWx1ZSwKICAgICAgb25DaGFuZ2U6ICh2KSA9PiB7IHNlYXJjaFZhbHVlID0gdjsgcGFnZSA9IDE7IHJlbmRlclRhYmxlKCk7IH0sCiAgICB9KSk7CiAgICByb290LmFwcGVuZENoaWxkKHRvb2xiYXIpOwoKICAgIHJvb3QuYXBwZW5kQ2hpbGQoYnVpbGRFeHBvcnRCdXR0b25zKHsKICAgICAgZ2V0Um93c0FuZENvbHVtbnM6ICgpID0+ICh7IHJvd3M6IGV4cG9ydFJvd3MoKSwgY29sdW1uczogRVhQT1JUX0NPTFVNTlMgfSksCiAgICAgIGZpbGVuYW1lQmFzZTogJ2ZvbGxvd2Vycy1kYXRhJywKICAgICAgc2hlZXROYW1lOiAnRm9sbG93ZXJzIERhdGEnLAogICAgfSkpOwoKICAgIGNvbnN0IHRhYmxlQ2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGFibGVDYXJkLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGNvbnN0IHRhYmxlV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGFibGVXcmFwLmNsYXNzTmFtZSA9ICd0YWJsZS1zY3JvbGwnOwogICAgdGFibGVXcmFwLmlkID0gJ2ZvbGxvd2Vyc1RhYmxlV3JhcCc7CiAgICB0YWJsZUNhcmQuYXBwZW5kQ2hpbGQodGFibGVXcmFwKTsKICAgIGNvbnN0IHBhZ2VyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBwYWdlci5jbGFzc05hbWUgPSAncGFnaW5hdGlvbi1yb3cnOwogICAgcGFnZXIuaWQgPSAnZm9sbG93ZXJzUGFnZXInOwogICAgdGFibGVDYXJkLmFwcGVuZENoaWxkKHBhZ2VyKTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQodGFibGVDYXJkKTsKCiAgICByZW5kZXJGb3JtKCk7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJGb3JtKCkgewogICAgY29uc3QgY2FyZCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmb2xsb3dlcnNGb3JtQ2FyZCcpOwogICAgaWYgKCFjYXJkKSByZXR1cm47CiAgICBjYXJkLmlubmVySFRNTCA9ICcnOwogICAgY29uc3QgaGVhZGVyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBoZWFkZXIuY2xhc3NOYW1lID0gJ2NhcmQtaGVhZGVyJzsKICAgIGhlYWRlci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2gzJywgZWRpdGluZ0lkICE9PSBudWxsID8gJ0VkaXQgZW50cnknIDogJ0FkZCBhIHdlZWtseSBlbnRyeScpKTsKICAgIGNhcmQuYXBwZW5kQ2hpbGQoaGVhZGVyKTsKCiAgICBjb25zdCBncmlkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBncmlkLmNsYXNzTmFtZSA9ICdmb3JtLWdyaWQnOwoKICAgIGNvbnN0IHBsYXRmb3JtRmllbGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHBsYXRmb3JtRmllbGQuY2xhc3NOYW1lID0gJ2Zvcm0tZmllbGQnOwogICAgcGxhdGZvcm1GaWVsZC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2xhYmVsJywgJ1BsYXRmb3JtJykpOwogICAgY29uc3QgcGxhdGZvcm1TZWxlY3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzZWxlY3QnKTsKICAgIHBsYXRmb3JtU2VsZWN0LmlkID0gJ2ZvbGxvd2Vyc1BsYXRmb3JtSW5wdXQnOwogICAgYWxsUGxhdGZvcm1zKCkuZm9yRWFjaCgocCkgPT4gewogICAgICBjb25zdCBvcHQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsKICAgICAgb3B0LnZhbHVlID0gcC5pZDsKICAgICAgb3B0LnRleHRDb250ZW50ID0gcC5sYWJlbDsKICAgICAgcGxhdGZvcm1TZWxlY3QuYXBwZW5kQ2hpbGQob3B0KTsKICAgIH0pOwogICAgcGxhdGZvcm1GaWVsZC5hcHBlbmRDaGlsZChwbGF0Zm9ybVNlbGVjdCk7CgogICAgY29uc3QgZGF0ZUZpZWxkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBkYXRlRmllbGQuY2xhc3NOYW1lID0gJ2Zvcm0tZmllbGQnOwogICAgZGF0ZUZpZWxkLmFwcGVuZENoaWxkKHRleHRFbCgnbGFiZWwnLCAnV2VlayAvIERhdGUnKSk7CiAgICBjb25zdCBkYXRlSW5wdXQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOwogICAgZGF0ZUlucHV0LnR5cGUgPSAnZGF0ZSc7CiAgICBkYXRlSW5wdXQuaWQgPSAnZm9sbG93ZXJzRGF0ZUlucHV0JzsKICAgIGRhdGVGaWVsZC5hcHBlbmRDaGlsZChkYXRlSW5wdXQpOwoKICAgIGNvbnN0IGNvdW50RmllbGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNvdW50RmllbGQuY2xhc3NOYW1lID0gJ2Zvcm0tZmllbGQnOwogICAgY291bnRGaWVsZC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2xhYmVsJywgJ0ZvbGxvd2VycyBDb3VudCcpKTsKICAgIGNvbnN0IGNvdW50SW5wdXQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOwogICAgY291bnRJbnB1dC50eXBlID0gJ251bWJlcic7CiAgICBjb3VudElucHV0Lm1pbiA9ICcwJzsKICAgIGNvdW50SW5wdXQuc3RlcCA9ICcxJzsKICAgIGNvdW50SW5wdXQuaWQgPSAnZm9sbG93ZXJzQ291bnRJbnB1dCc7CiAgICBjb3VudEZpZWxkLmFwcGVuZENoaWxkKGNvdW50SW5wdXQpOwoKICAgIGdyaWQuYXBwZW5kKHBsYXRmb3JtRmllbGQsIGRhdGVGaWVsZCwgY291bnRGaWVsZCk7CiAgICBjYXJkLmFwcGVuZENoaWxkKGdyaWQpOwoKICAgIGNvbnN0IGVkaXRSb3cgPSBlZGl0aW5nSWQgIT09IG51bGwgPyBjdXJyZW50Um93cy5maW5kKChyKSA9PiByLmlkID09PSBlZGl0aW5nSWQpIDogbnVsbDsKICAgIGlmIChlZGl0Um93KSB7CiAgICAgIHBsYXRmb3JtU2VsZWN0LnZhbHVlID0gZWRpdFJvdy5wbGF0Zm9ybTsKICAgICAgZGF0ZUlucHV0LnZhbHVlID0gZWRpdFJvdy5lbnRyeV9kYXRlOwogICAgICBjb3VudElucHV0LnZhbHVlID0gU3RyaW5nKGVkaXRSb3cuZm9sbG93ZXJzX2NvdW50KTsKICAgIH0gZWxzZSB7CiAgICAgIGRhdGVJbnB1dC52YWx1ZSA9IG5ldyBEYXRlKCkudG9JU09TdHJpbmcoKS5zbGljZSgwLCAxMCk7CiAgICB9CgogICAgY29uc3QgYWN0aW9ucyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYWN0aW9ucy5jbGFzc05hbWUgPSAnbW9kYWwtYWN0aW9ucyc7CiAgICBjb25zdCBlcnJvckVsID0gdGV4dEVsKCdzcGFuJywgJycsICdtdXRlZCcpOwogICAgZXJyb3JFbC5pZCA9ICdmb2xsb3dlcnNGb3JtRXJyb3InOwogICAgZXJyb3JFbC5zdHlsZS5jb2xvciA9ICd2YXIoLS1zdGF0dXMtY3JpdGljYWwpJzsKCiAgICBjb25zdCBidG5Sb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGJ0blJvdy5jbGFzc05hbWUgPSAnYnRuLXJvdyc7CiAgICBjb25zdCBzYXZlQnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBzYXZlQnRuLmNsYXNzTmFtZSA9ICdidG4gcHJpbWFyeSc7CiAgICBzYXZlQnRuLnRleHRDb250ZW50ID0gZWRpdGluZ0lkICE9PSBudWxsID8gJ1NhdmUgY2hhbmdlcycgOiAnQWRkIGVudHJ5JzsKICAgIHNhdmVCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBzdWJtaXRGb3JtKHNhdmVCdG4pKTsKICAgIGJ0blJvdy5hcHBlbmRDaGlsZChzYXZlQnRuKTsKICAgIGlmIChlZGl0aW5nSWQgIT09IG51bGwpIHsKICAgICAgY29uc3QgY2FuY2VsQnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICAgIGNhbmNlbEJ0bi5jbGFzc05hbWUgPSAnYnRuJzsKICAgICAgY2FuY2VsQnRuLnRleHRDb250ZW50ID0gJ0NhbmNlbCc7CiAgICAgIGNhbmNlbEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsgZWRpdGluZ0lkID0gbnVsbDsgcmVuZGVyRm9ybSgpOyB9KTsKICAgICAgYnRuUm93LmFwcGVuZENoaWxkKGNhbmNlbEJ0bik7CiAgICB9CiAgICBhY3Rpb25zLmFwcGVuZChlcnJvckVsLCBidG5Sb3cpOwogICAgY2FyZC5hcHBlbmRDaGlsZChhY3Rpb25zKTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIHN1Ym1pdEZvcm0oYnRuKSB7CiAgICBjb25zdCBlcnJvckVsID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZvbGxvd2Vyc0Zvcm1FcnJvcicpOwogICAgZXJyb3JFbC50ZXh0Q29udGVudCA9ICcnOwogICAgY29uc3QgcGxhdGZvcm0gPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZm9sbG93ZXJzUGxhdGZvcm1JbnB1dCcpLnZhbHVlOwogICAgY29uc3QgZW50cnlEYXRlID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZvbGxvd2Vyc0RhdGVJbnB1dCcpLnZhbHVlOwogICAgY29uc3QgZm9sbG93ZXJzQ291bnQgPSBOdW1iZXIoZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZvbGxvd2Vyc0NvdW50SW5wdXQnKS52YWx1ZSk7CiAgICBidG4uZGlzYWJsZWQgPSB0cnVlOwogICAgdHJ5IHsKICAgICAgaWYgKGVkaXRpbmdJZCAhPT0gbnVsbCkgewogICAgICAgIGF3YWl0IEFwaS51cGRhdGVGb2xsb3dlcnMoZWRpdGluZ0lkLCB7IHBsYXRmb3JtLCBlbnRyeURhdGUsIGZvbGxvd2Vyc0NvdW50IH0pOwogICAgICAgIFRvYXN0LnNob3coJ0VudHJ5IHVwZGF0ZWQuJywgJ3N1Y2Nlc3MnKTsKICAgICAgfSBlbHNlIHsKICAgICAgICBhd2FpdCBBcGkuc2F2ZUZvbGxvd2Vycyh7IHBsYXRmb3JtLCBlbnRyeURhdGUsIGZvbGxvd2Vyc0NvdW50IH0pOwogICAgICAgIFRvYXN0LnNob3coJ0VudHJ5IHNhdmVkLicsICdzdWNjZXNzJyk7CiAgICAgIH0KICAgICAgZWRpdGluZ0lkID0gbnVsbDsKICAgICAgYXdhaXQgbG9hZCgpOwogICAgICB3aW5kb3cuZGlzcGF0Y2hFdmVudChuZXcgQ3VzdG9tRXZlbnQoJ2xyczpkYXRhLXVwZGF0ZWQnKSk7CiAgICB9IGNhdGNoIChlcnIpIHsKICAgICAgZXJyb3JFbC50ZXh0Q29udGVudCA9IGVyci5tZXNzYWdlOwogICAgICBidG4uZGlzYWJsZWQgPSBmYWxzZTsKICAgIH0KICB9CgogIGZ1bmN0aW9uIHN0YXJ0RWRpdChyb3cpIHsKICAgIGVkaXRpbmdJZCA9IHJvdy5pZDsKICAgIHJlbmRlckZvcm0oKTsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmb2xsb3dlcnNGb3JtQ2FyZCcpLnNjcm9sbEludG9WaWV3KHsgYmVoYXZpb3I6ICdzbW9vdGgnLCBibG9jazogJ3N0YXJ0JyB9KTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIGhhbmRsZURlbGV0ZShyb3cpIHsKICAgIGNvbnN0IHN1cmUgPSB3aW5kb3cuY29uZmlybShgRGVsZXRlIHRoZSAke3BsYXRmb3JtTWV0YUZvcihyb3cucGxhdGZvcm0pLmxhYmVsfSBlbnRyeSBmb3IgJHtGb3JtYXQuZGF0ZShyb3cuZW50cnlfZGF0ZSl9P2ApOwogICAgaWYgKCFzdXJlKSByZXR1cm47CiAgICB0cnkgewogICAgICBhd2FpdCBBcGkuZGVsZXRlRm9sbG93ZXJzKHJvdy5pZCk7CiAgICAgIFRvYXN0LnNob3coJ0VudHJ5IGRlbGV0ZWQuJywgJ3N1Y2Nlc3MnKTsKICAgICAgaWYgKGVkaXRpbmdJZCA9PT0gcm93LmlkKSBlZGl0aW5nSWQgPSBudWxsOwogICAgICBhd2FpdCBsb2FkKCk7CiAgICAgIHdpbmRvdy5kaXNwYXRjaEV2ZW50KG5ldyBDdXN0b21FdmVudCgnbHJzOmRhdGEtdXBkYXRlZCcpKTsKICAgIH0gY2F0Y2ggKGVycikgewogICAgICBUb2FzdC5zaG93KGVyci5tZXNzYWdlLCAnZXJyb3InKTsKICAgIH0KICB9CgogIGZ1bmN0aW9uIHNvcnRhYmxlSGVhZGVyKGxhYmVsLCBrZXksIHR5cGUpIHsKICAgIGNvbnN0IHRoID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGgnKTsKICAgIGlmICh0eXBlID09PSAnbnVtYmVyJykgdGguY2xhc3NOYW1lID0gJ251bSc7CiAgICB0aC5jbGFzc0xpc3QuYWRkKCdzb3J0YWJsZS10aCcpOwogICAgY29uc3QgaXNBY3RpdmUgPSBzb3J0U3RhdGUua2V5ID09PSBrZXk7CiAgICB0aC5hcHBlbmRDaGlsZChkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShsYWJlbCkpOwogICAgdGguYXBwZW5kQ2hpbGQodGV4dEVsKCdzcGFuJywgaXNBY3RpdmUgPyAoc29ydFN0YXRlLmRpciA9PT0gJ2FzYycgPyAnIOKGkScgOiAnIOKGkycpIDogJyDihpUnLCAnc29ydC1hcnJvdycpKTsKICAgIHRoLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gewogICAgICBzb3J0U3RhdGUgPSB7IGtleSwgZGlyOiBzb3J0U3RhdGUua2V5ID09PSBrZXkgJiYgc29ydFN0YXRlLmRpciA9PT0gJ2FzYycgPyAnZGVzYycgOiAnYXNjJywgdHlwZSB9OwogICAgICByZW5kZXJUYWJsZSgpOwogICAgfSk7CiAgICByZXR1cm4gdGg7CiAgfQoKICBmdW5jdGlvbiBmaWx0ZXJlZFJvd3MoKSB7CiAgICBjb25zdCBxID0gc2VhcmNoVmFsdWUudHJpbSgpLnRvTG93ZXJDYXNlKCk7CiAgICBpZiAoIXEpIHJldHVybiBjdXJyZW50Um93czsKICAgIHJldHVybiBjdXJyZW50Um93cy5maWx0ZXIoKHJvdykgPT4gewogICAgICBjb25zdCBsYWJlbCA9IHBsYXRmb3JtTWV0YUZvcihyb3cucGxhdGZvcm0pLmxhYmVsLnRvTG93ZXJDYXNlKCk7CiAgICAgIHJldHVybiBsYWJlbC5pbmNsdWRlcyhxKSB8fCByb3cuZW50cnlfZGF0ZS5pbmNsdWRlcyhxKSB8fCBTdHJpbmcocm93LmZvbGxvd2Vyc19jb3VudCkuaW5jbHVkZXMocSk7CiAgICB9KTsKICB9CgogIGZ1bmN0aW9uIHNvcnRlZFJvd3MoKSB7CiAgICBjb25zdCB7IGtleSwgZGlyLCB0eXBlIH0gPSBzb3J0U3RhdGU7CiAgICBjb25zdCBmYWN0b3IgPSBkaXIgPT09ICdhc2MnID8gMSA6IC0xOwogICAgcmV0dXJuIFsuLi5maWx0ZXJlZFJvd3MoKV0uc29ydCgoYSwgYikgPT4gewogICAgICBjb25zdCBhdiA9IGFba2V5XTsKICAgICAgY29uc3QgYnYgPSBiW2tleV07CiAgICAgIGlmIChhdiA9PT0gbnVsbCB8fCBhdiA9PT0gdW5kZWZpbmVkKSByZXR1cm4gMTsKICAgICAgaWYgKGJ2ID09PSBudWxsIHx8IGJ2ID09PSB1bmRlZmluZWQpIHJldHVybiAtMTsKICAgICAgaWYgKHR5cGUgPT09ICdudW1iZXInKSByZXR1cm4gKGF2IC0gYnYpICogZmFjdG9yOwogICAgICByZXR1cm4gU3RyaW5nKGF2KS5sb2NhbGVDb21wYXJlKFN0cmluZyhidikpICogZmFjdG9yOwogICAgfSk7CiAgfQoKICAvKiogRXZlcnkgY3VycmVudGx5LWZpbHRlcmVkL3NvcnRlZCByb3csIHNoYXBlZCBmb3IgZXhwb3J0IChub3QganVzdCB0aGUgY3VycmVudCBwYWdlKS4gKi8KICBmdW5jdGlvbiBleHBvcnRSb3dzKCkgewogICAgcmV0dXJuIHNvcnRlZFJvd3MoKS5tYXAoKHJvdykgPT4gKHsKICAgICAgcGxhdGZvcm1fbGFiZWw6IHBsYXRmb3JtTWV0YUZvcihyb3cucGxhdGZvcm0pLmxhYmVsLAogICAgICBlbnRyeV9kYXRlOiByb3cuZW50cnlfZGF0ZSwKICAgICAgZm9sbG93ZXJzX2NvdW50OiByb3cuZm9sbG93ZXJzX2NvdW50LAogICAgICB1cGRhdGVkX2F0OiByb3cudXBkYXRlZF9hdCwKICAgIH0pKTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlclRhYmxlKCkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmb2xsb3dlcnNUYWJsZVdyYXAnKTsKICAgIGNvbnN0IHBhZ2VyRWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZm9sbG93ZXJzUGFnZXInKTsKICAgIGlmICghd3JhcCkgcmV0dXJuOwogICAgY29uc3QgYWxsU29ydGVkID0gc29ydGVkUm93cygpOwogICAgaWYgKCFhbGxTb3J0ZWQubGVuZ3RoKSB7CiAgICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQoZW1wdHlTdGF0ZSh7CiAgICAgICAgaWNvbjogJ3VzZXJzJywKICAgICAgICB0aXRsZTogY3VycmVudFJvd3MubGVuZ3RoID8gJ05vIGVudHJpZXMgbWF0Y2ggeW91ciBzZWFyY2gnIDogJ05vIGZvbGxvd2VyIGVudHJpZXMgeWV0JywKICAgICAgICBtZXNzYWdlOiBjdXJyZW50Um93cy5sZW5ndGggPyAnVHJ5IGEgZGlmZmVyZW50IHBsYXRmb3JtIG5hbWUgb3IgZGF0ZS4nIDogJ0FkZCB5b3VyIGZpcnN0IHdlZWtseSBmb2xsb3dlciBjb3VudCBhYm92ZSBmb3IgYW55IHBsYXRmb3JtLicsCiAgICAgIH0pKTsKICAgICAgaWYgKHBhZ2VyRWwpIHBhZ2VyRWwuaW5uZXJIVE1MID0gJyc7CiAgICAgIHJldHVybjsKICAgIH0KICAgIGNvbnN0IHsgcGFnZVJvd3MsIHRvdGFsUGFnZXMsIHNhZmVQYWdlLCB0b3RhbCB9ID0gcGFnaW5hdGVDbGllbnRTaWRlKGFsbFNvcnRlZCwgcGFnZSwgcGFnZVNpemUpOwogICAgcGFnZSA9IHNhZmVQYWdlOwoKICAgIGNvbnN0IHRhYmxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGFibGUnKTsKICAgIHRhYmxlLmNsYXNzTmFtZSA9ICdkYXRhLXRhYmxlJzsKICAgIGNvbnN0IHRoZWFkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGhlYWQnKTsKICAgIGNvbnN0IGhlYWRUciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICBoZWFkVHIuYXBwZW5kKAogICAgICBzb3J0YWJsZUhlYWRlcignUGxhdGZvcm0nLCAncGxhdGZvcm0nLCAnc3RyaW5nJyksCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdXZWVrIC8gRGF0ZScsICdlbnRyeV9kYXRlJywgJ3N0cmluZycpLAogICAgICBzb3J0YWJsZUhlYWRlcignRm9sbG93ZXJzIENvdW50JywgJ2ZvbGxvd2Vyc19jb3VudCcsICdudW1iZXInKSwKICAgICAgdGV4dEVsKCd0aCcsICdMYXN0IFVwZGF0ZWQnKSwKICAgICAgdGV4dEVsKCd0aCcsICdBY3Rpb25zJykKICAgICk7CiAgICB0aGVhZC5hcHBlbmRDaGlsZChoZWFkVHIpOwogICAgY29uc3QgdGJvZHkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0Ym9keScpOwogICAgcGFnZVJvd3MuZm9yRWFjaCgocm93KSA9PiB7CiAgICAgIGNvbnN0IHRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgY29uc3QgcGxhdGZvcm1UZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIGNvbnN0IG1ldGEgPSBwbGF0Zm9ybU1ldGFGb3Iocm93LnBsYXRmb3JtKTsKICAgICAgY29uc3QgcGlsbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgICAgcGlsbC5jbGFzc05hbWUgPSAncGxhdGZvcm0tcGlsbCc7CiAgICAgIGNvbnN0IGRvdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgICAgZG90LmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1kb3QnOwogICAgICBkb3Quc3R5bGUuYmFja2dyb3VuZCA9IG1ldGEuY29sb3I7CiAgICAgIHBpbGwuYXBwZW5kKGRvdCwgZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUobWV0YS5sYWJlbCkpOwogICAgICBwbGF0Zm9ybVRkLmFwcGVuZENoaWxkKHBpbGwpOwoKICAgICAgY29uc3QgYWN0aW9uc1RkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgY29uc3Qgcm93QWN0aW9ucyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICByb3dBY3Rpb25zLmNsYXNzTmFtZSA9ICdyb3ctYWN0aW9ucyc7CiAgICAgIGNvbnN0IGVkaXRCdG4gPSBpY29uQnRuKCdidG4nLCAncGVuY2lsJywgJ0VkaXQnKTsKICAgICAgZWRpdEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHN0YXJ0RWRpdChyb3cpKTsKICAgICAgY29uc3QgZGVsZXRlQnRuID0gaWNvbkJ0bignYnRuIGRhbmdlcicsICd0cmFzaC0yJywgJ0RlbGV0ZScpOwogICAgICBkZWxldGVCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBoYW5kbGVEZWxldGUocm93KSk7CiAgICAgIHJvd0FjdGlvbnMuYXBwZW5kKGVkaXRCdG4sIGRlbGV0ZUJ0bik7CiAgICAgIGFjdGlvbnNUZC5hcHBlbmRDaGlsZChyb3dBY3Rpb25zKTsKCiAgICAgIHRyLmFwcGVuZCgKICAgICAgICBwbGF0Zm9ybVRkLAogICAgICAgIHRleHRFbCgndGQnLCBGb3JtYXQuZGF0ZShyb3cuZW50cnlfZGF0ZSkpLAogICAgICAgIHRleHRFbCgndGQnLCBGb3JtYXQubnVtYmVyKHJvdy5mb2xsb3dlcnNfY291bnQpLCAnbnVtJyksCiAgICAgICAgdGV4dEVsKCd0ZCcsIHJvdy51cGRhdGVkX2F0KSwKICAgICAgICBhY3Rpb25zVGQKICAgICAgKTsKICAgICAgdGJvZHkuYXBwZW5kQ2hpbGQodHIpOwogICAgfSk7CiAgICB0YWJsZS5hcHBlbmQodGhlYWQsIHRib2R5KTsKICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CiAgICB3cmFwLmFwcGVuZENoaWxkKHRhYmxlKTsKCiAgICBpZiAocGFnZXJFbCkgewogICAgICBwYWdlckVsLmlubmVySFRNTCA9ICcnOwogICAgICBwYWdlckVsLmFwcGVuZENoaWxkKGJ1aWxkUGFnZXIoewogICAgICAgIHBhZ2U6IHNhZmVQYWdlLAogICAgICAgIHRvdGFsUGFnZXMsCiAgICAgICAgdG90YWwsCiAgICAgICAgb25QcmV2OiAoKSA9PiB7IHBhZ2UgLT0gMTsgcmVuZGVyVGFibGUoKTsgfSwKICAgICAgICBvbk5leHQ6ICgpID0+IHsgcGFnZSArPSAxOyByZW5kZXJUYWJsZSgpOyB9LAogICAgICB9KSk7CiAgICB9CiAgfQoKICBhc3luYyBmdW5jdGlvbiBsb2FkKCkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmb2xsb3dlcnNUYWJsZVdyYXAnKTsKICAgIGlmICh3cmFwKSB7IHdyYXAuaW5uZXJIVE1MID0gJyc7IHdyYXAuYXBwZW5kQ2hpbGQoc2tlbGV0b25Sb3dzKDQpKTsgfQogICAgY3VycmVudFJvd3MgPSBhd2FpdCBBcGkubGlzdEZvbGxvd2Vycyh7fSk7CiAgICByZW5kZXJUYWJsZSgpOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gcmVuZGVyKCkgewogICAgcm9vdCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd2aWV3LWZvbGxvd2VycycpOwogICAgZWRpdGluZ0lkID0gbnVsbDsKICAgIHNoZWxsKCk7CiAgICBhd2FpdCBsb2FkKCk7CiAgfQoKICByZXR1cm4geyByZW5kZXIgfTsKfSkoKTsKCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBBcHAgYm9vdHN0cmFwOiB0YWIgcm91dGluZywgZmlsdGVyIGJhciB3aXJpbmcsIHRoZW1lIHRvZ2dsZS4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCigoKSA9PiB7CiAgY29uc3QgVklFV1MgPSB7CiAgICBkYXNoYm9hcmQ6IERhc2hib2FyZCwKICAgIHJlY29yZHM6IFJlY29yZHMsCiAgICBmb2xsb3dlcnM6IEZvbGxvd2VycywKICAgIGNvbXBhcmlzb246IENvbXBhcmlzb24sCiAgICB1cGxvYWQ6IFVwbG9hZCwKICAgIGhpc3Rvcnk6IEhpc3RvcnksCiAgfTsKCiAgbGV0IGFjdGl2ZVRhYiA9ICdkYXNoYm9hcmQnOwoKICBmdW5jdGlvbiBzd2l0Y2hUYWIodGFiKSB7CiAgICBhY3RpdmVUYWIgPSB0YWI7CiAgICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcudGFiLWJ0bicpLmZvckVhY2goKGJ0bikgPT4gewogICAgICBjb25zdCBpc0FjdGl2ZSA9IGJ0bi5kYXRhc2V0LnRhYiA9PT0gdGFiOwogICAgICBidG4uY2xhc3NMaXN0LnRvZ2dsZSgnaXMtYWN0aXZlJywgaXNBY3RpdmUpOwogICAgICBidG4uc2V0QXR0cmlidXRlKCdhcmlhLXNlbGVjdGVkJywgU3RyaW5nKGlzQWN0aXZlKSk7CiAgICB9KTsKICAgIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJy52aWV3JykuZm9yRWFjaCgodmlldykgPT4gewogICAgICB2aWV3LmNsYXNzTGlzdC50b2dnbGUoJ2lzLWFjdGl2ZScsIHZpZXcuaWQgPT09IGB2aWV3LSR7dGFifWApOwogICAgfSk7CiAgICAvLyBGaWx0ZXJzIGFwcGx5IHRvIERhc2hib2FyZCBhbmQgRGF0YSBSZWNvcmRzIChDb21wYXJpc29ucyBoYXMgaXRzIG93biByYW5nZSBjb250cm9scykuCiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmlsdGVyQmFyJykuc3R5bGUuZGlzcGxheSA9ICh0YWIgPT09ICdkYXNoYm9hcmQnIHx8IHRhYiA9PT0gJ3JlY29yZHMnKSA/ICdmbGV4JyA6ICdub25lJzsKICAgIHJlbmRlckFjdGl2ZVZpZXcoKTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlckFjdGl2ZVZpZXcoKSB7CiAgICBjb25zdCB2aWV3ID0gVklFV1NbYWN0aXZlVGFiXTsKICAgIGlmICh2aWV3ICYmIHZpZXcucmVuZGVyKSB2aWV3LnJlbmRlcigpOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gbG9hZEZpbHRlck9wdGlvbnMoKSB7CiAgICBjb25zdCBvcHRpb25zID0gYXdhaXQgQXBpLmZpbHRlck9wdGlvbnMoKTsKICAgIHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZSA9IG9wdGlvbnM7CgogICAgY29uc3QgcGxhdGZvcm1TZWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmlsdGVyUGxhdGZvcm0nKTsKICAgIHBsYXRmb3JtU2VsLmxlbmd0aCA9IDE7CiAgICBvcHRpb25zLnBsYXRmb3Jtcy5mb3JFYWNoKChwKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOwogICAgICBvcHQudmFsdWUgPSBwLmlkOwogICAgICBvcHQudGV4dENvbnRlbnQgPSBwLmxhYmVsOwogICAgICBwbGF0Zm9ybVNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CgogICAgY29uc3QgY2FtcGFpZ25TZWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmlsdGVyQ2FtcGFpZ24nKTsKICAgIGNhbXBhaWduU2VsLmxlbmd0aCA9IDE7CiAgICBvcHRpb25zLmNhbXBhaWduVHlwZXMuZm9yRWFjaCgoYykgPT4gewogICAgICBjb25zdCBvcHQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsKICAgICAgb3B0LnZhbHVlID0gYzsKICAgICAgb3B0LnRleHRDb250ZW50ID0gYzsKICAgICAgY2FtcGFpZ25TZWwuYXBwZW5kQ2hpbGQob3B0KTsKICAgIH0pOwoKICAgIGNvbnN0IGNvbnRlbnRTZWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmlsdGVyQ29udGVudFR5cGUnKTsKICAgIGNvbnRlbnRTZWwubGVuZ3RoID0gMTsKICAgIG9wdGlvbnMuY29udGVudFR5cGVzLmZvckVhY2goKGMpID0+IHsKICAgICAgY29uc3Qgb3B0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7CiAgICAgIG9wdC52YWx1ZSA9IGM7CiAgICAgIG9wdC50ZXh0Q29udGVudCA9IGM7CiAgICAgIGNvbnRlbnRTZWwuYXBwZW5kQ2hpbGQob3B0KTsKICAgIH0pOwogIH0KCiAgZnVuY3Rpb24gd2lyZUZpbHRlckJhcigpIHsKICAgIGNvbnN0IGRhdGVGcm9tID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZpbHRlckRhdGVGcm9tJyk7CiAgICBjb25zdCBkYXRlVG8gPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmlsdGVyRGF0ZVRvJyk7CiAgICBjb25zdCBwbGF0Zm9ybSA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJQbGF0Zm9ybScpOwogICAgY29uc3QgY2FtcGFpZ24gPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmlsdGVyQ2FtcGFpZ24nKTsKICAgIGNvbnN0IGNvbnRlbnRUeXBlID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZpbHRlckNvbnRlbnRUeXBlJyk7CiAgICBjb25zdCBmID0gU3RhdGUuZ2V0RmlsdGVycygpOwogICAgZGF0ZUZyb20udmFsdWUgPSBmLmRhdGVGcm9tOwogICAgZGF0ZVRvLnZhbHVlID0gZi5kYXRlVG87CgogICAgZnVuY3Rpb24gYXBwbHkoKSB7CiAgICAgIFN0YXRlLnNldEZpbHRlcnMoewogICAgICAgIGRhdGVGcm9tOiBkYXRlRnJvbS52YWx1ZSwKICAgICAgICBkYXRlVG86IGRhdGVUby52YWx1ZSwKICAgICAgICBwbGF0Zm9ybTogcGxhdGZvcm0udmFsdWUsCiAgICAgICAgY2FtcGFpZ25UeXBlOiBjYW1wYWlnbi52YWx1ZSwKICAgICAgICBjb250ZW50VHlwZTogY29udGVudFR5cGUudmFsdWUsCiAgICAgIH0pOwogICAgfQogICAgW2RhdGVGcm9tLCBkYXRlVG8sIHBsYXRmb3JtLCBjYW1wYWlnbiwgY29udGVudFR5cGVdLmZvckVhY2goKGVsKSA9PiBlbC5hZGRFdmVudExpc3RlbmVyKCdjaGFuZ2UnLCBhcHBseSkpOwoKICAgIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJyNmaWx0ZXJQcmVzZXRzIGJ1dHRvbicpLmZvckVhY2goKGJ0bikgPT4gewogICAgICBidG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7CiAgICAgICAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnI2ZpbHRlclByZXNldHMgYnV0dG9uJykuZm9yRWFjaCgoYikgPT4gYi5jbGFzc0xpc3QucmVtb3ZlKCdpcy1hY3RpdmUnKSk7CiAgICAgICAgYnRuLmNsYXNzTGlzdC5hZGQoJ2lzLWFjdGl2ZScpOwogICAgICAgIGNvbnN0IHByZXNldCA9IGJ0bi5kYXRhc2V0LnByZXNldDsKICAgICAgICBjb25zdCB0b2RheSA9IG5ldyBEYXRlKCk7CiAgICAgICAgY29uc3QgdG8gPSB0b2RheS50b0lTT1N0cmluZygpLnNsaWNlKDAsIDEwKTsKICAgICAgICBsZXQgZnJvbTsKICAgICAgICBpZiAocHJlc2V0ID09PSAnYWxsJykgewogICAgICAgICAgY29uc3QgbWluID0gKHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZSAmJiB3aW5kb3cuX19maWx0ZXJPcHRpb25zQ2FjaGUuZGF0ZVJhbmdlLm1pbikgfHwgdG87CiAgICAgICAgICBmcm9tID0gbWluOwogICAgICAgIH0gZWxzZSB7CiAgICAgICAgICBjb25zdCBkID0gbmV3IERhdGUodG9kYXkpOwogICAgICAgICAgZC5zZXREYXRlKGQuZ2V0RGF0ZSgpIC0gKE51bWJlcihwcmVzZXQpIC0gMSkpOwogICAgICAgICAgZnJvbSA9IGQudG9JU09TdHJpbmcoKS5zbGljZSgwLCAxMCk7CiAgICAgICAgfQogICAgICAgIGRhdGVGcm9tLnZhbHVlID0gZnJvbTsKICAgICAgICBkYXRlVG8udmFsdWUgPSB0bzsKICAgICAgICBhcHBseSgpOwogICAgICB9KTsKICAgIH0pOwoKICAgIFN0YXRlLm9uQ2hhbmdlKCgpID0+IHsKICAgICAgaWYgKGFjdGl2ZVRhYiA9PT0gJ2Rhc2hib2FyZCcpIERhc2hib2FyZC5yZW5kZXIoKTsKICAgICAgaWYgKGFjdGl2ZVRhYiA9PT0gJ3JlY29yZHMnKSBSZWNvcmRzLnJlbmRlcigpOwogICAgfSk7CiAgfQoKICBmdW5jdGlvbiB3aXJlVGFicygpIHsKICAgIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJy50YWItYnRuJykuZm9yRWFjaCgoYnRuKSA9PiB7CiAgICAgIGJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHN3aXRjaFRhYihidG4uZGF0YXNldC50YWIpKTsKICAgIH0pOwogIH0KCiAgd2luZG93LmFkZEV2ZW50TGlzdGVuZXIoJ2xyczpkYXRhLXVwZGF0ZWQnLCBhc3luYyAoKSA9PiB7CiAgICBhd2FpdCBsb2FkRmlsdGVyT3B0aW9ucygpOwogICAgcmVuZGVyQWN0aXZlVmlldygpOwogIH0pOwoKICAvLyAtLS0tLS0tLS0tIEF1dGggc2NyZWVuIC0tLS0tLS0tLS0KICBsZXQgYXBwSW5pdGlhbGl6ZWQgPSBmYWxzZTsKCiAgZnVuY3Rpb24gc2hvd0F1dGhTY3JlZW4oKSB7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXV0aFNjcmVlbicpLnN0eWxlLmRpc3BsYXkgPSAnZmxleCc7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXBwU2hlbGwnKS5zdHlsZS5kaXNwbGF5ID0gJ25vbmUnOwogICAgY29uc3QgY29kZUlucHV0ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2F1dGhDb2RlJyk7CiAgICBjb2RlSW5wdXQudmFsdWUgPSAnJzsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhdXRoRXJyb3InKS50ZXh0Q29udGVudCA9ICcnOwogICAgY29kZUlucHV0LmZvY3VzKCk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiBzaG93QXBwKCkgewogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2F1dGhTY3JlZW4nKS5zdHlsZS5kaXNwbGF5ID0gJ25vbmUnOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2FwcFNoZWxsJykuc3R5bGUuZGlzcGxheSA9ICcnOwogICAgaWYgKCFhcHBJbml0aWFsaXplZCkgewogICAgICBhcHBJbml0aWFsaXplZCA9IHRydWU7CiAgICAgIHdpcmVUYWJzKCk7CiAgICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdsb2dvdXRCdG4nKS5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIGFzeW5jICgpID0+IHsKICAgICAgICBhd2FpdCBBcGkuYXV0aExvZ291dCgpOwogICAgICAgIGFwcEluaXRpYWxpemVkID0gZmFsc2U7CiAgICAgICAgc2hvd0F1dGhTY3JlZW4oKTsKICAgICAgfSk7CiAgICAgIGF3YWl0IGxvYWRGaWx0ZXJPcHRpb25zKCk7CiAgICAgIHdpcmVGaWx0ZXJCYXIoKTsKICAgICAgc3dpdGNoVGFiKCdkYXNoYm9hcmQnKTsKICAgIH0gZWxzZSB7CiAgICAgIGF3YWl0IGxvYWRGaWx0ZXJPcHRpb25zKCk7CiAgICAgIHJlbmRlckFjdGl2ZVZpZXcoKTsKICAgIH0KICB9CgogIGFzeW5jIGZ1bmN0aW9uIHN1Ym1pdEF1dGgoKSB7CiAgICBjb25zdCBlcnJvckVsID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2F1dGhFcnJvcicpOwogICAgY29uc3QgYnRuID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2F1dGhTdWJtaXRCdG4nKTsKICAgIGNvbnN0IGNvZGVJbnB1dCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhdXRoQ29kZScpOwogICAgZXJyb3JFbC50ZXh0Q29udGVudCA9ICcnOwogICAgYnRuLmRpc2FibGVkID0gdHJ1ZTsKICAgIGJ0bi50ZXh0Q29udGVudCA9ICdDaGVja2luZ+KApic7CiAgICB0cnkgewogICAgICBhd2FpdCBBcGkuYXV0aExvZ2luKGNvZGVJbnB1dC52YWx1ZSk7CiAgICAgIGF3YWl0IHNob3dBcHAoKTsKICAgIH0gY2F0Y2ggKGVycikgewogICAgICBlcnJvckVsLnRleHRDb250ZW50ID0gZXJyLm1lc3NhZ2U7CiAgICB9IGZpbmFsbHkgewogICAgICBidG4uZGlzYWJsZWQgPSBmYWxzZTsKICAgICAgYnRuLmlubmVySFRNTCA9ICc8aSBkYXRhLWx1Y2lkZT0iYXJyb3ctcmlnaHQiIHN0eWxlPSJ3aWR0aDoxNHB4O2hlaWdodDoxNHB4OyI+PC9pPiBFbnRlcic7CiAgICB9CiAgfQoKICBmdW5jdGlvbiB3aXJlQXV0aEZvcm0oKSB7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXV0aFN1Ym1pdEJ0bicpLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgc3VibWl0QXV0aCk7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXV0aENvZGUnKS5hZGRFdmVudExpc3RlbmVyKCdrZXlkb3duJywgKGUpID0+IHsgaWYgKGUua2V5ID09PSAnRW50ZXInKSBzdWJtaXRBdXRoKCk7IH0pOwogIH0KCiAgd2luZG93LmFkZEV2ZW50TGlzdGVuZXIoJ2xyczpzaWduZWQtb3V0JywgKCkgPT4gewogICAgYXBwSW5pdGlhbGl6ZWQgPSBmYWxzZTsKICAgIHNob3dBdXRoU2NyZWVuKCk7CiAgfSk7CgogIGFzeW5jIGZ1bmN0aW9uIGluaXQoKSB7CiAgICBhcHBseUJyYW5kaW5nKCk7CiAgICB3aXJlQXV0aEZvcm0oKTsKICAgIGNvbnN0IHsgYXV0aGVudGljYXRlZCB9ID0gYXdhaXQgQXBpLmF1dGhNZSgpOwogICAgaWYgKGF1dGhlbnRpY2F0ZWQpIGF3YWl0IHNob3dBcHAoKTsKICAgIGVsc2Ugc2hvd0F1dGhTY3JlZW4oKTsKICB9CgogIC8vIEljb25zIGFyZSBwbGFjZWQgYXMgPGkgZGF0YS1sdWNpZGU9Ii4uLiI+IHBsYWNlaG9sZGVycyB0aHJvdWdob3V0IHRoZSBkeW5hbWljYWxseQogIC8vIHJlbmRlcmVkIFVJOyBMdWNpZGUgcmVwbGFjZXMgZWFjaCB3aXRoIGFuIGlubGluZSBTVkcuIFJhdGhlciB0aGFuIHJlbWVtYmVyaW5nIHRvIGNhbGwKICAvLyB0aGlzIGFmdGVyIGV2ZXJ5IHNpbmdsZSByZW5kZXIsIG9uZSBvYnNlcnZlciBjYXRjaGVzIGV2ZXJ5IERPTSBjaGFuZ2UgdGhhdCBjb3VsZCBoYXZlCiAgLy8gaW50cm9kdWNlZCBhIG5ldyBwbGFjZWhvbGRlci4KICBpZiAod2luZG93Lmx1Y2lkZSkgewogICAgd2luZG93Lmx1Y2lkZS5jcmVhdGVJY29ucygpOwogICAgLy8gY3JlYXRlSWNvbnMoKSByZXBsYWNlcyA8aSBkYXRhLWx1Y2lkZT4gcGxhY2Vob2xkZXJzIHdpdGggPHN2Zz4g4oCUIGl0c2VsZiBhIERPTQogICAgLy8gbXV0YXRpb24uIFdpdGhvdXQgZGlzY29ubmVjdGluZyBmaXJzdCwgdGhhdCB3cml0ZSByZS10cmlnZ2VycyB0aGlzIHNhbWUgb2JzZXJ2ZXIKICAgIC8vIGZvcmV2ZXIgKGFuIGluZmluaXRlIG11dGF0ZS9vYnNlcnZlIGxvb3AgdGhhdCBwZWdzIHRoZSBDUFUgYW5kIGNyYXNoZXMgdGhlIHRhYikuCiAgICAvLyBEaXNjb25uZWN0aW5nIGJlZm9yZSBlYWNoIHBhc3MgYW5kIHJlY29ubmVjdGluZyBhZnRlciwgcGx1cyBiYXRjaGluZyBidXJzdHMgb2YKICAgIC8vIG11dGF0aW9ucyBpbnRvIGEgc2luZ2xlIG1pY3JvdGFzaywgYnJlYWtzIHRoZSBjeWNsZS4KICAgIGxldCBpY29uc1NjaGVkdWxlZCA9IGZhbHNlOwogICAgY29uc3QgaWNvbk9ic2VydmVyID0gbmV3IE11dGF0aW9uT2JzZXJ2ZXIoKCkgPT4gewogICAgICBpZiAoaWNvbnNTY2hlZHVsZWQpIHJldHVybjsKICAgICAgaWNvbnNTY2hlZHVsZWQgPSB0cnVlOwogICAgICBxdWV1ZU1pY3JvdGFzaygoKSA9PiB7CiAgICAgICAgaWNvbnNTY2hlZHVsZWQgPSBmYWxzZTsKICAgICAgICBpY29uT2JzZXJ2ZXIuZGlzY29ubmVjdCgpOwogICAgICAgIHdpbmRvdy5sdWNpZGUuY3JlYXRlSWNvbnMoKTsKICAgICAgICBpY29uT2JzZXJ2ZXIub2JzZXJ2ZShkb2N1bWVudC5ib2R5LCB7IGNoaWxkTGlzdDogdHJ1ZSwgc3VidHJlZTogdHJ1ZSB9KTsKICAgICAgfSk7CiAgICB9KTsKICAgIGljb25PYnNlcnZlci5vYnNlcnZlKGRvY3VtZW50LmJvZHksIHsgY2hpbGRMaXN0OiB0cnVlLCBzdWJ0cmVlOiB0cnVlIH0pOwogIH0KCiAgaW5pdCgpOwp9KSgpOwo8L3NjcmlwdD4KPC9ib2R5Pgo8L2h0bWw+Cg==';
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
