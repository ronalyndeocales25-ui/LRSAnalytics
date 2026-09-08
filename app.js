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
const INDEX_HTML_BASE64 = 'PCFkb2N0eXBlIGh0bWw+CjxodG1sIGxhbmc9ImVuIj4KPGhlYWQ+CjxtZXRhIGNoYXJzZXQ9IlVURi04IiAvPgo8bWV0YSBuYW1lPSJ2aWV3cG9ydCIgY29udGVudD0id2lkdGg9ZGV2aWNlLXdpZHRoLCBpbml0aWFsLXNjYWxlPTEuMCIgLz4KPHRpdGxlPkxSUyBBbmFseXRpY3MgRGFzaGJvYXJkPC90aXRsZT4KPGxpbmsgcmVsPSJpY29uIiB0eXBlPSJpbWFnZS9wbmciIGlkPSJmYXZpY29uTGluayIgLz4KPGxpbmsgcmVsPSJwcmVjb25uZWN0IiBocmVmPSJodHRwczovL2ZvbnRzLmdvb2dsZWFwaXMuY29tIiAvPgo8bGluayByZWw9InByZWNvbm5lY3QiIGhyZWY9Imh0dHBzOi8vZm9udHMuZ3N0YXRpYy5jb20iIGNyb3Nzb3JpZ2luIC8+CjxsaW5rIGhyZWY9Imh0dHBzOi8vZm9udHMuZ29vZ2xlYXBpcy5jb20vY3NzMj9mYW1pbHk9SW50ZXI6d2dodEA0MDA7NTAwOzYwMDs3MDA7ODAwJmRpc3BsYXk9c3dhcCIgcmVsPSJzdHlsZXNoZWV0IiAvPgo8c2NyaXB0IHNyYz0iaHR0cHM6Ly9jZG4uanNkZWxpdnIubmV0L25wbS9jaGFydC5qc0A0LjQuNC9kaXN0L2NoYXJ0LnVtZC5taW4uanMiPjwvc2NyaXB0Pgo8c2NyaXB0IHNyYz0iaHR0cHM6Ly9jZG4uanNkZWxpdnIubmV0L25wbS9jaGFydGpzLXBsdWdpbi1kYXRhbGFiZWxzQDIvZGlzdC9jaGFydGpzLXBsdWdpbi1kYXRhbGFiZWxzLm1pbi5qcyI+PC9zY3JpcHQ+CjxzY3JpcHQgc3JjPSJodHRwczovL2Nkbi5qc2RlbGl2ci5uZXQvbnBtL2x1Y2lkZUAwLjQ2Mi4wL2Rpc3QvdW1kL2x1Y2lkZS5taW4uanMiPjwvc2NyaXB0Pgo8c3R5bGU+Ci8qIC0tLS0tLS0tLS0gRGVzaWduIHRva2VuczogZGFyayBuYXZ5ICsgZ29sZCBicmFuZGVkIHRoZW1lIChzaW5nbGUsIHBlcm1hbmVudCDigJQgbm8gbGlnaHQgdmFyaWFudCkgLS0tLS0tLS0tLSAqLwo6cm9vdCB7CiAgY29sb3Itc2NoZW1lOiBkYXJrOwogIC0tZm9udC1zYW5zOiAnSW50ZXInLCAtYXBwbGUtc3lzdGVtLCBCbGlua01hY1N5c3RlbUZvbnQsICdTRiBQcm8gRGlzcGxheScsICdTZWdvZSBVSScsIFJvYm90bywgc2Fucy1zZXJpZjsKCiAgLS1wYWdlLXBsYW5lOiBsaW5lYXItZ3JhZGllbnQoMTgwZGVnLCAjMGEwZjFjIDAlLCAjMGQxNDI0IDEwMCUpOwogIC0tcGFnZS1wbGFuZS1zb2xpZDogIzBhMGYxYzsKICAtLXNpZGViYXItYmc6ICMwYjEyMjA7CiAgLS1zdXJmYWNlLTE6IHJnYmEoMjMsIDMxLCA1MSwgMC42Mik7IC8qIGdsYXNzOiBjYXJkcywgS1BJIHRpbGVzLCBmaWx0ZXIgYmFyICovCiAgLS1zdXJmYWNlLTI6IHJnYmEoMjU1LCAyNTUsIDI1NSwgMC4wNik7IC8qIGdsYXNzOiBpbnB1dHMsIG5lc3RlZCByb3dzLCBwaWxscyAqLwogIC0tc3VyZmFjZS1zb2xpZDogIzEzMWIyZTsKICAtLWdsYXNzLWJsdXI6IGJsdXIoMjBweCk7CiAgLS1ib3JkZXI6IHJnYmEoMjU1LCAyNTUsIDI1NSwgMC4wOSk7CiAgLS10ZXh0LXByaW1hcnk6ICNmNGY1Zjc7CiAgLS10ZXh0LXNlY29uZGFyeTogI2I4YmJjNDsKICAtLXRleHQtbXV0ZWQ6ICM4Mjg2OGY7CiAgLS1ncmlkbGluZTogcmdiYSgyNTUsIDI1NSwgMjU1LCAwLjA4KTsKICAtLWJhc2VsaW5lOiByZ2JhKDI1NSwgMjU1LCAyNTUsIDAuMik7CiAgLS1zdWNjZXNzLXRleHQ6ICMzNGM3NmY7CgogIC0tc3RhdHVzLWdvb2Q6ICMyZmI4NjI7CiAgLS1zdGF0dXMtd2FybmluZzogI2YwYTEzYTsKICAtLXN0YXR1cy1zZXJpb3VzOiAjZWM4MzVhOwogIC0tc3RhdHVzLWNyaXRpY2FsOiAjZTA2MDVmOwoKICAtLWFjY2VudC1nb2xkOiAjZjJiMzBlOyAvKiBMUlMgYnJhbmQgZ29sZCDigJQgYWN0aXZlIG5hdiBpdGVtLCBwcmltYXJ5IGludGVyYWN0aXZlIGFjY2VudCAqLwoKICAtLXNlcmllcy0xOiAjMzk4N2U1OyAvKiBmYWNlYm9vayAqLwogIC0tc2VyaWVzLTI6ICMwMDgzMDA7IC8qIGluc3RhZ3JhbSAqLwogIC0tc2VyaWVzLTM6ICNkNTUxODE7IC8qIHRpa3RvayAqLwogIC0tc2VyaWVzLTQ6ICNjOTg1MDA7IC8qIGxpbmtlZGluICovCiAgLS1zZXJpZXMtNTogIzE5OWU3MDsgLyogdGhyZWFkcyAqLwogIC0tc2VyaWVzLTY6ICNkOTU5MjY7IC8qIHlvdXR1YmUgKi8KICAtLXNlcmllcy03OiAjOTA4NWU5OyAvKiByZXNlcnZlZCAqLwogIC0tc2VyaWVzLTg6ICNlNjY3Njc7IC8qIHJlc2VydmVkICovCgogIC0tcmFkaXVzLXNtOiAxMHB4OwogIC0tcmFkaXVzLW1kOiAxNHB4OwogIC0tcmFkaXVzLWxnOiAxOHB4OwoKICAtLXNoYWRvdy1jYXJkOiAwIDFweCAycHggcmdiYSgwLDAsMCwwLjIpLCAwIDhweCAyNHB4IC0xMHB4IHJnYmEoMCwwLDAsMC41KTsKICAtLXNoYWRvdy1ob3ZlcjogMCA2cHggMTJweCAtMnB4IHJnYmEoMCwwLDAsMC4zKSwgMCAxOHB4IDQwcHggLTE0cHggcmdiYSgwLDAsMCwwLjYpOwogIC0tc2hhZG93LW1vZGFsOiAwIDI0cHggNjRweCAtMTJweCByZ2JhKDAsMCwwLDAuNyk7CiAgLS1lYXNlOiBjdWJpYy1iZXppZXIoMC40LCAwLCAwLjIsIDEpOwp9CgoqIHsgYm94LXNpemluZzogYm9yZGVyLWJveDsgfQpodG1sLCBib2R5IHsgaGVpZ2h0OiAxMDAlOyB9CmJvZHkgewogIG1hcmdpbjogMDsKICBmb250LWZhbWlseTogdmFyKC0tZm9udC1zYW5zKTsKICBiYWNrZ3JvdW5kOiB2YXIoLS1wYWdlLXBsYW5lKTsKICBiYWNrZ3JvdW5kLWF0dGFjaG1lbnQ6IGZpeGVkOwogIGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOwogIC13ZWJraXQtZm9udC1zbW9vdGhpbmc6IGFudGlhbGlhc2VkOwogIC1tb3otb3N4LWZvbnQtc21vb3RoaW5nOiBncmF5c2NhbGU7Cn0KYnV0dG9uLCBzZWxlY3QsIGlucHV0LCB0ZXh0YXJlYSB7IGZvbnQtZmFtaWx5OiBpbmhlcml0OyB9CmgxLCBoMiwgaDMsIGg0IHsgZm9udC13ZWlnaHQ6IDcwMDsgbGV0dGVyLXNwYWNpbmc6IC0wLjAxZW07IH0KCjo6c2VsZWN0aW9uIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSAzMCUsIHRyYW5zcGFyZW50KTsgfQoKLyogQ3VzdG9tIHNjcm9sbGJhciDigJQgdGhpbiwgdW5vYnRydXNpdmUsIGZpdHMgdGhlIGdsYXNzIGFlc3RoZXRpYyAqLwo6Oi13ZWJraXQtc2Nyb2xsYmFyIHsgd2lkdGg6IDEwcHg7IGhlaWdodDogMTBweDsgfQo6Oi13ZWJraXQtc2Nyb2xsYmFyLXRyYWNrIHsgYmFja2dyb3VuZDogdHJhbnNwYXJlbnQ7IH0KOjotd2Via2l0LXNjcm9sbGJhci10aHVtYiB7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS10ZXh0LW11dGVkKSA0MCUsIHRyYW5zcGFyZW50KTsgYm9yZGVyLXJhZGl1czogMjBweDsgYm9yZGVyOiAycHggc29saWQgdHJhbnNwYXJlbnQ7IGJhY2tncm91bmQtY2xpcDogcGFkZGluZy1ib3g7IH0KOjotd2Via2l0LXNjcm9sbGJhci10aHVtYjpob3ZlciB7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS10ZXh0LW11dGVkKSA2MCUsIHRyYW5zcGFyZW50KTsgYmFja2dyb3VuZC1jbGlwOiBwYWRkaW5nLWJveDsgfQoKLmFwcC1zaGVsbCB7IGhlaWdodDogMTAwdmg7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiByb3c7IG92ZXJmbG93OiBoaWRkZW47IH0KCi8qIC0tLS0tLS0tLS0gU2lkZWJhciAtLS0tLS0tLS0tICovCi5zaWRlYmFyIHsKICBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBnYXA6IDRweDsKICB3aWR0aDogMjQwcHg7IGZsZXg6IDAgMCBhdXRvOwogIHBhZGRpbmc6IDIwcHggMTRweDsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zaWRlYmFyLWJnKTsKICBib3JkZXItcmlnaHQ6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOwogIHotaW5kZXg6IDIwOwogIG92ZXJmbG93LXk6IGF1dG87Cn0KLnNpZGViYXItYnJhbmQgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDhweDsgd2hpdGUtc3BhY2U6IG5vd3JhcDsgcGFkZGluZzogNHB4IDEwcHggMjBweDsgfQouYnJhbmQtbG9nbyB7IGhlaWdodDogMjhweDsgd2lkdGg6IGF1dG87IGRpc3BsYXk6IGJsb2NrOyBmbGV4LXNocmluazogMDsgb2JqZWN0LWZpdDogY29udGFpbjsgfQouYnJhbmQtdGl0bGUgeyBmb250LXdlaWdodDogNjAwOyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgbGV0dGVyLXNwYWNpbmc6IC0wLjAxZW07IH0KCi50YWJzIHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiAycHg7IGZsZXg6IDE7IHBvc2l0aW9uOiByZWxhdGl2ZTsgfQoudGFiLWJ0biB7CiAgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiAxMHB4OwogIGJvcmRlcjogbm9uZTsgYmFja2dyb3VuZDogdHJhbnNwYXJlbnQ7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7CiAgcGFkZGluZzogMTBweCAxMnB4OyBib3JkZXItcmFkaXVzOiAxMHB4OyBjdXJzb3I6IHBvaW50ZXI7IGZvbnQtc2l6ZTogMTRweDsgZm9udC13ZWlnaHQ6IDUwMDsKICB3aGl0ZS1zcGFjZTogbm93cmFwOyBwb3NpdGlvbjogcmVsYXRpdmU7IHRleHQtYWxpZ246IGxlZnQ7IHdpZHRoOiAxMDAlOwogIHRyYW5zaXRpb246IGNvbG9yIDE4MG1zIHZhcigtLWVhc2UpLCBiYWNrZ3JvdW5kIDE4MG1zIHZhcigtLWVhc2UpOwp9Ci50YWItYnRuIHN2ZyB7IGZsZXgtc2hyaW5rOiAwOyBvcGFjaXR5OiAwLjg7IH0KLnRhYi1idG4uaXMtYWN0aXZlIHN2ZyB7IG9wYWNpdHk6IDE7IGNvbG9yOiB2YXIoLS1zZXJpZXMtMSk7IH0KLnRhYi1idG46aG92ZXIgeyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDglLCB0cmFuc3BhcmVudCk7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyB9Ci50YWItYnRuLmlzLWFjdGl2ZSB7CiAgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IGZvbnQtd2VpZ2h0OiA2MDA7CiAgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSAxNiUsIHRyYW5zcGFyZW50KTsKICBhbmltYXRpb246IHRhYkluZGljYXRvckluIDIyMG1zIHZhcigtLWVhc2UpOwp9CkBrZXlmcmFtZXMgdGFiSW5kaWNhdG9ySW4geyBmcm9tIHsgb3BhY2l0eTogMDsgdHJhbnNmb3JtOiB0cmFuc2xhdGVYKC00cHgpOyB9IHRvIHsgb3BhY2l0eTogMTsgdHJhbnNmb3JtOiB0cmFuc2xhdGVYKDApOyB9IH0KLnNpZGViYXItZm9vdGVyIHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiAxMHB4OyBwYWRkaW5nLXRvcDogMTRweDsgbWFyZ2luLXRvcDogMTRweDsgYm9yZGVyLXRvcDogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IH0KCi5zaWRlYmFyLXVzZXIgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDEwcHg7IGZvbnQtc2l6ZTogMTNweDsgfQoKLyogLS0tLS0tLS0tLSBBdXRoIHNjcmVlbiAtLS0tLS0tLS0tICovCi5hdXRoLXNjcmVlbiB7CiAgbWluLWhlaWdodDogMTAwdmg7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogY2VudGVyOwogIGJhY2tncm91bmQ6IHZhcigtLXBhZ2UtcGxhbmUpOyBwYWRkaW5nOiAyMHB4Owp9Ci5hdXRoLWNhcmQgewogIHdpZHRoOiAxMDAlOyBtYXgtd2lkdGg6IDQwMHB4OyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTEpOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOwogIGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtbGcpOyBwYWRkaW5nOiAzMnB4OyBib3gtc2hhZG93OiB2YXIoLS1zaGFkb3ctbW9kYWwpOwogIGFuaW1hdGlvbjogbW9kYWxQYW5lbEluIDI2MG1zIHZhcigtLWVhc2UpOwp9Ci5hdXRoLWJyYW5kIHsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA4cHg7IG1hcmdpbi1ib3R0b206IDIycHg7IH0KLmF1dGgtYnJhbmQgLmJyYW5kLXRpdGxlIHsgZm9udC13ZWlnaHQ6IDcwMDsgZm9udC1zaXplOiAxN3B4OyB9Ci5hdXRoLWJyYW5kIC5icmFuZC1sb2dvIHsgaGVpZ2h0OiAzNnB4OyB9Ci5hdXRoLWZvcm0geyBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBnYXA6IDE0cHg7IG1hcmdpbi10b3A6IDE2cHg7IH0KLmF1dGgtZm9ybSAuZm9ybS1maWVsZCBpbnB1dCB7IHdpZHRoOiAxMDAlOyB9Ci5hdXRoLWVycm9yIHsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IGZvbnQtc2l6ZTogMTJweDsgbWluLWhlaWdodDogMTZweDsgfQoKLyogLS0tLS0tLS0tLSBGaWx0ZXIgYmFyIC0tLS0tLS0tLS0gKi8KLmZpbHRlci1iYXIgewogIGRpc3BsYXk6IGZsZXg7IGZsZXgtd3JhcDogd3JhcDsgYWxpZ24taXRlbXM6IGVuZDsgZ2FwOiAxNnB4OwogIHBhZGRpbmc6IDE0cHggMjBweDsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgYm9yZGVyLWJvdHRvbTogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7CiAgcG9zaXRpb246IHN0aWNreTsgdG9wOiAwOyB6LWluZGV4OiAxOTsKfQouZmlsdGVyLWZpZWxkIHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiA1cHg7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgfQouZmlsdGVyLWZpZWxkIGxhYmVsIHsgZm9udC13ZWlnaHQ6IDYwMDsgfQouZmlsdGVyLXByZXNldHMgeyBmbGV4LWRpcmVjdGlvbjogcm93OyBnYXA6IDZweDsgfQouZmlsdGVyLXByZXNldHMgYnV0dG9uIHsKICBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTIpOyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOwogIGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBib3JkZXItcmFkaXVzOiAyMHB4OyBwYWRkaW5nOiA3cHggMTNweDsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNTAwOyBjdXJzb3I6IHBvaW50ZXI7CiAgdHJhbnNpdGlvbjogY29sb3IgMTgwbXMgdmFyKC0tZWFzZSksIGJhY2tncm91bmQgMTgwbXMgdmFyKC0tZWFzZSksIHRyYW5zZm9ybSAxNTBtcyB2YXIoLS1lYXNlKTsKfQouZmlsdGVyLXByZXNldHMgYnV0dG9uOmhvdmVyIHsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgtMXB4KTsgfQouZmlsdGVyLXByZXNldHMgYnV0dG9uOmFjdGl2ZSB7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgwKSBzY2FsZSgwLjk2KTsgfQouZmlsdGVyLXByZXNldHMgYnV0dG9uLmlzLWFjdGl2ZSB7IGJhY2tncm91bmQ6IHZhcigtLXNlcmllcy0xKTsgY29sb3I6ICNmZmY7IGJvcmRlci1jb2xvcjogdHJhbnNwYXJlbnQ7IGJveC1zaGFkb3c6IDAgNHB4IDE0cHggLTVweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDYwJSwgdHJhbnNwYXJlbnQpOyB9CgovKiAtLS0tLS0tLS0tIE1haW4gY29sdW1uIChzaXRzIGJlc2lkZSB0aGUgc2lkZWJhcjsgc2Nyb2xscyBpbmRlcGVuZGVudGx5IHNvIHRoZSBzaWRlYmFyIHN0YXlzIGZ1bGx5IHZpc2libGUpIC0tLS0tLS0tLS0gKi8KLm1haW4tY29sIHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZmxleDogMSAxIGF1dG87IG1pbi13aWR0aDogMDsgaGVpZ2h0OiAxMDAlOyBvdmVyZmxvdy15OiBhdXRvOyB9CgovKiAtLS0tLS0tLS0tIFZpZXcgYXJlYSAtLS0tLS0tLS0tICovCi52aWV3LWFyZWEgeyBmbGV4OiAxOyBwYWRkaW5nOiAyNHB4OyBtYXgtd2lkdGg6IDE4MDBweDsgd2lkdGg6IDEwMCU7IG1hcmdpbjogMCBhdXRvOyB9Ci52aWV3IHsgZGlzcGxheTogbm9uZTsgfQoudmlldy5pcy1hY3RpdmUgeyBkaXNwbGF5OiBibG9jazsgYW5pbWF0aW9uOiB2aWV3RmFkZUluIDI2MG1zIHZhcigtLWVhc2UpOyB9CkBrZXlmcmFtZXMgdmlld0ZhZGVJbiB7CiAgZnJvbSB7IG9wYWNpdHk6IDA7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSg2cHgpOyB9CiAgdG8geyBvcGFjaXR5OiAxOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMCk7IH0KfQoKLnNlY3Rpb24tdGl0bGUgeyBmb250LXNpemU6IDE2cHg7IGZvbnQtd2VpZ2h0OiA3MDA7IG1hcmdpbjogMzJweCAwIDE0cHg7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyBsZXR0ZXItc3BhY2luZzogLTAuMDFlbTsgfQouc2VjdGlvbi10aXRsZTpmaXJzdC1jaGlsZCB7IG1hcmdpbi10b3A6IDA7IH0KCi8qIC0tLS0tLS0tLS0gSW5wdXRzIOKAlCBvbmUgc2hhcmVkIGdsYXNzIHRyZWF0bWVudCBmb3IgZXZlcnkgdGV4dCBpbnB1dCwgc2VsZWN0LCBhbmQgZGF0ZSBwaWNrZXIgLS0tLS0tLS0tLSAqLwouZmlsdGVyLWZpZWxkIHNlbGVjdCwgLmZpbHRlci1maWVsZCBpbnB1dFt0eXBlPSJkYXRlIl0sCi5mb3JtLWZpZWxkIGlucHV0LCAuZm9ybS1maWVsZCBzZWxlY3QsIC5mb3JtLWZpZWxkIHRleHRhcmVhLAouZGFzaGJvYXJkLWNvbnRyb2xzIHNlbGVjdCwgLnJlY29yZHMtc2VhcmNoIGlucHV0LAouZmllbGQtaW5saW5lIHNlbGVjdCwgLmZpZWxkLWlubGluZSBpbnB1dCwKLmNvbmZsaWN0LXJvdyBzZWxlY3QsIC5jYXJkLWhlYWRlciBzZWxlY3QgewogIGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1zbSk7CiAgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0yKTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IGZvbnQtc2l6ZTogMTNweDsKICBwYWRkaW5nOiA4cHggMTJweDsgbWluLXdpZHRoOiAxNDBweDsKICB0cmFuc2l0aW9uOiBib3JkZXItY29sb3IgMTYwbXMgdmFyKC0tZWFzZSksIGJveC1zaGFkb3cgMTYwbXMgdmFyKC0tZWFzZSk7Cn0KLmZpbHRlci1maWVsZCBzZWxlY3Q6aG92ZXIsIC5maWx0ZXItZmllbGQgaW5wdXRbdHlwZT0iZGF0ZSJdOmhvdmVyLAouZm9ybS1maWVsZCBpbnB1dDpob3ZlciwgLmZvcm0tZmllbGQgc2VsZWN0OmhvdmVyLAouZGFzaGJvYXJkLWNvbnRyb2xzIHNlbGVjdDpob3ZlciwgLnJlY29yZHMtc2VhcmNoIGlucHV0OmhvdmVyLAouZmllbGQtaW5saW5lIHNlbGVjdDpob3ZlciwgLmZpZWxkLWlubGluZSBpbnB1dDpob3ZlciwKLmNvbmZsaWN0LXJvdyBzZWxlY3Q6aG92ZXIsIC5jYXJkLWhlYWRlciBzZWxlY3Q6aG92ZXIgewogIGJvcmRlci1jb2xvcjogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSAzNSUsIHZhcigtLWJvcmRlcikpOwp9Ci8qIEEgPHNlbGVjdD4ncyBvd24gYmFja2dyb3VuZCBpcyBhIHRyYW5zbHVjZW50IGdsYXNzIHRpbnQgbWVhbnQgdG8gYmxlbmQgd2l0aAogICB0aGUgcGFnZSBiZWhpbmQgaXQg4oCUIGJ1dCBpdHMgZHJvcGRvd24gcG9wdXAgcmVuZGVycyBvbiBhbiBpc29sYXRlZCBvcGFxdWUKICAgY2FudmFzLCBzbyB0aGF0IHNhbWUgdHJhbnNsdWNlbnQgdmFsdWUgc2hvd3MgdXAgdGhlcmUgYXMgcGxhaW4gd2hpdGUKICAgaW5zdGVhZCBvZiBkYXJrLiBFdmVyeSA8b3B0aW9uPiwgaW4gZXZlcnkgc2VsZWN0IGluIHRoZSBhcHAsIG5lZWRzIGFuCiAgIGV4cGxpY2l0IHNvbGlkIGRhcmsgYmFja2dyb3VuZC90ZXh0IGNvbG9yIHNvIHRoZSBwb3B1cCBtYXRjaGVzIHRoZSB0aGVtZS4gKi8Kb3B0aW9uIHsgYmFja2dyb3VuZC1jb2xvcjogdmFyKC0tc3VyZmFjZS1zb2xpZCk7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyB9CgouZmlsdGVyLWZpZWxkIHNlbGVjdDpmb2N1cywgLmZpbHRlci1maWVsZCBpbnB1dFt0eXBlPSJkYXRlIl06Zm9jdXMsCi5mb3JtLWZpZWxkIGlucHV0OmZvY3VzLCAuZm9ybS1maWVsZCBzZWxlY3Q6Zm9jdXMsIC5mb3JtLWZpZWxkIHRleHRhcmVhOmZvY3VzLAouZGFzaGJvYXJkLWNvbnRyb2xzIHNlbGVjdDpmb2N1cywgLnJlY29yZHMtc2VhcmNoIGlucHV0OmZvY3VzLAouZmllbGQtaW5saW5lIHNlbGVjdDpmb2N1cywgLmZpZWxkLWlubGluZSBpbnB1dDpmb2N1cywKLmNvbmZsaWN0LXJvdyBzZWxlY3Q6Zm9jdXMsIC5jYXJkLWhlYWRlciBzZWxlY3Q6Zm9jdXMsCi5hdXRoLWZvcm0gaW5wdXQ6Zm9jdXMgewogIG91dGxpbmU6IG5vbmU7IGJvcmRlci1jb2xvcjogdmFyKC0tc2VyaWVzLTEpOwogIGJveC1zaGFkb3c6IDAgMCAwIDNweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDE4JSwgdHJhbnNwYXJlbnQpOwp9CgovKiAtLS0tLS0tLS0tIFN0YXQgdGlsZXMgLS0tLS0tLS0tLSAqLwouc3RhdC1ncmlkIHsKICBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdChhdXRvLWZpdCwgbWlubWF4KDE4MHB4LCAxZnIpKTsgZ2FwOiAxNHB4Owp9Ci5zdGF0LXRpbGUgewogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1tZCk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHBhZGRpbmc6IDE2cHggMThweDsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWNhcmQpOwogIHRyYW5zaXRpb246IHRyYW5zZm9ybSAyMDBtcyB2YXIoLS1lYXNlKSwgYm94LXNoYWRvdyAyMDBtcyB2YXIoLS1lYXNlKTsKICBhbmltYXRpb246IGNhcmRJbiAzMjBtcyB2YXIoLS1lYXNlKSBiYWNrd2FyZHM7Cn0KLnN0YXQtdGlsZTpob3ZlciB7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgtM3B4KTsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWhvdmVyKTsgfQouc3RhdC1sYWJlbCB7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgZm9udC13ZWlnaHQ6IDYwMDsgfQouc3RhdC12YWx1ZSB7IGZvbnQtc2l6ZTogMjdweDsgZm9udC13ZWlnaHQ6IDcwMDsgbWFyZ2luLXRvcDogNXB4OyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgbGV0dGVyLXNwYWNpbmc6IC0wLjAyZW07IH0KLnN0YXQtZGVsdGEgeyBmb250LXNpemU6IDEycHg7IG1hcmdpbi10b3A6IDdweDsgZm9udC13ZWlnaHQ6IDYwMDsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA0cHg7IH0KLnN0YXQtZGVsdGEudXAgeyBjb2xvcjogdmFyKC0tc3VjY2Vzcy10ZXh0KTsgfQouc3RhdC1kZWx0YS5kb3duIHsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IH0KLnN0YXQtZGVsdGEuZmxhdCB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgfQouc3RhdC1kZWx0YS51cDo6YmVmb3JlIHsgY29udGVudDogJ+KGkSc7IH0KLnN0YXQtZGVsdGEuZG93bjo6YmVmb3JlIHsgY29udGVudDogJ+KGkyc7IH0KCi8qIC0tLS0tLS0tLS0gRGFzaGJvYXJkIEtQSSBncmlkIOKAlCBjb21wYWN0LCBzaW5nbGUtcm93LW9uLWRlc2t0b3AgbGF5b3V0LgogICBTY29wZWQgdG8gI2twaUdyaWQgc3BlY2lmaWNhbGx5IChub3QgdGhlIHNoYXJlZCAuc3RhdC1ncmlkLy5zdGF0LXRpbGUKICAgY2xhc3Nlcywgd2hpY2ggQ29tcGFyaXNvbnMgYW5kIHRoZSBVcGxvYWQgcHJldmlldyBzdW1tYXJ5IGFsc28gdXNlKSBzbwogICB0aGlzIGNvbXBhY3RpbmcgZG9lc24ndCBhZmZlY3QgdGhvc2Ugb3RoZXIgc3RhdC10aWxlIGdyaWRzLiAxMCBncmlkIHVuaXRzCiAgIHRvdGFsOiA3IHN0YW5kYXJkIEtQSSB0aWxlcyBhdCAxIHVuaXQgZWFjaCArIEJlc3QgUGVyZm9ybWluZyBQb3N0IGF0IDMKICAgdW5pdHMgKGEgM3gtd2lkZSBsYW5kc2NhcGUgY2FyZCksIGFsbCBzaGFyaW5nIG9uZSBmaXhlZCByb3cgaGVpZ2h0LiAtLS0tLS0tLS0tICovCiNrcGlHcmlkIHsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiByZXBlYXQoMTAsIG1pbm1heCgwLCAxZnIpKTsgZ2FwOiAxMnB4OyB9CiNrcGlHcmlkIC5zdGF0LXRpbGUgewogIGhlaWdodDogMTMycHg7IGJveC1zaXppbmc6IGJvcmRlci1ib3g7IG92ZXJmbG93OiBoaWRkZW47CiAgcGFkZGluZzogMTZweCAxOHB4OwogIGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGp1c3RpZnktY29udGVudDogY2VudGVyOwp9CiNrcGlHcmlkIC5zdGF0LWxhYmVsIHsgZm9udC1zaXplOiAxMnB4OyBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBhbGlnbi1pdGVtczogZmxleC1zdGFydDsgZ2FwOiA4cHg7IH0KI2twaUdyaWQgLnN0YXQtdmFsdWUgeyBmb250LXNpemU6IDMycHg7IG1hcmdpbi10b3A6IDhweDsgbGluZS1oZWlnaHQ6IDEuMTsgfQoja3BpR3JpZCAuc3RhdC1kZWx0YSB7IGZvbnQtc2l6ZTogMTNweDsgbWFyZ2luLXRvcDogOHB4OyB9Cgouc3RhdC1pY29uIHsKICB3aWR0aDogMjhweDsgaGVpZ2h0OiAyOHB4OyBmbGV4OiAwIDAgYXV0bzsgYm9yZGVyLXJhZGl1czogNTAlOwogIGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogY2VudGVyOwogIGNvbG9yOiAjZmZmOwp9Ci5zdGF0LWljb24udjEgeyBiYWNrZ3JvdW5kOiB2YXIoLS1zZXJpZXMtMSk7IH0KLnN0YXQtaWNvbi52MiB7IGJhY2tncm91bmQ6IHZhcigtLXNlcmllcy0yKTsgfQouc3RhdC1pY29uLnYzIHsgYmFja2dyb3VuZDogdmFyKC0tc2VyaWVzLTMpOyB9Ci5zdGF0LWljb24udjQgeyBiYWNrZ3JvdW5kOiB2YXIoLS1zZXJpZXMtNCk7IH0KLnN0YXQtaWNvbi52NSB7IGJhY2tncm91bmQ6IHZhcigtLXNlcmllcy01KTsgfQouc3RhdC1pY29uLnY2IHsgYmFja2dyb3VuZDogdmFyKC0tc2VyaWVzLTYpOyB9Ci5zdGF0LWljb24uZ29sZCB7IGJhY2tncm91bmQ6IHZhcigtLWFjY2VudC1nb2xkKTsgfQpAbWVkaWEgKG1heC13aWR0aDogOTAwcHgpIHsgI2twaUdyaWQgeyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdCg1LCBtaW5tYXgoMCwgMWZyKSk7IH0gfQpAbWVkaWEgKG1heC13aWR0aDogNjQwcHgpIHsgI2twaUdyaWQgeyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdCgyLCBtaW5tYXgoMCwgMWZyKSk7IH0gfQoKLmluc2lnaHRzLWxpc3QgeyBsaXN0LXN0eWxlOiBub25lOyBtYXJnaW46IDA7IHBhZGRpbmc6IDA7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGdhcDogMTBweDsgfQouaW5zaWdodHMtbGlzdCBsaSB7CiAgZm9udC1zaXplOiAxM3B4OyBsaW5lLWhlaWdodDogMS41OyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyBwYWRkaW5nLWxlZnQ6IDE4cHg7IHBvc2l0aW9uOiByZWxhdGl2ZTsKfQouaW5zaWdodHMtbGlzdCBsaTo6YmVmb3JlIHsKICBjb250ZW50OiAn4pymJzsgcG9zaXRpb246IGFic29sdXRlOyBsZWZ0OiAwOyBjb2xvcjogdmFyKC0tc2VyaWVzLTEpOyBmb250LXNpemU6IDExcHg7IHRvcDogMnB4Owp9CgpAa2V5ZnJhbWVzIGNhcmRJbiB7CiAgZnJvbSB7IG9wYWNpdHk6IDA7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgxMHB4KTsgfQogIHRvIHsgb3BhY2l0eTogMTsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDApOyB9Cn0KCi8qIC0tLS0tLS0tLS0gQ2FyZHMgLyBjaGFydHMgLS0tLS0tLS0tLSAqLwouY2FyZC1ncmlkIHsgZGlzcGxheTogZ3JpZDsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiAyZnIgMWZyOyBnYXA6IDE2cHg7IGFsaWduLWl0ZW1zOiBzdGFydDsgfQouY2FyZC1ncmlkLmV2ZW4geyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IDFmciAxZnI7IH0KQG1lZGlhIChtYXgtd2lkdGg6IDkwMHB4KSB7IC5jYXJkLWdyaWQsIC5jYXJkLWdyaWQuZXZlbiB7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogMWZyOyB9IH0KLmNhcmQgewogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1sZyk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHBhZGRpbmc6IDE4cHg7IGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdy1jYXJkKTsKICB0cmFuc2l0aW9uOiBib3gtc2hhZG93IDIyMG1zIHZhcigtLWVhc2UpLCB0cmFuc2Zvcm0gMjIwbXMgdmFyKC0tZWFzZSk7CiAgYW5pbWF0aW9uOiBjYXJkSW4gMzIwbXMgdmFyKC0tZWFzZSkgYmFja3dhcmRzOwp9Ci5jYXJkOmhvdmVyIHsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWhvdmVyKTsgfQouY2FyZC1oZWFkZXIgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6IHNwYWNlLWJldHdlZW47IGdhcDogOHB4OyBtYXJnaW4tYm90dG9tOiAxNHB4OyB9Ci5jYXJkLWhlYWRlciBoMyB7IGZvbnQtc2l6ZTogMTRweDsgbWFyZ2luOiAwOyBmb250LXdlaWdodDogNzAwOyBsZXR0ZXItc3BhY2luZzogLTAuMDA1ZW07IH0KLmNhcmQtaGVhZGVyIHNlbGVjdCB7IGZvbnQtc2l6ZTogMTJweDsgcGFkZGluZzogNnB4IDEwcHg7IG1pbi13aWR0aDogMDsgfQouY2hhcnQtd3JhcCB7IHBvc2l0aW9uOiByZWxhdGl2ZTsgaGVpZ2h0OiAyODBweDsgfQouY2hhcnQtd3JhcC50YWxsIHsgaGVpZ2h0OiAzNDBweDsgfQoKLmxlZ2VuZC1yb3cgeyBkaXNwbGF5OiBmbGV4OyBmbGV4LXdyYXA6IHdyYXA7IGdhcDogMTJweDsgbWFyZ2luLXRvcDogMTBweDsgZm9udC1zaXplOiAxMnB4OyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyB9Ci5sZWdlbmQtaXRlbSB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogNnB4OyB9Ci5sZWdlbmQtc3dhdGNoIHsgd2lkdGg6IDEwcHg7IGhlaWdodDogMTBweDsgYm9yZGVyLXJhZGl1czogM3B4OyBkaXNwbGF5OiBpbmxpbmUtYmxvY2s7IH0KLmxlZ2VuZC1saW5lIHsgd2lkdGg6IDE0cHg7IGhlaWdodDogMnB4OyBib3JkZXItcmFkaXVzOiAycHg7IGRpc3BsYXk6IGlubGluZS1ibG9jazsgfQoKLyogLS0tLS0tLS0tLSBUYWJsZXMg4oCUIHByZW1pdW0gZGF0YWJhc2UgZmVlbCwgbm90IGEgc3ByZWFkc2hlZXQgLS0tLS0tLS0tLSAqLwoudGFibGUtc2Nyb2xsIHsKICBvdmVyZmxvdy14OiBhdXRvOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtbWQpOwogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMik7Cn0KLmRhdGEtdGFibGUgeyB3aWR0aDogMTAwJTsgYm9yZGVyLWNvbGxhcHNlOiBzZXBhcmF0ZTsgYm9yZGVyLXNwYWNpbmc6IDA7IGZvbnQtc2l6ZTogMTNweDsgfQouZGF0YS10YWJsZSB0aCwgLmRhdGEtdGFibGUgdGQgeyB0ZXh0LWFsaWduOiBsZWZ0OyBwYWRkaW5nOiAxMXB4IDE0cHg7IGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1ncmlkbGluZSk7IHdoaXRlLXNwYWNlOiBub3dyYXA7IH0KLmRhdGEtdGFibGUgdGQud3JhcCB7IHdoaXRlLXNwYWNlOiBub3JtYWw7IH0KLmRhdGEtdGFibGUgdGhlYWQgdGggewogIGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IGZvbnQtd2VpZ2h0OiA2MDA7IGZvbnQtc2l6ZTogMTFweDsgdGV4dC10cmFuc2Zvcm06IHVwcGVyY2FzZTsgbGV0dGVyLXNwYWNpbmc6IDAuMDRlbTsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTEpOyBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgcG9zaXRpb246IHN0aWNreTsgdG9wOiAwOyB6LWluZGV4OiAxOwp9Ci5kYXRhLXRhYmxlIHRoZWFkIHRoLnNvcnRhYmxlLXRoIHsgY3Vyc29yOiBwb2ludGVyOyB1c2VyLXNlbGVjdDogbm9uZTsgdHJhbnNpdGlvbjogY29sb3IgMTUwbXMgdmFyKC0tZWFzZSk7IH0KLmRhdGEtdGFibGUgdGhlYWQgdGguc29ydGFibGUtdGg6aG92ZXIgeyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgfQouZGF0YS10YWJsZSB0aGVhZCB0aCAuc29ydC1hcnJvdyB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgZm9udC1zaXplOiAxMHB4OyBtYXJnaW4tbGVmdDogMnB4OyB9Ci5kYXRhLXRhYmxlIHRoZWFkIHRoLnNvcnRhYmxlLXRoOmhvdmVyIC5zb3J0LWFycm93IHsgY29sb3I6IHZhcigtLXNlcmllcy0xKTsgfQouZGF0YS10YWJsZSB0Ym9keSB0cjpudGgtY2hpbGQoZXZlbikgeyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tdGV4dC1tdXRlZCkgNCUsIHRyYW5zcGFyZW50KTsgfQouZGF0YS10YWJsZSB0ZC5udW0geyBmb250LXZhcmlhbnQtbnVtZXJpYzogdGFidWxhci1udW1zOyB0ZXh0LWFsaWduOiByaWdodDsgfQouZGF0YS10YWJsZSB0aC5udW0geyB0ZXh0LWFsaWduOiByaWdodDsgfQouZGF0YS10YWJsZSB0Ym9keSB0ciB7IHRyYW5zaXRpb246IGJhY2tncm91bmQgMTUwbXMgdmFyKC0tZWFzZSk7IH0KLmRhdGEtdGFibGUgdGJvZHkgdHI6aG92ZXIgeyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDclLCB0cmFuc3BhcmVudCk7IH0KLmRhdGEtdGFibGUgdGJvZHkgdHI6bGFzdC1jaGlsZCB0ZCB7IGJvcmRlci1ib3R0b206IG5vbmU7IH0KLnBsYXRmb3JtLXBpbGwgewogIGRpc3BsYXk6IGlubGluZS1mbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDZweDsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNjAwOwogIHBhZGRpbmc6IDRweCAxMHB4OyBib3JkZXItcmFkaXVzOiAyMHB4OyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTEpOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOwp9Ci5wbGF0Zm9ybS1kb3QgeyB3aWR0aDogOHB4OyBoZWlnaHQ6IDhweDsgYm9yZGVyLXJhZGl1czogNTAlOyB9CgovKiAtLS0tLS0tLS0tIEJ1dHRvbnMg4oCUIG5ldmVyIGZsYXQ6IHNvZnQgc2hhZG93LCBob3ZlciBsaWZ0LCBwcmVzcyBzY2FsZSAtLS0tLS0tLS0tICovCi5idG4gewogIGRpc3BsYXk6IGlubGluZS1mbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6IGNlbnRlcjsgZ2FwOiA2cHg7CiAgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0yKTsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHBhZGRpbmc6IDlweCAxN3B4OyBib3JkZXItcmFkaXVzOiAxMXB4OyBjdXJzb3I6IHBvaW50ZXI7IGZvbnQtc2l6ZTogMTNweDsgZm9udC13ZWlnaHQ6IDYwMDsKICBib3gtc2hhZG93OiAwIDFweCAycHggcmdiYSgxNSwxNywyMSwwLjA0KTsKICB0cmFuc2l0aW9uOiB0cmFuc2Zvcm0gMTUwbXMgdmFyKC0tZWFzZSksIGJveC1zaGFkb3cgMTUwbXMgdmFyKC0tZWFzZSksIGZpbHRlciAxNTBtcyB2YXIoLS1lYXNlKSwgYmFja2dyb3VuZCAxNTBtcyB2YXIoLS1lYXNlKTsKfQouYnRuIHN2ZyB7IGZsZXgtc2hyaW5rOiAwOyB9Ci5idG46aG92ZXIgeyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoLTFweCk7IGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdy1ob3Zlcik7IGZpbHRlcjogYnJpZ2h0bmVzcygxLjAyKTsgfQouYnRuOmFjdGl2ZSB7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgwKSBzY2FsZSgwLjk2KTsgYm94LXNoYWRvdzogMCAxcHggMnB4IHJnYmEoMTUsMTcsMjEsMC4wNik7IH0KLmJ0bi5wcmltYXJ5IHsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zZXJpZXMtMSk7IGNvbG9yOiAjZmZmOyBib3JkZXItY29sb3I6IHRyYW5zcGFyZW50OwogIGJveC1zaGFkb3c6IDAgNHB4IDE0cHggLTVweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDY1JSwgdHJhbnNwYXJlbnQpOwp9Ci5idG4ucHJpbWFyeTpob3ZlciB7IGZpbHRlcjogYnJpZ2h0bmVzcygxLjA3KTsgYm94LXNoYWRvdzogMCA4cHggMjJweCAtNnB4IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zZXJpZXMtMSkgNzAlLCB0cmFuc3BhcmVudCk7IH0KLmJ0bi5kYW5nZXIgewogIGJhY2tncm91bmQ6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IGNvbG9yOiAjZmZmOyBib3JkZXItY29sb3I6IHRyYW5zcGFyZW50OwogIGJveC1zaGFkb3c6IDAgNHB4IDE0cHggLTVweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc3RhdHVzLWNyaXRpY2FsKSA1NSUsIHRyYW5zcGFyZW50KTsKfQouYnRuLmRhbmdlcjpob3ZlciB7IGZpbHRlcjogYnJpZ2h0bmVzcygxLjA2KTsgYm94LXNoYWRvdzogMCA4cHggMjJweCAtNnB4IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zdGF0dXMtY3JpdGljYWwpIDYwJSwgdHJhbnNwYXJlbnQpOyB9Ci5idG4uc3VjY2VzcyB7CiAgYmFja2dyb3VuZDogdmFyKC0tc3RhdHVzLWdvb2QpOyBjb2xvcjogI2ZmZjsgYm9yZGVyLWNvbG9yOiB0cmFuc3BhcmVudDsKICBib3gtc2hhZG93OiAwIDRweCAxNHB4IC01cHggY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXN0YXR1cy1nb29kKSA1NSUsIHRyYW5zcGFyZW50KTsKfQouYnRuLnN1Y2Nlc3M6aG92ZXIgeyBmaWx0ZXI6IGJyaWdodG5lc3MoMS4wNik7IGJveC1zaGFkb3c6IDAgOHB4IDIycHggLTZweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc3RhdHVzLWdvb2QpIDYwJSwgdHJhbnNwYXJlbnQpOyB9Ci5idG46ZGlzYWJsZWQgeyBvcGFjaXR5OiAwLjQ1OyBjdXJzb3I6IG5vdC1hbGxvd2VkOyB0cmFuc2Zvcm06IG5vbmU7IGJveC1zaGFkb3c6IG5vbmU7IGZpbHRlcjogbm9uZTsgfQouYnRuLXJvdyB7IGRpc3BsYXk6IGZsZXg7IGdhcDogOHB4OyBmbGV4LXdyYXA6IHdyYXA7IH0KCi8qIC0tLS0tLS0tLS0gVXBsb2FkIC0tLS0tLS0tLS0gKi8KLmRyb3B6b25lIHsKICBib3JkZXI6IDJweCBkYXNoZWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLWxnKTsgcGFkZGluZzogNDBweCAyMHB4OwogIHRleHQtYWxpZ246IGNlbnRlcjsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIGN1cnNvcjogcG9pbnRlcjsgdHJhbnNpdGlvbjogYm9yZGVyLWNvbG9yIDIwMG1zIHZhcigtLWVhc2UpLCBiYWNrZ3JvdW5kIDIwMG1zIHZhcigtLWVhc2UpLCB0cmFuc2Zvcm0gMjAwbXMgdmFyKC0tZWFzZSk7Cn0KLmRyb3B6b25lOmhvdmVyIHsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKC0xcHgpOyB9Ci5kcm9wem9uZS5pcy1kcmFnIHsgYm9yZGVyLWNvbG9yOiB2YXIoLS1zZXJpZXMtMSk7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zZXJpZXMtMSkgNiUsIHZhcigtLXN1cmZhY2UtMikpOyB0cmFuc2Zvcm06IHNjYWxlKDEuMDA1KTsgfQouZHJvcHpvbmUgaDMgeyBtYXJnaW46IDAgMCA2cHg7IGZvbnQtc2l6ZTogMTVweDsgfQouZHJvcHpvbmUgcCB7IG1hcmdpbjogMDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgZm9udC1zaXplOiAxM3B4OyB9Ci5kcm9wem9uZSBpbnB1dFt0eXBlPSJmaWxlIl0geyBkaXNwbGF5OiBub25lOyB9CgouY29uZmxpY3QtbGlzdCB7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGdhcDogOHB4OyBtYXJnaW46IDEycHggMDsgfQouY29uZmxpY3Qtcm93IHsKICBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6IHNwYWNlLWJldHdlZW47IGdhcDogMTJweDsKICBwYWRkaW5nOiAxMXB4IDE0cHg7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1zbSk7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMik7CiAgdHJhbnNpdGlvbjogYm94LXNoYWRvdyAxODBtcyB2YXIoLS1lYXNlKTsKfQouY29uZmxpY3Qtcm93OmhvdmVyIHsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWNhcmQpOyB9Ci5jb25mbGljdC1yb3cgLndlZWstbGFiZWwgeyBmb250LXdlaWdodDogNjAwOyBmb250LXNpemU6IDEzcHg7IH0KLmNvbmZsaWN0LXJvdyAud2Vlay1tZXRhIHsgZm9udC1zaXplOiAxMnB4OyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyB9Ci5jb25mbGljdC1yb3cgc2VsZWN0IHsgbWluLXdpZHRoOiAwOyB9CgouYmFkZ2UgeyBkaXNwbGF5OiBpbmxpbmUtYmxvY2s7IHBhZGRpbmc6IDNweCAxMHB4OyBib3JkZXItcmFkaXVzOiAyMHB4OyBmb250LXNpemU6IDExcHg7IGZvbnQtd2VpZ2h0OiA3MDA7IH0KLmJhZGdlLnN1Y2Nlc3MgeyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc3RhdHVzLWdvb2QpIDE4JSwgdHJhbnNwYXJlbnQpOyBjb2xvcjogdmFyKC0tc3RhdHVzLWdvb2QpOyB9Ci5iYWRnZS5wYXJ0aWFsIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXN0YXR1cy13YXJuaW5nKSAyNSUsIHRyYW5zcGFyZW50KTsgY29sb3I6ICM4YTYzMDA7IH0KLmJhZGdlLmZhaWxlZCB7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zdGF0dXMtY3JpdGljYWwpIDE4JSwgdHJhbnNwYXJlbnQpOyBjb2xvcjogdmFyKC0tc3RhdHVzLWNyaXRpY2FsKTsgfQouYmFkZ2UuZXJyb3Itc2V2IHsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IH0KLmJhZGdlLndhcm5pbmctc2V2IHsgY29sb3I6ICM4YTYzMDA7IH0KLmJhZGdlLnNraXAtc2V2IHsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyB9CgouaXNzdWVzLWxpc3QgeyBtYXgtaGVpZ2h0OiAyMjBweDsgb3ZlcmZsb3cteTogYXV0bzsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLXNtKTsgfQouaXNzdWUtcm93IHsgcGFkZGluZzogOXB4IDE0cHg7IGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1ncmlkbGluZSk7IGZvbnQtc2l6ZTogMTJweDsgfQouaXNzdWUtcm93Omxhc3QtY2hpbGQgeyBib3JkZXItYm90dG9tOiBub25lOyB9Ci5pc3N1ZS1yb3cgLnJvdy1ubyB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgbWFyZ2luLXJpZ2h0OiA2cHg7IH0KCi8qIC0tLS0tLS0tLS0gVG9hc3QgLS0tLS0tLS0tLSAqLwoudG9hc3Qtcm9vdCB7IHBvc2l0aW9uOiBmaXhlZDsgYm90dG9tOiAyMHB4OyByaWdodDogMjBweDsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiA4cHg7IHotaW5kZXg6IDEwMDsgfQoudG9hc3QgewogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1zbSk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHBhZGRpbmc6IDEycHggMTZweDsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LW1vZGFsKTsgZm9udC1zaXplOiAxM3B4OyBtYXgtd2lkdGg6IDM0MHB4OwogIGFuaW1hdGlvbjogdG9hc3QtaW4gMjIwbXMgdmFyKC0tZWFzZSk7Cn0KLnRvYXN0LnN1Y2Nlc3MgeyBib3JkZXItbGVmdDogM3B4IHNvbGlkIHZhcigtLXN0YXR1cy1nb29kKTsgfQoudG9hc3QuZXJyb3IgeyBib3JkZXItbGVmdDogM3B4IHNvbGlkIHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IH0KQGtleWZyYW1lcyB0b2FzdC1pbiB7IGZyb20geyBvcGFjaXR5OiAwOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMTBweCkgc2NhbGUoMC45OCk7IH0gdG8geyBvcGFjaXR5OiAxOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMCkgc2NhbGUoMSk7IH0gfQoKLyogLS0tLS0tLS0tLSBNaXNjIC0tLS0tLS0tLS0gKi8KLm11dGVkIHsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyB9Ci5lbXB0eS1zdGF0ZSB7CiAgcGFkZGluZzogNTZweCAyNHB4OyB0ZXh0LWFsaWduOiBjZW50ZXI7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7CiAgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiAxMnB4OwogIGFuaW1hdGlvbjogY2FyZEluIDI2MG1zIHZhcigtLWVhc2UpOwp9Ci5lbXB0eS1zdGF0ZSAuZW1wdHktaWNvbiB7CiAgd2lkdGg6IDUycHg7IGhlaWdodDogNTJweDsgYm9yZGVyLXJhZGl1czogMTZweDsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsganVzdGlmeS1jb250ZW50OiBjZW50ZXI7CiAgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSAxMCUsIHRyYW5zcGFyZW50KTsgY29sb3I6IHZhcigtLXNlcmllcy0xKTsKfQouZW1wdHktc3RhdGUgLmVtcHR5LXRpdGxlIHsgZm9udC1zaXplOiAxNHB4OyBmb250LXdlaWdodDogNjAwOyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgfQouZW1wdHktc3RhdGUgLmVtcHR5LW1lc3NhZ2UgeyBmb250LXNpemU6IDEzcHg7IG1heC13aWR0aDogMzYwcHg7IH0KLnNwaW5uZXIgeyB3aWR0aDogMTZweDsgaGVpZ2h0OiAxNnB4OyBib3JkZXItcmFkaXVzOiA1MCU7IGJvcmRlcjogMnB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci10b3AtY29sb3I6IHZhcigtLXNlcmllcy0xKTsgYW5pbWF0aW9uOiBzcGluIC42cyBsaW5lYXIgaW5maW5pdGU7IGRpc3BsYXk6IGlubGluZS1ibG9jazsgfQpAa2V5ZnJhbWVzIHNwaW4geyB0byB7IHRyYW5zZm9ybTogcm90YXRlKDM2MGRlZyk7IH0gfQoubG9hZGluZy1yb3cgeyBwYWRkaW5nOiA0MHB4IDIwcHg7IHRleHQtYWxpZ246IGNlbnRlcjsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgfQoKLyogU2tlbGV0b24gbG9hZGVycyDigJQgc2hpbW1lcmluZyBwbGFjZWhvbGRlcnMgc2hvd24gd2hpbGUgYSBzZWN0aW9uJ3MgZGF0YSBpcyBpbiBmbGlnaHQgKi8KLnNrZWxldG9uIHsKICBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtc20pOwogIGJhY2tncm91bmQ6IGxpbmVhci1ncmFkaWVudCgxMDBkZWcsIGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS10ZXh0LW11dGVkKSAxMiUsIHRyYW5zcGFyZW50KSAzMCUsIGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS10ZXh0LW11dGVkKSAyMiUsIHRyYW5zcGFyZW50KSA1MCUsIGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS10ZXh0LW11dGVkKSAxMiUsIHRyYW5zcGFyZW50KSA3MCUpOwogIGJhY2tncm91bmQtc2l6ZTogMjAwJSAxMDAlOwogIGFuaW1hdGlvbjogc2tlbGV0b25TaGltbWVyIDEuNHMgZWFzZS1pbi1vdXQgaW5maW5pdGU7Cn0KQGtleWZyYW1lcyBza2VsZXRvblNoaW1tZXIgeyBmcm9tIHsgYmFja2dyb3VuZC1wb3NpdGlvbjogMTUwJSAwOyB9IHRvIHsgYmFja2dyb3VuZC1wb3NpdGlvbjogLTUwJSAwOyB9IH0KLnNrZWxldG9uLXN0YXQtZ3JpZCB7IGRpc3BsYXk6IGdyaWQ7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogcmVwZWF0KGF1dG8tZml0LCBtaW5tYXgoMTgwcHgsIDFmcikpOyBnYXA6IDE0cHg7IH0KLnNrZWxldG9uLXRpbGUgeyBoZWlnaHQ6IDg0cHg7IH0KLnNrZWxldG9uLWNoYXJ0IHsgaGVpZ2h0OiAyODBweDsgd2lkdGg6IDEwMCU7IH0KLnNrZWxldG9uLXJvdyB7IGhlaWdodDogNDBweDsgbWFyZ2luLWJvdHRvbTogOHB4OyB9CgovKiBBbmltYXRlZCBob3Jpem9udGFsIGNvbXBhcmlzb24gYmFyIOKAlCBhIGxhYmVsZWQgcm93IHdpdGggYSB0cmFjayB0aGF0IGZpbGxzIGluIG9uIGluc2VydGlvbiAqLwouYmFyLXJvdyB7IGRpc3BsYXk6IGdyaWQ7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogbWlubWF4KDkwcHgsIDE0MHB4KSAxZnIgYXV0bzsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiAxMHB4OyBwYWRkaW5nOiA1cHggMDsgfQouYmFyLWxhYmVsIHsgZm9udC1zaXplOiAxMnB4OyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyBmb250LXdlaWdodDogNjAwOyB9Ci5iYXItdHJhY2sgeyBoZWlnaHQ6IDhweDsgYm9yZGVyLXJhZGl1czogNXB4OyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tdGV4dC1tdXRlZCkgMTQlLCB0cmFuc3BhcmVudCk7IG92ZXJmbG93OiBoaWRkZW47IH0KLmJhci1maWxsIHsgaGVpZ2h0OiAxMDAlOyB3aWR0aDogMCU7IGJvcmRlci1yYWRpdXM6IDVweDsgdHJhbnNpdGlvbjogd2lkdGggNzAwbXMgY3ViaWMtYmV6aWVyKDAuMTYsIDEsIDAuMywgMSk7IH0KLmJhci12YWx1ZSB7IGZvbnQtc2l6ZTogMTJweDsgZm9udC13ZWlnaHQ6IDcwMDsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IGZvbnQtdmFyaWFudC1udW1lcmljOiB0YWJ1bGFyLW51bXM7IHRleHQtYWxpZ246IHJpZ2h0OyBtaW4td2lkdGg6IDU2cHg7IH0KCkBtZWRpYSAocHJlZmVycy1yZWR1Y2VkLW1vdGlvbjogcmVkdWNlKSB7CiAgLmJhci1maWxsIHsgdHJhbnNpdGlvbi1kdXJhdGlvbjogMW1zOyB9CiAgLnNrZWxldG9uIHsgYW5pbWF0aW9uLWR1cmF0aW9uOiAxbXM7IH0KICAuY2FyZCwgLnN0YXQtdGlsZSB7IGFuaW1hdGlvbi1kdXJhdGlvbjogMW1zOyB9Cn0KCi50d28tY29sIHsgZGlzcGxheTogZ3JpZDsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiAxZnIgMWZyOyBnYXA6IDE2cHg7IH0KQG1lZGlhIChtYXgtd2lkdGg6IDkwMHB4KSB7IC50d28tY29sIHsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiAxZnI7IH0gfQoKLm1vZGUtdGFicyB7IGRpc3BsYXk6IGZsZXg7IGdhcDogNnB4OyBmbGV4LXdyYXA6IHdyYXA7IG1hcmdpbi1ib3R0b206IDE2cHg7IH0KLm1vZGUtdGFicyBidXR0b24gewogIGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHBhZGRpbmc6IDdweCAxNHB4OyBib3JkZXItcmFkaXVzOiAyMHB4OyBmb250LXNpemU6IDEycHg7IGZvbnQtd2VpZ2h0OiA2MDA7IGN1cnNvcjogcG9pbnRlcjsKICB0cmFuc2l0aW9uOiBjb2xvciAxODBtcyB2YXIoLS1lYXNlKSwgYmFja2dyb3VuZCAxODBtcyB2YXIoLS1lYXNlKSwgdHJhbnNmb3JtIDE1MG1zIHZhcigtLWVhc2UpOwp9Ci5tb2RlLXRhYnMgYnV0dG9uOmhvdmVyIHsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKC0xcHgpOyB9Ci5tb2RlLXRhYnMgYnV0dG9uLmlzLWFjdGl2ZSB7IGJhY2tncm91bmQ6IHZhcigtLXNlcmllcy0xKTsgY29sb3I6ICNmZmY7IGJvcmRlci1jb2xvcjogdHJhbnNwYXJlbnQ7IGJveC1zaGFkb3c6IDAgNHB4IDE0cHggLTVweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDYwJSwgdHJhbnNwYXJlbnQpOyB9CgouZmllbGQtaW5saW5lIHsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA4cHg7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgfQouZmllbGQtaW5saW5lIHNlbGVjdCwgLmZpZWxkLWlubGluZSBpbnB1dCB7IG1pbi13aWR0aDogMDsgcGFkZGluZzogNnB4IDEwcHg7IH0KCi8qIC0tLS0tLS0tLS0gUGxhdGZvcm0gUGVyZm9ybWFuY2UgQ29tcGFyaXNvbiBjYXJkcyAtLS0tLS0tLS0tICovCi5wY2Mtc2VjdGlvbiB7IG1hcmdpbi10b3A6IDI0cHg7IH0KLnBjYy1jb250cm9scyB7IGRpc3BsYXk6IGZsZXg7IGZsZXgtd3JhcDogd3JhcDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiAxNnB4OyBtYXJnaW4tYm90dG9tOiAxNnB4OyB9Ci5wbGF0Zm9ybS1jb21wYXJlLWdyaWQgewogIGRpc3BsYXk6IGdyaWQ7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogcmVwZWF0KDIsIDFmcik7IGdhcDogMTZweDsKfQpAbWVkaWEgKG1heC13aWR0aDogOTAwcHgpIHsgLnBsYXRmb3JtLWNvbXBhcmUtZ3JpZCB7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogMWZyOyB9IH0KLnBsYXRmb3JtLWNvbXBhcmUtY2FyZCB7CiAgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLWxnKTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgcGFkZGluZzogMThweDsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWNhcmQpOwogIHRyYW5zaXRpb246IGJveC1zaGFkb3cgMjIwbXMgdmFyKC0tZWFzZSksIHRyYW5zZm9ybSAyMjBtcyB2YXIoLS1lYXNlKTsKICBhbmltYXRpb246IGNhcmRJbiAzMjBtcyB2YXIoLS1lYXNlKSBiYWNrd2FyZHM7Cn0KLnBsYXRmb3JtLWNvbXBhcmUtY2FyZDpob3ZlciB7IGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdy1ob3Zlcik7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgtMnB4KTsgfQoucGNjLWhlYWRlciB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogc3BhY2UtYmV0d2VlbjsgZ2FwOiAxMHB4OyB9Ci5wY2MtaGVhZGVyLW5hbWUgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDhweDsgfQoucGNjLW5hbWUgeyBmb250LXNpemU6IDE1cHg7IGZvbnQtd2VpZ2h0OiA3MDA7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyB9Ci5wY2MtYmFkZ2UgeyBmb250LXNpemU6IDEzcHg7IGZvbnQtd2VpZ2h0OiA3MDA7IHBhZGRpbmc6IDRweCAxMHB4OyBib3JkZXItcmFkaXVzOiAyMHB4OyB9Ci5wY2MtYmFkZ2UudXAgeyBjb2xvcjogdmFyKC0tc3VjY2Vzcy10ZXh0KTsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXN0YXR1cy1nb29kKSAxNCUsIHRyYW5zcGFyZW50KTsgfQoucGNjLWJhZGdlLmRvd24geyBjb2xvcjogdmFyKC0tc3RhdHVzLWNyaXRpY2FsKTsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXN0YXR1cy1jcml0aWNhbCkgMTIlLCB0cmFuc3BhcmVudCk7IH0KLnBjYy1iYWRnZS5mbGF0IHsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tdGV4dC1tdXRlZCkgMTIlLCB0cmFuc3BhcmVudCk7IH0KLnBjYy1jYXB0aW9uIHsgZm9udC1zaXplOiAxMnB4OyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyBtYXJnaW4tdG9wOiA2cHg7IH0KLnBjYy1tZXRyaWNzIHsgbWFyZ2luLXRvcDogMTZweDsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiAxNHB4OyB9Ci5wY2MtbWV0cmljLXJvdyB7IHBhZGRpbmctdG9wOiAxMnB4OyBib3JkZXItdG9wOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgfQoucGNjLW1ldHJpYy1yb3c6Zmlyc3QtY2hpbGQgeyBwYWRkaW5nLXRvcDogMDsgYm9yZGVyLXRvcDogbm9uZTsgfQoucGNjLW1ldHJpYy1oZWFkZXIgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6IHNwYWNlLWJldHdlZW47IGdhcDogOHB4OyBtYXJnaW4tYm90dG9tOiA2cHg7IH0KLnBjYy1tZXRyaWMtbGFiZWwgeyBmb250LXNpemU6IDEycHg7IGZvbnQtd2VpZ2h0OiA3MDA7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyB9Ci5wY2MtbWV0cmljLWRpZmYgeyBmb250LXNpemU6IDEycHg7IGZvbnQtd2VpZ2h0OiA3MDA7IH0KLnBjYy1tZXRyaWMtZGlmZi51cCB7IGNvbG9yOiB2YXIoLS1zdWNjZXNzLXRleHQpOyB9Ci5wY2MtbWV0cmljLWRpZmYuZG93biB7IGNvbG9yOiB2YXIoLS1zdGF0dXMtY3JpdGljYWwpOyB9Ci5wY2MtbWV0cmljLWRpZmYuZmxhdCB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgfQoucGNjLWZvb3RlciB7IG1hcmdpbi10b3A6IDE2cHg7IHBhZGRpbmctdG9wOiAxNHB4OyBib3JkZXItdG9wOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgfQoucGNjLWZvb3Rlci1sYWJlbCB7IGZvbnQtc2l6ZTogMTFweDsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyBmb250LXdlaWdodDogNjAwOyB0ZXh0LXRyYW5zZm9ybTogdXBwZXJjYXNlOyBsZXR0ZXItc3BhY2luZzogMC4wM2VtOyB9Ci5wY2MtZm9vdGVyLXZhbHVlIHsgZm9udC1zaXplOiAxNXB4OyBmb250LXdlaWdodDogNzAwOyBtYXJnaW4tdG9wOiA0cHg7IH0KLnBjYy1mb290ZXItdmFsdWUudXAgeyBjb2xvcjogdmFyKC0tc3VjY2Vzcy10ZXh0KTsgfQoucGNjLWZvb3Rlci12YWx1ZS5kb3duIHsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IH0KLnBjYy1mb290ZXItdmFsdWUuZmxhdCB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgfQoucGNjLWZvb3Rlci1kZXRhaWwgeyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IG1hcmdpbi10b3A6IDRweDsgfQoucGNjLXZpZXctbGluayB7CiAgZGlzcGxheTogaW5saW5lLWZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogNHB4OyBtYXJnaW4tdG9wOiAxNHB4OwogIGJhY2tncm91bmQ6IG5vbmU7IGJvcmRlcjogbm9uZTsgY29sb3I6IHZhcigtLXNlcmllcy0xKTsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNzAwOyBjdXJzb3I6IHBvaW50ZXI7IHBhZGRpbmc6IDA7CiAgdHJhbnNpdGlvbjogb3BhY2l0eSAxNTBtcyB2YXIoLS1lYXNlKTsKfQoucGNjLXZpZXctbGluazpob3ZlciB7IG9wYWNpdHk6IDAuNzU7IHRleHQtZGVjb3JhdGlvbjogdW5kZXJsaW5lOyB9CgovKiAtLS0tLS0tLS0tIFBhZ2luYXRpb24gLS0tLS0tLS0tLSAqLwoucGFnaW5hdGlvbi1yb3cgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDEycHg7IG1hcmdpbi10b3A6IDE0cHg7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgfQoucGFnaW5hdGlvbi1yb3cgLmJ0biB7IHBhZGRpbmc6IDZweCAxMnB4OyB9Ci5leHBvcnQtYnV0dG9ucyB7IGRpc3BsYXk6IGZsZXg7IGdhcDogOHB4OyBmbGV4LXdyYXA6IHdyYXA7IG1hcmdpbi1ib3R0b206IDEycHg7IH0KLmV4cG9ydC1idXR0b25zIC5idG4geyBwYWRkaW5nOiA3cHggMTNweDsgZm9udC1zaXplOiAxMnB4OyB9CgovKiAtLS0tLS0tLS0tIERhc2hib2FyZCBjb250cm9scyAvIG1ldHJpYy1mb2N1c2VkIEtQSXMgLS0tLS0tLS0tLSAqLwouZGFzaGJvYXJkLWNvbnRyb2xzIHsKICBkaXNwbGF5OiBmbGV4OyBmbGV4LXdyYXA6IHdyYXA7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogMTBweDsgbWFyZ2luLWJvdHRvbTogMThweDsKfQouZGFzaGJvYXJkLWNvbnRyb2xzIGxhYmVsIHsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNjAwOyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyBtYXJnaW4tcmlnaHQ6IDZweDsgfQouZGFzaGJvYXJkLWNvbnRyb2xzIHNlbGVjdCB7IGZvbnQtd2VpZ2h0OiA2MDA7IH0KLyogQmVzdCBQZXJmb3JtaW5nIFBvc3Qg4oCUIGEgZmVhdHVyZWQgbGFuZHNjYXBlIGNhcmQgc3Bhbm5pbmcgMyBLUEktdGlsZS13aWR0aHMKICAgKGEgc3RhbmRhcmQgdGlsZSBpcyAxIHVuaXQ7ICNrcGlHcmlkIGhhcyAxMCB1bml0cyB0b3RhbCksIHNhbWUgZml4ZWQKICAgaGVpZ2h0IGFzIHRoZSByZXN0IG9mICNrcGlHcmlkOiBjYXB0aW9uL3BsYXRmb3JtL2RhdGUgc2l0IG9uIHRoZSBsZWZ0LAogICB3aXRoIHRoZSBzZWxlY3RlZCBtZXRyaWMgKGxhcmdlKSBhbmQgQ3VycmVudCBGb2xsb3dlcnMgKHNtYWxsZXIsIGJlbG93IGEKICAgZGl2aWRlcikgc3RhY2tlZCBpbiBhIG5hcnJvd2VyIGNvbHVtbiBvbiB0aGUgcmlnaHQuICovCiNrcGlHcmlkIC5wb3N0LXRpbGUgewogIGdyaWQtY29sdW1uOiBzcGFuIDM7CiAgZmxleC1kaXJlY3Rpb246IHJvdzsKICBhbGlnbi1pdGVtczogc3RyZXRjaDsKICBqdXN0aWZ5LWNvbnRlbnQ6IGZsZXgtc3RhcnQ7CiAgZ2FwOiAyMHB4OwogIHBhZGRpbmc6IDE0cHggMjBweDsKfQoja3BpR3JpZCAucG9zdC10aWxlLW1haW4geyBmbGV4OiAxIDEgYXV0bzsgbWluLXdpZHRoOiAwOyBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBqdXN0aWZ5LWNvbnRlbnQ6IGNlbnRlcjsgZ2FwOiA2cHg7IH0KI2twaUdyaWQgLnBvc3QtdGlsZS1jYXB0aW9uIHsKICBmb250LXNpemU6IDEzcHg7IGZvbnQtd2VpZ2h0OiA2MDA7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOwogIGxpbmUtaGVpZ2h0OiAxLjQ7CiAgZGlzcGxheTogLXdlYmtpdC1ib3g7IC13ZWJraXQtbGluZS1jbGFtcDogMzsgLXdlYmtpdC1ib3gtb3JpZW50OiB2ZXJ0aWNhbDsgb3ZlcmZsb3c6IGhpZGRlbjsKfQoja3BpR3JpZCAucG9zdC10aWxlLWNhcHRpb24ubXV0ZWQgeyBjb2xvcjogdmFyKC0tdGV4dC1tdXRlZCk7IGZvbnQtd2VpZ2h0OiA1MDA7IC13ZWJraXQtbGluZS1jbGFtcDogMTsgfQoja3BpR3JpZCAucG9zdC10aWxlLW1ldGEgeyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA2cHg7IHdoaXRlLXNwYWNlOiBub3dyYXA7IG92ZXJmbG93OiBoaWRkZW47IHRleHQtb3ZlcmZsb3c6IGVsbGlwc2lzOyB9CiNrcGlHcmlkIC5wb3N0LXRpbGUtZGl2aWRlciB7IGZsZXg6IDAgMCBhdXRvOyB3aWR0aDogMXB4OyBhbGlnbi1zZWxmOiBzdHJldGNoOyBiYWNrZ3JvdW5kOiB2YXIoLS1ib3JkZXIpOyB9CiNrcGlHcmlkIC5wb3N0LXRpbGUtbWV0cmljcyB7CiAgZmxleDogMCAwIGF1dG87IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGp1c3RpZnktY29udGVudDogY2VudGVyOyBnYXA6IDhweDsgbWluLXdpZHRoOiAxMjBweDsKfQoja3BpR3JpZCAucG9zdC10aWxlLW1ldHJpYy1ibG9jayB7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGdhcDogM3B4OyB9CiNrcGlHcmlkIC5wb3N0LXRpbGUtbWV0cmljLWJsb2NrLnNlY29uZGFyeSB7IHBhZGRpbmctdG9wOiA4cHg7IGJvcmRlci10b3A6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyB9CiNrcGlHcmlkIC5wb3N0LXRpbGUtbWV0cmljLWxhYmVsIHsgZm9udC1zaXplOiAxMXB4OyBjb2xvcjogdmFyKC0tdGV4dC1tdXRlZCk7IGZvbnQtd2VpZ2h0OiA2MDA7IHRleHQtdHJhbnNmb3JtOiB1cHBlcmNhc2U7IGxldHRlci1zcGFjaW5nOiAwLjAzZW07IH0KI2twaUdyaWQgLnBvc3QtdGlsZS1tZXRyaWMtdmFsdWUgeyBmb250LXNpemU6IDIycHg7IGZvbnQtd2VpZ2h0OiA3MDA7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyBmb250LXZhcmlhbnQtbnVtZXJpYzogdGFidWxhci1udW1zOyB9CiNrcGlHcmlkIC5wb3N0LXRpbGUtbWV0cmljLWJsb2NrLnNlY29uZGFyeSAucG9zdC10aWxlLW1ldHJpYy12YWx1ZSB7IGZvbnQtc2l6ZTogMTVweDsgfQpAbWVkaWEgKG1heC13aWR0aDogNjQwcHgpIHsgI2twaUdyaWQgLnBvc3QtdGlsZSB7IGdyaWQtY29sdW1uOiBzcGFuIDI7IH0gfQoKLnN0YXQtdmFsdWUtbXV0ZWQgeyBmb250LXNpemU6IDE1cHggIWltcG9ydGFudDsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyBmb250LXdlaWdodDogNjAwOyB9Ci5jYXB0aW9uLWxpbmsgeyBjb2xvcjogdmFyKC0tc2VyaWVzLTEpOyB0ZXh0LWRlY29yYXRpb246IG5vbmU7IH0KLmNhcHRpb24tbGluazpob3ZlciB7IHRleHQtZGVjb3JhdGlvbjogdW5kZXJsaW5lOyB9CgovKiAtLS0tLS0tLS0tIERhdGEgUmVjb3JkcyAocGxhdGZvcm0tZ3JvdXBlZCkgLS0tLS0tLS0tLSAqLwoucmVjb3Jkcy10b29sYmFyIHsKICBkaXNwbGF5OiBmbGV4OyBmbGV4LXdyYXA6IHdyYXA7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogc3BhY2UtYmV0d2VlbjsKICBnYXA6IDEycHg7IG1hcmdpbi1ib3R0b206IDE0cHg7Cn0KLnBsYXRmb3JtLWZpbHRlci1waWxscyB7IGRpc3BsYXk6IGZsZXg7IGZsZXgtd3JhcDogd3JhcDsgZ2FwOiA2cHg7IH0KLnBsYXRmb3JtLWZpbHRlci1waWxscyBidXR0b24gewogIGRpc3BsYXk6IGlubGluZS1mbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDZweDsKICBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTEpOyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOwogIGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBwYWRkaW5nOiA3cHggMTRweDsgYm9yZGVyLXJhZGl1czogMjBweDsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNjAwOyBjdXJzb3I6IHBvaW50ZXI7CiAgdHJhbnNpdGlvbjogY29sb3IgMTgwbXMgdmFyKC0tZWFzZSksIGJhY2tncm91bmQgMTgwbXMgdmFyKC0tZWFzZSksIHRyYW5zZm9ybSAxNTBtcyB2YXIoLS1lYXNlKSwgYm94LXNoYWRvdyAxODBtcyB2YXIoLS1lYXNlKTsKfQoucGxhdGZvcm0tZmlsdGVyLXBpbGxzIGJ1dHRvbjpob3ZlciB7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoLTFweCk7IH0KLnBsYXRmb3JtLWZpbHRlci1waWxscyBidXR0b246YWN0aXZlIHsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDApIHNjYWxlKDAuOTYpOyB9Ci5wbGF0Zm9ybS1maWx0ZXItcGlsbHMgYnV0dG9uLmlzLWFjdGl2ZSB7IGJhY2tncm91bmQ6IHZhcigtLXNlcmllcy0xKTsgY29sb3I6ICNmZmY7IGJvcmRlci1jb2xvcjogdHJhbnNwYXJlbnQ7IGJveC1zaGFkb3c6IDAgNHB4IDE0cHggLTVweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDYwJSwgdHJhbnNwYXJlbnQpOyB9Ci5wbGF0Zm9ybS1maWx0ZXItcGlsbHMgYnV0dG9uLmlzLWFjdGl2ZSAucGxhdGZvcm0tZG90IHsgYm94LXNoYWRvdzogMCAwIDAgMnB4IHJnYmEoMjU1LDI1NSwyNTUsMC41KTsgfQoucmVjb3Jkcy1zZWFyY2ggaW5wdXQgeyBib3JkZXItcmFkaXVzOiAyMHB4OyBtaW4td2lkdGg6IDIyMHB4OyB9Ci5zdGF0dXMtcGlsbCB7IGRpc3BsYXk6IGlubGluZS1ibG9jazsgcGFkZGluZzogM3B4IDEwcHg7IGJvcmRlci1yYWRpdXM6IDIwcHg7IGZvbnQtc2l6ZTogMTFweDsgZm9udC13ZWlnaHQ6IDcwMDsgfQouc3RhdHVzLXBpbGwub3JpZ2luYWwgeyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tdGV4dC1tdXRlZCkgMTUlLCB0cmFuc3BhcmVudCk7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IH0KLnN0YXR1cy1waWxsLmVkaXRlZCB7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zdGF0dXMtd2FybmluZykgMjIlLCB0cmFuc3BhcmVudCk7IGNvbG9yOiAjOGE2MzAwOyB9Ci5yb3ctYWN0aW9ucyB7IGRpc3BsYXk6IGZsZXg7IGdhcDogNnB4OyBmbGV4LXdyYXA6IG5vd3JhcDsgfQoucm93LWFjdGlvbnMgLmJ0biB7IHBhZGRpbmc6IDVweCAxMHB4OyBmb250LXNpemU6IDEycHg7IH0KLmxpbmstY2VsbCBhIHsgY29sb3I6IHZhcigtLXNlcmllcy0xKTsgdGV4dC1kZWNvcmF0aW9uOiBub25lOyBmb250LXdlaWdodDogNjAwOyBmb250LXNpemU6IDEycHg7IH0KLmxpbmstY2VsbCBhOmhvdmVyIHsgdGV4dC1kZWNvcmF0aW9uOiB1bmRlcmxpbmU7IH0KLnJlY29yZC1zZWN0aW9uIHsKICBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtc20pOyBwYWRkaW5nOiAxNnB4OyBtYXJnaW4tYm90dG9tOiAxNHB4OwogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMik7IGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKfQoucmVjb3JkLXNlY3Rpb24gaDQgeyBtYXJnaW46IDAgMCAxMnB4OyBmb250LXNpemU6IDEycHg7IGZvbnQtd2VpZ2h0OiA3MDA7IGxldHRlci1zcGFjaW5nOiAwLjAzZW07IHRleHQtdHJhbnNmb3JtOiB1cHBlcmNhc2U7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IH0KLnJlY29yZC1zZWN0aW9uIC5mb3JtLWdyaWQgeyBtYXJnaW4tYm90dG9tOiAwOyB9Ci5yZWNvcmQtc2VjdGlvbiAudmlldy1maWVsZCB7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGdhcDogMnB4OyBmb250LXNpemU6IDEzcHg7IH0KLnJlY29yZC1zZWN0aW9uIC52aWV3LWZpZWxkIC52aWV3LWxhYmVsIHsgZm9udC1zaXplOiAxMXB4OyBmb250LXdlaWdodDogNjAwOyBjb2xvcjogdmFyKC0tdGV4dC1tdXRlZCk7IH0KLnJlY29yZC1zZWN0aW9uIC52aWV3LWZpZWxkIC52aWV3LXZhbHVlIHsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IHdvcmQtYnJlYWs6IGJyZWFrLXdvcmQ7IH0KQG1lZGlhIChtYXgtd2lkdGg6IDY0MHB4KSB7CiAgLnJlY29yZHMtdG9vbGJhciB7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGFsaWduLWl0ZW1zOiBzdHJldGNoOyB9CiAgLnJlY29yZHMtc2VhcmNoIGlucHV0IHsgd2lkdGg6IDEwMCU7IH0KfQoKLyogLS0tLS0tLS0tLSBNb2RhbCAocmVjb3JkIGVkaXRvcikgLS0tLS0tLS0tLSAqLwoubW9kYWwtb3ZlcmxheSB7CiAgcG9zaXRpb246IGZpeGVkOyBpbnNldDogMDsgYmFja2dyb3VuZDogcmdiYSgxMCwxMSwxMywwLjUpOwogIGJhY2tkcm9wLWZpbHRlcjogYmx1cig2cHgpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogYmx1cig2cHgpOwogIGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBmbGV4LXN0YXJ0OyBqdXN0aWZ5LWNvbnRlbnQ6IGNlbnRlcjsKICBwYWRkaW5nOiA0MHB4IDE2cHg7IG92ZXJmbG93LXk6IGF1dG87IHotaW5kZXg6IDIwMDsKICBhbmltYXRpb246IG92ZXJsYXlJbiAyMDBtcyB2YXIoLS1lYXNlKTsKfQpAa2V5ZnJhbWVzIG92ZXJsYXlJbiB7IGZyb20geyBvcGFjaXR5OiAwOyB9IHRvIHsgb3BhY2l0eTogMTsgfSB9CkBrZXlmcmFtZXMgbW9kYWxQYW5lbEluIHsKICBmcm9tIHsgb3BhY2l0eTogMDsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDE0cHgpIHNjYWxlKDAuOTcpOyB9CiAgdG8geyBvcGFjaXR5OiAxOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMCkgc2NhbGUoMSk7IH0KfQoubW9kYWwtcGFuZWwgewogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1sZyk7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHBhZGRpbmc6IDI0cHg7IHdpZHRoOiAxMDAlOyBtYXgtd2lkdGg6IDcyMHB4OyBib3gtc2hhZG93OiB2YXIoLS1zaGFkb3ctbW9kYWwpOwogIG1heC1oZWlnaHQ6IGNhbGMoMTAwdmggLSA4MHB4KTsgb3ZlcmZsb3cteTogYXV0bzsKICBhbmltYXRpb246IG1vZGFsUGFuZWxJbiAyNDBtcyB2YXIoLS1lYXNlKTsKfQoubW9kYWwtcGFuZWwud2lkZSB7IG1heC13aWR0aDogMTEwMHB4OyB9Ci5tb2RhbC1wYW5lbCBoMiB7IG1hcmdpbjogMCAwIDRweDsgZm9udC1zaXplOiAxN3B4OyBsZXR0ZXItc3BhY2luZzogLTAuMDFlbTsgfQoubW9kYWwtcGFuZWwgLm1vZGFsLXN1YiB7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IGZvbnQtc2l6ZTogMTJweDsgbWFyZ2luOiAwIDAgMThweDsgfQouZm9ybS1ncmlkIHsgZGlzcGxheTogZ3JpZDsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiByZXBlYXQoYXV0by1maXQsIG1pbm1heCgyMDBweCwgMWZyKSk7IGdhcDogMTJweDsgbWFyZ2luLWJvdHRvbTogMTZweDsgfQouZm9ybS1ncmlkLmZ1bGwgeyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IDFmcjsgfQpAbWVkaWEgKG1heC13aWR0aDogNjQwcHgpIHsgLmZvcm0tZ3JpZCB7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogMWZyOyB9IH0KLmZvcm0tZmllbGQgeyBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBnYXA6IDVweDsgZm9udC1zaXplOiAxMnB4OyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyB9Ci5mb3JtLWZpZWxkIGxhYmVsIHsgZm9udC13ZWlnaHQ6IDYwMDsgfQouZm9ybS1maWVsZCB0ZXh0YXJlYSB7IHJlc2l6ZTogdmVydGljYWw7IG1pbi1oZWlnaHQ6IDYwcHg7IH0KCi5wbGF0Zm9ybS1lZGl0LXJvdyB7CiAgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLXNtKTsgcGFkZGluZzogMTRweDsgbWFyZ2luLWJvdHRvbTogMTBweDsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0yKTsKfQoucGxhdGZvcm0tZWRpdC1yb3cgLnBsYXRmb3JtLWVkaXQtaGVhZCB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogc3BhY2UtYmV0d2VlbjsgZ2FwOiA4cHg7IG1hcmdpbi1ib3R0b206IDEwcHg7IH0KLnBsYXRmb3JtLWVkaXQtcm93IC5tZXRyaWNzLWdyaWQgeyBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdChhdXRvLWZpdCwgbWlubWF4KDEyMHB4LCAxZnIpKTsgZ2FwOiA4cHg7IH0KLnJlbW92ZS1wbGF0Zm9ybS1idG4geyBib3JkZXI6IG5vbmU7IGJhY2tncm91bmQ6IHRyYW5zcGFyZW50OyBjb2xvcjogdmFyKC0tc3RhdHVzLWNyaXRpY2FsKTsgY3Vyc29yOiBwb2ludGVyOyBmb250LXNpemU6IDEycHg7IGZvbnQtd2VpZ2h0OiA2MDA7IHRyYW5zaXRpb246IG9wYWNpdHkgMTUwbXMgdmFyKC0tZWFzZSk7IH0KLnJlbW92ZS1wbGF0Zm9ybS1idG46aG92ZXIgeyBvcGFjaXR5OiAwLjc7IH0KLm1vZGFsLWFjdGlvbnMgeyBkaXNwbGF5OiBmbGV4OyBqdXN0aWZ5LWNvbnRlbnQ6IHNwYWNlLWJldHdlZW47IGFsaWduLWl0ZW1zOiBjZW50ZXI7IG1hcmdpbi10b3A6IDE4cHg7IGdhcDogOHB4OyBmbGV4LXdyYXA6IHdyYXA7IH0KCi8qIC0tLS0tLS0tLS0gUmVzcG9uc2l2ZSB0aWdodGVuaW5nIC0tLS0tLS0tLS0gKi8KQG1lZGlhIChtYXgtd2lkdGg6IDcyMHB4KSB7CiAgLmFwcC1zaGVsbCB7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IH0KICAuc2lkZWJhciB7IHdpZHRoOiAxMDAlOyBoZWlnaHQ6IGF1dG87IHBvc2l0aW9uOiBzdGF0aWM7IGZsZXgtZGlyZWN0aW9uOiByb3c7IGZsZXgtd3JhcDogd3JhcDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA4cHg7IHBhZGRpbmc6IDEwcHggMTRweDsgfQogIC5zaWRlYmFyLWJyYW5kIHsgcGFkZGluZzogMDsgbWFyZ2luLXJpZ2h0OiBhdXRvOyB9CiAgLnRhYnMgeyBmbGV4LWRpcmVjdGlvbjogcm93OyB3aWR0aDogMTAwJTsgb3ZlcmZsb3cteDogYXV0bzsgb3JkZXI6IDM7IH0KICAuc2lkZWJhci1mb290ZXIgeyBmbGV4LWRpcmVjdGlvbjogcm93OyBib3JkZXItdG9wOiBub25lOyBtYXJnaW4tdG9wOiAwOyBwYWRkaW5nLXRvcDogMDsgfQogIC52aWV3LWFyZWEgeyBwYWRkaW5nOiAxNHB4OyB9CiAgLmZpbHRlci1iYXIgeyB0b3A6IGF1dG87IHBvc2l0aW9uOiBzdGF0aWM7IHBhZGRpbmc6IDEycHggMTRweDsgfQogIC5zdGF0LWdyaWQgeyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdChhdXRvLWZpdCwgbWlubWF4KDE0MHB4LCAxZnIpKTsgfQogIC5icmFuZC1sb2dvIHsgaGVpZ2h0OiAyMnB4OyB9Cn0KCi8qIC0tLS0tLS0tLS0gUmVwb3J0IEdlbmVyYXRvciAtLS0tLS0tLS0tICovCi5yZXBvcnQtY29udHJvbHMgeyBkaXNwbGF5OiBmbGV4OyBmbGV4LXdyYXA6IHdyYXA7IGdhcDogMTRweCAxOHB4OyBhbGlnbi1pdGVtczogZW5kOyB9Ci5yZXBvcnQtYWN0aW9ucyB7IGRpc3BsYXk6IGZsZXg7IGdhcDogOHB4OyBmbGV4LXdyYXA6IHdyYXA7IG1hcmdpbi10b3A6IDRweDsgfQoucmVwb3J0LWRvYy1oZWFkIHsgbWFyZ2luLWJvdHRvbTogNHB4OyB9Ci5yZXBvcnQtZG9jLWhlYWQgLnJlcG9ydC10aXRsZSB7IGZvbnQtc2l6ZTogMjBweDsgZm9udC13ZWlnaHQ6IDcwMDsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IGxldHRlci1zcGFjaW5nOiAtMC4wMmVtOyB9Ci5yZXBvcnQtZG9jLWhlYWQgLnJlcG9ydC1yYW5nZSB7IGZvbnQtc2l6ZTogMTNweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgbWFyZ2luLXRvcDogNHB4OyB9Ci5yZXBvcnQtZG9jLWhlYWQgLnJlcG9ydC1nZW5lcmF0ZWQgeyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgbWFyZ2luLXRvcDogMnB4OyB9Ci5yZXBvcnQtbWV0cmljLXNlbGVjdCB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogOHB4OyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IG1hcmdpbi1ib3R0b206IDEycHg7IH0KLnJlcG9ydC1zdW1tYXJ5IHsgbGluZS1oZWlnaHQ6IDEuNjsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgZm9udC1zaXplOiAxMy41cHg7IH0KLnJlcG9ydC1zdW1tYXJ5IGg0IHsgZm9udC1zaXplOiAxM3B4OyBmb250LXdlaWdodDogNzAwOyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgdGV4dC10cmFuc2Zvcm06IHVwcGVyY2FzZTsgbGV0dGVyLXNwYWNpbmc6IDAuMDRlbTsgbWFyZ2luOiAxOHB4IDAgOHB4OyB9Ci5yZXBvcnQtc3VtbWFyeSBoNDpmaXJzdC1jaGlsZCB7IG1hcmdpbi10b3A6IDA7IH0KLnJlcG9ydC1zdW1tYXJ5IHAgeyBtYXJnaW46IDAgMCA2cHg7IH0KLnJlcG9ydC1zdW1tYXJ5IHVsIHsgbWFyZ2luOiAwOyBwYWRkaW5nLWxlZnQ6IDE4cHg7IH0KLnJlcG9ydC1zdW1tYXJ5IGxpIHsgbWFyZ2luOiA0cHggMDsgfQoucmVwb3J0LXN1bW1hcnkgLnRyZW5kLXVwIHsgY29sb3I6IHZhcigtLXN1Y2Nlc3MtdGV4dCk7IGZvbnQtd2VpZ2h0OiA2MDA7IH0KLnJlcG9ydC1zdW1tYXJ5IC50cmVuZC1kb3duIHsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IGZvbnQtd2VpZ2h0OiA2MDA7IH0KLmRhdGEtdGFibGUgdGQudHJlbmQtdXAgeyBjb2xvcjogdmFyKC0tc3VjY2Vzcy10ZXh0KTsgZm9udC13ZWlnaHQ6IDYwMDsgfQouZGF0YS10YWJsZSB0ZC50cmVuZC1kb3duIHsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IGZvbnQtd2VpZ2h0OiA2MDA7IH0KLmRhdGEtdGFibGUgdGQudHJlbmQtZmxhdCB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgfQoKLyogLS0tLS0tLS0tLSBQcmludCAvIEV4cG9ydCAoUmVwb3J0IEdlbmVyYXRvciAiUHJpbnQgLyBFeHBvcnQgUmVwb3J0IikgLS0tLS0tLS0tLSAqLwpAbWVkaWEgcHJpbnQgewogIC5zaWRlYmFyLCAjZmlsdGVyQmFyLCAjdG9hc3RSb290LCAjdmlldy1yZXBvcnQgPiAuY2FyZCwgLnJlcG9ydC1tZXRyaWMtc2VsZWN0IHsgZGlzcGxheTogbm9uZSAhaW1wb3J0YW50OyB9CiAgLmFwcC1zaGVsbCwgLm1haW4tY29sLCAudmlldy1hcmVhIHsgZGlzcGxheTogYmxvY2sgIWltcG9ydGFudDsgYmFja2dyb3VuZDogI2ZmZiAhaW1wb3J0YW50OyB9CiAgLnZpZXctYXJlYSB7IHBhZGRpbmc6IDAgIWltcG9ydGFudDsgfQogIC52aWV3Om5vdCguaXMtYWN0aXZlKSB7IGRpc3BsYXk6IG5vbmUgIWltcG9ydGFudDsgfQogIC52aWV3LmlzLWFjdGl2ZSB7IGRpc3BsYXk6IGJsb2NrICFpbXBvcnRhbnQ7IH0KICBib2R5IHsgYmFja2dyb3VuZDogI2ZmZiAhaW1wb3J0YW50OyBjb2xvcjogIzExMSAhaW1wb3J0YW50OyB9CiAgLmNhcmQgeyBiYWNrZ3JvdW5kOiAjZmZmICFpbXBvcnRhbnQ7IGJvcmRlcjogMXB4IHNvbGlkICNkZGQgIWltcG9ydGFudDsgYm94LXNoYWRvdzogbm9uZSAhaW1wb3J0YW50OyBiYWNrZHJvcC1maWx0ZXI6IG5vbmUgIWltcG9ydGFudDsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IG5vbmUgIWltcG9ydGFudDsgYnJlYWstaW5zaWRlOiBhdm9pZDsgfQogIC5zZWN0aW9uLXRpdGxlLCAuc3RhdC12YWx1ZSwgLnJlcG9ydC1kb2MtaGVhZCAucmVwb3J0LXRpdGxlLCAucmVwb3J0LXN1bW1hcnkgaDQgeyBjb2xvcjogIzExMSAhaW1wb3J0YW50OyB9CiAgLnN0YXQtdGlsZSB7IGJhY2tncm91bmQ6ICNmZmYgIWltcG9ydGFudDsgYm9yZGVyOiAxcHggc29saWQgI2U2ZTZlNiAhaW1wb3J0YW50OyBib3gtc2hhZG93OiBub25lICFpbXBvcnRhbnQ7IH0KICAuc3RhdC1sYWJlbCwgLnN0YXQtZGVsdGEsIC5yZXBvcnQtZG9jLWhlYWQgLnJlcG9ydC1yYW5nZSwgLnJlcG9ydC1kb2MtaGVhZCAucmVwb3J0LWdlbmVyYXRlZCwgLnJlcG9ydC1zdW1tYXJ5IHsgY29sb3I6ICM0NDQgIWltcG9ydGFudDsgfQogIC5zdGF0LWRlbHRhLnVwLCAuZGF0YS10YWJsZSB0ZC50cmVuZC11cCB7IGNvbG9yOiAjMWE3ZjM3ICFpbXBvcnRhbnQ7IH0KICAuc3RhdC1kZWx0YS5kb3duLCAuZGF0YS10YWJsZSB0ZC50cmVuZC1kb3duIHsgY29sb3I6ICNiMzI2MWUgIWltcG9ydGFudDsgfQogIC5kYXRhLXRhYmxlIHRoLCAuZGF0YS10YWJsZSB0ZCB7IGNvbG9yOiAjMTExICFpbXBvcnRhbnQ7IGJvcmRlci1jb2xvcjogI2RkZCAhaW1wb3J0YW50OyB9CiAgLmRhdGEtdGFibGUgdGhlYWQgdGggeyBjb2xvcjogIzQ0NCAhaW1wb3J0YW50OyB9CiAgYSwgLmNhcHRpb24tbGluayB7IGNvbG9yOiAjMWE0ZjhiICFpbXBvcnRhbnQ7IH0KICAuY2hhcnQtd3JhcCB7IGhlaWdodDogMzQwcHggIWltcG9ydGFudDsgcGFnZS1icmVhay1pbnNpZGU6IGF2b2lkOyB9CiAgQHBhZ2UgeyBtYXJnaW46IDE0bW07IH0KfQo8L3N0eWxlPgo8L2hlYWQ+Cjxib2R5Pgo8ZGl2IGNsYXNzPSJhdXRoLXNjcmVlbiIgaWQ9ImF1dGhTY3JlZW4iPgogIDxkaXYgY2xhc3M9ImF1dGgtY2FyZCI+CiAgICA8ZGl2IGNsYXNzPSJhdXRoLWJyYW5kIj4KICAgICAgPGltZyBjbGFzcz0iYnJhbmQtbG9nbyIgYWx0PSJMaWdvbi1SYXpvbiBTb2x1dGlvbnMgbG9nbyIgLz4KICAgICAgPHNwYW4gY2xhc3M9ImJyYW5kLXRpdGxlIj5Tb2NpYWwgTWVkaWEgQW5hbHl0aWNzPC9zcGFuPgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJhdXRoLWZvcm0iPgogICAgICA8ZGl2IGNsYXNzPSJmb3JtLWZpZWxkIj4KICAgICAgICA8bGFiZWwgZm9yPSJhdXRoQ29kZSI+QWNjZXNzIGNvZGU8L2xhYmVsPgogICAgICAgIDxpbnB1dCB0eXBlPSJwYXNzd29yZCIgaWQ9ImF1dGhDb2RlIiBhdXRvY29tcGxldGU9Im9mZiIgLz4KICAgICAgPC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImF1dGgtZXJyb3IiIGlkPSJhdXRoRXJyb3IiPjwvZGl2PgogICAgICA8YnV0dG9uIGNsYXNzPSJidG4gcHJpbWFyeSIgaWQ9ImF1dGhTdWJtaXRCdG4iIHR5cGU9ImJ1dHRvbiI+PGkgZGF0YS1sdWNpZGU9ImFycm93LXJpZ2h0IiBzdHlsZT0id2lkdGg6MTRweDtoZWlnaHQ6MTRweDsiPjwvaT4gRW50ZXI8L2J1dHRvbj4KICAgIDwvZGl2PgogIDwvZGl2Pgo8L2Rpdj4KCjxkaXYgY2xhc3M9ImFwcC1zaGVsbCIgaWQ9ImFwcFNoZWxsIiBzdHlsZT0iZGlzcGxheTpub25lOyI+CiAgPGFzaWRlIGNsYXNzPSJzaWRlYmFyIj4KICAgIDxkaXYgY2xhc3M9InNpZGViYXItYnJhbmQiPgogICAgICA8aW1nIGNsYXNzPSJicmFuZC1sb2dvIiBhbHQ9IkxpZ29uLVJhem9uIFNvbHV0aW9ucyBsb2dvIiAvPgogICAgICA8c3BhbiBjbGFzcz0iYnJhbmQtdGl0bGUiPlNvY2lhbCBNZWRpYSBBbmFseXRpY3M8L3NwYW4+CiAgICA8L2Rpdj4KICAgIDxuYXYgY2xhc3M9InRhYnMiIHJvbGU9InRhYmxpc3QiIGFyaWEtbGFiZWw9IlNlY3Rpb25zIj4KICAgICAgPGJ1dHRvbiBjbGFzcz0idGFiLWJ0biBpcy1hY3RpdmUiIGRhdGEtdGFiPSJkYXNoYm9hcmQiIHJvbGU9InRhYiIgYXJpYS1zZWxlY3RlZD0idHJ1ZSI+PGkgZGF0YS1sdWNpZGU9ImxheW91dC1kYXNoYm9hcmQiIHN0eWxlPSJ3aWR0aDoxNHB4O2hlaWdodDoxNHB4OyI+PC9pPiBEYXNoYm9hcmQ8L2J1dHRvbj4KICAgICAgPGJ1dHRvbiBjbGFzcz0idGFiLWJ0biIgZGF0YS10YWI9InJlY29yZHMiIHJvbGU9InRhYiIgYXJpYS1zZWxlY3RlZD0iZmFsc2UiPjxpIGRhdGEtbHVjaWRlPSJkYXRhYmFzZSIgc3R5bGU9IndpZHRoOjE0cHg7aGVpZ2h0OjE0cHg7Ij48L2k+IERhdGEgUmVjb3JkczwvYnV0dG9uPgogICAgICA8YnV0dG9uIGNsYXNzPSJ0YWItYnRuIiBkYXRhLXRhYj0iZm9sbG93ZXJzIiByb2xlPSJ0YWIiIGFyaWEtc2VsZWN0ZWQ9ImZhbHNlIj48aSBkYXRhLWx1Y2lkZT0idXNlcnMiIHN0eWxlPSJ3aWR0aDoxNHB4O2hlaWdodDoxNHB4OyI+PC9pPiBGb2xsb3dlcnMgRGF0YTwvYnV0dG9uPgogICAgICA8YnV0dG9uIGNsYXNzPSJ0YWItYnRuIiBkYXRhLXRhYj0iY29tcGFyaXNvbiIgcm9sZT0idGFiIiBhcmlhLXNlbGVjdGVkPSJmYWxzZSI+PGkgZGF0YS1sdWNpZGU9ImdpdC1jb21wYXJlIiBzdHlsZT0id2lkdGg6MTRweDtoZWlnaHQ6MTRweDsiPjwvaT4gQ29tcGFyaXNvbnM8L2J1dHRvbj4KICAgICAgPGJ1dHRvbiBjbGFzcz0idGFiLWJ0biIgZGF0YS10YWI9InJlcG9ydCIgcm9sZT0idGFiIiBhcmlhLXNlbGVjdGVkPSJmYWxzZSI+PGkgZGF0YS1sdWNpZGU9ImZpbGUtdGV4dCIgc3R5bGU9IndpZHRoOjE0cHg7aGVpZ2h0OjE0cHg7Ij48L2k+IFJlcG9ydCBHZW5lcmF0b3I8L2J1dHRvbj4KICAgICAgPGJ1dHRvbiBjbGFzcz0idGFiLWJ0biIgZGF0YS10YWI9InVwbG9hZCIgcm9sZT0idGFiIiBhcmlhLXNlbGVjdGVkPSJmYWxzZSI+PGkgZGF0YS1sdWNpZGU9InVwbG9hZC1jbG91ZCIgc3R5bGU9IndpZHRoOjE0cHg7aGVpZ2h0OjE0cHg7Ij48L2k+IFVwbG9hZCBEYXRhPC9idXR0b24+CiAgICAgIDxidXR0b24gY2xhc3M9InRhYi1idG4iIGRhdGEtdGFiPSJoaXN0b3J5IiByb2xlPSJ0YWIiIGFyaWEtc2VsZWN0ZWQ9ImZhbHNlIj48aSBkYXRhLWx1Y2lkZT0iaGlzdG9yeSIgc3R5bGU9IndpZHRoOjE0cHg7aGVpZ2h0OjE0cHg7Ij48L2k+IFVwbG9hZCBIaXN0b3J5PC9idXR0b24+CiAgICA8L25hdj4KICAgIDxkaXYgY2xhc3M9InNpZGViYXItZm9vdGVyIj4KICAgICAgPGJ1dHRvbiBjbGFzcz0iYnRuIiBpZD0ibG9nb3V0QnRuIiB0eXBlPSJidXR0b24iPjxpIGRhdGEtbHVjaWRlPSJsb2NrIiBzdHlsZT0id2lkdGg6MTRweDtoZWlnaHQ6MTRweDsiPjwvaT4gTG9jazwvYnV0dG9uPgogICAgPC9kaXY+CiAgPC9hc2lkZT4KCiAgPGRpdiBjbGFzcz0ibWFpbi1jb2wiPgogIDxzZWN0aW9uIGNsYXNzPSJmaWx0ZXItYmFyIiBpZD0iZmlsdGVyQmFyIj4KICAgIDxkaXYgY2xhc3M9ImZpbHRlci1maWVsZCI+CiAgICAgIDxsYWJlbCBmb3I9ImZpbHRlckRhdGVGcm9tIj5Gcm9tPC9sYWJlbD4KICAgICAgPGlucHV0IHR5cGU9ImRhdGUiIGlkPSJmaWx0ZXJEYXRlRnJvbSIgLz4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0iZmlsdGVyLWZpZWxkIj4KICAgICAgPGxhYmVsIGZvcj0iZmlsdGVyRGF0ZVRvIj5UbzwvbGFiZWw+CiAgICAgIDxpbnB1dCB0eXBlPSJkYXRlIiBpZD0iZmlsdGVyRGF0ZVRvIiAvPgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJmaWx0ZXItZmllbGQgZmlsdGVyLXByZXNldHMiIGlkPSJmaWx0ZXJQcmVzZXRzIj4KICAgICAgPGJ1dHRvbiB0eXBlPSJidXR0b24iIGRhdGEtcHJlc2V0PSI3Ij5MYXN0IDcgZGF5czwvYnV0dG9uPgogICAgICA8YnV0dG9uIHR5cGU9ImJ1dHRvbiIgZGF0YS1wcmVzZXQ9IjMwIj5MYXN0IDMwIGRheXM8L2J1dHRvbj4KICAgICAgPGJ1dHRvbiB0eXBlPSJidXR0b24iIGRhdGEtcHJlc2V0PSI5MCI+TGFzdCA5MCBkYXlzPC9idXR0b24+CiAgICAgIDxidXR0b24gdHlwZT0iYnV0dG9uIiBkYXRhLXByZXNldD0iYWxsIj5BbGwgdGltZTwvYnV0dG9uPgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJmaWx0ZXItZmllbGQiPgogICAgICA8bGFiZWwgZm9yPSJmaWx0ZXJQbGF0Zm9ybSI+UGxhdGZvcm08L2xhYmVsPgogICAgICA8c2VsZWN0IGlkPSJmaWx0ZXJQbGF0Zm9ybSI+PG9wdGlvbiB2YWx1ZT0iYWxsIj5BbGwgcGxhdGZvcm1zPC9vcHRpb24+PC9zZWxlY3Q+CiAgICA8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImZpbHRlci1maWVsZCI+CiAgICAgIDxsYWJlbCBmb3I9ImZpbHRlckNhbXBhaWduIj5DYW1wYWlnbjwvbGFiZWw+CiAgICAgIDxzZWxlY3QgaWQ9ImZpbHRlckNhbXBhaWduIj48b3B0aW9uIHZhbHVlPSJhbGwiPkFsbCBjYW1wYWlnbnM8L29wdGlvbj48L3NlbGVjdD4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0iZmlsdGVyLWZpZWxkIj4KICAgICAgPGxhYmVsIGZvcj0iZmlsdGVyQ29udGVudFR5cGUiPkNvbnRlbnQgdHlwZTwvbGFiZWw+CiAgICAgIDxzZWxlY3QgaWQ9ImZpbHRlckNvbnRlbnRUeXBlIj48b3B0aW9uIHZhbHVlPSJhbGwiPkFsbCBjb250ZW50IHR5cGVzPC9vcHRpb24+PC9zZWxlY3Q+CiAgICA8L2Rpdj4KICA8L3NlY3Rpb24+CgogIDxtYWluIGNsYXNzPSJ2aWV3LWFyZWEiPgogICAgPHNlY3Rpb24gaWQ9InZpZXctZGFzaGJvYXJkIiBjbGFzcz0idmlldyBpcy1hY3RpdmUiPjwvc2VjdGlvbj4KICAgIDxzZWN0aW9uIGlkPSJ2aWV3LXJlY29yZHMiIGNsYXNzPSJ2aWV3Ij48L3NlY3Rpb24+CiAgICA8c2VjdGlvbiBpZD0idmlldy1mb2xsb3dlcnMiIGNsYXNzPSJ2aWV3Ij48L3NlY3Rpb24+CiAgICA8c2VjdGlvbiBpZD0idmlldy1jb21wYXJpc29uIiBjbGFzcz0idmlldyI+PC9zZWN0aW9uPgogICAgPHNlY3Rpb24gaWQ9InZpZXctcmVwb3J0IiBjbGFzcz0idmlldyI+PC9zZWN0aW9uPgogICAgPHNlY3Rpb24gaWQ9InZpZXctdXBsb2FkIiBjbGFzcz0idmlldyI+PC9zZWN0aW9uPgogICAgPHNlY3Rpb24gaWQ9InZpZXctaGlzdG9yeSIgY2xhc3M9InZpZXciPjwvc2VjdGlvbj4KICA8L21haW4+CiAgPC9kaXY+CjwvZGl2PgoKPGRpdiBpZD0idG9hc3RSb290IiBjbGFzcz0idG9hc3Qtcm9vdCIgYXJpYS1saXZlPSJwb2xpdGUiPjwvZGl2PgoKPHNjcmlwdD4KLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIEJyYW5kIGxvZ28g4oCUIGVtYmVkZGVkIG9uY2UgaGVyZSBhbmQgd2lyZWQgb250byBldmVyeSAuYnJhbmQtbG9nbwogICA8aW1nPiBhbmQgdGhlIGZhdmljb24gPGxpbms+IGF0IGJvb3RzdHJhcCwgc28gdGhlIGJhc2U2NCBwYXlsb2FkCiAgIGFwcGVhcnMgZXhhY3RseSBvbmNlIGluIHRoaXMgZmlsZSBpbnN0ZWFkIG9mIG9uY2UgcGVyIHVzYWdlIHNpdGUuCiAgID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwpjb25zdCBMT0dPX0RBVEFfVVJJID0gJ2RhdGE6aW1hZ2UvcG5nO2Jhc2U2NCxpVkJPUncwS0dnb0FBQUFOU1VoRVVnQUFDRm9BQUFkekNBWUFBQUJuYjhvM0FBQUFDWEJJV1hNQUFDNGpBQUF1SXdGNHBUOTJBQUFnQUVsRVFWUjRuT3pkVFc3YldMb0c0T1BnenVOYWdWMHJNR3NGVVUwOFRTNElHSnBGTzRoM1FIRUh6Z3JhTlJNSUdKVk1PYm5PQ29wWlFUczdTRmFnaXhNeHNWWDVvMjM5SEpMUEEzall5TkYzSkhXaitlcDdENWJMWlFBQUFBQUFBQUFBNE5lZW1CRUFBQUFBQUFBQVFEZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBRDBVRjdWV1Y3VkYrNE9BQUIyNjMvTUd3QUFBQUNnbHk1ZEd3QUE3SjZnQlFBQUFBQkF6K1JWUFE4aG5MZzNBQURZUGRVaEFBQUFBQUE5a2xmMWNRamgvTXVKODZxZXVEOEFBTmdkUVFzQUFBQUFnSDZKbFNGUDc1dzRjMzhBQUxBN2doWUFBQUFBQUQyUlYzWGNaUEhzWDZlMTBRSUFBSFpJMEFJQUFBQUFvQWZ5cWo0TUljeS9jMUliTFFBQVlJY0VMUUFBQUFBQSt1SGZsU0ZmSExVaERBQUFZQWNFTFFBQUFBQUFFcGRYOVlzUXd2T2ZuRko5Q0FBQTdJaWdCUUFBQUFCQXd0cHRGWmUvT0tINkVBQUEyQkZCQ3dBQUFBQ0F0TTEvVUJseWw0MFdBQUN3STRJV0FBQUFBQUNKeXFzNkJpaGVkVGlkalJZQUFMQWpnaFlBQUFBQUFBbnFXQm55eGRPOHFvVXRBQUJnQndRdEFBQUFBQURTZEI1Q09Mckh5UVF0QUFCZ0J3UXRBQUFBQUFBUzAyNm5LTzU1cW9sN0JBQ0E3Uk8wQUFBQUFBQklUOWZLa0x0c3RBQUFnQjBRdEFBQUFBQUFTRWhlMWZNUXdza0RUdlNRL3d3QUFIQlBnaFlBQUFBQUFJbklxL280aEhEKzBOUGtWYTArQkFBQXRrelFBZ0FBQUFBZ0hiRXk1T2tqVGlOb0FRQUFXeVpvQVFBQUFBQ1FnTHlxNHlhTFo0ODhTZVl1QVFCZ3V3UXRBQUFBQUFEMkxLL3F3eERDZkFPbkVMUUFBSUF0RTdRQUFBQUFBTmkveDFhR2ZIR1VWL1d4K3dRQWdPMFJ0QUFBQUFBQTJLTzhxbCtFRUo1djhBUzJXZ0FBd0JZSldnQUFBQUFBN0VsYkdYSzU0WDlkMEFJQUFMWkkwQUlBQUFBQVlIL21HNm9NdVd2aVBnRUFZSHNFTFFBQUFBQUE5aUN2NmhpSWVMV0ZmL21aK3dRQWdPMFJ0QUFBQUFBQTJMRXRWWVo4bFZlMStoQUFBTmdTUVFzQUFBQUFnTjA3RHlFY2JmRmZWUjhDQUFCYkltZ0JBQUFBQUxCRDdiYUpZc3Yvb28wV0FBQ3dKWUlXQUFBQUFBQzd0YlhLa0RzRUxRQUFZRXNFTFFBQUFBQ0FyY2tXNVN4YmxCNzZ0L0txbm9jUVRuYndUNTNrVlgyNGczOEhBQUJHNTM5Y09RQUFBQUN3YVRGZ0VVS0lvWUtiWmxwTURQaHp5T0o0QjVVaGQ4V0F5L1VPL3owQUFCZ0ZRUXNBQUFBQVlHUHVCQ3lPUWdpZlFnZ3owLzFxRjVVaGQwMEVMUUFBWVBNRUxRQUFBQUNBUjhzVzVZc1F3a1Vic1BoaTNreUxHOVA5dk0zaVBJVHdiTWYvck1vV0FBRFlBa0VMQUFBQUFPREJza1U1YVRkWS9EdEU4SzZaRmhjbSs3VXlaTDZIZjFwbEN3QUFiTUVUUXdVQUFBQUE3aXNHTExKRkdXc3AvdThIbXhyT0RmV3JHRGg1dW9kLzkya2I4Z0FBQURiSVJnc0FBQUFBb0xPZmJMQzRxMnltUldPcW43ZFp4RXFWNTNzOFFxd1BVZDhDQUFBYkpHZ0JBQUFBQVB4U3g0QkY5TDZaRnZ1b3lVaE9YdFdISVlUTFBaOHIzdHViQk1jREFBQzlKV2dCQUFBQUFQeFF0aWlQMjRERnk0NVRtcG5tVi91cURMa3IyL08vRHdBQWd5Tm9BUUFBQUFCODR3RUJpK2kxeXBDVnZLb245NXpkdHZ4cUF3a0FBSEJQQjh2bDBzd0FBQUFBZ004ZUdMQ0lQc1QvZURNdFBvNTlrbTFsU0F5Y0hDVnduT2pQcTdQVDZ3VE9BUUFBZzJDakJRQUFBQUR3bUlERkZ6TWhpNi9tQ1lVc1Fsc2ZJbWdCQUFBYkltZ0JBQUFBQUNPV0xjcTRmZUU4aEZBOFlncXhNc1NEL05VMml4aHFlSlhBVWU3SzBqa0tBQUQwbjZBRkFBQUFBSXpRbllCRi9IdjZpQWw4YURjNHNIS1o0QndtQ1p3QkFBQUc0MkM1WExwTkFBQUFBQmlKRFFZc3Z2amZabHE4OGY3NXZNMWkvc2pOSU52MDI5WFpxV29YQUFEWUFCc3RBQUFBQUdBRXRoQ3dpTjRLV2F6a1ZYMmNjTWdpdFBVaDZsMEFBR0FEbmhnaUFBQUFBQXhidGlobklZU2JOZ2l3cVpERnB4REN6RnZucXhRclErNVNId0lBQUJ0aW93VUFBQUFBREZRYnNJaDFGa2RiZUlXelpscW9vbGh0czRoYlFwNGxjSlNmRWJRQUFJQU5FYlFBQUFBQWdJSFpjc0FpZXFjeVpLV3RESm1uY0paZnlKSStIUUFBOUlqcUVBQUFBQUFZaUJpd3lCWmxyQWo1enhaREZpcEQxbDFzc0k1bG01NjJvUkFBQU9DUmJMUUFBQUFBZ0o3TEZtV3NoYmpjWXJqaXJua3pMVzY4Wno1dnMzZ1JRbmlld0ZHNit2SStBUUFBSGtIUUFnQUFBQUI2cWcxWXhOcUtaenQ2QmJFeTVNTDc1WFBJNHJDSG9RWDFJUUFBc0FHQ0ZnQUFBQURRTTNzSVdBU1ZJZC9vUzJYSVhaTjBqZ0lBQVAwbGFBRUFBQUFBUGJHbmdNVVhGeXBEVnZLcWp2ZndNb1d6M05OSnIwNExBQUNKT2xndWwrNEdBQUFBQUJLV0xjcmp0cVppSHdHTDZIMHpMZFJPM0ZhR05DR0Vvd1NPOHhCL1hwMmRYdmZ2MkFBQWtBNGJMUUFBQUFBZ1VXM0FZcDdBOWdTVkliZm1QUTVaUkRFd0kyZ0JBQUNQSUdnQkFBQUFBSWxKS0dBUmxjMjBhQkk0eDk2MWxTR3ZldjR5NG11NFNPQWNBQURRVzRJV0FBQUFBSkNJeEFJVzBZZG1Xc3dUT0VjcWhoQlFVQUVEQUFDUEpHZ0JBQUFBQUh1V0xjckQ5aUYrS2dHTEwxU0d0UEtxam9HVGt5UU84emhIZVZVZlhwMmRmdXp6aXdBQWdIMFN0QUFBQUFDQVBXa0RGdWZ0MzlQRTd1RjFNeTJ1RXpqSDN1VlZIYmRBRkFONlNiRSs1RTBDNXdBQWdGNFN0QUFBQUFDQUhVczhZQkY5YUN0TVdCbENaY2hkbWFBRkFBQThuS0FGQUFBQUFPeElEd0lXWDh5YWFhRmFZclhOSXQ3VnN3U09za21UNGJ3VUFBRFlQVUVMQUFBQUFOaXlIZ1Vzb3I5VWhxemtWWDA4ME0wZVF3dU9BQURBVGoweGJnQUFBQURZbm14UnprSUlUUWloNkVISTRsTWJCbUhsc2dkMzlpQjVWV2M5UERZQUFDVEJSZ3NBQUFBQTJJSTJZQkczSVJ6MWFMNHFRMXA1VmI4WStPYUhyQTBBQVFBQTl5Um9BUUFBQUFBYjFOT0FSZlMybVJadkVqakgzdVZWZmRodXN4aXl5UWhlSXdBQWJJV2dCUUFBQUFCc1FJOERGcUd0REprbGNJNVVETFl5NUE3VklRQUE4RUNDRmdBQUFBRHdDTm1pbkxRQml6N1hUTXhWaHF6a1ZSM3Y4M2tLWjlteWswRy9PZ0FBMktLRDVYSnB2Z0FBQUFCd1R3TUpXRVR2bW1reFNlQWNlOWRXaGpROTNVcnlFSDllbloxZTkrL1lBQUN3WHpaYUFBQUFBTUE5RENoZ0VWU0dmS092MVM4UEZkL0xnaFlBQUhCUGdoWUFBQUFBME1IQUFoWmZ4TXFRbXpTT3NsOXRaY2lya2Izc0xJRXpBQUJBN3doYUFBQUFBTUJQWkl2eU9JUndFVUo0UHJBNXhjcVFpd1RPa1lveHprTFFBZ0FBSHVDSm9RRUFBQURBdDJMQUlsdVVseUdFL3c0d1pCR2RKM0NHSk9SVkhUZVZuSXp3cFIvbFZYMmN3RGtBQUtCWGJMUUFBQUFBZ0R2YURSYnh3ZnZMQWMrbGJLWkZrOEE1OWk2djZyalZvUmp4Q09MclZ4OERBQUQzSUdnQkFBQUFBT01KV0VUdm0ya3hUK0FjcVJoN2Zjb2toUEFtZ1hNQUFFQnZDRm9BQUFBQU1HclpvanhzQXhhdlJqS0hXUUpuU0VKZTFiRSs1ZG5JeDVBbGNBWUFBT2dWUVFzQUFBQUFScWtOV0p5M2YwOUhNb1BYS2tOVzhxcitzc0ZrN01ZZU5BRUFnSHNUdEFBQUFBQmdWRVlhc0lnK0NCYXN1UnpaL2Y5UVh0WFoxZG1wQUE0QUFIUWthQUVBQUFEQUtJdzRZUEhGckprV0g5TTR5bjdsVmYzQ0pvYzFreENDb0FVQUFIUWthQUVBQUFEQTRHV0w4cnpkNWpEV0RRYXhNdVE2Z1hQc1hWN1ZoKzAyQzI1bFpnRUFBTjBKV2dBQUFBQXdXTm1pbkxVQmk2TVIzN0xLa0hVcVE3NGxhQUVBQVBkd3NGd3V6UXNBQUFDQVFSR3dXUE8vemJSNGs5QjU5cWF0RFBsN3BDLy9WMzY3T2p0VkxRTUFBQjNZYUFFQUFBREFZQWhZZk9PdGtNVktXeGx5a2NKWkVoVzNXcWlYQVFDQURnUXRBQUFBQU9pOWJGRythQitpQzFqYytoUkNtS1Z5bUFRSTRQemNSTkFDQUFDNkViUUFBQUFBb0xleVJUbHBINkEvYzR2Zm1EWFRRaFhFYXB0RmZKKzhTdUFvS2N2R1BnQUFBT2hLMEFJQUFBQ0EzaEd3K0tWM0trUFdYQ1owbGxSTnhqNEFBQURvNm9sSkFRQUFBTkFYTVdDUkxjcFliL0IvUWhZL3BETGtqcnlxVllaMDh6U3Y2dU0rSEJRQUFQYk5SZ3NBQUFBQWtwY3R5bGhyY0NGYzBjbThtUlkzUFRqbjF1VlZIZDgzeGNCZjVpWk5iUDhBQUlCZkU3UUFBQUFBSUZuWm9qeHVLMEpldXFWT1ltWElSUS9PdVN0Q0EvZVQ5ZW13QUFDd0w0SVdBQUFBQUNSSHdPSkJWSWJja1ZmMWVRamhKSmtEOVlPZ0JRQUFkQ0JvQVFBQUFFQXlCQ3dlNVVKbHlFcGUxVi9lUjl5UGFoNEFBT2pnWUxsY21oTUFBQUFBZXlWZzhXanZtMmxoRzBFcnIrcHJvWUVIKy9QcTdQUzZwMmNIQUlDZHNORUNBQUFBZ0wzSkZ1VmhDT0c4L1h2cUpoNU1aVWdycitxWmtNV2p4TUNPb0FVQUFQeUVvQVVBQUFBQU95ZGdzVkZsTXkyYUFiMmVCOHVyT3I2dkxucDYvRlRZakFJQUFMOGdhQUVBQUFEQXpnaFliTnlIWmxyTUIvYWFIdVBTKytyUkpqMC9Qd0FBYkoyZ0JRQUFBQUJiSjJDeE5TcERXbmxWdndnaFBFL2lNUDEyRkRlRFhKMmRmaHo3SUFBQTRFZWVtQXdBQUFBQTI1UXR5aGdHdUFraEZFSVdHL1c2bVJiWEEzbzlENll5Wk9QVWh3QUF3RS9ZYUFFQUFBREFWclFCaTFocmNXVENHL2VoblMwcjNtZWJGZXREaEhnQUFPQUhCQzBBQUFBQTJDZ0JpNTJZTmROQ3RjTnFtMFVNQmJ4SzRDaERNaG43QUFBQTRHY0VMUUFBQUFEWUNBR0xuZmxMWmNpYXk0VE9NaFNxUXdBQTRDZWVHQTRBQUFBQWo1RXR5a20yS0pzUXduK0VMTGJ1VXdqaGZPQ3ZzYk84cWdWN3R1TnBYdFhDRmdBQThBTTJXZ0FBQUFEd0lERmcwVzZ3ZUdhQ082TXlwTlVHQVlva0RqTk1jYjdOMkljQUFBRGZJMmdCQUFBQXdMMElXT3pOMjJaYXZCbnBhLzhlbFNIYlphTUZBQUQ4Z0tBRkFBQUFBSjBJV094VnJBeVpqZmoxcjhtck90YW5uQ1IwcENHYWpIMEFBQUR3STRJV0FBQUFBUHhVdGlpUDIrMEJBaGI3TTFjWnNwSlg5WEViK0dHN0JGa0FBT0FIbmhnTUFBQUFBTjhUQXhiWm9vd0JpLzhLV2V6VnUyWmFYSXo0OWY5YmZFOCtUZXRJdzVSWHRhMFdBQUR3SFRaYUFBQUFBTENtM1dBUk53YThOSm05VXhseVIxc1pJdlN6TzFrSTRYb3NMeFlBQUxvU3RBQUFBQURnTXdHTEpNWEtrSnV4RHlHc1FoYUhLa04yTG02MHNFMEZBQUQrUmRBQ0FBQUFZT1N5UmZubEFmYXJzYzhpTVNwRDFxa00yYjFzYkM4WUFBQzZPRmd1bHdZRkFBQUFNRUp0d09LOC9mTUFPejEvTk5PaUdmc1F3bXFieFlzUXd0OEpIR1dNZnJzNk8vMDQ5aUVBQU1CZE5sb0FBQUFBakl5QVJTK1VRaFlyYldYSVpRcG5HYWxZSC9KbTdFTUFBSUM3QkMwQUFBQUFSa0xBb2pmZU45TmlQdlloM0RIM2Z0MnJUTkFDQUFEV0NWb0FBQUFBakVDMktPY0NGcjB4Ry9zQXZzaXJPbTVUZUpYR2FVWnJNdllCQUFEQXZ3bGFBQUFBQUF4WXRpaG43VWFBSS9mY0M2OVZocXlvREVuR3M3RVBBQUFBL3UySmlRQUFBQUFNVHd4WVpJdnlKb1R3SHlHTDN2alFobUpZT2ZmZVRVTmUxZG5ZWndBQUFIZlphQUVBQUFBd0lEWlk5TnFzbVJZZnh6NkVjUHRndjBqZ0tLekUrN0JwQlFBQVdvSVdBQUFBQUFNZ1lORjdzVExrZXV4RHVFTmxTRm9tN2dRQUFHNEpXZ0FBQUFEMFdMWW9KMjNBNHBsNzdDMlZJWGZrVlIxbmNaTE1nUWp0UmdzQUFLQWxhQUVBQUFEUVF3SVdnM0t1TW1RbHIrcmpPSThVenNJYXdSY0FBTGpqaVdFQUFBQUE5RWNNV0dTTE1sWk0vSitReFNDOGJhYkZtN0VQNFk1WVQvRTBtZFB3VlY3VkU5TUFBSUFWR3kwQUFBQUFlc0FHaTBINkZFS1lqWDBJWCtSVmZlNzluYlQ0SFhROTlpRUFBRUFRdEFBQUFBQklXN1lvajl1QXhVdFhOVGd6bFNFcmVWVWZ0dTl6MHBXNUd3QUFXQkcwQUFBQUFFaVFnTVhndlZNWnNrWmxTUHBVaHdBQVFPdUpRUUFBQUFDa0l3WXNza1VaSHpyL1Y4aGlzRlNHM0pGWDlZc1F3dk5rRHNTUFBNMnIrdGgwQUFEQVJnc0FBQUNBSk5oZ01TcnpabHJjakgwSTRiWXk1REtCbzlCTnJBL3gzZ1VBWVBRRUxRQUFBQUQyS0Z1VThVSHplUWloY0ErakVDdERMc1kraER2bUtrTjZKZGFIcUx3QkFHRDBCQzBBQUFBQTl1Qk93T0xjZytiUlVCbHlSMTdWOGFIOXEyUU9SQmVaS1FFQWdLQUZBQUFBd0U0SldJemFoY3FRRlpVaHZmVnM3QU1BQUlBZ2FBRUFBQUN3R3dJV28vZSttUmJ6c1EvaGp2ZzVPRXJtTkhTV1YzVjJkWGJhbUJnQUFHUDJ4TzBEQUFBQWJGZTJLR05kUk54a1VBaFpqSmJLa0ZaOFVOOStGdWluaVhzREFHRHNiTFFBQUFBQTJKSTJZREgzeS8zUks1dHBZUVBBTFpVaC9aYU5mUUFBQUNCb0FRQUFBTEJoQWhiYzhVRmx5SzI4cXVNc1RsSTVEdzhpYUFFQXdPZ2RMSmZMc2M4QUFBQUFZQ01FTFBpT1A1dHBjVzB3bjBNV3h5R0VSbjNPSVB4MmRYYjZjZXhEQUFCZ3ZHeTBBQUFBQUhpa2JGRk8yam9FQVF2dWVpMWtzZVpTeUdJdzRsWUw3MjBBQUVaTDBBSUFBQURnZ2RxQVJkeGc4Y3dNK1pjUDdYdUQxVGFMYzUrVFFaa0lXZ0FBTUdhQ0ZnQUFBQUQzSkdCQkI3Tm1XcWhXdUswTUVUb1psc25ZQndBQXdMZ0pXZ0FBQUFCMEpHQkJSMytwREZsem9USmtjTEt4RHdBQWdIRTdXQzZYWTU4QkFBQUF3RTlsaXpKckh4WUxXUEFybjBJSXg3WlpyT1JWL1NLRThIY0taMkhqZnI4Nk83MHhWZ0FBeHNoR0N3QUFBSUFmeUJibGw4cURsMlpFUnlwRFdubFZINFlRTHBNNEROc3djYjhBQUl5Vm9BVUFBQURBdndoWThFQnZtMm54eHZDK1Voa3liT3BEQUFBWUxVRUxBQUFBZ0phQUJZOFFLME5tQnJpU1YvWEU1Mmp3QkMwQUFCZ3RRUXNBQUFCZzlMSkZlZGorK3Q2RFlSNXFyakprUldYSWFEd2Ird0FBQUJpdkorNGVBQUFBR0tzWXNNZ1daZHhnY1NOa3dTTzhhNmJGaFFGK0ZUOVRSNG1jaFMxcU41Y0FBTURvMkdnQkFBQUFqRTY3d2VLOC9YdnFIY0FqcUF5NUk2L3FXQ2Z4S3BrRHNXM3h2cTlOR1FDQXNSRzBBQUFBQUVaRHdJSXRpSlVoTndiN2xjcVFjY25HUGdBQUFNWkowQUlBQUFBWVBBRUx0a1JseUIxNVZjZktrSk5rRHNRdXFBNEJBR0NVbnJoMkFBQUFZTWl5UlJsckhab1FRaUZrd1lhZEcraEtYdFhIN1dlTWNUbktxL3JRblFNQU1EWTJXZ0FBQUFDRDFBWXM0aS9zajl3d1cxQTIwNkl4Mks5VWhveFgzR3J4WnV4REFBQmdYQVF0QUFBQWdFRVJzR0FIM2pmVFltN1FLM2xWeDgwZXoxSTRDM3VSQ1ZvQUFEQTJnaFlBQUFEQUlBaFlzRU16dzE1cEswT0VUc1p0TXZZQkFBQXdQb0lXQUFBQVFLOWxpM0xTUHVqMWkzcDI0YlhLa0RVWElZU25DWjJIM2N2TUhBQ0FzVGxZTHBjdUhRQUFBT2dkQVF2MjRFTjg2elhUNHFQaGY5NW04U0tFOEhjQ1IySC8vcmc2T3hWQUFnQmdOR3kwQUFBQUFIcEZ3SUk5bWdsWnJPUlZmUmhDdUV6aExDUWhiclVRdEFBQVlEUUVMUUFBQUlCZUVMQmd6MkpseUxWTCtFcGxDSGVwRHdFQVlGUUVMUUFBQUlDa1pZdnl1SDJvKzl4TnNTY2YycEFQcTIwV01mVDAwaXk0WTJJWUFBQ01pYUFGQUFBQWtLUTJZREgzUUpjRW5Lc01XVkVad2crY0dBd0FBR1B5eEcwREFBQUFLWWtCaTJ4UnhnZTUveFd5SUFGdm0ybnh4a1Y4RmNOUFI0bWNoWVMwbTA0QUFHQVViTFFBQUFBQWttQ0RCUW42RkVLWXVaaVZ2S3F6RU1LckZNNUNrdUw3NDlyVkFBQXdCb0lXQUFBQXdGNWxpL0t3RFZoNGdFdHFaaXBEMXFnTTRXZmlSb3NMRXdJQVlBd0VMUUFBQUlDOWFBTVc1KzNmVTdkQVl0NnBETG1WVjNVTVE1MmtjaDZTbExrV0FBREc0bUM1WExwc0FBQUFZR2NFTE9pQldCbVNOZFBpeG1WOXJRejVKNEdqa0w3ZnI4NU9mVzRBQUJnOEd5MEFBQUNBblJDd29FZm1RaFpyMUVIUVZRemwrT3dBQURCNFQxd3hBQUFBc0czWm9qeHZINzRWUWhZa0xsYUdDQmEwOHFxT245MW5TUnlHUGxBZkFnREFLTmhvQVFBQUFHeE50aWhuY1R0QUNPSElsT21CV0JreWMxRXJlVlVmdDU5ZjZHcGlVZ0FBaklHZ0JRQUFBTEJ4QWhiMDFJWEtrRFdYTnRCd1Q3YWZBQUF3Q2dmTDVkSk5Bd0FBQUJzaFlFR1B2VyttaGRxRFZsN1ZMMElJZnlkeEdQcm1qNnV6MDhhdEFRQXdaRFphQUFBQUFJK1dMY3I0VVBaQ3dJSWVVeG5TeXF2NnNOMW1BUThSQTB1Q0ZnQUFESnFnQlFBQUFQQmcyYUtjdEJzc3JJdW56OHBtV25nd2ZFdGxDSTh4RWRRQkFHRG9CQzBBQUFDQWV4T3dZRUErTk5OaTdrSlg4cXFPbiszbktaeUYzbExCQXdEQTRCMHNsMHUzREFBQUFIUWlZTUVBL2RsTWkyc1grN1V5cEZFQnhBYjhkblYyK3RFZ0FRQVlLaHN0QUFBQWdGL0tGbVg4aGZLRmdBVUQ4MXJJWXMxY3lJSU5pZitkNGJNRkFNQmdDVm9BQUFBQVA1UXR5dVAyNGV0TFUySmdQclR2Ylc0clExNlpCUnN5RWJRQUFHRElCQzBBQUFDQWJ3aFlNQUt6Wmxxb05yaDFrY3BCR0lUTU5RSUFNR1NDRmdBQUFNQlhBaGFNeEY4cVEyN2xWUjAvOHllcG5JZEJtTGhHQUFDRzdHQzVYTHBnQUFBQUdEa0JDMGJrVXdqaDJEYUxsYnlxNCthQmYxSTRDNFB6KzlYWjZZMXJCUUJnaUd5MEFBQUFnQkhMRnVWaENPRzgvWHZxdmNBSXFBeFpwektFYllraEhrRUxBQUFHU2RBQ0FBQUFSa2pBZ3BGNjIweUxOeTUvSmEvcStQbC9sc0paR0tSWUgrTHpCZ0RBSUFsYUFBQUF3SWdJV0RCaXNUSms1ZzJ3a2xmMWw3b2cySmJNWkFFQUdDcEJDd0FBQUJnQkFRc0ljNVVoYXk1OUY3Qmx0cVVBQURCWUI4dmwwdTBDQUFEQWdHV0xNdjZLLzhKRFZVYnNYVE10SnQ0QUszbFZ2d2doL0ozQ1dSaThQNjdPVGh2WERBREEwTmhvQVFBQUFBUFZCaXhpTmNDUk8yYkVWSWJja1ZmMVlidk5BblloQnB3RUxRQUFHQnhCQ3dBQUFCZ1lBUXRZRXl0RGJvemtLNVVoN0ZKbTJnQUFESkdnQlFBQUFBeUVnQVY4STFhR1hCakxTbDdWY2J2QTh4VE93bWlvN0FFQVlKQU9sc3VsbXdVQUFJQWV5eFpsZkpBVkh5YWZ1RWRZODBjekxkUVczRmFHTklKWTdNRnZWMmVuSHcwZUFJQWhzZEVDQUFBQWVxb05XTVFORnMvY0lYeWpGTEpZWTlzTit4THJRNjVOSHdDQUlSRzBBQUFBZ0o0UnNJQmZldDlNaTdreHJiU1ZJYTlTT0F1ak5CRzBBQUJnYUFRdEFBQUFvQ2NFTEtDem1WR3R1VXpvTEl6UHhKMERBREEwZ2hZQUFBQ1F1R3hSSHJjUFNnVXM0TmRlcXd5NWxWZTF5aEQyTFhNREFBQU16Y0Z5dVhTcEFBQUFrS0EyWUJFZmtyNTBQOURKaC9qUmFhYkZSK1A2SExLSUQ3ai9TZUFvOFB2VjJlbk42S2NBQU1CZzJHZ0JBQUFBaVJHd2dBZWJDVm1zVVJsQ0tpYmVqd0FBREltZ0JRQUFBQ1JDd0FJZUpWYUdYQnZoU2w3VjV5R0VreFRPQXVwREFBQVlHa0VMQUFBQTJMTnNVUjYyQVl0WDdnSWU1RVA3R1dJVnNqZzJEeElqYUFFQXdLQUlXZ0FBQU1DZXRBR0w4L2J2cVh1QUJ6dFhHYkxtMG5jS2lYbm1RZ0FBR0pLRDVYTHBRZ0VBQUdDSEJDeGdvOTQyMCtLRmthN2tWVDBMSWZ3bmhiUEF2L3g1ZFhhcTNnY0FnRUd3MFFJQUFBQjJSTUFDTnU1VENHRm1yQ3Q1VmNmdm1Jc1V6Z0xmRWV0REJDMEFBQmlFSjY0UkFBQUF0aTlibFBNUXdrMElvUkN5Z0kyWnFReFpvektFbEUzY0RnQUFRMkdqQlFBQUFHeFJ0aWpqcisxanlPTEluR0dqM2pYVDRvMlJydVJWSGV0VG5xZHdGdmlCekdBQUFCaUtnK1Z5NlRJQkFBQmd3d1FzWUt0aVpValdUSXNiWS81YUdkTDR2cUVIZnJzNk83V0ZCZ0NBM3JQUkFnQUFBRFpJd0FKMllpNWtzY1ozRG4wUjYwTnNvZ0VBb1BjRUxRQUFBR0FEQkN4Z1oySmx5SVZ4citSVkhSOWN2MHJoTE5CQkptZ0JBTUFRQ0ZvQUFBREFJMlNMY3RJR0xKNlpJMnhkckF5WkdmT2F5NFRPQXI4eU1TRUFBSVpBMEFJQUFBQWVRTUFDOXVKQ1pjaXR2S3B0MGFGdk1qY0dBTUFRSEN5WFN4Y0pBQUFBSFFsWXdONjhiNmFGaDdTdHZLcmpMUDVKNGpCd1AzOWNuWjAyWmdZQVFKL1phQUVBQUFBZENGakEzcWtNV2FjeWhMNktJU0ZCQ3dBQWVrM1FBZ0FBQUg0aVc1VEhiY0RpcFRuQjNwVE50UEJndHBWWDlYa0k0U1NKdzhEOTJVd0RBRUR2Q1ZvQUFBREFkd2hZUURJK05OTmk3anBXOHFyKzh0MEVmVFZ4Y3dBQTlKMmdCUUFBQU53aFlBSEpVUm15TGxhR1BFM3BRSEJQdHJFQUFOQjdCOHZsMGkwQ0FBQXdlZ0lXa0tUWHpiUTRkelVyZVZYSDBNbC9VamdMUE5LZlYyZW4xNFlJQUVCZjJXZ0JBQURBcUdXTDhqQ0VFQi9rRm1PZkJTVG1nNHFNVzNsVngrK3FpMVRPQTQ4VTYwTUVMUUFBNkMxQkN3QUFBTU05aHNzQUFDQUFTVVJCVkVicFRzRGkzQnArU05Lc21SWWZYYzFYS2tNWWtzeHRBZ0RRWjRJV0FBQUFqSXFBQmZUQ1g4MjA4R3YzVmw3VkwwSUl6NU00REd5R29BVUFBTDEyc0Z3dTNTQUFBQUNESjJBQnZmRXBoSEJzbThWS1d4bHk0M3VMQWZyOTZ1ejB4c1VDQU5CSE5sb0FBQUF3ZU5taW5JVVFManlvaEY1UUdiSnU3cnVMZ2NyYUVCRUFBUFNPb0FVQUFBQ0QxUVlzNGtQS0k3Y012ZkMybVJadlhOVktYdFdURU1LckZNNENXeENERmo3dkFBRDBrcUFGQUFBQWd5TmdBYjBVSzBObXJtNmxyUXk1VE9Fc3NDVVRnd1VBb0s4RUxRQUFBQmdNQVF2b3RibktrRFhudnNzWXVHY3VHQUNBdmpwWUxwY3VEd0FBZ0Y3TEZ1V2svZVczaDVMUVQrK2FhZUhYN2EyOHFtT2x3ajlKSEFhMjY0K3JzOVBHakFFQTZCc2JMUUFBQU9pdE5tQXg5NnRZNkRXVklkOVNHY0pZeEZDUm9BVUFBTDBqYUFFQUFFRHZDRmpBb01US2tCdFh1cEpYZGZ4dU8wbmhMTEFERThFaUFBRDZTTkFDQUFDQTNoQ3dnTUdKbFNFWHJuVWxyK3JqRU1KNUNtZUJIY2tNR2dDQVBoSzBBQUFBSUhuWm9vd1BZaTRFTEdCd2hBcld4Vi8yUDAzcFFMQmxKM2xWSDE2ZG5YNDBhQUFBK2tUUUFnQUFnR1JsaS9LNDNXRHgwaTNCNEpUTnRHaGM2MHBlMWVmQ1pJeFVERk5ldTN3QUFQcEUwQUlBQUlEa0NGakE0TDF2cHNYY05hL0VYL1MzMzNrd1JoTkJDd0FBK2tiUUFnQUFnR1FJV01Cb3pGejFHcFVoakZubTlnRUE2QnRCQ3dBQUFQWXVXNVR4MTl3WEFoWXdDcTlWaHR6S3EvcEZDT0Y1S3VlQlBaZ1lPZ0FBZlhPd1hDNWRHZ0FBQUh2UkJpek8yeisvNW9iaCt4QS8rczIwK09pdXYxYUczUGorZy9ENzFkbnBqVEVBQU5BWE5sb0FBQUN3Y3dJV01Gb3pJWXMxYzkrQjhGbldobzRBQUtBWEJDMEFBQURZR1FFTEdMVllHWEk5OWlGOGtWZDFyRXQ0bGNacFlPL2k1K0dOYXdBQW9DOEVMUUFBQU5nNkFRc1l2US90OWdadUswTXV6UUsreW93Q0FJQStFYlFBQUFCZ3E3SkZPV3Nmc0I2Wk5Jeld1Y3FRTmVlK0UySE5NK01BQUtCUERwYkxwUXNEQUFCZzR3UXNnTmJiWmxxOE1JeVZ2S3JqTC9mL1NlRXNrSmcvcjg1TzFRc0JBTkFMTmxvQUFBQ3dVUUlXd0IyZlFnZ3pBMW1qTWdTK0w0YVFCQzBBQU9nRlFRc0FBQUEyUXNBQytJNlp5cEJiZVZYSDc4aVRWTTREaWNsY0NBQUFmU0ZvQVFBQXdLTmtpM0xTQml6MHF3TjN2V3VteFJzVFdjbXIramlFY0o3Q1dTQlJFeGNEQUVCZkNGb0FBQUR3SUFJV3dFK29EUGxXckF4NW10cWhJQ0ZIZVZVZlhwMmQyb0lEQUVEeUJDMEFBQUM0RndFTG9JTjVNeTF1REdvbHIrcHozNW5RU2F3UHVUWXFBQUJTSjJnQkFBQkFKd0lXUUVleE11VENzRmJpTC9UYjcwN2cxeWFDRmdBQTlJR2dCUUFBQUQrVkxjcmpFRUo4YVByY3BJQmZVQm55TFpVaDBOM0VyQUFBNkFOQkN3QUFBTDZyRFZqRVgyRy9OQ0dnb3d1VkliZnlxbjRocEFiM2toa1hBQUI5Y0xCY0xsMFVBQUFBWHdsWUFBLzB2cGtXSHBLMjJzcVFHOXNzNE41K3Z6bzdGZGdDQUNCcE5sb0FBQUR3bVlBRjhFZ3FROVpkQ0ZuQWcwemF5aDBBQUVpV29BVUFBTURJWll2eXNBMVl2QnI3TElBSEs1dHAwUmpmU2w3VkU2RTFlRENiY1FBQVNKNmdCUUFBd0VpMUFZdno5cyt2cm9HSCt0Qk1pN25wcmJTVklYNk5EdzgzTVRzQUFGSW5hQUVBQURBeUFoYkFocWtNV1JkREowY3BIUWg2NXNTRkFRQ1F1b1BsY3VtU0FBQUFSa0RBQXRpQzE4MjBPRGZZbGJ5cVkrWEJQeW1jQlhydXo2dXowMnVYQ0FCQXFwNjRHUUFBZ09ITEZtVjhFSG9UUWlpRUxJQU4rZEJ1YitDV3loRFlqTXdjQVFCSW1lb1FBQUNBQWNzVzVjd2FlMkJMWnMyMCtHaTRLM2xWejFVZXdNWk1RZ2dYeGdrQVFLb0VMUUFBQUFaSXdBTFlzcithYVdHdGZ5dXY2dU4yWXhDd0dUWmFBQUNRTkVFTEFBQ0FBUkd3QUhiZ1V3amgzS0RYcUF5QnpUcktxL3J3NnV6VTFod0FBSklrYUFFQUFEQUEyYUo4MGE3WUZyQUF0azFseUIxNVZjZlF5Yk5rRGdUREVldEQzcmhQQUFCU0pHZ0JBQURRWTltaW5MUWJMRHprQTNiaGJUTXRQUGhzdFpVaDh5UU9BOE9UQ1ZvQUFKQXFRUXNBQUlBZUVyQUE5aUJXaHN3TWZrM2NKUFEwb2ZQQWtFemNKZ0FBcVJLMEFBQUE2QkVCQzJDUDVpcERidVZWSFN1Ym5xZHlIaGlnektVQ0FKQ3FnK1Z5NlhJQUFBQVNseTNLclAzbHRJQUZzQS92bW1uaDErV3R2S29QUXdnM3RsbkExdjF4ZFhiYUdETUFBS214MFFJQUFDQmgyYUk4YmpkWXZIUlB3SjZvRFBtV3loRFlqUmcwRmJRQUFDQTVnaFlBQUFBSkVyQUFFaElyUTI1Y3lFcGUxUlBmemJBejhmTjJhZHdBQUtSRzBBSUFBQ0FoQWhaQVltSmx5SVZMV1drclF6ejBoZDNKekJvQWdCUUpXZ0FBQUNSQXdBSkkxTG1MV1JPL3A0OFNPZzhNM1lrYkJnQWdSUWZMNWRMRkFBQUE3RW0yS0EvYkI1bm4rdjZCeEpUTnRKaTdsSlc4cXVNdjYvOUo0U3d3TW45ZW5aMWV1M1FBQUZKaW93VUFBTUFlQ0ZnQWlYc3ZaUEVObFNHd0g1TVFncUFGQUFCSkViUUFBQURZSVFFTG9DZG1MdXBXWHRWekZRYXdONW5SQXdDUUdrRUxBQUNBSFJDd0FIcmtkVE10R2hlMmtsZjFjUWloU09Fc01GS0NGZ0FBSk9lSkt3RUFBTml1YkZIR1g0YmZ0QS9xaEN5QWxIMElJYWdNV2FjeUJQYnJxQTA4QVFCQU1teTBBQUFBMkpJMllCRWZXQjZaTWRBVHMyWmFmSFJaSzNsVnh5MUV6MUk0QzR4YzFvWldBUUFnQ1lJV0FBQUFHeVpnQWZSVXJBeTVkbmtyN1Mvb2JmZUFOTVNneFJ0M0FRQkFLZ1F0QUFBQU5rVEFBdWd4bFNIZnVsVDNCTW1ZdUFvQUFGSWlhQUVBQVBCSTJhS00vK2YvUlFqaHhDeUJuanBYR1hJcnIrb1hLa01nS1Q2UEFBQWs1V0M1WExvUkFBQ0FCMmdERm5QLzV6L1FjMitiYWZIQ0phN2tWWDBZUXJpeHpRS1M4OGZWMldualdnQUFTSUdORmdBQUFQY2tZQUVNeUtjUXdzeUZybEVaQW1tSy8vdEwwQUlBZ0NRSVdnQUFBSFFrWUFFTTBFeGx5SzI4cXVQMy9QTlV6Z09zeVl3REFJQlVDRm9BQUFEOFFyWW9qOXRmT0F0WUFFUHlycGtXYjl6b1Nsc1pjcG5DV1lEdkVyUUFBQ0FaVDF3RkFBREE5OFdBUmJZbzQwTzMvd3BaQUFPak11UmJjV1BSVVdxSEFyNDZhUU5SQUFDd2R6WmFBQUFBL0V1N3dTSStjSHRwTnNCQXpadHBjZU55VjlyS2tGY3BuQVg0cWJqVjR0cUlBQURZTjBFTEFBQ0Fsb0FGTUJLeE11VENaYTh4RCtpSGlhQUZBQUFwRUxRQUFBQkdMMXVVaDIzQXdxK1pnYUZUR2ZJdmVWWEg3LytUcEE0Ri9FaG1NZ0FBcEVEUUFnQUFHSzAyWUhIZS9qMzFUZ0JHNEVKbHlLMjhxdU5EMnlLVjh3Qy9OREVpQUFCU2NMQmNMbDBFQUFBd0tnSVd3RWk5YjZhRlg0UGZrVmQxckNCNGxzeUJnQzUrdnpvN0ZSZ0RBR0N2YkxRQUFBQkdROEFDR0RtVklYZmtWWDB1WkFHOUZMZGFYTG82QUFEMlNkQUNBQUFZaFd4UnpnVXNnQkVybTJuUmVBT3M1RlY5SEVLWXAzQVc0TjVzNWdFQVlPOEVMUUFBZ0VITEZ1V3NmWmgyNUthQmtmclFUQXVoZ25XWGduZlFXNElXQUFEc25hQUZBQUF3U0FJV0FGK3BETGtqcitvWEtrT2cxM3grQVFEWXU0UGxjdWtXQUFDQXdSQ3dBRmp6dXBrVzUwYXlrbGYxWVFqaHhqWUw2TDAvcjg1T3IxMGpBQUQ3WXFNRkFBQXdDQUlXQU4vNDBINHZja3RsQ0F4RHJBOFJ0QUFBWUc4RUxRQUFnRjdMRnVXa2ZaQm9qVFRBdWxrekxUNmF5VXBlMWZHL0w1Nm5jQmJnMFRJakJBQmdud1F0QUFDQVhoS3dBUGlwdjVwcDRkZmVyYll5NURLSnd3Q2JNREZGQUFEMlNkQUNBQURvRlFFTGdGLzZGRUk0TjZZMXFxVmdXSTVpZ09ycTdOVFdIZ0FBOWtMUUFnQUE2QVVCQzRET1ZJYmMwVmFHdkVybVFNQ214UG9RbTNzQUFOZ0xRUXNBQUNCcDJhSThiZ01XTDkwVXdDKzliYWJGRzJOYWM1SFFXWURObVFoYThQL3MzVTl1M05iV0wrenQ0UFNsT3dMcGpNQThJMUNsdzY3OWdRQ2hudldPSUxvalVOVUlqaktDMSs0VkNCRFg2Ykx6U2lPNDVSRmNhUWJXQ1BSaEovU0o0L2lQL2xUVjV0NThIc0ROQU9SYUZaRlYvSEV0QUlCVUJDMEFBSUJKRXJBQWVMUzRNdVJNMmY3VWRFTzhqcnljeXZFQVc3VlFUZ0FBVWhHMEFBQUFKa1hBQXVESmxsYUcvS25waHJoVzRHSXF4d05zWGFXa0FBQ2s4dUwrL2w3eEFRQ0E1QVFzQUo3bGVuTjY0ZTN1enpUZHNESE5Bb3Izcjc2dE45b01BTUMrbVdnQkFBQWtWYTFYaHlHRWMyOGRBenlabFNGZmFMcmhYTWdDWmlGT3RSQzBBQUJnN3dRdEFBQ0FKRDRMV01SL0I3b0E4R1J4WmNpTjh2Mmg2WVpQRTVLQThsa2ZBZ0JBRW9JV0FBREFYZ2xZQUd4VlhCbHlxYVIvOGRiMUJXYkR5aVFBQUpJUXRBQUFBUFpDd0FKZ0o4NlY5VTlOTjhRVktpZFRPUjVnNTZ3SUFnQWdpWitVSFFBQTJMVnF2WW9QQXVOWSt3c2hDNEN0V1cxT0x6YksrWWVtRzJLZ3ozUVBtSm1tRzB5MUFBQmc3MHkwQUFBQWRxWmFyODdHUGZsSHFneXdWUjgycHhkTEpmMExLME5nbnFvUXdwWGVBd0N3VDRJV0FBREExZ2xZQU96Y21STC9xZW1HMXlHRVYxTTVIbUN2RnFiWkFBQ3diNElXQUFEQTFnaFlBT3pGcjFhRy9NbktFSmk5YXU0RkFBQmcvd1F0QUFDQVo2dldxOFU0c2wzQUFtQzNic2RBRzM4UzhJTjVPNHFCcTc2dFA4NjlFQUFBN0krZ0JRQUE4R1Jqd0NJKzREcFJSWUM5T051Y1huaVlPR3E2SVY2SGZwbkV3UUFweGI4RjczVUFBSUI5RWJRQUFBQWVUY0FDSUltNE11Uks2Zi9pN1lTT0JVaW5FclFBQUdDZkJDMEFBSUFIRTdBQVNNYktrQzgwM1dCbENQREpRaVVBQU5nblFRc0FBT0NIcXZVcXZpVjRLV0FCa015NWxTRi9hcm9oWHBjdXBuSThRSEx1VVFFQTJLdWZsQnNBQVBpV2FyMDZydGFyT0piOS8vb0JHeUNaM3phbkYwYmkvNVdWSWNCZmpBRXNBQURZQ3hNdEFBQ0F2NGtCaTNGRS9SdlZBVWpxTG9Sd3BnVi9hcnJoUElUd2NpckhBMHhHREZwc3RBTUFnSDBRdEFBQUFQNUR3QUpnY3M2c0RQbFQwdzJmcmxNQVgxcVlkZ01Bd0w0SVdnQUFBREZnY1JoQ3VCU3dBSmlVYXl0RC9pWStSRDJZMkRFQjAyQjFDQUFBZXlOb0FRQUFNellHTE03SGZ4NWNBVXlIbFNGZmFMb2gxdU5rVWdjRlRJbVZRZ0FBN00yTCsvdDcxUVlBZ0prUnNBQ1l2UCs5T2IyNDFLWS9OTjBRcjFzM3JsbkFEL3pjdC9XVklnRUFzR3NtV2dBQXdJd0lXQUJrNFZySTRtK3NEQUVlWWhGQ0VMUUFBR0RuQkMwQUFHQUdCQ3dBc21GbHlCZWFibmdkUW5nMXFZTUNwcXJTR1FBQTlrSFFBZ0FBQ2xldFYvR0IzVEtFY0tUWEFKTjN1VG05dU5HbVA0d3JRMHozQUI1SzBBSUFnTDE0Y1g5L3I5SUFBRkFnQVF1QTdIelluRjU0U1BpWnBodGl5T0tYeVJ3UWtJTi85bTB0c0FZQXdFNlphQUVBQUlVUnNBRElscFVobjJtNllTRmtBVHhCREt3SldnQUFzRk9DRmdBQVVBZ0JDNENzclRhbkZ4c3QvTU80TXVUdEZJNEZ5RTRNYWIzWE5nQUFka25RQWdBQU1sZXRWNHR4Zi8xTHZRVEkwdTNtOUdLcGRYOXhMamdJUEpFVlRBQUE3SnlnQlFBQVpHb01XTVFIY3lkNkNKQTFLME0rMDNSRGZFaDZNWmtEQW5MajNoZ0FnSjM3U1lrQkFDQXZNV0JSclZkWElZVC84VU15UVBaKzNaeGVYR25qWDFnWkFqekxHTmdDQUlDZE1kRUNBQUF5WVlJRlFIRnV4Ny9yakpwdVdGcUZCV3hCdkcvZUtDUUFBTHNpYUFFQUFCTlhyVmZISVlUTEVNSXJ2UUlveXRubTlPS2psdjZoNllaNHZUdWZ3ckVBMlRQUkFnQ0FuUkswQUFDQWlSb0RGdkhOM2pkNkJGQ2NkMWFHL0UxY0dYSXdzV01DOGlSb0FRREFUcjI0djc5WFlRQUFtQkFCQzREaTNZVVFqazJ6K0ZQVERYR1N4YituY2p4QUVmNVgzOWIremdJQXNCTW1XZ0FBd0VRSVdBRE13cTJWSVgvVmRNUGhlUDBEMktZNDFjTGtJQUFBZGtMUUFnQUFFcXZXcTA4UG1IN1JDNEJpeFlERmNuTjY4VmFMLzhiS0VHQVhGb0lXQUFEc2lxQUZBQUFrTWdZc3pzZC9IakFCbEVuQTRqdWFibmdkUW5nMTJRTUVjbGJwSGdBQXV5Sm9BUUFBZXlaZ0FUQUxBaFkvTUs0TVVSOWdWeFlxQ3dEQXJnaGFBQURBbmdoWUFNeUNnTVhETFYwUGdSMDZhTHJodUcvckcwVUdBR0RiQkMwQUFHQVBxdlhxM0FNbGdLSUpXRHhDMHczeFRmTmZzamxnSUZjTGszTUFBTmdGUVFzQUFOaWhhcjA2R3dNV1Irb01VQ1FCaTBleU1nVFlvMHF4QVFEWUJVRUxBQURZQVFFTGdPSUpXRHpkdWVzanNDZUNGZ0FBN0lTZ0JRQUFiSkdBQlVEeEJDeWVvZW1HK05Eekl0c1RBSEp6b21NQUFPeUNvQVVBQUd4QnRWNjlEaUZjQ2xnQUZFdkFZanZVRDlpcnBoc1dmVnRmcVRvQUFOc2thQUVBQU05UXJWZUxjWUtGdCtVQXlpUmdzU1ZOTjhTVklTK0xPQmtnSjNHU2pxQUZBQUJiSldnQkFBQlBJR0FCVUR3Qml5MXF1dUZ3dkc0QzdGdWw0Z0FBYkp1Z0JRQUFQSUtBQlVEeEJDeDJJNjdYT2lqeHhJREpXMmdSQUFEYjl1TCsvbDVSQVFEZ0J3UXNBSXIzSVlZQkJDeTJyK21HK0RiNS95M3R2SUNzL0srK3JUOXFHUUFBMjJLaUJRQUFmRWUxWGgyUEFZczM2Z1JRcE90eGdvWDkvYnR6V2VxSkFkbUlvZW4zMmdVQXdMWUlXZ0FBd0ZjSVdBQVVUOEJpRDVwdWVHMGFGREFCbGFBRkFBRGJKR2dCQUFDZkViQUFLSjZBeFg2WlpnRk13VUlYQUFEWUprRUxBQUFRc0FDWUF3R0xQV3U2SVY1WGoyWjEwc0JVVlRvREFNQTJ2YmkvdjFkUUFBQm1xMXF2RGtNSTUrTy9BNThFZ09JSVdDVFFkRU1NTUc1Y1c0RUorVmZmMWhzTkFRQmdHMHkwQUFCZ2xnUXNBSW9uWUpIVzB2VVZtSmhxRElBQkFNQ3pDVm9BQURBckFoWUF4Uk93U0t6cGhvVlZYTUFFV1I4Q0FNRFdDRm9BQURBTEFoWUF4Uk93bUk3bDNBc0FUTkpDV3dBQTJCWkJDd0FBaWxldFYyY2hoRXNCQzRBaUNWaE1TTk1OOFpwN012YzZBSlAwVWxzQUFOaVdGL2YzOTRvSkFFQ1J4b0JGZkt2MlNJY0JpaU5nTVRGTk44VHBVUnZYWFdEQ2Z1N2IyblVEQUlCbk05RUNBSURpQ0ZnQUZFM0FZcnJPWFh1QmlhdENDSzRmQUFBOG02QUZBQURGRUxBQUtKcUF4WVExM1hBY1FyaVlleDJBeVZ1TUt3VUJBT0JaQkMwQUFNaGV0VjU5K3NIVTNtV0E4Z2hZNU1HRFN5QUhsUzRCQUxBTkwrN3Y3eFVTQUlBc2pRR0xPTUhpUkFjQmlpTmdrWW1tRytMMStIL21YZ2NnRy8vczIvcEd1d0FBZUE0VExRQUF5STZBQlVEUkJDenlZNW9Ga0pNNDFVTFFBZ0NBWnhHMEFBQWdHd0lXQUVVVHNNaFEwdzFuVm5jQm1ZbEJpL2VhQmdEQWN3aGFBQUF3ZWRWNmRSeENlQ3RnQVZBa0FZdE1OZDF3YUpvRmtLR0ZwZ0VBOEZ5Q0ZnQUFUTllZc0lnVExON29Fa0J4QkN6eUY2L1JCM012QXBBZDRXMEFBSjd0eGYzOXZTb0NBREFwQWhZQVJST3dLRURURGZGYS9mL21YZ2NnVy8vcTIzcWpmUUFBUEpXSkZnQUFUSWFBQlVEUkJDeks4bmJ1QlFDeVZvVVFCQzBBQUhneVFRc0FBSktyMXF0UE85NEZMQURLSTJCUm1LWWJGa2J2QTVsYkNJd0JBUEFjZ2hZQUFDUXpCaXpPeDM5MnZBT1VSY0NpWEI1T0FybXJkQkFBZ09jUXRBQUFZTzhFTEFDS0ptQlJzS1liNG9xdm83blhBY2plU3kwRUFPQTVYdHpmM3lzZ0FBQjdJV0FCVURRQmk4STEzUkN2NHpldTRVQWhmdTdiMmpVTEFJQW5NZEVDQUlDOXFOYXJwWUFGUUpFRUxPYmowblVjS01naWhPRGFCUURBa3doYUFBQ3dVOVY2ZFJZZndCa3pEbEFjQVlzWmFicWhDaUc4bVhzZGdLSlUyZ2tBd0ZNSldnQUFzQk1DRmdERkVyQ1lwOHU1RndBb3prSkxBUUI0cWhmMzkvZUtCd0RBMWdoWUFCUkx3R0ttbW01NEhVTDRQM092QTFDa2YvWnRmYU8xQUFBOGxva1dBQUJzaFlBRlFMRitpOU1NQkN6bXFlbUdROU1zZ0lMRjlTR0NGZ0FBUEpxZ0JRQUF6MUt0VjRzeFlIR2lrZ0JGZVRkT3NQQUFhdDdPaFNpQmdzWHZNdTgxR0FDQXh4SzBBQURnU1FRc0FJb2xZTUh2bW00NEhvTVdBS1dxZEJZQWdLY1F0QUFBNEZFRUxBQ0tKV0RCbCtMMS9rQlZnSUw1VGdNQXdKTzh1TCsvVnprQUFINUl3QUtnV0FJVy9FM1REZkc2L3o4cUE4ekF2L3EyM21nMEFBQ1BZYUlGQUFEZlZhMVh4MlBBNG8xS0FSUkZ3SUx2V2FvT01CTXhXQ1pvQVFEQW93aGFBQUR3VlFJV0FNVVNzT0M3bW00NE04RUttSkZLc3dFQWVDeEJDd0FBL2tMQUFxQllBaGI4VU5NTmh5R0VTNVVDWmtUUUFnQ0FSeE8wQUFEZ2R3SVdBTVVTc09BeHprTUlCeW9Hek1qTEdETHIyL3FqcGdNQThGQ0NGZ0FBTTFldFY0ZmpRNVdMdWRjQ29EQUNGanhLMHczSDdnZUFtWXBUTGE0MEh3Q0FoeEswQUFDWXFjOENGdDVjQlNpTGdBVlBaV1VJTUZjTFFRc0FBQjVEMEFJQVlHWUVMQUNLSldEQmt6WGRFQjh5dmxKQllLWVdHZzhBd0dNSVdnQUF6SVNBQlVDeEJDellCdE1zZ0RtcmRCOEFnTWQ0Y1g5L3IyQUFBSVdyMXFzWXJsZ0tXQUFVUmNDQ3JXaTZJZDRuL0ZzMWdabjdaOS9XcnFrQUFEeUlpUllBQUFXcjFxdXpNV0J4cE04QXhSQ3dZR3VhYmpnYzd4VUE1aTZ1RDNrNzl5SUFBUEF3Z2hZQUFBVVNzQUFva29BRnUyRGlGY0FmckE4QkFPREJCQzBBQUFvaVlBRlFKQUVMZHFMcGh1TVF3aStxQy9BN1FRc0FBQjVNMEFJQW9BRFZldlZweksyQUJVQTVCQ3pZTlNQeUFmNTBvaFlBQUR6VWkvdjdlOFVDQU1qVUdMQlkrbEVRb0NnQ0Z1eGMwdzJ2UXdqL1I2VUIvdUxudnEydmxBUUFnQjh4MFFJQUlFTUNGZ0JGRXJCZ255NVZHK0J2NHZvUVFRc0FBSDVJMEFJQUlDTUNGZ0JGRXJCZ3I1cHVXRm8zQnZCVmxiSUFBUEFRZ2hZQUFCbW8xcXRxZlBOVXdBS2dIQUlXN0YzVERZY2hoSE9WQi9pcWhiSUFBUEFRZ2hZQUFCTldyVmZINHdTTE4vb0VVQXdCQzFLS3djMERIUUQ0cXFNWVNPdmIrcVB5QUFEd1BZSVdBQUFUSkdBQlVDUUJDNUpxdXFGeWJ3SHdRM0dxeFh0bEFnRGdld1F0QUFBbVJNQUNvRWdDRmt6RnBVNEEvRkFsYUFFQXdJOElXZ0FBVEVDMVhoMk9EejhFTEFES0lXREJaRFRkY0JaQ09ORVJnQjlhS0JFQUFEOGlhQUVBa05BWXNEZ2YvOW1YRGxBR0FRc21wZW1HdzNGaUZnQS9WcWtSQUFBL0ltZ0JBSkNBZ0FWQWtRUXNtS3A0djNHa093QVBjdEIwUTlXMzlVYTVBQUQ0RmtFTEFJQTlFckFBS0pLQUJaUFZkTVB4ZU44QndNUEZxUmFDRmdBQWZKT2dCUURBSGdoWUFCVG5Mb1R3Tm9Sd0tXREJ4QzNkZXdBOG12VWhBQUI4bDZBRkFNQ09WZXZWMmZpUXc4aHVnUHpGZ01YbEdMRDRxSjlNV2RNTml4RENHMDBDZUxTRmtnRUE4RDJDRmdBQU95SmdBVkFVQVF0eWRLbHJBRS95VXRrQUFQaWVGL2YzOXdvRUFMQkZBaFlBUlJHd0lFdE5OOFQ3a2YvV1BZQW4rN2x2Nnl2bEF3RGdhMHkwQUFEWUVnRUxnS0lJV0pDdHBoc09UYk1BZUxhNFBrVFFBZ0NBcnhLMEFBQjRwbXE5V293UE00eVhCY2lmZ0FVbE9BOGhIT2drd0xOVXlnY0F3TGNJV2dBQVBORVlzSWdUTEU3VUVDQjdBaFlVb2VtRzR4RENoVzRDUEp1Z0JRQUEzeVJvQVFEd1NBSVdBRVVSc0tBMGIzVVVZQ3VPWW5pdGIrc2I1UVFBNEV1Q0ZnQUFEeVJnQVZBVUFRdUswM1REd24wS3dGYkZxUmFDRmdBQS9JMmdCUURBRDFUcjFmSDRNTzZWV2dGa1Q4Q0NrcGxtQWJCZE1XanhYazBCQVBpU29BVUF3RGVNQVlzNHdlS05HZ0ZrVDhDQ29qWGRjQjdIM09zeXdGWXRsQk1BZ0s4UnRBQUErSUtBQlVCUkJDd29YdE1OaCtPOUN3RGJaUjBUQUFCZkpXZ0JBREFTc0FBb2lvQUZjeEx2WHc1MEhHRDdtbTZvK3JiZUtDMEFBSjhUdEFBQVpxOWFyejY5QmZyTDNHc0JVQUFCQzJZbFBnQjBEd093VS9IdnJLQUZBQUIvSVdnQkFNeldHTEE0SC85NUN4UWdid0lXek5XbHpsTzQ2L0ZCdC90MVVsbUVFTjZxUGdBQW54TzBBQUJtUjhBQ29DZ0NGc3hXMHcydlF3Z25QZ0VVN256OE8rK3pUaXFWeWdNQThDVkJDd0JnTmdRc0FJb2lZQUdtV1ZDK1gvdTIzalRkY0NWb1FVSXZtMjQ0N052YS9RWUFBUDhoYUFFQXpFSzFYc1Z3eFZMQUFpQjdBaGJ3eHpTTGVGOXpwQllVN0c2OGY0OWkwT0pDczBtb0dqK0hBQUR3TzBFTEFLQm8xWHAxTnY1QTYwRUVRTjRFTEdBVTM2d2VKM1JCeVphZlRSRFk2RFNKTFFRdEFBRDRuS0FGQUZBa0FRdUFZZ2hZd045ZG10SkY0Vzc3dHY3UGFwd1l1R2k2NFVOYzRhRHhKRklwUEFBQW54TzBBQUNLSW1BQlVBd0JDL2lLcGh2aVc5VnYxSWJDblgzbDlEYUNGaVMwVUh3QUFENG5hQUVBRktGYXIxNlBEK1FFTEFEeUptQUIzN2RVSHdyM1c5L1dYMXZSc0JFeUlxR0RwaHVPKzdhKzBRUUFBSUtnQlFDUXUycTlXb3dQSEU0MEV5QnJBaGJ3QTAwM25Mbm5ZUWJPdjNHS1h3dGZ3RDdGOVNHQ0ZnQUEvRTdRQWdESWtvQUZRREVFTE9BQm1tNDROTTJDR1ZoOWEySkEzOWFicGh0OEJrZ3BmZ2Q5cndNQUFBUkJDd0FnTndJV0FNVVFzSURIT2JjaWpjSjl1aTU4ejdYdkFTUlVLVDRBQUo4SVdnQUFXUkN3QUNpR2dBVThVdE1OeDk5WnB3Q2xPTy9iK2tmWGhTdmZCMGpJWnc4QWdQOFF0QUFBSnExYXI0N0hnTVViblFMSW1vQUZQRjM4ZitkQS9TalloNzZ0M3o3ZzlEWStCS1RVZEVNVjE5aG9BZ0FBZ2hZQXdDUUpXQUFVUThBQ25xSHBoampWNjVVYVVyaUhUbXp4Z0p2VUZqNkhBQUFFUVFzQVlHb0VMQUNLSVdBQjIzR3BqaFR1WGQvV1Z3ODV4YjZ0YjVwdXVBMGhIUGxRa0VpbDhBQUFCRUVMQUdBcUJDd0FpbkU3L2oxL0wyQUJ6OU4wdzFrSTRhVXlVckM3OFpyeEdCdEJDeEphS0Q0QUFFSFFBZ0JJclZxdkRzZFJ3ZWQyandOazdmZUF4ZWIwNGlFNzlvRWZhTHJoMERRTFp1QXlUcWw0NUdsZVdhZERRa2Z4NzNQZjFzS2tBQUF6SjJnQkFDUWhZQUZRREFFTDJJMmxleVFLZC92RU1OSEdCNFBFcWpId0F3REFqQWxhQUFCN0pXQUJVQXdCQzlpUnBodmlTclZmMUpmQ0xaOHlGYUJ2NjZ1bUczdzJTR2toYUFFQWdLQUZBTEFYQWhZQXhSQ3dnTjN6L3hlbHUrN2Iram1mOHc4aGhKYytKU1N5VUhnQUFBUXRBSUNkcTlhcnMzRXNzSUFGUUw0RUxHQVBtbTZJRC9CTzFKckNuVC96OUs0RUxVaW9VbndBQUFRdEFJQ2RHUU1XY2IvNGtTb0RaRXZBQXZiTC8ydVU3bDNmMXB0bm51TnovM3Q0am9PNDRxbHY2eHRWQkFDWUwwRUxBR0RyQkN3QWlpQmdBWHZXZE1PNSt5Y0tkN2VGYVJaaG5HZ0JLUzBFNHdBQTVrM1FBZ0RZR2dFTGdDSUlXRUFDVFRjY2p2ZFJVTExMdnEwL1B2Zjg0aVNCcGh2dXJDWWtJZXREQUFCbVR0QUNBSGkyYXIyS2IvTmMycE1Na0RVQkMwanIwa05qQ25mYnQvVTJ3MFJ4ZmNpSkR3MkpDRm9BQU15Y29BVUE4R1Jqd0dMcEIwNkFyQWxZUUdKTk44UUhkbS8wZ2NKdFkyWEk1NjU4RHlFaG56MEFnSmtUdEFBQUhrM0FBcUFJQWhZd0haZDZRZUd1KzdaK3YrVlRqRUdMQ3g4Y1VtbTZZZEczOVpVR0FBRE1rNkFGQVBCZ0FoWUFSUkN3Z0FscHV1RzFleXRtNEd3SHA3anh3U0d4YWd6OEFBQXdRNElXQU1BUFZldlZjUWpoclljQUFGa1RzSUJwTXMyQzB2M2F0L1hOdHMreGIrdVBUVGQ4Q0NHODlBa2lrWVcvNFFBQTh5Vm9BUUI4MHhpd1dOb1pEcEExQVF1WXFLWWI0bjNXa2Y1UXNMdngrOFN1YkFRdFNLaFNmQUNBK1JLMEFBRCtSc0FDb0FnQ0ZqQmhUVGZFKzYxelBhSnd5emg1WW9lbmVPVTdDd2tkTmQxd2Nxc2toUUFBSUFCSlJFRlV1T1BQT0FBQUV5Vm9BUUQ4aDRBRlFCRUVMQ0FQOFo3clFLOG8ySWUrclhlOVZtSGpBMFJpY1gzSWUwMEFBSmdmUVFzQUlBWXNEc2Zkc2dJV0FQa1NzSUJNTk4yd2NOL0ZET3g4WWt2ZjFwdW1HM3lXU0trU3RBQUFtQ2RCQ3dDWXNURmdjVDcrODBZbFFKNEVMQ0EvU3oyamNMLzFiWDIxcDFPOERpR2MrRUNSeUVMaEFRRG1TZEFDQUdaSXdBS2dDQUlXa0tHbUc4NDhGR1lHZGo3TjRqTlgvcDhpb1VyeEFRRG1TZEFDQUdaRXdBS2dDQUlXa0ttbUd3NU5zMkFHVm4xYjMrenhORGMrVkNSMDBIUkRGZGZZYUFJQXdMd0lXZ0RBVEZUcjFWTEFBaUJyQWhhUXYzZ3ZkcVNQRk93dWhIQzU1OVBiMTRvUytKWks0QWNBWUg1ZTNOL2ZhenNBRkt4YXI4N0dOeWY5cUErUUp3RUxLRURURGNjaGhQK25seFR1di9xMjN2djFxdW1HRzk5M1NPaGQzOVpuR2dBQU1DOG1XZ0JBb1FRc0FMSW5ZQUZsMmZkYi9yQnYxeWxDRnFPTjd6MGtWQ2srQU1EOENGb0FRR0VFTEFDeUoyQUJoV202WVJGQ2VLV3ZGRzZaOFBTdS9EOUdRaThWSHdCZ2ZnUXRBS0FRQWhZQTJST3dnSEtaWmtIcDR1cUVxNFRudVBFSkk2VVlxRXY4L3dBQUFIc21hQUVBbWF2V3E4VVlzRGpSUzRBc0NWaEF3WnB1T1BlMk00VzdTenpOSXNRSDNFMDMrSnlSMG1LY3JBSUF3RXdJV2dCQXBnUXNBTEluWUFHRmE3cmhNUFVEYU5pRHk3NnRieVpRNkd2ZmpVaW9VbndBZ0hrUnRBQ0F6QWhZQUdSUHdBTG1JOTZ6SGVnM0JidWQwR3FjamU5SUpDUm9BUUF3TTRJV0FKQUpBUXVBN0YyUEFRdGpwV0VHbW00NERpSDhvdGNVN3J4djY0OFRPY1hOQkk2QitUcUtmL2NuTXQwRkFJQTlFTFFBZ0ltcjFxdmpNV0R4UnE4QXNpUmdBZk5rYWcybHUrN2IrdjJFenRGMWx0VGlWQXRCQ3dDQW1SQzBBSUNKRXJBQXlKNkFCY3hVMHcwTFU4aVlnZk1wbldLY0pOQjB3NTExUFNRVWd4WlRDaDhCQUxCRGdoWUFNREVDRmdEWkU3QUFUTE9nZE8vNnRwN2lxbzU0N1gwMWdlTmduaGI2RGdBd0g0SVdBREFSQWhZQTJST3dBT0kwaTNnL2Q2UVNGT3h1YXRNc1ByTVJ0Q0FoazR3QUFHWkUwQUlBRXF2V3E4UHhoOG9MdlFESWtvQUY4THVtR3c0bi9BQWF0dVd5Yit1UEU2M21sZTlWcE5SMFF6WFJhUzhBQUd5Wm9BVUFKUEpad09MY0htR0FMQWxZQUYrNmRGOUg0Vzc3dGw1TzlSVDd0cjVxdW1FQ1I4S01MY2JKS2dBQUZFN1FBZ0QyVE1BQ0lIc0NGc0RmeExlWXJZQmpCczR5T01VUElZU1hFemdPNXFuU2R3Q0FlUkMwQUlBOUViQUF5SjZBQmZBOWw2cEQ0YTdqeElnTVRuRWphRUZDZ2hZQUFEUHhrMFlEd081VjYxVU1WOXlNKzRLRkxBRHlFZ01XUDI5T0x4WkNGc0RYTk4zd09vUndvamdVTG9kcEZwRnJOU205YkxyaFVBY0FBTXBub2dVQTdGQzFYc1VmSStNTzR5TjFCc2lPQ1JiQUQ0MFAxRXl6b0hTLzltMTlrOGs1Ymlad0RNeGJKZkFEQUZBK1FRc0EyQUVCQzRDc0NWZ0FqM0h1bm8vQzNZM2ZiYkxRdC9XbTZZWTdrd1JKYUNGb0FRQlFQa0VMQU5naUFRdUFyQWxZQUkvU2RNUHhHTFNBa3AzM2JmMHhzL1BiV09kRFFwWGlBd0NVVDlBQ0FMYWdXcS9pR3l0dkJTd0FzaVJnQVR6VjBsdnpGTzVEMzladk16ekZLMEVMRWxvb1BnQkErUVF0QU9BWnhvREYwbzk0QUZrU3NBQ2VyT21HZUIvNFJnVXBYSzRUV3pZVE9BYm02eUJPUE9yYitzWm5BQUNnWElJV0FQQUVBaFlBV1JPd0FMWmhxWW9VN3JlK3JYTzlWcnJHazFwY0h5Sm9BUUJRTUVFTEFIZ0VBUXVBckFsWUFGdlJkTU9aKzBGbUlOZHBGcUZ2NjQ5Tk45eGE3VWhDOGJlRDl4b0FBRkF1UVFzQWVJQnF2WXB2bzF6NlFSMGdTd0lXd05ZMDNYQTQzaGRDeVZZRnJEMjRzdDZIaENyRkJ3QW9tNkFGQUh4SHRWNGRqeE1zL0VBSGtCOEJDMkFYNGx2K0J5cEx3VzRMQ1JOdGZJOGpJUzlwQUFBVVR0QUNBTDVDd0FJZ2F3SVd3RTQwM1JEdkVTOVVsOEl0NCtxTkFrNXhNNEZqWU1hYWJsajBiZTErRkFDZ1VJSVdBUEFaQVF1QXJBbFlBTHRtWlFpbHUrN2IrbTBKNXhnZmNEZmRNSUVqWWNhcWNZVU5BQUFGRXJRQWdEOENGcDkyYlF0WUFPUkh3QUxZdWZobWNnamhsVXBUdUdWaHAzZHRoUU1KVllvUEFGQXVRUXNBWm0wTVdKemJ0UTJRSlFFTFlKK0tlTXNmdnVOZGdXc09Ob0lXSkxSUWZBQ0FjZ2xhQURCTEFoWUFXUk93QVBhcTZZWjR6M2lrNmhUc3JzQnBGbUZjMi9ETEJJNkRlVHBxdXVHd2IrdVArZzhBVUI1QkN3Qm1SY0FDSUdzQ0ZzRGV4WWRraFQ2QWhzOWQ5bTE5VTJCRk5oTTRCdWF0R2dNL0FBQVVSdEFDZ0ZrUXNBREltb0FGa05MUy9TT0Z1KzNidXNnd1VReVBOTjF3NS85aEVsb0lXZ0FBbEVuUUFvRGlWZXZWMmZnRHVYSFBBSGtSc0FDU2FycmgyTm9CWnVDODhGT005eEd2Sm5BY3pOTkMzd0VBeWlSb0FVQ3hCQ3dBc2lWZ0FVekZXNTJnY05kOVc3OHYvQnczZ2hZa1ZDaytBRUNaQkMwQUtJNkFCVUMyZm9zNzRnVXNnQ2xvdXVGMUNPRkVNeWhjNmRNc3dqalI0bUlDeDhFOEhjVHBTSEdOamY0REFKUkYwQUtBWWdoWUFHVHIzVGpCd2cvUXdKUmM2Z2FGZTllMzlhYjBrK3piK3FycGhna2NDVE8yTUNFSkFLQThnaFlBWks5YXJ4YmpEK0V2ZFJNZ0t3SVd3Q1ExM1NDOFMrbnVaakxONHBNUHZpK1NrUFVoQUFBRkVyUUFJRnRqd0dKcHBETkFkZ1FzZ01scXV1RndaZytnbWFkbDM5WWZaM1RtVjRJV0pMUlFmQUNBOGdoYUFKQWRBUXVBYkFsWUFEbUlrOUlPZElxQzNmWnRQYmZWT01XdlNHSFNoSHdBQUFva2FBRkFOZ1FzQUxJbFlBRmtvZW1HT043OWpXNVJ1TE1aTmxqUWdxU2FibGowYlgybEN3QUE1UkMwQUdEeXF2WHFlSHl6OEpWdUFXUkZ3QUxJemR6ZThtZCtydWY0c0xkdjYwM1REWGVtMVpCUU5hNndBUUNnRUlJV0FFeldHTEJZZXFzUUlEc0NGa0IybW00NE16bU5HWmpqTkl0UE52NGZKNkdGTUI4QVFGa0VMUUNZSEFFTGdHd0pXQUJaYXJyaGNMei9oSkw5MnJmMW5LL1JWNElXSkZRcFBnQkFXUVF0QUpnTUFRdUFiQWxZQUxrN0R5RWM2U0lGdXhNbStqMW9jVEdCNDJDZWptS29yMi9yai9vUEFGQUdRUXNBa3F2V3EwOXZFUDZpR3dCWkViQUFzdGQwdy9FWXRJQ1NuWHZBKy92cUVFZ3ByZzk1cndNQUFHVVF0QUFnbVRGZ2NUNytPOUFKZ0d3SVdBQWx1WFF2U3VFKzlHMzlkdTVOamtHVHBodHVUYThob1VyUUFnQ2dISUlXQU95ZGdBVkF0Z1FzZ0tJMDNSRGZMbjZscXhUT3hKWS9YVmxWU1VJTHhRY0FLSWVnQlFCN0kyQUJrQzBCQzZCVWx6cEw0WDdyMi9wS2svOWpJMmhCUXBYaUF3Q1VROUFDZ0wybzFxdWxnQVZBZGdRc2dHSTEzWEFXUW5pcHd4VHN6alNMdnhFNklhV0RwaHVxdnEwM3VnQUFrRDlCQ3dCMnFscXY0Zy9ZUzN0d0FiSWlZQUVVcmVtR1E5TXNtSUhMdnExZHl6OFRIM0EzM1RDWjQyR1dxbkd5Q2dBQW1STzBBR0FuQkN3QXNpUmdBY3lGU1d1VTdsYVk2SnV1UXdnbkV6MDJ5cmNJSWJ6Vlp3Q0EvQWxhQUxCVkFoWUFXUkt3QUdhajZZYmpFTUtGamxPNFpkL1dIelg1cXphQ0ZpUlVLVDRBUUJrRUxRRFlpbXE5ZWoyK01TVmdBWkFQQVF0Z2pyeEpUT211KzdiMk9mKzJxeERDTDFNOU9JcjNVb3NCQU1vZ2FBSEFzMVRyMVdLY1lPR05JSUI4Q0ZnQXM5UjB3OEo5S3pPdzFPVHYya3o0MkppQmVDM3EyL3BLcndFQThpWm9BY0NUQ0ZnQVpFbkFBcGc3Yi9sVHVuY2U0SDVmMzlZM1RUZmNtc1pJUW90eHNnb0FBQmtUdEFEZ1VRUXNBTElrWUFITVh0TU41eDZzVXJpN0VNSzVKai9JeHQ4REVxb1VId0FnZjRJV0FEeUlnQVZBbGdRc0FQNElXUnhhcDhBTVhQWnQvVkdqSHlRR0xWNWxjSnlVU2RBQ0FLQUFnaFlBZkZlMVhoMlBQMHEvVVNtQWJBaFlBUHhWdko4OVVCTUtkdHUzdFREUnc4VzFEUmU1SEN6Rk9XcTY0VGl1c2RGYUFJQjhDVm9BOEZVQ0ZnQlpFckFBK0VMVERmSE40Vi9VaGNKWkdmSUlmVnRmTmQyUXpmRlNwSGh0Y3M4T0FKQXhRUXNBL2tMQUFpQkxBaFlBMzNhcE5oVHV1bS9yOTVyOGFCOUNDQzh6TzJiS0VkZXordjhXQUNCamdoWUEvRTdBQWlCTEFoWUEzOUYwdytzUXdva2FVVGpUTEo3bVN0Q0NoQ3JGQndESW02QUZ3TXhWNjlYaCtNUGN1YjNWQU5rUXNBQjRHTk1zS04ydmZWdHZkUGxKMUkyVWhBQUJBREluYUFFd1V3SVdBRmtTc0FCNG9LWWI0clMySS9XaVlIZmpWRUtlNWtyZFNLbnBoa3BRQ2dBZ1g0SVdBRE1qWUFHUW5idHhmN09BQmNBRE5kMXdhSjBDTTdEczIvcWpSajlOMzlZM1RUZmMrVjVNUWd1VFZRQUE4aVZvQVRBVEFoWUEyYmtiUjk1ZmJrNHZQRVFCZUp4TDk3d1U3clp2YTZ0eG5tOWpoUU1KVllvUEFKQXZRUXVBR2FqV3F6TS9OZ05rUThBQzRCbWFib2h2Q0w5UlF3cDNwc0ZiY1NWb1FVS0NGZ0FBR1JPMEFDallHTEN3bXhvZ0R3SVdBTnV4VkVjS2Q5MjM5WlVtYjBXczQwVUI1MEdlWHNaVlYxWUFBUURrU2RBQ29FQUNGZ0JaRWJBQTJKS21HODY4bmM0TW1HYXhQWnRTVG9Sc1ZXUGdCd0NBekFoYUFCUkV3QUlnS3dJV0FGc1Uzd28yellJWldQVnRmYVBSMnhFbkNUVGQ4Q0ZPRmlqaGZNalNRdEFDQUNCUGdoWUFCYWpXcS9qRi9LMkFCVUFXQkN3QWR1UGMvVENGKzNRUHdYWnRCQzFJcUZKOEFJQThDVm9BWkd3TVdDeU5Sd2JJZ29BRndJNDAzWEFjUXJoUVh3cDNIaWN3YVBMV3hhREZtOExPaVh3czlBb0FJRStDRmdBWkVyQUF5SXFBQmNEdWVjdWYwbjNvMi9xdEx1K0V0UTJrZEJERGdsWUNBUURrUjlBQ0lDTUNGZ0JaRWJBQTJJT21HK0k5OGl1MXBuRG5HcndiZlZ0dm1tNG84ZFRJeDZkMXNBQUFaRVRRQWlBRDFYcDFQSDdwRnJBQW1ENEJDNEQ5TXMyQzByM3IyOXJVaGQyNjluMmJoQ3JGQndESWo2QUZ3SVNOQVl1bGZiRUFXUkN3QU5penBodk9RZ2d2MVoyQzNZM2ZDZG10SzBFTEVoSzBBQURJa0tBRndBUUpXQUJrUmNBQ0lJR21HdzVOczJBR0x2dTJ2dEhvbmRzVWZuNU1tNUFQQUVDR0JDMEFKa1RBQWlBckFoWUFhY1g3NWdNOW9HQzN3a1I3STJoQlVrMDNMS3dJQWdESWk2QUZ3QVJVNjlXbnQvRUVMQUNtVDhBQ0lMR21HMkpBK1JkOW9IREx2cTNkYSt4Qm5CclNkRU1NdGh3VmY3Sk1WVFd1c0FFQUlCT0NGZ0FKalFHTDgvR2Z0L0VBcGszQUFtQTYzdW9GaGJ2dTI5cm5mTDgyZ2hZa1ZDaytBRUJlQkMwQUVoblhoR3dFTEFBbVQ4QUNZRUxpZUhYNzdKbUJjMDNldXpoTjROWE16cG5wV09nRkFFQmVmdEl2Z0RRMnB4YzNJWVFiNVFlWXJCaXdXSVVRamplbkYwc2hDNERKOEpZL3BYdlh0L1ZHbC9kT3pVbnBxT21HUXgwQUFNaUhvQVZBV240a0JwZ2VBUXVBaVdxNjRkeG9md3AzWjVwRkduMWJYODN4dkprVTYwTUFBREppZFFoQVd1OURDUC9XQTRCSnNDSUVZTUxHTjMyWGVrVGhMdnUyZGgrU3pyWFZSQ1MwR0ZmWUFBQ1FBUk10QUJJYTE0ZDgwQU9BcEV5d0FNaERETU1kNkJVRnUrM2JXcGdvTGV0RFNHbWgrZ0FBK1REUkFpQzl0NlphQUNSaGdnVkFKcHB1aU9QVTMrZ1hoYk15SkQxQkMxS3lPZ1FBSUNNbVdnQ2s5MTRQQVBiS0JBdUEvRnpxR1lXNzd0dmFkOFAwckcwZ3BZTXhXQWdBUUFZRUxRQVNzejRFWUc4RUxBQXkxSFRENnhEQ2lkNVJ1RE1OVHE5djY1dnhuaEZTRWJRQUFNaUUxU0VBMDJCOUNNRHVXQkVDa0RmVExDamRyK01EZnFZaFRyVjRwUmNrSW1nQkFKQUpRUXVBYVhndmFBR3dkUUlXQUpscnVtRVpRampTUndvVzcxZVdHandwRzBFTEVsb29QZ0JBSHF3T0FaaUFjWDNJdFY0QWJJVVZJUUFGYUxyaE9JUndycGNVYnRtM3RYdVZhYm1hZXdGSTZxWHlBd0RrUWRBQ1lEcmU2Z1hBc3doWUFKUWx2dVYvb0tjVTdMWnZhNnR4cG1jejl3S1FWdE1OcGxvQUFHUkEwQUpnT3Q3ckJjQ1RDRmdBRkdaOHlQUkdYeW5jbVFaUHp6aGg1TVBjNjBCU2xmSURBRXlmb0FYQVJJd1BCbi9URDRBSHV3MGgvRzhCQzRBaUxiV1Z3djNXdDdVVkZkTmxxZ1VwbVdnQkFKQ0JmMmdTd0tURXFSYXZ0QVRndTJMQUlnWXJyRndDS0ZEVERmRXQveE85cFhEbkdqeHBWNmJxa0pDSkZnQUFHVERSQW1CYXJBOEIrTFlZc1BpdnplbkZzWkFGUUptYWJqZ01JVnhxTDRWYjlXMTlvOG1UWnFJRktSMk4xME1BQUNaTTBBSmdRcXdQQWZncUFRdUErWWh2K1Ivb053VzdFeWFhdnI2dEJTMUl6Zm9RQUlDSnN6b0VZSHFzRHdINGd4VWhBRFBTZE1OeENPRkN6eW5jZWQvV0h6VTVDOWZXR0pGUVplb3BBTUMwQ1ZvQVRFLzhJdjNmK2dMTW1JQUZ3RHg1eTUvU1hmZHQ3ZjRtSDFlQ0ZpUmtvZ1VBd01SWkhRSXdNZGFIQURObVJRakFURFhkc0REVmpSbFlhbkpXckE4aEpTRWZBSUNKRTdRQW1DWVBHSUU1RWJBQXdEUUxTdmV1YitzclhjNktmcEZVMHcyVkRnQUFUSmVnQmNBRWJVNHY0dnFRTzcwQkNpZGdBVUI4a0hRZVFuaXBFaFRzempTTC9QUnQvWEc4WDRWVUJDMEFBQ2JzSDVvRE1Ga3hiUEZHZTRBQ3hSK3NsOElWQURUZGNPZ0JORE53MmJmMWpVWm5LYTRQT1pwN0VVaG1ZZUlwQU1CMENWb0FUSmVnQlZBYUFRc0F2aFJERmdlcVFzRnVyY2JKV2x3ZjhtcnVSU0FaRXkwQUFDYk02aENBaWJJK0JDaUlGU0VBL0UzVERjY2hoRjlVaHNLZGp5c295Tk5HMzBqSVdpMEFnQWtUdEFDWXR2ZjZBMlJNd0FLQTczRnRvSFRYZlZ2N1RwZXh2cTJ2NWw0RDBtcTZZYUVGQUFEVFpIVUl3TFJaSHdMa3lJb1FBTDVyZkhCMG9rb1U3bHlEaTNEdDd4VUpMY1lWTmdBQVRJeUpGZ0FUWm4wSWtCa1RMQUI0S05jSlN2ZXViMnRySjhxZ2o2UlVxVDRBd0RTWmFBRXdmYVphQUZObmdnVUFEOVowd3pLRWNLUmlGT3pPTkl1aXhHa0N2OHk5Q0NRamFBRUFNRkVtV2dCTW53ZVh3RlNaWUFIQW96VGRjT2dCTkROdzJiZjFSNDB1aG9rV3BIVFVkTU94RGdBQVRJK2dCY0RFYlU0dnJzYUhtUUJUSVdBQndGTmRoaEFPVkkrQzNmWnR2ZFRnY3ZSdGZXT2xKNG1aYWdFQU1FR0NGZ0I1ZUs5UHdBUUlXQUR3WkUwM1ZGYmlNUU5ubWx5a3E3a1hnS1FXeWc4QU1EMy8wQk9BTEx5MUV4WklLQVlzbHNJVkFEelRwUUpTdU91K3JUMlFMMU5jSC9KcTdrVWdHUk10QUFBbXlFUUxnQXhzVGk4MjFvY0FDWmhnQWNCV05OM3dPb1J3b3BvVXpqU0xjZ25Ra0pMckp3REFCQWxhQU9URCtoQmdYd1FzQU5pYXBoc09UYk5nQm43dDIvcEdvOHRrVWdtcGpldTNBQUNZRUt0REFQSmhmUWl3YTFhRUFMQUw1eUdFSTVXbFlIZnhIa3FEaS9jaGhQQnk3a1VnbWNXNHdnWUFnSWt3MFFJZ0U5YUhBRHRrZ2dVQU85RjB3L0VZdElDU0xmdTIvcWpEeGZPUW01Uk10QUFBbUJnVExRRHk4dDVVQzJDTFRMQUFZTmZpVy80SHFrekJQdlJ0YlRYT1BNVDFJVy9tWGdTU0ViUUFBSmdZRXkwQTh1SUhQR0FiVExBQVlPZWFibGg0S01rTW1OZ3lIeVpha05MTHBoc09kUUFBWURvRUxRQXlzam05dUJuM3dnSThoWUFGQVBza0pFenBmdXZiK2txWDU2RnY2eGkwdUp0N0hVaktWQXNBZ0FteE9nUWdQL0hoNkwvMURYZ0VLMElBMkt1bUc4N2kyN2VxVHVGTXM1aWZHTFk0bVhzUlNHWXhyckFCQUdBQ0JDMEE4dk5lMEFKNElBRUxBUFp1SEcxdW1nV2xXL1Z0ZmFQTHMzTWxhRUZDQzhVSEFKZ09xME1BTW1OOUNQQUExeUdFLzgrS0VBQVNpVy81SHlnK0Jic1RKcHF0emR3TFFGSldod0FBVElpZ0JVQ2VQRGdGdmlZR0xIN2VuRjRzTnFjWDcxVUlnSDFydXVFNGhIQ2g4QlR1dkcvcmo1bzhTOVkya05MQmVKMEZBR0FDQkMwQTh1UUJLdkM1endNV2Z2d0ZJQ1Z2K1ZPNjY3NnRCZDluYWd6WTNNNjlEaVJsZlFnQXdFUUlXZ0JreVBvUVlDUmdBY0JrTk4wUUgvNjgwaEVLdDlUZzJYUGZUVXJXaHdBQVRNUS9OQUlnVy9FdHFuOXJIOHhTREZnc2hTc0FtQmh2K1ZPNmQzMWJ1LzlpRTBKNE0vc3FrSXFnQlFEQVJBaGFBT1RydmFBRnpJNkFCUUNUMUhURGVRamhTSGNvMkoxcEZvemNpNVBTaWVvREFFekRpL3Y3ZTYwQXlGUzFYcjAzbmhsbVFjQUNnTWxxdXVFd2hCQlgyeDNvRWdWYjlXMHRhTUh2bW03d2d5b3AvV3k2RGdCQWVpWmFBT1JOMEFMS0ptQUJRQTZXUWhZVTdsYklnaTljbXl4QVFwWEpLZ0FBNlFsYUFPUXRCaTMrV3craE9BSVdBR1NoNlliakVNSXZ1a1hoempXWUwyd0VMVWlvVW53QWdQUiswZ09BZkcxT0x6NkdFSDdUUWloR0RGajh2RG05V0FoWkFKQ0p0eHBGNGE3N3RuNnZ5WHpCdlRvcExWUWZBQ0E5UVF1QS9QblJEL0luWUFGQWRwcHVlTzJOYm1iQU5BdStacU1xSkhUVWRNT2hCZ0FBcENWb0FaQS9RUXZJbDRBRkFEbTcxRDBLOTY1dmF3L1UrWnUrclc5Q0NMY3FRMEttV2dBQUpDWm9BWkE1NjBNZ1N3SVdBR1N0NllabGZLTldGeW5ZbldrVy9JQVFEaWxWcWc4QWtOWS8xQitnQ0hHcXhTdXRoTW1MQVl1bGNBVUFPUnZIbFhzQVRlbVdmVnQvMUdXK1krTjdPQW1aYUFFQWtKaWdCVUFaWXREaXYvVVNKa3ZBQW9DU3hKVWhCenBLd1c3N3RyWWFoeCtKOS9ZWHFrUWlKbG9BQUNSbWRRaEFBY2IxSWUvMEVpYkhpaEFBaXRKMFEzeUQ5bzJ1VXJnekRlWkgrcloyZjA5S0IwMDNDRnNBQUNSa29nVkFPZDc3MFJzbXd3UUxBRXExMUZrS2QrMEJPby93SVlUd1VzRklwQnBYMkFBQWtJQ0pGZ0NGMkp4ZXhLREZuWDVDVWlaWUFGQ3NwaHZpVy80bk9remhUTFBnTWR6ems1S0pGZ0FBQ1psb0FWQVdVeTBnRFJNc0FDaGEwdzJIcGxrd0E3LzJiWDJqMFR5Q2FRS2t0RkI5QUlCMEJDMEF5aUpvQWZzbFlBSEFYSnlIRUk1MG00TGRDUlB4QklJV3BHUnREUUJBUWkvdTcrL1ZINkFnMVhyMU1ZUndvS2V3VXdJV0FNeEcwdzNINDhORTk1aVU3TC82dG42cnd6eFcwdzIrZzVQU3ozMWIrMTRLQUpDQWlSWUE1VEhWQW5aSHdBS0FPYnIwRUpIQ2ZSQ3k0QmxpRU8xRUFVbWtDaUg0ZmdvQWtNQlBpZzVRblBkYUNsc1hBeFkvYjA0dkZrSVdBTXhKMHcxeC8vc3JUYWR3NXhyTU0vaCtRRW9MMVFjQVNFUFFBcUF3bTlPTDkrTitZZUQ1QkN3QW1Mdkx1UmVBNHYxbTdEN1A1UE5EU3BYcUF3Q2tJV2dCVUNaVExlQjVCQ3dBbUwybUc4NUNDQy9uWGdlS1o1b0Z6N1ZSUVJJNmFycmhXQU1BQVBaUDBBS2dUTjQ4aEtjUnNBQ0FQMElXaCs0cG1ZRlYzOVkzR3MxejlHMzlNWVJ3cTRna1pLb0ZBRUFDZ2hZQUJkcWNYbXo4MEFPUEltQUJBSDhWMy9JL1VCTUtkaXRNeEJiNURrRktnaFlBQUFuOFE5RUJpaFhYaC95aXZmQmRNV0N4Rks0QWdEK05JOGd2bElUQ0xjZEpCTEFOOFdXSE55cEpJZ3VGQndEWVAwRUxnSEs5RmJTQWJ4S3dBSUJ2ZTZzMkZPNjZiMnVmYzdiSjl3cFNPbEY5QUlEOWUzRi9mNi9zQUlXcTFxdTRiL2hJZitFLzNvMEJDN3U0QWVBcm1tNkliOFgrajlwUXVKLzd0dlpnbksxcXVzR1ByS1Qwcjc2dE56b0FBTEEvSmxvQWxNMzZFUGlEZ0FVQVBJeTMvQ25kT3lFTGR1VGFaQUVTcXNZVk5nQUE3TWxQQ2cxUU5EK1VNM2N4WVBIUHplbkZtWkFGQUh4ZjB3M25wcUZSdUxzUXdya21zeU1DUEtTMFVIMEFnUDBTdEFBbzJPYjBJcjdOY0t2SHpKQ0FCUUE4UXRNTmgzSDZrNXBSdU11K3JUOXFNanRpbWdBcFZhb1BBTEJmVm9jQWxNLzZFT2JFaWhBQWVKb1lzamhRT3dwMjI3ZTFNQkc3SkdoQlNpOVZId0JndjB5MEFDamZwUjR6QXlaWUFNQVROZDFRQ2VZeUExYUdzRk45VzkrWUtFbEtUVGRZSHdJQXNFZUNGZ0NGR3g4NmY5Qm5DaVZnQVFEUEo1aEw2YTc3dG42dnkreUJxUmFrSkdnQkFMQkhWb2NBek1QYkVNSy85WnFDV0JFQ0FGdlFkTVByRU1LSldsSTQweXpZbDZzUXdpdlZKcEZLNFFFQTlrZlFBbUFlM2d0YVVBZ0JDd0RZTHRNc0tOMnZmVnViTXNDKytLeVJrb2tXQUFCNzlPTCsvbDY5QVdhZ1dxL2lEejR2OVpwTUNWZ0F3SlkxM2JBTUlWeW9Ld1c3Q3lFYzkyMzlVWlBabDZZYi9OaEtTdi9zMjlyM1pnQ0FQVERSQW1BK3JBOGhSd0lXQUxBRFRUY2NXNmZBREN5RkxFamdnNWNjU0NpdUQvSDlHUUJnRHdRdEFPYkQraEJ5SW1BQkFMc1ZwMWtjcURFRnUrM2IybW9jVXJnU3RDQ2h4Zmo3RHdBQU8vYVRBZ1BNdy9qQStvTjJNM0V4WVBIUHplbkZtWkFGQU94RzB3M3hJY3diNWFWd1p4cE1JaHVGSjZGSzhRRUE5c05FQzRCNXNUNkVxVExCQWdEMlo2bldGTzY2YitzclRTWVJuejFTT2xGOUFJRDlNTkVDWUY2TWoyUnFUTEFBZ0QxcXV1SE1ReGhtd0RRTGt1bmJPbjZ2dWRNQlVtbTZ3VlFMQUlBOUVMUUFtSkh4UWZadmVzNUVmQkN3QUlEOWFicmgwRFFMWm1BMVB1aUdsS3dQSWFXRjZnTUE3SjZnQmNEOG1HckJWTHlzMXF0ajNRQ0F2VGtQSVJ3cE53V0xVd1F1TlpnSnNENkVsRXkwQUFEWUEwRUxnUGtSdEdCS1h1c0dBT3hlMHcweDNIaWgxQlR1dkcvcmo1ck1CQWhha0pLSkZnQUFleUJvQVRBem05T0xqOWFITUNIMlp3UEFmbmpMbjlKOTZOdjZyUzR6RVZhSGtOTFJ1QzRNQUlBZEVyUUFtQ2RUTFpnSzYwTUFZTWVhYm9odnRyNVNad3AzcnNGTXhUaFo1WU9Ha0pEMUlRQUFPeVpvQVRCUGdoWk1pZlVoQUxCYnBsbFF1dC82dHJhcWdha3gxWUtVckE4QkFOZ3hRUXVBR2JJK2hJbng5aUVBN0VqVERYRk4xMHYxcFdCMzdpZVpLRUVMVWhLMEFBRFlNVUVMZ1BreTFZS3BPS3JXSzJOTkFXREx4djNzcGxsUXVzdStyVzkwbVFreVpZV1VmTWNHQU5neFFRdUErWG8vdnYwRlUzQ21Dd0N3ZGNzUXdvR3lVckJiWVNLbXFtOXJFeTFJNmFEcGhtTWRBQURZSFVFTGdKa2ExNGVZYXNGVXZOWUpBTmllOGVIS0wwcEs0Wlo5VzMvVVpDYnNXbk5JeVBvUUFJQWRFclFBbURkQkM2YkMraEFBMks2MzZrbmhydnUyOWpsbjZxd1BJU1hmc1FFQWRralFBbURHTnFjWDFvY3dKZGFIQU1BV05OMFEzMkE5VVVzS3Q5UmdNbUI5Q0NrSldnQUE3SkNnQlFDbVdqQVYxb2NBd0haNHk1L1N2ZXZiMnFRQWN1QnpTa3BDbHdBQU95Um9BWUNnQlZNUjE0ZllJUXNBejlCMHczbThwcW9oQllzVCtjNDFtQnowYmYweGhIQ3JXYVF5VHJrQ0FHQUhCQzBBWnM3NkVDYkcraEFBZUtLbUd3NnRVMkFHTHNlSDE1QUw2ME5JeWZvUUFJQWRFYlFBSUpocXdZUllId0lBVDNjWlFqaFFQd3AyMjdlMU1CRzVzVDZFbEV5MEFBRFlFVUVMQUlLZ0JSTnlVSzFYd2hZQThFaE5OOFEzVnQrb0c0V3pNb1FjbVdoQlNpWmFBQURzaUtBRkFOYUhNRFdDRmdEd2VKZHFSdUd1KzdZV0VDYzdmVnViYUVGS1IrTnFNUUFBdGt6UUFvQlAzcW9FRXlGb0FRQ1AwSFJEdkhhZXFCbUZPOU5nTW5hdGVTUmtmUWdBd0E0SVdnRHdpYUFGVTJGOUNBQTgwUGlXcW1rV2xPN1h2cTF2ZEptTVdSOUNTdGFIQUFEc2dLQUZBTC9ibkY3RUgzNXVWWU9KRUxRQWdJYzVqMlBCMVlxQ3hSV0hTdzBtYzRJV3BHU2lCUURBRGdoYUFQQTVPNCtaaWpmVmVtV1BMQUI4UjlNTngyUFFBa3EyN052Nm93NlR1U3NOSkNFVExRQUFka0RRQW9EUFdSL0NsSmhxQVFEZkY5L3lQMUFqQ25iYnQ3WFZPR1J2WEgxenA1TWtjdEIwZzdBRkFNQ1dDVm9BOEIvV2h6QXhnaFlBOEExTk44UXg0Ry9VaDhLZGFUQUZNZFdDbEFRdEFBQzJUTkFDZ0M5Wkg4SlV2TEkrQkFDK2FhazBGTzYzdnEwOW1LWWtHOTBrSVVFTEFJQXRFN1NvTzZVMEFBQWdBRWxFUVZRQTRFdldoekFscGxvQXdCZWFib2h2K1orb0M0VTcxMkFLSXpoRVNndlZCd0RZTGtFTEFQN2kvMmZ2Ym00aXliSTJBSWNIaVFmZ0FYZ0FxOWhXU1NHRjJJRUhuUjZRZUpEdFFiRkRJWVZVdFkxVjRVSGlBWGhBZXZBcFp1S2I2WjZ1SDM0eTg1NjQ5M2xNT0llaWdIenZlZFdIRUl5Z0JRRDhSZE1ONDdXbnRabVF1ZHUrclo4c21jeTRhRUZLcDZZUEFMQmJnaFlBL0lpckZrUXgxb2VjMkFZQS9NZjR5bjloSEdSc0sweEVqdnEyZnFtcTZ0RnlTYVhwQmxjdEFBQjJTTkFDZ0I4UnRDQVNWeTBBNE44ZmtJemh3eHV6SUhQTDZRTnB5SkdyRnFRa2FBRUFzRU9DRmdEOHcrYnk1c2xMR3dLNXRnd0ErQmV2L01uZFk5L1dRdC9rN0x2dGt0Q1o0UU1BN0k2Z0JRQS80dytjUkhHcVBnU0EwazNudmorVlBnZXl0N1JpTXVlaUJTa0pXZ0FBN0pDZ0JRQS84OVZrQ0VSOUNBQ2xjODJDM04zMWJlMjFQMW5yMjNvTVdteHRtVVNPcHhveUFBQjJRTkFDZ0I5U0gwSXc2a01BS0ZiVERlTXIvMU5mQVdScy9PQjVaY0VVd2xVTFVuTFZBZ0JnUndRdEFQZ1Y5U0ZFb1Q0RWdDSTEzWERrQTJnS3NPN2Irc21pS1lUTExhUWthQUVBc0NPQ0ZnRDhpdm9RSW5IVkFvQVNqU0dMaGMyVHNXZlZPQlRHUlF0U3VqQjlBSURkRUxRQTRLZlVoeENNb0FVQVJabDYxUCt3ZFRLMzZ0djZ4WklwaUlzV3BIUnUrZ0FBdXlGb0FjRHZxQThoaXVPeisxdG5UZ0VvaVovRHlOMUQzOWEremluS0ZDeDZ0blZTYWJyQjc5VUFBRHNnYUFIQTcvakRKNUc0YWdGQUVacHV1UERxbEFJc0xabEN1V3BCU29JV0FBQTdJR2dCd0M5dExtL0cxemJmVElrZ1Bsc0VBSVVRZGlWM2QzMWJiMnlaUXZuYUo2VUwwd2NBK0RoQkN3QmU0NnNwRVlUNkVBQ3kxM1REYXZ3L3o2YkoyTlkxQ3dvbmFFRktmcWNHQU5nQlFRc0FYa1BRZ2tqVWh3Q1FyYVliam53QVRRSFdmVnUvV0RTbDZ0dGFkUWdwblU0L2J3QUE4QUdDRmdEOGx2b1FnbEVmQWtETzFsVlZMV3lZakQzM2JiMnlZS2dlaklDRVhMVUFBUGdnUVFzQVhzdFZDNklZNjBPRUxRRElUdE1ONDRjZVZ6Wkw1bHduZzM5VEgwSktGNllQQVBBeGdoWUF2SmFnQlpFSVdnQ1FvN1d0a3JrSGxRbndILzR0a0pLTEZnQUFIeVJvQWNDcnFBOGhHRUVMQUxMU2RNUDR5di9jVnNtY2F4YndYeTVha0pLTEZnQUFIeVJvQWNCYnVHcEJGQXYxSVFEa291bUdvNnFxVmhaSzV2N3MyL3JKa3VIZnBuOFBXK01na1VYVERTZUdEd0R3Zm9JV0FMekZWMzhJSWhCQkN3QnlzYXlxNnRnMnlkaFdtQWgrU0gwSUtha1BBUUQ0QUVFTEFGNXRxZzl4MVlJb0JDMEFtTDNwTmVuU0pzbmNxbS9yRjB1R2YxQWZRa3JxUXdBQVBrRFFBb0MzRXJRZ0N2VWhBT1JnZk9XL3NFa3k5dGkzOWRxQzRZZGN0Q0FsRnkwQUFENUEwQUtBTjlsYzNxZ1BJWkpyMndCZ3JwcHVHRitTWGxrZ21YT3hCWDZpYjJ0QkMxSTZOMzBBZ1BjVHRBRGdQVnkxSUlwUFovZTNSN1lCd0V4NTVVL3V2dmtnR1g3cjBZaElwZWtHVnkwQUFONUowQUtBOXhDMElCTDFJUURNVHRNTjQxV21VNXNqYzY1WndPOXR6SWlFTGd3ZkFPQjlCQzBBZURQMUlRUWphQUhBckRUZGNPU2FCUVc0N2R2NnlhTGh0MXg5SVNVWExRQUEza25RQW9EM2N0V0NLTlNIQURBMzR5di9oYTJSc2Ewd0VieWFpeGFrNUtJRkFNQTdDVm9BOEY2Q0ZrVGlxZ1VBczlCMHcwbFZWVGUyUmVhV2ZWdS9XREw4WHQvV0d4Y2pTZWg0dXJRRkFNQWJDVm9BOEM1VGZjaXo2UkhFdFVVQU1CTmUrWk83aDc2dHY5Z3l2SW1yRnFTa1BnUUE0QjBFTFFENENGY3RpT0w4N1A3MnhEWUFpS3pwaHZFODl5ZExJbk1yQzRZMysyNWtKS1ErQkFEZ0hRUXRBUGdJTDlXSVJIMElBTkg1MlluYzNmVnQ3UU5qZUR2L2JraEowQUlBNEIwRUxRQjR0ODNselVaOUNJR29Ed0VncktZYmxtTVB1ZzJSc2ExckZ2QnVxa05JU1hVSUFNQTdDRm9BOEZIcVE0amlWSDBJQUJFMTNYRGtBMmdLc083YitzbWk0ZTM2dG43eGlJR0VGazAzK0YwYUFPQ05CQzBBK0NnbnNJbEVmUWdBRVkwaGk0WE5rTEh4QStLMUJjT0hxQThoSmZVaEFBQnZKR2dCd0llb0R5RVk5U0VBaE5KMHczaU8rdzliSVhQTDZVVSs4SDdxUTBoSmZRZ0F3QnNKV2dDd0MrcERpRUo5Q0FEUmVPVlA3aDc2dHZiN0FIeWNpeGFrNUtJRkFNQWJDVm9Bc0F2cVE0aGthUnNBUk5CMHcxaHBkVzRaWk03UFhyQURmVnU3YUVGS3A2WVBBUEEyZ2hZQWZKajZFSUw1YkNFQUJPR2FCYm03OCtFdzdOU0RjWkpLMHcydVdnQUF2SUdnQlFDNzRvTUVvamcrdTcvVkx3dEFVazAzck1iL2syeUJqRzFkczRDZEUxd2lKYjlIQXdDOGdhQUZBTHVpbDVsSXJtMERnRlNhYmpqeUFUUUZXUFZ0L1dMUnNGUGZqWk9FWExRQUFIZ0RRUXNBZG1KemVmTlVWZFdqYVJLRStoQUFVaG92ZlMxc2dJdzk5MjN0b2gzc25vc1dwT1NpQlFEQUd3aGFBTEJMWDB5VElOU0hBSkRFMUc5K1pmcGt6dlV3MklPK3JjY0hETTltU3lMSDAxVXVBQUJlUWRBQ2dGMVNIMElrUGdBQUlJV1ZxWk81aDc2dDFSdkEvcmhxUVVycVF3QUFYa25RQW9DZFVSOUNNSUlXQUJ4VTB3M2ovejNucGs3bS9Jd0YreVZvUVVvdVF3SUF2SktnQlFDN3BqNkVLQlpuOTdlZmJRT0FRNWhPYmJ0bVFlNytuS29OZ1AxeE1ZYVVYTFFBQUhnbFFRc0FkazE5Q0pFSVdnQndLTXV4Mjl5MHlkaFdtQWoyVHpVUGlibG9BUUR3U29JV0FPeVUraENDRWJRQVlPK2FiamlaZ2hhUXMyWGYxaTgyREFmaGQycFNXVFRkSUd3QkFQQUtnaFlBN01QYVZBbENmUWdBaHpEKzdMTXdhVEwyMkxlMWlrQTRIRmN0U0VuUUFnRGdGUVF0QU5nSDlTRkVJbWdCd040MDNUQjJtWDh5WVRMbllnc2Mxc2E4U2VqQzhBRUFmay9RQW9DZDIxemVqQ2VGdjVrc1FYdyt1Nzg5c2d3QTlzUWxMM0wzclc5cnIrdmhzUHliSXlVWExRQUFYa0hRQW9COWNkV0NLQmF1V2dDd0QwMDNYRmRWZFdxNFpNNDFDeml3dnEyZnFxcmFtanVKK05rR0FPQVZCQzBBMkJkQkN5SVJ0QUJncDVwdU9ITE5nZ0xjVGgvNEFvZW5Qb1JrcG1vMEFBQitRZEFDZ0wxUUgwSXduOVNIQUxCankrbHFFdVRxV1pnSWtsSWZRa3FDRmdBQXZ5Rm9BY0ErdVdwQkpLNWFBTEFUVFRlY1ZGVjFZNXBrYnRXMzlZc2xRektDRnFSMFp2b0FBTDhtYUFIQVBnbGFFSW1nQlFDNzhzVWt5ZHhEMzlhK3ppRXQxU0drSkdnQkFQQWJnaFlBN0kzNkVJSlJId0xBaDAyZDVlY21TZVpXRmd4cFRSZGxIcTJCUkk2bkMxNEFBUHlFb0FVQSsrYXFCWkc0YWdIQVIzbmxUKzd1K3JaV1dRQXh1R3BCU3E1YUFBRDhncUFGQUh1MXVid1pQNHpZbWpKQkxDMENnUGRxdW1IOGYrVFlBTW5ZMWpVTENFWFFncFFFTFFBQWZrSFFBb0JEY05XQ0tFN1A3bStkUHdYZ3pacHVPUElCTkFWWTkyMzlaTkVRaHVzeXBIUmgrZ0FBUHlkb0FjQWhDRm9RaWZvUUFONWpYVlhWd3VUSTJIUGYxc0pFRUVqZjFpNWFrTks1NlFNQS9KeWdCUUI3dDdtOCthbytoRUN1TFFPQXQyaTZZVHlkZldWb1pFN0ZHc1QwWUMra012ME1CQURBRHdoYUFIQW9ybG9RaGZvUUFONXFiV0prN3FGdmF6K3ZRMHpxUTBoSmZRZ0F3RThJV2dCd0tQNXdTeVRxUXdCNGxhWWJQanVkVFFGY3M0QzQxSWVRa29zV0FBQS9JV2dCd0VHb0R5RVk5U0VBdkpackZ1VHVybTlySCtSQ1hQNTlrcEtnQlFEQVR3aGFBSEJJcmxvUXhWZ2Y0ZzlHQVB4UzB3MnJxcXFPVFltTWJWMnpnTmo2dG42cXF1clpta2prdE9tR0k4TUhBUGduUVFzQURrblFna2hjdFFEZ3A1cHVPUEVCTkFWWTlXMzlZdEVRbnFzV3BPU1JBZ0RBRHdoYUFIQXdVMzJJbHpoRThka21BUGlGOFpyRndvREkySFBmMXFweFlCNisyeE1KWFJnK0FNQS9DVm9BY0dpdVdoREZzZm9RQUg2azZZYnhBNFVyd3lGenJudkJmTGhvUVVwK2J3WUErQUZCQ3dBTzdZdUpFNGdQR0FENGtaV3BrTG1IdnEyOWtJZVo4TytWeEZ5MEFBRDRBVUVMQUE1cWMzbXpVUjlDSU9wREFQaWJwaHZHRU42NXFaQTVZVk9ZbjBjN0k1RkYwdzBuaGc4QThIZUNGZ0Nrb0Q2RUtNYjZFSzl6QVBpWHBodU9YTE9nQUxkOVd6OVpOTXlPcXhha3BENEVBT0IvQ0ZvQWtJTDZFQ0x4b2hPQS83Y2NRM2ltUWNhMlZWV3RMUmhtYVdOdEpPU0JBZ0RBL3hDMEFPRGcxSWNRalBvUUFLcnBKUGFOU1pDNVpkL1dMNVlNcytTaUJTbTVhQUVBOEQ4RUxRQklSWDBJVVN6TzdtK0ZMUUR3eXAvY1BmWnQ3Ykljek5SVStiTzFQeEk1TjNnQWdMOFR0QUFnRlgva0pSSkJDNENDTmQwd25zUCs1R3VBekMwdEdHYlBWUXVTbVg1ZUFnQmdJbWdCUUJKVGZjaWo2Uk9Fb0FWQTJWeXpJSGZmK3JiMkFTM00zOFlPU1VoOUNBREFYd2hhQUpDU3F4WkVvVDRFb0ZCTk4xeFhWWFZxLzJSczY1b0ZaRU5naXBRRUxRQUEva0xRQW9DVXZwbytnUWhhQUJTbTZZWWoxeXdvd0xwdjZ5ZUxoaXk0YUVGS3FrTUFBUDVDMEFLQVpEYVhOMC9xUXdqazZ1eis5c2hDQUlxeUdxOGFXVGtaZXhZbWduejBiZjNpZDJnU09wNUNxZ0FBeGFzRUxRQUlRSDBJa2JocUFWQ0lwaHRPcXFyNnc3N0ozR3I2WUJiSWg2c1dwS1ErQkFCZ0ltZ0JRR3JxUTRoRTBBS2dITUtlNU82aGIydGY1NUNmNzNaS1F1cERBQUFtZ2hZQUpLVStoR0ErcVE4QnlGL1REZU9IQk9kV1RlWldGZ3haY3RHQ2xBUXRBQUFtZ2hZQVJPQ2xIWkc0YWdHUVB6OTdrTHU3dnEyOWVvY005VzB0YUVGS3FrTUFBQ2FDRmdCRW9ENkVTQVF0QURMV2RNUDR5di9ZanNuWXRxcXFwUVZEMWg2c2wwUVdUVGVjR0Q0QWdLQUZBQUZNOVNIK1VFUVU2a01BTXRWMHc1RVBvQ25BdW0vckY0dUdyTGxZUTBycVF3Q0E0bFdDRmdBRTRvUTNrVnpiQmtDVzF1TkxUS3NsWTg5OVc2OHNHTEtuUG9TVTFJY0FBTVdyQkMwQUNFUjlDSkVJV2dCa3B1bUc4VU9CSzNzbGN5NjJRQmxjdENBbEZ5MEFnT0pWZ2hZQVJMRzV2Qm5QRzMrekVJSTRQYnUvMVRzTGtKZTFmWks1aDc2dGhaZWhBRk05MExOZGs4aXB3UU1BQ0ZvQUVJcy9EQlBKWjlzQXlFUFREZVAzOUhQckpIT3VXVUJaMUllUVROTU5ybG9BQU1VVHRBQWdFa0VMSWxFZkFwQ0JwaHVPWExPZ0FILzJiZTFEVnlpTCtoQlNPak45QUtCMGdoWUFoS0UraEdEVWh3RGtZWHpsZjJ5WFpHeGJWZFhLZ3FFNHdsV2s1S0lGQUZBOFFRc0FvbkhWZ2tqVWh3RE1XTk1OSitvVUtNQ3FiK3NYaTRheTlHM3RvZ1VwdVdnQkFCUlAwQUtBYUFRdGlNU0hjd0R6TnI3eVg5Z2hHWHZ1MjFvMURwVHJ3ZTVKNUhpcVp3TUFLSmFnQlFDaHFBOGhtT096KzFzdmRRQm1xT21HOGFUMWxkMlJ1V3NMaHFLcER5RWw5U0VBUU5FRUxRQ0k2SXV0RUlnUE1BRG1hV1Z2Wk82YjZnQW9ucUFGS1htVUFBQVVUZEFDZ0hBMmx6ZGpmY2pXWmdqaXMwVUF6RXZURFdOSTd0emF5SnlLTTBEWWlwUmN0QUFBaWlab0FVQlVYMjJHSU5TSEFNekkxQmUrdGpNeWQ5dTM5Wk1sUTltbTd3TWVLWkNLVUNzQVVEUkJDd0NpRXJRZ0V2VWhBUE14dnZKZjJCY1oyd29UQVgvaHFnWEpOTjNnVVFJQVVDeEJDd0JDVWg5Q01PcERBR2FnNllhVHFxcHU3SXJNTGZ1MmZyRmtZTEl4Q0JJU3RBQUFpaVZvQVVCa3Jsb1F4VmdmSW13QkVKOVgvdVR1c1cvckw3WU0vSVdMRnFSMFlmb0FRS2tFTFFDSVROQ0NTQVF0QUFKcnVtSDhRLzhuT3lKelN3c0cvcXB2YTBFTFVuTFJBZ0FvbHFBRkFHR3BEeUVZUVF1QTJGeXpJSGQzUGxBRmZ1TFJZRWprMU9BQmdGSUpXZ0FRbmFzV1JMRlFId0lRVTlNTlMzL29KM05qK0hobHljQlBiQXlHVkthcllnQUF4UkcwQUNBNkhkUkVJbWdCRUV6VERVYytnS1lBNjc2dG55d2ErQW5YYmtoSjBBSUFLSktnQlFDaGJTNXZ4ajhZUGRzU1FRaGFBTVF6aGl3VzlrTEdubFhqQUwvaG9nVXBuWmsrQUZBaVFRc0E1a0I5Q0ZHb0R3RUlwT21HazZxcS9yQVRNcmZxMi9yRmtvR2Y2ZHQ2TTFVTVFRcUNGZ0JBa1FRdEFKZ0Q5U0ZFY20wYkFHSDRHWUhjUGZSdDdlc2NlQTFYTFVqbGVBcS9BZ0FVUmRBQ2dQQTJsemNiOVNFRTh1bnMvdmJJUWdEU2FycGh2REIwYmcxa2JtbkJ3Q3Q5TnlnU2N0VUNBQ2lPb0FVQWM2RStoRWpVaHdDa3Q3WURNbmMzMVFFQXZJYnZGNlIwWWZvQVFHa0VMUUNZQ3llVGlVVFFBaUNocGh0VzQ1bHFPeUJqVzljc2dEZHkwWUtVWExRQUFJb2phQUhBTEtnUElSajFJUUNKTk4xdzVBTm9DckR1Mi9yRm9vSFhtcjVuK0oyWlZOUzVBUURGRWJRQVlFN1VoeENKcXhZQWFZeVZJUXV6SjJQUGZWdXZMQmg0QjFjdFNLYnBCbGN0QUlDaUNGb0FNQ2ZxUTRoRTBBTGd3S1kvNEYrWk81bHpzUVY0cjQzSmtkQ0Y0UU1BSlJHMEFHQTJwdnFRUnhzamlMRSs1TVF5QUE1cWJkeGs3cUZ2YTFmY2dQY1N0Q0FsRnkwQWdLSUlXZ0F3TjY1YUVJbXJGZ0FIMG5URHRmNXZDbkJ0eWNCNzlXMnRPb1NVQkMwQWdLSUlXZ0F3TjE3NEVZa1BRd0FPb09tR282cXFWbVpONXY3czIvckprb0VQZWpCQUVqbWRmbVlEQUNpQ29BVUFzN0s1dkhsU0gwSWdwK3BEQUE1aVdWWFZzVkdUc2Ewd0ViQWo2a05JeVZVTEFLQVlnaFlBekpINkVDSlJId0t3UjAwM25FeEJDOGpacW0vckZ4c0dka0I5Q0NsZG1ENEFVQXBCQ3dEbVNIMElrYWdQQWRpdjhaWC93b3pKMkdQZjFtc0xCbmJFUlF0U2N0RUNBQ2lHb0FVQXM2TStoR0RVaHdEc1NkTU40NnZJSy9NbGN5NjJBRHZUdC9YNCsvS3ppWktJaXhZQVFERUVMUUNZSy9VaFJPSURFb0Q5OE1xZjNIM3IyOXFaZjJEWFhMVWdsY1ZVK3dZQWtEMUJDd0RtU24wSWtYeTJEWURkYXJwaHJHWTZOVll5SjZ3SjdJT2dCU201YWdFQUZFSFFBb0JabXVwRHZ0a2VRUnlmM2QvcW9nWFlrYVliamx5em9BQzMwNGwvZ0YxektZZVUvRzRNQUJSQjBBS0FPWFBWZ2tpdWJRTmdaOFpYL2d2akpHTmJZU0pnWDFRU2taaWdCUUJRQkVFTEFPWk0wSUpJMUljQTdNRFU2MzFqbG1SdTJiZjFpeVVEZS9Sb3VDUnlidkFBUUFrRUxRQ1lyYzNsell2NkVBSlJId0t3RzEvTWtjdzk5RzN0Nnh6WU4xY3RTS2JwaGd2VEJ3QnlKMmdCd055NWFrRWs2a01BUG1ENm83eFhrT1J1WmNQQUFXd01tWVE4UWdBQXNpZG9BY0RjQ1ZvUWlhQUZ3TWQ0NVUvdTd2cTI5c29jT0FSQkMxSVN0QUFBc2lkb0FjQ3NxUThobU1YWi9lMW5Td0Y0dTZZYmxtTU5rOUdSc2ExckZzQ2g5RzI5bWI3dlFBcXFRd0NBN0FsYUFKQURWeTJJUk5BQzRJMmFiamp5QVRRRldQZHQvV1RSd0FHNWFrRXF4OVBQZHdBQTJSSzBBQ0FIZ2haRUltZ0I4SFpqeUdKaGJtVHNlUXhhV0RCd1lLcUtTRWw5Q0FDUU5VRUxBR1p2cWcrNXMwbUNVQjhDOEFaTk40eC9oUC9Eek1qY3NtL3JGMHNHRGt6UWdwVFVod0FBV1JPMEFDQVhybG9RaWFBRndPdDU1VS91SHZxMjlyTXFrSUxxRUZJU3RBQUFzaVpvQVVBV05wYzM0eCt2dDdaSkVJSVdBSy9RZE1QNC9mTGNyTWpjMG9LQkZLWkxPcytHVHlLcVF3Q0FyQWxhQUpBVEx3V0pZcXdQdWJZTmdOOXl6WUxjM2ZWdDdVVTVrSkw2RUZKWlRCVnhBQUJaRXJRQUlDZUNGa1RpcWdYQUx6VGRzS3FxNnRpTXlOaldOUXNnQUdFdlVoSzBBQUN5SldnQlFEYlVoeERNcDdQNzJ5TkxBZmlucGh1T2ZBQk5BZGJUMlg2QWxGeTBJQ1ZCQ3dBZ1c0SVdBT1RHVlFzaWNkVUM0TWZHeXBDRjJaQ3g1NzZ0VnhZTXBLYStpTVF1TEFBQXlKV2dCUUM1RWJRZ0VrRUxnUC9SZE1QNEIvY3JjeUZ6MXhZTUJQSmdHU1J5YXZBQVFLNEVMUURJaXZvUWdsRWZBdkJQWHZtVHU0ZStyWjNxQnlKeDFZSmtwcEF0QUVCMkJDMEF5SkdyRmtUaXFnWEFwT21HOFpYL3VYbVFPZGNzZ0dpRXYwanB6UFFCZ0J3SldnQ1FvN1d0RXNqU01nRCtGYkk0Y3MyQ0F2elp0L1dUUlFQQnVHaEJTaTVhQUFCWkVyUUFJRHVieTV2eGowalBOa3NRcDJmM3R5ZVdBZkN2NE5teE1aQ3hyVEFSRU5FVUFQTTdNcW00YUFFQVpFblFBb0JjcVE4aEV2VWhRTkdhYmpoeDRZY0NyUHEyZnJGb0lDaFhMVWpsZVBwWkVBQWdLNElXQU9UcWk4MFNpSzUyb0hSanJkZWk5Q0dRdGNlK3JkWFhBWkY5dHgwU2N0VUNBTWlPb0FVQVdWSWZRakRxUTRCaU5kMHc5bkovOGhWQTVseHNBYUp6MFlLVUJDMEFnT3dJV2dDUU0vVWhSS0krQkNpVlYvN2s3bHZmMWw2S0E2SDVQa1ZpRnhZQUFPUkcwQUtBbktrUElSTDFJVUJ4bW00WXYvZWQyanlaYzgwQ21JdEhteUtSYzRNSEFISWphQUZBdHRTSEVNeFlIK0pjS2xDTXBodU9YTE9nQUxkOVd6OVpOREFUcmxxUVROTU5maDhHQUxJaWFBRkE3dFNIRUltckZrQkpWbFZWTFd5Y2pEMExFd0V6czdFd0VoSzBBQUN5SW1nQlFPNzg4WnRJUHRzR1VJS21HMDZxcXZyRHNzbmNxbS9yRjBzR1pzUkZDMUs2TUgwQUlDZUNGZ0JrYlhONTg2U0hsa0NPMVljQWhmaGkwV1R1b1c5clgrZkFyRXhWUjF0Ykl4Ry9Dd01BV1JHMEFLQUUvZ2hPSk9wRGdLdzEzVEMrVmp5M1pUSzNzbUJncHRTSGtNcXB5UU1BT1JHMEFLQUVYMjJaUU5TSEFMa1RjQ1IzZDMxYk83OFB6Slh2WHlRekJYSUJBTElnYUFGQTl0U0hFSXo2RUNCYlRUY3N4Kzl6Tmt6R3RxNVpBRE1uYUVGS2doWUFRRFlFTFFBb2hkZTFSTEswRFNBM1RUY2MrUUNhQXF6N3RuNnlhR0RHVkllUWtrY0hBRUEyQkMwQUtJWDZFQ0pSSHdMa2FGMVYxY0pteWRoejM5YkNSTUNzOVczOTR1SWpDYmxvQVFCa1E5QUNnQ0tvRHlHWXhkbjlyYkFGa0kybUc4YlhpVmMyU3VaY3BBSnk0YW9GcVN5YWJqZ3hmUUFnQjRJV0FKUkVmUWlSQ0ZvQU9WbmJKcGw3Nk52YWhUUWdGNElXcEtRK0JBRElncUFGQUNYeHgzRWlFYlFBc3RCMHcvajk3TncyeVp4ckZrQk92dHNtQ2FrUEFRQ3lJR2dCUURHbStwQnZOazRRNmtPQVhMaG1RZTd1K3JiMitodklodTlwSk9haUJRQ1FCVUVMQUVyanFnV1JDRm9BczlaMHc2cXFxbU5iSkdOYjF5eUFURDFZTEltNGhBWUFaRUhRQW9EU0NGb1F5ZFhaL2UyUmpRQnoxSFREaVErZ0tjQ3FiK3NYaXdZeXBENkVaSnB1Y05VQ0FKZzlRUXNBaXJLNXZIbFJIMEl3cmxvQWN6VmVzMWpZSGhsNzd0dGFOUTZRSy9VaHBIUmgrZ0RBM0FsYUFGQWlWeTJJUk5BQ21KMm1HOFkvamwvWkhKbTd0bUFnWXk1YWtKS0xGZ0RBN0FsYUFGQWlRUXNpK2FRK0JKaWhsYVdSdVllK3JYMElDV1JycWtWNnRtRVNFYlFBQUdaUDBBS0E0cWdQSVNCWExZRFphTHBoZk9WL2JtTmt6alVMb0FUcVEwamx0T2tHRHc0QWdGa1R0QUNnVks1YUVJbWdCVEFMMHgvRVhiTWdkMy8yYmYxa3kwQUJYTzRoSlZjdEFJQlpFN1FBb0ZTQ0ZrU2lQZ1NZaTJWVlZjZTJSY2Eyd2tSQVFWeTBJS1VMMHdjQTVrelFBb0FpVGZVaGQ3WlBJRTZVQTZFMTNYQlNWZFdOTFpHNVpkL1dMNVlNbEtCdmF4Y3RTRW5RQWdDWU5VRUxBRXJtcWdXUkNGb0EwYTF0aU13OTltMzl4WktCd2p4WU9JbW9EZ0VBWmszUUFvQmliUzV2dms3bm9TR0MwN1A3MnhPYkFDSnF1bUY4Y2ZqSmNzamMwb0tCQXFrUElaWEZkREVOQUdDV0JDMEFLSjJyRmtUeTJUYUFvRnl6SUhmZm5OQUhDaVZvUVVycVF3Q0EyUkswQUtCMGdoWkVvajRFQ0tmcGh2R1YvNm5Oa0xHdGF4WkF3WVRNU0VsOUNBQXdXNElXQUJSTmZRakJxQThCUW1tNjRhaXFxcFd0a0xsMTM5WlBsZ3lVYVByKzUzZGlVaEcwQUFCbVM5QUNBRnkxSUJiMUlVQWtZOGhpWVNOazdGazFEb0NyRmlSemJ2UUF3RndKV2dDQW9BV3hPRjBPaE5CMHczaGg1dy9iSUhPcnZxMWZMQmtvM0tiMEFaQk8wdzBYeGc4QXpKR2dCUURGVXg5Q01NZG45N2ZPcHdJUmZMRUZNdmZRdDdXdmN3QVhMVWpMNzc4QXdDd0pXZ0RBdi9rak81RmMyd2FRMHZTeTBDbG5jcmV5WVlCL2NkR0NsQVF0QUlCWkVyUUFnSDhUdENDU3o3WUJKT2IvUlhKMzE3ZTFGOXdBVlZWTkZVcVBaa0VpcWtNQWdGa1N0QUNBZjllSGpDOTRuczJDSU5TSEFNazAzVEMrOGorMkFUSTJWc1l0TFJqZ2IxeTFJSlhqcGh1T1RCOEFtQnRCQ3dENHI2OW1RU0RxUTRDRG0vN0k3UU5vY3JlZVhtOEQ4Rit1L0pDU3F4WUF3T3dJV2dEQWZ6bVRUaVRxUTRBVTFsVlZMVXllakQzM2JiMnlZSUIvY05HQ2xGeDBCQUJtUjlBQ0FDYnFRd2htckEveHFnYzRtS1lieGo5d1g1azRtWE94QmVBSCtyWVd0Q0Fsdi9zQ0FMTWphQUVBZjZjK2hFalVod0NIdERadE12ZlF0N1dmOVFCKzdzRnNTTVJGQ3dCZ2RnUXRBT0R2MUljUWlmb1E0Q0NhYmhpLzM1eWJOcGx6elFMZzE3NmJENGtzcHV0cUFBQ3pJV2dCQUgraFBvUmdGbWYzdDhJV3dGNDEzWERrbWdVRitOTlpmSURmOG4yU2xBUXRBSUJaRWJRQWdIOXkxWUpJQkMyQWZSdGYrUitiTWhuYlZsVzFzbUNBMzNMUmdwUUVMUUNBV1JHMEFJQi9FclFnRWtFTFlHK2FiamhScDBBQlZuMWJ2MWcwd0s5TjN5dGRlQ1NWQzVNSEFPWkUwQUlBL3NmbTh1YXBxcXBIY3lFSTlTSEFQbzJ2L0JjbVRNYWUrN1pXalFQd2VxNWFrTXFweVFNQWN5Sm9BUUEvNXFvRmtWemJCckJyVFRlTXJ3YXZESmJNK1Q4VTRHMDI1a1VxMDgrbkFBQ3pJR2dCQUQvMjFWd0k1TlBaL2UyUmhRQTd0akpRTXZmUXQ3V1gyUUJ2STJoQlNtZW1Ed0RNaGFBRkFQeUEraEFDVWg4QzdFelREZU1yLzNNVEpYT3VXUUM4a1lBYWlibG9BUURNaHFBRkFQeWMraEFpRWJRQWRxTHBodkZDenRvMHlkeHQzOVpQbGd6d0xnL0dSaUl1V2dBQXN5Rm9BUUEvcHo2RVNOU0hBTHV5cktwcVlacGtiQ3RNQlBBaDZrTkk1YmpwaGhQVEJ3RG1RTkFDQUg1Q2ZRZ0J1V29CZk1qMGgrc2JVeVJ6eTc2dFh5d1o0TjNVaDVDU3F4WUF3Q3dJV2dEQXI2a1BJUkpCQytDanZQSW5kNDk5Vy92NURlQmpYTFFnSlVFTEFHQVdCQzBBNE5mOG9aNUl4dm9RWjFTQmQybTY0V0w4UG1KNlpHNXB3UUFmMDdmMTAxVERCQ2xjbURvQU1BZUNGZ0R3QzV2TG0vSHM5RGN6SWhCWExZRDNFaDRrZDNkOVd6dDNEN0FidnArU3lybkpBd0J6SUdnQkFMLzMxWXdJNU5veWdMZHF1bUY4NVg5c2NHUnNmSG05c21DQW5WRWZRakpOTjZnUEFRRENFN1FBZ044VHRDQ1NVL1Vod0ZzMDNYRGtBMmdLc0o1TzNRT3dHeTVha0pLZ0JRQVFucUFGQVB5RytoQUNVaDhDdk1VWXNsaVlHQmw3SG9NV0ZneXdPNnFZU096Q0FnQ0E2QVF0QU9CMVhMVWdFdlVod0tzMDNUQmV3UG5EdE1qY3FtL3JGMHNHMkxsSEl5VVJGeTBBZ1BBRUxRRGdkUVF0aUVSOUNQQmFYMHlLekQzMGJlM3JIR0EvTnVaS0lxZFQvUjBBUUZpQ0ZnRHdDdXBEQ01oVkMrQ1htbTRZYTRiT1RZbk1MUzBZWUcvVWg1Q1NxeFlBUUdpQ0ZnRHdlcTVhRUltZ0JmQTdheE1pYzNkOVczdHREYkEvdnNlUzBvWHBBd0NSQ1ZvQXdPdU5RWXV0ZVJIRThkbjlyUmMrd0E4MTNiQWF2MCtZRGhuYnVtWUJzRjlUbU0zdndLVGk5MTBBSURSQkN3QjRwYWsreEZVTEluSFZBdmlIcWMvYUI5RGtidDIzOVlzdEEreWRxeGFrNHFJRkFCQ2FvQVVBdkkyZ0JaRjh0ZzNnQjhiS2tJWEJrTEhudnExWEZneHdFTitObVVRV1RUZWNHRDRBRUpXZ0JRQzh3ZWJ5Um4wSWthZ1BBZjZtNllieGU4S1ZxWkE1RjFzQURzZEZDMUx5K3k0QUVKYWdCUUM4bmFzV1JLSStCUGlydFdtUXVXSXdybjRBQUNBQVNVUkJWSWUrcmYwc0JuQTRMbHFRa3ZvUUFDQXNRUXNBZUR0LzNDY1NRUXZnWDVwdUdMOGZuSnNHbWZQL0hzQUI5VzM5TWxZMm1UbUp1R2dCQUlRbGFBRUFiNlEraEdBV1ovZTNueTBGeXRaMHcxRlZWYXZTNTBEMi91emIrc21hQVE3T1ZRdFNFU0lHQU1JU3RBQ0E5M0hWZ2tnRUxZQmxWVlhIeFUrQm5HMkZpUUNTMlJnOXFUVGQ0S29GQUJDU29BVUF2SStnQlpFSVdrREJtbTQ0bVlJV2tMUFZkTDRlZ01OejBZS1VMa3dmQUloSTBBSUEza0Y5Q01Hb0Q0R3lqYS84RjZVUGdhdzk5MjI5dG1LQU5QcTJkdEdDbEZ5MEFBQkNFclFBZ1BmN1luWUVJbWdCQldxNllYemhkMlgzWk83YWdnR1NlN0FDRW5IUkFnQUlTZEFDQU41UDBJSklCQzJnVEY3NWs3dHZmVnM3V1ErUW5xc1dwSExjZE1PUjZRTUEwUWhhQU1BN2JTNXZ4ajgwUFpzZlFZejFJVjc4UWtHYWJoai96Wi9hT1psYldqQkFDRUp2cEtRK0JBQUlSOUFDQUQ3bXEva1JpS3NXVUlqcFZaOXJGdVR1dG0vckoxc0dDTUZGQzFKU0h3SUFoQ05vQVFBZm96NkVTRDZkM2Q4NnFRcGxHRi81TCt5YWpHMkZpUURpbUlKdkxqcVNpcUFGQUJDT29BVUFmSUQ2RUFKeTFRSXkxM1REU1ZWVk4vWk01cFo5Vzc5WU1rQW9ybHFRaXVvUUFDQWNRUXNBK0RqMUlVUWlhQUg1YzAySjNEMzJiZTNySENBZVFRdFNXVXhoWXdDQU1BUXRBT0RqZkJCQUpPcERJR05OTjR4bms4L3RtTXd0TFJnZ3BPL1dRa0xxUXdDQVVBUXRBT0NEMUljUWtLc1drQy9oUG5KMzE3ZTFEL0lBQXZMOW1jVFVod0FBb1FoYUFNQnVyTTJSUUs0dEEvTFRkTVA0eXYvWWFzbll0cXFxbFFVRGhQWm9QU1FpYUFFQWhDSm9BUUM3OGRVY0NlVDg3UDVXZnkxa3BPbUdJeDlBVTRCMTM5WlBGZzBRbXFzV3BLSStEd0FJUmRBQ0FIWmdjM256NUdVUHdhZ1BnYnlNSVl1Rm5aS3haeGZDQUdaaFkwMmswblREaGVFREFGRUlXZ0RBN3VqTkp4TDFJWkNKcGh2R004bC8yQ2VaVy9adC9XTEpBT0VKV3BDUytoQUFJQXhCQ3dEWUhmVWhSSEtxUGdTeTRaVS91WHZvMjlyUFVRQXowTGYxR0xUWTJoV0p1R2dCQUlRaGFBRUFPNkkraElEVWg4RE1OZDN3V1I4MUJWaGFNc0NzdUdwQktpNWFBQUJoQ0ZvQXdHNnBEeUVTOVNFd2Y2NVprTHU3NlhVMEFQUHgzYTVJNUxqcGhpUERCd0FpRUxRQWdOMXk5cHBJeHZvUUwzNWdwcHB1V0kxL1RMWS9NcloxelFKZ2xnUXRTRWw5Q0FBUWdxQUZBT3lRK2hBQ2N0VUNabWg2cWVjRGFISzM3dHY2eFpZQlpzY2xJbEx5bUFBQUNFSFFBZ0IyejVsM0l2bHNHekJMNC84bEM2c2pZODk5VzY4c0dHQitwcENjQndhazRxSUZBQkNDb0FVQTdKNzZFQ0k1Vmg4Qzg5SjB3L2pINHl0ckkzTXVMZ0hNbTZzV3BPTDNXd0FnQkVFTEFOaXh6ZVhOK0xybm03a1NpQSt6WUY2ODhpZDNEMzFiNi9jSG1EZEJDMUpaTk4wZ2JBRUFKQ2RvQVFENzRhb0ZrYWdQZ1psb3VtRU1ScDNiRjVrVEFBU1lQNEU1VWhLMEFBQ1NFN1FBZ1AwUXRDQVM5U0V3QTAwM0hMbG1RUUgrN052NnlhSUI1cTF2YXhjdFNNbnZ0d0JBY29JV0FMQUg2a01JYUdrcEVONzQ3L1RZbXNqWVZwZ0lJQ3NQMWtraUZ3WVBBS1FtYUFFQSsrT3FCWkdvRDRIQW1tNDRxYXJxeG83STNLcHY2eGRMQnNpRytoQlNPVFY1QUNBMVFRc0EyQjlCQ3lKWm5OM2ZDbHRBWEd1N0lYT1BmVnY3T2dmSWkvb1FrbW02d1ZVTEFDQXBRUXNBMkJQMUlRUWthQUVCVFg4ay9tUTNaRTZGRlVCK0JDMUlTZEFDQUVoSzBBSUE5c3RWQ3lJUnRJQ1l2UEluZDkvNnRuWmVIaUF6ZlZzL1ZWWDFiSzhrY21id0FFQktnaFlBc0VlYnk1c3ZWVlZ0elpnZzFJZEFNRTAzWE91WXBnQ3VXUURreTFVTFVoRzBBQUNTRXJRQWdQMXoxWUpJQkMwZ2lLWWJqbHl6b0FDMzA0dG5BUExrWWhHcEhEZmRjR0w2QUVBcWdoWUFzSCtDRmtUeStleis5c2hHSUlUVmVHbkdLc2pZVnBnSUlIc3VXcENTcXhZQVFES0NGZ0N3WjV2TG02L3FRd2hrNGFvRnBEZTl2dnZES3NqY3NtL3JGMHNHeUZmZjFpNWFrSktnQlFDUWpLQUZBQnlHcXhaRUltZ0I2WDJ4QXpMMzBMZTFyM09BTWp6YU00bGNHRHdBa0lxZ0JRQWNocUFGa1h4U0h3THBOTjB3L2tINDNBckkzTXFDQVlyaHFnV3ArSmthQUVoRzBBSUFEa0I5Q0FHNWFnSHBlT1ZQN3U2Y2tnY295c2E2U2FYcEJ2VWhBRUFTZ2hZQWNEaXVXaENKb0FVazBIVERzcXFxWTdNblkxdlhMQUNLSTF4SFN1cERBSUFrQkMwQTRIQUVMWWhFZlFnY1dOTU5SejZBcGdEcnZxMmZMQnFnSE5QM2ZSY2NTY1ZGQ3dBZ0NVRUxBRGlRcVQ3azJid0p4RlVMT0t4MVZWVUxNeWRqejMxYkN4TUJsRWw5Q0trSVdnQUFTUWhhQU1CaHVXcEJKRXZiZ01PWXVxT3ZqSnZNK1g4Rm9GenFRMGpsZExvY0J3QndVSUlXQUhCWVg4eWJRRTdQN205UExBUU9ZbTNNWk82aGIydUJVb0J5Q1ZxUWtxc1dBTURCQ1ZvQXdBRnRMbTgyNmtNSVJuMEk3Rm5URGVPL3MzTnpKbk91V1FDVVRYVUlLVjJZUGdCd2FJSVdBSEI0WG5zU3liVnR3TjY1WmtIdTd2cTI5Z0ViUU1INnRuNnBxdXJSMXdDSnVHZ0JBQnljb0FVQUhKNzZFQ0pSSHdKNzFIVERxcXFxWXpNbVkxdlhMQUNZQ04yUmlvc1dBTURCQ1ZvQXdJR3BEeUVnOVNHd0IwMDNuUGdBbWdLc3BsZk1BUEM5K0FtUXltTDYyUnNBNEdBRUxRQWdEZlVoUk9LRFlOaVA4WnJGd216SjJIUGYxcXB4QVBoL0xscVFrdm9RQU9DZ0JDMEFJQTMxSVVSeWZIWi82NDlTc0VOTk40em5pNi9NbE14ZFd6QUEvNjl2YTBFTFVsSWZBZ0FjbEtBRkFDU2dQb1NBZkZnR3U3VXlUekwzMExlMUUvRUEvSzhIRXlFUmp3Y0FnSU1TdEFDQWRKemFKcExQdGdHNzBYVERHRnc2TjA0eUo2QUh3SThJNFpHS243OEJnSU1TdEFDQWRMNmFQWUdvRDRFZGFMcmhTSkNPQXZ6WnQvV1RSUVB3QStwRFNHYXE3d01BT0FoQkN3QklaSE41TTM1QThXaitCT0oxTW56Y3NxcXFoVG1Tc2ExcUhBQit3VVVMVXZKNEFBQTRHRUVMQUVqcmkva1RpUG9RK0lDbUcwNnFxcm94UXpLMzdOdjZ4WklCK0pIcC80aG53eUVSUVFzQTRHQUVMUUFnTGZVaFJETFdoemkxQ3Urbk1vVGNQZlp0TFNRS3dPK29EeUVWdjg4Q0FBY2phQUVBQ2FrUElTRDFJZkFPVXgvMEo3TWpjMHNMQnVBVjFJZVF5bkhURFVlbUR3QWNncUFGQUtUblpTaVJxQStCOTNITmd0eDk2OXZhQjJjQXZJYUxGcVNrUGdRQU9BaEJDd0JJVDMwSWtTek83bStGTGVBTm1tNFlYL21mbWhtWmM4MENnRmNSekNNeDlTRUF3RUVJV2dCQVlsTjl5SU05RUlpZ0JielNkSnA0WlY1azdyWnY2eWRMQnVBTi9JNUxLb0lXQU1CQkNGb0FRQXpxUTRoRTBBSmVid3haTE15TGpEMnJ4Z0hnSGRTSGtJcnFFQURnSUFRdEFDQUc5U0ZFb2o0RVhxSHBocE9xcXY0d0t6SzM2dHY2eFpJQmVDTkJDMUpaVEQrbkF3RHNsYUFGQUFTd3Vid1pQOEQ0WmhjRUltZ0J2K2NhRWJsNzZOdmExemtBNy9IZDFFaElmUWdBc0hlQ0ZnQVFoNnNXUkhKMWRuOTdaQ1B3WTAwM2pIKzhQVGNlTXJleVlBRGVvMi9ycDZxcXRvWkhJdXBEQUlDOUU3UUFnRGdFTFlqR1ZRdjRPYS84eWQxZDM5WmVJd1B3RWY0ZklSVVhMUUNBdlJPMEFJQWcxSWNRa0tBRi9FRFREZU1yLzJPeklXUGpDK1NsQlFQd1FSc0RKSkZUZ3djQTlrM1FBZ0JpY2RXQ1NENnBENEcvYTdyaHlBZlFGR0RkdC9XTFJRUHdRUzVha014VTlRY0FzRGVDRmdBUWk2QUYwYmhxQVgrM3JxcHFZU1prN0xsdjY1VUZBL0JSS3FoSTdNd0NBSUI5RXJRQWdFRFVoeENRb0FWTW1tNFkvMWg3WlI1a3pzVVdBSGJwMFRSSnhFVUxBR0N2QkMwQUlKNHZka0lnWTMzSWlZWEF2NnlOZ2N3OTlHM3R1aFlBdTdReFRSSngwUUlBMkN0QkN3QUlabk41TTM3QXNiVVhBbkhWZ3VJMTNURCtPemd2ZlE1a3p6VUxBSFpOZlFpcEhEZmRjR1Q2QU1DK0NGb0FRRXhla3hMSnRXMVFzdWtQdEs1WmtMcy8rN2IyNmhpQVhmTi9DeW1wRHdFQTlrYlFBZ0JpRXJRZ2tsUDFJUlJ1Zk9WL1hQb1F5TnA0U1d0bHhRRHMyaFRpYzdHUlZOU0hBQUI3STJnQkFBR3BEeUVnOVNFVXFlbUdFM1VLRkdEVnQvV0xSUU93SjY1YWtJcUxGZ0RBM2doYUFFQmNybG9RaWZvUVNqVys4bC9ZUGhsNzd0dGFOUTRBKy9UZGRFbkVSUXNBWUc4RUxRQWdMa0VMSWxFZlFuR2FiaGhmd0YzWlBKa1RwQU5nMzF5MElKVkYwdzNDRmdEQVhnaGFBRUJRNmtNSXlJZHhsTVlyZjNMMzBMZTFWOFlBN0p2L2EwaEowQUlBMkF0QkN3Q0l6VlVMSWhHMG9CaE5ONHhmNzZjMlR1WjhYd2RnNy9xMmZobXJxa3lhUkM0TUhnRFlCMEVMQUlqdGkvMFF5UEhaL2EzWFFHU3Y2WVlqMXl3b3dHM2YxazhXRGNDQnVHcEJLbjZIQlFEMlF0QUNBQUxiWE41ODkvS0hZTHgrcGdUTHNjL1pwc25ZVnBnSWdBUGJHRGlKdUZJSEFPeUZvQVVBeEtjK2hFZysyd1k1YTdyaHBLcXFHMHNtYzh2cGpEc0FISXFnQmNrMDNhQStCQURZT1VFTEFJaFBmUWlScUE4aGQxNzVrN3ZIdnEzOWJBSEFRZlZ0clRxRWxBUXRBSUNkRTdRQWdPQTJsemNiOVNFRW96NkVMRTB2M1Q3WkxwbGJXakFBaVR3WVBJbDRMQUFBN0p5Z0JRRE1nL29RSWxFZlFxNjg4aWQzMzd3b0JpQWg5U0drSW1nQkFPeWNvQVVBeklNUC80aGtyQThSdGlBclRUZU1yL3lQYlpXTWJWMnpBQ0F4WVQ5U09XNjY0Y1QwQVlCZEVyUUFnQmxRSDBKQWdoWmtvK21HbzZxcVZqWks1dFo5V3o5Wk1nQUp1V2hCU3E1YUFBQTdKV2dCQVBPaFBvUklCQzNJeVJpeVdOZ29HUnZEbW1zTEJpQ2xLZkMzdFFRU0ViUUFBSFpLMEFJQTVrTjlDSkVzMUllUWcrbUU4QitXU2VaV2ZWdS9XRElBQWFnUElaVUxrd2NBZGtuUUFnQm1ZcW9QZWJRdkFoRzBJQWRDYk9UdW9XOXJYK2NBUktFK2hGVE9UUjRBMkNWQkN3Q1lGeCtVRUltZ0JiUFdkTU5uZjNDbEFDdExCaUFRRnkxSXB1a0c5U0VBd000SVdnREF2SHkxTHdJWjYwT3VMWVFaVzFzZW1idnIyOW9IV2dDRTRmOGxFbE1mQWdEc2pLQUZBTXpJNXZMbVNYMEl3Ymhxd1N3MTNUQys4aisyUFRLMnJhcHFhY0VBQk9SM1dsSngwUUlBMkJsQkN3Q1lIL1VoUlBMcDdQNzJ5RWFZazZZYmpud0FUUUhXZlZ1L1dEUUFBYmxxUVNxQ0ZnREF6Z2hhQU1EOHFBOGhHbGN0bUp1eE1tUmhhMlRzdVcvcmxRVURFTlRHWWtqa2RBcGRBd0I4bUtBRkFNeU0raEFDRXJSZ05wcHVHRit4WGRrWW1YT3hCWURJQkMxSXlWVUxBR0FuQkMwQVlKN1VoeENKK2hEbVpHMWJaTzZoYjJ2WHJ3QUlxMi9yTVdpeHRTRVN1VEI0QUdBWEJDMEFZSjU4Z0VJMHJsb1FYdE1OMTFWVm5kc1VtYnUyWUFCbXdGVUxVbkhSQWdEWUNVRUxBSmlocVQ3a3dlNEl4QWQ3aERaMU1hOXNpY3o5MmJmMWt5VURNQVBmTFlsRVhMUUFBSFpDMEFJQTVrdDlDSkdjbjkzZm50Z0lnUzJycWpxMklESzJGU1lDWUVZRUxVaGwwWFNEMzEwQmdBOFR0QUNBK1ZJZlFqVHFRd2hwK2tQcTBuYkkzS3B2NnhkTEJtQW1WSWVRa3FzV0FNQ0hDVm9Bd0V4dExtL0dEMU8rMlIrQnFBOGhxdlg0Y3MxMnlOaHozOVpyQ3daZ0xxWnc0TE9Ga2NpWndRTUFIeVZvQVFEejVxb0ZrWnlxRHlHYXBodkcxMnFmTEliTUNib0JNRWZxUTBoRjBBSUErREJCQ3dDWU4wRUxvbEVmUWpSZStaTzdiMzFiKzZBS2dEbFNIMElxNXlZUEFIeVVvQVVBekpqNkVBTHlxcG93bW00WXZ4NVBiWVRNTFMwWWdKa1NGQ1NaNmZJZEFNQzdDVm9Bd1B5NWFrRWs2a01Jb2VtR0k5Y3NLTUJ0MzlaUEZnM0FIUFZ0N2FJRktha1BBUUErUk5BQ0FPWlAwSUpvdks0bWd2SHJjR0VUWkd3clRBUkFCaDRza1VRRUxRQ0FEeEcwQUlDWm0rcEQ3dXlSUUQ1YkJpazEzVEJlVmJteEJESzM3TnY2eFpJQm1EbFhMVWhGZFFnQThDR0NGZ0NRQjFjdGlPVDQ3UDdXNnlCUyttTDZaTzZ4YjJ0ZjV3RGs0THN0a3NqeFZEY0lBUEF1Z2hZQWtJSE41YzNYNllRNFJIRnRFNlRRZE1QNE11M2M4TW1jaWlZQWN1R2lCU2w1SUFBQXZKdWdCUURrdzFVTElsRWZRaXBlK1pPN3U3NnR2ZjRGSUF0OVd6OVZWZlZzbXlTaVBnUUFlRGRCQ3dESWg2QUZrYWdQNGVDYWJoaGYrUitiUEJrYnIxZXRMQmlBekxocVFTcUNGZ0RBdXdsYUFFQW0xSWNRa1BvUURtYnFWL1lCTkxsYlR5OS9BU0FuZ2hhazRuRUFBUEJ1Z2hZQWtCZFhMWWhFMElKREdrTVdDeE1uWStOWjliVUZBNUFobFZpa3NtaTZRZGdDQUhnWFFRc0F5SXVnQlpFc3p1NXZQOXNJK3piOWNmUVBneVp6cTc2dFh5d1pnTnowYlMxb1FVcUNGZ0RBdXdoYUFFQkcxSWNRa0tBRmgrQ1ZQN2w3Nk52Nml5MERrTEZIeXlVUlFRc0E0RjBFTFFBZ1A2NWFFSW1nQlh2VmRNUDROWFp1eW1SdWFjRUFaTTVWQzFLNU1Ia0E0RDBFTFFBZ1AxNTJFNG42RVBiTjl6eHlkOWUzOWNhV0FjaWMvK3RJNWRUa0FZRDNFTFFBZ014c0xtL0dQMUE5Mnl1QkNGcXdGMDAzcktxcU9qWmRNcloxelFLQVFyaG9RVEpOTjdocUFRQzhtYUFGQU9SSmZRaVJmRDY3dnoyeUVYYXA2WVlUSDBCVGdIWGYxaThXRFVEdStyWittZ0tHa01LWnFRTUFieVZvQVFCNSttS3ZCTEp3MVlJOVdFMWZXNUNyNTc2dFY3WUxRRUhVaDVDS2l4WUF3SnNKV2dCQWh0U0hFSkNnQlRzem5mYTlNbEV5ZDIzQkFCUkdmUWlwdUdnQkFMeVpvQVVBNUV0OUNKRjhVaC9DRG5ubFQrNGUrcmIyWVJNQXBmRi9INmtjTjkzZzkxVUE0RTBFTFFBZ1grcERpTVpWQ3o2czZZYnhsZis1U1pJNTF5d0FLSkhxRUZKU0h3SUF2SW1nQlFCa1NuMElBUWxhOENIVEt6UFhMTWpkbjMxYlA5a3lBS1hwMi9xbHFxcEhpeWNSOVNFQXdKc0lXZ0JBM3RTSEVJbjZFRDVxT1o3MU5VVXl0aFVtQXFCd3JscVFpb3NXQU1DYkNGb0FRTjdXOWtzd3JscndMazAzbkZSVmRXTjZaRzQxdmVZRmdGSUpXcENLZWtJQTRFMEVMUUFnWTV2TG15ZW5Wd2xtYVNHOGsrQVl1WHZzMjlyWE9RQ2wrMTc2QUVpbjZRYjFJUURBcXdsYUFFRCt2dGd4Z1p5ZTNkK2VXQWh2MFhURGVNYjNrNkdST1VFMEFJclh0N1dMRnFRa2FBRUF2SnFnQlFEazc2c2RFNHo2RU43S0szOXk5NjF2YXk5NEFlRGZIc3lCUkM0TUhnQjRMVUVMQU1pYytoQUN1cllVWHF2cGh2SHI1ZFRBeUp4ckZnRHdYOEtIcE9LaUJRRHdhb0lXQUZBRzlTRkVvajZFVjJtNjRjZzFDd3B3MjdmMWswVUR3SCtvRHlFVkFXOEE0TlVFTFFDZ0RPcERpRVo5Q0sreHFxcHFZVkprYkN0TUJBRC9JR2hCTWswM3FBOEJBRjVGMEFJQUNxQStoSURVaC9CTFRUZU1WMC8rTUNVeXQremIrc1dTQWVDL3BrdFB6MFpDSW9JV0FNQ3JDRm9BUURuVWh4REpXQitpLzVaZjhUMkwzRDMwYmUzckhBQit6RlVMVXZGN0tnRHdLb0lXQUZBTzlTRkU0Nm9GUHpTZDZ6MDNIVEszc21BQStLbnZSa01pZ2hZQXdLc0lXZ0JBSWFiNmtHLzJUU0NmTFlPZjhNcWYzTjMxYmUwREpBRDRPUmN0U09WNHFqRUVBUGdsUVFzQUtJdXJGa1J5ckQ2RS85VjB3M0w4MmpBWU1yWjF6UUlBZmswZ2tjVDhuZ29BL0phZ0JRQ1VSZENDYU5TSDhCOU5OeHo1QUpvQ3JQdTJmckpvQVBpdEJ5TWlrUXVEQndCK1I5QUNBQXF5dWJ4NVVSOUNNT3BEK0t0MVZWVUxFeUZqejlQWE9RRHdlK3BEU01WRkN3RGd0d1F0QUtBOHJsb1F5VmdmNHJVUTR6V0w4WStaVnlaQjVwWjlXNzlZTWdDOGlxQUZxWnliUEFEd080SVdBRkFlUVF1aVVSOUM1WlUvQlhqbzI5ci93UUR3ZXQvTmlsU21JRGdBd0U4SldnQkFZZFNIRUpENmtNSTEzZkRacXpFS3NMUmtBSGk5dnEyZnFxcmFHaG1KdUx3SUFQeVNvQVVBbE1tTFdpSlpuTjNmQ2xzVXF1bUdJOWNzS01CZDM5Yk9ud1BBMjdscVFTb3VXZ0FBdnlSb0FRQmxFclFnR2tHTGNvMnYvSTlMSHdKWjI3cG1BUUR2SnFoSUtvSVdBTUF2Q1ZvQVFJR20rcEE3dXljUVFZc0NOZDF3NGdOb0NyRHEyL3JGb2dIZ1hWeTBJSlhUNmZvZUFNQVBDVm9BUUxsY3RTQVM5U0ZsV28yN0wzMElaTzI1YjJ2Vk9BRHdmaTVha0pLckZnREFUd2xhQUVDaE5wYzNYNmR6NWhDRm9FVkJtbTY0cUtycXF2UTVrTDFyS3dhQTk1dXVRajBhSVlsY0dEd0E4RE9DRmdCUU5sY3RpT1RxN1A3V2FkWnlyRW9mQU5sNzZOdmF1WE1BK0RoWExVakZSUXNBNEtjRUxRQ2diSUlXL0I5Nzk1TVRTWkx0QzlqeTZjN2hyZ0I2QlVtdkFIcmkwK1RKSlZmTU1tb0ZGYldDQ0Zad3lSVTB6Rkl1dVI1TWZkTEpDdHBaUWNFS0xxd2ducnpiczVxcXlqK1FSSVNibTMyZmhIclNVa1djRThrZnQ1K2RFeHRUTFRKUTFtMS95Lzg0OXpxUVBOTXNBR0F6QkJjWmk0a1dBTUJYQ1ZvQVFNYXNEeUZDZ2hhSksrdTJuMXB5bm5zZFNONkhwaXJ1dEJrQU5zSkVDOGF5VjlidG9lb0RBRjhpYUFFQW1HcEJUTjVaSDVLOFJmL0FNdmNpa0xSSHEzRUFZSE9hcWhDMFlFeW1XZ0FBWHlSb0FRQUlXaEFiVXkwU05kd0dXK1plQjVLM2FLcmlRWnNCWUtOdWxKT1JIQ2s4QVBBbGdoWUFrRG5yUTRpUW9FVzZyQXdoZGJkTlZWem9NZ0JzM0NjbFpTU0NGZ0RBRndsYUFBREJWQXNpWTMxSWdzcTY3VWZ1dnN1OURpUnZvY1VBc0JYV2h6Q1dZNVVIQUw1RTBBSUFDRzZaRTZHNXBpVEg5eGxTZDkxVWhkdTJBTEFkZnNZeW1pRTBEZ0R3TzRJV0FFQy9QcVMvSFhTdkVrUkUwQ0loWmQzMnQvemY1bDRIa21lYUJRQnNTVk1WRC81bVpVVFdod0FBZnlKb0FRQjhabjBJTVhsNzlQSHNVRWVtcjZ6YmZnM01LdmM2a0x5enBpcnV0QmtBdHNyNkVNWWlhQUVBL0ltZ0JRRHcyWVZLRUpsVERVbENIN0xZeTcwSUpPM2VhaHdBMkFuclF4aUwxU0VBd0o4SVdnQUEvMko5Q0JHeVBtVGl5cnJ0cDVMOG5Ic2RTTjVxR0djT0FHeVhpUmFNNVdDWTFBY0E4QnRCQ3dEZ0tldERpSW4xSWROblVnNnB1Mm1xd3VjY0FIYWdxUW9UTFJpVHFSWUF3TzhJV2dBQVR6a3NJamFtV2t4VVdiZjlnOGpqM090QThsWmFEQUE3ZGFQY2pPUkk0UUdBcHdRdEFJRGZXQjlDaEFRdHBrdHdpOVJkdWxrTEFEdG5mUWhqTWRFQ0FQZ2RRUXNBNEk4Y2poS1RPOTJZbnJKdSsxditCN25YZ2FROW1tWUJBS01RY21Rc0psb0FBTDhqYUFFQS9KR2dCYkhvRHpKUGRXTmF5cnJkRHlFc2NxOER5VHR2cWtJUURBQjJ6MFFMeHJKWDFxMndCUUR3RzBFTEFPQjN1dG15UHppNlZSVWljTnJObGc4YU1Ubm4vVVBJM0l0QTB1NmJxakROQWdCR01BUWRIOVdla1FoYUFBQy9FYlFBQUw3RVZBdkdkdGJObHNZQ1Q4eHd3K3Q5N25VZ2VTYTJBTUM0L0ozQVdBUXRBSURmQ0ZvQUFGOXlwU3FNNkthYkxkMFduNmJ6M0F0QThtNmFxdkF6RWdER1pYMElZemxSZVFEZ00wRUxBT0JQckE5aFJQMFk0Rk1ObUo2eWJ1Y2hoT1BjNjBEeVRMTUFnUEdaYU1GWTNxbzhBUENab0FVQThEWFdoekNHMDI2MmZGRDVhU25yZGorRVlBb0pxYnRzcXNJTldnQVlXVk1WZ2hhTXBxeGJVeTBBZ0g4UnRBQUF2c1pvZEhidHJKc3RQVFNkcHY2Vy8wSHVSU0JwajZaWkFFQlVUR0JrTEVjcUR3QUVRUXNBNEd1c0QySEhicnJaMGtTRUNTcnI5dEFCTkJsWU5WVmgyZzRBeE1PVUtjWmlvZ1VBOEMrQ0ZnREF0MWdmd2k3ME44VlBWWHF5K29ETVh1NUZJR24zVFZXY2F6RUFSTVVrUE1aaW9nVUE4QytDRmdEQXR3aGFzQXVuM1d6cHB2Z0VEZnVKMytkZUI1STMxMklBaUk2SkZvemxZSmpxQndCa1R0QUNBUGlxNGZEN1dvWFlvck51dG5RYmJicmM4aWQxTjAxVitCNEZBSkZwcXFJYkp1UEJHRXkxQUFBRUxRQ0E3N3BTSXJia3Bwc3RWNG83VFdYZDlyZjgzK1plQjVKbm1nVUF4TXRVQzhZaWFBRUFDRm9BQU44bGFNRTI5TGZQVGxWMm1zcTYzVGZOZ2d5Y05WVnhwOUVBRUMxVHB4akxpY29EQUlJV0FNQTNXUi9DbHB3T255Mm1hUkZDMk5NN0V2WW9UQVFBMFRQUmdyRWNxendBSUdnQkFEeUhxUlpzMG9kdXRuVDdiS0xLdWowTUlTeHpyd1BKV3pSVklRd0dBSEh6TndXaktldlcraEFBeUp5Z0JRRHdISUlXYk1wdE4xc3VWSFBTM1BJbmRiZE5WVnpvTWdERWJRaEYzbXNUSXhHMEFJRE1DVm9BQU45bGZRZ2Iwby9pUDFYTTZTcnJ0dDlGL0M3M09wQThZVEFBbUE1VExSakxpY29EUU40RUxRQ0E1ekxWZ3RlYWQ3UGxuU3BPbWx2K3BPNjZxUW9ITmdBd0haMWVNUklUTFFBZ2M0SVdBTUJ6WFEwVENlQkhmT2htUzJHZENTdnJ0ci9sZjVCN0hVamFvMmtXQURBNUFwS001YTNLQTBEZUJDMEFnR2NaMW9jNEtPZEgzSGF6cGNQTENTdnJkaitFc01xOURpVHZ2S2tLVTNjQVlFS2FxakRSZ3RFTXF4VUJnRXdKV2dBQUx5Rm93VXYxTjhSUFZXM3krcERGWHU1RklHbjNmZEJDaXdGZ2ttNjBqWkVJV2dCQXhnUXRBSUJuRzFZL1dCL0NTOHk3MmRJTjhRa3I2L1l3aFBCejduVWdlYXVtS2g2MEdRQW15VlFMeG5LazhnQ1FMMEVMQU9DbFRMWGd1VDRNNFJ5bTdVTC9TTnhOVXhVKzV3QXdYWi8wanBHWWFBRUFHUk8wQUFCZXlzRTV6M0hielpZTGxacTJzbTc3dFMvSHVkZUI1SzIwR0FBbXpVUUx4ckkzVEFBRUFESWthQUVBdklqMUlUeEQvL2s0VmFna25PZGVBSkozMlZTRlc3QUFNR0ZOVmZTckN1LzFrSkZZSHdJQW1SSzBBQUIraEtrV2ZNdThteTN2VkdqYXlycnRiL2tmNUY0SGt0YUh3a3plQVlBMG1HckJXS3dQQVlCTUNWb0FBRDlDMElLditUQk1QV0hDeXJyZGR3Qk5CczZicW5qUWFBQklncUFGWXpIUkFnQXlKV2dCQUx5WTlTRjh4VzAzV3pxY1QwTy9NbVF2OXlLUXRQdW1LbFphREFESnNBcU1zUnlyUEFEa1NkQUNBUGhSRnlySEUzM3c1bFJCcHErczIzNzA3ZnZjNjBEeWhNSUFJQ0ZOVlFoYU1KcXliazIxQUlBTUNWb0FBRDlLMElLbjV0MXNlYWNpU1hETG45VGRORlZoeFJFQXBPZFdUeG5KaWNJRFFINEVMUUNBSDlMTmx2ME8zSHZWSTRUd1lWZ253OFNWZFRzMytwWU1tR1lCQUdreTFZS3htR2dCQUJrU3RBQUFYc1BoT3JmZGJPblFNZ0ZsM2U2YlprRUdQalJWMFdrMEFDVEp6M2pHSW1nQkFCa1N0QUFBWHNQNmtMdzloaEJPY3k5Q1F2ckF6RUh1UlNCcGo4SkVBSkEwUVF2RzhuWUlyZ01BR1JHMEFBQittUFVoMlp0M3MrVmQ3a1ZJUVZtM2g5WXBrSUZWVXhVUEdnMEFhUnFtVmoxcUx5TXgxUUlBTWlOb0FRQzhsdlVoZWZyUXpaWjZuNDd6RU1KZTdrVWdhZmROVlp4ck1RQWt6MVFMeG5LaThnQ1FGMEVMQU9DMXJBL0p6MjAzVzVwK2tJaXlidnNIZ3U5eXJ3UEptMnN4QUdUaGt6WXpFa0VMQU1pTW9BVUE4Q3JXaDJTbkg4VjdtbnNSRXVPV1A2bTdicXJDb1FzQTVNSFBmTVppZFFnQVpFYlFBZ0RZQkZNdDhqSHZac3U3M0l1UWlySnUrMXYrYjNPdkE4a3pnUWNBOG1GMUNHUFpLK3YyVVBVQklCK0NGZ0RBSmdoYTVPRkROMXRlNVY2RVZKUjF1MithQlJrNGE2cENPQXdBTXRGVXhZT0ppNHpJK2hBQXlJaWdCUUR3YXNPRWcxdVZUTnB0TjF1NkZaNld2cDk3dVJlQnBEMEtFd0ZBbHF3UFlTeldod0JBUmdRdEFJQk5NZFVpWGYxaDVXbnVSVWpKTU5KMm1Yc2RTTjVpdU5VS0FPVEYraERHSW1nQkFCa1J0QUFBTnNWS2lYVE5oNmtscEVNd2l0VGRObFhoY3c0QWVUTFJnckVjcXp3QTVFUFFBZ0RZQ090RGt2V2hteTJGYUJKUzF1MkpCNEJrd0tvakFNaFVVeFVtdTFTVTdRQUFJQUJKUkVGVVdqQ2E0ZTh0QUNBRGdoWUF3Q2E1UFp5VzIyNjJkRmlaSHY5T1NkMWxVeFZ1c2dKQTNtNXlMd0Nqc1Q0RUFESWhhQUVBYkpMSkIrbDREQ0djNWw2RTFKUjEyd2RuRG5LdkEwbnJ2M2V0dEJnQXNpZDB5VmhNdEFDQVRBaGFBQUFiWTMxSVV1WkRQMGxFV2JmN0RxREp3SGxURmI1M0FRRFdoekFXRXkwQUlCT0NGZ0RBcHAycjZPUjk2R1pMMDBuUzAvL2IzTXU5Q0NUdDNzOGdBR0FnYU1GWURvYVFPd0NRT0VFTEFHRFRITkJQMjIwM1d5NXlMMEpxeXJydGIxVzl6NzBPSkcvVlZNV0ROZ01BdzRTcisrd0x3VmlzRHdHQURBaGFBQUFiMWMyVy9TSFh0YXBPMG1NSTRUVDNJaVRLTFg5U2Q5TlV4WVV1QXdCUG1HckJXS3dQQVlBTUNGb0FBTnRncXNVMHpidlo4aTczSXFTbXJOcytQSE9jZXgxSW5razhBTUFmZlZJUlJtS2lCUUJrUU5BQ0FOZ0dRWXZwK2RETmx2cVdKdE1zU04xbFV4VnVyQUlBZitUM0E4WmlvZ1VBWkVEUUFnRFlPT3RESnVlMm15M2RCazlRV2JlckVNSkI3blVnYVkrbVdRQUFYOUpVaFlrV2pHV3ZyRnRoQ3dCSW5LQUZBTEF0cGlOTVEzOUllWnA3RVZKVTF1MmhBMmd5Y041VXhZTkdBd0JmY2Fzd2pFVFFBZ0FTSjJnQkFHeUxvTVUwekx2WjhpNzNJaVNxbjJheGwzc1JTTnA5VXhVckxRWUF2c0ZVQzhZaWFBRUFpUk8wQUFDMnd2cVFTZmpRelpZQ01Ra3E2L1lraFBBKzl6cVFQQk5iQUlEdjZWU0lrWndvUEFDa1RkQUNBTmdtaC9qeHV1MW1TNGVVNlhMTG45VGRORlhoWnd3QThEMkNGb3pscmNvRFFOb0VMUUNBcmVsbXk0c1F3cU1LUjZmdnlXbnVSVWhWV2JmekVNSng3blVnZVhNdEJnQytwS3piL2Y1MzRySnUrMURtUHhXSnNReVRCZ0dBUlAyWHhnSUFXM1psaFVGMDV0MXNlWmQ3RVZMVVAxUTJ6WUlNZkdpcXd2Y3dBT0EzdysvQnA4UFhPNVVoRW4zUTRwTm1BRUNhQkMwQWdHMFR0SWpMaDI2Mk5HNC9YZjA2bUlQY2kwRFNIb1dKQUlBZ1hNRTBIT2tUQUtUcnpYcTkxbDRBWUt1T1BwNDloQkQyVkhsMHQ5MXM2VUZQb3NxNlBRd2gvSnA3SFVqZUwwMVZuR3N6QU9SSnVJS0p1VytxNGxEVEFDQk4vMGRmQVlBZE1FRmhmUDB0OEhudVJVaWN3MmRTZHl0a0FRQjVLdXYydEt6Yi91L0svdzBoL0YzSWdvazRHQUx4QUVDQ3JBNEJBSGJCK3BEeExiclpzc3U5Q0trcTYvYkV3Mll5c05Ca0FNaEhINjU0TXIzQ2hFU21xcDhxZWFkN0FKQWVRUXNBWU91NjJmTHE2T1BabzRkam83bnNac3VMVE45N0x0enlKM1hYVFZWODBtVUFTSnR3QlFrNk11VVRBTklrYUFFQTdJcXBGdU80ZFFzOGJXWGQ5aXRoM3VaZUI1TG4reGdBSkVxNGdzU2RhREFBcEVuUUFnRFlGVUdMM2V1bmlNeTcyZklodHplZWk3SnU5MDJ6SUFOblRWVVl0d3dBQ1JHdUlDUEhtZzBBYVhxelhxKzFGZ0RZaWFPUFovMUIyWUZxNzh4UFZvYWtyYXpiUG1UeGMrNTFJR2w5WU95d3FRcUJNUUNZT09FS012YlhwaW82SHdBQVNJdUpGZ0RBTGwwNUZONlpTeUdMdEpWMWUramZFeGxZQ0ZrQXdIUUpWOEMvSElVUUJDMEFJREdDRmdEQUxsMDRHTjZKMi81d01vUDNtVHRCR2xKMzAxU0Z6emtBVEl4d0JmekppYi9mQUNBOWdoWUF3TTUwczJWMzlQSHMzdnFRcmVySDdNKzcyZElOOElTVmRYdGkxeThaV0dreUFFeERXYmY5amYzNUVLN3c5eDc4M3BGNkFFQjZCQzBBZ0YyelBtUzdGbjJnSmVVM3lMKzREVVhxTHB1cStLVExBQkF2NFFwNHRyZGwzZTViaVFjQWFSRzBBQUIyemZxUTdibnNaa3NIOElrcjYzYmxRVGFKZXpUTkFnRGlKRndCUDZ6L3R5TklEQUFKRWJRQUFIYksrcEN0dWUybldTVDYzaGowdDZEMG1ReWNOMVZ4cDlFQUVBZmhDdGlJRTBFTEFFaUxvQVVBTUFiclF6YXJ2LzA5NzJaTFkwalRkeDVDMk11OUNDVHRmdmljQXdBakVxNkFqVHRTVWdCSWk2QUZBREFHNjBNMmE5RlBDa25wRGZGbnc4UHU5MHBENGhaMlZ3UEFPSVFyWUt0T2xCY0EwdkptdlY1cktRQ3djMGNmeis0OHZOdUl5MjYybkNmd1B2aU9zbTc3TWJQSDZrVENicHFxOEFBYUFIWkl1QUoyNmk5VzVBRkFPa3kwQUFERzBvK0cveC9WZjVYYi92YjNoRjgvejFUVzdhbVFCUm53L1F3QWRxQ3MyOE1oWERFWHJvQ2Q2b05OZ2hZQWtBaEJDd0JnTEZlQ0ZxL3kyRDhZN1daTEkvWVRWOWJ0L2hCTWdwUmRObFZoQlJJQWJNa1Fyamdkd2hWdjFSbEdjVEk4Q3dFQUVpQm9BUUNNb3BzdDc0NCtudDE2eVBmREZ0MXM2VkF5RHdzM0RVbmNvMmtXQUxCNXdoVVFuU010QVlCMENGb0FBR082TU5YaWgxeDJzK1hGQkY4M0x6UThISGNBVGVyT202b3duUWNBTmtDNEFxSm1IU1FBSkVUUUFnQVlrL1VoTDNmcjREMHJxeERDWHU1RklHbjNUVldzdEJnQWZweHdCVXhIV2JkSFZ1WUJRQm9FTFFDQTBWZ2Y4bUw5ZVAxNU4xdTYrWjJCc203Ny9iM3ZjNjhEeVp0ck1RQzhuSEFGVEZiL2Q1NmdCUUFrUU5BQ0FCaWI5U0hQdCtobVN3OWs4dUdXUDZtN2FhcmlreTREd1BNSVYwQVNqclFSQU5JZ2FBRUFqTTM2a09lNTdHYkxpeW04VUY2dnJOdTUvYjFrd0RRTEFQZ080UXBJem9tV0FrQWEzcXpYYTYwRUFFWjE5UEdzODlEd20vcjFLaWRXaHVTaHJOdjlFTUpkQ0dFdjkxcVF0QTlOVlN5MEdBRCtiUGg5OEhUNGVxZEVrSnovYnFyQzMvY0FNSEVtV2dBQU1UZ1BJZnhkSjc3b3NiKzlKbVNSbFlXUUJZbDd0Qm9IQUg1UHVBS3kwcThQc1VJUEFDWk8wQUlBaU1HVm9NVlhMYnJac292MHRiRmh3MmpvcGJxU3VKVWJmQUFnWEFFWk94RzBBSURwRTdRQUFFYlhUMnM0K25oMjdlSGluMXgycytWRlpLK0o3VHBYWHhKMzIxU0Z6emtBMlJLdUFJYWdCUUF3Y1lJV0FFQXNyanhvL0ozYllZVUVtU2pyOXNTL0FUTGcreG9BMlJHdUFQN2dTRUVBWVBvRUxRQ0FXRmdmOGgrUElZUjVQK2tqbGhmRVRyamxUK3F1bTZvd0lobUFMQWhYQU4rdzE2K05iS3JpVHBFQVlMb0VMUUNBS0ZnZjhqdUxicmJzSW5vOWJGbFp0LzB0LzdmcVRPSk1zd0FnYWNJVndBdjBFdzJ0Q2dXQUNSTzBBQUJpWW4xSUNKZmRiT2xoUzBhR0IvS3IzT3RBOHM3YzJBTWdWV1hkOXNHS3ViOWxnQmV3UGdRQUprN1FBZ0NJU2U3clEyN2QrTTVTSDdMWXk3MElKTzNlYWh3QVVqT0VLejUvK1YwT2VDbEJDd0NZdURmcjlWb1BBWUJvSEgwOHkzV3F4V00vT3RUS2tMejBlM2xEQ0wvbVhnZVM5MU5URlNiMUFEQjV3aFhBSmpWVjhVWkJBV0M2VExRQUFHSnprV25RWWlGa2tTV0h6NlR1UnNnQ2dDa1RyZ0MycGF6Yms2WXFQaWt3QUV5VG9BVUFFSlZ1dHJ3NituajJtTmxEek10dXRuUVFtWm5ob2YxeDduVWdlU3N0Qm1CcWhDdUFIZW5YaHdoYUFNQkVDVm9BQURIcTE0ZTh6NlF6dC8wMGl3aGVCN3QzcnVZazd0SU5QUUNtUXJnQ0dNR0p2d3NCWUxvRUxRQ0FHT1VTdE9nbmQ4eTcyZkloZ3RmQ0RwVjEyOS95UDFCekV2Wm9tZ1VBc1JPdUFFWjJwQUVBTUYxdjF1dTE5Z0VBMFRuNmVQYVF3Y1BPbjZ3TXlVOVp0L3NoaERzUDgwbmNXVk1WZ2hZQVJLZXMyLzRHK1Z5NEFvakVmemRWNGZJRkFFeVFpUllBUUt4U24ycHhLV1NSclhNUDlVbmN2WkFGQURFcDYvYm9TYmpDVkRFZ0ppZkQ4dzhBWUdJRUxRQ0FXS1VjdExnTklTd2llQjNzMlBDUVA0ZTFPT1ROOXpjQVJpZGNBVXpFa2FBRkFFeVRvQVVBRUtWdXRydzYrbmoybU9ETi8vNDl6YnZaMG1qUVBKM25YZ0NTZDlOVWhRZkZBSXhDdUFLWW9CTk5BNEJwRXJRQUFHS1c0bFNMUlRkYmRoRzhEbmFzck52K29mK3h1cE00MHl3QTJDbmhDbURpampRUUFLYnB6WHE5MWpvQUlFcEhIOC82bXgzL1NLZzdsOTFzT1kvZ2RiQmpaZDN1OTROYVBQd25jWmROVmZnZUI4RFdDVmNBaWZsclV4VXVaQURBeEpob0FRQkVxNXN0UHgxOVBMdFA1T0hwclp2ZVdWczRCQ0J4ajc3SEFiQk53aFZBd282R1lENEFNQ0dDRmdCQTdQcjFJVDlQdkV2OUFlUzhteTBmSW5ndDdGaFp0NGNPb01uQXFxa0szK01BMkNqaENpQVQvVFRQQzgwR2dHa1J0QUFBWW5lUlFOQmkwYzJXYnFma2F4VkMyTXU5Q0NUdHZxbUtjeTBHWUJPR2tPcXBpV0JBUm80MEd3Q21SOUFDQUloYUgxQ1krUHFRNjI2MmRETWxVMlhkOWplVDN1ZGVCNUkzMTJJQVh1Tkp1S0wvbWZKV01ZSE0rTDRIQUJNa2FBRUFUTUZVMTRmY080RE1ubHYrcE82bXFZcFB1Z3pBU3dsWEFQeEhIOUwzZXpVQVRJdWdCUUF3QlZOZEgzTGF6WllQRWJ3T1JsRFdyVU1EY2lCTUJzQ3pDVmNBZkZVL0RWSFFBZ0FtUk5BQ0FJamVSTmVIL05LLzdnaGVCeU1vNjNiZk5Bc3k4S0dwaWp1TkJ1QmJoQ3NBbnVWSW1RQmdXZ1F0QUlDcG1OTDZrT3R1dG5USW5yZEZDR0V2OXlLUXRNY1F3a3FMQWZnUzRRcUFGeE8wQUlDSkViUUFBS1ppS3V0RDdvM1N6OXR3c0xETXZRNGtiOUZVaGRWSUFQeEd1QUxnVlE3Njc2TW14Z0hBZEFoYUFBQ1RNS3dQdVozQVE5dlRiclowK0pnMzAweEkzVzFURlJlNkRNQ3dMbTB1WEFHd0VmMVVDMEVMQUpnSVFRc0FZRXI2ZzczL2lmajEvdElIUWlKNEhZeWtyTnVURU1JNzlTZHhDdzBHeU5jUXJqZ2R2dnplQTdBNVI4UGFWQUJnQWdRdEFJQXB1WW80YUhIZHpaWW1HZUNXUDZtN2JxcmlreTRENUVXNEFtQW5UcFFaQUtaRDBBSUFtSXh1dHJ5TGRIM0kvVEF1bVl5VmRkdmY4ai93R1NCaGo2WlpBT1JEdUFKZzU0NlZIQUNtUTlBQ0FKaWFHTmVIbkhhejVVTUVyNE9SREFjUksvVW5jZWROVmRnWkRaQXc0UXFBY1pWMWU5UlVoWldrQURBQmdoWUF3TlRFdGo3a2wyNjI5QkNFUG1TeGwzMFZTRmsvdWNkNkpJQUVDVmNBUktWZkgrSVpBd0JNZ0tBRkFEQXBrYTBQdWU1bVN3ZVBtZXR2SElVUWZzNjlEaVJ2MVZTRnlUMEFpUkN1QUlqV2tkWUF3RFFJV2dBQVV4VEQrcEQrZHZkOG11Vmp3NFJ0U04xTlV4VVh1Z3d3YmNJVkFKTWdhQUVBRS9GL05Bb0FtS0NyQ0Y3eWFUZGJ1dDJkdWJKdSs0T0s0OXpyUVBKV1dnd3dYZjN2SzJYZDlvRzUvdzBoL0YzSUFpQnFiNGRnSEFBUU9STXRBSURKR2RhSDNJeDR3UDFMTjF2YW1Vb3d6WUlNWERaVjhVbWpBYVpsQ0lOKy90clRQb0JKNmFkYStCMGNBQ0luYUFFQVROWEZTRUdMNjI2MmRMaE9mNERSMy9JL1VBa1M5aGhDV0dnd3dEUUlWd0FrNDBUUUFnRGlKMmdCQUV6VjFURDZlSmZ1UXdoem54aUdVYTRPb0VuZGVWTVZWaVFCUkV5NEFpQkpSOW9LQVBFVHRBQUFKcW1iTFIrT1BwNWQ3M2pIOUduLzMvV0pZVmdaNGpDRGxOMDNWYkhTWVlENENGY0FKTzlFaXdFZ2ZvSVdBTUNVWGUwd2FQRkxOMXQyUGkyVWRkcy85SHFmZlNGSW5Za3RBQkVScmdESXlsNVp0NGROVmR4cE93REVTOUFDQUppeVhhMFB1ZTVteTNPZkZBWnUrWk82bTZZcXJuUVpZRnpDRlFCWjY5ZUhDRm9BUU1RRUxRQ0F5ZHJSK3BEN0VNTGNwNFR3N3dPUC9yTndyQmdrempRTGdKRU1rN1Btd2hVQTJUc1pMcGNBQUpFU3RBQUFwbTdiNjBOTyswQ0hUd2xsM2U2YlprRUdQalJWWVUwU3dBNlZkWHYwSkZ4eG9QWUFEQk10QUlDSUNWb0FBRk8zemZVaHYzU3pwUU5IUGxzNC9DQnhqOEpFQUxzaFhBSEFkNWlrQ0FDUmU3TmVyL1VJQUppMG80OW4yNWhxY2QzTmxxYytHWVIvSDRZYzl0dHFqUEFtY2I4MFZYR3V5UURiSVZ3QndBdjlyYW1LVDRvR0FIRXkwUUlBU01IRmhvTVc5OE5EY1Bqc1hNaUN4TjBMV1FCc25uQUZBSy9RL3d3UnRBQ0FTQWxhQUFDVDE4MldWMGNmeng0M2VCQisyczJXRHo0WmhIOGZrSnhzWVdJS3hFYTRER0JEaENzQTJKQWpoUVNBZUFsYUFBQ3A2TmVIdk4vQWUvbWxteTA3bndxZWNNdWYxTjBZU1F6d09zSVZBR3pCaWFJQ1FMd0VMUUNBVkd3aWFISGR6WllPMWZsTldiZjlnY2xiRlNGeHBsa0EvQURoQ2dDMjdLQ3MyLzJtS2t6Y0JJQUlDVm9BQUVuWXdQcVFlNGVOUE5VLzBETE5nZ3ljTlZWeHA5RUF6MVBXN2VFUXJGZ0lWd0N3QTMyb3ovUTVBSWlRb0FVQWtKTFhUTFU0N1daTHQwUjRhdldLNEE1TXdhTXdFY0QzUFFsWG1IUUZ3SzZkQ0ZvQVFKd0VMUUNBbFB4bzBPS1hicmJzZkJMNGJEaFErVmxCU056Q0dHS0FMeE91QUNBU0p4b0JBSEVTdEFBQWt2R0Q2ME91dTluU2pXNys2RUpGU054dFV4VSs1d0JQQ0ZjQUVLRWpUUUdBT0FsYUFBQ3BlY2xVaS92aFFUcjhwcXpiL3NiUXNZcVF1SVVHQXdoWEFCQzl2ZjVuVlZNVmQxb0ZBSEVSdEFBQVVuUCtncURGYVRkYkdwdlBIN25sVCtvdW02cXc1eG5JbG5BRkFCTno0dTlVQUlpUG9BVUFrSlJ1dHV5T1BwNzFreW9PdnZPK2Z1bi92N3JQVTJYZExwN3gyWUVwNjljcnJYUVF5STF3QlFBVFpuMElBRVJJMEFJQVNGRy9QdVRuYjd5djYyNjJQTmQ1bmlycmR0OEJOQms0TjNZWXlNWHdzLzEwV0pja1hBSEFWSjNvSEFERVI5QUNBRWpSeFRlQ0Z2ZkRUVWI0b3o1OHM2Y3FKT3grK0p3REpPdEp1S0wvZXFmVEFDUkFXQkFBSXZSbXZWN3JDd0NRbktPUFozZGZXUUh4Vnl0RCtLT3lidnRSclA5VUdCTDNVMU1WZGpzRHlSR3VBQ0FEZjJ1cTRwTkdBMEE4VExRQUFGTDFwZlVodndoWjhCVnUrWk82R3lFTElDWENGUUJrcHI4Y0lHZ0JBQkVSdEFBQVV2WEg5U0hYM1d6cE1KMC9LZXUyUDZBNVZoa1N0OUJnWU9xRUt3REkySWtMQWdBUUYwRUxBQ0JKL2VTS280OW45OFA2a1A1LzV6ck5WM2hZUmVvdW02b3d6UWVZSk9FS0FQaVhJMlVBZ0xnSVdnQUFLZnU4UHVTMG15MGZkSm8vS3V0Mk5ZUnhJRldQcGxrQVV5TmNBUUIvY3REL2ZHeXF3ck1OQUlpRW9BVUFrTEorVWtIWFQ3ZlFaZjZvck50REI5Qms0TnpEV0dBS2hDc0E0THRPaGdzbEFFQUVCQzBBZ0dSMXMrVmRDT0ZDaC9tS2ZwckZudUtRc1B1bUtsWWFETVNzck52VEp3RUxQNWNCNE91T0JDMEFJQjZDRmdBQVpLZXMyLzRtMEh1ZEozRW10Z0JSRXE0QWdCOXlvbXdBRUE5QkN3QUFjdVNXUDZtN2FhckNiVGNnR3NJVkFQQnFSMG9JQVBGNHMxNnZ0UU1BZ0d5VWRUc1BJZnhkeDBuY1g1cXF1Tk5rWUV6Q0ZRQ3djWDl0cXFKVFZnQVluNGtXQUFCa282emJmZE1zeU1BSElRdGdMTUlWQUxCVi9WUUxRUXNBaUlDZ0JRQUFPVm1FRUE1MG5JUTlDaE1CdXlaY0FRQTdjeEpDdUZCdUFCaWZvQVVBQUZrbzYvWXdoTERVYlJLM2FxcmlRWk9CYlJPdUFJQlJIQ2s3QU1SQjBBSUFnRnljNnpTSnUyK3F3dWNjMkpxeWJvK0c2VkRDRlFBd2pyZnFEZ0J4RUxRQUFDQjVaZDMyNDFYZjZUU0ptMnN3c0dsRHVHSStoQ3VzM3dLQWNkMzNQNXVicXVqMEFRREdKV2dCQUVBTzNQSW5kZGROVlh6U1pXQVRoQ3NBSUNyM0lZU3JFTUtGZ0FVQXhFUFFBZ0NBcEpWMXV6QmVsUXdzTkJsNERlRUtBSWlLY0FVQVJFN1FBZ0NBWkpWMXV4OUNXT2t3aVR0cnF1Sk9rNEdYRXE0QWdLZ0lWd0RBaEFoYUFBQ1FzajVrc2FmREpPelJhaHpnSllRckFDQXFqMC9DRlZZQkFzQ0V2Rm12MS9vRkFFQnl5cm85RENIOHFyTWs3cWVtS2k0MEdmZ1c0UW9BaU1ybmNNVlZVeFZYV2dNQTAyU2lCUUFBcVhMNFRPcHVoU3lBcnhHdUFJQ29DRmNBUUdJRUxRQUFTRTVadHljaGhHT2RKWEVMRFFhZUdxWTVuUTRCaTdlS0F3Q2pFcTRBZ0lRSldnQUFrQ0szL0VuZHBSM09RQkN1QUlEWUNGY0FRQ1lFTFFBQVNFcFp0eXNqMGtsYy8vQjJwY21RTCtFS0FJaUtjQVVBWkVqUUFnQ0FaSlIxdTIrZEFoazRiNnJpVHFNaEw4SVZBQkFWNFFvQXlKeWdCUUFBS1RrUEllenBLQW03SHo3blFBYUVLd0FnT3RmOXFrcmhDZ0RnelhxOXpyNElBQUJNWDFtM1J5R0VmMm9saWZ1L0h1cEMyb1FyQUNBNjEwK21WenhvRHdBUVRMUUFBQ0FoYnZtVHVoc2hDMGlUY0FVQVJFZTRBZ0Q0SmtFTEFBQW1yNnpiL25EcVdDZEozRUtESVIxbDNlNFA0WXFGY0FVQVJFRzRBZ0I0TmtFTEFBQW1iVGlvTXMyQzFGMDJWZEhwTWt6YmszQkYvL1ZPT3dGZ2RNSVZBTUFQRWJRQUFHRHErcHZBQjdwSXdoNU5zNERwRXE0QWdPZ0lWd0FBcnlab0FRREFaQTA3N1IxQWs3cHpENEJoV29RckFDQTZ3aFVBd0VZSldnQUFNR1dyRU1LZURwS3crNllxVmhvTThST3VBSURvQ0ZjQUFGc2phQUVBd0NTVmRYc1NRbml2ZXlSdXJzRVFMK0VLQUlqT2JRamhZZ2hYM0drUEFMQXRnaFlBQUV5VlcvNms3cWFwaWsrNkRIRVJyZ0NBNkFoWEFBQTdKMmdCQU1Ea2xIWGIzL0kvMWprU1o1b0ZSR1Q0MlNOY0FRQnhFSzRBQUVZbGFBRUF3S1FNTjRuUGRZM0VmZkRBR01aWDF1M3BrK2tWZTFvQ0FLTVNyZ0FBb2lGb0FRREExQ3djZHBHNFI2dHhZRHpDRlFBUUZlRUtBQ0JLYjlicnRjNEFBREFKWmQwZWhoQisxUzBTOTB0VEZhYTJ3QTRKVndCQVZJUXJBSURvbVdnQkFNQ1VPSHdtZGJkQ0ZyQWJ3aFVBRUJYaENnQmdVZ1F0QUFDWWhMSnVUMElJNzNTTHhDMDBHTFpIdUFJQW9uTGZCeXY2UUwxd0JRQXdOWUlXQUFCTXhZVk9rYmpycGlvK2FUSnNsbkFGQUVUbGM3amlvcW1LVG1zQWdLa1N0QUFBSUhwbDNmYTMvQTkwaXNTWlpnRWJJbHdCQUZFUnJnQUFraU5vQVFCQTFNcTYzUThockhTSnhKMFpsd3l2VTlidFVRaGhQbndKVndEQXVJUXJBSUNrQ1ZvQUFCQzdsUU16RXZmWTc2WFdaSGk1SitHS1U1T1BBR0Iwd2hVQVFEYmVyTmRyM1FZQUlFcGwzUjZHRUg3VkhSTDNVMU1WRjVvTXp5TmNBUUJSRWE0QUFMSmtvZ1VBQURGeitFenFib1FzNFB1RUt3QWdLc0lWQUVEMkJDMEFBSWhTV2JmOVlkcXg3cEM0bFFiRGx3bFhBRUJVSG9jZ3ZIQUZBSkM5SUdnQkFFREV6aldIeEYwMlZmRkprK0UvaENzQUlDcVB3K1NLcTZZcXJyUUdBT0EvQkMwQUFJaE9XYmNyQjJ3azd0RTBDL2czNFFvQWlJcHdCUURBTXdoYUFBQVFsYkp1OTBNSUMxMGhjZWROVmR4cE1ya3E2L1p3K0Y0dlhBRUE0eE91QUFCNElVRUxBQUJpMDY4TTJkTVZFbmJmVklWcEZtUm5DRmVjRHRNcjN2b0VBTUNvaENzQUFGNUIwQUlBZ0dnTTQrUGY2d2lKTTdHRmJBaFhBRUJVaENzQUFEWkUwQUlBZ0ppYzZ3YUp1L0ZRbTlRSlZ3QkFWSVFyQUFDMlFOQUNBSUFvbEhYYkg4Z2Q2d2FKTTgyQ0pBbFhBRUJVaENzQUFMWk0wQUlBZ05HVmRic2ZRbGpwQkltN2JLcWkwMlJTSVZ3QkFORzVmaEt3ZU5BZUFJRHRFYlFBQUNBRy9TMy9BNTBnWVkrbVdaQUM0UW9BaUk1d0JRREFDQVF0QUFBWTFYQm81d0NhMUswOCtHYXFoQ3NBSURyQ0ZRQUFJeE8wQUFCZ2JQM0trRDFkSUdIM1RWV2NhekJUTXF4MCtoeXVPTlk4QUJpZGNBVUFRRVFFTFFBQUdFMVp0eWNoaFBjNlFPTG1Hc3dVUEFsWDlGL3ZOQTBBUmlkY0FRQVFLVUVMQUFERzVKWS9xYnRwcXVLVExoTXI0UW9BaUk1d0JRREFCQWhhQUFBd2lySnU3Zm9uQjZaWkVCM2hDZ0NJam5BRkFNREVDRm9BQUxCend5R2ZhUmFrN2tOVEZYZTZUQXlFS3dBZ09yZkQzMFRDRlFBQUV5Um9BUURBR0JZaGhEMlZKMkdQSVlTVkJqTW00UW9BaUU0ZnJyZ1l3aFVDdVFBQUUvWm12VjdySHdBQU8xUFc3V0VJNFZjVkozRS9OVlZ4b2Nuc21uQUZBRVJIdUFJQUlFRW1XZ0FBc0dzT24wbmRyWkFGdTFiVzdWeTRBZ0NpSVZ3QkFKQTRRUXNBQUhhbXJOdVRFTUt4aXBPNGhRYXpDMlhkbmo2WlhtRWRFd0NNUzdnQ0FDQWpnaFlBQU95U1cvNms3cnFwaWsrNnpMWUlWd0JBVklRckFBQXlKV2dCQU1CT2xIWGIzL0kvVUcwU1o1b0ZHeWRjQVFCUkVhNEFBRURRQWdDQTdTdnJkaitFc0ZKcUVuZm1ZVHViSWx3QkFGRzVEeUdjQzFjQUFQQ1pvQVVBQUx1d2NsQkk0ajQvZkljZkpsd0JBRkhwZjcrNzZxZFhORlhSYVEwQUFFKzlXYS9YQ2dJQXdOYVVkWHNVUXZpbkNwTzRuNXFxdU5Ca1hrcTRBZ0NpSWx3QkFNQ3ptR2dCQU1DMnVlVlA2bTZFTEhnSjRRb0FpSXB3QlFBQUx5Wm9BUURBMWd5SGljY3FUT0pXR3N6M0ROTjk1c09YY0FVQWpFdTRBZ0NBVnhHMEFBQmdtMHl6SUhXWFRWVjgwbVcrNUVtNG9nK2RIU2dTQUl4S3VBSUFnSTBSdEFBQVlDdkt1bDA1V0NSeGo2Wlo4RWZDRlFBUUZlRUtBQUMyUXRBQ0FJQ05LK3QyUDRTd1VGa1NkOTVVeFowbUkxd0JBRkVScmdBQVlPc0VMUUFBMklaK1pjaWV5cEt3KzZZcVRMUEltSEFGQUVUbDhVbTR3bG8zQUFDMlR0QUNBSUNOS3V2MkpJVHdYbFZKbklrdEdSS3VBSUNvZkE1WFhEVlZjYVUxQUFEc2txQUZBQUNiNXBZL3FidnhNRDhmd2hVQUVCWGhDZ0FBb2lCb0FRREF4cFIxMng5R0hxc29pVFBOSW5GbDNSNE9mUmF1QUlEeENWY0FBQkFkUVFzQUFEYWlyTnQ5MHl6SXdJZW1LanFOVHM4UXJqZ2RwbGU4emIwZUFEQXk0UW9BQUtJbWFBRUF3S1lzM1B3bWNZL0NSR2tScmdDQXFBaFhBQUF3R1cvVzY3VnVBUUR3S3NOaDVhK3FTT0orYWFyaVhKT25UYmdDQUtJaVhBRUF3Q1NaYUFFQXdDWTRmQ1oxOTBJVzB5VmNBUURSdVE0aFhBaFhBQUF3VllJV0FBQzhTbG0zSnlHRWQ2cEk0dVlhUEMzQ0ZRQVFuZXNuMHlzZXRBY0FnQ2tUdEFBQTRMWGM4aWQxTjAxVmZOTGwrQWxYQUVCMGhDc0FBRWlTb0FVQUFEK3NyRnVIbWVUQU5JdUlsWFc3L3lSY2NaeDdQUUFnQXNJVkFBQWtUOUFDQUlBZk1oeHVtbVpCNnM2YXFyalQ1Ymc4Q1ZlY1dsMEVBRkVRcmdBQUlDdUNGZ0FBL0toVkNHRlA5VWpZb3pCUlBJUXJBQ0E2d2hVQUFHVHJ6WHE5MW4wQUFGNmtyTnZERU1LdnFrYmlmbXFxNGtLVHh5TmNBUURSRWE0QUFDQjd3VVFMQUFCK2tNTm5VbmNyWkRFTzRRb0FpSTV3QlFBQS9JR2dCUUFBTDFMVzdVa0k0VmpWU054Q2czZEh1QUlBb25NN2hLdjdjTVdkOWdBQXdPOEpXZ0FBOEZKdStaTzY2NllxUHVueTlwVjFPeGV1QUlCb0NGY0FBTUF6Q1ZvQUFQQnNaZDMydC93UFZJeUVQWnBtc1YxbDNaNCttVjZ4bC9KN0JZQUpFSzRBQUlBZklHZ0JBTUN6REtQOVY2cEY0czRkTW15ZWNBVUFSRVc0QWdBQVhrblFBZ0NBNXpwM1FFcmk3b2ZQT1JzZ1hBRUFVUkd1QUFDQURSSzBBQURndThxNlBRb2h2RmNwRXJkcXF1SkJrMytjY0FVQVJFVzRBZ0FBdGtUUUFnQ0E1M0RMbjlUZE5GVnhvY3N2SjF3QkFGRVJyZ0FBZ0IwUXRBQUE0SnVHUTlSalZTSnhLdzErUHVFS0FJaEt2LzdzcWc5SEMxY0FBTUJ1Q0ZvQUFQQTlwbG1RdXN1bUtqN3A4cmNOSzRUbXc1ZHdCUUNNNjNPNDRxS3BpazR2QUFCZ3R3UXRBQUQ0cXJKdSsxditCeXBFd2g1RENBc04vckluNFlwVDN3c0FZSFRDRlFBQUVBbEJDd0FBdnFpczIwTUgwR1NnSDdIOW9OSC9JVndCQUZFUnJnQUFnQWdKV2dBQThEVXI2d0ZJM0gxVEZTdE5GcTRBZ01nSVZ3QUFRT1FFTFFBQStKT3liazlDQ085VmhzUmxQYkZGdUFJQW9pSmNBUUFBRXlKb0FRREFsN2psVCtwdW1xcTR5cTNMd2hVQUVCWGhDZ0FBbUNoQkN3QUFmcWVzMi80UTlsaFZTTnc4bHdhWGRYczRUTzhRcmdDQThUMzJ3UXJoQ2dBQW1EWkJDd0FBZmxQVzdYNEk0VnhGU055SHBpcnVVbjZMUTdqaWRBaVV2STNnSlFGQXpoNkh5UlZYT1U3VUFnQ0FGQWxhQUFEd1ZIL3JmVTlGU05oanFxdHhoQ3NBSUNyQ0ZRQUFrTEEzNi9WYWZ3RUErSHhJKzZ0S2tMaGZtcXBJWm1xTGNBVUFSRVc0QWdBQU1tR2lCUUFBbjFrWlF1cnVVd2haQ0ZjQVFGU0VLd0FBSUVPQ0ZnQUE5QWUzSnlHRWR5cVcxSVZvQUFBZ0FFbEVRVlJCNHVaVGZYdkNGUUFRRmVFS0FBREluS0FGQUFEQk5Bc3ljTjFVeGFjcHZVM2hDZ0NJaW5BRkFBRHdHMEVMQUlETWxYVzdjSWhMQmhaVGVJdGwzZTRQNFlwVFUyWUFZSFRDRlFBQXdCY0pXZ0FBWkd3NDFGMzVESkM0czZZcTdtSjlpOElWQUJDZDZ5RmNjYUUxQUFEQWx3aGFBQURrclE5WjdPVmVCSkwyR09OcUhPRUtBSWpPOVpQcEZRL2FBd0FBZk11YjlYcXRRQUFBR1NycjlqQ0U4S3ZlazdpZllybU5LbHdCQU5FUnJnQUFBSDZJaVJZQUFQa3lDcG5VM1k0ZHNoQ3VBSURvQ0ZjQUFBQ3ZKbWdCQUpDaHNtNVBRZ2pIZWsvaUZtTzhQZUVLQUlpT2NBVUFBTEJSZ2hZQUFIa3l6WUxVWFRaVjhXbVg3N0dzMno1WU1SZXVBSUFvQ0ZjQUFBQmJJMmdCQUpDWnNtNVhJWVFEZlNkaGp5R0UxUzdlM2hDdStQeTE1ME1GQUtNU3JnQUFBSFpDMEFJQUlDUERTb05SMWluQURwMDNWWEczcmYrY2NBVUFST1cyLzlrdlhBRUFBT3lTb0FVQVFGN09IUXlUdVB2aGM3NVJ3aFVBRUpYYllSWGUxVGJEbFFBQUFGOGphQUVBa0lteWJvOUNDTy8xbThTdE5uV2JWYmdDQUtJaVhBRUFBRVJEMEFJQUlCOGJ2K1VQa2JscHF1TGlOUzlKdUFJQW9pSmNBUUFBUkVuUUFnQWdBOFBoOGJGZWs3akZqN3c5NFFvQWlJcHdCUUFBRUQxQkN3Q0F4SlYxdTIrYUJSbTRiS3FpZSs3YkhGYnB6SWR3eFlFUENBQ01TcmdDQUFDWUZFRUxBSUQwTFJ3a2s3akg1MHl6RUs0QWdLZ0lWd0FBQUpNbGFBRUFrTEN5Ymc5L2RKMENUTWg1VXhVUFgzcTV3aFVBRUJYaENnQUFJQW1DRmdBQWFWdUZFUGIwbUlUZE4xV3hldnIyaENzQUlDcjNmYkNpRDFpOFpNMFhBQUJBekFRdEFBQVNWZGJ0U1FqaHZmNlN1RDVRSVZ3QkFIRVJyZ0FBQUpMMlpyMWU2ekFBUUlMS3V1MGZhci9WV3hMMitSQkh1QUlBeGlkY0FRQUFaRVBRQWdBZ1FXWGQ5amY3LzY2M0FBQnNrWEFGQUFDUUpVRUxBSURFbEhXN0gwSzRDeUhzNlMwQUFCc21YQUVBQUdUdnYzSXZBQUJBZ2haQ0ZnQUFiSkJ3QlFBQXdCTW1XZ0FBSktTczI4TVF3cTk2Q2dEQUt3bFhBQUFBZklXSkZnQUFhVG5YVHdBQWZ0RGprM0RGSjBVRUFBRDRNaE10QUFBU1VkYnRTUWpoSC9vSkFNQUxmQTVYWERWVmNhVndBQUFBMzJlaUJRQkFPaTcwRWdDQVp4Q3VBQUFBZUFWQkN3Q0FCSlIxdXdnaEhPZ2xBQUJmSVZ3QkFBQ3dJVmFIQUFCTVhGbTMreUdFdXhEQ25sNENBUENFY0FVQUFNQVdtR2dCQURCOUt5RUxBQUFHd2hVQUFBQmJacUlGQU1DRWxYVjdHRUw0VlE4QkFMSW1YQUVBQUxCREpsb0FBRXpiaGY0QkFHVHJVcmdDQUFCZzl3UXRBQUFtcXF6YjB4RENzZjRCQUdUbCtzbjBpZ2V0QndBQTJEMUJDd0NBNlRyWE93Q0FMQWhYQUFBQVJFVFFBZ0JnZ3NxNlhZVVFEdlFPQUNCWndoVUFBQUNSZXJOZXIvVUdBR0JDeXJyZER5SGNoUkQyOUEwQUlDbkNGUUFBQUJOZ29nVUF3UFNjQzFrQUFDUkR1QUlBQUdCaVRMUUFBSmlRc202UFFnai8xRE1BZ0VrVHJnQUFBSmd3RXkwQUFLYmxYTDhBQUNaSnVBSUFBQ0FSZ2hZQUFCTlIxdTA4aEhDc1h3QUFrM0ViUXJnWXdoVjMyZ1lBQUpBR3EwTUFBQ2Fnck52OUVFSVhRampRTHdDQXFBbFhBQUFBSk01RUN3Q0FhVmdJV1FBQVJFdTRBZ0FBSUNNbVdnQUFSSzZzMjhOaG1zV2VYZ0VBUkVPNEFnQUFJRk1tV2dBQXhPOWN5QUlBSUFyQ0ZRQUFBSmhvQVFBUXM3SnVUMElJLzlBa0FJRFJDRmNBQUFEd095WmFBQURFN1Z4L0FBQjJUcmdDQUFDQXJ4SzBBQUNJVkZtMzh4RENXLzBCQU5pSit5RmNjU0ZjQVFBQXdMZFlIUUlBRUtHeWJ2ZERDUDBEL2ozOUFRRFltajVjY1RXRUt6cGxCZ0FBNERsTXRBQUFpTk5DeUFJQVlDdUVLd0FBQUhnVkV5MEFBQ0pUMXUxaENPRlhmUUVBMkJqaENnQUFBRGJHUkFzQWdQaGM2QWtBd0tzSlZ3QUFBTEFWZ2hZQUFCRXA2L1lraEhDc0p3QUFQMFM0QWdBQWdLMFR0QUFBaUl0cEZnQUFMeU5jQVFBQXdFNEpXZ0FBUktLczIwVUk0VUEvQUFDKzYzRUlxQXBYQUFBQXNITnYxdXUxcWdNQWpLeXMyLzBRd2wwSVlVOHZBQUMrNkhHWVhISFZWTVdWRWdFQUFEQVdFeTBBQU9Ld0VySUFBUGdUNFFvQUFBQ2lZNklGQU1ESXlybzlDaUg4VXg5STNGa0k0VENFY0NwVUJNQjNDRmNBQUFBUU5STXRBQURHZDY0SEpPNm1xWXJWNTdkWTF1M3BFTGdRdWdEZ1MvcWZEUThoaEUrcUF3QUFRSXhNdEFBQUdORnc0UHovOUlERS9iV3BpdTZQYjdHczIvMG5nWXQzUGdRQS9NRjlDR0hlVklYQUJRQUFBRkVSdEFBQUdGRlp0M2NoaEFNOUlHR1hUVlhNdi9mMmhDNEErSVlQSVlSVlV4VVBpZ1FBQUVBTUJDMEFBRVpTMW0yL1NtR3AvaVNzMzdGLytOS0RzU2VoaTBVSTRhMFBDQUFoaE50aHVzV2ZKaVFCQUFEQXJnbGFBQUNNb0t6Ynd4QkNOK3dnaDFTZE5WV3hlczE3Ry82dDlLR0x1ZEFGQUp2NDJRSUFBQUN2SldnQkFEQ0NzbTR2UWdqdjFaNkUzVGRWY2JqSnR5ZDBBY0RBZEFzQUFBQkdKV2dCQUxCalpkMmVoQkQrb2U0azdtOU5WWHphMWxzczYvWm9DRnowd1lzREh5YUE3UFRycVZaTlZaeHJQUUFBQUxzbWFBRUFzR05sM2ZhSHo4ZnFUc0p1bXFvNDJkWGJFN29BeU5yTk1OM2lMdmRDQUFBQXNEdUNGZ0FBTzFUV2JYOFkvSGMxSjNGL0dldkFhd2hkTEliUXhaNFBHa0FXSG9ld3haVjJBd0FBc0F1Q0ZnQUFPMUxXN1g0SW9YUGpuc1I5YUtwaUVjTmJMT3YyZEFoY0NGMEE1T0Y2Q0Z3ODZEY0FBQURiSkdnQkFMQWpaZDJ1UWdoTDlTWmgvWTNpd3hnUHVJUXVBTEp4UDRRdFBtazVBQUFBMnlKb0FRQ3dBMlhkSG9ZUWZsVnJFdmRMVXhYbk1iL0ZZYkxNNThERnV3aGVFZ0RiY2RaVXhVcHRBUUFBMkFaQkN3Q0FIU2pyOXNxaExvbTdiYXJpYUVwdlVlZ0NJSG0zL2ZmNHBpcnV0Qm9BQUlCTkVyUUFBTml5c201UFFnai9VR2NTOTdjcGoybC9FcnFZaHhDT0kzaEpBR3hHdjlacTBWVEZoWG9DQUFDd0tZSVdBQUJiVnRadEYwSjRxODRrN0xxcGl0TlUzdDZ3NnVkejZNSy9YWUEwWFBmZjE1dXFlTkJQQUFBQVhrdlFBZ0JnaThxNjdROXEvNjdHSk80dnFZNWxGN29BU01yOXNFcWswMVlBQUFCZVE5QUNBR0JMaGxVRS9lSHpuaHFUc0xPbUtsWTVOSGdJWFN5RzRNVkJCQzhKZ0IrVHpjOHVBQUFBdGtQUUFnQmdTOHE2UFE4aC9LeStKS3pmZTMrWTR4ajJzbTZQaGlrWFFoY0EwM1F6VExld1NnUUFBSUFYRTdRQUFOaUM0ZWI3cjJwTDRuNXFxdUlpOXlZL0NWM01UYkFCbUpUSElXenhTZHNBQUFCNENVRUxBSUF0S091MmYyQi9yTFlrN0thcGloTU4vcjJ5YmsrSEtSZW5RaGNBazJHVkNBQUFBQzhpYUFFQXNHRmwzZmFIei85UVZ4TDNOemVBdjAzb0FtQlNyQklCQUFEZzJRUXRBQUEyckt6YnV4RENnYnFTc011bUt1WWEvSHhsM2M2SHdNVzdxYnhtZ0F4WkpRSUFBTUN6Q0ZvQUFHeFFXYmVMRU1ML3FDa0o2dytoanBxcXVOUGtseXZyZHYvSmxBdWhDNEE0V1NVQ0FBREFOd2xhQUFCc3lIQ0FlbWRGQUlseitMUWhUMElYL2JTTDR5VGVGRUE2ckJJQkFBRGdxd1F0Z1AvUDN0MGsxWEdrYlFOT085NDVmQ3NBcjBCNEJhQkpUb1dqSW5JcXRBTGpGZmhvQlkxV1lKaFdSTVlyVFhOaVdFSERDZ3c3RUN2Z2kvS2Jkc3R1U1M1K0RxY3E2N29pdklCK2JocU9UdDc1SkFCUHBPdkxhUWpodFhuU3NKdWM0cTZBbjE3WGw5MVBTaGN2V3Z2ZkJ6QlRuaElCQUFEZ3N4UXRBQUNlUU5lWHZSREN2ODJTeHYyUVUzd3Y1UFZTdWdDWW5KOXlpaWRpQVFBQTRBK0tGZ0FBVDZEcnk3blYvelR1SXFkNElPVG5WVXNYeDdWNHNiT2svKzBBRS9OaEtNQjVTZ1FBQUlDZ2FBRUE4SGhkWDRZRDBQODFTaHIzZlU3eFVzaWJVemZuSENsZEFHek1UWDFLeE45REFBQ0FoVk8wQUFCNGhLNHYyeUdFUzRlZU5PNHNwM2drNU9uNHBIUXgvTGUxOUhrQVBLUGJZZE5RVHZIVTBBRUFBSlpMMFFJQTRCRzZ2cXhDQ0QrYklRMGJEcFIyclVxZnJycFY1NC8vbEM0QW5zZTduT0t4V1FNQUFDeVRvZ1VBd0FOMWZkbXQyeXdjYk5LeW4zS0tKeEtlQjZVTGdHZDFGVUk0VUVZRUFBQllIa1VMQUlBSDZ2b3lySXgrYlg0MDdDYW51Q3ZnZWVyNmNsUUxGNitXUGd1QU5icXRaWXRMUXdZQUFGZ09SUXNBZ0FmbytuSVFRdmpWN0dqY3k1eml1WkRucmV2TDlpZGJMcFF1QU5ialRVN3gxR3dCQUFDV1FkRUNBT0FCdXI0TWg4Lzdaa2ZETG5LS0J3SnV5eWVsaXlPL3d3Q2UzRmxPOGNoWUFRQUEycWRvQVFCd1QzVWQveS9tUnVPK3l5bGVDN2xkWFY5MlB5bGR2Rmo2UEFDZXlGVjlTdVNqZ1FJQUFMUkwwUUlBNEI3cWJmRGg4SG5MM0dqWXU1emlzWUNYUStrQzRFbmQxckxGcGJFQ0FBQzBTZEVDQU9BZXVyNnNRZ2cvbXhrTkd3NkhkdDNFWGE1YXVqaXV4WXVkcGM4RDRJR0d2NmZIT2NWVEF3UUFBR2lQb2dVQXdFajE4UEUzODZKeGJ4d0s4WWV1TDN0MXk0WFNCY0REMkJJRkFBRFFJRVVMQUlDUnVyNjhEeUc4TWk4YWRwVlQzQk13bi9OSjZlTEk4MGtBOS9KaCtOMXBXeFFBQUVBN0ZDMEFBRWJvK25JUVF2alZyR2pjeTV6aXVaRDVKMTFmRHV1V2kwT2xDNEJScm9iZm1UbkZhK01DQUFDWVAwVUxBSUFSdXI1Y2hoQmVtQlVOKzVCVFBCUXc5NlYwQVREYWJRamhJS2Q0YVdRQUFBRHpwbWdCQVBBUHVyNE03MnIveTV4bzJIRHdzK2VXTFkvVjllV29GaTQ4c3dUd1pXOXlpcWZtQXdBQU1GK0tGZ0FBWDlIMVpUdUVjTzJXTm8xN20xTmNDWm1uVW45M0hpcGRBSHpSdTV6aXNmRUFBQURNazZJRkFNQlhkSDA1Q1NIOGFFWTA3S1p1cy9nb1pOWkI2UUxnaXo2RUVJNzhEUVlBQUpnZlJRc0FnQy9vK3JJYlF2ak5mR2ljOWVVOG0vcDdkU2hjREUrTXZEQjVnSEFWUWpoUXRnQUFBSmdYUlFzQWdDL28rbkllUXRnM0h4cDJrVk04RURDYm9IUUI4S2ZiV3JhNE5CSUFBSUI1VUxRQUFQaU1yaS9ENGZPdlprUGpYdVlVejRYTXB0WFN4Vkg5YjBjZ3dBSU5aWXREZjVjQkFBRG1RZEVDQU9BenVyNWNPK3lqY1djNXhTTWhNelZkWC9acTRlTFE3MkZnZ1R6cEJRQUFNQU9LRmdBQWY5UDFaUlZDK05sY2FOaHdhM2JYZS9CTW5kSUZzRkR2Y29ySHdnY0FBSmd1UlFzQWdFOTBmZGtPSVF6YkxMYk1oWWE5elNtdUJNeWNkSDA1cklXTFE3K2pnUVd3ZVFvQUFHRENGQzBBQUQ3UjlXVlkxZnphVEdqWVRVNXhWOERNbWRJRnNCQlhJWVFERzZnQUFBQ21SOUVDQUtDcUsrci9iUjQwN29lYzRuc2gwNHBhdWhodWZiOFNLdEFnWlFzQUFJQUpVclFBQUtpNnZweUhFUGJOZzRaZDVCUVBCRXlMNnROUGYyeTVVTG9BV25KYnl4YVhVZ1VBQUpnR1JRc0FnUDg3b0J0dVEvOWlGalR1ZTRjMExJSFNCZEFnWlFzQUFJQUpVYlFBQUJhdkhzZ05YMXJ2TEgwV05PMWRUdkZZeEN4TjE1ZmRXcmdZQ25Vdi9BQUFNemFVTFk1emlxZENCQUFBMkN4RkN3Qmc4YnErckVJSVB5OTlEalJ0T0pqWjliNDdTNmQwQVRUaWpiSUZBQURBWmlsYUFBQ0xWZy9kaG0wV1cwdWZCVTM3S2FkNEltTDRqL3I3LzZqK1o2TVJNRGYrdGdNQUFHeVFvZ1VBc0doZFg0YmJnSytYUGdlYWRwTlQzQlV4ZkZuWGw3MWF1RGhVdWdCbTVDeW5lQ1F3QUFDQTU2ZG9BUUFzVnRlWGd4RENyMzRDYU56TG5PSzVrR0VjcFF0Z1pwUXRBQUFBTmtEUkFnQllySzR2bDk3b3AzRVhPY1VESWNQRGRIMDVySVdMUTA5TUFSTjJGa0k0emlsK0ZCSUFBTUR6VUxRQUFCYXA2OHR3OCs4WDZkTzQ3M0tLMTBLR3g2dC9ONGJDeFN2akJDYm9Lb1J3b0d3QkFBRHdQQlF0QUlERjZmcXlIVUs0ZGp1WnhyM05LYTZFREUrci9nMDVxdi9aaWdSTWliSUZBQURBTTFHMEFBQVdwK3ZMY1BqOHMrUnAyRzBJWWRkQkM2eFgxNWU5VDBvWHludkFGQ2hiQUFBQVBBTkZDd0JnVWJxKzdJWVFmcE02alh1VFV6d1ZNandmVDRzQUU2SnNBUUFBc0dhS0ZnREFvblI5ZWU4UWpNWmQ1UlQzaEF5YlVRdDlmMnk1MkJFRHNDRzN0V3h4S1FBQUFJQ25wMmdCQUN4RzE1ZURFTUt2RXFkeEwzT0s1MEtHemV2NmNsZ0xGd3Ard0NZb1d3QUFBS3lKb2dVQXNCaGRYNjdkTHFaeFp6bkZJeUhEdE5oeUFXeVFzZ1VBQU1BYUtGb0FBSXZROWVVNGhQQXZhZE93NFNCbEw2ZDRMV1NZcnE0dmZ4UXU5c1VFUEJObEN3QUFnQ2VtYUFFQU5LL3J5M1lJWVRoODNwSTJEWHViVTF3SkdPYWhicms0cnFVTGY1K0FkVk8yQUFBQWVFTGZHaVlBc0FBcmgxZzA3aWFFY0NKa21JOWgrMHhPY1NoYURJV0xOL1gveHdEck1ud1dQdS82c21mQ0FBQUFqMmVqQlFEUXRIcGorRGNwMDdnM09jVlRJY084ZFgwNXFGc3VYb2tTV0JPYkxRQUFBSjZBb2dVQTBMU3VMK2Zld2FkeEZ6bkZBeUZET3p3ckFxeVpzZ1VBQU1BaktWb0FBTTNxK25JWVF2aGZDZE80N3gyVVFKdTZ2bXlIRUE3ckUxZzdZZ2Fla0xJRkFBREFJeWhhQUFETjZ2cHk3V0NLeHAzbEZJK0VETzJyNWNGalc1cUFKNlJzQVFBQThFRGZHaHdBMEtLdUwyNy8wcnJiZXVnS0xFQk84WDE5SnVqN29XUWxjK0FKREU4VG5YZDkyVE5NQUFDQSs3SFJBZ0JvVGwyMWZ1MWRleHIzTnFlNEVqSXNVOWVYM1JEQ1VTMWMrWHNIUEliTkZnQUFBUGVrYUFFQU5LZnJ5MmtJNGJWa2FkaE5UbkZYd0VBdEYvNVJ1TERKQ1hnb1pRc0FBSUI3VUxRQUFKclM5V1ZZcS82clZHbmNEOE16QWtJR1B0WDE1WS9DeFF1REFSNUEyUUlBQUdBa1JRc0FvQ2xkWDg1RENQdFNwV0VYT2NVREFRTmZVa3VISzM4UGdRZFF0Z0FBQUJoQjBRSUFhRWE5eWZ1TFJHbmNkem5GYXlFRC82UVdMbzQ4cHdYYzAxQzIyUE41QXdBQTRNc1VMUUNBSnRRMzZpKzlUMC9qM3VVVWo0VU0zRWZYbDkyNjRVTGhBaGpycW02MitHaGlBQUFBLysxYk13RUFHbkdzWkVIamJ1dEJLY0M5RExmU2M0ckRab3Z2UWdobjlmY0p3TmU4Q0NHYzF6SXpBQUFBZjJPakJRQXdlL1dtN3JETllrdWFOT3lubk9LSmdJSEhxZ2VueC9VL2Z6dUJyN0haQWdBQTRETVVMUUNBMmV2NjhqNkU4RXFTTk93cXA3Z25ZT0FwS1Z3QUl5bGJBQUFBL0kyblF3Q0FXZXY2Y3FCa3dRSWNDeGw0YXNPaGFVNXhlSkpvMkF6MTFwTWl3QmNNejRqWXFnVUFBUEFKR3kwQWdGbnIrbkpadi95RlZuM0lLUjVLRjFnM0d5NkFmM0NXVXp3eUpBQUFBRVVMQUdER3VyNE1YL1QrSWtNYTkxMU84VnJJd0hPcGhZdWg0RFZzdTlneGVPQVR5aFlBQU1EaUJVVUxBR0N1NmlIUXRSdTNOTzV0WGVzUHNCRzExS2h3QVh6cXA1eWlwMFFBQUlCRiszYnBBd0FBWnN0YWMxcDM2ejEwWU5OeWlxYzV4ZDJoK0ZWL0x3SDhxNWF3QUFBQUZzdEdDd0JnZHJxK0RBYyt2MG1PeHIwWkRqaUZERXhGM1NaMXJPd0lWRDZyQUFBQWk2Vm9BUURNVHRlWDh4REN2dVJvMkVWTzhVREF3QlFwWEFEVnNPWG1JS2Q0YVNBQUFNRFNLRm9BQUxQUzlXVTRmUDVWYWpUdVpVN3hYTWpBbENsY0FNb1dBQURBVWlsYUFBQ3owdlhsT29Td0l6VWFkcFpUOU80NU1Cc0tGN0I0TnlHRXZaeml4NlVQQWdBQVdJNXZaUTBBekVYWGwyTWxDeG8zM0FwZENSbVlrK0Z3TmFjNC9PN2FHOHBpd29QRkdUNmZuOWZTRlFBQXdDSW9XZ0FBczFDL3VIVUFUZXRPY29yWFVnYm1hUGo5VlRmeWZLZHdBWXZ6SW9Sd0tuWUFBR0FwRkMwQWdMbFlXVWRPNDRhMTJ5ZENCdWJ1YjRXTEM0SENZcnpxK3FKc0FRQUFMTUkzZDNkM2tnWUFKcTNyeTdDSy9OOVNvbkUvNUJUZkN4bG9UZGVYZzFxWTNCY3VMTUpQT1VYbFVRQUFvR21LRmdEQTVIVjlPWGM0UStNdWNvb0hRZ1phMXZYbHNHN3UyUkUwTk85TlR0RjJDd0FBb0ZtS0ZnREFwTlZEbWYrVkVvMzdQcWQ0S1dSZ0NicStITlVORndvWDBLN2JFTUtCenpjQUFFQ3J2cFVzQURCeDFnN1R1ak9IRU1DUzFGdnV3N05nYit0aExOQ2VyUkRDZWRlWFhka0NBQUF0VXJRQUFDYXI2NHZicnJSdU9HQThsakt3TkRuRmp6bkY0ZS84Y0FoNzVnY0FtalNVTGQ1M2Zka1dMd0FBMEJwRkN3QmdrdXJ0TndmUXRHNDFIRFpLR1ZpcVdyZ1luaEw1TG9SdzRRY0Jtdk1paEhBcVZnQUFvRFdLRmdEQVZLM3FMVGhvMVUxTzBkTTRBUDlYdUxqT0tSNkVFRjZHRUs3TUJKcnlxdXVMenp3QUFFQlRGQzBBZ01ucCtqSWN0THlXREkwN0VqREFYK1VVejNPS2V5R0VOME1oelhpZ0dUOTJmZkhaQndBQWFJYWlCUUF3UlN1cDBMaUw0VEJSeUFDZmwxTWNuaG9ZQ2hkdlF3aTN4Z1JOK0tYcnk1NG9BUUNBRm54emQzY25TQUJnTXVwTnQxOGtRdU8rRzlia0N4bmduM1Y5MmEwbFROdXVZUDZHNHRSdVR2R2pMQUVBZ0RtejBRSUFtSXl1TDl1MldiQUE3NVFzQU1ZYmZtZm1GSWNpNXN0aEk1RFJ3YXh0aFJCczlRSUFBR1pQMFFJQW1KTGpFTUtPUkdqWXJUSVJ3TU1NVHk3bEZBOUNDRzg4SndLejlxTHJ5NmtJQVFDQU9mTjBDQUF3Q1hVdCtHL1NvSEZ2Y29vT0ZnQWVxVzdCR2dxYVA1c2x6TlpQT2NVVDhRRUFBSE9rYUFFQVRFTFhsL2NoaEZmU29HRlhPY1U5QVFNOG5WclVIQXBzKzhZS3MvUnkyRllqT2dBQVlHNFVMUUNBamV2Nk1xd0IvMVVTTk01QkFzQ2ExTThTcDU0Z2c5a1puZ0hhelNsK0ZCMEFBREFuMzBvTEFKZ0FLNE5wM1FjbEM0RDFHWDdINWhTSDdSWnY2OEV0TUE5YklRU2ZrUUFBZ05sUnRBQUFOcXJyeTFFSTRZVVVhTnl4Z0FIV0w2ZTRDaUVNenpSOU1HNllqUmRkWDA3RkJRQUF6SW1uUXdDQWplbjZzaDFDdUs0MzJhQlZiK3ZCSHdEUHlITWlNRHR2Y29vS0Z3QUF3Q3pZYUFFQWJOSkt5WUxHM1hnYUIyQXovdmFjQ0RCOUoxMWY5dVFFQUFETWdZMFdBTUJHZEgwWkRqNStNMzBhNTJZbXdBVFV6eDNENytOOWVjQ2tYWVVRRG5LS0g4VUVBQUJNbVkwV0FNQ21PSHltZFJkS0ZnRFRrRk84emlrT1Q0bjhVTGNOQWRQMHdyOFRBQUNBT1ZDMEFBQ2VYWDB6M1kxU1dyZVNNTUMwNUJUZmh4Q0dwd25laVFZbTYxWFhsMlB4QUFBQVUrYnBFQURnMlhWOXVRNGg3Smc4RFR2TEtSNEpHR0M2dXI3czFadnpMOFFFay9SOVR2RlNOQUFBd0JUWmFBRUFQS3V1THlzbEN4cDNhNXNGd1BRTkI3ZzV4YUZzOFZQOTNRMU15L3V1TDlzeUFRQUFwa2pSQWdCNE52V0xVbXVBYWQxSlR2RmF5Z0R6a0ZNOHFjK0pYSWdNSm1XbmJwMEJBQUNZSEUrSEFBRFBwdXZMOEVYcGF4T25ZVGM1eFYwQkE4eFQxNWZEZXJDN0pVS1lqRGM1UllVTEFBQmdVbXkwQUFDZVJYMEhYY21DMXRuWUFqQmpPY1gzSVlTaE1IY21SNWlNay9wdkNRQUFnTW13MFFJQWVCWmRYODVEQ1B1bVRjTXVjb29IQWdab1E5ZVhnN3JkWWtla3NIRlhJWVNEbk9KSFVRQUFBRk5nb3dVQXNIWjFEYmVTQmEyenpRS2dJVG5Gb1NRNjNLSi9KMWZZdUJjaGhKVVlBQUNBcWJEUkFnQllxNjR2MnlHRVM3ZEJhZHk3bktLaUJVQ2piTGVBeWZpaFB2RURBQUN3VVRaYUFBRHJkdXhRZ3NiZHVtRUowTFpQdGx1OEZUVnMxR2t0Y2dNQUFHeVVvZ1VBc0RaZFgzWTlwOEFDckx3WER0Qys0WGQ5VG5FbzFuMGZRcmdTT1d6RVZnakJSZ3NBQUdEakZDMEFnSFZhMVM5RG9WVTNPY1VUNlFJc1IwN3hNcWRvdXdWc3puN1hGMlZ1QUFCZ294UXRBSUMxcUcrWnZ6WmRHbmNrWUlCbHN0MENObXJWOVdWUEJBQUF3S1lvV2dBQTY3SXlXUnAzVWQvc0IyQ2hodTBXSVlTaFhQck96d0E4cTJGcjNxbVJBd0FBbTZKb0FRQTh1YTR2d3kzL2ZaT2xjYlpaQURDVUxUN21GSWRuREY0T1QwcVpDRHliRjExZmxMc0JBSUNOK09idTdzN2tBWUFuMC9WbE80UndYVytaUWF2ZTFwWHhBUENuK2psbytQdndvNm5Bcy9tK2JwY0JBQUI0TmpaYUFBQlA3VmpKZ3NiZGhoQk9oQXpBMzMyeTNlS0grdmNDV0Q5UGlBQUFBTTlPMFFJQWVESmRYM1pEQ0QrYktJMDdIZzdTaEF6QWwrUVUzNGNRaHM5Rkh3d0oxbTU0UWtRSkZnQUFlRmFlRGdFQW5relhsK0ZRNFpXSjByQ3JuT0tlZ0FFWXErdkxjWDFPeE1ZdldLK1hPY1Z6TXdZQUFKNkRqUllBd0pQbytuS2daTUVDSEFzWmdQdklLUTQzN1llUzNwWEJ3VnFkZG4zWk5tSUFBT0E1S0ZvQUFFL0Z1bDVhOThFdFNRQWVJcWQ0WFRjaXZUVkFXSnNkcFZnQUFPQzVLRm9BQUk5V1YySy9NRWthZHV1TGV3QWVLNmM0UENIeU1vUndZNWl3Rmo5M2ZmSE1Hd0FBc0hhS0ZnREFvOVQxdkN0VHBIRW53MjFrSVFQd1dIVTcwbkFRL01Fd1lTMU9qUlVBQUZnM1JRc0E0TEdHa3NXV0tkS3dHMC9qQVBDVWNvb2ZjNHFISVlRM2RXc1M4SFJlZEgxUkJBY0FBTmJxbTd1N094TUdBQjZrNjh0dUNPRTMwNk54YjNLS2JrWUNzQmIxbVlOVHo3REJrL3MrcDNocHJBQUF3RHJZYUFFQVBJYkRaMXAzb1dRQndEclZnK0NERU1JN2c0WW5aU01aQUFDd05vb1dBTUNEZEgwWjFsM3ZteDZOT3hZd0FPdFdueElaL3ViODRDa1JlREw3WFY5OGxnTUFBTlpDMFFJQWVDZzN4R2pkbVhYVEFEeW5uT0w3RU1Md2xNaVZ3Y09UV0hWOTJUWktBQURncVNsYUFBRDMxdlZsRlVMWU1Ua2FkbXViQlFDYmtGTzh6aW51ZVVvRW5zU1c1dzRCQUlCMVVMUUFBTzZsM2doekFFM3JUb1kxN2xJR1lGTThKUUpQNWxYWGx3UGpCQUFBbnBLaUJRQndYeWYxWmhpMDZpYW51Skl1QUp0V254STU4SlFJUE5xcEowUUFBSUNucEdnQkFJelc5V1ZZWS8zYXhHaWNqUzBBVEVaTzhiS1dMYzZrQWcrMjR6TWVBQUR3bEw2NXU3c3pVQUJnbEs0djV5R0VmZE9pWVJjNVJhdWxBWmlrcmk5SElZUmZwQU1QOWwxTzhkcjRBQUNBeDdMUkFnQVlwWDZ4cjJSQjY0NGtETUJVNVJSUFF3amZEODljQ1FrZTVOVFlBQUNBcDZCb0FRRDhvL3FlOGNxa2FOdzdOeHdCbUxyNmxNanduTnVGc09EZTlydStIQm9iQUFEd1dJb1dBTUFZeC9WZFkyalZyVElSQUhPUlUveFluN3A2SnpTNHQ1TmFKQWNBQUhnd1JRc0E0S3U2dnV6V29nVzBiRFVjV2trWWdEbkpLUTZmMGQ3VXdpQXd6bzUvM3dBQUFJK2xhQUVBL0pQaGx2K1dLZEd3bTV6aWlZQUJtS09jNG1rSVlkaHVjU05BR08yNEZzb0JBQUFlUk5FQ0FQaWlyaS9EbC9hdlRZakdIUWtZZ0RuTEtWNkdFUFpDQ0JlQ2hGRzJQQnNIQUFBOGhxSUZBUEExYnZuVHVnODV4WE1wQXpCM3d4TllPY1doSlB0T21EREs2MW9zQndBQXVEZEZDd0RnczdxK0RMZjhYNWdPamZNK053Qk55U2tPZjl2ZVNCVkdzZFVDQUFCNEVFVUxBT0MvZEgzWnRzMkNCWGliVTd3V05BQ3R5U21laGhDK0R5SGNDaGUrYXI4V3pBRUFBTzVGMFFJQStKemorbTR4dE9wV21RaUFsdVVVTDBNSWV5R0VLMEhEVjYxcTBSd0FBR0EwUlFzQTRDKzZ2dXlHRUg0MkZScDNQTHhsTDJRQVdsWTNOeDJFRUQ0SUdyNW94M055QUFEQWZYMXpkM2RuYUFEQW43cSt2QThodkRJUkduYVZVOXdUTUFCTDB2VmwyT1QwbzlEaHM0WnRaN3VLdUFBQXdGZzJXZ0FBZityNmNxQmt3UUs0c1FqQTR1UVVoNzkvYnlRUG43WGxXVGtBQU9BK0ZDMEFnRStkbWdhTk84c3BuZ3NaZ0NYS0tRNmY5VjdXMi92QVg3MnV6eWdDQUFEOEkwVUxBT0IzWFYrTzYvdkUwS3JoVUdrbFhRQ1dyQllPRDVRdDRMTjhWZ1FBQUVaUnRBQUFocExGdGk4VldZQ1RuT0sxb0FGWXVwemlaUWhodUxsL3RmUlp3Tis4cnM4cEFnQUFmSldpQlFBUWFzbGl5eVJvMkJaTkh0VUFBQmJVU1VSQlZJMTN0d0hnUDNLS0grdG1pd3RqZ2I5UVFBY0FBUDdSTjNkM2Q2WUVBQXZXOVdVdmhQQnZQd00wN29lYzRuc2hBOEIvNi9weU90emtOeHI0MDh2NnpBNEFBTUJuMldnQkFMamxUK3N1bEN3QTRNdHlpa2NoaEhkR0JIK3kxUUlBQVBncVJRc0FXTEN1TDRjaGhIMC9BelR1V01BQThIVTV4ZUh2NVJ0amd0L3QxMzhyQVFBQWZKYWlCUUFzbTIwV3RPNHNwM2dwWlFENFp6bkZVMlVMK0pOL0t3RUFBRitrYUFFQUM5WDFaVmlIdXlOL0duWnJtd1VBM0U4dFczeGYvNDdDa3UxMGZUbnlFd0FBQUh5T29nVUFMRkRYbDIwSDBDekFTVTd4bzZBQjRIN3FOcWdEWlFzSUt5TUFBQUErUjlFQ0FKWnBXSU83SlhzYWRwTlQ5TVU0QUR5UXNnWDh6bFlMQUFEZ3N4UXRBR0JodXI0TVg1aS9sanVOODRVNEFEeFNMVnZzaFJDdXpKSUZVOTRGQUFEK2k2SUZBQ3lQTHdwcDNVVk84VnpLQVBCNE9jWHJ1dGxDMllLbHN0VUNBQUQ0TDRvV0FMQWc5UXZDZlpuVE9GK0VBOEFUeWlsK1ZMWmc0WlRWQVFDQXYxQzBBSUNGNlBxeTdRdENGdUJkdlhrTEFEd2haUXNXYnRocWNiajBJUUFBQVAraGFBRUF5M0U4ZkVFb2J4cDJxMHdFQU91amJNSENIUzk5QUFBQXdIOG9XZ0RBQW5SOTJmWEZJQXV3cWdkQUFNQ2FLRnV3WVB0ZFh3NzhBQUFBQUVIUkFnQVc0eVNFc0NWdUduYVZVendSTUFDczN5ZGxpdy9HemNMWW5nWUFBUHhPMFFJQUdsZHZYYjJTTTQyenNRVUFudEZRdHNncEhvWVF6c3lkQmJIVkFnQUErSjJpQlFDMHp5MS9XdmNocDNndVpRQjRmam5GSTJVTEZ1Wkk0QUFBZ0tJRkFEU3M2OHZ3SmVBTEdkTTQyeXdBWUlOcTJlSkNCaXpFNjY0dnU4SUdBSUJsVTdRQWdFWjFmZG0yellJRmVKdFR2QlkwQUd6YzhJeklsUmhZaUpXZ0FRQmcyUlF0QUtCZHd5My9MZm5Tc0J0bElnQ1locHppeHhEQ2diSUZDL0c2RnRzQkFJQ0ZVclFBZ0FiVlZiWS95NWJHcmVxaERnQXdBY29XTEl6bjZ3QUFZTUVVTFFDZ1RhZHlwWEVYT1VVLzV3QXdNY29XTElpaUJRQUFMSmlpQlFBMHB1dkw4TVgydmx4cG5IZXhBV0NpbEMxWWlLMnVMMGZDQmdDQVpWSzBBSUQydU9WUDY4NXlpdWRTQm9EcCtxUnNjU01tR21hckJRQUFMSlNpQlFBMHBPdkw4RVhmamt4cDJLMXRGZ0F3RDdWc2NWai9ma09MWHRTTmdnQUF3TUlvV2dCQUk3cStiRHVBWmdGT2NvclhnZ2FBZWNncFh0Yk5Gc29XdE1yeklRQUFzRUNLRmdEUWpwUGhuV0I1MHJDYm5LSXlFUURNakxJRmpYdmQ5V1ZYeUFBQXNDeUtGZ0RRZ0s0dmU4TVhmTEtrY2Q3QUJvQ1pxbVVMZjh0cGxhMFdBQUN3TUlvV0FOQ0dFem5TdUl1YzRuc2hBOEI4NVJSUFF3aHZSRWlEbElnQUFHQmhGQzBBWU9hNnZoeUdFUGJsU09OOGVRMEFEYWhsaTdleXBERmJYVjlzdFFBQWdBVlJ0QUNBK2JQTmd0YWQxWFhqQUVBRGNvcXI0ZSs3TEdtTW9nVUFBQ3lJb2dVQXpGalhsK0ZMNmgwWjByQmIyeXdBb0QwNXhlRlEra0swTkdTLzY4dXVRQUVBWUJrVUxRQmdwdXFYZUE2Z2FkMHFwL2hSeWdEUXBPRUp2Q3ZSMGhEL1BnTUFnSVZRdEFDQStScTJXV3pKajRiZDVCUTlqUU1BamFwbHlzTzZ3UXBhNFBrUUFBQllDRVVMQUppaHJpOEhJWVRYc3FOeHZxZ0dnTWJsRks5RENBZktGalJpcSt1THo3QUFBTEFBaWhZQU1FOHJ1ZEc0aTV6aXVaQUJvSDA1eFV0UEx0Q1FRMkVDQUVEN0ZDMEFZR2JxRGFsOXVkRTROd0VCWUVGeWlxY2hoTGN5cHdHdnVyN3NDaElBQU5xbWFBRUFNOUwxWmRzMkN4YmdYVjBqRGdBc1NFNXgrSno3UWVZMHdGWUxBQUJvbktJRkFNekxzRko1UjJZMDdGYVpDQUFXYmRocWRiWDBJVEI3bnNJQkFJREdmWE4zZHlkakFKaUJ1bjcyTjFuUnVEZDFkVGdBc0ZEMWMrOWxDR0hMendBejluMU84VktBQUFEUUpoc3RBR0ErVG1SRjQ2NlVMQUNBK29TWXB4ZVlPMXN0QUFDZ1lZb1dBREFEWFY4T1FnaXZaRVhqZkJrTkFQd3VwM2dlUXZqSk5KZ3haU0VBQUdpWW9nVUF6SU50RnJUdVF6MVFBUUQ0WFU1eCtBejh3VFNZcWEydUw4b1dBQURRS0VVTEFKaTRyaS9ETGY4WGNxSmh0N1paQUFCZmNCUkN1REVjWmtyUkFnQUFHdlhOM2QyZGJBRmdvcnErYkljUWhqZXF0MlJFdzk3bUZGY0NCZ0ErcCt2TFhnamgzR2RpWnVyLzVSUS9DZzhBQU5waW93VUFUTnZLRjhvMDdzYlRPQURBMStRVUwyMi9Zc1pzdFFBQWdBWXBXZ0RBUkhWOTJRMGgvQ2dmR3JkeXd3OEErQ2M1eGRNUXdwbEJNVU9LRmdBQTBDQkZDd0NZcmxQWjBMaUxlbWdDQURER3NOWGl5cVNZbVZmMVNVZ0FBS0FoaWhZQU1FRmRYdzVDQ1B1eW9YRXJBUU1BWTlVdFdFY0d4Z3paYWdFQUFJMVJ0QUNBYVhMTG45YWQ1UlRQcFF3QTNFZE84VEtFOEpPaE1UT0tGZ0FBMEJoRkN3Q1ltSzR2d3kzL0hiblFzTnU2K2hzQTRONXlpaWZERTJRbXg0eDRQZ1FBQUJxamFBRUFFMUsvZkhNQVRldE82dXB2QUlDSE9xemxUWmdMV3kwQUFLQWhpaFlBTUMzRDdid3RtZEN3bTV6aVNzQUF3R1BVMHVhUklUSWppaFlBQU5BUVJRc0FtSWl1TDNzaGhOZnlvSEUydGdBQVR5S24rRDZFY0dhYXpNUXJRUUVBUURzVUxRQmdPazVrUWVNdTZvRUlBTUJUR1VxY042YkpISFI5c2RVQ0FBQWFvV2dCQUJOUXYzRGJsd1dOczgwQ0FIaFNuaEJoWmhRdEFBQ2dFWW9XQUxCaFhWKzJiYk5nQWQ3bEZDOEZEUUE4dFp6aStmQlp3MkNaQVVVTEFBQm9oS0lGQUd6ZWNNdC9SdzQwN0RhRXNCSXdBTEJHSzArSU1BTmJYVi8yQkFVQUFQT25hQUVBRzlUMVpkZHpDaXpBcXE3MUJnQllDMCtJTUNPMldnQUFRQU1VTFFCZ3M0YWJkMXN5b0dFM09VVlA0d0FBYTFlZkVEa3phU1pPMFFJQUFCcWdhQUVBRzlMMTVTQ0U4TnI4YVp5YnBRREFjenF1ejViQlZMMm9tdzBCQUlBWlU3UUFnTTFabVQyTisxQnZsZ0lBUEl2NmhJaW4rWmk2QXdrQkFNQzhLVm9Bd0FaMGZSbHUrZStiUFkxenlBRUFQTHVjNG1rSTRjTGttVEJGQ3dBQW1EbEZDd0I0WmwxZnRrTUlKK1pPNDk3bUZLK0ZEQUJzaU9mTG1MSkQ2UUFBd0x3cFdnREE4eHR1K1crWk93MjdWU1lDQURhcEZqN2ZDb0dKMnVyNnNpY2NBQUNZTDBVTEFIaEdYVjkyUXdnL216bU5PNjd2b3dNQWJOSlEvTHlSQUJQbCtSQUFBSmd4UlFzQWVGNXUrZE82cS9vdU9nREFSdFhpNTdFVW1DalBod0FBd0l3cFdnREFNK242TXR4WWVtWGVOTTVoQmdBd0dUbkY5eUdFQzRrd1FmdENBUUNBK1ZLMEFJRG40NVkvclR2TEtaNUxHUUNZR0VWUUpxbVc4UUVBZ0JsU3RBQ0FaOUQxWmZoeWQ4ZXNhZGh0Q0dFbFlBQmdhbktLbDBNaFZEQk1rS0lGQUFETWxLSUZBS3haMTVkdEI5QXN3RWxPOFZyUUFNQkVyV294RktaRTBRSUFBR1pLMFFJQTFtLzRVbmZMbkduWXpWQzBFREFBTUZXMUVPcnpDbE96THhFQUFKZ25SUXNBV0tPdUw3c2hoQi9ObU1hdGNvb2ZoUXdBVE55SnJSWk1UZGNYV3kwQUFHQ0dGQzBBWUwxT3paZkdYZVFVL1p3REFKTlhpNkcyV2pBMWloWUFBREJEaWhZQXNDWmRYdzZ0Z21VQmpvVU1BTXhGVG5GVm56MkRxVkMwQUFDQUdWSzBBSUQxY1Z1TzFwM2xGQytsREFETXpFcGdUSWh5UGdBQXpKQ2lCUUNzUWRlWDRjdmJIYk9sWWJlMldRQUFjMVNmUGJQVmdzbm8rckluRFFBQW1CZEZDd0I0WWwxZnRoMUFzd0FuOVoxekFJQTVzdFdDS2ZGOENBQUF6SXlpQlFBOHZlSEprQzF6cFdFMzlYMXpBSUJac3RXQ2liSFJBZ0FBWmtiUkFnQ2VVRjM1K3RwTWFkeVJnQUdBQmlpT01oVTJXZ0FBd013b1dnREEwem94VHhwM2tWTThGeklBTUhkMXE4V3RJSm1BbmZvRUpRQUFNQk9LRmdEd1JMcStETGY4OTgyVHh0bG1BUUMwUkZHYXFmQjhDQUFBeklpaUJRQThnWHI3eU9waFd2Y3VwM2d0WlFDZ0lTZTJXakFSbmc4QkFJQVpVYlFBZ0tkeFBLeDdOVXNhZHF0TUJBQzBKcWY0TVlSd0tsZ213RVlMQUFDWUVVVUxBSGlrcmkrN3RXZ0JMVnZWZ3dnQWdOWjRQb1FwVUxRQUFJQVpVYlFBZ01jYmJ2bHZtU01OdThvcE9vQUFBSnBVbjBZN2t5NGJ0bE9mcEFRQUFHWkEwUUlBSHFIcnkvQ083bXN6cEhFMnRnQUFyZk44Q0ZOZ3F3VUFBTXlFb2dVQVBJNWIvclR1UTA3eFhNb0FRTXZxNTUwckliTmhCd0lBQUlCNVVMUUFnQWZxK25JVVFuaGhmalRPTmdzQVlDbVVxTm0wWFFrQUFNQThLRm9Bd0FQVXQzTjlFVXZyM3RZM3l3RUFsdUI5Q09GVzBteVFwME1BQUdBbUZDMEE0R0dHVy81YlprZkRicFdKQUlBbHlTbCtyR1VMMkJRYkV3RUFZQ1lVTFFEZ25ycStET3RjZnpZM0duZGNEeHNBQUpaRTBaU042dnBpcXdVQUFNeUFvZ1VBM04rcG1kRzRpNXlpbjNNQVlIRnlpcGNoaEJ2SnMwRzdoZzhBQU5PbmFBRUE5OUQxNVNDRXNHOW1ORzRsWUFCZ3dXeTFZSk5zdEFBQWdCbFF0QUNBKzNITG45YWQ1UlRQcFF3QUxOaDc0Yk5CaWhZQUFEQURpaFlBTUZMWGwrTVF3bzU1MGJCYjJ5d0FnS1hMS1Y2SEVENHNmUTVzaktkREFBQmdCaFF0QUdDRXJpL2JEcUJaZ0pONnNBQUFzSFMyV3JBcEwwd2VBQUNtVDlFQ0FNWVpTaFpiWmtYRGJyeEhEZ0R3SjBVTE5xYnJpNjBXQUFBd2NZb1dBUEFQdXI0TWIrVCthRTQwN2ppbitGSElBQUMvUHgveTBmTWhiSkNpQlFBQVRKeWlCUUQ4TTdmOGFkMUZUdEd0VFFDQXYvTDVpRTA1TUhrQUFKZzJSUXNBK0lxdUw0Y2hoSDB6b25ISEFnWUErQytLRm16S3Rza0RBTUMwS1ZvQXdOZlpaa0hyem5LS2wxSUdBUGlyK256SWxiR3dBWHVHRGdBQTA2Wm9BUUJmMFBWbEZVTFlNUjhhZG11YkJRREFWNTBhRHh0Z293VUFBRXpjL3dnSUFMN29Pb1R3MW5obzJHVzlxUWtBd09lOWQrZ05BQURBMzMxemQzZG5LQUFBQUFBQUFBQUFJM2c2QkFBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEErUC90Mm9FQUFBQUF3NkQ3VTE5aEFNVVJBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBVUd3N0VienhEVmFPOXBNQUFBQUFTVVZPUks1Q1lJST0nOwpmdW5jdGlvbiBhcHBseUJyYW5kaW5nKCkgewogIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJy5icmFuZC1sb2dvJykuZm9yRWFjaCgoaW1nKSA9PiB7IGltZy5zcmMgPSBMT0dPX0RBVEFfVVJJOyB9KTsKICBjb25zdCBmYXZpY29uID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2Zhdmljb25MaW5rJyk7CiAgaWYgKGZhdmljb24pIGZhdmljb24uaHJlZiA9IExPR09fREFUQV9VUkk7Cn0KCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBBcGkg4oCUIHRoaW4gZmV0Y2ggd3JhcHBlcnMgYXJvdW5kIHRoZSBSRVNUIEFQSS4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmNvbnN0IEFwaSA9ICgoKSA9PiB7CiAgYXN5bmMgZnVuY3Rpb24gcmVxdWVzdChwYXRoLCBvcHRpb25zKSB7CiAgICBjb25zdCByZXMgPSBhd2FpdCBmZXRjaChwYXRoLCBvcHRpb25zKTsKICAgIGxldCBib2R5OwogICAgdHJ5IHsKICAgICAgYm9keSA9IGF3YWl0IHJlcy5qc29uKCk7CiAgICB9IGNhdGNoIHsKICAgICAgYm9keSA9IG51bGw7CiAgICB9CiAgICBpZiAocmVzLnN0YXR1cyA9PT0gNDAxICYmICFwYXRoLnN0YXJ0c1dpdGgoJy9hcGkvYXV0aC8nKSkgewogICAgICB3aW5kb3cuZGlzcGF0Y2hFdmVudChuZXcgQ3VzdG9tRXZlbnQoJ2xyczpzaWduZWQtb3V0JykpOwogICAgfQogICAgaWYgKCFyZXMub2spIHsKICAgICAgY29uc3QgbWVzc2FnZSA9IChib2R5ICYmIGJvZHkuZXJyb3IpIHx8IGBSZXF1ZXN0IGZhaWxlZCAoJHtyZXMuc3RhdHVzfSlgOwogICAgICB0aHJvdyBuZXcgRXJyb3IobWVzc2FnZSk7CiAgICB9CiAgICByZXR1cm4gYm9keTsKICB9CgogIGZ1bmN0aW9uIHFzKHBhcmFtcykgewogICAgY29uc3QgdXNwID0gbmV3IFVSTFNlYXJjaFBhcmFtcygpOwogICAgT2JqZWN0LmVudHJpZXMocGFyYW1zIHx8IHt9KS5mb3JFYWNoKChbaywgdl0pID0+IHsKICAgICAgaWYgKHYgIT09IHVuZGVmaW5lZCAmJiB2ICE9PSBudWxsICYmIHYgIT09ICcnKSB1c3Auc2V0KGssIHYpOwogICAgfSk7CiAgICBjb25zdCBzID0gdXNwLnRvU3RyaW5nKCk7CiAgICByZXR1cm4gcyA/IGA/JHtzfWAgOiAnJzsKICB9CgogIHJldHVybiB7CiAgICBhdXRoTWU6ICgpID0+IHJlcXVlc3QoJy9hcGkvYXV0aC9tZScpLAogICAgYXV0aExvZ2luOiAoY29kZSkgPT4KICAgICAgcmVxdWVzdCgnL2FwaS9hdXRoL2xvZ2luJywgeyBtZXRob2Q6ICdQT1NUJywgaGVhZGVyczogeyAnQ29udGVudC1UeXBlJzogJ2FwcGxpY2F0aW9uL2pzb24nIH0sIGJvZHk6IEpTT04uc3RyaW5naWZ5KHsgY29kZSB9KSB9KSwKICAgIGF1dGhMb2dvdXQ6ICgpID0+IHJlcXVlc3QoJy9hcGkvYXV0aC9sb2dvdXQnLCB7IG1ldGhvZDogJ1BPU1QnIH0pLAoKICAgIGZpbHRlck9wdGlvbnM6ICgpID0+IHJlcXVlc3QoJy9hcGkvYW5hbHl0aWNzL2ZpbHRlci1vcHRpb25zJyksCiAgICBrcGlzOiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9rcGlzJHtxcyhwYXJhbXMpfWApLAogICAgcGxhdGZvcm1CcmVha2Rvd246IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL3BsYXRmb3JtLWJyZWFrZG93biR7cXMocGFyYW1zKX1gKSwKICAgIGNhbXBhaWduQnJlYWtkb3duOiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9jYW1wYWlnbi1icmVha2Rvd24ke3FzKHBhcmFtcyl9YCksCiAgICBjb250ZW50VHlwZUJyZWFrZG93bjogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9hbmFseXRpY3MvY29udGVudC10eXBlLWJyZWFrZG93biR7cXMocGFyYW1zKX1gKSwKICAgIG1ldHJpY09wdGlvbnM6IChwbGF0Zm9ybSkgPT4gcmVxdWVzdChgL2FwaS9hbmFseXRpY3MvbWV0cmljLW9wdGlvbnMke3FzKHsgcGxhdGZvcm0gfSl9YCksCiAgICBtZXRyaWNTdW1tYXJ5OiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9tZXRyaWMtc3VtbWFyeSR7cXMocGFyYW1zKX1gKSwKICAgIHRyZW5kOiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy90cmVuZCR7cXMocGFyYW1zKX1gKSwKICAgIHRvcFBvc3RzOiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy90b3AtcG9zdHMke3FzKHBhcmFtcyl9YCksCiAgICBjb21wYXJlOiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9jb21wYXJlJHtxcyhwYXJhbXMpfWApLAogICAgbW9udGhseTogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9hbmFseXRpY3MvbW9udGhseSR7cXMocGFyYW1zKX1gKSwKICAgIHF1YXJ0ZXJseTogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9hbmFseXRpY3MvcXVhcnRlcmx5JHtxcyhwYXJhbXMpfWApLAogICAgeXRkOiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy95dGQke3FzKHBhcmFtcyl9YCksCiAgICBwbGF0Zm9ybVJlcG9ydDogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9hbmFseXRpY3MvcGxhdGZvcm0tcmVwb3J0JHtxcyhwYXJhbXMpfWApLAoKICAgIHByZXZpZXdVcGxvYWQ6IChmaWxlKSA9PiB7CiAgICAgIGNvbnN0IGZvcm0gPSBuZXcgRm9ybURhdGEoKTsKICAgICAgZm9ybS5hcHBlbmQoJ2ZpbGUnLCBmaWxlKTsKICAgICAgcmV0dXJuIHJlcXVlc3QoJy9hcGkvdXBsb2Fkcy9wcmV2aWV3JywgeyBtZXRob2Q6ICdQT1NUJywgYm9keTogZm9ybSB9KTsKICAgIH0sCiAgICBjb21taXRVcGxvYWQ6IChwYXlsb2FkKSA9PgogICAgICByZXF1ZXN0KCcvYXBpL3VwbG9hZHMvY29tbWl0JywgewogICAgICAgIG1ldGhvZDogJ1BPU1QnLAogICAgICAgIGhlYWRlcnM6IHsgJ0NvbnRlbnQtVHlwZSc6ICdhcHBsaWNhdGlvbi9qc29uJyB9LAogICAgICAgIGJvZHk6IEpTT04uc3RyaW5naWZ5KHBheWxvYWQpLAogICAgICB9KSwKICAgIHVwbG9hZEhpc3Rvcnk6ICgpID0+IHJlcXVlc3QoJy9hcGkvdXBsb2Fkcy9oaXN0b3J5JyksCiAgICB1cGxvYWRFcnJvcnM6IChpZCkgPT4gcmVxdWVzdChgL2FwaS91cGxvYWRzLyR7aWR9L2Vycm9yc2ApLAogICAgdXBsb2FkUmF3Um93czogKGlkKSA9PiByZXF1ZXN0KGAvYXBpL3VwbG9hZHMvJHtpZH0vcmF3LXJvd3NgKSwKCiAgICBsaXN0UmVjb3JkczogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9yZWNvcmRzJHtxcyhwYXJhbXMpfWApLAogICAgcmVjb3Jkc1RhYmxlOiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL3JlY29yZHMvdGFibGUke3FzKHBhcmFtcyl9YCksCiAgICBnZXRSZWNvcmQ6IChpZCkgPT4gcmVxdWVzdChgL2FwaS9yZWNvcmRzLyR7aWR9YCksCiAgICB1cGRhdGVSZWNvcmQ6IChpZCwgdmFsdWVzKSA9PgogICAgICByZXF1ZXN0KGAvYXBpL3JlY29yZHMvJHtpZH1gLCB7CiAgICAgICAgbWV0aG9kOiAnUFVUJywKICAgICAgICBoZWFkZXJzOiB7ICdDb250ZW50LVR5cGUnOiAnYXBwbGljYXRpb24vanNvbicgfSwKICAgICAgICBib2R5OiBKU09OLnN0cmluZ2lmeSh7IHZhbHVlcyB9KSwKICAgICAgfSksCiAgICBkZWxldGVSZWNvcmRQb3N0OiAocG9zdElkKSA9PiByZXF1ZXN0KGAvYXBpL3JlY29yZHMvcG9zdC8ke3Bvc3RJZH1gLCB7IG1ldGhvZDogJ0RFTEVURScgfSksCiAgICBkZWxldGVSZWNvcmRQbGF0Zm9ybTogKHBvc3RJZCwgcGxhdGZvcm0pID0+CiAgICAgIHJlcXVlc3QoYC9hcGkvcmVjb3Jkcy9wb3N0LyR7cG9zdElkfS9wbGF0Zm9ybS8ke3BsYXRmb3JtfWAsIHsgbWV0aG9kOiAnREVMRVRFJyB9KSwKICAgIGR1cGxpY2F0ZVJlY29yZHNQcmV2aWV3OiAoKSA9PiByZXF1ZXN0KCcvYXBpL3JlY29yZHMvZHVwbGljYXRlcycpLAogICAgcmVzb2x2ZUR1cGxpY2F0ZVJlY29yZHM6ICgpID0+IHJlcXVlc3QoJy9hcGkvcmVjb3Jkcy9kdXBsaWNhdGVzL3Jlc29sdmUnLCB7IG1ldGhvZDogJ1BPU1QnIH0pLAogICAgd2lwZVVwbG9hZGVkUmVjb3JkczogKCkgPT4KICAgICAgcmVxdWVzdCgnL2FwaS9yZWNvcmRzL3dpcGUnLCB7CiAgICAgICAgbWV0aG9kOiAnUE9TVCcsCiAgICAgICAgaGVhZGVyczogeyAnQ29udGVudC1UeXBlJzogJ2FwcGxpY2F0aW9uL2pzb24nIH0sCiAgICAgICAgYm9keTogSlNPTi5zdHJpbmdpZnkoeyBjb25maXJtOiAnREVMRVRFJyB9KSwKICAgICAgfSksCgogICAgcmVzdG9yZUJhY2t1cDogKGZvcm0pID0+IHJlcXVlc3QoJy9hcGkvYmFja3VwL3Jlc3RvcmUnLCB7IG1ldGhvZDogJ1BPU1QnLCBib2R5OiBmb3JtIH0pLAoKICAgIGxpc3RGb2xsb3dlcnM6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvZm9sbG93ZXJzJHtxcyhwYXJhbXMpfWApLAogICAgZm9sbG93ZXJzR3Jvd3RoOiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2ZvbGxvd2Vycy9ncm93dGgke3FzKHBhcmFtcyl9YCksCiAgICBmb2xsb3dlcnNLcGlzOiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2ZvbGxvd2Vycy9rcGlzJHtxcyhwYXJhbXMpfWApLAogICAgc2F2ZUZvbGxvd2VyczogKHBheWxvYWQpID0+CiAgICAgIHJlcXVlc3QoJy9hcGkvZm9sbG93ZXJzJywgeyBtZXRob2Q6ICdQT1NUJywgaGVhZGVyczogeyAnQ29udGVudC1UeXBlJzogJ2FwcGxpY2F0aW9uL2pzb24nIH0sIGJvZHk6IEpTT04uc3RyaW5naWZ5KHBheWxvYWQpIH0pLAogICAgdXBkYXRlRm9sbG93ZXJzOiAoaWQsIHBheWxvYWQpID0+CiAgICAgIHJlcXVlc3QoYC9hcGkvZm9sbG93ZXJzLyR7aWR9YCwgeyBtZXRob2Q6ICdQVVQnLCBoZWFkZXJzOiB7ICdDb250ZW50LVR5cGUnOiAnYXBwbGljYXRpb24vanNvbicgfSwgYm9keTogSlNPTi5zdHJpbmdpZnkocGF5bG9hZCkgfSksCiAgICBkZWxldGVGb2xsb3dlcnM6IChpZCkgPT4gcmVxdWVzdChgL2FwaS9mb2xsb3dlcnMvJHtpZH1gLCB7IG1ldGhvZDogJ0RFTEVURScgfSksCiAgfTsKfSkoKTsKCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBTdGF0ZSAvIEZvcm1hdCAvIFRvYXN0IOKAlCBzaGFyZWQgYXBwIHN0YXRlICsgc21hbGwgdXRpbGl0aWVzLgogICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KY29uc3QgU3RhdGUgPSAoKCkgPT4gewogIGNvbnN0IHRvZGF5ID0gbmV3IERhdGUoKTsKICBjb25zdCBpc28gPSAoZCkgPT4gZC50b0lTT1N0cmluZygpLnNsaWNlKDAsIDEwKTsKICBjb25zdCB0aGlydHlEYXlzQWdvID0gbmV3IERhdGUodG9kYXkpOwogIHRoaXJ0eURheXNBZ28uc2V0RGF0ZSh0aGlydHlEYXlzQWdvLmdldERhdGUoKSAtIDI5KTsKCiAgY29uc3QgZmlsdGVycyA9IHsKICAgIGRhdGVGcm9tOiBpc28odGhpcnR5RGF5c0FnbyksCiAgICBkYXRlVG86IGlzbyh0b2RheSksCiAgICBwbGF0Zm9ybTogJ2FsbCcsCiAgICBjYW1wYWlnblR5cGU6ICdhbGwnLAogICAgY29udGVudFR5cGU6ICdhbGwnLAogIH07CgogIGNvbnN0IGxpc3RlbmVycyA9IFtdOwoKICByZXR1cm4gewogICAgZ2V0RmlsdGVyczogKCkgPT4gKHsgLi4uZmlsdGVycyB9KSwKICAgIHNldEZpbHRlcnMocGFydGlhbCkgewogICAgICBPYmplY3QuYXNzaWduKGZpbHRlcnMsIHBhcnRpYWwpOwogICAgICBsaXN0ZW5lcnMuZm9yRWFjaCgoZm4pID0+IGZuKHRoaXMuZ2V0RmlsdGVycygpKSk7CiAgICB9LAogICAgb25DaGFuZ2UoZm4pIHsKICAgICAgbGlzdGVuZXJzLnB1c2goZm4pOwogICAgfSwKICB9Owp9KSgpOwoKY29uc3QgRm9ybWF0ID0gewogIG51bWJlcihuKSB7CiAgICBpZiAobiA9PT0gbnVsbCB8fCBuID09PSB1bmRlZmluZWQpIHJldHVybiAn4oCUJzsKICAgIHJldHVybiBNYXRoLnJvdW5kKG4pLnRvTG9jYWxlU3RyaW5nKCdlbi1VUycpOwogIH0sCiAgY29tcGFjdChuKSB7CiAgICBpZiAobiA9PT0gbnVsbCB8fCBuID09PSB1bmRlZmluZWQpIHJldHVybiAn4oCUJzsKICAgIGNvbnN0IGFicyA9IE1hdGguYWJzKG4pOwogICAgaWYgKGFicyA+PSAxXzAwMF8wMDApIHJldHVybiBgJHsobiAvIDFfMDAwXzAwMCkudG9GaXhlZCgxKS5yZXBsYWNlKC9cLjAkLywgJycpfU1gOwogICAgaWYgKGFicyA+PSAxXzAwMCkgcmV0dXJuIGAkeyhuIC8gMV8wMDApLnRvRml4ZWQoMSkucmVwbGFjZSgvXC4wJC8sICcnKX1LYDsKICAgIHJldHVybiBgJHtNYXRoLnJvdW5kKG4pfWA7CiAgfSwKICAvKiogRGFzaGJvYXJkLXdpZGUgInByb2Zlc3Npb25hbCIgbnVtYmVyIGZvcm1hdDogcGxhaW4gdW5kZXIgMSwwMDA7IGNvbW1hLWdyb3VwZWQKICAgICAgdXAgdG8gMTAsMDAwOyBhYmJyZXZpYXRlZCAoSy9NKSBiZXlvbmQgdGhhdCDigJQgZS5nLiA4NTAsIDEsMjUwLCAxMi41SywgMTU2SywgMS4yNU0uICovCiAgc21hcnQobikgewogICAgaWYgKG4gPT09IG51bGwgfHwgbiA9PT0gdW5kZWZpbmVkKSByZXR1cm4gJ+KAlCc7CiAgICBjb25zdCBhYnMgPSBNYXRoLmFicyhuKTsKICAgIGlmIChhYnMgPCAxMDAwKSByZXR1cm4gYCR7TWF0aC5yb3VuZChuKX1gOwogICAgaWYgKGFicyA8IDEwMDAwKSByZXR1cm4gTWF0aC5yb3VuZChuKS50b0xvY2FsZVN0cmluZygnZW4tVVMnKTsKICAgIGlmIChhYnMgPCAxXzAwMF8wMDApIHJldHVybiBgJHsobiAvIDEwMDApLnRvRml4ZWQoMSkucmVwbGFjZSgvXC4wJC8sICcnKX1LYDsKICAgIHJldHVybiBgJHsobiAvIDFfMDAwXzAwMCkudG9GaXhlZCgyKS5yZXBsYWNlKC9cLj8wKyQvLCAnJyl9TWA7CiAgfSwKICBwZXJjZW50KG4pIHsKICAgIGlmIChuID09PSBudWxsIHx8IG4gPT09IHVuZGVmaW5lZCkgcmV0dXJuICfigJQnOwogICAgcmV0dXJuIGAke051bWJlcihuKS50b0ZpeGVkKDEpLnJlcGxhY2UoL1wuMCQvLCAnJyl9JWA7CiAgfSwKICBwY3QobikgewogICAgaWYgKG4gPT09IG51bGwgfHwgbiA9PT0gdW5kZWZpbmVkKSByZXR1cm4gJ+KAlCc7CiAgICBjb25zdCBzaWduID0gbiA+IDAgPyAnKycgOiAnJzsKICAgIHJldHVybiBgJHtzaWdufSR7bi50b0ZpeGVkKDEpfSVgOwogIH0sCiAgZGF0ZShpc29fKSB7CiAgICBpZiAoIWlzb18pIHJldHVybiAn4oCUJzsKICAgIGNvbnN0IFt5LCBtLCBkXSA9IGlzb18uc3BsaXQoJy0nKS5tYXAoTnVtYmVyKTsKICAgIHJldHVybiBuZXcgRGF0ZSh5LCBtIC0gMSwgZCkudG9Mb2NhbGVEYXRlU3RyaW5nKCdlbi1VUycsIHsgbW9udGg6ICdzaG9ydCcsIGRheTogJ251bWVyaWMnLCB5ZWFyOiAnbnVtZXJpYycgfSk7CiAgfSwKICBkdXJhdGlvbihzZWNvbmRzKSB7CiAgICBpZiAoc2Vjb25kcyA9PT0gbnVsbCB8fCBzZWNvbmRzID09PSB1bmRlZmluZWQpIHJldHVybiAn4oCUJzsKICAgIGNvbnN0IHMgPSBNYXRoLnJvdW5kKHNlY29uZHMpOwogICAgaWYgKHMgPCA2MCkgcmV0dXJuIGAke3N9c2A7CiAgICBpZiAocyA8IDM2MDApIHJldHVybiBgJHtNYXRoLmZsb29yKHMgLyA2MCl9bSAke3MgJSA2MH1zYDsKICAgIGNvbnN0IGggPSBNYXRoLmZsb29yKHMgLyAzNjAwKTsKICAgIGNvbnN0IG0gPSBNYXRoLnJvdW5kKChzICUgMzYwMCkgLyA2MCk7CiAgICByZXR1cm4gYCR7aH1oICR7bX1tYDsKICB9LAogIGRlbHRhQ2xhc3MobikgewogICAgaWYgKG4gPT09IG51bGwgfHwgbiA9PT0gdW5kZWZpbmVkKSByZXR1cm4gJ2ZsYXQnOwogICAgaWYgKG4gPiAwLjUpIHJldHVybiAndXAnOwogICAgaWYgKG4gPCAtMC41KSByZXR1cm4gJ2Rvd24nOwogICAgcmV0dXJuICdmbGF0JzsKICB9LAp9OwoKY29uc3QgVG9hc3QgPSB7CiAgc2hvdyhtZXNzYWdlLCB0eXBlID0gJ3N1Y2Nlc3MnKSB7CiAgICBjb25zdCByb290ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3RvYXN0Um9vdCcpOwogICAgY29uc3QgZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGVsLmNsYXNzTmFtZSA9IGB0b2FzdCAke3R5cGV9YDsKICAgIGVsLnRleHRDb250ZW50ID0gbWVzc2FnZTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQoZWwpOwogICAgc2V0VGltZW91dCgoKSA9PiBlbC5yZW1vdmUoKSwgNTAwMCk7CiAgfSwKfTsKCi8qKiBTYWZlbHkgYnVpbGRzIERPTSB0ZXh0IG5vZGVzIGZvciB1bnRydXN0ZWQgc3RyaW5ncyAoY2FwdGlvbnMsIGZpbGVuYW1lcywgcGxhdGZvcm0gbGFiZWxzIGZyb20gZGF0YSkuICovCmZ1bmN0aW9uIHRleHRFbCh0YWcsIHRleHQsIGNsYXNzTmFtZSkgewogIGNvbnN0IGVsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCh0YWcpOwogIGlmIChjbGFzc05hbWUpIGVsLmNsYXNzTmFtZSA9IGNsYXNzTmFtZTsKICBlbC5hcHBlbmRDaGlsZChkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZSh0ZXh0ID8/ICcnKSk7CiAgcmV0dXJuIGVsOwp9CgovKiogQSBwcmVtaXVtIGVtcHR5IHN0YXRlOiBpY29uICsgZXhwbGFuYXRpb24gKyBvcHRpb25hbCBhY3Rpb24sIGluc3RlYWQgb2YgYSBibGFuayBhcmVhLgogICAgSWNvbnMgcmVuZGVyIHZpYSB0aGUgcGFnZS13aWRlIE11dGF0aW9uT2JzZXJ2ZXIgdGhhdCBjYWxscyBsdWNpZGUuY3JlYXRlSWNvbnMoKSAoc2VlIGJvb3RzdHJhcCkuICovCmZ1bmN0aW9uIGVtcHR5U3RhdGUoeyBpY29uID0gJ2luYm94JywgdGl0bGUsIG1lc3NhZ2UsIGFjdGlvbkxhYmVsLCBvbkFjdGlvbiB9KSB7CiAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogIHdyYXAuY2xhc3NOYW1lID0gJ2VtcHR5LXN0YXRlJzsKICBjb25zdCBpY29uV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogIGljb25XcmFwLmNsYXNzTmFtZSA9ICdlbXB0eS1pY29uJzsKICBpY29uV3JhcC5pbm5lckhUTUwgPSBgPGkgZGF0YS1sdWNpZGU9IiR7aWNvbn0iIHN0eWxlPSJ3aWR0aDoyMnB4O2hlaWdodDoyMnB4OyI+PC9pPmA7CiAgd3JhcC5hcHBlbmRDaGlsZChpY29uV3JhcCk7CiAgaWYgKHRpdGxlKSB3cmFwLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgdGl0bGUsICdlbXB0eS10aXRsZScpKTsKICBpZiAobWVzc2FnZSkgd3JhcC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIG1lc3NhZ2UsICdlbXB0eS1tZXNzYWdlJykpOwogIGlmIChhY3Rpb25MYWJlbCAmJiBvbkFjdGlvbikgewogICAgY29uc3QgYnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBidG4uY2xhc3NOYW1lID0gJ2J0biBwcmltYXJ5JzsKICAgIGJ0bi50ZXh0Q29udGVudCA9IGFjdGlvbkxhYmVsOwogICAgYnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgb25BY3Rpb24pOwogICAgd3JhcC5hcHBlbmRDaGlsZChidG4pOwogIH0KICByZXR1cm4gd3JhcDsKfQoKLyoqIEEgPGJ1dHRvbj4gd2l0aCBhIHNtYWxsIGxlYWRpbmcgTHVjaWRlIGljb24gYmVmb3JlIGl0cyBsYWJlbCAobGFiZWwgaXMgYWx3YXlzIGEgc3RhdGljLCBkZXZlbG9wZXItc3VwcGxpZWQgc3RyaW5nIGF0IGNhbGwgc2l0ZXMsIG5ldmVyIHVzZXIgZGF0YSDigJQgaW5zZXJ0ZWQgdmlhIGNyZWF0ZVRleHROb2RlIHJlZ2FyZGxlc3MpLiAqLwpmdW5jdGlvbiBpY29uQnRuKGNsYXNzTmFtZSwgaWNvbk5hbWUsIGxhYmVsKSB7CiAgY29uc3QgYnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgYnRuLmNsYXNzTmFtZSA9IGNsYXNzTmFtZTsKICBjb25zdCBpY29uID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaScpOwogIGljb24uc2V0QXR0cmlidXRlKCdkYXRhLWx1Y2lkZScsIGljb25OYW1lKTsKICBpY29uLnN0eWxlLndpZHRoID0gJzEzcHgnOwogIGljb24uc3R5bGUuaGVpZ2h0ID0gJzEzcHgnOwogIGJ0bi5hcHBlbmRDaGlsZChpY29uKTsKICBidG4uYXBwZW5kQ2hpbGQoZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUoYCAke2xhYmVsfWApKTsKICByZXR1cm4gYnRuOwp9CgovKiogU2hpbW1lcmluZyBwbGFjZWhvbGRlcnMgc2hvd24gdGhlIGluc3RhbnQgYSBzZWN0aW9uIHN0YXJ0cyBsb2FkaW5nLCBzd2FwcGVkIGZvciByZWFsCiAgICBjb250ZW50IChvciBhbiBlbXB0eSBzdGF0ZSkgb25jZSB0aGUgZmV0Y2ggcmVzb2x2ZXMg4oCUIG5vIGJsYW5rIGFyZWFzIHdoaWxlIHdhaXRpbmcuICovCmZ1bmN0aW9uIHNrZWxldG9uU3RhdEdyaWQoY291bnQgPSA2KSB7CiAgY29uc3QgZ3JpZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogIGdyaWQuY2xhc3NOYW1lID0gJ3NrZWxldG9uLXN0YXQtZ3JpZCc7CiAgZm9yIChsZXQgaSA9IDA7IGkgPCBjb3VudDsgaSArPSAxKSB7CiAgICBjb25zdCB0aWxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0aWxlLmNsYXNzTmFtZSA9ICdza2VsZXRvbiBza2VsZXRvbi10aWxlJzsKICAgIGdyaWQuYXBwZW5kQ2hpbGQodGlsZSk7CiAgfQogIHJldHVybiBncmlkOwp9CmZ1bmN0aW9uIHNrZWxldG9uQ2hhcnQoKSB7CiAgY29uc3QgZGl2ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgZGl2LmNsYXNzTmFtZSA9ICdza2VsZXRvbiBza2VsZXRvbi1jaGFydCc7CiAgcmV0dXJuIGRpdjsKfQpmdW5jdGlvbiBza2VsZXRvblJvd3MoY291bnQgPSA2KSB7CiAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogIGZvciAobGV0IGkgPSAwOyBpIDwgY291bnQ7IGkgKz0gMSkgewogICAgY29uc3Qgcm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICByb3cuY2xhc3NOYW1lID0gJ3NrZWxldG9uIHNrZWxldG9uLXJvdyc7CiAgICB3cmFwLmFwcGVuZENoaWxkKHJvdyk7CiAgfQogIHJldHVybiB3cmFwOwp9CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgU2hhcmVkIGFuaW1hdGlvbiBwcmltaXRpdmVzIOKAlCBhIGNvdW50LXVwIGZvciBLUEkgbnVtYmVycyBhbmQgYQogICBDU1Mgd2lkdGgtdHJhbnNpdGlvbiBiYXIsIGJvdGggcmV1c2VkIGFjcm9zcyB0aGUgRGFzaGJvYXJkIGFuZAogICBDb21wYXJpc29ucyBwYWdlcy4gQm90aCByZXNwZWN0IHByZWZlcnMtcmVkdWNlZC1tb3Rpb24gKGd1YXJkZWQKICAgaW4gQ1NTLCBzZWUgLmJhci1maWxsIC8gdGhlIGFuaW1hdGVDb3VudCBkdXJhdGlvbiBjaGVjayBiZWxvdykuCiAgID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwpjb25zdCBQUkVGRVJTX1JFRFVDRURfTU9USU9OID0gd2luZG93Lm1hdGNoTWVkaWEgJiYgd2luZG93Lm1hdGNoTWVkaWEoJyhwcmVmZXJzLXJlZHVjZWQtbW90aW9uOiByZWR1Y2UpJykubWF0Y2hlczsKCi8qKiBTaHJpbmtzIGBlbGAncyBmb250IHNpemUganVzdCBlbm91Z2ggZm9yIGl0cyBjdXJyZW50IHRleHQgdG8gZml0IGl0cyBvd24gd2lkdGgg4oCUIGEgS1BJIHRpbGUncyBib3ggaXMgYSBmaXhlZCBzaXplLCBidXQgdGhlIHZhbHVlIGluc2lkZSBpdCBpc24ndCAoYSBmb2xsb3dlciBjb3VudCBjYW4gYmUgIjAiIG9yICIxLDA0OCw1NzYiKSwgc28gYSBzaW5nbGUgZml4ZWQgZm9udC1zaXplIHdpbGwgZXZlbnR1YWxseSBvdmVyZmxvdy4gUmVzZXRzIHRvIHRoZSBDU1MtZGVmaW5lZCBzaXplIGZpcnN0LCB0aGVuIHN0ZXBzIGRvd24gYnkgMXB4IGF0IGEgdGltZSB1bnRpbCBpdCBmaXRzIG9yIGhpdHMgYG1pblNpemVgLiAqLwpmdW5jdGlvbiBmaXRTdGF0VmFsdWUoZWwsIG1pblNpemUgPSAxOCkgewogIGlmICghZWwpIHJldHVybjsKICBlbC5zdHlsZS5mb250U2l6ZSA9ICcnOwogIGNvbnN0IG1heFdpZHRoID0gZWwuY2xpZW50V2lkdGg7CiAgaWYgKCFtYXhXaWR0aCkgcmV0dXJuOwogIGxldCBzaXplID0gcGFyc2VGbG9hdChnZXRDb21wdXRlZFN0eWxlKGVsKS5mb250U2l6ZSk7CiAgd2hpbGUgKGVsLnNjcm9sbFdpZHRoID4gbWF4V2lkdGggJiYgc2l6ZSA+IG1pblNpemUpIHsKICAgIHNpemUgLT0gMTsKICAgIGVsLnN0eWxlLmZvbnRTaXplID0gYCR7c2l6ZX1weGA7CiAgfQp9CgovKiogQW5pbWF0ZXMgYSBudW1iZXIgZnJvbSBgZnJvbWAgdG8gYHRvYCBpbnNpZGUgYGVsYCBvdmVyIGBkdXJhdGlvbmBtcywgZm9ybWF0dGluZyBlYWNoIGZyYW1lIHdpdGggYGZvcm1hdGAgKGRlZmF1bHRzIHRvIGEgcGxhaW4gcm91bmRlZCBpbnRlZ2VyKS4gU2tpcHMgc3RyYWlnaHQgdG8gdGhlIGZpbmFsIHZhbHVlIHVuZGVyIHByZWZlcnMtcmVkdWNlZC1tb3Rpb24uIFNocmlua3MgdGhlIGZvbnQgdG8gZml0IG9uY2UgdGhlIGZpbmFsIHZhbHVlIGxhbmRzLCBzaW5jZSB0aGUgYW5pbWF0ZWQgZGlnaXRzIGNhbiBiZSBhIGRpZmZlcmVudCB3aWR0aCB0aGFuIHRoZSBzZXR0bGVkIHZhbHVlIOKAlCBkZWZlcnJlZCBhIGZyYW1lIGJlY2F1c2UgYGVsYCBpcyB0eXBpY2FsbHkgc3RpbGwgZGV0YWNoZWQgZnJvbSB0aGUgZG9jdW1lbnQgKG1pZC1jb25zdHJ1Y3Rpb24gYnkgaXRzIGNhbGxlcikgd2hlbiBhbmltYXRlQ291bnQgaXMgZmlyc3QgaW52b2tlZCwgYW5kIGNsaWVudFdpZHRoIHJlYWRzIDAgdW50aWwgaXQncyBhY3R1YWxseSBhdHRhY2hlZCBhbmQgbGFpZCBvdXQuICovCmZ1bmN0aW9uIGFuaW1hdGVDb3VudChlbCwgZnJvbSwgdG8sIGR1cmF0aW9uID0gOTAwLCBmb3JtYXQpIHsKICBpZiAoIWVsKSByZXR1cm47CiAgY29uc3QgZm10ID0gZm9ybWF0IHx8ICgodikgPT4gTWF0aC5yb3VuZCh2KS50b0xvY2FsZVN0cmluZygnZW4tVVMnKSk7CiAgaWYgKFBSRUZFUlNfUkVEVUNFRF9NT1RJT04gfHwgZnJvbSA9PT0gdG8gfHwgIU51bWJlci5pc0Zpbml0ZShmcm9tKSB8fCAhTnVtYmVyLmlzRmluaXRlKHRvKSkgewogICAgZWwudGV4dENvbnRlbnQgPSBmbXQodG8pOwogICAgcmVxdWVzdEFuaW1hdGlvbkZyYW1lKCgpID0+IGZpdFN0YXRWYWx1ZShlbCkpOwogICAgcmV0dXJuOwogIH0KICBjb25zdCBzdGFydCA9IHBlcmZvcm1hbmNlLm5vdygpOwogIGZ1bmN0aW9uIHRpY2sobm93KSB7CiAgICBjb25zdCBlbGFwc2VkID0gbm93IC0gc3RhcnQ7CiAgICBjb25zdCBwcm9ncmVzcyA9IE1hdGgubWluKDEsIGVsYXBzZWQgLyBkdXJhdGlvbik7CiAgICBjb25zdCBlYXNlZCA9IDEgLSBNYXRoLnBvdygxIC0gcHJvZ3Jlc3MsIDMpOyAvLyBlYXNlT3V0Q3ViaWMKICAgIGVsLnRleHRDb250ZW50ID0gZm10KGZyb20gKyAodG8gLSBmcm9tKSAqIGVhc2VkKTsKICAgIGlmIChwcm9ncmVzcyA8IDEpIHsKICAgICAgcmVxdWVzdEFuaW1hdGlvbkZyYW1lKHRpY2spOwogICAgfSBlbHNlIHsKICAgICAgZml0U3RhdFZhbHVlKGVsKTsKICAgIH0KICB9CiAgcmVxdWVzdEFuaW1hdGlvbkZyYW1lKHRpY2spOwp9CgovKiogQSBsYWJlbGVkIGhvcml6b250YWwgYmFyIHRoYXQgYW5pbWF0ZXMgaXRzIHdpZHRoIGluIG9uIGluc2VydGlvbiDigJQgdXNlZCBmb3IgdGhlIENvbXBhcmlzb25zIHBhZ2UncyBwYWlyZWQgUmFuZ2UgQS9CIGJhcnMuIGB2YWx1ZWAvYG1heGAgZHJpdmUgdGhlIGZpbGwgcGVyY2VudGFnZTsgYGNvbG9yVmFyYCBpcyBhIENTUyBjdXN0b20gcHJvcGVydHkgbmFtZSAoZS5nLiAnLS1zZXJpZXMtMScpLiAqLwpmdW5jdGlvbiBidWlsZEJhcih7IGxhYmVsLCB2YWx1ZSwgbWF4LCBjb2xvclZhciwgZm9ybWF0VmFsdWUgfSkgewogIGNvbnN0IHJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogIHJvdy5jbGFzc05hbWUgPSAnYmFyLXJvdyc7CiAgY29uc3QgbGFiZWxFbCA9IHRleHRFbCgnZGl2JywgbGFiZWwsICdiYXItbGFiZWwnKTsKICBjb25zdCB0cmFjayA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogIHRyYWNrLmNsYXNzTmFtZSA9ICdiYXItdHJhY2snOwogIGNvbnN0IGZpbGwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICBmaWxsLmNsYXNzTmFtZSA9ICdiYXItZmlsbCc7CiAgZmlsbC5zdHlsZS5iYWNrZ3JvdW5kID0gY29sb3JWYXIgPyBgdmFyKCR7Y29sb3JWYXJ9KWAgOiAndmFyKC0tc2VyaWVzLTEpJzsKICB0cmFjay5hcHBlbmRDaGlsZChmaWxsKTsKICBjb25zdCB2YWx1ZUVsID0gdGV4dEVsKCdkaXYnLCBmb3JtYXRWYWx1ZSA/IGZvcm1hdFZhbHVlKHZhbHVlKSA6IFN0cmluZyh2YWx1ZSksICdiYXItdmFsdWUnKTsKICByb3cuYXBwZW5kKGxhYmVsRWwsIHRyYWNrLCB2YWx1ZUVsKTsKICBjb25zdCBwY3QgPSBtYXggPiAwID8gTWF0aC5taW4oMTAwLCBNYXRoLnJvdW5kKCh2YWx1ZSAvIG1heCkgKiAxMDAwKSAvIDEwKSA6IDA7CiAgcmVxdWVzdEFuaW1hdGlvbkZyYW1lKCgpID0+IHsgZmlsbC5zdHlsZS53aWR0aCA9IGAke3BjdH0lYDsgfSk7CiAgcmV0dXJuIHJvdzsKfQoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIFNoYXJlZCB0YWJsZSB0b29sYmFyIHBpZWNlcyDigJQgc2VhcmNoIGJveCwgY2xpZW50LXNpZGUgcGFnZXIsCiAgIGFuZCBDU1YvWExTWCBleHBvcnQg4oCUIHJldXNlZCBieSB0aGUgRm9sbG93ZXJzIERhdGEgYW5kIFVwbG9hZAogICBIaXN0b3J5IHRhYnMgKGJvdGggbG9hZCB0aGVpciBmdWxsIGRhdGFzZXQgb25jZSBhbmQgc2VhcmNoLwogICBzb3J0L3BhZ2luYXRlIGl0IGluIHRoZSBicm93c2VyLCB1bmxpa2UgRGF0YSBSZWNvcmRzIHdoaWNoIGlzCiAgIHNlcnZlci1wYWdpbmF0ZWQpLiBDU1YgbmVlZHMgbm8gc2VydmVyIHJvdW5kIHRyaXAgYXQgYWxsOyBYTFNYCiAgIGdvZXMgdGhyb3VnaCBQT1NUIC9hcGkvZXhwb3J0IHNvIGV4Y2VsanMgKGFscmVhZHkgYSBkZXBlbmRlbmN5KQogICBjYW4gZ2VuZXJhdGUgYSByZWFsIC54bHN4IHdpdGhvdXQgYWRkaW5nIGEgY2xpZW50LXNpZGUgbGlicmFyeS4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmZ1bmN0aW9uIGJ1aWxkU2VhcmNoQm94KHsgcGxhY2Vob2xkZXIsIHZhbHVlLCBvbkNoYW5nZSB9KSB7CiAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogIHdyYXAuY2xhc3NOYW1lID0gJ3JlY29yZHMtc2VhcmNoJzsKICBjb25zdCBpbnB1dCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7CiAgaW5wdXQudHlwZSA9ICdzZWFyY2gnOwogIGlucHV0LnBsYWNlaG9sZGVyID0gcGxhY2Vob2xkZXI7CiAgaW5wdXQudmFsdWUgPSB2YWx1ZSB8fCAnJzsKICBsZXQgZGVib3VuY2UgPSBudWxsOwogIGlucHV0LmFkZEV2ZW50TGlzdGVuZXIoJ2lucHV0JywgKCkgPT4gewogICAgY2xlYXJUaW1lb3V0KGRlYm91bmNlKTsKICAgIGRlYm91bmNlID0gc2V0VGltZW91dCgoKSA9PiBvbkNoYW5nZShpbnB1dC52YWx1ZSksIDMwMCk7CiAgfSk7CiAgd3JhcC5hcHBlbmRDaGlsZChpbnB1dCk7CiAgcmV0dXJuIHdyYXA7Cn0KCi8qKiBTbGljZXMgYW4gYWxyZWFkeS1sb2FkZWQsIGFscmVhZHktZmlsdGVyZWQvc29ydGVkIGFycmF5IGZvciBjbGllbnQtc2lkZSBwYWdpbmF0aW9uIOKAlCB0aGUgY291bnRlcnBhcnQgdG8gdGhlIHNlcnZlci1zaWRlIHBhZ2luYXRlKCkgaW4gYXBwLmpzLCBmb3IgdGFibGVzIHRoYXQgZG9uJ3QgaGF2ZSBhIHBhZ2luYXRlZCBlbmRwb2ludC4gKi8KZnVuY3Rpb24gcGFnaW5hdGVDbGllbnRTaWRlKHJvd3MsIHBhZ2UsIHBhZ2VTaXplKSB7CiAgY29uc3QgdG90YWxQYWdlcyA9IE1hdGgubWF4KDEsIE1hdGguY2VpbChyb3dzLmxlbmd0aCAvIHBhZ2VTaXplKSk7CiAgY29uc3Qgc2FmZVBhZ2UgPSBNYXRoLm1pbihNYXRoLm1heCgxLCBwYWdlKSwgdG90YWxQYWdlcyk7CiAgY29uc3Qgc3RhcnQgPSAoc2FmZVBhZ2UgLSAxKSAqIHBhZ2VTaXplOwogIHJldHVybiB7IHBhZ2VSb3dzOiByb3dzLnNsaWNlKHN0YXJ0LCBzdGFydCArIHBhZ2VTaXplKSwgdG90YWxQYWdlcywgc2FmZVBhZ2UsIHRvdGFsOiByb3dzLmxlbmd0aCB9Owp9CgpmdW5jdGlvbiBidWlsZFBhZ2VyKHsgcGFnZSwgdG90YWxQYWdlcywgdG90YWwsIG9uUHJldiwgb25OZXh0IH0pIHsKICBjb25zdCBwYWdlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogIHBhZ2VyLmNsYXNzTmFtZSA9ICdwYWdpbmF0aW9uLXJvdyc7CiAgY29uc3QgcHJldkJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogIHByZXZCdG4uY2xhc3NOYW1lID0gJ2J0bic7IHByZXZCdG4udHlwZSA9ICdidXR0b24nOyBwcmV2QnRuLnRleHRDb250ZW50ID0gJ1ByZXZpb3VzJzsKICBwcmV2QnRuLmRpc2FibGVkID0gcGFnZSA8PSAxOwogIHByZXZCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBvblByZXYpOwogIGNvbnN0IG5leHRCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICBuZXh0QnRuLmNsYXNzTmFtZSA9ICdidG4nOyBuZXh0QnRuLnR5cGUgPSAnYnV0dG9uJzsgbmV4dEJ0bi50ZXh0Q29udGVudCA9ICdOZXh0JzsKICBuZXh0QnRuLmRpc2FibGVkID0gcGFnZSA+PSB0b3RhbFBhZ2VzOwogIG5leHRCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBvbk5leHQpOwogIHBhZ2VyLmFwcGVuZChwcmV2QnRuLCB0ZXh0RWwoJ3NwYW4nLCBgUGFnZSAke3BhZ2V9IG9mICR7dG90YWxQYWdlc30g4oCUICR7dG90YWx9IHJlY29yZChzKWApLCBuZXh0QnRuKTsKICByZXR1cm4gcGFnZXI7Cn0KCmZ1bmN0aW9uIGRvd25sb2FkQmxvYihmaWxlbmFtZSwgbWltZVR5cGUsIGNvbnRlbnQpIHsKICBjb25zdCBibG9iID0gbmV3IEJsb2IoW2NvbnRlbnRdLCB7IHR5cGU6IG1pbWVUeXBlIH0pOwogIGNvbnN0IHVybCA9IFVSTC5jcmVhdGVPYmplY3RVUkwoYmxvYik7CiAgY29uc3QgYSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2EnKTsKICBhLmhyZWYgPSB1cmw7CiAgYS5kb3dubG9hZCA9IGZpbGVuYW1lOwogIGRvY3VtZW50LmJvZHkuYXBwZW5kQ2hpbGQoYSk7CiAgYS5jbGljaygpOwogIGEucmVtb3ZlKCk7CiAgVVJMLnJldm9rZU9iamVjdFVSTCh1cmwpOwp9CgovKiogUXVvdGVkLWNvbW1hLWpvaW4gQ1NWIOKAlCB0aGUgY2xpZW50LXNpZGUgbWlycm9yIG9mIGFwcC5qcydzIHRvQ1NWKCksIGZvciB0YWJsZXMgd2hvc2UgZnVsbCBkYXRhc2V0IGlzIGFscmVhZHkgbG9hZGVkIGluIHRoZSBicm93c2VyLiAqLwpmdW5jdGlvbiB0b0NTVkNsaWVudFNpZGUocm93cywgY29sdW1ucykgewogIGNvbnN0IGVzY2FwZSA9ICh2KSA9PiB7CiAgICBjb25zdCBzID0gdiA9PT0gbnVsbCB8fCB2ID09PSB1bmRlZmluZWQgPyAnJyA6IFN0cmluZyh2KTsKICAgIHJldHVybiAvWyIsXHJcbl0vLnRlc3QocykgPyBgIiR7cy5yZXBsYWNlKC8iL2csICciIicpfSJgIDogczsKICB9OwogIGNvbnN0IGxpbmVzID0gW2NvbHVtbnMubWFwKChjKSA9PiBlc2NhcGUoYy5sYWJlbCkpLmpvaW4oJywnKV07CiAgcm93cy5mb3JFYWNoKChyb3cpID0+IGxpbmVzLnB1c2goY29sdW1ucy5tYXAoKGMpID0+IGVzY2FwZShyb3dbYy5rZXldKSkuam9pbignLCcpKSk7CiAgcmV0dXJuIGxpbmVzLmpvaW4oJ1xyXG4nKTsKfQoKLyoqIEEgc21hbGwgIkV4cG9ydCBDU1YgLyBFeHBvcnQgRXhjZWwiIGJ1dHRvbiBwYWlyLiBgZ2V0Um93c0FuZENvbHVtbnMoKWAgaXMgY2FsbGVkIGF0IGNsaWNrIHRpbWUgc28gaXQgYWx3YXlzIGV4cG9ydHMgd2hhdGV2ZXIncyBjdXJyZW50bHkgZmlsdGVyZWQvc29ydGVkLCBuZXZlciBhIHN0YWxlIHNuYXBzaG90LiAqLwpmdW5jdGlvbiBidWlsZEV4cG9ydEJ1dHRvbnMoeyBnZXRSb3dzQW5kQ29sdW1ucywgZmlsZW5hbWVCYXNlLCBzaGVldE5hbWUgfSkgewogIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICB3cmFwLmNsYXNzTmFtZSA9ICdleHBvcnQtYnV0dG9ucyc7CiAgY29uc3QgY3N2QnRuID0gaWNvbkJ0bignYnRuJywgJ2ZpbGUtZG93bicsICdFeHBvcnQgQ1NWJyk7CiAgY3N2QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gewogICAgY29uc3QgeyByb3dzLCBjb2x1bW5zIH0gPSBnZXRSb3dzQW5kQ29sdW1ucygpOwogICAgaWYgKCFyb3dzLmxlbmd0aCkgeyBUb2FzdC5zaG93KCdOb3RoaW5nIHRvIGV4cG9ydC4nLCAnZXJyb3InKTsgcmV0dXJuOyB9CiAgICBkb3dubG9hZEJsb2IoYCR7ZmlsZW5hbWVCYXNlfS5jc3ZgLCAndGV4dC9jc3Y7Y2hhcnNldD11dGYtOCcsIHRvQ1NWQ2xpZW50U2lkZShyb3dzLCBjb2x1bW5zKSk7CiAgfSk7CiAgY29uc3QgeGxzeEJ0biA9IGljb25CdG4oJ2J0bicsICdmaWxlLXNwcmVhZHNoZWV0JywgJ0V4cG9ydCBFeGNlbCcpOwogIHhsc3hCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBhc3luYyAoKSA9PiB7CiAgICBjb25zdCB7IHJvd3MsIGNvbHVtbnMgfSA9IGdldFJvd3NBbmRDb2x1bW5zKCk7CiAgICBpZiAoIXJvd3MubGVuZ3RoKSB7IFRvYXN0LnNob3coJ05vdGhpbmcgdG8gZXhwb3J0LicsICdlcnJvcicpOyByZXR1cm47IH0KICAgIHRyeSB7CiAgICAgIGNvbnN0IHJlcyA9IGF3YWl0IGZldGNoKCcvYXBpL2V4cG9ydCcsIHsKICAgICAgICBtZXRob2Q6ICdQT1NUJywKICAgICAgICBoZWFkZXJzOiB7ICdDb250ZW50LVR5cGUnOiAnYXBwbGljYXRpb24vanNvbicgfSwKICAgICAgICBib2R5OiBKU09OLnN0cmluZ2lmeSh7IHJvd3MsIGNvbHVtbnMsIGZvcm1hdDogJ3hsc3gnLCBmaWxlbmFtZTogZmlsZW5hbWVCYXNlLCBzaGVldE5hbWU6IHNoZWV0TmFtZSB8fCBmaWxlbmFtZUJhc2UgfSksCiAgICAgIH0pOwogICAgICBpZiAoIXJlcy5vaykgdGhyb3cgbmV3IEVycm9yKCdFeHBvcnQgZmFpbGVkLicpOwogICAgICBjb25zdCBibG9iID0gYXdhaXQgcmVzLmJsb2IoKTsKICAgICAgY29uc3QgdXJsID0gVVJMLmNyZWF0ZU9iamVjdFVSTChibG9iKTsKICAgICAgY29uc3QgYSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2EnKTsKICAgICAgYS5ocmVmID0gdXJsOwogICAgICBhLmRvd25sb2FkID0gYCR7ZmlsZW5hbWVCYXNlfS54bHN4YDsKICAgICAgZG9jdW1lbnQuYm9keS5hcHBlbmRDaGlsZChhKTsKICAgICAgYS5jbGljaygpOwogICAgICBhLnJlbW92ZSgpOwogICAgICBVUkwucmV2b2tlT2JqZWN0VVJMKHVybCk7CiAgICB9IGNhdGNoIChlcnIpIHsKICAgICAgVG9hc3Quc2hvdyhlcnIubWVzc2FnZSB8fCAnRXhwb3J0IGZhaWxlZC4nLCAnZXJyb3InKTsKICAgIH0KICB9KTsKICB3cmFwLmFwcGVuZChjc3ZCdG4sIHhsc3hCdG4pOwogIHJldHVybiB3cmFwOwp9CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgQ2hhcnRzIOKAlCBDaGFydC5qcyBidWlsZGVycyAodmFsaWRhdGVkIGNhdGVnb3JpY2FsIHBhbGV0dGUsCiAgIGhhaXJsaW5lIHJlY2Vzc2l2ZSBncmlkbGluZXMsIHNpbmdsZSBheGlzLCBsZWdlbmQgYWx3YXlzCiAgIHByZXNlbnQgZm9yIDIrIHNlcmllcywgaW5kZXgtbW9kZSB0b29sdGlwcykuCiAgID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwppZiAod2luZG93LkNoYXJ0RGF0YUxhYmVscykgQ2hhcnQucmVnaXN0ZXIod2luZG93LkNoYXJ0RGF0YUxhYmVscyk7Cgpjb25zdCBDaGFydHMgPSAoKCkgPT4gewogIGNvbnN0IHJlZ2lzdHJ5ID0gbmV3IE1hcCgpOyAvLyBjYW52YXNJZCAtPiBDaGFydCBpbnN0YW5jZSwgc28gcmUtcmVuZGVycyBkZXN0cm95IHRoZSBvbGQgb25lIGZpcnN0CiAgY29uc3QgTUFYX0xBQkVMRURfSVRFTVMgPSAyMDsgLy8gYmV5b25kIHRoaXMsIHBlci1pdGVtIHZhbHVlIGxhYmVscyB3b3VsZCBvdmVybGFwIOKAlCByZWx5IG9uIHRvb2x0aXBzIGluc3RlYWQKCiAgZnVuY3Rpb24gY3NzVmFyKG5hbWUpIHsKICAgIHJldHVybiBnZXRDb21wdXRlZFN0eWxlKGRvY3VtZW50LmRvY3VtZW50RWxlbWVudCkuZ2V0UHJvcGVydHlWYWx1ZShuYW1lKS50cmltKCk7CiAgfQoKICBjb25zdCBTRVJJRVNfVkFSUyA9IFsnLS1zZXJpZXMtMScsICctLXNlcmllcy0yJywgJy0tc2VyaWVzLTMnLCAnLS1zZXJpZXMtNCcsICctLXNlcmllcy01JywgJy0tc2VyaWVzLTYnLCAnLS1zZXJpZXMtNycsICctLXNlcmllcy04J107CiAgZnVuY3Rpb24gc2VyaWVzQ29sb3IoaW5kZXgpIHsKICAgIHJldHVybiBjc3NWYXIoU0VSSUVTX1ZBUlNbaW5kZXggJSBTRVJJRVNfVkFSUy5sZW5ndGhdKTsKICB9CgogIGZ1bmN0aW9uIGJhc2VHcmlkKCkgewogICAgcmV0dXJuIHsKICAgICAgY29sb3I6IGNzc1ZhcignLS1ncmlkbGluZScpLAogICAgICBkcmF3VGlja3M6IGZhbHNlLAogICAgfTsKICB9CiAgZnVuY3Rpb24gYmFzZVRpY2tzKCkgewogICAgcmV0dXJuIHsgY29sb3I6IGNzc1ZhcignLS10ZXh0LW11dGVkJyksIGZvbnQ6IHsgc2l6ZTogMTEgfSB9OwogIH0KICBmdW5jdGlvbiBiYXNlVG9vbHRpcCgpIHsKICAgIHJldHVybiB7CiAgICAgIGJhY2tncm91bmRDb2xvcjogY3NzVmFyKCctLXN1cmZhY2UtMScpLAogICAgICB0aXRsZUNvbG9yOiBjc3NWYXIoJy0tdGV4dC1wcmltYXJ5JyksCiAgICAgIGJvZHlDb2xvcjogY3NzVmFyKCctLXRleHQtc2Vjb25kYXJ5JyksCiAgICAgIGJvcmRlckNvbG9yOiBjc3NWYXIoJy0tYm9yZGVyJyksCiAgICAgIGJvcmRlcldpZHRoOiAxLAogICAgICBjb3JuZXJSYWRpdXM6IDEwLAogICAgICBwYWRkaW5nOiAxMiwKICAgICAgYm94UGFkZGluZzogNCwKICAgICAgdGl0bGVGb250OiB7IHNpemU6IDEyLCB3ZWlnaHQ6ICc3MDAnIH0sCiAgICAgIGJvZHlGb250OiB7IHNpemU6IDEyIH0sCiAgICB9OwogIH0KICBmdW5jdGlvbiBsYWJlbENvbG9yKCkgewogICAgcmV0dXJuIGNzc1ZhcignLS10ZXh0LXByaW1hcnknKTsKICB9CiAgLyoqIFNuYXBweSwgc3VidGxlIG1vdGlvbiDigJQgaW4gdGhlIDE1MC0zMDBtcyByYW5nZSB0aGUgcmVkZXNpZ24gY2FsbHMgZm9yLCBuZXZlciBib3VuY3kuICovCiAgZnVuY3Rpb24gYmFzZUFuaW1hdGlvbigpIHsKICAgIHJldHVybiB7IGR1cmF0aW9uOiAyODAsIGVhc2luZzogJ2Vhc2VPdXRRdWFydCcgfTsKICB9CgogIGZ1bmN0aW9uIGRlc3Ryb3koY2FudmFzSWQpIHsKICAgIGlmIChyZWdpc3RyeS5oYXMoY2FudmFzSWQpKSB7CiAgICAgIHJlZ2lzdHJ5LmdldChjYW52YXNJZCkuZGVzdHJveSgpOwogICAgICByZWdpc3RyeS5kZWxldGUoY2FudmFzSWQpOwogICAgfQogIH0KCiAgLyoqIE11bHRpLXNlcmllcyBsaW5lIGNoYXJ0IChlLmcuIHdlZWtseSB0cmVuZCBwZXIgcGxhdGZvcm0pLiBPbmUgc2VyaWVzIG5lZWRzIG5vIGxlZ2VuZCBib3guCiAgICAgIFBlci1wb2ludCB2YWx1ZSBsYWJlbHMgYXJlIHNob3duIG9ubHkgZm9yIGEgc2luZ2xlIHNlcmllcyDigJQgd2l0aCBzZXZlcmFsIHNlcmllcyBvdmVybGFpZCwKICAgICAgbGFiZWxpbmcgZXZlcnkgcG9pbnQgd291bGQgb3ZlcmxhcCwgc28gdGhvc2UgcmVseSBvbiB0aGUgKHN0aWxsLXByZXNlbnQpIGhvdmVyIHRvb2x0aXAuICovCiAgZnVuY3Rpb24gdHJlbmRDaGFydChjYW52YXNJZCwgeyBsYWJlbHMsIHNlcmllcywgZm9ybWF0VmFsdWUgfSkgewogICAgZGVzdHJveShjYW52YXNJZCk7CiAgICBjb25zdCBjdHggPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZChjYW52YXNJZCk7CiAgICBpZiAoIWN0eCkgcmV0dXJuIG51bGw7CiAgICBjb25zdCBmbXQgPSBmb3JtYXRWYWx1ZSB8fCAoKHYpID0+IEZvcm1hdC5zbWFydCh2KSk7CiAgICBjb25zdCBzaG93TGFiZWxzID0gc2VyaWVzLmxlbmd0aCA9PT0gMSAmJiBsYWJlbHMubGVuZ3RoIDw9IE1BWF9MQUJFTEVEX0lURU1TOwoKICAgIGNvbnN0IGRhdGFzZXRzID0gc2VyaWVzLm1hcCgocywgaSkgPT4gKHsKICAgICAgbGFiZWw6IHMubGFiZWwsCiAgICAgIGRhdGE6IHMuZGF0YSwKICAgICAgYm9yZGVyQ29sb3I6IHMuY29sb3IgfHwgc2VyaWVzQ29sb3IoaSksCiAgICAgIGJhY2tncm91bmRDb2xvcjogcy5jb2xvciB8fCBzZXJpZXNDb2xvcihpKSwKICAgICAgYm9yZGVyV2lkdGg6IDIsCiAgICAgIHBvaW50UmFkaXVzOiBzaG93TGFiZWxzID8gMyA6IDAsCiAgICAgIHBvaW50SG92ZXJSYWRpdXM6IDQsCiAgICAgIHBvaW50SGl0UmFkaXVzOiAxMiwKICAgICAgdGVuc2lvbjogMC4yNSwKICAgICAgZmlsbDogZmFsc2UsCiAgICB9KSk7CgogICAgY29uc3QgY2hhcnQgPSBuZXcgQ2hhcnQoY3R4LCB7CiAgICAgIHR5cGU6ICdsaW5lJywKICAgICAgZGF0YTogeyBsYWJlbHMsIGRhdGFzZXRzIH0sCiAgICAgIG9wdGlvbnM6IHsKICAgICAgICByZXNwb25zaXZlOiB0cnVlLAogICAgICAgIG1haW50YWluQXNwZWN0UmF0aW86IGZhbHNlLAogICAgICAgIGludGVyYWN0aW9uOiB7IG1vZGU6ICdpbmRleCcsIGludGVyc2VjdDogZmFsc2UgfSwKICAgICAgICBsYXlvdXQ6IHsgcGFkZGluZzogeyB0b3A6IHNob3dMYWJlbHMgPyAyMCA6IDggfSB9LAogICAgICAgIGFuaW1hdGlvbjogYmFzZUFuaW1hdGlvbigpLAogICAgICAgIHBsdWdpbnM6IHsKICAgICAgICAgIGxlZ2VuZDogewogICAgICAgICAgICBkaXNwbGF5OiBzZXJpZXMubGVuZ3RoID4gMSwKICAgICAgICAgICAgcG9zaXRpb246ICdib3R0b20nLAogICAgICAgICAgICBsYWJlbHM6IHsgY29sb3I6IGNzc1ZhcignLS10ZXh0LXNlY29uZGFyeScpLCB1c2VQb2ludFN0eWxlOiB0cnVlLCBwb2ludFN0eWxlOiAnbGluZScsIGJveFdpZHRoOiAxNiwgcGFkZGluZzogMTYsIGZvbnQ6IHsgc2l6ZTogMTEgfSB9LAogICAgICAgICAgfSwKICAgICAgICAgIHRvb2x0aXA6IHsgLi4uYmFzZVRvb2x0aXAoKSwgdXNlUG9pbnRTdHlsZTogdHJ1ZSB9LAogICAgICAgICAgZGF0YWxhYmVsczogc2hvd0xhYmVscwogICAgICAgICAgICA/IHsgYWxpZ246ICd0b3AnLCBhbmNob3I6ICdlbmQnLCBjb2xvcjogbGFiZWxDb2xvcigpLCBmb250OiB7IHNpemU6IDExLCB3ZWlnaHQ6ICc2MDAnIH0sIGZvcm1hdHRlcjogKHYpID0+IGZtdCh2KSB9CiAgICAgICAgICAgIDogeyBkaXNwbGF5OiBmYWxzZSB9LAogICAgICAgIH0sCiAgICAgICAgc2NhbGVzOiB7CiAgICAgICAgICB4OiB7IGdyaWQ6IHsgZGlzcGxheTogZmFsc2UgfSwgdGlja3M6IGJhc2VUaWNrcygpIH0sCiAgICAgICAgICB5OiB7IGdyaWQ6IGJhc2VHcmlkKCksIHRpY2tzOiBiYXNlVGlja3MoKSwgYm9yZGVyOiB7IGRpc3BsYXk6IGZhbHNlIH0sIGJlZ2luQXRaZXJvOiB0cnVlIH0sCiAgICAgICAgfSwKICAgICAgfSwKICAgIH0pOwogICAgcmVnaXN0cnkuc2V0KGNhbnZhc0lkLCBjaGFydCk7CiAgICByZXR1cm4gY2hhcnQ7CiAgfQoKICAvKiogU2luZ2xlLW1ldHJpYyBiYXIgY2hhcnQgYWNyb3NzIHBsYXRmb3JtcyAoaWRlbnRpdHkgZW5jb2Rpbmcg4oCUIGVhY2ggYmFyIElTIGEgcGxhdGZvcm0pLiAqLwogIGZ1bmN0aW9uIHBsYXRmb3JtQmFyQ2hhcnQoY2FudmFzSWQsIHsgbGFiZWxzLCBkYXRhLCBjb2xvcnMsIGZvcm1hdFZhbHVlIH0pIHsKICAgIGRlc3Ryb3koY2FudmFzSWQpOwogICAgY29uc3QgY3R4ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoY2FudmFzSWQpOwogICAgaWYgKCFjdHgpIHJldHVybiBudWxsOwogICAgY29uc3QgZm10ID0gZm9ybWF0VmFsdWUgfHwgKCh2KSA9PiBGb3JtYXQuc21hcnQodikpOwogICAgY29uc3Qgc2hvd0xhYmVscyA9IGxhYmVscy5sZW5ndGggPD0gTUFYX0xBQkVMRURfSVRFTVM7CgogICAgY29uc3QgY2hhcnQgPSBuZXcgQ2hhcnQoY3R4LCB7CiAgICAgIHR5cGU6ICdiYXInLAogICAgICBkYXRhOiB7CiAgICAgICAgbGFiZWxzLAogICAgICAgIGRhdGFzZXRzOiBbCiAgICAgICAgICB7CiAgICAgICAgICAgIGRhdGEsCiAgICAgICAgICAgIGJhY2tncm91bmRDb2xvcjogY29sb3JzLAogICAgICAgICAgICBib3JkZXJSYWRpdXM6IDQsCiAgICAgICAgICAgIG1heEJhclRoaWNrbmVzczogMjgsCiAgICAgICAgICAgIGJvcmRlclNraXBwZWQ6ICdib3R0b20nLAogICAgICAgICAgfSwKICAgICAgICBdLAogICAgICB9LAogICAgICBvcHRpb25zOiB7CiAgICAgICAgcmVzcG9uc2l2ZTogdHJ1ZSwKICAgICAgICBtYWludGFpbkFzcGVjdFJhdGlvOiBmYWxzZSwKICAgICAgICBsYXlvdXQ6IHsgcGFkZGluZzogeyB0b3A6IHNob3dMYWJlbHMgPyAyMCA6IDggfSB9LAogICAgICAgIGFuaW1hdGlvbjogYmFzZUFuaW1hdGlvbigpLAogICAgICAgIHBsdWdpbnM6IHsKICAgICAgICAgIGxlZ2VuZDogeyBkaXNwbGF5OiBmYWxzZSB9LAogICAgICAgICAgdG9vbHRpcDogYmFzZVRvb2x0aXAoKSwKICAgICAgICAgIGRhdGFsYWJlbHM6IHNob3dMYWJlbHMKICAgICAgICAgICAgPyB7IGFsaWduOiAnZW5kJywgYW5jaG9yOiAnZW5kJywgY29sb3I6IGxhYmVsQ29sb3IoKSwgZm9udDogeyBzaXplOiAxMSwgd2VpZ2h0OiAnNjAwJyB9LCBmb3JtYXR0ZXI6ICh2KSA9PiBmbXQodikgfQogICAgICAgICAgICA6IHsgZGlzcGxheTogZmFsc2UgfSwKICAgICAgICB9LAogICAgICAgIHNjYWxlczogewogICAgICAgICAgeDogeyBncmlkOiB7IGRpc3BsYXk6IGZhbHNlIH0sIHRpY2tzOiBiYXNlVGlja3MoKSB9LAogICAgICAgICAgeTogeyBncmlkOiBiYXNlR3JpZCgpLCB0aWNrczogYmFzZVRpY2tzKCksIGJvcmRlcjogeyBkaXNwbGF5OiBmYWxzZSB9LCBiZWdpbkF0WmVybzogdHJ1ZSB9LAogICAgICAgIH0sCiAgICAgIH0sCiAgICB9KTsKICAgIHJlZ2lzdHJ5LnNldChjYW52YXNJZCwgY2hhcnQpOwogICAgcmV0dXJuIGNoYXJ0OwogIH0KCiAgLyoqIFBpZSBjaGFydCAoYSBoYW5kZnVsIG9mIGNhdGVnb3JpZXMgb25seSDigJQgZS5nLiBDYW1wYWlnbiBQZXJmb3JtYW5jZSdzIEFkcy9PcmdhbmljIHNwbGl0KS4KICAgICAgU2xpY2UgbGFiZWxzIHNob3cgYm90aCBzaGFyZS1vZi13aG9sZSBhbmQgdGhlIGFjdHVhbCB2YWx1ZSwgcGVyIHRoZSAibm8gaG92ZXIgcmVxdWlyZWQiIGdvYWwuICovCiAgZnVuY3Rpb24gcGllQ2hhcnQoY2FudmFzSWQsIHsgbGFiZWxzLCBkYXRhLCBjb2xvcnMsIGZvcm1hdFZhbHVlIH0pIHsKICAgIGRlc3Ryb3koY2FudmFzSWQpOwogICAgY29uc3QgY3R4ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoY2FudmFzSWQpOwogICAgaWYgKCFjdHgpIHJldHVybiBudWxsOwogICAgY29uc3QgZm10ID0gZm9ybWF0VmFsdWUgfHwgKCh2KSA9PiBGb3JtYXQuc21hcnQodikpOwogICAgY29uc3QgdG90YWwgPSBkYXRhLnJlZHVjZSgoc3VtLCB2KSA9PiBzdW0gKyAodiB8fCAwKSwgMCk7CgogICAgY29uc3QgY2hhcnQgPSBuZXcgQ2hhcnQoY3R4LCB7CiAgICAgIHR5cGU6ICdwaWUnLAogICAgICBkYXRhOiB7CiAgICAgICAgbGFiZWxzLAogICAgICAgIGRhdGFzZXRzOiBbeyBkYXRhLCBiYWNrZ3JvdW5kQ29sb3I6IGNvbG9ycywgYm9yZGVyQ29sb3I6IGNzc1ZhcignLS1zdXJmYWNlLTEnKSwgYm9yZGVyV2lkdGg6IDIgfV0sCiAgICAgIH0sCiAgICAgIG9wdGlvbnM6IHsKICAgICAgICByZXNwb25zaXZlOiB0cnVlLAogICAgICAgIG1haW50YWluQXNwZWN0UmF0aW86IGZhbHNlLAogICAgICAgIGFuaW1hdGlvbjogYmFzZUFuaW1hdGlvbigpLAogICAgICAgIHBsdWdpbnM6IHsKICAgICAgICAgIGxlZ2VuZDogeyBkaXNwbGF5OiB0cnVlLCBwb3NpdGlvbjogJ2JvdHRvbScsIGxhYmVsczogeyBjb2xvcjogY3NzVmFyKCctLXRleHQtc2Vjb25kYXJ5JyksIGJveFdpZHRoOiAxMiwgcGFkZGluZzogMTYsIGZvbnQ6IHsgc2l6ZTogMTEgfSB9IH0sCiAgICAgICAgICB0b29sdGlwOiBiYXNlVG9vbHRpcCgpLAogICAgICAgICAgZGF0YWxhYmVsczogewogICAgICAgICAgICBjb2xvcjogJyNmZmYnLAogICAgICAgICAgICBmb250OiB7IHNpemU6IDEyLCB3ZWlnaHQ6ICc3MDAnIH0sCiAgICAgICAgICAgIGZvcm1hdHRlcjogKHYpID0+IHsKICAgICAgICAgICAgICBjb25zdCBwY3QgPSB0b3RhbCA/IE1hdGgucm91bmQoKHYgLyB0b3RhbCkgKiAxMDAwKSAvIDEwIDogMDsKICAgICAgICAgICAgICByZXR1cm4gYCR7cGN0fSVcbiR7Zm10KHYpfWA7CiAgICAgICAgICAgIH0sCiAgICAgICAgICB9LAogICAgICAgIH0sCiAgICAgIH0sCiAgICB9KTsKICAgIHJlZ2lzdHJ5LnNldChjYW52YXNJZCwgY2hhcnQpOwogICAgcmV0dXJuIGNoYXJ0OwogIH0KCiAgLyoqIEdyb3VwZWQgdmVydGljYWwgYmFyIGNoYXJ0IOKAlCBhIGZldyBjYXRlZ29yaWVzLCAyKyBuYW1lZCBzZXJpZXMgc2hvd24gc2lkZSBieSBzaWRlCiAgICAgIChlLmcuIFRoaXMgV2VlayB2cyBMYXN0IFdlZWsgYWNyb3NzIG1ldHJpY3MpLiBTYW1lIHZpc3VhbCBsYW5ndWFnZSBhcyBwbGF0Zm9ybUJhckNoYXJ0OwogICAgICBsZWdlbmQgaXMgYWx3YXlzIG9uIHNpbmNlIHRoZSBzZXJpZXMgbmFtZXMgY2FycnkgdGhlIG1lYW5pbmcuICovCiAgZnVuY3Rpb24gZ3JvdXBlZEJhckNoYXJ0KGNhbnZhc0lkLCB7IGxhYmVscywgc2VyaWVzLCBmb3JtYXRWYWx1ZSB9KSB7CiAgICBkZXN0cm95KGNhbnZhc0lkKTsKICAgIGNvbnN0IGN0eCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGNhbnZhc0lkKTsKICAgIGlmICghY3R4KSByZXR1cm4gbnVsbDsKICAgIGNvbnN0IGZtdCA9IGZvcm1hdFZhbHVlIHx8ICgodikgPT4gRm9ybWF0LnNtYXJ0KHYpKTsKICAgIGNvbnN0IHNob3dMYWJlbHMgPSBsYWJlbHMubGVuZ3RoICogc2VyaWVzLmxlbmd0aCA8PSBNQVhfTEFCRUxFRF9JVEVNUzsKCiAgICBjb25zdCBjaGFydCA9IG5ldyBDaGFydChjdHgsIHsKICAgICAgdHlwZTogJ2JhcicsCiAgICAgIGRhdGE6IHsKICAgICAgICBsYWJlbHMsCiAgICAgICAgZGF0YXNldHM6IHNlcmllcy5tYXAoKHMsIGkpID0+ICh7CiAgICAgICAgICBsYWJlbDogcy5sYWJlbCwKICAgICAgICAgIGRhdGE6IHMuZGF0YSwKICAgICAgICAgIGJhY2tncm91bmRDb2xvcjogcy5jb2xvciB8fCBzZXJpZXNDb2xvcihpKSwKICAgICAgICAgIGJvcmRlclJhZGl1czogNCwKICAgICAgICAgIG1heEJhclRoaWNrbmVzczogMzQsCiAgICAgICAgICBib3JkZXJTa2lwcGVkOiAnYm90dG9tJywKICAgICAgICB9KSksCiAgICAgIH0sCiAgICAgIG9wdGlvbnM6IHsKICAgICAgICByZXNwb25zaXZlOiB0cnVlLAogICAgICAgIG1haW50YWluQXNwZWN0UmF0aW86IGZhbHNlLAogICAgICAgIGxheW91dDogeyBwYWRkaW5nOiB7IHRvcDogc2hvd0xhYmVscyA/IDIwIDogOCB9IH0sCiAgICAgICAgYW5pbWF0aW9uOiBiYXNlQW5pbWF0aW9uKCksCiAgICAgICAgcGx1Z2luczogewogICAgICAgICAgbGVnZW5kOiB7CiAgICAgICAgICAgIGRpc3BsYXk6IHRydWUsCiAgICAgICAgICAgIHBvc2l0aW9uOiAnYm90dG9tJywKICAgICAgICAgICAgbGFiZWxzOiB7IGNvbG9yOiBjc3NWYXIoJy0tdGV4dC1zZWNvbmRhcnknKSwgdXNlUG9pbnRTdHlsZTogdHJ1ZSwgcG9pbnRTdHlsZTogJ3JlY3RSb3VuZGVkJywgYm94V2lkdGg6IDEyLCBwYWRkaW5nOiAxNiwgZm9udDogeyBzaXplOiAxMSB9IH0sCiAgICAgICAgICB9LAogICAgICAgICAgdG9vbHRpcDogeyAuLi5iYXNlVG9vbHRpcCgpLCBjYWxsYmFja3M6IHsgbGFiZWw6IChjKSA9PiBgICR7Yy5kYXRhc2V0LmxhYmVsfTogJHtmbXQoYy5wYXJzZWQueSl9YCB9IH0sCiAgICAgICAgICBkYXRhbGFiZWxzOiBzaG93TGFiZWxzCiAgICAgICAgICAgID8geyBhbGlnbjogJ2VuZCcsIGFuY2hvcjogJ2VuZCcsIGNvbG9yOiBsYWJlbENvbG9yKCksIGZvbnQ6IHsgc2l6ZTogMTAsIHdlaWdodDogJzYwMCcgfSwgZm9ybWF0dGVyOiAodikgPT4gZm10KHYpIH0KICAgICAgICAgICAgOiB7IGRpc3BsYXk6IGZhbHNlIH0sCiAgICAgICAgfSwKICAgICAgICBzY2FsZXM6IHsKICAgICAgICAgIHg6IHsgZ3JpZDogeyBkaXNwbGF5OiBmYWxzZSB9LCB0aWNrczogYmFzZVRpY2tzKCkgfSwKICAgICAgICAgIHk6IHsgZ3JpZDogYmFzZUdyaWQoKSwgdGlja3M6IGJhc2VUaWNrcygpLCBib3JkZXI6IHsgZGlzcGxheTogZmFsc2UgfSwgYmVnaW5BdFplcm86IHRydWUgfSwKICAgICAgICB9LAogICAgICB9LAogICAgfSk7CiAgICByZWdpc3RyeS5zZXQoY2FudmFzSWQsIGNoYXJ0KTsKICAgIHJldHVybiBjaGFydDsKICB9CgogIGZ1bmN0aW9uIGRlc3Ryb3lBbGwoKSB7CiAgICBbLi4ucmVnaXN0cnkua2V5cygpXS5mb3JFYWNoKGRlc3Ryb3kpOwogIH0KCiAgcmV0dXJuIHsgdHJlbmRDaGFydCwgcGxhdGZvcm1CYXJDaGFydCwgZ3JvdXBlZEJhckNoYXJ0LCBwaWVDaGFydCwgc2VyaWVzQ29sb3IsIGRlc3Ryb3ksIGRlc3Ryb3lBbGwgfTsKfSkoKTsKCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBEYXNoYm9hcmQgdGFiOiBhIG1ldHJpYy1mb2N1c2VkIHByZW1pdW0gQkkgZGFzaGJvYXJkLiBBIHNpbmdsZQogICBNZXRyaWMgc2VsZWN0b3IgKGR5bmFtaWNhbGx5IHBvcHVsYXRlZCBmcm9tIHdoYXRldmVyIHRoZQogICBzZWxlY3RlZCBwbGF0Zm9ybSdzIGRhdGEgYWN0dWFsbHkgaGFzIOKAlCBuZXZlciBoYXJkY29kZWQpIGRyaXZlcwogICB0aGUgS1BJIGNhcmRzLCB3ZWVrbHkgdHJlbmQsIHBsYXRmb3JtL2NhbXBhaWduL2NvbnRlbnQtdHlwZQogICBicmVha2Rvd25zLCBhbmQgdGhlIFRvcCBQZXJmb3JtaW5nIFBvc3RzIHJhbmtpbmcgdG9nZXRoZXI7CiAgIFBsYXRmb3JtL2RhdGUvY2FtcGFpZ24vY29udGVudC10eXBlIGZpbHRlcmluZyBjb21lcyBmcm9tIHRoZQogICBzaGFyZWQgZmlsdGVyIGJhci4gRXZlcnkgY2hhcnQgc2hvd3MgaXRzIHZhbHVlcyBkaXJlY3RseSAodmlhCiAgIGNoYXJ0anMtcGx1Z2luLWRhdGFsYWJlbHMpIHNvIG5vdGhpbmcgcmVxdWlyZXMgYSBob3ZlciB0byByZWFkLgogICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KY29uc3QgRGFzaGJvYXJkID0gKCgpID0+IHsKICBsZXQgcm9vdDsKICBsZXQgbWV0cmljID0gJ3ZpZXdzJzsKICBsZXQgbWV0cmljT3B0aW9ucyA9IFtdOwoKICBmdW5jdGlvbiBvcHRpb25Gb3Ioa2V5KSB7CiAgICByZXR1cm4gbWV0cmljT3B0aW9ucy5maW5kKChtKSA9PiBtLmtleSA9PT0ga2V5KTsKICB9CiAgZnVuY3Rpb24gbWV0cmljTGFiZWwoa2V5KSB7CiAgICBjb25zdCBvcHQgPSBvcHRpb25Gb3Ioa2V5KTsKICAgIHJldHVybiBvcHQgPyBvcHQubGFiZWwgOiBrZXk7CiAgfQogIGZ1bmN0aW9uIG1ldHJpY1VuaXQoa2V5KSB7CiAgICBjb25zdCBvcHQgPSBvcHRpb25Gb3Ioa2V5KTsKICAgIHJldHVybiBvcHQgPyBvcHQudW5pdCA6ICdudW1iZXInOwogIH0KICBmdW5jdGlvbiBmb3JtYXRNZXRyaWNWYWx1ZShrZXksIHZhbHVlKSB7CiAgICBjb25zdCB1bml0ID0gbWV0cmljVW5pdChrZXkpOwogICAgaWYgKHZhbHVlID09PSBudWxsIHx8IHZhbHVlID09PSB1bmRlZmluZWQpIHJldHVybiAn4oCUJzsKICAgIGlmICh1bml0ID09PSAnZHVyYXRpb24nKSByZXR1cm4gRm9ybWF0LmR1cmF0aW9uKHZhbHVlKTsKICAgIHJldHVybiBGb3JtYXQuc21hcnQodmFsdWUpOwogIH0KCiAgZnVuY3Rpb24gc2hlbGwoKSB7CiAgICByb290LmlubmVySFRNTCA9ICcnOwoKICAgIGNvbnN0IGNvbnRyb2xzID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjb250cm9scy5jbGFzc05hbWUgPSAnZGFzaGJvYXJkLWNvbnRyb2xzJzsKICAgIGNvbnN0IGxhYmVsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnbGFiZWwnKTsKICAgIGxhYmVsLnRleHRDb250ZW50ID0gJ01ldHJpYyc7CiAgICBsYWJlbC5zZXRBdHRyaWJ1dGUoJ2ZvcicsICdkYXNoYm9hcmRNZXRyaWNTZWxlY3QnKTsKICAgIGNvbnN0IG1ldHJpY1NlbGVjdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NlbGVjdCcpOwogICAgbWV0cmljU2VsZWN0LmlkID0gJ2Rhc2hib2FyZE1ldHJpY1NlbGVjdCc7CiAgICBtZXRyaWNPcHRpb25zLmZvckVhY2goKG0pID0+IHsKICAgICAgY29uc3Qgb3B0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7CiAgICAgIG9wdC52YWx1ZSA9IG0ua2V5OwogICAgICBvcHQudGV4dENvbnRlbnQgPSBtLmxhYmVsOwogICAgICBpZiAobS5rZXkgPT09IG1ldHJpYykgb3B0LnNlbGVjdGVkID0gdHJ1ZTsKICAgICAgbWV0cmljU2VsZWN0LmFwcGVuZENoaWxkKG9wdCk7CiAgICB9KTsKICAgIG1ldHJpY1NlbGVjdC5hZGRFdmVudExpc3RlbmVyKCdjaGFuZ2UnLCAoKSA9PiB7CiAgICAgIG1ldHJpYyA9IG1ldHJpY1NlbGVjdC52YWx1ZTsKICAgICAgcmVmcmVzaEZvck1ldHJpYygpOwogICAgfSk7CiAgICBjb250cm9scy5hcHBlbmQobGFiZWwsIG1ldHJpY1NlbGVjdCk7CiAgICByb290LmFwcGVuZENoaWxkKGNvbnRyb2xzKTsKCiAgICBjb25zdCBrcGlUaXRsZSA9IHRleHRFbCgnZGl2JywgJ0tleSBwZXJmb3JtYW5jZSBpbmRpY2F0b3JzJywgJ3NlY3Rpb24tdGl0bGUnKTsKICAgIGNvbnN0IGtwaUdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGtwaUdyaWQuY2xhc3NOYW1lID0gJ3N0YXQtZ3JpZCc7CiAgICBrcGlHcmlkLmlkID0gJ2twaUdyaWQnOwogICAgcm9vdC5hcHBlbmQoa3BpVGl0bGUsIGtwaUdyaWQpOwoKICAgIGNvbnN0IGNoYXJ0c1RpdGxlID0gdGV4dEVsKCdkaXYnLCAnVHJlbmQgJiBwZXJmb3JtYW5jZSBicmVha2Rvd24nLCAnc2VjdGlvbi10aXRsZScpOwogICAgcm9vdC5hcHBlbmQoY2hhcnRzVGl0bGUpOwoKICAgIGNvbnN0IHRyZW5kQ2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdHJlbmRDYXJkLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGNvbnN0IHRyZW5kSGVhZGVyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0cmVuZEhlYWRlci5jbGFzc05hbWUgPSAnY2FyZC1oZWFkZXInOwogICAgdHJlbmRIZWFkZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMycsICdXZWVrbHkgcGVyZm9ybWFuY2UnKSk7CiAgICB0cmVuZEhlYWRlci5maXJzdENoaWxkLmlkID0gJ3RyZW5kQ2FyZFRpdGxlJzsKICAgIHRyZW5kQ2FyZC5hcHBlbmRDaGlsZCh0cmVuZEhlYWRlcik7CiAgICBjb25zdCB0cmVuZENoYXJ0V3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdHJlbmRDaGFydFdyYXAuY2xhc3NOYW1lID0gJ2NoYXJ0LXdyYXAgdGFsbCc7CiAgICB0cmVuZENoYXJ0V3JhcC5pZCA9ICd0cmVuZENoYXJ0V3JhcCc7CiAgICB0cmVuZENoYXJ0V3JhcC5pbm5lckhUTUwgPSAnPGNhbnZhcyBpZD0idHJlbmRDYW52YXMiPjwvY2FudmFzPic7CiAgICB0cmVuZENhcmQuYXBwZW5kQ2hpbGQodHJlbmRDaGFydFdyYXApOwogICAgcm9vdC5hcHBlbmRDaGlsZCh0cmVuZENhcmQpOwoKICAgIGNvbnN0IGdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGdyaWQuY2xhc3NOYW1lID0gJ2NhcmQtZ3JpZCBldmVuJzsKICAgIGdyaWQuc3R5bGUubWFyZ2luVG9wID0gJzE2cHgnOwoKICAgIGNvbnN0IGJyZWFrZG93bkNhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGJyZWFrZG93bkNhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgY29uc3QgYnJlYWtkb3duSGVhZGVyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBicmVha2Rvd25IZWFkZXIuY2xhc3NOYW1lID0gJ2NhcmQtaGVhZGVyJzsKICAgIGJyZWFrZG93bkhlYWRlci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2gzJywgJycpKTsKICAgIGJyZWFrZG93bkhlYWRlci5maXJzdENoaWxkLmlkID0gJ2JyZWFrZG93bkNhcmRUaXRsZSc7CiAgICBicmVha2Rvd25DYXJkLmFwcGVuZENoaWxkKGJyZWFrZG93bkhlYWRlcik7CiAgICBjb25zdCBicmVha2Rvd25XcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBicmVha2Rvd25XcmFwLmNsYXNzTmFtZSA9ICdjaGFydC13cmFwJzsKICAgIGJyZWFrZG93bldyYXAuaWQgPSAnYnJlYWtkb3duQ2hhcnRXcmFwJzsKICAgIGJyZWFrZG93bldyYXAuaW5uZXJIVE1MID0gJzxjYW52YXMgaWQ9ImJyZWFrZG93bkNhbnZhcyI+PC9jYW52YXM+JzsKICAgIGJyZWFrZG93bkNhcmQuYXBwZW5kQ2hpbGQoYnJlYWtkb3duV3JhcCk7CgogICAgY29uc3QgY29udGVudFR5cGVDYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjb250ZW50VHlwZUNhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgY29uc3QgY29udGVudFR5cGVIZWFkZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNvbnRlbnRUeXBlSGVhZGVyLmNsYXNzTmFtZSA9ICdjYXJkLWhlYWRlcic7CiAgICBjb250ZW50VHlwZUhlYWRlci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2gzJywgJycpKTsKICAgIGNvbnRlbnRUeXBlSGVhZGVyLmZpcnN0Q2hpbGQuaWQgPSAnY29udGVudFR5cGVDYXJkVGl0bGUnOwogICAgY29udGVudFR5cGVDYXJkLmFwcGVuZENoaWxkKGNvbnRlbnRUeXBlSGVhZGVyKTsKICAgIGNvbnN0IGNvbnRlbnRUeXBlV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY29udGVudFR5cGVXcmFwLmNsYXNzTmFtZSA9ICdjaGFydC13cmFwJzsKICAgIGNvbnRlbnRUeXBlV3JhcC5pZCA9ICdjb250ZW50VHlwZUNoYXJ0V3JhcCc7CiAgICBjb250ZW50VHlwZVdyYXAuaW5uZXJIVE1MID0gJzxjYW52YXMgaWQ9ImNvbnRlbnRUeXBlQ2FudmFzIj48L2NhbnZhcz4nOwogICAgY29udGVudFR5cGVDYXJkLmFwcGVuZENoaWxkKGNvbnRlbnRUeXBlV3JhcCk7CgogICAgZ3JpZC5hcHBlbmQoYnJlYWtkb3duQ2FyZCwgY29udGVudFR5cGVDYXJkKTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQoZ3JpZCk7CgogICAgY29uc3QgdG9wVGl0bGUgPSB0ZXh0RWwoJ2RpdicsICdUb3AtcGVyZm9ybWluZyBwb3N0cycsICdzZWN0aW9uLXRpdGxlJyk7CiAgICBjb25zdCB0b3BDYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0b3BDYXJkLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGNvbnN0IHRvcEhlYWRlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdG9wSGVhZGVyLmNsYXNzTmFtZSA9ICdjYXJkLWhlYWRlcic7CiAgICB0b3BIZWFkZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMycsICdSYW5rZWQgYnkgc2VsZWN0ZWQgbWV0cmljJykpOwogICAgdG9wQ2FyZC5hcHBlbmRDaGlsZCh0b3BIZWFkZXIpOwogICAgY29uc3QgdGFibGVXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0YWJsZVdyYXAuY2xhc3NOYW1lID0gJ3RhYmxlLXNjcm9sbCc7CiAgICB0YWJsZVdyYXAuaWQgPSAndG9wUG9zdHNUYWJsZSc7CiAgICB0b3BDYXJkLmFwcGVuZENoaWxkKHRhYmxlV3JhcCk7CiAgICByb290LmFwcGVuZCh0b3BUaXRsZSwgdG9wQ2FyZCk7CiAgfQoKICAvKiogQSBzbWFsbCBsYWJlbCt2YWx1ZSBwYWlyIHVzZWQgaW5zaWRlIHRoZSBCZXN0IFBlcmZvcm1pbmcgUG9zdCBjYXJkJ3MgbWV0cmljcyBjb2x1bW4uIGB2YXJpYW50YCAoJ3ByaW1hcnknLydzZWNvbmRhcnknKSBjb250cm9scyBzaXplIGFuZCB3aGV0aGVyIGEgZGl2aWRlciBydWxlIHNpdHMgYWJvdmUgaXQuICovCiAgZnVuY3Rpb24gbWV0cmljQmxvY2sobGFiZWwsIHZhbHVlLCB2YXJpYW50KSB7CiAgICBjb25zdCBibG9jayA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYmxvY2suY2xhc3NOYW1lID0gdmFyaWFudCA9PT0gJ3NlY29uZGFyeScgPyAncG9zdC10aWxlLW1ldHJpYy1ibG9jayBzZWNvbmRhcnknIDogJ3Bvc3QtdGlsZS1tZXRyaWMtYmxvY2snOwogICAgY29uc3QgdmFsdWVFbCA9IHRleHRFbCgnZGl2JywgdmFsdWUsICdwb3N0LXRpbGUtbWV0cmljLXZhbHVlJyk7CiAgICBibG9jay5hcHBlbmQodGV4dEVsKCdkaXYnLCBsYWJlbCwgJ3Bvc3QtdGlsZS1tZXRyaWMtbGFiZWwnKSwgdmFsdWVFbCk7CiAgICByZXF1ZXN0QW5pbWF0aW9uRnJhbWUoKCkgPT4gZml0U3RhdFZhbHVlKHZhbHVlRWwsIDEyKSk7CiAgICByZXR1cm4gYmxvY2s7CiAgfQoKICAvKiogQSBmZWF0dXJlZCBsYW5kc2NhcGUgY2FyZCAoMyBLUEktdGlsZS13aWR0aHMsIHNhbWUgZml4ZWQgaGVpZ2h0IGFzIHRoZSByZXN0IG9mIHRoZQogICAgICByb3cpOiB0aGUgdG9wLXRpZWQgcG9zdCdzIGNhcHRpb24gKHdyYXBzIHVwIHRvIDMgbGluZXMpIHdpdGggYSBwbGF0Zm9ybS1jb2xvciBkb3QKICAgICAgKyBwbGF0Zm9ybSBuYW1lICsgZGF0ZSBvbiB0aGUgbGVmdCAod2hlbiB0aGVyZSdzIG1vcmUgdGhhbiBvbmUgdGllLCB0aGUgZXh0cmEgY291bnQKICAgICAgaXMgZm9sZGVkIGludG8gdGhhdCBzYW1lIG1ldGEgbGluZSByYXRoZXIgdGhhbiBsaXN0aW5nIGV2ZXJ5IHRpZWQgcG9zdCwgc28gdGhlIGNhcmQKICAgICAgbmV2ZXIgaGFzIHRvIGdyb3cgdGFsbGVyIHRoYW4gaXRzIG5laWdoYm9ycyk7IHRoZSBzZWxlY3RlZCBtZXRyaWMgKGxhcmdlKSBhbmQKICAgICAgQ3VycmVudCBGb2xsb3dlcnMgKHNtYWxsZXIsIGJlbG93IGEgZGl2aWRlcikgc3RhY2tlZCBpbiBhIG5hcnJvdyBjb2x1bW4gb24gdGhlIHJpZ2h0LiAqLwogIGZ1bmN0aW9uIGJlc3RQb3N0c1RpbGUobGFiZWwsIHBvc3RzLCBjdXJyZW50Rm9sbG93ZXJzKSB7CiAgICBjb25zdCB0aWxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0aWxlLmNsYXNzTmFtZSA9ICdzdGF0LXRpbGUgcG9zdC10aWxlJzsKCiAgICBjb25zdCBtYWluID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBtYWluLmNsYXNzTmFtZSA9ICdwb3N0LXRpbGUtbWFpbic7CiAgICBtYWluLmFwcGVuZENoaWxkKHN0YXRMYWJlbEVsKGxhYmVsLCAndHJvcGh5JywgJ2dvbGQnKSk7CgogICAgY29uc3QgaGFzUG9zdHMgPSBCb29sZWFuKHBvc3RzICYmIHBvc3RzLmxlbmd0aCk7CiAgICBpZiAoaGFzUG9zdHMpIHsKICAgICAgY29uc3QgcGxhdGZvcm1PcHRpb25zID0gKHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZSB8fCB7IHBsYXRmb3JtczogW10gfSkucGxhdGZvcm1zOwogICAgICBjb25zdCBwcmltYXJ5ID0gcG9zdHNbMF07CiAgICAgIGNvbnN0IHBsYXRNZXRhID0gcGxhdGZvcm1PcHRpb25zLmZpbmQoKHApID0+IHAuaWQgPT09IHByaW1hcnkucGxhdGZvcm0pIHx8IHsgbGFiZWw6IHByaW1hcnkucGxhdGZvcm0sIGNvbG9yOiAndmFyKC0tc2VyaWVzLTEpJyB9OwogICAgICBjb25zdCBjYXB0aW9uID0gcHJpbWFyeS5jYXB0aW9uIHx8ICcobm8gY2FwdGlvbiknOwogICAgICBjb25zdCBjYXB0aW9uRWwgPSB0ZXh0RWwoJ2RpdicsIGNhcHRpb24sICdwb3N0LXRpbGUtY2FwdGlvbicpOwogICAgICBjYXB0aW9uRWwudGl0bGUgPSBjYXB0aW9uOwogICAgICBtYWluLmFwcGVuZENoaWxkKGNhcHRpb25FbCk7CiAgICAgIGNvbnN0IHRpZWROb3RlID0gcG9zdHMubGVuZ3RoID4gMSA/IGAgwrcgKyR7cG9zdHMubGVuZ3RoIC0gMX0gbW9yZSB0aWVkYCA6ICcnOwogICAgICBjb25zdCBtZXRhTGluZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBtZXRhTGluZS5jbGFzc05hbWUgPSAncG9zdC10aWxlLW1ldGEnOwogICAgICBjb25zdCBkb3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICAgIGRvdC5jbGFzc05hbWUgPSAncGxhdGZvcm0tZG90JzsKICAgICAgZG90LnN0eWxlLmJhY2tncm91bmQgPSBwbGF0TWV0YS5jb2xvcjsKICAgICAgbWV0YUxpbmUuYXBwZW5kKGRvdCwgZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUoYCR7cGxhdE1ldGEubGFiZWx9IMK3ICR7Rm9ybWF0LmRhdGUocHJpbWFyeS5wdWJsaXNoX2RhdGUpfSR7dGllZE5vdGV9YCkpOwogICAgICBtYWluLmFwcGVuZENoaWxkKG1ldGFMaW5lKTsKICAgIH0gZWxzZSB7CiAgICAgIG1haW4uYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCAnTm8gZGF0YSB5ZXQnLCAncG9zdC10aWxlLWNhcHRpb24gbXV0ZWQnKSk7CiAgICB9CiAgICB0aWxlLmFwcGVuZENoaWxkKG1haW4pOwoKICAgIGNvbnN0IGhhc0ZvbGxvd2VycyA9IGN1cnJlbnRGb2xsb3dlcnMgIT09IG51bGwgJiYgY3VycmVudEZvbGxvd2VycyAhPT0gdW5kZWZpbmVkOwogICAgaWYgKGhhc1Bvc3RzIHx8IGhhc0ZvbGxvd2VycykgewogICAgICB0aWxlLmFwcGVuZENoaWxkKE9iamVjdC5hc3NpZ24oZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2JyksIHsgY2xhc3NOYW1lOiAncG9zdC10aWxlLWRpdmlkZXInIH0pKTsKICAgICAgY29uc3QgbWV0cmljcyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBtZXRyaWNzLmNsYXNzTmFtZSA9ICdwb3N0LXRpbGUtbWV0cmljcyc7CiAgICAgIGlmIChoYXNQb3N0cykgbWV0cmljcy5hcHBlbmRDaGlsZChtZXRyaWNCbG9jayhtZXRyaWNMYWJlbChtZXRyaWMpLCBmb3JtYXRNZXRyaWNWYWx1ZShtZXRyaWMsIHBvc3RzWzBdLnZhbHVlKSwgJ3ByaW1hcnknKSk7CiAgICAgIGlmIChoYXNGb2xsb3dlcnMpIG1ldHJpY3MuYXBwZW5kQ2hpbGQobWV0cmljQmxvY2soJ0N1cnJlbnQgRm9sbG93ZXJzJywgRm9ybWF0Lm51bWJlcihjdXJyZW50Rm9sbG93ZXJzKSwgJ3NlY29uZGFyeScpKTsKICAgICAgdGlsZS5hcHBlbmRDaGlsZChtZXRyaWNzKTsKICAgIH0KICAgIHJldHVybiB0aWxlOwogIH0KCiAgLyoqIEljb24tYmFkZ2UgKyB0ZXh0IGxhYmVsIHJvdywgc2hhcmVkIGJ5IGV2ZXJ5IEtQSSB0aWxlIGJlbG93IChtYXRjaGVzIHRoZSByZWZlcmVuY2UgZGFzaGJvYXJkJ3MgY29sb3JlZCBwZXItY2FyZCBpY29ucykuIGB2YXJpYW50YCBwaWNrcyB0aGUgYmFkZ2UgY29sb3I6IHYxLXY2IG1hcCB0byB0aGUgcGxhdGZvcm0gc2VyaWVzIHBhbGV0dGUsICdnb2xkJyBpcyByZXNlcnZlZCBmb3IgdGhlIGJyYW5kLWFjY2VudCB0aWxlLiAqLwogIGZ1bmN0aW9uIHN0YXRMYWJlbEVsKHRleHQsIGljb24sIHZhcmlhbnQpIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHdyYXAuY2xhc3NOYW1lID0gJ3N0YXQtbGFiZWwnOwogICAgY29uc3QgYmFkZ2UgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICBiYWRnZS5jbGFzc05hbWUgPSBgc3RhdC1pY29uICR7dmFyaWFudH1gOwogICAgYmFkZ2UuaW5uZXJIVE1MID0gYDxpIGRhdGEtbHVjaWRlPSIke2ljb259IiBzdHlsZT0id2lkdGg6MTZweDtoZWlnaHQ6MTZweDsiPjwvaT5gOwogICAgd3JhcC5hcHBlbmQoYmFkZ2UsIGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKHRleHQpKTsKICAgIHJldHVybiB3cmFwOwogIH0KCiAgZnVuY3Rpb24gc3RhdFRpbGUobGFiZWwsIHZhbHVlLCBmb3JtYXRGbiwgaWNvbiwgdmFyaWFudCkgewogICAgY29uc3QgdGlsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGlsZS5jbGFzc05hbWUgPSAnc3RhdC10aWxlJzsKICAgIGNvbnN0IHZhbHVlRWwgPSB0ZXh0RWwoJ2RpdicsICcnLCAnc3RhdC12YWx1ZScpOwogICAgdGlsZS5hcHBlbmQoc3RhdExhYmVsRWwobGFiZWwsIGljb24sIHZhcmlhbnQpLCB2YWx1ZUVsKTsKICAgIGNvbnN0IGZtdCA9IGZvcm1hdEZuIHx8ICgodikgPT4gRm9ybWF0Lm51bWJlcih2KSk7CiAgICBpZiAodHlwZW9mIHZhbHVlID09PSAnbnVtYmVyJyAmJiBOdW1iZXIuaXNGaW5pdGUodmFsdWUpKSB7CiAgICAgIGFuaW1hdGVDb3VudCh2YWx1ZUVsLCAwLCB2YWx1ZSwgOTAwLCBmbXQpOwogICAgfSBlbHNlIHsKICAgICAgdmFsdWVFbC50ZXh0Q29udGVudCA9IGZtdCh2YWx1ZSk7CiAgICAgIHJlcXVlc3RBbmltYXRpb25GcmFtZSgoKSA9PiBmaXRTdGF0VmFsdWUodmFsdWVFbCkpOwogICAgfQogICAgcmV0dXJuIHRpbGU7CiAgfQoKICAvKiogIkZvbGxvd2VycyBHcm93dGgiIHRpbGU6IGFuIGFic29sdXRlLWRpZmZlcmVuY2Ugc3RhdC12YWx1ZSBwbHVzIGEgcGVyY2VudGFnZSBkZWx0YSBsaW5lIChhcnJvdyArIGNvbG9yIGRyaXZlbiBieSBGb3JtYXQuZGVsdGFDbGFzcywgc2FtZSBjb252ZW50aW9uIGFzIHRoZSBDb21wYXJpc29ucyBwYWdlJ3Mgc3RhdCB0aWxlcykuICovCiAgZnVuY3Rpb24gZm9sbG93ZXJzR3Jvd3RoVGlsZShjaGFuZ2UsIGNoYW5nZVBjdCkgewogICAgY29uc3QgdGlsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGlsZS5jbGFzc05hbWUgPSAnc3RhdC10aWxlJzsKICAgIGNvbnN0IHZhbHVlRWwgPSB0ZXh0RWwoJ2RpdicsICcnLCAnc3RhdC12YWx1ZScpOwogICAgdGlsZS5hcHBlbmQoc3RhdExhYmVsRWwoJ0ZvbGxvd2VycyBHcm93dGgnLCAndHJlbmRpbmctdXAnLCAndjMnKSwgdmFsdWVFbCk7CiAgICBpZiAoY2hhbmdlID09PSBudWxsIHx8IGNoYW5nZSA9PT0gdW5kZWZpbmVkKSB7CiAgICAgIHZhbHVlRWwudGV4dENvbnRlbnQgPSAn4oCUJzsKICAgICAgcmVxdWVzdEFuaW1hdGlvbkZyYW1lKCgpID0+IGZpdFN0YXRWYWx1ZSh2YWx1ZUVsKSk7CiAgICB9IGVsc2UgewogICAgICBhbmltYXRlQ291bnQodmFsdWVFbCwgMCwgY2hhbmdlLCA5MDAsICh2KSA9PiBgJHt2ID4gMCA/ICcrJyA6ICcnfSR7Rm9ybWF0Lm51bWJlcihNYXRoLnJvdW5kKHYpKX1gKTsKICAgIH0KICAgIGNvbnN0IGRlbHRhVGV4dCA9IGNoYW5nZVBjdCA9PT0gbnVsbCB8fCBjaGFuZ2VQY3QgPT09IHVuZGVmaW5lZCA/ICfigJQnIDogRm9ybWF0LnBjdChjaGFuZ2VQY3QpOwogICAgdGlsZS5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIGRlbHRhVGV4dCwgYHN0YXQtZGVsdGEgJHtGb3JtYXQuZGVsdGFDbGFzcyhjaGFuZ2VQY3QpfWApKTsKICAgIHJldHVybiB0aWxlOwogIH0KCiAgLyoqICJOZXcgRm9sbG93ZXJzIiB0aWxlOiBmb2xsb3dlcnMgZ2FpbmVkIHdpdGhpbiB0aGUgY3VycmVudGx5IHNlbGVjdGVkIGRhdGUgcmFuZ2Ug4oCUIHNob3dzICJObyBmb2xsb3dlciB1cGRhdGUiIHJhdGhlciB0aGFuIDAgd2hlbiBub3RoaW5nIGlzIGNvbXB1dGFibGUgZm9yIHRoZSByYW5nZSAocGVyIHNwZWMpLCB3aGljaCBpcyBkaWZmZXJlbnQgZnJvbSBhIGdlbnVpbmUgemVyby4gKi8KICBmdW5jdGlvbiBuZXdGb2xsb3dlcnNUaWxlKG5ld0ZvbGxvd2VycykgewogICAgY29uc3QgdGlsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGlsZS5jbGFzc05hbWUgPSAnc3RhdC10aWxlJzsKICAgIGNvbnN0IHZhbHVlRWwgPSB0ZXh0RWwoJ2RpdicsICcnLCAnc3RhdC12YWx1ZScpOwogICAgdGlsZS5hcHBlbmQoc3RhdExhYmVsRWwoJ05ldyBGb2xsb3dlcnMnLCAndXNlci1wbHVzJywgJ3YxJyksIHZhbHVlRWwpOwogICAgaWYgKG5ld0ZvbGxvd2VycyA9PT0gbnVsbCB8fCBuZXdGb2xsb3dlcnMgPT09IHVuZGVmaW5lZCkgewogICAgICB2YWx1ZUVsLnRleHRDb250ZW50ID0gJ05vIGZvbGxvd2VyIHVwZGF0ZSc7CiAgICAgIHZhbHVlRWwuY2xhc3NMaXN0LmFkZCgnc3RhdC12YWx1ZS1tdXRlZCcpOwogICAgICByZXF1ZXN0QW5pbWF0aW9uRnJhbWUoKCkgPT4gZml0U3RhdFZhbHVlKHZhbHVlRWwpKTsKICAgIH0gZWxzZSB7CiAgICAgIGFuaW1hdGVDb3VudCh2YWx1ZUVsLCAwLCBuZXdGb2xsb3dlcnMsIDkwMCwgKHYpID0+IGAke3YgPiAwID8gJysnIDogJyd9JHtGb3JtYXQubnVtYmVyKE1hdGgucm91bmQodikpfWApOwogICAgfQogICAgcmV0dXJuIHRpbGU7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJLcGlzKHN1bW1hcnksIGZvbGxvd2VycykgewogICAgY29uc3QgZ3JpZCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdrcGlHcmlkJyk7CiAgICBpZiAoIWdyaWQpIHJldHVybjsKICAgIGdyaWQuaW5uZXJIVE1MID0gJyc7CgogICAgZ3JpZC5hcHBlbmRDaGlsZChzdGF0VGlsZSgnSGlnaGVzdCBWYWx1ZScsIHN1bW1hcnkuaGlnaGVzdCwgKHYpID0+IGZvcm1hdE1ldHJpY1ZhbHVlKG1ldHJpYywgdiksICd0cmVuZGluZy11cCcsICd2MScpKTsKICAgIGdyaWQuYXBwZW5kQ2hpbGQoc3RhdFRpbGUoJ0F2ZXJhZ2UgVmFsdWUnLCBzdW1tYXJ5LmF2ZXJhZ2UsICh2KSA9PiBmb3JtYXRNZXRyaWNWYWx1ZShtZXRyaWMsIHYpLCAnYmFyLWNoYXJ0LTInLCAndjQnKSk7CiAgICBncmlkLmFwcGVuZENoaWxkKHN0YXRUaWxlKCdUb3RhbCBWYWx1ZScsIHN1bW1hcnkudG90YWwsICh2KSA9PiBmb3JtYXRNZXRyaWNWYWx1ZShtZXRyaWMsIHYpLCAnbGF5ZXJzJywgJ3Y1JykpOwogICAgZ3JpZC5hcHBlbmRDaGlsZChzdGF0VGlsZSgnTnVtYmVyIG9mIFBvc3RzJywgc3VtbWFyeS5wb3N0Q291bnQsICh2KSA9PiBGb3JtYXQubnVtYmVyKHYpLCAnZmlsZS10ZXh0JywgJ3Y2JykpOwogICAgZ3JpZC5hcHBlbmRDaGlsZChzdGF0VGlsZSgnQ3VycmVudCBGb2xsb3dlcnMnLCBmb2xsb3dlcnMuY3VycmVudEZvbGxvd2VycywgKHYpID0+ICh2ID09PSBudWxsIHx8IHYgPT09IHVuZGVmaW5lZCA/ICfigJQnIDogRm9ybWF0Lm51bWJlcih2KSksICd1c2VycycsICd2MicpKTsKICAgIGdyaWQuYXBwZW5kQ2hpbGQoZm9sbG93ZXJzR3Jvd3RoVGlsZShmb2xsb3dlcnMuZm9sbG93ZXJzQ2hhbmdlLCBmb2xsb3dlcnMuZm9sbG93ZXJzQ2hhbmdlUGN0KSk7CiAgICBncmlkLmFwcGVuZENoaWxkKG5ld0ZvbGxvd2Vyc1RpbGUoZm9sbG93ZXJzLm5ld0ZvbGxvd2VycykpOwogICAgZ3JpZC5hcHBlbmRDaGlsZChiZXN0UG9zdHNUaWxlKCdCZXN0IFBlcmZvcm1pbmcgUG9zdCcsIHN1bW1hcnkuYmVzdFBvc3RzLCBmb2xsb3dlcnMuY3VycmVudEZvbGxvd2VycykpOwogIH0KCgogIC8qKiBTd2FwcyBhIGNoYXJ0IGNhcmQncyBjYW52YXMgZm9yIGFuIGVtcHR5LXN0YXRlIG1lc3NhZ2UsIG9yIHJlc3RvcmVzIHRoZSBjYW52YXMg4oCUIHNpbmNlCiAgICAgIHJlLXJlbmRlcmluZyBhIENoYXJ0LmpzIGluc3RhbmNlIG5lZWRzIGEgbGl2ZSA8Y2FudmFzPiwgbm90IHdoYXRldmVyIHRoZSBsYXN0IHJlbmRlciBsZWZ0IHRoZXJlLiAqLwogIGZ1bmN0aW9uIGNoYXJ0T3JFbXB0eSh3cmFwSWQsIGNhbnZhc0lkLCBoYXNEYXRhLCBlbXB0eU1lc3NhZ2UsIHJlbmRlckZuKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQod3JhcElkKTsKICAgIGlmICghd3JhcCkgcmV0dXJuOwogICAgQ2hhcnRzLmRlc3Ryb3koY2FudmFzSWQpOwogICAgaWYgKCFoYXNEYXRhKSB7CiAgICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQoZW1wdHlTdGF0ZSh7IGljb246ICdiYXItY2hhcnQtMycsIG1lc3NhZ2U6IGVtcHR5TWVzc2FnZSB9KSk7CiAgICAgIHJldHVybjsKICAgIH0KICAgIHdyYXAuaW5uZXJIVE1MID0gYDxjYW52YXMgaWQ9IiR7Y2FudmFzSWR9Ij48L2NhbnZhcz5gOwogICAgcmVuZGVyRm4oKTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIHJlbmRlclRyZW5kKGZpbHRlcnMpIHsKICAgIGNvbnN0IHBsYXRmb3JtT3B0aW9ucyA9ICh3aW5kb3cuX19maWx0ZXJPcHRpb25zQ2FjaGUgfHwgeyBwbGF0Zm9ybXM6IFtdIH0pLnBsYXRmb3JtczsKICAgIGNvbnN0IG1MYWJlbCA9IG1ldHJpY0xhYmVsKG1ldHJpYyk7CiAgICBjb25zdCBwbGF0Zm9ybXNUb0ZldGNoID0gZmlsdGVycy5wbGF0Zm9ybSA9PT0gJ2FsbCcgPyBwbGF0Zm9ybU9wdGlvbnMubWFwKChwKSA9PiBwLmlkKSA6IFtmaWx0ZXJzLnBsYXRmb3JtXTsKICAgIGNvbnN0IHRyZW5kUmVzcG9uc2VzID0gYXdhaXQgUHJvbWlzZS5hbGwoCiAgICAgIHBsYXRmb3Jtc1RvRmV0Y2gubWFwKChwKSA9PgogICAgICAgIEFwaS50cmVuZCh7IGRhdGVGcm9tOiBmaWx0ZXJzLmRhdGVGcm9tLCBkYXRlVG86IGZpbHRlcnMuZGF0ZVRvLCBwbGF0Zm9ybTogcCwgY2FtcGFpZ25UeXBlOiBmaWx0ZXJzLmNhbXBhaWduVHlwZSwgY29udGVudFR5cGU6IGZpbHRlcnMuY29udGVudFR5cGUgfSkKICAgICAgKQogICAgKTsKICAgIGNvbnN0IHdlZWtTZXQgPSBuZXcgU2V0KCk7CiAgICB0cmVuZFJlc3BvbnNlcy5mb3JFYWNoKChyb3dzKSA9PiByb3dzLmZvckVhY2goKHIpID0+IHdlZWtTZXQuYWRkKHIucGVyaW9kKSkpOwogICAgY29uc3Qgd2Vla3MgPSBbLi4ud2Vla1NldF0uc29ydCgpOwogICAgY29uc3Qgc2VyaWVzID0gcGxhdGZvcm1zVG9GZXRjaC5tYXAoKHAsIGkpID0+IHsKICAgICAgY29uc3QgbWV0YSA9IHBsYXRmb3JtT3B0aW9ucy5maW5kKChwbCkgPT4gcGwuaWQgPT09IHApIHx8IHsgbGFiZWw6IHAgfTsKICAgICAgY29uc3QgYnlXZWVrID0gT2JqZWN0LmZyb21FbnRyaWVzKHRyZW5kUmVzcG9uc2VzW2ldLm1hcCgocikgPT4gW3IucGVyaW9kLCByW21ldHJpY11dKSk7CiAgICAgIHJldHVybiB7IGxhYmVsOiBtZXRhLmxhYmVsLCBjb2xvcjogbWV0YS5jb2xvciwgZGF0YTogd2Vla3MubWFwKCh3KSA9PiAoYnlXZWVrW3ddID09PSB1bmRlZmluZWQgPyBudWxsIDogYnlXZWVrW3ddKSkgfTsKICAgIH0pOwoKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd0cmVuZENhcmRUaXRsZScpLnRleHRDb250ZW50ID0KICAgICAgZmlsdGVycy5wbGF0Zm9ybSA9PT0gJ2FsbCcgPyBgV2Vla2x5ICR7bUxhYmVsfSBieSBQbGF0Zm9ybWAgOiBgJHttTGFiZWx9IFRyZW5kYDsKCiAgICBjaGFydE9yRW1wdHkoJ3RyZW5kQ2hhcnRXcmFwJywgJ3RyZW5kQ2FudmFzJywgd2Vla3MubGVuZ3RoID4gMCwgJ05vIGRhdGEgaW4gdGhpcyByYW5nZSB5ZXQuJywgKCkgPT4gewogICAgICBDaGFydHMudHJlbmRDaGFydCgndHJlbmRDYW52YXMnLCB7IGxhYmVsczogd2Vla3MubWFwKEZvcm1hdC5kYXRlKSwgc2VyaWVzLCBmb3JtYXRWYWx1ZTogKHYpID0+IGZvcm1hdE1ldHJpY1ZhbHVlKG1ldHJpYywgdikgfSk7CiAgICB9KTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIHJlbmRlckJyZWFrZG93bihmaWx0ZXJzKSB7CiAgICBjb25zdCBtTGFiZWwgPSBtZXRyaWNMYWJlbChtZXRyaWMpOwogICAgY29uc3QgdGl0bGVFbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdicmVha2Rvd25DYXJkVGl0bGUnKTsKCiAgICBpZiAoZmlsdGVycy5wbGF0Zm9ybSA9PT0gJ2FsbCcpIHsKICAgICAgdGl0bGVFbC50ZXh0Q29udGVudCA9IGBQbGF0Zm9ybSBDb21wYXJpc29uIOKAlCAke21MYWJlbH1gOwogICAgICBjb25zdCBicmVha2Rvd24gPSBhd2FpdCBBcGkucGxhdGZvcm1CcmVha2Rvd24oZmlsdGVycyk7CiAgICAgIGNvbnN0IHNvcnRlZCA9IGJyZWFrZG93bi5maWx0ZXIoKHApID0+IHBbbWV0cmljXSAhPT0gbnVsbCAmJiBwW21ldHJpY10gIT09IHVuZGVmaW5lZCkuc29ydCgoYSwgYikgPT4gYlttZXRyaWNdIC0gYVttZXRyaWNdKTsKICAgICAgY2hhcnRPckVtcHR5KCdicmVha2Rvd25DaGFydFdyYXAnLCAnYnJlYWtkb3duQ2FudmFzJywgc29ydGVkLmxlbmd0aCA+IDAsICdObyBkYXRhIGluIHRoaXMgcmFuZ2UgeWV0LicsICgpID0+IHsKICAgICAgICBDaGFydHMucGxhdGZvcm1CYXJDaGFydCgnYnJlYWtkb3duQ2FudmFzJywgewogICAgICAgICAgbGFiZWxzOiBzb3J0ZWQubWFwKChwKSA9PiBwLmxhYmVsKSwKICAgICAgICAgIGRhdGE6IHNvcnRlZC5tYXAoKHApID0+IHBbbWV0cmljXSksCiAgICAgICAgICBjb2xvcnM6IHNvcnRlZC5tYXAoKHApID0+IHAuY29sb3IpLAogICAgICAgICAgZm9ybWF0VmFsdWU6ICh2KSA9PiBmb3JtYXRNZXRyaWNWYWx1ZShtZXRyaWMsIHYpLAogICAgICAgIH0pOwogICAgICB9KTsKICAgIH0gZWxzZSB7CiAgICAgIHRpdGxlRWwudGV4dENvbnRlbnQgPSBgQ2FtcGFpZ24gUGVyZm9ybWFuY2Ug4oCUICR7bUxhYmVsfWA7CiAgICAgIGNvbnN0IGNhbXBhaWducyA9IGF3YWl0IEFwaS5jYW1wYWlnbkJyZWFrZG93bihmaWx0ZXJzKTsKICAgICAgY29uc3Qgd2l0aFZhbHVlID0gY2FtcGFpZ25zLmZpbHRlcigoYykgPT4gY1ttZXRyaWNdICE9PSBudWxsICYmIGNbbWV0cmljXSAhPT0gdW5kZWZpbmVkICYmIGNbbWV0cmljXSA+IDApOwogICAgICBjaGFydE9yRW1wdHkoJ2JyZWFrZG93bkNoYXJ0V3JhcCcsICdicmVha2Rvd25DYW52YXMnLCB3aXRoVmFsdWUubGVuZ3RoID4gMCwgJ05vIGNhbXBhaWduIGRhdGEgaW4gdGhpcyByYW5nZSB5ZXQuJywgKCkgPT4gewogICAgICAgIENoYXJ0cy5waWVDaGFydCgnYnJlYWtkb3duQ2FudmFzJywgewogICAgICAgICAgbGFiZWxzOiB3aXRoVmFsdWUubWFwKChjKSA9PiBjLmNhbXBhaWduX3R5cGUpLAogICAgICAgICAgZGF0YTogd2l0aFZhbHVlLm1hcCgoYykgPT4gY1ttZXRyaWNdKSwKICAgICAgICAgIGNvbG9yczogd2l0aFZhbHVlLm1hcCgoXywgaSkgPT4gQ2hhcnRzLnNlcmllc0NvbG9yKGkpKSwKICAgICAgICAgIGZvcm1hdFZhbHVlOiAodikgPT4gZm9ybWF0TWV0cmljVmFsdWUobWV0cmljLCB2KSwKICAgICAgICB9KTsKICAgICAgfSk7CiAgICB9CiAgfQoKICBhc3luYyBmdW5jdGlvbiByZW5kZXJDb250ZW50VHlwZUJyZWFrZG93bihmaWx0ZXJzKSB7CiAgICBjb25zdCBtTGFiZWwgPSBtZXRyaWNMYWJlbChtZXRyaWMpOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2NvbnRlbnRUeXBlQ2FyZFRpdGxlJykudGV4dENvbnRlbnQgPSBgQ29udGVudCBUeXBlIFBlcmZvcm1hbmNlIOKAlCAke21MYWJlbH1gOwogICAgY29uc3Qgcm93cyA9IGF3YWl0IEFwaS5jb250ZW50VHlwZUJyZWFrZG93bihmaWx0ZXJzKTsKICAgIGNvbnN0IHNvcnRlZCA9IHJvd3MuZmlsdGVyKChjKSA9PiBjW21ldHJpY10gIT09IG51bGwgJiYgY1ttZXRyaWNdICE9PSB1bmRlZmluZWQpLnNvcnQoKGEsIGIpID0+IGJbbWV0cmljXSAtIGFbbWV0cmljXSk7CiAgICBjaGFydE9yRW1wdHkoJ2NvbnRlbnRUeXBlQ2hhcnRXcmFwJywgJ2NvbnRlbnRUeXBlQ2FudmFzJywgc29ydGVkLmxlbmd0aCA+IDAsICdObyBkYXRhIGluIHRoaXMgcmFuZ2UgeWV0LicsICgpID0+IHsKICAgICAgQ2hhcnRzLnBsYXRmb3JtQmFyQ2hhcnQoJ2NvbnRlbnRUeXBlQ2FudmFzJywgewogICAgICAgIGxhYmVsczogc29ydGVkLm1hcCgoYykgPT4gYy5jb250ZW50X3R5cGUpLAogICAgICAgIGRhdGE6IHNvcnRlZC5tYXAoKGMpID0+IGNbbWV0cmljXSksCiAgICAgICAgY29sb3JzOiBzb3J0ZWQubWFwKChfLCBpKSA9PiBDaGFydHMuc2VyaWVzQ29sb3IoaSkpLAogICAgICAgIGZvcm1hdFZhbHVlOiAodikgPT4gZm9ybWF0TWV0cmljVmFsdWUobWV0cmljLCB2KSwKICAgICAgfSk7CiAgICB9KTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIHJlbmRlclRvcFBvc3RzKGZpbHRlcnMpIHsKICAgIGNvbnN0IHBvc3RzID0gYXdhaXQgQXBpLnRvcFBvc3RzKHsgLi4uZmlsdGVycywgc29ydEJ5OiBtZXRyaWMsIGxpbWl0OiAxMCB9KTsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndG9wUG9zdHNUYWJsZScpOwogICAgaWYgKCF3cmFwKSByZXR1cm47CiAgICBpZiAoIXBvc3RzLmxlbmd0aCkgewogICAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgICB3cmFwLmFwcGVuZENoaWxkKGVtcHR5U3RhdGUoewogICAgICAgIGljb246ICd0cm9waHknLAogICAgICAgIHRpdGxlOiAnTm8gcG9zdHMgaW4gdGhpcyByYW5nZSB5ZXQnLAogICAgICAgIG1lc3NhZ2U6ICdVcGxvYWQgYSB3ZWVrbHkgZXhwb3J0LCBvciB3aWRlbiB0aGUgZGF0ZSByYW5nZSwgdG8gc2VlIHRvcCBwZXJmb3JtZXJzIGhlcmUuJywKICAgICAgICBhY3Rpb25MYWJlbDogJ1VwbG9hZCBkYXRhJywKICAgICAgICBvbkFjdGlvbjogKCkgPT4gZG9jdW1lbnQucXVlcnlTZWxlY3RvcignLnRhYi1idG5bZGF0YS10YWI9InVwbG9hZCJdJyk/LmNsaWNrKCksCiAgICAgIH0pKTsKICAgICAgcmV0dXJuOwogICAgfQogICAgY29uc3QgcGxhdGZvcm1PcHRpb25zID0gKHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZSB8fCB7IHBsYXRmb3JtczogW10gfSkucGxhdGZvcm1zOwoKICAgIGNvbnN0IHRhYmxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGFibGUnKTsKICAgIHRhYmxlLmNsYXNzTmFtZSA9ICdkYXRhLXRhYmxlJzsKICAgIGNvbnN0IHRoZWFkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGhlYWQnKTsKICAgIGNvbnN0IGhlYWRUciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICBoZWFkVHIuYXBwZW5kKAogICAgICB0ZXh0RWwoJ3RoJywgJ1JhbmsnKSwKICAgICAgdGV4dEVsKCd0aCcsICdEYXRlJyksCiAgICAgIHRleHRFbCgndGgnLCAnUGxhdGZvcm0nKSwKICAgICAgdGV4dEVsKCd0aCcsICdDYW1wYWlnbicpLAogICAgICB0ZXh0RWwoJ3RoJywgJ0NvbnRlbnQgVHlwZScpLAogICAgICB0ZXh0RWwoJ3RoJywgJ0NhcHRpb24nKSwKICAgICAgdGV4dEVsKCd0aCcsIG1ldHJpY0xhYmVsKG1ldHJpYyksICdudW0nKQogICAgKTsKICAgIGhlYWRUci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3RoJywgJycpKTsKICAgIHRoZWFkLmFwcGVuZENoaWxkKGhlYWRUcik7CgogICAgY29uc3QgdGJvZHkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0Ym9keScpOwogICAgcG9zdHMuZm9yRWFjaCgocCwgaSkgPT4gewogICAgICBjb25zdCB0ciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAgIGNvbnN0IG1ldGEgPSBwbGF0Zm9ybU9wdGlvbnMuZmluZCgocGwpID0+IHBsLmlkID09PSBwLnBsYXRmb3JtKSB8fCB7IGxhYmVsOiBwLnBsYXRmb3JtLCBjb2xvcjogJyM5OTknIH07CiAgICAgIGNvbnN0IHBsYXRmb3JtVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBjb25zdCBwaWxsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICBwaWxsLmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1waWxsJzsKICAgICAgY29uc3QgZG90ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICBkb3QuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLWRvdCc7CiAgICAgIGRvdC5zdHlsZS5iYWNrZ3JvdW5kID0gbWV0YS5jb2xvcjsKICAgICAgcGlsbC5hcHBlbmQoZG90LCBkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShtZXRhLmxhYmVsKSk7CiAgICAgIHBsYXRmb3JtVGQuYXBwZW5kQ2hpbGQocGlsbCk7CgogICAgICBjb25zdCBjYXB0aW9uID0gcC5jYXB0aW9uIHx8ICcobm8gY2FwdGlvbiknOwogICAgICBjb25zdCB0cnVuY2F0ZWQgPSBjYXB0aW9uLmxlbmd0aCA+IDYwID8gYCR7Y2FwdGlvbi5zbGljZSgwLCA2MCl94oCmYCA6IGNhcHRpb247CiAgICAgIGNvbnN0IGNhcHRpb25UZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIGlmIChwLnBvc3RpbmdfbGluaykgewogICAgICAgIGNvbnN0IGxpbmsgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdhJyk7CiAgICAgICAgbGluay5jbGFzc05hbWUgPSAnY2FwdGlvbi1saW5rJzsKICAgICAgICBsaW5rLmhyZWYgPSBwLnBvc3RpbmdfbGluazsKICAgICAgICBsaW5rLnRhcmdldCA9ICdfYmxhbmsnOwogICAgICAgIGxpbmsucmVsID0gJ25vb3BlbmVyIG5vcmVmZXJyZXInOwogICAgICAgIGxpbmsudGl0bGUgPSBjYXB0aW9uOwogICAgICAgIGxpbmsuYXBwZW5kQ2hpbGQoZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUodHJ1bmNhdGVkKSk7CiAgICAgICAgY2FwdGlvblRkLmFwcGVuZENoaWxkKGxpbmspOwogICAgICB9IGVsc2UgewogICAgICAgIGNhcHRpb25UZC5hcHBlbmRDaGlsZChkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZSh0cnVuY2F0ZWQpKTsKICAgICAgICBjYXB0aW9uVGQudGl0bGUgPSBjYXB0aW9uOwogICAgICB9CgogICAgICB0ci5hcHBlbmQoCiAgICAgICAgdGV4dEVsKCd0ZCcsIGAjJHtpICsgMX1gKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgRm9ybWF0LmRhdGUocC5wdWJsaXNoX2RhdGUpKSwKICAgICAgICBwbGF0Zm9ybVRkLAogICAgICAgIHRleHRFbCgndGQnLCBwLmNhbXBhaWduX3R5cGUgfHwgJ+KAlCcpLAogICAgICAgIHRleHRFbCgndGQnLCBwLmNvbnRlbnRfdHlwZSB8fCAn4oCUJyksCiAgICAgICAgY2FwdGlvblRkLAogICAgICAgIHRleHRFbCgndGQnLCBmb3JtYXRNZXRyaWNWYWx1ZShtZXRyaWMsIHAubWV0cmljX3ZhbHVlKSwgJ251bScpCiAgICAgICk7CgogICAgICBjb25zdCBhY3Rpb25UZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIGNvbnN0IHZpZXdCdG4gPSBpY29uQnRuKCdidG4nLCAnZXllJywgJ1ZpZXcgRGV0YWlscycpOwogICAgICB2aWV3QnRuLmRpc2FibGVkID0gIXAucmF3X3Jvd19pZDsKICAgICAgdmlld0J0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IFJlY29yZHMub3BlblZpZXcocC5yYXdfcm93X2lkKSk7CiAgICAgIGFjdGlvblRkLmFwcGVuZENoaWxkKHZpZXdCdG4pOwogICAgICB0ci5hcHBlbmRDaGlsZChhY3Rpb25UZCk7CgogICAgICB0Ym9keS5hcHBlbmRDaGlsZCh0cik7CiAgICB9KTsKICAgIHRhYmxlLmFwcGVuZCh0aGVhZCwgdGJvZHkpOwogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGFibGUpOwogIH0KCiAgLyoqIE1ldHJpYyAob3IgYW55IGZpbHRlcikgY2hhbmdlZCBidXQgdGhlIHBsYXRmb3JtIOKAlCBhbmQgdGhlcmVmb3JlIHRoZSBhdmFpbGFibGUgbWV0cmljIGxpc3Qg4oCUIGRpZG4ndDogbm8gbmVlZCB0byByZS1mZXRjaCBtZXRyaWMtb3B0aW9ucyBvciByZWJ1aWxkIHRoZSBzaGVsbCwganVzdCByZWZyZXNoIHRoZSBkYXRhLiAqLwogIGFzeW5jIGZ1bmN0aW9uIHJlZnJlc2hGb3JNZXRyaWMoKSB7CiAgICBjb25zdCBmaWx0ZXJzID0gU3RhdGUuZ2V0RmlsdGVycygpOwogICAgY29uc3QgW3N1bW1hcnksIGZvbGxvd2Vyc10gPSBhd2FpdCBQcm9taXNlLmFsbChbCiAgICAgIEFwaS5tZXRyaWNTdW1tYXJ5KHsgLi4uZmlsdGVycywgbWV0cmljIH0pLAogICAgICBBcGkuZm9sbG93ZXJzS3BpcyhmaWx0ZXJzKSwKICAgIF0pOwogICAgcmVuZGVyS3BpcyhzdW1tYXJ5LCBmb2xsb3dlcnMpOwogICAgYXdhaXQgUHJvbWlzZS5hbGwoWwogICAgICByZW5kZXJUcmVuZChmaWx0ZXJzKSwgcmVuZGVyQnJlYWtkb3duKGZpbHRlcnMpLCByZW5kZXJDb250ZW50VHlwZUJyZWFrZG93bihmaWx0ZXJzKSwgcmVuZGVyVG9wUG9zdHMoZmlsdGVycyksCiAgICBdKTsKICB9CgogIGZ1bmN0aW9uIHNob3dTa2VsZXRvbnMoKSB7CiAgICBjb25zdCBrcGlHcmlkID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2twaUdyaWQnKTsKICAgIGlmIChrcGlHcmlkKSB7IGtwaUdyaWQuaW5uZXJIVE1MID0gJyc7IGtwaUdyaWQuYXBwZW5kQ2hpbGQoc2tlbGV0b25TdGF0R3JpZCg4KSk7IH0KICAgIFsndHJlbmRDaGFydFdyYXAnLCAnYnJlYWtkb3duQ2hhcnRXcmFwJywgJ2NvbnRlbnRUeXBlQ2hhcnRXcmFwJ10uZm9yRWFjaCgoaWQpID0+IHsKICAgICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGlkKTsKICAgICAgaWYgKHdyYXApIHsgd3JhcC5pbm5lckhUTUwgPSAnJzsgd3JhcC5hcHBlbmRDaGlsZChza2VsZXRvbkNoYXJ0KCkpOyB9CiAgICB9KTsKICAgIGNvbnN0IHRvcFBvc3RzID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3RvcFBvc3RzVGFibGUnKTsKICAgIGlmICh0b3BQb3N0cykgeyB0b3BQb3N0cy5pbm5lckhUTUwgPSAnJzsgdG9wUG9zdHMuYXBwZW5kQ2hpbGQoc2tlbGV0b25Sb3dzKDYpKTsgfQogIH0KCiAgYXN5bmMgZnVuY3Rpb24gcmVuZGVyKCkgewogICAgcm9vdCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd2aWV3LWRhc2hib2FyZCcpOwogICAgY29uc3QgZmlsdGVycyA9IFN0YXRlLmdldEZpbHRlcnMoKTsKICAgIGNvbnN0IHsgb3B0aW9ucyB9ID0gYXdhaXQgQXBpLm1ldHJpY09wdGlvbnMoZmlsdGVycy5wbGF0Zm9ybSk7CiAgICBtZXRyaWNPcHRpb25zID0gb3B0aW9uczsKICAgIGlmICghbWV0cmljT3B0aW9ucy5zb21lKChtKSA9PiBtLmtleSA9PT0gbWV0cmljKSkgewogICAgICBtZXRyaWMgPSBtZXRyaWNPcHRpb25zLmxlbmd0aCA/IG1ldHJpY09wdGlvbnNbMF0ua2V5IDogJ3ZpZXdzJzsKICAgIH0KICAgIHNoZWxsKCk7CiAgICBzaG93U2tlbGV0b25zKCk7CiAgICBhd2FpdCByZWZyZXNoRm9yTWV0cmljKCk7CiAgfQoKICByZXR1cm4geyByZW5kZXIgfTsKfSkoKTsKCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBEYXRhIFJlY29yZHMgdGFiOiBhIENSTS1zdHlsZSwgcGxhdGZvcm0tZ3JvdXBlZCBicm93c2VyIGJhY2tlZAogICBieSBwb3N0cy9wb3N0X21ldHJpY3MgKHRoZSBzYW1lIG5vcm1hbGl6ZWQgZGF0YSB0aGUgZGFzaGJvYXJkLAogICBjb21wYXJpc29ucywgYW5kIHJlcG9ydHMgcmVhZCkg4oCUICJBbGwgUGxhdGZvcm1zIiBzaG93cyBhIGNvbW1vbgogICBjcm9zcy1wbGF0Zm9ybSBzdW1tYXJ5LCBhIHNwZWNpZmljIHBsYXRmb3JtIHNob3dzIG9ubHkgdGhhdAogICBwbGF0Zm9ybSdzIGN1cmF0ZWQgbWV0cmljcy4gRXZlcnkgZmllbGQgb2YgYSByZWNvcmQgKGV4YWN0bHkgYXMKICAgaW1wb3J0ZWQpIGlzIGFsd2F5cyByZWFjaGFibGUgdmlhIFZpZXcvRWRpdCByZWdhcmRsZXNzIG9mIHRoZQogICB0YWJsZSdzIGN1cmF0aW9uLCB3aGljaCByZWFkcyB0aGUgcmF3X3Jvd3MgbWlycm9yIGFuZCwgb24gc2F2ZSwKICAgcmUtc3luY3MgcG9zdHMvcG9zdF9tZXRyaWNzIHNvIGV2ZXJ5IHZpZXcgcmVmbGVjdHMgdGhlIGNoYW5nZQogICBpbW1lZGlhdGVseS4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmNvbnN0IFJlY29yZHMgPSAoKCkgPT4gewogIGxldCByb290OwogIGxldCBwYWdlID0gMTsKICBjb25zdCBwYWdlU2l6ZSA9IDI1OwogIGxldCBzZWFyY2hWYWx1ZSA9ICcnOwogIGxldCBzZWFyY2hEZWJvdW5jZSA9IG51bGw7CiAgbGV0IG1vZGFsU3RhdGUgPSBudWxsOyAvLyB7IHJlY29yZCwgdmFsdWVzOiBbLi4uXSB9IOKAlCBFZGl0IG1vZGFsIG9ubHkKICBsZXQgY3VycmVudFJlc3VsdCA9IG51bGw7IC8vIGxhc3QtbG9hZGVkIHBhZ2UsIGtlcHQgc28gc29ydGluZyBjYW4gcmUtcmVuZGVyIHdpdGhvdXQgYSBuZXR3b3JrIHJvdW5kLXRyaXAKICBsZXQgc29ydFN0YXRlID0geyBrZXk6IG51bGwsIGRpcjogJ2FzYycsIHR5cGU6ICdzdHJpbmcnIH07CgogIC8qKiBTb3J0cyBhIGNvcHkgb2YgYHJvd3NgIGJ5IGEgKHBvc3NpYmx5IGRvdHRlZCwgZS5nLiAibWV0cmljcy5yZWFjaCIpIGtleSBwYXRoLiBOdWxscyBhbHdheXMgc29ydCBsYXN0IHJlZ2FyZGxlc3Mgb2YgZGlyZWN0aW9uLiAqLwogIGZ1bmN0aW9uIHNvcnRSb3dzKHJvd3MsIGtleSwgZGlyLCB0eXBlKSB7CiAgICBjb25zdCBmYWN0b3IgPSBkaXIgPT09ICdhc2MnID8gMSA6IC0xOwogICAgY29uc3QgcmVhZCA9IChyb3cpID0+IGtleS5zcGxpdCgnLicpLnJlZHVjZSgobywgaykgPT4gKG8gPT09IG51bGwgfHwgbyA9PT0gdW5kZWZpbmVkID8gdW5kZWZpbmVkIDogb1trXSksIHJvdyk7CiAgICByZXR1cm4gWy4uLnJvd3NdLnNvcnQoKGEsIGIpID0+IHsKICAgICAgY29uc3QgYXYgPSByZWFkKGEpOwogICAgICBjb25zdCBidiA9IHJlYWQoYik7CiAgICAgIGNvbnN0IGFNaXNzaW5nID0gYXYgPT09IG51bGwgfHwgYXYgPT09IHVuZGVmaW5lZCB8fCBhdiA9PT0gJyc7CiAgICAgIGNvbnN0IGJNaXNzaW5nID0gYnYgPT09IG51bGwgfHwgYnYgPT09IHVuZGVmaW5lZCB8fCBidiA9PT0gJyc7CiAgICAgIGlmIChhTWlzc2luZyAmJiBiTWlzc2luZykgcmV0dXJuIDA7CiAgICAgIGlmIChhTWlzc2luZykgcmV0dXJuIDE7CiAgICAgIGlmIChiTWlzc2luZykgcmV0dXJuIC0xOwogICAgICBpZiAodHlwZSA9PT0gJ251bWJlcicpIHJldHVybiAoYXYgLSBidikgKiBmYWN0b3I7CiAgICAgIHJldHVybiBTdHJpbmcoYXYpLmxvY2FsZUNvbXBhcmUoU3RyaW5nKGJ2KSkgKiBmYWN0b3I7CiAgICB9KTsKICB9CgogIC8qKiBBIDx0aD4gdGhhdCB0b2dnbGVzIGFzY2VuZGluZy9kZXNjZW5kaW5nIG9uIGNsaWNrIGFuZCBzaG93cyBhbiBhcnJvdyBvbiB3aGljaGV2ZXIgY29sdW1uIGlzIGFjdGl2ZSDigJQgc29ydHMgdGhlIGFscmVhZHktbG9hZGVkIHBhZ2UgaW5zdGFudGx5LCBubyByZWxvYWQuICovCiAgZnVuY3Rpb24gc29ydGFibGVIZWFkZXIobGFiZWwsIGtleSwgdHlwZSkgewogICAgY29uc3QgdGggPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0aCcpOwogICAgaWYgKHR5cGUgPT09ICdudW1iZXInKSB0aC5jbGFzc05hbWUgPSAnbnVtJzsKICAgIHRoLmNsYXNzTGlzdC5hZGQoJ3NvcnRhYmxlLXRoJyk7CiAgICBjb25zdCBpc0FjdGl2ZSA9IHNvcnRTdGF0ZS5rZXkgPT09IGtleTsKICAgIHRoLmFwcGVuZENoaWxkKGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKGxhYmVsKSk7CiAgICB0aC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3NwYW4nLCBpc0FjdGl2ZSA/IChzb3J0U3RhdGUuZGlyID09PSAnYXNjJyA/ICcg4oaRJyA6ICcg4oaTJykgOiAnIOKGlScsICdzb3J0LWFycm93JykpOwogICAgdGguYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7CiAgICAgIHNvcnRTdGF0ZSA9IHsga2V5LCBkaXI6IHNvcnRTdGF0ZS5rZXkgPT09IGtleSAmJiBzb3J0U3RhdGUuZGlyID09PSAnYXNjJyA/ICdkZXNjJyA6ICdhc2MnLCB0eXBlIH07CiAgICAgIGlmIChjdXJyZW50UmVzdWx0KSByZW5kZXJUYWJsZShjdXJyZW50UmVzdWx0KTsKICAgIH0pOwogICAgcmV0dXJuIHRoOwogIH0KCiAgZnVuY3Rpb24gcGxhdGZvcm1NZXRhKCkgewogICAgcmV0dXJuICh3aW5kb3cuX19maWx0ZXJPcHRpb25zQ2FjaGUgfHwgeyBwbGF0Zm9ybXM6IFtdIH0pLnBsYXRmb3JtczsKICB9CgogIGZ1bmN0aW9uIHBsYXRmb3JtTGFiZWwoaWQpIHsKICAgIGNvbnN0IG0gPSBwbGF0Zm9ybU1ldGEoKS5maW5kKChwKSA9PiBwLmlkID09PSBpZCk7CiAgICByZXR1cm4gbSA/IG0ubGFiZWwgOiBpZDsKICB9CgogIGZ1bmN0aW9uIHNoZWxsKCkgewogICAgcm9vdC5pbm5lckhUTUwgPSAnJzsKICAgIHJvb3QuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCAnRGF0YSBSZWNvcmRzJywgJ3NlY3Rpb24tdGl0bGUnKSk7CiAgICByb290LmFwcGVuZENoaWxkKHRleHRFbCgKICAgICAgJ2RpdicsCiAgICAgICdCcm93c2UgYnkgcGxhdGZvcm0gdG8gc2VlIG9ubHkgaXRzIG1ldHJpY3MsIG9yIHN0YXkgb24gQWxsIFBsYXRmb3JtcyBmb3IgYSBjcm9zcy1wbGF0Zm9ybSBzdW1tYXJ5LiBFdmVyeSByZWNvcmQgc3RheXMgZnVsbHkgZWRpdGFibGUg4oCUIFZpZXcgb3IgRWRpdCBhbHdheXMgb3BlbnMgZXZlcnkgZmllbGQgaW1wb3J0ZWQgZnJvbSB0aGUgc3ByZWFkc2hlZXQsIG5vdCBqdXN0IHdoYXTigJlzIGluIHRoZSB0YWJsZS4nLAogICAgICAnbXV0ZWQnCiAgICApKTsKCiAgICBjb25zdCB0b29sYmFyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0b29sYmFyLmNsYXNzTmFtZSA9ICdyZWNvcmRzLXRvb2xiYXInOwogICAgY29uc3QgcGlsbHMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHBpbGxzLmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1maWx0ZXItcGlsbHMnOwogICAgcGlsbHMuaWQgPSAncmVjb3Jkc1BsYXRmb3JtUGlsbHMnOwogICAgY29uc3Qgc2VhcmNoID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBzZWFyY2guY2xhc3NOYW1lID0gJ3JlY29yZHMtc2VhcmNoJzsKICAgIGNvbnN0IHNlYXJjaElucHV0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsKICAgIHNlYXJjaElucHV0LnR5cGUgPSAnc2VhcmNoJzsKICAgIHNlYXJjaElucHV0LnBsYWNlaG9sZGVyID0gJ1NlYXJjaCBjYXB0aW9ucywgY2FtcGFpZ25zLCBjb250ZW50IHR5cGXigKYnOwogICAgc2VhcmNoSW5wdXQuaWQgPSAncmVjb3Jkc1NlYXJjaElucHV0JzsKICAgIHNlYXJjaElucHV0LnZhbHVlID0gc2VhcmNoVmFsdWU7CiAgICBzZWFyY2hJbnB1dC5hZGRFdmVudExpc3RlbmVyKCdpbnB1dCcsICgpID0+IHsKICAgICAgY2xlYXJUaW1lb3V0KHNlYXJjaERlYm91bmNlKTsKICAgICAgc2VhcmNoRGVib3VuY2UgPSBzZXRUaW1lb3V0KCgpID0+IHsKICAgICAgICBzZWFyY2hWYWx1ZSA9IHNlYXJjaElucHV0LnZhbHVlOwogICAgICAgIHBhZ2UgPSAxOwogICAgICAgIGxvYWQoKTsKICAgICAgfSwgMzAwKTsKICAgIH0pOwogICAgc2VhcmNoLmFwcGVuZENoaWxkKHNlYXJjaElucHV0KTsKICAgIHRvb2xiYXIuYXBwZW5kKHBpbGxzLCBzZWFyY2gpOwogICAgcm9vdC5hcHBlbmRDaGlsZCh0b29sYmFyKTsKCiAgICBjb25zdCBleHBvcnRSb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGV4cG9ydFJvdy5jbGFzc05hbWUgPSAnZXhwb3J0LWJ1dHRvbnMnOwogICAgY29uc3QgZXhwb3J0Q3N2QnRuID0gaWNvbkJ0bignYnRuJywgJ2ZpbGUtZG93bicsICdFeHBvcnQgQ1NWJyk7CiAgICBleHBvcnRDc3ZCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB0cmlnZ2VyUmVjb3Jkc0V4cG9ydCgnY3N2JykpOwogICAgY29uc3QgZXhwb3J0WGxzeEJ0biA9IGljb25CdG4oJ2J0bicsICdmaWxlLXNwcmVhZHNoZWV0JywgJ0V4cG9ydCBFeGNlbCcpOwogICAgZXhwb3J0WGxzeEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHRyaWdnZXJSZWNvcmRzRXhwb3J0KCd4bHN4JykpOwogICAgY29uc3QgZGVkdXBlQnRuID0gaWNvbkJ0bignYnRuJywgJ2NvcHkteCcsICdSZW1vdmUgZHVwbGljYXRlcycpOwogICAgZGVkdXBlQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gaGFuZGxlRGVkdXBlKGRlZHVwZUJ0bikpOwogICAgZXhwb3J0Um93LmFwcGVuZChleHBvcnRDc3ZCdG4sIGV4cG9ydFhsc3hCdG4sIGRlZHVwZUJ0bik7CiAgICByb290LmFwcGVuZENoaWxkKGV4cG9ydFJvdyk7CgogICAgY29uc3QgY2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY2FyZC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICBjb25zdCB0YWJsZVdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRhYmxlV3JhcC5jbGFzc05hbWUgPSAndGFibGUtc2Nyb2xsJzsKICAgIHRhYmxlV3JhcC5pZCA9ICdyZWNvcmRzVGFibGVXcmFwJzsKICAgIGNhcmQuYXBwZW5kQ2hpbGQodGFibGVXcmFwKTsKICAgIGNvbnN0IHBhZ2VyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBwYWdlci5jbGFzc05hbWUgPSAncGFnaW5hdGlvbi1yb3cnOwogICAgcGFnZXIuaWQgPSAncmVjb3Jkc1BhZ2VyJzsKICAgIGNhcmQuYXBwZW5kQ2hpbGQocGFnZXIpOwogICAgcm9vdC5hcHBlbmRDaGlsZChjYXJkKTsKCiAgICByZW5kZXJQaWxscygpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyUGlsbHMoKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3JlY29yZHNQbGF0Zm9ybVBpbGxzJyk7CiAgICBpZiAoIXdyYXApIHJldHVybjsKICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CiAgICBjb25zdCBjdXJyZW50ID0gU3RhdGUuZ2V0RmlsdGVycygpLnBsYXRmb3JtIHx8ICdhbGwnOwogICAgY29uc3Qgb3B0aW9ucyA9IFt7IGlkOiAnYWxsJywgbGFiZWw6ICdBbGwgUGxhdGZvcm1zJywgY29sb3I6IG51bGwgfSwgLi4ucGxhdGZvcm1NZXRhKCldOwogICAgb3B0aW9ucy5mb3JFYWNoKChvcHQpID0+IHsKICAgICAgY29uc3QgYnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICAgIGJ0bi50eXBlID0gJ2J1dHRvbic7CiAgICAgIGJ0bi5jbGFzc0xpc3QudG9nZ2xlKCdpcy1hY3RpdmUnLCBjdXJyZW50ID09PSBvcHQuaWQpOwogICAgICBpZiAob3B0LmNvbG9yKSB7CiAgICAgICAgY29uc3QgZG90ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICAgIGRvdC5jbGFzc05hbWUgPSAncGxhdGZvcm0tZG90JzsKICAgICAgICBkb3Quc3R5bGUuYmFja2dyb3VuZCA9IG9wdC5jb2xvcjsKICAgICAgICBidG4uYXBwZW5kQ2hpbGQoZG90KTsKICAgICAgfQogICAgICBidG4uYXBwZW5kQ2hpbGQoZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUob3B0LmxhYmVsKSk7CiAgICAgIGJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsKICAgICAgICBpZiAoY3VycmVudCA9PT0gb3B0LmlkKSByZXR1cm47CiAgICAgICAgY29uc3QgZmlsdGVyU2VsZWN0ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZpbHRlclBsYXRmb3JtJyk7CiAgICAgICAgaWYgKGZpbHRlclNlbGVjdCkgZmlsdGVyU2VsZWN0LnZhbHVlID0gb3B0LmlkOwogICAgICAgIHBhZ2UgPSAxOwogICAgICAgIFN0YXRlLnNldEZpbHRlcnMoeyBwbGF0Zm9ybTogb3B0LmlkIH0pOwogICAgICB9KTsKICAgICAgd3JhcC5hcHBlbmRDaGlsZChidG4pOwogICAgfSk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiBsb2FkKCkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdyZWNvcmRzVGFibGVXcmFwJyk7CiAgICBpZiAod3JhcCkgeyB3cmFwLmlubmVySFRNTCA9ICcnOyB3cmFwLmFwcGVuZENoaWxkKHNrZWxldG9uUm93cyg4KSk7IH0KICAgIGNvbnN0IGZpbHRlcnMgPSBTdGF0ZS5nZXRGaWx0ZXJzKCk7CiAgICBjb25zdCByZXN1bHQgPSBhd2FpdCBBcGkucmVjb3Jkc1RhYmxlKHsgLi4uZmlsdGVycywgc2VhcmNoOiBzZWFyY2hWYWx1ZSwgcGFnZSwgcGFnZVNpemUgfSk7CiAgICByZW5kZXJUYWJsZShyZXN1bHQpOwogICAgcmVuZGVyUGFnZXIocmVzdWx0KTsKICB9CgogIC8qKiBSZWNvcmRzIGlzIHNlcnZlci1wYWdpbmF0ZWQvc2VhcmNoZWQsIHNvIGl0cyBleHBvcnQgaXMgYSBkaXJlY3QgbmF2aWdhdGlvbiB0byBhIGJhY2tlbmQgcm91dGUgdGhhdCByZXVzZXMgdGhlIGV4YWN0IHNhbWUgZmlsdGVyLWJ1aWxkaW5nIHRoZSBsaXN0IGVuZHBvaW50IGRvZXMg4oCUIGV4cG9ydHMgdGhlIGZ1bGwgbWF0Y2hpbmcgZGF0YXNldCwgbm90IGp1c3QgdGhlIGN1cnJlbnQgcGFnZS4gKi8KICBmdW5jdGlvbiB0cmlnZ2VyUmVjb3Jkc0V4cG9ydChmb3JtYXQpIHsKICAgIGNvbnN0IGZpbHRlcnMgPSBTdGF0ZS5nZXRGaWx0ZXJzKCk7CiAgICBjb25zdCBwYXJhbXMgPSBuZXcgVVJMU2VhcmNoUGFyYW1zKHsgLi4uZmlsdGVycywgc2VhcmNoOiBzZWFyY2hWYWx1ZSwgZm9ybWF0IH0pOwogICAgd2luZG93LmxvY2F0aW9uLmhyZWYgPSBgL2FwaS9yZWNvcmRzL2V4cG9ydD8ke3BhcmFtcy50b1N0cmluZygpfWA7CiAgfQoKICBmdW5jdGlvbiBjb2x1bW5MYWJlbHNGb3IocmVjb3JkKSB7CiAgICByZXR1cm4gcmVjb3JkLmhlYWRlcnMgJiYgcmVjb3JkLmhlYWRlcnMubGVuZ3RoCiAgICAgID8gcmVjb3JkLmhlYWRlcnMubWFwKChoKSA9PiAoaCAmJiBoLnRyaW0oKSA/IGggOiAnKHVubGFiZWxlZCBjb2x1bW4pJykpCiAgICAgIDogcmVjb3JkLnZhbHVlcy5tYXAoKF8sIGkpID0+IGBDb2x1bW4gJHtpICsgMX1gKTsKICB9CgogIC8qKiBHcm91cHMgYSByYXcgcmVjb3JkJ3MgZmllbGRzIGJ5IHRoZSBxdWFsaWZpZWQgaGVhZGVyJ3MgcGxhdGZvcm0tZ3JvdXAgcHJlZml4CiAgICAgIChlLmcuICJGQUNFQk9PSyDigJQgVmlld3MiKSwgc28gdGhlIFZpZXcvRWRpdCBwb3B1cCByZWFkcyBhcyBzZWN0aW9ucyBpbnN0ZWFkCiAgICAgIG9mIG9uZSBsb25nIGZsYXQgbGlzdCDigJQgZmFsbHMgYmFjayB0byBhIHNpbmdsZSAiRGV0YWlscyIgc2VjdGlvbiBmb3IKICAgICAgaWRlbnRpZmllciBjb2x1bW5zIGFuZCBmb3IgdGhlIHNpbXBsZSAob25lLXBsYXRmb3JtLXBlci1yb3cpIGZvcm1hdC4gKi8KICBmdW5jdGlvbiBncm91cEZpZWxkUm93cyhsYWJlbHMsIHZhbHVlcykgewogICAgY29uc3QgZ3JvdXBzID0gW107CiAgICBjb25zdCBpbmRleCA9IG5ldyBNYXAoKTsKICAgIGxhYmVscy5mb3JFYWNoKChsYWJlbCwgaWR4KSA9PiB7CiAgICAgIGNvbnN0IHNlcElkeCA9IGxhYmVsLmluZGV4T2YoJyDigJQgJyk7CiAgICAgIGNvbnN0IGdyb3VwTmFtZSA9IHNlcElkeCA+PSAwID8gbGFiZWwuc2xpY2UoMCwgc2VwSWR4KSA6ICdEZXRhaWxzJzsKICAgICAgY29uc3QgZmllbGRMYWJlbCA9IHNlcElkeCA+PSAwID8gbGFiZWwuc2xpY2Uoc2VwSWR4ICsgMykgOiBsYWJlbDsKICAgICAgaWYgKCFpbmRleC5oYXMoZ3JvdXBOYW1lKSkgewogICAgICAgIGluZGV4LnNldChncm91cE5hbWUsIHsgZ3JvdXA6IGdyb3VwTmFtZSwgZmllbGRzOiBbXSB9KTsKICAgICAgICBncm91cHMucHVzaChpbmRleC5nZXQoZ3JvdXBOYW1lKSk7CiAgICAgIH0KICAgICAgaW5kZXguZ2V0KGdyb3VwTmFtZSkuZmllbGRzLnB1c2goeyBpZHgsIGxhYmVsOiBmaWVsZExhYmVsIHx8IGBDb2x1bW4gJHtpZHggKyAxfWAsIHZhbHVlOiB2YWx1ZXNbaWR4XSB9KTsKICAgIH0pOwogICAgcmV0dXJuIGdyb3VwczsKICB9CgogIGZ1bmN0aW9uIHBsYXRmb3JtQmFkZ2VzKGlkcykgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgd3JhcC5zdHlsZS5kaXNwbGF5ID0gJ2ZsZXgnOwogICAgd3JhcC5zdHlsZS5mbGV4V3JhcCA9ICd3cmFwJzsKICAgIHdyYXAuc3R5bGUuZ2FwID0gJzRweCc7CiAgICBpZiAoIWlkcy5sZW5ndGgpIHJldHVybiB0ZXh0RWwoJ3NwYW4nLCAn4oCUJywgJ211dGVkJyk7CiAgICBjb25zdCBtZXRhID0gcGxhdGZvcm1NZXRhKCk7CiAgICBpZHMuZm9yRWFjaCgoaWQpID0+IHsKICAgICAgY29uc3QgbSA9IG1ldGEuZmluZCgocCkgPT4gcC5pZCA9PT0gaWQpIHx8IHsgbGFiZWw6IGlkLCBjb2xvcjogJyM5OTknIH07CiAgICAgIGNvbnN0IHBpbGwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICAgIHBpbGwuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLXBpbGwnOwogICAgICBjb25zdCBkb3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICAgIGRvdC5jbGFzc05hbWUgPSAncGxhdGZvcm0tZG90JzsKICAgICAgZG90LnN0eWxlLmJhY2tncm91bmQgPSBtLmNvbG9yOwogICAgICBwaWxsLmFwcGVuZChkb3QsIGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKG0ubGFiZWwpKTsKICAgICAgd3JhcC5hcHBlbmRDaGlsZChwaWxsKTsKICAgIH0pOwogICAgcmV0dXJuIHdyYXA7CiAgfQoKICBmdW5jdGlvbiBzdGF0dXNQaWxsKHN0YXR1cykgewogICAgY29uc3Qgc3BhbiA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgIHNwYW4uY2xhc3NOYW1lID0gYHN0YXR1cy1waWxsICR7c3RhdHVzfWA7CiAgICBzcGFuLnRleHRDb250ZW50ID0gc3RhdHVzID09PSAnZWRpdGVkJyA/ICdFZGl0ZWQnIDogJ09yaWdpbmFsJzsKICAgIHJldHVybiBzcGFuOwogIH0KCiAgZnVuY3Rpb24gbWV0cmljQ2VsbChrZXksIHZhbHVlKSB7CiAgICBpZiAoa2V5ID09PSAncG9zdGluZ19saW5rJykgewogICAgICBjb25zdCB0ZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIHRkLmNsYXNzTmFtZSA9ICdsaW5rLWNlbGwnOwogICAgICBpZiAodmFsdWUpIHsKICAgICAgICBjb25zdCBhID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYScpOwogICAgICAgIGEuaHJlZiA9IHZhbHVlOwogICAgICAgIGEudGFyZ2V0ID0gJ19ibGFuayc7CiAgICAgICAgYS5yZWwgPSAnbm9vcGVuZXIgbm9yZWZlcnJlcic7CiAgICAgICAgYS50ZXh0Q29udGVudCA9ICdPcGVuIOKGlyc7CiAgICAgICAgdGQuYXBwZW5kQ2hpbGQoYSk7CiAgICAgIH0gZWxzZSB7CiAgICAgICAgdGQuYXBwZW5kQ2hpbGQoZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUoJ+KAlCcpKTsKICAgICAgfQogICAgICByZXR1cm4gdGQ7CiAgICB9CiAgICBjb25zdCBkaXNwbGF5ID0ga2V5ID09PSAnd2F0Y2hfdGltZV9zZWNvbmRzJyA/IEZvcm1hdC5kdXJhdGlvbih2YWx1ZSkgOiBGb3JtYXQubnVtYmVyKHZhbHVlKTsKICAgIHJldHVybiB0ZXh0RWwoJ3RkJywgZGlzcGxheSwgJ251bScpOwogIH0KCiAgZnVuY3Rpb24gYWN0aW9uQnV0dG9ucyhyb3csIHBsYXRmb3JtKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB3cmFwLmNsYXNzTmFtZSA9ICdyb3ctYWN0aW9ucyc7CiAgICBjb25zdCB2aWV3QnRuID0gaWNvbkJ0bignYnRuJywgJ2V5ZScsICdWaWV3Jyk7CiAgICB2aWV3QnRuLmRpc2FibGVkID0gIXJvdy5yYXdSb3dJZDsKICAgIHZpZXdCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBvcGVuVmlldyhyb3cucmF3Um93SWQpKTsKICAgIGNvbnN0IGVkaXRCdG4gPSBpY29uQnRuKCdidG4nLCAncGVuY2lsJywgJ0VkaXQnKTsKICAgIGVkaXRCdG4uZGlzYWJsZWQgPSAhcm93LnJhd1Jvd0lkOwogICAgZWRpdEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IG9wZW5FZGl0b3Iocm93LnJhd1Jvd0lkKSk7CiAgICBjb25zdCBkZWxldGVCdG4gPSBpY29uQnRuKCdidG4gZGFuZ2VyJywgJ3RyYXNoLTInLCAnRGVsZXRlJyk7CiAgICBkZWxldGVCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBoYW5kbGVEZWxldGUocm93LCBwbGF0Zm9ybSkpOwogICAgd3JhcC5hcHBlbmQodmlld0J0biwgZWRpdEJ0biwgZGVsZXRlQnRuKTsKICAgIHJldHVybiB3cmFwOwogIH0KCiAgZnVuY3Rpb24gY2FwdGlvbkNlbGwoY2FwdGlvbikgewogICAgY29uc3QgdGV4dCA9IGNhcHRpb24gfHwgJyhubyBjYXB0aW9uKSc7CiAgICByZXR1cm4gdGV4dEVsKCd0ZCcsIHRleHQubGVuZ3RoID4gNzAgPyBgJHt0ZXh0LnNsaWNlKDAsIDcwKX3igKZgIDogdGV4dCk7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJTdW1tYXJ5VGFibGUocmVzdWx0KSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3JlY29yZHNUYWJsZVdyYXAnKTsKICAgIGlmICghcmVzdWx0LnJvd3MubGVuZ3RoKSB7CiAgICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQoZW1wdHlTdGF0ZSh7CiAgICAgICAgaWNvbjogJ2RhdGFiYXNlJywKICAgICAgICB0aXRsZTogJ05vIHJlY29yZHMgbWF0Y2ggdGhlc2UgZmlsdGVycyB5ZXQnLAogICAgICAgIG1lc3NhZ2U6ICdVcGxvYWQgYSB3ZWVrbHkgZXhwb3J0LCBvciB3aWRlbiB0aGUgZGF0ZSByYW5nZSwgdG8gc2VlIHJlY29yZHMgaGVyZS4nLAogICAgICAgIGFjdGlvbkxhYmVsOiAnVXBsb2FkIGRhdGEnLAogICAgICAgIG9uQWN0aW9uOiAoKSA9PiBkb2N1bWVudC5xdWVyeVNlbGVjdG9yKCcudGFiLWJ0bltkYXRhLXRhYj0idXBsb2FkIl0nKT8uY2xpY2soKSwKICAgICAgfSkpOwogICAgICByZXR1cm47CiAgICB9CiAgICBjb25zdCB0YWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICB0YWJsZS5jbGFzc05hbWUgPSAnZGF0YS10YWJsZSc7CiAgICBjb25zdCB0aGVhZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RoZWFkJyk7CiAgICBjb25zdCBoZWFkVHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgaGVhZFRyLmFwcGVuZCgKICAgICAgc29ydGFibGVIZWFkZXIoJ0RhdGUnLCAncHVibGlzaERhdGUnLCAnc3RyaW5nJyksCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdQbGF0Zm9ybXMnLCAncGxhdGZvcm1JZHMuMCcsICdzdHJpbmcnKSwKICAgICAgdGV4dEVsKCd0aCcsICdDYXB0aW9uJyksCiAgICAgIHRleHRFbCgndGgnLCAnQ2FtcGFpZ24nKSwKICAgICAgdGV4dEVsKCd0aCcsICdDb250ZW50IFR5cGUnKSwKICAgICAgdGV4dEVsKCd0aCcsICdTdGF0dXMnKSwKICAgICAgc29ydGFibGVIZWFkZXIoJ0xhc3QgVXBkYXRlZCcsICd1cGRhdGVkQXQnLCAnc3RyaW5nJyksCiAgICAgIHRleHRFbCgndGgnLCAnQWN0aW9ucycpCiAgICApOwogICAgdGhlYWQuYXBwZW5kQ2hpbGQoaGVhZFRyKTsKICAgIGNvbnN0IHRib2R5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGJvZHknKTsKICAgIGNvbnN0IHJvd3MgPSBzb3J0U3RhdGUua2V5ID8gc29ydFJvd3MocmVzdWx0LnJvd3MsIHNvcnRTdGF0ZS5rZXksIHNvcnRTdGF0ZS5kaXIsIHNvcnRTdGF0ZS50eXBlKSA6IHJlc3VsdC5yb3dzOwogICAgcm93cy5mb3JFYWNoKChyKSA9PiB7CiAgICAgIGNvbnN0IHRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgY29uc3QgcGxhdGZvcm1zVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBwbGF0Zm9ybXNUZC5hcHBlbmRDaGlsZChwbGF0Zm9ybUJhZGdlcyhyLnBsYXRmb3JtSWRzKSk7CiAgICAgIGNvbnN0IHN0YXR1c1RkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgc3RhdHVzVGQuYXBwZW5kQ2hpbGQoc3RhdHVzUGlsbChyLnN0YXR1cykpOwogICAgICBjb25zdCBhY3Rpb25zVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBhY3Rpb25zVGQuYXBwZW5kQ2hpbGQoYWN0aW9uQnV0dG9ucyhyLCAnYWxsJykpOwogICAgICB0ci5hcHBlbmQoCiAgICAgICAgdGV4dEVsKCd0ZCcsIEZvcm1hdC5kYXRlKHIucHVibGlzaERhdGUpKSwKICAgICAgICBwbGF0Zm9ybXNUZCwKICAgICAgICBjYXB0aW9uQ2VsbChyLmNhcHRpb24pLAogICAgICAgIHRleHRFbCgndGQnLCByLmNhbXBhaWduVHlwZSB8fCAn4oCUJyksCiAgICAgICAgdGV4dEVsKCd0ZCcsIHIuY29udGVudFR5cGUgfHwgJ+KAlCcpLAogICAgICAgIHN0YXR1c1RkLAogICAgICAgIHRleHRFbCgndGQnLCByLnVwZGF0ZWRBdCksCiAgICAgICAgYWN0aW9uc1RkCiAgICAgICk7CiAgICAgIHRib2R5LmFwcGVuZENoaWxkKHRyKTsKICAgIH0pOwogICAgdGFibGUuYXBwZW5kKHRoZWFkLCB0Ym9keSk7CiAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgd3JhcC5hcHBlbmRDaGlsZCh0YWJsZSk7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJQbGF0Zm9ybVRhYmxlKHJlc3VsdCkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdyZWNvcmRzVGFibGVXcmFwJyk7CiAgICBpZiAoIXJlc3VsdC5yb3dzLmxlbmd0aCkgewogICAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgICB3cmFwLmFwcGVuZENoaWxkKGVtcHR5U3RhdGUoewogICAgICAgIGljb246ICdkYXRhYmFzZScsCiAgICAgICAgdGl0bGU6IGBObyAke3BsYXRmb3JtTGFiZWwocmVzdWx0LnBsYXRmb3JtKX0gcmVjb3JkcyBtYXRjaCB0aGVzZSBmaWx0ZXJzIHlldGAsCiAgICAgICAgbWVzc2FnZTogJ1RyeSBhIGRpZmZlcmVudCBwbGF0Zm9ybSwgb3Igd2lkZW4gdGhlIGRhdGUgcmFuZ2UuJywKICAgICAgfSkpOwogICAgICByZXR1cm47CiAgICB9CiAgICBjb25zdCB0YWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICB0YWJsZS5jbGFzc05hbWUgPSAnZGF0YS10YWJsZSc7CiAgICBjb25zdCB0aGVhZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RoZWFkJyk7CiAgICBjb25zdCBoZWFkVHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgaGVhZFRyLmFwcGVuZChzb3J0YWJsZUhlYWRlcignRGF0ZScsICdwdWJsaXNoRGF0ZScsICdzdHJpbmcnKSwgdGV4dEVsKCd0aCcsICdDYXB0aW9uJyksIHRleHRFbCgndGgnLCAnQ2FtcGFpZ24nKSwgdGV4dEVsKCd0aCcsICdDb250ZW50IFR5cGUnKSk7CiAgICByZXN1bHQuY29sdW1ucy5mb3JFYWNoKChjKSA9PiB7CiAgICAgIGlmIChjLmtleSA9PT0gJ3Bvc3RpbmdfbGluaycpIHsKICAgICAgICBoZWFkVHIuYXBwZW5kQ2hpbGQodGV4dEVsKCd0aCcsIGMubGFiZWwpKTsKICAgICAgfSBlbHNlIHsKICAgICAgICBoZWFkVHIuYXBwZW5kQ2hpbGQoc29ydGFibGVIZWFkZXIoYy5sYWJlbCwgYG1ldHJpY3MuJHtjLmtleX1gLCAnbnVtYmVyJykpOwogICAgICB9CiAgICB9KTsKICAgIGhlYWRUci5hcHBlbmQodGV4dEVsKCd0aCcsICdTdGF0dXMnKSwgdGV4dEVsKCd0aCcsICdBY3Rpb25zJykpOwogICAgdGhlYWQuYXBwZW5kQ2hpbGQoaGVhZFRyKTsKICAgIGNvbnN0IHRib2R5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGJvZHknKTsKICAgIGNvbnN0IHJvd3MgPSBzb3J0U3RhdGUua2V5ID8gc29ydFJvd3MocmVzdWx0LnJvd3MsIHNvcnRTdGF0ZS5rZXksIHNvcnRTdGF0ZS5kaXIsIHNvcnRTdGF0ZS50eXBlKSA6IHJlc3VsdC5yb3dzOwogICAgcm93cy5mb3JFYWNoKChyKSA9PiB7CiAgICAgIGNvbnN0IHRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgdHIuYXBwZW5kKHRleHRFbCgndGQnLCBGb3JtYXQuZGF0ZShyLnB1Ymxpc2hEYXRlKSksIGNhcHRpb25DZWxsKHIuY2FwdGlvbiksIHRleHRFbCgndGQnLCByLmNhbXBhaWduVHlwZSB8fCAn4oCUJyksIHRleHRFbCgndGQnLCByLmNvbnRlbnRUeXBlIHx8ICfigJQnKSk7CiAgICAgIHJlc3VsdC5jb2x1bW5zLmZvckVhY2goKGMpID0+IHRyLmFwcGVuZENoaWxkKG1ldHJpY0NlbGwoYy5rZXksIHIubWV0cmljc1tjLmtleV0pKSk7CiAgICAgIGNvbnN0IHN0YXR1c1RkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgc3RhdHVzVGQuYXBwZW5kQ2hpbGQoc3RhdHVzUGlsbChyLnN0YXR1cykpOwogICAgICB0ci5hcHBlbmRDaGlsZChzdGF0dXNUZCk7CiAgICAgIGNvbnN0IGFjdGlvbnNUZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIGFjdGlvbnNUZC5hcHBlbmRDaGlsZChhY3Rpb25CdXR0b25zKHIsIHJlc3VsdC5wbGF0Zm9ybSkpOwogICAgICB0ci5hcHBlbmRDaGlsZChhY3Rpb25zVGQpOwogICAgICB0Ym9keS5hcHBlbmRDaGlsZCh0cik7CiAgICB9KTsKICAgIHRhYmxlLmFwcGVuZCh0aGVhZCwgdGJvZHkpOwogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGFibGUpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyVGFibGUocmVzdWx0KSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3JlY29yZHNUYWJsZVdyYXAnKTsKICAgIGlmICghd3JhcCkgcmV0dXJuOwogICAgY3VycmVudFJlc3VsdCA9IHJlc3VsdDsKICAgIGlmIChyZXN1bHQucGxhdGZvcm0gPT09ICdhbGwnKSByZW5kZXJTdW1tYXJ5VGFibGUocmVzdWx0KTsKICAgIGVsc2UgcmVuZGVyUGxhdGZvcm1UYWJsZShyZXN1bHQpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyUGFnZXIocmVzdWx0KSB7CiAgICBjb25zdCBwYWdlciA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdyZWNvcmRzUGFnZXInKTsKICAgIGlmICghcGFnZXIpIHJldHVybjsKICAgIHBhZ2VyLmlubmVySFRNTCA9ICcnOwogICAgY29uc3QgdG90YWxQYWdlcyA9IE1hdGgubWF4KDEsIE1hdGguY2VpbChyZXN1bHQudG90YWwgLyByZXN1bHQucGFnZVNpemUpKTsKICAgIGNvbnN0IHByZXZCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIHByZXZCdG4uY2xhc3NOYW1lID0gJ2J0bic7CiAgICBwcmV2QnRuLnRleHRDb250ZW50ID0gJ1ByZXZpb3VzJzsKICAgIHByZXZCdG4uZGlzYWJsZWQgPSByZXN1bHQucGFnZSA8PSAxOwogICAgcHJldkJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsgcGFnZSAtPSAxOyBsb2FkKCk7IH0pOwogICAgY29uc3QgbmV4dEJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgbmV4dEJ0bi5jbGFzc05hbWUgPSAnYnRuJzsKICAgIG5leHRCdG4udGV4dENvbnRlbnQgPSAnTmV4dCc7CiAgICBuZXh0QnRuLmRpc2FibGVkID0gcmVzdWx0LnBhZ2UgPj0gdG90YWxQYWdlczsKICAgIG5leHRCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7IHBhZ2UgKz0gMTsgbG9hZCgpOyB9KTsKICAgIHBhZ2VyLmFwcGVuZChwcmV2QnRuLCB0ZXh0RWwoJ3NwYW4nLCBgUGFnZSAke3Jlc3VsdC5wYWdlfSBvZiAke3RvdGFsUGFnZXN9IOKAlCAke3Jlc3VsdC50b3RhbH0gcmVjb3JkKHMpYCksIG5leHRCdG4pOwogIH0KCiAgLyoqICJSZW1vdmUgZHVwbGljYXRlcyI6IGNvbGxhcHNlIHJlY29yZHMgdGhhdCBzaGFyZSBhIHB1Ymxpc2ggZGF0ZSArIGNhcHRpb24KICAgICAgKHRoZSBjb3BpZXMgbGVmdCBiZWhpbmQgd2hlbiBhIHdlZWtseSBzaGVldCBpcyByZS11cGxvYWRlZCBhbmQgYSBwb3N0J3MKICAgICAgbnVtYmVycyBjaGFuZ2VkKSwga2VlcGluZyB0aGUgbW9zdCByZWNlbnRseSBpbXBvcnRlZCBvbmUuIFByZXZpZXdzIHRoZQogICAgICBjb3VudCBmaXJzdCwgdGhlbiBkZWxldGVzIG9uIGNvbmZpcm0g4oCUIHRoZSBvcmlnaW5hbHMgc3RheSBpbiBVcGxvYWQKICAgICAgSGlzdG9yeSdzIHJhdy1yb3cgdmlld2VyIGVpdGhlciB3YXkuICovCiAgYXN5bmMgZnVuY3Rpb24gaGFuZGxlRGVkdXBlKGJ0bikgewogICAgaWYgKGJ0biAmJiBidG4uZGlzYWJsZWQpIHJldHVybjsKICAgIHRyeSB7CiAgICAgIGlmIChidG4pIGJ0bi5kaXNhYmxlZCA9IHRydWU7CiAgICAgIGNvbnN0IHByZXZpZXcgPSBhd2FpdCBBcGkuZHVwbGljYXRlUmVjb3Jkc1ByZXZpZXcoKTsKICAgICAgaWYgKCFwcmV2aWV3LnJlbW92ZUNvdW50KSB7CiAgICAgICAgVG9hc3Quc2hvdygnTm8gZHVwbGljYXRlIHJlY29yZHMgZm91bmQuJywgJ3N1Y2Nlc3MnKTsKICAgICAgICByZXR1cm47CiAgICAgIH0KICAgICAgY29uc3QgZWRpdGVkTm90ZSA9IHByZXZpZXcuZWRpdGVkSW5SZW1vdmVDb3VudAogICAgICAgID8gYFxuXG4ke3ByZXZpZXcuZWRpdGVkSW5SZW1vdmVDb3VudH0gb2YgdGhlIGNvcGllcyB0byBiZSByZW1vdmVkIHdlcmUgaGFuZC1lZGl0ZWQgYWZ0ZXIgaW1wb3J0IOKAlCB0aGUgbW9zdCByZWNlbnQgaW1wb3J0IGlzIHN0aWxsIHRoZSBvbmUga2VwdC5gCiAgICAgICAgOiAnJzsKICAgICAgY29uc3QgbWVzc2FnZSA9CiAgICAgICAgYEZvdW5kICR7cHJldmlldy5yZW1vdmVDb3VudH0gZHVwbGljYXRlIHJlY29yZChzKSBhY3Jvc3MgJHtwcmV2aWV3Lmdyb3VwQ291bnR9IHBvc3QocykgYCArCiAgICAgICAgYChzYW1lIHB1Ymxpc2ggZGF0ZSBhbmQgY2FwdGlvbikuIFJlbW92ZSB0aGUgb2xkZXIgY29waWVzIGFuZCBrZWVwIHRoZSBtb3N0IHJlY2VudGx5IHVwbG9hZGVkIG9uZSBmb3IgZWFjaD9gICsKICAgICAgICBgXG5cblRoaXMgdXBkYXRlcyB0aGUgZGFzaGJvYXJkLCBjb21wYXJpc29ucywgYW5kIHJlcG9ydHMuIEV2ZXJ5IG9yaWdpbmFsIGltcG9ydCBzdGF5cyBpbiBVcGxvYWQgSGlzdG9yeS5gICsKICAgICAgICBlZGl0ZWROb3RlOwogICAgICBpZiAoIXdpbmRvdy5jb25maXJtKG1lc3NhZ2UpKSByZXR1cm47CiAgICAgIGNvbnN0IHJlc3VsdCA9IGF3YWl0IEFwaS5yZXNvbHZlRHVwbGljYXRlUmVjb3JkcygpOwogICAgICBUb2FzdC5zaG93KGBSZW1vdmVkICR7cmVzdWx0LnJlbW92ZWRDb3VudH0gZHVwbGljYXRlIHJlY29yZChzKS5gLCAnc3VjY2VzcycpOwogICAgICBwYWdlID0gMTsKICAgICAgYXdhaXQgbG9hZCgpOwogICAgICB3aW5kb3cuZGlzcGF0Y2hFdmVudChuZXcgQ3VzdG9tRXZlbnQoJ2xyczpkYXRhLXVwZGF0ZWQnKSk7CiAgICB9IGNhdGNoIChlcnIpIHsKICAgICAgVG9hc3Quc2hvdyhlcnIubWVzc2FnZSB8fCAnRHVwbGljYXRlIGNsZWFudXAgZmFpbGVkLicsICdlcnJvcicpOwogICAgfSBmaW5hbGx5IHsKICAgICAgaWYgKGJ0bikgYnRuLmRpc2FibGVkID0gZmFsc2U7CiAgICB9CiAgfQoKICBhc3luYyBmdW5jdGlvbiBoYW5kbGVEZWxldGUocm93LCBwbGF0Zm9ybSkgewogICAgY29uc3QgY2FwdGlvbiA9IChyb3cuY2FwdGlvbiB8fCAnKG5vIGNhcHRpb24pJykuc2xpY2UoMCwgNjApOwogICAgY29uc3QgbWVzc2FnZSA9IHBsYXRmb3JtID09PSAnYWxsJwogICAgICA/IGBEZWxldGUgdGhpcyBlbnRpcmUgcmVjb3JkIOKAlCAiJHtjYXB0aW9ufSIg4oCUIGFjcm9zcyBldmVyeSBwbGF0Zm9ybT8gSXRzIG9yaWdpbmFsIGltcG9ydCBzdGF5cyBpbiBVcGxvYWQgSGlzdG9yeSwgYnV0IGl0IHdpbGwgZGlzYXBwZWFyIGZyb20gdGhlIGRhc2hib2FyZCwgY29tcGFyaXNvbnMsIGFuZCByZXBvcnRzLmAKICAgICAgOiBgUmVtb3ZlIHRoaXMgcmVjb3JkJ3MgJHtwbGF0Zm9ybUxhYmVsKHBsYXRmb3JtKX0gZGF0YSDigJQgIiR7Y2FwdGlvbn0iPyBJZiB0aGlzIGlzIGl0cyBvbmx5IHBsYXRmb3JtLCB0aGUgd2hvbGUgcmVjb3JkIHdpbGwgYmUgcmVtb3ZlZCBmcm9tIHRoZSBkYXNoYm9hcmQuYDsKICAgIGlmICghd2luZG93LmNvbmZpcm0obWVzc2FnZSkpIHJldHVybjsKICAgIHRyeSB7CiAgICAgIGlmIChwbGF0Zm9ybSA9PT0gJ2FsbCcpIGF3YWl0IEFwaS5kZWxldGVSZWNvcmRQb3N0KHJvdy5wb3N0SWQpOwogICAgICBlbHNlIGF3YWl0IEFwaS5kZWxldGVSZWNvcmRQbGF0Zm9ybShyb3cucG9zdElkLCBwbGF0Zm9ybSk7CiAgICAgIFRvYXN0LnNob3coJ1JlY29yZCBkZWxldGVkLicsICdzdWNjZXNzJyk7CiAgICAgIGF3YWl0IGxvYWQoKTsKICAgICAgd2luZG93LmRpc3BhdGNoRXZlbnQobmV3IEN1c3RvbUV2ZW50KCdscnM6ZGF0YS11cGRhdGVkJykpOwogICAgfSBjYXRjaCAoZXJyKSB7CiAgICAgIFRvYXN0LnNob3coZXJyLm1lc3NhZ2UsICdlcnJvcicpOwogICAgfQogIH0KCiAgZnVuY3Rpb24gcmVtb3ZlRXhpc3RpbmdPdmVybGF5KCkgewogICAgY29uc3Qgb3ZlcmxheSA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdyZWNvcmRNb2RhbE92ZXJsYXknKTsKICAgIGlmIChvdmVybGF5KSBvdmVybGF5LnJlbW92ZSgpOwogIH0KCiAgZnVuY3Rpb24gY2xvc2VNb2RhbCgpIHsKICAgIHJlbW92ZUV4aXN0aW5nT3ZlcmxheSgpOwogICAgbW9kYWxTdGF0ZSA9IG51bGw7CiAgfQoKICAvLyBPbmx5IGNsZWFycyB0aGUgc3RhbGUgRE9NIG5vZGUg4oCUIE5PVCBtb2RhbFN0YXRlLiByZW5kZXJFZGl0TW9kYWwgcmVhZHMKICAvLyBtb2RhbFN0YXRlIHJpZ2h0IGFmdGVyIGNhbGxpbmcgdGhpcyB0byBidWlsZCB0aGUgZm9ybTsgaWYgdGhpcyBjYWxsZWQKICAvLyB0aGUgcmVhbCBjbG9zZU1vZGFsKCkgKGFzIGl0IHVzZWQgdG8pLCB0aGF0IHJlc2V0IG1vZGFsU3RhdGUgdG8gbnVsbCBvdXQKICAvLyBmcm9tIHVuZGVyIGl0IGJlZm9yZSB0aGUgcmVhZCwgd2hpY2ggaXMgZXhhY3RseSB3aHkgRWRpdCB3YXMgYnJva2VuLgogIGZ1bmN0aW9uIG1vZGFsU2hlbGwodGl0bGVUZXh0KSB7CiAgICByZW1vdmVFeGlzdGluZ092ZXJsYXkoKTsKICAgIGNvbnN0IG92ZXJsYXkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIG92ZXJsYXkuY2xhc3NOYW1lID0gJ21vZGFsLW92ZXJsYXknOwogICAgb3ZlcmxheS5pZCA9ICdyZWNvcmRNb2RhbE92ZXJsYXknOwogICAgb3ZlcmxheS5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIChlKSA9PiB7IGlmIChlLnRhcmdldCA9PT0gb3ZlcmxheSkgY2xvc2VNb2RhbCgpOyB9KTsKICAgIGNvbnN0IHBhbmVsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBwYW5lbC5jbGFzc05hbWUgPSAnbW9kYWwtcGFuZWwgd2lkZSc7CiAgICBwYW5lbC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2gyJywgdGl0bGVUZXh0KSk7CiAgICBvdmVybGF5LmFwcGVuZENoaWxkKHBhbmVsKTsKICAgIHJldHVybiB7IG92ZXJsYXksIHBhbmVsIH07CiAgfQoKICBmdW5jdGlvbiByZWNvcmRTdWJ0aXRsZShyKSB7CiAgICByZXR1cm4gYFNoZWV0ICIke3Iuc2hlZXROYW1lfSIsIHJvdyAke3Iucm93TnVtYmVyfSR7ci5wb3N0SWQgPyBgIOKAlCBsaW5rZWQgdG8gZGFzaGJvYXJkIHBvc3QgIyR7ci5wb3N0SWR9YCA6ICcg4oCUIG5vdCBwYXJ0IG9mIHRoZSBkYXNoYm9hcmQgKGUuZy4gbmVlZHMgYSB2YWxpZCBkYXRlKSd9YDsKICB9CgogIC8vIC0tLS0tLS0tLS0gVmlldyBwb3B1cDogcmVhZC1vbmx5LCBldmVyeSBmaWVsZCwgZ3JvdXBlZCBpbnRvIHNlY3Rpb25zIC0tLS0tLS0tLS0KICBhc3luYyBmdW5jdGlvbiBvcGVuVmlldyhpZCkgewogICAgY29uc3QgcmVjb3JkID0gYXdhaXQgQXBpLmdldFJlY29yZChpZCk7CiAgICBjb25zdCB7IG92ZXJsYXksIHBhbmVsIH0gPSBtb2RhbFNoZWxsKCdSZWNvcmQgZGV0YWlscycpOwogICAgcGFuZWwuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCByZWNvcmRTdWJ0aXRsZShyZWNvcmQpLCAnbW9kYWwtc3ViJykpOwoKICAgIGNvbnN0IGdyb3VwcyA9IGdyb3VwRmllbGRSb3dzKGNvbHVtbkxhYmVsc0ZvcihyZWNvcmQpLCByZWNvcmQudmFsdWVzKTsKICAgIGdyb3Vwcy5mb3JFYWNoKChnKSA9PiB7CiAgICAgIGNvbnN0IHNlY3Rpb24gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgc2VjdGlvbi5jbGFzc05hbWUgPSAncmVjb3JkLXNlY3Rpb24nOwogICAgICBzZWN0aW9uLmFwcGVuZENoaWxkKHRleHRFbCgnaDQnLCBnLmdyb3VwKSk7CiAgICAgIGNvbnN0IGdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgZ3JpZC5jbGFzc05hbWUgPSAnZm9ybS1ncmlkJzsKICAgICAgZy5maWVsZHMuZm9yRWFjaCgoZikgPT4gewogICAgICAgIGNvbnN0IGZpZWxkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgICAgZmllbGQuY2xhc3NOYW1lID0gJ3ZpZXctZmllbGQnOwogICAgICAgIGZpZWxkLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgZi5sYWJlbCwgJ3ZpZXctbGFiZWwnKSk7CiAgICAgICAgY29uc3QgdmFsID0gZi52YWx1ZSA9PT0gdW5kZWZpbmVkIHx8IGYudmFsdWUgPT09IG51bGwgfHwgZi52YWx1ZSA9PT0gJycgPyAn4oCUJyA6IFN0cmluZyhmLnZhbHVlKTsKICAgICAgICBmaWVsZC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIHZhbCwgJ3ZpZXctdmFsdWUnKSk7CiAgICAgICAgZ3JpZC5hcHBlbmRDaGlsZChmaWVsZCk7CiAgICAgIH0pOwogICAgICBzZWN0aW9uLmFwcGVuZENoaWxkKGdyaWQpOwogICAgICBwYW5lbC5hcHBlbmRDaGlsZChzZWN0aW9uKTsKICAgIH0pOwoKICAgIGNvbnN0IGFjdGlvbnMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGFjdGlvbnMuY2xhc3NOYW1lID0gJ21vZGFsLWFjdGlvbnMnOwogICAgY29uc3QgYnRuUm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBidG5Sb3cuY2xhc3NOYW1lID0gJ2J0bi1yb3cnOwogICAgY29uc3QgY2xvc2VCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIGNsb3NlQnRuLmNsYXNzTmFtZSA9ICdidG4nOwogICAgY2xvc2VCdG4udGV4dENvbnRlbnQgPSAnQ2xvc2UnOwogICAgY2xvc2VCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBjbG9zZU1vZGFsKTsKICAgIGNvbnN0IGVkaXRCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIGVkaXRCdG4uY2xhc3NOYW1lID0gJ2J0biBwcmltYXJ5JzsKICAgIGVkaXRCdG4udGV4dENvbnRlbnQgPSAnRWRpdCB0aGlzIHJlY29yZCc7CiAgICBlZGl0QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gb3BlbkVkaXRvcihyZWNvcmQuaWQpKTsKICAgIGJ0blJvdy5hcHBlbmQoY2xvc2VCdG4sIGVkaXRCdG4pOwogICAgYWN0aW9ucy5hcHBlbmRDaGlsZChidG5Sb3cpOwogICAgcGFuZWwuYXBwZW5kQ2hpbGQoYWN0aW9ucyk7CgogICAgZG9jdW1lbnQuYm9keS5hcHBlbmRDaGlsZChvdmVybGF5KTsKICB9CgogIC8vIC0tLS0tLS0tLS0gRWRpdCBwb3B1cDogZXZlcnkgZmllbGQsIGdyb3VwZWQgaW50byBzZWN0aW9ucywgYWxsIGVkaXRhYmxlIC0tLS0tLS0tLS0KICBhc3luYyBmdW5jdGlvbiBvcGVuRWRpdG9yKGlkKSB7CiAgICBjb25zdCByZWNvcmQgPSBhd2FpdCBBcGkuZ2V0UmVjb3JkKGlkKTsKICAgIG1vZGFsU3RhdGUgPSB7IHJlY29yZCwgdmFsdWVzOiBbLi4ucmVjb3JkLnZhbHVlc10gfTsKICAgIHJlbmRlckVkaXRNb2RhbCgpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyRWRpdE1vZGFsKCkgewogICAgY29uc3QgciA9IG1vZGFsU3RhdGUucmVjb3JkOwogICAgY29uc3QgeyBvdmVybGF5LCBwYW5lbCB9ID0gbW9kYWxTaGVsbCgnRWRpdCByZWNvcmQnKTsKICAgIHBhbmVsLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgcmVjb3JkU3VidGl0bGUociksICdtb2RhbC1zdWInKSk7CgogICAgY29uc3QgZ3JvdXBzID0gZ3JvdXBGaWVsZFJvd3MoY29sdW1uTGFiZWxzRm9yKHIpLCBtb2RhbFN0YXRlLnZhbHVlcyk7CiAgICBncm91cHMuZm9yRWFjaCgoZykgPT4gewogICAgICBjb25zdCBzZWN0aW9uID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIHNlY3Rpb24uY2xhc3NOYW1lID0gJ3JlY29yZC1zZWN0aW9uJzsKICAgICAgc2VjdGlvbi5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2g0JywgZy5ncm91cCkpOwogICAgICBjb25zdCBncmlkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIGdyaWQuY2xhc3NOYW1lID0gJ2Zvcm0tZ3JpZCc7CiAgICAgIGcuZmllbGRzLmZvckVhY2goKGYpID0+IHsKICAgICAgICBjb25zdCBmaWVsZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICAgIGZpZWxkLmNsYXNzTmFtZSA9ICdmb3JtLWZpZWxkJzsKICAgICAgICBmaWVsZC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2xhYmVsJywgZi5sYWJlbCkpOwogICAgICAgIGNvbnN0IHN0clZhbCA9IGYudmFsdWUgPT09IHVuZGVmaW5lZCB8fCBmLnZhbHVlID09PSBudWxsID8gJycgOiBTdHJpbmcoZi52YWx1ZSk7CiAgICAgICAgY29uc3QgaXNMb25nID0gc3RyVmFsLmxlbmd0aCA+IDgwOwogICAgICAgIGNvbnN0IGlucHV0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudChpc0xvbmcgPyAndGV4dGFyZWEnIDogJ2lucHV0Jyk7CiAgICAgICAgaWYgKCFpc0xvbmcpIGlucHV0LnR5cGUgPSAndGV4dCc7CiAgICAgICAgZWxzZSBmaWVsZC5zdHlsZS5ncmlkQ29sdW1uID0gJzEgLyAtMSc7CiAgICAgICAgaW5wdXQudmFsdWUgPSBzdHJWYWw7CiAgICAgICAgaW5wdXQuYWRkRXZlbnRMaXN0ZW5lcignaW5wdXQnLCAoKSA9PiB7IG1vZGFsU3RhdGUudmFsdWVzW2YuaWR4XSA9IGlucHV0LnZhbHVlOyB9KTsKICAgICAgICBmaWVsZC5hcHBlbmRDaGlsZChpbnB1dCk7CiAgICAgICAgZ3JpZC5hcHBlbmRDaGlsZChmaWVsZCk7CiAgICAgIH0pOwogICAgICBzZWN0aW9uLmFwcGVuZENoaWxkKGdyaWQpOwogICAgICBwYW5lbC5hcHBlbmRDaGlsZChzZWN0aW9uKTsKICAgIH0pOwoKICAgIGNvbnN0IGFjdGlvbnMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGFjdGlvbnMuY2xhc3NOYW1lID0gJ21vZGFsLWFjdGlvbnMnOwogICAgY29uc3QgZXJyb3JNc2cgPSB0ZXh0RWwoJ3NwYW4nLCAnJywgJ211dGVkJyk7CiAgICBlcnJvck1zZy5pZCA9ICdtb2RhbEVycm9yTXNnJzsKICAgIGNvbnN0IGJ0blJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYnRuUm93LmNsYXNzTmFtZSA9ICdidG4tcm93JzsKICAgIGNvbnN0IGNhbmNlbEJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgY2FuY2VsQnRuLmNsYXNzTmFtZSA9ICdidG4nOwogICAgY2FuY2VsQnRuLnRleHRDb250ZW50ID0gJ0NhbmNlbCc7CiAgICBjYW5jZWxCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBjbG9zZU1vZGFsKTsKICAgIGNvbnN0IHNhdmVCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIHNhdmVCdG4uY2xhc3NOYW1lID0gJ2J0biBwcmltYXJ5JzsKICAgIHNhdmVCdG4udGV4dENvbnRlbnQgPSAnU2F2ZSBjaGFuZ2VzJzsKICAgIHNhdmVCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBzYXZlRWRpdChzYXZlQnRuKSk7CiAgICBidG5Sb3cuYXBwZW5kKGNhbmNlbEJ0biwgc2F2ZUJ0bik7CiAgICBhY3Rpb25zLmFwcGVuZChlcnJvck1zZywgYnRuUm93KTsKICAgIHBhbmVsLmFwcGVuZENoaWxkKGFjdGlvbnMpOwoKICAgIGRvY3VtZW50LmJvZHkuYXBwZW5kQ2hpbGQob3ZlcmxheSk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiBzYXZlRWRpdChidG4pIHsKICAgIGNvbnN0IGVycm9yRWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbW9kYWxFcnJvck1zZycpOwogICAgZXJyb3JFbC50ZXh0Q29udGVudCA9ICcnOwogICAgYnRuLmRpc2FibGVkID0gdHJ1ZTsKICAgIGJ0bi50ZXh0Q29udGVudCA9ICdTYXZpbmfigKYnOwogICAgdHJ5IHsKICAgICAgYXdhaXQgQXBpLnVwZGF0ZVJlY29yZChtb2RhbFN0YXRlLnJlY29yZC5pZCwgbW9kYWxTdGF0ZS52YWx1ZXMpOwogICAgICBUb2FzdC5zaG93KCdSZWNvcmQgdXBkYXRlZC4nLCAnc3VjY2VzcycpOwogICAgICBjbG9zZU1vZGFsKCk7CiAgICAgIGF3YWl0IGxvYWQoKTsKICAgICAgd2luZG93LmRpc3BhdGNoRXZlbnQobmV3IEN1c3RvbUV2ZW50KCdscnM6ZGF0YS11cGRhdGVkJykpOwogICAgfSBjYXRjaCAoZXJyKSB7CiAgICAgIGVycm9yRWwudGV4dENvbnRlbnQgPSBlcnIubWVzc2FnZTsKICAgICAgZXJyb3JFbC5zdHlsZS5jb2xvciA9ICd2YXIoLS1zdGF0dXMtY3JpdGljYWwpJzsKICAgICAgYnRuLmRpc2FibGVkID0gZmFsc2U7CiAgICAgIGJ0bi50ZXh0Q29udGVudCA9ICdTYXZlIGNoYW5nZXMnOwogICAgfQogIH0KCiAgYXN5bmMgZnVuY3Rpb24gcmVuZGVyKCkgewogICAgcm9vdCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd2aWV3LXJlY29yZHMnKTsKICAgIHBhZ2UgPSAxOwogICAgc2hlbGwoKTsKICAgIGF3YWl0IGxvYWQoKTsKICB9CgogIHJldHVybiB7IHJlbmRlciwgcmVsb2FkOiBsb2FkLCBvcGVuVmlldyB9Owp9KSgpOwoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIENvbXBhcmlzb25zIHRhYjogd2Vlay12cy13ZWVrLCBjdXN0b20gcmFuZ2UsIG1vbnRobHksCiAgIHF1YXJ0ZXJseSwgWVRELgogICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KY29uc3QgQ29tcGFyaXNvbiA9ICgoKSA9PiB7CiAgY29uc3QgTU9ERVMgPSBbCiAgICB7IGtleTogJ3BsYXRmb3JtcycsIGxhYmVsOiAnQWxsIFBsYXRmb3JtcycgfSwKICAgIHsga2V5OiAnd2VlaycsIGxhYmVsOiAnV2VlayB2cyBXZWVrJyB9LAogICAgeyBrZXk6ICdjdXN0b20nLCBsYWJlbDogJ0N1c3RvbSBSYW5nZScgfSwKICAgIHsga2V5OiAnbW9udGgnLCBsYWJlbDogJ01vbnRobHknIH0sCiAgICB7IGtleTogJ3F1YXJ0ZXInLCBsYWJlbDogJ1F1YXJ0ZXJseScgfSwKICAgIHsga2V5OiAneXRkJywgbGFiZWw6ICdZZWFyIHRvIERhdGUnIH0sCiAgXTsKICBjb25zdCBNRVRSSUNfUk9XUyA9IFsKICAgIHsga2V5OiAndmlld3MnLCBsYWJlbDogJ1ZpZXdzJyB9LAogICAgeyBrZXk6ICdyZWFjaCcsIGxhYmVsOiAnUmVhY2gnIH0sCiAgICB7IGtleTogJ2ltcHJlc3Npb25zJywgbGFiZWw6ICdJbXByZXNzaW9ucycgfSwKICAgIHsga2V5OiAnZW5nYWdlbWVudCcsIGxhYmVsOiAnRW5nYWdlbWVudCcgfSwKICAgIHsga2V5OiAnY2xpY2tzJywgbGFiZWw6ICdDbGlja3MnIH0sCiAgICB7IGtleTogJ2ZvbGxvd2Vyc19nYWluZWQnLCBsYWJlbDogJ0ZvbGxvd2VycyBHYWluZWQnIH0sCiAgICB7IGtleTogJ3dhdGNoX3RpbWVfc2Vjb25kcycsIGxhYmVsOiAnV2F0Y2ggVGltZScgfSwKICAgIHsga2V5OiAnc2hhcmVzJywgbGFiZWw6ICdTaGFyZXMnIH0sCiAgICB7IGtleTogJ2NvbW1lbnRzJywgbGFiZWw6ICdDb21tZW50cycgfSwKICAgIHsga2V5OiAnc2F2ZXMnLCBsYWJlbDogJ1NhdmVzJyB9LAogIF07CgogIGxldCBtb2RlID0gJ3BsYXRmb3Jtcyc7CiAgbGV0IHJvb3Q7CiAgbGV0IHBsYXRmb3JtQ2hhcnRNZXRyaWMgPSAnZW5nYWdlbWVudCc7CiAgbGV0IGNhcmRTb3J0TW9kZSA9ICdvdmVyYWxsJzsKICBsZXQgY2FyZFBsYXRmb3JtRmlsdGVyID0gJ2FsbCc7CiAgLy8gQ29tcGFyaXNvbnMga2VlcHMgaXRzIG93biBQbGF0Zm9ybSBmaWx0ZXIsIGluZGVwZW5kZW50IG9mIHRoZSBEYXNoYm9hcmQncyDigJQgdGhlCiAgLy8gc2hhcmVkIGZpbHRlciBiYXIgaXMgaGlkZGVuIG9uIHRoaXMgdGFiIChzZWUgc3dpdGNoVGFiKSwgYnV0IGl0cyBzdGF0ZSBwZXJzaXN0cwogIC8vIGluIG1lbW9yeSwgc28gd2l0aG91dCB0aGlzIGV2ZXJ5IG1vZGUgaGVyZSB3b3VsZCBzaWxlbnRseSBrZWVwIHdoYXRldmVyIHBsYXRmb3JtCiAgLy8gd2FzIGxhc3QgcGlja2VkIG9uIERhc2hib2FyZCB3aXRoIG5vIHZpc2libGUgY29udHJvbCB0byBzZWUgb3IgY2hhbmdlIGl0LgogIGxldCBjb21wYXJpc29uUGxhdGZvcm0gPSAnYWxsJzsKCiAgZnVuY3Rpb24gbW9uZGF5T2YoZGF0ZVN0cikgewogICAgY29uc3QgZCA9IG5ldyBEYXRlKGRhdGVTdHIpOwogICAgY29uc3QgZGF5ID0gZC5nZXREYXkoKTsKICAgIGNvbnN0IGRpZmYgPSBkYXkgPT09IDAgPyA2IDogZGF5IC0gMTsKICAgIGQuc2V0RGF0ZShkLmdldERhdGUoKSAtIGRpZmYpOwogICAgcmV0dXJuIGQudG9JU09TdHJpbmcoKS5zbGljZSgwLCAxMCk7CiAgfQogIGZ1bmN0aW9uIGFkZERheXMoZGF0ZVN0ciwgbikgewogICAgY29uc3QgZCA9IG5ldyBEYXRlKGRhdGVTdHIpOwogICAgZC5zZXREYXRlKGQuZ2V0RGF0ZSgpICsgbik7CiAgICByZXR1cm4gZC50b0lTT1N0cmluZygpLnNsaWNlKDAsIDEwKTsKICB9CgogIGZ1bmN0aW9uIHNoZWxsKCkgewogICAgcm9vdC5pbm5lckhUTUwgPSAnJzsKCiAgICBjb25zdCB0YWJzID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0YWJzLmNsYXNzTmFtZSA9ICdtb2RlLXRhYnMnOwogICAgTU9ERVMuZm9yRWFjaCgobSkgPT4gewogICAgICBjb25zdCBidG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgICAgYnRuLnRleHRDb250ZW50ID0gbS5sYWJlbDsKICAgICAgYnRuLnR5cGUgPSAnYnV0dG9uJzsKICAgICAgaWYgKG0ua2V5ID09PSBtb2RlKSBidG4uY2xhc3NMaXN0LmFkZCgnaXMtYWN0aXZlJyk7CiAgICAgIGJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsKICAgICAgICBtb2RlID0gbS5rZXk7CiAgICAgICAgc2hlbGwoKTsKICAgICAgfSk7CiAgICAgIHRhYnMuYXBwZW5kQ2hpbGQoYnRuKTsKICAgIH0pOwogICAgcm9vdC5hcHBlbmRDaGlsZCh0YWJzKTsKCiAgICBjb25zdCBjb250cm9scyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY29udHJvbHMuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgY29udHJvbHMuaWQgPSAnY29tcGFyaXNvbkNvbnRyb2xzJzsKICAgIHJvb3QuYXBwZW5kQ2hpbGQoY29udHJvbHMpOwoKICAgIGNvbnN0IHJlc3VsdHMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHJlc3VsdHMuaWQgPSAnY29tcGFyaXNvblJlc3VsdHMnOwogICAgcm9vdC5hcHBlbmRDaGlsZChyZXN1bHRzKTsKCiAgICByZW5kZXJDb250cm9scygpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyQ29udHJvbHMoKSB7CiAgICBjb25zdCBjb250cm9scyA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdjb21wYXJpc29uQ29udHJvbHMnKTsKICAgIGNvbnRyb2xzLmlubmVySFRNTCA9ICcnOwogICAgY29uc3Qgcm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICByb3cuY2xhc3NOYW1lID0gJ2J0bi1yb3cnOwogICAgcm93LnN0eWxlLmFsaWduSXRlbXMgPSAnZW5kJzsKCiAgICBjb25zdCB0b2RheSA9IG5ldyBEYXRlKCkudG9JU09TdHJpbmcoKS5zbGljZSgwLCAxMCk7CiAgICBjb25zdCB0aGlzWWVhciA9IG5ldyBEYXRlKCkuZ2V0RnVsbFllYXIoKTsKCiAgICBpZiAobW9kZSA9PT0gJ3BsYXRmb3JtcycpIHsKICAgICAgY29uc3QgZkZyb20gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOyBmRnJvbS50eXBlID0gJ2RhdGUnOyBmRnJvbS5pZCA9ICdwbGF0Zm9ybVJlcG9ydEZyb20nOwogICAgICBjb25zdCBmVG8gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOyBmVG8udHlwZSA9ICdkYXRlJzsgZlRvLmlkID0gJ3BsYXRmb3JtUmVwb3J0VG8nOwogICAgICBjb25zdCBhcHBseUJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgICBhcHBseUJ0bi5jbGFzc05hbWUgPSAnYnRuIHByaW1hcnknOwogICAgICBhcHBseUJ0bi50eXBlID0gJ2J1dHRvbic7CiAgICAgIGFwcGx5QnRuLnRleHRDb250ZW50ID0gJ0FwcGx5IFJhbmdlJzsKICAgICAgYXBwbHlCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBsb2FkUGxhdGZvcm1SZXBvcnQoeyBkYXRlRnJvbTogZkZyb20udmFsdWUsIGRhdGVUbzogZlRvLnZhbHVlIH0pKTsKICAgICAgY29uc3QgY2xlYXJCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgICAgY2xlYXJCdG4uY2xhc3NOYW1lID0gJ2J0bic7CiAgICAgIGNsZWFyQnRuLnR5cGUgPSAnYnV0dG9uJzsKICAgICAgY2xlYXJCdG4udGV4dENvbnRlbnQgPSAnQWxsIFRpbWUnOwogICAgICBjbGVhckJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsgZkZyb20udmFsdWUgPSAnJzsgZlRvLnZhbHVlID0gJyc7IGxvYWRQbGF0Zm9ybVJlcG9ydCh7fSk7IH0pOwogICAgICByb3cuYXBwZW5kKAogICAgICAgIGxhYmVsZWQoJ0Zyb20gKG9wdGlvbmFsKScsIGZGcm9tKSwKICAgICAgICBsYWJlbGVkKCdUbyAob3B0aW9uYWwpJywgZlRvKSwKICAgICAgICBhcHBseUJ0biwKICAgICAgICBjbGVhckJ0bgogICAgICApOwogICAgICBjb250cm9scy5hcHBlbmRDaGlsZChyb3cpOwogICAgICBsb2FkUGxhdGZvcm1SZXBvcnQoe30pOwogICAgICByZXR1cm47CiAgICB9IGVsc2UgaWYgKG1vZGUgPT09ICd3ZWVrJykgewogICAgICAvLyBXZWVrIEEgaXMgYWx3YXlzIHRoZSBQcmV2aW91cyBwZXJpb2QsIFdlZWsgQiBpcyBhbHdheXMgdGhlIEN1cnJlbnQvbW9zdCByZWNlbnQKICAgICAgLy8gcGVyaW9kIOKAlCBydW5Db21wYXJlKCkncyBmaXJzdCBhcmd1bWVudCBpcyB0aGUgImN1cnJlbnQiIHNsb3QgZXZlcnkgb3RoZXIgbW9kZQogICAgICAvLyBpbiB0aGlzIHRhYiBmZWVkcyBpdCAocGVyY2VudENoYW5nZSA9IChjdXJyZW50IC0gcHJldmlvdXMpIC8gcHJldmlvdXMpLCBzbyBCIGdvZXMKICAgICAgLy8gaW4gZmlyc3QgYW5kIEEgc2Vjb25kLCByZWdhcmRsZXNzIG9mIHdoaWNoIGNhbGVuZGFyIHdlZWsgaXMgZWFybGllciBvciBsYXRlci4KICAgICAgY29uc3Qgd0EgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOyB3QS50eXBlID0gJ2RhdGUnOyB3QS52YWx1ZSA9IG1vbmRheU9mKGFkZERheXModG9kYXksIC03KSk7CiAgICAgIGNvbnN0IHdCID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsgd0IudHlwZSA9ICdkYXRlJzsgd0IudmFsdWUgPSBtb25kYXlPZih0b2RheSk7CiAgICAgIHJvdy5hcHBlbmQobGFiZWxlZCgnV2VlayBBIChQcmV2aW91cyknLCB3QSksIGxhYmVsZWQoJ1dlZWsgQiAoQ3VycmVudCknLCB3QiksIHBsYXRmb3JtRmlsdGVyRmllbGQoKSwgcnVuQnRuKCgpID0+IHsKICAgICAgICBjb25zdCByYW5nZUEgPSB7IGZyb206IG1vbmRheU9mKHdBLnZhbHVlKSwgdG86IGFkZERheXMobW9uZGF5T2Yod0EudmFsdWUpLCA2KSB9OwogICAgICAgIGNvbnN0IHJhbmdlQiA9IHsgZnJvbTogbW9uZGF5T2Yod0IudmFsdWUpLCB0bzogYWRkRGF5cyhtb25kYXlPZih3Qi52YWx1ZSksIDYpIH07CiAgICAgICAgcnVuQ29tcGFyZShyYW5nZUIsIHJhbmdlQSwgYFdlZWsgb2YgJHtGb3JtYXQuZGF0ZShyYW5nZUIuZnJvbSl9YCwgYFdlZWsgb2YgJHtGb3JtYXQuZGF0ZShyYW5nZUEuZnJvbSl9YCk7CiAgICAgIH0pKTsKICAgIH0gZWxzZSBpZiAobW9kZSA9PT0gJ2N1c3RvbScpIHsKICAgICAgY29uc3QgZkEgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOyBmQS50eXBlID0gJ2RhdGUnOyBmQS52YWx1ZSA9IGFkZERheXModG9kYXksIC0xMyk7CiAgICAgIGNvbnN0IHRBID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsgdEEudHlwZSA9ICdkYXRlJzsgdEEudmFsdWUgPSB0b2RheTsKICAgICAgY29uc3QgZkIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOyBmQi50eXBlID0gJ2RhdGUnOyBmQi52YWx1ZSA9IGFkZERheXModG9kYXksIC0yNyk7CiAgICAgIGNvbnN0IHRCID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsgdEIudHlwZSA9ICdkYXRlJzsgdEIudmFsdWUgPSBhZGREYXlzKHRvZGF5LCAtMTQpOwogICAgICByb3cuYXBwZW5kKAogICAgICAgIGxhYmVsZWQoJ1JhbmdlIEEgZnJvbScsIGZBKSwgbGFiZWxlZCgndG8nLCB0QSksCiAgICAgICAgbGFiZWxlZCgnUmFuZ2UgQiBmcm9tJywgZkIpLCBsYWJlbGVkKCd0bycsIHRCKSwKICAgICAgICBwbGF0Zm9ybUZpbHRlckZpZWxkKCksCiAgICAgICAgcnVuQnRuKCgpID0+IHJ1bkNvbXBhcmUoeyBmcm9tOiBmQS52YWx1ZSwgdG86IHRBLnZhbHVlIH0sIHsgZnJvbTogZkIudmFsdWUsIHRvOiB0Qi52YWx1ZSB9LCAnUmFuZ2UgQScsICdSYW5nZSBCJykpCiAgICAgICk7CiAgICB9IGVsc2UgaWYgKG1vZGUgPT09ICdtb250aCcpIHsKICAgICAgY29uc3QgeSA9IHllYXJTZWxlY3QodGhpc1llYXIpOyBjb25zdCBtID0gbW9udGhTZWxlY3QobmV3IERhdGUoKS5nZXRNb250aCgpICsgMSk7CiAgICAgIGNvbnN0IHRvZ2dsZSA9IHBlcmlvZFRvZ2dsZSgpOwogICAgICByb3cuYXBwZW5kKGxhYmVsZWQoJ1llYXInLCB5KSwgbGFiZWxlZCgnTW9udGgnLCBtKSwgdG9nZ2xlLmVsLCBwbGF0Zm9ybUZpbHRlckZpZWxkKCksIHJ1bkJ0bihhc3luYyAoKSA9PiB7CiAgICAgICAgY29uc3QgcmVwb3J0ID0gYXdhaXQgQXBpLm1vbnRobHkoeyB5ZWFyOiB5LnZhbHVlLCBtb250aDogbS52YWx1ZSwgLi4uU3RhdGUuZ2V0RmlsdGVycygpLCBwbGF0Zm9ybTogY29tcGFyaXNvblBsYXRmb3JtIH0pOwogICAgICAgIHJlbmRlclBlcmlvZFJlcG9ydChyZXBvcnQsIHRvZ2dsZS5nZXQoKSk7CiAgICAgIH0pKTsKICAgIH0gZWxzZSBpZiAobW9kZSA9PT0gJ3F1YXJ0ZXInKSB7CiAgICAgIGNvbnN0IHkgPSB5ZWFyU2VsZWN0KHRoaXNZZWFyKTsgY29uc3QgcSA9IHF1YXJ0ZXJTZWxlY3QoKTsKICAgICAgY29uc3QgdG9nZ2xlID0gcGVyaW9kVG9nZ2xlKCk7CiAgICAgIHJvdy5hcHBlbmQobGFiZWxlZCgnWWVhcicsIHkpLCBsYWJlbGVkKCdRdWFydGVyJywgcSksIHRvZ2dsZS5lbCwgcGxhdGZvcm1GaWx0ZXJGaWVsZCgpLCBydW5CdG4oYXN5bmMgKCkgPT4gewogICAgICAgIGNvbnN0IHJlcG9ydCA9IGF3YWl0IEFwaS5xdWFydGVybHkoeyB5ZWFyOiB5LnZhbHVlLCBxdWFydGVyOiBxLnZhbHVlLCAuLi5TdGF0ZS5nZXRGaWx0ZXJzKCksIHBsYXRmb3JtOiBjb21wYXJpc29uUGxhdGZvcm0gfSk7CiAgICAgICAgcmVuZGVyUGVyaW9kUmVwb3J0KHJlcG9ydCwgdG9nZ2xlLmdldCgpKTsKICAgICAgfSkpOwogICAgfSBlbHNlIGlmIChtb2RlID09PSAneXRkJykgewogICAgICBjb25zdCB5ID0geWVhclNlbGVjdCh0aGlzWWVhcik7CiAgICAgIHJvdy5hcHBlbmQobGFiZWxlZCgnWWVhcicsIHkpLCBwbGF0Zm9ybUZpbHRlckZpZWxkKCksIHJ1bkJ0bihhc3luYyAoKSA9PiB7CiAgICAgICAgY29uc3QgcmVwb3J0ID0gYXdhaXQgQXBpLnl0ZCh7IHllYXI6IHkudmFsdWUsIC4uLlN0YXRlLmdldEZpbHRlcnMoKSwgcGxhdGZvcm06IGNvbXBhcmlzb25QbGF0Zm9ybSB9KTsKICAgICAgICByZW5kZXJQZXJpb2RSZXBvcnQocmVwb3J0LCAndnNMYXN0WWVhcicpOwogICAgICB9KSk7CiAgICB9CgogICAgY29udHJvbHMuYXBwZW5kQ2hpbGQocm93KTsKICB9CgogIC8qKiBUaGUgQ29tcGFyaXNvbnMgdGFiJ3Mgb3duIFBsYXRmb3JtIGZpbHRlciDigJQgcmV1c2VzIHRoZSBzYW1lIHBsYXRmb3JtIGxpc3QgdGhlCiAgICAgIERhc2hib2FyZCdzIGZpbHRlciBiYXIgc2hvd3MsIGJ1dCB3cml0ZXMgdG8gY29tcGFyaXNvblBsYXRmb3JtLCBub3QgU3RhdGUsIHNvCiAgICAgIHRoZSB0d28gc3RheSBmdWxseSBpbmRlcGVuZGVudC4gKi8KICBmdW5jdGlvbiBwbGF0Zm9ybUZpbHRlckZpZWxkKCkgewogICAgY29uc3Qgc2VsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc2VsZWN0Jyk7CiAgICBjb25zdCBhbGxPcHQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsKICAgIGFsbE9wdC52YWx1ZSA9ICdhbGwnOwogICAgYWxsT3B0LnRleHRDb250ZW50ID0gJ0FsbCBwbGF0Zm9ybXMnOwogICAgc2VsLmFwcGVuZENoaWxkKGFsbE9wdCk7CiAgICBjb25zdCBvcHRpb25zID0gKHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZSB8fCB7IHBsYXRmb3JtczogW10gfSkucGxhdGZvcm1zIHx8IFtdOwogICAgb3B0aW9ucy5mb3JFYWNoKChwKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOwogICAgICBvcHQudmFsdWUgPSBwLmlkOwogICAgICBvcHQudGV4dENvbnRlbnQgPSBwLmxhYmVsOwogICAgICBzZWwuYXBwZW5kQ2hpbGQob3B0KTsKICAgIH0pOwogICAgc2VsLnZhbHVlID0gY29tcGFyaXNvblBsYXRmb3JtOwogICAgc2VsLmFkZEV2ZW50TGlzdGVuZXIoJ2NoYW5nZScsICgpID0+IHsgY29tcGFyaXNvblBsYXRmb3JtID0gc2VsLnZhbHVlOyB9KTsKICAgIHJldHVybiBsYWJlbGVkKCdQbGF0Zm9ybScsIHNlbCk7CiAgfQoKICBmdW5jdGlvbiBsYWJlbGVkKGxhYmVsLCBlbCkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgd3JhcC5jbGFzc05hbWUgPSAnZmllbGQtaW5saW5lJzsKICAgIHdyYXAuYXBwZW5kKHRleHRFbCgnbGFiZWwnLCBsYWJlbCksIGVsKTsKICAgIHJldHVybiB3cmFwOwogIH0KICBmdW5jdGlvbiBydW5CdG4ob25DbGljaykgewogICAgY29uc3QgYnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBidG4uY2xhc3NOYW1lID0gJ2J0biBwcmltYXJ5JzsKICAgIGJ0bi50ZXh0Q29udGVudCA9ICdDb21wYXJlJzsKICAgIGJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IG9uQ2xpY2soKSk7CiAgICByZXR1cm4gYnRuOwogIH0KICBmdW5jdGlvbiB5ZWFyU2VsZWN0KGRlZmF1bHRZZWFyKSB7CiAgICBjb25zdCBzZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzZWxlY3QnKTsKICAgIGZvciAobGV0IHkgPSBkZWZhdWx0WWVhciAtIDM7IHkgPD0gZGVmYXVsdFllYXIgKyAxOyB5ICs9IDEpIHsKICAgICAgY29uc3Qgb3B0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7IG9wdC52YWx1ZSA9IHk7IG9wdC50ZXh0Q29udGVudCA9IHk7CiAgICAgIGlmICh5ID09PSBkZWZhdWx0WWVhcikgb3B0LnNlbGVjdGVkID0gdHJ1ZTsKICAgICAgc2VsLmFwcGVuZENoaWxkKG9wdCk7CiAgICB9CiAgICByZXR1cm4gc2VsOwogIH0KICBmdW5jdGlvbiBtb250aFNlbGVjdChkZWZhdWx0TW9udGgpIHsKICAgIGNvbnN0IG5hbWVzID0gWydKYW51YXJ5JywnRmVicnVhcnknLCdNYXJjaCcsJ0FwcmlsJywnTWF5JywnSnVuZScsJ0p1bHknLCdBdWd1c3QnLCdTZXB0ZW1iZXInLCdPY3RvYmVyJywnTm92ZW1iZXInLCdEZWNlbWJlciddOwogICAgY29uc3Qgc2VsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc2VsZWN0Jyk7CiAgICBuYW1lcy5mb3JFYWNoKChuLCBpKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOyBvcHQudmFsdWUgPSBpICsgMTsgb3B0LnRleHRDb250ZW50ID0gbjsKICAgICAgaWYgKGkgKyAxID09PSBkZWZhdWx0TW9udGgpIG9wdC5zZWxlY3RlZCA9IHRydWU7CiAgICAgIHNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgICByZXR1cm4gc2VsOwogIH0KICBmdW5jdGlvbiBxdWFydGVyU2VsZWN0KCkgewogICAgY29uc3QgY3VycmVudFEgPSBNYXRoLmZsb29yKG5ldyBEYXRlKCkuZ2V0TW9udGgoKSAvIDMpICsgMTsKICAgIGNvbnN0IHNlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NlbGVjdCcpOwogICAgWzEsIDIsIDMsIDRdLmZvckVhY2goKHEpID0+IHsKICAgICAgY29uc3Qgb3B0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7IG9wdC52YWx1ZSA9IHE7IG9wdC50ZXh0Q29udGVudCA9IGBRJHtxfWA7CiAgICAgIGlmIChxID09PSBjdXJyZW50USkgb3B0LnNlbGVjdGVkID0gdHJ1ZTsKICAgICAgc2VsLmFwcGVuZENoaWxkKG9wdCk7CiAgICB9KTsKICAgIHJldHVybiBzZWw7CiAgfQogIGZ1bmN0aW9uIHBlcmlvZFRvZ2dsZSgpIHsKICAgIGNvbnN0IHNlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NlbGVjdCcpOwogICAgW1sndnNQcmV2aW91c1BlcmlvZCcsICd2cyBQcmV2aW91cyBQZXJpb2QnXSwgWyd2c0xhc3RZZWFyJywgJ3ZzIFNhbWUgUGVyaW9kIExhc3QgWWVhciddXS5mb3JFYWNoKChbdiwgbF0pID0+IHsKICAgICAgY29uc3Qgb3B0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7IG9wdC52YWx1ZSA9IHY7IG9wdC50ZXh0Q29udGVudCA9IGw7CiAgICAgIHNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgICByZXR1cm4geyBlbDogbGFiZWxlZCgnQ29tcGFyZScsIHNlbCksIGdldDogKCkgPT4gc2VsLnZhbHVlIH07CiAgfQoKICBhc3luYyBmdW5jdGlvbiBydW5Db21wYXJlKHJhbmdlQSwgcmFuZ2VCLCBsYWJlbEEsIGxhYmVsQikgewogICAgY29uc3QgZmlsdGVycyA9IFN0YXRlLmdldEZpbHRlcnMoKTsKICAgIGNvbnN0IHJlc3VsdCA9IGF3YWl0IEFwaS5jb21wYXJlKHsKICAgICAgZnJvbUE6IHJhbmdlQS5mcm9tLCB0b0E6IHJhbmdlQS50bywgZnJvbUI6IHJhbmdlQi5mcm9tLCB0b0I6IHJhbmdlQi50bywKICAgICAgcGxhdGZvcm06IGNvbXBhcmlzb25QbGF0Zm9ybSwgY2FtcGFpZ25UeXBlOiBmaWx0ZXJzLmNhbXBhaWduVHlwZSwgY29udGVudFR5cGU6IGZpbHRlcnMuY29udGVudFR5cGUsCiAgICB9KTsKICAgIHJlbmRlckNvbXBhcmVSZXN1bHQocmVzdWx0LCBsYWJlbEEsIGxhYmVsQik7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJQZXJpb2RSZXBvcnQocmVwb3J0LCB3aGljaCkgewogICAgY29uc3QgY21wID0gcmVwb3J0W3doaWNoXTsKICAgIGNvbnN0IGxhYmVsQSA9ICdDdXJyZW50IHBlcmlvZCc7CiAgICBjb25zdCBsYWJlbEIgPSB3aGljaCA9PT0gJ3ZzTGFzdFllYXInID8gJ1NhbWUgcGVyaW9kIGxhc3QgeWVhcicgOiAnUHJldmlvdXMgcGVyaW9kJzsKICAgIHJlbmRlckNvbXBhcmVSZXN1bHQoY21wLCBsYWJlbEEsIGxhYmVsQiwgcmVwb3J0LnJhbmdlKTsKICB9CgogIGZ1bmN0aW9uIHN0YXRUaWxlKGxhYmVsLCBjdXJyZW50LCBwcmV2aW91cywgZ3Jvd3RoLCBpc0R1cmF0aW9uKSB7CiAgICBjb25zdCB0aWxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0aWxlLmNsYXNzTmFtZSA9ICdzdGF0LXRpbGUnOwogICAgY29uc3QgY3VyRGlzcGxheSA9IGlzRHVyYXRpb24gPyBGb3JtYXQuZHVyYXRpb24oY3VycmVudCkgOiBGb3JtYXQuY29tcGFjdChjdXJyZW50KTsKICAgIGNvbnN0IHByZXZEaXNwbGF5ID0gaXNEdXJhdGlvbiA/IEZvcm1hdC5kdXJhdGlvbihwcmV2aW91cykgOiBGb3JtYXQuY29tcGFjdChwcmV2aW91cyk7CiAgICB0aWxlLmFwcGVuZCgKICAgICAgdGV4dEVsKCdkaXYnLCBsYWJlbCwgJ3N0YXQtbGFiZWwnKSwKICAgICAgdGV4dEVsKCdkaXYnLCBjdXJEaXNwbGF5LCAnc3RhdC12YWx1ZScpLAogICAgICB0ZXh0RWwoJ2RpdicsIGAke0Zvcm1hdC5wY3QoZ3Jvd3RoKX0gwrcgd2FzICR7cHJldkRpc3BsYXl9YCwgYHN0YXQtZGVsdGEgJHtGb3JtYXQuZGVsdGFDbGFzcyhncm93dGgpfWApCiAgICApOwogICAgcmV0dXJuIHRpbGU7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJDb21wYXJlUmVzdWx0KHJlc3VsdCwgbGFiZWxBLCBsYWJlbEIsIGhlYWRsaW5lKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2NvbXBhcmlzb25SZXN1bHRzJyk7CiAgICB3cmFwLmlubmVySFRNTCA9ICcnOwoKICAgIGNvbnN0IHRpdGxlID0gdGV4dEVsKCdkaXYnLCBoZWFkbGluZQogICAgICA/IGAke0Zvcm1hdC5kYXRlKHJlc3VsdC5yYW5nZUEuZnJvbSl9IOKAkyAke0Zvcm1hdC5kYXRlKHJlc3VsdC5yYW5nZUEudG8pfWAKICAgICAgOiBgJHtsYWJlbEF9OiAke0Zvcm1hdC5kYXRlKHJlc3VsdC5yYW5nZUEuZnJvbSl9IOKAkyAke0Zvcm1hdC5kYXRlKHJlc3VsdC5yYW5nZUEudG8pfSAgdnMgICR7bGFiZWxCfTogJHtGb3JtYXQuZGF0ZShyZXN1bHQucmFuZ2VCLmZyb20pfSDigJMgJHtGb3JtYXQuZGF0ZShyZXN1bHQucmFuZ2VCLnRvKX1gLAogICAgICAnc2VjdGlvbi10aXRsZScpOwogICAgd3JhcC5hcHBlbmRDaGlsZCh0aXRsZSk7CgogICAgY29uc3QgZ3JpZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZ3JpZC5jbGFzc05hbWUgPSAnc3RhdC1ncmlkJzsKICAgIGdyaWQuYXBwZW5kQ2hpbGQoc3RhdFRpbGUoJ1Bvc3RzJywgcmVzdWx0LnJhbmdlQS50b3RhbHMucG9zdF9jb3VudCwgcmVzdWx0LnJhbmdlQi50b3RhbHMucG9zdF9jb3VudCwgcmVzdWx0Lmdyb3d0aC5wb3N0X2NvdW50LCBmYWxzZSkpOwogICAgTUVUUklDX1JPV1MuZm9yRWFjaCgobSkgPT4gewogICAgICBncmlkLmFwcGVuZENoaWxkKHN0YXRUaWxlKG0ubGFiZWwsIHJlc3VsdC5yYW5nZUEudG90YWxzW20ua2V5XSwgcmVzdWx0LnJhbmdlQi50b3RhbHNbbS5rZXldLCByZXN1bHQuZ3Jvd3RoW20ua2V5XSwgbS5rZXkgPT09ICd3YXRjaF90aW1lX3NlY29uZHMnKSk7CiAgICB9KTsKICAgIHdyYXAuYXBwZW5kQ2hpbGQoZ3JpZCk7CgogICAgcmVuZGVyUGxhdGZvcm1Db21wYXJpc29uQ2FyZHMod3JhcCwgcmVzdWx0LCBsYWJlbEEsIGxhYmVsQik7CiAgfQoKICAvKioKICAgKiAiQWxsIFBsYXRmb3JtcyIgcmVwb3J0IOKAlCB0aGUgaGVhZGxpbmUgQ29tcGFyaXNvbnMgdmlldy4gVW5saWtlIHRoZQogICAqIHdlZWsvY3VzdG9tL21vbnRoL3F1YXJ0ZXIveXRkIHRvb2xzIGFib3ZlLCB0aGlzIGlnbm9yZXMgdGhlIHNoYXJlZAogICAqIHBsYXRmb3JtL2NhbXBhaWduL2NvbnRlbnQtdHlwZSBmaWx0ZXIgYmFyIGVudGlyZWx5IGFuZCBuZWVkcyBubyBkYXRlCiAgICogcmFuZ2U6IGl0IGFsd2F5cyBjb3ZlcnMgZXZlcnkgcGxhdGZvcm0gd2l0aCBhbnkgZGF0YSAodXBsb2FkZWQgcG9zdHMKICAgKiBhbmQvb3IgbWFudWFsbHktZW50ZXJlZCBGb2xsb3dlcnMgRGF0YSBSZWNvcmQgaGlzdG9yeSkuCiAgICovCiAgYXN5bmMgZnVuY3Rpb24gbG9hZFBsYXRmb3JtUmVwb3J0KHBhcmFtcykgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdjb21wYXJpc29uUmVzdWx0cycpOwogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgIHdyYXAuYXBwZW5kQ2hpbGQoc2tlbGV0b25TdGF0R3JpZCgyKSk7CiAgICB3cmFwLmFwcGVuZENoaWxkKHNrZWxldG9uQ2hhcnQoKSk7CiAgICBjb25zdCBoYXNFeHBsaWNpdFJhbmdlID0gcGFyYW1zICYmIHBhcmFtcy5kYXRlRnJvbSAmJiBwYXJhbXMuZGF0ZVRvOwogICAgY29uc3QgcmVwb3J0ID0gYXdhaXQgQXBpLnBsYXRmb3JtUmVwb3J0KGhhc0V4cGxpY2l0UmFuZ2UgPyBwYXJhbXMgOiB7fSk7CiAgICByZW5kZXJQbGF0Zm9ybVJlcG9ydChyZXBvcnQpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyUGxhdGZvcm1SZXBvcnQocmVwb3J0KSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2NvbXBhcmlzb25SZXN1bHRzJyk7CiAgICB3cmFwLmlubmVySFRNTCA9ICcnOwoKICAgIGlmICghcmVwb3J0LnBsYXRmb3Jtcy5sZW5ndGgpIHsKICAgICAgd3JhcC5hcHBlbmRDaGlsZChlbXB0eVN0YXRlKHsKICAgICAgICBpY29uOiAnZ2l0LWNvbXBhcmUnLAogICAgICAgIHRpdGxlOiAnTm8gcGxhdGZvcm0gZGF0YSB5ZXQnLAogICAgICAgIG1lc3NhZ2U6ICdVcGxvYWQgcG9zdHMgb3IgYWRkIEZvbGxvd2VycyBEYXRhIFJlY29yZCBlbnRyaWVzIHRvIHNlZSBhIGNyb3NzLXBsYXRmb3JtIGNvbXBhcmlzb24gaGVyZS4nLAogICAgICB9KSk7CiAgICAgIHJldHVybjsKICAgIH0KCiAgICBjb25zdCByYW5nZUxhYmVsID0gcmVwb3J0LnJhbmdlLmlzRXhwbGljaXQKICAgICAgPyBgJHtGb3JtYXQuZGF0ZShyZXBvcnQucmFuZ2UuZnJvbSl9IOKAkyAke0Zvcm1hdC5kYXRlKHJlcG9ydC5yYW5nZS50byl9YAogICAgICA6IGBBbGwgdGltZSAoJHtGb3JtYXQuZGF0ZShyZXBvcnQucmFuZ2UuZnJvbSl9IOKAkyAke0Zvcm1hdC5kYXRlKHJlcG9ydC5yYW5nZS50byl9KWA7CiAgICB3cmFwLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgYFBsYXRmb3JtIENvbXBhcmlzb24gUmVwb3J0IOKAlCAke3JhbmdlTGFiZWx9YCwgJ3NlY3Rpb24tdGl0bGUnKSk7CgogICAgY29uc3QgYmVzdFAgPSByZXBvcnQucGxhdGZvcm1zLmZpbmQoKHApID0+IHAucGxhdGZvcm0gPT09IHJlcG9ydC5iZXN0UGxhdGZvcm0pOwogICAgY29uc3Qgd29yc3RQID0gcmVwb3J0LnBsYXRmb3Jtcy5maW5kKChwKSA9PiBwLnBsYXRmb3JtID09PSByZXBvcnQud29yc3RQbGF0Zm9ybSk7CiAgICBpZiAoYmVzdFAgfHwgd29yc3RQKSB7CiAgICAgIGNvbnN0IGhpZ2hsaWdodEdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgaGlnaGxpZ2h0R3JpZC5jbGFzc05hbWUgPSAnc3RhdC1ncmlkJzsKICAgICAgaWYgKGJlc3RQKSB7CiAgICAgICAgY29uc3QgdGlsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICAgIHRpbGUuY2xhc3NOYW1lID0gJ3N0YXQtdGlsZSc7CiAgICAgICAgdGlsZS5hcHBlbmQoCiAgICAgICAgICB0ZXh0RWwoJ2RpdicsICdCZXN0LVBlcmZvcm1pbmcgUGxhdGZvcm0nLCAnc3RhdC1sYWJlbCcpLAogICAgICAgICAgdGV4dEVsKCdkaXYnLCBiZXN0UC5sYWJlbCwgJ3N0YXQtdmFsdWUnKSwKICAgICAgICAgIHRleHRFbCgnZGl2JywgYFJlYWNoICR7Rm9ybWF0LnNtYXJ0KGJlc3RQLnRvdGFscy5yZWFjaCl9IMK3IEVuZ2FnZW1lbnQgJHtGb3JtYXQuc21hcnQoYmVzdFAudG90YWxzLmVuZ2FnZW1lbnQpfWAsICdwb3N0LW1ldGEnKQogICAgICAgICk7CiAgICAgICAgaGlnaGxpZ2h0R3JpZC5hcHBlbmRDaGlsZCh0aWxlKTsKICAgICAgfQogICAgICBpZiAod29yc3RQKSB7CiAgICAgICAgY29uc3QgdGlsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICAgIHRpbGUuY2xhc3NOYW1lID0gJ3N0YXQtdGlsZSc7CiAgICAgICAgdGlsZS5hcHBlbmQoCiAgICAgICAgICB0ZXh0RWwoJ2RpdicsICdMb3dlc3QtUGVyZm9ybWluZyBQbGF0Zm9ybScsICdzdGF0LWxhYmVsJyksCiAgICAgICAgICB0ZXh0RWwoJ2RpdicsIHdvcnN0UC5sYWJlbCwgJ3N0YXQtdmFsdWUnKSwKICAgICAgICAgIHRleHRFbCgnZGl2JywgYFJlYWNoICR7Rm9ybWF0LnNtYXJ0KHdvcnN0UC50b3RhbHMucmVhY2gpfSDCtyBFbmdhZ2VtZW50ICR7Rm9ybWF0LnNtYXJ0KHdvcnN0UC50b3RhbHMuZW5nYWdlbWVudCl9YCwgJ3Bvc3QtbWV0YScpCiAgICAgICAgKTsKICAgICAgICBoaWdobGlnaHRHcmlkLmFwcGVuZENoaWxkKHRpbGUpOwogICAgICB9CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQoaGlnaGxpZ2h0R3JpZCk7CiAgICB9CgogICAgY29uc3QgdGFibGVDYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0YWJsZUNhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgdGFibGVDYXJkLmFwcGVuZENoaWxkKHRleHRFbCgnaDMnLCAnUGxhdGZvcm0gUmFua2luZycpKTsKICAgIGNvbnN0IHRhYmxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGFibGUnKTsKICAgIHRhYmxlLmNsYXNzTmFtZSA9ICdkYXRhLXRhYmxlJzsKICAgIGNvbnN0IHRoZWFkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGhlYWQnKTsKICAgIGNvbnN0IGhlYWRSb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgWydSYW5rJywgJ1BsYXRmb3JtJywgJ1Bvc3RzJywgJ1JlYWNoJywgJ0VuZ2FnZW1lbnQnLCAnSW1wcmVzc2lvbnMnLCAnRm9sbG93ZXIgR3Jvd3RoJ10uZm9yRWFjaCgobGFiZWwsIGkpID0+IHsKICAgICAgY29uc3QgdGggPSB0ZXh0RWwoJ3RoJywgbGFiZWwpOwogICAgICBpZiAoaSA+PSAyKSB0aC5jbGFzc0xpc3QuYWRkKCdudW0nKTsKICAgICAgaGVhZFJvdy5hcHBlbmRDaGlsZCh0aCk7CiAgICB9KTsKICAgIHRoZWFkLmFwcGVuZENoaWxkKGhlYWRSb3cpOwogICAgdGFibGUuYXBwZW5kQ2hpbGQodGhlYWQpOwogICAgY29uc3QgdGJvZHkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0Ym9keScpOwogICAgcmVwb3J0LnBsYXRmb3Jtcy5mb3JFYWNoKChwKSA9PiB7CiAgICAgIGNvbnN0IHRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgdHIuYXBwZW5kQ2hpbGQodGV4dEVsKCd0ZCcsIHAub3ZlcmFsbFJhbmsgPyBgIyR7cC5vdmVyYWxsUmFua31gIDogJ+KAlCcpKTsKICAgICAgY29uc3QgcGxhdFRkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgY29uc3QgcGlsbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsgcGlsbC5jbGFzc05hbWUgPSAncGxhdGZvcm0tcGlsbCc7CiAgICAgIGNvbnN0IGRvdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsgZG90LmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1kb3QnOyBkb3Quc3R5bGUuYmFja2dyb3VuZCA9IHAuY29sb3I7CiAgICAgIHBpbGwuYXBwZW5kKGRvdCwgZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUocC5sYWJlbCkpOwogICAgICBwbGF0VGQuYXBwZW5kQ2hpbGQocGlsbCk7CiAgICAgIHRyLmFwcGVuZENoaWxkKHBsYXRUZCk7CiAgICAgIHRyLmFwcGVuZENoaWxkKHRleHRFbCgndGQnLCBGb3JtYXQubnVtYmVyKHAucG9zdENvdW50KSwgJ251bScpKTsKICAgICAgdHIuYXBwZW5kQ2hpbGQodGV4dEVsKCd0ZCcsIEZvcm1hdC5zbWFydChwLnRvdGFscy5yZWFjaCksICdudW0nKSk7CiAgICAgIHRyLmFwcGVuZENoaWxkKHRleHRFbCgndGQnLCBGb3JtYXQuc21hcnQocC50b3RhbHMuZW5nYWdlbWVudCksICdudW0nKSk7CiAgICAgIHRyLmFwcGVuZENoaWxkKHRleHRFbCgndGQnLCBGb3JtYXQuc21hcnQocC50b3RhbHMuaW1wcmVzc2lvbnMpLCAnbnVtJykpOwogICAgICBjb25zdCBmb2xsb3dlclRkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgZm9sbG93ZXJUZC5jbGFzc05hbWUgPSAnbnVtJzsKICAgICAgaWYgKHAuZm9sbG93ZXJzLmNoYW5nZSA9PT0gbnVsbCkgewogICAgICAgIGZvbGxvd2VyVGQuYXBwZW5kQ2hpbGQoZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUocC5mb2xsb3dlcnMubGF0ZXN0ICE9PSBudWxsID8gRm9ybWF0Lm51bWJlcihwLmZvbGxvd2Vycy5sYXRlc3QpIDogJ+KAlCcpKTsKICAgICAgfSBlbHNlIHsKICAgICAgICBjb25zdCBmb2xsb3dlcldyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICAgICAgZm9sbG93ZXJXcmFwLnN0eWxlLmRpc3BsYXkgPSAnaW5saW5lLWZsZXgnOwogICAgICAgIGZvbGxvd2VyV3JhcC5zdHlsZS5hbGlnbkl0ZW1zID0gJ2NlbnRlcic7CiAgICAgICAgZm9sbG93ZXJXcmFwLnN0eWxlLmdhcCA9ICc2cHgnOwogICAgICAgIGZvbGxvd2VyV3JhcC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3NwYW4nLCBgJHtwLmZvbGxvd2Vycy5jaGFuZ2UgPiAwID8gJysnIDogJyd9JHtGb3JtYXQubnVtYmVyKHAuZm9sbG93ZXJzLmNoYW5nZSl9YCwgYHN0YXQtZGVsdGEgJHtGb3JtYXQuZGVsdGFDbGFzcyhwLmZvbGxvd2Vycy5jaGFuZ2UpfWApKTsKICAgICAgICBpZiAocC5mb2xsb3dlcnMuY2hhbmdlUGN0ICE9PSBudWxsKSBmb2xsb3dlcldyYXAuYXBwZW5kQ2hpbGQodGV4dEVsKCdzcGFuJywgYCgke0Zvcm1hdC5wY3QocC5mb2xsb3dlcnMuY2hhbmdlUGN0KX0pYCwgJ3Bvc3QtbWV0YScpKTsKICAgICAgICBmb2xsb3dlclRkLmFwcGVuZENoaWxkKGZvbGxvd2VyV3JhcCk7CiAgICAgIH0KICAgICAgdHIuYXBwZW5kQ2hpbGQoZm9sbG93ZXJUZCk7CiAgICAgIHRib2R5LmFwcGVuZENoaWxkKHRyKTsKICAgIH0pOwogICAgdGFibGUuYXBwZW5kQ2hpbGQodGJvZHkpOwogICAgY29uc3QgdGFibGVTY3JvbGwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRhYmxlU2Nyb2xsLmNsYXNzTmFtZSA9ICd0YWJsZS1zY3JvbGwnOwogICAgdGFibGVTY3JvbGwuYXBwZW5kQ2hpbGQodGFibGUpOwogICAgdGFibGVDYXJkLmFwcGVuZENoaWxkKHRhYmxlU2Nyb2xsKTsKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGFibGVDYXJkKTsKCiAgICBjb25zdCBjaGFydENhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNoYXJ0Q2FyZC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICBjb25zdCBjaGFydEhlYWRlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY2hhcnRIZWFkZXIuY2xhc3NOYW1lID0gJ2NhcmQtaGVhZGVyJzsKICAgIGNoYXJ0SGVhZGVyLmFwcGVuZENoaWxkKHRleHRFbCgnaDMnLCAnTWV0cmljIENvbXBhcmlzb24nKSk7CiAgICBjb25zdCBtZXRyaWNTZWxlY3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzZWxlY3QnKTsKICAgIE1FVFJJQ19ST1dTLmZvckVhY2goKG0pID0+IHsKICAgICAgY29uc3Qgb3B0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7IG9wdC52YWx1ZSA9IG0ua2V5OyBvcHQudGV4dENvbnRlbnQgPSBtLmxhYmVsOwogICAgICBpZiAobS5rZXkgPT09IHBsYXRmb3JtQ2hhcnRNZXRyaWMpIG9wdC5zZWxlY3RlZCA9IHRydWU7CiAgICAgIG1ldHJpY1NlbGVjdC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgICBtZXRyaWNTZWxlY3QuYWRkRXZlbnRMaXN0ZW5lcignY2hhbmdlJywgKCkgPT4gewogICAgICBwbGF0Zm9ybUNoYXJ0TWV0cmljID0gbWV0cmljU2VsZWN0LnZhbHVlOwogICAgICBkcmF3UGxhdGZvcm1SZXBvcnRDaGFydChyZXBvcnQpOwogICAgfSk7CiAgICBjaGFydEhlYWRlci5hcHBlbmRDaGlsZChtZXRyaWNTZWxlY3QpOwogICAgY29uc3QgY2hhcnRXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjaGFydFdyYXAuY2xhc3NOYW1lID0gJ2NoYXJ0LXdyYXAgdGFsbCc7CiAgICBjaGFydFdyYXAuaW5uZXJIVE1MID0gJzxjYW52YXMgaWQ9InBsYXRmb3JtUmVwb3J0Q2FudmFzIj48L2NhbnZhcz4nOwogICAgY2hhcnRDYXJkLmFwcGVuZChjaGFydEhlYWRlciwgY2hhcnRXcmFwKTsKICAgIHdyYXAuYXBwZW5kQ2hpbGQoY2hhcnRDYXJkKTsKICAgIGRyYXdQbGF0Zm9ybVJlcG9ydENoYXJ0KHJlcG9ydCk7CgogICAgY29uc3Qgd2l0aEZvbGxvd2VycyA9IHJlcG9ydC5wbGF0Zm9ybXMuZmlsdGVyKChwKSA9PiBwLmZvbGxvd2Vycy5sYXRlc3QgIT09IG51bGwpOwogICAgaWYgKHdpdGhGb2xsb3dlcnMubGVuZ3RoKSB7CiAgICAgIGNvbnN0IGZvbGxvd2VyQ2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBmb2xsb3dlckNhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgICBmb2xsb3dlckNhcmQuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMycsICdGb2xsb3dlciBHcm93dGggYnkgUGxhdGZvcm0nKSk7CiAgICAgIGNvbnN0IGZDaGFydFdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgZkNoYXJ0V3JhcC5jbGFzc05hbWUgPSAnY2hhcnQtd3JhcCB0YWxsJzsKICAgICAgZkNoYXJ0V3JhcC5pbm5lckhUTUwgPSAnPGNhbnZhcyBpZD0icGxhdGZvcm1Gb2xsb3dlckNhbnZhcyI+PC9jYW52YXM+JzsKICAgICAgZm9sbG93ZXJDYXJkLmFwcGVuZENoaWxkKGZDaGFydFdyYXApOwogICAgICB3cmFwLmFwcGVuZENoaWxkKGZvbGxvd2VyQ2FyZCk7CiAgICAgIENoYXJ0cy5wbGF0Zm9ybUJhckNoYXJ0KCdwbGF0Zm9ybUZvbGxvd2VyQ2FudmFzJywgewogICAgICAgIGxhYmVsczogd2l0aEZvbGxvd2Vycy5tYXAoKHApID0+IHAubGFiZWwpLAogICAgICAgIGRhdGE6IHdpdGhGb2xsb3dlcnMubWFwKChwKSA9PiBwLmZvbGxvd2Vycy5sYXRlc3QgfHwgMCksCiAgICAgICAgY29sb3JzOiB3aXRoRm9sbG93ZXJzLm1hcCgocCkgPT4gcC5jb2xvciksCiAgICAgICAgZm9ybWF0VmFsdWU6ICh2KSA9PiBGb3JtYXQuc21hcnQodiksCiAgICAgIH0pOwogICAgfQoKICAgIGNvbnN0IGluc2lnaHRzQ2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgaW5zaWdodHNDYXJkLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGluc2lnaHRzQ2FyZC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2gzJywgJ0luc2lnaHRzICYgU3VtbWFyeScpKTsKICAgIGNvbnN0IGxpc3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd1bCcpOwogICAgbGlzdC5jbGFzc05hbWUgPSAnaW5zaWdodHMtbGlzdCc7CiAgICByZXBvcnQuaW5zaWdodHMuZm9yRWFjaCgobGluZSkgPT4gbGlzdC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2xpJywgbGluZSkpKTsKICAgIGluc2lnaHRzQ2FyZC5hcHBlbmRDaGlsZChsaXN0KTsKICAgIHdyYXAuYXBwZW5kQ2hpbGQoaW5zaWdodHNDYXJkKTsKICB9CgogIGZ1bmN0aW9uIGRyYXdQbGF0Zm9ybVJlcG9ydENoYXJ0KHJlcG9ydCkgewogICAgQ2hhcnRzLnBsYXRmb3JtQmFyQ2hhcnQoJ3BsYXRmb3JtUmVwb3J0Q2FudmFzJywgewogICAgICBsYWJlbHM6IHJlcG9ydC5wbGF0Zm9ybXMubWFwKChwKSA9PiBwLmxhYmVsKSwKICAgICAgZGF0YTogcmVwb3J0LnBsYXRmb3Jtcy5tYXAoKHApID0+IHAudG90YWxzW3BsYXRmb3JtQ2hhcnRNZXRyaWNdIHx8IDApLAogICAgICBjb2xvcnM6IHJlcG9ydC5wbGF0Zm9ybXMubWFwKChwKSA9PiBwLmNvbG9yKSwKICAgICAgZm9ybWF0VmFsdWU6ICh2KSA9PiAocGxhdGZvcm1DaGFydE1ldHJpYyA9PT0gJ3dhdGNoX3RpbWVfc2Vjb25kcycgPyBGb3JtYXQuZHVyYXRpb24odikgOiBGb3JtYXQuc21hcnQodikpLAogICAgfSk7CiAgfQoKICAvKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgICAgUGxhdGZvcm0gUGVyZm9ybWFuY2UgQ29tcGFyaXNvbiDigJQgcmVwbGFjZXMgdGhlIG9sZCBncm91cGVkCiAgICAgIlJhbmdlIEEgdnMgUmFuZ2UgQiBieSBQbGF0Zm9ybSIgY2hhcnQuIE9uZSBjYXJkIHBlciBwbGF0Zm9ybQogICAgIHdpdGggYW55IGRhdGEgaW4gZWl0aGVyIHJhbmdlLCBidWlsdCBlbnRpcmVseSBmcm9tIHRoZSBzYW1lCiAgICAgY29tcGFyZVJhbmdlcygpIHJlc3BvbnNlIHRoZSBzdGF0LXRpbGUgZ3JpZCBhYm92ZSBhbHJlYWR5CiAgICAgdXNlcyAocmVzdWx0LnJhbmdlQS5wbGF0Zm9ybXMgLyByZXN1bHQucmFuZ2VCLnBsYXRmb3Jtcykg4oCUIG5vCiAgICAgZXh0cmEgZmV0Y2guCiAgICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwogIGNvbnN0IEFMTF9DQVJEX01FVFJJQ1MgPSBbeyBrZXk6ICdwb3N0X2NvdW50JywgbGFiZWw6ICdQb3N0cycgfSwgLi4uTUVUUklDX1JPV1NdOwogIGNvbnN0IENBUkRfU09SVF9NT0RFUyA9IFsKICAgIHsga2V5OiAnb3ZlcmFsbCcsIGxhYmVsOiAnT3ZlcmFsbCBQZXJmb3JtYW5jZScgfSwKICAgIHsga2V5OiAnZ3Jvd3RoJywgbGFiZWw6ICdIaWdoZXN0IEdyb3d0aCcgfSwKICAgIHsga2V5OiAnZW5nYWdlbWVudCcsIGxhYmVsOiAnSGlnaGVzdCBFbmdhZ2VtZW50JyB9LAogICAgeyBrZXk6ICdmb2xsb3dlcnMnLCBsYWJlbDogJ01vc3QgRm9sbG93ZXJzJyB9LAogICAgeyBrZXk6ICdwb3N0cycsIGxhYmVsOiAnTW9zdCBQb3N0cycgfSwKICAgIHsga2V5OiAnYWxwaGEnLCBsYWJlbDogJ0FscGhhYmV0aWNhbCcgfSwKICBdOwoKICAvKiogUGVyLW1ldHJpYyB7YSwgYiwgZGlmZiwgcGN0RGlmZn0gYWNyb3NzIGJvdGggcmFuZ2VzIGZvciBvbmUgcGxhdGZvcm0sIHNraXBwaW5nIGFueSBtZXRyaWMgdGhhdCdzIHplcm8gaW4gYm90aCDigJQgYSBwbGF0Zm9ybSdzIGNhcmQgc2hvdWxkIG9ubHkgZXZlciBzaG93IG1ldHJpY3MgaXQgYWN0dWFsbHkgaGFzLiAqLwogIGZ1bmN0aW9uIGNvbXB1dGVDYXJkTWV0cmljcyhwbGF0Zm9ybUEsIHBsYXRmb3JtQikgewogICAgY29uc3QgbWV0cmljcyA9IFtdOwogICAgQUxMX0NBUkRfTUVUUklDUy5mb3JFYWNoKCh7IGtleSwgbGFiZWwgfSkgPT4gewogICAgICBjb25zdCBhID0gKHBsYXRmb3JtQSAmJiBwbGF0Zm9ybUFba2V5XSkgfHwgMDsKICAgICAgY29uc3QgYiA9IChwbGF0Zm9ybUIgJiYgcGxhdGZvcm1CW2tleV0pIHx8IDA7CiAgICAgIGlmIChhID09PSAwICYmIGIgPT09IDApIHJldHVybjsKICAgICAgY29uc3QgZGlmZiA9IGEgLSBiOwogICAgICBjb25zdCBwY3REaWZmID0gYiA/IE1hdGgucm91bmQoKGRpZmYgLyBiKSAqIDEwMDApIC8gMTAgOiAoYSA+IDAgPyBudWxsIDogMCk7CiAgICAgIG1ldHJpY3MucHVzaCh7IGtleSwgbGFiZWwsIGEsIGIsIGRpZmYsIHBjdERpZmYsIGlzRHVyYXRpb246IGtleSA9PT0gJ3dhdGNoX3RpbWVfc2Vjb25kcycgfSk7CiAgICB9KTsKICAgIHJldHVybiBtZXRyaWNzOwogIH0KCiAgLyoqIEEgc2luZ2xlICJob3cgZGlkIHRoaXMgcGxhdGZvcm0gZG8gb3ZlcmFsbCIgbnVtYmVyOiB0aGUgYXZlcmFnZSAlIGNoYW5nZSBhY3Jvc3MgZXZlcnkgbWV0cmljIHRoYXQgaGFzIGEgY29tcHV0YWJsZSBwZXJjZW50YWdlIChhIG1ldHJpYyBnb2luZyBmcm9tIDAgdG8gc29tZXRoaW5nIGhhcyBubyBwZXJjZW50YWdlIOKAlCAibmV3Iiwgbm90IGNvdW50ZWQgZWl0aGVyIHdheSkuICovCiAgZnVuY3Rpb24gb3ZlcmFsbFBjdENoYW5nZShtZXRyaWNzKSB7CiAgICBjb25zdCB3aXRoUGN0ID0gbWV0cmljcy5maWx0ZXIoKG0pID0+IG0ucGN0RGlmZiAhPT0gbnVsbCk7CiAgICBpZiAoIXdpdGhQY3QubGVuZ3RoKSByZXR1cm4gbnVsbDsKICAgIHJldHVybiBNYXRoLnJvdW5kKCh3aXRoUGN0LnJlZHVjZSgoc3VtLCBtKSA9PiBzdW0gKyBtLnBjdERpZmYsIDApIC8gd2l0aFBjdC5sZW5ndGgpICogMTApIC8gMTA7CiAgfQoKICBmdW5jdGlvbiBiZXN0V2Vha2VzdE1ldHJpYyhtZXRyaWNzKSB7CiAgICBjb25zdCB3aXRoUGN0ID0gbWV0cmljcy5maWx0ZXIoKG0pID0+IG0ucGN0RGlmZiAhPT0gbnVsbCk7CiAgICBpZiAoIXdpdGhQY3QubGVuZ3RoKSByZXR1cm4geyBiZXN0OiBudWxsLCB3ZWFrZXN0OiBudWxsIH07CiAgICBjb25zdCBiZXN0ID0gd2l0aFBjdC5yZWR1Y2UoKGEsIGIpID0+IChiLnBjdERpZmYgPiBhLnBjdERpZmYgPyBiIDogYSkpOwogICAgY29uc3Qgd2Vha2VzdCA9IHdpdGhQY3QucmVkdWNlKChhLCBiKSA9PiAoYi5wY3REaWZmIDwgYS5wY3REaWZmID8gYiA6IGEpKTsKICAgIHJldHVybiB7IGJlc3QsIHdlYWtlc3QgfTsKICB9CgogIGZ1bmN0aW9uIHRyZW5kRGlyZWN0aW9uKHBjdCkgewogICAgaWYgKHBjdCA9PT0gbnVsbCB8fCBwY3QgPT09IHVuZGVmaW5lZCkgcmV0dXJuICdmbGF0JzsKICAgIGlmIChwY3QgPiAwLjUpIHJldHVybiAndXAnOwogICAgaWYgKHBjdCA8IC0wLjUpIHJldHVybiAnZG93bic7CiAgICByZXR1cm4gJ2ZsYXQnOwogIH0KCiAgZnVuY3Rpb24gYnVpbGRQbGF0Zm9ybUNhcmRzKHJlc3VsdCkgewogICAgY29uc3QgcGxhdGZvcm1PcHRpb25zID0gKHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZSB8fCB7IGFsbFBsYXRmb3JtczogW10gfSkuYWxsUGxhdGZvcm1zOwogICAgY29uc3QgaWRzID0gWy4uLm5ldyBTZXQoWy4uLnJlc3VsdC5yYW5nZUEucGxhdGZvcm1zLCAuLi5yZXN1bHQucmFuZ2VCLnBsYXRmb3Jtc10ubWFwKChwKSA9PiBwLnBsYXRmb3JtKSldOwogICAgY29uc3QgYnlJZEEgPSBPYmplY3QuZnJvbUVudHJpZXMocmVzdWx0LnJhbmdlQS5wbGF0Zm9ybXMubWFwKChwKSA9PiBbcC5wbGF0Zm9ybSwgcF0pKTsKICAgIGNvbnN0IGJ5SWRCID0gT2JqZWN0LmZyb21FbnRyaWVzKHJlc3VsdC5yYW5nZUIucGxhdGZvcm1zLm1hcCgocCkgPT4gW3AucGxhdGZvcm0sIHBdKSk7CgogICAgcmV0dXJuIGlkcwogICAgICAubWFwKChpZCkgPT4gewogICAgICAgIGNvbnN0IG1ldGEgPSBwbGF0Zm9ybU9wdGlvbnMuZmluZCgocCkgPT4gcC5pZCA9PT0gaWQpIHx8IHsgaWQsIGxhYmVsOiBpZCwgY29sb3I6ICd2YXIoLS1zZXJpZXMtMSknIH07CiAgICAgICAgY29uc3QgYSA9IGJ5SWRBW2lkXSB8fCBudWxsOwogICAgICAgIGNvbnN0IGIgPSBieUlkQltpZF0gfHwgbnVsbDsKICAgICAgICBjb25zdCBtZXRyaWNzID0gY29tcHV0ZUNhcmRNZXRyaWNzKGEsIGIpOwogICAgICAgIGNvbnN0IHsgYmVzdCwgd2Vha2VzdCB9ID0gYmVzdFdlYWtlc3RNZXRyaWMobWV0cmljcyk7CiAgICAgICAgcmV0dXJuIHsKICAgICAgICAgIHBsYXRmb3JtOiBpZCwKICAgICAgICAgIGxhYmVsOiBtZXRhLmxhYmVsLAogICAgICAgICAgY29sb3I6IG1ldGEuY29sb3IsCiAgICAgICAgICBtZXRyaWNzLAogICAgICAgICAgb3ZlcmFsbDogb3ZlcmFsbFBjdENoYW5nZShtZXRyaWNzKSwKICAgICAgICAgIGJlc3QsCiAgICAgICAgICB3ZWFrZXN0LAogICAgICAgICAgZm9sbG93ZXJzR2FpbmVkOiAoYSA/IGEuZm9sbG93ZXJzX2dhaW5lZCB8fCAwIDogMCkgKyAoYiA/IGIuZm9sbG93ZXJzX2dhaW5lZCB8fCAwIDogMCksCiAgICAgICAgICBwb3N0czogKGEgPyBhLnBvc3RfY291bnQgfHwgMCA6IDApICsgKGIgPyBiLnBvc3RfY291bnQgfHwgMCA6IDApLAogICAgICAgICAgZW5nYWdlbWVudFRvdGFsOiAoYSA/IGEuZW5nYWdlbWVudCB8fCAwIDogMCkgKyAoYiA/IGIuZW5nYWdlbWVudCB8fCAwIDogMCksCiAgICAgICAgfTsKICAgICAgfSkKICAgICAgLmZpbHRlcigoY2FyZCkgPT4gY2FyZC5tZXRyaWNzLmxlbmd0aCA+IDApOwogIH0KCiAgZnVuY3Rpb24gc29ydENhcmRzKGNhcmRzLCBzb3J0TW9kZSkgewogICAgY29uc3QgYXJyID0gWy4uLmNhcmRzXTsKICAgIGlmIChzb3J0TW9kZSA9PT0gJ2VuZ2FnZW1lbnQnKSByZXR1cm4gYXJyLnNvcnQoKHgsIHkpID0+IHkuZW5nYWdlbWVudFRvdGFsIC0geC5lbmdhZ2VtZW50VG90YWwpOwogICAgaWYgKHNvcnRNb2RlID09PSAnZm9sbG93ZXJzJykgcmV0dXJuIGFyci5zb3J0KCh4LCB5KSA9PiB5LmZvbGxvd2Vyc0dhaW5lZCAtIHguZm9sbG93ZXJzR2FpbmVkKTsKICAgIGlmIChzb3J0TW9kZSA9PT0gJ3Bvc3RzJykgcmV0dXJuIGFyci5zb3J0KCh4LCB5KSA9PiB5LnBvc3RzIC0geC5wb3N0cyk7CiAgICBpZiAoc29ydE1vZGUgPT09ICdhbHBoYScpIHJldHVybiBhcnIuc29ydCgoeCwgeSkgPT4geC5sYWJlbC5sb2NhbGVDb21wYXJlKHkubGFiZWwpKTsKICAgIC8vICdvdmVyYWxsJyBhbmQgJ2dyb3d0aCcgYm90aCByYW5rIGJ5IHRoZSBzYW1lIGNvbXBvc2l0ZSAlIGNoYW5nZSDigJQgdGhlIHR3byBsYWJlbHMKICAgIC8vIHJlYWQgZGlmZmVyZW50bHkgb24gdGhlIHNhbWUgdW5kZXJseWluZyBudW1iZXIsIHBlciB0aGUgcmVxdWVzdGVkIG9wdGlvbiBsaXN0LgogICAgcmV0dXJuIGFyci5zb3J0KCh4LCB5KSA9PiAoeS5vdmVyYWxsID8/IC1JbmZpbml0eSkgLSAoeC5vdmVyYWxsID8/IC1JbmZpbml0eSkpOwogIH0KCiAgZnVuY3Rpb24gYnVpbGRNZXRyaWNSb3cobSkgewogICAgY29uc3Qgcm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICByb3cuY2xhc3NOYW1lID0gJ3BjYy1tZXRyaWMtcm93JzsKICAgIGNvbnN0IGhlYWRlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgaGVhZGVyLmNsYXNzTmFtZSA9ICdwY2MtbWV0cmljLWhlYWRlcic7CiAgICBjb25zdCBmbXQgPSAodikgPT4gKG0uaXNEdXJhdGlvbiA/IEZvcm1hdC5kdXJhdGlvbih2KSA6IEZvcm1hdC5zbWFydCh2KSk7CiAgICBjb25zdCBkaWZmVGV4dCA9IG0ucGN0RGlmZiA9PT0gbnVsbAogICAgICA/IGAke20uZGlmZiA+IDAgPyAnKycgOiAnJ30ke2ZtdChtLmRpZmYpfSAobmV3KWAKICAgICAgOiBgJHttLmRpZmYgPiAwID8gJysnIDogJyd9JHtmbXQobS5kaWZmKX0gKCR7Rm9ybWF0LnBjdChtLnBjdERpZmYpfSlgOwogICAgaGVhZGVyLmFwcGVuZCgKICAgICAgdGV4dEVsKCdzcGFuJywgbS5sYWJlbCwgJ3BjYy1tZXRyaWMtbGFiZWwnKSwKICAgICAgdGV4dEVsKCdzcGFuJywgZGlmZlRleHQsIGBwY2MtbWV0cmljLWRpZmYgJHtGb3JtYXQuZGVsdGFDbGFzcyhtLnBjdERpZmYpfWApCiAgICApOwogICAgcm93LmFwcGVuZENoaWxkKGhlYWRlcik7CiAgICBjb25zdCBtYXggPSBNYXRoLm1heChtLmEsIG0uYiwgMSk7CiAgICAvLyBtLmEgaXMgYWx3YXlzIHRoZSBjdXJyZW50IHBlcmlvZCBhbmQgbS5iIGFsd2F5cyB0aGUgcHJldmlvdXMgcGVyaW9kIChzZWUKICAgIC8vIGNvbXB1dGVDYXJkTWV0cmljcykgcmVnYXJkbGVzcyBvZiBjb21wYXJpc29uIG1vZGUsIHNvIHRoZXNlIGxhYmVscyBjYW4gYmUKICAgIC8vIGhhcmRjb2RlZCByYXRoZXIgdGhhbiBuZWVkaW5nIHRoZSBtb2RlLXNwZWNpZmljIGxhYmVsQS9sYWJlbEIgdGV4dCDigJQgdW5saWtlCiAgICAvLyB0aGUgZ2VuZXJpYyAiUmFuZ2UgQSIvIlJhbmdlIEIiIHdvcmRpbmcgdGhpcyByZXBsYWNlZCwgd2hpY2ggcmVhZCBhcwogICAgLy8gYXJiaXRyYXJ5IGxldHRlcnMgd2l0aCBubyBpbmRpY2F0aW9uIG9mIHdoaWNoIHNpZGUgd2FzIG1vcmUgcmVjZW50LgogICAgcm93LmFwcGVuZENoaWxkKGJ1aWxkQmFyKHsgbGFiZWw6ICdDdXJyZW50JywgdmFsdWU6IG0uYSwgbWF4LCBjb2xvclZhcjogJy0tc2VyaWVzLTEnLCBmb3JtYXRWYWx1ZTogZm10IH0pKTsKICAgIHJvdy5hcHBlbmRDaGlsZChidWlsZEJhcih7IGxhYmVsOiAnUHJldmlvdXMnLCB2YWx1ZTogbS5iLCBtYXgsIGNvbG9yVmFyOiAnLS10ZXh0LW11dGVkJywgZm9ybWF0VmFsdWU6IGZtdCB9KSk7CiAgICByZXR1cm4gcm93OwogIH0KCiAgZnVuY3Rpb24gYnVpbGRDYXJkRm9vdGVyKGNhcmQpIHsKICAgIGNvbnN0IGZvb3RlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZm9vdGVyLmNsYXNzTmFtZSA9ICdwY2MtZm9vdGVyJzsKICAgIGNvbnN0IGRpciA9IHRyZW5kRGlyZWN0aW9uKGNhcmQub3ZlcmFsbCk7CiAgICBjb25zdCByZXN1bHRUZXh0ID0gY2FyZC5vdmVyYWxsID09PSBudWxsCiAgICAgID8gJ05vdCBlbm91Z2ggZGF0YSB0byBjb21wYXJlJwogICAgICA6IGAke2RpciA9PT0gJ3VwJyA/ICdJbXByb3ZlZCcgOiBkaXIgPT09ICdkb3duJyA/ICdEZWNsaW5lZCcgOiAnTm8gc2lnbmlmaWNhbnQgY2hhbmdlJ30ke2RpciAhPT0gJ2ZsYXQnID8gYCBieSAke01hdGguYWJzKGNhcmQub3ZlcmFsbCl9JWAgOiAnJ31gOwogICAgZm9vdGVyLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgJ092ZXJhbGwgUmVzdWx0JywgJ3BjYy1mb290ZXItbGFiZWwnKSk7CiAgICBmb290ZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCByZXN1bHRUZXh0LCBgcGNjLWZvb3Rlci12YWx1ZSAke2Rpcn1gKSk7CiAgICBpZiAoY2FyZC5iZXN0KSBmb290ZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBgQmVzdCBNZXRyaWM6ICR7Y2FyZC5iZXN0LmxhYmVsfSAoJHtGb3JtYXQucGN0KGNhcmQuYmVzdC5wY3REaWZmKX0pYCwgJ3BjYy1mb290ZXItZGV0YWlsJykpOwogICAgaWYgKGNhcmQud2Vha2VzdCAmJiBjYXJkLndlYWtlc3QgIT09IGNhcmQuYmVzdCkgewogICAgICBmb290ZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBgV2Vha2VzdCBNZXRyaWM6ICR7Y2FyZC53ZWFrZXN0LmxhYmVsfSAoJHtGb3JtYXQucGN0KGNhcmQud2Vha2VzdC5wY3REaWZmKX0pYCwgJ3BjYy1mb290ZXItZGV0YWlsJykpOwogICAgfQogICAgcmV0dXJuIGZvb3RlcjsKICB9CgogIC8qKiBTZWxmLWNvbnRhaW5lZCBtb2RhbCBmb3IgIlZpZXcgRnVsbCBDb21wYXJpc29uIiDigJQgYSBzZXBhcmF0ZSBvdmVybGF5IGlkIGZyb20gdGhlIERhdGEgUmVjb3JkcyBFZGl0IG1vZGFsIChSZWNvcmRzLm1vZGFsU2hlbGwgaXMgYSBwcml2YXRlIGNsb3N1cmUgb2YgdGhhdCBtb2R1bGUsIG5vdCBzaGFyZWQgc3RhdGUpLCBzYW1lIHZpc3VhbCBsYW5ndWFnZSAoLm1vZGFsLW92ZXJsYXkgLyAubW9kYWwtcGFuZWwpIHNvIGl0IGxvb2tzIGlkZW50aWNhbC4gKi8KICBmdW5jdGlvbiBjbG9zZUNhcmRNb2RhbCgpIHsKICAgIGNvbnN0IG92ZXJsYXkgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnY29tcGFyaXNvbk1vZGFsT3ZlcmxheScpOwogICAgaWYgKG92ZXJsYXkpIG92ZXJsYXkucmVtb3ZlKCk7CiAgfQoKICBmdW5jdGlvbiBvcGVuQ2FyZE1vZGFsKGNhcmQsIGxhYmVsQSwgbGFiZWxCKSB7CiAgICBjbG9zZUNhcmRNb2RhbCgpOwogICAgY29uc3Qgb3ZlcmxheSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgb3ZlcmxheS5jbGFzc05hbWUgPSAnbW9kYWwtb3ZlcmxheSc7CiAgICBvdmVybGF5LmlkID0gJ2NvbXBhcmlzb25Nb2RhbE92ZXJsYXknOwogICAgb3ZlcmxheS5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIChlKSA9PiB7IGlmIChlLnRhcmdldCA9PT0gb3ZlcmxheSkgY2xvc2VDYXJkTW9kYWwoKTsgfSk7CiAgICBjb25zdCBwYW5lbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcGFuZWwuY2xhc3NOYW1lID0gJ21vZGFsLXBhbmVsIHdpZGUnOwogICAgcGFuZWwuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMicsIGAke2NhcmQubGFiZWx9IOKAlCBGdWxsIENvbXBhcmlzb25gKSk7CiAgICBwYW5lbC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIGAke2xhYmVsQX0gdnMgJHtsYWJlbEJ9YCwgJ21vZGFsLXN1YicpKTsKCiAgICBjb25zdCB0YWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICB0YWJsZS5jbGFzc05hbWUgPSAnZGF0YS10YWJsZSc7CiAgICBjb25zdCB0aGVhZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RoZWFkJyk7CiAgICBjb25zdCBoZWFkUm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgIC8vIGxhYmVsQS9sYWJlbEIgYXJlIHRoZSBzYW1lIG1vZGUtc3BlY2lmaWMgdGV4dCBzaG93biBpbiB0aGUgbW9kYWwncyBzdWJ0aXRsZQogICAgLy8gYWJvdmUgKGUuZy4gIldlZWsgb2YgSmFuIDEzLCAyMDI2IiAvICJXZWVrIG9mIERlYyAzMCwgMjAyNSIsIG9yICJDdXJyZW50CiAgICAvLyBwZXJpb2QiIC8gIlByZXZpb3VzIHBlcmlvZCIpIOKAlCByZXVzZWQgaGVyZSBpbnN0ZWFkIG9mIGdlbmVyaWMgIlJhbmdlIEEiLwogICAgLy8gIlJhbmdlIEIiIHNvIHRoZSBjb2x1bW4gaGVhZGVycyBhbHdheXMgc2F5IHdoYXQgcGVyaW9kIHRoZXkgYWN0dWFsbHkgaG9sZC4KICAgIGNvbnN0IG51bVRoID0gKHRleHQpID0+IHsgY29uc3QgdGggPSB0ZXh0RWwoJ3RoJywgdGV4dCk7IHRoLmNsYXNzTGlzdC5hZGQoJ251bScpOyByZXR1cm4gdGg7IH07CiAgICBoZWFkUm93LmFwcGVuZCh0ZXh0RWwoJ3RoJywgJ01ldHJpYycpLCBudW1UaChsYWJlbEEpLCBudW1UaChsYWJlbEIpLCBudW1UaCgnRGlmZmVyZW5jZScpLCBudW1UaCgnJSBEaWZmZXJlbmNlJyksIHRleHRFbCgndGgnLCAnVHJlbmQnKSk7CiAgICB0aGVhZC5hcHBlbmRDaGlsZChoZWFkUm93KTsKICAgIGNvbnN0IHRib2R5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGJvZHknKTsKICAgIGNhcmQubWV0cmljcy5mb3JFYWNoKChtKSA9PiB7CiAgICAgIGNvbnN0IGZtdCA9ICh2KSA9PiAobS5pc0R1cmF0aW9uID8gRm9ybWF0LmR1cmF0aW9uKHYpIDogRm9ybWF0LnNtYXJ0KHYpKTsKICAgICAgY29uc3QgdHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgICBjb25zdCB0cmVuZEVsID0gdGV4dEVsKCdzcGFuJywgdHJlbmREaXJlY3Rpb24obS5wY3REaWZmKSA9PT0gJ3VwJyA/ICfilrInIDogdHJlbmREaXJlY3Rpb24obS5wY3REaWZmKSA9PT0gJ2Rvd24nID8gJ+KWvCcgOiAn4oCUJywgYHN0YXQtZGVsdGEgJHtGb3JtYXQuZGVsdGFDbGFzcyhtLnBjdERpZmYpfWApOwogICAgICBjb25zdCB0cmVuZFRkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgdHJlbmRUZC5hcHBlbmRDaGlsZCh0cmVuZEVsKTsKICAgICAgdHIuYXBwZW5kKAogICAgICAgIHRleHRFbCgndGQnLCBtLmxhYmVsKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgZm10KG0uYSksICdudW0nKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgZm10KG0uYiksICdudW0nKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgYCR7bS5kaWZmID4gMCA/ICcrJyA6ICcnfSR7Zm10KG0uZGlmZil9YCwgJ251bScpLAogICAgICAgIHRleHRFbCgndGQnLCBtLnBjdERpZmYgPT09IG51bGwgPyAnbmV3JyA6IEZvcm1hdC5wY3QobS5wY3REaWZmKSwgJ251bScpLAogICAgICAgIHRyZW5kVGQKICAgICAgKTsKICAgICAgdGJvZHkuYXBwZW5kQ2hpbGQodHIpOwogICAgfSk7CiAgICB0YWJsZS5hcHBlbmQodGhlYWQsIHRib2R5KTsKICAgIGNvbnN0IHRhYmxlU2Nyb2xsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0YWJsZVNjcm9sbC5jbGFzc05hbWUgPSAndGFibGUtc2Nyb2xsJzsKICAgIHRhYmxlU2Nyb2xsLmFwcGVuZENoaWxkKHRhYmxlKTsKICAgIHBhbmVsLmFwcGVuZENoaWxkKHRhYmxlU2Nyb2xsKTsKCiAgICBjb25zdCBhY3Rpb25zID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBhY3Rpb25zLmNsYXNzTmFtZSA9ICdtb2RhbC1hY3Rpb25zJzsKICAgIGNvbnN0IGNsb3NlQnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBjbG9zZUJ0bi5jbGFzc05hbWUgPSAnYnRuJzsKICAgIGNsb3NlQnRuLnR5cGUgPSAnYnV0dG9uJzsKICAgIGNsb3NlQnRuLnRleHRDb250ZW50ID0gJ0Nsb3NlJzsKICAgIGNsb3NlQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgY2xvc2VDYXJkTW9kYWwpOwogICAgYWN0aW9ucy5hcHBlbmRDaGlsZChjbG9zZUJ0bik7CiAgICBwYW5lbC5hcHBlbmRDaGlsZChhY3Rpb25zKTsKCiAgICBvdmVybGF5LmFwcGVuZENoaWxkKHBhbmVsKTsKICAgIGRvY3VtZW50LmJvZHkuYXBwZW5kQ2hpbGQob3ZlcmxheSk7CiAgfQoKICBmdW5jdGlvbiBidWlsZFBsYXRmb3JtQ2FyZChjYXJkLCBsYWJlbEEsIGxhYmVsQikgewogICAgY29uc3QgZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGVsLmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1jb21wYXJlLWNhcmQnOwoKICAgIGNvbnN0IGhlYWRlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgaGVhZGVyLmNsYXNzTmFtZSA9ICdwY2MtaGVhZGVyJzsKICAgIGNvbnN0IG5hbWVXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBuYW1lV3JhcC5jbGFzc05hbWUgPSAncGNjLWhlYWRlci1uYW1lJzsKICAgIGNvbnN0IGRvdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgIGRvdC5jbGFzc05hbWUgPSAncGxhdGZvcm0tZG90JzsKICAgIGRvdC5zdHlsZS5iYWNrZ3JvdW5kID0gY2FyZC5jb2xvcjsKICAgIG5hbWVXcmFwLmFwcGVuZChkb3QsIHRleHRFbCgnc3BhbicsIGNhcmQubGFiZWwsICdwY2MtbmFtZScpKTsKICAgIGhlYWRlci5hcHBlbmRDaGlsZChuYW1lV3JhcCk7CiAgICBjb25zdCBkaXIgPSB0cmVuZERpcmVjdGlvbihjYXJkLm92ZXJhbGwpOwogICAgaGVhZGVyLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgY2FyZC5vdmVyYWxsID09PSBudWxsID8gJ+KAlCcgOiBgJHtkaXIgPT09ICd1cCcgPyAn4payJyA6IGRpciA9PT0gJ2Rvd24nID8gJ+KWvCcgOiAn4oCUJ30gJHtGb3JtYXQucGN0KGNhcmQub3ZlcmFsbCl9YCwgYHBjYy1iYWRnZSAke2Rpcn1gKSk7CiAgICBlbC5hcHBlbmRDaGlsZChoZWFkZXIpOwoKICAgIGNvbnN0IGNhcHRpb24gPSBjYXJkLm92ZXJhbGwgPT09IG51bGwKICAgICAgPyAnTm90IGVub3VnaCBkYXRhIHRvIGNvbXBhcmUgeWV0JwogICAgICA6IGRpciA9PT0gJ3VwJyA/ICdJbXByb3ZlZCBjb21wYXJlZCB0byBwcmV2aW91cyBwZXJpb2QnCiAgICAgIDogZGlyID09PSAnZG93bicgPyAnTG93ZXIgdGhhbiBwcmV2aW91cyBwZXJpb2QnCiAgICAgIDogJ0Fib3V0IHRoZSBzYW1lIGFzIHRoZSBwcmV2aW91cyBwZXJpb2QnOwogICAgZWwuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBjYXB0aW9uLCAncGNjLWNhcHRpb24nKSk7CgogICAgY29uc3QgbWV0cmljc1dyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIG1ldHJpY3NXcmFwLmNsYXNzTmFtZSA9ICdwY2MtbWV0cmljcyc7CiAgICBjYXJkLm1ldHJpY3MuZm9yRWFjaCgobSkgPT4gbWV0cmljc1dyYXAuYXBwZW5kQ2hpbGQoYnVpbGRNZXRyaWNSb3cobSkpKTsKICAgIGVsLmFwcGVuZENoaWxkKG1ldHJpY3NXcmFwKTsKCiAgICBlbC5hcHBlbmRDaGlsZChidWlsZENhcmRGb290ZXIoY2FyZCkpOwoKICAgIGNvbnN0IHZpZXdMaW5rID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICB2aWV3TGluay50eXBlID0gJ2J1dHRvbic7CiAgICB2aWV3TGluay5jbGFzc05hbWUgPSAncGNjLXZpZXctbGluayc7CiAgICB2aWV3TGluay50ZXh0Q29udGVudCA9ICdWaWV3IEZ1bGwgQ29tcGFyaXNvbiDihpInOwogICAgdmlld0xpbmsuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBvcGVuQ2FyZE1vZGFsKGNhcmQsIGxhYmVsQSwgbGFiZWxCKSk7CiAgICBlbC5hcHBlbmRDaGlsZCh2aWV3TGluayk7CgogICAgcmV0dXJuIGVsOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyUGxhdGZvcm1Db21wYXJpc29uQ2FyZHMod3JhcCwgcmVzdWx0LCBsYWJlbEEsIGxhYmVsQikgewogICAgY29uc3QgYWxsQ2FyZHMgPSBidWlsZFBsYXRmb3JtQ2FyZHMocmVzdWx0KTsKCiAgICBjb25zdCBzZWN0aW9uID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBzZWN0aW9uLmNsYXNzTmFtZSA9ICdwY2Mtc2VjdGlvbic7CiAgICBzZWN0aW9uLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgJ1BsYXRmb3JtIFBlcmZvcm1hbmNlIENvbXBhcmlzb24nLCAnc2VjdGlvbi10aXRsZScpKTsKCiAgICBpZiAoIWFsbENhcmRzLmxlbmd0aCkgewogICAgICBzZWN0aW9uLmFwcGVuZENoaWxkKGVtcHR5U3RhdGUoewogICAgICAgIGljb246ICdnaXQtY29tcGFyZScsCiAgICAgICAgdGl0bGU6ICdObyBkYXRhIGF2YWlsYWJsZSBmb3IgdGhlIHNlbGVjdGVkIGRhdGUgcmFuZ2VzLicsCiAgICAgICAgbWVzc2FnZTogJ1RyeSBhIHdpZGVyIHJhbmdlLCBvciBjaGVjayB0aGF0IHBvc3RzIGV4aXN0IGZvciBhdCBsZWFzdCBvbmUgcGxhdGZvcm0gaW4gUmFuZ2UgQSBvciBSYW5nZSBCLicsCiAgICAgIH0pKTsKICAgICAgd3JhcC5hcHBlbmRDaGlsZChzZWN0aW9uKTsKICAgICAgcmV0dXJuOwogICAgfQoKICAgIGNvbnN0IGNvbnRyb2xzID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjb250cm9scy5jbGFzc05hbWUgPSAncGNjLWNvbnRyb2xzJzsKCiAgICBjb25zdCBzb3J0U2VsZWN0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc2VsZWN0Jyk7CiAgICBDQVJEX1NPUlRfTU9ERVMuZm9yRWFjaCgobSkgPT4gewogICAgICBjb25zdCBvcHQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsgb3B0LnZhbHVlID0gbS5rZXk7IG9wdC50ZXh0Q29udGVudCA9IG0ubGFiZWw7CiAgICAgIGlmIChtLmtleSA9PT0gY2FyZFNvcnRNb2RlKSBvcHQuc2VsZWN0ZWQgPSB0cnVlOwogICAgICBzb3J0U2VsZWN0LmFwcGVuZENoaWxkKG9wdCk7CiAgICB9KTsKICAgIHNvcnRTZWxlY3QuYWRkRXZlbnRMaXN0ZW5lcignY2hhbmdlJywgKCkgPT4geyBjYXJkU29ydE1vZGUgPSBzb3J0U2VsZWN0LnZhbHVlOyByZW5kZXJDYXJkR3JpZCgpOyB9KTsKICAgIGNvbnRyb2xzLmFwcGVuZENoaWxkKGxhYmVsZWQoJ1NvcnQgQnknLCBzb3J0U2VsZWN0KSk7CgogICAgY29uc3QgZmlsdGVyUGlsbHMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGZpbHRlclBpbGxzLmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1maWx0ZXItcGlsbHMnOwogICAgY29uc3QgYWxsQnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBhbGxCdG4udHlwZSA9ICdidXR0b24nOwogICAgYWxsQnRuLmRhdGFzZXQuZmlsdGVyID0gJ2FsbCc7CiAgICBhbGxCdG4udGV4dENvbnRlbnQgPSAnQWxsIFBsYXRmb3Jtcyc7CiAgICBhbGxCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7IGNhcmRQbGF0Zm9ybUZpbHRlciA9ICdhbGwnOyByZW5kZXJDYXJkR3JpZCgpOyB9KTsKICAgIGZpbHRlclBpbGxzLmFwcGVuZENoaWxkKGFsbEJ0bik7CiAgICBhbGxDYXJkcy5mb3JFYWNoKChjYXJkKSA9PiB7CiAgICAgIGNvbnN0IGJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgICBidG4udHlwZSA9ICdidXR0b24nOwogICAgICBidG4uZGF0YXNldC5maWx0ZXIgPSBjYXJkLnBsYXRmb3JtOwogICAgICBjb25zdCBkb3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICAgIGRvdC5jbGFzc05hbWUgPSAncGxhdGZvcm0tZG90JzsKICAgICAgZG90LnN0eWxlLmJhY2tncm91bmQgPSBjYXJkLmNvbG9yOwogICAgICBidG4uYXBwZW5kKGRvdCwgZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUoY2FyZC5sYWJlbCkpOwogICAgICBidG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7IGNhcmRQbGF0Zm9ybUZpbHRlciA9IGNhcmQucGxhdGZvcm07IHJlbmRlckNhcmRHcmlkKCk7IH0pOwogICAgICBmaWx0ZXJQaWxscy5hcHBlbmRDaGlsZChidG4pOwogICAgfSk7CiAgICBjb250cm9scy5hcHBlbmRDaGlsZChmaWx0ZXJQaWxscyk7CiAgICBzZWN0aW9uLmFwcGVuZENoaWxkKGNvbnRyb2xzKTsKCiAgICBjb25zdCBncmlkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBncmlkLmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1jb21wYXJlLWdyaWQnOwogICAgZ3JpZC5pZCA9ICdwbGF0Zm9ybUNvbXBhcmVHcmlkJzsKICAgIHNlY3Rpb24uYXBwZW5kQ2hpbGQoZ3JpZCk7CiAgICB3cmFwLmFwcGVuZENoaWxkKHNlY3Rpb24pOwoKICAgIGZ1bmN0aW9uIHJlbmRlckNhcmRHcmlkKCkgewogICAgICBjb25zdCBncmlkRWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncGxhdGZvcm1Db21wYXJlR3JpZCcpOwogICAgICBpZiAoIWdyaWRFbCkgcmV0dXJuOwogICAgICBncmlkRWwuaW5uZXJIVE1MID0gJyc7CiAgICAgIGNvbnN0IHZpc2libGUgPSBjYXJkUGxhdGZvcm1GaWx0ZXIgPT09ICdhbGwnID8gYWxsQ2FyZHMgOiBhbGxDYXJkcy5maWx0ZXIoKGMpID0+IGMucGxhdGZvcm0gPT09IGNhcmRQbGF0Zm9ybUZpbHRlcik7CiAgICAgIGNvbnN0IHNvcnRlZCA9IHNvcnRDYXJkcyh2aXNpYmxlLCBjYXJkU29ydE1vZGUpOwogICAgICBpZiAoIXNvcnRlZC5sZW5ndGgpIHsKICAgICAgICBncmlkRWwuYXBwZW5kQ2hpbGQoZW1wdHlTdGF0ZSh7IGljb246ICdnaXQtY29tcGFyZScsIG1lc3NhZ2U6ICdObyBkYXRhIGZvciB0aGlzIHBsYXRmb3JtIGluIHRoZSBzZWxlY3RlZCBkYXRlIHJhbmdlcy4nIH0pKTsKICAgICAgfSBlbHNlIHsKICAgICAgICBzb3J0ZWQuZm9yRWFjaCgoY2FyZCkgPT4gZ3JpZEVsLmFwcGVuZENoaWxkKGJ1aWxkUGxhdGZvcm1DYXJkKGNhcmQsIGxhYmVsQSwgbGFiZWxCKSkpOwogICAgICB9CiAgICAgIGZpbHRlclBpbGxzLnF1ZXJ5U2VsZWN0b3JBbGwoJ2J1dHRvbicpLmZvckVhY2goKGJ0bikgPT4gewogICAgICAgIGJ0bi5jbGFzc0xpc3QudG9nZ2xlKCdpcy1hY3RpdmUnLCBidG4uZGF0YXNldC5maWx0ZXIgPT09IGNhcmRQbGF0Zm9ybUZpbHRlcik7CiAgICAgIH0pOwogICAgfQogICAgcmVuZGVyQ2FyZEdyaWQoKTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlcigpIHsKICAgIHJvb3QgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndmlldy1jb21wYXJpc29uJyk7CiAgICBzaGVsbCgpOwogIH0KCiAgcmV0dXJuIHsgcmVuZGVyIH07Cn0pKCk7CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgVXBsb2FkIHRhYjogZHJhZy1kcm9wLCB2YWxpZGF0aW9uIHByZXZpZXcsIHBlci13ZWVrIGNvbmZsaWN0CiAgIHJlc29sdXRpb24sIGNvbW1pdCDigJQgcGx1cyB0aGUgVXBsb2FkIEhpc3RvcnkgdGFiLgogICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KY29uc3QgVXBsb2FkID0gKCgpID0+IHsKICBsZXQgcm9vdDsKICBsZXQgY3VycmVudFByZXZpZXcgPSBudWxsOyAvLyB7IGZpbGVQYXRoLCBvcmlnaW5hbE5hbWUsIGR1cGxpY2F0ZXMsIGlzc3Vlcywgc2FtcGxlLCAuLi4gfQogIGNvbnN0IGR1cGxpY2F0ZUFjdGlvbk92ZXJyaWRlcyA9IHt9OwoKICBmdW5jdGlvbiBzaGVsbCgpIHsKICAgIHJvb3QuaW5uZXJIVE1MID0gJyc7CgogICAgY29uc3QgaW50cm8gPSB0ZXh0RWwoJ2RpdicsICdVcGxvYWQgYSB3ZWVrbHkgZXhwb3J0JywgJ3NlY3Rpb24tdGl0bGUnKTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQoaW50cm8pOwoKICAgIGNvbnN0IGRyb3B6b25lID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBkcm9wem9uZS5jbGFzc05hbWUgPSAnZHJvcHpvbmUnOwogICAgZHJvcHpvbmUuaWQgPSAnZHJvcHpvbmUnOwogICAgZHJvcHpvbmUuaW5uZXJIVE1MID0gYAogICAgICA8ZGl2IGNsYXNzPSJlbXB0eS1pY29uIiBzdHlsZT0ibWFyZ2luOiAwIGF1dG8gMTRweDsiPjxpIGRhdGEtbHVjaWRlPSJ1cGxvYWQtY2xvdWQiIHN0eWxlPSJ3aWR0aDoyMnB4O2hlaWdodDoyMnB4OyI+PC9pPjwvZGl2PgogICAgICA8aDM+RHJhZyAmYW1wOyBkcm9wIHlvdXIgLmNzdiBvciAueGxzeCBmaWxlIGhlcmU8L2gzPgogICAgICA8cD5vciBjbGljayB0byBicm93c2Ug4oCUIGZpbGVzIGFyZSB2YWxpZGF0ZWQgYmVmb3JlIGFueXRoaW5nIGlzIHNhdmVkPC9wPgogICAgICA8aW5wdXQgdHlwZT0iZmlsZSIgaWQ9ImZpbGVJbnB1dCIgYWNjZXB0PSIuY3N2LC54bHN4LC54bHMiIC8+CiAgICBgOwogICAgcm9vdC5hcHBlbmRDaGlsZChkcm9wem9uZSk7CgogICAgY29uc3QgcHJldmlld0FyZWEgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHByZXZpZXdBcmVhLmlkID0gJ3ByZXZpZXdBcmVhJzsKICAgIHJvb3QuYXBwZW5kQ2hpbGQocHJldmlld0FyZWEpOwoKICAgIHdpcmVEcm9wem9uZShkcm9wem9uZSk7CiAgfQoKICBmdW5jdGlvbiB3aXJlRHJvcHpvbmUoZHJvcHpvbmUpIHsKICAgIGNvbnN0IGlucHV0ID0gZHJvcHpvbmUucXVlcnlTZWxlY3RvcignI2ZpbGVJbnB1dCcpOwogICAgZHJvcHpvbmUuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBpbnB1dC5jbGljaygpKTsKICAgIGlucHV0LmFkZEV2ZW50TGlzdGVuZXIoJ2NoYW5nZScsICgpID0+IHsKICAgICAgaWYgKGlucHV0LmZpbGVzWzBdKSBoYW5kbGVGaWxlKGlucHV0LmZpbGVzWzBdKTsKICAgIH0pOwogICAgWydkcmFnZW50ZXInLCAnZHJhZ292ZXInXS5mb3JFYWNoKChldnQpID0+CiAgICAgIGRyb3B6b25lLmFkZEV2ZW50TGlzdGVuZXIoZXZ0LCAoZSkgPT4geyBlLnByZXZlbnREZWZhdWx0KCk7IGRyb3B6b25lLmNsYXNzTGlzdC5hZGQoJ2lzLWRyYWcnKTsgfSkKICAgICk7CiAgICBbJ2RyYWdsZWF2ZScsICdkcm9wJ10uZm9yRWFjaCgoZXZ0KSA9PgogICAgICBkcm9wem9uZS5hZGRFdmVudExpc3RlbmVyKGV2dCwgKGUpID0+IHsgZS5wcmV2ZW50RGVmYXVsdCgpOyBkcm9wem9uZS5jbGFzc0xpc3QucmVtb3ZlKCdpcy1kcmFnJyk7IH0pCiAgICApOwogICAgZHJvcHpvbmUuYWRkRXZlbnRMaXN0ZW5lcignZHJvcCcsIChlKSA9PiB7CiAgICAgIGNvbnN0IGZpbGUgPSBlLmRhdGFUcmFuc2Zlci5maWxlc1swXTsKICAgICAgaWYgKGZpbGUpIGhhbmRsZUZpbGUoZmlsZSk7CiAgICB9KTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIGhhbmRsZUZpbGUoZmlsZSkgewogICAgY29uc3QgYXJlYSA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdwcmV2aWV3QXJlYScpOwogICAgYXJlYS5pbm5lckhUTUwgPSAnJzsKICAgIGFyZWEuYXBwZW5kQ2hpbGQocm93V2l0aFNwaW5uZXIoJ1ZhbGlkYXRpbmcgZmlsZeKApicpKTsKICAgIE9iamVjdC5rZXlzKGR1cGxpY2F0ZUFjdGlvbk92ZXJyaWRlcykuZm9yRWFjaCgoaykgPT4gZGVsZXRlIGR1cGxpY2F0ZUFjdGlvbk92ZXJyaWRlc1trXSk7CiAgICB0cnkgewogICAgICBjdXJyZW50UHJldmlldyA9IGF3YWl0IEFwaS5wcmV2aWV3VXBsb2FkKGZpbGUpOwogICAgICByZW5kZXJQcmV2aWV3KCk7CiAgICB9IGNhdGNoIChlcnIpIHsKICAgICAgYXJlYS5pbm5lckhUTUwgPSAnJzsKICAgICAgYXJlYS5hcHBlbmRDaGlsZChlcnJvckJhbm5lcihlcnIubWVzc2FnZSkpOwogICAgfQogIH0KCiAgZnVuY3Rpb24gcm93V2l0aFNwaW5uZXIodGV4dCkgewogICAgY29uc3QgZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGVsLmNsYXNzTmFtZSA9ICdsb2FkaW5nLXJvdyc7CiAgICBjb25zdCBzcGlubmVyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgc3Bpbm5lci5jbGFzc05hbWUgPSAnc3Bpbm5lcic7CiAgICBlbC5hcHBlbmQoc3Bpbm5lciwgZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUoYCAke3RleHR9YCkpOwogICAgcmV0dXJuIGVsOwogIH0KICBmdW5jdGlvbiBlcnJvckJhbm5lcihtZXNzYWdlKSB7CiAgICBjb25zdCBlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZWwuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgZWwuc3R5bGUuYm9yZGVyTGVmdCA9ICczcHggc29saWQgdmFyKC0tc3RhdHVzLWNyaXRpY2FsKSc7CiAgICBlbC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIGBDb3VsZCBub3QgcmVhZCB0aGlzIGZpbGU6ICR7bWVzc2FnZX1gLCAnbXV0ZWQnKSk7CiAgICByZXR1cm4gZWw7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJQcmV2aWV3KCkgewogICAgY29uc3QgYXJlYSA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdwcmV2aWV3QXJlYScpOwogICAgYXJlYS5pbm5lckhUTUwgPSAnJzsKICAgIGNvbnN0IHAgPSBjdXJyZW50UHJldmlldzsKCiAgICBjb25zdCBzdW1tYXJ5VGl0bGUgPSB0ZXh0RWwoJ2RpdicsICdWYWxpZGF0aW9uIHN1bW1hcnknLCAnc2VjdGlvbi10aXRsZScpOwogICAgY29uc3Qgc3VtbWFyeUdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHN1bW1hcnlHcmlkLmNsYXNzTmFtZSA9ICdzdGF0LWdyaWQnOwogICAgc3VtbWFyeUdyaWQuYXBwZW5kKAogICAgICBzdGF0VGlsZSgnRmlsZScsIHAub3JpZ2luYWxOYW1lKSwKICAgICAgc3RhdFRpbGUoJ1NoZWV0cyBmb3VuZCcsIHAuc2hlZXRzLmxlbmd0aCksCiAgICAgIHN0YXRUaWxlKCdUb3RhbCByb3dzIChhbGwgc2hlZXRzKScsIHAudG90YWxEYXRhUm93cyksCiAgICAgIHN0YXRUaWxlKCdOZXcgcmVjb3JkcycsIHAubmV3UmVjb3Jkc0NvdW50KSwKICAgICAgc3RhdFRpbGUoJ0V4YWN0IGR1cGxpY2F0ZXMgZm91bmQnLCBwLmR1cGxpY2F0ZXMubGVuZ3RoKSwKICAgICAgc3RhdFRpbGUoJ0R1cGxpY2F0ZSByb3dzIGluIGZpbGUnLCBwLmR1cGxpY2F0ZVJvd3NJbkZpbGUpLAogICAgICBzdGF0VGlsZSgnUm93cyB3aXRoIGVycm9ycycsIHAuZXJyb3JSb3dzKQogICAgKTsKICAgIGFyZWEuYXBwZW5kKHN1bW1hcnlUaXRsZSwgc3VtbWFyeUdyaWQpOwoKICAgIGlmIChwLnNoZWV0cy5sZW5ndGgpIHsKICAgICAgY29uc3Qgc2hlZXRzVGl0bGUgPSB0ZXh0RWwoJ2RpdicsICdTaGVldCBicmVha2Rvd24nLCAnc2VjdGlvbi10aXRsZScpOwogICAgICBjb25zdCBzaGVldHNUYWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICAgIHNoZWV0c1RhYmxlLmNsYXNzTmFtZSA9ICdkYXRhLXRhYmxlJzsKICAgICAgc2hlZXRzVGFibGUuaW5uZXJIVE1MID0gJzx0aGVhZD48dHI+PHRoPlNoZWV0PC90aD48dGg+TGF5b3V0IGRldGVjdGVkPC90aD48dGggY2xhc3M9Im51bSI+Um93czwvdGg+PHRoIGNsYXNzPSJudW0iPlZhbGlkPC90aD48dGggY2xhc3M9Im51bSI+RXJyb3JzPC90aD48L3RyPjwvdGhlYWQ+JzsKICAgICAgY29uc3Qgc2hlZXRzQm9keSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3Rib2R5Jyk7CiAgICAgIHAuc2hlZXRzLmZvckVhY2goKHMpID0+IHsKICAgICAgICBjb25zdCB0ciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAgICAgY29uc3QgbGF5b3V0TGFiZWwgPSBzLmZvcm1hdCA9PT0gJ2FnZW5kYScgPyAnTFJTIGFnZW5kYSB0cmFja2VyJyA6IHMuZm9ybWF0ID09PSAnc2ltcGxlJyA/ICdTaW1wbGUgcGxhdGZvcm0gdGFibGUnIDogJ05vdCByZWNvZ25pemVkIOKAlCBzYXZlZCBhcyByYXcgZGF0YSBvbmx5JzsKICAgICAgICB0ci5hcHBlbmQoCiAgICAgICAgICB0ZXh0RWwoJ3RkJywgcy5uYW1lKSwKICAgICAgICAgIHRleHRFbCgndGQnLCBsYXlvdXRMYWJlbCksCiAgICAgICAgICB0ZXh0RWwoJ3RkJywgU3RyaW5nKHMudG90YWxSb3dzKSwgJ251bScpLAogICAgICAgICAgdGV4dEVsKCd0ZCcsIFN0cmluZyhzLnZhbGlkUm93cyksICdudW0nKSwKICAgICAgICAgIHRleHRFbCgndGQnLCBTdHJpbmcocy5lcnJvclJvd3MpLCAnbnVtJykKICAgICAgICApOwogICAgICAgIHNoZWV0c0JvZHkuYXBwZW5kQ2hpbGQodHIpOwogICAgICB9KTsKICAgICAgc2hlZXRzVGFibGUuYXBwZW5kQ2hpbGQoc2hlZXRzQm9keSk7CiAgICAgIGNvbnN0IHNoZWV0c1dyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgc2hlZXRzV3JhcC5jbGFzc05hbWUgPSAndGFibGUtc2Nyb2xsJzsKICAgICAgc2hlZXRzV3JhcC5hcHBlbmRDaGlsZChzaGVldHNUYWJsZSk7CiAgICAgIGFyZWEuYXBwZW5kKHNoZWV0c1RpdGxlLCBzaGVldHNXcmFwKTsKICAgIH0KCiAgICBpZiAocC5kdXBsaWNhdGVzLmxlbmd0aCkgewogICAgICBjb25zdCBkdXBUaXRsZSA9IHRleHRFbCgnZGl2JywgYEV4YWN0IGR1cGxpY2F0ZXMgZm91bmQgKCR7cC5kdXBsaWNhdGVzLmxlbmd0aH0pYCwgJ3NlY3Rpb24tdGl0bGUnKTsKICAgICAgYXJlYS5hcHBlbmRDaGlsZChkdXBUaXRsZSk7CiAgICAgIGFyZWEuYXBwZW5kQ2hpbGQodGV4dEVsKAogICAgICAgICdkaXYnLAogICAgICAgICdFYWNoIG9mIHRoZXNlIHJvd3MgaXMgYnl0ZS1mb3ItYnl0ZSBpZGVudGljYWwgdG8gYW4gYWxyZWFkeS1zYXZlZCByZWNvcmQg4oCUIGV2ZXJ5IGZpZWxkIG1hdGNoZXMsIGluY2x1ZGluZyBldmVyeSBtZXRyaWMsIG5vdCBqdXN0IHRoZSBkYXRlL2NhcHRpb24vcGxhdGZvcm0uIENob29zZSB3aGF0IHRvIGRvIHdpdGggZWFjaCDigJQgb3Igc2V0IGEgZGVmYXVsdCBmb3IgYWxsIG9mIHRoZW0uIChBIHJvdyB0aGF0IHNoYXJlcyB0aGUgc2FtZSBkYXRlL2NhcHRpb24vcGxhdGZvcm0gYnV0IGhhcyBkaWZmZXJlbnQgbnVtYmVycyBpcyBub3Qgc2hvd24gaGVyZSDigJQgaXTigJlzIGltcG9ydGVkIGF1dG9tYXRpY2FsbHkgYXMgaXRzIG93biBuZXcgcmVjb3JkLCBzaW5jZSBpdHMgYW5hbHl0aWNzIGNoYW5nZWQuKScsCiAgICAgICAgJ211dGVkJwogICAgICApKTsKICAgICAgY29uc3QgZGVmYXVsdFJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBkZWZhdWx0Um93LmNsYXNzTmFtZSA9ICdmaWVsZC1pbmxpbmUnOwogICAgICBkZWZhdWx0Um93LnN0eWxlLm1hcmdpbiA9ICcxMHB4IDAnOwogICAgICBjb25zdCBkZWZhdWx0U2VsZWN0ID0gYWN0aW9uU2VsZWN0KCdza2lwJyk7CiAgICAgIGRlZmF1bHRTZWxlY3QuaWQgPSAnZGVmYXVsdER1cGxpY2F0ZUFjdGlvblNlbGVjdCc7CiAgICAgIGRlZmF1bHRTZWxlY3QuYWRkRXZlbnRMaXN0ZW5lcignY2hhbmdlJywgKCkgPT4gewogICAgICAgIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJy5jb25mbGljdC1yb3cgc2VsZWN0W2RhdGEtaGFzaF0nKS5mb3JFYWNoKChzZWwpID0+IHsKICAgICAgICAgIGlmICghZHVwbGljYXRlQWN0aW9uT3ZlcnJpZGVzW3NlbC5kYXRhc2V0Lmhhc2hdKSBzZWwudmFsdWUgPSBkZWZhdWx0U2VsZWN0LnZhbHVlOwogICAgICAgIH0pOwogICAgICB9KTsKICAgICAgZGVmYXVsdFJvdy5hcHBlbmQodGV4dEVsKCdsYWJlbCcsICdEZWZhdWx0IGFjdGlvbiBmb3IgYWxsIG1hdGNoZXMnKSwgZGVmYXVsdFNlbGVjdCk7CiAgICAgIGFyZWEuYXBwZW5kQ2hpbGQoZGVmYXVsdFJvdyk7CgogICAgICBjb25zdCBsaXN0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIGxpc3QuY2xhc3NOYW1lID0gJ2NvbmZsaWN0LWxpc3QnOwogICAgICBwLmR1cGxpY2F0ZXMuZm9yRWFjaCgoZCkgPT4gewogICAgICAgIGNvbnN0IHJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICAgIHJvdy5jbGFzc05hbWUgPSAnY29uZmxpY3Qtcm93JzsKICAgICAgICBjb25zdCBsZWZ0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgICAgbGVmdC5hcHBlbmQoCiAgICAgICAgICB0ZXh0RWwoJ2RpdicsIGAke0Zvcm1hdC5kYXRlKGQucHVibGlzaERhdGUpfSDigJQgJHsoZC5jYXB0aW9uIHx8ICcobm8gY2FwdGlvbiknKS5zbGljZSgwLCA3MCl9YCwgJ3dlZWstbGFiZWwnKSwKICAgICAgICAgIHRleHRFbCgnZGl2JywgYEV4YWN0IG1hdGNoIG9mIGV4aXN0aW5nIHJlY29yZCAjJHtkLmV4aXN0aW5nLnBvc3RJZH0gKGxhc3QgdXBkYXRlZCAke2QuZXhpc3RpbmcudXBkYXRlZEF0fSlgLCAnd2Vlay1tZXRhJykKICAgICAgICApOwogICAgICAgIHJvdy5hcHBlbmRDaGlsZChsZWZ0KTsKICAgICAgICBjb25zdCBzZWwgPSBhY3Rpb25TZWxlY3QoJ3NraXAnKTsKICAgICAgICBzZWwuZGF0YXNldC5oYXNoID0gZC5oYXNoOwogICAgICAgIHNlbC5hZGRFdmVudExpc3RlbmVyKCdjaGFuZ2UnLCAoKSA9PiB7IGR1cGxpY2F0ZUFjdGlvbk92ZXJyaWRlc1tkLmhhc2hdID0gc2VsLnZhbHVlOyB9KTsKICAgICAgICByb3cuYXBwZW5kQ2hpbGQoc2VsKTsKICAgICAgICBsaXN0LmFwcGVuZENoaWxkKHJvdyk7CiAgICAgIH0pOwogICAgICBhcmVhLmFwcGVuZENoaWxkKGxpc3QpOwogICAgfQoKICAgIGNvbnN0IG5vdGVzRmllbGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIG5vdGVzRmllbGQuY2xhc3NOYW1lID0gJ2Zvcm0tZmllbGQnOwogICAgbm90ZXNGaWVsZC5zdHlsZS5tYXJnaW4gPSAnMTJweCAwJzsKICAgIG5vdGVzRmllbGQuYXBwZW5kQ2hpbGQodGV4dEVsKCdsYWJlbCcsICdVcGxvYWQgbm90ZXMgKG9wdGlvbmFsKScpKTsKICAgIGNvbnN0IG5vdGVzSW5wdXQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOwogICAgbm90ZXNJbnB1dC50eXBlID0gJ3RleHQnOwogICAgbm90ZXNJbnB1dC5pZCA9ICd1cGxvYWROb3Rlc0lucHV0JzsKICAgIG5vdGVzSW5wdXQucGxhY2Vob2xkZXIgPSAnZS5nLiAiV2VlayAzIGV4cG9ydCwgaW5jbHVkZXMgY29ycmVjdGVkIFRpa1RvayBudW1iZXJzIic7CiAgICBub3Rlc0ZpZWxkLmFwcGVuZENoaWxkKG5vdGVzSW5wdXQpOwogICAgYXJlYS5hcHBlbmRDaGlsZChub3Rlc0ZpZWxkKTsKCiAgICBpZiAocC5pc3N1ZXMubGVuZ3RoKSB7CiAgICAgIGNvbnN0IGlzc3Vlc1RpdGxlID0gdGV4dEVsKCdkaXYnLCBgUm93cyBza2lwcGVkIG9yIGZsYWdnZWQgKCR7cC5pc3N1ZXMubGVuZ3RofSlgLCAnc2VjdGlvbi10aXRsZScpOwogICAgICBjb25zdCBpc3N1ZXNDYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIGlzc3Vlc0NhcmQuY2xhc3NOYW1lID0gJ2lzc3Vlcy1saXN0JzsKICAgICAgcC5pc3N1ZXMuZm9yRWFjaCgoaXNzdWUpID0+IHsKICAgICAgICBjb25zdCByb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgICByb3cuY2xhc3NOYW1lID0gJ2lzc3VlLXJvdyc7CiAgICAgICAgaWYgKGlzc3VlLnJvd051bWJlcikgcm93LmFwcGVuZENoaWxkKHRleHRFbCgnc3BhbicsIGBSb3cgJHtpc3N1ZS5yb3dOdW1iZXJ9YCwgJ3Jvdy1ubycpKTsKICAgICAgICByb3cuYXBwZW5kQ2hpbGQoZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUoaXNzdWUubWVzc2FnZSkpOwogICAgICAgIGlzc3Vlc0NhcmQuYXBwZW5kQ2hpbGQocm93KTsKICAgICAgfSk7CiAgICAgIGFyZWEuYXBwZW5kKGlzc3Vlc1RpdGxlLCBpc3N1ZXNDYXJkKTsKICAgIH0KCiAgICBpZiAocC5uZXdSZWNvcmRzLmxlbmd0aCkgewogICAgICBjb25zdCBuZXdUaXRsZSA9IHRleHRFbCgnZGl2JywgYE5ldyByZWNvcmRzIHRvIGltcG9ydCAoJHtwLm5ld1JlY29yZHMubGVuZ3RofSlgLCAnc2VjdGlvbi10aXRsZScpOwogICAgICBhcmVhLmFwcGVuZENoaWxkKG5ld1RpdGxlKTsKICAgICAgYXJlYS5hcHBlbmRDaGlsZCh0ZXh0RWwoCiAgICAgICAgJ2RpdicsCiAgICAgICAgJ1RoZXNlIHJvd3MgZG9u4oCZdCBtYXRjaCBhbnl0aGluZyBhbHJlYWR5IHNhdmVkLCBzbyB0aGV54oCZbGwgYmUgaW1wb3J0ZWQgYXV0b21hdGljYWxseSDigJQgbm8gZGVjaXNpb24gbmVlZGVkLCB1bmxpa2UgdGhlIGV4YWN0LWR1cGxpY2F0ZSBtYXRjaGVzIGFib3ZlLicsCiAgICAgICAgJ211dGVkJwogICAgICApKTsKICAgICAgY29uc3QgdGFibGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0YWJsZScpOwogICAgICB0YWJsZS5jbGFzc05hbWUgPSAnZGF0YS10YWJsZSc7CiAgICAgIHRhYmxlLmlubmVySFRNTCA9ICc8dGhlYWQ+PHRyPjx0aD5EYXRlPC90aD48dGg+Q2FwdGlvbjwvdGg+PHRoPlR5cGU8L3RoPjx0aD5DYW1wYWlnbjwvdGg+PHRoPlBsYXRmb3JtczwvdGg+PC90cj48L3RoZWFkPic7CiAgICAgIGNvbnN0IHRib2R5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGJvZHknKTsKICAgICAgcC5uZXdSZWNvcmRzLmZvckVhY2goKHMpID0+IHsKICAgICAgICBjb25zdCB0ciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAgICAgdHIuYXBwZW5kKAogICAgICAgICAgdGV4dEVsKCd0ZCcsIEZvcm1hdC5kYXRlKHMucHVibGlzaERhdGUpKSwKICAgICAgICAgIHRleHRFbCgndGQnLCBzLmNhcHRpb24gfHwgJ+KAlCcpLAogICAgICAgICAgdGV4dEVsKCd0ZCcsIHMuY29udGVudFR5cGUgfHwgJ+KAlCcpLAogICAgICAgICAgdGV4dEVsKCd0ZCcsIHMuY2FtcGFpZ25UeXBlIHx8ICdVbnNwZWNpZmllZCcpLAogICAgICAgICAgdGV4dEVsKCd0ZCcsIHMucGxhdGZvcm1zLmpvaW4oJywgJykpCiAgICAgICAgKTsKICAgICAgICB0Ym9keS5hcHBlbmRDaGlsZCh0cik7CiAgICAgIH0pOwogICAgICB0YWJsZS5hcHBlbmRDaGlsZCh0Ym9keSk7CiAgICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgd3JhcC5jbGFzc05hbWUgPSAndGFibGUtc2Nyb2xsJzsKICAgICAgd3JhcC5hcHBlbmRDaGlsZCh0YWJsZSk7CiAgICAgIGFyZWEuYXBwZW5kKHdyYXApOwogICAgfQoKICAgIGNvbnN0IGFjdGlvbnMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGFjdGlvbnMuY2xhc3NOYW1lID0gJ2J0bi1yb3cnOwogICAgYWN0aW9ucy5zdHlsZS5tYXJnaW5Ub3AgPSAnMTZweCc7CiAgICBjb25zdCBjb21taXRCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIGNvbW1pdEJ0bi5jbGFzc05hbWUgPSAnYnRuIHByaW1hcnknOwogICAgY29tbWl0QnRuLnRleHRDb250ZW50ID0gcC52YWxpZFJvd3MgPiAwID8gYEltcG9ydCAke3AudmFsaWRSb3dzfSByb3cocylgIDogJ05vdGhpbmcgdG8gaW1wb3J0JzsKICAgIGNvbW1pdEJ0bi5kaXNhYmxlZCA9IHAudmFsaWRSb3dzID09PSAwOwogICAgY29tbWl0QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gY29tbWl0KGNvbW1pdEJ0bikpOwogICAgY29uc3QgY2FuY2VsQnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBjYW5jZWxCdG4uY2xhc3NOYW1lID0gJ2J0bic7CiAgICBjYW5jZWxCdG4udGV4dENvbnRlbnQgPSAnQ2FuY2VsJzsKICAgIGNhbmNlbEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsgY3VycmVudFByZXZpZXcgPSBudWxsOyBzaGVsbCgpOyB9KTsKICAgIGFjdGlvbnMuYXBwZW5kKGNvbW1pdEJ0biwgY2FuY2VsQnRuKTsKICAgIGFyZWEuYXBwZW5kQ2hpbGQoYWN0aW9ucyk7CiAgfQoKICBmdW5jdGlvbiBzdGF0VGlsZShsYWJlbCwgdmFsdWUpIHsKICAgIGNvbnN0IHRpbGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRpbGUuY2xhc3NOYW1lID0gJ3N0YXQtdGlsZSc7CiAgICB0aWxlLmFwcGVuZCh0ZXh0RWwoJ2RpdicsIGxhYmVsLCAnc3RhdC1sYWJlbCcpLCB0ZXh0RWwoJ2RpdicsIFN0cmluZyh2YWx1ZSksICdzdGF0LXZhbHVlJykpOwogICAgcmV0dXJuIHRpbGU7CiAgfQogIGZ1bmN0aW9uIGFjdGlvblNlbGVjdChkZWZhdWx0VmFsKSB7CiAgICBjb25zdCBzZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzZWxlY3QnKTsKICAgIFtbJ3NraXAnLCAnU2tpcCAoa2VlcCBleGlzdGluZyByZWNvcmQgdW5jaGFuZ2VkKSddLCBbJ3VwZGF0ZScsICdVcGRhdGUgZXhpc3RpbmcgcmVjb3JkJ10sIFsnY3JlYXRlJywgJ0NyZWF0ZSBhcyBhIG5ldywgc2VwYXJhdGUgcmVjb3JkJ11dLmZvckVhY2goKFt2LCBsXSkgPT4gewogICAgICBjb25zdCBvcHQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsgb3B0LnZhbHVlID0gdjsgb3B0LnRleHRDb250ZW50ID0gbDsKICAgICAgaWYgKHYgPT09IGRlZmF1bHRWYWwpIG9wdC5zZWxlY3RlZCA9IHRydWU7CiAgICAgIHNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgICByZXR1cm4gc2VsOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gY29tbWl0KGJ0bikgewogICAgYnRuLmRpc2FibGVkID0gdHJ1ZTsKICAgIGJ0bi50ZXh0Q29udGVudCA9ICdJbXBvcnRpbmfigKYnOwogICAgY29uc3QgZGVmYXVsdER1cGxpY2F0ZUFjdGlvbiA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdkZWZhdWx0RHVwbGljYXRlQWN0aW9uU2VsZWN0Jyk/LnZhbHVlIHx8ICdza2lwJzsKICAgIGNvbnN0IG5vdGVzID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3VwbG9hZE5vdGVzSW5wdXQnKT8udmFsdWUgfHwgbnVsbDsKICAgIHRyeSB7CiAgICAgIGNvbnN0IHJlc3VsdCA9IGF3YWl0IEFwaS5jb21taXRVcGxvYWQoewogICAgICAgIGZpbGVQYXRoOiBjdXJyZW50UHJldmlldy5maWxlUGF0aCwKICAgICAgICBvcmlnaW5hbE5hbWU6IGN1cnJlbnRQcmV2aWV3Lm9yaWdpbmFsTmFtZSwKICAgICAgICBkZWZhdWx0RHVwbGljYXRlQWN0aW9uLAogICAgICAgIGR1cGxpY2F0ZUFjdGlvbnM6IGR1cGxpY2F0ZUFjdGlvbk92ZXJyaWRlcywKICAgICAgICBub3RlcywKICAgICAgfSk7CiAgICAgIFRvYXN0LnNob3coCiAgICAgICAgYEltcG9ydGVkOiAke3Jlc3VsdC5pbXBvcnRlZFJvd3N9IG5ldywgJHtyZXN1bHQudXBkYXRlZFJvd3N9IHVwZGF0ZWQsICR7cmVzdWx0LnNraXBwZWRSb3dzfSBza2lwcGVkLmAsCiAgICAgICAgcmVzdWx0LmVycm9yQ291bnQgPiAwID8gJ2Vycm9yJyA6ICdzdWNjZXNzJwogICAgICApOwogICAgICBjdXJyZW50UHJldmlldyA9IG51bGw7CiAgICAgIHNoZWxsKCk7CiAgICAgIHdpbmRvdy5kaXNwYXRjaEV2ZW50KG5ldyBDdXN0b21FdmVudCgnbHJzOmRhdGEtdXBkYXRlZCcpKTsKICAgIH0gY2F0Y2ggKGVycikgewogICAgICBUb2FzdC5zaG93KGVyci5tZXNzYWdlLCAnZXJyb3InKTsKICAgICAgYnRuLmRpc2FibGVkID0gZmFsc2U7CiAgICAgIGJ0bi50ZXh0Q29udGVudCA9ICdSZXRyeSBpbXBvcnQnOwogICAgfQogIH0KCiAgZnVuY3Rpb24gcmVuZGVyKCkgewogICAgcm9vdCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd2aWV3LXVwbG9hZCcpOwogICAgc2hlbGwoKTsKICB9CgogIHJldHVybiB7IHJlbmRlciB9Owp9KSgpOwoKY29uc3QgSGlzdG9yeSA9ICgoKSA9PiB7CiAgbGV0IHJvb3Q7CiAgbGV0IGN1cnJlbnRVcGxvYWRzID0gW107CiAgbGV0IHNlYXJjaFZhbHVlID0gJyc7CiAgbGV0IHBhZ2UgPSAxOwogIGNvbnN0IHBhZ2VTaXplID0gMTU7CiAgbGV0IHNvcnRTdGF0ZSA9IHsga2V5OiAndXBsb2FkZWRfYXQnLCBkaXI6ICdkZXNjJywgdHlwZTogJ3N0cmluZycgfTsKICBjb25zdCBFWFBPUlRfQ09MVU1OUyA9IFsKICAgIHsga2V5OiAnZmlsZW5hbWUnLCBsYWJlbDogJ0ZpbGUnIH0sCiAgICB7IGtleTogJ3VwbG9hZGVkX2F0JywgbGFiZWw6ICdVcGxvYWRlZCcgfSwKICAgIHsga2V5OiAnc3RhdHVzJywgbGFiZWw6ICdTdGF0dXMnIH0sCiAgICB7IGtleTogJ2ltcG9ydGVkX3Jvd3MnLCBsYWJlbDogJ0ltcG9ydGVkJyB9LAogICAgeyBrZXk6ICd1cGRhdGVkX3Jvd3MnLCBsYWJlbDogJ1VwZGF0ZWQnIH0sCiAgICB7IGtleTogJ3NraXBwZWRfcm93cycsIGxhYmVsOiAnU2tpcHBlZCcgfSwKICAgIHsga2V5OiAnZXJyb3JfY291bnQnLCBsYWJlbDogJ0Vycm9ycycgfSwKICAgIHsga2V5OiAnd2Vla3MnLCBsYWJlbDogJ1dlZWtzJyB9LAogICAgeyBrZXk6ICdub3RlcycsIGxhYmVsOiAnTm90ZXMnIH0sCiAgXTsKCiAgZnVuY3Rpb24gYmFkZ2VDbGFzcyhzdGF0dXMpIHsKICAgIGlmIChzdGF0dXMgPT09ICdzdWNjZXNzJykgcmV0dXJuICdzdWNjZXNzJzsKICAgIGlmIChzdGF0dXMgPT09ICdwYXJ0aWFsJykgcmV0dXJuICdwYXJ0aWFsJzsKICAgIHJldHVybiAnZmFpbGVkJzsKICB9CgogIGZ1bmN0aW9uIHNvcnRhYmxlSGVhZGVyKGxhYmVsLCBrZXksIHR5cGUpIHsKICAgIGNvbnN0IHRoID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGgnKTsKICAgIGlmICh0eXBlID09PSAnbnVtYmVyJykgdGguY2xhc3NOYW1lID0gJ251bSc7CiAgICB0aC5jbGFzc0xpc3QuYWRkKCdzb3J0YWJsZS10aCcpOwogICAgY29uc3QgaXNBY3RpdmUgPSBzb3J0U3RhdGUua2V5ID09PSBrZXk7CiAgICB0aC5hcHBlbmRDaGlsZChkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShsYWJlbCkpOwogICAgdGguYXBwZW5kQ2hpbGQodGV4dEVsKCdzcGFuJywgaXNBY3RpdmUgPyAoc29ydFN0YXRlLmRpciA9PT0gJ2FzYycgPyAnIOKGkScgOiAnIOKGkycpIDogJyDihpUnLCAnc29ydC1hcnJvdycpKTsKICAgIHRoLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gewogICAgICBzb3J0U3RhdGUgPSB7IGtleSwgZGlyOiBzb3J0U3RhdGUua2V5ID09PSBrZXkgJiYgc29ydFN0YXRlLmRpciA9PT0gJ2FzYycgPyAnZGVzYycgOiAnYXNjJywgdHlwZSB9OwogICAgICByZW5kZXJMaXN0KCk7CiAgICB9KTsKICAgIHJldHVybiB0aDsKICB9CgogIGZ1bmN0aW9uIGZpbHRlcmVkVXBsb2FkcygpIHsKICAgIGNvbnN0IHEgPSBzZWFyY2hWYWx1ZS50cmltKCkudG9Mb3dlckNhc2UoKTsKICAgIGlmICghcSkgcmV0dXJuIGN1cnJlbnRVcGxvYWRzOwogICAgcmV0dXJuIGN1cnJlbnRVcGxvYWRzLmZpbHRlcigodSkgPT4gKAogICAgICB1LmZpbGVuYW1lLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkKICAgICAgfHwgKHUubm90ZXMgfHwgJycpLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkKICAgICAgfHwgdS5zdGF0dXMudG9Mb3dlckNhc2UoKS5pbmNsdWRlcyhxKQogICAgKSk7CiAgfQoKICBmdW5jdGlvbiBzb3J0ZWRVcGxvYWRzKCkgewogICAgY29uc3QgeyBrZXksIGRpciwgdHlwZSB9ID0gc29ydFN0YXRlOwogICAgY29uc3QgZmFjdG9yID0gZGlyID09PSAnYXNjJyA/IDEgOiAtMTsKICAgIHJldHVybiBbLi4uZmlsdGVyZWRVcGxvYWRzKCldLnNvcnQoKGEsIGIpID0+IHsKICAgICAgY29uc3QgYXYgPSBhW2tleV07CiAgICAgIGNvbnN0IGJ2ID0gYltrZXldOwogICAgICBpZiAoYXYgPT09IG51bGwgfHwgYXYgPT09IHVuZGVmaW5lZCkgcmV0dXJuIDE7CiAgICAgIGlmIChidiA9PT0gbnVsbCB8fCBidiA9PT0gdW5kZWZpbmVkKSByZXR1cm4gLTE7CiAgICAgIGlmICh0eXBlID09PSAnbnVtYmVyJykgcmV0dXJuIChhdiAtIGJ2KSAqIGZhY3RvcjsKICAgICAgcmV0dXJuIFN0cmluZyhhdikubG9jYWxlQ29tcGFyZShTdHJpbmcoYnYpKSAqIGZhY3RvcjsKICAgIH0pOwogIH0KCiAgZnVuY3Rpb24gZXhwb3J0Um93cygpIHsKICAgIHJldHVybiBzb3J0ZWRVcGxvYWRzKCkubWFwKCh1KSA9PiAoewogICAgICBmaWxlbmFtZTogdS5maWxlbmFtZSwKICAgICAgdXBsb2FkZWRfYXQ6IHUudXBsb2FkZWRfYXQsCiAgICAgIHN0YXR1czogdS5zdGF0dXMsCiAgICAgIGltcG9ydGVkX3Jvd3M6IHUuaW1wb3J0ZWRfcm93cywKICAgICAgdXBkYXRlZF9yb3dzOiB1LnVwZGF0ZWRfcm93cywKICAgICAgc2tpcHBlZF9yb3dzOiB1LnNraXBwZWRfcm93cywKICAgICAgZXJyb3JfY291bnQ6IHUuZXJyb3JfY291bnQsCiAgICAgIHdlZWtzOiB1LndlZWtzX2FmZmVjdGVkLm1hcCgodykgPT4gRm9ybWF0LmRhdGUodykpLmpvaW4oJywgJyksCiAgICAgIG5vdGVzOiB1Lm5vdGVzIHx8ICcnLAogICAgfSkpOwogIH0KCiAgZnVuY3Rpb24gYnVpbGRCYWNrdXBDYXJkKCkgewogICAgY29uc3QgY2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY2FyZC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICBjYXJkLnN0eWxlLm1hcmdpbkJvdHRvbSA9ICcyMHB4JzsKICAgIGNvbnN0IGhlYWRlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgaGVhZGVyLmNsYXNzTmFtZSA9ICdjYXJkLWhlYWRlcic7CiAgICBoZWFkZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMycsICdCYWNrdXAgJiBSZXN0b3JlJykpOwogICAgY2FyZC5hcHBlbmRDaGlsZChoZWFkZXIpOwogICAgY2FyZC5hcHBlbmRDaGlsZCh0ZXh0RWwoCiAgICAgICdkaXYnLAogICAgICAnRG93bmxvYWQgYSBmdWxsIHNuYXBzaG90IG9mIHRoZSBkYXRhYmFzZSBhbnkgdGltZS4gUmVzdG9yaW5nIHJlcGxhY2VzIEFMTCBjdXJyZW50IGRhdGEgd2l0aCB0aGUgdXBsb2FkZWQgYmFja3VwIGFuZCByZXN0YXJ0cyB0aGUgc2VydmVyLiAiRGVsZXRlIGFsbCB1cGxvYWRlZCByZWNvcmRzIiBjbGVhcnMgZXZlcnkgcmVjb3JkIGFuZCB0aGUgdXBsb2FkIGhpc3RvcnkgKGJ1dCBub3QgRm9sbG93ZXJzIERhdGEgUmVjb3JkKSBzbyB5b3UgY2FuIHJlYnVpbGQgZnJvbSBhIGZyZXNoIGltcG9ydC4gQm90aCBhY3Rpb25zIGNhbm5vdCBiZSB1bmRvbmUg4oCUIGRvd25sb2FkIGEgYmFja3VwIGZpcnN0LicsCiAgICAgICdtdXRlZCcKICAgICkpOwoKICAgIGNvbnN0IGFjdGlvbnMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGFjdGlvbnMuY2xhc3NOYW1lID0gJ2J0bi1yb3cnOwogICAgYWN0aW9ucy5zdHlsZS5tYXJnaW5Ub3AgPSAnMTRweCc7CgogICAgY29uc3QgZG93bmxvYWRCdG4gPSBpY29uQnRuKCdidG4gcHJpbWFyeScsICdkb3dubG9hZCcsICdEb3dubG9hZCBCYWNrdXAnKTsKICAgIGRvd25sb2FkQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4geyB3aW5kb3cubG9jYXRpb24uaHJlZiA9ICcvYXBpL2JhY2t1cC9leHBvcnQnOyB9KTsKCiAgICBjb25zdCByZXN0b3JlSW5wdXQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOwogICAgcmVzdG9yZUlucHV0LnR5cGUgPSAnZmlsZSc7CiAgICByZXN0b3JlSW5wdXQuYWNjZXB0ID0gJy5kYic7CiAgICByZXN0b3JlSW5wdXQuc3R5bGUuZGlzcGxheSA9ICdub25lJzsKCiAgICBjb25zdCByZXN0b3JlQnRuID0gaWNvbkJ0bignYnRuIGRhbmdlcicsICd1cGxvYWQnLCAnUmVzdG9yZSBmcm9tIEJhY2t1cCcpOwogICAgcmVzdG9yZUJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHJlc3RvcmVJbnB1dC5jbGljaygpKTsKCiAgICByZXN0b3JlSW5wdXQuYWRkRXZlbnRMaXN0ZW5lcignY2hhbmdlJywgYXN5bmMgKCkgPT4gewogICAgICBjb25zdCBmaWxlID0gcmVzdG9yZUlucHV0LmZpbGVzWzBdOwogICAgICBpZiAoIWZpbGUpIHJldHVybjsKICAgICAgY29uc3Qgc3VyZSA9IHdpbmRvdy5jb25maXJtKAogICAgICAgICdSZXN0b3Jpbmcgd2lsbCBSRVBMQUNFIGFsbCBjdXJyZW50IGRhdGEgd2l0aCB0aGlzIGJhY2t1cCBmaWxlIGFuZCByZXN0YXJ0IHRoZSBzZXJ2ZXIuIFRoaXMgY2Fubm90IGJlIHVuZG9uZS4gQ29udGludWU/JwogICAgICApOwogICAgICBpZiAoIXN1cmUpIHsKICAgICAgICByZXN0b3JlSW5wdXQudmFsdWUgPSAnJzsKICAgICAgICByZXR1cm47CiAgICAgIH0KICAgICAgcmVzdG9yZUJ0bi5kaXNhYmxlZCA9IHRydWU7CiAgICAgIHRyeSB7CiAgICAgICAgY29uc3QgZm9ybSA9IG5ldyBGb3JtRGF0YSgpOwogICAgICAgIGZvcm0uYXBwZW5kKCdmaWxlJywgZmlsZSk7CiAgICAgICAgY29uc3QgcmVzdWx0ID0gYXdhaXQgQXBpLnJlc3RvcmVCYWNrdXAoZm9ybSk7CiAgICAgICAgVG9hc3Quc2hvdyhyZXN1bHQubWVzc2FnZSB8fCAnQmFja3VwIHJlc3RvcmVkLiBUaGUgc2VydmVyIGlzIHJlc3RhcnRpbmcuJywgJ3N1Y2Nlc3MnKTsKICAgICAgfSBjYXRjaCAoZXJyKSB7CiAgICAgICAgVG9hc3Quc2hvdyhlcnIubWVzc2FnZSwgJ2Vycm9yJyk7CiAgICAgICAgcmVzdG9yZUJ0bi5kaXNhYmxlZCA9IGZhbHNlOwogICAgICB9IGZpbmFsbHkgewogICAgICAgIHJlc3RvcmVJbnB1dC52YWx1ZSA9ICcnOwogICAgICB9CiAgICB9KTsKCiAgICBjb25zdCB3aXBlQnRuID0gaWNvbkJ0bignYnRuIGRhbmdlcicsICd0cmFzaC0yJywgJ0RlbGV0ZSBhbGwgdXBsb2FkZWQgcmVjb3JkcycpOwogICAgd2lwZUJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIGFzeW5jICgpID0+IHsKICAgICAgY29uc3QgdHlwZWQgPSB3aW5kb3cucHJvbXB0KAogICAgICAgICdUaGlzIHBlcm1hbmVudGx5IGRlbGV0ZXMgRVZFUlkgdXBsb2FkZWQgcmVjb3JkIGFuZCBhbGwgdXBsb2FkIGhpc3RvcnksIHNvIHlvdSBjYW4gcmVidWlsZCB0aGUgZGF0YSBmcm9tIGEgZnJlc2ggaW1wb3J0LiBGb2xsb3dlcnMgRGF0YSBSZWNvcmQgaXMgbm90IGFmZmVjdGVkLiBUaGlzIGNhbm5vdCBiZSB1bmRvbmUuXG5cblR5cGUgREVMRVRFIHRvIGNvbmZpcm06JwogICAgICApOwogICAgICBpZiAodHlwZWQgIT09ICdERUxFVEUnKSB7CiAgICAgICAgaWYgKHR5cGVkICE9PSBudWxsKSBUb2FzdC5zaG93KCdDb25maXJtYXRpb24gcGhyYXNlIGRpZCBub3QgbWF0Y2gg4oCUIG5vdGhpbmcgd2FzIGRlbGV0ZWQuJywgJ2Vycm9yJyk7CiAgICAgICAgcmV0dXJuOwogICAgICB9CiAgICAgIHdpcGVCdG4uZGlzYWJsZWQgPSB0cnVlOwogICAgICB0cnkgewogICAgICAgIGNvbnN0IHJlc3VsdCA9IGF3YWl0IEFwaS53aXBlVXBsb2FkZWRSZWNvcmRzKCk7CiAgICAgICAgY29uc3QgbiA9IHJlc3VsdC5kZWxldGVkID8gcmVzdWx0LmRlbGV0ZWQucG9zdHMgOiAwOwogICAgICAgIFRvYXN0LnNob3coYERlbGV0ZWQgJHtufSByZWNvcmQocykuIFVwbG9hZCBhIGZpbGUgdG8gbG9hZCBmcmVzaCBkYXRhLmAsICdzdWNjZXNzJyk7CiAgICAgICAgd2luZG93LmRpc3BhdGNoRXZlbnQobmV3IEN1c3RvbUV2ZW50KCdscnM6ZGF0YS11cGRhdGVkJykpOwogICAgICAgIHJlbmRlcigpOwogICAgICB9IGNhdGNoIChlcnIpIHsKICAgICAgICBUb2FzdC5zaG93KGVyci5tZXNzYWdlIHx8ICdEZWxldGUgZmFpbGVkLicsICdlcnJvcicpOwogICAgICB9IGZpbmFsbHkgewogICAgICAgIHdpcGVCdG4uZGlzYWJsZWQgPSBmYWxzZTsKICAgICAgfQogICAgfSk7CgogICAgYWN0aW9ucy5hcHBlbmQoZG93bmxvYWRCdG4sIHJlc3RvcmVCdG4sIHJlc3RvcmVJbnB1dCwgd2lwZUJ0bik7CiAgICBjYXJkLmFwcGVuZENoaWxkKGFjdGlvbnMpOwogICAgcmV0dXJuIGNhcmQ7CiAgfQoKICBhc3luYyBmdW5jdGlvbiByZW5kZXIoKSB7CiAgICByb290ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3ZpZXctaGlzdG9yeScpOwogICAgcm9vdC5pbm5lckhUTUwgPSAnJzsKICAgIHJvb3QuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCAnVXBsb2FkIGhpc3RvcnknLCAnc2VjdGlvbi10aXRsZScpKTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQoYnVpbGRCYWNrdXBDYXJkKCkpOwoKICAgIGN1cnJlbnRVcGxvYWRzID0gYXdhaXQgQXBpLnVwbG9hZEhpc3RvcnkoKTsKICAgIGlmICghY3VycmVudFVwbG9hZHMubGVuZ3RoKSB7CiAgICAgIHJvb3QuYXBwZW5kQ2hpbGQoZW1wdHlTdGF0ZSh7CiAgICAgICAgaWNvbjogJ3VwbG9hZC1jbG91ZCcsCiAgICAgICAgdGl0bGU6ICdObyB1cGxvYWRzIHlldCcsCiAgICAgICAgbWVzc2FnZTogJ0ltcG9ydCB5b3VyIGZpcnN0IHdlZWtseSBleHBvcnQgdG8gc3RhcnQgc2VlaW5nIGRhdGEgYWNyb3NzIHRoZSBhcHAuJywKICAgICAgICBhY3Rpb25MYWJlbDogJ1VwbG9hZCBkYXRhJywKICAgICAgICBvbkFjdGlvbjogKCkgPT4gZG9jdW1lbnQucXVlcnlTZWxlY3RvcignLnRhYi1idG5bZGF0YS10YWI9InVwbG9hZCJdJyk/LmNsaWNrKCksCiAgICAgIH0pKTsKICAgICAgcmV0dXJuOwogICAgfQoKICAgIGNvbnN0IHRvb2xiYXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRvb2xiYXIuY2xhc3NOYW1lID0gJ3JlY29yZHMtdG9vbGJhcic7CiAgICB0b29sYmFyLmFwcGVuZENoaWxkKGJ1aWxkU2VhcmNoQm94KHsKICAgICAgcGxhY2Vob2xkZXI6ICdTZWFyY2ggZmlsZW5hbWUsIHN0YXR1cywgb3Igbm90ZXPigKYnLAogICAgICB2YWx1ZTogc2VhcmNoVmFsdWUsCiAgICAgIG9uQ2hhbmdlOiAodikgPT4geyBzZWFyY2hWYWx1ZSA9IHY7IHBhZ2UgPSAxOyByZW5kZXJMaXN0KCk7IH0sCiAgICB9KSk7CiAgICByb290LmFwcGVuZENoaWxkKHRvb2xiYXIpOwoKICAgIHJvb3QuYXBwZW5kQ2hpbGQoYnVpbGRFeHBvcnRCdXR0b25zKHsKICAgICAgZ2V0Um93c0FuZENvbHVtbnM6ICgpID0+ICh7IHJvd3M6IGV4cG9ydFJvd3MoKSwgY29sdW1uczogRVhQT1JUX0NPTFVNTlMgfSksCiAgICAgIGZpbGVuYW1lQmFzZTogJ3VwbG9hZC1oaXN0b3J5JywKICAgICAgc2hlZXROYW1lOiAnVXBsb2FkIEhpc3RvcnknLAogICAgfSkpOwoKICAgIGNvbnN0IGNhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgd3JhcC5jbGFzc05hbWUgPSAndGFibGUtc2Nyb2xsJzsKICAgIHdyYXAuaWQgPSAnaGlzdG9yeVRhYmxlV3JhcCc7CiAgICBjYXJkLmFwcGVuZENoaWxkKHdyYXApOwogICAgY29uc3QgcGFnZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHBhZ2VyLmNsYXNzTmFtZSA9ICdwYWdpbmF0aW9uLXJvdyc7CiAgICBwYWdlci5pZCA9ICdoaXN0b3J5UGFnZXInOwogICAgY2FyZC5hcHBlbmRDaGlsZChwYWdlcik7CiAgICByb290LmFwcGVuZENoaWxkKGNhcmQpOwoKICAgIHJlbmRlckxpc3QoKTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlckxpc3QoKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2hpc3RvcnlUYWJsZVdyYXAnKTsKICAgIGNvbnN0IHBhZ2VyRWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnaGlzdG9yeVBhZ2VyJyk7CiAgICBpZiAoIXdyYXApIHJldHVybjsKICAgIGNvbnN0IGFsbFNvcnRlZCA9IHNvcnRlZFVwbG9hZHMoKTsKICAgIGlmICghYWxsU29ydGVkLmxlbmd0aCkgewogICAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgICB3cmFwLmFwcGVuZENoaWxkKGVtcHR5U3RhdGUoeyBpY29uOiAndXBsb2FkLWNsb3VkJywgbWVzc2FnZTogJ05vIHVwbG9hZHMgbWF0Y2ggeW91ciBzZWFyY2guJyB9KSk7CiAgICAgIGlmIChwYWdlckVsKSBwYWdlckVsLmlubmVySFRNTCA9ICcnOwogICAgICByZXR1cm47CiAgICB9CiAgICBjb25zdCB7IHBhZ2VSb3dzLCB0b3RhbFBhZ2VzLCBzYWZlUGFnZSwgdG90YWwgfSA9IHBhZ2luYXRlQ2xpZW50U2lkZShhbGxTb3J0ZWQsIHBhZ2UsIHBhZ2VTaXplKTsKICAgIHBhZ2UgPSBzYWZlUGFnZTsKCiAgICBjb25zdCB0YWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICB0YWJsZS5jbGFzc05hbWUgPSAnZGF0YS10YWJsZSc7CiAgICBjb25zdCB0aGVhZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RoZWFkJyk7CiAgICBjb25zdCBoZWFkVHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgaGVhZFRyLmFwcGVuZCgKICAgICAgc29ydGFibGVIZWFkZXIoJ0ZpbGUnLCAnZmlsZW5hbWUnLCAnc3RyaW5nJyksCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdVcGxvYWRlZCcsICd1cGxvYWRlZF9hdCcsICdzdHJpbmcnKSwKICAgICAgc29ydGFibGVIZWFkZXIoJ1N0YXR1cycsICdzdGF0dXMnLCAnc3RyaW5nJyksCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdJbXBvcnRlZCcsICdpbXBvcnRlZF9yb3dzJywgJ251bWJlcicpLAogICAgICBzb3J0YWJsZUhlYWRlcignVXBkYXRlZCcsICd1cGRhdGVkX3Jvd3MnLCAnbnVtYmVyJyksCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdTa2lwcGVkJywgJ3NraXBwZWRfcm93cycsICdudW1iZXInKSwKICAgICAgc29ydGFibGVIZWFkZXIoJ0Vycm9ycycsICdlcnJvcl9jb3VudCcsICdudW1iZXInKSwKICAgICAgdGV4dEVsKCd0aCcsICdXZWVrcycpLAogICAgICB0ZXh0RWwoJ3RoJywgJ05vdGVzJykKICAgICk7CiAgICB0aGVhZC5hcHBlbmRDaGlsZChoZWFkVHIpOwogICAgY29uc3QgdGJvZHkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0Ym9keScpOwogICAgcGFnZVJvd3MuZm9yRWFjaCgodSkgPT4gewogICAgICBjb25zdCB0ciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAgIHRyLnN0eWxlLmN1cnNvciA9ICdwb2ludGVyJzsKICAgICAgY29uc3QgYmFkZ2UgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICAgIGJhZGdlLmNsYXNzTmFtZSA9IGBiYWRnZSAke2JhZGdlQ2xhc3ModS5zdGF0dXMpfWA7CiAgICAgIGJhZGdlLnRleHRDb250ZW50ID0gdS5zdGF0dXM7CiAgICAgIGNvbnN0IHN0YXR1c1RkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgc3RhdHVzVGQuYXBwZW5kQ2hpbGQoYmFkZ2UpOwogICAgICB0ci5hcHBlbmQoCiAgICAgICAgdGV4dEVsKCd0ZCcsIHUuZmlsZW5hbWUpLAogICAgICAgIHRleHRFbCgndGQnLCB1LnVwbG9hZGVkX2F0KSwKICAgICAgICBzdGF0dXNUZCwKICAgICAgICB0ZXh0RWwoJ3RkJywgU3RyaW5nKHUuaW1wb3J0ZWRfcm93cyksICdudW0nKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgU3RyaW5nKHUudXBkYXRlZF9yb3dzKSwgJ251bScpLAogICAgICAgIHRleHRFbCgndGQnLCBTdHJpbmcodS5za2lwcGVkX3Jvd3MpLCAnbnVtJyksCiAgICAgICAgdGV4dEVsKCd0ZCcsIFN0cmluZyh1LmVycm9yX2NvdW50KSwgJ251bScpLAogICAgICAgIHRleHRFbCgndGQnLCB1LndlZWtzX2FmZmVjdGVkLm1hcCgodykgPT4gRm9ybWF0LmRhdGUodykpLmpvaW4oJywgJykgfHwgJ+KAlCcpLAogICAgICAgIHRleHRFbCgndGQnLCB1Lm5vdGVzIHx8ICfigJQnKQogICAgICApOwogICAgICB0ci5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHRvZ2dsZUVycm9ycyh1LmlkLCB0cikpOwogICAgICB0Ym9keS5hcHBlbmRDaGlsZCh0cik7CiAgICB9KTsKICAgIHRhYmxlLmFwcGVuZCh0aGVhZCwgdGJvZHkpOwogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGFibGUpOwoKICAgIGlmIChwYWdlckVsKSB7CiAgICAgIHBhZ2VyRWwuaW5uZXJIVE1MID0gJyc7CiAgICAgIHBhZ2VyRWwuYXBwZW5kQ2hpbGQoYnVpbGRQYWdlcih7CiAgICAgICAgcGFnZTogc2FmZVBhZ2UsCiAgICAgICAgdG90YWxQYWdlcywKICAgICAgICB0b3RhbCwKICAgICAgICBvblByZXY6ICgpID0+IHsgcGFnZSAtPSAxOyByZW5kZXJMaXN0KCk7IH0sCiAgICAgICAgb25OZXh0OiAoKSA9PiB7IHBhZ2UgKz0gMTsgcmVuZGVyTGlzdCgpOyB9LAogICAgICB9KSk7CiAgICB9CiAgfQoKICBhc3luYyBmdW5jdGlvbiB0b2dnbGVFcnJvcnModXBsb2FkSWQsIHRyKSB7CiAgICBjb25zdCBleGlzdGluZyA9IHRyLm5leHRFbGVtZW50U2libGluZzsKICAgIGlmIChleGlzdGluZyAmJiBleGlzdGluZy5jbGFzc0xpc3QuY29udGFpbnMoJ2Vycm9yLWxvZy1yb3cnKSkgewogICAgICBleGlzdGluZy5yZW1vdmUoKTsKICAgICAgcmV0dXJuOwogICAgfQogICAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnLmVycm9yLWxvZy1yb3cnKS5mb3JFYWNoKChlbCkgPT4gZWwucmVtb3ZlKCkpOwogICAgY29uc3QgZXJyb3JzID0gYXdhaXQgQXBpLnVwbG9hZEVycm9ycyh1cGxvYWRJZCk7CiAgICBjb25zdCBsb2dSb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgbG9nUm93LmNsYXNzTmFtZSA9ICdlcnJvci1sb2ctcm93JzsKICAgIGNvbnN0IHRkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgIHRkLmNvbFNwYW4gPSA5OwogICAgaWYgKCFlcnJvcnMubGVuZ3RoKSB7CiAgICAgIHRkLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgJ05vIGlzc3VlcyBsb2dnZWQgZm9yIHRoaXMgdXBsb2FkLicsICdtdXRlZCcpKTsKICAgIH0gZWxzZSB7CiAgICAgIGNvbnN0IGxpc3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgbGlzdC5jbGFzc05hbWUgPSAnaXNzdWVzLWxpc3QnOwogICAgICBlcnJvcnMuZm9yRWFjaCgoZSkgPT4gewogICAgICAgIGNvbnN0IHJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICAgIHJvdy5jbGFzc05hbWUgPSAnaXNzdWUtcm93JzsKICAgICAgICBjb25zdCBiYWRnZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgICAgICBiYWRnZS5jbGFzc05hbWUgPSBgYmFkZ2UgJHtlLnNldmVyaXR5fS1zZXZgOwogICAgICAgIGJhZGdlLnRleHRDb250ZW50ID0gZS5zZXZlcml0eTsKICAgICAgICByb3cuYXBwZW5kKGJhZGdlLCBkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShgICR7ZS5yb3dfbnVtYmVyID8gYFJvdyAke2Uucm93X251bWJlcn06IGAgOiAnJ30ke2UubWVzc2FnZX1gKSk7CiAgICAgICAgbGlzdC5hcHBlbmRDaGlsZChyb3cpOwogICAgICB9KTsKICAgICAgdGQuYXBwZW5kQ2hpbGQobGlzdCk7CiAgICB9CgogICAgY29uc3QgcmF3QnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICByYXdCdG4uY2xhc3NOYW1lID0gJ2J0bic7CiAgICByYXdCdG4uc3R5bGUubWFyZ2luVG9wID0gJzEwcHgnOwogICAgcmF3QnRuLnRleHRDb250ZW50ID0gJ1ZpZXcgZXZlcnkgcmF3IHNvdXJjZSByb3cgZnJvbSB0aGlzIHVwbG9hZCc7CiAgICByYXdCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBsb2FkUmF3Um93cyh1cGxvYWRJZCwgcmF3QnRuKSk7CiAgICB0ZC5hcHBlbmRDaGlsZChyYXdCdG4pOwogICAgY29uc3QgcmF3V3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcmF3V3JhcC5pZCA9IGByYXdXcmFwLSR7dXBsb2FkSWR9YDsKICAgIHRkLmFwcGVuZENoaWxkKHJhd1dyYXApOwoKICAgIGxvZ1Jvdy5hcHBlbmRDaGlsZCh0ZCk7CiAgICB0ci5hZnRlcihsb2dSb3cpOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gbG9hZFJhd1Jvd3ModXBsb2FkSWQsIGJ0bikgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGByYXdXcmFwLSR7dXBsb2FkSWR9YCk7CiAgICBpZiAoIXdyYXApIHJldHVybjsKICAgIGlmICh3cmFwLmRhdGFzZXQubG9hZGVkKSB7CiAgICAgIHdyYXAuc3R5bGUuZGlzcGxheSA9IHdyYXAuc3R5bGUuZGlzcGxheSA9PT0gJ25vbmUnID8gJ2Jsb2NrJyA6ICdub25lJzsKICAgICAgcmV0dXJuOwogICAgfQogICAgYnRuLnRleHRDb250ZW50ID0gJ0xvYWRpbmfigKYnOwogICAgY29uc3QgeyByb3dzLCB0b3RhbCB9ID0gYXdhaXQgQXBpLnVwbG9hZFJhd1Jvd3ModXBsb2FkSWQpOwogICAgd3JhcC5kYXRhc2V0LmxvYWRlZCA9ICcxJzsKICAgIGJ0bi50ZXh0Q29udGVudCA9IGBTaG93aW5nICR7cm93cy5sZW5ndGh9IG9mICR7dG90YWx9IHJhdyByb3cocylgOwoKICAgIGNvbnN0IGJ5U2hlZXQgPSBuZXcgTWFwKCk7CiAgICByb3dzLmZvckVhY2goKHIpID0+IHsKICAgICAgaWYgKCFieVNoZWV0LmhhcyhyLnNoZWV0X25hbWUpKSBieVNoZWV0LnNldChyLnNoZWV0X25hbWUsIFtdKTsKICAgICAgYnlTaGVldC5nZXQoci5zaGVldF9uYW1lKS5wdXNoKHIpOwogICAgfSk7CgogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgIHdyYXAuc3R5bGUubWFyZ2luVG9wID0gJzEwcHgnOwogICAgYnlTaGVldC5mb3JFYWNoKChzaGVldFJvd3MsIHNoZWV0TmFtZSkgPT4gewogICAgICB3cmFwLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgYFNoZWV0OiAke3NoZWV0TmFtZX0gKCR7c2hlZXRSb3dzLmxlbmd0aH0gcm93KHMpKWAsICdzdGF0LWxhYmVsJykpOwogICAgICBjb25zdCBoZWFkZXJzID0gc2hlZXRSb3dzWzBdLmhlYWRlcnM7CiAgICAgIGNvbnN0IHRhYmxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGFibGUnKTsKICAgICAgdGFibGUuY2xhc3NOYW1lID0gJ2RhdGEtdGFibGUnOwogICAgICBjb25zdCB0aGVhZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAgIHRoZWFkLmFwcGVuZCh0ZXh0RWwoJ3RoJywgJ1JvdyAjJyksIHRleHRFbCgndGgnLCAnTGlua2VkIHRvIHBvc3QnKSk7CiAgICAgIGNvbnN0IGNvbENvdW50ID0gaGVhZGVycyA/IGhlYWRlcnMubGVuZ3RoIDogTWF0aC5tYXgoLi4uc2hlZXRSb3dzLm1hcCgocikgPT4gci5yYXcubGVuZ3RoKSk7CiAgICAgIGZvciAobGV0IGkgPSAwOyBpIDwgY29sQ291bnQ7IGkgKz0gMSkgdGhlYWQuYXBwZW5kQ2hpbGQodGV4dEVsKCd0aCcsIGhlYWRlcnMgJiYgaGVhZGVyc1tpXSA/IFN0cmluZyhoZWFkZXJzW2ldKSA6IGBDb2wgJHtpICsgMX1gKSk7CiAgICAgIGNvbnN0IHRoZWFkV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RoZWFkJyk7CiAgICAgIHRoZWFkV3JhcC5hcHBlbmRDaGlsZCh0aGVhZCk7CiAgICAgIGNvbnN0IHRib2R5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGJvZHknKTsKICAgICAgc2hlZXRSb3dzLmZvckVhY2goKHIpID0+IHsKICAgICAgICBjb25zdCB0cjIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgICAgIHRyMi5hcHBlbmQodGV4dEVsKCd0ZCcsIFN0cmluZyhyLnJvd19udW1iZXIpKSwgdGV4dEVsKCd0ZCcsIHIucG9zdF9pZCA/IGAjJHtyLnBvc3RfaWR9YCA6ICfigJQnKSk7CiAgICAgICAgZm9yIChsZXQgaSA9IDA7IGkgPCBjb2xDb3VudDsgaSArPSAxKSB7CiAgICAgICAgICBjb25zdCB2YWwgPSByLnJhd1tpXTsKICAgICAgICAgIHRyMi5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3RkJywgdmFsID09PSB1bmRlZmluZWQgfHwgdmFsID09PSBudWxsID8gJycgOiBTdHJpbmcodmFsKS5zbGljZSgwLCA2MCkpKTsKICAgICAgICB9CiAgICAgICAgdGJvZHkuYXBwZW5kQ2hpbGQodHIyKTsKICAgICAgfSk7CiAgICAgIHRhYmxlLmFwcGVuZCh0aGVhZFdyYXAsIHRib2R5KTsKICAgICAgY29uc3Qgc2Nyb2xsV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBzY3JvbGxXcmFwLmNsYXNzTmFtZSA9ICd0YWJsZS1zY3JvbGwnOwogICAgICBzY3JvbGxXcmFwLnN0eWxlLm1hcmdpbkJvdHRvbSA9ICcxNnB4JzsKICAgICAgc2Nyb2xsV3JhcC5hcHBlbmRDaGlsZCh0YWJsZSk7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQoc2Nyb2xsV3JhcCk7CiAgICB9KTsKICB9CgogIHJldHVybiB7IHJlbmRlciB9Owp9KSgpOwoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIEZvbGxvd2VycyBEYXRhIHRhYjogbWFudWFsIHdlZWtseSBmb2xsb3dlci1jb3VudCBlbnRyeSBwZXIKICAgcGxhdGZvcm0g4oCUIGVudGlyZWx5IGluZGVwZW5kZW50IG9mIHNwcmVhZHNoZWV0IHVwbG9hZHMgKGl0cyBvd24KICAgdGFibGUsIGl0cyBvd24gQVBJLCBuZXZlciB0b3VjaGVkIGJ5IHRoZSBpbXBvcnQgcGlwZWxpbmUpLiBQb3dlcnMKICAgRm9sbG93ZXIgR3Jvd3RoIGNoYXJ0cy9jb21wYXJpc29ucyBlbHNld2hlcmUgaW4gdGhlIGFwcC4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmNvbnN0IEZvbGxvd2VycyA9ICgoKSA9PiB7CiAgbGV0IHJvb3Q7CiAgbGV0IGVkaXRpbmdJZCA9IG51bGw7IC8vIG5vbi1udWxsIHdoaWxlIHRoZSBmb3JtIGlzIGVkaXRpbmcgYW4gZXhpc3RpbmcgZW50cnkgcmF0aGVyIHRoYW4gYWRkaW5nIGEgbmV3IG9uZQogIGxldCBzb3J0U3RhdGUgPSB7IGtleTogJ2VudHJ5X2RhdGUnLCBkaXI6ICdkZXNjJywgdHlwZTogJ3N0cmluZycgfTsKICBsZXQgY3VycmVudFJvd3MgPSBbXTsKICBsZXQgc2VhcmNoVmFsdWUgPSAnJzsKICBsZXQgcGFnZSA9IDE7CiAgY29uc3QgcGFnZVNpemUgPSAxMDsKICBjb25zdCBFWFBPUlRfQ09MVU1OUyA9IFsKICAgIHsga2V5OiAncGxhdGZvcm1fbGFiZWwnLCBsYWJlbDogJ1BsYXRmb3JtJyB9LAogICAgeyBrZXk6ICdlbnRyeV9kYXRlJywgbGFiZWw6ICdXZWVrIC8gRGF0ZScgfSwKICAgIHsga2V5OiAnZm9sbG93ZXJzX2NvdW50JywgbGFiZWw6ICdGb2xsb3dlcnMgQ291bnQnIH0sCiAgICB7IGtleTogJ3VwZGF0ZWRfYXQnLCBsYWJlbDogJ0xhc3QgVXBkYXRlZCcgfSwKICBdOwoKICBmdW5jdGlvbiBhbGxQbGF0Zm9ybXMoKSB7CiAgICByZXR1cm4gKHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZSB8fCB7IGFsbFBsYXRmb3JtczogW10gfSkuYWxsUGxhdGZvcm1zIHx8IFtdOwogIH0KCiAgZnVuY3Rpb24gcGxhdGZvcm1NZXRhRm9yKGlkKSB7CiAgICByZXR1cm4gYWxsUGxhdGZvcm1zKCkuZmluZCgocCkgPT4gcC5pZCA9PT0gaWQpIHx8IHsgbGFiZWw6IGlkLCBjb2xvcjogJyM5OTknIH07CiAgfQoKICBmdW5jdGlvbiBzaGVsbCgpIHsKICAgIHJvb3QuaW5uZXJIVE1MID0gJyc7CiAgICByb290LmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgJ0ZvbGxvd2VycyBEYXRhIFJlY29yZCcsICdzZWN0aW9uLXRpdGxlJykpOwogICAgcm9vdC5hcHBlbmRDaGlsZCh0ZXh0RWwoCiAgICAgICdkaXYnLAogICAgICAnTWFudWFsbHkgbG9nIGVhY2ggcGxhdGZvcm3igJlzIHRvdGFsIGZvbGxvd2VyIGNvdW50IG9uY2UgYSB3ZWVrLiBUaGlzIGlzIGluZGVwZW5kZW50IG9mIHNwcmVhZHNoZWV0IHVwbG9hZHMg4oCUIGl0IHBvd2VycyBGb2xsb3dlciBHcm93dGggY2hhcnRzIGFuZCBjb21wYXJpc29ucyBlbHNld2hlcmUgaW4gdGhlIGFwcC4nLAogICAgICAnbXV0ZWQnCiAgICApKTsKCiAgICBjb25zdCBmb3JtQ2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZm9ybUNhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgZm9ybUNhcmQuc3R5bGUubWFyZ2luQm90dG9tID0gJzIwcHgnOwogICAgZm9ybUNhcmQuaWQgPSAnZm9sbG93ZXJzRm9ybUNhcmQnOwogICAgcm9vdC5hcHBlbmRDaGlsZChmb3JtQ2FyZCk7CgogICAgY29uc3QgdG9vbGJhciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdG9vbGJhci5jbGFzc05hbWUgPSAncmVjb3Jkcy10b29sYmFyJzsKICAgIHRvb2xiYXIuYXBwZW5kQ2hpbGQoYnVpbGRTZWFyY2hCb3goewogICAgICBwbGFjZWhvbGRlcjogJ1NlYXJjaCBwbGF0Zm9ybSBvciBkYXRl4oCmJywKICAgICAgdmFsdWU6IHNlYXJjaFZhbHVlLAogICAgICBvbkNoYW5nZTogKHYpID0+IHsgc2VhcmNoVmFsdWUgPSB2OyBwYWdlID0gMTsgcmVuZGVyVGFibGUoKTsgfSwKICAgIH0pKTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQodG9vbGJhcik7CgogICAgcm9vdC5hcHBlbmRDaGlsZChidWlsZEV4cG9ydEJ1dHRvbnMoewogICAgICBnZXRSb3dzQW5kQ29sdW1uczogKCkgPT4gKHsgcm93czogZXhwb3J0Um93cygpLCBjb2x1bW5zOiBFWFBPUlRfQ09MVU1OUyB9KSwKICAgICAgZmlsZW5hbWVCYXNlOiAnZm9sbG93ZXJzLWRhdGEnLAogICAgICBzaGVldE5hbWU6ICdGb2xsb3dlcnMgRGF0YScsCiAgICB9KSk7CgogICAgY29uc3QgdGFibGVDYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0YWJsZUNhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgY29uc3QgdGFibGVXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0YWJsZVdyYXAuY2xhc3NOYW1lID0gJ3RhYmxlLXNjcm9sbCc7CiAgICB0YWJsZVdyYXAuaWQgPSAnZm9sbG93ZXJzVGFibGVXcmFwJzsKICAgIHRhYmxlQ2FyZC5hcHBlbmRDaGlsZCh0YWJsZVdyYXApOwogICAgY29uc3QgcGFnZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHBhZ2VyLmNsYXNzTmFtZSA9ICdwYWdpbmF0aW9uLXJvdyc7CiAgICBwYWdlci5pZCA9ICdmb2xsb3dlcnNQYWdlcic7CiAgICB0YWJsZUNhcmQuYXBwZW5kQ2hpbGQocGFnZXIpOwogICAgcm9vdC5hcHBlbmRDaGlsZCh0YWJsZUNhcmQpOwoKICAgIHJlbmRlckZvcm0oKTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlckZvcm0oKSB7CiAgICBjb25zdCBjYXJkID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZvbGxvd2Vyc0Zvcm1DYXJkJyk7CiAgICBpZiAoIWNhcmQpIHJldHVybjsKICAgIGNhcmQuaW5uZXJIVE1MID0gJyc7CiAgICBjb25zdCBoZWFkZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGhlYWRlci5jbGFzc05hbWUgPSAnY2FyZC1oZWFkZXInOwogICAgaGVhZGVyLmFwcGVuZENoaWxkKHRleHRFbCgnaDMnLCBlZGl0aW5nSWQgIT09IG51bGwgPyAnRWRpdCBlbnRyeScgOiAnQWRkIGEgd2Vla2x5IGVudHJ5JykpOwogICAgY2FyZC5hcHBlbmRDaGlsZChoZWFkZXIpOwoKICAgIGNvbnN0IGdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGdyaWQuY2xhc3NOYW1lID0gJ2Zvcm0tZ3JpZCc7CgogICAgY29uc3QgcGxhdGZvcm1GaWVsZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcGxhdGZvcm1GaWVsZC5jbGFzc05hbWUgPSAnZm9ybS1maWVsZCc7CiAgICBwbGF0Zm9ybUZpZWxkLmFwcGVuZENoaWxkKHRleHRFbCgnbGFiZWwnLCAnUGxhdGZvcm0nKSk7CiAgICBjb25zdCBwbGF0Zm9ybVNlbGVjdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NlbGVjdCcpOwogICAgcGxhdGZvcm1TZWxlY3QuaWQgPSAnZm9sbG93ZXJzUGxhdGZvcm1JbnB1dCc7CiAgICBhbGxQbGF0Zm9ybXMoKS5mb3JFYWNoKChwKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOwogICAgICBvcHQudmFsdWUgPSBwLmlkOwogICAgICBvcHQudGV4dENvbnRlbnQgPSBwLmxhYmVsOwogICAgICBwbGF0Zm9ybVNlbGVjdC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgICBwbGF0Zm9ybUZpZWxkLmFwcGVuZENoaWxkKHBsYXRmb3JtU2VsZWN0KTsKCiAgICBjb25zdCBkYXRlRmllbGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGRhdGVGaWVsZC5jbGFzc05hbWUgPSAnZm9ybS1maWVsZCc7CiAgICBkYXRlRmllbGQuYXBwZW5kQ2hpbGQodGV4dEVsKCdsYWJlbCcsICdXZWVrIC8gRGF0ZScpKTsKICAgIGNvbnN0IGRhdGVJbnB1dCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7CiAgICBkYXRlSW5wdXQudHlwZSA9ICdkYXRlJzsKICAgIGRhdGVJbnB1dC5pZCA9ICdmb2xsb3dlcnNEYXRlSW5wdXQnOwogICAgZGF0ZUZpZWxkLmFwcGVuZENoaWxkKGRhdGVJbnB1dCk7CgogICAgY29uc3QgY291bnRGaWVsZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY291bnRGaWVsZC5jbGFzc05hbWUgPSAnZm9ybS1maWVsZCc7CiAgICBjb3VudEZpZWxkLmFwcGVuZENoaWxkKHRleHRFbCgnbGFiZWwnLCAnRm9sbG93ZXJzIENvdW50JykpOwogICAgY29uc3QgY291bnRJbnB1dCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7CiAgICBjb3VudElucHV0LnR5cGUgPSAnbnVtYmVyJzsKICAgIGNvdW50SW5wdXQubWluID0gJzAnOwogICAgY291bnRJbnB1dC5zdGVwID0gJzEnOwogICAgY291bnRJbnB1dC5pZCA9ICdmb2xsb3dlcnNDb3VudElucHV0JzsKICAgIGNvdW50RmllbGQuYXBwZW5kQ2hpbGQoY291bnRJbnB1dCk7CgogICAgZ3JpZC5hcHBlbmQocGxhdGZvcm1GaWVsZCwgZGF0ZUZpZWxkLCBjb3VudEZpZWxkKTsKICAgIGNhcmQuYXBwZW5kQ2hpbGQoZ3JpZCk7CgogICAgY29uc3QgZWRpdFJvdyA9IGVkaXRpbmdJZCAhPT0gbnVsbCA/IGN1cnJlbnRSb3dzLmZpbmQoKHIpID0+IHIuaWQgPT09IGVkaXRpbmdJZCkgOiBudWxsOwogICAgaWYgKGVkaXRSb3cpIHsKICAgICAgcGxhdGZvcm1TZWxlY3QudmFsdWUgPSBlZGl0Um93LnBsYXRmb3JtOwogICAgICBkYXRlSW5wdXQudmFsdWUgPSBlZGl0Um93LmVudHJ5X2RhdGU7CiAgICAgIGNvdW50SW5wdXQudmFsdWUgPSBTdHJpbmcoZWRpdFJvdy5mb2xsb3dlcnNfY291bnQpOwogICAgfSBlbHNlIHsKICAgICAgZGF0ZUlucHV0LnZhbHVlID0gbmV3IERhdGUoKS50b0lTT1N0cmluZygpLnNsaWNlKDAsIDEwKTsKICAgIH0KCiAgICBjb25zdCBhY3Rpb25zID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBhY3Rpb25zLmNsYXNzTmFtZSA9ICdtb2RhbC1hY3Rpb25zJzsKICAgIGNvbnN0IGVycm9yRWwgPSB0ZXh0RWwoJ3NwYW4nLCAnJywgJ211dGVkJyk7CiAgICBlcnJvckVsLmlkID0gJ2ZvbGxvd2Vyc0Zvcm1FcnJvcic7CiAgICBlcnJvckVsLnN0eWxlLmNvbG9yID0gJ3ZhcigtLXN0YXR1cy1jcml0aWNhbCknOwoKICAgIGNvbnN0IGJ0blJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYnRuUm93LmNsYXNzTmFtZSA9ICdidG4tcm93JzsKICAgIGNvbnN0IHNhdmVCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIHNhdmVCdG4uY2xhc3NOYW1lID0gJ2J0biBwcmltYXJ5JzsKICAgIHNhdmVCdG4udGV4dENvbnRlbnQgPSBlZGl0aW5nSWQgIT09IG51bGwgPyAnU2F2ZSBjaGFuZ2VzJyA6ICdBZGQgZW50cnknOwogICAgc2F2ZUJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHN1Ym1pdEZvcm0oc2F2ZUJ0bikpOwogICAgYnRuUm93LmFwcGVuZENoaWxkKHNhdmVCdG4pOwogICAgaWYgKGVkaXRpbmdJZCAhPT0gbnVsbCkgewogICAgICBjb25zdCBjYW5jZWxCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgICAgY2FuY2VsQnRuLmNsYXNzTmFtZSA9ICdidG4nOwogICAgICBjYW5jZWxCdG4udGV4dENvbnRlbnQgPSAnQ2FuY2VsJzsKICAgICAgY2FuY2VsQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4geyBlZGl0aW5nSWQgPSBudWxsOyByZW5kZXJGb3JtKCk7IH0pOwogICAgICBidG5Sb3cuYXBwZW5kQ2hpbGQoY2FuY2VsQnRuKTsKICAgIH0KICAgIGFjdGlvbnMuYXBwZW5kKGVycm9yRWwsIGJ0blJvdyk7CiAgICBjYXJkLmFwcGVuZENoaWxkKGFjdGlvbnMpOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gc3VibWl0Rm9ybShidG4pIHsKICAgIGNvbnN0IGVycm9yRWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZm9sbG93ZXJzRm9ybUVycm9yJyk7CiAgICBlcnJvckVsLnRleHRDb250ZW50ID0gJyc7CiAgICBjb25zdCBwbGF0Zm9ybSA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmb2xsb3dlcnNQbGF0Zm9ybUlucHV0JykudmFsdWU7CiAgICBjb25zdCBlbnRyeURhdGUgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZm9sbG93ZXJzRGF0ZUlucHV0JykudmFsdWU7CiAgICBjb25zdCBmb2xsb3dlcnNDb3VudCA9IE51bWJlcihkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZm9sbG93ZXJzQ291bnRJbnB1dCcpLnZhbHVlKTsKICAgIGJ0bi5kaXNhYmxlZCA9IHRydWU7CiAgICB0cnkgewogICAgICBpZiAoZWRpdGluZ0lkICE9PSBudWxsKSB7CiAgICAgICAgYXdhaXQgQXBpLnVwZGF0ZUZvbGxvd2VycyhlZGl0aW5nSWQsIHsgcGxhdGZvcm0sIGVudHJ5RGF0ZSwgZm9sbG93ZXJzQ291bnQgfSk7CiAgICAgICAgVG9hc3Quc2hvdygnRW50cnkgdXBkYXRlZC4nLCAnc3VjY2VzcycpOwogICAgICB9IGVsc2UgewogICAgICAgIGF3YWl0IEFwaS5zYXZlRm9sbG93ZXJzKHsgcGxhdGZvcm0sIGVudHJ5RGF0ZSwgZm9sbG93ZXJzQ291bnQgfSk7CiAgICAgICAgVG9hc3Quc2hvdygnRW50cnkgc2F2ZWQuJywgJ3N1Y2Nlc3MnKTsKICAgICAgfQogICAgICBlZGl0aW5nSWQgPSBudWxsOwogICAgICBhd2FpdCBsb2FkKCk7CiAgICAgIHdpbmRvdy5kaXNwYXRjaEV2ZW50KG5ldyBDdXN0b21FdmVudCgnbHJzOmRhdGEtdXBkYXRlZCcpKTsKICAgIH0gY2F0Y2ggKGVycikgewogICAgICBlcnJvckVsLnRleHRDb250ZW50ID0gZXJyLm1lc3NhZ2U7CiAgICAgIGJ0bi5kaXNhYmxlZCA9IGZhbHNlOwogICAgfQogIH0KCiAgZnVuY3Rpb24gc3RhcnRFZGl0KHJvdykgewogICAgZWRpdGluZ0lkID0gcm93LmlkOwogICAgcmVuZGVyRm9ybSgpOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZvbGxvd2Vyc0Zvcm1DYXJkJykuc2Nyb2xsSW50b1ZpZXcoeyBiZWhhdmlvcjogJ3Ntb290aCcsIGJsb2NrOiAnc3RhcnQnIH0pOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gaGFuZGxlRGVsZXRlKHJvdykgewogICAgY29uc3Qgc3VyZSA9IHdpbmRvdy5jb25maXJtKGBEZWxldGUgdGhlICR7cGxhdGZvcm1NZXRhRm9yKHJvdy5wbGF0Zm9ybSkubGFiZWx9IGVudHJ5IGZvciAke0Zvcm1hdC5kYXRlKHJvdy5lbnRyeV9kYXRlKX0/YCk7CiAgICBpZiAoIXN1cmUpIHJldHVybjsKICAgIHRyeSB7CiAgICAgIGF3YWl0IEFwaS5kZWxldGVGb2xsb3dlcnMocm93LmlkKTsKICAgICAgVG9hc3Quc2hvdygnRW50cnkgZGVsZXRlZC4nLCAnc3VjY2VzcycpOwogICAgICBpZiAoZWRpdGluZ0lkID09PSByb3cuaWQpIGVkaXRpbmdJZCA9IG51bGw7CiAgICAgIGF3YWl0IGxvYWQoKTsKICAgICAgd2luZG93LmRpc3BhdGNoRXZlbnQobmV3IEN1c3RvbUV2ZW50KCdscnM6ZGF0YS11cGRhdGVkJykpOwogICAgfSBjYXRjaCAoZXJyKSB7CiAgICAgIFRvYXN0LnNob3coZXJyLm1lc3NhZ2UsICdlcnJvcicpOwogICAgfQogIH0KCiAgZnVuY3Rpb24gc29ydGFibGVIZWFkZXIobGFiZWwsIGtleSwgdHlwZSkgewogICAgY29uc3QgdGggPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0aCcpOwogICAgaWYgKHR5cGUgPT09ICdudW1iZXInKSB0aC5jbGFzc05hbWUgPSAnbnVtJzsKICAgIHRoLmNsYXNzTGlzdC5hZGQoJ3NvcnRhYmxlLXRoJyk7CiAgICBjb25zdCBpc0FjdGl2ZSA9IHNvcnRTdGF0ZS5rZXkgPT09IGtleTsKICAgIHRoLmFwcGVuZENoaWxkKGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKGxhYmVsKSk7CiAgICB0aC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3NwYW4nLCBpc0FjdGl2ZSA/IChzb3J0U3RhdGUuZGlyID09PSAnYXNjJyA/ICcg4oaRJyA6ICcg4oaTJykgOiAnIOKGlScsICdzb3J0LWFycm93JykpOwogICAgdGguYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7CiAgICAgIHNvcnRTdGF0ZSA9IHsga2V5LCBkaXI6IHNvcnRTdGF0ZS5rZXkgPT09IGtleSAmJiBzb3J0U3RhdGUuZGlyID09PSAnYXNjJyA/ICdkZXNjJyA6ICdhc2MnLCB0eXBlIH07CiAgICAgIHJlbmRlclRhYmxlKCk7CiAgICB9KTsKICAgIHJldHVybiB0aDsKICB9CgogIGZ1bmN0aW9uIGZpbHRlcmVkUm93cygpIHsKICAgIGNvbnN0IHEgPSBzZWFyY2hWYWx1ZS50cmltKCkudG9Mb3dlckNhc2UoKTsKICAgIGlmICghcSkgcmV0dXJuIGN1cnJlbnRSb3dzOwogICAgcmV0dXJuIGN1cnJlbnRSb3dzLmZpbHRlcigocm93KSA9PiB7CiAgICAgIGNvbnN0IGxhYmVsID0gcGxhdGZvcm1NZXRhRm9yKHJvdy5wbGF0Zm9ybSkubGFiZWwudG9Mb3dlckNhc2UoKTsKICAgICAgcmV0dXJuIGxhYmVsLmluY2x1ZGVzKHEpIHx8IHJvdy5lbnRyeV9kYXRlLmluY2x1ZGVzKHEpIHx8IFN0cmluZyhyb3cuZm9sbG93ZXJzX2NvdW50KS5pbmNsdWRlcyhxKTsKICAgIH0pOwogIH0KCiAgZnVuY3Rpb24gc29ydGVkUm93cygpIHsKICAgIGNvbnN0IHsga2V5LCBkaXIsIHR5cGUgfSA9IHNvcnRTdGF0ZTsKICAgIGNvbnN0IGZhY3RvciA9IGRpciA9PT0gJ2FzYycgPyAxIDogLTE7CiAgICByZXR1cm4gWy4uLmZpbHRlcmVkUm93cygpXS5zb3J0KChhLCBiKSA9PiB7CiAgICAgIGNvbnN0IGF2ID0gYVtrZXldOwogICAgICBjb25zdCBidiA9IGJba2V5XTsKICAgICAgaWYgKGF2ID09PSBudWxsIHx8IGF2ID09PSB1bmRlZmluZWQpIHJldHVybiAxOwogICAgICBpZiAoYnYgPT09IG51bGwgfHwgYnYgPT09IHVuZGVmaW5lZCkgcmV0dXJuIC0xOwogICAgICBpZiAodHlwZSA9PT0gJ251bWJlcicpIHJldHVybiAoYXYgLSBidikgKiBmYWN0b3I7CiAgICAgIHJldHVybiBTdHJpbmcoYXYpLmxvY2FsZUNvbXBhcmUoU3RyaW5nKGJ2KSkgKiBmYWN0b3I7CiAgICB9KTsKICB9CgogIC8qKiBFdmVyeSBjdXJyZW50bHktZmlsdGVyZWQvc29ydGVkIHJvdywgc2hhcGVkIGZvciBleHBvcnQgKG5vdCBqdXN0IHRoZSBjdXJyZW50IHBhZ2UpLiAqLwogIGZ1bmN0aW9uIGV4cG9ydFJvd3MoKSB7CiAgICByZXR1cm4gc29ydGVkUm93cygpLm1hcCgocm93KSA9PiAoewogICAgICBwbGF0Zm9ybV9sYWJlbDogcGxhdGZvcm1NZXRhRm9yKHJvdy5wbGF0Zm9ybSkubGFiZWwsCiAgICAgIGVudHJ5X2RhdGU6IHJvdy5lbnRyeV9kYXRlLAogICAgICBmb2xsb3dlcnNfY291bnQ6IHJvdy5mb2xsb3dlcnNfY291bnQsCiAgICAgIHVwZGF0ZWRfYXQ6IHJvdy51cGRhdGVkX2F0LAogICAgfSkpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyVGFibGUoKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZvbGxvd2Vyc1RhYmxlV3JhcCcpOwogICAgY29uc3QgcGFnZXJFbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmb2xsb3dlcnNQYWdlcicpOwogICAgaWYgKCF3cmFwKSByZXR1cm47CiAgICBjb25zdCBhbGxTb3J0ZWQgPSBzb3J0ZWRSb3dzKCk7CiAgICBpZiAoIWFsbFNvcnRlZC5sZW5ndGgpIHsKICAgICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgICAgd3JhcC5hcHBlbmRDaGlsZChlbXB0eVN0YXRlKHsKICAgICAgICBpY29uOiAndXNlcnMnLAogICAgICAgIHRpdGxlOiBjdXJyZW50Um93cy5sZW5ndGggPyAnTm8gZW50cmllcyBtYXRjaCB5b3VyIHNlYXJjaCcgOiAnTm8gZm9sbG93ZXIgZW50cmllcyB5ZXQnLAogICAgICAgIG1lc3NhZ2U6IGN1cnJlbnRSb3dzLmxlbmd0aCA/ICdUcnkgYSBkaWZmZXJlbnQgcGxhdGZvcm0gbmFtZSBvciBkYXRlLicgOiAnQWRkIHlvdXIgZmlyc3Qgd2Vla2x5IGZvbGxvd2VyIGNvdW50IGFib3ZlIGZvciBhbnkgcGxhdGZvcm0uJywKICAgICAgfSkpOwogICAgICBpZiAocGFnZXJFbCkgcGFnZXJFbC5pbm5lckhUTUwgPSAnJzsKICAgICAgcmV0dXJuOwogICAgfQogICAgY29uc3QgeyBwYWdlUm93cywgdG90YWxQYWdlcywgc2FmZVBhZ2UsIHRvdGFsIH0gPSBwYWdpbmF0ZUNsaWVudFNpZGUoYWxsU29ydGVkLCBwYWdlLCBwYWdlU2l6ZSk7CiAgICBwYWdlID0gc2FmZVBhZ2U7CgogICAgY29uc3QgdGFibGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0YWJsZScpOwogICAgdGFibGUuY2xhc3NOYW1lID0gJ2RhdGEtdGFibGUnOwogICAgY29uc3QgdGhlYWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0aGVhZCcpOwogICAgY29uc3QgaGVhZFRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgIGhlYWRUci5hcHBlbmQoCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdQbGF0Zm9ybScsICdwbGF0Zm9ybScsICdzdHJpbmcnKSwKICAgICAgc29ydGFibGVIZWFkZXIoJ1dlZWsgLyBEYXRlJywgJ2VudHJ5X2RhdGUnLCAnc3RyaW5nJyksCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdGb2xsb3dlcnMgQ291bnQnLCAnZm9sbG93ZXJzX2NvdW50JywgJ251bWJlcicpLAogICAgICB0ZXh0RWwoJ3RoJywgJ0xhc3QgVXBkYXRlZCcpLAogICAgICB0ZXh0RWwoJ3RoJywgJ0FjdGlvbnMnKQogICAgKTsKICAgIHRoZWFkLmFwcGVuZENoaWxkKGhlYWRUcik7CiAgICBjb25zdCB0Ym9keSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3Rib2R5Jyk7CiAgICBwYWdlUm93cy5mb3JFYWNoKChyb3cpID0+IHsKICAgICAgY29uc3QgdHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgICBjb25zdCBwbGF0Zm9ybVRkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgY29uc3QgbWV0YSA9IHBsYXRmb3JtTWV0YUZvcihyb3cucGxhdGZvcm0pOwogICAgICBjb25zdCBwaWxsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICBwaWxsLmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1waWxsJzsKICAgICAgY29uc3QgZG90ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICBkb3QuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLWRvdCc7CiAgICAgIGRvdC5zdHlsZS5iYWNrZ3JvdW5kID0gbWV0YS5jb2xvcjsKICAgICAgcGlsbC5hcHBlbmQoZG90LCBkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShtZXRhLmxhYmVsKSk7CiAgICAgIHBsYXRmb3JtVGQuYXBwZW5kQ2hpbGQocGlsbCk7CgogICAgICBjb25zdCBhY3Rpb25zVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBjb25zdCByb3dBY3Rpb25zID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIHJvd0FjdGlvbnMuY2xhc3NOYW1lID0gJ3Jvdy1hY3Rpb25zJzsKICAgICAgY29uc3QgZWRpdEJ0biA9IGljb25CdG4oJ2J0bicsICdwZW5jaWwnLCAnRWRpdCcpOwogICAgICBlZGl0QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gc3RhcnRFZGl0KHJvdykpOwogICAgICBjb25zdCBkZWxldGVCdG4gPSBpY29uQnRuKCdidG4gZGFuZ2VyJywgJ3RyYXNoLTInLCAnRGVsZXRlJyk7CiAgICAgIGRlbGV0ZUJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IGhhbmRsZURlbGV0ZShyb3cpKTsKICAgICAgcm93QWN0aW9ucy5hcHBlbmQoZWRpdEJ0biwgZGVsZXRlQnRuKTsKICAgICAgYWN0aW9uc1RkLmFwcGVuZENoaWxkKHJvd0FjdGlvbnMpOwoKICAgICAgdHIuYXBwZW5kKAogICAgICAgIHBsYXRmb3JtVGQsCiAgICAgICAgdGV4dEVsKCd0ZCcsIEZvcm1hdC5kYXRlKHJvdy5lbnRyeV9kYXRlKSksCiAgICAgICAgdGV4dEVsKCd0ZCcsIEZvcm1hdC5udW1iZXIocm93LmZvbGxvd2Vyc19jb3VudCksICdudW0nKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgcm93LnVwZGF0ZWRfYXQpLAogICAgICAgIGFjdGlvbnNUZAogICAgICApOwogICAgICB0Ym9keS5hcHBlbmRDaGlsZCh0cik7CiAgICB9KTsKICAgIHRhYmxlLmFwcGVuZCh0aGVhZCwgdGJvZHkpOwogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGFibGUpOwoKICAgIGlmIChwYWdlckVsKSB7CiAgICAgIHBhZ2VyRWwuaW5uZXJIVE1MID0gJyc7CiAgICAgIHBhZ2VyRWwuYXBwZW5kQ2hpbGQoYnVpbGRQYWdlcih7CiAgICAgICAgcGFnZTogc2FmZVBhZ2UsCiAgICAgICAgdG90YWxQYWdlcywKICAgICAgICB0b3RhbCwKICAgICAgICBvblByZXY6ICgpID0+IHsgcGFnZSAtPSAxOyByZW5kZXJUYWJsZSgpOyB9LAogICAgICAgIG9uTmV4dDogKCkgPT4geyBwYWdlICs9IDE7IHJlbmRlclRhYmxlKCk7IH0sCiAgICAgIH0pKTsKICAgIH0KICB9CgogIGFzeW5jIGZ1bmN0aW9uIGxvYWQoKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZvbGxvd2Vyc1RhYmxlV3JhcCcpOwogICAgaWYgKHdyYXApIHsgd3JhcC5pbm5lckhUTUwgPSAnJzsgd3JhcC5hcHBlbmRDaGlsZChza2VsZXRvblJvd3MoNCkpOyB9CiAgICBjdXJyZW50Um93cyA9IGF3YWl0IEFwaS5saXN0Rm9sbG93ZXJzKHt9KTsKICAgIHJlbmRlclRhYmxlKCk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiByZW5kZXIoKSB7CiAgICByb290ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3ZpZXctZm9sbG93ZXJzJyk7CiAgICBlZGl0aW5nSWQgPSBudWxsOwogICAgc2hlbGwoKTsKICAgIGF3YWl0IGxvYWQoKTsKICB9CgogIHJldHVybiB7IHJlbmRlciB9Owp9KSgpOwoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIFJlcG9ydCBHZW5lcmF0b3IgdGFiOiBvbmUtY2xpY2sgV2Vla2x5IFNvY2lhbCBNZWRpYSBBbmFseXRpY3MKICAgUmVwb3J0LiBUaGUgdXNlciBwaWNrcyBhIHJlcG9ydGluZyB3ZWVrOyB0aGUgaW1tZWRpYXRlbHktCiAgIHByZWNlZGluZyBlcXVhbC1sZW5ndGggd2VlayBpcyBkZXJpdmVkIGF1dG9tYXRpY2FsbHkuIEV2ZXJ5CiAgIG51bWJlciBjb21lcyBmcm9tIHRoZSBTQU1FIGVuZHBvaW50cyB0aGUgcmVzdCBvZiB0aGUgYXBwIHVzZXMg4oCUCiAgIC9hcGkvYW5hbHl0aWNzL2NvbXBhcmUgKHBvc3QgbWV0cmljcyArIHBlci1wbGF0Zm9ybSBzcGxpdCBmb3IKICAgYm90aCB3ZWVrcyBpbiBhIHNpbmdsZSBjYWxsKSwgL2FwaS9hbmFseXRpY3MvdG9wLXBvc3RzICh0b3AKICAgY29udGVudCksIGFuZCAvYXBpL2ZvbGxvd2VycyAodGhlIEZvbGxvd2VycyBEYXRhIFJlY29yZCwgcmVhZAogICBvbmx5LCByZWR1Y2VkIGhlcmUgdG8gYXMtb2YtZGF0ZSB0b3RhbHMpIOKAlCBzbyB0aGVyZSBpcyBubyBzZWNvbmQKICAgYW5hbHl0aWNzIHN5c3RlbSBhbmQgbm90aGluZyBpcyBldmVyIHdyaXR0ZW4gYmFjay4gVGhlIHdyaXR0ZW4KICAgc3VtbWFyeSdzIHdvcmRpbmcgaXMgZGVyaXZlZCBmcm9tIHRoZSBhY3R1YWwgZGVsdGFzLCBuZXZlciBhCiAgIGZpeGVkICJwZXJmb3JtYW5jZSBpbXByb3ZlZCIuIFByaW50IC8gRXhwb3J0IFJlcG9ydCBpcwogICB3aW5kb3cucHJpbnQoKSBhZ2FpbnN0IHRoZSBAbWVkaWEgcHJpbnQgYmxvY2sgaW4gPHN0eWxlPi4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmNvbnN0IFJlcG9ydCA9ICgoKSA9PiB7CiAgLy8gUm93cyBpbiB0aGUgbWV0cmljIGNvbXBhcmlzb24gdGFibGUgKyBzdW1tYXJ5LiBgc291cmNlYCBub3RlcyB3aGVyZSB0aGUgdmFsdWUKICAvLyBjb21lcyBmcm9tOyBgZm10YCBwaWNrcyB0aGUgZGlzcGxheSBmb3JtYXQuCiAgY29uc3QgTUVUUklDX1JPV1MgPSBbCiAgICB7IGtleTogJ3JlYWNoJywgbGFiZWw6ICdSZWFjaCcsIHNvdXJjZTogJ3Bvc3RzJywgZm10OiAnc21hcnQnIH0sCiAgICB7IGtleTogJ2VuZ2FnZW1lbnQnLCBsYWJlbDogJ0VuZ2FnZW1lbnQnLCBzb3VyY2U6ICdwb3N0cycsIGZtdDogJ3NtYXJ0JyB9LAogICAgeyBrZXk6ICdpbXByZXNzaW9ucycsIGxhYmVsOiAnSW1wcmVzc2lvbnMnLCBzb3VyY2U6ICdwb3N0cycsIGZtdDogJ3NtYXJ0JyB9LAogICAgeyBrZXk6ICdmb2xsb3dlcnMnLCBsYWJlbDogJ0ZvbGxvd2VycycsIHNvdXJjZTogJ2ZvbGxvd2VycycsIGZtdDogJ3NtYXJ0JyB9LAogICAgeyBrZXk6ICdmb2xsb3dlcnNfZ3Jvd3RoJywgbGFiZWw6ICdGb2xsb3dlcnMgR3Jvd3RoJywgc291cmNlOiAnZ3Jvd3RoJywgZm10OiAnc2lnbmVkJyB9LAogICAgeyBrZXk6ICdwb3N0cycsIGxhYmVsOiAnUG9zdHMgUHVibGlzaGVkJywgc291cmNlOiAnY291bnQnLCBmbXQ6ICdpbnQnIH0sCiAgXTsKICAvLyAiUmFuayB0b3AgcG9zdHMgYnkiIOKAlCBvbmx5IG1ldHJpY3MgL2FwaS9hbmFseXRpY3MvdG9wLXBvc3RzIGNhbiBhY3R1YWxseSBzb3J0IG9uLgogIGNvbnN0IFJBTktfTUVUUklDUyA9IFsKICAgIHsga2V5OiAnZW5nYWdlbWVudCcsIGxhYmVsOiAnRW5nYWdlbWVudCcgfSwKICAgIHsga2V5OiAncmVhY2gnLCBsYWJlbDogJ1JlYWNoJyB9LAogICAgeyBrZXk6ICdpbXByZXNzaW9ucycsIGxhYmVsOiAnSW1wcmVzc2lvbnMnIH0sCiAgICB7IGtleTogJ3ZpZXdzJywgbGFiZWw6ICdWaWV3cycgfSwKICAgIHsga2V5OiAnY2xpY2tzJywgbGFiZWw6ICdDbGlja3MnIH0sCiAgICB7IGtleTogJ3NoYXJlcycsIGxhYmVsOiAnU2hhcmVzJyB9LAogICAgeyBrZXk6ICdjb21tZW50cycsIGxhYmVsOiAnQ29tbWVudHMnIH0sCiAgICB7IGtleTogJ3NhdmVzJywgbGFiZWw6ICdTYXZlcycgfSwKICBdOwogIGNvbnN0IENIQVJUX01FVFJJQ1MgPSBbCiAgICB7IGtleTogJ2FsbCcsIGxhYmVsOiAnUmVhY2ggwrcgRW5nYWdlbWVudCDCtyBJbXByZXNzaW9ucycgfSwKICAgIHsga2V5OiAncmVhY2gnLCBsYWJlbDogJ1JlYWNoJyB9LAogICAgeyBrZXk6ICdlbmdhZ2VtZW50JywgbGFiZWw6ICdFbmdhZ2VtZW50JyB9LAogICAgeyBrZXk6ICdpbXByZXNzaW9ucycsIGxhYmVsOiAnSW1wcmVzc2lvbnMnIH0sCiAgICB7IGtleTogJ2ZvbGxvd2VycycsIGxhYmVsOiAnRm9sbG93ZXJzJyB9LAogICAgeyBrZXk6ICdmb2xsb3dlcnNfZ3Jvd3RoJywgbGFiZWw6ICdGb2xsb3dlcnMgR3Jvd3RoJyB9LAogICAgeyBrZXk6ICdwb3N0cycsIGxhYmVsOiAnUG9zdHMgUHVibGlzaGVkJyB9LAogIF07CgogIGxldCByb290OwogIGNvbnN0IHRvZGF5SXNvID0gbmV3IERhdGUoKS50b0lTT1N0cmluZygpLnNsaWNlKDAsIDEwKTsKICBsZXQgcmFuZ2VTdGFydCA9IG1vbmRheU9mKHRvZGF5SXNvKTsKICBsZXQgcmFuZ2VFbmQgPSBhZGREYXlzKHJhbmdlU3RhcnQsIDYpOwogIGxldCByYW5rTWV0cmljID0gJ2VuZ2FnZW1lbnQnOwogIGxldCBjaGFydE1ldHJpYyA9ICdhbGwnOwogIGxldCBsYXN0TW9kZWwgPSBudWxsOwogIGxldCBidXN5ID0gZmFsc2U7CgogIGZ1bmN0aW9uIG1vbmRheU9mKGRhdGVTdHIpIHsKICAgIGNvbnN0IGQgPSBuZXcgRGF0ZShkYXRlU3RyICsgJ1QwMDowMDowMCcpOwogICAgY29uc3QgZGF5ID0gZC5nZXREYXkoKTsKICAgIGQuc2V0RGF0ZShkLmdldERhdGUoKSAtIChkYXkgPT09IDAgPyA2IDogZGF5IC0gMSkpOwogICAgcmV0dXJuIGQudG9JU09TdHJpbmcoKS5zbGljZSgwLCAxMCk7CiAgfQogIGZ1bmN0aW9uIGFkZERheXMoZGF0ZVN0ciwgbikgewogICAgY29uc3QgZCA9IG5ldyBEYXRlKGRhdGVTdHIgKyAnVDAwOjAwOjAwJyk7CiAgICBkLnNldERhdGUoZC5nZXREYXRlKCkgKyBuKTsKICAgIHJldHVybiBkLnRvSVNPU3RyaW5nKCkuc2xpY2UoMCwgMTApOwogIH0KICBmdW5jdGlvbiBkYXlzSW5jbHVzaXZlKGZyb20sIHRvKSB7CiAgICByZXR1cm4gTWF0aC5yb3VuZCgobmV3IERhdGUodG8gKyAnVDAwOjAwOjAwJykgLSBuZXcgRGF0ZShmcm9tICsgJ1QwMDowMDowMCcpKSAvIDg2NDAwMDAwKSArIDE7CiAgfQogIC8qKiBCYWNrZW5kLWlkZW50aWNhbCB6ZXJvLXNhZmUgcGVyY2VudCBjaGFuZ2U6IG51bGwgd2hlbiB0aGVyZSBpcyBubyBiYXNlbGluZQogICAgICB0byBjb21wYXJlIGFnYWluc3QgKHByZXZpb3VzIDAsIGN1cnJlbnQgPiAwKTsgMCB3aGVuIGJvdGggc2lkZXMgYXJlIDAuICovCiAgZnVuY3Rpb24gcGN0Q2hhbmdlKGN1ciwgcHJldikgewogICAgY3VyID0gY3VyIHx8IDA7IHByZXYgPSBwcmV2IHx8IDA7CiAgICBpZiAoIXByZXYpIHJldHVybiBjdXIgPiAwID8gbnVsbCA6IDA7CiAgICByZXR1cm4gTWF0aC5yb3VuZCgoKGN1ciAtIHByZXYpIC8gcHJldikgKiAxMDAwKSAvIDEwOwogIH0KICBmdW5jdGlvbiBmbXRWYWx1ZShraW5kLCB2KSB7CiAgICBpZiAodiA9PT0gbnVsbCB8fCB2ID09PSB1bmRlZmluZWQpIHJldHVybiAn4oCUJzsKICAgIGlmIChraW5kID09PSAnaW50JykgcmV0dXJuIEZvcm1hdC5udW1iZXIodik7CiAgICBpZiAoa2luZCA9PT0gJ3NpZ25lZCcpIHJldHVybiAodiA+IDAgPyAnKycgOiAnJykgKyBGb3JtYXQuc21hcnQodik7CiAgICByZXR1cm4gRm9ybWF0LnNtYXJ0KHYpOwogIH0KICBmdW5jdGlvbiBmbXREaWZmKGtpbmQsIHYpIHsKICAgIGlmICh2ID09PSBudWxsIHx8IHYgPT09IHVuZGVmaW5lZCkgcmV0dXJuICfigJQnOwogICAgY29uc3QgYmFzZSA9IGtpbmQgPT09ICdpbnQnID8gRm9ybWF0Lm51bWJlcihNYXRoLmFicyh2KSkgOiBGb3JtYXQuc21hcnQoTWF0aC5hYnModikpOwogICAgaWYgKHYgPiAwKSByZXR1cm4gJysnICsgYmFzZTsKICAgIGlmICh2IDwgMCkgcmV0dXJuICfiiJInICsgYmFzZTsKICAgIHJldHVybiAnMCc7CiAgfQogIGZ1bmN0aW9uIHBjdFRleHQocGN0KSB7CiAgICBpZiAocGN0ID09PSBudWxsKSByZXR1cm4gJ24vYSc7CiAgICByZXR1cm4gKHBjdCA+IDAgPyAnKycgOiAnJykgKyBOdW1iZXIocGN0KS50b0ZpeGVkKDEpLnJlcGxhY2UoL1wuMCQvLCAnJykgKyAnJSc7CiAgfQogIGZ1bmN0aW9uIGFic1BjdFRleHQocGN0KSB7CiAgICBpZiAocGN0ID09PSBudWxsKSByZXR1cm4gJ24vYSc7CiAgICByZXR1cm4gTWF0aC5hYnMocGN0KS50b0ZpeGVkKDEpLnJlcGxhY2UoL1wuMCQvLCAnJykgKyAnJSc7CiAgfQogIGZ1bmN0aW9uIGFycm93Rm9yKGRpZmYpIHsgcmV0dXJuIGRpZmYgPiAwID8gJ+KGkScgOiBkaWZmIDwgMCA/ICfihpMnIDogJ+KGkic7IH0KICBmdW5jdGlvbiB0cmVuZENsYXNzKGRpZmYpIHsgcmV0dXJuIGRpZmYgPiAwID8gJ3RyZW5kLXVwJyA6IGRpZmYgPCAwID8gJ3RyZW5kLWRvd24nIDogJ3RyZW5kLWZsYXQnOyB9CiAgZnVuY3Rpb24gd29yZEZvcihkaWZmKSB7IHJldHVybiBkaWZmID4gMCA/ICdpbmNyZWFzZWQnIDogZGlmZiA8IDAgPyAnZGVjcmVhc2VkJyA6ICdoZWxkIGZsYXQnOyB9CgogIGZ1bmN0aW9uIHBlcmlvZHMoKSB7CiAgICBjb25zdCBjdXIgPSB7IGZyb206IHJhbmdlU3RhcnQsIHRvOiByYW5nZUVuZCB9OwogICAgY29uc3QgbGVuID0gZGF5c0luY2x1c2l2ZShyYW5nZVN0YXJ0LCByYW5nZUVuZCk7CiAgICBjb25zdCBwcmV2ID0geyBmcm9tOiBhZGREYXlzKHJhbmdlU3RhcnQsIC1sZW4pLCB0bzogYWRkRGF5cyhyYW5nZVN0YXJ0LCAtMSkgfTsKICAgIHJldHVybiB7IGN1ciwgcHJldiwgbGVuIH07CiAgfQoKICAvKiogTGF0ZXN0IEZvbGxvd2VycyBEYXRhIFJlY29yZCBlbnRyeSBwZXIgcGxhdGZvcm0gb24gb3IgYmVmb3JlIGBkYXRlU3RyYCwgc3VtbWVkLiBSZWFkLW9ubHkuICovCiAgZnVuY3Rpb24gZm9sbG93ZXJzQXNPZihyb3dzLCBkYXRlU3RyKSB7CiAgICBjb25zdCBieVBsYXQgPSB7fTsKICAgIHJvd3MuZm9yRWFjaCgocikgPT4gewogICAgICBpZiAoci5lbnRyeV9kYXRlID4gZGF0ZVN0cikgcmV0dXJuOwogICAgICBjb25zdCBjdXIgPSBieVBsYXRbci5wbGF0Zm9ybV07CiAgICAgIGlmICghY3VyIHx8IHIuZW50cnlfZGF0ZSA+IGN1ci5lbnRyeV9kYXRlKSBieVBsYXRbci5wbGF0Zm9ybV0gPSByOwogICAgfSk7CiAgICBjb25zdCBwbGF0Zm9ybXMgPSBPYmplY3Qua2V5cyhieVBsYXQpOwogICAgcmV0dXJuIHsKICAgICAgdG90YWw6IHBsYXRmb3Jtcy5yZWR1Y2UoKHMsIHApID0+IHMgKyBieVBsYXRbcF0uZm9sbG93ZXJzX2NvdW50LCAwKSwKICAgICAgYnlQbGF0LCBwbGF0Zm9ybXMsIGhhc0RhdGE6IHBsYXRmb3Jtcy5sZW5ndGggPiAwLAogICAgfTsKICB9CiAgLyoqIE5ldCBmb2xsb3dlciBjaGFuZ2Ugb3ZlciAoZGF5IGJlZm9yZSBgZnJvbWApIC4uIGB0b2AsIGNvdW50aW5nIG9ubHkgcGxhdGZvcm1zIHdpdGggYW4KICAgICAgZW50cnkgb24gQk9USCBzaWRlcyBzbyBhIG1pc3NpbmcgYmFzZWxpbmUgY2FuJ3QgbWFzcXVlcmFkZSBhcyBncm93dGguICovCiAgZnVuY3Rpb24gZm9sbG93ZXJzRGVsdGEocm93cywgZnJvbSwgdG8pIHsKICAgIGNvbnN0IGJhc2UgPSBmb2xsb3dlcnNBc09mKHJvd3MsIGFkZERheXMoZnJvbSwgLTEpKTsKICAgIGNvbnN0IGVuZCA9IGZvbGxvd2Vyc0FzT2Yocm93cywgdG8pOwogICAgbGV0IGRlbHRhID0gMDsgbGV0IGNvdW50ZWQgPSAwOwogICAgZW5kLnBsYXRmb3Jtcy5mb3JFYWNoKChwKSA9PiB7CiAgICAgIGlmIChiYXNlLmJ5UGxhdFtwXSkgeyBkZWx0YSArPSBlbmQuYnlQbGF0W3BdLmZvbGxvd2Vyc19jb3VudCAtIGJhc2UuYnlQbGF0W3BdLmZvbGxvd2Vyc19jb3VudDsgY291bnRlZCArPSAxOyB9CiAgICB9KTsKICAgIHJldHVybiB7IGRlbHRhLCBoYXNEYXRhOiBjb3VudGVkID4gMCB9OwogIH0KCiAgZnVuY3Rpb24gbWVyZ2VQbGF0Zm9ybXMoYSwgYikgewogICAgY29uc3QgbWFwID0gbmV3IE1hcCgpOwogICAgY29uc3QgcHV0ID0gKGFyciwgc2lkZSkgPT4gKGFyciB8fCBbXSkuZm9yRWFjaCgocCkgPT4gewogICAgICBjb25zdCBlID0gbWFwLmdldChwLnBsYXRmb3JtKSB8fCB7IHBsYXRmb3JtOiBwLnBsYXRmb3JtLCBsYWJlbDogcC5sYWJlbCwgY29sb3I6IHAuY29sb3IgfTsKICAgICAgZVtzaWRlXSA9IHA7IGUubGFiZWwgPSBwLmxhYmVsOyBlLmNvbG9yID0gcC5jb2xvcjsKICAgICAgbWFwLnNldChwLnBsYXRmb3JtLCBlKTsKICAgIH0pOwogICAgcHV0KGEsICdjdXInKTsgcHV0KGIsICdwcmV2Jyk7CiAgICByZXR1cm4gWy4uLm1hcC52YWx1ZXMoKV0ubWFwKChlKSA9PiB7CiAgICAgIGNvbnN0IGMgPSBlLmN1ciB8fCB7fTsgY29uc3QgcHIgPSBlLnByZXYgfHwge307CiAgICAgIHJldHVybiB7CiAgICAgICAgcGxhdGZvcm06IGUucGxhdGZvcm0sIGxhYmVsOiBlLmxhYmVsIHx8IGUucGxhdGZvcm0sIGNvbG9yOiBlLmNvbG9yIHx8ICcjOTk5JywKICAgICAgICBwb3N0c0N1cjogYy5wb3N0X2NvdW50IHx8IDAsIHBvc3RzUHJldjogcHIucG9zdF9jb3VudCB8fCAwLAogICAgICAgIHJlYWNoQ3VyOiBjLnJlYWNoIHx8IDAsIHJlYWNoUHJldjogcHIucmVhY2ggfHwgMCwKICAgICAgICBlbmdDdXI6IGMuZW5nYWdlbWVudCB8fCAwLCBlbmdQcmV2OiBwci5lbmdhZ2VtZW50IHx8IDAsCiAgICAgICAgZW5nUGN0OiBwY3RDaGFuZ2UoYy5lbmdhZ2VtZW50IHx8IDAsIHByLmVuZ2FnZW1lbnQgfHwgMCksCiAgICAgIH07CiAgICB9KS5zb3J0KCh4LCB5KSA9PiB5LmVuZ0N1ciAtIHguZW5nQ3VyKTsKICB9CgogIGZ1bmN0aW9uIGJ1aWxkTW9kZWwoY21wLCBmb2xsb3dlclJvd3MsIHRvcFBvc3RzLCBjdXIsIHByZXYpIHsKICAgIGNvbnN0IHRBID0gY21wLnJhbmdlQS50b3RhbHM7IGNvbnN0IHRCID0gY21wLnJhbmdlQi50b3RhbHM7CiAgICBjb25zdCBmVGhpcyA9IGZvbGxvd2Vyc0FzT2YoZm9sbG93ZXJSb3dzLCBjdXIudG8pOwogICAgY29uc3QgZkxhc3QgPSBmb2xsb3dlcnNBc09mKGZvbGxvd2VyUm93cywgcHJldi50byk7CiAgICBjb25zdCBnVGhpcyA9IGZvbGxvd2Vyc0RlbHRhKGZvbGxvd2VyUm93cywgY3VyLmZyb20sIGN1ci50byk7CiAgICBjb25zdCBnTGFzdCA9IGZvbGxvd2Vyc0RlbHRhKGZvbGxvd2VyUm93cywgcHJldi5mcm9tLCBwcmV2LnRvKTsKICAgIGNvbnN0IHBvc3RzSGF2ZURhdGEgPSAodEEucG9zdF9jb3VudCB8fCAwKSA+IDAgfHwgKHRCLnBvc3RfY291bnQgfHwgMCkgPiAwOwoKICAgIGNvbnN0IHJhdyA9IHsKICAgICAgcmVhY2g6IHsgY3VyOiB0QS5yZWFjaCB8fCAwLCBwcmV2OiB0Qi5yZWFjaCB8fCAwLCBoYXNEYXRhOiBwb3N0c0hhdmVEYXRhIH0sCiAgICAgIGVuZ2FnZW1lbnQ6IHsgY3VyOiB0QS5lbmdhZ2VtZW50IHx8IDAsIHByZXY6IHRCLmVuZ2FnZW1lbnQgfHwgMCwgaGFzRGF0YTogcG9zdHNIYXZlRGF0YSB9LAogICAgICBpbXByZXNzaW9uczogeyBjdXI6IHRBLmltcHJlc3Npb25zIHx8IDAsIHByZXY6IHRCLmltcHJlc3Npb25zIHx8IDAsIGhhc0RhdGE6IHBvc3RzSGF2ZURhdGEgfSwKICAgICAgZm9sbG93ZXJzOiB7IGN1cjogZlRoaXMudG90YWwsIHByZXY6IGZMYXN0LnRvdGFsLCBoYXNEYXRhOiBmVGhpcy5oYXNEYXRhIHx8IGZMYXN0Lmhhc0RhdGEgfSwKICAgICAgZm9sbG93ZXJzX2dyb3d0aDogeyBjdXI6IGdUaGlzLmRlbHRhLCBwcmV2OiBnTGFzdC5kZWx0YSwgaGFzRGF0YTogZ1RoaXMuaGFzRGF0YSB8fCBnTGFzdC5oYXNEYXRhIH0sCiAgICAgIHBvc3RzOiB7IGN1cjogdEEucG9zdF9jb3VudCB8fCAwLCBwcmV2OiB0Qi5wb3N0X2NvdW50IHx8IDAsIGhhc0RhdGE6IHBvc3RzSGF2ZURhdGEgfSwKICAgIH07CiAgICBjb25zdCByb3dzID0gTUVUUklDX1JPV1MubWFwKChtKSA9PiB7CiAgICAgIGNvbnN0IHIgPSByYXdbbS5rZXldOwogICAgICBjb25zdCBkaWZmID0gci5jdXIgLSByLnByZXY7CiAgICAgIHJldHVybiB7IC4uLm0sIGN1cjogci5jdXIsIHByZXY6IHIucHJldiwgZGlmZiwgcGN0OiBwY3RDaGFuZ2Uoci5jdXIsIHIucHJldiksIGhhc0RhdGE6IHIuaGFzRGF0YSB9OwogICAgfSk7CiAgICBjb25zdCBwbGF0Zm9ybXMgPSBtZXJnZVBsYXRmb3JtcyhjbXAucmFuZ2VBLnBsYXRmb3JtcywgY21wLnJhbmdlQi5wbGF0Zm9ybXMpOwogICAgY29uc3QgYW55RGF0YSA9IHJvd3Muc29tZSgocikgPT4gci5oYXNEYXRhKSB8fCBwbGF0Zm9ybXMubGVuZ3RoID4gMCB8fCB0b3BQb3N0cy5sZW5ndGggPiAwOwogICAgcmV0dXJuIHsgY3VyLCBwcmV2LCByb3dzLCByYXcsIHBsYXRmb3JtcywgdG9wUG9zdHMsIGZvbGxvd2Vyc1RoaXM6IGZUaGlzLCBhbnlEYXRhIH07CiAgfQoKICBhc3luYyBmdW5jdGlvbiBnZW5lcmF0ZSgpIHsKICAgIGlmIChidXN5KSByZXR1cm47CiAgICBidXN5ID0gdHJ1ZTsKICAgIGNvbnN0IHJlc3VsdHMgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncmVwb3J0UmVzdWx0cycpOwogICAgaWYgKHJlc3VsdHMpIHsKICAgICAgcmVzdWx0cy5pbm5lckhUTUwgPSAnJzsKICAgICAgcmVzdWx0cy5hcHBlbmRDaGlsZChza2VsZXRvblN0YXRHcmlkKDYpKTsKICAgICAgcmVzdWx0cy5hcHBlbmRDaGlsZChza2VsZXRvbkNoYXJ0KCkpOwogICAgICByZXN1bHRzLmFwcGVuZENoaWxkKHNrZWxldG9uUm93cyg1KSk7CiAgICB9CiAgICB0cnkgewogICAgICBjb25zdCB7IGN1ciwgcHJldiB9ID0gcGVyaW9kcygpOwogICAgICBjb25zdCBbY21wLCBmb2xsb3dlclJvd3MsIHRvcFBvc3RzXSA9IGF3YWl0IFByb21pc2UuYWxsKFsKICAgICAgICBBcGkuY29tcGFyZSh7IGZyb21BOiBjdXIuZnJvbSwgdG9BOiBjdXIudG8sIGZyb21COiBwcmV2LmZyb20sIHRvQjogcHJldi50byB9KSwKICAgICAgICBBcGkubGlzdEZvbGxvd2Vycyh7fSksCiAgICAgICAgQXBpLnRvcFBvc3RzKHsgZGF0ZUZyb206IGN1ci5mcm9tLCBkYXRlVG86IGN1ci50bywgc29ydEJ5OiByYW5rTWV0cmljLCBsaW1pdDogMTAgfSksCiAgICAgIF0pOwogICAgICBsYXN0TW9kZWwgPSBidWlsZE1vZGVsKGNtcCwgZm9sbG93ZXJSb3dzLCB0b3BQb3N0cywgY3VyLCBwcmV2KTsKICAgICAgcmVuZGVyUmVwb3J0KCk7CiAgICB9IGNhdGNoIChlcnIpIHsKICAgICAgaWYgKHJlc3VsdHMpIHsKICAgICAgICByZXN1bHRzLmlubmVySFRNTCA9ICcnOwogICAgICAgIHJlc3VsdHMuYXBwZW5kQ2hpbGQoZW1wdHlTdGF0ZSh7CiAgICAgICAgICBpY29uOiAnYWxlcnQtdHJpYW5nbGUnLAogICAgICAgICAgdGl0bGU6ICdDb3VsZCBub3QgZ2VuZXJhdGUgdGhlIHJlcG9ydCcsCiAgICAgICAgICBtZXNzYWdlOiBlcnIubWVzc2FnZSB8fCAnU29tZXRoaW5nIHdlbnQgd3Jvbmcgd2hpbGUgZmV0Y2hpbmcgdGhlIGRhdGEuJywKICAgICAgICB9KSk7CiAgICAgIH0KICAgIH0gZmluYWxseSB7CiAgICAgIGJ1c3kgPSBmYWxzZTsKICAgIH0KICB9CgogIGZ1bmN0aW9uIGxhYmVsZWQobGFiZWwsIGVsKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB3cmFwLmNsYXNzTmFtZSA9ICdmaWVsZC1pbmxpbmUnOwogICAgd3JhcC5hcHBlbmQodGV4dEVsKCdsYWJlbCcsIGxhYmVsKSwgZWwpOwogICAgcmV0dXJuIHdyYXA7CiAgfQoKICBmdW5jdGlvbiBzaGVsbCgpIHsKICAgIHJvb3QuaW5uZXJIVE1MID0gJyc7CgogICAgY29uc3QgY29udHJvbHMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNvbnRyb2xzLmNsYXNzTmFtZSA9ICdjYXJkJzsKCiAgICBjb25zdCBjUm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjUm93LmNsYXNzTmFtZSA9ICdyZXBvcnQtY29udHJvbHMnOwoKICAgIGNvbnN0IGZyb21JbnB1dCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7CiAgICBmcm9tSW5wdXQudHlwZSA9ICdkYXRlJzsgZnJvbUlucHV0LnZhbHVlID0gcmFuZ2VTdGFydDsKICAgIGNvbnN0IHRvSW5wdXQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOwogICAgdG9JbnB1dC50eXBlID0gJ2RhdGUnOyB0b0lucHV0LnZhbHVlID0gcmFuZ2VFbmQ7CiAgICBjb25zdCByYW5rU2VsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc2VsZWN0Jyk7CiAgICBSQU5LX01FVFJJQ1MuZm9yRWFjaCgobSkgPT4gewogICAgICBjb25zdCBvID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7CiAgICAgIG8udmFsdWUgPSBtLmtleTsgby50ZXh0Q29udGVudCA9IG0ubGFiZWw7CiAgICAgIGlmIChtLmtleSA9PT0gcmFua01ldHJpYykgby5zZWxlY3RlZCA9IHRydWU7CiAgICAgIHJhbmtTZWwuYXBwZW5kQ2hpbGQobyk7CiAgICB9KTsKCiAgICBjb25zdCByZXNvbHZlZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcmVzb2x2ZWQuc3R5bGUuY3NzVGV4dCA9ICdtYXJnaW4tdG9wOjEwcHg7Zm9udC1zaXplOjEycHg7Y29sb3I6dmFyKC0tdGV4dC1tdXRlZCknOwogICAgZnVuY3Rpb24gdXBkYXRlUmVzb2x2ZWQoKSB7CiAgICAgIGNvbnN0IHMgPSBmcm9tSW5wdXQudmFsdWU7IGNvbnN0IGUgPSB0b0lucHV0LnZhbHVlIHx8IHM7CiAgICAgIGlmICghcyB8fCAhZSB8fCBlIDwgcykgeyByZXNvbHZlZC50ZXh0Q29udGVudCA9ICdQaWNrIGEgdmFsaWQgd2VlayDigJQgdGhlIFRvIGRhdGUgbXVzdCBiZSBvbiBvciBhZnRlciB0aGUgRnJvbSBkYXRlLic7IHJldHVybjsgfQogICAgICBjb25zdCBsZW4gPSBkYXlzSW5jbHVzaXZlKHMsIGUpOwogICAgICByZXNvbHZlZC50ZXh0Q29udGVudCA9ICdUaGlzIHdlZWs6ICcgKyBGb3JtYXQuZGF0ZShzKSArICcg4oCTICcgKyBGb3JtYXQuZGF0ZShlKQogICAgICAgICsgJyAgICAgwrcgICAgIFByZXZpb3VzIHdlZWs6ICcgKyBGb3JtYXQuZGF0ZShhZGREYXlzKHMsIC1sZW4pKSArICcg4oCTICcgKyBGb3JtYXQuZGF0ZShhZGREYXlzKHMsIC0xKSk7CiAgICB9CiAgICBmcm9tSW5wdXQuYWRkRXZlbnRMaXN0ZW5lcignY2hhbmdlJywgdXBkYXRlUmVzb2x2ZWQpOwogICAgdG9JbnB1dC5hZGRFdmVudExpc3RlbmVyKCdjaGFuZ2UnLCB1cGRhdGVSZXNvbHZlZCk7CiAgICB1cGRhdGVSZXNvbHZlZCgpOwoKICAgIGNSb3cuYXBwZW5kKAogICAgICBsYWJlbGVkKCdXZWVrIGZyb20nLCBmcm9tSW5wdXQpLAogICAgICBsYWJlbGVkKCdXZWVrIHRvJywgdG9JbnB1dCksCiAgICAgIGxhYmVsZWQoJ1JhbmsgdG9wIHBvc3RzIGJ5JywgcmFua1NlbCksCiAgICApOwogICAgY29udHJvbHMuYXBwZW5kQ2hpbGQoY1Jvdyk7CiAgICBjb250cm9scy5hcHBlbmRDaGlsZChyZXNvbHZlZCk7CgogICAgY29uc3QgYWN0aW9ucyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYWN0aW9ucy5jbGFzc05hbWUgPSAncmVwb3J0LWFjdGlvbnMnOwogICAgY29uc3QgZ2VuQnRuID0gaWNvbkJ0bignYnRuIHByaW1hcnknLCAnc3BhcmtsZXMnLCAnR2VuZXJhdGUgUmVwb3J0Jyk7CiAgICBnZW5CdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7CiAgICAgIGNvbnN0IHMgPSBmcm9tSW5wdXQudmFsdWU7IGNvbnN0IGUgPSB0b0lucHV0LnZhbHVlIHx8IHM7CiAgICAgIGlmICghcyB8fCAhZSB8fCBlIDwgcykgeyBUb2FzdC5zaG93KCdQaWNrIGEgdmFsaWQgd2VlayDigJQgdGhlIFRvIGRhdGUgbXVzdCBiZSBvbiBvciBhZnRlciB0aGUgRnJvbSBkYXRlLicsICdlcnJvcicpOyByZXR1cm47IH0KICAgICAgcmFuZ2VTdGFydCA9IHM7IHJhbmdlRW5kID0gZTsgcmFua01ldHJpYyA9IHJhbmtTZWwudmFsdWU7CiAgICAgIGdlbmVyYXRlKCk7CiAgICB9KTsKICAgIGNvbnN0IHJlZnJlc2hCdG4gPSBpY29uQnRuKCdidG4nLCAncmVmcmVzaC1jdycsICdSZWZyZXNoIERhdGEnKTsKICAgIHJlZnJlc2hCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7IHJhbmtNZXRyaWMgPSByYW5rU2VsLnZhbHVlOyBnZW5lcmF0ZSgpOyB9KTsKICAgIGNvbnN0IHByaW50QnRuID0gaWNvbkJ0bignYnRuJywgJ3ByaW50ZXInLCAnUHJpbnQgLyBFeHBvcnQgUmVwb3J0Jyk7CiAgICBwcmludEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsKICAgICAgaWYgKCFsYXN0TW9kZWwpIHsgVG9hc3Quc2hvdygnR2VuZXJhdGUgdGhlIHJlcG9ydCBmaXJzdC4nLCAnZXJyb3InKTsgcmV0dXJuOyB9CiAgICAgIHdpbmRvdy5wcmludCgpOwogICAgfSk7CiAgICBhY3Rpb25zLmFwcGVuZChnZW5CdG4sIHJlZnJlc2hCdG4sIHByaW50QnRuKTsKICAgIGNvbnRyb2xzLmFwcGVuZENoaWxkKGFjdGlvbnMpOwogICAgcm9vdC5hcHBlbmRDaGlsZChjb250cm9scyk7CgogICAgY29uc3QgcmVzdWx0cyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcmVzdWx0cy5pZCA9ICdyZXBvcnRSZXN1bHRzJzsKICAgIHJvb3QuYXBwZW5kQ2hpbGQocmVzdWx0cyk7CiAgfQoKICBmdW5jdGlvbiBtZXRyaWNUaWxlKHIpIHsKICAgIGNvbnN0IHRpbGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRpbGUuY2xhc3NOYW1lID0gJ3N0YXQtdGlsZSc7CiAgICB0aWxlLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2Jywgci5sYWJlbCwgJ3N0YXQtbGFiZWwnKSk7CiAgICBpZiAoIXIuaGFzRGF0YSkgewogICAgICB0aWxlLmFwcGVuZCh0ZXh0RWwoJ2RpdicsICfigJQnLCAnc3RhdC12YWx1ZScpLCB0ZXh0RWwoJ2RpdicsICdubyBkYXRhJywgJ3N0YXQtZGVsdGEgZmxhdCcpKTsKICAgICAgcmV0dXJuIHRpbGU7CiAgICB9CiAgICB0aWxlLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgZm10VmFsdWUoci5mbXQsIHIuY3VyKSwgJ3N0YXQtdmFsdWUnKSk7CiAgICBjb25zdCBkZWx0YVRleHQgPSAoci5wY3QgPT09IG51bGwgPyAnbmV3IHZzIGxhc3Qgd2VlaycgOiBwY3RUZXh0KHIucGN0KSkgKyAnIMK3IHdhcyAnICsgZm10VmFsdWUoci5mbXQsIHIucHJldik7CiAgICBjb25zdCBjbHMgPSByLnBjdCA9PT0gbnVsbCA/ICd1cCcgOiBGb3JtYXQuZGVsdGFDbGFzcyhyLnBjdCk7CiAgICB0aWxlLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgZGVsdGFUZXh0LCAnc3RhdC1kZWx0YSAnICsgY2xzKSk7CiAgICByZXR1cm4gdGlsZTsKICB9CgogIGZ1bmN0aW9uIG1ldHJpY1RhYmxlKHJvd3MpIHsKICAgIGNvbnN0IHRhYmxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGFibGUnKTsKICAgIHRhYmxlLmNsYXNzTmFtZSA9ICdkYXRhLXRhYmxlJzsKICAgIGNvbnN0IHRoZWFkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGhlYWQnKTsKICAgIGNvbnN0IGh0ciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICBodHIuYXBwZW5kKAogICAgICB0ZXh0RWwoJ3RoJywgJ01ldHJpYycpLAogICAgICB0ZXh0RWwoJ3RoJywgJ1RoaXMgV2VlaycsICdudW0nKSwKICAgICAgdGV4dEVsKCd0aCcsICdMYXN0IFdlZWsnLCAnbnVtJyksCiAgICAgIHRleHRFbCgndGgnLCAnRGlmZmVyZW5jZScsICdudW0nKSwKICAgICAgdGV4dEVsKCd0aCcsICclIENoYW5nZScsICdudW0nKSwKICAgICAgdGV4dEVsKCd0aCcsICdUcmVuZCcpLAogICAgKTsKICAgIHRoZWFkLmFwcGVuZENoaWxkKGh0cik7CiAgICBjb25zdCB0Ym9keSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3Rib2R5Jyk7CiAgICByb3dzLmZvckVhY2goKHIpID0+IHsKICAgICAgY29uc3QgdHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgICBpZiAoIXIuaGFzRGF0YSkgewogICAgICAgIHRyLmFwcGVuZCgKICAgICAgICAgIHRleHRFbCgndGQnLCByLmxhYmVsKSwgdGV4dEVsKCd0ZCcsICfigJQnLCAnbnVtJyksIHRleHRFbCgndGQnLCAn4oCUJywgJ251bScpLAogICAgICAgICAgdGV4dEVsKCd0ZCcsICfigJQnLCAnbnVtJyksIHRleHRFbCgndGQnLCAn4oCUJywgJ251bScpLCB0ZXh0RWwoJ3RkJywgJ+KAlCcpLAogICAgICAgICk7CiAgICAgICAgdGJvZHkuYXBwZW5kQ2hpbGQodHIpOwogICAgICAgIHJldHVybjsKICAgICAgfQogICAgICB0ci5hcHBlbmQoCiAgICAgICAgdGV4dEVsKCd0ZCcsIHIubGFiZWwpLAogICAgICAgIHRleHRFbCgndGQnLCBmbXRWYWx1ZShyLmZtdCwgci5jdXIpLCAnbnVtJyksCiAgICAgICAgdGV4dEVsKCd0ZCcsIGZtdFZhbHVlKHIuZm10LCByLnByZXYpLCAnbnVtJyksCiAgICAgICAgdGV4dEVsKCd0ZCcsIGZtdERpZmYoci5mbXQsIHIuZGlmZiksICdudW0gJyArIHRyZW5kQ2xhc3Moci5kaWZmKSksCiAgICAgICAgdGV4dEVsKCd0ZCcsIHIucGN0ID09PSBudWxsID8gJ24vYScgOiBwY3RUZXh0KHIucGN0KSwgJ251bSAnICsgdHJlbmRDbGFzcyhyLmRpZmYpKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgYXJyb3dGb3Ioci5kaWZmKSwgdHJlbmRDbGFzcyhyLmRpZmYpKSwKICAgICAgKTsKICAgICAgdGJvZHkuYXBwZW5kQ2hpbGQodHIpOwogICAgfSk7CiAgICB0YWJsZS5hcHBlbmQodGhlYWQsIHRib2R5KTsKICAgIHJldHVybiB0YWJsZTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlckNvbXBhcmVDaGFydCgpIHsKICAgIGNvbnN0IGhvc3QgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncmVwb3J0Q2hhcnRXcmFwJyk7CiAgICBpZiAoIWhvc3QgfHwgIWxhc3RNb2RlbCkgcmV0dXJuOwogICAgQ2hhcnRzLmRlc3Ryb3koJ3JlcG9ydENvbXBhcmVDYW52YXMnKTsKICAgIGNvbnN0IG0gPSBsYXN0TW9kZWw7CiAgICBjb25zdCBjVGhpcyA9IENoYXJ0cy5zZXJpZXNDb2xvcigwKTsKICAgIGNvbnN0IGNMYXN0ID0gQ2hhcnRzLnNlcmllc0NvbG9yKDEpOwogICAgbGV0IGxhYmVsczsgbGV0IHNlcmllczsgbGV0IGZtdCA9ICh2KSA9PiBGb3JtYXQuc21hcnQodik7CgogICAgaWYgKGNoYXJ0TWV0cmljID09PSAnYWxsJykgewogICAgICBjb25zdCBrZXlzID0gWydyZWFjaCcsICdlbmdhZ2VtZW50JywgJ2ltcHJlc3Npb25zJ107CiAgICAgIGNvbnN0IGhhcyA9IGtleXMuc29tZSgoaykgPT4gbS5yYXdba10uaGFzRGF0YSAmJiAobS5yYXdba10uY3VyIHx8IG0ucmF3W2tdLnByZXYpKTsKICAgICAgaWYgKCFoYXMpIHsKICAgICAgICBob3N0LmlubmVySFRNTCA9ICcnOwogICAgICAgIGhvc3QuYXBwZW5kQ2hpbGQoZW1wdHlTdGF0ZSh7IGljb246ICdiYXItY2hhcnQtMycsIHRpdGxlOiAnTm8gcG9zdCBtZXRyaWNzIGZvciB0aGlzIHdlZWsnLCBtZXNzYWdlOiAnUmVhY2gsIGVuZ2FnZW1lbnQgYW5kIGltcHJlc3Npb25zIGFyZSBhbGwgemVybyBmb3IgYm90aCB3ZWVrcy4nIH0pKTsKICAgICAgICByZXR1cm47CiAgICAgIH0KICAgICAgbGFiZWxzID0gWydSZWFjaCcsICdFbmdhZ2VtZW50JywgJ0ltcHJlc3Npb25zJ107CiAgICAgIHNlcmllcyA9IFsKICAgICAgICB7IGxhYmVsOiAnVGhpcyBXZWVrJywgY29sb3I6IGNUaGlzLCBkYXRhOiBrZXlzLm1hcCgoaykgPT4gbS5yYXdba10uY3VyKSB9LAogICAgICAgIHsgbGFiZWw6ICdMYXN0IFdlZWsnLCBjb2xvcjogY0xhc3QsIGRhdGE6IGtleXMubWFwKChrKSA9PiBtLnJhd1trXS5wcmV2KSB9LAogICAgICBdOwogICAgfSBlbHNlIHsKICAgICAgY29uc3Qgcm93ID0gbS5yb3dzLmZpbmQoKHIpID0+IHIua2V5ID09PSBjaGFydE1ldHJpYyk7CiAgICAgIGlmICghcm93IHx8ICFyb3cuaGFzRGF0YSkgewogICAgICAgIGhvc3QuaW5uZXJIVE1MID0gJyc7CiAgICAgICAgaG9zdC5hcHBlbmRDaGlsZChlbXB0eVN0YXRlKHsgaWNvbjogJ2Jhci1jaGFydC0zJywgdGl0bGU6ICdObyBkYXRhIGZvciAnICsgKHJvdyA/IHJvdy5sYWJlbCA6IGNoYXJ0TWV0cmljKSwgbWVzc2FnZTogJ1BpY2sgYW5vdGhlciBtZXRyaWMgb3IgYW5vdGhlciB3ZWVrLicgfSkpOwogICAgICAgIHJldHVybjsKICAgICAgfQogICAgICBpZiAocm93LmZtdCA9PT0gJ2ludCcpIGZtdCA9ICh2KSA9PiBGb3JtYXQubnVtYmVyKHYpOwogICAgICBsYWJlbHMgPSBbJ1RoaXMgV2VlaycsICdMYXN0IFdlZWsnXTsKICAgICAgc2VyaWVzID0gW3sgbGFiZWw6IHJvdy5sYWJlbCwgY29sb3I6IFtjVGhpcywgY0xhc3RdLCBkYXRhOiBbcm93LmN1ciwgcm93LnByZXZdIH1dOwogICAgfQogICAgaG9zdC5pbm5lckhUTUwgPSAnPGNhbnZhcyBpZD0icmVwb3J0Q29tcGFyZUNhbnZhcyI+PC9jYW52YXM+JzsKICAgIENoYXJ0cy5ncm91cGVkQmFyQ2hhcnQoJ3JlcG9ydENvbXBhcmVDYW52YXMnLCB7IGxhYmVscywgc2VyaWVzLCBmb3JtYXRWYWx1ZTogZm10IH0pOwogIH0KCiAgZnVuY3Rpb24gcGxhdGZvcm1TZWN0aW9uKHBsYXRmb3JtcykgewogICAgaWYgKCFwbGF0Zm9ybXMubGVuZ3RoKSB7CiAgICAgIHJldHVybiBlbXB0eVN0YXRlKHsgaWNvbjogJ2xheWVycycsIHRpdGxlOiAnTm8gcGxhdGZvcm0gYWN0aXZpdHknLCBtZXNzYWdlOiAnTm8gcG9zdHMgd2VyZSByZWNvcmRlZCBvbiBhbnkgcGxhdGZvcm0gaW4gZWl0aGVyIHdlZWsuJyB9KTsKICAgIH0KICAgIGNvbnN0IHRhYmxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGFibGUnKTsKICAgIHRhYmxlLmNsYXNzTmFtZSA9ICdkYXRhLXRhYmxlJzsKICAgIGNvbnN0IHRoZWFkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGhlYWQnKTsKICAgIGNvbnN0IGh0ciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICBodHIuYXBwZW5kKAogICAgICB0ZXh0RWwoJ3RoJywgJ1BsYXRmb3JtJyksCiAgICAgIHRleHRFbCgndGgnLCAnUG9zdHMgKFRoaXMgLyBMYXN0KScsICdudW0nKSwKICAgICAgdGV4dEVsKCd0aCcsICdSZWFjaCAoVGhpcyAvIExhc3QpJywgJ251bScpLAogICAgICB0ZXh0RWwoJ3RoJywgJ0VuZ2FnZW1lbnQgKFRoaXMgLyBMYXN0KScsICdudW0nKSwKICAgICAgdGV4dEVsKCd0aCcsICdFbmdhZ2VtZW50ICUgQ2hnJywgJ251bScpLAogICAgICB0ZXh0RWwoJ3RoJywgJ1RyZW5kJyksCiAgICApOwogICAgdGhlYWQuYXBwZW5kQ2hpbGQoaHRyKTsKICAgIGNvbnN0IHRib2R5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGJvZHknKTsKICAgIHBsYXRmb3Jtcy5mb3JFYWNoKChwKSA9PiB7CiAgICAgIGNvbnN0IHRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgY29uc3QgcFRkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgY29uc3QgcGlsbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsgcGlsbC5jbGFzc05hbWUgPSAncGxhdGZvcm0tcGlsbCc7CiAgICAgIGNvbnN0IGRvdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsgZG90LmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1kb3QnOyBkb3Quc3R5bGUuYmFja2dyb3VuZCA9IHAuY29sb3I7CiAgICAgIHBpbGwuYXBwZW5kKGRvdCwgZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUocC5sYWJlbCkpOwogICAgICBwVGQuYXBwZW5kQ2hpbGQocGlsbCk7CiAgICAgIGNvbnN0IGRpZmYgPSBwLmVuZ0N1ciAtIHAuZW5nUHJldjsKICAgICAgdHIuYXBwZW5kKAogICAgICAgIHBUZCwKICAgICAgICB0ZXh0RWwoJ3RkJywgRm9ybWF0Lm51bWJlcihwLnBvc3RzQ3VyKSArICcgLyAnICsgRm9ybWF0Lm51bWJlcihwLnBvc3RzUHJldiksICdudW0nKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgRm9ybWF0LnNtYXJ0KHAucmVhY2hDdXIpICsgJyAvICcgKyBGb3JtYXQuc21hcnQocC5yZWFjaFByZXYpLCAnbnVtJyksCiAgICAgICAgdGV4dEVsKCd0ZCcsIEZvcm1hdC5zbWFydChwLmVuZ0N1cikgKyAnIC8gJyArIEZvcm1hdC5zbWFydChwLmVuZ1ByZXYpLCAnbnVtJyksCiAgICAgICAgdGV4dEVsKCd0ZCcsIHAuZW5nUGN0ID09PSBudWxsID8gJ24vYScgOiBwY3RUZXh0KHAuZW5nUGN0KSwgJ251bSAnICsgdHJlbmRDbGFzcyhkaWZmKSksCiAgICAgICAgdGV4dEVsKCd0ZCcsIGFycm93Rm9yKGRpZmYpLCB0cmVuZENsYXNzKGRpZmYpKSwKICAgICAgKTsKICAgICAgdGJvZHkuYXBwZW5kQ2hpbGQodHIpOwogICAgfSk7CiAgICB0YWJsZS5hcHBlbmQodGhlYWQsIHRib2R5KTsKICAgIHJldHVybiB0YWJsZTsKICB9CgogIGZ1bmN0aW9uIHRvcFBvc3RzU2VjdGlvbihwb3N0cykgewogICAgaWYgKCFwb3N0cy5sZW5ndGgpIHsKICAgICAgcmV0dXJuIGVtcHR5U3RhdGUoeyBpY29uOiAndHJvcGh5JywgdGl0bGU6ICdObyByYW5rZWQgcG9zdHMgdGhpcyB3ZWVrJywgbWVzc2FnZTogJ05vIHBvc3RzIGluIHRoZSBzZWxlY3RlZCB3ZWVrIGhhdmUgYSB2YWx1ZSBmb3IgdGhlIGNob3NlbiByYW5raW5nIG1ldHJpYy4nIH0pOwogICAgfQogICAgY29uc3QgcGxhdGZvcm1PcHRpb25zID0gKHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZSB8fCB7IHBsYXRmb3JtczogW10gfSkucGxhdGZvcm1zIHx8IFtdOwogICAgY29uc3QgbGFiZWwgPSAoUkFOS19NRVRSSUNTLmZpbmQoKHgpID0+IHgua2V5ID09PSByYW5rTWV0cmljKSB8fCB7IGxhYmVsOiByYW5rTWV0cmljIH0pLmxhYmVsOwogICAgY29uc3QgdGFibGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0YWJsZScpOwogICAgdGFibGUuY2xhc3NOYW1lID0gJ2RhdGEtdGFibGUnOwogICAgY29uc3QgdGhlYWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0aGVhZCcpOwogICAgY29uc3QgaHRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgIGh0ci5hcHBlbmQoCiAgICAgIHRleHRFbCgndGgnLCAnUmFuaycpLCB0ZXh0RWwoJ3RoJywgJ0RhdGUnKSwgdGV4dEVsKCd0aCcsICdQbGF0Zm9ybScpLCB0ZXh0RWwoJ3RoJywgJ0NhbXBhaWduJyksCiAgICAgIHRleHRFbCgndGgnLCAnQ29udGVudCBUeXBlJyksIHRleHRFbCgndGgnLCAnQ2FwdGlvbicpLCB0ZXh0RWwoJ3RoJywgbGFiZWwsICdudW0nKSwKICAgICk7CiAgICB0aGVhZC5hcHBlbmRDaGlsZChodHIpOwogICAgY29uc3QgdGJvZHkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0Ym9keScpOwogICAgcG9zdHMuZm9yRWFjaCgocCwgaSkgPT4gewogICAgICBjb25zdCB0ciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAgIGNvbnN0IG1ldGEgPSBwbGF0Zm9ybU9wdGlvbnMuZmluZCgocGwpID0+IHBsLmlkID09PSBwLnBsYXRmb3JtKSB8fCB7IGxhYmVsOiBwLnBsYXRmb3JtLCBjb2xvcjogJyM5OTknIH07CiAgICAgIGNvbnN0IHBUZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIGNvbnN0IHBpbGwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7IHBpbGwuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLXBpbGwnOwogICAgICBjb25zdCBkb3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7IGRvdC5jbGFzc05hbWUgPSAncGxhdGZvcm0tZG90JzsgZG90LnN0eWxlLmJhY2tncm91bmQgPSBtZXRhLmNvbG9yOwogICAgICBwaWxsLmFwcGVuZChkb3QsIGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKG1ldGEubGFiZWwpKTsKICAgICAgcFRkLmFwcGVuZENoaWxkKHBpbGwpOwogICAgICBjb25zdCBjYXB0aW9uID0gcC5jYXB0aW9uIHx8ICcobm8gY2FwdGlvbiknOwogICAgICBjb25zdCBzaG9ydCA9IGNhcHRpb24ubGVuZ3RoID4gNzAgPyBjYXB0aW9uLnNsaWNlKDAsIDcwKSArICfigKYnIDogY2FwdGlvbjsKICAgICAgY29uc3QgY2FwVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBpZiAocC5wb3N0aW5nX2xpbmspIHsKICAgICAgICBjb25zdCBhID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYScpOwogICAgICAgIGEuY2xhc3NOYW1lID0gJ2NhcHRpb24tbGluayc7IGEuaHJlZiA9IHAucG9zdGluZ19saW5rOyBhLnRhcmdldCA9ICdfYmxhbmsnOyBhLnJlbCA9ICdub29wZW5lciBub3JlZmVycmVyJzsKICAgICAgICBhLnRpdGxlID0gY2FwdGlvbjsgYS5hcHBlbmRDaGlsZChkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShzaG9ydCkpOwogICAgICAgIGNhcFRkLmFwcGVuZENoaWxkKGEpOwogICAgICB9IGVsc2UgewogICAgICAgIGNhcFRkLmFwcGVuZENoaWxkKGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKHNob3J0KSk7CiAgICAgICAgY2FwVGQudGl0bGUgPSBjYXB0aW9uOwogICAgICB9CiAgICAgIHRyLmFwcGVuZCgKICAgICAgICB0ZXh0RWwoJ3RkJywgJyMnICsgKGkgKyAxKSksCiAgICAgICAgdGV4dEVsKCd0ZCcsIEZvcm1hdC5kYXRlKHAucHVibGlzaF9kYXRlKSksCiAgICAgICAgcFRkLAogICAgICAgIHRleHRFbCgndGQnLCBwLmNhbXBhaWduX3R5cGUgfHwgJ+KAlCcpLAogICAgICAgIHRleHRFbCgndGQnLCBwLmNvbnRlbnRfdHlwZSB8fCAn4oCUJyksCiAgICAgICAgY2FwVGQsCiAgICAgICAgdGV4dEVsKCd0ZCcsIHAubWV0cmljX3ZhbHVlID09PSBudWxsIHx8IHAubWV0cmljX3ZhbHVlID09PSB1bmRlZmluZWQgPyAn4oCUJyA6IEZvcm1hdC5zbWFydChwLm1ldHJpY192YWx1ZSksICdudW0nKSwKICAgICAgKTsKICAgICAgdGJvZHkuYXBwZW5kQ2hpbGQodHIpOwogICAgfSk7CiAgICB0YWJsZS5hcHBlbmQodGhlYWQsIHRib2R5KTsKICAgIHJldHVybiB0YWJsZTsKICB9CgogIGZ1bmN0aW9uIHdyaXR0ZW5TdW1tYXJ5KG0pIHsKICAgIGNvbnN0IGJveCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYm94LmNsYXNzTmFtZSA9ICdjYXJkIHJlcG9ydC1zdW1tYXJ5JzsKICAgIGNvbnN0IGdldCA9IChrKSA9PiBtLnJvd3MuZmluZCgocikgPT4gci5rZXkgPT09IGspOwogICAgY29uc3QgcmVhY2ggPSBnZXQoJ3JlYWNoJyk7IGNvbnN0IGVuZyA9IGdldCgnZW5nYWdlbWVudCcpOyBjb25zdCBpbXByID0gZ2V0KCdpbXByZXNzaW9ucycpOwogICAgY29uc3QgZm9sID0gZ2V0KCdmb2xsb3dlcnMnKTsgY29uc3QgZ3JvdyA9IGdldCgnZm9sbG93ZXJzX2dyb3d0aCcpOyBjb25zdCBwb3N0cyA9IGdldCgncG9zdHMnKTsKCiAgICAvLyBIZWFkbGluZSBkaXJlY3Rpb246IGEgdm90ZSBhY3Jvc3MgdGhlIG1ldHJpY3MgdGhhdCBoYXZlIGEgcmVhbCBiYXNlbGluZS4KICAgIGNvbnN0IHZvdGVycyA9IFtyZWFjaCwgZW5nLCBpbXByLCBncm93XS5maWx0ZXIoKHIpID0+IHIuaGFzRGF0YSAmJiByLnBjdCAhPT0gbnVsbCAmJiByLmRpZmYgIT09IDApOwogICAgY29uc3QgdXBzID0gdm90ZXJzLmZpbHRlcigocikgPT4gci5kaWZmID4gMCkubGVuZ3RoOwogICAgY29uc3QgZG93bnMgPSB2b3RlcnMuZmlsdGVyKChyKSA9PiByLmRpZmYgPCAwKS5sZW5ndGg7CiAgICBsZXQgaGVhZGxpbmU7CiAgICBpZiAoIXZvdGVycy5sZW5ndGgpIGhlYWRsaW5lID0gJ1RoZXJlIGlzIG5vdCBlbm91Z2ggY29tcGFyYWJsZSBkYXRhIHRoaXMgd2VlayB0byBqdWRnZSB0aGUgb3ZlcmFsbCBkaXJlY3Rpb24gb2YgcGVyZm9ybWFuY2UgYWdhaW5zdCB0aGUgcHJldmlvdXMgd2Vlay4nOwogICAgZWxzZSBpZiAodXBzID4gZG93bnMpIGhlYWRsaW5lID0gJ1RoaXMgd2Vlaywgc29jaWFsIG1lZGlhIHBlcmZvcm1hbmNlIGltcHJvdmVkIGNvbXBhcmVkIHdpdGggdGhlIHByZXZpb3VzIHdlZWsuJzsKICAgIGVsc2UgaWYgKGRvd25zID4gdXBzKSBoZWFkbGluZSA9ICdUaGlzIHdlZWssIHNvY2lhbCBtZWRpYSBwZXJmb3JtYW5jZSBkZWNsaW5lZCBjb21wYXJlZCB3aXRoIHRoZSBwcmV2aW91cyB3ZWVrLic7CiAgICBlbHNlIGhlYWRsaW5lID0gJ1RoaXMgd2Vlaywgc29jaWFsIG1lZGlhIHBlcmZvcm1hbmNlIHdhcyBtaXhlZCBjb21wYXJlZCB3aXRoIHRoZSBwcmV2aW91cyB3ZWVrIOKAlCBzb21lIG1ldHJpY3Mgcm9zZSB3aGlsZSBvdGhlcnMgZmVsbC4nOwoKICAgIGNvbnN0IGgxID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaDQnKTsgaDEudGV4dENvbnRlbnQgPSAnV2Vla2x5IFBlcmZvcm1hbmNlIFN1bW1hcnknOwogICAgY29uc3QgcDEgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdwJyk7IHAxLnRleHRDb250ZW50ID0gaGVhZGxpbmU7CiAgICBib3guYXBwZW5kKGgxLCBwMSk7CgogICAgY29uc3QgaDIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdoNCcpOyBoMi50ZXh0Q29udGVudCA9ICdLZXkgSW5zaWdodHMnOwogICAgY29uc3QgdWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd1bCcpOwogICAgY29uc3QgYWRkSW5zaWdodCA9ICh0ZXh0LCBkaWZmKSA9PiB7CiAgICAgIGNvbnN0IGxpID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnbGknKTsKICAgICAgaWYgKGRpZmYgIT09IHVuZGVmaW5lZCAmJiBkaWZmICE9PSBudWxsICYmIGRpZmYgIT09IDApIHsKICAgICAgICBjb25zdCBzcGFuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICAgIHNwYW4uY2xhc3NOYW1lID0gZGlmZiA+IDAgPyAndHJlbmQtdXAnIDogJ3RyZW5kLWRvd24nOwogICAgICAgIHNwYW4udGV4dENvbnRlbnQgPSBkaWZmID4gMCA/ICfilrIgJyA6ICfilrwgJzsKICAgICAgICBsaS5hcHBlbmRDaGlsZChzcGFuKTsKICAgICAgfQogICAgICBsaS5hcHBlbmRDaGlsZChkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZSh0ZXh0KSk7CiAgICAgIHVsLmFwcGVuZENoaWxkKGxpKTsKICAgIH07CiAgICBbcmVhY2gsIGVuZywgaW1wcl0uZm9yRWFjaCgocikgPT4gewogICAgICBpZiAoIXIuaGFzRGF0YSkgeyBhZGRJbnNpZ2h0KHIubGFiZWwgKyAnOiBubyBwb3N0IGRhdGEgdGhpcyB3ZWVrLicpOyByZXR1cm47IH0KICAgICAgaWYgKHIucGN0ID09PSBudWxsKSB7IGFkZEluc2lnaHQoci5sYWJlbCArICcgd2FzICcgKyBmbXRWYWx1ZShyLmZtdCwgci5jdXIpICsgJyB0aGlzIHdlZWsgKG5vdGhpbmcgcmVjb3JkZWQgbGFzdCB3ZWVrIHRvIGNvbXBhcmUgYWdhaW5zdCkuJywgci5kaWZmKTsgcmV0dXJuOyB9CiAgICAgIGlmIChyLmRpZmYgPT09IDApIHsgYWRkSW5zaWdodChyLmxhYmVsICsgJyBoZWxkIGZsYXQgdmVyc3VzIGxhc3Qgd2VlayBhdCAnICsgZm10VmFsdWUoci5mbXQsIHIuY3VyKSArICcuJyk7IHJldHVybjsgfQogICAgICBjb25zdCBhYnNQY3QgPSBNYXRoLmFicyhyLnBjdCkudG9GaXhlZCgxKS5yZXBsYWNlKC9cLjAkLywgJycpICsgJyUnOwogICAgICBhZGRJbnNpZ2h0KHIubGFiZWwgKyAnICcgKyB3b3JkRm9yKHIuZGlmZikgKyAnIGJ5ICcgKyBhYnNQY3QgKyAnICgnICsgZm10VmFsdWUoci5mbXQsIHIucHJldikgKyAnIOKGkiAnICsgZm10VmFsdWUoci5mbXQsIHIuY3VyKSArICcpLicsIHIuZGlmZik7CiAgICB9KTsKICAgIGlmIChmb2wuaGFzRGF0YSkgewogICAgICBpZiAoZ3Jvdy5oYXNEYXRhICYmIGdyb3cuY3VyICE9PSAwKSB7CiAgICAgICAgYWRkSW5zaWdodCgnRm9sbG93ZXJzICcgKyAoZ3Jvdy5jdXIgPiAwID8gJ2dyZXcnIDogJ2ZlbGwnKSArICcgYnkgJyArIGZtdERpZmYoZ3Jvdy5mbXQsIGdyb3cuY3VyKS5yZXBsYWNlKCcrJywgJycpLnJlcGxhY2UoJ+KIkicsICcnKSArICcgdGhpcyB3ZWVrLCBlbmRpbmcgYXQgJyArIGZtdFZhbHVlKGZvbC5mbXQsIGZvbC5jdXIpICsgJy4nLCBncm93LmN1cik7CiAgICAgIH0gZWxzZSB7CiAgICAgICAgYWRkSW5zaWdodCgnRm9sbG93ZXJzIHdlcmUgdW5jaGFuZ2VkIHRoaXMgd2VlayBhdCAnICsgZm10VmFsdWUoZm9sLmZtdCwgZm9sLmN1cikgKyAnLicpOwogICAgICB9CiAgICB9CiAgICBhZGRJbnNpZ2h0KCdQb3N0cyBwdWJsaXNoZWQ6ICcgKyBGb3JtYXQubnVtYmVyKHBvc3RzLmN1cikgKyAnIHRoaXMgd2VlayB2ZXJzdXMgJyArIEZvcm1hdC5udW1iZXIocG9zdHMucHJldikgKyAnIGxhc3Qgd2Vlay4nLCBwb3N0cy5kaWZmKTsKICAgIGNvbnN0IGJlc3QgPSBtLnBsYXRmb3Jtcy5maW5kKChwKSA9PiBwLmVuZ0N1ciA+IDApOwogICAgaWYgKGJlc3QpIGFkZEluc2lnaHQoJ1RoZSBiZXN0LXBlcmZvcm1pbmcgcGxhdGZvcm0gd2FzICcgKyBiZXN0LmxhYmVsICsgJyB3aXRoICcgKyBGb3JtYXQuc21hcnQoYmVzdC5lbmdDdXIpICsgJyBlbmdhZ2VtZW50LicpOwogICAgaWYgKG0udG9wUG9zdHMubGVuZ3RoKSB7CiAgICAgIGNvbnN0IHQgPSBtLnRvcFBvc3RzWzBdOwogICAgICBjb25zdCBjYXAgPSAodC5jYXB0aW9uIHx8ICcobm8gY2FwdGlvbiknKS5zbGljZSgwLCA4MCk7CiAgICAgIGNvbnN0IGxibCA9IChSQU5LX01FVFJJQ1MuZmluZCgoeCkgPT4geC5rZXkgPT09IHJhbmtNZXRyaWMpIHx8IHsgbGFiZWw6IHJhbmtNZXRyaWMgfSkubGFiZWw7CiAgICAgIGNvbnN0IHRwID0gKCh3aW5kb3cuX19maWx0ZXJPcHRpb25zQ2FjaGUgfHwgeyBwbGF0Zm9ybXM6IFtdIH0pLnBsYXRmb3JtcyB8fCBbXSkuZmluZCgoeCkgPT4geC5pZCA9PT0gdC5wbGF0Zm9ybSk7CiAgICAgIGNvbnN0IHRwTGFiZWwgPSB0cCA/IHRwLmxhYmVsIDogKHQucGxhdGZvcm0gfHwgJ3Vua25vd24nKTsKICAgICAgYWRkSW5zaWdodCgnVGhlIHRvcC1wZXJmb3JtaW5nIHBvc3Qgd2FzICInICsgY2FwICsgJyIgb24gJyArIHRwTGFiZWwgKyAnICgnICsgbGJsICsgJzogJyArICh0Lm1ldHJpY192YWx1ZSA9PSBudWxsID8gJ+KAlCcgOiBGb3JtYXQuc21hcnQodC5tZXRyaWNfdmFsdWUpKSArICcpLicpOwogICAgfQogICAgYm94LmFwcGVuZChoMiwgdWwpOwoKICAgIGNvbnN0IGgzID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaDQnKTsgaDMudGV4dENvbnRlbnQgPSAnUmVjb21tZW5kYXRpb25zJzsKICAgIGNvbnN0IHJsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndWwnKTsKICAgIGNvbnN0IHJlY3MgPSBbXTsKICAgIGlmIChlbmcuaGFzRGF0YSAmJiBlbmcucGN0ICE9PSBudWxsICYmIGVuZy5wY3QgPD0gLTEwKSByZWNzLnB1c2goJ0VuZ2FnZW1lbnQgZmVsbCAnICsgYWJzUGN0VGV4dChlbmcucGN0KSArICcuIFJldmlzaXQgdGhlIHBvc3QgZm9ybWF0cyBhbmQgcHVibGlzaCB0aW1lcyB0aGF0IHBlcmZvcm1lZCB3ZWxsIGluIGVhcmxpZXIgd2Vla3MgYW5kIGxlYW4gaW50byB0aGVtLicpOwogICAgaWYgKHJlYWNoLmhhc0RhdGEgJiYgcmVhY2gucGN0ICE9PSBudWxsICYmIHJlYWNoLnBjdCA8PSAtMTApIHJlY3MucHVzaCgnUmVhY2ggZHJvcHBlZCAnICsgYWJzUGN0VGV4dChyZWFjaC5wY3QpICsgJy4gSW5jcmVhc2luZyBwb3N0aW5nIGZyZXF1ZW5jeSBvciBwcmlvcml0aXNpbmcgc2hhcmVhYmxlIGZvcm1hdHMgKHNob3J0IHZpZGVvLCBjYXJvdXNlbHMpIGNhbiBoZWxwIHJlY292ZXIgaXQuJyk7CiAgICBpZiAocG9zdHMuaGFzRGF0YSAmJiBwb3N0cy5kaWZmIDwgMCkgcmVjcy5wdXNoKCdZb3UgcHVibGlzaGVkIGZld2VyIHBvc3RzIHRoaXMgd2VlayAoJyArIEZvcm1hdC5udW1iZXIocG9zdHMuY3VyKSArICcgdnMgJyArIEZvcm1hdC5udW1iZXIocG9zdHMucHJldikgKyAnKS4gSG9sZGluZyBhIHN0ZWFkeSBjYWRlbmNlIHN1cHBvcnRzIGJvdGggcmVhY2ggYW5kIGVuZ2FnZW1lbnQuJyk7CiAgICBpZiAoZ3Jvdy5oYXNEYXRhICYmIGdyb3cuY3VyIDw9IDApIHJlY3MucHVzaCgnRm9sbG93ZXIgZ3Jvd3RoIHN0YWxsZWQgdGhpcyB3ZWVrLiBDb25zaWRlciBhIGZvbGxvd2VyLWZvY3VzZWQgY2FtcGFpZ24gb3IgYW4gZXhwbGljaXQgY2FsbC10by1mb2xsb3cgaW4geW91ciBoaWdoZXN0LXJlYWNoIHBvc3RzLicpOwogICAgaWYgKG0ucGxhdGZvcm1zLmxlbmd0aCA+PSAyKSB7CiAgICAgIGNvbnN0IHRvcCA9IG0ucGxhdGZvcm1zWzBdOwogICAgICBjb25zdCBhY3RpdmUgPSBtLnBsYXRmb3Jtcy5maWx0ZXIoKHApID0+IHAuZW5nQ3VyID4gMCk7CiAgICAgIGNvbnN0IHdlYWtlc3QgPSBhY3RpdmUubGVuZ3RoID49IDIgPyBhY3RpdmVbYWN0aXZlLmxlbmd0aCAtIDFdIDogbnVsbDsKICAgICAgaWYgKHRvcCAmJiB3ZWFrZXN0ICYmIHRvcC5lbmdDdXIgPj0gd2Vha2VzdC5lbmdDdXIgKiAzKSB7CiAgICAgICAgcmVjcy5wdXNoKHRvcC5sYWJlbCArICcgaXMgY2FycnlpbmcgbW9zdCBvZiB0aGUgZW5nYWdlbWVudC4gVGVzdCByZXB1cnBvc2luZyBpdHMgYmVzdCBjb250ZW50IG9uICcgKyB3ZWFrZXN0LmxhYmVsICsgJyB0byBsaWZ0IHRoZSB3ZWFrZXIgY2hhbm5lbHMuJyk7CiAgICAgIH0KICAgIH0KICAgIGlmIChyZWFjaC5oYXNEYXRhICYmIGVuZy5oYXNEYXRhICYmIGltcHIuaGFzRGF0YSAmJiAocmVhY2gucGN0IHx8IDApID4gMCAmJiAoZW5nLnBjdCB8fCAwKSA+IDAgJiYgKGltcHIucGN0IHx8IDApID4gMCkgewogICAgICByZWNzLnB1c2goJ1JlYWNoLCBlbmdhZ2VtZW50IGFuZCBpbXByZXNzaW9ucyBhbGwgcm9zZSB0aGlzIHdlZWsg4oCUIGtlZXAgdGhlIGN1cnJlbnQgY29udGVudCBtaXggYW5kIHBvc3RpbmcgY2FkZW5jZS4nKTsKICAgIH0KICAgIGlmICghcmVjcy5sZW5ndGgpIHJlY3MucHVzaCgnTm8gbWV0cmljIG1vdmVkIHNoYXJwbHkgdGhpcyB3ZWVrLiBNYWludGFpbiB0aGUgY3VycmVudCBhcHByb2FjaCBhbmQga2VlcCByZXZpZXdpbmcgcGVyZm9ybWFuY2Ugd2VlayBvdmVyIHdlZWsuJyk7CiAgICByZWNzLmZvckVhY2goKHQpID0+IHsgY29uc3QgbGkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdsaScpOyBsaS50ZXh0Q29udGVudCA9IHQ7IHJsLmFwcGVuZENoaWxkKGxpKTsgfSk7CiAgICBib3guYXBwZW5kKGgzLCBybCk7CiAgICByZXR1cm4gYm94OwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyUmVwb3J0KCkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdyZXBvcnRSZXN1bHRzJyk7CiAgICBpZiAoIXdyYXAgfHwgIWxhc3RNb2RlbCkgcmV0dXJuOwogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgIGNvbnN0IG0gPSBsYXN0TW9kZWw7CgogICAgY29uc3QgaGVhZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgaGVhZC5jbGFzc05hbWUgPSAnY2FyZCByZXBvcnQtZG9jLWhlYWQnOwogICAgaGVhZC5hcHBlbmQoCiAgICAgIHRleHRFbCgnZGl2JywgJ1dlZWtseSBTb2NpYWwgTWVkaWEgQW5hbHl0aWNzIFJlcG9ydCcsICdyZXBvcnQtdGl0bGUnKSwKICAgICAgdGV4dEVsKCdkaXYnLCBGb3JtYXQuZGF0ZShtLmN1ci5mcm9tKSArICcg4oCTICcgKyBGb3JtYXQuZGF0ZShtLmN1ci50bykKICAgICAgICArICcgICB2cyAgICcgKyBGb3JtYXQuZGF0ZShtLnByZXYuZnJvbSkgKyAnIOKAkyAnICsgRm9ybWF0LmRhdGUobS5wcmV2LnRvKSwgJ3JlcG9ydC1yYW5nZScpLAogICAgICB0ZXh0RWwoJ2RpdicsICdHZW5lcmF0ZWQgJyArIG5ldyBEYXRlKCkudG9Mb2NhbGVTdHJpbmcoJ2VuLVVTJywgeyBkYXRlU3R5bGU6ICdtZWRpdW0nLCB0aW1lU3R5bGU6ICdzaG9ydCcgfSksICdyZXBvcnQtZ2VuZXJhdGVkJyksCiAgICApOwogICAgd3JhcC5hcHBlbmRDaGlsZChoZWFkKTsKCiAgICBpZiAoIW0uYW55RGF0YSkgewogICAgICB3cmFwLmFwcGVuZENoaWxkKGVtcHR5U3RhdGUoewogICAgICAgIGljb246ICdmaWxlLXRleHQnLAogICAgICAgIHRpdGxlOiAnTm8gZGF0YSBmb3IgdGhlIHNlbGVjdGVkIHdlZWsnLAogICAgICAgIG1lc3NhZ2U6ICdUaGVyZSBhcmUgbm8gcG9zdHMgb3IgRm9sbG93ZXJzIERhdGEgUmVjb3JkIGVudHJpZXMgaW4gdGhpcyBkYXRlIHJhbmdlLiBVcGxvYWQgYSB3ZWVrbHkgZXhwb3J0IG9yIHBpY2sgYW5vdGhlciB3ZWVrLicsCiAgICAgICAgYWN0aW9uTGFiZWw6ICdVcGxvYWQgZGF0YScsCiAgICAgICAgb25BY3Rpb246ICgpID0+IHsgY29uc3QgYiA9IGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3IoJy50YWItYnRuW2RhdGEtdGFiPSJ1cGxvYWQiXScpOyBpZiAoYikgYi5jbGljaygpOyB9LAogICAgICB9KSk7CiAgICAgIHJldHVybjsKICAgIH0KCiAgICB3cmFwLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgJ01ldHJpYyBDb21wYXJpc29uJywgJ3NlY3Rpb24tdGl0bGUnKSk7CiAgICBjb25zdCBncmlkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBncmlkLmNsYXNzTmFtZSA9ICdzdGF0LWdyaWQnOwogICAgbS5yb3dzLmZvckVhY2goKHIpID0+IGdyaWQuYXBwZW5kQ2hpbGQobWV0cmljVGlsZShyKSkpOwogICAgd3JhcC5hcHBlbmRDaGlsZChncmlkKTsKICAgIHdyYXAuYXBwZW5kQ2hpbGQobWV0cmljVGFibGUobS5yb3dzKSk7CgogICAgd3JhcC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsICdUaGlzIFdlZWsgdnMgTGFzdCBXZWVrJywgJ3NlY3Rpb24tdGl0bGUnKSk7CiAgICBjb25zdCBjaGFydENhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNoYXJ0Q2FyZC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICBjb25zdCBzZWxSb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHNlbFJvdy5jbGFzc05hbWUgPSAncmVwb3J0LW1ldHJpYy1zZWxlY3QnOwogICAgY29uc3QgY1NlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NlbGVjdCcpOwogICAgQ0hBUlRfTUVUUklDUy5mb3JFYWNoKChjKSA9PiB7CiAgICAgIGNvbnN0IG8gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsKICAgICAgby52YWx1ZSA9IGMua2V5OyBvLnRleHRDb250ZW50ID0gYy5sYWJlbDsKICAgICAgaWYgKGMua2V5ID09PSBjaGFydE1ldHJpYykgby5zZWxlY3RlZCA9IHRydWU7CiAgICAgIGNTZWwuYXBwZW5kQ2hpbGQobyk7CiAgICB9KTsKICAgIGNTZWwuYWRkRXZlbnRMaXN0ZW5lcignY2hhbmdlJywgKCkgPT4geyBjaGFydE1ldHJpYyA9IGNTZWwudmFsdWU7IHJlbmRlckNvbXBhcmVDaGFydCgpOyB9KTsKICAgIHNlbFJvdy5hcHBlbmQodGV4dEVsKCdsYWJlbCcsICdNZXRyaWMnKSwgY1NlbCk7CiAgICBjaGFydENhcmQuYXBwZW5kQ2hpbGQoc2VsUm93KTsKICAgIGNvbnN0IGNoYXJ0V3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY2hhcnRXcmFwLmNsYXNzTmFtZSA9ICdjaGFydC13cmFwJzsKICAgIGNoYXJ0V3JhcC5pZCA9ICdyZXBvcnRDaGFydFdyYXAnOwogICAgY2hhcnRDYXJkLmFwcGVuZENoaWxkKGNoYXJ0V3JhcCk7CiAgICB3cmFwLmFwcGVuZENoaWxkKGNoYXJ0Q2FyZCk7CiAgICByZW5kZXJDb21wYXJlQ2hhcnQoKTsKCiAgICB3cmFwLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgJ1BsYXRmb3JtIFBlcmZvcm1hbmNlJywgJ3NlY3Rpb24tdGl0bGUnKSk7CiAgICB3cmFwLmFwcGVuZENoaWxkKHBsYXRmb3JtU2VjdGlvbihtLnBsYXRmb3JtcykpOwoKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCAnVG9wLVBlcmZvcm1pbmcgQ29udGVudCDigJQgYnkgJyArIChSQU5LX01FVFJJQ1MuZmluZCgoeCkgPT4geC5rZXkgPT09IHJhbmtNZXRyaWMpIHx8IHsgbGFiZWw6IHJhbmtNZXRyaWMgfSkubGFiZWwsICdzZWN0aW9uLXRpdGxlJykpOwogICAgd3JhcC5hcHBlbmRDaGlsZCh0b3BQb3N0c1NlY3Rpb24obS50b3BQb3N0cykpOwoKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCAnQXV0by1HZW5lcmF0ZWQgU3VtbWFyeScsICdzZWN0aW9uLXRpdGxlJykpOwogICAgd3JhcC5hcHBlbmRDaGlsZCh3cml0dGVuU3VtbWFyeShtKSk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiByZW5kZXIoKSB7CiAgICByb290ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3ZpZXctcmVwb3J0Jyk7CiAgICBzaGVsbCgpOwogICAgaWYgKGxhc3RNb2RlbCkgcmVuZGVyUmVwb3J0KCk7CiAgICBlbHNlIGF3YWl0IGdlbmVyYXRlKCk7CiAgfQoKICByZXR1cm4geyByZW5kZXIgfTsKfSkoKTsKCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBBcHAgYm9vdHN0cmFwOiB0YWIgcm91dGluZywgZmlsdGVyIGJhciB3aXJpbmcsIHRoZW1lIHRvZ2dsZS4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCigoKSA9PiB7CiAgY29uc3QgVklFV1MgPSB7CiAgICBkYXNoYm9hcmQ6IERhc2hib2FyZCwKICAgIHJlY29yZHM6IFJlY29yZHMsCiAgICBmb2xsb3dlcnM6IEZvbGxvd2VycywKICAgIGNvbXBhcmlzb246IENvbXBhcmlzb24sCiAgICByZXBvcnQ6IFJlcG9ydCwKICAgIHVwbG9hZDogVXBsb2FkLAogICAgaGlzdG9yeTogSGlzdG9yeSwKICB9OwoKICBsZXQgYWN0aXZlVGFiID0gJ2Rhc2hib2FyZCc7CgogIGZ1bmN0aW9uIHN3aXRjaFRhYih0YWIpIHsKICAgIGFjdGl2ZVRhYiA9IHRhYjsKICAgIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJy50YWItYnRuJykuZm9yRWFjaCgoYnRuKSA9PiB7CiAgICAgIGNvbnN0IGlzQWN0aXZlID0gYnRuLmRhdGFzZXQudGFiID09PSB0YWI7CiAgICAgIGJ0bi5jbGFzc0xpc3QudG9nZ2xlKCdpcy1hY3RpdmUnLCBpc0FjdGl2ZSk7CiAgICAgIGJ0bi5zZXRBdHRyaWJ1dGUoJ2FyaWEtc2VsZWN0ZWQnLCBTdHJpbmcoaXNBY3RpdmUpKTsKICAgIH0pOwogICAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnLnZpZXcnKS5mb3JFYWNoKCh2aWV3KSA9PiB7CiAgICAgIHZpZXcuY2xhc3NMaXN0LnRvZ2dsZSgnaXMtYWN0aXZlJywgdmlldy5pZCA9PT0gYHZpZXctJHt0YWJ9YCk7CiAgICB9KTsKICAgIC8vIEZpbHRlcnMgYXBwbHkgdG8gRGFzaGJvYXJkIGFuZCBEYXRhIFJlY29yZHMgKENvbXBhcmlzb25zIGhhcyBpdHMgb3duIHJhbmdlIGNvbnRyb2xzKS4KICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJCYXInKS5zdHlsZS5kaXNwbGF5ID0gKHRhYiA9PT0gJ2Rhc2hib2FyZCcgfHwgdGFiID09PSAncmVjb3JkcycpID8gJ2ZsZXgnIDogJ25vbmUnOwogICAgcmVuZGVyQWN0aXZlVmlldygpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyQWN0aXZlVmlldygpIHsKICAgIGNvbnN0IHZpZXcgPSBWSUVXU1thY3RpdmVUYWJdOwogICAgaWYgKHZpZXcgJiYgdmlldy5yZW5kZXIpIHZpZXcucmVuZGVyKCk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiBsb2FkRmlsdGVyT3B0aW9ucygpIHsKICAgIGNvbnN0IG9wdGlvbnMgPSBhd2FpdCBBcGkuZmlsdGVyT3B0aW9ucygpOwogICAgd2luZG93Ll9fZmlsdGVyT3B0aW9uc0NhY2hlID0gb3B0aW9uczsKCiAgICBjb25zdCBwbGF0Zm9ybVNlbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJQbGF0Zm9ybScpOwogICAgcGxhdGZvcm1TZWwubGVuZ3RoID0gMTsKICAgIG9wdGlvbnMucGxhdGZvcm1zLmZvckVhY2goKHApID0+IHsKICAgICAgY29uc3Qgb3B0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7CiAgICAgIG9wdC52YWx1ZSA9IHAuaWQ7CiAgICAgIG9wdC50ZXh0Q29udGVudCA9IHAubGFiZWw7CiAgICAgIHBsYXRmb3JtU2VsLmFwcGVuZENoaWxkKG9wdCk7CiAgICB9KTsKCiAgICBjb25zdCBjYW1wYWlnblNlbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJDYW1wYWlnbicpOwogICAgY2FtcGFpZ25TZWwubGVuZ3RoID0gMTsKICAgIG9wdGlvbnMuY2FtcGFpZ25UeXBlcy5mb3JFYWNoKChjKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOwogICAgICBvcHQudmFsdWUgPSBjOwogICAgICBvcHQudGV4dENvbnRlbnQgPSBjOwogICAgICBjYW1wYWlnblNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CgogICAgY29uc3QgY29udGVudFNlbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJDb250ZW50VHlwZScpOwogICAgY29udGVudFNlbC5sZW5ndGggPSAxOwogICAgb3B0aW9ucy5jb250ZW50VHlwZXMuZm9yRWFjaCgoYykgPT4gewogICAgICBjb25zdCBvcHQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsKICAgICAgb3B0LnZhbHVlID0gYzsKICAgICAgb3B0LnRleHRDb250ZW50ID0gYzsKICAgICAgY29udGVudFNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgfQoKICBmdW5jdGlvbiB3aXJlRmlsdGVyQmFyKCkgewogICAgY29uc3QgZGF0ZUZyb20gPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmlsdGVyRGF0ZUZyb20nKTsKICAgIGNvbnN0IGRhdGVUbyA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJEYXRlVG8nKTsKICAgIGNvbnN0IHBsYXRmb3JtID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZpbHRlclBsYXRmb3JtJyk7CiAgICBjb25zdCBjYW1wYWlnbiA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJDYW1wYWlnbicpOwogICAgY29uc3QgY29udGVudFR5cGUgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmlsdGVyQ29udGVudFR5cGUnKTsKICAgIGNvbnN0IGYgPSBTdGF0ZS5nZXRGaWx0ZXJzKCk7CiAgICBkYXRlRnJvbS52YWx1ZSA9IGYuZGF0ZUZyb207CiAgICBkYXRlVG8udmFsdWUgPSBmLmRhdGVUbzsKCiAgICBmdW5jdGlvbiBhcHBseSgpIHsKICAgICAgU3RhdGUuc2V0RmlsdGVycyh7CiAgICAgICAgZGF0ZUZyb206IGRhdGVGcm9tLnZhbHVlLAogICAgICAgIGRhdGVUbzogZGF0ZVRvLnZhbHVlLAogICAgICAgIHBsYXRmb3JtOiBwbGF0Zm9ybS52YWx1ZSwKICAgICAgICBjYW1wYWlnblR5cGU6IGNhbXBhaWduLnZhbHVlLAogICAgICAgIGNvbnRlbnRUeXBlOiBjb250ZW50VHlwZS52YWx1ZSwKICAgICAgfSk7CiAgICB9CiAgICBbZGF0ZUZyb20sIGRhdGVUbywgcGxhdGZvcm0sIGNhbXBhaWduLCBjb250ZW50VHlwZV0uZm9yRWFjaCgoZWwpID0+IGVsLmFkZEV2ZW50TGlzdGVuZXIoJ2NoYW5nZScsIGFwcGx5KSk7CgogICAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnI2ZpbHRlclByZXNldHMgYnV0dG9uJykuZm9yRWFjaCgoYnRuKSA9PiB7CiAgICAgIGJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsKICAgICAgICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcjZmlsdGVyUHJlc2V0cyBidXR0b24nKS5mb3JFYWNoKChiKSA9PiBiLmNsYXNzTGlzdC5yZW1vdmUoJ2lzLWFjdGl2ZScpKTsKICAgICAgICBidG4uY2xhc3NMaXN0LmFkZCgnaXMtYWN0aXZlJyk7CiAgICAgICAgY29uc3QgcHJlc2V0ID0gYnRuLmRhdGFzZXQucHJlc2V0OwogICAgICAgIGNvbnN0IHRvZGF5ID0gbmV3IERhdGUoKTsKICAgICAgICBjb25zdCB0byA9IHRvZGF5LnRvSVNPU3RyaW5nKCkuc2xpY2UoMCwgMTApOwogICAgICAgIGxldCBmcm9tOwogICAgICAgIGlmIChwcmVzZXQgPT09ICdhbGwnKSB7CiAgICAgICAgICBjb25zdCBtaW4gPSAod2luZG93Ll9fZmlsdGVyT3B0aW9uc0NhY2hlICYmIHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZS5kYXRlUmFuZ2UubWluKSB8fCB0bzsKICAgICAgICAgIGZyb20gPSBtaW47CiAgICAgICAgfSBlbHNlIHsKICAgICAgICAgIGNvbnN0IGQgPSBuZXcgRGF0ZSh0b2RheSk7CiAgICAgICAgICBkLnNldERhdGUoZC5nZXREYXRlKCkgLSAoTnVtYmVyKHByZXNldCkgLSAxKSk7CiAgICAgICAgICBmcm9tID0gZC50b0lTT1N0cmluZygpLnNsaWNlKDAsIDEwKTsKICAgICAgICB9CiAgICAgICAgZGF0ZUZyb20udmFsdWUgPSBmcm9tOwogICAgICAgIGRhdGVUby52YWx1ZSA9IHRvOwogICAgICAgIGFwcGx5KCk7CiAgICAgIH0pOwogICAgfSk7CgogICAgU3RhdGUub25DaGFuZ2UoKCkgPT4gewogICAgICBpZiAoYWN0aXZlVGFiID09PSAnZGFzaGJvYXJkJykgRGFzaGJvYXJkLnJlbmRlcigpOwogICAgICBpZiAoYWN0aXZlVGFiID09PSAncmVjb3JkcycpIFJlY29yZHMucmVuZGVyKCk7CiAgICB9KTsKICB9CgogIGZ1bmN0aW9uIHdpcmVUYWJzKCkgewogICAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnLnRhYi1idG4nKS5mb3JFYWNoKChidG4pID0+IHsKICAgICAgYnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gc3dpdGNoVGFiKGJ0bi5kYXRhc2V0LnRhYikpOwogICAgfSk7CiAgfQoKICB3aW5kb3cuYWRkRXZlbnRMaXN0ZW5lcignbHJzOmRhdGEtdXBkYXRlZCcsIGFzeW5jICgpID0+IHsKICAgIGF3YWl0IGxvYWRGaWx0ZXJPcHRpb25zKCk7CiAgICByZW5kZXJBY3RpdmVWaWV3KCk7CiAgfSk7CgogIC8vIC0tLS0tLS0tLS0gQXV0aCBzY3JlZW4gLS0tLS0tLS0tLQogIGxldCBhcHBJbml0aWFsaXplZCA9IGZhbHNlOwoKICBmdW5jdGlvbiBzaG93QXV0aFNjcmVlbigpIHsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhdXRoU2NyZWVuJykuc3R5bGUuZGlzcGxheSA9ICdmbGV4JzsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhcHBTaGVsbCcpLnN0eWxlLmRpc3BsYXkgPSAnbm9uZSc7CiAgICBjb25zdCBjb2RlSW5wdXQgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXV0aENvZGUnKTsKICAgIGNvZGVJbnB1dC52YWx1ZSA9ICcnOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2F1dGhFcnJvcicpLnRleHRDb250ZW50ID0gJyc7CiAgICBjb2RlSW5wdXQuZm9jdXMoKTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIHNob3dBcHAoKSB7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXV0aFNjcmVlbicpLnN0eWxlLmRpc3BsYXkgPSAnbm9uZSc7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXBwU2hlbGwnKS5zdHlsZS5kaXNwbGF5ID0gJyc7CiAgICBpZiAoIWFwcEluaXRpYWxpemVkKSB7CiAgICAgIGFwcEluaXRpYWxpemVkID0gdHJ1ZTsKICAgICAgd2lyZVRhYnMoKTsKICAgICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2xvZ291dEJ0bicpLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgYXN5bmMgKCkgPT4gewogICAgICAgIGF3YWl0IEFwaS5hdXRoTG9nb3V0KCk7CiAgICAgICAgYXBwSW5pdGlhbGl6ZWQgPSBmYWxzZTsKICAgICAgICBzaG93QXV0aFNjcmVlbigpOwogICAgICB9KTsKICAgICAgYXdhaXQgbG9hZEZpbHRlck9wdGlvbnMoKTsKICAgICAgd2lyZUZpbHRlckJhcigpOwogICAgICBzd2l0Y2hUYWIoJ2Rhc2hib2FyZCcpOwogICAgfSBlbHNlIHsKICAgICAgYXdhaXQgbG9hZEZpbHRlck9wdGlvbnMoKTsKICAgICAgcmVuZGVyQWN0aXZlVmlldygpOwogICAgfQogIH0KCiAgYXN5bmMgZnVuY3Rpb24gc3VibWl0QXV0aCgpIHsKICAgIGNvbnN0IGVycm9yRWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXV0aEVycm9yJyk7CiAgICBjb25zdCBidG4gPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXV0aFN1Ym1pdEJ0bicpOwogICAgY29uc3QgY29kZUlucHV0ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2F1dGhDb2RlJyk7CiAgICBlcnJvckVsLnRleHRDb250ZW50ID0gJyc7CiAgICBidG4uZGlzYWJsZWQgPSB0cnVlOwogICAgYnRuLnRleHRDb250ZW50ID0gJ0NoZWNraW5n4oCmJzsKICAgIHRyeSB7CiAgICAgIGF3YWl0IEFwaS5hdXRoTG9naW4oY29kZUlucHV0LnZhbHVlKTsKICAgICAgYXdhaXQgc2hvd0FwcCgpOwogICAgfSBjYXRjaCAoZXJyKSB7CiAgICAgIGVycm9yRWwudGV4dENvbnRlbnQgPSBlcnIubWVzc2FnZTsKICAgIH0gZmluYWxseSB7CiAgICAgIGJ0bi5kaXNhYmxlZCA9IGZhbHNlOwogICAgICBidG4uaW5uZXJIVE1MID0gJzxpIGRhdGEtbHVjaWRlPSJhcnJvdy1yaWdodCIgc3R5bGU9IndpZHRoOjE0cHg7aGVpZ2h0OjE0cHg7Ij48L2k+IEVudGVyJzsKICAgIH0KICB9CgogIGZ1bmN0aW9uIHdpcmVBdXRoRm9ybSgpIHsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhdXRoU3VibWl0QnRuJykuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBzdWJtaXRBdXRoKTsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhdXRoQ29kZScpLmFkZEV2ZW50TGlzdGVuZXIoJ2tleWRvd24nLCAoZSkgPT4geyBpZiAoZS5rZXkgPT09ICdFbnRlcicpIHN1Ym1pdEF1dGgoKTsgfSk7CiAgfQoKICB3aW5kb3cuYWRkRXZlbnRMaXN0ZW5lcignbHJzOnNpZ25lZC1vdXQnLCAoKSA9PiB7CiAgICBhcHBJbml0aWFsaXplZCA9IGZhbHNlOwogICAgc2hvd0F1dGhTY3JlZW4oKTsKICB9KTsKCiAgYXN5bmMgZnVuY3Rpb24gaW5pdCgpIHsKICAgIGFwcGx5QnJhbmRpbmcoKTsKICAgIHdpcmVBdXRoRm9ybSgpOwogICAgY29uc3QgeyBhdXRoZW50aWNhdGVkIH0gPSBhd2FpdCBBcGkuYXV0aE1lKCk7CiAgICBpZiAoYXV0aGVudGljYXRlZCkgYXdhaXQgc2hvd0FwcCgpOwogICAgZWxzZSBzaG93QXV0aFNjcmVlbigpOwogIH0KCiAgLy8gSWNvbnMgYXJlIHBsYWNlZCBhcyA8aSBkYXRhLWx1Y2lkZT0iLi4uIj4gcGxhY2Vob2xkZXJzIHRocm91Z2hvdXQgdGhlIGR5bmFtaWNhbGx5CiAgLy8gcmVuZGVyZWQgVUk7IEx1Y2lkZSByZXBsYWNlcyBlYWNoIHdpdGggYW4gaW5saW5lIFNWRy4gUmF0aGVyIHRoYW4gcmVtZW1iZXJpbmcgdG8gY2FsbAogIC8vIHRoaXMgYWZ0ZXIgZXZlcnkgc2luZ2xlIHJlbmRlciwgb25lIG9ic2VydmVyIGNhdGNoZXMgZXZlcnkgRE9NIGNoYW5nZSB0aGF0IGNvdWxkIGhhdmUKICAvLyBpbnRyb2R1Y2VkIGEgbmV3IHBsYWNlaG9sZGVyLgogIGlmICh3aW5kb3cubHVjaWRlKSB7CiAgICB3aW5kb3cubHVjaWRlLmNyZWF0ZUljb25zKCk7CiAgICAvLyBjcmVhdGVJY29ucygpIHJlcGxhY2VzIDxpIGRhdGEtbHVjaWRlPiBwbGFjZWhvbGRlcnMgd2l0aCA8c3ZnPiDigJQgaXRzZWxmIGEgRE9NCiAgICAvLyBtdXRhdGlvbi4gV2l0aG91dCBkaXNjb25uZWN0aW5nIGZpcnN0LCB0aGF0IHdyaXRlIHJlLXRyaWdnZXJzIHRoaXMgc2FtZSBvYnNlcnZlcgogICAgLy8gZm9yZXZlciAoYW4gaW5maW5pdGUgbXV0YXRlL29ic2VydmUgbG9vcCB0aGF0IHBlZ3MgdGhlIENQVSBhbmQgY3Jhc2hlcyB0aGUgdGFiKS4KICAgIC8vIERpc2Nvbm5lY3RpbmcgYmVmb3JlIGVhY2ggcGFzcyBhbmQgcmVjb25uZWN0aW5nIGFmdGVyLCBwbHVzIGJhdGNoaW5nIGJ1cnN0cyBvZgogICAgLy8gbXV0YXRpb25zIGludG8gYSBzaW5nbGUgbWljcm90YXNrLCBicmVha3MgdGhlIGN5Y2xlLgogICAgbGV0IGljb25zU2NoZWR1bGVkID0gZmFsc2U7CiAgICBjb25zdCBpY29uT2JzZXJ2ZXIgPSBuZXcgTXV0YXRpb25PYnNlcnZlcigoKSA9PiB7CiAgICAgIGlmIChpY29uc1NjaGVkdWxlZCkgcmV0dXJuOwogICAgICBpY29uc1NjaGVkdWxlZCA9IHRydWU7CiAgICAgIHF1ZXVlTWljcm90YXNrKCgpID0+IHsKICAgICAgICBpY29uc1NjaGVkdWxlZCA9IGZhbHNlOwogICAgICAgIGljb25PYnNlcnZlci5kaXNjb25uZWN0KCk7CiAgICAgICAgd2luZG93Lmx1Y2lkZS5jcmVhdGVJY29ucygpOwogICAgICAgIGljb25PYnNlcnZlci5vYnNlcnZlKGRvY3VtZW50LmJvZHksIHsgY2hpbGRMaXN0OiB0cnVlLCBzdWJ0cmVlOiB0cnVlIH0pOwogICAgICB9KTsKICAgIH0pOwogICAgaWNvbk9ic2VydmVyLm9ic2VydmUoZG9jdW1lbnQuYm9keSwgeyBjaGlsZExpc3Q6IHRydWUsIHN1YnRyZWU6IHRydWUgfSk7CiAgfQoKICBpbml0KCk7Cn0pKCk7Cjwvc2NyaXB0Pgo8L2JvZHk+CjwvaHRtbD4K';
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
