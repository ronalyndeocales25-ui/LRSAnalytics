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
// Per-metric "does any row in scope actually carry this metric" flag. Lets callers tell a
// genuine zero (metric reported, value 0) apart from N/A (metric never provided for these
// accounts/platforms) — SUM() alone collapses both to NULL. 1 = applicable, 0 = not.
const PRESENCE_METRICS_SQL = CANONICAL_METRIC_KEYS.map((k) => `MAX(pm.${k} IS NOT NULL) AS has_${k}`).join(', ');
const METRIC_LABELS = Object.fromEntries(CANONICAL_METRICS.map((m) => [m.key, m.label]));

/** Pulls the has_<key> presence flags off a query row into { key: boolean }. */
function applicableMap(row) {
  const out = {};
  for (const key of CANONICAL_METRIC_KEYS) out[key] = row[`has_${key}`] === 1;
  return out;
}

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
    SELECT ${SUM_METRICS_SQL}, ${PRESENCE_METRICS_SQL}, COUNT(DISTINCT p.id) AS post_count
    FROM post_metrics pm JOIN posts p ON p.id = pm.post_id
    ${where}`;
  const row = db.prepare(sql).get(...params);
  // `applicable[key]` distinguishes a real zero from N/A; the numeric fields stay
  // coerced to 0 for backward compatibility, so callers that don't care are unaffected.
  const result = { post_count: row.post_count || 0, applicable: applicableMap(row) };
  for (const key of CANONICAL_METRIC_KEYS) result[key] = row[key] || 0;
  return result;
}

function platformBreakdown(filters) {
  const { where, params } = buildFilter(filters);
  const sql = `
    SELECT pm.platform, ${SUM_METRICS_SQL}, ${PRESENCE_METRICS_SQL}, COUNT(DISTINCT p.id) AS post_count
    FROM post_metrics pm JOIN posts p ON p.id = pm.post_id
    ${where}
    GROUP BY pm.platform`;
  const rows = db.prepare(sql).all(...params);
  const byId = Object.fromEntries(rows.map((r) => [r.platform, r]));
  return PLATFORM_IDS.filter((id) => byId[id]).map((id) => {
    const r = byId[id];
    const meta = PLATFORMS.find((p) => p.id === id);
    const out = { platform: id, label: meta.label, color: meta.color, post_count: r.post_count, applicable: applicableMap(r) };
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
    // applicable=false means no post in this slice carries the metric at all — the
    // Total/Average/Highest below are then N/A, not a genuine zero.
    applicable: agg.highest !== null,
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
    const applicable = {};
    for (const key of CANONICAL_METRIC_KEYS) {
      totals[key] = row[key] || 0;
      growth[key] = explicitRange ? pctChange(row[key] || 0, prev[key] || 0) : null;
      applicable[key] = Boolean(row.applicable && row.applicable[key]);
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
      applicable,
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
const INDEX_HTML_BASE64 = 'PCFkb2N0eXBlIGh0bWw+CjxodG1sIGxhbmc9ImVuIj4KPGhlYWQ+CjxtZXRhIGNoYXJzZXQ9IlVURi04IiAvPgo8bWV0YSBuYW1lPSJ2aWV3cG9ydCIgY29udGVudD0id2lkdGg9ZGV2aWNlLXdpZHRoLCBpbml0aWFsLXNjYWxlPTEuMCIgLz4KPHRpdGxlPkxSUyBBbmFseXRpY3MgRGFzaGJvYXJkPC90aXRsZT4KPGxpbmsgcmVsPSJpY29uIiB0eXBlPSJpbWFnZS9wbmciIGlkPSJmYXZpY29uTGluayIgLz4KPGxpbmsgcmVsPSJwcmVjb25uZWN0IiBocmVmPSJodHRwczovL2ZvbnRzLmdvb2dsZWFwaXMuY29tIiAvPgo8bGluayByZWw9InByZWNvbm5lY3QiIGhyZWY9Imh0dHBzOi8vZm9udHMuZ3N0YXRpYy5jb20iIGNyb3Nzb3JpZ2luIC8+CjxsaW5rIGhyZWY9Imh0dHBzOi8vZm9udHMuZ29vZ2xlYXBpcy5jb20vY3NzMj9mYW1pbHk9SW50ZXI6d2dodEA0MDA7NTAwOzYwMDs3MDA7ODAwJmRpc3BsYXk9c3dhcCIgcmVsPSJzdHlsZXNoZWV0IiAvPgo8c2NyaXB0IHNyYz0iaHR0cHM6Ly9jZG4uanNkZWxpdnIubmV0L25wbS9jaGFydC5qc0A0LjQuNC9kaXN0L2NoYXJ0LnVtZC5taW4uanMiPjwvc2NyaXB0Pgo8c2NyaXB0IHNyYz0iaHR0cHM6Ly9jZG4uanNkZWxpdnIubmV0L25wbS9jaGFydGpzLXBsdWdpbi1kYXRhbGFiZWxzQDIvZGlzdC9jaGFydGpzLXBsdWdpbi1kYXRhbGFiZWxzLm1pbi5qcyI+PC9zY3JpcHQ+CjxzY3JpcHQgc3JjPSJodHRwczovL2Nkbi5qc2RlbGl2ci5uZXQvbnBtL2x1Y2lkZUAwLjQ2Mi4wL2Rpc3QvdW1kL2x1Y2lkZS5taW4uanMiPjwvc2NyaXB0Pgo8c3R5bGU+Ci8qIC0tLS0tLS0tLS0gRGVzaWduIHRva2VuczogZGFyayBuYXZ5ICsgZ29sZCBicmFuZGVkIHRoZW1lIChzaW5nbGUsIHBlcm1hbmVudCDigJQgbm8gbGlnaHQgdmFyaWFudCkgLS0tLS0tLS0tLSAqLwo6cm9vdCB7CiAgY29sb3Itc2NoZW1lOiBkYXJrOwogIC0tZm9udC1zYW5zOiAnSW50ZXInLCAtYXBwbGUtc3lzdGVtLCBCbGlua01hY1N5c3RlbUZvbnQsICdTRiBQcm8gRGlzcGxheScsICdTZWdvZSBVSScsIFJvYm90bywgc2Fucy1zZXJpZjsKCiAgLS1wYWdlLXBsYW5lOiBsaW5lYXItZ3JhZGllbnQoMTgwZGVnLCAjMGEwZjFjIDAlLCAjMGQxNDI0IDEwMCUpOwogIC0tcGFnZS1wbGFuZS1zb2xpZDogIzBhMGYxYzsKICAtLXNpZGViYXItYmc6ICMwYjEyMjA7CiAgLS1zdXJmYWNlLTE6IHJnYmEoMjMsIDMxLCA1MSwgMC42Mik7IC8qIGdsYXNzOiBjYXJkcywgS1BJIHRpbGVzLCBmaWx0ZXIgYmFyICovCiAgLS1zdXJmYWNlLTI6IHJnYmEoMjU1LCAyNTUsIDI1NSwgMC4wNik7IC8qIGdsYXNzOiBpbnB1dHMsIG5lc3RlZCByb3dzLCBwaWxscyAqLwogIC0tc3VyZmFjZS1zb2xpZDogIzEzMWIyZTsKICAtLWdsYXNzLWJsdXI6IGJsdXIoMjBweCk7CiAgLS1ib3JkZXI6IHJnYmEoMjU1LCAyNTUsIDI1NSwgMC4wOSk7CiAgLS10ZXh0LXByaW1hcnk6ICNmNGY1Zjc7CiAgLS10ZXh0LXNlY29uZGFyeTogI2I4YmJjNDsKICAtLXRleHQtbXV0ZWQ6ICM4Mjg2OGY7CiAgLS1ncmlkbGluZTogcmdiYSgyNTUsIDI1NSwgMjU1LCAwLjA4KTsKICAtLWJhc2VsaW5lOiByZ2JhKDI1NSwgMjU1LCAyNTUsIDAuMik7CiAgLS1zdWNjZXNzLXRleHQ6ICMzNGM3NmY7CgogIC0tc3RhdHVzLWdvb2Q6ICMyZmI4NjI7CiAgLS1zdGF0dXMtd2FybmluZzogI2YwYTEzYTsKICAtLXN0YXR1cy1zZXJpb3VzOiAjZWM4MzVhOwogIC0tc3RhdHVzLWNyaXRpY2FsOiAjZTA2MDVmOwoKICAtLWFjY2VudC1nb2xkOiAjZjJiMzBlOyAvKiBMUlMgYnJhbmQgZ29sZCDigJQgYWN0aXZlIG5hdiBpdGVtLCBwcmltYXJ5IGludGVyYWN0aXZlIGFjY2VudCAqLwoKICAtLXNlcmllcy0xOiAjMzk4N2U1OyAvKiBmYWNlYm9vayAqLwogIC0tc2VyaWVzLTI6ICMwMDgzMDA7IC8qIGluc3RhZ3JhbSAqLwogIC0tc2VyaWVzLTM6ICNkNTUxODE7IC8qIHRpa3RvayAqLwogIC0tc2VyaWVzLTQ6ICNjOTg1MDA7IC8qIGxpbmtlZGluICovCiAgLS1zZXJpZXMtNTogIzE5OWU3MDsgLyogdGhyZWFkcyAqLwogIC0tc2VyaWVzLTY6ICNkOTU5MjY7IC8qIHlvdXR1YmUgKi8KICAtLXNlcmllcy03OiAjOTA4NWU5OyAvKiByZXNlcnZlZCAqLwogIC0tc2VyaWVzLTg6ICNlNjY3Njc7IC8qIHJlc2VydmVkICovCgogIC0tcmFkaXVzLXNtOiAxMHB4OwogIC0tcmFkaXVzLW1kOiAxNHB4OwogIC0tcmFkaXVzLWxnOiAxOHB4OwoKICAtLXNoYWRvdy1jYXJkOiAwIDFweCAycHggcmdiYSgwLDAsMCwwLjIpLCAwIDhweCAyNHB4IC0xMHB4IHJnYmEoMCwwLDAsMC41KTsKICAtLXNoYWRvdy1ob3ZlcjogMCA2cHggMTJweCAtMnB4IHJnYmEoMCwwLDAsMC4zKSwgMCAxOHB4IDQwcHggLTE0cHggcmdiYSgwLDAsMCwwLjYpOwogIC0tc2hhZG93LW1vZGFsOiAwIDI0cHggNjRweCAtMTJweCByZ2JhKDAsMCwwLDAuNyk7CiAgLS1lYXNlOiBjdWJpYy1iZXppZXIoMC40LCAwLCAwLjIsIDEpOwp9CgoqIHsgYm94LXNpemluZzogYm9yZGVyLWJveDsgfQpodG1sLCBib2R5IHsgaGVpZ2h0OiAxMDAlOyB9CmJvZHkgewogIG1hcmdpbjogMDsKICBmb250LWZhbWlseTogdmFyKC0tZm9udC1zYW5zKTsKICBiYWNrZ3JvdW5kOiB2YXIoLS1wYWdlLXBsYW5lKTsKICBiYWNrZ3JvdW5kLWF0dGFjaG1lbnQ6IGZpeGVkOwogIGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOwogIC13ZWJraXQtZm9udC1zbW9vdGhpbmc6IGFudGlhbGlhc2VkOwogIC1tb3otb3N4LWZvbnQtc21vb3RoaW5nOiBncmF5c2NhbGU7Cn0KYnV0dG9uLCBzZWxlY3QsIGlucHV0LCB0ZXh0YXJlYSB7IGZvbnQtZmFtaWx5OiBpbmhlcml0OyB9CmgxLCBoMiwgaDMsIGg0IHsgZm9udC13ZWlnaHQ6IDcwMDsgbGV0dGVyLXNwYWNpbmc6IC0wLjAxZW07IH0KCjo6c2VsZWN0aW9uIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSAzMCUsIHRyYW5zcGFyZW50KTsgfQoKLyogQ3VzdG9tIHNjcm9sbGJhciDigJQgdGhpbiwgdW5vYnRydXNpdmUsIGZpdHMgdGhlIGdsYXNzIGFlc3RoZXRpYyAqLwo6Oi13ZWJraXQtc2Nyb2xsYmFyIHsgd2lkdGg6IDEwcHg7IGhlaWdodDogMTBweDsgfQo6Oi13ZWJraXQtc2Nyb2xsYmFyLXRyYWNrIHsgYmFja2dyb3VuZDogdHJhbnNwYXJlbnQ7IH0KOjotd2Via2l0LXNjcm9sbGJhci10aHVtYiB7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS10ZXh0LW11dGVkKSA0MCUsIHRyYW5zcGFyZW50KTsgYm9yZGVyLXJhZGl1czogMjBweDsgYm9yZGVyOiAycHggc29saWQgdHJhbnNwYXJlbnQ7IGJhY2tncm91bmQtY2xpcDogcGFkZGluZy1ib3g7IH0KOjotd2Via2l0LXNjcm9sbGJhci10aHVtYjpob3ZlciB7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS10ZXh0LW11dGVkKSA2MCUsIHRyYW5zcGFyZW50KTsgYmFja2dyb3VuZC1jbGlwOiBwYWRkaW5nLWJveDsgfQoKLmFwcC1zaGVsbCB7IGhlaWdodDogMTAwdmg7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiByb3c7IG92ZXJmbG93OiBoaWRkZW47IH0KCi8qIC0tLS0tLS0tLS0gU2lkZWJhciAtLS0tLS0tLS0tICovCi5zaWRlYmFyIHsKICBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBnYXA6IDRweDsKICB3aWR0aDogMjQwcHg7IGZsZXg6IDAgMCBhdXRvOwogIHBhZGRpbmc6IDIwcHggMTRweDsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zaWRlYmFyLWJnKTsKICBib3JkZXItcmlnaHQ6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOwogIHotaW5kZXg6IDIwOwogIG92ZXJmbG93LXk6IGF1dG87Cn0KLnNpZGViYXItYnJhbmQgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDhweDsgd2hpdGUtc3BhY2U6IG5vd3JhcDsgcGFkZGluZzogNHB4IDEwcHggMjBweDsgfQouYnJhbmQtbG9nbyB7IGhlaWdodDogMjhweDsgd2lkdGg6IGF1dG87IGRpc3BsYXk6IGJsb2NrOyBmbGV4LXNocmluazogMDsgb2JqZWN0LWZpdDogY29udGFpbjsgfQouYnJhbmQtdGl0bGUgeyBmb250LXdlaWdodDogNjAwOyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgbGV0dGVyLXNwYWNpbmc6IC0wLjAxZW07IH0KCi50YWJzIHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiAycHg7IGZsZXg6IDE7IHBvc2l0aW9uOiByZWxhdGl2ZTsgfQoudGFiLWJ0biB7CiAgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiAxMHB4OwogIGJvcmRlcjogbm9uZTsgYmFja2dyb3VuZDogdHJhbnNwYXJlbnQ7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7CiAgcGFkZGluZzogMTBweCAxMnB4OyBib3JkZXItcmFkaXVzOiAxMHB4OyBjdXJzb3I6IHBvaW50ZXI7IGZvbnQtc2l6ZTogMTRweDsgZm9udC13ZWlnaHQ6IDUwMDsKICB3aGl0ZS1zcGFjZTogbm93cmFwOyBwb3NpdGlvbjogcmVsYXRpdmU7IHRleHQtYWxpZ246IGxlZnQ7IHdpZHRoOiAxMDAlOwogIHRyYW5zaXRpb246IGNvbG9yIDE4MG1zIHZhcigtLWVhc2UpLCBiYWNrZ3JvdW5kIDE4MG1zIHZhcigtLWVhc2UpOwp9Ci50YWItYnRuIHN2ZyB7IGZsZXgtc2hyaW5rOiAwOyBvcGFjaXR5OiAwLjg7IH0KLnRhYi1idG4uaXMtYWN0aXZlIHN2ZyB7IG9wYWNpdHk6IDE7IGNvbG9yOiB2YXIoLS1zZXJpZXMtMSk7IH0KLnRhYi1idG46aG92ZXIgeyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDglLCB0cmFuc3BhcmVudCk7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyB9Ci50YWItYnRuLmlzLWFjdGl2ZSB7CiAgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IGZvbnQtd2VpZ2h0OiA2MDA7CiAgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSAxNiUsIHRyYW5zcGFyZW50KTsKICBhbmltYXRpb246IHRhYkluZGljYXRvckluIDIyMG1zIHZhcigtLWVhc2UpOwp9CkBrZXlmcmFtZXMgdGFiSW5kaWNhdG9ySW4geyBmcm9tIHsgb3BhY2l0eTogMDsgdHJhbnNmb3JtOiB0cmFuc2xhdGVYKC00cHgpOyB9IHRvIHsgb3BhY2l0eTogMTsgdHJhbnNmb3JtOiB0cmFuc2xhdGVYKDApOyB9IH0KLnNpZGViYXItZm9vdGVyIHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiAxMHB4OyBwYWRkaW5nLXRvcDogMTRweDsgbWFyZ2luLXRvcDogMTRweDsgYm9yZGVyLXRvcDogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IH0KCi5zaWRlYmFyLXVzZXIgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDEwcHg7IGZvbnQtc2l6ZTogMTNweDsgfQoKLyogLS0tLS0tLS0tLSBBdXRoIHNjcmVlbiAtLS0tLS0tLS0tICovCi5hdXRoLXNjcmVlbiB7CiAgbWluLWhlaWdodDogMTAwdmg7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogY2VudGVyOwogIGJhY2tncm91bmQ6IHZhcigtLXBhZ2UtcGxhbmUpOyBwYWRkaW5nOiAyMHB4Owp9Ci5hdXRoLWNhcmQgewogIHdpZHRoOiAxMDAlOyBtYXgtd2lkdGg6IDQwMHB4OyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTEpOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOwogIGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtbGcpOyBwYWRkaW5nOiAzMnB4OyBib3gtc2hhZG93OiB2YXIoLS1zaGFkb3ctbW9kYWwpOwogIGFuaW1hdGlvbjogbW9kYWxQYW5lbEluIDI2MG1zIHZhcigtLWVhc2UpOwp9Ci5hdXRoLWJyYW5kIHsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA4cHg7IG1hcmdpbi1ib3R0b206IDIycHg7IH0KLmF1dGgtYnJhbmQgLmJyYW5kLXRpdGxlIHsgZm9udC13ZWlnaHQ6IDcwMDsgZm9udC1zaXplOiAxN3B4OyB9Ci5hdXRoLWJyYW5kIC5icmFuZC1sb2dvIHsgaGVpZ2h0OiAzNnB4OyB9Ci5hdXRoLWZvcm0geyBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBnYXA6IDE0cHg7IG1hcmdpbi10b3A6IDE2cHg7IH0KLmF1dGgtZm9ybSAuZm9ybS1maWVsZCBpbnB1dCB7IHdpZHRoOiAxMDAlOyB9Ci5hdXRoLWVycm9yIHsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IGZvbnQtc2l6ZTogMTJweDsgbWluLWhlaWdodDogMTZweDsgfQoKLyogLS0tLS0tLS0tLSBGaWx0ZXIgYmFyIC0tLS0tLS0tLS0gKi8KLmZpbHRlci1iYXIgewogIGRpc3BsYXk6IGZsZXg7IGZsZXgtd3JhcDogd3JhcDsgYWxpZ24taXRlbXM6IGVuZDsgZ2FwOiAxNnB4OwogIHBhZGRpbmc6IDE0cHggMjBweDsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgYm9yZGVyLWJvdHRvbTogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7CiAgcG9zaXRpb246IHN0aWNreTsgdG9wOiAwOyB6LWluZGV4OiAxOTsKfQouZmlsdGVyLWZpZWxkIHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiA1cHg7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgfQouZmlsdGVyLWZpZWxkIGxhYmVsIHsgZm9udC13ZWlnaHQ6IDYwMDsgfQouZmlsdGVyLXByZXNldHMgeyBmbGV4LWRpcmVjdGlvbjogcm93OyBnYXA6IDZweDsgfQouZmlsdGVyLXByZXNldHMgYnV0dG9uIHsKICBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTIpOyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOwogIGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBib3JkZXItcmFkaXVzOiAyMHB4OyBwYWRkaW5nOiA3cHggMTNweDsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNTAwOyBjdXJzb3I6IHBvaW50ZXI7CiAgdHJhbnNpdGlvbjogY29sb3IgMTgwbXMgdmFyKC0tZWFzZSksIGJhY2tncm91bmQgMTgwbXMgdmFyKC0tZWFzZSksIHRyYW5zZm9ybSAxNTBtcyB2YXIoLS1lYXNlKTsKfQouZmlsdGVyLXByZXNldHMgYnV0dG9uOmhvdmVyIHsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgtMXB4KTsgfQouZmlsdGVyLXByZXNldHMgYnV0dG9uOmFjdGl2ZSB7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgwKSBzY2FsZSgwLjk2KTsgfQouZmlsdGVyLXByZXNldHMgYnV0dG9uLmlzLWFjdGl2ZSB7IGJhY2tncm91bmQ6IHZhcigtLXNlcmllcy0xKTsgY29sb3I6ICNmZmY7IGJvcmRlci1jb2xvcjogdHJhbnNwYXJlbnQ7IGJveC1zaGFkb3c6IDAgNHB4IDE0cHggLTVweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDYwJSwgdHJhbnNwYXJlbnQpOyB9CgovKiAtLS0tLS0tLS0tIE1haW4gY29sdW1uIChzaXRzIGJlc2lkZSB0aGUgc2lkZWJhcjsgc2Nyb2xscyBpbmRlcGVuZGVudGx5IHNvIHRoZSBzaWRlYmFyIHN0YXlzIGZ1bGx5IHZpc2libGUpIC0tLS0tLS0tLS0gKi8KLm1haW4tY29sIHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZmxleDogMSAxIGF1dG87IG1pbi13aWR0aDogMDsgaGVpZ2h0OiAxMDAlOyBvdmVyZmxvdy15OiBhdXRvOyB9CgovKiAtLS0tLS0tLS0tIFZpZXcgYXJlYSAtLS0tLS0tLS0tICovCi52aWV3LWFyZWEgeyBmbGV4OiAxOyBwYWRkaW5nOiAyNHB4OyBtYXgtd2lkdGg6IDE4MDBweDsgd2lkdGg6IDEwMCU7IG1hcmdpbjogMCBhdXRvOyB9Ci52aWV3IHsgZGlzcGxheTogbm9uZTsgfQoudmlldy5pcy1hY3RpdmUgeyBkaXNwbGF5OiBibG9jazsgYW5pbWF0aW9uOiB2aWV3RmFkZUluIDI2MG1zIHZhcigtLWVhc2UpOyB9CkBrZXlmcmFtZXMgdmlld0ZhZGVJbiB7CiAgZnJvbSB7IG9wYWNpdHk6IDA7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSg2cHgpOyB9CiAgdG8geyBvcGFjaXR5OiAxOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMCk7IH0KfQoKLnNlY3Rpb24tdGl0bGUgeyBmb250LXNpemU6IDE2cHg7IGZvbnQtd2VpZ2h0OiA3MDA7IG1hcmdpbjogMzJweCAwIDE0cHg7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyBsZXR0ZXItc3BhY2luZzogLTAuMDFlbTsgfQouc2VjdGlvbi10aXRsZTpmaXJzdC1jaGlsZCB7IG1hcmdpbi10b3A6IDA7IH0KCi8qIC0tLS0tLS0tLS0gSW5wdXRzIOKAlCBvbmUgc2hhcmVkIGdsYXNzIHRyZWF0bWVudCBmb3IgZXZlcnkgdGV4dCBpbnB1dCwgc2VsZWN0LCBhbmQgZGF0ZSBwaWNrZXIgLS0tLS0tLS0tLSAqLwouZmlsdGVyLWZpZWxkIHNlbGVjdCwgLmZpbHRlci1maWVsZCBpbnB1dFt0eXBlPSJkYXRlIl0sCi5mb3JtLWZpZWxkIGlucHV0LCAuZm9ybS1maWVsZCBzZWxlY3QsIC5mb3JtLWZpZWxkIHRleHRhcmVhLAouZGFzaGJvYXJkLWNvbnRyb2xzIHNlbGVjdCwgLnJlY29yZHMtc2VhcmNoIGlucHV0LAouZmllbGQtaW5saW5lIHNlbGVjdCwgLmZpZWxkLWlubGluZSBpbnB1dCwKLmNvbmZsaWN0LXJvdyBzZWxlY3QsIC5jYXJkLWhlYWRlciBzZWxlY3QgewogIGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1zbSk7CiAgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0yKTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IGZvbnQtc2l6ZTogMTNweDsKICBwYWRkaW5nOiA4cHggMTJweDsgbWluLXdpZHRoOiAxNDBweDsKICB0cmFuc2l0aW9uOiBib3JkZXItY29sb3IgMTYwbXMgdmFyKC0tZWFzZSksIGJveC1zaGFkb3cgMTYwbXMgdmFyKC0tZWFzZSk7Cn0KLmZpbHRlci1maWVsZCBzZWxlY3Q6aG92ZXIsIC5maWx0ZXItZmllbGQgaW5wdXRbdHlwZT0iZGF0ZSJdOmhvdmVyLAouZm9ybS1maWVsZCBpbnB1dDpob3ZlciwgLmZvcm0tZmllbGQgc2VsZWN0OmhvdmVyLAouZGFzaGJvYXJkLWNvbnRyb2xzIHNlbGVjdDpob3ZlciwgLnJlY29yZHMtc2VhcmNoIGlucHV0OmhvdmVyLAouZmllbGQtaW5saW5lIHNlbGVjdDpob3ZlciwgLmZpZWxkLWlubGluZSBpbnB1dDpob3ZlciwKLmNvbmZsaWN0LXJvdyBzZWxlY3Q6aG92ZXIsIC5jYXJkLWhlYWRlciBzZWxlY3Q6aG92ZXIgewogIGJvcmRlci1jb2xvcjogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSAzNSUsIHZhcigtLWJvcmRlcikpOwp9Ci8qIEEgPHNlbGVjdD4ncyBvd24gYmFja2dyb3VuZCBpcyBhIHRyYW5zbHVjZW50IGdsYXNzIHRpbnQgbWVhbnQgdG8gYmxlbmQgd2l0aAogICB0aGUgcGFnZSBiZWhpbmQgaXQg4oCUIGJ1dCBpdHMgZHJvcGRvd24gcG9wdXAgcmVuZGVycyBvbiBhbiBpc29sYXRlZCBvcGFxdWUKICAgY2FudmFzLCBzbyB0aGF0IHNhbWUgdHJhbnNsdWNlbnQgdmFsdWUgc2hvd3MgdXAgdGhlcmUgYXMgcGxhaW4gd2hpdGUKICAgaW5zdGVhZCBvZiBkYXJrLiBFdmVyeSA8b3B0aW9uPiwgaW4gZXZlcnkgc2VsZWN0IGluIHRoZSBhcHAsIG5lZWRzIGFuCiAgIGV4cGxpY2l0IHNvbGlkIGRhcmsgYmFja2dyb3VuZC90ZXh0IGNvbG9yIHNvIHRoZSBwb3B1cCBtYXRjaGVzIHRoZSB0aGVtZS4gKi8Kb3B0aW9uIHsgYmFja2dyb3VuZC1jb2xvcjogdmFyKC0tc3VyZmFjZS1zb2xpZCk7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyB9CgouZmlsdGVyLWZpZWxkIHNlbGVjdDpmb2N1cywgLmZpbHRlci1maWVsZCBpbnB1dFt0eXBlPSJkYXRlIl06Zm9jdXMsCi5mb3JtLWZpZWxkIGlucHV0OmZvY3VzLCAuZm9ybS1maWVsZCBzZWxlY3Q6Zm9jdXMsIC5mb3JtLWZpZWxkIHRleHRhcmVhOmZvY3VzLAouZGFzaGJvYXJkLWNvbnRyb2xzIHNlbGVjdDpmb2N1cywgLnJlY29yZHMtc2VhcmNoIGlucHV0OmZvY3VzLAouZmllbGQtaW5saW5lIHNlbGVjdDpmb2N1cywgLmZpZWxkLWlubGluZSBpbnB1dDpmb2N1cywKLmNvbmZsaWN0LXJvdyBzZWxlY3Q6Zm9jdXMsIC5jYXJkLWhlYWRlciBzZWxlY3Q6Zm9jdXMsCi5hdXRoLWZvcm0gaW5wdXQ6Zm9jdXMgewogIG91dGxpbmU6IG5vbmU7IGJvcmRlci1jb2xvcjogdmFyKC0tc2VyaWVzLTEpOwogIGJveC1zaGFkb3c6IDAgMCAwIDNweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDE4JSwgdHJhbnNwYXJlbnQpOwp9CgovKiAtLS0tLS0tLS0tIFN0YXQgdGlsZXMgLS0tLS0tLS0tLSAqLwouc3RhdC1ncmlkIHsKICBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdChhdXRvLWZpdCwgbWlubWF4KDE4MHB4LCAxZnIpKTsgZ2FwOiAxNHB4Owp9Ci5zdGF0LXRpbGUgewogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1tZCk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHBhZGRpbmc6IDE2cHggMThweDsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWNhcmQpOwogIHRyYW5zaXRpb246IHRyYW5zZm9ybSAyMDBtcyB2YXIoLS1lYXNlKSwgYm94LXNoYWRvdyAyMDBtcyB2YXIoLS1lYXNlKTsKICBhbmltYXRpb246IGNhcmRJbiAzMjBtcyB2YXIoLS1lYXNlKSBiYWNrd2FyZHM7Cn0KLnN0YXQtdGlsZTpob3ZlciB7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgtM3B4KTsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWhvdmVyKTsgfQouc3RhdC1sYWJlbCB7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgZm9udC13ZWlnaHQ6IDYwMDsgfQouc3RhdC12YWx1ZSB7IGZvbnQtc2l6ZTogMjdweDsgZm9udC13ZWlnaHQ6IDcwMDsgbWFyZ2luLXRvcDogNXB4OyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgbGV0dGVyLXNwYWNpbmc6IC0wLjAyZW07IH0KLnN0YXQtZGVsdGEgeyBmb250LXNpemU6IDEycHg7IG1hcmdpbi10b3A6IDdweDsgZm9udC13ZWlnaHQ6IDYwMDsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA0cHg7IH0KLnN0YXQtZGVsdGEudXAgeyBjb2xvcjogdmFyKC0tc3VjY2Vzcy10ZXh0KTsgfQouc3RhdC1kZWx0YS5kb3duIHsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IH0KLnN0YXQtZGVsdGEuZmxhdCB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgfQouc3RhdC1kZWx0YS51cDo6YmVmb3JlIHsgY29udGVudDogJ+KGkSc7IH0KLnN0YXQtZGVsdGEuZG93bjo6YmVmb3JlIHsgY29udGVudDogJ+KGkyc7IH0KCi8qIC0tLS0tLS0tLS0gRGFzaGJvYXJkIEtQSSBncmlkIOKAlCBjb21wYWN0LCBzaW5nbGUtcm93LW9uLWRlc2t0b3AgbGF5b3V0LgogICBTY29wZWQgdG8gI2twaUdyaWQgc3BlY2lmaWNhbGx5IChub3QgdGhlIHNoYXJlZCAuc3RhdC1ncmlkLy5zdGF0LXRpbGUKICAgY2xhc3Nlcywgd2hpY2ggQ29tcGFyaXNvbnMgYW5kIHRoZSBVcGxvYWQgcHJldmlldyBzdW1tYXJ5IGFsc28gdXNlKSBzbwogICB0aGlzIGNvbXBhY3RpbmcgZG9lc24ndCBhZmZlY3QgdGhvc2Ugb3RoZXIgc3RhdC10aWxlIGdyaWRzLiAxMCBncmlkIHVuaXRzCiAgIHRvdGFsOiA3IHN0YW5kYXJkIEtQSSB0aWxlcyBhdCAxIHVuaXQgZWFjaCArIEJlc3QgUGVyZm9ybWluZyBQb3N0IGF0IDMKICAgdW5pdHMgKGEgM3gtd2lkZSBsYW5kc2NhcGUgY2FyZCksIGFsbCBzaGFyaW5nIG9uZSBmaXhlZCByb3cgaGVpZ2h0LiAtLS0tLS0tLS0tICovCiNrcGlHcmlkIHsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiByZXBlYXQoMTAsIG1pbm1heCgwLCAxZnIpKTsgZ2FwOiAxMnB4OyB9CiNrcGlHcmlkIC5zdGF0LXRpbGUgewogIGhlaWdodDogMTMycHg7IGJveC1zaXppbmc6IGJvcmRlci1ib3g7IG92ZXJmbG93OiBoaWRkZW47CiAgcGFkZGluZzogMTZweCAxOHB4OwogIGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGp1c3RpZnktY29udGVudDogY2VudGVyOwp9CiNrcGlHcmlkIC5zdGF0LWxhYmVsIHsgZm9udC1zaXplOiAxMnB4OyBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBhbGlnbi1pdGVtczogZmxleC1zdGFydDsgZ2FwOiA4cHg7IH0KI2twaUdyaWQgLnN0YXQtdmFsdWUgeyBmb250LXNpemU6IDMycHg7IG1hcmdpbi10b3A6IDhweDsgbGluZS1oZWlnaHQ6IDEuMTsgfQoja3BpR3JpZCAuc3RhdC1kZWx0YSB7IGZvbnQtc2l6ZTogMTNweDsgbWFyZ2luLXRvcDogOHB4OyB9Cgouc3RhdC1pY29uIHsKICB3aWR0aDogMjhweDsgaGVpZ2h0OiAyOHB4OyBmbGV4OiAwIDAgYXV0bzsgYm9yZGVyLXJhZGl1czogNTAlOwogIGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogY2VudGVyOwogIGNvbG9yOiAjZmZmOwp9Ci5zdGF0LWljb24udjEgeyBiYWNrZ3JvdW5kOiB2YXIoLS1zZXJpZXMtMSk7IH0KLnN0YXQtaWNvbi52MiB7IGJhY2tncm91bmQ6IHZhcigtLXNlcmllcy0yKTsgfQouc3RhdC1pY29uLnYzIHsgYmFja2dyb3VuZDogdmFyKC0tc2VyaWVzLTMpOyB9Ci5zdGF0LWljb24udjQgeyBiYWNrZ3JvdW5kOiB2YXIoLS1zZXJpZXMtNCk7IH0KLnN0YXQtaWNvbi52NSB7IGJhY2tncm91bmQ6IHZhcigtLXNlcmllcy01KTsgfQouc3RhdC1pY29uLnY2IHsgYmFja2dyb3VuZDogdmFyKC0tc2VyaWVzLTYpOyB9Ci5zdGF0LWljb24uZ29sZCB7IGJhY2tncm91bmQ6IHZhcigtLWFjY2VudC1nb2xkKTsgfQpAbWVkaWEgKG1heC13aWR0aDogOTAwcHgpIHsgI2twaUdyaWQgeyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdCg1LCBtaW5tYXgoMCwgMWZyKSk7IH0gfQpAbWVkaWEgKG1heC13aWR0aDogNjQwcHgpIHsgI2twaUdyaWQgeyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdCgyLCBtaW5tYXgoMCwgMWZyKSk7IH0gfQoKLmluc2lnaHRzLWxpc3QgeyBsaXN0LXN0eWxlOiBub25lOyBtYXJnaW46IDA7IHBhZGRpbmc6IDA7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGdhcDogMTBweDsgfQouaW5zaWdodHMtbGlzdCBsaSB7CiAgZm9udC1zaXplOiAxM3B4OyBsaW5lLWhlaWdodDogMS41OyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyBwYWRkaW5nLWxlZnQ6IDE4cHg7IHBvc2l0aW9uOiByZWxhdGl2ZTsKfQouaW5zaWdodHMtbGlzdCBsaTo6YmVmb3JlIHsKICBjb250ZW50OiAn4pymJzsgcG9zaXRpb246IGFic29sdXRlOyBsZWZ0OiAwOyBjb2xvcjogdmFyKC0tc2VyaWVzLTEpOyBmb250LXNpemU6IDExcHg7IHRvcDogMnB4Owp9CgpAa2V5ZnJhbWVzIGNhcmRJbiB7CiAgZnJvbSB7IG9wYWNpdHk6IDA7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgxMHB4KTsgfQogIHRvIHsgb3BhY2l0eTogMTsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDApOyB9Cn0KCi8qIC0tLS0tLS0tLS0gQ2FyZHMgLyBjaGFydHMgLS0tLS0tLS0tLSAqLwouY2FyZC1ncmlkIHsgZGlzcGxheTogZ3JpZDsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiAyZnIgMWZyOyBnYXA6IDE2cHg7IGFsaWduLWl0ZW1zOiBzdGFydDsgfQouY2FyZC1ncmlkLmV2ZW4geyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IDFmciAxZnI7IH0KQG1lZGlhIChtYXgtd2lkdGg6IDkwMHB4KSB7IC5jYXJkLWdyaWQsIC5jYXJkLWdyaWQuZXZlbiB7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogMWZyOyB9IH0KLmNhcmQgewogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1sZyk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHBhZGRpbmc6IDE4cHg7IGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdy1jYXJkKTsKICB0cmFuc2l0aW9uOiBib3gtc2hhZG93IDIyMG1zIHZhcigtLWVhc2UpLCB0cmFuc2Zvcm0gMjIwbXMgdmFyKC0tZWFzZSk7CiAgYW5pbWF0aW9uOiBjYXJkSW4gMzIwbXMgdmFyKC0tZWFzZSkgYmFja3dhcmRzOwp9Ci5jYXJkOmhvdmVyIHsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWhvdmVyKTsgfQouY2FyZC1oZWFkZXIgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6IHNwYWNlLWJldHdlZW47IGdhcDogOHB4OyBtYXJnaW4tYm90dG9tOiAxNHB4OyB9Ci5jYXJkLWhlYWRlciBoMyB7IGZvbnQtc2l6ZTogMTRweDsgbWFyZ2luOiAwOyBmb250LXdlaWdodDogNzAwOyBsZXR0ZXItc3BhY2luZzogLTAuMDA1ZW07IH0KLmNhcmQtaGVhZGVyIHNlbGVjdCB7IGZvbnQtc2l6ZTogMTJweDsgcGFkZGluZzogNnB4IDEwcHg7IG1pbi13aWR0aDogMDsgfQouY2hhcnQtd3JhcCB7IHBvc2l0aW9uOiByZWxhdGl2ZTsgaGVpZ2h0OiAyODBweDsgfQouY2hhcnQtd3JhcC50YWxsIHsgaGVpZ2h0OiAzNDBweDsgfQoKLmxlZ2VuZC1yb3cgeyBkaXNwbGF5OiBmbGV4OyBmbGV4LXdyYXA6IHdyYXA7IGdhcDogMTJweDsgbWFyZ2luLXRvcDogMTBweDsgZm9udC1zaXplOiAxMnB4OyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyB9Ci5sZWdlbmQtaXRlbSB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogNnB4OyB9Ci5sZWdlbmQtc3dhdGNoIHsgd2lkdGg6IDEwcHg7IGhlaWdodDogMTBweDsgYm9yZGVyLXJhZGl1czogM3B4OyBkaXNwbGF5OiBpbmxpbmUtYmxvY2s7IH0KLmxlZ2VuZC1saW5lIHsgd2lkdGg6IDE0cHg7IGhlaWdodDogMnB4OyBib3JkZXItcmFkaXVzOiAycHg7IGRpc3BsYXk6IGlubGluZS1ibG9jazsgfQoKLyogLS0tLS0tLS0tLSBUYWJsZXMg4oCUIHByZW1pdW0gZGF0YWJhc2UgZmVlbCwgbm90IGEgc3ByZWFkc2hlZXQgLS0tLS0tLS0tLSAqLwoudGFibGUtc2Nyb2xsIHsKICBvdmVyZmxvdy14OiBhdXRvOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtbWQpOwogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMik7Cn0KLmRhdGEtdGFibGUgeyB3aWR0aDogMTAwJTsgYm9yZGVyLWNvbGxhcHNlOiBzZXBhcmF0ZTsgYm9yZGVyLXNwYWNpbmc6IDA7IGZvbnQtc2l6ZTogMTNweDsgfQouZGF0YS10YWJsZSB0aCwgLmRhdGEtdGFibGUgdGQgeyB0ZXh0LWFsaWduOiBsZWZ0OyBwYWRkaW5nOiAxMXB4IDE0cHg7IGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1ncmlkbGluZSk7IHdoaXRlLXNwYWNlOiBub3dyYXA7IH0KLmRhdGEtdGFibGUgdGQud3JhcCB7IHdoaXRlLXNwYWNlOiBub3JtYWw7IH0KLmRhdGEtdGFibGUgdGhlYWQgdGggewogIGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IGZvbnQtd2VpZ2h0OiA2MDA7IGZvbnQtc2l6ZTogMTFweDsgdGV4dC10cmFuc2Zvcm06IHVwcGVyY2FzZTsgbGV0dGVyLXNwYWNpbmc6IDAuMDRlbTsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTEpOyBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgcG9zaXRpb246IHN0aWNreTsgdG9wOiAwOyB6LWluZGV4OiAxOwp9Ci5kYXRhLXRhYmxlIHRoZWFkIHRoLnNvcnRhYmxlLXRoIHsgY3Vyc29yOiBwb2ludGVyOyB1c2VyLXNlbGVjdDogbm9uZTsgdHJhbnNpdGlvbjogY29sb3IgMTUwbXMgdmFyKC0tZWFzZSk7IH0KLmRhdGEtdGFibGUgdGhlYWQgdGguc29ydGFibGUtdGg6aG92ZXIgeyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgfQouZGF0YS10YWJsZSB0aGVhZCB0aCAuc29ydC1hcnJvdyB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgZm9udC1zaXplOiAxMHB4OyBtYXJnaW4tbGVmdDogMnB4OyB9Ci5kYXRhLXRhYmxlIHRoZWFkIHRoLnNvcnRhYmxlLXRoOmhvdmVyIC5zb3J0LWFycm93IHsgY29sb3I6IHZhcigtLXNlcmllcy0xKTsgfQouZGF0YS10YWJsZSB0Ym9keSB0cjpudGgtY2hpbGQoZXZlbikgeyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tdGV4dC1tdXRlZCkgNCUsIHRyYW5zcGFyZW50KTsgfQouZGF0YS10YWJsZSB0ZC5udW0geyBmb250LXZhcmlhbnQtbnVtZXJpYzogdGFidWxhci1udW1zOyB0ZXh0LWFsaWduOiByaWdodDsgfQouZGF0YS10YWJsZSB0aC5udW0geyB0ZXh0LWFsaWduOiByaWdodDsgfQouZGF0YS10YWJsZSB0Ym9keSB0ciB7IHRyYW5zaXRpb246IGJhY2tncm91bmQgMTUwbXMgdmFyKC0tZWFzZSk7IH0KLmRhdGEtdGFibGUgdGJvZHkgdHI6aG92ZXIgeyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDclLCB0cmFuc3BhcmVudCk7IH0KLmRhdGEtdGFibGUgdGJvZHkgdHI6bGFzdC1jaGlsZCB0ZCB7IGJvcmRlci1ib3R0b206IG5vbmU7IH0KLnBsYXRmb3JtLXBpbGwgewogIGRpc3BsYXk6IGlubGluZS1mbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDZweDsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNjAwOwogIHBhZGRpbmc6IDRweCAxMHB4OyBib3JkZXItcmFkaXVzOiAyMHB4OyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTEpOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOwp9Ci5wbGF0Zm9ybS1kb3QgeyB3aWR0aDogOHB4OyBoZWlnaHQ6IDhweDsgYm9yZGVyLXJhZGl1czogNTAlOyB9CgovKiAtLS0tLS0tLS0tIEJ1dHRvbnMg4oCUIG5ldmVyIGZsYXQ6IHNvZnQgc2hhZG93LCBob3ZlciBsaWZ0LCBwcmVzcyBzY2FsZSAtLS0tLS0tLS0tICovCi5idG4gewogIGRpc3BsYXk6IGlubGluZS1mbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6IGNlbnRlcjsgZ2FwOiA2cHg7CiAgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0yKTsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHBhZGRpbmc6IDlweCAxN3B4OyBib3JkZXItcmFkaXVzOiAxMXB4OyBjdXJzb3I6IHBvaW50ZXI7IGZvbnQtc2l6ZTogMTNweDsgZm9udC13ZWlnaHQ6IDYwMDsKICBib3gtc2hhZG93OiAwIDFweCAycHggcmdiYSgxNSwxNywyMSwwLjA0KTsKICB0cmFuc2l0aW9uOiB0cmFuc2Zvcm0gMTUwbXMgdmFyKC0tZWFzZSksIGJveC1zaGFkb3cgMTUwbXMgdmFyKC0tZWFzZSksIGZpbHRlciAxNTBtcyB2YXIoLS1lYXNlKSwgYmFja2dyb3VuZCAxNTBtcyB2YXIoLS1lYXNlKTsKfQouYnRuIHN2ZyB7IGZsZXgtc2hyaW5rOiAwOyB9Ci5idG46aG92ZXIgeyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoLTFweCk7IGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdy1ob3Zlcik7IGZpbHRlcjogYnJpZ2h0bmVzcygxLjAyKTsgfQouYnRuOmFjdGl2ZSB7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgwKSBzY2FsZSgwLjk2KTsgYm94LXNoYWRvdzogMCAxcHggMnB4IHJnYmEoMTUsMTcsMjEsMC4wNik7IH0KLmJ0bi5wcmltYXJ5IHsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zZXJpZXMtMSk7IGNvbG9yOiAjZmZmOyBib3JkZXItY29sb3I6IHRyYW5zcGFyZW50OwogIGJveC1zaGFkb3c6IDAgNHB4IDE0cHggLTVweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDY1JSwgdHJhbnNwYXJlbnQpOwp9Ci5idG4ucHJpbWFyeTpob3ZlciB7IGZpbHRlcjogYnJpZ2h0bmVzcygxLjA3KTsgYm94LXNoYWRvdzogMCA4cHggMjJweCAtNnB4IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zZXJpZXMtMSkgNzAlLCB0cmFuc3BhcmVudCk7IH0KLmJ0bi5kYW5nZXIgewogIGJhY2tncm91bmQ6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IGNvbG9yOiAjZmZmOyBib3JkZXItY29sb3I6IHRyYW5zcGFyZW50OwogIGJveC1zaGFkb3c6IDAgNHB4IDE0cHggLTVweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc3RhdHVzLWNyaXRpY2FsKSA1NSUsIHRyYW5zcGFyZW50KTsKfQouYnRuLmRhbmdlcjpob3ZlciB7IGZpbHRlcjogYnJpZ2h0bmVzcygxLjA2KTsgYm94LXNoYWRvdzogMCA4cHggMjJweCAtNnB4IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zdGF0dXMtY3JpdGljYWwpIDYwJSwgdHJhbnNwYXJlbnQpOyB9Ci5idG4uc3VjY2VzcyB7CiAgYmFja2dyb3VuZDogdmFyKC0tc3RhdHVzLWdvb2QpOyBjb2xvcjogI2ZmZjsgYm9yZGVyLWNvbG9yOiB0cmFuc3BhcmVudDsKICBib3gtc2hhZG93OiAwIDRweCAxNHB4IC01cHggY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXN0YXR1cy1nb29kKSA1NSUsIHRyYW5zcGFyZW50KTsKfQouYnRuLnN1Y2Nlc3M6aG92ZXIgeyBmaWx0ZXI6IGJyaWdodG5lc3MoMS4wNik7IGJveC1zaGFkb3c6IDAgOHB4IDIycHggLTZweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc3RhdHVzLWdvb2QpIDYwJSwgdHJhbnNwYXJlbnQpOyB9Ci5idG46ZGlzYWJsZWQgeyBvcGFjaXR5OiAwLjQ1OyBjdXJzb3I6IG5vdC1hbGxvd2VkOyB0cmFuc2Zvcm06IG5vbmU7IGJveC1zaGFkb3c6IG5vbmU7IGZpbHRlcjogbm9uZTsgfQouYnRuLXJvdyB7IGRpc3BsYXk6IGZsZXg7IGdhcDogOHB4OyBmbGV4LXdyYXA6IHdyYXA7IH0KCi8qIC0tLS0tLS0tLS0gVXBsb2FkIC0tLS0tLS0tLS0gKi8KLmRyb3B6b25lIHsKICBib3JkZXI6IDJweCBkYXNoZWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLWxnKTsgcGFkZGluZzogNDBweCAyMHB4OwogIHRleHQtYWxpZ246IGNlbnRlcjsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIGN1cnNvcjogcG9pbnRlcjsgdHJhbnNpdGlvbjogYm9yZGVyLWNvbG9yIDIwMG1zIHZhcigtLWVhc2UpLCBiYWNrZ3JvdW5kIDIwMG1zIHZhcigtLWVhc2UpLCB0cmFuc2Zvcm0gMjAwbXMgdmFyKC0tZWFzZSk7Cn0KLmRyb3B6b25lOmhvdmVyIHsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKC0xcHgpOyB9Ci5kcm9wem9uZS5pcy1kcmFnIHsgYm9yZGVyLWNvbG9yOiB2YXIoLS1zZXJpZXMtMSk7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zZXJpZXMtMSkgNiUsIHZhcigtLXN1cmZhY2UtMikpOyB0cmFuc2Zvcm06IHNjYWxlKDEuMDA1KTsgfQouZHJvcHpvbmUgaDMgeyBtYXJnaW46IDAgMCA2cHg7IGZvbnQtc2l6ZTogMTVweDsgfQouZHJvcHpvbmUgcCB7IG1hcmdpbjogMDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgZm9udC1zaXplOiAxM3B4OyB9Ci5kcm9wem9uZSBpbnB1dFt0eXBlPSJmaWxlIl0geyBkaXNwbGF5OiBub25lOyB9CgouY29uZmxpY3QtbGlzdCB7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGdhcDogOHB4OyBtYXJnaW46IDEycHggMDsgfQouY29uZmxpY3Qtcm93IHsKICBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6IHNwYWNlLWJldHdlZW47IGdhcDogMTJweDsKICBwYWRkaW5nOiAxMXB4IDE0cHg7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1zbSk7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMik7CiAgdHJhbnNpdGlvbjogYm94LXNoYWRvdyAxODBtcyB2YXIoLS1lYXNlKTsKfQouY29uZmxpY3Qtcm93OmhvdmVyIHsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWNhcmQpOyB9Ci5jb25mbGljdC1yb3cgLndlZWstbGFiZWwgeyBmb250LXdlaWdodDogNjAwOyBmb250LXNpemU6IDEzcHg7IH0KLmNvbmZsaWN0LXJvdyAud2Vlay1tZXRhIHsgZm9udC1zaXplOiAxMnB4OyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyB9Ci5jb25mbGljdC1yb3cgc2VsZWN0IHsgbWluLXdpZHRoOiAwOyB9CgouYmFkZ2UgeyBkaXNwbGF5OiBpbmxpbmUtYmxvY2s7IHBhZGRpbmc6IDNweCAxMHB4OyBib3JkZXItcmFkaXVzOiAyMHB4OyBmb250LXNpemU6IDExcHg7IGZvbnQtd2VpZ2h0OiA3MDA7IH0KLmJhZGdlLnN1Y2Nlc3MgeyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc3RhdHVzLWdvb2QpIDE4JSwgdHJhbnNwYXJlbnQpOyBjb2xvcjogdmFyKC0tc3RhdHVzLWdvb2QpOyB9Ci5iYWRnZS5wYXJ0aWFsIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXN0YXR1cy13YXJuaW5nKSAyNSUsIHRyYW5zcGFyZW50KTsgY29sb3I6ICM4YTYzMDA7IH0KLmJhZGdlLmZhaWxlZCB7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zdGF0dXMtY3JpdGljYWwpIDE4JSwgdHJhbnNwYXJlbnQpOyBjb2xvcjogdmFyKC0tc3RhdHVzLWNyaXRpY2FsKTsgfQouYmFkZ2UuZXJyb3Itc2V2IHsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IH0KLmJhZGdlLndhcm5pbmctc2V2IHsgY29sb3I6ICM4YTYzMDA7IH0KLmJhZGdlLnNraXAtc2V2IHsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyB9CgouaXNzdWVzLWxpc3QgeyBtYXgtaGVpZ2h0OiAyMjBweDsgb3ZlcmZsb3cteTogYXV0bzsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLXNtKTsgfQouaXNzdWUtcm93IHsgcGFkZGluZzogOXB4IDE0cHg7IGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1ncmlkbGluZSk7IGZvbnQtc2l6ZTogMTJweDsgfQouaXNzdWUtcm93Omxhc3QtY2hpbGQgeyBib3JkZXItYm90dG9tOiBub25lOyB9Ci5pc3N1ZS1yb3cgLnJvdy1ubyB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgbWFyZ2luLXJpZ2h0OiA2cHg7IH0KCi8qIC0tLS0tLS0tLS0gVG9hc3QgLS0tLS0tLS0tLSAqLwoudG9hc3Qtcm9vdCB7IHBvc2l0aW9uOiBmaXhlZDsgYm90dG9tOiAyMHB4OyByaWdodDogMjBweDsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiA4cHg7IHotaW5kZXg6IDEwMDsgfQoudG9hc3QgewogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1zbSk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHBhZGRpbmc6IDEycHggMTZweDsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LW1vZGFsKTsgZm9udC1zaXplOiAxM3B4OyBtYXgtd2lkdGg6IDM0MHB4OwogIGFuaW1hdGlvbjogdG9hc3QtaW4gMjIwbXMgdmFyKC0tZWFzZSk7Cn0KLnRvYXN0LnN1Y2Nlc3MgeyBib3JkZXItbGVmdDogM3B4IHNvbGlkIHZhcigtLXN0YXR1cy1nb29kKTsgfQoudG9hc3QuZXJyb3IgeyBib3JkZXItbGVmdDogM3B4IHNvbGlkIHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IH0KQGtleWZyYW1lcyB0b2FzdC1pbiB7IGZyb20geyBvcGFjaXR5OiAwOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMTBweCkgc2NhbGUoMC45OCk7IH0gdG8geyBvcGFjaXR5OiAxOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMCkgc2NhbGUoMSk7IH0gfQoKLyogLS0tLS0tLS0tLSBNaXNjIC0tLS0tLS0tLS0gKi8KLm11dGVkIHsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyB9Ci5lbXB0eS1zdGF0ZSB7CiAgcGFkZGluZzogNTZweCAyNHB4OyB0ZXh0LWFsaWduOiBjZW50ZXI7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7CiAgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiAxMnB4OwogIGFuaW1hdGlvbjogY2FyZEluIDI2MG1zIHZhcigtLWVhc2UpOwp9Ci5lbXB0eS1zdGF0ZSAuZW1wdHktaWNvbiB7CiAgd2lkdGg6IDUycHg7IGhlaWdodDogNTJweDsgYm9yZGVyLXJhZGl1czogMTZweDsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsganVzdGlmeS1jb250ZW50OiBjZW50ZXI7CiAgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSAxMCUsIHRyYW5zcGFyZW50KTsgY29sb3I6IHZhcigtLXNlcmllcy0xKTsKfQouZW1wdHktc3RhdGUgLmVtcHR5LXRpdGxlIHsgZm9udC1zaXplOiAxNHB4OyBmb250LXdlaWdodDogNjAwOyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgfQouZW1wdHktc3RhdGUgLmVtcHR5LW1lc3NhZ2UgeyBmb250LXNpemU6IDEzcHg7IG1heC13aWR0aDogMzYwcHg7IH0KLnNwaW5uZXIgeyB3aWR0aDogMTZweDsgaGVpZ2h0OiAxNnB4OyBib3JkZXItcmFkaXVzOiA1MCU7IGJvcmRlcjogMnB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci10b3AtY29sb3I6IHZhcigtLXNlcmllcy0xKTsgYW5pbWF0aW9uOiBzcGluIC42cyBsaW5lYXIgaW5maW5pdGU7IGRpc3BsYXk6IGlubGluZS1ibG9jazsgfQpAa2V5ZnJhbWVzIHNwaW4geyB0byB7IHRyYW5zZm9ybTogcm90YXRlKDM2MGRlZyk7IH0gfQoubG9hZGluZy1yb3cgeyBwYWRkaW5nOiA0MHB4IDIwcHg7IHRleHQtYWxpZ246IGNlbnRlcjsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgfQoKLyogU2tlbGV0b24gbG9hZGVycyDigJQgc2hpbW1lcmluZyBwbGFjZWhvbGRlcnMgc2hvd24gd2hpbGUgYSBzZWN0aW9uJ3MgZGF0YSBpcyBpbiBmbGlnaHQgKi8KLnNrZWxldG9uIHsKICBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtc20pOwogIGJhY2tncm91bmQ6IGxpbmVhci1ncmFkaWVudCgxMDBkZWcsIGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS10ZXh0LW11dGVkKSAxMiUsIHRyYW5zcGFyZW50KSAzMCUsIGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS10ZXh0LW11dGVkKSAyMiUsIHRyYW5zcGFyZW50KSA1MCUsIGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS10ZXh0LW11dGVkKSAxMiUsIHRyYW5zcGFyZW50KSA3MCUpOwogIGJhY2tncm91bmQtc2l6ZTogMjAwJSAxMDAlOwogIGFuaW1hdGlvbjogc2tlbGV0b25TaGltbWVyIDEuNHMgZWFzZS1pbi1vdXQgaW5maW5pdGU7Cn0KQGtleWZyYW1lcyBza2VsZXRvblNoaW1tZXIgeyBmcm9tIHsgYmFja2dyb3VuZC1wb3NpdGlvbjogMTUwJSAwOyB9IHRvIHsgYmFja2dyb3VuZC1wb3NpdGlvbjogLTUwJSAwOyB9IH0KLnNrZWxldG9uLXN0YXQtZ3JpZCB7IGRpc3BsYXk6IGdyaWQ7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogcmVwZWF0KGF1dG8tZml0LCBtaW5tYXgoMTgwcHgsIDFmcikpOyBnYXA6IDE0cHg7IH0KLnNrZWxldG9uLXRpbGUgeyBoZWlnaHQ6IDg0cHg7IH0KLnNrZWxldG9uLWNoYXJ0IHsgaGVpZ2h0OiAyODBweDsgd2lkdGg6IDEwMCU7IH0KLnNrZWxldG9uLXJvdyB7IGhlaWdodDogNDBweDsgbWFyZ2luLWJvdHRvbTogOHB4OyB9CgovKiBBbmltYXRlZCBob3Jpem9udGFsIGNvbXBhcmlzb24gYmFyIOKAlCBhIGxhYmVsZWQgcm93IHdpdGggYSB0cmFjayB0aGF0IGZpbGxzIGluIG9uIGluc2VydGlvbiAqLwouYmFyLXJvdyB7IGRpc3BsYXk6IGdyaWQ7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogbWlubWF4KDkwcHgsIDE0MHB4KSAxZnIgYXV0bzsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiAxMHB4OyBwYWRkaW5nOiA1cHggMDsgfQouYmFyLWxhYmVsIHsgZm9udC1zaXplOiAxMnB4OyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyBmb250LXdlaWdodDogNjAwOyB9Ci5iYXItdHJhY2sgeyBoZWlnaHQ6IDhweDsgYm9yZGVyLXJhZGl1czogNXB4OyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tdGV4dC1tdXRlZCkgMTQlLCB0cmFuc3BhcmVudCk7IG92ZXJmbG93OiBoaWRkZW47IH0KLmJhci1maWxsIHsgaGVpZ2h0OiAxMDAlOyB3aWR0aDogMCU7IGJvcmRlci1yYWRpdXM6IDVweDsgdHJhbnNpdGlvbjogd2lkdGggNzAwbXMgY3ViaWMtYmV6aWVyKDAuMTYsIDEsIDAuMywgMSk7IH0KLmJhci12YWx1ZSB7IGZvbnQtc2l6ZTogMTJweDsgZm9udC13ZWlnaHQ6IDcwMDsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IGZvbnQtdmFyaWFudC1udW1lcmljOiB0YWJ1bGFyLW51bXM7IHRleHQtYWxpZ246IHJpZ2h0OyBtaW4td2lkdGg6IDU2cHg7IH0KCkBtZWRpYSAocHJlZmVycy1yZWR1Y2VkLW1vdGlvbjogcmVkdWNlKSB7CiAgLmJhci1maWxsIHsgdHJhbnNpdGlvbi1kdXJhdGlvbjogMW1zOyB9CiAgLnNrZWxldG9uIHsgYW5pbWF0aW9uLWR1cmF0aW9uOiAxbXM7IH0KICAuY2FyZCwgLnN0YXQtdGlsZSB7IGFuaW1hdGlvbi1kdXJhdGlvbjogMW1zOyB9Cn0KCi50d28tY29sIHsgZGlzcGxheTogZ3JpZDsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiAxZnIgMWZyOyBnYXA6IDE2cHg7IH0KQG1lZGlhIChtYXgtd2lkdGg6IDkwMHB4KSB7IC50d28tY29sIHsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiAxZnI7IH0gfQoKLm1vZGUtdGFicyB7IGRpc3BsYXk6IGZsZXg7IGdhcDogNnB4OyBmbGV4LXdyYXA6IHdyYXA7IG1hcmdpbi1ib3R0b206IDE2cHg7IH0KLm1vZGUtdGFicyBidXR0b24gewogIGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHBhZGRpbmc6IDdweCAxNHB4OyBib3JkZXItcmFkaXVzOiAyMHB4OyBmb250LXNpemU6IDEycHg7IGZvbnQtd2VpZ2h0OiA2MDA7IGN1cnNvcjogcG9pbnRlcjsKICB0cmFuc2l0aW9uOiBjb2xvciAxODBtcyB2YXIoLS1lYXNlKSwgYmFja2dyb3VuZCAxODBtcyB2YXIoLS1lYXNlKSwgdHJhbnNmb3JtIDE1MG1zIHZhcigtLWVhc2UpOwp9Ci5tb2RlLXRhYnMgYnV0dG9uOmhvdmVyIHsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKC0xcHgpOyB9Ci5tb2RlLXRhYnMgYnV0dG9uLmlzLWFjdGl2ZSB7IGJhY2tncm91bmQ6IHZhcigtLXNlcmllcy0xKTsgY29sb3I6ICNmZmY7IGJvcmRlci1jb2xvcjogdHJhbnNwYXJlbnQ7IGJveC1zaGFkb3c6IDAgNHB4IDE0cHggLTVweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDYwJSwgdHJhbnNwYXJlbnQpOyB9CgouZmllbGQtaW5saW5lIHsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA4cHg7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgfQouZmllbGQtaW5saW5lIHNlbGVjdCwgLmZpZWxkLWlubGluZSBpbnB1dCB7IG1pbi13aWR0aDogMDsgcGFkZGluZzogNnB4IDEwcHg7IH0KCi8qIC0tLS0tLS0tLS0gUGxhdGZvcm0gUGVyZm9ybWFuY2UgQ29tcGFyaXNvbiBjYXJkcyAtLS0tLS0tLS0tICovCi5wY2Mtc2VjdGlvbiB7IG1hcmdpbi10b3A6IDI0cHg7IH0KLnBjYy1jb250cm9scyB7IGRpc3BsYXk6IGZsZXg7IGZsZXgtd3JhcDogd3JhcDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiAxNnB4OyBtYXJnaW4tYm90dG9tOiAxNnB4OyB9Ci5wbGF0Zm9ybS1jb21wYXJlLWdyaWQgewogIGRpc3BsYXk6IGdyaWQ7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogcmVwZWF0KDIsIDFmcik7IGdhcDogMTZweDsKfQpAbWVkaWEgKG1heC13aWR0aDogOTAwcHgpIHsgLnBsYXRmb3JtLWNvbXBhcmUtZ3JpZCB7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogMWZyOyB9IH0KLnBsYXRmb3JtLWNvbXBhcmUtY2FyZCB7CiAgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLWxnKTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgcGFkZGluZzogMThweDsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWNhcmQpOwogIHRyYW5zaXRpb246IGJveC1zaGFkb3cgMjIwbXMgdmFyKC0tZWFzZSksIHRyYW5zZm9ybSAyMjBtcyB2YXIoLS1lYXNlKTsKICBhbmltYXRpb246IGNhcmRJbiAzMjBtcyB2YXIoLS1lYXNlKSBiYWNrd2FyZHM7Cn0KLnBsYXRmb3JtLWNvbXBhcmUtY2FyZDpob3ZlciB7IGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdy1ob3Zlcik7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgtMnB4KTsgfQoucGNjLWhlYWRlciB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogc3BhY2UtYmV0d2VlbjsgZ2FwOiAxMHB4OyB9Ci5wY2MtaGVhZGVyLW5hbWUgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDhweDsgfQoucGNjLW5hbWUgeyBmb250LXNpemU6IDE1cHg7IGZvbnQtd2VpZ2h0OiA3MDA7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyB9Ci5wY2MtYmFkZ2UgeyBmb250LXNpemU6IDEzcHg7IGZvbnQtd2VpZ2h0OiA3MDA7IHBhZGRpbmc6IDRweCAxMHB4OyBib3JkZXItcmFkaXVzOiAyMHB4OyB9Ci5wY2MtYmFkZ2UudXAgeyBjb2xvcjogdmFyKC0tc3VjY2Vzcy10ZXh0KTsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXN0YXR1cy1nb29kKSAxNCUsIHRyYW5zcGFyZW50KTsgfQoucGNjLWJhZGdlLmRvd24geyBjb2xvcjogdmFyKC0tc3RhdHVzLWNyaXRpY2FsKTsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXN0YXR1cy1jcml0aWNhbCkgMTIlLCB0cmFuc3BhcmVudCk7IH0KLnBjYy1iYWRnZS5mbGF0IHsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tdGV4dC1tdXRlZCkgMTIlLCB0cmFuc3BhcmVudCk7IH0KLnBjYy1jYXB0aW9uIHsgZm9udC1zaXplOiAxMnB4OyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyBtYXJnaW4tdG9wOiA2cHg7IH0KLnBjYy1tZXRyaWNzIHsgbWFyZ2luLXRvcDogMTZweDsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiAxNHB4OyB9Ci5wY2MtbWV0cmljLXJvdyB7IHBhZGRpbmctdG9wOiAxMnB4OyBib3JkZXItdG9wOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgfQoucGNjLW1ldHJpYy1yb3c6Zmlyc3QtY2hpbGQgeyBwYWRkaW5nLXRvcDogMDsgYm9yZGVyLXRvcDogbm9uZTsgfQoucGNjLW1ldHJpYy1oZWFkZXIgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6IHNwYWNlLWJldHdlZW47IGdhcDogOHB4OyBtYXJnaW4tYm90dG9tOiA2cHg7IH0KLnBjYy1tZXRyaWMtbGFiZWwgeyBmb250LXNpemU6IDEycHg7IGZvbnQtd2VpZ2h0OiA3MDA7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyB9Ci5wY2MtbWV0cmljLWRpZmYgeyBmb250LXNpemU6IDEycHg7IGZvbnQtd2VpZ2h0OiA3MDA7IH0KLnBjYy1tZXRyaWMtZGlmZi51cCB7IGNvbG9yOiB2YXIoLS1zdWNjZXNzLXRleHQpOyB9Ci5wY2MtbWV0cmljLWRpZmYuZG93biB7IGNvbG9yOiB2YXIoLS1zdGF0dXMtY3JpdGljYWwpOyB9Ci5wY2MtbWV0cmljLWRpZmYuZmxhdCB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgfQoucGNjLWZvb3RlciB7IG1hcmdpbi10b3A6IDE2cHg7IHBhZGRpbmctdG9wOiAxNHB4OyBib3JkZXItdG9wOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgfQoucGNjLWZvb3Rlci1sYWJlbCB7IGZvbnQtc2l6ZTogMTFweDsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyBmb250LXdlaWdodDogNjAwOyB0ZXh0LXRyYW5zZm9ybTogdXBwZXJjYXNlOyBsZXR0ZXItc3BhY2luZzogMC4wM2VtOyB9Ci5wY2MtZm9vdGVyLXZhbHVlIHsgZm9udC1zaXplOiAxNXB4OyBmb250LXdlaWdodDogNzAwOyBtYXJnaW4tdG9wOiA0cHg7IH0KLnBjYy1mb290ZXItdmFsdWUudXAgeyBjb2xvcjogdmFyKC0tc3VjY2Vzcy10ZXh0KTsgfQoucGNjLWZvb3Rlci12YWx1ZS5kb3duIHsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IH0KLnBjYy1mb290ZXItdmFsdWUuZmxhdCB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgfQoucGNjLWZvb3Rlci1kZXRhaWwgeyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IG1hcmdpbi10b3A6IDRweDsgfQoucGNjLXZpZXctbGluayB7CiAgZGlzcGxheTogaW5saW5lLWZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogNHB4OyBtYXJnaW4tdG9wOiAxNHB4OwogIGJhY2tncm91bmQ6IG5vbmU7IGJvcmRlcjogbm9uZTsgY29sb3I6IHZhcigtLXNlcmllcy0xKTsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNzAwOyBjdXJzb3I6IHBvaW50ZXI7IHBhZGRpbmc6IDA7CiAgdHJhbnNpdGlvbjogb3BhY2l0eSAxNTBtcyB2YXIoLS1lYXNlKTsKfQoucGNjLXZpZXctbGluazpob3ZlciB7IG9wYWNpdHk6IDAuNzU7IHRleHQtZGVjb3JhdGlvbjogdW5kZXJsaW5lOyB9CgovKiAtLS0tLS0tLS0tIFBhZ2luYXRpb24gLS0tLS0tLS0tLSAqLwoucGFnaW5hdGlvbi1yb3cgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDEycHg7IG1hcmdpbi10b3A6IDE0cHg7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgfQoucGFnaW5hdGlvbi1yb3cgLmJ0biB7IHBhZGRpbmc6IDZweCAxMnB4OyB9Ci5leHBvcnQtYnV0dG9ucyB7IGRpc3BsYXk6IGZsZXg7IGdhcDogOHB4OyBmbGV4LXdyYXA6IHdyYXA7IG1hcmdpbi1ib3R0b206IDEycHg7IH0KLmV4cG9ydC1idXR0b25zIC5idG4geyBwYWRkaW5nOiA3cHggMTNweDsgZm9udC1zaXplOiAxMnB4OyB9CgovKiAtLS0tLS0tLS0tIERhc2hib2FyZCBjb250cm9scyAvIG1ldHJpYy1mb2N1c2VkIEtQSXMgLS0tLS0tLS0tLSAqLwouZGFzaGJvYXJkLWNvbnRyb2xzIHsKICBkaXNwbGF5OiBmbGV4OyBmbGV4LXdyYXA6IHdyYXA7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogMTBweDsgbWFyZ2luLWJvdHRvbTogMThweDsKfQouZGFzaGJvYXJkLWNvbnRyb2xzIGxhYmVsIHsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNjAwOyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyBtYXJnaW4tcmlnaHQ6IDZweDsgfQouZGFzaGJvYXJkLWNvbnRyb2xzIHNlbGVjdCB7IGZvbnQtd2VpZ2h0OiA2MDA7IH0KLyogQmVzdCBQZXJmb3JtaW5nIFBvc3Qg4oCUIGEgZmVhdHVyZWQgbGFuZHNjYXBlIGNhcmQgc3Bhbm5pbmcgMyBLUEktdGlsZS13aWR0aHMKICAgKGEgc3RhbmRhcmQgdGlsZSBpcyAxIHVuaXQ7ICNrcGlHcmlkIGhhcyAxMCB1bml0cyB0b3RhbCksIHNhbWUgZml4ZWQKICAgaGVpZ2h0IGFzIHRoZSByZXN0IG9mICNrcGlHcmlkOiBjYXB0aW9uL3BsYXRmb3JtL2RhdGUgc2l0IG9uIHRoZSBsZWZ0LAogICB3aXRoIHRoZSBzZWxlY3RlZCBtZXRyaWMgKGxhcmdlKSBhbmQgQ3VycmVudCBGb2xsb3dlcnMgKHNtYWxsZXIsIGJlbG93IGEKICAgZGl2aWRlcikgc3RhY2tlZCBpbiBhIG5hcnJvd2VyIGNvbHVtbiBvbiB0aGUgcmlnaHQuICovCiNrcGlHcmlkIC5wb3N0LXRpbGUgewogIGdyaWQtY29sdW1uOiBzcGFuIDM7CiAgZmxleC1kaXJlY3Rpb246IHJvdzsKICBhbGlnbi1pdGVtczogc3RyZXRjaDsKICBqdXN0aWZ5LWNvbnRlbnQ6IGZsZXgtc3RhcnQ7CiAgZ2FwOiAyMHB4OwogIHBhZGRpbmc6IDE0cHggMjBweDsKfQoja3BpR3JpZCAucG9zdC10aWxlLW1haW4geyBmbGV4OiAxIDEgYXV0bzsgbWluLXdpZHRoOiAwOyBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBqdXN0aWZ5LWNvbnRlbnQ6IGNlbnRlcjsgZ2FwOiA2cHg7IH0KI2twaUdyaWQgLnBvc3QtdGlsZS1jYXB0aW9uIHsKICBmb250LXNpemU6IDEzcHg7IGZvbnQtd2VpZ2h0OiA2MDA7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOwogIGxpbmUtaGVpZ2h0OiAxLjQ7CiAgZGlzcGxheTogLXdlYmtpdC1ib3g7IC13ZWJraXQtbGluZS1jbGFtcDogMzsgLXdlYmtpdC1ib3gtb3JpZW50OiB2ZXJ0aWNhbDsgb3ZlcmZsb3c6IGhpZGRlbjsKfQoja3BpR3JpZCAucG9zdC10aWxlLWNhcHRpb24ubXV0ZWQgeyBjb2xvcjogdmFyKC0tdGV4dC1tdXRlZCk7IGZvbnQtd2VpZ2h0OiA1MDA7IC13ZWJraXQtbGluZS1jbGFtcDogMTsgfQoja3BpR3JpZCAucG9zdC10aWxlLW1ldGEgeyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA2cHg7IHdoaXRlLXNwYWNlOiBub3dyYXA7IG92ZXJmbG93OiBoaWRkZW47IHRleHQtb3ZlcmZsb3c6IGVsbGlwc2lzOyB9CiNrcGlHcmlkIC5wb3N0LXRpbGUtZGl2aWRlciB7IGZsZXg6IDAgMCBhdXRvOyB3aWR0aDogMXB4OyBhbGlnbi1zZWxmOiBzdHJldGNoOyBiYWNrZ3JvdW5kOiB2YXIoLS1ib3JkZXIpOyB9CiNrcGlHcmlkIC5wb3N0LXRpbGUtbWV0cmljcyB7CiAgZmxleDogMCAwIGF1dG87IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGp1c3RpZnktY29udGVudDogY2VudGVyOyBnYXA6IDhweDsgbWluLXdpZHRoOiAxMjBweDsKfQoja3BpR3JpZCAucG9zdC10aWxlLW1ldHJpYy1ibG9jayB7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGdhcDogM3B4OyB9CiNrcGlHcmlkIC5wb3N0LXRpbGUtbWV0cmljLWJsb2NrLnNlY29uZGFyeSB7IHBhZGRpbmctdG9wOiA4cHg7IGJvcmRlci10b3A6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyB9CiNrcGlHcmlkIC5wb3N0LXRpbGUtbWV0cmljLWxhYmVsIHsgZm9udC1zaXplOiAxMXB4OyBjb2xvcjogdmFyKC0tdGV4dC1tdXRlZCk7IGZvbnQtd2VpZ2h0OiA2MDA7IHRleHQtdHJhbnNmb3JtOiB1cHBlcmNhc2U7IGxldHRlci1zcGFjaW5nOiAwLjAzZW07IH0KI2twaUdyaWQgLnBvc3QtdGlsZS1tZXRyaWMtdmFsdWUgeyBmb250LXNpemU6IDIycHg7IGZvbnQtd2VpZ2h0OiA3MDA7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyBmb250LXZhcmlhbnQtbnVtZXJpYzogdGFidWxhci1udW1zOyB9CiNrcGlHcmlkIC5wb3N0LXRpbGUtbWV0cmljLWJsb2NrLnNlY29uZGFyeSAucG9zdC10aWxlLW1ldHJpYy12YWx1ZSB7IGZvbnQtc2l6ZTogMTVweDsgfQpAbWVkaWEgKG1heC13aWR0aDogNjQwcHgpIHsgI2twaUdyaWQgLnBvc3QtdGlsZSB7IGdyaWQtY29sdW1uOiBzcGFuIDI7IH0gfQoKLnN0YXQtdmFsdWUtbXV0ZWQgeyBmb250LXNpemU6IDE1cHggIWltcG9ydGFudDsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyBmb250LXdlaWdodDogNjAwOyB9Ci5jYXB0aW9uLWxpbmsgeyBjb2xvcjogdmFyKC0tc2VyaWVzLTEpOyB0ZXh0LWRlY29yYXRpb246IG5vbmU7IH0KLmNhcHRpb24tbGluazpob3ZlciB7IHRleHQtZGVjb3JhdGlvbjogdW5kZXJsaW5lOyB9CgovKiAtLS0tLS0tLS0tIERhdGEgUmVjb3JkcyAocGxhdGZvcm0tZ3JvdXBlZCkgLS0tLS0tLS0tLSAqLwoucmVjb3Jkcy10b29sYmFyIHsKICBkaXNwbGF5OiBmbGV4OyBmbGV4LXdyYXA6IHdyYXA7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogc3BhY2UtYmV0d2VlbjsKICBnYXA6IDEycHg7IG1hcmdpbi1ib3R0b206IDE0cHg7Cn0KLnBsYXRmb3JtLWZpbHRlci1waWxscyB7IGRpc3BsYXk6IGZsZXg7IGZsZXgtd3JhcDogd3JhcDsgZ2FwOiA2cHg7IH0KLnBsYXRmb3JtLWZpbHRlci1waWxscyBidXR0b24gewogIGRpc3BsYXk6IGlubGluZS1mbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDZweDsKICBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTEpOyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOwogIGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBwYWRkaW5nOiA3cHggMTRweDsgYm9yZGVyLXJhZGl1czogMjBweDsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNjAwOyBjdXJzb3I6IHBvaW50ZXI7CiAgdHJhbnNpdGlvbjogY29sb3IgMTgwbXMgdmFyKC0tZWFzZSksIGJhY2tncm91bmQgMTgwbXMgdmFyKC0tZWFzZSksIHRyYW5zZm9ybSAxNTBtcyB2YXIoLS1lYXNlKSwgYm94LXNoYWRvdyAxODBtcyB2YXIoLS1lYXNlKTsKfQoucGxhdGZvcm0tZmlsdGVyLXBpbGxzIGJ1dHRvbjpob3ZlciB7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoLTFweCk7IH0KLnBsYXRmb3JtLWZpbHRlci1waWxscyBidXR0b246YWN0aXZlIHsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDApIHNjYWxlKDAuOTYpOyB9Ci5wbGF0Zm9ybS1maWx0ZXItcGlsbHMgYnV0dG9uLmlzLWFjdGl2ZSB7IGJhY2tncm91bmQ6IHZhcigtLXNlcmllcy0xKTsgY29sb3I6ICNmZmY7IGJvcmRlci1jb2xvcjogdHJhbnNwYXJlbnQ7IGJveC1zaGFkb3c6IDAgNHB4IDE0cHggLTVweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDYwJSwgdHJhbnNwYXJlbnQpOyB9Ci5wbGF0Zm9ybS1maWx0ZXItcGlsbHMgYnV0dG9uLmlzLWFjdGl2ZSAucGxhdGZvcm0tZG90IHsgYm94LXNoYWRvdzogMCAwIDAgMnB4IHJnYmEoMjU1LDI1NSwyNTUsMC41KTsgfQoucmVjb3Jkcy1zZWFyY2ggaW5wdXQgeyBib3JkZXItcmFkaXVzOiAyMHB4OyBtaW4td2lkdGg6IDIyMHB4OyB9Ci5zdGF0dXMtcGlsbCB7IGRpc3BsYXk6IGlubGluZS1ibG9jazsgcGFkZGluZzogM3B4IDEwcHg7IGJvcmRlci1yYWRpdXM6IDIwcHg7IGZvbnQtc2l6ZTogMTFweDsgZm9udC13ZWlnaHQ6IDcwMDsgfQouc3RhdHVzLXBpbGwub3JpZ2luYWwgeyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tdGV4dC1tdXRlZCkgMTUlLCB0cmFuc3BhcmVudCk7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IH0KLnN0YXR1cy1waWxsLmVkaXRlZCB7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zdGF0dXMtd2FybmluZykgMjIlLCB0cmFuc3BhcmVudCk7IGNvbG9yOiAjOGE2MzAwOyB9Ci5yb3ctYWN0aW9ucyB7IGRpc3BsYXk6IGZsZXg7IGdhcDogNnB4OyBmbGV4LXdyYXA6IG5vd3JhcDsgfQoucm93LWFjdGlvbnMgLmJ0biB7IHBhZGRpbmc6IDVweCAxMHB4OyBmb250LXNpemU6IDEycHg7IH0KLmxpbmstY2VsbCBhIHsgY29sb3I6IHZhcigtLXNlcmllcy0xKTsgdGV4dC1kZWNvcmF0aW9uOiBub25lOyBmb250LXdlaWdodDogNjAwOyBmb250LXNpemU6IDEycHg7IH0KLmxpbmstY2VsbCBhOmhvdmVyIHsgdGV4dC1kZWNvcmF0aW9uOiB1bmRlcmxpbmU7IH0KLnJlY29yZC1zZWN0aW9uIHsKICBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtc20pOyBwYWRkaW5nOiAxNnB4OyBtYXJnaW4tYm90dG9tOiAxNHB4OwogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMik7IGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKfQoucmVjb3JkLXNlY3Rpb24gaDQgeyBtYXJnaW46IDAgMCAxMnB4OyBmb250LXNpemU6IDEycHg7IGZvbnQtd2VpZ2h0OiA3MDA7IGxldHRlci1zcGFjaW5nOiAwLjAzZW07IHRleHQtdHJhbnNmb3JtOiB1cHBlcmNhc2U7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IH0KLnJlY29yZC1zZWN0aW9uIC5mb3JtLWdyaWQgeyBtYXJnaW4tYm90dG9tOiAwOyB9Ci5yZWNvcmQtc2VjdGlvbiAudmlldy1maWVsZCB7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGdhcDogMnB4OyBmb250LXNpemU6IDEzcHg7IH0KLnJlY29yZC1zZWN0aW9uIC52aWV3LWZpZWxkIC52aWV3LWxhYmVsIHsgZm9udC1zaXplOiAxMXB4OyBmb250LXdlaWdodDogNjAwOyBjb2xvcjogdmFyKC0tdGV4dC1tdXRlZCk7IH0KLnJlY29yZC1zZWN0aW9uIC52aWV3LWZpZWxkIC52aWV3LXZhbHVlIHsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IHdvcmQtYnJlYWs6IGJyZWFrLXdvcmQ7IH0KQG1lZGlhIChtYXgtd2lkdGg6IDY0MHB4KSB7CiAgLnJlY29yZHMtdG9vbGJhciB7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGFsaWduLWl0ZW1zOiBzdHJldGNoOyB9CiAgLnJlY29yZHMtc2VhcmNoIGlucHV0IHsgd2lkdGg6IDEwMCU7IH0KfQoKLyogLS0tLS0tLS0tLSBNb2RhbCAocmVjb3JkIGVkaXRvcikgLS0tLS0tLS0tLSAqLwoubW9kYWwtb3ZlcmxheSB7CiAgcG9zaXRpb246IGZpeGVkOyBpbnNldDogMDsgYmFja2dyb3VuZDogcmdiYSgxMCwxMSwxMywwLjUpOwogIGJhY2tkcm9wLWZpbHRlcjogYmx1cig2cHgpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogYmx1cig2cHgpOwogIGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBmbGV4LXN0YXJ0OyBqdXN0aWZ5LWNvbnRlbnQ6IGNlbnRlcjsKICBwYWRkaW5nOiA0MHB4IDE2cHg7IG92ZXJmbG93LXk6IGF1dG87IHotaW5kZXg6IDIwMDsKICBhbmltYXRpb246IG92ZXJsYXlJbiAyMDBtcyB2YXIoLS1lYXNlKTsKfQpAa2V5ZnJhbWVzIG92ZXJsYXlJbiB7IGZyb20geyBvcGFjaXR5OiAwOyB9IHRvIHsgb3BhY2l0eTogMTsgfSB9CkBrZXlmcmFtZXMgbW9kYWxQYW5lbEluIHsKICBmcm9tIHsgb3BhY2l0eTogMDsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDE0cHgpIHNjYWxlKDAuOTcpOyB9CiAgdG8geyBvcGFjaXR5OiAxOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMCkgc2NhbGUoMSk7IH0KfQoubW9kYWwtcGFuZWwgewogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1sZyk7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHBhZGRpbmc6IDI0cHg7IHdpZHRoOiAxMDAlOyBtYXgtd2lkdGg6IDcyMHB4OyBib3gtc2hhZG93OiB2YXIoLS1zaGFkb3ctbW9kYWwpOwogIG1heC1oZWlnaHQ6IGNhbGMoMTAwdmggLSA4MHB4KTsgb3ZlcmZsb3cteTogYXV0bzsKICBhbmltYXRpb246IG1vZGFsUGFuZWxJbiAyNDBtcyB2YXIoLS1lYXNlKTsKfQoubW9kYWwtcGFuZWwud2lkZSB7IG1heC13aWR0aDogMTEwMHB4OyB9Ci5tb2RhbC1wYW5lbCBoMiB7IG1hcmdpbjogMCAwIDRweDsgZm9udC1zaXplOiAxN3B4OyBsZXR0ZXItc3BhY2luZzogLTAuMDFlbTsgfQoubW9kYWwtcGFuZWwgLm1vZGFsLXN1YiB7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IGZvbnQtc2l6ZTogMTJweDsgbWFyZ2luOiAwIDAgMThweDsgfQouZm9ybS1ncmlkIHsgZGlzcGxheTogZ3JpZDsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiByZXBlYXQoYXV0by1maXQsIG1pbm1heCgyMDBweCwgMWZyKSk7IGdhcDogMTJweDsgbWFyZ2luLWJvdHRvbTogMTZweDsgfQouZm9ybS1ncmlkLmZ1bGwgeyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IDFmcjsgfQpAbWVkaWEgKG1heC13aWR0aDogNjQwcHgpIHsgLmZvcm0tZ3JpZCB7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogMWZyOyB9IH0KLmZvcm0tZmllbGQgeyBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBnYXA6IDVweDsgZm9udC1zaXplOiAxMnB4OyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyB9Ci5mb3JtLWZpZWxkIGxhYmVsIHsgZm9udC13ZWlnaHQ6IDYwMDsgfQouZm9ybS1maWVsZCB0ZXh0YXJlYSB7IHJlc2l6ZTogdmVydGljYWw7IG1pbi1oZWlnaHQ6IDYwcHg7IH0KCi5wbGF0Zm9ybS1lZGl0LXJvdyB7CiAgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLXNtKTsgcGFkZGluZzogMTRweDsgbWFyZ2luLWJvdHRvbTogMTBweDsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0yKTsKfQoucGxhdGZvcm0tZWRpdC1yb3cgLnBsYXRmb3JtLWVkaXQtaGVhZCB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogc3BhY2UtYmV0d2VlbjsgZ2FwOiA4cHg7IG1hcmdpbi1ib3R0b206IDEwcHg7IH0KLnBsYXRmb3JtLWVkaXQtcm93IC5tZXRyaWNzLWdyaWQgeyBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdChhdXRvLWZpdCwgbWlubWF4KDEyMHB4LCAxZnIpKTsgZ2FwOiA4cHg7IH0KLnJlbW92ZS1wbGF0Zm9ybS1idG4geyBib3JkZXI6IG5vbmU7IGJhY2tncm91bmQ6IHRyYW5zcGFyZW50OyBjb2xvcjogdmFyKC0tc3RhdHVzLWNyaXRpY2FsKTsgY3Vyc29yOiBwb2ludGVyOyBmb250LXNpemU6IDEycHg7IGZvbnQtd2VpZ2h0OiA2MDA7IHRyYW5zaXRpb246IG9wYWNpdHkgMTUwbXMgdmFyKC0tZWFzZSk7IH0KLnJlbW92ZS1wbGF0Zm9ybS1idG46aG92ZXIgeyBvcGFjaXR5OiAwLjc7IH0KLm1vZGFsLWFjdGlvbnMgeyBkaXNwbGF5OiBmbGV4OyBqdXN0aWZ5LWNvbnRlbnQ6IHNwYWNlLWJldHdlZW47IGFsaWduLWl0ZW1zOiBjZW50ZXI7IG1hcmdpbi10b3A6IDE4cHg7IGdhcDogOHB4OyBmbGV4LXdyYXA6IHdyYXA7IH0KCi8qIC0tLS0tLS0tLS0gUmVzcG9uc2l2ZSB0aWdodGVuaW5nIC0tLS0tLS0tLS0gKi8KQG1lZGlhIChtYXgtd2lkdGg6IDcyMHB4KSB7CiAgLmFwcC1zaGVsbCB7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IH0KICAuc2lkZWJhciB7IHdpZHRoOiAxMDAlOyBoZWlnaHQ6IGF1dG87IHBvc2l0aW9uOiBzdGF0aWM7IGZsZXgtZGlyZWN0aW9uOiByb3c7IGZsZXgtd3JhcDogd3JhcDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA4cHg7IHBhZGRpbmc6IDEwcHggMTRweDsgfQogIC5zaWRlYmFyLWJyYW5kIHsgcGFkZGluZzogMDsgbWFyZ2luLXJpZ2h0OiBhdXRvOyB9CiAgLnRhYnMgeyBmbGV4LWRpcmVjdGlvbjogcm93OyB3aWR0aDogMTAwJTsgb3ZlcmZsb3cteDogYXV0bzsgb3JkZXI6IDM7IH0KICAuc2lkZWJhci1mb290ZXIgeyBmbGV4LWRpcmVjdGlvbjogcm93OyBib3JkZXItdG9wOiBub25lOyBtYXJnaW4tdG9wOiAwOyBwYWRkaW5nLXRvcDogMDsgfQogIC52aWV3LWFyZWEgeyBwYWRkaW5nOiAxNHB4OyB9CiAgLmZpbHRlci1iYXIgeyB0b3A6IGF1dG87IHBvc2l0aW9uOiBzdGF0aWM7IHBhZGRpbmc6IDEycHggMTRweDsgfQogIC5zdGF0LWdyaWQgeyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdChhdXRvLWZpdCwgbWlubWF4KDE0MHB4LCAxZnIpKTsgfQogIC5icmFuZC1sb2dvIHsgaGVpZ2h0OiAyMnB4OyB9Cn0KCi8qIC0tLS0tLS0tLS0gUmVwb3J0IEdlbmVyYXRvciAtLS0tLS0tLS0tICovCi5yZXBvcnQtY29udHJvbHMgeyBkaXNwbGF5OiBmbGV4OyBmbGV4LXdyYXA6IHdyYXA7IGdhcDogMTRweCAxOHB4OyBhbGlnbi1pdGVtczogZW5kOyB9Ci5yZXBvcnQtYWN0aW9ucyB7IGRpc3BsYXk6IGZsZXg7IGdhcDogOHB4OyBmbGV4LXdyYXA6IHdyYXA7IG1hcmdpbi10b3A6IDRweDsgfQoucmVwb3J0LWRvYy1oZWFkIHsgbWFyZ2luLWJvdHRvbTogNHB4OyB9Ci5yZXBvcnQtZG9jLWhlYWQgLnJlcG9ydC10aXRsZSB7IGZvbnQtc2l6ZTogMjBweDsgZm9udC13ZWlnaHQ6IDcwMDsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IGxldHRlci1zcGFjaW5nOiAtMC4wMmVtOyB9Ci5yZXBvcnQtZG9jLWhlYWQgLnJlcG9ydC1yYW5nZSB7IGZvbnQtc2l6ZTogMTNweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgbWFyZ2luLXRvcDogNHB4OyB9Ci5yZXBvcnQtZG9jLWhlYWQgLnJlcG9ydC1nZW5lcmF0ZWQgeyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgbWFyZ2luLXRvcDogMnB4OyB9Ci5yZXBvcnQtbWV0cmljLXNlbGVjdCB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogOHB4OyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IG1hcmdpbi1ib3R0b206IDEycHg7IH0KLnJlcG9ydC1zdW1tYXJ5IHsgbGluZS1oZWlnaHQ6IDEuNjsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgZm9udC1zaXplOiAxMy41cHg7IH0KLnJlcG9ydC1zdW1tYXJ5IGg0IHsgZm9udC1zaXplOiAxM3B4OyBmb250LXdlaWdodDogNzAwOyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgdGV4dC10cmFuc2Zvcm06IHVwcGVyY2FzZTsgbGV0dGVyLXNwYWNpbmc6IDAuMDRlbTsgbWFyZ2luOiAxOHB4IDAgOHB4OyB9Ci5yZXBvcnQtc3VtbWFyeSBoNDpmaXJzdC1jaGlsZCB7IG1hcmdpbi10b3A6IDA7IH0KLnJlcG9ydC1zdW1tYXJ5IHAgeyBtYXJnaW46IDAgMCA2cHg7IH0KLnJlcG9ydC1zdW1tYXJ5IHVsIHsgbWFyZ2luOiAwOyBwYWRkaW5nLWxlZnQ6IDE4cHg7IH0KLnJlcG9ydC1zdW1tYXJ5IGxpIHsgbWFyZ2luOiA0cHggMDsgfQoucmVwb3J0LXN1bW1hcnkgLnRyZW5kLXVwIHsgY29sb3I6IHZhcigtLXN1Y2Nlc3MtdGV4dCk7IGZvbnQtd2VpZ2h0OiA2MDA7IH0KLnJlcG9ydC1zdW1tYXJ5IC50cmVuZC1kb3duIHsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IGZvbnQtd2VpZ2h0OiA2MDA7IH0KLmRhdGEtdGFibGUgdGQudHJlbmQtdXAgeyBjb2xvcjogdmFyKC0tc3VjY2Vzcy10ZXh0KTsgZm9udC13ZWlnaHQ6IDYwMDsgfQouZGF0YS10YWJsZSB0ZC50cmVuZC1kb3duIHsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IGZvbnQtd2VpZ2h0OiA2MDA7IH0KLmRhdGEtdGFibGUgdGQudHJlbmQtZmxhdCB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgfQoucmVwb3J0LXBsYXRmb3JtLWNhcmQgeyBtYXJnaW4tYm90dG9tOiA4cHg7IH0KLnJlcG9ydC1wbGF0Zm9ybS1jYXJkICsgLnRhYmxlLXNjcm9sbCB7IG1hcmdpbi10b3A6IDhweDsgfQoudGFibGUtc2Nyb2xsICsgLnNlY3Rpb24tdGl0bGUgeyBtYXJnaW4tdG9wOiAyNnB4OyB9Ci5yZXBvcnQtc3RhdHVzIHsgZGlzcGxheTogaW5saW5lLWJsb2NrOyBwYWRkaW5nOiAycHggOXB4OyBib3JkZXItcmFkaXVzOiA5OTlweDsgZm9udC1zaXplOiAxMXB4OyBmb250LXdlaWdodDogNzAwOyB3aGl0ZS1zcGFjZTogbm93cmFwOyB9Ci5yZXBvcnQtc3RhdHVzLnRyZW5kLXVwIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXN1Y2Nlc3MtdGV4dCkgMTYlLCB0cmFuc3BhcmVudCk7IGNvbG9yOiB2YXIoLS1zdWNjZXNzLXRleHQpOyB9Ci5yZXBvcnQtc3RhdHVzLnRyZW5kLWRvd24geyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc3RhdHVzLWNyaXRpY2FsKSAxNiUsIHRyYW5zcGFyZW50KTsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IH0KLnJlcG9ydC1zdGF0dXMudHJlbmQtZmxhdCB7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMik7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgfQoKLyogLS0tLS0tLS0tLSBQcmludCAvIEV4cG9ydCAoUmVwb3J0IEdlbmVyYXRvciAiUHJpbnQgLyBFeHBvcnQgUmVwb3J0IikgLS0tLS0tLS0tLSAqLwpAbWVkaWEgcHJpbnQgewogIC5zaWRlYmFyLCAjZmlsdGVyQmFyLCAjdG9hc3RSb290LCAjdmlldy1yZXBvcnQgPiAuY2FyZCwgLnJlcG9ydC1tZXRyaWMtc2VsZWN0IHsgZGlzcGxheTogbm9uZSAhaW1wb3J0YW50OyB9CiAgLmFwcC1zaGVsbCwgLm1haW4tY29sLCAudmlldy1hcmVhIHsgZGlzcGxheTogYmxvY2sgIWltcG9ydGFudDsgYmFja2dyb3VuZDogI2ZmZiAhaW1wb3J0YW50OyB9CiAgLnZpZXctYXJlYSB7IHBhZGRpbmc6IDAgIWltcG9ydGFudDsgfQogIC52aWV3Om5vdCguaXMtYWN0aXZlKSB7IGRpc3BsYXk6IG5vbmUgIWltcG9ydGFudDsgfQogIC52aWV3LmlzLWFjdGl2ZSB7IGRpc3BsYXk6IGJsb2NrICFpbXBvcnRhbnQ7IH0KICBib2R5IHsgYmFja2dyb3VuZDogI2ZmZiAhaW1wb3J0YW50OyBjb2xvcjogIzExMSAhaW1wb3J0YW50OyB9CiAgLmNhcmQgeyBiYWNrZ3JvdW5kOiAjZmZmICFpbXBvcnRhbnQ7IGJvcmRlcjogMXB4IHNvbGlkICNkZGQgIWltcG9ydGFudDsgYm94LXNoYWRvdzogbm9uZSAhaW1wb3J0YW50OyBiYWNrZHJvcC1maWx0ZXI6IG5vbmUgIWltcG9ydGFudDsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IG5vbmUgIWltcG9ydGFudDsgYnJlYWstaW5zaWRlOiBhdm9pZDsgfQogIC5zZWN0aW9uLXRpdGxlLCAuc3RhdC12YWx1ZSwgLnJlcG9ydC1kb2MtaGVhZCAucmVwb3J0LXRpdGxlLCAucmVwb3J0LXN1bW1hcnkgaDQgeyBjb2xvcjogIzExMSAhaW1wb3J0YW50OyB9CiAgLnN0YXQtdGlsZSB7IGJhY2tncm91bmQ6ICNmZmYgIWltcG9ydGFudDsgYm9yZGVyOiAxcHggc29saWQgI2U2ZTZlNiAhaW1wb3J0YW50OyBib3gtc2hhZG93OiBub25lICFpbXBvcnRhbnQ7IH0KICAuc3RhdC1sYWJlbCwgLnN0YXQtZGVsdGEsIC5yZXBvcnQtZG9jLWhlYWQgLnJlcG9ydC1yYW5nZSwgLnJlcG9ydC1kb2MtaGVhZCAucmVwb3J0LWdlbmVyYXRlZCwgLnJlcG9ydC1zdW1tYXJ5IHsgY29sb3I6ICM0NDQgIWltcG9ydGFudDsgfQogIC5zdGF0LWRlbHRhLnVwLCAuZGF0YS10YWJsZSB0ZC50cmVuZC11cCB7IGNvbG9yOiAjMWE3ZjM3ICFpbXBvcnRhbnQ7IH0KICAuc3RhdC1kZWx0YS5kb3duLCAuZGF0YS10YWJsZSB0ZC50cmVuZC1kb3duIHsgY29sb3I6ICNiMzI2MWUgIWltcG9ydGFudDsgfQogIC5yZXBvcnQtc3RhdHVzLnRyZW5kLXVwIHsgYmFja2dyb3VuZDogI2UzZjRlOCAhaW1wb3J0YW50OyBjb2xvcjogIzFhN2YzNyAhaW1wb3J0YW50OyB9CiAgLnJlcG9ydC1zdGF0dXMudHJlbmQtZG93biB7IGJhY2tncm91bmQ6ICNmYmU2ZTUgIWltcG9ydGFudDsgY29sb3I6ICNiMzI2MWUgIWltcG9ydGFudDsgfQogIC5yZXBvcnQtc3RhdHVzLnRyZW5kLWZsYXQgeyBiYWNrZ3JvdW5kOiAjZWVlICFpbXBvcnRhbnQ7IGNvbG9yOiAjNDQ0ICFpbXBvcnRhbnQ7IH0KICAuZGF0YS10YWJsZSB0aCwgLmRhdGEtdGFibGUgdGQgeyBjb2xvcjogIzExMSAhaW1wb3J0YW50OyBib3JkZXItY29sb3I6ICNkZGQgIWltcG9ydGFudDsgfQogIC5kYXRhLXRhYmxlIHRoZWFkIHRoIHsgY29sb3I6ICM0NDQgIWltcG9ydGFudDsgfQogIGEsIC5jYXB0aW9uLWxpbmsgeyBjb2xvcjogIzFhNGY4YiAhaW1wb3J0YW50OyB9CiAgLmNoYXJ0LXdyYXAsIC5jaGFydC13cmFwLnRhbGwgeyBoZWlnaHQ6IDMyMHB4ICFpbXBvcnRhbnQ7IHBhZ2UtYnJlYWstaW5zaWRlOiBhdm9pZDsgfQogIC5zZWN0aW9uLXRpdGxlIHsgcGFnZS1icmVhay1hZnRlcjogYXZvaWQ7IH0KICBAcGFnZSB7IG1hcmdpbjogMTRtbTsgfQp9Cjwvc3R5bGU+CjwvaGVhZD4KPGJvZHk+CjxkaXYgY2xhc3M9ImF1dGgtc2NyZWVuIiBpZD0iYXV0aFNjcmVlbiI+CiAgPGRpdiBjbGFzcz0iYXV0aC1jYXJkIj4KICAgIDxkaXYgY2xhc3M9ImF1dGgtYnJhbmQiPgogICAgICA8aW1nIGNsYXNzPSJicmFuZC1sb2dvIiBhbHQ9IkxpZ29uLVJhem9uIFNvbHV0aW9ucyBsb2dvIiAvPgogICAgICA8c3BhbiBjbGFzcz0iYnJhbmQtdGl0bGUiPlNvY2lhbCBNZWRpYSBBbmFseXRpY3M8L3NwYW4+CiAgICA8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImF1dGgtZm9ybSI+CiAgICAgIDxkaXYgY2xhc3M9ImZvcm0tZmllbGQiPgogICAgICAgIDxsYWJlbCBmb3I9ImF1dGhDb2RlIj5BY2Nlc3MgY29kZTwvbGFiZWw+CiAgICAgICAgPGlucHV0IHR5cGU9InBhc3N3b3JkIiBpZD0iYXV0aENvZGUiIGF1dG9jb21wbGV0ZT0ib2ZmIiAvPgogICAgICA8L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0iYXV0aC1lcnJvciIgaWQ9ImF1dGhFcnJvciI+PC9kaXY+CiAgICAgIDxidXR0b24gY2xhc3M9ImJ0biBwcmltYXJ5IiBpZD0iYXV0aFN1Ym1pdEJ0biIgdHlwZT0iYnV0dG9uIj48aSBkYXRhLWx1Y2lkZT0iYXJyb3ctcmlnaHQiIHN0eWxlPSJ3aWR0aDoxNHB4O2hlaWdodDoxNHB4OyI+PC9pPiBFbnRlcjwvYnV0dG9uPgogICAgPC9kaXY+CiAgPC9kaXY+CjwvZGl2PgoKPGRpdiBjbGFzcz0iYXBwLXNoZWxsIiBpZD0iYXBwU2hlbGwiIHN0eWxlPSJkaXNwbGF5Om5vbmU7Ij4KICA8YXNpZGUgY2xhc3M9InNpZGViYXIiPgogICAgPGRpdiBjbGFzcz0ic2lkZWJhci1icmFuZCI+CiAgICAgIDxpbWcgY2xhc3M9ImJyYW5kLWxvZ28iIGFsdD0iTGlnb24tUmF6b24gU29sdXRpb25zIGxvZ28iIC8+CiAgICAgIDxzcGFuIGNsYXNzPSJicmFuZC10aXRsZSI+U29jaWFsIE1lZGlhIEFuYWx5dGljczwvc3Bhbj4KICAgIDwvZGl2PgogICAgPG5hdiBjbGFzcz0idGFicyIgcm9sZT0idGFibGlzdCIgYXJpYS1sYWJlbD0iU2VjdGlvbnMiPgogICAgICA8YnV0dG9uIGNsYXNzPSJ0YWItYnRuIGlzLWFjdGl2ZSIgZGF0YS10YWI9ImRhc2hib2FyZCIgcm9sZT0idGFiIiBhcmlhLXNlbGVjdGVkPSJ0cnVlIj48aSBkYXRhLWx1Y2lkZT0ibGF5b3V0LWRhc2hib2FyZCIgc3R5bGU9IndpZHRoOjE0cHg7aGVpZ2h0OjE0cHg7Ij48L2k+IERhc2hib2FyZDwvYnV0dG9uPgogICAgICA8YnV0dG9uIGNsYXNzPSJ0YWItYnRuIiBkYXRhLXRhYj0icmVjb3JkcyIgcm9sZT0idGFiIiBhcmlhLXNlbGVjdGVkPSJmYWxzZSI+PGkgZGF0YS1sdWNpZGU9ImRhdGFiYXNlIiBzdHlsZT0id2lkdGg6MTRweDtoZWlnaHQ6MTRweDsiPjwvaT4gRGF0YSBSZWNvcmRzPC9idXR0b24+CiAgICAgIDxidXR0b24gY2xhc3M9InRhYi1idG4iIGRhdGEtdGFiPSJmb2xsb3dlcnMiIHJvbGU9InRhYiIgYXJpYS1zZWxlY3RlZD0iZmFsc2UiPjxpIGRhdGEtbHVjaWRlPSJ1c2VycyIgc3R5bGU9IndpZHRoOjE0cHg7aGVpZ2h0OjE0cHg7Ij48L2k+IEZvbGxvd2VycyBEYXRhPC9idXR0b24+CiAgICAgIDxidXR0b24gY2xhc3M9InRhYi1idG4iIGRhdGEtdGFiPSJjb21wYXJpc29uIiByb2xlPSJ0YWIiIGFyaWEtc2VsZWN0ZWQ9ImZhbHNlIj48aSBkYXRhLWx1Y2lkZT0iZ2l0LWNvbXBhcmUiIHN0eWxlPSJ3aWR0aDoxNHB4O2hlaWdodDoxNHB4OyI+PC9pPiBDb21wYXJpc29uczwvYnV0dG9uPgogICAgICA8YnV0dG9uIGNsYXNzPSJ0YWItYnRuIiBkYXRhLXRhYj0icmVwb3J0IiByb2xlPSJ0YWIiIGFyaWEtc2VsZWN0ZWQ9ImZhbHNlIj48aSBkYXRhLWx1Y2lkZT0iZmlsZS10ZXh0IiBzdHlsZT0id2lkdGg6MTRweDtoZWlnaHQ6MTRweDsiPjwvaT4gUmVwb3J0IEdlbmVyYXRvcjwvYnV0dG9uPgogICAgICA8YnV0dG9uIGNsYXNzPSJ0YWItYnRuIiBkYXRhLXRhYj0idXBsb2FkIiByb2xlPSJ0YWIiIGFyaWEtc2VsZWN0ZWQ9ImZhbHNlIj48aSBkYXRhLWx1Y2lkZT0idXBsb2FkLWNsb3VkIiBzdHlsZT0id2lkdGg6MTRweDtoZWlnaHQ6MTRweDsiPjwvaT4gVXBsb2FkIERhdGE8L2J1dHRvbj4KICAgICAgPGJ1dHRvbiBjbGFzcz0idGFiLWJ0biIgZGF0YS10YWI9Imhpc3RvcnkiIHJvbGU9InRhYiIgYXJpYS1zZWxlY3RlZD0iZmFsc2UiPjxpIGRhdGEtbHVjaWRlPSJoaXN0b3J5IiBzdHlsZT0id2lkdGg6MTRweDtoZWlnaHQ6MTRweDsiPjwvaT4gVXBsb2FkIEhpc3Rvcnk8L2J1dHRvbj4KICAgIDwvbmF2PgogICAgPGRpdiBjbGFzcz0ic2lkZWJhci1mb290ZXIiPgogICAgICA8YnV0dG9uIGNsYXNzPSJidG4iIGlkPSJsb2dvdXRCdG4iIHR5cGU9ImJ1dHRvbiI+PGkgZGF0YS1sdWNpZGU9ImxvY2siIHN0eWxlPSJ3aWR0aDoxNHB4O2hlaWdodDoxNHB4OyI+PC9pPiBMb2NrPC9idXR0b24+CiAgICA8L2Rpdj4KICA8L2FzaWRlPgoKICA8ZGl2IGNsYXNzPSJtYWluLWNvbCI+CiAgPHNlY3Rpb24gY2xhc3M9ImZpbHRlci1iYXIiIGlkPSJmaWx0ZXJCYXIiPgogICAgPGRpdiBjbGFzcz0iZmlsdGVyLWZpZWxkIj4KICAgICAgPGxhYmVsIGZvcj0iZmlsdGVyRGF0ZUZyb20iPkZyb208L2xhYmVsPgogICAgICA8aW5wdXQgdHlwZT0iZGF0ZSIgaWQ9ImZpbHRlckRhdGVGcm9tIiAvPgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJmaWx0ZXItZmllbGQiPgogICAgICA8bGFiZWwgZm9yPSJmaWx0ZXJEYXRlVG8iPlRvPC9sYWJlbD4KICAgICAgPGlucHV0IHR5cGU9ImRhdGUiIGlkPSJmaWx0ZXJEYXRlVG8iIC8+CiAgICA8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImZpbHRlci1maWVsZCBmaWx0ZXItcHJlc2V0cyIgaWQ9ImZpbHRlclByZXNldHMiPgogICAgICA8YnV0dG9uIHR5cGU9ImJ1dHRvbiIgZGF0YS1wcmVzZXQ9IjciPkxhc3QgNyBkYXlzPC9idXR0b24+CiAgICAgIDxidXR0b24gdHlwZT0iYnV0dG9uIiBkYXRhLXByZXNldD0iMzAiPkxhc3QgMzAgZGF5czwvYnV0dG9uPgogICAgICA8YnV0dG9uIHR5cGU9ImJ1dHRvbiIgZGF0YS1wcmVzZXQ9IjkwIj5MYXN0IDkwIGRheXM8L2J1dHRvbj4KICAgICAgPGJ1dHRvbiB0eXBlPSJidXR0b24iIGRhdGEtcHJlc2V0PSJhbGwiPkFsbCB0aW1lPC9idXR0b24+CiAgICA8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImZpbHRlci1maWVsZCI+CiAgICAgIDxsYWJlbCBmb3I9ImZpbHRlclBsYXRmb3JtIj5QbGF0Zm9ybTwvbGFiZWw+CiAgICAgIDxzZWxlY3QgaWQ9ImZpbHRlclBsYXRmb3JtIj48b3B0aW9uIHZhbHVlPSJhbGwiPkFsbCBwbGF0Zm9ybXM8L29wdGlvbj48L3NlbGVjdD4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0iZmlsdGVyLWZpZWxkIj4KICAgICAgPGxhYmVsIGZvcj0iZmlsdGVyQ2FtcGFpZ24iPkNhbXBhaWduPC9sYWJlbD4KICAgICAgPHNlbGVjdCBpZD0iZmlsdGVyQ2FtcGFpZ24iPjxvcHRpb24gdmFsdWU9ImFsbCI+QWxsIGNhbXBhaWduczwvb3B0aW9uPjwvc2VsZWN0PgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJmaWx0ZXItZmllbGQiPgogICAgICA8bGFiZWwgZm9yPSJmaWx0ZXJDb250ZW50VHlwZSI+Q29udGVudCB0eXBlPC9sYWJlbD4KICAgICAgPHNlbGVjdCBpZD0iZmlsdGVyQ29udGVudFR5cGUiPjxvcHRpb24gdmFsdWU9ImFsbCI+QWxsIGNvbnRlbnQgdHlwZXM8L29wdGlvbj48L3NlbGVjdD4KICAgIDwvZGl2PgogIDwvc2VjdGlvbj4KCiAgPG1haW4gY2xhc3M9InZpZXctYXJlYSI+CiAgICA8c2VjdGlvbiBpZD0idmlldy1kYXNoYm9hcmQiIGNsYXNzPSJ2aWV3IGlzLWFjdGl2ZSI+PC9zZWN0aW9uPgogICAgPHNlY3Rpb24gaWQ9InZpZXctcmVjb3JkcyIgY2xhc3M9InZpZXciPjwvc2VjdGlvbj4KICAgIDxzZWN0aW9uIGlkPSJ2aWV3LWZvbGxvd2VycyIgY2xhc3M9InZpZXciPjwvc2VjdGlvbj4KICAgIDxzZWN0aW9uIGlkPSJ2aWV3LWNvbXBhcmlzb24iIGNsYXNzPSJ2aWV3Ij48L3NlY3Rpb24+CiAgICA8c2VjdGlvbiBpZD0idmlldy1yZXBvcnQiIGNsYXNzPSJ2aWV3Ij48L3NlY3Rpb24+CiAgICA8c2VjdGlvbiBpZD0idmlldy11cGxvYWQiIGNsYXNzPSJ2aWV3Ij48L3NlY3Rpb24+CiAgICA8c2VjdGlvbiBpZD0idmlldy1oaXN0b3J5IiBjbGFzcz0idmlldyI+PC9zZWN0aW9uPgogIDwvbWFpbj4KICA8L2Rpdj4KPC9kaXY+Cgo8ZGl2IGlkPSJ0b2FzdFJvb3QiIGNsYXNzPSJ0b2FzdC1yb290IiBhcmlhLWxpdmU9InBvbGl0ZSI+PC9kaXY+Cgo8c2NyaXB0PgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgQnJhbmQgbG9nbyDigJQgZW1iZWRkZWQgb25jZSBoZXJlIGFuZCB3aXJlZCBvbnRvIGV2ZXJ5IC5icmFuZC1sb2dvCiAgIDxpbWc+IGFuZCB0aGUgZmF2aWNvbiA8bGluaz4gYXQgYm9vdHN0cmFwLCBzbyB0aGUgYmFzZTY0IHBheWxvYWQKICAgYXBwZWFycyBleGFjdGx5IG9uY2UgaW4gdGhpcyBmaWxlIGluc3RlYWQgb2Ygb25jZSBwZXIgdXNhZ2Ugc2l0ZS4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmNvbnN0IExPR09fREFUQV9VUkkgPSAnZGF0YTppbWFnZS9wbmc7YmFzZTY0LGlWQk9SdzBLR2dvQUFBQU5TVWhFVWdBQUNGb0FBQWR6Q0FZQUFBQm5iOG8zQUFBQUNYQklXWE1BQUM0akFBQXVJd0Y0cFQ5MkFBQWdBRWxFUVZSNG5PemRUVzdiV0xvRzRPUGd6dU5hZ1Ywck1Hc0ZVVTA4VFM0SUdKcEZPNGgzUUhFSHpncmFOUk1JR0pWTU9ibk9Db3BaUVRzN1NGYWdpeE14c1ZYNW8yMzlISkxQQTNqWXlORjNKSFdqK2VwN0Q1YkxaUUFBQUFBQUFBQUE0TmVlbUJFQUFBQUFBQUFBUURlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFEMFVGN1ZXVjdWRis0T0FBQjI2My9NR3dBQUFBQ2dseTVkR3dBQTdKNmdCUUFBQUFCQXorUlZQUThobkxnM0FBRFlQZFVoQUFBQUFBQTlrbGYxY1FqaC9NdUo4NnFldUQ4QUFOZ2RRUXNBQUFBQWdINkpsU0ZQNzV3NGMzOEFBTEE3Z2hZQUFBQUFBRDJSVjNYY1pQSHNYNmUxMFFJQUFIWkkwQUlBQUFBQW9BZnlxajRNSWN5L2MxSWJMUUFBWUljRUxRQUFBQUFBK3VIZmxTRmZITFVoREFBQVlBY0VMUUFBQUFBQUVwZFg5WXNRd3ZPZm5GSjlDQUFBN0lpZ0JRQUFBQUJBd3RwdEZaZS9PS0g2RUFBQTJCRkJDd0FBQUFDQXRNMS9VQmx5bDQwV0FBQ3dJNElXQUFBQUFBQ0p5cXM2QmloZWRUaWRqUllBQUxBamdoWUFBQUFBQUFucVdCbnl4ZE84cW9VdEFBQmdCd1F0QUFBQUFBRFNkQjVDT0xySHlRUXRBQUJnQndRdEFBQUFBQUFTMDI2bktPNTVxb2w3QkFDQTdSTzBBQUFBQUFCSVQ5ZktrTHRzdEFBQWdCMFF0QUFBQUFBQVNFaGUxZk1Rd3NrRFR2U1Evd3dBQUhCUGdoWUFBQUFBQUluSXEvbzRoSEQrME5Qa1ZhMCtCQUFBdGt6UUFnQUFBQUFnSGJFeTVPa2pUaU5vQVFBQVd5Wm9BUUFBQUFDUWdMeXE0eWFMWjQ4OFNlWXVBUUJndXdRdEFBQUFBQUQyTEsvcXd4RENmQU9uRUxRQUFJQXRFN1FBQUFBQUFOaS94MWFHZkhHVVYvV3grd1FBZ08wUnRBQUFBQUFBMktPOHFsK0VFSjV2OEFTMldnQUF3QllKV2dBQUFBQUE3RWxiR1hLNTRYOWQwQUlBQUxaSTBBSUFBQUFBWUgvbUc2b011V3ZpUGdFQVlIc0VMUUFBQUFBQTlpQ3Y2aGlJZUxXRmYvbVord1FBZ08wUnRBQUFBQUFBMkxFdFZZWjhsVmUxK2hBQUFOZ1NRUXNBQUFBQWdOMDdEeUVjYmZGZlZSOENBQUJiSW1nQkFBQUFBTEJEN2JhSllzdi9vbzBXQUFDd0pZSVdBQUFBQUFDN3RiWEtrRHNFTFFBQVlFc0VMUUFBQUFDQXJja1c1U3hibEI3NnQvS3Fub2NRVG5id1Q1M2tWWDI0ZzM4SEFBQkc1MzljT1FBQUFBQ3dhVEZnRVVLSW9ZS2JabHBNRFBoenlPSjRCNVVoZDhXQXkvVU8vejBBQUJnRlFRc0FBQUFBWUdQdUJDeU9RZ2lmUWdnejAvMXFGNVVoZDAwRUxRQUFZUE1FTFFBQUFBQ0FSOHNXNVlzUXdrVWJzUGhpM2t5TEc5UDl2TTNpUElUd2JNZi9yTW9XQUFEWUFrRUxBQUFBQU9EQnNrVTVhVGRZL0R0RThLNlpGaGNtKzdVeVpMNkhmMXBsQ3dBQWJNRVRRd1VBQUFBQTdpc0dMTEpGR1dzcC91OEhteHJPRGZXckdEaDV1b2QvOTJrYjhnQUFBRGJJUmdzQUFBQUFvTE9mYkxDNHEyeW1SV09xbjdkWnhFcVY1M3M4UXF3UFVkOENBQUFiSkdnQkFBQUFBUHhTeDRCRjlMNlpGdnVveVVoT1h0V0hJWVRMUFo4cjN0dWJCTWNEQUFDOUpXZ0JBQUFBQVB4UXRpaVAyNERGeTQ1VG1wbm1WL3VxRExrcjIvTy9Ed0FBZ3lOb0FRQUFBQUI4NHdFQmkraTF5cENWdktvbjk1emR0dnhxQXdrQUFIQlBCOHZsMHN3QUFBQUFnTThlR0xDSVBzVC9lRE10UG81OWttMWxTQXljSENWd25PalBxN1BUNndUT0FRQUFnMkNqQlFBQUFBRHdtSURGRnpNaGk2L21DWVVzUWxzZkltZ0JBQUFiSW1nQkFBQUFBQ09XTGNxNGZlRThoRkE4WWdxeE1zU0QvTlUyaXhocWVKWEFVZTdLMGprS0FBRDBuNkFGQUFBQUFJelFuWUJGL0h2NmlBbDhhRGM0c0hLWjRCd21DWndCQUFBRzQyQzVYTHBOQUFBQUFCaUpEUVlzdnZqZlpscTg4Zjc1dk0xaS9zak5JTnYwMjlYWnFXb1hBQURZQUJzdEFBQUFBR0FFdGhDd2lONEtXYXprVlgyY2NNZ2l0UFVoNmwwQUFHQURuaGdpQUFBQUFBeGJ0aWhuSVlTYk5naXdxWkRGcHhEQ3pGdm5xeFFyUSs1U0h3SUFBQnRpb3dVQUFBQUFERlFic0loMUZrZGJlSVd6Wmxxb29saHRzNGhiUXA0bGNKU2ZFYlFBQUlBTkViUUFBQUFBZ0lIWmNzQWllcWN5WktXdERKbW5jSlpmeUpJK0hRQUE5SWpxRUFBQUFBQVlpQml3eUJabHJBajV6eFpERmlwRDFsMXNzSTVsbTU2Mm9SQUFBT0NSYkxRQUFBQUFnSjdMRm1Xc2hiamNZcmppcm5rekxXNjhaejV2czNnUlFuaWV3Rkc2K3ZJK0FRQUFIa0hRQWdBQUFBQjZxZzFZeE5xS1p6dDZCYkV5NU1MNzVYUEk0ckNIb1FYMUlRQUFzQUdDRmdBQUFBRFFNM3NJV0FTVklkL29TMlhJWFpOMGpnSUFBUDBsYUFFQUFBQUFQYkduZ01VWEZ5cERWdktxanZmd01vV3ozTk5KcjA0TEFBQ0pPbGd1bCs0R0FBQUFBQktXTGNyanRxWmlId0dMNkgwekxkUk8zRmFHTkNHRW93U084eEIvWHAyZFh2ZnYyQUFBa0E0YkxRQUFBQUFnVVczQVlwN0E5Z1NWSWJmbVBRNVpSREV3STJnQkFBQ1BJR2dCQUFBQUFJbEpLR0FSbGMyMGFCSTR4OTYxbFNHdmV2NHk0bXU0U09BY0FBRFFXNElXQUFBQUFKQ0l4QUlXMFlkbVdzd1RPRWNxaGhCUVVBRURBQUNQSkdnQkFBQUFBSHVXTGNyRDlpRitLZ0dMTDFTR3RQS3Fqb0dUa3lRTzh6aEhlVlVmWHAyZGZ1enppd0FBZ0gwU3RBQUFBQUNBUFdrREZ1ZnQzOVBFN3VGMU15MnVFempIM3VWVkhiZEFGQU42U2JFKzVFMEM1d0FBZ0Y0U3RBQUFBQUNBSFVzOFlCRjlhQ3RNV0JsQ1pjaGRtYUFGQUFBOG5LQUZBQUFBQU94SUR3SVdYOHlhYWFGYVlyWE5JdDdWc3dTT3NrbVQ0YndVQUFEWVBVRUxBQUFBQU5peUhnVXNvcjlVaHF6a1ZYMDgwTTBlUXd1T0FBREFUajB4YmdBQUFBRFlubXhSemtJSVRRaWg2RUhJNGxNYkJtSGxzZ2QzOWlCNVZXYzlQRFlBQUNUQlJnc0FBQUFBMklJMllCRzNJUnoxYUw0cVExcDVWYjhZK09hSHJBMEFBUUFBOXlSb0FRQUFBQUFiMU5PQVJmUzJtUlp2RWpqSDN1VlZmZGh1c3hpeXlRaGVJd0FBYklXZ0JRQUFBQUJzUUk4REZxR3RESmtsY0k1VURMWXk1QTdWSVFBQThFQ0NGZ0FBQUFEd0NObWluTFFCaXo3WFRNeFZocXprVlIzdjgza0taOW15azBHL09nQUEyS0tENVhKcHZnQUFBQUJ3VHdNSldFVHZtbWt4U2VBY2U5ZFdoalE5M1VyeUVIOWVuWjFlOSsvWUFBQ3dYelphQUFBQUFNQTlEQ2hnRVZTR2ZLT3YxUzhQRmQvTGdoWUFBSEJQZ2hZQUFBQUEwTUhBQWhaZnhNcVFtelNPc2w5dFpjaXJrYjNzTElFekFBQkE3d2hhQUFBQUFNQlBaSXZ5T0lSd0VVSjRQckE1eGNxUWl3VE9rWW94emtMUUFnQUFIdUNKb1FFQUFBREF0MkxBSWx1VWx5R0UvdzR3WkJHZEozQ0dKT1JWSFRlVm5JendwUi9sVlgyY3dEa0FBS0JYYkxRQUFBQUFnRHZhRFJieHdmdkxBYytsYktaRms4QTU5aTZ2NnJqVm9SanhDT0xyVng4REFBRDNJR2dCQUFBQUFPTUpXRVR2bTJreFQrQWNxUmg3ZmNva2hQQW1nWE1BQUVCdkNGb0FBQUFBTUdyWm9qeHNBeGF2UmpLSFdRSm5TRUplMWJFKzVkbkl4NUFsY0FZQUFPZ1ZRUXNBQUFBQVJxa05XSnkzZjA5SE1vUFhLa05XOHFyK3NzRms3TVllTkFFQWdIc1R0QUFBQUFCZ1ZFWWFzSWcrQ0Jhc3VSelovZjlRWHRYWjFkbXBBQTRBQUhRa2FBRUFBQURBS0l3NFlQSEZySmtXSDlNNHluN2xWZjNDSm9jMWt4Q0NvQVVBQUhRa2FBRUFBQURBNEdXTDhyemQ1akRXRFFheE11UTZnWFBzWFY3VmgrMDJDMjVsWmdFQUFOMEpXZ0FBQUFBd1dObWluTFVCaTZNUjM3TEtrSFVxUTc0bGFBRUFBUGR3c0Z3dXpRc0FBQUNBUVJHd1dQTy96YlI0azlCNTlxYXREUGw3cEMvL1YzNjdPanRWTFFNQUFCM1lhQUVBQUFEQVlBaFlmT090a01WS1d4bHlrY0paRWhXM1dxaVhBUUNBRGdRdEFBQUFBT2k5YkZHK2FCK2lDMWpjK2hSQ21LVnltQVFJNFB6Y1JOQUNBQUM2RWJRQUFBQUFvTGV5UlRscEg2QS9jNHZmbURYVFFoWEVhcHRGZkorOFN1QW9LY3ZHUGdBQUFPaEswQUlBQUFDQTNoR3crS1YzS2tQV1hDWjBsbFJOeGo0QUFBRG82b2xKQVFBQUFOQVhNV0NSTGNwWWIvQi9RaFkvcERMa2pyeXFWWVowOHpTdjZ1TStIQlFBQVBiTlJnc0FBQUFBa3BjdHlsaHJjQ0ZjMGNtOG1SWTNQVGpuMXVWVkhkODN4Y0JmNWlaTmJQOEFBSUJmRTdRQUFBQUFJRm5ab2p4dUswSmV1cVZPWW1YSVJRL091U3RDQS9lVDllbXdBQUN3TDRJV0FBQUFBQ1JId09KQlZJYmNrVmYxZVFqaEpKa0Q5WU9nQlFBQWRDQm9BUUFBQUVBeUJDd2U1VUpseUVwZTFWL2VSOXlQYWg0QUFPamdZTGxjbWhNQUFBQUFleVZnOFdqdm0ybGhHMEVycitwcm9ZRUgrL1BxN1BTNnAyY0hBSUNkc05FQ0FBQUFnTDNKRnVWaENPRzgvWHZxSmg1TVpVZ3JyK3Faa01XanhNQ09vQVVBQVB5RW9BVUFBQUFBT3lkZ3NWRmxNeTJhQWIyZUI4dXJPcjZ2TG5wNi9GVFlqQUlBQUw4Z2FBRUFBQURBemdoWWJOeUhabHJNQi9hYUh1UFMrK3JSSmowL1B3QUFiSjJnQlFBQUFBQmJKMkN4TlNwRFdubFZ2d2doUEUvaU1QMTJGRGVEWEoyZGZoejdJQUFBNEVlZW1Bd0FBQUFBMjVRdHloZ0d1QWtoRkVJV0cvVzZtUmJYQTNvOUQ2WXlaT1BVaHdBQXdFL1lhQUVBQUFEQVZyUUJpMWhyY1dUQ0cvZWhuUzByM21lYkZldERoSGdBQU9BSEJDMEFBQUFBMkNnQmk1MllOZE5DdGNOcW0wVU1CYnhLNENoRE1objdBQUFBNEdjRUxRQUFBQURZQ0FHTG5mbExaY2lheTRUT01oU3FRd0FBNENlZUdBNEFBQUFBajVFdHlrbTJLSnNRd24rRUxMYnVVd2poZk9DdnNiTzhxZ1Y3dHVOcFh0WENGZ0FBOEFNMldnQUFBQUR3SURGZzBXNndlR2FDTzZNeXBOVUdBWW9rRGpOTWNiN04ySWNBQUFEZkkyZ0JBQUFBd0wwSVdPek4yMlphdkJucGEvOGVsU0hiWmFNRkFBRDhnS0FGQUFBQUFKMElXT3hWckF5WmpmajFyOG1yT3Rhbm5DUjBwQ0dhakgwQUFBRHdJNElXQUFBQUFQeFV0aWlQMiswQkFoYjdNMWNac3BKWDlYRWIrR0c3QkZrQUFPQUhuaGdNQUFBQUFOOFRBeGJab293QmkvOEtXZXpWdTJaYVhJejQ5ZjliZkU4K1RldEl3NVJYdGEwV0FBRHdIVFphQUFBQUFMQ20zV0FSTndhOE5KbTlVeGx5UjFzWkl2U3pPMWtJNFhvc0x4WUFBTG9TdEFBQUFBRGdNd0dMSk1YS2tKdXhEeUdzUWhhSEtrTjJMbTYwc0UwRkFBRCtSZEFDQUFBQVlPU3lSZm5sQWZhcnNjOGlNU3BEMXFrTTJiMXNiQzhZQUFDNk9GZ3Vsd1lGQUFBQU1FSnR3T0s4L2ZNQU96MS9OTk9pR2ZzUXdtcWJ4WXNRd3Q4SkhHV01mcnM2Ty8wNDlpRUFBTUJkTmxvQUFBQUFqSXlBUlMrVVFoWXJiV1hJWlFwbkdhbFlIL0ptN0VNQUFJQzdCQzBBQUFBQVJrTEFvamZlTjlOaVB2WWgzREgzZnQyclROQUNBQURXQ1ZvQUFBQUFqRUMyS09jQ0ZyMHhHL3NBdnNpck9tNVRlSlhHYVVack12WUJBQURBdndsYUFBQUFBQXhZdGlobjdVYUFJL2ZjQzY5VmhxeW9ERW5HczdFUEFBQUEvdTJKaVFBQUFBQU1Ud3hZWkl2eUpvVHdIeUdMM3ZqUWhtSllPZmZlVFVOZTFkbllad0FBQUhmWmFBRUFBQUF3SURaWTlOcXNtUllmeHo2RWNQdGd2MGpnS0t6RSs3QnBCUUFBV29JV0FBQUFBQU1nWU5GN3NUTGtldXhEdUVObFNGb203Z1FBQUc0SldnQUFBQUQwV0xZb0oyM0E0cGw3N0MyVklYZmtWUjFuY1pMTWdRanRSZ3NBQUtBbGFBRUFBQURRUXdJV2czS3VNbVFscityak9JOFV6c0lhd1JjQUFMamppV0VBQUFBQTlFY01XR1NMTWxaTS9KK1F4U0M4YmFiRm03RVA0WTVZVC9FMG1kUHdWVjdWRTlNQUFJQVZHeTBBQUFBQWVzQUdpMEg2RkVLWWpYMElYK1JWZmU3OW5iVDRIWFE5OWlFQUFFQVF0QUFBQUFCSVc3WW9qOXVBeFV0WE5UZ3psU0VyZVZVZnR1OXowcFc1R3dBQVdCRzBBQUFBQUVpUWdNWGd2Vk1ac2tabFNQcFVod0FBUU91SlFRQUFBQUNrSXdZc3NrVVpIenIvVjhoaXNGU0czSkZYOVlzUXd2TmtEc1NQUE0ycit0aDBBQURBUmdzQUFBQ0FKTmhnTVNyelpscmNqSDBJNGJZeTVES0JvOUJOckEveDNnVUFZUFFFTFFBQUFBRDJLRnVVOFVIemVRaWhjQStqRUN0RExzWStoRHZtS2tONkpkYUhxTHdCQUdEMEJDMEFBQUFBOXVCT3dPTGNnK2JSVUJseVIxN1Y4YUg5cTJRT1JCZVpLUUVBZ0tBRkFBQUF3RTRKV0l6YWhjcVFGWlVodmZWczdBTUFBSUFnYUFFQUFBQ3dHd0lXby9lK21SYnpzUS9oanZnNU9Fcm1OSFNXVjNWMmRYYmFtQmdBQUdQMnhPMERBQUFBYkZlMktHTmRSTnhrVUFoWmpKYktrRlo4VU45K0Z1aW5pWHNEQUdEc2JMUUFBQUFBMkpJMllESDN5LzNSSzV0cFlRUEFMWlVoL1phTmZRQUFBQ0JvQVFBQUFMQmhBaGJjOFVGbHlLMjhxdU1zVGxJNUR3OGlhQUVBd09nZExKZkxzYzhBQUFBQVlDTUVMUGlPUDV0cGNXMHduME1XeHlHRVJuM09JUHgyZFhiNmNleERBQUJndkd5MEFBQUFBSGlrYkZGTzJqb0VBUXZ1ZWkxa3NlWlN5R0l3NGxZTDcyMEFBRVpMMEFJQUFBRGdnZHFBUmR4Zzhjd00rWmNQN1h1RDFUYUxjNStUUVprSVdnQUFNR2FDRmdBQUFBRDNKR0JCQjdObVdxaFd1SzBNRVRvWmxzbllCd0FBd0xnSldnQUFBQUIwSkdCQlIzK3BERmx6b1RKa2NMS3hEd0FBZ0hFN1dDNlhZNThCQUFBQXdFOWxpekpySHhZTFdQQXJuMElJeDdaWnJPUlYvU0tFOEhjS1oySGpmcjg2TzcweFZnQUF4c2hHQ3dBQUFJQWZ5QmJsbDhxRGwyWkVSeXBEV25sVkg0WVFMcE00RE5zd2NiOEFBSXlWb0FVQUFBREF2d2hZOEVCdm0ybnh4dkMrVWhreWJPcERBQUFZTFVFTEFBQUFnSmFBQlk4UUswTm1CcmlTVi9YRTUyandCQzBBQUJndFFRc0FBQUJnOUxKRmVkaisrdDZEWVI1cXJqSmtSV1hJYUR3Yit3QUFBQml2Sis0ZUFBQUFHS3NZc01nV1pkeGdjU05rd1NPOGE2YkZoUUYrRlQ5VFI0bWNoUzFxTjVjQUFNRG8yR2dCQUFBQWpFNjd3ZUs4L1h2cUhjQWpxQXk1STYvcVdDZnhLcGtEc1czeHZxOU5HUUNBc1JHMEFBQUFBRVpEd0lJdGlKVWhOd2I3bGNxUWNjbkdQZ0FBQU1aSjBBSUFBQUFZUEFFTHRrUmx5QjE1VmNmS2tKTmtEc1F1cUE0QkFHQ1VucmgyQUFBQVlNaXlSUmxySFpvUVFpRmt3WWFkRytoS1h0WEg3V2VNY1RuS3EvclFuUU1BTURZMldnQUFBQUNEMUFZczRpL3NqOXd3VzFBMjA2SXgySzlVaG94WDNHcnhadXhEQUFCZ1hBUXRBQUFBZ0VFUnNHQUgzamZUWW03UUszbFZ4ODBlejFJNEMzdVJDVm9BQURBMmdoWUFBQURBSUFoWXNFTXp3MTVwSzBPRVRzWnRNdllCQUFBd1BvSVdBQUFBUUs5bGkzTFNQdWoxaTNwMjRiWEtrRFVYSVlTbkNaMkgzY3ZNSEFDQXNUbFlMcGN1SFFBQUFPZ2RBUXYyNEVOODZ6WFQ0cVBoZjk1bThTS0U4SGNDUjJILy9yZzZPeFZBQWdCZ05HeTBBQUFBQUhwRndJSTltZ2xack9SVmZSaEN1RXpoTENRaGJyVVF0QUFBWURRRUxRQUFBSUJlRUxCZ3oySmx5TFZMK0VwbENIZXBEd0VBWUZRRUxRQUFBSUNrWll2eXVIMm8rOXhOc1NjZjJwQVBxMjBXTWZUMDBpeTRZMklZQUFDTWlhQUZBQUFBa0tRMllESDNRSmNFbktzTVdWRVp3ZytjR0F3QUFHUHl4RzBEQUFBQUtZa0JpMnhSeGdlNS94V3lJQUZ2bTJueHhrVjhGY05QUjRtY2hZUzBtMDRBQUdBVWJMUUFBQUFBa21DREJRbjZGRUtZdVppVnZLcXpFTUtyRk01Q2t1TDc0OXJWQUFBd0JvSVdBQUFBd0Y1bGkvS3dEVmg0Z0V0cVppcEQxcWdNNFdmaVJvc0xFd0lBWUF3RUxRQUFBSUM5YUFNVzUrM2ZVN2RBWXQ2cERMbVZWM1VNUTUya2NoNlNsTGtXQUFERzRtQzVYTHBzQUFBQVlHY0VMT2lCV0JtU05kUGl4bVY5clF6NUo0R2prTDdmcjg1T2ZXNEFBQmc4R3kwQUFBQ0FuUkN3b0VmbVFoWnIxRUhRVlF6bCtPd0FBREI0VDF3eEFBQUFzRzNab2p4dkg3NFZRaFlrTGxhR0NCYTA4cXFPbjkxblNSeUdQbEFmQWdEQUtOaG9BUUFBQUd4TnRpaG5jVHRBQ09ISWxPbUJXQmt5YzFFcmVWVWZ0NTlmNkdwaVVnQUFqSUdnQlFBQUFMQnhBaGIwMUlYS2tEV1hOdEJ3VDdhZkFBQXdDZ2ZMNWRKTkF3QUFBQnNoWUVHUHZXK21oZHFEVmw3VkwwSUlmeWR4R1BybWo2dXowOGF0QVFBd1pEWmFBQUFBQUkrV0xjcjRVUFpDd0lJZVV4blN5cXY2c04xbUFROFJBMHVDRmdBQURKcWdCUUFBQVBCZzJhS2N0QnNzckl1bno4cG1Xbmd3ZkV0bENJOHhFZFFCQUdEb0JDMEFBQUNBZXhPd1lFQStOTk5pN2tKWDhxcU9uKzNuS1p5RjNsTEJBd0RBNEIwc2wwdTNEQUFBQUhRaVlNRUEvZGxNaTJzWCs3VXlwRkVCeEFiOGRuVjIrdEVnQVFBWUtoc3RBQUFBZ0YvS0ZtWDhoZktGZ0FVRDgxcklZczFjeUlJTmlmK2Q0Yk1GQU1CZ0NWb0FBQUFBUDVRdHl1UDI0ZXRMVTJKZ1ByVHZiVzRyUTE2WkJSc3lFYlFBQUdESUJDMEFBQUNBYndoWU1BS3pabHFvTnJoMWtjcEJHSVRNTlFJQU1HU0NGZ0FBQU1CWEFoYU14RjhxUTI3bFZSMC84eWVwbklkQm1MaEdBQUNHN0dDNVhMcGdBQUFBR0RrQkMwYmtVd2poMkRhTGxieXE0K2FCZjFJNEM0UHorOVhaNlkxckJRQmdpR3kwQUFBQWdCSExGdVZoQ09HOC9YdnF2Y0FJcUF4WnB6S0ViWWtoSGtFTEFBQUdTZEFDQUFBQVJrakFncEY2MjB5TE55NS9KYS9xK1BsL2xzSlpHS1JZSCtMekJnREFJQWxhQUFBQXdJZ0lXREJpc1RKazVnMndrbGYxbDdvZzJKYk1aQUVBR0NwQkN3QUFBQmdCQVFzSWM1VWhheTU5RjdCbHRxVUFBREJZQjh2bDB1MENBQURBZ0dXTE12NksvOEpEVlVic1hUTXRKdDRBSzNsVnZ3Z2gvSjNDV1JpOFA2N09UaHZYREFEQTBOaG9BUUFBQUFQVkJpeGlOY0NSTzJiRVZJYmNrVmYxWWJ2TkFuWWhCcHdFTFFBQUdCeEJDd0FBQUJnWUFRdFlFeXREYm96a0s1VWg3RkptMmdBQURKR2dCUUFBQUF5RWdBVjhJMWFHWEJqTFNsN1ZjYnZBOHhUT3dtaW83QUVBWUpBT2xzdWxtd1VBQUlBZXl4WmxmSkFWSHlhZnVFZFk4MGN6TGRRVzNGYUdOSUpZN01GdlYyZW5IdzBlQUlBaHNkRUNBQUFBZXFvTldNUU5Gcy9jSVh5akZMSllZOXNOK3hMclE2NU5Id0NBSVJHMEFBQUFnSjRSc0lCZmV0OU1pN2t4cmJTVklhOVNPQXVqTkJHMEFBQmdhQVF0QUFBQW9DY0VMS0N6bVZHdHVVem9MSXpQeEowREFEQTBnaFlBQUFDUXVHeFJIcmNQU2dVczROZGVxd3k1bFZlMXloRDJMWE1EQUFBTXpjRnl1WFNwQUFBQWtLQTJZQkVma3I1MFA5REpoL2pSYWFiRlIrUDZITEtJRDdqL1NlQW84UHZWMmVuTjZLY0FBTUJnMkdnQkFBQUFpUkd3Z0FlYkNWbXNVUmxDS2liZWp3QUFESW1nQlFBQUFDUkN3QUllSlZhR1hCdmhTbDdWNXlHRWt4VE9BdXBEQUFBWUdrRUxBQUFBMkxOc1VSNjJBWXRYN2dJZTVFUDdHV0lWc2pnMkR4SWphQUVBd0tBSVdnQUFBTUNldEFHTDgvYnZxWHVBQnp0WEdiTG0wbmNLaVhubVFnQUFHSktENVhMcFFnRUFBR0NIQkN4Z285NDIwK0tGa2E3a1ZUMExJZnduaGJQQXYveDVkWGFxM2djQWdFR3cwUUlBQUFCMlJNQUNOdTVUQ0dGbXJDdDVWY2Z2bUlzVXpnTGZFZXREQkMwQUFCaUVKNjRSQUFBQXRpOWJsUE1Rd2swSW9SQ3lnSTJacVF4Wm96S0VsRTNjRGdBQVEyR2pCUUFBQUd4UnRpampyKzFqeU9MSW5HR2ozalhUNG8yUnJ1UlZIZXRUbnFkd0Z2aUJ6R0FBQUJpS2crVnk2VElCQUFCZ3d3UXNZS3RpWlVqV1RJc2JZLzVhR2RMNHZxRUhmcnM2TzdXRkJnQ0EzclBSQWdBQUFEWkl3QUoyWWk1a3NjWjNEbjBSNjBOc29nRUFvUGNFTFFBQUFHQURCQ3hnWjJKbHlJVnhyK1JWSFI5Y3YwcmhMTkJCSm1nQkFNQVFDRm9BQUFEQUkyU0xjdElHTEo2WkkyeGRyQXlaR2ZPYXk0VE9Bcjh5TVNFQUFJWkEwQUlBQUFBZVFNQUM5dUpDWmNpdHZLcHQwYUZ2TWpjR0FNQVFIQ3lYU3hjSkFBQUFIUWxZd042OGI2YUZoN1N0dktyakxQNUo0akJ3UDM5Y25aMDJaZ1lBUUovWmFBRUFBQUFkQ0ZqQTNxa01XYWN5aEw2S0lTRkJDd0FBZWszUUFnQUFBSDRpVzVUSGJjRGlwVG5CM3BUTnRQQmd0cFZYOVhrSTRTU0p3OEQ5MlV3REFFRHZDVm9BQUFEQWR3aFlRREkrTk5OaTdqcFc4cXIrOHQwRWZUVnhjd0FBOUoyZ0JRQUFBTndoWUFISlVSbXlMbGFHUEUzcFFIQlB0ckVBQU5CN0I4dmwwaTBDQUFBd2VnSVdrS1RYemJRNGR6VXJlVlhIME1sL1VqZ0xQTktmVjJlbjE0WUlBRUJmMldnQkFBREFxR1dMOGpDRUVCL2tGbU9mQlNUbWc0cU1XM2xWeCsrcWkxVE9BNDhVNjBNRUxRQUE2QzFCQ3dBQUFNTTloc3NBQUNBQVNVUkJWRWJwVHNEaTNCcCtTTktzbVJZZlhjMVhLa01Za3N4dEFnRFFaNElXQUFBQWpJcUFCZlRDWDgyMDhHdjNWbDdWTDBJSXo1TTRER3lHb0FVQUFMMTJzRnd1M1NBQUFBQ0RKMkFCdmZFcGhIQnNtOFZLV3hseTQzdUxBZnI5NnV6MHhzVUNBTkJITmxvQUFBQXdlTm1pbklVUUxqeW9oRjVRR2JKdTdydUxnY3JhRUJFQUFQU09vQVVBQUFDRDFRWXM0a1BLSTdjTXZmQzJtUlp2WE5WS1h0V1RFTUtyRk00Q1d4Q0RGajd2QUFEMGtxQUZBQUFBZ3lOZ0FiMFVLME5tcm02bHJReTVUT0Vzc0NVVGd3VUFvSzhFTFFBQUFCZ01BUXZvdGJuS2tEWG52c3NZdUdjdUdBQ0F2anBZTHBjdUR3QUFnRjdMRnVXay9lVzNoNUxRVCsrYWFlSFg3YTI4cW1PbHdqOUpIQWEyNjQrcnM5UEdqQUVBNkJzYkxRQUFBT2l0Tm1BeDk2dFk2RFdWSWQ5U0djSll4RkNSb0FVQUFMMGphQUVBQUVEdkNGakFvTVRLa0J0WHVwSlhkZnh1TzBuaExMQURFOEVpQUFENlNOQUNBQUNBM2hDd2dNR0psU0VYcm5VbHIrcmpFTUo1Q21lQkhja01HZ0NBUGhLMEFBQUFJSG5ab293UFlpNEVMR0J3aEFyV3hWLzJQMDNwUUxCbEozbFZIMTZkblg0MGFBQUEra1RRQWdBQWdHUmxpL0s0M1dEeDBpM0I0SlROdEdoYzYwcGUxZWZDWkl4VURGTmV1M3dBQVBwRTBBSUFBSURrQ0ZqQTRMMXZwc1hjTmEvRVgvUzMzM2t3UmhOQkN3QUEra2JRQWdBQWdHUUlXTUJvekZ6MUdwVWhqRm5tOWdFQTZCdEJDd0FBQVBZdVc1VHgxOXdYQWhZd0NxOVZodHpLcS9wRkNPRjVLdWVCUFpnWU9nQUFmWE93WEM1ZEdnQUFBSHZSQml6TzJ6Ky81b2JoK3hBLytzMjArT2l1djFhRzNQaitnL0Q3MWRucGpURUFBTkFYTmxvQUFBQ3djd0lXTUZveklZczFjOStCOEZuV2hvNEFBS0FYQkMwQUFBRFlHUUVMR0xWWUdYSTk5aUY4a1ZkMXJFdDRsY1pwWU8vaTUrR05hd0FBb0M4RUxRQUFBTmc2QVFzWXZRL3Q5Z1p1SzBNdXpRSyt5b3dDQUlBK0ViUUFBQUJncTdKRk9Xc2ZzQjZaTkl6V3VjcVFOZWUrRTJITk0rTUFBS0JQRHBiTHBRc0RBQUJnNHdRc2dOYmJabHE4TUl5VnZLcmpML2YvU2VFc2tKZy9yODVPMVFzQkFOQUxObG9BQUFDd1VRSVd3QjJmUWdnekExbWpNZ1MrTDRhUUJDMEFBT2dGUVFzQUFBQTJRc0FDK0k2WnlwQmJlVlhINzhpVFZNNERpY2xjQ0FBQWZTRm9BUUFBd0tOa2kzTFNCaXowcXdOM3ZXdW14UnNUV2NtcitqaUVjSjdDV1NCUkV4Y0RBRUJmQ0ZvQUFBRHdJQUlXd0Urb0RQbFdyQXg1bXRxaElDRkhlVlVmWHAyZDJvSURBRUR5QkMwQUFBQzRGd0VMb0lONU15MXVER29scitwejM1blFTYXdQdVRZcUFBQlNKMmdCQUFCQUp3SVdRRWV4TXVUQ3NGYmlML1RiNzA3ZzF5YUNGZ0FBOUlHZ0JRQUFBRCtWTGNyakVFSjhhUHJjcElCZlVCbnlMWlVoME4zRXJBQUE2QU5CQ3dBQUFMNnJEVmpFWDJHL05DR2dvd3VWSWJmeXFuNGhwQWIza2hrWEFBQjljTEJjTGwwVUFBQUFYd2xZQUEvMHZwa1dIcEsyMnNxUUc5c3M0TjUrdnpvN0ZkZ0NBQ0JwTmxvQUFBRHdtWUFGOEVncVE5WmRDRm5BZzB6YXloMEFBRWlXb0FVQUFNRElaWXZ5c0ExWXZCcjdMSUFISzV0cDBSamZTbDdWRTZFMWVEQ2JjUUFBU0o2Z0JRQUF3RWkxQVl2ejlzK3Zyb0dIK3RCTWk3bnByYlNWSVg2TkR3ODNNVHNBQUZJbmFBRUFBREF5QWhiQWhxa01XUmRESjBjcEhRaDY1c1NGQVFDUXVvUGxjdW1TQUFBQVJrREFBdGlDMTgyME9EZllsYnlxWStYQlB5bWNCWHJ1ejZ1ejAydVhDQUJBcXA2NEdRQUFnT0hMRm1WOEVIb1RRaWlFTElBTitkQnViK0NXeWhEWWpNd2NBUUJJbWVvUUFBQ0FBY3NXNWN3YWUyQkxaczIwK0dpNEszbFZ6MVVld01aTVFnZ1h4Z2tBUUtvRUxRQUFBQVpJd0FMWXNyK2FhV0d0Znl1djZ1TjJZeEN3R1RaYUFBQ1FORUVMQUFDQUFSR3dBSGJnVXdqaDNLRFhxQXlCelRyS3Evcnc2dXpVMWh3QUFKSWthQUVBQURBQTJhSjgwYTdZRnJBQXRrMWx5QjE1VmNmUXliTmtEZ1RERWV0RDNyaFBBQUJTSkdnQkFBRFFZOW1pbkxRYkxEemtBM2JoYlRNdFBQaHN0WlVoOHlRT0E4T1RDVm9BQUpBcVFRc0FBSUFlRXJBQTlpQldoc3dNZmszY0pQUTBvZlBBa0V6Y0pnQUFxUkswQUFBQTZCRUJDMkNQNWlwRGJ1VlZIU3VibnFkeUhoaWd6S1VDQUpDcWcrVnk2WElBQUFBU2x5M0tyUDNsdElBRnNBL3ZtbW5oMStXdHZLb1BRd2czdGxuQTF2MXhkWGJhR0RNQUFLbXgwUUlBQUNCaDJhSThiamRZdkhSUHdKNm9EUG1XeWhEWWpSZzBGYlFBQUNBNWdoWUFBQUFKRXJBQUVoSXJRMjVjeUVwZTFSUGZ6YkF6OGZOMmFkd0FBS1JHMEFJQUFDQWhBaFpBWW1KbHlJVkxXV2tyUXp6MGhkM0p6Qm9BZ0JRSldnQUFBQ1JBd0FKSTFMbUxXUk8vcDQ4U09nOE0zWWtiQmdBZ1JRZkw1ZExGQUFBQTdFbTJLQS9iQjVubit2NkJ4SlROdEppN2xKVzhxdU12Ni85SjRTd3dNbjllbloxZXUzUUFBRkppb3dVQUFNQWVDRmdBaVhzdlpQRU5sU0d3SDVNUWdxQUZBQUJKRWJRQUFBRFlJUUVMb0NkbUx1cFdYdFZ6RlFhd041blJBd0NRR2tFTEFBQ0FIUkN3QUhya2RUTXRHaGUya2xmMWNRaWhTT0VzTUZLQ0ZnQUFKT2VKS3dFQUFOaXViRkhHWDRiZnRBL3FoQ3lBbEgwSUlhZ01XYWN5QlBicnFBMDhBUUJBTW15MEFBQUEySkkyWUJFZldCNlpNZEFUczJaYWZIUlpLM2xWeHkxRXoxSTRDNHhjMW9aV0FRQWdDWUlXQUFBQUd5WmdBZlJVckF5NWRua3I3Uy9vYmZlQU5NU2d4UnQzQVFCQUtnUXRBQUFBTmtUQUF1Z3hsU0hmdWxUM0JNbVl1QW9BQUZJaWFBRUFBUEJJMmFLTS8rZi9SUWpoeEN5Qm5qcFhHWElycitvWEtrTWdLVDZQQUFBazVXQzVYTG9SQUFDQUIyZ0RGblAvNXovUWMyK2JhZkhDSmE3a1ZYMFlRcml4elFLUzg4ZlYyV25qV2dBQVNJR05GZ0FBQVBja1lBRU15S2NRd3N5RnJsRVpBbW1LLy90TDBBSUFnQ1FJV2dBQUFIUWtZQUVNMEV4bHlLMjhxdVAzL1BOVXpnT3N5WXdEQUlCVUNGb0FBQUQ4UXJZb2o5dGZPQXRZQUVQeXJwa1diOXpvU2xzWmNwbkNXWUR2RXJRQUFDQVpUMXdGQUFEQTk4V0FSYllvNDBPMy93cFpBQU9qTXVSYmNXUFJVV3FIQXI0NmFRTlJBQUN3ZHpaYUFBQUEvRXU3d1NJK2NIdHBOc0JBelp0cGNlTnlWOXJLa0ZjcG5BWDRxYmpWNHRxSUFBRFlOMEVMQUFDQWxvQUZNQkt4TXVUQ1phOHhEK2lIaWFBRkFBQXBFTFFBQUFCR0wxdVVoMjNBd3ErWmdhRlRHZkl2ZVZYSDcvK1RwQTRGL0VobU1nQUFwRURRQWdBQUdLMDJZSEhlL2ozMVRnQkc0RUpseUsyOHF1TkQyeUtWOHdDL05ERWlBQUJTY0xCY0xsMEVBQUF3S2dJV3dFaTliNmFGWDRQZmtWZDFyQ0I0bHN5QmdDNSt2em83RlJnREFHQ3ZiTFFBQUFCR1E4QUNHRG1WSVhma1ZYMHVaQUc5RkxkYVhMbzZBQUQyU2RBQ0FBQVloV3hSemdVc2dCRXJtMm5SZUFPczVGVjlIRUtZcDNBVzRONXM1Z0VBWU84RUxRQUFnRUhMRnVXc2ZaaDI1S2FCa2ZyUVRBdWhnbldYZ25mUVc0SVdBQURzbmFBRkFBQXdTQUlXQUYrcERMa2pyK29YS2tPZzEzeCtBUURZdTRQbGN1a1dBQUNBd1JDd0FGanp1cGtXNTBheWtsZjFZUWpoeGpZTDZMMC9yODVPcjEwakFBRDdZcU1GQUFBd0NBSVdBTi80MEg0dmNrdGxDQXhEckE4UnRBQUFZRzhFTFFBQWdGN0xGdVdrZlpCb2pUVEF1bGt6TFQ2YXlVcGUxZkcvTDU2bmNCYmcwVElqQkFCZ253UXRBQUNBWGhLd0FQaXB2NXBwNGRmZXJiWXk1REtKd3dDYk1ERkZBQUQyU2RBQ0FBRG9GUUVMZ0YvNkZFSTRONlkxcXFWZ1dJNWlnT3JxN05UV0hnQUE5a0xRQWdBQTZBVUJDNERPVkliYzBWYUd2RXJtUU1DbXhQb1FtM3NBQU5nTFFRc0FBQ0JwMmFJOGJnTVdMOTBVd0MrOWJhYkZHMk5hYzVIUVdZRE5tUWhhOFAvczNVOXUzTmJXTCt6dDRQU2xPd0xwak1BOEkxQ2x3Njc5Z1FDaG52V09JTG9qVU5VSWpqS0MxKzRWQ0JEWDZiTHpTaU80NVJGY2FRYldDUFJoSi9TSjQvaVAvbFRWNXQ1OEhzRE5BT1JhRlpGVi9IRXRBSUJVQkMwQUFJQkpFckFBZUxTNE11Uk0yZjdVZEVPOGpyeWN5dkVBVzdWUVRnQUFVaEcwQUFBQUprWEFBdURKbGxhRy9LbnBocmhXNEdJcXh3TnNYYVdrQUFDazh1TCsvbDd4QVFDQTVBUXNBSjdsZW5ONjRlM3V6elRkc0RITkFvcjNyNzZ0TjlvTUFNQyttV2dCQUFBa1ZhMVhoeUdFYzI4ZEF6eVpsU0ZmYUxyaFhNZ0NaaUZPdFJDMEFBQmc3d1F0QUFDQUpENExXTVIvQjdvQThHUnhaY2lOOHYyaDZZWlBFNUtBOGxrZkFnQkFFb0lXQUFEQVhnbFlBR3hWWEJseXFhUi84ZGIxQldiRHlpUUFBSklRdEFBQUFQWkN3QUpnSjg2VjlVOU5OOFFWS2lkVE9SNWc1NndJQWdBZ2laK1VIUUFBMkxWcXZZb1BBdU5ZK3dzaEM0Q3RXVzFPTHpiSytZZW1HMktnejNRUG1KbW1HMHkxQUFCZzcweTBBQUFBZHFaYXI4N0dQZmxIcWd5d1ZSODJweGRMSmYwTEswTmducW9Rd3BYZUF3Q3dUNElXQUFEQTFnbFlBT3pjbVJML3FlbUcxeUdFVjFNNUhtQ3ZGcWJaQUFDd2I0SVdBQURBMWdoWUFPekZyMWFHL01uS0VKaTlhdTRGQUFCZy93UXRBQUNBWjZ2V3E4VTRzbDNBQW1DM2JzZEFHMzhTOElONU80cUJxNzZ0UDg2OUVBQUE3SStnQlFBQThHUmp3Q0krNERwUlJZQzlPTnVjWG5pWU9HcTZJVjZIZnBuRXdRQXB4YjhGNzNVQUFJQjlFYlFBQUFBZVRjQUNJSW00TXVSSzZmL2k3WVNPQlVpbkVyUUFBR0NmQkMwQUFJQUhFN0FBU01iS2tDODAzV0JsQ1BESlFpVUFBTmduUVFzQUFPQ0hxdlVxdmlWNEtXQUJrTXk1bFNGL2Fyb2hYcGN1cG5JOFFITHVVUUVBMkt1ZmxCc0FBUGlXYXIwNnJ0YXJPSmI5Ly9vQkd5Q1ozemFuRjBiaS81V1ZJY0JmakFFc0FBRFlDeE10QUFDQXY0a0JpM0ZFL1J2VkFVanFMb1J3cGdWL2FycmhQSVR3Y2lySEEweEdERnBzdEFNQWdIMFF0QUFBQVA1RHdBSmdjczZzRFBsVDB3MmZybE1BWDFxWWRnTUF3TDRJV2dBQUFERmdjUmhDdUJTd0FKaVVheXREL2lZK1JEMlkyREVCMDJCMUNBQUFleU5vQVFBQU16WUdMTTdIZng1Y0FVeUhsU0ZmYUxvaDF1TmtVZ2NGVEltVlFnQUE3TTJMKy90NzFRWUFnSmtSc0FDWXZQKzlPYjI0MUtZL05OMFFyMXMzcmxuQUQvemN0L1dWSWdFQXNHc21XZ0FBd0l3SVdBQms0VnJJNG0rc0RBRWVZaEZDRUxRQUFHRG5CQzBBQUdBR0JDd0FzbUZseUJlYWJuZ2RRbmcxcVlNQ3BxclNHUUFBOWtIUUFnQUFDbGV0Vi9HQjNUS0VjS1RYQUpOM3VUbTl1TkdtUDR3clEwejNBQjVLMEFJQWdMMTRjWDkvcjlJQUFGQWdBUXVBN0h6WW5GNTRTUGlacGh0aXlPS1h5UndRa0lOLzltMHRzQVlBd0U2WmFBRUFBSVVSc0FESWxwVWhuMm02WVNGa0FUeEJES3dKV2dBQXNGT0NGZ0FBVUFnQkM0Q3NyVGFuRnhzdC9NTzRNdVR0Rkk0RnlFNE1hYjNYTmdBQWRrblFBZ0FBTWxldFY0dHhmLzFMdlFUSTB1M205R0twZFg5eExqZ0lQSkVWVEFBQTdKeWdCUUFBWkdvTVdNUUhjeWQ2Q0pBMUswTSswM1JEZkVoNk1aa0RBbkxqM2hnQWdKMzdTWWtCQUNBdk1XQlJyVmRYSVlULzhVTXlRUForM1p4ZVhHbmpYMWdaQWp6TEdOZ0NBSUNkTWRFQ0FBQXlZWUlGUUhGdXg3L3JqSnB1V0ZxRkJXeEJ2Ry9lS0NRQUFMc2lhQUVBQUJOWHJWZkhJWVRMRU1JcnZRSW95dG5tOU9Lamx2Nmg2WVo0dlR1ZndyRUEyVFBSQWdDQW5SSzBBQUNBaVJvREZ2SE4zamQ2QkZDY2QxYUcvRTFjR1hJd3NXTUM4aVJvQVFEQVRyMjR2NzlYWVFBQW1CQUJDNERpM1lVUWprMnorRlBURFhHU3hiK25janhBRWY1WDM5Yit6Z0lBc0JNbVdnQUF3RVFJV0FETXdxMlZJWC9WZE1QaGVQMEQyS1k0MWNMa0lBQUFka0xRQWdBQUVxdldxMDhQbUg3UkM0Qml4WURGY25ONjhWYUwvOGJLRUdBWEZvSVdBQURzaXFBRkFBQWtNZ1lzenNkL0hqQUJsRW5BNGp1YWJuZ2RRbmcxMlFNRWNsYnBIZ0FBdXlKb0FRQUFleVpnQVRBTEFoWS9NSzRNVVI5Z1Z4WXFDd0RBcmdoYUFBREFuZ2hZQU15Q2dNWERMVjBQZ1IwNmFMcmh1Ry9yRzBVR0FHRGJCQzBBQUdBUHF2WHEzQU1sZ0tJSldEeEMwdzN4VGZOZnNqbGdJRmNMazNNQUFOZ0ZRUXNBQU5paGFyMDZHd01XUitvTVVDUUJpMGV5TWdUWW8wcXhBUURZQlVFTEFBRFlBUUVMZ09JSldEemR1ZXNqc0NlQ0ZnQUE3SVNnQlFBQWJKR0FCVUR4QkN5ZW9lbUcrTkR6SXRzVEFISnpvbU1BQU95Q29BVUFBR3hCdFY2OURpRmNDbGdBRkV2QVlqdlVEOWlycGhzV2ZWdGZxVG9BQU5za2FBRUFBTTlRclZlTGNZS0Z0K1VBeWlSZ3NTVk5OOFNWSVMrTE9Ca2dKM0dTanFBRkFBQmJKV2dCQUFCUElHQUJVRHdCaXkxcXV1Rnd2RzRDN0Z1bDRnQUFiSnVnQlFBQVBJS0FCVUR4QkN4Mkk2N1hPaWp4eElESlcyZ1JBQURiOXVMKy9sNVJBUURnQndRc0FJcjNJWVlCQkN5MnIrbUcrRGI1L3kzdHZJQ3MvSysrclQ5cUdRQUEyMktpQlFBQWZFZTFYaDJQQVlzMzZnUlFwT3R4Z29YOS9idHpXZXFKQWRtSW9lbjMyZ1VBd0xZSVdnQUF3RmNJV0FBVVQ4QmlENXB1ZUcwYUZEQUJsYUFGQUFEYkpHZ0JBQUNmRWJBQUtKNkF4WDZaWmdGTXdVSVhBQURZSmtFTEFBQVFzQUNZQXdHTFBXdTZJVjVYajJaMTBzQlVWVG9EQU1BMnZiaS92MWRRQUFCbXExcXZEa01JNStPL0E1OEVnT0lJV0NUUWRFTU1NRzVjVzRFSitWZmYxaHNOQVFCZ0cweTBBQUJnbGdRc0FJb25ZSkhXMHZVVm1KaHFESUFCQU1DekNWb0FBREFyQWhZQXhST3dTS3pwaG9WVlhNQUVXUjhDQU1EV0NGb0FBREFMQWhZQXhST3dtSTdsM0FzQVROSkNXd0FBMkJaQkN3QUFpbGV0VjJjaGhFc0JDNEFpQ1ZoTVNOTU44WnA3TXZjNkFKUDBVbHNBQU5pV0YvZjM5NG9KQUVDUnhvQkZmS3YyU0ljQmlpTmdNVEZOTjhUcFVSdlhYV0RDZnU3YjJuVURBSUJuTTlFQ0FJRGlDRmdBRkUzQVlyck9YWHVCaWF0Q0NLNGZBQUE4bTZBRkFBREZFTEFBS0pxQXhZUTEzWEFjUXJpWWV4MkF5VnVNS3dVQkFPQlpCQzBBQU1oZXRWNTkrc0hVM21XQThnaFk1TUdEU3lBSGxTNEJBTEFOTCs3djd4VVNBSUFzalFHTE9NSGlSQWNCaWlOZ2tZbW1HK0wxK0gvbVhnY2dHLy9zMi9wR3V3QUFlQTRUTFFBQXlJNkFCVURSQkN6eVk1b0ZrSk00MVVMUUFnQ0FaeEcwQUFBZ0d3SVdBRVVUc01oUTB3MW5WbmNCbVlsQmkvZWFCZ0RBY3doYUFBQXdlZFY2ZFJ4Q2VDdGdBVkFrQVl0TU5kMXdhSm9Ga0tHRnBnRUE4RnlDRmdBQVROWVlzSWdUTE43b0VrQnhCQ3p5RjYvUkIzTXZBcEFkNFcwQUFKN3R4ZjM5dlNvQ0FEQXBBaFlBUlJPd0tFRFREZkZhL2YvbVhnY2dXLy9xMjNxamZRQUFQSldKRmdBQVRJYUFCVURSQkN6SzhuYnVCUUN5Vm9VUUJDMEFBSGd5UVFzQUFKS3IxcXRQTzk0RkxBREtJMkJSbUtZYkZrYnZBNWxiQ0l3QkFQQWNnaFlBQUNRekJpek94MzkydkFPVVJjQ2lYQjVPQXJtcmRCQUFnT2NRdEFBQVlPOEVMQUNLSm1CUnNLWWI0b3F2bzduWEFjamVTeTBFQU9BNVh0emYzeXNnQUFCN0lXQUJVRFFCaThJMTNSQ3Y0emV1NFVBaGZ1N2IyalVMQUlBbk1kRUNBSUM5cU5hcnBZQUZRSkVFTE9iajBuVWNLTWdpaE9EYUJRREFrd2hhQUFDd1U5VjZkUllmd0JrekRsQWNBWXNaYWJxaENpRzhtWHNkZ0tKVTJna0F3Rk1KV2dBQXNCTUNGZ0RGRXJDWXA4dTVGd0FvemtKTEFRQjRxaGYzOS9lS0J3REExZ2hZQUJSTHdHS21tbTU0SFVMNFAzT3ZBMUNrZi9adGZhTzFBQUE4bG9rV0FBQnNoWUFGUUxGK2k5TU1CQ3ptcWVtR1E5TXNnSUxGOVNHQ0ZnQUFQSnFnQlFBQXoxS3RWNHN4WUhHaWtnQkZlVGRPc1BBQWF0N09oU2lCZ3NYdk11ODFHQUNBeHhLMEFBRGdTUVFzQUlvbFlNSHZtbTQ0SG9NV0FLV3FkQllBZ0tjUXRBQUE0RkVFTEFDS0pXREJsK0wxL2tCVmdJTDVUZ01Bd0pPOHVMKy9WemtBQUg1SXdBS2dXQUlXL0UzVERmRzYvejhxQTh6QXYvcTIzbWcwQUFDUFlhSUZBQURmVmExWHgyUEE0bzFLQVJSRndJTHZXYW9PTUJNeFdDWm9BUURBb3doYUFBRHdWUUlXQU1VU3NPQzdtbTQ0TThFS21KRktzd0VBZUN4QkN3QUEva0xBQXFCWUFoYjhVTk1OaHlHRVM1VUNaa1RRQWdDQVJ4TzBBQURnZHdJV0FNVVNzT0F4emtNSUJ5b0d6TWpMR0RMcjIvcWpwZ01BOEZDQ0ZnQUFNMWV0VjRmalE1V0x1ZGNDb0RBQ0ZqeEswdzNIN2dlQW1ZcFRMYTQwSHdDQWh4SzBBQUNZcWM4Q0Z0NWNCU2lMZ0FWUFpXVUlNRmNMUVFzQUFCNUQwQUlBWUdZRUxBQ0tKV0RCa3pYZEVCOHl2bEpCWUtZV0dnOEF3R01JV2dBQXpJU0FCVUN4QkN6WUJ0TXNnRG1yZEI4QWdNZDRjWDkvcjJBQUFJV3IxcXNZcmxnS1dBQVVSY0NDcldpNklkNG4vRnMxZ1puN1o5L1dycWtBQUR5SWlSWUFBQVdyMXF1ek1XQnhwTThBeFJDd1lHdWFiamdjN3hVQTVpNnVEM2s3OXlJQUFQQXdnaFlBQUFVU3NBQW9rb0FGdTJEaUZjQWZyQThCQU9EQkJDMEFBQW9pWUFGUUpBRUxkcUxwaHVNUXdpK3FDL0E3UVFzQUFCNU0wQUlBb0FEVmV2VnB6SzJBQlVBNUJDellOU1B5QWY1MG9oWUFBRHpVaS92N2U4VUNBTWpVR0xCWStsRVFvQ2dDRnV4YzB3MnZRd2ovUjZVQi91TG52cTJ2bEFRQWdCOHgwUUlBSUVNQ0ZnQkZFckJnbnk1VkcrQnY0dm9RUVFzQUFINUkwQUlBSUNNQ0ZnQkZFckJncjVwdVdGbzNCdkJWbGJJQUFQQVFnaFlBQUJtbzFxdHFmUE5Vd0FLZ0hBSVc3RjNURFljaGhIT1ZCL2lxaGJJQUFQQVFnaFlBQUJOV3JWZkg0d1NMTi9vRVVBd0JDMUtLd2MwREhRRDRxcU1ZU092YitxUHlBQUR3UFlJV0FBQVRKR0FCVUNRQkM1SnF1cUZ5YndId1EzR3F4WHRsQWdEZ2V3UXRBQUFtUk1BQ29FZ0NGa3pGcFU0QS9GQWxhQUVBd0k4SVdnQUFURUMxWGgyT0R6OEVMQURLSVdEQlpEVGRjQlpDT05FUmdCOWFLQkVBQUQ4aWFBRUFrTkFZc0RnZi85bVhEbEFHQVFzbXBlbUd3M0ZpRmdBL1Zxa1JBQUEvSW1nQkFKQ0FnQVZBa1FRc21LcDR2M0drT3dBUGN0QjBROVczOVVhNUFBRDRGa0VMQUlBOUVyQUFLSktBQlpQVmRNUHhlTjhCd01QRnFSYUNGZ0FBZkpPZ0JRREFIZ2hZQUJUbkxvVHdOb1J3S1dEQnhDM2Rld0E4bXZVaEFBQjhsNkFGQU1DT1ZldlYyZmlRdzhodWdQekZnTVhsR0xENHFKOU1XZE1OaXhEQ0cwMENlTFNGa2dFQThEMkNGZ0FBT3lKZ0FWQVVBUXR5ZEtsckFFL3lVdGtBQVBpZUYvZjM5d29FQUxCRkFoWUFSUkd3SUV0Tk44VDdrZi9XUFlBbis3bHY2eXZsQXdEZ2EweTBBQURZRWdFTGdLSUlXSkN0cGhzT1RiTUFlTGE0UGtUUUFnQ0FyeEswQUFCNHBtcTlXb3dQTTR5WEJjaWZnQVVsT0E4aEhPZ2t3TE5VeWdjQXdMY0lXZ0FBUE5FWXNJZ1RMRTdVRUNCN0FoWVVvZW1HNHhEQ2hXNENQSnVnQlFBQTN5Um9BUUR3U0FJV0FFVVJzS0EwYjNVVVlDdU9Zbml0YitzYjVRUUE0RXVDRmdBQUR5UmdBVkFVQVF1SzAzVER3bjBLd0ZiRnFSYUNGZ0FBL0kyZ0JRREFEMVRyMWZINE1PNlZXZ0ZrVDhDQ2twbG1BYkJkTVdqeFhrMEJBUGlTb0FVQXdEZU1BWXM0d2VLTkdnRmtUOENDb2pYZGNCN0gzT3N5d0ZZdGxCTUFnSzhSdEFBQStJS0FCVUJSQkN3b1h0TU5oK085Q3dEYlpSMFRBQUJmSldnQkFEQVNzQUFvaW9BRmN4THZYdzUwSEdEN21tNm8rcmJlS0MwQUFKOFR0QUFBWnE5YXJ6NjlCZnJMM0dzQlVBQUJDMllsUGdCMER3T3dVL0h2cktBRkFBQi9JV2dCQU16V0dMQTRILzk1Q3hRZ2J3SVd6TldsemxPNDYvRkJ0L3QxVWxtRUVONnFQZ0FBbnhPMEFBQm1SOEFDb0NnQ0ZzeFcwdzJ2UXdnblBnRVU3bno4TysrelRpcVZ5Z01BOENWQkN3QmdOZ1FzQUlvaVlBR21XVkMrWC91MjNqVGRjQ1ZvUVVJdm0yNDQ3TnZhL1FZQUFQOGhhQUVBekVLMVhzVnd4VkxBQWlCN0FoYnd4elNMZUY5enBCWVU3RzY4ZjQ5aTBPSkNzMG1vR2orSEFBRHdPMEVMQUtCbzFYcDFOdjVBNjBFRVFONEVMR0FVMzZ3ZUozUkJ5WmFmVFJEWTZEU0pMUVF0QUFENG5LQUZBRkFrQVF1QVlnaFl3TjlkbXRKRjRXNzd0djdQYXB3WXVHaTY0VU5jNGFEeEpGSXBQQUFBbnhPMEFBQ0tJbUFCVUF3QkMvaUtwaHZpVzlWdjFJYkNuWDNsOURhQ0ZpUzBVSHdBQUQ0bmFBRUFGS0ZhcjE2UEQrUUVMQUR5Sm1BQjM3ZFVId3IzVzkvV1gxdlJzQkV5SXFHRHBodU8rN2ErMFFRQUFJS2dCUUNRdTJxOVdvd1BIRTQwRXlCckFoYndBMDAzbkxubllRYk92M0dLWHd0ZndEN0Y5U0dDRmdBQS9FN1FBZ0RJa29BRlFERUVMT0FCbW00NE5NMkNHVmg5YTJKQTM5YWJwaHQ4QmtncGZnZDlyd01BQUFSQkN3QWdOd0lXQU1VUXNJREhPYmNpamNKOXVpNTh6N1h2QVNSVUtUNEFBSjhJV2dBQVdSQ3dBQ2lHZ0FVOFV0TU54OTlacHdDbE9PL2Ira2ZYaFN2ZkIwakladzhBZ1A4UXRBQUFKcTFhcjQ3SGdNVWJuUUxJbW9BRlBGMzhmK2RBL1NqWWg3NnQzejdnOURZK0JLVFVkRU1WMTlob0FnQUFnaFlBd0NRSldBQVVROEFDbnFIcGhqalY2NVVhVXJpSFRtenhnSnZVRmo2SEFBQUVRUXNBWUdvRUxBQ0tJV0FCMjNHcGpoVHVYZC9XVnc4NXhiNnRiNXB1dUEwaEhQbFFrRWlsOEFBQUJFRUxBR0FxQkN3QWluRTcvajEvTDJBQno5TjB3MWtJNGFVeVVyQzc4WnJ4R0J0QkN4SmFLRDRBQUVIUUFnQklyVnF2RHNkUndlZDJqd05rN2ZlQXhlYjA0aUU3OW9FZmFMcmgwRFFMWnVBeVRxbDQ1R2xlV2FkRFFrZng3M1BmMXNLa0FBQXpKMmdCQUNRaFlBRlFEQUVMMkkybGV5UUtkL3ZFTU5IR0I0UEVxakh3QXdEQWpBbGFBQUI3SldBQlVBd0JDOWlScGh2aVNyVmYxSmZDTFo4eUZhQnY2NnVtRzN3MlNHa2hhQUVBZ0tBRkFMQVhBaFlBeFJDd2dOM3oveGVsdSs3YitqbWY4dzhoaEpjK0pTU3lVSGdBQUFRdEFJQ2RxOWFyczNFc3NJQUZRTDRFTEdBUG1tNklEL0JPMUpyQ25UL3o5SzRFTFVpb1Vud0FBQVF0QUlDZEdRTVdjYi80a1NvRFpFdkFBdmJMLzJ1VTdsM2YxcHRubnVOei8zdDRqb080NHFsdjZ4dFZCQUNZTDBFTEFHRHJCQ3dBaWlCZ0FYdldkTU81K3ljS2Q3ZUZhUlpobkdnQktTMEU0d0FBNWszUUFnRFlHZ0VMZ0NJSVdFQUNUVGNjanZkUlVMTEx2cTAvUHZmODRpU0JwaHZ1ckNZa0lldERBQUJtVHRBQ0FIaTJhcjJLYi9OYzJwTU1rRFVCQzBqcjBrTmpDbmZidC9VMncwUnhmY2lKRHcySkNGb0FBTXljb0FVQThHUmp3R0xwQjA2QXJBbFlRR0pOTjhRSGRtLzBnY0p0WTJYSTU2NThEeUVobnowQWdKa1R0QUFBSGszQUFxQUlBaFl3SFpkNlFlR3UrN1orditWVGpFR0xDeDhjVW1tNllkRzM5WlVHQUFETWs2QUZBUEJnQWhZQVJSQ3dnQWxwdXVHMWV5dG00R3dIcDdqeHdTR3hhZ3o4QUFBd1E0SVdBTUFQVmV2VmNRamhyWWNBQUZrVHNJQnBNczJDMHYzYXQvWE50cyt4Yit1UFRUZDhDQ0c4OUFraWtZVy80UUFBOHlWb0FRQjgweGl3V05vWkRwQTFBUXVZcUtZYjRuM1drZjVRc0x2eCs4U3ViQVF0U0toU2ZBQ0ErUkswQUFEK1JzQUNvQWdDRmpCaFRUZkUrNjF6UGFKd3l6aDVZb2VuZU9VN0N3a2ROZDF3Y3Fza2hRQUFJQUJKUkVGVXVPUFBPQUFBRXlWb0FRRDhoNEFGUUJFRUxDQVA4WjdyUUs4bzJJZStyWGU5Vm1IakEwUmljWDNJZTAwQUFKZ2ZRUXNBSUFZc0RzZmRzZ0lXQVBrU3NJQk1OTjJ3Y04vRkRPeDhZa3ZmMXB1bUczeVdTS2tTdEFBQW1DZEJDd0NZc1RGZ2NUNys4MFlsUUo0RUxDQS9TejJqY0wvMWJYMjFwMU84RGlHYytFQ1J5RUxoQVFEbVNkQUNBR1pJd0FLZ0NBSVdrS0dtRzg0OEZHWUdkajdONGpOWC9wOGlvVXJ4QVFEbVNkQUNBR1pFd0FLZ0NBSVdrS21tR3c1TnMyQUdWbjFiMyt6eE5EYytWQ1IwMEhSREZkZllhQUlBd0x3SVdnREFURlRyMVZMQUFpQnJBaGFRdjNndmRxU1BGT3d1aEhDNTU5UGIxNG9TK0paSzRBY0FZSDVlM04vZmF6c0FGS3hhcjg3R055ZjlxQStRSndFTEtFRFREY2NoaFArbmx4VHV2L3EyM3Z2MXF1bUdHOTkzU09oZDM5Wm5HZ0FBTUM4bVdnQkFvUVFzQUxJbllBRmwyZmRiL3JCdjF5bENGcU9ON3owa1ZDaytBTUQ4Q0ZvQVFHRUVMQUN5SjJBQmhXbTZZUkZDZUtXdkZHNlo4UFN1L0Q5R1FpOFZId0JnZmdRdEFLQVFBaFlBMlJPd2dIS1paa0hwNHVxRXE0VG51UEVKSTZVWXFFdjgvd0FBQUhzbWFBRUFtYXZXcThVWXNEalJTNEFzQ1ZoQXdacHVPUGUyTTRXN1N6ek5Jc1FIM0UwMytKeVIwbUtjckFJQXdFd0lXZ0JBcGdRc0FMSW5ZQUdGYTdyaE1QVURhTmlEeTc2dGJ5WlE2R3ZmalVpb1Vud0FnSGtSdEFDQXpBaFlBR1JQd0FMbUk5NnpIZWczQmJ1ZDBHcWNqZTlJSkNSb0FRQXdNNElXQUpBSkFRdUE3RjJQQVF0anBXRUdtbTQ0RGlIOG90Y1U3cnh2NjQ4VE9jWE5CSTZCK1RxS2YvY25NdDBGQUlBOUVMUUFnSW1yMXF2ak1XRHhScThBc2lSZ0FmTmthZzJsdSs3Yit2MkV6dEYxbHRUaVZBdEJDd0NBbVJDMEFJQ0pFckFBeUo2QUJjeFUwdzBMVThpWWdmTXBuV0tjSk5CMHc1MTFQU1FVZ3haVENoOEJBTEJEZ2hZQU1ERUNGZ0RaRTdBQVRMT2dkTy82dHA3aXFvNTQ3WDAxZ2VOZ25oYjZEZ0F3SDRJV0FEQVJBaFlBMlJPd0FPSTBpM2cvZDZRU0ZPeHVhdE1zUHJNUnRDQWhrNHdBQUdaRTBBSUFFcXZXcThQeGg4b0x2UURJa29BRjhMdW1HdzRuL0FBYXR1V3liK3VQRTYzbWxlOVZwTlIwUXpYUmFTOEFBR3lab0FVQUpQSlp3T0xjSG1HQUxBbFlBRis2ZEY5SDRXNzd0bDVPOVJUN3RyNXF1bUVDUjhLTUxjYkpLZ0FBRkU3UUFnRDJUTUFDSUhzQ0ZzRGZ4TGVZcllCakJzNHlPTVVQSVlTWEV6Z081cW5TZHdDQWVSQzBBSUE5RWJBQXlKNkFCZkE5bDZwRDRhN2p4SWdNVG5FamFFRkNnaFlBQURQeGswWUR3TzVWNjFVTVY5eU0rNEtGTEFEeUVnTVdQMjlPTHhaQ0ZzRFhOTjN3T29Sd29qZ1VMb2RwRnBGck5TbTliTHJoVUFjQUFNcG5vZ1VBN0ZDMVhzVWZJK01PNHlOMUJzaU9DUmJBRDQwUDFFeXpvSFMvOW0xOWs4azViaVp3RE14YkpmQURBRkErUVFzQTJBRUJDNENzQ1ZnQWozSHVuby9DM1kzZmJiTFF0L1dtNllZN2t3UkphQ0ZvQVFCUVBrRUxBTmdpQVF1QXJBbFlBSS9TZE1QeEdMU0FrcDMzYmYweHMvUGJXT2REUXBYaUF3Q1VUOUFDQUxhZ1dxL2lHeXR2QlN3QXNpUmdBVHpWMGx2ekZPNUQzOVp2TXp6RkswRUxFbG9vUGdCQStRUXRBT0FaeG9ERjBvOTRBRmtTc0FDZXJPbUdlQi80UmdVcFhLNFRXellUT0FibTZ5Qk9QT3JiK3NabkFBQ2dYSUlXQVBBRUFoWUFXUk93QUxaaHFZb1U3cmUrclhPOVZyckdrMXBjSHlKb0FRQlFNRUVMQUhnRUFRdUFyQWxZQUZ2UmRNT1orMEZtSU5kcEZxRnY2NDlOTjl4YTdVaEM4YmVEOXhvQUFGQXVRUXNBZUlCcXZZcHZvMXo2UVIwZ1N3SVd3TlkwM1hBNDNoZEN5VllGckQyNHN0NkhoQ3JGQndBb202QUZBSHhIdFY0ZGp4TXMvRUFIa0I4QkMyQVg0bHYrQnlwTHdXNExDUk50Zkk4aklTOXBBQUFVVHRBQ0FMNUN3QUlnYXdJV3dFNDAzUkR2RVM5VWw4SXQ0K3FOQWs1eE00RmpZTWFhYmxqMGJlMStGQUNnVUlJV0FQQVpBUXVBckFsWUFMdG1aUWlsdSs3YittMEo1eGdmY0RmZE1JRWpZY2FxY1lVTkFBQUZFclFBZ0Q4Q0ZwOTJiUXRZQU9SSHdBTFl1ZmhtY2dqaGxVcFR1R1ZocDNkdGhRTUpWWW9QQUZBdVFRc0FabTBNV0p6YnRRMlFKUUVMWUorS2VNc2Z2dU5kZ1dzT05vSVdKTFJRZkFDQWNnbGFBREJMQWhZQVdST3dBUGFxNllaNHozaWs2aFRzcnNCcEZtRmMyL0RMQkk2RGVUcHF1dUd3Yit1UCtnOEFVQjVCQ3dCbVJjQUNJR3NDRnNEZXhZZGtoVDZBaHM5ZDltMTlVMkJGTmhNNEJ1YXRHZ00vQUFBVVJ0QUNnRmtRc0FESW1vQUZrTkxTL1NPRnUrM2J1c2d3VVF5UE5OMXc1LzloRWxvSVdnQUFsRW5RQW9EaVZldlYyZmdEdVhIUEFIa1JzQUNTYXJyaDJOb0JadUM4OEZPTTl4R3ZKbkFjek5OQzN3RUF5aVJvQVVDeEJDd0FzaVZnQVV6Rlc1MmdjTmQ5Vzc4di9CdzNnaFlrVkNrK0FFQ1pCQzBBS0k2QUJVQzJmb3M3NGdVc2dDbG91dUYxQ09GRU15aGM2ZE1zd2pqUjRtSUN4OEU4SGNUcFNIR05qZjREQUpSRjBBS0FZZ2hZQUdUcjNUakJ3Zy9Rd0pSYzZnYUZlOWUzOWFiMGsremIrcXJwaGdrY0NUTzJNQ0VKQUtBOGdoWUFaSzlhcnhiakQrRXZkUk1nS3dJV3dDUTEzU0M4UytudVpqTE40cE1QdmkrU2tQVWhBQUFGRXJRQUlGdGp3R0pwcEROQWRnUXNnTWxxdXVGd1pnK2dtYWRsMzlZZlozVG1WNElXSkxSUWZBQ0E4Z2hhQUpBZEFRdUFiQWxZQURtSWs5SU9kSXFDM2ZadFBiZlZPTVd2U0dIU2hId0FBQW9rYUFGQU5nUXNBTElsWUFGa29lbUdPTjc5alc1UnVMTVpObGpRZ3FTYWJsajBiWDJsQ3dBQTVSQzBBR0R5cXZYcWVIeXo4SlZ1QVdSRndBTEl6ZHplOG1kK3J1ZjRzTGR2NjAzVERYZW0xWkJRTmE2d0FRQ2dFSUlXQUV6V0dMQlllcXNRSURzQ0ZrQjJtbTQ0TXptTkdaampOSXRQTnY0Zko2R0ZNQjhBUUZrRUxRQ1lIQUVMZ0d3SldBQlphcnJoY0x6L2hKTDkycmYxbksvUlY0SVdKRlFwUGdCQVdRUXRBSmdNQVF1QWJBbFlBTGs3RHlFYzZTSUZ1eE1tK2oxb2NUR0I0MkNlam1Lb3IyL3JqL29QQUZBR1FRc0FrcXZXcTA5dkVQNmlHd0JaRWJBQXN0ZDB3L0VZdElDU25YdkErL3ZxRUVncHJnOTVyd01BQUdVUXRBQWdtVEZnY1Q3K085QUpnR3dJV0FBbHVYUXZTdUUrOUczOWR1NU5qa0dUcGh0dVRhOGhvVXJRQWdDZ0hJSVdBT3lkZ0FWQXRnUXNnS0kwM1JEZkxuNmxxeFRPeEpZL1hWbFZTVUlMeFFjQUtJZWdCUUI3STJBQmtDMEJDNkJVbHpwTDRYN3IyL3BLay85akkyaEJRcFhpQXdDVVE5QUNnTDJvMXF1bGdBVkFkZ1FzZ0dJMTNYQVdRbmlwd3hUc3pqU0x2eEU2SWFXRHBodXF2cTAzdWdBQWtEOUJDd0IycWxxdjRnL1lTM3R3QWJJaVlBRVVyZW1HUTlNc21JSEx2cTFkeXo4VEgzQTMzVENaNDJHV3FuR3lDZ0FBbVJPMEFHQW5CQ3dBc2lSZ0FjeUZTV3VVN2xhWTZKdXVRd2duRXowMnlyY0lJYnpWWndDQS9BbGFBTEJWQWhZQVdSS3dBR2FqNlliakVNS0ZqbE80WmQvV0h6WDVxemFDRmlSVUtUNEFRQmtFTFFEWWltcTllajIrTVNWZ0FaQVBBUXRnanJ4SlRPbXUrN2IyT2YrMnF4RENMMU05T0lyM1Vvc0JBTW9nYUFIQXMxVHIxV0tjWU9HTklJQjhDRmdBczlSMHc4SjlLek93MU9UdjJrejQySmlCZUMzcTIvcEtyd0VBOGlab0FjQ1RDRmdBWkVuQUFwZzdiL2xUdW5jZTRINWYzOVkzVFRmY21zWklRb3R4c2dvQUFCa1R0QURnVVFRc0FMSWtZQUhNWHRNTjV4NnNVcmk3RU1LNUpqL0l4dDhERXFvVUh3QWdmNElXQUR5SWdBVkFsZ1FzQVA0SVdSeGFwOEFNWFBadC9WR2pIeVFHTFY1bGNKeVVTZEFDQUtBQWdoWUFmRmUxWGgyUFAwcS9VU21BYkFoWUFQeFZ2Sjg5VUJNS2R0dTN0VERSdzhXMURSZTVIQ3pGT1dxNjRUaXVzZEZhQUlCOENWb0E4RlVDRmdCWkVyQUErRUxURGZITjRWL1VoY0paR2ZJSWZWdGZOZDJRemZGU3BIaHRjczhPQUpBeFFRc0Eva0xBQWlCTEFoWUEzM2FwTmhUdXVtL3I5NXI4YUI5Q0NDOHpPMmJLRWRleit2OFdBQ0JqZ2hZQS9FN0FBaUJMQWhZQTM5RjB3K3NRd29rYVVUalRMSjdtU3RDQ2hDckZCd0RJbTZBRndNeFY2OVhoK01QY3ViM1ZBTmtRc0FCNEdOTXNLTjJ2ZlZ0dmRQbEoxSTJVaEFBQkFESW5hQUV3VXdJV0FGa1NzQUI0b0tZYjRyUzJJL1dpWUhmalZFS2U1a3JkU0tucGhrcFFDZ0FnWDRJV0FETWpZQUdRbmJ0eGY3T0FCY0FETmQxd2FKMENNN0RzMi9xalJqOU4zOVkzVFRmYytWNU1RZ3VUVlFBQThpVm9BVEFUQWhZQTJia2JSOTVmYms0dlBFUUJlSnhMOTd3VTdyWnZhNnR4bm05amhRTUpWWW9QQUpBdlFRdUFHYWpXcXpNL05nTmtROEFDNEJtYWJvaHZDTDlSUXdwM3BzRmJjU1ZvUVVLQ0ZnQUFHUk8wQUNqWUdMQ3dteG9nRHdJV0FOdXhWRWNLZDkyMzlaVW1iMFdzNDBVQjUwR2VYc1pWVjFZQUFRRGtTZEFDb0VBQ0ZnQlpFYkFBMkpLbUc4NjhuYzRNbUdheFBadFNUb1JzVldQZ0J3Q0F6QWhhQUJSRXdBSWdLd0lXQUZzVTN3bzJ6WUlaV1BWdGZhUFIyeEVuQ1RUZDhDRk9GaWpoZk1qU1F0QUNBQ0JQZ2hZQUJhaldxL2pGL0syQUJVQVdCQ3dBZHVQYy9UQ0YrM1FQd1hadEJDMUlxRko4QUlBOENWb0FaR3dNV0N5TlJ3Yklnb0FGd0k0MDNYQWNRcmhRWHdwM0hpY3dhUExXeGFERm04TE9pWHdzOUFvQUlFK0NGZ0FaRXJBQXlJcUFCY0R1ZWN1ZjBuM28yL3F0THUrRXRRMmtkQkREZ2xZQ0FRRGtSOUFDSUNNQ0ZnQlpFYkFBMklPbUcrSTk4aXUxcG5Ebkdyd2JmVnR2bW00bzhkVEl4NmQxc0FBQVpFVFFBaUFEMVhwMVBIN3BGckFBbUQ0QkM0RDlNczJDMHIzcjI5clVoZDI2OW4yYmhDckZCd0RJajZBRndJU05BWXVsZmJFQVdSQ3dBTml6cGh2T1FnZ3YxWjJDM1kzZkNkbXRLMEVMRWhLMEFBRElrS0FGd0FRSldBQmtSY0FDSUlHbUd3NU5zMkFHTHZ1MnZ0SG9uZHNVZm41TW01QVBBRUNHQkMwQUprVEFBaUFyQWhZQWFjWDc1Z005b0dDM3drUjdJMmhCVWswM0xLd0lBZ0RJaTZBRndBUlU2OVdudC9FRUxBQ21UOEFDSUxHbUcySkErUmQ5b0hETHZxM2RhK3hCbkJyU2RFTU10aHdWZjdKTVZUV3VzQUVBSUJPQ0ZnQUpqUUdMOC9HZnQvRUFwazNBQW1BNjN1b0ZoYnZ1MjlybmZMODJnaFlrVkNrK0FFQmVCQzBBRWhuWGhHd0VMQUFtVDhBQ1lFTGllSFg3N0ptQmMwM2V1emhONE5YTXpwbnBXT2dGQUVCZWZ0SXZnRFEycHhjM0lZUWI1UWVZckJpd1dJVVFqamVuRjBzaEM0REo4SlkvcFh2WHQvVkdsL2RPelVucHFPbUdReDBBQU1pSG9BVkFXbjRrQnBnZUFRdUFpV3E2NGR4b2Z3cDNaNXBGR24xYlg4M3h2SmtVNjBNQUFESmlkUWhBV3U5RENQL1dBNEJKc0NJRVlNTEdOMzJYZWtUaEx2dTJkaCtTenJYVlJDUzBHRmZZQUFDUUFSTXRBQklhMTRkODBBT0FwRXl3QU1oRERNTWQ2QlVGdSszYldwZ29MZXREU0dtaCtnQUErVERSQWlDOXQ2WmFBQ1JoZ2dWQUpwcHVpT1BVMytnWGhiTXlKRDFCQzFLeU9nUUFJQ01tV2dDazkxNFBBUGJLQkF1QS9GenFHWVc3N3R2YWQ4UDByRzBncFlNeFdBZ0FRQVlFTFFBU3N6NEVZRzhFTEFBeTFIVEQ2eERDaWQ1UnVETU5UcTl2NjV2eG5oRlNFYlFBQU1pRTFTRUEwMkI5Q01EdVdCRUNrRGZUTENqZHIrTURmcVloVHJWNHBSY2tJbWdCQUpBSlFRdUFhWGd2YUFHd2RRSVdBSmxydW1FWlFqalNSd29XNzFlV0dqd3BHMEVMRWxvb1BnQkFIcXdPQVppQWNYM0l0VjRBYklVVklRQUZhTHJoT0lSd3JwY1VidG0zdFh1VmFibWFld0ZJNnFYeUF3RGtRZEFDWURyZTZnWEFzd2hZQUpRbHZ1Vi9vS2NVN0xadmE2dHhwbWN6OXdLUVZ0TU5wbG9BQUdSQTBBSmdPdDdyQmNDVENGZ0FGR1o4eVBSR1h5bmNtUVpQenpoaDVNUGM2MEJTbGZJREFFeWZvQVhBUkl3UEJuL1RENEFIdXcwaC9HOEJDNEFpTGJXVnd2M1d0N1VWRmRObHFnVXBtV2dCQUpDQmYyZ1N3S1RFcVJhdnRBVGd1MkxBSWdZcnJGd0NLRkRURGZFdC94TzlwWERuR2p4cFY2YnFrSkNKRmdBQUdURFJBbUJhckE4QitMWVlzUGl2emVuRnNaQUZRSm1hYmpnTUlWeHFMNFZiOVcxOW84bVRacUlGS1IyTjEwTUFBQ1pNMEFKZ1Fxd1BBZmdxQVF1QStZaHYrUi9vTndXN0V5YWF2cjZ0QlMxSXpmb1FBSUNKc3pvRVlIcXNEd0g0Z3hVaEFEUFNkTU54Q09GQ3p5bmNlZC9XSHpVNUM5ZldHSkZRWmVvcEFNQzBDVm9BVEUvOEl2M2YrZ0xNbUlBRndEeDV5NS9TWGZkdDdmNG1IMWVDRmlSa29nVUF3TVJaSFFJd01kYUhBRE5tUlFqQVREWGRzRERWalJsWWFuSldyQThoSlNFZkFJQ0pFN1FBbUNZUEdJRTVFYkFBd0RRTFN2ZXViK3NyWGM2S2ZwRlUwdzJWRGdBQVRKZWdCY0FFYlU0djR2cVFPNzBCQ2lkZ0FVQjhrSFFlUW5pcEVoVHN6alNML1BSdC9YRzhYNFZVQkMwQUFDYnNINW9ETUZreGJQRkdlNEFDeFIrc2w4SVZBRFRkY09nQk5ETncyYmYxalVabkthNFBPWnA3RVVobVllSXBBTUIwQ1ZvQVRKZWdCVkFhQVFzQXZoUkRGZ2VxUXNGdXJjYkpXbHdmOG1ydVJTQVpFeTBBQUNiTTZoQ0FpYkkrQkNpSUZTRUEvRTNURGNjaGhGOVVoc0tkanlzb3lOTkczMGpJV2kwQWdBa1R0QUNZdHZmNkEyUk13QUtBNzNGdG9IVFhmVnY3VHBleHZxMnY1bDREMG1xNllhRUZBQURUWkhVSXdMUlpId0xreUlvUUFMNXJmSEIwb2tvVTdseURpM0R0N3hVSkxjWVZOZ0FBVEl5SkZnQVRabjBJa0JrVExBQjRLTmNKU3ZldWIydHJKOHFnajZSVXFUNEF3RFNaYUFFd2ZhWmFBRk5uZ2dVQUQ5WjB3ektFY0tSaUZPek9OSXVpeEdrQ3Y4eTlDQ1FqYUFFQU1GRW1XZ0JNbndlWHdGU1pZQUhBb3pUZGNPZ0JORE53MmJmMVI0MHVob2tXcEhUVWRNT3hEZ0FBVEkrZ0JjREViVTR2cnNhSG1RQlRJV0FCd0ZOZGhoQU9WSStDM2ZadHZkVGdjdlJ0ZldPbEo0bVphZ0VBTUVHQ0ZnQjVlSzlQd0FRSVdBRHdaRTAzVkZiaU1RTm5tbHlrcTdrWGdLUVd5ZzhBTUQzLzBCT0FMTHkxRXhaSUtBWXNsc0lWQUR6VHBRSlN1T3UrclQyUUwxTmNIL0pxN2tVZ0dSTXRBQUFteUVRTGdBeHNUaTgyMW9jQUNaaGdBY0JXTk4zd09vUndvcG9VempTTGNnblFrSkxySndEQUJBbGFBT1REK2hCZ1h3UXNBTmlhcGhzT1RiTmdCbjd0Mi9wR284dGtVZ21wamV1M0FBQ1lFS3REQVBKaGZRaXdhMWFFQUxBTDV5R0VJNVdsWUhmeEhrcURpL2NoaFBCeTdrVWdtY1c0d2dZQWdJa3cwUUlnRTlhSEFEdGtnZ1VBTzlGMHcvRVl0SUNTTGZ1Mi9xakR4Zk9RbTVSTXRBQUFtQmdUTFFEeTh0NVVDMkNMVExBQVlOZmlXLzRIcWt6QlB2UnRiVFhPUE1UMUlXL21YZ1NTRWJRQUFKZ1lFeTBBOHVJSFBHQWJUTEFBWU9lYWJsaDRLTWtNbU5neUh5WmFrTkxMcGhzT2RRQUFZRG9FTFFBeXNqbTl1Qm4zd2dJOGhZQUZBUHNrSkV6cGZ1dmIra3FYNTZGdjZ4aTB1SnQ3SFVqS1ZBc0FnQW14T2dRZ1AvSGg2TC8xRFhnRUswSUEyS3VtRzg3aTI3ZXFUdUZNczVpZkdMWTRtWHNSU0dZeHJyQUJBR0FDQkMwQTh2TmUwQUo0SUFFTEFQWnVIRzF1bWdXbFcvVnRmYVBMczNNbGFFRkNDOFVIQUpnT3EwTUFNbU45Q1BBQTF5R0UvOCtLRUFBU2lXLzVIeWcrQmJzVEpwcXR6ZHdMUUZKV2h3QUFUSWlnQlVDZVBEZ0Z2aVlHTEg3ZW5GNHNOcWNYNzFVSWdIMXJ1dUU0aEhDaDhCVHV2Ry9yajVvOFM5WTJrTkxCZUowRkFHQUNCQzBBOHVRQkt2QzV6d01XZnZ3RklDVnYrVk82Njc2dEJkOW5hZ3pZM002OURpUmxmUWdBd0VRSVdnQmt5UG9RWUNSZ0FjQmtOTjBRSC82ODBoRUt0OVRnMlhQZlRVcldod0FBVE1RL05BSWdXL0V0cW45ckg4eFNERmdzaFNzQW1CaHYrVk82ZDMxYnUvOWlFMEo0TS9zcWtJcWdCUURBUkFoYUFPVHJ2YUFGekk2QUJRQ1QxSFREZVFqaFNIY28ySjFwRm96Y2k1UFNpZW9EQUV6RGkvdjdlNjBBeUZTMVhyMDNuaGxtUWNBQ2dNbHF1dUV3aEJCWDJ4M29FZ1ZiOVcwdGFNSHZtbTd3Z3lvcC9XeTZEZ0JBZWlaYUFPUk4wQUxLSm1BQlFBNldRaFlVN2xiSWdpOWNteXhBUXBYSktnQUE2UWxhQU9RdEJpMytXdytoT0FJV0FHU2g2WWJqRU1JdnVrWGh6aldZTDJ3RUxVaW9VbndBZ1BSKzBnT0FmRzFPTHo2R0VIN1RRaWhHREZqOHZEbTlXQWhaQUpDSnR4cEY0YTc3dG42dnlYekJ2VG9wTFZRZkFDQTlRUXVBL1BuUkQvSW5ZQUZBZHBwdWVPMk5ibWJBTkF1K1pxTXFKSFRVZE1PaEJnQUFwQ1ZvQVpBL1FRdklsNEFGQURtNzFEMEs5NjV2YXcvVStadStyVzlDQ0xjcVEwS21XZ0FBSkNab0FaQTU2ME1nU3dJV0FHU3Q2WVpsZktOV0Z5blluV2tXL0lBUURpbFZxZzhBa05ZLzFCK2dDSEdxeFN1dGhNbUxBWXVsY0FVQU9SdkhsWHNBVGVtV2ZWdC8xR1crWStON09BbVphQUVBa0ppZ0JVQVpZdERpdi9VU0prdkFBb0NTeEpVaEJ6cEt3Vzc3dHJZYWh4K0o5L1lYcWtRaUpsb0FBQ1JtZFFoQUFjYjFJZS8wRWliSGloQUFpdEowUTN5RDlvMnVVcmd6RGVaSCtyWjJmMDlLQjAwM0NGc0FBQ1Jrb2dWQU9kNzcwUnNtd3dRTEFFcTExRmtLZCswQk9vL3dJWVR3VXNGSXBCcFgyQUFBa0lDSkZnQ0YySnhleEtERm5YNUNVaVpZQUZDc3BodmlXLzRuT2t6aFRMUGdNZHp6azVLSkZnQUFDWmxvQVZBV1V5MGdEUk1zQUNoYTB3MkhwbGt3QTcvMmJYMmowVHlDYVFLa3RGQjlBSUIwQkMwQXlpSm9BZnNsWUFIQVhKeUhFSTUwbTRMZENSUHhCSUlXcEdSdERRQkFRaS91NysvVkg2QWcxWHIxTVlSd29LZXdVd0lXQU14RzB3M0g0OE5FOTVpVTdMLzZ0bjZyd3p4VzB3MitnNVBTejMxYisxNEtBSkNBaVJZQTVUSFZBblpId0FLQU9icjBFSkhDZlJDeTRCbGlFTzFFQVVta0NpSDRmZ29Ba01CUGlnNVFuUGRhQ2xzWEF4WS9iMDR2RmtJV0FNeEowdzF4Ly9zclRhZHc1eHJNTS9oK1FFb0wxUWNBU0VQUUFxQXdtOU9MOStOK1llRDVCQ3dBbUx2THVSZUE0djFtN0Q3UDVQTkRTcFhxQXdDa0lXZ0JVQ1pUTGVCNUJDd0FtTDJtRzg1Q0NDL25YZ2VLWjVvRno3VlJRUkk2YXJyaFdBTUFBUFpQMEFLZ1RONDhoS2NSc0FDQVAwSVdoKzRwbVlGVjM5WTNHczF6OUczOU1ZUndxNGdrWktvRkFFQUNnaFlBQmRxY1htejgwQU9QSW1BQkFIOFYzL0kvVUJNS2RpdE14QmI1RGtGS2doWUFBQW44UTlFQmloWFhoL3lpdmZCZE1XQ3hGSzRBZ0QrTkk4Z3ZsSVRDTGNkSkJMQU44V1dITnlwSklndUZCd0RZUDBFTGdISzlGYlNBYnhLd0FJQnZlNnMyRk82NmIydWZjN2JKOXdwU09sRjlBSUQ5ZTNGL2Y2L3NBSVdxMXF1NGIvaElmK0UvM28wQkM3dTRBZUFybW02SWI4WCtqOXBRdUovN3R2WmduSzFxdXNHUHJLVDByNzZ0TnpvQUFMQS9KbG9BbE0zNkVQaURnQVVBUEl5My9DbmRPeUVMZHVUYVpBRVNxc1lWTmdBQTdNbFBDZzFRTkQrVU0zY3hZUEhQemVuRm1aQUZBSHhmMHczbnBxRlJ1THNRd3JrbXN5TUNQS1MwVUgwQWdQMFN0QUFvMk9iMElyN05jS3ZIekpDQUJRQThRdE1OaDNINms1cFJ1TXUrclQ5cU1qdGltZ0FwVmFvUEFMQmZWb2NBbE0vNkVPYkVpaEFBZUpvWXNqaFFPd3AyMjdlMU1CRzdKR2hCU2k5Vkh3Qmd2MHkwQUNqZnBSNHpBeVpZQU1BVE5kMVFDZVl5QTFhR3NGTjlXOStZS0VsS1RUZFlId0lBc0VlQ0ZnQ0ZHeDg2ZjlCbkNpVmdBUURQSjVoTDZhNzd0bjZ2eSt5QnFSYWtKR2dCQUxCSFZvY0F6TVBiRU1LLzlacUNXQkVDQUZ2UWRNUHJFTUtKV2xJNDB5ellsNnNRd2l2VkpwRks0UUVBOWtmUUFtQWUzZ3RhVUFnQkN3RFlMdE1zS04ydmZWdWJNc0MrK0t5Umtva1dBQUI3OU9MKy9sNjlBV2FnV3EvaUR6NHY5WnBNQ1ZnQXdKWTEzYkFNSVZ5b0t3VzdDeUVjOTIzOVVaUFpsNlliL05oS1N2L3MyOXIzWmdDQVBURFJBbUErckE4aFJ3SVdBTEFEVFRjY1c2ZkFEQ3lGTEVqZ2c1Y2NTQ2l1RC9IOUdRQmdEd1F0QU9iRCtoQnlJbUFCQUxzVnAxa2NxREVGdSszYjJtb2NVcmdTdENDaHhmajdEd0FBTy9hVEFnUE13L2pBK29OMk0zRXhZUEhQemVuRm1aQUZBT3hHMHczeEljd2I1YVZ3WnhwTUlodUZKNkZLOFFFQTlzTkVDNEI1c1Q2RXFUTEJBZ0QyWjZuV0ZPNjZiK3NyVFNZUm56MVNPbEY5QUlEOU1ORUNZRjZNajJScVRMQUFnRDFxdXVITVF4aG13RFFMa3VuYk9uNnZ1ZE1CVW1tNndWUUxBSUE5RUxRQW1KSHhRZlp2ZXM1RWZCQ3dBSUQ5YWJyaDBEUUxabUExUHVpR2xLd1BJYVdGNmdNQTdKNmdCY0Q4bUdyQlZMeXMxcXRqM1FDQXZUa1BJUndwTndXTFV3UXVOWmdKc0Q2RWxFeTBBQURZQTBFTGdQa1J0R0JLWHVzR0FPeGUwdzB4M0hpaDFCVHV2Ry9yajVyTUJBaGFrSktKRmdBQWV5Qm9BVEF6bTlPTGo5YUhNQ0gyWndQQWZuakxuOUo5Nk52NnJTNHpFVmFIa05MUnVDNE1BSUFkRXJRQW1DZFRMWmdLNjBNQVlNZWFib2h2dHI1U1p3cDNyc0ZNeFRoWjVZT0drSkQxSVFBQU95Wm9BVEJQZ2haTWlmVWhBTEJicGxsUXV0LzZ0cmFxZ2FreDFZS1VyQThCQU5neFFRdUFHYkkraElueDlpRUE3RWpURFhGTjEwdjFwV0IzN2llWktFRUxVaEswQUFEWU1VRUxnUGt5MVlLcE9LcldLMk5OQVdETHh2M3NwbGxRdXN1K3JXOTBtUWt5WllXVWZNY0dBTmd4UVF1QStYby92djBGVTNDbUN3Q3dkY3NRd29HeVVyQmJZU0ttcW05ckV5MUk2YURwaG1NZEFBRFlIVUVMZ0prYTE0ZVlhc0ZVdk5ZSkFOaWU4ZUhLTDBwSzRaWjlXMy9VWkNic1duTkl5UG9RQUlBZEVyUUFtRGRCQzZiQytoQUEySzYzNmtuaHJ2dTI5amxuNnF3UElTWGZzUUVBZGtqUUFtREdOcWNYMW9jd0pkYUhBTUFXTk4wUTMyQTlVVXNLdDlSZ01tQjlDQ2tKV2dBQTdKQ2dCUUNtV2pBVjFvY0F3SFo0eTUvU3ZldmIycVFBY3VCelNrcENsd0FBT3lSb0FZQ2dCVk1SMTRmWUlRc0F6OUIwdzNtOHBxb2hCWXNUK2M0MW1CejBiZjB4aEhDcldhUXlUcmtDQUdBSEJDMEFaczc2RUNiRytoQUFlS0ttR3c2dFUyQUdMc2VIMTVBTDYwTkl5Zm9RQUlBZEViUUFJSmhxd1lSWUh3SUFUM2NaUWpoUVB3cDIyN2UxTUJHNXNUNkVsRXkwQUFEWUVVRUxBSUtnQlJOeVVLMVh3aFlBOEVoTk44UTNWdCtvRzRXek1vUWNtV2hCU2laYUFBRHNpS0FGQU5hSE1EV0NGZ0R3ZUpkcVJ1R3UrN1lXRUNjN2ZWdWJhRUZLUitOcU1RQUF0a3pRQW9CUDNxb0VFeUZvQVFDUDBIUkR2SGFlcUJtRk85TmdNbmF0ZVNSa2ZRZ0F3QTRJV2dEd2lhQUZVMkY5Q0FBODBQaVdxbWtXbE83WHZxMXZkSm1NV1I5Q1N0YUhBQURzZ0tBRkFML2JuRjdFSDM1dVZZT0pFTFFBZ0ljNWoyUEIxWXFDeFJXSFN3MG1jNElXcEdTaUJRREFEZ2hhQVBBNU80K1ppamZWZW1XUExBQjhSOU1OeDJQUUFrcTI3TnY2b3c2VHVTc05KQ0VUTFFBQWRrRFFBb0RQV1IvQ2xKaHFBUURmRjkveVAxQWpDbmJidDdYVk9HUnZYSDF6cDVNa2N0QjBnN0FGQU1DV0NWb0E4Qi9XaHpBeGdoWUE4QTFOTjhReDRHL1VoOEtkYVRBRk1kV0NsQVF0QUFDMlROQUNnQzlaSDhKVXZMSStCQUMrYWFrMEZPNjN2cTA5bUtZa0c5MGtJVUVMQUlBdEU3U29PNlUwQUFBZ0FFbEVRVlFBNEV2V2h6QWxwbG9Bd0JlYWJvaHYrWitvQzRVNzEyQUtJemhFU2d2VkJ3RFlMa0VMQVA3aS8yZnZibTRpeWJJMkFJY0hpUWZnQVhnQXE5aFdTU0dGMklFSG5SNlFlSkR0UWJGRElZVlV0WTFWNFVIaUFYaEFldkFwWnVLYjZaNnVIMzR5ODU2NDkzbE1PSWVpZ0h6dmVkV0hFSXlnQlFEOFJkTU40N1dudFptUXVkdStyWjhzbWN5NGFFRktwNllQQUxCYmdoWUEvSWlyRmtReDFvZWMyQVlBL01mNHluOWhIR1JzSzB4RWp2cTJmcW1xNnRGeVNhWHBCbGN0QUFCMlNOQUNnQjhSdENBU1Z5MEE0Tjhma0l6aHd4dXpJSFBMNlFOcHlKR3JGcVFrYUFFQXNFT0NGZ0Q4dytieTVzbExHd0s1dGd3QStCZXYvTW5kWTkvV1F0L2s3THZ0a3RDWjRRTUE3STZnQlFBLzR3K2NSSEdxUGdTQTBrM252aitWUGdleXQ3UmlNdWVpQlNrSldnQUE3SkNnQlFBLzg5VmtDRVI5Q0FDbGM4MkMzTjMxYmUyMVAxbnIyM29NV214dG1VU09weG95QUFCMlFOQUNnQjlTSDBJdzZrTUFLRmJURGVNci8xTmZBV1JzL09CNVpjRVV3bFVMVW5MVkFnQmdSd1F0QVBnVjlTRkVvVDRFZ0NJMTNYRGtBMmdLc083YitzbWlLWVRMTGFRa2FBRUFzQ09DRmdEOGl2b1FJbkhWQW9BU2pTR0xoYzJUc1dmVk9CVEdSUXRTdWpCOUFJRGRFTFFBNEtmVWh4Q01vQVVBUlpsNjFQK3dkVEszNnR2NnhaSXBpSXNXcEhSdStnQUF1eUZvQWNEdnFBOGhpdU96KzF0blRnRW9pWi9EeU4xRDM5YSt6aW5LRkN4NnRuVlNhYnJCNzlVQUFEc2dhQUhBNy9qREo1RzRhZ0ZBRVpwdXVQRHFsQUlzTFpsQ3VXcEJTb0lXQUFBN0lHZ0J3Qzl0TG0vRzF6YmZUSWtnUGxzRUFJVVFkaVYzZDMxYmIyeVpRdm5hSjZVTDB3Y0ErRGhCQ3dCZTQ2c3BFWVQ2RUFDeTEzVERhdncvejZiSjJOWTFDd29uYUVGS2ZxY0dBTmdCUVFzQVhrUFFna2pVaHdDUXJhWWJqbndBVFFIV2ZWdS9XRFNsNnR0YWRRZ3BuVTQvYndBQThBR0NGZ0Q4bHZvUWdsRWZBa0RPMWxWVkxXeVlqRDMzYmIyeVlLZ2VqSUNFWExVQUFQZ2dRUXNBWHN0VkM2SVk2ME9FTFFESVR0TU40NGNlVnpaTDVsd25nMzlUSDBKS0Y2WVBBUEF4Z2hZQXZKYWdCWkVJV2dDUW83V3RrcmtIbFFud0gvNHRrSktMRmdBQUh5Um9BY0NycUE4aEdFRUxBTExTZE1QNHl2L2NWc21jYXhid1h5NWFrSktMRmdBQUh5Um9BY0JidUdwQkZBdjFJUURrb3VtR282cXFWaFpLNXY3czIvckprdUhmcG44UFcrTWdrVVhURFNlR0R3RHdmb0lXQUx6RlYzOElJaEJCQ3dCeXNheXE2dGcyeWRoV21BaCtTSDBJS2FrUEFRRDRBRUVMQUY1dHFnOXgxWUlvQkMwQW1MM3BOZW5TSnNuY3FtL3JGMHVHZjFBZlFrcnFRd0FBUGtEUUFvQzNFclFnQ3ZVaEFPUmdmT1cvc0VreTl0aTM5ZHFDNFlkY3RDQWxGeTBBQUQ1QTBBS0FOOWxjM3FnUElaSnIyd0JncnBwdUdGK1NYbGtnbVhPeEJYNmliMnRCQzFJNk4zMEFnUGNUdEFEZ1BWeTFJSXBQWi9lM1I3WUJ3RXg1NVUvdXZ2a2dHWDdyMFloSXBla0dWeTBBQU41SjBBS0E5eEMwSUJMMUlRRE1UdE1ONDFXbVU1c2pjNjVad085dHpJaUVMZ3dmQU9COUJDMEFlRFAxSVFRamFBSEFyRFRkY09TYUJRVzQ3ZHY2eWFMaHQxeDlJU1VYTFFBQTNrblFBb0QzY3RXQ0tOU0hBREEzNHl2L2hhMlJzYTB3RWJ5YWl4YWs1S0lGQU1BN0NWb0E4RjZDRmtUaXFnVUFzOUIwdzBsVlZUZTJSZWFXZlZ1L1dETDhYdC9XR3hjalNlaDR1clFGQU1BYkNWb0E4QzVUZmNpejZSSEV0VVVBTUJOZStaTzdoNzZ0djlneXZJbXJGcVNrUGdRQTRCMEVMUUQ0Q0ZjdGlPTDg3UDcyeERZQWlLenBodkU4OXlkTEluTXJDNFkzKzI1a0pLUStCQURnSFFRdEFQZ0lMOVdJUkgwSUFOSDUyWW5jM2ZWdDdRTmplRHYvYmtoSjBBSUE0QjBFTFFCNHQ4M2x6VVo5Q0lHb0R3RWdyS1libG1NUHVnMlJzYTFyRnZCdXFrTklTWFVJQU1BN0NGb0E4RkhxUTRqaVZIMElBQkUxM1hEa0EyZ0tzTzdiK3NtaTRlMzZ0bjd4aUlHRUZrMDMrRjBhQU9DTkJDMEErQ2duc0lsRWZRZ0FFWTBoaTRYTmtMSHhBK0sxQmNPSHFBOGhKZlVoQUFCdkpHZ0J3SWVvRHlFWTlTRUFoTkowdzNpTyt3OWJJWFBMNlVVKzhIN3FRMGhKZlFnQXdCc0pXZ0N3QytwRGlFSjlDQURSZU9WUDdoNzZ0dmI3QUh5Y2l4YWs1S0lGQU1BYkNWb0FzQXZxUTRoa2FSc0FSTkIwdzFocGRXNFpaTTdQWHJBRGZWdTdhRUZLcDZZUEFQQTJnaFlBZkpqNkVJTDViQ0VBQk9HYUJibTc4K0V3N05TRGNaSkswdzJ1V2dBQXZJR2dCUUM3NG9NRW9qZyt1Ny9WTHd0QVVrMDNyTWIvazJ5QmpHMWRzNENkRTF3aUpiOUhBd0M4Z2FBRkFMdWlsNWxJcm0wRGdGU2Fiamp5QVRRRldQVnQvV0xSc0ZQZmpaT0VYTFFBQUhnRFFRc0FkbUp6ZWZOVVZkV2phUktFK2hBQVVob3ZmUzFzZ0l3OTkyM3RvaDNzbm9zV3BPU2lCUURBR3doYUFMQkxYMHlUSU5TSEFKREUxRzkrWmZwa3p2VXcySU8rcmNjSERNOW1TeUxIMDFVdUFBQmVRZEFDZ0YxU0gwSWtQZ0FBSUlXVnFaTzVoNzZ0MVJ2QS9yaHFRVXJxUXdBQVhrblFBb0NkVVI5Q01JSVdBQnhVMHczai96M25wazdtL0l3Rit5Vm9RVW91UXdJQXZKS2dCUUM3cGo2RUtCWm45N2VmYlFPQVE1aE9iYnRtUWU3K25Lb05nUDF4TVlhVVhMUUFBSGdsUVFzQWRrMTlDSkVJV2dCd0tNdXgyOXkweWRoV21BajJUelVQaWJsb0FRRHdTb0lXQU95VStoQ0NFYlFBWU8rYWJqaVpnaGFRczJYZjFpODJEQWZoZDJwU1dUVGRJR3dCQVBBS2doWUE3TVBhVkFsQ2ZRZ0FoekQrN0xNd2FUTDIyTGUxaWtBNEhGY3RTRW5RQWdEZ0ZRUXRBTmdIOVNGRUltZ0J3TjQwM1RCMm1YOHlZVExuWWdzYzFzYThTZWpDOEFFQWZrL1FBb0NkMjF6ZWpDZUZ2NWtzUVh3K3U3ODlzZ3dBOXNRbEwzTDNyVzlycit2aHNQeWJJeVVYTFFBQVhrSFFBb0I5Y2RXQ0tCYXVXZ0N3RDAwM1hGZFZkV3E0Wk00MUN6aXd2cTJmcXFyYW1qdUorTmtHQU9BVkJDMEEyQmRCQ3lJUnRBQmdwNXB1T0hMTmdnTGNUaC80QW9lblBvUmtwbW8wQUFCK1FkQUNnTDFRSDBJd245U0hBTEJqeStscUV1VHFXWmdJa2xJZlFrcUNGZ0FBdnlGb0FjQSt1V3BCSks1YUFMQVRUVGVjVkZWMVk1cGtidFczOVlzbFF6S0NGcVIwWnZvQUFMOG1hQUhBUGdsYUVJbWdCUUM3OHNVa3lkeEQzOWEremlFdDFTR2tKR2dCQVBBYmdoWUE3STM2RUlKUkh3TEFoMDJkNWVjbVNlWldGZ3hwVFJkbEhxMkJSSTZuQzE0QUFQeUVvQVVBKythcUJaRzRhZ0hBUjNubFQrN3UrclpXV1FBeHVHcEJTcTVhQUFEOGdxQUZBSHUxdWJ3WlA0elltakpCTEMwQ2dQZHF1bUg4ZitUWUFNblkxalVMQ0VYUWdwUUVMUUFBZmtIUUFvQkRjTldDS0U3UDdtK2RQd1hnelpwdU9QSUJOQVZZOTIzOVpORVFodXN5cEhSaCtnQUFQeWRvQWNBaENGb1FpZm9RQU41alhWWFZ3dVRJMkhQZjFzSkVFRWpmMWk1YWtOSzU2UU1BL0p5Z0JRQjd0N204K2FvK2hFQ3VMUU9BdDJpNllUeWRmV1ZvWkU3RkdzVDBZQytrTXYwTUJBREFEd2hhQUhBb3Jsb1FoZm9RQU41cWJXSms3cUZ2YXordlEwenFRMGhKZlFnQXdFOElXZ0J3S1A1d1N5VHFRd0I0bGFZYlBqdWRUUUZjczRDNDFJZVFrb3NXQUFBL0lXZ0J3RUdvRHlFWTlTRUF2SlpyRnVUdXJtOXJIK1JDWFA1OWtwS2dCUURBVHdoYUFIQklybG9ReFZnZjRnOUdBUHhTMHcycnFxcU9UWW1NYlYyemdOajZ0bjZxcXVyWm1ramt0T21HSThNSEFQZ25RUXNBRGtuUWdraGN0UURncDVwdU9QRUJOQVZZOVczOVl0RVFucXNXcE9TUkFnREFEd2hhQUhBd1UzMklsemhFOGRrbUFQaUY4WnJGd29ESTJIUGYxcXB4WUI2KzJ4TUpYUmcrQU1BL0NWb0FjR2l1V2hERnNmb1FBSDZrNllieEE0VXJ3eUZ6cm52QmZMaG9RVXArYndZQStBRkJDd0FPN1l1SkU0Z1BHQUQ0a1pXcGtMbUh2cTI5a0llWjhPK1Z4RnkwQUFENEFVRUxBQTVxYzNtelVSOUNJT3BEQVBpYnBodkdFTjY1cVpBNVlWT1luMGM3STVGRjB3MG5oZzhBOEhlQ0ZnQ2tvRDZFS01iNkVLOXpBUGlYcGh1T1hMT2dBTGQ5V3o5Wk5NeU9xeGFrcEQ0RUFPQi9DRm9Ba0lMNkVDTHhvaE9BLzdjY1EzaW1RY2EyVlZXdExSaG1hV050Sk9TQkFnREEveEMwQU9EZzFJY1FqUG9RQUtycEpQYU5TWkM1WmQvV0w1WU1zK1NpQlNtNWFBRUE4RDhFTFFCSVJYMElVU3pPN20rRkxRRHd5cC9jUGZadDdiSWN6TlJVK2JPMVB4STVOM2dBZ0w4VHRBQWdGWC9rSlJKQkM0Q0NOZDB3bnNQKzVHdUF6QzB0R0diUFZRdVNtWDVlQWdCZ0ltZ0JRQkpUZmNpajZST0VvQVZBMlZ5eklIZmYrcmIyQVMzTTM4WU9TVWg5Q0FEQVh3aGFBSkNTcXhaRW9UNEVvRkJOTjF4WFZYVnEvMlJzNjVvRlpFTmdpcFFFTFFBQS9rTFFBb0NVdnBvK2dRaGFBQlNtNllZajF5d293THB2NnllTGhpeTRhRUZLcWtNQUFQNUMwQUtBWkRhWE4wL3FRd2prNnV6KzlzaENBSXF5R3E4YVdUa1pleFltZ256MGJmM2lkMmdTT3A1Q3FnQUF4YXNFTFFBSVFIMElrYmhxQVZDSXBodE9xcXI2dzc3SjNHcjZZQmJJaDZzV3BLUStCQUJnSW1nQlFHcnFRNGhFMEFLZ0hNS2U1TzZoYjJ0ZjU1Q2Y3M1pLUXVwREFBQW1naFlBSktVK2hHQStxUThCeUYvVERlT0hCT2RXVGVaV0ZneFpjdEdDbEFRdEFBQW1naFlBUk9DbEhaRzRhZ0dRUHo5N2tMdTd2cTI5ZW9jTTlXMHRhRUZLcWtNQUFDYUNGZ0JFb0Q2RVNBUXRBRExXZE1QNHl2L1lqc25ZdHFxcXBRVkQxaDZzbDBRV1RUZWNHRDRBZ0tBRkFBRk05U0grVUVRVTZrTUFNdFYwdzVFUG9DbkF1bS9yRjR1R3JMbFlRMHJxUXdDQTRsV0NGZ0FFNG9RM2tWemJCa0NXMXVOTFRLc2xZODk5VzY4c0dMS25Qb1NVMUljQUFNV3JCQzBBQ0VSOUNKRUlXZ0JrcHVtRzhVT0JLM3NsY3k2MlFCbGN0Q0FsRnkwQWdPSlZnaFlBUkxHNXZCblBHMyt6RUlJNFBidS8xVHNMa0plMWZaSzVoNzZ0aFplaEFGTTkwTE5kazhpcHdRTUFDRm9BRUlzL0RCUEpaOXNBeUVQVERlUDM5SFBySkhPdVdVQloxSWVRVE5NTnJsb0FBTVVUdEFBZ0VrRUxJbEVmQXBDQnBodU9YTE9nQUgvMmJlMURWeWlMK2hCU09qTjlBS0IwZ2hZQWhLRStoR0RVaHdEa1lYemxmMnlYWkd4YlZkWEtncUU0d2xXazVLSUZBRkE4UVFzQW9uSFZna2pVaHdETVdOTU5KK29VS01DcWIrc1hpNGF5OUczdG9nVXB1V2dCQUJSUDBBS0FhQVF0aU1TSGN3RHpOcjd5WDlnaEdYdnUyMW8xRHBUcndlNUo1SGlxWndNQUtKYWdCUUNocUE4aG1PT3orMXN2ZFFCbXFPbUc4YVQxbGQyUnVXc0xocUtwRHlFbDlTRUFRTkVFTFFDSTZJdXRFSWdQTUFEbWFXVnZaTzZiNmdBb25xQUZLWG1VQUFBVVRkQUNnSEEybHpkamZjaldaZ2ppczBVQXpFdlREV05JN3R6YXlKeUtNMERZaXBSY3RBQUFpaVpvQVVCVVgyMkdJTlNIQU16STFCZSt0ak15ZDl1MzlaTWxROW1tN3dNZUtaQ0tVQ3NBVURSQkN3Q2lFclFnRXZVaEFQTXh2dkpmMkJjWjJ3b1RBWC9ocWdYSk5OM2dVUUlBVUN4QkN3QkNVaDlDTU9wREFHYWc2WWFUcXFwdTdJck1MZnUyZnJGa1lMSXhDQklTdEFBQWlpVm9BVUJrcmxvUXhWZ2ZJbXdCRUo5WC91VHVzVy9yTDdZTS9JV0xGcVIwWWZvQVFLa0VMUUNJVE5DQ1NBUXRBQUpydW1IOFEvOG5PeUp6U3dzRy9xcHZhMEVMVW5MUkFnQW9scUFGQUdHcER5RVlRUXVBMkZ5eklIZDNQbEFGZnVMUllFamsxT0FCZ0ZJSldnQVFuYXNXUkxGUUh3SVFVOU1OUzMvb0ozTmorSGhseWNCUGJBeUdWS2FyWWdBQXhSRzBBQ0E2SGRSRUltZ0JFRXpURFVjK2dLWUE2NzZ0bnl3YStBblhia2hKMEFJQUtKS2dCUUNoYlM1dnhqOFlQZHNTUVFoYUFNUXpoaXdXOWtMR25sWGpBTC9ob2dVcG5aaytBRkFpUVFzQTVrQjlDRkdvRHdFSXBPbUdrNnFxL3JBVE1yZnEyL3JGa29HZjZkdDZNMVVNUVFxQ0ZnQkFrUVF0QUpnRDlTRkVjbTBiQUdINEdZSGNQZlJ0N2VzY2VBMVhMVWpsZUFxL0FnQVVSZEFDZ1BBMmx6Y2I5U0VFOHVucy92YklRZ0RTYXJwaHZEQjBiZzFrYm1uQndDdDlOeWdTY3RVQ0FDaU9vQVVBYzZFK2hFalVod0NrdDdZRE1uYzMxUUVBdklidkY2UjBZZm9BUUdrRUxRQ1lDeWVUaVVUUUFpQ2hwaHRXNDVscU95QmpXOWNzZ0RkeTBZS1VYTFFBQUlvamFBSEFMS2dQSVJqMUlRQ0pOTjF3NUFOb0NyRHUyL3JGb29IWG1yNW4rSjJaVk5TNUFRREZFYlFBWUU3VWh4Q0pxeFlBYVl5VklRdXpKMlBQZlZ1dkxCaDRCMWN0U0ticEJsY3RBSUNpQ0ZvQU1DZnFRNGhFMEFMZ3dLWS80RitaTzVsenNRVjRyNDNKa2RDRjRRTUFKUkcwQUdBMnB2cVFSeHNqaUxFKzVNUXlBQTVxYmR4azdxRnZhMWZjZ1BjU3RDQWxGeTBBZ0tJSVdnQXdONjVhRUltckZnQUgwblREdGY1dkNuQnR5Y0I3OVcydE9vU1VCQzBBZ0tJSVdnQXdOMTc0RVlrUFF3QU9vT21HbzZxcVZtWk41djdzMi9ySmtvRVBlakJBRWptZGZtWURBQ2lDb0FVQXM3SzV2SGxTSDBJZ3ArcERBQTVpV1ZYVnNWR1RzYTB3RWJBajZrTkl5VlVMQUtBWWdoWUF6Skg2RUNKUkh3S3dSMDAzbkV4QkM4alpxbS9yRnhzR2RrQjlDQ2xkbUQ0QVVBcEJDd0RtU0gwSWthZ1BBZGl2OFpYL3dvekoyR1BmMW1zTEJuYkVSUXRTY3RFQ0FDaUdvQVVBczZNK2hHRFVod0RzU2RNTjQ2dklLL01sY3k2MkFEdlR0L1g0Ky9LemlaS0lpeFlBUURFRUxRQ1lLL1VoUk9JREVvRDk4TXFmM0gzcjI5cVpmMkRYWExVZ2xjVlUrd1lBa0QxQkN3RG1TbjBJa1h5MkRZRGRhcnBockdZNk5WWXlKNndKN0lPZ0JTbTVhZ0VBRkVIUUFvQlptdXBEdnRrZVFSeWYzZC9xb2dYWWthWWJqbHl6b0FDMzA0bC9nRjF6S1llVS9HNE1BQlJCMEFLQU9YUFZna2l1YlFOZ1o4WlgvZ3ZqSkdOYllTSmdYMVFTa1ppZ0JRQlFCRUVMQU9aTTBJSkkxSWNBN01EVTYzMWpsbVJ1MmJmMWl5VURlL1JvdUNSeWJ2QUFRQWtFTFFDWXJjM2x6WXY2RUFKUkh3S3dHMS9Na2N3OTlHM3Q2eHpZTjFjdFNLYnBoZ3ZUQndCeUoyZ0J3Tnk1YWtFazZrTUFQbUQ2bzd4WGtPUnVaY1BBQVd3TW1ZUThRZ0FBc2lkb0FjRGNDVm9RaWFBRndNZDQ1VS91N3ZxMjlzb2NPQVJCQzFJU3RBQUFzaWRvQWNDc3FROGhtTVhaL2UxblN3RjR1NllibG1NTms5R1JzYTFyRnNDaDlHMjltYjd2UUFxcVF3Q0E3QWxhQUpBRFZ5MklSTkFDNEkyYWJqanlBVFFGV1BkdC9XVFJ3QUc1YWtFcXg5UFBkd0FBMlJLMEFDQUhnaFpFSW1nQjhIWmp5R0poYm1Uc2VReGFXREJ3WUtxS1NFbDlDQUNRTlVFTEFHWnZxZys1czBtQ1VCOEM4QVpOTjR4L2hQL0R6TWpjc20vckYwc0dEa3pRZ3BUVWh3QUFXUk8wQUNBWHJsb1FpYUFGd090NTVVL3VIdnEyOXJNcWtJTHFFRklTdEFBQXNpWm9BVUFXTnBjMzR4K3Z0N1pKRUlJV0FLL1FkTVA0L2ZMY3JNamMwb0tCRktaTE9zK0dUeUtxUXdDQXJBbGFBSkFUTHdXSllxd1B1YllOZ045eXpZTGMzZlZ0N1VVNWtKTDZFRkpaVEJWeEFBQlpFclFBSUNlQ0ZrVGlxZ1hBTHpUZHNLcXE2dGlNeU5qV05Rc2dBR0V2VWhLMEFBQ3lKV2dCUURiVWh4RE1wN1A3MnlOTEFmaW5waHVPZkFCTkFkYlQyWDZBbEZ5MElDVkJDd0FnVzRJV0FPVEdWUXNpY2RVQzRNZkd5cENGMlpDeDU3NnRWeFlNcEthK2lNUXVMQUFBeUpXZ0JRQzVFYlFnRWtFTGdQL1JkTVA0Qi9jcmN5RnoxeFlNQlBKZ0dTUnlhdkFBUUs0RUxRRElpdm9RZ2xFZkF2QlBYdm1UdTRlK3JaM3FCeUp4MVlKa3BwQXRBRUIyQkMwQXlKR3JGa1RpcWdYQXBPbUc4WlgvdVhtUU9kY3NnR2lFdjBqcHpQUUJnQndKV2dDUW83V3RFc2pTTWdEK0ZiSTRjczJDQXZ6WnQvV1RSUVBCdUdoQlNpNWFBQUJaRXJRQUlEdWJ5NXZ4ajBqUE5rc1FwMmYzdHllV0FmQ3Y0Tm14TVpDeHJUQVJFTkVVQVBNN01xbTRhQUVBWkVuUUFvQmNxUThoRXZVaFFOR2Fiamh4NFljQ3JQcTJmckZvSUNoWExVamxlUHBaRUFBZ0s0SVdBT1RxaTgwU2lLNTJvSFJqcmRlaTlDR1F0Y2UrcmRYWEFaRjl0eDBTY3RVQ0FNaU9vQVVBV1ZJZlFqRHFRNEJpTmQwdzluSi84aFZBNWx4c0FhSnowWUtVQkMwQWdPd0lXZ0NRTS9VaFJLSStCQ2lWVi83azdsdmYxbDZLQTZINVBrVmlGeFlBQU9SRzBBS0FuS2tQSVJMMUlVQnhtbTRZdi9lZDJqeVpjODBDbUl0SG15S1JjNE1IQUhJamFBRkF0dFNIRU14WUgrSmNLbENNcGh1T1hMT2dBTGQ5V3o5Wk5EQVRybHFRVE5NTmZoOEdBTElpYUFGQTd0U0hFSW1yRmtCSlZsVlZMV3ljakQwTEV3RXpzN0V3RWhLMEFBQ3lJbWdCUU83ODhadElQdHNHVUlLbUcwNnFxdnJEc3NuY3FtL3JGMHNHWnNSRkMxSzZNSDBBSUNlQ0ZnQmtiWE41ODZTSGxrQ08xWWNBaGZoaTBXVHVvVzlyWCtmQXJFeFZSMXRiSXhHL0N3TUFXUkcwQUtBRS9naE9KT3BEZ0t3MTNUQytWankzWlRLM3NtQmdwdFNIa01xcHlRTUFPUkcwQUtBRVgyMlpRTlNIQUxrVGNDUjNkMzFiTzc4UHpKWHZYeVF6QlhJQkFMSWdhQUZBOXRTSEVJejZFQ0JiVFRjc3grOXpOa3pHdHE1WkFETW5hRUZLZ2hZQVFEWUVMUUFvaGRlMVJMSzBEU0EzVFRjYytRQ2FBcXo3dG42eWFHREdWSWVRa2tjSEFFQTJCQzBBS0lYNkVDSlJId0xrYUYxVjFjSm15ZGh6MzliQ1JNQ3M5VzM5NHVJakNibG9BUUJrUTlBQ2dDS29EeUdZeGRuOXJiQUZrSTJtRzhiWGlWYzJTdVpjcEFKeTRhb0ZxU3lhYmpneGZRQWdCNElXQUpSRWZRaVJDRm9BT1ZuYkpwbDc2TnZhaFRRZ0Y0SVdwS1ErQkFESWdxQUZBQ1h4eDNFaUViUUFzdEIwdy9qOTdOdzJ5WnhyRmtCT3Z0c21DYWtQQVFDeUlHZ0JRREdtK3BCdk5rNFE2a09BWExobVFlN3UrcmIyK2h2SWh1OXBKT2FpQlFDUUJVRUxBRXJqcWdXUkNGb0FzOVowdzZxcXFtTmJKR05iMXl5QVREMVlMSW00aEFZQVpFSFFBb0RTQ0ZvUXlkWFovZTJSalFCejFIVERpUStnS2NDcWIrc1hpd1l5cEQ2RVpKcHVjTlVDQUpnOVFRc0Fpcks1dkhsUkgwSXdybG9BY3pWZXMxallIaGw3N3R0YU5RNlFLL1VocEhSaCtnREEzQWxhQUZBaVZ5MklSTkFDbUoybUc4WS9qbC9aSEptN3RtQWdZeTVha0pLTEZnREE3QWxhQUZBaVFRc2krYVErQkppaGxhV1J1WWUrclgwSUNXUnJxa1Y2dG1FU0ViUUFBR1pQMEFLQTRxZ1BJU0JYTFlEWmFMcGhmT1YvYm1Oa3pqVUxvQVRxUTBqbHRPa0dEdzRBZ0ZrVHRBQ2dWSzVhRUltZ0JUQUwweC9FWGJNZ2QzLzJiZjFreTBBQlhPNGhKVmN0QUlCWkU3UUFvRlNDRmtTaVBnU1lpMlZWVmNlMlJjYTJ3a1JBUVZ5MElLVUwwd2NBNWt6UUFvQWlUZlVoZDdaUElFNlVBNkUxM1hCU1ZkV05MWkc1WmQvV0w1WU1sS0J2YXhjdFNFblFBZ0NZTlVFTEFFcm1xZ1dSQ0ZvQTBhMXRpTXc5OW0zOXhaS0J3anhZT0ltb0RnRUFaazNRQW9CaWJTNXZ2azdub1NHQzA3UDcyeE9iQUNKcXVtRjhjZmpKY3NqYzBvS0JBcWtQSVpYRmRERU5BR0NXQkMwQUtKMnJGa1R5MlRhQW9GeXpJSGZmbk5BSENpVm9RVXJxUXdDQTJSSzBBS0IwZ2haRW9qNEVDS2ZwaHZHVi82bk5rTEd0YXhaQXdZVE1TRWw5Q0FBd1c0SVdBQlJOZlFqQnFBOEJRbW02NGFpcXFwV3RrTGwxMzlaUGxneVVhUHIrNTNkaVVoRzBBQUJtUzlBQ0FGeTFJQmIxSVVBa1k4aGlZU05rN0ZrMURvQ3JGaVJ6YnZRQXdGd0pXZ0NBb0FXeE9GME9oTkIwdzNoaDV3L2JJSE9ydnExZkxCa28zS2IwQVpCTzB3MFh4ZzhBekpHZ0JRREZVeDlDTU1kbjk3Zk9wd0lSZkxFRk12ZlF0N1d2Y3dBWExVakw3NzhBd0N3SldnREF2L2tqTzVGYzJ3YVEwdlN5MENsbmNyZXlZWUIvY2RHQ2xBUXRBSUJaRXJRQWdIOFR0Q0NTejdZQkpPYi9SWEozMTdlMUY5d0FWVlZORlVxUFprRWlxa01BZ0ZrU3RBQ0FmOWVIakM5NG5zMkNJTlNIQU1rMDNUQys4aisyQVRJMlZzWXRMUmpnYjF5MUlKWGpwaHVPVEI4QW1CdEJDd0Q0cjY5bVFTRHFRNENEbS83STdRTm9jcmVlWG04RDhGK3UvSkNTcXhZQXdPd0lXZ0RBZnptVFRpVHFRNEFVMWxWVkxVeWVqRDMzYmIyeVlJQi9jTkdDbEZ4MEJBQm1SOUFDQUNicVF3aG1yQS94cWdjNG1LWWJ4ajl3WDVrNG1YT3hCZUFIK3JZV3RDQWx2L3NDQUxNamFBRUFmNmMraEVqVWh3Q0h0RFp0TXZmUXQ3V2Y5UUIrN3NGc1NNUkZDd0JnZGdRdEFPRHYxSWNRaWZvUTRDQ2FiaGkvMzV5Yk5wbHp6UUxnMTc2YkQ0a3NwdXRxQUFDeklXZ0JBSCtoUG9SZ0ZtZjN0OElXd0Y0MTNYRGttZ1VGK05OWmZJRGY4bjJTbEFRdEFJQlpFYlFBZ0g5eTFZSklCQzJBZlJ0ZitSK2JNaG5iVmxXMXNtQ0EzM0xSZ3BRRUxRQ0FXUkcwQUlCL0VyUWdFa0VMWUcrYWJqaFJwMEFCVm4xYnYxZzB3SzlOM3l0ZGVDU1ZDNU1IQU9aRTBBSUEvc2ZtOHVhcHFxcEhjeUVJOVNIQVBvMnYvQmNtVE1hZSs3WldqUVB3ZXE1YWtNcXB5UU1BY3lKb0FRQS81cW9Ga1Z6YkJyQnJUVGVNcndhdkRKYk0rVDhVNEcwMjVrVXEwOCtuQUFDeklHZ0JBRC8yMVZ3STVOUFovZTJSaFFBN3RqSlFNdmZRdDdXWDJRQnZJMmhCU21lbUR3RE1oYUFGQVB5QStoQUNVaDhDN0V6VERlTXIvM01USlhPdVdRQzhrWUFhaWJsb0FRRE1ocUFGQVB5YytoQWlFYlFBZHFMcGh2RkN6dG8weWR4dDM5WlBsZ3p3TGcvR1JpSXVXZ0FBc3lGb0FRQS9wejZFU05TSEFMdXlyS3BxWVpwa2JDdE1CUEFoNmtOSTVianBoaFBUQndEbVFOQUNBSDVDZlFnQnVXb0JmTWowaCtzYlV5Unp5NzZ0WHl3WjROM1VoNUNTcXhZQXdDd0lXZ0RBcjZrUElSSkJDK0NqdlBJbmQ0OTlXL3Y1RGVCalhMUWdKVUVMQUdBV0JDMEE0TmY4b1o1SXh2b1FaMVNCZDJtNjRXTDhQbUo2Wkc1cHdRQWYwN2YxMDFUREJDbGNtRG9BTUFlQ0ZnRHdDNXZMbS9IczlEY3pJaEJYTFlEM0VoNGtkM2Q5V3p0M0Q3QWJ2cCtTeXJuSkF3QnpJR2dCQUwvMzFZd0k1Tm95Z0xkcXVtRjg1WDlzY0dSc2ZIbTlzbUNBblZFZlFqSk5ONmdQQVFEQ0U3UUFnTjhUdENDU1UvVWh3RnMwM1hEa0EyZ0tzSjVPM1FPd0d5NWFrSktnQlFBUW5xQUZBUHlHK2hBQ1VoOEN2TVVZc2xpWUdCbDdIb01XRmd5d082cVlTT3pDQWdDQTZBUXRBT0IxWExVZ0V2VWh3S3MwM1RCZXdQbkR0TWpjcW0vckYwc0cyTGxISXlVUkZ5MEFnUEFFTFFEZ2RRUXRpRVI5Q1BCYVgweUt6RDMwYmUzckhHQS9OdVpLSXFkVC9SMEFRRmlDRmdEd0N1cERDTWhWQytDWG1tNFlhNGJPVFluTUxTMFlZRy9VaDVDU3F4WUFRR2lDRmdEd2VxNWFFSW1nQmZBN2F4TWljM2Q5VzN0dERiQS92c2VTMG9YcEF3Q1JDVm9Bd091TlFZdXRlUkhFOGRuOXJSYyt3QTgxM2JBYXYwK1lEaG5idW1ZQnNGOVRtTTN2d0tUaTkxMEFJRFJCQ3dCNHBhayt4RlVMSW5IVkF2aUhxYy9hQjlEa2J0MjM5WXN0QSt5ZHF4YWs0cUlGQUJDYW9BVUF2STJnQlpGOHRnM2dCOGJLa0lYQmtMSG52cTFYRmd4d0VOK05tVVFXVFRlY0dENEFFSldnQlFDOHdlYnlSbjBJa2FnUEFmNm02WWJ4ZThLVnFaQTVGMXNBRHNkRkMxTHkreTRBRUphZ0JRQzhuYXNXUktJK0JQaXJ0V21RdVdJd3JuNEFBQ0FBU1VSQlZJZStyZjBzQm5BNExscVFrdm9RQUNBc1FRc0FlRHQvM0NjU1FRdmdYNXB1R0w4Zm5Kc0dtZlAvSHNBQjlXMzlNbFkybVRtSnVHZ0JBSVFsYUFFQWI2UStoR0FXWi9lM255MEZ5dFowdzFGVlZhdlM1MEQyL3V6YitzbWFBUTdPVlF0U0VTSUdBTUlTdEFDQTkzSFZna2dFTFlCbFZWWEh4VStCbkcyRmlRQ1MyUmc5cVRUZDRLb0ZBQkNTb0FVQXZJK2dCWkVJV2tEQm1tNDRtWUlXa0xQVmRMNGVnTU56MFlLVUxrd2ZBSWhJMEFJQTNrRjlDTUdvRDRHeWphLzhGNlVQZ2F3OTkyMjl0bUtBTlBxMmR0R0NsRnkwQUFCQ0VyUUFnUGY3WW5ZRUltZ0JCV3E2WVh6aGQyWDNaTzdhZ2dHU2U3QUNFbkhSQWdBSVNkQUNBTjVQMElKSUJDMmdURjc1azd0dmZWczdXUStRbnFzV3BITGNkTU9SNlFNQTBRaGFBTUE3YlM1dnhqODBQWnNmUVl6MUlWNzhRa0dhYmhqL3paL2FPWmxiV2pCQUNFSnZwS1ErQkFBSVI5QUNBRDdtcS9rUmlLc1dVSWpwVlo5ckZ1VHV0bS9ySjFzR0NNRkZDMUpTSHdJQWhDTm9BUUFmb3o2RVNENmQzZDg2cVFwbEdGLzVMK3lhakcyRmlRRGltSUp2TGpxU2lxQUZBQkNPb0FVQWZJRDZFQUp5MVFJeTEzVERTVlZWTi9aTTVwWjlXNzlZTWtBb3JscVFpdW9RQUNBY1FRc0ErRGoxSVVRaWFBSDVjMDJKM0QzMmJlM3JIQ0FlUVF0U1dVeGhZd0NBTUFRdEFPRGpmQkJBSk9wRElHTk5ONHhuazgvdG1Nd3RMUmdncE8vV1FrTHFRd0NBVUFRdEFPQ0QxSWNRa0tzV2tDL2hQbkozMTdlMUQvSUFBdkw5bWNUVWh3QUFvUWhhQU1CdXJNMlJRSzR0QS9MVGRNUDR5di9ZYXNuWXRxcXFsUVVEaFBab1BTUWlhQUVBaENKb0FRQzc4ZFVjQ2VUODdQNVdmeTFrcE9tR0l4OUFVNEIxMzlaUEZnMFFtcXNXcEtJK0R3QUlSZEFDQUhaZ2Mzbno1R1VQd2FnUGdieU1JWXVGblpLeFp4ZkNBR1poWTAyazBuVERoZUVEQUZFSVdnREE3dWpOSnhMMUlaQ0pwaHZHTThsLzJDZVpXL1p0L1dMSkFPRUpXcENTK2hBQUlBeEJDd0RZSGZVaFJIS3FQZ1N5NFpVL3VYdm8yOXJQVVFBejBMZjFHTFRZMmhXSnVHZ0JBSVFoYUFFQU82SStoSURVaDhETU5kM3dXUjgxQlZoYU1zQ3N1R3BCS2k1YUFBQmhDRm9Bd0c2cER5RVM5U0V3ZjY1WmtMdTc2WFUwQVBQeDNhNUk1TGpwaGlQREJ3QWlFTFFBZ04xeTlwcEl4dm9RTDM1Z3BwcHVXSTEvVExZL01yWjF6UUpnbGdRdFNFbDlDQUFRZ3FBRkFPeVEraEFDY3RVQ1ptaDZxZWNEYUhLMzd0djZ4WllCWnNjbElsTHltQUFBQ0VIUUFnQjJ6NWwzSXZsc0d6Qkw0LzhsQzZzalk4OTlXNjhzR0dCK3BwQ2NCd2FrNHFJRkFCQ0NvQVVBN0o3NkVDSTVWaDhDODlKMHcvakg0eXRySTNNdUxnSE1tNnNXcE9MM1d3QWdCRUVMQU5peHplWE4rTHJubTdrU2lBK3pZRjY4OGlkM0QzMWI2L2NIbURkQkMxSlpOTjBnYkFFQUpDZG9BUUQ3NGFvRmthZ1BnWmxvdW1FTVJwM2JGNWtUQUFTWVA0RTVVaEswQUFDU0U3UUFnUDBRdENBUzlTRXdBMDAzSExsbVFRSCs3TnY2eWFJQjVxMXZheGN0U01udnR3QkFjb0lXQUxBSDZrTUlhR2twRU43NDcvVFltc2pZVnBnSUlDc1Axa2tpRndZUEFLUW1hQUVBKytPcUJaR29ENEhBbW00NHFhcnF4bzdJM0twdjZ4ZExCc2lHK2hCU09UVjVBQ0ExUVFzQTJCOUJDeUpabk4zZkNsdEFYR3U3SVhPUGZWdjdPZ2ZJaS9vUWttbTZ3VlVMQUNBcFFRc0EyQlAxSVFRa2FBRUJUWDhrL21RM1pFNkZGVUIrQkMxSVNkQUNBRWhLMEFJQTlzdFZDeUlSdElDWXZQSW5kOS82dG5aZUhpQXpmVnMvVlZYMWJLOGtjbWJ3QUVCS2doWUFzRWVieTVzdlZWVnR6WmdnMUlkQU1FMDNYT3VZcGdDdVdRRGt5MVVMVWhHMEFBQ1NFclFBZ1AxejFZSklCQzBnaUtZYmpseXpvQUMzMDR0bkFQTGtZaEdwSERmZGNHTDZBRUFxZ2hZQXNIK0NGa1R5K2V6KzlzaEdJSVRWZUduR0tzallWcGdJSUhzdVdwQ1NxeFlBUURLQ0ZnQ3daNXZMbTYvcVF3aGs0YW9GcERlOXZ2dkRLc2pjc20vckYwc0d5RmZmMWk1YWtKS2dCUUNRaktBRkFCeUdxeFpFSW1nQjZYMnhBekwzMExlMXIzT0FNanphTTRsY0dEd0FrSXFnQlFBY2hxQUZrWHhTSHdMcE5OMHcva0g0M0FySTNNcUNBWXJocWdXcCtKa2FBRWhHMEFJQURrQjlDQUc1YWdIcGVPVlA3dTZja2djb3lzYTZTYVhwQnZVaEFFQVNnaFlBY0RpdVdoQ0pvQVVrMEhURHNxcXFZN01uWTF2WExBQ0tJMXhIU3VwREFJQWtCQzBBNEhBRUxZaEVmUWdjV05NTlJ6NkFwZ0RydnEyZkxCcWdITlAzZlJjY1NjVkZDd0FnQ1VFTEFEaVFxVDdrMmJ3SnhGVUxPS3gxVlZVTE15ZGp6MzFiQ3hNQmxFbDlDS2tJV2dBQVNRaGFBTUJodVdwQkpFdmJnTU9ZdXFPdmpKdk0rWDhGb0Z6cVEwamxkTG9jQndCd1VJSVdBSEJZWDh5YlFFN1A3bTlQTEFRT1ltM01aTzZoYjJ1QlVvQnlDVnFRa3FzV0FNREJDVm9Bd0FGdExtODI2a01JUm4wSTdGblREZU8vczNOekpuT3VXUUNVVFhVSUtWMllQZ0J3YUlJV0FIQjRYbnNTeWJWdHdONjVaa0h1N3ZxMjlnRWJRTUg2dG42cHF1clIxd0NKdUdnQkFCeWNvQVVBSEo3NkVDSlJId0o3MUhURHFxcXFZek1tWTF2WExBQ1lDTjJSaW9zV0FNREJDVm9Bd0lHcER5RWc5U0d3QjAwM25QZ0FtZ0tzcGxmTUFQQzkrQW1ReW1MNjJSc0E0R0FFTFFBZ0RmVWhST0tEWU5pUDhackZ3bXpKMkhQZjFxcHhBUGgvTGxxUWt2b1FBT0NnQkMwQUlBMzFJVVJ5ZkhaLzY0OVNzRU5OTjR6bmk2L01sTXhkV3pBQS82OXZhMEVMVWxJZkFnQWNsS0FGQUNTZ1BvU0FmRmdHdTdVeVR6TDMwTGUxRS9FQS9LOEhFeUVSandjQWdJTVN0QUNBZEp6YUpwTFB0Z0c3MFhUREdGdzZOMDR5SjZBSHdJOEk0WkdLbjc4QmdJTVN0QUNBZEw2YVBZR29ENEVkYUxyaFNKQ09BdnpadC9XVFJRUHdBK3BEU0dhcTd3TUFPQWhCQ3dCSVpITjVNMzVBOFdqK0JPSjFNbnpjc3FxcWhUbVNzYTFxSEFCK3dVVUxVdko0QUFBNEdFRUxBRWpyaS9rVGlQb1ErSUNtRzA2cXFyb3hRekszN052NnhaSUIrSkhwLzRobnd5RVJRUXNBNEdBRUxRQWdMZlVoUkRMV2h6aTFDdStuTW9UY1BmWnRMU1FLd08rb0R5RVZ2ODhDQUFjamFBRUFDYWtQSVNEMUlmQU9VeC8wSjdNamMwc0xCdUFWMUllUXluSFREVWVtRHdBY2dxQUZBS1RuWlNpUnFBK0I5M0hOZ3R4OTY5dmFCMmNBdklhTEZxU2tQZ1FBT0FoQkN3QklUMzBJa1N6TzdtK0ZMZUFObW00WVgvbWZtaG1aYzgwQ2dGY1J6Q014OVNFQXdFRUlXZ0JBWWxOOXlJTTlFSWlnQmJ6U2RKcDRaVjVrN3JadjZ5ZExCdUFOL0k1TEtvSVdBTUJCQ0ZvQVFBenFRNGhFMEFKZWJ3eFpMTXlMakQycnhnSGdIZFNIa0lycUVBRGdJQVF0QUNBRzlTRkVvajRFWHFIcGhwT3FxdjR3S3pLMzZ0djZ4WklCZUNOQkMxSlpURCtuQXdEc2xhQUZBQVN3dWJ3WlA4RDRaaGNFSW1nQnYrY2FFYmw3Nk52YTF6a0E3L0hkMUVoSWZRZ0FzSGVDRmdBUWg2c1dSSEoxZG45N1pDUHdZMDAzakgrOFBUY2VNcmV5WUFEZW8yL3JwNnFxdG9aSEl1cERBSUM5RTdRQWdEZ0VMWWpHVlF2NE9hLzh5ZDFkMzlaZUl3UHdFZjRmSVJVWExRQ0F2Uk8wQUlBZzFJY1FrS0FGL0VEVERlTXIvMk96SVdQakMrU2xCUVB3UVJzREpKRlRnd2NBOWszUUFnQmljZFdDU0Q2cEQ0Ry9hN3JoeUFmUUZHRGR0L1dMUlFQd1FTNWFrTXhVOVFjQXNEZUNGZ0FRaTZBRjBiaHFBWCszcnFwcVlTWms3TGx2NjVVRkEvQlJLcWhJN013Q0FJQjlFclFBZ0VEVWh4Q1FvQVZNbW00WS8xaDdaUjVrenNVV0FIYnAwVFJKeEVVTEFHQ3ZCQzBBSUo0dmRrSWdZMzNJaVlYQXY2eU5nY3c5OUczdHVoWUF1N1F4VFJKeDBRSUEyQ3RCQ3dBSVpuTjVNMzdBc2JVWEFuSFZndUkxM1REK096Z3ZmUTVrenpVTEFIWk5mUWlwSERmZGNHVDZBTUMrQ0ZvQVFFeGVreExKdFcxUXN1a1B0SzVaa0xzLys3YjI2aGlBWGZOL0N5bXBEd0VBOWtiUUFnQmlFclFna2xQMUlSUnVmT1YvWFBvUXlOcDRTV3RseFFEczJoVGljN0dSVk5TSEFBQjdJMmdCQUFHcER5RWc5U0VVcWVtR0UzVUtGR0RWdC9XTFJRT3dKNjVha0lxTEZnREEzZ2hhQUVCY3Jsb1FpZm9RU2pXKzhsL1lQaGw3N3R0YU5RNEErL1RkZEVuRVJRc0FZRzhFTFFBZ0xrRUxJbEVmUW5HYWJoaGZ3RjNaUEprVHBBTmczMXkwSUpWRjB3M0NGZ0RBWGdoYUFFQlE2a01JeUlkeGxNWXJmM0wzMExlMVY4WUE3SnYvYTBoSjBBSUEyQXRCQ3dDSXpWVUxJaEcwb0JoTk40eGY3NmMyVHVaOFh3ZGc3L3EyZmhtcnFreWFSQzRNSGdEWUIwRUxBSWp0aS8wUXlQSFovYTNYUUdTdjZZWWoxeXdvd0czZjFrOFdEY0NCdUdwQktuNkhCUUQyUXRBQ0FBTGJYTjU4OS9LSFlMeCtwZ1RMc2MvWnBzbllWcGdJZ0FQYkdEaUp1RklIQU95Rm9BVUF4S2MraEVnKzJ3WTVhN3JocEtxcUcwc21jOHZwakRzQUhJcWdCY2swM2FBK0JBRFlPVUVMQUloUGZRaVJxQThoZDE3NWs3dkh2cTM5YkFIQVFmVnRyVHFFbEFRdEFJQ2RFN1FBZ09BMmx6Y2I5U0VFb3o2RUxFMHYzVDdaTHBsYldqQUFpVHdZUElsNExBQUE3SnlnQlFETWcvb1FJbEVmUXE2ODhpZDMzN3dvQmlBaDlTR2tJbWdCQU95Y29BVUF6SU1QLzRoa3JBOFJ0aUFyVFRlTXIveVBiWldNYlYyekFDQXhZVDlTT1c2NjRjVDBBWUJkRXJRQWdCbFFIMEpBZ2haa28rbUdvNnFxVmpaSzV0WjlXejlaTWdBSnVXaEJTcTVhQUFBN0pXZ0JBUE9oUG9SSUJDM0l5Uml5V05nb0dSdkRtbXNMQmlDbEtmQzN0UVFTRWJRQUFIWkswQUlBNWtOOUNKRXMxSWVRZyttRThCK1dTZVpXZlZ1L1dESUFBYWdQSVpVTGt3Y0Fka25RQWdCbVlxb1BlYlF2QWhHMElBZENiT1R1b1c5clgrY0FSS0UraEZUT1RSNEEyQ1ZCQ3dDWUZ4K1VFSW1nQmJQV2RNTm5mM0NsQUN0TEJpQVFGeTFJcHVrRzlTRUF3TTRJV2dEQXZIeTFMd0laNjBPdUxZUVpXMXNlbWJ2cjI5b0hXZ0NFNGY4bEVsTWZBZ0RzaktBRkFNekk1dkxtU1gwSXdiaHF3U3cxM1RDKzhqKzJQVEsycmFwcWFjRUFCT1IzV2xKeDBRSUEyQmxCQ3dDWUgvVWhSUExwN1A3MnlFYVlrNlliam53QVRRSFdmVnUvV0RRQUFibHFRU3FDRmdEQXpnaGFBTUQ4cUE4aEdsY3RtSnV4TW1SaGEyVHN1Vy9ybFFVREVOVEdZa2prZEFwZEF3QjhtS0FGQU15TStoQUNFclJnTnBwdUdGK3hYZGtZbVhPeEJZRElCQzFJeVZVTEFHQW5CQzBBWUo3VWh4Q0oraERtWkcxYlpPNmhiMnZYcndBSXEyL3JNV2l4dFNFU3VUQjRBR0FYQkMwQVlKNThnRUkwcmxvUVh0TU4xMVZWbmRzVW1idTJZQUJtd0ZVTFVuSFJBZ0RZQ1VFTEFKaWhxVDdrd2U0SXhBZDdoRFoxTWE5c2ljejkyYmYxa3lVRE1BUGZMWWxFWExRQUFIWkMwQUlBNWt0OUNKR2NuOTNmbnRnSWdTMnJxanEySURLMkZTWUNZRVlFTFVobDBYU0QzMTBCZ0E4VHRBQ0ErVklmUWpUcVF3aHAra1BxMG5iSTNLcHY2eGRMQm1BbVZJZVFrcXNXQU1DSENWb0F3RXh0TG0vR0QxTysyUitCcUE4aHF2WDRjczEyeU5oejM5WnJDd1pnTHFadzRMT0ZrY2lad1FNQUh5Vm9BUUR6NXFvRmtaeXFEeUdhcGh2RzEycWZMSWJNQ2JvQk1FZnFRMGhGMEFJQStEQkJDd0NZTjBFTG9sRWZRalJlK1pPN2IzMWIrNkFLZ0RsU0gwSXE1eVlQQUh5VW9BVUF6Smo2RUFMeXFwb3dtbTRZdng1UGJZVE1MUzBZZ0prU0ZDU1o2ZklkQU1DN0NWb0F3UHk1YWtFazZrTUlvZW1HSTljc0tNQnQzOVpQRmczQUhQVnQ3YUlGS2FrUEFRQStSTkFDQU9aUDBJSm92SzRtZ3ZIcmNHRVRaR3dyVEFSQUJoNHNrVVFFTFFDQUR4RzBBSUNabStwRDd1eVJRRDViQmlrMTNUQmVWYm14QkRLMzdOdjZ4WklCbURsWExVaEZkUWdBOENHQ0ZnQ1FCMWN0aU9UNDdQN1c2eUJTK21MNlpPNnhiMnRmNXdEazRMc3Rrc2p4VkRjSUFQQXVnaFlBa0lITjVjM1g2WVE0UkhGdEU2VFFkTVA0TXUzYzhNbWNpaVlBY3VHaUJTbDVJQUFBdkp1Z0JRRGt3MVVMSWxFZlFpcGUrWk83dTc2dHZmNEZJQXQ5V3o5VlZmVnNteVNpUGdRQWVEZEJDd0RJaDZBRmthZ1A0ZUNhYmhoZitSK2JQQmticjFldExCaUF6TGhxUVNxQ0ZnREF1d2xhQUVBbTFJY1FrUG9RRG1icVYvWUJOTGxiVHk5L0FTQW5naGFrNG5FQUFQQnVnaFlBa0JkWExZaEUwSUpER2tNV0N4TW5ZK05aOWJVRkE1QWhsVmlrc21pNlFkZ0NBSGdYUVFzQXlJdWdCWkVzenU1dlA5c0kremI5Y2ZRUGd5WnpxNzZ0WHl3WmdOejBiUzFvUVVxQ0ZnREF1d2hhQUVCRzFJY1FrS0FGaCtDVlA3bDc2TnY2aXkwRGtMRkh5eVVSUVFzQTRGMEVMUUFnUDY1YUVJbWdCWHZWZE1QNE5YWnV5bVJ1YWNFQVpNNVZDMUs1TUhrQTREMEVMUUFnUDE1MkU0bjZFUGJOOXp4eWQ5ZTM5Y2FXQWNpYy8rdEk1ZFRrQVlEM0VMUUFnTXhzTG0vR1AxQTkyeXVCQ0Zxd0YwMDNyS3FxT2paZE1yWjF6UUtBUXJob1FUSk5ON2hxQVFDOG1hQUZBT1JKZlFpUmZENjd2ejJ5RVhhcDZZWVRIMEJUZ0hYZjFpOFdEVUR1K3JaK21nS0drTUtacVFNQWJ5Vm9BUUI1K21LdkJMSncxWUk5V0UxZlc1Q3I1NzZ0VjdZTFFFSFVoNUNLaXhZQXdKc0pXZ0JBaHRTSEVKQ2dCVHN6bmZhOU1sRXlkMjNCQUJSR2ZRaXB1R2dCQUx5Wm9BVUE1RXQ5Q0pGOFVoL0NEbm5sVCs0ZStyYjJZUk1BcGZGL0g2a2NOOTNnOTFVQTRFMEVMUUFnWCtwRGlNWlZDejZzNllieGxmKzVTWkk1MXl3QUtKSHFFRkpTSHdJQXZJbWdCUUJrU24wSUFRbGE4Q0hUS3pQWExNamRuMzFiUDlreUFLWHAyL3FscXFwSGl5Y1I5U0VBd0pzSVdnQkEzdFNIRUluNkVENXFPWjcxTlVVeXRoVW1BcUJ3cmxxUWlvc1dBTUNiQ0ZvQVFON1c5a3N3cmxyd0xrMDNuRlJWZFdONlpHNDF2ZVlGZ0ZJSldwQ0tla0lBNEUwRUxRQWdZNXZMbXllblZ3bG1hU0c4aytBWXVYdnMyOXJYT1FDbCsxNzZBRWluNlFiMUlRREFxd2xhQUVEK3Z0Z3hnWnllM2QrZVdBaHYwWFREZU1iM2s2R1JPVUUwQUlyWHQ3V0xGcVFrYUFFQXZKcWdCUURrNzZzZEU0ejZFTjdLSzM5eTk2MXZheTk0QWVEZkhzeUJSQzRNSGdCNExVRUxBTWljK2hBQ3VyWVVYcXZwaHZIcjVkVEF5SnhyRmdEd1g4S0hwT0tpQlFEd2FvSVdBRkFHOVNGRW9qNkVWMm02NGNnMUN3cHcyN2YxazBVRHdIK29EeUVWQVc4QTROVUVMUUNnRE9wRGlFWjlDSyt4cXFwcVlWSmtiQ3RNQkFEL0lHaEJNazAzcUE4QkFGNUYwQUlBQ3FBK2hJRFVoL0JMVFRlTVYwLytNQ1V5dCt6YitzV1NBZUMvcGt0UHowWkNJb0lXQU1DckNGb0FRRG5VaHhESldCK2kvNVpmOFQyTDNEMzBiZTNySEFCK3pGVUxVdkY3S2dEd0tvSVdBRkFPOVNGRTQ2b0ZQelNkNnowM0hUSzNzbUFBK0tudlJrTWlnaFlBd0tzSVdnQkFJYWI2a0cvMlRTQ2ZMWU9mOE1xZjNOMzFiZTBESkFENE9SY3RTT1Y0cWpFRUFQZ2xRUXNBS0l1ckZrUnlyRDZFLzlWMHczTDgyakFZTXJaMXpRSUFmazBna2NUOG5nb0EvSmFnQlFDVVJkQ0NhTlNIOEI5Tk54ejVBSm9DclB1MmZySm9BUGl0QnlNaWtRdURCd0IrUjlBQ0FBcXl1Yng1VVI5Q01PcEQrS3QxVlZVTEV5Rmp6OVBYT1FEd2UrcERTTVZGQ3dEZ3R3UXRBS0E4cmxvUXlWZ2Y0clVRNHpXTDhZK1pWeVpCNXBaOVc3OVlNZ0M4aXFBRnFaeWJQQUR3TzRJV0FGQWVRUXVpVVI5QzVaVS9CWGpvMjlyL3dRRHdldC9OaWxTbUlEZ0F3RThKV2dCQVlkU0hFSkQ2a01JMTNmRFpxekVLc0xSa0FIaTl2cTJmcXFyYUdobUp1THdJQVB5U29BVUFsTW1MV2lKWm5OM2ZDbHNVcXVtR0k5Y3NLTUJkMzliT253UEEyN2xxUVNvdVdnQUF2eVJvQVFCbEVyUWdHa0dMY28ydi9JOUxId0paMjdwbUFRRHZKcWhJS29JV0FNQXZDVm9BUUlHbStwQTd1eWNRUVlzQ05kMXc0Z05vQ3JEcTIvckZvZ0hnWFZ5MElKWFQ2Zm9lQU1BUENWb0FRTGxjdFNBUzlTRmxXbzI3TDMwSVpPMjViMnZWT0FEd2ZpNWFrSktyRmdEQVR3bGFBRUNoTnBjM1g2ZHo1aENGb0VWQm1tNjRxS3JxcXZRNWtMMXJLd2FBOTV1dVFqMGFJWWxjR0R3QThET0NGZ0JRTmxjdGlPVHE3UDdXYWRaeXJFb2ZBTmw3Nk52YXVYTUErRGhYTFVqRlJRc0E0S2NFTFFDZ2JJSVcvQjk3OTVNVFNaTHRDOWp5NmM3aHJnQjZCVW12QUhyaTArVEpKVmZNTW1vRkZiV0NDRlp3eVJVMHpGSXV1UjVNZmRMSkN0cFpRY0VLTHF3Z25yemJzNXFxeWorUVJJU2JtMzJmaEhyU1VrV2NFOGtmdDUrZEV4dFRMVEpRMW0xL3kvODQ5enFRUE5Nc0FHQXpCQmNaaTRrV0FNQlhDVm9BUU1hc0R5RkNnaGFKSyt1Mm4xcHlubnNkU042SHBpcnV0QmtBTnNKRUM4YXlWOWJ0b2VvREFGOGlhQUVBbUdwQlRONVpINUs4UmYvQU12Y2lrTFJIcTNFQVlIT2FxaEMwWUV5bVdnQUFYeVJvQVFBSVdoQWJVeTBTTmR3R1crWmVCNUszYUtyaVFac0JZS051bEpPUkhDazhBUEFsZ2hZQWtEbnJRNGlRb0VXNnJBd2hkYmROVlZ6b01nQnMzQ2NsWlNTQ0ZnREFGd2xhQUFEQlZBc2lZMzFJZ3NxNjdVZnV2c3U5RGlSdm9jVUFzQlhXaHpDV1k1VUhBTDVFMEFJQUNHNlpFNkc1cGlUSDl4bFNkOTFVaGR1MkFMQWRmc1l5bWlFMERnRHdPNElXQUVDL1BxUy9IWFN2RWtSRTBDSWhaZDMydC96ZjVsNEhrbWVhQlFCc1NWTVZELzVtWlVUV2h3QUFmeUpvQVFCOFpuMElNWGw3OVBIc1VFZW1yNnpiZmczTUt2YzZrTHl6cGlydXRCa0F0c3I2RU1ZaWFBRUEvSW1nQlFEdzJZVktFSmxURFVsQ0g3TFl5NzBJSk8zZWFod0EyQW5yUXhpTDFTRUF3SjhJV2dBQS8ySjlDQkd5UG1UaXlycnRwNUw4bkhzZFNONXFHR2NPQUd5WGlSYU01V0NZMUFjQThCdEJDd0RnS2V0RGlJbjFJZE5uVWc2cHUybXF3dWNjQUhhZ3FRb1RMUmlUcVJZQXdPOElXZ0FBVHprc0lqYW1Xa3hVV2JmOWc4amozT3RBOGxaYURBQTdkYVBjak9SSTRRR0Fwd1F0QUlEZldCOUNoQVF0cGt0d2k5UmR1bGtMQUR0bmZRaGpNZEVDQVBnZFFRc0E0SThjamhLVE85Mlluckp1KzF2K0I3blhnYVE5bW1ZQkFLTVFjbVFzSmxvQUFMOGphQUVBL0pHZ0JiSG9EekpQZFdOYXlycmREeUVzY3E4RHlUdHZxa0lRREFCMnowUUx4ckpYMXEyd0JRRHdHMEVMQU9CM3V0bXlQemk2VlJVaWNOck5sZzhhTVRubi9VUEkzSXRBMHU2YnFqRE5BZ0JHTUFRZEg5V2VrUWhhQUFDL0ViUUFBTDdFVkF2R2R0Yk5sc1lDVDh4d3crdDk3blVnZVNhMkFNQzQvSjNBV0FRdEFJRGZDRm9BQUY5eXBTcU02S2FiTGQwV242YnozQXRBOG02YXF2QXpFZ0RHWlgwSVl6bFJlUURnTTBFTEFPQlByQTloUlAwWTRGTU5tSjZ5YnVjaGhPUGM2MER5VExNQWdQR1phTUZZM3FvOEFQQ1pvQVVBOERYV2h6Q0cwMjYyZkZENWFTbnJkaitFWUFvSnFidHNxc0lOV2dBWVdWTVZnaGFNcHF4YlV5MEFnSDhSdEFBQXZzWm9kSGJ0ckpzdFBUU2RwdjZXLzBIdVJTQnBqNlpaQUVCVVRHQmtMRWNxRHdBRVFRc0E0R3VzRDJISGJyclowa1NFQ1Nycjl0QUJOQmxZTlZWaDJnNEF4TU9VS2NaaW9nVUE4QytDRmdEQXQxZ2Z3aTcwTjhWUFZYcXkrb0RNWHU1RklHbjNUVldjYXpFQVJNVWtQTVppb2dVQThDK0NGZ0RBdHdoYXNBdW4zV3pwcHZnRURmdUozK2RlQjVJMzEySUFpSTZKRm96bFlKanFCd0JrVHRBQ0FQaXE0ZkQ3V29YWW9yTnV0blFiYmJyYzhpZDFOMDFWK0I0RkFKRnBxcUliSnVQQkdFeTFBQUFFTFFDQTc3cFNJcmJrcHBzdFY0bzdUV1hkOXJmODMrWmVCNUpubWdVQXhNdFVDOFlpYUFFQUNGb0FBTjhsYU1FMjlMZlBUbFYybXNxNjNUZk5nZ3ljTlZWeHA5RUFFQzFUcHhqTGljb0RBSUlXQU1BM1dSL0NscHdPbnkybWFSRkMyTk03RXZZb1RBUUEwVFBSZ3JFY3F6d0FJR2dCQUR5SHFSWnMwb2R1dG5UN2JLTEt1ajBNSVN4enJ3UEpXelJWSVF3R0FISHpOd1dqS2V2VytoQUF5SnlnQlFEd0hJSVdiTXB0TjFzdVZIUFMzUEluZGJkTlZWem9NZ0RFYlFoRjNtc1RJeEcwQUlETUNWb0FBTjlsZlFnYjBvL2lQMVhNNlNycnR0OUYvQzczT3BBOFlUQUFtQTVUTFJqTGljb0RRTjRFTFFDQTV6TFZndGVhZDdQbG5TcE9tbHYrcE82NnFRb0hOZ0F3SFoxZU1SSVRMUUFnYzRJV0FNQnpYUTBUQ2VCSGZPaG1TMkdkQ1N2cnRyL2xmNUI3SFVqYW8ya1dBREE1QXBLTTVhM0tBMERlQkMwQWdHY1oxb2M0S09kSDNIYXpwY1BMQ1N2cmRqK0VzTXE5RGlUdnZLa0tVM2NBWUVLYXFqRFJndEVNcXhVQmdFd0pXZ0FBTHlGb3dVdjFOOFJQVlczeStwREZYdTVGSUduM2ZkQkNpd0Zna202MGpaRUlXZ0JBeGdRdEFJQm5HMVkvV0IvQ1M4eTcyZElOOFFrcjYvWXdoUEJ6N25VZ2VhdW1LaDYwR1FBbXlWUUx4bktrOGdDUUwwRUxBT0NsVExYZ3VUNE00UnltN1VML1NOeE5VeFUrNXdBd1haLzBqcEdZYUFFQUdSTzBBQUJleXNFNXozSGJ6WllMbFpxMnNtNzd0Uy9IdWRlQjVLMjBHQUFtelVRTHhySTNUQUFFQURJa2FBRUF2SWoxSVR4RC8vazRWYWdrbk9kZUFKSjMyVlNGVzdBQU1HRk5WZlNyQ3UvMWtKRllId0lBbVJLMEFBQitoS2tXZk11OG15M3ZWR2pheXJydGIva2Y1RjRIa3RhSHdremVBWUEwbUdyQldLd1BBWUJNQ1ZvQUFEOUMwSUt2K1RCTVBXSEN5cnJkZHdCTkJzNmJxbmpRYUFCSWdxQUZZekhSQWdBeUpXZ0JBTHlZOVNGOHhXMDNXenFjVDBPL01tUXY5eUtRdFB1bUtsWmFEQURKc0FxTXNSeXJQQURrU2RBQ0FQaFJGeXJIRTMzdzVsUkJwcStzMjM3MDdmdmM2MER5aE1JQUlDRk5WUWhhTUpxeWJrMjFBSUFNQ1ZvQUFEOUswSUtuNXQxc2VhY2lTWERMbjlUZE5GVmh4UkVBcE9kV1R4bkppY0lEUUg0RUxRQ0FIOUxObHYwTzNIdlZJNFR3WVZnbnc4U1ZkVHMzK3BZTW1HWUJBR2t5MVlLeG1HZ0JBQmtTdEFBQVhzUGhPcmZkYk9uUU1nRmwzZTZiWmtFR1BqUlYwV2swQUNUSnozakdJbWdCQUJrU3RBQUFYc1A2a0x3OWhoQk9jeTlDUXZyQXpFSHVSU0JwajhKRUFKQTBRUXZHOG5ZSXJnTUFHUkcwQUFCK21QVWgyWnQzcytWZDdrVklRVm0zaDlZcGtJRlZVeFVQR2cwQWFScW1WajFxTHlNeDFRSUFNaU5vQVFDOGx2VWhlZnJRelpaNm40N3pFTUplN2tVZ2FmZE5WWnhyTVFBa3oxUUx4bktpOGdDUUYwRUxBT0MxckEvSnoyMDNXNXAra0lpeWJ2c0hndTl5cndQSm0yc3hBR1Roa3pZekVrRUxBTWlNb0FVQThDcldoMlNuSDhWN21uc1JFdU9XUDZtN2JxckNvUXNBNU1IUGZNWmlkUWdBWkViUUFnRFlCRk10OGpIdlpzdTczSXVRaXJKdSsxditiM092QThremdRY0E4bUYxQ0dQWksrdjJVUFVCSUIrQ0ZnREFKZ2hhNU9GRE4xdGU1VjZFVkpSMXUyK2FCUms0YTZwQ09Bd0FNdEZVeFlPSmk0ekkraEFBeUlpZ0JRRHdhc09FZzF1VlROcHROMXU2Rlo2V3ZwOTd1UmVCcEQwS0V3RkFscXdQWVN6V2h3QkFSZ1F0QUlCTk1kVWlYZjFoNVdudVJVakpNTkoybVhzZFNONWl1TlVLQU9URitoREdJbWdCQUJrUnRBQUFOc1ZLaVhUTmg2a2xwRU13aXRUZE5sWGhjdzRBZVRMUmdyRWNxendBNUVQUUFnRFlDT3REa3ZXaG15MkZhQkpTMXUySkI0Qmt3S29qQU1oVVV4VW11MVNVN1FBQUlBQkpSRUZVV2pDYTRlOHRBQ0FEZ2hZQXdDYTVQWnlXMjI2MmRGaVpIdjlPU2QxbFV4VnVzZ0pBM201eUx3Q2pzVDRFQURJaGFBRUFiSkxKQitsNERDR2M1bDZFMUpSMTJ3ZG5Ebkt2QTBucnYzZXR0QmdBc2lkMHlWaE10QUNBVEFoYUFBQWJZMzFJVXVaRFAwbEVXYmY3RHFESndIbFRGYjUzQVFEV2h6QVdFeTBBSUJPQ0ZnREFwcDJyNk9SOTZHWkwwMG5TMC8vYjNNdTlDQ1R0M3M4Z0FHQWdhTUZZRG9hUU93Q1FPRUVMQUdEVEhOQlAyMjAzV3k1eUwwSnF5cnJ0YjFXOXo3ME9KRy9WVk1XRE5nTUF3NFNyKyt3THdWaXNEd0dBREFoYUFBQWIxYzJXL1NIWHRhcE8wbU1JNFRUM0lpVEtMWDlTZDlOVXhZVXVBd0JQbUdyQldLd1BBWUFNQ0ZvQUFOdGdxc1UwemJ2WjhpNzNJcVNtck5zK1BIT2NleDFJbmtrOEFNQWZmVklSUm1LaUJRQmtRTkFDQU5nR1FZdnArZERObHZxV0p0TXNTTjFsVXhWdXJBSUFmK1QzQThaaW9nVUFaRURRQWdEWU9PdERKdWUybXkzZEJrOVFXYmVyRU1KQjduVWdhWSttV1FBQVg5SlVoWWtXakdXdnJGdGhDd0JJbktBRkFMQXRwaU5NUTM5SWVacDdFVkpVMXUyaEEyZ3ljTjVVeFlOR0F3QmZjYXN3akVUUUFnQVNKMmdCQUd5TG9NVTB6THZaOGk3M0lpU3FuMmF4bDNzUlNOcDlVeFVyTFFZQXZzRlVDOFlpYUFFQWlSTzBBQUMyd3ZxUVNmalF6WllDTVFrcTYvWWtoUEErOXpxUVBCTmJBSUR2NlZTSWtad29QQUNrVGRBQ0FOZ21oL2p4dXUxbVM0ZVU2WExMbjlUZE5GWGhad3dBOEQyQ0ZvemxyY29EUU5vRUxRQ0FyZWxteTRzUXdxTUtSNmZ2eVdudVJVaFZXYmZ6RU1KeDduVWdlWE10QmdDK3BLemIvZjUzNHJKdSsxRG1QeFdKc1F5VEJnR0FSUDJYeGdJQVczWmxoVUYwNXQxc2VaZDdFVkxVUDFRMnpZSU1mR2lxd3Zjd0FPQTN3Ky9CcDhQWE81VWhFbjNRNHBObUFFQ2FCQzBBZ0cwVHRJakxoMjYyTkc0L1hmMDZtSVBjaTBEU0hvV0pBSUFnWE1FMEhPa1RBS1RyelhxOTFsNEFZS3VPUHA0OWhCRDJWSGwwdDkxczZVRlBvc3E2UFF3aC9KcDdIVWplTDAxVm5Hc3pBT1JKdUlLSnVXK3E0bERUQUNCTi8wZGZBWUFkTUVGaGZQMHQ4SG51UlVpY3cyZFNkeXRrQVFCNUt1djJ0S3piL3UvSy93MGgvRjNJZ29rNEdBTHhBRUNDckE0QkFIYkIrcER4TGJyWnNzdTlDS2txNi9iRXcyWXlzTkJrQU1oSEg2NTRNcjNDaEVTbXFwOHFlYWQ3QUpBZVFRc0FZT3U2MmZMcTZPUFpvNGRqbzduc1pzdUxUTjk3THR6eUozWFhUVlY4MG1VQVNKdHdCUWs2TXVVVEFOSWthQUVBN0lxcEZ1TzRkUXM4YldYZDlpdGgzdVplQjVMbit4Z0FKRXE0Z3NTZGFEQUFwRW5RQWdEWUZVR0wzZXVuaU15NzJmSWh0emVlaTdKdTkwMnpJQU5uVFZVWXR3d0FDUkd1SUNQSG1nMEFhWHF6WHErMUZnRFlpYU9QWi8xQjJZRnE3OHhQVm9ha3JhemJQbVR4Yys1MUlHbDlZT3l3cVFxQk1RQ1lPT0VLTXZiWHBpbzZId0FBU0l1SkZnREFMbDA1Rk42WlN5R0x0SlYxZStqZkV4bFlDRmtBd0hRSlY4Qy9ISVVRQkMwQUlER0NGZ0RBTGwwNEdONkoyLzV3TW9QM21UdEJHbEozMDFTRnp6a0FUSXh3QmZ6SmliL2ZBQ0E5Z2hZQXdNNTBzMlYzOVBIczN2cVFyZXJIN00rNzJkSU44SVNWZFh0aTF5OFpXR2t5QUV4RFdiZjlqZjM1RUs3dzl4NzgzcEY2QUVCNkJDMEFnRjJ6UG1TN0ZuMmdKZVUzeUwrNERVWHFMcHVxK0tUTEFCQXY0UXA0dHJkbDNlNWJpUWNBYVJHMEFBQjJ6ZnFRN2Juc1prc0g4SWtyNjNibFFUYUplelROQWdEaUpGd0JQNnovdHlOSURBQUpFYlFBQUhiSytwQ3R1ZTJuV1NUNjNoajB0NkQwbVF5Y04xVnhwOUVBRUFmaEN0aUlFMEVMQUVpTG9BVUFNQWJyUXphcnYvMDk3MlpMWTBqVGR4NUMyTXU5Q0NUdGZ2aWNBd0FqRXE2QWpUdFNVZ0JJaTZBRkFEQUc2ME0yYTlGUENrbnBEZkZudzhQdTkwcEQ0aFoyVndQQU9JUXJZS3RPbEJjQTB2Sm12VjVyS1FDd2MwY2Z6KzQ4dk51SXkyNjJuQ2Z3UHZpT3NtNzdNYlBINmtUQ2JwcXE4QUFhQUhaSXVBSjI2aTlXNUFGQU9reTBBQURHMG8rRy94L1ZmNVhiL3ZiM2hGOC96MVRXN2FtUUJSbncvUXdBZHFDczI4TWhYREVYcm9DZDZvTk5naFlBa0FoQkN3QmdMRmVDRnEveTJEOFk3V1pMSS9ZVFY5YnQvaEJNZ3BSZE5sVmhCUklBYk1rUXJqZ2R3aFZ2MVJsR2NUSThDd0VBRWlCb0FRQ01vcHN0NzQ0K250MTZ5UGZERnQxczZWQXlEd3MzRFVuY28ya1dBTEI1d2hVUW5TTXRBWUIwQ0ZvQUFHTzZNTlhpaDF4MnMrWEZCRjgzTHpROEhIY0FUZXJPbTZvd25RY0FOa0M0QXFKbUhTUUFKRVRRQWdBWWsvVWhMM2ZyNEQwcnF4RENYdTVGSUduM1RWV3N0QmdBZnB4d0JVeEhXYmRIVnVZQlFCb0VMUUNBMFZnZjhtTDllUDE1TjF1NitaMkJzbTc3L2IzdmM2OER5WnRyTVFDOG5IQUZURmIvZDU2Z0JRQWtRTkFDQUJpYjlTSFB0K2htU3c5azh1R1dQNm03YWFyaWt5NER3UE1JVjBBU2pyUVJBTklnYUFFQWpNMzZrT2U1N0diTGl5bThVRjZ2ck51NS9iMWt3RFFMQVBnTzRRcEl6b21XQWtBYTNxelhhNjBFQUVaMTlQR3M4OUR3bS9yMUtpZFdodVNock52OUVNSmRDR0V2OTFxUXRBOU5WU3kwR0FEK2JQaDk4SFQ0ZXFkRWtKei9icXJDMy9jQU1IRW1XZ0FBTVRnUElmeGRKNzdvc2IrOUptU1JsWVdRQllsN3RCb0hBSDVQdUFLeTBxOFBzVUlQQUNaTzBBSUFpTUdWb01WWExiclpzb3YwdGJGaHcyam9wYnFTdUpVYmZBQWdYQUVaT3hHMEFJRHBFN1FBQUViWFQyczQrbmgyN2VIaW4xeDJzK1ZGWksrSjdUcFhYeEozMjFTRnp6a0EyUkt1QUlhZ0JRQXdjWUlXQUVBc3JqeG8vSjNiWVlVRW1TanI5c1MvQVRMZyt4b0EyUkd1QVA3Z1NFRUFZUG9FTFFDQVdGZ2Y4aCtQSVlSNVAra2psaGZFVHJqbFQrcXVtNm93SWhtQUxBaFhBTit3MTYrTmJLcmlUcEVBWUxvRUxRQ0FLRmdmOGp1TGJyYnNJbm85YkZsWnQvMHQvN2ZxVE9KTXN3QWdhY0lWd0F2MEV3MnRDZ1dBQ1JPMEFBQmlZbjFJQ0pmZGJPbGhTMGFHQi9LcjNPdEE4czdjMkFNZ1ZXWGQ5c0dLdWI5bGdCZXdQZ1FBSms3UUFnQ0lTZTdyUTI3ZCtNNVNIN0xZeTcwSUpPM2VhaHdBVWpPRUt6NS8rVjBPZUNsQkN3Q1l1RGZyOVZvUEFZQm9ISDA4eTNXcXhXTS9PdFRLa0x6MGUzbERDTC9tWGdlUzkxTlRGU2IxQURCNXdoWEFKalZWOFVaQkFXQzZUTFFBQUdKemtXblFZaUZra1NXSHo2VHVSc2dDZ0NrVHJnQzJwYXpiazZZcVBpa3dBRXlUb0FVQUVKVnV0cnc2K25qMm1ObER6TXR1dG5RUW1abmhvZjF4N25VZ2VTc3RCbUJxaEN1QUhlblhod2hhQU1CRUNWb0FBREhxMTRlOHo2UXp0LzAwaXdoZUI3dDNydVlrN3RJTlBRQ21RcmdDR01HSnZ3c0JZTG9FTFFDQUdPVVN0T2duZDh5NzJmSWhndGZDRHBWMTI5L3lQMUJ6RXZab21nVUFzUk91QUVaMnBBRUFNRjF2MXV1MTlnRUEwVG42ZVBhUXdjUE9uNndNeVU5WnQvc2hoRHNQODBuY1dWTVZnaFlBUktlczIvNEcrVnk0QW9qRWZ6ZFY0ZklGQUV5UWlSWUFRS3hTbjJweEtXU1JyWE1QOVVuY3ZaQUZBREVwNi9ib1NiakNWREVnSmlmRDh3OEFZR0lFTFFDQVdLVWN0TGdOSVN3aWVCM3MyUENRUDRlMU9PVE45emNBUmlkY0FVekVrYUFGQUV5VG9BVUFFS1Z1dHJ3NituajJtT0ROLy80OXpidlowbWpRUEozblhnQ1NkOU5VaFFmRkFJeEN1QUtZb0JOTkE0QnBFclFBQUdLVzRsU0xSVGRiZGhHOERuYXNyTnYrb2YreHVwTTQweXdBMkNuaENtRGlqalFRQUticHpYcTkxam9BSUVwSEg4LzZteDMvU0tnN2w5MXNPWS9nZGJCalpkM3U5NE5hUFB3bmNaZE5WZmdlQjhEV0NWY0FpZmxyVXhVdVpBREF4SmhvQVFCRXE1c3RQeDE5UEx0UDVPSHByWnZlV1ZzNEJDQnhqNzdIQWJCTndoVkF3bzZHWUQ0QU1DR0NGZ0JBN1ByMUlUOVB2RXY5QWVTOG15MGZJbmd0N0ZoWnQ0Y09vTW5BcXFrSzMrTUEyQ2poQ2lBVC9UVFBDODBHZ0drUnRBQUFZbmVSUU5CaTBjMldicWZrYXhWQzJNdTlDQ1R0dnFtS2N5MEdZQk9Ha09xcGlXQkFSbzQwR3dDbVI5QUNBSWhhSDFDWStQcVE2MjYyZERNbFUyWGQ5amVUM3VkZUI1STMxMklBWHVOSnVLTC9tZkpXTVlITStMNEhBQk1rYUFFQVRNRlUxNGZjTzRETW5sditwTzZtcVlwUHVnekFTd2xYQVB4SEg5TDNlelVBVEl1Z0JRQXdCVk5kSDNMYXpaWVBFYndPUmxEV3JVTURjaUJNQnNDekNWY0FmRlUvRFZIUUFnQW1STkFDQUlqZVJOZUgvTksvN2doZUJ5TW82M2JmTkFzeThLR3BpanVOQnVCYmhDc0FudVZJbVFCZ1dnUXRBSUNwbU5MNmtPdHV0blRJbnJkRkNHRXY5eUtRdE1jUXdrcUxBZmdTNFFxQUZ4TzBBSUNKRWJRQUFLWmlLdXREN28zU3o5dHdzTERNdlE0a2I5RlVoZFZJQVB4R3VBTGdWUTc2NzZNbXhnSEFkQWhhQUFDVE1Ld1B1WjNBUTl2VGJyWjArSmczMDB4STNXMVRGUmU2RE1Dd0xtMHVYQUd3RWYxVUMwRUxBSmdJUVFzQVlFcjZnNzMvaWZqMS90SUhRaUo0SFl5a3JOdVRFTUk3OVNkeEN3MEd5TmNRcmpnZHZ2emVBN0E1UjhQYVZBQmdBZ1F0QUlBcHVZbzRhSEhkelpZbUdlQ1dQNm03YnFyaWt5NEQ1RVc0QW1BblRwUVpBS1pEMEFJQW1JeHV0cnlMZEgzSS9UQXVtWXlWZGR2ZjhqL3dHU0JoajZaWkFPUkR1QUpnNTQ2VkhBQ21ROUFDQUppYUdOZUhuSGF6NVVNRXI0T1JEQWNSSy9VbmNlZE5WZGdaRFpBdzRRcUFjWlYxZTlSVWhaV2tBREFCZ2hZQXdOVEV0ajdrbDI2MjlCQ0VQbVN4bDMwVlNGay91Y2Q2SklBRUNWY0FSS1ZmSCtJWkF3Qk1nS0FGQURBcGthMFB1ZTVtU3dlUG1ldHZISVVRZnM2OURpUnYxVlNGeVQwQWlSQ3VBSWpXa2RZQXdEUUlXZ0FBVXhURCtwRCtkdmQ4bXVWanc0UnRTTjFOVXhVWHVnd3diY0lWQUpNZ2FBRUFFL0YvTkFvQW1LQ3JDRjd5YVRkYnV0MmR1Ykp1KzRPSzQ5enJRUEpXV2d3d1hmM3ZLMlhkOW9HNS93MGgvRjNJQWlCcWI0ZGdIQUFRT1JNdEFJREpHZGFIM0l4NHdQMUxOMXZhbVVvd3pZSU1YRFpWOFVtakFhWmxDSU4rL3RyVFBvQko2YWRhK0IwY0FDSW5hQUVBVE5YRlNFR0w2MjYyZExoT2Y0RFIzL0kvVUFrUzloaENXR2d3d0RRSVZ3QWs0MFRRQWdEaUoyZ0JBRXpWMVRENmVKZnVRd2h6bnhpR1VhNE9vRW5kZVZNVlZpUUJSRXk0QWlCSlI5b0tBUEVUdEFBQUpxbWJMUitPUHA1ZDczakg5R24vMy9XSllWZ1o0akNEbE4wM1ZiSFNZWUQ0Q0ZjQUpPOUVpd0VnZm9JV0FNQ1VYZTB3YVBGTE4xdDJQaTJVZGRzLzlIcWZmU0ZJbllrdEFCRVJyZ0RJeWw1WnQ0ZE5WZHhwT3dERVM5QUNBSml5WGEwUHVlNW15M09mRkFadStaTzZtNllxcm5RWllGekNGUUJaNjllSENGb0FRTVFFTFFDQXlkclIrcEQ3RU1MY3A0VHc3d09QL3JOd3JCZ2t6alFMZ0pFTWs3UG13aFVBMlRzWkxwY0FBSkVTdEFBQXBtN2I2ME5PKzBDSFR3bGwzZTZiWmtFR1BqUlZZVTBTd0E2VmRYdjBKRnh4b1BZQURCTXRBSUNJQ1ZvQUFGTzN6ZlVodjNTenBRTkhQbHM0L0NCeGo4SkVBTHNoWEFIQWQ1aWtDQUNSZTdOZXIvVUlBSmkwbzQ5bjI1aHFjZDNObHFjK0dZUi9INFljOXR0cWpQQW1jYjgwVlhHdXlRRGJJVndCd0F2OXJhbUtUNG9HQUhFeTBRSUFTTUhGaG9NVzk4TkRjUGpzWE1pQ3hOMExXUUJzbm5BRkFLL1Evd3dSdEFDQVNBbGFBQUNUMTgyV1YwY2Z6eDQzZUJCKzJzMldEejRaaEg4ZmtKeHNZV0lLeEVhNERHQkRoQ3NBMkpBamhRU0FlQWxhQUFDcDZOZUh2Ti9BZS9tbG15MDdud3FlY011ZjFOMFlTUXp3T3NJVkFHekJpYUlDUUx3RUxRQ0FWR3dpYUhIZHpaWU8xZmxOV2JmOWdjbGJGU0Z4cGxrQS9BRGhDZ0MyN0tDczIvMm1La3pjQklBSUNWb0FBRW5Zd1BxUWU0ZU5QTlUvMERMTmdneWNOVlZ4cDlFQXoxUFc3ZUVRckZnSVZ3Q3dBMzJvei9RNUFJaVFvQVVBa0pMWFRMVTQ3V1pMdDBSNGF2V0s0QTVNd2FNd0VjRDNQUWxYbUhRRndLNmRDRm9BUUp3RUxRQ0FsUHhvME9LWGJyYnNmQkw0YkRoUStWbEJTTnpDR0dLQUx4T3VBQ0FTSnhvQkFIRVN0QUFBa3ZHRDYwT3V1OW5Talc3KzZFSkZTTnh0VXhVKzV3QlBDRmNBRUtFalRRR0FPQWxhQUFDcGVjbFVpL3ZoUVRyOHBxemIvc2JRc1lxUXVJVUdBd2hYQUJDOXZmNW5WVk1WZDFvRkFIRVJ0QUFBVW5QK2dxREZhVGRiR3B2UEg3bmxUK291bTZxdzV4bklsbkFGQUJOejR1OVVBSWlQb0FVQWtKUnV0dXlPUHA3MWt5b092dk8rZnVuL3Y3clBVMlhkTHA3eDJZRXA2OWNyclhRUXlJMXdCUUFUWm4wSUFFUkkwQUlBU0ZHL1B1VG5iN3l2NjI2MlBOZDVuaXJyZHQ4Qk5CazROM1lZeU1Yd3MvMTBXSmNrWEFIQVZKM29IQURFUjlBQ0FFalJ4VGVDRnZmRFRVYjRvejU4czZjcUpPeCsrSndESk90SnVLTC9lcWZUQUNSQVdCQUFJdlJtdlY3ckN3Q1FuS09QWjNkZldRSHhWeXREK0tPeWJ2dFJyUDlVR0JMM1UxTVZkanNEeVJHdUFDQURmMnVxNHBOR0EwQThUTFFBQUZMMXBmVWh2d2haOEJWdStaTzZHeUVMSUNYQ0ZRQmtwcjhjSUdnQkFCRVJ0QUFBVXZYSDlTSFgzV3pwTUowL0tldTJQNkE1VmhrU3Q5QmdZT3FFS3dESTJJa0xBZ0FRRjBFTEFDQkovZVNLbzQ5bjk4UDZrUDUvNXpyTlYzaFlSZW91bTZvd3pRZVlKT0VLQVBpWEkyVUFnTGdJV2dBQUtmdThQdVMwbXkwZmRKby9LdXQyTllSeElGV1BwbGtBVXlOY0FRQi9jdEQvZkd5cXdyTU5BSWlFb0FVQWtMSitVa0hYVDdmUVpmNm9yTnREQjlCazROekRXR0FLaENzQTRMdE9oZ3NsQUVBRUJDMEFnR1IxcytWZENPRkNoL21LZnByRm51S1FzUHVtS2xZYURNU3NyTnZUSndFTFA1Y0I0T3VPQkMwQUlCNkNGZ0FBWktlczIvNG0wSHVkSjNFbXRnQlJFcTRBZ0I5eW9td0FFQTlCQ3dBQWN1U1dQNm03YWFyQ2JUY2dHc0lWQVBCcVIwb0lBUEY0czE2dnRRTUFnR3lVZFRzUElmeGR4MG5jWDVxcXVOTmtZRXpDRlFDd2NYOXRxcUpUVmdBWW40a1dBQUJrbzZ6YmZkTXN5TUFISVF0Z0xNSVZBTEJWL1ZRTFFRc0FpSUNnQlFBQU9WbUVFQTUwbklROUNoTUJ1eVpjQVFBN2N4SkN1RkJ1QUJpZm9BVUFBRmtvNi9Zd2hMRFViUkszYXFyaVFaT0JiUk91QUlCUkhDazdBTVJCMEFJQWdGeWM2elNKdTIrcXd1Y2MySnF5Ym8rRzZWRENGUUF3anJmcURnQnhFTFFBQUNCNVpkMzI0MVhmNlRTSm0yc3dzR2xEdUdJK2hDdXMzd0tBY2QzM1A1dWJxdWowQVFER0pXZ0JBRUFPM1BJbmRkZE5WWHpTWldBVGhDc0FJQ3IzSVlTckVNS0ZnQVVBeEVQUUFnQ0FwSlYxdXpCZWxRd3NOQmw0RGVFS0FJaUtjQVVBUkU3UUFnQ0FaSlYxdXg5Q1dPa3dpVHRycXVKT2s0R1hFcTRBZ0tnSVZ3REFoQWhhQUFDUXNqNWtzYWZESk96UmFoemdKWVFyQUNBcWowL0NGVllCQXNDRXZGbXYxL29GQUVCeXlybzlEQ0g4cXJNazdxZW1LaTQwR2ZnVzRRb0FpTXJuY01WVlV4VlhXZ01BMDJTaUJRQUFxWEw0VE9wdWhTeUFyeEd1QUlDb0NGY0FRR0lFTFFBQVNFNVp0eWNoaEdPZEpYRUxEUWFlR3FZNW5RNEJpN2VLQXdDakVxNEFnSVFKV2dBQWtDSzMvRW5kcFIzT1FCQ3VBSURZQ0ZjQVFDWUVMUUFBU0VwWnR5c2owa2xjLy9CMnBjbVFMK0VLQUlpS2NBVUFaRWpRQWdDQVpKUjF1MitkQWhrNGI2cmlUcU1oTDhJVkFCQVY0UW9BeUp5Z0JRQUFLVGtQSWV6cEtBbTdIejduUUFhRUt3QWdPdGY5cWtyaENnRGd6WHE5enI0SUFBQk1YMW0zUnlHRWYyb2xpZnUvSHVwQzJvUXJBQ0E2MTArbVZ6eG9Ed0FRVExRQUFDQWhidm1UdWhzaEMwaVRjQVVBUkVlNEFnRDRKa0VMQUFBbXI2emIvbkRxV0NkSjNFS0RJUjFsM2U0UDRZcUZjQVVBUkVHNEFnQjROa0VMQUFBbWJUaW9NczJDMUYwMlZkSHBNa3piazNCRi8vVk9Pd0ZnZE1JVkFNQVBFYlFBQUdEcStwdkFCN3BJd2g1TnM0RHBFcTRBZ09nSVZ3QUFyeVpvQVFEQVpBMDc3UjFBazdwekQ0QmhXb1FyQUNBNndoVUF3RVlKV2dBQU1HV3JFTUtlRHBLdys2WXFWaG9NOFJPdUFJRG9DRmNBQUZzamFBRUF3Q1NWZFhzU1FuaXZleVJ1cnNFUUwrRUtBSWpPYlFqaFlnaFgzR2tQQUxBdGdoWUFBRXlWVy82azdxYXBpays2REhFUnJnQ0E2QWhYQUFBN0oyZ0JBTURrbEhYYjMvSS8xamtTWjVvRlJHVDQyU05jQVFCeEVLNEFBRVlsYUFFQXdLUU1ONG5QZFkzRWZmREFHTVpYMXUzcGsra1ZlMW9DQUtNU3JnQUFvaUZvQVFEQTFDd2NkcEc0UjZ0eFlEekNGUUFRRmVFS0FDQktiOWJydGM0QUFEQUpaZDBlaGhCKzFTMFM5MHRURmFhMndBNEpWd0JBVklRckFJRG9tV2dCQU1DVU9Id21kYmRDRnJBYndoVUFFQlhoQ2dCZ1VnUXRBQUNZaExKdVQwSUk3M1NMeEMwMEdMWkh1QUlBb25MZkJ5djZRTDF3QlFBd05ZSVdBQUJNeFlWT2tianJwaW8rYVRKc2xuQUZBRVRsYzdqaW9xbUtUbXNBZ0trU3RBQUFJSHBsM2ZhMy9BOTBpc1NaWmdFYklsd0JBRkVScmdBQWtpTm9BUUJBMU1xNjNROGhySFNKeEowWmx3eXZVOWJ0VVFoaFBud0pWd0RBdUlRckFJQ2tDVm9BQUJDN2xRTXpFdmZZNzZYV1pIaTVKK0dLVTVPUEFHQjB3aFVBUURiZXJOZHIzUVlBSUVwbDNSNkdFSDdWSFJMM1UxTVZGNW9NenlOY0FRQlJFYTRBQUxKa29nVUFBREZ6K0V6cWJvUXM0UHVFS3dBZ0tzSVZBRUQyQkMwQUFJaFNXYmY5WWRxeDdwQzRsUWJEbHdsWEFFQlVIb2NndkhBRkFKQzlJR2dCQUVERXpqV0h4RjAyVmZGSmsrRS9oQ3NBSUNxUHcrU0txNllxcnJRR0FPQS9CQzBBQUloT1diY3JCMndrN3RFMEMvZzM0UW9BaUlwd0JRREFNd2hhQUFBUWxiSnU5ME1JQzEwaGNlZE5WZHhwTXJrcTYvWncrRjR2WEFFQTR4T3VBQUI0SVVFTEFBQmkwNjhNMmRNVkVuYmZWSVZwRm1SbkNGZWNEdE1yM3ZvRUFNQ29oQ3NBQUY1QjBBSUFnR2dNNCtQZjZ3aUpNN0dGYkFoWEFFQlVoQ3NBQURaRTBBSUFnSmljNndhSnUvRlFtOVFKVndCQVZJUXJBQUMyUU5BQ0FJQW9sSFhiSDhnZDZ3YUpNODJDSkFsWEFFQlVoQ3NBQUxaTTBBSUFnTkdWZGJzZlFsanBCSW03YktxaTAyUlNJVndCQU5HNWZoS3dlTkFlQUlEdEViUUFBQ0FHL1MzL0E1MGdZWSttV1pBQzRRb0FpSTV3QlFEQUNBUXRBQUFZMVhCbzV3Q2ExSzA4K0dhcWhDc0FJRHJDRlFBQUl4TzBBQUJnYlAzS2tEMWRJR0gzVFZXY2F6QlRNcXgwK2h5dU9OWThBQmlkY0FVQVFFUUVMUUFBR0UxWnR5Y2hoUGM2UU9MbUdzd1VQQWxYOUYvdk5BMEFSaWRjQVFBUUtVRUxBQURHNUpZL3FidHBxdUtUTGhNcjRRb0FpSTV3QlFEQUJBaGFBQUF3aXJKdTdmb25CNlpaRUIzaENnQ0lqbkFGQU1ERUNGb0FBTEJ6d3lHZmFSYWs3a05URlhlNlRBeUVLd0FnT3JmRDMwVENGUUFBRXlSb0FRREFHQlloaEQyVkoyR1BJWVNWQmpNbTRRb0FpRTRmcnJnWXdoVUN1UUFBRS9abXZWN3JId0FBTzFQVzdXRUk0VmNWSjNFL05WVnhvY25zbW5BRkFFUkh1QUlBSUVFbVdnQUFzR3NPbjBuZHJaQUZ1MWJXN1Z5NEFnQ2lJVndCQUpBNFFRc0FBSGFtck51VEVNS3hpcE80aFFhekMyWGRuajZaWG1FZEV3Q01TN2dDQUNBamdoWUFBT3lTVy82azdycXBpays2ekxZSVZ3QkFWSVFyQUFBeUpXZ0JBTUJPbEhYYjMvSS9VRzBTWjVvRkd5ZGNBUUJSRWE0QUFFRFFBZ0NBN1N2cmRqK0VzRkpxRW5mbVlUdWJJbHdCQUZHNUR5R2NDMWNBQVBDWm9BVUFBTHV3Y2xCSTRqNC9mSWNmSmx3QkFGSHBmNys3NnFkWE5GWFJhUTBBQUUrOVdhL1hDZ0lBd05hVWRYc1VRdmluQ3BPNG41cXF1TkJrWGtxNEFnQ2lJbHdCQU1Dem1HZ0JBTUMydWVWUDZtNkVMSGdKNFFvQWlJcHdCUUFBTHlab0FRREExZ3lIaWNjcVRPSldHc3ozRE5OOTVzT1hjQVVBakV1NEFnQ0FWeEcwQUFCZ20weXpJSFdYVFZWODBtVys1RW00b2crZEhTZ1NBSXhLdUFJQWdJMFJ0QUFBWUN2S3VsMDVXQ1J4ajZaWjhFZkNGUUFRRmVFS0FBQzJRdEFDQUlDTksrdDJQNFN3VUZrU2Q5NVV4WjBtSTF3QkFGRVJyZ0FBWU9zRUxRQUEySVorWmNpZXlwS3crNllxVExQSW1IQUZBRVRsOFVtNHdsbzNBQUMyVHRBQ0FJQ05LdXYySklUd1hsVkpuSWt0R1JLdUFJQ29mQTVYWERWVmNhVTFBQURza3FBRkFBQ2I1cFkvcWJ2eE1EOGZ3aFVBRUJYaENnQUFvaUJvQVFEQXhwUjEyeDlHSHFzb2lUUE5JbkZsM1I0T2ZSYXVBSUR4Q1ZjQUFCQWRRUXNBQURhaXJOdDkweXpJd0llbUtqcU5UczhRcmpnZHBsZTh6YjBlQURBeTRRb0FBS0ltYUFFQXdLWXMzUHdtY1kvQ1JHa1JyZ0NBcUFoWEFBQXdHVy9XNjdWdUFRRHdLc05oNWErcVNPSithYXJpWEpPblRiZ0NBS0lpWEFFQXdDU1phQUVBd0NZNGZDWjE5MElXMHlWY0FRRFJ1UTRoWEFoWEFBQXdWWUlXQUFDOFNsbTNKeUdFZDZwSTR1WWFQQzNDRlFBUW5lc24weXNldEFjQWdDa1R0QUFBNExYYzhpZDFOMDFWZk5MbCtBbFhBRUIwaENzQUFFaVNvQVVBQUQrc3JGdUhtZVRBTkl1SWxYVzcveVJjY1p4N1BRQWdBc0lWQUFBa1Q5QUNBSUFmTWh4dW1tWkI2czZhcXJqVDViZzhDVmVjV2wwRUFGRVFyZ0FBSUN1Q0ZnQUEvS2hWQ0dGUDlVallvekJSUElRckFDQTZ3aFVBQUdUcnpYcTkxbjBBQUY2a3JOdkRFTUt2cWtiaWZtcXE0a0tUeHlOY0FRRFJFYTRBQUNCN3dVUUxBQUIra01OblVuY3JaREVPNFFvQWlJNXdCUUFBL0lHZ0JRQUFMMUxXN1VrSTRWalZTTnhDZzNkSHVBSUFvbk03aEt2N2NNV2Q5Z0FBd084SldnQUE4Rkp1K1pPNjY2WXFQdW55OXBWMU94ZXVBSUJvQ0ZjQUFNQXpDVm9BQVBCc1pkMzJ0L3dQVkl5RVBacG1zVjFsM1o0K21WNnhsL0o3QllBSkVLNEFBSUFmSUdnQkFNQ3pES1A5VjZwRjRzNGRNbXllY0FVQVJFVzRBZ0FBWGtuUUFnQ0E1enAzUUVyaTdvZlBPUnNnWEFFQVVSR3VBQUNBRFJLMEFBRGd1OHE2UFFvaHZGY3BFcmRxcXVKQmszK2NjQVVBUkVXNEFnQUF0a1RRQWdDQTUzRExuOVRkTkZWeG9jc3ZKMXdCQUZFUnJnQUFnQjBRdEFBQTRKdUdROVJqVlNKeEt3MStQdUVLQUloS3YvN3NxZzlIQzFjQUFNQnVDRm9BQVBBOXBsbVF1c3VtS2o3cDhyY05LNFRtdzVkd0JRQ002M080NHFLcGlrNHZBQUJndHdRdEFBRDRxckp1KzF2K0J5cEV3aDVEQ0FzTi9ySW40WXBUM3dzQVlIVENGUUFBRUFsQkN3QUF2cWlzMjBNSDBHU2dIN0g5b05IL0lWd0JBRkVScmdBQWdBZ0pXZ0FBOERVcjZ3RkkzSDFURlN0TkZxNEFnTWdJVndBQVFPUUVMUUFBK0pPeWJrOUNDTzlWaHNSbFBiRkZ1QUlBb2lKY0FRQUFFeUpvQVFEQWw3amxUK3B1bXFxNHlxM0x3aFVBRUJYaENnQUFtQ2hCQ3dBQWZxZXMyLzRROWxoVlNOdzhsd2FYZFhzNFRPOFFyZ0NBOFQzMndRcmhDZ0FBbURaQkN3QUFmbFBXN1g0STRWeEZTTnlIcGlydVVuNkxRN2ppZEFpVXZJM2dKUUZBemg2SHlSVlhPVTdVQWdDQUZBbGFBQUR3VkgvcmZVOUZTTmhqcXF0eGhDc0FJQ3JDRlFBQWtMQTM2L1ZhZndFQStIeEkrNnRLa0xoZm1xcElabXFMY0FVQVJFVzRBZ0FBTW1HaUJRQUFuMWtaUXVydVV3aFpDRmNBUUZTRUt3QUFJRU9DRmdBQTlBZTNKeUdFZHlxVzFJVm9BQUFnQUVsRVFWUkI0dVpUZlh2Q0ZRQVFGZUVLQUFESW5LQUZBQURCTkFzeWNOMVV4YWNwdlUzaENnQ0lpbkFGQUFEd0cwRUxBSURNbFhXN2NJaExCaFpUZUl0bDNlNFA0WXBUVTJZQVlIVENGUUFBd0JjSldnQUFaR3c0MUYzNURKQzRzNllxN21KOWk4SVZBQkNkNnlGY2NhRTFBQURBbHdoYUFBRGtyUTlaN09WZUJKTDJHT05xSE9FS0FJak85WlBwRlEvYUF3QUFmTXViOVhxdFFBQUFHU3JyOWpDRThLdmVrN2lmWXJtTktsd0JBTkVScmdBQUFINklpUllBQVBreUNwblUzWTRkc2hDdUFJRG9DRmNBQUFDdkptZ0JBSkNoc201UFFnakhlay9pRm1POFBlRUtBSWlPY0FVQUFMQlJnaFlBQUhreXpZTFVYVFpWOFdtWDc3R3MyejVZTVJldUFJQW9DRmNBQUFCYkkyZ0JBSkNac201WElZUURmU2RoanlHRTFTN2UzaEN1K1B5MTUwTUZBS01TcmdBQUFIWkMwQUlBSUNQRFNvTlIxaW5BRHAwM1ZYRzNyZitjY0FVQVJPVzIvOWt2WEFFQUFPeVNvQVVBUUY3T0hReVR1UHZoYzc1UndoVUFFSlhiWVJYZTFUYkRsUUFBQUY4amFBRUFrSW15Ym85Q0NPLzFtOFN0Tm5XYlZiZ0NBS0lpWEFFQUFFUkQwQUlBSUI4YnYrVVBrYmxwcXVMaU5TOUp1QUlBb2lKY0FRQUFSRW5RQWdBZ0E4UGg4YkZlazdqRmo3dzk0UW9BaUlwd0JRQUFFRDFCQ3dDQXhKVjF1MithQlJtNGJLcWllKzdiSEZicHpJZHd4WUVQQ0FDTVNyZ0NBQUNZRkVFTEFJRDBMUndrazdqSDUweXpFSzRBZ0tnSVZ3QUFBSk1sYUFFQWtMQ3liZzkvZEowQ1RNaDVVeFVQWDNxNXdoVUFFQlhoQ2dBQUlBbUNGZ0FBYVZ1RkVQYjBtSVRkTjFXeGV2cjJoQ3NBSUNyM2ZiQ2lEMWk4Wk0wWEFBQkF6QVF0QUFBU1ZkYnRTUWpodmY2U3VENVFJVndCQUhFUnJnQUFBSkwyWnIxZTZ6QUFRSUxLdXUwZmFyL1ZXeEwyK1JCSHVBSUF4aWRjQVFBQVpFUFFBZ0FnUVdYZDlqZjcvNjYzQUFCc2tYQUZBQUNRSlVFTEFJREVsSFc3SDBLNEN5SHM2UzBBQUJzbVhBRUFBR1R2djNJdkFBQkFnaFpDRmdBQWJKQndCUUFBd0JNbVdnQUFKS1NzMjhNUXdxOTZDZ0RBS3dsWEFBQUFmSVdKRmdBQWFUblhUd0FBZnREamszREZKMFVFQUFENE1oTXRBQUFTVWRidFNRamhIL29KQU1BTGZBNVhYRFZWY2FWd0FBQUEzMmVpQlFCQU9pNzBFZ0NBWnhDdUFBQUFlQVZCQ3dDQUJKUjF1d2doSE9nbEFBQmZJVndCQUFDd0lWYUhBQUJNWEZtMyt5R0V1eERDbmw0Q0FQQ0VjQVVBQU1BV21HZ0JBREI5S3lFTEFBQUd3aFVBQUFCYlpxSUZBTUNFbFhWN0dFTDRWUThCQUxJbVhBRUFBTEJESmxvQUFFemJoZjRCQUdUclVyZ0NBQUJnOXdRdEFBQW1xcXpiMHhEQ3NmNEJBR1RsK3NuMGlnZXRCd0FBMkQxQkN3Q0E2VHJYT3dDQUxBaFhBQUFBUkVUUUFnQmdnc3E2WFlVUUR2UU9BQ0Jad2hVQUFBQ1Jlck5lci9VR0FHQkN5cnJkRHlIY2hSRDI5QTBBSUNuQ0ZRQUFBQk5nb2dVQXdQU2NDMWtBQUNSRHVBSUFBR0JpVExRQUFKaVFzbTZQUWdqLzFETUFnRWtUcmdBQUFKZ3dFeTBBQUtibFhMOEFBQ1pKdUFJQUFDQVJnaFlBQUJOUjF1MDhoSENzWHdBQWszRWJRcmdZd2hWMzJnWUFBSkFHcTBNQUFDYWdyTnY5RUVJWFFqalFMd0NBcUFsWEFBQUFKTTVFQ3dDQWFWZ0lXUUFBUkV1NEFnQUFJQ01tV2dBQVJLNnMyOE5obXNXZVhnRUFSRU80QWdBQUlGTW1XZ0FBeE85Y3lBSUFJQXJDRlFBQUFKaG9BUUFRczdKdVQwSUkvOUFrQUlEUkNGY0FBQUR3T3laYUFBREU3VngvQUFCMlRyZ0NBQUNBcnhLMEFBQ0lWRm0zOHhEQ1cvMEJBTmlKK3lGY2NTRmNBUUFBd0xkWUhRSUFFS0d5YnZkRENQMEQvajM5QVFEWW1qNWNjVFdFS3pwbEJnQUE0RGxNdEFBQWlOTkN5QUlBWUN1RUt3QUFBSGdWRXkwQUFDSlQxdTFoQ09GWGZRRUEyQmpoQ2dBQUFEYkdSQXNBZ1BoYzZBa0F3S3NKVndBQUFMQVZnaFlBQUJFcDYvWWtoSENzSndBQVAwUzRBZ0FBZ0swVHRBQUFpSXRwRmdBQUx5TmNBUUFBd0U0SldnQUFSS0tzMjBVSTRVQS9BQUMrNjNFSXFBcFhBQUFBc0hOdjF1dTFxZ01Bakt5czIvMFF3bDBJWVU4dkFBQys2SEdZWEhIVlZNV1ZFZ0VBQURBV0V5MEFBT0t3RXJJQUFQZ1Q0UW9BQUFDaVk2SUZBTURJeXJvOUNpSDhVeDlJM0ZrSTRUQ0VjQ3BVQk1CM0NGY0FBQUFRTlJNdEFBREdkNjRISk82bXFZclY1N2RZMXUzcEVMZ1F1Z0RnUy9xZkRROGhoRStxQXdBQVFJeE10QUFBR05GdzRQei85SURFL2JXcGl1NlBiN0dzMi8wbmdZdDNQZ1FBL01GOUNHSGVWSVhBQlFBQUFGRVJ0QUFBR0ZGWnQzY2hoQU05SUdHWFRWWE12L2YyaEM0QStJWVBJWVJWVXhVUGlnUUFBRUFNQkMwQUFFWlMxbTIvU21HcC9pU3MzN0YvK05LRHNTZWhpMFVJNGEwUENBQWhoTnRodXNXZkppUUJBQURBcmdsYUFBQ01vS3pid3hCQ04rd2doMVNkTlZXeGVzMTdHLzZ0OUtHTHVkQUZBSnY0MlFJQUFBQ3ZKV2dCQURDQ3NtNHZRZ2p2MVo2RTNUZFZjYmpKdHlkMEFjREFkQXNBQUFCR0pXZ0JBTEJqWmQyZWhCRCtvZTRrN205TlZYemExbHNzNi9ab0NGejB3WXNESHlhQTdQVHJxVlpOVlp4clBRQUFBTHNtYUFFQXNHTmwzZmFIejhmcVRzSnVtcW80MmRYYkU3b0F5TnJOTU4zaUx2ZENBQUFBc0R1Q0ZnQUFPMVRXYlg4WS9IYzFKM0YvR2V2QWF3aGRMSWJReFo0UEdrQVdIb2V3eFpWMkF3QUFzQXVDRmdBQU8xTFc3WDRJb1hQam5zUjlhS3BpRWNOYkxPdjJkQWhjQ0YwQTVPRjZDRnc4NkRjQUFBRGJKR2dCQUxBalpkMnVRZ2hMOVNaaC9ZM2l3eGdQdUlRdUFMSnhQNFF0UG1rNUFBQUEyeUpvQVFDd0EyWGRIb1lRZmxWckV2ZExVeFhuTWIvRlliTE01OERGdXdoZUVnRGJjZFpVeFVwdEFRQUEyQVpCQ3dDQUhTanI5c3FoTG9tN2JhcmlhRXB2VWVnQ0lIbTMvZmY0cGlydXRCb0FBSUJORXJRQUFOaXlzbTVQUWdqL1VHY1M5N2NwajJsL0VycVloeENPSTNoSkFHeEd2OVpxMFZURmhYb0NBQUN3S1lJV0FBQmJWdFp0RjBKNHE4NGs3THFwaXROVTN0Nnc2dWR6Nk1LL1hZQTBYUGZmMTV1cWVOQlBBQUFBWGt2UUFnQmdpOHE2N1E5cS82N0dKTzR2cVk1bEY3b0FTTXI5c0VxazAxWUFBQUJlUTlBQ0FHQkxobFVFL2VIem5ocVRzTE9tS2xZNU5IZ0lYU3lHNE1WQkJDOEpnQitUemM4dUFBQUF0a1BRQWdCZ1M4cTZQUThoL0t5K0pLemZlMytZNHhqMnNtNlBoaWtYUWhjQTAzUXpUTGV3U2dRQUFJQVhFN1FBQU5pQzRlYjdyMnBMNG41cXF1SWk5eVkvQ1YzTVRiQUJtSlRISVd6eFNkc0FBQUI0Q1VFTEFJQXRLT3UyZjJCL3JMWWs3S2FwaWhNTi9yMnliaytIS1JlblFoY0FrMkdWQ0FBQUFDOGlhQUVBc0dGbDNmYUh6LzlRVnhMM056ZUF2MDNvQW1CU3JCSUJBQURnMlFRdEFBQTJyS3pidXhEQ2dicVNzTXVtS3VZYS9IeGwzYzZId01XN3FieG1nQXhaSlFJQUFNQ3pDRm9BQUd4UVdiZUxFTUwvcUNrSjZ3K2hqcHFxdU5Qa2x5dnJkdi9KbEF1aEM0QTRXU1VDQUFEQU53bGFBQUJzeUhDQWVtZEZBSWx6K0xRaFQwSVgvYlNMNHlUZUZFQTZyQklCQUFEZ3F3UXRnUC9QM3QwazFYR2tiUU5PTzk0NWZDc0FyMEI0QmFCSlRvV2pJbklxdEFMakZmaG9CWTFXWUpoV1JNWXJUWE5pV0VIRENndzdFQ3ZnaS9LYmRzdHVTUzUrRHFjcTY3b2l2SUIrYmhxT1R0NzVKQUJQcE92TGFRamh0WG5Tc0p1YzRxNkFuMTdYbDkxUFNoY3ZXdnZmQnpCVG5oSUJBQURnc3hRdEFBQ2VRTmVYdlJEQ3Y4MlN4djJRVTN3djVQVlN1Z0NZbko5eWlpZGlBUUFBNEErS0ZnQUFUNkRyeTduVi96VHVJcWQ0SU9UblZVc1h4N1Y0c2JPay8rMEFFL05oS01CNVNnUUFBSUNnYUFFQThIaGRYNFlEMFA4MVNocjNmVTd4VXNpYlV6Zm5IQ2xkQUd6TVRYMUt4TjlEQUFDQWhWTzBBQUI0aEs0djJ5R0VTNGVlTk80c3AzZ2s1T240cEhReC9MZTE5SGtBUEtQYllkTlFUdkhVMEFFQUFKWkwwUUlBNEJHNnZxeENDRCtiSVEwYkRwUjJyVXFmcnJwVjU0Ly9sQzRBbnNlN25PS3hXUU1BQUN5VG9nVUF3QU4xZmRtdDJ5d2NiTkt5bjNLS0p4S2VCNlVMZ0dkMUZVSTRVRVlFQUFCWUhrVUxBSUFINnZveXJJeCtiWDQwN0NhbnVDdmdlZXI2Y2xRTEY2K1dQZ3VBTmJxdFpZdExRd1lBQUZnT1JRc0FnQWZvK25JUVF2alY3R2pjeTV6aXVaRG5yZXZMOWlkYkxwUXVBTmJqVFU3eDFHd0JBQUNXUWRFQ0FPQUJ1cjRNaDgvN1prZkRMbktLQndKdXl5ZWxpeU8vd3dDZTNGbE84Y2hZQVFBQTJxZG9BUUJ3VDNVZC95L21SdU8reXlsZUM3bGRYVjkyUHlsZHZGajZQQUNleUZWOVN1U2pnUUlBQUxSTDBRSUE0QjdxYmZEaDhIbkwzR2pZdTV6aXNZQ1hRK2tDNEVuZDFyTEZwYkVDQUFDMFNkRUNBT0FldXI2c1FnZy9teGtOR3c2SGR0M0VYYTVhdWppdXhZdWRwYzhENElHR3Y2ZkhPY1ZUQXdRQUFHaVBvZ1VBd0VqMThQRTM4Nkp4Ynh3SzhZZXVMM3QxeTRYU0JjREQyQklGQUFEUUlFVUxBSUNSdXI2OER5RzhNaThhZHBWVDNCTXduL05KNmVMSTgwa0E5L0poK04xcFd4UUFBRUE3RkMwQUFFYm8rbklRUXZqVnJHamN5NXppdVpENUoxMWZEdXVXaTBPbEM0QlJyb2JmbVRuRmErTUNBQUNZUDBVTEFJQVJ1cjVjaGhCZW1CVU4rNUJUUEJRdzk2VjBBVERhYlFqaElLZDRhV1FBQUFEenBtZ0JBUEFQdXI0TTcyci95NXhvMkhEd3MrZVdMWS9WOWVXb0ZpNDhzd1R3Wlc5eWlxZm1Bd0FBTUYrS0ZnQUFYOUgxWlR1RWNPMldObzE3bTFOY0NabW5VbjkzSGlwZEFIelJ1NXppc2ZFQUFBRE1rNklGQU1CWGRIMDVDU0g4YUVZMDdLWnVzL2dvWk5aQjZRTGdpejZFRUk3OERRWUFBSmdmUlFzQWdDL28rckliUXZqTmZHaWM5ZVU4bS9wN2RTaGNERStNdkRCNWdIQVZRamhRdGdBQUFKZ1hSUXNBZ0MvbytuSWVRdGczSHhwMmtWTThFRENib0hRQjhLZmJXcmE0TkJJQUFJQjVVTFFBQVBpTXJpL0Q0Zk92WmtQalh1WVV6NFhNcHRYU3hWSDliMGNnd0FJTlpZdERmNWNCQUFEbVFkRUNBT0F6dXI1Y08reWpjV2M1eFNNaE16VmRYL1pxNGVMUTcyRmdnVHpwQlFBQU1BT0tGZ0FBZjlQMVpSVkMrTmxjYU5od2EzYlhlL0JNbmRJRnNGRHZjb3JId2djQUFKZ3VSUXNBZ0U5MGZka09JUXpiTExiTWhZYTl6U211Qk15Y2RIMDVySVdMUTcramdRV3dlUW9BQUdEQ0ZDMEFBRDdSOVdWWTFmemFUR2pZVFU1eFY4RE1tZElGc0JCWElZUURHNmdBQUFDbVI5RUNBS0NxSytyL2JSNDA3b2VjNG5zaDA0cGF1aGh1ZmI4U0t0QWdaUXNBQUlBSlVyUUFBS2k2dnB5SEVQYk5nNFpkNUJRUEJFeUw2dE5QZjJ5NVVMb0FXbkpieXhhWFVnVUFBSmdHUlFzQWdQODdvQnR1US85aUZqVHVlNGMwTElIU0JkQWdaUXNBQUlBSlViUUFBQmF2SHNnTlgxcnZMSDBXTk8xZFR2Rll4Q3hOMTVmZFdyZ1lDblV2L0FBQU16YVVMWTV6aXFkQ0JBQUEyQ3hGQ3dCZzhicStyRUlJUHk5OURqUnRPSmpaOWI0N1M2ZDBBVFRpamJJRkFBREFaaWxhQUFDTFZnL2RobTBXVzB1ZkJVMzdLYWQ0SW1MNGovcjcvNmorWjZNUk1EZit0Z01BQUd5UW9nVUFzR2hkWDRiYmdLK1hQZ2VhZHBOVDNCVXhmRm5YbDcxYXVEaFV1Z0JtNUN5bmVDUXdBQUNBNTZkb0FRQXNWdGVYZ3hEQ3IzNENhTnpMbk9LNWtHRWNwUXRnWnBRdEFBQUFOa0RSQWdCWXJLNHZsOTdvcDNFWE9jVURJY1BEZEgwNXJJV0xRMDlNQVJOMkZrSTR6aWwrRkJJQUFNRHpVTFFBQUJhcDY4dHc4KzhYNmRPNDczS0sxMEtHeDZ0L040YkN4U3ZqQkNib0tvUndvR3dCQUFEd1BCUXRBSURGNmZxeUhVSzRkanVaeHIzTkthNkVERStyL2cwNXF2L1ppZ1JNaWJJRkFBREFNMUcwQUFBV3ArdkxjUGo4cytScDJHMElZZGRCQzZ4WDE1ZTlUMG9YeW52QUZDaGJBQUFBUEFORkN3QmdVYnErN0lZUWZwTTZqWHVUVXp3Vk1qd2ZUNHNBRTZKc0FRQUFzR2FLRmdEQW9uUjllZThRak1aZDVSVDNoQXliVVF0OWYyeTUyQkVEc0NHM3RXeHhLUUFBQUlDbnAyZ0JBQ3hHMTVlREVNS3ZFcWR4TDNPSzUwS0d6ZXY2Y2xnTEZ3cCt3Q1lvV3dBQUFLeUpvZ1VBc0JoZFg2N2RMcVp4WnpuRkl5SER0Tmh5QVd5UXNnVUFBTUFhS0ZvQUFJdlE5ZVU0aFBBdmFkT3c0U0JsTDZkNExXU1lycTR2ZnhRdTlzVUVQQk5sQ3dBQWdDZW1hQUVBTksvcnkzWUlZVGg4M3BJMkRYdWJVMXdKR09haGJyazRycVVMZjUrQWRWTzJBQUFBZUVMZkdpWUFzQUFyaDFnMDdpYUVjQ0prbUk5aCsweE9jU2hhRElXTE4vWC94d0RyTW53V1B1LzZzbWZDQUFBQWoyZWpCUURRdEhwaitEY3AwN2czT2NWVEljTzhkWDA1cUZzdVhva1NXQk9iTFFBQUFKNkFvZ1VBMExTdUwrZmV3YWR4RnpuRkF5RkRPendyQXF5WnNnVUFBTUFqS1ZvQUFNM3ErbklZUXZoZkNkTzQ3eDJVUUp1NnZteUhFQTdyRTFnN1lnYWVrTElGQUFEQUl5aGFBQURONnZweTdXQ0t4cDNsRkkrRURPMnI1Y0ZqVzVxQUo2UnNBUUFBOEVEZkdod0EwS0t1TDI3LzBycmJldWdLTEVCTzhYMTlKdWo3b1dRbGMrQUpERThUblhkOTJUTk1BQUNBKzdIUkFnQm9UbDIxZnUxZGV4cjNOcWU0RWpJc1U5ZVgzUkRDVVMxYytYc0hQSWJORmdBQUFQZWthQUVBTktmcnkya0k0YlZrYWRoTlRuRlh3RUF0Ri81UnVMREpDWGdvWlFzQUFJQjdVTFFBQUpyUzlXVllxLzZyVkduY0Q4TXpBa0lHUHRYMTVZL0N4UXVEQVI1QTJRSUFBR0FrUlFzQW9DbGRYODVEQ1B0U3BXRVhPY1VEQVFOZlVrdUhLMzhQZ1FkUXRnQUFBQmhCMFFJQWFFYTl5ZnVMUkduY2R6bkZheUVELzZRV0xvNDhwd1hjMDFDMjJQTjVBd0FBNE1zVUxRQ0FKdFEzNmkrOVQwL2ozdVVVajRVTTNFZlhsOTI2NFVMaEFoanJxbTYyK0doaUFBQUEvKzFiTXdFQUduR3NaRUhqYnV0QktjQzlETGZTYzRyRFpvdnZRZ2huOWZjSndOZThDQ0djMXpJekFBQUFmMk9qQlFBd2UvV203ckROWWt1YU5PeW5uT0tKZ0lISHFnZW54L1UvZnp1QnI3SFpBZ0FBNERNVUxRQ0EyZXY2OGo2RThFcVNOT3dxcDdnbllPQXBLVndBSXlsYkFBQUEvSTJuUXdDQVdldjZjcUJrd1FJY0N4bDRhc09oYVU1eGVKSm8yQXoxMXBNaXdCY016NGpZcWdVQUFQQUpHeTBBZ0ZucituSlp2L3lGVm4zSUtSNUtGMWczR3k2QWYzQ1dVend5SkFBQUFFVUxBR0RHdXI0TVgvVCtJa01hOTExTzhWckl3SE9waFl1aDREVnN1OWd4ZU9BVHloWUFBTURpQlVVTEFHQ3U2aUhRdFJ1M05PNXRYZXNQc0JHMTFLaHdBWHpxcDV5aXAwUUFBSUJGKzNicEF3QUFac3RhYzFwMzZ6MTBZTk55aXFjNXhkMmgrRlYvTHdIOHE1YXdBQUFBRnN0R0N3QmdkcnErREFjK3YwbU94cjBaRGppRkRFeEYzU1oxck93SVZENnJBQUFBaTZWb0FRRE1UdGVYOHhEQ3Z1Um8yRVZPOFVEQXdCUXBYQURWc09YbUlLZDRhU0FBQU1EU0tGb0FBTFBTOVdVNGZQNVZhalR1WlU3eFhNakFsQ2xjQU1vV0FBREFVaWxhQUFDejB2WGxPb1N3SXpVYWRwWlQ5TzQ1TUJzS0Y3QjROeUdFdlp6aXg2VVBBZ0FBV0k1dlpRMEF6RVhYbDJNbEN4bzMzQXBkQ1JtWWsrRndOYWM0L083YUc4cGl3b1BGR1Q2Zm45ZlNGUUFBd0NJb1dnQUFzMUMvdUhVQVRldE9jb3JYVWdibWFQajlWVGZ5Zktkd0FZdnpJb1J3S25ZQUFHQXBGQzBBZ0xsWVdVZE80NGExMnlkQ0J1YnViNFdMQzRIQ1lyenErcUpzQVFBQUxNSTNkM2Qza2dZQUpxM3J5N0NLL045U29uRS81QlRmQ3hsb1RkZVhnMXFZM0JjdUxNSlBPVVhsVVFBQW9HbUtGZ0RBNUhWOU9YYzRRK011Y29vSFFnWmExdlhsc0c3dTJSRTBOTzlOVHRGMkN3QUFvRm1LRmdEQXBOVkRtZitWRW8zN1BxZDRLV1JnQ2JxK0hOVU5Gd29YMEs3YkVNS0J6emNBQUVDcnZwVXNBREJ4MWc3VHVqT0hFTUNTMUZ2dXc3TmdiK3RoTE5DZXJSRENlZGVYWGRrQ0FBQXRVclFBQUNhcjY0dmJyclJ1T0dBOGxqS3dORG5GanpuRjRlLzhjQWg3NWdjQW1qU1VMZDUzZmRrV0x3QUEwQnBGQ3dCZ2t1cnROd2ZRdEc0MUhEWktHVmlxV3JnWW5oTDVMb1J3NFFjQm12TWloSEFxVmdBQW9EV0tGZ0RBVkszcUxUaG8xVTFPMGRNNEFQOVh1TGpPS1I2RUVGNkdFSzdNQkpyeXF1dUx6endBQUVCVEZDMEFnTW5wK2pJY3RMeVdESTA3RWpEQVgrVVV6M09LZXlHRU4wTWh6WGlnR1Q5MmZmSFpCd0FBYUlhaUJRQXdSU3VwMExpTDRUQlJ5QUNmbDFNY25ob1lDaGR2UXdpM3hnUk4rS1hyeTU0b0FRQ0FGbnh6ZDNjblNBQmdNdXBOdDE4a1F1TytHOWJrQ3huZ24zVjkyYTBsVE51dVlQNkc0dFJ1VHZHakxBRUFnRG16MFFJQW1JeXVMOXUyV2JBQTc1UXNBTVliZm1mbUZJY2k1c3RoSTVEUndheHRoUkJzOVFJQUFHWlAwUUlBbUpMakVNS09SR2pZclRJUndNTU1UeTdsRkE5Q0NHODhKd0t6OXFMcnk2a0lBUUNBT2ZOMENBQXdDWFV0K0cvU29IRnZjb29PRmdBZXFXN0JHZ3FhUDVzbHpOWlBPY1VUOFFFQUFIT2thQUVBVEVMWGwvY2hoRmZTb0dGWE9jVTlBUU04blZyVUhBcHMrOFlLcy9SeTJGWWpPZ0FBWUc0VUxRQ0FqZXY2TXF3Qi8xVVNOTTVCQXNDYTFNOFNwNTRnZzlrWm5nSGF6U2wrRkIwQUFEQW4zMG9MQUpnQUs0TnAzUWNsQzREMUdYN0g1aFNIN1JadjY4RXRNQTliSVFTZmtRQUFnTmxSdEFBQU5xcnJ5MUVJNFlVVWFOeXhnQUhXTDZlNENpRU16elI5TUc2WWpSZGRYMDdGQlFBQXpJbW5Rd0NBamVuNnNoMUN1SzQzMmFCVmIrdkJId0RQeUhNaU1EdHZjb29LRndBQXdDellhQUVBYk5KS3lZTEczWGdhQjJBei92YWNDREI5SjExZjl1UUVBQURNZ1kwV0FNQkdkSDBaRGo1K00zMGE1Mlltd0FUVXp4M0Q3K045ZWNDa1hZVVFEbktLSDhVRUFBQk1tWTBXQU1DbU9IeW1kUmRLRmdEVGtGTzh6aWtPVDRuOFVMY05BZFAwd3I4VEFBQ0FPVkMwQUFDZVhYMHozWTFTV3JlU01NQzA1QlRmaHhDR3B3bmVpUVltNjFYWGwyUHhBQUFBVSticEVBRGcyWFY5dVE0aDdKZzhEVHZMS1I0SkdHQzZ1cjdzMVp2ekw4UUVrL1I5VHZGU05BQUF3QlRaYUFFQVBLdXVMeXNsQ3hwM2E1c0Z3UFFOQjdnNXhhRnM4VlA5M1ExTXkvdXVMOXN5QVFBQXBralJBZ0I0TnZXTFVtdUFhZDFKVHZGYXlnRHprRk04cWMrSlhJZ01KbVduYnAwQkFBQ1lIRStIQUFEUHB1dkw4RVhwYXhPbllUYzV4VjBCQTh4VDE1ZkRlckM3SlVLWWpEYzVSWVVMQUFCZ1VteTBBQUNlUlgwSFhjbUMxdG5ZQWpCak9jWDNJWVNoTUhjbVI1aU1rL3B2Q1FBQWdNbXcwUUlBZUJaZFg4NURDUHVtVGNNdWNvb0hBZ1pvUTllWGc3cmRZa2Vrc0hGWElZU0RuT0pIVVFBQUFGTmdvd1VBc0haMURiZVNCYTJ6elFLZ0lUbkZvU1E2M0tKL0oxZll1QmNoaEpVWUFBQ0FxYkRSQWdCWXE2NHYyeUdFUzdkQmFkeTduS0tpQlVDamJMZUF5ZmloUHZFREFBQ3dVVFphQUFEcmR1eFFnc2JkdW1FSjBMWlB0bHU4RlRWczFHa3RjZ01BQUd5VW9nVUFzRFpkWDNZOXA4QUNyTHdYRHRDKzRYZDlUbkVvMW4wZlFyZ1NPV3pFVmdqQlJnc0FBR0RqRkMwQWdIVmExUzlEb1ZVM09jVVQ2UUlzUjA3eE1xZG91d1Zzem43WEYyVnVBQUJnb3hRdEFJQzFxRytadnpaZEduY2tZSUJsc3QwQ05tclY5V1ZQQkFBQXdLWW9XZ0FBNjdJeVdScDNVZC9zQjJDaGh1MFdJWVNoWFByT3p3QThxMkZyM3FtUkF3QUFtNkpvQVFBOHVhNHZ3eTMvZlpPbGNiWlpBRENVTFQ3bUZJZG5ERjRPVDBxWkNEeWJGMTFmbExzQkFJQ04rT2J1N3M3a0FZQW4wL1ZsTzRSd1hXK1pRYXZlMXBYeEFQQ24ramxvK1B2d282bkFzL20rYnBjQkFBQjROalphQUFCUDdWakpnc2JkaGhCT2hBekEzMzJ5M2VLSCt2Y0NXRDlQaUFBQUFNOU8wUUlBZURKZFgzWkRDRCtiS0kwN0hnN1NoQXpBbCtRVTM0Y1FoczlGSHd3SjFtNTRRa1FKRmdBQWVGYWVEZ0VBbmt6WGwrRlE0WldKMHJDcm5PS2VnQUVZcSt2TGNYMU94TVl2V0srWE9jVnpNd1lBQUo2RGpSWUF3SlBvK25LZ1pNRUNIQXNaZ1B2SUtRNDM3WWVTM3BYQndWcWRkbjNaTm1JQUFPQTVLRm9BQUUvRnVsNWE5OEV0U1FBZUlxZDRYVGNpdlRWQVdKc2RwVmdBQU9DNUtGb0FBSTlXVjJLL01Fa2FkdXVMZXdBZUs2YzRQQ0h5TW9Sd1k1aXdGajkzZmZITUd3QUFzSGFLRmdEQW85VDF2Q3RUcEhFbncyMWtJUVB3V0hVNzBuQVEvTUV3WVMxT2pSVUFBRmczUlFzQTRMR0drc1dXS2RLd0cwL2pBUENVY29vZmM0cUhJWVEzZFdzUzhIUmVkSDFSQkFjQUFOYnFtN3U3T3hNR0FCNms2OHR1Q09FMzA2TnhiM0tLYmtZQ3NCYjFtWU5UejdEQmsvcytwM2hwckFBQXdEcllhQUVBUEliRFoxcDNvV1FCd0RyVmcrQ0RFTUk3ZzRZblpTTVpBQUN3Tm9vV0FNQ0RkSDBaMWwzdm14Nk5PeFl3QU90V254SVovdWI4NENrUmVETDdYVjk4bGdNQUFOWkMwUUlBZUNnM3hHamRtWFhUQUR5bm5PTDdFTUx3bE1pVndjT1RXSFY5MlRaS0FBRGdxU2xhQUFEMzF2VmxGVUxZTVRrYWRtdWJCUUNia0ZPOHppbnVlVW9FbnNTVzV3NEJBSUIxVUxRQUFPNmwzZ2h6QUUzclRvWTE3bElHWUZNOEpRSlA1bFhYbHdQakJBQUFucEtpQlFCd1h5ZjFaaGkwNmlhbnVKSXVBSnRXbnhJNThKUUlQTnFwSjBRQUFJQ25wR2dCQUl6VzlXVllZLzNheEdpY2pTMEFURVpPOGJLV0xjNmtBZysyNHpNZUFBRHdsTDY1dTdzelVBQmdsSzR2NXlHRWZkT2lZUmM1UmF1bEFaaWtyaTlISVlSZnBBTVA5bDFPOGRyNEFBQ0F4N0xSQWdBWXBYNnhyMlJCNjQ0a0RNQlU1UlJQUXdqZkQ4OWNDUWtlNU5UWUFBQ0FwNkJvQVFEOG8vcWU4Y3FrYU53N054d0JtTHI2bE1qd25OdUZzT0RlOXJ1K0hCb2JBQUR3V0lvV0FNQVl4L1ZkWTJqVnJUSVJBSE9SVS94WW43cDZKelM0dDVOYUpBY0FBSGd3UlFzQTRLdTZ2dXpXb2dXMGJEVWNXa2tZZ0RuSktRNmYwZDdVd2lBd3pvNS8zd0FBQUkrbGFBRUEvSlBobHYrV0tkR3dtNXppaVlBQm1LT2M0bWtJWWRodWNTTkFHTzI0RnNvQkFBQWVSTkVDQVBpaXJpL0RsL2F2VFlqR0hRa1lnRG5MS1Y2R0VQWkNDQmVDaEZHMlBCc0hBQUE4aHFJRkFQQTFidm5UdWc4NXhYTXBBekIzd3hOWU9jV2hKUHRPbURESzYxb3NCd0FBdURkRkN3RGdzN3ErRExmOFg1Z09qZk0rTndCTnlTa09mOXZlU0JWR3NkVUNBQUI0RUVVTEFPQy9kSDNadHMyQ0JYaWJVN3dXTkFDdHlTbWVoaEMrRHlIY0NoZSthcjhXekFFQUFPNUYwUUlBK0p6aittNHh0T3BXbVFpQWx1VVVMME1JZXlHRUswSERWNjFxMFJ3QUFHQTBSUXNBNEMrNnZ1eUdFSDQyRlJwM1BMeGxMMlFBV2xZM054MkVFRDRJR3I1b3gzTnlBQURBZlgxemQzZG5hQURBbjdxK3ZBOGh2RElSR25hVlU5d1RNQUJMMHZWbDJPVDBvOURoczRadFo3dUt1QUFBd0ZnMldnQUFmK3I2Y3FCa3dRSzRzUWpBNHVRVWg3OS9ieVFQbjdYbFdUa0FBT0ErRkMwQWdFK2RtZ2FOTzhzcG5nc1pnQ1hLS1E2ZjlWN1cyL3ZBWDcydXp5Z0NBQUQ4STBVTEFPQjNYVitPNi92RTBLcmhVR2tsWFFDV3JCWU9ENVF0NExOOFZnUUFBRVpSdEFBQWhwTEZ0aThWV1lDVG5PSzFvQUZZdXB6aVpRaGh1TGwvdGZSWndOKzhyczhwQWdBQWZKV2lCUUFRYXNsaXl5Um8yQlpOSHRVQUFCYlVTVVJCVkkxM3R3SGdQM0tLSCt0bWl3dGpnYjlRUUFjQUFQN1JOM2QzZDZZRUFBdlc5V1V2aFBCdlB3TTA3b2VjNG5zaEE4Qi82L3B5T3R6a054cjQwOHY2ekE0QUFNQm4yV2dCQUxqbFQrc3VsQ3dBNE10eWlrY2hoSGRHQkgreTFRSUFBUGdxUlFzQVdMQ3VMNGNoaEgwL0F6VHVXTUFBOEhVNXhlSHY1UnRqZ3QvdDEzOHJBUUFBZkphaUJRQXNtMjBXdE80c3AzZ3BaUUQ0WnpuRlUyVUwrSk4vS3dFQUFGK2thQUVBQzlYMVpWaUh1eU4vR25acm13VUEzRTh0VzN4Zi80N0NrdTEwZlRueUV3QUFBSHlPb2dVQUxGRFhsMjBIMEN6QVNVN3hvNkFCNEg3cU5xZ0RaUXNJS3lNQUFBQStSOUVDQUpacFdJTzdKWHNhZHBOVDlNVTRBRHlRc2dYOHpsWUxBQURnc3hRdEFHQmh1cjRNWDVpL2xqdU44NFU0QUR4U0xWdnNoUkN1ekpJRlU5NEZBQUQraTZJRkFDeVBMd3BwM1VWTzhWektBUEI0T2NYcnV0bEMyWUtsc3RVQ0FBRDRMNG9XQUxBZzlRdkNmWm5UT0YrRUE4QVR5aWwrVkxaZzRaVFZBUUNBdjFDMEFJQ0Y2UHF5N1F0Q0Z1QmR2WGtMQUR3aFpRc1didGhxY2JqMElRQUFBUCtoYUFFQXkzRThmRUVvYnhwMnEwd0VBT3VqYk1IQ0hTOTlBQUFBd0g4b1dnREFBblI5MmZYRklBdXdxZ2RBQU1DYUtGdXdZUHRkWHc3OEFBQUFBRUhSQWdBVzR5U0VzQ1Z1R25hVlV6d1JNQUNzM3lkbGl3L0d6Y0xZbmdZQUFQeE8wUUlBR2xkdlhiMlNNNDJ6c1FVQW50RlF0c2dwSG9ZUXpzeWRCYkhWQWdBQStKMmlCUUMwenkxL1d2Y2hwM2d1WlFCNGZqbkZJMlVMRnVaSTRBQUFnS0lGQURTczY4dndKZUFMR2RNNDJ5d0FZSU5xMmVKQ0JpekU2NjR2dThJR0FJQmxVN1FBZ0VaMWZkbTJ6WUlGZUp0VHZCWTBBR3pjOEl6SWxSaFlpSldnQVFCZzJSUXRBS0Jkd3kzL0xmblNzQnRsSWdDWWhweml4eERDZ2JJRkMvRzZGdHNCQUlDRlVyUUFnQWJWVmJZL3k1YkdyZXFoRGdBd0Fjb1dMSXpuNndBQVlNRVVMUUNnVGFkeXBYRVhPVVUvNXdBd01jb1dMSWlpQlFBQUxKaWlCUUEwcHV2TDhNWDJ2bHhwbkhleEFXQ2lsQzFZaUsydUwwZkNCZ0NBWlZLMEFJRDJ1T1ZQNjg1eWl1ZFNCb0RwK3FSc2NTTW1HbWFyQlFBQUxKU2lCUUEwcE92TDhFWGZqa3hwMksxdEZnQXdEN1ZzY1ZqL2ZrT0xYdFNOZ2dBQXdNSW9XZ0JBSTdxK2JEdUFaZ0ZPY29yWGdnYUFlY2dwWHRiTkZzb1d0TXJ6SVFBQXNFQ0tGZ0RRanBQaG5XQjUwckNibktJeUVRRE1qTElGalh2ZDlXVlh5QUFBc0N5S0ZnRFFnSzR2ZThNWGZMS2tjZDdBQm9DWnFtVUxmOHRwbGEwV0FBQ3dNSW9XQU5DR0V6blN1SXVjNG5zaEE4Qjg1UlJQUXdodlJFaURsSWdBQUdCaEZDMEFZT2E2dmh5R0VQYmxTT044ZVEwQURhaGxpN2V5cERGYlhWOXN0UUFBZ0FWUnRBQ0ErYlBOZ3RhZDFYWGpBRUFEY29xcjRlKzdMR21Nb2dVQUFDeUlvZ1VBekZqWGwrRkw2aDBaMHJCYjJ5d0FvRDA1eGVGUStrSzBOR1MvNjh1dVFBRUFZQmtVTFFCZ3B1cVhlQTZnYWQwcXAvaFJ5Z0RRcE9FSnZDdlIwaEQvUGdNQWdJVlF0QUNBK1JxMldXekpqNGJkNUJROWpRTUFqYXBseXNPNndRcGE0UGtRQUFCWUNFVUxBSmlocmk4SElZVFhzcU54dnFnR2dNYmxGSzlEQ0FmS0ZqUmlxK3VMejdBQUFMQUFpaFlBTUU4cnVkRzRpNXppdVpBQm9IMDV4VXRQTHRDUVEyRUNBRUQ3RkMwQVlHYnFEYWw5dWRFNE53RUJZRUZ5aXFjaGhMY3lwd0d2dXI3c0NoSUFBTnFtYUFFQU05TDFaZHMyQ3hiZ1hWMGpEZ0FzU0U1eCtKejdRZVkwd0ZZTEFBQm9uS0lGQU16THNGSjVSMlkwN0ZhWkNBQVdiZGhxZGJYMElUQjduc0lCQUlER2ZYTjNkeWRqQUppQnVuNzJOMW5SdURkMWRUZ0FzRkQxYys5bENHSEx6d0F6OW4xTzhWS0FBQURRSmhzdEFHQStUbVJGNDY2VUxBQ0Erb1NZcHhlWU8xc3RBQUNnWVlvV0FEQURYVjhPUWdpdlpFWGpmQmtOQVB3dXAzZ2VRdmpKTkpneFpTRUFBR2lZb2dVQXpJTnRGclR1UXoxUUFRRDRYVTV4K0F6OHdUU1lxYTJ1TDhvV0FBRFFLRVVMQUppNHJpL0RMZjhYY3FKaHQ3WlpBQUJmY0JSQ3VERWNaa3JSQWdBQUd2WE4zZDJkYkFGZ29ycStiSWNRaGplcXQyUkV3OTdtRkZjQ0JnQStwK3ZMWGdqaDNHZGladXIvNVJRL0NnOEFBTnBpb3dVQVROdktGOG8wN3NiVE9BREExK1FVTDIyL1lzWnN0UUFBZ0FZcFdnREFSSFY5MlEwaC9DZ2ZHcmR5d3c4QStDYzV4ZE1Rd3BsQk1VT0tGZ0FBMENCRkN3Q1lybFBaMExpTGVtZ0NBRERHc05YaXlxU1ltVmYxU1VnQUFLQWhpaFlBTUVGZFh3NUNDUHV5b1hFckFRTUFZOVV0V0VjR3hnelphZ0VBQUkxUnRBQ0FhWExMbjlhZDVSVFBwUXdBM0VkTzhUS0U4Sk9oTVRPS0ZnQUEwQmhGQ3dDWW1LNHZ3eTMvSGJuUXNOdTYraHNBNE41eWlpZkRFMlFteDR4NFBnUUFBQnFqYUFFQUUxSy9mSE1BVGV0TzZ1cHZBSUNIT3F6bFRaZ0xXeTBBQUtBaGloWUFNQzNEN2J3dG1kQ3dtNXppU3NBQXdHUFUwdWFSSVRJamloWUFBTkFRUlFzQW1JaXVMM3NoaE5meW9IRTJ0Z0FBVHlLbitENkVjR2Fhek1RclFRRUFRRHNVTFFCZ09rNWtRZU11Nm9FSUFNQlRHVXFjTjZiSkhIUjlzZFVDQUFBYW9XZ0JBQk5RdjNEYmx3V05zODBDQUhoU25oQmhaaFF0QUFDZ0VZb1dBTEJoWFYrMmJiTmdBZDdsRkM4RkRRQTh0WnppK2ZCWncyQ1pBVVVMQUFCb2hLSUZBR3plY010L1J3NDA3RGFFc0JJd0FMQkdLMCtJTUFOYlhWLzJCQVVBQVBPbmFBRUFHOVQxWmRkekNpekFxcTcxQmdCWUMwK0lNQ08yV2dBQVFBTVVMUUJnczRhYmQxc3lvR0UzT1VWUDR3QUFhMWVmRURremFTWk8wUUlBQUJxZ2FBRUFHOUwxNVNDRThOcjhhWnlicFFEQWN6cXV6NWJCVkwyb213MEJBSUFaVTdRQWdNMVptVDJOKzFCdmxnSUFQSXY2aElpbitaaTZBd2tCQU1DOEtWb0F3QVowZlJsdStlK2JQWTF6eUFFQVBMdWM0bWtJNGNMa21UQkZDd0FBbURsRkN3QjRabDFmdGtNSUorWk80OTdtRksrRkRBQnNpT2ZMbUxKRDZRQUF3THdwV2dEQTh4dHUrVytaT3cyN1ZTWUNBRGFwRmo3ZkNvR0oydXI2c2ljY0FBQ1lMMFVMQUhoR1hWOTJRd2cvbXptTk82N3Zvd01BYk5KUS9MeVJBQlBsK1JBQUFKZ3hSUXNBZUY1dStkTzZxL291T2dEQVJ0WGk1N0VVbUNqUGh3QUF3SXdwV2dEQU0rbjZNdHhZZW1YZU5NNWhCZ0F3R1RuRjl5R0VDNGt3UWZ0Q0FRQ0ErVkswQUlEbjQ1WS9yVHZMS1o1TEdRQ1lHRVZRSnFtVzhRRUFnQmxTdEFDQVo5RDFaZmh5ZDhlc2FkaHRDR0VsWUFCZ2FuS0tsME1oVkRCTWtLSUZBQURNbEtJRkFLeFoxNWR0QjlBc3dFbE84VnJRQU1CRXJXb3hGS1pFMFFJQUFHWkswUUlBMW0vNFVuZkxuR25ZelZDMEVEQUFNRlcxRU9yekNsT3pMeEVBQUpnblJRc0FXS091TDdzaGhCL05tTWF0Y29vZmhRd0FUTnlKclJaTVRkY1hXeTBBQUdDR0ZDMEFZTDFPelpmR1hlUVUvWndEQUpOWGk2RzJXakExaWhZQUFEQkRpaFlBc0NaZFh3NnRnbVVCam9VTUFNeEZUbkZWbnoyRHFWQzBBQUNBR1ZLMEFJRDFjVnVPMXAzbEZDK2xEQURNekVwZ1RJaHlQZ0FBekpDaUJRQ3NRZGVYNGN2YkhiT2xZYmUyV1FBQWMxU2ZQYlBWZ3NubytySW5EUUFBbUJkRkN3QjRZbDFmdGgxQXN3QW45WjF6QUlBNXN0V0NLZkY4Q0FBQXpJeWlCUUE4dmVISmtDMXpwV0UzOVgxekFJQlpzdFdDaWJIUkFnQUFaa2JSQWdDZVVGMzUrdHBNYWR5UmdBR0FCaWlPTWhVMldnQUF3TXdvV2dEQTB6b3hUeHAza1ZNOEZ6SUFNSGQxcThXdElKbUFuZm9FSlFBQU1CT0tGZ0R3UkxxK0RMZjg5ODJUeHRsbUFRQzBSRkdhcWZCOENBQUF6SWlpQlFBOGdYcjd5T3BoV3ZjdXAzZ3RaUUNnSVNlMldqQVJuZzhCQUlBWlViUUFnS2R4UEt4N05Vc2FkcXRNQkFDMEpxZjRNWVJ3S2xnbXdFWUxBQUNZRVVVTEFIaWtyaSs3dFdnQkxWdlZnd2dBZ05aNFBvUXBVTFFBQUlBWlViUUFnTWNiYnZsdm1TTU51OG9wT29BQUFKcFVuMFk3a3k0YnRsT2ZwQVFBQUdaQTBRSUFIcUhyeS9DTzdtc3pwSEUydGdBQXJmTjhDRk5ncXdVQUFNeUVvZ1VBUEk1Yi9yVHVRMDd4WE1vQVFNdnE1NTBySWJOaEJ3SUFBSUI1VUxRQWdBZnErbklVUW5oaGZqVE9OZ3NBWUNtVXFObTBYUWtBQU1BOEtGb0F3QVBVdDNOOUVVdnIzdFkzeXdFQWx1QjlDT0ZXMG15UXAwTUFBR0FtRkMwQTRHR0dXLzViWmtmRGJwV0pBSUFseVNsK3JHVUwyQlFiRXdFQVlDWVVMUURnbnJxK0RPdGNmelkzR25kY0R4c0FBSlpFMFpTTjZ2cGlxd1VBQU15QW9nVUEzTitwbWRHNGk1eWluM01BWUhGeWlwY2hoQnZKczBHN2hnOEFBTk9uYUFFQTk5RDE1U0NFc0c5bU5HNGxZQUJnd1d5MVlKTnN0QUFBZ0JsUXRBQ0ErM0hMbjlhZDVSVFBwUXdBTE5oNzRiTkJpaFlBQURBRGloWUFNRkxYbCtNUXdvNTUwYkJiMnl3QWdLWExLVjZIRUQ0c2ZRNXNqS2REQUFCZ0JoUXRBR0NFcmkvYkRxQlpnSk42c0FBQXNIUzJXckFwTDB3ZUFBQ21UOUVDQU1ZWlNoWmJaa1hEYnJ4SERnRHdKMFVMTnFicmk2MFdBQUF3Y1lvV0FQQVB1cjRNYitUK2FFNDA3amluK0ZISUFBQy9QeC95MGZNaGJKQ2lCUUFBVEp5aUJRRDhNN2Y4YWQxRlR0R3RUUUNBdi9MNWlFMDVNSGtBQUpnMlJRc0ErSXF1TDRjaGhIMHpvbkhIQWdZQStDK0tGbXpLdHNrREFNQzBLVm9Bd05mWlprSHJ6bktLbDFJR0FQaXIrbnpJbGJHd0FYdUdEZ0FBMDZab0FRQmYwUFZsRlVMWU1SOGFkbXViQlFEQVY1MGFEeHRnb3dVQUFFemMvd2dJQUw3b09vVHcxbmhvMkdXOXFRa0F3T2U5ZCtnTkFBREEzMzF6ZDNkbktBQUFBQUFBQUFBQUkzZzZCQUFBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQStQL3Qyb0VBQUFBQXc2RDdVMTloQU1VUkFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFVR3c3RWJ6eERWYU85cE1BQUFBQVNVVk9SSzVDWUlJPSc7CmZ1bmN0aW9uIGFwcGx5QnJhbmRpbmcoKSB7CiAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnLmJyYW5kLWxvZ28nKS5mb3JFYWNoKChpbWcpID0+IHsgaW1nLnNyYyA9IExPR09fREFUQV9VUkk7IH0pOwogIGNvbnN0IGZhdmljb24gPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmF2aWNvbkxpbmsnKTsKICBpZiAoZmF2aWNvbikgZmF2aWNvbi5ocmVmID0gTE9HT19EQVRBX1VSSTsKfQoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIEFwaSDigJQgdGhpbiBmZXRjaCB3cmFwcGVycyBhcm91bmQgdGhlIFJFU1QgQVBJLgogICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KY29uc3QgQXBpID0gKCgpID0+IHsKICBhc3luYyBmdW5jdGlvbiByZXF1ZXN0KHBhdGgsIG9wdGlvbnMpIHsKICAgIGNvbnN0IHJlcyA9IGF3YWl0IGZldGNoKHBhdGgsIG9wdGlvbnMpOwogICAgbGV0IGJvZHk7CiAgICB0cnkgewogICAgICBib2R5ID0gYXdhaXQgcmVzLmpzb24oKTsKICAgIH0gY2F0Y2ggewogICAgICBib2R5ID0gbnVsbDsKICAgIH0KICAgIGlmIChyZXMuc3RhdHVzID09PSA0MDEgJiYgIXBhdGguc3RhcnRzV2l0aCgnL2FwaS9hdXRoLycpKSB7CiAgICAgIHdpbmRvdy5kaXNwYXRjaEV2ZW50KG5ldyBDdXN0b21FdmVudCgnbHJzOnNpZ25lZC1vdXQnKSk7CiAgICB9CiAgICBpZiAoIXJlcy5vaykgewogICAgICBjb25zdCBtZXNzYWdlID0gKGJvZHkgJiYgYm9keS5lcnJvcikgfHwgYFJlcXVlc3QgZmFpbGVkICgke3Jlcy5zdGF0dXN9KWA7CiAgICAgIHRocm93IG5ldyBFcnJvcihtZXNzYWdlKTsKICAgIH0KICAgIHJldHVybiBib2R5OwogIH0KCiAgZnVuY3Rpb24gcXMocGFyYW1zKSB7CiAgICBjb25zdCB1c3AgPSBuZXcgVVJMU2VhcmNoUGFyYW1zKCk7CiAgICBPYmplY3QuZW50cmllcyhwYXJhbXMgfHwge30pLmZvckVhY2goKFtrLCB2XSkgPT4gewogICAgICBpZiAodiAhPT0gdW5kZWZpbmVkICYmIHYgIT09IG51bGwgJiYgdiAhPT0gJycpIHVzcC5zZXQoaywgdik7CiAgICB9KTsKICAgIGNvbnN0IHMgPSB1c3AudG9TdHJpbmcoKTsKICAgIHJldHVybiBzID8gYD8ke3N9YCA6ICcnOwogIH0KCiAgcmV0dXJuIHsKICAgIGF1dGhNZTogKCkgPT4gcmVxdWVzdCgnL2FwaS9hdXRoL21lJyksCiAgICBhdXRoTG9naW46IChjb2RlKSA9PgogICAgICByZXF1ZXN0KCcvYXBpL2F1dGgvbG9naW4nLCB7IG1ldGhvZDogJ1BPU1QnLCBoZWFkZXJzOiB7ICdDb250ZW50LVR5cGUnOiAnYXBwbGljYXRpb24vanNvbicgfSwgYm9keTogSlNPTi5zdHJpbmdpZnkoeyBjb2RlIH0pIH0pLAogICAgYXV0aExvZ291dDogKCkgPT4gcmVxdWVzdCgnL2FwaS9hdXRoL2xvZ291dCcsIHsgbWV0aG9kOiAnUE9TVCcgfSksCgogICAgZmlsdGVyT3B0aW9uczogKCkgPT4gcmVxdWVzdCgnL2FwaS9hbmFseXRpY3MvZmlsdGVyLW9wdGlvbnMnKSwKICAgIGtwaXM6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL2twaXMke3FzKHBhcmFtcyl9YCksCiAgICBwbGF0Zm9ybUJyZWFrZG93bjogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9hbmFseXRpY3MvcGxhdGZvcm0tYnJlYWtkb3duJHtxcyhwYXJhbXMpfWApLAogICAgY2FtcGFpZ25CcmVha2Rvd246IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL2NhbXBhaWduLWJyZWFrZG93biR7cXMocGFyYW1zKX1gKSwKICAgIGNvbnRlbnRUeXBlQnJlYWtkb3duOiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9jb250ZW50LXR5cGUtYnJlYWtkb3duJHtxcyhwYXJhbXMpfWApLAogICAgbWV0cmljT3B0aW9uczogKHBsYXRmb3JtKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9tZXRyaWMtb3B0aW9ucyR7cXMoeyBwbGF0Zm9ybSB9KX1gKSwKICAgIG1ldHJpY1N1bW1hcnk6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL21ldHJpYy1zdW1tYXJ5JHtxcyhwYXJhbXMpfWApLAogICAgdHJlbmQ6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL3RyZW5kJHtxcyhwYXJhbXMpfWApLAogICAgdG9wUG9zdHM6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL3RvcC1wb3N0cyR7cXMocGFyYW1zKX1gKSwKICAgIGNvbXBhcmU6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL2NvbXBhcmUke3FzKHBhcmFtcyl9YCksCiAgICBtb250aGx5OiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9tb250aGx5JHtxcyhwYXJhbXMpfWApLAogICAgcXVhcnRlcmx5OiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9xdWFydGVybHkke3FzKHBhcmFtcyl9YCksCiAgICB5dGQ6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL3l0ZCR7cXMocGFyYW1zKX1gKSwKICAgIHBsYXRmb3JtUmVwb3J0OiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9wbGF0Zm9ybS1yZXBvcnQke3FzKHBhcmFtcyl9YCksCgogICAgcHJldmlld1VwbG9hZDogKGZpbGUpID0+IHsKICAgICAgY29uc3QgZm9ybSA9IG5ldyBGb3JtRGF0YSgpOwogICAgICBmb3JtLmFwcGVuZCgnZmlsZScsIGZpbGUpOwogICAgICByZXR1cm4gcmVxdWVzdCgnL2FwaS91cGxvYWRzL3ByZXZpZXcnLCB7IG1ldGhvZDogJ1BPU1QnLCBib2R5OiBmb3JtIH0pOwogICAgfSwKICAgIGNvbW1pdFVwbG9hZDogKHBheWxvYWQpID0+CiAgICAgIHJlcXVlc3QoJy9hcGkvdXBsb2Fkcy9jb21taXQnLCB7CiAgICAgICAgbWV0aG9kOiAnUE9TVCcsCiAgICAgICAgaGVhZGVyczogeyAnQ29udGVudC1UeXBlJzogJ2FwcGxpY2F0aW9uL2pzb24nIH0sCiAgICAgICAgYm9keTogSlNPTi5zdHJpbmdpZnkocGF5bG9hZCksCiAgICAgIH0pLAogICAgdXBsb2FkSGlzdG9yeTogKCkgPT4gcmVxdWVzdCgnL2FwaS91cGxvYWRzL2hpc3RvcnknKSwKICAgIHVwbG9hZEVycm9yczogKGlkKSA9PiByZXF1ZXN0KGAvYXBpL3VwbG9hZHMvJHtpZH0vZXJyb3JzYCksCiAgICB1cGxvYWRSYXdSb3dzOiAoaWQpID0+IHJlcXVlc3QoYC9hcGkvdXBsb2Fkcy8ke2lkfS9yYXctcm93c2ApLAoKICAgIGxpc3RSZWNvcmRzOiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL3JlY29yZHMke3FzKHBhcmFtcyl9YCksCiAgICByZWNvcmRzVGFibGU6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvcmVjb3Jkcy90YWJsZSR7cXMocGFyYW1zKX1gKSwKICAgIGdldFJlY29yZDogKGlkKSA9PiByZXF1ZXN0KGAvYXBpL3JlY29yZHMvJHtpZH1gKSwKICAgIHVwZGF0ZVJlY29yZDogKGlkLCB2YWx1ZXMpID0+CiAgICAgIHJlcXVlc3QoYC9hcGkvcmVjb3Jkcy8ke2lkfWAsIHsKICAgICAgICBtZXRob2Q6ICdQVVQnLAogICAgICAgIGhlYWRlcnM6IHsgJ0NvbnRlbnQtVHlwZSc6ICdhcHBsaWNhdGlvbi9qc29uJyB9LAogICAgICAgIGJvZHk6IEpTT04uc3RyaW5naWZ5KHsgdmFsdWVzIH0pLAogICAgICB9KSwKICAgIGRlbGV0ZVJlY29yZFBvc3Q6IChwb3N0SWQpID0+IHJlcXVlc3QoYC9hcGkvcmVjb3Jkcy9wb3N0LyR7cG9zdElkfWAsIHsgbWV0aG9kOiAnREVMRVRFJyB9KSwKICAgIGRlbGV0ZVJlY29yZFBsYXRmb3JtOiAocG9zdElkLCBwbGF0Zm9ybSkgPT4KICAgICAgcmVxdWVzdChgL2FwaS9yZWNvcmRzL3Bvc3QvJHtwb3N0SWR9L3BsYXRmb3JtLyR7cGxhdGZvcm19YCwgeyBtZXRob2Q6ICdERUxFVEUnIH0pLAogICAgZHVwbGljYXRlUmVjb3Jkc1ByZXZpZXc6ICgpID0+IHJlcXVlc3QoJy9hcGkvcmVjb3Jkcy9kdXBsaWNhdGVzJyksCiAgICByZXNvbHZlRHVwbGljYXRlUmVjb3JkczogKCkgPT4gcmVxdWVzdCgnL2FwaS9yZWNvcmRzL2R1cGxpY2F0ZXMvcmVzb2x2ZScsIHsgbWV0aG9kOiAnUE9TVCcgfSksCiAgICB3aXBlVXBsb2FkZWRSZWNvcmRzOiAoKSA9PgogICAgICByZXF1ZXN0KCcvYXBpL3JlY29yZHMvd2lwZScsIHsKICAgICAgICBtZXRob2Q6ICdQT1NUJywKICAgICAgICBoZWFkZXJzOiB7ICdDb250ZW50LVR5cGUnOiAnYXBwbGljYXRpb24vanNvbicgfSwKICAgICAgICBib2R5OiBKU09OLnN0cmluZ2lmeSh7IGNvbmZpcm06ICdERUxFVEUnIH0pLAogICAgICB9KSwKCiAgICByZXN0b3JlQmFja3VwOiAoZm9ybSkgPT4gcmVxdWVzdCgnL2FwaS9iYWNrdXAvcmVzdG9yZScsIHsgbWV0aG9kOiAnUE9TVCcsIGJvZHk6IGZvcm0gfSksCgogICAgbGlzdEZvbGxvd2VyczogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9mb2xsb3dlcnMke3FzKHBhcmFtcyl9YCksCiAgICBmb2xsb3dlcnNHcm93dGg6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvZm9sbG93ZXJzL2dyb3d0aCR7cXMocGFyYW1zKX1gKSwKICAgIGZvbGxvd2Vyc0twaXM6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvZm9sbG93ZXJzL2twaXMke3FzKHBhcmFtcyl9YCksCiAgICBzYXZlRm9sbG93ZXJzOiAocGF5bG9hZCkgPT4KICAgICAgcmVxdWVzdCgnL2FwaS9mb2xsb3dlcnMnLCB7IG1ldGhvZDogJ1BPU1QnLCBoZWFkZXJzOiB7ICdDb250ZW50LVR5cGUnOiAnYXBwbGljYXRpb24vanNvbicgfSwgYm9keTogSlNPTi5zdHJpbmdpZnkocGF5bG9hZCkgfSksCiAgICB1cGRhdGVGb2xsb3dlcnM6IChpZCwgcGF5bG9hZCkgPT4KICAgICAgcmVxdWVzdChgL2FwaS9mb2xsb3dlcnMvJHtpZH1gLCB7IG1ldGhvZDogJ1BVVCcsIGhlYWRlcnM6IHsgJ0NvbnRlbnQtVHlwZSc6ICdhcHBsaWNhdGlvbi9qc29uJyB9LCBib2R5OiBKU09OLnN0cmluZ2lmeShwYXlsb2FkKSB9KSwKICAgIGRlbGV0ZUZvbGxvd2VyczogKGlkKSA9PiByZXF1ZXN0KGAvYXBpL2ZvbGxvd2Vycy8ke2lkfWAsIHsgbWV0aG9kOiAnREVMRVRFJyB9KSwKICB9Owp9KSgpOwoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIFN0YXRlIC8gRm9ybWF0IC8gVG9hc3Qg4oCUIHNoYXJlZCBhcHAgc3RhdGUgKyBzbWFsbCB1dGlsaXRpZXMuCiAgID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwpjb25zdCBTdGF0ZSA9ICgoKSA9PiB7CiAgY29uc3QgdG9kYXkgPSBuZXcgRGF0ZSgpOwogIGNvbnN0IGlzbyA9IChkKSA9PiBkLnRvSVNPU3RyaW5nKCkuc2xpY2UoMCwgMTApOwogIGNvbnN0IHRoaXJ0eURheXNBZ28gPSBuZXcgRGF0ZSh0b2RheSk7CiAgdGhpcnR5RGF5c0Fnby5zZXREYXRlKHRoaXJ0eURheXNBZ28uZ2V0RGF0ZSgpIC0gMjkpOwoKICBjb25zdCBmaWx0ZXJzID0gewogICAgZGF0ZUZyb206IGlzbyh0aGlydHlEYXlzQWdvKSwKICAgIGRhdGVUbzogaXNvKHRvZGF5KSwKICAgIHBsYXRmb3JtOiAnYWxsJywKICAgIGNhbXBhaWduVHlwZTogJ2FsbCcsCiAgICBjb250ZW50VHlwZTogJ2FsbCcsCiAgfTsKCiAgY29uc3QgbGlzdGVuZXJzID0gW107CgogIHJldHVybiB7CiAgICBnZXRGaWx0ZXJzOiAoKSA9PiAoeyAuLi5maWx0ZXJzIH0pLAogICAgc2V0RmlsdGVycyhwYXJ0aWFsKSB7CiAgICAgIE9iamVjdC5hc3NpZ24oZmlsdGVycywgcGFydGlhbCk7CiAgICAgIGxpc3RlbmVycy5mb3JFYWNoKChmbikgPT4gZm4odGhpcy5nZXRGaWx0ZXJzKCkpKTsKICAgIH0sCiAgICBvbkNoYW5nZShmbikgewogICAgICBsaXN0ZW5lcnMucHVzaChmbik7CiAgICB9LAogIH07Cn0pKCk7Cgpjb25zdCBGb3JtYXQgPSB7CiAgbnVtYmVyKG4pIHsKICAgIGlmIChuID09PSBudWxsIHx8IG4gPT09IHVuZGVmaW5lZCkgcmV0dXJuICfigJQnOwogICAgcmV0dXJuIE1hdGgucm91bmQobikudG9Mb2NhbGVTdHJpbmcoJ2VuLVVTJyk7CiAgfSwKICBjb21wYWN0KG4pIHsKICAgIGlmIChuID09PSBudWxsIHx8IG4gPT09IHVuZGVmaW5lZCkgcmV0dXJuICfigJQnOwogICAgY29uc3QgYWJzID0gTWF0aC5hYnMobik7CiAgICBpZiAoYWJzID49IDFfMDAwXzAwMCkgcmV0dXJuIGAkeyhuIC8gMV8wMDBfMDAwKS50b0ZpeGVkKDEpLnJlcGxhY2UoL1wuMCQvLCAnJyl9TWA7CiAgICBpZiAoYWJzID49IDFfMDAwKSByZXR1cm4gYCR7KG4gLyAxXzAwMCkudG9GaXhlZCgxKS5yZXBsYWNlKC9cLjAkLywgJycpfUtgOwogICAgcmV0dXJuIGAke01hdGgucm91bmQobil9YDsKICB9LAogIC8qKiBEYXNoYm9hcmQtd2lkZSAicHJvZmVzc2lvbmFsIiBudW1iZXIgZm9ybWF0OiBwbGFpbiB1bmRlciAxLDAwMDsgY29tbWEtZ3JvdXBlZAogICAgICB1cCB0byAxMCwwMDA7IGFiYnJldmlhdGVkIChLL00pIGJleW9uZCB0aGF0IOKAlCBlLmcuIDg1MCwgMSwyNTAsIDEyLjVLLCAxNTZLLCAxLjI1TS4gKi8KICBzbWFydChuKSB7CiAgICBpZiAobiA9PT0gbnVsbCB8fCBuID09PSB1bmRlZmluZWQpIHJldHVybiAn4oCUJzsKICAgIGNvbnN0IGFicyA9IE1hdGguYWJzKG4pOwogICAgaWYgKGFicyA8IDEwMDApIHJldHVybiBgJHtNYXRoLnJvdW5kKG4pfWA7CiAgICBpZiAoYWJzIDwgMTAwMDApIHJldHVybiBNYXRoLnJvdW5kKG4pLnRvTG9jYWxlU3RyaW5nKCdlbi1VUycpOwogICAgaWYgKGFicyA8IDFfMDAwXzAwMCkgcmV0dXJuIGAkeyhuIC8gMTAwMCkudG9GaXhlZCgxKS5yZXBsYWNlKC9cLjAkLywgJycpfUtgOwogICAgcmV0dXJuIGAkeyhuIC8gMV8wMDBfMDAwKS50b0ZpeGVkKDIpLnJlcGxhY2UoL1wuPzArJC8sICcnKX1NYDsKICB9LAogIHBlcmNlbnQobikgewogICAgaWYgKG4gPT09IG51bGwgfHwgbiA9PT0gdW5kZWZpbmVkKSByZXR1cm4gJ+KAlCc7CiAgICByZXR1cm4gYCR7TnVtYmVyKG4pLnRvRml4ZWQoMSkucmVwbGFjZSgvXC4wJC8sICcnKX0lYDsKICB9LAogIHBjdChuKSB7CiAgICBpZiAobiA9PT0gbnVsbCB8fCBuID09PSB1bmRlZmluZWQpIHJldHVybiAn4oCUJzsKICAgIGNvbnN0IHNpZ24gPSBuID4gMCA/ICcrJyA6ICcnOwogICAgcmV0dXJuIGAke3NpZ259JHtuLnRvRml4ZWQoMSl9JWA7CiAgfSwKICBkYXRlKGlzb18pIHsKICAgIGlmICghaXNvXykgcmV0dXJuICfigJQnOwogICAgY29uc3QgW3ksIG0sIGRdID0gaXNvXy5zcGxpdCgnLScpLm1hcChOdW1iZXIpOwogICAgcmV0dXJuIG5ldyBEYXRlKHksIG0gLSAxLCBkKS50b0xvY2FsZURhdGVTdHJpbmcoJ2VuLVVTJywgeyBtb250aDogJ3Nob3J0JywgZGF5OiAnbnVtZXJpYycsIHllYXI6ICdudW1lcmljJyB9KTsKICB9LAogIGR1cmF0aW9uKHNlY29uZHMpIHsKICAgIGlmIChzZWNvbmRzID09PSBudWxsIHx8IHNlY29uZHMgPT09IHVuZGVmaW5lZCkgcmV0dXJuICfigJQnOwogICAgY29uc3QgcyA9IE1hdGgucm91bmQoc2Vjb25kcyk7CiAgICBpZiAocyA8IDYwKSByZXR1cm4gYCR7c31zYDsKICAgIGlmIChzIDwgMzYwMCkgcmV0dXJuIGAke01hdGguZmxvb3IocyAvIDYwKX1tICR7cyAlIDYwfXNgOwogICAgY29uc3QgaCA9IE1hdGguZmxvb3IocyAvIDM2MDApOwogICAgY29uc3QgbSA9IE1hdGgucm91bmQoKHMgJSAzNjAwKSAvIDYwKTsKICAgIHJldHVybiBgJHtofWggJHttfW1gOwogIH0sCiAgZGVsdGFDbGFzcyhuKSB7CiAgICBpZiAobiA9PT0gbnVsbCB8fCBuID09PSB1bmRlZmluZWQpIHJldHVybiAnZmxhdCc7CiAgICBpZiAobiA+IDAuNSkgcmV0dXJuICd1cCc7CiAgICBpZiAobiA8IC0wLjUpIHJldHVybiAnZG93bic7CiAgICByZXR1cm4gJ2ZsYXQnOwogIH0sCn07Cgpjb25zdCBUb2FzdCA9IHsKICBzaG93KG1lc3NhZ2UsIHR5cGUgPSAnc3VjY2VzcycpIHsKICAgIGNvbnN0IHJvb3QgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndG9hc3RSb290Jyk7CiAgICBjb25zdCBlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZWwuY2xhc3NOYW1lID0gYHRvYXN0ICR7dHlwZX1gOwogICAgZWwudGV4dENvbnRlbnQgPSBtZXNzYWdlOwogICAgcm9vdC5hcHBlbmRDaGlsZChlbCk7CiAgICBzZXRUaW1lb3V0KCgpID0+IGVsLnJlbW92ZSgpLCA1MDAwKTsKICB9LAp9OwoKLyoqIFNhZmVseSBidWlsZHMgRE9NIHRleHQgbm9kZXMgZm9yIHVudHJ1c3RlZCBzdHJpbmdzIChjYXB0aW9ucywgZmlsZW5hbWVzLCBwbGF0Zm9ybSBsYWJlbHMgZnJvbSBkYXRhKS4gKi8KZnVuY3Rpb24gdGV4dEVsKHRhZywgdGV4dCwgY2xhc3NOYW1lKSB7CiAgY29uc3QgZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KHRhZyk7CiAgaWYgKGNsYXNzTmFtZSkgZWwuY2xhc3NOYW1lID0gY2xhc3NOYW1lOwogIGVsLmFwcGVuZENoaWxkKGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKHRleHQgPz8gJycpKTsKICByZXR1cm4gZWw7Cn0KCi8qKiBBIHByZW1pdW0gZW1wdHkgc3RhdGU6IGljb24gKyBleHBsYW5hdGlvbiArIG9wdGlvbmFsIGFjdGlvbiwgaW5zdGVhZCBvZiBhIGJsYW5rIGFyZWEuCiAgICBJY29ucyByZW5kZXIgdmlhIHRoZSBwYWdlLXdpZGUgTXV0YXRpb25PYnNlcnZlciB0aGF0IGNhbGxzIGx1Y2lkZS5jcmVhdGVJY29ucygpIChzZWUgYm9vdHN0cmFwKS4gKi8KZnVuY3Rpb24gZW1wdHlTdGF0ZSh7IGljb24gPSAnaW5ib3gnLCB0aXRsZSwgbWVzc2FnZSwgYWN0aW9uTGFiZWwsIG9uQWN0aW9uIH0pIHsKICBjb25zdCB3cmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgd3JhcC5jbGFzc05hbWUgPSAnZW1wdHktc3RhdGUnOwogIGNvbnN0IGljb25XcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgaWNvbldyYXAuY2xhc3NOYW1lID0gJ2VtcHR5LWljb24nOwogIGljb25XcmFwLmlubmVySFRNTCA9IGA8aSBkYXRhLWx1Y2lkZT0iJHtpY29ufSIgc3R5bGU9IndpZHRoOjIycHg7aGVpZ2h0OjIycHg7Ij48L2k+YDsKICB3cmFwLmFwcGVuZENoaWxkKGljb25XcmFwKTsKICBpZiAodGl0bGUpIHdyYXAuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCB0aXRsZSwgJ2VtcHR5LXRpdGxlJykpOwogIGlmIChtZXNzYWdlKSB3cmFwLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgbWVzc2FnZSwgJ2VtcHR5LW1lc3NhZ2UnKSk7CiAgaWYgKGFjdGlvbkxhYmVsICYmIG9uQWN0aW9uKSB7CiAgICBjb25zdCBidG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIGJ0bi5jbGFzc05hbWUgPSAnYnRuIHByaW1hcnknOwogICAgYnRuLnRleHRDb250ZW50ID0gYWN0aW9uTGFiZWw7CiAgICBidG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBvbkFjdGlvbik7CiAgICB3cmFwLmFwcGVuZENoaWxkKGJ0bik7CiAgfQogIHJldHVybiB3cmFwOwp9CgovKiogQSA8YnV0dG9uPiB3aXRoIGEgc21hbGwgbGVhZGluZyBMdWNpZGUgaWNvbiBiZWZvcmUgaXRzIGxhYmVsIChsYWJlbCBpcyBhbHdheXMgYSBzdGF0aWMsIGRldmVsb3Blci1zdXBwbGllZCBzdHJpbmcgYXQgY2FsbCBzaXRlcywgbmV2ZXIgdXNlciBkYXRhIOKAlCBpbnNlcnRlZCB2aWEgY3JlYXRlVGV4dE5vZGUgcmVnYXJkbGVzcykuICovCmZ1bmN0aW9uIGljb25CdG4oY2xhc3NOYW1lLCBpY29uTmFtZSwgbGFiZWwpIHsKICBjb25zdCBidG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICBidG4uY2xhc3NOYW1lID0gY2xhc3NOYW1lOwogIGNvbnN0IGljb24gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpJyk7CiAgaWNvbi5zZXRBdHRyaWJ1dGUoJ2RhdGEtbHVjaWRlJywgaWNvbk5hbWUpOwogIGljb24uc3R5bGUud2lkdGggPSAnMTNweCc7CiAgaWNvbi5zdHlsZS5oZWlnaHQgPSAnMTNweCc7CiAgYnRuLmFwcGVuZENoaWxkKGljb24pOwogIGJ0bi5hcHBlbmRDaGlsZChkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShgICR7bGFiZWx9YCkpOwogIHJldHVybiBidG47Cn0KCi8qKiBTaGltbWVyaW5nIHBsYWNlaG9sZGVycyBzaG93biB0aGUgaW5zdGFudCBhIHNlY3Rpb24gc3RhcnRzIGxvYWRpbmcsIHN3YXBwZWQgZm9yIHJlYWwKICAgIGNvbnRlbnQgKG9yIGFuIGVtcHR5IHN0YXRlKSBvbmNlIHRoZSBmZXRjaCByZXNvbHZlcyDigJQgbm8gYmxhbmsgYXJlYXMgd2hpbGUgd2FpdGluZy4gKi8KZnVuY3Rpb24gc2tlbGV0b25TdGF0R3JpZChjb3VudCA9IDYpIHsKICBjb25zdCBncmlkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgZ3JpZC5jbGFzc05hbWUgPSAnc2tlbGV0b24tc3RhdC1ncmlkJzsKICBmb3IgKGxldCBpID0gMDsgaSA8IGNvdW50OyBpICs9IDEpIHsKICAgIGNvbnN0IHRpbGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRpbGUuY2xhc3NOYW1lID0gJ3NrZWxldG9uIHNrZWxldG9uLXRpbGUnOwogICAgZ3JpZC5hcHBlbmRDaGlsZCh0aWxlKTsKICB9CiAgcmV0dXJuIGdyaWQ7Cn0KZnVuY3Rpb24gc2tlbGV0b25DaGFydCgpIHsKICBjb25zdCBkaXYgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICBkaXYuY2xhc3NOYW1lID0gJ3NrZWxldG9uIHNrZWxldG9uLWNoYXJ0JzsKICByZXR1cm4gZGl2Owp9CmZ1bmN0aW9uIHNrZWxldG9uUm93cyhjb3VudCA9IDYpIHsKICBjb25zdCB3cmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgZm9yIChsZXQgaSA9IDA7IGkgPCBjb3VudDsgaSArPSAxKSB7CiAgICBjb25zdCByb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHJvdy5jbGFzc05hbWUgPSAnc2tlbGV0b24gc2tlbGV0b24tcm93JzsKICAgIHdyYXAuYXBwZW5kQ2hpbGQocm93KTsKICB9CiAgcmV0dXJuIHdyYXA7Cn0KCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBTaGFyZWQgYW5pbWF0aW9uIHByaW1pdGl2ZXMg4oCUIGEgY291bnQtdXAgZm9yIEtQSSBudW1iZXJzIGFuZCBhCiAgIENTUyB3aWR0aC10cmFuc2l0aW9uIGJhciwgYm90aCByZXVzZWQgYWNyb3NzIHRoZSBEYXNoYm9hcmQgYW5kCiAgIENvbXBhcmlzb25zIHBhZ2VzLiBCb3RoIHJlc3BlY3QgcHJlZmVycy1yZWR1Y2VkLW1vdGlvbiAoZ3VhcmRlZAogICBpbiBDU1MsIHNlZSAuYmFyLWZpbGwgLyB0aGUgYW5pbWF0ZUNvdW50IGR1cmF0aW9uIGNoZWNrIGJlbG93KS4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmNvbnN0IFBSRUZFUlNfUkVEVUNFRF9NT1RJT04gPSB3aW5kb3cubWF0Y2hNZWRpYSAmJiB3aW5kb3cubWF0Y2hNZWRpYSgnKHByZWZlcnMtcmVkdWNlZC1tb3Rpb246IHJlZHVjZSknKS5tYXRjaGVzOwoKLyoqIFNocmlua3MgYGVsYCdzIGZvbnQgc2l6ZSBqdXN0IGVub3VnaCBmb3IgaXRzIGN1cnJlbnQgdGV4dCB0byBmaXQgaXRzIG93biB3aWR0aCDigJQgYSBLUEkgdGlsZSdzIGJveCBpcyBhIGZpeGVkIHNpemUsIGJ1dCB0aGUgdmFsdWUgaW5zaWRlIGl0IGlzbid0IChhIGZvbGxvd2VyIGNvdW50IGNhbiBiZSAiMCIgb3IgIjEsMDQ4LDU3NiIpLCBzbyBhIHNpbmdsZSBmaXhlZCBmb250LXNpemUgd2lsbCBldmVudHVhbGx5IG92ZXJmbG93LiBSZXNldHMgdG8gdGhlIENTUy1kZWZpbmVkIHNpemUgZmlyc3QsIHRoZW4gc3RlcHMgZG93biBieSAxcHggYXQgYSB0aW1lIHVudGlsIGl0IGZpdHMgb3IgaGl0cyBgbWluU2l6ZWAuICovCmZ1bmN0aW9uIGZpdFN0YXRWYWx1ZShlbCwgbWluU2l6ZSA9IDE4KSB7CiAgaWYgKCFlbCkgcmV0dXJuOwogIGVsLnN0eWxlLmZvbnRTaXplID0gJyc7CiAgY29uc3QgbWF4V2lkdGggPSBlbC5jbGllbnRXaWR0aDsKICBpZiAoIW1heFdpZHRoKSByZXR1cm47CiAgbGV0IHNpemUgPSBwYXJzZUZsb2F0KGdldENvbXB1dGVkU3R5bGUoZWwpLmZvbnRTaXplKTsKICB3aGlsZSAoZWwuc2Nyb2xsV2lkdGggPiBtYXhXaWR0aCAmJiBzaXplID4gbWluU2l6ZSkgewogICAgc2l6ZSAtPSAxOwogICAgZWwuc3R5bGUuZm9udFNpemUgPSBgJHtzaXplfXB4YDsKICB9Cn0KCi8qKiBBbmltYXRlcyBhIG51bWJlciBmcm9tIGBmcm9tYCB0byBgdG9gIGluc2lkZSBgZWxgIG92ZXIgYGR1cmF0aW9uYG1zLCBmb3JtYXR0aW5nIGVhY2ggZnJhbWUgd2l0aCBgZm9ybWF0YCAoZGVmYXVsdHMgdG8gYSBwbGFpbiByb3VuZGVkIGludGVnZXIpLiBTa2lwcyBzdHJhaWdodCB0byB0aGUgZmluYWwgdmFsdWUgdW5kZXIgcHJlZmVycy1yZWR1Y2VkLW1vdGlvbi4gU2hyaW5rcyB0aGUgZm9udCB0byBmaXQgb25jZSB0aGUgZmluYWwgdmFsdWUgbGFuZHMsIHNpbmNlIHRoZSBhbmltYXRlZCBkaWdpdHMgY2FuIGJlIGEgZGlmZmVyZW50IHdpZHRoIHRoYW4gdGhlIHNldHRsZWQgdmFsdWUg4oCUIGRlZmVycmVkIGEgZnJhbWUgYmVjYXVzZSBgZWxgIGlzIHR5cGljYWxseSBzdGlsbCBkZXRhY2hlZCBmcm9tIHRoZSBkb2N1bWVudCAobWlkLWNvbnN0cnVjdGlvbiBieSBpdHMgY2FsbGVyKSB3aGVuIGFuaW1hdGVDb3VudCBpcyBmaXJzdCBpbnZva2VkLCBhbmQgY2xpZW50V2lkdGggcmVhZHMgMCB1bnRpbCBpdCdzIGFjdHVhbGx5IGF0dGFjaGVkIGFuZCBsYWlkIG91dC4gKi8KZnVuY3Rpb24gYW5pbWF0ZUNvdW50KGVsLCBmcm9tLCB0bywgZHVyYXRpb24gPSA5MDAsIGZvcm1hdCkgewogIGlmICghZWwpIHJldHVybjsKICBjb25zdCBmbXQgPSBmb3JtYXQgfHwgKCh2KSA9PiBNYXRoLnJvdW5kKHYpLnRvTG9jYWxlU3RyaW5nKCdlbi1VUycpKTsKICBpZiAoUFJFRkVSU19SRURVQ0VEX01PVElPTiB8fCBmcm9tID09PSB0byB8fCAhTnVtYmVyLmlzRmluaXRlKGZyb20pIHx8ICFOdW1iZXIuaXNGaW5pdGUodG8pKSB7CiAgICBlbC50ZXh0Q29udGVudCA9IGZtdCh0byk7CiAgICByZXF1ZXN0QW5pbWF0aW9uRnJhbWUoKCkgPT4gZml0U3RhdFZhbHVlKGVsKSk7CiAgICByZXR1cm47CiAgfQogIGNvbnN0IHN0YXJ0ID0gcGVyZm9ybWFuY2Uubm93KCk7CiAgZnVuY3Rpb24gdGljayhub3cpIHsKICAgIGNvbnN0IGVsYXBzZWQgPSBub3cgLSBzdGFydDsKICAgIGNvbnN0IHByb2dyZXNzID0gTWF0aC5taW4oMSwgZWxhcHNlZCAvIGR1cmF0aW9uKTsKICAgIGNvbnN0IGVhc2VkID0gMSAtIE1hdGgucG93KDEgLSBwcm9ncmVzcywgMyk7IC8vIGVhc2VPdXRDdWJpYwogICAgZWwudGV4dENvbnRlbnQgPSBmbXQoZnJvbSArICh0byAtIGZyb20pICogZWFzZWQpOwogICAgaWYgKHByb2dyZXNzIDwgMSkgewogICAgICByZXF1ZXN0QW5pbWF0aW9uRnJhbWUodGljayk7CiAgICB9IGVsc2UgewogICAgICBmaXRTdGF0VmFsdWUoZWwpOwogICAgfQogIH0KICByZXF1ZXN0QW5pbWF0aW9uRnJhbWUodGljayk7Cn0KCi8qKiBBIGxhYmVsZWQgaG9yaXpvbnRhbCBiYXIgdGhhdCBhbmltYXRlcyBpdHMgd2lkdGggaW4gb24gaW5zZXJ0aW9uIOKAlCB1c2VkIGZvciB0aGUgQ29tcGFyaXNvbnMgcGFnZSdzIHBhaXJlZCBSYW5nZSBBL0IgYmFycy4gYHZhbHVlYC9gbWF4YCBkcml2ZSB0aGUgZmlsbCBwZXJjZW50YWdlOyBgY29sb3JWYXJgIGlzIGEgQ1NTIGN1c3RvbSBwcm9wZXJ0eSBuYW1lIChlLmcuICctLXNlcmllcy0xJykuICovCmZ1bmN0aW9uIGJ1aWxkQmFyKHsgbGFiZWwsIHZhbHVlLCBtYXgsIGNvbG9yVmFyLCBmb3JtYXRWYWx1ZSB9KSB7CiAgY29uc3Qgcm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgcm93LmNsYXNzTmFtZSA9ICdiYXItcm93JzsKICBjb25zdCBsYWJlbEVsID0gdGV4dEVsKCdkaXYnLCBsYWJlbCwgJ2Jhci1sYWJlbCcpOwogIGNvbnN0IHRyYWNrID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgdHJhY2suY2xhc3NOYW1lID0gJ2Jhci10cmFjayc7CiAgY29uc3QgZmlsbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogIGZpbGwuY2xhc3NOYW1lID0gJ2Jhci1maWxsJzsKICBmaWxsLnN0eWxlLmJhY2tncm91bmQgPSBjb2xvclZhciA/IGB2YXIoJHtjb2xvclZhcn0pYCA6ICd2YXIoLS1zZXJpZXMtMSknOwogIHRyYWNrLmFwcGVuZENoaWxkKGZpbGwpOwogIGNvbnN0IHZhbHVlRWwgPSB0ZXh0RWwoJ2RpdicsIGZvcm1hdFZhbHVlID8gZm9ybWF0VmFsdWUodmFsdWUpIDogU3RyaW5nKHZhbHVlKSwgJ2Jhci12YWx1ZScpOwogIHJvdy5hcHBlbmQobGFiZWxFbCwgdHJhY2ssIHZhbHVlRWwpOwogIGNvbnN0IHBjdCA9IG1heCA+IDAgPyBNYXRoLm1pbigxMDAsIE1hdGgucm91bmQoKHZhbHVlIC8gbWF4KSAqIDEwMDApIC8gMTApIDogMDsKICByZXF1ZXN0QW5pbWF0aW9uRnJhbWUoKCkgPT4geyBmaWxsLnN0eWxlLndpZHRoID0gYCR7cGN0fSVgOyB9KTsKICByZXR1cm4gcm93Owp9CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgU2hhcmVkIHRhYmxlIHRvb2xiYXIgcGllY2VzIOKAlCBzZWFyY2ggYm94LCBjbGllbnQtc2lkZSBwYWdlciwKICAgYW5kIENTVi9YTFNYIGV4cG9ydCDigJQgcmV1c2VkIGJ5IHRoZSBGb2xsb3dlcnMgRGF0YSBhbmQgVXBsb2FkCiAgIEhpc3RvcnkgdGFicyAoYm90aCBsb2FkIHRoZWlyIGZ1bGwgZGF0YXNldCBvbmNlIGFuZCBzZWFyY2gvCiAgIHNvcnQvcGFnaW5hdGUgaXQgaW4gdGhlIGJyb3dzZXIsIHVubGlrZSBEYXRhIFJlY29yZHMgd2hpY2ggaXMKICAgc2VydmVyLXBhZ2luYXRlZCkuIENTViBuZWVkcyBubyBzZXJ2ZXIgcm91bmQgdHJpcCBhdCBhbGw7IFhMU1gKICAgZ29lcyB0aHJvdWdoIFBPU1QgL2FwaS9leHBvcnQgc28gZXhjZWxqcyAoYWxyZWFkeSBhIGRlcGVuZGVuY3kpCiAgIGNhbiBnZW5lcmF0ZSBhIHJlYWwgLnhsc3ggd2l0aG91dCBhZGRpbmcgYSBjbGllbnQtc2lkZSBsaWJyYXJ5LgogICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KZnVuY3Rpb24gYnVpbGRTZWFyY2hCb3goeyBwbGFjZWhvbGRlciwgdmFsdWUsIG9uQ2hhbmdlIH0pIHsKICBjb25zdCB3cmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgd3JhcC5jbGFzc05hbWUgPSAncmVjb3Jkcy1zZWFyY2gnOwogIGNvbnN0IGlucHV0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsKICBpbnB1dC50eXBlID0gJ3NlYXJjaCc7CiAgaW5wdXQucGxhY2Vob2xkZXIgPSBwbGFjZWhvbGRlcjsKICBpbnB1dC52YWx1ZSA9IHZhbHVlIHx8ICcnOwogIGxldCBkZWJvdW5jZSA9IG51bGw7CiAgaW5wdXQuYWRkRXZlbnRMaXN0ZW5lcignaW5wdXQnLCAoKSA9PiB7CiAgICBjbGVhclRpbWVvdXQoZGVib3VuY2UpOwogICAgZGVib3VuY2UgPSBzZXRUaW1lb3V0KCgpID0+IG9uQ2hhbmdlKGlucHV0LnZhbHVlKSwgMzAwKTsKICB9KTsKICB3cmFwLmFwcGVuZENoaWxkKGlucHV0KTsKICByZXR1cm4gd3JhcDsKfQoKLyoqIFNsaWNlcyBhbiBhbHJlYWR5LWxvYWRlZCwgYWxyZWFkeS1maWx0ZXJlZC9zb3J0ZWQgYXJyYXkgZm9yIGNsaWVudC1zaWRlIHBhZ2luYXRpb24g4oCUIHRoZSBjb3VudGVycGFydCB0byB0aGUgc2VydmVyLXNpZGUgcGFnaW5hdGUoKSBpbiBhcHAuanMsIGZvciB0YWJsZXMgdGhhdCBkb24ndCBoYXZlIGEgcGFnaW5hdGVkIGVuZHBvaW50LiAqLwpmdW5jdGlvbiBwYWdpbmF0ZUNsaWVudFNpZGUocm93cywgcGFnZSwgcGFnZVNpemUpIHsKICBjb25zdCB0b3RhbFBhZ2VzID0gTWF0aC5tYXgoMSwgTWF0aC5jZWlsKHJvd3MubGVuZ3RoIC8gcGFnZVNpemUpKTsKICBjb25zdCBzYWZlUGFnZSA9IE1hdGgubWluKE1hdGgubWF4KDEsIHBhZ2UpLCB0b3RhbFBhZ2VzKTsKICBjb25zdCBzdGFydCA9IChzYWZlUGFnZSAtIDEpICogcGFnZVNpemU7CiAgcmV0dXJuIHsgcGFnZVJvd3M6IHJvd3Muc2xpY2Uoc3RhcnQsIHN0YXJ0ICsgcGFnZVNpemUpLCB0b3RhbFBhZ2VzLCBzYWZlUGFnZSwgdG90YWw6IHJvd3MubGVuZ3RoIH07Cn0KCmZ1bmN0aW9uIGJ1aWxkUGFnZXIoeyBwYWdlLCB0b3RhbFBhZ2VzLCB0b3RhbCwgb25QcmV2LCBvbk5leHQgfSkgewogIGNvbnN0IHBhZ2VyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgcGFnZXIuY2xhc3NOYW1lID0gJ3BhZ2luYXRpb24tcm93JzsKICBjb25zdCBwcmV2QnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgcHJldkJ0bi5jbGFzc05hbWUgPSAnYnRuJzsgcHJldkJ0bi50eXBlID0gJ2J1dHRvbic7IHByZXZCdG4udGV4dENvbnRlbnQgPSAnUHJldmlvdXMnOwogIHByZXZCdG4uZGlzYWJsZWQgPSBwYWdlIDw9IDE7CiAgcHJldkJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIG9uUHJldik7CiAgY29uc3QgbmV4dEJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogIG5leHRCdG4uY2xhc3NOYW1lID0gJ2J0bic7IG5leHRCdG4udHlwZSA9ICdidXR0b24nOyBuZXh0QnRuLnRleHRDb250ZW50ID0gJ05leHQnOwogIG5leHRCdG4uZGlzYWJsZWQgPSBwYWdlID49IHRvdGFsUGFnZXM7CiAgbmV4dEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIG9uTmV4dCk7CiAgcGFnZXIuYXBwZW5kKHByZXZCdG4sIHRleHRFbCgnc3BhbicsIGBQYWdlICR7cGFnZX0gb2YgJHt0b3RhbFBhZ2VzfSDigJQgJHt0b3RhbH0gcmVjb3JkKHMpYCksIG5leHRCdG4pOwogIHJldHVybiBwYWdlcjsKfQoKZnVuY3Rpb24gZG93bmxvYWRCbG9iKGZpbGVuYW1lLCBtaW1lVHlwZSwgY29udGVudCkgewogIGNvbnN0IGJsb2IgPSBuZXcgQmxvYihbY29udGVudF0sIHsgdHlwZTogbWltZVR5cGUgfSk7CiAgY29uc3QgdXJsID0gVVJMLmNyZWF0ZU9iamVjdFVSTChibG9iKTsKICBjb25zdCBhID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYScpOwogIGEuaHJlZiA9IHVybDsKICBhLmRvd25sb2FkID0gZmlsZW5hbWU7CiAgZG9jdW1lbnQuYm9keS5hcHBlbmRDaGlsZChhKTsKICBhLmNsaWNrKCk7CiAgYS5yZW1vdmUoKTsKICBVUkwucmV2b2tlT2JqZWN0VVJMKHVybCk7Cn0KCi8qKiBRdW90ZWQtY29tbWEtam9pbiBDU1Yg4oCUIHRoZSBjbGllbnQtc2lkZSBtaXJyb3Igb2YgYXBwLmpzJ3MgdG9DU1YoKSwgZm9yIHRhYmxlcyB3aG9zZSBmdWxsIGRhdGFzZXQgaXMgYWxyZWFkeSBsb2FkZWQgaW4gdGhlIGJyb3dzZXIuICovCmZ1bmN0aW9uIHRvQ1NWQ2xpZW50U2lkZShyb3dzLCBjb2x1bW5zKSB7CiAgY29uc3QgZXNjYXBlID0gKHYpID0+IHsKICAgIGNvbnN0IHMgPSB2ID09PSBudWxsIHx8IHYgPT09IHVuZGVmaW5lZCA/ICcnIDogU3RyaW5nKHYpOwogICAgcmV0dXJuIC9bIixcclxuXS8udGVzdChzKSA/IGAiJHtzLnJlcGxhY2UoLyIvZywgJyIiJyl9ImAgOiBzOwogIH07CiAgY29uc3QgbGluZXMgPSBbY29sdW1ucy5tYXAoKGMpID0+IGVzY2FwZShjLmxhYmVsKSkuam9pbignLCcpXTsKICByb3dzLmZvckVhY2goKHJvdykgPT4gbGluZXMucHVzaChjb2x1bW5zLm1hcCgoYykgPT4gZXNjYXBlKHJvd1tjLmtleV0pKS5qb2luKCcsJykpKTsKICByZXR1cm4gbGluZXMuam9pbignXHJcbicpOwp9CgovKiogQSBzbWFsbCAiRXhwb3J0IENTViAvIEV4cG9ydCBFeGNlbCIgYnV0dG9uIHBhaXIuIGBnZXRSb3dzQW5kQ29sdW1ucygpYCBpcyBjYWxsZWQgYXQgY2xpY2sgdGltZSBzbyBpdCBhbHdheXMgZXhwb3J0cyB3aGF0ZXZlcidzIGN1cnJlbnRseSBmaWx0ZXJlZC9zb3J0ZWQsIG5ldmVyIGEgc3RhbGUgc25hcHNob3QuICovCmZ1bmN0aW9uIGJ1aWxkRXhwb3J0QnV0dG9ucyh7IGdldFJvd3NBbmRDb2x1bW5zLCBmaWxlbmFtZUJhc2UsIHNoZWV0TmFtZSB9KSB7CiAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogIHdyYXAuY2xhc3NOYW1lID0gJ2V4cG9ydC1idXR0b25zJzsKICBjb25zdCBjc3ZCdG4gPSBpY29uQnRuKCdidG4nLCAnZmlsZS1kb3duJywgJ0V4cG9ydCBDU1YnKTsKICBjc3ZCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7CiAgICBjb25zdCB7IHJvd3MsIGNvbHVtbnMgfSA9IGdldFJvd3NBbmRDb2x1bW5zKCk7CiAgICBpZiAoIXJvd3MubGVuZ3RoKSB7IFRvYXN0LnNob3coJ05vdGhpbmcgdG8gZXhwb3J0LicsICdlcnJvcicpOyByZXR1cm47IH0KICAgIGRvd25sb2FkQmxvYihgJHtmaWxlbmFtZUJhc2V9LmNzdmAsICd0ZXh0L2NzdjtjaGFyc2V0PXV0Zi04JywgdG9DU1ZDbGllbnRTaWRlKHJvd3MsIGNvbHVtbnMpKTsKICB9KTsKICBjb25zdCB4bHN4QnRuID0gaWNvbkJ0bignYnRuJywgJ2ZpbGUtc3ByZWFkc2hlZXQnLCAnRXhwb3J0IEV4Y2VsJyk7CiAgeGxzeEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIGFzeW5jICgpID0+IHsKICAgIGNvbnN0IHsgcm93cywgY29sdW1ucyB9ID0gZ2V0Um93c0FuZENvbHVtbnMoKTsKICAgIGlmICghcm93cy5sZW5ndGgpIHsgVG9hc3Quc2hvdygnTm90aGluZyB0byBleHBvcnQuJywgJ2Vycm9yJyk7IHJldHVybjsgfQogICAgdHJ5IHsKICAgICAgY29uc3QgcmVzID0gYXdhaXQgZmV0Y2goJy9hcGkvZXhwb3J0JywgewogICAgICAgIG1ldGhvZDogJ1BPU1QnLAogICAgICAgIGhlYWRlcnM6IHsgJ0NvbnRlbnQtVHlwZSc6ICdhcHBsaWNhdGlvbi9qc29uJyB9LAogICAgICAgIGJvZHk6IEpTT04uc3RyaW5naWZ5KHsgcm93cywgY29sdW1ucywgZm9ybWF0OiAneGxzeCcsIGZpbGVuYW1lOiBmaWxlbmFtZUJhc2UsIHNoZWV0TmFtZTogc2hlZXROYW1lIHx8IGZpbGVuYW1lQmFzZSB9KSwKICAgICAgfSk7CiAgICAgIGlmICghcmVzLm9rKSB0aHJvdyBuZXcgRXJyb3IoJ0V4cG9ydCBmYWlsZWQuJyk7CiAgICAgIGNvbnN0IGJsb2IgPSBhd2FpdCByZXMuYmxvYigpOwogICAgICBjb25zdCB1cmwgPSBVUkwuY3JlYXRlT2JqZWN0VVJMKGJsb2IpOwogICAgICBjb25zdCBhID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYScpOwogICAgICBhLmhyZWYgPSB1cmw7CiAgICAgIGEuZG93bmxvYWQgPSBgJHtmaWxlbmFtZUJhc2V9Lnhsc3hgOwogICAgICBkb2N1bWVudC5ib2R5LmFwcGVuZENoaWxkKGEpOwogICAgICBhLmNsaWNrKCk7CiAgICAgIGEucmVtb3ZlKCk7CiAgICAgIFVSTC5yZXZva2VPYmplY3RVUkwodXJsKTsKICAgIH0gY2F0Y2ggKGVycikgewogICAgICBUb2FzdC5zaG93KGVyci5tZXNzYWdlIHx8ICdFeHBvcnQgZmFpbGVkLicsICdlcnJvcicpOwogICAgfQogIH0pOwogIHdyYXAuYXBwZW5kKGNzdkJ0biwgeGxzeEJ0bik7CiAgcmV0dXJuIHdyYXA7Cn0KCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBDaGFydHMg4oCUIENoYXJ0LmpzIGJ1aWxkZXJzICh2YWxpZGF0ZWQgY2F0ZWdvcmljYWwgcGFsZXR0ZSwKICAgaGFpcmxpbmUgcmVjZXNzaXZlIGdyaWRsaW5lcywgc2luZ2xlIGF4aXMsIGxlZ2VuZCBhbHdheXMKICAgcHJlc2VudCBmb3IgMisgc2VyaWVzLCBpbmRleC1tb2RlIHRvb2x0aXBzKS4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmlmICh3aW5kb3cuQ2hhcnREYXRhTGFiZWxzKSBDaGFydC5yZWdpc3Rlcih3aW5kb3cuQ2hhcnREYXRhTGFiZWxzKTsKCmNvbnN0IENoYXJ0cyA9ICgoKSA9PiB7CiAgY29uc3QgcmVnaXN0cnkgPSBuZXcgTWFwKCk7IC8vIGNhbnZhc0lkIC0+IENoYXJ0IGluc3RhbmNlLCBzbyByZS1yZW5kZXJzIGRlc3Ryb3kgdGhlIG9sZCBvbmUgZmlyc3QKICBjb25zdCBNQVhfTEFCRUxFRF9JVEVNUyA9IDIwOyAvLyBiZXlvbmQgdGhpcywgcGVyLWl0ZW0gdmFsdWUgbGFiZWxzIHdvdWxkIG92ZXJsYXAg4oCUIHJlbHkgb24gdG9vbHRpcHMgaW5zdGVhZAoKICBmdW5jdGlvbiBjc3NWYXIobmFtZSkgewogICAgcmV0dXJuIGdldENvbXB1dGVkU3R5bGUoZG9jdW1lbnQuZG9jdW1lbnRFbGVtZW50KS5nZXRQcm9wZXJ0eVZhbHVlKG5hbWUpLnRyaW0oKTsKICB9CgogIGNvbnN0IFNFUklFU19WQVJTID0gWyctLXNlcmllcy0xJywgJy0tc2VyaWVzLTInLCAnLS1zZXJpZXMtMycsICctLXNlcmllcy00JywgJy0tc2VyaWVzLTUnLCAnLS1zZXJpZXMtNicsICctLXNlcmllcy03JywgJy0tc2VyaWVzLTgnXTsKICBmdW5jdGlvbiBzZXJpZXNDb2xvcihpbmRleCkgewogICAgcmV0dXJuIGNzc1ZhcihTRVJJRVNfVkFSU1tpbmRleCAlIFNFUklFU19WQVJTLmxlbmd0aF0pOwogIH0KCiAgZnVuY3Rpb24gYmFzZUdyaWQoKSB7CiAgICByZXR1cm4gewogICAgICBjb2xvcjogY3NzVmFyKCctLWdyaWRsaW5lJyksCiAgICAgIGRyYXdUaWNrczogZmFsc2UsCiAgICB9OwogIH0KICBmdW5jdGlvbiBiYXNlVGlja3MoKSB7CiAgICByZXR1cm4geyBjb2xvcjogY3NzVmFyKCctLXRleHQtbXV0ZWQnKSwgZm9udDogeyBzaXplOiAxMSB9IH07CiAgfQogIGZ1bmN0aW9uIGJhc2VUb29sdGlwKCkgewogICAgcmV0dXJuIHsKICAgICAgYmFja2dyb3VuZENvbG9yOiBjc3NWYXIoJy0tc3VyZmFjZS0xJyksCiAgICAgIHRpdGxlQ29sb3I6IGNzc1ZhcignLS10ZXh0LXByaW1hcnknKSwKICAgICAgYm9keUNvbG9yOiBjc3NWYXIoJy0tdGV4dC1zZWNvbmRhcnknKSwKICAgICAgYm9yZGVyQ29sb3I6IGNzc1ZhcignLS1ib3JkZXInKSwKICAgICAgYm9yZGVyV2lkdGg6IDEsCiAgICAgIGNvcm5lclJhZGl1czogMTAsCiAgICAgIHBhZGRpbmc6IDEyLAogICAgICBib3hQYWRkaW5nOiA0LAogICAgICB0aXRsZUZvbnQ6IHsgc2l6ZTogMTIsIHdlaWdodDogJzcwMCcgfSwKICAgICAgYm9keUZvbnQ6IHsgc2l6ZTogMTIgfSwKICAgIH07CiAgfQogIGZ1bmN0aW9uIGxhYmVsQ29sb3IoKSB7CiAgICByZXR1cm4gY3NzVmFyKCctLXRleHQtcHJpbWFyeScpOwogIH0KICAvKiogU25hcHB5LCBzdWJ0bGUgbW90aW9uIOKAlCBpbiB0aGUgMTUwLTMwMG1zIHJhbmdlIHRoZSByZWRlc2lnbiBjYWxscyBmb3IsIG5ldmVyIGJvdW5jeS4gKi8KICBmdW5jdGlvbiBiYXNlQW5pbWF0aW9uKCkgewogICAgcmV0dXJuIHsgZHVyYXRpb246IDI4MCwgZWFzaW5nOiAnZWFzZU91dFF1YXJ0JyB9OwogIH0KCiAgZnVuY3Rpb24gZGVzdHJveShjYW52YXNJZCkgewogICAgaWYgKHJlZ2lzdHJ5LmhhcyhjYW52YXNJZCkpIHsKICAgICAgcmVnaXN0cnkuZ2V0KGNhbnZhc0lkKS5kZXN0cm95KCk7CiAgICAgIHJlZ2lzdHJ5LmRlbGV0ZShjYW52YXNJZCk7CiAgICB9CiAgfQoKICAvKiogTXVsdGktc2VyaWVzIGxpbmUgY2hhcnQgKGUuZy4gd2Vla2x5IHRyZW5kIHBlciBwbGF0Zm9ybSkuIE9uZSBzZXJpZXMgbmVlZHMgbm8gbGVnZW5kIGJveC4KICAgICAgUGVyLXBvaW50IHZhbHVlIGxhYmVscyBhcmUgc2hvd24gb25seSBmb3IgYSBzaW5nbGUgc2VyaWVzIOKAlCB3aXRoIHNldmVyYWwgc2VyaWVzIG92ZXJsYWlkLAogICAgICBsYWJlbGluZyBldmVyeSBwb2ludCB3b3VsZCBvdmVybGFwLCBzbyB0aG9zZSByZWx5IG9uIHRoZSAoc3RpbGwtcHJlc2VudCkgaG92ZXIgdG9vbHRpcC4gKi8KICBmdW5jdGlvbiB0cmVuZENoYXJ0KGNhbnZhc0lkLCB7IGxhYmVscywgc2VyaWVzLCBmb3JtYXRWYWx1ZSB9KSB7CiAgICBkZXN0cm95KGNhbnZhc0lkKTsKICAgIGNvbnN0IGN0eCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGNhbnZhc0lkKTsKICAgIGlmICghY3R4KSByZXR1cm4gbnVsbDsKICAgIGNvbnN0IGZtdCA9IGZvcm1hdFZhbHVlIHx8ICgodikgPT4gRm9ybWF0LnNtYXJ0KHYpKTsKICAgIGNvbnN0IHNob3dMYWJlbHMgPSBzZXJpZXMubGVuZ3RoID09PSAxICYmIGxhYmVscy5sZW5ndGggPD0gTUFYX0xBQkVMRURfSVRFTVM7CgogICAgY29uc3QgZGF0YXNldHMgPSBzZXJpZXMubWFwKChzLCBpKSA9PiAoewogICAgICBsYWJlbDogcy5sYWJlbCwKICAgICAgZGF0YTogcy5kYXRhLAogICAgICBib3JkZXJDb2xvcjogcy5jb2xvciB8fCBzZXJpZXNDb2xvcihpKSwKICAgICAgYmFja2dyb3VuZENvbG9yOiBzLmNvbG9yIHx8IHNlcmllc0NvbG9yKGkpLAogICAgICBib3JkZXJXaWR0aDogMiwKICAgICAgcG9pbnRSYWRpdXM6IHNob3dMYWJlbHMgPyAzIDogMCwKICAgICAgcG9pbnRIb3ZlclJhZGl1czogNCwKICAgICAgcG9pbnRIaXRSYWRpdXM6IDEyLAogICAgICB0ZW5zaW9uOiAwLjI1LAogICAgICBmaWxsOiBmYWxzZSwKICAgIH0pKTsKCiAgICBjb25zdCBjaGFydCA9IG5ldyBDaGFydChjdHgsIHsKICAgICAgdHlwZTogJ2xpbmUnLAogICAgICBkYXRhOiB7IGxhYmVscywgZGF0YXNldHMgfSwKICAgICAgb3B0aW9uczogewogICAgICAgIHJlc3BvbnNpdmU6IHRydWUsCiAgICAgICAgbWFpbnRhaW5Bc3BlY3RSYXRpbzogZmFsc2UsCiAgICAgICAgaW50ZXJhY3Rpb246IHsgbW9kZTogJ2luZGV4JywgaW50ZXJzZWN0OiBmYWxzZSB9LAogICAgICAgIGxheW91dDogeyBwYWRkaW5nOiB7IHRvcDogc2hvd0xhYmVscyA/IDIwIDogOCB9IH0sCiAgICAgICAgYW5pbWF0aW9uOiBiYXNlQW5pbWF0aW9uKCksCiAgICAgICAgcGx1Z2luczogewogICAgICAgICAgbGVnZW5kOiB7CiAgICAgICAgICAgIGRpc3BsYXk6IHNlcmllcy5sZW5ndGggPiAxLAogICAgICAgICAgICBwb3NpdGlvbjogJ2JvdHRvbScsCiAgICAgICAgICAgIGxhYmVsczogeyBjb2xvcjogY3NzVmFyKCctLXRleHQtc2Vjb25kYXJ5JyksIHVzZVBvaW50U3R5bGU6IHRydWUsIHBvaW50U3R5bGU6ICdsaW5lJywgYm94V2lkdGg6IDE2LCBwYWRkaW5nOiAxNiwgZm9udDogeyBzaXplOiAxMSB9IH0sCiAgICAgICAgICB9LAogICAgICAgICAgdG9vbHRpcDogeyAuLi5iYXNlVG9vbHRpcCgpLCB1c2VQb2ludFN0eWxlOiB0cnVlIH0sCiAgICAgICAgICBkYXRhbGFiZWxzOiBzaG93TGFiZWxzCiAgICAgICAgICAgID8geyBhbGlnbjogJ3RvcCcsIGFuY2hvcjogJ2VuZCcsIGNvbG9yOiBsYWJlbENvbG9yKCksIGZvbnQ6IHsgc2l6ZTogMTEsIHdlaWdodDogJzYwMCcgfSwgZm9ybWF0dGVyOiAodikgPT4gZm10KHYpIH0KICAgICAgICAgICAgOiB7IGRpc3BsYXk6IGZhbHNlIH0sCiAgICAgICAgfSwKICAgICAgICBzY2FsZXM6IHsKICAgICAgICAgIHg6IHsgZ3JpZDogeyBkaXNwbGF5OiBmYWxzZSB9LCB0aWNrczogYmFzZVRpY2tzKCkgfSwKICAgICAgICAgIHk6IHsgZ3JpZDogYmFzZUdyaWQoKSwgdGlja3M6IGJhc2VUaWNrcygpLCBib3JkZXI6IHsgZGlzcGxheTogZmFsc2UgfSwgYmVnaW5BdFplcm86IHRydWUgfSwKICAgICAgICB9LAogICAgICB9LAogICAgfSk7CiAgICByZWdpc3RyeS5zZXQoY2FudmFzSWQsIGNoYXJ0KTsKICAgIHJldHVybiBjaGFydDsKICB9CgogIC8qKiBTaW5nbGUtbWV0cmljIGJhciBjaGFydCBhY3Jvc3MgcGxhdGZvcm1zIChpZGVudGl0eSBlbmNvZGluZyDigJQgZWFjaCBiYXIgSVMgYSBwbGF0Zm9ybSkuICovCiAgZnVuY3Rpb24gcGxhdGZvcm1CYXJDaGFydChjYW52YXNJZCwgeyBsYWJlbHMsIGRhdGEsIGNvbG9ycywgZm9ybWF0VmFsdWUgfSkgewogICAgZGVzdHJveShjYW52YXNJZCk7CiAgICBjb25zdCBjdHggPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZChjYW52YXNJZCk7CiAgICBpZiAoIWN0eCkgcmV0dXJuIG51bGw7CiAgICBjb25zdCBmbXQgPSBmb3JtYXRWYWx1ZSB8fCAoKHYpID0+IEZvcm1hdC5zbWFydCh2KSk7CiAgICBjb25zdCBzaG93TGFiZWxzID0gbGFiZWxzLmxlbmd0aCA8PSBNQVhfTEFCRUxFRF9JVEVNUzsKCiAgICBjb25zdCBjaGFydCA9IG5ldyBDaGFydChjdHgsIHsKICAgICAgdHlwZTogJ2JhcicsCiAgICAgIGRhdGE6IHsKICAgICAgICBsYWJlbHMsCiAgICAgICAgZGF0YXNldHM6IFsKICAgICAgICAgIHsKICAgICAgICAgICAgZGF0YSwKICAgICAgICAgICAgYmFja2dyb3VuZENvbG9yOiBjb2xvcnMsCiAgICAgICAgICAgIGJvcmRlclJhZGl1czogNCwKICAgICAgICAgICAgbWF4QmFyVGhpY2tuZXNzOiAyOCwKICAgICAgICAgICAgYm9yZGVyU2tpcHBlZDogJ2JvdHRvbScsCiAgICAgICAgICB9LAogICAgICAgIF0sCiAgICAgIH0sCiAgICAgIG9wdGlvbnM6IHsKICAgICAgICByZXNwb25zaXZlOiB0cnVlLAogICAgICAgIG1haW50YWluQXNwZWN0UmF0aW86IGZhbHNlLAogICAgICAgIGxheW91dDogeyBwYWRkaW5nOiB7IHRvcDogc2hvd0xhYmVscyA/IDIwIDogOCB9IH0sCiAgICAgICAgYW5pbWF0aW9uOiBiYXNlQW5pbWF0aW9uKCksCiAgICAgICAgcGx1Z2luczogewogICAgICAgICAgbGVnZW5kOiB7IGRpc3BsYXk6IGZhbHNlIH0sCiAgICAgICAgICB0b29sdGlwOiBiYXNlVG9vbHRpcCgpLAogICAgICAgICAgZGF0YWxhYmVsczogc2hvd0xhYmVscwogICAgICAgICAgICA/IHsgYWxpZ246ICdlbmQnLCBhbmNob3I6ICdlbmQnLCBjb2xvcjogbGFiZWxDb2xvcigpLCBmb250OiB7IHNpemU6IDExLCB3ZWlnaHQ6ICc2MDAnIH0sIGZvcm1hdHRlcjogKHYpID0+IGZtdCh2KSB9CiAgICAgICAgICAgIDogeyBkaXNwbGF5OiBmYWxzZSB9LAogICAgICAgIH0sCiAgICAgICAgc2NhbGVzOiB7CiAgICAgICAgICB4OiB7IGdyaWQ6IHsgZGlzcGxheTogZmFsc2UgfSwgdGlja3M6IGJhc2VUaWNrcygpIH0sCiAgICAgICAgICB5OiB7IGdyaWQ6IGJhc2VHcmlkKCksIHRpY2tzOiBiYXNlVGlja3MoKSwgYm9yZGVyOiB7IGRpc3BsYXk6IGZhbHNlIH0sIGJlZ2luQXRaZXJvOiB0cnVlIH0sCiAgICAgICAgfSwKICAgICAgfSwKICAgIH0pOwogICAgcmVnaXN0cnkuc2V0KGNhbnZhc0lkLCBjaGFydCk7CiAgICByZXR1cm4gY2hhcnQ7CiAgfQoKICAvKiogUGllIGNoYXJ0IChhIGhhbmRmdWwgb2YgY2F0ZWdvcmllcyBvbmx5IOKAlCBlLmcuIENhbXBhaWduIFBlcmZvcm1hbmNlJ3MgQWRzL09yZ2FuaWMgc3BsaXQpLgogICAgICBTbGljZSBsYWJlbHMgc2hvdyBib3RoIHNoYXJlLW9mLXdob2xlIGFuZCB0aGUgYWN0dWFsIHZhbHVlLCBwZXIgdGhlICJubyBob3ZlciByZXF1aXJlZCIgZ29hbC4gKi8KICBmdW5jdGlvbiBwaWVDaGFydChjYW52YXNJZCwgeyBsYWJlbHMsIGRhdGEsIGNvbG9ycywgZm9ybWF0VmFsdWUgfSkgewogICAgZGVzdHJveShjYW52YXNJZCk7CiAgICBjb25zdCBjdHggPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZChjYW52YXNJZCk7CiAgICBpZiAoIWN0eCkgcmV0dXJuIG51bGw7CiAgICBjb25zdCBmbXQgPSBmb3JtYXRWYWx1ZSB8fCAoKHYpID0+IEZvcm1hdC5zbWFydCh2KSk7CiAgICBjb25zdCB0b3RhbCA9IGRhdGEucmVkdWNlKChzdW0sIHYpID0+IHN1bSArICh2IHx8IDApLCAwKTsKCiAgICBjb25zdCBjaGFydCA9IG5ldyBDaGFydChjdHgsIHsKICAgICAgdHlwZTogJ3BpZScsCiAgICAgIGRhdGE6IHsKICAgICAgICBsYWJlbHMsCiAgICAgICAgZGF0YXNldHM6IFt7IGRhdGEsIGJhY2tncm91bmRDb2xvcjogY29sb3JzLCBib3JkZXJDb2xvcjogY3NzVmFyKCctLXN1cmZhY2UtMScpLCBib3JkZXJXaWR0aDogMiB9XSwKICAgICAgfSwKICAgICAgb3B0aW9uczogewogICAgICAgIHJlc3BvbnNpdmU6IHRydWUsCiAgICAgICAgbWFpbnRhaW5Bc3BlY3RSYXRpbzogZmFsc2UsCiAgICAgICAgYW5pbWF0aW9uOiBiYXNlQW5pbWF0aW9uKCksCiAgICAgICAgcGx1Z2luczogewogICAgICAgICAgbGVnZW5kOiB7IGRpc3BsYXk6IHRydWUsIHBvc2l0aW9uOiAnYm90dG9tJywgbGFiZWxzOiB7IGNvbG9yOiBjc3NWYXIoJy0tdGV4dC1zZWNvbmRhcnknKSwgYm94V2lkdGg6IDEyLCBwYWRkaW5nOiAxNiwgZm9udDogeyBzaXplOiAxMSB9IH0gfSwKICAgICAgICAgIHRvb2x0aXA6IGJhc2VUb29sdGlwKCksCiAgICAgICAgICBkYXRhbGFiZWxzOiB7CiAgICAgICAgICAgIGNvbG9yOiAnI2ZmZicsCiAgICAgICAgICAgIGZvbnQ6IHsgc2l6ZTogMTIsIHdlaWdodDogJzcwMCcgfSwKICAgICAgICAgICAgZm9ybWF0dGVyOiAodikgPT4gewogICAgICAgICAgICAgIGNvbnN0IHBjdCA9IHRvdGFsID8gTWF0aC5yb3VuZCgodiAvIHRvdGFsKSAqIDEwMDApIC8gMTAgOiAwOwogICAgICAgICAgICAgIHJldHVybiBgJHtwY3R9JVxuJHtmbXQodil9YDsKICAgICAgICAgICAgfSwKICAgICAgICAgIH0sCiAgICAgICAgfSwKICAgICAgfSwKICAgIH0pOwogICAgcmVnaXN0cnkuc2V0KGNhbnZhc0lkLCBjaGFydCk7CiAgICByZXR1cm4gY2hhcnQ7CiAgfQoKICAvKiogR3JvdXBlZCB2ZXJ0aWNhbCBiYXIgY2hhcnQg4oCUIGEgZmV3IGNhdGVnb3JpZXMsIDIrIG5hbWVkIHNlcmllcyBzaG93biBzaWRlIGJ5IHNpZGUKICAgICAgKGUuZy4gVGhpcyBXZWVrIHZzIExhc3QgV2VlayBhY3Jvc3MgbWV0cmljcykuIFNhbWUgdmlzdWFsIGxhbmd1YWdlIGFzIHBsYXRmb3JtQmFyQ2hhcnQ7CiAgICAgIGxlZ2VuZCBpcyBhbHdheXMgb24gc2luY2UgdGhlIHNlcmllcyBuYW1lcyBjYXJyeSB0aGUgbWVhbmluZy4gKi8KICBmdW5jdGlvbiBncm91cGVkQmFyQ2hhcnQoY2FudmFzSWQsIHsgbGFiZWxzLCBzZXJpZXMsIGZvcm1hdFZhbHVlIH0pIHsKICAgIGRlc3Ryb3koY2FudmFzSWQpOwogICAgY29uc3QgY3R4ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoY2FudmFzSWQpOwogICAgaWYgKCFjdHgpIHJldHVybiBudWxsOwogICAgY29uc3QgZm10ID0gZm9ybWF0VmFsdWUgfHwgKCh2KSA9PiBGb3JtYXQuc21hcnQodikpOwogICAgY29uc3Qgc2hvd0xhYmVscyA9IGxhYmVscy5sZW5ndGggKiBzZXJpZXMubGVuZ3RoIDw9IE1BWF9MQUJFTEVEX0lURU1TOwoKICAgIGNvbnN0IGNoYXJ0ID0gbmV3IENoYXJ0KGN0eCwgewogICAgICB0eXBlOiAnYmFyJywKICAgICAgZGF0YTogewogICAgICAgIGxhYmVscywKICAgICAgICBkYXRhc2V0czogc2VyaWVzLm1hcCgocywgaSkgPT4gKHsKICAgICAgICAgIGxhYmVsOiBzLmxhYmVsLAogICAgICAgICAgZGF0YTogcy5kYXRhLAogICAgICAgICAgYmFja2dyb3VuZENvbG9yOiBzLmNvbG9yIHx8IHNlcmllc0NvbG9yKGkpLAogICAgICAgICAgYm9yZGVyUmFkaXVzOiA0LAogICAgICAgICAgbWF4QmFyVGhpY2tuZXNzOiAzNCwKICAgICAgICAgIGJvcmRlclNraXBwZWQ6ICdib3R0b20nLAogICAgICAgIH0pKSwKICAgICAgfSwKICAgICAgb3B0aW9uczogewogICAgICAgIHJlc3BvbnNpdmU6IHRydWUsCiAgICAgICAgbWFpbnRhaW5Bc3BlY3RSYXRpbzogZmFsc2UsCiAgICAgICAgbGF5b3V0OiB7IHBhZGRpbmc6IHsgdG9wOiBzaG93TGFiZWxzID8gMjAgOiA4IH0gfSwKICAgICAgICBhbmltYXRpb246IGJhc2VBbmltYXRpb24oKSwKICAgICAgICBwbHVnaW5zOiB7CiAgICAgICAgICBsZWdlbmQ6IHsKICAgICAgICAgICAgZGlzcGxheTogdHJ1ZSwKICAgICAgICAgICAgcG9zaXRpb246ICdib3R0b20nLAogICAgICAgICAgICBsYWJlbHM6IHsgY29sb3I6IGNzc1ZhcignLS10ZXh0LXNlY29uZGFyeScpLCB1c2VQb2ludFN0eWxlOiB0cnVlLCBwb2ludFN0eWxlOiAncmVjdFJvdW5kZWQnLCBib3hXaWR0aDogMTIsIHBhZGRpbmc6IDE2LCBmb250OiB7IHNpemU6IDExIH0gfSwKICAgICAgICAgIH0sCiAgICAgICAgICB0b29sdGlwOiB7IC4uLmJhc2VUb29sdGlwKCksIGNhbGxiYWNrczogeyBsYWJlbDogKGMpID0+IGAgJHtjLmRhdGFzZXQubGFiZWx9OiAke2ZtdChjLnBhcnNlZC55KX1gIH0gfSwKICAgICAgICAgIGRhdGFsYWJlbHM6IHNob3dMYWJlbHMKICAgICAgICAgICAgPyB7CiAgICAgICAgICAgICAgICBhbGlnbjogJ2VuZCcsIGFuY2hvcjogJ2VuZCcsIGNvbG9yOiBsYWJlbENvbG9yKCksIGZvbnQ6IHsgc2l6ZTogMTAsIHdlaWdodDogJzYwMCcgfSwKICAgICAgICAgICAgICAgIC8vIEEgbnVsbC91bmRlZmluZWQgdmFsdWUgaXMgIk4vQSIg4oCUIGRyYXcgbm8gYmFyIGFuZCBubyBudW1iZXIsIG5ldmVyIGEgIjAiIG9yICLigJQiLgogICAgICAgICAgICAgICAgZGlzcGxheTogKGMpID0+IGMuZGF0YXNldC5kYXRhW2MuZGF0YUluZGV4XSAhPSBudWxsLAogICAgICAgICAgICAgICAgZm9ybWF0dGVyOiAodikgPT4gKHYgPT0gbnVsbCA/ICcnIDogZm10KHYpKSwKICAgICAgICAgICAgICB9CiAgICAgICAgICAgIDogeyBkaXNwbGF5OiBmYWxzZSB9LAogICAgICAgIH0sCiAgICAgICAgc2NhbGVzOiB7CiAgICAgICAgICB4OiB7IGdyaWQ6IHsgZGlzcGxheTogZmFsc2UgfSwgdGlja3M6IGJhc2VUaWNrcygpIH0sCiAgICAgICAgICB5OiB7IGdyaWQ6IGJhc2VHcmlkKCksIHRpY2tzOiBiYXNlVGlja3MoKSwgYm9yZGVyOiB7IGRpc3BsYXk6IGZhbHNlIH0sIGJlZ2luQXRaZXJvOiB0cnVlIH0sCiAgICAgICAgfSwKICAgICAgfSwKICAgIH0pOwogICAgcmVnaXN0cnkuc2V0KGNhbnZhc0lkLCBjaGFydCk7CiAgICByZXR1cm4gY2hhcnQ7CiAgfQoKICBmdW5jdGlvbiBkZXN0cm95QWxsKCkgewogICAgWy4uLnJlZ2lzdHJ5LmtleXMoKV0uZm9yRWFjaChkZXN0cm95KTsKICB9CgogIHJldHVybiB7IHRyZW5kQ2hhcnQsIHBsYXRmb3JtQmFyQ2hhcnQsIGdyb3VwZWRCYXJDaGFydCwgcGllQ2hhcnQsIHNlcmllc0NvbG9yLCBkZXN0cm95LCBkZXN0cm95QWxsIH07Cn0pKCk7CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgRGFzaGJvYXJkIHRhYjogYSBtZXRyaWMtZm9jdXNlZCBwcmVtaXVtIEJJIGRhc2hib2FyZC4gQSBzaW5nbGUKICAgTWV0cmljIHNlbGVjdG9yIChkeW5hbWljYWxseSBwb3B1bGF0ZWQgZnJvbSB3aGF0ZXZlciB0aGUKICAgc2VsZWN0ZWQgcGxhdGZvcm0ncyBkYXRhIGFjdHVhbGx5IGhhcyDigJQgbmV2ZXIgaGFyZGNvZGVkKSBkcml2ZXMKICAgdGhlIEtQSSBjYXJkcywgd2Vla2x5IHRyZW5kLCBwbGF0Zm9ybS9jYW1wYWlnbi9jb250ZW50LXR5cGUKICAgYnJlYWtkb3ducywgYW5kIHRoZSBUb3AgUGVyZm9ybWluZyBQb3N0cyByYW5raW5nIHRvZ2V0aGVyOwogICBQbGF0Zm9ybS9kYXRlL2NhbXBhaWduL2NvbnRlbnQtdHlwZSBmaWx0ZXJpbmcgY29tZXMgZnJvbSB0aGUKICAgc2hhcmVkIGZpbHRlciBiYXIuIEV2ZXJ5IGNoYXJ0IHNob3dzIGl0cyB2YWx1ZXMgZGlyZWN0bHkgKHZpYQogICBjaGFydGpzLXBsdWdpbi1kYXRhbGFiZWxzKSBzbyBub3RoaW5nIHJlcXVpcmVzIGEgaG92ZXIgdG8gcmVhZC4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmNvbnN0IERhc2hib2FyZCA9ICgoKSA9PiB7CiAgbGV0IHJvb3Q7CiAgbGV0IG1ldHJpYyA9ICd2aWV3cyc7CiAgbGV0IG1ldHJpY09wdGlvbnMgPSBbXTsKCiAgZnVuY3Rpb24gb3B0aW9uRm9yKGtleSkgewogICAgcmV0dXJuIG1ldHJpY09wdGlvbnMuZmluZCgobSkgPT4gbS5rZXkgPT09IGtleSk7CiAgfQogIGZ1bmN0aW9uIG1ldHJpY0xhYmVsKGtleSkgewogICAgY29uc3Qgb3B0ID0gb3B0aW9uRm9yKGtleSk7CiAgICByZXR1cm4gb3B0ID8gb3B0LmxhYmVsIDoga2V5OwogIH0KICBmdW5jdGlvbiBtZXRyaWNVbml0KGtleSkgewogICAgY29uc3Qgb3B0ID0gb3B0aW9uRm9yKGtleSk7CiAgICByZXR1cm4gb3B0ID8gb3B0LnVuaXQgOiAnbnVtYmVyJzsKICB9CiAgZnVuY3Rpb24gZm9ybWF0TWV0cmljVmFsdWUoa2V5LCB2YWx1ZSkgewogICAgY29uc3QgdW5pdCA9IG1ldHJpY1VuaXQoa2V5KTsKICAgIGlmICh2YWx1ZSA9PT0gbnVsbCB8fCB2YWx1ZSA9PT0gdW5kZWZpbmVkKSByZXR1cm4gJ+KAlCc7CiAgICBpZiAodW5pdCA9PT0gJ2R1cmF0aW9uJykgcmV0dXJuIEZvcm1hdC5kdXJhdGlvbih2YWx1ZSk7CiAgICByZXR1cm4gRm9ybWF0LnNtYXJ0KHZhbHVlKTsKICB9CgogIGZ1bmN0aW9uIHNoZWxsKCkgewogICAgcm9vdC5pbm5lckhUTUwgPSAnJzsKCiAgICBjb25zdCBjb250cm9scyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY29udHJvbHMuY2xhc3NOYW1lID0gJ2Rhc2hib2FyZC1jb250cm9scyc7CiAgICBjb25zdCBsYWJlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2xhYmVsJyk7CiAgICBsYWJlbC50ZXh0Q29udGVudCA9ICdNZXRyaWMnOwogICAgbGFiZWwuc2V0QXR0cmlidXRlKCdmb3InLCAnZGFzaGJvYXJkTWV0cmljU2VsZWN0Jyk7CiAgICBjb25zdCBtZXRyaWNTZWxlY3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzZWxlY3QnKTsKICAgIG1ldHJpY1NlbGVjdC5pZCA9ICdkYXNoYm9hcmRNZXRyaWNTZWxlY3QnOwogICAgbWV0cmljT3B0aW9ucy5mb3JFYWNoKChtKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOwogICAgICBvcHQudmFsdWUgPSBtLmtleTsKICAgICAgb3B0LnRleHRDb250ZW50ID0gbS5sYWJlbDsKICAgICAgaWYgKG0ua2V5ID09PSBtZXRyaWMpIG9wdC5zZWxlY3RlZCA9IHRydWU7CiAgICAgIG1ldHJpY1NlbGVjdC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgICBtZXRyaWNTZWxlY3QuYWRkRXZlbnRMaXN0ZW5lcignY2hhbmdlJywgKCkgPT4gewogICAgICBtZXRyaWMgPSBtZXRyaWNTZWxlY3QudmFsdWU7CiAgICAgIHJlZnJlc2hGb3JNZXRyaWMoKTsKICAgIH0pOwogICAgY29udHJvbHMuYXBwZW5kKGxhYmVsLCBtZXRyaWNTZWxlY3QpOwogICAgcm9vdC5hcHBlbmRDaGlsZChjb250cm9scyk7CgogICAgY29uc3Qga3BpVGl0bGUgPSB0ZXh0RWwoJ2RpdicsICdLZXkgcGVyZm9ybWFuY2UgaW5kaWNhdG9ycycsICdzZWN0aW9uLXRpdGxlJyk7CiAgICBjb25zdCBrcGlHcmlkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBrcGlHcmlkLmNsYXNzTmFtZSA9ICdzdGF0LWdyaWQnOwogICAga3BpR3JpZC5pZCA9ICdrcGlHcmlkJzsKICAgIHJvb3QuYXBwZW5kKGtwaVRpdGxlLCBrcGlHcmlkKTsKCiAgICBjb25zdCBjaGFydHNUaXRsZSA9IHRleHRFbCgnZGl2JywgJ1RyZW5kICYgcGVyZm9ybWFuY2UgYnJlYWtkb3duJywgJ3NlY3Rpb24tdGl0bGUnKTsKICAgIHJvb3QuYXBwZW5kKGNoYXJ0c1RpdGxlKTsKCiAgICBjb25zdCB0cmVuZENhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRyZW5kQ2FyZC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICBjb25zdCB0cmVuZEhlYWRlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdHJlbmRIZWFkZXIuY2xhc3NOYW1lID0gJ2NhcmQtaGVhZGVyJzsKICAgIHRyZW5kSGVhZGVyLmFwcGVuZENoaWxkKHRleHRFbCgnaDMnLCAnV2Vla2x5IHBlcmZvcm1hbmNlJykpOwogICAgdHJlbmRIZWFkZXIuZmlyc3RDaGlsZC5pZCA9ICd0cmVuZENhcmRUaXRsZSc7CiAgICB0cmVuZENhcmQuYXBwZW5kQ2hpbGQodHJlbmRIZWFkZXIpOwogICAgY29uc3QgdHJlbmRDaGFydFdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRyZW5kQ2hhcnRXcmFwLmNsYXNzTmFtZSA9ICdjaGFydC13cmFwIHRhbGwnOwogICAgdHJlbmRDaGFydFdyYXAuaWQgPSAndHJlbmRDaGFydFdyYXAnOwogICAgdHJlbmRDaGFydFdyYXAuaW5uZXJIVE1MID0gJzxjYW52YXMgaWQ9InRyZW5kQ2FudmFzIj48L2NhbnZhcz4nOwogICAgdHJlbmRDYXJkLmFwcGVuZENoaWxkKHRyZW5kQ2hhcnRXcmFwKTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQodHJlbmRDYXJkKTsKCiAgICBjb25zdCBncmlkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBncmlkLmNsYXNzTmFtZSA9ICdjYXJkLWdyaWQgZXZlbic7CiAgICBncmlkLnN0eWxlLm1hcmdpblRvcCA9ICcxNnB4JzsKCiAgICBjb25zdCBicmVha2Rvd25DYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBicmVha2Rvd25DYXJkLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGNvbnN0IGJyZWFrZG93bkhlYWRlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYnJlYWtkb3duSGVhZGVyLmNsYXNzTmFtZSA9ICdjYXJkLWhlYWRlcic7CiAgICBicmVha2Rvd25IZWFkZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMycsICcnKSk7CiAgICBicmVha2Rvd25IZWFkZXIuZmlyc3RDaGlsZC5pZCA9ICdicmVha2Rvd25DYXJkVGl0bGUnOwogICAgYnJlYWtkb3duQ2FyZC5hcHBlbmRDaGlsZChicmVha2Rvd25IZWFkZXIpOwogICAgY29uc3QgYnJlYWtkb3duV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYnJlYWtkb3duV3JhcC5jbGFzc05hbWUgPSAnY2hhcnQtd3JhcCc7CiAgICBicmVha2Rvd25XcmFwLmlkID0gJ2JyZWFrZG93bkNoYXJ0V3JhcCc7CiAgICBicmVha2Rvd25XcmFwLmlubmVySFRNTCA9ICc8Y2FudmFzIGlkPSJicmVha2Rvd25DYW52YXMiPjwvY2FudmFzPic7CiAgICBicmVha2Rvd25DYXJkLmFwcGVuZENoaWxkKGJyZWFrZG93bldyYXApOwoKICAgIGNvbnN0IGNvbnRlbnRUeXBlQ2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY29udGVudFR5cGVDYXJkLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGNvbnN0IGNvbnRlbnRUeXBlSGVhZGVyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjb250ZW50VHlwZUhlYWRlci5jbGFzc05hbWUgPSAnY2FyZC1oZWFkZXInOwogICAgY29udGVudFR5cGVIZWFkZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMycsICcnKSk7CiAgICBjb250ZW50VHlwZUhlYWRlci5maXJzdENoaWxkLmlkID0gJ2NvbnRlbnRUeXBlQ2FyZFRpdGxlJzsKICAgIGNvbnRlbnRUeXBlQ2FyZC5hcHBlbmRDaGlsZChjb250ZW50VHlwZUhlYWRlcik7CiAgICBjb25zdCBjb250ZW50VHlwZVdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNvbnRlbnRUeXBlV3JhcC5jbGFzc05hbWUgPSAnY2hhcnQtd3JhcCc7CiAgICBjb250ZW50VHlwZVdyYXAuaWQgPSAnY29udGVudFR5cGVDaGFydFdyYXAnOwogICAgY29udGVudFR5cGVXcmFwLmlubmVySFRNTCA9ICc8Y2FudmFzIGlkPSJjb250ZW50VHlwZUNhbnZhcyI+PC9jYW52YXM+JzsKICAgIGNvbnRlbnRUeXBlQ2FyZC5hcHBlbmRDaGlsZChjb250ZW50VHlwZVdyYXApOwoKICAgIGdyaWQuYXBwZW5kKGJyZWFrZG93bkNhcmQsIGNvbnRlbnRUeXBlQ2FyZCk7CiAgICByb290LmFwcGVuZENoaWxkKGdyaWQpOwoKICAgIGNvbnN0IHRvcFRpdGxlID0gdGV4dEVsKCdkaXYnLCAnVG9wLXBlcmZvcm1pbmcgcG9zdHMnLCAnc2VjdGlvbi10aXRsZScpOwogICAgY29uc3QgdG9wQ2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdG9wQ2FyZC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICBjb25zdCB0b3BIZWFkZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRvcEhlYWRlci5jbGFzc05hbWUgPSAnY2FyZC1oZWFkZXInOwogICAgdG9wSGVhZGVyLmFwcGVuZENoaWxkKHRleHRFbCgnaDMnLCAnUmFua2VkIGJ5IHNlbGVjdGVkIG1ldHJpYycpKTsKICAgIHRvcENhcmQuYXBwZW5kQ2hpbGQodG9wSGVhZGVyKTsKICAgIGNvbnN0IHRhYmxlV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGFibGVXcmFwLmNsYXNzTmFtZSA9ICd0YWJsZS1zY3JvbGwnOwogICAgdGFibGVXcmFwLmlkID0gJ3RvcFBvc3RzVGFibGUnOwogICAgdG9wQ2FyZC5hcHBlbmRDaGlsZCh0YWJsZVdyYXApOwogICAgcm9vdC5hcHBlbmQodG9wVGl0bGUsIHRvcENhcmQpOwogIH0KCiAgLyoqIEEgc21hbGwgbGFiZWwrdmFsdWUgcGFpciB1c2VkIGluc2lkZSB0aGUgQmVzdCBQZXJmb3JtaW5nIFBvc3QgY2FyZCdzIG1ldHJpY3MgY29sdW1uLiBgdmFyaWFudGAgKCdwcmltYXJ5Jy8nc2Vjb25kYXJ5JykgY29udHJvbHMgc2l6ZSBhbmQgd2hldGhlciBhIGRpdmlkZXIgcnVsZSBzaXRzIGFib3ZlIGl0LiAqLwogIGZ1bmN0aW9uIG1ldHJpY0Jsb2NrKGxhYmVsLCB2YWx1ZSwgdmFyaWFudCkgewogICAgY29uc3QgYmxvY2sgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGJsb2NrLmNsYXNzTmFtZSA9IHZhcmlhbnQgPT09ICdzZWNvbmRhcnknID8gJ3Bvc3QtdGlsZS1tZXRyaWMtYmxvY2sgc2Vjb25kYXJ5JyA6ICdwb3N0LXRpbGUtbWV0cmljLWJsb2NrJzsKICAgIGNvbnN0IHZhbHVlRWwgPSB0ZXh0RWwoJ2RpdicsIHZhbHVlLCAncG9zdC10aWxlLW1ldHJpYy12YWx1ZScpOwogICAgYmxvY2suYXBwZW5kKHRleHRFbCgnZGl2JywgbGFiZWwsICdwb3N0LXRpbGUtbWV0cmljLWxhYmVsJyksIHZhbHVlRWwpOwogICAgcmVxdWVzdEFuaW1hdGlvbkZyYW1lKCgpID0+IGZpdFN0YXRWYWx1ZSh2YWx1ZUVsLCAxMikpOwogICAgcmV0dXJuIGJsb2NrOwogIH0KCiAgLyoqIEEgZmVhdHVyZWQgbGFuZHNjYXBlIGNhcmQgKDMgS1BJLXRpbGUtd2lkdGhzLCBzYW1lIGZpeGVkIGhlaWdodCBhcyB0aGUgcmVzdCBvZiB0aGUKICAgICAgcm93KTogdGhlIHRvcC10aWVkIHBvc3QncyBjYXB0aW9uICh3cmFwcyB1cCB0byAzIGxpbmVzKSB3aXRoIGEgcGxhdGZvcm0tY29sb3IgZG90CiAgICAgICsgcGxhdGZvcm0gbmFtZSArIGRhdGUgb24gdGhlIGxlZnQgKHdoZW4gdGhlcmUncyBtb3JlIHRoYW4gb25lIHRpZSwgdGhlIGV4dHJhIGNvdW50CiAgICAgIGlzIGZvbGRlZCBpbnRvIHRoYXQgc2FtZSBtZXRhIGxpbmUgcmF0aGVyIHRoYW4gbGlzdGluZyBldmVyeSB0aWVkIHBvc3QsIHNvIHRoZSBjYXJkCiAgICAgIG5ldmVyIGhhcyB0byBncm93IHRhbGxlciB0aGFuIGl0cyBuZWlnaGJvcnMpOyB0aGUgc2VsZWN0ZWQgbWV0cmljIChsYXJnZSkgYW5kCiAgICAgIEN1cnJlbnQgRm9sbG93ZXJzIChzbWFsbGVyLCBiZWxvdyBhIGRpdmlkZXIpIHN0YWNrZWQgaW4gYSBuYXJyb3cgY29sdW1uIG9uIHRoZSByaWdodC4gKi8KICBmdW5jdGlvbiBiZXN0UG9zdHNUaWxlKGxhYmVsLCBwb3N0cywgY3VycmVudEZvbGxvd2VycykgewogICAgY29uc3QgdGlsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGlsZS5jbGFzc05hbWUgPSAnc3RhdC10aWxlIHBvc3QtdGlsZSc7CgogICAgY29uc3QgbWFpbiA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgbWFpbi5jbGFzc05hbWUgPSAncG9zdC10aWxlLW1haW4nOwogICAgbWFpbi5hcHBlbmRDaGlsZChzdGF0TGFiZWxFbChsYWJlbCwgJ3Ryb3BoeScsICdnb2xkJykpOwoKICAgIGNvbnN0IGhhc1Bvc3RzID0gQm9vbGVhbihwb3N0cyAmJiBwb3N0cy5sZW5ndGgpOwogICAgaWYgKGhhc1Bvc3RzKSB7CiAgICAgIGNvbnN0IHBsYXRmb3JtT3B0aW9ucyA9ICh3aW5kb3cuX19maWx0ZXJPcHRpb25zQ2FjaGUgfHwgeyBwbGF0Zm9ybXM6IFtdIH0pLnBsYXRmb3JtczsKICAgICAgY29uc3QgcHJpbWFyeSA9IHBvc3RzWzBdOwogICAgICBjb25zdCBwbGF0TWV0YSA9IHBsYXRmb3JtT3B0aW9ucy5maW5kKChwKSA9PiBwLmlkID09PSBwcmltYXJ5LnBsYXRmb3JtKSB8fCB7IGxhYmVsOiBwcmltYXJ5LnBsYXRmb3JtLCBjb2xvcjogJ3ZhcigtLXNlcmllcy0xKScgfTsKICAgICAgY29uc3QgY2FwdGlvbiA9IHByaW1hcnkuY2FwdGlvbiB8fCAnKG5vIGNhcHRpb24pJzsKICAgICAgY29uc3QgY2FwdGlvbkVsID0gdGV4dEVsKCdkaXYnLCBjYXB0aW9uLCAncG9zdC10aWxlLWNhcHRpb24nKTsKICAgICAgY2FwdGlvbkVsLnRpdGxlID0gY2FwdGlvbjsKICAgICAgbWFpbi5hcHBlbmRDaGlsZChjYXB0aW9uRWwpOwogICAgICBjb25zdCB0aWVkTm90ZSA9IHBvc3RzLmxlbmd0aCA+IDEgPyBgIMK3ICske3Bvc3RzLmxlbmd0aCAtIDF9IG1vcmUgdGllZGAgOiAnJzsKICAgICAgY29uc3QgbWV0YUxpbmUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgbWV0YUxpbmUuY2xhc3NOYW1lID0gJ3Bvc3QtdGlsZS1tZXRhJzsKICAgICAgY29uc3QgZG90ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICBkb3QuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLWRvdCc7CiAgICAgIGRvdC5zdHlsZS5iYWNrZ3JvdW5kID0gcGxhdE1ldGEuY29sb3I7CiAgICAgIG1ldGFMaW5lLmFwcGVuZChkb3QsIGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKGAke3BsYXRNZXRhLmxhYmVsfSDCtyAke0Zvcm1hdC5kYXRlKHByaW1hcnkucHVibGlzaF9kYXRlKX0ke3RpZWROb3RlfWApKTsKICAgICAgbWFpbi5hcHBlbmRDaGlsZChtZXRhTGluZSk7CiAgICB9IGVsc2UgewogICAgICBtYWluLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgJ05vIGRhdGEgeWV0JywgJ3Bvc3QtdGlsZS1jYXB0aW9uIG11dGVkJykpOwogICAgfQogICAgdGlsZS5hcHBlbmRDaGlsZChtYWluKTsKCiAgICBjb25zdCBoYXNGb2xsb3dlcnMgPSBjdXJyZW50Rm9sbG93ZXJzICE9PSBudWxsICYmIGN1cnJlbnRGb2xsb3dlcnMgIT09IHVuZGVmaW5lZDsKICAgIGlmIChoYXNQb3N0cyB8fCBoYXNGb2xsb3dlcnMpIHsKICAgICAgdGlsZS5hcHBlbmRDaGlsZChPYmplY3QuYXNzaWduKGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpLCB7IGNsYXNzTmFtZTogJ3Bvc3QtdGlsZS1kaXZpZGVyJyB9KSk7CiAgICAgIGNvbnN0IG1ldHJpY3MgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgbWV0cmljcy5jbGFzc05hbWUgPSAncG9zdC10aWxlLW1ldHJpY3MnOwogICAgICBpZiAoaGFzUG9zdHMpIG1ldHJpY3MuYXBwZW5kQ2hpbGQobWV0cmljQmxvY2sobWV0cmljTGFiZWwobWV0cmljKSwgZm9ybWF0TWV0cmljVmFsdWUobWV0cmljLCBwb3N0c1swXS52YWx1ZSksICdwcmltYXJ5JykpOwogICAgICBpZiAoaGFzRm9sbG93ZXJzKSBtZXRyaWNzLmFwcGVuZENoaWxkKG1ldHJpY0Jsb2NrKCdDdXJyZW50IEZvbGxvd2VycycsIEZvcm1hdC5udW1iZXIoY3VycmVudEZvbGxvd2VycyksICdzZWNvbmRhcnknKSk7CiAgICAgIHRpbGUuYXBwZW5kQ2hpbGQobWV0cmljcyk7CiAgICB9CiAgICByZXR1cm4gdGlsZTsKICB9CgogIC8qKiBJY29uLWJhZGdlICsgdGV4dCBsYWJlbCByb3csIHNoYXJlZCBieSBldmVyeSBLUEkgdGlsZSBiZWxvdyAobWF0Y2hlcyB0aGUgcmVmZXJlbmNlIGRhc2hib2FyZCdzIGNvbG9yZWQgcGVyLWNhcmQgaWNvbnMpLiBgdmFyaWFudGAgcGlja3MgdGhlIGJhZGdlIGNvbG9yOiB2MS12NiBtYXAgdG8gdGhlIHBsYXRmb3JtIHNlcmllcyBwYWxldHRlLCAnZ29sZCcgaXMgcmVzZXJ2ZWQgZm9yIHRoZSBicmFuZC1hY2NlbnQgdGlsZS4gKi8KICBmdW5jdGlvbiBzdGF0TGFiZWxFbCh0ZXh0LCBpY29uLCB2YXJpYW50KSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB3cmFwLmNsYXNzTmFtZSA9ICdzdGF0LWxhYmVsJzsKICAgIGNvbnN0IGJhZGdlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgYmFkZ2UuY2xhc3NOYW1lID0gYHN0YXQtaWNvbiAke3ZhcmlhbnR9YDsKICAgIGJhZGdlLmlubmVySFRNTCA9IGA8aSBkYXRhLWx1Y2lkZT0iJHtpY29ufSIgc3R5bGU9IndpZHRoOjE2cHg7aGVpZ2h0OjE2cHg7Ij48L2k+YDsKICAgIHdyYXAuYXBwZW5kKGJhZGdlLCBkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZSh0ZXh0KSk7CiAgICByZXR1cm4gd3JhcDsKICB9CgogIGZ1bmN0aW9uIHN0YXRUaWxlKGxhYmVsLCB2YWx1ZSwgZm9ybWF0Rm4sIGljb24sIHZhcmlhbnQpIHsKICAgIGNvbnN0IHRpbGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRpbGUuY2xhc3NOYW1lID0gJ3N0YXQtdGlsZSc7CiAgICBjb25zdCB2YWx1ZUVsID0gdGV4dEVsKCdkaXYnLCAnJywgJ3N0YXQtdmFsdWUnKTsKICAgIHRpbGUuYXBwZW5kKHN0YXRMYWJlbEVsKGxhYmVsLCBpY29uLCB2YXJpYW50KSwgdmFsdWVFbCk7CiAgICBjb25zdCBmbXQgPSBmb3JtYXRGbiB8fCAoKHYpID0+IEZvcm1hdC5udW1iZXIodikpOwogICAgaWYgKHR5cGVvZiB2YWx1ZSA9PT0gJ251bWJlcicgJiYgTnVtYmVyLmlzRmluaXRlKHZhbHVlKSkgewogICAgICBhbmltYXRlQ291bnQodmFsdWVFbCwgMCwgdmFsdWUsIDkwMCwgZm10KTsKICAgIH0gZWxzZSB7CiAgICAgIHZhbHVlRWwudGV4dENvbnRlbnQgPSBmbXQodmFsdWUpOwogICAgICByZXF1ZXN0QW5pbWF0aW9uRnJhbWUoKCkgPT4gZml0U3RhdFZhbHVlKHZhbHVlRWwpKTsKICAgIH0KICAgIHJldHVybiB0aWxlOwogIH0KCiAgLyoqICJGb2xsb3dlcnMgR3Jvd3RoIiB0aWxlOiBhbiBhYnNvbHV0ZS1kaWZmZXJlbmNlIHN0YXQtdmFsdWUgcGx1cyBhIHBlcmNlbnRhZ2UgZGVsdGEgbGluZSAoYXJyb3cgKyBjb2xvciBkcml2ZW4gYnkgRm9ybWF0LmRlbHRhQ2xhc3MsIHNhbWUgY29udmVudGlvbiBhcyB0aGUgQ29tcGFyaXNvbnMgcGFnZSdzIHN0YXQgdGlsZXMpLiAqLwogIGZ1bmN0aW9uIGZvbGxvd2Vyc0dyb3d0aFRpbGUoY2hhbmdlLCBjaGFuZ2VQY3QpIHsKICAgIGNvbnN0IHRpbGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRpbGUuY2xhc3NOYW1lID0gJ3N0YXQtdGlsZSc7CiAgICBjb25zdCB2YWx1ZUVsID0gdGV4dEVsKCdkaXYnLCAnJywgJ3N0YXQtdmFsdWUnKTsKICAgIHRpbGUuYXBwZW5kKHN0YXRMYWJlbEVsKCdGb2xsb3dlcnMgR3Jvd3RoJywgJ3RyZW5kaW5nLXVwJywgJ3YzJyksIHZhbHVlRWwpOwogICAgaWYgKGNoYW5nZSA9PT0gbnVsbCB8fCBjaGFuZ2UgPT09IHVuZGVmaW5lZCkgewogICAgICB2YWx1ZUVsLnRleHRDb250ZW50ID0gJ+KAlCc7CiAgICAgIHJlcXVlc3RBbmltYXRpb25GcmFtZSgoKSA9PiBmaXRTdGF0VmFsdWUodmFsdWVFbCkpOwogICAgfSBlbHNlIHsKICAgICAgYW5pbWF0ZUNvdW50KHZhbHVlRWwsIDAsIGNoYW5nZSwgOTAwLCAodikgPT4gYCR7diA+IDAgPyAnKycgOiAnJ30ke0Zvcm1hdC5udW1iZXIoTWF0aC5yb3VuZCh2KSl9YCk7CiAgICB9CiAgICBjb25zdCBkZWx0YVRleHQgPSBjaGFuZ2VQY3QgPT09IG51bGwgfHwgY2hhbmdlUGN0ID09PSB1bmRlZmluZWQgPyAn4oCUJyA6IEZvcm1hdC5wY3QoY2hhbmdlUGN0KTsKICAgIHRpbGUuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBkZWx0YVRleHQsIGBzdGF0LWRlbHRhICR7Rm9ybWF0LmRlbHRhQ2xhc3MoY2hhbmdlUGN0KX1gKSk7CiAgICByZXR1cm4gdGlsZTsKICB9CgogIC8qKiAiTmV3IEZvbGxvd2VycyIgdGlsZTogZm9sbG93ZXJzIGdhaW5lZCB3aXRoaW4gdGhlIGN1cnJlbnRseSBzZWxlY3RlZCBkYXRlIHJhbmdlIOKAlCBzaG93cyAiTm8gZm9sbG93ZXIgdXBkYXRlIiByYXRoZXIgdGhhbiAwIHdoZW4gbm90aGluZyBpcyBjb21wdXRhYmxlIGZvciB0aGUgcmFuZ2UgKHBlciBzcGVjKSwgd2hpY2ggaXMgZGlmZmVyZW50IGZyb20gYSBnZW51aW5lIHplcm8uICovCiAgZnVuY3Rpb24gbmV3Rm9sbG93ZXJzVGlsZShuZXdGb2xsb3dlcnMpIHsKICAgIGNvbnN0IHRpbGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRpbGUuY2xhc3NOYW1lID0gJ3N0YXQtdGlsZSc7CiAgICBjb25zdCB2YWx1ZUVsID0gdGV4dEVsKCdkaXYnLCAnJywgJ3N0YXQtdmFsdWUnKTsKICAgIHRpbGUuYXBwZW5kKHN0YXRMYWJlbEVsKCdOZXcgRm9sbG93ZXJzJywgJ3VzZXItcGx1cycsICd2MScpLCB2YWx1ZUVsKTsKICAgIGlmIChuZXdGb2xsb3dlcnMgPT09IG51bGwgfHwgbmV3Rm9sbG93ZXJzID09PSB1bmRlZmluZWQpIHsKICAgICAgdmFsdWVFbC50ZXh0Q29udGVudCA9ICdObyBmb2xsb3dlciB1cGRhdGUnOwogICAgICB2YWx1ZUVsLmNsYXNzTGlzdC5hZGQoJ3N0YXQtdmFsdWUtbXV0ZWQnKTsKICAgICAgcmVxdWVzdEFuaW1hdGlvbkZyYW1lKCgpID0+IGZpdFN0YXRWYWx1ZSh2YWx1ZUVsKSk7CiAgICB9IGVsc2UgewogICAgICBhbmltYXRlQ291bnQodmFsdWVFbCwgMCwgbmV3Rm9sbG93ZXJzLCA5MDAsICh2KSA9PiBgJHt2ID4gMCA/ICcrJyA6ICcnfSR7Rm9ybWF0Lm51bWJlcihNYXRoLnJvdW5kKHYpKX1gKTsKICAgIH0KICAgIHJldHVybiB0aWxlOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyS3BpcyhzdW1tYXJ5LCBmb2xsb3dlcnMpIHsKICAgIGNvbnN0IGdyaWQgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgna3BpR3JpZCcpOwogICAgaWYgKCFncmlkKSByZXR1cm47CiAgICBncmlkLmlubmVySFRNTCA9ICcnOwoKICAgIC8vIGFwcGxpY2FibGUgPT09IGZhbHNlOiBubyBwb3N0IGluIHRoaXMgc2xpY2UgY2FycmllcyB0aGUgbWV0cmljIGF0IGFsbCDigJQgSGlnaGVzdC8KICAgIC8vIEF2ZXJhZ2UvVG90YWwgYXJlIE4vQSwgbm90IGEgZ2VudWluZSAwLiAoT2xkZXIgQVBJIHdpdGhvdXQgdGhlIGZsYWcg4oaSIHRyZWF0IGFzIGFwcGxpY2FibGUuKQogICAgY29uc3QgbWV0cmljQXBwbGljYWJsZSA9IHN1bW1hcnkuYXBwbGljYWJsZSAhPT0gZmFsc2U7CiAgICBjb25zdCBtZXRyaWNGbXQgPSBtZXRyaWNBcHBsaWNhYmxlID8gKHYpID0+IGZvcm1hdE1ldHJpY1ZhbHVlKG1ldHJpYywgdikgOiAoKSA9PiAnTi9BJzsKICAgIGdyaWQuYXBwZW5kQ2hpbGQoc3RhdFRpbGUoJ0hpZ2hlc3QgVmFsdWUnLCBtZXRyaWNBcHBsaWNhYmxlID8gc3VtbWFyeS5oaWdoZXN0IDogbnVsbCwgbWV0cmljRm10LCAndHJlbmRpbmctdXAnLCAndjEnKSk7CiAgICBncmlkLmFwcGVuZENoaWxkKHN0YXRUaWxlKCdBdmVyYWdlIFZhbHVlJywgbWV0cmljQXBwbGljYWJsZSA/IHN1bW1hcnkuYXZlcmFnZSA6IG51bGwsIG1ldHJpY0ZtdCwgJ2Jhci1jaGFydC0yJywgJ3Y0JykpOwogICAgZ3JpZC5hcHBlbmRDaGlsZChzdGF0VGlsZSgnVG90YWwgVmFsdWUnLCBtZXRyaWNBcHBsaWNhYmxlID8gc3VtbWFyeS50b3RhbCA6IG51bGwsIG1ldHJpY0ZtdCwgJ2xheWVycycsICd2NScpKTsKICAgIGdyaWQuYXBwZW5kQ2hpbGQoc3RhdFRpbGUoJ051bWJlciBvZiBQb3N0cycsIHN1bW1hcnkucG9zdENvdW50LCAodikgPT4gRm9ybWF0Lm51bWJlcih2KSwgJ2ZpbGUtdGV4dCcsICd2NicpKTsKICAgIGdyaWQuYXBwZW5kQ2hpbGQoc3RhdFRpbGUoJ0N1cnJlbnQgRm9sbG93ZXJzJywgZm9sbG93ZXJzLmN1cnJlbnRGb2xsb3dlcnMsICh2KSA9PiAodiA9PT0gbnVsbCB8fCB2ID09PSB1bmRlZmluZWQgPyAn4oCUJyA6IEZvcm1hdC5udW1iZXIodikpLCAndXNlcnMnLCAndjInKSk7CiAgICBncmlkLmFwcGVuZENoaWxkKGZvbGxvd2Vyc0dyb3d0aFRpbGUoZm9sbG93ZXJzLmZvbGxvd2Vyc0NoYW5nZSwgZm9sbG93ZXJzLmZvbGxvd2Vyc0NoYW5nZVBjdCkpOwogICAgZ3JpZC5hcHBlbmRDaGlsZChuZXdGb2xsb3dlcnNUaWxlKGZvbGxvd2Vycy5uZXdGb2xsb3dlcnMpKTsKICAgIGdyaWQuYXBwZW5kQ2hpbGQoYmVzdFBvc3RzVGlsZSgnQmVzdCBQZXJmb3JtaW5nIFBvc3QnLCBzdW1tYXJ5LmJlc3RQb3N0cywgZm9sbG93ZXJzLmN1cnJlbnRGb2xsb3dlcnMpKTsKICB9CgoKICAvKiogU3dhcHMgYSBjaGFydCBjYXJkJ3MgY2FudmFzIGZvciBhbiBlbXB0eS1zdGF0ZSBtZXNzYWdlLCBvciByZXN0b3JlcyB0aGUgY2FudmFzIOKAlCBzaW5jZQogICAgICByZS1yZW5kZXJpbmcgYSBDaGFydC5qcyBpbnN0YW5jZSBuZWVkcyBhIGxpdmUgPGNhbnZhcz4sIG5vdCB3aGF0ZXZlciB0aGUgbGFzdCByZW5kZXIgbGVmdCB0aGVyZS4gKi8KICBmdW5jdGlvbiBjaGFydE9yRW1wdHkod3JhcElkLCBjYW52YXNJZCwgaGFzRGF0YSwgZW1wdHlNZXNzYWdlLCByZW5kZXJGbikgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKHdyYXBJZCk7CiAgICBpZiAoIXdyYXApIHJldHVybjsKICAgIENoYXJ0cy5kZXN0cm95KGNhbnZhc0lkKTsKICAgIGlmICghaGFzRGF0YSkgewogICAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgICB3cmFwLmFwcGVuZENoaWxkKGVtcHR5U3RhdGUoeyBpY29uOiAnYmFyLWNoYXJ0LTMnLCBtZXNzYWdlOiBlbXB0eU1lc3NhZ2UgfSkpOwogICAgICByZXR1cm47CiAgICB9CiAgICB3cmFwLmlubmVySFRNTCA9IGA8Y2FudmFzIGlkPSIke2NhbnZhc0lkfSI+PC9jYW52YXM+YDsKICAgIHJlbmRlckZuKCk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiByZW5kZXJUcmVuZChmaWx0ZXJzKSB7CiAgICBjb25zdCBwbGF0Zm9ybU9wdGlvbnMgPSAod2luZG93Ll9fZmlsdGVyT3B0aW9uc0NhY2hlIHx8IHsgcGxhdGZvcm1zOiBbXSB9KS5wbGF0Zm9ybXM7CiAgICBjb25zdCBtTGFiZWwgPSBtZXRyaWNMYWJlbChtZXRyaWMpOwogICAgY29uc3QgcGxhdGZvcm1zVG9GZXRjaCA9IGZpbHRlcnMucGxhdGZvcm0gPT09ICdhbGwnID8gcGxhdGZvcm1PcHRpb25zLm1hcCgocCkgPT4gcC5pZCkgOiBbZmlsdGVycy5wbGF0Zm9ybV07CiAgICBjb25zdCB0cmVuZFJlc3BvbnNlcyA9IGF3YWl0IFByb21pc2UuYWxsKAogICAgICBwbGF0Zm9ybXNUb0ZldGNoLm1hcCgocCkgPT4KICAgICAgICBBcGkudHJlbmQoeyBkYXRlRnJvbTogZmlsdGVycy5kYXRlRnJvbSwgZGF0ZVRvOiBmaWx0ZXJzLmRhdGVUbywgcGxhdGZvcm06IHAsIGNhbXBhaWduVHlwZTogZmlsdGVycy5jYW1wYWlnblR5cGUsIGNvbnRlbnRUeXBlOiBmaWx0ZXJzLmNvbnRlbnRUeXBlIH0pCiAgICAgICkKICAgICk7CiAgICBjb25zdCB3ZWVrU2V0ID0gbmV3IFNldCgpOwogICAgdHJlbmRSZXNwb25zZXMuZm9yRWFjaCgocm93cykgPT4gcm93cy5mb3JFYWNoKChyKSA9PiB3ZWVrU2V0LmFkZChyLnBlcmlvZCkpKTsKICAgIGNvbnN0IHdlZWtzID0gWy4uLndlZWtTZXRdLnNvcnQoKTsKICAgIGNvbnN0IHNlcmllcyA9IHBsYXRmb3Jtc1RvRmV0Y2gubWFwKChwLCBpKSA9PiB7CiAgICAgIGNvbnN0IG1ldGEgPSBwbGF0Zm9ybU9wdGlvbnMuZmluZCgocGwpID0+IHBsLmlkID09PSBwKSB8fCB7IGxhYmVsOiBwIH07CiAgICAgIGNvbnN0IGJ5V2VlayA9IE9iamVjdC5mcm9tRW50cmllcyh0cmVuZFJlc3BvbnNlc1tpXS5tYXAoKHIpID0+IFtyLnBlcmlvZCwgclttZXRyaWNdXSkpOwogICAgICByZXR1cm4geyBsYWJlbDogbWV0YS5sYWJlbCwgY29sb3I6IG1ldGEuY29sb3IsIGRhdGE6IHdlZWtzLm1hcCgodykgPT4gKGJ5V2Vla1t3XSA9PT0gdW5kZWZpbmVkID8gbnVsbCA6IGJ5V2Vla1t3XSkpIH07CiAgICB9KTsKCiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndHJlbmRDYXJkVGl0bGUnKS50ZXh0Q29udGVudCA9CiAgICAgIGZpbHRlcnMucGxhdGZvcm0gPT09ICdhbGwnID8gYFdlZWtseSAke21MYWJlbH0gYnkgUGxhdGZvcm1gIDogYCR7bUxhYmVsfSBUcmVuZGA7CgogICAgY2hhcnRPckVtcHR5KCd0cmVuZENoYXJ0V3JhcCcsICd0cmVuZENhbnZhcycsIHdlZWtzLmxlbmd0aCA+IDAsICdObyBkYXRhIGluIHRoaXMgcmFuZ2UgeWV0LicsICgpID0+IHsKICAgICAgQ2hhcnRzLnRyZW5kQ2hhcnQoJ3RyZW5kQ2FudmFzJywgeyBsYWJlbHM6IHdlZWtzLm1hcChGb3JtYXQuZGF0ZSksIHNlcmllcywgZm9ybWF0VmFsdWU6ICh2KSA9PiBmb3JtYXRNZXRyaWNWYWx1ZShtZXRyaWMsIHYpIH0pOwogICAgfSk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiByZW5kZXJCcmVha2Rvd24oZmlsdGVycykgewogICAgY29uc3QgbUxhYmVsID0gbWV0cmljTGFiZWwobWV0cmljKTsKICAgIGNvbnN0IHRpdGxlRWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYnJlYWtkb3duQ2FyZFRpdGxlJyk7CgogICAgaWYgKGZpbHRlcnMucGxhdGZvcm0gPT09ICdhbGwnKSB7CiAgICAgIHRpdGxlRWwudGV4dENvbnRlbnQgPSBgUGxhdGZvcm0gQ29tcGFyaXNvbiDigJQgJHttTGFiZWx9YDsKICAgICAgY29uc3QgYnJlYWtkb3duID0gYXdhaXQgQXBpLnBsYXRmb3JtQnJlYWtkb3duKGZpbHRlcnMpOwogICAgICAvLyBPbmx5IHBsb3QgcGxhdGZvcm1zIHRoYXQgYWN0dWFsbHkgY2FycnkgdGhpcyBtZXRyaWMg4oCUIGEgcGxhdGZvcm0gd2hlcmUgaXQncyBOL0EKICAgICAgLy8gaXMgbGVmdCBvZmYgdGhlIGNoYXJ0IGVudGlyZWx5IHJhdGhlciB0aGFuIGRyYXduIGFzIGEgMCBiYXIuCiAgICAgIGNvbnN0IHNvcnRlZCA9IGJyZWFrZG93bgogICAgICAgIC5maWx0ZXIoKHApID0+IChwLmFwcGxpY2FibGUgPyBwLmFwcGxpY2FibGVbbWV0cmljXSA6IChwW21ldHJpY10gIT09IG51bGwgJiYgcFttZXRyaWNdICE9PSB1bmRlZmluZWQpKSkKICAgICAgICAuc29ydCgoYSwgYikgPT4gYlttZXRyaWNdIC0gYVttZXRyaWNdKTsKICAgICAgY2hhcnRPckVtcHR5KCdicmVha2Rvd25DaGFydFdyYXAnLCAnYnJlYWtkb3duQ2FudmFzJywgc29ydGVkLmxlbmd0aCA+IDAsICdObyBkYXRhIGluIHRoaXMgcmFuZ2UgeWV0LicsICgpID0+IHsKICAgICAgICBDaGFydHMucGxhdGZvcm1CYXJDaGFydCgnYnJlYWtkb3duQ2FudmFzJywgewogICAgICAgICAgbGFiZWxzOiBzb3J0ZWQubWFwKChwKSA9PiBwLmxhYmVsKSwKICAgICAgICAgIGRhdGE6IHNvcnRlZC5tYXAoKHApID0+IHBbbWV0cmljXSksCiAgICAgICAgICBjb2xvcnM6IHNvcnRlZC5tYXAoKHApID0+IHAuY29sb3IpLAogICAgICAgICAgZm9ybWF0VmFsdWU6ICh2KSA9PiBmb3JtYXRNZXRyaWNWYWx1ZShtZXRyaWMsIHYpLAogICAgICAgIH0pOwogICAgICB9KTsKICAgIH0gZWxzZSB7CiAgICAgIHRpdGxlRWwudGV4dENvbnRlbnQgPSBgQ2FtcGFpZ24gUGVyZm9ybWFuY2Ug4oCUICR7bUxhYmVsfWA7CiAgICAgIGNvbnN0IGNhbXBhaWducyA9IGF3YWl0IEFwaS5jYW1wYWlnbkJyZWFrZG93bihmaWx0ZXJzKTsKICAgICAgY29uc3Qgd2l0aFZhbHVlID0gY2FtcGFpZ25zLmZpbHRlcigoYykgPT4gY1ttZXRyaWNdICE9PSBudWxsICYmIGNbbWV0cmljXSAhPT0gdW5kZWZpbmVkICYmIGNbbWV0cmljXSA+IDApOwogICAgICBjaGFydE9yRW1wdHkoJ2JyZWFrZG93bkNoYXJ0V3JhcCcsICdicmVha2Rvd25DYW52YXMnLCB3aXRoVmFsdWUubGVuZ3RoID4gMCwgJ05vIGNhbXBhaWduIGRhdGEgaW4gdGhpcyByYW5nZSB5ZXQuJywgKCkgPT4gewogICAgICAgIENoYXJ0cy5waWVDaGFydCgnYnJlYWtkb3duQ2FudmFzJywgewogICAgICAgICAgbGFiZWxzOiB3aXRoVmFsdWUubWFwKChjKSA9PiBjLmNhbXBhaWduX3R5cGUpLAogICAgICAgICAgZGF0YTogd2l0aFZhbHVlLm1hcCgoYykgPT4gY1ttZXRyaWNdKSwKICAgICAgICAgIGNvbG9yczogd2l0aFZhbHVlLm1hcCgoXywgaSkgPT4gQ2hhcnRzLnNlcmllc0NvbG9yKGkpKSwKICAgICAgICAgIGZvcm1hdFZhbHVlOiAodikgPT4gZm9ybWF0TWV0cmljVmFsdWUobWV0cmljLCB2KSwKICAgICAgICB9KTsKICAgICAgfSk7CiAgICB9CiAgfQoKICBhc3luYyBmdW5jdGlvbiByZW5kZXJDb250ZW50VHlwZUJyZWFrZG93bihmaWx0ZXJzKSB7CiAgICBjb25zdCBtTGFiZWwgPSBtZXRyaWNMYWJlbChtZXRyaWMpOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2NvbnRlbnRUeXBlQ2FyZFRpdGxlJykudGV4dENvbnRlbnQgPSBgQ29udGVudCBUeXBlIFBlcmZvcm1hbmNlIOKAlCAke21MYWJlbH1gOwogICAgY29uc3Qgcm93cyA9IGF3YWl0IEFwaS5jb250ZW50VHlwZUJyZWFrZG93bihmaWx0ZXJzKTsKICAgIGNvbnN0IHNvcnRlZCA9IHJvd3MuZmlsdGVyKChjKSA9PiBjW21ldHJpY10gIT09IG51bGwgJiYgY1ttZXRyaWNdICE9PSB1bmRlZmluZWQpLnNvcnQoKGEsIGIpID0+IGJbbWV0cmljXSAtIGFbbWV0cmljXSk7CiAgICBjaGFydE9yRW1wdHkoJ2NvbnRlbnRUeXBlQ2hhcnRXcmFwJywgJ2NvbnRlbnRUeXBlQ2FudmFzJywgc29ydGVkLmxlbmd0aCA+IDAsICdObyBkYXRhIGluIHRoaXMgcmFuZ2UgeWV0LicsICgpID0+IHsKICAgICAgQ2hhcnRzLnBsYXRmb3JtQmFyQ2hhcnQoJ2NvbnRlbnRUeXBlQ2FudmFzJywgewogICAgICAgIGxhYmVsczogc29ydGVkLm1hcCgoYykgPT4gYy5jb250ZW50X3R5cGUpLAogICAgICAgIGRhdGE6IHNvcnRlZC5tYXAoKGMpID0+IGNbbWV0cmljXSksCiAgICAgICAgY29sb3JzOiBzb3J0ZWQubWFwKChfLCBpKSA9PiBDaGFydHMuc2VyaWVzQ29sb3IoaSkpLAogICAgICAgIGZvcm1hdFZhbHVlOiAodikgPT4gZm9ybWF0TWV0cmljVmFsdWUobWV0cmljLCB2KSwKICAgICAgfSk7CiAgICB9KTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIHJlbmRlclRvcFBvc3RzKGZpbHRlcnMpIHsKICAgIGNvbnN0IHBvc3RzID0gYXdhaXQgQXBpLnRvcFBvc3RzKHsgLi4uZmlsdGVycywgc29ydEJ5OiBtZXRyaWMsIGxpbWl0OiAxMCB9KTsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndG9wUG9zdHNUYWJsZScpOwogICAgaWYgKCF3cmFwKSByZXR1cm47CiAgICBpZiAoIXBvc3RzLmxlbmd0aCkgewogICAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgICB3cmFwLmFwcGVuZENoaWxkKGVtcHR5U3RhdGUoewogICAgICAgIGljb246ICd0cm9waHknLAogICAgICAgIHRpdGxlOiAnTm8gcG9zdHMgaW4gdGhpcyByYW5nZSB5ZXQnLAogICAgICAgIG1lc3NhZ2U6ICdVcGxvYWQgYSB3ZWVrbHkgZXhwb3J0LCBvciB3aWRlbiB0aGUgZGF0ZSByYW5nZSwgdG8gc2VlIHRvcCBwZXJmb3JtZXJzIGhlcmUuJywKICAgICAgICBhY3Rpb25MYWJlbDogJ1VwbG9hZCBkYXRhJywKICAgICAgICBvbkFjdGlvbjogKCkgPT4gZG9jdW1lbnQucXVlcnlTZWxlY3RvcignLnRhYi1idG5bZGF0YS10YWI9InVwbG9hZCJdJyk/LmNsaWNrKCksCiAgICAgIH0pKTsKICAgICAgcmV0dXJuOwogICAgfQogICAgY29uc3QgcGxhdGZvcm1PcHRpb25zID0gKHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZSB8fCB7IHBsYXRmb3JtczogW10gfSkucGxhdGZvcm1zOwoKICAgIGNvbnN0IHRhYmxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGFibGUnKTsKICAgIHRhYmxlLmNsYXNzTmFtZSA9ICdkYXRhLXRhYmxlJzsKICAgIGNvbnN0IHRoZWFkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGhlYWQnKTsKICAgIGNvbnN0IGhlYWRUciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICBoZWFkVHIuYXBwZW5kKAogICAgICB0ZXh0RWwoJ3RoJywgJ1JhbmsnKSwKICAgICAgdGV4dEVsKCd0aCcsICdEYXRlJyksCiAgICAgIHRleHRFbCgndGgnLCAnUGxhdGZvcm0nKSwKICAgICAgdGV4dEVsKCd0aCcsICdDYW1wYWlnbicpLAogICAgICB0ZXh0RWwoJ3RoJywgJ0NvbnRlbnQgVHlwZScpLAogICAgICB0ZXh0RWwoJ3RoJywgJ0NhcHRpb24nKSwKICAgICAgdGV4dEVsKCd0aCcsIG1ldHJpY0xhYmVsKG1ldHJpYyksICdudW0nKQogICAgKTsKICAgIGhlYWRUci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3RoJywgJycpKTsKICAgIHRoZWFkLmFwcGVuZENoaWxkKGhlYWRUcik7CgogICAgY29uc3QgdGJvZHkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0Ym9keScpOwogICAgcG9zdHMuZm9yRWFjaCgocCwgaSkgPT4gewogICAgICBjb25zdCB0ciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAgIGNvbnN0IG1ldGEgPSBwbGF0Zm9ybU9wdGlvbnMuZmluZCgocGwpID0+IHBsLmlkID09PSBwLnBsYXRmb3JtKSB8fCB7IGxhYmVsOiBwLnBsYXRmb3JtLCBjb2xvcjogJyM5OTknIH07CiAgICAgIGNvbnN0IHBsYXRmb3JtVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBjb25zdCBwaWxsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICBwaWxsLmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1waWxsJzsKICAgICAgY29uc3QgZG90ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICBkb3QuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLWRvdCc7CiAgICAgIGRvdC5zdHlsZS5iYWNrZ3JvdW5kID0gbWV0YS5jb2xvcjsKICAgICAgcGlsbC5hcHBlbmQoZG90LCBkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShtZXRhLmxhYmVsKSk7CiAgICAgIHBsYXRmb3JtVGQuYXBwZW5kQ2hpbGQocGlsbCk7CgogICAgICBjb25zdCBjYXB0aW9uID0gcC5jYXB0aW9uIHx8ICcobm8gY2FwdGlvbiknOwogICAgICBjb25zdCB0cnVuY2F0ZWQgPSBjYXB0aW9uLmxlbmd0aCA+IDYwID8gYCR7Y2FwdGlvbi5zbGljZSgwLCA2MCl94oCmYCA6IGNhcHRpb247CiAgICAgIGNvbnN0IGNhcHRpb25UZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIGlmIChwLnBvc3RpbmdfbGluaykgewogICAgICAgIGNvbnN0IGxpbmsgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdhJyk7CiAgICAgICAgbGluay5jbGFzc05hbWUgPSAnY2FwdGlvbi1saW5rJzsKICAgICAgICBsaW5rLmhyZWYgPSBwLnBvc3RpbmdfbGluazsKICAgICAgICBsaW5rLnRhcmdldCA9ICdfYmxhbmsnOwogICAgICAgIGxpbmsucmVsID0gJ25vb3BlbmVyIG5vcmVmZXJyZXInOwogICAgICAgIGxpbmsudGl0bGUgPSBjYXB0aW9uOwogICAgICAgIGxpbmsuYXBwZW5kQ2hpbGQoZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUodHJ1bmNhdGVkKSk7CiAgICAgICAgY2FwdGlvblRkLmFwcGVuZENoaWxkKGxpbmspOwogICAgICB9IGVsc2UgewogICAgICAgIGNhcHRpb25UZC5hcHBlbmRDaGlsZChkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZSh0cnVuY2F0ZWQpKTsKICAgICAgICBjYXB0aW9uVGQudGl0bGUgPSBjYXB0aW9uOwogICAgICB9CgogICAgICB0ci5hcHBlbmQoCiAgICAgICAgdGV4dEVsKCd0ZCcsIGAjJHtpICsgMX1gKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgRm9ybWF0LmRhdGUocC5wdWJsaXNoX2RhdGUpKSwKICAgICAgICBwbGF0Zm9ybVRkLAogICAgICAgIHRleHRFbCgndGQnLCBwLmNhbXBhaWduX3R5cGUgfHwgJ+KAlCcpLAogICAgICAgIHRleHRFbCgndGQnLCBwLmNvbnRlbnRfdHlwZSB8fCAn4oCUJyksCiAgICAgICAgY2FwdGlvblRkLAogICAgICAgIHRleHRFbCgndGQnLCBmb3JtYXRNZXRyaWNWYWx1ZShtZXRyaWMsIHAubWV0cmljX3ZhbHVlKSwgJ251bScpCiAgICAgICk7CgogICAgICBjb25zdCBhY3Rpb25UZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIGNvbnN0IHZpZXdCdG4gPSBpY29uQnRuKCdidG4nLCAnZXllJywgJ1ZpZXcgRGV0YWlscycpOwogICAgICB2aWV3QnRuLmRpc2FibGVkID0gIXAucmF3X3Jvd19pZDsKICAgICAgdmlld0J0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IFJlY29yZHMub3BlblZpZXcocC5yYXdfcm93X2lkKSk7CiAgICAgIGFjdGlvblRkLmFwcGVuZENoaWxkKHZpZXdCdG4pOwogICAgICB0ci5hcHBlbmRDaGlsZChhY3Rpb25UZCk7CgogICAgICB0Ym9keS5hcHBlbmRDaGlsZCh0cik7CiAgICB9KTsKICAgIHRhYmxlLmFwcGVuZCh0aGVhZCwgdGJvZHkpOwogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGFibGUpOwogIH0KCiAgLyoqIE1ldHJpYyAob3IgYW55IGZpbHRlcikgY2hhbmdlZCBidXQgdGhlIHBsYXRmb3JtIOKAlCBhbmQgdGhlcmVmb3JlIHRoZSBhdmFpbGFibGUgbWV0cmljIGxpc3Qg4oCUIGRpZG4ndDogbm8gbmVlZCB0byByZS1mZXRjaCBtZXRyaWMtb3B0aW9ucyBvciByZWJ1aWxkIHRoZSBzaGVsbCwganVzdCByZWZyZXNoIHRoZSBkYXRhLiAqLwogIGFzeW5jIGZ1bmN0aW9uIHJlZnJlc2hGb3JNZXRyaWMoKSB7CiAgICBjb25zdCBmaWx0ZXJzID0gU3RhdGUuZ2V0RmlsdGVycygpOwogICAgY29uc3QgW3N1bW1hcnksIGZvbGxvd2Vyc10gPSBhd2FpdCBQcm9taXNlLmFsbChbCiAgICAgIEFwaS5tZXRyaWNTdW1tYXJ5KHsgLi4uZmlsdGVycywgbWV0cmljIH0pLAogICAgICBBcGkuZm9sbG93ZXJzS3BpcyhmaWx0ZXJzKSwKICAgIF0pOwogICAgcmVuZGVyS3BpcyhzdW1tYXJ5LCBmb2xsb3dlcnMpOwogICAgYXdhaXQgUHJvbWlzZS5hbGwoWwogICAgICByZW5kZXJUcmVuZChmaWx0ZXJzKSwgcmVuZGVyQnJlYWtkb3duKGZpbHRlcnMpLCByZW5kZXJDb250ZW50VHlwZUJyZWFrZG93bihmaWx0ZXJzKSwgcmVuZGVyVG9wUG9zdHMoZmlsdGVycyksCiAgICBdKTsKICB9CgogIGZ1bmN0aW9uIHNob3dTa2VsZXRvbnMoKSB7CiAgICBjb25zdCBrcGlHcmlkID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2twaUdyaWQnKTsKICAgIGlmIChrcGlHcmlkKSB7IGtwaUdyaWQuaW5uZXJIVE1MID0gJyc7IGtwaUdyaWQuYXBwZW5kQ2hpbGQoc2tlbGV0b25TdGF0R3JpZCg4KSk7IH0KICAgIFsndHJlbmRDaGFydFdyYXAnLCAnYnJlYWtkb3duQ2hhcnRXcmFwJywgJ2NvbnRlbnRUeXBlQ2hhcnRXcmFwJ10uZm9yRWFjaCgoaWQpID0+IHsKICAgICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGlkKTsKICAgICAgaWYgKHdyYXApIHsgd3JhcC5pbm5lckhUTUwgPSAnJzsgd3JhcC5hcHBlbmRDaGlsZChza2VsZXRvbkNoYXJ0KCkpOyB9CiAgICB9KTsKICAgIGNvbnN0IHRvcFBvc3RzID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3RvcFBvc3RzVGFibGUnKTsKICAgIGlmICh0b3BQb3N0cykgeyB0b3BQb3N0cy5pbm5lckhUTUwgPSAnJzsgdG9wUG9zdHMuYXBwZW5kQ2hpbGQoc2tlbGV0b25Sb3dzKDYpKTsgfQogIH0KCiAgYXN5bmMgZnVuY3Rpb24gcmVuZGVyKCkgewogICAgcm9vdCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd2aWV3LWRhc2hib2FyZCcpOwogICAgY29uc3QgZmlsdGVycyA9IFN0YXRlLmdldEZpbHRlcnMoKTsKICAgIGNvbnN0IHsgb3B0aW9ucyB9ID0gYXdhaXQgQXBpLm1ldHJpY09wdGlvbnMoZmlsdGVycy5wbGF0Zm9ybSk7CiAgICBtZXRyaWNPcHRpb25zID0gb3B0aW9uczsKICAgIGlmICghbWV0cmljT3B0aW9ucy5zb21lKChtKSA9PiBtLmtleSA9PT0gbWV0cmljKSkgewogICAgICBtZXRyaWMgPSBtZXRyaWNPcHRpb25zLmxlbmd0aCA/IG1ldHJpY09wdGlvbnNbMF0ua2V5IDogJ3ZpZXdzJzsKICAgIH0KICAgIHNoZWxsKCk7CiAgICBzaG93U2tlbGV0b25zKCk7CiAgICBhd2FpdCByZWZyZXNoRm9yTWV0cmljKCk7CiAgfQoKICByZXR1cm4geyByZW5kZXIgfTsKfSkoKTsKCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBEYXRhIFJlY29yZHMgdGFiOiBhIENSTS1zdHlsZSwgcGxhdGZvcm0tZ3JvdXBlZCBicm93c2VyIGJhY2tlZAogICBieSBwb3N0cy9wb3N0X21ldHJpY3MgKHRoZSBzYW1lIG5vcm1hbGl6ZWQgZGF0YSB0aGUgZGFzaGJvYXJkLAogICBjb21wYXJpc29ucywgYW5kIHJlcG9ydHMgcmVhZCkg4oCUICJBbGwgUGxhdGZvcm1zIiBzaG93cyBhIGNvbW1vbgogICBjcm9zcy1wbGF0Zm9ybSBzdW1tYXJ5LCBhIHNwZWNpZmljIHBsYXRmb3JtIHNob3dzIG9ubHkgdGhhdAogICBwbGF0Zm9ybSdzIGN1cmF0ZWQgbWV0cmljcy4gRXZlcnkgZmllbGQgb2YgYSByZWNvcmQgKGV4YWN0bHkgYXMKICAgaW1wb3J0ZWQpIGlzIGFsd2F5cyByZWFjaGFibGUgdmlhIFZpZXcvRWRpdCByZWdhcmRsZXNzIG9mIHRoZQogICB0YWJsZSdzIGN1cmF0aW9uLCB3aGljaCByZWFkcyB0aGUgcmF3X3Jvd3MgbWlycm9yIGFuZCwgb24gc2F2ZSwKICAgcmUtc3luY3MgcG9zdHMvcG9zdF9tZXRyaWNzIHNvIGV2ZXJ5IHZpZXcgcmVmbGVjdHMgdGhlIGNoYW5nZQogICBpbW1lZGlhdGVseS4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmNvbnN0IFJlY29yZHMgPSAoKCkgPT4gewogIGxldCByb290OwogIGxldCBwYWdlID0gMTsKICBjb25zdCBwYWdlU2l6ZSA9IDI1OwogIGxldCBzZWFyY2hWYWx1ZSA9ICcnOwogIGxldCBzZWFyY2hEZWJvdW5jZSA9IG51bGw7CiAgbGV0IG1vZGFsU3RhdGUgPSBudWxsOyAvLyB7IHJlY29yZCwgdmFsdWVzOiBbLi4uXSB9IOKAlCBFZGl0IG1vZGFsIG9ubHkKICBsZXQgY3VycmVudFJlc3VsdCA9IG51bGw7IC8vIGxhc3QtbG9hZGVkIHBhZ2UsIGtlcHQgc28gc29ydGluZyBjYW4gcmUtcmVuZGVyIHdpdGhvdXQgYSBuZXR3b3JrIHJvdW5kLXRyaXAKICBsZXQgc29ydFN0YXRlID0geyBrZXk6IG51bGwsIGRpcjogJ2FzYycsIHR5cGU6ICdzdHJpbmcnIH07CgogIC8qKiBTb3J0cyBhIGNvcHkgb2YgYHJvd3NgIGJ5IGEgKHBvc3NpYmx5IGRvdHRlZCwgZS5nLiAibWV0cmljcy5yZWFjaCIpIGtleSBwYXRoLiBOdWxscyBhbHdheXMgc29ydCBsYXN0IHJlZ2FyZGxlc3Mgb2YgZGlyZWN0aW9uLiAqLwogIGZ1bmN0aW9uIHNvcnRSb3dzKHJvd3MsIGtleSwgZGlyLCB0eXBlKSB7CiAgICBjb25zdCBmYWN0b3IgPSBkaXIgPT09ICdhc2MnID8gMSA6IC0xOwogICAgY29uc3QgcmVhZCA9IChyb3cpID0+IGtleS5zcGxpdCgnLicpLnJlZHVjZSgobywgaykgPT4gKG8gPT09IG51bGwgfHwgbyA9PT0gdW5kZWZpbmVkID8gdW5kZWZpbmVkIDogb1trXSksIHJvdyk7CiAgICByZXR1cm4gWy4uLnJvd3NdLnNvcnQoKGEsIGIpID0+IHsKICAgICAgY29uc3QgYXYgPSByZWFkKGEpOwogICAgICBjb25zdCBidiA9IHJlYWQoYik7CiAgICAgIGNvbnN0IGFNaXNzaW5nID0gYXYgPT09IG51bGwgfHwgYXYgPT09IHVuZGVmaW5lZCB8fCBhdiA9PT0gJyc7CiAgICAgIGNvbnN0IGJNaXNzaW5nID0gYnYgPT09IG51bGwgfHwgYnYgPT09IHVuZGVmaW5lZCB8fCBidiA9PT0gJyc7CiAgICAgIGlmIChhTWlzc2luZyAmJiBiTWlzc2luZykgcmV0dXJuIDA7CiAgICAgIGlmIChhTWlzc2luZykgcmV0dXJuIDE7CiAgICAgIGlmIChiTWlzc2luZykgcmV0dXJuIC0xOwogICAgICBpZiAodHlwZSA9PT0gJ251bWJlcicpIHJldHVybiAoYXYgLSBidikgKiBmYWN0b3I7CiAgICAgIHJldHVybiBTdHJpbmcoYXYpLmxvY2FsZUNvbXBhcmUoU3RyaW5nKGJ2KSkgKiBmYWN0b3I7CiAgICB9KTsKICB9CgogIC8qKiBBIDx0aD4gdGhhdCB0b2dnbGVzIGFzY2VuZGluZy9kZXNjZW5kaW5nIG9uIGNsaWNrIGFuZCBzaG93cyBhbiBhcnJvdyBvbiB3aGljaGV2ZXIgY29sdW1uIGlzIGFjdGl2ZSDigJQgc29ydHMgdGhlIGFscmVhZHktbG9hZGVkIHBhZ2UgaW5zdGFudGx5LCBubyByZWxvYWQuICovCiAgZnVuY3Rpb24gc29ydGFibGVIZWFkZXIobGFiZWwsIGtleSwgdHlwZSkgewogICAgY29uc3QgdGggPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0aCcpOwogICAgaWYgKHR5cGUgPT09ICdudW1iZXInKSB0aC5jbGFzc05hbWUgPSAnbnVtJzsKICAgIHRoLmNsYXNzTGlzdC5hZGQoJ3NvcnRhYmxlLXRoJyk7CiAgICBjb25zdCBpc0FjdGl2ZSA9IHNvcnRTdGF0ZS5rZXkgPT09IGtleTsKICAgIHRoLmFwcGVuZENoaWxkKGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKGxhYmVsKSk7CiAgICB0aC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3NwYW4nLCBpc0FjdGl2ZSA/IChzb3J0U3RhdGUuZGlyID09PSAnYXNjJyA/ICcg4oaRJyA6ICcg4oaTJykgOiAnIOKGlScsICdzb3J0LWFycm93JykpOwogICAgdGguYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7CiAgICAgIHNvcnRTdGF0ZSA9IHsga2V5LCBkaXI6IHNvcnRTdGF0ZS5rZXkgPT09IGtleSAmJiBzb3J0U3RhdGUuZGlyID09PSAnYXNjJyA/ICdkZXNjJyA6ICdhc2MnLCB0eXBlIH07CiAgICAgIGlmIChjdXJyZW50UmVzdWx0KSByZW5kZXJUYWJsZShjdXJyZW50UmVzdWx0KTsKICAgIH0pOwogICAgcmV0dXJuIHRoOwogIH0KCiAgZnVuY3Rpb24gcGxhdGZvcm1NZXRhKCkgewogICAgcmV0dXJuICh3aW5kb3cuX19maWx0ZXJPcHRpb25zQ2FjaGUgfHwgeyBwbGF0Zm9ybXM6IFtdIH0pLnBsYXRmb3JtczsKICB9CgogIGZ1bmN0aW9uIHBsYXRmb3JtTGFiZWwoaWQpIHsKICAgIGNvbnN0IG0gPSBwbGF0Zm9ybU1ldGEoKS5maW5kKChwKSA9PiBwLmlkID09PSBpZCk7CiAgICByZXR1cm4gbSA/IG0ubGFiZWwgOiBpZDsKICB9CgogIGZ1bmN0aW9uIHNoZWxsKCkgewogICAgcm9vdC5pbm5lckhUTUwgPSAnJzsKICAgIHJvb3QuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCAnRGF0YSBSZWNvcmRzJywgJ3NlY3Rpb24tdGl0bGUnKSk7CiAgICByb290LmFwcGVuZENoaWxkKHRleHRFbCgKICAgICAgJ2RpdicsCiAgICAgICdCcm93c2UgYnkgcGxhdGZvcm0gdG8gc2VlIG9ubHkgaXRzIG1ldHJpY3MsIG9yIHN0YXkgb24gQWxsIFBsYXRmb3JtcyBmb3IgYSBjcm9zcy1wbGF0Zm9ybSBzdW1tYXJ5LiBFdmVyeSByZWNvcmQgc3RheXMgZnVsbHkgZWRpdGFibGUg4oCUIFZpZXcgb3IgRWRpdCBhbHdheXMgb3BlbnMgZXZlcnkgZmllbGQgaW1wb3J0ZWQgZnJvbSB0aGUgc3ByZWFkc2hlZXQsIG5vdCBqdXN0IHdoYXTigJlzIGluIHRoZSB0YWJsZS4nLAogICAgICAnbXV0ZWQnCiAgICApKTsKCiAgICBjb25zdCB0b29sYmFyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0b29sYmFyLmNsYXNzTmFtZSA9ICdyZWNvcmRzLXRvb2xiYXInOwogICAgY29uc3QgcGlsbHMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHBpbGxzLmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1maWx0ZXItcGlsbHMnOwogICAgcGlsbHMuaWQgPSAncmVjb3Jkc1BsYXRmb3JtUGlsbHMnOwogICAgY29uc3Qgc2VhcmNoID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBzZWFyY2guY2xhc3NOYW1lID0gJ3JlY29yZHMtc2VhcmNoJzsKICAgIGNvbnN0IHNlYXJjaElucHV0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsKICAgIHNlYXJjaElucHV0LnR5cGUgPSAnc2VhcmNoJzsKICAgIHNlYXJjaElucHV0LnBsYWNlaG9sZGVyID0gJ1NlYXJjaCBjYXB0aW9ucywgY2FtcGFpZ25zLCBjb250ZW50IHR5cGXigKYnOwogICAgc2VhcmNoSW5wdXQuaWQgPSAncmVjb3Jkc1NlYXJjaElucHV0JzsKICAgIHNlYXJjaElucHV0LnZhbHVlID0gc2VhcmNoVmFsdWU7CiAgICBzZWFyY2hJbnB1dC5hZGRFdmVudExpc3RlbmVyKCdpbnB1dCcsICgpID0+IHsKICAgICAgY2xlYXJUaW1lb3V0KHNlYXJjaERlYm91bmNlKTsKICAgICAgc2VhcmNoRGVib3VuY2UgPSBzZXRUaW1lb3V0KCgpID0+IHsKICAgICAgICBzZWFyY2hWYWx1ZSA9IHNlYXJjaElucHV0LnZhbHVlOwogICAgICAgIHBhZ2UgPSAxOwogICAgICAgIGxvYWQoKTsKICAgICAgfSwgMzAwKTsKICAgIH0pOwogICAgc2VhcmNoLmFwcGVuZENoaWxkKHNlYXJjaElucHV0KTsKICAgIHRvb2xiYXIuYXBwZW5kKHBpbGxzLCBzZWFyY2gpOwogICAgcm9vdC5hcHBlbmRDaGlsZCh0b29sYmFyKTsKCiAgICBjb25zdCBleHBvcnRSb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGV4cG9ydFJvdy5jbGFzc05hbWUgPSAnZXhwb3J0LWJ1dHRvbnMnOwogICAgY29uc3QgZXhwb3J0Q3N2QnRuID0gaWNvbkJ0bignYnRuJywgJ2ZpbGUtZG93bicsICdFeHBvcnQgQ1NWJyk7CiAgICBleHBvcnRDc3ZCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB0cmlnZ2VyUmVjb3Jkc0V4cG9ydCgnY3N2JykpOwogICAgY29uc3QgZXhwb3J0WGxzeEJ0biA9IGljb25CdG4oJ2J0bicsICdmaWxlLXNwcmVhZHNoZWV0JywgJ0V4cG9ydCBFeGNlbCcpOwogICAgZXhwb3J0WGxzeEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHRyaWdnZXJSZWNvcmRzRXhwb3J0KCd4bHN4JykpOwogICAgY29uc3QgZGVkdXBlQnRuID0gaWNvbkJ0bignYnRuJywgJ2NvcHkteCcsICdSZW1vdmUgZHVwbGljYXRlcycpOwogICAgZGVkdXBlQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gaGFuZGxlRGVkdXBlKGRlZHVwZUJ0bikpOwogICAgZXhwb3J0Um93LmFwcGVuZChleHBvcnRDc3ZCdG4sIGV4cG9ydFhsc3hCdG4sIGRlZHVwZUJ0bik7CiAgICByb290LmFwcGVuZENoaWxkKGV4cG9ydFJvdyk7CgogICAgY29uc3QgY2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY2FyZC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICBjb25zdCB0YWJsZVdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRhYmxlV3JhcC5jbGFzc05hbWUgPSAndGFibGUtc2Nyb2xsJzsKICAgIHRhYmxlV3JhcC5pZCA9ICdyZWNvcmRzVGFibGVXcmFwJzsKICAgIGNhcmQuYXBwZW5kQ2hpbGQodGFibGVXcmFwKTsKICAgIGNvbnN0IHBhZ2VyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBwYWdlci5jbGFzc05hbWUgPSAncGFnaW5hdGlvbi1yb3cnOwogICAgcGFnZXIuaWQgPSAncmVjb3Jkc1BhZ2VyJzsKICAgIGNhcmQuYXBwZW5kQ2hpbGQocGFnZXIpOwogICAgcm9vdC5hcHBlbmRDaGlsZChjYXJkKTsKCiAgICByZW5kZXJQaWxscygpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyUGlsbHMoKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3JlY29yZHNQbGF0Zm9ybVBpbGxzJyk7CiAgICBpZiAoIXdyYXApIHJldHVybjsKICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CiAgICBjb25zdCBjdXJyZW50ID0gU3RhdGUuZ2V0RmlsdGVycygpLnBsYXRmb3JtIHx8ICdhbGwnOwogICAgY29uc3Qgb3B0aW9ucyA9IFt7IGlkOiAnYWxsJywgbGFiZWw6ICdBbGwgUGxhdGZvcm1zJywgY29sb3I6IG51bGwgfSwgLi4ucGxhdGZvcm1NZXRhKCldOwogICAgb3B0aW9ucy5mb3JFYWNoKChvcHQpID0+IHsKICAgICAgY29uc3QgYnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICAgIGJ0bi50eXBlID0gJ2J1dHRvbic7CiAgICAgIGJ0bi5jbGFzc0xpc3QudG9nZ2xlKCdpcy1hY3RpdmUnLCBjdXJyZW50ID09PSBvcHQuaWQpOwogICAgICBpZiAob3B0LmNvbG9yKSB7CiAgICAgICAgY29uc3QgZG90ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICAgIGRvdC5jbGFzc05hbWUgPSAncGxhdGZvcm0tZG90JzsKICAgICAgICBkb3Quc3R5bGUuYmFja2dyb3VuZCA9IG9wdC5jb2xvcjsKICAgICAgICBidG4uYXBwZW5kQ2hpbGQoZG90KTsKICAgICAgfQogICAgICBidG4uYXBwZW5kQ2hpbGQoZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUob3B0LmxhYmVsKSk7CiAgICAgIGJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsKICAgICAgICBpZiAoY3VycmVudCA9PT0gb3B0LmlkKSByZXR1cm47CiAgICAgICAgY29uc3QgZmlsdGVyU2VsZWN0ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZpbHRlclBsYXRmb3JtJyk7CiAgICAgICAgaWYgKGZpbHRlclNlbGVjdCkgZmlsdGVyU2VsZWN0LnZhbHVlID0gb3B0LmlkOwogICAgICAgIHBhZ2UgPSAxOwogICAgICAgIFN0YXRlLnNldEZpbHRlcnMoeyBwbGF0Zm9ybTogb3B0LmlkIH0pOwogICAgICB9KTsKICAgICAgd3JhcC5hcHBlbmRDaGlsZChidG4pOwogICAgfSk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiBsb2FkKCkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdyZWNvcmRzVGFibGVXcmFwJyk7CiAgICBpZiAod3JhcCkgeyB3cmFwLmlubmVySFRNTCA9ICcnOyB3cmFwLmFwcGVuZENoaWxkKHNrZWxldG9uUm93cyg4KSk7IH0KICAgIGNvbnN0IGZpbHRlcnMgPSBTdGF0ZS5nZXRGaWx0ZXJzKCk7CiAgICBjb25zdCByZXN1bHQgPSBhd2FpdCBBcGkucmVjb3Jkc1RhYmxlKHsgLi4uZmlsdGVycywgc2VhcmNoOiBzZWFyY2hWYWx1ZSwgcGFnZSwgcGFnZVNpemUgfSk7CiAgICByZW5kZXJUYWJsZShyZXN1bHQpOwogICAgcmVuZGVyUGFnZXIocmVzdWx0KTsKICB9CgogIC8qKiBSZWNvcmRzIGlzIHNlcnZlci1wYWdpbmF0ZWQvc2VhcmNoZWQsIHNvIGl0cyBleHBvcnQgaXMgYSBkaXJlY3QgbmF2aWdhdGlvbiB0byBhIGJhY2tlbmQgcm91dGUgdGhhdCByZXVzZXMgdGhlIGV4YWN0IHNhbWUgZmlsdGVyLWJ1aWxkaW5nIHRoZSBsaXN0IGVuZHBvaW50IGRvZXMg4oCUIGV4cG9ydHMgdGhlIGZ1bGwgbWF0Y2hpbmcgZGF0YXNldCwgbm90IGp1c3QgdGhlIGN1cnJlbnQgcGFnZS4gKi8KICBmdW5jdGlvbiB0cmlnZ2VyUmVjb3Jkc0V4cG9ydChmb3JtYXQpIHsKICAgIGNvbnN0IGZpbHRlcnMgPSBTdGF0ZS5nZXRGaWx0ZXJzKCk7CiAgICBjb25zdCBwYXJhbXMgPSBuZXcgVVJMU2VhcmNoUGFyYW1zKHsgLi4uZmlsdGVycywgc2VhcmNoOiBzZWFyY2hWYWx1ZSwgZm9ybWF0IH0pOwogICAgd2luZG93LmxvY2F0aW9uLmhyZWYgPSBgL2FwaS9yZWNvcmRzL2V4cG9ydD8ke3BhcmFtcy50b1N0cmluZygpfWA7CiAgfQoKICBmdW5jdGlvbiBjb2x1bW5MYWJlbHNGb3IocmVjb3JkKSB7CiAgICByZXR1cm4gcmVjb3JkLmhlYWRlcnMgJiYgcmVjb3JkLmhlYWRlcnMubGVuZ3RoCiAgICAgID8gcmVjb3JkLmhlYWRlcnMubWFwKChoKSA9PiAoaCAmJiBoLnRyaW0oKSA/IGggOiAnKHVubGFiZWxlZCBjb2x1bW4pJykpCiAgICAgIDogcmVjb3JkLnZhbHVlcy5tYXAoKF8sIGkpID0+IGBDb2x1bW4gJHtpICsgMX1gKTsKICB9CgogIC8qKiBHcm91cHMgYSByYXcgcmVjb3JkJ3MgZmllbGRzIGJ5IHRoZSBxdWFsaWZpZWQgaGVhZGVyJ3MgcGxhdGZvcm0tZ3JvdXAgcHJlZml4CiAgICAgIChlLmcuICJGQUNFQk9PSyDigJQgVmlld3MiKSwgc28gdGhlIFZpZXcvRWRpdCBwb3B1cCByZWFkcyBhcyBzZWN0aW9ucyBpbnN0ZWFkCiAgICAgIG9mIG9uZSBsb25nIGZsYXQgbGlzdCDigJQgZmFsbHMgYmFjayB0byBhIHNpbmdsZSAiRGV0YWlscyIgc2VjdGlvbiBmb3IKICAgICAgaWRlbnRpZmllciBjb2x1bW5zIGFuZCBmb3IgdGhlIHNpbXBsZSAob25lLXBsYXRmb3JtLXBlci1yb3cpIGZvcm1hdC4gKi8KICBmdW5jdGlvbiBncm91cEZpZWxkUm93cyhsYWJlbHMsIHZhbHVlcykgewogICAgY29uc3QgZ3JvdXBzID0gW107CiAgICBjb25zdCBpbmRleCA9IG5ldyBNYXAoKTsKICAgIGxhYmVscy5mb3JFYWNoKChsYWJlbCwgaWR4KSA9PiB7CiAgICAgIGNvbnN0IHNlcElkeCA9IGxhYmVsLmluZGV4T2YoJyDigJQgJyk7CiAgICAgIGNvbnN0IGdyb3VwTmFtZSA9IHNlcElkeCA+PSAwID8gbGFiZWwuc2xpY2UoMCwgc2VwSWR4KSA6ICdEZXRhaWxzJzsKICAgICAgY29uc3QgZmllbGRMYWJlbCA9IHNlcElkeCA+PSAwID8gbGFiZWwuc2xpY2Uoc2VwSWR4ICsgMykgOiBsYWJlbDsKICAgICAgaWYgKCFpbmRleC5oYXMoZ3JvdXBOYW1lKSkgewogICAgICAgIGluZGV4LnNldChncm91cE5hbWUsIHsgZ3JvdXA6IGdyb3VwTmFtZSwgZmllbGRzOiBbXSB9KTsKICAgICAgICBncm91cHMucHVzaChpbmRleC5nZXQoZ3JvdXBOYW1lKSk7CiAgICAgIH0KICAgICAgaW5kZXguZ2V0KGdyb3VwTmFtZSkuZmllbGRzLnB1c2goeyBpZHgsIGxhYmVsOiBmaWVsZExhYmVsIHx8IGBDb2x1bW4gJHtpZHggKyAxfWAsIHZhbHVlOiB2YWx1ZXNbaWR4XSB9KTsKICAgIH0pOwogICAgcmV0dXJuIGdyb3VwczsKICB9CgogIGZ1bmN0aW9uIHBsYXRmb3JtQmFkZ2VzKGlkcykgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgd3JhcC5zdHlsZS5kaXNwbGF5ID0gJ2ZsZXgnOwogICAgd3JhcC5zdHlsZS5mbGV4V3JhcCA9ICd3cmFwJzsKICAgIHdyYXAuc3R5bGUuZ2FwID0gJzRweCc7CiAgICBpZiAoIWlkcy5sZW5ndGgpIHJldHVybiB0ZXh0RWwoJ3NwYW4nLCAn4oCUJywgJ211dGVkJyk7CiAgICBjb25zdCBtZXRhID0gcGxhdGZvcm1NZXRhKCk7CiAgICBpZHMuZm9yRWFjaCgoaWQpID0+IHsKICAgICAgY29uc3QgbSA9IG1ldGEuZmluZCgocCkgPT4gcC5pZCA9PT0gaWQpIHx8IHsgbGFiZWw6IGlkLCBjb2xvcjogJyM5OTknIH07CiAgICAgIGNvbnN0IHBpbGwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICAgIHBpbGwuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLXBpbGwnOwogICAgICBjb25zdCBkb3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICAgIGRvdC5jbGFzc05hbWUgPSAncGxhdGZvcm0tZG90JzsKICAgICAgZG90LnN0eWxlLmJhY2tncm91bmQgPSBtLmNvbG9yOwogICAgICBwaWxsLmFwcGVuZChkb3QsIGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKG0ubGFiZWwpKTsKICAgICAgd3JhcC5hcHBlbmRDaGlsZChwaWxsKTsKICAgIH0pOwogICAgcmV0dXJuIHdyYXA7CiAgfQoKICBmdW5jdGlvbiBzdGF0dXNQaWxsKHN0YXR1cykgewogICAgY29uc3Qgc3BhbiA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgIHNwYW4uY2xhc3NOYW1lID0gYHN0YXR1cy1waWxsICR7c3RhdHVzfWA7CiAgICBzcGFuLnRleHRDb250ZW50ID0gc3RhdHVzID09PSAnZWRpdGVkJyA/ICdFZGl0ZWQnIDogJ09yaWdpbmFsJzsKICAgIHJldHVybiBzcGFuOwogIH0KCiAgZnVuY3Rpb24gbWV0cmljQ2VsbChrZXksIHZhbHVlKSB7CiAgICBpZiAoa2V5ID09PSAncG9zdGluZ19saW5rJykgewogICAgICBjb25zdCB0ZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIHRkLmNsYXNzTmFtZSA9ICdsaW5rLWNlbGwnOwogICAgICBpZiAodmFsdWUpIHsKICAgICAgICBjb25zdCBhID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYScpOwogICAgICAgIGEuaHJlZiA9IHZhbHVlOwogICAgICAgIGEudGFyZ2V0ID0gJ19ibGFuayc7CiAgICAgICAgYS5yZWwgPSAnbm9vcGVuZXIgbm9yZWZlcnJlcic7CiAgICAgICAgYS50ZXh0Q29udGVudCA9ICdPcGVuIOKGlyc7CiAgICAgICAgdGQuYXBwZW5kQ2hpbGQoYSk7CiAgICAgIH0gZWxzZSB7CiAgICAgICAgdGQuYXBwZW5kQ2hpbGQoZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUoJ+KAlCcpKTsKICAgICAgfQogICAgICByZXR1cm4gdGQ7CiAgICB9CiAgICBjb25zdCBkaXNwbGF5ID0ga2V5ID09PSAnd2F0Y2hfdGltZV9zZWNvbmRzJyA/IEZvcm1hdC5kdXJhdGlvbih2YWx1ZSkgOiBGb3JtYXQubnVtYmVyKHZhbHVlKTsKICAgIHJldHVybiB0ZXh0RWwoJ3RkJywgZGlzcGxheSwgJ251bScpOwogIH0KCiAgZnVuY3Rpb24gYWN0aW9uQnV0dG9ucyhyb3csIHBsYXRmb3JtKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB3cmFwLmNsYXNzTmFtZSA9ICdyb3ctYWN0aW9ucyc7CiAgICBjb25zdCB2aWV3QnRuID0gaWNvbkJ0bignYnRuJywgJ2V5ZScsICdWaWV3Jyk7CiAgICB2aWV3QnRuLmRpc2FibGVkID0gIXJvdy5yYXdSb3dJZDsKICAgIHZpZXdCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBvcGVuVmlldyhyb3cucmF3Um93SWQpKTsKICAgIGNvbnN0IGVkaXRCdG4gPSBpY29uQnRuKCdidG4nLCAncGVuY2lsJywgJ0VkaXQnKTsKICAgIGVkaXRCdG4uZGlzYWJsZWQgPSAhcm93LnJhd1Jvd0lkOwogICAgZWRpdEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IG9wZW5FZGl0b3Iocm93LnJhd1Jvd0lkKSk7CiAgICBjb25zdCBkZWxldGVCdG4gPSBpY29uQnRuKCdidG4gZGFuZ2VyJywgJ3RyYXNoLTInLCAnRGVsZXRlJyk7CiAgICBkZWxldGVCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBoYW5kbGVEZWxldGUocm93LCBwbGF0Zm9ybSkpOwogICAgd3JhcC5hcHBlbmQodmlld0J0biwgZWRpdEJ0biwgZGVsZXRlQnRuKTsKICAgIHJldHVybiB3cmFwOwogIH0KCiAgZnVuY3Rpb24gY2FwdGlvbkNlbGwoY2FwdGlvbikgewogICAgY29uc3QgdGV4dCA9IGNhcHRpb24gfHwgJyhubyBjYXB0aW9uKSc7CiAgICByZXR1cm4gdGV4dEVsKCd0ZCcsIHRleHQubGVuZ3RoID4gNzAgPyBgJHt0ZXh0LnNsaWNlKDAsIDcwKX3igKZgIDogdGV4dCk7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJTdW1tYXJ5VGFibGUocmVzdWx0KSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3JlY29yZHNUYWJsZVdyYXAnKTsKICAgIGlmICghcmVzdWx0LnJvd3MubGVuZ3RoKSB7CiAgICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQoZW1wdHlTdGF0ZSh7CiAgICAgICAgaWNvbjogJ2RhdGFiYXNlJywKICAgICAgICB0aXRsZTogJ05vIHJlY29yZHMgbWF0Y2ggdGhlc2UgZmlsdGVycyB5ZXQnLAogICAgICAgIG1lc3NhZ2U6ICdVcGxvYWQgYSB3ZWVrbHkgZXhwb3J0LCBvciB3aWRlbiB0aGUgZGF0ZSByYW5nZSwgdG8gc2VlIHJlY29yZHMgaGVyZS4nLAogICAgICAgIGFjdGlvbkxhYmVsOiAnVXBsb2FkIGRhdGEnLAogICAgICAgIG9uQWN0aW9uOiAoKSA9PiBkb2N1bWVudC5xdWVyeVNlbGVjdG9yKCcudGFiLWJ0bltkYXRhLXRhYj0idXBsb2FkIl0nKT8uY2xpY2soKSwKICAgICAgfSkpOwogICAgICByZXR1cm47CiAgICB9CiAgICBjb25zdCB0YWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICB0YWJsZS5jbGFzc05hbWUgPSAnZGF0YS10YWJsZSc7CiAgICBjb25zdCB0aGVhZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RoZWFkJyk7CiAgICBjb25zdCBoZWFkVHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgaGVhZFRyLmFwcGVuZCgKICAgICAgc29ydGFibGVIZWFkZXIoJ0RhdGUnLCAncHVibGlzaERhdGUnLCAnc3RyaW5nJyksCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdQbGF0Zm9ybXMnLCAncGxhdGZvcm1JZHMuMCcsICdzdHJpbmcnKSwKICAgICAgdGV4dEVsKCd0aCcsICdDYXB0aW9uJyksCiAgICAgIHRleHRFbCgndGgnLCAnQ2FtcGFpZ24nKSwKICAgICAgdGV4dEVsKCd0aCcsICdDb250ZW50IFR5cGUnKSwKICAgICAgdGV4dEVsKCd0aCcsICdTdGF0dXMnKSwKICAgICAgc29ydGFibGVIZWFkZXIoJ0xhc3QgVXBkYXRlZCcsICd1cGRhdGVkQXQnLCAnc3RyaW5nJyksCiAgICAgIHRleHRFbCgndGgnLCAnQWN0aW9ucycpCiAgICApOwogICAgdGhlYWQuYXBwZW5kQ2hpbGQoaGVhZFRyKTsKICAgIGNvbnN0IHRib2R5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGJvZHknKTsKICAgIGNvbnN0IHJvd3MgPSBzb3J0U3RhdGUua2V5ID8gc29ydFJvd3MocmVzdWx0LnJvd3MsIHNvcnRTdGF0ZS5rZXksIHNvcnRTdGF0ZS5kaXIsIHNvcnRTdGF0ZS50eXBlKSA6IHJlc3VsdC5yb3dzOwogICAgcm93cy5mb3JFYWNoKChyKSA9PiB7CiAgICAgIGNvbnN0IHRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgY29uc3QgcGxhdGZvcm1zVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBwbGF0Zm9ybXNUZC5hcHBlbmRDaGlsZChwbGF0Zm9ybUJhZGdlcyhyLnBsYXRmb3JtSWRzKSk7CiAgICAgIGNvbnN0IHN0YXR1c1RkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgc3RhdHVzVGQuYXBwZW5kQ2hpbGQoc3RhdHVzUGlsbChyLnN0YXR1cykpOwogICAgICBjb25zdCBhY3Rpb25zVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBhY3Rpb25zVGQuYXBwZW5kQ2hpbGQoYWN0aW9uQnV0dG9ucyhyLCAnYWxsJykpOwogICAgICB0ci5hcHBlbmQoCiAgICAgICAgdGV4dEVsKCd0ZCcsIEZvcm1hdC5kYXRlKHIucHVibGlzaERhdGUpKSwKICAgICAgICBwbGF0Zm9ybXNUZCwKICAgICAgICBjYXB0aW9uQ2VsbChyLmNhcHRpb24pLAogICAgICAgIHRleHRFbCgndGQnLCByLmNhbXBhaWduVHlwZSB8fCAn4oCUJyksCiAgICAgICAgdGV4dEVsKCd0ZCcsIHIuY29udGVudFR5cGUgfHwgJ+KAlCcpLAogICAgICAgIHN0YXR1c1RkLAogICAgICAgIHRleHRFbCgndGQnLCByLnVwZGF0ZWRBdCksCiAgICAgICAgYWN0aW9uc1RkCiAgICAgICk7CiAgICAgIHRib2R5LmFwcGVuZENoaWxkKHRyKTsKICAgIH0pOwogICAgdGFibGUuYXBwZW5kKHRoZWFkLCB0Ym9keSk7CiAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgd3JhcC5hcHBlbmRDaGlsZCh0YWJsZSk7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJQbGF0Zm9ybVRhYmxlKHJlc3VsdCkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdyZWNvcmRzVGFibGVXcmFwJyk7CiAgICBpZiAoIXJlc3VsdC5yb3dzLmxlbmd0aCkgewogICAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgICB3cmFwLmFwcGVuZENoaWxkKGVtcHR5U3RhdGUoewogICAgICAgIGljb246ICdkYXRhYmFzZScsCiAgICAgICAgdGl0bGU6IGBObyAke3BsYXRmb3JtTGFiZWwocmVzdWx0LnBsYXRmb3JtKX0gcmVjb3JkcyBtYXRjaCB0aGVzZSBmaWx0ZXJzIHlldGAsCiAgICAgICAgbWVzc2FnZTogJ1RyeSBhIGRpZmZlcmVudCBwbGF0Zm9ybSwgb3Igd2lkZW4gdGhlIGRhdGUgcmFuZ2UuJywKICAgICAgfSkpOwogICAgICByZXR1cm47CiAgICB9CiAgICBjb25zdCB0YWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICB0YWJsZS5jbGFzc05hbWUgPSAnZGF0YS10YWJsZSc7CiAgICBjb25zdCB0aGVhZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RoZWFkJyk7CiAgICBjb25zdCBoZWFkVHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgaGVhZFRyLmFwcGVuZChzb3J0YWJsZUhlYWRlcignRGF0ZScsICdwdWJsaXNoRGF0ZScsICdzdHJpbmcnKSwgdGV4dEVsKCd0aCcsICdDYXB0aW9uJyksIHRleHRFbCgndGgnLCAnQ2FtcGFpZ24nKSwgdGV4dEVsKCd0aCcsICdDb250ZW50IFR5cGUnKSk7CiAgICByZXN1bHQuY29sdW1ucy5mb3JFYWNoKChjKSA9PiB7CiAgICAgIGlmIChjLmtleSA9PT0gJ3Bvc3RpbmdfbGluaycpIHsKICAgICAgICBoZWFkVHIuYXBwZW5kQ2hpbGQodGV4dEVsKCd0aCcsIGMubGFiZWwpKTsKICAgICAgfSBlbHNlIHsKICAgICAgICBoZWFkVHIuYXBwZW5kQ2hpbGQoc29ydGFibGVIZWFkZXIoYy5sYWJlbCwgYG1ldHJpY3MuJHtjLmtleX1gLCAnbnVtYmVyJykpOwogICAgICB9CiAgICB9KTsKICAgIGhlYWRUci5hcHBlbmQodGV4dEVsKCd0aCcsICdTdGF0dXMnKSwgdGV4dEVsKCd0aCcsICdBY3Rpb25zJykpOwogICAgdGhlYWQuYXBwZW5kQ2hpbGQoaGVhZFRyKTsKICAgIGNvbnN0IHRib2R5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGJvZHknKTsKICAgIGNvbnN0IHJvd3MgPSBzb3J0U3RhdGUua2V5ID8gc29ydFJvd3MocmVzdWx0LnJvd3MsIHNvcnRTdGF0ZS5rZXksIHNvcnRTdGF0ZS5kaXIsIHNvcnRTdGF0ZS50eXBlKSA6IHJlc3VsdC5yb3dzOwogICAgcm93cy5mb3JFYWNoKChyKSA9PiB7CiAgICAgIGNvbnN0IHRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgdHIuYXBwZW5kKHRleHRFbCgndGQnLCBGb3JtYXQuZGF0ZShyLnB1Ymxpc2hEYXRlKSksIGNhcHRpb25DZWxsKHIuY2FwdGlvbiksIHRleHRFbCgndGQnLCByLmNhbXBhaWduVHlwZSB8fCAn4oCUJyksIHRleHRFbCgndGQnLCByLmNvbnRlbnRUeXBlIHx8ICfigJQnKSk7CiAgICAgIHJlc3VsdC5jb2x1bW5zLmZvckVhY2goKGMpID0+IHRyLmFwcGVuZENoaWxkKG1ldHJpY0NlbGwoYy5rZXksIHIubWV0cmljc1tjLmtleV0pKSk7CiAgICAgIGNvbnN0IHN0YXR1c1RkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgc3RhdHVzVGQuYXBwZW5kQ2hpbGQoc3RhdHVzUGlsbChyLnN0YXR1cykpOwogICAgICB0ci5hcHBlbmRDaGlsZChzdGF0dXNUZCk7CiAgICAgIGNvbnN0IGFjdGlvbnNUZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIGFjdGlvbnNUZC5hcHBlbmRDaGlsZChhY3Rpb25CdXR0b25zKHIsIHJlc3VsdC5wbGF0Zm9ybSkpOwogICAgICB0ci5hcHBlbmRDaGlsZChhY3Rpb25zVGQpOwogICAgICB0Ym9keS5hcHBlbmRDaGlsZCh0cik7CiAgICB9KTsKICAgIHRhYmxlLmFwcGVuZCh0aGVhZCwgdGJvZHkpOwogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGFibGUpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyVGFibGUocmVzdWx0KSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3JlY29yZHNUYWJsZVdyYXAnKTsKICAgIGlmICghd3JhcCkgcmV0dXJuOwogICAgY3VycmVudFJlc3VsdCA9IHJlc3VsdDsKICAgIGlmIChyZXN1bHQucGxhdGZvcm0gPT09ICdhbGwnKSByZW5kZXJTdW1tYXJ5VGFibGUocmVzdWx0KTsKICAgIGVsc2UgcmVuZGVyUGxhdGZvcm1UYWJsZShyZXN1bHQpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyUGFnZXIocmVzdWx0KSB7CiAgICBjb25zdCBwYWdlciA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdyZWNvcmRzUGFnZXInKTsKICAgIGlmICghcGFnZXIpIHJldHVybjsKICAgIHBhZ2VyLmlubmVySFRNTCA9ICcnOwogICAgY29uc3QgdG90YWxQYWdlcyA9IE1hdGgubWF4KDEsIE1hdGguY2VpbChyZXN1bHQudG90YWwgLyByZXN1bHQucGFnZVNpemUpKTsKICAgIGNvbnN0IHByZXZCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIHByZXZCdG4uY2xhc3NOYW1lID0gJ2J0bic7CiAgICBwcmV2QnRuLnRleHRDb250ZW50ID0gJ1ByZXZpb3VzJzsKICAgIHByZXZCdG4uZGlzYWJsZWQgPSByZXN1bHQucGFnZSA8PSAxOwogICAgcHJldkJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsgcGFnZSAtPSAxOyBsb2FkKCk7IH0pOwogICAgY29uc3QgbmV4dEJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgbmV4dEJ0bi5jbGFzc05hbWUgPSAnYnRuJzsKICAgIG5leHRCdG4udGV4dENvbnRlbnQgPSAnTmV4dCc7CiAgICBuZXh0QnRuLmRpc2FibGVkID0gcmVzdWx0LnBhZ2UgPj0gdG90YWxQYWdlczsKICAgIG5leHRCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7IHBhZ2UgKz0gMTsgbG9hZCgpOyB9KTsKICAgIHBhZ2VyLmFwcGVuZChwcmV2QnRuLCB0ZXh0RWwoJ3NwYW4nLCBgUGFnZSAke3Jlc3VsdC5wYWdlfSBvZiAke3RvdGFsUGFnZXN9IOKAlCAke3Jlc3VsdC50b3RhbH0gcmVjb3JkKHMpYCksIG5leHRCdG4pOwogIH0KCiAgLyoqICJSZW1vdmUgZHVwbGljYXRlcyI6IGNvbGxhcHNlIHJlY29yZHMgdGhhdCBzaGFyZSBhIHB1Ymxpc2ggZGF0ZSArIGNhcHRpb24KICAgICAgKHRoZSBjb3BpZXMgbGVmdCBiZWhpbmQgd2hlbiBhIHdlZWtseSBzaGVldCBpcyByZS11cGxvYWRlZCBhbmQgYSBwb3N0J3MKICAgICAgbnVtYmVycyBjaGFuZ2VkKSwga2VlcGluZyB0aGUgbW9zdCByZWNlbnRseSBpbXBvcnRlZCBvbmUuIFByZXZpZXdzIHRoZQogICAgICBjb3VudCBmaXJzdCwgdGhlbiBkZWxldGVzIG9uIGNvbmZpcm0g4oCUIHRoZSBvcmlnaW5hbHMgc3RheSBpbiBVcGxvYWQKICAgICAgSGlzdG9yeSdzIHJhdy1yb3cgdmlld2VyIGVpdGhlciB3YXkuICovCiAgYXN5bmMgZnVuY3Rpb24gaGFuZGxlRGVkdXBlKGJ0bikgewogICAgaWYgKGJ0biAmJiBidG4uZGlzYWJsZWQpIHJldHVybjsKICAgIHRyeSB7CiAgICAgIGlmIChidG4pIGJ0bi5kaXNhYmxlZCA9IHRydWU7CiAgICAgIGNvbnN0IHByZXZpZXcgPSBhd2FpdCBBcGkuZHVwbGljYXRlUmVjb3Jkc1ByZXZpZXcoKTsKICAgICAgaWYgKCFwcmV2aWV3LnJlbW92ZUNvdW50KSB7CiAgICAgICAgVG9hc3Quc2hvdygnTm8gZHVwbGljYXRlIHJlY29yZHMgZm91bmQuJywgJ3N1Y2Nlc3MnKTsKICAgICAgICByZXR1cm47CiAgICAgIH0KICAgICAgY29uc3QgZWRpdGVkTm90ZSA9IHByZXZpZXcuZWRpdGVkSW5SZW1vdmVDb3VudAogICAgICAgID8gYFxuXG4ke3ByZXZpZXcuZWRpdGVkSW5SZW1vdmVDb3VudH0gb2YgdGhlIGNvcGllcyB0byBiZSByZW1vdmVkIHdlcmUgaGFuZC1lZGl0ZWQgYWZ0ZXIgaW1wb3J0IOKAlCB0aGUgbW9zdCByZWNlbnQgaW1wb3J0IGlzIHN0aWxsIHRoZSBvbmUga2VwdC5gCiAgICAgICAgOiAnJzsKICAgICAgY29uc3QgbWVzc2FnZSA9CiAgICAgICAgYEZvdW5kICR7cHJldmlldy5yZW1vdmVDb3VudH0gZHVwbGljYXRlIHJlY29yZChzKSBhY3Jvc3MgJHtwcmV2aWV3Lmdyb3VwQ291bnR9IHBvc3QocykgYCArCiAgICAgICAgYChzYW1lIHB1Ymxpc2ggZGF0ZSBhbmQgY2FwdGlvbikuIFJlbW92ZSB0aGUgb2xkZXIgY29waWVzIGFuZCBrZWVwIHRoZSBtb3N0IHJlY2VudGx5IHVwbG9hZGVkIG9uZSBmb3IgZWFjaD9gICsKICAgICAgICBgXG5cblRoaXMgdXBkYXRlcyB0aGUgZGFzaGJvYXJkLCBjb21wYXJpc29ucywgYW5kIHJlcG9ydHMuIEV2ZXJ5IG9yaWdpbmFsIGltcG9ydCBzdGF5cyBpbiBVcGxvYWQgSGlzdG9yeS5gICsKICAgICAgICBlZGl0ZWROb3RlOwogICAgICBpZiAoIXdpbmRvdy5jb25maXJtKG1lc3NhZ2UpKSByZXR1cm47CiAgICAgIGNvbnN0IHJlc3VsdCA9IGF3YWl0IEFwaS5yZXNvbHZlRHVwbGljYXRlUmVjb3JkcygpOwogICAgICBUb2FzdC5zaG93KGBSZW1vdmVkICR7cmVzdWx0LnJlbW92ZWRDb3VudH0gZHVwbGljYXRlIHJlY29yZChzKS5gLCAnc3VjY2VzcycpOwogICAgICBwYWdlID0gMTsKICAgICAgYXdhaXQgbG9hZCgpOwogICAgICB3aW5kb3cuZGlzcGF0Y2hFdmVudChuZXcgQ3VzdG9tRXZlbnQoJ2xyczpkYXRhLXVwZGF0ZWQnKSk7CiAgICB9IGNhdGNoIChlcnIpIHsKICAgICAgVG9hc3Quc2hvdyhlcnIubWVzc2FnZSB8fCAnRHVwbGljYXRlIGNsZWFudXAgZmFpbGVkLicsICdlcnJvcicpOwogICAgfSBmaW5hbGx5IHsKICAgICAgaWYgKGJ0bikgYnRuLmRpc2FibGVkID0gZmFsc2U7CiAgICB9CiAgfQoKICBhc3luYyBmdW5jdGlvbiBoYW5kbGVEZWxldGUocm93LCBwbGF0Zm9ybSkgewogICAgY29uc3QgY2FwdGlvbiA9IChyb3cuY2FwdGlvbiB8fCAnKG5vIGNhcHRpb24pJykuc2xpY2UoMCwgNjApOwogICAgY29uc3QgbWVzc2FnZSA9IHBsYXRmb3JtID09PSAnYWxsJwogICAgICA/IGBEZWxldGUgdGhpcyBlbnRpcmUgcmVjb3JkIOKAlCAiJHtjYXB0aW9ufSIg4oCUIGFjcm9zcyBldmVyeSBwbGF0Zm9ybT8gSXRzIG9yaWdpbmFsIGltcG9ydCBzdGF5cyBpbiBVcGxvYWQgSGlzdG9yeSwgYnV0IGl0IHdpbGwgZGlzYXBwZWFyIGZyb20gdGhlIGRhc2hib2FyZCwgY29tcGFyaXNvbnMsIGFuZCByZXBvcnRzLmAKICAgICAgOiBgUmVtb3ZlIHRoaXMgcmVjb3JkJ3MgJHtwbGF0Zm9ybUxhYmVsKHBsYXRmb3JtKX0gZGF0YSDigJQgIiR7Y2FwdGlvbn0iPyBJZiB0aGlzIGlzIGl0cyBvbmx5IHBsYXRmb3JtLCB0aGUgd2hvbGUgcmVjb3JkIHdpbGwgYmUgcmVtb3ZlZCBmcm9tIHRoZSBkYXNoYm9hcmQuYDsKICAgIGlmICghd2luZG93LmNvbmZpcm0obWVzc2FnZSkpIHJldHVybjsKICAgIHRyeSB7CiAgICAgIGlmIChwbGF0Zm9ybSA9PT0gJ2FsbCcpIGF3YWl0IEFwaS5kZWxldGVSZWNvcmRQb3N0KHJvdy5wb3N0SWQpOwogICAgICBlbHNlIGF3YWl0IEFwaS5kZWxldGVSZWNvcmRQbGF0Zm9ybShyb3cucG9zdElkLCBwbGF0Zm9ybSk7CiAgICAgIFRvYXN0LnNob3coJ1JlY29yZCBkZWxldGVkLicsICdzdWNjZXNzJyk7CiAgICAgIGF3YWl0IGxvYWQoKTsKICAgICAgd2luZG93LmRpc3BhdGNoRXZlbnQobmV3IEN1c3RvbUV2ZW50KCdscnM6ZGF0YS11cGRhdGVkJykpOwogICAgfSBjYXRjaCAoZXJyKSB7CiAgICAgIFRvYXN0LnNob3coZXJyLm1lc3NhZ2UsICdlcnJvcicpOwogICAgfQogIH0KCiAgZnVuY3Rpb24gcmVtb3ZlRXhpc3RpbmdPdmVybGF5KCkgewogICAgY29uc3Qgb3ZlcmxheSA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdyZWNvcmRNb2RhbE92ZXJsYXknKTsKICAgIGlmIChvdmVybGF5KSBvdmVybGF5LnJlbW92ZSgpOwogIH0KCiAgZnVuY3Rpb24gY2xvc2VNb2RhbCgpIHsKICAgIHJlbW92ZUV4aXN0aW5nT3ZlcmxheSgpOwogICAgbW9kYWxTdGF0ZSA9IG51bGw7CiAgfQoKICAvLyBPbmx5IGNsZWFycyB0aGUgc3RhbGUgRE9NIG5vZGUg4oCUIE5PVCBtb2RhbFN0YXRlLiByZW5kZXJFZGl0TW9kYWwgcmVhZHMKICAvLyBtb2RhbFN0YXRlIHJpZ2h0IGFmdGVyIGNhbGxpbmcgdGhpcyB0byBidWlsZCB0aGUgZm9ybTsgaWYgdGhpcyBjYWxsZWQKICAvLyB0aGUgcmVhbCBjbG9zZU1vZGFsKCkgKGFzIGl0IHVzZWQgdG8pLCB0aGF0IHJlc2V0IG1vZGFsU3RhdGUgdG8gbnVsbCBvdXQKICAvLyBmcm9tIHVuZGVyIGl0IGJlZm9yZSB0aGUgcmVhZCwgd2hpY2ggaXMgZXhhY3RseSB3aHkgRWRpdCB3YXMgYnJva2VuLgogIGZ1bmN0aW9uIG1vZGFsU2hlbGwodGl0bGVUZXh0KSB7CiAgICByZW1vdmVFeGlzdGluZ092ZXJsYXkoKTsKICAgIGNvbnN0IG92ZXJsYXkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIG92ZXJsYXkuY2xhc3NOYW1lID0gJ21vZGFsLW92ZXJsYXknOwogICAgb3ZlcmxheS5pZCA9ICdyZWNvcmRNb2RhbE92ZXJsYXknOwogICAgb3ZlcmxheS5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIChlKSA9PiB7IGlmIChlLnRhcmdldCA9PT0gb3ZlcmxheSkgY2xvc2VNb2RhbCgpOyB9KTsKICAgIGNvbnN0IHBhbmVsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBwYW5lbC5jbGFzc05hbWUgPSAnbW9kYWwtcGFuZWwgd2lkZSc7CiAgICBwYW5lbC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2gyJywgdGl0bGVUZXh0KSk7CiAgICBvdmVybGF5LmFwcGVuZENoaWxkKHBhbmVsKTsKICAgIHJldHVybiB7IG92ZXJsYXksIHBhbmVsIH07CiAgfQoKICBmdW5jdGlvbiByZWNvcmRTdWJ0aXRsZShyKSB7CiAgICByZXR1cm4gYFNoZWV0ICIke3Iuc2hlZXROYW1lfSIsIHJvdyAke3Iucm93TnVtYmVyfSR7ci5wb3N0SWQgPyBgIOKAlCBsaW5rZWQgdG8gZGFzaGJvYXJkIHBvc3QgIyR7ci5wb3N0SWR9YCA6ICcg4oCUIG5vdCBwYXJ0IG9mIHRoZSBkYXNoYm9hcmQgKGUuZy4gbmVlZHMgYSB2YWxpZCBkYXRlKSd9YDsKICB9CgogIC8vIC0tLS0tLS0tLS0gVmlldyBwb3B1cDogcmVhZC1vbmx5LCBldmVyeSBmaWVsZCwgZ3JvdXBlZCBpbnRvIHNlY3Rpb25zIC0tLS0tLS0tLS0KICBhc3luYyBmdW5jdGlvbiBvcGVuVmlldyhpZCkgewogICAgY29uc3QgcmVjb3JkID0gYXdhaXQgQXBpLmdldFJlY29yZChpZCk7CiAgICBjb25zdCB7IG92ZXJsYXksIHBhbmVsIH0gPSBtb2RhbFNoZWxsKCdSZWNvcmQgZGV0YWlscycpOwogICAgcGFuZWwuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCByZWNvcmRTdWJ0aXRsZShyZWNvcmQpLCAnbW9kYWwtc3ViJykpOwoKICAgIGNvbnN0IGdyb3VwcyA9IGdyb3VwRmllbGRSb3dzKGNvbHVtbkxhYmVsc0ZvcihyZWNvcmQpLCByZWNvcmQudmFsdWVzKTsKICAgIGdyb3Vwcy5mb3JFYWNoKChnKSA9PiB7CiAgICAgIGNvbnN0IHNlY3Rpb24gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgc2VjdGlvbi5jbGFzc05hbWUgPSAncmVjb3JkLXNlY3Rpb24nOwogICAgICBzZWN0aW9uLmFwcGVuZENoaWxkKHRleHRFbCgnaDQnLCBnLmdyb3VwKSk7CiAgICAgIGNvbnN0IGdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgZ3JpZC5jbGFzc05hbWUgPSAnZm9ybS1ncmlkJzsKICAgICAgZy5maWVsZHMuZm9yRWFjaCgoZikgPT4gewogICAgICAgIGNvbnN0IGZpZWxkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgICAgZmllbGQuY2xhc3NOYW1lID0gJ3ZpZXctZmllbGQnOwogICAgICAgIGZpZWxkLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgZi5sYWJlbCwgJ3ZpZXctbGFiZWwnKSk7CiAgICAgICAgY29uc3QgdmFsID0gZi52YWx1ZSA9PT0gdW5kZWZpbmVkIHx8IGYudmFsdWUgPT09IG51bGwgfHwgZi52YWx1ZSA9PT0gJycgPyAn4oCUJyA6IFN0cmluZyhmLnZhbHVlKTsKICAgICAgICBmaWVsZC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIHZhbCwgJ3ZpZXctdmFsdWUnKSk7CiAgICAgICAgZ3JpZC5hcHBlbmRDaGlsZChmaWVsZCk7CiAgICAgIH0pOwogICAgICBzZWN0aW9uLmFwcGVuZENoaWxkKGdyaWQpOwogICAgICBwYW5lbC5hcHBlbmRDaGlsZChzZWN0aW9uKTsKICAgIH0pOwoKICAgIGNvbnN0IGFjdGlvbnMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGFjdGlvbnMuY2xhc3NOYW1lID0gJ21vZGFsLWFjdGlvbnMnOwogICAgY29uc3QgYnRuUm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBidG5Sb3cuY2xhc3NOYW1lID0gJ2J0bi1yb3cnOwogICAgY29uc3QgY2xvc2VCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIGNsb3NlQnRuLmNsYXNzTmFtZSA9ICdidG4nOwogICAgY2xvc2VCdG4udGV4dENvbnRlbnQgPSAnQ2xvc2UnOwogICAgY2xvc2VCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBjbG9zZU1vZGFsKTsKICAgIGNvbnN0IGVkaXRCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIGVkaXRCdG4uY2xhc3NOYW1lID0gJ2J0biBwcmltYXJ5JzsKICAgIGVkaXRCdG4udGV4dENvbnRlbnQgPSAnRWRpdCB0aGlzIHJlY29yZCc7CiAgICBlZGl0QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gb3BlbkVkaXRvcihyZWNvcmQuaWQpKTsKICAgIGJ0blJvdy5hcHBlbmQoY2xvc2VCdG4sIGVkaXRCdG4pOwogICAgYWN0aW9ucy5hcHBlbmRDaGlsZChidG5Sb3cpOwogICAgcGFuZWwuYXBwZW5kQ2hpbGQoYWN0aW9ucyk7CgogICAgZG9jdW1lbnQuYm9keS5hcHBlbmRDaGlsZChvdmVybGF5KTsKICB9CgogIC8vIC0tLS0tLS0tLS0gRWRpdCBwb3B1cDogZXZlcnkgZmllbGQsIGdyb3VwZWQgaW50byBzZWN0aW9ucywgYWxsIGVkaXRhYmxlIC0tLS0tLS0tLS0KICBhc3luYyBmdW5jdGlvbiBvcGVuRWRpdG9yKGlkKSB7CiAgICBjb25zdCByZWNvcmQgPSBhd2FpdCBBcGkuZ2V0UmVjb3JkKGlkKTsKICAgIG1vZGFsU3RhdGUgPSB7IHJlY29yZCwgdmFsdWVzOiBbLi4ucmVjb3JkLnZhbHVlc10gfTsKICAgIHJlbmRlckVkaXRNb2RhbCgpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyRWRpdE1vZGFsKCkgewogICAgY29uc3QgciA9IG1vZGFsU3RhdGUucmVjb3JkOwogICAgY29uc3QgeyBvdmVybGF5LCBwYW5lbCB9ID0gbW9kYWxTaGVsbCgnRWRpdCByZWNvcmQnKTsKICAgIHBhbmVsLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgcmVjb3JkU3VidGl0bGUociksICdtb2RhbC1zdWInKSk7CgogICAgY29uc3QgZ3JvdXBzID0gZ3JvdXBGaWVsZFJvd3MoY29sdW1uTGFiZWxzRm9yKHIpLCBtb2RhbFN0YXRlLnZhbHVlcyk7CiAgICBncm91cHMuZm9yRWFjaCgoZykgPT4gewogICAgICBjb25zdCBzZWN0aW9uID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIHNlY3Rpb24uY2xhc3NOYW1lID0gJ3JlY29yZC1zZWN0aW9uJzsKICAgICAgc2VjdGlvbi5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2g0JywgZy5ncm91cCkpOwogICAgICBjb25zdCBncmlkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIGdyaWQuY2xhc3NOYW1lID0gJ2Zvcm0tZ3JpZCc7CiAgICAgIGcuZmllbGRzLmZvckVhY2goKGYpID0+IHsKICAgICAgICBjb25zdCBmaWVsZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICAgIGZpZWxkLmNsYXNzTmFtZSA9ICdmb3JtLWZpZWxkJzsKICAgICAgICBmaWVsZC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2xhYmVsJywgZi5sYWJlbCkpOwogICAgICAgIGNvbnN0IHN0clZhbCA9IGYudmFsdWUgPT09IHVuZGVmaW5lZCB8fCBmLnZhbHVlID09PSBudWxsID8gJycgOiBTdHJpbmcoZi52YWx1ZSk7CiAgICAgICAgY29uc3QgaXNMb25nID0gc3RyVmFsLmxlbmd0aCA+IDgwOwogICAgICAgIGNvbnN0IGlucHV0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudChpc0xvbmcgPyAndGV4dGFyZWEnIDogJ2lucHV0Jyk7CiAgICAgICAgaWYgKCFpc0xvbmcpIGlucHV0LnR5cGUgPSAndGV4dCc7CiAgICAgICAgZWxzZSBmaWVsZC5zdHlsZS5ncmlkQ29sdW1uID0gJzEgLyAtMSc7CiAgICAgICAgaW5wdXQudmFsdWUgPSBzdHJWYWw7CiAgICAgICAgaW5wdXQuYWRkRXZlbnRMaXN0ZW5lcignaW5wdXQnLCAoKSA9PiB7IG1vZGFsU3RhdGUudmFsdWVzW2YuaWR4XSA9IGlucHV0LnZhbHVlOyB9KTsKICAgICAgICBmaWVsZC5hcHBlbmRDaGlsZChpbnB1dCk7CiAgICAgICAgZ3JpZC5hcHBlbmRDaGlsZChmaWVsZCk7CiAgICAgIH0pOwogICAgICBzZWN0aW9uLmFwcGVuZENoaWxkKGdyaWQpOwogICAgICBwYW5lbC5hcHBlbmRDaGlsZChzZWN0aW9uKTsKICAgIH0pOwoKICAgIGNvbnN0IGFjdGlvbnMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGFjdGlvbnMuY2xhc3NOYW1lID0gJ21vZGFsLWFjdGlvbnMnOwogICAgY29uc3QgZXJyb3JNc2cgPSB0ZXh0RWwoJ3NwYW4nLCAnJywgJ211dGVkJyk7CiAgICBlcnJvck1zZy5pZCA9ICdtb2RhbEVycm9yTXNnJzsKICAgIGNvbnN0IGJ0blJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYnRuUm93LmNsYXNzTmFtZSA9ICdidG4tcm93JzsKICAgIGNvbnN0IGNhbmNlbEJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgY2FuY2VsQnRuLmNsYXNzTmFtZSA9ICdidG4nOwogICAgY2FuY2VsQnRuLnRleHRDb250ZW50ID0gJ0NhbmNlbCc7CiAgICBjYW5jZWxCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBjbG9zZU1vZGFsKTsKICAgIGNvbnN0IHNhdmVCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIHNhdmVCdG4uY2xhc3NOYW1lID0gJ2J0biBwcmltYXJ5JzsKICAgIHNhdmVCdG4udGV4dENvbnRlbnQgPSAnU2F2ZSBjaGFuZ2VzJzsKICAgIHNhdmVCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBzYXZlRWRpdChzYXZlQnRuKSk7CiAgICBidG5Sb3cuYXBwZW5kKGNhbmNlbEJ0biwgc2F2ZUJ0bik7CiAgICBhY3Rpb25zLmFwcGVuZChlcnJvck1zZywgYnRuUm93KTsKICAgIHBhbmVsLmFwcGVuZENoaWxkKGFjdGlvbnMpOwoKICAgIGRvY3VtZW50LmJvZHkuYXBwZW5kQ2hpbGQob3ZlcmxheSk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiBzYXZlRWRpdChidG4pIHsKICAgIGNvbnN0IGVycm9yRWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbW9kYWxFcnJvck1zZycpOwogICAgZXJyb3JFbC50ZXh0Q29udGVudCA9ICcnOwogICAgYnRuLmRpc2FibGVkID0gdHJ1ZTsKICAgIGJ0bi50ZXh0Q29udGVudCA9ICdTYXZpbmfigKYnOwogICAgdHJ5IHsKICAgICAgYXdhaXQgQXBpLnVwZGF0ZVJlY29yZChtb2RhbFN0YXRlLnJlY29yZC5pZCwgbW9kYWxTdGF0ZS52YWx1ZXMpOwogICAgICBUb2FzdC5zaG93KCdSZWNvcmQgdXBkYXRlZC4nLCAnc3VjY2VzcycpOwogICAgICBjbG9zZU1vZGFsKCk7CiAgICAgIGF3YWl0IGxvYWQoKTsKICAgICAgd2luZG93LmRpc3BhdGNoRXZlbnQobmV3IEN1c3RvbUV2ZW50KCdscnM6ZGF0YS11cGRhdGVkJykpOwogICAgfSBjYXRjaCAoZXJyKSB7CiAgICAgIGVycm9yRWwudGV4dENvbnRlbnQgPSBlcnIubWVzc2FnZTsKICAgICAgZXJyb3JFbC5zdHlsZS5jb2xvciA9ICd2YXIoLS1zdGF0dXMtY3JpdGljYWwpJzsKICAgICAgYnRuLmRpc2FibGVkID0gZmFsc2U7CiAgICAgIGJ0bi50ZXh0Q29udGVudCA9ICdTYXZlIGNoYW5nZXMnOwogICAgfQogIH0KCiAgYXN5bmMgZnVuY3Rpb24gcmVuZGVyKCkgewogICAgcm9vdCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd2aWV3LXJlY29yZHMnKTsKICAgIHBhZ2UgPSAxOwogICAgc2hlbGwoKTsKICAgIGF3YWl0IGxvYWQoKTsKICB9CgogIHJldHVybiB7IHJlbmRlciwgcmVsb2FkOiBsb2FkLCBvcGVuVmlldyB9Owp9KSgpOwoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIENvbXBhcmlzb25zIHRhYjogd2Vlay12cy13ZWVrLCBjdXN0b20gcmFuZ2UsIG1vbnRobHksCiAgIHF1YXJ0ZXJseSwgWVRELgogICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KY29uc3QgQ29tcGFyaXNvbiA9ICgoKSA9PiB7CiAgY29uc3QgTU9ERVMgPSBbCiAgICB7IGtleTogJ3BsYXRmb3JtcycsIGxhYmVsOiAnQWxsIFBsYXRmb3JtcycgfSwKICAgIHsga2V5OiAnd2VlaycsIGxhYmVsOiAnV2VlayB2cyBXZWVrJyB9LAogICAgeyBrZXk6ICdjdXN0b20nLCBsYWJlbDogJ0N1c3RvbSBSYW5nZScgfSwKICAgIHsga2V5OiAnbW9udGgnLCBsYWJlbDogJ01vbnRobHknIH0sCiAgICB7IGtleTogJ3F1YXJ0ZXInLCBsYWJlbDogJ1F1YXJ0ZXJseScgfSwKICAgIHsga2V5OiAneXRkJywgbGFiZWw6ICdZZWFyIHRvIERhdGUnIH0sCiAgXTsKICBjb25zdCBNRVRSSUNfUk9XUyA9IFsKICAgIHsga2V5OiAndmlld3MnLCBsYWJlbDogJ1ZpZXdzJyB9LAogICAgeyBrZXk6ICdyZWFjaCcsIGxhYmVsOiAnUmVhY2gnIH0sCiAgICB7IGtleTogJ2ltcHJlc3Npb25zJywgbGFiZWw6ICdJbXByZXNzaW9ucycgfSwKICAgIHsga2V5OiAnZW5nYWdlbWVudCcsIGxhYmVsOiAnRW5nYWdlbWVudCcgfSwKICAgIHsga2V5OiAnY2xpY2tzJywgbGFiZWw6ICdDbGlja3MnIH0sCiAgICB7IGtleTogJ2ZvbGxvd2Vyc19nYWluZWQnLCBsYWJlbDogJ0ZvbGxvd2VycyBHYWluZWQnIH0sCiAgICB7IGtleTogJ3dhdGNoX3RpbWVfc2Vjb25kcycsIGxhYmVsOiAnV2F0Y2ggVGltZScgfSwKICAgIHsga2V5OiAnc2hhcmVzJywgbGFiZWw6ICdTaGFyZXMnIH0sCiAgICB7IGtleTogJ2NvbW1lbnRzJywgbGFiZWw6ICdDb21tZW50cycgfSwKICAgIHsga2V5OiAnc2F2ZXMnLCBsYWJlbDogJ1NhdmVzJyB9LAogIF07CgogIGxldCBtb2RlID0gJ3BsYXRmb3Jtcyc7CiAgbGV0IHJvb3Q7CiAgbGV0IHBsYXRmb3JtQ2hhcnRNZXRyaWMgPSAnZW5nYWdlbWVudCc7CiAgbGV0IGNhcmRTb3J0TW9kZSA9ICdvdmVyYWxsJzsKICBsZXQgY2FyZFBsYXRmb3JtRmlsdGVyID0gJ2FsbCc7CiAgLy8gQ29tcGFyaXNvbnMga2VlcHMgaXRzIG93biBQbGF0Zm9ybSBmaWx0ZXIsIGluZGVwZW5kZW50IG9mIHRoZSBEYXNoYm9hcmQncyDigJQgdGhlCiAgLy8gc2hhcmVkIGZpbHRlciBiYXIgaXMgaGlkZGVuIG9uIHRoaXMgdGFiIChzZWUgc3dpdGNoVGFiKSwgYnV0IGl0cyBzdGF0ZSBwZXJzaXN0cwogIC8vIGluIG1lbW9yeSwgc28gd2l0aG91dCB0aGlzIGV2ZXJ5IG1vZGUgaGVyZSB3b3VsZCBzaWxlbnRseSBrZWVwIHdoYXRldmVyIHBsYXRmb3JtCiAgLy8gd2FzIGxhc3QgcGlja2VkIG9uIERhc2hib2FyZCB3aXRoIG5vIHZpc2libGUgY29udHJvbCB0byBzZWUgb3IgY2hhbmdlIGl0LgogIGxldCBjb21wYXJpc29uUGxhdGZvcm0gPSAnYWxsJzsKCiAgZnVuY3Rpb24gbW9uZGF5T2YoZGF0ZVN0cikgewogICAgY29uc3QgZCA9IG5ldyBEYXRlKGRhdGVTdHIpOwogICAgY29uc3QgZGF5ID0gZC5nZXREYXkoKTsKICAgIGNvbnN0IGRpZmYgPSBkYXkgPT09IDAgPyA2IDogZGF5IC0gMTsKICAgIGQuc2V0RGF0ZShkLmdldERhdGUoKSAtIGRpZmYpOwogICAgcmV0dXJuIGQudG9JU09TdHJpbmcoKS5zbGljZSgwLCAxMCk7CiAgfQogIGZ1bmN0aW9uIGFkZERheXMoZGF0ZVN0ciwgbikgewogICAgY29uc3QgZCA9IG5ldyBEYXRlKGRhdGVTdHIpOwogICAgZC5zZXREYXRlKGQuZ2V0RGF0ZSgpICsgbik7CiAgICByZXR1cm4gZC50b0lTT1N0cmluZygpLnNsaWNlKDAsIDEwKTsKICB9CgogIGZ1bmN0aW9uIHNoZWxsKCkgewogICAgcm9vdC5pbm5lckhUTUwgPSAnJzsKCiAgICBjb25zdCB0YWJzID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0YWJzLmNsYXNzTmFtZSA9ICdtb2RlLXRhYnMnOwogICAgTU9ERVMuZm9yRWFjaCgobSkgPT4gewogICAgICBjb25zdCBidG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgICAgYnRuLnRleHRDb250ZW50ID0gbS5sYWJlbDsKICAgICAgYnRuLnR5cGUgPSAnYnV0dG9uJzsKICAgICAgaWYgKG0ua2V5ID09PSBtb2RlKSBidG4uY2xhc3NMaXN0LmFkZCgnaXMtYWN0aXZlJyk7CiAgICAgIGJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsKICAgICAgICBtb2RlID0gbS5rZXk7CiAgICAgICAgc2hlbGwoKTsKICAgICAgfSk7CiAgICAgIHRhYnMuYXBwZW5kQ2hpbGQoYnRuKTsKICAgIH0pOwogICAgcm9vdC5hcHBlbmRDaGlsZCh0YWJzKTsKCiAgICBjb25zdCBjb250cm9scyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY29udHJvbHMuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgY29udHJvbHMuaWQgPSAnY29tcGFyaXNvbkNvbnRyb2xzJzsKICAgIHJvb3QuYXBwZW5kQ2hpbGQoY29udHJvbHMpOwoKICAgIGNvbnN0IHJlc3VsdHMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHJlc3VsdHMuaWQgPSAnY29tcGFyaXNvblJlc3VsdHMnOwogICAgcm9vdC5hcHBlbmRDaGlsZChyZXN1bHRzKTsKCiAgICByZW5kZXJDb250cm9scygpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyQ29udHJvbHMoKSB7CiAgICBjb25zdCBjb250cm9scyA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdjb21wYXJpc29uQ29udHJvbHMnKTsKICAgIGNvbnRyb2xzLmlubmVySFRNTCA9ICcnOwogICAgY29uc3Qgcm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICByb3cuY2xhc3NOYW1lID0gJ2J0bi1yb3cnOwogICAgcm93LnN0eWxlLmFsaWduSXRlbXMgPSAnZW5kJzsKCiAgICBjb25zdCB0b2RheSA9IG5ldyBEYXRlKCkudG9JU09TdHJpbmcoKS5zbGljZSgwLCAxMCk7CiAgICBjb25zdCB0aGlzWWVhciA9IG5ldyBEYXRlKCkuZ2V0RnVsbFllYXIoKTsKCiAgICBpZiAobW9kZSA9PT0gJ3BsYXRmb3JtcycpIHsKICAgICAgY29uc3QgZkZyb20gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOyBmRnJvbS50eXBlID0gJ2RhdGUnOyBmRnJvbS5pZCA9ICdwbGF0Zm9ybVJlcG9ydEZyb20nOwogICAgICBjb25zdCBmVG8gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOyBmVG8udHlwZSA9ICdkYXRlJzsgZlRvLmlkID0gJ3BsYXRmb3JtUmVwb3J0VG8nOwogICAgICBjb25zdCBhcHBseUJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgICBhcHBseUJ0bi5jbGFzc05hbWUgPSAnYnRuIHByaW1hcnknOwogICAgICBhcHBseUJ0bi50eXBlID0gJ2J1dHRvbic7CiAgICAgIGFwcGx5QnRuLnRleHRDb250ZW50ID0gJ0FwcGx5IFJhbmdlJzsKICAgICAgYXBwbHlCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBsb2FkUGxhdGZvcm1SZXBvcnQoeyBkYXRlRnJvbTogZkZyb20udmFsdWUsIGRhdGVUbzogZlRvLnZhbHVlIH0pKTsKICAgICAgY29uc3QgY2xlYXJCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgICAgY2xlYXJCdG4uY2xhc3NOYW1lID0gJ2J0bic7CiAgICAgIGNsZWFyQnRuLnR5cGUgPSAnYnV0dG9uJzsKICAgICAgY2xlYXJCdG4udGV4dENvbnRlbnQgPSAnQWxsIFRpbWUnOwogICAgICBjbGVhckJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsgZkZyb20udmFsdWUgPSAnJzsgZlRvLnZhbHVlID0gJyc7IGxvYWRQbGF0Zm9ybVJlcG9ydCh7fSk7IH0pOwogICAgICByb3cuYXBwZW5kKAogICAgICAgIGxhYmVsZWQoJ0Zyb20gKG9wdGlvbmFsKScsIGZGcm9tKSwKICAgICAgICBsYWJlbGVkKCdUbyAob3B0aW9uYWwpJywgZlRvKSwKICAgICAgICBhcHBseUJ0biwKICAgICAgICBjbGVhckJ0bgogICAgICApOwogICAgICBjb250cm9scy5hcHBlbmRDaGlsZChyb3cpOwogICAgICBsb2FkUGxhdGZvcm1SZXBvcnQoe30pOwogICAgICByZXR1cm47CiAgICB9IGVsc2UgaWYgKG1vZGUgPT09ICd3ZWVrJykgewogICAgICAvLyBXZWVrIEEgaXMgYWx3YXlzIHRoZSBQcmV2aW91cyBwZXJpb2QsIFdlZWsgQiBpcyBhbHdheXMgdGhlIEN1cnJlbnQvbW9zdCByZWNlbnQKICAgICAgLy8gcGVyaW9kIOKAlCBydW5Db21wYXJlKCkncyBmaXJzdCBhcmd1bWVudCBpcyB0aGUgImN1cnJlbnQiIHNsb3QgZXZlcnkgb3RoZXIgbW9kZQogICAgICAvLyBpbiB0aGlzIHRhYiBmZWVkcyBpdCAocGVyY2VudENoYW5nZSA9IChjdXJyZW50IC0gcHJldmlvdXMpIC8gcHJldmlvdXMpLCBzbyBCIGdvZXMKICAgICAgLy8gaW4gZmlyc3QgYW5kIEEgc2Vjb25kLCByZWdhcmRsZXNzIG9mIHdoaWNoIGNhbGVuZGFyIHdlZWsgaXMgZWFybGllciBvciBsYXRlci4KICAgICAgY29uc3Qgd0EgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOyB3QS50eXBlID0gJ2RhdGUnOyB3QS52YWx1ZSA9IG1vbmRheU9mKGFkZERheXModG9kYXksIC03KSk7CiAgICAgIGNvbnN0IHdCID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsgd0IudHlwZSA9ICdkYXRlJzsgd0IudmFsdWUgPSBtb25kYXlPZih0b2RheSk7CiAgICAgIHJvdy5hcHBlbmQobGFiZWxlZCgnV2VlayBBIChQcmV2aW91cyknLCB3QSksIGxhYmVsZWQoJ1dlZWsgQiAoQ3VycmVudCknLCB3QiksIHBsYXRmb3JtRmlsdGVyRmllbGQoKSwgcnVuQnRuKCgpID0+IHsKICAgICAgICBjb25zdCByYW5nZUEgPSB7IGZyb206IG1vbmRheU9mKHdBLnZhbHVlKSwgdG86IGFkZERheXMobW9uZGF5T2Yod0EudmFsdWUpLCA2KSB9OwogICAgICAgIGNvbnN0IHJhbmdlQiA9IHsgZnJvbTogbW9uZGF5T2Yod0IudmFsdWUpLCB0bzogYWRkRGF5cyhtb25kYXlPZih3Qi52YWx1ZSksIDYpIH07CiAgICAgICAgcnVuQ29tcGFyZShyYW5nZUIsIHJhbmdlQSwgYFdlZWsgb2YgJHtGb3JtYXQuZGF0ZShyYW5nZUIuZnJvbSl9YCwgYFdlZWsgb2YgJHtGb3JtYXQuZGF0ZShyYW5nZUEuZnJvbSl9YCk7CiAgICAgIH0pKTsKICAgIH0gZWxzZSBpZiAobW9kZSA9PT0gJ2N1c3RvbScpIHsKICAgICAgY29uc3QgZkEgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOyBmQS50eXBlID0gJ2RhdGUnOyBmQS52YWx1ZSA9IGFkZERheXModG9kYXksIC0xMyk7CiAgICAgIGNvbnN0IHRBID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsgdEEudHlwZSA9ICdkYXRlJzsgdEEudmFsdWUgPSB0b2RheTsKICAgICAgY29uc3QgZkIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOyBmQi50eXBlID0gJ2RhdGUnOyBmQi52YWx1ZSA9IGFkZERheXModG9kYXksIC0yNyk7CiAgICAgIGNvbnN0IHRCID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsgdEIudHlwZSA9ICdkYXRlJzsgdEIudmFsdWUgPSBhZGREYXlzKHRvZGF5LCAtMTQpOwogICAgICByb3cuYXBwZW5kKAogICAgICAgIGxhYmVsZWQoJ1JhbmdlIEEgZnJvbScsIGZBKSwgbGFiZWxlZCgndG8nLCB0QSksCiAgICAgICAgbGFiZWxlZCgnUmFuZ2UgQiBmcm9tJywgZkIpLCBsYWJlbGVkKCd0bycsIHRCKSwKICAgICAgICBwbGF0Zm9ybUZpbHRlckZpZWxkKCksCiAgICAgICAgcnVuQnRuKCgpID0+IHJ1bkNvbXBhcmUoeyBmcm9tOiBmQS52YWx1ZSwgdG86IHRBLnZhbHVlIH0sIHsgZnJvbTogZkIudmFsdWUsIHRvOiB0Qi52YWx1ZSB9LCAnUmFuZ2UgQScsICdSYW5nZSBCJykpCiAgICAgICk7CiAgICB9IGVsc2UgaWYgKG1vZGUgPT09ICdtb250aCcpIHsKICAgICAgY29uc3QgeSA9IHllYXJTZWxlY3QodGhpc1llYXIpOyBjb25zdCBtID0gbW9udGhTZWxlY3QobmV3IERhdGUoKS5nZXRNb250aCgpICsgMSk7CiAgICAgIGNvbnN0IHRvZ2dsZSA9IHBlcmlvZFRvZ2dsZSgpOwogICAgICByb3cuYXBwZW5kKGxhYmVsZWQoJ1llYXInLCB5KSwgbGFiZWxlZCgnTW9udGgnLCBtKSwgdG9nZ2xlLmVsLCBwbGF0Zm9ybUZpbHRlckZpZWxkKCksIHJ1bkJ0bihhc3luYyAoKSA9PiB7CiAgICAgICAgY29uc3QgcmVwb3J0ID0gYXdhaXQgQXBpLm1vbnRobHkoeyB5ZWFyOiB5LnZhbHVlLCBtb250aDogbS52YWx1ZSwgLi4uU3RhdGUuZ2V0RmlsdGVycygpLCBwbGF0Zm9ybTogY29tcGFyaXNvblBsYXRmb3JtIH0pOwogICAgICAgIHJlbmRlclBlcmlvZFJlcG9ydChyZXBvcnQsIHRvZ2dsZS5nZXQoKSk7CiAgICAgIH0pKTsKICAgIH0gZWxzZSBpZiAobW9kZSA9PT0gJ3F1YXJ0ZXInKSB7CiAgICAgIGNvbnN0IHkgPSB5ZWFyU2VsZWN0KHRoaXNZZWFyKTsgY29uc3QgcSA9IHF1YXJ0ZXJTZWxlY3QoKTsKICAgICAgY29uc3QgdG9nZ2xlID0gcGVyaW9kVG9nZ2xlKCk7CiAgICAgIHJvdy5hcHBlbmQobGFiZWxlZCgnWWVhcicsIHkpLCBsYWJlbGVkKCdRdWFydGVyJywgcSksIHRvZ2dsZS5lbCwgcGxhdGZvcm1GaWx0ZXJGaWVsZCgpLCBydW5CdG4oYXN5bmMgKCkgPT4gewogICAgICAgIGNvbnN0IHJlcG9ydCA9IGF3YWl0IEFwaS5xdWFydGVybHkoeyB5ZWFyOiB5LnZhbHVlLCBxdWFydGVyOiBxLnZhbHVlLCAuLi5TdGF0ZS5nZXRGaWx0ZXJzKCksIHBsYXRmb3JtOiBjb21wYXJpc29uUGxhdGZvcm0gfSk7CiAgICAgICAgcmVuZGVyUGVyaW9kUmVwb3J0KHJlcG9ydCwgdG9nZ2xlLmdldCgpKTsKICAgICAgfSkpOwogICAgfSBlbHNlIGlmIChtb2RlID09PSAneXRkJykgewogICAgICBjb25zdCB5ID0geWVhclNlbGVjdCh0aGlzWWVhcik7CiAgICAgIHJvdy5hcHBlbmQobGFiZWxlZCgnWWVhcicsIHkpLCBwbGF0Zm9ybUZpbHRlckZpZWxkKCksIHJ1bkJ0bihhc3luYyAoKSA9PiB7CiAgICAgICAgY29uc3QgcmVwb3J0ID0gYXdhaXQgQXBpLnl0ZCh7IHllYXI6IHkudmFsdWUsIC4uLlN0YXRlLmdldEZpbHRlcnMoKSwgcGxhdGZvcm06IGNvbXBhcmlzb25QbGF0Zm9ybSB9KTsKICAgICAgICByZW5kZXJQZXJpb2RSZXBvcnQocmVwb3J0LCAndnNMYXN0WWVhcicpOwogICAgICB9KSk7CiAgICB9CgogICAgY29udHJvbHMuYXBwZW5kQ2hpbGQocm93KTsKICB9CgogIC8qKiBUaGUgQ29tcGFyaXNvbnMgdGFiJ3Mgb3duIFBsYXRmb3JtIGZpbHRlciDigJQgcmV1c2VzIHRoZSBzYW1lIHBsYXRmb3JtIGxpc3QgdGhlCiAgICAgIERhc2hib2FyZCdzIGZpbHRlciBiYXIgc2hvd3MsIGJ1dCB3cml0ZXMgdG8gY29tcGFyaXNvblBsYXRmb3JtLCBub3QgU3RhdGUsIHNvCiAgICAgIHRoZSB0d28gc3RheSBmdWxseSBpbmRlcGVuZGVudC4gKi8KICBmdW5jdGlvbiBwbGF0Zm9ybUZpbHRlckZpZWxkKCkgewogICAgY29uc3Qgc2VsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc2VsZWN0Jyk7CiAgICBjb25zdCBhbGxPcHQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsKICAgIGFsbE9wdC52YWx1ZSA9ICdhbGwnOwogICAgYWxsT3B0LnRleHRDb250ZW50ID0gJ0FsbCBwbGF0Zm9ybXMnOwogICAgc2VsLmFwcGVuZENoaWxkKGFsbE9wdCk7CiAgICBjb25zdCBvcHRpb25zID0gKHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZSB8fCB7IHBsYXRmb3JtczogW10gfSkucGxhdGZvcm1zIHx8IFtdOwogICAgb3B0aW9ucy5mb3JFYWNoKChwKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOwogICAgICBvcHQudmFsdWUgPSBwLmlkOwogICAgICBvcHQudGV4dENvbnRlbnQgPSBwLmxhYmVsOwogICAgICBzZWwuYXBwZW5kQ2hpbGQob3B0KTsKICAgIH0pOwogICAgc2VsLnZhbHVlID0gY29tcGFyaXNvblBsYXRmb3JtOwogICAgc2VsLmFkZEV2ZW50TGlzdGVuZXIoJ2NoYW5nZScsICgpID0+IHsgY29tcGFyaXNvblBsYXRmb3JtID0gc2VsLnZhbHVlOyB9KTsKICAgIHJldHVybiBsYWJlbGVkKCdQbGF0Zm9ybScsIHNlbCk7CiAgfQoKICBmdW5jdGlvbiBsYWJlbGVkKGxhYmVsLCBlbCkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgd3JhcC5jbGFzc05hbWUgPSAnZmllbGQtaW5saW5lJzsKICAgIHdyYXAuYXBwZW5kKHRleHRFbCgnbGFiZWwnLCBsYWJlbCksIGVsKTsKICAgIHJldHVybiB3cmFwOwogIH0KICBmdW5jdGlvbiBydW5CdG4ob25DbGljaykgewogICAgY29uc3QgYnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBidG4uY2xhc3NOYW1lID0gJ2J0biBwcmltYXJ5JzsKICAgIGJ0bi50ZXh0Q29udGVudCA9ICdDb21wYXJlJzsKICAgIGJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IG9uQ2xpY2soKSk7CiAgICByZXR1cm4gYnRuOwogIH0KICBmdW5jdGlvbiB5ZWFyU2VsZWN0KGRlZmF1bHRZZWFyKSB7CiAgICBjb25zdCBzZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzZWxlY3QnKTsKICAgIGZvciAobGV0IHkgPSBkZWZhdWx0WWVhciAtIDM7IHkgPD0gZGVmYXVsdFllYXIgKyAxOyB5ICs9IDEpIHsKICAgICAgY29uc3Qgb3B0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7IG9wdC52YWx1ZSA9IHk7IG9wdC50ZXh0Q29udGVudCA9IHk7CiAgICAgIGlmICh5ID09PSBkZWZhdWx0WWVhcikgb3B0LnNlbGVjdGVkID0gdHJ1ZTsKICAgICAgc2VsLmFwcGVuZENoaWxkKG9wdCk7CiAgICB9CiAgICByZXR1cm4gc2VsOwogIH0KICBmdW5jdGlvbiBtb250aFNlbGVjdChkZWZhdWx0TW9udGgpIHsKICAgIGNvbnN0IG5hbWVzID0gWydKYW51YXJ5JywnRmVicnVhcnknLCdNYXJjaCcsJ0FwcmlsJywnTWF5JywnSnVuZScsJ0p1bHknLCdBdWd1c3QnLCdTZXB0ZW1iZXInLCdPY3RvYmVyJywnTm92ZW1iZXInLCdEZWNlbWJlciddOwogICAgY29uc3Qgc2VsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc2VsZWN0Jyk7CiAgICBuYW1lcy5mb3JFYWNoKChuLCBpKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOyBvcHQudmFsdWUgPSBpICsgMTsgb3B0LnRleHRDb250ZW50ID0gbjsKICAgICAgaWYgKGkgKyAxID09PSBkZWZhdWx0TW9udGgpIG9wdC5zZWxlY3RlZCA9IHRydWU7CiAgICAgIHNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgICByZXR1cm4gc2VsOwogIH0KICBmdW5jdGlvbiBxdWFydGVyU2VsZWN0KCkgewogICAgY29uc3QgY3VycmVudFEgPSBNYXRoLmZsb29yKG5ldyBEYXRlKCkuZ2V0TW9udGgoKSAvIDMpICsgMTsKICAgIGNvbnN0IHNlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NlbGVjdCcpOwogICAgWzEsIDIsIDMsIDRdLmZvckVhY2goKHEpID0+IHsKICAgICAgY29uc3Qgb3B0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7IG9wdC52YWx1ZSA9IHE7IG9wdC50ZXh0Q29udGVudCA9IGBRJHtxfWA7CiAgICAgIGlmIChxID09PSBjdXJyZW50USkgb3B0LnNlbGVjdGVkID0gdHJ1ZTsKICAgICAgc2VsLmFwcGVuZENoaWxkKG9wdCk7CiAgICB9KTsKICAgIHJldHVybiBzZWw7CiAgfQogIGZ1bmN0aW9uIHBlcmlvZFRvZ2dsZSgpIHsKICAgIGNvbnN0IHNlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NlbGVjdCcpOwogICAgW1sndnNQcmV2aW91c1BlcmlvZCcsICd2cyBQcmV2aW91cyBQZXJpb2QnXSwgWyd2c0xhc3RZZWFyJywgJ3ZzIFNhbWUgUGVyaW9kIExhc3QgWWVhciddXS5mb3JFYWNoKChbdiwgbF0pID0+IHsKICAgICAgY29uc3Qgb3B0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7IG9wdC52YWx1ZSA9IHY7IG9wdC50ZXh0Q29udGVudCA9IGw7CiAgICAgIHNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgICByZXR1cm4geyBlbDogbGFiZWxlZCgnQ29tcGFyZScsIHNlbCksIGdldDogKCkgPT4gc2VsLnZhbHVlIH07CiAgfQoKICBhc3luYyBmdW5jdGlvbiBydW5Db21wYXJlKHJhbmdlQSwgcmFuZ2VCLCBsYWJlbEEsIGxhYmVsQikgewogICAgY29uc3QgZmlsdGVycyA9IFN0YXRlLmdldEZpbHRlcnMoKTsKICAgIGNvbnN0IHJlc3VsdCA9IGF3YWl0IEFwaS5jb21wYXJlKHsKICAgICAgZnJvbUE6IHJhbmdlQS5mcm9tLCB0b0E6IHJhbmdlQS50bywgZnJvbUI6IHJhbmdlQi5mcm9tLCB0b0I6IHJhbmdlQi50bywKICAgICAgcGxhdGZvcm06IGNvbXBhcmlzb25QbGF0Zm9ybSwgY2FtcGFpZ25UeXBlOiBmaWx0ZXJzLmNhbXBhaWduVHlwZSwgY29udGVudFR5cGU6IGZpbHRlcnMuY29udGVudFR5cGUsCiAgICB9KTsKICAgIHJlbmRlckNvbXBhcmVSZXN1bHQocmVzdWx0LCBsYWJlbEEsIGxhYmVsQik7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJQZXJpb2RSZXBvcnQocmVwb3J0LCB3aGljaCkgewogICAgY29uc3QgY21wID0gcmVwb3J0W3doaWNoXTsKICAgIGNvbnN0IGxhYmVsQSA9ICdDdXJyZW50IHBlcmlvZCc7CiAgICBjb25zdCBsYWJlbEIgPSB3aGljaCA9PT0gJ3ZzTGFzdFllYXInID8gJ1NhbWUgcGVyaW9kIGxhc3QgeWVhcicgOiAnUHJldmlvdXMgcGVyaW9kJzsKICAgIHJlbmRlckNvbXBhcmVSZXN1bHQoY21wLCBsYWJlbEEsIGxhYmVsQiwgcmVwb3J0LnJhbmdlKTsKICB9CgogIGZ1bmN0aW9uIHN0YXRUaWxlKGxhYmVsLCBjdXJyZW50LCBwcmV2aW91cywgZ3Jvd3RoLCBpc0R1cmF0aW9uLCBhcHBsaWNhYmxlVGhpcyA9IHRydWUsIGFwcGxpY2FibGVMYXN0ID0gdHJ1ZSkgewogICAgY29uc3QgdGlsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGlsZS5jbGFzc05hbWUgPSAnc3RhdC10aWxlJzsKICAgIGNvbnN0IG5hID0gJ04vQSc7CiAgICBjb25zdCBmbXRPbmUgPSAodikgPT4gKGlzRHVyYXRpb24gPyBGb3JtYXQuZHVyYXRpb24odikgOiBGb3JtYXQuY29tcGFjdCh2KSk7CiAgICBjb25zdCBjdXJEaXNwbGF5ID0gYXBwbGljYWJsZVRoaXMgPyBmbXRPbmUoY3VycmVudCkgOiBuYTsKICAgIGNvbnN0IHByZXZEaXNwbGF5ID0gYXBwbGljYWJsZUxhc3QgPyBmbXRPbmUocHJldmlvdXMpIDogbmE7CiAgICBjb25zdCBjb21wYXJhYmxlID0gYXBwbGljYWJsZVRoaXMgJiYgYXBwbGljYWJsZUxhc3Q7CiAgICBsZXQgZGVsdGFUZXh0OyBsZXQgZGVsdGFDbHM7CiAgICBpZiAoY29tcGFyYWJsZSkgewogICAgICBkZWx0YVRleHQgPSBgJHtGb3JtYXQucGN0KGdyb3d0aCl9IMK3IHdhcyAke3ByZXZEaXNwbGF5fWA7CiAgICAgIGRlbHRhQ2xzID0gRm9ybWF0LmRlbHRhQ2xhc3MoZ3Jvd3RoKTsKICAgIH0gZWxzZSBpZiAoIWFwcGxpY2FibGVUaGlzICYmICFhcHBsaWNhYmxlTGFzdCkgewogICAgICBkZWx0YVRleHQgPSAnbm90IHRyYWNrZWQgZm9yIHRoZXNlIGFjY291bnRzJzsKICAgICAgZGVsdGFDbHMgPSAnZmxhdCc7CiAgICB9IGVsc2UgewogICAgICBkZWx0YVRleHQgPSBgd2FzICR7cHJldkRpc3BsYXl9YDsKICAgICAgZGVsdGFDbHMgPSAnZmxhdCc7CiAgICB9CiAgICB0aWxlLmFwcGVuZCgKICAgICAgdGV4dEVsKCdkaXYnLCBsYWJlbCwgJ3N0YXQtbGFiZWwnKSwKICAgICAgdGV4dEVsKCdkaXYnLCBjdXJEaXNwbGF5LCAnc3RhdC12YWx1ZScpLAogICAgICB0ZXh0RWwoJ2RpdicsIGRlbHRhVGV4dCwgYHN0YXQtZGVsdGEgJHtkZWx0YUNsc31gKQogICAgKTsKICAgIHJldHVybiB0aWxlOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyQ29tcGFyZVJlc3VsdChyZXN1bHQsIGxhYmVsQSwgbGFiZWxCLCBoZWFkbGluZSkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdjb21wYXJpc29uUmVzdWx0cycpOwogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKCiAgICBjb25zdCB0aXRsZSA9IHRleHRFbCgnZGl2JywgaGVhZGxpbmUKICAgICAgPyBgJHtGb3JtYXQuZGF0ZShyZXN1bHQucmFuZ2VBLmZyb20pfSDigJMgJHtGb3JtYXQuZGF0ZShyZXN1bHQucmFuZ2VBLnRvKX1gCiAgICAgIDogYCR7bGFiZWxBfTogJHtGb3JtYXQuZGF0ZShyZXN1bHQucmFuZ2VBLmZyb20pfSDigJMgJHtGb3JtYXQuZGF0ZShyZXN1bHQucmFuZ2VBLnRvKX0gIHZzICAke2xhYmVsQn06ICR7Rm9ybWF0LmRhdGUocmVzdWx0LnJhbmdlQi5mcm9tKX0g4oCTICR7Rm9ybWF0LmRhdGUocmVzdWx0LnJhbmdlQi50byl9YCwKICAgICAgJ3NlY3Rpb24tdGl0bGUnKTsKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGl0bGUpOwoKICAgIGNvbnN0IGdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGdyaWQuY2xhc3NOYW1lID0gJ3N0YXQtZ3JpZCc7CiAgICAvLyBgYXBwbGljYWJsZWAgbWFwcyAoZnJvbSB0aGUgYmFja2VuZCkgdGVsbCBhIHJlYWwgMCBhcGFydCBmcm9tIE4vQSBwZXIgcmFuZ2U7IGB8fCB7fWAKICAgIC8vIGZhbGxzIGJhY2sgdG8gImV2ZXJ5dGhpbmcgYXBwbGljYWJsZSIgYWdhaW5zdCBhbiBvbGRlciBBUEkgd2l0aG91dCB0aGUgZmxhZy4KICAgIGNvbnN0IGFwQSA9IHJlc3VsdC5yYW5nZUEudG90YWxzLmFwcGxpY2FibGUgfHwgbnVsbDsKICAgIGNvbnN0IGFwQiA9IHJlc3VsdC5yYW5nZUIudG90YWxzLmFwcGxpY2FibGUgfHwgbnVsbDsKICAgIGNvbnN0IG9rQSA9IChrKSA9PiAoYXBBID8gISFhcEFba10gOiB0cnVlKTsKICAgIGNvbnN0IG9rQiA9IChrKSA9PiAoYXBCID8gISFhcEJba10gOiB0cnVlKTsKICAgIGdyaWQuYXBwZW5kQ2hpbGQoc3RhdFRpbGUoJ1Bvc3RzJywgcmVzdWx0LnJhbmdlQS50b3RhbHMucG9zdF9jb3VudCwgcmVzdWx0LnJhbmdlQi50b3RhbHMucG9zdF9jb3VudCwgcmVzdWx0Lmdyb3d0aC5wb3N0X2NvdW50LCBmYWxzZSkpOwogICAgTUVUUklDX1JPV1MuZm9yRWFjaCgobSkgPT4gewogICAgICBncmlkLmFwcGVuZENoaWxkKHN0YXRUaWxlKG0ubGFiZWwsIHJlc3VsdC5yYW5nZUEudG90YWxzW20ua2V5XSwgcmVzdWx0LnJhbmdlQi50b3RhbHNbbS5rZXldLCByZXN1bHQuZ3Jvd3RoW20ua2V5XSwgbS5rZXkgPT09ICd3YXRjaF90aW1lX3NlY29uZHMnLCBva0EobS5rZXkpLCBva0IobS5rZXkpKSk7CiAgICB9KTsKICAgIHdyYXAuYXBwZW5kQ2hpbGQoZ3JpZCk7CgogICAgcmVuZGVyUGxhdGZvcm1Db21wYXJpc29uQ2FyZHMod3JhcCwgcmVzdWx0LCBsYWJlbEEsIGxhYmVsQik7CiAgfQoKICAvKioKICAgKiAiQWxsIFBsYXRmb3JtcyIgcmVwb3J0IOKAlCB0aGUgaGVhZGxpbmUgQ29tcGFyaXNvbnMgdmlldy4gVW5saWtlIHRoZQogICAqIHdlZWsvY3VzdG9tL21vbnRoL3F1YXJ0ZXIveXRkIHRvb2xzIGFib3ZlLCB0aGlzIGlnbm9yZXMgdGhlIHNoYXJlZAogICAqIHBsYXRmb3JtL2NhbXBhaWduL2NvbnRlbnQtdHlwZSBmaWx0ZXIgYmFyIGVudGlyZWx5IGFuZCBuZWVkcyBubyBkYXRlCiAgICogcmFuZ2U6IGl0IGFsd2F5cyBjb3ZlcnMgZXZlcnkgcGxhdGZvcm0gd2l0aCBhbnkgZGF0YSAodXBsb2FkZWQgcG9zdHMKICAgKiBhbmQvb3IgbWFudWFsbHktZW50ZXJlZCBGb2xsb3dlcnMgRGF0YSBSZWNvcmQgaGlzdG9yeSkuCiAgICovCiAgYXN5bmMgZnVuY3Rpb24gbG9hZFBsYXRmb3JtUmVwb3J0KHBhcmFtcykgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdjb21wYXJpc29uUmVzdWx0cycpOwogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgIHdyYXAuYXBwZW5kQ2hpbGQoc2tlbGV0b25TdGF0R3JpZCgyKSk7CiAgICB3cmFwLmFwcGVuZENoaWxkKHNrZWxldG9uQ2hhcnQoKSk7CiAgICBjb25zdCBoYXNFeHBsaWNpdFJhbmdlID0gcGFyYW1zICYmIHBhcmFtcy5kYXRlRnJvbSAmJiBwYXJhbXMuZGF0ZVRvOwogICAgY29uc3QgcmVwb3J0ID0gYXdhaXQgQXBpLnBsYXRmb3JtUmVwb3J0KGhhc0V4cGxpY2l0UmFuZ2UgPyBwYXJhbXMgOiB7fSk7CiAgICByZW5kZXJQbGF0Zm9ybVJlcG9ydChyZXBvcnQpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyUGxhdGZvcm1SZXBvcnQocmVwb3J0KSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2NvbXBhcmlzb25SZXN1bHRzJyk7CiAgICB3cmFwLmlubmVySFRNTCA9ICcnOwoKICAgIGlmICghcmVwb3J0LnBsYXRmb3Jtcy5sZW5ndGgpIHsKICAgICAgd3JhcC5hcHBlbmRDaGlsZChlbXB0eVN0YXRlKHsKICAgICAgICBpY29uOiAnZ2l0LWNvbXBhcmUnLAogICAgICAgIHRpdGxlOiAnTm8gcGxhdGZvcm0gZGF0YSB5ZXQnLAogICAgICAgIG1lc3NhZ2U6ICdVcGxvYWQgcG9zdHMgb3IgYWRkIEZvbGxvd2VycyBEYXRhIFJlY29yZCBlbnRyaWVzIHRvIHNlZSBhIGNyb3NzLXBsYXRmb3JtIGNvbXBhcmlzb24gaGVyZS4nLAogICAgICB9KSk7CiAgICAgIHJldHVybjsKICAgIH0KCiAgICBjb25zdCByYW5nZUxhYmVsID0gcmVwb3J0LnJhbmdlLmlzRXhwbGljaXQKICAgICAgPyBgJHtGb3JtYXQuZGF0ZShyZXBvcnQucmFuZ2UuZnJvbSl9IOKAkyAke0Zvcm1hdC5kYXRlKHJlcG9ydC5yYW5nZS50byl9YAogICAgICA6IGBBbGwgdGltZSAoJHtGb3JtYXQuZGF0ZShyZXBvcnQucmFuZ2UuZnJvbSl9IOKAkyAke0Zvcm1hdC5kYXRlKHJlcG9ydC5yYW5nZS50byl9KWA7CiAgICB3cmFwLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgYFBsYXRmb3JtIENvbXBhcmlzb24gUmVwb3J0IOKAlCAke3JhbmdlTGFiZWx9YCwgJ3NlY3Rpb24tdGl0bGUnKSk7CgogICAgY29uc3QgYmVzdFAgPSByZXBvcnQucGxhdGZvcm1zLmZpbmQoKHApID0+IHAucGxhdGZvcm0gPT09IHJlcG9ydC5iZXN0UGxhdGZvcm0pOwogICAgY29uc3Qgd29yc3RQID0gcmVwb3J0LnBsYXRmb3Jtcy5maW5kKChwKSA9PiBwLnBsYXRmb3JtID09PSByZXBvcnQud29yc3RQbGF0Zm9ybSk7CiAgICBpZiAoYmVzdFAgfHwgd29yc3RQKSB7CiAgICAgIGNvbnN0IGhpZ2hsaWdodEdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgaGlnaGxpZ2h0R3JpZC5jbGFzc05hbWUgPSAnc3RhdC1ncmlkJzsKICAgICAgaWYgKGJlc3RQKSB7CiAgICAgICAgY29uc3QgdGlsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICAgIHRpbGUuY2xhc3NOYW1lID0gJ3N0YXQtdGlsZSc7CiAgICAgICAgdGlsZS5hcHBlbmQoCiAgICAgICAgICB0ZXh0RWwoJ2RpdicsICdCZXN0LVBlcmZvcm1pbmcgUGxhdGZvcm0nLCAnc3RhdC1sYWJlbCcpLAogICAgICAgICAgdGV4dEVsKCdkaXYnLCBiZXN0UC5sYWJlbCwgJ3N0YXQtdmFsdWUnKSwKICAgICAgICAgIHRleHRFbCgnZGl2JywgYFJlYWNoICR7Rm9ybWF0LnNtYXJ0KGJlc3RQLnRvdGFscy5yZWFjaCl9IMK3IEVuZ2FnZW1lbnQgJHtGb3JtYXQuc21hcnQoYmVzdFAudG90YWxzLmVuZ2FnZW1lbnQpfWAsICdwb3N0LW1ldGEnKQogICAgICAgICk7CiAgICAgICAgaGlnaGxpZ2h0R3JpZC5hcHBlbmRDaGlsZCh0aWxlKTsKICAgICAgfQogICAgICBpZiAod29yc3RQKSB7CiAgICAgICAgY29uc3QgdGlsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICAgIHRpbGUuY2xhc3NOYW1lID0gJ3N0YXQtdGlsZSc7CiAgICAgICAgdGlsZS5hcHBlbmQoCiAgICAgICAgICB0ZXh0RWwoJ2RpdicsICdMb3dlc3QtUGVyZm9ybWluZyBQbGF0Zm9ybScsICdzdGF0LWxhYmVsJyksCiAgICAgICAgICB0ZXh0RWwoJ2RpdicsIHdvcnN0UC5sYWJlbCwgJ3N0YXQtdmFsdWUnKSwKICAgICAgICAgIHRleHRFbCgnZGl2JywgYFJlYWNoICR7Rm9ybWF0LnNtYXJ0KHdvcnN0UC50b3RhbHMucmVhY2gpfSDCtyBFbmdhZ2VtZW50ICR7Rm9ybWF0LnNtYXJ0KHdvcnN0UC50b3RhbHMuZW5nYWdlbWVudCl9YCwgJ3Bvc3QtbWV0YScpCiAgICAgICAgKTsKICAgICAgICBoaWdobGlnaHRHcmlkLmFwcGVuZENoaWxkKHRpbGUpOwogICAgICB9CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQoaGlnaGxpZ2h0R3JpZCk7CiAgICB9CgogICAgY29uc3QgdGFibGVDYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0YWJsZUNhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgdGFibGVDYXJkLmFwcGVuZENoaWxkKHRleHRFbCgnaDMnLCAnUGxhdGZvcm0gUmFua2luZycpKTsKICAgIGNvbnN0IHRhYmxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGFibGUnKTsKICAgIHRhYmxlLmNsYXNzTmFtZSA9ICdkYXRhLXRhYmxlJzsKICAgIGNvbnN0IHRoZWFkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGhlYWQnKTsKICAgIGNvbnN0IGhlYWRSb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgWydSYW5rJywgJ1BsYXRmb3JtJywgJ1Bvc3RzJywgJ1JlYWNoJywgJ0VuZ2FnZW1lbnQnLCAnSW1wcmVzc2lvbnMnLCAnRm9sbG93ZXIgR3Jvd3RoJ10uZm9yRWFjaCgobGFiZWwsIGkpID0+IHsKICAgICAgY29uc3QgdGggPSB0ZXh0RWwoJ3RoJywgbGFiZWwpOwogICAgICBpZiAoaSA+PSAyKSB0aC5jbGFzc0xpc3QuYWRkKCdudW0nKTsKICAgICAgaGVhZFJvdy5hcHBlbmRDaGlsZCh0aCk7CiAgICB9KTsKICAgIHRoZWFkLmFwcGVuZENoaWxkKGhlYWRSb3cpOwogICAgdGFibGUuYXBwZW5kQ2hpbGQodGhlYWQpOwogICAgY29uc3QgdGJvZHkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0Ym9keScpOwogICAgcmVwb3J0LnBsYXRmb3Jtcy5mb3JFYWNoKChwKSA9PiB7CiAgICAgIGNvbnN0IHRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgdHIuYXBwZW5kQ2hpbGQodGV4dEVsKCd0ZCcsIHAub3ZlcmFsbFJhbmsgPyBgIyR7cC5vdmVyYWxsUmFua31gIDogJ+KAlCcpKTsKICAgICAgY29uc3QgcGxhdFRkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgY29uc3QgcGlsbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsgcGlsbC5jbGFzc05hbWUgPSAncGxhdGZvcm0tcGlsbCc7CiAgICAgIGNvbnN0IGRvdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsgZG90LmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1kb3QnOyBkb3Quc3R5bGUuYmFja2dyb3VuZCA9IHAuY29sb3I7CiAgICAgIHBpbGwuYXBwZW5kKGRvdCwgZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUocC5sYWJlbCkpOwogICAgICBwbGF0VGQuYXBwZW5kQ2hpbGQocGlsbCk7CiAgICAgIHRyLmFwcGVuZENoaWxkKHBsYXRUZCk7CiAgICAgIHRyLmFwcGVuZENoaWxkKHRleHRFbCgndGQnLCBGb3JtYXQubnVtYmVyKHAucG9zdENvdW50KSwgJ251bScpKTsKICAgICAgdHIuYXBwZW5kQ2hpbGQodGV4dEVsKCd0ZCcsIEZvcm1hdC5zbWFydChwLnRvdGFscy5yZWFjaCksICdudW0nKSk7CiAgICAgIHRyLmFwcGVuZENoaWxkKHRleHRFbCgndGQnLCBGb3JtYXQuc21hcnQocC50b3RhbHMuZW5nYWdlbWVudCksICdudW0nKSk7CiAgICAgIHRyLmFwcGVuZENoaWxkKHRleHRFbCgndGQnLCBGb3JtYXQuc21hcnQocC50b3RhbHMuaW1wcmVzc2lvbnMpLCAnbnVtJykpOwogICAgICBjb25zdCBmb2xsb3dlclRkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgZm9sbG93ZXJUZC5jbGFzc05hbWUgPSAnbnVtJzsKICAgICAgaWYgKHAuZm9sbG93ZXJzLmNoYW5nZSA9PT0gbnVsbCkgewogICAgICAgIGZvbGxvd2VyVGQuYXBwZW5kQ2hpbGQoZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUocC5mb2xsb3dlcnMubGF0ZXN0ICE9PSBudWxsID8gRm9ybWF0Lm51bWJlcihwLmZvbGxvd2Vycy5sYXRlc3QpIDogJ+KAlCcpKTsKICAgICAgfSBlbHNlIHsKICAgICAgICBjb25zdCBmb2xsb3dlcldyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICAgICAgZm9sbG93ZXJXcmFwLnN0eWxlLmRpc3BsYXkgPSAnaW5saW5lLWZsZXgnOwogICAgICAgIGZvbGxvd2VyV3JhcC5zdHlsZS5hbGlnbkl0ZW1zID0gJ2NlbnRlcic7CiAgICAgICAgZm9sbG93ZXJXcmFwLnN0eWxlLmdhcCA9ICc2cHgnOwogICAgICAgIGZvbGxvd2VyV3JhcC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3NwYW4nLCBgJHtwLmZvbGxvd2Vycy5jaGFuZ2UgPiAwID8gJysnIDogJyd9JHtGb3JtYXQubnVtYmVyKHAuZm9sbG93ZXJzLmNoYW5nZSl9YCwgYHN0YXQtZGVsdGEgJHtGb3JtYXQuZGVsdGFDbGFzcyhwLmZvbGxvd2Vycy5jaGFuZ2UpfWApKTsKICAgICAgICBpZiAocC5mb2xsb3dlcnMuY2hhbmdlUGN0ICE9PSBudWxsKSBmb2xsb3dlcldyYXAuYXBwZW5kQ2hpbGQodGV4dEVsKCdzcGFuJywgYCgke0Zvcm1hdC5wY3QocC5mb2xsb3dlcnMuY2hhbmdlUGN0KX0pYCwgJ3Bvc3QtbWV0YScpKTsKICAgICAgICBmb2xsb3dlclRkLmFwcGVuZENoaWxkKGZvbGxvd2VyV3JhcCk7CiAgICAgIH0KICAgICAgdHIuYXBwZW5kQ2hpbGQoZm9sbG93ZXJUZCk7CiAgICAgIHRib2R5LmFwcGVuZENoaWxkKHRyKTsKICAgIH0pOwogICAgdGFibGUuYXBwZW5kQ2hpbGQodGJvZHkpOwogICAgY29uc3QgdGFibGVTY3JvbGwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRhYmxlU2Nyb2xsLmNsYXNzTmFtZSA9ICd0YWJsZS1zY3JvbGwnOwogICAgdGFibGVTY3JvbGwuYXBwZW5kQ2hpbGQodGFibGUpOwogICAgdGFibGVDYXJkLmFwcGVuZENoaWxkKHRhYmxlU2Nyb2xsKTsKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGFibGVDYXJkKTsKCiAgICBjb25zdCBjaGFydENhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNoYXJ0Q2FyZC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICBjb25zdCBjaGFydEhlYWRlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY2hhcnRIZWFkZXIuY2xhc3NOYW1lID0gJ2NhcmQtaGVhZGVyJzsKICAgIGNoYXJ0SGVhZGVyLmFwcGVuZENoaWxkKHRleHRFbCgnaDMnLCAnTWV0cmljIENvbXBhcmlzb24nKSk7CiAgICBjb25zdCBtZXRyaWNTZWxlY3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzZWxlY3QnKTsKICAgIE1FVFJJQ19ST1dTLmZvckVhY2goKG0pID0+IHsKICAgICAgY29uc3Qgb3B0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7IG9wdC52YWx1ZSA9IG0ua2V5OyBvcHQudGV4dENvbnRlbnQgPSBtLmxhYmVsOwogICAgICBpZiAobS5rZXkgPT09IHBsYXRmb3JtQ2hhcnRNZXRyaWMpIG9wdC5zZWxlY3RlZCA9IHRydWU7CiAgICAgIG1ldHJpY1NlbGVjdC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgICBtZXRyaWNTZWxlY3QuYWRkRXZlbnRMaXN0ZW5lcignY2hhbmdlJywgKCkgPT4gewogICAgICBwbGF0Zm9ybUNoYXJ0TWV0cmljID0gbWV0cmljU2VsZWN0LnZhbHVlOwogICAgICBkcmF3UGxhdGZvcm1SZXBvcnRDaGFydChyZXBvcnQpOwogICAgfSk7CiAgICBjaGFydEhlYWRlci5hcHBlbmRDaGlsZChtZXRyaWNTZWxlY3QpOwogICAgY29uc3QgY2hhcnRXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjaGFydFdyYXAuY2xhc3NOYW1lID0gJ2NoYXJ0LXdyYXAgdGFsbCc7CiAgICBjaGFydFdyYXAuaW5uZXJIVE1MID0gJzxjYW52YXMgaWQ9InBsYXRmb3JtUmVwb3J0Q2FudmFzIj48L2NhbnZhcz4nOwogICAgY2hhcnRDYXJkLmFwcGVuZChjaGFydEhlYWRlciwgY2hhcnRXcmFwKTsKICAgIHdyYXAuYXBwZW5kQ2hpbGQoY2hhcnRDYXJkKTsKICAgIGRyYXdQbGF0Zm9ybVJlcG9ydENoYXJ0KHJlcG9ydCk7CgogICAgY29uc3Qgd2l0aEZvbGxvd2VycyA9IHJlcG9ydC5wbGF0Zm9ybXMuZmlsdGVyKChwKSA9PiBwLmZvbGxvd2Vycy5sYXRlc3QgIT09IG51bGwpOwogICAgaWYgKHdpdGhGb2xsb3dlcnMubGVuZ3RoKSB7CiAgICAgIGNvbnN0IGZvbGxvd2VyQ2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBmb2xsb3dlckNhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgICBmb2xsb3dlckNhcmQuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMycsICdGb2xsb3dlciBHcm93dGggYnkgUGxhdGZvcm0nKSk7CiAgICAgIGNvbnN0IGZDaGFydFdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgZkNoYXJ0V3JhcC5jbGFzc05hbWUgPSAnY2hhcnQtd3JhcCB0YWxsJzsKICAgICAgZkNoYXJ0V3JhcC5pbm5lckhUTUwgPSAnPGNhbnZhcyBpZD0icGxhdGZvcm1Gb2xsb3dlckNhbnZhcyI+PC9jYW52YXM+JzsKICAgICAgZm9sbG93ZXJDYXJkLmFwcGVuZENoaWxkKGZDaGFydFdyYXApOwogICAgICB3cmFwLmFwcGVuZENoaWxkKGZvbGxvd2VyQ2FyZCk7CiAgICAgIENoYXJ0cy5wbGF0Zm9ybUJhckNoYXJ0KCdwbGF0Zm9ybUZvbGxvd2VyQ2FudmFzJywgewogICAgICAgIGxhYmVsczogd2l0aEZvbGxvd2Vycy5tYXAoKHApID0+IHAubGFiZWwpLAogICAgICAgIGRhdGE6IHdpdGhGb2xsb3dlcnMubWFwKChwKSA9PiBwLmZvbGxvd2Vycy5sYXRlc3QgfHwgMCksCiAgICAgICAgY29sb3JzOiB3aXRoRm9sbG93ZXJzLm1hcCgocCkgPT4gcC5jb2xvciksCiAgICAgICAgZm9ybWF0VmFsdWU6ICh2KSA9PiBGb3JtYXQuc21hcnQodiksCiAgICAgIH0pOwogICAgfQoKICAgIGNvbnN0IGluc2lnaHRzQ2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgaW5zaWdodHNDYXJkLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGluc2lnaHRzQ2FyZC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2gzJywgJ0luc2lnaHRzICYgU3VtbWFyeScpKTsKICAgIGNvbnN0IGxpc3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd1bCcpOwogICAgbGlzdC5jbGFzc05hbWUgPSAnaW5zaWdodHMtbGlzdCc7CiAgICByZXBvcnQuaW5zaWdodHMuZm9yRWFjaCgobGluZSkgPT4gbGlzdC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2xpJywgbGluZSkpKTsKICAgIGluc2lnaHRzQ2FyZC5hcHBlbmRDaGlsZChsaXN0KTsKICAgIHdyYXAuYXBwZW5kQ2hpbGQoaW5zaWdodHNDYXJkKTsKICB9CgogIGZ1bmN0aW9uIGRyYXdQbGF0Zm9ybVJlcG9ydENoYXJ0KHJlcG9ydCkgewogICAgQ2hhcnRzLnBsYXRmb3JtQmFyQ2hhcnQoJ3BsYXRmb3JtUmVwb3J0Q2FudmFzJywgewogICAgICBsYWJlbHM6IHJlcG9ydC5wbGF0Zm9ybXMubWFwKChwKSA9PiBwLmxhYmVsKSwKICAgICAgZGF0YTogcmVwb3J0LnBsYXRmb3Jtcy5tYXAoKHApID0+IHAudG90YWxzW3BsYXRmb3JtQ2hhcnRNZXRyaWNdIHx8IDApLAogICAgICBjb2xvcnM6IHJlcG9ydC5wbGF0Zm9ybXMubWFwKChwKSA9PiBwLmNvbG9yKSwKICAgICAgZm9ybWF0VmFsdWU6ICh2KSA9PiAocGxhdGZvcm1DaGFydE1ldHJpYyA9PT0gJ3dhdGNoX3RpbWVfc2Vjb25kcycgPyBGb3JtYXQuZHVyYXRpb24odikgOiBGb3JtYXQuc21hcnQodikpLAogICAgfSk7CiAgfQoKICAvKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgICAgUGxhdGZvcm0gUGVyZm9ybWFuY2UgQ29tcGFyaXNvbiDigJQgcmVwbGFjZXMgdGhlIG9sZCBncm91cGVkCiAgICAgIlJhbmdlIEEgdnMgUmFuZ2UgQiBieSBQbGF0Zm9ybSIgY2hhcnQuIE9uZSBjYXJkIHBlciBwbGF0Zm9ybQogICAgIHdpdGggYW55IGRhdGEgaW4gZWl0aGVyIHJhbmdlLCBidWlsdCBlbnRpcmVseSBmcm9tIHRoZSBzYW1lCiAgICAgY29tcGFyZVJhbmdlcygpIHJlc3BvbnNlIHRoZSBzdGF0LXRpbGUgZ3JpZCBhYm92ZSBhbHJlYWR5CiAgICAgdXNlcyAocmVzdWx0LnJhbmdlQS5wbGF0Zm9ybXMgLyByZXN1bHQucmFuZ2VCLnBsYXRmb3Jtcykg4oCUIG5vCiAgICAgZXh0cmEgZmV0Y2guCiAgICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwogIGNvbnN0IEFMTF9DQVJEX01FVFJJQ1MgPSBbeyBrZXk6ICdwb3N0X2NvdW50JywgbGFiZWw6ICdQb3N0cycgfSwgLi4uTUVUUklDX1JPV1NdOwogIGNvbnN0IENBUkRfU09SVF9NT0RFUyA9IFsKICAgIHsga2V5OiAnb3ZlcmFsbCcsIGxhYmVsOiAnT3ZlcmFsbCBQZXJmb3JtYW5jZScgfSwKICAgIHsga2V5OiAnZ3Jvd3RoJywgbGFiZWw6ICdIaWdoZXN0IEdyb3d0aCcgfSwKICAgIHsga2V5OiAnZW5nYWdlbWVudCcsIGxhYmVsOiAnSGlnaGVzdCBFbmdhZ2VtZW50JyB9LAogICAgeyBrZXk6ICdmb2xsb3dlcnMnLCBsYWJlbDogJ01vc3QgRm9sbG93ZXJzJyB9LAogICAgeyBrZXk6ICdwb3N0cycsIGxhYmVsOiAnTW9zdCBQb3N0cycgfSwKICAgIHsga2V5OiAnYWxwaGEnLCBsYWJlbDogJ0FscGhhYmV0aWNhbCcgfSwKICBdOwoKICAvKiogUGVyLW1ldHJpYyB7YSwgYiwgZGlmZiwgcGN0RGlmZn0gYWNyb3NzIGJvdGggcmFuZ2VzIGZvciBvbmUgcGxhdGZvcm0sIHNraXBwaW5nIGFueSBtZXRyaWMgdGhhdCdzIHplcm8gaW4gYm90aCDigJQgYSBwbGF0Zm9ybSdzIGNhcmQgc2hvdWxkIG9ubHkgZXZlciBzaG93IG1ldHJpY3MgaXQgYWN0dWFsbHkgaGFzLiAqLwogIGZ1bmN0aW9uIGNvbXB1dGVDYXJkTWV0cmljcyhwbGF0Zm9ybUEsIHBsYXRmb3JtQikgewogICAgY29uc3QgbWV0cmljcyA9IFtdOwogICAgQUxMX0NBUkRfTUVUUklDUy5mb3JFYWNoKCh7IGtleSwgbGFiZWwgfSkgPT4gewogICAgICBjb25zdCBhID0gKHBsYXRmb3JtQSAmJiBwbGF0Zm9ybUFba2V5XSkgfHwgMDsKICAgICAgY29uc3QgYiA9IChwbGF0Zm9ybUIgJiYgcGxhdGZvcm1CW2tleV0pIHx8IDA7CiAgICAgIGlmIChhID09PSAwICYmIGIgPT09IDApIHJldHVybjsKICAgICAgY29uc3QgZGlmZiA9IGEgLSBiOwogICAgICBjb25zdCBwY3REaWZmID0gYiA/IE1hdGgucm91bmQoKGRpZmYgLyBiKSAqIDEwMDApIC8gMTAgOiAoYSA+IDAgPyBudWxsIDogMCk7CiAgICAgIG1ldHJpY3MucHVzaCh7IGtleSwgbGFiZWwsIGEsIGIsIGRpZmYsIHBjdERpZmYsIGlzRHVyYXRpb246IGtleSA9PT0gJ3dhdGNoX3RpbWVfc2Vjb25kcycgfSk7CiAgICB9KTsKICAgIHJldHVybiBtZXRyaWNzOwogIH0KCiAgLyoqIEEgc2luZ2xlICJob3cgZGlkIHRoaXMgcGxhdGZvcm0gZG8gb3ZlcmFsbCIgbnVtYmVyOiB0aGUgYXZlcmFnZSAlIGNoYW5nZSBhY3Jvc3MgZXZlcnkgbWV0cmljIHRoYXQgaGFzIGEgY29tcHV0YWJsZSBwZXJjZW50YWdlIChhIG1ldHJpYyBnb2luZyBmcm9tIDAgdG8gc29tZXRoaW5nIGhhcyBubyBwZXJjZW50YWdlIOKAlCAibmV3Iiwgbm90IGNvdW50ZWQgZWl0aGVyIHdheSkuICovCiAgZnVuY3Rpb24gb3ZlcmFsbFBjdENoYW5nZShtZXRyaWNzKSB7CiAgICBjb25zdCB3aXRoUGN0ID0gbWV0cmljcy5maWx0ZXIoKG0pID0+IG0ucGN0RGlmZiAhPT0gbnVsbCk7CiAgICBpZiAoIXdpdGhQY3QubGVuZ3RoKSByZXR1cm4gbnVsbDsKICAgIHJldHVybiBNYXRoLnJvdW5kKCh3aXRoUGN0LnJlZHVjZSgoc3VtLCBtKSA9PiBzdW0gKyBtLnBjdERpZmYsIDApIC8gd2l0aFBjdC5sZW5ndGgpICogMTApIC8gMTA7CiAgfQoKICBmdW5jdGlvbiBiZXN0V2Vha2VzdE1ldHJpYyhtZXRyaWNzKSB7CiAgICBjb25zdCB3aXRoUGN0ID0gbWV0cmljcy5maWx0ZXIoKG0pID0+IG0ucGN0RGlmZiAhPT0gbnVsbCk7CiAgICBpZiAoIXdpdGhQY3QubGVuZ3RoKSByZXR1cm4geyBiZXN0OiBudWxsLCB3ZWFrZXN0OiBudWxsIH07CiAgICBjb25zdCBiZXN0ID0gd2l0aFBjdC5yZWR1Y2UoKGEsIGIpID0+IChiLnBjdERpZmYgPiBhLnBjdERpZmYgPyBiIDogYSkpOwogICAgY29uc3Qgd2Vha2VzdCA9IHdpdGhQY3QucmVkdWNlKChhLCBiKSA9PiAoYi5wY3REaWZmIDwgYS5wY3REaWZmID8gYiA6IGEpKTsKICAgIHJldHVybiB7IGJlc3QsIHdlYWtlc3QgfTsKICB9CgogIGZ1bmN0aW9uIHRyZW5kRGlyZWN0aW9uKHBjdCkgewogICAgaWYgKHBjdCA9PT0gbnVsbCB8fCBwY3QgPT09IHVuZGVmaW5lZCkgcmV0dXJuICdmbGF0JzsKICAgIGlmIChwY3QgPiAwLjUpIHJldHVybiAndXAnOwogICAgaWYgKHBjdCA8IC0wLjUpIHJldHVybiAnZG93bic7CiAgICByZXR1cm4gJ2ZsYXQnOwogIH0KCiAgZnVuY3Rpb24gYnVpbGRQbGF0Zm9ybUNhcmRzKHJlc3VsdCkgewogICAgY29uc3QgcGxhdGZvcm1PcHRpb25zID0gKHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZSB8fCB7IGFsbFBsYXRmb3JtczogW10gfSkuYWxsUGxhdGZvcm1zOwogICAgY29uc3QgaWRzID0gWy4uLm5ldyBTZXQoWy4uLnJlc3VsdC5yYW5nZUEucGxhdGZvcm1zLCAuLi5yZXN1bHQucmFuZ2VCLnBsYXRmb3Jtc10ubWFwKChwKSA9PiBwLnBsYXRmb3JtKSldOwogICAgY29uc3QgYnlJZEEgPSBPYmplY3QuZnJvbUVudHJpZXMocmVzdWx0LnJhbmdlQS5wbGF0Zm9ybXMubWFwKChwKSA9PiBbcC5wbGF0Zm9ybSwgcF0pKTsKICAgIGNvbnN0IGJ5SWRCID0gT2JqZWN0LmZyb21FbnRyaWVzKHJlc3VsdC5yYW5nZUIucGxhdGZvcm1zLm1hcCgocCkgPT4gW3AucGxhdGZvcm0sIHBdKSk7CgogICAgcmV0dXJuIGlkcwogICAgICAubWFwKChpZCkgPT4gewogICAgICAgIGNvbnN0IG1ldGEgPSBwbGF0Zm9ybU9wdGlvbnMuZmluZCgocCkgPT4gcC5pZCA9PT0gaWQpIHx8IHsgaWQsIGxhYmVsOiBpZCwgY29sb3I6ICd2YXIoLS1zZXJpZXMtMSknIH07CiAgICAgICAgY29uc3QgYSA9IGJ5SWRBW2lkXSB8fCBudWxsOwogICAgICAgIGNvbnN0IGIgPSBieUlkQltpZF0gfHwgbnVsbDsKICAgICAgICBjb25zdCBtZXRyaWNzID0gY29tcHV0ZUNhcmRNZXRyaWNzKGEsIGIpOwogICAgICAgIGNvbnN0IHsgYmVzdCwgd2Vha2VzdCB9ID0gYmVzdFdlYWtlc3RNZXRyaWMobWV0cmljcyk7CiAgICAgICAgcmV0dXJuIHsKICAgICAgICAgIHBsYXRmb3JtOiBpZCwKICAgICAgICAgIGxhYmVsOiBtZXRhLmxhYmVsLAogICAgICAgICAgY29sb3I6IG1ldGEuY29sb3IsCiAgICAgICAgICBtZXRyaWNzLAogICAgICAgICAgb3ZlcmFsbDogb3ZlcmFsbFBjdENoYW5nZShtZXRyaWNzKSwKICAgICAgICAgIGJlc3QsCiAgICAgICAgICB3ZWFrZXN0LAogICAgICAgICAgZm9sbG93ZXJzR2FpbmVkOiAoYSA/IGEuZm9sbG93ZXJzX2dhaW5lZCB8fCAwIDogMCkgKyAoYiA/IGIuZm9sbG93ZXJzX2dhaW5lZCB8fCAwIDogMCksCiAgICAgICAgICBwb3N0czogKGEgPyBhLnBvc3RfY291bnQgfHwgMCA6IDApICsgKGIgPyBiLnBvc3RfY291bnQgfHwgMCA6IDApLAogICAgICAgICAgZW5nYWdlbWVudFRvdGFsOiAoYSA/IGEuZW5nYWdlbWVudCB8fCAwIDogMCkgKyAoYiA/IGIuZW5nYWdlbWVudCB8fCAwIDogMCksCiAgICAgICAgfTsKICAgICAgfSkKICAgICAgLmZpbHRlcigoY2FyZCkgPT4gY2FyZC5tZXRyaWNzLmxlbmd0aCA+IDApOwogIH0KCiAgZnVuY3Rpb24gc29ydENhcmRzKGNhcmRzLCBzb3J0TW9kZSkgewogICAgY29uc3QgYXJyID0gWy4uLmNhcmRzXTsKICAgIGlmIChzb3J0TW9kZSA9PT0gJ2VuZ2FnZW1lbnQnKSByZXR1cm4gYXJyLnNvcnQoKHgsIHkpID0+IHkuZW5nYWdlbWVudFRvdGFsIC0geC5lbmdhZ2VtZW50VG90YWwpOwogICAgaWYgKHNvcnRNb2RlID09PSAnZm9sbG93ZXJzJykgcmV0dXJuIGFyci5zb3J0KCh4LCB5KSA9PiB5LmZvbGxvd2Vyc0dhaW5lZCAtIHguZm9sbG93ZXJzR2FpbmVkKTsKICAgIGlmIChzb3J0TW9kZSA9PT0gJ3Bvc3RzJykgcmV0dXJuIGFyci5zb3J0KCh4LCB5KSA9PiB5LnBvc3RzIC0geC5wb3N0cyk7CiAgICBpZiAoc29ydE1vZGUgPT09ICdhbHBoYScpIHJldHVybiBhcnIuc29ydCgoeCwgeSkgPT4geC5sYWJlbC5sb2NhbGVDb21wYXJlKHkubGFiZWwpKTsKICAgIC8vICdvdmVyYWxsJyBhbmQgJ2dyb3d0aCcgYm90aCByYW5rIGJ5IHRoZSBzYW1lIGNvbXBvc2l0ZSAlIGNoYW5nZSDigJQgdGhlIHR3byBsYWJlbHMKICAgIC8vIHJlYWQgZGlmZmVyZW50bHkgb24gdGhlIHNhbWUgdW5kZXJseWluZyBudW1iZXIsIHBlciB0aGUgcmVxdWVzdGVkIG9wdGlvbiBsaXN0LgogICAgcmV0dXJuIGFyci5zb3J0KCh4LCB5KSA9PiAoeS5vdmVyYWxsID8/IC1JbmZpbml0eSkgLSAoeC5vdmVyYWxsID8/IC1JbmZpbml0eSkpOwogIH0KCiAgZnVuY3Rpb24gYnVpbGRNZXRyaWNSb3cobSkgewogICAgY29uc3Qgcm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICByb3cuY2xhc3NOYW1lID0gJ3BjYy1tZXRyaWMtcm93JzsKICAgIGNvbnN0IGhlYWRlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgaGVhZGVyLmNsYXNzTmFtZSA9ICdwY2MtbWV0cmljLWhlYWRlcic7CiAgICBjb25zdCBmbXQgPSAodikgPT4gKG0uaXNEdXJhdGlvbiA/IEZvcm1hdC5kdXJhdGlvbih2KSA6IEZvcm1hdC5zbWFydCh2KSk7CiAgICBjb25zdCBkaWZmVGV4dCA9IG0ucGN0RGlmZiA9PT0gbnVsbAogICAgICA/IGAke20uZGlmZiA+IDAgPyAnKycgOiAnJ30ke2ZtdChtLmRpZmYpfSAobmV3KWAKICAgICAgOiBgJHttLmRpZmYgPiAwID8gJysnIDogJyd9JHtmbXQobS5kaWZmKX0gKCR7Rm9ybWF0LnBjdChtLnBjdERpZmYpfSlgOwogICAgaGVhZGVyLmFwcGVuZCgKICAgICAgdGV4dEVsKCdzcGFuJywgbS5sYWJlbCwgJ3BjYy1tZXRyaWMtbGFiZWwnKSwKICAgICAgdGV4dEVsKCdzcGFuJywgZGlmZlRleHQsIGBwY2MtbWV0cmljLWRpZmYgJHtGb3JtYXQuZGVsdGFDbGFzcyhtLnBjdERpZmYpfWApCiAgICApOwogICAgcm93LmFwcGVuZENoaWxkKGhlYWRlcik7CiAgICBjb25zdCBtYXggPSBNYXRoLm1heChtLmEsIG0uYiwgMSk7CiAgICAvLyBtLmEgaXMgYWx3YXlzIHRoZSBjdXJyZW50IHBlcmlvZCBhbmQgbS5iIGFsd2F5cyB0aGUgcHJldmlvdXMgcGVyaW9kIChzZWUKICAgIC8vIGNvbXB1dGVDYXJkTWV0cmljcykgcmVnYXJkbGVzcyBvZiBjb21wYXJpc29uIG1vZGUsIHNvIHRoZXNlIGxhYmVscyBjYW4gYmUKICAgIC8vIGhhcmRjb2RlZCByYXRoZXIgdGhhbiBuZWVkaW5nIHRoZSBtb2RlLXNwZWNpZmljIGxhYmVsQS9sYWJlbEIgdGV4dCDigJQgdW5saWtlCiAgICAvLyB0aGUgZ2VuZXJpYyAiUmFuZ2UgQSIvIlJhbmdlIEIiIHdvcmRpbmcgdGhpcyByZXBsYWNlZCwgd2hpY2ggcmVhZCBhcwogICAgLy8gYXJiaXRyYXJ5IGxldHRlcnMgd2l0aCBubyBpbmRpY2F0aW9uIG9mIHdoaWNoIHNpZGUgd2FzIG1vcmUgcmVjZW50LgogICAgcm93LmFwcGVuZENoaWxkKGJ1aWxkQmFyKHsgbGFiZWw6ICdDdXJyZW50JywgdmFsdWU6IG0uYSwgbWF4LCBjb2xvclZhcjogJy0tc2VyaWVzLTEnLCBmb3JtYXRWYWx1ZTogZm10IH0pKTsKICAgIHJvdy5hcHBlbmRDaGlsZChidWlsZEJhcih7IGxhYmVsOiAnUHJldmlvdXMnLCB2YWx1ZTogbS5iLCBtYXgsIGNvbG9yVmFyOiAnLS10ZXh0LW11dGVkJywgZm9ybWF0VmFsdWU6IGZtdCB9KSk7CiAgICByZXR1cm4gcm93OwogIH0KCiAgZnVuY3Rpb24gYnVpbGRDYXJkRm9vdGVyKGNhcmQpIHsKICAgIGNvbnN0IGZvb3RlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZm9vdGVyLmNsYXNzTmFtZSA9ICdwY2MtZm9vdGVyJzsKICAgIGNvbnN0IGRpciA9IHRyZW5kRGlyZWN0aW9uKGNhcmQub3ZlcmFsbCk7CiAgICBjb25zdCByZXN1bHRUZXh0ID0gY2FyZC5vdmVyYWxsID09PSBudWxsCiAgICAgID8gJ05vdCBlbm91Z2ggZGF0YSB0byBjb21wYXJlJwogICAgICA6IGAke2RpciA9PT0gJ3VwJyA/ICdJbXByb3ZlZCcgOiBkaXIgPT09ICdkb3duJyA/ICdEZWNsaW5lZCcgOiAnTm8gc2lnbmlmaWNhbnQgY2hhbmdlJ30ke2RpciAhPT0gJ2ZsYXQnID8gYCBieSAke01hdGguYWJzKGNhcmQub3ZlcmFsbCl9JWAgOiAnJ31gOwogICAgZm9vdGVyLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgJ092ZXJhbGwgUmVzdWx0JywgJ3BjYy1mb290ZXItbGFiZWwnKSk7CiAgICBmb290ZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCByZXN1bHRUZXh0LCBgcGNjLWZvb3Rlci12YWx1ZSAke2Rpcn1gKSk7CiAgICBpZiAoY2FyZC5iZXN0KSBmb290ZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBgQmVzdCBNZXRyaWM6ICR7Y2FyZC5iZXN0LmxhYmVsfSAoJHtGb3JtYXQucGN0KGNhcmQuYmVzdC5wY3REaWZmKX0pYCwgJ3BjYy1mb290ZXItZGV0YWlsJykpOwogICAgaWYgKGNhcmQud2Vha2VzdCAmJiBjYXJkLndlYWtlc3QgIT09IGNhcmQuYmVzdCkgewogICAgICBmb290ZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBgV2Vha2VzdCBNZXRyaWM6ICR7Y2FyZC53ZWFrZXN0LmxhYmVsfSAoJHtGb3JtYXQucGN0KGNhcmQud2Vha2VzdC5wY3REaWZmKX0pYCwgJ3BjYy1mb290ZXItZGV0YWlsJykpOwogICAgfQogICAgcmV0dXJuIGZvb3RlcjsKICB9CgogIC8qKiBTZWxmLWNvbnRhaW5lZCBtb2RhbCBmb3IgIlZpZXcgRnVsbCBDb21wYXJpc29uIiDigJQgYSBzZXBhcmF0ZSBvdmVybGF5IGlkIGZyb20gdGhlIERhdGEgUmVjb3JkcyBFZGl0IG1vZGFsIChSZWNvcmRzLm1vZGFsU2hlbGwgaXMgYSBwcml2YXRlIGNsb3N1cmUgb2YgdGhhdCBtb2R1bGUsIG5vdCBzaGFyZWQgc3RhdGUpLCBzYW1lIHZpc3VhbCBsYW5ndWFnZSAoLm1vZGFsLW92ZXJsYXkgLyAubW9kYWwtcGFuZWwpIHNvIGl0IGxvb2tzIGlkZW50aWNhbC4gKi8KICBmdW5jdGlvbiBjbG9zZUNhcmRNb2RhbCgpIHsKICAgIGNvbnN0IG92ZXJsYXkgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnY29tcGFyaXNvbk1vZGFsT3ZlcmxheScpOwogICAgaWYgKG92ZXJsYXkpIG92ZXJsYXkucmVtb3ZlKCk7CiAgfQoKICBmdW5jdGlvbiBvcGVuQ2FyZE1vZGFsKGNhcmQsIGxhYmVsQSwgbGFiZWxCKSB7CiAgICBjbG9zZUNhcmRNb2RhbCgpOwogICAgY29uc3Qgb3ZlcmxheSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgb3ZlcmxheS5jbGFzc05hbWUgPSAnbW9kYWwtb3ZlcmxheSc7CiAgICBvdmVybGF5LmlkID0gJ2NvbXBhcmlzb25Nb2RhbE92ZXJsYXknOwogICAgb3ZlcmxheS5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIChlKSA9PiB7IGlmIChlLnRhcmdldCA9PT0gb3ZlcmxheSkgY2xvc2VDYXJkTW9kYWwoKTsgfSk7CiAgICBjb25zdCBwYW5lbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcGFuZWwuY2xhc3NOYW1lID0gJ21vZGFsLXBhbmVsIHdpZGUnOwogICAgcGFuZWwuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMicsIGAke2NhcmQubGFiZWx9IOKAlCBGdWxsIENvbXBhcmlzb25gKSk7CiAgICBwYW5lbC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIGAke2xhYmVsQX0gdnMgJHtsYWJlbEJ9YCwgJ21vZGFsLXN1YicpKTsKCiAgICBjb25zdCB0YWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICB0YWJsZS5jbGFzc05hbWUgPSAnZGF0YS10YWJsZSc7CiAgICBjb25zdCB0aGVhZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RoZWFkJyk7CiAgICBjb25zdCBoZWFkUm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgIC8vIGxhYmVsQS9sYWJlbEIgYXJlIHRoZSBzYW1lIG1vZGUtc3BlY2lmaWMgdGV4dCBzaG93biBpbiB0aGUgbW9kYWwncyBzdWJ0aXRsZQogICAgLy8gYWJvdmUgKGUuZy4gIldlZWsgb2YgSmFuIDEzLCAyMDI2IiAvICJXZWVrIG9mIERlYyAzMCwgMjAyNSIsIG9yICJDdXJyZW50CiAgICAvLyBwZXJpb2QiIC8gIlByZXZpb3VzIHBlcmlvZCIpIOKAlCByZXVzZWQgaGVyZSBpbnN0ZWFkIG9mIGdlbmVyaWMgIlJhbmdlIEEiLwogICAgLy8gIlJhbmdlIEIiIHNvIHRoZSBjb2x1bW4gaGVhZGVycyBhbHdheXMgc2F5IHdoYXQgcGVyaW9kIHRoZXkgYWN0dWFsbHkgaG9sZC4KICAgIGNvbnN0IG51bVRoID0gKHRleHQpID0+IHsgY29uc3QgdGggPSB0ZXh0RWwoJ3RoJywgdGV4dCk7IHRoLmNsYXNzTGlzdC5hZGQoJ251bScpOyByZXR1cm4gdGg7IH07CiAgICBoZWFkUm93LmFwcGVuZCh0ZXh0RWwoJ3RoJywgJ01ldHJpYycpLCBudW1UaChsYWJlbEEpLCBudW1UaChsYWJlbEIpLCBudW1UaCgnRGlmZmVyZW5jZScpLCBudW1UaCgnJSBEaWZmZXJlbmNlJyksIHRleHRFbCgndGgnLCAnVHJlbmQnKSk7CiAgICB0aGVhZC5hcHBlbmRDaGlsZChoZWFkUm93KTsKICAgIGNvbnN0IHRib2R5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGJvZHknKTsKICAgIGNhcmQubWV0cmljcy5mb3JFYWNoKChtKSA9PiB7CiAgICAgIGNvbnN0IGZtdCA9ICh2KSA9PiAobS5pc0R1cmF0aW9uID8gRm9ybWF0LmR1cmF0aW9uKHYpIDogRm9ybWF0LnNtYXJ0KHYpKTsKICAgICAgY29uc3QgdHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgICBjb25zdCB0cmVuZEVsID0gdGV4dEVsKCdzcGFuJywgdHJlbmREaXJlY3Rpb24obS5wY3REaWZmKSA9PT0gJ3VwJyA/ICfilrInIDogdHJlbmREaXJlY3Rpb24obS5wY3REaWZmKSA9PT0gJ2Rvd24nID8gJ+KWvCcgOiAn4oCUJywgYHN0YXQtZGVsdGEgJHtGb3JtYXQuZGVsdGFDbGFzcyhtLnBjdERpZmYpfWApOwogICAgICBjb25zdCB0cmVuZFRkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgdHJlbmRUZC5hcHBlbmRDaGlsZCh0cmVuZEVsKTsKICAgICAgdHIuYXBwZW5kKAogICAgICAgIHRleHRFbCgndGQnLCBtLmxhYmVsKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgZm10KG0uYSksICdudW0nKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgZm10KG0uYiksICdudW0nKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgYCR7bS5kaWZmID4gMCA/ICcrJyA6ICcnfSR7Zm10KG0uZGlmZil9YCwgJ251bScpLAogICAgICAgIHRleHRFbCgndGQnLCBtLnBjdERpZmYgPT09IG51bGwgPyAnbmV3JyA6IEZvcm1hdC5wY3QobS5wY3REaWZmKSwgJ251bScpLAogICAgICAgIHRyZW5kVGQKICAgICAgKTsKICAgICAgdGJvZHkuYXBwZW5kQ2hpbGQodHIpOwogICAgfSk7CiAgICB0YWJsZS5hcHBlbmQodGhlYWQsIHRib2R5KTsKICAgIGNvbnN0IHRhYmxlU2Nyb2xsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0YWJsZVNjcm9sbC5jbGFzc05hbWUgPSAndGFibGUtc2Nyb2xsJzsKICAgIHRhYmxlU2Nyb2xsLmFwcGVuZENoaWxkKHRhYmxlKTsKICAgIHBhbmVsLmFwcGVuZENoaWxkKHRhYmxlU2Nyb2xsKTsKCiAgICBjb25zdCBhY3Rpb25zID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBhY3Rpb25zLmNsYXNzTmFtZSA9ICdtb2RhbC1hY3Rpb25zJzsKICAgIGNvbnN0IGNsb3NlQnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBjbG9zZUJ0bi5jbGFzc05hbWUgPSAnYnRuJzsKICAgIGNsb3NlQnRuLnR5cGUgPSAnYnV0dG9uJzsKICAgIGNsb3NlQnRuLnRleHRDb250ZW50ID0gJ0Nsb3NlJzsKICAgIGNsb3NlQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgY2xvc2VDYXJkTW9kYWwpOwogICAgYWN0aW9ucy5hcHBlbmRDaGlsZChjbG9zZUJ0bik7CiAgICBwYW5lbC5hcHBlbmRDaGlsZChhY3Rpb25zKTsKCiAgICBvdmVybGF5LmFwcGVuZENoaWxkKHBhbmVsKTsKICAgIGRvY3VtZW50LmJvZHkuYXBwZW5kQ2hpbGQob3ZlcmxheSk7CiAgfQoKICBmdW5jdGlvbiBidWlsZFBsYXRmb3JtQ2FyZChjYXJkLCBsYWJlbEEsIGxhYmVsQikgewogICAgY29uc3QgZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGVsLmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1jb21wYXJlLWNhcmQnOwoKICAgIGNvbnN0IGhlYWRlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgaGVhZGVyLmNsYXNzTmFtZSA9ICdwY2MtaGVhZGVyJzsKICAgIGNvbnN0IG5hbWVXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBuYW1lV3JhcC5jbGFzc05hbWUgPSAncGNjLWhlYWRlci1uYW1lJzsKICAgIGNvbnN0IGRvdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgIGRvdC5jbGFzc05hbWUgPSAncGxhdGZvcm0tZG90JzsKICAgIGRvdC5zdHlsZS5iYWNrZ3JvdW5kID0gY2FyZC5jb2xvcjsKICAgIG5hbWVXcmFwLmFwcGVuZChkb3QsIHRleHRFbCgnc3BhbicsIGNhcmQubGFiZWwsICdwY2MtbmFtZScpKTsKICAgIGhlYWRlci5hcHBlbmRDaGlsZChuYW1lV3JhcCk7CiAgICBjb25zdCBkaXIgPSB0cmVuZERpcmVjdGlvbihjYXJkLm92ZXJhbGwpOwogICAgaGVhZGVyLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgY2FyZC5vdmVyYWxsID09PSBudWxsID8gJ+KAlCcgOiBgJHtkaXIgPT09ICd1cCcgPyAn4payJyA6IGRpciA9PT0gJ2Rvd24nID8gJ+KWvCcgOiAn4oCUJ30gJHtGb3JtYXQucGN0KGNhcmQub3ZlcmFsbCl9YCwgYHBjYy1iYWRnZSAke2Rpcn1gKSk7CiAgICBlbC5hcHBlbmRDaGlsZChoZWFkZXIpOwoKICAgIGNvbnN0IGNhcHRpb24gPSBjYXJkLm92ZXJhbGwgPT09IG51bGwKICAgICAgPyAnTm90IGVub3VnaCBkYXRhIHRvIGNvbXBhcmUgeWV0JwogICAgICA6IGRpciA9PT0gJ3VwJyA/ICdJbXByb3ZlZCBjb21wYXJlZCB0byBwcmV2aW91cyBwZXJpb2QnCiAgICAgIDogZGlyID09PSAnZG93bicgPyAnTG93ZXIgdGhhbiBwcmV2aW91cyBwZXJpb2QnCiAgICAgIDogJ0Fib3V0IHRoZSBzYW1lIGFzIHRoZSBwcmV2aW91cyBwZXJpb2QnOwogICAgZWwuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBjYXB0aW9uLCAncGNjLWNhcHRpb24nKSk7CgogICAgY29uc3QgbWV0cmljc1dyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIG1ldHJpY3NXcmFwLmNsYXNzTmFtZSA9ICdwY2MtbWV0cmljcyc7CiAgICBjYXJkLm1ldHJpY3MuZm9yRWFjaCgobSkgPT4gbWV0cmljc1dyYXAuYXBwZW5kQ2hpbGQoYnVpbGRNZXRyaWNSb3cobSkpKTsKICAgIGVsLmFwcGVuZENoaWxkKG1ldHJpY3NXcmFwKTsKCiAgICBlbC5hcHBlbmRDaGlsZChidWlsZENhcmRGb290ZXIoY2FyZCkpOwoKICAgIGNvbnN0IHZpZXdMaW5rID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICB2aWV3TGluay50eXBlID0gJ2J1dHRvbic7CiAgICB2aWV3TGluay5jbGFzc05hbWUgPSAncGNjLXZpZXctbGluayc7CiAgICB2aWV3TGluay50ZXh0Q29udGVudCA9ICdWaWV3IEZ1bGwgQ29tcGFyaXNvbiDihpInOwogICAgdmlld0xpbmsuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBvcGVuQ2FyZE1vZGFsKGNhcmQsIGxhYmVsQSwgbGFiZWxCKSk7CiAgICBlbC5hcHBlbmRDaGlsZCh2aWV3TGluayk7CgogICAgcmV0dXJuIGVsOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyUGxhdGZvcm1Db21wYXJpc29uQ2FyZHMod3JhcCwgcmVzdWx0LCBsYWJlbEEsIGxhYmVsQikgewogICAgY29uc3QgYWxsQ2FyZHMgPSBidWlsZFBsYXRmb3JtQ2FyZHMocmVzdWx0KTsKCiAgICBjb25zdCBzZWN0aW9uID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBzZWN0aW9uLmNsYXNzTmFtZSA9ICdwY2Mtc2VjdGlvbic7CiAgICBzZWN0aW9uLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgJ1BsYXRmb3JtIFBlcmZvcm1hbmNlIENvbXBhcmlzb24nLCAnc2VjdGlvbi10aXRsZScpKTsKCiAgICBpZiAoIWFsbENhcmRzLmxlbmd0aCkgewogICAgICBzZWN0aW9uLmFwcGVuZENoaWxkKGVtcHR5U3RhdGUoewogICAgICAgIGljb246ICdnaXQtY29tcGFyZScsCiAgICAgICAgdGl0bGU6ICdObyBkYXRhIGF2YWlsYWJsZSBmb3IgdGhlIHNlbGVjdGVkIGRhdGUgcmFuZ2VzLicsCiAgICAgICAgbWVzc2FnZTogJ1RyeSBhIHdpZGVyIHJhbmdlLCBvciBjaGVjayB0aGF0IHBvc3RzIGV4aXN0IGZvciBhdCBsZWFzdCBvbmUgcGxhdGZvcm0gaW4gUmFuZ2UgQSBvciBSYW5nZSBCLicsCiAgICAgIH0pKTsKICAgICAgd3JhcC5hcHBlbmRDaGlsZChzZWN0aW9uKTsKICAgICAgcmV0dXJuOwogICAgfQoKICAgIGNvbnN0IGNvbnRyb2xzID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjb250cm9scy5jbGFzc05hbWUgPSAncGNjLWNvbnRyb2xzJzsKCiAgICBjb25zdCBzb3J0U2VsZWN0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc2VsZWN0Jyk7CiAgICBDQVJEX1NPUlRfTU9ERVMuZm9yRWFjaCgobSkgPT4gewogICAgICBjb25zdCBvcHQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsgb3B0LnZhbHVlID0gbS5rZXk7IG9wdC50ZXh0Q29udGVudCA9IG0ubGFiZWw7CiAgICAgIGlmIChtLmtleSA9PT0gY2FyZFNvcnRNb2RlKSBvcHQuc2VsZWN0ZWQgPSB0cnVlOwogICAgICBzb3J0U2VsZWN0LmFwcGVuZENoaWxkKG9wdCk7CiAgICB9KTsKICAgIHNvcnRTZWxlY3QuYWRkRXZlbnRMaXN0ZW5lcignY2hhbmdlJywgKCkgPT4geyBjYXJkU29ydE1vZGUgPSBzb3J0U2VsZWN0LnZhbHVlOyByZW5kZXJDYXJkR3JpZCgpOyB9KTsKICAgIGNvbnRyb2xzLmFwcGVuZENoaWxkKGxhYmVsZWQoJ1NvcnQgQnknLCBzb3J0U2VsZWN0KSk7CgogICAgY29uc3QgZmlsdGVyUGlsbHMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGZpbHRlclBpbGxzLmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1maWx0ZXItcGlsbHMnOwogICAgY29uc3QgYWxsQnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBhbGxCdG4udHlwZSA9ICdidXR0b24nOwogICAgYWxsQnRuLmRhdGFzZXQuZmlsdGVyID0gJ2FsbCc7CiAgICBhbGxCdG4udGV4dENvbnRlbnQgPSAnQWxsIFBsYXRmb3Jtcyc7CiAgICBhbGxCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7IGNhcmRQbGF0Zm9ybUZpbHRlciA9ICdhbGwnOyByZW5kZXJDYXJkR3JpZCgpOyB9KTsKICAgIGZpbHRlclBpbGxzLmFwcGVuZENoaWxkKGFsbEJ0bik7CiAgICBhbGxDYXJkcy5mb3JFYWNoKChjYXJkKSA9PiB7CiAgICAgIGNvbnN0IGJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgICBidG4udHlwZSA9ICdidXR0b24nOwogICAgICBidG4uZGF0YXNldC5maWx0ZXIgPSBjYXJkLnBsYXRmb3JtOwogICAgICBjb25zdCBkb3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICAgIGRvdC5jbGFzc05hbWUgPSAncGxhdGZvcm0tZG90JzsKICAgICAgZG90LnN0eWxlLmJhY2tncm91bmQgPSBjYXJkLmNvbG9yOwogICAgICBidG4uYXBwZW5kKGRvdCwgZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUoY2FyZC5sYWJlbCkpOwogICAgICBidG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7IGNhcmRQbGF0Zm9ybUZpbHRlciA9IGNhcmQucGxhdGZvcm07IHJlbmRlckNhcmRHcmlkKCk7IH0pOwogICAgICBmaWx0ZXJQaWxscy5hcHBlbmRDaGlsZChidG4pOwogICAgfSk7CiAgICBjb250cm9scy5hcHBlbmRDaGlsZChmaWx0ZXJQaWxscyk7CiAgICBzZWN0aW9uLmFwcGVuZENoaWxkKGNvbnRyb2xzKTsKCiAgICBjb25zdCBncmlkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBncmlkLmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1jb21wYXJlLWdyaWQnOwogICAgZ3JpZC5pZCA9ICdwbGF0Zm9ybUNvbXBhcmVHcmlkJzsKICAgIHNlY3Rpb24uYXBwZW5kQ2hpbGQoZ3JpZCk7CiAgICB3cmFwLmFwcGVuZENoaWxkKHNlY3Rpb24pOwoKICAgIGZ1bmN0aW9uIHJlbmRlckNhcmRHcmlkKCkgewogICAgICBjb25zdCBncmlkRWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncGxhdGZvcm1Db21wYXJlR3JpZCcpOwogICAgICBpZiAoIWdyaWRFbCkgcmV0dXJuOwogICAgICBncmlkRWwuaW5uZXJIVE1MID0gJyc7CiAgICAgIGNvbnN0IHZpc2libGUgPSBjYXJkUGxhdGZvcm1GaWx0ZXIgPT09ICdhbGwnID8gYWxsQ2FyZHMgOiBhbGxDYXJkcy5maWx0ZXIoKGMpID0+IGMucGxhdGZvcm0gPT09IGNhcmRQbGF0Zm9ybUZpbHRlcik7CiAgICAgIGNvbnN0IHNvcnRlZCA9IHNvcnRDYXJkcyh2aXNpYmxlLCBjYXJkU29ydE1vZGUpOwogICAgICBpZiAoIXNvcnRlZC5sZW5ndGgpIHsKICAgICAgICBncmlkRWwuYXBwZW5kQ2hpbGQoZW1wdHlTdGF0ZSh7IGljb246ICdnaXQtY29tcGFyZScsIG1lc3NhZ2U6ICdObyBkYXRhIGZvciB0aGlzIHBsYXRmb3JtIGluIHRoZSBzZWxlY3RlZCBkYXRlIHJhbmdlcy4nIH0pKTsKICAgICAgfSBlbHNlIHsKICAgICAgICBzb3J0ZWQuZm9yRWFjaCgoY2FyZCkgPT4gZ3JpZEVsLmFwcGVuZENoaWxkKGJ1aWxkUGxhdGZvcm1DYXJkKGNhcmQsIGxhYmVsQSwgbGFiZWxCKSkpOwogICAgICB9CiAgICAgIGZpbHRlclBpbGxzLnF1ZXJ5U2VsZWN0b3JBbGwoJ2J1dHRvbicpLmZvckVhY2goKGJ0bikgPT4gewogICAgICAgIGJ0bi5jbGFzc0xpc3QudG9nZ2xlKCdpcy1hY3RpdmUnLCBidG4uZGF0YXNldC5maWx0ZXIgPT09IGNhcmRQbGF0Zm9ybUZpbHRlcik7CiAgICAgIH0pOwogICAgfQogICAgcmVuZGVyQ2FyZEdyaWQoKTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlcigpIHsKICAgIHJvb3QgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndmlldy1jb21wYXJpc29uJyk7CiAgICBzaGVsbCgpOwogIH0KCiAgcmV0dXJuIHsgcmVuZGVyIH07Cn0pKCk7CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgVXBsb2FkIHRhYjogZHJhZy1kcm9wLCB2YWxpZGF0aW9uIHByZXZpZXcsIHBlci13ZWVrIGNvbmZsaWN0CiAgIHJlc29sdXRpb24sIGNvbW1pdCDigJQgcGx1cyB0aGUgVXBsb2FkIEhpc3RvcnkgdGFiLgogICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KY29uc3QgVXBsb2FkID0gKCgpID0+IHsKICBsZXQgcm9vdDsKICBsZXQgY3VycmVudFByZXZpZXcgPSBudWxsOyAvLyB7IGZpbGVQYXRoLCBvcmlnaW5hbE5hbWUsIGR1cGxpY2F0ZXMsIGlzc3Vlcywgc2FtcGxlLCAuLi4gfQogIGNvbnN0IGR1cGxpY2F0ZUFjdGlvbk92ZXJyaWRlcyA9IHt9OwoKICBmdW5jdGlvbiBzaGVsbCgpIHsKICAgIHJvb3QuaW5uZXJIVE1MID0gJyc7CgogICAgY29uc3QgaW50cm8gPSB0ZXh0RWwoJ2RpdicsICdVcGxvYWQgYSB3ZWVrbHkgZXhwb3J0JywgJ3NlY3Rpb24tdGl0bGUnKTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQoaW50cm8pOwoKICAgIGNvbnN0IGRyb3B6b25lID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBkcm9wem9uZS5jbGFzc05hbWUgPSAnZHJvcHpvbmUnOwogICAgZHJvcHpvbmUuaWQgPSAnZHJvcHpvbmUnOwogICAgZHJvcHpvbmUuaW5uZXJIVE1MID0gYAogICAgICA8ZGl2IGNsYXNzPSJlbXB0eS1pY29uIiBzdHlsZT0ibWFyZ2luOiAwIGF1dG8gMTRweDsiPjxpIGRhdGEtbHVjaWRlPSJ1cGxvYWQtY2xvdWQiIHN0eWxlPSJ3aWR0aDoyMnB4O2hlaWdodDoyMnB4OyI+PC9pPjwvZGl2PgogICAgICA8aDM+RHJhZyAmYW1wOyBkcm9wIHlvdXIgLmNzdiBvciAueGxzeCBmaWxlIGhlcmU8L2gzPgogICAgICA8cD5vciBjbGljayB0byBicm93c2Ug4oCUIGZpbGVzIGFyZSB2YWxpZGF0ZWQgYmVmb3JlIGFueXRoaW5nIGlzIHNhdmVkPC9wPgogICAgICA8aW5wdXQgdHlwZT0iZmlsZSIgaWQ9ImZpbGVJbnB1dCIgYWNjZXB0PSIuY3N2LC54bHN4LC54bHMiIC8+CiAgICBgOwogICAgcm9vdC5hcHBlbmRDaGlsZChkcm9wem9uZSk7CgogICAgY29uc3QgcHJldmlld0FyZWEgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHByZXZpZXdBcmVhLmlkID0gJ3ByZXZpZXdBcmVhJzsKICAgIHJvb3QuYXBwZW5kQ2hpbGQocHJldmlld0FyZWEpOwoKICAgIHdpcmVEcm9wem9uZShkcm9wem9uZSk7CiAgfQoKICBmdW5jdGlvbiB3aXJlRHJvcHpvbmUoZHJvcHpvbmUpIHsKICAgIGNvbnN0IGlucHV0ID0gZHJvcHpvbmUucXVlcnlTZWxlY3RvcignI2ZpbGVJbnB1dCcpOwogICAgZHJvcHpvbmUuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBpbnB1dC5jbGljaygpKTsKICAgIGlucHV0LmFkZEV2ZW50TGlzdGVuZXIoJ2NoYW5nZScsICgpID0+IHsKICAgICAgaWYgKGlucHV0LmZpbGVzWzBdKSBoYW5kbGVGaWxlKGlucHV0LmZpbGVzWzBdKTsKICAgIH0pOwogICAgWydkcmFnZW50ZXInLCAnZHJhZ292ZXInXS5mb3JFYWNoKChldnQpID0+CiAgICAgIGRyb3B6b25lLmFkZEV2ZW50TGlzdGVuZXIoZXZ0LCAoZSkgPT4geyBlLnByZXZlbnREZWZhdWx0KCk7IGRyb3B6b25lLmNsYXNzTGlzdC5hZGQoJ2lzLWRyYWcnKTsgfSkKICAgICk7CiAgICBbJ2RyYWdsZWF2ZScsICdkcm9wJ10uZm9yRWFjaCgoZXZ0KSA9PgogICAgICBkcm9wem9uZS5hZGRFdmVudExpc3RlbmVyKGV2dCwgKGUpID0+IHsgZS5wcmV2ZW50RGVmYXVsdCgpOyBkcm9wem9uZS5jbGFzc0xpc3QucmVtb3ZlKCdpcy1kcmFnJyk7IH0pCiAgICApOwogICAgZHJvcHpvbmUuYWRkRXZlbnRMaXN0ZW5lcignZHJvcCcsIChlKSA9PiB7CiAgICAgIGNvbnN0IGZpbGUgPSBlLmRhdGFUcmFuc2Zlci5maWxlc1swXTsKICAgICAgaWYgKGZpbGUpIGhhbmRsZUZpbGUoZmlsZSk7CiAgICB9KTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIGhhbmRsZUZpbGUoZmlsZSkgewogICAgY29uc3QgYXJlYSA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdwcmV2aWV3QXJlYScpOwogICAgYXJlYS5pbm5lckhUTUwgPSAnJzsKICAgIGFyZWEuYXBwZW5kQ2hpbGQocm93V2l0aFNwaW5uZXIoJ1ZhbGlkYXRpbmcgZmlsZeKApicpKTsKICAgIE9iamVjdC5rZXlzKGR1cGxpY2F0ZUFjdGlvbk92ZXJyaWRlcykuZm9yRWFjaCgoaykgPT4gZGVsZXRlIGR1cGxpY2F0ZUFjdGlvbk92ZXJyaWRlc1trXSk7CiAgICB0cnkgewogICAgICBjdXJyZW50UHJldmlldyA9IGF3YWl0IEFwaS5wcmV2aWV3VXBsb2FkKGZpbGUpOwogICAgICByZW5kZXJQcmV2aWV3KCk7CiAgICB9IGNhdGNoIChlcnIpIHsKICAgICAgYXJlYS5pbm5lckhUTUwgPSAnJzsKICAgICAgYXJlYS5hcHBlbmRDaGlsZChlcnJvckJhbm5lcihlcnIubWVzc2FnZSkpOwogICAgfQogIH0KCiAgZnVuY3Rpb24gcm93V2l0aFNwaW5uZXIodGV4dCkgewogICAgY29uc3QgZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGVsLmNsYXNzTmFtZSA9ICdsb2FkaW5nLXJvdyc7CiAgICBjb25zdCBzcGlubmVyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgc3Bpbm5lci5jbGFzc05hbWUgPSAnc3Bpbm5lcic7CiAgICBlbC5hcHBlbmQoc3Bpbm5lciwgZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUoYCAke3RleHR9YCkpOwogICAgcmV0dXJuIGVsOwogIH0KICBmdW5jdGlvbiBlcnJvckJhbm5lcihtZXNzYWdlKSB7CiAgICBjb25zdCBlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZWwuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgZWwuc3R5bGUuYm9yZGVyTGVmdCA9ICczcHggc29saWQgdmFyKC0tc3RhdHVzLWNyaXRpY2FsKSc7CiAgICBlbC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIGBDb3VsZCBub3QgcmVhZCB0aGlzIGZpbGU6ICR7bWVzc2FnZX1gLCAnbXV0ZWQnKSk7CiAgICByZXR1cm4gZWw7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJQcmV2aWV3KCkgewogICAgY29uc3QgYXJlYSA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdwcmV2aWV3QXJlYScpOwogICAgYXJlYS5pbm5lckhUTUwgPSAnJzsKICAgIGNvbnN0IHAgPSBjdXJyZW50UHJldmlldzsKCiAgICBjb25zdCBzdW1tYXJ5VGl0bGUgPSB0ZXh0RWwoJ2RpdicsICdWYWxpZGF0aW9uIHN1bW1hcnknLCAnc2VjdGlvbi10aXRsZScpOwogICAgY29uc3Qgc3VtbWFyeUdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHN1bW1hcnlHcmlkLmNsYXNzTmFtZSA9ICdzdGF0LWdyaWQnOwogICAgc3VtbWFyeUdyaWQuYXBwZW5kKAogICAgICBzdGF0VGlsZSgnRmlsZScsIHAub3JpZ2luYWxOYW1lKSwKICAgICAgc3RhdFRpbGUoJ1NoZWV0cyBmb3VuZCcsIHAuc2hlZXRzLmxlbmd0aCksCiAgICAgIHN0YXRUaWxlKCdUb3RhbCByb3dzIChhbGwgc2hlZXRzKScsIHAudG90YWxEYXRhUm93cyksCiAgICAgIHN0YXRUaWxlKCdOZXcgcmVjb3JkcycsIHAubmV3UmVjb3Jkc0NvdW50KSwKICAgICAgc3RhdFRpbGUoJ0V4YWN0IGR1cGxpY2F0ZXMgZm91bmQnLCBwLmR1cGxpY2F0ZXMubGVuZ3RoKSwKICAgICAgc3RhdFRpbGUoJ0R1cGxpY2F0ZSByb3dzIGluIGZpbGUnLCBwLmR1cGxpY2F0ZVJvd3NJbkZpbGUpLAogICAgICBzdGF0VGlsZSgnUm93cyB3aXRoIGVycm9ycycsIHAuZXJyb3JSb3dzKQogICAgKTsKICAgIGFyZWEuYXBwZW5kKHN1bW1hcnlUaXRsZSwgc3VtbWFyeUdyaWQpOwoKICAgIGlmIChwLnNoZWV0cy5sZW5ndGgpIHsKICAgICAgY29uc3Qgc2hlZXRzVGl0bGUgPSB0ZXh0RWwoJ2RpdicsICdTaGVldCBicmVha2Rvd24nLCAnc2VjdGlvbi10aXRsZScpOwogICAgICBjb25zdCBzaGVldHNUYWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICAgIHNoZWV0c1RhYmxlLmNsYXNzTmFtZSA9ICdkYXRhLXRhYmxlJzsKICAgICAgc2hlZXRzVGFibGUuaW5uZXJIVE1MID0gJzx0aGVhZD48dHI+PHRoPlNoZWV0PC90aD48dGg+TGF5b3V0IGRldGVjdGVkPC90aD48dGggY2xhc3M9Im51bSI+Um93czwvdGg+PHRoIGNsYXNzPSJudW0iPlZhbGlkPC90aD48dGggY2xhc3M9Im51bSI+RXJyb3JzPC90aD48L3RyPjwvdGhlYWQ+JzsKICAgICAgY29uc3Qgc2hlZXRzQm9keSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3Rib2R5Jyk7CiAgICAgIHAuc2hlZXRzLmZvckVhY2goKHMpID0+IHsKICAgICAgICBjb25zdCB0ciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAgICAgY29uc3QgbGF5b3V0TGFiZWwgPSBzLmZvcm1hdCA9PT0gJ2FnZW5kYScgPyAnTFJTIGFnZW5kYSB0cmFja2VyJyA6IHMuZm9ybWF0ID09PSAnc2ltcGxlJyA/ICdTaW1wbGUgcGxhdGZvcm0gdGFibGUnIDogJ05vdCByZWNvZ25pemVkIOKAlCBzYXZlZCBhcyByYXcgZGF0YSBvbmx5JzsKICAgICAgICB0ci5hcHBlbmQoCiAgICAgICAgICB0ZXh0RWwoJ3RkJywgcy5uYW1lKSwKICAgICAgICAgIHRleHRFbCgndGQnLCBsYXlvdXRMYWJlbCksCiAgICAgICAgICB0ZXh0RWwoJ3RkJywgU3RyaW5nKHMudG90YWxSb3dzKSwgJ251bScpLAogICAgICAgICAgdGV4dEVsKCd0ZCcsIFN0cmluZyhzLnZhbGlkUm93cyksICdudW0nKSwKICAgICAgICAgIHRleHRFbCgndGQnLCBTdHJpbmcocy5lcnJvclJvd3MpLCAnbnVtJykKICAgICAgICApOwogICAgICAgIHNoZWV0c0JvZHkuYXBwZW5kQ2hpbGQodHIpOwogICAgICB9KTsKICAgICAgc2hlZXRzVGFibGUuYXBwZW5kQ2hpbGQoc2hlZXRzQm9keSk7CiAgICAgIGNvbnN0IHNoZWV0c1dyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgc2hlZXRzV3JhcC5jbGFzc05hbWUgPSAndGFibGUtc2Nyb2xsJzsKICAgICAgc2hlZXRzV3JhcC5hcHBlbmRDaGlsZChzaGVldHNUYWJsZSk7CiAgICAgIGFyZWEuYXBwZW5kKHNoZWV0c1RpdGxlLCBzaGVldHNXcmFwKTsKICAgIH0KCiAgICBpZiAocC5kdXBsaWNhdGVzLmxlbmd0aCkgewogICAgICBjb25zdCBkdXBUaXRsZSA9IHRleHRFbCgnZGl2JywgYEV4YWN0IGR1cGxpY2F0ZXMgZm91bmQgKCR7cC5kdXBsaWNhdGVzLmxlbmd0aH0pYCwgJ3NlY3Rpb24tdGl0bGUnKTsKICAgICAgYXJlYS5hcHBlbmRDaGlsZChkdXBUaXRsZSk7CiAgICAgIGFyZWEuYXBwZW5kQ2hpbGQodGV4dEVsKAogICAgICAgICdkaXYnLAogICAgICAgICdFYWNoIG9mIHRoZXNlIHJvd3MgaXMgYnl0ZS1mb3ItYnl0ZSBpZGVudGljYWwgdG8gYW4gYWxyZWFkeS1zYXZlZCByZWNvcmQg4oCUIGV2ZXJ5IGZpZWxkIG1hdGNoZXMsIGluY2x1ZGluZyBldmVyeSBtZXRyaWMsIG5vdCBqdXN0IHRoZSBkYXRlL2NhcHRpb24vcGxhdGZvcm0uIENob29zZSB3aGF0IHRvIGRvIHdpdGggZWFjaCDigJQgb3Igc2V0IGEgZGVmYXVsdCBmb3IgYWxsIG9mIHRoZW0uIChBIHJvdyB0aGF0IHNoYXJlcyB0aGUgc2FtZSBkYXRlL2NhcHRpb24vcGxhdGZvcm0gYnV0IGhhcyBkaWZmZXJlbnQgbnVtYmVycyBpcyBub3Qgc2hvd24gaGVyZSDigJQgaXTigJlzIGltcG9ydGVkIGF1dG9tYXRpY2FsbHkgYXMgaXRzIG93biBuZXcgcmVjb3JkLCBzaW5jZSBpdHMgYW5hbHl0aWNzIGNoYW5nZWQuKScsCiAgICAgICAgJ211dGVkJwogICAgICApKTsKICAgICAgY29uc3QgZGVmYXVsdFJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBkZWZhdWx0Um93LmNsYXNzTmFtZSA9ICdmaWVsZC1pbmxpbmUnOwogICAgICBkZWZhdWx0Um93LnN0eWxlLm1hcmdpbiA9ICcxMHB4IDAnOwogICAgICBjb25zdCBkZWZhdWx0U2VsZWN0ID0gYWN0aW9uU2VsZWN0KCdza2lwJyk7CiAgICAgIGRlZmF1bHRTZWxlY3QuaWQgPSAnZGVmYXVsdER1cGxpY2F0ZUFjdGlvblNlbGVjdCc7CiAgICAgIGRlZmF1bHRTZWxlY3QuYWRkRXZlbnRMaXN0ZW5lcignY2hhbmdlJywgKCkgPT4gewogICAgICAgIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJy5jb25mbGljdC1yb3cgc2VsZWN0W2RhdGEtaGFzaF0nKS5mb3JFYWNoKChzZWwpID0+IHsKICAgICAgICAgIGlmICghZHVwbGljYXRlQWN0aW9uT3ZlcnJpZGVzW3NlbC5kYXRhc2V0Lmhhc2hdKSBzZWwudmFsdWUgPSBkZWZhdWx0U2VsZWN0LnZhbHVlOwogICAgICAgIH0pOwogICAgICB9KTsKICAgICAgZGVmYXVsdFJvdy5hcHBlbmQodGV4dEVsKCdsYWJlbCcsICdEZWZhdWx0IGFjdGlvbiBmb3IgYWxsIG1hdGNoZXMnKSwgZGVmYXVsdFNlbGVjdCk7CiAgICAgIGFyZWEuYXBwZW5kQ2hpbGQoZGVmYXVsdFJvdyk7CgogICAgICBjb25zdCBsaXN0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIGxpc3QuY2xhc3NOYW1lID0gJ2NvbmZsaWN0LWxpc3QnOwogICAgICBwLmR1cGxpY2F0ZXMuZm9yRWFjaCgoZCkgPT4gewogICAgICAgIGNvbnN0IHJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICAgIHJvdy5jbGFzc05hbWUgPSAnY29uZmxpY3Qtcm93JzsKICAgICAgICBjb25zdCBsZWZ0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgICAgbGVmdC5hcHBlbmQoCiAgICAgICAgICB0ZXh0RWwoJ2RpdicsIGAke0Zvcm1hdC5kYXRlKGQucHVibGlzaERhdGUpfSDigJQgJHsoZC5jYXB0aW9uIHx8ICcobm8gY2FwdGlvbiknKS5zbGljZSgwLCA3MCl9YCwgJ3dlZWstbGFiZWwnKSwKICAgICAgICAgIHRleHRFbCgnZGl2JywgYEV4YWN0IG1hdGNoIG9mIGV4aXN0aW5nIHJlY29yZCAjJHtkLmV4aXN0aW5nLnBvc3RJZH0gKGxhc3QgdXBkYXRlZCAke2QuZXhpc3RpbmcudXBkYXRlZEF0fSlgLCAnd2Vlay1tZXRhJykKICAgICAgICApOwogICAgICAgIHJvdy5hcHBlbmRDaGlsZChsZWZ0KTsKICAgICAgICBjb25zdCBzZWwgPSBhY3Rpb25TZWxlY3QoJ3NraXAnKTsKICAgICAgICBzZWwuZGF0YXNldC5oYXNoID0gZC5oYXNoOwogICAgICAgIHNlbC5hZGRFdmVudExpc3RlbmVyKCdjaGFuZ2UnLCAoKSA9PiB7IGR1cGxpY2F0ZUFjdGlvbk92ZXJyaWRlc1tkLmhhc2hdID0gc2VsLnZhbHVlOyB9KTsKICAgICAgICByb3cuYXBwZW5kQ2hpbGQoc2VsKTsKICAgICAgICBsaXN0LmFwcGVuZENoaWxkKHJvdyk7CiAgICAgIH0pOwogICAgICBhcmVhLmFwcGVuZENoaWxkKGxpc3QpOwogICAgfQoKICAgIGNvbnN0IG5vdGVzRmllbGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIG5vdGVzRmllbGQuY2xhc3NOYW1lID0gJ2Zvcm0tZmllbGQnOwogICAgbm90ZXNGaWVsZC5zdHlsZS5tYXJnaW4gPSAnMTJweCAwJzsKICAgIG5vdGVzRmllbGQuYXBwZW5kQ2hpbGQodGV4dEVsKCdsYWJlbCcsICdVcGxvYWQgbm90ZXMgKG9wdGlvbmFsKScpKTsKICAgIGNvbnN0IG5vdGVzSW5wdXQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOwogICAgbm90ZXNJbnB1dC50eXBlID0gJ3RleHQnOwogICAgbm90ZXNJbnB1dC5pZCA9ICd1cGxvYWROb3Rlc0lucHV0JzsKICAgIG5vdGVzSW5wdXQucGxhY2Vob2xkZXIgPSAnZS5nLiAiV2VlayAzIGV4cG9ydCwgaW5jbHVkZXMgY29ycmVjdGVkIFRpa1RvayBudW1iZXJzIic7CiAgICBub3Rlc0ZpZWxkLmFwcGVuZENoaWxkKG5vdGVzSW5wdXQpOwogICAgYXJlYS5hcHBlbmRDaGlsZChub3Rlc0ZpZWxkKTsKCiAgICBpZiAocC5pc3N1ZXMubGVuZ3RoKSB7CiAgICAgIGNvbnN0IGlzc3Vlc1RpdGxlID0gdGV4dEVsKCdkaXYnLCBgUm93cyBza2lwcGVkIG9yIGZsYWdnZWQgKCR7cC5pc3N1ZXMubGVuZ3RofSlgLCAnc2VjdGlvbi10aXRsZScpOwogICAgICBjb25zdCBpc3N1ZXNDYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIGlzc3Vlc0NhcmQuY2xhc3NOYW1lID0gJ2lzc3Vlcy1saXN0JzsKICAgICAgcC5pc3N1ZXMuZm9yRWFjaCgoaXNzdWUpID0+IHsKICAgICAgICBjb25zdCByb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgICByb3cuY2xhc3NOYW1lID0gJ2lzc3VlLXJvdyc7CiAgICAgICAgaWYgKGlzc3VlLnJvd051bWJlcikgcm93LmFwcGVuZENoaWxkKHRleHRFbCgnc3BhbicsIGBSb3cgJHtpc3N1ZS5yb3dOdW1iZXJ9YCwgJ3Jvdy1ubycpKTsKICAgICAgICByb3cuYXBwZW5kQ2hpbGQoZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUoaXNzdWUubWVzc2FnZSkpOwogICAgICAgIGlzc3Vlc0NhcmQuYXBwZW5kQ2hpbGQocm93KTsKICAgICAgfSk7CiAgICAgIGFyZWEuYXBwZW5kKGlzc3Vlc1RpdGxlLCBpc3N1ZXNDYXJkKTsKICAgIH0KCiAgICBpZiAocC5uZXdSZWNvcmRzLmxlbmd0aCkgewogICAgICBjb25zdCBuZXdUaXRsZSA9IHRleHRFbCgnZGl2JywgYE5ldyByZWNvcmRzIHRvIGltcG9ydCAoJHtwLm5ld1JlY29yZHMubGVuZ3RofSlgLCAnc2VjdGlvbi10aXRsZScpOwogICAgICBhcmVhLmFwcGVuZENoaWxkKG5ld1RpdGxlKTsKICAgICAgYXJlYS5hcHBlbmRDaGlsZCh0ZXh0RWwoCiAgICAgICAgJ2RpdicsCiAgICAgICAgJ1RoZXNlIHJvd3MgZG9u4oCZdCBtYXRjaCBhbnl0aGluZyBhbHJlYWR5IHNhdmVkLCBzbyB0aGV54oCZbGwgYmUgaW1wb3J0ZWQgYXV0b21hdGljYWxseSDigJQgbm8gZGVjaXNpb24gbmVlZGVkLCB1bmxpa2UgdGhlIGV4YWN0LWR1cGxpY2F0ZSBtYXRjaGVzIGFib3ZlLicsCiAgICAgICAgJ211dGVkJwogICAgICApKTsKICAgICAgY29uc3QgdGFibGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0YWJsZScpOwogICAgICB0YWJsZS5jbGFzc05hbWUgPSAnZGF0YS10YWJsZSc7CiAgICAgIHRhYmxlLmlubmVySFRNTCA9ICc8dGhlYWQ+PHRyPjx0aD5EYXRlPC90aD48dGg+Q2FwdGlvbjwvdGg+PHRoPlR5cGU8L3RoPjx0aD5DYW1wYWlnbjwvdGg+PHRoPlBsYXRmb3JtczwvdGg+PC90cj48L3RoZWFkPic7CiAgICAgIGNvbnN0IHRib2R5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGJvZHknKTsKICAgICAgcC5uZXdSZWNvcmRzLmZvckVhY2goKHMpID0+IHsKICAgICAgICBjb25zdCB0ciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAgICAgdHIuYXBwZW5kKAogICAgICAgICAgdGV4dEVsKCd0ZCcsIEZvcm1hdC5kYXRlKHMucHVibGlzaERhdGUpKSwKICAgICAgICAgIHRleHRFbCgndGQnLCBzLmNhcHRpb24gfHwgJ+KAlCcpLAogICAgICAgICAgdGV4dEVsKCd0ZCcsIHMuY29udGVudFR5cGUgfHwgJ+KAlCcpLAogICAgICAgICAgdGV4dEVsKCd0ZCcsIHMuY2FtcGFpZ25UeXBlIHx8ICdVbnNwZWNpZmllZCcpLAogICAgICAgICAgdGV4dEVsKCd0ZCcsIHMucGxhdGZvcm1zLmpvaW4oJywgJykpCiAgICAgICAgKTsKICAgICAgICB0Ym9keS5hcHBlbmRDaGlsZCh0cik7CiAgICAgIH0pOwogICAgICB0YWJsZS5hcHBlbmRDaGlsZCh0Ym9keSk7CiAgICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgd3JhcC5jbGFzc05hbWUgPSAndGFibGUtc2Nyb2xsJzsKICAgICAgd3JhcC5hcHBlbmRDaGlsZCh0YWJsZSk7CiAgICAgIGFyZWEuYXBwZW5kKHdyYXApOwogICAgfQoKICAgIGNvbnN0IGFjdGlvbnMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGFjdGlvbnMuY2xhc3NOYW1lID0gJ2J0bi1yb3cnOwogICAgYWN0aW9ucy5zdHlsZS5tYXJnaW5Ub3AgPSAnMTZweCc7CiAgICBjb25zdCBjb21taXRCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIGNvbW1pdEJ0bi5jbGFzc05hbWUgPSAnYnRuIHByaW1hcnknOwogICAgY29tbWl0QnRuLnRleHRDb250ZW50ID0gcC52YWxpZFJvd3MgPiAwID8gYEltcG9ydCAke3AudmFsaWRSb3dzfSByb3cocylgIDogJ05vdGhpbmcgdG8gaW1wb3J0JzsKICAgIGNvbW1pdEJ0bi5kaXNhYmxlZCA9IHAudmFsaWRSb3dzID09PSAwOwogICAgY29tbWl0QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gY29tbWl0KGNvbW1pdEJ0bikpOwogICAgY29uc3QgY2FuY2VsQnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBjYW5jZWxCdG4uY2xhc3NOYW1lID0gJ2J0bic7CiAgICBjYW5jZWxCdG4udGV4dENvbnRlbnQgPSAnQ2FuY2VsJzsKICAgIGNhbmNlbEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsgY3VycmVudFByZXZpZXcgPSBudWxsOyBzaGVsbCgpOyB9KTsKICAgIGFjdGlvbnMuYXBwZW5kKGNvbW1pdEJ0biwgY2FuY2VsQnRuKTsKICAgIGFyZWEuYXBwZW5kQ2hpbGQoYWN0aW9ucyk7CiAgfQoKICBmdW5jdGlvbiBzdGF0VGlsZShsYWJlbCwgdmFsdWUpIHsKICAgIGNvbnN0IHRpbGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRpbGUuY2xhc3NOYW1lID0gJ3N0YXQtdGlsZSc7CiAgICB0aWxlLmFwcGVuZCh0ZXh0RWwoJ2RpdicsIGxhYmVsLCAnc3RhdC1sYWJlbCcpLCB0ZXh0RWwoJ2RpdicsIFN0cmluZyh2YWx1ZSksICdzdGF0LXZhbHVlJykpOwogICAgcmV0dXJuIHRpbGU7CiAgfQogIGZ1bmN0aW9uIGFjdGlvblNlbGVjdChkZWZhdWx0VmFsKSB7CiAgICBjb25zdCBzZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzZWxlY3QnKTsKICAgIFtbJ3NraXAnLCAnU2tpcCAoa2VlcCBleGlzdGluZyByZWNvcmQgdW5jaGFuZ2VkKSddLCBbJ3VwZGF0ZScsICdVcGRhdGUgZXhpc3RpbmcgcmVjb3JkJ10sIFsnY3JlYXRlJywgJ0NyZWF0ZSBhcyBhIG5ldywgc2VwYXJhdGUgcmVjb3JkJ11dLmZvckVhY2goKFt2LCBsXSkgPT4gewogICAgICBjb25zdCBvcHQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsgb3B0LnZhbHVlID0gdjsgb3B0LnRleHRDb250ZW50ID0gbDsKICAgICAgaWYgKHYgPT09IGRlZmF1bHRWYWwpIG9wdC5zZWxlY3RlZCA9IHRydWU7CiAgICAgIHNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgICByZXR1cm4gc2VsOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gY29tbWl0KGJ0bikgewogICAgYnRuLmRpc2FibGVkID0gdHJ1ZTsKICAgIGJ0bi50ZXh0Q29udGVudCA9ICdJbXBvcnRpbmfigKYnOwogICAgY29uc3QgZGVmYXVsdER1cGxpY2F0ZUFjdGlvbiA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdkZWZhdWx0RHVwbGljYXRlQWN0aW9uU2VsZWN0Jyk/LnZhbHVlIHx8ICdza2lwJzsKICAgIGNvbnN0IG5vdGVzID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3VwbG9hZE5vdGVzSW5wdXQnKT8udmFsdWUgfHwgbnVsbDsKICAgIHRyeSB7CiAgICAgIGNvbnN0IHJlc3VsdCA9IGF3YWl0IEFwaS5jb21taXRVcGxvYWQoewogICAgICAgIGZpbGVQYXRoOiBjdXJyZW50UHJldmlldy5maWxlUGF0aCwKICAgICAgICBvcmlnaW5hbE5hbWU6IGN1cnJlbnRQcmV2aWV3Lm9yaWdpbmFsTmFtZSwKICAgICAgICBkZWZhdWx0RHVwbGljYXRlQWN0aW9uLAogICAgICAgIGR1cGxpY2F0ZUFjdGlvbnM6IGR1cGxpY2F0ZUFjdGlvbk92ZXJyaWRlcywKICAgICAgICBub3RlcywKICAgICAgfSk7CiAgICAgIFRvYXN0LnNob3coCiAgICAgICAgYEltcG9ydGVkOiAke3Jlc3VsdC5pbXBvcnRlZFJvd3N9IG5ldywgJHtyZXN1bHQudXBkYXRlZFJvd3N9IHVwZGF0ZWQsICR7cmVzdWx0LnNraXBwZWRSb3dzfSBza2lwcGVkLmAsCiAgICAgICAgcmVzdWx0LmVycm9yQ291bnQgPiAwID8gJ2Vycm9yJyA6ICdzdWNjZXNzJwogICAgICApOwogICAgICBjdXJyZW50UHJldmlldyA9IG51bGw7CiAgICAgIHNoZWxsKCk7CiAgICAgIHdpbmRvdy5kaXNwYXRjaEV2ZW50KG5ldyBDdXN0b21FdmVudCgnbHJzOmRhdGEtdXBkYXRlZCcpKTsKICAgIH0gY2F0Y2ggKGVycikgewogICAgICBUb2FzdC5zaG93KGVyci5tZXNzYWdlLCAnZXJyb3InKTsKICAgICAgYnRuLmRpc2FibGVkID0gZmFsc2U7CiAgICAgIGJ0bi50ZXh0Q29udGVudCA9ICdSZXRyeSBpbXBvcnQnOwogICAgfQogIH0KCiAgZnVuY3Rpb24gcmVuZGVyKCkgewogICAgcm9vdCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd2aWV3LXVwbG9hZCcpOwogICAgc2hlbGwoKTsKICB9CgogIHJldHVybiB7IHJlbmRlciB9Owp9KSgpOwoKY29uc3QgSGlzdG9yeSA9ICgoKSA9PiB7CiAgbGV0IHJvb3Q7CiAgbGV0IGN1cnJlbnRVcGxvYWRzID0gW107CiAgbGV0IHNlYXJjaFZhbHVlID0gJyc7CiAgbGV0IHBhZ2UgPSAxOwogIGNvbnN0IHBhZ2VTaXplID0gMTU7CiAgbGV0IHNvcnRTdGF0ZSA9IHsga2V5OiAndXBsb2FkZWRfYXQnLCBkaXI6ICdkZXNjJywgdHlwZTogJ3N0cmluZycgfTsKICBjb25zdCBFWFBPUlRfQ09MVU1OUyA9IFsKICAgIHsga2V5OiAnZmlsZW5hbWUnLCBsYWJlbDogJ0ZpbGUnIH0sCiAgICB7IGtleTogJ3VwbG9hZGVkX2F0JywgbGFiZWw6ICdVcGxvYWRlZCcgfSwKICAgIHsga2V5OiAnc3RhdHVzJywgbGFiZWw6ICdTdGF0dXMnIH0sCiAgICB7IGtleTogJ2ltcG9ydGVkX3Jvd3MnLCBsYWJlbDogJ0ltcG9ydGVkJyB9LAogICAgeyBrZXk6ICd1cGRhdGVkX3Jvd3MnLCBsYWJlbDogJ1VwZGF0ZWQnIH0sCiAgICB7IGtleTogJ3NraXBwZWRfcm93cycsIGxhYmVsOiAnU2tpcHBlZCcgfSwKICAgIHsga2V5OiAnZXJyb3JfY291bnQnLCBsYWJlbDogJ0Vycm9ycycgfSwKICAgIHsga2V5OiAnd2Vla3MnLCBsYWJlbDogJ1dlZWtzJyB9LAogICAgeyBrZXk6ICdub3RlcycsIGxhYmVsOiAnTm90ZXMnIH0sCiAgXTsKCiAgZnVuY3Rpb24gYmFkZ2VDbGFzcyhzdGF0dXMpIHsKICAgIGlmIChzdGF0dXMgPT09ICdzdWNjZXNzJykgcmV0dXJuICdzdWNjZXNzJzsKICAgIGlmIChzdGF0dXMgPT09ICdwYXJ0aWFsJykgcmV0dXJuICdwYXJ0aWFsJzsKICAgIHJldHVybiAnZmFpbGVkJzsKICB9CgogIGZ1bmN0aW9uIHNvcnRhYmxlSGVhZGVyKGxhYmVsLCBrZXksIHR5cGUpIHsKICAgIGNvbnN0IHRoID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGgnKTsKICAgIGlmICh0eXBlID09PSAnbnVtYmVyJykgdGguY2xhc3NOYW1lID0gJ251bSc7CiAgICB0aC5jbGFzc0xpc3QuYWRkKCdzb3J0YWJsZS10aCcpOwogICAgY29uc3QgaXNBY3RpdmUgPSBzb3J0U3RhdGUua2V5ID09PSBrZXk7CiAgICB0aC5hcHBlbmRDaGlsZChkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShsYWJlbCkpOwogICAgdGguYXBwZW5kQ2hpbGQodGV4dEVsKCdzcGFuJywgaXNBY3RpdmUgPyAoc29ydFN0YXRlLmRpciA9PT0gJ2FzYycgPyAnIOKGkScgOiAnIOKGkycpIDogJyDihpUnLCAnc29ydC1hcnJvdycpKTsKICAgIHRoLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gewogICAgICBzb3J0U3RhdGUgPSB7IGtleSwgZGlyOiBzb3J0U3RhdGUua2V5ID09PSBrZXkgJiYgc29ydFN0YXRlLmRpciA9PT0gJ2FzYycgPyAnZGVzYycgOiAnYXNjJywgdHlwZSB9OwogICAgICByZW5kZXJMaXN0KCk7CiAgICB9KTsKICAgIHJldHVybiB0aDsKICB9CgogIGZ1bmN0aW9uIGZpbHRlcmVkVXBsb2FkcygpIHsKICAgIGNvbnN0IHEgPSBzZWFyY2hWYWx1ZS50cmltKCkudG9Mb3dlckNhc2UoKTsKICAgIGlmICghcSkgcmV0dXJuIGN1cnJlbnRVcGxvYWRzOwogICAgcmV0dXJuIGN1cnJlbnRVcGxvYWRzLmZpbHRlcigodSkgPT4gKAogICAgICB1LmZpbGVuYW1lLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkKICAgICAgfHwgKHUubm90ZXMgfHwgJycpLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkKICAgICAgfHwgdS5zdGF0dXMudG9Mb3dlckNhc2UoKS5pbmNsdWRlcyhxKQogICAgKSk7CiAgfQoKICBmdW5jdGlvbiBzb3J0ZWRVcGxvYWRzKCkgewogICAgY29uc3QgeyBrZXksIGRpciwgdHlwZSB9ID0gc29ydFN0YXRlOwogICAgY29uc3QgZmFjdG9yID0gZGlyID09PSAnYXNjJyA/IDEgOiAtMTsKICAgIHJldHVybiBbLi4uZmlsdGVyZWRVcGxvYWRzKCldLnNvcnQoKGEsIGIpID0+IHsKICAgICAgY29uc3QgYXYgPSBhW2tleV07CiAgICAgIGNvbnN0IGJ2ID0gYltrZXldOwogICAgICBpZiAoYXYgPT09IG51bGwgfHwgYXYgPT09IHVuZGVmaW5lZCkgcmV0dXJuIDE7CiAgICAgIGlmIChidiA9PT0gbnVsbCB8fCBidiA9PT0gdW5kZWZpbmVkKSByZXR1cm4gLTE7CiAgICAgIGlmICh0eXBlID09PSAnbnVtYmVyJykgcmV0dXJuIChhdiAtIGJ2KSAqIGZhY3RvcjsKICAgICAgcmV0dXJuIFN0cmluZyhhdikubG9jYWxlQ29tcGFyZShTdHJpbmcoYnYpKSAqIGZhY3RvcjsKICAgIH0pOwogIH0KCiAgZnVuY3Rpb24gZXhwb3J0Um93cygpIHsKICAgIHJldHVybiBzb3J0ZWRVcGxvYWRzKCkubWFwKCh1KSA9PiAoewogICAgICBmaWxlbmFtZTogdS5maWxlbmFtZSwKICAgICAgdXBsb2FkZWRfYXQ6IHUudXBsb2FkZWRfYXQsCiAgICAgIHN0YXR1czogdS5zdGF0dXMsCiAgICAgIGltcG9ydGVkX3Jvd3M6IHUuaW1wb3J0ZWRfcm93cywKICAgICAgdXBkYXRlZF9yb3dzOiB1LnVwZGF0ZWRfcm93cywKICAgICAgc2tpcHBlZF9yb3dzOiB1LnNraXBwZWRfcm93cywKICAgICAgZXJyb3JfY291bnQ6IHUuZXJyb3JfY291bnQsCiAgICAgIHdlZWtzOiB1LndlZWtzX2FmZmVjdGVkLm1hcCgodykgPT4gRm9ybWF0LmRhdGUodykpLmpvaW4oJywgJyksCiAgICAgIG5vdGVzOiB1Lm5vdGVzIHx8ICcnLAogICAgfSkpOwogIH0KCiAgZnVuY3Rpb24gYnVpbGRCYWNrdXBDYXJkKCkgewogICAgY29uc3QgY2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY2FyZC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICBjYXJkLnN0eWxlLm1hcmdpbkJvdHRvbSA9ICcyMHB4JzsKICAgIGNvbnN0IGhlYWRlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgaGVhZGVyLmNsYXNzTmFtZSA9ICdjYXJkLWhlYWRlcic7CiAgICBoZWFkZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMycsICdCYWNrdXAgJiBSZXN0b3JlJykpOwogICAgY2FyZC5hcHBlbmRDaGlsZChoZWFkZXIpOwogICAgY2FyZC5hcHBlbmRDaGlsZCh0ZXh0RWwoCiAgICAgICdkaXYnLAogICAgICAnRG93bmxvYWQgYSBmdWxsIHNuYXBzaG90IG9mIHRoZSBkYXRhYmFzZSBhbnkgdGltZS4gUmVzdG9yaW5nIHJlcGxhY2VzIEFMTCBjdXJyZW50IGRhdGEgd2l0aCB0aGUgdXBsb2FkZWQgYmFja3VwIGFuZCByZXN0YXJ0cyB0aGUgc2VydmVyLiAiRGVsZXRlIGFsbCB1cGxvYWRlZCByZWNvcmRzIiBjbGVhcnMgZXZlcnkgcmVjb3JkIGFuZCB0aGUgdXBsb2FkIGhpc3RvcnkgKGJ1dCBub3QgRm9sbG93ZXJzIERhdGEgUmVjb3JkKSBzbyB5b3UgY2FuIHJlYnVpbGQgZnJvbSBhIGZyZXNoIGltcG9ydC4gQm90aCBhY3Rpb25zIGNhbm5vdCBiZSB1bmRvbmUg4oCUIGRvd25sb2FkIGEgYmFja3VwIGZpcnN0LicsCiAgICAgICdtdXRlZCcKICAgICkpOwoKICAgIGNvbnN0IGFjdGlvbnMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGFjdGlvbnMuY2xhc3NOYW1lID0gJ2J0bi1yb3cnOwogICAgYWN0aW9ucy5zdHlsZS5tYXJnaW5Ub3AgPSAnMTRweCc7CgogICAgY29uc3QgZG93bmxvYWRCdG4gPSBpY29uQnRuKCdidG4gcHJpbWFyeScsICdkb3dubG9hZCcsICdEb3dubG9hZCBCYWNrdXAnKTsKICAgIGRvd25sb2FkQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4geyB3aW5kb3cubG9jYXRpb24uaHJlZiA9ICcvYXBpL2JhY2t1cC9leHBvcnQnOyB9KTsKCiAgICBjb25zdCByZXN0b3JlSW5wdXQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOwogICAgcmVzdG9yZUlucHV0LnR5cGUgPSAnZmlsZSc7CiAgICByZXN0b3JlSW5wdXQuYWNjZXB0ID0gJy5kYic7CiAgICByZXN0b3JlSW5wdXQuc3R5bGUuZGlzcGxheSA9ICdub25lJzsKCiAgICBjb25zdCByZXN0b3JlQnRuID0gaWNvbkJ0bignYnRuIGRhbmdlcicsICd1cGxvYWQnLCAnUmVzdG9yZSBmcm9tIEJhY2t1cCcpOwogICAgcmVzdG9yZUJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHJlc3RvcmVJbnB1dC5jbGljaygpKTsKCiAgICByZXN0b3JlSW5wdXQuYWRkRXZlbnRMaXN0ZW5lcignY2hhbmdlJywgYXN5bmMgKCkgPT4gewogICAgICBjb25zdCBmaWxlID0gcmVzdG9yZUlucHV0LmZpbGVzWzBdOwogICAgICBpZiAoIWZpbGUpIHJldHVybjsKICAgICAgY29uc3Qgc3VyZSA9IHdpbmRvdy5jb25maXJtKAogICAgICAgICdSZXN0b3Jpbmcgd2lsbCBSRVBMQUNFIGFsbCBjdXJyZW50IGRhdGEgd2l0aCB0aGlzIGJhY2t1cCBmaWxlIGFuZCByZXN0YXJ0IHRoZSBzZXJ2ZXIuIFRoaXMgY2Fubm90IGJlIHVuZG9uZS4gQ29udGludWU/JwogICAgICApOwogICAgICBpZiAoIXN1cmUpIHsKICAgICAgICByZXN0b3JlSW5wdXQudmFsdWUgPSAnJzsKICAgICAgICByZXR1cm47CiAgICAgIH0KICAgICAgcmVzdG9yZUJ0bi5kaXNhYmxlZCA9IHRydWU7CiAgICAgIHRyeSB7CiAgICAgICAgY29uc3QgZm9ybSA9IG5ldyBGb3JtRGF0YSgpOwogICAgICAgIGZvcm0uYXBwZW5kKCdmaWxlJywgZmlsZSk7CiAgICAgICAgY29uc3QgcmVzdWx0ID0gYXdhaXQgQXBpLnJlc3RvcmVCYWNrdXAoZm9ybSk7CiAgICAgICAgVG9hc3Quc2hvdyhyZXN1bHQubWVzc2FnZSB8fCAnQmFja3VwIHJlc3RvcmVkLiBUaGUgc2VydmVyIGlzIHJlc3RhcnRpbmcuJywgJ3N1Y2Nlc3MnKTsKICAgICAgfSBjYXRjaCAoZXJyKSB7CiAgICAgICAgVG9hc3Quc2hvdyhlcnIubWVzc2FnZSwgJ2Vycm9yJyk7CiAgICAgICAgcmVzdG9yZUJ0bi5kaXNhYmxlZCA9IGZhbHNlOwogICAgICB9IGZpbmFsbHkgewogICAgICAgIHJlc3RvcmVJbnB1dC52YWx1ZSA9ICcnOwogICAgICB9CiAgICB9KTsKCiAgICBjb25zdCB3aXBlQnRuID0gaWNvbkJ0bignYnRuIGRhbmdlcicsICd0cmFzaC0yJywgJ0RlbGV0ZSBhbGwgdXBsb2FkZWQgcmVjb3JkcycpOwogICAgd2lwZUJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIGFzeW5jICgpID0+IHsKICAgICAgY29uc3QgdHlwZWQgPSB3aW5kb3cucHJvbXB0KAogICAgICAgICdUaGlzIHBlcm1hbmVudGx5IGRlbGV0ZXMgRVZFUlkgdXBsb2FkZWQgcmVjb3JkIGFuZCBhbGwgdXBsb2FkIGhpc3RvcnksIHNvIHlvdSBjYW4gcmVidWlsZCB0aGUgZGF0YSBmcm9tIGEgZnJlc2ggaW1wb3J0LiBGb2xsb3dlcnMgRGF0YSBSZWNvcmQgaXMgbm90IGFmZmVjdGVkLiBUaGlzIGNhbm5vdCBiZSB1bmRvbmUuXG5cblR5cGUgREVMRVRFIHRvIGNvbmZpcm06JwogICAgICApOwogICAgICBpZiAodHlwZWQgIT09ICdERUxFVEUnKSB7CiAgICAgICAgaWYgKHR5cGVkICE9PSBudWxsKSBUb2FzdC5zaG93KCdDb25maXJtYXRpb24gcGhyYXNlIGRpZCBub3QgbWF0Y2gg4oCUIG5vdGhpbmcgd2FzIGRlbGV0ZWQuJywgJ2Vycm9yJyk7CiAgICAgICAgcmV0dXJuOwogICAgICB9CiAgICAgIHdpcGVCdG4uZGlzYWJsZWQgPSB0cnVlOwogICAgICB0cnkgewogICAgICAgIGNvbnN0IHJlc3VsdCA9IGF3YWl0IEFwaS53aXBlVXBsb2FkZWRSZWNvcmRzKCk7CiAgICAgICAgY29uc3QgbiA9IHJlc3VsdC5kZWxldGVkID8gcmVzdWx0LmRlbGV0ZWQucG9zdHMgOiAwOwogICAgICAgIFRvYXN0LnNob3coYERlbGV0ZWQgJHtufSByZWNvcmQocykuIFVwbG9hZCBhIGZpbGUgdG8gbG9hZCBmcmVzaCBkYXRhLmAsICdzdWNjZXNzJyk7CiAgICAgICAgd2luZG93LmRpc3BhdGNoRXZlbnQobmV3IEN1c3RvbUV2ZW50KCdscnM6ZGF0YS11cGRhdGVkJykpOwogICAgICAgIHJlbmRlcigpOwogICAgICB9IGNhdGNoIChlcnIpIHsKICAgICAgICBUb2FzdC5zaG93KGVyci5tZXNzYWdlIHx8ICdEZWxldGUgZmFpbGVkLicsICdlcnJvcicpOwogICAgICB9IGZpbmFsbHkgewogICAgICAgIHdpcGVCdG4uZGlzYWJsZWQgPSBmYWxzZTsKICAgICAgfQogICAgfSk7CgogICAgYWN0aW9ucy5hcHBlbmQoZG93bmxvYWRCdG4sIHJlc3RvcmVCdG4sIHJlc3RvcmVJbnB1dCwgd2lwZUJ0bik7CiAgICBjYXJkLmFwcGVuZENoaWxkKGFjdGlvbnMpOwogICAgcmV0dXJuIGNhcmQ7CiAgfQoKICBhc3luYyBmdW5jdGlvbiByZW5kZXIoKSB7CiAgICByb290ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3ZpZXctaGlzdG9yeScpOwogICAgcm9vdC5pbm5lckhUTUwgPSAnJzsKICAgIHJvb3QuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCAnVXBsb2FkIGhpc3RvcnknLCAnc2VjdGlvbi10aXRsZScpKTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQoYnVpbGRCYWNrdXBDYXJkKCkpOwoKICAgIGN1cnJlbnRVcGxvYWRzID0gYXdhaXQgQXBpLnVwbG9hZEhpc3RvcnkoKTsKICAgIGlmICghY3VycmVudFVwbG9hZHMubGVuZ3RoKSB7CiAgICAgIHJvb3QuYXBwZW5kQ2hpbGQoZW1wdHlTdGF0ZSh7CiAgICAgICAgaWNvbjogJ3VwbG9hZC1jbG91ZCcsCiAgICAgICAgdGl0bGU6ICdObyB1cGxvYWRzIHlldCcsCiAgICAgICAgbWVzc2FnZTogJ0ltcG9ydCB5b3VyIGZpcnN0IHdlZWtseSBleHBvcnQgdG8gc3RhcnQgc2VlaW5nIGRhdGEgYWNyb3NzIHRoZSBhcHAuJywKICAgICAgICBhY3Rpb25MYWJlbDogJ1VwbG9hZCBkYXRhJywKICAgICAgICBvbkFjdGlvbjogKCkgPT4gZG9jdW1lbnQucXVlcnlTZWxlY3RvcignLnRhYi1idG5bZGF0YS10YWI9InVwbG9hZCJdJyk/LmNsaWNrKCksCiAgICAgIH0pKTsKICAgICAgcmV0dXJuOwogICAgfQoKICAgIGNvbnN0IHRvb2xiYXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRvb2xiYXIuY2xhc3NOYW1lID0gJ3JlY29yZHMtdG9vbGJhcic7CiAgICB0b29sYmFyLmFwcGVuZENoaWxkKGJ1aWxkU2VhcmNoQm94KHsKICAgICAgcGxhY2Vob2xkZXI6ICdTZWFyY2ggZmlsZW5hbWUsIHN0YXR1cywgb3Igbm90ZXPigKYnLAogICAgICB2YWx1ZTogc2VhcmNoVmFsdWUsCiAgICAgIG9uQ2hhbmdlOiAodikgPT4geyBzZWFyY2hWYWx1ZSA9IHY7IHBhZ2UgPSAxOyByZW5kZXJMaXN0KCk7IH0sCiAgICB9KSk7CiAgICByb290LmFwcGVuZENoaWxkKHRvb2xiYXIpOwoKICAgIHJvb3QuYXBwZW5kQ2hpbGQoYnVpbGRFeHBvcnRCdXR0b25zKHsKICAgICAgZ2V0Um93c0FuZENvbHVtbnM6ICgpID0+ICh7IHJvd3M6IGV4cG9ydFJvd3MoKSwgY29sdW1uczogRVhQT1JUX0NPTFVNTlMgfSksCiAgICAgIGZpbGVuYW1lQmFzZTogJ3VwbG9hZC1oaXN0b3J5JywKICAgICAgc2hlZXROYW1lOiAnVXBsb2FkIEhpc3RvcnknLAogICAgfSkpOwoKICAgIGNvbnN0IGNhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgd3JhcC5jbGFzc05hbWUgPSAndGFibGUtc2Nyb2xsJzsKICAgIHdyYXAuaWQgPSAnaGlzdG9yeVRhYmxlV3JhcCc7CiAgICBjYXJkLmFwcGVuZENoaWxkKHdyYXApOwogICAgY29uc3QgcGFnZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHBhZ2VyLmNsYXNzTmFtZSA9ICdwYWdpbmF0aW9uLXJvdyc7CiAgICBwYWdlci5pZCA9ICdoaXN0b3J5UGFnZXInOwogICAgY2FyZC5hcHBlbmRDaGlsZChwYWdlcik7CiAgICByb290LmFwcGVuZENoaWxkKGNhcmQpOwoKICAgIHJlbmRlckxpc3QoKTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlckxpc3QoKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2hpc3RvcnlUYWJsZVdyYXAnKTsKICAgIGNvbnN0IHBhZ2VyRWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnaGlzdG9yeVBhZ2VyJyk7CiAgICBpZiAoIXdyYXApIHJldHVybjsKICAgIGNvbnN0IGFsbFNvcnRlZCA9IHNvcnRlZFVwbG9hZHMoKTsKICAgIGlmICghYWxsU29ydGVkLmxlbmd0aCkgewogICAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgICB3cmFwLmFwcGVuZENoaWxkKGVtcHR5U3RhdGUoeyBpY29uOiAndXBsb2FkLWNsb3VkJywgbWVzc2FnZTogJ05vIHVwbG9hZHMgbWF0Y2ggeW91ciBzZWFyY2guJyB9KSk7CiAgICAgIGlmIChwYWdlckVsKSBwYWdlckVsLmlubmVySFRNTCA9ICcnOwogICAgICByZXR1cm47CiAgICB9CiAgICBjb25zdCB7IHBhZ2VSb3dzLCB0b3RhbFBhZ2VzLCBzYWZlUGFnZSwgdG90YWwgfSA9IHBhZ2luYXRlQ2xpZW50U2lkZShhbGxTb3J0ZWQsIHBhZ2UsIHBhZ2VTaXplKTsKICAgIHBhZ2UgPSBzYWZlUGFnZTsKCiAgICBjb25zdCB0YWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICB0YWJsZS5jbGFzc05hbWUgPSAnZGF0YS10YWJsZSc7CiAgICBjb25zdCB0aGVhZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RoZWFkJyk7CiAgICBjb25zdCBoZWFkVHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgaGVhZFRyLmFwcGVuZCgKICAgICAgc29ydGFibGVIZWFkZXIoJ0ZpbGUnLCAnZmlsZW5hbWUnLCAnc3RyaW5nJyksCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdVcGxvYWRlZCcsICd1cGxvYWRlZF9hdCcsICdzdHJpbmcnKSwKICAgICAgc29ydGFibGVIZWFkZXIoJ1N0YXR1cycsICdzdGF0dXMnLCAnc3RyaW5nJyksCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdJbXBvcnRlZCcsICdpbXBvcnRlZF9yb3dzJywgJ251bWJlcicpLAogICAgICBzb3J0YWJsZUhlYWRlcignVXBkYXRlZCcsICd1cGRhdGVkX3Jvd3MnLCAnbnVtYmVyJyksCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdTa2lwcGVkJywgJ3NraXBwZWRfcm93cycsICdudW1iZXInKSwKICAgICAgc29ydGFibGVIZWFkZXIoJ0Vycm9ycycsICdlcnJvcl9jb3VudCcsICdudW1iZXInKSwKICAgICAgdGV4dEVsKCd0aCcsICdXZWVrcycpLAogICAgICB0ZXh0RWwoJ3RoJywgJ05vdGVzJykKICAgICk7CiAgICB0aGVhZC5hcHBlbmRDaGlsZChoZWFkVHIpOwogICAgY29uc3QgdGJvZHkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0Ym9keScpOwogICAgcGFnZVJvd3MuZm9yRWFjaCgodSkgPT4gewogICAgICBjb25zdCB0ciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAgIHRyLnN0eWxlLmN1cnNvciA9ICdwb2ludGVyJzsKICAgICAgY29uc3QgYmFkZ2UgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICAgIGJhZGdlLmNsYXNzTmFtZSA9IGBiYWRnZSAke2JhZGdlQ2xhc3ModS5zdGF0dXMpfWA7CiAgICAgIGJhZGdlLnRleHRDb250ZW50ID0gdS5zdGF0dXM7CiAgICAgIGNvbnN0IHN0YXR1c1RkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgc3RhdHVzVGQuYXBwZW5kQ2hpbGQoYmFkZ2UpOwogICAgICB0ci5hcHBlbmQoCiAgICAgICAgdGV4dEVsKCd0ZCcsIHUuZmlsZW5hbWUpLAogICAgICAgIHRleHRFbCgndGQnLCB1LnVwbG9hZGVkX2F0KSwKICAgICAgICBzdGF0dXNUZCwKICAgICAgICB0ZXh0RWwoJ3RkJywgU3RyaW5nKHUuaW1wb3J0ZWRfcm93cyksICdudW0nKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgU3RyaW5nKHUudXBkYXRlZF9yb3dzKSwgJ251bScpLAogICAgICAgIHRleHRFbCgndGQnLCBTdHJpbmcodS5za2lwcGVkX3Jvd3MpLCAnbnVtJyksCiAgICAgICAgdGV4dEVsKCd0ZCcsIFN0cmluZyh1LmVycm9yX2NvdW50KSwgJ251bScpLAogICAgICAgIHRleHRFbCgndGQnLCB1LndlZWtzX2FmZmVjdGVkLm1hcCgodykgPT4gRm9ybWF0LmRhdGUodykpLmpvaW4oJywgJykgfHwgJ+KAlCcpLAogICAgICAgIHRleHRFbCgndGQnLCB1Lm5vdGVzIHx8ICfigJQnKQogICAgICApOwogICAgICB0ci5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHRvZ2dsZUVycm9ycyh1LmlkLCB0cikpOwogICAgICB0Ym9keS5hcHBlbmRDaGlsZCh0cik7CiAgICB9KTsKICAgIHRhYmxlLmFwcGVuZCh0aGVhZCwgdGJvZHkpOwogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGFibGUpOwoKICAgIGlmIChwYWdlckVsKSB7CiAgICAgIHBhZ2VyRWwuaW5uZXJIVE1MID0gJyc7CiAgICAgIHBhZ2VyRWwuYXBwZW5kQ2hpbGQoYnVpbGRQYWdlcih7CiAgICAgICAgcGFnZTogc2FmZVBhZ2UsCiAgICAgICAgdG90YWxQYWdlcywKICAgICAgICB0b3RhbCwKICAgICAgICBvblByZXY6ICgpID0+IHsgcGFnZSAtPSAxOyByZW5kZXJMaXN0KCk7IH0sCiAgICAgICAgb25OZXh0OiAoKSA9PiB7IHBhZ2UgKz0gMTsgcmVuZGVyTGlzdCgpOyB9LAogICAgICB9KSk7CiAgICB9CiAgfQoKICBhc3luYyBmdW5jdGlvbiB0b2dnbGVFcnJvcnModXBsb2FkSWQsIHRyKSB7CiAgICBjb25zdCBleGlzdGluZyA9IHRyLm5leHRFbGVtZW50U2libGluZzsKICAgIGlmIChleGlzdGluZyAmJiBleGlzdGluZy5jbGFzc0xpc3QuY29udGFpbnMoJ2Vycm9yLWxvZy1yb3cnKSkgewogICAgICBleGlzdGluZy5yZW1vdmUoKTsKICAgICAgcmV0dXJuOwogICAgfQogICAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnLmVycm9yLWxvZy1yb3cnKS5mb3JFYWNoKChlbCkgPT4gZWwucmVtb3ZlKCkpOwogICAgY29uc3QgZXJyb3JzID0gYXdhaXQgQXBpLnVwbG9hZEVycm9ycyh1cGxvYWRJZCk7CiAgICBjb25zdCBsb2dSb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgbG9nUm93LmNsYXNzTmFtZSA9ICdlcnJvci1sb2ctcm93JzsKICAgIGNvbnN0IHRkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgIHRkLmNvbFNwYW4gPSA5OwogICAgaWYgKCFlcnJvcnMubGVuZ3RoKSB7CiAgICAgIHRkLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgJ05vIGlzc3VlcyBsb2dnZWQgZm9yIHRoaXMgdXBsb2FkLicsICdtdXRlZCcpKTsKICAgIH0gZWxzZSB7CiAgICAgIGNvbnN0IGxpc3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgbGlzdC5jbGFzc05hbWUgPSAnaXNzdWVzLWxpc3QnOwogICAgICBlcnJvcnMuZm9yRWFjaCgoZSkgPT4gewogICAgICAgIGNvbnN0IHJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICAgIHJvdy5jbGFzc05hbWUgPSAnaXNzdWUtcm93JzsKICAgICAgICBjb25zdCBiYWRnZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgICAgICBiYWRnZS5jbGFzc05hbWUgPSBgYmFkZ2UgJHtlLnNldmVyaXR5fS1zZXZgOwogICAgICAgIGJhZGdlLnRleHRDb250ZW50ID0gZS5zZXZlcml0eTsKICAgICAgICByb3cuYXBwZW5kKGJhZGdlLCBkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShgICR7ZS5yb3dfbnVtYmVyID8gYFJvdyAke2Uucm93X251bWJlcn06IGAgOiAnJ30ke2UubWVzc2FnZX1gKSk7CiAgICAgICAgbGlzdC5hcHBlbmRDaGlsZChyb3cpOwogICAgICB9KTsKICAgICAgdGQuYXBwZW5kQ2hpbGQobGlzdCk7CiAgICB9CgogICAgY29uc3QgcmF3QnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICByYXdCdG4uY2xhc3NOYW1lID0gJ2J0bic7CiAgICByYXdCdG4uc3R5bGUubWFyZ2luVG9wID0gJzEwcHgnOwogICAgcmF3QnRuLnRleHRDb250ZW50ID0gJ1ZpZXcgZXZlcnkgcmF3IHNvdXJjZSByb3cgZnJvbSB0aGlzIHVwbG9hZCc7CiAgICByYXdCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBsb2FkUmF3Um93cyh1cGxvYWRJZCwgcmF3QnRuKSk7CiAgICB0ZC5hcHBlbmRDaGlsZChyYXdCdG4pOwogICAgY29uc3QgcmF3V3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcmF3V3JhcC5pZCA9IGByYXdXcmFwLSR7dXBsb2FkSWR9YDsKICAgIHRkLmFwcGVuZENoaWxkKHJhd1dyYXApOwoKICAgIGxvZ1Jvdy5hcHBlbmRDaGlsZCh0ZCk7CiAgICB0ci5hZnRlcihsb2dSb3cpOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gbG9hZFJhd1Jvd3ModXBsb2FkSWQsIGJ0bikgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGByYXdXcmFwLSR7dXBsb2FkSWR9YCk7CiAgICBpZiAoIXdyYXApIHJldHVybjsKICAgIGlmICh3cmFwLmRhdGFzZXQubG9hZGVkKSB7CiAgICAgIHdyYXAuc3R5bGUuZGlzcGxheSA9IHdyYXAuc3R5bGUuZGlzcGxheSA9PT0gJ25vbmUnID8gJ2Jsb2NrJyA6ICdub25lJzsKICAgICAgcmV0dXJuOwogICAgfQogICAgYnRuLnRleHRDb250ZW50ID0gJ0xvYWRpbmfigKYnOwogICAgY29uc3QgeyByb3dzLCB0b3RhbCB9ID0gYXdhaXQgQXBpLnVwbG9hZFJhd1Jvd3ModXBsb2FkSWQpOwogICAgd3JhcC5kYXRhc2V0LmxvYWRlZCA9ICcxJzsKICAgIGJ0bi50ZXh0Q29udGVudCA9IGBTaG93aW5nICR7cm93cy5sZW5ndGh9IG9mICR7dG90YWx9IHJhdyByb3cocylgOwoKICAgIGNvbnN0IGJ5U2hlZXQgPSBuZXcgTWFwKCk7CiAgICByb3dzLmZvckVhY2goKHIpID0+IHsKICAgICAgaWYgKCFieVNoZWV0LmhhcyhyLnNoZWV0X25hbWUpKSBieVNoZWV0LnNldChyLnNoZWV0X25hbWUsIFtdKTsKICAgICAgYnlTaGVldC5nZXQoci5zaGVldF9uYW1lKS5wdXNoKHIpOwogICAgfSk7CgogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgIHdyYXAuc3R5bGUubWFyZ2luVG9wID0gJzEwcHgnOwogICAgYnlTaGVldC5mb3JFYWNoKChzaGVldFJvd3MsIHNoZWV0TmFtZSkgPT4gewogICAgICB3cmFwLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgYFNoZWV0OiAke3NoZWV0TmFtZX0gKCR7c2hlZXRSb3dzLmxlbmd0aH0gcm93KHMpKWAsICdzdGF0LWxhYmVsJykpOwogICAgICBjb25zdCBoZWFkZXJzID0gc2hlZXRSb3dzWzBdLmhlYWRlcnM7CiAgICAgIGNvbnN0IHRhYmxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGFibGUnKTsKICAgICAgdGFibGUuY2xhc3NOYW1lID0gJ2RhdGEtdGFibGUnOwogICAgICBjb25zdCB0aGVhZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAgIHRoZWFkLmFwcGVuZCh0ZXh0RWwoJ3RoJywgJ1JvdyAjJyksIHRleHRFbCgndGgnLCAnTGlua2VkIHRvIHBvc3QnKSk7CiAgICAgIGNvbnN0IGNvbENvdW50ID0gaGVhZGVycyA/IGhlYWRlcnMubGVuZ3RoIDogTWF0aC5tYXgoLi4uc2hlZXRSb3dzLm1hcCgocikgPT4gci5yYXcubGVuZ3RoKSk7CiAgICAgIGZvciAobGV0IGkgPSAwOyBpIDwgY29sQ291bnQ7IGkgKz0gMSkgdGhlYWQuYXBwZW5kQ2hpbGQodGV4dEVsKCd0aCcsIGhlYWRlcnMgJiYgaGVhZGVyc1tpXSA/IFN0cmluZyhoZWFkZXJzW2ldKSA6IGBDb2wgJHtpICsgMX1gKSk7CiAgICAgIGNvbnN0IHRoZWFkV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RoZWFkJyk7CiAgICAgIHRoZWFkV3JhcC5hcHBlbmRDaGlsZCh0aGVhZCk7CiAgICAgIGNvbnN0IHRib2R5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGJvZHknKTsKICAgICAgc2hlZXRSb3dzLmZvckVhY2goKHIpID0+IHsKICAgICAgICBjb25zdCB0cjIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgICAgIHRyMi5hcHBlbmQodGV4dEVsKCd0ZCcsIFN0cmluZyhyLnJvd19udW1iZXIpKSwgdGV4dEVsKCd0ZCcsIHIucG9zdF9pZCA/IGAjJHtyLnBvc3RfaWR9YCA6ICfigJQnKSk7CiAgICAgICAgZm9yIChsZXQgaSA9IDA7IGkgPCBjb2xDb3VudDsgaSArPSAxKSB7CiAgICAgICAgICBjb25zdCB2YWwgPSByLnJhd1tpXTsKICAgICAgICAgIHRyMi5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3RkJywgdmFsID09PSB1bmRlZmluZWQgfHwgdmFsID09PSBudWxsID8gJycgOiBTdHJpbmcodmFsKS5zbGljZSgwLCA2MCkpKTsKICAgICAgICB9CiAgICAgICAgdGJvZHkuYXBwZW5kQ2hpbGQodHIyKTsKICAgICAgfSk7CiAgICAgIHRhYmxlLmFwcGVuZCh0aGVhZFdyYXAsIHRib2R5KTsKICAgICAgY29uc3Qgc2Nyb2xsV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBzY3JvbGxXcmFwLmNsYXNzTmFtZSA9ICd0YWJsZS1zY3JvbGwnOwogICAgICBzY3JvbGxXcmFwLnN0eWxlLm1hcmdpbkJvdHRvbSA9ICcxNnB4JzsKICAgICAgc2Nyb2xsV3JhcC5hcHBlbmRDaGlsZCh0YWJsZSk7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQoc2Nyb2xsV3JhcCk7CiAgICB9KTsKICB9CgogIHJldHVybiB7IHJlbmRlciB9Owp9KSgpOwoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIEZvbGxvd2VycyBEYXRhIHRhYjogbWFudWFsIHdlZWtseSBmb2xsb3dlci1jb3VudCBlbnRyeSBwZXIKICAgcGxhdGZvcm0g4oCUIGVudGlyZWx5IGluZGVwZW5kZW50IG9mIHNwcmVhZHNoZWV0IHVwbG9hZHMgKGl0cyBvd24KICAgdGFibGUsIGl0cyBvd24gQVBJLCBuZXZlciB0b3VjaGVkIGJ5IHRoZSBpbXBvcnQgcGlwZWxpbmUpLiBQb3dlcnMKICAgRm9sbG93ZXIgR3Jvd3RoIGNoYXJ0cy9jb21wYXJpc29ucyBlbHNld2hlcmUgaW4gdGhlIGFwcC4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmNvbnN0IEZvbGxvd2VycyA9ICgoKSA9PiB7CiAgbGV0IHJvb3Q7CiAgbGV0IGVkaXRpbmdJZCA9IG51bGw7IC8vIG5vbi1udWxsIHdoaWxlIHRoZSBmb3JtIGlzIGVkaXRpbmcgYW4gZXhpc3RpbmcgZW50cnkgcmF0aGVyIHRoYW4gYWRkaW5nIGEgbmV3IG9uZQogIGxldCBzb3J0U3RhdGUgPSB7IGtleTogJ2VudHJ5X2RhdGUnLCBkaXI6ICdkZXNjJywgdHlwZTogJ3N0cmluZycgfTsKICBsZXQgY3VycmVudFJvd3MgPSBbXTsKICBsZXQgc2VhcmNoVmFsdWUgPSAnJzsKICBsZXQgcGFnZSA9IDE7CiAgY29uc3QgcGFnZVNpemUgPSAxMDsKICBjb25zdCBFWFBPUlRfQ09MVU1OUyA9IFsKICAgIHsga2V5OiAncGxhdGZvcm1fbGFiZWwnLCBsYWJlbDogJ1BsYXRmb3JtJyB9LAogICAgeyBrZXk6ICdlbnRyeV9kYXRlJywgbGFiZWw6ICdXZWVrIC8gRGF0ZScgfSwKICAgIHsga2V5OiAnZm9sbG93ZXJzX2NvdW50JywgbGFiZWw6ICdGb2xsb3dlcnMgQ291bnQnIH0sCiAgICB7IGtleTogJ3VwZGF0ZWRfYXQnLCBsYWJlbDogJ0xhc3QgVXBkYXRlZCcgfSwKICBdOwoKICBmdW5jdGlvbiBhbGxQbGF0Zm9ybXMoKSB7CiAgICByZXR1cm4gKHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZSB8fCB7IGFsbFBsYXRmb3JtczogW10gfSkuYWxsUGxhdGZvcm1zIHx8IFtdOwogIH0KCiAgZnVuY3Rpb24gcGxhdGZvcm1NZXRhRm9yKGlkKSB7CiAgICByZXR1cm4gYWxsUGxhdGZvcm1zKCkuZmluZCgocCkgPT4gcC5pZCA9PT0gaWQpIHx8IHsgbGFiZWw6IGlkLCBjb2xvcjogJyM5OTknIH07CiAgfQoKICBmdW5jdGlvbiBzaGVsbCgpIHsKICAgIHJvb3QuaW5uZXJIVE1MID0gJyc7CiAgICByb290LmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgJ0ZvbGxvd2VycyBEYXRhIFJlY29yZCcsICdzZWN0aW9uLXRpdGxlJykpOwogICAgcm9vdC5hcHBlbmRDaGlsZCh0ZXh0RWwoCiAgICAgICdkaXYnLAogICAgICAnTWFudWFsbHkgbG9nIGVhY2ggcGxhdGZvcm3igJlzIHRvdGFsIGZvbGxvd2VyIGNvdW50IG9uY2UgYSB3ZWVrLiBUaGlzIGlzIGluZGVwZW5kZW50IG9mIHNwcmVhZHNoZWV0IHVwbG9hZHMg4oCUIGl0IHBvd2VycyBGb2xsb3dlciBHcm93dGggY2hhcnRzIGFuZCBjb21wYXJpc29ucyBlbHNld2hlcmUgaW4gdGhlIGFwcC4nLAogICAgICAnbXV0ZWQnCiAgICApKTsKCiAgICBjb25zdCBmb3JtQ2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZm9ybUNhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgZm9ybUNhcmQuc3R5bGUubWFyZ2luQm90dG9tID0gJzIwcHgnOwogICAgZm9ybUNhcmQuaWQgPSAnZm9sbG93ZXJzRm9ybUNhcmQnOwogICAgcm9vdC5hcHBlbmRDaGlsZChmb3JtQ2FyZCk7CgogICAgY29uc3QgdG9vbGJhciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdG9vbGJhci5jbGFzc05hbWUgPSAncmVjb3Jkcy10b29sYmFyJzsKICAgIHRvb2xiYXIuYXBwZW5kQ2hpbGQoYnVpbGRTZWFyY2hCb3goewogICAgICBwbGFjZWhvbGRlcjogJ1NlYXJjaCBwbGF0Zm9ybSBvciBkYXRl4oCmJywKICAgICAgdmFsdWU6IHNlYXJjaFZhbHVlLAogICAgICBvbkNoYW5nZTogKHYpID0+IHsgc2VhcmNoVmFsdWUgPSB2OyBwYWdlID0gMTsgcmVuZGVyVGFibGUoKTsgfSwKICAgIH0pKTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQodG9vbGJhcik7CgogICAgcm9vdC5hcHBlbmRDaGlsZChidWlsZEV4cG9ydEJ1dHRvbnMoewogICAgICBnZXRSb3dzQW5kQ29sdW1uczogKCkgPT4gKHsgcm93czogZXhwb3J0Um93cygpLCBjb2x1bW5zOiBFWFBPUlRfQ09MVU1OUyB9KSwKICAgICAgZmlsZW5hbWVCYXNlOiAnZm9sbG93ZXJzLWRhdGEnLAogICAgICBzaGVldE5hbWU6ICdGb2xsb3dlcnMgRGF0YScsCiAgICB9KSk7CgogICAgY29uc3QgdGFibGVDYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0YWJsZUNhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgY29uc3QgdGFibGVXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0YWJsZVdyYXAuY2xhc3NOYW1lID0gJ3RhYmxlLXNjcm9sbCc7CiAgICB0YWJsZVdyYXAuaWQgPSAnZm9sbG93ZXJzVGFibGVXcmFwJzsKICAgIHRhYmxlQ2FyZC5hcHBlbmRDaGlsZCh0YWJsZVdyYXApOwogICAgY29uc3QgcGFnZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHBhZ2VyLmNsYXNzTmFtZSA9ICdwYWdpbmF0aW9uLXJvdyc7CiAgICBwYWdlci5pZCA9ICdmb2xsb3dlcnNQYWdlcic7CiAgICB0YWJsZUNhcmQuYXBwZW5kQ2hpbGQocGFnZXIpOwogICAgcm9vdC5hcHBlbmRDaGlsZCh0YWJsZUNhcmQpOwoKICAgIHJlbmRlckZvcm0oKTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlckZvcm0oKSB7CiAgICBjb25zdCBjYXJkID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZvbGxvd2Vyc0Zvcm1DYXJkJyk7CiAgICBpZiAoIWNhcmQpIHJldHVybjsKICAgIGNhcmQuaW5uZXJIVE1MID0gJyc7CiAgICBjb25zdCBoZWFkZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGhlYWRlci5jbGFzc05hbWUgPSAnY2FyZC1oZWFkZXInOwogICAgaGVhZGVyLmFwcGVuZENoaWxkKHRleHRFbCgnaDMnLCBlZGl0aW5nSWQgIT09IG51bGwgPyAnRWRpdCBlbnRyeScgOiAnQWRkIGEgd2Vla2x5IGVudHJ5JykpOwogICAgY2FyZC5hcHBlbmRDaGlsZChoZWFkZXIpOwoKICAgIGNvbnN0IGdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGdyaWQuY2xhc3NOYW1lID0gJ2Zvcm0tZ3JpZCc7CgogICAgY29uc3QgcGxhdGZvcm1GaWVsZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcGxhdGZvcm1GaWVsZC5jbGFzc05hbWUgPSAnZm9ybS1maWVsZCc7CiAgICBwbGF0Zm9ybUZpZWxkLmFwcGVuZENoaWxkKHRleHRFbCgnbGFiZWwnLCAnUGxhdGZvcm0nKSk7CiAgICBjb25zdCBwbGF0Zm9ybVNlbGVjdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NlbGVjdCcpOwogICAgcGxhdGZvcm1TZWxlY3QuaWQgPSAnZm9sbG93ZXJzUGxhdGZvcm1JbnB1dCc7CiAgICBhbGxQbGF0Zm9ybXMoKS5mb3JFYWNoKChwKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOwogICAgICBvcHQudmFsdWUgPSBwLmlkOwogICAgICBvcHQudGV4dENvbnRlbnQgPSBwLmxhYmVsOwogICAgICBwbGF0Zm9ybVNlbGVjdC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgICBwbGF0Zm9ybUZpZWxkLmFwcGVuZENoaWxkKHBsYXRmb3JtU2VsZWN0KTsKCiAgICBjb25zdCBkYXRlRmllbGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGRhdGVGaWVsZC5jbGFzc05hbWUgPSAnZm9ybS1maWVsZCc7CiAgICBkYXRlRmllbGQuYXBwZW5kQ2hpbGQodGV4dEVsKCdsYWJlbCcsICdXZWVrIC8gRGF0ZScpKTsKICAgIGNvbnN0IGRhdGVJbnB1dCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7CiAgICBkYXRlSW5wdXQudHlwZSA9ICdkYXRlJzsKICAgIGRhdGVJbnB1dC5pZCA9ICdmb2xsb3dlcnNEYXRlSW5wdXQnOwogICAgZGF0ZUZpZWxkLmFwcGVuZENoaWxkKGRhdGVJbnB1dCk7CgogICAgY29uc3QgY291bnRGaWVsZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY291bnRGaWVsZC5jbGFzc05hbWUgPSAnZm9ybS1maWVsZCc7CiAgICBjb3VudEZpZWxkLmFwcGVuZENoaWxkKHRleHRFbCgnbGFiZWwnLCAnRm9sbG93ZXJzIENvdW50JykpOwogICAgY29uc3QgY291bnRJbnB1dCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7CiAgICBjb3VudElucHV0LnR5cGUgPSAnbnVtYmVyJzsKICAgIGNvdW50SW5wdXQubWluID0gJzAnOwogICAgY291bnRJbnB1dC5zdGVwID0gJzEnOwogICAgY291bnRJbnB1dC5pZCA9ICdmb2xsb3dlcnNDb3VudElucHV0JzsKICAgIGNvdW50RmllbGQuYXBwZW5kQ2hpbGQoY291bnRJbnB1dCk7CgogICAgZ3JpZC5hcHBlbmQocGxhdGZvcm1GaWVsZCwgZGF0ZUZpZWxkLCBjb3VudEZpZWxkKTsKICAgIGNhcmQuYXBwZW5kQ2hpbGQoZ3JpZCk7CgogICAgY29uc3QgZWRpdFJvdyA9IGVkaXRpbmdJZCAhPT0gbnVsbCA/IGN1cnJlbnRSb3dzLmZpbmQoKHIpID0+IHIuaWQgPT09IGVkaXRpbmdJZCkgOiBudWxsOwogICAgaWYgKGVkaXRSb3cpIHsKICAgICAgcGxhdGZvcm1TZWxlY3QudmFsdWUgPSBlZGl0Um93LnBsYXRmb3JtOwogICAgICBkYXRlSW5wdXQudmFsdWUgPSBlZGl0Um93LmVudHJ5X2RhdGU7CiAgICAgIGNvdW50SW5wdXQudmFsdWUgPSBTdHJpbmcoZWRpdFJvdy5mb2xsb3dlcnNfY291bnQpOwogICAgfSBlbHNlIHsKICAgICAgZGF0ZUlucHV0LnZhbHVlID0gbmV3IERhdGUoKS50b0lTT1N0cmluZygpLnNsaWNlKDAsIDEwKTsKICAgIH0KCiAgICBjb25zdCBhY3Rpb25zID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBhY3Rpb25zLmNsYXNzTmFtZSA9ICdtb2RhbC1hY3Rpb25zJzsKICAgIGNvbnN0IGVycm9yRWwgPSB0ZXh0RWwoJ3NwYW4nLCAnJywgJ211dGVkJyk7CiAgICBlcnJvckVsLmlkID0gJ2ZvbGxvd2Vyc0Zvcm1FcnJvcic7CiAgICBlcnJvckVsLnN0eWxlLmNvbG9yID0gJ3ZhcigtLXN0YXR1cy1jcml0aWNhbCknOwoKICAgIGNvbnN0IGJ0blJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYnRuUm93LmNsYXNzTmFtZSA9ICdidG4tcm93JzsKICAgIGNvbnN0IHNhdmVCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIHNhdmVCdG4uY2xhc3NOYW1lID0gJ2J0biBwcmltYXJ5JzsKICAgIHNhdmVCdG4udGV4dENvbnRlbnQgPSBlZGl0aW5nSWQgIT09IG51bGwgPyAnU2F2ZSBjaGFuZ2VzJyA6ICdBZGQgZW50cnknOwogICAgc2F2ZUJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHN1Ym1pdEZvcm0oc2F2ZUJ0bikpOwogICAgYnRuUm93LmFwcGVuZENoaWxkKHNhdmVCdG4pOwogICAgaWYgKGVkaXRpbmdJZCAhPT0gbnVsbCkgewogICAgICBjb25zdCBjYW5jZWxCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgICAgY2FuY2VsQnRuLmNsYXNzTmFtZSA9ICdidG4nOwogICAgICBjYW5jZWxCdG4udGV4dENvbnRlbnQgPSAnQ2FuY2VsJzsKICAgICAgY2FuY2VsQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4geyBlZGl0aW5nSWQgPSBudWxsOyByZW5kZXJGb3JtKCk7IH0pOwogICAgICBidG5Sb3cuYXBwZW5kQ2hpbGQoY2FuY2VsQnRuKTsKICAgIH0KICAgIGFjdGlvbnMuYXBwZW5kKGVycm9yRWwsIGJ0blJvdyk7CiAgICBjYXJkLmFwcGVuZENoaWxkKGFjdGlvbnMpOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gc3VibWl0Rm9ybShidG4pIHsKICAgIGNvbnN0IGVycm9yRWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZm9sbG93ZXJzRm9ybUVycm9yJyk7CiAgICBlcnJvckVsLnRleHRDb250ZW50ID0gJyc7CiAgICBjb25zdCBwbGF0Zm9ybSA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmb2xsb3dlcnNQbGF0Zm9ybUlucHV0JykudmFsdWU7CiAgICBjb25zdCBlbnRyeURhdGUgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZm9sbG93ZXJzRGF0ZUlucHV0JykudmFsdWU7CiAgICBjb25zdCBmb2xsb3dlcnNDb3VudCA9IE51bWJlcihkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZm9sbG93ZXJzQ291bnRJbnB1dCcpLnZhbHVlKTsKICAgIGJ0bi5kaXNhYmxlZCA9IHRydWU7CiAgICB0cnkgewogICAgICBpZiAoZWRpdGluZ0lkICE9PSBudWxsKSB7CiAgICAgICAgYXdhaXQgQXBpLnVwZGF0ZUZvbGxvd2VycyhlZGl0aW5nSWQsIHsgcGxhdGZvcm0sIGVudHJ5RGF0ZSwgZm9sbG93ZXJzQ291bnQgfSk7CiAgICAgICAgVG9hc3Quc2hvdygnRW50cnkgdXBkYXRlZC4nLCAnc3VjY2VzcycpOwogICAgICB9IGVsc2UgewogICAgICAgIGF3YWl0IEFwaS5zYXZlRm9sbG93ZXJzKHsgcGxhdGZvcm0sIGVudHJ5RGF0ZSwgZm9sbG93ZXJzQ291bnQgfSk7CiAgICAgICAgVG9hc3Quc2hvdygnRW50cnkgc2F2ZWQuJywgJ3N1Y2Nlc3MnKTsKICAgICAgfQogICAgICBlZGl0aW5nSWQgPSBudWxsOwogICAgICBhd2FpdCBsb2FkKCk7CiAgICAgIHdpbmRvdy5kaXNwYXRjaEV2ZW50KG5ldyBDdXN0b21FdmVudCgnbHJzOmRhdGEtdXBkYXRlZCcpKTsKICAgIH0gY2F0Y2ggKGVycikgewogICAgICBlcnJvckVsLnRleHRDb250ZW50ID0gZXJyLm1lc3NhZ2U7CiAgICAgIGJ0bi5kaXNhYmxlZCA9IGZhbHNlOwogICAgfQogIH0KCiAgZnVuY3Rpb24gc3RhcnRFZGl0KHJvdykgewogICAgZWRpdGluZ0lkID0gcm93LmlkOwogICAgcmVuZGVyRm9ybSgpOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZvbGxvd2Vyc0Zvcm1DYXJkJykuc2Nyb2xsSW50b1ZpZXcoeyBiZWhhdmlvcjogJ3Ntb290aCcsIGJsb2NrOiAnc3RhcnQnIH0pOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gaGFuZGxlRGVsZXRlKHJvdykgewogICAgY29uc3Qgc3VyZSA9IHdpbmRvdy5jb25maXJtKGBEZWxldGUgdGhlICR7cGxhdGZvcm1NZXRhRm9yKHJvdy5wbGF0Zm9ybSkubGFiZWx9IGVudHJ5IGZvciAke0Zvcm1hdC5kYXRlKHJvdy5lbnRyeV9kYXRlKX0/YCk7CiAgICBpZiAoIXN1cmUpIHJldHVybjsKICAgIHRyeSB7CiAgICAgIGF3YWl0IEFwaS5kZWxldGVGb2xsb3dlcnMocm93LmlkKTsKICAgICAgVG9hc3Quc2hvdygnRW50cnkgZGVsZXRlZC4nLCAnc3VjY2VzcycpOwogICAgICBpZiAoZWRpdGluZ0lkID09PSByb3cuaWQpIGVkaXRpbmdJZCA9IG51bGw7CiAgICAgIGF3YWl0IGxvYWQoKTsKICAgICAgd2luZG93LmRpc3BhdGNoRXZlbnQobmV3IEN1c3RvbUV2ZW50KCdscnM6ZGF0YS11cGRhdGVkJykpOwogICAgfSBjYXRjaCAoZXJyKSB7CiAgICAgIFRvYXN0LnNob3coZXJyLm1lc3NhZ2UsICdlcnJvcicpOwogICAgfQogIH0KCiAgZnVuY3Rpb24gc29ydGFibGVIZWFkZXIobGFiZWwsIGtleSwgdHlwZSkgewogICAgY29uc3QgdGggPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0aCcpOwogICAgaWYgKHR5cGUgPT09ICdudW1iZXInKSB0aC5jbGFzc05hbWUgPSAnbnVtJzsKICAgIHRoLmNsYXNzTGlzdC5hZGQoJ3NvcnRhYmxlLXRoJyk7CiAgICBjb25zdCBpc0FjdGl2ZSA9IHNvcnRTdGF0ZS5rZXkgPT09IGtleTsKICAgIHRoLmFwcGVuZENoaWxkKGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKGxhYmVsKSk7CiAgICB0aC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3NwYW4nLCBpc0FjdGl2ZSA/IChzb3J0U3RhdGUuZGlyID09PSAnYXNjJyA/ICcg4oaRJyA6ICcg4oaTJykgOiAnIOKGlScsICdzb3J0LWFycm93JykpOwogICAgdGguYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7CiAgICAgIHNvcnRTdGF0ZSA9IHsga2V5LCBkaXI6IHNvcnRTdGF0ZS5rZXkgPT09IGtleSAmJiBzb3J0U3RhdGUuZGlyID09PSAnYXNjJyA/ICdkZXNjJyA6ICdhc2MnLCB0eXBlIH07CiAgICAgIHJlbmRlclRhYmxlKCk7CiAgICB9KTsKICAgIHJldHVybiB0aDsKICB9CgogIGZ1bmN0aW9uIGZpbHRlcmVkUm93cygpIHsKICAgIGNvbnN0IHEgPSBzZWFyY2hWYWx1ZS50cmltKCkudG9Mb3dlckNhc2UoKTsKICAgIGlmICghcSkgcmV0dXJuIGN1cnJlbnRSb3dzOwogICAgcmV0dXJuIGN1cnJlbnRSb3dzLmZpbHRlcigocm93KSA9PiB7CiAgICAgIGNvbnN0IGxhYmVsID0gcGxhdGZvcm1NZXRhRm9yKHJvdy5wbGF0Zm9ybSkubGFiZWwudG9Mb3dlckNhc2UoKTsKICAgICAgcmV0dXJuIGxhYmVsLmluY2x1ZGVzKHEpIHx8IHJvdy5lbnRyeV9kYXRlLmluY2x1ZGVzKHEpIHx8IFN0cmluZyhyb3cuZm9sbG93ZXJzX2NvdW50KS5pbmNsdWRlcyhxKTsKICAgIH0pOwogIH0KCiAgZnVuY3Rpb24gc29ydGVkUm93cygpIHsKICAgIGNvbnN0IHsga2V5LCBkaXIsIHR5cGUgfSA9IHNvcnRTdGF0ZTsKICAgIGNvbnN0IGZhY3RvciA9IGRpciA9PT0gJ2FzYycgPyAxIDogLTE7CiAgICByZXR1cm4gWy4uLmZpbHRlcmVkUm93cygpXS5zb3J0KChhLCBiKSA9PiB7CiAgICAgIGNvbnN0IGF2ID0gYVtrZXldOwogICAgICBjb25zdCBidiA9IGJba2V5XTsKICAgICAgaWYgKGF2ID09PSBudWxsIHx8IGF2ID09PSB1bmRlZmluZWQpIHJldHVybiAxOwogICAgICBpZiAoYnYgPT09IG51bGwgfHwgYnYgPT09IHVuZGVmaW5lZCkgcmV0dXJuIC0xOwogICAgICBpZiAodHlwZSA9PT0gJ251bWJlcicpIHJldHVybiAoYXYgLSBidikgKiBmYWN0b3I7CiAgICAgIHJldHVybiBTdHJpbmcoYXYpLmxvY2FsZUNvbXBhcmUoU3RyaW5nKGJ2KSkgKiBmYWN0b3I7CiAgICB9KTsKICB9CgogIC8qKiBFdmVyeSBjdXJyZW50bHktZmlsdGVyZWQvc29ydGVkIHJvdywgc2hhcGVkIGZvciBleHBvcnQgKG5vdCBqdXN0IHRoZSBjdXJyZW50IHBhZ2UpLiAqLwogIGZ1bmN0aW9uIGV4cG9ydFJvd3MoKSB7CiAgICByZXR1cm4gc29ydGVkUm93cygpLm1hcCgocm93KSA9PiAoewogICAgICBwbGF0Zm9ybV9sYWJlbDogcGxhdGZvcm1NZXRhRm9yKHJvdy5wbGF0Zm9ybSkubGFiZWwsCiAgICAgIGVudHJ5X2RhdGU6IHJvdy5lbnRyeV9kYXRlLAogICAgICBmb2xsb3dlcnNfY291bnQ6IHJvdy5mb2xsb3dlcnNfY291bnQsCiAgICAgIHVwZGF0ZWRfYXQ6IHJvdy51cGRhdGVkX2F0LAogICAgfSkpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyVGFibGUoKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZvbGxvd2Vyc1RhYmxlV3JhcCcpOwogICAgY29uc3QgcGFnZXJFbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmb2xsb3dlcnNQYWdlcicpOwogICAgaWYgKCF3cmFwKSByZXR1cm47CiAgICBjb25zdCBhbGxTb3J0ZWQgPSBzb3J0ZWRSb3dzKCk7CiAgICBpZiAoIWFsbFNvcnRlZC5sZW5ndGgpIHsKICAgICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgICAgd3JhcC5hcHBlbmRDaGlsZChlbXB0eVN0YXRlKHsKICAgICAgICBpY29uOiAndXNlcnMnLAogICAgICAgIHRpdGxlOiBjdXJyZW50Um93cy5sZW5ndGggPyAnTm8gZW50cmllcyBtYXRjaCB5b3VyIHNlYXJjaCcgOiAnTm8gZm9sbG93ZXIgZW50cmllcyB5ZXQnLAogICAgICAgIG1lc3NhZ2U6IGN1cnJlbnRSb3dzLmxlbmd0aCA/ICdUcnkgYSBkaWZmZXJlbnQgcGxhdGZvcm0gbmFtZSBvciBkYXRlLicgOiAnQWRkIHlvdXIgZmlyc3Qgd2Vla2x5IGZvbGxvd2VyIGNvdW50IGFib3ZlIGZvciBhbnkgcGxhdGZvcm0uJywKICAgICAgfSkpOwogICAgICBpZiAocGFnZXJFbCkgcGFnZXJFbC5pbm5lckhUTUwgPSAnJzsKICAgICAgcmV0dXJuOwogICAgfQogICAgY29uc3QgeyBwYWdlUm93cywgdG90YWxQYWdlcywgc2FmZVBhZ2UsIHRvdGFsIH0gPSBwYWdpbmF0ZUNsaWVudFNpZGUoYWxsU29ydGVkLCBwYWdlLCBwYWdlU2l6ZSk7CiAgICBwYWdlID0gc2FmZVBhZ2U7CgogICAgY29uc3QgdGFibGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0YWJsZScpOwogICAgdGFibGUuY2xhc3NOYW1lID0gJ2RhdGEtdGFibGUnOwogICAgY29uc3QgdGhlYWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0aGVhZCcpOwogICAgY29uc3QgaGVhZFRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgIGhlYWRUci5hcHBlbmQoCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdQbGF0Zm9ybScsICdwbGF0Zm9ybScsICdzdHJpbmcnKSwKICAgICAgc29ydGFibGVIZWFkZXIoJ1dlZWsgLyBEYXRlJywgJ2VudHJ5X2RhdGUnLCAnc3RyaW5nJyksCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdGb2xsb3dlcnMgQ291bnQnLCAnZm9sbG93ZXJzX2NvdW50JywgJ251bWJlcicpLAogICAgICB0ZXh0RWwoJ3RoJywgJ0xhc3QgVXBkYXRlZCcpLAogICAgICB0ZXh0RWwoJ3RoJywgJ0FjdGlvbnMnKQogICAgKTsKICAgIHRoZWFkLmFwcGVuZENoaWxkKGhlYWRUcik7CiAgICBjb25zdCB0Ym9keSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3Rib2R5Jyk7CiAgICBwYWdlUm93cy5mb3JFYWNoKChyb3cpID0+IHsKICAgICAgY29uc3QgdHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgICBjb25zdCBwbGF0Zm9ybVRkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgY29uc3QgbWV0YSA9IHBsYXRmb3JtTWV0YUZvcihyb3cucGxhdGZvcm0pOwogICAgICBjb25zdCBwaWxsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICBwaWxsLmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1waWxsJzsKICAgICAgY29uc3QgZG90ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICBkb3QuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLWRvdCc7CiAgICAgIGRvdC5zdHlsZS5iYWNrZ3JvdW5kID0gbWV0YS5jb2xvcjsKICAgICAgcGlsbC5hcHBlbmQoZG90LCBkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShtZXRhLmxhYmVsKSk7CiAgICAgIHBsYXRmb3JtVGQuYXBwZW5kQ2hpbGQocGlsbCk7CgogICAgICBjb25zdCBhY3Rpb25zVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBjb25zdCByb3dBY3Rpb25zID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIHJvd0FjdGlvbnMuY2xhc3NOYW1lID0gJ3Jvdy1hY3Rpb25zJzsKICAgICAgY29uc3QgZWRpdEJ0biA9IGljb25CdG4oJ2J0bicsICdwZW5jaWwnLCAnRWRpdCcpOwogICAgICBlZGl0QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gc3RhcnRFZGl0KHJvdykpOwogICAgICBjb25zdCBkZWxldGVCdG4gPSBpY29uQnRuKCdidG4gZGFuZ2VyJywgJ3RyYXNoLTInLCAnRGVsZXRlJyk7CiAgICAgIGRlbGV0ZUJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IGhhbmRsZURlbGV0ZShyb3cpKTsKICAgICAgcm93QWN0aW9ucy5hcHBlbmQoZWRpdEJ0biwgZGVsZXRlQnRuKTsKICAgICAgYWN0aW9uc1RkLmFwcGVuZENoaWxkKHJvd0FjdGlvbnMpOwoKICAgICAgdHIuYXBwZW5kKAogICAgICAgIHBsYXRmb3JtVGQsCiAgICAgICAgdGV4dEVsKCd0ZCcsIEZvcm1hdC5kYXRlKHJvdy5lbnRyeV9kYXRlKSksCiAgICAgICAgdGV4dEVsKCd0ZCcsIEZvcm1hdC5udW1iZXIocm93LmZvbGxvd2Vyc19jb3VudCksICdudW0nKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgcm93LnVwZGF0ZWRfYXQpLAogICAgICAgIGFjdGlvbnNUZAogICAgICApOwogICAgICB0Ym9keS5hcHBlbmRDaGlsZCh0cik7CiAgICB9KTsKICAgIHRhYmxlLmFwcGVuZCh0aGVhZCwgdGJvZHkpOwogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGFibGUpOwoKICAgIGlmIChwYWdlckVsKSB7CiAgICAgIHBhZ2VyRWwuaW5uZXJIVE1MID0gJyc7CiAgICAgIHBhZ2VyRWwuYXBwZW5kQ2hpbGQoYnVpbGRQYWdlcih7CiAgICAgICAgcGFnZTogc2FmZVBhZ2UsCiAgICAgICAgdG90YWxQYWdlcywKICAgICAgICB0b3RhbCwKICAgICAgICBvblByZXY6ICgpID0+IHsgcGFnZSAtPSAxOyByZW5kZXJUYWJsZSgpOyB9LAogICAgICAgIG9uTmV4dDogKCkgPT4geyBwYWdlICs9IDE7IHJlbmRlclRhYmxlKCk7IH0sCiAgICAgIH0pKTsKICAgIH0KICB9CgogIGFzeW5jIGZ1bmN0aW9uIGxvYWQoKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZvbGxvd2Vyc1RhYmxlV3JhcCcpOwogICAgaWYgKHdyYXApIHsgd3JhcC5pbm5lckhUTUwgPSAnJzsgd3JhcC5hcHBlbmRDaGlsZChza2VsZXRvblJvd3MoNCkpOyB9CiAgICBjdXJyZW50Um93cyA9IGF3YWl0IEFwaS5saXN0Rm9sbG93ZXJzKHt9KTsKICAgIHJlbmRlclRhYmxlKCk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiByZW5kZXIoKSB7CiAgICByb290ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3ZpZXctZm9sbG93ZXJzJyk7CiAgICBlZGl0aW5nSWQgPSBudWxsOwogICAgc2hlbGwoKTsKICAgIGF3YWl0IGxvYWQoKTsKICB9CgogIHJldHVybiB7IHJlbmRlciB9Owp9KSgpOwoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIFJlcG9ydCBHZW5lcmF0b3IgdGFiOiBhIHBlci1wbGF0Zm9ybSBXZWVrbHkgU29jaWFsIE1lZGlhCiAgIFBlcmZvcm1hbmNlIFJlcG9ydC4gUGljayBhIGN1cnJlbnQgd2VlayAoZGVmYXVsdCBBdWcgMzEg4oCTIFNlcCA2LAogICAyMDI2KTsgdGhlIGltbWVkaWF0ZWx5LXByZWNlZGluZyBlcXVhbC1sZW5ndGggd2VlayBpcyBkZXJpdmVkCiAgIGF1dG9tYXRpY2FsbHkuIEZvciBldmVyeSBwbGF0Zm9ybSBpdCBkcmF3cyBPTkUgZ3JvdXBlZCBiYXIgY2hhcnQKICAgKGl0cyBhcHBsaWNhYmxlIGNhdGVnb3JpZXMgb24gdGhlIHgtYXhpcywgdHdvIGJhcnMgcGVyIGNhdGVnb3J5IOKAlAogICBwcmV2aW91cyB3ZWVrIC8gY3VycmVudCB3ZWVrLCBhY3R1YWwgdmFsdWVzIHByaW50ZWQgb24gdGhlIGJhcnMpCiAgIGZvbGxvd2VkIGJ5IGEgZGV0YWlsIHRhYmxlIChwcmV2aW91cywgY3VycmVudCwgY2hhbmdlLCAlIGNoYW5nZSwKICAgc3RhdHVzKS4gQSB0cmFpbGluZyAiRm9sbG93ZXJzIEdyb3d0aCIgc2VjdGlvbiBjb21wYXJlcyB0b3RhbAogICBmb2xsb3dlciBjb3VudHMgcGVyIHBsYXRmb3JtLiBFdmVyeXRoaW5nIGNvbWVzIGZyb20gdGhlIFNBTUUKICAgZW5kcG9pbnRzIHRoZSByZXN0IG9mIHRoZSBhcHAgdXNlcyDigJQgL2FwaS9hbmFseXRpY3MvY29tcGFyZQogICAocGVyLXBsYXRmb3JtIHBvc3QgbWV0cmljcyArIE4vQSBmbGFncyBmb3IgYm90aCB3ZWVrcyBpbiBvbmUKICAgY2FsbCkgYW5kIC9hcGkvZm9sbG93ZXJzICh0aGUgRm9sbG93ZXJzIERhdGEgUmVjb3JkLCByZWFkLW9ubHkpIOKAlAogICBzbyB0aGVyZSBpcyBubyBzZWNvbmQgYW5hbHl0aWNzIHN5c3RlbSwgbm90aGluZyBpcyB3cml0dGVuIGJhY2ssCiAgIGFuZCB0aGUgRm9sbG93ZXJzIERhdGEgUmVjb3JkIGlzIG5ldmVyIG1vZGlmaWVkLiBOL0EgaXMgc2hvd24KICAgd2hlbmV2ZXIgYSBtZXRyaWMgaXMgbm90IGFwcGxpY2FibGUgdG8gYSBwbGF0Zm9ybSBmb3IgYSB3ZWVrOyBpdAogICBpcyBuZXZlciB0cmVhdGVkIGFzIDAsIGFuZCBubyBjaGFuZ2UgLyAlIGNoYW5nZSBpcyBjb21wdXRlZAogICBhZ2FpbnN0IGl0LiAlIGNoYW5nZSBpcyBhbHNvIE4vQSB3aGVuIHRoZSBwcmV2aW91cyB3ZWVrIGlzIDAuCiAgIFByaW50IC8gRXhwb3J0IFJlcG9ydCBpcyB3aW5kb3cucHJpbnQoKSBhZ2FpbnN0IHRoZSBAbWVkaWEgcHJpbnQKICAgYmxvY2sgaW4gPHN0eWxlPi4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmNvbnN0IFJlcG9ydCA9ICgoKSA9PiB7CiAgLy8gQXBwbGljYWJsZSBjYXRlZ29yaWVzIHBlciBwbGF0Zm9ybSAobWF0Y2hlcyB0aGUgcG9ydGFsJ3Mgb3duIHBlci1wbGF0Zm9ybQogIC8vIGNvbHVtbiBsYWJlbHMpLiBBIHRyYWlsaW5nICJGb2xsb3dlcnMiIGNhdGVnb3J5LCBzb3VyY2VkIGZyb20gdGhlIEZvbGxvd2VycwogIC8vIERhdGEgUmVjb3JkLCBpcyBhcHBlbmRlZCB0byBldmVyeSBwbGF0Zm9ybSBhdXRvbWF0aWNhbGx5LgogIGNvbnN0IFJFUE9SVF9QTEFURk9STVMgPSBbCiAgICB7IGlkOiAnZmFjZWJvb2snLCAgbGFiZWw6ICdGYWNlYm9vaycsICAgICAgIGNhdHM6IFtbJ3ZpZXdzJywgJ1ZpZXdzJ10sIFsncmVhY2gnLCAnUmVhY2gnXSwgWydlbmdhZ2VtZW50JywgJ0VuZ2FnZW1lbnQnXV0gfSwKICAgIHsgaWQ6ICdmYl9ncm91cCcsICBsYWJlbDogJ0ZhY2Vib29rIEdyb3VwJywgY2F0czogW1snZW5nYWdlbWVudCcsICdSZWFjdGlvbnMnXSwgWydjb21tZW50cycsICdDb21tZW50cyddXSB9LAogICAgeyBpZDogJ2luc3RhZ3JhbScsIGxhYmVsOiAnSW5zdGFncmFtJywgICAgICBjYXRzOiBbWyd2aWV3cycsICdWaWV3cyddLCBbJ3JlYWNoJywgJ1JlYWNoJ10sIFsnZW5nYWdlbWVudCcsICdJbnRlcmFjdGlvbnMnXV0gfSwKICAgIHsgaWQ6ICd0aWt0b2snLCAgICBsYWJlbDogJ1Rpa1RvaycsICAgICAgICAgY2F0czogW1sndmlld3MnLCAnVmlld3MnXSwgWydlbmdhZ2VtZW50JywgJ0VuZ2FnZW1lbnRzJ10sIFsnZm9sbG93ZXJzX2dhaW5lZCcsICdGb2xsb3dlcnMgR2FpbmVkJ11dIH0sCiAgICB7IGlkOiAnbGlua2VkaW4nLCAgbGFiZWw6ICdMaW5rZWRJbicsICAgICAgIGNhdHM6IFtbJ2ltcHJlc3Npb25zJywgJ0ltcHJlc3Npb25zJ10sIFsncmVhY2gnLCAnUmVhY2gnXV0gfSwKICAgIHsgaWQ6ICd0aHJlYWRzJywgICBsYWJlbDogJ1RocmVhZHMnLCAgICAgICAgY2F0czogW1sndmlld3MnLCAnVmlld3MnXSwgWydlbmdhZ2VtZW50JywgJ0ludGVyYWN0aW9ucyddXSB9LAogICAgeyBpZDogJ3lvdXR1YmUnLCAgIGxhYmVsOiAnWW91VHViZScsICAgICAgICBjYXRzOiBbWyd2aWV3cycsICdWaWV3cyddLCBbJ2ltcHJlc3Npb25zJywgJ0ltcHJlc3Npb25zJ11dIH0sCiAgXTsKICBjb25zdCBDSEFSVF9JRFMgPSBSRVBPUlRfUExBVEZPUk1TLm1hcCgocCkgPT4gJ3JlcG9ydENoYXJ0XycgKyBwLmlkKS5jb25jYXQoWydyZXBvcnRGb2xsb3dlcnNDaGFydCddKTsKICBjb25zdCBOQSA9ICdOL0EnOwoKICBsZXQgcm9vdDsKICAvLyBEZWZhdWx0IHBlciBzcGVjOiBBdWcgMzEg4oCTIFNlcCA2LCAyMDI2IChjdXJyZW50KSB2cyBBdWcgMjQg4oCTIDMwLCAyMDI2IChwcmV2aW91cykuCiAgbGV0IHJhbmdlU3RhcnQgPSAnMjAyNi0wOC0zMSc7CiAgbGV0IHJhbmdlRW5kID0gJzIwMjYtMDktMDYnOwogIGxldCBsYXN0TW9kZWwgPSBudWxsOwogIGxldCBidXN5ID0gZmFsc2U7CgogIC8qIC0tLS0tLS0tLS0gZGF0ZSBoZWxwZXJzIC0tLS0tLS0tLS0gKi8KICBmdW5jdGlvbiBtb25kYXlPZihkYXRlU3RyKSB7CiAgICBjb25zdCBkID0gbmV3IERhdGUoZGF0ZVN0ciArICdUMDA6MDA6MDAnKTsKICAgIGNvbnN0IGRheSA9IGQuZ2V0RGF5KCk7CiAgICBkLnNldERhdGUoZC5nZXREYXRlKCkgLSAoZGF5ID09PSAwID8gNiA6IGRheSAtIDEpKTsKICAgIHJldHVybiBkLnRvSVNPU3RyaW5nKCkuc2xpY2UoMCwgMTApOwogIH0KICBmdW5jdGlvbiBhZGREYXlzKGRhdGVTdHIsIG4pIHsKICAgIGNvbnN0IGQgPSBuZXcgRGF0ZShkYXRlU3RyICsgJ1QwMDowMDowMCcpOwogICAgZC5zZXREYXRlKGQuZ2V0RGF0ZSgpICsgbik7CiAgICByZXR1cm4gZC50b0lTT1N0cmluZygpLnNsaWNlKDAsIDEwKTsKICB9CiAgZnVuY3Rpb24gZGF5c0luY2x1c2l2ZShmcm9tLCB0bykgewogICAgcmV0dXJuIE1hdGgucm91bmQoKG5ldyBEYXRlKHRvICsgJ1QwMDowMDowMCcpIC0gbmV3IERhdGUoZnJvbSArICdUMDA6MDA6MDAnKSkgLyA4NjQwMDAwMCkgKyAxOwogIH0KICBmdW5jdGlvbiB5ZWFyT2YoZGF0ZVN0cikgeyByZXR1cm4gZGF0ZVN0ci5zbGljZSgwLCA0KTsgfQogIGNvbnN0IE1PTiA9IFsnSmFuJywgJ0ZlYicsICdNYXInLCAnQXByJywgJ01heScsICdKdW4nLCAnSnVsJywgJ0F1ZycsICdTZXAnLCAnT2N0JywgJ05vdicsICdEZWMnXTsKICBmdW5jdGlvbiBzaG9ydFJhbmdlKGZyb20sIHRvKSB7CiAgICBjb25zdCBbLCBtMSwgZDFdID0gZnJvbS5zcGxpdCgnLScpLm1hcChOdW1iZXIpOwogICAgY29uc3QgWywgbTIsIGQyXSA9IHRvLnNwbGl0KCctJykubWFwKE51bWJlcik7CiAgICBpZiAobTEgPT09IG0yKSByZXR1cm4gTU9OW20xIC0gMV0gKyAnICcgKyBkMSArICfigJMnICsgZDI7CiAgICByZXR1cm4gTU9OW20xIC0gMV0gKyAnICcgKyBkMSArICcg4oCTICcgKyBNT05bbTIgLSAxXSArICcgJyArIGQyOwogIH0KICBmdW5jdGlvbiBwZXJpb2RzKCkgewogICAgY29uc3QgY3VyID0geyBmcm9tOiByYW5nZVN0YXJ0LCB0bzogcmFuZ2VFbmQgfTsKICAgIGNvbnN0IGxlbiA9IGRheXNJbmNsdXNpdmUocmFuZ2VTdGFydCwgcmFuZ2VFbmQpOwogICAgY29uc3QgcHJldiA9IHsgZnJvbTogYWRkRGF5cyhyYW5nZVN0YXJ0LCAtbGVuKSwgdG86IGFkZERheXMocmFuZ2VTdGFydCwgLTEpIH07CiAgICByZXR1cm4geyBjdXIsIHByZXYsIGxlbiB9OwogIH0KCiAgLyogLS0tLS0tLS0tLSBudW1iZXIgLyBzdGF0dXMgZm9ybWF0dGluZyAtLS0tLS0tLS0tICovCiAgZnVuY3Rpb24gZm10Q2VsbCh2KSB7IHJldHVybiB2ID09PSBudWxsIHx8IHYgPT09IHVuZGVmaW5lZCA/IE5BIDogRm9ybWF0Lm51bWJlcih2KTsgfQogIGZ1bmN0aW9uIGZtdENoYW5nZSh2KSB7CiAgICBpZiAodiA9PT0gbnVsbCB8fCB2ID09PSB1bmRlZmluZWQpIHJldHVybiBOQTsKICAgIGlmICh2ID4gMCkgcmV0dXJuICcrJyArIEZvcm1hdC5udW1iZXIodik7CiAgICByZXR1cm4gRm9ybWF0Lm51bWJlcih2KTsgLy8gRm9ybWF0Lm51bWJlciBhbHJlYWR5IGNhcnJpZXMgdGhlIG1pbnVzIHNpZ24KICB9CiAgZnVuY3Rpb24gZm10UGN0KHYpIHsKICAgIGlmICh2ID09PSBudWxsIHx8IHYgPT09IHVuZGVmaW5lZCkgcmV0dXJuIE5BOwogICAgcmV0dXJuICh2ID4gMCA/ICcrJyA6ICcnKSArIE51bWJlcih2KS50b0ZpeGVkKDEpLnJlcGxhY2UoL1wuMCQvLCAnJykgKyAnJSc7CiAgfQogIGZ1bmN0aW9uIGNoYW5nZUNsYXNzKHJvdykgewogICAgaWYgKHJvdy5uYSB8fCByb3cuY2hhbmdlID09PSBudWxsKSByZXR1cm4gJ3RyZW5kLWZsYXQnOwogICAgaWYgKHJvdy5jaGFuZ2UgPiAwKSByZXR1cm4gJ3RyZW5kLXVwJzsKICAgIGlmIChyb3cuY2hhbmdlIDwgMCkgcmV0dXJuICd0cmVuZC1kb3duJzsKICAgIHJldHVybiAndHJlbmQtZmxhdCc7CiAgfQogIGZ1bmN0aW9uIHBsYXRmb3JtQ29sb3IoaWQpIHsKICAgIGNvbnN0IGNhY2hlID0gd2luZG93Ll9fZmlsdGVyT3B0aW9uc0NhY2hlIHx8IHt9OwogICAgY29uc3QgYWxsID0gY2FjaGUuYWxsUGxhdGZvcm1zIHx8IGNhY2hlLnBsYXRmb3JtcyB8fCBbXTsKICAgIGNvbnN0IHAgPSBhbGwuZmluZCgoeCkgPT4geC5pZCA9PT0gaWQpOwogICAgcmV0dXJuIHAgPyBwLmNvbG9yIDogJ3ZhcigtLXNlcmllcy0xKSc7CiAgfQoKICAvKiAtLS0tLS0tLS0tIEZvbGxvd2VycyBEYXRhIFJlY29yZDogbGF0ZXN0IGNvdW50IHBlciBwbGF0Zm9ybSBvbi9iZWZvcmUgYSBkYXRlIC0tLS0tLS0tLS0gKi8KICBmdW5jdGlvbiBmb2xsb3dlckFzT2Yocm93cywgcGxhdGZvcm1JZCwgZGF0ZVN0cikgewogICAgbGV0IGJlc3QgPSBudWxsOwogICAgZm9yIChjb25zdCByIG9mIHJvd3MpIHsKICAgICAgaWYgKHIucGxhdGZvcm0gIT09IHBsYXRmb3JtSWQgfHwgci5lbnRyeV9kYXRlID4gZGF0ZVN0cikgY29udGludWU7CiAgICAgIGlmICghYmVzdCB8fCByLmVudHJ5X2RhdGUgPiBiZXN0LmVudHJ5X2RhdGUpIGJlc3QgPSByOwogICAgfQogICAgcmV0dXJuIGJlc3QgPyBiZXN0LmZvbGxvd2Vyc19jb3VudCA6IG51bGw7IC8vIG51bGwgPT09IE4vQSAobm8gcmVjb3JkIGZvciB0aGlzIHBsYXRmb3JtIHlldCkKICB9CgogIC8qIC0tLS0tLS0tLS0gb25lIGNvbXBhcmlzb24gcm93IChhIGNhdGVnb3J5LCBvciBhIHBsYXRmb3JtIGluIHRoZSBGb2xsb3dlcnMgR3Jvd3RoIHRhYmxlKSAtLS0tLS0tLS0tCiAgICAgbmEgID0gZWl0aGVyIHdlZWsgaGFzIG5vIGNvbXBhcmFibGUgdmFsdWUg4oaSIGNoYW5nZSAvICUgY2hhbmdlIGFyZSBOL0EsIG5ldmVyIDAuCiAgICAgJSBjaGFuZ2UgaXMgYWxzbyBOL0Egd2hlbiB0aGUgcHJldmlvdXMgd2VlayBpcyBleGFjdGx5IDAgKHBlciB0aGUgc3BlYydzIGZvcm11bGEgZ3VhcmQpLiAqLwogIGZ1bmN0aW9uIG1ha2VSb3cobGFiZWwsIHByZXZWLCBjdXJWKSB7CiAgICBjb25zdCBuYSA9IHByZXZWID09PSBudWxsIHx8IHByZXZWID09PSB1bmRlZmluZWQgfHwgY3VyViA9PT0gbnVsbCB8fCBjdXJWID09PSB1bmRlZmluZWQ7CiAgICBjb25zdCBjaGFuZ2UgPSBuYSA/IG51bGwgOiAoY3VyViAtIHByZXZWKTsKICAgIGNvbnN0IHBjdCA9IChuYSB8fCBwcmV2ViA9PT0gMCkgPyBudWxsIDogTWF0aC5yb3VuZCgoKGN1clYgLSBwcmV2VikgLyBwcmV2VikgKiAxMDAwKSAvIDEwOwogICAgbGV0IHN0YXR1czsKICAgIGlmIChuYSkgc3RhdHVzID0gTkE7CiAgICBlbHNlIGlmIChjaGFuZ2UgPiAwKSBzdGF0dXMgPSAnSW1wcm92ZWQnOwogICAgZWxzZSBpZiAoY2hhbmdlIDwgMCkgc3RhdHVzID0gJ0RlY2xpbmVkJzsKICAgIGVsc2Ugc3RhdHVzID0gJ05vIGNoYW5nZSc7CiAgICByZXR1cm4geyBsYWJlbCwgcHJldjogcHJldlYgPT09IHVuZGVmaW5lZCA/IG51bGwgOiBwcmV2ViwgY3VyOiBjdXJWID09PSB1bmRlZmluZWQgPyBudWxsIDogY3VyViwgbmEsIGNoYW5nZSwgcGN0LCBzdGF0dXMgfTsKICB9CgogIGZ1bmN0aW9uIGJ1aWxkTW9kZWwoY21wLCBmb2xsb3dlclJvd3MsIGN1ciwgcHJldikgewogICAgY29uc3QgYUJ5ID0gT2JqZWN0LmZyb21FbnRyaWVzKChjbXAucmFuZ2VBLnBsYXRmb3JtcyB8fCBbXSkubWFwKChwKSA9PiBbcC5wbGF0Zm9ybSwgcF0pKTsKICAgIGNvbnN0IGJCeSA9IE9iamVjdC5mcm9tRW50cmllcygoY21wLnJhbmdlQi5wbGF0Zm9ybXMgfHwgW10pLm1hcCgocCkgPT4gW3AucGxhdGZvcm0sIHBdKSk7CgogICAgY29uc3QgcGxhdGZvcm1zID0gUkVQT1JUX1BMQVRGT1JNUy5tYXAoKGRlZikgPT4gewogICAgICBjb25zdCBwYSA9IGFCeVtkZWYuaWRdIHx8IG51bGw7CiAgICAgIGNvbnN0IHBiID0gYkJ5W2RlZi5pZF0gfHwgbnVsbDsKCiAgICAgIGNvbnN0IHJvd3MgPSBkZWYuY2F0cy5tYXAoKFtrZXksIGNhdExhYmVsXSkgPT4gewogICAgICAgIC8vIGFwcGxpY2FibGUgPT09IGZhbHNlIOKGkiB0aGUgcGxhdGZvcm0gZG9lc24ndCBjYXJyeSB0aGlzIG1ldHJpYyB0aGF0IHdlZWsg4oaSIE4vQSAobm90IDApLgogICAgICAgIGNvbnN0IG9rQSA9IHBhID8gKHBhLmFwcGxpY2FibGUgPyAhIXBhLmFwcGxpY2FibGVba2V5XSA6IHRydWUpIDogZmFsc2U7CiAgICAgICAgY29uc3Qgb2tCID0gcGIgPyAocGIuYXBwbGljYWJsZSA/ICEhcGIuYXBwbGljYWJsZVtrZXldIDogdHJ1ZSkgOiBmYWxzZTsKICAgICAgICByZXR1cm4gbWFrZVJvdyhjYXRMYWJlbCwgb2tCID8gKHBiW2tleV0gfHwgMCkgOiBudWxsLCBva0EgPyAocGFba2V5XSB8fCAwKSA6IG51bGwpOwogICAgICB9KTsKCiAgICAgIGNvbnN0IGZDdXIgPSBmb2xsb3dlckFzT2YoZm9sbG93ZXJSb3dzLCBkZWYuaWQsIGN1ci50byk7CiAgICAgIGNvbnN0IGZQcmV2ID0gZm9sbG93ZXJBc09mKGZvbGxvd2VyUm93cywgZGVmLmlkLCBwcmV2LnRvKTsKICAgICAgcm93cy5wdXNoKG1ha2VSb3coJ0ZvbGxvd2VycycsIGZQcmV2LCBmQ3VyKSk7CgogICAgICBjb25zdCBoYXNQb3N0cyA9IEJvb2xlYW4oKHBhICYmIHBhLnBvc3RfY291bnQpIHx8IChwYiAmJiBwYi5wb3N0X2NvdW50KSk7CiAgICAgIGNvbnN0IGhhc0ZvbGxvd2VycyA9IGZDdXIgIT09IG51bGwgfHwgZlByZXYgIT09IG51bGw7CiAgICAgIHJldHVybiB7CiAgICAgICAgaWQ6IGRlZi5pZCwgbGFiZWw6IGRlZi5sYWJlbCwgY29sb3I6IChwYSAmJiBwYS5jb2xvcikgfHwgKHBiICYmIHBiLmNvbG9yKSB8fCBwbGF0Zm9ybUNvbG9yKGRlZi5pZCksCiAgICAgICAgcm93cywgaGFzUG9zdHMsIGhhc0ZvbGxvd2VycywgcHJlc2VudDogaGFzUG9zdHMgfHwgaGFzRm9sbG93ZXJzLAogICAgICB9OwogICAgfSk7CgogICAgY29uc3QgZm9sbG93ZXJzR3Jvd3RoID0gUkVQT1JUX1BMQVRGT1JNUwogICAgICAubWFwKChkZWYpID0+IHsKICAgICAgICBjb25zdCByID0gbWFrZVJvdyhkZWYubGFiZWwsIGZvbGxvd2VyQXNPZihmb2xsb3dlclJvd3MsIGRlZi5pZCwgcHJldi50byksIGZvbGxvd2VyQXNPZihmb2xsb3dlclJvd3MsIGRlZi5pZCwgY3VyLnRvKSk7CiAgICAgICAgcmV0dXJuIHsgaWQ6IGRlZi5pZCwgY29sb3I6IHBsYXRmb3JtQ29sb3IoZGVmLmlkKSwgLi4uciB9OwogICAgICB9KQogICAgICAuZmlsdGVyKChyKSA9PiByLnByZXYgIT09IG51bGwgfHwgci5jdXIgIT09IG51bGwpOwoKICAgIHJldHVybiB7CiAgICAgIGN1ciwgcHJldiwKICAgICAgY3VyTGFiZWw6IHNob3J0UmFuZ2UoY3VyLmZyb20sIGN1ci50byksCiAgICAgIHByZXZMYWJlbDogc2hvcnRSYW5nZShwcmV2LmZyb20sIHByZXYudG8pLAogICAgICBwbGF0Zm9ybXMsIGZvbGxvd2Vyc0dyb3d0aCwKICAgIH07CiAgfQoKICAvKiAtLS0tLS0tLS0tIHJlbmRlcmluZyAtLS0tLS0tLS0tICovCiAgZnVuY3Rpb24gc3RhdHVzQ2VsbChzdGF0dXMpIHsKICAgIGNvbnN0IHRkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgIGNvbnN0IHNwYW4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICBjb25zdCBjbHMgPSBzdGF0dXMgPT09ICdJbXByb3ZlZCcgPyAndHJlbmQtdXAnIDogc3RhdHVzID09PSAnRGVjbGluZWQnID8gJ3RyZW5kLWRvd24nIDogJ3RyZW5kLWZsYXQnOwogICAgc3Bhbi5jbGFzc05hbWUgPSAncmVwb3J0LXN0YXR1cyAnICsgY2xzOwogICAgc3Bhbi50ZXh0Q29udGVudCA9IHN0YXR1czsKICAgIHRkLmFwcGVuZENoaWxkKHNwYW4pOwogICAgcmV0dXJuIHRkOwogIH0KCiAgZnVuY3Rpb24gZGV0YWlsVGFibGUocm93cywgZmlyc3RDb2wsIHByZXZMYWJlbCwgY3VyTGFiZWwpIHsKICAgIGNvbnN0IHRhYmxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGFibGUnKTsKICAgIHRhYmxlLmNsYXNzTmFtZSA9ICdkYXRhLXRhYmxlJzsKICAgIGNvbnN0IHRoZWFkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGhlYWQnKTsKICAgIGNvbnN0IGh0ciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICBodHIuYXBwZW5kKAogICAgICB0ZXh0RWwoJ3RoJywgZmlyc3RDb2wpLAogICAgICB0ZXh0RWwoJ3RoJywgcHJldkxhYmVsLCAnbnVtJyksCiAgICAgIHRleHRFbCgndGgnLCBjdXJMYWJlbCwgJ251bScpLAogICAgICB0ZXh0RWwoJ3RoJywgJ0NoYW5nZScsICdudW0nKSwKICAgICAgdGV4dEVsKCd0aCcsICclIENoYW5nZScsICdudW0nKSwKICAgICAgdGV4dEVsKCd0aCcsICdTdGF0dXMnKSwKICAgICk7CiAgICB0aGVhZC5hcHBlbmRDaGlsZChodHIpOwogICAgY29uc3QgdGJvZHkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0Ym9keScpOwogICAgcm93cy5mb3JFYWNoKChyKSA9PiB7CiAgICAgIGNvbnN0IHRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgY29uc3QgY2xzID0gY2hhbmdlQ2xhc3Mocik7CiAgICAgIHRyLmFwcGVuZCgKICAgICAgICB0ZXh0RWwoJ3RkJywgci5sYWJlbCksCiAgICAgICAgdGV4dEVsKCd0ZCcsIGZtdENlbGwoci5wcmV2KSwgJ251bScpLAogICAgICAgIHRleHRFbCgndGQnLCBmbXRDZWxsKHIuY3VyKSwgJ251bScpLAogICAgICAgIHRleHRFbCgndGQnLCBmbXRDaGFuZ2Uoci5jaGFuZ2UpLCAnbnVtICcgKyBjbHMpLAogICAgICAgIHRleHRFbCgndGQnLCBmbXRQY3Qoci5wY3QpLCAnbnVtICcgKyBjbHMpLAogICAgICAgIHN0YXR1c0NlbGwoci5zdGF0dXMpLAogICAgICApOwogICAgICB0Ym9keS5hcHBlbmRDaGlsZCh0cik7CiAgICB9KTsKICAgIHRhYmxlLmFwcGVuZCh0aGVhZCwgdGJvZHkpOwogICAgY29uc3Qgc2Nyb2xsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBzY3JvbGwuY2xhc3NOYW1lID0gJ3RhYmxlLXNjcm9sbCc7CiAgICBzY3JvbGwuYXBwZW5kQ2hpbGQodGFibGUpOwogICAgcmV0dXJuIHNjcm9sbDsKICB9CgogIGZ1bmN0aW9uIGRyYXdHcm91cGVkQ2hhcnQoY2FudmFzSWQsIGxhYmVscywgcm93cywgcHJldkxhYmVsLCBjdXJMYWJlbCkgewogICAgQ2hhcnRzLmdyb3VwZWRCYXJDaGFydChjYW52YXNJZCwgewogICAgICBsYWJlbHMsCiAgICAgIHNlcmllczogWwogICAgICAgIHsgbGFiZWw6IHByZXZMYWJlbCwgY29sb3I6IENoYXJ0cy5zZXJpZXNDb2xvcigwKSwgZGF0YTogcm93cy5tYXAoKHIpID0+IHIucHJldikgfSwKICAgICAgICB7IGxhYmVsOiBjdXJMYWJlbCwgY29sb3I6IENoYXJ0cy5zZXJpZXNDb2xvcigxKSwgZGF0YTogcm93cy5tYXAoKHIpID0+IHIuY3VyKSB9LAogICAgICBdLAogICAgICBmb3JtYXRWYWx1ZTogKHYpID0+IEZvcm1hdC5udW1iZXIodiksCiAgICB9KTsKICB9CgogIGZ1bmN0aW9uIGxpc3QoYXJyKSB7CiAgICBpZiAoYXJyLmxlbmd0aCA9PT0gMSkgcmV0dXJuIGFyclswXTsKICAgIGlmIChhcnIubGVuZ3RoID09PSAyKSByZXR1cm4gYXJyWzBdICsgJyBhbmQgJyArIGFyclsxXTsKICAgIHJldHVybiBhcnIuc2xpY2UoMCwgLTEpLmpvaW4oJywgJykgKyAnLCBhbmQgJyArIGFyclthcnIubGVuZ3RoIC0gMV07CiAgfQoKICAvKiogT25lLXRvLXR3byBzZW50ZW5jZXMgZ2VuZXJhdGVkIHB1cmVseSBmcm9tIHRoZSBjYWxjdWxhdGVkIGRlbHRhcyDigJQgbmV2ZXIgZml4ZWQgdGV4dC4gKi8KICBmdW5jdGlvbiBidWlsZFN1bW1hcnkobW9kZWwsIHByZXNlbnQpIHsKICAgIGNvbnN0IFZJUyA9IFsnVmlld3MnLCAnSW1wcmVzc2lvbnMnXTsKICAgIGNvbnN0IEVORyA9IFsnRW5nYWdlbWVudCcsICdFbmdhZ2VtZW50cycsICdJbnRlcmFjdGlvbnMnLCAnUmVhY3Rpb25zJ107CiAgICBjb25zdCBpbXByb3ZlZCA9IFtdOyBjb25zdCBkZWNsaW5lZCA9IFtdOwogICAgcHJlc2VudC5mb3JFYWNoKChwKSA9PiB7CiAgICAgIGNvbnN0IHNpZ25hbHMgPSBwLnJvd3MuZmlsdGVyKChyKSA9PiAoVklTLmluY2x1ZGVzKHIubGFiZWwpIHx8IEVORy5pbmNsdWRlcyhyLmxhYmVsKSkgJiYgIXIubmEgJiYgci5jaGFuZ2UgIT09IDApOwogICAgICBpZiAoIXNpZ25hbHMubGVuZ3RoKSByZXR1cm47CiAgICAgIGNvbnN0IHVwID0gc2lnbmFscy5maWx0ZXIoKHIpID0+IHIuY2hhbmdlID4gMCkubGVuZ3RoOwogICAgICBjb25zdCBkbiA9IHNpZ25hbHMuZmlsdGVyKChyKSA9PiByLmNoYW5nZSA8IDApLmxlbmd0aDsKICAgICAgaWYgKHVwID4gZG4pIGltcHJvdmVkLnB1c2gocC5sYWJlbCk7CiAgICAgIGVsc2UgaWYgKGRuID4gdXApIGRlY2xpbmVkLnB1c2gocC5sYWJlbCk7CiAgICB9KTsKCiAgICBjb25zdCBsaW5lcyA9IFtdOwogICAgaWYgKGltcHJvdmVkLmxlbmd0aCB8fCBkZWNsaW5lZC5sZW5ndGgpIHsKICAgICAgbGV0IHMgPSAnJzsKICAgICAgaWYgKGltcHJvdmVkLmxlbmd0aCkgcyArPSBsaXN0KGltcHJvdmVkKSArICcgaW1wcm92ZWQgaW4gdmlzaWJpbGl0eSBhbmQgZW5nYWdlbWVudCc7CiAgICAgIGlmIChpbXByb3ZlZC5sZW5ndGggJiYgZGVjbGluZWQubGVuZ3RoKSBzICs9ICcsIHdoaWxlICc7CiAgICAgIGlmIChkZWNsaW5lZC5sZW5ndGgpIHMgKz0gbGlzdChkZWNsaW5lZCkgKyAnIGRlY2xpbmVkJzsKICAgICAgbGluZXMucHVzaChzLmNoYXJBdCgwKS50b1VwcGVyQ2FzZSgpICsgcy5zbGljZSgxKSArICcuJyk7CiAgICB9IGVsc2UgewogICAgICBsaW5lcy5wdXNoKCdXZWVrLW92ZXItd2VlayB2aXNpYmlsaXR5IGFuZCBlbmdhZ2VtZW50IHdlcmUgYnJvYWRseSBmbGF0IGFjcm9zcyBwbGF0Zm9ybXMg4oCUIG5vIG1ldHJpYyBtb3ZlZCBkZWNpc2l2ZWx5LicpOwogICAgfQoKICAgIGlmIChtb2RlbC5mb2xsb3dlcnNHcm93dGgubGVuZ3RoKSB7CiAgICAgIGNvbnN0IG5ldCA9IG1vZGVsLmZvbGxvd2Vyc0dyb3d0aC5yZWR1Y2UoKGFjYywgcikgPT4gYWNjICsgKHIubmEgPyAwIDogci5jaGFuZ2UpLCAwKTsKICAgICAgY29uc3QgdXAgPSBtb2RlbC5mb2xsb3dlcnNHcm93dGguZmlsdGVyKChyKSA9PiAhci5uYSAmJiByLmNoYW5nZSA+IDApLm1hcCgocikgPT4gci5sYWJlbCk7CiAgICAgIGNvbnN0IGRuID0gbW9kZWwuZm9sbG93ZXJzR3Jvd3RoLmZpbHRlcigocikgPT4gIXIubmEgJiYgci5jaGFuZ2UgPCAwKS5tYXAoKHIpID0+IHIubGFiZWwpOwogICAgICBsZXQgZiA9ICdGb2xsb3dlcnM6IG5ldCAnICsgKG5ldCA+IDAgPyAnKycgOiBuZXQgPCAwID8gJ+KIkicgOiAnJykgKyBGb3JtYXQubnVtYmVyKE1hdGguYWJzKG5ldCkpICsgJyBhY3Jvc3MgdHJhY2tlZCBwbGF0Zm9ybXMnOwogICAgICBpZiAodXAubGVuZ3RoKSBmICs9ICcgKHVwIG9uICcgKyBsaXN0KHVwKSArIChkbi5sZW5ndGggPyAnOyBkb3duIG9uICcgKyBsaXN0KGRuKSA6ICcnKSArICcpJzsKICAgICAgZWxzZSBpZiAoZG4ubGVuZ3RoKSBmICs9ICcgKGRvd24gb24gJyArIGxpc3QoZG4pICsgJyknOwogICAgICBmICs9ICcuIFNlZSB0aGUgRm9sbG93ZXJzIEdyb3d0aCBzZWN0aW9uIGZvciB0aGUgcGVyLXBsYXRmb3JtIGJyZWFrZG93bi4nOwogICAgICBsaW5lcy5wdXNoKGYpOwogICAgfQogICAgcmV0dXJuIGxpbmVzOwogIH0KCiAgZnVuY3Rpb24gcmVwb3J0Qm9keShtb2RlbCkgewogICAgQ0hBUlRfSURTLmZvckVhY2goKGlkKSA9PiBDaGFydHMuZGVzdHJveShpZCkpOwogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdyZXBvcnRSZXN1bHRzJyk7CiAgICBpZiAoIXdyYXApIHJldHVybjsKICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CgogICAgY29uc3QgaGVhZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgaGVhZC5jbGFzc05hbWUgPSAnY2FyZCByZXBvcnQtZG9jLWhlYWQnOwogICAgaGVhZC5hcHBlbmQoCiAgICAgIHRleHRFbCgnZGl2JywgJ1dlZWtseSBTb2NpYWwgTWVkaWEgUGVyZm9ybWFuY2UgUmVwb3J0JywgJ3JlcG9ydC10aXRsZScpLAogICAgICB0ZXh0RWwoJ2RpdicsIG1vZGVsLmN1ckxhYmVsICsgJywgJyArIHllYXJPZihtb2RlbC5jdXIudG8pICsgJ+KAg3Zz4oCDJyArIG1vZGVsLnByZXZMYWJlbCArICcsICcgKyB5ZWFyT2YobW9kZWwucHJldi50byksICdyZXBvcnQtcmFuZ2UnKSwKICAgICAgdGV4dEVsKCdkaXYnLCAnR2VuZXJhdGVkICcgKyBuZXcgRGF0ZSgpLnRvTG9jYWxlU3RyaW5nKCdlbi1VUycsIHsgZGF0ZVN0eWxlOiAnbWVkaXVtJywgdGltZVN0eWxlOiAnc2hvcnQnIH0pLCAncmVwb3J0LWdlbmVyYXRlZCcpLAogICAgKTsKICAgIHdyYXAuYXBwZW5kQ2hpbGQoaGVhZCk7CgogICAgY29uc3QgcHJlc2VudCA9IG1vZGVsLnBsYXRmb3Jtcy5maWx0ZXIoKHApID0+IHAucHJlc2VudCk7CiAgICBpZiAoIXByZXNlbnQubGVuZ3RoICYmICFtb2RlbC5mb2xsb3dlcnNHcm93dGgubGVuZ3RoKSB7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQoZW1wdHlTdGF0ZSh7CiAgICAgICAgaWNvbjogJ2ZpbGUtdGV4dCcsCiAgICAgICAgdGl0bGU6ICdObyBkYXRhIGZvciB0aGUgc2VsZWN0ZWQgd2Vla3MnLAogICAgICAgIG1lc3NhZ2U6ICdObyBwb3N0cyBvciBGb2xsb3dlcnMgRGF0YSBSZWNvcmQgZW50cmllcyBmYWxsIGluIHRoaXMgcmFuZ2UuIFBpY2sgYW5vdGhlciB3ZWVrLCBvciB1cGxvYWQgZGF0YS4nLAogICAgICAgIGFjdGlvbkxhYmVsOiAnVXBsb2FkIGRhdGEnLAogICAgICAgIG9uQWN0aW9uOiAoKSA9PiB7IGNvbnN0IGIgPSBkb2N1bWVudC5xdWVyeVNlbGVjdG9yKCcudGFiLWJ0bltkYXRhLXRhYj0idXBsb2FkIl0nKTsgaWYgKGIpIGIuY2xpY2soKTsgfSwKICAgICAgfSkpOwogICAgICByZXR1cm47CiAgICB9CgogICAgd3JhcC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsICdPdmVyYWxsIFN1bW1hcnknLCAnc2VjdGlvbi10aXRsZScpKTsKICAgIGNvbnN0IHN1bSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgc3VtLmNsYXNzTmFtZSA9ICdjYXJkIHJlcG9ydC1zdW1tYXJ5JzsKICAgIGJ1aWxkU3VtbWFyeShtb2RlbCwgcHJlc2VudCkuZm9yRWFjaCgocCkgPT4geyBjb25zdCBlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3AnKTsgZWwudGV4dENvbnRlbnQgPSBwOyBzdW0uYXBwZW5kQ2hpbGQoZWwpOyB9KTsKICAgIHdyYXAuYXBwZW5kQ2hpbGQoc3VtKTsKCiAgICBwcmVzZW50LmZvckVhY2goKHApID0+IHsKICAgICAgd3JhcC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIHAubGFiZWwsICdzZWN0aW9uLXRpdGxlJykpOwogICAgICBjb25zdCBjYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIGNhcmQuY2xhc3NOYW1lID0gJ2NhcmQgcmVwb3J0LXBsYXRmb3JtLWNhcmQnOwogICAgICBjb25zdCBjdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBjdy5jbGFzc05hbWUgPSAnY2hhcnQtd3JhcCB0YWxsJzsKICAgICAgY29uc3QgY2lkID0gJ3JlcG9ydENoYXJ0XycgKyBwLmlkOwogICAgICBjdy5pbm5lckhUTUwgPSAnPGNhbnZhcyBpZD0iJyArIGNpZCArICciPjwvY2FudmFzPic7CiAgICAgIGNhcmQuYXBwZW5kQ2hpbGQoY3cpOwogICAgICB3cmFwLmFwcGVuZENoaWxkKGNhcmQpOwogICAgICB3cmFwLmFwcGVuZENoaWxkKGRldGFpbFRhYmxlKHAucm93cywgJ0NhdGVnb3J5JywgbW9kZWwucHJldkxhYmVsLCBtb2RlbC5jdXJMYWJlbCkpOwogICAgICBkcmF3R3JvdXBlZENoYXJ0KGNpZCwgcC5yb3dzLm1hcCgocikgPT4gci5sYWJlbCksIHAucm93cywgbW9kZWwucHJldkxhYmVsLCBtb2RlbC5jdXJMYWJlbCk7CiAgICB9KTsKCiAgICBpZiAobW9kZWwuZm9sbG93ZXJzR3Jvd3RoLmxlbmd0aCkgewogICAgICB3cmFwLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgJ0ZvbGxvd2VycyBHcm93dGgnLCAnc2VjdGlvbi10aXRsZScpKTsKICAgICAgY29uc3QgY2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBjYXJkLmNsYXNzTmFtZSA9ICdjYXJkIHJlcG9ydC1wbGF0Zm9ybS1jYXJkJzsKICAgICAgY29uc3QgY3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgY3cuY2xhc3NOYW1lID0gJ2NoYXJ0LXdyYXAgdGFsbCc7CiAgICAgIGN3LmlubmVySFRNTCA9ICc8Y2FudmFzIGlkPSJyZXBvcnRGb2xsb3dlcnNDaGFydCI+PC9jYW52YXM+JzsKICAgICAgY2FyZC5hcHBlbmRDaGlsZChjdyk7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQoY2FyZCk7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQoZGV0YWlsVGFibGUobW9kZWwuZm9sbG93ZXJzR3Jvd3RoLCAnUGxhdGZvcm0nLCBtb2RlbC5wcmV2TGFiZWwsIG1vZGVsLmN1ckxhYmVsKSk7CiAgICAgIGRyYXdHcm91cGVkQ2hhcnQoJ3JlcG9ydEZvbGxvd2Vyc0NoYXJ0JywgbW9kZWwuZm9sbG93ZXJzR3Jvd3RoLm1hcCgocikgPT4gci5sYWJlbCksIG1vZGVsLmZvbGxvd2Vyc0dyb3d0aCwgbW9kZWwucHJldkxhYmVsLCBtb2RlbC5jdXJMYWJlbCk7CiAgICB9CiAgfQoKICAvKiAtLS0tLS0tLS0tIGRhdGEgZmV0Y2ggLS0tLS0tLS0tLSAqLwogIGFzeW5jIGZ1bmN0aW9uIGdlbmVyYXRlKCkgewogICAgaWYgKGJ1c3kpIHJldHVybjsKICAgIGJ1c3kgPSB0cnVlOwogICAgY29uc3QgcmVzdWx0cyA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdyZXBvcnRSZXN1bHRzJyk7CiAgICBpZiAocmVzdWx0cykgewogICAgICByZXN1bHRzLmlubmVySFRNTCA9ICcnOwogICAgICByZXN1bHRzLmFwcGVuZENoaWxkKHNrZWxldG9uQ2hhcnQoKSk7CiAgICAgIHJlc3VsdHMuYXBwZW5kQ2hpbGQoc2tlbGV0b25Sb3dzKDYpKTsKICAgICAgcmVzdWx0cy5hcHBlbmRDaGlsZChza2VsZXRvbkNoYXJ0KCkpOwogICAgfQogICAgdHJ5IHsKICAgICAgY29uc3QgeyBjdXIsIHByZXYgfSA9IHBlcmlvZHMoKTsKICAgICAgY29uc3QgW2NtcCwgZm9sbG93ZXJSb3dzXSA9IGF3YWl0IFByb21pc2UuYWxsKFsKICAgICAgICBBcGkuY29tcGFyZSh7IGZyb21BOiBjdXIuZnJvbSwgdG9BOiBjdXIudG8sIGZyb21COiBwcmV2LmZyb20sIHRvQjogcHJldi50byB9KSwKICAgICAgICBBcGkubGlzdEZvbGxvd2Vycyh7fSksCiAgICAgIF0pOwogICAgICBsYXN0TW9kZWwgPSBidWlsZE1vZGVsKGNtcCwgZm9sbG93ZXJSb3dzLCBjdXIsIHByZXYpOwogICAgICByZXBvcnRCb2R5KGxhc3RNb2RlbCk7CiAgICB9IGNhdGNoIChlcnIpIHsKICAgICAgaWYgKHJlc3VsdHMpIHsKICAgICAgICByZXN1bHRzLmlubmVySFRNTCA9ICcnOwogICAgICAgIHJlc3VsdHMuYXBwZW5kQ2hpbGQoZW1wdHlTdGF0ZSh7CiAgICAgICAgICBpY29uOiAnYWxlcnQtdHJpYW5nbGUnLAogICAgICAgICAgdGl0bGU6ICdDb3VsZCBub3QgZ2VuZXJhdGUgdGhlIHJlcG9ydCcsCiAgICAgICAgICBtZXNzYWdlOiBlcnIubWVzc2FnZSB8fCAnU29tZXRoaW5nIHdlbnQgd3Jvbmcgd2hpbGUgZmV0Y2hpbmcgdGhlIGRhdGEuJywKICAgICAgICB9KSk7CiAgICAgIH0KICAgIH0gZmluYWxseSB7CiAgICAgIGJ1c3kgPSBmYWxzZTsKICAgIH0KICB9CgogIC8qIC0tLS0tLS0tLS0gY29udHJvbHMgLS0tLS0tLS0tLSAqLwogIGZ1bmN0aW9uIGxhYmVsZWQobGFiZWwsIGVsKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB3cmFwLmNsYXNzTmFtZSA9ICdmaWVsZC1pbmxpbmUnOwogICAgd3JhcC5hcHBlbmQodGV4dEVsKCdsYWJlbCcsIGxhYmVsKSwgZWwpOwogICAgcmV0dXJuIHdyYXA7CiAgfQoKICBmdW5jdGlvbiBzaGVsbCgpIHsKICAgIHJvb3QuaW5uZXJIVE1MID0gJyc7CgogICAgY29uc3QgY29udHJvbHMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNvbnRyb2xzLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGNvbnN0IGNSb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNSb3cuY2xhc3NOYW1lID0gJ3JlcG9ydC1jb250cm9scyc7CgogICAgY29uc3QgZnJvbUlucHV0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsKICAgIGZyb21JbnB1dC50eXBlID0gJ2RhdGUnOyBmcm9tSW5wdXQudmFsdWUgPSByYW5nZVN0YXJ0OwogICAgY29uc3QgdG9JbnB1dCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7CiAgICB0b0lucHV0LnR5cGUgPSAnZGF0ZSc7IHRvSW5wdXQudmFsdWUgPSByYW5nZUVuZDsKCiAgICBjb25zdCByZXNvbHZlZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcmVzb2x2ZWQuc3R5bGUuY3NzVGV4dCA9ICdtYXJnaW4tdG9wOjEwcHg7Zm9udC1zaXplOjEycHg7Y29sb3I6dmFyKC0tdGV4dC1tdXRlZCknOwogICAgZnVuY3Rpb24gdXBkYXRlUmVzb2x2ZWQoKSB7CiAgICAgIGNvbnN0IHMgPSBmcm9tSW5wdXQudmFsdWU7IGNvbnN0IGUgPSB0b0lucHV0LnZhbHVlIHx8IHM7CiAgICAgIGlmICghcyB8fCAhZSB8fCBlIDwgcykgeyByZXNvbHZlZC50ZXh0Q29udGVudCA9ICdQaWNrIGEgdmFsaWQgd2VlayDigJQgdGhlIFRvIGRhdGUgbXVzdCBiZSBvbiBvciBhZnRlciB0aGUgRnJvbSBkYXRlLic7IHJldHVybjsgfQogICAgICBjb25zdCBsZW4gPSBkYXlzSW5jbHVzaXZlKHMsIGUpOwogICAgICByZXNvbHZlZC50ZXh0Q29udGVudCA9ICdDdXJyZW50IHdlZWs6ICcgKyBGb3JtYXQuZGF0ZShzKSArICcg4oCTICcgKyBGb3JtYXQuZGF0ZShlKQogICAgICAgICsgJyAgICAgwrcgICAgIFByZXZpb3VzIHdlZWs6ICcgKyBGb3JtYXQuZGF0ZShhZGREYXlzKHMsIC1sZW4pKSArICcg4oCTICcgKyBGb3JtYXQuZGF0ZShhZGREYXlzKHMsIC0xKSk7CiAgICB9CiAgICBmcm9tSW5wdXQuYWRkRXZlbnRMaXN0ZW5lcignY2hhbmdlJywgdXBkYXRlUmVzb2x2ZWQpOwogICAgdG9JbnB1dC5hZGRFdmVudExpc3RlbmVyKCdjaGFuZ2UnLCB1cGRhdGVSZXNvbHZlZCk7CiAgICB1cGRhdGVSZXNvbHZlZCgpOwoKICAgIGNSb3cuYXBwZW5kKGxhYmVsZWQoJ1dlZWsgZnJvbScsIGZyb21JbnB1dCksIGxhYmVsZWQoJ1dlZWsgdG8nLCB0b0lucHV0KSk7CiAgICBjb250cm9scy5hcHBlbmRDaGlsZChjUm93KTsKICAgIGNvbnRyb2xzLmFwcGVuZENoaWxkKHJlc29sdmVkKTsKCiAgICBjb25zdCBhY3Rpb25zID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBhY3Rpb25zLmNsYXNzTmFtZSA9ICdyZXBvcnQtYWN0aW9ucyc7CiAgICBjb25zdCBnZW5CdG4gPSBpY29uQnRuKCdidG4gcHJpbWFyeScsICdzcGFya2xlcycsICdHZW5lcmF0ZSBSZXBvcnQnKTsKICAgIGdlbkJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsKICAgICAgY29uc3QgcyA9IGZyb21JbnB1dC52YWx1ZTsgY29uc3QgZSA9IHRvSW5wdXQudmFsdWUgfHwgczsKICAgICAgaWYgKCFzIHx8ICFlIHx8IGUgPCBzKSB7IFRvYXN0LnNob3coJ1BpY2sgYSB2YWxpZCB3ZWVrIOKAlCB0aGUgVG8gZGF0ZSBtdXN0IGJlIG9uIG9yIGFmdGVyIHRoZSBGcm9tIGRhdGUuJywgJ2Vycm9yJyk7IHJldHVybjsgfQogICAgICByYW5nZVN0YXJ0ID0gczsgcmFuZ2VFbmQgPSBlOwogICAgICBnZW5lcmF0ZSgpOwogICAgfSk7CiAgICBjb25zdCByZWZyZXNoQnRuID0gaWNvbkJ0bignYnRuJywgJ3JlZnJlc2gtY3cnLCAnUmVmcmVzaCBEYXRhJyk7CiAgICByZWZyZXNoQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gZ2VuZXJhdGUoKSk7CiAgICBjb25zdCBwcmludEJ0biA9IGljb25CdG4oJ2J0bicsICdwcmludGVyJywgJ1ByaW50IC8gRXhwb3J0IFJlcG9ydCcpOwogICAgcHJpbnRCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7CiAgICAgIGlmICghbGFzdE1vZGVsKSB7IFRvYXN0LnNob3coJ0dlbmVyYXRlIHRoZSByZXBvcnQgZmlyc3QuJywgJ2Vycm9yJyk7IHJldHVybjsgfQogICAgICB3aW5kb3cucHJpbnQoKTsKICAgIH0pOwogICAgYWN0aW9ucy5hcHBlbmQoZ2VuQnRuLCByZWZyZXNoQnRuLCBwcmludEJ0bik7CiAgICBjb250cm9scy5hcHBlbmRDaGlsZChhY3Rpb25zKTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQoY29udHJvbHMpOwoKICAgIGNvbnN0IHJlc3VsdHMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHJlc3VsdHMuaWQgPSAncmVwb3J0UmVzdWx0cyc7CiAgICByb290LmFwcGVuZENoaWxkKHJlc3VsdHMpOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gcmVuZGVyKCkgewogICAgcm9vdCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd2aWV3LXJlcG9ydCcpOwogICAgc2hlbGwoKTsKICAgIGlmIChsYXN0TW9kZWwpIHJlcG9ydEJvZHkobGFzdE1vZGVsKTsKICAgIGVsc2UgYXdhaXQgZ2VuZXJhdGUoKTsKICB9CgogIHJldHVybiB7IHJlbmRlciB9Owp9KSgpOwoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIEFwcCBib290c3RyYXA6IHRhYiByb3V0aW5nLCBmaWx0ZXIgYmFyIHdpcmluZywgdGhlbWUgdG9nZ2xlLgogICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KKCgpID0+IHsKICBjb25zdCBWSUVXUyA9IHsKICAgIGRhc2hib2FyZDogRGFzaGJvYXJkLAogICAgcmVjb3JkczogUmVjb3JkcywKICAgIGZvbGxvd2VyczogRm9sbG93ZXJzLAogICAgY29tcGFyaXNvbjogQ29tcGFyaXNvbiwKICAgIHJlcG9ydDogUmVwb3J0LAogICAgdXBsb2FkOiBVcGxvYWQsCiAgICBoaXN0b3J5OiBIaXN0b3J5LAogIH07CgogIGxldCBhY3RpdmVUYWIgPSAnZGFzaGJvYXJkJzsKCiAgZnVuY3Rpb24gc3dpdGNoVGFiKHRhYikgewogICAgYWN0aXZlVGFiID0gdGFiOwogICAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnLnRhYi1idG4nKS5mb3JFYWNoKChidG4pID0+IHsKICAgICAgY29uc3QgaXNBY3RpdmUgPSBidG4uZGF0YXNldC50YWIgPT09IHRhYjsKICAgICAgYnRuLmNsYXNzTGlzdC50b2dnbGUoJ2lzLWFjdGl2ZScsIGlzQWN0aXZlKTsKICAgICAgYnRuLnNldEF0dHJpYnV0ZSgnYXJpYS1zZWxlY3RlZCcsIFN0cmluZyhpc0FjdGl2ZSkpOwogICAgfSk7CiAgICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcudmlldycpLmZvckVhY2goKHZpZXcpID0+IHsKICAgICAgdmlldy5jbGFzc0xpc3QudG9nZ2xlKCdpcy1hY3RpdmUnLCB2aWV3LmlkID09PSBgdmlldy0ke3RhYn1gKTsKICAgIH0pOwogICAgLy8gRmlsdGVycyBhcHBseSB0byBEYXNoYm9hcmQgYW5kIERhdGEgUmVjb3JkcyAoQ29tcGFyaXNvbnMgaGFzIGl0cyBvd24gcmFuZ2UgY29udHJvbHMpLgogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZpbHRlckJhcicpLnN0eWxlLmRpc3BsYXkgPSAodGFiID09PSAnZGFzaGJvYXJkJyB8fCB0YWIgPT09ICdyZWNvcmRzJykgPyAnZmxleCcgOiAnbm9uZSc7CiAgICByZW5kZXJBY3RpdmVWaWV3KCk7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJBY3RpdmVWaWV3KCkgewogICAgY29uc3QgdmlldyA9IFZJRVdTW2FjdGl2ZVRhYl07CiAgICBpZiAodmlldyAmJiB2aWV3LnJlbmRlcikgdmlldy5yZW5kZXIoKTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIGxvYWRGaWx0ZXJPcHRpb25zKCkgewogICAgY29uc3Qgb3B0aW9ucyA9IGF3YWl0IEFwaS5maWx0ZXJPcHRpb25zKCk7CiAgICB3aW5kb3cuX19maWx0ZXJPcHRpb25zQ2FjaGUgPSBvcHRpb25zOwoKICAgIGNvbnN0IHBsYXRmb3JtU2VsID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZpbHRlclBsYXRmb3JtJyk7CiAgICBwbGF0Zm9ybVNlbC5sZW5ndGggPSAxOwogICAgb3B0aW9ucy5wbGF0Zm9ybXMuZm9yRWFjaCgocCkgPT4gewogICAgICBjb25zdCBvcHQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsKICAgICAgb3B0LnZhbHVlID0gcC5pZDsKICAgICAgb3B0LnRleHRDb250ZW50ID0gcC5sYWJlbDsKICAgICAgcGxhdGZvcm1TZWwuYXBwZW5kQ2hpbGQob3B0KTsKICAgIH0pOwoKICAgIGNvbnN0IGNhbXBhaWduU2VsID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZpbHRlckNhbXBhaWduJyk7CiAgICBjYW1wYWlnblNlbC5sZW5ndGggPSAxOwogICAgb3B0aW9ucy5jYW1wYWlnblR5cGVzLmZvckVhY2goKGMpID0+IHsKICAgICAgY29uc3Qgb3B0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7CiAgICAgIG9wdC52YWx1ZSA9IGM7CiAgICAgIG9wdC50ZXh0Q29udGVudCA9IGM7CiAgICAgIGNhbXBhaWduU2VsLmFwcGVuZENoaWxkKG9wdCk7CiAgICB9KTsKCiAgICBjb25zdCBjb250ZW50U2VsID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZpbHRlckNvbnRlbnRUeXBlJyk7CiAgICBjb250ZW50U2VsLmxlbmd0aCA9IDE7CiAgICBvcHRpb25zLmNvbnRlbnRUeXBlcy5mb3JFYWNoKChjKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOwogICAgICBvcHQudmFsdWUgPSBjOwogICAgICBvcHQudGV4dENvbnRlbnQgPSBjOwogICAgICBjb250ZW50U2VsLmFwcGVuZENoaWxkKG9wdCk7CiAgICB9KTsKICB9CgogIGZ1bmN0aW9uIHdpcmVGaWx0ZXJCYXIoKSB7CiAgICBjb25zdCBkYXRlRnJvbSA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJEYXRlRnJvbScpOwogICAgY29uc3QgZGF0ZVRvID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZpbHRlckRhdGVUbycpOwogICAgY29uc3QgcGxhdGZvcm0gPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmlsdGVyUGxhdGZvcm0nKTsKICAgIGNvbnN0IGNhbXBhaWduID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZpbHRlckNhbXBhaWduJyk7CiAgICBjb25zdCBjb250ZW50VHlwZSA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJDb250ZW50VHlwZScpOwogICAgY29uc3QgZiA9IFN0YXRlLmdldEZpbHRlcnMoKTsKICAgIGRhdGVGcm9tLnZhbHVlID0gZi5kYXRlRnJvbTsKICAgIGRhdGVUby52YWx1ZSA9IGYuZGF0ZVRvOwoKICAgIGZ1bmN0aW9uIGFwcGx5KCkgewogICAgICBTdGF0ZS5zZXRGaWx0ZXJzKHsKICAgICAgICBkYXRlRnJvbTogZGF0ZUZyb20udmFsdWUsCiAgICAgICAgZGF0ZVRvOiBkYXRlVG8udmFsdWUsCiAgICAgICAgcGxhdGZvcm06IHBsYXRmb3JtLnZhbHVlLAogICAgICAgIGNhbXBhaWduVHlwZTogY2FtcGFpZ24udmFsdWUsCiAgICAgICAgY29udGVudFR5cGU6IGNvbnRlbnRUeXBlLnZhbHVlLAogICAgICB9KTsKICAgIH0KICAgIFtkYXRlRnJvbSwgZGF0ZVRvLCBwbGF0Zm9ybSwgY2FtcGFpZ24sIGNvbnRlbnRUeXBlXS5mb3JFYWNoKChlbCkgPT4gZWwuYWRkRXZlbnRMaXN0ZW5lcignY2hhbmdlJywgYXBwbHkpKTsKCiAgICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcjZmlsdGVyUHJlc2V0cyBidXR0b24nKS5mb3JFYWNoKChidG4pID0+IHsKICAgICAgYnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gewogICAgICAgIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJyNmaWx0ZXJQcmVzZXRzIGJ1dHRvbicpLmZvckVhY2goKGIpID0+IGIuY2xhc3NMaXN0LnJlbW92ZSgnaXMtYWN0aXZlJykpOwogICAgICAgIGJ0bi5jbGFzc0xpc3QuYWRkKCdpcy1hY3RpdmUnKTsKICAgICAgICBjb25zdCBwcmVzZXQgPSBidG4uZGF0YXNldC5wcmVzZXQ7CiAgICAgICAgY29uc3QgdG9kYXkgPSBuZXcgRGF0ZSgpOwogICAgICAgIGNvbnN0IHRvID0gdG9kYXkudG9JU09TdHJpbmcoKS5zbGljZSgwLCAxMCk7CiAgICAgICAgbGV0IGZyb207CiAgICAgICAgaWYgKHByZXNldCA9PT0gJ2FsbCcpIHsKICAgICAgICAgIGNvbnN0IG1pbiA9ICh3aW5kb3cuX19maWx0ZXJPcHRpb25zQ2FjaGUgJiYgd2luZG93Ll9fZmlsdGVyT3B0aW9uc0NhY2hlLmRhdGVSYW5nZS5taW4pIHx8IHRvOwogICAgICAgICAgZnJvbSA9IG1pbjsKICAgICAgICB9IGVsc2UgewogICAgICAgICAgY29uc3QgZCA9IG5ldyBEYXRlKHRvZGF5KTsKICAgICAgICAgIGQuc2V0RGF0ZShkLmdldERhdGUoKSAtIChOdW1iZXIocHJlc2V0KSAtIDEpKTsKICAgICAgICAgIGZyb20gPSBkLnRvSVNPU3RyaW5nKCkuc2xpY2UoMCwgMTApOwogICAgICAgIH0KICAgICAgICBkYXRlRnJvbS52YWx1ZSA9IGZyb207CiAgICAgICAgZGF0ZVRvLnZhbHVlID0gdG87CiAgICAgICAgYXBwbHkoKTsKICAgICAgfSk7CiAgICB9KTsKCiAgICBTdGF0ZS5vbkNoYW5nZSgoKSA9PiB7CiAgICAgIGlmIChhY3RpdmVUYWIgPT09ICdkYXNoYm9hcmQnKSBEYXNoYm9hcmQucmVuZGVyKCk7CiAgICAgIGlmIChhY3RpdmVUYWIgPT09ICdyZWNvcmRzJykgUmVjb3Jkcy5yZW5kZXIoKTsKICAgIH0pOwogIH0KCiAgZnVuY3Rpb24gd2lyZVRhYnMoKSB7CiAgICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcudGFiLWJ0bicpLmZvckVhY2goKGJ0bikgPT4gewogICAgICBidG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBzd2l0Y2hUYWIoYnRuLmRhdGFzZXQudGFiKSk7CiAgICB9KTsKICB9CgogIHdpbmRvdy5hZGRFdmVudExpc3RlbmVyKCdscnM6ZGF0YS11cGRhdGVkJywgYXN5bmMgKCkgPT4gewogICAgYXdhaXQgbG9hZEZpbHRlck9wdGlvbnMoKTsKICAgIHJlbmRlckFjdGl2ZVZpZXcoKTsKICB9KTsKCiAgLy8gLS0tLS0tLS0tLSBBdXRoIHNjcmVlbiAtLS0tLS0tLS0tCiAgbGV0IGFwcEluaXRpYWxpemVkID0gZmFsc2U7CgogIGZ1bmN0aW9uIHNob3dBdXRoU2NyZWVuKCkgewogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2F1dGhTY3JlZW4nKS5zdHlsZS5kaXNwbGF5ID0gJ2ZsZXgnOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2FwcFNoZWxsJykuc3R5bGUuZGlzcGxheSA9ICdub25lJzsKICAgIGNvbnN0IGNvZGVJbnB1dCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhdXRoQ29kZScpOwogICAgY29kZUlucHV0LnZhbHVlID0gJyc7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXV0aEVycm9yJykudGV4dENvbnRlbnQgPSAnJzsKICAgIGNvZGVJbnB1dC5mb2N1cygpOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gc2hvd0FwcCgpIHsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhdXRoU2NyZWVuJykuc3R5bGUuZGlzcGxheSA9ICdub25lJzsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhcHBTaGVsbCcpLnN0eWxlLmRpc3BsYXkgPSAnJzsKICAgIGlmICghYXBwSW5pdGlhbGl6ZWQpIHsKICAgICAgYXBwSW5pdGlhbGl6ZWQgPSB0cnVlOwogICAgICB3aXJlVGFicygpOwogICAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbG9nb3V0QnRuJykuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBhc3luYyAoKSA9PiB7CiAgICAgICAgYXdhaXQgQXBpLmF1dGhMb2dvdXQoKTsKICAgICAgICBhcHBJbml0aWFsaXplZCA9IGZhbHNlOwogICAgICAgIHNob3dBdXRoU2NyZWVuKCk7CiAgICAgIH0pOwogICAgICBhd2FpdCBsb2FkRmlsdGVyT3B0aW9ucygpOwogICAgICB3aXJlRmlsdGVyQmFyKCk7CiAgICAgIHN3aXRjaFRhYignZGFzaGJvYXJkJyk7CiAgICB9IGVsc2UgewogICAgICBhd2FpdCBsb2FkRmlsdGVyT3B0aW9ucygpOwogICAgICByZW5kZXJBY3RpdmVWaWV3KCk7CiAgICB9CiAgfQoKICBhc3luYyBmdW5jdGlvbiBzdWJtaXRBdXRoKCkgewogICAgY29uc3QgZXJyb3JFbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhdXRoRXJyb3InKTsKICAgIGNvbnN0IGJ0biA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhdXRoU3VibWl0QnRuJyk7CiAgICBjb25zdCBjb2RlSW5wdXQgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXV0aENvZGUnKTsKICAgIGVycm9yRWwudGV4dENvbnRlbnQgPSAnJzsKICAgIGJ0bi5kaXNhYmxlZCA9IHRydWU7CiAgICBidG4udGV4dENvbnRlbnQgPSAnQ2hlY2tpbmfigKYnOwogICAgdHJ5IHsKICAgICAgYXdhaXQgQXBpLmF1dGhMb2dpbihjb2RlSW5wdXQudmFsdWUpOwogICAgICBhd2FpdCBzaG93QXBwKCk7CiAgICB9IGNhdGNoIChlcnIpIHsKICAgICAgZXJyb3JFbC50ZXh0Q29udGVudCA9IGVyci5tZXNzYWdlOwogICAgfSBmaW5hbGx5IHsKICAgICAgYnRuLmRpc2FibGVkID0gZmFsc2U7CiAgICAgIGJ0bi5pbm5lckhUTUwgPSAnPGkgZGF0YS1sdWNpZGU9ImFycm93LXJpZ2h0IiBzdHlsZT0id2lkdGg6MTRweDtoZWlnaHQ6MTRweDsiPjwvaT4gRW50ZXInOwogICAgfQogIH0KCiAgZnVuY3Rpb24gd2lyZUF1dGhGb3JtKCkgewogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2F1dGhTdWJtaXRCdG4nKS5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIHN1Ym1pdEF1dGgpOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2F1dGhDb2RlJykuYWRkRXZlbnRMaXN0ZW5lcigna2V5ZG93bicsIChlKSA9PiB7IGlmIChlLmtleSA9PT0gJ0VudGVyJykgc3VibWl0QXV0aCgpOyB9KTsKICB9CgogIHdpbmRvdy5hZGRFdmVudExpc3RlbmVyKCdscnM6c2lnbmVkLW91dCcsICgpID0+IHsKICAgIGFwcEluaXRpYWxpemVkID0gZmFsc2U7CiAgICBzaG93QXV0aFNjcmVlbigpOwogIH0pOwoKICBhc3luYyBmdW5jdGlvbiBpbml0KCkgewogICAgYXBwbHlCcmFuZGluZygpOwogICAgd2lyZUF1dGhGb3JtKCk7CiAgICBjb25zdCB7IGF1dGhlbnRpY2F0ZWQgfSA9IGF3YWl0IEFwaS5hdXRoTWUoKTsKICAgIGlmIChhdXRoZW50aWNhdGVkKSBhd2FpdCBzaG93QXBwKCk7CiAgICBlbHNlIHNob3dBdXRoU2NyZWVuKCk7CiAgfQoKICAvLyBJY29ucyBhcmUgcGxhY2VkIGFzIDxpIGRhdGEtbHVjaWRlPSIuLi4iPiBwbGFjZWhvbGRlcnMgdGhyb3VnaG91dCB0aGUgZHluYW1pY2FsbHkKICAvLyByZW5kZXJlZCBVSTsgTHVjaWRlIHJlcGxhY2VzIGVhY2ggd2l0aCBhbiBpbmxpbmUgU1ZHLiBSYXRoZXIgdGhhbiByZW1lbWJlcmluZyB0byBjYWxsCiAgLy8gdGhpcyBhZnRlciBldmVyeSBzaW5nbGUgcmVuZGVyLCBvbmUgb2JzZXJ2ZXIgY2F0Y2hlcyBldmVyeSBET00gY2hhbmdlIHRoYXQgY291bGQgaGF2ZQogIC8vIGludHJvZHVjZWQgYSBuZXcgcGxhY2Vob2xkZXIuCiAgaWYgKHdpbmRvdy5sdWNpZGUpIHsKICAgIHdpbmRvdy5sdWNpZGUuY3JlYXRlSWNvbnMoKTsKICAgIC8vIGNyZWF0ZUljb25zKCkgcmVwbGFjZXMgPGkgZGF0YS1sdWNpZGU+IHBsYWNlaG9sZGVycyB3aXRoIDxzdmc+IOKAlCBpdHNlbGYgYSBET00KICAgIC8vIG11dGF0aW9uLiBXaXRob3V0IGRpc2Nvbm5lY3RpbmcgZmlyc3QsIHRoYXQgd3JpdGUgcmUtdHJpZ2dlcnMgdGhpcyBzYW1lIG9ic2VydmVyCiAgICAvLyBmb3JldmVyIChhbiBpbmZpbml0ZSBtdXRhdGUvb2JzZXJ2ZSBsb29wIHRoYXQgcGVncyB0aGUgQ1BVIGFuZCBjcmFzaGVzIHRoZSB0YWIpLgogICAgLy8gRGlzY29ubmVjdGluZyBiZWZvcmUgZWFjaCBwYXNzIGFuZCByZWNvbm5lY3RpbmcgYWZ0ZXIsIHBsdXMgYmF0Y2hpbmcgYnVyc3RzIG9mCiAgICAvLyBtdXRhdGlvbnMgaW50byBhIHNpbmdsZSBtaWNyb3Rhc2ssIGJyZWFrcyB0aGUgY3ljbGUuCiAgICBsZXQgaWNvbnNTY2hlZHVsZWQgPSBmYWxzZTsKICAgIGNvbnN0IGljb25PYnNlcnZlciA9IG5ldyBNdXRhdGlvbk9ic2VydmVyKCgpID0+IHsKICAgICAgaWYgKGljb25zU2NoZWR1bGVkKSByZXR1cm47CiAgICAgIGljb25zU2NoZWR1bGVkID0gdHJ1ZTsKICAgICAgcXVldWVNaWNyb3Rhc2soKCkgPT4gewogICAgICAgIGljb25zU2NoZWR1bGVkID0gZmFsc2U7CiAgICAgICAgaWNvbk9ic2VydmVyLmRpc2Nvbm5lY3QoKTsKICAgICAgICB3aW5kb3cubHVjaWRlLmNyZWF0ZUljb25zKCk7CiAgICAgICAgaWNvbk9ic2VydmVyLm9ic2VydmUoZG9jdW1lbnQuYm9keSwgeyBjaGlsZExpc3Q6IHRydWUsIHN1YnRyZWU6IHRydWUgfSk7CiAgICAgIH0pOwogICAgfSk7CiAgICBpY29uT2JzZXJ2ZXIub2JzZXJ2ZShkb2N1bWVudC5ib2R5LCB7IGNoaWxkTGlzdDogdHJ1ZSwgc3VidHJlZTogdHJ1ZSB9KTsKICB9CgogIGluaXQoKTsKfSkoKTsKPC9zY3JpcHQ+CjwvYm9keT4KPC9odG1sPgo=';
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
