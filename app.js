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
const INDEX_HTML_BASE64 = 'PCFkb2N0eXBlIGh0bWw+CjxodG1sIGxhbmc9ImVuIj4KPGhlYWQ+CjxtZXRhIGNoYXJzZXQ9IlVURi04IiAvPgo8bWV0YSBuYW1lPSJ2aWV3cG9ydCIgY29udGVudD0id2lkdGg9ZGV2aWNlLXdpZHRoLCBpbml0aWFsLXNjYWxlPTEuMCIgLz4KPHRpdGxlPkxSUyBBbmFseXRpY3MgRGFzaGJvYXJkPC90aXRsZT4KPGxpbmsgcmVsPSJpY29uIiB0eXBlPSJpbWFnZS9wbmciIGlkPSJmYXZpY29uTGluayIgLz4KPGxpbmsgcmVsPSJwcmVjb25uZWN0IiBocmVmPSJodHRwczovL2ZvbnRzLmdvb2dsZWFwaXMuY29tIiAvPgo8bGluayByZWw9InByZWNvbm5lY3QiIGhyZWY9Imh0dHBzOi8vZm9udHMuZ3N0YXRpYy5jb20iIGNyb3Nzb3JpZ2luIC8+CjxsaW5rIGhyZWY9Imh0dHBzOi8vZm9udHMuZ29vZ2xlYXBpcy5jb20vY3NzMj9mYW1pbHk9SW50ZXI6d2dodEA0MDA7NTAwOzYwMDs3MDA7ODAwJmRpc3BsYXk9c3dhcCIgcmVsPSJzdHlsZXNoZWV0IiAvPgo8c2NyaXB0IHNyYz0iaHR0cHM6Ly9jZG4uanNkZWxpdnIubmV0L25wbS9jaGFydC5qc0A0LjQuNC9kaXN0L2NoYXJ0LnVtZC5taW4uanMiPjwvc2NyaXB0Pgo8c2NyaXB0IHNyYz0iaHR0cHM6Ly9jZG4uanNkZWxpdnIubmV0L25wbS9jaGFydGpzLXBsdWdpbi1kYXRhbGFiZWxzQDIvZGlzdC9jaGFydGpzLXBsdWdpbi1kYXRhbGFiZWxzLm1pbi5qcyI+PC9zY3JpcHQ+CjxzY3JpcHQgc3JjPSJodHRwczovL2Nkbi5qc2RlbGl2ci5uZXQvbnBtL2x1Y2lkZUAwLjQ2Mi4wL2Rpc3QvdW1kL2x1Y2lkZS5taW4uanMiPjwvc2NyaXB0Pgo8c3R5bGU+Ci8qIC0tLS0tLS0tLS0gRGVzaWduIHRva2VuczogZGFyayBuYXZ5ICsgZ29sZCBicmFuZGVkIHRoZW1lIChzaW5nbGUsIHBlcm1hbmVudCDigJQgbm8gbGlnaHQgdmFyaWFudCkgLS0tLS0tLS0tLSAqLwo6cm9vdCB7CiAgY29sb3Itc2NoZW1lOiBkYXJrOwogIC0tZm9udC1zYW5zOiAnSW50ZXInLCAtYXBwbGUtc3lzdGVtLCBCbGlua01hY1N5c3RlbUZvbnQsICdTRiBQcm8gRGlzcGxheScsICdTZWdvZSBVSScsIFJvYm90bywgc2Fucy1zZXJpZjsKCiAgLS1wYWdlLXBsYW5lOiBsaW5lYXItZ3JhZGllbnQoMTgwZGVnLCAjMGEwZjFjIDAlLCAjMGQxNDI0IDEwMCUpOwogIC0tcGFnZS1wbGFuZS1zb2xpZDogIzBhMGYxYzsKICAtLXNpZGViYXItYmc6ICMwYjEyMjA7CiAgLS1zdXJmYWNlLTE6IHJnYmEoMjMsIDMxLCA1MSwgMC42Mik7IC8qIGdsYXNzOiBjYXJkcywgS1BJIHRpbGVzLCBmaWx0ZXIgYmFyICovCiAgLS1zdXJmYWNlLTI6IHJnYmEoMjU1LCAyNTUsIDI1NSwgMC4wNik7IC8qIGdsYXNzOiBpbnB1dHMsIG5lc3RlZCByb3dzLCBwaWxscyAqLwogIC0tc3VyZmFjZS1zb2xpZDogIzEzMWIyZTsKICAtLWdsYXNzLWJsdXI6IGJsdXIoMjBweCk7CiAgLS1ib3JkZXI6IHJnYmEoMjU1LCAyNTUsIDI1NSwgMC4wOSk7CiAgLS10ZXh0LXByaW1hcnk6ICNmNGY1Zjc7CiAgLS10ZXh0LXNlY29uZGFyeTogI2I4YmJjNDsKICAtLXRleHQtbXV0ZWQ6ICM4Mjg2OGY7CiAgLS1ncmlkbGluZTogcmdiYSgyNTUsIDI1NSwgMjU1LCAwLjA4KTsKICAtLWJhc2VsaW5lOiByZ2JhKDI1NSwgMjU1LCAyNTUsIDAuMik7CiAgLS1zdWNjZXNzLXRleHQ6ICMzNGM3NmY7CgogIC0tc3RhdHVzLWdvb2Q6ICMyZmI4NjI7CiAgLS1zdGF0dXMtd2FybmluZzogI2YwYTEzYTsKICAtLXN0YXR1cy1zZXJpb3VzOiAjZWM4MzVhOwogIC0tc3RhdHVzLWNyaXRpY2FsOiAjZTA2MDVmOwoKICAtLWFjY2VudC1nb2xkOiAjZjJiMzBlOyAvKiBMUlMgYnJhbmQgZ29sZCDigJQgYWN0aXZlIG5hdiBpdGVtLCBwcmltYXJ5IGludGVyYWN0aXZlIGFjY2VudCAqLwoKICAtLXNlcmllcy0xOiAjMzk4N2U1OyAvKiBmYWNlYm9vayAqLwogIC0tc2VyaWVzLTI6ICMwMDgzMDA7IC8qIGluc3RhZ3JhbSAqLwogIC0tc2VyaWVzLTM6ICNkNTUxODE7IC8qIHRpa3RvayAqLwogIC0tc2VyaWVzLTQ6ICNjOTg1MDA7IC8qIGxpbmtlZGluICovCiAgLS1zZXJpZXMtNTogIzE5OWU3MDsgLyogdGhyZWFkcyAqLwogIC0tc2VyaWVzLTY6ICNkOTU5MjY7IC8qIHlvdXR1YmUgKi8KICAtLXNlcmllcy03OiAjOTA4NWU5OyAvKiByZXNlcnZlZCAqLwogIC0tc2VyaWVzLTg6ICNlNjY3Njc7IC8qIHJlc2VydmVkICovCgogIC0tcmFkaXVzLXNtOiAxMHB4OwogIC0tcmFkaXVzLW1kOiAxNHB4OwogIC0tcmFkaXVzLWxnOiAxOHB4OwoKICAtLXNoYWRvdy1jYXJkOiAwIDFweCAycHggcmdiYSgwLDAsMCwwLjIpLCAwIDhweCAyNHB4IC0xMHB4IHJnYmEoMCwwLDAsMC41KTsKICAtLXNoYWRvdy1ob3ZlcjogMCA2cHggMTJweCAtMnB4IHJnYmEoMCwwLDAsMC4zKSwgMCAxOHB4IDQwcHggLTE0cHggcmdiYSgwLDAsMCwwLjYpOwogIC0tc2hhZG93LW1vZGFsOiAwIDI0cHggNjRweCAtMTJweCByZ2JhKDAsMCwwLDAuNyk7CiAgLS1lYXNlOiBjdWJpYy1iZXppZXIoMC40LCAwLCAwLjIsIDEpOwp9CgoqIHsgYm94LXNpemluZzogYm9yZGVyLWJveDsgfQpodG1sLCBib2R5IHsgaGVpZ2h0OiAxMDAlOyB9CmJvZHkgewogIG1hcmdpbjogMDsKICBmb250LWZhbWlseTogdmFyKC0tZm9udC1zYW5zKTsKICBiYWNrZ3JvdW5kOiB2YXIoLS1wYWdlLXBsYW5lKTsKICBiYWNrZ3JvdW5kLWF0dGFjaG1lbnQ6IGZpeGVkOwogIGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOwogIC13ZWJraXQtZm9udC1zbW9vdGhpbmc6IGFudGlhbGlhc2VkOwogIC1tb3otb3N4LWZvbnQtc21vb3RoaW5nOiBncmF5c2NhbGU7Cn0KYnV0dG9uLCBzZWxlY3QsIGlucHV0LCB0ZXh0YXJlYSB7IGZvbnQtZmFtaWx5OiBpbmhlcml0OyB9CmgxLCBoMiwgaDMsIGg0IHsgZm9udC13ZWlnaHQ6IDcwMDsgbGV0dGVyLXNwYWNpbmc6IC0wLjAxZW07IH0KCjo6c2VsZWN0aW9uIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSAzMCUsIHRyYW5zcGFyZW50KTsgfQoKLyogQ3VzdG9tIHNjcm9sbGJhciDigJQgdGhpbiwgdW5vYnRydXNpdmUsIGZpdHMgdGhlIGdsYXNzIGFlc3RoZXRpYyAqLwo6Oi13ZWJraXQtc2Nyb2xsYmFyIHsgd2lkdGg6IDEwcHg7IGhlaWdodDogMTBweDsgfQo6Oi13ZWJraXQtc2Nyb2xsYmFyLXRyYWNrIHsgYmFja2dyb3VuZDogdHJhbnNwYXJlbnQ7IH0KOjotd2Via2l0LXNjcm9sbGJhci10aHVtYiB7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS10ZXh0LW11dGVkKSA0MCUsIHRyYW5zcGFyZW50KTsgYm9yZGVyLXJhZGl1czogMjBweDsgYm9yZGVyOiAycHggc29saWQgdHJhbnNwYXJlbnQ7IGJhY2tncm91bmQtY2xpcDogcGFkZGluZy1ib3g7IH0KOjotd2Via2l0LXNjcm9sbGJhci10aHVtYjpob3ZlciB7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS10ZXh0LW11dGVkKSA2MCUsIHRyYW5zcGFyZW50KTsgYmFja2dyb3VuZC1jbGlwOiBwYWRkaW5nLWJveDsgfQoKLmFwcC1zaGVsbCB7IGhlaWdodDogMTAwdmg7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiByb3c7IG92ZXJmbG93OiBoaWRkZW47IH0KCi8qIC0tLS0tLS0tLS0gU2lkZWJhciAtLS0tLS0tLS0tICovCi5zaWRlYmFyIHsKICBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBnYXA6IDRweDsKICB3aWR0aDogMjQwcHg7IGZsZXg6IDAgMCBhdXRvOwogIHBhZGRpbmc6IDIwcHggMTRweDsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zaWRlYmFyLWJnKTsKICBib3JkZXItcmlnaHQ6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOwogIHotaW5kZXg6IDIwOwogIG92ZXJmbG93LXk6IGF1dG87Cn0KLnNpZGViYXItYnJhbmQgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDhweDsgd2hpdGUtc3BhY2U6IG5vd3JhcDsgcGFkZGluZzogNHB4IDEwcHggMjBweDsgfQouYnJhbmQtbG9nbyB7IGhlaWdodDogMjhweDsgd2lkdGg6IGF1dG87IGRpc3BsYXk6IGJsb2NrOyBmbGV4LXNocmluazogMDsgb2JqZWN0LWZpdDogY29udGFpbjsgfQouYnJhbmQtdGl0bGUgeyBmb250LXdlaWdodDogNjAwOyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgbGV0dGVyLXNwYWNpbmc6IC0wLjAxZW07IH0KCi50YWJzIHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiAycHg7IGZsZXg6IDE7IHBvc2l0aW9uOiByZWxhdGl2ZTsgfQoudGFiLWJ0biB7CiAgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiAxMHB4OwogIGJvcmRlcjogbm9uZTsgYmFja2dyb3VuZDogdHJhbnNwYXJlbnQ7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7CiAgcGFkZGluZzogMTBweCAxMnB4OyBib3JkZXItcmFkaXVzOiAxMHB4OyBjdXJzb3I6IHBvaW50ZXI7IGZvbnQtc2l6ZTogMTRweDsgZm9udC13ZWlnaHQ6IDUwMDsKICB3aGl0ZS1zcGFjZTogbm93cmFwOyBwb3NpdGlvbjogcmVsYXRpdmU7IHRleHQtYWxpZ246IGxlZnQ7IHdpZHRoOiAxMDAlOwogIHRyYW5zaXRpb246IGNvbG9yIDE4MG1zIHZhcigtLWVhc2UpLCBiYWNrZ3JvdW5kIDE4MG1zIHZhcigtLWVhc2UpOwp9Ci50YWItYnRuIHN2ZyB7IGZsZXgtc2hyaW5rOiAwOyBvcGFjaXR5OiAwLjg7IH0KLnRhYi1idG4uaXMtYWN0aXZlIHN2ZyB7IG9wYWNpdHk6IDE7IGNvbG9yOiB2YXIoLS1zZXJpZXMtMSk7IH0KLnRhYi1idG46aG92ZXIgeyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDglLCB0cmFuc3BhcmVudCk7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyB9Ci50YWItYnRuLmlzLWFjdGl2ZSB7CiAgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IGZvbnQtd2VpZ2h0OiA2MDA7CiAgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSAxNiUsIHRyYW5zcGFyZW50KTsKICBhbmltYXRpb246IHRhYkluZGljYXRvckluIDIyMG1zIHZhcigtLWVhc2UpOwp9CkBrZXlmcmFtZXMgdGFiSW5kaWNhdG9ySW4geyBmcm9tIHsgb3BhY2l0eTogMDsgdHJhbnNmb3JtOiB0cmFuc2xhdGVYKC00cHgpOyB9IHRvIHsgb3BhY2l0eTogMTsgdHJhbnNmb3JtOiB0cmFuc2xhdGVYKDApOyB9IH0KLnNpZGViYXItZm9vdGVyIHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiAxMHB4OyBwYWRkaW5nLXRvcDogMTRweDsgbWFyZ2luLXRvcDogMTRweDsgYm9yZGVyLXRvcDogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IH0KCi5zaWRlYmFyLXVzZXIgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDEwcHg7IGZvbnQtc2l6ZTogMTNweDsgfQoKLyogLS0tLS0tLS0tLSBBdXRoIHNjcmVlbiAtLS0tLS0tLS0tICovCi5hdXRoLXNjcmVlbiB7CiAgbWluLWhlaWdodDogMTAwdmg7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogY2VudGVyOwogIGJhY2tncm91bmQ6IHZhcigtLXBhZ2UtcGxhbmUpOyBwYWRkaW5nOiAyMHB4Owp9Ci5hdXRoLWNhcmQgewogIHdpZHRoOiAxMDAlOyBtYXgtd2lkdGg6IDQwMHB4OyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTEpOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOwogIGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtbGcpOyBwYWRkaW5nOiAzMnB4OyBib3gtc2hhZG93OiB2YXIoLS1zaGFkb3ctbW9kYWwpOwogIGFuaW1hdGlvbjogbW9kYWxQYW5lbEluIDI2MG1zIHZhcigtLWVhc2UpOwp9Ci5hdXRoLWJyYW5kIHsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA4cHg7IG1hcmdpbi1ib3R0b206IDIycHg7IH0KLmF1dGgtYnJhbmQgLmJyYW5kLXRpdGxlIHsgZm9udC13ZWlnaHQ6IDcwMDsgZm9udC1zaXplOiAxN3B4OyB9Ci5hdXRoLWJyYW5kIC5icmFuZC1sb2dvIHsgaGVpZ2h0OiAzNnB4OyB9Ci5hdXRoLWZvcm0geyBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBnYXA6IDE0cHg7IG1hcmdpbi10b3A6IDE2cHg7IH0KLmF1dGgtZm9ybSAuZm9ybS1maWVsZCBpbnB1dCB7IHdpZHRoOiAxMDAlOyB9Ci5hdXRoLWVycm9yIHsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IGZvbnQtc2l6ZTogMTJweDsgbWluLWhlaWdodDogMTZweDsgfQoKLyogLS0tLS0tLS0tLSBGaWx0ZXIgYmFyIC0tLS0tLS0tLS0gKi8KLmZpbHRlci1iYXIgewogIGRpc3BsYXk6IGZsZXg7IGZsZXgtd3JhcDogd3JhcDsgYWxpZ24taXRlbXM6IGVuZDsgZ2FwOiAxNnB4OwogIHBhZGRpbmc6IDE0cHggMjBweDsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgYm9yZGVyLWJvdHRvbTogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7CiAgcG9zaXRpb246IHN0aWNreTsgdG9wOiAwOyB6LWluZGV4OiAxOTsKfQouZmlsdGVyLWZpZWxkIHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiA1cHg7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgfQouZmlsdGVyLWZpZWxkIGxhYmVsIHsgZm9udC13ZWlnaHQ6IDYwMDsgfQouZmlsdGVyLXByZXNldHMgeyBmbGV4LWRpcmVjdGlvbjogcm93OyBnYXA6IDZweDsgfQouZmlsdGVyLXByZXNldHMgYnV0dG9uIHsKICBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTIpOyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOwogIGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBib3JkZXItcmFkaXVzOiAyMHB4OyBwYWRkaW5nOiA3cHggMTNweDsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNTAwOyBjdXJzb3I6IHBvaW50ZXI7CiAgdHJhbnNpdGlvbjogY29sb3IgMTgwbXMgdmFyKC0tZWFzZSksIGJhY2tncm91bmQgMTgwbXMgdmFyKC0tZWFzZSksIHRyYW5zZm9ybSAxNTBtcyB2YXIoLS1lYXNlKTsKfQouZmlsdGVyLXByZXNldHMgYnV0dG9uOmhvdmVyIHsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgtMXB4KTsgfQouZmlsdGVyLXByZXNldHMgYnV0dG9uOmFjdGl2ZSB7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgwKSBzY2FsZSgwLjk2KTsgfQouZmlsdGVyLXByZXNldHMgYnV0dG9uLmlzLWFjdGl2ZSB7IGJhY2tncm91bmQ6IHZhcigtLXNlcmllcy0xKTsgY29sb3I6ICNmZmY7IGJvcmRlci1jb2xvcjogdHJhbnNwYXJlbnQ7IGJveC1zaGFkb3c6IDAgNHB4IDE0cHggLTVweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDYwJSwgdHJhbnNwYXJlbnQpOyB9CgovKiAtLS0tLS0tLS0tIE1haW4gY29sdW1uIChzaXRzIGJlc2lkZSB0aGUgc2lkZWJhcjsgc2Nyb2xscyBpbmRlcGVuZGVudGx5IHNvIHRoZSBzaWRlYmFyIHN0YXlzIGZ1bGx5IHZpc2libGUpIC0tLS0tLS0tLS0gKi8KLm1haW4tY29sIHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZmxleDogMSAxIGF1dG87IG1pbi13aWR0aDogMDsgaGVpZ2h0OiAxMDAlOyBvdmVyZmxvdy15OiBhdXRvOyB9CgovKiAtLS0tLS0tLS0tIFZpZXcgYXJlYSAtLS0tLS0tLS0tICovCi52aWV3LWFyZWEgeyBmbGV4OiAxOyBwYWRkaW5nOiAyNHB4OyBtYXgtd2lkdGg6IDE0MDBweDsgd2lkdGg6IDEwMCU7IG1hcmdpbjogMCBhdXRvOyB9Ci52aWV3IHsgZGlzcGxheTogbm9uZTsgfQoudmlldy5pcy1hY3RpdmUgeyBkaXNwbGF5OiBibG9jazsgYW5pbWF0aW9uOiB2aWV3RmFkZUluIDI2MG1zIHZhcigtLWVhc2UpOyB9CkBrZXlmcmFtZXMgdmlld0ZhZGVJbiB7CiAgZnJvbSB7IG9wYWNpdHk6IDA7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSg2cHgpOyB9CiAgdG8geyBvcGFjaXR5OiAxOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMCk7IH0KfQoKLnNlY3Rpb24tdGl0bGUgeyBmb250LXNpemU6IDE2cHg7IGZvbnQtd2VpZ2h0OiA3MDA7IG1hcmdpbjogMzJweCAwIDE0cHg7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyBsZXR0ZXItc3BhY2luZzogLTAuMDFlbTsgfQouc2VjdGlvbi10aXRsZTpmaXJzdC1jaGlsZCB7IG1hcmdpbi10b3A6IDA7IH0KCi8qIC0tLS0tLS0tLS0gSW5wdXRzIOKAlCBvbmUgc2hhcmVkIGdsYXNzIHRyZWF0bWVudCBmb3IgZXZlcnkgdGV4dCBpbnB1dCwgc2VsZWN0LCBhbmQgZGF0ZSBwaWNrZXIgLS0tLS0tLS0tLSAqLwouZmlsdGVyLWZpZWxkIHNlbGVjdCwgLmZpbHRlci1maWVsZCBpbnB1dFt0eXBlPSJkYXRlIl0sCi5mb3JtLWZpZWxkIGlucHV0LCAuZm9ybS1maWVsZCBzZWxlY3QsIC5mb3JtLWZpZWxkIHRleHRhcmVhLAouZGFzaGJvYXJkLWNvbnRyb2xzIHNlbGVjdCwgLnJlY29yZHMtc2VhcmNoIGlucHV0LAouZmllbGQtaW5saW5lIHNlbGVjdCwgLmZpZWxkLWlubGluZSBpbnB1dCwKLmNvbmZsaWN0LXJvdyBzZWxlY3QsIC5jYXJkLWhlYWRlciBzZWxlY3QgewogIGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1zbSk7CiAgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0yKTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IGZvbnQtc2l6ZTogMTNweDsKICBwYWRkaW5nOiA4cHggMTJweDsgbWluLXdpZHRoOiAxNDBweDsKICB0cmFuc2l0aW9uOiBib3JkZXItY29sb3IgMTYwbXMgdmFyKC0tZWFzZSksIGJveC1zaGFkb3cgMTYwbXMgdmFyKC0tZWFzZSk7Cn0KLmZpbHRlci1maWVsZCBzZWxlY3Q6aG92ZXIsIC5maWx0ZXItZmllbGQgaW5wdXRbdHlwZT0iZGF0ZSJdOmhvdmVyLAouZm9ybS1maWVsZCBpbnB1dDpob3ZlciwgLmZvcm0tZmllbGQgc2VsZWN0OmhvdmVyLAouZGFzaGJvYXJkLWNvbnRyb2xzIHNlbGVjdDpob3ZlciwgLnJlY29yZHMtc2VhcmNoIGlucHV0OmhvdmVyLAouZmllbGQtaW5saW5lIHNlbGVjdDpob3ZlciwgLmZpZWxkLWlubGluZSBpbnB1dDpob3ZlciwKLmNvbmZsaWN0LXJvdyBzZWxlY3Q6aG92ZXIsIC5jYXJkLWhlYWRlciBzZWxlY3Q6aG92ZXIgewogIGJvcmRlci1jb2xvcjogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSAzNSUsIHZhcigtLWJvcmRlcikpOwp9Ci5maWx0ZXItZmllbGQgc2VsZWN0OmZvY3VzLCAuZmlsdGVyLWZpZWxkIGlucHV0W3R5cGU9ImRhdGUiXTpmb2N1cywKLmZvcm0tZmllbGQgaW5wdXQ6Zm9jdXMsIC5mb3JtLWZpZWxkIHNlbGVjdDpmb2N1cywgLmZvcm0tZmllbGQgdGV4dGFyZWE6Zm9jdXMsCi5kYXNoYm9hcmQtY29udHJvbHMgc2VsZWN0OmZvY3VzLCAucmVjb3Jkcy1zZWFyY2ggaW5wdXQ6Zm9jdXMsCi5maWVsZC1pbmxpbmUgc2VsZWN0OmZvY3VzLCAuZmllbGQtaW5saW5lIGlucHV0OmZvY3VzLAouY29uZmxpY3Qtcm93IHNlbGVjdDpmb2N1cywgLmNhcmQtaGVhZGVyIHNlbGVjdDpmb2N1cywKLmF1dGgtZm9ybSBpbnB1dDpmb2N1cyB7CiAgb3V0bGluZTogbm9uZTsgYm9yZGVyLWNvbG9yOiB2YXIoLS1zZXJpZXMtMSk7CiAgYm94LXNoYWRvdzogMCAwIDAgM3B4IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zZXJpZXMtMSkgMTglLCB0cmFuc3BhcmVudCk7Cn0KCi8qIC0tLS0tLS0tLS0gU3RhdCB0aWxlcyAtLS0tLS0tLS0tICovCi5zdGF0LWdyaWQgewogIGRpc3BsYXk6IGdyaWQ7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogcmVwZWF0KGF1dG8tZml0LCBtaW5tYXgoMTgwcHgsIDFmcikpOyBnYXA6IDE0cHg7Cn0KLnN0YXQtdGlsZSB7CiAgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLW1kKTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgcGFkZGluZzogMTZweCAxOHB4OyBib3gtc2hhZG93OiB2YXIoLS1zaGFkb3ctY2FyZCk7CiAgdHJhbnNpdGlvbjogdHJhbnNmb3JtIDIwMG1zIHZhcigtLWVhc2UpLCBib3gtc2hhZG93IDIwMG1zIHZhcigtLWVhc2UpOwogIGFuaW1hdGlvbjogY2FyZEluIDMyMG1zIHZhcigtLWVhc2UpIGJhY2t3YXJkczsKfQouc3RhdC10aWxlOmhvdmVyIHsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKC0zcHgpOyBib3gtc2hhZG93OiB2YXIoLS1zaGFkb3ctaG92ZXIpOyB9Ci5zdGF0LWxhYmVsIHsgZm9udC1zaXplOiAxMnB4OyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyBmb250LXdlaWdodDogNjAwOyB9Ci5zdGF0LXZhbHVlIHsgZm9udC1zaXplOiAyN3B4OyBmb250LXdlaWdodDogNzAwOyBtYXJnaW4tdG9wOiA1cHg7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyBsZXR0ZXItc3BhY2luZzogLTAuMDJlbTsgfQouc3RhdC1kZWx0YSB7IGZvbnQtc2l6ZTogMTJweDsgbWFyZ2luLXRvcDogN3B4OyBmb250LXdlaWdodDogNjAwOyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDRweDsgfQouc3RhdC1kZWx0YS51cCB7IGNvbG9yOiB2YXIoLS1zdWNjZXNzLXRleHQpOyB9Ci5zdGF0LWRlbHRhLmRvd24geyBjb2xvcjogdmFyKC0tc3RhdHVzLWNyaXRpY2FsKTsgfQouc3RhdC1kZWx0YS5mbGF0IHsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyB9Ci5zdGF0LWRlbHRhLnVwOjpiZWZvcmUgeyBjb250ZW50OiAn4oaRJzsgfQouc3RhdC1kZWx0YS5kb3duOjpiZWZvcmUgeyBjb250ZW50OiAn4oaTJzsgfQoKLyogLS0tLS0tLS0tLSBEYXNoYm9hcmQgS1BJIGdyaWQg4oCUIGNvbXBhY3QsIHNpbmdsZS1yb3ctb24tZGVza3RvcCBsYXlvdXQuCiAgIFNjb3BlZCB0byAja3BpR3JpZCBzcGVjaWZpY2FsbHkgKG5vdCB0aGUgc2hhcmVkIC5zdGF0LWdyaWQvLnN0YXQtdGlsZQogICBjbGFzc2VzLCB3aGljaCBDb21wYXJpc29ucyBhbmQgdGhlIFVwbG9hZCBwcmV2aWV3IHN1bW1hcnkgYWxzbyB1c2UpIHNvCiAgIHRoaXMgY29tcGFjdGluZyBkb2Vzbid0IGFmZmVjdCB0aG9zZSBvdGhlciBzdGF0LXRpbGUgZ3JpZHMuIDEwIGdyaWQgdW5pdHMKICAgdG90YWw6IDcgc3RhbmRhcmQgS1BJIHRpbGVzIGF0IDEgdW5pdCBlYWNoICsgQmVzdCBQZXJmb3JtaW5nIFBvc3QgYXQgMwogICB1bml0cyAoYSAzeC13aWRlIGxhbmRzY2FwZSBjYXJkKSwgYWxsIHNoYXJpbmcgb25lIGZpeGVkIHJvdyBoZWlnaHQuIC0tLS0tLS0tLS0gKi8KI2twaUdyaWQgeyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdCgxMCwgbWlubWF4KDAsIDFmcikpOyBnYXA6IDEycHg7IH0KI2twaUdyaWQgLnN0YXQtdGlsZSB7CiAgaGVpZ2h0OiAxMzJweDsgYm94LXNpemluZzogYm9yZGVyLWJveDsgb3ZlcmZsb3c6IGhpZGRlbjsKICBwYWRkaW5nOiAxNnB4IDE4cHg7CiAgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsganVzdGlmeS1jb250ZW50OiBjZW50ZXI7Cn0KI2twaUdyaWQgLnN0YXQtbGFiZWwgeyBmb250LXNpemU6IDEycHg7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGFsaWduLWl0ZW1zOiBmbGV4LXN0YXJ0OyBnYXA6IDhweDsgfQoja3BpR3JpZCAuc3RhdC12YWx1ZSB7IGZvbnQtc2l6ZTogMzJweDsgbWFyZ2luLXRvcDogOHB4OyBsaW5lLWhlaWdodDogMS4xOyB9CiNrcGlHcmlkIC5zdGF0LWRlbHRhIHsgZm9udC1zaXplOiAxM3B4OyBtYXJnaW4tdG9wOiA4cHg7IH0KCi5zdGF0LWljb24gewogIHdpZHRoOiAyOHB4OyBoZWlnaHQ6IDI4cHg7IGZsZXg6IDAgMCBhdXRvOyBib3JkZXItcmFkaXVzOiA1MCU7CiAgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsganVzdGlmeS1jb250ZW50OiBjZW50ZXI7CiAgY29sb3I6ICNmZmY7Cn0KLnN0YXQtaWNvbi52MSB7IGJhY2tncm91bmQ6IHZhcigtLXNlcmllcy0xKTsgfQouc3RhdC1pY29uLnYyIHsgYmFja2dyb3VuZDogdmFyKC0tc2VyaWVzLTIpOyB9Ci5zdGF0LWljb24udjMgeyBiYWNrZ3JvdW5kOiB2YXIoLS1zZXJpZXMtMyk7IH0KLnN0YXQtaWNvbi52NCB7IGJhY2tncm91bmQ6IHZhcigtLXNlcmllcy00KTsgfQouc3RhdC1pY29uLnY1IHsgYmFja2dyb3VuZDogdmFyKC0tc2VyaWVzLTUpOyB9Ci5zdGF0LWljb24udjYgeyBiYWNrZ3JvdW5kOiB2YXIoLS1zZXJpZXMtNik7IH0KLnN0YXQtaWNvbi5nb2xkIHsgYmFja2dyb3VuZDogdmFyKC0tYWNjZW50LWdvbGQpOyB9CkBtZWRpYSAobWF4LXdpZHRoOiA5MDBweCkgeyAja3BpR3JpZCB7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogcmVwZWF0KDUsIG1pbm1heCgwLCAxZnIpKTsgfSB9CkBtZWRpYSAobWF4LXdpZHRoOiA2NDBweCkgeyAja3BpR3JpZCB7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogcmVwZWF0KDIsIG1pbm1heCgwLCAxZnIpKTsgfSB9CgouaW5zaWdodHMtbGlzdCB7IGxpc3Qtc3R5bGU6IG5vbmU7IG1hcmdpbjogMDsgcGFkZGluZzogMDsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiAxMHB4OyB9Ci5pbnNpZ2h0cy1saXN0IGxpIHsKICBmb250LXNpemU6IDEzcHg7IGxpbmUtaGVpZ2h0OiAxLjU7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IHBhZGRpbmctbGVmdDogMThweDsgcG9zaXRpb246IHJlbGF0aXZlOwp9Ci5pbnNpZ2h0cy1saXN0IGxpOjpiZWZvcmUgewogIGNvbnRlbnQ6ICfinKYnOyBwb3NpdGlvbjogYWJzb2x1dGU7IGxlZnQ6IDA7IGNvbG9yOiB2YXIoLS1zZXJpZXMtMSk7IGZvbnQtc2l6ZTogMTFweDsgdG9wOiAycHg7Cn0KCkBrZXlmcmFtZXMgY2FyZEluIHsKICBmcm9tIHsgb3BhY2l0eTogMDsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDEwcHgpOyB9CiAgdG8geyBvcGFjaXR5OiAxOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMCk7IH0KfQoKLyogLS0tLS0tLS0tLSBDYXJkcyAvIGNoYXJ0cyAtLS0tLS0tLS0tICovCi5jYXJkLWdyaWQgeyBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IDJmciAxZnI7IGdhcDogMTZweDsgYWxpZ24taXRlbXM6IHN0YXJ0OyB9Ci5jYXJkLWdyaWQuZXZlbiB7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogMWZyIDFmcjsgfQpAbWVkaWEgKG1heC13aWR0aDogOTAwcHgpIHsgLmNhcmQtZ3JpZCwgLmNhcmQtZ3JpZC5ldmVuIHsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiAxZnI7IH0gfQouY2FyZCB7CiAgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLWxnKTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgcGFkZGluZzogMThweDsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWNhcmQpOwogIHRyYW5zaXRpb246IGJveC1zaGFkb3cgMjIwbXMgdmFyKC0tZWFzZSksIHRyYW5zZm9ybSAyMjBtcyB2YXIoLS1lYXNlKTsKICBhbmltYXRpb246IGNhcmRJbiAzMjBtcyB2YXIoLS1lYXNlKSBiYWNrd2FyZHM7Cn0KLmNhcmQ6aG92ZXIgeyBib3gtc2hhZG93OiB2YXIoLS1zaGFkb3ctaG92ZXIpOyB9Ci5jYXJkLWhlYWRlciB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogc3BhY2UtYmV0d2VlbjsgZ2FwOiA4cHg7IG1hcmdpbi1ib3R0b206IDE0cHg7IH0KLmNhcmQtaGVhZGVyIGgzIHsgZm9udC1zaXplOiAxNHB4OyBtYXJnaW46IDA7IGZvbnQtd2VpZ2h0OiA3MDA7IGxldHRlci1zcGFjaW5nOiAtMC4wMDVlbTsgfQouY2FyZC1oZWFkZXIgc2VsZWN0IHsgZm9udC1zaXplOiAxMnB4OyBwYWRkaW5nOiA2cHggMTBweDsgbWluLXdpZHRoOiAwOyB9Ci5jaGFydC13cmFwIHsgcG9zaXRpb246IHJlbGF0aXZlOyBoZWlnaHQ6IDI4MHB4OyB9Ci5jaGFydC13cmFwLnRhbGwgeyBoZWlnaHQ6IDM0MHB4OyB9CgoubGVnZW5kLXJvdyB7IGRpc3BsYXk6IGZsZXg7IGZsZXgtd3JhcDogd3JhcDsgZ2FwOiAxMnB4OyBtYXJnaW4tdG9wOiAxMHB4OyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IH0KLmxlZ2VuZC1pdGVtIHsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA2cHg7IH0KLmxlZ2VuZC1zd2F0Y2ggeyB3aWR0aDogMTBweDsgaGVpZ2h0OiAxMHB4OyBib3JkZXItcmFkaXVzOiAzcHg7IGRpc3BsYXk6IGlubGluZS1ibG9jazsgfQoubGVnZW5kLWxpbmUgeyB3aWR0aDogMTRweDsgaGVpZ2h0OiAycHg7IGJvcmRlci1yYWRpdXM6IDJweDsgZGlzcGxheTogaW5saW5lLWJsb2NrOyB9CgovKiAtLS0tLS0tLS0tIFRhYmxlcyDigJQgcHJlbWl1bSBkYXRhYmFzZSBmZWVsLCBub3QgYSBzcHJlYWRzaGVldCAtLS0tLS0tLS0tICovCi50YWJsZS1zY3JvbGwgewogIG92ZXJmbG93LXg6IGF1dG87IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1tZCk7CiAgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0yKTsKfQouZGF0YS10YWJsZSB7IHdpZHRoOiAxMDAlOyBib3JkZXItY29sbGFwc2U6IHNlcGFyYXRlOyBib3JkZXItc3BhY2luZzogMDsgZm9udC1zaXplOiAxM3B4OyB9Ci5kYXRhLXRhYmxlIHRoLCAuZGF0YS10YWJsZSB0ZCB7IHRleHQtYWxpZ246IGxlZnQ7IHBhZGRpbmc6IDExcHggMTRweDsgYm9yZGVyLWJvdHRvbTogMXB4IHNvbGlkIHZhcigtLWdyaWRsaW5lKTsgd2hpdGUtc3BhY2U6IG5vd3JhcDsgfQouZGF0YS10YWJsZSB0ZC53cmFwIHsgd2hpdGUtc3BhY2U6IG5vcm1hbDsgfQouZGF0YS10YWJsZSB0aGVhZCB0aCB7CiAgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgZm9udC13ZWlnaHQ6IDYwMDsgZm9udC1zaXplOiAxMXB4OyB0ZXh0LXRyYW5zZm9ybTogdXBwZXJjYXNlOyBsZXR0ZXItc3BhY2luZzogMC4wNGVtOwogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBwb3NpdGlvbjogc3RpY2t5OyB0b3A6IDA7IHotaW5kZXg6IDE7Cn0KLmRhdGEtdGFibGUgdGhlYWQgdGguc29ydGFibGUtdGggeyBjdXJzb3I6IHBvaW50ZXI7IHVzZXItc2VsZWN0OiBub25lOyB0cmFuc2l0aW9uOiBjb2xvciAxNTBtcyB2YXIoLS1lYXNlKTsgfQouZGF0YS10YWJsZSB0aGVhZCB0aC5zb3J0YWJsZS10aDpob3ZlciB7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyB9Ci5kYXRhLXRhYmxlIHRoZWFkIHRoIC5zb3J0LWFycm93IHsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyBmb250LXNpemU6IDEwcHg7IG1hcmdpbi1sZWZ0OiAycHg7IH0KLmRhdGEtdGFibGUgdGhlYWQgdGguc29ydGFibGUtdGg6aG92ZXIgLnNvcnQtYXJyb3cgeyBjb2xvcjogdmFyKC0tc2VyaWVzLTEpOyB9Ci5kYXRhLXRhYmxlIHRib2R5IHRyOm50aC1jaGlsZChldmVuKSB7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS10ZXh0LW11dGVkKSA0JSwgdHJhbnNwYXJlbnQpOyB9Ci5kYXRhLXRhYmxlIHRkLm51bSB7IGZvbnQtdmFyaWFudC1udW1lcmljOiB0YWJ1bGFyLW51bXM7IHRleHQtYWxpZ246IHJpZ2h0OyB9Ci5kYXRhLXRhYmxlIHRoLm51bSB7IHRleHQtYWxpZ246IHJpZ2h0OyB9Ci5kYXRhLXRhYmxlIHRib2R5IHRyIHsgdHJhbnNpdGlvbjogYmFja2dyb3VuZCAxNTBtcyB2YXIoLS1lYXNlKTsgfQouZGF0YS10YWJsZSB0Ym9keSB0cjpob3ZlciB7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zZXJpZXMtMSkgNyUsIHRyYW5zcGFyZW50KTsgfQouZGF0YS10YWJsZSB0Ym9keSB0cjpsYXN0LWNoaWxkIHRkIHsgYm9yZGVyLWJvdHRvbTogbm9uZTsgfQoucGxhdGZvcm0tcGlsbCB7CiAgZGlzcGxheTogaW5saW5lLWZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogNnB4OyBmb250LXNpemU6IDEycHg7IGZvbnQtd2VpZ2h0OiA2MDA7CiAgcGFkZGluZzogNHB4IDEwcHg7IGJvcmRlci1yYWRpdXM6IDIwcHg7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7Cn0KLnBsYXRmb3JtLWRvdCB7IHdpZHRoOiA4cHg7IGhlaWdodDogOHB4OyBib3JkZXItcmFkaXVzOiA1MCU7IH0KCi8qIC0tLS0tLS0tLS0gQnV0dG9ucyDigJQgbmV2ZXIgZmxhdDogc29mdCBzaGFkb3csIGhvdmVyIGxpZnQsIHByZXNzIHNjYWxlIC0tLS0tLS0tLS0gKi8KLmJ0biB7CiAgZGlzcGxheTogaW5saW5lLWZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogY2VudGVyOyBnYXA6IDZweDsKICBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTIpOyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgcGFkZGluZzogOXB4IDE3cHg7IGJvcmRlci1yYWRpdXM6IDExcHg7IGN1cnNvcjogcG9pbnRlcjsgZm9udC1zaXplOiAxM3B4OyBmb250LXdlaWdodDogNjAwOwogIGJveC1zaGFkb3c6IDAgMXB4IDJweCByZ2JhKDE1LDE3LDIxLDAuMDQpOwogIHRyYW5zaXRpb246IHRyYW5zZm9ybSAxNTBtcyB2YXIoLS1lYXNlKSwgYm94LXNoYWRvdyAxNTBtcyB2YXIoLS1lYXNlKSwgZmlsdGVyIDE1MG1zIHZhcigtLWVhc2UpLCBiYWNrZ3JvdW5kIDE1MG1zIHZhcigtLWVhc2UpOwp9Ci5idG4gc3ZnIHsgZmxleC1zaHJpbms6IDA7IH0KLmJ0bjpob3ZlciB7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgtMXB4KTsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWhvdmVyKTsgZmlsdGVyOiBicmlnaHRuZXNzKDEuMDIpOyB9Ci5idG46YWN0aXZlIHsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDApIHNjYWxlKDAuOTYpOyBib3gtc2hhZG93OiAwIDFweCAycHggcmdiYSgxNSwxNywyMSwwLjA2KTsgfQouYnRuLnByaW1hcnkgewogIGJhY2tncm91bmQ6IHZhcigtLXNlcmllcy0xKTsgY29sb3I6ICNmZmY7IGJvcmRlci1jb2xvcjogdHJhbnNwYXJlbnQ7CiAgYm94LXNoYWRvdzogMCA0cHggMTRweCAtNXB4IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zZXJpZXMtMSkgNjUlLCB0cmFuc3BhcmVudCk7Cn0KLmJ0bi5wcmltYXJ5OmhvdmVyIHsgZmlsdGVyOiBicmlnaHRuZXNzKDEuMDcpOyBib3gtc2hhZG93OiAwIDhweCAyMnB4IC02cHggY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSA3MCUsIHRyYW5zcGFyZW50KTsgfQouYnRuLmRhbmdlciB7CiAgYmFja2dyb3VuZDogdmFyKC0tc3RhdHVzLWNyaXRpY2FsKTsgY29sb3I6ICNmZmY7IGJvcmRlci1jb2xvcjogdHJhbnNwYXJlbnQ7CiAgYm94LXNoYWRvdzogMCA0cHggMTRweCAtNXB4IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zdGF0dXMtY3JpdGljYWwpIDU1JSwgdHJhbnNwYXJlbnQpOwp9Ci5idG4uZGFuZ2VyOmhvdmVyIHsgZmlsdGVyOiBicmlnaHRuZXNzKDEuMDYpOyBib3gtc2hhZG93OiAwIDhweCAyMnB4IC02cHggY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXN0YXR1cy1jcml0aWNhbCkgNjAlLCB0cmFuc3BhcmVudCk7IH0KLmJ0bi5zdWNjZXNzIHsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdGF0dXMtZ29vZCk7IGNvbG9yOiAjZmZmOyBib3JkZXItY29sb3I6IHRyYW5zcGFyZW50OwogIGJveC1zaGFkb3c6IDAgNHB4IDE0cHggLTVweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc3RhdHVzLWdvb2QpIDU1JSwgdHJhbnNwYXJlbnQpOwp9Ci5idG4uc3VjY2Vzczpob3ZlciB7IGZpbHRlcjogYnJpZ2h0bmVzcygxLjA2KTsgYm94LXNoYWRvdzogMCA4cHggMjJweCAtNnB4IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zdGF0dXMtZ29vZCkgNjAlLCB0cmFuc3BhcmVudCk7IH0KLmJ0bjpkaXNhYmxlZCB7IG9wYWNpdHk6IDAuNDU7IGN1cnNvcjogbm90LWFsbG93ZWQ7IHRyYW5zZm9ybTogbm9uZTsgYm94LXNoYWRvdzogbm9uZTsgZmlsdGVyOiBub25lOyB9Ci5idG4tcm93IHsgZGlzcGxheTogZmxleDsgZ2FwOiA4cHg7IGZsZXgtd3JhcDogd3JhcDsgfQoKLyogLS0tLS0tLS0tLSBVcGxvYWQgLS0tLS0tLS0tLSAqLwouZHJvcHpvbmUgewogIGJvcmRlcjogMnB4IGRhc2hlZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtbGcpOyBwYWRkaW5nOiA0MHB4IDIwcHg7CiAgdGV4dC1hbGlnbjogY2VudGVyOyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTEpOyBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgY3Vyc29yOiBwb2ludGVyOyB0cmFuc2l0aW9uOiBib3JkZXItY29sb3IgMjAwbXMgdmFyKC0tZWFzZSksIGJhY2tncm91bmQgMjAwbXMgdmFyKC0tZWFzZSksIHRyYW5zZm9ybSAyMDBtcyB2YXIoLS1lYXNlKTsKfQouZHJvcHpvbmU6aG92ZXIgeyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoLTFweCk7IH0KLmRyb3B6b25lLmlzLWRyYWcgeyBib3JkZXItY29sb3I6IHZhcigtLXNlcmllcy0xKTsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSA2JSwgdmFyKC0tc3VyZmFjZS0yKSk7IHRyYW5zZm9ybTogc2NhbGUoMS4wMDUpOyB9Ci5kcm9wem9uZSBoMyB7IG1hcmdpbjogMCAwIDZweDsgZm9udC1zaXplOiAxNXB4OyB9Ci5kcm9wem9uZSBwIHsgbWFyZ2luOiAwOyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyBmb250LXNpemU6IDEzcHg7IH0KLmRyb3B6b25lIGlucHV0W3R5cGU9ImZpbGUiXSB7IGRpc3BsYXk6IG5vbmU7IH0KCi5jb25mbGljdC1saXN0IHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiA4cHg7IG1hcmdpbjogMTJweCAwOyB9Ci5jb25mbGljdC1yb3cgewogIGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogc3BhY2UtYmV0d2VlbjsgZ2FwOiAxMnB4OwogIHBhZGRpbmc6IDExcHggMTRweDsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLXNtKTsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0yKTsKICB0cmFuc2l0aW9uOiBib3gtc2hhZG93IDE4MG1zIHZhcigtLWVhc2UpOwp9Ci5jb25mbGljdC1yb3c6aG92ZXIgeyBib3gtc2hhZG93OiB2YXIoLS1zaGFkb3ctY2FyZCk7IH0KLmNvbmZsaWN0LXJvdyAud2Vlay1sYWJlbCB7IGZvbnQtd2VpZ2h0OiA2MDA7IGZvbnQtc2l6ZTogMTNweDsgfQouY29uZmxpY3Qtcm93IC53ZWVrLW1ldGEgeyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IH0KLmNvbmZsaWN0LXJvdyBzZWxlY3QgeyBtaW4td2lkdGg6IDA7IH0KCi5iYWRnZSB7IGRpc3BsYXk6IGlubGluZS1ibG9jazsgcGFkZGluZzogM3B4IDEwcHg7IGJvcmRlci1yYWRpdXM6IDIwcHg7IGZvbnQtc2l6ZTogMTFweDsgZm9udC13ZWlnaHQ6IDcwMDsgfQouYmFkZ2Uuc3VjY2VzcyB7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zdGF0dXMtZ29vZCkgMTglLCB0cmFuc3BhcmVudCk7IGNvbG9yOiB2YXIoLS1zdGF0dXMtZ29vZCk7IH0KLmJhZGdlLnBhcnRpYWwgeyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc3RhdHVzLXdhcm5pbmcpIDI1JSwgdHJhbnNwYXJlbnQpOyBjb2xvcjogIzhhNjMwMDsgfQouYmFkZ2UuZmFpbGVkIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXN0YXR1cy1jcml0aWNhbCkgMTglLCB0cmFuc3BhcmVudCk7IGNvbG9yOiB2YXIoLS1zdGF0dXMtY3JpdGljYWwpOyB9Ci5iYWRnZS5lcnJvci1zZXYgeyBjb2xvcjogdmFyKC0tc3RhdHVzLWNyaXRpY2FsKTsgfQouYmFkZ2Uud2FybmluZy1zZXYgeyBjb2xvcjogIzhhNjMwMDsgfQouYmFkZ2Uuc2tpcC1zZXYgeyBjb2xvcjogdmFyKC0tdGV4dC1tdXRlZCk7IH0KCi5pc3N1ZXMtbGlzdCB7IG1heC1oZWlnaHQ6IDIyMHB4OyBvdmVyZmxvdy15OiBhdXRvOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtc20pOyB9Ci5pc3N1ZS1yb3cgeyBwYWRkaW5nOiA5cHggMTRweDsgYm9yZGVyLWJvdHRvbTogMXB4IHNvbGlkIHZhcigtLWdyaWRsaW5lKTsgZm9udC1zaXplOiAxMnB4OyB9Ci5pc3N1ZS1yb3c6bGFzdC1jaGlsZCB7IGJvcmRlci1ib3R0b206IG5vbmU7IH0KLmlzc3VlLXJvdyAucm93LW5vIHsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyBtYXJnaW4tcmlnaHQ6IDZweDsgfQoKLyogLS0tLS0tLS0tLSBUb2FzdCAtLS0tLS0tLS0tICovCi50b2FzdC1yb290IHsgcG9zaXRpb246IGZpeGVkOyBib3R0b206IDIwcHg7IHJpZ2h0OiAyMHB4OyBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBnYXA6IDhweDsgei1pbmRleDogMTAwOyB9Ci50b2FzdCB7CiAgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLXNtKTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgcGFkZGluZzogMTJweCAxNnB4OyBib3gtc2hhZG93OiB2YXIoLS1zaGFkb3ctbW9kYWwpOyBmb250LXNpemU6IDEzcHg7IG1heC13aWR0aDogMzQwcHg7CiAgYW5pbWF0aW9uOiB0b2FzdC1pbiAyMjBtcyB2YXIoLS1lYXNlKTsKfQoudG9hc3Quc3VjY2VzcyB7IGJvcmRlci1sZWZ0OiAzcHggc29saWQgdmFyKC0tc3RhdHVzLWdvb2QpOyB9Ci50b2FzdC5lcnJvciB7IGJvcmRlci1sZWZ0OiAzcHggc29saWQgdmFyKC0tc3RhdHVzLWNyaXRpY2FsKTsgfQpAa2V5ZnJhbWVzIHRvYXN0LWluIHsgZnJvbSB7IG9wYWNpdHk6IDA7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgxMHB4KSBzY2FsZSgwLjk4KTsgfSB0byB7IG9wYWNpdHk6IDE7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgwKSBzY2FsZSgxKTsgfSB9CgovKiAtLS0tLS0tLS0tIE1pc2MgLS0tLS0tLS0tLSAqLwoubXV0ZWQgeyBjb2xvcjogdmFyKC0tdGV4dC1tdXRlZCk7IH0KLmVtcHR5LXN0YXRlIHsKICBwYWRkaW5nOiA1NnB4IDI0cHg7IHRleHQtYWxpZ246IGNlbnRlcjsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsKICBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDEycHg7CiAgYW5pbWF0aW9uOiBjYXJkSW4gMjYwbXMgdmFyKC0tZWFzZSk7Cn0KLmVtcHR5LXN0YXRlIC5lbXB0eS1pY29uIHsKICB3aWR0aDogNTJweDsgaGVpZ2h0OiA1MnB4OyBib3JkZXItcmFkaXVzOiAxNnB4OyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6IGNlbnRlcjsKICBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDEwJSwgdHJhbnNwYXJlbnQpOyBjb2xvcjogdmFyKC0tc2VyaWVzLTEpOwp9Ci5lbXB0eS1zdGF0ZSAuZW1wdHktdGl0bGUgeyBmb250LXNpemU6IDE0cHg7IGZvbnQtd2VpZ2h0OiA2MDA7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyB9Ci5lbXB0eS1zdGF0ZSAuZW1wdHktbWVzc2FnZSB7IGZvbnQtc2l6ZTogMTNweDsgbWF4LXdpZHRoOiAzNjBweDsgfQouc3Bpbm5lciB7IHdpZHRoOiAxNnB4OyBoZWlnaHQ6IDE2cHg7IGJvcmRlci1yYWRpdXM6IDUwJTsgYm9yZGVyOiAycHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXRvcC1jb2xvcjogdmFyKC0tc2VyaWVzLTEpOyBhbmltYXRpb246IHNwaW4gLjZzIGxpbmVhciBpbmZpbml0ZTsgZGlzcGxheTogaW5saW5lLWJsb2NrOyB9CkBrZXlmcmFtZXMgc3BpbiB7IHRvIHsgdHJhbnNmb3JtOiByb3RhdGUoMzYwZGVnKTsgfSB9Ci5sb2FkaW5nLXJvdyB7IHBhZGRpbmc6IDQwcHggMjBweDsgdGV4dC1hbGlnbjogY2VudGVyOyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyB9CgovKiBTa2VsZXRvbiBsb2FkZXJzIOKAlCBzaGltbWVyaW5nIHBsYWNlaG9sZGVycyBzaG93biB3aGlsZSBhIHNlY3Rpb24ncyBkYXRhIGlzIGluIGZsaWdodCAqLwouc2tlbGV0b24gewogIGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1zbSk7CiAgYmFja2dyb3VuZDogbGluZWFyLWdyYWRpZW50KDEwMGRlZywgY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXRleHQtbXV0ZWQpIDEyJSwgdHJhbnNwYXJlbnQpIDMwJSwgY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXRleHQtbXV0ZWQpIDIyJSwgdHJhbnNwYXJlbnQpIDUwJSwgY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXRleHQtbXV0ZWQpIDEyJSwgdHJhbnNwYXJlbnQpIDcwJSk7CiAgYmFja2dyb3VuZC1zaXplOiAyMDAlIDEwMCU7CiAgYW5pbWF0aW9uOiBza2VsZXRvblNoaW1tZXIgMS40cyBlYXNlLWluLW91dCBpbmZpbml0ZTsKfQpAa2V5ZnJhbWVzIHNrZWxldG9uU2hpbW1lciB7IGZyb20geyBiYWNrZ3JvdW5kLXBvc2l0aW9uOiAxNTAlIDA7IH0gdG8geyBiYWNrZ3JvdW5kLXBvc2l0aW9uOiAtNTAlIDA7IH0gfQouc2tlbGV0b24tc3RhdC1ncmlkIHsgZGlzcGxheTogZ3JpZDsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiByZXBlYXQoYXV0by1maXQsIG1pbm1heCgxODBweCwgMWZyKSk7IGdhcDogMTRweDsgfQouc2tlbGV0b24tdGlsZSB7IGhlaWdodDogODRweDsgfQouc2tlbGV0b24tY2hhcnQgeyBoZWlnaHQ6IDI4MHB4OyB3aWR0aDogMTAwJTsgfQouc2tlbGV0b24tcm93IHsgaGVpZ2h0OiA0MHB4OyBtYXJnaW4tYm90dG9tOiA4cHg7IH0KCi8qIEFuaW1hdGVkIGhvcml6b250YWwgY29tcGFyaXNvbiBiYXIg4oCUIGEgbGFiZWxlZCByb3cgd2l0aCBhIHRyYWNrIHRoYXQgZmlsbHMgaW4gb24gaW5zZXJ0aW9uICovCi5iYXItcm93IHsgZGlzcGxheTogZ3JpZDsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiBtaW5tYXgoOTBweCwgMTQwcHgpIDFmciBhdXRvOyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDEwcHg7IHBhZGRpbmc6IDVweCAwOyB9Ci5iYXItbGFiZWwgeyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IGZvbnQtd2VpZ2h0OiA2MDA7IH0KLmJhci10cmFjayB7IGhlaWdodDogOHB4OyBib3JkZXItcmFkaXVzOiA1cHg7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS10ZXh0LW11dGVkKSAxNCUsIHRyYW5zcGFyZW50KTsgb3ZlcmZsb3c6IGhpZGRlbjsgfQouYmFyLWZpbGwgeyBoZWlnaHQ6IDEwMCU7IHdpZHRoOiAwJTsgYm9yZGVyLXJhZGl1czogNXB4OyB0cmFuc2l0aW9uOiB3aWR0aCA3MDBtcyBjdWJpYy1iZXppZXIoMC4xNiwgMSwgMC4zLCAxKTsgfQouYmFyLXZhbHVlIHsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNzAwOyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgZm9udC12YXJpYW50LW51bWVyaWM6IHRhYnVsYXItbnVtczsgdGV4dC1hbGlnbjogcmlnaHQ7IG1pbi13aWR0aDogNTZweDsgfQoKQG1lZGlhIChwcmVmZXJzLXJlZHVjZWQtbW90aW9uOiByZWR1Y2UpIHsKICAuYmFyLWZpbGwgeyB0cmFuc2l0aW9uLWR1cmF0aW9uOiAxbXM7IH0KICAuc2tlbGV0b24geyBhbmltYXRpb24tZHVyYXRpb246IDFtczsgfQogIC5jYXJkLCAuc3RhdC10aWxlIHsgYW5pbWF0aW9uLWR1cmF0aW9uOiAxbXM7IH0KfQoKLnR3by1jb2wgeyBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IDFmciAxZnI7IGdhcDogMTZweDsgfQpAbWVkaWEgKG1heC13aWR0aDogOTAwcHgpIHsgLnR3by1jb2wgeyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IDFmcjsgfSB9CgoubW9kZS10YWJzIHsgZGlzcGxheTogZmxleDsgZ2FwOiA2cHg7IGZsZXgtd3JhcDogd3JhcDsgbWFyZ2luLWJvdHRvbTogMTZweDsgfQoubW9kZS10YWJzIGJ1dHRvbiB7CiAgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgcGFkZGluZzogN3B4IDE0cHg7IGJvcmRlci1yYWRpdXM6IDIwcHg7IGZvbnQtc2l6ZTogMTJweDsgZm9udC13ZWlnaHQ6IDYwMDsgY3Vyc29yOiBwb2ludGVyOwogIHRyYW5zaXRpb246IGNvbG9yIDE4MG1zIHZhcigtLWVhc2UpLCBiYWNrZ3JvdW5kIDE4MG1zIHZhcigtLWVhc2UpLCB0cmFuc2Zvcm0gMTUwbXMgdmFyKC0tZWFzZSk7Cn0KLm1vZGUtdGFicyBidXR0b246aG92ZXIgeyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoLTFweCk7IH0KLm1vZGUtdGFicyBidXR0b24uaXMtYWN0aXZlIHsgYmFja2dyb3VuZDogdmFyKC0tc2VyaWVzLTEpOyBjb2xvcjogI2ZmZjsgYm9yZGVyLWNvbG9yOiB0cmFuc3BhcmVudDsgYm94LXNoYWRvdzogMCA0cHggMTRweCAtNXB4IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zZXJpZXMtMSkgNjAlLCB0cmFuc3BhcmVudCk7IH0KCi5maWVsZC1pbmxpbmUgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDhweDsgZm9udC1zaXplOiAxMnB4OyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyB9Ci5maWVsZC1pbmxpbmUgc2VsZWN0LCAuZmllbGQtaW5saW5lIGlucHV0IHsgbWluLXdpZHRoOiAwOyBwYWRkaW5nOiA2cHggMTBweDsgfQoKLyogLS0tLS0tLS0tLSBQbGF0Zm9ybSBQZXJmb3JtYW5jZSBDb21wYXJpc29uIGNhcmRzIC0tLS0tLS0tLS0gKi8KLnBjYy1zZWN0aW9uIHsgbWFyZ2luLXRvcDogMjRweDsgfQoucGNjLWNvbnRyb2xzIHsgZGlzcGxheTogZmxleDsgZmxleC13cmFwOiB3cmFwOyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDE2cHg7IG1hcmdpbi1ib3R0b206IDE2cHg7IH0KLnBsYXRmb3JtLWNvbXBhcmUtZ3JpZCB7CiAgZGlzcGxheTogZ3JpZDsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiByZXBlYXQoMiwgMWZyKTsgZ2FwOiAxNnB4Owp9CkBtZWRpYSAobWF4LXdpZHRoOiA5MDBweCkgeyAucGxhdGZvcm0tY29tcGFyZS1ncmlkIHsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiAxZnI7IH0gfQoucGxhdGZvcm0tY29tcGFyZS1jYXJkIHsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTEpOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtbGcpOwogIGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBwYWRkaW5nOiAxOHB4OyBib3gtc2hhZG93OiB2YXIoLS1zaGFkb3ctY2FyZCk7CiAgdHJhbnNpdGlvbjogYm94LXNoYWRvdyAyMjBtcyB2YXIoLS1lYXNlKSwgdHJhbnNmb3JtIDIyMG1zIHZhcigtLWVhc2UpOwogIGFuaW1hdGlvbjogY2FyZEluIDMyMG1zIHZhcigtLWVhc2UpIGJhY2t3YXJkczsKfQoucGxhdGZvcm0tY29tcGFyZS1jYXJkOmhvdmVyIHsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWhvdmVyKTsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKC0ycHgpOyB9Ci5wY2MtaGVhZGVyIHsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsganVzdGlmeS1jb250ZW50OiBzcGFjZS1iZXR3ZWVuOyBnYXA6IDEwcHg7IH0KLnBjYy1oZWFkZXItbmFtZSB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogOHB4OyB9Ci5wY2MtbmFtZSB7IGZvbnQtc2l6ZTogMTVweDsgZm9udC13ZWlnaHQ6IDcwMDsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IH0KLnBjYy1iYWRnZSB7IGZvbnQtc2l6ZTogMTNweDsgZm9udC13ZWlnaHQ6IDcwMDsgcGFkZGluZzogNHB4IDEwcHg7IGJvcmRlci1yYWRpdXM6IDIwcHg7IH0KLnBjYy1iYWRnZS51cCB7IGNvbG9yOiB2YXIoLS1zdWNjZXNzLXRleHQpOyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc3RhdHVzLWdvb2QpIDE0JSwgdHJhbnNwYXJlbnQpOyB9Ci5wY2MtYmFkZ2UuZG93biB7IGNvbG9yOiB2YXIoLS1zdGF0dXMtY3JpdGljYWwpOyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc3RhdHVzLWNyaXRpY2FsKSAxMiUsIHRyYW5zcGFyZW50KTsgfQoucGNjLWJhZGdlLmZsYXQgeyBjb2xvcjogdmFyKC0tdGV4dC1tdXRlZCk7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS10ZXh0LW11dGVkKSAxMiUsIHRyYW5zcGFyZW50KTsgfQoucGNjLWNhcHRpb24geyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IG1hcmdpbi10b3A6IDZweDsgfQoucGNjLW1ldHJpY3MgeyBtYXJnaW4tdG9wOiAxNnB4OyBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBnYXA6IDE0cHg7IH0KLnBjYy1tZXRyaWMtcm93IHsgcGFkZGluZy10b3A6IDEycHg7IGJvcmRlci10b3A6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyB9Ci5wY2MtbWV0cmljLXJvdzpmaXJzdC1jaGlsZCB7IHBhZGRpbmctdG9wOiAwOyBib3JkZXItdG9wOiBub25lOyB9Ci5wY2MtbWV0cmljLWhlYWRlciB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogc3BhY2UtYmV0d2VlbjsgZ2FwOiA4cHg7IG1hcmdpbi1ib3R0b206IDZweDsgfQoucGNjLW1ldHJpYy1sYWJlbCB7IGZvbnQtc2l6ZTogMTJweDsgZm9udC13ZWlnaHQ6IDcwMDsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IH0KLnBjYy1tZXRyaWMtZGlmZiB7IGZvbnQtc2l6ZTogMTJweDsgZm9udC13ZWlnaHQ6IDcwMDsgfQoucGNjLW1ldHJpYy1kaWZmLnVwIHsgY29sb3I6IHZhcigtLXN1Y2Nlc3MtdGV4dCk7IH0KLnBjYy1tZXRyaWMtZGlmZi5kb3duIHsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IH0KLnBjYy1tZXRyaWMtZGlmZi5mbGF0IHsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyB9Ci5wY2MtZm9vdGVyIHsgbWFyZ2luLXRvcDogMTZweDsgcGFkZGluZy10b3A6IDE0cHg7IGJvcmRlci10b3A6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyB9Ci5wY2MtZm9vdGVyLWxhYmVsIHsgZm9udC1zaXplOiAxMXB4OyBjb2xvcjogdmFyKC0tdGV4dC1tdXRlZCk7IGZvbnQtd2VpZ2h0OiA2MDA7IHRleHQtdHJhbnNmb3JtOiB1cHBlcmNhc2U7IGxldHRlci1zcGFjaW5nOiAwLjAzZW07IH0KLnBjYy1mb290ZXItdmFsdWUgeyBmb250LXNpemU6IDE1cHg7IGZvbnQtd2VpZ2h0OiA3MDA7IG1hcmdpbi10b3A6IDRweDsgfQoucGNjLWZvb3Rlci12YWx1ZS51cCB7IGNvbG9yOiB2YXIoLS1zdWNjZXNzLXRleHQpOyB9Ci5wY2MtZm9vdGVyLXZhbHVlLmRvd24geyBjb2xvcjogdmFyKC0tc3RhdHVzLWNyaXRpY2FsKTsgfQoucGNjLWZvb3Rlci12YWx1ZS5mbGF0IHsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyB9Ci5wY2MtZm9vdGVyLWRldGFpbCB7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgbWFyZ2luLXRvcDogNHB4OyB9Ci5wY2Mtdmlldy1saW5rIHsKICBkaXNwbGF5OiBpbmxpbmUtZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA0cHg7IG1hcmdpbi10b3A6IDE0cHg7CiAgYmFja2dyb3VuZDogbm9uZTsgYm9yZGVyOiBub25lOyBjb2xvcjogdmFyKC0tc2VyaWVzLTEpOyBmb250LXNpemU6IDEycHg7IGZvbnQtd2VpZ2h0OiA3MDA7IGN1cnNvcjogcG9pbnRlcjsgcGFkZGluZzogMDsKICB0cmFuc2l0aW9uOiBvcGFjaXR5IDE1MG1zIHZhcigtLWVhc2UpOwp9Ci5wY2Mtdmlldy1saW5rOmhvdmVyIHsgb3BhY2l0eTogMC43NTsgdGV4dC1kZWNvcmF0aW9uOiB1bmRlcmxpbmU7IH0KCi8qIC0tLS0tLS0tLS0gUGFnaW5hdGlvbiAtLS0tLS0tLS0tICovCi5wYWdpbmF0aW9uLXJvdyB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogMTJweDsgbWFyZ2luLXRvcDogMTRweDsgZm9udC1zaXplOiAxMnB4OyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyB9Ci5wYWdpbmF0aW9uLXJvdyAuYnRuIHsgcGFkZGluZzogNnB4IDEycHg7IH0KLmV4cG9ydC1idXR0b25zIHsgZGlzcGxheTogZmxleDsgZ2FwOiA4cHg7IGZsZXgtd3JhcDogd3JhcDsgbWFyZ2luLWJvdHRvbTogMTJweDsgfQouZXhwb3J0LWJ1dHRvbnMgLmJ0biB7IHBhZGRpbmc6IDdweCAxM3B4OyBmb250LXNpemU6IDEycHg7IH0KCi8qIC0tLS0tLS0tLS0gRGFzaGJvYXJkIGNvbnRyb2xzIC8gbWV0cmljLWZvY3VzZWQgS1BJcyAtLS0tLS0tLS0tICovCi5kYXNoYm9hcmQtY29udHJvbHMgewogIGRpc3BsYXk6IGZsZXg7IGZsZXgtd3JhcDogd3JhcDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiAxMHB4OyBtYXJnaW4tYm90dG9tOiAxOHB4Owp9Ci5kYXNoYm9hcmQtY29udHJvbHMgbGFiZWwgeyBmb250LXNpemU6IDEycHg7IGZvbnQtd2VpZ2h0OiA2MDA7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IG1hcmdpbi1yaWdodDogNnB4OyB9Ci5kYXNoYm9hcmQtY29udHJvbHMgc2VsZWN0IHsgZm9udC13ZWlnaHQ6IDYwMDsgfQovKiBCZXN0IFBlcmZvcm1pbmcgUG9zdCDigJQgYSBmZWF0dXJlZCBsYW5kc2NhcGUgY2FyZCBzcGFubmluZyAzIEtQSS10aWxlLXdpZHRocwogICAoYSBzdGFuZGFyZCB0aWxlIGlzIDEgdW5pdDsgI2twaUdyaWQgaGFzIDEwIHVuaXRzIHRvdGFsKSwgc2FtZSBmaXhlZAogICBoZWlnaHQgYXMgdGhlIHJlc3Qgb2YgI2twaUdyaWQ6IGNhcHRpb24vcGxhdGZvcm0vZGF0ZSBzaXQgb24gdGhlIGxlZnQsCiAgIHdpdGggdGhlIHNlbGVjdGVkIG1ldHJpYyAobGFyZ2UpIGFuZCBDdXJyZW50IEZvbGxvd2VycyAoc21hbGxlciwgYmVsb3cgYQogICBkaXZpZGVyKSBzdGFja2VkIGluIGEgbmFycm93ZXIgY29sdW1uIG9uIHRoZSByaWdodC4gKi8KI2twaUdyaWQgLnBvc3QtdGlsZSB7CiAgZ3JpZC1jb2x1bW46IHNwYW4gMzsKICBmbGV4LWRpcmVjdGlvbjogcm93OwogIGFsaWduLWl0ZW1zOiBzdHJldGNoOwogIGp1c3RpZnktY29udGVudDogZmxleC1zdGFydDsKICBnYXA6IDIwcHg7CiAgcGFkZGluZzogMTRweCAyMHB4Owp9CiNrcGlHcmlkIC5wb3N0LXRpbGUtbWFpbiB7IGZsZXg6IDEgMSBhdXRvOyBtaW4td2lkdGg6IDA7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGp1c3RpZnktY29udGVudDogY2VudGVyOyBnYXA6IDZweDsgfQoja3BpR3JpZCAucG9zdC10aWxlLWNhcHRpb24gewogIGZvbnQtc2l6ZTogMTNweDsgZm9udC13ZWlnaHQ6IDYwMDsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7CiAgbGluZS1oZWlnaHQ6IDEuNDsKICBkaXNwbGF5OiAtd2Via2l0LWJveDsgLXdlYmtpdC1saW5lLWNsYW1wOiAzOyAtd2Via2l0LWJveC1vcmllbnQ6IHZlcnRpY2FsOyBvdmVyZmxvdzogaGlkZGVuOwp9CiNrcGlHcmlkIC5wb3N0LXRpbGUtY2FwdGlvbi5tdXRlZCB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgZm9udC13ZWlnaHQ6IDUwMDsgLXdlYmtpdC1saW5lLWNsYW1wOiAxOyB9CiNrcGlHcmlkIC5wb3N0LXRpbGUtbWV0YSB7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDZweDsgd2hpdGUtc3BhY2U6IG5vd3JhcDsgb3ZlcmZsb3c6IGhpZGRlbjsgdGV4dC1vdmVyZmxvdzogZWxsaXBzaXM7IH0KI2twaUdyaWQgLnBvc3QtdGlsZS1kaXZpZGVyIHsgZmxleDogMCAwIGF1dG87IHdpZHRoOiAxcHg7IGFsaWduLXNlbGY6IHN0cmV0Y2g7IGJhY2tncm91bmQ6IHZhcigtLWJvcmRlcik7IH0KI2twaUdyaWQgLnBvc3QtdGlsZS1tZXRyaWNzIHsKICBmbGV4OiAwIDAgYXV0bzsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsganVzdGlmeS1jb250ZW50OiBjZW50ZXI7IGdhcDogOHB4OyBtaW4td2lkdGg6IDEyMHB4Owp9CiNrcGlHcmlkIC5wb3N0LXRpbGUtbWV0cmljLWJsb2NrIHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiAzcHg7IH0KI2twaUdyaWQgLnBvc3QtdGlsZS1tZXRyaWMtYmxvY2suc2Vjb25kYXJ5IHsgcGFkZGluZy10b3A6IDhweDsgYm9yZGVyLXRvcDogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IH0KI2twaUdyaWQgLnBvc3QtdGlsZS1tZXRyaWMtbGFiZWwgeyBmb250LXNpemU6IDExcHg7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgZm9udC13ZWlnaHQ6IDYwMDsgdGV4dC10cmFuc2Zvcm06IHVwcGVyY2FzZTsgbGV0dGVyLXNwYWNpbmc6IDAuMDNlbTsgfQoja3BpR3JpZCAucG9zdC10aWxlLW1ldHJpYy12YWx1ZSB7IGZvbnQtc2l6ZTogMjJweDsgZm9udC13ZWlnaHQ6IDcwMDsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IGZvbnQtdmFyaWFudC1udW1lcmljOiB0YWJ1bGFyLW51bXM7IH0KI2twaUdyaWQgLnBvc3QtdGlsZS1tZXRyaWMtYmxvY2suc2Vjb25kYXJ5IC5wb3N0LXRpbGUtbWV0cmljLXZhbHVlIHsgZm9udC1zaXplOiAxNXB4OyB9CkBtZWRpYSAobWF4LXdpZHRoOiA2NDBweCkgeyAja3BpR3JpZCAucG9zdC10aWxlIHsgZ3JpZC1jb2x1bW46IHNwYW4gMjsgfSB9Cgouc3RhdC12YWx1ZS1tdXRlZCB7IGZvbnQtc2l6ZTogMTVweCAhaW1wb3J0YW50OyBjb2xvcjogdmFyKC0tdGV4dC1tdXRlZCk7IGZvbnQtd2VpZ2h0OiA2MDA7IH0KLmNhcHRpb24tbGluayB7IGNvbG9yOiB2YXIoLS1zZXJpZXMtMSk7IHRleHQtZGVjb3JhdGlvbjogbm9uZTsgfQouY2FwdGlvbi1saW5rOmhvdmVyIHsgdGV4dC1kZWNvcmF0aW9uOiB1bmRlcmxpbmU7IH0KCi8qIC0tLS0tLS0tLS0gRGF0YSBSZWNvcmRzIChwbGF0Zm9ybS1ncm91cGVkKSAtLS0tLS0tLS0tICovCi5yZWNvcmRzLXRvb2xiYXIgewogIGRpc3BsYXk6IGZsZXg7IGZsZXgtd3JhcDogd3JhcDsgYWxpZ24taXRlbXM6IGNlbnRlcjsganVzdGlmeS1jb250ZW50OiBzcGFjZS1iZXR3ZWVuOwogIGdhcDogMTJweDsgbWFyZ2luLWJvdHRvbTogMTRweDsKfQoucGxhdGZvcm0tZmlsdGVyLXBpbGxzIHsgZGlzcGxheTogZmxleDsgZmxleC13cmFwOiB3cmFwOyBnYXA6IDZweDsgfQoucGxhdGZvcm0tZmlsdGVyLXBpbGxzIGJ1dHRvbiB7CiAgZGlzcGxheTogaW5saW5lLWZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogNnB4OwogIGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHBhZGRpbmc6IDdweCAxNHB4OyBib3JkZXItcmFkaXVzOiAyMHB4OyBmb250LXNpemU6IDEycHg7IGZvbnQtd2VpZ2h0OiA2MDA7IGN1cnNvcjogcG9pbnRlcjsKICB0cmFuc2l0aW9uOiBjb2xvciAxODBtcyB2YXIoLS1lYXNlKSwgYmFja2dyb3VuZCAxODBtcyB2YXIoLS1lYXNlKSwgdHJhbnNmb3JtIDE1MG1zIHZhcigtLWVhc2UpLCBib3gtc2hhZG93IDE4MG1zIHZhcigtLWVhc2UpOwp9Ci5wbGF0Zm9ybS1maWx0ZXItcGlsbHMgYnV0dG9uOmhvdmVyIHsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgtMXB4KTsgfQoucGxhdGZvcm0tZmlsdGVyLXBpbGxzIGJ1dHRvbjphY3RpdmUgeyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMCkgc2NhbGUoMC45Nik7IH0KLnBsYXRmb3JtLWZpbHRlci1waWxscyBidXR0b24uaXMtYWN0aXZlIHsgYmFja2dyb3VuZDogdmFyKC0tc2VyaWVzLTEpOyBjb2xvcjogI2ZmZjsgYm9yZGVyLWNvbG9yOiB0cmFuc3BhcmVudDsgYm94LXNoYWRvdzogMCA0cHggMTRweCAtNXB4IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zZXJpZXMtMSkgNjAlLCB0cmFuc3BhcmVudCk7IH0KLnBsYXRmb3JtLWZpbHRlci1waWxscyBidXR0b24uaXMtYWN0aXZlIC5wbGF0Zm9ybS1kb3QgeyBib3gtc2hhZG93OiAwIDAgMCAycHggcmdiYSgyNTUsMjU1LDI1NSwwLjUpOyB9Ci5yZWNvcmRzLXNlYXJjaCBpbnB1dCB7IGJvcmRlci1yYWRpdXM6IDIwcHg7IG1pbi13aWR0aDogMjIwcHg7IH0KLnN0YXR1cy1waWxsIHsgZGlzcGxheTogaW5saW5lLWJsb2NrOyBwYWRkaW5nOiAzcHggMTBweDsgYm9yZGVyLXJhZGl1czogMjBweDsgZm9udC1zaXplOiAxMXB4OyBmb250LXdlaWdodDogNzAwOyB9Ci5zdGF0dXMtcGlsbC5vcmlnaW5hbCB7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS10ZXh0LW11dGVkKSAxNSUsIHRyYW5zcGFyZW50KTsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgfQouc3RhdHVzLXBpbGwuZWRpdGVkIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXN0YXR1cy13YXJuaW5nKSAyMiUsIHRyYW5zcGFyZW50KTsgY29sb3I6ICM4YTYzMDA7IH0KLnJvdy1hY3Rpb25zIHsgZGlzcGxheTogZmxleDsgZ2FwOiA2cHg7IGZsZXgtd3JhcDogbm93cmFwOyB9Ci5yb3ctYWN0aW9ucyAuYnRuIHsgcGFkZGluZzogNXB4IDEwcHg7IGZvbnQtc2l6ZTogMTJweDsgfQoubGluay1jZWxsIGEgeyBjb2xvcjogdmFyKC0tc2VyaWVzLTEpOyB0ZXh0LWRlY29yYXRpb246IG5vbmU7IGZvbnQtd2VpZ2h0OiA2MDA7IGZvbnQtc2l6ZTogMTJweDsgfQoubGluay1jZWxsIGE6aG92ZXIgeyB0ZXh0LWRlY29yYXRpb246IHVuZGVybGluZTsgfQoucmVjb3JkLXNlY3Rpb24gewogIGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1zbSk7IHBhZGRpbmc6IDE2cHg7IG1hcmdpbi1ib3R0b206IDE0cHg7CiAgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0yKTsgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwp9Ci5yZWNvcmQtc2VjdGlvbiBoNCB7IG1hcmdpbjogMCAwIDEycHg7IGZvbnQtc2l6ZTogMTJweDsgZm9udC13ZWlnaHQ6IDcwMDsgbGV0dGVyLXNwYWNpbmc6IDAuMDNlbTsgdGV4dC10cmFuc2Zvcm06IHVwcGVyY2FzZTsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgfQoucmVjb3JkLXNlY3Rpb24gLmZvcm0tZ3JpZCB7IG1hcmdpbi1ib3R0b206IDA7IH0KLnJlY29yZC1zZWN0aW9uIC52aWV3LWZpZWxkIHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiAycHg7IGZvbnQtc2l6ZTogMTNweDsgfQoucmVjb3JkLXNlY3Rpb24gLnZpZXctZmllbGQgLnZpZXctbGFiZWwgeyBmb250LXNpemU6IDExcHg7IGZvbnQtd2VpZ2h0OiA2MDA7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgfQoucmVjb3JkLXNlY3Rpb24gLnZpZXctZmllbGQgLnZpZXctdmFsdWUgeyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgd29yZC1icmVhazogYnJlYWstd29yZDsgfQpAbWVkaWEgKG1heC13aWR0aDogNjQwcHgpIHsKICAucmVjb3Jkcy10b29sYmFyIHsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgYWxpZ24taXRlbXM6IHN0cmV0Y2g7IH0KICAucmVjb3Jkcy1zZWFyY2ggaW5wdXQgeyB3aWR0aDogMTAwJTsgfQp9CgovKiAtLS0tLS0tLS0tIE1vZGFsIChyZWNvcmQgZWRpdG9yKSAtLS0tLS0tLS0tICovCi5tb2RhbC1vdmVybGF5IHsKICBwb3NpdGlvbjogZml4ZWQ7IGluc2V0OiAwOyBiYWNrZ3JvdW5kOiByZ2JhKDEwLDExLDEzLDAuNSk7CiAgYmFja2Ryb3AtZmlsdGVyOiBibHVyKDZweCk7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiBibHVyKDZweCk7CiAgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGZsZXgtc3RhcnQ7IGp1c3RpZnktY29udGVudDogY2VudGVyOwogIHBhZGRpbmc6IDQwcHggMTZweDsgb3ZlcmZsb3cteTogYXV0bzsgei1pbmRleDogMjAwOwogIGFuaW1hdGlvbjogb3ZlcmxheUluIDIwMG1zIHZhcigtLWVhc2UpOwp9CkBrZXlmcmFtZXMgb3ZlcmxheUluIHsgZnJvbSB7IG9wYWNpdHk6IDA7IH0gdG8geyBvcGFjaXR5OiAxOyB9IH0KQGtleWZyYW1lcyBtb2RhbFBhbmVsSW4gewogIGZyb20geyBvcGFjaXR5OiAwOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMTRweCkgc2NhbGUoMC45Nyk7IH0KICB0byB7IG9wYWNpdHk6IDE7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgwKSBzY2FsZSgxKTsgfQp9Ci5tb2RhbC1wYW5lbCB7CiAgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLWxnKTsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgcGFkZGluZzogMjRweDsgd2lkdGg6IDEwMCU7IG1heC13aWR0aDogNzIwcHg7IGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdy1tb2RhbCk7CiAgbWF4LWhlaWdodDogY2FsYygxMDB2aCAtIDgwcHgpOyBvdmVyZmxvdy15OiBhdXRvOwogIGFuaW1hdGlvbjogbW9kYWxQYW5lbEluIDI0MG1zIHZhcigtLWVhc2UpOwp9Ci5tb2RhbC1wYW5lbC53aWRlIHsgbWF4LXdpZHRoOiAxMTAwcHg7IH0KLm1vZGFsLXBhbmVsIGgyIHsgbWFyZ2luOiAwIDAgNHB4OyBmb250LXNpemU6IDE3cHg7IGxldHRlci1zcGFjaW5nOiAtMC4wMWVtOyB9Ci5tb2RhbC1wYW5lbCAubW9kYWwtc3ViIHsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgZm9udC1zaXplOiAxMnB4OyBtYXJnaW46IDAgMCAxOHB4OyB9Ci5mb3JtLWdyaWQgeyBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdChhdXRvLWZpdCwgbWlubWF4KDIwMHB4LCAxZnIpKTsgZ2FwOiAxMnB4OyBtYXJnaW4tYm90dG9tOiAxNnB4OyB9Ci5mb3JtLWdyaWQuZnVsbCB7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogMWZyOyB9CkBtZWRpYSAobWF4LXdpZHRoOiA2NDBweCkgeyAuZm9ybS1ncmlkIHsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiAxZnI7IH0gfQouZm9ybS1maWVsZCB7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGdhcDogNXB4OyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IH0KLmZvcm0tZmllbGQgbGFiZWwgeyBmb250LXdlaWdodDogNjAwOyB9Ci5mb3JtLWZpZWxkIHRleHRhcmVhIHsgcmVzaXplOiB2ZXJ0aWNhbDsgbWluLWhlaWdodDogNjBweDsgfQoKLnBsYXRmb3JtLWVkaXQtcm93IHsKICBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtc20pOyBwYWRkaW5nOiAxNHB4OyBtYXJnaW4tYm90dG9tOiAxMHB4OyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTIpOwp9Ci5wbGF0Zm9ybS1lZGl0LXJvdyAucGxhdGZvcm0tZWRpdC1oZWFkIHsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsganVzdGlmeS1jb250ZW50OiBzcGFjZS1iZXR3ZWVuOyBnYXA6IDhweDsgbWFyZ2luLWJvdHRvbTogMTBweDsgfQoucGxhdGZvcm0tZWRpdC1yb3cgLm1ldHJpY3MtZ3JpZCB7IGRpc3BsYXk6IGdyaWQ7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogcmVwZWF0KGF1dG8tZml0LCBtaW5tYXgoMTIwcHgsIDFmcikpOyBnYXA6IDhweDsgfQoucmVtb3ZlLXBsYXRmb3JtLWJ0biB7IGJvcmRlcjogbm9uZTsgYmFja2dyb3VuZDogdHJhbnNwYXJlbnQ7IGNvbG9yOiB2YXIoLS1zdGF0dXMtY3JpdGljYWwpOyBjdXJzb3I6IHBvaW50ZXI7IGZvbnQtc2l6ZTogMTJweDsgZm9udC13ZWlnaHQ6IDYwMDsgdHJhbnNpdGlvbjogb3BhY2l0eSAxNTBtcyB2YXIoLS1lYXNlKTsgfQoucmVtb3ZlLXBsYXRmb3JtLWJ0bjpob3ZlciB7IG9wYWNpdHk6IDAuNzsgfQoubW9kYWwtYWN0aW9ucyB7IGRpc3BsYXk6IGZsZXg7IGp1c3RpZnktY29udGVudDogc3BhY2UtYmV0d2VlbjsgYWxpZ24taXRlbXM6IGNlbnRlcjsgbWFyZ2luLXRvcDogMThweDsgZ2FwOiA4cHg7IGZsZXgtd3JhcDogd3JhcDsgfQoKLyogLS0tLS0tLS0tLSBSZXNwb25zaXZlIHRpZ2h0ZW5pbmcgLS0tLS0tLS0tLSAqLwpAbWVkaWEgKG1heC13aWR0aDogNzIwcHgpIHsKICAuYXBwLXNoZWxsIHsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgfQogIC5zaWRlYmFyIHsgd2lkdGg6IDEwMCU7IGhlaWdodDogYXV0bzsgcG9zaXRpb246IHN0YXRpYzsgZmxleC1kaXJlY3Rpb246IHJvdzsgZmxleC13cmFwOiB3cmFwOyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDhweDsgcGFkZGluZzogMTBweCAxNHB4OyB9CiAgLnNpZGViYXItYnJhbmQgeyBwYWRkaW5nOiAwOyBtYXJnaW4tcmlnaHQ6IGF1dG87IH0KICAudGFicyB7IGZsZXgtZGlyZWN0aW9uOiByb3c7IHdpZHRoOiAxMDAlOyBvdmVyZmxvdy14OiBhdXRvOyBvcmRlcjogMzsgfQogIC5zaWRlYmFyLWZvb3RlciB7IGZsZXgtZGlyZWN0aW9uOiByb3c7IGJvcmRlci10b3A6IG5vbmU7IG1hcmdpbi10b3A6IDA7IHBhZGRpbmctdG9wOiAwOyB9CiAgLnZpZXctYXJlYSB7IHBhZGRpbmc6IDE0cHg7IH0KICAuZmlsdGVyLWJhciB7IHRvcDogYXV0bzsgcG9zaXRpb246IHN0YXRpYzsgcGFkZGluZzogMTJweCAxNHB4OyB9CiAgLnN0YXQtZ3JpZCB7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogcmVwZWF0KGF1dG8tZml0LCBtaW5tYXgoMTQwcHgsIDFmcikpOyB9CiAgLmJyYW5kLWxvZ28geyBoZWlnaHQ6IDIycHg7IH0KfQo8L3N0eWxlPgo8L2hlYWQ+Cjxib2R5Pgo8ZGl2IGNsYXNzPSJhdXRoLXNjcmVlbiIgaWQ9ImF1dGhTY3JlZW4iPgogIDxkaXYgY2xhc3M9ImF1dGgtY2FyZCI+CiAgICA8ZGl2IGNsYXNzPSJhdXRoLWJyYW5kIj4KICAgICAgPGltZyBjbGFzcz0iYnJhbmQtbG9nbyIgYWx0PSJMaWdvbi1SYXpvbiBTb2x1dGlvbnMgbG9nbyIgLz4KICAgICAgPHNwYW4gY2xhc3M9ImJyYW5kLXRpdGxlIj5Tb2NpYWwgTWVkaWEgQW5hbHl0aWNzPC9zcGFuPgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJhdXRoLWZvcm0iPgogICAgICA8ZGl2IGNsYXNzPSJmb3JtLWZpZWxkIj4KICAgICAgICA8bGFiZWwgZm9yPSJhdXRoQ29kZSI+QWNjZXNzIGNvZGU8L2xhYmVsPgogICAgICAgIDxpbnB1dCB0eXBlPSJwYXNzd29yZCIgaWQ9ImF1dGhDb2RlIiBhdXRvY29tcGxldGU9Im9mZiIgLz4KICAgICAgPC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImF1dGgtZXJyb3IiIGlkPSJhdXRoRXJyb3IiPjwvZGl2PgogICAgICA8YnV0dG9uIGNsYXNzPSJidG4gcHJpbWFyeSIgaWQ9ImF1dGhTdWJtaXRCdG4iIHR5cGU9ImJ1dHRvbiI+PGkgZGF0YS1sdWNpZGU9ImFycm93LXJpZ2h0IiBzdHlsZT0id2lkdGg6MTRweDtoZWlnaHQ6MTRweDsiPjwvaT4gRW50ZXI8L2J1dHRvbj4KICAgIDwvZGl2PgogIDwvZGl2Pgo8L2Rpdj4KCjxkaXYgY2xhc3M9ImFwcC1zaGVsbCIgaWQ9ImFwcFNoZWxsIiBzdHlsZT0iZGlzcGxheTpub25lOyI+CiAgPGFzaWRlIGNsYXNzPSJzaWRlYmFyIj4KICAgIDxkaXYgY2xhc3M9InNpZGViYXItYnJhbmQiPgogICAgICA8aW1nIGNsYXNzPSJicmFuZC1sb2dvIiBhbHQ9IkxpZ29uLVJhem9uIFNvbHV0aW9ucyBsb2dvIiAvPgogICAgICA8c3BhbiBjbGFzcz0iYnJhbmQtdGl0bGUiPlNvY2lhbCBNZWRpYSBBbmFseXRpY3M8L3NwYW4+CiAgICA8L2Rpdj4KICAgIDxuYXYgY2xhc3M9InRhYnMiIHJvbGU9InRhYmxpc3QiIGFyaWEtbGFiZWw9IlNlY3Rpb25zIj4KICAgICAgPGJ1dHRvbiBjbGFzcz0idGFiLWJ0biBpcy1hY3RpdmUiIGRhdGEtdGFiPSJkYXNoYm9hcmQiIHJvbGU9InRhYiIgYXJpYS1zZWxlY3RlZD0idHJ1ZSI+PGkgZGF0YS1sdWNpZGU9ImxheW91dC1kYXNoYm9hcmQiIHN0eWxlPSJ3aWR0aDoxNHB4O2hlaWdodDoxNHB4OyI+PC9pPiBEYXNoYm9hcmQ8L2J1dHRvbj4KICAgICAgPGJ1dHRvbiBjbGFzcz0idGFiLWJ0biIgZGF0YS10YWI9InJlY29yZHMiIHJvbGU9InRhYiIgYXJpYS1zZWxlY3RlZD0iZmFsc2UiPjxpIGRhdGEtbHVjaWRlPSJkYXRhYmFzZSIgc3R5bGU9IndpZHRoOjE0cHg7aGVpZ2h0OjE0cHg7Ij48L2k+IERhdGEgUmVjb3JkczwvYnV0dG9uPgogICAgICA8YnV0dG9uIGNsYXNzPSJ0YWItYnRuIiBkYXRhLXRhYj0iZm9sbG93ZXJzIiByb2xlPSJ0YWIiIGFyaWEtc2VsZWN0ZWQ9ImZhbHNlIj48aSBkYXRhLWx1Y2lkZT0idXNlcnMiIHN0eWxlPSJ3aWR0aDoxNHB4O2hlaWdodDoxNHB4OyI+PC9pPiBGb2xsb3dlcnMgRGF0YTwvYnV0dG9uPgogICAgICA8YnV0dG9uIGNsYXNzPSJ0YWItYnRuIiBkYXRhLXRhYj0iY29tcGFyaXNvbiIgcm9sZT0idGFiIiBhcmlhLXNlbGVjdGVkPSJmYWxzZSI+PGkgZGF0YS1sdWNpZGU9ImdpdC1jb21wYXJlIiBzdHlsZT0id2lkdGg6MTRweDtoZWlnaHQ6MTRweDsiPjwvaT4gQ29tcGFyaXNvbnM8L2J1dHRvbj4KICAgICAgPGJ1dHRvbiBjbGFzcz0idGFiLWJ0biIgZGF0YS10YWI9InVwbG9hZCIgcm9sZT0idGFiIiBhcmlhLXNlbGVjdGVkPSJmYWxzZSI+PGkgZGF0YS1sdWNpZGU9InVwbG9hZC1jbG91ZCIgc3R5bGU9IndpZHRoOjE0cHg7aGVpZ2h0OjE0cHg7Ij48L2k+IFVwbG9hZCBEYXRhPC9idXR0b24+CiAgICAgIDxidXR0b24gY2xhc3M9InRhYi1idG4iIGRhdGEtdGFiPSJoaXN0b3J5IiByb2xlPSJ0YWIiIGFyaWEtc2VsZWN0ZWQ9ImZhbHNlIj48aSBkYXRhLWx1Y2lkZT0iaGlzdG9yeSIgc3R5bGU9IndpZHRoOjE0cHg7aGVpZ2h0OjE0cHg7Ij48L2k+IFVwbG9hZCBIaXN0b3J5PC9idXR0b24+CiAgICA8L25hdj4KICAgIDxkaXYgY2xhc3M9InNpZGViYXItZm9vdGVyIj4KICAgICAgPGJ1dHRvbiBjbGFzcz0iYnRuIiBpZD0ibG9nb3V0QnRuIiB0eXBlPSJidXR0b24iPjxpIGRhdGEtbHVjaWRlPSJsb2NrIiBzdHlsZT0id2lkdGg6MTRweDtoZWlnaHQ6MTRweDsiPjwvaT4gTG9jazwvYnV0dG9uPgogICAgPC9kaXY+CiAgPC9hc2lkZT4KCiAgPGRpdiBjbGFzcz0ibWFpbi1jb2wiPgogIDxzZWN0aW9uIGNsYXNzPSJmaWx0ZXItYmFyIiBpZD0iZmlsdGVyQmFyIj4KICAgIDxkaXYgY2xhc3M9ImZpbHRlci1maWVsZCI+CiAgICAgIDxsYWJlbCBmb3I9ImZpbHRlckRhdGVGcm9tIj5Gcm9tPC9sYWJlbD4KICAgICAgPGlucHV0IHR5cGU9ImRhdGUiIGlkPSJmaWx0ZXJEYXRlRnJvbSIgLz4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0iZmlsdGVyLWZpZWxkIj4KICAgICAgPGxhYmVsIGZvcj0iZmlsdGVyRGF0ZVRvIj5UbzwvbGFiZWw+CiAgICAgIDxpbnB1dCB0eXBlPSJkYXRlIiBpZD0iZmlsdGVyRGF0ZVRvIiAvPgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJmaWx0ZXItZmllbGQgZmlsdGVyLXByZXNldHMiIGlkPSJmaWx0ZXJQcmVzZXRzIj4KICAgICAgPGJ1dHRvbiB0eXBlPSJidXR0b24iIGRhdGEtcHJlc2V0PSI3Ij5MYXN0IDcgZGF5czwvYnV0dG9uPgogICAgICA8YnV0dG9uIHR5cGU9ImJ1dHRvbiIgZGF0YS1wcmVzZXQ9IjMwIj5MYXN0IDMwIGRheXM8L2J1dHRvbj4KICAgICAgPGJ1dHRvbiB0eXBlPSJidXR0b24iIGRhdGEtcHJlc2V0PSI5MCI+TGFzdCA5MCBkYXlzPC9idXR0b24+CiAgICAgIDxidXR0b24gdHlwZT0iYnV0dG9uIiBkYXRhLXByZXNldD0iYWxsIj5BbGwgdGltZTwvYnV0dG9uPgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJmaWx0ZXItZmllbGQiPgogICAgICA8bGFiZWwgZm9yPSJmaWx0ZXJQbGF0Zm9ybSI+UGxhdGZvcm08L2xhYmVsPgogICAgICA8c2VsZWN0IGlkPSJmaWx0ZXJQbGF0Zm9ybSI+PG9wdGlvbiB2YWx1ZT0iYWxsIj5BbGwgcGxhdGZvcm1zPC9vcHRpb24+PC9zZWxlY3Q+CiAgICA8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImZpbHRlci1maWVsZCI+CiAgICAgIDxsYWJlbCBmb3I9ImZpbHRlckNhbXBhaWduIj5DYW1wYWlnbjwvbGFiZWw+CiAgICAgIDxzZWxlY3QgaWQ9ImZpbHRlckNhbXBhaWduIj48b3B0aW9uIHZhbHVlPSJhbGwiPkFsbCBjYW1wYWlnbnM8L29wdGlvbj48L3NlbGVjdD4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0iZmlsdGVyLWZpZWxkIj4KICAgICAgPGxhYmVsIGZvcj0iZmlsdGVyQ29udGVudFR5cGUiPkNvbnRlbnQgdHlwZTwvbGFiZWw+CiAgICAgIDxzZWxlY3QgaWQ9ImZpbHRlckNvbnRlbnRUeXBlIj48b3B0aW9uIHZhbHVlPSJhbGwiPkFsbCBjb250ZW50IHR5cGVzPC9vcHRpb24+PC9zZWxlY3Q+CiAgICA8L2Rpdj4KICA8L3NlY3Rpb24+CgogIDxtYWluIGNsYXNzPSJ2aWV3LWFyZWEiPgogICAgPHNlY3Rpb24gaWQ9InZpZXctZGFzaGJvYXJkIiBjbGFzcz0idmlldyBpcy1hY3RpdmUiPjwvc2VjdGlvbj4KICAgIDxzZWN0aW9uIGlkPSJ2aWV3LXJlY29yZHMiIGNsYXNzPSJ2aWV3Ij48L3NlY3Rpb24+CiAgICA8c2VjdGlvbiBpZD0idmlldy1mb2xsb3dlcnMiIGNsYXNzPSJ2aWV3Ij48L3NlY3Rpb24+CiAgICA8c2VjdGlvbiBpZD0idmlldy1jb21wYXJpc29uIiBjbGFzcz0idmlldyI+PC9zZWN0aW9uPgogICAgPHNlY3Rpb24gaWQ9InZpZXctdXBsb2FkIiBjbGFzcz0idmlldyI+PC9zZWN0aW9uPgogICAgPHNlY3Rpb24gaWQ9InZpZXctaGlzdG9yeSIgY2xhc3M9InZpZXciPjwvc2VjdGlvbj4KICA8L21haW4+CiAgPC9kaXY+CjwvZGl2PgoKPGRpdiBpZD0idG9hc3RSb290IiBjbGFzcz0idG9hc3Qtcm9vdCIgYXJpYS1saXZlPSJwb2xpdGUiPjwvZGl2PgoKPHNjcmlwdD4KLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIEJyYW5kIGxvZ28g4oCUIGVtYmVkZGVkIG9uY2UgaGVyZSBhbmQgd2lyZWQgb250byBldmVyeSAuYnJhbmQtbG9nbwogICA8aW1nPiBhbmQgdGhlIGZhdmljb24gPGxpbms+IGF0IGJvb3RzdHJhcCwgc28gdGhlIGJhc2U2NCBwYXlsb2FkCiAgIGFwcGVhcnMgZXhhY3RseSBvbmNlIGluIHRoaXMgZmlsZSBpbnN0ZWFkIG9mIG9uY2UgcGVyIHVzYWdlIHNpdGUuCiAgID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwpjb25zdCBMT0dPX0RBVEFfVVJJID0gJ2RhdGE6aW1hZ2UvcG5nO2Jhc2U2NCxpVkJPUncwS0dnb0FBQUFOU1VoRVVnQUFDRm9BQUFkekNBWUFBQUJuYjhvM0FBQUFDWEJJV1hNQUFDNGpBQUF1SXdGNHBUOTJBQUFnQUVsRVFWUjRuT3pkVFc3YldMb0c0T1BnenVOYWdWMHJNR3NGVVUwOFRTNElHSnBGTzRoM1FIRUh6Z3JhTlJNSUdKVk1PYm5PQ29wWlFUczdTRmFnaXhNeHNWWDVvMjM5SEpMUEEzall5TkYzSkhXaitlcDdENWJMWlFBQUFBQUFBQUFBNE5lZW1CRUFBQUFBQUFBQVFEZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBQUFBQUIwSkdnQkFBQUFBQUFBQU5DUm9BVUFBQUFBQUFBQVFFZUNGZ0FBQUFBQUFBQUFIUWxhQUFBQUFBRDBVRjdWV1Y3VkYrNE9BQUIyNjMvTUd3QUFBQUNnbHk1ZEd3QUE3SjZnQlFBQUFBQkF6K1JWUFE4aG5MZzNBQURZUGRVaEFBQUFBQUE5a2xmMWNRamgvTXVKODZxZXVEOEFBTmdkUVFzQUFBQUFnSDZKbFNGUDc1dzRjMzhBQUxBN2doWUFBQUFBQUQyUlYzWGNaUEhzWDZlMTBRSUFBSFpJMEFJQUFBQUFvQWZ5cWo0TUljeS9jMUliTFFBQVlJY0VMUUFBQUFBQSt1SGZsU0ZmSExVaERBQUFZQWNFTFFBQUFBQUFFcGRYOVlzUXd2T2ZuRko5Q0FBQTdJaWdCUUFBQUFCQXd0cHRGWmUvT0tINkVBQUEyQkZCQ3dBQUFBQ0F0TTEvVUJseWw0MFdBQUN3STRJV0FBQUFBQUNKeXFzNkJpaGVkVGlkalJZQUFMQWpnaFlBQUFBQUFBbnFXQm55eGRPOHFvVXRBQUJnQndRdEFBQUFBQURTZEI1Q09Mckh5UVF0QUFCZ0J3UXRBQUFBQUFBUzAyNm5LTzU1cW9sN0JBQ0E3Uk8wQUFBQUFBQklUOWZLa0x0c3RBQUFnQjBRdEFBQUFBQUFTRWhlMWZNUXdza0RUdlNRL3d3QUFIQlBnaFlBQUFBQUFJbklxL280aEhEKzBOUGtWYTArQkFBQXRrelFBZ0FBQUFBZ0hiRXk1T2tqVGlOb0FRQUFXeVpvQVFBQUFBQ1FnTHlxNHlhTFo0ODhTZVl1QVFCZ3V3UXRBQUFBQUFEMkxLL3F3eERDZkFPbkVMUUFBSUF0RTdRQUFBQUFBTmkveDFhR2ZIR1VWL1d4K3dRQWdPMFJ0QUFBQUFBQTJLTzhxbCtFRUo1djhBUzJXZ0FBd0JZSldnQUFBQUFBN0VsYkdYSzU0WDlkMEFJQUFMWkkwQUlBQUFBQVlIL21HNm9NdVd2aVBnRUFZSHNFTFFBQUFBQUE5aUN2NmhpSWVMV0ZmL21aK3dRQWdPMFJ0QUFBQUFBQTJMRXRWWVo4bFZlMStoQUFBTmdTUVFzQUFBQUFnTjA3RHlFY2JmRmZWUjhDQUFCYkltZ0JBQUFBQUxCRDdiYUpZc3Yvb28wV0FBQ3dKWUlXQUFBQUFBQzd0YlhLa0RzRUxRQUFZRXNFTFFBQUFBQ0FyY2tXNVN4YmxCNzZ0L0txbm9jUVRuYndUNTNrVlgyNGczOEhBQUJHNTM5Y09RQUFBQUN3YVRGZ0VVS0lvWUtiWmxwTURQaHp5T0o0QjVVaGQ4V0F5L1VPL3owQUFCZ0ZRUXNBQUFBQVlHUHVCQ3lPUWdpZlFnZ3owLzFxRjVVaGQwMEVMUUFBWVBNRUxRQUFBQUNBUjhzVzVZc1F3a1Vic1BoaTNreUxHOVA5dk0zaVBJVHdiTWYvck1vV0FBRFlBa0VMQUFBQUFPREJza1U1YVRkWS9EdEU4SzZaRmhjbSs3VXlaTDZIZjFwbEN3QUFiTUVUUXdVQUFBQUE3aXNHTExKRkdXc3AvdThIbXhyT0RmV3JHRGg1dW9kLzkya2I4Z0FBQURiSVJnc0FBQUFBb0xPZmJMQzRxMnltUldPcW43ZFp4RXFWNTNzOFFxd1BVZDhDQUFBYkpHZ0JBQUFBQVB4U3g0QkY5TDZaRnZ1b3lVaE9YdFdISVlUTFBaOHIzdHViQk1jREFBQzlKV2dCQUFBQUFQeFF0aWlQMjRERnk0NVRtcG5tVi91cURMa3IyL08vRHdBQWd5Tm9BUUFBQUFCODR3RUJpK2kxeXBDVnZLb245NXpkdHZ4cUF3a0FBSEJQQjh2bDBzd0FBQUFBZ004ZUdMQ0lQc1QvZURNdFBvNTlrbTFsU0F5Y0hDVnduT2pQcTdQVDZ3VE9BUUFBZzJDakJRQUFBQUR3bUlERkZ6TWhpNi9tQ1lVc1Fsc2ZJbWdCQUFBYkltZ0JBQUFBQUNPV0xjcTRmZUU4aEZBOFlncXhNc1NEL05VMml4aHFlSlhBVWU3SzBqa0tBQUQwbjZBRkFBQUFBSXpRbllCRi9IdjZpQWw4YURjNHNIS1o0QndtQ1p3QkFBQUc0MkM1WExwTkFBQUFBQmlKRFFZc3Z2amZabHE4OGY3NXZNMWkvc2pOSU52MDI5WFpxV29YQUFEWUFCc3RBQUFBQUdBRXRoQ3dpTjRLV2F6a1ZYMmNjTWdpdFBVaDZsMEFBR0FEbmhnaUFBQUFBQXhidGlobklZU2JOZ2l3cVpERnB4REN6RnZucXhRclErNVNId0lBQUJ0aW93VUFBQUFBREZRYnNJaDFGa2RiZUlXelpscW9vbGh0czRoYlFwNGxjSlNmRWJRQUFJQU5FYlFBQUFBQWdJSFpjc0FpZXFjeVpLV3RESm1uY0paZnlKSStIUUFBOUlqcUVBQUFBQUFZaUJpd3lCWmxyQWo1enhaREZpcEQxbDFzc0k1bG01NjJvUkFBQU9DUmJMUUFBQUFBZ0o3TEZtV3NoYmpjWXJqaXJua3pMVzY4Wno1dnMzZ1JRbmlld0ZHNit2SStBUUFBSGtIUUFnQUFBQUI2cWcxWXhOcUtaenQ2QmJFeTVNTDc1WFBJNHJDSG9RWDFJUUFBc0FHQ0ZnQUFBQURRTTNzSVdBU1ZJZC9vUzJYSVhaTjBqZ0lBQVAwbGFBRUFBQUFBUGJHbmdNVVhGeXBEVnZLcWp2ZndNb1d6M05OSnIwNExBQUNKT2xndWwrNEdBQUFBQUJLV0xjcmp0cVppSHdHTDZIMHpMZFJPM0ZhR05DR0Vvd1NPOHhCL1hwMmRYdmZ2MkFBQWtBNGJMUUFBQUFBZ1VXM0FZcDdBOWdTVkliZm1QUTVaUkRFd0kyZ0JBQUNQSUdnQkFBQUFBSWxKS0dBUmxjMjBhQkk0eDk2MWxTR3ZldjR5NG11NFNPQWNBQURRVzRJV0FBQUFBSkNJeEFJVzBZZG1Xc3dUT0VjcWhoQlFVQUVEQUFDUEpHZ0JBQUFBQUh1V0xjckQ5aUYrS2dHTEwxU0d0UEtxam9HVGt5UU84emhIZVZVZlhwMmRmdXp6aXdBQWdIMFN0QUFBQUFDQVBXa0RGdWZ0MzlQRTd1RjFNeTJ1RXpqSDN1VlZIYmRBRkFONlNiRSs1RTBDNXdBQWdGNFN0QUFBQUFDQUhVczhZQkY5YUN0TVdCbENaY2hkbWFBRkFBQThuS0FGQUFBQUFPeElEd0lXWDh5YWFhRmFZclhOSXQ3VnN3U09za21UNGJ3VUFBRFlQVUVMQUFBQUFOaXlIZ1Vzb3I5VWhxemtWWDA4ME0wZVF3dU9BQURBVGoweGJnQUFBQURZbm14UnprSUlUUWloNkVISTRsTWJCbUhsc2dkMzlpQjVWV2M5UERZQUFDVEJSZ3NBQUFBQTJJSTJZQkczSVJ6MWFMNHFRMXA1VmI4WStPYUhyQTBBQVFBQTl5Um9BUUFBQUFBYjFOT0FSZlMybVJadkVqakgzdVZWZmRodXN4aXl5UWhlSXdBQWJJV2dCUUFBQUFCc1FJOERGcUd0REprbGNJNVVETFl5NUE3VklRQUE4RUNDRmdBQUFBRHdDTm1pbkxRQml6N1hUTXhWaHF6a1ZSM3Y4M2tLWjlteWswRy9PZ0FBMktLRDVYSnB2Z0FBQUFCd1R3TUpXRVR2bW1reFNlQWNlOWRXaGpROTNVcnlFSDllbloxZTkrL1lBQUN3WHpaYUFBQUFBTUE5RENoZ0VWU0dmS092MVM4UEZkL0xnaFlBQUhCUGdoWUFBQUFBME1IQUFoWmZ4TXFRbXpTT3NsOXRaY2lya2Izc0xJRXpBQUJBN3doYUFBQUFBTUJQWkl2eU9JUndFVUo0UHJBNXhjcVFpd1RPa1lveHprTFFBZ0FBSHVDSm9RRUFBQURBdDJMQUlsdVVseUdFL3c0d1pCR2RKM0NHSk9SVkhUZVZuSXp3cFIvbFZYMmN3RGtBQUtCWGJMUUFBQUFBZ0R2YURSYnh3ZnZMQWMrbGJLWkZrOEE1OWk2djZyalZvUmp4Q09MclZ4OERBQUQzSUdnQkFBQUFBT01KV0VUdm0ya3hUK0FjcVJoN2Zjb2toUEFtZ1hNQUFFQnZDRm9BQUFBQU1HclpvanhzQXhhdlJqS0hXUUpuU0VKZTFiRSs1ZG5JeDVBbGNBWUFBT2dWUVFzQUFBQUFScWtOV0p5M2YwOUhNb1BYS2tOVzhxcitzc0ZrN01ZZU5BRUFnSHNUdEFBQUFBQmdWRVlhc0lnK0NCYXN1UnpaL2Y5UVh0WFoxZG1wQUE0QUFIUWthQUVBQUFEQUtJdzRZUEhGckprV0g5TTR5bjdsVmYzQ0pvYzFreENDb0FVQUFIUWthQUVBQUFEQTRHV0w4cnpkNWpEV0RRYXhNdVE2Z1hQc1hWN1ZoKzAyQzI1bFpnRUFBTjBKV2dBQUFBQXdXTm1pbkxVQmk2TVIzN0xLa0hVcVE3NGxhQUVBQVBkd3NGd3V6UXNBQUFDQVFSR3dXUE8vemJSNGs5QjU5cWF0RFBsN3BDLy9WMzY3T2p0VkxRTUFBQjNZYUFFQUFBREFZQWhZZk9PdGtNVktXeGx5a2NKWkVoVzNXcWlYQVFDQURnUXRBQUFBQU9pOWJGRythQitpQzFqYytoUkNtS1Z5bUFRSTRQemNSTkFDQUFDNkViUUFBQUFBb0xleVJUbHBINkEvYzR2Zm1EWFRRaFhFYXB0RmZKKzhTdUFvS2N2R1BnQUFBT2hLMEFJQUFBQ0EzaEd3K0tWM0trUFdYQ1owbGxSTnhqNEFBQURvNm9sSkFRQUFBTkFYTVdDUkxjcFliL0IvUWhZL3BETGtqcnlxVllaMDh6U3Y2dU0rSEJRQUFQYk5SZ3NBQUFBQWtwY3R5bGhyY0NGYzBjbThtUlkzUFRqbjF1VlZIZDgzeGNCZjVpWk5iUDhBQUlCZkU3UUFBQUFBSUZuWm9qeHVLMEpldXFWT1ltWElSUS9PdVN0Q0EvZVQ5ZW13QUFDd0w0SVdBQUFBQUNSSHdPSkJWSWJja1ZmMWVRamhKSmtEOVlPZ0JRQUFkQ0JvQVFBQUFFQXlCQ3dlNVVKbHlFcGUxVi9lUjl5UGFoNEFBT2pnWUxsY21oTUFBQUFBZXlWZzhXanZtMmxoRzBFcnIrcHJvWUVIKy9QcTdQUzZwMmNIQUlDZHNORUNBQUFBZ0wzSkZ1VmhDT0c4L1h2cUpoNU1aVWdycitxWmtNV2p4TUNPb0FVQUFQeUVvQVVBQUFBQU95ZGdzVkZsTXkyYUFiMmVCOHVyT3I2dkxucDYvRlRZakFJQUFMOGdhQUVBQUFEQXpnaFliTnlIWmxyTUIvYWFIdVBTKytyUkpqMC9Qd0FBYkoyZ0JRQUFBQUJiSjJDeE5TcERXbmxWdndnaFBFL2lNUDEyRkRlRFhKMmRmaHo3SUFBQTRFZWVtQXdBQUFBQTI1UXR5aGdHdUFraEZFSVdHL1c2bVJiWEEzbzlENll5Wk9QVWh3QUF3RS9ZYUFFQUFBREFWclFCaTFocmNXVENHL2VoblMwcjNtZWJGZXREaEhnQUFPQUhCQzBBQUFBQTJDZ0JpNTJZTmROQ3RjTnFtMFVNQmJ4SzRDaERNaG43QUFBQTRHY0VMUUFBQUFEWUNBR0xuZmxMWmNpYXk0VE9NaFNxUXdBQTRDZWVHQTRBQUFBQWo1RXR5a20yS0pzUXduK0VMTGJ1VXdqaGZPQ3ZzYk84cWdWN3R1TnBYdFhDRmdBQThBTTJXZ0FBQUFEd0lERmcwVzZ3ZUdhQ082TXlwTlVHQVlva0RqTk1jYjdOMkljQUFBRGZJMmdCQUFBQXdMMElXT3pOMjJaYXZCbnBhLzhlbFNIYlphTUZBQUQ4Z0tBRkFBQUFBSjBJV094VnJBeVpqZmoxcjhtck90YW5uQ1IwcENHYWpIMEFBQUR3STRJV0FBQUFBUHhVdGlpUDIrMEJBaGI3TTFjWnNwSlg5WEViK0dHN0JGa0FBT0FIbmhnTUFBQUFBTjhUQXhiWm9vd0JpLzhLV2V6VnUyWmFYSXo0OWY5YmZFOCtUZXRJdzVSWHRhMFdBQUR3SFRaYUFBQUFBTENtM1dBUk53YThOSm05VXhseVIxc1pJdlN6TzFrSTRYb3NMeFlBQUxvU3RBQUFBQURnTXdHTEpNWEtrSnV4RHlHc1FoYUhLa04yTG02MHNFMEZBQUQrUmRBQ0FBQUFZT1N5UmZubEFmYXJzYzhpTVNwRDFxa00yYjFzYkM4WUFBQzZPRmd1bHdZRkFBQUFNRUp0d09LOC9mTUFPejEvTk5PaUdmc1F3bXFieFlzUXd0OEpIR1dNZnJzNk8vMDQ5aUVBQU1CZE5sb0FBQUFBakl5QVJTK1VRaFlyYldYSVpRcG5HYWxZSC9KbTdFTUFBSUM3QkMwQUFBQUFSa0xBb2pmZU45TmlQdlloM0RIM2Z0MnJUTkFDQUFEV0NWb0FBQUFBakVDMktPY0NGcjB4Ry9zQXZzaXJPbTVUZUpYR2FVWnJNdllCQUFEQXZ3bGFBQUFBQUF4WXRpaG43VWFBSS9mY0M2OVZocXlvREVuR3M3RVBBQUFBL3UySmlRQUFBQUFNVHd4WVpJdnlKb1R3SHlHTDN2alFobUpZT2ZmZVRVTmUxZG5ZWndBQUFIZlphQUVBQUFBd0lEWlk5TnFzbVJZZnh6NkVjUHRndjBqZ0tLekUrN0JwQlFBQVdvSVdBQUFBQUFNZ1lORjdzVExrZXV4RHVFTmxTRm9tN2dRQUFHNEpXZ0FBQUFEMFdMWW9KMjNBNHBsNzdDMlZJWGZrVlIxbmNaTE1nUWp0UmdzQUFLQWxhQUVBQUFEUVF3SVdnM0t1TW1RbHIrcmpPSThVenNJYXdSY0FBTGpqaVdFQUFBQUE5RWNNV0dTTE1sWk0vSitReFNDOGJhYkZtN0VQNFk1WVQvRTBtZFB3VlY3VkU5TUFBSUFWR3kwQUFBQUFlc0FHaTBINkZFS1lqWDBJWCtSVmZlNzluYlQ0SFhROTlpRUFBRUFRdEFBQUFBQklXN1lvajl1QXhVdFhOVGd6bFNFcmVWVWZ0dTl6MHBXNUd3QUFXQkcwQUFBQUFFaVFnTVhndlZNWnNrWmxTUHBVaHdBQVFPdUpRUUFBQUFDa0l3WXNza1VaSHpyL1Y4aGlzRlNHM0pGWDlZc1F3dk5rRHNTUFBNMnIrdGgwQUFEQVJnc0FBQUNBSk5oZ01TcnpabHJjakgwSTRiWXk1REtCbzlCTnJBL3gzZ1VBWVBRRUxRQUFBQUQyS0Z1VThVSHplUWloY0ErakVDdERMc1kraER2bUtrTjZKZGFIcUx3QkFHRDBCQzBBQUFBQTl1Qk93T0xjZytiUlVCbHlSMTdWOGFIOXEyUU9SQmVaS1FFQWdLQUZBQUFBd0U0SldJemFoY3FRRlpVaHZmVnM3QU1BQUlBZ2FBRUFBQUN3R3dJV28vZSttUmJ6c1EvaGp2ZzVPRXJtTkhTV1YzVjJkWGJhbUJnQUFHUDJ4TzBEQUFBQWJGZTJLR05kUk54a1VBaFpqSmJLa0ZaOFVOOStGdWluaVhzREFHRHNiTFFBQUFBQTJKSTJZREgzeS8zUks1dHBZUVBBTFpVaC9aYU5mUUFBQUNCb0FRQUFBTEJoQWhiYzhVRmx5SzI4cXVNc1RsSTVEdzhpYUFFQXdPZ2RMSmZMc2M4QUFBQUFZQ01FTFBpT1A1dHBjVzB3bjBNV3h5R0VSbjNPSVB4MmRYYjZjZXhEQUFCZ3ZHeTBBQUFBQUhpa2JGRk8yam9FQVF2dWVpMWtzZVpTeUdJdzRsWUw3MjBBQUVaTDBBSUFBQURnZ2RxQVJkeGc4Y3dNK1pjUDdYdUQxVGFMYzUrVFFaa0lXZ0FBTUdhQ0ZnQUFBQUQzSkdCQkI3Tm1XcWhXdUswTUVUb1psc25ZQndBQXdMZ0pXZ0FBQUFCMEpHQkJSMytwREZsem9USmtjTEt4RHdBQWdIRTdXQzZYWTU4QkFBQUF3RTlsaXpKckh4WUxXUEFybjBJSXg3WlpyT1JWL1NLRThIY0taMkhqZnI4Nk83MHhWZ0FBeHNoR0N3QUFBSUFmeUJibGw4cURsMlpFUnlwRFdubFZINFlRTHBNNEROc3djYjhBQUl5Vm9BVUFBQURBdndoWThFQnZtMm54eHZDK1Voa3liT3BEQUFBWUxVRUxBQUFBZ0phQUJZOFFLME5tQnJpU1YvWEU1Mmp3QkMwQUFCZ3RRUXNBQUFCZzlMSkZlZGorK3Q2RFlSNXFyakprUldYSWFEd2Ird0FBQUJpdkorNGVBQUFBR0tzWXNNZ1daZHhnY1NOa3dTTzhhNmJGaFFGK0ZUOVRSNG1jaFMxcU41Y0FBTURvMkdnQkFBQUFqRTY3d2VLOC9YdnFIY0FqcUF5NUk2L3FXQ2Z4S3BrRHNXM3h2cTlOR1FDQXNSRzBBQUFBQUVaRHdJSXRpSlVoTndiN2xjcVFjY25HUGdBQUFNWkowQUlBQUFBWVBBRUx0a1JseUIxNVZjZktrSk5rRHNRdXFBNEJBR0NVbnJoMkFBQUFZTWl5UlJsckhab1FRaUZrd1lhZEcraEtYdFhIN1dlTWNUbktxL3JRblFNQU1EWTJXZ0FBQUFDRDFBWXM0aS9zajl3d1cxQTIwNkl4Mks5VWhveFgzR3J4WnV4REFBQmdYQVF0QUFBQWdFRVJzR0FIM2pmVFltN1FLM2xWeDgwZXoxSTRDM3VSQ1ZvQUFEQTJnaFlBQUFEQUlBaFlzRU16dzE1cEswT0VUc1p0TXZZQkFBQXdQb0lXQUFBQVFLOWxpM0xTUHVqMWkzcDI0YlhLa0RVWElZU25DWjJIM2N2TUhBQ0FzVGxZTHBjdUhRQUFBT2dkQVF2MjRFTjg2elhUNHFQaGY5NW04U0tFOEhjQ1IySC8vcmc2T3hWQUFnQmdOR3kwQUFBQUFIcEZ3SUk5bWdsWnJPUlZmUmhDdUV6aExDUWhiclVRdEFBQVlEUUVMUUFBQUlCZUVMQmd6MkpseUxWTCtFcGxDSGVwRHdFQVlGUUVMUUFBQUlDa1pZdnl1SDJvKzl4TnNTY2YycEFQcTIwV01mVDAwaXk0WTJJWUFBQ01pYUFGQUFBQWtLUTJZREgzUUpjRW5Lc01XVkVad2crY0dBd0FBR1B5eEcwREFBQUFLWWtCaTJ4UnhnZTUveFd5SUFGdm0ybnh4a1Y4RmNOUFI0bWNoWVMwbTA0QUFHQVViTFFBQUFBQWttQ0RCUW42RkVLWXVaaVZ2S3F6RU1LckZNNUNrdUw3NDlyVkFBQXdCb0lXQUFBQXdGNWxpL0t3RFZoNGdFdHFaaXBEMXFnTTRXZmlSb3NMRXdJQVlBd0VMUUFBQUlDOWFBTVc1KzNmVTdkQVl0NnBETG1WVjNVTVE1MmtjaDZTbExrV0FBREc0bUM1WExwc0FBQUFZR2NFTE9pQldCbVNOZFBpeG1WOXJRejVKNEdqa0w3ZnI4NU9mVzRBQUJnOEd5MEFBQUNBblJDd29FZm1RaFpyMUVIUVZRemwrT3dBQURCNFQxd3hBQUFBc0czWm9qeHZINzRWUWhZa0xsYUdDQmEwOHFxT245MW5TUnlHUGxBZkFnREFLTmhvQVFBQUFHeE50aWhuY1R0QUNPSElsT21CV0JreWMxRXJlVlVmdDU5ZjZHcGlVZ0FBaklHZ0JRQUFBTEJ4QWhiMDFJWEtrRFdYTnRCd1Q3YWZBQUF3Q2dmTDVkSk5Bd0FBQUJzaFlFR1B2VyttaGRxRFZsN1ZMMElJZnlkeEdQcm1qNnV6MDhhdEFRQXdaRFphQUFBQUFJK1dMY3I0VVBaQ3dJSWVVeG5TeXF2NnNOMW1BUThSQTB1Q0ZnQUFESnFnQlFBQUFQQmcyYUtjdEJzc3JJdW56OHBtV25nd2ZFdGxDSTh4RWRRQkFHRG9CQzBBQUFDQWV4T3dZRUErTk5OaTdrSlg4cXFPbiszbktaeUYzbExCQXdEQTRCMHNsMHUzREFBQUFIUWlZTUVBL2RsTWkyc1grN1V5cEZFQnhBYjhkblYyK3RFZ0FRQVlLaHN0QUFBQWdGL0tGbVg4aGZLRmdBVUQ4MXJJWXMxY3lJSU5pZitkNGJNRkFNQmdDVm9BQUFBQVA1UXR5dVAyNGV0TFUySmdQclR2Ylc0clExNlpCUnN5RWJRQUFHRElCQzBBQUFDQWJ3aFlNQUt6Wmxxb05yaDFrY3BCR0lUTU5RSUFNR1NDRmdBQUFNQlhBaGFNeEY4cVEyN2xWUjAvOHllcG5JZEJtTGhHQUFDRzdHQzVYTHBnQUFBQUdEa0JDMGJrVXdqaDJEYUxsYnlxNCthQmYxSTRDNFB6KzlYWjZZMXJCUUJnaUd5MEFBQUFnQkhMRnVWaENPRzgvWHZxdmNBSXFBeFpwektFYllraEhrRUxBQUFHU2RBQ0FBQUFSa2pBZ3BGNjIweUxOeTUvSmEvcStQbC9sc0paR0tSWUgrTHpCZ0RBSUFsYUFBQUF3SWdJV0RCaXNUSms1ZzJ3a2xmMWw3b2cySmJNWkFFQUdDcEJDd0FBQUJnQkFRc0ljNVVoYXk1OUY3Qmx0cVVBQURCWUI4dmwwdTBDQUFEQWdHV0xNdjZLLzhKRFZVYnNYVE10SnQ0QUszbFZ2d2doL0ozQ1dSaThQNjdPVGh2WERBREEwTmhvQVFBQUFBUFZCaXhpTmNDUk8yYkVWSWJja1ZmMVlidk5BblloQnB3RUxRQUFHQnhCQ3dBQUFCZ1lBUXRZRXl0RGJvemtLNVVoN0ZKbTJnQUFESkdnQlFBQUFBeUVnQVY4STFhR1hCakxTbDdWY2J2QTh4VE93bWlvN0FFQVlKQU9sc3VsbXdVQUFJQWV5eFpsZkpBVkh5YWZ1RWRZODBjekxkUVczRmFHTklKWTdNRnZWMmVuSHcwZUFJQWhzZEVDQUFBQWVxb05XTVFORnMvY0lYeWpGTEpZWTlzTit4THJRNjVOSHdDQUlSRzBBQUFBZ0o0UnNJQmZldDlNaTdreHJiU1ZJYTlTT0F1ak5CRzBBQUJnYUFRdEFBQUFvQ2NFTEtDem1WR3R1VXpvTEl6UHhKMERBREEwZ2hZQUFBQ1F1R3hSSHJjUFNnVXM0TmRlcXd5NWxWZTF5aEQyTFhNREFBQU16Y0Z5dVhTcEFBQUFrS0EyWUJFZmtyNTBQOURKaC9qUmFhYkZSK1A2SExLSUQ3ai9TZUFvOFB2VjJlbk42S2NBQU1CZzJHZ0JBQUFBaVJHd2dBZWJDVm1zVVJsQ0tpYmVqd0FBREltZ0JRQUFBQ1JDd0FJZUpWYUdYQnZoU2w3VjV5R0VreFRPQXVwREFBQVlHa0VMQUFBQTJMTnNVUjYyQVl0WDdnSWU1RVA3R1dJVnNqZzJEeElqYUFFQXdLQUlXZ0FBQU1DZXRBR0w4L2J2cVh1QUJ6dFhHYkxtMG5jS2lYbm1RZ0FBR0pLRDVYTHBRZ0VBQUdDSEJDeGdvOTQyMCtLRmthN2tWVDBMSWZ3bmhiUEF2L3g1ZFhhcTNnY0FnRUd3MFFJQUFBQjJSTUFDTnU1VENHRm1yQ3Q1VmNmdm1Jc1V6Z0xmRWV0REJDMEFBQmlFSjY0UkFBQUF0aTlibFBNUXdrMElvUkN5Z0kyWnFReFpvektFbEUzY0RnQUFRMkdqQlFBQUFHeFJ0aWpqcisxanlPTEluR0dqM2pYVDRvMlJydVJWSGV0VG5xZHdGdmlCekdBQUFCaUtnK1Z5NlRJQkFBQmd3d1FzWUt0aVpValdUSXNiWS81YUdkTDR2cUVIZnJzNk83V0ZCZ0NBM3JQUkFnQUFBRFpJd0FKMllpNWtzY1ozRG4wUjYwTnNvZ0VBb1BjRUxRQUFBR0FEQkN4Z1oySmx5SVZ4citSVkhSOWN2MHJoTE5CQkptZ0JBTUFRQ0ZvQUFBREFJMlNMY3RJR0xKNlpJMnhkckF5WkdmT2F5NFRPQXI4eU1TRUFBSVpBMEFJQUFBQWVRTUFDOXVKQ1pjaXR2S3B0MGFGdk1qY0dBTUFRSEN5WFN4Y0pBQUFBSFFsWXdONjhiNmFGaDdTdHZLcmpMUDVKNGpCd1AzOWNuWjAyWmdZQVFKL1phQUVBQUFBZENGakEzcWtNV2FjeWhMNktJU0ZCQ3dBQWVrM1FBZ0FBQUg0aVc1VEhiY0RpcFRuQjNwVE50UEJndHBWWDlYa0k0U1NKdzhEOTJVd0RBRUR2Q1ZvQUFBREFkd2hZUURJK05OTmk3anBXOHFyKzh0MEVmVFZ4Y3dBQTlKMmdCUUFBQU53aFlBSEpVUm15TGxhR1BFM3BRSEJQdHJFQUFOQjdCOHZsMGkwQ0FBQXdlZ0lXa0tUWHpiUTRkelVyZVZYSDBNbC9VamdMUE5LZlYyZW4xNFlJQUVCZjJXZ0JBQURBcUdXTDhqQ0VFQi9rRm1PZkJTVG1nNHFNVzNsVngrK3FpMVRPQTQ4VTYwTUVMUUFBNkMxQkN3QUFBTU05aHNzQUFDQUFTVVJCVkVicFRzRGkzQnArU05Lc21SWWZYYzFYS2tNWWtzeHRBZ0RRWjRJV0FBQUFqSXFBQmZUQ1g4MjA4R3YzVmw3VkwwSUl6NU00REd5R29BVUFBTDEyc0Z3dTNTQUFBQUNESjJBQnZmRXBoSEJzbThWS1d4bHk0M3VMQWZyOTZ1ejB4c1VDQU5CSE5sb0FBQUF3ZU5taW5JVVFManlvaEY1UUdiSnU3cnVMZ2NyYUVCRUFBUFNPb0FVQUFBQ0QxUVlzNGtQS0k3Y012ZkMybVJadlhOVktYdFdURU1LckZNNENXeENERmo3dkFBRDBrcUFGQUFBQWd5TmdBYjBVSzBObXJtNmxyUXk1VE9Fc3NDVVRnd1VBb0s4RUxRQUFBQmdNQVF2b3RibktrRFhudnNzWXVHY3VHQUNBdmpwWUxwY3VEd0FBZ0Y3TEZ1V2svZVczaDVMUVQrK2FhZUhYN2EyOHFtT2x3ajlKSEFhMjY0K3JzOVBHakFFQTZCc2JMUUFBQU9pdE5tQXg5NnRZNkRXVklkOVNHY0pZeEZDUm9BVUFBTDBqYUFFQUFFRHZDRmpBb01US2tCdFh1cEpYZGZ4dU8wbmhMTEFERThFaUFBRDZTTkFDQUFDQTNoQ3dnTUdKbFNFWHJuVWxyK3JqRU1KNUNtZUJIY2tNR2dDQVBoSzBBQUFBSUhuWm9vd1BZaTRFTEdCd2hBcld4Vi8yUDAzcFFMQmxKM2xWSDE2ZG5YNDBhQUFBK2tUUUFnQUFnR1JsaS9LNDNXRHgwaTNCNEpUTnRHaGM2MHBlMWVmQ1pJeFVERk5ldTN3QUFQcEUwQUlBQUlEa0NGakE0TDF2cHNYY05hL0VYL1MzMzNrd1JoTkJDd0FBK2tiUUFnQUFnR1FJV01Cb3pGejFHcFVoakZubTlnRUE2QnRCQ3dBQUFQWXVXNVR4MTl3WEFoWXdDcTlWaHR6S3EvcEZDT0Y1S3VlQlBaZ1lPZ0FBZlhPd1hDNWRHZ0FBQUh2UkJpek8yeisvNW9iaCt4QS8rczIwK09pdXYxYUczUGorZy9ENzFkbnBqVEVBQU5BWE5sb0FBQUN3Y3dJV01Gb3pJWXMxYzkrQjhGbldobzRBQUtBWEJDMEFBQURZR1FFTEdMVllHWEk5OWlGOGtWZDFyRXQ0bGNacFlPL2k1K0dOYXdBQW9DOEVMUUFBQU5nNkFRc1l2US90OWdadUswTXV6UUsreW93Q0FJQStFYlFBQUFCZ3E3SkZPV3Nmc0I2Wk5Jeld1Y3FRTmVlK0UySE5NK01BQUtCUERwYkxwUXNEQUFCZzR3UXNnTmJiWmxxOE1JeVZ2S3JqTC9mL1NlRXNrSmcvcjg1TzFRc0JBTkFMTmxvQUFBQ3dVUUlXd0IyZlFnZ3pBMW1qTWdTK0w0YVFCQzBBQU9nRlFRc0FBQUEyUXNBQytJNlp5cEJiZVZYSDc4aVRWTTREaWNsY0NBQUFmU0ZvQVFBQXdLTmtpM0xTQml6MHF3TjN2V3VteFJzVFdjbXIramlFY0o3Q1dTQlJFeGNEQUVCZkNGb0FBQUR3SUFJV3dFK29EUGxXckF4NW10cWhJQ0ZIZVZVZlhwMmQyb0lEQUVEeUJDMEFBQUM0RndFTG9JTjVNeTF1REdvbHIrcHozNW5RU2F3UHVUWXFBQUJTSjJnQkFBQkFKd0lXUUVleE11VENzRmJpTC9UYjcwN2cxeWFDRmdBQTlJR2dCUUFBQUQrVkxjcmpFRUo4YVByY3BJQmZVQm55TFpVaDBOM0VyQUFBNkFOQkN3QUFBTDZyRFZqRVgyRy9OQ0dnb3d1VkliZnlxbjRocEFiM2toa1hBQUI5Y0xCY0xsMFVBQUFBWHdsWUFBLzB2cGtXSHBLMjJzcVFHOXNzNE41K3Z6bzdGZGdDQUNCcE5sb0FBQUR3bVlBRjhFZ3FROVpkQ0ZuQWcwemF5aDBBQUVpV29BVUFBTURJWll2eXNBMVl2QnI3TElBSEs1dHAwUmpmU2w3VkU2RTFlRENiY1FBQVNKNmdCUUFBd0VpMUFZdno5cyt2cm9HSCt0Qk1pN25wcmJTVklYNk5EdzgzTVRzQUFGSW5hQUVBQURBeUFoYkFocWtNV1JkREowY3BIUWg2NXNTRkFRQ1F1b1BsY3VtU0FBQUFSa0RBQXRpQzE4MjBPRGZZbGJ5cVkrWEJQeW1jQlhydXo2dXowMnVYQ0FCQXFwNjRHUUFBZ09ITEZtVjhFSG9UUWlpRUxJQU4rZEJ1YitDV3loRFlqTXdjQVFCSW1lb1FBQUNBQWNzVzVjd2FlMkJMWnMyMCtHaTRLM2xWejFVZXdNWk1RZ2dYeGdrQVFLb0VMUUFBQUFaSXdBTFlzcithYVdHdGZ5dXY2dU4yWXhDd0dUWmFBQUNRTkVFTEFBQ0FBUkd3QUhiZ1V3amgzS0RYcUF5QnpUcktxL3J3NnV6VTFod0FBSklrYUFFQUFEQUEyYUo4MGE3WUZyQUF0azFseUIxNVZjZlF5Yk5rRGdUREVldEQzcmhQQUFCU0pHZ0JBQURRWTltaW5MUWJMRHprQTNiaGJUTXRQUGhzdFpVaDh5UU9BOE9UQ1ZvQUFKQXFRUXNBQUlBZUVyQUE5aUJXaHN3TWZrM2NKUFEwb2ZQQWtFemNKZ0FBcVJLMEFBQUE2QkVCQzJDUDVpcERidVZWSFN1Ym5xZHlIaGlnektVQ0FKQ3FnK1Z5NlhJQUFBQVNseTNLclAzbHRJQUZzQS92bW1uaDErV3R2S29QUXdnM3RsbkExdjF4ZFhiYUdETUFBS214MFFJQUFDQmgyYUk4YmpkWXZIUlB3SjZvRFBtV3loRFlqUmcwRmJRQUFDQTVnaFlBQUFBSkVyQUFFaElyUTI1Y3lFcGUxUlBmemJBejhmTjJhZHdBQUtSRzBBSUFBQ0FoQWhaQVltSmx5SVZMV1drclF6ejBoZDNKekJvQWdCUUpXZ0FBQUNSQXdBSkkxTG1MV1JPL3A0OFNPZzhNM1lrYkJnQWdSUWZMNWRMRkFBQUE3RW0yS0EvYkI1bm4rdjZCeEpUTnRKaTdsSlc4cXVNdjYvOUo0U3d3TW45ZW5aMWV1M1FBQUZKaW93VUFBTUFlQ0ZnQWlYc3ZaUEVObFNHd0g1TVFncUFGQUFCSkViUUFBQURZSVFFTG9DZG1MdXBXWHRWekZRYXdONW5SQXdDUUdrRUxBQUNBSFJDd0FIcmtkVE10R2hlMmtsZjFjUWloU09Fc01GS0NGZ0FBSk9lSkt3RUFBTml1YkZIR1g0YmZ0QS9xaEN5QWxIMElJYWdNV2FjeUJQYnJxQTA4QVFCQU1teTBBQUFBMkpJMllCRWZXQjZaTWRBVHMyWmFmSFJaSzNsVnh5MUV6MUk0QzR4YzFvWldBUUFnQ1lJV0FBQUFHeVpnQWZSVXJBeTVkbmtyN1Mvb2JmZUFOTVNneFJ0M0FRQkFLZ1F0QUFBQU5rVEFBdWd4bFNIZnVsVDNCTW1ZdUFvQUFGSWlhQUVBQVBCSTJhS00vK2YvUlFqaHhDeUJuanBYR1hJcnIrb1hLa01nS1Q2UEFBQWs1V0M1WExvUkFBQ0FCMmdERm5QLzV6L1FjMitiYWZIQ0phN2tWWDBZUXJpeHpRS1M4OGZWMldualdnQUFTSUdORmdBQUFQY2tZQUVNeUtjUXdzeUZybEVaQW1tSy8vdEwwQUlBZ0NRSVdnQUFBSFFrWUFFTTBFeGx5SzI4cXVQMy9QTlV6Z09zeVl3REFJQlVDRm9BQUFEOFFyWW9qOXRmT0F0WUFFUHlycGtXYjl6b1Nsc1pjcG5DV1lEdkVyUUFBQ0FaVDF3RkFBREE5OFdBUmJZbzQwTzMvd3BaQUFPak11UmJjV1BSVVdxSEFyNDZhUU5SQUFDd2R6WmFBQUFBL0V1N3dTSStjSHRwTnNCQXpadHBjZU55VjlyS2tGY3BuQVg0cWJqVjR0cUlBQURZTjBFTEFBQ0Fsb0FGTUJLeE11VENaYTh4RCtpSGlhQUZBQUFwRUxRQUFBQkdMMXVVaDIzQXdxK1pnYUZUR2ZJdmVWWEg3LytUcEE0Ri9FaG1NZ0FBcEVEUUFnQUFHSzAyWUhIZS9qMzFUZ0JHNEVKbHlLMjhxdU5EMnlLVjh3Qy9OREVpQUFCU2NMQmNMbDBFQUFBd0tnSVd3RWk5YjZhRlg0UGZrVmQxckNCNGxzeUJnQzUrdnpvN0ZSZ0RBR0N2YkxRQUFBQkdROEFDR0RtVklYZmtWWDB1WkFHOUZMZGFYTG82QUFEMlNkQUNBQUFZaFd4UnpnVXNnQkVybTJuUmVBT3M1RlY5SEVLWXAzQVc0TjVzNWdFQVlPOEVMUUFBZ0VITEZ1V3NmWmgyNUthQmtmclFUQXVoZ25XWGduZlFXNElXQUFEc25hQUZBQUF3U0FJV0FGK3BETGtqcitvWEtrT2cxM3grQVFEWXU0UGxjdWtXQUFDQXdSQ3dBRmp6dXBrVzUwYXlrbGYxWVFqaHhqWUw2TDAvcjg1T3IxMGpBQUQ3WXFNRkFBQXdDQUlXQU4vNDBINHZja3RsQ0F4RHJBOFJ0QUFBWUc4RUxRQUFnRjdMRnVXa2ZaQm9qVFRBdWxrekxUNmF5VXBlMWZHL0w1Nm5jQmJnMFRJakJBQmdud1F0QUFDQVhoS3dBUGlwdjVwcDRkZmVyYll5NURLSnd3Q2JNREZGQUFEMlNkQUNBQURvRlFFTGdGLzZGRUk0TjZZMXFxVmdXSTVpZ09ycTdOVFdIZ0FBOWtMUUFnQUE2QVVCQzRET1ZJYmMwVmFHdkVybVFNQ214UG9RbTNzQUFOZ0xRUXNBQUNCcDJhSThiZ01XTDkwVXdDKzliYWJGRzJOYWM1SFFXWURObVFoYThQL3MzVTl1M05iV0wrenQ0UFNsT3dMcGpNQThJMUNsdzY3OWdRQ2hudldPSUxvalVOVUlqaktDMSs0VkNCRFg2Ykx6U2lPNDVSRmNhUWJXQ1BSaEovU0o0L2lQL2xUVjV0NThIc0ROQU9SYUZaRlYvSEV0QUlCVUJDMEFBSUJKRXJBQWVMUzRNdVJNMmY3VWRFTzhqcnljeXZFQVc3VlFUZ0FBVWhHMEFBQUFKa1hBQXVESmxsYUcvS25waHJoVzRHSXF4d05zWGFXa0FBQ2s4dUwrL2w3eEFRQ0E1QVFzQUo3bGVuTjY0ZTN1enpUZHNESE5Bb3Izcjc2dE45b01BTUMrbVdnQkFBQWtWYTFYaHlHRWMyOGRBenlabFNGZmFMcmhYTWdDWmlGT3RSQzBBQUJnN3dRdEFBQ0FKRDRMV01SL0I3b0E4R1J4WmNpTjh2Mmg2WVpQRTVLQThsa2ZBZ0JBRW9JV0FBREFYZ2xZQUd4VlhCbHlxYVIvOGRiMUJXYkR5aVFBQUpJUXRBQUFBUFpDd0FKZ0o4NlY5VTlOTjhRVktpZFRPUjVnNTZ3SUFnQWdpWitVSFFBQTJMVnF2WW9QQXVOWSt3c2hDNEN0V1cxT0x6YksrWWVtRzJLZ3ozUVBtSm1tRzB5MUFBQmc3MHkwQUFBQWRxWmFyODdHUGZsSHFneXdWUjgycHhkTEpmMExLME5nbnFvUXdwWGVBd0N3VDRJV0FBREExZ2xZQU96Y21STC9xZW1HMXlHRVYxTTVIbUN2RnFiWkFBQ3diNElXQUFEQTFnaFlBT3pGcjFhRy9NbktFSmk5YXU0RkFBQmcvd1F0QUFDQVo2dldxOFU0c2wzQUFtQzNic2RBRzM4UzhJTjVPNHFCcTc2dFA4NjlFQUFBN0krZ0JRQUE4R1Jqd0NJKzREcFJSWUM5T051Y1huaVlPR3E2SVY2SGZwbkV3UUFweGI4RjczVUFBSUI5RWJRQUFBQWVUY0FDSUltNE11Uks2Zi9pN1lTT0JVaW5FclFBQUdDZkJDMEFBSUFIRTdBQVNNYktrQzgwM1dCbENQREpRaVVBQU5nblFRc0FBT0NIcXZVcXZpVjRLV0FCa015NWxTRi9hcm9oWHBjdXBuSThRSEx1VVFFQTJLdWZsQnNBQVBpV2FyMDZydGFyT0piOS8vb0JHeUNaM3phbkYwYmkvNVdWSWNCZmpBRXNBQURZQ3hNdEFBQ0F2NGtCaTNGRS9SdlZBVWpxTG9Sd3BnVi9hcnJoUElUd2NpckhBMHhHREZwc3RBTUFnSDBRdEFBQUFQNUR3QUpnY3M2c0RQbFQwdzJmcmxNQVgxcVlkZ01Bd0w0SVdnQUFBREZnY1JoQ3VCU3dBSmlVYXl0RC9pWStSRDJZMkRFQjAyQjFDQUFBZXlOb0FRQUFNellHTE03SGZ4NWNBVXlIbFNGZmFMb2gxdU5rVWdjRlRJbVZRZ0FBN00yTCsvdDcxUVlBZ0prUnNBQ1l2UCs5T2IyNDFLWS9OTjBRcjFzM3JsbkFEL3pjdC9XVklnRUFzR3NtV2dBQXdJd0lXQUJrNFZySTRtK3NEQUVlWWhGQ0VMUUFBR0RuQkMwQUFHQUdCQ3dBc21GbHlCZWFibmdkUW5nMXFZTUNwcXJTR1FBQTlrSFFBZ0FBQ2xldFYvR0IzVEtFY0tUWEFKTjN1VG05dU5HbVA0d3JRMHozQUI1SzBBSUFnTDE0Y1g5L3I5SUFBRkFnQVF1QTdIelluRjU0U1BpWnBodGl5T0tYeVJ3UWtJTi85bTB0c0FZQXdFNlphQUVBQUlVUnNBRElscFVobjJtNllTRmtBVHhCREt3SldnQUFzRk9DRmdBQVVBZ0JDNENzclRhbkZ4c3QvTU80TXVUdEZJNEZ5RTRNYWIzWE5nQUFka25RQWdBQU1sZXRWNHR4Zi8xTHZRVEkwdTNtOUdLcGRYOXhMamdJUEpFVlRBQUE3SnlnQlFBQVpHb01XTVFIY3lkNkNKQTFLME0rMDNSRGZFaDZNWmtEQW5MajNoZ0FnSjM3U1lrQkFDQXZNV0JSclZkWElZVC84VU15UVBaKzNaeGVYR25qWDFnWkFqekxHTmdDQUlDZE1kRUNBQUF5WVlJRlFIRnV4Ny9yakpwdVdGcUZCV3hCdkcvZUtDUUFBTHNpYUFFQUFCTlhyVmZISVlUTEVNSXJ2UUlveXRubTlPS2psdjZoNllaNHZUdWZ3ckVBMlRQUkFnQ0FuUkswQUFDQWlSb0RGdkhOM2pkNkJGQ2NkMWFHL0UxY0dYSXdzV01DOGlSb0FRREFUcjI0djc5WFlRQUFtQkFCQzREaTNZVVFqazJ6K0ZQVERYR1N4YituY2p4QUVmNVgzOWIremdJQXNCTW1XZ0FBd0VRSVdBRE13cTJWSVgvVmRNUGhlUDBEMktZNDFjTGtJQUFBZGtMUUFnQUFFcXZXcTA4UG1IN1JDNEJpeFlERmNuTjY4VmFMLzhiS0VHQVhGb0lXQUFEc2lxQUZBQUFrTWdZc3pzZC9IakFCbEVuQTRqdWFibmdkUW5nMTJRTUVjbGJwSGdBQXV5Sm9BUUFBZXlaZ0FUQUxBaFkvTUs0TVVSOWdWeFlxQ3dEQXJnaGFBQURBbmdoWUFNeUNnTVhETFYwUGdSMDZhTHJodUcvckcwVUdBR0RiQkMwQUFHQVBxdlhxM0FNbGdLSUpXRHhDMHczeFRmTmZzamxnSUZjTGszTUFBTmdGUVFzQUFOaWhhcjA2R3dNV1Irb01VQ1FCaTBleU1nVFlvMHF4QVFEWUJVRUxBQURZQVFFTGdPSUpXRHpkdWVzanNDZUNGZ0FBN0lTZ0JRQUFiSkdBQlVEeEJDeWVvZW1HK05Eekl0c1RBSEp6b21NQUFPeUNvQVVBQUd4QnRWNjlEaUZjQ2xnQUZFdkFZanZVRDlpcnBoc1dmVnRmcVRvQUFOc2thQUVBQU05UXJWZUxjWUtGdCtVQXlpUmdzU1ZOTjhTVklTK0xPQmtnSjNHU2pxQUZBQUJiSldnQkFBQlBJR0FCVUR3Qml5MXF1dUZ3dkc0QzdGdWw0Z0FBYkp1Z0JRQUFQSUtBQlVEeEJDeDJJNjdYT2lqeHhJREpXMmdSQUFEYjl1TCsvbDVSQVFEZ0J3UXNBSXIzSVlZQkJDeTJyK21HK0RiNS95M3R2SUNzL0srK3JUOXFHUUFBMjJLaUJRQUFmRWUxWGgyUEFZczM2Z1JRcE90eGdvWDkvYnR6V2VxSkFkbUlvZW4zMmdVQXdMWUlXZ0FBd0ZjSVdBQVVUOEJpRDVwdWVHMGFGREFCbGFBRkFBRGJKR2dCQUFDZkViQUFLSjZBeFg2WlpnRk13VUlYQUFEWUprRUxBQUFRc0FDWUF3R0xQV3U2SVY1WGoyWjEwc0JVVlRvREFNQTJ2YmkvdjFkUUFBQm1xMXF2RGtNSTUrTy9BNThFZ09JSVdDVFFkRU1NTUc1Y1c0RUorVmZmMWhzTkFRQmdHMHkwQUFCZ2xnUXNBSW9uWUpIVzB2VVZtSmhxRElBQkFNQ3pDVm9BQURBckFoWUF4Uk93U0t6cGhvVlZYTUFFV1I4Q0FNRFdDRm9BQURBTEFoWUF4Uk93bUk3bDNBc0FUTkpDV3dBQTJCWkJDd0FBaWxldFYyY2hoRXNCQzRBaUNWaE1TTk1OOFpwN012YzZBSlAwVWxzQUFOaVdGL2YzOTRvSkFFQ1J4b0JGZkt2MlNJY0JpaU5nTVRGTk44VHBVUnZYWFdEQ2Z1N2IyblVEQUlCbk05RUNBSURpQ0ZnQUZFM0FZcnJPWFh1QmlhdENDSzRmQUFBOG02QUZBQURGRUxBQUtKcUF4WVExM1hBY1FyaVlleDJBeVZ1TUt3VUJBT0JaQkMwQUFNaGV0VjU5K3NIVTNtV0E4Z2hZNU1HRFN5QUhsUzRCQUxBTkwrN3Y3eFVTQUlBc2pRR0xPTUhpUkFjQmlpTmdrWW1tRytMMStIL21YZ2NnRy8vczIvcEd1d0FBZUE0VExRQUF5STZBQlVEUkJDenlZNW9Ga0pNNDFVTFFBZ0NBWnhHMEFBQWdHd0lXQUVVVHNNaFEwdzFuVm5jQm1ZbEJpL2VhQmdEQWN3aGFBQUF3ZWRWNmRSeENlQ3RnQVZBa0FZdE1OZDF3YUpvRmtLR0ZwZ0VBOEZ5Q0ZnQUFUTllZc0lnVExON29Fa0J4QkN6eUY2L1JCM012QXBBZDRXMEFBSjd0eGYzOXZTb0NBREFwQWhZQVJST3dLRURURGZGYS9mL21YZ2NnVy8vcTIzcWpmUUFBUEpXSkZnQUFUSWFBQlVEUkJDeks4bmJ1QlFDeVZvVVFCQzBBQUhneVFRc0FBSktyMXF0UE85NEZMQURLSTJCUm1LWWJGa2J2QTVsYkNJd0JBUEFjZ2hZQUFDUXpCaXpPeDM5MnZBT1VSY0NpWEI1T0FybXJkQkFBZ09jUXRBQUFZTzhFTEFDS0ptQlJzS1liNG9xdm83blhBY2plU3kwRUFPQTVYdHpmM3lzZ0FBQjdJV0FCVURRQmk4STEzUkN2NHpldTRVQWhmdTdiMmpVTEFJQW5NZEVDQUlDOXFOYXJwWUFGUUpFRUxPYmowblVjS01naWhPRGFCUURBa3doYUFBQ3dVOVY2ZFJZZndCa3pEbEFjQVlzWmFicWhDaUc4bVhzZGdLSlUyZ2tBd0ZNSldnQUFzQk1DRmdERkVyQ1lwOHU1RndBb3prSkxBUUI0cWhmMzkvZUtCd0RBMWdoWUFCUkx3R0ttbW01NEhVTDRQM092QTFDa2YvWnRmYU8xQUFBOGxva1dBQUJzaFlBRlFMRitpOU1NQkN6bXFlbUdROU1zZ0lMRjlTR0NGZ0FBUEpxZ0JRQUF6MUt0VjRzeFlIR2lrZ0JGZVRkT3NQQUFhdDdPaFNpQmdzWHZNdTgxR0FDQXh4SzBBQURnU1FRc0FJb2xZTUh2bW00NEhvTVdBS1dxZEJZQWdLY1F0QUFBNEZFRUxBQ0tKV0RCbCtMMS9rQlZnSUw1VGdNQXdKTzh1TCsvVnprQUFINUl3QUtnV0FJVy9FM1REZkc2L3o4cUE4ekF2L3EyM21nMEFBQ1BZYUlGQUFEZlZhMVh4MlBBNG8xS0FSUkZ3SUx2V2FvT01CTXhXQ1pvQVFEQW93aGFBQUR3VlFJV0FNVVNzT0M3bW00NE04RUttSkZLc3dFQWVDeEJDd0FBL2tMQUFxQllBaGI4VU5NTmh5R0VTNVVDWmtUUUFnQ0FSeE8wQUFEZ2R3SVdBTVVTc09BeHprTUlCeW9Hek1qTEdETHIyL3FqcGdNQThGQ0NGZ0FBTTFldFY0ZmpRNVdMdWRjQ29EQUNGanhLMHczSDdnZUFtWXBUTGE0MEh3Q0FoeEswQUFDWXFjOENGdDVjQlNpTGdBVlBaV1VJTUZjTFFRc0FBQjVEMEFJQVlHWUVMQUNLSldEQmt6WGRFQjh5dmxKQllLWVdHZzhBd0dNSVdnQUF6SVNBQlVDeEJDellCdE1zZ0RtcmRCOEFnTWQ0Y1g5L3IyQUFBSVdyMXFzWXJsZ0tXQUFVUmNDQ3JXaTZJZDRuL0ZzMWdabjdaOS9XcnFrQUFEeUlpUllBQUFXcjFxdXpNV0J4cE04QXhSQ3dZR3VhYmpnYzd4VUE1aTZ1RDNrNzl5SUFBUEF3Z2hZQUFBVVNzQUFva29BRnUyRGlGY0FmckE4QkFPREJCQzBBQUFvaVlBRlFKQUVMZHFMcGh1TVF3aStxQy9BN1FRc0FBQjVNMEFJQW9BRFZldlZweksyQUJVQTVCQ3pZTlNQeUFmNTBvaFlBQUR6VWkvdjdlOFVDQU1qVUdMQlkrbEVRb0NnQ0Z1eGMwdzJ2UXdqL1I2VUIvdUxudnEydmxBUUFnQjh4MFFJQUlFTUNGZ0JGRXJCZ255NVZHK0J2NHZvUVFRc0FBSDVJMEFJQUlDTUNGZ0JGRXJCZ3I1cHVXRm8zQnZCVmxiSUFBUEFRZ2hZQUFCbW8xcXRxZlBOVXdBS2dIQUlXN0YzVERZY2hoSE9WQi9pcWhiSUFBUEFRZ2hZQUFCTldyVmZINHdTTE4vb0VVQXdCQzFLS3djMERIUUQ0cXFNWVNPdmIrcVB5QUFEd1BZSVdBQUFUSkdBQlVDUUJDNUpxdXFGeWJ3SHdRM0dxeFh0bEFnRGdld1F0QUFBbVJNQUNvRWdDRmt6RnBVNEEvRkFsYUFFQXdJOElXZ0FBVEVDMVhoMk9EejhFTEFES0lXREJaRFRkY0JaQ09ORVJnQjlhS0JFQUFEOGlhQUVBa05BWXNEZ2YvOW1YRGxBR0FRc21wZW1HdzNGaUZnQS9WcWtSQUFBL0ltZ0JBSkNBZ0FWQWtRUXNtS3A0djNHa093QVBjdEIwUTlXMzlVYTVBQUQ0RmtFTEFJQTlFckFBS0pLQUJaUFZkTVB4ZU44QndNUEZxUmFDRmdBQWZKT2dCUURBSGdoWUFCVG5Mb1R3Tm9Sd0tXREJ4QzNkZXdBOG12VWhBQUI4bDZBRkFNQ09WZXZWMmZpUXc4aHVnUHpGZ01YbEdMRDRxSjlNV2RNTml4RENHMDBDZUxTRmtnRUE4RDJDRmdBQU95SmdBVkFVQVF0eWRLbHJBRS95VXRrQUFQaWVGL2YzOXdvRUFMQkZBaFlBUlJHd0lFdE5OOFQ3a2YvV1BZQW4rN2x2Nnl2bEF3RGdhMHkwQUFEWUVnRUxnS0lJV0pDdHBoc09UYk1BZUxhNFBrVFFBZ0NBcnhLMEFBQjRwbXE5V293UE00eVhCY2lmZ0FVbE9BOGhIT2drd0xOVXlnY0F3TGNJV2dBQVBORVlzSWdUTEU3VUVDQjdBaFlVb2VtRzR4RENoVzRDUEp1Z0JRQUEzeVJvQVFEd1NBSVdBRVVSc0tBMGIzVVVZQ3VPWW5pdGIrc2I1UVFBNEV1Q0ZnQUFEeVJnQVZBVUFRdUswM1REd24wS3dGYkZxUmFDRmdBQS9JMmdCUURBRDFUcjFmSDRNTzZWV2dGa1Q4Q0NrcGxtQWJCZE1XanhYazBCQVBpU29BVUF3RGVNQVlzNHdlS05HZ0ZrVDhDQ29qWGRjQjdIM09zeXdGWXRsQk1BZ0s4UnRBQUErSUtBQlVCUkJDd29YdE1OaCtPOUN3RGJaUjBUQUFCZkpXZ0JBREFTc0FBb2lvQUZjeEx2WHc1MEhHRDdtbTZvK3JiZUtDMEFBSjhUdEFBQVpxOWFyejY5QmZyTDNHc0JVQUFCQzJZbFBnQjBEd093VS9IdnJLQUZBQUIvSVdnQkFNeldHTEE0SC85NUN4UWdid0lXek5XbHpsTzQ2L0ZCdC90MVVsbUVFTjZxUGdBQW54TzBBQUJtUjhBQ29DZ0NGc3hXMHcydlF3Z25QZ0VVN256OE8rK3pUaXFWeWdNQThDVkJDd0JnTmdRc0FJb2lZQUdtV1ZDK1gvdTIzalRkY0NWb1FVSXZtMjQ0N052YS9RWUFBUDhoYUFFQXpFSzFYc1Z3eFZMQUFpQjdBaGJ3eHpTTGVGOXpwQllVN0c2OGY0OWkwT0pDczBtb0dqK0hBQUR3TzBFTEFLQm8xWHAxTnY1QTYwRUVRTjRFTEdBVTM2d2VKM1JCeVphZlRSRFk2RFNKTFFRdEFBRDRuS0FGQUZBa0FRdUFZZ2hZd045ZG10SkY0Vzc3dHY3UGFwd1l1R2k2NFVOYzRhRHhKRklwUEFBQW54TzBBQUNLSW1BQlVBd0JDL2lLcGh2aVc5VnYxSWJDblgzbDlEYUNGaVMwVUh3QUFENG5hQUVBRktGYXIxNlBEK1FFTEFEeUptQUIzN2RVSHdyM1c5L1dYMXZSc0JFeUlxR0RwaHVPKzdhKzBRUUFBSUtnQlFDUXUycTlXb3dQSEU0MEV5QnJBaGJ3QTAwM25Mbm5ZUWJPdjNHS1h3dGZ3RDdGOVNHQ0ZnQUEvRTdRQWdESWtvQUZRREVFTE9BQm1tNDROTTJDR1ZoOWEySkEzOWFicGh0OEJrZ3BmZ2Q5cndNQUFBUkJDd0FnTndJV0FNVVFzSURIT2JjaWpjSjl1aTU4ejdYdkFTUlVLVDRBQUo4SVdnQUFXUkN3QUNpR2dBVThVdE1OeDk5WnB3Q2xPTy9iK2tmWGhTdmZCMGpJWnc4QWdQOFF0QUFBSnExYXI0N0hnTVViblFMSW1vQUZQRjM4ZitkQS9TalloNzZ0M3o3ZzlEWStCS1RVZEVNVjE5aG9BZ0FBZ2hZQXdDUUpXQUFVUThBQ25xSHBoampWNjVVYVVyaUhUbXp4Z0p2VUZqNkhBQUFFUVFzQVlHb0VMQUNLSVdBQjIzR3BqaFR1WGQvV1Z3ODV4YjZ0YjVwdXVBMGhIUGxRa0VpbDhBQUFCRUVMQUdBcUJDd0FpbkU3L2oxL0wyQUJ6OU4wdzFrSTRhVXlVckM3OFpyeEdCdEJDeEphS0Q0QUFFSFFBZ0JJclZxdkRzZFJ3ZWQyandOazdmZUF4ZWIwNGlFNzlvRWZhTHJoMERRTFp1QXlUcWw0NUdsZVdhZERRa2Z4NzNQZjFzS2tBQUF6SjJnQkFDUWhZQUZRREFFTDJJMmxleVFLZC92RU1OSEdCNFBFcWpId0F3REFqQWxhQUFCN0pXQUJVQXdCQzlpUnBodmlTclZmMUpmQ0xaOHlGYUJ2NjZ1bUczdzJTR2toYUFFQWdLQUZBTEFYQWhZQXhSQ3dnTjN6L3hlbHUrN2Iram1mOHc4aGhKYytKU1N5VUhnQUFBUXRBSUNkcTlhcnMzRXNzSUFGUUw0RUxHQVBtbTZJRC9CTzFKckNuVC96OUs0RUxVaW9VbndBQUFRdEFJQ2RHUU1XY2IvNGtTb0RaRXZBQXZiTC8ydVU3bDNmMXB0bm51TnovM3Q0am9PNDRxbHY2eHRWQkFDWUwwRUxBR0RyQkN3QWlpQmdBWHZXZE1PNSt5Y0tkN2VGYVJaaG5HZ0JLUzBFNHdBQTVrM1FBZ0RZR2dFTGdDSUlXRUFDVFRjY2p2ZFJVTExMdnEwL1B2Zjg0aVNCcGh2dXJDWWtJZXREQUFCbVR0QUNBSGkyYXIyS2IvTmMycE1Na0RVQkMwanIwa05qQ25mYnQvVTJ3MFJ4ZmNpSkR3MkpDRm9BQU15Y29BVUE4R1Jqd0dMcEIwNkFyQWxZUUdKTk44UUhkbS8wZ2NKdFkyWEk1NjU4RHlFaG56MEFnSmtUdEFBQUhrM0FBcUFJQWhZd0haZDZRZUd1KzdaK3YrVlRqRUdMQ3g4Y1VtbTZZZEczOVpVR0FBRE1rNkFGQVBCZ0FoWUFSUkN3Z0FscHV1RzFleXRtNEd3SHA3anh3U0d4YWd6OEFBQXdRNElXQU1BUFZldlZjUWpoclljQUFGa1RzSUJwTXMyQzB2M2F0L1hOdHMreGIrdVBUVGQ4Q0NHODlBa2lrWVcvNFFBQTh5Vm9BUUI4MHhpd1dOb1pEcEExQVF1WXFLWWI0bjNXa2Y1UXNMdngrOFN1YkFRdFNLaFNmQUNBK1JLMEFBRCtSc0FDb0FnQ0ZqQmhUVGZFKzYxelBhSnd5emg1WW9lbmVPVTdDd2tkTmQxd2Nxc2toUUFBSUFCSlJFRlV1T1BQT0FBQUV5Vm9BUUQ4aDRBRlFCRUVMQ0FQOFo3clFLOG8ySWUrclhlOVZtSGpBMFJpY1gzSWUwMEFBSmdmUVFzQUlBWXNEc2Zkc2dJV0FQa1NzSUJNTk4yd2NOL0ZET3g4WWt2ZjFwdW1HM3lXU0trU3RBQUFtQ2RCQ3dDWXNURmdjVDcrODBZbFFKNEVMQ0EvU3oyamNMLzFiWDIxcDFPOERpR2MrRUNSeUVMaEFRRG1TZEFDQUdaSXdBS2dDQUlXa0tHbUc4NDhGR1lHZGo3TjRqTlgvcDhpb1VyeEFRRG1TZEFDQUdaRXdBS2dDQUlXa0ttbUd3NU5zMkFHVm4xYjMrenhORGMrVkNSMDBIUkRGZGZZYUFJQXdMd0lXZ0RBVEZUcjFWTEFBaUJyQWhhUXYzZ3ZkcVNQRk93dWhIQzU1OVBiMTRvUytKWks0QWNBWUg1ZTNOL2ZhenNBRkt4YXI4N0dOeWY5cUErUUp3RUxLRURURGNjaGhQK25seFR1di9xMjN2djFxdW1HRzk5M1NPaGQzOVpuR2dBQU1DOG1XZ0JBb1FRc0FMSW5ZQUZsMmZkYi9yQnYxeWxDRnFPTjd6MGtWQ2srQU1EOENGb0FRR0VFTEFDeUoyQUJoV202WVJGQ2VLV3ZGRzZaOFBTdS9EOUdRaThWSHdCZ2ZnUXRBS0FRQWhZQTJST3dnSEtaWmtIcDR1cUVxNFRudVBFSkk2VVlxRXY4L3dBQUFIc21hQUVBbWF2V3E4VVlzRGpSUzRBc0NWaEF3WnB1T1BlMk00VzdTenpOSXNRSDNFMDMrSnlSMG1LY3JBSUF3RXdJV2dCQXBnUXNBTEluWUFHRmE3cmhNUFVEYU5pRHk3NnRieVpRNkd2ZmpVaW9VbndBZ0hrUnRBQ0F6QWhZQUdSUHdBTG1JOTZ6SGVnM0JidWQwR3FjamU5SUpDUm9BUUF3TTRJV0FKQUpBUXVBN0YyUEFRdGpwV0VHbW00NERpSDhvdGNVN3J4djY0OFRPY1hOQkk2QitUcUtmL2NuTXQwRkFJQTlFTFFBZ0ltcjFxdmpNV0R4UnE4QXNpUmdBZk5rYWcybHUrN2IrdjJFenRGMWx0VGlWQXRCQ3dDQW1SQzBBSUNKRXJBQXlKNkFCY3hVMHcwTFU4aVlnZk1wbldLY0pOQjB3NTExUFNRVWd4WlRDaDhCQUxCRGdoWUFNREVDRmdEWkU3QUFUTE9nZE8vNnRwN2lxbzU0N1gwMWdlTmduaGI2RGdBd0g0SVdBREFSQWhZQTJST3dBT0kwaTNnL2Q2UVNGT3h1YXRNc1ByTVJ0Q0FoazR3QUFHWkUwQUlBRXF2V3E4UHhoOG9MdlFESWtvQUY4THVtR3c0bi9BQWF0dVd5Yit1UEU2M21sZTlWcE5SMFF6WFJhUzhBQUd5Wm9BVUFKUEpad09MY0htR0FMQWxZQUYrNmRGOUg0Vzc3dGw1TzlSVDd0cjVxdW1FQ1I4S01MY2JKS2dBQUZFN1FBZ0QyVE1BQ0lIc0NGc0RmeExlWXJZQmpCczR5T01VUElZU1hFemdPNXFuU2R3Q0FlUkMwQUlBOUViQUF5SjZBQmZBOWw2cEQ0YTdqeElnTVRuRWphRUZDZ2hZQUFEUHhrMFlEd081VjYxVU1WOXlNKzRLRkxBRHlFZ01XUDI5T0x4WkNGc0RYTk4zd09vUndvamdVTG9kcEZwRnJOU205YkxyaFVBY0FBTXBub2dVQTdGQzFYc1VmSStNTzR5TjFCc2lPQ1JiQUQ0MFAxRXl6b0hTLzltMTlrOGs1Ymlad0RNeGJKZkFEQUZBK1FRc0EyQUVCQzRDc0NWZ0FqM0h1bm8vQzNZM2ZiYkxRdC9XbTZZWTdrd1JKYUNGb0FRQlFQa0VMQU5naUFRdUFyQWxZQUkvU2RNUHhHTFNBa3AzM2JmMHhzL1BiV09kRFFwWGlBd0NVVDlBQ0FMYWdXcS9pR3l0dkJTd0FzaVJnQVR6VjBsdnpGTzVEMzladk16ekZLMEVMRWxvb1BnQkErUVF0QU9BWnhvREYwbzk0QUZrU3NBQ2VyT21HZUIvNFJnVXBYSzRUV3pZVE9BYm02eUJPUE9yYitzWm5BQUNnWElJV0FQQUVBaFlBV1JPd0FMWmhxWW9VN3JlK3JYTzlWcnJHazFwY0h5Sm9BUUJRTUVFTEFIZ0VBUXVBckFsWUFGdlJkTU9aKzBGbUlOZHBGcUZ2NjQ5Tk45eGE3VWhDOGJlRDl4b0FBRkF1UVFzQWVJQnF2WXB2bzF6NlFSMGdTd0lXd05ZMDNYQTQzaGRDeVZZRnJEMjRzdDZIaENyRkJ3QW9tNkFGQUh4SHRWNGRqeE1zL0VBSGtCOEJDMkFYNGx2K0J5cEx3VzRMQ1JOdGZJOGpJUzlwQUFBVVR0QUNBTDVDd0FJZ2F3SVd3RTQwM1JEdkVTOVVsOEl0NCtxTkFrNXhNNEZqWU1hYWJsajBiZTErRkFDZ1VJSVdBUEFaQVF1QXJBbFlBTHRtWlFpbHUrN2IrbTBKNXhnZmNEZmRNSUVqWWNhcWNZVU5BQUFGRXJRQWdEOENGcDkyYlF0WUFPUkh3QUxZdWZobWNnamhsVXBUdUdWaHAzZHRoUU1KVllvUEFGQXVRUXNBWm0wTVdKemJ0UTJRSlFFTFlKK0tlTXNmdnVOZGdXc09Ob0lXSkxSUWZBQ0FjZ2xhQURCTEFoWUFXUk93QVBhcTZZWjR6M2lrNmhUc3JzQnBGbUZjMi9ETEJJNkRlVHBxdXVHd2IrdVArZzhBVUI1QkN3Qm1SY0FDSUdzQ0ZzRGV4WWRraFQ2QWhzOWQ5bTE5VTJCRk5oTTRCdWF0R2dNL0FBQVVSdEFDZ0ZrUXNBREltb0FGa05MUy9TT0Z1KzNidXNnd1VReVBOTjF3NS85aEVsb0lXZ0FBbEVuUUFvRGlWZXZWMmZnRHVYSFBBSGtSc0FDU2FycmgyTm9CWnVDODhGT005eEd2Sm5BY3pOTkMzd0VBeWlSb0FVQ3hCQ3dBc2lWZ0FVekZXNTJnY05kOVc3OHYvQnczZ2hZa1ZDaytBRUNaQkMwQUtJNkFCVUMyZm9zNzRnVXNnQ2xvdXVGMUNPRkVNeWhjNmRNc3dqalI0bUlDeDhFOEhjVHBTSEdOamY0REFKUkYwQUtBWWdoWUFHVHIzVGpCd2cvUXdKUmM2Z2FGZTllMzlhYjBrK3piK3FycGhna2NDVE8yTUNFSkFLQThnaFlBWks5YXJ4YmpEK0V2ZFJNZ0t3SVd3Q1ExM1NDOFMrbnVaakxONHBNUHZpK1NrUFVoQUFBRkVyUUFJRnRqd0dKcHBETkFkZ1FzZ01scXV1RndaZytnbWFkbDM5WWZaM1RtVjRJV0pMUlFmQUNBOGdoYUFKQWRBUXVBYkFsWUFEbUlrOUlPZElxQzNmWnRQYmZWT01XdlNHSFNoSHdBQUFva2FBRkFOZ1FzQUxJbFlBRmtvZW1HT043OWpXNVJ1TE1aTmxqUWdxU2FibGowYlgybEN3QUE1UkMwQUdEeXF2WHFlSHl6OEpWdUFXUkZ3QUxJemR6ZThtZCtydWY0c0xkdjYwM1REWGVtMVpCUU5hNndBUUNnRUlJV0FFeldHTEJZZXFzUUlEc0NGa0IybW00NE16bU5HWmpqTkl0UE52NGZKNkdGTUI4QVFGa0VMUUNZSEFFTGdHd0pXQUJaYXJyaGNMei9oSkw5MnJmMW5LL1JWNElXSkZRcFBnQkFXUVF0QUpnTUFRdUFiQWxZQUxrN0R5RWM2U0lGdXhNbStqMW9jVEdCNDJDZWptS29yMi9yai9vUEFGQUdRUXNBa3F2V3EwOXZFUDZpR3dCWkViQUFzdGQwdy9FWXRJQ1NuWHZBKy92cUVFZ3ByZzk1cndNQUFHVVF0QUFnbVRGZ2NUNytPOUFKZ0d3SVdBQWx1WFF2U3VFKzlHMzlkdTVOamtHVHBodHVUYThob1VyUUFnQ2dISUlXQU95ZGdBVkF0Z1FzZ0tJMDNSRGZMbjZscXhUT3hKWS9YVmxWU1VJTHhRY0FLSWVnQlFCN0kyQUJrQzBCQzZCVWx6cEw0WDdyMi9wS2svOWpJMmhCUXBYaUF3Q1VROUFDZ0wybzFxdWxnQVZBZGdRc2dHSTEzWEFXUW5pcHd4VHN6alNMdnhFNklhV0RwaHVxdnEwM3VnQUFrRDlCQ3dCMnFscXY0Zy9ZUzN0d0FiSWlZQUVVcmVtR1E5TXNtSUhMdnExZHl6OFRIM0EzM1RDWjQyR1dxbkd5Q2dBQW1STzBBR0FuQkN3QXNpUmdBY3lGU1d1VTdsYVk2SnV1UXdnbkV6MDJ5cmNJSWJ6Vlp3Q0EvQWxhQUxCVkFoWUFXUkt3QUdhajZZYmpFTUtGamxPNFpkL1dIelg1cXphQ0ZpUlVLVDRBUUJrRUxRRFlpbXE5ZWoyK01TVmdBWkFQQVF0Z2pyeEpUT211KzdiMk9mKzJxeERDTDFNOU9JcjNVb3NCQU1vZ2FBSEFzMVRyMVdLY1lPR05JSUI4Q0ZnQXM5UjB3OEo5S3pPdzFPVHYya3o0MkppQmVDM3EyL3BLcndFQThpWm9BY0NUQ0ZnQVpFbkFBcGc3Yi9sVHVuY2U0SDVmMzlZM1RUZmNtc1pJUW90eHNnb0FBQmtUdEFEZ1VRUXNBTElrWUFITVh0TU41eDZzVXJpN0VNSzVKai9JeHQ4REVxb1VId0FnZjRJV0FEeUlnQVZBbGdRc0FQNElXUnhhcDhBTVhQWnQvVkdqSHlRR0xWNWxjSnlVU2RBQ0FLQUFnaFlBZkZlMVhoMlBQMHEvVVNtQWJBaFlBUHhWdko4OVVCTUtkdHUzdFREUnc4VzFEUmU1SEN6Rk9XcTY0VGl1c2RGYUFJQjhDVm9BOEZVQ0ZnQlpFckFBK0VMVERmSE40Vi9VaGNKWkdmSUlmVnRmTmQyUXpmRlNwSGh0Y3M4T0FKQXhRUXNBL2tMQUFpQkxBaFlBMzNhcE5oVHV1bS9yOTVyOGFCOUNDQzh6TzJiS0VkZXordjhXQUNCamdoWUEvRTdBQWlCTEFoWUEzOUYwdytzUXdva2FVVGpUTEo3bVN0Q0NoQ3JGQndESW02QUZ3TXhWNjlYaCtNUGN1YjNWQU5rUXNBQjRHTk1zS04ydmZWdHZkUGxKMUkyVWhBQUJBREluYUFFd1V3SVdBRmtTc0FCNG9LWWI0clMySS9XaVlIZmpWRUtlNWtyZFNLbnBoa3BRQ2dBZ1g0SVdBRE1qWUFHUW5idHhmN09BQmNBRE5kMXdhSjBDTTdEczIvcWpSajlOMzlZM1RUZmMrVjVNUWd1VFZRQUE4aVZvQVRBVEFoWUEyYmtiUjk1ZmJrNHZQRVFCZUp4TDk3d1U3clp2YTZ0eG5tOWpoUU1KVllvUEFKQXZRUXVBR2FqV3F6TS9OZ05rUThBQzRCbWFib2h2Q0w5UlF3cDNwc0ZiY1NWb1FVS0NGZ0FBR1JPMEFDallHTEN3bXhvZ0R3SVdBTnV4VkVjS2Q5MjM5WlVtYjBXczQwVUI1MEdlWHNaVlYxWUFBUURrU2RBQ29FQUNGZ0JaRWJBQTJKS21HODY4bmM0TW1HYXhQWnRTVG9Sc1ZXUGdCd0NBekFoYUFCUkV3QUlnS3dJV0FGc1Uzd28yellJWldQVnRmYVBSMnhFbkNUVGQ4Q0ZPRmlqaGZNalNRdEFDQUNCUGdoWUFCYWpXcS9qRi9LMkFCVUFXQkN3QWR1UGMvVENGKzNRUHdYWnRCQzFJcUZKOEFJQThDVm9BWkd3TVdDeU5Sd2JJZ29BRndJNDAzWEFjUXJoUVh3cDNIaWN3YVBMV3hhREZtOExPaVh3czlBb0FJRStDRmdBWkVyQUF5SXFBQmNEdWVjdWYwbjNvMi9xdEx1K0V0UTJrZEJERGdsWUNBUURrUjlBQ0lDTUNGZ0JaRWJBQTJJT21HK0k5OGl1MXBuRG5HcndiZlZ0dm1tNG84ZFRJeDZkMXNBQUFaRVRRQWlBRDFYcDFQSDdwRnJBQW1ENEJDNEQ5TXMyQzByM3IyOXJVaGQyNjluMmJoQ3JGQndESWo2QUZ3SVNOQVl1bGZiRUFXUkN3QU5penBodk9RZ2d2MVoyQzNZM2ZDZG10SzBFTEVoSzBBQURJa0tBRndBUUpXQUJrUmNBQ0lJR21HdzVOczJBR0x2dTJ2dEhvbmRzVWZuNU1tNUFQQUVDR0JDMEFKa1RBQWlBckFoWUFhY1g3NWdNOW9HQzN3a1I3STJoQlVrMDNMS3dJQWdESWk2QUZ3QVJVNjlXbnQvRUVMQUNtVDhBQ0lMR21HMkpBK1JkOW9IREx2cTNkYSt4Qm5CclNkRU1NdGh3VmY3Sk1WVFd1c0FFQUlCT0NGZ0FKalFHTDgvR2Z0L0VBcGszQUFtQTYzdW9GaGJ2dTI5cm5mTDgyZ2hZa1ZDaytBRUJlQkMwQUVoblhoR3dFTEFBbVQ4QUNZRUxpZUhYNzdKbUJjMDNldXpoTjROWE16cG5wV09nRkFFQmVmdEl2Z0RRMnB4YzNJWVFiNVFlWXJCaXdXSVVRamplbkYwc2hDNERKOEpZL3BYdlh0L1ZHbC9kT3pVbnBxT21HUXgwQUFNaUhvQVZBV240a0JwZ2VBUXVBaVdxNjRkeG9md3AzWjVwRkduMWJYODN4dkprVTYwTUFBREppZFFoQVd1OURDUC9XQTRCSnNDSUVZTUxHTjMyWGVrVGhMdnUyZGgrU3pyWFZSQ1MwR0ZmWUFBQ1FBUk10QUJJYTE0ZDgwQU9BcEV5d0FNaERETU1kNkJVRnUrM2JXcGdvTGV0RFNHbWgrZ0FBK1REUkFpQzl0NlphQUNSaGdnVkFKcHB1aU9QVTMrZ1hoYk15SkQxQkMxS3lPZ1FBSUNNbVdnQ2s5MTRQQVBiS0JBdUEvRnpxR1lXNzd0dmFkOFAwckcwZ3BZTXhXQWdBUUFZRUxRQVNzejRFWUc4RUxBQXkxSFRENnhEQ2lkNVJ1RE1OVHE5djY1dnhuaEZTRWJRQUFNaUUxU0VBMDJCOUNNRHVXQkVDa0RmVExDamRyK01EZnFZaFRyVjRwUmNrSW1nQkFKQUpRUXVBYVhndmFBR3dkUUlXQUpscnVtRVpRampTUndvVzcxZVdHandwRzBFTEVsb29QZ0JBSHF3T0FaaUFjWDNJdFY0QWJJVVZJUUFGYUxyaE9JUndycGNVYnRtM3RYdVZhYm1hZXdGSTZxWHlBd0RrUWRBQ1lEcmU2Z1hBc3doWUFKUWx2dVYvb0tjVTdMWnZhNnR4cG1jejl3S1FWdE1OcGxvQUFHUkEwQUpnT3Q3ckJjQ1RDRmdBRkdaOHlQUkdYeW5jbVFaUHp6aGg1TVBjNjBCU2xmSURBRXlmb0FYQVJJd1BCbi9URDRBSHV3MGgvRzhCQzRBaUxiV1Z3djNXdDdVVkZkTmxxZ1VwbVdnQkFKQ0JmMmdTd0tURXFSYXZ0QVRndTJMQUlnWXJyRndDS0ZEVERmRXQveE85cFhEbkdqeHBWNmJxa0pDSkZnQUFHVERSQW1CYXJBOEIrTFlZc1BpdnplbkZzWkFGUUptYWJqZ01JVnhxTDRWYjlXMTlvOG1UWnFJRktSMk4xME1BQUNaTTBBSmdRcXdQQWZncUFRdUErWWh2K1Ivb053VzdFeWFhdnI2dEJTMUl6Zm9RQUlDSnN6b0VZSHFzRHdINGd4VWhBRFBTZE1OeENPRkN6eW5jZWQvV0h6VTVDOWZXR0pGUVplb3BBTUMwQ1ZvQVRFLzhJdjNmK2dMTW1JQUZ3RHg1eTUvU1hmZHQ3ZjRtSDFlQ0ZpUmtvZ1VBd01SWkhRSXdNZGFIQURObVJRakFURFhkc0REVmpSbFlhbkpXckE4aEpTRWZBSUNKRTdRQW1DWVBHSUU1RWJBQXdEUUxTdmV1YitzclhjNktmcEZVMHcyVkRnQUFUSmVnQmNBRWJVNHY0dnFRTzcwQkNpZGdBVUI4a0hRZVFuaXBFaFRzempTTC9QUnQvWEc4WDRWVUJDMEFBQ2JzSDVvRE1Ga3hiUEZHZTRBQ3hSK3NsOElWQURUZGNPZ0JORE53MmJmMWpVWm5LYTRQT1pwN0VVaG1ZZUlwQU1CMENWb0FUSmVnQlZBYUFRc0F2aFJERmdlcVFzRnVyY2JKV2x3ZjhtcnVSU0FaRXkwQUFDYk02aENBaWJJK0JDaUlGU0VBL0UzVERjY2hoRjlVaHNLZGp5c295Tk5HMzBqSVdpMEFnQWtUdEFDWXR2ZjZBMlJNd0FLQTczRnRvSFRYZlZ2N1RwZXh2cTJ2NWw0RDBtcTZZYUVGQUFEVFpIVUl3TFJaSHdMa3lJb1FBTDVyZkhCMG9rb1U3bHlEaTNEdDd4VUpMY1lWTmdBQVRJeUpGZ0FUWm4wSWtCa1RMQUI0S05jSlN2ZXViMnRySjhxZ2o2UlVxVDRBd0RTWmFBRXdmYVphQUZObmdnVUFEOVowd3pLRWNLUmlGT3pPTkl1aXhHa0N2OHk5Q0NRamFBRUFNRkVtV2dCTW53ZVh3RlNaWUFIQW96VGRjT2dCTkROdzJiZjFSNDB1aG9rV3BIVFVkTU94RGdBQVRJK2dCY0RFYlU0dnJzYUhtUUJUSVdBQndGTmRoaEFPVkkrQzNmWnR2ZFRnY3ZSdGZXT2xKNG1aYWdFQU1FR0NGZ0I1ZUs5UHdBUUlXQUR3WkUwM1ZGYmlNUU5ubWx5a3E3a1hnS1FXeWc4QU1EMy8wQk9BTEx5MUV4WklLQVlzbHNJVkFEelRwUUpTdU91K3JUMlFMMU5jSC9KcTdrVWdHUk10QUFBbXlFUUxnQXhzVGk4MjFvY0FDWmhnQWNCV05OM3dPb1J3b3BvVXpqU0xjZ25Ra0pMckp3REFCQWxhQU9URCtoQmdYd1FzQU5pYXBoc09UYk5nQm43dDIvcEdvOHRrVWdtcGpldTNBQUNZRUt0REFQSmhmUWl3YTFhRUFMQUw1eUdFSTVXbFlIZnhIa3FEaS9jaGhQQnk3a1VnbWNXNHdnWUFnSWt3MFFJZ0U5YUhBRHRrZ2dVQU85RjB3L0VZdElDU0xmdTIvcWpEeGZPUW01Uk10QUFBbUJnVExRRHk4dDVVQzJDTFRMQUFZTmZpVy80SHFrekJQdlJ0YlRYT1BNVDFJVy9tWGdTU0ViUUFBSmdZRXkwQTh1SUhQR0FiVExBQVlPZWFibGg0S01rTW1OZ3lIeVpha05MTHBoc09kUUFBWURvRUxRQXlzam05dUJuM3dnSThoWUFGQVBza0pFenBmdXZiK2txWDU2RnY2eGkwdUp0N0hVaktWQXNBZ0FteE9nUWdQL0hoNkwvMURYZ0VLMElBMkt1bUc4N2kyN2VxVHVGTXM1aWZHTFk0bVhzUlNHWXhyckFCQUdBQ0JDMEE4dk5lMEFKNElBRUxBUFp1SEcxdW1nV2xXL1Z0ZmFQTHMzTWxhRUZDQzhVSEFKZ09xME1BTW1OOUNQQUExeUdFLzgrS0VBQVNpVy81SHlnK0Jic1RKcHF0emR3TFFGSldod0FBVElpZ0JVQ2VQRGdGdmlZR0xIN2VuRjRzTnFjWDcxVUlnSDFydXVFNGhIQ2g4QlR1dkcvcmo1bzhTOVkya05MQmVKMEZBR0FDQkMwQTh1UUJLdkM1endNV2Z2d0ZJQ1Z2K1ZPNjY3NnRCZDluYWd6WTNNNjlEaVJsZlFnQXdFUUlXZ0JreVBvUVlDUmdBY0JrTk4wUUgvNjgwaEVLdDlUZzJYUGZUVXJXaHdBQVRNUS9OQUlnVy9FdHFuOXJIOHhTREZnc2hTc0FtQmh2K1ZPNmQzMWJ1LzlpRTBKNE0vc3FrSXFnQlFEQVJBaGFBT1RydmFBRnpJNkFCUUNUMUhURGVRamhTSGNvMkoxcEZvemNpNVBTaWVvREFFekRpL3Y3ZTYwQXlGUzFYcjAzbmhsbVFjQUNnTWxxdXVFd2hCQlgyeDNvRWdWYjlXMHRhTUh2bW03d2d5b3AvV3k2RGdCQWVpWmFBT1JOMEFMS0ptQUJRQTZXUWhZVTdsYklnaTljbXl4QVFwWEpLZ0FBNlFsYUFPUXRCaTMrV3craE9BSVdBR1NoNlliakVNSXZ1a1hoempXWUwyd0VMVWlvVW53QWdQUiswZ09BZkcxT0x6NkdFSDdUUWloR0RGajh2RG05V0FoWkFKQ0p0eHBGNGE3N3RuNnZ5WHpCdlRvcExWUWZBQ0E5UVF1QS9QblJEL0luWUFGQWRwcHVlTzJOYm1iQU5BdStacU1xSkhUVWRNT2hCZ0FBcENWb0FaQS9RUXZJbDRBRkFEbTcxRDBLOTY1dmF3L1UrWnUrclc5Q0NMY3FRMEttV2dBQUpDWm9BWkE1NjBNZ1N3SVdBR1N0NllabGZLTldGeW5ZbldrVy9JQVFEaWxWcWc4QWtOWS8xQitnQ0hHcXhTdXRoTW1MQVl1bGNBVUFPUnZIbFhzQVRlbVdmVnQvMUdXK1krTjdPQW1aYUFFQWtKaWdCVUFaWXREaXYvVVNKa3ZBQW9DU3hKVWhCenBLd1c3N3RyWWFoeCtKOS9ZWHFrUWlKbG9BQUNSbWRRaEFBY2IxSWUvMEVpYkhpaEFBaXRKMFEzeUQ5bzJ1VXJnekRlWkgrcloyZjA5S0IwMDNDRnNBQUNSa29nVkFPZDc3MFJzbXd3UUxBRXExMUZrS2QrMEJPby93SVlUd1VzRklwQnBYMkFBQWtJQ0pGZ0NGMkp4ZXhLREZuWDVDVWlaWUFGQ3NwaHZpVy80bk9remhUTFBnTWR6ems1S0pGZ0FBQ1psb0FWQVdVeTBnRFJNc0FDaGEwdzJIcGxrd0E3LzJiWDJqMFR5Q2FRS2t0RkI5QUlCMEJDMEF5aUpvQWZzbFlBSEFYSnlIRUk1MG00TGRDUlB4QklJV3BHUnREUUJBUWkvdTcrL1ZINkFnMVhyMU1ZUndvS2V3VXdJV0FNeEcwdzNINDhORTk1aVU3TC82dG42cnd6eFcwdzIrZzVQU3ozMWIrMTRLQUpDQWlSWUE1VEhWQW5aSHdBS0FPYnIwRUpIQ2ZSQ3k0QmxpRU8xRUFVbWtDaUg0ZmdvQWtNQlBpZzVRblBkYUNsc1hBeFkvYjA0dkZrSVdBTXhKMHcxeC8vc3JUYWR3NXhyTU0vaCtRRW9MMVFjQVNFUFFBcUF3bTlPTDkrTitZZUQ1QkN3QW1Mdkx1UmVBNHYxbTdEN1A1UE5EU3BYcUF3Q2tJV2dCVUNaVExlQjVCQ3dBbUwybUc4NUNDQy9uWGdlS1o1b0Z6N1ZSUVJJNmFycmhXQU1BQVBaUDBBS2dUTjQ4aEtjUnNBQ0FQMElXaCs0cG1ZRlYzOVkzR3MxejlHMzlNWVJ3cTRna1pLb0ZBRUFDZ2hZQUJkcWNYbXo4MEFPUEltQUJBSDhWMy9JL1VCTUtkaXRNeEJiNURrRktnaFlBQUFuOFE5RUJpaFhYaC95aXZmQmRNV0N4Rks0QWdEK05JOGd2bElUQ0xjZEpCTEFOOFdXSE55cEpJZ3VGQndEWVAwRUxnSEs5RmJTQWJ4S3dBSUJ2ZTZzMkZPNjZiMnVmYzdiSjl3cFNPbEY5QUlEOWUzRi9mNi9zQUlXcTFxdTRiL2hJZitFLzNvMEJDN3U0QWVBcm1tNkliOFgrajlwUXVKLzd0dlpnbksxcXVzR1ByS1Qwcjc2dE56b0FBTEEvSmxvQWxNMzZFUGlEZ0FVQVBJeTMvQ25kT3lFTGR1VGFaQUVTcXNZVk5nQUE3TWxQQ2cxUU5EK1VNM2N4WVBIUHplbkZtWkFGQUh4ZjB3M25wcUZSdUxzUXdya21zeU1DUEtTMFVIMEFnUDBTdEFBbzJPYjBJcjdOY0t2SHpKQ0FCUUE4UXRNTmgzSDZrNXBSdU11K3JUOXFNanRpbWdBcFZhb1BBTEJmVm9jQWxNLzZFT2JFaWhBQWVKb1lzamhRT3dwMjI3ZTFNQkc3SkdoQlNpOVZId0JndjB5MEFDamZwUjR6QXlaWUFNQVROZDFRQ2VZeUExYUdzRk45VzkrWUtFbEtUVGRZSHdJQXNFZUNGZ0NGR3g4NmY5Qm5DaVZnQVFEUEo1aEw2YTc3dG42dnkreUJxUmFrSkdnQkFMQkhWb2NBek1QYkVNSy85WnFDV0JFQ0FGdlFkTVByRU1LSldsSTQweXpZbDZzUXdpdlZKcEZLNFFFQTlrZlFBbUFlM2d0YVVBZ0JDd0RZTHRNc0tOMnZmVnViTXNDKytLeVJrb2tXQUFCNzlPTCsvbDY5QVdhZ1dxL2lEejR2OVpwTUNWZ0F3SlkxM2JBTUlWeW9Ld1c3Q3lFYzkyMzlVWlBabDZZYi9OaEtTdi9zMjlyM1pnQ0FQVERSQW1BK3JBOGhSd0lXQUxBRFRUY2NXNmZBREN5RkxFamdnNWNjU0NpdUQvSDlHUUJnRHdRdEFPYkQraEJ5SW1BQkFMc1ZwMWtjcURFRnUrM2IybW9jVXJnU3RDQ2h4Zmo3RHdBQU8vYVRBZ1BNdy9qQStvTjJNM0V4WVBIUHplbkZtWkFGQU94RzB3M3hJY3diNWFWd1p4cE1JaHVGSjZGSzhRRUE5c05FQzRCNXNUNkVxVExCQWdEMlo2bldGTzY2YitzclRTWVJuejFTT2xGOUFJRDlNTkVDWUY2TWoyUnFUTEFBZ0QxcXV1SE1ReGhtd0RRTGt1bmJPbjZ2dWRNQlVtbTZ3VlFMQUlBOUVMUUFtSkh4UWZadmVzNUVmQkN3QUlEOWFicmgwRFFMWm1BMVB1aUdsS3dQSWFXRjZnTUE3SjZnQmNEOG1HckJWTHlzMXF0ajNRQ0F2VGtQSVJ3cE53V0xVd1F1TlpnSnNENkVsRXkwQUFEWUEwRUxnUGtSdEdCS1h1c0dBT3hlMHcweDNIaWgxQlR1dkcvcmo1ck1CQWhha0pLSkZnQUFleUJvQVRBem05T0xqOWFITUNIMlp3UEFmbmpMbjlKOTZOdjZyUzR6RVZhSGtOTFJ1QzRNQUlBZEVyUUFtQ2RUTFpnSzYwTUFZTWVhYm9odnRyNVNad3AzcnNGTXhUaFo1WU9Ha0pEMUlRQUFPeVpvQVRCUGdoWk1pZlVoQUxCYnBsbFF1dC82dHJhcWdha3gxWUtVckE4QkFOZ3hRUXVBR2JJK2hJbng5aUVBN0VqVERYRk4xMHYxcFdCMzdpZVpLRUVMVWhLMEFBRFlNVUVMZ1BreTFZS3BPS3JXSzJOTkFXREx4djNzcGxsUXVzdStyVzkwbVFreVpZV1VmTWNHQU5neFFRdUErWG8vdnYwRlUzQ21Dd0N3ZGNzUXdvR3lVckJiWVNLbXFtOXJFeTFJNmFEcGhtTWRBQURZSFVFTGdKa2ExNGVZYXNGVXZOWUpBTmllOGVIS0wwcEs0Wlo5VzMvVVpDYnNXbk5JeVBvUUFJQWRFclFBbURkQkM2YkMraEFBMks2MzZrbmhydnUyOWpsbjZxd1BJU1hmc1FFQWRralFBbURHTnFjWDFvY3dKZGFIQU1BV05OMFEzMkE5VVVzS3Q5UmdNbUI5Q0NrSldnQUE3SkNnQlFDbVdqQVYxb2NBd0haNHk1L1N2ZXZiMnFRQWN1QnpTa3BDbHdBQU95Um9BWUNnQlZNUjE0ZllJUXNBejlCMHczbThwcW9oQllzVCtjNDFtQnowYmYweGhIQ3JXYVF5VHJrQ0FHQUhCQzBBWnM3NkVDYkcraEFBZUtLbUd3NnRVMkFHTHNlSDE1QUw2ME5JeWZvUUFJQWRFYlFBSUpocXdZUllId0lBVDNjWlFqaFFQd3AyMjdlMU1CRzVzVDZFbEV5MEFBRFlFVUVMQUlLZ0JSTnlVSzFYd2hZQThFaE5OOFEzVnQrb0c0V3pNb1FjbVdoQlNpWmFBQURzaUtBRkFOYUhNRFdDRmdEd2VKZHFSdUd1KzdZV0VDYzdmVnViYUVGS1IrTnFNUUFBdGt6UUFvQlAzcW9FRXlGb0FRQ1AwSFJEdkhhZXFCbUZPOU5nTW5hdGVTUmtmUWdBd0E0SVdnRHdpYUFGVTJGOUNBQTgwUGlXcW1rV2xPN1h2cTF2ZEptTVdSOUNTdGFIQUFEc2dLQUZBTC9ibkY3RUgzNXVWWU9KRUxRQWdJYzVqMlBCMVlxQ3hSV0hTdzBtYzRJV3BHU2lCUURBRGdoYUFQQTVPNCtaaWpmVmVtV1BMQUI4UjlNTngyUFFBa3EyN052Nm93NlR1U3NOSkNFVExRQUFka0RRQW9EUFdSL0NsSmhxQVFEZkY5L3lQMUFqQ25iYnQ3WFZPR1J2WEgxenA1TWtjdEIwZzdBRkFNQ1dDVm9BOEIvV2h6QXhnaFlBOEExTk44UXg0Ry9VaDhLZGFUQUZNZFdDbEFRdEFBQzJUTkFDZ0M5Wkg4SlV2TEkrQkFDK2FhazBGTzYzdnEwOW1LWWtHOTBrSVVFTEFJQXRFN1NvTzZVMEFBQWdBRWxFUVZRQTRFdldoekFscGxvQXdCZWFib2h2K1orb0M0VTcxMkFLSXpoRVNndlZCd0RZTGtFTEFQN2kvMmZ2Ym00aXliSTJBSWNIaVFmZ0FYZ0FxOWhXU1NHRjJJRUhuUjZRZUpEdFFiRkRJWVZVdFkxVjRVSGlBWGhBZXZBcFp1S2I2WjZ1SDM0eTg1NjQ5M2xNT0llaWdIenZlZFdIRUl5Z0JRRDhSZE1ONDdXbnRabVF1ZHUrclo4c21jeTRhRUZLcDZZUEFMQmJnaFlBL0lpckZrUXgxb2VjMkFZQS9NZjR5bjloSEdSc0sweEVqdnEyZnFtcTZ0RnlTYVhwQmxjdEFBQjJTTkFDZ0I4UnRDQVNWeTBBNE44ZmtJemh3eHV6SUhQTDZRTnB5SkdyRnFRa2FBRUFzRU9DRmdEOHcrYnk1c2xMR3dLNXRnd0ErQmV2L01uZFk5L1dRdC9rN0x2dGt0Q1o0UU1BN0k2Z0JRQS80dytjUkhHcVBnU0EwazNudmorVlBnZXl0N1JpTXVlaUJTa0pXZ0FBN0pDZ0JRQS84OVZrQ0VSOUNBQ2xjODJDM04zMWJlMjFQMW5yMjNvTVdteHRtVVNPcHhveUFBQjJRTkFDZ0I5U0gwSXc2a01BS0ZiVERlTXIvMU5mQVdScy9PQjVaY0VVd2xVTFVuTFZBZ0JnUndRdEFQZ1Y5U0ZFb1Q0RWdDSTEzWERrQTJnS3NPN2Irc21pS1lUTExhUWthQUVBc0NPQ0ZnRDhpdm9RSW5IVkFvQVNqU0dMaGMyVHNXZlZPQlRHUlF0U3VqQjlBSURkRUxRQTRLZlVoeENNb0FVQVJabDYxUCt3ZFRLMzZ0djZ4WklwaUlzV3BIUnUrZ0FBdXlGb0FjRHZxQThoaXVPeisxdG5UZ0VvaVovRHlOMUQzOWEremluS0ZDeDZ0blZTYWJyQjc5VUFBRHNnYUFIQTcvakRKNUc0YWdGQUVacHV1UERxbEFJc0xabEN1V3BCU29JV0FBQTdJR2dCd0M5dExtL0cxemJmVElrZ1Bsc0VBSVVRZGlWM2QzMWJiMnlaUXZuYUo2VUwwd2NBK0RoQkN3QmU0NnNwRVlUNkVBQ3kxM1REYXZ3L3o2YkoyTlkxQ3dvbmFFRktmcWNHQU5nQlFRc0FYa1BRZ2tqVWh3Q1FyYVliam53QVRRSFdmVnUvV0RTbDZ0dGFkUWdwblU0L2J3QUE4QUdDRmdEOGx2b1FnbEVmQWtETzFsVlZMV3lZakQzM2JiMnlZS2dlaklDRVhMVUFBUGdnUVFzQVhzdFZDNklZNjBPRUxRRElUdE1ONDRjZVZ6Wkw1bHduZzM5VEgwSktGNllQQVBBeGdoWUF2SmFnQlpFSVdnQ1FvN1d0a3JrSGxRbndILzR0a0pLTEZnQUFIeVJvQWNDcnFBOGhHRUVMQUxMU2RNUDR5di9jVnNtY2F4YndYeTVha0pLTEZnQUFIeVJvQWNCYnVHcEJGQXYxSVFEa291bUdvNnFxVmhaSzV2N3MyL3JKa3VIZnBuOFBXK01na1VYVERTZUdEd0R3Zm9JV0FMekZWMzhJSWhCQkN3QnlzYXlxNnRnMnlkaFdtQWgrU0gwSUtha1BBUUQ0QUVFTEFGNXRxZzl4MVlJb0JDMEFtTDNwTmVuU0pzbmNxbS9yRjB1R2YxQWZRa3JxUXdBQVBrRFFBb0MzRXJRZ0N2VWhBT1JnZk9XL3NFa3k5dGkzOWRxQzRZZGN0Q0FsRnkwQUFENUEwQUtBTjlsYzNxZ1BJWkpyMndCZ3JwcHVHRitTWGxrZ21YT3hCWDZpYjJ0QkMxSTZOMzBBZ1BjVHRBRGdQVnkxSUlwUFovZTNSN1lCd0V4NTVVL3V2dmtnR1g3cjBZaElwZWtHVnkwQUFONUowQUtBOXhDMElCTDFJUURNVHRNTjQxV21VNXNqYzY1WndPOXR6SWlFTGd3ZkFPQjlCQzBBZURQMUlRUWphQUhBckRUZGNPU2FCUVc0N2R2NnlhTGh0MXg5SVNVWExRQUEza25RQW9EM2N0V0NLTlNIQURBMzR5di9oYTJSc2Ewd0VieWFpeGFrNUtJRkFNQTdDVm9BOEY2Q0ZrVGlxZ1VBczlCMHcwbFZWVGUyUmVhV2ZWdS9XREw4WHQvV0d4Y2pTZWg0dXJRRkFNQWJDVm9BOEM1VGZjaXo2UkhFdFVVQU1CTmUrWk83aDc2dHY5Z3l2SW1yRnFTa1BnUUE0QjBFTFFENENGY3RpT0w4N1A3MnhEWUFpS3pwaHZFODl5ZExJbk1yQzRZMysyNWtKS1ErQkFEZ0hRUXRBUGdJTDlXSVJIMElBTkg1MlluYzNmVnQ3UU5qZUR2L2JraEowQUlBNEIwRUxRQjR0ODNselVaOUNJR29Ed0VncktZYmxtTVB1ZzJSc2ExckZ2QnVxa05JU1hVSUFNQTdDRm9BOEZIcVE0amlWSDBJQUJFMTNYRGtBMmdLc083YitzbWk0ZTM2dG43eGlJR0VGazAzK0YwYUFPQ05CQzBBK0NnbnNJbEVmUWdBRVkwaGk0WE5rTEh4QStLMUJjT0hxQThoSmZVaEFBQnZKR2dCd0llb0R5RVk5U0VBaE5KMHczaU8rdzliSVhQTDZVVSs4SDdxUTBoSmZRZ0F3QnNKV2dDd0MrcERpRUo5Q0FEUmVPVlA3aDc2dHZiN0FIeWNpeGFrNUtJRkFNQWJDVm9Bc0F2cVE0aGthUnNBUk5CMHcxaHBkVzRaWk03UFhyQURmVnU3YUVGS3A2WVBBUEEyZ2hZQWZKajZFSUw1YkNFQUJPR2FCYm03OCtFdzdOU0RjWkpLMHcydVdnQUF2SUdnQlFDNzRvTUVvamcrdTcvVkx3dEFVazAzck1iL2syeUJqRzFkczRDZEUxd2lKYjlIQXdDOGdhQUZBTHVpbDVsSXJtMERnRlNhYmpqeUFUUUZXUFZ0L1dMUnNGUGZqWk9FWExRQUFIZ0RRUXNBZG1KemVmTlVWZFdqYVJLRStoQUFVaG92ZlMxc2dJdzk5MjN0b2gzc25vc1dwT1NpQlFEQUd3aGFBTEJMWDB5VElOU0hBSkRFMUc5K1pmcGt6dlV3MklPK3JjY0hETTltU3lMSDAxVXVBQUJlUWRBQ2dGMVNIMElrUGdBQUlJV1ZxWk81aDc2dDFSdkEvcmhxUVVycVF3QUFYa25RQW9DZFVSOUNNSUlXQUJ4VTB3M2ovejNucGs3bS9Jd0YreVZvUVVvdVF3SUF2SktnQlFDN3BqNkVLQlpuOTdlZmJRT0FRNWhPYmJ0bVFlNytuS29OZ1AxeE1ZYVVYTFFBQUhnbFFRc0FkazE5Q0pFSVdnQndLTXV4Mjl5MHlkaFdtQWoyVHpVUGlibG9BUUR3U29JV0FPeVUraENDRWJRQVlPK2FiamlaZ2hhUXMyWGYxaTgyREFmaGQycFNXVFRkSUd3QkFQQUtnaFlBN01QYVZBbENmUWdBaHpEKzdMTXdhVEwyMkxlMWlrQTRIRmN0U0VuUUFnRGdGUVF0QU5nSDlTRkVJbWdCd040MDNUQjJtWDh5WVRMbllnc2Mxc2E4U2VqQzhBRUFmay9RQW9DZDIxemVqQ2VGdjVrc1FYdyt1Nzg5c2d3QTlzUWxMM0wzclc5cnIrdmhzUHliSXlVWExRQUFYa0hRQW9COWNkV0NLQmF1V2dDd0QwMDNYRmRWZFdxNFpNNDFDeml3dnEyZnFxcmFtanVKK05rR0FPQVZCQzBBMkJkQkN5SVJ0QUJncDVwdU9ITE5nZ0xjVGgvNEFvZW5Qb1JrcG1vMEFBQitRZEFDZ0wxUUgwSXduOVNIQUxCankrbHFFdVRxV1pnSWtsSWZRa3FDRmdBQXZ5Rm9BY0ErdVdwQkpLNWFBTEFUVFRlY1ZGVjFZNXBrYnRXMzlZc2xRektDRnFSMFp2b0FBTDhtYUFIQVBnbGFFSW1nQlFDNzhzVWt5ZHhEMzlhK3ppRXQxU0drSkdnQkFQQWJnaFlBN0kzNkVJSlJId0xBaDAyZDVlY21TZVpXRmd4cFRSZGxIcTJCUkk2bkMxNEFBUHlFb0FVQSsrYXFCWkc0YWdIQVIzbmxUKzd1K3JaV1dRQXh1R3BCU3E1YUFBRDhncUFGQUh1MXVid1pQNHpZbWpKQkxDMENnUGRxdW1IOGYrVFlBTW5ZMWpVTENFWFFncFFFTFFBQWZrSFFBb0JEY05XQ0tFN1A3bStkUHdYZ3pacHVPUElCTkFWWTkyMzlaTkVRaHVzeXBIUmgrZ0FBUHlkb0FjQWhDRm9RaWZvUUFONWpYVlhWd3VUSTJIUGYxc0pFRUVqZjFpNWFrTks1NlFNQS9KeWdCUUI3dDdtOCthbytoRUN1TFFPQXQyaTZZVHlkZldWb1pFN0ZHc1QwWUMra012ME1CQURBRHdoYUFIQW9ybG9RaGZvUUFONXFiV0prN3FGdmF6K3ZRMHpxUTBoSmZRZ0F3RThJV2dCd0tQNXdTeVRxUXdCNGxhWWJQanVkVFFGY3M0QzQxSWVRa29zV0FBQS9JV2dCd0VHb0R5RVk5U0VBdkpackZ1VHVybTlySCtSQ1hQNTlrcEtnQlFEQVR3aGFBSEJJcmxvUXhWZ2Y0ZzlHQVB4UzB3MnJxcXFPVFltTWJWMnpnTmo2dG42cXF1clpta2prdE9tR0k4TUhBUGduUVFzQURrblFna2hjdFFEZ3A1cHVPUEVCTkFWWTlXMzlZdEVRbnFzV3BPU1JBZ0RBRHdoYUFIQXdVMzJJbHpoRThka21BUGlGOFpyRndvREkySFBmMXFweFlCNisyeE1KWFJnK0FNQS9DVm9BY0dpdVdoREZzZm9RQUg2azZZYnhBNFVyd3lGenJudkJmTGhvUVVwK2J3WUErQUZCQ3dBTzdZdUpFNGdQR0FENGtaV3BrTG1IdnEyOWtJZVo4TytWeEZ5MEFBRDRBVUVMQUE1cWMzbXpVUjlDSU9wREFQaWJwaHZHRU42NXFaQTVZVk9ZbjBjN0k1RkYwdzBuaGc4QThIZUNGZ0Nrb0Q2RUtNYjZFSzl6QVBpWHBodU9YTE9nQUxkOVd6OVpOTXlPcXhha3BENEVBT0IvQ0ZvQWtJTDZFQ0x4b2hPQS83Y2NRM2ltUWNhMlZWV3RMUmhtYVdOdEpPU0JBZ0RBL3hDMEFPRGcxSWNRalBvUUFLcnBKUGFOU1pDNVpkL1dMNVlNcytTaUJTbTVhQUVBOEQ4RUxRQklSWDBJVVN6TzdtK0ZMUUR3eXAvY1BmWnQ3Ykljek5SVStiTzFQeEk1TjNnQWdMOFR0QUFnRlgva0pSSkJDNENDTmQwd25zUCs1R3VBekMwdEdHYlBWUXVTbVg1ZUFnQmdJbWdCUUJKVGZjaWo2Uk9Fb0FWQTJWeXpJSGZmK3JiMkFTM00zOFlPU1VoOUNBREFYd2hhQUpDU3F4WkVvVDRFb0ZCTk4xeFhWWFZxLzJSczY1b0ZaRU5naXBRRUxRQUEva0xRQW9DVXZwbytnUWhhQUJTbTZZWWoxeXdvd0xwdjZ5ZUxoaXk0YUVGS3FrTUFBUDVDMEFLQVpEYVhOMC9xUXdqazZ1eis5c2hDQUlxeUdxOGFXVGtaZXhZbWduejBiZjNpZDJnU09wNUNxZ0FBeGFzRUxRQUlRSDBJa2JocUFWQ0lwaHRPcXFyNnc3N0ozR3I2WUJiSWg2c1dwS1ErQkFCZ0ltZ0JRR3JxUTRoRTBBS2dITUtlNU82aGIydGY1NUNmNzNaS1F1cERBQUFtZ2hZQUpLVStoR0ErcVE4QnlGL1REZU9IQk9kV1RlWldGZ3haY3RHQ2xBUXRBQUFtZ2hZQVJPQ2xIWkc0YWdHUVB6OTdrTHU3dnEyOWVvY005VzB0YUVGS3FrTUFBQ2FDRmdCRW9ENkVTQVF0QURMV2RNUDR5di9ZanNuWXRxcXFwUVZEMWg2c2wwUVdUVGVjR0Q0QWdLQUZBQUZNOVNIK1VFUVU2a01BTXRWMHc1RVBvQ25BdW0vckY0dUdyTGxZUTBycVF3Q0E0bFdDRmdBRTRvUTNrVnpiQmtDVzF1TkxUS3NsWTg5OVc2OHNHTEtuUG9TVTFJY0FBTVdyQkMwQUNFUjlDSkVJV2dCa3B1bUc4VU9CSzNzbGN5NjJRQmxjdENBbEZ5MEFnT0pWZ2hZQVJMRzV2Qm5QRzMrekVJSTRQYnUvMVRzTGtKZTFmWks1aDc2dGhaZWhBRk05MExOZGs4aXB3UU1BQ0ZvQUVJcy9EQlBKWjlzQXlFUFREZVAzOUhQckpIT3VXVUJaMUllUVROTU5ybG9BQU1VVHRBQWdFa0VMSWxFZkFwQ0JwaHVPWExPZ0FILzJiZTFEVnlpTCtoQlNPak45QUtCMGdoWUFoS0UraEdEVWh3RGtZWHpsZjJ5WFpHeGJWZFhLZ3FFNHdsV2s1S0lGQUZBOFFRc0FvbkhWZ2tqVWh3RE1XTk1OSitvVUtNQ3FiK3NYaTRheTlHM3RvZ1VwdVdnQkFCUlAwQUtBYUFRdGlNU0hjd0R6TnI3eVg5Z2hHWHZ1MjFvMURwVHJ3ZTVKNUhpcVp3TUFLSmFnQlFDaHFBOGhtT096KzFzdmRRQm1xT21HOGFUMWxkMlJ1V3NMaHFLcER5RWw5U0VBUU5FRUxRQ0k2SXV0RUlnUE1BRG1hV1Z2Wk82YjZnQW9ucUFGS1htVUFBQVVUZEFDZ0hBMmx6ZGpmY2pXWmdqaXMwVUF6RXZURFdOSTd0emF5SnlLTTBEWWlwUmN0QUFBaWlab0FVQlVYMjJHSU5TSEFNekkxQmUrdGpNeWQ5dTM5Wk1sUTltbTd3TWVLWkNLVUNzQVVEUkJDd0NpRXJRZ0V2VWhBUE14dnZKZjJCY1oyd29UQVgvaHFnWEpOTjNnVVFJQVVDeEJDd0JDVWg5Q01PcERBR2FnNllhVHFxcHU3SXJNTGZ1MmZyRmtZTEl4Q0JJU3RBQUFpaVZvQVVCa3Jsb1F4VmdmSW13QkVKOVgvdVR1c1cvckw3WU0vSVdMRnFSMFlmb0FRS2tFTFFDSVROQ0NTQVF0QUFKcnVtSDhRLzhuT3lKelN3c0cvcXB2YTBFTFVuTFJBZ0FvbHFBRkFHR3BEeUVZUVF1QTJGeXpJSGQzUGxBRmZ1TFJZRWprMU9BQmdGSUpXZ0FRbmFzV1JMRlFId0lRVTlNTlMzL29KM05qK0hobHljQlBiQXlHVkthcllnQUF4UkcwQUNBNkhkUkVJbWdCRUV6VERVYytnS1lBNjc2dG55d2ErQW5YYmtoSjBBSUFLSktnQlFDaGJTNXZ4ajhZUGRzU1FRaGFBTVF6aGl3VzlrTEdubFhqQUwvaG9nVXBuWmsrQUZBaVFRc0E1a0I5Q0ZHb0R3RUlwT21HazZxcS9yQVRNcmZxMi9yRmtvR2Y2ZHQ2TTFVTVFRcUNGZ0JBa1FRdEFKZ0Q5U0ZFY20wYkFHSDRHWUhjUGZSdDdlc2NlQTFYTFVqbGVBcS9BZ0FVUmRBQ2dQQTJsemNiOVNFRTh1bnMvdmJJUWdEU2FycGh2REIwYmcxa2JtbkJ3Q3Q5TnlnU2N0VUNBQ2lPb0FVQWM2RStoRWpVaHdDa3Q3WURNbmMzMVFFQXZJYnZGNlIwWWZvQVFHa0VMUUNZQ3llVGlVVFFBaUNocGh0VzQ1bHFPeUJqVzljc2dEZHkwWUtVWExRQUFJb2phQUhBTEtnUElSajFJUUNKTk4xdzVBTm9DckR1Mi9yRm9vSFhtcjVuK0oyWlZOUzVBUURGRWJRQVlFN1VoeENKcXhZQWFZeVZJUXV6SjJQUGZWdXZMQmg0QjFjdFNLYnBCbGN0QUlDaUNGb0FNQ2ZxUTRoRTBBTGd3S1kvNEYrWk81bHpzUVY0cjQzSmtkQ0Y0UU1BSlJHMEFHQTJwdnFRUnhzamlMRSs1TVF5QUE1cWJkeGs3cUZ2YTFmY2dQY1N0Q0FsRnkwQWdLSUlXZ0F3TjY1YUVJbXJGZ0FIMG5URHRmNXZDbkJ0eWNCNzlXMnRPb1NVQkMwQWdLSUlXZ0F3TjE3NEVZa1BRd0FPb09tR282cXFWbVpONXY3czIvckprb0VQZWpCQUVqbWRmbVlEQUNpQ29BVUFzN0s1dkhsU0gwSWdwK3BEQUE1aVdWWFZzVkdUc2Ewd0ViQWo2a05JeVZVTEFLQVlnaFlBekpINkVDSlJId0t3UjAwM25FeEJDOGpacW0vckZ4c0dka0I5Q0NsZG1ENEFVQXBCQ3dEbVNIMElrYWdQQWRpdjhaWC93b3pKMkdQZjFtc0xCbmJFUlF0U2N0RUNBQ2lHb0FVQXM2TStoR0RVaHdEc1NkTU40NnZJSy9NbGN5NjJBRHZUdC9YNCsvS3ppWktJaXhZQVFERUVMUUNZSy9VaFJPSURFb0Q5OE1xZjNIM3IyOXFaZjJEWFhMVWdsY1ZVK3dZQWtEMUJDd0RtU24wSWtYeTJEWURkYXJwaHJHWTZOVll5SjZ3SjdJT2dCU201YWdFQUZFSFFBb0JabXVwRHZ0a2VRUnlmM2QvcW9nWFlrYVliamx5em9BQzMwNGwvZ0YxektZZVUvRzRNQUJSQjBBS0FPWFBWZ2tpdWJRTmdaOFpYL2d2akpHTmJZU0pnWDFRU2taaWdCUUJRQkVFTEFPWk0wSUpJMUljQTdNRFU2MzFqbG1SdTJiZjFpeVVEZS9Sb3VDUnlidkFBUUFrRUxRQ1lyYzNsell2NkVBSlJId0t3RzEvTWtjdzk5RzN0Nnh6WU4xY3RTS2JwaGd2VEJ3QnlKMmdCd055NWFrRWs2a01BUG1ENm83eFhrT1J1WmNQQUFXd01tWVE4UWdBQXNpZG9BY0RjQ1ZvUWlhQUZ3TWQ0NVUvdTd2cTI5c29jT0FSQkMxSVN0QUFBc2lkb0FjQ3NxUThobU1YWi9lMW5Td0Y0dTZZYmxtTU5rOUdSc2ExckZzQ2g5RzI5bWI3dlFBcXFRd0NBN0FsYUFKQURWeTJJUk5BQzRJMmFiamp5QVRRRldQZHQvV1RSd0FHNWFrRXF4OVBQZHdBQTJSSzBBQ0FIZ2haRUltZ0I4SFpqeUdKaGJtVHNlUXhhV0RCd1lLcUtTRWw5Q0FDUU5VRUxBR1p2cWcrNXMwbUNVQjhDOEFaTk40eC9oUC9Eek1qY3NtL3JGMHNHRGt6UWdwVFVod0FBV1JPMEFDQVhybG9RaWFBRndPdDU1VS91SHZxMjlyTXFrSUxxRUZJU3RBQUFzaVpvQVVBV05wYzM0eCt2dDdaSkVJSVdBSy9RZE1QNC9mTGNyTWpjMG9LQkZLWkxPcytHVHlLcVF3Q0FyQWxhQUpBVEx3V0pZcXdQdWJZTmdOOXl6WUxjM2ZWdDdVVTVrSkw2RUZKWlRCVnhBQUJaRXJRQUlDZUNGa1RpcWdYQUx6VGRzS3FxNnRpTXlOaldOUXNnQUdFdlVoSzBBQUN5SldnQlFEYlVoeERNcDdQNzJ5TkxBZmlucGh1T2ZBQk5BZGJUMlg2QWxGeTBJQ1ZCQ3dBZ1c0SVdBT1RHVlFzaWNkVUM0TWZHeXBDRjJaQ3g1NzZ0VnhZTXBLYStpTVF1TEFBQXlKV2dCUUM1RWJRZ0VrRUxnUC9SZE1QNEIvY3JjeUZ6MXhZTUJQSmdHU1J5YXZBQVFLNEVMUURJaXZvUWdsRWZBdkJQWHZtVHU0ZStyWjNxQnlKeDFZSmtwcEF0QUVCMkJDMEF5SkdyRmtUaXFnWEFwT21HOFpYL3VYbVFPZGNzZ0dpRXYwanB6UFFCZ0J3SldnQ1FvN1d0RXNqU01nRCtGYkk0Y3MyQ0F2elp0L1dUUlFQQnVHaEJTaTVhQUFCWkVyUUFJRHVieTV2eGowalBOa3NRcDJmM3R5ZVdBZkN2NE5teE1aQ3hyVEFSRU5FVUFQTTdNcW00YUFFQVpFblFBb0JjcVE4aEV2VWhRTkdhYmpoeDRZY0NyUHEyZnJGb0lDaFhMVWpsZVBwWkVBQWdLNElXQU9UcWk4MFNpSzUyb0hSanJkZWk5Q0dRdGNlK3JkWFhBWkY5dHgwU2N0VUNBTWlPb0FVQVdWSWZRakRxUTRCaU5kMHc5bkovOGhWQTVseHNBYUp6MFlLVUJDMEFnT3dJV2dDUU0vVWhSS0krQkNpVlYvN2s3bHZmMWw2S0E2SDVQa1ZpRnhZQUFPUkcwQUtBbktrUElSTDFJVUJ4bW00WXYvZWQyanlaYzgwQ21JdEhteUtSYzRNSEFISWphQUZBdHRTSEVNeFlIK0pjS2xDTXBodU9YTE9nQUxkOVd6OVpOREFUcmxxUVROTU5maDhHQUxJaWFBRkE3dFNIRUltckZrQkpWbFZWTFd5Y2pEMExFd0V6czdFd0VoSzBBQUN5SW1nQlFPNzg4WnRJUHRzR1VJS21HMDZxcXZyRHNzbmNxbS9yRjBzR1pzUkZDMUs2TUgwQUlDZUNGZ0JrYlhONTg2U0hsa0NPMVljQWhmaGkwV1R1b1c5clgrZkFyRXhWUjF0Ykl4Ry9Dd01BV1JHMEFLQUUvZ2hPSk9wRGdLdzEzVEMrVmp5M1pUSzNzbUJncHRTSGtNcXB5UU1BT1JHMEFLQUVYMjJaUU5TSEFMa1RjQ1IzZDMxYk83OFB6Slh2WHlRekJYSUJBTElnYUFGQTl0U0hFSXo2RUNCYlRUY3N4Kzl6Tmt6R3RxNVpBRE1uYUVGS2doWUFRRFlFTFFBb2hkZTFSTEswRFNBM1RUY2MrUUNhQXF6N3RuNnlhR0RHVkllUWtrY0hBRUEyQkMwQUtJWDZFQ0pSSHdMa2FGMVYxY0pteWRoejM5YkNSTUNzOVczOTR1SWpDYmxvQVFCa1E5QUNnQ0tvRHlHWXhkbjlyYkFGa0kybUc4YlhpVmMyU3VaY3BBSnk0YW9GcVN5YWJqZ3hmUUFnQjRJV0FKUkVmUWlSQ0ZvQU9WbmJKcGw3Nk52YWhUUWdGNElXcEtRK0JBRElncUFGQUNYeHgzRWlFYlFBc3RCMHcvajk3TncyeVp4ckZrQk92dHNtQ2FrUEFRQ3lJR2dCUURHbStwQnZOazRRNmtPQVhMaG1RZTd1K3JiMitodklodTlwSk9haUJRQ1FCVUVMQUVyanFnV1JDRm9BczlaMHc2cXFxbU5iSkdOYjF5eUFURDFZTEltNGhBWUFaRUhRQW9EU0NGb1F5ZFhaL2UyUmpRQnoxSFREaVErZ0tjQ3FiK3NYaXdZeXBENkVaSnB1Y05VQ0FKZzlRUXNBaXJLNXZIbFJIMEl3cmxvQWN6VmVzMWpZSGhsNzd0dGFOUTZRSy9VaHBIUmgrZ0RBM0FsYUFGQWlWeTJJUk5BQ21KMm1HOFkvamwvWkhKbTd0bUFnWXk1YWtKS0xGZ0RBN0FsYUFGQWlRUXNpK2FRK0JKaWhsYVdSdVllK3JYMElDV1JycWtWNnRtRVNFYlFBQUdaUDBBS0E0cWdQSVNCWExZRFphTHBoZk9WL2JtTmt6alVMb0FUcVEwamx0T2tHRHc0QWdGa1R0QUNnVks1YUVJbWdCVEFMMHgvRVhiTWdkMy8yYmYxa3kwQUJYTzRoSlZjdEFJQlpFN1FBb0ZTQ0ZrU2lQZ1NZaTJWVlZjZTJSY2Eyd2tSQVFWeTBJS1VMMHdjQTVrelFBb0FpVGZVaGQ3WlBJRTZVQTZFMTNYQlNWZFdOTFpHNVpkL1dMNVlNbEtCdmF4Y3RTRW5RQWdDWU5VRUxBRXJtcWdXUkNGb0EwYTF0aU13OTltMzl4WktCd2p4WU9JbW9EZ0VBWmszUUFvQmliUzV2dms3bm9TR0MwN1A3MnhPYkFDSnF1bUY4Y2ZqSmNzamMwb0tCQXFrUElaWEZkREVOQUdDV0JDMEFLSjJyRmtUeTJUYUFvRnl6SUhmZm5OQUhDaVZvUVVycVF3Q0EyUkswQUtCMGdoWkVvajRFQ0tmcGh2R1YvNm5Oa0xHdGF4WkF3WVRNU0VsOUNBQXdXNElXQUJSTmZRakJxQThCUW1tNjRhaXFxcFd0a0xsMTM5WlBsZ3lVYVByKzUzZGlVaEcwQUFCbVM5QUNBRnkxSUJiMUlVQWtZOGhpWVNOazdGazFEb0NyRmlSemJ2UUF3RndKV2dDQW9BV3hPRjBPaE5CMHczaGg1dy9iSUhPcnZxMWZMQmtvM0tiMEFaQk8wdzBYeGc4QXpKR2dCUURGVXg5Q01NZG45N2ZPcHdJUmZMRUZNdmZRdDdXdmN3QVhMVWpMNzc4QXdDd0pXZ0RBdi9rak81RmMyd2FRMHZTeTBDbG5jcmV5WVlCL2NkR0NsQVF0QUlCWkVyUUFnSDhUdENDU3o3WUJKT2IvUlhKMzE3ZTFGOXdBVlZWTkZVcVBaa0VpcWtNQWdGa1N0QUNBZjllSGpDOTRuczJDSU5TSEFNazAzVEMrOGorMkFUSTJWc1l0TFJqZ2IxeTFJSlhqcGh1T1RCOEFtQnRCQ3dENHI2OW1RU0RxUTRDRG0vN0k3UU5vY3JlZVhtOEQ4Rit1L0pDU3F4WUF3T3dJV2dEQWZ6bVRUaVRxUTRBVTFsVlZMVXllakQzM2JiMnlZSUIvY05HQ2xGeDBCQUJtUjlBQ0FDYnFRd2htckEveHFnYzRtS1lieGo5d1g1azRtWE94QmVBSCtyWVd0Q0Fsdi9zQ0FMTWphQUVBZjZjK2hFalVod0NIdERadE12ZlF0N1dmOVFCKzdzRnNTTVJGQ3dCZ2RnUXRBT0R2MUljUWlmb1E0Q0NhYmhpLzM1eWJOcGx6elFMZzE3NmJENGtzcHV0cUFBQ3pJV2dCQUgraFBvUmdGbWYzdDhJV3dGNDEzWERrbWdVRitOTlpmSURmOG4yU2xBUXRBSUJaRWJRQWdIOXkxWUpJQkMyQWZSdGYrUitiTWhuYlZsVzFzbUNBMzNMUmdwUUVMUUNBV1JHMEFJQi9FclFnRWtFTFlHK2FiamhScDBBQlZuMWJ2MWcwd0s5TjN5dGRlQ1NWQzVNSEFPWkUwQUlBL3NmbTh1YXBxcXBIY3lFSTlTSEFQbzJ2L0JjbVRNYWUrN1pXalFQd2VxNWFrTXFweVFNQWN5Sm9BUUEvNXFvRmtWemJCckJyVFRlTXJ3YXZESmJNK1Q4VTRHMDI1a1VxMDgrbkFBQ3pJR2dCQUQvMjFWd0k1TlBaL2UyUmhRQTd0akpRTXZmUXQ3V1gyUUJ2STJoQlNtZW1Ed0RNaGFBRkFQeUEraEFDVWg4QzdFelREZU1yLzNNVEpYT3VXUUM4a1lBYWlibG9BUURNaHFBRkFQeWMraEFpRWJRQWRxTHBodkZDenRvMHlkeHQzOVpQbGd6d0xnL0dSaUl1V2dBQXN5Rm9BUUEvcHo2RVNOU0hBTHV5cktwcVlacGtiQ3RNQlBBaDZrTkk1YmpwaGhQVEJ3RG1RTkFDQUg1Q2ZRZ0J1V29CZk1qMGgrc2JVeVJ6eTc2dFh5d1o0TjNVaDVDU3F4WUF3Q3dJV2dEQXI2a1BJUkpCQytDanZQSW5kNDk5Vy92NURlQmpYTFFnSlVFTEFHQVdCQzBBNE5mOG9aNUl4dm9RWjFTQmQybTY0V0w4UG1KNlpHNXB3UUFmMDdmMTAxVERCQ2xjbURvQU1BZUNGZ0R3QzV2TG0vSHM5RGN6SWhCWExZRDNFaDRrZDNkOVd6dDNEN0FidnArU3lybkpBd0J6SUdnQkFMLzMxWXdJNU5veWdMZHF1bUY4NVg5c2NHUnNmSG05c21DQW5WRWZRakpOTjZnUEFRRENFN1FBZ044VHRDQ1NVL1Vod0ZzMDNYRGtBMmdLc0o1TzNRT3dHeTVha0pLZ0JRQVFucUFGQVB5RytoQUNVaDhDdk1VWXNsaVlHQmw3SG9NV0ZneXdPNnFZU096Q0FnQ0E2QVF0QU9CMVhMVWdFdlVod0tzMDNUQmV3UG5EdE1qY3FtL3JGMHNHMkxsSEl5VVJGeTBBZ1BBRUxRRGdkUVF0aUVSOUNQQmFYMHlLekQzMGJlM3JIR0EvTnVaS0lxZFQvUjBBUUZpQ0ZnRHdDdXBEQ01oVkMrQ1htbTRZYTRiT1RZbk1MUzBZWUcvVWg1Q1NxeFlBUUdpQ0ZnRHdlcTVhRUltZ0JmQTdheE1pYzNkOVczdHREYkEvdnNlUzBvWHBBd0NSQ1ZvQXdPdU5RWXV0ZVJIRThkbjlyUmMrd0E4MTNiQWF2MCtZRGhuYnVtWUJzRjlUbU0zdndLVGk5MTBBSURSQkN3QjRwYWsreEZVTEluSFZBdmlIcWMvYUI5RGtidDIzOVlzdEEreWRxeGFrNHFJRkFCQ2FvQVVBdkkyZ0JaRjh0ZzNnQjhiS2tJWEJrTEhudnExWEZneHdFTitObVVRV1RUZWNHRDRBRUpXZ0JRQzh3ZWJ5Um4wSWthZ1BBZjZtNllieGU4S1ZxWkE1RjFzQURzZEZDMUx5K3k0QUVKYWdCUUM4bmFzV1JLSStCUGlydFdtUXVXSXdybjRBQUNBQVNVUkJWSWUrcmYwc0JuQTRMbHFRa3ZvUUFDQXNRUXNBZUR0LzNDY1NRUXZnWDVwdUdMOGZuSnNHbWZQL0hzQUI5VzM5TWxZMm1UbUp1R2dCQUlRbGFBRUFiNlEraEdBV1ovZTNueTBGeXRaMHcxRlZWYXZTNTBEMi91emIrc21hQVE3T1ZRdFNFU0lHQU1JU3RBQ0E5M0hWZ2tnRUxZQmxWVlhIeFUrQm5HMkZpUUNTMlJnOXFUVGQ0S29GQUJDU29BVUF2SStnQlpFSVdrREJtbTQ0bVlJV2tMUFZkTDRlZ01OejBZS1VMa3dmQUloSTBBSUEza0Y5Q01Hb0Q0R3lqYS84RjZVUGdhdzk5MjI5dG1LQU5QcTJkdEdDbEZ5MEFBQkNFclFBZ1BmN1luWUVJbWdCQldxNllYemhkMlgzWk83YWdnR1NlN0FDRW5IUkFnQUlTZEFDQU41UDBJSklCQzJnVEY3NWs3dHZmVnM3V1ErUW5xc1dwSExjZE1PUjZRTUEwUWhhQU1BN2JTNXZ4ajgwUFpzZlFZejFJVjc4UWtHYWJoai96Wi9hT1psYldqQkFDRUp2cEtRK0JBQUlSOUFDQUQ3bXEva1JpS3NXVUlqcFZaOXJGdVR1dG0vckoxc0dDTUZGQzFKU0h3SUFoQ05vQVFBZm96NkVTRDZkM2Q4NnFRcGxHRi81TCt5YWpHMkZpUURpbUlKdkxqcVNpcUFGQUJDT29BVUFmSUQ2RUFKeTFRSXkxM1REU1ZWVk4vWk01cFo5Vzc5WU1rQW9ybHFRaXVvUUFDQWNRUXNBK0RqMUlVUWlhQUg1YzAySjNEMzJiZTNySENBZVFRdFNXVXhoWXdDQU1BUXRBT0RqZkJCQUpPcERJR05OTjR4bms4L3RtTXd0TFJnZ3BPL1dRa0xxUXdDQVVBUXRBT0NEMUljUWtLc1drQy9oUG5KMzE3ZTFEL0lBQXZMOW1jVFVod0FBb1FoYUFNQnVyTTJSUUs0dEEvTFRkTVA0eXYvWWFzbll0cXFxbFFVRGhQWm9QU1FpYUFFQWhDSm9BUUM3OGRVY0NlVDg3UDVXZnkxa3BPbUdJeDlBVTRCMTM5WlBGZzBRbXFzV3BLSStEd0FJUmRBQ0FIWmdjM256NUdVUHdhZ1BnYnlNSVl1Rm5aS3haeGZDQUdaaFkwMmswblREaGVFREFGRUlXZ0RBN3VqTkp4TDFJWkNKcGh2R004bC8yQ2VaVy9adC9XTEpBT0VKV3BDUytoQUFJQXhCQ3dEWUhmVWhSSEtxUGdTeTRaVS91WHZvMjlyUFVRQXowTGYxR0xUWTJoV0p1R2dCQUlRaGFBRUFPNkkraElEVWg4RE1OZDN3V1I4MUJWaGFNc0NzdUdwQktpNWFBQUJoQ0ZvQXdHNnBEeUVTOVNFd2Y2NVprTHU3NlhVMEFQUHgzYTVJNUxqcGhpUERCd0FpRUxRQWdOMXk5cHBJeHZvUUwzNWdwcHB1V0kxL1RMWS9NcloxelFKZ2xnUXRTRWw5Q0FBUWdxQUZBT3lRK2hBQ2N0VUNabWg2cWVjRGFISzM3dHY2eFpZQlpzY2xJbEx5bUFBQUNFSFFBZ0IyejVsM0l2bHNHekJMNC84bEM2c2pZODk5VzY4c0dHQitwcENjQndhazRxSUZBQkNDb0FVQTdKNzZFQ0k1Vmg4Qzg5SjB3L2pINHl0ckkzTXVMZ0hNbTZzV3BPTDNXd0FnQkVFTEFOaXh6ZVhOK0xybm03a1NpQSt6WUY2ODhpZDNEMzFiNi9jSG1EZEJDMUpaTk4wZ2JBRUFKQ2RvQVFENzRhb0ZrYWdQZ1psb3VtRU1ScDNiRjVrVEFBU1lQNEU1VWhLMEFBQ1NFN1FBZ1AwUXRDQVM5U0V3QTAwM0hMbG1RUUgrN052NnlhSUI1cTF2YXhjdFNNbnZ0d0JBY29JV0FMQUg2a01JYUdrcEVONzQ3L1RZbXNqWVZwZ0lJQ3NQMWtraUZ3WVBBS1FtYUFFQSsrT3FCWkdvRDRIQW1tNDRxYXJxeG83STNLcHY2eGRMQnNpRytoQlNPVFY1QUNBMVFRc0EyQjlCQ3lKWm5OM2ZDbHRBWEd1N0lYT1BmVnY3T2dmSWkvb1FrbW02d1ZVTEFDQXBRUXNBMkJQMUlRUWthQUVCVFg4ay9tUTNaRTZGRlVCK0JDMUlTZEFDQUVoSzBBSUE5c3RWQ3lJUnRJQ1l2UEluZDkvNnRuWmVIaUF6ZlZzL1ZWWDFiSzhrY21id0FFQktnaFlBc0VlYnk1c3ZWVlZ0elpnZzFJZEFNRTAzWE91WXBnQ3VXUURreTFVTFVoRzBBQUNTRXJRQWdQMXoxWUpJQkMwZ2lLWWJqbHl6b0FDMzA0dG5BUExrWWhHcEhEZmRjR0w2QUVBcWdoWUFzSCtDRmtUeStleis5c2hHSUlUVmVHbkdLc2pZVnBnSUlIc3VXcENTcXhZQVFES0NGZ0N3WjV2TG02L3FRd2hrNGFvRnBEZTl2dnZES3NqY3NtL3JGMHNHeUZmZjFpNWFrSktnQlFDUWpLQUZBQnlHcXhaRUltZ0I2WDJ4QXpMMzBMZTFyM09BTWp6YU00bGNHRHdBa0lxZ0JRQWNocUFGa1h4U0h3THBOTjB3L2tINDNBckkzTXFDQVlyaHFnV3ArSmthQUVoRzBBSUFEa0I5Q0FHNWFnSHBlT1ZQN3U2Y2tnY295c2E2U2FYcEJ2VWhBRUFTZ2hZQWNEaXVXaENKb0FVazBIVERzcXFxWTdNblkxdlhMQUNLSTF4SFN1cERBSUFrQkMwQTRIQUVMWWhFZlFnY1dOTU5SejZBcGdEcnZxMmZMQnFnSE5QM2ZSY2NTY1ZGQ3dBZ0NVRUxBRGlRcVQ3azJid0p4RlVMT0t4MVZWVUxNeWRqejMxYkN4TUJsRWw5Q0trSVdnQUFTUWhhQU1CaHVXcEJKRXZiZ01PWXVxT3ZqSnZNK1g4Rm9GenFRMGpsZExvY0J3QndVSUlXQUhCWVg4eWJRRTdQN205UExBUU9ZbTNNWk82aGIydUJVb0J5Q1ZxUWtxc1dBTURCQ1ZvQXdBRnRMbTgyNmtNSVJuMEk3Rm5URGVPL3MzTnpKbk91V1FDVVRYVUlLVjJZUGdCd2FJSVdBSEI0WG5zU3liVnR3TjY1WmtIdTd2cTI5Z0ViUU1INnRuNnBxdXJSMXdDSnVHZ0JBQnljb0FVQUhKNzZFQ0pSSHdKNzFIVERxcXFxWXpNbVkxdlhMQUNZQ04yUmlvc1dBTURCQ1ZvQXdJR3BEeUVnOVNHd0IwMDNuUGdBbWdLc3BsZk1BUEM5K0FtUXltTDYyUnNBNEdBRUxRQWdEZlVoUk9LRFlOaVA4WnJGd216SjJIUGYxcXB4QVBoL0xscVFrdm9RQU9DZ0JDMEFJQTMxSVVSeWZIWi82NDlTc0VOTk40em5pNi9NbE14ZFd6QUEvNjl2YTBFTFVsSWZBZ0FjbEtBRkFDU2dQb1NBZkZnR3U3VXlUekwzMExlMUUvRUEvSzhIRXlFUmp3Y0FnSU1TdEFDQWRKemFKcExQdGdHNzBYVERHRnc2TjA0eUo2QUh3SThJNFpHS243OEJnSU1TdEFDQWRMNmFQWUdvRDRFZGFMcmhTSkNPQXZ6WnQvV1RSUVB3QStwRFNHYXE3d01BT0FoQkN3QklaSE41TTM1QThXaitCT0oxTW56Y3NxcXFoVG1Tc2ExcUhBQit3VVVMVXZKNEFBQTRHRUVMQUVqcmkva1RpUG9RK0lDbUcwNnFxcm94UXpLMzdOdjZ4WklCK0pIcC80aG53eUVSUVFzQTRHQUVMUUFnTGZVaFJETFdoemkxQ3Urbk1vVGNQZlp0TFNRS3dPK29EeUVWdjg4Q0FBY2phQUVBQ2FrUElTRDFJZkFPVXgvMEo3TWpjMHNMQnVBVjFJZVF5bkhURFVlbUR3QWNncUFGQUtUblpTaVJxQStCOTNITmd0eDk2OXZhQjJjQXZJYUxGcVNrUGdRQU9BaEJDd0JJVDMwSWtTek83bStGTGVBTm1tNFlYL21mbWhtWmM4MENnRmNSekNNeDlTRUF3RUVJV2dCQVlsTjl5SU05RUlpZ0JielNkSnA0WlY1azdyWnY2eWRMQnVBTi9JNUxLb0lXQU1CQkNGb0FRQXpxUTRoRTBBSmVid3haTE15TGpEMnJ4Z0hnSGRTSGtJcnFFQURnSUFRdEFDQUc5U0ZFb2o0RVhxSHBocE9xcXY0d0t6SzM2dHY2eFpJQmVDTkJDMUpaVEQrbkF3RHNsYUFGQUFTd3Vid1pQOEQ0WmhjRUltZ0J2K2NhRWJsNzZOdmExemtBNy9IZDFFaElmUWdBc0hlQ0ZnQVFoNnNXUkhKMWRuOTdaQ1B3WTAwM2pIKzhQVGNlTXJleVlBRGVvMi9ycDZxcXRvWkhJdXBEQUlDOUU3UUFnRGdFTFlqR1ZRdjRPYS84eWQxZDM5WmVJd1B3RWY0ZklSVVhMUUNBdlJPMEFJQWcxSWNRa0tBRi9FRFREZU1yLzJPeklXUGpDK1NsQlFQd1FSc0RKSkZUZ3djQTlrM1FBZ0JpY2RXQ1NENnBENEcvYTdyaHlBZlFGR0RkdC9XTFJRUHdRUzVha014VTlRY0FzRGVDRmdBUWk2QUYwYmhxQVgrM3JxcHFZU1prN0xsdjY1VUZBL0JSS3FoSTdNd0NBSUI5RXJRQWdFRFVoeENRb0FWTW1tNFkvMWg3WlI1a3pzVVdBSGJwMFRSSnhFVUxBR0N2QkMwQUlKNHZka0lnWTMzSWlZWEF2NnlOZ2N3OTlHM3R1aFlBdTdReFRSSngwUUlBMkN0QkN3QUlabk41TTM3QXNiVVhBbkhWZ3VJMTNURCtPemd2ZlE1a3p6VUxBSFpOZlFpcEhEZmRjR1Q2QU1DK0NGb0FRRXhla3hMSnRXMVFzdWtQdEs1WmtMcy8rN2IyNmhpQVhmTi9DeW1wRHdFQTlrYlFBZ0JpRXJRZ2tsUDFJUlJ1Zk9WL1hQb1F5TnA0U1d0bHhRRHMyaFRpYzdHUlZOU0hBQUI3STJnQkFBR3BEeUVnOVNFVXFlbUdFM1VLRkdEVnQvV0xSUU93SjY1YWtJcUxGZ0RBM2doYUFFQmNybG9RaWZvUVNqVys4bC9ZUGhsNzd0dGFOUTRBKy9UZGRFbkVSUXNBWUc4RUxRQWdMa0VMSWxFZlFuR2FiaGhmd0YzWlBKa1RwQU5nMzF5MElKVkYwdzNDRmdEQVhnaGFBRUJRNmtNSXlJZHhsTVlyZjNMMzBMZTFWOFlBN0p2L2EwaEowQUlBMkF0QkN3Q0l6VlVMSWhHMG9CaE5ONHhmNzZjMlR1WjhYd2RnNy9xMmZobXJxa3lhUkM0TUhnRFlCMEVMQUlqdGkvMFF5UEhaL2EzWFFHU3Y2WVlqMXl3b3dHM2YxazhXRGNDQnVHcEJLbjZIQlFEMlF0QUNBQUxiWE41ODkvS0hZTHgrcGdUTHNjL1pwc25ZVnBnSWdBUGJHRGlKdUZJSEFPeUZvQVVBeEtjK2hFZysyd1k1YTdyaHBLcXFHMHNtYzh2cGpEc0FISXFnQmNrMDNhQStCQURZT1VFTEFJaFBmUWlScUE4aGQxNzVrN3ZIdnEzOWJBSEFRZlZ0clRxRWxBUXRBSUNkRTdRQWdPQTJsemNiOVNFRW96NkVMRTB2M1Q3WkxwbGJXakFBaVR3WVBJbDRMQUFBN0p5Z0JRRE1nL29RSWxFZlFxNjg4aWQzMzd3b0JpQWg5U0drSW1nQkFPeWNvQVVBeklNUC80aGtyQThSdGlBclRUZU1yL3lQYlpXTWJWMnpBQ0F4WVQ5U09XNjY0Y1QwQVlCZEVyUUFnQmxRSDBKQWdoWmtvK21HbzZxcVZqWks1dFo5V3o5Wk1nQUp1V2hCU3E1YUFBQTdKV2dCQVBPaFBvUklCQzNJeVJpeVdOZ29HUnZEbW1zTEJpQ2xLZkMzdFFRU0ViUUFBSFpLMEFJQTVrTjlDSkVzMUllUWcrbUU4QitXU2VaV2ZWdS9XRElBQWFnUElaVUxrd2NBZGtuUUFnQm1ZcW9QZWJRdkFoRzBJQWRDYk9UdW9XOXJYK2NBUktFK2hGVE9UUjRBMkNWQkN3Q1lGeCtVRUltZ0JiUFdkTU5uZjNDbEFDdExCaUFRRnkxSXB1a0c5U0VBd000SVdnREF2SHkxTHdJWjYwT3VMWVFaVzFzZW1idnIyOW9IV2dDRTRmOGxFbE1mQWdEc2pLQUZBTXpJNXZMbVNYMEl3Ymhxd1N3MTNUQys4aisyUFRLMnJhcHFhY0VBQk9SM1dsSngwUUlBMkJsQkN3Q1lIL1VoUlBMcDdQNzJ5RWFZazZZYmpud0FUUUhXZlZ1L1dEUUFBYmxxUVNxQ0ZnREF6Z2hhQU1EOHFBOGhHbGN0bUp1eE1tUmhhMlRzdVcvcmxRVURFTlRHWWtqa2RBcGRBd0I4bUtBRkFNeU0raEFDRXJSZ05wcHVHRit4WGRrWW1YT3hCWURJQkMxSXlWVUxBR0FuQkMwQVlKN1VoeENKK2hEbVpHMWJaTzZoYjJ2WHJ3QUlxMi9yTVdpeHRTRVN1VEI0QUdBWEJDMEFZSjU4Z0VJMHJsb1FYdE1OMTFWVm5kc1VtYnUyWUFCbXdGVUxVbkhSQWdEWUNVRUxBSmlocVQ3a3dlNEl4QWQ3aERaMU1hOXNpY3o5MmJmMWt5VURNQVBmTFlsRVhMUUFBSFpDMEFJQTVrdDlDSkdjbjkzZm50Z0lnUzJycWpxMklESzJGU1lDWUVZRUxVaGwwWFNEMzEwQmdBOFR0QUNBK1ZJZlFqVHFRd2hwK2tQcTBuYkkzS3B2NnhkTEJtQW1WSWVRa3FzV0FNQ0hDVm9Bd0V4dExtL0dEMU8rMlIrQnFBOGhxdlg0Y3MxMnlOaHozOVpyQ3daZ0xxWnc0TE9Ga2NpWndRTUFIeVZvQVFEejVxb0ZrWnlxRHlHYXBodkcxMnFmTEliTUNib0JNRWZxUTBoRjBBSUErREJCQ3dDWU4wRUxvbEVmUWpSZStaTzdiMzFiKzZBS2dEbFNIMElxNXlZUEFIeVVvQVVBekpqNkVBTHlxcG93bW00WXZ4NVBiWVRNTFMwWWdKa1NGQ1NaNmZJZEFNQzdDVm9Bd1B5NWFrRWs2a01Jb2VtR0k5Y3NLTUJ0MzlaUEZnM0FIUFZ0N2FJRktha1BBUUErUk5BQ0FPWlAwSUpvdks0bWd2SHJjR0VUWkd3clRBUkFCaDRza1VRRUxRQ0FEeEcwQUlDWm0rcEQ3dXlSUUQ1YkJpazEzVEJlVmJteEJESzM3TnY2eFpJQm1EbFhMVWhGZFFnQThDR0NGZ0NRQjFjdGlPVDQ3UDdXNnlCUyttTDZaTzZ4YjJ0ZjV3RGs0THN0a3NqeFZEY0lBUEF1Z2hZQWtJSE41YzNYNllRNFJIRnRFNlRRZE1QNE11M2M4TW1jaWlZQWN1R2lCU2w1SUFBQXZKdWdCUURrdzFVTElsRWZRaXBlK1pPN3U3NnR2ZjRGSUF0OVd6OVZWZlZzbXlTaVBnUUFlRGRCQ3dESWg2QUZrYWdQNGVDYWJoaGYrUitiUEJrYnIxZXRMQmlBekxocVFTcUNGZ0RBdXdsYUFFQW0xSWNRa1BvUURtYnFWL1lCTkxsYlR5OS9BU0FuZ2hhazRuRUFBUEJ1Z2hZQWtCZFhMWWhFMElKREdrTVdDeE1uWStOWjliVUZBNUFobFZpa3NtaTZRZGdDQUhnWFFRc0F5SXVnQlpFc3p1NXZQOXNJK3piOWNmUVBneVp6cTc2dFh5d1pnTnowYlMxb1FVcUNGZ0RBdXdoYUFFQkcxSWNRa0tBRmgrQ1ZQN2w3Nk52Nml5MERrTEZIeXlVUlFRc0E0RjBFTFFBZ1A2NWFFSW1nQlh2VmRNUDROWFp1eW1SdWFjRUFaTTVWQzFLNU1Ia0E0RDBFTFFBZ1AxNTJFNG42RVBiTjl6eHlkOWUzOWNhV0FjaWMvK3RJNWRUa0FZRDNFTFFBZ014c0xtL0dQMUE5Mnl1QkNGcXdGMDAzcktxcU9qWmRNcloxelFLQVFyaG9RVEpOTjdocUFRQzhtYUFGQU9SSmZRaVJmRDY3dnoyeUVYYXA2WVlUSDBCVGdIWGYxaThXRFVEdStyWittZ0tHa01LWnFRTUFieVZvQVFCNSttS3ZCTEp3MVlJOVdFMWZXNUNyNTc2dFY3WUxRRUhVaDVDS2l4WUF3SnNKV2dCQWh0U0hFSkNnQlRzem5mYTlNbEV5ZDIzQkFCUkdmUWlwdUdnQkFMeVpvQVVBNUV0OUNKRjhVaC9DRG5ubFQrNGUrcmIyWVJNQXBmRi9INmtjTjkzZzkxVUE0RTBFTFFBZ1grcERpTVpWQ3o2czZZYnhsZis1U1pJNTF5d0FLSkhxRUZKU0h3SUF2SW1nQlFCa1NuMElBUWxhOENIVEt6UFhMTWpkbjMxYlA5a3lBS1hwMi9xbHFxcEhpeWNSOVNFQXdKc0lXZ0JBM3RTSEVJbjZFRDVxT1o3MU5VVXl0aFVtQXFCd3JscVFpb3NXQU1DYkNGb0FRTjdXOWtzd3JscndMazAzbkZSVmRXTjZaRzQxdmVZRmdGSUpXcENLZWtJQTRFMEVMUUFnWTV2TG15ZW5Wd2xtYVNHOGsrQVl1WHZzMjlyWE9RQ2wrMTc2QUVpbjZRYjFJUURBcXdsYUFFRCt2dGd4Z1p5ZTNkK2VXQWh2MFhURGVNYjNrNkdST1VFMEFJclh0N1dMRnFRa2FBRUF2SnFnQlFEazc2c2RFNHo2RU43S0szOXk5NjF2YXk5NEFlRGZIc3lCUkM0TUhnQjRMVUVMQU1pYytoQUN1cllVWHF2cGh2SHI1ZFRBeUp4ckZnRHdYOEtIcE9LaUJRRHdhb0lXQUZBRzlTRkVvajZFVjJtNjRjZzFDd3B3MjdmMWswVUR3SCtvRHlFVkFXOEE0TlVFTFFDZ0RPcERpRVo5Q0sreHFxcHFZVkprYkN0TUJBRC9JR2hCTWswM3FBOEJBRjVGMEFJQUNxQStoSURVaC9CTFRUZU1WMC8rTUNVeXQremIrc1dTQWVDL3BrdFB6MFpDSW9JV0FNQ3JDRm9BUURuVWh4REpXQitpLzVaZjhUMkwzRDMwYmUzckhBQit6RlVMVXZGN0tnRHdLb0lXQUZBTzlTRkU0Nm9GUHpTZDZ6MDNIVEszc21BQStLbnZSa01pZ2hZQXdLc0lXZ0JBSWFiNmtHLzJUU0NmTFlPZjhNcWYzTjMxYmUwREpBRDRPUmN0U09WNHFqRUVBUGdsUVFzQUtJdXJGa1J5ckQ2RS85VjB3M0w4MmpBWU1yWjF6UUlBZmswZ2tjVDhuZ29BL0phZ0JRQ1VSZENDYU5TSDhCOU5OeHo1QUpvQ3JQdTJmckpvQVBpdEJ5TWlrUXVEQndCK1I5QUNBQXF5dWJ4NVVSOUNNT3BEK0t0MVZWVUxFeUZqejlQWE9RRHdlK3BEU01WRkN3RGd0d1F0QUtBOHJsb1F5VmdmNHJVUTR6V0w4WStaVnlaQjVwWjlXNzlZTWdDOGlxQUZxWnliUEFEd080SVdBRkFlUVF1aVVSOUM1WlUvQlhqbzI5ci93UUR3ZXQvTmlsU21JRGdBd0U4SldnQkFZZFNIRUpENmtNSTEzZkRacXpFS3NMUmtBSGk5dnEyZnFxcmFHaG1KdUx3SUFQeVNvQVVBbE1tTFdpSlpuTjNmQ2xzVXF1bUdJOWNzS01CZDM5Yk9ud1BBMjdscVFTb3VXZ0FBdnlSb0FRQmxFclFnR2tHTGNvMnYvSTlMSHdKWjI3cG1BUUR2SnFoSUtvSVdBTUF2Q1ZvQVFJR20rcEE3dXljUVFZc0NOZDF3NGdOb0NyRHEyL3JGb2dIZ1hWeTBJSlhUNmZvZUFNQVBDVm9BUUxsY3RTQVM5U0ZsV28yN0wzMElaTzI1YjJ2Vk9BRHdmaTVha0pLckZnREFUd2xhQUVDaE5wYzNYNmR6NWhDRm9FVkJtbTY0cUtycXF2UTVrTDFyS3dhQTk1dXVRajBhSVlsY0dEd0E4RE9DRmdCUU5sY3RpT1RxN1A3V2FkWnlyRW9mQU5sNzZOdmF1WE1BK0RoWExVakZSUXNBNEtjRUxRQ2diSUlXL0I5Nzk1TVRTWkx0QzlqeTZjN2hyZ0I2QlVtdkFIcmkwK1RKSlZmTU1tb0ZGYldDQ0Zad3lSVTB6Rkl1dVI1TWZkTEpDdHBaUWNFS0xxd2ducnpiczVxcXlqK1FSSVNibTMyZmhIclNVa1djRThrZnQ1K2RFeHRUTFRKUTFtMS95Lzg0OXpxUVBOTXNBR0F6QkJjWmk0a1dBTUJYQ1ZvQVFNYXNEeUZDZ2hhSksrdTJuMXB5bm5zZFNONkhwaXJ1dEJrQU5zSkVDOGF5VjlidG9lb0RBRjhpYUFFQW1HcEJUTjVaSDVLOFJmL0FNdmNpa0xSSHEzRUFZSE9hcWhDMFlFeW1XZ0FBWHlSb0FRQUlXaEFiVXkwU05kd0dXK1plQjVLM2FLcmlRWnNCWUtOdWxKT1JIQ2s4QVBBbGdoWUFrRG5yUTRpUW9FVzZyQXdoZGJkTlZWem9NZ0JzM0NjbFpTU0NGZ0RBRndsYUFBREJWQXNpWTMxSWdzcTY3VWZ1dnN1OURpUnZvY1VBc0JYV2h6Q1dZNVVIQUw1RTBBSUFDRzZaRTZHNXBpVEg5eGxTZDkxVWhkdTJBTEFkZnNZeW1pRTBEZ0R3TzRJV0FFQy9QcVMvSFhTdkVrUkUwQ0loWmQzMnQvemY1bDRIa21lYUJRQnNTVk1WRC81bVpVVFdod0FBZnlKb0FRQjhabjBJTVhsNzlQSHNVRWVtcjZ6YmZnM01LdmM2a0x5enBpcnV0QmtBdHNyNkVNWWlhQUVBL0ltZ0JRRHcyWVZLRUpsVERVbENIN0xZeTcwSUpPM2VhaHdBMkFuclF4aUwxU0VBd0o4SVdnQUEvMko5Q0JHeVBtVGl5cnJ0cDVMOG5Ic2RTTjVxR0djT0FHeVhpUmFNNVdDWTFBY0E4QnRCQ3dEZ0tldERpSW4xSWROblVnNnB1Mm1xd3VjY0FIYWdxUW9UTFJpVHFSWUF3TzhJV2dBQVR6a3NJamFtV2t4VVdiZjlnOGpqM090QThsWmFEQUE3ZGFQY2pPUkk0UUdBcHdRdEFJRGZXQjlDaEFRdHBrdHdpOVJkdWxrTEFEdG5mUWhqTWRFQ0FQZ2RRUXNBNEk4Y2poS1RPOTJZbnJKdSsxditCN25YZ2FROW1tWUJBS01RY21Rc0psb0FBTDhqYUFFQS9KR2dCYkhvRHpKUGRXTmF5cnJkRHlFc2NxOER5VHR2cWtJUURBQjJ6MFFMeHJKWDFxMndCUUR3RzBFTEFPQjN1dG15UHppNlZSVWljTnJObGc4YU1Ubm4vVVBJM0l0QTB1NmJxakROQWdCR01BUWRIOVdla1FoYUFBQy9FYlFBQUw3RVZBdkdkdGJObHNZQ1Q4eHd3K3Q5N25VZ2VTYTJBTUM0L0ozQVdBUXRBSURmQ0ZvQUFGOXlwU3FNNkthYkxkMFduNmJ6M0F0QThtNmFxdkF6RWdER1pYMElZemxSZVFEZ00wRUxBT0JQckE5aFJQMFk0Rk1ObUo2eWJ1Y2hoT1BjNjBEeVRMTUFnUEdaYU1GWTNxbzhBUENab0FVQThEWFdoekNHMDI2MmZGRDVhU25yZGorRVlBb0pxYnRzcXNJTldnQVlXVk1WZ2hhTXBxeGJVeTBBZ0g4UnRBQUF2c1pvZEhidHJKc3RQVFNkcHY2Vy8wSHVSU0JwajZaWkFFQlVUR0JrTEVjcUR3QUVRUXNBNEd1c0QySEhicnJaMGtTRUNTcnI5dEFCTkJsWU5WVmgyZzRBeE1PVUtjWmlvZ1VBOEMrQ0ZnREF0MWdmd2k3ME44VlBWWHF5K29ETVh1NUZJR24zVFZXY2F6RUFSTVVrUE1aaW9nVUE4QytDRmdEQXR3aGFzQXVuM1d6cHB2Z0VEZnVKMytkZUI1STMxMklBaUk2SkZvemxZSmpxQndCa1R0QUNBUGlxNGZEN1dvWFlvck51dG5RYmJicmM4aWQxTjAxVitCNEZBSkZwcXFJYkp1UEJHRXkxQUFBRUxRQ0E3N3BTSXJia3Bwc3RWNG83VFdYZDlyZjgzK1plQjVKbm1nVUF4TXRVQzhZaWFBRUFDRm9BQU44bGFNRTI5TGZQVGxWMm1zcTYzVGZOZ2d5Y05WVnhwOUVBRUMxVHB4akxpY29EQUlJV0FNQTNXUi9DbHB3T255Mm1hUkZDMk5NN0V2WW9UQVFBMFRQUmdyRWNxendBSUdnQkFEeUhxUlpzMG9kdXRuVDdiS0xLdWowTUlTeHpyd1BKV3pSVklRd0dBSEh6TndXaktldlcraEFBeUp5Z0JRRHdISUlXYk1wdE4xc3VWSFBTM1BJbmRiZE5WVnpvTWdERWJRaEYzbXNUSXhHMEFJRE1DVm9BQU45bGZRZ2Iwby9pUDFYTTZTcnJ0dDlGL0M3M09wQThZVEFBbUE1VExSakxpY29EUU40RUxRQ0E1ekxWZ3RlYWQ3UGxuU3BPbWx2K3BPNjZxUW9ITmdBd0haMWVNUklUTFFBZ2M0SVdBTUJ6WFEwVENlQkhmT2htUzJHZENTdnJ0ci9sZjVCN0hVamFvMmtXQURBNUFwS001YTNLQTBEZUJDMEFnR2NaMW9jNEtPZEgzSGF6cGNQTENTdnJkaitFc01xOURpVHZ2S2tLVTNjQVlFS2FxakRSZ3RFTXF4VUJnRXdKV2dBQUx5Rm93VXYxTjhSUFZXM3krcERGWHU1RklHbjNmZEJDaXdGZ2ttNjBqWkVJV2dCQXhnUXRBSUJuRzFZL1dCL0NTOHk3MmRJTjhRa3I2L1l3aFBCejduVWdlYXVtS2g2MEdRQW15VlFMeG5LazhnQ1FMMEVMQU9DbFRMWGd1VDRNNFJ5bTdVTC9TTnhOVXhVKzV3QXdYWi8wanBHWWFBRUFHUk8wQUFCZXlzRTV6M0hielpZTGxacTJzbTc3dFMvSHVkZUI1SzIwR0FBbXpVUUx4ckkzVEFBRUFESWthQUVBdklqMUlUeEQvL2s0VmFna25PZGVBSkozMlZTRlc3QUFNR0ZOVmZTckN1LzFrSkZZSHdJQW1SSzBBQUIraEtrV2ZNdThteTN2VkdqYXlycnRiL2tmNUY0SGt0YUh3a3plQVlBMG1HckJXS3dQQVlCTUNWb0FBRDlDMElLditUQk1QV0hDeXJyZGR3Qk5CczZicW5qUWFBQklncUFGWXpIUkFnQXlKV2dCQUx5WTlTRjh4VzAzV3pxY1QwTy9NbVF2OXlLUXRQdW1LbFphREFESnNBcU1zUnlyUEFEa1NkQUNBUGhSRnlySEUzM3c1bFJCcHErczIzNzA3ZnZjNjBEeWhNSUFJQ0ZOVlFoYU1KcXliazIxQUlBTUNWb0FBRDlLMElLbjV0MXNlYWNpU1hETG45VGRORlZoeFJFQXBPZFdUeG5KaWNJRFFINEVMUUNBSDlMTmx2ME8zSHZWSTRUd1lWZ253OFNWZFRzMytwWU1tR1lCQUdreTFZS3htR2dCQUJrU3RBQUFYc1BoT3JmZGJPblFNZ0ZsM2U2YlprRUdQalJWMFdrMEFDVEp6M2pHSW1nQkFCa1N0QUFBWHNQNmtMdzloaEJPY3k5Q1F2ckF6RUh1UlNCcGo4SkVBSkEwUVF2RzhuWUlyZ01BR1JHMEFBQittUFVoMlp0M3MrVmQ3a1ZJUVZtM2g5WXBrSUZWVXhVUEdnMEFhUnFtVmoxcUx5TXgxUUlBTWlOb0FRQzhsdlVoZWZyUXpaWjZuNDd6RU1KZTdrVWdhZmROVlp4ck1RQWt6MVFMeG5LaThnQ1FGMEVMQU9DMXJBL0p6MjAzVzVwK2tJaXlidnNIZ3U5eXJ3UEptMnN4QUdUaGt6WXpFa0VMQU1pTW9BVUE4Q3JXaDJTbkg4VjdtbnNSRXVPV1A2bTdicXJDb1FzQTVNSFBmTVppZFFnQVpFYlFBZ0RZQkZNdDhqSHZac3U3M0l1UWlySnUrMXYrYjNPdkE4a3pnUWNBOG1GMUNHUFpLK3YyVVBVQklCK0NGZ0RBSmdoYTVPRkROMXRlNVY2RVZKUjF1MithQlJrNGE2cENPQXdBTXRGVXhZT0ppNHpJK2hBQXlJaWdCUUR3YXNPRWcxdVZUTnB0TjF1NkZaNld2cDk3dVJlQnBEMEtFd0ZBbHF3UFlTeldod0JBUmdRdEFJQk5NZFVpWGYxaDVXbnVSVWpKTU5KMm1Yc2RTTjVpdU5VS0FPVEYraERHSW1nQkFCa1J0QUFBTnNWS2lYVE5oNmtscEVNd2l0VGRObFhoY3c0QWVUTFJnckVjcXp3QTVFUFFBZ0RZQ090RGt2V2hteTJGYUJKUzF1MkpCNEJrd0tvakFNaFVVeFVtdTFTVTdRQUFJQUJKUkVGVVdqQ2E0ZTh0QUNBRGdoWUF3Q2E1UFp5VzIyNjJkRmlaSHY5T1NkMWxVeFZ1c2dKQTNtNXlMd0Nqc1Q0RUFESWhhQUVBYkpMSkIrbDREQ0djNWw2RTFKUjEyd2RuRG5LdkEwbnJ2M2V0dEJnQXNpZDB5VmhNdEFDQVRBaGFBQUFiWTMxSVV1WkRQMGxFV2JmN0RxREp3SGxURmI1M0FRRFdoekFXRXkwQUlCT0NGZ0RBcHAycjZPUjk2R1pMMDBuUzAvL2IzTXU5Q0NUdDNzOGdBR0FnYU1GWURvYVFPd0NRT0VFTEFHRFRITkJQMjIwM1d5NXlMMEpxeXJydGIxVzl6NzBPSkcvVlZNV0ROZ01BdzRTcisrd0x3VmlzRHdHQURBaGFBQUFiMWMyVy9TSFh0YXBPMG1NSTRUVDNJaVRLTFg5U2Q5TlV4WVV1QXdCUG1HckJXS3dQQVlBTUNGb0FBTnRncXNVMHpidlo4aTczSXFTbXJOcytQSE9jZXgxSW5razhBTUFmZlZJUlJtS2lCUUJrUU5BQ0FOZ0dRWXZwK2RETmx2cVdKdE1zU04xbFV4VnVyQUlBZitUM0E4WmlvZ1VBWkVEUUFnRFlPT3RESnVlMm15M2RCazlRV2JlckVNSkI3blVnYVkrbVdRQUFYOUpVaFlrV2pHV3ZyRnRoQ3dCSW5LQUZBTEF0cGlOTVEzOUllWnA3RVZKVTF1MmhBMmd5Y041VXhZTkdBd0JmY2Fzd2pFVFFBZ0FTSjJnQkFHeUxvTVUwekx2WjhpNzNJaVNxbjJheGwzc1JTTnA5VXhVckxRWUF2c0ZVQzhZaWFBRUFpUk8wQUFDMnd2cVFTZmpRelpZQ01Ra3E2L1lraFBBKzl6cVFQQk5iQUlEdjZWU0lrWndvUEFDa1RkQUNBTmdtaC9qeHV1MW1TNGVVNlhMTG45VGRORlhoWnd3QThEMkNGb3pscmNvRFFOb0VMUUNBcmVsbXk0c1F3cU1LUjZmdnlXbnVSVWhWV2JmekVNSng3blVnZVhNdEJnQytwS3piL2Y1MzRySnUrMURtUHhXSnNReVRCZ0dBUlAyWHhnSUFXM1psaFVGMDV0MXNlWmQ3RVZMVVAxUTJ6WUlNZkdpcXd2Y3dBT0EzdysvQnA4UFhPNVVoRW4zUTRwTm1BRUNhQkMwQWdHMFR0SWpMaDI2Mk5HNC9YZjA2bUlQY2kwRFNIb1dKQUlBZ1hNRTBIT2tUQUtUcnpYcTkxbDRBWUt1T1BwNDloQkQyVkhsMHQ5MXM2VUZQb3NxNlBRd2gvSnA3SFVqZUwwMVZuR3N6QU9SSnVJS0p1VytxNGxEVEFDQk4vMGRmQVlBZE1FRmhmUDB0OEhudVJVaWN3MmRTZHl0a0FRQjVLdXYydEt6Yi91L0svdzBoL0YzSWdvazRHQUx4QUVDQ3JBNEJBSGJCK3BEeExiclpzc3U5Q0trcTYvYkV3Mll5c05Ca0FNaEhINjU0TXIzQ2hFU21xcDhxZWFkN0FKQWVRUXNBWU91NjJmTHE2T1BabzRkam83bnNac3VMVE45N0x0enlKM1hYVFZWODBtVUFTSnR3QlFrNk11VVRBTklrYUFFQTdJcXBGdU80ZFFzOGJXWGQ5aXRoM3VaZUI1TG4reGdBSkVxNGdzU2RhREFBcEVuUUFnRFlGVUdMM2V1bmlNeTcyZklodHplZWk3SnU5MDJ6SUFOblRWVVl0d3dBQ1JHdUlDUEhtZzBBYVhxelhxKzFGZ0RZaWFPUFovMUIyWUZxNzh4UFZvYWtyYXpiUG1UeGMrNTFJR2w5WU95d3FRcUJNUUNZT09FS012YlhwaW82SHdBQVNJdUpGZ0RBTGwwNUZONlpTeUdMdEpWMWUramZFeGxZQ0ZrQXdIUUpWOEMvSElVUUJDMEFJREdDRmdEQUxsMDRHTjZKMi81d01vUDNtVHRCR2xKMzAxU0Z6emtBVEl4d0JmekppYi9mQUNBOWdoWUF3TTUwczJWMzlQSHMzdnFRcmVySDdNKzcyZElOOElTVmRYdGkxeThaV0dreUFFeERXYmY5amYzNUVLN3c5eDc4M3BGNkFFQjZCQzBBZ0YyelBtUzdGbjJnSmVVM3lMKzREVVhxTHB1cStLVExBQkF2NFFwNHRyZGwzZTViaVFjQWFSRzBBQUIyemZxUTdibnNaa3NIOElrcjYzYmxRVGFKZXpUTkFnRGlKRndCUDZ6L3R5TklEQUFKRWJRQUFIYksrcEN0dWUybldTVDYzaGowdDZEMG1ReWNOMVZ4cDlFQUVBZmhDdGlJRTBFTEFFaUxvQVVBTUFiclF6YXJ2LzA5NzJaTFkwalRkeDVDMk11OUNDVHRmdmljQXdBakVxNkFqVHRTVWdCSWk2QUZBREFHNjBNMmE5RlBDa25wRGZGbnc4UHU5MHBENGhaMlZ3UEFPSVFyWUt0T2xCY0EwdkptdlY1cktRQ3djMGNmeis0OHZOdUl5MjYybkNmd1B2aU9zbTc3TWJQSDZrVENicHFxOEFBYUFIWkl1QUoyNmk5VzVBRkFPa3kwQUFERzBvK0cveC9WZjVYYi92YjNoRjgvejFUVzdhbVFCUm53L1F3QWRxQ3MyOE1oWERFWHJvQ2Q2b05OZ2hZQWtBaEJDd0JnTEZlQ0ZxL3kyRDhZN1daTEkvWVRWOWJ0L2hCTWdwUmRObFZoQlJJQWJNa1Fyamdkd2hWdjFSbEdjVEk4Q3dFQUVpQm9BUUNNb3BzdDc0NCtudDE2eVBmREZ0MXM2VkF5RHdzM0RVbmNvMmtXQUxCNXdoVVFuU010QVlCMENGb0FBR082TU5YaWgxeDJzK1hGQkY4M0x6UThISGNBVGVyT202b3duUWNBTmtDNEFxSm1IU1FBSkVUUUFnQVlrL1VoTDNmcjREMHJxeERDWHU1RklHbjNUVldzdEJnQWZweHdCVXhIV2JkSFZ1WUJRQm9FTFFDQTBWZ2Y4bUw5ZVAxNU4xdTYrWjJCc203Ny9iM3ZjNjhEeVp0ck1RQzhuSEFGVEZiL2Q1NmdCUUFrUU5BQ0FCaWI5U0hQdCtobVN3OWs4dUdXUDZtN2FhcmlreTREd1BNSVYwQVNqclFSQU5JZ2FBRUFqTTM2a09lNTdHYkxpeW04VUY2dnJOdTUvYjFrd0RRTEFQZ080UXBJem9tV0FrQWEzcXpYYTYwRUFFWjE5UEdzODlEd20vcjFLaWRXaHVTaHJOdjlFTUpkQ0dFdjkxcVF0QTlOVlN5MEdBRCtiUGg5OEhUNGVxZEVrSnovYnFyQzMvY0FNSEVtV2dBQU1UZ1BJZnhkSjc3b3NiKzlKbVNSbFlXUUJZbDd0Qm9IQUg1UHVBS3kwcThQc1VJUEFDWk8wQUlBaU1HVm9NVlhMYnJac292MHRiRmh3MmpvcGJxU3VKVWJmQUFnWEFFWk94RzBBSURwRTdRQUFFYlhUMnM0K25oMjdlSGluMXgycytWRlpLK0o3VHBYWHhKMzIxU0Z6emtBMlJLdUFJYWdCUUF3Y1lJV0FFQXNyanhvL0ozYllZVUVtU2pyOXNTL0FUTGcreG9BMlJHdUFQN2dTRUVBWVBvRUxRQ0FXRmdmOGgrUElZUjVQK2tqbGhmRVRyamxUK3F1bTZvd0lobUFMQWhYQU4rdzE2K05iS3JpVHBFQVlMb0VMUUNBS0ZnZjhqdUxicmJzSW5vOWJGbFp0LzB0LzdmcVRPSk1zd0FnYWNJVndBdjBFdzJ0Q2dXQUNSTzBBQUJpWW4xSUNKZmRiT2xoUzBhR0IvS3IzT3RBOHM3YzJBTWdWV1hkOXNHS3ViOWxnQmV3UGdRQUprN1FBZ0NJU2U3clEyN2QrTTVTSDdMWXk3MElKTzNlYWh3QVVqT0VLejUvK1YwT2VDbEJDd0NZdURmcjlWb1BBWUJvSEgwOHkzV3F4V00vT3RUS2tMejBlM2xEQ0wvbVhnZVM5MU5URlNiMUFEQjV3aFhBSmpWVjhVWkJBV0M2VExRQUFHSnprV25RWWlGa2tTV0h6NlR1UnNnQ2dDa1RyZ0MycGF6Yms2WXFQaWt3QUV5VG9BVUFFSlZ1dHJ3NituajJtTmxEek10dXRuUVFtWm5ob2YxeDduVWdlU3N0Qm1CcWhDdUFIZW5YaHdoYUFNQkVDVm9BQURIcTE0ZTh6NlF6dC8wMGl3aGVCN3QzcnVZazd0SU5QUUNtUXJnQ0dNR0p2d3NCWUxvRUxRQ0FHT1VTdE9nbmQ4eTcyZkloZ3RmQ0RwVjEyOS95UDFCekV2Wm9tZ1VBc1JPdUFFWjJwQUVBTUYxdjF1dTE5Z0VBMFRuNmVQYVF3Y1BPbjZ3TXlVOVp0L3NoaERzUDgwbmNXVk1WZ2hZQVJLZXMyLzRHK1Z5NEFvakVmemRWNGZJRkFFeVFpUllBUUt4U24ycHhLV1NSclhNUDlVbmN2WkFGQURFcDYvYm9TYmpDVkRFZ0ppZkQ4dzhBWUdJRUxRQ0FXS1VjdExnTklTd2llQjNzMlBDUVA0ZTFPT1ROOXpjQVJpZGNBVXpFa2FBRkFFeVRvQVVBRUtWdXRydzYrbmoybU9ETi8vNDl6YnZaMG1qUVBKM25YZ0NTZDlOVWhRZkZBSXhDdUFLWW9CTk5BNEJwRXJRQUFHS1c0bFNMUlRkYmRoRzhEbmFzck52K29mK3h1cE00MHl3QTJDbmhDbURpampRUUFLYnB6WHE5MWpvQUlFcEhIOC82bXgzL1NLZzdsOTFzT1kvZ2RiQmpaZDN1OTROYVBQd25jWmROVmZnZUI4RFdDVmNBaWZsclV4VXVaQURBeEpob0FRQkVxNXN0UHgxOVBMdFA1T0hwclp2ZVdWczRCQ0J4ajc3SEFiQk53aFZBd282R1lENEFNQ0dDRmdCQTdQcjFJVDlQdkV2OUFlUzhteTBmSW5ndDdGaFp0NGNPb01uQXFxa0szK01BMkNqaENpQVQvVFRQQzgwR2dHa1J0QUFBWW5lUlFOQmkwYzJXYnFma2F4VkMyTXU5Q0NUdHZxbUtjeTBHWUJPR2tPcXBpV0JBUm80MEd3Q21SOUFDQUloYUgxQ1krUHFRNjI2MmRETWxVMlhkOWplVDN1ZGVCNUkzMTJJQVh1Tkp1S0wvbWZKV01ZSE0rTDRIQUJNa2FBRUFUTUZVMTRmY080RE1ubHYrcE82bXFZcFB1Z3pBU3dsWEFQeEhIOUwzZXpVQVRJdWdCUUF3QlZOZEgzTGF6WllQRWJ3T1JsRFdyVU1EY2lCTUJzQ3pDVmNBZkZVL0RWSFFBZ0FtUk5BQ0FJamVSTmVIL05LLzdnaGVCeU1vNjNiZk5Bc3k4S0dwaWp1TkJ1QmJoQ3NBbnVWSW1RQmdXZ1F0QUlDcG1OTDZrT3R1dG5USW5yZEZDR0V2OXlLUXRNY1F3a3FMQWZnUzRRcUFGeE8wQUlDSkViUUFBS1ppS3V0RDdvM1N6OXR3c0xETXZRNGtiOUZVaGRWSUFQeEd1QUxnVlE3Njc2TW14Z0hBZEFoYUFBQ1RNS3dQdVozQVE5dlRiclowK0pnMzAweEkzVzFURlJlNkRNQ3dMbTB1WEFHd0VmMVVDMEVMQUpnSVFRc0FZRXI2ZzczL2lmajEvdElIUWlKNEhZeWtyTnVURU1JNzlTZHhDdzBHeU5jUXJqZ2R2dnplQTdBNVI4UGFWQUJnQWdRdEFJQXB1WW80YUhIZHpaWW1HZUNXUDZtN2JxcmlreTRENUVXNEFtQW5UcFFaQUtaRDBBSUFtSXh1dHJ5TGRIM0kvVEF1bVl5VmRkdmY4ai93R1NCaGo2WlpBT1JEdUFKZzU0NlZIQUNtUTlBQ0FKaWFHTmVIbkhhejVVTUVyNE9SREFjUksvVW5jZWROVmRnWkRaQXc0UXFBY1pWMWU5UlVoWldrQURBQmdoWUF3TlRFdGo3a2wyNjI5QkNFUG1TeGwzMFZTRmsvdWNkNkpJQUVDVmNBUktWZkgrSVpBd0JNZ0tBRkFEQXBrYTBQdWU1bVN3ZVBtZXR2SElVUWZzNjlEaVJ2MVZTRnlUMEFpUkN1QUlqV2tkWUF3RFFJV2dBQVV4VEQrcEQrZHZkOG11Vmp3NFJ0U04xTlV4VVh1Z3d3YmNJVkFKTWdhQUVBRS9GL05Bb0FtS0NyQ0Y3eWFUZGJ1dDJkdWJKdSs0T0s0OXpyUVBKV1dnd3dYZjN2SzJYZDlvRzUvdzBoL0YzSUFpQnFiNGRnSEFBUU9STXRBSURKR2RhSDNJeDR3UDFMTjF2YW1Vb3d6WUlNWERaVjhVbWpBYVpsQ0lOKy90clRQb0JKNmFkYStCMGNBQ0luYUFFQVROWEZTRUdMNjI2MmRMaE9mNERSMy9JL1VBa1M5aGhDV0dnd3dEUUlWd0FrNDBUUUFnRGlKMmdCQUV6VjFURDZlSmZ1UXdoem54aUdVYTRPb0VuZGVWTVZWaVFCUkV5NEFpQkpSOW9LQVBFVHRBQUFKcW1iTFIrT1BwNWQ3M2pIOUduLzMvV0pZVmdaNGpDRGxOMDNWYkhTWVlENENGY0FKTzlFaXdFZ2ZvSVdBTUNVWGUwd2FQRkxOMXQyUGkyVWRkcy85SHFmZlNGSW5Za3RBQkVScmdESXlsNVp0NGROVmR4cE93REVTOUFDQUppeVhhMFB1ZTVteTNPZkZBWnUrWk82bTZZcXJuUVpZRnpDRlFCWjY5ZUhDRm9BUU1RRUxRQ0F5ZHJSK3BEN0VNTGNwNFR3N3dPUC9yTndyQmdrempRTGdKRU1rN1Btd2hVQTJUc1pMcGNBQUpFU3RBQUFwbTdiNjBOTyswQ0hUd2xsM2U2YlprRUdQalJWWVUwU3dBNlZkWHYwSkZ4eG9QWUFEQk10QUlDSUNWb0FBRk8zemZVaHYzU3pwUU5IUGxzNC9DQnhqOEpFQUxzaFhBSEFkNWlrQ0FDUmU3TmVyL1VJQUppMG80OW4yNWhxY2QzTmxxYytHWVIvSDRZYzl0dHFqUEFtY2I4MFZYR3V5UURiSVZ3QndBdjlyYW1LVDRvR0FIRXkwUUlBU01IRmhvTVc5OE5EY1Bqc1hNaUN4TjBMV1FCc25uQUZBSy9RL3d3UnRBQ0FTQWxhQUFDVDE4MldWMGNmeng0M2VCQisyczJXRHo0WmhIOGZrSnhzWVdJS3hFYTRER0JEaENzQTJKQWpoUVNBZUFsYUFBQ3A2TmVIdk4vQWUvbWxteTA3bndxZWNNdWYxTjBZU1F6d09zSVZBR3pCaWFJQ1FMd0VMUUNBVkd3aWFISGR6WllPMWZsTldiZjlnY2xiRlNGeHBsa0EvQURoQ2dDMjdLQ3MyLzJtS2t6Y0JJQUlDVm9BQUVuWXdQcVFlNGVOUE5VLzBETE5nZ3ljTlZWeHA5RUF6MVBXN2VFUXJGZ0lWd0N3QTMyb3ovUTVBSWlRb0FVQWtKTFhUTFU0N1daTHQwUjRhdldLNEE1TXdhTXdFY0QzUFFsWG1IUUZ3SzZkQ0ZvQVFKd0VMUUNBbFB4bzBPS1hicmJzZkJMNGJEaFErVmxCU056Q0dHS0FMeE91QUNBU0p4b0JBSEVTdEFBQWt2R0Q2ME91dTluU2pXNys2RUpGU054dFV4VSs1d0JQQ0ZjQUVLRWpUUUdBT0FsYUFBQ3BlY2xVaS92aFFUcjhwcXpiL3NiUXNZcVF1SVVHQXdoWEFCQzl2ZjVuVlZNVmQxb0ZBSEVSdEFBQVVuUCtncURGYVRkYkdwdlBIN25sVCtvdW02cXc1eG5JbG5BRkFCTno0dTlVQUlpUG9BVUFrSlJ1dHV5T1BwNzFreW9PdnZPK2Z1bi92N3JQVTJYZExwN3gyWUVwNjljcnJYUVF5STF3QlFBVFpuMElBRVJJMEFJQVNGRy9QdVRuYjd5djYyNjJQTmQ1bmlycmR0OEJOQms0TjNZWXlNWHdzLzEwV0pja1hBSEFWSjNvSEFERVI5QUNBRWpSeFRlQ0Z2ZkRUVWI0b3o1OHM2Y3FKT3grK0p3REpPdEp1S0wvZXFmVEFDUkFXQkFBSXZSbXZWN3JDd0NRbktPUFozZGZXUUh4Vnl0RCtLT3lidnRSclA5VUdCTDNVMU1WZGpzRHlSR3VBQ0FEZjJ1cTRwTkdBMEE4VExRQUFGTDFwZlVodndoWjhCVnUrWk82R3lFTElDWENGUUJrcHI4Y0lHZ0JBQkVSdEFBQVV2WEg5U0hYM1d6cE1KMC9LZXUyUDZBNVZoa1N0OUJnWU9xRUt3REkySWtMQWdBUUYwRUxBQ0JKL2VTS280OW45OFA2a1A1LzV6ck5WM2hZUmVvdW02b3d6UWVZSk9FS0FQaVhJMlVBZ0xnSVdnQUFLZnU4UHVTMG15MGZkSm8vS3V0Mk5ZUnhJRldQcGxrQVV5TmNBUUIvY3REL2ZHeXF3ck1OQUlpRW9BVUFrTEorVWtIWFQ3ZlFaZjZvck50REI5Qms0TnpEV0dBS2hDc0E0THRPaGdzbEFFQUVCQzBBZ0dSMXMrVmRDT0ZDaC9tS2ZwckZudUtRc1B1bUtsWWFETVNzck52VEp3RUxQNWNCNE91T0JDMEFJQjZDRmdBQVpLZXMyLzRtMEh1ZEozRW10Z0JSRXE0QWdCOXlvbXdBRUE5QkN3QUFjdVNXUDZtN2FhckNiVGNnR3NJVkFQQnFSMG9JQVBGNHMxNnZ0UU1BZ0d5VWRUc1BJZnhkeDBuY1g1cXF1Tk5rWUV6Q0ZRQ3djWDl0cXFKVFZnQVluNGtXQUFCa282emJmZE1zeU1BSElRdGdMTUlWQUxCVi9WUUxRUXNBaUlDZ0JRQUFPVm1FRUE1MG5JUTlDaE1CdXlaY0FRQTdjeEpDdUZCdUFCaWZvQVVBQUZrbzYvWXdoTERVYlJLM2FxcmlRWk9CYlJPdUFJQlJIQ2s3QU1SQjBBSUFnRnljNnpTSnUyK3F3dWNjMkpxeWJvK0c2VkRDRlFBd2pyZnFEZ0J4RUxRQUFDQjVaZDMyNDFYZjZUU0ptMnN3c0dsRHVHSStoQ3VzM3dLQWNkMzNQNXVicXVqMEFRREdKV2dCQUVBTzNQSW5kZGROVlh6U1pXQVRoQ3NBSUNyM0lZU3JFTUtGZ0FVQXhFUFFBZ0NBcEpWMXV6QmVsUXdzTkJsNERlRUtBSWlLY0FVQVJFN1FBZ0NBWkpWMXV4OUNXT2t3aVR0cnF1Sk9rNEdYRXE0QWdLZ0lWd0RBaEFoYUFBQ1FzajVrc2FmREpPelJhaHpnSllRckFDQXFqMC9DRlZZQkFzQ0V2Rm12MS9vRkFFQnl5cm85RENIOHFyTWs3cWVtS2k0MEdmZ1c0UW9BaU1ybmNNVlZVeFZYV2dNQTAyU2lCUUFBcVhMNFRPcHVoU3lBcnhHdUFJQ29DRmNBUUdJRUxRQUFTRTVadHljaGhHT2RKWEVMRFFhZUdxWTVuUTRCaTdlS0F3Q2pFcTRBZ0lRSldnQUFrQ0szL0VuZHBSM09RQkN1QUlEWUNGY0FRQ1lFTFFBQVNFcFp0eXNqMGtsYy8vQjJwY21RTCtFS0FJaUtjQVVBWkVqUUFnQ0FaSlIxdTIrZEFoazRiNnJpVHFNaEw4SVZBQkFWNFFvQXlKeWdCUUFBS1RrUEllenBLQW03SHo3blFBYUVLd0FnT3RmOXFrcmhDZ0RnelhxOXpyNElBQUJNWDFtM1J5R0VmMm9saWZ1L0h1cEMyb1FyQUNBNjEwK21WenhvRHdBUVRMUUFBQ0FoYnZtVHVoc2hDMGlUY0FVQVJFZTRBZ0Q0SmtFTEFBQW1yNnpiL25EcVdDZEozRUtESVIxbDNlNFA0WXFGY0FVQVJFRzRBZ0I0TmtFTEFBQW1iVGlvTXMyQzFGMDJWZEhwTWt6YmszQkYvL1ZPT3dGZ2RNSVZBTUFQRWJRQUFHRHErcHZBQjdwSXdoNU5zNERwRXE0QWdPZ0lWd0FBcnlab0FRREFaQTA3N1IxQWs3cHpENEJoV29RckFDQTZ3aFVBd0VZSldnQUFNR1dyRU1LZURwS3crNllxVmhvTThST3VBSURvQ0ZjQUFGc2phQUVBd0NTVmRYc1NRbml2ZXlSdXJzRVFMK0VLQUlqT2JRamhZZ2hYM0drUEFMQXRnaFlBQUV5VlcvNms3cWFwaWsrNkRIRVJyZ0NBNkFoWEFBQTdKMmdCQU1Ea2xIWGIzL0kvMWprU1o1b0ZSR1Q0MlNOY0FRQnhFSzRBQUVZbGFBRUF3S1FNTjRuUGRZM0VmZkRBR01aWDF1M3BrK2tWZTFvQ0FLTVNyZ0FBb2lGb0FRREExQ3djZHBHNFI2dHhZRHpDRlFBUUZlRUtBQ0JLYjlicnRjNEFBREFKWmQwZWhoQisxUzBTOTB0VEZhYTJ3QTRKVndCQVZJUXJBSURvbVdnQkFNQ1VPSHdtZGJkQ0ZyQWJ3aFVBRUJYaENnQmdVZ1F0QUFDWWhMSnVUMElJNzNTTHhDMDBHTFpIdUFJQW9uTGZCeXY2UUwxd0JRQXdOWUlXQUFCTXhZVk9rYmpycGlvK2FUSnNsbkFGQUVUbGM3amlvcW1LVG1zQWdLa1N0QUFBSUhwbDNmYTMvQTkwaXNTWlpnRWJJbHdCQUZFUnJnQUFraU5vQVFCQTFNcTYzUThockhTSnhKMFpsd3l2VTlidFVRaGhQbndKVndEQXVJUXJBSUNrQ1ZvQUFCQzdsUU16RXZmWTc2WFdaSGk1SitHS1U1T1BBR0Iwd2hVQVFEYmVyTmRyM1FZQUlFcGwzUjZHRUg3VkhSTDNVMU1WRjVvTXp5TmNBUUJSRWE0QUFMSmtvZ1VBQURGeitFenFib1FzNFB1RUt3QWdLc0lWQUVEMkJDMEFBSWhTV2JmOVlkcXg3cEM0bFFiRGx3bFhBRUJVSG9jZ3ZIQUZBSkM5SUdnQkFFREV6aldIeEYwMlZmRkprK0UvaENzQUlDcVB3K1NLcTZZcXJyUUdBT0EvQkMwQUFJaE9XYmNyQjJ3azd0RTBDL2czNFFvQWlJcHdCUURBTXdoYUFBQVFsYkp1OTBNSUMxMGhjZWROVmR4cE1ya3E2L1p3K0Y0dlhBRUE0eE91QUFCNElVRUxBQUJpMDY4TTJkTVZFbmJmVklWcEZtUm5DRmVjRHRNcjN2b0VBTUNvaENzQUFGNUIwQUlBZ0dnTTQrUGY2d2lKTTdHRmJBaFhBRUJVaENzQUFEWkUwQUlBZ0ppYzZ3YUp1L0ZRbTlRSlZ3QkFWSVFyQUFDMlFOQUNBSUFvbEhYYkg4Z2Q2d2FKTTgyQ0pBbFhBRUJVaENzQUFMWk0wQUlBZ05HVmRic2ZRbGpwQkltN2JLcWkwMlJTSVZ3QkFORzVmaEt3ZU5BZUFJRHRFYlFBQUNBRy9TMy9BNTBnWVkrbVdaQUM0UW9BaUk1d0JRREFDQVF0QUFBWTFYQm81d0NhMUswOCtHYXFoQ3NBSURyQ0ZRQUFJeE8wQUFCZ2JQM0trRDFkSUdIM1RWV2NhekJUTXF4MCtoeXVPTlk4QUJpZGNBVUFRRVFFTFFBQUdFMVp0eWNoaFBjNlFPTG1Hc3dVUEFsWDlGL3ZOQTBBUmlkY0FRQVFLVUVMQUFERzVKWS9xYnRwcXVLVExoTXI0UW9BaUk1d0JRREFCQWhhQUFBd2lySnU3Zm9uQjZaWkVCM2hDZ0NJam5BRkFNREVDRm9BQUxCend5R2ZhUmFrN2tOVEZYZTZUQXlFS3dBZ09yZkQzMFRDRlFBQUV5Um9BUURBR0JZaGhEMlZKMkdQSVlTVkJqTW00UW9BaUU0ZnJyZ1l3aFVDdVFBQUUvWm12VjdySHdBQU8xUFc3V0VJNFZjVkozRS9OVlZ4b2Nuc21uQUZBRVJIdUFJQUlFRW1XZ0FBc0dzT24wbmRyWkFGdTFiVzdWeTRBZ0NpSVZ3QkFKQTRRUXNBQUhhbXJOdVRFTUt4aXBPNGhRYXpDMlhkbmo2WlhtRWRFd0NNUzdnQ0FDQWpnaFlBQU95U1cvNms3cnFwaWsrNnpMWUlWd0JBVklRckFBQXlKV2dCQU1CT2xIWGIzL0kvVUcwU1o1b0ZHeWRjQVFCUkVhNEFBRURRQWdDQTdTdnJkaitFc0ZKcUVuZm1ZVHViSWx3QkFGRzVEeUdjQzFjQUFQQ1pvQVVBQUx1d2NsQkk0ajQvZkljZkpsd0JBRkhwZjcrNzZxZFhORlhSYVEwQUFFKzlXYS9YQ2dJQXdOYVVkWHNVUXZpbkNwTzRuNXFxdU5Ca1hrcTRBZ0NpSWx3QkFNQ3ptR2dCQU1DMnVlVlA2bTZFTEhnSjRRb0FpSXB3QlFBQUx5Wm9BUURBMWd5SGljY3FUT0pXR3N6M0ROTjk1c09YY0FVQWpFdTRBZ0NBVnhHMEFBQmdtMHl6SUhXWFRWVjgwbVcrNUVtNG9nK2RIU2dTQUl4S3VBSUFnSTBSdEFBQVlDdkt1bDA1V0NSeGo2Wlo4RWZDRlFBUUZlRUtBQUMyUXRBQ0FJQ05LK3QyUDRTd1VGa1NkOTVVeFowbUkxd0JBRkVScmdBQVlPc0VMUUFBMklaK1pjaWV5cEt3KzZZcVRMUEltSEFGQUVUbDhVbTR3bG8zQUFDMlR0QUNBSUNOS3V2MkpJVHdYbFZKbklrdEdSS3VBSUNvZkE1WFhEVlZjYVUxQUFEc2txQUZBQUNiNXBZL3FidnhNRDhmd2hVQUVCWGhDZ0FBb2lCb0FRREF4cFIxMng5R0hxc29pVFBOSW5GbDNSNE9mUmF1QUlEeENWY0FBQkFkUVFzQUFEYWlyTnQ5MHl6SXdJZW1LanFOVHM4UXJqZ2RwbGU4emIwZUFEQXk0UW9BQUtJbWFBRUF3S1lzM1B3bWNZL0NSR2tScmdDQXFBaFhBQUF3R1cvVzY3VnVBUUR3S3NOaDVhK3FTT0orYWFyaVhKT25UYmdDQUtJaVhBRUF3Q1NaYUFFQXdDWTRmQ1oxOTBJVzB5VmNBUURSdVE0aFhBaFhBQUF3VllJV0FBQzhTbG0zSnlHRWQ2cEk0dVlhUEMzQ0ZRQVFuZXNuMHlzZXRBY0FnQ2tUdEFBQTRMWGM4aWQxTjAxVmZOTGwrQWxYQUVCMGhDc0FBRWlTb0FVQUFEK3NyRnVIbWVUQU5JdUlsWFc3L3lSY2NaeDdQUUFnQXNJVkFBQWtUOUFDQUlBZk1oeHVtbVpCNnM2YXFyalQ1Ymc4Q1ZlY1dsMEVBRkVRcmdBQUlDdUNGZ0FBL0toVkNHRlA5VWpZb3pCUlBJUXJBQ0E2d2hVQUFHVHJ6WHE5MW4wQUFGNmtyTnZERU1LdnFrYmlmbXFxNGtLVHh5TmNBUURSRWE0QUFDQjd3VVFMQUFCK2tNTm5VbmNyWkRFTzRRb0FpSTV3QlFBQS9JR2dCUUFBTDFMVzdVa0k0VmpWU054Q2czZEh1QUlBb25NN2hLdjdjTVdkOWdBQXdPOEpXZ0FBOEZKdStaTzY2NllxUHVueTlwVjFPeGV1QUlCb0NGY0FBTUF6Q1ZvQUFQQnNaZDMydC93UFZJeUVQWnBtc1YxbDNaNCttVjZ4bC9KN0JZQUpFSzRBQUlBZklHZ0JBTUN6REtQOVY2cEY0czRkTW15ZWNBVUFSRVc0QWdBQVhrblFBZ0NBNXpwM1FFcmk3b2ZQT1JzZ1hBRUFVUkd1QUFDQURSSzBBQURndThxNlBRb2h2RmNwRXJkcXF1SkJrMytjY0FVQVJFVzRBZ0FBdGtUUUFnQ0E1M0RMbjlUZE5GVnhvY3N2SjF3QkFGRVJyZ0FBZ0IwUXRBQUE0SnVHUTlSalZTSnhLdzErUHVFS0FJaEt2LzdzcWc5SEMxY0FBTUJ1Q0ZvQUFQQTlwbG1RdXN1bUtqN3A4cmNOSzRUbXc1ZHdCUUNNNjNPNDRxS3BpazR2QUFCZ3R3UXRBQUQ0cXJKdSsxditCeXBFd2g1RENBc04vckluNFlwVDN3c0FZSFRDRlFBQUVBbEJDd0FBdnFpczIwTUgwR1NnSDdIOW9OSC9JVndCQUZFUnJnQUFnQWdKV2dBQThEVXI2d0ZJM0gxVEZTdE5GcTRBZ01nSVZ3QUFRT1FFTFFBQStKT3liazlDQ085VmhzUmxQYkZGdUFJQW9pSmNBUUFBRXlKb0FRREFsN2psVCtwdW1xcTR5cTNMd2hVQUVCWGhDZ0FBbUNoQkN3QUFmcWVzMi80UTlsaFZTTnc4bHdhWGRYczRUTzhRcmdDQThUMzJ3UXJoQ2dBQW1EWkJDd0FBZmxQVzdYNEk0VnhGU055SHBpcnVVbjZMUTdqaWRBaVV2STNnSlFGQXpoNkh5UlZYT1U3VUFnQ0FGQWxhQUFEd1ZIL3JmVTlGU05oanFxdHhoQ3NBSUNyQ0ZRQUFrTEEzNi9WYWZ3RUErSHhJKzZ0S2tMaGZtcXBJWm1xTGNBVUFSRVc0QWdBQU1tR2lCUUFBbjFrWlF1cnVVd2haQ0ZjQVFGU0VLd0FBSUVPQ0ZnQUE5QWUzSnlHRWR5cVcxSVZvQUFBZ0FFbEVRVlJCNHVaVGZYdkNGUUFRRmVFS0FBREluS0FGQUFEQk5Bc3ljTjFVeGFjcHZVM2hDZ0NJaW5BRkFBRHdHMEVMQUlETWxYVzdjSWhMQmhaVGVJdGwzZTRQNFlwVFUyWUFZSFRDRlFBQXdCY0pXZ0FBWkd3NDFGMzVESkM0czZZcTdtSjlpOElWQUJDZDZ5RmNjYUUxQUFEQWx3aGFBQURrclE5WjdPVmVCSkwyR09OcUhPRUtBSWpPOVpQcEZRL2FBd0FBZk11YjlYcXRRQUFBR1NycjlqQ0U4S3ZlazdpZllybU5LbHdCQU5FUnJnQUFBSDZJaVJZQUFQa3lDcG5VM1k0ZHNoQ3VBSURvQ0ZjQUFBQ3ZKbWdCQUpDaHNtNVBRZ2pIZWsvaUZtTzhQZUVLQUlpT2NBVUFBTEJSZ2hZQUFIa3l6WUxVWFRaVjhXbVg3N0dzMno1WU1SZXVBSUFvQ0ZjQUFBQmJJMmdCQUpDWnNtNVhJWVFEZlNkaGp5R0UxUzdlM2hDdStQeTE1ME1GQUtNU3JnQUFBSFpDMEFJQUlDUERTb05SMWluQURwMDNWWEczcmYrY2NBVUFST1cyLzlrdlhBRUFBT3lTb0FVQVFGN09IUXlUdVB2aGM3NVJ3aFVBRUpYYllSWGUxVGJEbFFBQUFGOGphQUVBa0lteWJvOUNDTy8xbThTdE5uV2JWYmdDQUtJaVhBRUFBRVJEMEFJQUlCOGJ2K1VQa2JscHF1TGlOUzlKdUFJQW9pSmNBUUFBUkVuUUFnQWdBOFBoOGJGZWs3akZqN3c5NFFvQWlJcHdCUUFBRUQxQkN3Q0F4SlYxdTIrYUJSbTRiS3FpZSs3YkhGYnB6SWR3eFlFUENBQ01TcmdDQUFDWUZFRUxBSUQwTFJ3a2s3akg1MHl6RUs0QWdLZ0lWd0FBQUpNbGFBRUFrTEN5Ymc5L2RKMENUTWg1VXhVUFgzcTV3aFVBRUJYaENnQUFJQW1DRmdBQWFWdUZFUGIwbUlUZE4xV3hldnIyaENzQUlDcjNmYkNpRDFpOFpNMFhBQUJBekFRdEFBQVNWZGJ0U1FqaHZmNlN1RDVRSVZ3QkFIRVJyZ0FBQUpMMlpyMWU2ekFBUUlMS3V1MGZhci9WV3hMMitSQkh1QUlBeGlkY0FRQUFaRVBRQWdBZ1FXWGQ5amY3LzY2M0FBQnNrWEFGQUFDUUpVRUxBSURFbEhXN0gwSzRDeUhzNlMwQUFCc21YQUVBQUdUdnYzSXZBQUJBZ2haQ0ZnQUFiSkJ3QlFBQXdCTW1XZ0FBSktTczI4TVF3cTk2Q2dEQUt3bFhBQUFBZklXSkZnQUFhVG5YVHdBQWZ0RGprM0RGSjBVRUFBRDRNaE10QUFBU1VkYnRTUWpoSC9vSkFNQUxmQTVYWERWVmNhVndBQUFBMzJlaUJRQkFPaTcwRWdDQVp4Q3VBQUFBZUFWQkN3Q0FCSlIxdXdnaEhPZ2xBQUJmSVZ3QkFBQ3dJVmFIQUFCTVhGbTMreUdFdXhEQ25sNENBUENFY0FVQUFNQVdtR2dCQURCOUt5RUxBQUFHd2hVQUFBQmJacUlGQU1DRWxYVjdHRUw0VlE4QkFMSW1YQUVBQUxCREpsb0FBRXpiaGY0QkFHVHJVcmdDQUFCZzl3UXRBQUFtcXF6YjB4RENzZjRCQUdUbCtzbjBpZ2V0QndBQTJEMUJDd0NBNlRyWE93Q0FMQWhYQUFBQVJFVFFBZ0JnZ3NxNlhZVVFEdlFPQUNCWndoVUFBQUNSZXJOZXIvVUdBR0JDeXJyZER5SGNoUkQyOUEwQUlDbkNGUUFBQUJOZ29nVUF3UFNjQzFrQUFDUkR1QUlBQUdCaVRMUUFBSmlRc202UFFnai8xRE1BZ0VrVHJnQUFBSmd3RXkwQUFLYmxYTDhBQUNaSnVBSUFBQ0FSZ2hZQUFCTlIxdTA4aEhDc1h3QUFrM0ViUXJnWXdoVjMyZ1lBQUpBR3EwTUFBQ2Fnck52OUVFSVhRampRTHdDQXFBbFhBQUFBSk01RUN3Q0FhVmdJV1FBQVJFdTRBZ0FBSUNNbVdnQUFSSzZzMjhOaG1zV2VYZ0VBUkVPNEFnQUFJRk1tV2dBQXhPOWN5QUlBSUFyQ0ZRQUFBSmhvQVFBUXM3SnVUMElJLzlBa0FJRFJDRmNBQUFEd095WmFBQURFN1Z4L0FBQjJUcmdDQUFDQXJ4SzBBQUNJVkZtMzh4RENXLzBCQU5pSit5RmNjU0ZjQVFBQXdMZFlIUUlBRUtHeWJ2ZERDUDBEL2ozOUFRRFltajVjY1RXRUt6cGxCZ0FBNERsTXRBQUFpTk5DeUFJQVlDdUVLd0FBQUhnVkV5MEFBQ0pUMXUxaENPRlhmUUVBMkJqaENnQUFBRGJHUkFzQWdQaGM2QWtBd0tzSlZ3QUFBTEFWZ2hZQUFCRXA2L1lraEhDc0p3QUFQMFM0QWdBQWdLMFR0QUFBaUl0cEZnQUFMeU5jQVFBQXdFNEpXZ0FBUktLczIwVUk0VUEvQUFDKzYzRUlxQXBYQUFBQXNITnYxdXUxcWdNQWpLeXMyLzBRd2wwSVlVOHZBQUMrNkhHWVhISFZWTVdWRWdFQUFEQVdFeTBBQU9Ld0VySUFBUGdUNFFvQUFBQ2lZNklGQU1ESXlybzlDaUg4VXg5STNGa0k0VENFY0NwVUJNQjNDRmNBQUFBUU5STXRBQURHZDY0SEpPNm1xWXJWNTdkWTF1M3BFTGdRdWdEZ1MvcWZEUThoaEUrcUF3QUFRSXhNdEFBQUdORnc0UHovOUlERS9iV3BpdTZQYjdHczIvMG5nWXQzUGdRQS9NRjlDR0hlVklYQUJRQUFBRkVSdEFBQUdGRlp0M2NoaEFNOUlHR1hUVlhNdi9mMmhDNEErSVlQSVlSVlV4VVBpZ1FBQUVBTUJDMEFBRVpTMW0yL1NtR3AvaVNzMzdGLytOS0RzU2VoaTBVSTRhMFBDQUFoaE50aHVzV2ZKaVFCQUFEQXJnbGFBQUNNb0t6Ynd4QkNOK3dnaDFTZE5WV3hlczE3Ry82dDlLR0x1ZEFGQUp2NDJRSUFBQUN2SldnQkFEQ0NzbTR2UWdqdjFaNkUzVGRWY2JqSnR5ZDBBY0RBZEFzQUFBQkdKV2dCQUxCalpkMmVoQkQrb2U0azdtOU5WWHphMWxzczYvWm9DRnowd1lzREh5YUE3UFRycVZaTlZaeHJQUUFBQUxzbWFBRUFzR05sM2ZhSHo4ZnFUc0p1bXFvNDJkWGJFN29BeU5yTk1OM2lMdmRDQUFBQXNEdUNGZ0FBTzFUV2JYOFkvSGMxSjNGL0dldkFhd2hkTEliUXhaNFBHa0FXSG9ld3haVjJBd0FBc0F1Q0ZnQUFPMUxXN1g0SW9YUGpuc1I5YUtwaUVjTmJMT3YyZEFoY0NGMEE1T0Y2Q0Z3ODZEY0FBQURiSkdnQkFMQWpaZDJ1UWdoTDlTWmgvWTNpd3hnUHVJUXVBTEp4UDRRdFBtazVBQUFBMnlKb0FRQ3dBMlhkSG9ZUWZsVnJFdmRMVXhYbk1iL0ZZYkxNNThERnV3aGVFZ0RiY2RaVXhVcHRBUUFBMkFaQkN3Q0FIU2pyOXNxaExvbTdiYXJpYUVwdlVlZ0NJSG0zL2ZmNHBpcnV0Qm9BQUlCTkVyUUFBTml5c201UFFnai9VR2NTOTdjcGoybC9FcnFZaHhDT0kzaEpBR3hHdjlacTBWVEZoWG9DQUFDd0tZSVdBQUJiVnRadEYwSjRxODRrN0xxcGl0TlUzdDZ3NnVkejZNSy9YWUEwWFBmZjE1dXFlTkJQQUFBQVhrdlFBZ0JnaThxNjdROXEvNjdHSk80dnFZNWxGN29BU01yOXNFcWswMVlBQUFCZVE5QUNBR0JMaGxVRS9lSHpuaHFUc0xPbUtsWTVOSGdJWFN5RzRNVkJCQzhKZ0IrVHpjOHVBQUFBdGtQUUFnQmdTOHE2UFE4aC9LeStKS3pmZTMrWTR4ajJzbTZQaGlrWFFoY0EwM1F6VExld1NnUUFBSUFYRTdRQUFOaUM0ZWI3cjJwTDRuNXFxdUlpOXlZL0NWM01UYkFCbUpUSElXenhTZHNBQUFCNENVRUxBSUF0S091MmYyQi9yTFlrN0thcGloTU4vcjJ5YmsrSEtSZW5RaGNBazJHVkNBQUFBQzhpYUFFQXNHRmwzZmFIei85UVZ4TDNOemVBdjAzb0FtQlNyQklCQUFEZzJRUXRBQUEyckt6YnV4RENnYnFTc011bUt1WWEvSHhsM2M2SHdNVzdxYnhtZ0F4WkpRSUFBTUN6Q0ZvQUFHeFFXYmVMRU1ML3FDa0o2dytoanBxcXVOUGtseXZyZHYvSmxBdWhDNEE0V1NVQ0FBREFOd2xhQUFCc3lIQ0FlbWRGQUlseitMUWhUMElYL2JTTDR5VGVGRUE2ckJJQkFBRGdxd1F0Z1AvUDN0MGsxWEdrYlFOT085NDVmQ3NBcjBCNEJhQkpUb1dqSW5JcXRBTGpGZmhvQlkxV1lKaFdSTVlyVFhOaVdFSERDZ3c3RUN2Z2kvS2Jkc3R1U1M1K0RxY3E2N29pdklCK2JocU9UdDc1SkFCUHBPdkxhUWpodFhuU3NKdWM0cTZBbjE3WGw5MVBTaGN2V3Z2ZkJ6QlRuaElCQUFEZ3N4UXRBQUNlUU5lWHZSREN2ODJTeHYyUVUzd3Y1UFZTdWdDWW5KOXlpaWRpQVFBQTRBK0tGZ0FBVDZEcnk3blYvelR1SXFkNElPVG5WVXNYeDdWNHNiT2svKzBBRS9OaEtNQjVTZ1FBQUlDZ2FBRUE4SGhkWDRZRDBQODFTaHIzZlU3eFVzaWJVemZuSENsZEFHek1UWDFLeE45REFBQ0FoVk8wQUFCNGhLNHYyeUdFUzRlZU5PNHNwM2drNU9uNHBIUXgvTGUxOUhrQVBLUGJZZE5RVHZIVTBBRUFBSlpMMFFJQTRCRzZ2cXhDQ0QrYklRMGJEcFIyclVxZnJycFY1NC8vbEM0QW5zZTduT0t4V1FNQUFDeVRvZ1VBd0FOMWZkbXQyeXdjYk5LeW4zS0tKeEtlQjZVTGdHZDFGVUk0VUVZRUFBQllIa1VMQUlBSDZ2b3lySXgrYlg0MDdDYW51Q3ZnZWVyNmNsUUxGNitXUGd1QU5icXRaWXRMUXdZQUFGZ09SUXNBZ0FmbytuSVFRdmpWN0dqY3k1eml1WkRucmV2TDlpZGJMcFF1QU5ialRVN3gxR3dCQUFDV1FkRUNBT0FCdXI0TWg4Lzdaa2ZETG5LS0J3SnV5eWVsaXlPL3d3Q2UzRmxPOGNoWUFRQUEycWRvQVFCd1QzVWQveS9tUnVPK3l5bGVDN2xkWFY5MlB5bGR2Rmo2UEFDZXlGVjlTdVNqZ1FJQUFMUkwwUUlBNEI3cWJmRGg4SG5MM0dqWXU1emlzWUNYUStrQzRFbmQxckxGcGJFQ0FBQzBTZEVDQU9BZXVyNnNRZ2cvbXhrTkd3NkhkdDNFWGE1YXVqaXV4WXVkcGM4RDRJR0d2NmZIT2NWVEF3UUFBR2lQb2dVQXdFajE4UEUzODZKeGJ4d0s4WWV1TDN0MXk0WFNCY0REMkJJRkFBRFFJRVVMQUlDUnVyNjhEeUc4TWk4YWRwVlQzQk13bi9OSjZlTEk4MGtBOS9KaCtOMXBXeFFBQUVBN0ZDMEFBRWJvK25JUVF2alZyR2pjeTV6aXVaRDVKMTFmRHV1V2kwT2xDNEJScm9iZm1UbkZhK01DQUFDWVAwVUxBSUFSdXI1Y2hoQmVtQlVOKzVCVFBCUXc5NlYwQVREYWJRamhJS2Q0YVdRQUFBRHpwbWdCQVBBUHVyNE03MnIveTV4bzJIRHdzK2VXTFkvVjllV29GaTQ4c3dUd1pXOXlpcWZtQXdBQU1GK0tGZ0FBWDlIMVpUdUVjTzJXTm8xN20xTmNDWm1uVW45M0hpcGRBSHpSdTV6aXNmRUFBQURNazZJRkFNQlhkSDA1Q1NIOGFFWTA3S1p1cy9nb1pOWkI2UUxnaXo2RUVJNzhEUVlBQUpnZlJRc0FnQy9vK3JJYlF2ak5mR2ljOWVVOG0vcDdkU2hjREUrTXZEQjVnSEFWUWpoUXRnQUFBSmdYUlFzQWdDL28rbkllUXRnM0h4cDJrVk04RURDYm9IUUI4S2ZiV3JhNE5CSUFBSUI1VUxRQUFQaU1yaS9ENGZPdlprUGpYdVlVejRYTXB0WFN4Vkg5YjBjZ3dBSU5aWXREZjVjQkFBRG1RZEVDQU9BenVyNWNPK3lqY1djNXhTTWhNelZkWC9acTRlTFE3MkZnZ1R6cEJRQUFNQU9LRmdBQWY5UDFaUlZDK05sY2FOaHdhM2JYZS9CTW5kSUZzRkR2Y29ySHdnY0FBSmd1UlFzQWdFOTBmZGtPSVF6YkxMYk1oWWE5elNtdUJNeWNkSDA1cklXTFE3K2pnUVd3ZVFvQUFHRENGQzBBQUQ3UjlXVlkxZnphVEdqWVRVNXhWOERNbWRJRnNCQlhJWVFERzZnQUFBQ21SOUVDQUtDcUsrci9iUjQwN29lYzRuc2gwNHBhdWhodWZiOFNLdEFnWlFzQUFJQUpVclFBQUtpNnZweUhFUGJOZzRaZDVCUVBCRXlMNnROUGYyeTVVTG9BV25KYnl4YVhVZ1VBQUpnR1JRc0FnUDg3b0J0dVEvOWlGalR1ZTRjMExJSFNCZEFnWlFzQUFJQUpVYlFBQUJhdkhzZ05YMXJ2TEgwV05PMWRUdkZZeEN4TjE1ZmRXcmdZQ25Vdi9BQUFNemFVTFk1emlxZENCQUFBMkN4RkN3Qmc4YnErckVJSVB5OTlEalJ0T0pqWjliNDdTNmQwQVRUaWpiSUZBQURBWmlsYUFBQ0xWZy9kaG0wV1cwdWZCVTM3S2FkNEltTDRqL3I3LzZqK1o2TVJNRGYrdGdNQUFHeVFvZ1VBc0doZFg0YmJnSytYUGdlYWRwTlQzQlV4ZkZuWGw3MWF1RGhVdWdCbTVDeW5lQ1F3QUFDQTU2ZG9BUUFzVnRlWGd4RENyMzRDYU56TG5PSzVrR0VjcFF0Z1pwUXRBQUFBTmtEUkFnQllySzR2bDk3b3AzRVhPY1VESWNQRGRIMDVySVdMUTA5TUFSTjJGa0k0emlsK0ZCSUFBTUR6VUxRQUFCYXA2OHR3OCs4WDZkTzQ3M0tLMTBLR3g2dC9ONGJDeFN2akJDYm9Lb1J3b0d3QkFBRHdQQlF0QUlERjZmcXlIVUs0ZGp1WnhyM05LYTZFREUrci9nMDVxdi9aaWdSTWliSUZBQURBTTFHMEFBQVdwK3ZMY1BqOHMrUnAyRzBJWWRkQkM2eFgxNWU5VDBvWHludkFGQ2hiQUFBQVBBTkZDd0JnVWJxKzdJWVFmcE02alh1VFV6d1ZNandmVDRzQUU2SnNBUUFBc0dhS0ZnREFvblI5ZWU4UWpNWmQ1UlQzaEF5YlVRdDlmMnk1MkJFRHNDRzN0V3h4S1FBQUFJQ25wMmdCQUN4RzE1ZURFTUt2RXFkeEwzT0s1MEtHemV2NmNsZ0xGd3Ard0NZb1d3QUFBS3lKb2dVQXNCaGRYNjdkTHFaeFp6bkZJeUhEdE5oeUFXeVFzZ1VBQU1BYUtGb0FBSXZROWVVNGhQQXZhZE93NFNCbEw2ZDRMV1NZcnE0dmZ4UXU5c1VFUEJObEN3QUFnQ2VtYUFFQU5LL3J5M1lJWVRoODNwSTJEWHViVTF3SkdPYWhicms0cnFVTGY1K0FkVk8yQUFBQWVFTGZHaVlBc0FBcmgxZzA3aWFFY0NKa21JOWgrMHhPY1NoYURJV0xOL1gveHdEck1ud1dQdS82c21mQ0FBQUFqMmVqQlFEUXRIcGorRGNwMDdnM09jVlRJY084ZFgwNXFGc3VYb2tTV0JPYkxRQUFBSjZBb2dVQTBMU3VMK2Zld2FkeEZ6bkZBeUZET3p3ckFxeVpzZ1VBQU1BaktWb0FBTTNxK25JWVF2aGZDZE80N3gyVVFKdTZ2bXlIRUE3ckUxZzdZZ2Fla0xJRkFBREFJeWhhQUFETjZ2cHk3V0NLeHAzbEZJK0VETzJyNWNGalc1cUFKNlJzQVFBQThFRGZHaHdBMEtLdUwyNy8wcnJiZXVnS0xFQk84WDE5SnVqN29XUWxjK0FKREU4VG5YZDkyVE5NQUFDQSs3SFJBZ0JvVGwyMWZ1MWRleHIzTnFlNEVqSXNVOWVYM1JEQ1VTMWMrWHNIUEliTkZnQUFBUGVrYUFFQU5LZnJ5MmtJNGJWa2FkaE5UbkZYd0VBdEYvNVJ1TERKQ1hnb1pRc0FBSUI3VUxRQUFKclM5V1ZZcS82clZHbmNEOE16QWtJR1B0WDE1WS9DeFF1REFSNUEyUUlBQUdBa1JRc0FvQ2xkWDg1RENQdFNwV0VYT2NVREFRTmZVa3VISzM4UGdRZFF0Z0FBQUJoQjBRSUFhRWE5eWZ1TFJHbmNkem5GYXlFRC82UVdMbzQ4cHdYYzAxQzIyUE41QXdBQTRNc1VMUUNBSnRRMzZpKzlUMC9qM3VVVWo0VU0zRWZYbDkyNjRVTGhBaGpycW02MitHaGlBQUFBLysxYk13RUFHbkdzWkVIamJ1dEJLY0M5RExmU2M0ckRab3Z2UWdobjlmY0p3TmU4Q0NHYzF6SXpBQUFBZjJPakJRQXdlL1dtN3JETllrdWFOT3lubk9LSmdJSEhxZ2VueC9VL2Z6dUJyN0haQWdBQTRETVVMUUNBMmV2NjhqNkU4RXFTTk93cXA3Z25ZT0FwS1Z3QUl5bGJBQUFBL0kyblF3Q0FXZXY2Y3FCa3dRSWNDeGw0YXNPaGFVNXhlSkpvMkF6MTFwTWl3QmNNejRqWXFnVUFBUEFKR3kwQWdGbnIrbkpadi95RlZuM0lLUjVLRjFnM0d5NkFmM0NXVXp3eUpBQUFBRVVMQUdER3VyNE1YL1QrSWtNYTkxMU84VnJJd0hPcGhZdWg0RFZzdTlneGVPQVR5aFlBQU1EaUJVVUxBR0N1NmlIUXRSdTNOTzV0WGVzUHNCRzExS2h3QVh6cXA1eWlwMFFBQUlCRiszYnBBd0FBWnN0YWMxcDM2ejEwWU5OeWlxYzV4ZDJoK0ZWL0x3SDhxNWF3QUFBQUZzdEdDd0JnZHJxK0RBYyt2MG1PeHIwWkRqaUZERXhGM1NaMXJPd0lWRDZyQUFBQWk2Vm9BUURNVHRlWDh4REN2dVJvMkVWTzhVREF3QlFwWEFEVnNPWG1JS2Q0YVNBQUFNRFNLRm9BQUxQUzlXVTRmUDVWYWpUdVpVN3hYTWpBbENsY0FNb1dBQURBVWlsYUFBQ3owdlhsT29Td0l6VWFkcFpUOU80NU1Cc0tGN0I0TnlHRXZaeml4NlVQQWdBQVdJNXZaUTBBekVYWGwyTWxDeG8zM0FwZENSbVlrK0Z3TmFjNC9PN2FHOHBpd29QRkdUNmZuOWZTRlFBQXdDSW9XZ0FBczFDL3VIVUFUZXRPY29yWFVnYm1hUGo5VlRmeWZLZHdBWXZ6SW9Sd0tuWUFBR0FwRkMwQWdMbFlXVWRPNDRhMTJ5ZENCdWJ1YjRXTEM0SENZcnpxK3FKc0FRQUFMTUkzZDNkM2tnWUFKcTNyeTdDSy9OOVNvbkUvNUJUZkN4bG9UZGVYZzFxWTNCY3VMTUpQT1VYbFVRQUFvR21LRmdEQTVIVjlPWGM0UStNdWNvb0hRZ1phMXZYbHNHN3UyUkUwTk85TlR0RjJDd0FBb0ZtS0ZnREFwTlZEbWYrVkVvMzdQcWQ0S1dSZ0NicStITlVORndvWDBLN2JFTUtCenpjQUFFQ3J2cFVzQURCeDFnN1R1ak9IRU1DUzFGdnV3N05nYit0aExOQ2VyUkRDZWRlWFhka0NBQUF0VXJRQUFDYXI2NHZicnJSdU9HQThsakt3TkRuRmp6bkY0ZS84Y0FoNzVnY0FtalNVTGQ1M2Zka1dMd0FBMEJwRkN3QmdrdXJ0TndmUXRHNDFIRFpLR1ZpcVdyZ1luaEw1TG9SdzRRY0Jtdk1paEhBcVZnQUFvRFdLRmdEQVZLM3FMVGhvMVUxTzBkTTRBUDlYdUxqT0tSNkVFRjZHRUs3TUJKcnlxdXVMenp3QUFFQlRGQzBBZ01ucCtqSWN0THlXREkwN0VqREFYK1VVejNPS2V5R0VOME1oelhpZ0dUOTJmZkhaQndBQWFJYWlCUUF3UlN1cDBMaUw0VEJSeUFDZmwxTWNuaG9ZQ2hkdlF3aTN4Z1JOK0tYcnk1NG9BUUNBRm54emQzY25TQUJnTXVwTnQxOGtRdU8rRzlia0N4bmduM1Y5MmEwbFROdXVZUDZHNHRSdVR2R2pMQUVBZ0RtejBRSUFtSXl1TDl1MldiQUE3NVFzQU1ZYmZtZm1GSWNpNXN0aEk1RFJ3YXh0aFJCczlRSUFBR1pQMFFJQW1KTGpFTUtPUkdqWXJUSVJ3TU1NVHk3bEZBOUNDRzg4SndLejlxTHJ5NmtJQVFDQU9mTjBDQUF3Q1hVdCtHL1NvSEZ2Y29vT0ZnQWVxVzdCR2dxYVA1c2x6TlpQT2NVVDhRRUFBSE9rYUFFQVRFTFhsL2NoaEZmU29HRlhPY1U5QVFNOG5WclVIQXBzKzhZS3MvUnkyRllqT2dBQVlHNFVMUUNBamV2Nk1xd0IvMVVTTk01QkFzQ2ExTThTcDU0Z2c5a1puZ0hhelNsK0ZCMEFBREFuMzBvTEFKZ0FLNE5wM1FjbEM0RDFHWDdINWhTSDdSWnY2OEV0TUE5YklRU2ZrUUFBZ05sUnRBQUFOcXJyeTFFSTRZVVVhTnl4Z0FIV0w2ZTRDaUVNenpSOU1HNllqUmRkWDA3RkJRQUF6SW1uUXdDQWplbjZzaDFDdUs0MzJhQlZiK3ZCSHdEUHlITWlNRHR2Y29vS0Z3QUF3Q3pZYUFFQWJOSkt5WUxHM1hnYUIyQXovdmFjQ0RCOUoxMWY5dVFFQUFETWdZMFdBTUJHZEgwWkRqNStNMzBhNTJZbXdBVFV6eDNENytOOWVjQ2tYWVVRRG5LS0g4VUVBQUJNbVkwV0FNQ21PSHltZFJkS0ZnRFRrRk84emlrT1Q0bjhVTGNOQWRQMHdyOFRBQUNBT1ZDMEFBQ2VYWDB6M1kxU1dyZVNNTUMwNUJUZmh4Q0dwd25laVFZbTYxWFhsMlB4QUFBQVUrYnBFQURnMlhWOXVRNGg3Smc4RFR2TEtSNEpHR0M2dXI3czFadnpMOFFFay9SOVR2RlNOQUFBd0JUWmFBRUFQS3V1THlzbEN4cDNhNXNGd1BRTkI3ZzV4YUZzOFZQOTNRMU15L3V1TDlzeUFRQUFwa2pSQWdCNE52V0xVbXVBYWQxSlR2RmF5Z0R6a0ZNOHFjK0pYSWdNSm1XbmJwMEJBQUNZSEUrSEFBRFBwdXZMOEVYcGF4T25ZVGM1eFYwQkE4eFQxNWZEZXJDN0pVS1lqRGM1UllVTEFBQmdVbXkwQUFDZVJYMEhYY21DMXRuWUFqQmpPY1gzSVlTaE1IY21SNWlNay9wdkNRQUFnTW13MFFJQWVCWmRYODVEQ1B1bVRjTXVjb29IQWdab1E5ZVhnN3JkWWtla3NIRlhJWVNEbk9KSFVRQUFBRk5nb3dVQXNIWjFEYmVTQmEyenpRS2dJVG5Gb1NRNjNLSi9KMWZZdUJjaGhKVVlBQUNBcWJEUkFnQllxNjR2MnlHRVM3ZEJhZHk3bktLaUJVQ2piTGVBeWZpaFB2RURBQUN3VVRaYUFBRHJkdXhRZ3NiZHVtRUowTFpQdGx1OEZUVnMxR2t0Y2dNQUFHeVVvZ1VBc0RaZFgzWTlwOEFDckx3WER0Qys0WGQ5VG5FbzFuMGZRcmdTT1d6RVZnakJSZ3NBQUdEakZDMEFnSFZhMVM5RG9WVTNPY1VUNlFJc1IwN3hNcWRvdXdWc3puN1hGMlZ1QUFCZ294UXRBSUMxcUcrWnZ6WmRHbmNrWUlCbHN0MENObXJWOVdWUEJBQUF3S1lvV2dBQTY3SXlXUnAzVWQvc0IyQ2hodTBXSVlTaFhQck96d0E4cTJGcjNxbVJBd0FBbTZKb0FRQTh1YTR2d3kzL2ZaT2xjYlpaQURDVUxUN21GSWRuREY0T1QwcVpDRHliRjExZmxMc0JBSUNOK09idTdzN2tBWUFuMC9WbE80UndYVytaUWF2ZTFwWHhBUENuK2psbytQdndvNm5Bcy9tK2JwY0JBQUI0TmpaYUFBQlA3VmpKZ3NiZGhoQk9oQXpBMzMyeTNlS0grdmNDV0Q5UGlBQUFBTTlPMFFJQWVESmRYM1pEQ0QrYktJMDdIZzdTaEF6QWwrUVUzNGNRaHM5Rkh3d0oxbTU0UWtRSkZnQUFlRmFlRGdFQW5relhsK0ZRNFpXSjByQ3JuT0tlZ0FFWXErdkxjWDFPeE1ZdldLK1hPY1Z6TXdZQUFKNkRqUllBd0pQbytuS2daTUVDSEFzWmdQdklLUTQzN1llUzNwWEJ3VnFkZG4zWk5tSUFBT0E1S0ZvQUFFL0Z1bDVhOThFdFNRQWVJcWQ0WFRjaXZUVkFXSnNkcFZnQUFPQzVLRm9BQUk5V1YySy9NRWthZHV1TGV3QWVLNmM0UENIeU1vUndZNWl3Rmo5M2ZmSE1Hd0FBc0hhS0ZnREFvOVQxdkN0VHBIRW53MjFrSVFQd1dIVTcwbkFRL01Fd1lTMU9qUlVBQUZnM1JRc0E0TEdHa3NXV0tkS3dHMC9qQVBDVWNvb2ZjNHFISVlRM2RXc1M4SFJlZEgxUkJBY0FBTmJxbTd1N094TUdBQjZrNjh0dUNPRTMwNk54YjNLS2JrWUNzQmIxbVlOVHo3REJrL3MrcDNocHJBQUF3RHJZYUFFQVBJYkRaMXAzb1dRQndEclZnK0NERU1JN2c0WW5aU01aQUFDd05vb1dBTUNEZEgwWjFsM3ZteDZOT3hZd0FPdFdueElaL3ViODRDa1JlREw3WFY5OGxnTUFBTlpDMFFJQWVDZzN4R2pkbVhYVEFEeW5uT0w3RU1Md2xNaVZ3Y09UV0hWOTJUWktBQURncVNsYUFBRDMxdlZsRlVMWU1Ua2FkbXViQlFDYmtGTzh6aW51ZVVvRW5zU1c1dzRCQUlCMVVMUUFBTzZsM2doekFFM3JUb1kxN2xJR1lGTThKUUpQNWxYWGx3UGpCQUFBbnBLaUJRQndYeWYxWmhpMDZpYW51Skl1QUp0V254STU4SlFJUE5xcEowUUFBSUNucEdnQkFJelc5V1ZZWS8zYXhHaWNqUzBBVEVaTzhiS1dMYzZrQWcrMjR6TWVBQUR3bEw2NXU3c3pVQUJnbEs0djV5R0VmZE9pWVJjNVJhdWxBWmlrcmk5SElZUmZwQU1QOWwxTzhkcjRBQUNBeDdMUkFnQVlwWDZ4cjJSQjY0NGtETUJVNVJSUFF3amZEODljQ1FrZTVOVFlBQUNBcDZCb0FRRDhvL3FlOGNxa2FOdzdOeHdCbUxyNmxNanduTnVGc09EZTlydStIQm9iQUFEd1dJb1dBTUFZeC9WZFkyalZyVElSQUhPUlUveFluN3A2SnpTNHQ1TmFKQWNBQUhnd1JRc0E0S3U2dnV6V29nVzBiRFVjV2trWWdEbkpLUTZmMGQ3VXdpQXd6bzUvM3dBQUFJK2xhQUVBL0pQaGx2K1dLZEd3bTV6aWlZQUJtS09jNG1rSVlkaHVjU05BR08yNEZzb0JBQUFlUk5FQ0FQaWlyaS9EbC9hdlRZakdIUWtZZ0RuTEtWNkdFUFpDQ0JlQ2hGRzJQQnNIQUFBOGhxSUZBUEExYnZuVHVnODV4WE1wQXpCM3d4TllPY1doSlB0T21EREs2MW9zQndBQXVEZEZDd0RnczdxK0RMZjhYNWdPamZNK053Qk55U2tPZjl2ZVNCVkdzZFVDQUFCNEVFVUxBT0MvZEgzWnRzMkNCWGliVTd3V05BQ3R5U21laGhDK0R5SGNDaGUrYXI4V3pBRUFBTzVGMFFJQStKemorbTR4dE9wV21RaUFsdVVVTDBNSWV5R0VLMEhEVjYxcTBSd0FBR0EwUlFzQTRDKzZ2dXlHRUg0MkZScDNQTHhsTDJRQVdsWTNOeDJFRUQ0SUdyNW94M055QUFEQWZYMXpkM2RuYUFEQW43cSt2QThodkRJUkduYVZVOXdUTUFCTDB2VmwyT1QwbzlEaHM0WnRaN3VLdUFBQXdGZzJXZ0FBZityNmNxQmt3UUs0c1FqQTR1UVVoNzkvYnlRUG43WGxXVGtBQU9BK0ZDMEFnRStkbWdhTk84c3BuZ3NaZ0NYS0tRNmY5VjdXMi92QVg3MnV6eWdDQUFEOEkwVUxBT0IzWFYrTzYvdkUwS3JoVUdrbFhRQ1dyQllPRDVRdDRMTjhWZ1FBQUVaUnRBQUFocExGdGk4VldZQ1RuT0sxb0FGWXVwemlaUWhodUxsL3RmUlp3Tis4cnM4cEFnQUFmSldpQlFBUWFzbGl5eVJvMkJaTkh0VUFBQmJVU1VSQlZJMTN0d0hnUDNLS0grdG1pd3RqZ2I5UVFBY0FBUDdSTjNkM2Q2WUVBQXZXOVdVdmhQQnZQd00wN29lYzRuc2hBOEIvNi9weU90emtOeHI0MDh2NnpBNEFBTUJuMldnQkFMamxUK3N1bEN3QTRNdHlpa2NoaEhkR0JIK3kxUUlBQVBncVJRc0FXTEN1TDRjaGhIMC9BelR1V01BQThIVTV4ZUh2NVJ0amd0L3QxMzhyQVFBQWZKYWlCUUFzbTIwV3RPNHNwM2dwWlFENFp6bkZVMlVMK0pOL0t3RUFBRitrYUFFQUM5WDFaVmlIdXlOL0duWnJtd1VBM0U4dFczeGYvNDdDa3UxMGZUbnlFd0FBQUh5T29nVUFMRkRYbDIwSDBDekFTVTd4bzZBQjRIN3FOcWdEWlFzSUt5TUFBQUErUjlFQ0FKWnBXSU83SlhzYWRwTlQ5TVU0QUR5UXNnWDh6bFlMQUFEZ3N4UXRBR0JodXI0TVg1aS9sanVOODRVNEFEeFNMVnZzaFJDdXpKSUZVOTRGQUFEK2k2SUZBQ3lQTHdwcDNVVk84VnpLQVBCNE9jWHJ1dGxDMllLbHN0VUNBQUQ0TDRvV0FMQWc5UXZDZlpuVE9GK0VBOEFUeWlsK1ZMWmc0WlRWQVFDQXYxQzBBSUNGNlBxeTdRdENGdUJkdlhrTEFEd2haUXNXYnRocWNiajBJUUFBQVAraGFBRUF5M0U4ZkVFb2J4cDJxMHdFQU91amJNSENIUzk5QUFBQXdIOG9XZ0RBQW5SOTJmWEZJQXV3cWdkQUFNQ2FLRnV3WVB0ZFh3NzhBQUFBQUVIUkFnQVc0eVNFc0NWdUduYVZVendSTUFDczN5ZGxpdy9HemNMWW5nWUFBUHhPMFFJQUdsZHZYYjJTTTQyenNRVUFudEZRdHNncEhvWVF6c3lkQmJIVkFnQUErSjJpQlFDMHp5MS9XdmNocDNndVpRQjRmam5GSTJVTEZ1Wkk0QUFBZ0tJRkFEU3M2OHZ3SmVBTEdkTTQyeXdBWUlOcTJlSkNCaXpFNjY0dnU4SUdBSUJsVTdRQWdFWjFmZG0yellJRmVKdFR2QlkwQUd6YzhJeklsUmhZaUpXZ0FRQmcyUlF0QUtCZHd5My9MZm5Tc0J0bElnQ1locHppeHhEQ2diSUZDL0c2RnRzQkFJQ0ZVclFBZ0FiVlZiWS95NWJHcmVxaERnQXdBY29XTEl6bjZ3QUFZTUVVTFFDZ1RhZHlwWEVYT1VVLzV3QXdNY29XTElpaUJRQUFMSmlpQlFBMHB1dkw4TVgydmx4cG5IZXhBV0NpbEMxWWlLMnVMMGZDQmdDQVpWSzBBSUQydU9WUDY4NXlpdWRTQm9EcCtxUnNjU01tR21hckJRQUFMSlNpQlFBMHBPdkw4RVhmamt4cDJLMXRGZ0F3RDdWc2NWai9ma09MWHRTTmdnQUF3TUlvV2dCQUk3cStiRHVBWmdGT2NvclhnZ2FBZWNncFh0Yk5Gc29XdE1yeklRQUFzRUNLRmdEUWpwUGhuV0I1MHJDYm5LSXlFUURNakxJRmpYdmQ5V1ZYeUFBQXNDeUtGZ0RRZ0s0dmU4TVhmTEtrY2Q3QUJvQ1pxbVVMZjh0cGxhMFdBQUN3TUlvV0FOQ0dFem5TdUl1YzRuc2hBOEI4NVJSUFF3aHZSRWlEbElnQUFHQmhGQzBBWU9hNnZoeUdFUGJsU09OOGVRMEFEYWhsaTdleXBERmJYVjlzdFFBQWdBVlJ0QUNBK2JQTmd0YWQxWFhqQUVBRGNvcXI0ZSs3TEdtTW9nVUFBQ3lJb2dVQXpGalhsK0ZMNmgwWjByQmIyeXdBb0QwNXhlRlEra0swTkdTLzY4dXVRQUVBWUJrVUxRQmdwdXFYZUE2Z2FkMHFwL2hSeWdEUXBPRUp2Q3ZSMGhEL1BnTUFnSVZRdEFDQStScTJXV3pKajRiZDVCUTlqUU1BamFwbHlzTzZ3UXBhNFBrUUFBQllDRVVMQUppaHJpOEhJWVRYc3FOeHZxZ0dnTWJsRks5RENBZktGalJpcSt1THo3QUFBTEFBaWhZQU1FOHJ1ZEc0aTV6aXVaQUJvSDA1eFV0UEx0Q1FRMkVDQUVEN0ZDMEFZR2JxRGFsOXVkRTROd0VCWUVGeWlxY2hoTGN5cHdHdnVyN3NDaElBQU5xbWFBRUFNOUwxWmRzMkN4YmdYVjBqRGdBc1NFNXgrSno3UWVZMHdGWUxBQUJvbktJRkFNekxzRko1UjJZMDdGYVpDQUFXYmRocWRiWDBJVEI3bnNJQkFJREdmWE4zZHlkakFKaUJ1bjcyTjFuUnVEZDFkVGdBc0ZEMWMrOWxDR0hMendBejluMU84VktBQUFEUUpoc3RBR0ErVG1SRjQ2NlVMQUNBK29TWXB4ZVlPMXN0QUFDZ1lZb1dBREFEWFY4T1FnaXZaRVhqZkJrTkFQd3VwM2dlUXZqSk5KZ3haU0VBQUdpWW9nVUF6SU50RnJUdVF6MVFBUUQ0WFU1eCtBejh3VFNZcWEydUw4b1dBQURRS0VVTEFKaTRyaS9ETGY4WGNxSmh0N1paQUFCZmNCUkN1REVjWmtyUkFnQUFHdlhOM2QyZGJBRmdvcnErYkljUWhqZXF0MlJFdzk3bUZGY0NCZ0ErcCt2TFhnamgzR2RpWnVyLzVSUS9DZzhBQU5waW93VUFUTnZLRjhvMDdzYlRPQURBMStRVUwyMi9Zc1pzdFFBQWdBWXBXZ0RBUkhWOTJRMGgvQ2dmR3JkeXd3OEErQ2M1eGRNUXdwbEJNVU9LRmdBQTBDQkZDd0NZcmxQWjBMaUxlbWdDQURER3NOWGl5cVNZbVZmMVNVZ0FBS0FoaWhZQU1FRmRYdzVDQ1B1eW9YRXJBUU1BWTlVdFdFY0d4Z3paYWdFQUFJMVJ0QUNBYVhMTG45YWQ1UlRQcFF3QTNFZE84VEtFOEpPaE1UT0tGZ0FBMEJoRkN3Q1ltSzR2d3kzL0hiblFzTnU2K2hzQTRONXlpaWZERTJRbXg0eDRQZ1FBQUJxamFBRUFFMUsvZkhNQVRldE82dXB2QUlDSE9xemxUWmdMV3kwQUFLQWhpaFlBTUMzRDdid3RtZEN3bTV6aVNzQUF3R1BVMHVhUklUSWppaFlBQU5BUVJRc0FtSWl1TDNzaGhOZnlvSEUydGdBQVR5S24rRDZFY0dhYXpNUXJRUUVBUURzVUxRQmdPazVrUWVNdTZvRUlBTUJUR1VxY042YkpISFI5c2RVQ0FBQWFvV2dCQUJOUXYzRGJsd1dOczgwQ0FIaFNuaEJoWmhRdEFBQ2dFWW9XQUxCaFhWKzJiYk5nQWQ3bEZDOEZEUUE4dFp6aStmQlp3MkNaQVVVTEFBQm9oS0lGQUd6ZWNNdC9SdzQwN0RhRXNCSXdBTEJHSzArSU1BTmJYVi8yQkFVQUFQT25hQUVBRzlUMVpkZHpDaXpBcXE3MUJnQllDMCtJTUNPMldnQUFRQU1VTFFCZ3M0YWJkMXN5b0dFM09VVlA0d0FBYTFlZkVEa3phU1pPMFFJQUFCcWdhQUVBRzlMMTVTQ0U4TnI4YVp5YnBRREFjenF1ejViQlZMMm9tdzBCQUlBWlU3UUFnTTFabVQyTisxQnZsZ0lBUEl2NmhJaW4rWmk2QXdrQkFNQzhLVm9Bd0FaMGZSbHUrZStiUFkxenlBRUFQTHVjNG1rSTRjTGttVEJGQ3dBQW1EbEZDd0I0WmwxZnRrTUlKK1pPNDk3bUZLK0ZEQUJzaU9mTG1MSkQ2UUFBd0x3cFdnREE4eHR1K1crWk93MjdWU1lDQURhcEZqN2ZDb0dKMnVyNnNpY2NBQUNZTDBVTEFIaEdYVjkyUXdnL216bU5PNjd2b3dNQWJOSlEvTHlSQUJQbCtSQUFBSmd4UlFzQWVGNXUrZE82cS9vdU9nREFSdFhpNTdFVW1DalBod0FBd0l3cFdnREFNK242TXR4WWVtWGVOTTVoQmdBd0dUbkY5eUdFQzRrd1FmdENBUUNBK1ZLMEFJRG40NVkvclR2TEtaNUxHUUNZR0VWUUpxbVc4UUVBZ0JsU3RBQ0FaOUQxWmZoeWQ4ZXNhZGh0Q0dFbFlBQmdhbktLbDBNaFZEQk1rS0lGQUFETWxLSUZBS3haMTVkdEI5QXN3RWxPOFZyUUFNQkVyV294RktaRTBRSUFBR1pLMFFJQTFtLzRVbmZMbkduWXpWQzBFREFBTUZXMUVPcnpDbE96THhFQUFKZ25SUXNBV0tPdUw3c2hoQi9ObU1hdGNvb2ZoUXdBVE55SnJSWk1UZGNYV3kwQUFHQ0dGQzBBWUwxT3paZkdYZVFVL1p3REFKTlhpNkcyV2pBMWloWUFBREJEaWhZQXNDWmRYdzZ0Z21VQmpvVU1BTXhGVG5GVm56MkRxVkMwQUFDQUdWSzBBSUQxY1Z1TzFwM2xGQytsREFETXpFcGdUSWh5UGdBQXpKQ2lCUUNzUWRlWDRjdmJIYk9sWWJlMldRQUFjMVNmUGJQVmdzbm8rckluRFFBQW1CZEZDd0I0WWwxZnRoMUFzd0FuOVoxekFJQTVzdFdDS2ZGOENBQUF6SXlpQlFBOHZlSEprQzF6cFdFMzlYMXpBSUJac3RXQ2liSFJBZ0FBWmtiUkFnQ2VVRjM1K3RwTWFkeVJnQUdBQmlpT01oVTJXZ0FBd013b1dnREEwem94VHhwM2tWTThGeklBTUhkMXE4V3RJSm1BbmZvRUpRQUFNQk9LRmdEd1JMcStETGY4OTgyVHh0bG1BUUMwUkZHYXFmQjhDQUFBeklpaUJRQThnWHI3eU9waFd2Y3VwM2d0WlFDZ0lTZTJXakFSbmc4QkFJQVpVYlFBZ0tkeFBLeDdOVXNhZHF0TUJBQzBKcWY0TVlSd0tsZ213RVlMQUFDWUVVVUxBSGlrcmkrN3RXZ0JMVnZWZ3dnQWdOWjRQb1FwVUxRQUFJQVpVYlFBZ01jYmJ2bHZtU01OdThvcE9vQUFBSnBVbjBZN2t5NGJ0bE9mcEFRQUFHWkEwUUlBSHFIcnkvQ083bXN6cEhFMnRnQUFyZk44Q0ZOZ3F3VUFBTXlFb2dVQVBJNWIvclR1UTA3eFhNb0FRTXZxNTUwckliTmhCd0lBQUlCNVVMUUFnQWZxK25JVVFuaGhmalRPTmdzQVlDbVVxTm0wWFFrQUFNQThLRm9Bd0FQVXQzTjlFVXZyM3RZM3l3RUFsdUI5Q09GVzBteVFwME1BQUdBbUZDMEE0R0dHVy81YlprZkRicFdKQUlBbHlTbCtyR1VMMkJRYkV3RUFZQ1lVTFFEZ25ycStET3RjZnpZM0duZGNEeHNBQUpaRTBaU042dnBpcXdVQUFNeUFvZ1VBM04rcG1kRzRpNXlpbjNNQVlIRnlpcGNoaEJ2SnMwRzdoZzhBQU5PbmFBRUE5OUQxNVNDRXNHOW1ORzRsWUFCZ3dXeTFZSk5zdEFBQWdCbFF0QUNBKzNITG45YWQ1UlRQcFF3QUxOaDc0Yk5CaWhZQUFEQURpaFlBTUZMWGwrTVF3bzU1MGJCYjJ5d0FnS1hMS1Y2SEVENHNmUTVzaktkREFBQmdCaFF0QUdDRXJpL2JEcUJaZ0pONnNBQUFzSFMyV3JBcEwwd2VBQUNtVDlFQ0FNWVpTaFpiWmtYRGJyeEhEZ0R3SjBVTE5xYnJpNjBXQUFBd2NZb1dBUEFQdXI0TWIrVCthRTQwN2ppbitGSElBQUMvUHgveTBmTWhiSkNpQlFBQVRKeWlCUUQ4TTdmOGFkMUZUdEd0VFFDQXYvTDVpRTA1TUhrQUFKZzJSUXNBK0lxdUw0Y2hoSDB6b25ISEFnWUErQytLRm16S3Rza0RBTUMwS1ZvQXdOZlpaa0hyem5LS2wxSUdBUGlyK256SWxiR3dBWHVHRGdBQTA2Wm9BUUJmMFBWbEZVTFlNUjhhZG11YkJRREFWNTBhRHh0Z293VUFBRXpjL3dnSUFMN29Pb1R3MW5obzJHVzlxUWtBd09lOWQrZ05BQURBMzMxemQzZG5LQUFBQUFBQUFBQUFJM2c2QkFBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEFBQUFBQUFBQUFFWlN0QUFBQUFBQUFBQUFHRW5SQWdBQUFBQUFBQUJnSkVVTEFBQUFBQUFBQUlDUkZDMEErUC90Mm9FQUFBQUF3NkQ3VTE5aEFNVVJBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBVUd3N0VienhEVmFPOXBNQUFBQUFTVVZPUks1Q1lJST0nOwpmdW5jdGlvbiBhcHBseUJyYW5kaW5nKCkgewogIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJy5icmFuZC1sb2dvJykuZm9yRWFjaCgoaW1nKSA9PiB7IGltZy5zcmMgPSBMT0dPX0RBVEFfVVJJOyB9KTsKICBjb25zdCBmYXZpY29uID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2Zhdmljb25MaW5rJyk7CiAgaWYgKGZhdmljb24pIGZhdmljb24uaHJlZiA9IExPR09fREFUQV9VUkk7Cn0KCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBBcGkg4oCUIHRoaW4gZmV0Y2ggd3JhcHBlcnMgYXJvdW5kIHRoZSBSRVNUIEFQSS4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmNvbnN0IEFwaSA9ICgoKSA9PiB7CiAgYXN5bmMgZnVuY3Rpb24gcmVxdWVzdChwYXRoLCBvcHRpb25zKSB7CiAgICBjb25zdCByZXMgPSBhd2FpdCBmZXRjaChwYXRoLCBvcHRpb25zKTsKICAgIGxldCBib2R5OwogICAgdHJ5IHsKICAgICAgYm9keSA9IGF3YWl0IHJlcy5qc29uKCk7CiAgICB9IGNhdGNoIHsKICAgICAgYm9keSA9IG51bGw7CiAgICB9CiAgICBpZiAocmVzLnN0YXR1cyA9PT0gNDAxICYmICFwYXRoLnN0YXJ0c1dpdGgoJy9hcGkvYXV0aC8nKSkgewogICAgICB3aW5kb3cuZGlzcGF0Y2hFdmVudChuZXcgQ3VzdG9tRXZlbnQoJ2xyczpzaWduZWQtb3V0JykpOwogICAgfQogICAgaWYgKCFyZXMub2spIHsKICAgICAgY29uc3QgbWVzc2FnZSA9IChib2R5ICYmIGJvZHkuZXJyb3IpIHx8IGBSZXF1ZXN0IGZhaWxlZCAoJHtyZXMuc3RhdHVzfSlgOwogICAgICB0aHJvdyBuZXcgRXJyb3IobWVzc2FnZSk7CiAgICB9CiAgICByZXR1cm4gYm9keTsKICB9CgogIGZ1bmN0aW9uIHFzKHBhcmFtcykgewogICAgY29uc3QgdXNwID0gbmV3IFVSTFNlYXJjaFBhcmFtcygpOwogICAgT2JqZWN0LmVudHJpZXMocGFyYW1zIHx8IHt9KS5mb3JFYWNoKChbaywgdl0pID0+IHsKICAgICAgaWYgKHYgIT09IHVuZGVmaW5lZCAmJiB2ICE9PSBudWxsICYmIHYgIT09ICcnKSB1c3Auc2V0KGssIHYpOwogICAgfSk7CiAgICBjb25zdCBzID0gdXNwLnRvU3RyaW5nKCk7CiAgICByZXR1cm4gcyA/IGA/JHtzfWAgOiAnJzsKICB9CgogIHJldHVybiB7CiAgICBhdXRoTWU6ICgpID0+IHJlcXVlc3QoJy9hcGkvYXV0aC9tZScpLAogICAgYXV0aExvZ2luOiAoY29kZSkgPT4KICAgICAgcmVxdWVzdCgnL2FwaS9hdXRoL2xvZ2luJywgeyBtZXRob2Q6ICdQT1NUJywgaGVhZGVyczogeyAnQ29udGVudC1UeXBlJzogJ2FwcGxpY2F0aW9uL2pzb24nIH0sIGJvZHk6IEpTT04uc3RyaW5naWZ5KHsgY29kZSB9KSB9KSwKICAgIGF1dGhMb2dvdXQ6ICgpID0+IHJlcXVlc3QoJy9hcGkvYXV0aC9sb2dvdXQnLCB7IG1ldGhvZDogJ1BPU1QnIH0pLAoKICAgIGZpbHRlck9wdGlvbnM6ICgpID0+IHJlcXVlc3QoJy9hcGkvYW5hbHl0aWNzL2ZpbHRlci1vcHRpb25zJyksCiAgICBrcGlzOiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9rcGlzJHtxcyhwYXJhbXMpfWApLAogICAgcGxhdGZvcm1CcmVha2Rvd246IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL3BsYXRmb3JtLWJyZWFrZG93biR7cXMocGFyYW1zKX1gKSwKICAgIGNhbXBhaWduQnJlYWtkb3duOiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9jYW1wYWlnbi1icmVha2Rvd24ke3FzKHBhcmFtcyl9YCksCiAgICBjb250ZW50VHlwZUJyZWFrZG93bjogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9hbmFseXRpY3MvY29udGVudC10eXBlLWJyZWFrZG93biR7cXMocGFyYW1zKX1gKSwKICAgIG1ldHJpY09wdGlvbnM6IChwbGF0Zm9ybSkgPT4gcmVxdWVzdChgL2FwaS9hbmFseXRpY3MvbWV0cmljLW9wdGlvbnMke3FzKHsgcGxhdGZvcm0gfSl9YCksCiAgICBtZXRyaWNTdW1tYXJ5OiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9tZXRyaWMtc3VtbWFyeSR7cXMocGFyYW1zKX1gKSwKICAgIHRyZW5kOiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy90cmVuZCR7cXMocGFyYW1zKX1gKSwKICAgIHRvcFBvc3RzOiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy90b3AtcG9zdHMke3FzKHBhcmFtcyl9YCksCiAgICBjb21wYXJlOiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9jb21wYXJlJHtxcyhwYXJhbXMpfWApLAogICAgbW9udGhseTogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9hbmFseXRpY3MvbW9udGhseSR7cXMocGFyYW1zKX1gKSwKICAgIHF1YXJ0ZXJseTogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9hbmFseXRpY3MvcXVhcnRlcmx5JHtxcyhwYXJhbXMpfWApLAogICAgeXRkOiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy95dGQke3FzKHBhcmFtcyl9YCksCiAgICBwbGF0Zm9ybVJlcG9ydDogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9hbmFseXRpY3MvcGxhdGZvcm0tcmVwb3J0JHtxcyhwYXJhbXMpfWApLAoKICAgIHByZXZpZXdVcGxvYWQ6IChmaWxlKSA9PiB7CiAgICAgIGNvbnN0IGZvcm0gPSBuZXcgRm9ybURhdGEoKTsKICAgICAgZm9ybS5hcHBlbmQoJ2ZpbGUnLCBmaWxlKTsKICAgICAgcmV0dXJuIHJlcXVlc3QoJy9hcGkvdXBsb2Fkcy9wcmV2aWV3JywgeyBtZXRob2Q6ICdQT1NUJywgYm9keTogZm9ybSB9KTsKICAgIH0sCiAgICBjb21taXRVcGxvYWQ6IChwYXlsb2FkKSA9PgogICAgICByZXF1ZXN0KCcvYXBpL3VwbG9hZHMvY29tbWl0JywgewogICAgICAgIG1ldGhvZDogJ1BPU1QnLAogICAgICAgIGhlYWRlcnM6IHsgJ0NvbnRlbnQtVHlwZSc6ICdhcHBsaWNhdGlvbi9qc29uJyB9LAogICAgICAgIGJvZHk6IEpTT04uc3RyaW5naWZ5KHBheWxvYWQpLAogICAgICB9KSwKICAgIHVwbG9hZEhpc3Rvcnk6ICgpID0+IHJlcXVlc3QoJy9hcGkvdXBsb2Fkcy9oaXN0b3J5JyksCiAgICB1cGxvYWRFcnJvcnM6IChpZCkgPT4gcmVxdWVzdChgL2FwaS91cGxvYWRzLyR7aWR9L2Vycm9yc2ApLAogICAgdXBsb2FkUmF3Um93czogKGlkKSA9PiByZXF1ZXN0KGAvYXBpL3VwbG9hZHMvJHtpZH0vcmF3LXJvd3NgKSwKCiAgICBsaXN0UmVjb3JkczogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9yZWNvcmRzJHtxcyhwYXJhbXMpfWApLAogICAgcmVjb3Jkc1RhYmxlOiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL3JlY29yZHMvdGFibGUke3FzKHBhcmFtcyl9YCksCiAgICBnZXRSZWNvcmQ6IChpZCkgPT4gcmVxdWVzdChgL2FwaS9yZWNvcmRzLyR7aWR9YCksCiAgICB1cGRhdGVSZWNvcmQ6IChpZCwgdmFsdWVzKSA9PgogICAgICByZXF1ZXN0KGAvYXBpL3JlY29yZHMvJHtpZH1gLCB7CiAgICAgICAgbWV0aG9kOiAnUFVUJywKICAgICAgICBoZWFkZXJzOiB7ICdDb250ZW50LVR5cGUnOiAnYXBwbGljYXRpb24vanNvbicgfSwKICAgICAgICBib2R5OiBKU09OLnN0cmluZ2lmeSh7IHZhbHVlcyB9KSwKICAgICAgfSksCiAgICBkZWxldGVSZWNvcmRQb3N0OiAocG9zdElkKSA9PiByZXF1ZXN0KGAvYXBpL3JlY29yZHMvcG9zdC8ke3Bvc3RJZH1gLCB7IG1ldGhvZDogJ0RFTEVURScgfSksCiAgICBkZWxldGVSZWNvcmRQbGF0Zm9ybTogKHBvc3RJZCwgcGxhdGZvcm0pID0+CiAgICAgIHJlcXVlc3QoYC9hcGkvcmVjb3Jkcy9wb3N0LyR7cG9zdElkfS9wbGF0Zm9ybS8ke3BsYXRmb3JtfWAsIHsgbWV0aG9kOiAnREVMRVRFJyB9KSwKCiAgICByZXN0b3JlQmFja3VwOiAoZm9ybSkgPT4gcmVxdWVzdCgnL2FwaS9iYWNrdXAvcmVzdG9yZScsIHsgbWV0aG9kOiAnUE9TVCcsIGJvZHk6IGZvcm0gfSksCgogICAgbGlzdEZvbGxvd2VyczogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9mb2xsb3dlcnMke3FzKHBhcmFtcyl9YCksCiAgICBmb2xsb3dlcnNHcm93dGg6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvZm9sbG93ZXJzL2dyb3d0aCR7cXMocGFyYW1zKX1gKSwKICAgIGZvbGxvd2Vyc0twaXM6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvZm9sbG93ZXJzL2twaXMke3FzKHBhcmFtcyl9YCksCiAgICBzYXZlRm9sbG93ZXJzOiAocGF5bG9hZCkgPT4KICAgICAgcmVxdWVzdCgnL2FwaS9mb2xsb3dlcnMnLCB7IG1ldGhvZDogJ1BPU1QnLCBoZWFkZXJzOiB7ICdDb250ZW50LVR5cGUnOiAnYXBwbGljYXRpb24vanNvbicgfSwgYm9keTogSlNPTi5zdHJpbmdpZnkocGF5bG9hZCkgfSksCiAgICB1cGRhdGVGb2xsb3dlcnM6IChpZCwgcGF5bG9hZCkgPT4KICAgICAgcmVxdWVzdChgL2FwaS9mb2xsb3dlcnMvJHtpZH1gLCB7IG1ldGhvZDogJ1BVVCcsIGhlYWRlcnM6IHsgJ0NvbnRlbnQtVHlwZSc6ICdhcHBsaWNhdGlvbi9qc29uJyB9LCBib2R5OiBKU09OLnN0cmluZ2lmeShwYXlsb2FkKSB9KSwKICAgIGRlbGV0ZUZvbGxvd2VyczogKGlkKSA9PiByZXF1ZXN0KGAvYXBpL2ZvbGxvd2Vycy8ke2lkfWAsIHsgbWV0aG9kOiAnREVMRVRFJyB9KSwKICB9Owp9KSgpOwoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIFN0YXRlIC8gRm9ybWF0IC8gVG9hc3Qg4oCUIHNoYXJlZCBhcHAgc3RhdGUgKyBzbWFsbCB1dGlsaXRpZXMuCiAgID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwpjb25zdCBTdGF0ZSA9ICgoKSA9PiB7CiAgY29uc3QgdG9kYXkgPSBuZXcgRGF0ZSgpOwogIGNvbnN0IGlzbyA9IChkKSA9PiBkLnRvSVNPU3RyaW5nKCkuc2xpY2UoMCwgMTApOwogIGNvbnN0IHRoaXJ0eURheXNBZ28gPSBuZXcgRGF0ZSh0b2RheSk7CiAgdGhpcnR5RGF5c0Fnby5zZXREYXRlKHRoaXJ0eURheXNBZ28uZ2V0RGF0ZSgpIC0gMjkpOwoKICBjb25zdCBmaWx0ZXJzID0gewogICAgZGF0ZUZyb206IGlzbyh0aGlydHlEYXlzQWdvKSwKICAgIGRhdGVUbzogaXNvKHRvZGF5KSwKICAgIHBsYXRmb3JtOiAnYWxsJywKICAgIGNhbXBhaWduVHlwZTogJ2FsbCcsCiAgICBjb250ZW50VHlwZTogJ2FsbCcsCiAgfTsKCiAgY29uc3QgbGlzdGVuZXJzID0gW107CgogIHJldHVybiB7CiAgICBnZXRGaWx0ZXJzOiAoKSA9PiAoeyAuLi5maWx0ZXJzIH0pLAogICAgc2V0RmlsdGVycyhwYXJ0aWFsKSB7CiAgICAgIE9iamVjdC5hc3NpZ24oZmlsdGVycywgcGFydGlhbCk7CiAgICAgIGxpc3RlbmVycy5mb3JFYWNoKChmbikgPT4gZm4odGhpcy5nZXRGaWx0ZXJzKCkpKTsKICAgIH0sCiAgICBvbkNoYW5nZShmbikgewogICAgICBsaXN0ZW5lcnMucHVzaChmbik7CiAgICB9LAogIH07Cn0pKCk7Cgpjb25zdCBGb3JtYXQgPSB7CiAgbnVtYmVyKG4pIHsKICAgIGlmIChuID09PSBudWxsIHx8IG4gPT09IHVuZGVmaW5lZCkgcmV0dXJuICfigJQnOwogICAgcmV0dXJuIE1hdGgucm91bmQobikudG9Mb2NhbGVTdHJpbmcoJ2VuLVVTJyk7CiAgfSwKICBjb21wYWN0KG4pIHsKICAgIGlmIChuID09PSBudWxsIHx8IG4gPT09IHVuZGVmaW5lZCkgcmV0dXJuICfigJQnOwogICAgY29uc3QgYWJzID0gTWF0aC5hYnMobik7CiAgICBpZiAoYWJzID49IDFfMDAwXzAwMCkgcmV0dXJuIGAkeyhuIC8gMV8wMDBfMDAwKS50b0ZpeGVkKDEpLnJlcGxhY2UoL1wuMCQvLCAnJyl9TWA7CiAgICBpZiAoYWJzID49IDFfMDAwKSByZXR1cm4gYCR7KG4gLyAxXzAwMCkudG9GaXhlZCgxKS5yZXBsYWNlKC9cLjAkLywgJycpfUtgOwogICAgcmV0dXJuIGAke01hdGgucm91bmQobil9YDsKICB9LAogIC8qKiBEYXNoYm9hcmQtd2lkZSAicHJvZmVzc2lvbmFsIiBudW1iZXIgZm9ybWF0OiBwbGFpbiB1bmRlciAxLDAwMDsgY29tbWEtZ3JvdXBlZAogICAgICB1cCB0byAxMCwwMDA7IGFiYnJldmlhdGVkIChLL00pIGJleW9uZCB0aGF0IOKAlCBlLmcuIDg1MCwgMSwyNTAsIDEyLjVLLCAxNTZLLCAxLjI1TS4gKi8KICBzbWFydChuKSB7CiAgICBpZiAobiA9PT0gbnVsbCB8fCBuID09PSB1bmRlZmluZWQpIHJldHVybiAn4oCUJzsKICAgIGNvbnN0IGFicyA9IE1hdGguYWJzKG4pOwogICAgaWYgKGFicyA8IDEwMDApIHJldHVybiBgJHtNYXRoLnJvdW5kKG4pfWA7CiAgICBpZiAoYWJzIDwgMTAwMDApIHJldHVybiBNYXRoLnJvdW5kKG4pLnRvTG9jYWxlU3RyaW5nKCdlbi1VUycpOwogICAgaWYgKGFicyA8IDFfMDAwXzAwMCkgcmV0dXJuIGAkeyhuIC8gMTAwMCkudG9GaXhlZCgxKS5yZXBsYWNlKC9cLjAkLywgJycpfUtgOwogICAgcmV0dXJuIGAkeyhuIC8gMV8wMDBfMDAwKS50b0ZpeGVkKDIpLnJlcGxhY2UoL1wuPzArJC8sICcnKX1NYDsKICB9LAogIHBlcmNlbnQobikgewogICAgaWYgKG4gPT09IG51bGwgfHwgbiA9PT0gdW5kZWZpbmVkKSByZXR1cm4gJ+KAlCc7CiAgICByZXR1cm4gYCR7TnVtYmVyKG4pLnRvRml4ZWQoMSkucmVwbGFjZSgvXC4wJC8sICcnKX0lYDsKICB9LAogIHBjdChuKSB7CiAgICBpZiAobiA9PT0gbnVsbCB8fCBuID09PSB1bmRlZmluZWQpIHJldHVybiAn4oCUJzsKICAgIGNvbnN0IHNpZ24gPSBuID4gMCA/ICcrJyA6ICcnOwogICAgcmV0dXJuIGAke3NpZ259JHtuLnRvRml4ZWQoMSl9JWA7CiAgfSwKICBkYXRlKGlzb18pIHsKICAgIGlmICghaXNvXykgcmV0dXJuICfigJQnOwogICAgY29uc3QgW3ksIG0sIGRdID0gaXNvXy5zcGxpdCgnLScpLm1hcChOdW1iZXIpOwogICAgcmV0dXJuIG5ldyBEYXRlKHksIG0gLSAxLCBkKS50b0xvY2FsZURhdGVTdHJpbmcoJ2VuLVVTJywgeyBtb250aDogJ3Nob3J0JywgZGF5OiAnbnVtZXJpYycsIHllYXI6ICdudW1lcmljJyB9KTsKICB9LAogIGR1cmF0aW9uKHNlY29uZHMpIHsKICAgIGlmIChzZWNvbmRzID09PSBudWxsIHx8IHNlY29uZHMgPT09IHVuZGVmaW5lZCkgcmV0dXJuICfigJQnOwogICAgY29uc3QgcyA9IE1hdGgucm91bmQoc2Vjb25kcyk7CiAgICBpZiAocyA8IDYwKSByZXR1cm4gYCR7c31zYDsKICAgIGlmIChzIDwgMzYwMCkgcmV0dXJuIGAke01hdGguZmxvb3IocyAvIDYwKX1tICR7cyAlIDYwfXNgOwogICAgY29uc3QgaCA9IE1hdGguZmxvb3IocyAvIDM2MDApOwogICAgY29uc3QgbSA9IE1hdGgucm91bmQoKHMgJSAzNjAwKSAvIDYwKTsKICAgIHJldHVybiBgJHtofWggJHttfW1gOwogIH0sCiAgZGVsdGFDbGFzcyhuKSB7CiAgICBpZiAobiA9PT0gbnVsbCB8fCBuID09PSB1bmRlZmluZWQpIHJldHVybiAnZmxhdCc7CiAgICBpZiAobiA+IDAuNSkgcmV0dXJuICd1cCc7CiAgICBpZiAobiA8IC0wLjUpIHJldHVybiAnZG93bic7CiAgICByZXR1cm4gJ2ZsYXQnOwogIH0sCn07Cgpjb25zdCBUb2FzdCA9IHsKICBzaG93KG1lc3NhZ2UsIHR5cGUgPSAnc3VjY2VzcycpIHsKICAgIGNvbnN0IHJvb3QgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndG9hc3RSb290Jyk7CiAgICBjb25zdCBlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZWwuY2xhc3NOYW1lID0gYHRvYXN0ICR7dHlwZX1gOwogICAgZWwudGV4dENvbnRlbnQgPSBtZXNzYWdlOwogICAgcm9vdC5hcHBlbmRDaGlsZChlbCk7CiAgICBzZXRUaW1lb3V0KCgpID0+IGVsLnJlbW92ZSgpLCA1MDAwKTsKICB9LAp9OwoKLyoqIFNhZmVseSBidWlsZHMgRE9NIHRleHQgbm9kZXMgZm9yIHVudHJ1c3RlZCBzdHJpbmdzIChjYXB0aW9ucywgZmlsZW5hbWVzLCBwbGF0Zm9ybSBsYWJlbHMgZnJvbSBkYXRhKS4gKi8KZnVuY3Rpb24gdGV4dEVsKHRhZywgdGV4dCwgY2xhc3NOYW1lKSB7CiAgY29uc3QgZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KHRhZyk7CiAgaWYgKGNsYXNzTmFtZSkgZWwuY2xhc3NOYW1lID0gY2xhc3NOYW1lOwogIGVsLmFwcGVuZENoaWxkKGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKHRleHQgPz8gJycpKTsKICByZXR1cm4gZWw7Cn0KCi8qKiBBIHByZW1pdW0gZW1wdHkgc3RhdGU6IGljb24gKyBleHBsYW5hdGlvbiArIG9wdGlvbmFsIGFjdGlvbiwgaW5zdGVhZCBvZiBhIGJsYW5rIGFyZWEuCiAgICBJY29ucyByZW5kZXIgdmlhIHRoZSBwYWdlLXdpZGUgTXV0YXRpb25PYnNlcnZlciB0aGF0IGNhbGxzIGx1Y2lkZS5jcmVhdGVJY29ucygpIChzZWUgYm9vdHN0cmFwKS4gKi8KZnVuY3Rpb24gZW1wdHlTdGF0ZSh7IGljb24gPSAnaW5ib3gnLCB0aXRsZSwgbWVzc2FnZSwgYWN0aW9uTGFiZWwsIG9uQWN0aW9uIH0pIHsKICBjb25zdCB3cmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgd3JhcC5jbGFzc05hbWUgPSAnZW1wdHktc3RhdGUnOwogIGNvbnN0IGljb25XcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgaWNvbldyYXAuY2xhc3NOYW1lID0gJ2VtcHR5LWljb24nOwogIGljb25XcmFwLmlubmVySFRNTCA9IGA8aSBkYXRhLWx1Y2lkZT0iJHtpY29ufSIgc3R5bGU9IndpZHRoOjIycHg7aGVpZ2h0OjIycHg7Ij48L2k+YDsKICB3cmFwLmFwcGVuZENoaWxkKGljb25XcmFwKTsKICBpZiAodGl0bGUpIHdyYXAuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCB0aXRsZSwgJ2VtcHR5LXRpdGxlJykpOwogIGlmIChtZXNzYWdlKSB3cmFwLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgbWVzc2FnZSwgJ2VtcHR5LW1lc3NhZ2UnKSk7CiAgaWYgKGFjdGlvbkxhYmVsICYmIG9uQWN0aW9uKSB7CiAgICBjb25zdCBidG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIGJ0bi5jbGFzc05hbWUgPSAnYnRuIHByaW1hcnknOwogICAgYnRuLnRleHRDb250ZW50ID0gYWN0aW9uTGFiZWw7CiAgICBidG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBvbkFjdGlvbik7CiAgICB3cmFwLmFwcGVuZENoaWxkKGJ0bik7CiAgfQogIHJldHVybiB3cmFwOwp9CgovKiogQSA8YnV0dG9uPiB3aXRoIGEgc21hbGwgbGVhZGluZyBMdWNpZGUgaWNvbiBiZWZvcmUgaXRzIGxhYmVsIChsYWJlbCBpcyBhbHdheXMgYSBzdGF0aWMsIGRldmVsb3Blci1zdXBwbGllZCBzdHJpbmcgYXQgY2FsbCBzaXRlcywgbmV2ZXIgdXNlciBkYXRhIOKAlCBpbnNlcnRlZCB2aWEgY3JlYXRlVGV4dE5vZGUgcmVnYXJkbGVzcykuICovCmZ1bmN0aW9uIGljb25CdG4oY2xhc3NOYW1lLCBpY29uTmFtZSwgbGFiZWwpIHsKICBjb25zdCBidG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICBidG4uY2xhc3NOYW1lID0gY2xhc3NOYW1lOwogIGNvbnN0IGljb24gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpJyk7CiAgaWNvbi5zZXRBdHRyaWJ1dGUoJ2RhdGEtbHVjaWRlJywgaWNvbk5hbWUpOwogIGljb24uc3R5bGUud2lkdGggPSAnMTNweCc7CiAgaWNvbi5zdHlsZS5oZWlnaHQgPSAnMTNweCc7CiAgYnRuLmFwcGVuZENoaWxkKGljb24pOwogIGJ0bi5hcHBlbmRDaGlsZChkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShgICR7bGFiZWx9YCkpOwogIHJldHVybiBidG47Cn0KCi8qKiBTaGltbWVyaW5nIHBsYWNlaG9sZGVycyBzaG93biB0aGUgaW5zdGFudCBhIHNlY3Rpb24gc3RhcnRzIGxvYWRpbmcsIHN3YXBwZWQgZm9yIHJlYWwKICAgIGNvbnRlbnQgKG9yIGFuIGVtcHR5IHN0YXRlKSBvbmNlIHRoZSBmZXRjaCByZXNvbHZlcyDigJQgbm8gYmxhbmsgYXJlYXMgd2hpbGUgd2FpdGluZy4gKi8KZnVuY3Rpb24gc2tlbGV0b25TdGF0R3JpZChjb3VudCA9IDYpIHsKICBjb25zdCBncmlkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgZ3JpZC5jbGFzc05hbWUgPSAnc2tlbGV0b24tc3RhdC1ncmlkJzsKICBmb3IgKGxldCBpID0gMDsgaSA8IGNvdW50OyBpICs9IDEpIHsKICAgIGNvbnN0IHRpbGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRpbGUuY2xhc3NOYW1lID0gJ3NrZWxldG9uIHNrZWxldG9uLXRpbGUnOwogICAgZ3JpZC5hcHBlbmRDaGlsZCh0aWxlKTsKICB9CiAgcmV0dXJuIGdyaWQ7Cn0KZnVuY3Rpb24gc2tlbGV0b25DaGFydCgpIHsKICBjb25zdCBkaXYgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICBkaXYuY2xhc3NOYW1lID0gJ3NrZWxldG9uIHNrZWxldG9uLWNoYXJ0JzsKICByZXR1cm4gZGl2Owp9CmZ1bmN0aW9uIHNrZWxldG9uUm93cyhjb3VudCA9IDYpIHsKICBjb25zdCB3cmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgZm9yIChsZXQgaSA9IDA7IGkgPCBjb3VudDsgaSArPSAxKSB7CiAgICBjb25zdCByb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHJvdy5jbGFzc05hbWUgPSAnc2tlbGV0b24gc2tlbGV0b24tcm93JzsKICAgIHdyYXAuYXBwZW5kQ2hpbGQocm93KTsKICB9CiAgcmV0dXJuIHdyYXA7Cn0KCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBTaGFyZWQgYW5pbWF0aW9uIHByaW1pdGl2ZXMg4oCUIGEgY291bnQtdXAgZm9yIEtQSSBudW1iZXJzIGFuZCBhCiAgIENTUyB3aWR0aC10cmFuc2l0aW9uIGJhciwgYm90aCByZXVzZWQgYWNyb3NzIHRoZSBEYXNoYm9hcmQgYW5kCiAgIENvbXBhcmlzb25zIHBhZ2VzLiBCb3RoIHJlc3BlY3QgcHJlZmVycy1yZWR1Y2VkLW1vdGlvbiAoZ3VhcmRlZAogICBpbiBDU1MsIHNlZSAuYmFyLWZpbGwgLyB0aGUgYW5pbWF0ZUNvdW50IGR1cmF0aW9uIGNoZWNrIGJlbG93KS4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmNvbnN0IFBSRUZFUlNfUkVEVUNFRF9NT1RJT04gPSB3aW5kb3cubWF0Y2hNZWRpYSAmJiB3aW5kb3cubWF0Y2hNZWRpYSgnKHByZWZlcnMtcmVkdWNlZC1tb3Rpb246IHJlZHVjZSknKS5tYXRjaGVzOwoKLyoqIFNocmlua3MgYGVsYCdzIGZvbnQgc2l6ZSBqdXN0IGVub3VnaCBmb3IgaXRzIGN1cnJlbnQgdGV4dCB0byBmaXQgaXRzIG93biB3aWR0aCDigJQgYSBLUEkgdGlsZSdzIGJveCBpcyBhIGZpeGVkIHNpemUsIGJ1dCB0aGUgdmFsdWUgaW5zaWRlIGl0IGlzbid0IChhIGZvbGxvd2VyIGNvdW50IGNhbiBiZSAiMCIgb3IgIjEsMDQ4LDU3NiIpLCBzbyBhIHNpbmdsZSBmaXhlZCBmb250LXNpemUgd2lsbCBldmVudHVhbGx5IG92ZXJmbG93LiBSZXNldHMgdG8gdGhlIENTUy1kZWZpbmVkIHNpemUgZmlyc3QsIHRoZW4gc3RlcHMgZG93biBieSAxcHggYXQgYSB0aW1lIHVudGlsIGl0IGZpdHMgb3IgaGl0cyBgbWluU2l6ZWAuICovCmZ1bmN0aW9uIGZpdFN0YXRWYWx1ZShlbCwgbWluU2l6ZSA9IDE4KSB7CiAgaWYgKCFlbCkgcmV0dXJuOwogIGVsLnN0eWxlLmZvbnRTaXplID0gJyc7CiAgY29uc3QgbWF4V2lkdGggPSBlbC5jbGllbnRXaWR0aDsKICBpZiAoIW1heFdpZHRoKSByZXR1cm47CiAgbGV0IHNpemUgPSBwYXJzZUZsb2F0KGdldENvbXB1dGVkU3R5bGUoZWwpLmZvbnRTaXplKTsKICB3aGlsZSAoZWwuc2Nyb2xsV2lkdGggPiBtYXhXaWR0aCAmJiBzaXplID4gbWluU2l6ZSkgewogICAgc2l6ZSAtPSAxOwogICAgZWwuc3R5bGUuZm9udFNpemUgPSBgJHtzaXplfXB4YDsKICB9Cn0KCi8qKiBBbmltYXRlcyBhIG51bWJlciBmcm9tIGBmcm9tYCB0byBgdG9gIGluc2lkZSBgZWxgIG92ZXIgYGR1cmF0aW9uYG1zLCBmb3JtYXR0aW5nIGVhY2ggZnJhbWUgd2l0aCBgZm9ybWF0YCAoZGVmYXVsdHMgdG8gYSBwbGFpbiByb3VuZGVkIGludGVnZXIpLiBTa2lwcyBzdHJhaWdodCB0byB0aGUgZmluYWwgdmFsdWUgdW5kZXIgcHJlZmVycy1yZWR1Y2VkLW1vdGlvbi4gU2hyaW5rcyB0aGUgZm9udCB0byBmaXQgb25jZSB0aGUgZmluYWwgdmFsdWUgbGFuZHMsIHNpbmNlIHRoZSBhbmltYXRlZCBkaWdpdHMgY2FuIGJlIGEgZGlmZmVyZW50IHdpZHRoIHRoYW4gdGhlIHNldHRsZWQgdmFsdWUg4oCUIGRlZmVycmVkIGEgZnJhbWUgYmVjYXVzZSBgZWxgIGlzIHR5cGljYWxseSBzdGlsbCBkZXRhY2hlZCBmcm9tIHRoZSBkb2N1bWVudCAobWlkLWNvbnN0cnVjdGlvbiBieSBpdHMgY2FsbGVyKSB3aGVuIGFuaW1hdGVDb3VudCBpcyBmaXJzdCBpbnZva2VkLCBhbmQgY2xpZW50V2lkdGggcmVhZHMgMCB1bnRpbCBpdCdzIGFjdHVhbGx5IGF0dGFjaGVkIGFuZCBsYWlkIG91dC4gKi8KZnVuY3Rpb24gYW5pbWF0ZUNvdW50KGVsLCBmcm9tLCB0bywgZHVyYXRpb24gPSA5MDAsIGZvcm1hdCkgewogIGlmICghZWwpIHJldHVybjsKICBjb25zdCBmbXQgPSBmb3JtYXQgfHwgKCh2KSA9PiBNYXRoLnJvdW5kKHYpLnRvTG9jYWxlU3RyaW5nKCdlbi1VUycpKTsKICBpZiAoUFJFRkVSU19SRURVQ0VEX01PVElPTiB8fCBmcm9tID09PSB0byB8fCAhTnVtYmVyLmlzRmluaXRlKGZyb20pIHx8ICFOdW1iZXIuaXNGaW5pdGUodG8pKSB7CiAgICBlbC50ZXh0Q29udGVudCA9IGZtdCh0byk7CiAgICByZXF1ZXN0QW5pbWF0aW9uRnJhbWUoKCkgPT4gZml0U3RhdFZhbHVlKGVsKSk7CiAgICByZXR1cm47CiAgfQogIGNvbnN0IHN0YXJ0ID0gcGVyZm9ybWFuY2Uubm93KCk7CiAgZnVuY3Rpb24gdGljayhub3cpIHsKICAgIGNvbnN0IGVsYXBzZWQgPSBub3cgLSBzdGFydDsKICAgIGNvbnN0IHByb2dyZXNzID0gTWF0aC5taW4oMSwgZWxhcHNlZCAvIGR1cmF0aW9uKTsKICAgIGNvbnN0IGVhc2VkID0gMSAtIE1hdGgucG93KDEgLSBwcm9ncmVzcywgMyk7IC8vIGVhc2VPdXRDdWJpYwogICAgZWwudGV4dENvbnRlbnQgPSBmbXQoZnJvbSArICh0byAtIGZyb20pICogZWFzZWQpOwogICAgaWYgKHByb2dyZXNzIDwgMSkgewogICAgICByZXF1ZXN0QW5pbWF0aW9uRnJhbWUodGljayk7CiAgICB9IGVsc2UgewogICAgICBmaXRTdGF0VmFsdWUoZWwpOwogICAgfQogIH0KICByZXF1ZXN0QW5pbWF0aW9uRnJhbWUodGljayk7Cn0KCi8qKiBBIGxhYmVsZWQgaG9yaXpvbnRhbCBiYXIgdGhhdCBhbmltYXRlcyBpdHMgd2lkdGggaW4gb24gaW5zZXJ0aW9uIOKAlCB1c2VkIGZvciB0aGUgQ29tcGFyaXNvbnMgcGFnZSdzIHBhaXJlZCBSYW5nZSBBL0IgYmFycy4gYHZhbHVlYC9gbWF4YCBkcml2ZSB0aGUgZmlsbCBwZXJjZW50YWdlOyBgY29sb3JWYXJgIGlzIGEgQ1NTIGN1c3RvbSBwcm9wZXJ0eSBuYW1lIChlLmcuICctLXNlcmllcy0xJykuICovCmZ1bmN0aW9uIGJ1aWxkQmFyKHsgbGFiZWwsIHZhbHVlLCBtYXgsIGNvbG9yVmFyLCBmb3JtYXRWYWx1ZSB9KSB7CiAgY29uc3Qgcm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgcm93LmNsYXNzTmFtZSA9ICdiYXItcm93JzsKICBjb25zdCBsYWJlbEVsID0gdGV4dEVsKCdkaXYnLCBsYWJlbCwgJ2Jhci1sYWJlbCcpOwogIGNvbnN0IHRyYWNrID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgdHJhY2suY2xhc3NOYW1lID0gJ2Jhci10cmFjayc7CiAgY29uc3QgZmlsbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogIGZpbGwuY2xhc3NOYW1lID0gJ2Jhci1maWxsJzsKICBmaWxsLnN0eWxlLmJhY2tncm91bmQgPSBjb2xvclZhciA/IGB2YXIoJHtjb2xvclZhcn0pYCA6ICd2YXIoLS1zZXJpZXMtMSknOwogIHRyYWNrLmFwcGVuZENoaWxkKGZpbGwpOwogIGNvbnN0IHZhbHVlRWwgPSB0ZXh0RWwoJ2RpdicsIGZvcm1hdFZhbHVlID8gZm9ybWF0VmFsdWUodmFsdWUpIDogU3RyaW5nKHZhbHVlKSwgJ2Jhci12YWx1ZScpOwogIHJvdy5hcHBlbmQobGFiZWxFbCwgdHJhY2ssIHZhbHVlRWwpOwogIGNvbnN0IHBjdCA9IG1heCA+IDAgPyBNYXRoLm1pbigxMDAsIE1hdGgucm91bmQoKHZhbHVlIC8gbWF4KSAqIDEwMDApIC8gMTApIDogMDsKICByZXF1ZXN0QW5pbWF0aW9uRnJhbWUoKCkgPT4geyBmaWxsLnN0eWxlLndpZHRoID0gYCR7cGN0fSVgOyB9KTsKICByZXR1cm4gcm93Owp9CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgU2hhcmVkIHRhYmxlIHRvb2xiYXIgcGllY2VzIOKAlCBzZWFyY2ggYm94LCBjbGllbnQtc2lkZSBwYWdlciwKICAgYW5kIENTVi9YTFNYIGV4cG9ydCDigJQgcmV1c2VkIGJ5IHRoZSBGb2xsb3dlcnMgRGF0YSBhbmQgVXBsb2FkCiAgIEhpc3RvcnkgdGFicyAoYm90aCBsb2FkIHRoZWlyIGZ1bGwgZGF0YXNldCBvbmNlIGFuZCBzZWFyY2gvCiAgIHNvcnQvcGFnaW5hdGUgaXQgaW4gdGhlIGJyb3dzZXIsIHVubGlrZSBEYXRhIFJlY29yZHMgd2hpY2ggaXMKICAgc2VydmVyLXBhZ2luYXRlZCkuIENTViBuZWVkcyBubyBzZXJ2ZXIgcm91bmQgdHJpcCBhdCBhbGw7IFhMU1gKICAgZ29lcyB0aHJvdWdoIFBPU1QgL2FwaS9leHBvcnQgc28gZXhjZWxqcyAoYWxyZWFkeSBhIGRlcGVuZGVuY3kpCiAgIGNhbiBnZW5lcmF0ZSBhIHJlYWwgLnhsc3ggd2l0aG91dCBhZGRpbmcgYSBjbGllbnQtc2lkZSBsaWJyYXJ5LgogICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KZnVuY3Rpb24gYnVpbGRTZWFyY2hCb3goeyBwbGFjZWhvbGRlciwgdmFsdWUsIG9uQ2hhbmdlIH0pIHsKICBjb25zdCB3cmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgd3JhcC5jbGFzc05hbWUgPSAncmVjb3Jkcy1zZWFyY2gnOwogIGNvbnN0IGlucHV0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsKICBpbnB1dC50eXBlID0gJ3NlYXJjaCc7CiAgaW5wdXQucGxhY2Vob2xkZXIgPSBwbGFjZWhvbGRlcjsKICBpbnB1dC52YWx1ZSA9IHZhbHVlIHx8ICcnOwogIGxldCBkZWJvdW5jZSA9IG51bGw7CiAgaW5wdXQuYWRkRXZlbnRMaXN0ZW5lcignaW5wdXQnLCAoKSA9PiB7CiAgICBjbGVhclRpbWVvdXQoZGVib3VuY2UpOwogICAgZGVib3VuY2UgPSBzZXRUaW1lb3V0KCgpID0+IG9uQ2hhbmdlKGlucHV0LnZhbHVlKSwgMzAwKTsKICB9KTsKICB3cmFwLmFwcGVuZENoaWxkKGlucHV0KTsKICByZXR1cm4gd3JhcDsKfQoKLyoqIFNsaWNlcyBhbiBhbHJlYWR5LWxvYWRlZCwgYWxyZWFkeS1maWx0ZXJlZC9zb3J0ZWQgYXJyYXkgZm9yIGNsaWVudC1zaWRlIHBhZ2luYXRpb24g4oCUIHRoZSBjb3VudGVycGFydCB0byB0aGUgc2VydmVyLXNpZGUgcGFnaW5hdGUoKSBpbiBhcHAuanMsIGZvciB0YWJsZXMgdGhhdCBkb24ndCBoYXZlIGEgcGFnaW5hdGVkIGVuZHBvaW50LiAqLwpmdW5jdGlvbiBwYWdpbmF0ZUNsaWVudFNpZGUocm93cywgcGFnZSwgcGFnZVNpemUpIHsKICBjb25zdCB0b3RhbFBhZ2VzID0gTWF0aC5tYXgoMSwgTWF0aC5jZWlsKHJvd3MubGVuZ3RoIC8gcGFnZVNpemUpKTsKICBjb25zdCBzYWZlUGFnZSA9IE1hdGgubWluKE1hdGgubWF4KDEsIHBhZ2UpLCB0b3RhbFBhZ2VzKTsKICBjb25zdCBzdGFydCA9IChzYWZlUGFnZSAtIDEpICogcGFnZVNpemU7CiAgcmV0dXJuIHsgcGFnZVJvd3M6IHJvd3Muc2xpY2Uoc3RhcnQsIHN0YXJ0ICsgcGFnZVNpemUpLCB0b3RhbFBhZ2VzLCBzYWZlUGFnZSwgdG90YWw6IHJvd3MubGVuZ3RoIH07Cn0KCmZ1bmN0aW9uIGJ1aWxkUGFnZXIoeyBwYWdlLCB0b3RhbFBhZ2VzLCB0b3RhbCwgb25QcmV2LCBvbk5leHQgfSkgewogIGNvbnN0IHBhZ2VyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgcGFnZXIuY2xhc3NOYW1lID0gJ3BhZ2luYXRpb24tcm93JzsKICBjb25zdCBwcmV2QnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgcHJldkJ0bi5jbGFzc05hbWUgPSAnYnRuJzsgcHJldkJ0bi50eXBlID0gJ2J1dHRvbic7IHByZXZCdG4udGV4dENvbnRlbnQgPSAnUHJldmlvdXMnOwogIHByZXZCdG4uZGlzYWJsZWQgPSBwYWdlIDw9IDE7CiAgcHJldkJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIG9uUHJldik7CiAgY29uc3QgbmV4dEJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogIG5leHRCdG4uY2xhc3NOYW1lID0gJ2J0bic7IG5leHRCdG4udHlwZSA9ICdidXR0b24nOyBuZXh0QnRuLnRleHRDb250ZW50ID0gJ05leHQnOwogIG5leHRCdG4uZGlzYWJsZWQgPSBwYWdlID49IHRvdGFsUGFnZXM7CiAgbmV4dEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIG9uTmV4dCk7CiAgcGFnZXIuYXBwZW5kKHByZXZCdG4sIHRleHRFbCgnc3BhbicsIGBQYWdlICR7cGFnZX0gb2YgJHt0b3RhbFBhZ2VzfSDigJQgJHt0b3RhbH0gcmVjb3JkKHMpYCksIG5leHRCdG4pOwogIHJldHVybiBwYWdlcjsKfQoKZnVuY3Rpb24gZG93bmxvYWRCbG9iKGZpbGVuYW1lLCBtaW1lVHlwZSwgY29udGVudCkgewogIGNvbnN0IGJsb2IgPSBuZXcgQmxvYihbY29udGVudF0sIHsgdHlwZTogbWltZVR5cGUgfSk7CiAgY29uc3QgdXJsID0gVVJMLmNyZWF0ZU9iamVjdFVSTChibG9iKTsKICBjb25zdCBhID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYScpOwogIGEuaHJlZiA9IHVybDsKICBhLmRvd25sb2FkID0gZmlsZW5hbWU7CiAgZG9jdW1lbnQuYm9keS5hcHBlbmRDaGlsZChhKTsKICBhLmNsaWNrKCk7CiAgYS5yZW1vdmUoKTsKICBVUkwucmV2b2tlT2JqZWN0VVJMKHVybCk7Cn0KCi8qKiBRdW90ZWQtY29tbWEtam9pbiBDU1Yg4oCUIHRoZSBjbGllbnQtc2lkZSBtaXJyb3Igb2YgYXBwLmpzJ3MgdG9DU1YoKSwgZm9yIHRhYmxlcyB3aG9zZSBmdWxsIGRhdGFzZXQgaXMgYWxyZWFkeSBsb2FkZWQgaW4gdGhlIGJyb3dzZXIuICovCmZ1bmN0aW9uIHRvQ1NWQ2xpZW50U2lkZShyb3dzLCBjb2x1bW5zKSB7CiAgY29uc3QgZXNjYXBlID0gKHYpID0+IHsKICAgIGNvbnN0IHMgPSB2ID09PSBudWxsIHx8IHYgPT09IHVuZGVmaW5lZCA/ICcnIDogU3RyaW5nKHYpOwogICAgcmV0dXJuIC9bIixcclxuXS8udGVzdChzKSA/IGAiJHtzLnJlcGxhY2UoLyIvZywgJyIiJyl9ImAgOiBzOwogIH07CiAgY29uc3QgbGluZXMgPSBbY29sdW1ucy5tYXAoKGMpID0+IGVzY2FwZShjLmxhYmVsKSkuam9pbignLCcpXTsKICByb3dzLmZvckVhY2goKHJvdykgPT4gbGluZXMucHVzaChjb2x1bW5zLm1hcCgoYykgPT4gZXNjYXBlKHJvd1tjLmtleV0pKS5qb2luKCcsJykpKTsKICByZXR1cm4gbGluZXMuam9pbignXHJcbicpOwp9CgovKiogQSBzbWFsbCAiRXhwb3J0IENTViAvIEV4cG9ydCBFeGNlbCIgYnV0dG9uIHBhaXIuIGBnZXRSb3dzQW5kQ29sdW1ucygpYCBpcyBjYWxsZWQgYXQgY2xpY2sgdGltZSBzbyBpdCBhbHdheXMgZXhwb3J0cyB3aGF0ZXZlcidzIGN1cnJlbnRseSBmaWx0ZXJlZC9zb3J0ZWQsIG5ldmVyIGEgc3RhbGUgc25hcHNob3QuICovCmZ1bmN0aW9uIGJ1aWxkRXhwb3J0QnV0dG9ucyh7IGdldFJvd3NBbmRDb2x1bW5zLCBmaWxlbmFtZUJhc2UsIHNoZWV0TmFtZSB9KSB7CiAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogIHdyYXAuY2xhc3NOYW1lID0gJ2V4cG9ydC1idXR0b25zJzsKICBjb25zdCBjc3ZCdG4gPSBpY29uQnRuKCdidG4nLCAnZmlsZS1kb3duJywgJ0V4cG9ydCBDU1YnKTsKICBjc3ZCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7CiAgICBjb25zdCB7IHJvd3MsIGNvbHVtbnMgfSA9IGdldFJvd3NBbmRDb2x1bW5zKCk7CiAgICBpZiAoIXJvd3MubGVuZ3RoKSB7IFRvYXN0LnNob3coJ05vdGhpbmcgdG8gZXhwb3J0LicsICdlcnJvcicpOyByZXR1cm47IH0KICAgIGRvd25sb2FkQmxvYihgJHtmaWxlbmFtZUJhc2V9LmNzdmAsICd0ZXh0L2NzdjtjaGFyc2V0PXV0Zi04JywgdG9DU1ZDbGllbnRTaWRlKHJvd3MsIGNvbHVtbnMpKTsKICB9KTsKICBjb25zdCB4bHN4QnRuID0gaWNvbkJ0bignYnRuJywgJ2ZpbGUtc3ByZWFkc2hlZXQnLCAnRXhwb3J0IEV4Y2VsJyk7CiAgeGxzeEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIGFzeW5jICgpID0+IHsKICAgIGNvbnN0IHsgcm93cywgY29sdW1ucyB9ID0gZ2V0Um93c0FuZENvbHVtbnMoKTsKICAgIGlmICghcm93cy5sZW5ndGgpIHsgVG9hc3Quc2hvdygnTm90aGluZyB0byBleHBvcnQuJywgJ2Vycm9yJyk7IHJldHVybjsgfQogICAgdHJ5IHsKICAgICAgY29uc3QgcmVzID0gYXdhaXQgZmV0Y2goJy9hcGkvZXhwb3J0JywgewogICAgICAgIG1ldGhvZDogJ1BPU1QnLAogICAgICAgIGhlYWRlcnM6IHsgJ0NvbnRlbnQtVHlwZSc6ICdhcHBsaWNhdGlvbi9qc29uJyB9LAogICAgICAgIGJvZHk6IEpTT04uc3RyaW5naWZ5KHsgcm93cywgY29sdW1ucywgZm9ybWF0OiAneGxzeCcsIGZpbGVuYW1lOiBmaWxlbmFtZUJhc2UsIHNoZWV0TmFtZTogc2hlZXROYW1lIHx8IGZpbGVuYW1lQmFzZSB9KSwKICAgICAgfSk7CiAgICAgIGlmICghcmVzLm9rKSB0aHJvdyBuZXcgRXJyb3IoJ0V4cG9ydCBmYWlsZWQuJyk7CiAgICAgIGNvbnN0IGJsb2IgPSBhd2FpdCByZXMuYmxvYigpOwogICAgICBjb25zdCB1cmwgPSBVUkwuY3JlYXRlT2JqZWN0VVJMKGJsb2IpOwogICAgICBjb25zdCBhID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYScpOwogICAgICBhLmhyZWYgPSB1cmw7CiAgICAgIGEuZG93bmxvYWQgPSBgJHtmaWxlbmFtZUJhc2V9Lnhsc3hgOwogICAgICBkb2N1bWVudC5ib2R5LmFwcGVuZENoaWxkKGEpOwogICAgICBhLmNsaWNrKCk7CiAgICAgIGEucmVtb3ZlKCk7CiAgICAgIFVSTC5yZXZva2VPYmplY3RVUkwodXJsKTsKICAgIH0gY2F0Y2ggKGVycikgewogICAgICBUb2FzdC5zaG93KGVyci5tZXNzYWdlIHx8ICdFeHBvcnQgZmFpbGVkLicsICdlcnJvcicpOwogICAgfQogIH0pOwogIHdyYXAuYXBwZW5kKGNzdkJ0biwgeGxzeEJ0bik7CiAgcmV0dXJuIHdyYXA7Cn0KCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBDaGFydHMg4oCUIENoYXJ0LmpzIGJ1aWxkZXJzICh2YWxpZGF0ZWQgY2F0ZWdvcmljYWwgcGFsZXR0ZSwKICAgaGFpcmxpbmUgcmVjZXNzaXZlIGdyaWRsaW5lcywgc2luZ2xlIGF4aXMsIGxlZ2VuZCBhbHdheXMKICAgcHJlc2VudCBmb3IgMisgc2VyaWVzLCBpbmRleC1tb2RlIHRvb2x0aXBzKS4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmlmICh3aW5kb3cuQ2hhcnREYXRhTGFiZWxzKSBDaGFydC5yZWdpc3Rlcih3aW5kb3cuQ2hhcnREYXRhTGFiZWxzKTsKCmNvbnN0IENoYXJ0cyA9ICgoKSA9PiB7CiAgY29uc3QgcmVnaXN0cnkgPSBuZXcgTWFwKCk7IC8vIGNhbnZhc0lkIC0+IENoYXJ0IGluc3RhbmNlLCBzbyByZS1yZW5kZXJzIGRlc3Ryb3kgdGhlIG9sZCBvbmUgZmlyc3QKICBjb25zdCBNQVhfTEFCRUxFRF9JVEVNUyA9IDIwOyAvLyBiZXlvbmQgdGhpcywgcGVyLWl0ZW0gdmFsdWUgbGFiZWxzIHdvdWxkIG92ZXJsYXAg4oCUIHJlbHkgb24gdG9vbHRpcHMgaW5zdGVhZAoKICBmdW5jdGlvbiBjc3NWYXIobmFtZSkgewogICAgcmV0dXJuIGdldENvbXB1dGVkU3R5bGUoZG9jdW1lbnQuZG9jdW1lbnRFbGVtZW50KS5nZXRQcm9wZXJ0eVZhbHVlKG5hbWUpLnRyaW0oKTsKICB9CgogIGNvbnN0IFNFUklFU19WQVJTID0gWyctLXNlcmllcy0xJywgJy0tc2VyaWVzLTInLCAnLS1zZXJpZXMtMycsICctLXNlcmllcy00JywgJy0tc2VyaWVzLTUnLCAnLS1zZXJpZXMtNicsICctLXNlcmllcy03JywgJy0tc2VyaWVzLTgnXTsKICBmdW5jdGlvbiBzZXJpZXNDb2xvcihpbmRleCkgewogICAgcmV0dXJuIGNzc1ZhcihTRVJJRVNfVkFSU1tpbmRleCAlIFNFUklFU19WQVJTLmxlbmd0aF0pOwogIH0KCiAgZnVuY3Rpb24gYmFzZUdyaWQoKSB7CiAgICByZXR1cm4gewogICAgICBjb2xvcjogY3NzVmFyKCctLWdyaWRsaW5lJyksCiAgICAgIGRyYXdUaWNrczogZmFsc2UsCiAgICB9OwogIH0KICBmdW5jdGlvbiBiYXNlVGlja3MoKSB7CiAgICByZXR1cm4geyBjb2xvcjogY3NzVmFyKCctLXRleHQtbXV0ZWQnKSwgZm9udDogeyBzaXplOiAxMSB9IH07CiAgfQogIGZ1bmN0aW9uIGJhc2VUb29sdGlwKCkgewogICAgcmV0dXJuIHsKICAgICAgYmFja2dyb3VuZENvbG9yOiBjc3NWYXIoJy0tc3VyZmFjZS0xJyksCiAgICAgIHRpdGxlQ29sb3I6IGNzc1ZhcignLS10ZXh0LXByaW1hcnknKSwKICAgICAgYm9keUNvbG9yOiBjc3NWYXIoJy0tdGV4dC1zZWNvbmRhcnknKSwKICAgICAgYm9yZGVyQ29sb3I6IGNzc1ZhcignLS1ib3JkZXInKSwKICAgICAgYm9yZGVyV2lkdGg6IDEsCiAgICAgIGNvcm5lclJhZGl1czogMTAsCiAgICAgIHBhZGRpbmc6IDEyLAogICAgICBib3hQYWRkaW5nOiA0LAogICAgICB0aXRsZUZvbnQ6IHsgc2l6ZTogMTIsIHdlaWdodDogJzcwMCcgfSwKICAgICAgYm9keUZvbnQ6IHsgc2l6ZTogMTIgfSwKICAgIH07CiAgfQogIGZ1bmN0aW9uIGxhYmVsQ29sb3IoKSB7CiAgICByZXR1cm4gY3NzVmFyKCctLXRleHQtcHJpbWFyeScpOwogIH0KICAvKiogU25hcHB5LCBzdWJ0bGUgbW90aW9uIOKAlCBpbiB0aGUgMTUwLTMwMG1zIHJhbmdlIHRoZSByZWRlc2lnbiBjYWxscyBmb3IsIG5ldmVyIGJvdW5jeS4gKi8KICBmdW5jdGlvbiBiYXNlQW5pbWF0aW9uKCkgewogICAgcmV0dXJuIHsgZHVyYXRpb246IDI4MCwgZWFzaW5nOiAnZWFzZU91dFF1YXJ0JyB9OwogIH0KCiAgZnVuY3Rpb24gZGVzdHJveShjYW52YXNJZCkgewogICAgaWYgKHJlZ2lzdHJ5LmhhcyhjYW52YXNJZCkpIHsKICAgICAgcmVnaXN0cnkuZ2V0KGNhbnZhc0lkKS5kZXN0cm95KCk7CiAgICAgIHJlZ2lzdHJ5LmRlbGV0ZShjYW52YXNJZCk7CiAgICB9CiAgfQoKICAvKiogTXVsdGktc2VyaWVzIGxpbmUgY2hhcnQgKGUuZy4gd2Vla2x5IHRyZW5kIHBlciBwbGF0Zm9ybSkuIE9uZSBzZXJpZXMgbmVlZHMgbm8gbGVnZW5kIGJveC4KICAgICAgUGVyLXBvaW50IHZhbHVlIGxhYmVscyBhcmUgc2hvd24gb25seSBmb3IgYSBzaW5nbGUgc2VyaWVzIOKAlCB3aXRoIHNldmVyYWwgc2VyaWVzIG92ZXJsYWlkLAogICAgICBsYWJlbGluZyBldmVyeSBwb2ludCB3b3VsZCBvdmVybGFwLCBzbyB0aG9zZSByZWx5IG9uIHRoZSAoc3RpbGwtcHJlc2VudCkgaG92ZXIgdG9vbHRpcC4gKi8KICBmdW5jdGlvbiB0cmVuZENoYXJ0KGNhbnZhc0lkLCB7IGxhYmVscywgc2VyaWVzLCBmb3JtYXRWYWx1ZSB9KSB7CiAgICBkZXN0cm95KGNhbnZhc0lkKTsKICAgIGNvbnN0IGN0eCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGNhbnZhc0lkKTsKICAgIGlmICghY3R4KSByZXR1cm4gbnVsbDsKICAgIGNvbnN0IGZtdCA9IGZvcm1hdFZhbHVlIHx8ICgodikgPT4gRm9ybWF0LnNtYXJ0KHYpKTsKICAgIGNvbnN0IHNob3dMYWJlbHMgPSBzZXJpZXMubGVuZ3RoID09PSAxICYmIGxhYmVscy5sZW5ndGggPD0gTUFYX0xBQkVMRURfSVRFTVM7CgogICAgY29uc3QgZGF0YXNldHMgPSBzZXJpZXMubWFwKChzLCBpKSA9PiAoewogICAgICBsYWJlbDogcy5sYWJlbCwKICAgICAgZGF0YTogcy5kYXRhLAogICAgICBib3JkZXJDb2xvcjogcy5jb2xvciB8fCBzZXJpZXNDb2xvcihpKSwKICAgICAgYmFja2dyb3VuZENvbG9yOiBzLmNvbG9yIHx8IHNlcmllc0NvbG9yKGkpLAogICAgICBib3JkZXJXaWR0aDogMiwKICAgICAgcG9pbnRSYWRpdXM6IHNob3dMYWJlbHMgPyAzIDogMCwKICAgICAgcG9pbnRIb3ZlclJhZGl1czogNCwKICAgICAgcG9pbnRIaXRSYWRpdXM6IDEyLAogICAgICB0ZW5zaW9uOiAwLjI1LAogICAgICBmaWxsOiBmYWxzZSwKICAgIH0pKTsKCiAgICBjb25zdCBjaGFydCA9IG5ldyBDaGFydChjdHgsIHsKICAgICAgdHlwZTogJ2xpbmUnLAogICAgICBkYXRhOiB7IGxhYmVscywgZGF0YXNldHMgfSwKICAgICAgb3B0aW9uczogewogICAgICAgIHJlc3BvbnNpdmU6IHRydWUsCiAgICAgICAgbWFpbnRhaW5Bc3BlY3RSYXRpbzogZmFsc2UsCiAgICAgICAgaW50ZXJhY3Rpb246IHsgbW9kZTogJ2luZGV4JywgaW50ZXJzZWN0OiBmYWxzZSB9LAogICAgICAgIGxheW91dDogeyBwYWRkaW5nOiB7IHRvcDogc2hvd0xhYmVscyA/IDIwIDogOCB9IH0sCiAgICAgICAgYW5pbWF0aW9uOiBiYXNlQW5pbWF0aW9uKCksCiAgICAgICAgcGx1Z2luczogewogICAgICAgICAgbGVnZW5kOiB7CiAgICAgICAgICAgIGRpc3BsYXk6IHNlcmllcy5sZW5ndGggPiAxLAogICAgICAgICAgICBwb3NpdGlvbjogJ2JvdHRvbScsCiAgICAgICAgICAgIGxhYmVsczogeyBjb2xvcjogY3NzVmFyKCctLXRleHQtc2Vjb25kYXJ5JyksIHVzZVBvaW50U3R5bGU6IHRydWUsIHBvaW50U3R5bGU6ICdsaW5lJywgYm94V2lkdGg6IDE2LCBwYWRkaW5nOiAxNiwgZm9udDogeyBzaXplOiAxMSB9IH0sCiAgICAgICAgICB9LAogICAgICAgICAgdG9vbHRpcDogeyAuLi5iYXNlVG9vbHRpcCgpLCB1c2VQb2ludFN0eWxlOiB0cnVlIH0sCiAgICAgICAgICBkYXRhbGFiZWxzOiBzaG93TGFiZWxzCiAgICAgICAgICAgID8geyBhbGlnbjogJ3RvcCcsIGFuY2hvcjogJ2VuZCcsIGNvbG9yOiBsYWJlbENvbG9yKCksIGZvbnQ6IHsgc2l6ZTogMTEsIHdlaWdodDogJzYwMCcgfSwgZm9ybWF0dGVyOiAodikgPT4gZm10KHYpIH0KICAgICAgICAgICAgOiB7IGRpc3BsYXk6IGZhbHNlIH0sCiAgICAgICAgfSwKICAgICAgICBzY2FsZXM6IHsKICAgICAgICAgIHg6IHsgZ3JpZDogeyBkaXNwbGF5OiBmYWxzZSB9LCB0aWNrczogYmFzZVRpY2tzKCkgfSwKICAgICAgICAgIHk6IHsgZ3JpZDogYmFzZUdyaWQoKSwgdGlja3M6IGJhc2VUaWNrcygpLCBib3JkZXI6IHsgZGlzcGxheTogZmFsc2UgfSwgYmVnaW5BdFplcm86IHRydWUgfSwKICAgICAgICB9LAogICAgICB9LAogICAgfSk7CiAgICByZWdpc3RyeS5zZXQoY2FudmFzSWQsIGNoYXJ0KTsKICAgIHJldHVybiBjaGFydDsKICB9CgogIC8qKiBTaW5nbGUtbWV0cmljIGJhciBjaGFydCBhY3Jvc3MgcGxhdGZvcm1zIChpZGVudGl0eSBlbmNvZGluZyDigJQgZWFjaCBiYXIgSVMgYSBwbGF0Zm9ybSkuICovCiAgZnVuY3Rpb24gcGxhdGZvcm1CYXJDaGFydChjYW52YXNJZCwgeyBsYWJlbHMsIGRhdGEsIGNvbG9ycywgZm9ybWF0VmFsdWUgfSkgewogICAgZGVzdHJveShjYW52YXNJZCk7CiAgICBjb25zdCBjdHggPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZChjYW52YXNJZCk7CiAgICBpZiAoIWN0eCkgcmV0dXJuIG51bGw7CiAgICBjb25zdCBmbXQgPSBmb3JtYXRWYWx1ZSB8fCAoKHYpID0+IEZvcm1hdC5zbWFydCh2KSk7CiAgICBjb25zdCBzaG93TGFiZWxzID0gbGFiZWxzLmxlbmd0aCA8PSBNQVhfTEFCRUxFRF9JVEVNUzsKCiAgICBjb25zdCBjaGFydCA9IG5ldyBDaGFydChjdHgsIHsKICAgICAgdHlwZTogJ2JhcicsCiAgICAgIGRhdGE6IHsKICAgICAgICBsYWJlbHMsCiAgICAgICAgZGF0YXNldHM6IFsKICAgICAgICAgIHsKICAgICAgICAgICAgZGF0YSwKICAgICAgICAgICAgYmFja2dyb3VuZENvbG9yOiBjb2xvcnMsCiAgICAgICAgICAgIGJvcmRlclJhZGl1czogNCwKICAgICAgICAgICAgbWF4QmFyVGhpY2tuZXNzOiAyOCwKICAgICAgICAgICAgYm9yZGVyU2tpcHBlZDogJ2JvdHRvbScsCiAgICAgICAgICB9LAogICAgICAgIF0sCiAgICAgIH0sCiAgICAgIG9wdGlvbnM6IHsKICAgICAgICByZXNwb25zaXZlOiB0cnVlLAogICAgICAgIG1haW50YWluQXNwZWN0UmF0aW86IGZhbHNlLAogICAgICAgIGxheW91dDogeyBwYWRkaW5nOiB7IHRvcDogc2hvd0xhYmVscyA/IDIwIDogOCB9IH0sCiAgICAgICAgYW5pbWF0aW9uOiBiYXNlQW5pbWF0aW9uKCksCiAgICAgICAgcGx1Z2luczogewogICAgICAgICAgbGVnZW5kOiB7IGRpc3BsYXk6IGZhbHNlIH0sCiAgICAgICAgICB0b29sdGlwOiBiYXNlVG9vbHRpcCgpLAogICAgICAgICAgZGF0YWxhYmVsczogc2hvd0xhYmVscwogICAgICAgICAgICA/IHsgYWxpZ246ICdlbmQnLCBhbmNob3I6ICdlbmQnLCBjb2xvcjogbGFiZWxDb2xvcigpLCBmb250OiB7IHNpemU6IDExLCB3ZWlnaHQ6ICc2MDAnIH0sIGZvcm1hdHRlcjogKHYpID0+IGZtdCh2KSB9CiAgICAgICAgICAgIDogeyBkaXNwbGF5OiBmYWxzZSB9LAogICAgICAgIH0sCiAgICAgICAgc2NhbGVzOiB7CiAgICAgICAgICB4OiB7IGdyaWQ6IHsgZGlzcGxheTogZmFsc2UgfSwgdGlja3M6IGJhc2VUaWNrcygpIH0sCiAgICAgICAgICB5OiB7IGdyaWQ6IGJhc2VHcmlkKCksIHRpY2tzOiBiYXNlVGlja3MoKSwgYm9yZGVyOiB7IGRpc3BsYXk6IGZhbHNlIH0sIGJlZ2luQXRaZXJvOiB0cnVlIH0sCiAgICAgICAgfSwKICAgICAgfSwKICAgIH0pOwogICAgcmVnaXN0cnkuc2V0KGNhbnZhc0lkLCBjaGFydCk7CiAgICByZXR1cm4gY2hhcnQ7CiAgfQoKICAvKiogUGllIGNoYXJ0IChhIGhhbmRmdWwgb2YgY2F0ZWdvcmllcyBvbmx5IOKAlCBlLmcuIENhbXBhaWduIFBlcmZvcm1hbmNlJ3MgQWRzL09yZ2FuaWMgc3BsaXQpLgogICAgICBTbGljZSBsYWJlbHMgc2hvdyBib3RoIHNoYXJlLW9mLXdob2xlIGFuZCB0aGUgYWN0dWFsIHZhbHVlLCBwZXIgdGhlICJubyBob3ZlciByZXF1aXJlZCIgZ29hbC4gKi8KICBmdW5jdGlvbiBwaWVDaGFydChjYW52YXNJZCwgeyBsYWJlbHMsIGRhdGEsIGNvbG9ycywgZm9ybWF0VmFsdWUgfSkgewogICAgZGVzdHJveShjYW52YXNJZCk7CiAgICBjb25zdCBjdHggPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZChjYW52YXNJZCk7CiAgICBpZiAoIWN0eCkgcmV0dXJuIG51bGw7CiAgICBjb25zdCBmbXQgPSBmb3JtYXRWYWx1ZSB8fCAoKHYpID0+IEZvcm1hdC5zbWFydCh2KSk7CiAgICBjb25zdCB0b3RhbCA9IGRhdGEucmVkdWNlKChzdW0sIHYpID0+IHN1bSArICh2IHx8IDApLCAwKTsKCiAgICBjb25zdCBjaGFydCA9IG5ldyBDaGFydChjdHgsIHsKICAgICAgdHlwZTogJ3BpZScsCiAgICAgIGRhdGE6IHsKICAgICAgICBsYWJlbHMsCiAgICAgICAgZGF0YXNldHM6IFt7IGRhdGEsIGJhY2tncm91bmRDb2xvcjogY29sb3JzLCBib3JkZXJDb2xvcjogY3NzVmFyKCctLXN1cmZhY2UtMScpLCBib3JkZXJXaWR0aDogMiB9XSwKICAgICAgfSwKICAgICAgb3B0aW9uczogewogICAgICAgIHJlc3BvbnNpdmU6IHRydWUsCiAgICAgICAgbWFpbnRhaW5Bc3BlY3RSYXRpbzogZmFsc2UsCiAgICAgICAgYW5pbWF0aW9uOiBiYXNlQW5pbWF0aW9uKCksCiAgICAgICAgcGx1Z2luczogewogICAgICAgICAgbGVnZW5kOiB7IGRpc3BsYXk6IHRydWUsIHBvc2l0aW9uOiAnYm90dG9tJywgbGFiZWxzOiB7IGNvbG9yOiBjc3NWYXIoJy0tdGV4dC1zZWNvbmRhcnknKSwgYm94V2lkdGg6IDEyLCBwYWRkaW5nOiAxNiwgZm9udDogeyBzaXplOiAxMSB9IH0gfSwKICAgICAgICAgIHRvb2x0aXA6IGJhc2VUb29sdGlwKCksCiAgICAgICAgICBkYXRhbGFiZWxzOiB7CiAgICAgICAgICAgIGNvbG9yOiAnI2ZmZicsCiAgICAgICAgICAgIGZvbnQ6IHsgc2l6ZTogMTIsIHdlaWdodDogJzcwMCcgfSwKICAgICAgICAgICAgZm9ybWF0dGVyOiAodikgPT4gewogICAgICAgICAgICAgIGNvbnN0IHBjdCA9IHRvdGFsID8gTWF0aC5yb3VuZCgodiAvIHRvdGFsKSAqIDEwMDApIC8gMTAgOiAwOwogICAgICAgICAgICAgIHJldHVybiBgJHtwY3R9JVxuJHtmbXQodil9YDsKICAgICAgICAgICAgfSwKICAgICAgICAgIH0sCiAgICAgICAgfSwKICAgICAgfSwKICAgIH0pOwogICAgcmVnaXN0cnkuc2V0KGNhbnZhc0lkLCBjaGFydCk7CiAgICByZXR1cm4gY2hhcnQ7CiAgfQoKICBmdW5jdGlvbiBkZXN0cm95QWxsKCkgewogICAgWy4uLnJlZ2lzdHJ5LmtleXMoKV0uZm9yRWFjaChkZXN0cm95KTsKICB9CgogIHJldHVybiB7IHRyZW5kQ2hhcnQsIHBsYXRmb3JtQmFyQ2hhcnQsIHBpZUNoYXJ0LCBzZXJpZXNDb2xvciwgZGVzdHJveSwgZGVzdHJveUFsbCB9Owp9KSgpOwoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIERhc2hib2FyZCB0YWI6IGEgbWV0cmljLWZvY3VzZWQgcHJlbWl1bSBCSSBkYXNoYm9hcmQuIEEgc2luZ2xlCiAgIE1ldHJpYyBzZWxlY3RvciAoZHluYW1pY2FsbHkgcG9wdWxhdGVkIGZyb20gd2hhdGV2ZXIgdGhlCiAgIHNlbGVjdGVkIHBsYXRmb3JtJ3MgZGF0YSBhY3R1YWxseSBoYXMg4oCUIG5ldmVyIGhhcmRjb2RlZCkgZHJpdmVzCiAgIHRoZSBLUEkgY2FyZHMsIHdlZWtseSB0cmVuZCwgcGxhdGZvcm0vY2FtcGFpZ24vY29udGVudC10eXBlCiAgIGJyZWFrZG93bnMsIGFuZCB0aGUgVG9wIFBlcmZvcm1pbmcgUG9zdHMgcmFua2luZyB0b2dldGhlcjsKICAgUGxhdGZvcm0vZGF0ZS9jYW1wYWlnbi9jb250ZW50LXR5cGUgZmlsdGVyaW5nIGNvbWVzIGZyb20gdGhlCiAgIHNoYXJlZCBmaWx0ZXIgYmFyLiBFdmVyeSBjaGFydCBzaG93cyBpdHMgdmFsdWVzIGRpcmVjdGx5ICh2aWEKICAgY2hhcnRqcy1wbHVnaW4tZGF0YWxhYmVscykgc28gbm90aGluZyByZXF1aXJlcyBhIGhvdmVyIHRvIHJlYWQuCiAgID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwpjb25zdCBEYXNoYm9hcmQgPSAoKCkgPT4gewogIGxldCByb290OwogIGxldCBtZXRyaWMgPSAndmlld3MnOwogIGxldCBtZXRyaWNPcHRpb25zID0gW107CgogIGZ1bmN0aW9uIG9wdGlvbkZvcihrZXkpIHsKICAgIHJldHVybiBtZXRyaWNPcHRpb25zLmZpbmQoKG0pID0+IG0ua2V5ID09PSBrZXkpOwogIH0KICBmdW5jdGlvbiBtZXRyaWNMYWJlbChrZXkpIHsKICAgIGNvbnN0IG9wdCA9IG9wdGlvbkZvcihrZXkpOwogICAgcmV0dXJuIG9wdCA/IG9wdC5sYWJlbCA6IGtleTsKICB9CiAgZnVuY3Rpb24gbWV0cmljVW5pdChrZXkpIHsKICAgIGNvbnN0IG9wdCA9IG9wdGlvbkZvcihrZXkpOwogICAgcmV0dXJuIG9wdCA/IG9wdC51bml0IDogJ251bWJlcic7CiAgfQogIGZ1bmN0aW9uIGZvcm1hdE1ldHJpY1ZhbHVlKGtleSwgdmFsdWUpIHsKICAgIGNvbnN0IHVuaXQgPSBtZXRyaWNVbml0KGtleSk7CiAgICBpZiAodmFsdWUgPT09IG51bGwgfHwgdmFsdWUgPT09IHVuZGVmaW5lZCkgcmV0dXJuICfigJQnOwogICAgaWYgKHVuaXQgPT09ICdkdXJhdGlvbicpIHJldHVybiBGb3JtYXQuZHVyYXRpb24odmFsdWUpOwogICAgcmV0dXJuIEZvcm1hdC5zbWFydCh2YWx1ZSk7CiAgfQoKICBmdW5jdGlvbiBzaGVsbCgpIHsKICAgIHJvb3QuaW5uZXJIVE1MID0gJyc7CgogICAgY29uc3QgY29udHJvbHMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNvbnRyb2xzLmNsYXNzTmFtZSA9ICdkYXNoYm9hcmQtY29udHJvbHMnOwogICAgY29uc3QgbGFiZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdsYWJlbCcpOwogICAgbGFiZWwudGV4dENvbnRlbnQgPSAnTWV0cmljJzsKICAgIGxhYmVsLnNldEF0dHJpYnV0ZSgnZm9yJywgJ2Rhc2hib2FyZE1ldHJpY1NlbGVjdCcpOwogICAgY29uc3QgbWV0cmljU2VsZWN0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc2VsZWN0Jyk7CiAgICBtZXRyaWNTZWxlY3QuaWQgPSAnZGFzaGJvYXJkTWV0cmljU2VsZWN0JzsKICAgIG1ldHJpY09wdGlvbnMuZm9yRWFjaCgobSkgPT4gewogICAgICBjb25zdCBvcHQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsKICAgICAgb3B0LnZhbHVlID0gbS5rZXk7CiAgICAgIG9wdC50ZXh0Q29udGVudCA9IG0ubGFiZWw7CiAgICAgIGlmIChtLmtleSA9PT0gbWV0cmljKSBvcHQuc2VsZWN0ZWQgPSB0cnVlOwogICAgICBtZXRyaWNTZWxlY3QuYXBwZW5kQ2hpbGQob3B0KTsKICAgIH0pOwogICAgbWV0cmljU2VsZWN0LmFkZEV2ZW50TGlzdGVuZXIoJ2NoYW5nZScsICgpID0+IHsKICAgICAgbWV0cmljID0gbWV0cmljU2VsZWN0LnZhbHVlOwogICAgICByZWZyZXNoRm9yTWV0cmljKCk7CiAgICB9KTsKICAgIGNvbnRyb2xzLmFwcGVuZChsYWJlbCwgbWV0cmljU2VsZWN0KTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQoY29udHJvbHMpOwoKICAgIGNvbnN0IGtwaVRpdGxlID0gdGV4dEVsKCdkaXYnLCAnS2V5IHBlcmZvcm1hbmNlIGluZGljYXRvcnMnLCAnc2VjdGlvbi10aXRsZScpOwogICAgY29uc3Qga3BpR3JpZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAga3BpR3JpZC5jbGFzc05hbWUgPSAnc3RhdC1ncmlkJzsKICAgIGtwaUdyaWQuaWQgPSAna3BpR3JpZCc7CiAgICByb290LmFwcGVuZChrcGlUaXRsZSwga3BpR3JpZCk7CgogICAgY29uc3QgY2hhcnRzVGl0bGUgPSB0ZXh0RWwoJ2RpdicsICdUcmVuZCAmIHBlcmZvcm1hbmNlIGJyZWFrZG93bicsICdzZWN0aW9uLXRpdGxlJyk7CiAgICByb290LmFwcGVuZChjaGFydHNUaXRsZSk7CgogICAgY29uc3QgdHJlbmRDYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0cmVuZENhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgY29uc3QgdHJlbmRIZWFkZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRyZW5kSGVhZGVyLmNsYXNzTmFtZSA9ICdjYXJkLWhlYWRlcic7CiAgICB0cmVuZEhlYWRlci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2gzJywgJ1dlZWtseSBwZXJmb3JtYW5jZScpKTsKICAgIHRyZW5kSGVhZGVyLmZpcnN0Q2hpbGQuaWQgPSAndHJlbmRDYXJkVGl0bGUnOwogICAgdHJlbmRDYXJkLmFwcGVuZENoaWxkKHRyZW5kSGVhZGVyKTsKICAgIGNvbnN0IHRyZW5kQ2hhcnRXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0cmVuZENoYXJ0V3JhcC5jbGFzc05hbWUgPSAnY2hhcnQtd3JhcCB0YWxsJzsKICAgIHRyZW5kQ2hhcnRXcmFwLmlkID0gJ3RyZW5kQ2hhcnRXcmFwJzsKICAgIHRyZW5kQ2hhcnRXcmFwLmlubmVySFRNTCA9ICc8Y2FudmFzIGlkPSJ0cmVuZENhbnZhcyI+PC9jYW52YXM+JzsKICAgIHRyZW5kQ2FyZC5hcHBlbmRDaGlsZCh0cmVuZENoYXJ0V3JhcCk7CiAgICByb290LmFwcGVuZENoaWxkKHRyZW5kQ2FyZCk7CgogICAgY29uc3QgZ3JpZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZ3JpZC5jbGFzc05hbWUgPSAnY2FyZC1ncmlkIGV2ZW4nOwogICAgZ3JpZC5zdHlsZS5tYXJnaW5Ub3AgPSAnMTZweCc7CgogICAgY29uc3QgYnJlYWtkb3duQ2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYnJlYWtkb3duQ2FyZC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICBjb25zdCBicmVha2Rvd25IZWFkZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGJyZWFrZG93bkhlYWRlci5jbGFzc05hbWUgPSAnY2FyZC1oZWFkZXInOwogICAgYnJlYWtkb3duSGVhZGVyLmFwcGVuZENoaWxkKHRleHRFbCgnaDMnLCAnJykpOwogICAgYnJlYWtkb3duSGVhZGVyLmZpcnN0Q2hpbGQuaWQgPSAnYnJlYWtkb3duQ2FyZFRpdGxlJzsKICAgIGJyZWFrZG93bkNhcmQuYXBwZW5kQ2hpbGQoYnJlYWtkb3duSGVhZGVyKTsKICAgIGNvbnN0IGJyZWFrZG93bldyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGJyZWFrZG93bldyYXAuY2xhc3NOYW1lID0gJ2NoYXJ0LXdyYXAnOwogICAgYnJlYWtkb3duV3JhcC5pZCA9ICdicmVha2Rvd25DaGFydFdyYXAnOwogICAgYnJlYWtkb3duV3JhcC5pbm5lckhUTUwgPSAnPGNhbnZhcyBpZD0iYnJlYWtkb3duQ2FudmFzIj48L2NhbnZhcz4nOwogICAgYnJlYWtkb3duQ2FyZC5hcHBlbmRDaGlsZChicmVha2Rvd25XcmFwKTsKCiAgICBjb25zdCBjb250ZW50VHlwZUNhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNvbnRlbnRUeXBlQ2FyZC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICBjb25zdCBjb250ZW50VHlwZUhlYWRlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY29udGVudFR5cGVIZWFkZXIuY2xhc3NOYW1lID0gJ2NhcmQtaGVhZGVyJzsKICAgIGNvbnRlbnRUeXBlSGVhZGVyLmFwcGVuZENoaWxkKHRleHRFbCgnaDMnLCAnJykpOwogICAgY29udGVudFR5cGVIZWFkZXIuZmlyc3RDaGlsZC5pZCA9ICdjb250ZW50VHlwZUNhcmRUaXRsZSc7CiAgICBjb250ZW50VHlwZUNhcmQuYXBwZW5kQ2hpbGQoY29udGVudFR5cGVIZWFkZXIpOwogICAgY29uc3QgY29udGVudFR5cGVXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjb250ZW50VHlwZVdyYXAuY2xhc3NOYW1lID0gJ2NoYXJ0LXdyYXAnOwogICAgY29udGVudFR5cGVXcmFwLmlkID0gJ2NvbnRlbnRUeXBlQ2hhcnRXcmFwJzsKICAgIGNvbnRlbnRUeXBlV3JhcC5pbm5lckhUTUwgPSAnPGNhbnZhcyBpZD0iY29udGVudFR5cGVDYW52YXMiPjwvY2FudmFzPic7CiAgICBjb250ZW50VHlwZUNhcmQuYXBwZW5kQ2hpbGQoY29udGVudFR5cGVXcmFwKTsKCiAgICBncmlkLmFwcGVuZChicmVha2Rvd25DYXJkLCBjb250ZW50VHlwZUNhcmQpOwogICAgcm9vdC5hcHBlbmRDaGlsZChncmlkKTsKCiAgICBjb25zdCB0b3BUaXRsZSA9IHRleHRFbCgnZGl2JywgJ1RvcC1wZXJmb3JtaW5nIHBvc3RzJywgJ3NlY3Rpb24tdGl0bGUnKTsKICAgIGNvbnN0IHRvcENhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRvcENhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgY29uc3QgdG9wSGVhZGVyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0b3BIZWFkZXIuY2xhc3NOYW1lID0gJ2NhcmQtaGVhZGVyJzsKICAgIHRvcEhlYWRlci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2gzJywgJ1JhbmtlZCBieSBzZWxlY3RlZCBtZXRyaWMnKSk7CiAgICB0b3BDYXJkLmFwcGVuZENoaWxkKHRvcEhlYWRlcik7CiAgICBjb25zdCB0YWJsZVdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRhYmxlV3JhcC5jbGFzc05hbWUgPSAndGFibGUtc2Nyb2xsJzsKICAgIHRhYmxlV3JhcC5pZCA9ICd0b3BQb3N0c1RhYmxlJzsKICAgIHRvcENhcmQuYXBwZW5kQ2hpbGQodGFibGVXcmFwKTsKICAgIHJvb3QuYXBwZW5kKHRvcFRpdGxlLCB0b3BDYXJkKTsKICB9CgogIC8qKiBBIHNtYWxsIGxhYmVsK3ZhbHVlIHBhaXIgdXNlZCBpbnNpZGUgdGhlIEJlc3QgUGVyZm9ybWluZyBQb3N0IGNhcmQncyBtZXRyaWNzIGNvbHVtbi4gYHZhcmlhbnRgICgncHJpbWFyeScvJ3NlY29uZGFyeScpIGNvbnRyb2xzIHNpemUgYW5kIHdoZXRoZXIgYSBkaXZpZGVyIHJ1bGUgc2l0cyBhYm92ZSBpdC4gKi8KICBmdW5jdGlvbiBtZXRyaWNCbG9jayhsYWJlbCwgdmFsdWUsIHZhcmlhbnQpIHsKICAgIGNvbnN0IGJsb2NrID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBibG9jay5jbGFzc05hbWUgPSB2YXJpYW50ID09PSAnc2Vjb25kYXJ5JyA/ICdwb3N0LXRpbGUtbWV0cmljLWJsb2NrIHNlY29uZGFyeScgOiAncG9zdC10aWxlLW1ldHJpYy1ibG9jayc7CiAgICBjb25zdCB2YWx1ZUVsID0gdGV4dEVsKCdkaXYnLCB2YWx1ZSwgJ3Bvc3QtdGlsZS1tZXRyaWMtdmFsdWUnKTsKICAgIGJsb2NrLmFwcGVuZCh0ZXh0RWwoJ2RpdicsIGxhYmVsLCAncG9zdC10aWxlLW1ldHJpYy1sYWJlbCcpLCB2YWx1ZUVsKTsKICAgIHJlcXVlc3RBbmltYXRpb25GcmFtZSgoKSA9PiBmaXRTdGF0VmFsdWUodmFsdWVFbCwgMTIpKTsKICAgIHJldHVybiBibG9jazsKICB9CgogIC8qKiBBIGZlYXR1cmVkIGxhbmRzY2FwZSBjYXJkICgzIEtQSS10aWxlLXdpZHRocywgc2FtZSBmaXhlZCBoZWlnaHQgYXMgdGhlIHJlc3Qgb2YgdGhlCiAgICAgIHJvdyk6IHRoZSB0b3AtdGllZCBwb3N0J3MgY2FwdGlvbiAod3JhcHMgdXAgdG8gMyBsaW5lcykgd2l0aCBhIHBsYXRmb3JtLWNvbG9yIGRvdAogICAgICArIHBsYXRmb3JtIG5hbWUgKyBkYXRlIG9uIHRoZSBsZWZ0ICh3aGVuIHRoZXJlJ3MgbW9yZSB0aGFuIG9uZSB0aWUsIHRoZSBleHRyYSBjb3VudAogICAgICBpcyBmb2xkZWQgaW50byB0aGF0IHNhbWUgbWV0YSBsaW5lIHJhdGhlciB0aGFuIGxpc3RpbmcgZXZlcnkgdGllZCBwb3N0LCBzbyB0aGUgY2FyZAogICAgICBuZXZlciBoYXMgdG8gZ3JvdyB0YWxsZXIgdGhhbiBpdHMgbmVpZ2hib3JzKTsgdGhlIHNlbGVjdGVkIG1ldHJpYyAobGFyZ2UpIGFuZAogICAgICBDdXJyZW50IEZvbGxvd2VycyAoc21hbGxlciwgYmVsb3cgYSBkaXZpZGVyKSBzdGFja2VkIGluIGEgbmFycm93IGNvbHVtbiBvbiB0aGUgcmlnaHQuICovCiAgZnVuY3Rpb24gYmVzdFBvc3RzVGlsZShsYWJlbCwgcG9zdHMsIGN1cnJlbnRGb2xsb3dlcnMpIHsKICAgIGNvbnN0IHRpbGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRpbGUuY2xhc3NOYW1lID0gJ3N0YXQtdGlsZSBwb3N0LXRpbGUnOwoKICAgIGNvbnN0IG1haW4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIG1haW4uY2xhc3NOYW1lID0gJ3Bvc3QtdGlsZS1tYWluJzsKICAgIG1haW4uYXBwZW5kQ2hpbGQoc3RhdExhYmVsRWwobGFiZWwsICd0cm9waHknLCAnZ29sZCcpKTsKCiAgICBjb25zdCBoYXNQb3N0cyA9IEJvb2xlYW4ocG9zdHMgJiYgcG9zdHMubGVuZ3RoKTsKICAgIGlmIChoYXNQb3N0cykgewogICAgICBjb25zdCBwbGF0Zm9ybU9wdGlvbnMgPSAod2luZG93Ll9fZmlsdGVyT3B0aW9uc0NhY2hlIHx8IHsgcGxhdGZvcm1zOiBbXSB9KS5wbGF0Zm9ybXM7CiAgICAgIGNvbnN0IHByaW1hcnkgPSBwb3N0c1swXTsKICAgICAgY29uc3QgcGxhdE1ldGEgPSBwbGF0Zm9ybU9wdGlvbnMuZmluZCgocCkgPT4gcC5pZCA9PT0gcHJpbWFyeS5wbGF0Zm9ybSkgfHwgeyBsYWJlbDogcHJpbWFyeS5wbGF0Zm9ybSwgY29sb3I6ICd2YXIoLS1zZXJpZXMtMSknIH07CiAgICAgIGNvbnN0IGNhcHRpb24gPSBwcmltYXJ5LmNhcHRpb24gfHwgJyhubyBjYXB0aW9uKSc7CiAgICAgIGNvbnN0IGNhcHRpb25FbCA9IHRleHRFbCgnZGl2JywgY2FwdGlvbiwgJ3Bvc3QtdGlsZS1jYXB0aW9uJyk7CiAgICAgIGNhcHRpb25FbC50aXRsZSA9IGNhcHRpb247CiAgICAgIG1haW4uYXBwZW5kQ2hpbGQoY2FwdGlvbkVsKTsKICAgICAgY29uc3QgdGllZE5vdGUgPSBwb3N0cy5sZW5ndGggPiAxID8gYCDCtyArJHtwb3N0cy5sZW5ndGggLSAxfSBtb3JlIHRpZWRgIDogJyc7CiAgICAgIGNvbnN0IG1ldGFMaW5lID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIG1ldGFMaW5lLmNsYXNzTmFtZSA9ICdwb3N0LXRpbGUtbWV0YSc7CiAgICAgIGNvbnN0IGRvdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgICAgZG90LmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1kb3QnOwogICAgICBkb3Quc3R5bGUuYmFja2dyb3VuZCA9IHBsYXRNZXRhLmNvbG9yOwogICAgICBtZXRhTGluZS5hcHBlbmQoZG90LCBkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShgJHtwbGF0TWV0YS5sYWJlbH0gwrcgJHtGb3JtYXQuZGF0ZShwcmltYXJ5LnB1Ymxpc2hfZGF0ZSl9JHt0aWVkTm90ZX1gKSk7CiAgICAgIG1haW4uYXBwZW5kQ2hpbGQobWV0YUxpbmUpOwogICAgfSBlbHNlIHsKICAgICAgbWFpbi5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsICdObyBkYXRhIHlldCcsICdwb3N0LXRpbGUtY2FwdGlvbiBtdXRlZCcpKTsKICAgIH0KICAgIHRpbGUuYXBwZW5kQ2hpbGQobWFpbik7CgogICAgY29uc3QgaGFzRm9sbG93ZXJzID0gY3VycmVudEZvbGxvd2VycyAhPT0gbnVsbCAmJiBjdXJyZW50Rm9sbG93ZXJzICE9PSB1bmRlZmluZWQ7CiAgICBpZiAoaGFzUG9zdHMgfHwgaGFzRm9sbG93ZXJzKSB7CiAgICAgIHRpbGUuYXBwZW5kQ2hpbGQoT2JqZWN0LmFzc2lnbihkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKSwgeyBjbGFzc05hbWU6ICdwb3N0LXRpbGUtZGl2aWRlcicgfSkpOwogICAgICBjb25zdCBtZXRyaWNzID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIG1ldHJpY3MuY2xhc3NOYW1lID0gJ3Bvc3QtdGlsZS1tZXRyaWNzJzsKICAgICAgaWYgKGhhc1Bvc3RzKSBtZXRyaWNzLmFwcGVuZENoaWxkKG1ldHJpY0Jsb2NrKG1ldHJpY0xhYmVsKG1ldHJpYyksIGZvcm1hdE1ldHJpY1ZhbHVlKG1ldHJpYywgcG9zdHNbMF0udmFsdWUpLCAncHJpbWFyeScpKTsKICAgICAgaWYgKGhhc0ZvbGxvd2VycykgbWV0cmljcy5hcHBlbmRDaGlsZChtZXRyaWNCbG9jaygnQ3VycmVudCBGb2xsb3dlcnMnLCBGb3JtYXQubnVtYmVyKGN1cnJlbnRGb2xsb3dlcnMpLCAnc2Vjb25kYXJ5JykpOwogICAgICB0aWxlLmFwcGVuZENoaWxkKG1ldHJpY3MpOwogICAgfQogICAgcmV0dXJuIHRpbGU7CiAgfQoKICAvKiogSWNvbi1iYWRnZSArIHRleHQgbGFiZWwgcm93LCBzaGFyZWQgYnkgZXZlcnkgS1BJIHRpbGUgYmVsb3cgKG1hdGNoZXMgdGhlIHJlZmVyZW5jZSBkYXNoYm9hcmQncyBjb2xvcmVkIHBlci1jYXJkIGljb25zKS4gYHZhcmlhbnRgIHBpY2tzIHRoZSBiYWRnZSBjb2xvcjogdjEtdjYgbWFwIHRvIHRoZSBwbGF0Zm9ybSBzZXJpZXMgcGFsZXR0ZSwgJ2dvbGQnIGlzIHJlc2VydmVkIGZvciB0aGUgYnJhbmQtYWNjZW50IHRpbGUuICovCiAgZnVuY3Rpb24gc3RhdExhYmVsRWwodGV4dCwgaWNvbiwgdmFyaWFudCkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgd3JhcC5jbGFzc05hbWUgPSAnc3RhdC1sYWJlbCc7CiAgICBjb25zdCBiYWRnZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgIGJhZGdlLmNsYXNzTmFtZSA9IGBzdGF0LWljb24gJHt2YXJpYW50fWA7CiAgICBiYWRnZS5pbm5lckhUTUwgPSBgPGkgZGF0YS1sdWNpZGU9IiR7aWNvbn0iIHN0eWxlPSJ3aWR0aDoxNnB4O2hlaWdodDoxNnB4OyI+PC9pPmA7CiAgICB3cmFwLmFwcGVuZChiYWRnZSwgZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUodGV4dCkpOwogICAgcmV0dXJuIHdyYXA7CiAgfQoKICBmdW5jdGlvbiBzdGF0VGlsZShsYWJlbCwgdmFsdWUsIGZvcm1hdEZuLCBpY29uLCB2YXJpYW50KSB7CiAgICBjb25zdCB0aWxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0aWxlLmNsYXNzTmFtZSA9ICdzdGF0LXRpbGUnOwogICAgY29uc3QgdmFsdWVFbCA9IHRleHRFbCgnZGl2JywgJycsICdzdGF0LXZhbHVlJyk7CiAgICB0aWxlLmFwcGVuZChzdGF0TGFiZWxFbChsYWJlbCwgaWNvbiwgdmFyaWFudCksIHZhbHVlRWwpOwogICAgY29uc3QgZm10ID0gZm9ybWF0Rm4gfHwgKCh2KSA9PiBGb3JtYXQubnVtYmVyKHYpKTsKICAgIGlmICh0eXBlb2YgdmFsdWUgPT09ICdudW1iZXInICYmIE51bWJlci5pc0Zpbml0ZSh2YWx1ZSkpIHsKICAgICAgYW5pbWF0ZUNvdW50KHZhbHVlRWwsIDAsIHZhbHVlLCA5MDAsIGZtdCk7CiAgICB9IGVsc2UgewogICAgICB2YWx1ZUVsLnRleHRDb250ZW50ID0gZm10KHZhbHVlKTsKICAgICAgcmVxdWVzdEFuaW1hdGlvbkZyYW1lKCgpID0+IGZpdFN0YXRWYWx1ZSh2YWx1ZUVsKSk7CiAgICB9CiAgICByZXR1cm4gdGlsZTsKICB9CgogIC8qKiAiRm9sbG93ZXJzIEdyb3d0aCIgdGlsZTogYW4gYWJzb2x1dGUtZGlmZmVyZW5jZSBzdGF0LXZhbHVlIHBsdXMgYSBwZXJjZW50YWdlIGRlbHRhIGxpbmUgKGFycm93ICsgY29sb3IgZHJpdmVuIGJ5IEZvcm1hdC5kZWx0YUNsYXNzLCBzYW1lIGNvbnZlbnRpb24gYXMgdGhlIENvbXBhcmlzb25zIHBhZ2UncyBzdGF0IHRpbGVzKS4gKi8KICBmdW5jdGlvbiBmb2xsb3dlcnNHcm93dGhUaWxlKGNoYW5nZSwgY2hhbmdlUGN0KSB7CiAgICBjb25zdCB0aWxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0aWxlLmNsYXNzTmFtZSA9ICdzdGF0LXRpbGUnOwogICAgY29uc3QgdmFsdWVFbCA9IHRleHRFbCgnZGl2JywgJycsICdzdGF0LXZhbHVlJyk7CiAgICB0aWxlLmFwcGVuZChzdGF0TGFiZWxFbCgnRm9sbG93ZXJzIEdyb3d0aCcsICd0cmVuZGluZy11cCcsICd2MycpLCB2YWx1ZUVsKTsKICAgIGlmIChjaGFuZ2UgPT09IG51bGwgfHwgY2hhbmdlID09PSB1bmRlZmluZWQpIHsKICAgICAgdmFsdWVFbC50ZXh0Q29udGVudCA9ICfigJQnOwogICAgICByZXF1ZXN0QW5pbWF0aW9uRnJhbWUoKCkgPT4gZml0U3RhdFZhbHVlKHZhbHVlRWwpKTsKICAgIH0gZWxzZSB7CiAgICAgIGFuaW1hdGVDb3VudCh2YWx1ZUVsLCAwLCBjaGFuZ2UsIDkwMCwgKHYpID0+IGAke3YgPiAwID8gJysnIDogJyd9JHtGb3JtYXQubnVtYmVyKE1hdGgucm91bmQodikpfWApOwogICAgfQogICAgY29uc3QgZGVsdGFUZXh0ID0gY2hhbmdlUGN0ID09PSBudWxsIHx8IGNoYW5nZVBjdCA9PT0gdW5kZWZpbmVkID8gJ+KAlCcgOiBGb3JtYXQucGN0KGNoYW5nZVBjdCk7CiAgICB0aWxlLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgZGVsdGFUZXh0LCBgc3RhdC1kZWx0YSAke0Zvcm1hdC5kZWx0YUNsYXNzKGNoYW5nZVBjdCl9YCkpOwogICAgcmV0dXJuIHRpbGU7CiAgfQoKICAvKiogIk5ldyBGb2xsb3dlcnMiIHRpbGU6IGZvbGxvd2VycyBnYWluZWQgd2l0aGluIHRoZSBjdXJyZW50bHkgc2VsZWN0ZWQgZGF0ZSByYW5nZSDigJQgc2hvd3MgIk5vIGZvbGxvd2VyIHVwZGF0ZSIgcmF0aGVyIHRoYW4gMCB3aGVuIG5vdGhpbmcgaXMgY29tcHV0YWJsZSBmb3IgdGhlIHJhbmdlIChwZXIgc3BlYyksIHdoaWNoIGlzIGRpZmZlcmVudCBmcm9tIGEgZ2VudWluZSB6ZXJvLiAqLwogIGZ1bmN0aW9uIG5ld0ZvbGxvd2Vyc1RpbGUobmV3Rm9sbG93ZXJzKSB7CiAgICBjb25zdCB0aWxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0aWxlLmNsYXNzTmFtZSA9ICdzdGF0LXRpbGUnOwogICAgY29uc3QgdmFsdWVFbCA9IHRleHRFbCgnZGl2JywgJycsICdzdGF0LXZhbHVlJyk7CiAgICB0aWxlLmFwcGVuZChzdGF0TGFiZWxFbCgnTmV3IEZvbGxvd2VycycsICd1c2VyLXBsdXMnLCAndjEnKSwgdmFsdWVFbCk7CiAgICBpZiAobmV3Rm9sbG93ZXJzID09PSBudWxsIHx8IG5ld0ZvbGxvd2VycyA9PT0gdW5kZWZpbmVkKSB7CiAgICAgIHZhbHVlRWwudGV4dENvbnRlbnQgPSAnTm8gZm9sbG93ZXIgdXBkYXRlJzsKICAgICAgdmFsdWVFbC5jbGFzc0xpc3QuYWRkKCdzdGF0LXZhbHVlLW11dGVkJyk7CiAgICAgIHJlcXVlc3RBbmltYXRpb25GcmFtZSgoKSA9PiBmaXRTdGF0VmFsdWUodmFsdWVFbCkpOwogICAgfSBlbHNlIHsKICAgICAgYW5pbWF0ZUNvdW50KHZhbHVlRWwsIDAsIG5ld0ZvbGxvd2VycywgOTAwLCAodikgPT4gYCR7diA+IDAgPyAnKycgOiAnJ30ke0Zvcm1hdC5udW1iZXIoTWF0aC5yb3VuZCh2KSl9YCk7CiAgICB9CiAgICByZXR1cm4gdGlsZTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlcktwaXMoc3VtbWFyeSwgZm9sbG93ZXJzKSB7CiAgICBjb25zdCBncmlkID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2twaUdyaWQnKTsKICAgIGlmICghZ3JpZCkgcmV0dXJuOwogICAgZ3JpZC5pbm5lckhUTUwgPSAnJzsKCiAgICBncmlkLmFwcGVuZENoaWxkKHN0YXRUaWxlKCdIaWdoZXN0IFZhbHVlJywgc3VtbWFyeS5oaWdoZXN0LCAodikgPT4gZm9ybWF0TWV0cmljVmFsdWUobWV0cmljLCB2KSwgJ3RyZW5kaW5nLXVwJywgJ3YxJykpOwogICAgZ3JpZC5hcHBlbmRDaGlsZChzdGF0VGlsZSgnQXZlcmFnZSBWYWx1ZScsIHN1bW1hcnkuYXZlcmFnZSwgKHYpID0+IGZvcm1hdE1ldHJpY1ZhbHVlKG1ldHJpYywgdiksICdiYXItY2hhcnQtMicsICd2NCcpKTsKICAgIGdyaWQuYXBwZW5kQ2hpbGQoc3RhdFRpbGUoJ1RvdGFsIFZhbHVlJywgc3VtbWFyeS50b3RhbCwgKHYpID0+IGZvcm1hdE1ldHJpY1ZhbHVlKG1ldHJpYywgdiksICdsYXllcnMnLCAndjUnKSk7CiAgICBncmlkLmFwcGVuZENoaWxkKHN0YXRUaWxlKCdOdW1iZXIgb2YgUG9zdHMnLCBzdW1tYXJ5LnBvc3RDb3VudCwgKHYpID0+IEZvcm1hdC5udW1iZXIodiksICdmaWxlLXRleHQnLCAndjYnKSk7CiAgICBncmlkLmFwcGVuZENoaWxkKHN0YXRUaWxlKCdDdXJyZW50IEZvbGxvd2VycycsIGZvbGxvd2Vycy5jdXJyZW50Rm9sbG93ZXJzLCAodikgPT4gKHYgPT09IG51bGwgfHwgdiA9PT0gdW5kZWZpbmVkID8gJ+KAlCcgOiBGb3JtYXQubnVtYmVyKHYpKSwgJ3VzZXJzJywgJ3YyJykpOwogICAgZ3JpZC5hcHBlbmRDaGlsZChmb2xsb3dlcnNHcm93dGhUaWxlKGZvbGxvd2Vycy5mb2xsb3dlcnNDaGFuZ2UsIGZvbGxvd2Vycy5mb2xsb3dlcnNDaGFuZ2VQY3QpKTsKICAgIGdyaWQuYXBwZW5kQ2hpbGQobmV3Rm9sbG93ZXJzVGlsZShmb2xsb3dlcnMubmV3Rm9sbG93ZXJzKSk7CiAgICBncmlkLmFwcGVuZENoaWxkKGJlc3RQb3N0c1RpbGUoJ0Jlc3QgUGVyZm9ybWluZyBQb3N0Jywgc3VtbWFyeS5iZXN0UG9zdHMsIGZvbGxvd2Vycy5jdXJyZW50Rm9sbG93ZXJzKSk7CiAgfQoKCiAgLyoqIFN3YXBzIGEgY2hhcnQgY2FyZCdzIGNhbnZhcyBmb3IgYW4gZW1wdHktc3RhdGUgbWVzc2FnZSwgb3IgcmVzdG9yZXMgdGhlIGNhbnZhcyDigJQgc2luY2UKICAgICAgcmUtcmVuZGVyaW5nIGEgQ2hhcnQuanMgaW5zdGFuY2UgbmVlZHMgYSBsaXZlIDxjYW52YXM+LCBub3Qgd2hhdGV2ZXIgdGhlIGxhc3QgcmVuZGVyIGxlZnQgdGhlcmUuICovCiAgZnVuY3Rpb24gY2hhcnRPckVtcHR5KHdyYXBJZCwgY2FudmFzSWQsIGhhc0RhdGEsIGVtcHR5TWVzc2FnZSwgcmVuZGVyRm4pIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCh3cmFwSWQpOwogICAgaWYgKCF3cmFwKSByZXR1cm47CiAgICBDaGFydHMuZGVzdHJveShjYW52YXNJZCk7CiAgICBpZiAoIWhhc0RhdGEpIHsKICAgICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgICAgd3JhcC5hcHBlbmRDaGlsZChlbXB0eVN0YXRlKHsgaWNvbjogJ2Jhci1jaGFydC0zJywgbWVzc2FnZTogZW1wdHlNZXNzYWdlIH0pKTsKICAgICAgcmV0dXJuOwogICAgfQogICAgd3JhcC5pbm5lckhUTUwgPSBgPGNhbnZhcyBpZD0iJHtjYW52YXNJZH0iPjwvY2FudmFzPmA7CiAgICByZW5kZXJGbigpOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gcmVuZGVyVHJlbmQoZmlsdGVycykgewogICAgY29uc3QgcGxhdGZvcm1PcHRpb25zID0gKHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZSB8fCB7IHBsYXRmb3JtczogW10gfSkucGxhdGZvcm1zOwogICAgY29uc3QgbUxhYmVsID0gbWV0cmljTGFiZWwobWV0cmljKTsKICAgIGNvbnN0IHBsYXRmb3Jtc1RvRmV0Y2ggPSBmaWx0ZXJzLnBsYXRmb3JtID09PSAnYWxsJyA/IHBsYXRmb3JtT3B0aW9ucy5tYXAoKHApID0+IHAuaWQpIDogW2ZpbHRlcnMucGxhdGZvcm1dOwogICAgY29uc3QgdHJlbmRSZXNwb25zZXMgPSBhd2FpdCBQcm9taXNlLmFsbCgKICAgICAgcGxhdGZvcm1zVG9GZXRjaC5tYXAoKHApID0+CiAgICAgICAgQXBpLnRyZW5kKHsgZGF0ZUZyb206IGZpbHRlcnMuZGF0ZUZyb20sIGRhdGVUbzogZmlsdGVycy5kYXRlVG8sIHBsYXRmb3JtOiBwLCBjYW1wYWlnblR5cGU6IGZpbHRlcnMuY2FtcGFpZ25UeXBlLCBjb250ZW50VHlwZTogZmlsdGVycy5jb250ZW50VHlwZSB9KQogICAgICApCiAgICApOwogICAgY29uc3Qgd2Vla1NldCA9IG5ldyBTZXQoKTsKICAgIHRyZW5kUmVzcG9uc2VzLmZvckVhY2goKHJvd3MpID0+IHJvd3MuZm9yRWFjaCgocikgPT4gd2Vla1NldC5hZGQoci5wZXJpb2QpKSk7CiAgICBjb25zdCB3ZWVrcyA9IFsuLi53ZWVrU2V0XS5zb3J0KCk7CiAgICBjb25zdCBzZXJpZXMgPSBwbGF0Zm9ybXNUb0ZldGNoLm1hcCgocCwgaSkgPT4gewogICAgICBjb25zdCBtZXRhID0gcGxhdGZvcm1PcHRpb25zLmZpbmQoKHBsKSA9PiBwbC5pZCA9PT0gcCkgfHwgeyBsYWJlbDogcCB9OwogICAgICBjb25zdCBieVdlZWsgPSBPYmplY3QuZnJvbUVudHJpZXModHJlbmRSZXNwb25zZXNbaV0ubWFwKChyKSA9PiBbci5wZXJpb2QsIHJbbWV0cmljXV0pKTsKICAgICAgcmV0dXJuIHsgbGFiZWw6IG1ldGEubGFiZWwsIGNvbG9yOiBtZXRhLmNvbG9yLCBkYXRhOiB3ZWVrcy5tYXAoKHcpID0+IChieVdlZWtbd10gPT09IHVuZGVmaW5lZCA/IG51bGwgOiBieVdlZWtbd10pKSB9OwogICAgfSk7CgogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3RyZW5kQ2FyZFRpdGxlJykudGV4dENvbnRlbnQgPQogICAgICBmaWx0ZXJzLnBsYXRmb3JtID09PSAnYWxsJyA/IGBXZWVrbHkgJHttTGFiZWx9IGJ5IFBsYXRmb3JtYCA6IGAke21MYWJlbH0gVHJlbmRgOwoKICAgIGNoYXJ0T3JFbXB0eSgndHJlbmRDaGFydFdyYXAnLCAndHJlbmRDYW52YXMnLCB3ZWVrcy5sZW5ndGggPiAwLCAnTm8gZGF0YSBpbiB0aGlzIHJhbmdlIHlldC4nLCAoKSA9PiB7CiAgICAgIENoYXJ0cy50cmVuZENoYXJ0KCd0cmVuZENhbnZhcycsIHsgbGFiZWxzOiB3ZWVrcy5tYXAoRm9ybWF0LmRhdGUpLCBzZXJpZXMsIGZvcm1hdFZhbHVlOiAodikgPT4gZm9ybWF0TWV0cmljVmFsdWUobWV0cmljLCB2KSB9KTsKICAgIH0pOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gcmVuZGVyQnJlYWtkb3duKGZpbHRlcnMpIHsKICAgIGNvbnN0IG1MYWJlbCA9IG1ldHJpY0xhYmVsKG1ldHJpYyk7CiAgICBjb25zdCB0aXRsZUVsID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2JyZWFrZG93bkNhcmRUaXRsZScpOwoKICAgIGlmIChmaWx0ZXJzLnBsYXRmb3JtID09PSAnYWxsJykgewogICAgICB0aXRsZUVsLnRleHRDb250ZW50ID0gYFBsYXRmb3JtIENvbXBhcmlzb24g4oCUICR7bUxhYmVsfWA7CiAgICAgIGNvbnN0IGJyZWFrZG93biA9IGF3YWl0IEFwaS5wbGF0Zm9ybUJyZWFrZG93bihmaWx0ZXJzKTsKICAgICAgY29uc3Qgc29ydGVkID0gYnJlYWtkb3duLmZpbHRlcigocCkgPT4gcFttZXRyaWNdICE9PSBudWxsICYmIHBbbWV0cmljXSAhPT0gdW5kZWZpbmVkKS5zb3J0KChhLCBiKSA9PiBiW21ldHJpY10gLSBhW21ldHJpY10pOwogICAgICBjaGFydE9yRW1wdHkoJ2JyZWFrZG93bkNoYXJ0V3JhcCcsICdicmVha2Rvd25DYW52YXMnLCBzb3J0ZWQubGVuZ3RoID4gMCwgJ05vIGRhdGEgaW4gdGhpcyByYW5nZSB5ZXQuJywgKCkgPT4gewogICAgICAgIENoYXJ0cy5wbGF0Zm9ybUJhckNoYXJ0KCdicmVha2Rvd25DYW52YXMnLCB7CiAgICAgICAgICBsYWJlbHM6IHNvcnRlZC5tYXAoKHApID0+IHAubGFiZWwpLAogICAgICAgICAgZGF0YTogc29ydGVkLm1hcCgocCkgPT4gcFttZXRyaWNdKSwKICAgICAgICAgIGNvbG9yczogc29ydGVkLm1hcCgocCkgPT4gcC5jb2xvciksCiAgICAgICAgICBmb3JtYXRWYWx1ZTogKHYpID0+IGZvcm1hdE1ldHJpY1ZhbHVlKG1ldHJpYywgdiksCiAgICAgICAgfSk7CiAgICAgIH0pOwogICAgfSBlbHNlIHsKICAgICAgdGl0bGVFbC50ZXh0Q29udGVudCA9IGBDYW1wYWlnbiBQZXJmb3JtYW5jZSDigJQgJHttTGFiZWx9YDsKICAgICAgY29uc3QgY2FtcGFpZ25zID0gYXdhaXQgQXBpLmNhbXBhaWduQnJlYWtkb3duKGZpbHRlcnMpOwogICAgICBjb25zdCB3aXRoVmFsdWUgPSBjYW1wYWlnbnMuZmlsdGVyKChjKSA9PiBjW21ldHJpY10gIT09IG51bGwgJiYgY1ttZXRyaWNdICE9PSB1bmRlZmluZWQgJiYgY1ttZXRyaWNdID4gMCk7CiAgICAgIGNoYXJ0T3JFbXB0eSgnYnJlYWtkb3duQ2hhcnRXcmFwJywgJ2JyZWFrZG93bkNhbnZhcycsIHdpdGhWYWx1ZS5sZW5ndGggPiAwLCAnTm8gY2FtcGFpZ24gZGF0YSBpbiB0aGlzIHJhbmdlIHlldC4nLCAoKSA9PiB7CiAgICAgICAgQ2hhcnRzLnBpZUNoYXJ0KCdicmVha2Rvd25DYW52YXMnLCB7CiAgICAgICAgICBsYWJlbHM6IHdpdGhWYWx1ZS5tYXAoKGMpID0+IGMuY2FtcGFpZ25fdHlwZSksCiAgICAgICAgICBkYXRhOiB3aXRoVmFsdWUubWFwKChjKSA9PiBjW21ldHJpY10pLAogICAgICAgICAgY29sb3JzOiB3aXRoVmFsdWUubWFwKChfLCBpKSA9PiBDaGFydHMuc2VyaWVzQ29sb3IoaSkpLAogICAgICAgICAgZm9ybWF0VmFsdWU6ICh2KSA9PiBmb3JtYXRNZXRyaWNWYWx1ZShtZXRyaWMsIHYpLAogICAgICAgIH0pOwogICAgICB9KTsKICAgIH0KICB9CgogIGFzeW5jIGZ1bmN0aW9uIHJlbmRlckNvbnRlbnRUeXBlQnJlYWtkb3duKGZpbHRlcnMpIHsKICAgIGNvbnN0IG1MYWJlbCA9IG1ldHJpY0xhYmVsKG1ldHJpYyk7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnY29udGVudFR5cGVDYXJkVGl0bGUnKS50ZXh0Q29udGVudCA9IGBDb250ZW50IFR5cGUgUGVyZm9ybWFuY2Ug4oCUICR7bUxhYmVsfWA7CiAgICBjb25zdCByb3dzID0gYXdhaXQgQXBpLmNvbnRlbnRUeXBlQnJlYWtkb3duKGZpbHRlcnMpOwogICAgY29uc3Qgc29ydGVkID0gcm93cy5maWx0ZXIoKGMpID0+IGNbbWV0cmljXSAhPT0gbnVsbCAmJiBjW21ldHJpY10gIT09IHVuZGVmaW5lZCkuc29ydCgoYSwgYikgPT4gYlttZXRyaWNdIC0gYVttZXRyaWNdKTsKICAgIGNoYXJ0T3JFbXB0eSgnY29udGVudFR5cGVDaGFydFdyYXAnLCAnY29udGVudFR5cGVDYW52YXMnLCBzb3J0ZWQubGVuZ3RoID4gMCwgJ05vIGRhdGEgaW4gdGhpcyByYW5nZSB5ZXQuJywgKCkgPT4gewogICAgICBDaGFydHMucGxhdGZvcm1CYXJDaGFydCgnY29udGVudFR5cGVDYW52YXMnLCB7CiAgICAgICAgbGFiZWxzOiBzb3J0ZWQubWFwKChjKSA9PiBjLmNvbnRlbnRfdHlwZSksCiAgICAgICAgZGF0YTogc29ydGVkLm1hcCgoYykgPT4gY1ttZXRyaWNdKSwKICAgICAgICBjb2xvcnM6IHNvcnRlZC5tYXAoKF8sIGkpID0+IENoYXJ0cy5zZXJpZXNDb2xvcihpKSksCiAgICAgICAgZm9ybWF0VmFsdWU6ICh2KSA9PiBmb3JtYXRNZXRyaWNWYWx1ZShtZXRyaWMsIHYpLAogICAgICB9KTsKICAgIH0pOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gcmVuZGVyVG9wUG9zdHMoZmlsdGVycykgewogICAgY29uc3QgcG9zdHMgPSBhd2FpdCBBcGkudG9wUG9zdHMoeyAuLi5maWx0ZXJzLCBzb3J0Qnk6IG1ldHJpYywgbGltaXQ6IDEwIH0pOwogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd0b3BQb3N0c1RhYmxlJyk7CiAgICBpZiAoIXdyYXApIHJldHVybjsKICAgIGlmICghcG9zdHMubGVuZ3RoKSB7CiAgICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQoZW1wdHlTdGF0ZSh7CiAgICAgICAgaWNvbjogJ3Ryb3BoeScsCiAgICAgICAgdGl0bGU6ICdObyBwb3N0cyBpbiB0aGlzIHJhbmdlIHlldCcsCiAgICAgICAgbWVzc2FnZTogJ1VwbG9hZCBhIHdlZWtseSBleHBvcnQsIG9yIHdpZGVuIHRoZSBkYXRlIHJhbmdlLCB0byBzZWUgdG9wIHBlcmZvcm1lcnMgaGVyZS4nLAogICAgICAgIGFjdGlvbkxhYmVsOiAnVXBsb2FkIGRhdGEnLAogICAgICAgIG9uQWN0aW9uOiAoKSA9PiBkb2N1bWVudC5xdWVyeVNlbGVjdG9yKCcudGFiLWJ0bltkYXRhLXRhYj0idXBsb2FkIl0nKT8uY2xpY2soKSwKICAgICAgfSkpOwogICAgICByZXR1cm47CiAgICB9CiAgICBjb25zdCBwbGF0Zm9ybU9wdGlvbnMgPSAod2luZG93Ll9fZmlsdGVyT3B0aW9uc0NhY2hlIHx8IHsgcGxhdGZvcm1zOiBbXSB9KS5wbGF0Zm9ybXM7CgogICAgY29uc3QgdGFibGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0YWJsZScpOwogICAgdGFibGUuY2xhc3NOYW1lID0gJ2RhdGEtdGFibGUnOwogICAgY29uc3QgdGhlYWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0aGVhZCcpOwogICAgY29uc3QgaGVhZFRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgIGhlYWRUci5hcHBlbmQoCiAgICAgIHRleHRFbCgndGgnLCAnUmFuaycpLAogICAgICB0ZXh0RWwoJ3RoJywgJ0RhdGUnKSwKICAgICAgdGV4dEVsKCd0aCcsICdQbGF0Zm9ybScpLAogICAgICB0ZXh0RWwoJ3RoJywgJ0NhbXBhaWduJyksCiAgICAgIHRleHRFbCgndGgnLCAnQ29udGVudCBUeXBlJyksCiAgICAgIHRleHRFbCgndGgnLCAnQ2FwdGlvbicpLAogICAgICB0ZXh0RWwoJ3RoJywgbWV0cmljTGFiZWwobWV0cmljKSwgJ251bScpCiAgICApOwogICAgaGVhZFRyLmFwcGVuZENoaWxkKHRleHRFbCgndGgnLCAnJykpOwogICAgdGhlYWQuYXBwZW5kQ2hpbGQoaGVhZFRyKTsKCiAgICBjb25zdCB0Ym9keSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3Rib2R5Jyk7CiAgICBwb3N0cy5mb3JFYWNoKChwLCBpKSA9PiB7CiAgICAgIGNvbnN0IHRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgY29uc3QgbWV0YSA9IHBsYXRmb3JtT3B0aW9ucy5maW5kKChwbCkgPT4gcGwuaWQgPT09IHAucGxhdGZvcm0pIHx8IHsgbGFiZWw6IHAucGxhdGZvcm0sIGNvbG9yOiAnIzk5OScgfTsKICAgICAgY29uc3QgcGxhdGZvcm1UZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIGNvbnN0IHBpbGwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICAgIHBpbGwuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLXBpbGwnOwogICAgICBjb25zdCBkb3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICAgIGRvdC5jbGFzc05hbWUgPSAncGxhdGZvcm0tZG90JzsKICAgICAgZG90LnN0eWxlLmJhY2tncm91bmQgPSBtZXRhLmNvbG9yOwogICAgICBwaWxsLmFwcGVuZChkb3QsIGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKG1ldGEubGFiZWwpKTsKICAgICAgcGxhdGZvcm1UZC5hcHBlbmRDaGlsZChwaWxsKTsKCiAgICAgIGNvbnN0IGNhcHRpb24gPSBwLmNhcHRpb24gfHwgJyhubyBjYXB0aW9uKSc7CiAgICAgIGNvbnN0IHRydW5jYXRlZCA9IGNhcHRpb24ubGVuZ3RoID4gNjAgPyBgJHtjYXB0aW9uLnNsaWNlKDAsIDYwKX3igKZgIDogY2FwdGlvbjsKICAgICAgY29uc3QgY2FwdGlvblRkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgaWYgKHAucG9zdGluZ19saW5rKSB7CiAgICAgICAgY29uc3QgbGluayA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2EnKTsKICAgICAgICBsaW5rLmNsYXNzTmFtZSA9ICdjYXB0aW9uLWxpbmsnOwogICAgICAgIGxpbmsuaHJlZiA9IHAucG9zdGluZ19saW5rOwogICAgICAgIGxpbmsudGFyZ2V0ID0gJ19ibGFuayc7CiAgICAgICAgbGluay5yZWwgPSAnbm9vcGVuZXIgbm9yZWZlcnJlcic7CiAgICAgICAgbGluay50aXRsZSA9IGNhcHRpb247CiAgICAgICAgbGluay5hcHBlbmRDaGlsZChkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZSh0cnVuY2F0ZWQpKTsKICAgICAgICBjYXB0aW9uVGQuYXBwZW5kQ2hpbGQobGluayk7CiAgICAgIH0gZWxzZSB7CiAgICAgICAgY2FwdGlvblRkLmFwcGVuZENoaWxkKGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKHRydW5jYXRlZCkpOwogICAgICAgIGNhcHRpb25UZC50aXRsZSA9IGNhcHRpb247CiAgICAgIH0KCiAgICAgIHRyLmFwcGVuZCgKICAgICAgICB0ZXh0RWwoJ3RkJywgYCMke2kgKyAxfWApLAogICAgICAgIHRleHRFbCgndGQnLCBGb3JtYXQuZGF0ZShwLnB1Ymxpc2hfZGF0ZSkpLAogICAgICAgIHBsYXRmb3JtVGQsCiAgICAgICAgdGV4dEVsKCd0ZCcsIHAuY2FtcGFpZ25fdHlwZSB8fCAn4oCUJyksCiAgICAgICAgdGV4dEVsKCd0ZCcsIHAuY29udGVudF90eXBlIHx8ICfigJQnKSwKICAgICAgICBjYXB0aW9uVGQsCiAgICAgICAgdGV4dEVsKCd0ZCcsIGZvcm1hdE1ldHJpY1ZhbHVlKG1ldHJpYywgcC5tZXRyaWNfdmFsdWUpLCAnbnVtJykKICAgICAgKTsKCiAgICAgIGNvbnN0IGFjdGlvblRkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgY29uc3Qgdmlld0J0biA9IGljb25CdG4oJ2J0bicsICdleWUnLCAnVmlldyBEZXRhaWxzJyk7CiAgICAgIHZpZXdCdG4uZGlzYWJsZWQgPSAhcC5yYXdfcm93X2lkOwogICAgICB2aWV3QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gUmVjb3Jkcy5vcGVuVmlldyhwLnJhd19yb3dfaWQpKTsKICAgICAgYWN0aW9uVGQuYXBwZW5kQ2hpbGQodmlld0J0bik7CiAgICAgIHRyLmFwcGVuZENoaWxkKGFjdGlvblRkKTsKCiAgICAgIHRib2R5LmFwcGVuZENoaWxkKHRyKTsKICAgIH0pOwogICAgdGFibGUuYXBwZW5kKHRoZWFkLCB0Ym9keSk7CiAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgd3JhcC5hcHBlbmRDaGlsZCh0YWJsZSk7CiAgfQoKICAvKiogTWV0cmljIChvciBhbnkgZmlsdGVyKSBjaGFuZ2VkIGJ1dCB0aGUgcGxhdGZvcm0g4oCUIGFuZCB0aGVyZWZvcmUgdGhlIGF2YWlsYWJsZSBtZXRyaWMgbGlzdCDigJQgZGlkbid0OiBubyBuZWVkIHRvIHJlLWZldGNoIG1ldHJpYy1vcHRpb25zIG9yIHJlYnVpbGQgdGhlIHNoZWxsLCBqdXN0IHJlZnJlc2ggdGhlIGRhdGEuICovCiAgYXN5bmMgZnVuY3Rpb24gcmVmcmVzaEZvck1ldHJpYygpIHsKICAgIGNvbnN0IGZpbHRlcnMgPSBTdGF0ZS5nZXRGaWx0ZXJzKCk7CiAgICBjb25zdCBbc3VtbWFyeSwgZm9sbG93ZXJzXSA9IGF3YWl0IFByb21pc2UuYWxsKFsKICAgICAgQXBpLm1ldHJpY1N1bW1hcnkoeyAuLi5maWx0ZXJzLCBtZXRyaWMgfSksCiAgICAgIEFwaS5mb2xsb3dlcnNLcGlzKGZpbHRlcnMpLAogICAgXSk7CiAgICByZW5kZXJLcGlzKHN1bW1hcnksIGZvbGxvd2Vycyk7CiAgICBhd2FpdCBQcm9taXNlLmFsbChbCiAgICAgIHJlbmRlclRyZW5kKGZpbHRlcnMpLCByZW5kZXJCcmVha2Rvd24oZmlsdGVycyksIHJlbmRlckNvbnRlbnRUeXBlQnJlYWtkb3duKGZpbHRlcnMpLCByZW5kZXJUb3BQb3N0cyhmaWx0ZXJzKSwKICAgIF0pOwogIH0KCiAgZnVuY3Rpb24gc2hvd1NrZWxldG9ucygpIHsKICAgIGNvbnN0IGtwaUdyaWQgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgna3BpR3JpZCcpOwogICAgaWYgKGtwaUdyaWQpIHsga3BpR3JpZC5pbm5lckhUTUwgPSAnJzsga3BpR3JpZC5hcHBlbmRDaGlsZChza2VsZXRvblN0YXRHcmlkKDgpKTsgfQogICAgWyd0cmVuZENoYXJ0V3JhcCcsICdicmVha2Rvd25DaGFydFdyYXAnLCAnY29udGVudFR5cGVDaGFydFdyYXAnXS5mb3JFYWNoKChpZCkgPT4gewogICAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoaWQpOwogICAgICBpZiAod3JhcCkgeyB3cmFwLmlubmVySFRNTCA9ICcnOyB3cmFwLmFwcGVuZENoaWxkKHNrZWxldG9uQ2hhcnQoKSk7IH0KICAgIH0pOwogICAgY29uc3QgdG9wUG9zdHMgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndG9wUG9zdHNUYWJsZScpOwogICAgaWYgKHRvcFBvc3RzKSB7IHRvcFBvc3RzLmlubmVySFRNTCA9ICcnOyB0b3BQb3N0cy5hcHBlbmRDaGlsZChza2VsZXRvblJvd3MoNikpOyB9CiAgfQoKICBhc3luYyBmdW5jdGlvbiByZW5kZXIoKSB7CiAgICByb290ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3ZpZXctZGFzaGJvYXJkJyk7CiAgICBjb25zdCBmaWx0ZXJzID0gU3RhdGUuZ2V0RmlsdGVycygpOwogICAgY29uc3QgeyBvcHRpb25zIH0gPSBhd2FpdCBBcGkubWV0cmljT3B0aW9ucyhmaWx0ZXJzLnBsYXRmb3JtKTsKICAgIG1ldHJpY09wdGlvbnMgPSBvcHRpb25zOwogICAgaWYgKCFtZXRyaWNPcHRpb25zLnNvbWUoKG0pID0+IG0ua2V5ID09PSBtZXRyaWMpKSB7CiAgICAgIG1ldHJpYyA9IG1ldHJpY09wdGlvbnMubGVuZ3RoID8gbWV0cmljT3B0aW9uc1swXS5rZXkgOiAndmlld3MnOwogICAgfQogICAgc2hlbGwoKTsKICAgIHNob3dTa2VsZXRvbnMoKTsKICAgIGF3YWl0IHJlZnJlc2hGb3JNZXRyaWMoKTsKICB9CgogIHJldHVybiB7IHJlbmRlciB9Owp9KSgpOwoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIERhdGEgUmVjb3JkcyB0YWI6IGEgQ1JNLXN0eWxlLCBwbGF0Zm9ybS1ncm91cGVkIGJyb3dzZXIgYmFja2VkCiAgIGJ5IHBvc3RzL3Bvc3RfbWV0cmljcyAodGhlIHNhbWUgbm9ybWFsaXplZCBkYXRhIHRoZSBkYXNoYm9hcmQsCiAgIGNvbXBhcmlzb25zLCBhbmQgcmVwb3J0cyByZWFkKSDigJQgIkFsbCBQbGF0Zm9ybXMiIHNob3dzIGEgY29tbW9uCiAgIGNyb3NzLXBsYXRmb3JtIHN1bW1hcnksIGEgc3BlY2lmaWMgcGxhdGZvcm0gc2hvd3Mgb25seSB0aGF0CiAgIHBsYXRmb3JtJ3MgY3VyYXRlZCBtZXRyaWNzLiBFdmVyeSBmaWVsZCBvZiBhIHJlY29yZCAoZXhhY3RseSBhcwogICBpbXBvcnRlZCkgaXMgYWx3YXlzIHJlYWNoYWJsZSB2aWEgVmlldy9FZGl0IHJlZ2FyZGxlc3Mgb2YgdGhlCiAgIHRhYmxlJ3MgY3VyYXRpb24sIHdoaWNoIHJlYWRzIHRoZSByYXdfcm93cyBtaXJyb3IgYW5kLCBvbiBzYXZlLAogICByZS1zeW5jcyBwb3N0cy9wb3N0X21ldHJpY3Mgc28gZXZlcnkgdmlldyByZWZsZWN0cyB0aGUgY2hhbmdlCiAgIGltbWVkaWF0ZWx5LgogICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KY29uc3QgUmVjb3JkcyA9ICgoKSA9PiB7CiAgbGV0IHJvb3Q7CiAgbGV0IHBhZ2UgPSAxOwogIGNvbnN0IHBhZ2VTaXplID0gMjU7CiAgbGV0IHNlYXJjaFZhbHVlID0gJyc7CiAgbGV0IHNlYXJjaERlYm91bmNlID0gbnVsbDsKICBsZXQgbW9kYWxTdGF0ZSA9IG51bGw7IC8vIHsgcmVjb3JkLCB2YWx1ZXM6IFsuLi5dIH0g4oCUIEVkaXQgbW9kYWwgb25seQogIGxldCBjdXJyZW50UmVzdWx0ID0gbnVsbDsgLy8gbGFzdC1sb2FkZWQgcGFnZSwga2VwdCBzbyBzb3J0aW5nIGNhbiByZS1yZW5kZXIgd2l0aG91dCBhIG5ldHdvcmsgcm91bmQtdHJpcAogIGxldCBzb3J0U3RhdGUgPSB7IGtleTogbnVsbCwgZGlyOiAnYXNjJywgdHlwZTogJ3N0cmluZycgfTsKCiAgLyoqIFNvcnRzIGEgY29weSBvZiBgcm93c2AgYnkgYSAocG9zc2libHkgZG90dGVkLCBlLmcuICJtZXRyaWNzLnJlYWNoIikga2V5IHBhdGguIE51bGxzIGFsd2F5cyBzb3J0IGxhc3QgcmVnYXJkbGVzcyBvZiBkaXJlY3Rpb24uICovCiAgZnVuY3Rpb24gc29ydFJvd3Mocm93cywga2V5LCBkaXIsIHR5cGUpIHsKICAgIGNvbnN0IGZhY3RvciA9IGRpciA9PT0gJ2FzYycgPyAxIDogLTE7CiAgICBjb25zdCByZWFkID0gKHJvdykgPT4ga2V5LnNwbGl0KCcuJykucmVkdWNlKChvLCBrKSA9PiAobyA9PT0gbnVsbCB8fCBvID09PSB1bmRlZmluZWQgPyB1bmRlZmluZWQgOiBvW2tdKSwgcm93KTsKICAgIHJldHVybiBbLi4ucm93c10uc29ydCgoYSwgYikgPT4gewogICAgICBjb25zdCBhdiA9IHJlYWQoYSk7CiAgICAgIGNvbnN0IGJ2ID0gcmVhZChiKTsKICAgICAgY29uc3QgYU1pc3NpbmcgPSBhdiA9PT0gbnVsbCB8fCBhdiA9PT0gdW5kZWZpbmVkIHx8IGF2ID09PSAnJzsKICAgICAgY29uc3QgYk1pc3NpbmcgPSBidiA9PT0gbnVsbCB8fCBidiA9PT0gdW5kZWZpbmVkIHx8IGJ2ID09PSAnJzsKICAgICAgaWYgKGFNaXNzaW5nICYmIGJNaXNzaW5nKSByZXR1cm4gMDsKICAgICAgaWYgKGFNaXNzaW5nKSByZXR1cm4gMTsKICAgICAgaWYgKGJNaXNzaW5nKSByZXR1cm4gLTE7CiAgICAgIGlmICh0eXBlID09PSAnbnVtYmVyJykgcmV0dXJuIChhdiAtIGJ2KSAqIGZhY3RvcjsKICAgICAgcmV0dXJuIFN0cmluZyhhdikubG9jYWxlQ29tcGFyZShTdHJpbmcoYnYpKSAqIGZhY3RvcjsKICAgIH0pOwogIH0KCiAgLyoqIEEgPHRoPiB0aGF0IHRvZ2dsZXMgYXNjZW5kaW5nL2Rlc2NlbmRpbmcgb24gY2xpY2sgYW5kIHNob3dzIGFuIGFycm93IG9uIHdoaWNoZXZlciBjb2x1bW4gaXMgYWN0aXZlIOKAlCBzb3J0cyB0aGUgYWxyZWFkeS1sb2FkZWQgcGFnZSBpbnN0YW50bHksIG5vIHJlbG9hZC4gKi8KICBmdW5jdGlvbiBzb3J0YWJsZUhlYWRlcihsYWJlbCwga2V5LCB0eXBlKSB7CiAgICBjb25zdCB0aCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RoJyk7CiAgICBpZiAodHlwZSA9PT0gJ251bWJlcicpIHRoLmNsYXNzTmFtZSA9ICdudW0nOwogICAgdGguY2xhc3NMaXN0LmFkZCgnc29ydGFibGUtdGgnKTsKICAgIGNvbnN0IGlzQWN0aXZlID0gc29ydFN0YXRlLmtleSA9PT0ga2V5OwogICAgdGguYXBwZW5kQ2hpbGQoZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUobGFiZWwpKTsKICAgIHRoLmFwcGVuZENoaWxkKHRleHRFbCgnc3BhbicsIGlzQWN0aXZlID8gKHNvcnRTdGF0ZS5kaXIgPT09ICdhc2MnID8gJyDihpEnIDogJyDihpMnKSA6ICcg4oaVJywgJ3NvcnQtYXJyb3cnKSk7CiAgICB0aC5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsKICAgICAgc29ydFN0YXRlID0geyBrZXksIGRpcjogc29ydFN0YXRlLmtleSA9PT0ga2V5ICYmIHNvcnRTdGF0ZS5kaXIgPT09ICdhc2MnID8gJ2Rlc2MnIDogJ2FzYycsIHR5cGUgfTsKICAgICAgaWYgKGN1cnJlbnRSZXN1bHQpIHJlbmRlclRhYmxlKGN1cnJlbnRSZXN1bHQpOwogICAgfSk7CiAgICByZXR1cm4gdGg7CiAgfQoKICBmdW5jdGlvbiBwbGF0Zm9ybU1ldGEoKSB7CiAgICByZXR1cm4gKHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZSB8fCB7IHBsYXRmb3JtczogW10gfSkucGxhdGZvcm1zOwogIH0KCiAgZnVuY3Rpb24gcGxhdGZvcm1MYWJlbChpZCkgewogICAgY29uc3QgbSA9IHBsYXRmb3JtTWV0YSgpLmZpbmQoKHApID0+IHAuaWQgPT09IGlkKTsKICAgIHJldHVybiBtID8gbS5sYWJlbCA6IGlkOwogIH0KCiAgZnVuY3Rpb24gc2hlbGwoKSB7CiAgICByb290LmlubmVySFRNTCA9ICcnOwogICAgcm9vdC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsICdEYXRhIFJlY29yZHMnLCAnc2VjdGlvbi10aXRsZScpKTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQodGV4dEVsKAogICAgICAnZGl2JywKICAgICAgJ0Jyb3dzZSBieSBwbGF0Zm9ybSB0byBzZWUgb25seSBpdHMgbWV0cmljcywgb3Igc3RheSBvbiBBbGwgUGxhdGZvcm1zIGZvciBhIGNyb3NzLXBsYXRmb3JtIHN1bW1hcnkuIEV2ZXJ5IHJlY29yZCBzdGF5cyBmdWxseSBlZGl0YWJsZSDigJQgVmlldyBvciBFZGl0IGFsd2F5cyBvcGVucyBldmVyeSBmaWVsZCBpbXBvcnRlZCBmcm9tIHRoZSBzcHJlYWRzaGVldCwgbm90IGp1c3Qgd2hhdOKAmXMgaW4gdGhlIHRhYmxlLicsCiAgICAgICdtdXRlZCcKICAgICkpOwoKICAgIGNvbnN0IHRvb2xiYXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRvb2xiYXIuY2xhc3NOYW1lID0gJ3JlY29yZHMtdG9vbGJhcic7CiAgICBjb25zdCBwaWxscyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcGlsbHMuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLWZpbHRlci1waWxscyc7CiAgICBwaWxscy5pZCA9ICdyZWNvcmRzUGxhdGZvcm1QaWxscyc7CiAgICBjb25zdCBzZWFyY2ggPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHNlYXJjaC5jbGFzc05hbWUgPSAncmVjb3Jkcy1zZWFyY2gnOwogICAgY29uc3Qgc2VhcmNoSW5wdXQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOwogICAgc2VhcmNoSW5wdXQudHlwZSA9ICdzZWFyY2gnOwogICAgc2VhcmNoSW5wdXQucGxhY2Vob2xkZXIgPSAnU2VhcmNoIGNhcHRpb25zLCBjYW1wYWlnbnMsIGNvbnRlbnQgdHlwZeKApic7CiAgICBzZWFyY2hJbnB1dC5pZCA9ICdyZWNvcmRzU2VhcmNoSW5wdXQnOwogICAgc2VhcmNoSW5wdXQudmFsdWUgPSBzZWFyY2hWYWx1ZTsKICAgIHNlYXJjaElucHV0LmFkZEV2ZW50TGlzdGVuZXIoJ2lucHV0JywgKCkgPT4gewogICAgICBjbGVhclRpbWVvdXQoc2VhcmNoRGVib3VuY2UpOwogICAgICBzZWFyY2hEZWJvdW5jZSA9IHNldFRpbWVvdXQoKCkgPT4gewogICAgICAgIHNlYXJjaFZhbHVlID0gc2VhcmNoSW5wdXQudmFsdWU7CiAgICAgICAgcGFnZSA9IDE7CiAgICAgICAgbG9hZCgpOwogICAgICB9LCAzMDApOwogICAgfSk7CiAgICBzZWFyY2guYXBwZW5kQ2hpbGQoc2VhcmNoSW5wdXQpOwogICAgdG9vbGJhci5hcHBlbmQocGlsbHMsIHNlYXJjaCk7CiAgICByb290LmFwcGVuZENoaWxkKHRvb2xiYXIpOwoKICAgIGNvbnN0IGV4cG9ydFJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZXhwb3J0Um93LmNsYXNzTmFtZSA9ICdleHBvcnQtYnV0dG9ucyc7CiAgICBjb25zdCBleHBvcnRDc3ZCdG4gPSBpY29uQnRuKCdidG4nLCAnZmlsZS1kb3duJywgJ0V4cG9ydCBDU1YnKTsKICAgIGV4cG9ydENzdkJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHRyaWdnZXJSZWNvcmRzRXhwb3J0KCdjc3YnKSk7CiAgICBjb25zdCBleHBvcnRYbHN4QnRuID0gaWNvbkJ0bignYnRuJywgJ2ZpbGUtc3ByZWFkc2hlZXQnLCAnRXhwb3J0IEV4Y2VsJyk7CiAgICBleHBvcnRYbHN4QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gdHJpZ2dlclJlY29yZHNFeHBvcnQoJ3hsc3gnKSk7CiAgICBleHBvcnRSb3cuYXBwZW5kKGV4cG9ydENzdkJ0biwgZXhwb3J0WGxzeEJ0bik7CiAgICByb290LmFwcGVuZENoaWxkKGV4cG9ydFJvdyk7CgogICAgY29uc3QgY2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY2FyZC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICBjb25zdCB0YWJsZVdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRhYmxlV3JhcC5jbGFzc05hbWUgPSAndGFibGUtc2Nyb2xsJzsKICAgIHRhYmxlV3JhcC5pZCA9ICdyZWNvcmRzVGFibGVXcmFwJzsKICAgIGNhcmQuYXBwZW5kQ2hpbGQodGFibGVXcmFwKTsKICAgIGNvbnN0IHBhZ2VyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBwYWdlci5jbGFzc05hbWUgPSAncGFnaW5hdGlvbi1yb3cnOwogICAgcGFnZXIuaWQgPSAncmVjb3Jkc1BhZ2VyJzsKICAgIGNhcmQuYXBwZW5kQ2hpbGQocGFnZXIpOwogICAgcm9vdC5hcHBlbmRDaGlsZChjYXJkKTsKCiAgICByZW5kZXJQaWxscygpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyUGlsbHMoKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3JlY29yZHNQbGF0Zm9ybVBpbGxzJyk7CiAgICBpZiAoIXdyYXApIHJldHVybjsKICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CiAgICBjb25zdCBjdXJyZW50ID0gU3RhdGUuZ2V0RmlsdGVycygpLnBsYXRmb3JtIHx8ICdhbGwnOwogICAgY29uc3Qgb3B0aW9ucyA9IFt7IGlkOiAnYWxsJywgbGFiZWw6ICdBbGwgUGxhdGZvcm1zJywgY29sb3I6IG51bGwgfSwgLi4ucGxhdGZvcm1NZXRhKCldOwogICAgb3B0aW9ucy5mb3JFYWNoKChvcHQpID0+IHsKICAgICAgY29uc3QgYnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICAgIGJ0bi50eXBlID0gJ2J1dHRvbic7CiAgICAgIGJ0bi5jbGFzc0xpc3QudG9nZ2xlKCdpcy1hY3RpdmUnLCBjdXJyZW50ID09PSBvcHQuaWQpOwogICAgICBpZiAob3B0LmNvbG9yKSB7CiAgICAgICAgY29uc3QgZG90ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICAgIGRvdC5jbGFzc05hbWUgPSAncGxhdGZvcm0tZG90JzsKICAgICAgICBkb3Quc3R5bGUuYmFja2dyb3VuZCA9IG9wdC5jb2xvcjsKICAgICAgICBidG4uYXBwZW5kQ2hpbGQoZG90KTsKICAgICAgfQogICAgICBidG4uYXBwZW5kQ2hpbGQoZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUob3B0LmxhYmVsKSk7CiAgICAgIGJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsKICAgICAgICBpZiAoY3VycmVudCA9PT0gb3B0LmlkKSByZXR1cm47CiAgICAgICAgY29uc3QgZmlsdGVyU2VsZWN0ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZpbHRlclBsYXRmb3JtJyk7CiAgICAgICAgaWYgKGZpbHRlclNlbGVjdCkgZmlsdGVyU2VsZWN0LnZhbHVlID0gb3B0LmlkOwogICAgICAgIHBhZ2UgPSAxOwogICAgICAgIFN0YXRlLnNldEZpbHRlcnMoeyBwbGF0Zm9ybTogb3B0LmlkIH0pOwogICAgICB9KTsKICAgICAgd3JhcC5hcHBlbmRDaGlsZChidG4pOwogICAgfSk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiBsb2FkKCkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdyZWNvcmRzVGFibGVXcmFwJyk7CiAgICBpZiAod3JhcCkgeyB3cmFwLmlubmVySFRNTCA9ICcnOyB3cmFwLmFwcGVuZENoaWxkKHNrZWxldG9uUm93cyg4KSk7IH0KICAgIGNvbnN0IGZpbHRlcnMgPSBTdGF0ZS5nZXRGaWx0ZXJzKCk7CiAgICBjb25zdCByZXN1bHQgPSBhd2FpdCBBcGkucmVjb3Jkc1RhYmxlKHsgLi4uZmlsdGVycywgc2VhcmNoOiBzZWFyY2hWYWx1ZSwgcGFnZSwgcGFnZVNpemUgfSk7CiAgICByZW5kZXJUYWJsZShyZXN1bHQpOwogICAgcmVuZGVyUGFnZXIocmVzdWx0KTsKICB9CgogIC8qKiBSZWNvcmRzIGlzIHNlcnZlci1wYWdpbmF0ZWQvc2VhcmNoZWQsIHNvIGl0cyBleHBvcnQgaXMgYSBkaXJlY3QgbmF2aWdhdGlvbiB0byBhIGJhY2tlbmQgcm91dGUgdGhhdCByZXVzZXMgdGhlIGV4YWN0IHNhbWUgZmlsdGVyLWJ1aWxkaW5nIHRoZSBsaXN0IGVuZHBvaW50IGRvZXMg4oCUIGV4cG9ydHMgdGhlIGZ1bGwgbWF0Y2hpbmcgZGF0YXNldCwgbm90IGp1c3QgdGhlIGN1cnJlbnQgcGFnZS4gKi8KICBmdW5jdGlvbiB0cmlnZ2VyUmVjb3Jkc0V4cG9ydChmb3JtYXQpIHsKICAgIGNvbnN0IGZpbHRlcnMgPSBTdGF0ZS5nZXRGaWx0ZXJzKCk7CiAgICBjb25zdCBwYXJhbXMgPSBuZXcgVVJMU2VhcmNoUGFyYW1zKHsgLi4uZmlsdGVycywgc2VhcmNoOiBzZWFyY2hWYWx1ZSwgZm9ybWF0IH0pOwogICAgd2luZG93LmxvY2F0aW9uLmhyZWYgPSBgL2FwaS9yZWNvcmRzL2V4cG9ydD8ke3BhcmFtcy50b1N0cmluZygpfWA7CiAgfQoKICBmdW5jdGlvbiBjb2x1bW5MYWJlbHNGb3IocmVjb3JkKSB7CiAgICByZXR1cm4gcmVjb3JkLmhlYWRlcnMgJiYgcmVjb3JkLmhlYWRlcnMubGVuZ3RoCiAgICAgID8gcmVjb3JkLmhlYWRlcnMubWFwKChoKSA9PiAoaCAmJiBoLnRyaW0oKSA/IGggOiAnKHVubGFiZWxlZCBjb2x1bW4pJykpCiAgICAgIDogcmVjb3JkLnZhbHVlcy5tYXAoKF8sIGkpID0+IGBDb2x1bW4gJHtpICsgMX1gKTsKICB9CgogIC8qKiBHcm91cHMgYSByYXcgcmVjb3JkJ3MgZmllbGRzIGJ5IHRoZSBxdWFsaWZpZWQgaGVhZGVyJ3MgcGxhdGZvcm0tZ3JvdXAgcHJlZml4CiAgICAgIChlLmcuICJGQUNFQk9PSyDigJQgVmlld3MiKSwgc28gdGhlIFZpZXcvRWRpdCBwb3B1cCByZWFkcyBhcyBzZWN0aW9ucyBpbnN0ZWFkCiAgICAgIG9mIG9uZSBsb25nIGZsYXQgbGlzdCDigJQgZmFsbHMgYmFjayB0byBhIHNpbmdsZSAiRGV0YWlscyIgc2VjdGlvbiBmb3IKICAgICAgaWRlbnRpZmllciBjb2x1bW5zIGFuZCBmb3IgdGhlIHNpbXBsZSAob25lLXBsYXRmb3JtLXBlci1yb3cpIGZvcm1hdC4gKi8KICBmdW5jdGlvbiBncm91cEZpZWxkUm93cyhsYWJlbHMsIHZhbHVlcykgewogICAgY29uc3QgZ3JvdXBzID0gW107CiAgICBjb25zdCBpbmRleCA9IG5ldyBNYXAoKTsKICAgIGxhYmVscy5mb3JFYWNoKChsYWJlbCwgaWR4KSA9PiB7CiAgICAgIGNvbnN0IHNlcElkeCA9IGxhYmVsLmluZGV4T2YoJyDigJQgJyk7CiAgICAgIGNvbnN0IGdyb3VwTmFtZSA9IHNlcElkeCA+PSAwID8gbGFiZWwuc2xpY2UoMCwgc2VwSWR4KSA6ICdEZXRhaWxzJzsKICAgICAgY29uc3QgZmllbGRMYWJlbCA9IHNlcElkeCA+PSAwID8gbGFiZWwuc2xpY2Uoc2VwSWR4ICsgMykgOiBsYWJlbDsKICAgICAgaWYgKCFpbmRleC5oYXMoZ3JvdXBOYW1lKSkgewogICAgICAgIGluZGV4LnNldChncm91cE5hbWUsIHsgZ3JvdXA6IGdyb3VwTmFtZSwgZmllbGRzOiBbXSB9KTsKICAgICAgICBncm91cHMucHVzaChpbmRleC5nZXQoZ3JvdXBOYW1lKSk7CiAgICAgIH0KICAgICAgaW5kZXguZ2V0KGdyb3VwTmFtZSkuZmllbGRzLnB1c2goeyBpZHgsIGxhYmVsOiBmaWVsZExhYmVsIHx8IGBDb2x1bW4gJHtpZHggKyAxfWAsIHZhbHVlOiB2YWx1ZXNbaWR4XSB9KTsKICAgIH0pOwogICAgcmV0dXJuIGdyb3VwczsKICB9CgogIGZ1bmN0aW9uIHBsYXRmb3JtQmFkZ2VzKGlkcykgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgd3JhcC5zdHlsZS5kaXNwbGF5ID0gJ2ZsZXgnOwogICAgd3JhcC5zdHlsZS5mbGV4V3JhcCA9ICd3cmFwJzsKICAgIHdyYXAuc3R5bGUuZ2FwID0gJzRweCc7CiAgICBpZiAoIWlkcy5sZW5ndGgpIHJldHVybiB0ZXh0RWwoJ3NwYW4nLCAn4oCUJywgJ211dGVkJyk7CiAgICBjb25zdCBtZXRhID0gcGxhdGZvcm1NZXRhKCk7CiAgICBpZHMuZm9yRWFjaCgoaWQpID0+IHsKICAgICAgY29uc3QgbSA9IG1ldGEuZmluZCgocCkgPT4gcC5pZCA9PT0gaWQpIHx8IHsgbGFiZWw6IGlkLCBjb2xvcjogJyM5OTknIH07CiAgICAgIGNvbnN0IHBpbGwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICAgIHBpbGwuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLXBpbGwnOwogICAgICBjb25zdCBkb3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICAgIGRvdC5jbGFzc05hbWUgPSAncGxhdGZvcm0tZG90JzsKICAgICAgZG90LnN0eWxlLmJhY2tncm91bmQgPSBtLmNvbG9yOwogICAgICBwaWxsLmFwcGVuZChkb3QsIGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKG0ubGFiZWwpKTsKICAgICAgd3JhcC5hcHBlbmRDaGlsZChwaWxsKTsKICAgIH0pOwogICAgcmV0dXJuIHdyYXA7CiAgfQoKICBmdW5jdGlvbiBzdGF0dXNQaWxsKHN0YXR1cykgewogICAgY29uc3Qgc3BhbiA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgIHNwYW4uY2xhc3NOYW1lID0gYHN0YXR1cy1waWxsICR7c3RhdHVzfWA7CiAgICBzcGFuLnRleHRDb250ZW50ID0gc3RhdHVzID09PSAnZWRpdGVkJyA/ICdFZGl0ZWQnIDogJ09yaWdpbmFsJzsKICAgIHJldHVybiBzcGFuOwogIH0KCiAgZnVuY3Rpb24gbWV0cmljQ2VsbChrZXksIHZhbHVlKSB7CiAgICBpZiAoa2V5ID09PSAncG9zdGluZ19saW5rJykgewogICAgICBjb25zdCB0ZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIHRkLmNsYXNzTmFtZSA9ICdsaW5rLWNlbGwnOwogICAgICBpZiAodmFsdWUpIHsKICAgICAgICBjb25zdCBhID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYScpOwogICAgICAgIGEuaHJlZiA9IHZhbHVlOwogICAgICAgIGEudGFyZ2V0ID0gJ19ibGFuayc7CiAgICAgICAgYS5yZWwgPSAnbm9vcGVuZXIgbm9yZWZlcnJlcic7CiAgICAgICAgYS50ZXh0Q29udGVudCA9ICdPcGVuIOKGlyc7CiAgICAgICAgdGQuYXBwZW5kQ2hpbGQoYSk7CiAgICAgIH0gZWxzZSB7CiAgICAgICAgdGQuYXBwZW5kQ2hpbGQoZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUoJ+KAlCcpKTsKICAgICAgfQogICAgICByZXR1cm4gdGQ7CiAgICB9CiAgICBjb25zdCBkaXNwbGF5ID0ga2V5ID09PSAnd2F0Y2hfdGltZV9zZWNvbmRzJyA/IEZvcm1hdC5kdXJhdGlvbih2YWx1ZSkgOiBGb3JtYXQubnVtYmVyKHZhbHVlKTsKICAgIHJldHVybiB0ZXh0RWwoJ3RkJywgZGlzcGxheSwgJ251bScpOwogIH0KCiAgZnVuY3Rpb24gYWN0aW9uQnV0dG9ucyhyb3csIHBsYXRmb3JtKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB3cmFwLmNsYXNzTmFtZSA9ICdyb3ctYWN0aW9ucyc7CiAgICBjb25zdCB2aWV3QnRuID0gaWNvbkJ0bignYnRuJywgJ2V5ZScsICdWaWV3Jyk7CiAgICB2aWV3QnRuLmRpc2FibGVkID0gIXJvdy5yYXdSb3dJZDsKICAgIHZpZXdCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBvcGVuVmlldyhyb3cucmF3Um93SWQpKTsKICAgIGNvbnN0IGVkaXRCdG4gPSBpY29uQnRuKCdidG4nLCAncGVuY2lsJywgJ0VkaXQnKTsKICAgIGVkaXRCdG4uZGlzYWJsZWQgPSAhcm93LnJhd1Jvd0lkOwogICAgZWRpdEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IG9wZW5FZGl0b3Iocm93LnJhd1Jvd0lkKSk7CiAgICBjb25zdCBkZWxldGVCdG4gPSBpY29uQnRuKCdidG4gZGFuZ2VyJywgJ3RyYXNoLTInLCAnRGVsZXRlJyk7CiAgICBkZWxldGVCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBoYW5kbGVEZWxldGUocm93LCBwbGF0Zm9ybSkpOwogICAgd3JhcC5hcHBlbmQodmlld0J0biwgZWRpdEJ0biwgZGVsZXRlQnRuKTsKICAgIHJldHVybiB3cmFwOwogIH0KCiAgZnVuY3Rpb24gY2FwdGlvbkNlbGwoY2FwdGlvbikgewogICAgY29uc3QgdGV4dCA9IGNhcHRpb24gfHwgJyhubyBjYXB0aW9uKSc7CiAgICByZXR1cm4gdGV4dEVsKCd0ZCcsIHRleHQubGVuZ3RoID4gNzAgPyBgJHt0ZXh0LnNsaWNlKDAsIDcwKX3igKZgIDogdGV4dCk7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJTdW1tYXJ5VGFibGUocmVzdWx0KSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3JlY29yZHNUYWJsZVdyYXAnKTsKICAgIGlmICghcmVzdWx0LnJvd3MubGVuZ3RoKSB7CiAgICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQoZW1wdHlTdGF0ZSh7CiAgICAgICAgaWNvbjogJ2RhdGFiYXNlJywKICAgICAgICB0aXRsZTogJ05vIHJlY29yZHMgbWF0Y2ggdGhlc2UgZmlsdGVycyB5ZXQnLAogICAgICAgIG1lc3NhZ2U6ICdVcGxvYWQgYSB3ZWVrbHkgZXhwb3J0LCBvciB3aWRlbiB0aGUgZGF0ZSByYW5nZSwgdG8gc2VlIHJlY29yZHMgaGVyZS4nLAogICAgICAgIGFjdGlvbkxhYmVsOiAnVXBsb2FkIGRhdGEnLAogICAgICAgIG9uQWN0aW9uOiAoKSA9PiBkb2N1bWVudC5xdWVyeVNlbGVjdG9yKCcudGFiLWJ0bltkYXRhLXRhYj0idXBsb2FkIl0nKT8uY2xpY2soKSwKICAgICAgfSkpOwogICAgICByZXR1cm47CiAgICB9CiAgICBjb25zdCB0YWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICB0YWJsZS5jbGFzc05hbWUgPSAnZGF0YS10YWJsZSc7CiAgICBjb25zdCB0aGVhZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RoZWFkJyk7CiAgICBjb25zdCBoZWFkVHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgaGVhZFRyLmFwcGVuZCgKICAgICAgc29ydGFibGVIZWFkZXIoJ0RhdGUnLCAncHVibGlzaERhdGUnLCAnc3RyaW5nJyksCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdQbGF0Zm9ybXMnLCAncGxhdGZvcm1JZHMuMCcsICdzdHJpbmcnKSwKICAgICAgdGV4dEVsKCd0aCcsICdDYXB0aW9uJyksCiAgICAgIHRleHRFbCgndGgnLCAnQ2FtcGFpZ24nKSwKICAgICAgdGV4dEVsKCd0aCcsICdDb250ZW50IFR5cGUnKSwKICAgICAgdGV4dEVsKCd0aCcsICdTdGF0dXMnKSwKICAgICAgc29ydGFibGVIZWFkZXIoJ0xhc3QgVXBkYXRlZCcsICd1cGRhdGVkQXQnLCAnc3RyaW5nJyksCiAgICAgIHRleHRFbCgndGgnLCAnQWN0aW9ucycpCiAgICApOwogICAgdGhlYWQuYXBwZW5kQ2hpbGQoaGVhZFRyKTsKICAgIGNvbnN0IHRib2R5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGJvZHknKTsKICAgIGNvbnN0IHJvd3MgPSBzb3J0U3RhdGUua2V5ID8gc29ydFJvd3MocmVzdWx0LnJvd3MsIHNvcnRTdGF0ZS5rZXksIHNvcnRTdGF0ZS5kaXIsIHNvcnRTdGF0ZS50eXBlKSA6IHJlc3VsdC5yb3dzOwogICAgcm93cy5mb3JFYWNoKChyKSA9PiB7CiAgICAgIGNvbnN0IHRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgY29uc3QgcGxhdGZvcm1zVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBwbGF0Zm9ybXNUZC5hcHBlbmRDaGlsZChwbGF0Zm9ybUJhZGdlcyhyLnBsYXRmb3JtSWRzKSk7CiAgICAgIGNvbnN0IHN0YXR1c1RkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgc3RhdHVzVGQuYXBwZW5kQ2hpbGQoc3RhdHVzUGlsbChyLnN0YXR1cykpOwogICAgICBjb25zdCBhY3Rpb25zVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBhY3Rpb25zVGQuYXBwZW5kQ2hpbGQoYWN0aW9uQnV0dG9ucyhyLCAnYWxsJykpOwogICAgICB0ci5hcHBlbmQoCiAgICAgICAgdGV4dEVsKCd0ZCcsIEZvcm1hdC5kYXRlKHIucHVibGlzaERhdGUpKSwKICAgICAgICBwbGF0Zm9ybXNUZCwKICAgICAgICBjYXB0aW9uQ2VsbChyLmNhcHRpb24pLAogICAgICAgIHRleHRFbCgndGQnLCByLmNhbXBhaWduVHlwZSB8fCAn4oCUJyksCiAgICAgICAgdGV4dEVsKCd0ZCcsIHIuY29udGVudFR5cGUgfHwgJ+KAlCcpLAogICAgICAgIHN0YXR1c1RkLAogICAgICAgIHRleHRFbCgndGQnLCByLnVwZGF0ZWRBdCksCiAgICAgICAgYWN0aW9uc1RkCiAgICAgICk7CiAgICAgIHRib2R5LmFwcGVuZENoaWxkKHRyKTsKICAgIH0pOwogICAgdGFibGUuYXBwZW5kKHRoZWFkLCB0Ym9keSk7CiAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgd3JhcC5hcHBlbmRDaGlsZCh0YWJsZSk7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJQbGF0Zm9ybVRhYmxlKHJlc3VsdCkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdyZWNvcmRzVGFibGVXcmFwJyk7CiAgICBpZiAoIXJlc3VsdC5yb3dzLmxlbmd0aCkgewogICAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgICB3cmFwLmFwcGVuZENoaWxkKGVtcHR5U3RhdGUoewogICAgICAgIGljb246ICdkYXRhYmFzZScsCiAgICAgICAgdGl0bGU6IGBObyAke3BsYXRmb3JtTGFiZWwocmVzdWx0LnBsYXRmb3JtKX0gcmVjb3JkcyBtYXRjaCB0aGVzZSBmaWx0ZXJzIHlldGAsCiAgICAgICAgbWVzc2FnZTogJ1RyeSBhIGRpZmZlcmVudCBwbGF0Zm9ybSwgb3Igd2lkZW4gdGhlIGRhdGUgcmFuZ2UuJywKICAgICAgfSkpOwogICAgICByZXR1cm47CiAgICB9CiAgICBjb25zdCB0YWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICB0YWJsZS5jbGFzc05hbWUgPSAnZGF0YS10YWJsZSc7CiAgICBjb25zdCB0aGVhZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RoZWFkJyk7CiAgICBjb25zdCBoZWFkVHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgaGVhZFRyLmFwcGVuZChzb3J0YWJsZUhlYWRlcignRGF0ZScsICdwdWJsaXNoRGF0ZScsICdzdHJpbmcnKSwgdGV4dEVsKCd0aCcsICdDYXB0aW9uJyksIHRleHRFbCgndGgnLCAnQ2FtcGFpZ24nKSwgdGV4dEVsKCd0aCcsICdDb250ZW50IFR5cGUnKSk7CiAgICByZXN1bHQuY29sdW1ucy5mb3JFYWNoKChjKSA9PiB7CiAgICAgIGlmIChjLmtleSA9PT0gJ3Bvc3RpbmdfbGluaycpIHsKICAgICAgICBoZWFkVHIuYXBwZW5kQ2hpbGQodGV4dEVsKCd0aCcsIGMubGFiZWwpKTsKICAgICAgfSBlbHNlIHsKICAgICAgICBoZWFkVHIuYXBwZW5kQ2hpbGQoc29ydGFibGVIZWFkZXIoYy5sYWJlbCwgYG1ldHJpY3MuJHtjLmtleX1gLCAnbnVtYmVyJykpOwogICAgICB9CiAgICB9KTsKICAgIGhlYWRUci5hcHBlbmQodGV4dEVsKCd0aCcsICdTdGF0dXMnKSwgdGV4dEVsKCd0aCcsICdBY3Rpb25zJykpOwogICAgdGhlYWQuYXBwZW5kQ2hpbGQoaGVhZFRyKTsKICAgIGNvbnN0IHRib2R5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGJvZHknKTsKICAgIGNvbnN0IHJvd3MgPSBzb3J0U3RhdGUua2V5ID8gc29ydFJvd3MocmVzdWx0LnJvd3MsIHNvcnRTdGF0ZS5rZXksIHNvcnRTdGF0ZS5kaXIsIHNvcnRTdGF0ZS50eXBlKSA6IHJlc3VsdC5yb3dzOwogICAgcm93cy5mb3JFYWNoKChyKSA9PiB7CiAgICAgIGNvbnN0IHRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgdHIuYXBwZW5kKHRleHRFbCgndGQnLCBGb3JtYXQuZGF0ZShyLnB1Ymxpc2hEYXRlKSksIGNhcHRpb25DZWxsKHIuY2FwdGlvbiksIHRleHRFbCgndGQnLCByLmNhbXBhaWduVHlwZSB8fCAn4oCUJyksIHRleHRFbCgndGQnLCByLmNvbnRlbnRUeXBlIHx8ICfigJQnKSk7CiAgICAgIHJlc3VsdC5jb2x1bW5zLmZvckVhY2goKGMpID0+IHRyLmFwcGVuZENoaWxkKG1ldHJpY0NlbGwoYy5rZXksIHIubWV0cmljc1tjLmtleV0pKSk7CiAgICAgIGNvbnN0IHN0YXR1c1RkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgc3RhdHVzVGQuYXBwZW5kQ2hpbGQoc3RhdHVzUGlsbChyLnN0YXR1cykpOwogICAgICB0ci5hcHBlbmRDaGlsZChzdGF0dXNUZCk7CiAgICAgIGNvbnN0IGFjdGlvbnNUZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIGFjdGlvbnNUZC5hcHBlbmRDaGlsZChhY3Rpb25CdXR0b25zKHIsIHJlc3VsdC5wbGF0Zm9ybSkpOwogICAgICB0ci5hcHBlbmRDaGlsZChhY3Rpb25zVGQpOwogICAgICB0Ym9keS5hcHBlbmRDaGlsZCh0cik7CiAgICB9KTsKICAgIHRhYmxlLmFwcGVuZCh0aGVhZCwgdGJvZHkpOwogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGFibGUpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyVGFibGUocmVzdWx0KSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3JlY29yZHNUYWJsZVdyYXAnKTsKICAgIGlmICghd3JhcCkgcmV0dXJuOwogICAgY3VycmVudFJlc3VsdCA9IHJlc3VsdDsKICAgIGlmIChyZXN1bHQucGxhdGZvcm0gPT09ICdhbGwnKSByZW5kZXJTdW1tYXJ5VGFibGUocmVzdWx0KTsKICAgIGVsc2UgcmVuZGVyUGxhdGZvcm1UYWJsZShyZXN1bHQpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyUGFnZXIocmVzdWx0KSB7CiAgICBjb25zdCBwYWdlciA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdyZWNvcmRzUGFnZXInKTsKICAgIGlmICghcGFnZXIpIHJldHVybjsKICAgIHBhZ2VyLmlubmVySFRNTCA9ICcnOwogICAgY29uc3QgdG90YWxQYWdlcyA9IE1hdGgubWF4KDEsIE1hdGguY2VpbChyZXN1bHQudG90YWwgLyByZXN1bHQucGFnZVNpemUpKTsKICAgIGNvbnN0IHByZXZCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIHByZXZCdG4uY2xhc3NOYW1lID0gJ2J0bic7CiAgICBwcmV2QnRuLnRleHRDb250ZW50ID0gJ1ByZXZpb3VzJzsKICAgIHByZXZCdG4uZGlzYWJsZWQgPSByZXN1bHQucGFnZSA8PSAxOwogICAgcHJldkJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsgcGFnZSAtPSAxOyBsb2FkKCk7IH0pOwogICAgY29uc3QgbmV4dEJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgbmV4dEJ0bi5jbGFzc05hbWUgPSAnYnRuJzsKICAgIG5leHRCdG4udGV4dENvbnRlbnQgPSAnTmV4dCc7CiAgICBuZXh0QnRuLmRpc2FibGVkID0gcmVzdWx0LnBhZ2UgPj0gdG90YWxQYWdlczsKICAgIG5leHRCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7IHBhZ2UgKz0gMTsgbG9hZCgpOyB9KTsKICAgIHBhZ2VyLmFwcGVuZChwcmV2QnRuLCB0ZXh0RWwoJ3NwYW4nLCBgUGFnZSAke3Jlc3VsdC5wYWdlfSBvZiAke3RvdGFsUGFnZXN9IOKAlCAke3Jlc3VsdC50b3RhbH0gcmVjb3JkKHMpYCksIG5leHRCdG4pOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gaGFuZGxlRGVsZXRlKHJvdywgcGxhdGZvcm0pIHsKICAgIGNvbnN0IGNhcHRpb24gPSAocm93LmNhcHRpb24gfHwgJyhubyBjYXB0aW9uKScpLnNsaWNlKDAsIDYwKTsKICAgIGNvbnN0IG1lc3NhZ2UgPSBwbGF0Zm9ybSA9PT0gJ2FsbCcKICAgICAgPyBgRGVsZXRlIHRoaXMgZW50aXJlIHJlY29yZCDigJQgIiR7Y2FwdGlvbn0iIOKAlCBhY3Jvc3MgZXZlcnkgcGxhdGZvcm0/IEl0cyBvcmlnaW5hbCBpbXBvcnQgc3RheXMgaW4gVXBsb2FkIEhpc3RvcnksIGJ1dCBpdCB3aWxsIGRpc2FwcGVhciBmcm9tIHRoZSBkYXNoYm9hcmQsIGNvbXBhcmlzb25zLCBhbmQgcmVwb3J0cy5gCiAgICAgIDogYFJlbW92ZSB0aGlzIHJlY29yZCdzICR7cGxhdGZvcm1MYWJlbChwbGF0Zm9ybSl9IGRhdGEg4oCUICIke2NhcHRpb259Ij8gSWYgdGhpcyBpcyBpdHMgb25seSBwbGF0Zm9ybSwgdGhlIHdob2xlIHJlY29yZCB3aWxsIGJlIHJlbW92ZWQgZnJvbSB0aGUgZGFzaGJvYXJkLmA7CiAgICBpZiAoIXdpbmRvdy5jb25maXJtKG1lc3NhZ2UpKSByZXR1cm47CiAgICB0cnkgewogICAgICBpZiAocGxhdGZvcm0gPT09ICdhbGwnKSBhd2FpdCBBcGkuZGVsZXRlUmVjb3JkUG9zdChyb3cucG9zdElkKTsKICAgICAgZWxzZSBhd2FpdCBBcGkuZGVsZXRlUmVjb3JkUGxhdGZvcm0ocm93LnBvc3RJZCwgcGxhdGZvcm0pOwogICAgICBUb2FzdC5zaG93KCdSZWNvcmQgZGVsZXRlZC4nLCAnc3VjY2VzcycpOwogICAgICBhd2FpdCBsb2FkKCk7CiAgICAgIHdpbmRvdy5kaXNwYXRjaEV2ZW50KG5ldyBDdXN0b21FdmVudCgnbHJzOmRhdGEtdXBkYXRlZCcpKTsKICAgIH0gY2F0Y2ggKGVycikgewogICAgICBUb2FzdC5zaG93KGVyci5tZXNzYWdlLCAnZXJyb3InKTsKICAgIH0KICB9CgogIGZ1bmN0aW9uIHJlbW92ZUV4aXN0aW5nT3ZlcmxheSgpIHsKICAgIGNvbnN0IG92ZXJsYXkgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncmVjb3JkTW9kYWxPdmVybGF5Jyk7CiAgICBpZiAob3ZlcmxheSkgb3ZlcmxheS5yZW1vdmUoKTsKICB9CgogIGZ1bmN0aW9uIGNsb3NlTW9kYWwoKSB7CiAgICByZW1vdmVFeGlzdGluZ092ZXJsYXkoKTsKICAgIG1vZGFsU3RhdGUgPSBudWxsOwogIH0KCiAgLy8gT25seSBjbGVhcnMgdGhlIHN0YWxlIERPTSBub2RlIOKAlCBOT1QgbW9kYWxTdGF0ZS4gcmVuZGVyRWRpdE1vZGFsIHJlYWRzCiAgLy8gbW9kYWxTdGF0ZSByaWdodCBhZnRlciBjYWxsaW5nIHRoaXMgdG8gYnVpbGQgdGhlIGZvcm07IGlmIHRoaXMgY2FsbGVkCiAgLy8gdGhlIHJlYWwgY2xvc2VNb2RhbCgpIChhcyBpdCB1c2VkIHRvKSwgdGhhdCByZXNldCBtb2RhbFN0YXRlIHRvIG51bGwgb3V0CiAgLy8gZnJvbSB1bmRlciBpdCBiZWZvcmUgdGhlIHJlYWQsIHdoaWNoIGlzIGV4YWN0bHkgd2h5IEVkaXQgd2FzIGJyb2tlbi4KICBmdW5jdGlvbiBtb2RhbFNoZWxsKHRpdGxlVGV4dCkgewogICAgcmVtb3ZlRXhpc3RpbmdPdmVybGF5KCk7CiAgICBjb25zdCBvdmVybGF5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBvdmVybGF5LmNsYXNzTmFtZSA9ICdtb2RhbC1vdmVybGF5JzsKICAgIG92ZXJsYXkuaWQgPSAncmVjb3JkTW9kYWxPdmVybGF5JzsKICAgIG92ZXJsYXkuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoZSkgPT4geyBpZiAoZS50YXJnZXQgPT09IG92ZXJsYXkpIGNsb3NlTW9kYWwoKTsgfSk7CiAgICBjb25zdCBwYW5lbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcGFuZWwuY2xhc3NOYW1lID0gJ21vZGFsLXBhbmVsIHdpZGUnOwogICAgcGFuZWwuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMicsIHRpdGxlVGV4dCkpOwogICAgb3ZlcmxheS5hcHBlbmRDaGlsZChwYW5lbCk7CiAgICByZXR1cm4geyBvdmVybGF5LCBwYW5lbCB9OwogIH0KCiAgZnVuY3Rpb24gcmVjb3JkU3VidGl0bGUocikgewogICAgcmV0dXJuIGBTaGVldCAiJHtyLnNoZWV0TmFtZX0iLCByb3cgJHtyLnJvd051bWJlcn0ke3IucG9zdElkID8gYCDigJQgbGlua2VkIHRvIGRhc2hib2FyZCBwb3N0ICMke3IucG9zdElkfWAgOiAnIOKAlCBub3QgcGFydCBvZiB0aGUgZGFzaGJvYXJkIChlLmcuIG5lZWRzIGEgdmFsaWQgZGF0ZSknfWA7CiAgfQoKICAvLyAtLS0tLS0tLS0tIFZpZXcgcG9wdXA6IHJlYWQtb25seSwgZXZlcnkgZmllbGQsIGdyb3VwZWQgaW50byBzZWN0aW9ucyAtLS0tLS0tLS0tCiAgYXN5bmMgZnVuY3Rpb24gb3BlblZpZXcoaWQpIHsKICAgIGNvbnN0IHJlY29yZCA9IGF3YWl0IEFwaS5nZXRSZWNvcmQoaWQpOwogICAgY29uc3QgeyBvdmVybGF5LCBwYW5lbCB9ID0gbW9kYWxTaGVsbCgnUmVjb3JkIGRldGFpbHMnKTsKICAgIHBhbmVsLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgcmVjb3JkU3VidGl0bGUocmVjb3JkKSwgJ21vZGFsLXN1YicpKTsKCiAgICBjb25zdCBncm91cHMgPSBncm91cEZpZWxkUm93cyhjb2x1bW5MYWJlbHNGb3IocmVjb3JkKSwgcmVjb3JkLnZhbHVlcyk7CiAgICBncm91cHMuZm9yRWFjaCgoZykgPT4gewogICAgICBjb25zdCBzZWN0aW9uID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIHNlY3Rpb24uY2xhc3NOYW1lID0gJ3JlY29yZC1zZWN0aW9uJzsKICAgICAgc2VjdGlvbi5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2g0JywgZy5ncm91cCkpOwogICAgICBjb25zdCBncmlkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIGdyaWQuY2xhc3NOYW1lID0gJ2Zvcm0tZ3JpZCc7CiAgICAgIGcuZmllbGRzLmZvckVhY2goKGYpID0+IHsKICAgICAgICBjb25zdCBmaWVsZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICAgIGZpZWxkLmNsYXNzTmFtZSA9ICd2aWV3LWZpZWxkJzsKICAgICAgICBmaWVsZC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIGYubGFiZWwsICd2aWV3LWxhYmVsJykpOwogICAgICAgIGNvbnN0IHZhbCA9IGYudmFsdWUgPT09IHVuZGVmaW5lZCB8fCBmLnZhbHVlID09PSBudWxsIHx8IGYudmFsdWUgPT09ICcnID8gJ+KAlCcgOiBTdHJpbmcoZi52YWx1ZSk7CiAgICAgICAgZmllbGQuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCB2YWwsICd2aWV3LXZhbHVlJykpOwogICAgICAgIGdyaWQuYXBwZW5kQ2hpbGQoZmllbGQpOwogICAgICB9KTsKICAgICAgc2VjdGlvbi5hcHBlbmRDaGlsZChncmlkKTsKICAgICAgcGFuZWwuYXBwZW5kQ2hpbGQoc2VjdGlvbik7CiAgICB9KTsKCiAgICBjb25zdCBhY3Rpb25zID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBhY3Rpb25zLmNsYXNzTmFtZSA9ICdtb2RhbC1hY3Rpb25zJzsKICAgIGNvbnN0IGJ0blJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYnRuUm93LmNsYXNzTmFtZSA9ICdidG4tcm93JzsKICAgIGNvbnN0IGNsb3NlQnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBjbG9zZUJ0bi5jbGFzc05hbWUgPSAnYnRuJzsKICAgIGNsb3NlQnRuLnRleHRDb250ZW50ID0gJ0Nsb3NlJzsKICAgIGNsb3NlQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgY2xvc2VNb2RhbCk7CiAgICBjb25zdCBlZGl0QnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBlZGl0QnRuLmNsYXNzTmFtZSA9ICdidG4gcHJpbWFyeSc7CiAgICBlZGl0QnRuLnRleHRDb250ZW50ID0gJ0VkaXQgdGhpcyByZWNvcmQnOwogICAgZWRpdEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IG9wZW5FZGl0b3IocmVjb3JkLmlkKSk7CiAgICBidG5Sb3cuYXBwZW5kKGNsb3NlQnRuLCBlZGl0QnRuKTsKICAgIGFjdGlvbnMuYXBwZW5kQ2hpbGQoYnRuUm93KTsKICAgIHBhbmVsLmFwcGVuZENoaWxkKGFjdGlvbnMpOwoKICAgIGRvY3VtZW50LmJvZHkuYXBwZW5kQ2hpbGQob3ZlcmxheSk7CiAgfQoKICAvLyAtLS0tLS0tLS0tIEVkaXQgcG9wdXA6IGV2ZXJ5IGZpZWxkLCBncm91cGVkIGludG8gc2VjdGlvbnMsIGFsbCBlZGl0YWJsZSAtLS0tLS0tLS0tCiAgYXN5bmMgZnVuY3Rpb24gb3BlbkVkaXRvcihpZCkgewogICAgY29uc3QgcmVjb3JkID0gYXdhaXQgQXBpLmdldFJlY29yZChpZCk7CiAgICBtb2RhbFN0YXRlID0geyByZWNvcmQsIHZhbHVlczogWy4uLnJlY29yZC52YWx1ZXNdIH07CiAgICByZW5kZXJFZGl0TW9kYWwoKTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlckVkaXRNb2RhbCgpIHsKICAgIGNvbnN0IHIgPSBtb2RhbFN0YXRlLnJlY29yZDsKICAgIGNvbnN0IHsgb3ZlcmxheSwgcGFuZWwgfSA9IG1vZGFsU2hlbGwoJ0VkaXQgcmVjb3JkJyk7CiAgICBwYW5lbC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIHJlY29yZFN1YnRpdGxlKHIpLCAnbW9kYWwtc3ViJykpOwoKICAgIGNvbnN0IGdyb3VwcyA9IGdyb3VwRmllbGRSb3dzKGNvbHVtbkxhYmVsc0ZvcihyKSwgbW9kYWxTdGF0ZS52YWx1ZXMpOwogICAgZ3JvdXBzLmZvckVhY2goKGcpID0+IHsKICAgICAgY29uc3Qgc2VjdGlvbiA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBzZWN0aW9uLmNsYXNzTmFtZSA9ICdyZWNvcmQtc2VjdGlvbic7CiAgICAgIHNlY3Rpb24uYXBwZW5kQ2hpbGQodGV4dEVsKCdoNCcsIGcuZ3JvdXApKTsKICAgICAgY29uc3QgZ3JpZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBncmlkLmNsYXNzTmFtZSA9ICdmb3JtLWdyaWQnOwogICAgICBnLmZpZWxkcy5mb3JFYWNoKChmKSA9PiB7CiAgICAgICAgY29uc3QgZmllbGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgICBmaWVsZC5jbGFzc05hbWUgPSAnZm9ybS1maWVsZCc7CiAgICAgICAgZmllbGQuYXBwZW5kQ2hpbGQodGV4dEVsKCdsYWJlbCcsIGYubGFiZWwpKTsKICAgICAgICBjb25zdCBzdHJWYWwgPSBmLnZhbHVlID09PSB1bmRlZmluZWQgfHwgZi52YWx1ZSA9PT0gbnVsbCA/ICcnIDogU3RyaW5nKGYudmFsdWUpOwogICAgICAgIGNvbnN0IGlzTG9uZyA9IHN0clZhbC5sZW5ndGggPiA4MDsKICAgICAgICBjb25zdCBpbnB1dCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoaXNMb25nID8gJ3RleHRhcmVhJyA6ICdpbnB1dCcpOwogICAgICAgIGlmICghaXNMb25nKSBpbnB1dC50eXBlID0gJ3RleHQnOwogICAgICAgIGVsc2UgZmllbGQuc3R5bGUuZ3JpZENvbHVtbiA9ICcxIC8gLTEnOwogICAgICAgIGlucHV0LnZhbHVlID0gc3RyVmFsOwogICAgICAgIGlucHV0LmFkZEV2ZW50TGlzdGVuZXIoJ2lucHV0JywgKCkgPT4geyBtb2RhbFN0YXRlLnZhbHVlc1tmLmlkeF0gPSBpbnB1dC52YWx1ZTsgfSk7CiAgICAgICAgZmllbGQuYXBwZW5kQ2hpbGQoaW5wdXQpOwogICAgICAgIGdyaWQuYXBwZW5kQ2hpbGQoZmllbGQpOwogICAgICB9KTsKICAgICAgc2VjdGlvbi5hcHBlbmRDaGlsZChncmlkKTsKICAgICAgcGFuZWwuYXBwZW5kQ2hpbGQoc2VjdGlvbik7CiAgICB9KTsKCiAgICBjb25zdCBhY3Rpb25zID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBhY3Rpb25zLmNsYXNzTmFtZSA9ICdtb2RhbC1hY3Rpb25zJzsKICAgIGNvbnN0IGVycm9yTXNnID0gdGV4dEVsKCdzcGFuJywgJycsICdtdXRlZCcpOwogICAgZXJyb3JNc2cuaWQgPSAnbW9kYWxFcnJvck1zZyc7CiAgICBjb25zdCBidG5Sb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGJ0blJvdy5jbGFzc05hbWUgPSAnYnRuLXJvdyc7CiAgICBjb25zdCBjYW5jZWxCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIGNhbmNlbEJ0bi5jbGFzc05hbWUgPSAnYnRuJzsKICAgIGNhbmNlbEJ0bi50ZXh0Q29udGVudCA9ICdDYW5jZWwnOwogICAgY2FuY2VsQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgY2xvc2VNb2RhbCk7CiAgICBjb25zdCBzYXZlQnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBzYXZlQnRuLmNsYXNzTmFtZSA9ICdidG4gcHJpbWFyeSc7CiAgICBzYXZlQnRuLnRleHRDb250ZW50ID0gJ1NhdmUgY2hhbmdlcyc7CiAgICBzYXZlQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gc2F2ZUVkaXQoc2F2ZUJ0bikpOwogICAgYnRuUm93LmFwcGVuZChjYW5jZWxCdG4sIHNhdmVCdG4pOwogICAgYWN0aW9ucy5hcHBlbmQoZXJyb3JNc2csIGJ0blJvdyk7CiAgICBwYW5lbC5hcHBlbmRDaGlsZChhY3Rpb25zKTsKCiAgICBkb2N1bWVudC5ib2R5LmFwcGVuZENoaWxkKG92ZXJsYXkpOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gc2F2ZUVkaXQoYnRuKSB7CiAgICBjb25zdCBlcnJvckVsID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ21vZGFsRXJyb3JNc2cnKTsKICAgIGVycm9yRWwudGV4dENvbnRlbnQgPSAnJzsKICAgIGJ0bi5kaXNhYmxlZCA9IHRydWU7CiAgICBidG4udGV4dENvbnRlbnQgPSAnU2F2aW5n4oCmJzsKICAgIHRyeSB7CiAgICAgIGF3YWl0IEFwaS51cGRhdGVSZWNvcmQobW9kYWxTdGF0ZS5yZWNvcmQuaWQsIG1vZGFsU3RhdGUudmFsdWVzKTsKICAgICAgVG9hc3Quc2hvdygnUmVjb3JkIHVwZGF0ZWQuJywgJ3N1Y2Nlc3MnKTsKICAgICAgY2xvc2VNb2RhbCgpOwogICAgICBhd2FpdCBsb2FkKCk7CiAgICAgIHdpbmRvdy5kaXNwYXRjaEV2ZW50KG5ldyBDdXN0b21FdmVudCgnbHJzOmRhdGEtdXBkYXRlZCcpKTsKICAgIH0gY2F0Y2ggKGVycikgewogICAgICBlcnJvckVsLnRleHRDb250ZW50ID0gZXJyLm1lc3NhZ2U7CiAgICAgIGVycm9yRWwuc3R5bGUuY29sb3IgPSAndmFyKC0tc3RhdHVzLWNyaXRpY2FsKSc7CiAgICAgIGJ0bi5kaXNhYmxlZCA9IGZhbHNlOwogICAgICBidG4udGV4dENvbnRlbnQgPSAnU2F2ZSBjaGFuZ2VzJzsKICAgIH0KICB9CgogIGFzeW5jIGZ1bmN0aW9uIHJlbmRlcigpIHsKICAgIHJvb3QgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndmlldy1yZWNvcmRzJyk7CiAgICBwYWdlID0gMTsKICAgIHNoZWxsKCk7CiAgICBhd2FpdCBsb2FkKCk7CiAgfQoKICByZXR1cm4geyByZW5kZXIsIHJlbG9hZDogbG9hZCwgb3BlblZpZXcgfTsKfSkoKTsKCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBDb21wYXJpc29ucyB0YWI6IHdlZWstdnMtd2VlaywgY3VzdG9tIHJhbmdlLCBtb250aGx5LAogICBxdWFydGVybHksIFlURC4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmNvbnN0IENvbXBhcmlzb24gPSAoKCkgPT4gewogIGNvbnN0IE1PREVTID0gWwogICAgeyBrZXk6ICdwbGF0Zm9ybXMnLCBsYWJlbDogJ0FsbCBQbGF0Zm9ybXMnIH0sCiAgICB7IGtleTogJ3dlZWsnLCBsYWJlbDogJ1dlZWsgdnMgV2VlaycgfSwKICAgIHsga2V5OiAnY3VzdG9tJywgbGFiZWw6ICdDdXN0b20gUmFuZ2UnIH0sCiAgICB7IGtleTogJ21vbnRoJywgbGFiZWw6ICdNb250aGx5JyB9LAogICAgeyBrZXk6ICdxdWFydGVyJywgbGFiZWw6ICdRdWFydGVybHknIH0sCiAgICB7IGtleTogJ3l0ZCcsIGxhYmVsOiAnWWVhciB0byBEYXRlJyB9LAogIF07CiAgY29uc3QgTUVUUklDX1JPV1MgPSBbCiAgICB7IGtleTogJ3ZpZXdzJywgbGFiZWw6ICdWaWV3cycgfSwKICAgIHsga2V5OiAncmVhY2gnLCBsYWJlbDogJ1JlYWNoJyB9LAogICAgeyBrZXk6ICdpbXByZXNzaW9ucycsIGxhYmVsOiAnSW1wcmVzc2lvbnMnIH0sCiAgICB7IGtleTogJ2VuZ2FnZW1lbnQnLCBsYWJlbDogJ0VuZ2FnZW1lbnQnIH0sCiAgICB7IGtleTogJ2NsaWNrcycsIGxhYmVsOiAnQ2xpY2tzJyB9LAogICAgeyBrZXk6ICdmb2xsb3dlcnNfZ2FpbmVkJywgbGFiZWw6ICdGb2xsb3dlcnMgR2FpbmVkJyB9LAogICAgeyBrZXk6ICd3YXRjaF90aW1lX3NlY29uZHMnLCBsYWJlbDogJ1dhdGNoIFRpbWUnIH0sCiAgICB7IGtleTogJ3NoYXJlcycsIGxhYmVsOiAnU2hhcmVzJyB9LAogICAgeyBrZXk6ICdjb21tZW50cycsIGxhYmVsOiAnQ29tbWVudHMnIH0sCiAgICB7IGtleTogJ3NhdmVzJywgbGFiZWw6ICdTYXZlcycgfSwKICBdOwoKICBsZXQgbW9kZSA9ICdwbGF0Zm9ybXMnOwogIGxldCByb290OwogIGxldCBwbGF0Zm9ybUNoYXJ0TWV0cmljID0gJ2VuZ2FnZW1lbnQnOwogIGxldCBjYXJkU29ydE1vZGUgPSAnb3ZlcmFsbCc7CiAgbGV0IGNhcmRQbGF0Zm9ybUZpbHRlciA9ICdhbGwnOwoKICBmdW5jdGlvbiBtb25kYXlPZihkYXRlU3RyKSB7CiAgICBjb25zdCBkID0gbmV3IERhdGUoZGF0ZVN0cik7CiAgICBjb25zdCBkYXkgPSBkLmdldERheSgpOwogICAgY29uc3QgZGlmZiA9IGRheSA9PT0gMCA/IDYgOiBkYXkgLSAxOwogICAgZC5zZXREYXRlKGQuZ2V0RGF0ZSgpIC0gZGlmZik7CiAgICByZXR1cm4gZC50b0lTT1N0cmluZygpLnNsaWNlKDAsIDEwKTsKICB9CiAgZnVuY3Rpb24gYWRkRGF5cyhkYXRlU3RyLCBuKSB7CiAgICBjb25zdCBkID0gbmV3IERhdGUoZGF0ZVN0cik7CiAgICBkLnNldERhdGUoZC5nZXREYXRlKCkgKyBuKTsKICAgIHJldHVybiBkLnRvSVNPU3RyaW5nKCkuc2xpY2UoMCwgMTApOwogIH0KCiAgZnVuY3Rpb24gc2hlbGwoKSB7CiAgICByb290LmlubmVySFRNTCA9ICcnOwoKICAgIGNvbnN0IHRhYnMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRhYnMuY2xhc3NOYW1lID0gJ21vZGUtdGFicyc7CiAgICBNT0RFUy5mb3JFYWNoKChtKSA9PiB7CiAgICAgIGNvbnN0IGJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgICBidG4udGV4dENvbnRlbnQgPSBtLmxhYmVsOwogICAgICBidG4udHlwZSA9ICdidXR0b24nOwogICAgICBpZiAobS5rZXkgPT09IG1vZGUpIGJ0bi5jbGFzc0xpc3QuYWRkKCdpcy1hY3RpdmUnKTsKICAgICAgYnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gewogICAgICAgIG1vZGUgPSBtLmtleTsKICAgICAgICBzaGVsbCgpOwogICAgICB9KTsKICAgICAgdGFicy5hcHBlbmRDaGlsZChidG4pOwogICAgfSk7CiAgICByb290LmFwcGVuZENoaWxkKHRhYnMpOwoKICAgIGNvbnN0IGNvbnRyb2xzID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjb250cm9scy5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICBjb250cm9scy5pZCA9ICdjb21wYXJpc29uQ29udHJvbHMnOwogICAgcm9vdC5hcHBlbmRDaGlsZChjb250cm9scyk7CgogICAgY29uc3QgcmVzdWx0cyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcmVzdWx0cy5pZCA9ICdjb21wYXJpc29uUmVzdWx0cyc7CiAgICByb290LmFwcGVuZENoaWxkKHJlc3VsdHMpOwoKICAgIHJlbmRlckNvbnRyb2xzKCk7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJDb250cm9scygpIHsKICAgIGNvbnN0IGNvbnRyb2xzID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2NvbXBhcmlzb25Db250cm9scycpOwogICAgY29udHJvbHMuaW5uZXJIVE1MID0gJyc7CiAgICBjb25zdCByb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHJvdy5jbGFzc05hbWUgPSAnYnRuLXJvdyc7CiAgICByb3cuc3R5bGUuYWxpZ25JdGVtcyA9ICdlbmQnOwoKICAgIGNvbnN0IHRvZGF5ID0gbmV3IERhdGUoKS50b0lTT1N0cmluZygpLnNsaWNlKDAsIDEwKTsKICAgIGNvbnN0IHRoaXNZZWFyID0gbmV3IERhdGUoKS5nZXRGdWxsWWVhcigpOwoKICAgIGlmIChtb2RlID09PSAncGxhdGZvcm1zJykgewogICAgICBjb25zdCBmRnJvbSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7IGZGcm9tLnR5cGUgPSAnZGF0ZSc7IGZGcm9tLmlkID0gJ3BsYXRmb3JtUmVwb3J0RnJvbSc7CiAgICAgIGNvbnN0IGZUbyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7IGZUby50eXBlID0gJ2RhdGUnOyBmVG8uaWQgPSAncGxhdGZvcm1SZXBvcnRUbyc7CiAgICAgIGNvbnN0IGFwcGx5QnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICAgIGFwcGx5QnRuLmNsYXNzTmFtZSA9ICdidG4gcHJpbWFyeSc7CiAgICAgIGFwcGx5QnRuLnR5cGUgPSAnYnV0dG9uJzsKICAgICAgYXBwbHlCdG4udGV4dENvbnRlbnQgPSAnQXBwbHkgUmFuZ2UnOwogICAgICBhcHBseUJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IGxvYWRQbGF0Zm9ybVJlcG9ydCh7IGRhdGVGcm9tOiBmRnJvbS52YWx1ZSwgZGF0ZVRvOiBmVG8udmFsdWUgfSkpOwogICAgICBjb25zdCBjbGVhckJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgICBjbGVhckJ0bi5jbGFzc05hbWUgPSAnYnRuJzsKICAgICAgY2xlYXJCdG4udHlwZSA9ICdidXR0b24nOwogICAgICBjbGVhckJ0bi50ZXh0Q29udGVudCA9ICdBbGwgVGltZSc7CiAgICAgIGNsZWFyQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4geyBmRnJvbS52YWx1ZSA9ICcnOyBmVG8udmFsdWUgPSAnJzsgbG9hZFBsYXRmb3JtUmVwb3J0KHt9KTsgfSk7CiAgICAgIHJvdy5hcHBlbmQoCiAgICAgICAgbGFiZWxlZCgnRnJvbSAob3B0aW9uYWwpJywgZkZyb20pLAogICAgICAgIGxhYmVsZWQoJ1RvIChvcHRpb25hbCknLCBmVG8pLAogICAgICAgIGFwcGx5QnRuLAogICAgICAgIGNsZWFyQnRuCiAgICAgICk7CiAgICAgIGNvbnRyb2xzLmFwcGVuZENoaWxkKHJvdyk7CiAgICAgIGxvYWRQbGF0Zm9ybVJlcG9ydCh7fSk7CiAgICAgIHJldHVybjsKICAgIH0gZWxzZSBpZiAobW9kZSA9PT0gJ3dlZWsnKSB7CiAgICAgIC8vIFdlZWsgQSBpcyBhbHdheXMgdGhlIFByZXZpb3VzIHBlcmlvZCwgV2VlayBCIGlzIGFsd2F5cyB0aGUgQ3VycmVudC9tb3N0IHJlY2VudAogICAgICAvLyBwZXJpb2Qg4oCUIHJ1bkNvbXBhcmUoKSdzIGZpcnN0IGFyZ3VtZW50IGlzIHRoZSAiY3VycmVudCIgc2xvdCBldmVyeSBvdGhlciBtb2RlCiAgICAgIC8vIGluIHRoaXMgdGFiIGZlZWRzIGl0IChwZXJjZW50Q2hhbmdlID0gKGN1cnJlbnQgLSBwcmV2aW91cykgLyBwcmV2aW91cyksIHNvIEIgZ29lcwogICAgICAvLyBpbiBmaXJzdCBhbmQgQSBzZWNvbmQsIHJlZ2FyZGxlc3Mgb2Ygd2hpY2ggY2FsZW5kYXIgd2VlayBpcyBlYXJsaWVyIG9yIGxhdGVyLgogICAgICBjb25zdCB3QSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7IHdBLnR5cGUgPSAnZGF0ZSc7IHdBLnZhbHVlID0gbW9uZGF5T2YoYWRkRGF5cyh0b2RheSwgLTcpKTsKICAgICAgY29uc3Qgd0IgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOyB3Qi50eXBlID0gJ2RhdGUnOyB3Qi52YWx1ZSA9IG1vbmRheU9mKHRvZGF5KTsKICAgICAgcm93LmFwcGVuZChsYWJlbGVkKCdXZWVrIEEgKFByZXZpb3VzKScsIHdBKSwgbGFiZWxlZCgnV2VlayBCIChDdXJyZW50KScsIHdCKSwgcnVuQnRuKCgpID0+IHsKICAgICAgICBjb25zdCByYW5nZUEgPSB7IGZyb206IG1vbmRheU9mKHdBLnZhbHVlKSwgdG86IGFkZERheXMobW9uZGF5T2Yod0EudmFsdWUpLCA2KSB9OwogICAgICAgIGNvbnN0IHJhbmdlQiA9IHsgZnJvbTogbW9uZGF5T2Yod0IudmFsdWUpLCB0bzogYWRkRGF5cyhtb25kYXlPZih3Qi52YWx1ZSksIDYpIH07CiAgICAgICAgcnVuQ29tcGFyZShyYW5nZUIsIHJhbmdlQSwgYFdlZWsgb2YgJHtGb3JtYXQuZGF0ZShyYW5nZUIuZnJvbSl9YCwgYFdlZWsgb2YgJHtGb3JtYXQuZGF0ZShyYW5nZUEuZnJvbSl9YCk7CiAgICAgIH0pKTsKICAgIH0gZWxzZSBpZiAobW9kZSA9PT0gJ2N1c3RvbScpIHsKICAgICAgY29uc3QgZkEgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOyBmQS50eXBlID0gJ2RhdGUnOyBmQS52YWx1ZSA9IGFkZERheXModG9kYXksIC0xMyk7CiAgICAgIGNvbnN0IHRBID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsgdEEudHlwZSA9ICdkYXRlJzsgdEEudmFsdWUgPSB0b2RheTsKICAgICAgY29uc3QgZkIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOyBmQi50eXBlID0gJ2RhdGUnOyBmQi52YWx1ZSA9IGFkZERheXModG9kYXksIC0yNyk7CiAgICAgIGNvbnN0IHRCID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsgdEIudHlwZSA9ICdkYXRlJzsgdEIudmFsdWUgPSBhZGREYXlzKHRvZGF5LCAtMTQpOwogICAgICByb3cuYXBwZW5kKAogICAgICAgIGxhYmVsZWQoJ1JhbmdlIEEgZnJvbScsIGZBKSwgbGFiZWxlZCgndG8nLCB0QSksCiAgICAgICAgbGFiZWxlZCgnUmFuZ2UgQiBmcm9tJywgZkIpLCBsYWJlbGVkKCd0bycsIHRCKSwKICAgICAgICBydW5CdG4oKCkgPT4gcnVuQ29tcGFyZSh7IGZyb206IGZBLnZhbHVlLCB0bzogdEEudmFsdWUgfSwgeyBmcm9tOiBmQi52YWx1ZSwgdG86IHRCLnZhbHVlIH0sICdSYW5nZSBBJywgJ1JhbmdlIEInKSkKICAgICAgKTsKICAgIH0gZWxzZSBpZiAobW9kZSA9PT0gJ21vbnRoJykgewogICAgICBjb25zdCB5ID0geWVhclNlbGVjdCh0aGlzWWVhcik7IGNvbnN0IG0gPSBtb250aFNlbGVjdChuZXcgRGF0ZSgpLmdldE1vbnRoKCkgKyAxKTsKICAgICAgY29uc3QgdG9nZ2xlID0gcGVyaW9kVG9nZ2xlKCk7CiAgICAgIHJvdy5hcHBlbmQobGFiZWxlZCgnWWVhcicsIHkpLCBsYWJlbGVkKCdNb250aCcsIG0pLCB0b2dnbGUuZWwsIHJ1bkJ0bihhc3luYyAoKSA9PiB7CiAgICAgICAgY29uc3QgcmVwb3J0ID0gYXdhaXQgQXBpLm1vbnRobHkoeyB5ZWFyOiB5LnZhbHVlLCBtb250aDogbS52YWx1ZSwgLi4uU3RhdGUuZ2V0RmlsdGVycygpIH0pOwogICAgICAgIHJlbmRlclBlcmlvZFJlcG9ydChyZXBvcnQsIHRvZ2dsZS5nZXQoKSk7CiAgICAgIH0pKTsKICAgIH0gZWxzZSBpZiAobW9kZSA9PT0gJ3F1YXJ0ZXInKSB7CiAgICAgIGNvbnN0IHkgPSB5ZWFyU2VsZWN0KHRoaXNZZWFyKTsgY29uc3QgcSA9IHF1YXJ0ZXJTZWxlY3QoKTsKICAgICAgY29uc3QgdG9nZ2xlID0gcGVyaW9kVG9nZ2xlKCk7CiAgICAgIHJvdy5hcHBlbmQobGFiZWxlZCgnWWVhcicsIHkpLCBsYWJlbGVkKCdRdWFydGVyJywgcSksIHRvZ2dsZS5lbCwgcnVuQnRuKGFzeW5jICgpID0+IHsKICAgICAgICBjb25zdCByZXBvcnQgPSBhd2FpdCBBcGkucXVhcnRlcmx5KHsgeWVhcjogeS52YWx1ZSwgcXVhcnRlcjogcS52YWx1ZSwgLi4uU3RhdGUuZ2V0RmlsdGVycygpIH0pOwogICAgICAgIHJlbmRlclBlcmlvZFJlcG9ydChyZXBvcnQsIHRvZ2dsZS5nZXQoKSk7CiAgICAgIH0pKTsKICAgIH0gZWxzZSBpZiAobW9kZSA9PT0gJ3l0ZCcpIHsKICAgICAgY29uc3QgeSA9IHllYXJTZWxlY3QodGhpc1llYXIpOwogICAgICByb3cuYXBwZW5kKGxhYmVsZWQoJ1llYXInLCB5KSwgcnVuQnRuKGFzeW5jICgpID0+IHsKICAgICAgICBjb25zdCByZXBvcnQgPSBhd2FpdCBBcGkueXRkKHsgeWVhcjogeS52YWx1ZSwgLi4uU3RhdGUuZ2V0RmlsdGVycygpIH0pOwogICAgICAgIHJlbmRlclBlcmlvZFJlcG9ydChyZXBvcnQsICd2c0xhc3RZZWFyJyk7CiAgICAgIH0pKTsKICAgIH0KCiAgICBjb250cm9scy5hcHBlbmRDaGlsZChyb3cpOwogIH0KCiAgZnVuY3Rpb24gbGFiZWxlZChsYWJlbCwgZWwpIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHdyYXAuY2xhc3NOYW1lID0gJ2ZpZWxkLWlubGluZSc7CiAgICB3cmFwLmFwcGVuZCh0ZXh0RWwoJ2xhYmVsJywgbGFiZWwpLCBlbCk7CiAgICByZXR1cm4gd3JhcDsKICB9CiAgZnVuY3Rpb24gcnVuQnRuKG9uQ2xpY2spIHsKICAgIGNvbnN0IGJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgYnRuLmNsYXNzTmFtZSA9ICdidG4gcHJpbWFyeSc7CiAgICBidG4udGV4dENvbnRlbnQgPSAnQ29tcGFyZSc7CiAgICBidG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBvbkNsaWNrKCkpOwogICAgcmV0dXJuIGJ0bjsKICB9CiAgZnVuY3Rpb24geWVhclNlbGVjdChkZWZhdWx0WWVhcikgewogICAgY29uc3Qgc2VsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc2VsZWN0Jyk7CiAgICBmb3IgKGxldCB5ID0gZGVmYXVsdFllYXIgLSAzOyB5IDw9IGRlZmF1bHRZZWFyICsgMTsgeSArPSAxKSB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOyBvcHQudmFsdWUgPSB5OyBvcHQudGV4dENvbnRlbnQgPSB5OwogICAgICBpZiAoeSA9PT0gZGVmYXVsdFllYXIpIG9wdC5zZWxlY3RlZCA9IHRydWU7CiAgICAgIHNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfQogICAgcmV0dXJuIHNlbDsKICB9CiAgZnVuY3Rpb24gbW9udGhTZWxlY3QoZGVmYXVsdE1vbnRoKSB7CiAgICBjb25zdCBuYW1lcyA9IFsnSmFudWFyeScsJ0ZlYnJ1YXJ5JywnTWFyY2gnLCdBcHJpbCcsJ01heScsJ0p1bmUnLCdKdWx5JywnQXVndXN0JywnU2VwdGVtYmVyJywnT2N0b2JlcicsJ05vdmVtYmVyJywnRGVjZW1iZXInXTsKICAgIGNvbnN0IHNlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NlbGVjdCcpOwogICAgbmFtZXMuZm9yRWFjaCgobiwgaSkgPT4gewogICAgICBjb25zdCBvcHQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsgb3B0LnZhbHVlID0gaSArIDE7IG9wdC50ZXh0Q29udGVudCA9IG47CiAgICAgIGlmIChpICsgMSA9PT0gZGVmYXVsdE1vbnRoKSBvcHQuc2VsZWN0ZWQgPSB0cnVlOwogICAgICBzZWwuYXBwZW5kQ2hpbGQob3B0KTsKICAgIH0pOwogICAgcmV0dXJuIHNlbDsKICB9CiAgZnVuY3Rpb24gcXVhcnRlclNlbGVjdCgpIHsKICAgIGNvbnN0IGN1cnJlbnRRID0gTWF0aC5mbG9vcihuZXcgRGF0ZSgpLmdldE1vbnRoKCkgLyAzKSArIDE7CiAgICBjb25zdCBzZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzZWxlY3QnKTsKICAgIFsxLCAyLCAzLCA0XS5mb3JFYWNoKChxKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOyBvcHQudmFsdWUgPSBxOyBvcHQudGV4dENvbnRlbnQgPSBgUSR7cX1gOwogICAgICBpZiAocSA9PT0gY3VycmVudFEpIG9wdC5zZWxlY3RlZCA9IHRydWU7CiAgICAgIHNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgICByZXR1cm4gc2VsOwogIH0KICBmdW5jdGlvbiBwZXJpb2RUb2dnbGUoKSB7CiAgICBjb25zdCBzZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzZWxlY3QnKTsKICAgIFtbJ3ZzUHJldmlvdXNQZXJpb2QnLCAndnMgUHJldmlvdXMgUGVyaW9kJ10sIFsndnNMYXN0WWVhcicsICd2cyBTYW1lIFBlcmlvZCBMYXN0IFllYXInXV0uZm9yRWFjaCgoW3YsIGxdKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOyBvcHQudmFsdWUgPSB2OyBvcHQudGV4dENvbnRlbnQgPSBsOwogICAgICBzZWwuYXBwZW5kQ2hpbGQob3B0KTsKICAgIH0pOwogICAgcmV0dXJuIHsgZWw6IGxhYmVsZWQoJ0NvbXBhcmUnLCBzZWwpLCBnZXQ6ICgpID0+IHNlbC52YWx1ZSB9OwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gcnVuQ29tcGFyZShyYW5nZUEsIHJhbmdlQiwgbGFiZWxBLCBsYWJlbEIpIHsKICAgIGNvbnN0IGZpbHRlcnMgPSBTdGF0ZS5nZXRGaWx0ZXJzKCk7CiAgICBjb25zdCByZXN1bHQgPSBhd2FpdCBBcGkuY29tcGFyZSh7CiAgICAgIGZyb21BOiByYW5nZUEuZnJvbSwgdG9BOiByYW5nZUEudG8sIGZyb21COiByYW5nZUIuZnJvbSwgdG9COiByYW5nZUIudG8sCiAgICAgIHBsYXRmb3JtOiBmaWx0ZXJzLnBsYXRmb3JtLCBjYW1wYWlnblR5cGU6IGZpbHRlcnMuY2FtcGFpZ25UeXBlLCBjb250ZW50VHlwZTogZmlsdGVycy5jb250ZW50VHlwZSwKICAgIH0pOwogICAgcmVuZGVyQ29tcGFyZVJlc3VsdChyZXN1bHQsIGxhYmVsQSwgbGFiZWxCKTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlclBlcmlvZFJlcG9ydChyZXBvcnQsIHdoaWNoKSB7CiAgICBjb25zdCBjbXAgPSByZXBvcnRbd2hpY2hdOwogICAgY29uc3QgbGFiZWxBID0gJ0N1cnJlbnQgcGVyaW9kJzsKICAgIGNvbnN0IGxhYmVsQiA9IHdoaWNoID09PSAndnNMYXN0WWVhcicgPyAnU2FtZSBwZXJpb2QgbGFzdCB5ZWFyJyA6ICdQcmV2aW91cyBwZXJpb2QnOwogICAgcmVuZGVyQ29tcGFyZVJlc3VsdChjbXAsIGxhYmVsQSwgbGFiZWxCLCByZXBvcnQucmFuZ2UpOwogIH0KCiAgZnVuY3Rpb24gc3RhdFRpbGUobGFiZWwsIGN1cnJlbnQsIHByZXZpb3VzLCBncm93dGgsIGlzRHVyYXRpb24pIHsKICAgIGNvbnN0IHRpbGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRpbGUuY2xhc3NOYW1lID0gJ3N0YXQtdGlsZSc7CiAgICBjb25zdCBjdXJEaXNwbGF5ID0gaXNEdXJhdGlvbiA/IEZvcm1hdC5kdXJhdGlvbihjdXJyZW50KSA6IEZvcm1hdC5jb21wYWN0KGN1cnJlbnQpOwogICAgY29uc3QgcHJldkRpc3BsYXkgPSBpc0R1cmF0aW9uID8gRm9ybWF0LmR1cmF0aW9uKHByZXZpb3VzKSA6IEZvcm1hdC5jb21wYWN0KHByZXZpb3VzKTsKICAgIHRpbGUuYXBwZW5kKAogICAgICB0ZXh0RWwoJ2RpdicsIGxhYmVsLCAnc3RhdC1sYWJlbCcpLAogICAgICB0ZXh0RWwoJ2RpdicsIGN1ckRpc3BsYXksICdzdGF0LXZhbHVlJyksCiAgICAgIHRleHRFbCgnZGl2JywgYCR7Rm9ybWF0LnBjdChncm93dGgpfSDCtyB3YXMgJHtwcmV2RGlzcGxheX1gLCBgc3RhdC1kZWx0YSAke0Zvcm1hdC5kZWx0YUNsYXNzKGdyb3d0aCl9YCkKICAgICk7CiAgICByZXR1cm4gdGlsZTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlckNvbXBhcmVSZXN1bHQocmVzdWx0LCBsYWJlbEEsIGxhYmVsQiwgaGVhZGxpbmUpIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnY29tcGFyaXNvblJlc3VsdHMnKTsKICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CgogICAgY29uc3QgdGl0bGUgPSB0ZXh0RWwoJ2RpdicsIGhlYWRsaW5lCiAgICAgID8gYCR7Rm9ybWF0LmRhdGUocmVzdWx0LnJhbmdlQS5mcm9tKX0g4oCTICR7Rm9ybWF0LmRhdGUocmVzdWx0LnJhbmdlQS50byl9YAogICAgICA6IGAke2xhYmVsQX06ICR7Rm9ybWF0LmRhdGUocmVzdWx0LnJhbmdlQS5mcm9tKX0g4oCTICR7Rm9ybWF0LmRhdGUocmVzdWx0LnJhbmdlQS50byl9ICB2cyAgJHtsYWJlbEJ9OiAke0Zvcm1hdC5kYXRlKHJlc3VsdC5yYW5nZUIuZnJvbSl9IOKAkyAke0Zvcm1hdC5kYXRlKHJlc3VsdC5yYW5nZUIudG8pfWAsCiAgICAgICdzZWN0aW9uLXRpdGxlJyk7CiAgICB3cmFwLmFwcGVuZENoaWxkKHRpdGxlKTsKCiAgICBjb25zdCBncmlkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBncmlkLmNsYXNzTmFtZSA9ICdzdGF0LWdyaWQnOwogICAgZ3JpZC5hcHBlbmRDaGlsZChzdGF0VGlsZSgnUG9zdHMnLCByZXN1bHQucmFuZ2VBLnRvdGFscy5wb3N0X2NvdW50LCByZXN1bHQucmFuZ2VCLnRvdGFscy5wb3N0X2NvdW50LCByZXN1bHQuZ3Jvd3RoLnBvc3RfY291bnQsIGZhbHNlKSk7CiAgICBNRVRSSUNfUk9XUy5mb3JFYWNoKChtKSA9PiB7CiAgICAgIGdyaWQuYXBwZW5kQ2hpbGQoc3RhdFRpbGUobS5sYWJlbCwgcmVzdWx0LnJhbmdlQS50b3RhbHNbbS5rZXldLCByZXN1bHQucmFuZ2VCLnRvdGFsc1ttLmtleV0sIHJlc3VsdC5ncm93dGhbbS5rZXldLCBtLmtleSA9PT0gJ3dhdGNoX3RpbWVfc2Vjb25kcycpKTsKICAgIH0pOwogICAgd3JhcC5hcHBlbmRDaGlsZChncmlkKTsKCiAgICByZW5kZXJQbGF0Zm9ybUNvbXBhcmlzb25DYXJkcyh3cmFwLCByZXN1bHQsIGxhYmVsQSwgbGFiZWxCKTsKICB9CgogIC8qKgogICAqICJBbGwgUGxhdGZvcm1zIiByZXBvcnQg4oCUIHRoZSBoZWFkbGluZSBDb21wYXJpc29ucyB2aWV3LiBVbmxpa2UgdGhlCiAgICogd2Vlay9jdXN0b20vbW9udGgvcXVhcnRlci95dGQgdG9vbHMgYWJvdmUsIHRoaXMgaWdub3JlcyB0aGUgc2hhcmVkCiAgICogcGxhdGZvcm0vY2FtcGFpZ24vY29udGVudC10eXBlIGZpbHRlciBiYXIgZW50aXJlbHkgYW5kIG5lZWRzIG5vIGRhdGUKICAgKiByYW5nZTogaXQgYWx3YXlzIGNvdmVycyBldmVyeSBwbGF0Zm9ybSB3aXRoIGFueSBkYXRhICh1cGxvYWRlZCBwb3N0cwogICAqIGFuZC9vciBtYW51YWxseS1lbnRlcmVkIEZvbGxvd2VycyBEYXRhIFJlY29yZCBoaXN0b3J5KS4KICAgKi8KICBhc3luYyBmdW5jdGlvbiBsb2FkUGxhdGZvcm1SZXBvcnQocGFyYW1zKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2NvbXBhcmlzb25SZXN1bHRzJyk7CiAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgd3JhcC5hcHBlbmRDaGlsZChza2VsZXRvblN0YXRHcmlkKDIpKTsKICAgIHdyYXAuYXBwZW5kQ2hpbGQoc2tlbGV0b25DaGFydCgpKTsKICAgIGNvbnN0IGhhc0V4cGxpY2l0UmFuZ2UgPSBwYXJhbXMgJiYgcGFyYW1zLmRhdGVGcm9tICYmIHBhcmFtcy5kYXRlVG87CiAgICBjb25zdCByZXBvcnQgPSBhd2FpdCBBcGkucGxhdGZvcm1SZXBvcnQoaGFzRXhwbGljaXRSYW5nZSA/IHBhcmFtcyA6IHt9KTsKICAgIHJlbmRlclBsYXRmb3JtUmVwb3J0KHJlcG9ydCk7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJQbGF0Zm9ybVJlcG9ydChyZXBvcnQpIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnY29tcGFyaXNvblJlc3VsdHMnKTsKICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CgogICAgaWYgKCFyZXBvcnQucGxhdGZvcm1zLmxlbmd0aCkgewogICAgICB3cmFwLmFwcGVuZENoaWxkKGVtcHR5U3RhdGUoewogICAgICAgIGljb246ICdnaXQtY29tcGFyZScsCiAgICAgICAgdGl0bGU6ICdObyBwbGF0Zm9ybSBkYXRhIHlldCcsCiAgICAgICAgbWVzc2FnZTogJ1VwbG9hZCBwb3N0cyBvciBhZGQgRm9sbG93ZXJzIERhdGEgUmVjb3JkIGVudHJpZXMgdG8gc2VlIGEgY3Jvc3MtcGxhdGZvcm0gY29tcGFyaXNvbiBoZXJlLicsCiAgICAgIH0pKTsKICAgICAgcmV0dXJuOwogICAgfQoKICAgIGNvbnN0IHJhbmdlTGFiZWwgPSByZXBvcnQucmFuZ2UuaXNFeHBsaWNpdAogICAgICA/IGAke0Zvcm1hdC5kYXRlKHJlcG9ydC5yYW5nZS5mcm9tKX0g4oCTICR7Rm9ybWF0LmRhdGUocmVwb3J0LnJhbmdlLnRvKX1gCiAgICAgIDogYEFsbCB0aW1lICgke0Zvcm1hdC5kYXRlKHJlcG9ydC5yYW5nZS5mcm9tKX0g4oCTICR7Rm9ybWF0LmRhdGUocmVwb3J0LnJhbmdlLnRvKX0pYDsKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBgUGxhdGZvcm0gQ29tcGFyaXNvbiBSZXBvcnQg4oCUICR7cmFuZ2VMYWJlbH1gLCAnc2VjdGlvbi10aXRsZScpKTsKCiAgICBjb25zdCBiZXN0UCA9IHJlcG9ydC5wbGF0Zm9ybXMuZmluZCgocCkgPT4gcC5wbGF0Zm9ybSA9PT0gcmVwb3J0LmJlc3RQbGF0Zm9ybSk7CiAgICBjb25zdCB3b3JzdFAgPSByZXBvcnQucGxhdGZvcm1zLmZpbmQoKHApID0+IHAucGxhdGZvcm0gPT09IHJlcG9ydC53b3JzdFBsYXRmb3JtKTsKICAgIGlmIChiZXN0UCB8fCB3b3JzdFApIHsKICAgICAgY29uc3QgaGlnaGxpZ2h0R3JpZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBoaWdobGlnaHRHcmlkLmNsYXNzTmFtZSA9ICdzdGF0LWdyaWQnOwogICAgICBpZiAoYmVzdFApIHsKICAgICAgICBjb25zdCB0aWxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgICAgdGlsZS5jbGFzc05hbWUgPSAnc3RhdC10aWxlJzsKICAgICAgICB0aWxlLmFwcGVuZCgKICAgICAgICAgIHRleHRFbCgnZGl2JywgJ0Jlc3QtUGVyZm9ybWluZyBQbGF0Zm9ybScsICdzdGF0LWxhYmVsJyksCiAgICAgICAgICB0ZXh0RWwoJ2RpdicsIGJlc3RQLmxhYmVsLCAnc3RhdC12YWx1ZScpLAogICAgICAgICAgdGV4dEVsKCdkaXYnLCBgUmVhY2ggJHtGb3JtYXQuc21hcnQoYmVzdFAudG90YWxzLnJlYWNoKX0gwrcgRW5nYWdlbWVudCAke0Zvcm1hdC5zbWFydChiZXN0UC50b3RhbHMuZW5nYWdlbWVudCl9YCwgJ3Bvc3QtbWV0YScpCiAgICAgICAgKTsKICAgICAgICBoaWdobGlnaHRHcmlkLmFwcGVuZENoaWxkKHRpbGUpOwogICAgICB9CiAgICAgIGlmICh3b3JzdFApIHsKICAgICAgICBjb25zdCB0aWxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgICAgdGlsZS5jbGFzc05hbWUgPSAnc3RhdC10aWxlJzsKICAgICAgICB0aWxlLmFwcGVuZCgKICAgICAgICAgIHRleHRFbCgnZGl2JywgJ0xvd2VzdC1QZXJmb3JtaW5nIFBsYXRmb3JtJywgJ3N0YXQtbGFiZWwnKSwKICAgICAgICAgIHRleHRFbCgnZGl2Jywgd29yc3RQLmxhYmVsLCAnc3RhdC12YWx1ZScpLAogICAgICAgICAgdGV4dEVsKCdkaXYnLCBgUmVhY2ggJHtGb3JtYXQuc21hcnQod29yc3RQLnRvdGFscy5yZWFjaCl9IMK3IEVuZ2FnZW1lbnQgJHtGb3JtYXQuc21hcnQod29yc3RQLnRvdGFscy5lbmdhZ2VtZW50KX1gLCAncG9zdC1tZXRhJykKICAgICAgICApOwogICAgICAgIGhpZ2hsaWdodEdyaWQuYXBwZW5kQ2hpbGQodGlsZSk7CiAgICAgIH0KICAgICAgd3JhcC5hcHBlbmRDaGlsZChoaWdobGlnaHRHcmlkKTsKICAgIH0KCiAgICBjb25zdCB0YWJsZUNhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRhYmxlQ2FyZC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICB0YWJsZUNhcmQuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMycsICdQbGF0Zm9ybSBSYW5raW5nJykpOwogICAgY29uc3QgdGFibGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0YWJsZScpOwogICAgdGFibGUuY2xhc3NOYW1lID0gJ2RhdGEtdGFibGUnOwogICAgY29uc3QgdGhlYWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0aGVhZCcpOwogICAgY29uc3QgaGVhZFJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICBbJ1JhbmsnLCAnUGxhdGZvcm0nLCAnUG9zdHMnLCAnUmVhY2gnLCAnRW5nYWdlbWVudCcsICdJbXByZXNzaW9ucycsICdGb2xsb3dlciBHcm93dGgnXS5mb3JFYWNoKChsYWJlbCwgaSkgPT4gewogICAgICBjb25zdCB0aCA9IHRleHRFbCgndGgnLCBsYWJlbCk7CiAgICAgIGlmIChpID49IDIpIHRoLmNsYXNzTGlzdC5hZGQoJ251bScpOwogICAgICBoZWFkUm93LmFwcGVuZENoaWxkKHRoKTsKICAgIH0pOwogICAgdGhlYWQuYXBwZW5kQ2hpbGQoaGVhZFJvdyk7CiAgICB0YWJsZS5hcHBlbmRDaGlsZCh0aGVhZCk7CiAgICBjb25zdCB0Ym9keSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3Rib2R5Jyk7CiAgICByZXBvcnQucGxhdGZvcm1zLmZvckVhY2goKHApID0+IHsKICAgICAgY29uc3QgdHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgICB0ci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3RkJywgcC5vdmVyYWxsUmFuayA/IGAjJHtwLm92ZXJhbGxSYW5rfWAgOiAn4oCUJykpOwogICAgICBjb25zdCBwbGF0VGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBjb25zdCBwaWxsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOyBwaWxsLmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1waWxsJzsKICAgICAgY29uc3QgZG90ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOyBkb3QuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLWRvdCc7IGRvdC5zdHlsZS5iYWNrZ3JvdW5kID0gcC5jb2xvcjsKICAgICAgcGlsbC5hcHBlbmQoZG90LCBkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShwLmxhYmVsKSk7CiAgICAgIHBsYXRUZC5hcHBlbmRDaGlsZChwaWxsKTsKICAgICAgdHIuYXBwZW5kQ2hpbGQocGxhdFRkKTsKICAgICAgdHIuYXBwZW5kQ2hpbGQodGV4dEVsKCd0ZCcsIEZvcm1hdC5udW1iZXIocC5wb3N0Q291bnQpLCAnbnVtJykpOwogICAgICB0ci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3RkJywgRm9ybWF0LnNtYXJ0KHAudG90YWxzLnJlYWNoKSwgJ251bScpKTsKICAgICAgdHIuYXBwZW5kQ2hpbGQodGV4dEVsKCd0ZCcsIEZvcm1hdC5zbWFydChwLnRvdGFscy5lbmdhZ2VtZW50KSwgJ251bScpKTsKICAgICAgdHIuYXBwZW5kQ2hpbGQodGV4dEVsKCd0ZCcsIEZvcm1hdC5zbWFydChwLnRvdGFscy5pbXByZXNzaW9ucyksICdudW0nKSk7CiAgICAgIGNvbnN0IGZvbGxvd2VyVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBmb2xsb3dlclRkLmNsYXNzTmFtZSA9ICdudW0nOwogICAgICBpZiAocC5mb2xsb3dlcnMuY2hhbmdlID09PSBudWxsKSB7CiAgICAgICAgZm9sbG93ZXJUZC5hcHBlbmRDaGlsZChkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShwLmZvbGxvd2Vycy5sYXRlc3QgIT09IG51bGwgPyBGb3JtYXQubnVtYmVyKHAuZm9sbG93ZXJzLmxhdGVzdCkgOiAn4oCUJykpOwogICAgICB9IGVsc2UgewogICAgICAgIGNvbnN0IGZvbGxvd2VyV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgICAgICBmb2xsb3dlcldyYXAuc3R5bGUuZGlzcGxheSA9ICdpbmxpbmUtZmxleCc7CiAgICAgICAgZm9sbG93ZXJXcmFwLnN0eWxlLmFsaWduSXRlbXMgPSAnY2VudGVyJzsKICAgICAgICBmb2xsb3dlcldyYXAuc3R5bGUuZ2FwID0gJzZweCc7CiAgICAgICAgZm9sbG93ZXJXcmFwLmFwcGVuZENoaWxkKHRleHRFbCgnc3BhbicsIGAke3AuZm9sbG93ZXJzLmNoYW5nZSA+IDAgPyAnKycgOiAnJ30ke0Zvcm1hdC5udW1iZXIocC5mb2xsb3dlcnMuY2hhbmdlKX1gLCBgc3RhdC1kZWx0YSAke0Zvcm1hdC5kZWx0YUNsYXNzKHAuZm9sbG93ZXJzLmNoYW5nZSl9YCkpOwogICAgICAgIGlmIChwLmZvbGxvd2Vycy5jaGFuZ2VQY3QgIT09IG51bGwpIGZvbGxvd2VyV3JhcC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3NwYW4nLCBgKCR7Rm9ybWF0LnBjdChwLmZvbGxvd2Vycy5jaGFuZ2VQY3QpfSlgLCAncG9zdC1tZXRhJykpOwogICAgICAgIGZvbGxvd2VyVGQuYXBwZW5kQ2hpbGQoZm9sbG93ZXJXcmFwKTsKICAgICAgfQogICAgICB0ci5hcHBlbmRDaGlsZChmb2xsb3dlclRkKTsKICAgICAgdGJvZHkuYXBwZW5kQ2hpbGQodHIpOwogICAgfSk7CiAgICB0YWJsZS5hcHBlbmRDaGlsZCh0Ym9keSk7CiAgICBjb25zdCB0YWJsZVNjcm9sbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGFibGVTY3JvbGwuY2xhc3NOYW1lID0gJ3RhYmxlLXNjcm9sbCc7CiAgICB0YWJsZVNjcm9sbC5hcHBlbmRDaGlsZCh0YWJsZSk7CiAgICB0YWJsZUNhcmQuYXBwZW5kQ2hpbGQodGFibGVTY3JvbGwpOwogICAgd3JhcC5hcHBlbmRDaGlsZCh0YWJsZUNhcmQpOwoKICAgIGNvbnN0IGNoYXJ0Q2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY2hhcnRDYXJkLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGNvbnN0IGNoYXJ0SGVhZGVyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjaGFydEhlYWRlci5jbGFzc05hbWUgPSAnY2FyZC1oZWFkZXInOwogICAgY2hhcnRIZWFkZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMycsICdNZXRyaWMgQ29tcGFyaXNvbicpKTsKICAgIGNvbnN0IG1ldHJpY1NlbGVjdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NlbGVjdCcpOwogICAgTUVUUklDX1JPV1MuZm9yRWFjaCgobSkgPT4gewogICAgICBjb25zdCBvcHQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsgb3B0LnZhbHVlID0gbS5rZXk7IG9wdC50ZXh0Q29udGVudCA9IG0ubGFiZWw7CiAgICAgIGlmIChtLmtleSA9PT0gcGxhdGZvcm1DaGFydE1ldHJpYykgb3B0LnNlbGVjdGVkID0gdHJ1ZTsKICAgICAgbWV0cmljU2VsZWN0LmFwcGVuZENoaWxkKG9wdCk7CiAgICB9KTsKICAgIG1ldHJpY1NlbGVjdC5hZGRFdmVudExpc3RlbmVyKCdjaGFuZ2UnLCAoKSA9PiB7CiAgICAgIHBsYXRmb3JtQ2hhcnRNZXRyaWMgPSBtZXRyaWNTZWxlY3QudmFsdWU7CiAgICAgIGRyYXdQbGF0Zm9ybVJlcG9ydENoYXJ0KHJlcG9ydCk7CiAgICB9KTsKICAgIGNoYXJ0SGVhZGVyLmFwcGVuZENoaWxkKG1ldHJpY1NlbGVjdCk7CiAgICBjb25zdCBjaGFydFdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNoYXJ0V3JhcC5jbGFzc05hbWUgPSAnY2hhcnQtd3JhcCB0YWxsJzsKICAgIGNoYXJ0V3JhcC5pbm5lckhUTUwgPSAnPGNhbnZhcyBpZD0icGxhdGZvcm1SZXBvcnRDYW52YXMiPjwvY2FudmFzPic7CiAgICBjaGFydENhcmQuYXBwZW5kKGNoYXJ0SGVhZGVyLCBjaGFydFdyYXApOwogICAgd3JhcC5hcHBlbmRDaGlsZChjaGFydENhcmQpOwogICAgZHJhd1BsYXRmb3JtUmVwb3J0Q2hhcnQocmVwb3J0KTsKCiAgICBjb25zdCB3aXRoRm9sbG93ZXJzID0gcmVwb3J0LnBsYXRmb3Jtcy5maWx0ZXIoKHApID0+IHAuZm9sbG93ZXJzLmxhdGVzdCAhPT0gbnVsbCk7CiAgICBpZiAod2l0aEZvbGxvd2Vycy5sZW5ndGgpIHsKICAgICAgY29uc3QgZm9sbG93ZXJDYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIGZvbGxvd2VyQ2FyZC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICAgIGZvbGxvd2VyQ2FyZC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2gzJywgJ0ZvbGxvd2VyIEdyb3d0aCBieSBQbGF0Zm9ybScpKTsKICAgICAgY29uc3QgZkNoYXJ0V3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBmQ2hhcnRXcmFwLmNsYXNzTmFtZSA9ICdjaGFydC13cmFwIHRhbGwnOwogICAgICBmQ2hhcnRXcmFwLmlubmVySFRNTCA9ICc8Y2FudmFzIGlkPSJwbGF0Zm9ybUZvbGxvd2VyQ2FudmFzIj48L2NhbnZhcz4nOwogICAgICBmb2xsb3dlckNhcmQuYXBwZW5kQ2hpbGQoZkNoYXJ0V3JhcCk7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQoZm9sbG93ZXJDYXJkKTsKICAgICAgQ2hhcnRzLnBsYXRmb3JtQmFyQ2hhcnQoJ3BsYXRmb3JtRm9sbG93ZXJDYW52YXMnLCB7CiAgICAgICAgbGFiZWxzOiB3aXRoRm9sbG93ZXJzLm1hcCgocCkgPT4gcC5sYWJlbCksCiAgICAgICAgZGF0YTogd2l0aEZvbGxvd2Vycy5tYXAoKHApID0+IHAuZm9sbG93ZXJzLmxhdGVzdCB8fCAwKSwKICAgICAgICBjb2xvcnM6IHdpdGhGb2xsb3dlcnMubWFwKChwKSA9PiBwLmNvbG9yKSwKICAgICAgICBmb3JtYXRWYWx1ZTogKHYpID0+IEZvcm1hdC5zbWFydCh2KSwKICAgICAgfSk7CiAgICB9CgogICAgY29uc3QgaW5zaWdodHNDYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBpbnNpZ2h0c0NhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgaW5zaWdodHNDYXJkLmFwcGVuZENoaWxkKHRleHRFbCgnaDMnLCAnSW5zaWdodHMgJiBTdW1tYXJ5JykpOwogICAgY29uc3QgbGlzdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3VsJyk7CiAgICBsaXN0LmNsYXNzTmFtZSA9ICdpbnNpZ2h0cy1saXN0JzsKICAgIHJlcG9ydC5pbnNpZ2h0cy5mb3JFYWNoKChsaW5lKSA9PiBsaXN0LmFwcGVuZENoaWxkKHRleHRFbCgnbGknLCBsaW5lKSkpOwogICAgaW5zaWdodHNDYXJkLmFwcGVuZENoaWxkKGxpc3QpOwogICAgd3JhcC5hcHBlbmRDaGlsZChpbnNpZ2h0c0NhcmQpOwogIH0KCiAgZnVuY3Rpb24gZHJhd1BsYXRmb3JtUmVwb3J0Q2hhcnQocmVwb3J0KSB7CiAgICBDaGFydHMucGxhdGZvcm1CYXJDaGFydCgncGxhdGZvcm1SZXBvcnRDYW52YXMnLCB7CiAgICAgIGxhYmVsczogcmVwb3J0LnBsYXRmb3Jtcy5tYXAoKHApID0+IHAubGFiZWwpLAogICAgICBkYXRhOiByZXBvcnQucGxhdGZvcm1zLm1hcCgocCkgPT4gcC50b3RhbHNbcGxhdGZvcm1DaGFydE1ldHJpY10gfHwgMCksCiAgICAgIGNvbG9yczogcmVwb3J0LnBsYXRmb3Jtcy5tYXAoKHApID0+IHAuY29sb3IpLAogICAgICBmb3JtYXRWYWx1ZTogKHYpID0+IChwbGF0Zm9ybUNoYXJ0TWV0cmljID09PSAnd2F0Y2hfdGltZV9zZWNvbmRzJyA/IEZvcm1hdC5kdXJhdGlvbih2KSA6IEZvcm1hdC5zbWFydCh2KSksCiAgICB9KTsKICB9CgogIC8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgICBQbGF0Zm9ybSBQZXJmb3JtYW5jZSBDb21wYXJpc29uIOKAlCByZXBsYWNlcyB0aGUgb2xkIGdyb3VwZWQKICAgICAiUmFuZ2UgQSB2cyBSYW5nZSBCIGJ5IFBsYXRmb3JtIiBjaGFydC4gT25lIGNhcmQgcGVyIHBsYXRmb3JtCiAgICAgd2l0aCBhbnkgZGF0YSBpbiBlaXRoZXIgcmFuZ2UsIGJ1aWx0IGVudGlyZWx5IGZyb20gdGhlIHNhbWUKICAgICBjb21wYXJlUmFuZ2VzKCkgcmVzcG9uc2UgdGhlIHN0YXQtdGlsZSBncmlkIGFib3ZlIGFscmVhZHkKICAgICB1c2VzIChyZXN1bHQucmFuZ2VBLnBsYXRmb3JtcyAvIHJlc3VsdC5yYW5nZUIucGxhdGZvcm1zKSDigJQgbm8KICAgICBleHRyYSBmZXRjaC4KICAgICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCiAgY29uc3QgQUxMX0NBUkRfTUVUUklDUyA9IFt7IGtleTogJ3Bvc3RfY291bnQnLCBsYWJlbDogJ1Bvc3RzJyB9LCAuLi5NRVRSSUNfUk9XU107CiAgY29uc3QgQ0FSRF9TT1JUX01PREVTID0gWwogICAgeyBrZXk6ICdvdmVyYWxsJywgbGFiZWw6ICdPdmVyYWxsIFBlcmZvcm1hbmNlJyB9LAogICAgeyBrZXk6ICdncm93dGgnLCBsYWJlbDogJ0hpZ2hlc3QgR3Jvd3RoJyB9LAogICAgeyBrZXk6ICdlbmdhZ2VtZW50JywgbGFiZWw6ICdIaWdoZXN0IEVuZ2FnZW1lbnQnIH0sCiAgICB7IGtleTogJ2ZvbGxvd2VycycsIGxhYmVsOiAnTW9zdCBGb2xsb3dlcnMnIH0sCiAgICB7IGtleTogJ3Bvc3RzJywgbGFiZWw6ICdNb3N0IFBvc3RzJyB9LAogICAgeyBrZXk6ICdhbHBoYScsIGxhYmVsOiAnQWxwaGFiZXRpY2FsJyB9LAogIF07CgogIC8qKiBQZXItbWV0cmljIHthLCBiLCBkaWZmLCBwY3REaWZmfSBhY3Jvc3MgYm90aCByYW5nZXMgZm9yIG9uZSBwbGF0Zm9ybSwgc2tpcHBpbmcgYW55IG1ldHJpYyB0aGF0J3MgemVybyBpbiBib3RoIOKAlCBhIHBsYXRmb3JtJ3MgY2FyZCBzaG91bGQgb25seSBldmVyIHNob3cgbWV0cmljcyBpdCBhY3R1YWxseSBoYXMuICovCiAgZnVuY3Rpb24gY29tcHV0ZUNhcmRNZXRyaWNzKHBsYXRmb3JtQSwgcGxhdGZvcm1CKSB7CiAgICBjb25zdCBtZXRyaWNzID0gW107CiAgICBBTExfQ0FSRF9NRVRSSUNTLmZvckVhY2goKHsga2V5LCBsYWJlbCB9KSA9PiB7CiAgICAgIGNvbnN0IGEgPSAocGxhdGZvcm1BICYmIHBsYXRmb3JtQVtrZXldKSB8fCAwOwogICAgICBjb25zdCBiID0gKHBsYXRmb3JtQiAmJiBwbGF0Zm9ybUJba2V5XSkgfHwgMDsKICAgICAgaWYgKGEgPT09IDAgJiYgYiA9PT0gMCkgcmV0dXJuOwogICAgICBjb25zdCBkaWZmID0gYSAtIGI7CiAgICAgIGNvbnN0IHBjdERpZmYgPSBiID8gTWF0aC5yb3VuZCgoZGlmZiAvIGIpICogMTAwMCkgLyAxMCA6IChhID4gMCA/IG51bGwgOiAwKTsKICAgICAgbWV0cmljcy5wdXNoKHsga2V5LCBsYWJlbCwgYSwgYiwgZGlmZiwgcGN0RGlmZiwgaXNEdXJhdGlvbjoga2V5ID09PSAnd2F0Y2hfdGltZV9zZWNvbmRzJyB9KTsKICAgIH0pOwogICAgcmV0dXJuIG1ldHJpY3M7CiAgfQoKICAvKiogQSBzaW5nbGUgImhvdyBkaWQgdGhpcyBwbGF0Zm9ybSBkbyBvdmVyYWxsIiBudW1iZXI6IHRoZSBhdmVyYWdlICUgY2hhbmdlIGFjcm9zcyBldmVyeSBtZXRyaWMgdGhhdCBoYXMgYSBjb21wdXRhYmxlIHBlcmNlbnRhZ2UgKGEgbWV0cmljIGdvaW5nIGZyb20gMCB0byBzb21ldGhpbmcgaGFzIG5vIHBlcmNlbnRhZ2Ug4oCUICJuZXciLCBub3QgY291bnRlZCBlaXRoZXIgd2F5KS4gKi8KICBmdW5jdGlvbiBvdmVyYWxsUGN0Q2hhbmdlKG1ldHJpY3MpIHsKICAgIGNvbnN0IHdpdGhQY3QgPSBtZXRyaWNzLmZpbHRlcigobSkgPT4gbS5wY3REaWZmICE9PSBudWxsKTsKICAgIGlmICghd2l0aFBjdC5sZW5ndGgpIHJldHVybiBudWxsOwogICAgcmV0dXJuIE1hdGgucm91bmQoKHdpdGhQY3QucmVkdWNlKChzdW0sIG0pID0+IHN1bSArIG0ucGN0RGlmZiwgMCkgLyB3aXRoUGN0Lmxlbmd0aCkgKiAxMCkgLyAxMDsKICB9CgogIGZ1bmN0aW9uIGJlc3RXZWFrZXN0TWV0cmljKG1ldHJpY3MpIHsKICAgIGNvbnN0IHdpdGhQY3QgPSBtZXRyaWNzLmZpbHRlcigobSkgPT4gbS5wY3REaWZmICE9PSBudWxsKTsKICAgIGlmICghd2l0aFBjdC5sZW5ndGgpIHJldHVybiB7IGJlc3Q6IG51bGwsIHdlYWtlc3Q6IG51bGwgfTsKICAgIGNvbnN0IGJlc3QgPSB3aXRoUGN0LnJlZHVjZSgoYSwgYikgPT4gKGIucGN0RGlmZiA+IGEucGN0RGlmZiA/IGIgOiBhKSk7CiAgICBjb25zdCB3ZWFrZXN0ID0gd2l0aFBjdC5yZWR1Y2UoKGEsIGIpID0+IChiLnBjdERpZmYgPCBhLnBjdERpZmYgPyBiIDogYSkpOwogICAgcmV0dXJuIHsgYmVzdCwgd2Vha2VzdCB9OwogIH0KCiAgZnVuY3Rpb24gdHJlbmREaXJlY3Rpb24ocGN0KSB7CiAgICBpZiAocGN0ID09PSBudWxsIHx8IHBjdCA9PT0gdW5kZWZpbmVkKSByZXR1cm4gJ2ZsYXQnOwogICAgaWYgKHBjdCA+IDAuNSkgcmV0dXJuICd1cCc7CiAgICBpZiAocGN0IDwgLTAuNSkgcmV0dXJuICdkb3duJzsKICAgIHJldHVybiAnZmxhdCc7CiAgfQoKICBmdW5jdGlvbiBidWlsZFBsYXRmb3JtQ2FyZHMocmVzdWx0KSB7CiAgICBjb25zdCBwbGF0Zm9ybU9wdGlvbnMgPSAod2luZG93Ll9fZmlsdGVyT3B0aW9uc0NhY2hlIHx8IHsgYWxsUGxhdGZvcm1zOiBbXSB9KS5hbGxQbGF0Zm9ybXM7CiAgICBjb25zdCBpZHMgPSBbLi4ubmV3IFNldChbLi4ucmVzdWx0LnJhbmdlQS5wbGF0Zm9ybXMsIC4uLnJlc3VsdC5yYW5nZUIucGxhdGZvcm1zXS5tYXAoKHApID0+IHAucGxhdGZvcm0pKV07CiAgICBjb25zdCBieUlkQSA9IE9iamVjdC5mcm9tRW50cmllcyhyZXN1bHQucmFuZ2VBLnBsYXRmb3Jtcy5tYXAoKHApID0+IFtwLnBsYXRmb3JtLCBwXSkpOwogICAgY29uc3QgYnlJZEIgPSBPYmplY3QuZnJvbUVudHJpZXMocmVzdWx0LnJhbmdlQi5wbGF0Zm9ybXMubWFwKChwKSA9PiBbcC5wbGF0Zm9ybSwgcF0pKTsKCiAgICByZXR1cm4gaWRzCiAgICAgIC5tYXAoKGlkKSA9PiB7CiAgICAgICAgY29uc3QgbWV0YSA9IHBsYXRmb3JtT3B0aW9ucy5maW5kKChwKSA9PiBwLmlkID09PSBpZCkgfHwgeyBpZCwgbGFiZWw6IGlkLCBjb2xvcjogJ3ZhcigtLXNlcmllcy0xKScgfTsKICAgICAgICBjb25zdCBhID0gYnlJZEFbaWRdIHx8IG51bGw7CiAgICAgICAgY29uc3QgYiA9IGJ5SWRCW2lkXSB8fCBudWxsOwogICAgICAgIGNvbnN0IG1ldHJpY3MgPSBjb21wdXRlQ2FyZE1ldHJpY3MoYSwgYik7CiAgICAgICAgY29uc3QgeyBiZXN0LCB3ZWFrZXN0IH0gPSBiZXN0V2Vha2VzdE1ldHJpYyhtZXRyaWNzKTsKICAgICAgICByZXR1cm4gewogICAgICAgICAgcGxhdGZvcm06IGlkLAogICAgICAgICAgbGFiZWw6IG1ldGEubGFiZWwsCiAgICAgICAgICBjb2xvcjogbWV0YS5jb2xvciwKICAgICAgICAgIG1ldHJpY3MsCiAgICAgICAgICBvdmVyYWxsOiBvdmVyYWxsUGN0Q2hhbmdlKG1ldHJpY3MpLAogICAgICAgICAgYmVzdCwKICAgICAgICAgIHdlYWtlc3QsCiAgICAgICAgICBmb2xsb3dlcnNHYWluZWQ6IChhID8gYS5mb2xsb3dlcnNfZ2FpbmVkIHx8IDAgOiAwKSArIChiID8gYi5mb2xsb3dlcnNfZ2FpbmVkIHx8IDAgOiAwKSwKICAgICAgICAgIHBvc3RzOiAoYSA/IGEucG9zdF9jb3VudCB8fCAwIDogMCkgKyAoYiA/IGIucG9zdF9jb3VudCB8fCAwIDogMCksCiAgICAgICAgICBlbmdhZ2VtZW50VG90YWw6IChhID8gYS5lbmdhZ2VtZW50IHx8IDAgOiAwKSArIChiID8gYi5lbmdhZ2VtZW50IHx8IDAgOiAwKSwKICAgICAgICB9OwogICAgICB9KQogICAgICAuZmlsdGVyKChjYXJkKSA9PiBjYXJkLm1ldHJpY3MubGVuZ3RoID4gMCk7CiAgfQoKICBmdW5jdGlvbiBzb3J0Q2FyZHMoY2FyZHMsIHNvcnRNb2RlKSB7CiAgICBjb25zdCBhcnIgPSBbLi4uY2FyZHNdOwogICAgaWYgKHNvcnRNb2RlID09PSAnZW5nYWdlbWVudCcpIHJldHVybiBhcnIuc29ydCgoeCwgeSkgPT4geS5lbmdhZ2VtZW50VG90YWwgLSB4LmVuZ2FnZW1lbnRUb3RhbCk7CiAgICBpZiAoc29ydE1vZGUgPT09ICdmb2xsb3dlcnMnKSByZXR1cm4gYXJyLnNvcnQoKHgsIHkpID0+IHkuZm9sbG93ZXJzR2FpbmVkIC0geC5mb2xsb3dlcnNHYWluZWQpOwogICAgaWYgKHNvcnRNb2RlID09PSAncG9zdHMnKSByZXR1cm4gYXJyLnNvcnQoKHgsIHkpID0+IHkucG9zdHMgLSB4LnBvc3RzKTsKICAgIGlmIChzb3J0TW9kZSA9PT0gJ2FscGhhJykgcmV0dXJuIGFyci5zb3J0KCh4LCB5KSA9PiB4LmxhYmVsLmxvY2FsZUNvbXBhcmUoeS5sYWJlbCkpOwogICAgLy8gJ292ZXJhbGwnIGFuZCAnZ3Jvd3RoJyBib3RoIHJhbmsgYnkgdGhlIHNhbWUgY29tcG9zaXRlICUgY2hhbmdlIOKAlCB0aGUgdHdvIGxhYmVscwogICAgLy8gcmVhZCBkaWZmZXJlbnRseSBvbiB0aGUgc2FtZSB1bmRlcmx5aW5nIG51bWJlciwgcGVyIHRoZSByZXF1ZXN0ZWQgb3B0aW9uIGxpc3QuCiAgICByZXR1cm4gYXJyLnNvcnQoKHgsIHkpID0+ICh5Lm92ZXJhbGwgPz8gLUluZmluaXR5KSAtICh4Lm92ZXJhbGwgPz8gLUluZmluaXR5KSk7CiAgfQoKICBmdW5jdGlvbiBidWlsZE1ldHJpY1JvdyhtKSB7CiAgICBjb25zdCByb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHJvdy5jbGFzc05hbWUgPSAncGNjLW1ldHJpYy1yb3cnOwogICAgY29uc3QgaGVhZGVyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBoZWFkZXIuY2xhc3NOYW1lID0gJ3BjYy1tZXRyaWMtaGVhZGVyJzsKICAgIGNvbnN0IGZtdCA9ICh2KSA9PiAobS5pc0R1cmF0aW9uID8gRm9ybWF0LmR1cmF0aW9uKHYpIDogRm9ybWF0LnNtYXJ0KHYpKTsKICAgIGNvbnN0IGRpZmZUZXh0ID0gbS5wY3REaWZmID09PSBudWxsCiAgICAgID8gYCR7bS5kaWZmID4gMCA/ICcrJyA6ICcnfSR7Zm10KG0uZGlmZil9IChuZXcpYAogICAgICA6IGAke20uZGlmZiA+IDAgPyAnKycgOiAnJ30ke2ZtdChtLmRpZmYpfSAoJHtGb3JtYXQucGN0KG0ucGN0RGlmZil9KWA7CiAgICBoZWFkZXIuYXBwZW5kKAogICAgICB0ZXh0RWwoJ3NwYW4nLCBtLmxhYmVsLCAncGNjLW1ldHJpYy1sYWJlbCcpLAogICAgICB0ZXh0RWwoJ3NwYW4nLCBkaWZmVGV4dCwgYHBjYy1tZXRyaWMtZGlmZiAke0Zvcm1hdC5kZWx0YUNsYXNzKG0ucGN0RGlmZil9YCkKICAgICk7CiAgICByb3cuYXBwZW5kQ2hpbGQoaGVhZGVyKTsKICAgIGNvbnN0IG1heCA9IE1hdGgubWF4KG0uYSwgbS5iLCAxKTsKICAgIHJvdy5hcHBlbmRDaGlsZChidWlsZEJhcih7IGxhYmVsOiAnQScsIHZhbHVlOiBtLmEsIG1heCwgY29sb3JWYXI6ICctLXNlcmllcy0xJywgZm9ybWF0VmFsdWU6IGZtdCB9KSk7CiAgICByb3cuYXBwZW5kQ2hpbGQoYnVpbGRCYXIoeyBsYWJlbDogJ0InLCB2YWx1ZTogbS5iLCBtYXgsIGNvbG9yVmFyOiAnLS10ZXh0LW11dGVkJywgZm9ybWF0VmFsdWU6IGZtdCB9KSk7CiAgICByZXR1cm4gcm93OwogIH0KCiAgZnVuY3Rpb24gYnVpbGRDYXJkRm9vdGVyKGNhcmQpIHsKICAgIGNvbnN0IGZvb3RlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZm9vdGVyLmNsYXNzTmFtZSA9ICdwY2MtZm9vdGVyJzsKICAgIGNvbnN0IGRpciA9IHRyZW5kRGlyZWN0aW9uKGNhcmQub3ZlcmFsbCk7CiAgICBjb25zdCByZXN1bHRUZXh0ID0gY2FyZC5vdmVyYWxsID09PSBudWxsCiAgICAgID8gJ05vdCBlbm91Z2ggZGF0YSB0byBjb21wYXJlJwogICAgICA6IGAke2RpciA9PT0gJ3VwJyA/ICdJbXByb3ZlZCcgOiBkaXIgPT09ICdkb3duJyA/ICdEZWNsaW5lZCcgOiAnTm8gc2lnbmlmaWNhbnQgY2hhbmdlJ30ke2RpciAhPT0gJ2ZsYXQnID8gYCBieSAke01hdGguYWJzKGNhcmQub3ZlcmFsbCl9JWAgOiAnJ31gOwogICAgZm9vdGVyLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgJ092ZXJhbGwgUmVzdWx0JywgJ3BjYy1mb290ZXItbGFiZWwnKSk7CiAgICBmb290ZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCByZXN1bHRUZXh0LCBgcGNjLWZvb3Rlci12YWx1ZSAke2Rpcn1gKSk7CiAgICBpZiAoY2FyZC5iZXN0KSBmb290ZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBgQmVzdCBNZXRyaWM6ICR7Y2FyZC5iZXN0LmxhYmVsfSAoJHtGb3JtYXQucGN0KGNhcmQuYmVzdC5wY3REaWZmKX0pYCwgJ3BjYy1mb290ZXItZGV0YWlsJykpOwogICAgaWYgKGNhcmQud2Vha2VzdCAmJiBjYXJkLndlYWtlc3QgIT09IGNhcmQuYmVzdCkgewogICAgICBmb290ZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBgV2Vha2VzdCBNZXRyaWM6ICR7Y2FyZC53ZWFrZXN0LmxhYmVsfSAoJHtGb3JtYXQucGN0KGNhcmQud2Vha2VzdC5wY3REaWZmKX0pYCwgJ3BjYy1mb290ZXItZGV0YWlsJykpOwogICAgfQogICAgcmV0dXJuIGZvb3RlcjsKICB9CgogIC8qKiBTZWxmLWNvbnRhaW5lZCBtb2RhbCBmb3IgIlZpZXcgRnVsbCBDb21wYXJpc29uIiDigJQgYSBzZXBhcmF0ZSBvdmVybGF5IGlkIGZyb20gdGhlIERhdGEgUmVjb3JkcyBFZGl0IG1vZGFsIChSZWNvcmRzLm1vZGFsU2hlbGwgaXMgYSBwcml2YXRlIGNsb3N1cmUgb2YgdGhhdCBtb2R1bGUsIG5vdCBzaGFyZWQgc3RhdGUpLCBzYW1lIHZpc3VhbCBsYW5ndWFnZSAoLm1vZGFsLW92ZXJsYXkgLyAubW9kYWwtcGFuZWwpIHNvIGl0IGxvb2tzIGlkZW50aWNhbC4gKi8KICBmdW5jdGlvbiBjbG9zZUNhcmRNb2RhbCgpIHsKICAgIGNvbnN0IG92ZXJsYXkgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnY29tcGFyaXNvbk1vZGFsT3ZlcmxheScpOwogICAgaWYgKG92ZXJsYXkpIG92ZXJsYXkucmVtb3ZlKCk7CiAgfQoKICBmdW5jdGlvbiBvcGVuQ2FyZE1vZGFsKGNhcmQsIGxhYmVsQSwgbGFiZWxCKSB7CiAgICBjbG9zZUNhcmRNb2RhbCgpOwogICAgY29uc3Qgb3ZlcmxheSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgb3ZlcmxheS5jbGFzc05hbWUgPSAnbW9kYWwtb3ZlcmxheSc7CiAgICBvdmVybGF5LmlkID0gJ2NvbXBhcmlzb25Nb2RhbE92ZXJsYXknOwogICAgb3ZlcmxheS5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIChlKSA9PiB7IGlmIChlLnRhcmdldCA9PT0gb3ZlcmxheSkgY2xvc2VDYXJkTW9kYWwoKTsgfSk7CiAgICBjb25zdCBwYW5lbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcGFuZWwuY2xhc3NOYW1lID0gJ21vZGFsLXBhbmVsIHdpZGUnOwogICAgcGFuZWwuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMicsIGAke2NhcmQubGFiZWx9IOKAlCBGdWxsIENvbXBhcmlzb25gKSk7CiAgICBwYW5lbC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIGAke2xhYmVsQX0gdnMgJHtsYWJlbEJ9YCwgJ21vZGFsLXN1YicpKTsKCiAgICBjb25zdCB0YWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICB0YWJsZS5jbGFzc05hbWUgPSAnZGF0YS10YWJsZSc7CiAgICBjb25zdCB0aGVhZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RoZWFkJyk7CiAgICB0aGVhZC5pbm5lckhUTUwgPSAnPHRyPjx0aD5NZXRyaWM8L3RoPjx0aCBjbGFzcz0ibnVtIj5SYW5nZSBBPC90aD48dGggY2xhc3M9Im51bSI+UmFuZ2UgQjwvdGg+PHRoIGNsYXNzPSJudW0iPkRpZmZlcmVuY2U8L3RoPjx0aCBjbGFzcz0ibnVtIj4lIERpZmZlcmVuY2U8L3RoPjx0aD5UcmVuZDwvdGg+PC90cj4nOwogICAgY29uc3QgdGJvZHkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0Ym9keScpOwogICAgY2FyZC5tZXRyaWNzLmZvckVhY2goKG0pID0+IHsKICAgICAgY29uc3QgZm10ID0gKHYpID0+IChtLmlzRHVyYXRpb24gPyBGb3JtYXQuZHVyYXRpb24odikgOiBGb3JtYXQuc21hcnQodikpOwogICAgICBjb25zdCB0ciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAgIGNvbnN0IHRyZW5kRWwgPSB0ZXh0RWwoJ3NwYW4nLCB0cmVuZERpcmVjdGlvbihtLnBjdERpZmYpID09PSAndXAnID8gJ+KWsicgOiB0cmVuZERpcmVjdGlvbihtLnBjdERpZmYpID09PSAnZG93bicgPyAn4pa8JyA6ICfigJQnLCBgc3RhdC1kZWx0YSAke0Zvcm1hdC5kZWx0YUNsYXNzKG0ucGN0RGlmZil9YCk7CiAgICAgIGNvbnN0IHRyZW5kVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICB0cmVuZFRkLmFwcGVuZENoaWxkKHRyZW5kRWwpOwogICAgICB0ci5hcHBlbmQoCiAgICAgICAgdGV4dEVsKCd0ZCcsIG0ubGFiZWwpLAogICAgICAgIHRleHRFbCgndGQnLCBmbXQobS5hKSwgJ251bScpLAogICAgICAgIHRleHRFbCgndGQnLCBmbXQobS5iKSwgJ251bScpLAogICAgICAgIHRleHRFbCgndGQnLCBgJHttLmRpZmYgPiAwID8gJysnIDogJyd9JHtmbXQobS5kaWZmKX1gLCAnbnVtJyksCiAgICAgICAgdGV4dEVsKCd0ZCcsIG0ucGN0RGlmZiA9PT0gbnVsbCA/ICduZXcnIDogRm9ybWF0LnBjdChtLnBjdERpZmYpLCAnbnVtJyksCiAgICAgICAgdHJlbmRUZAogICAgICApOwogICAgICB0Ym9keS5hcHBlbmRDaGlsZCh0cik7CiAgICB9KTsKICAgIHRhYmxlLmFwcGVuZCh0aGVhZCwgdGJvZHkpOwogICAgY29uc3QgdGFibGVTY3JvbGwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRhYmxlU2Nyb2xsLmNsYXNzTmFtZSA9ICd0YWJsZS1zY3JvbGwnOwogICAgdGFibGVTY3JvbGwuYXBwZW5kQ2hpbGQodGFibGUpOwogICAgcGFuZWwuYXBwZW5kQ2hpbGQodGFibGVTY3JvbGwpOwoKICAgIGNvbnN0IGFjdGlvbnMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGFjdGlvbnMuY2xhc3NOYW1lID0gJ21vZGFsLWFjdGlvbnMnOwogICAgY29uc3QgY2xvc2VCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIGNsb3NlQnRuLmNsYXNzTmFtZSA9ICdidG4nOwogICAgY2xvc2VCdG4udHlwZSA9ICdidXR0b24nOwogICAgY2xvc2VCdG4udGV4dENvbnRlbnQgPSAnQ2xvc2UnOwogICAgY2xvc2VCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBjbG9zZUNhcmRNb2RhbCk7CiAgICBhY3Rpb25zLmFwcGVuZENoaWxkKGNsb3NlQnRuKTsKICAgIHBhbmVsLmFwcGVuZENoaWxkKGFjdGlvbnMpOwoKICAgIG92ZXJsYXkuYXBwZW5kQ2hpbGQocGFuZWwpOwogICAgZG9jdW1lbnQuYm9keS5hcHBlbmRDaGlsZChvdmVybGF5KTsKICB9CgogIGZ1bmN0aW9uIGJ1aWxkUGxhdGZvcm1DYXJkKGNhcmQsIGxhYmVsQSwgbGFiZWxCKSB7CiAgICBjb25zdCBlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZWwuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLWNvbXBhcmUtY2FyZCc7CgogICAgY29uc3QgaGVhZGVyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBoZWFkZXIuY2xhc3NOYW1lID0gJ3BjYy1oZWFkZXInOwogICAgY29uc3QgbmFtZVdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIG5hbWVXcmFwLmNsYXNzTmFtZSA9ICdwY2MtaGVhZGVyLW5hbWUnOwogICAgY29uc3QgZG90ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgZG90LmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1kb3QnOwogICAgZG90LnN0eWxlLmJhY2tncm91bmQgPSBjYXJkLmNvbG9yOwogICAgbmFtZVdyYXAuYXBwZW5kKGRvdCwgdGV4dEVsKCdzcGFuJywgY2FyZC5sYWJlbCwgJ3BjYy1uYW1lJykpOwogICAgaGVhZGVyLmFwcGVuZENoaWxkKG5hbWVXcmFwKTsKICAgIGNvbnN0IGRpciA9IHRyZW5kRGlyZWN0aW9uKGNhcmQub3ZlcmFsbCk7CiAgICBoZWFkZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBjYXJkLm92ZXJhbGwgPT09IG51bGwgPyAn4oCUJyA6IGAke2RpciA9PT0gJ3VwJyA/ICfilrInIDogZGlyID09PSAnZG93bicgPyAn4pa8JyA6ICfigJQnfSAke0Zvcm1hdC5wY3QoY2FyZC5vdmVyYWxsKX1gLCBgcGNjLWJhZGdlICR7ZGlyfWApKTsKICAgIGVsLmFwcGVuZENoaWxkKGhlYWRlcik7CgogICAgY29uc3QgY2FwdGlvbiA9IGNhcmQub3ZlcmFsbCA9PT0gbnVsbAogICAgICA/ICdOb3QgZW5vdWdoIGRhdGEgdG8gY29tcGFyZSB5ZXQnCiAgICAgIDogZGlyID09PSAndXAnID8gJ0ltcHJvdmVkIGNvbXBhcmVkIHRvIHByZXZpb3VzIHBlcmlvZCcKICAgICAgOiBkaXIgPT09ICdkb3duJyA/ICdMb3dlciB0aGFuIHByZXZpb3VzIHBlcmlvZCcKICAgICAgOiAnQWJvdXQgdGhlIHNhbWUgYXMgdGhlIHByZXZpb3VzIHBlcmlvZCc7CiAgICBlbC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIGNhcHRpb24sICdwY2MtY2FwdGlvbicpKTsKCiAgICBjb25zdCBtZXRyaWNzV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgbWV0cmljc1dyYXAuY2xhc3NOYW1lID0gJ3BjYy1tZXRyaWNzJzsKICAgIGNhcmQubWV0cmljcy5mb3JFYWNoKChtKSA9PiBtZXRyaWNzV3JhcC5hcHBlbmRDaGlsZChidWlsZE1ldHJpY1JvdyhtKSkpOwogICAgZWwuYXBwZW5kQ2hpbGQobWV0cmljc1dyYXApOwoKICAgIGVsLmFwcGVuZENoaWxkKGJ1aWxkQ2FyZEZvb3RlcihjYXJkKSk7CgogICAgY29uc3Qgdmlld0xpbmsgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIHZpZXdMaW5rLnR5cGUgPSAnYnV0dG9uJzsKICAgIHZpZXdMaW5rLmNsYXNzTmFtZSA9ICdwY2Mtdmlldy1saW5rJzsKICAgIHZpZXdMaW5rLnRleHRDb250ZW50ID0gJ1ZpZXcgRnVsbCBDb21wYXJpc29uIOKGkic7CiAgICB2aWV3TGluay5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IG9wZW5DYXJkTW9kYWwoY2FyZCwgbGFiZWxBLCBsYWJlbEIpKTsKICAgIGVsLmFwcGVuZENoaWxkKHZpZXdMaW5rKTsKCiAgICByZXR1cm4gZWw7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJQbGF0Zm9ybUNvbXBhcmlzb25DYXJkcyh3cmFwLCByZXN1bHQsIGxhYmVsQSwgbGFiZWxCKSB7CiAgICBjb25zdCBhbGxDYXJkcyA9IGJ1aWxkUGxhdGZvcm1DYXJkcyhyZXN1bHQpOwoKICAgIGNvbnN0IHNlY3Rpb24gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHNlY3Rpb24uY2xhc3NOYW1lID0gJ3BjYy1zZWN0aW9uJzsKICAgIHNlY3Rpb24uYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCAnUGxhdGZvcm0gUGVyZm9ybWFuY2UgQ29tcGFyaXNvbicsICdzZWN0aW9uLXRpdGxlJykpOwoKICAgIGlmICghYWxsQ2FyZHMubGVuZ3RoKSB7CiAgICAgIHNlY3Rpb24uYXBwZW5kQ2hpbGQoZW1wdHlTdGF0ZSh7CiAgICAgICAgaWNvbjogJ2dpdC1jb21wYXJlJywKICAgICAgICB0aXRsZTogJ05vIGRhdGEgYXZhaWxhYmxlIGZvciB0aGUgc2VsZWN0ZWQgZGF0ZSByYW5nZXMuJywKICAgICAgICBtZXNzYWdlOiAnVHJ5IGEgd2lkZXIgcmFuZ2UsIG9yIGNoZWNrIHRoYXQgcG9zdHMgZXhpc3QgZm9yIGF0IGxlYXN0IG9uZSBwbGF0Zm9ybSBpbiBSYW5nZSBBIG9yIFJhbmdlIEIuJywKICAgICAgfSkpOwogICAgICB3cmFwLmFwcGVuZENoaWxkKHNlY3Rpb24pOwogICAgICByZXR1cm47CiAgICB9CgogICAgY29uc3QgY29udHJvbHMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNvbnRyb2xzLmNsYXNzTmFtZSA9ICdwY2MtY29udHJvbHMnOwoKICAgIGNvbnN0IHNvcnRTZWxlY3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzZWxlY3QnKTsKICAgIENBUkRfU09SVF9NT0RFUy5mb3JFYWNoKChtKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOyBvcHQudmFsdWUgPSBtLmtleTsgb3B0LnRleHRDb250ZW50ID0gbS5sYWJlbDsKICAgICAgaWYgKG0ua2V5ID09PSBjYXJkU29ydE1vZGUpIG9wdC5zZWxlY3RlZCA9IHRydWU7CiAgICAgIHNvcnRTZWxlY3QuYXBwZW5kQ2hpbGQob3B0KTsKICAgIH0pOwogICAgc29ydFNlbGVjdC5hZGRFdmVudExpc3RlbmVyKCdjaGFuZ2UnLCAoKSA9PiB7IGNhcmRTb3J0TW9kZSA9IHNvcnRTZWxlY3QudmFsdWU7IHJlbmRlckNhcmRHcmlkKCk7IH0pOwogICAgY29udHJvbHMuYXBwZW5kQ2hpbGQobGFiZWxlZCgnU29ydCBCeScsIHNvcnRTZWxlY3QpKTsKCiAgICBjb25zdCBmaWx0ZXJQaWxscyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZmlsdGVyUGlsbHMuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLWZpbHRlci1waWxscyc7CiAgICBjb25zdCBhbGxCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIGFsbEJ0bi50eXBlID0gJ2J1dHRvbic7CiAgICBhbGxCdG4uZGF0YXNldC5maWx0ZXIgPSAnYWxsJzsKICAgIGFsbEJ0bi50ZXh0Q29udGVudCA9ICdBbGwgUGxhdGZvcm1zJzsKICAgIGFsbEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsgY2FyZFBsYXRmb3JtRmlsdGVyID0gJ2FsbCc7IHJlbmRlckNhcmRHcmlkKCk7IH0pOwogICAgZmlsdGVyUGlsbHMuYXBwZW5kQ2hpbGQoYWxsQnRuKTsKICAgIGFsbENhcmRzLmZvckVhY2goKGNhcmQpID0+IHsKICAgICAgY29uc3QgYnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICAgIGJ0bi50eXBlID0gJ2J1dHRvbic7CiAgICAgIGJ0bi5kYXRhc2V0LmZpbHRlciA9IGNhcmQucGxhdGZvcm07CiAgICAgIGNvbnN0IGRvdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgICAgZG90LmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1kb3QnOwogICAgICBkb3Quc3R5bGUuYmFja2dyb3VuZCA9IGNhcmQuY29sb3I7CiAgICAgIGJ0bi5hcHBlbmQoZG90LCBkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShjYXJkLmxhYmVsKSk7CiAgICAgIGJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsgY2FyZFBsYXRmb3JtRmlsdGVyID0gY2FyZC5wbGF0Zm9ybTsgcmVuZGVyQ2FyZEdyaWQoKTsgfSk7CiAgICAgIGZpbHRlclBpbGxzLmFwcGVuZENoaWxkKGJ0bik7CiAgICB9KTsKICAgIGNvbnRyb2xzLmFwcGVuZENoaWxkKGZpbHRlclBpbGxzKTsKICAgIHNlY3Rpb24uYXBwZW5kQ2hpbGQoY29udHJvbHMpOwoKICAgIGNvbnN0IGdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGdyaWQuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLWNvbXBhcmUtZ3JpZCc7CiAgICBncmlkLmlkID0gJ3BsYXRmb3JtQ29tcGFyZUdyaWQnOwogICAgc2VjdGlvbi5hcHBlbmRDaGlsZChncmlkKTsKICAgIHdyYXAuYXBwZW5kQ2hpbGQoc2VjdGlvbik7CgogICAgZnVuY3Rpb24gcmVuZGVyQ2FyZEdyaWQoKSB7CiAgICAgIGNvbnN0IGdyaWRFbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdwbGF0Zm9ybUNvbXBhcmVHcmlkJyk7CiAgICAgIGlmICghZ3JpZEVsKSByZXR1cm47CiAgICAgIGdyaWRFbC5pbm5lckhUTUwgPSAnJzsKICAgICAgY29uc3QgdmlzaWJsZSA9IGNhcmRQbGF0Zm9ybUZpbHRlciA9PT0gJ2FsbCcgPyBhbGxDYXJkcyA6IGFsbENhcmRzLmZpbHRlcigoYykgPT4gYy5wbGF0Zm9ybSA9PT0gY2FyZFBsYXRmb3JtRmlsdGVyKTsKICAgICAgY29uc3Qgc29ydGVkID0gc29ydENhcmRzKHZpc2libGUsIGNhcmRTb3J0TW9kZSk7CiAgICAgIGlmICghc29ydGVkLmxlbmd0aCkgewogICAgICAgIGdyaWRFbC5hcHBlbmRDaGlsZChlbXB0eVN0YXRlKHsgaWNvbjogJ2dpdC1jb21wYXJlJywgbWVzc2FnZTogJ05vIGRhdGEgZm9yIHRoaXMgcGxhdGZvcm0gaW4gdGhlIHNlbGVjdGVkIGRhdGUgcmFuZ2VzLicgfSkpOwogICAgICB9IGVsc2UgewogICAgICAgIHNvcnRlZC5mb3JFYWNoKChjYXJkKSA9PiBncmlkRWwuYXBwZW5kQ2hpbGQoYnVpbGRQbGF0Zm9ybUNhcmQoY2FyZCwgbGFiZWxBLCBsYWJlbEIpKSk7CiAgICAgIH0KICAgICAgZmlsdGVyUGlsbHMucXVlcnlTZWxlY3RvckFsbCgnYnV0dG9uJykuZm9yRWFjaCgoYnRuKSA9PiB7CiAgICAgICAgYnRuLmNsYXNzTGlzdC50b2dnbGUoJ2lzLWFjdGl2ZScsIGJ0bi5kYXRhc2V0LmZpbHRlciA9PT0gY2FyZFBsYXRmb3JtRmlsdGVyKTsKICAgICAgfSk7CiAgICB9CiAgICByZW5kZXJDYXJkR3JpZCgpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyKCkgewogICAgcm9vdCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd2aWV3LWNvbXBhcmlzb24nKTsKICAgIHNoZWxsKCk7CiAgfQoKICByZXR1cm4geyByZW5kZXIgfTsKfSkoKTsKCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBVcGxvYWQgdGFiOiBkcmFnLWRyb3AsIHZhbGlkYXRpb24gcHJldmlldywgcGVyLXdlZWsgY29uZmxpY3QKICAgcmVzb2x1dGlvbiwgY29tbWl0IOKAlCBwbHVzIHRoZSBVcGxvYWQgSGlzdG9yeSB0YWIuCiAgID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwpjb25zdCBVcGxvYWQgPSAoKCkgPT4gewogIGxldCByb290OwogIGxldCBjdXJyZW50UHJldmlldyA9IG51bGw7IC8vIHsgZmlsZVBhdGgsIG9yaWdpbmFsTmFtZSwgZHVwbGljYXRlcywgaXNzdWVzLCBzYW1wbGUsIC4uLiB9CiAgY29uc3QgZHVwbGljYXRlQWN0aW9uT3ZlcnJpZGVzID0ge307CgogIGZ1bmN0aW9uIHNoZWxsKCkgewogICAgcm9vdC5pbm5lckhUTUwgPSAnJzsKCiAgICBjb25zdCBpbnRybyA9IHRleHRFbCgnZGl2JywgJ1VwbG9hZCBhIHdlZWtseSBleHBvcnQnLCAnc2VjdGlvbi10aXRsZScpOwogICAgcm9vdC5hcHBlbmRDaGlsZChpbnRybyk7CgogICAgY29uc3QgZHJvcHpvbmUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGRyb3B6b25lLmNsYXNzTmFtZSA9ICdkcm9wem9uZSc7CiAgICBkcm9wem9uZS5pZCA9ICdkcm9wem9uZSc7CiAgICBkcm9wem9uZS5pbm5lckhUTUwgPSBgCiAgICAgIDxkaXYgY2xhc3M9ImVtcHR5LWljb24iIHN0eWxlPSJtYXJnaW46IDAgYXV0byAxNHB4OyI+PGkgZGF0YS1sdWNpZGU9InVwbG9hZC1jbG91ZCIgc3R5bGU9IndpZHRoOjIycHg7aGVpZ2h0OjIycHg7Ij48L2k+PC9kaXY+CiAgICAgIDxoMz5EcmFnICZhbXA7IGRyb3AgeW91ciAuY3N2IG9yIC54bHN4IGZpbGUgaGVyZTwvaDM+CiAgICAgIDxwPm9yIGNsaWNrIHRvIGJyb3dzZSDigJQgZmlsZXMgYXJlIHZhbGlkYXRlZCBiZWZvcmUgYW55dGhpbmcgaXMgc2F2ZWQ8L3A+CiAgICAgIDxpbnB1dCB0eXBlPSJmaWxlIiBpZD0iZmlsZUlucHV0IiBhY2NlcHQ9Ii5jc3YsLnhsc3gsLnhscyIgLz4KICAgIGA7CiAgICByb290LmFwcGVuZENoaWxkKGRyb3B6b25lKTsKCiAgICBjb25zdCBwcmV2aWV3QXJlYSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcHJldmlld0FyZWEuaWQgPSAncHJldmlld0FyZWEnOwogICAgcm9vdC5hcHBlbmRDaGlsZChwcmV2aWV3QXJlYSk7CgogICAgd2lyZURyb3B6b25lKGRyb3B6b25lKTsKICB9CgogIGZ1bmN0aW9uIHdpcmVEcm9wem9uZShkcm9wem9uZSkgewogICAgY29uc3QgaW5wdXQgPSBkcm9wem9uZS5xdWVyeVNlbGVjdG9yKCcjZmlsZUlucHV0Jyk7CiAgICBkcm9wem9uZS5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IGlucHV0LmNsaWNrKCkpOwogICAgaW5wdXQuYWRkRXZlbnRMaXN0ZW5lcignY2hhbmdlJywgKCkgPT4gewogICAgICBpZiAoaW5wdXQuZmlsZXNbMF0pIGhhbmRsZUZpbGUoaW5wdXQuZmlsZXNbMF0pOwogICAgfSk7CiAgICBbJ2RyYWdlbnRlcicsICdkcmFnb3ZlciddLmZvckVhY2goKGV2dCkgPT4KICAgICAgZHJvcHpvbmUuYWRkRXZlbnRMaXN0ZW5lcihldnQsIChlKSA9PiB7IGUucHJldmVudERlZmF1bHQoKTsgZHJvcHpvbmUuY2xhc3NMaXN0LmFkZCgnaXMtZHJhZycpOyB9KQogICAgKTsKICAgIFsnZHJhZ2xlYXZlJywgJ2Ryb3AnXS5mb3JFYWNoKChldnQpID0+CiAgICAgIGRyb3B6b25lLmFkZEV2ZW50TGlzdGVuZXIoZXZ0LCAoZSkgPT4geyBlLnByZXZlbnREZWZhdWx0KCk7IGRyb3B6b25lLmNsYXNzTGlzdC5yZW1vdmUoJ2lzLWRyYWcnKTsgfSkKICAgICk7CiAgICBkcm9wem9uZS5hZGRFdmVudExpc3RlbmVyKCdkcm9wJywgKGUpID0+IHsKICAgICAgY29uc3QgZmlsZSA9IGUuZGF0YVRyYW5zZmVyLmZpbGVzWzBdOwogICAgICBpZiAoZmlsZSkgaGFuZGxlRmlsZShmaWxlKTsKICAgIH0pOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gaGFuZGxlRmlsZShmaWxlKSB7CiAgICBjb25zdCBhcmVhID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3ByZXZpZXdBcmVhJyk7CiAgICBhcmVhLmlubmVySFRNTCA9ICcnOwogICAgYXJlYS5hcHBlbmRDaGlsZChyb3dXaXRoU3Bpbm5lcignVmFsaWRhdGluZyBmaWxl4oCmJykpOwogICAgT2JqZWN0LmtleXMoZHVwbGljYXRlQWN0aW9uT3ZlcnJpZGVzKS5mb3JFYWNoKChrKSA9PiBkZWxldGUgZHVwbGljYXRlQWN0aW9uT3ZlcnJpZGVzW2tdKTsKICAgIHRyeSB7CiAgICAgIGN1cnJlbnRQcmV2aWV3ID0gYXdhaXQgQXBpLnByZXZpZXdVcGxvYWQoZmlsZSk7CiAgICAgIHJlbmRlclByZXZpZXcoKTsKICAgIH0gY2F0Y2ggKGVycikgewogICAgICBhcmVhLmlubmVySFRNTCA9ICcnOwogICAgICBhcmVhLmFwcGVuZENoaWxkKGVycm9yQmFubmVyKGVyci5tZXNzYWdlKSk7CiAgICB9CiAgfQoKICBmdW5jdGlvbiByb3dXaXRoU3Bpbm5lcih0ZXh0KSB7CiAgICBjb25zdCBlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZWwuY2xhc3NOYW1lID0gJ2xvYWRpbmctcm93JzsKICAgIGNvbnN0IHNwaW5uZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICBzcGlubmVyLmNsYXNzTmFtZSA9ICdzcGlubmVyJzsKICAgIGVsLmFwcGVuZChzcGlubmVyLCBkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShgICR7dGV4dH1gKSk7CiAgICByZXR1cm4gZWw7CiAgfQogIGZ1bmN0aW9uIGVycm9yQmFubmVyKG1lc3NhZ2UpIHsKICAgIGNvbnN0IGVsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBlbC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICBlbC5zdHlsZS5ib3JkZXJMZWZ0ID0gJzNweCBzb2xpZCB2YXIoLS1zdGF0dXMtY3JpdGljYWwpJzsKICAgIGVsLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgYENvdWxkIG5vdCByZWFkIHRoaXMgZmlsZTogJHttZXNzYWdlfWAsICdtdXRlZCcpKTsKICAgIHJldHVybiBlbDsKICB9CgogIGZ1bmN0aW9uIHJlbmRlclByZXZpZXcoKSB7CiAgICBjb25zdCBhcmVhID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3ByZXZpZXdBcmVhJyk7CiAgICBhcmVhLmlubmVySFRNTCA9ICcnOwogICAgY29uc3QgcCA9IGN1cnJlbnRQcmV2aWV3OwoKICAgIGNvbnN0IHN1bW1hcnlUaXRsZSA9IHRleHRFbCgnZGl2JywgJ1ZhbGlkYXRpb24gc3VtbWFyeScsICdzZWN0aW9uLXRpdGxlJyk7CiAgICBjb25zdCBzdW1tYXJ5R3JpZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgc3VtbWFyeUdyaWQuY2xhc3NOYW1lID0gJ3N0YXQtZ3JpZCc7CiAgICBzdW1tYXJ5R3JpZC5hcHBlbmQoCiAgICAgIHN0YXRUaWxlKCdGaWxlJywgcC5vcmlnaW5hbE5hbWUpLAogICAgICBzdGF0VGlsZSgnU2hlZXRzIGZvdW5kJywgcC5zaGVldHMubGVuZ3RoKSwKICAgICAgc3RhdFRpbGUoJ1RvdGFsIHJvd3MgKGFsbCBzaGVldHMpJywgcC50b3RhbERhdGFSb3dzKSwKICAgICAgc3RhdFRpbGUoJ05ldyByZWNvcmRzJywgcC5uZXdSZWNvcmRzQ291bnQpLAogICAgICBzdGF0VGlsZSgnRXhhY3QgZHVwbGljYXRlcyBmb3VuZCcsIHAuZHVwbGljYXRlcy5sZW5ndGgpLAogICAgICBzdGF0VGlsZSgnRHVwbGljYXRlIHJvd3MgaW4gZmlsZScsIHAuZHVwbGljYXRlUm93c0luRmlsZSksCiAgICAgIHN0YXRUaWxlKCdSb3dzIHdpdGggZXJyb3JzJywgcC5lcnJvclJvd3MpCiAgICApOwogICAgYXJlYS5hcHBlbmQoc3VtbWFyeVRpdGxlLCBzdW1tYXJ5R3JpZCk7CgogICAgaWYgKHAuc2hlZXRzLmxlbmd0aCkgewogICAgICBjb25zdCBzaGVldHNUaXRsZSA9IHRleHRFbCgnZGl2JywgJ1NoZWV0IGJyZWFrZG93bicsICdzZWN0aW9uLXRpdGxlJyk7CiAgICAgIGNvbnN0IHNoZWV0c1RhYmxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGFibGUnKTsKICAgICAgc2hlZXRzVGFibGUuY2xhc3NOYW1lID0gJ2RhdGEtdGFibGUnOwogICAgICBzaGVldHNUYWJsZS5pbm5lckhUTUwgPSAnPHRoZWFkPjx0cj48dGg+U2hlZXQ8L3RoPjx0aD5MYXlvdXQgZGV0ZWN0ZWQ8L3RoPjx0aCBjbGFzcz0ibnVtIj5Sb3dzPC90aD48dGggY2xhc3M9Im51bSI+VmFsaWQ8L3RoPjx0aCBjbGFzcz0ibnVtIj5FcnJvcnM8L3RoPjwvdHI+PC90aGVhZD4nOwogICAgICBjb25zdCBzaGVldHNCb2R5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGJvZHknKTsKICAgICAgcC5zaGVldHMuZm9yRWFjaCgocykgPT4gewogICAgICAgIGNvbnN0IHRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgICBjb25zdCBsYXlvdXRMYWJlbCA9IHMuZm9ybWF0ID09PSAnYWdlbmRhJyA/ICdMUlMgYWdlbmRhIHRyYWNrZXInIDogcy5mb3JtYXQgPT09ICdzaW1wbGUnID8gJ1NpbXBsZSBwbGF0Zm9ybSB0YWJsZScgOiAnTm90IHJlY29nbml6ZWQg4oCUIHNhdmVkIGFzIHJhdyBkYXRhIG9ubHknOwogICAgICAgIHRyLmFwcGVuZCgKICAgICAgICAgIHRleHRFbCgndGQnLCBzLm5hbWUpLAogICAgICAgICAgdGV4dEVsKCd0ZCcsIGxheW91dExhYmVsKSwKICAgICAgICAgIHRleHRFbCgndGQnLCBTdHJpbmcocy50b3RhbFJvd3MpLCAnbnVtJyksCiAgICAgICAgICB0ZXh0RWwoJ3RkJywgU3RyaW5nKHMudmFsaWRSb3dzKSwgJ251bScpLAogICAgICAgICAgdGV4dEVsKCd0ZCcsIFN0cmluZyhzLmVycm9yUm93cyksICdudW0nKQogICAgICAgICk7CiAgICAgICAgc2hlZXRzQm9keS5hcHBlbmRDaGlsZCh0cik7CiAgICAgIH0pOwogICAgICBzaGVldHNUYWJsZS5hcHBlbmRDaGlsZChzaGVldHNCb2R5KTsKICAgICAgY29uc3Qgc2hlZXRzV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBzaGVldHNXcmFwLmNsYXNzTmFtZSA9ICd0YWJsZS1zY3JvbGwnOwogICAgICBzaGVldHNXcmFwLmFwcGVuZENoaWxkKHNoZWV0c1RhYmxlKTsKICAgICAgYXJlYS5hcHBlbmQoc2hlZXRzVGl0bGUsIHNoZWV0c1dyYXApOwogICAgfQoKICAgIGlmIChwLmR1cGxpY2F0ZXMubGVuZ3RoKSB7CiAgICAgIGNvbnN0IGR1cFRpdGxlID0gdGV4dEVsKCdkaXYnLCBgRXhhY3QgZHVwbGljYXRlcyBmb3VuZCAoJHtwLmR1cGxpY2F0ZXMubGVuZ3RofSlgLCAnc2VjdGlvbi10aXRsZScpOwogICAgICBhcmVhLmFwcGVuZENoaWxkKGR1cFRpdGxlKTsKICAgICAgYXJlYS5hcHBlbmRDaGlsZCh0ZXh0RWwoCiAgICAgICAgJ2RpdicsCiAgICAgICAgJ0VhY2ggb2YgdGhlc2Ugcm93cyBpcyBieXRlLWZvci1ieXRlIGlkZW50aWNhbCB0byBhbiBhbHJlYWR5LXNhdmVkIHJlY29yZCDigJQgZXZlcnkgZmllbGQgbWF0Y2hlcywgaW5jbHVkaW5nIGV2ZXJ5IG1ldHJpYywgbm90IGp1c3QgdGhlIGRhdGUvY2FwdGlvbi9wbGF0Zm9ybS4gQ2hvb3NlIHdoYXQgdG8gZG8gd2l0aCBlYWNoIOKAlCBvciBzZXQgYSBkZWZhdWx0IGZvciBhbGwgb2YgdGhlbS4gKEEgcm93IHRoYXQgc2hhcmVzIHRoZSBzYW1lIGRhdGUvY2FwdGlvbi9wbGF0Zm9ybSBidXQgaGFzIGRpZmZlcmVudCBudW1iZXJzIGlzIG5vdCBzaG93biBoZXJlIOKAlCBpdOKAmXMgaW1wb3J0ZWQgYXV0b21hdGljYWxseSBhcyBpdHMgb3duIG5ldyByZWNvcmQsIHNpbmNlIGl0cyBhbmFseXRpY3MgY2hhbmdlZC4pJywKICAgICAgICAnbXV0ZWQnCiAgICAgICkpOwogICAgICBjb25zdCBkZWZhdWx0Um93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIGRlZmF1bHRSb3cuY2xhc3NOYW1lID0gJ2ZpZWxkLWlubGluZSc7CiAgICAgIGRlZmF1bHRSb3cuc3R5bGUubWFyZ2luID0gJzEwcHggMCc7CiAgICAgIGNvbnN0IGRlZmF1bHRTZWxlY3QgPSBhY3Rpb25TZWxlY3QoJ3NraXAnKTsKICAgICAgZGVmYXVsdFNlbGVjdC5pZCA9ICdkZWZhdWx0RHVwbGljYXRlQWN0aW9uU2VsZWN0JzsKICAgICAgZGVmYXVsdFNlbGVjdC5hZGRFdmVudExpc3RlbmVyKCdjaGFuZ2UnLCAoKSA9PiB7CiAgICAgICAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnLmNvbmZsaWN0LXJvdyBzZWxlY3RbZGF0YS1oYXNoXScpLmZvckVhY2goKHNlbCkgPT4gewogICAgICAgICAgaWYgKCFkdXBsaWNhdGVBY3Rpb25PdmVycmlkZXNbc2VsLmRhdGFzZXQuaGFzaF0pIHNlbC52YWx1ZSA9IGRlZmF1bHRTZWxlY3QudmFsdWU7CiAgICAgICAgfSk7CiAgICAgIH0pOwogICAgICBkZWZhdWx0Um93LmFwcGVuZCh0ZXh0RWwoJ2xhYmVsJywgJ0RlZmF1bHQgYWN0aW9uIGZvciBhbGwgbWF0Y2hlcycpLCBkZWZhdWx0U2VsZWN0KTsKICAgICAgYXJlYS5hcHBlbmRDaGlsZChkZWZhdWx0Um93KTsKCiAgICAgIGNvbnN0IGxpc3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgbGlzdC5jbGFzc05hbWUgPSAnY29uZmxpY3QtbGlzdCc7CiAgICAgIHAuZHVwbGljYXRlcy5mb3JFYWNoKChkKSA9PiB7CiAgICAgICAgY29uc3Qgcm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgICAgcm93LmNsYXNzTmFtZSA9ICdjb25mbGljdC1yb3cnOwogICAgICAgIGNvbnN0IGxlZnQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgICBsZWZ0LmFwcGVuZCgKICAgICAgICAgIHRleHRFbCgnZGl2JywgYCR7Rm9ybWF0LmRhdGUoZC5wdWJsaXNoRGF0ZSl9IOKAlCAkeyhkLmNhcHRpb24gfHwgJyhubyBjYXB0aW9uKScpLnNsaWNlKDAsIDcwKX1gLCAnd2Vlay1sYWJlbCcpLAogICAgICAgICAgdGV4dEVsKCdkaXYnLCBgRXhhY3QgbWF0Y2ggb2YgZXhpc3RpbmcgcmVjb3JkICMke2QuZXhpc3RpbmcucG9zdElkfSAobGFzdCB1cGRhdGVkICR7ZC5leGlzdGluZy51cGRhdGVkQXR9KWAsICd3ZWVrLW1ldGEnKQogICAgICAgICk7CiAgICAgICAgcm93LmFwcGVuZENoaWxkKGxlZnQpOwogICAgICAgIGNvbnN0IHNlbCA9IGFjdGlvblNlbGVjdCgnc2tpcCcpOwogICAgICAgIHNlbC5kYXRhc2V0Lmhhc2ggPSBkLmhhc2g7CiAgICAgICAgc2VsLmFkZEV2ZW50TGlzdGVuZXIoJ2NoYW5nZScsICgpID0+IHsgZHVwbGljYXRlQWN0aW9uT3ZlcnJpZGVzW2QuaGFzaF0gPSBzZWwudmFsdWU7IH0pOwogICAgICAgIHJvdy5hcHBlbmRDaGlsZChzZWwpOwogICAgICAgIGxpc3QuYXBwZW5kQ2hpbGQocm93KTsKICAgICAgfSk7CiAgICAgIGFyZWEuYXBwZW5kQ2hpbGQobGlzdCk7CiAgICB9CgogICAgY29uc3Qgbm90ZXNGaWVsZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgbm90ZXNGaWVsZC5jbGFzc05hbWUgPSAnZm9ybS1maWVsZCc7CiAgICBub3Rlc0ZpZWxkLnN0eWxlLm1hcmdpbiA9ICcxMnB4IDAnOwogICAgbm90ZXNGaWVsZC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2xhYmVsJywgJ1VwbG9hZCBub3RlcyAob3B0aW9uYWwpJykpOwogICAgY29uc3Qgbm90ZXNJbnB1dCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7CiAgICBub3Rlc0lucHV0LnR5cGUgPSAndGV4dCc7CiAgICBub3Rlc0lucHV0LmlkID0gJ3VwbG9hZE5vdGVzSW5wdXQnOwogICAgbm90ZXNJbnB1dC5wbGFjZWhvbGRlciA9ICdlLmcuICJXZWVrIDMgZXhwb3J0LCBpbmNsdWRlcyBjb3JyZWN0ZWQgVGlrVG9rIG51bWJlcnMiJzsKICAgIG5vdGVzRmllbGQuYXBwZW5kQ2hpbGQobm90ZXNJbnB1dCk7CiAgICBhcmVhLmFwcGVuZENoaWxkKG5vdGVzRmllbGQpOwoKICAgIGlmIChwLmlzc3Vlcy5sZW5ndGgpIHsKICAgICAgY29uc3QgaXNzdWVzVGl0bGUgPSB0ZXh0RWwoJ2RpdicsIGBSb3dzIHNraXBwZWQgb3IgZmxhZ2dlZCAoJHtwLmlzc3Vlcy5sZW5ndGh9KWAsICdzZWN0aW9uLXRpdGxlJyk7CiAgICAgIGNvbnN0IGlzc3Vlc0NhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgaXNzdWVzQ2FyZC5jbGFzc05hbWUgPSAnaXNzdWVzLWxpc3QnOwogICAgICBwLmlzc3Vlcy5mb3JFYWNoKChpc3N1ZSkgPT4gewogICAgICAgIGNvbnN0IHJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICAgIHJvdy5jbGFzc05hbWUgPSAnaXNzdWUtcm93JzsKICAgICAgICBpZiAoaXNzdWUucm93TnVtYmVyKSByb3cuYXBwZW5kQ2hpbGQodGV4dEVsKCdzcGFuJywgYFJvdyAke2lzc3VlLnJvd051bWJlcn1gLCAncm93LW5vJykpOwogICAgICAgIHJvdy5hcHBlbmRDaGlsZChkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShpc3N1ZS5tZXNzYWdlKSk7CiAgICAgICAgaXNzdWVzQ2FyZC5hcHBlbmRDaGlsZChyb3cpOwogICAgICB9KTsKICAgICAgYXJlYS5hcHBlbmQoaXNzdWVzVGl0bGUsIGlzc3Vlc0NhcmQpOwogICAgfQoKICAgIGlmIChwLnNhbXBsZS5sZW5ndGgpIHsKICAgICAgY29uc3Qgc2FtcGxlVGl0bGUgPSB0ZXh0RWwoJ2RpdicsICdTYW1wbGUgb2YgcGFyc2VkIHJvd3MnLCAnc2VjdGlvbi10aXRsZScpOwogICAgICBjb25zdCB0YWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICAgIHRhYmxlLmNsYXNzTmFtZSA9ICdkYXRhLXRhYmxlJzsKICAgICAgdGFibGUuaW5uZXJIVE1MID0gJzx0aGVhZD48dHI+PHRoPkRhdGU8L3RoPjx0aD5DYXB0aW9uPC90aD48dGg+VHlwZTwvdGg+PHRoPkNhbXBhaWduPC90aD48dGg+UGxhdGZvcm1zPC90aD48L3RyPjwvdGhlYWQ+JzsKICAgICAgY29uc3QgdGJvZHkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0Ym9keScpOwogICAgICBwLnNhbXBsZS5mb3JFYWNoKChzKSA9PiB7CiAgICAgICAgY29uc3QgdHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgICAgIHRyLmFwcGVuZCgKICAgICAgICAgIHRleHRFbCgndGQnLCBGb3JtYXQuZGF0ZShzLnB1Ymxpc2hEYXRlKSksCiAgICAgICAgICB0ZXh0RWwoJ3RkJywgcy5jYXB0aW9uIHx8ICfigJQnKSwKICAgICAgICAgIHRleHRFbCgndGQnLCBzLmNvbnRlbnRUeXBlIHx8ICfigJQnKSwKICAgICAgICAgIHRleHRFbCgndGQnLCBzLmNhbXBhaWduVHlwZSB8fCAnVW5zcGVjaWZpZWQnKSwKICAgICAgICAgIHRleHRFbCgndGQnLCBzLnBsYXRmb3Jtcy5qb2luKCcsICcpKQogICAgICAgICk7CiAgICAgICAgdGJvZHkuYXBwZW5kQ2hpbGQodHIpOwogICAgICB9KTsKICAgICAgdGFibGUuYXBwZW5kQ2hpbGQodGJvZHkpOwogICAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIHdyYXAuY2xhc3NOYW1lID0gJ3RhYmxlLXNjcm9sbCc7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQodGFibGUpOwogICAgICBhcmVhLmFwcGVuZChzYW1wbGVUaXRsZSwgd3JhcCk7CiAgICB9CgogICAgY29uc3QgYWN0aW9ucyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYWN0aW9ucy5jbGFzc05hbWUgPSAnYnRuLXJvdyc7CiAgICBhY3Rpb25zLnN0eWxlLm1hcmdpblRvcCA9ICcxNnB4JzsKICAgIGNvbnN0IGNvbW1pdEJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgY29tbWl0QnRuLmNsYXNzTmFtZSA9ICdidG4gcHJpbWFyeSc7CiAgICBjb21taXRCdG4udGV4dENvbnRlbnQgPSBwLnZhbGlkUm93cyA+IDAgPyBgSW1wb3J0ICR7cC52YWxpZFJvd3N9IHJvdyhzKWAgOiAnTm90aGluZyB0byBpbXBvcnQnOwogICAgY29tbWl0QnRuLmRpc2FibGVkID0gcC52YWxpZFJvd3MgPT09IDA7CiAgICBjb21taXRCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBjb21taXQoY29tbWl0QnRuKSk7CiAgICBjb25zdCBjYW5jZWxCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIGNhbmNlbEJ0bi5jbGFzc05hbWUgPSAnYnRuJzsKICAgIGNhbmNlbEJ0bi50ZXh0Q29udGVudCA9ICdDYW5jZWwnOwogICAgY2FuY2VsQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4geyBjdXJyZW50UHJldmlldyA9IG51bGw7IHNoZWxsKCk7IH0pOwogICAgYWN0aW9ucy5hcHBlbmQoY29tbWl0QnRuLCBjYW5jZWxCdG4pOwogICAgYXJlYS5hcHBlbmRDaGlsZChhY3Rpb25zKTsKICB9CgogIGZ1bmN0aW9uIHN0YXRUaWxlKGxhYmVsLCB2YWx1ZSkgewogICAgY29uc3QgdGlsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGlsZS5jbGFzc05hbWUgPSAnc3RhdC10aWxlJzsKICAgIHRpbGUuYXBwZW5kKHRleHRFbCgnZGl2JywgbGFiZWwsICdzdGF0LWxhYmVsJyksIHRleHRFbCgnZGl2JywgU3RyaW5nKHZhbHVlKSwgJ3N0YXQtdmFsdWUnKSk7CiAgICByZXR1cm4gdGlsZTsKICB9CiAgZnVuY3Rpb24gYWN0aW9uU2VsZWN0KGRlZmF1bHRWYWwpIHsKICAgIGNvbnN0IHNlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NlbGVjdCcpOwogICAgW1snc2tpcCcsICdTa2lwIChrZWVwIGV4aXN0aW5nIHJlY29yZCB1bmNoYW5nZWQpJ10sIFsndXBkYXRlJywgJ1VwZGF0ZSBleGlzdGluZyByZWNvcmQnXSwgWydjcmVhdGUnLCAnQ3JlYXRlIGFzIGEgbmV3LCBzZXBhcmF0ZSByZWNvcmQnXV0uZm9yRWFjaCgoW3YsIGxdKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOyBvcHQudmFsdWUgPSB2OyBvcHQudGV4dENvbnRlbnQgPSBsOwogICAgICBpZiAodiA9PT0gZGVmYXVsdFZhbCkgb3B0LnNlbGVjdGVkID0gdHJ1ZTsKICAgICAgc2VsLmFwcGVuZENoaWxkKG9wdCk7CiAgICB9KTsKICAgIHJldHVybiBzZWw7CiAgfQoKICBhc3luYyBmdW5jdGlvbiBjb21taXQoYnRuKSB7CiAgICBidG4uZGlzYWJsZWQgPSB0cnVlOwogICAgYnRuLnRleHRDb250ZW50ID0gJ0ltcG9ydGluZ+KApic7CiAgICBjb25zdCBkZWZhdWx0RHVwbGljYXRlQWN0aW9uID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2RlZmF1bHREdXBsaWNhdGVBY3Rpb25TZWxlY3QnKT8udmFsdWUgfHwgJ3NraXAnOwogICAgY29uc3Qgbm90ZXMgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndXBsb2FkTm90ZXNJbnB1dCcpPy52YWx1ZSB8fCBudWxsOwogICAgdHJ5IHsKICAgICAgY29uc3QgcmVzdWx0ID0gYXdhaXQgQXBpLmNvbW1pdFVwbG9hZCh7CiAgICAgICAgZmlsZVBhdGg6IGN1cnJlbnRQcmV2aWV3LmZpbGVQYXRoLAogICAgICAgIG9yaWdpbmFsTmFtZTogY3VycmVudFByZXZpZXcub3JpZ2luYWxOYW1lLAogICAgICAgIGRlZmF1bHREdXBsaWNhdGVBY3Rpb24sCiAgICAgICAgZHVwbGljYXRlQWN0aW9uczogZHVwbGljYXRlQWN0aW9uT3ZlcnJpZGVzLAogICAgICAgIG5vdGVzLAogICAgICB9KTsKICAgICAgVG9hc3Quc2hvdygKICAgICAgICBgSW1wb3J0ZWQ6ICR7cmVzdWx0LmltcG9ydGVkUm93c30gbmV3LCAke3Jlc3VsdC51cGRhdGVkUm93c30gdXBkYXRlZCwgJHtyZXN1bHQuc2tpcHBlZFJvd3N9IHNraXBwZWQuYCwKICAgICAgICByZXN1bHQuZXJyb3JDb3VudCA+IDAgPyAnZXJyb3InIDogJ3N1Y2Nlc3MnCiAgICAgICk7CiAgICAgIGN1cnJlbnRQcmV2aWV3ID0gbnVsbDsKICAgICAgc2hlbGwoKTsKICAgICAgd2luZG93LmRpc3BhdGNoRXZlbnQobmV3IEN1c3RvbUV2ZW50KCdscnM6ZGF0YS11cGRhdGVkJykpOwogICAgfSBjYXRjaCAoZXJyKSB7CiAgICAgIFRvYXN0LnNob3coZXJyLm1lc3NhZ2UsICdlcnJvcicpOwogICAgICBidG4uZGlzYWJsZWQgPSBmYWxzZTsKICAgICAgYnRuLnRleHRDb250ZW50ID0gJ1JldHJ5IGltcG9ydCc7CiAgICB9CiAgfQoKICBmdW5jdGlvbiByZW5kZXIoKSB7CiAgICByb290ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3ZpZXctdXBsb2FkJyk7CiAgICBzaGVsbCgpOwogIH0KCiAgcmV0dXJuIHsgcmVuZGVyIH07Cn0pKCk7Cgpjb25zdCBIaXN0b3J5ID0gKCgpID0+IHsKICBsZXQgcm9vdDsKICBsZXQgY3VycmVudFVwbG9hZHMgPSBbXTsKICBsZXQgc2VhcmNoVmFsdWUgPSAnJzsKICBsZXQgcGFnZSA9IDE7CiAgY29uc3QgcGFnZVNpemUgPSAxNTsKICBsZXQgc29ydFN0YXRlID0geyBrZXk6ICd1cGxvYWRlZF9hdCcsIGRpcjogJ2Rlc2MnLCB0eXBlOiAnc3RyaW5nJyB9OwogIGNvbnN0IEVYUE9SVF9DT0xVTU5TID0gWwogICAgeyBrZXk6ICdmaWxlbmFtZScsIGxhYmVsOiAnRmlsZScgfSwKICAgIHsga2V5OiAndXBsb2FkZWRfYXQnLCBsYWJlbDogJ1VwbG9hZGVkJyB9LAogICAgeyBrZXk6ICdzdGF0dXMnLCBsYWJlbDogJ1N0YXR1cycgfSwKICAgIHsga2V5OiAnaW1wb3J0ZWRfcm93cycsIGxhYmVsOiAnSW1wb3J0ZWQnIH0sCiAgICB7IGtleTogJ3VwZGF0ZWRfcm93cycsIGxhYmVsOiAnVXBkYXRlZCcgfSwKICAgIHsga2V5OiAnc2tpcHBlZF9yb3dzJywgbGFiZWw6ICdTa2lwcGVkJyB9LAogICAgeyBrZXk6ICdlcnJvcl9jb3VudCcsIGxhYmVsOiAnRXJyb3JzJyB9LAogICAgeyBrZXk6ICd3ZWVrcycsIGxhYmVsOiAnV2Vla3MnIH0sCiAgICB7IGtleTogJ25vdGVzJywgbGFiZWw6ICdOb3RlcycgfSwKICBdOwoKICBmdW5jdGlvbiBiYWRnZUNsYXNzKHN0YXR1cykgewogICAgaWYgKHN0YXR1cyA9PT0gJ3N1Y2Nlc3MnKSByZXR1cm4gJ3N1Y2Nlc3MnOwogICAgaWYgKHN0YXR1cyA9PT0gJ3BhcnRpYWwnKSByZXR1cm4gJ3BhcnRpYWwnOwogICAgcmV0dXJuICdmYWlsZWQnOwogIH0KCiAgZnVuY3Rpb24gc29ydGFibGVIZWFkZXIobGFiZWwsIGtleSwgdHlwZSkgewogICAgY29uc3QgdGggPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0aCcpOwogICAgaWYgKHR5cGUgPT09ICdudW1iZXInKSB0aC5jbGFzc05hbWUgPSAnbnVtJzsKICAgIHRoLmNsYXNzTGlzdC5hZGQoJ3NvcnRhYmxlLXRoJyk7CiAgICBjb25zdCBpc0FjdGl2ZSA9IHNvcnRTdGF0ZS5rZXkgPT09IGtleTsKICAgIHRoLmFwcGVuZENoaWxkKGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKGxhYmVsKSk7CiAgICB0aC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3NwYW4nLCBpc0FjdGl2ZSA/IChzb3J0U3RhdGUuZGlyID09PSAnYXNjJyA/ICcg4oaRJyA6ICcg4oaTJykgOiAnIOKGlScsICdzb3J0LWFycm93JykpOwogICAgdGguYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7CiAgICAgIHNvcnRTdGF0ZSA9IHsga2V5LCBkaXI6IHNvcnRTdGF0ZS5rZXkgPT09IGtleSAmJiBzb3J0U3RhdGUuZGlyID09PSAnYXNjJyA/ICdkZXNjJyA6ICdhc2MnLCB0eXBlIH07CiAgICAgIHJlbmRlckxpc3QoKTsKICAgIH0pOwogICAgcmV0dXJuIHRoOwogIH0KCiAgZnVuY3Rpb24gZmlsdGVyZWRVcGxvYWRzKCkgewogICAgY29uc3QgcSA9IHNlYXJjaFZhbHVlLnRyaW0oKS50b0xvd2VyQ2FzZSgpOwogICAgaWYgKCFxKSByZXR1cm4gY3VycmVudFVwbG9hZHM7CiAgICByZXR1cm4gY3VycmVudFVwbG9hZHMuZmlsdGVyKCh1KSA9PiAoCiAgICAgIHUuZmlsZW5hbWUudG9Mb3dlckNhc2UoKS5pbmNsdWRlcyhxKQogICAgICB8fCAodS5ub3RlcyB8fCAnJykudG9Mb3dlckNhc2UoKS5pbmNsdWRlcyhxKQogICAgICB8fCB1LnN0YXR1cy50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpCiAgICApKTsKICB9CgogIGZ1bmN0aW9uIHNvcnRlZFVwbG9hZHMoKSB7CiAgICBjb25zdCB7IGtleSwgZGlyLCB0eXBlIH0gPSBzb3J0U3RhdGU7CiAgICBjb25zdCBmYWN0b3IgPSBkaXIgPT09ICdhc2MnID8gMSA6IC0xOwogICAgcmV0dXJuIFsuLi5maWx0ZXJlZFVwbG9hZHMoKV0uc29ydCgoYSwgYikgPT4gewogICAgICBjb25zdCBhdiA9IGFba2V5XTsKICAgICAgY29uc3QgYnYgPSBiW2tleV07CiAgICAgIGlmIChhdiA9PT0gbnVsbCB8fCBhdiA9PT0gdW5kZWZpbmVkKSByZXR1cm4gMTsKICAgICAgaWYgKGJ2ID09PSBudWxsIHx8IGJ2ID09PSB1bmRlZmluZWQpIHJldHVybiAtMTsKICAgICAgaWYgKHR5cGUgPT09ICdudW1iZXInKSByZXR1cm4gKGF2IC0gYnYpICogZmFjdG9yOwogICAgICByZXR1cm4gU3RyaW5nKGF2KS5sb2NhbGVDb21wYXJlKFN0cmluZyhidikpICogZmFjdG9yOwogICAgfSk7CiAgfQoKICBmdW5jdGlvbiBleHBvcnRSb3dzKCkgewogICAgcmV0dXJuIHNvcnRlZFVwbG9hZHMoKS5tYXAoKHUpID0+ICh7CiAgICAgIGZpbGVuYW1lOiB1LmZpbGVuYW1lLAogICAgICB1cGxvYWRlZF9hdDogdS51cGxvYWRlZF9hdCwKICAgICAgc3RhdHVzOiB1LnN0YXR1cywKICAgICAgaW1wb3J0ZWRfcm93czogdS5pbXBvcnRlZF9yb3dzLAogICAgICB1cGRhdGVkX3Jvd3M6IHUudXBkYXRlZF9yb3dzLAogICAgICBza2lwcGVkX3Jvd3M6IHUuc2tpcHBlZF9yb3dzLAogICAgICBlcnJvcl9jb3VudDogdS5lcnJvcl9jb3VudCwKICAgICAgd2Vla3M6IHUud2Vla3NfYWZmZWN0ZWQubWFwKCh3KSA9PiBGb3JtYXQuZGF0ZSh3KSkuam9pbignLCAnKSwKICAgICAgbm90ZXM6IHUubm90ZXMgfHwgJycsCiAgICB9KSk7CiAgfQoKICBmdW5jdGlvbiBidWlsZEJhY2t1cENhcmQoKSB7CiAgICBjb25zdCBjYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjYXJkLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGNhcmQuc3R5bGUubWFyZ2luQm90dG9tID0gJzIwcHgnOwogICAgY29uc3QgaGVhZGVyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBoZWFkZXIuY2xhc3NOYW1lID0gJ2NhcmQtaGVhZGVyJzsKICAgIGhlYWRlci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2gzJywgJ0JhY2t1cCAmIFJlc3RvcmUnKSk7CiAgICBjYXJkLmFwcGVuZENoaWxkKGhlYWRlcik7CiAgICBjYXJkLmFwcGVuZENoaWxkKHRleHRFbCgKICAgICAgJ2RpdicsCiAgICAgICdEb3dubG9hZCBhIGZ1bGwgc25hcHNob3Qgb2YgdGhlIGRhdGFiYXNlIGFueSB0aW1lLiBSZXN0b3JpbmcgcmVwbGFjZXMgQUxMIGN1cnJlbnQgZGF0YSB3aXRoIHRoZSB1cGxvYWRlZCBiYWNrdXAgYW5kIHJlc3RhcnRzIHRoZSBzZXJ2ZXIg4oCUIHRoaXMgY2Fubm90IGJlIHVuZG9uZS4nLAogICAgICAnbXV0ZWQnCiAgICApKTsKCiAgICBjb25zdCBhY3Rpb25zID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBhY3Rpb25zLmNsYXNzTmFtZSA9ICdidG4tcm93JzsKICAgIGFjdGlvbnMuc3R5bGUubWFyZ2luVG9wID0gJzE0cHgnOwoKICAgIGNvbnN0IGRvd25sb2FkQnRuID0gaWNvbkJ0bignYnRuIHByaW1hcnknLCAnZG93bmxvYWQnLCAnRG93bmxvYWQgQmFja3VwJyk7CiAgICBkb3dubG9hZEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsgd2luZG93LmxvY2F0aW9uLmhyZWYgPSAnL2FwaS9iYWNrdXAvZXhwb3J0JzsgfSk7CgogICAgY29uc3QgcmVzdG9yZUlucHV0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsKICAgIHJlc3RvcmVJbnB1dC50eXBlID0gJ2ZpbGUnOwogICAgcmVzdG9yZUlucHV0LmFjY2VwdCA9ICcuZGInOwogICAgcmVzdG9yZUlucHV0LnN0eWxlLmRpc3BsYXkgPSAnbm9uZSc7CgogICAgY29uc3QgcmVzdG9yZUJ0biA9IGljb25CdG4oJ2J0biBkYW5nZXInLCAndXBsb2FkJywgJ1Jlc3RvcmUgZnJvbSBCYWNrdXAnKTsKICAgIHJlc3RvcmVCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiByZXN0b3JlSW5wdXQuY2xpY2soKSk7CgogICAgcmVzdG9yZUlucHV0LmFkZEV2ZW50TGlzdGVuZXIoJ2NoYW5nZScsIGFzeW5jICgpID0+IHsKICAgICAgY29uc3QgZmlsZSA9IHJlc3RvcmVJbnB1dC5maWxlc1swXTsKICAgICAgaWYgKCFmaWxlKSByZXR1cm47CiAgICAgIGNvbnN0IHN1cmUgPSB3aW5kb3cuY29uZmlybSgKICAgICAgICAnUmVzdG9yaW5nIHdpbGwgUkVQTEFDRSBhbGwgY3VycmVudCBkYXRhIHdpdGggdGhpcyBiYWNrdXAgZmlsZSBhbmQgcmVzdGFydCB0aGUgc2VydmVyLiBUaGlzIGNhbm5vdCBiZSB1bmRvbmUuIENvbnRpbnVlPycKICAgICAgKTsKICAgICAgaWYgKCFzdXJlKSB7CiAgICAgICAgcmVzdG9yZUlucHV0LnZhbHVlID0gJyc7CiAgICAgICAgcmV0dXJuOwogICAgICB9CiAgICAgIHJlc3RvcmVCdG4uZGlzYWJsZWQgPSB0cnVlOwogICAgICB0cnkgewogICAgICAgIGNvbnN0IGZvcm0gPSBuZXcgRm9ybURhdGEoKTsKICAgICAgICBmb3JtLmFwcGVuZCgnZmlsZScsIGZpbGUpOwogICAgICAgIGNvbnN0IHJlc3VsdCA9IGF3YWl0IEFwaS5yZXN0b3JlQmFja3VwKGZvcm0pOwogICAgICAgIFRvYXN0LnNob3cocmVzdWx0Lm1lc3NhZ2UgfHwgJ0JhY2t1cCByZXN0b3JlZC4gVGhlIHNlcnZlciBpcyByZXN0YXJ0aW5nLicsICdzdWNjZXNzJyk7CiAgICAgIH0gY2F0Y2ggKGVycikgewogICAgICAgIFRvYXN0LnNob3coZXJyLm1lc3NhZ2UsICdlcnJvcicpOwogICAgICAgIHJlc3RvcmVCdG4uZGlzYWJsZWQgPSBmYWxzZTsKICAgICAgfSBmaW5hbGx5IHsKICAgICAgICByZXN0b3JlSW5wdXQudmFsdWUgPSAnJzsKICAgICAgfQogICAgfSk7CgogICAgYWN0aW9ucy5hcHBlbmQoZG93bmxvYWRCdG4sIHJlc3RvcmVCdG4sIHJlc3RvcmVJbnB1dCk7CiAgICBjYXJkLmFwcGVuZENoaWxkKGFjdGlvbnMpOwogICAgcmV0dXJuIGNhcmQ7CiAgfQoKICBhc3luYyBmdW5jdGlvbiByZW5kZXIoKSB7CiAgICByb290ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3ZpZXctaGlzdG9yeScpOwogICAgcm9vdC5pbm5lckhUTUwgPSAnJzsKICAgIHJvb3QuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCAnVXBsb2FkIGhpc3RvcnknLCAnc2VjdGlvbi10aXRsZScpKTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQoYnVpbGRCYWNrdXBDYXJkKCkpOwoKICAgIGN1cnJlbnRVcGxvYWRzID0gYXdhaXQgQXBpLnVwbG9hZEhpc3RvcnkoKTsKICAgIGlmICghY3VycmVudFVwbG9hZHMubGVuZ3RoKSB7CiAgICAgIHJvb3QuYXBwZW5kQ2hpbGQoZW1wdHlTdGF0ZSh7CiAgICAgICAgaWNvbjogJ3VwbG9hZC1jbG91ZCcsCiAgICAgICAgdGl0bGU6ICdObyB1cGxvYWRzIHlldCcsCiAgICAgICAgbWVzc2FnZTogJ0ltcG9ydCB5b3VyIGZpcnN0IHdlZWtseSBleHBvcnQgdG8gc3RhcnQgc2VlaW5nIGRhdGEgYWNyb3NzIHRoZSBhcHAuJywKICAgICAgICBhY3Rpb25MYWJlbDogJ1VwbG9hZCBkYXRhJywKICAgICAgICBvbkFjdGlvbjogKCkgPT4gZG9jdW1lbnQucXVlcnlTZWxlY3RvcignLnRhYi1idG5bZGF0YS10YWI9InVwbG9hZCJdJyk/LmNsaWNrKCksCiAgICAgIH0pKTsKICAgICAgcmV0dXJuOwogICAgfQoKICAgIGNvbnN0IHRvb2xiYXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRvb2xiYXIuY2xhc3NOYW1lID0gJ3JlY29yZHMtdG9vbGJhcic7CiAgICB0b29sYmFyLmFwcGVuZENoaWxkKGJ1aWxkU2VhcmNoQm94KHsKICAgICAgcGxhY2Vob2xkZXI6ICdTZWFyY2ggZmlsZW5hbWUsIHN0YXR1cywgb3Igbm90ZXPigKYnLAogICAgICB2YWx1ZTogc2VhcmNoVmFsdWUsCiAgICAgIG9uQ2hhbmdlOiAodikgPT4geyBzZWFyY2hWYWx1ZSA9IHY7IHBhZ2UgPSAxOyByZW5kZXJMaXN0KCk7IH0sCiAgICB9KSk7CiAgICByb290LmFwcGVuZENoaWxkKHRvb2xiYXIpOwoKICAgIHJvb3QuYXBwZW5kQ2hpbGQoYnVpbGRFeHBvcnRCdXR0b25zKHsKICAgICAgZ2V0Um93c0FuZENvbHVtbnM6ICgpID0+ICh7IHJvd3M6IGV4cG9ydFJvd3MoKSwgY29sdW1uczogRVhQT1JUX0NPTFVNTlMgfSksCiAgICAgIGZpbGVuYW1lQmFzZTogJ3VwbG9hZC1oaXN0b3J5JywKICAgICAgc2hlZXROYW1lOiAnVXBsb2FkIEhpc3RvcnknLAogICAgfSkpOwoKICAgIGNvbnN0IGNhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgd3JhcC5jbGFzc05hbWUgPSAndGFibGUtc2Nyb2xsJzsKICAgIHdyYXAuaWQgPSAnaGlzdG9yeVRhYmxlV3JhcCc7CiAgICBjYXJkLmFwcGVuZENoaWxkKHdyYXApOwogICAgY29uc3QgcGFnZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHBhZ2VyLmNsYXNzTmFtZSA9ICdwYWdpbmF0aW9uLXJvdyc7CiAgICBwYWdlci5pZCA9ICdoaXN0b3J5UGFnZXInOwogICAgY2FyZC5hcHBlbmRDaGlsZChwYWdlcik7CiAgICByb290LmFwcGVuZENoaWxkKGNhcmQpOwoKICAgIHJlbmRlckxpc3QoKTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlckxpc3QoKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2hpc3RvcnlUYWJsZVdyYXAnKTsKICAgIGNvbnN0IHBhZ2VyRWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnaGlzdG9yeVBhZ2VyJyk7CiAgICBpZiAoIXdyYXApIHJldHVybjsKICAgIGNvbnN0IGFsbFNvcnRlZCA9IHNvcnRlZFVwbG9hZHMoKTsKICAgIGlmICghYWxsU29ydGVkLmxlbmd0aCkgewogICAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgICB3cmFwLmFwcGVuZENoaWxkKGVtcHR5U3RhdGUoeyBpY29uOiAndXBsb2FkLWNsb3VkJywgbWVzc2FnZTogJ05vIHVwbG9hZHMgbWF0Y2ggeW91ciBzZWFyY2guJyB9KSk7CiAgICAgIGlmIChwYWdlckVsKSBwYWdlckVsLmlubmVySFRNTCA9ICcnOwogICAgICByZXR1cm47CiAgICB9CiAgICBjb25zdCB7IHBhZ2VSb3dzLCB0b3RhbFBhZ2VzLCBzYWZlUGFnZSwgdG90YWwgfSA9IHBhZ2luYXRlQ2xpZW50U2lkZShhbGxTb3J0ZWQsIHBhZ2UsIHBhZ2VTaXplKTsKICAgIHBhZ2UgPSBzYWZlUGFnZTsKCiAgICBjb25zdCB0YWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICB0YWJsZS5jbGFzc05hbWUgPSAnZGF0YS10YWJsZSc7CiAgICBjb25zdCB0aGVhZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RoZWFkJyk7CiAgICBjb25zdCBoZWFkVHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgaGVhZFRyLmFwcGVuZCgKICAgICAgc29ydGFibGVIZWFkZXIoJ0ZpbGUnLCAnZmlsZW5hbWUnLCAnc3RyaW5nJyksCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdVcGxvYWRlZCcsICd1cGxvYWRlZF9hdCcsICdzdHJpbmcnKSwKICAgICAgc29ydGFibGVIZWFkZXIoJ1N0YXR1cycsICdzdGF0dXMnLCAnc3RyaW5nJyksCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdJbXBvcnRlZCcsICdpbXBvcnRlZF9yb3dzJywgJ251bWJlcicpLAogICAgICBzb3J0YWJsZUhlYWRlcignVXBkYXRlZCcsICd1cGRhdGVkX3Jvd3MnLCAnbnVtYmVyJyksCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdTa2lwcGVkJywgJ3NraXBwZWRfcm93cycsICdudW1iZXInKSwKICAgICAgc29ydGFibGVIZWFkZXIoJ0Vycm9ycycsICdlcnJvcl9jb3VudCcsICdudW1iZXInKSwKICAgICAgdGV4dEVsKCd0aCcsICdXZWVrcycpLAogICAgICB0ZXh0RWwoJ3RoJywgJ05vdGVzJykKICAgICk7CiAgICB0aGVhZC5hcHBlbmRDaGlsZChoZWFkVHIpOwogICAgY29uc3QgdGJvZHkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0Ym9keScpOwogICAgcGFnZVJvd3MuZm9yRWFjaCgodSkgPT4gewogICAgICBjb25zdCB0ciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAgIHRyLnN0eWxlLmN1cnNvciA9ICdwb2ludGVyJzsKICAgICAgY29uc3QgYmFkZ2UgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICAgIGJhZGdlLmNsYXNzTmFtZSA9IGBiYWRnZSAke2JhZGdlQ2xhc3ModS5zdGF0dXMpfWA7CiAgICAgIGJhZGdlLnRleHRDb250ZW50ID0gdS5zdGF0dXM7CiAgICAgIGNvbnN0IHN0YXR1c1RkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgc3RhdHVzVGQuYXBwZW5kQ2hpbGQoYmFkZ2UpOwogICAgICB0ci5hcHBlbmQoCiAgICAgICAgdGV4dEVsKCd0ZCcsIHUuZmlsZW5hbWUpLAogICAgICAgIHRleHRFbCgndGQnLCB1LnVwbG9hZGVkX2F0KSwKICAgICAgICBzdGF0dXNUZCwKICAgICAgICB0ZXh0RWwoJ3RkJywgU3RyaW5nKHUuaW1wb3J0ZWRfcm93cyksICdudW0nKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgU3RyaW5nKHUudXBkYXRlZF9yb3dzKSwgJ251bScpLAogICAgICAgIHRleHRFbCgndGQnLCBTdHJpbmcodS5za2lwcGVkX3Jvd3MpLCAnbnVtJyksCiAgICAgICAgdGV4dEVsKCd0ZCcsIFN0cmluZyh1LmVycm9yX2NvdW50KSwgJ251bScpLAogICAgICAgIHRleHRFbCgndGQnLCB1LndlZWtzX2FmZmVjdGVkLm1hcCgodykgPT4gRm9ybWF0LmRhdGUodykpLmpvaW4oJywgJykgfHwgJ+KAlCcpLAogICAgICAgIHRleHRFbCgndGQnLCB1Lm5vdGVzIHx8ICfigJQnKQogICAgICApOwogICAgICB0ci5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHRvZ2dsZUVycm9ycyh1LmlkLCB0cikpOwogICAgICB0Ym9keS5hcHBlbmRDaGlsZCh0cik7CiAgICB9KTsKICAgIHRhYmxlLmFwcGVuZCh0aGVhZCwgdGJvZHkpOwogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGFibGUpOwoKICAgIGlmIChwYWdlckVsKSB7CiAgICAgIHBhZ2VyRWwuaW5uZXJIVE1MID0gJyc7CiAgICAgIHBhZ2VyRWwuYXBwZW5kQ2hpbGQoYnVpbGRQYWdlcih7CiAgICAgICAgcGFnZTogc2FmZVBhZ2UsCiAgICAgICAgdG90YWxQYWdlcywKICAgICAgICB0b3RhbCwKICAgICAgICBvblByZXY6ICgpID0+IHsgcGFnZSAtPSAxOyByZW5kZXJMaXN0KCk7IH0sCiAgICAgICAgb25OZXh0OiAoKSA9PiB7IHBhZ2UgKz0gMTsgcmVuZGVyTGlzdCgpOyB9LAogICAgICB9KSk7CiAgICB9CiAgfQoKICBhc3luYyBmdW5jdGlvbiB0b2dnbGVFcnJvcnModXBsb2FkSWQsIHRyKSB7CiAgICBjb25zdCBleGlzdGluZyA9IHRyLm5leHRFbGVtZW50U2libGluZzsKICAgIGlmIChleGlzdGluZyAmJiBleGlzdGluZy5jbGFzc0xpc3QuY29udGFpbnMoJ2Vycm9yLWxvZy1yb3cnKSkgewogICAgICBleGlzdGluZy5yZW1vdmUoKTsKICAgICAgcmV0dXJuOwogICAgfQogICAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnLmVycm9yLWxvZy1yb3cnKS5mb3JFYWNoKChlbCkgPT4gZWwucmVtb3ZlKCkpOwogICAgY29uc3QgZXJyb3JzID0gYXdhaXQgQXBpLnVwbG9hZEVycm9ycyh1cGxvYWRJZCk7CiAgICBjb25zdCBsb2dSb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgbG9nUm93LmNsYXNzTmFtZSA9ICdlcnJvci1sb2ctcm93JzsKICAgIGNvbnN0IHRkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgIHRkLmNvbFNwYW4gPSA5OwogICAgaWYgKCFlcnJvcnMubGVuZ3RoKSB7CiAgICAgIHRkLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgJ05vIGlzc3VlcyBsb2dnZWQgZm9yIHRoaXMgdXBsb2FkLicsICdtdXRlZCcpKTsKICAgIH0gZWxzZSB7CiAgICAgIGNvbnN0IGxpc3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgbGlzdC5jbGFzc05hbWUgPSAnaXNzdWVzLWxpc3QnOwogICAgICBlcnJvcnMuZm9yRWFjaCgoZSkgPT4gewogICAgICAgIGNvbnN0IHJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICAgIHJvdy5jbGFzc05hbWUgPSAnaXNzdWUtcm93JzsKICAgICAgICBjb25zdCBiYWRnZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgICAgICBiYWRnZS5jbGFzc05hbWUgPSBgYmFkZ2UgJHtlLnNldmVyaXR5fS1zZXZgOwogICAgICAgIGJhZGdlLnRleHRDb250ZW50ID0gZS5zZXZlcml0eTsKICAgICAgICByb3cuYXBwZW5kKGJhZGdlLCBkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShgICR7ZS5yb3dfbnVtYmVyID8gYFJvdyAke2Uucm93X251bWJlcn06IGAgOiAnJ30ke2UubWVzc2FnZX1gKSk7CiAgICAgICAgbGlzdC5hcHBlbmRDaGlsZChyb3cpOwogICAgICB9KTsKICAgICAgdGQuYXBwZW5kQ2hpbGQobGlzdCk7CiAgICB9CgogICAgY29uc3QgcmF3QnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICByYXdCdG4uY2xhc3NOYW1lID0gJ2J0bic7CiAgICByYXdCdG4uc3R5bGUubWFyZ2luVG9wID0gJzEwcHgnOwogICAgcmF3QnRuLnRleHRDb250ZW50ID0gJ1ZpZXcgZXZlcnkgcmF3IHNvdXJjZSByb3cgZnJvbSB0aGlzIHVwbG9hZCc7CiAgICByYXdCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBsb2FkUmF3Um93cyh1cGxvYWRJZCwgcmF3QnRuKSk7CiAgICB0ZC5hcHBlbmRDaGlsZChyYXdCdG4pOwogICAgY29uc3QgcmF3V3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcmF3V3JhcC5pZCA9IGByYXdXcmFwLSR7dXBsb2FkSWR9YDsKICAgIHRkLmFwcGVuZENoaWxkKHJhd1dyYXApOwoKICAgIGxvZ1Jvdy5hcHBlbmRDaGlsZCh0ZCk7CiAgICB0ci5hZnRlcihsb2dSb3cpOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gbG9hZFJhd1Jvd3ModXBsb2FkSWQsIGJ0bikgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGByYXdXcmFwLSR7dXBsb2FkSWR9YCk7CiAgICBpZiAoIXdyYXApIHJldHVybjsKICAgIGlmICh3cmFwLmRhdGFzZXQubG9hZGVkKSB7CiAgICAgIHdyYXAuc3R5bGUuZGlzcGxheSA9IHdyYXAuc3R5bGUuZGlzcGxheSA9PT0gJ25vbmUnID8gJ2Jsb2NrJyA6ICdub25lJzsKICAgICAgcmV0dXJuOwogICAgfQogICAgYnRuLnRleHRDb250ZW50ID0gJ0xvYWRpbmfigKYnOwogICAgY29uc3QgeyByb3dzLCB0b3RhbCB9ID0gYXdhaXQgQXBpLnVwbG9hZFJhd1Jvd3ModXBsb2FkSWQpOwogICAgd3JhcC5kYXRhc2V0LmxvYWRlZCA9ICcxJzsKICAgIGJ0bi50ZXh0Q29udGVudCA9IGBTaG93aW5nICR7cm93cy5sZW5ndGh9IG9mICR7dG90YWx9IHJhdyByb3cocylgOwoKICAgIGNvbnN0IGJ5U2hlZXQgPSBuZXcgTWFwKCk7CiAgICByb3dzLmZvckVhY2goKHIpID0+IHsKICAgICAgaWYgKCFieVNoZWV0LmhhcyhyLnNoZWV0X25hbWUpKSBieVNoZWV0LnNldChyLnNoZWV0X25hbWUsIFtdKTsKICAgICAgYnlTaGVldC5nZXQoci5zaGVldF9uYW1lKS5wdXNoKHIpOwogICAgfSk7CgogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgIHdyYXAuc3R5bGUubWFyZ2luVG9wID0gJzEwcHgnOwogICAgYnlTaGVldC5mb3JFYWNoKChzaGVldFJvd3MsIHNoZWV0TmFtZSkgPT4gewogICAgICB3cmFwLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgYFNoZWV0OiAke3NoZWV0TmFtZX0gKCR7c2hlZXRSb3dzLmxlbmd0aH0gcm93KHMpKWAsICdzdGF0LWxhYmVsJykpOwogICAgICBjb25zdCBoZWFkZXJzID0gc2hlZXRSb3dzWzBdLmhlYWRlcnM7CiAgICAgIGNvbnN0IHRhYmxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGFibGUnKTsKICAgICAgdGFibGUuY2xhc3NOYW1lID0gJ2RhdGEtdGFibGUnOwogICAgICBjb25zdCB0aGVhZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAgIHRoZWFkLmFwcGVuZCh0ZXh0RWwoJ3RoJywgJ1JvdyAjJyksIHRleHRFbCgndGgnLCAnTGlua2VkIHRvIHBvc3QnKSk7CiAgICAgIGNvbnN0IGNvbENvdW50ID0gaGVhZGVycyA/IGhlYWRlcnMubGVuZ3RoIDogTWF0aC5tYXgoLi4uc2hlZXRSb3dzLm1hcCgocikgPT4gci5yYXcubGVuZ3RoKSk7CiAgICAgIGZvciAobGV0IGkgPSAwOyBpIDwgY29sQ291bnQ7IGkgKz0gMSkgdGhlYWQuYXBwZW5kQ2hpbGQodGV4dEVsKCd0aCcsIGhlYWRlcnMgJiYgaGVhZGVyc1tpXSA/IFN0cmluZyhoZWFkZXJzW2ldKSA6IGBDb2wgJHtpICsgMX1gKSk7CiAgICAgIGNvbnN0IHRoZWFkV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RoZWFkJyk7CiAgICAgIHRoZWFkV3JhcC5hcHBlbmRDaGlsZCh0aGVhZCk7CiAgICAgIGNvbnN0IHRib2R5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGJvZHknKTsKICAgICAgc2hlZXRSb3dzLmZvckVhY2goKHIpID0+IHsKICAgICAgICBjb25zdCB0cjIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgICAgIHRyMi5hcHBlbmQodGV4dEVsKCd0ZCcsIFN0cmluZyhyLnJvd19udW1iZXIpKSwgdGV4dEVsKCd0ZCcsIHIucG9zdF9pZCA/IGAjJHtyLnBvc3RfaWR9YCA6ICfigJQnKSk7CiAgICAgICAgZm9yIChsZXQgaSA9IDA7IGkgPCBjb2xDb3VudDsgaSArPSAxKSB7CiAgICAgICAgICBjb25zdCB2YWwgPSByLnJhd1tpXTsKICAgICAgICAgIHRyMi5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3RkJywgdmFsID09PSB1bmRlZmluZWQgfHwgdmFsID09PSBudWxsID8gJycgOiBTdHJpbmcodmFsKS5zbGljZSgwLCA2MCkpKTsKICAgICAgICB9CiAgICAgICAgdGJvZHkuYXBwZW5kQ2hpbGQodHIyKTsKICAgICAgfSk7CiAgICAgIHRhYmxlLmFwcGVuZCh0aGVhZFdyYXAsIHRib2R5KTsKICAgICAgY29uc3Qgc2Nyb2xsV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBzY3JvbGxXcmFwLmNsYXNzTmFtZSA9ICd0YWJsZS1zY3JvbGwnOwogICAgICBzY3JvbGxXcmFwLnN0eWxlLm1hcmdpbkJvdHRvbSA9ICcxNnB4JzsKICAgICAgc2Nyb2xsV3JhcC5hcHBlbmRDaGlsZCh0YWJsZSk7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQoc2Nyb2xsV3JhcCk7CiAgICB9KTsKICB9CgogIHJldHVybiB7IHJlbmRlciB9Owp9KSgpOwoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIEZvbGxvd2VycyBEYXRhIHRhYjogbWFudWFsIHdlZWtseSBmb2xsb3dlci1jb3VudCBlbnRyeSBwZXIKICAgcGxhdGZvcm0g4oCUIGVudGlyZWx5IGluZGVwZW5kZW50IG9mIHNwcmVhZHNoZWV0IHVwbG9hZHMgKGl0cyBvd24KICAgdGFibGUsIGl0cyBvd24gQVBJLCBuZXZlciB0b3VjaGVkIGJ5IHRoZSBpbXBvcnQgcGlwZWxpbmUpLiBQb3dlcnMKICAgRm9sbG93ZXIgR3Jvd3RoIGNoYXJ0cy9jb21wYXJpc29ucyBlbHNld2hlcmUgaW4gdGhlIGFwcC4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmNvbnN0IEZvbGxvd2VycyA9ICgoKSA9PiB7CiAgbGV0IHJvb3Q7CiAgbGV0IGVkaXRpbmdJZCA9IG51bGw7IC8vIG5vbi1udWxsIHdoaWxlIHRoZSBmb3JtIGlzIGVkaXRpbmcgYW4gZXhpc3RpbmcgZW50cnkgcmF0aGVyIHRoYW4gYWRkaW5nIGEgbmV3IG9uZQogIGxldCBzb3J0U3RhdGUgPSB7IGtleTogJ2VudHJ5X2RhdGUnLCBkaXI6ICdkZXNjJywgdHlwZTogJ3N0cmluZycgfTsKICBsZXQgY3VycmVudFJvd3MgPSBbXTsKICBsZXQgc2VhcmNoVmFsdWUgPSAnJzsKICBsZXQgcGFnZSA9IDE7CiAgY29uc3QgcGFnZVNpemUgPSAxMDsKICBjb25zdCBFWFBPUlRfQ09MVU1OUyA9IFsKICAgIHsga2V5OiAncGxhdGZvcm1fbGFiZWwnLCBsYWJlbDogJ1BsYXRmb3JtJyB9LAogICAgeyBrZXk6ICdlbnRyeV9kYXRlJywgbGFiZWw6ICdXZWVrIC8gRGF0ZScgfSwKICAgIHsga2V5OiAnZm9sbG93ZXJzX2NvdW50JywgbGFiZWw6ICdGb2xsb3dlcnMgQ291bnQnIH0sCiAgICB7IGtleTogJ3VwZGF0ZWRfYXQnLCBsYWJlbDogJ0xhc3QgVXBkYXRlZCcgfSwKICBdOwoKICBmdW5jdGlvbiBhbGxQbGF0Zm9ybXMoKSB7CiAgICByZXR1cm4gKHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZSB8fCB7IGFsbFBsYXRmb3JtczogW10gfSkuYWxsUGxhdGZvcm1zIHx8IFtdOwogIH0KCiAgZnVuY3Rpb24gcGxhdGZvcm1NZXRhRm9yKGlkKSB7CiAgICByZXR1cm4gYWxsUGxhdGZvcm1zKCkuZmluZCgocCkgPT4gcC5pZCA9PT0gaWQpIHx8IHsgbGFiZWw6IGlkLCBjb2xvcjogJyM5OTknIH07CiAgfQoKICBmdW5jdGlvbiBzaGVsbCgpIHsKICAgIHJvb3QuaW5uZXJIVE1MID0gJyc7CiAgICByb290LmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgJ0ZvbGxvd2VycyBEYXRhIFJlY29yZCcsICdzZWN0aW9uLXRpdGxlJykpOwogICAgcm9vdC5hcHBlbmRDaGlsZCh0ZXh0RWwoCiAgICAgICdkaXYnLAogICAgICAnTWFudWFsbHkgbG9nIGVhY2ggcGxhdGZvcm3igJlzIHRvdGFsIGZvbGxvd2VyIGNvdW50IG9uY2UgYSB3ZWVrLiBUaGlzIGlzIGluZGVwZW5kZW50IG9mIHNwcmVhZHNoZWV0IHVwbG9hZHMg4oCUIGl0IHBvd2VycyBGb2xsb3dlciBHcm93dGggY2hhcnRzIGFuZCBjb21wYXJpc29ucyBlbHNld2hlcmUgaW4gdGhlIGFwcC4nLAogICAgICAnbXV0ZWQnCiAgICApKTsKCiAgICBjb25zdCBmb3JtQ2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZm9ybUNhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgZm9ybUNhcmQuc3R5bGUubWFyZ2luQm90dG9tID0gJzIwcHgnOwogICAgZm9ybUNhcmQuaWQgPSAnZm9sbG93ZXJzRm9ybUNhcmQnOwogICAgcm9vdC5hcHBlbmRDaGlsZChmb3JtQ2FyZCk7CgogICAgY29uc3QgdG9vbGJhciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdG9vbGJhci5jbGFzc05hbWUgPSAncmVjb3Jkcy10b29sYmFyJzsKICAgIHRvb2xiYXIuYXBwZW5kQ2hpbGQoYnVpbGRTZWFyY2hCb3goewogICAgICBwbGFjZWhvbGRlcjogJ1NlYXJjaCBwbGF0Zm9ybSBvciBkYXRl4oCmJywKICAgICAgdmFsdWU6IHNlYXJjaFZhbHVlLAogICAgICBvbkNoYW5nZTogKHYpID0+IHsgc2VhcmNoVmFsdWUgPSB2OyBwYWdlID0gMTsgcmVuZGVyVGFibGUoKTsgfSwKICAgIH0pKTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQodG9vbGJhcik7CgogICAgcm9vdC5hcHBlbmRDaGlsZChidWlsZEV4cG9ydEJ1dHRvbnMoewogICAgICBnZXRSb3dzQW5kQ29sdW1uczogKCkgPT4gKHsgcm93czogZXhwb3J0Um93cygpLCBjb2x1bW5zOiBFWFBPUlRfQ09MVU1OUyB9KSwKICAgICAgZmlsZW5hbWVCYXNlOiAnZm9sbG93ZXJzLWRhdGEnLAogICAgICBzaGVldE5hbWU6ICdGb2xsb3dlcnMgRGF0YScsCiAgICB9KSk7CgogICAgY29uc3QgdGFibGVDYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0YWJsZUNhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgY29uc3QgdGFibGVXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0YWJsZVdyYXAuY2xhc3NOYW1lID0gJ3RhYmxlLXNjcm9sbCc7CiAgICB0YWJsZVdyYXAuaWQgPSAnZm9sbG93ZXJzVGFibGVXcmFwJzsKICAgIHRhYmxlQ2FyZC5hcHBlbmRDaGlsZCh0YWJsZVdyYXApOwogICAgY29uc3QgcGFnZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHBhZ2VyLmNsYXNzTmFtZSA9ICdwYWdpbmF0aW9uLXJvdyc7CiAgICBwYWdlci5pZCA9ICdmb2xsb3dlcnNQYWdlcic7CiAgICB0YWJsZUNhcmQuYXBwZW5kQ2hpbGQocGFnZXIpOwogICAgcm9vdC5hcHBlbmRDaGlsZCh0YWJsZUNhcmQpOwoKICAgIHJlbmRlckZvcm0oKTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlckZvcm0oKSB7CiAgICBjb25zdCBjYXJkID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZvbGxvd2Vyc0Zvcm1DYXJkJyk7CiAgICBpZiAoIWNhcmQpIHJldHVybjsKICAgIGNhcmQuaW5uZXJIVE1MID0gJyc7CiAgICBjb25zdCBoZWFkZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGhlYWRlci5jbGFzc05hbWUgPSAnY2FyZC1oZWFkZXInOwogICAgaGVhZGVyLmFwcGVuZENoaWxkKHRleHRFbCgnaDMnLCBlZGl0aW5nSWQgIT09IG51bGwgPyAnRWRpdCBlbnRyeScgOiAnQWRkIGEgd2Vla2x5IGVudHJ5JykpOwogICAgY2FyZC5hcHBlbmRDaGlsZChoZWFkZXIpOwoKICAgIGNvbnN0IGdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGdyaWQuY2xhc3NOYW1lID0gJ2Zvcm0tZ3JpZCc7CgogICAgY29uc3QgcGxhdGZvcm1GaWVsZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcGxhdGZvcm1GaWVsZC5jbGFzc05hbWUgPSAnZm9ybS1maWVsZCc7CiAgICBwbGF0Zm9ybUZpZWxkLmFwcGVuZENoaWxkKHRleHRFbCgnbGFiZWwnLCAnUGxhdGZvcm0nKSk7CiAgICBjb25zdCBwbGF0Zm9ybVNlbGVjdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NlbGVjdCcpOwogICAgcGxhdGZvcm1TZWxlY3QuaWQgPSAnZm9sbG93ZXJzUGxhdGZvcm1JbnB1dCc7CiAgICBhbGxQbGF0Zm9ybXMoKS5mb3JFYWNoKChwKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOwogICAgICBvcHQudmFsdWUgPSBwLmlkOwogICAgICBvcHQudGV4dENvbnRlbnQgPSBwLmxhYmVsOwogICAgICBwbGF0Zm9ybVNlbGVjdC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgICBwbGF0Zm9ybUZpZWxkLmFwcGVuZENoaWxkKHBsYXRmb3JtU2VsZWN0KTsKCiAgICBjb25zdCBkYXRlRmllbGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGRhdGVGaWVsZC5jbGFzc05hbWUgPSAnZm9ybS1maWVsZCc7CiAgICBkYXRlRmllbGQuYXBwZW5kQ2hpbGQodGV4dEVsKCdsYWJlbCcsICdXZWVrIC8gRGF0ZScpKTsKICAgIGNvbnN0IGRhdGVJbnB1dCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7CiAgICBkYXRlSW5wdXQudHlwZSA9ICdkYXRlJzsKICAgIGRhdGVJbnB1dC5pZCA9ICdmb2xsb3dlcnNEYXRlSW5wdXQnOwogICAgZGF0ZUZpZWxkLmFwcGVuZENoaWxkKGRhdGVJbnB1dCk7CgogICAgY29uc3QgY291bnRGaWVsZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY291bnRGaWVsZC5jbGFzc05hbWUgPSAnZm9ybS1maWVsZCc7CiAgICBjb3VudEZpZWxkLmFwcGVuZENoaWxkKHRleHRFbCgnbGFiZWwnLCAnRm9sbG93ZXJzIENvdW50JykpOwogICAgY29uc3QgY291bnRJbnB1dCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7CiAgICBjb3VudElucHV0LnR5cGUgPSAnbnVtYmVyJzsKICAgIGNvdW50SW5wdXQubWluID0gJzAnOwogICAgY291bnRJbnB1dC5zdGVwID0gJzEnOwogICAgY291bnRJbnB1dC5pZCA9ICdmb2xsb3dlcnNDb3VudElucHV0JzsKICAgIGNvdW50RmllbGQuYXBwZW5kQ2hpbGQoY291bnRJbnB1dCk7CgogICAgZ3JpZC5hcHBlbmQocGxhdGZvcm1GaWVsZCwgZGF0ZUZpZWxkLCBjb3VudEZpZWxkKTsKICAgIGNhcmQuYXBwZW5kQ2hpbGQoZ3JpZCk7CgogICAgY29uc3QgZWRpdFJvdyA9IGVkaXRpbmdJZCAhPT0gbnVsbCA/IGN1cnJlbnRSb3dzLmZpbmQoKHIpID0+IHIuaWQgPT09IGVkaXRpbmdJZCkgOiBudWxsOwogICAgaWYgKGVkaXRSb3cpIHsKICAgICAgcGxhdGZvcm1TZWxlY3QudmFsdWUgPSBlZGl0Um93LnBsYXRmb3JtOwogICAgICBkYXRlSW5wdXQudmFsdWUgPSBlZGl0Um93LmVudHJ5X2RhdGU7CiAgICAgIGNvdW50SW5wdXQudmFsdWUgPSBTdHJpbmcoZWRpdFJvdy5mb2xsb3dlcnNfY291bnQpOwogICAgfSBlbHNlIHsKICAgICAgZGF0ZUlucHV0LnZhbHVlID0gbmV3IERhdGUoKS50b0lTT1N0cmluZygpLnNsaWNlKDAsIDEwKTsKICAgIH0KCiAgICBjb25zdCBhY3Rpb25zID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBhY3Rpb25zLmNsYXNzTmFtZSA9ICdtb2RhbC1hY3Rpb25zJzsKICAgIGNvbnN0IGVycm9yRWwgPSB0ZXh0RWwoJ3NwYW4nLCAnJywgJ211dGVkJyk7CiAgICBlcnJvckVsLmlkID0gJ2ZvbGxvd2Vyc0Zvcm1FcnJvcic7CiAgICBlcnJvckVsLnN0eWxlLmNvbG9yID0gJ3ZhcigtLXN0YXR1cy1jcml0aWNhbCknOwoKICAgIGNvbnN0IGJ0blJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYnRuUm93LmNsYXNzTmFtZSA9ICdidG4tcm93JzsKICAgIGNvbnN0IHNhdmVCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIHNhdmVCdG4uY2xhc3NOYW1lID0gJ2J0biBwcmltYXJ5JzsKICAgIHNhdmVCdG4udGV4dENvbnRlbnQgPSBlZGl0aW5nSWQgIT09IG51bGwgPyAnU2F2ZSBjaGFuZ2VzJyA6ICdBZGQgZW50cnknOwogICAgc2F2ZUJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHN1Ym1pdEZvcm0oc2F2ZUJ0bikpOwogICAgYnRuUm93LmFwcGVuZENoaWxkKHNhdmVCdG4pOwogICAgaWYgKGVkaXRpbmdJZCAhPT0gbnVsbCkgewogICAgICBjb25zdCBjYW5jZWxCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgICAgY2FuY2VsQnRuLmNsYXNzTmFtZSA9ICdidG4nOwogICAgICBjYW5jZWxCdG4udGV4dENvbnRlbnQgPSAnQ2FuY2VsJzsKICAgICAgY2FuY2VsQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4geyBlZGl0aW5nSWQgPSBudWxsOyByZW5kZXJGb3JtKCk7IH0pOwogICAgICBidG5Sb3cuYXBwZW5kQ2hpbGQoY2FuY2VsQnRuKTsKICAgIH0KICAgIGFjdGlvbnMuYXBwZW5kKGVycm9yRWwsIGJ0blJvdyk7CiAgICBjYXJkLmFwcGVuZENoaWxkKGFjdGlvbnMpOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gc3VibWl0Rm9ybShidG4pIHsKICAgIGNvbnN0IGVycm9yRWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZm9sbG93ZXJzRm9ybUVycm9yJyk7CiAgICBlcnJvckVsLnRleHRDb250ZW50ID0gJyc7CiAgICBjb25zdCBwbGF0Zm9ybSA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmb2xsb3dlcnNQbGF0Zm9ybUlucHV0JykudmFsdWU7CiAgICBjb25zdCBlbnRyeURhdGUgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZm9sbG93ZXJzRGF0ZUlucHV0JykudmFsdWU7CiAgICBjb25zdCBmb2xsb3dlcnNDb3VudCA9IE51bWJlcihkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZm9sbG93ZXJzQ291bnRJbnB1dCcpLnZhbHVlKTsKICAgIGJ0bi5kaXNhYmxlZCA9IHRydWU7CiAgICB0cnkgewogICAgICBpZiAoZWRpdGluZ0lkICE9PSBudWxsKSB7CiAgICAgICAgYXdhaXQgQXBpLnVwZGF0ZUZvbGxvd2VycyhlZGl0aW5nSWQsIHsgcGxhdGZvcm0sIGVudHJ5RGF0ZSwgZm9sbG93ZXJzQ291bnQgfSk7CiAgICAgICAgVG9hc3Quc2hvdygnRW50cnkgdXBkYXRlZC4nLCAnc3VjY2VzcycpOwogICAgICB9IGVsc2UgewogICAgICAgIGF3YWl0IEFwaS5zYXZlRm9sbG93ZXJzKHsgcGxhdGZvcm0sIGVudHJ5RGF0ZSwgZm9sbG93ZXJzQ291bnQgfSk7CiAgICAgICAgVG9hc3Quc2hvdygnRW50cnkgc2F2ZWQuJywgJ3N1Y2Nlc3MnKTsKICAgICAgfQogICAgICBlZGl0aW5nSWQgPSBudWxsOwogICAgICBhd2FpdCBsb2FkKCk7CiAgICAgIHdpbmRvdy5kaXNwYXRjaEV2ZW50KG5ldyBDdXN0b21FdmVudCgnbHJzOmRhdGEtdXBkYXRlZCcpKTsKICAgIH0gY2F0Y2ggKGVycikgewogICAgICBlcnJvckVsLnRleHRDb250ZW50ID0gZXJyLm1lc3NhZ2U7CiAgICAgIGJ0bi5kaXNhYmxlZCA9IGZhbHNlOwogICAgfQogIH0KCiAgZnVuY3Rpb24gc3RhcnRFZGl0KHJvdykgewogICAgZWRpdGluZ0lkID0gcm93LmlkOwogICAgcmVuZGVyRm9ybSgpOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZvbGxvd2Vyc0Zvcm1DYXJkJykuc2Nyb2xsSW50b1ZpZXcoeyBiZWhhdmlvcjogJ3Ntb290aCcsIGJsb2NrOiAnc3RhcnQnIH0pOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gaGFuZGxlRGVsZXRlKHJvdykgewogICAgY29uc3Qgc3VyZSA9IHdpbmRvdy5jb25maXJtKGBEZWxldGUgdGhlICR7cGxhdGZvcm1NZXRhRm9yKHJvdy5wbGF0Zm9ybSkubGFiZWx9IGVudHJ5IGZvciAke0Zvcm1hdC5kYXRlKHJvdy5lbnRyeV9kYXRlKX0/YCk7CiAgICBpZiAoIXN1cmUpIHJldHVybjsKICAgIHRyeSB7CiAgICAgIGF3YWl0IEFwaS5kZWxldGVGb2xsb3dlcnMocm93LmlkKTsKICAgICAgVG9hc3Quc2hvdygnRW50cnkgZGVsZXRlZC4nLCAnc3VjY2VzcycpOwogICAgICBpZiAoZWRpdGluZ0lkID09PSByb3cuaWQpIGVkaXRpbmdJZCA9IG51bGw7CiAgICAgIGF3YWl0IGxvYWQoKTsKICAgICAgd2luZG93LmRpc3BhdGNoRXZlbnQobmV3IEN1c3RvbUV2ZW50KCdscnM6ZGF0YS11cGRhdGVkJykpOwogICAgfSBjYXRjaCAoZXJyKSB7CiAgICAgIFRvYXN0LnNob3coZXJyLm1lc3NhZ2UsICdlcnJvcicpOwogICAgfQogIH0KCiAgZnVuY3Rpb24gc29ydGFibGVIZWFkZXIobGFiZWwsIGtleSwgdHlwZSkgewogICAgY29uc3QgdGggPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0aCcpOwogICAgaWYgKHR5cGUgPT09ICdudW1iZXInKSB0aC5jbGFzc05hbWUgPSAnbnVtJzsKICAgIHRoLmNsYXNzTGlzdC5hZGQoJ3NvcnRhYmxlLXRoJyk7CiAgICBjb25zdCBpc0FjdGl2ZSA9IHNvcnRTdGF0ZS5rZXkgPT09IGtleTsKICAgIHRoLmFwcGVuZENoaWxkKGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKGxhYmVsKSk7CiAgICB0aC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3NwYW4nLCBpc0FjdGl2ZSA/IChzb3J0U3RhdGUuZGlyID09PSAnYXNjJyA/ICcg4oaRJyA6ICcg4oaTJykgOiAnIOKGlScsICdzb3J0LWFycm93JykpOwogICAgdGguYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7CiAgICAgIHNvcnRTdGF0ZSA9IHsga2V5LCBkaXI6IHNvcnRTdGF0ZS5rZXkgPT09IGtleSAmJiBzb3J0U3RhdGUuZGlyID09PSAnYXNjJyA/ICdkZXNjJyA6ICdhc2MnLCB0eXBlIH07CiAgICAgIHJlbmRlclRhYmxlKCk7CiAgICB9KTsKICAgIHJldHVybiB0aDsKICB9CgogIGZ1bmN0aW9uIGZpbHRlcmVkUm93cygpIHsKICAgIGNvbnN0IHEgPSBzZWFyY2hWYWx1ZS50cmltKCkudG9Mb3dlckNhc2UoKTsKICAgIGlmICghcSkgcmV0dXJuIGN1cnJlbnRSb3dzOwogICAgcmV0dXJuIGN1cnJlbnRSb3dzLmZpbHRlcigocm93KSA9PiB7CiAgICAgIGNvbnN0IGxhYmVsID0gcGxhdGZvcm1NZXRhRm9yKHJvdy5wbGF0Zm9ybSkubGFiZWwudG9Mb3dlckNhc2UoKTsKICAgICAgcmV0dXJuIGxhYmVsLmluY2x1ZGVzKHEpIHx8IHJvdy5lbnRyeV9kYXRlLmluY2x1ZGVzKHEpIHx8IFN0cmluZyhyb3cuZm9sbG93ZXJzX2NvdW50KS5pbmNsdWRlcyhxKTsKICAgIH0pOwogIH0KCiAgZnVuY3Rpb24gc29ydGVkUm93cygpIHsKICAgIGNvbnN0IHsga2V5LCBkaXIsIHR5cGUgfSA9IHNvcnRTdGF0ZTsKICAgIGNvbnN0IGZhY3RvciA9IGRpciA9PT0gJ2FzYycgPyAxIDogLTE7CiAgICByZXR1cm4gWy4uLmZpbHRlcmVkUm93cygpXS5zb3J0KChhLCBiKSA9PiB7CiAgICAgIGNvbnN0IGF2ID0gYVtrZXldOwogICAgICBjb25zdCBidiA9IGJba2V5XTsKICAgICAgaWYgKGF2ID09PSBudWxsIHx8IGF2ID09PSB1bmRlZmluZWQpIHJldHVybiAxOwogICAgICBpZiAoYnYgPT09IG51bGwgfHwgYnYgPT09IHVuZGVmaW5lZCkgcmV0dXJuIC0xOwogICAgICBpZiAodHlwZSA9PT0gJ251bWJlcicpIHJldHVybiAoYXYgLSBidikgKiBmYWN0b3I7CiAgICAgIHJldHVybiBTdHJpbmcoYXYpLmxvY2FsZUNvbXBhcmUoU3RyaW5nKGJ2KSkgKiBmYWN0b3I7CiAgICB9KTsKICB9CgogIC8qKiBFdmVyeSBjdXJyZW50bHktZmlsdGVyZWQvc29ydGVkIHJvdywgc2hhcGVkIGZvciBleHBvcnQgKG5vdCBqdXN0IHRoZSBjdXJyZW50IHBhZ2UpLiAqLwogIGZ1bmN0aW9uIGV4cG9ydFJvd3MoKSB7CiAgICByZXR1cm4gc29ydGVkUm93cygpLm1hcCgocm93KSA9PiAoewogICAgICBwbGF0Zm9ybV9sYWJlbDogcGxhdGZvcm1NZXRhRm9yKHJvdy5wbGF0Zm9ybSkubGFiZWwsCiAgICAgIGVudHJ5X2RhdGU6IHJvdy5lbnRyeV9kYXRlLAogICAgICBmb2xsb3dlcnNfY291bnQ6IHJvdy5mb2xsb3dlcnNfY291bnQsCiAgICAgIHVwZGF0ZWRfYXQ6IHJvdy51cGRhdGVkX2F0LAogICAgfSkpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyVGFibGUoKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZvbGxvd2Vyc1RhYmxlV3JhcCcpOwogICAgY29uc3QgcGFnZXJFbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmb2xsb3dlcnNQYWdlcicpOwogICAgaWYgKCF3cmFwKSByZXR1cm47CiAgICBjb25zdCBhbGxTb3J0ZWQgPSBzb3J0ZWRSb3dzKCk7CiAgICBpZiAoIWFsbFNvcnRlZC5sZW5ndGgpIHsKICAgICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgICAgd3JhcC5hcHBlbmRDaGlsZChlbXB0eVN0YXRlKHsKICAgICAgICBpY29uOiAndXNlcnMnLAogICAgICAgIHRpdGxlOiBjdXJyZW50Um93cy5sZW5ndGggPyAnTm8gZW50cmllcyBtYXRjaCB5b3VyIHNlYXJjaCcgOiAnTm8gZm9sbG93ZXIgZW50cmllcyB5ZXQnLAogICAgICAgIG1lc3NhZ2U6IGN1cnJlbnRSb3dzLmxlbmd0aCA/ICdUcnkgYSBkaWZmZXJlbnQgcGxhdGZvcm0gbmFtZSBvciBkYXRlLicgOiAnQWRkIHlvdXIgZmlyc3Qgd2Vla2x5IGZvbGxvd2VyIGNvdW50IGFib3ZlIGZvciBhbnkgcGxhdGZvcm0uJywKICAgICAgfSkpOwogICAgICBpZiAocGFnZXJFbCkgcGFnZXJFbC5pbm5lckhUTUwgPSAnJzsKICAgICAgcmV0dXJuOwogICAgfQogICAgY29uc3QgeyBwYWdlUm93cywgdG90YWxQYWdlcywgc2FmZVBhZ2UsIHRvdGFsIH0gPSBwYWdpbmF0ZUNsaWVudFNpZGUoYWxsU29ydGVkLCBwYWdlLCBwYWdlU2l6ZSk7CiAgICBwYWdlID0gc2FmZVBhZ2U7CgogICAgY29uc3QgdGFibGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0YWJsZScpOwogICAgdGFibGUuY2xhc3NOYW1lID0gJ2RhdGEtdGFibGUnOwogICAgY29uc3QgdGhlYWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0aGVhZCcpOwogICAgY29uc3QgaGVhZFRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgIGhlYWRUci5hcHBlbmQoCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdQbGF0Zm9ybScsICdwbGF0Zm9ybScsICdzdHJpbmcnKSwKICAgICAgc29ydGFibGVIZWFkZXIoJ1dlZWsgLyBEYXRlJywgJ2VudHJ5X2RhdGUnLCAnc3RyaW5nJyksCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdGb2xsb3dlcnMgQ291bnQnLCAnZm9sbG93ZXJzX2NvdW50JywgJ251bWJlcicpLAogICAgICB0ZXh0RWwoJ3RoJywgJ0xhc3QgVXBkYXRlZCcpLAogICAgICB0ZXh0RWwoJ3RoJywgJ0FjdGlvbnMnKQogICAgKTsKICAgIHRoZWFkLmFwcGVuZENoaWxkKGhlYWRUcik7CiAgICBjb25zdCB0Ym9keSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3Rib2R5Jyk7CiAgICBwYWdlUm93cy5mb3JFYWNoKChyb3cpID0+IHsKICAgICAgY29uc3QgdHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgICBjb25zdCBwbGF0Zm9ybVRkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgY29uc3QgbWV0YSA9IHBsYXRmb3JtTWV0YUZvcihyb3cucGxhdGZvcm0pOwogICAgICBjb25zdCBwaWxsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICBwaWxsLmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1waWxsJzsKICAgICAgY29uc3QgZG90ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICBkb3QuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLWRvdCc7CiAgICAgIGRvdC5zdHlsZS5iYWNrZ3JvdW5kID0gbWV0YS5jb2xvcjsKICAgICAgcGlsbC5hcHBlbmQoZG90LCBkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShtZXRhLmxhYmVsKSk7CiAgICAgIHBsYXRmb3JtVGQuYXBwZW5kQ2hpbGQocGlsbCk7CgogICAgICBjb25zdCBhY3Rpb25zVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBjb25zdCByb3dBY3Rpb25zID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIHJvd0FjdGlvbnMuY2xhc3NOYW1lID0gJ3Jvdy1hY3Rpb25zJzsKICAgICAgY29uc3QgZWRpdEJ0biA9IGljb25CdG4oJ2J0bicsICdwZW5jaWwnLCAnRWRpdCcpOwogICAgICBlZGl0QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gc3RhcnRFZGl0KHJvdykpOwogICAgICBjb25zdCBkZWxldGVCdG4gPSBpY29uQnRuKCdidG4gZGFuZ2VyJywgJ3RyYXNoLTInLCAnRGVsZXRlJyk7CiAgICAgIGRlbGV0ZUJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IGhhbmRsZURlbGV0ZShyb3cpKTsKICAgICAgcm93QWN0aW9ucy5hcHBlbmQoZWRpdEJ0biwgZGVsZXRlQnRuKTsKICAgICAgYWN0aW9uc1RkLmFwcGVuZENoaWxkKHJvd0FjdGlvbnMpOwoKICAgICAgdHIuYXBwZW5kKAogICAgICAgIHBsYXRmb3JtVGQsCiAgICAgICAgdGV4dEVsKCd0ZCcsIEZvcm1hdC5kYXRlKHJvdy5lbnRyeV9kYXRlKSksCiAgICAgICAgdGV4dEVsKCd0ZCcsIEZvcm1hdC5udW1iZXIocm93LmZvbGxvd2Vyc19jb3VudCksICdudW0nKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgcm93LnVwZGF0ZWRfYXQpLAogICAgICAgIGFjdGlvbnNUZAogICAgICApOwogICAgICB0Ym9keS5hcHBlbmRDaGlsZCh0cik7CiAgICB9KTsKICAgIHRhYmxlLmFwcGVuZCh0aGVhZCwgdGJvZHkpOwogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGFibGUpOwoKICAgIGlmIChwYWdlckVsKSB7CiAgICAgIHBhZ2VyRWwuaW5uZXJIVE1MID0gJyc7CiAgICAgIHBhZ2VyRWwuYXBwZW5kQ2hpbGQoYnVpbGRQYWdlcih7CiAgICAgICAgcGFnZTogc2FmZVBhZ2UsCiAgICAgICAgdG90YWxQYWdlcywKICAgICAgICB0b3RhbCwKICAgICAgICBvblByZXY6ICgpID0+IHsgcGFnZSAtPSAxOyByZW5kZXJUYWJsZSgpOyB9LAogICAgICAgIG9uTmV4dDogKCkgPT4geyBwYWdlICs9IDE7IHJlbmRlclRhYmxlKCk7IH0sCiAgICAgIH0pKTsKICAgIH0KICB9CgogIGFzeW5jIGZ1bmN0aW9uIGxvYWQoKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZvbGxvd2Vyc1RhYmxlV3JhcCcpOwogICAgaWYgKHdyYXApIHsgd3JhcC5pbm5lckhUTUwgPSAnJzsgd3JhcC5hcHBlbmRDaGlsZChza2VsZXRvblJvd3MoNCkpOyB9CiAgICBjdXJyZW50Um93cyA9IGF3YWl0IEFwaS5saXN0Rm9sbG93ZXJzKHt9KTsKICAgIHJlbmRlclRhYmxlKCk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiByZW5kZXIoKSB7CiAgICByb290ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3ZpZXctZm9sbG93ZXJzJyk7CiAgICBlZGl0aW5nSWQgPSBudWxsOwogICAgc2hlbGwoKTsKICAgIGF3YWl0IGxvYWQoKTsKICB9CgogIHJldHVybiB7IHJlbmRlciB9Owp9KSgpOwoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIEFwcCBib290c3RyYXA6IHRhYiByb3V0aW5nLCBmaWx0ZXIgYmFyIHdpcmluZywgdGhlbWUgdG9nZ2xlLgogICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KKCgpID0+IHsKICBjb25zdCBWSUVXUyA9IHsKICAgIGRhc2hib2FyZDogRGFzaGJvYXJkLAogICAgcmVjb3JkczogUmVjb3JkcywKICAgIGZvbGxvd2VyczogRm9sbG93ZXJzLAogICAgY29tcGFyaXNvbjogQ29tcGFyaXNvbiwKICAgIHVwbG9hZDogVXBsb2FkLAogICAgaGlzdG9yeTogSGlzdG9yeSwKICB9OwoKICBsZXQgYWN0aXZlVGFiID0gJ2Rhc2hib2FyZCc7CgogIGZ1bmN0aW9uIHN3aXRjaFRhYih0YWIpIHsKICAgIGFjdGl2ZVRhYiA9IHRhYjsKICAgIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJy50YWItYnRuJykuZm9yRWFjaCgoYnRuKSA9PiB7CiAgICAgIGNvbnN0IGlzQWN0aXZlID0gYnRuLmRhdGFzZXQudGFiID09PSB0YWI7CiAgICAgIGJ0bi5jbGFzc0xpc3QudG9nZ2xlKCdpcy1hY3RpdmUnLCBpc0FjdGl2ZSk7CiAgICAgIGJ0bi5zZXRBdHRyaWJ1dGUoJ2FyaWEtc2VsZWN0ZWQnLCBTdHJpbmcoaXNBY3RpdmUpKTsKICAgIH0pOwogICAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnLnZpZXcnKS5mb3JFYWNoKCh2aWV3KSA9PiB7CiAgICAgIHZpZXcuY2xhc3NMaXN0LnRvZ2dsZSgnaXMtYWN0aXZlJywgdmlldy5pZCA9PT0gYHZpZXctJHt0YWJ9YCk7CiAgICB9KTsKICAgIC8vIEZpbHRlcnMgYXBwbHkgdG8gRGFzaGJvYXJkIGFuZCBEYXRhIFJlY29yZHMgKENvbXBhcmlzb25zIGhhcyBpdHMgb3duIHJhbmdlIGNvbnRyb2xzKS4KICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJCYXInKS5zdHlsZS5kaXNwbGF5ID0gKHRhYiA9PT0gJ2Rhc2hib2FyZCcgfHwgdGFiID09PSAncmVjb3JkcycpID8gJ2ZsZXgnIDogJ25vbmUnOwogICAgcmVuZGVyQWN0aXZlVmlldygpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyQWN0aXZlVmlldygpIHsKICAgIGNvbnN0IHZpZXcgPSBWSUVXU1thY3RpdmVUYWJdOwogICAgaWYgKHZpZXcgJiYgdmlldy5yZW5kZXIpIHZpZXcucmVuZGVyKCk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiBsb2FkRmlsdGVyT3B0aW9ucygpIHsKICAgIGNvbnN0IG9wdGlvbnMgPSBhd2FpdCBBcGkuZmlsdGVyT3B0aW9ucygpOwogICAgd2luZG93Ll9fZmlsdGVyT3B0aW9uc0NhY2hlID0gb3B0aW9uczsKCiAgICBjb25zdCBwbGF0Zm9ybVNlbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJQbGF0Zm9ybScpOwogICAgcGxhdGZvcm1TZWwubGVuZ3RoID0gMTsKICAgIG9wdGlvbnMucGxhdGZvcm1zLmZvckVhY2goKHApID0+IHsKICAgICAgY29uc3Qgb3B0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7CiAgICAgIG9wdC52YWx1ZSA9IHAuaWQ7CiAgICAgIG9wdC50ZXh0Q29udGVudCA9IHAubGFiZWw7CiAgICAgIHBsYXRmb3JtU2VsLmFwcGVuZENoaWxkKG9wdCk7CiAgICB9KTsKCiAgICBjb25zdCBjYW1wYWlnblNlbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJDYW1wYWlnbicpOwogICAgY2FtcGFpZ25TZWwubGVuZ3RoID0gMTsKICAgIG9wdGlvbnMuY2FtcGFpZ25UeXBlcy5mb3JFYWNoKChjKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOwogICAgICBvcHQudmFsdWUgPSBjOwogICAgICBvcHQudGV4dENvbnRlbnQgPSBjOwogICAgICBjYW1wYWlnblNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CgogICAgY29uc3QgY29udGVudFNlbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJDb250ZW50VHlwZScpOwogICAgY29udGVudFNlbC5sZW5ndGggPSAxOwogICAgb3B0aW9ucy5jb250ZW50VHlwZXMuZm9yRWFjaCgoYykgPT4gewogICAgICBjb25zdCBvcHQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsKICAgICAgb3B0LnZhbHVlID0gYzsKICAgICAgb3B0LnRleHRDb250ZW50ID0gYzsKICAgICAgY29udGVudFNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgfQoKICBmdW5jdGlvbiB3aXJlRmlsdGVyQmFyKCkgewogICAgY29uc3QgZGF0ZUZyb20gPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmlsdGVyRGF0ZUZyb20nKTsKICAgIGNvbnN0IGRhdGVUbyA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJEYXRlVG8nKTsKICAgIGNvbnN0IHBsYXRmb3JtID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZpbHRlclBsYXRmb3JtJyk7CiAgICBjb25zdCBjYW1wYWlnbiA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJDYW1wYWlnbicpOwogICAgY29uc3QgY29udGVudFR5cGUgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmlsdGVyQ29udGVudFR5cGUnKTsKICAgIGNvbnN0IGYgPSBTdGF0ZS5nZXRGaWx0ZXJzKCk7CiAgICBkYXRlRnJvbS52YWx1ZSA9IGYuZGF0ZUZyb207CiAgICBkYXRlVG8udmFsdWUgPSBmLmRhdGVUbzsKCiAgICBmdW5jdGlvbiBhcHBseSgpIHsKICAgICAgU3RhdGUuc2V0RmlsdGVycyh7CiAgICAgICAgZGF0ZUZyb206IGRhdGVGcm9tLnZhbHVlLAogICAgICAgIGRhdGVUbzogZGF0ZVRvLnZhbHVlLAogICAgICAgIHBsYXRmb3JtOiBwbGF0Zm9ybS52YWx1ZSwKICAgICAgICBjYW1wYWlnblR5cGU6IGNhbXBhaWduLnZhbHVlLAogICAgICAgIGNvbnRlbnRUeXBlOiBjb250ZW50VHlwZS52YWx1ZSwKICAgICAgfSk7CiAgICB9CiAgICBbZGF0ZUZyb20sIGRhdGVUbywgcGxhdGZvcm0sIGNhbXBhaWduLCBjb250ZW50VHlwZV0uZm9yRWFjaCgoZWwpID0+IGVsLmFkZEV2ZW50TGlzdGVuZXIoJ2NoYW5nZScsIGFwcGx5KSk7CgogICAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnI2ZpbHRlclByZXNldHMgYnV0dG9uJykuZm9yRWFjaCgoYnRuKSA9PiB7CiAgICAgIGJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsKICAgICAgICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcjZmlsdGVyUHJlc2V0cyBidXR0b24nKS5mb3JFYWNoKChiKSA9PiBiLmNsYXNzTGlzdC5yZW1vdmUoJ2lzLWFjdGl2ZScpKTsKICAgICAgICBidG4uY2xhc3NMaXN0LmFkZCgnaXMtYWN0aXZlJyk7CiAgICAgICAgY29uc3QgcHJlc2V0ID0gYnRuLmRhdGFzZXQucHJlc2V0OwogICAgICAgIGNvbnN0IHRvZGF5ID0gbmV3IERhdGUoKTsKICAgICAgICBjb25zdCB0byA9IHRvZGF5LnRvSVNPU3RyaW5nKCkuc2xpY2UoMCwgMTApOwogICAgICAgIGxldCBmcm9tOwogICAgICAgIGlmIChwcmVzZXQgPT09ICdhbGwnKSB7CiAgICAgICAgICBjb25zdCBtaW4gPSAod2luZG93Ll9fZmlsdGVyT3B0aW9uc0NhY2hlICYmIHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZS5kYXRlUmFuZ2UubWluKSB8fCB0bzsKICAgICAgICAgIGZyb20gPSBtaW47CiAgICAgICAgfSBlbHNlIHsKICAgICAgICAgIGNvbnN0IGQgPSBuZXcgRGF0ZSh0b2RheSk7CiAgICAgICAgICBkLnNldERhdGUoZC5nZXREYXRlKCkgLSAoTnVtYmVyKHByZXNldCkgLSAxKSk7CiAgICAgICAgICBmcm9tID0gZC50b0lTT1N0cmluZygpLnNsaWNlKDAsIDEwKTsKICAgICAgICB9CiAgICAgICAgZGF0ZUZyb20udmFsdWUgPSBmcm9tOwogICAgICAgIGRhdGVUby52YWx1ZSA9IHRvOwogICAgICAgIGFwcGx5KCk7CiAgICAgIH0pOwogICAgfSk7CgogICAgU3RhdGUub25DaGFuZ2UoKCkgPT4gewogICAgICBpZiAoYWN0aXZlVGFiID09PSAnZGFzaGJvYXJkJykgRGFzaGJvYXJkLnJlbmRlcigpOwogICAgICBpZiAoYWN0aXZlVGFiID09PSAncmVjb3JkcycpIFJlY29yZHMucmVuZGVyKCk7CiAgICB9KTsKICB9CgogIGZ1bmN0aW9uIHdpcmVUYWJzKCkgewogICAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnLnRhYi1idG4nKS5mb3JFYWNoKChidG4pID0+IHsKICAgICAgYnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gc3dpdGNoVGFiKGJ0bi5kYXRhc2V0LnRhYikpOwogICAgfSk7CiAgfQoKICB3aW5kb3cuYWRkRXZlbnRMaXN0ZW5lcignbHJzOmRhdGEtdXBkYXRlZCcsIGFzeW5jICgpID0+IHsKICAgIGF3YWl0IGxvYWRGaWx0ZXJPcHRpb25zKCk7CiAgICByZW5kZXJBY3RpdmVWaWV3KCk7CiAgfSk7CgogIC8vIC0tLS0tLS0tLS0gQXV0aCBzY3JlZW4gLS0tLS0tLS0tLQogIGxldCBhcHBJbml0aWFsaXplZCA9IGZhbHNlOwoKICBmdW5jdGlvbiBzaG93QXV0aFNjcmVlbigpIHsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhdXRoU2NyZWVuJykuc3R5bGUuZGlzcGxheSA9ICdmbGV4JzsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhcHBTaGVsbCcpLnN0eWxlLmRpc3BsYXkgPSAnbm9uZSc7CiAgICBjb25zdCBjb2RlSW5wdXQgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXV0aENvZGUnKTsKICAgIGNvZGVJbnB1dC52YWx1ZSA9ICcnOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2F1dGhFcnJvcicpLnRleHRDb250ZW50ID0gJyc7CiAgICBjb2RlSW5wdXQuZm9jdXMoKTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIHNob3dBcHAoKSB7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXV0aFNjcmVlbicpLnN0eWxlLmRpc3BsYXkgPSAnbm9uZSc7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXBwU2hlbGwnKS5zdHlsZS5kaXNwbGF5ID0gJyc7CiAgICBpZiAoIWFwcEluaXRpYWxpemVkKSB7CiAgICAgIGFwcEluaXRpYWxpemVkID0gdHJ1ZTsKICAgICAgd2lyZVRhYnMoKTsKICAgICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2xvZ291dEJ0bicpLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgYXN5bmMgKCkgPT4gewogICAgICAgIGF3YWl0IEFwaS5hdXRoTG9nb3V0KCk7CiAgICAgICAgYXBwSW5pdGlhbGl6ZWQgPSBmYWxzZTsKICAgICAgICBzaG93QXV0aFNjcmVlbigpOwogICAgICB9KTsKICAgICAgYXdhaXQgbG9hZEZpbHRlck9wdGlvbnMoKTsKICAgICAgd2lyZUZpbHRlckJhcigpOwogICAgICBzd2l0Y2hUYWIoJ2Rhc2hib2FyZCcpOwogICAgfSBlbHNlIHsKICAgICAgYXdhaXQgbG9hZEZpbHRlck9wdGlvbnMoKTsKICAgICAgcmVuZGVyQWN0aXZlVmlldygpOwogICAgfQogIH0KCiAgYXN5bmMgZnVuY3Rpb24gc3VibWl0QXV0aCgpIHsKICAgIGNvbnN0IGVycm9yRWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXV0aEVycm9yJyk7CiAgICBjb25zdCBidG4gPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXV0aFN1Ym1pdEJ0bicpOwogICAgY29uc3QgY29kZUlucHV0ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2F1dGhDb2RlJyk7CiAgICBlcnJvckVsLnRleHRDb250ZW50ID0gJyc7CiAgICBidG4uZGlzYWJsZWQgPSB0cnVlOwogICAgYnRuLnRleHRDb250ZW50ID0gJ0NoZWNraW5n4oCmJzsKICAgIHRyeSB7CiAgICAgIGF3YWl0IEFwaS5hdXRoTG9naW4oY29kZUlucHV0LnZhbHVlKTsKICAgICAgYXdhaXQgc2hvd0FwcCgpOwogICAgfSBjYXRjaCAoZXJyKSB7CiAgICAgIGVycm9yRWwudGV4dENvbnRlbnQgPSBlcnIubWVzc2FnZTsKICAgIH0gZmluYWxseSB7CiAgICAgIGJ0bi5kaXNhYmxlZCA9IGZhbHNlOwogICAgICBidG4uaW5uZXJIVE1MID0gJzxpIGRhdGEtbHVjaWRlPSJhcnJvdy1yaWdodCIgc3R5bGU9IndpZHRoOjE0cHg7aGVpZ2h0OjE0cHg7Ij48L2k+IEVudGVyJzsKICAgIH0KICB9CgogIGZ1bmN0aW9uIHdpcmVBdXRoRm9ybSgpIHsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhdXRoU3VibWl0QnRuJykuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBzdWJtaXRBdXRoKTsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhdXRoQ29kZScpLmFkZEV2ZW50TGlzdGVuZXIoJ2tleWRvd24nLCAoZSkgPT4geyBpZiAoZS5rZXkgPT09ICdFbnRlcicpIHN1Ym1pdEF1dGgoKTsgfSk7CiAgfQoKICB3aW5kb3cuYWRkRXZlbnRMaXN0ZW5lcignbHJzOnNpZ25lZC1vdXQnLCAoKSA9PiB7CiAgICBhcHBJbml0aWFsaXplZCA9IGZhbHNlOwogICAgc2hvd0F1dGhTY3JlZW4oKTsKICB9KTsKCiAgYXN5bmMgZnVuY3Rpb24gaW5pdCgpIHsKICAgIGFwcGx5QnJhbmRpbmcoKTsKICAgIHdpcmVBdXRoRm9ybSgpOwogICAgY29uc3QgeyBhdXRoZW50aWNhdGVkIH0gPSBhd2FpdCBBcGkuYXV0aE1lKCk7CiAgICBpZiAoYXV0aGVudGljYXRlZCkgYXdhaXQgc2hvd0FwcCgpOwogICAgZWxzZSBzaG93QXV0aFNjcmVlbigpOwogIH0KCiAgLy8gSWNvbnMgYXJlIHBsYWNlZCBhcyA8aSBkYXRhLWx1Y2lkZT0iLi4uIj4gcGxhY2Vob2xkZXJzIHRocm91Z2hvdXQgdGhlIGR5bmFtaWNhbGx5CiAgLy8gcmVuZGVyZWQgVUk7IEx1Y2lkZSByZXBsYWNlcyBlYWNoIHdpdGggYW4gaW5saW5lIFNWRy4gUmF0aGVyIHRoYW4gcmVtZW1iZXJpbmcgdG8gY2FsbAogIC8vIHRoaXMgYWZ0ZXIgZXZlcnkgc2luZ2xlIHJlbmRlciwgb25lIG9ic2VydmVyIGNhdGNoZXMgZXZlcnkgRE9NIGNoYW5nZSB0aGF0IGNvdWxkIGhhdmUKICAvLyBpbnRyb2R1Y2VkIGEgbmV3IHBsYWNlaG9sZGVyLgogIGlmICh3aW5kb3cubHVjaWRlKSB7CiAgICB3aW5kb3cubHVjaWRlLmNyZWF0ZUljb25zKCk7CiAgICAvLyBjcmVhdGVJY29ucygpIHJlcGxhY2VzIDxpIGRhdGEtbHVjaWRlPiBwbGFjZWhvbGRlcnMgd2l0aCA8c3ZnPiDigJQgaXRzZWxmIGEgRE9NCiAgICAvLyBtdXRhdGlvbi4gV2l0aG91dCBkaXNjb25uZWN0aW5nIGZpcnN0LCB0aGF0IHdyaXRlIHJlLXRyaWdnZXJzIHRoaXMgc2FtZSBvYnNlcnZlcgogICAgLy8gZm9yZXZlciAoYW4gaW5maW5pdGUgbXV0YXRlL29ic2VydmUgbG9vcCB0aGF0IHBlZ3MgdGhlIENQVSBhbmQgY3Jhc2hlcyB0aGUgdGFiKS4KICAgIC8vIERpc2Nvbm5lY3RpbmcgYmVmb3JlIGVhY2ggcGFzcyBhbmQgcmVjb25uZWN0aW5nIGFmdGVyLCBwbHVzIGJhdGNoaW5nIGJ1cnN0cyBvZgogICAgLy8gbXV0YXRpb25zIGludG8gYSBzaW5nbGUgbWljcm90YXNrLCBicmVha3MgdGhlIGN5Y2xlLgogICAgbGV0IGljb25zU2NoZWR1bGVkID0gZmFsc2U7CiAgICBjb25zdCBpY29uT2JzZXJ2ZXIgPSBuZXcgTXV0YXRpb25PYnNlcnZlcigoKSA9PiB7CiAgICAgIGlmIChpY29uc1NjaGVkdWxlZCkgcmV0dXJuOwogICAgICBpY29uc1NjaGVkdWxlZCA9IHRydWU7CiAgICAgIHF1ZXVlTWljcm90YXNrKCgpID0+IHsKICAgICAgICBpY29uc1NjaGVkdWxlZCA9IGZhbHNlOwogICAgICAgIGljb25PYnNlcnZlci5kaXNjb25uZWN0KCk7CiAgICAgICAgd2luZG93Lmx1Y2lkZS5jcmVhdGVJY29ucygpOwogICAgICAgIGljb25PYnNlcnZlci5vYnNlcnZlKGRvY3VtZW50LmJvZHksIHsgY2hpbGRMaXN0OiB0cnVlLCBzdWJ0cmVlOiB0cnVlIH0pOwogICAgICB9KTsKICAgIH0pOwogICAgaWNvbk9ic2VydmVyLm9ic2VydmUoZG9jdW1lbnQuYm9keSwgeyBjaGlsZExpc3Q6IHRydWUsIHN1YnRyZWU6IHRydWUgfSk7CiAgfQoKICBpbml0KCk7Cn0pKCk7Cjwvc2NyaXB0Pgo8L2JvZHk+CjwvaHRtbD4K';
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
