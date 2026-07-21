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
  x: [
    { key: 'impressions', label: 'Impressions' },
    { key: 'engagement', label: 'Engagements' },
    { key: 'clicks', label: 'Clicks' },
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

/** Every followers_history row, grouped by platform, oldest-to-newest — the shared, unfiltered base that followersKpis() and platformSummaryBar() both reduce over. Unlike followersGrowth() above (which is deliberately range-filterable for the growth chart), these need the true full history to find the actual latest/previous entry regardless of whatever date range the Dashboard filter bar has selected. */
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
 * Per-platform row for the Dashboard's Platform Summary Bar: posts/views/
 * reach/engagement from the existing platformBreakdown() analytics query,
 * joined with each platform's latest Followers Data Record entry. A
 * platform shows up if it has either side of that data (uploaded posts OR
 * follower entries) — both are legitimate signs of platform activity.
 */
function platformSummaryBar(filters = {}) {
  const breakdown = platformBreakdown(filters);
  const byId = Object.fromEntries(breakdown.map((r) => [r.platform, r]));
  const followers = allFollowersByPlatform();

  const activeIds = PLATFORM_IDS.filter((id) => byId[id] || followers[id]);
  return activeIds.map((id) => {
    const meta = PLATFORMS.find((p) => p.id === id);
    const row = byId[id] || {};
    const entries = followers[id] || [];
    return {
      platform: id,
      label: meta.label,
      color: meta.color,
      followers: entries.length ? entries[entries.length - 1].followers_count : null,
      posts: row.post_count || 0,
      views: row.views || 0,
      reach: row.reach || 0,
      engagement: row.engagement || 0,
    };
  });
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

analyticsRouter.get('/platform-summary', (req, res) => {
  const { dateFrom, dateTo } = req.query;
  res.json(platformSummaryBar({ ...filtersFromQuery(req.query), dateFrom, dateTo }));
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
const INDEX_HTML_BASE64 = 'PCFkb2N0eXBlIGh0bWw+CjxodG1sIGxhbmc9ImVuIj4KPGhlYWQ+CjxtZXRhIGNoYXJzZXQ9IlVURi04IiAvPgo8bWV0YSBuYW1lPSJ2aWV3cG9ydCIgY29udGVudD0id2lkdGg9ZGV2aWNlLXdpZHRoLCBpbml0aWFsLXNjYWxlPTEuMCIgLz4KPHRpdGxlPkxSUyBBbmFseXRpY3MgRGFzaGJvYXJkPC90aXRsZT4KPGxpbmsgcmVsPSJpY29uIiB0eXBlPSJpbWFnZS9wbmciIGlkPSJmYXZpY29uTGluayIgLz4KPGxpbmsgcmVsPSJwcmVjb25uZWN0IiBocmVmPSJodHRwczovL2ZvbnRzLmdvb2dsZWFwaXMuY29tIiAvPgo8bGluayByZWw9InByZWNvbm5lY3QiIGhyZWY9Imh0dHBzOi8vZm9udHMuZ3N0YXRpYy5jb20iIGNyb3Nzb3JpZ2luIC8+CjxsaW5rIGhyZWY9Imh0dHBzOi8vZm9udHMuZ29vZ2xlYXBpcy5jb20vY3NzMj9mYW1pbHk9SW50ZXI6d2dodEA0MDA7NTAwOzYwMDs3MDA7ODAwJmRpc3BsYXk9c3dhcCIgcmVsPSJzdHlsZXNoZWV0IiAvPgo8c2NyaXB0IHNyYz0iaHR0cHM6Ly9jZG4uanNkZWxpdnIubmV0L25wbS9jaGFydC5qc0A0LjQuNC9kaXN0L2NoYXJ0LnVtZC5taW4uanMiPjwvc2NyaXB0Pgo8c2NyaXB0IHNyYz0iaHR0cHM6Ly9jZG4uanNkZWxpdnIubmV0L25wbS9jaGFydGpzLXBsdWdpbi1kYXRhbGFiZWxzQDIvZGlzdC9jaGFydGpzLXBsdWdpbi1kYXRhbGFiZWxzLm1pbi5qcyI+PC9zY3JpcHQ+CjxzY3JpcHQgc3JjPSJodHRwczovL2Nkbi5qc2RlbGl2ci5uZXQvbnBtL2x1Y2lkZUAwLjQ2Mi4wL2Rpc3QvdW1kL2x1Y2lkZS5taW4uanMiPjwvc2NyaXB0Pgo8c3R5bGU+Ci8qIC0tLS0tLS0tLS0gRGVzaWduIHRva2VuczogZ2xhc3Ntb3JwaGlzbSBzdXJmYWNlcyArIHZhbGlkYXRlZCBjYXRlZ29yaWNhbC9zdGF0dXMgcGFsZXR0ZSAtLS0tLS0tLS0tICovCjpyb290IHsKICBjb2xvci1zY2hlbWU6IGxpZ2h0OwogIC0tZm9udC1zYW5zOiAnSW50ZXInLCAtYXBwbGUtc3lzdGVtLCBCbGlua01hY1N5c3RlbUZvbnQsICdTRiBQcm8gRGlzcGxheScsICdTZWdvZSBVSScsIFJvYm90bywgc2Fucy1zZXJpZjsKCiAgLS1wYWdlLXBsYW5lOiBsaW5lYXItZ3JhZGllbnQoMTgwZGVnLCAjZjdmOGZiIDAlLCAjZWNlZWYzIDEwMCUpOwogIC0tcGFnZS1wbGFuZS1zb2xpZDogI2VjZWVmMzsKICAtLXN1cmZhY2UtMTogcmdiYSgyNTUsIDI1NSwgMjU1LCAwLjY4KTsgLyogZ2xhc3M6IGNhcmRzLCBLUEkgdGlsZXMsIHRvcGJhciwgZmlsdGVyIGJhciAqLwogIC0tc3VyZmFjZS0yOiByZ2JhKDI1NSwgMjU1LCAyNTUsIDAuNTUpOyAvKiBnbGFzczogaW5wdXRzLCBuZXN0ZWQgcm93cywgcGlsbHMgKi8KICAtLXN1cmZhY2Utc29saWQ6ICNmZmZmZmY7CiAgLS1nbGFzcy1ibHVyOiBibHVyKDIwcHgpOwogIC0tYm9yZGVyOiByZ2JhKDE1LCAxNywgMjEsIDAuMDgpOwogIC0tdGV4dC1wcmltYXJ5OiAjMGYxMTE1OwogIC0tdGV4dC1zZWNvbmRhcnk6ICM1NjViNjY7CiAgLS10ZXh0LW11dGVkOiAjOGE4ZjlhOwogIC0tZ3JpZGxpbmU6IHJnYmEoMTUsIDE3LCAyMSwgMC4wOCk7CiAgLS1iYXNlbGluZTogcmdiYSgxNSwgMTcsIDIxLCAwLjIpOwogIC0tc3VjY2Vzcy10ZXh0OiAjMWE3ZDNjOwoKICAtLXN0YXR1cy1nb29kOiAjMWE5YzRhOwogIC0tc3RhdHVzLXdhcm5pbmc6ICNlMDhhMWY7CiAgLS1zdGF0dXMtc2VyaW91czogI2VjODM1YTsKICAtLXN0YXR1cy1jcml0aWNhbDogI2QxNDAzZjsKCiAgLS1zZXJpZXMtMTogIzJhNzhkNjsgLyogTFJTIGJsdWUg4oCUIGJyYW5kICsgZmFjZWJvb2sgKi8KICAtLXNlcmllcy0yOiAjMDA4MzAwOyAvKiBpbnN0YWdyYW0gKi8KICAtLXNlcmllcy0zOiAjZTg3YmE0OyAvKiB0aWt0b2sgKi8KICAtLXNlcmllcy00OiAjZWRhMTAwOyAvKiBsaW5rZWRpbiAqLwogIC0tc2VyaWVzLTU6ICMxYmFmN2E7IC8qIHRocmVhZHMgKi8KICAtLXNlcmllcy02OiAjZWI2ODM0OyAvKiB5b3V0dWJlICovCiAgLS1zZXJpZXMtNzogIzRhM2FhNzsgLyogcmVzZXJ2ZWQgKi8KICAtLXNlcmllcy04OiAjZTM0OTQ4OyAvKiByZXNlcnZlZCAqLwoKICAtLXJhZGl1cy1zbTogMTBweDsKICAtLXJhZGl1cy1tZDogMTRweDsKICAtLXJhZGl1cy1sZzogMThweDsKCiAgLS1zaGFkb3ctY2FyZDogMCAxcHggMnB4IHJnYmEoMTUsMTcsMjEsMC4wMyksIDAgOHB4IDI0cHggLTEwcHggcmdiYSgxNSwxNywyMSwwLjE0KTsKICAtLXNoYWRvdy1ob3ZlcjogMCA2cHggMTJweCAtMnB4IHJnYmEoMTUsMTcsMjEsMC4wOCksIDAgMThweCA0MHB4IC0xNHB4IHJnYmEoMTUsMTcsMjEsMC4yKTsKICAtLXNoYWRvdy1tb2RhbDogMCAyNHB4IDY0cHggLTEycHggcmdiYSgxNSwxNywyMSwwLjM1KTsKICAtLWVhc2U6IGN1YmljLWJlemllcigwLjQsIDAsIDAuMiwgMSk7Cn0KCkBtZWRpYSAocHJlZmVycy1jb2xvci1zY2hlbWU6IGRhcmspIHsKICA6cm9vdDp3aGVyZSg6bm90KFtkYXRhLXRoZW1lPSJsaWdodCJdKSkgewogICAgY29sb3Itc2NoZW1lOiBkYXJrOwogICAgLS1wYWdlLXBsYW5lOiBsaW5lYXItZ3JhZGllbnQoMTgwZGVnLCAjMGIwYzBmIDAlLCAjMTcxODFkIDEwMCUpOwogICAgLS1wYWdlLXBsYW5lLXNvbGlkOiAjMGIwYzBmOwogICAgLS1zdXJmYWNlLTE6IHJnYmEoMzIsIDM0LCA0MCwgMC42Mik7CiAgICAtLXN1cmZhY2UtMjogcmdiYSgyNTUsIDI1NSwgMjU1LCAwLjA2KTsKICAgIC0tc3VyZmFjZS1zb2xpZDogIzFjMWQyMzsKICAgIC0tYm9yZGVyOiByZ2JhKDI1NSwgMjU1LCAyNTUsIDAuMDkpOwogICAgLS10ZXh0LXByaW1hcnk6ICNmNGY1Zjc7CiAgICAtLXRleHQtc2Vjb25kYXJ5OiAjYjhiYmM0OwogICAgLS10ZXh0LW11dGVkOiAjODI4NjhmOwogICAgLS1ncmlkbGluZTogcmdiYSgyNTUsIDI1NSwgMjU1LCAwLjA4KTsKICAgIC0tYmFzZWxpbmU6IHJnYmEoMjU1LCAyNTUsIDI1NSwgMC4yKTsKICAgIC0tc3VjY2Vzcy10ZXh0OiAjMzRjNzZmOwoKICAgIC0tc3RhdHVzLWdvb2Q6ICMyZmI4NjI7CiAgICAtLXN0YXR1cy13YXJuaW5nOiAjZjBhMTNhOwogICAgLS1zdGF0dXMtc2VyaW91czogI2VjODM1YTsKICAgIC0tc3RhdHVzLWNyaXRpY2FsOiAjZTA2MDVmOwoKICAgIC0tc2VyaWVzLTE6ICMzOTg3ZTU7CiAgICAtLXNlcmllcy0yOiAjMDA4MzAwOwogICAgLS1zZXJpZXMtMzogI2Q1NTE4MTsKICAgIC0tc2VyaWVzLTQ6ICNjOTg1MDA7CiAgICAtLXNlcmllcy01OiAjMTk5ZTcwOwogICAgLS1zZXJpZXMtNjogI2Q5NTkyNjsKICAgIC0tc2VyaWVzLTc6ICM5MDg1ZTk7CiAgICAtLXNlcmllcy04OiAjZTY2NzY3OwoKICAgIC0tc2hhZG93LWNhcmQ6IDAgMXB4IDJweCByZ2JhKDAsMCwwLDAuMiksIDAgOHB4IDI0cHggLTEwcHggcmdiYSgwLDAsMCwwLjUpOwogICAgLS1zaGFkb3ctaG92ZXI6IDAgNnB4IDEycHggLTJweCByZ2JhKDAsMCwwLDAuMyksIDAgMThweCA0MHB4IC0xNHB4IHJnYmEoMCwwLDAsMC42KTsKICAgIC0tc2hhZG93LW1vZGFsOiAwIDI0cHggNjRweCAtMTJweCByZ2JhKDAsMCwwLDAuNyk7CiAgfQp9Cjpyb290W2RhdGEtdGhlbWU9ImRhcmsiXSB7CiAgY29sb3Itc2NoZW1lOiBkYXJrOwogIC0tcGFnZS1wbGFuZTogbGluZWFyLWdyYWRpZW50KDE4MGRlZywgIzBiMGMwZiAwJSwgIzE3MTgxZCAxMDAlKTsKICAtLXBhZ2UtcGxhbmUtc29saWQ6ICMwYjBjMGY7CiAgLS1zdXJmYWNlLTE6IHJnYmEoMzIsIDM0LCA0MCwgMC42Mik7CiAgLS1zdXJmYWNlLTI6IHJnYmEoMjU1LCAyNTUsIDI1NSwgMC4wNik7CiAgLS1zdXJmYWNlLXNvbGlkOiAjMWMxZDIzOwogIC0tYm9yZGVyOiByZ2JhKDI1NSwgMjU1LCAyNTUsIDAuMDkpOwogIC0tdGV4dC1wcmltYXJ5OiAjZjRmNWY3OwogIC0tdGV4dC1zZWNvbmRhcnk6ICNiOGJiYzQ7CiAgLS10ZXh0LW11dGVkOiAjODI4NjhmOwogIC0tZ3JpZGxpbmU6IHJnYmEoMjU1LCAyNTUsIDI1NSwgMC4wOCk7CiAgLS1iYXNlbGluZTogcmdiYSgyNTUsIDI1NSwgMjU1LCAwLjIpOwogIC0tc3VjY2Vzcy10ZXh0OiAjMzRjNzZmOwoKICAtLXN0YXR1cy1nb29kOiAjMmZiODYyOwogIC0tc3RhdHVzLXdhcm5pbmc6ICNmMGExM2E7CiAgLS1zdGF0dXMtc2VyaW91czogI2VjODM1YTsKICAtLXN0YXR1cy1jcml0aWNhbDogI2UwNjA1ZjsKCiAgLS1zZXJpZXMtMTogIzM5ODdlNTsKICAtLXNlcmllcy0yOiAjMDA4MzAwOwogIC0tc2VyaWVzLTM6ICNkNTUxODE7CiAgLS1zZXJpZXMtNDogI2M5ODUwMDsKICAtLXNlcmllcy01OiAjMTk5ZTcwOwogIC0tc2VyaWVzLTY6ICNkOTU5MjY7CiAgLS1zZXJpZXMtNzogIzkwODVlOTsKICAtLXNlcmllcy04OiAjZTY2NzY3OwoKICAtLXNoYWRvdy1jYXJkOiAwIDFweCAycHggcmdiYSgwLDAsMCwwLjIpLCAwIDhweCAyNHB4IC0xMHB4IHJnYmEoMCwwLDAsMC41KTsKICAtLXNoYWRvdy1ob3ZlcjogMCA2cHggMTJweCAtMnB4IHJnYmEoMCwwLDAsMC4zKSwgMCAxOHB4IDQwcHggLTE0cHggcmdiYSgwLDAsMCwwLjYpOwogIC0tc2hhZG93LW1vZGFsOiAwIDI0cHggNjRweCAtMTJweCByZ2JhKDAsMCwwLDAuNyk7Cn0KCiogeyBib3gtc2l6aW5nOiBib3JkZXItYm94OyB9Cmh0bWwsIGJvZHkgeyBoZWlnaHQ6IDEwMCU7IH0KYm9keSB7CiAgbWFyZ2luOiAwOwogIGZvbnQtZmFtaWx5OiB2YXIoLS1mb250LXNhbnMpOwogIGJhY2tncm91bmQ6IHZhcigtLXBhZ2UtcGxhbmUpOwogIGJhY2tncm91bmQtYXR0YWNobWVudDogZml4ZWQ7CiAgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7CiAgLXdlYmtpdC1mb250LXNtb290aGluZzogYW50aWFsaWFzZWQ7CiAgLW1vei1vc3gtZm9udC1zbW9vdGhpbmc6IGdyYXlzY2FsZTsKfQpidXR0b24sIHNlbGVjdCwgaW5wdXQsIHRleHRhcmVhIHsgZm9udC1mYW1pbHk6IGluaGVyaXQ7IH0KaDEsIGgyLCBoMywgaDQgeyBmb250LXdlaWdodDogNzAwOyBsZXR0ZXItc3BhY2luZzogLTAuMDFlbTsgfQoKOjpzZWxlY3Rpb24geyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDMwJSwgdHJhbnNwYXJlbnQpOyB9CgovKiBDdXN0b20gc2Nyb2xsYmFyIOKAlCB0aGluLCB1bm9idHJ1c2l2ZSwgZml0cyB0aGUgZ2xhc3MgYWVzdGhldGljICovCjo6LXdlYmtpdC1zY3JvbGxiYXIgeyB3aWR0aDogMTBweDsgaGVpZ2h0OiAxMHB4OyB9Cjo6LXdlYmtpdC1zY3JvbGxiYXItdHJhY2sgeyBiYWNrZ3JvdW5kOiB0cmFuc3BhcmVudDsgfQo6Oi13ZWJraXQtc2Nyb2xsYmFyLXRodW1iIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXRleHQtbXV0ZWQpIDQwJSwgdHJhbnNwYXJlbnQpOyBib3JkZXItcmFkaXVzOiAyMHB4OyBib3JkZXI6IDJweCBzb2xpZCB0cmFuc3BhcmVudDsgYmFja2dyb3VuZC1jbGlwOiBwYWRkaW5nLWJveDsgfQo6Oi13ZWJraXQtc2Nyb2xsYmFyLXRodW1iOmhvdmVyIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXRleHQtbXV0ZWQpIDYwJSwgdHJhbnNwYXJlbnQpOyBiYWNrZ3JvdW5kLWNsaXA6IHBhZGRpbmctYm94OyB9CgouYXBwLXNoZWxsIHsgbWluLWhlaWdodDogMTAwJTsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgfQoKLyogLS0tLS0tLS0tLSBUb3BiYXIgLS0tLS0tLS0tLSAqLwoudG9wYmFyIHsKICBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDI0cHg7CiAgcGFkZGluZzogMTJweCAyMHB4OwogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOwogIHBvc2l0aW9uOiBzdGlja3k7IHRvcDogMDsgei1pbmRleDogMjA7Cn0KLnRvcGJhci1icmFuZCB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogOHB4OyB3aGl0ZS1zcGFjZTogbm93cmFwOyB9Ci5icmFuZC1sb2dvIHsgaGVpZ2h0OiAyOHB4OyB3aWR0aDogYXV0bzsgZGlzcGxheTogYmxvY2s7IGZsZXgtc2hyaW5rOiAwOyBvYmplY3QtZml0OiBjb250YWluOyB9Ci5icmFuZC10aXRsZSB7IGZvbnQtd2VpZ2h0OiA2MDA7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyBsZXR0ZXItc3BhY2luZzogLTAuMDFlbTsgfQoKLnRhYnMgeyBkaXNwbGF5OiBmbGV4OyBnYXA6IDJweDsgZmxleDogMTsgb3ZlcmZsb3cteDogYXV0bzsgcG9zaXRpb246IHJlbGF0aXZlOyB9Ci50YWItYnRuIHsKICBkaXNwbGF5OiBpbmxpbmUtZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA3cHg7CiAgYm9yZGVyOiBub25lOyBiYWNrZ3JvdW5kOiB0cmFuc3BhcmVudDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsKICBwYWRkaW5nOiA5cHggMTZweDsgYm9yZGVyLXJhZGl1czogMTBweDsgY3Vyc29yOiBwb2ludGVyOyBmb250LXNpemU6IDE0cHg7IGZvbnQtd2VpZ2h0OiA1MDA7CiAgd2hpdGUtc3BhY2U6IG5vd3JhcDsgcG9zaXRpb246IHJlbGF0aXZlOwogIHRyYW5zaXRpb246IGNvbG9yIDE4MG1zIHZhcigtLWVhc2UpLCBiYWNrZ3JvdW5kIDE4MG1zIHZhcigtLWVhc2UpOwp9Ci50YWItYnRuIHN2ZyB7IGZsZXgtc2hyaW5rOiAwOyBvcGFjaXR5OiAwLjg7IH0KLnRhYi1idG4uaXMtYWN0aXZlIHN2ZyB7IG9wYWNpdHk6IDE7IGNvbG9yOiB2YXIoLS1zZXJpZXMtMSk7IH0KLnRhYi1idG46aG92ZXIgeyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDglLCB0cmFuc3BhcmVudCk7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyB9Ci50YWItYnRuLmlzLWFjdGl2ZSB7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyBmb250LXdlaWdodDogNjAwOyB9Ci50YWItYnRuLmlzLWFjdGl2ZTo6YWZ0ZXIgewogIGNvbnRlbnQ6ICcnOyBwb3NpdGlvbjogYWJzb2x1dGU7IGxlZnQ6IDEycHg7IHJpZ2h0OiAxMnB4OyBib3R0b206IC0xcHg7IGhlaWdodDogMnB4OwogIGJhY2tncm91bmQ6IHZhcigtLXNlcmllcy0xKTsgYm9yZGVyLXJhZGl1czogMnB4IDJweCAwIDA7CiAgYW5pbWF0aW9uOiB0YWJJbmRpY2F0b3JJbiAyMjBtcyB2YXIoLS1lYXNlKTsKfQpAa2V5ZnJhbWVzIHRhYkluZGljYXRvckluIHsgZnJvbSB7IG9wYWNpdHk6IDA7IHRyYW5zZm9ybTogc2NhbGVYKDAuNCk7IH0gdG8geyBvcGFjaXR5OiAxOyB0cmFuc2Zvcm06IHNjYWxlWCgxKTsgfSB9CgoudGhlbWUtdG9nZ2xlIHsKICBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTIpOyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgd2lkdGg6IDM2cHg7IGhlaWdodDogMzZweDsgYm9yZGVyLXJhZGl1czogMTBweDsgY3Vyc29yOiBwb2ludGVyOyBmb250LXNpemU6IDE1cHg7CiAgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsganVzdGlmeS1jb250ZW50OiBjZW50ZXI7CiAgdHJhbnNpdGlvbjogdHJhbnNmb3JtIDE4MG1zIHZhcigtLWVhc2UpLCBib3gtc2hhZG93IDE4MG1zIHZhcigtLWVhc2UpLCBiYWNrZ3JvdW5kIDE4MG1zIHZhcigtLWVhc2UpOwp9Ci50aGVtZS10b2dnbGU6aG92ZXIgeyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoLTFweCk7IGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdy1ob3Zlcik7IH0KLnRoZW1lLXRvZ2dsZTphY3RpdmUgeyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMCkgc2NhbGUoMC45NCk7IH0KLnRvcGJhci11c2VyIHsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiAxMHB4OyBmb250LXNpemU6IDEzcHg7IH0KCi8qIC0tLS0tLS0tLS0gQXV0aCBzY3JlZW4gLS0tLS0tLS0tLSAqLwouYXV0aC1zY3JlZW4gewogIG1pbi1oZWlnaHQ6IDEwMHZoOyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6IGNlbnRlcjsKICBiYWNrZ3JvdW5kOiB2YXIoLS1wYWdlLXBsYW5lKTsgcGFkZGluZzogMjBweDsKfQouYXV0aC1jYXJkIHsKICB3aWR0aDogMTAwJTsgbWF4LXdpZHRoOiA0MDBweDsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLWxnKTsgcGFkZGluZzogMzJweDsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LW1vZGFsKTsKICBhbmltYXRpb246IG1vZGFsUGFuZWxJbiAyNjBtcyB2YXIoLS1lYXNlKTsKfQouYXV0aC1icmFuZCB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogOHB4OyBtYXJnaW4tYm90dG9tOiAyMnB4OyB9Ci5hdXRoLWJyYW5kIC5icmFuZC10aXRsZSB7IGZvbnQtd2VpZ2h0OiA3MDA7IGZvbnQtc2l6ZTogMTdweDsgfQouYXV0aC1icmFuZCAuYnJhbmQtbG9nbyB7IGhlaWdodDogMzZweDsgfQouYXV0aC1mb3JtIHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiAxNHB4OyBtYXJnaW4tdG9wOiAxNnB4OyB9Ci5hdXRoLWZvcm0gLmZvcm0tZmllbGQgaW5wdXQgeyB3aWR0aDogMTAwJTsgfQouYXV0aC1lcnJvciB7IGNvbG9yOiB2YXIoLS1zdGF0dXMtY3JpdGljYWwpOyBmb250LXNpemU6IDEycHg7IG1pbi1oZWlnaHQ6IDE2cHg7IH0KCi8qIC0tLS0tLS0tLS0gRmlsdGVyIGJhciAtLS0tLS0tLS0tICovCi5maWx0ZXItYmFyIHsKICBkaXNwbGF5OiBmbGV4OyBmbGV4LXdyYXA6IHdyYXA7IGFsaWduLWl0ZW1zOiBlbmQ7IGdhcDogMTZweDsKICBwYWRkaW5nOiAxNHB4IDIwcHg7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOwogIHBvc2l0aW9uOiBzdGlja3k7IHRvcDogNTdweDsgei1pbmRleDogMTk7Cn0KLmZpbHRlci1maWVsZCB7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGdhcDogNXB4OyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IH0KLmZpbHRlci1maWVsZCBsYWJlbCB7IGZvbnQtd2VpZ2h0OiA2MDA7IH0KLmZpbHRlci1wcmVzZXRzIHsgZmxleC1kaXJlY3Rpb246IHJvdzsgZ2FwOiA2cHg7IH0KLmZpbHRlci1wcmVzZXRzIGJ1dHRvbiB7CiAgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0yKTsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgYm9yZGVyLXJhZGl1czogMjBweDsgcGFkZGluZzogN3B4IDEzcHg7IGZvbnQtc2l6ZTogMTJweDsgZm9udC13ZWlnaHQ6IDUwMDsgY3Vyc29yOiBwb2ludGVyOwogIHRyYW5zaXRpb246IGNvbG9yIDE4MG1zIHZhcigtLWVhc2UpLCBiYWNrZ3JvdW5kIDE4MG1zIHZhcigtLWVhc2UpLCB0cmFuc2Zvcm0gMTUwbXMgdmFyKC0tZWFzZSk7Cn0KLmZpbHRlci1wcmVzZXRzIGJ1dHRvbjpob3ZlciB7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoLTFweCk7IH0KLmZpbHRlci1wcmVzZXRzIGJ1dHRvbjphY3RpdmUgeyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMCkgc2NhbGUoMC45Nik7IH0KLmZpbHRlci1wcmVzZXRzIGJ1dHRvbi5pcy1hY3RpdmUgeyBiYWNrZ3JvdW5kOiB2YXIoLS1zZXJpZXMtMSk7IGNvbG9yOiAjZmZmOyBib3JkZXItY29sb3I6IHRyYW5zcGFyZW50OyBib3gtc2hhZG93OiAwIDRweCAxNHB4IC01cHggY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSA2MCUsIHRyYW5zcGFyZW50KTsgfQoKLyogLS0tLS0tLS0tLSBWaWV3IGFyZWEgLS0tLS0tLS0tLSAqLwoudmlldy1hcmVhIHsgZmxleDogMTsgcGFkZGluZzogMjRweDsgbWF4LXdpZHRoOiAxNDAwcHg7IHdpZHRoOiAxMDAlOyBtYXJnaW46IDAgYXV0bzsgfQoudmlldyB7IGRpc3BsYXk6IG5vbmU7IH0KLnZpZXcuaXMtYWN0aXZlIHsgZGlzcGxheTogYmxvY2s7IGFuaW1hdGlvbjogdmlld0ZhZGVJbiAyNjBtcyB2YXIoLS1lYXNlKTsgfQpAa2V5ZnJhbWVzIHZpZXdGYWRlSW4gewogIGZyb20geyBvcGFjaXR5OiAwOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoNnB4KTsgfQogIHRvIHsgb3BhY2l0eTogMTsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDApOyB9Cn0KCi5zZWN0aW9uLXRpdGxlIHsgZm9udC1zaXplOiAxNnB4OyBmb250LXdlaWdodDogNzAwOyBtYXJnaW46IDMycHggMCAxNHB4OyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgbGV0dGVyLXNwYWNpbmc6IC0wLjAxZW07IH0KLnNlY3Rpb24tdGl0bGU6Zmlyc3QtY2hpbGQgeyBtYXJnaW4tdG9wOiAwOyB9CgovKiAtLS0tLS0tLS0tIElucHV0cyDigJQgb25lIHNoYXJlZCBnbGFzcyB0cmVhdG1lbnQgZm9yIGV2ZXJ5IHRleHQgaW5wdXQsIHNlbGVjdCwgYW5kIGRhdGUgcGlja2VyIC0tLS0tLS0tLS0gKi8KLmZpbHRlci1maWVsZCBzZWxlY3QsIC5maWx0ZXItZmllbGQgaW5wdXRbdHlwZT0iZGF0ZSJdLAouZm9ybS1maWVsZCBpbnB1dCwgLmZvcm0tZmllbGQgc2VsZWN0LCAuZm9ybS1maWVsZCB0ZXh0YXJlYSwKLmRhc2hib2FyZC1jb250cm9scyBzZWxlY3QsIC5yZWNvcmRzLXNlYXJjaCBpbnB1dCwKLmZpZWxkLWlubGluZSBzZWxlY3QsIC5maWVsZC1pbmxpbmUgaW5wdXQsCi5jb25mbGljdC1yb3cgc2VsZWN0LCAuY2FyZC1oZWFkZXIgc2VsZWN0IHsKICBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtc20pOwogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMik7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyBmb250LXNpemU6IDEzcHg7CiAgcGFkZGluZzogOHB4IDEycHg7IG1pbi13aWR0aDogMTQwcHg7CiAgdHJhbnNpdGlvbjogYm9yZGVyLWNvbG9yIDE2MG1zIHZhcigtLWVhc2UpLCBib3gtc2hhZG93IDE2MG1zIHZhcigtLWVhc2UpOwp9Ci5maWx0ZXItZmllbGQgc2VsZWN0OmhvdmVyLCAuZmlsdGVyLWZpZWxkIGlucHV0W3R5cGU9ImRhdGUiXTpob3ZlciwKLmZvcm0tZmllbGQgaW5wdXQ6aG92ZXIsIC5mb3JtLWZpZWxkIHNlbGVjdDpob3ZlciwKLmRhc2hib2FyZC1jb250cm9scyBzZWxlY3Q6aG92ZXIsIC5yZWNvcmRzLXNlYXJjaCBpbnB1dDpob3ZlciwKLmZpZWxkLWlubGluZSBzZWxlY3Q6aG92ZXIsIC5maWVsZC1pbmxpbmUgaW5wdXQ6aG92ZXIsCi5jb25mbGljdC1yb3cgc2VsZWN0OmhvdmVyLCAuY2FyZC1oZWFkZXIgc2VsZWN0OmhvdmVyIHsKICBib3JkZXItY29sb3I6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zZXJpZXMtMSkgMzUlLCB2YXIoLS1ib3JkZXIpKTsKfQouZmlsdGVyLWZpZWxkIHNlbGVjdDpmb2N1cywgLmZpbHRlci1maWVsZCBpbnB1dFt0eXBlPSJkYXRlIl06Zm9jdXMsCi5mb3JtLWZpZWxkIGlucHV0OmZvY3VzLCAuZm9ybS1maWVsZCBzZWxlY3Q6Zm9jdXMsIC5mb3JtLWZpZWxkIHRleHRhcmVhOmZvY3VzLAouZGFzaGJvYXJkLWNvbnRyb2xzIHNlbGVjdDpmb2N1cywgLnJlY29yZHMtc2VhcmNoIGlucHV0OmZvY3VzLAouZmllbGQtaW5saW5lIHNlbGVjdDpmb2N1cywgLmZpZWxkLWlubGluZSBpbnB1dDpmb2N1cywKLmNvbmZsaWN0LXJvdyBzZWxlY3Q6Zm9jdXMsIC5jYXJkLWhlYWRlciBzZWxlY3Q6Zm9jdXMsCi5hdXRoLWZvcm0gaW5wdXQ6Zm9jdXMgewogIG91dGxpbmU6IG5vbmU7IGJvcmRlci1jb2xvcjogdmFyKC0tc2VyaWVzLTEpOwogIGJveC1zaGFkb3c6IDAgMCAwIDNweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDE4JSwgdHJhbnNwYXJlbnQpOwp9CgovKiAtLS0tLS0tLS0tIFN0YXQgdGlsZXMgLS0tLS0tLS0tLSAqLwouc3RhdC1ncmlkIHsKICBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdChhdXRvLWZpdCwgbWlubWF4KDE4MHB4LCAxZnIpKTsgZ2FwOiAxNHB4Owp9Ci5zdGF0LXRpbGUgewogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1tZCk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHBhZGRpbmc6IDE2cHggMThweDsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWNhcmQpOwogIHRyYW5zaXRpb246IHRyYW5zZm9ybSAyMDBtcyB2YXIoLS1lYXNlKSwgYm94LXNoYWRvdyAyMDBtcyB2YXIoLS1lYXNlKTsKICBhbmltYXRpb246IGNhcmRJbiAzMjBtcyB2YXIoLS1lYXNlKSBiYWNrd2FyZHM7Cn0KLnN0YXQtdGlsZTpob3ZlciB7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgtM3B4KTsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWhvdmVyKTsgfQouc3RhdC1sYWJlbCB7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgZm9udC13ZWlnaHQ6IDYwMDsgfQouc3RhdC12YWx1ZSB7IGZvbnQtc2l6ZTogMjdweDsgZm9udC13ZWlnaHQ6IDcwMDsgbWFyZ2luLXRvcDogNXB4OyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgbGV0dGVyLXNwYWNpbmc6IC0wLjAyZW07IH0KLnN0YXQtZGVsdGEgeyBmb250LXNpemU6IDEycHg7IG1hcmdpbi10b3A6IDdweDsgZm9udC13ZWlnaHQ6IDYwMDsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA0cHg7IH0KLnN0YXQtZGVsdGEudXAgeyBjb2xvcjogdmFyKC0tc3VjY2Vzcy10ZXh0KTsgfQouc3RhdC1kZWx0YS5kb3duIHsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IH0KLnN0YXQtZGVsdGEuZmxhdCB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgfQouc3RhdC1kZWx0YS51cDo6YmVmb3JlIHsgY29udGVudDogJ+KGkSc7IH0KLnN0YXQtZGVsdGEuZG93bjo6YmVmb3JlIHsgY29udGVudDogJ+KGkyc7IH0KCi5pbnNpZ2h0cy1saXN0IHsgbGlzdC1zdHlsZTogbm9uZTsgbWFyZ2luOiAwOyBwYWRkaW5nOiAwOyBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBnYXA6IDEwcHg7IH0KLmluc2lnaHRzLWxpc3QgbGkgewogIGZvbnQtc2l6ZTogMTNweDsgbGluZS1oZWlnaHQ6IDEuNTsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgcGFkZGluZy1sZWZ0OiAxOHB4OyBwb3NpdGlvbjogcmVsYXRpdmU7Cn0KLmluc2lnaHRzLWxpc3QgbGk6OmJlZm9yZSB7CiAgY29udGVudDogJ+Kcpic7IHBvc2l0aW9uOiBhYnNvbHV0ZTsgbGVmdDogMDsgY29sb3I6IHZhcigtLXNlcmllcy0xKTsgZm9udC1zaXplOiAxMXB4OyB0b3A6IDJweDsKfQoKQGtleWZyYW1lcyBjYXJkSW4gewogIGZyb20geyBvcGFjaXR5OiAwOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMTBweCk7IH0KICB0byB7IG9wYWNpdHk6IDE7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgwKTsgfQp9CgovKiAtLS0tLS0tLS0tIENhcmRzIC8gY2hhcnRzIC0tLS0tLS0tLS0gKi8KLmNhcmQtZ3JpZCB7IGRpc3BsYXk6IGdyaWQ7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogMmZyIDFmcjsgZ2FwOiAxNnB4OyBhbGlnbi1pdGVtczogc3RhcnQ7IH0KLmNhcmQtZ3JpZC5ldmVuIHsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiAxZnIgMWZyOyB9CkBtZWRpYSAobWF4LXdpZHRoOiA5MDBweCkgeyAuY2FyZC1ncmlkLCAuY2FyZC1ncmlkLmV2ZW4geyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IDFmcjsgfSB9Ci5jYXJkIHsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTEpOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtbGcpOwogIGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBwYWRkaW5nOiAxOHB4OyBib3gtc2hhZG93OiB2YXIoLS1zaGFkb3ctY2FyZCk7CiAgdHJhbnNpdGlvbjogYm94LXNoYWRvdyAyMjBtcyB2YXIoLS1lYXNlKSwgdHJhbnNmb3JtIDIyMG1zIHZhcigtLWVhc2UpOwogIGFuaW1hdGlvbjogY2FyZEluIDMyMG1zIHZhcigtLWVhc2UpIGJhY2t3YXJkczsKfQouY2FyZDpob3ZlciB7IGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdy1ob3Zlcik7IH0KLmNhcmQtaGVhZGVyIHsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsganVzdGlmeS1jb250ZW50OiBzcGFjZS1iZXR3ZWVuOyBnYXA6IDhweDsgbWFyZ2luLWJvdHRvbTogMTRweDsgfQouY2FyZC1oZWFkZXIgaDMgeyBmb250LXNpemU6IDE0cHg7IG1hcmdpbjogMDsgZm9udC13ZWlnaHQ6IDcwMDsgbGV0dGVyLXNwYWNpbmc6IC0wLjAwNWVtOyB9Ci5jYXJkLWhlYWRlciBzZWxlY3QgeyBmb250LXNpemU6IDEycHg7IHBhZGRpbmc6IDZweCAxMHB4OyBtaW4td2lkdGg6IDA7IH0KLmNoYXJ0LXdyYXAgeyBwb3NpdGlvbjogcmVsYXRpdmU7IGhlaWdodDogMjgwcHg7IH0KLmNoYXJ0LXdyYXAudGFsbCB7IGhlaWdodDogMzQwcHg7IH0KCi5sZWdlbmQtcm93IHsgZGlzcGxheTogZmxleDsgZmxleC13cmFwOiB3cmFwOyBnYXA6IDEycHg7IG1hcmdpbi10b3A6IDEwcHg7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgfQoubGVnZW5kLWl0ZW0geyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDZweDsgfQoubGVnZW5kLXN3YXRjaCB7IHdpZHRoOiAxMHB4OyBoZWlnaHQ6IDEwcHg7IGJvcmRlci1yYWRpdXM6IDNweDsgZGlzcGxheTogaW5saW5lLWJsb2NrOyB9Ci5sZWdlbmQtbGluZSB7IHdpZHRoOiAxNHB4OyBoZWlnaHQ6IDJweDsgYm9yZGVyLXJhZGl1czogMnB4OyBkaXNwbGF5OiBpbmxpbmUtYmxvY2s7IH0KCi8qIC0tLS0tLS0tLS0gVGFibGVzIOKAlCBwcmVtaXVtIGRhdGFiYXNlIGZlZWwsIG5vdCBhIHNwcmVhZHNoZWV0IC0tLS0tLS0tLS0gKi8KLnRhYmxlLXNjcm9sbCB7CiAgb3ZlcmZsb3cteDogYXV0bzsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLW1kKTsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTIpOwp9Ci5kYXRhLXRhYmxlIHsgd2lkdGg6IDEwMCU7IGJvcmRlci1jb2xsYXBzZTogc2VwYXJhdGU7IGJvcmRlci1zcGFjaW5nOiAwOyBmb250LXNpemU6IDEzcHg7IH0KLmRhdGEtdGFibGUgdGgsIC5kYXRhLXRhYmxlIHRkIHsgdGV4dC1hbGlnbjogbGVmdDsgcGFkZGluZzogMTFweCAxNHB4OyBib3JkZXItYm90dG9tOiAxcHggc29saWQgdmFyKC0tZ3JpZGxpbmUpOyB3aGl0ZS1zcGFjZTogbm93cmFwOyB9Ci5kYXRhLXRhYmxlIHRkLndyYXAgeyB3aGl0ZS1zcGFjZTogbm9ybWFsOyB9Ci5kYXRhLXRhYmxlIHRoZWFkIHRoIHsKICBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyBmb250LXdlaWdodDogNjAwOyBmb250LXNpemU6IDExcHg7IHRleHQtdHJhbnNmb3JtOiB1cHBlcmNhc2U7IGxldHRlci1zcGFjaW5nOiAwLjA0ZW07CiAgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHBvc2l0aW9uOiBzdGlja3k7IHRvcDogMDsgei1pbmRleDogMTsKfQouZGF0YS10YWJsZSB0aGVhZCB0aC5zb3J0YWJsZS10aCB7IGN1cnNvcjogcG9pbnRlcjsgdXNlci1zZWxlY3Q6IG5vbmU7IHRyYW5zaXRpb246IGNvbG9yIDE1MG1zIHZhcigtLWVhc2UpOyB9Ci5kYXRhLXRhYmxlIHRoZWFkIHRoLnNvcnRhYmxlLXRoOmhvdmVyIHsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IH0KLmRhdGEtdGFibGUgdGhlYWQgdGggLnNvcnQtYXJyb3cgeyBjb2xvcjogdmFyKC0tdGV4dC1tdXRlZCk7IGZvbnQtc2l6ZTogMTBweDsgbWFyZ2luLWxlZnQ6IDJweDsgfQouZGF0YS10YWJsZSB0aGVhZCB0aC5zb3J0YWJsZS10aDpob3ZlciAuc29ydC1hcnJvdyB7IGNvbG9yOiB2YXIoLS1zZXJpZXMtMSk7IH0KLmRhdGEtdGFibGUgdGJvZHkgdHI6bnRoLWNoaWxkKGV2ZW4pIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXRleHQtbXV0ZWQpIDQlLCB0cmFuc3BhcmVudCk7IH0KLmRhdGEtdGFibGUgdGQubnVtIHsgZm9udC12YXJpYW50LW51bWVyaWM6IHRhYnVsYXItbnVtczsgdGV4dC1hbGlnbjogcmlnaHQ7IH0KLmRhdGEtdGFibGUgdGgubnVtIHsgdGV4dC1hbGlnbjogcmlnaHQ7IH0KLmRhdGEtdGFibGUgdGJvZHkgdHIgeyB0cmFuc2l0aW9uOiBiYWNrZ3JvdW5kIDE1MG1zIHZhcigtLWVhc2UpOyB9Ci5kYXRhLXRhYmxlIHRib2R5IHRyOmhvdmVyIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSA3JSwgdHJhbnNwYXJlbnQpOyB9Ci5kYXRhLXRhYmxlIHRib2R5IHRyOmxhc3QtY2hpbGQgdGQgeyBib3JkZXItYm90dG9tOiBub25lOyB9Ci5wbGF0Zm9ybS1waWxsIHsKICBkaXNwbGF5OiBpbmxpbmUtZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA2cHg7IGZvbnQtc2l6ZTogMTJweDsgZm9udC13ZWlnaHQ6IDYwMDsKICBwYWRkaW5nOiA0cHggMTBweDsgYm9yZGVyLXJhZGl1czogMjBweDsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsKfQoucGxhdGZvcm0tZG90IHsgd2lkdGg6IDhweDsgaGVpZ2h0OiA4cHg7IGJvcmRlci1yYWRpdXM6IDUwJTsgfQoKLyogLS0tLS0tLS0tLSBCdXR0b25zIOKAlCBuZXZlciBmbGF0OiBzb2Z0IHNoYWRvdywgaG92ZXIgbGlmdCwgcHJlc3Mgc2NhbGUgLS0tLS0tLS0tLSAqLwouYnRuIHsKICBkaXNwbGF5OiBpbmxpbmUtZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsganVzdGlmeS1jb250ZW50OiBjZW50ZXI7IGdhcDogNnB4OwogIGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMik7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOwogIGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBwYWRkaW5nOiA5cHggMTdweDsgYm9yZGVyLXJhZGl1czogMTFweDsgY3Vyc29yOiBwb2ludGVyOyBmb250LXNpemU6IDEzcHg7IGZvbnQtd2VpZ2h0OiA2MDA7CiAgYm94LXNoYWRvdzogMCAxcHggMnB4IHJnYmEoMTUsMTcsMjEsMC4wNCk7CiAgdHJhbnNpdGlvbjogdHJhbnNmb3JtIDE1MG1zIHZhcigtLWVhc2UpLCBib3gtc2hhZG93IDE1MG1zIHZhcigtLWVhc2UpLCBmaWx0ZXIgMTUwbXMgdmFyKC0tZWFzZSksIGJhY2tncm91bmQgMTUwbXMgdmFyKC0tZWFzZSk7Cn0KLmJ0biBzdmcgeyBmbGV4LXNocmluazogMDsgfQouYnRuOmhvdmVyIHsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKC0xcHgpOyBib3gtc2hhZG93OiB2YXIoLS1zaGFkb3ctaG92ZXIpOyBmaWx0ZXI6IGJyaWdodG5lc3MoMS4wMik7IH0KLmJ0bjphY3RpdmUgeyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoMCkgc2NhbGUoMC45Nik7IGJveC1zaGFkb3c6IDAgMXB4IDJweCByZ2JhKDE1LDE3LDIxLDAuMDYpOyB9Ci5idG4ucHJpbWFyeSB7CiAgYmFja2dyb3VuZDogdmFyKC0tc2VyaWVzLTEpOyBjb2xvcjogI2ZmZjsgYm9yZGVyLWNvbG9yOiB0cmFuc3BhcmVudDsKICBib3gtc2hhZG93OiAwIDRweCAxNHB4IC01cHggY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSA2NSUsIHRyYW5zcGFyZW50KTsKfQouYnRuLnByaW1hcnk6aG92ZXIgeyBmaWx0ZXI6IGJyaWdodG5lc3MoMS4wNyk7IGJveC1zaGFkb3c6IDAgOHB4IDIycHggLTZweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDcwJSwgdHJhbnNwYXJlbnQpOyB9Ci5idG4uZGFuZ2VyIHsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdGF0dXMtY3JpdGljYWwpOyBjb2xvcjogI2ZmZjsgYm9yZGVyLWNvbG9yOiB0cmFuc3BhcmVudDsKICBib3gtc2hhZG93OiAwIDRweCAxNHB4IC01cHggY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXN0YXR1cy1jcml0aWNhbCkgNTUlLCB0cmFuc3BhcmVudCk7Cn0KLmJ0bi5kYW5nZXI6aG92ZXIgeyBmaWx0ZXI6IGJyaWdodG5lc3MoMS4wNik7IGJveC1zaGFkb3c6IDAgOHB4IDIycHggLTZweCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc3RhdHVzLWNyaXRpY2FsKSA2MCUsIHRyYW5zcGFyZW50KTsgfQouYnRuLnN1Y2Nlc3MgewogIGJhY2tncm91bmQ6IHZhcigtLXN0YXR1cy1nb29kKTsgY29sb3I6ICNmZmY7IGJvcmRlci1jb2xvcjogdHJhbnNwYXJlbnQ7CiAgYm94LXNoYWRvdzogMCA0cHggMTRweCAtNXB4IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zdGF0dXMtZ29vZCkgNTUlLCB0cmFuc3BhcmVudCk7Cn0KLmJ0bi5zdWNjZXNzOmhvdmVyIHsgZmlsdGVyOiBicmlnaHRuZXNzKDEuMDYpOyBib3gtc2hhZG93OiAwIDhweCAyMnB4IC02cHggY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXN0YXR1cy1nb29kKSA2MCUsIHRyYW5zcGFyZW50KTsgfQouYnRuOmRpc2FibGVkIHsgb3BhY2l0eTogMC40NTsgY3Vyc29yOiBub3QtYWxsb3dlZDsgdHJhbnNmb3JtOiBub25lOyBib3gtc2hhZG93OiBub25lOyBmaWx0ZXI6IG5vbmU7IH0KLmJ0bi1yb3cgeyBkaXNwbGF5OiBmbGV4OyBnYXA6IDhweDsgZmxleC13cmFwOiB3cmFwOyB9CgovKiAtLS0tLS0tLS0tIFVwbG9hZCAtLS0tLS0tLS0tICovCi5kcm9wem9uZSB7CiAgYm9yZGVyOiAycHggZGFzaGVkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1sZyk7IHBhZGRpbmc6IDQwcHggMjBweDsKICB0ZXh0LWFsaWduOiBjZW50ZXI7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBjdXJzb3I6IHBvaW50ZXI7IHRyYW5zaXRpb246IGJvcmRlci1jb2xvciAyMDBtcyB2YXIoLS1lYXNlKSwgYmFja2dyb3VuZCAyMDBtcyB2YXIoLS1lYXNlKSwgdHJhbnNmb3JtIDIwMG1zIHZhcigtLWVhc2UpOwp9Ci5kcm9wem9uZTpob3ZlciB7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgtMXB4KTsgfQouZHJvcHpvbmUuaXMtZHJhZyB7IGJvcmRlci1jb2xvcjogdmFyKC0tc2VyaWVzLTEpOyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc2VyaWVzLTEpIDYlLCB2YXIoLS1zdXJmYWNlLTIpKTsgdHJhbnNmb3JtOiBzY2FsZSgxLjAwNSk7IH0KLmRyb3B6b25lIGgzIHsgbWFyZ2luOiAwIDAgNnB4OyBmb250LXNpemU6IDE1cHg7IH0KLmRyb3B6b25lIHAgeyBtYXJnaW46IDA7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IGZvbnQtc2l6ZTogMTNweDsgfQouZHJvcHpvbmUgaW5wdXRbdHlwZT0iZmlsZSJdIHsgZGlzcGxheTogbm9uZTsgfQoKLmNvbmZsaWN0LWxpc3QgeyBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBnYXA6IDhweDsgbWFyZ2luOiAxMnB4IDA7IH0KLmNvbmZsaWN0LXJvdyB7CiAgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsganVzdGlmeS1jb250ZW50OiBzcGFjZS1iZXR3ZWVuOyBnYXA6IDEycHg7CiAgcGFkZGluZzogMTFweCAxNHB4OyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtc20pOyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTIpOwogIHRyYW5zaXRpb246IGJveC1zaGFkb3cgMTgwbXMgdmFyKC0tZWFzZSk7Cn0KLmNvbmZsaWN0LXJvdzpob3ZlciB7IGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdy1jYXJkKTsgfQouY29uZmxpY3Qtcm93IC53ZWVrLWxhYmVsIHsgZm9udC13ZWlnaHQ6IDYwMDsgZm9udC1zaXplOiAxM3B4OyB9Ci5jb25mbGljdC1yb3cgLndlZWstbWV0YSB7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgfQouY29uZmxpY3Qtcm93IHNlbGVjdCB7IG1pbi13aWR0aDogMDsgfQoKLmJhZGdlIHsgZGlzcGxheTogaW5saW5lLWJsb2NrOyBwYWRkaW5nOiAzcHggMTBweDsgYm9yZGVyLXJhZGl1czogMjBweDsgZm9udC1zaXplOiAxMXB4OyBmb250LXdlaWdodDogNzAwOyB9Ci5iYWRnZS5zdWNjZXNzIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXN0YXR1cy1nb29kKSAxOCUsIHRyYW5zcGFyZW50KTsgY29sb3I6IHZhcigtLXN0YXR1cy1nb29kKTsgfQouYmFkZ2UucGFydGlhbCB7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zdGF0dXMtd2FybmluZykgMjUlLCB0cmFuc3BhcmVudCk7IGNvbG9yOiAjOGE2MzAwOyB9Ci5iYWRnZS5mYWlsZWQgeyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc3RhdHVzLWNyaXRpY2FsKSAxOCUsIHRyYW5zcGFyZW50KTsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IH0KLmJhZGdlLmVycm9yLXNldiB7IGNvbG9yOiB2YXIoLS1zdGF0dXMtY3JpdGljYWwpOyB9Ci5iYWRnZS53YXJuaW5nLXNldiB7IGNvbG9yOiAjOGE2MzAwOyB9Ci5iYWRnZS5za2lwLXNldiB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgfQoKLmlzc3Vlcy1saXN0IHsgbWF4LWhlaWdodDogMjIwcHg7IG92ZXJmbG93LXk6IGF1dG87IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1zbSk7IH0KLmlzc3VlLXJvdyB7IHBhZGRpbmc6IDlweCAxNHB4OyBib3JkZXItYm90dG9tOiAxcHggc29saWQgdmFyKC0tZ3JpZGxpbmUpOyBmb250LXNpemU6IDEycHg7IH0KLmlzc3VlLXJvdzpsYXN0LWNoaWxkIHsgYm9yZGVyLWJvdHRvbTogbm9uZTsgfQouaXNzdWUtcm93IC5yb3ctbm8geyBjb2xvcjogdmFyKC0tdGV4dC1tdXRlZCk7IG1hcmdpbi1yaWdodDogNnB4OyB9CgovKiAtLS0tLS0tLS0tIFRvYXN0IC0tLS0tLS0tLS0gKi8KLnRvYXN0LXJvb3QgeyBwb3NpdGlvbjogZml4ZWQ7IGJvdHRvbTogMjBweDsgcmlnaHQ6IDIwcHg7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGdhcDogOHB4OyB6LWluZGV4OiAxMDA7IH0KLnRvYXN0IHsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTEpOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtc20pOwogIGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBwYWRkaW5nOiAxMnB4IDE2cHg7IGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdy1tb2RhbCk7IGZvbnQtc2l6ZTogMTNweDsgbWF4LXdpZHRoOiAzNDBweDsKICBhbmltYXRpb246IHRvYXN0LWluIDIyMG1zIHZhcigtLWVhc2UpOwp9Ci50b2FzdC5zdWNjZXNzIHsgYm9yZGVyLWxlZnQ6IDNweCBzb2xpZCB2YXIoLS1zdGF0dXMtZ29vZCk7IH0KLnRvYXN0LmVycm9yIHsgYm9yZGVyLWxlZnQ6IDNweCBzb2xpZCB2YXIoLS1zdGF0dXMtY3JpdGljYWwpOyB9CkBrZXlmcmFtZXMgdG9hc3QtaW4geyBmcm9tIHsgb3BhY2l0eTogMDsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDEwcHgpIHNjYWxlKDAuOTgpOyB9IHRvIHsgb3BhY2l0eTogMTsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDApIHNjYWxlKDEpOyB9IH0KCi8qIC0tLS0tLS0tLS0gTWlzYyAtLS0tLS0tLS0tICovCi5tdXRlZCB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgfQouZW1wdHktc3RhdGUgewogIHBhZGRpbmc6IDU2cHggMjRweDsgdGV4dC1hbGlnbjogY2VudGVyOyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOwogIGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogMTJweDsKICBhbmltYXRpb246IGNhcmRJbiAyNjBtcyB2YXIoLS1lYXNlKTsKfQouZW1wdHktc3RhdGUgLmVtcHR5LWljb24gewogIHdpZHRoOiA1MnB4OyBoZWlnaHQ6IDUycHg7IGJvcmRlci1yYWRpdXM6IDE2cHg7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogY2VudGVyOwogIGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zZXJpZXMtMSkgMTAlLCB0cmFuc3BhcmVudCk7IGNvbG9yOiB2YXIoLS1zZXJpZXMtMSk7Cn0KLmVtcHR5LXN0YXRlIC5lbXB0eS10aXRsZSB7IGZvbnQtc2l6ZTogMTRweDsgZm9udC13ZWlnaHQ6IDYwMDsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IH0KLmVtcHR5LXN0YXRlIC5lbXB0eS1tZXNzYWdlIHsgZm9udC1zaXplOiAxM3B4OyBtYXgtd2lkdGg6IDM2MHB4OyB9Ci5zcGlubmVyIHsgd2lkdGg6IDE2cHg7IGhlaWdodDogMTZweDsgYm9yZGVyLXJhZGl1czogNTAlOyBib3JkZXI6IDJweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItdG9wLWNvbG9yOiB2YXIoLS1zZXJpZXMtMSk7IGFuaW1hdGlvbjogc3BpbiAuNnMgbGluZWFyIGluZmluaXRlOyBkaXNwbGF5OiBpbmxpbmUtYmxvY2s7IH0KQGtleWZyYW1lcyBzcGluIHsgdG8geyB0cmFuc2Zvcm06IHJvdGF0ZSgzNjBkZWcpOyB9IH0KLmxvYWRpbmctcm93IHsgcGFkZGluZzogNDBweCAyMHB4OyB0ZXh0LWFsaWduOiBjZW50ZXI7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IH0KCi8qIFNrZWxldG9uIGxvYWRlcnMg4oCUIHNoaW1tZXJpbmcgcGxhY2Vob2xkZXJzIHNob3duIHdoaWxlIGEgc2VjdGlvbidzIGRhdGEgaXMgaW4gZmxpZ2h0ICovCi5za2VsZXRvbiB7CiAgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLXNtKTsKICBiYWNrZ3JvdW5kOiBsaW5lYXItZ3JhZGllbnQoMTAwZGVnLCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tdGV4dC1tdXRlZCkgMTIlLCB0cmFuc3BhcmVudCkgMzAlLCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tdGV4dC1tdXRlZCkgMjIlLCB0cmFuc3BhcmVudCkgNTAlLCBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tdGV4dC1tdXRlZCkgMTIlLCB0cmFuc3BhcmVudCkgNzAlKTsKICBiYWNrZ3JvdW5kLXNpemU6IDIwMCUgMTAwJTsKICBhbmltYXRpb246IHNrZWxldG9uU2hpbW1lciAxLjRzIGVhc2UtaW4tb3V0IGluZmluaXRlOwp9CkBrZXlmcmFtZXMgc2tlbGV0b25TaGltbWVyIHsgZnJvbSB7IGJhY2tncm91bmQtcG9zaXRpb246IDE1MCUgMDsgfSB0byB7IGJhY2tncm91bmQtcG9zaXRpb246IC01MCUgMDsgfSB9Ci5za2VsZXRvbi1zdGF0LWdyaWQgeyBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdChhdXRvLWZpdCwgbWlubWF4KDE4MHB4LCAxZnIpKTsgZ2FwOiAxNHB4OyB9Ci5za2VsZXRvbi10aWxlIHsgaGVpZ2h0OiA4NHB4OyB9Ci5za2VsZXRvbi1jaGFydCB7IGhlaWdodDogMjgwcHg7IHdpZHRoOiAxMDAlOyB9Ci5za2VsZXRvbi1yb3cgeyBoZWlnaHQ6IDQwcHg7IG1hcmdpbi1ib3R0b206IDhweDsgfQoKLyogQW5pbWF0ZWQgaG9yaXpvbnRhbCBjb21wYXJpc29uIGJhciDigJQgYSBsYWJlbGVkIHJvdyB3aXRoIGEgdHJhY2sgdGhhdCBmaWxscyBpbiBvbiBpbnNlcnRpb24gKi8KLmJhci1yb3cgeyBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IG1pbm1heCg5MHB4LCAxNDBweCkgMWZyIGF1dG87IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogMTBweDsgcGFkZGluZzogNXB4IDA7IH0KLmJhci1sYWJlbCB7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgZm9udC13ZWlnaHQ6IDYwMDsgfQouYmFyLXRyYWNrIHsgaGVpZ2h0OiA4cHg7IGJvcmRlci1yYWRpdXM6IDVweDsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXRleHQtbXV0ZWQpIDE0JSwgdHJhbnNwYXJlbnQpOyBvdmVyZmxvdzogaGlkZGVuOyB9Ci5iYXItZmlsbCB7IGhlaWdodDogMTAwJTsgd2lkdGg6IDAlOyBib3JkZXItcmFkaXVzOiA1cHg7IHRyYW5zaXRpb246IHdpZHRoIDcwMG1zIGN1YmljLWJlemllcigwLjE2LCAxLCAwLjMsIDEpOyB9Ci5iYXItdmFsdWUgeyBmb250LXNpemU6IDEycHg7IGZvbnQtd2VpZ2h0OiA3MDA7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyBmb250LXZhcmlhbnQtbnVtZXJpYzogdGFidWxhci1udW1zOyB0ZXh0LWFsaWduOiByaWdodDsgbWluLXdpZHRoOiA1NnB4OyB9CgpAbWVkaWEgKHByZWZlcnMtcmVkdWNlZC1tb3Rpb246IHJlZHVjZSkgewogIC5iYXItZmlsbCB7IHRyYW5zaXRpb24tZHVyYXRpb246IDFtczsgfQogIC5za2VsZXRvbiB7IGFuaW1hdGlvbi1kdXJhdGlvbjogMW1zOyB9CiAgLmNhcmQsIC5zdGF0LXRpbGUgeyBhbmltYXRpb24tZHVyYXRpb246IDFtczsgfQp9CgoudHdvLWNvbCB7IGRpc3BsYXk6IGdyaWQ7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogMWZyIDFmcjsgZ2FwOiAxNnB4OyB9CkBtZWRpYSAobWF4LXdpZHRoOiA5MDBweCkgeyAudHdvLWNvbCB7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogMWZyOyB9IH0KCi5tb2RlLXRhYnMgeyBkaXNwbGF5OiBmbGV4OyBnYXA6IDZweDsgZmxleC13cmFwOiB3cmFwOyBtYXJnaW4tYm90dG9tOiAxNnB4OyB9Ci5tb2RlLXRhYnMgYnV0dG9uIHsKICBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTEpOyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOwogIGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBwYWRkaW5nOiA3cHggMTRweDsgYm9yZGVyLXJhZGl1czogMjBweDsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNjAwOyBjdXJzb3I6IHBvaW50ZXI7CiAgdHJhbnNpdGlvbjogY29sb3IgMTgwbXMgdmFyKC0tZWFzZSksIGJhY2tncm91bmQgMTgwbXMgdmFyKC0tZWFzZSksIHRyYW5zZm9ybSAxNTBtcyB2YXIoLS1lYXNlKTsKfQoubW9kZS10YWJzIGJ1dHRvbjpob3ZlciB7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgtMXB4KTsgfQoubW9kZS10YWJzIGJ1dHRvbi5pcy1hY3RpdmUgeyBiYWNrZ3JvdW5kOiB2YXIoLS1zZXJpZXMtMSk7IGNvbG9yOiAjZmZmOyBib3JkZXItY29sb3I6IHRyYW5zcGFyZW50OyBib3gtc2hhZG93OiAwIDRweCAxNHB4IC01cHggY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSA2MCUsIHRyYW5zcGFyZW50KTsgfQoKLmZpZWxkLWlubGluZSB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogOHB4OyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IH0KLmZpZWxkLWlubGluZSBzZWxlY3QsIC5maWVsZC1pbmxpbmUgaW5wdXQgeyBtaW4td2lkdGg6IDA7IHBhZGRpbmc6IDZweCAxMHB4OyB9CgovKiAtLS0tLS0tLS0tIFBsYXRmb3JtIFBlcmZvcm1hbmNlIENvbXBhcmlzb24gY2FyZHMgLS0tLS0tLS0tLSAqLwoucGNjLXNlY3Rpb24geyBtYXJnaW4tdG9wOiAyNHB4OyB9Ci5wY2MtY29udHJvbHMgeyBkaXNwbGF5OiBmbGV4OyBmbGV4LXdyYXA6IHdyYXA7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogMTZweDsgbWFyZ2luLWJvdHRvbTogMTZweDsgfQoucGxhdGZvcm0tY29tcGFyZS1ncmlkIHsKICBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdCgyLCAxZnIpOyBnYXA6IDE2cHg7Cn0KQG1lZGlhIChtYXgtd2lkdGg6IDkwMHB4KSB7IC5wbGF0Zm9ybS1jb21wYXJlLWdyaWQgeyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IDFmcjsgfSB9Ci5wbGF0Zm9ybS1jb21wYXJlLWNhcmQgewogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMSk7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1sZyk7CiAgYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOwogIHBhZGRpbmc6IDE4cHg7IGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdy1jYXJkKTsKICB0cmFuc2l0aW9uOiBib3gtc2hhZG93IDIyMG1zIHZhcigtLWVhc2UpLCB0cmFuc2Zvcm0gMjIwbXMgdmFyKC0tZWFzZSk7CiAgYW5pbWF0aW9uOiBjYXJkSW4gMzIwbXMgdmFyKC0tZWFzZSkgYmFja3dhcmRzOwp9Ci5wbGF0Zm9ybS1jb21wYXJlLWNhcmQ6aG92ZXIgeyBib3gtc2hhZG93OiB2YXIoLS1zaGFkb3ctaG92ZXIpOyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoLTJweCk7IH0KLnBjYy1oZWFkZXIgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6IHNwYWNlLWJldHdlZW47IGdhcDogMTBweDsgfQoucGNjLWhlYWRlci1uYW1lIHsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA4cHg7IH0KLnBjYy1uYW1lIHsgZm9udC1zaXplOiAxNXB4OyBmb250LXdlaWdodDogNzAwOyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgfQoucGNjLWJhZGdlIHsgZm9udC1zaXplOiAxM3B4OyBmb250LXdlaWdodDogNzAwOyBwYWRkaW5nOiA0cHggMTBweDsgYm9yZGVyLXJhZGl1czogMjBweDsgfQoucGNjLWJhZGdlLnVwIHsgY29sb3I6IHZhcigtLXN1Y2Nlc3MtdGV4dCk7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zdGF0dXMtZ29vZCkgMTQlLCB0cmFuc3BhcmVudCk7IH0KLnBjYy1iYWRnZS5kb3duIHsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1zdGF0dXMtY3JpdGljYWwpIDEyJSwgdHJhbnNwYXJlbnQpOyB9Ci5wY2MtYmFkZ2UuZmxhdCB7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXRleHQtbXV0ZWQpIDEyJSwgdHJhbnNwYXJlbnQpOyB9Ci5wY2MtY2FwdGlvbiB7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgbWFyZ2luLXRvcDogNnB4OyB9Ci5wY2MtbWV0cmljcyB7IG1hcmdpbi10b3A6IDE2cHg7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGdhcDogMTRweDsgfQoucGNjLW1ldHJpYy1yb3cgeyBwYWRkaW5nLXRvcDogMTJweDsgYm9yZGVyLXRvcDogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IH0KLnBjYy1tZXRyaWMtcm93OmZpcnN0LWNoaWxkIHsgcGFkZGluZy10b3A6IDA7IGJvcmRlci10b3A6IG5vbmU7IH0KLnBjYy1tZXRyaWMtaGVhZGVyIHsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsganVzdGlmeS1jb250ZW50OiBzcGFjZS1iZXR3ZWVuOyBnYXA6IDhweDsgbWFyZ2luLWJvdHRvbTogNnB4OyB9Ci5wY2MtbWV0cmljLWxhYmVsIHsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNzAwOyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgfQoucGNjLW1ldHJpYy1kaWZmIHsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNzAwOyB9Ci5wY2MtbWV0cmljLWRpZmYudXAgeyBjb2xvcjogdmFyKC0tc3VjY2Vzcy10ZXh0KTsgfQoucGNjLW1ldHJpYy1kaWZmLmRvd24geyBjb2xvcjogdmFyKC0tc3RhdHVzLWNyaXRpY2FsKTsgfQoucGNjLW1ldHJpYy1kaWZmLmZsYXQgeyBjb2xvcjogdmFyKC0tdGV4dC1tdXRlZCk7IH0KLnBjYy1mb290ZXIgeyBtYXJnaW4tdG9wOiAxNnB4OyBwYWRkaW5nLXRvcDogMTRweDsgYm9yZGVyLXRvcDogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IH0KLnBjYy1mb290ZXItbGFiZWwgeyBmb250LXNpemU6IDExcHg7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgZm9udC13ZWlnaHQ6IDYwMDsgdGV4dC10cmFuc2Zvcm06IHVwcGVyY2FzZTsgbGV0dGVyLXNwYWNpbmc6IDAuMDNlbTsgfQoucGNjLWZvb3Rlci12YWx1ZSB7IGZvbnQtc2l6ZTogMTVweDsgZm9udC13ZWlnaHQ6IDcwMDsgbWFyZ2luLXRvcDogNHB4OyB9Ci5wY2MtZm9vdGVyLXZhbHVlLnVwIHsgY29sb3I6IHZhcigtLXN1Y2Nlc3MtdGV4dCk7IH0KLnBjYy1mb290ZXItdmFsdWUuZG93biB7IGNvbG9yOiB2YXIoLS1zdGF0dXMtY3JpdGljYWwpOyB9Ci5wY2MtZm9vdGVyLXZhbHVlLmZsYXQgeyBjb2xvcjogdmFyKC0tdGV4dC1tdXRlZCk7IH0KLnBjYy1mb290ZXItZGV0YWlsIHsgZm9udC1zaXplOiAxMnB4OyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyBtYXJnaW4tdG9wOiA0cHg7IH0KLnBjYy12aWV3LWxpbmsgewogIGRpc3BsYXk6IGlubGluZS1mbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDRweDsgbWFyZ2luLXRvcDogMTRweDsKICBiYWNrZ3JvdW5kOiBub25lOyBib3JkZXI6IG5vbmU7IGNvbG9yOiB2YXIoLS1zZXJpZXMtMSk7IGZvbnQtc2l6ZTogMTJweDsgZm9udC13ZWlnaHQ6IDcwMDsgY3Vyc29yOiBwb2ludGVyOyBwYWRkaW5nOiAwOwogIHRyYW5zaXRpb246IG9wYWNpdHkgMTUwbXMgdmFyKC0tZWFzZSk7Cn0KLnBjYy12aWV3LWxpbms6aG92ZXIgeyBvcGFjaXR5OiAwLjc1OyB0ZXh0LWRlY29yYXRpb246IHVuZGVybGluZTsgfQoKLyogLS0tLS0tLS0tLSBQYWdpbmF0aW9uIC0tLS0tLS0tLS0gKi8KLnBhZ2luYXRpb24tcm93IHsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiAxMnB4OyBtYXJnaW4tdG9wOiAxNHB4OyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0LXNlY29uZGFyeSk7IH0KLnBhZ2luYXRpb24tcm93IC5idG4geyBwYWRkaW5nOiA2cHggMTJweDsgfQoKLyogLS0tLS0tLS0tLSBEYXNoYm9hcmQgY29udHJvbHMgLyBtZXRyaWMtZm9jdXNlZCBLUElzIC0tLS0tLS0tLS0gKi8KLmRhc2hib2FyZC1jb250cm9scyB7CiAgZGlzcGxheTogZmxleDsgZmxleC13cmFwOiB3cmFwOyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDEwcHg7IG1hcmdpbi1ib3R0b206IDE4cHg7Cn0KLmRhc2hib2FyZC1jb250cm9scyBsYWJlbCB7IGZvbnQtc2l6ZTogMTJweDsgZm9udC13ZWlnaHQ6IDYwMDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgbWFyZ2luLXJpZ2h0OiA2cHg7IH0KLmRhc2hib2FyZC1jb250cm9scyBzZWxlY3QgeyBmb250LXdlaWdodDogNjAwOyB9Ci5zdGF0LXRpbGUucG9zdC10aWxlIC5zdGF0LXZhbHVlIHsgZm9udC1zaXplOiAxNXB4OyBsaW5lLWhlaWdodDogMS4zNTsgbWFyZ2luLXRvcDogNnB4OyB9Ci5zdGF0LXRpbGUucG9zdC10aWxlIC5wb3N0LW1ldGEgeyBmb250LXNpemU6IDExcHg7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgbWFyZ2luLXRvcDogNHB4OyB9Ci5zdGF0LXRpbGUucG9zdC10aWxlIC5wb3N0LW1ldHJpYy12YWx1ZSB7IGZvbnQtc2l6ZTogMjBweDsgZm9udC13ZWlnaHQ6IDcwMDsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IG1hcmdpbi10b3A6IDZweDsgfQouc3RhdC10aWxlLnBvc3QtdGlsZSAucG9zdC10aWxlLWVudHJ5ICsgLnBvc3QtdGlsZS1lbnRyeSB7IG1hcmdpbi10b3A6IDEwcHg7IHBhZGRpbmctdG9wOiAxMHB4OyBib3JkZXItdG9wOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgfQouc3RhdC10aWxlLnBvc3QtdGlsZSAucG9zdC10aWxlLWZvb3RlciB7CiAgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsganVzdGlmeS1jb250ZW50OiBzcGFjZS1iZXR3ZWVuOyBnYXA6IDhweDsKICBtYXJnaW4tdG9wOiAxMHB4OyBwYWRkaW5nLXRvcDogMTBweDsgYm9yZGVyLXRvcDogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7Cn0KLnBvc3QtdGlsZS1mb290ZXItbGFiZWwgeyBmb250LXNpemU6IDExcHg7IGNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgZm9udC13ZWlnaHQ6IDYwMDsgfQoucG9zdC10aWxlLWZvb3Rlci12YWx1ZSB7IGZvbnQtc2l6ZTogMTNweDsgZm9udC13ZWlnaHQ6IDcwMDsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IGZvbnQtdmFyaWFudC1udW1lcmljOiB0YWJ1bGFyLW51bXM7IH0KLnN0YXQtdmFsdWUtbXV0ZWQgeyBmb250LXNpemU6IDE1cHggIWltcG9ydGFudDsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyBmb250LXdlaWdodDogNjAwOyB9Ci5jYXB0aW9uLWxpbmsgeyBjb2xvcjogdmFyKC0tc2VyaWVzLTEpOyB0ZXh0LWRlY29yYXRpb246IG5vbmU7IH0KLmNhcHRpb24tbGluazpob3ZlciB7IHRleHQtZGVjb3JhdGlvbjogdW5kZXJsaW5lOyB9CgovKiAtLS0tLS0tLS0tIFBsYXRmb3JtIFN1bW1hcnkgQmFyIOKAlCBjb21wYWN0IHBlci1wbGF0Zm9ybSBzdHJpcCBhYm92ZSB0aGUgRGFzaGJvYXJkJ3MgS1BJIGdyaWQgLS0tLS0tLS0tLSAqLwoucGxhdGZvcm0tc3VtbWFyeS1iYXIgewogIGRpc3BsYXk6IGZsZXg7IGdhcDogMTJweDsgb3ZlcmZsb3cteDogYXV0bzsgcGFkZGluZy1ib3R0b206IDRweDsgbWFyZ2luLWJvdHRvbTogMjBweDsKfQoucGxhdGZvcm0tc3VtbWFyeS1jYXJkIHsKICBmbGV4LXNocmluazogMDsgbWluLXdpZHRoOiAyMjBweDsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTEpOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtbWQpOwogIGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBwYWRkaW5nOiAxNHB4IDE2cHg7IGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdy1jYXJkKTsKICB0cmFuc2l0aW9uOiB0cmFuc2Zvcm0gMjAwbXMgdmFyKC0tZWFzZSksIGJveC1zaGFkb3cgMjAwbXMgdmFyKC0tZWFzZSk7CiAgYW5pbWF0aW9uOiBjYXJkSW4gMzIwbXMgdmFyKC0tZWFzZSkgYmFja3dhcmRzOwp9Ci5wbGF0Zm9ybS1zdW1tYXJ5LWNhcmQ6aG92ZXIgeyB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoLTJweCk7IGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdy1ob3Zlcik7IH0KLnBsYXRmb3JtLXN1bW1hcnktaGVhZGVyIHsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA3cHg7IG1hcmdpbi1ib3R0b206IDEwcHg7IH0KLnBsYXRmb3JtLXN1bW1hcnktbmFtZSB7IGZvbnQtc2l6ZTogMTNweDsgZm9udC13ZWlnaHQ6IDcwMDsgY29sb3I6IHZhcigtLXRleHQtcHJpbWFyeSk7IH0KLnBsYXRmb3JtLXN1bW1hcnktbWV0cmljcyB7IGRpc3BsYXk6IGdyaWQ7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogcmVwZWF0KDMsIDFmcik7IGdhcDogOHB4IDEycHg7IH0KLnBsYXRmb3JtLXN1bW1hcnktbWV0cmljIHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiAycHg7IH0KLnBzbS1sYWJlbCB7IGZvbnQtc2l6ZTogMTBweDsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyBmb250LXdlaWdodDogNjAwOyB0ZXh0LXRyYW5zZm9ybTogdXBwZXJjYXNlOyBsZXR0ZXItc3BhY2luZzogMC4wM2VtOyB9Ci5wc20tdmFsdWUgeyBmb250LXNpemU6IDEzcHg7IGZvbnQtd2VpZ2h0OiA3MDA7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyBmb250LXZhcmlhbnQtbnVtZXJpYzogdGFidWxhci1udW1zOyB9CkBtZWRpYSAobWF4LXdpZHRoOiA2NDBweCkgewogIC5wbGF0Zm9ybS1zdW1tYXJ5LWNhcmQgeyBtaW4td2lkdGg6IDE4MHB4OyB9CiAgLnBsYXRmb3JtLXN1bW1hcnktbWV0cmljcyB7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogcmVwZWF0KDIsIDFmcik7IH0KfQoKLyogLS0tLS0tLS0tLSBEYXRhIFJlY29yZHMgKHBsYXRmb3JtLWdyb3VwZWQpIC0tLS0tLS0tLS0gKi8KLnJlY29yZHMtdG9vbGJhciB7CiAgZGlzcGxheTogZmxleDsgZmxleC13cmFwOiB3cmFwOyBhbGlnbi1pdGVtczogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6IHNwYWNlLWJldHdlZW47CiAgZ2FwOiAxMnB4OyBtYXJnaW4tYm90dG9tOiAxNHB4Owp9Ci5wbGF0Zm9ybS1maWx0ZXItcGlsbHMgeyBkaXNwbGF5OiBmbGV4OyBmbGV4LXdyYXA6IHdyYXA7IGdhcDogNnB4OyB9Ci5wbGF0Zm9ybS1maWx0ZXItcGlsbHMgYnV0dG9uIHsKICBkaXNwbGF5OiBpbmxpbmUtZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA2cHg7CiAgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0xKTsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsKICBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7CiAgcGFkZGluZzogN3B4IDE0cHg7IGJvcmRlci1yYWRpdXM6IDIwcHg7IGZvbnQtc2l6ZTogMTJweDsgZm9udC13ZWlnaHQ6IDYwMDsgY3Vyc29yOiBwb2ludGVyOwogIHRyYW5zaXRpb246IGNvbG9yIDE4MG1zIHZhcigtLWVhc2UpLCBiYWNrZ3JvdW5kIDE4MG1zIHZhcigtLWVhc2UpLCB0cmFuc2Zvcm0gMTUwbXMgdmFyKC0tZWFzZSksIGJveC1zaGFkb3cgMTgwbXMgdmFyKC0tZWFzZSk7Cn0KLnBsYXRmb3JtLWZpbHRlci1waWxscyBidXR0b246aG92ZXIgeyBjb2xvcjogdmFyKC0tdGV4dC1wcmltYXJ5KTsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKC0xcHgpOyB9Ci5wbGF0Zm9ybS1maWx0ZXItcGlsbHMgYnV0dG9uOmFjdGl2ZSB7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgwKSBzY2FsZSgwLjk2KTsgfQoucGxhdGZvcm0tZmlsdGVyLXBpbGxzIGJ1dHRvbi5pcy1hY3RpdmUgeyBiYWNrZ3JvdW5kOiB2YXIoLS1zZXJpZXMtMSk7IGNvbG9yOiAjZmZmOyBib3JkZXItY29sb3I6IHRyYW5zcGFyZW50OyBib3gtc2hhZG93OiAwIDRweCAxNHB4IC01cHggY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXNlcmllcy0xKSA2MCUsIHRyYW5zcGFyZW50KTsgfQoucGxhdGZvcm0tZmlsdGVyLXBpbGxzIGJ1dHRvbi5pcy1hY3RpdmUgLnBsYXRmb3JtLWRvdCB7IGJveC1zaGFkb3c6IDAgMCAwIDJweCByZ2JhKDI1NSwyNTUsMjU1LDAuNSk7IH0KLnJlY29yZHMtc2VhcmNoIGlucHV0IHsgYm9yZGVyLXJhZGl1czogMjBweDsgbWluLXdpZHRoOiAyMjBweDsgfQouc3RhdHVzLXBpbGwgeyBkaXNwbGF5OiBpbmxpbmUtYmxvY2s7IHBhZGRpbmc6IDNweCAxMHB4OyBib3JkZXItcmFkaXVzOiAyMHB4OyBmb250LXNpemU6IDExcHg7IGZvbnQtd2VpZ2h0OiA3MDA7IH0KLnN0YXR1cy1waWxsLm9yaWdpbmFsIHsgYmFja2dyb3VuZDogY29sb3ItbWl4KGluIHNyZ2IsIHZhcigtLXRleHQtbXV0ZWQpIDE1JSwgdHJhbnNwYXJlbnQpOyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyB9Ci5zdGF0dXMtcGlsbC5lZGl0ZWQgeyBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tc3RhdHVzLXdhcm5pbmcpIDIyJSwgdHJhbnNwYXJlbnQpOyBjb2xvcjogIzhhNjMwMDsgfQoucm93LWFjdGlvbnMgeyBkaXNwbGF5OiBmbGV4OyBnYXA6IDZweDsgZmxleC13cmFwOiBub3dyYXA7IH0KLnJvdy1hY3Rpb25zIC5idG4geyBwYWRkaW5nOiA1cHggMTBweDsgZm9udC1zaXplOiAxMnB4OyB9Ci5saW5rLWNlbGwgYSB7IGNvbG9yOiB2YXIoLS1zZXJpZXMtMSk7IHRleHQtZGVjb3JhdGlvbjogbm9uZTsgZm9udC13ZWlnaHQ6IDYwMDsgZm9udC1zaXplOiAxMnB4OyB9Ci5saW5rLWNlbGwgYTpob3ZlciB7IHRleHQtZGVjb3JhdGlvbjogdW5kZXJsaW5lOyB9Ci5yZWNvcmQtc2VjdGlvbiB7CiAgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLXNtKTsgcGFkZGluZzogMTZweDsgbWFyZ2luLWJvdHRvbTogMTRweDsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTIpOyBiYWNrZHJvcC1maWx0ZXI6IHZhcigtLWdsYXNzLWJsdXIpOyAtd2Via2l0LWJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7Cn0KLnJlY29yZC1zZWN0aW9uIGg0IHsgbWFyZ2luOiAwIDAgMTJweDsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNzAwOyBsZXR0ZXItc3BhY2luZzogMC4wM2VtOyB0ZXh0LXRyYW5zZm9ybTogdXBwZXJjYXNlOyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyB9Ci5yZWNvcmQtc2VjdGlvbiAuZm9ybS1ncmlkIHsgbWFyZ2luLWJvdHRvbTogMDsgfQoucmVjb3JkLXNlY3Rpb24gLnZpZXctZmllbGQgeyBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBnYXA6IDJweDsgZm9udC1zaXplOiAxM3B4OyB9Ci5yZWNvcmQtc2VjdGlvbiAudmlldy1maWVsZCAudmlldy1sYWJlbCB7IGZvbnQtc2l6ZTogMTFweDsgZm9udC13ZWlnaHQ6IDYwMDsgY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyB9Ci5yZWNvcmQtc2VjdGlvbiAudmlldy1maWVsZCAudmlldy12YWx1ZSB7IGNvbG9yOiB2YXIoLS10ZXh0LXByaW1hcnkpOyB3b3JkLWJyZWFrOiBicmVhay13b3JkOyB9CkBtZWRpYSAobWF4LXdpZHRoOiA2NDBweCkgewogIC5yZWNvcmRzLXRvb2xiYXIgeyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOyBhbGlnbi1pdGVtczogc3RyZXRjaDsgfQogIC5yZWNvcmRzLXNlYXJjaCBpbnB1dCB7IHdpZHRoOiAxMDAlOyB9Cn0KCi8qIC0tLS0tLS0tLS0gTW9kYWwgKHJlY29yZCBlZGl0b3IpIC0tLS0tLS0tLS0gKi8KLm1vZGFsLW92ZXJsYXkgewogIHBvc2l0aW9uOiBmaXhlZDsgaW5zZXQ6IDA7IGJhY2tncm91bmQ6IHJnYmEoMTAsMTEsMTMsMC41KTsKICBiYWNrZHJvcC1maWx0ZXI6IGJsdXIoNnB4KTsgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IGJsdXIoNnB4KTsKICBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogZmxleC1zdGFydDsganVzdGlmeS1jb250ZW50OiBjZW50ZXI7CiAgcGFkZGluZzogNDBweCAxNnB4OyBvdmVyZmxvdy15OiBhdXRvOyB6LWluZGV4OiAyMDA7CiAgYW5pbWF0aW9uOiBvdmVybGF5SW4gMjAwbXMgdmFyKC0tZWFzZSk7Cn0KQGtleWZyYW1lcyBvdmVybGF5SW4geyBmcm9tIHsgb3BhY2l0eTogMDsgfSB0byB7IG9wYWNpdHk6IDE7IH0gfQpAa2V5ZnJhbWVzIG1vZGFsUGFuZWxJbiB7CiAgZnJvbSB7IG9wYWNpdHk6IDA7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgxNHB4KSBzY2FsZSgwLjk3KTsgfQogIHRvIHsgb3BhY2l0eTogMTsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDApIHNjYWxlKDEpOyB9Cn0KLm1vZGFsLXBhbmVsIHsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTEpOyBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMtbGcpOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOwogIGJhY2tkcm9wLWZpbHRlcjogdmFyKC0tZ2xhc3MtYmx1cik7IC13ZWJraXQtYmFja2Ryb3AtZmlsdGVyOiB2YXIoLS1nbGFzcy1ibHVyKTsKICBwYWRkaW5nOiAyNHB4OyB3aWR0aDogMTAwJTsgbWF4LXdpZHRoOiA3MjBweDsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LW1vZGFsKTsKICBtYXgtaGVpZ2h0OiBjYWxjKDEwMHZoIC0gODBweCk7IG92ZXJmbG93LXk6IGF1dG87CiAgYW5pbWF0aW9uOiBtb2RhbFBhbmVsSW4gMjQwbXMgdmFyKC0tZWFzZSk7Cn0KLm1vZGFsLXBhbmVsLndpZGUgeyBtYXgtd2lkdGg6IDExMDBweDsgfQoubW9kYWwtcGFuZWwgaDIgeyBtYXJnaW46IDAgMCA0cHg7IGZvbnQtc2l6ZTogMTdweDsgbGV0dGVyLXNwYWNpbmc6IC0wLjAxZW07IH0KLm1vZGFsLXBhbmVsIC5tb2RhbC1zdWIgeyBjb2xvcjogdmFyKC0tdGV4dC1zZWNvbmRhcnkpOyBmb250LXNpemU6IDEycHg7IG1hcmdpbjogMCAwIDE4cHg7IH0KLmZvcm0tZ3JpZCB7IGRpc3BsYXk6IGdyaWQ7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogcmVwZWF0KGF1dG8tZml0LCBtaW5tYXgoMjAwcHgsIDFmcikpOyBnYXA6IDEycHg7IG1hcmdpbi1ib3R0b206IDE2cHg7IH0KLmZvcm0tZ3JpZC5mdWxsIHsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiAxZnI7IH0KQG1lZGlhIChtYXgtd2lkdGg6IDY0MHB4KSB7IC5mb3JtLWdyaWQgeyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IDFmcjsgfSB9Ci5mb3JtLWZpZWxkIHsgZGlzcGxheTogZmxleDsgZmxleC1kaXJlY3Rpb246IGNvbHVtbjsgZ2FwOiA1cHg7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLXRleHQtc2Vjb25kYXJ5KTsgfQouZm9ybS1maWVsZCBsYWJlbCB7IGZvbnQtd2VpZ2h0OiA2MDA7IH0KLmZvcm0tZmllbGQgdGV4dGFyZWEgeyByZXNpemU6IHZlcnRpY2FsOyBtaW4taGVpZ2h0OiA2MHB4OyB9CgoucGxhdGZvcm0tZWRpdC1yb3cgewogIGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1zbSk7IHBhZGRpbmc6IDE0cHg7IG1hcmdpbi1ib3R0b206IDEwcHg7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMik7Cn0KLnBsYXRmb3JtLWVkaXQtcm93IC5wbGF0Zm9ybS1lZGl0LWhlYWQgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6IHNwYWNlLWJldHdlZW47IGdhcDogOHB4OyBtYXJnaW4tYm90dG9tOiAxMHB4OyB9Ci5wbGF0Zm9ybS1lZGl0LXJvdyAubWV0cmljcy1ncmlkIHsgZGlzcGxheTogZ3JpZDsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiByZXBlYXQoYXV0by1maXQsIG1pbm1heCgxMjBweCwgMWZyKSk7IGdhcDogOHB4OyB9Ci5yZW1vdmUtcGxhdGZvcm0tYnRuIHsgYm9yZGVyOiBub25lOyBiYWNrZ3JvdW5kOiB0cmFuc3BhcmVudDsgY29sb3I6IHZhcigtLXN0YXR1cy1jcml0aWNhbCk7IGN1cnNvcjogcG9pbnRlcjsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNjAwOyB0cmFuc2l0aW9uOiBvcGFjaXR5IDE1MG1zIHZhcigtLWVhc2UpOyB9Ci5yZW1vdmUtcGxhdGZvcm0tYnRuOmhvdmVyIHsgb3BhY2l0eTogMC43OyB9Ci5tb2RhbC1hY3Rpb25zIHsgZGlzcGxheTogZmxleDsganVzdGlmeS1jb250ZW50OiBzcGFjZS1iZXR3ZWVuOyBhbGlnbi1pdGVtczogY2VudGVyOyBtYXJnaW4tdG9wOiAxOHB4OyBnYXA6IDhweDsgZmxleC13cmFwOiB3cmFwOyB9CgovKiAtLS0tLS0tLS0tIFJlc3BvbnNpdmUgdGlnaHRlbmluZyAtLS0tLS0tLS0tICovCkBtZWRpYSAobWF4LXdpZHRoOiA3MjBweCkgewogIC50b3BiYXIgeyBnYXA6IDEycHg7IHBhZGRpbmc6IDEwcHggMTRweDsgZmxleC13cmFwOiB3cmFwOyB9CiAgLnRvcGJhci1icmFuZCB7IG9yZGVyOiAxOyB9CiAgLnRvcGJhci11c2VyIHsgb3JkZXI6IDI7IG1hcmdpbi1sZWZ0OiBhdXRvOyB9CiAgLnRhYnMgeyBvcmRlcjogMzsgd2lkdGg6IDEwMCU7IH0KICAudmlldy1hcmVhIHsgcGFkZGluZzogMTRweDsgfQogIC5maWx0ZXItYmFyIHsgdG9wOiBhdXRvOyBwb3NpdGlvbjogc3RhdGljOyBwYWRkaW5nOiAxMnB4IDE0cHg7IH0KICAuc3RhdC1ncmlkIHsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiByZXBlYXQoYXV0by1maXQsIG1pbm1heCgxNDBweCwgMWZyKSk7IH0KICAuYnJhbmQtbG9nbyB7IGhlaWdodDogMjJweDsgfQp9Cjwvc3R5bGU+CjwvaGVhZD4KPGJvZHk+CjxkaXYgY2xhc3M9ImF1dGgtc2NyZWVuIiBpZD0iYXV0aFNjcmVlbiI+CiAgPGRpdiBjbGFzcz0iYXV0aC1jYXJkIj4KICAgIDxkaXYgY2xhc3M9ImF1dGgtYnJhbmQiPgogICAgICA8aW1nIGNsYXNzPSJicmFuZC1sb2dvIiBhbHQ9IkxpZ29uLVJhem9uIFNvbHV0aW9ucyBsb2dvIiAvPgogICAgICA8c3BhbiBjbGFzcz0iYnJhbmQtdGl0bGUiPlNvY2lhbCBNZWRpYSBBbmFseXRpY3M8L3NwYW4+CiAgICA8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImF1dGgtZm9ybSI+CiAgICAgIDxkaXYgY2xhc3M9ImZvcm0tZmllbGQiPgogICAgICAgIDxsYWJlbCBmb3I9ImF1dGhDb2RlIj5BY2Nlc3MgY29kZTwvbGFiZWw+CiAgICAgICAgPGlucHV0IHR5cGU9InBhc3N3b3JkIiBpZD0iYXV0aENvZGUiIGF1dG9jb21wbGV0ZT0ib2ZmIiAvPgogICAgICA8L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0iYXV0aC1lcnJvciIgaWQ9ImF1dGhFcnJvciI+PC9kaXY+CiAgICAgIDxidXR0b24gY2xhc3M9ImJ0biBwcmltYXJ5IiBpZD0iYXV0aFN1Ym1pdEJ0biIgdHlwZT0iYnV0dG9uIj48aSBkYXRhLWx1Y2lkZT0iYXJyb3ctcmlnaHQiIHN0eWxlPSJ3aWR0aDoxNHB4O2hlaWdodDoxNHB4OyI+PC9pPiBFbnRlcjwvYnV0dG9uPgogICAgPC9kaXY+CiAgPC9kaXY+CjwvZGl2PgoKPGRpdiBjbGFzcz0iYXBwLXNoZWxsIiBpZD0iYXBwU2hlbGwiIHN0eWxlPSJkaXNwbGF5Om5vbmU7Ij4KICA8aGVhZGVyIGNsYXNzPSJ0b3BiYXIiPgogICAgPGRpdiBjbGFzcz0idG9wYmFyLWJyYW5kIj4KICAgICAgPGltZyBjbGFzcz0iYnJhbmQtbG9nbyIgYWx0PSJMaWdvbi1SYXpvbiBTb2x1dGlvbnMgbG9nbyIgLz4KICAgICAgPHNwYW4gY2xhc3M9ImJyYW5kLXRpdGxlIj5Tb2NpYWwgTWVkaWEgQW5hbHl0aWNzPC9zcGFuPgogICAgPC9kaXY+CiAgICA8bmF2IGNsYXNzPSJ0YWJzIiByb2xlPSJ0YWJsaXN0IiBhcmlhLWxhYmVsPSJTZWN0aW9ucyI+CiAgICAgIDxidXR0b24gY2xhc3M9InRhYi1idG4gaXMtYWN0aXZlIiBkYXRhLXRhYj0iZGFzaGJvYXJkIiByb2xlPSJ0YWIiIGFyaWEtc2VsZWN0ZWQ9InRydWUiPjxpIGRhdGEtbHVjaWRlPSJsYXlvdXQtZGFzaGJvYXJkIiBzdHlsZT0id2lkdGg6MTRweDtoZWlnaHQ6MTRweDsiPjwvaT4gRGFzaGJvYXJkPC9idXR0b24+CiAgICAgIDxidXR0b24gY2xhc3M9InRhYi1idG4iIGRhdGEtdGFiPSJyZWNvcmRzIiByb2xlPSJ0YWIiIGFyaWEtc2VsZWN0ZWQ9ImZhbHNlIj48aSBkYXRhLWx1Y2lkZT0iZGF0YWJhc2UiIHN0eWxlPSJ3aWR0aDoxNHB4O2hlaWdodDoxNHB4OyI+PC9pPiBEYXRhIFJlY29yZHM8L2J1dHRvbj4KICAgICAgPGJ1dHRvbiBjbGFzcz0idGFiLWJ0biIgZGF0YS10YWI9ImZvbGxvd2VycyIgcm9sZT0idGFiIiBhcmlhLXNlbGVjdGVkPSJmYWxzZSI+PGkgZGF0YS1sdWNpZGU9InVzZXJzIiBzdHlsZT0id2lkdGg6MTRweDtoZWlnaHQ6MTRweDsiPjwvaT4gRm9sbG93ZXJzIERhdGE8L2J1dHRvbj4KICAgICAgPGJ1dHRvbiBjbGFzcz0idGFiLWJ0biIgZGF0YS10YWI9ImNvbXBhcmlzb24iIHJvbGU9InRhYiIgYXJpYS1zZWxlY3RlZD0iZmFsc2UiPjxpIGRhdGEtbHVjaWRlPSJnaXQtY29tcGFyZSIgc3R5bGU9IndpZHRoOjE0cHg7aGVpZ2h0OjE0cHg7Ij48L2k+IENvbXBhcmlzb25zPC9idXR0b24+CiAgICAgIDxidXR0b24gY2xhc3M9InRhYi1idG4iIGRhdGEtdGFiPSJ1cGxvYWQiIHJvbGU9InRhYiIgYXJpYS1zZWxlY3RlZD0iZmFsc2UiPjxpIGRhdGEtbHVjaWRlPSJ1cGxvYWQtY2xvdWQiIHN0eWxlPSJ3aWR0aDoxNHB4O2hlaWdodDoxNHB4OyI+PC9pPiBVcGxvYWQgRGF0YTwvYnV0dG9uPgogICAgICA8YnV0dG9uIGNsYXNzPSJ0YWItYnRuIiBkYXRhLXRhYj0iaGlzdG9yeSIgcm9sZT0idGFiIiBhcmlhLXNlbGVjdGVkPSJmYWxzZSI+PGkgZGF0YS1sdWNpZGU9Imhpc3RvcnkiIHN0eWxlPSJ3aWR0aDoxNHB4O2hlaWdodDoxNHB4OyI+PC9pPiBVcGxvYWQgSGlzdG9yeTwvYnV0dG9uPgogICAgPC9uYXY+CiAgICA8ZGl2IGNsYXNzPSJ0b3BiYXItdXNlciI+CiAgICAgIDxidXR0b24gY2xhc3M9ImJ0biIgaWQ9ImxvZ291dEJ0biIgdHlwZT0iYnV0dG9uIj48aSBkYXRhLWx1Y2lkZT0ibG9jayIgc3R5bGU9IndpZHRoOjE0cHg7aGVpZ2h0OjE0cHg7Ij48L2k+IExvY2s8L2J1dHRvbj4KICAgICAgPGJ1dHRvbiBjbGFzcz0idGhlbWUtdG9nZ2xlIiBpZD0idGhlbWVUb2dnbGUiIHR5cGU9ImJ1dHRvbiIgYXJpYS1sYWJlbD0iVG9nZ2xlIGRhcmsgbW9kZSI+CiAgICAgICAgPHNwYW4gaWQ9InRoZW1lVG9nZ2xlSWNvbiI+PGkgZGF0YS1sdWNpZGU9Im1vb24iIHN0eWxlPSJ3aWR0aDoxNnB4O2hlaWdodDoxNnB4OyI+PC9pPjwvc3Bhbj4KICAgICAgPC9idXR0b24+CiAgICA8L2Rpdj4KICA8L2hlYWRlcj4KCiAgPHNlY3Rpb24gY2xhc3M9ImZpbHRlci1iYXIiIGlkPSJmaWx0ZXJCYXIiPgogICAgPGRpdiBjbGFzcz0iZmlsdGVyLWZpZWxkIj4KICAgICAgPGxhYmVsIGZvcj0iZmlsdGVyRGF0ZUZyb20iPkZyb208L2xhYmVsPgogICAgICA8aW5wdXQgdHlwZT0iZGF0ZSIgaWQ9ImZpbHRlckRhdGVGcm9tIiAvPgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJmaWx0ZXItZmllbGQiPgogICAgICA8bGFiZWwgZm9yPSJmaWx0ZXJEYXRlVG8iPlRvPC9sYWJlbD4KICAgICAgPGlucHV0IHR5cGU9ImRhdGUiIGlkPSJmaWx0ZXJEYXRlVG8iIC8+CiAgICA8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImZpbHRlci1maWVsZCBmaWx0ZXItcHJlc2V0cyIgaWQ9ImZpbHRlclByZXNldHMiPgogICAgICA8YnV0dG9uIHR5cGU9ImJ1dHRvbiIgZGF0YS1wcmVzZXQ9IjciPkxhc3QgNyBkYXlzPC9idXR0b24+CiAgICAgIDxidXR0b24gdHlwZT0iYnV0dG9uIiBkYXRhLXByZXNldD0iMzAiPkxhc3QgMzAgZGF5czwvYnV0dG9uPgogICAgICA8YnV0dG9uIHR5cGU9ImJ1dHRvbiIgZGF0YS1wcmVzZXQ9IjkwIj5MYXN0IDkwIGRheXM8L2J1dHRvbj4KICAgICAgPGJ1dHRvbiB0eXBlPSJidXR0b24iIGRhdGEtcHJlc2V0PSJhbGwiPkFsbCB0aW1lPC9idXR0b24+CiAgICA8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImZpbHRlci1maWVsZCI+CiAgICAgIDxsYWJlbCBmb3I9ImZpbHRlclBsYXRmb3JtIj5QbGF0Zm9ybTwvbGFiZWw+CiAgICAgIDxzZWxlY3QgaWQ9ImZpbHRlclBsYXRmb3JtIj48b3B0aW9uIHZhbHVlPSJhbGwiPkFsbCBwbGF0Zm9ybXM8L29wdGlvbj48L3NlbGVjdD4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0iZmlsdGVyLWZpZWxkIj4KICAgICAgPGxhYmVsIGZvcj0iZmlsdGVyQ2FtcGFpZ24iPkNhbXBhaWduPC9sYWJlbD4KICAgICAgPHNlbGVjdCBpZD0iZmlsdGVyQ2FtcGFpZ24iPjxvcHRpb24gdmFsdWU9ImFsbCI+QWxsIGNhbXBhaWduczwvb3B0aW9uPjwvc2VsZWN0PgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJmaWx0ZXItZmllbGQiPgogICAgICA8bGFiZWwgZm9yPSJmaWx0ZXJDb250ZW50VHlwZSI+Q29udGVudCB0eXBlPC9sYWJlbD4KICAgICAgPHNlbGVjdCBpZD0iZmlsdGVyQ29udGVudFR5cGUiPjxvcHRpb24gdmFsdWU9ImFsbCI+QWxsIGNvbnRlbnQgdHlwZXM8L29wdGlvbj48L3NlbGVjdD4KICAgIDwvZGl2PgogIDwvc2VjdGlvbj4KCiAgPG1haW4gY2xhc3M9InZpZXctYXJlYSI+CiAgICA8c2VjdGlvbiBpZD0idmlldy1kYXNoYm9hcmQiIGNsYXNzPSJ2aWV3IGlzLWFjdGl2ZSI+PC9zZWN0aW9uPgogICAgPHNlY3Rpb24gaWQ9InZpZXctcmVjb3JkcyIgY2xhc3M9InZpZXciPjwvc2VjdGlvbj4KICAgIDxzZWN0aW9uIGlkPSJ2aWV3LWZvbGxvd2VycyIgY2xhc3M9InZpZXciPjwvc2VjdGlvbj4KICAgIDxzZWN0aW9uIGlkPSJ2aWV3LWNvbXBhcmlzb24iIGNsYXNzPSJ2aWV3Ij48L3NlY3Rpb24+CiAgICA8c2VjdGlvbiBpZD0idmlldy11cGxvYWQiIGNsYXNzPSJ2aWV3Ij48L3NlY3Rpb24+CiAgICA8c2VjdGlvbiBpZD0idmlldy1oaXN0b3J5IiBjbGFzcz0idmlldyI+PC9zZWN0aW9uPgogIDwvbWFpbj4KPC9kaXY+Cgo8ZGl2IGlkPSJ0b2FzdFJvb3QiIGNsYXNzPSJ0b2FzdC1yb290IiBhcmlhLWxpdmU9InBvbGl0ZSI+PC9kaXY+Cgo8c2NyaXB0PgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgQnJhbmQgbG9nbyDigJQgZW1iZWRkZWQgb25jZSBoZXJlIGFuZCB3aXJlZCBvbnRvIGV2ZXJ5IC5icmFuZC1sb2dvCiAgIDxpbWc+IGFuZCB0aGUgZmF2aWNvbiA8bGluaz4gYXQgYm9vdHN0cmFwLCBzbyB0aGUgYmFzZTY0IHBheWxvYWQKICAgYXBwZWFycyBleGFjdGx5IG9uY2UgaW4gdGhpcyBmaWxlIGluc3RlYWQgb2Ygb25jZSBwZXIgdXNhZ2Ugc2l0ZS4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmNvbnN0IExPR09fREFUQV9VUkkgPSAnZGF0YTppbWFnZS9wbmc7YmFzZTY0LGlWQk9SdzBLR2dvQUFBQU5TVWhFVWdBQUNGb0FBQWR6Q0FZQUFBQm5iOG8zQUFBQUNYQklXWE1BQUM0akFBQXVJd0Y0cFQ5MkFBQWdBRWxFUVZSNG5PemRUVzdiV0xvRzRPUGd6dU5hZ1Ywck1Hc0ZVVTA4VFM0SUdKcEZPNGgzUUhFSHpncmFOUk1JR0pWTU9ibk9Db3BaUVRzN1NGYWdpeE14c1ZYNW8yMzlISkxQQTNqWXlORjNKSFdqK2VwN0Q1YkxaUUFBQUFBQUFBQUE0TmVlbUJFQUFBQUFBQUFBUURlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFBQUFBQjBKR2dCQUFBQUFBQUFBTkNSb0FVQUFBQUFBQUFBUUVlQ0ZnQUFBQUFBQUFBQUhRbGFBQUFBQUFEMFVGN1ZXVjdWRis0T0FBQjI2My9NR3dBQUFBQ2dseTVkR3dBQTdKNmdCUUFBQUFCQXorUlZQUThobkxnM0FBRFlQZFVoQUFBQUFBQTlrbGYxY1FqaC9NdUo4NnFldUQ4QUFOZ2RRUXNBQUFBQWdINkpsU0ZQNzV3NGMzOEFBTEE3Z2hZQUFBQUFBRDJSVjNYY1pQSHNYNmUxMFFJQUFIWkkwQUlBQUFBQW9BZnlxajRNSWN5L2MxSWJMUUFBWUljRUxRQUFBQUFBK3VIZmxTRmZITFVoREFBQVlBY0VMUUFBQUFBQUVwZFg5WXNRd3ZPZm5GSjlDQUFBN0lpZ0JRQUFBQUJBd3RwdEZaZS9PS0g2RUFBQTJCRkJDd0FBQUFDQXRNMS9VQmx5bDQwV0FBQ3dJNElXQUFBQUFBQ0p5cXM2QmloZWRUaWRqUllBQUxBamdoWUFBQUFBQUFucVdCbnl4ZE84cW9VdEFBQmdCd1F0QUFBQUFBRFNkQjVDT0xySHlRUXRBQUJnQndRdEFBQUFBQUFTMDI2bktPNTVxb2w3QkFDQTdSTzBBQUFBQUFCSVQ5ZktrTHRzdEFBQWdCMFF0QUFBQUFBQVNFaGUxZk1Rd3NrRFR2U1Evd3dBQUhCUGdoWUFBQUFBQUluSXEvbzRoSEQrME5Qa1ZhMCtCQUFBdGt6UUFnQUFBQUFnSGJFeTVPa2pUaU5vQVFBQVd5Wm9BUUFBQUFDUWdMeXE0eWFMWjQ4OFNlWXVBUUJndXdRdEFBQUFBQUQyTEsvcXd4RENmQU9uRUxRQUFJQXRFN1FBQUFBQUFOaS94MWFHZkhHVVYvV3grd1FBZ08wUnRBQUFBQUFBMktPOHFsK0VFSjV2OEFTMldnQUF3QllKV2dBQUFBQUE3RWxiR1hLNTRYOWQwQUlBQUxaSTBBSUFBQUFBWUgvbUc2b011V3ZpUGdFQVlIc0VMUUFBQUFBQTlpQ3Y2aGlJZUxXRmYvbVord1FBZ08wUnRBQUFBQUFBMkxFdFZZWjhsVmUxK2hBQUFOZ1NRUXNBQUFBQWdOMDdEeUVjYmZGZlZSOENBQUJiSW1nQkFBQUFBTEJEN2JhSllzdi9vbzBXQUFDd0pZSVdBQUFBQUFDN3RiWEtrRHNFTFFBQVlFc0VMUUFBQUFDQXJja1c1U3hibEI3NnQvS3Fub2NRVG5id1Q1M2tWWDI0ZzM4SEFBQkc1MzljT1FBQUFBQ3dhVEZnRVVLSW9ZS2JabHBNRFBoenlPSjRCNVVoZDhXQXkvVU8vejBBQUJnRlFRc0FBQUFBWUdQdUJDeU9RZ2lmUWdnejAvMXFGNVVoZDAwRUxRQUFZUE1FTFFBQUFBQ0FSOHNXNVlzUXdrVWJzUGhpM2t5TEc5UDl2TTNpUElUd2JNZi9yTW9XQUFEWUFrRUxBQUFBQU9EQnNrVTVhVGRZL0R0RThLNlpGaGNtKzdVeVpMNkhmMXBsQ3dBQWJNRVRRd1VBQUFBQTdpc0dMTEpGR1dzcC91OEhteHJPRGZXckdEaDV1b2QvOTJrYjhnQUFBRGJJUmdzQUFBQUFvTE9mYkxDNHEyeW1SV09xbjdkWnhFcVY1M3M4UXF3UFVkOENBQUFiSkdnQkFBQUFBUHhTeDRCRjlMNlpGdnVveVVoT1h0V0hJWVRMUFo4cjN0dWJCTWNEQUFDOUpXZ0JBQUFBQVB4UXRpaVAyNERGeTQ1VG1wbm1WL3VxRExrcjIvTy9Ed0FBZ3lOb0FRQUFBQUI4NHdFQmkraTF5cENWdktvbjk1emR0dnhxQXdrQUFIQlBCOHZsMHN3QUFBQUFnTThlR0xDSVBzVC9lRE10UG81OWttMWxTQXljSENWd25PalBxN1BUNndUT0FRQUFnMkNqQlFBQUFBRHdtSURGRnpNaGk2L21DWVVzUWxzZkltZ0JBQUFiSW1nQkFBQUFBQ09XTGNxNGZlRThoRkE4WWdxeE1zU0QvTlUyaXhocWVKWEFVZTdLMGprS0FBRDBuNkFGQUFBQUFJelFuWUJGL0h2NmlBbDhhRGM0c0hLWjRCd21DWndCQUFBRzQyQzVYTHBOQUFBQUFCaUpEUVlzdnZqZlpscTg4Zjc1dk0xaS9zak5JTnYwMjlYWnFXb1hBQURZQUJzdEFBQUFBR0FFdGhDd2lONEtXYXprVlgyY2NNZ2l0UFVoNmwwQUFHQURuaGdpQUFBQUFBeGJ0aWhuSVlTYk5naXdxWkRGcHhEQ3pGdm5xeFFyUSs1U0h3SUFBQnRpb3dVQUFBQUFERlFic0loMUZrZGJlSVd6Wmxxb29saHRzNGhiUXA0bGNKU2ZFYlFBQUlBTkViUUFBQUFBZ0lIWmNzQWllcWN5WktXdERKbW5jSlpmeUpJK0hRQUE5SWpxRUFBQUFBQVlpQml3eUJabHJBajV6eFpERmlwRDFsMXNzSTVsbTU2Mm9SQUFBT0NSYkxRQUFBQUFnSjdMRm1Xc2hiamNZcmppcm5rekxXNjhaejV2czNnUlFuaWV3Rkc2K3ZJK0FRQUFIa0hRQWdBQUFBQjZxZzFZeE5xS1p6dDZCYkV5NU1MNzVYUEk0ckNIb1FYMUlRQUFzQUdDRmdBQUFBRFFNM3NJV0FTVklkL29TMlhJWFpOMGpnSUFBUDBsYUFFQUFBQUFQYkduZ01VWEZ5cERWdktxanZmd01vV3ozTk5KcjA0TEFBQ0pPbGd1bCs0R0FBQUFBQktXTGNyanRxWmlId0dMNkgwekxkUk8zRmFHTkNHRW93U084eEIvWHAyZFh2ZnYyQUFBa0E0YkxRQUFBQUFnVVczQVlwN0E5Z1NWSWJmbVBRNVpSREV3STJnQkFBQ1BJR2dCQUFBQUFJbEpLR0FSbGMyMGFCSTR4OTYxbFNHdmV2NHk0bXU0U09BY0FBRFFXNElXQUFBQUFKQ0l4QUlXMFlkbVdzd1RPRWNxaGhCUVVBRURBQUNQSkdnQkFBQUFBSHVXTGNyRDlpRitLZ0dMTDFTR3RQS3Fqb0dUa3lRTzh6aEhlVlVmWHAyZGZ1enppd0FBZ0gwU3RBQUFBQUNBUFdrREZ1ZnQzOVBFN3VGMU15MnVFempIM3VWVkhiZEFGQU42U2JFKzVFMEM1d0FBZ0Y0U3RBQUFBQUNBSFVzOFlCRjlhQ3RNV0JsQ1pjaGRtYUFGQUFBOG5LQUZBQUFBQU94SUR3SVdYOHlhYWFGYVlyWE5JdDdWc3dTT3NrbVQ0YndVQUFEWVBVRUxBQUFBQU5peUhnVXNvcjlVaHF6a1ZYMDgwTTBlUXd1T0FBREFUajB4YmdBQUFBRFlubXhSemtJSVRRaWg2RUhJNGxNYkJtSGxzZ2QzOWlCNVZXYzlQRFlBQUNUQlJnc0FBQUFBMklJMllCRzNJUnoxYUw0cVExcDVWYjhZK09hSHJBMEFBUUFBOXlSb0FRQUFBQUFiMU5PQVJmUzJtUlp2RWpqSDN1VlZmZGh1c3hpeXlRaGVJd0FBYklXZ0JRQUFBQUJzUUk4REZxR3RESmtsY0k1VURMWXk1QTdWSVFBQThFQ0NGZ0FBQUFEd0NObWluTFFCaXo3WFRNeFZocXprVlIzdjgza0taOW15azBHL09nQUEyS0tENVhKcHZnQUFBQUJ3VHdNSldFVHZtbWt4U2VBY2U5ZFdoalE5M1VyeUVIOWVuWjFlOSsvWUFBQ3dYelphQUFBQUFNQTlEQ2hnRVZTR2ZLT3YxUzhQRmQvTGdoWUFBSEJQZ2hZQUFBQUEwTUhBQWhaZnhNcVFtelNPc2w5dFpjaXJrYjNzTElFekFBQkE3d2hhQUFBQUFNQlBaSXZ5T0lSd0VVSjRQckE1eGNxUWl3VE9rWW94emtMUUFnQUFIdUNKb1FFQUFBREF0MkxBSWx1VWx5R0UvdzR3WkJHZEozQ0dKT1JWSFRlVm5JendwUi9sVlgyY3dEa0FBS0JYYkxRQUFBQUFnRHZhRFJieHdmdkxBYytsYktaRms4QTU5aTZ2NnJqVm9SanhDT0xyVng4REFBRDNJR2dCQUFBQUFPTUpXRVR2bTJreFQrQWNxUmg3ZmNva2hQQW1nWE1BQUVCdkNGb0FBQUFBTUdyWm9qeHNBeGF2UmpLSFdRSm5TRUplMWJFKzVkbkl4NUFsY0FZQUFPZ1ZRUXNBQUFBQVJxa05XSnkzZjA5SE1vUFhLa05XOHFyK3NzRms3TVllTkFFQWdIc1R0QUFBQUFCZ1ZFWWFzSWcrQ0Jhc3VSelovZjlRWHRYWjFkbXBBQTRBQUhRa2FBRUFBQURBS0l3NFlQSEZySmtXSDlNNHluN2xWZjNDSm9jMWt4Q0NvQVVBQUhRa2FBRUFBQURBNEdXTDhyemQ1akRXRFFheE11UTZnWFBzWFY3VmgrMDJDMjVsWmdFQUFOMEpXZ0FBQUFBd1dObWluTFVCaTZNUjM3TEtrSFVxUTc0bGFBRUFBUGR3c0Z3dXpRc0FBQUNBUVJHd1dQTy96YlI0azlCNTlxYXREUGw3cEMvL1YzNjdPanRWTFFNQUFCM1lhQUVBQUFEQVlBaFlmT090a01WS1d4bHlrY0paRWhXM1dxaVhBUUNBRGdRdEFBQUFBT2k5YkZHK2FCK2lDMWpjK2hSQ21LVnltQVFJNFB6Y1JOQUNBQUM2RWJRQUFBQUFvTGV5UlRscEg2QS9jNHZmbURYVFFoWEVhcHRGZkorOFN1QW9LY3ZHUGdBQUFPaEswQUlBQUFDQTNoR3crS1YzS2tQV1hDWjBsbFJOeGo0QUFBRG82b2xKQVFBQUFOQVhNV0NSTGNwWWIvQi9RaFkvcERMa2pyeXFWWVowOHpTdjZ1TStIQlFBQVBiTlJnc0FBQUFBa3BjdHlsaHJjQ0ZjMGNtOG1SWTNQVGpuMXVWVkhkODN4Y0JmNWlaTmJQOEFBSUJmRTdRQUFBQUFJRm5ab2p4dUswSmV1cVZPWW1YSVJRL091U3RDQS9lVDllbXdBQUN3TDRJV0FBQUFBQ1JId09KQlZJYmNrVmYxZVFqaEpKa0Q5WU9nQlFBQWRDQm9BUUFBQUVBeUJDd2U1VUpseUVwZTFWL2VSOXlQYWg0QUFPamdZTGxjbWhNQUFBQUFleVZnOFdqdm0ybGhHMEVycitwcm9ZRUgrL1BxN1BTNnAyY0hBSUNkc05FQ0FBQUFnTDNKRnVWaENPRzgvWHZxSmg1TVpVZ3JyK3Faa01XanhNQ09vQVVBQVB5RW9BVUFBQUFBT3lkZ3NWRmxNeTJhQWIyZUI4dXJPcjZ2TG5wNi9GVFlqQUlBQUw4Z2FBRUFBQURBemdoWWJOeUhabHJNQi9hYUh1UFMrK3JSSmowL1B3QUFiSjJnQlFBQUFBQmJKMkN4TlNwRFdubFZ2d2doUEUvaU1QMTJGRGVEWEoyZGZoejdJQUFBNEVlZW1Bd0FBQUFBMjVRdHloZ0d1QWtoRkVJV0cvVzZtUmJYQTNvOUQ2WXlaT1BVaHdBQXdFL1lhQUVBQUFEQVZyUUJpMWhyY1dUQ0cvZWhuUzByM21lYkZldERoSGdBQU9BSEJDMEFBQUFBMkNnQmk1MllOZE5DdGNOcW0wVU1CYnhLNENoRE1objdBQUFBNEdjRUxRQUFBQURZQ0FHTG5mbExaY2lheTRUT01oU3FRd0FBNENlZUdBNEFBQUFBajVFdHlrbTJLSnNRd24rRUxMYnVVd2poZk9DdnNiTzhxZ1Y3dHVOcFh0WENGZ0FBOEFNMldnQUFBQUR3SURGZzBXNndlR2FDTzZNeXBOVUdBWW9rRGpOTWNiN04ySWNBQUFEZkkyZ0JBQUFBd0wwSVdPek4yMlphdkJucGEvOGVsU0hiWmFNRkFBRDhnS0FGQUFBQUFKMElXT3hWckF5WmpmajFyOG1yT3Rhbm5DUjBwQ0dhakgwQUFBRHdJNElXQUFBQUFQeFV0aWlQMiswQkFoYjdNMWNac3BKWDlYRWIrR0c3QkZrQUFPQUhuaGdNQUFBQUFOOFRBeGJab293QmkvOEtXZXpWdTJaYVhJejQ5ZjliZkU4K1RldEl3NVJYdGEwV0FBRHdIVFphQUFBQUFMQ20zV0FSTndhOE5KbTlVeGx5UjFzWkl2U3pPMWtJNFhvc0x4WUFBTG9TdEFBQUFBRGdNd0dMSk1YS2tKdXhEeUdzUWhhSEtrTjJMbTYwc0UwRkFBRCtSZEFDQUFBQVlPU3lSZm5sQWZhcnNjOGlNU3BEMXFrTTJiMXNiQzhZQUFDNk9GZ3Vsd1lGQUFBQU1FSnR3T0s4L2ZNQU96MS9OTk9pR2ZzUXdtcWJ4WXNRd3Q4SkhHV01mcnM2Ty8wNDlpRUFBTUJkTmxvQUFBQUFqSXlBUlMrVVFoWXJiV1hJWlFwbkdhbFlIL0ptN0VNQUFJQzdCQzBBQUFBQVJrTEFvamZlTjlOaVB2WWgzREgzZnQyclROQUNBQURXQ1ZvQUFBQUFqRUMyS09jQ0ZyMHhHL3NBdnNpck9tNVRlSlhHYVVack12WUJBQURBdndsYUFBQUFBQXhZdGlobjdVYUFJL2ZjQzY5VmhxeW9ERW5HczdFUEFBQUEvdTJKaVFBQUFBQU1Ud3hZWkl2eUpvVHdIeUdMM3ZqUWhtSllPZmZlVFVOZTFkbllad0FBQUhmWmFBRUFBQUF3SURaWTlOcXNtUllmeHo2RWNQdGd2MGpnS0t6RSs3QnBCUUFBV29JV0FBQUFBQU1nWU5GN3NUTGtldXhEdUVObFNGb203Z1FBQUc0SldnQUFBQUQwV0xZb0oyM0E0cGw3N0MyVklYZmtWUjFuY1pMTWdRanRSZ3NBQUtBbGFBRUFBQURRUXdJV2czS3VNbVFscityak9JOFV6c0lhd1JjQUFMamppV0VBQUFBQTlFY01XR1NMTWxaTS9KK1F4U0M4YmFiRm03RVA0WTVZVC9FMG1kUHdWVjdWRTlNQUFJQVZHeTBBQUFBQWVzQUdpMEg2RkVLWWpYMElYK1JWZmU3OW5iVDRIWFE5OWlFQUFFQVF0QUFBQUFCSVc3WW9qOXVBeFV0WE5UZ3psU0VyZVZVZnR1OXowcFc1R3dBQVdCRzBBQUFBQUVpUWdNWGd2Vk1ac2tabFNQcFVod0FBUU91SlFRQUFBQUNrSXdZc3NrVVpIenIvVjhoaXNGU0czSkZYOVlzUXd2TmtEc1NQUE0ycit0aDBBQURBUmdzQUFBQ0FKTmhnTVNyelpscmNqSDBJNGJZeTVES0JvOUJOckEveDNnVUFZUFFFTFFBQUFBRDJLRnVVOFVIemVRaWhjQStqRUN0RExzWStoRHZtS2tONkpkYUhxTHdCQUdEMEJDMEFBQUFBOXVCT3dPTGNnK2JSVUJseVIxN1Y4YUg5cTJRT1JCZVpLUUVBZ0tBRkFBQUF3RTRKV0l6YWhjcVFGWlVodmZWczdBTUFBSUFnYUFFQUFBQ3dHd0lXby9lK21SYnpzUS9oanZnNU9Fcm1OSFNXVjNWMmRYYmFtQmdBQUdQMnhPMERBQUFBYkZlMktHTmRSTnhrVUFoWmpKYktrRlo4VU45K0Z1aW5pWHNEQUdEc2JMUUFBQUFBMkpJMllESDN5LzNSSzV0cFlRUEFMWlVoL1phTmZRQUFBQ0JvQVFBQUFMQmhBaGJjOFVGbHlLMjhxdU1zVGxJNUR3OGlhQUVBd09nZExKZkxzYzhBQUFBQVlDTUVMUGlPUDV0cGNXMHduME1XeHlHRVJuM09JUHgyZFhiNmNleERBQUJndkd5MEFBQUFBSGlrYkZGTzJqb0VBUXZ1ZWkxa3NlWlN5R0l3NGxZTDcyMEFBRVpMMEFJQUFBRGdnZHFBUmR4Zzhjd00rWmNQN1h1RDFUYUxjNStUUVprSVdnQUFNR2FDRmdBQUFBRDNKR0JCQjdObVdxaFd1SzBNRVRvWmxzbllCd0FBd0xnSldnQUFBQUIwSkdCQlIzK3BERmx6b1RKa2NMS3hEd0FBZ0hFN1dDNlhZNThCQUFBQXdFOWxpekpySHhZTFdQQXJuMElJeDdaWnJPUlYvU0tFOEhjS1oySGpmcjg2TzcweFZnQUF4c2hHQ3dBQUFJQWZ5QmJsbDhxRGwyWkVSeXBEV25sVkg0WVFMcE00RE5zd2NiOEFBSXlWb0FVQUFBREF2d2hZOEVCdm0ybnh4dkMrVWhreWJPcERBQUFZTFVFTEFBQUFnSmFBQlk4UUswTm1CcmlTVi9YRTUyandCQzBBQUJndFFRc0FBQUJnOUxKRmVkaisrdDZEWVI1cXJqSmtSV1hJYUR3Yit3QUFBQml2Sis0ZUFBQUFHS3NZc01nV1pkeGdjU05rd1NPOGE2YkZoUUYrRlQ5VFI0bWNoUzFxTjVjQUFNRG8yR2dCQUFBQWpFNjd3ZUs4L1h2cUhjQWpxQXk1STYvcVdDZnhLcGtEc1czeHZxOU5HUUNBc1JHMEFBQUFBRVpEd0lJdGlKVWhOd2I3bGNxUWNjbkdQZ0FBQU1aSjBBSUFBQUFZUEFFTHRrUmx5QjE1VmNmS2tKTmtEc1F1cUE0QkFHQ1VucmgyQUFBQVlNaXlSUmxySFpvUVFpRmt3WWFkRytoS1h0WEg3V2VNY1RuS3EvclFuUU1BTURZMldnQUFBQUNEMUFZczRpL3NqOXd3VzFBMjA2SXgySzlVaG94WDNHcnhadXhEQUFCZ1hBUXRBQUFBZ0VFUnNHQUgzamZUWW03UUszbFZ4ODBlejFJNEMzdVJDVm9BQURBMmdoWUFBQURBSUFoWXNFTXp3MTVwSzBPRVRzWnRNdllCQUFBd1BvSVdBQUFBUUs5bGkzTFNQdWoxaTNwMjRiWEtrRFVYSVlTbkNaMkgzY3ZNSEFDQXNUbFlMcGN1SFFBQUFPZ2RBUXYyNEVOODZ6WFQ0cVBoZjk1bThTS0U4SGNDUjJILy9yZzZPeFZBQWdCZ05HeTBBQUFBQUhwRndJSTltZ2xack9SVmZSaEN1RXpoTENRaGJyVVF0QUFBWURRRUxRQUFBSUJlRUxCZ3oySmx5TFZMK0VwbENIZXBEd0VBWUZRRUxRQUFBSUNrWll2eXVIMm8rOXhOc1NjZjJwQVBxMjBXTWZUMDBpeTRZMklZQUFDTWlhQUZBQUFBa0tRMllESDNRSmNFbktzTVdWRVp3ZytjR0F3QUFHUHl4RzBEQUFBQUtZa0JpMnhSeGdlNS94V3lJQUZ2bTJueHhrVjhGY05QUjRtY2hZUzBtMDRBQUdBVWJMUUFBQUFBa21DREJRbjZGRUtZdVppVnZLcXpFTUtyRk01Q2t1TDc0OXJWQUFBd0JvSVdBQUFBd0Y1bGkvS3dEVmg0Z0V0cVppcEQxcWdNNFdmaVJvc0xFd0lBWUF3RUxRQUFBSUM5YUFNVzUrM2ZVN2RBWXQ2cERMbVZWM1VNUTUya2NoNlNsTGtXQUFERzRtQzVYTHBzQUFBQVlHY0VMT2lCV0JtU05kUGl4bVY5clF6NUo0R2prTDdmcjg1T2ZXNEFBQmc4R3kwQUFBQ0FuUkN3b0VmbVFoWnIxRUhRVlF6bCtPd0FBREI0VDF3eEFBQUFzRzNab2p4dkg3NFZRaFlrTGxhR0NCYTA4cXFPbjkxblNSeUdQbEFmQWdEQUtOaG9BUUFBQUd4TnRpaG5jVHRBQ09ISWxPbUJXQmt5YzFFcmVWVWZ0NTlmNkdwaVVnQUFqSUdnQlFBQUFMQnhBaGIwMUlYS2tEV1hOdEJ3VDdhZkFBQXdDZ2ZMNWRKTkF3QUFBQnNoWUVHUHZXK21oZHFEVmw3VkwwSUlmeWR4R1BybWo2dXowOGF0QVFBd1pEWmFBQUFBQUkrV0xjcjRVUFpDd0lJZVV4blN5cXY2c04xbUFROFJBMHVDRmdBQURKcWdCUUFBQVBCZzJhS2N0QnNzckl1bno4cG1Xbmd3ZkV0bENJOHhFZFFCQUdEb0JDMEFBQUNBZXhPd1lFQStOTk5pN2tKWDhxcU9uKzNuS1p5RjNsTEJBd0RBNEIwc2wwdTNEQUFBQUhRaVlNRUEvZGxNaTJzWCs3VXlwRkVCeEFiOGRuVjIrdEVnQVFBWUtoc3RBQUFBZ0YvS0ZtWDhoZktGZ0FVRDgxcklZczFjeUlJTmlmK2Q0Yk1GQU1CZ0NWb0FBQUFBUDVRdHl1UDI0ZXRMVTJKZ1ByVHZiVzRyUTE2WkJSc3lFYlFBQUdESUJDMEFBQUNBYndoWU1BS3pabHFvTnJoMWtjcEJHSVRNTlFJQU1HU0NGZ0FBQU1CWEFoYU14RjhxUTI3bFZSMC84eWVwbklkQm1MaEdBQUNHN0dDNVhMcGdBQUFBR0RrQkMwYmtVd2poMkRhTGxieXE0K2FCZjFJNEM0UHorOVhaNlkxckJRQmdpR3kwQUFBQWdCSExGdVZoQ09HOC9YdnF2Y0FJcUF4WnB6S0ViWWtoSGtFTEFBQUdTZEFDQUFBQVJrakFncEY2MjB5TE55NS9KYS9xK1BsL2xzSlpHS1JZSCtMekJnREFJQWxhQUFBQXdJZ0lXREJpc1RKazVnMndrbGYxbDdvZzJKYk1aQUVBR0NwQkN3QUFBQmdCQVFzSWM1VWhheTU5RjdCbHRxVUFBREJZQjh2bDB1MENBQURBZ0dXTE12NksvOEpEVlVic1hUTXRKdDRBSzNsVnZ3Z2gvSjNDV1JpOFA2N09UaHZYREFEQTBOaG9BUUFBQUFQVkJpeGlOY0NSTzJiRVZJYmNrVmYxWWJ2TkFuWWhCcHdFTFFBQUdCeEJDd0FBQUJnWUFRdFlFeXREYm96a0s1VWg3RkptMmdBQURKR2dCUUFBQUF5RWdBVjhJMWFHWEJqTFNsN1ZjYnZBOHhUT3dtaW83QUVBWUpBT2xzdWxtd1VBQUlBZXl4WmxmSkFWSHlhZnVFZFk4MGN6TGRRVzNGYUdOSUpZN01GdlYyZW5IdzBlQUlBaHNkRUNBQUFBZXFvTldNUU5Gcy9jSVh5akZMSllZOXNOK3hMclE2NU5Id0NBSVJHMEFBQUFnSjRSc0lCZmV0OU1pN2t4cmJTVklhOVNPQXVqTkJHMEFBQmdhQVF0QUFBQW9DY0VMS0N6bVZHdHVVem9MSXpQeEowREFEQTBnaFlBQUFDUXVHeFJIcmNQU2dVczROZGVxd3k1bFZlMXloRDJMWE1EQUFBTXpjRnl1WFNwQUFBQWtLQTJZQkVma3I1MFA5REpoL2pSYWFiRlIrUDZITEtJRDdqL1NlQW84UHZWMmVuTjZLY0FBTUJnMkdnQkFBQUFpUkd3Z0FlYkNWbXNVUmxDS2liZWp3QUFESW1nQlFBQUFDUkN3QUllSlZhR1hCdmhTbDdWNXlHRWt4VE9BdXBEQUFBWUdrRUxBQUFBMkxOc1VSNjJBWXRYN2dJZTVFUDdHV0lWc2pnMkR4SWphQUVBd0tBSVdnQUFBTUNldEFHTDgvYnZxWHVBQnp0WEdiTG0wbmNLaVhubVFnQUFHSktENVhMcFFnRUFBR0NIQkN4Z285NDIwK0tGa2E3a1ZUMExJZnduaGJQQXYveDVkWGFxM2djQWdFR3cwUUlBQUFCMlJNQUNOdTVUQ0dGbXJDdDVWY2Z2bUlzVXpnTGZFZXREQkMwQUFCaUVKNjRSQUFBQXRpOWJsUE1Rd2swSW9SQ3lnSTJacVF4Wm96S0VsRTNjRGdBQVEyR2pCUUFBQUd4UnRpampyKzFqeU9MSW5HR2ozalhUNG8yUnJ1UlZIZXRUbnFkd0Z2aUJ6R0FBQUJpS2crVnk2VElCQUFCZ3d3UXNZS3RpWlVqV1RJc2JZLzVhR2RMNHZxRUhmcnM2TzdXRkJnQ0EzclBSQWdBQUFEWkl3QUoyWWk1a3NjWjNEbjBSNjBOc29nRUFvUGNFTFFBQUFHQURCQ3hnWjJKbHlJVnhyK1JWSFI5Y3YwcmhMTkJCSm1nQkFNQVFDRm9BQUFEQUkyU0xjdElHTEo2WkkyeGRyQXlaR2ZPYXk0VE9Bcjh5TVNFQUFJWkEwQUlBQUFBZVFNQUM5dUpDWmNpdHZLcHQwYUZ2TWpjR0FNQVFIQ3lYU3hjSkFBQUFIUWxZd042OGI2YUZoN1N0dktyakxQNUo0akJ3UDM5Y25aMDJaZ1lBUUovWmFBRUFBQUFkQ0ZqQTNxa01XYWN5aEw2S0lTRkJDd0FBZWszUUFnQUFBSDRpVzVUSGJjRGlwVG5CM3BUTnRQQmd0cFZYOVhrSTRTU0p3OEQ5MlV3REFFRHZDVm9BQUFEQWR3aFlRREkrTk5OaTdqcFc4cXIrOHQwRWZUVnhjd0FBOUoyZ0JRQUFBTndoWUFISlVSbXlMbGFHUEUzcFFIQlB0ckVBQU5CN0I4dmwwaTBDQUFBd2VnSVdrS1RYemJRNGR6VXJlVlhIME1sL1VqZ0xQTktmVjJlbjE0WUlBRUJmMldnQkFBREFxR1dMOGpDRUVCL2tGbU9mQlNUbWc0cU1XM2xWeCsrcWkxVE9BNDhVNjBNRUxRQUE2QzFCQ3dBQUFNTTloc3NBQUNBQVNVUkJWRWJwVHNEaTNCcCtTTktzbVJZZlhjMVhLa01Za3N4dEFnRFFaNElXQUFBQWpJcUFCZlRDWDgyMDhHdjNWbDdWTDBJSXo1TTRER3lHb0FVQUFMMTJzRnd1M1NBQUFBQ0RKMkFCdmZFcGhIQnNtOFZLV3hseTQzdUxBZnI5NnV6MHhzVUNBTkJITmxvQUFBQXdlTm1pbklVUUxqeW9oRjVRR2JKdTdydUxnY3JhRUJFQUFQU09vQVVBQUFDRDFRWXM0a1BLSTdjTXZmQzJtUlp2WE5WS1h0V1RFTUtyRk00Q1d4Q0RGajd2QUFEMGtxQUZBQUFBZ3lOZ0FiMFVLME5tcm02bHJReTVUT0Vzc0NVVGd3VUFvSzhFTFFBQUFCZ01BUXZvdGJuS2tEWG52c3NZdUdjdUdBQ0F2anBZTHBjdUR3QUFnRjdMRnVXay9lVzNoNUxRVCsrYWFlSFg3YTI4cW1PbHdqOUpIQWEyNjQrcnM5UEdqQUVBNkJzYkxRQUFBT2l0Tm1BeDk2dFk2RFdWSWQ5U0djSll4RkNSb0FVQUFMMGphQUVBQUVEdkNGakFvTVRLa0J0WHVwSlhkZnh1TzBuaExMQURFOEVpQUFENlNOQUNBQUNBM2hDd2dNR0psU0VYcm5VbHIrcmpFTUo1Q21lQkhja01HZ0NBUGhLMEFBQUFJSG5ab293UFlpNEVMR0J3aEFyV3hWLzJQMDNwUUxCbEozbFZIMTZkblg0MGFBQUEra1RRQWdBQWdHUmxpL0s0M1dEeDBpM0I0SlROdEdoYzYwcGUxZWZDWkl4VURGTmV1M3dBQVBwRTBBSUFBSURrQ0ZqQTRMMXZwc1hjTmEvRVgvUzMzM2t3UmhOQkN3QUEra2JRQWdBQWdHUUlXTUJvekZ6MUdwVWhqRm5tOWdFQTZCdEJDd0FBQVBZdVc1VHgxOXdYQWhZd0NxOVZodHpLcS9wRkNPRjVLdWVCUFpnWU9nQUFmWE93WEM1ZEdnQUFBSHZSQml6TzJ6Ky81b2JoK3hBLytzMjArT2l1djFhRzNQaitnL0Q3MWRucGpURUFBTkFYTmxvQUFBQ3djd0lXTUZveklZczFjOStCOEZuV2hvNEFBS0FYQkMwQUFBRFlHUUVMR0xWWUdYSTk5aUY4a1ZkMXJFdDRsY1pwWU8vaTUrR05hd0FBb0M4RUxRQUFBTmc2QVFzWXZRL3Q5Z1p1SzBNdXpRSyt5b3dDQUlBK0ViUUFBQUJncTdKRk9Xc2ZzQjZaTkl6V3VjcVFOZWUrRTJITk0rTUFBS0JQRHBiTHBRc0RBQUJnNHdRc2dOYmJabHE4TUl5VnZLcmpML2YvU2VFc2tKZy9yODVPMVFzQkFOQUxObG9BQUFDd1VRSVd3QjJmUWdnekExbWpNZ1MrTDRhUUJDMEFBT2dGUVFzQUFBQTJRc0FDK0k2WnlwQmJlVlhINzhpVFZNNERpY2xjQ0FBQWZTRm9BUUFBd0tOa2kzTFNCaXowcXdOM3ZXdW14UnNUV2NtcitqaUVjSjdDV1NCUkV4Y0RBRUJmQ0ZvQUFBRHdJQUlXd0Urb0RQbFdyQXg1bXRxaElDRkhlVlVmWHAyZDJvSURBRUR5QkMwQUFBQzRGd0VMb0lONU15MXVER29scitwejM1blFTYXdQdVRZcUFBQlNKMmdCQUFCQUp3SVdRRWV4TXVUQ3NGYmlML1RiNzA3ZzF5YUNGZ0FBOUlHZ0JRQUFBRCtWTGNyakVFSjhhUHJjcElCZlVCbnlMWlVoME4zRXJBQUE2QU5CQ3dBQUFMNnJEVmpFWDJHL05DR2dvd3VWSWJmeXFuNGhwQWIza2hrWEFBQjljTEJjTGwwVUFBQUFYd2xZQUEvMHZwa1dIcEsyMnNxUUc5c3M0TjUrdnpvN0ZkZ0NBQ0JwTmxvQUFBRHdtWUFGOEVncVE5WmRDRm5BZzB6YXloMEFBRWlXb0FVQUFNRElaWXZ5c0ExWXZCcjdMSUFISzV0cDBSamZTbDdWRTZFMWVEQ2JjUUFBU0o2Z0JRQUF3RWkxQVl2ejlzK3Zyb0dIK3RCTWk3bnByYlNWSVg2TkR3ODNNVHNBQUZJbmFBRUFBREF5QWhiQWhxa01XUmRESjBjcEhRaDY1c1NGQVFDUXVvUGxjdW1TQUFBQVJrREFBdGlDMTgyME9EZllsYnlxWStYQlB5bWNCWHJ1ejZ1ejAydVhDQUJBcXA2NEdRQUFnT0hMRm1WOEVIb1RRaWlFTElBTitkQnViK0NXeWhEWWpNd2NBUUJJbWVvUUFBQ0FBY3NXNWN3YWUyQkxaczIwK0dpNEszbFZ6MVVld01aTVFnZ1h4Z2tBUUtvRUxRQUFBQVpJd0FMWXNyK2FhV0d0Znl1djZ1TjJZeEN3R1RaYUFBQ1FORUVMQUFDQUFSR3dBSGJnVXdqaDNLRFhxQXlCelRyS3Evcnc2dXpVMWh3QUFKSWthQUVBQURBQTJhSjgwYTdZRnJBQXRrMWx5QjE1VmNmUXliTmtEZ1RERWV0RDNyaFBBQUJTSkdnQkFBRFFZOW1pbkxRYkxEemtBM2JoYlRNdFBQaHN0WlVoOHlRT0E4T1RDVm9BQUpBcVFRc0FBSUFlRXJBQTlpQldoc3dNZmszY0pQUTBvZlBBa0V6Y0pnQUFxUkswQUFBQTZCRUJDMkNQNWlwRGJ1VlZIU3VibnFkeUhoaWd6S1VDQUpDcWcrVnk2WElBQUFBU2x5M0tyUDNsdElBRnNBL3ZtbW5oMStXdHZLb1BRd2czdGxuQTF2MXhkWGJhR0RNQUFLbXgwUUlBQUNCaDJhSThiamRZdkhSUHdKNm9EUG1XeWhEWWpSZzBGYlFBQUNBNWdoWUFBQUFKRXJBQUVoSXJRMjVjeUVwZTFSUGZ6YkF6OGZOMmFkd0FBS1JHMEFJQUFDQWhBaFpBWW1KbHlJVkxXV2tyUXp6MGhkM0p6Qm9BZ0JRSldnQUFBQ1JBd0FKSTFMbUxXUk8vcDQ4U09nOE0zWWtiQmdBZ1JRZkw1ZExGQUFBQTdFbTJLQS9iQjVubit2NkJ4SlROdEppN2xKVzhxdU12Ni85SjRTd3dNbjllbloxZXUzUUFBRkppb3dVQUFNQWVDRmdBaVhzdlpQRU5sU0d3SDVNUWdxQUZBQUJKRWJRQUFBRFlJUUVMb0NkbUx1cFdYdFZ6RlFhd041blJBd0NRR2tFTEFBQ0FIUkN3QUhya2RUTXRHaGUya2xmMWNRaWhTT0VzTUZLQ0ZnQUFKT2VKS3dFQUFOaXViRkhHWDRiZnRBL3FoQ3lBbEgwSUlhZ01XYWN5QlBicnFBMDhBUUJBTW15MEFBQUEySkkyWUJFZldCNlpNZEFUczJaYWZIUlpLM2xWeHkxRXoxSTRDNHhjMW9aV0FRQWdDWUlXQUFBQUd5WmdBZlJVckF5NWRua3I3Uy9vYmZlQU5NU2d4UnQzQVFCQUtnUXRBQUFBTmtUQUF1Z3hsU0hmdWxUM0JNbVl1QW9BQUZJaWFBRUFBUEJJMmFLTS8rZi9SUWpoeEN5Qm5qcFhHWElycitvWEtrTWdLVDZQQUFBazVXQzVYTG9SQUFDQUIyZ0RGblAvNXovUWMyK2JhZkhDSmE3a1ZYMFlRcml4elFLUzg4ZlYyV25qV2dBQVNJR05GZ0FBQVBja1lBRU15S2NRd3N5RnJsRVpBbW1LLy90TDBBSUFnQ1FJV2dBQUFIUWtZQUVNMEV4bHlLMjhxdVAzL1BOVXpnT3N5WXdEQUlCVUNGb0FBQUQ4UXJZb2o5dGZPQXRZQUVQeXJwa1diOXpvU2xzWmNwbkNXWUR2RXJRQUFDQVpUMXdGQUFEQTk4V0FSYllvNDBPMy93cFpBQU9qTXVSYmNXUFJVV3FIQXI0NmFRTlJBQUN3ZHpaYUFBQUEvRXU3d1NJK2NIdHBOc0JBelp0cGNlTnlWOXJLa0ZjcG5BWDRxYmpWNHRxSUFBRFlOMEVMQUFDQWxvQUZNQkt4TXVUQ1phOHhEK2lIaWFBRkFBQXBFTFFBQUFCR0wxdVVoMjNBd3ErWmdhRlRHZkl2ZVZYSDcvK1RwQTRGL0VobU1nQUFwRURRQWdBQUdLMDJZSEhlL2ozMVRnQkc0RUpseUsyOHF1TkQyeUtWOHdDL05ERWlBQUJTY0xCY0xsMEVBQUF3S2dJV3dFaTliNmFGWDRQZmtWZDFyQ0I0bHN5QmdDNSt2em83RlJnREFHQ3ZiTFFBQUFCR1E4QUNHRG1WSVhma1ZYMHVaQUc5RkxkYVhMbzZBQUQyU2RBQ0FBQVloV3hSemdVc2dCRXJtMm5SZUFPczVGVjlIRUtZcDNBVzRONXM1Z0VBWU84RUxRQUFnRUhMRnVXc2ZaaDI1S2FCa2ZyUVRBdWhnbldYZ25mUVc0SVdBQURzbmFBRkFBQXdTQUlXQUYrcERMa2pyK29YS2tPZzEzeCtBUURZdTRQbGN1a1dBQUNBd1JDd0FGanp1cGtXNTBheWtsZjFZUWpoeGpZTDZMMC9yODVPcjEwakFBRDdZcU1GQUFBd0NBSVdBTi80MEg0dmNrdGxDQXhEckE4UnRBQUFZRzhFTFFBQWdGN0xGdVdrZlpCb2pUVEF1bGt6TFQ2YXlVcGUxZkcvTDU2bmNCYmcwVElqQkFCZ253UXRBQUNBWGhLd0FQaXB2NXBwNGRmZXJiWXk1REtKd3dDYk1ERkZBQUQyU2RBQ0FBRG9GUUVMZ0YvNkZFSTRONlkxcXFWZ1dJNWlnT3JxN05UV0hnQUE5a0xRQWdBQTZBVUJDNERPVkliYzBWYUd2RXJtUU1DbXhQb1FtM3NBQU5nTFFRc0FBQ0JwMmFJOGJnTVdMOTBVd0MrOWJhYkZHMk5hYzVIUVdZRE5tUWhhOFAvczNVOXUzTmJXTCt6dDRQU2xPd0xwak1BOEkxQ2x3Njc5Z1FDaG52V09JTG9qVU5VSWpqS0MxKzRWQ0JEWDZiTHpTaU80NVJGY2FRYldDUFJoSi9TSjQvaVAvbFRWNXQ1OEhzRE5BT1JhRlpGVi9IRXRBSUJVQkMwQUFJQkpFckFBZUxTNE11Uk0yZjdVZEVPOGpyeWN5dkVBVzdWUVRnQUFVaEcwQUFBQUprWEFBdURKbGxhRy9LbnBocmhXNEdJcXh3TnNYYVdrQUFDazh1TCsvbDd4QVFDQTVBUXNBSjdsZW5ONjRlM3V6elRkc0RITkFvcjNyNzZ0TjlvTUFNQyttV2dCQUFBa1ZhMVhoeUdFYzI4ZEF6eVpsU0ZmYUxyaFhNZ0NaaUZPdFJDMEFBQmc3d1F0QUFDQUpENExXTVIvQjdvQThHUnhaY2lOOHYyaDZZWlBFNUtBOGxrZkFnQkFFb0lXQUFEQVhnbFlBR3hWWEJseXFhUi84ZGIxQldiRHlpUUFBSklRdEFBQUFQWkN3QUpnSjg2VjlVOU5OOFFWS2lkVE9SNWc1NndJQWdBZ2laK1VIUUFBMkxWcXZZb1BBdU5ZK3dzaEM0Q3RXVzFPTHpiSytZZW1HMktnejNRUG1KbW1HMHkxQUFCZzcweTBBQUFBZHFaYXI4N0dQZmxIcWd5d1ZSODJweGRMSmYwTEswTmducW9Rd3BYZUF3Q3dUNElXQUFEQTFnbFlBT3pjbVJML3FlbUcxeUdFVjFNNUhtQ3ZGcWJaQUFDd2I0SVdBQURBMWdoWUFPekZyMWFHL01uS0VKaTlhdTRGQUFCZy93UXRBQUNBWjZ2V3E4VTRzbDNBQW1DM2JzZEFHMzhTOElONU80cUJxNzZ0UDg2OUVBQUE3SStnQlFBQThHUmp3Q0krNERwUlJZQzlPTnVjWG5pWU9HcTZJVjZIZnBuRXdRQXB4YjhGNzNVQUFJQjlFYlFBQUFBZVRjQUNJSW00TXVSSzZmL2k3WVNPQlVpbkVyUUFBR0NmQkMwQUFJQUhFN0FBU01iS2tDODAzV0JsQ1BESlFpVUFBTmduUVFzQUFPQ0hxdlVxdmlWNEtXQUJrTXk1bFNGL2Fyb2hYcGN1cG5JOFFITHVVUUVBMkt1ZmxCc0FBUGlXYXIwNnJ0YXJPSmI5Ly9vQkd5Q1ozemFuRjBiaS81V1ZJY0JmakFFc0FBRFlDeE10QUFDQXY0a0JpM0ZFL1J2VkFVanFMb1J3cGdWL2FycmhQSVR3Y2lySEEweEdERnBzdEFNQWdIMFF0QUFBQVA1RHdBSmdjczZzRFBsVDB3MmZybE1BWDFxWWRnTUF3TDRJV2dBQUFERmdjUmhDdUJTd0FKaVVheXREL2lZK1JEMlkyREVCMDJCMUNBQUFleU5vQVFBQU16WUdMTTdIZng1Y0FVeUhsU0ZmYUxvaDF1TmtVZ2NGVEltVlFnQUE3TTJMKy90NzFRWUFnSmtSc0FDWXZQKzlPYjI0MUtZL05OMFFyMXMzcmxuQUQvemN0L1dWSWdFQXNHc21XZ0FBd0l3SVdBQms0VnJJNG0rc0RBRWVZaEZDRUxRQUFHRG5CQzBBQUdBR0JDd0FzbUZseUJlYWJuZ2RRbmcxcVlNQ3BxclNHUUFBOWtIUUFnQUFDbGV0Vi9HQjNUS0VjS1RYQUpOM3VUbTl1TkdtUDR3clEwejNBQjVLMEFJQWdMMTRjWDkvcjlJQUFGQWdBUXVBN0h6WW5GNTRTUGlacGh0aXlPS1h5UndRa0lOLzltMHRzQVlBd0U2WmFBRUFBSVVSc0FESWxwVWhuMm02WVNGa0FUeEJES3dKV2dBQXNGT0NGZ0FBVUFnQkM0Q3NyVGFuRnhzdC9NTzRNdVR0Rkk0RnlFNE1hYjNYTmdBQWRrblFBZ0FBTWxldFY0dHhmLzFMdlFUSTB1M205R0twZFg5eExqZ0lQSkVWVEFBQTdKeWdCUUFBWkdvTVdNUUhjeWQ2Q0pBMUswTSswM1JEZkVoNk1aa0RBbkxqM2hnQWdKMzdTWWtCQUNBdk1XQlJyVmRYSVlULzhVTXlRUForM1p4ZVhHbmpYMWdaQWp6TEdOZ0NBSUNkTWRFQ0FBQXlZWUlGUUhGdXg3L3JqSnB1V0ZxRkJXeEJ2Ry9lS0NRQUFMc2lhQUVBQUJOWHJWZkhJWVRMRU1JcnZRSW95dG5tOU9Lamx2Nmg2WVo0dlR1ZndyRUEyVFBSQWdDQW5SSzBBQUNBaVJvREZ2SE4zamQ2QkZDY2QxYUcvRTFjR1hJd3NXTUM4aVJvQVFEQVRyMjR2NzlYWVFBQW1CQUJDNERpM1lVUWprMnorRlBURFhHU3hiK25janhBRWY1WDM5Yit6Z0lBc0JNbVdnQUF3RVFJV0FETXdxMlZJWC9WZE1QaGVQMEQyS1k0MWNMa0lBQUFka0xRQWdBQUVxdldxMDhQbUg3UkM0Qml4WURGY25ONjhWYUwvOGJLRUdBWEZvSVdBQURzaXFBRkFBQWtNZ1lzenNkL0hqQUJsRW5BNGp1YWJuZ2RRbmcxMlFNRWNsYnBIZ0FBdXlKb0FRQUFleVpnQVRBTEFoWS9NSzRNVVI5Z1Z4WXFDd0RBcmdoYUFBREFuZ2hZQU15Q2dNWERMVjBQZ1IwNmFMcmh1Ry9yRzBVR0FHRGJCQzBBQUdBUHF2WHEzQU1sZ0tJSldEeEMwdzN4VGZOZnNqbGdJRmNMazNNQUFOZ0ZRUXNBQU5paGFyMDZHd01XUitvTVVDUUJpMGV5TWdUWW8wcXhBUURZQlVFTEFBRFlBUUVMZ09JSldEemR1ZXNqc0NlQ0ZnQUE3SVNnQlFBQWJKR0FCVUR4QkN5ZW9lbUcrTkR6SXRzVEFISnpvbU1BQU95Q29BVUFBR3hCdFY2OURpRmNDbGdBRkV2QVlqdlVEOWlycGhzV2ZWdGZxVG9BQU5za2FBRUFBTTlRclZlTGNZS0Z0K1VBeWlSZ3NTVk5OOFNWSVMrTE9Ca2dKM0dTanFBRkFBQmJKV2dCQUFCUElHQUJVRHdCaXkxcXV1Rnd2RzRDN0Z1bDRnQUFiSnVnQlFBQVBJS0FCVUR4QkN4Mkk2N1hPaWp4eElESlcyZ1JBQURiOXVMKy9sNVJBUURnQndRc0FJcjNJWVlCQkN5MnIrbUcrRGI1L3kzdHZJQ3MvSysrclQ5cUdRQUEyMktpQlFBQWZFZTFYaDJQQVlzMzZnUlFwT3R4Z29YOS9idHpXZXFKQWRtSW9lbjMyZ1VBd0xZSVdnQUF3RmNJV0FBVVQ4QmlENXB1ZUcwYUZEQUJsYUFGQUFEYkpHZ0JBQUNmRWJBQUtKNkF4WDZaWmdGTXdVSVhBQURZSmtFTEFBQVFzQUNZQXdHTFBXdTZJVjVYajJaMTBzQlVWVG9EQU1BMnZiaS92MWRRQUFCbXExcXZEa01JNStPL0E1OEVnT0lJV0NUUWRFTU1NRzVjVzRFSitWZmYxaHNOQVFCZ0cweTBBQUJnbGdRc0FJb25ZSkhXMHZVVm1KaHFESUFCQU1DekNWb0FBREFyQWhZQXhST3dTS3pwaG9WVlhNQUVXUjhDQU1EV0NGb0FBREFMQWhZQXhST3dtSTdsM0FzQVROSkNXd0FBMkJaQkN3QUFpbGV0VjJjaGhFc0JDNEFpQ1ZoTVNOTU44WnA3TXZjNkFKUDBVbHNBQU5pV0YvZjM5NG9KQUVDUnhvQkZmS3YyU0ljQmlpTmdNVEZOTjhUcFVSdlhYV0RDZnU3YjJuVURBSUJuTTlFQ0FJRGlDRmdBRkUzQVlyck9YWHVCaWF0Q0NLNGZBQUE4bTZBRkFBREZFTEFBS0pxQXhZUTEzWEFjUXJpWWV4MkF5VnVNS3dVQkFPQlpCQzBBQU1oZXRWNTkrc0hVM21XQThnaFk1TUdEU3lBSGxTNEJBTEFOTCs3djd4VVNBSUFzalFHTE9NSGlSQWNCaWlOZ2tZbW1HK0wxK0gvbVhnY2dHLy9zMi9wR3V3QUFlQTRUTFFBQXlJNkFCVURSQkN6eVk1b0ZrSk00MVVMUUFnQ0FaeEcwQUFBZ0d3SVdBRVVUc01oUTB3MW5WbmNCbVlsQmkvZWFCZ0RBY3doYUFBQXdlZFY2ZFJ4Q2VDdGdBVkFrQVl0TU5kMXdhSm9Ga0tHRnBnRUE4RnlDRmdBQVROWVlzSWdUTE43b0VrQnhCQ3p5RjYvUkIzTXZBcEFkNFcwQUFKN3R4ZjM5dlNvQ0FEQXBBaFlBUlJPd0tFRFREZkZhL2YvbVhnY2dXLy9xMjNxamZRQUFQSldKRmdBQVRJYUFCVURSQkN6SzhuYnVCUUN5Vm9VUUJDMEFBSGd5UVFzQUFKS3IxcXRQTzk0RkxBREtJMkJSbUtZYkZrYnZBNWxiQ0l3QkFQQWNnaFlBQUNRekJpek94MzkydkFPVVJjQ2lYQjVPQXJtcmRCQUFnT2NRdEFBQVlPOEVMQUNLSm1CUnNLWWI0b3F2bzduWEFjamVTeTBFQU9BNVh0emYzeXNnQUFCN0lXQUJVRFFCaThJMTNSQ3Y0emV1NFVBaGZ1N2IyalVMQUlBbk1kRUNBSUM5cU5hcnBZQUZRSkVFTE9iajBuVWNLTWdpaE9EYUJRREFrd2hhQUFDd1U5VjZkUllmd0JrekRsQWNBWXNaYWJxaENpRzhtWHNkZ0tKVTJna0F3Rk1KV2dBQXNCTUNGZ0RGRXJDWXA4dTVGd0FvemtKTEFRQjRxaGYzOS9lS0J3REExZ2hZQUJSTHdHS21tbTU0SFVMNFAzT3ZBMUNrZi9adGZhTzFBQUE4bG9rV0FBQnNoWUFGUUxGK2k5TU1CQ3ptcWVtR1E5TXNnSUxGOVNHQ0ZnQUFQSnFnQlFBQXoxS3RWNHN4WUhHaWtnQkZlVGRPc1BBQWF0N09oU2lCZ3NYdk11ODFHQUNBeHhLMEFBRGdTUVFzQUlvbFlNSHZtbTQ0SG9NV0FLV3FkQllBZ0tjUXRBQUE0RkVFTEFDS0pXREJsK0wxL2tCVmdJTDVUZ01Bd0pPOHVMKy9WemtBQUg1SXdBS2dXQUlXL0UzVERmRzYvejhxQTh6QXYvcTIzbWcwQUFDUFlhSUZBQURmVmExWHgyUEE0bzFLQVJSRndJTHZXYW9PTUJNeFdDWm9BUURBb3doYUFBRHdWUUlXQU1VU3NPQzdtbTQ0TThFS21KRktzd0VBZUN4QkN3QUEva0xBQXFCWUFoYjhVTk1OaHlHRVM1VUNaa1RRQWdDQVJ4TzBBQURnZHdJV0FNVVNzT0F4emtNSUJ5b0d6TWpMR0RMcjIvcWpwZ01BOEZDQ0ZnQUFNMWV0VjRmalE1V0x1ZGNDb0RBQ0ZqeEswdzNIN2dlQW1ZcFRMYTQwSHdDQWh4SzBBQUNZcWM4Q0Z0NWNCU2lMZ0FWUFpXVUlNRmNMUVFzQUFCNUQwQUlBWUdZRUxBQ0tKV0RCa3pYZEVCOHl2bEpCWUtZV0dnOEF3R01JV2dBQXpJU0FCVUN4QkN6WUJ0TXNnRG1yZEI4QWdNZDRjWDkvcjJBQUFJV3IxcXNZcmxnS1dBQVVSY0NDcldpNklkNG4vRnMxZ1puN1o5L1dycWtBQUR5SWlSWUFBQVdyMXF1ek1XQnhwTThBeFJDd1lHdWFiamdjN3hVQTVpNnVEM2s3OXlJQUFQQXdnaFlBQUFVU3NBQW9rb0FGdTJEaUZjQWZyQThCQU9EQkJDMEFBQW9pWUFGUUpBRUxkcUxwaHVNUXdpK3FDL0E3UVFzQUFCNU0wQUlBb0FEVmV2VnB6SzJBQlVBNUJDellOU1B5QWY1MG9oWUFBRHpVaS92N2U4VUNBTWpVR0xCWStsRVFvQ2dDRnV4YzB3MnZRd2ovUjZVQi91TG52cTJ2bEFRQWdCOHgwUUlBSUVNQ0ZnQkZFckJnbnk1VkcrQnY0dm9RUVFzQUFINUkwQUlBSUNNQ0ZnQkZFckJncjVwdVdGbzNCdkJWbGJJQUFQQVFnaFlBQUJtbzFxdHFmUE5Vd0FLZ0hBSVc3RjNURFljaGhIT1ZCL2lxaGJJQUFQQVFnaFlBQUJOV3JWZkg0d1NMTi9vRVVBd0JDMUtLd2MwREhRRDRxcU1ZU092YitxUHlBQUR3UFlJV0FBQVRKR0FCVUNRQkM1SnF1cUZ5YndId1EzR3F4WHRsQWdEZ2V3UXRBQUFtUk1BQ29FZ0NGa3pGcFU0QS9GQWxhQUVBd0k4SVdnQUFURUMxWGgyT0R6OEVMQURLSVdEQlpEVGRjQlpDT05FUmdCOWFLQkVBQUQ4aWFBRUFrTkFZc0RnZi85bVhEbEFHQVFzbXBlbUd3M0ZpRmdBL1Zxa1JBQUEvSW1nQkFKQ0FnQVZBa1FRc21LcDR2M0drT3dBUGN0QjBROVczOVVhNUFBRDRGa0VMQUlBOUVyQUFLSktBQlpQVmRNUHhlTjhCd01QRnFSYUNGZ0FBZkpPZ0JRREFIZ2hZQUJUbkxvVHdOb1J3S1dEQnhDM2Rld0E4bXZVaEFBQjhsNkFGQU1DT1ZldlYyZmlRdzhodWdQekZnTVhsR0xENHFKOU1XZE1OaXhEQ0cwMENlTFNGa2dFQThEMkNGZ0FBT3lKZ0FWQVVBUXR5ZEtsckFFL3lVdGtBQVBpZUYvZjM5d29FQUxCRkFoWUFSUkd3SUV0Tk44VDdrZi9XUFlBbis3bHY2eXZsQXdEZ2EweTBBQURZRWdFTGdLSUlXSkN0cGhzT1RiTUFlTGE0UGtUUUFnQ0FyeEswQUFCNHBtcTlXb3dQTTR5WEJjaWZnQVVsT0E4aEhPZ2t3TE5VeWdjQXdMY0lXZ0FBUE5FWXNJZ1RMRTdVRUNCN0FoWVVvZW1HNHhEQ2hXNENQSnVnQlFBQTN5Um9BUUR3U0FJV0FFVVJzS0EwYjNVVVlDdU9Zbml0YitzYjVRUUE0RXVDRmdBQUR5UmdBVkFVQVF1SzAzVER3bjBLd0ZiRnFSYUNGZ0FBL0kyZ0JRREFEMVRyMWZINE1PNlZXZ0ZrVDhDQ2twbG1BYkJkTVdqeFhrMEJBUGlTb0FVQXdEZU1BWXM0d2VLTkdnRmtUOENDb2pYZGNCN0gzT3N5d0ZZdGxCTUFnSzhSdEFBQStJS0FCVUJSQkN3b1h0TU5oK085Q3dEYlpSMFRBQUJmSldnQkFEQVNzQUFvaW9BRmN4THZYdzUwSEdEN21tNm8rcmJlS0MwQUFKOFR0QUFBWnE5YXJ6NjlCZnJMM0dzQlVBQUJDMllsUGdCMER3T3dVL0h2cktBRkFBQi9JV2dCQU16V0dMQTRILzk1Q3hRZ2J3SVd6TldsemxPNDYvRkJ0L3QxVWxtRUVONnFQZ0FBbnhPMEFBQm1SOEFDb0NnQ0ZzeFcwdzJ2UXdnblBnRVU3bno4TysrelRpcVZ5Z01BOENWQkN3QmdOZ1FzQUlvaVlBR21XVkMrWC91MjNqVGRjQ1ZvUVVJdm0yNDQ3TnZhL1FZQUFQOGhhQUVBekVLMVhzVnd4VkxBQWlCN0FoYnd4elNMZUY5enBCWVU3RzY4ZjQ5aTBPSkNzMG1vR2orSEFBRHdPMEVMQUtCbzFYcDFOdjVBNjBFRVFONEVMR0FVMzZ3ZUozUkJ5WmFmVFJEWTZEU0pMUVF0QUFENG5LQUZBRkFrQVF1QVlnaFl3TjlkbXRKRjRXNzd0djdQYXB3WXVHaTY0VU5jNGFEeEpGSXBQQUFBbnhPMEFBQ0tJbUFCVUF3QkMvaUtwaHZpVzlWdjFJYkNuWDNsOURhQ0ZpUzBVSHdBQUQ0bmFBRUFGS0ZhcjE2UEQrUUVMQUR5Sm1BQjM3ZFVId3IzVzkvV1gxdlJzQkV5SXFHRHBodU8rN2ErMFFRQUFJS2dCUUNRdTJxOVdvd1BIRTQwRXlCckFoYndBMDAzbkxubllRYk92M0dLWHd0ZndEN0Y5U0dDRmdBQS9FN1FBZ0RJa29BRlFERUVMT0FCbW00NE5NMkNHVmg5YTJKQTM5YWJwaHQ4QmtncGZnZDlyd01BQUFSQkN3QWdOd0lXQU1VUXNJREhPYmNpamNKOXVpNTh6N1h2QVNSVUtUNEFBSjhJV2dBQVdSQ3dBQ2lHZ0FVOFV0TU54OTlacHdDbE9PL2Ira2ZYaFN2ZkIwakladzhBZ1A4UXRBQUFKcTFhcjQ3SGdNVWJuUUxJbW9BRlBGMzhmK2RBL1NqWWg3NnQzejdnOURZK0JLVFVkRU1WMTlob0FnQUFnaFlBd0NRSldBQVVROEFDbnFIcGhqalY2NVVhVXJpSFRtenhnSnZVRmo2SEFBQUVRUXNBWUdvRUxBQ0tJV0FCMjNHcGpoVHVYZC9XVnc4NXhiNnRiNXB1dUEwaEhQbFFrRWlsOEFBQUJFRUxBR0FxQkN3QWluRTcvajEvTDJBQno5TjB3MWtJNGFVeVVyQzc4WnJ4R0J0QkN4SmFLRDRBQUVIUUFnQklyVnF2RHNkUndlZDJqd05rN2ZlQXhlYjA0aUU3OW9FZmFMcmgwRFFMWnVBeVRxbDQ1R2xlV2FkRFFrZng3M1BmMXNLa0FBQXpKMmdCQUNRaFlBRlFEQUVMMkkybGV5UUtkL3ZFTU5IR0I0UEVxakh3QXdEQWpBbGFBQUI3SldBQlVBd0JDOWlScGh2aVNyVmYxSmZDTFo4eUZhQnY2NnVtRzN3MlNHa2hhQUVBZ0tBRkFMQVhBaFlBeFJDd2dOM3oveGVsdSs3YitqbWY4dzhoaEpjK0pTU3lVSGdBQUFRdEFJQ2RxOWFyczNFc3NJQUZRTDRFTEdBUG1tNklEL0JPMUpyQ25UL3o5SzRFTFVpb1Vud0FBQVF0QUlDZEdRTVdjYi80a1NvRFpFdkFBdmJMLzJ1VTdsM2YxcHRubnVOei8zdDRqb080NHFsdjZ4dFZCQUNZTDBFTEFHRHJCQ3dBaWlCZ0FYdldkTU81K3ljS2Q3ZUZhUlpobkdnQktTMEU0d0FBNWszUUFnRFlHZ0VMZ0NJSVdFQUNUVGNjanZkUlVMTEx2cTAvUHZmODRpU0JwaHZ1ckNZa0lldERBQUJtVHRBQ0FIaTJhcjJLYi9OYzJwTU1rRFVCQzBqcjBrTmpDbmZidC9VMncwUnhmY2lKRHcySkNGb0FBTXljb0FVQThHUmp3R0xwQjA2QXJBbFlRR0pOTjhRSGRtLzBnY0p0WTJYSTU2NThEeUVobnowQWdKa1R0QUFBSGszQUFxQUlBaFl3SFpkNlFlR3UrN1orditWVGpFR0xDeDhjVW1tNllkRzM5WlVHQUFETWs2QUZBUEJnQWhZQVJSQ3dnQWxwdXVHMWV5dG00R3dIcDdqeHdTR3hhZ3o4QUFBd1E0SVdBTUFQVmV2VmNRamhyWWNBQUZrVHNJQnBNczJDMHYzYXQvWE50cyt4Yit1UFRUZDhDQ0c4OUFraWtZVy80UUFBOHlWb0FRQjgweGl3V05vWkRwQTFBUXVZcUtZYjRuM1drZjVRc0x2eCs4U3ViQVF0U0toU2ZBQ0ErUkswQUFEK1JzQUNvQWdDRmpCaFRUZkUrNjF6UGFKd3l6aDVZb2VuZU9VN0N3a2ROZDF3Y3Fza2hRQUFJQUJKUkVGVXVPUFBPQUFBRXlWb0FRRDhoNEFGUUJFRUxDQVA4WjdyUUs4bzJJZStyWGU5Vm1IakEwUmljWDNJZTAwQUFKZ2ZRUXNBSUFZc0RzZmRzZ0lXQVBrU3NJQk1OTjJ3Y04vRkRPeDhZa3ZmMXB1bUczeVdTS2tTdEFBQW1DZEJDd0NZc1RGZ2NUNys4MFlsUUo0RUxDQS9TejJqY0wvMWJYMjFwMU84RGlHYytFQ1J5RUxoQVFEbVNkQUNBR1pJd0FLZ0NBSVdrS0dtRzg0OEZHWUdkajdONGpOWC9wOGlvVXJ4QVFEbVNkQUNBR1pFd0FLZ0NBSVdrS21tR3c1TnMyQUdWbjFiMyt6eE5EYytWQ1IwMEhSREZkZllhQUlBd0x3SVdnREFURlRyMVZMQUFpQnJBaGFRdjNndmRxU1BGT3d1aEhDNTU5UGIxNG9TK0paSzRBY0FZSDVlM04vZmF6c0FGS3hhcjg3R055ZjlxQStRSndFTEtFRFREY2NoaFArbmx4VHV2L3EyM3Z2MXF1bUdHOTkzU09oZDM5Wm5HZ0FBTUM4bVdnQkFvUVFzQUxJbllBRmwyZmRiL3JCdjF5bENGcU9ON3owa1ZDaytBTUQ4Q0ZvQVFHRUVMQUN5SjJBQmhXbTZZUkZDZUtXdkZHNlo4UFN1L0Q5R1FpOFZId0JnZmdRdEFLQVFBaFlBMlJPd2dIS1paa0hwNHVxRXE0VG51UEVKSTZVWXFFdjgvd0FBQUhzbWFBRUFtYXZXcThVWXNEalJTNEFzQ1ZoQXdacHVPUGUyTTRXN1N6ek5Jc1FIM0UwMytKeVIwbUtjckFJQXdFd0lXZ0JBcGdRc0FMSW5ZQUdGYTdyaE1QVURhTmlEeTc2dGJ5WlE2R3ZmalVpb1Vud0FnSGtSdEFDQXpBaFlBR1JQd0FMbUk5NnpIZWczQmJ1ZDBHcWNqZTlJSkNSb0FRQXdNNElXQUpBSkFRdUE3RjJQQVF0anBXRUdtbTQ0RGlIOG90Y1U3cnh2NjQ4VE9jWE5CSTZCK1RxS2YvY25NdDBGQUlBOUVMUUFnSW1yMXF2ak1XRHhScThBc2lSZ0FmTmthZzJsdSs3Yit2MkV6dEYxbHRUaVZBdEJDd0NBbVJDMEFJQ0pFckFBeUo2QUJjeFUwdzBMVThpWWdmTXBuV0tjSk5CMHc1MTFQU1FVZ3haVENoOEJBTEJEZ2hZQU1ERUNGZ0RaRTdBQVRMT2dkTy82dHA3aXFvNTQ3WDAxZ2VOZ25oYjZEZ0F3SDRJV0FEQVJBaFlBMlJPd0FPSTBpM2cvZDZRU0ZPeHVhdE1zUHJNUnRDQWhrNHdBQUdaRTBBSUFFcXZXcThQeGg4b0x2UURJa29BRjhMdW1HdzRuL0FBYXR1V3liK3VQRTYzbWxlOVZwTlIwUXpYUmFTOEFBR3lab0FVQUpQSlp3T0xjSG1HQUxBbFlBRis2ZEY5SDRXNzd0bDVPOVJUN3RyNXF1bUVDUjhLTUxjYkpLZ0FBRkU3UUFnRDJUTUFDSUhzQ0ZzRGZ4TGVZcllCakJzNHlPTVVQSVlTWEV6Z081cW5TZHdDQWVSQzBBSUE5RWJBQXlKNkFCZkE5bDZwRDRhN2p4SWdNVG5FamFFRkNnaFlBQURQeGswWUR3TzVWNjFVTVY5eU0rNEtGTEFEeUVnTVdQMjlPTHhaQ0ZzRFhOTjN3T29Sd29qZ1VMb2RwRnBGck5TbTliTHJoVUFjQUFNcG5vZ1VBN0ZDMVhzVWZJK01PNHlOMUJzaU9DUmJBRDQwUDFFeXpvSFMvOW0xOWs4azViaVp3RE14YkpmQURBRkErUVFzQTJBRUJDNENzQ1ZnQWozSHVuby9DM1kzZmJiTFF0L1dtNllZN2t3UkphQ0ZvQVFCUVBrRUxBTmdpQVF1QXJBbFlBSS9TZE1QeEdMU0FrcDMzYmYweHMvUGJXT2REUXBYaUF3Q1VUOUFDQUxhZ1dxL2lHeXR2QlN3QXNpUmdBVHpWMGx2ekZPNUQzOVp2TXp6RkswRUxFbG9vUGdCQStRUXRBT0FaeG9ERjBvOTRBRmtTc0FDZXJPbUdlQi80UmdVcFhLNFRXellUT0FibTZ5Qk9QT3JiK3NabkFBQ2dYSUlXQVBBRUFoWUFXUk93QUxaaHFZb1U3cmUrclhPOVZyckdrMXBjSHlKb0FRQlFNRUVMQUhnRUFRdUFyQWxZQUZ2UmRNT1orMEZtSU5kcEZxRnY2NDlOTjl4YTdVaEM4YmVEOXhvQUFGQXVRUXNBZUlCcXZZcHZvMXo2UVIwZ1N3SVd3TlkwM1hBNDNoZEN5VllGckQyNHN0NkhoQ3JGQndBb202QUZBSHhIdFY0ZGp4TXMvRUFIa0I4QkMyQVg0bHYrQnlwTHdXNExDUk50Zkk4aklTOXBBQUFVVHRBQ0FMNUN3QUlnYXdJV3dFNDAzUkR2RVM5VWw4SXQ0K3FOQWs1eE00RmpZTWFhYmxqMGJlMStGQUNnVUlJV0FQQVpBUXVBckFsWUFMdG1aUWlsdSs3YittMEo1eGdmY0RmZE1JRWpZY2FxY1lVTkFBQUZFclFBZ0Q4Q0ZwOTJiUXRZQU9SSHdBTFl1ZmhtY2dqaGxVcFR1R1ZocDNkdGhRTUpWWW9QQUZBdVFRc0FabTBNV0p6YnRRMlFKUUVMWUorS2VNc2Z2dU5kZ1dzT05vSVdKTFJRZkFDQWNnbGFBREJMQWhZQVdST3dBUGFxNllaNHozaWs2aFRzcnNCcEZtRmMyL0RMQkk2RGVUcHF1dUd3Yit1UCtnOEFVQjVCQ3dCbVJjQUNJR3NDRnNEZXhZZGtoVDZBaHM5ZDltMTlVMkJGTmhNNEJ1YXRHZ00vQUFBVVJ0QUNnRmtRc0FESW1vQUZrTkxTL1NPRnUrM2J1c2d3VVF5UE5OMXc1LzloRWxvSVdnQUFsRW5RQW9EaVZldlYyZmdEdVhIUEFIa1JzQUNTYXJyaDJOb0JadUM4OEZPTTl4R3ZKbkFjek5OQzN3RUF5aVJvQVVDeEJDd0FzaVZnQVV6Rlc1MmdjTmQ5Vzc4di9CdzNnaFlrVkNrK0FFQ1pCQzBBS0k2QUJVQzJmb3M3NGdVc2dDbG91dUYxQ09GRU15aGM2ZE1zd2pqUjRtSUN4OEU4SGNUcFNIR05qZjREQUpSRjBBS0FZZ2hZQUdUcjNUakJ3Zy9Rd0pSYzZnYUZlOWUzOWFiMGsremIrcXJwaGdrY0NUTzJNQ0VKQUtBOGdoWUFaSzlhcnhiakQrRXZkUk1nS3dJV3dDUTEzU0M4UytudVpqTE40cE1QdmkrU2tQVWhBQUFGRXJRQUlGdGp3R0pwcEROQWRnUXNnTWxxdXVGd1pnK2dtYWRsMzlZZlozVG1WNElXSkxSUWZBQ0E4Z2hhQUpBZEFRdUFiQWxZQURtSWs5SU9kSXFDM2ZadFBiZlZPTVd2U0dIU2hId0FBQW9rYUFGQU5nUXNBTElsWUFGa29lbUdPTjc5alc1UnVMTVpObGpRZ3FTYWJsajBiWDJsQ3dBQTVSQzBBR0R5cXZYcWVIeXo4SlZ1QVdSRndBTEl6ZHplOG1kK3J1ZjRzTGR2NjAzVERYZW0xWkJRTmE2d0FRQ2dFSUlXQUV6V0dMQlllcXNRSURzQ0ZrQjJtbTQ0TXptTkdaampOSXRQTnY0Zko2R0ZNQjhBUUZrRUxRQ1lIQUVMZ0d3SldBQlphcnJoY0x6L2hKTDkycmYxbksvUlY0SVdKRlFwUGdCQVdRUXRBSmdNQVF1QWJBbFlBTGs3RHlFYzZTSUZ1eE1tK2oxb2NUR0I0MkNlam1Lb3IyL3JqL29QQUZBR1FRc0FrcXZXcTA5dkVQNmlHd0JaRWJBQXN0ZDB3L0VZdElDU25YdkErL3ZxRUVncHJnOTVyd01BQUdVUXRBQWdtVEZnY1Q3K085QUpnR3dJV0FBbHVYUXZTdUUrOUczOWR1NU5qa0dUcGh0dVRhOGhvVXJRQWdDZ0hJSVdBT3lkZ0FWQXRnUXNnS0kwM1JEZkxuNmxxeFRPeEpZL1hWbFZTVUlMeFFjQUtJZWdCUUI3STJBQmtDMEJDNkJVbHpwTDRYN3IyL3BLay85akkyaEJRcFhpQXdDVVE5QUNnTDJvMXF1bGdBVkFkZ1FzZ0dJMTNYQVdRbmlwd3hUc3pqU0x2eEU2SWFXRHBodXF2cTAzdWdBQWtEOUJDd0IycWxxdjRnL1lTM3R3QWJJaVlBRVVyZW1HUTlNc21JSEx2cTFkeXo4VEgzQTMzVENaNDJHV3FuR3lDZ0FBbVJPMEFHQW5CQ3dBc2lSZ0FjeUZTV3VVN2xhWTZKdXVRd2duRXowMnlyY0lJYnpWWndDQS9BbGFBTEJWQWhZQVdSS3dBR2FqNlliakVNS0ZqbE80WmQvV0h6WDVxemFDRmlSVUtUNEFRQmtFTFFEWWltcTllajIrTVNWZ0FaQVBBUXRnanJ4SlRPbXUrN2IyT2YrMnF4RENMMU05T0lyM1Vvc0JBTW9nYUFIQXMxVHIxV0tjWU9HTklJQjhDRmdBczlSMHc4SjlLek93MU9UdjJrejQySmlCZUMzcTIvcEtyd0VBOGlab0FjQ1RDRmdBWkVuQUFwZzdiL2xUdW5jZTRINWYzOVkzVFRmY21zWklRb3R4c2dvQUFCa1R0QURnVVFRc0FMSWtZQUhNWHRNTjV4NnNVcmk3RU1LNUpqL0l4dDhERXFvVUh3QWdmNElXQUR5SWdBVkFsZ1FzQVA0SVdSeGFwOEFNWFBadC9WR2pIeVFHTFY1bGNKeVVTZEFDQUtBQWdoWUFmRmUxWGgyUFAwcS9VU21BYkFoWUFQeFZ2Sjg5VUJNS2R0dTN0VERSdzhXMURSZTVIQ3pGT1dxNjRUaXVzZEZhQUlCOENWb0E4RlVDRmdCWkVyQUErRUxURGZITjRWL1VoY0paR2ZJSWZWdGZOZDJRemZGU3BIaHRjczhPQUpBeFFRc0Eva0xBQWlCTEFoWUEzM2FwTmhUdXVtL3I5NXI4YUI5Q0NDOHpPMmJLRWRleit2OFdBQ0JqZ2hZQS9FN0FBaUJMQWhZQTM5RjB3K3NRd29rYVVUalRMSjdtU3RDQ2hDckZCd0RJbTZBRndNeFY2OVhoK01QY3ViM1ZBTmtRc0FCNEdOTXNLTjJ2ZlZ0dmRQbEoxSTJVaEFBQkFESW5hQUV3VXdJV0FGa1NzQUI0b0tZYjRyUzJJL1dpWUhmalZFS2U1a3JkU0tucGhrcFFDZ0FnWDRJV0FETWpZQUdRbmJ0eGY3T0FCY0FETmQxd2FKMENNN0RzMi9xalJqOU4zOVkzVFRmYytWNU1RZ3VUVlFBQThpVm9BVEFUQWhZQTJia2JSOTVmYms0dlBFUUJlSnhMOTd3VTdyWnZhNnR4bm05amhRTUpWWW9QQUpBdlFRdUFHYWpXcXpNL05nTmtROEFDNEJtYWJvaHZDTDlSUXdwM3BzRmJjU1ZvUVVLQ0ZnQUFHUk8wQUNqWUdMQ3dteG9nRHdJV0FOdXhWRWNLZDkyMzlaVW1iMFdzNDBVQjUwR2VYc1pWVjFZQUFRRGtTZEFDb0VBQ0ZnQlpFYkFBMkpLbUc4NjhuYzRNbUdheFBadFNUb1JzVldQZ0J3Q0F6QWhhQUJSRXdBSWdLd0lXQUZzVTN3bzJ6WUlaV1BWdGZhUFIyeEVuQ1RUZDhDRk9GaWpoZk1qU1F0QUNBQ0JQZ2hZQUJhaldxL2pGL0syQUJVQVdCQ3dBZHVQYy9UQ0YrM1FQd1hadEJDMUlxRko4QUlBOENWb0FaR3dNV0N5TlJ3Yklnb0FGd0k0MDNYQWNRcmhRWHdwM0hpY3dhUExXeGFERm04TE9pWHdzOUFvQUlFK0NGZ0FaRXJBQXlJcUFCY0R1ZWN1ZjBuM28yL3F0THUrRXRRMmtkQkREZ2xZQ0FRRGtSOUFDSUNNQ0ZnQlpFYkFBMklPbUcrSTk4aXUxcG5Ebkdyd2JmVnR2bW00bzhkVEl4NmQxc0FBQVpFVFFBaUFEMVhwMVBIN3BGckFBbUQ0QkM0RDlNczJDMHIzcjI5clVoZDI2OW4yYmhDckZCd0RJajZBRndJU05BWXVsZmJFQVdSQ3dBTml6cGh2T1FnZ3YxWjJDM1kzZkNkbXRLMEVMRWhLMEFBRElrS0FGd0FRSldBQmtSY0FDSUlHbUd3NU5zMkFHTHZ1MnZ0SG9uZHNVZm41TW01QVBBRUNHQkMwQUprVEFBaUFyQWhZQWFjWDc1Z005b0dDM3drUjdJMmhCVWswM0xLd0lBZ0RJaTZBRndBUlU2OVdudC9FRUxBQ21UOEFDSUxHbUcySkErUmQ5b0hETHZxM2RhK3hCbkJyU2RFTU10aHdWZjdKTVZUV3VzQUVBSUJPQ0ZnQUpqUUdMOC9HZnQvRUFwazNBQW1BNjN1b0ZoYnZ1MjlybmZMODJnaFlrVkNrK0FFQmVCQzBBRWhuWGhHd0VMQUFtVDhBQ1lFTGllSFg3N0ptQmMwM2V1emhONE5YTXpwbnBXT2dGQUVCZWZ0SXZnRFEycHhjM0lZUWI1UWVZckJpd1dJVVFqamVuRjBzaEM0REo4SlkvcFh2WHQvVkdsL2RPelVucHFPbUdReDBBQU1pSG9BVkFXbjRrQnBnZUFRdUFpV3E2NGR4b2Z3cDNaNXBGR24xYlg4M3h2SmtVNjBNQUFESmlkUWhBV3U5RENQL1dBNEJKc0NJRVlNTEdOMzJYZWtUaEx2dTJkaCtTenJYVlJDUzBHRmZZQUFDUUFSTXRBQklhMTRkODBBT0FwRXl3QU1oRERNTWQ2QlVGdSszYldwZ29MZXREU0dtaCtnQUErVERSQWlDOXQ2WmFBQ1JoZ2dWQUpwcHVpT1BVMytnWGhiTXlKRDFCQzFLeU9nUUFJQ01tV2dDazkxNFBBUGJLQkF1QS9GenFHWVc3N3R2YWQ4UDByRzBncFlNeFdBZ0FRQVlFTFFBU3N6NEVZRzhFTEFBeTFIVEQ2eERDaWQ1UnVETU5UcTl2NjV2eG5oRlNFYlFBQU1pRTFTRUEwMkI5Q01EdVdCRUNrRGZUTENqZHIrTURmcVloVHJWNHBSY2tJbWdCQUpBSlFRdUFhWGd2YUFHd2RRSVdBSmxydW1FWlFqalNSd29XNzFlV0dqd3BHMEVMRWxvb1BnQkFIcXdPQVppQWNYM0l0VjRBYklVVklRQUZhTHJoT0lSd3JwY1VidG0zdFh1VmFibWFld0ZJNnFYeUF3RGtRZEFDWURyZTZnWEFzd2hZQUpRbHZ1Vi9vS2NVN0xadmE2dHhwbWN6OXdLUVZ0TU5wbG9BQUdSQTBBSmdPdDdyQmNDVENGZ0FGR1o4eVBSR1h5bmNtUVpQenpoaDVNUGM2MEJTbGZJREFFeWZvQVhBUkl3UEJuL1RENEFIdXcwaC9HOEJDNEFpTGJXVnd2M1d0N1VWRmRObHFnVXBtV2dCQUpDQmYyZ1N3S1RFcVJhdnRBVGd1MkxBSWdZcnJGd0NLRkRURGZFdC94TzlwWERuR2p4cFY2YnFrSkNKRmdBQUdURFJBbUJhckE4QitMWVlzUGl2emVuRnNaQUZRSm1hYmpnTUlWeHFMNFZiOVcxOW84bVRacUlGS1IyTjEwTUFBQ1pNMEFKZ1Fxd1BBZmdxQVF1QStZaHYrUi9vTndXN0V5YWF2cjZ0QlMxSXpmb1FBSUNKc3pvRVlIcXNEd0g0Z3hVaEFEUFNkTU54Q09GQ3p5bmNlZC9XSHpVNUM5ZldHSkZRWmVvcEFNQzBDVm9BVEUvOEl2M2YrZ0xNbUlBRndEeDV5NS9TWGZkdDdmNG1IMWVDRmlSa29nVUF3TVJaSFFJd01kYUhBRE5tUlFqQVREWGRzRERWalJsWWFuSldyQThoSlNFZkFJQ0pFN1FBbUNZUEdJRTVFYkFBd0RRTFN2ZXViK3NyWGM2S2ZwRlUwdzJWRGdBQVRKZWdCY0FFYlU0djR2cVFPNzBCQ2lkZ0FVQjhrSFFlUW5pcEVoVHN6alNML1BSdC9YRzhYNFZVQkMwQUFDYnNINW9ETUZreGJQRkdlNEFDeFIrc2w4SVZBRFRkY09nQk5ETncyYmYxalVabkthNFBPWnA3RVVobVllSXBBTUIwQ1ZvQVRKZWdCVkFhQVFzQXZoUkRGZ2VxUXNGdXJjYkpXbHdmOG1ydVJTQVpFeTBBQUNiTTZoQ0FpYkkrQkNpSUZTRUEvRTNURGNjaGhGOVVoc0tkanlzb3lOTkczMGpJV2kwQWdBa1R0QUNZdHZmNkEyUk13QUtBNzNGdG9IVFhmVnY3VHBleHZxMnY1bDREMG1xNllhRUZBQURUWkhVSXdMUlpId0xreUlvUUFMNXJmSEIwb2tvVTdseURpM0R0N3hVSkxjWVZOZ0FBVEl5SkZnQVRabjBJa0JrVExBQjRLTmNKU3ZldWIydHJKOHFnajZSVXFUNEF3RFNaYUFFd2ZhWmFBRk5uZ2dVQUQ5WjB3ektFY0tSaUZPek9OSXVpeEdrQ3Y4eTlDQ1FqYUFFQU1GRW1XZ0JNbndlWHdGU1pZQUhBb3pUZGNPZ0JORE53MmJmMVI0MHVob2tXcEhUVWRNT3hEZ0FBVEkrZ0JjREViVTR2cnNhSG1RQlRJV0FCd0ZOZGhoQU9WSStDM2ZadHZkVGdjdlJ0ZldPbEo0bVphZ0VBTUVHQ0ZnQjVlSzlQd0FRSVdBRHdaRTAzVkZiaU1RTm5tbHlrcTdrWGdLUVd5ZzhBTUQzLzBCT0FMTHkxRXhaSUtBWXNsc0lWQUR6VHBRSlN1T3UrclQyUUwxTmNIL0pxN2tVZ0dSTXRBQUFteUVRTGdBeHNUaTgyMW9jQUNaaGdBY0JXTk4zd09vUndvcG9VempTTGNnblFrSkxySndEQUJBbGFBT1REK2hCZ1h3UXNBTmlhcGhzT1RiTmdCbjd0Mi9wR284dGtVZ21wamV1M0FBQ1lFS3REQVBKaGZRaXdhMWFFQUxBTDV5R0VJNVdsWUhmeEhrcURpL2NoaFBCeTdrVWdtY1c0d2dZQWdJa3cwUUlnRTlhSEFEdGtnZ1VBTzlGMHcvRVl0SUNTTGZ1Mi9xakR4Zk9RbTVSTXRBQUFtQmdUTFFEeTh0NVVDMkNMVExBQVlOZmlXLzRIcWt6QlB2UnRiVFhPUE1UMUlXL21YZ1NTRWJRQUFKZ1lFeTBBOHVJSFBHQWJUTEFBWU9lYWJsaDRLTWtNbU5neUh5WmFrTkxMcGhzT2RRQUFZRG9FTFFBeXNqbTl1Qm4zd2dJOGhZQUZBUHNrSkV6cGZ1dmIra3FYNTZGdjZ4aTB1SnQ3SFVqS1ZBc0FnQW14T2dRZ1AvSGg2TC8xRFhnRUswSUEyS3VtRzg3aTI3ZXFUdUZNczVpZkdMWTRtWHNSU0dZeHJyQUJBR0FDQkMwQTh2TmUwQUo0SUFFTEFQWnVIRzF1bWdXbFcvVnRmYVBMczNNbGFFRkNDOFVIQUpnT3EwTUFNbU45Q1BBQTF5R0UvOCtLRUFBU2lXLzVIeWcrQmJzVEpwcXR6ZHdMUUZKV2h3QUFUSWlnQlVDZVBEZ0Z2aVlHTEg3ZW5GNHNOcWNYNzFVSWdIMXJ1dUU0aEhDaDhCVHV2Ry9yajVvOFM5WTJrTkxCZUowRkFHQUNCQzBBOHVRQkt2QzV6d01XZnZ3RklDVnYrVk82Njc2dEJkOW5hZ3pZM002OURpUmxmUWdBd0VRSVdnQmt5UG9RWUNSZ0FjQmtOTjBRSC82ODBoRUt0OVRnMlhQZlRVcldod0FBVE1RL05BSWdXL0V0cW45ckg4eFNERmdzaFNzQW1CaHYrVk82ZDMxYnUvOWlFMEo0TS9zcWtJcWdCUURBUkFoYUFPVHJ2YUFGekk2QUJRQ1QxSFREZVFqaFNIY28ySjFwRm96Y2k1UFNpZW9EQUV6RGkvdjdlNjBBeUZTMVhyMDNuaGxtUWNBQ2dNbHF1dUV3aEJCWDJ4M29FZ1ZiOVcwdGFNSHZtbTd3Z3lvcC9XeTZEZ0JBZWlaYUFPUk4wQUxLSm1BQlFBNldRaFlVN2xiSWdpOWNteXhBUXBYSktnQUE2UWxhQU9RdEJpMytXdytoT0FJV0FHU2g2WWJqRU1JdnVrWGh6aldZTDJ3RUxVaW9VbndBZ1BSKzBnT0FmRzFPTHo2R0VIN1RRaWhHREZqOHZEbTlXQWhaQUpDSnR4cEY0YTc3dG42dnlYekJ2VG9wTFZRZkFDQTlRUXVBL1BuUkQvSW5ZQUZBZHBwdWVPMk5ibWJBTkF1K1pxTXFKSFRVZE1PaEJnQUFwQ1ZvQVpBL1FRdklsNEFGQURtNzFEMEs5NjV2YXcvVStadStyVzlDQ0xjcVEwS21XZ0FBSkNab0FaQTU2ME1nU3dJV0FHU3Q2WVpsZktOV0Z5blluV2tXL0lBUURpbFZxZzhBa05ZLzFCK2dDSEdxeFN1dGhNbUxBWXVsY0FVQU9SdkhsWHNBVGVtV2ZWdC8xR1crWStON09BbVphQUVBa0ppZ0JVQVpZdERpdi9VU0prdkFBb0NTeEpVaEJ6cEt3Vzc3dHJZYWh4K0o5L1lYcWtRaUpsb0FBQ1JtZFFoQUFjYjFJZS8wRWliSGloQUFpdEowUTN5RDlvMnVVcmd6RGVaSCtyWjJmMDlLQjAwM0NGc0FBQ1Jrb2dWQU9kNzcwUnNtd3dRTEFFcTExRmtLZCswQk9vL3dJWVR3VXNGSXBCcFgyQUFBa0lDSkZnQ0YySnhleEtERm5YNUNVaVpZQUZDc3BodmlXLzRuT2t6aFRMUGdNZHp6azVLSkZnQUFDWmxvQVZBV1V5MGdEUk1zQUNoYTB3MkhwbGt3QTcvMmJYMmowVHlDYVFLa3RGQjlBSUIwQkMwQXlpSm9BZnNsWUFIQVhKeUhFSTUwbTRMZENSUHhCSUlXcEdSdERRQkFRaS91NysvVkg2QWcxWHIxTVlSd29LZXdVd0lXQU14RzB3M0g0OE5FOTVpVTdMLzZ0bjZyd3p4VzB3MitnNVBTejMxYisxNEtBSkNBaVJZQTVUSFZBblpId0FLQU9icjBFSkhDZlJDeTRCbGlFTzFFQVVta0NpSDRmZ29Ba01CUGlnNVFuUGRhQ2xzWEF4WS9iMDR2RmtJV0FNeEowdzF4Ly9zclRhZHc1eHJNTS9oK1FFb0wxUWNBU0VQUUFxQXdtOU9MOStOK1llRDVCQ3dBbUx2THVSZUE0djFtN0Q3UDVQTkRTcFhxQXdDa0lXZ0JVQ1pUTGVCNUJDd0FtTDJtRzg1Q0NDL25YZ2VLWjVvRno3VlJRUkk2YXJyaFdBTUFBUFpQMEFLZ1RONDhoS2NSc0FDQVAwSVdoKzRwbVlGVjM5WTNHczF6OUczOU1ZUndxNGdrWktvRkFFQUNnaFlBQmRxY1htejgwQU9QSW1BQkFIOFYzL0kvVUJNS2RpdE14QmI1RGtGS2doWUFBQW44UTlFQmloWFhoL3lpdmZCZE1XQ3hGSzRBZ0QrTkk4Z3ZsSVRDTGNkSkJMQU44V1dITnlwSklndUZCd0RZUDBFTGdISzlGYlNBYnhLd0FJQnZlNnMyRk82NmIydWZjN2JKOXdwU09sRjlBSUQ5ZTNGL2Y2L3NBSVdxMXF1NGIvaElmK0UvM28wQkM3dTRBZUFybW02SWI4WCtqOXBRdUovN3R2WmduSzFxdXNHUHJLVDByNzZ0TnpvQUFMQS9KbG9BbE0zNkVQaURnQVVBUEl5My9DbmRPeUVMZHVUYVpBRVNxc1lWTmdBQTdNbFBDZzFRTkQrVU0zY3hZUEhQemVuRm1aQUZBSHhmMHczbnBxRlJ1THNRd3JrbXN5TUNQS1MwVUgwQWdQMFN0QUFvMk9iMElyN05jS3ZIekpDQUJRQThRdE1OaDNINms1cFJ1TXUrclQ5cU1qdGltZ0FwVmFvUEFMQmZWb2NBbE0vNkVPYkVpaEFBZUpvWXNqaFFPd3AyMjdlMU1CRzdKR2hCU2k5Vkh3Qmd2MHkwQUNqZnBSNHpBeVpZQU1BVE5kMVFDZVl5QTFhR3NGTjlXOStZS0VsS1RUZFlId0lBc0VlQ0ZnQ0ZHeDg2ZjlCbkNpVmdBUURQSjVoTDZhNzd0bjZ2eSt5QnFSYWtKR2dCQUxCSFZvY0F6TVBiRU1LLzlacUNXQkVDQUZ2UWRNUHJFTUtKV2xJNDB5ellsNnNRd2l2VkpwRks0UUVBOWtmUUFtQWUzZ3RhVUFnQkN3RFlMdE1zS04ydmZWdWJNc0MrK0t5Umtva1dBQUI3OU9MKy9sNjlBV2FnV3EvaUR6NHY5WnBNQ1ZnQXdKWTEzYkFNSVZ5b0t3VzdDeUVjOTIzOVVaUFpsNlliL05oS1N2L3MyOXIzWmdDQVBURFJBbUErckE4aFJ3SVdBTEFEVFRjY1c2ZkFEQ3lGTEVqZ2c1Y2NTQ2l1RC9IOUdRQmdEd1F0QU9iRCtoQnlJbUFCQUxzVnAxa2NxREVGdSszYjJtb2NVcmdTdENDaHhmajdEd0FBTy9hVEFnUE13L2pBK29OMk0zRXhZUEhQemVuRm1aQUZBT3hHMHczeEljd2I1YVZ3WnhwTUlodUZKNkZLOFFFQTlzTkVDNEI1c1Q2RXFUTEJBZ0QyWjZuV0ZPNjZiK3NyVFNZUm56MVNPbEY5QUlEOU1ORUNZRjZNajJScVRMQUFnRDFxdXVITVF4aG13RFFMa3VuYk9uNnZ1ZE1CVW1tNndWUUxBSUE5RUxRQW1KSHhRZlp2ZXM1RWZCQ3dBSUQ5YWJyaDBEUUxabUExUHVpR2xLd1BJYVdGNmdNQTdKNmdCY0Q4bUdyQlZMeXMxcXRqM1FDQXZUa1BJUndwTndXTFV3UXVOWmdKc0Q2RWxFeTBBQURZQTBFTGdQa1J0R0JLWHVzR0FPeGUwdzB4M0hpaDFCVHV2Ry9yajVyTUJBaGFrSktKRmdBQWV5Qm9BVEF6bTlPTGo5YUhNQ0gyWndQQWZuakxuOUo5Nk52NnJTNHpFVmFIa05MUnVDNE1BSUFkRXJRQW1DZFRMWmdLNjBNQVlNZWFib2h2dHI1U1p3cDNyc0ZNeFRoWjVZT0drSkQxSVFBQU95Wm9BVEJQZ2haTWlmVWhBTEJicGxsUXV0LzZ0cmFxZ2FreDFZS1VyQThCQU5neFFRdUFHYkkraElueDlpRUE3RWpURFhGTjEwdjFwV0IzN2llWktFRUxVaEswQUFEWU1VRUxnUGt5MVlLcE9LcldLMk5OQVdETHh2M3NwbGxRdXN1K3JXOTBtUWt5WllXVWZNY0dBTmd4UVF1QStYby92djBGVTNDbUN3Q3dkY3NRd29HeVVyQmJZU0ttcW05ckV5MUk2YURwaG1NZEFBRFlIVUVMZ0prYTE0ZVlhc0ZVdk5ZSkFOaWU4ZUhLTDBwSzRaWjlXMy9VWkNic1duTkl5UG9RQUlBZEVyUUFtRGRCQzZiQytoQUEySzYzNmtuaHJ2dTI5amxuNnF3UElTWGZzUUVBZGtqUUFtREdOcWNYMW9jd0pkYUhBTUFXTk4wUTMyQTlVVXNLdDlSZ01tQjlDQ2tKV2dBQTdKQ2dCUUNtV2pBVjFvY0F3SFo0eTUvU3ZldmIycVFBY3VCelNrcENsd0FBT3lSb0FZQ2dCVk1SMTRmWUlRc0F6OUIwdzNtOHBxb2hCWXNUK2M0MW1CejBiZjB4aEhDcldhUXlUcmtDQUdBSEJDMEFaczc2RUNiRytoQUFlS0ttR3c2dFUyQUdMc2VIMTVBTDYwTkl5Zm9RQUlBZEViUUFJSmhxd1lSWUh3SUFUM2NaUWpoUVB3cDIyN2UxTUJHNXNUNkVsRXkwQUFEWUVVRUxBSUtnQlJOeVVLMVh3aFlBOEVoTk44UTNWdCtvRzRXek1vUWNtV2hCU2laYUFBRHNpS0FGQU5hSE1EV0NGZ0R3ZUpkcVJ1R3UrN1lXRUNjN2ZWdWJhRUZLUitOcU1RQUF0a3pRQW9CUDNxb0VFeUZvQVFDUDBIUkR2SGFlcUJtRk85TmdNbmF0ZVNSa2ZRZ0F3QTRJV2dEd2lhQUZVMkY5Q0FBODBQaVdxbWtXbE83WHZxMXZkSm1NV1I5Q1N0YUhBQURzZ0tBRkFML2JuRjdFSDM1dVZZT0pFTFFBZ0ljNWoyUEIxWXFDeFJXSFN3MG1jNElXcEdTaUJRREFEZ2hhQVBBNU80K1ppamZWZW1XUExBQjhSOU1OeDJQUUFrcTI3TnY2b3c2VHVTc05KQ0VUTFFBQWRrRFFBb0RQV1IvQ2xKaHFBUURmRjkveVAxQWpDbmJidDdYVk9HUnZYSDF6cDVNa2N0QjBnN0FGQU1DV0NWb0E4Qi9XaHpBeGdoWUE4QTFOTjhReDRHL1VoOEtkYVRBRk1kV0NsQVF0QUFDMlROQUNnQzlaSDhKVXZMSStCQUMrYWFrMEZPNjN2cTA5bUtZa0c5MGtJVUVMQUlBdEU3U29PNlUwQUFBZ0FFbEVRVlFBNEV2V2h6QWxwbG9Bd0JlYWJvaHYrWitvQzRVNzEyQUtJemhFU2d2VkJ3RFlMa0VMQVA3aS8yZnZibTRpeWJJMkFJY0hpUWZnQVhnQXE5aFdTU0dGMklFSG5SNlFlSkR0UWJGRElZVlV0WTFWNFVIaUFYaEFldkFwWnVLYjZaNnVIMzR5ODU2NDkzbE1PSWVpZ0h6dmVkV0hFSXlnQlFEOFJkTU40N1dudFptUXVkdStyWjhzbWN5NGFFRktwNllQQUxCYmdoWUEvSWlyRmtReDFvZWMyQVlBL01mNHluOWhIR1JzSzB4RWp2cTJmcW1xNnRGeVNhWHBCbGN0QUFCMlNOQUNnQjhSdENBU1Z5MEE0Tjhma0l6aHd4dXpJSFBMNlFOcHlKR3JGcVFrYUFFQXNFT0NGZ0Q4dytieTVzbExHd0s1dGd3QStCZXYvTW5kWTkvV1F0L2s3THZ0a3RDWjRRTUE3STZnQlFBLzR3K2NSSEdxUGdTQTBrM252aitWUGdleXQ3UmlNdWVpQlNrSldnQUE3SkNnQlFBLzg5VmtDRVI5Q0FDbGM4MkMzTjMxYmUyMVAxbnIyM29NV214dG1VU09weG95QUFCMlFOQUNnQjlTSDBJdzZrTUFLRmJURGVNci8xTmZBV1JzL09CNVpjRVV3bFVMVW5MVkFnQmdSd1F0QVBnVjlTRkVvVDRFZ0NJMTNYRGtBMmdLc083YitzbWlLWVRMTGFRa2FBRUFzQ09DRmdEOGl2b1FJbkhWQW9BU2pTR0xoYzJUc1dmVk9CVEdSUXRTdWpCOUFJRGRFTFFBNEtmVWh4Q01vQVVBUlpsNjFQK3dkVEszNnR2NnhaSXBpSXNXcEhSdStnQUF1eUZvQWNEdnFBOGhpdU96KzF0blRnRW9pWi9EeU4xRDM5YSt6aW5LRkN4NnRuVlNhYnJCNzlVQUFEc2dhQUhBNy9qREo1RzRhZ0ZBRVpwdXVQRHFsQUlzTFpsQ3VXcEJTb0lXQUFBN0lHZ0J3Qzl0TG0vRzF6YmZUSWtnUGxzRUFJVVFkaVYzZDMxYmIyeVpRdm5hSjZVTDB3Y0ErRGhCQ3dCZTQ2c3BFWVQ2RUFDeTEzVERhdncvejZiSjJOWTFDd29uYUVGS2ZxY0dBTmdCUVFzQVhrUFFna2pVaHdDUXJhWWJqbndBVFFIV2ZWdS9XRFNsNnR0YWRRZ3BuVTQvYndBQThBR0NGZ0Q4bHZvUWdsRWZBa0RPMWxWVkxXeVlqRDMzYmIyeVlLZ2VqSUNFWExVQUFQZ2dRUXNBWHN0VkM2SVk2ME9FTFFESVR0TU40NGNlVnpaTDVsd25nMzlUSDBKS0Y2WVBBUEF4Z2hZQXZKYWdCWkVJV2dDUW83V3RrcmtIbFFud0gvNHRrSktMRmdBQUh5Um9BY0NycUE4aEdFRUxBTExTZE1QNHl2L2NWc21jYXhid1h5NWFrSktMRmdBQUh5Um9BY0JidUdwQkZBdjFJUURrb3VtR282cXFWaFpLNXY3czIvckprdUhmcG44UFcrTWdrVVhURFNlR0R3RHdmb0lXQUx6RlYzOElJaEJCQ3dCeXNheXE2dGcyeWRoV21BaCtTSDBJS2FrUEFRRDRBRUVMQUY1dHFnOXgxWUlvQkMwQW1MM3BOZW5TSnNuY3FtL3JGMHVHZjFBZlFrcnFRd0FBUGtEUUFvQzNFclFnQ3ZVaEFPUmdmT1cvc0VreTl0aTM5ZHFDNFlkY3RDQWxGeTBBQUQ1QTBBS0FOOWxjM3FnUElaSnIyd0JncnBwdUdGK1NYbGtnbVhPeEJYNmliMnRCQzFJNk4zMEFnUGNUdEFEZ1BWeTFJSXBQWi9lM1I3WUJ3RXg1NVUvdXZ2a2dHWDdyMFloSXBla0dWeTBBQU41SjBBS0E5eEMwSUJMMUlRRE1UdE1ONDFXbVU1c2pjNjVad085dHpJaUVMZ3dmQU9COUJDMEFlRFAxSVFRamFBSEFyRFRkY09TYUJRVzQ3ZHY2eWFMaHQxeDlJU1VYTFFBQTNrblFBb0QzY3RXQ0tOU0hBREEzNHl2L2hhMlJzYTB3RWJ5YWl4YWs1S0lGQU1BN0NWb0E4RjZDRmtUaXFnVUFzOUIwdzBsVlZUZTJSZWFXZlZ1L1dETDhYdC9XR3hjalNlaDR1clFGQU1BYkNWb0E4QzVUZmNpejZSSEV0VVVBTUJOZStaTzdoNzZ0djlneXZJbXJGcVNrUGdRQTRCMEVMUUQ0Q0ZjdGlPTDg3UDcyeERZQWlLenBodkU4OXlkTEluTXJDNFkzKzI1a0pLUStCQURnSFFRdEFQZ0lMOVdJUkgwSUFOSDUyWW5jM2ZWdDdRTmplRHYvYmtoSjBBSUE0QjBFTFFCNHQ4M2x6VVo5Q0lHb0R3RWdyS1libG1NUHVnMlJzYTFyRnZCdXFrTklTWFVJQU1BN0NGb0E4RkhxUTRqaVZIMElBQkUxM1hEa0EyZ0tzTzdiK3NtaTRlMzZ0bjd4aUlHRUZrMDMrRjBhQU9DTkJDMEErQ2duc0lsRWZRZ0FFWTBoaTRYTmtMSHhBK0sxQmNPSHFBOGhKZlVoQUFCdkpHZ0J3SWVvRHlFWTlTRUFoTkowdzNpTyt3OWJJWFBMNlVVKzhIN3FRMGhKZlFnQXdCc0pXZ0N3QytwRGlFSjlDQURSZU9WUDdoNzZ0dmI3QUh5Y2l4YWs1S0lGQU1BYkNWb0FzQXZxUTRoa2FSc0FSTkIwdzFocGRXNFpaTTdQWHJBRGZWdTdhRUZLcDZZUEFQQTJnaFlBZkpqNkVJTDViQ0VBQk9HYUJibTc4K0V3N05TRGNaSkswdzJ1V2dBQXZJR2dCUUM3NG9NRW9qZyt1Ny9WTHd0QVVrMDNyTWIvazJ5QmpHMWRzNENkRTF3aUpiOUhBd0M4Z2FBRkFMdWlsNWxJcm0wRGdGU2Fiamp5QVRRRldQVnQvV0xSc0ZQZmpaT0VYTFFBQUhnRFFRc0FkbUp6ZWZOVVZkV2phUktFK2hBQVVob3ZmUzFzZ0l3OTkyM3RvaDNzbm9zV3BPU2lCUURBR3doYUFMQkxYMHlUSU5TSEFKREUxRzkrWmZwa3p2VXcySU8rcmNjSERNOW1TeUxIMDFVdUFBQmVRZEFDZ0YxU0gwSWtQZ0FBSUlXVnFaTzVoNzZ0MVJ2QS9yaHFRVXJxUXdBQVhrblFBb0NkVVI5Q01JSVdBQnhVMHczai96M25wazdtL0l3Rit5Vm9RVW91UXdJQXZKS2dCUUM3cGo2RUtCWm45N2VmYlFPQVE1aE9iYnRtUWU3K25Lb05nUDF4TVlhVVhMUUFBSGdsUVFzQWRrMTlDSkVJV2dCd0tNdXgyOXkweWRoV21BajJUelVQaWJsb0FRRHdTb0lXQU95VStoQ0NFYlFBWU8rYWJqaVpnaGFRczJYZjFpODJEQWZoZDJwU1dUVGRJR3dCQVBBS2doWUE3TVBhVkFsQ2ZRZ0FoekQrN0xNd2FUTDIyTGUxaWtBNEhGY3RTRW5RQWdEZ0ZRUXRBTmdIOVNGRUltZ0J3TjQwM1RCMm1YOHlZVExuWWdzYzFzYThTZWpDOEFFQWZrL1FBb0NkMjF6ZWpDZUZ2NWtzUVh3K3U3ODlzZ3dBOXNRbEwzTDNyVzlycit2aHNQeWJJeVVYTFFBQVhrSFFBb0I5Y2RXQ0tCYXVXZ0N3RDAwM1hGZFZkV3E0Wk00MUN6aXd2cTJmcXFyYW1qdUorTmtHQU9BVkJDMEEyQmRCQ3lJUnRBQmdwNXB1T0hMTmdnTGNUaC80QW9lblBvUmtwbW8wQUFCK1FkQUNnTDFRSDBJd245U0hBTEJqeStscUV1VHFXWmdJa2xJZlFrcUNGZ0FBdnlGb0FjQSt1V3BCSks1YUFMQVRUVGVjVkZWMVk1cGtidFczOVlzbFF6S0NGcVIwWnZvQUFMOG1hQUhBUGdsYUVJbWdCUUM3OHNVa3lkeEQzOWEremlFdDFTR2tKR2dCQVBBYmdoWUE3STM2RUlKUkh3TEFoMDJkNWVjbVNlWldGZ3hwVFJkbEhxMkJSSTZuQzE0QUFQeUVvQVVBKythcUJaRzRhZ0hBUjNubFQrN3UrclpXV1FBeHVHcEJTcTVhQUFEOGdxQUZBSHUxdWJ3WlA0elltakpCTEMwQ2dQZHF1bUg4ZitUWUFNblkxalVMQ0VYUWdwUUVMUUFBZmtIUUFvQkRjTldDS0U3UDdtK2RQd1hnelpwdU9QSUJOQVZZOTIzOVpORVFodXN5cEhSaCtnQUFQeWRvQWNBaENGb1FpZm9RQU41alhWWFZ3dVRJMkhQZjFzSkVFRWpmMWk1YWtOSzU2UU1BL0p5Z0JRQjd0N204K2FvK2hFQ3VMUU9BdDJpNllUeWRmV1ZvWkU3RkdzVDBZQytrTXYwTUJBREFEd2hhQUhBb3Jsb1FoZm9RQU41cWJXSms3cUZ2YXordlEwenFRMGhKZlFnQXdFOElXZ0J3S1A1d1N5VHFRd0I0bGFZYlBqdWRUUUZjczRDNDFJZVFrb3NXQUFBL0lXZ0J3RUdvRHlFWTlTRUF2SlpyRnVUdXJtOXJIK1JDWFA1OWtwS2dCUURBVHdoYUFIQklybG9ReFZnZjRnOUdBUHhTMHcycnFxcU9UWW1NYlYyemdOajZ0bjZxcXVyWm1ramt0T21HSThNSEFQZ25RUXNBRGtuUWdraGN0UURncDVwdU9QRUJOQVZZOVczOVl0RVFucXNXcE9TUkFnREFEd2hhQUhBd1UzMklsemhFOGRrbUFQaUY4WnJGd29ESTJIUGYxcXB4WUI2KzJ4TUpYUmcrQU1BL0NWb0FjR2l1V2hERnNmb1FBSDZrNllieEE0VXJ3eUZ6cm52QmZMaG9RVXArYndZQStBRkJDd0FPN1l1SkU0Z1BHQUQ0a1pXcGtMbUh2cTI5a0llWjhPK1Z4RnkwQUFENEFVRUxBQTVxYzNtelVSOUNJT3BEQVBpYnBodkdFTjY1cVpBNVlWT1luMGM3STVGRjB3MG5oZzhBOEhlQ0ZnQ2tvRDZFS01iNkVLOXpBUGlYcGh1T1hMT2dBTGQ5V3o5Wk5NeU9xeGFrcEQ0RUFPQi9DRm9Ba0lMNkVDTHhvaE9BLzdjY1EzaW1RY2EyVlZXdExSaG1hV050Sk9TQkFnREEveEMwQU9EZzFJY1FqUG9RQUtycEpQYU5TWkM1WmQvV0w1WU1zK1NpQlNtNWFBRUE4RDhFTFFCSVJYMElVU3pPN20rRkxRRHd5cC9jUGZadDdiSWN6TlJVK2JPMVB4STVOM2dBZ0w4VHRBQWdGWC9rSlJKQkM0Q0NOZDB3bnNQKzVHdUF6QzB0R0diUFZRdVNtWDVlQWdCZ0ltZ0JRQkpUZmNpajZST0VvQVZBMlZ5eklIZmYrcmIyQVMzTTM4WU9TVWg5Q0FEQVh3aGFBSkNTcXhaRW9UNEVvRkJOTjF4WFZYVnEvMlJzNjVvRlpFTmdpcFFFTFFBQS9rTFFBb0NVdnBvK2dRaGFBQlNtNllZajF5d293THB2NnllTGhpeTRhRUZLcWtNQUFQNUMwQUtBWkRhWE4wL3FRd2prNnV6KzlzaENBSXF5R3E4YVdUa1pleFltZ256MGJmM2lkMmdTT3A1Q3FnQUF4YXNFTFFBSVFIMElrYmhxQVZDSXBodE9xcXI2dzc3SjNHcjZZQmJJaDZzV3BLUStCQUJnSW1nQlFHcnFRNGhFMEFLZ0hNS2U1TzZoYjJ0ZjU1Q2Y3M1pLUXVwREFBQW1naFlBSktVK2hHQStxUThCeUYvVERlT0hCT2RXVGVaV0ZneFpjdEdDbEFRdEFBQW1naFlBUk9DbEhaRzRhZ0dRUHo5N2tMdTd2cTI5ZW9jTTlXMHRhRUZLcWtNQUFDYUNGZ0JFb0Q2RVNBUXRBRExXZE1QNHl2L1lqc25ZdHFxcXBRVkQxaDZzbDBRV1RUZWNHRDRBZ0tBRkFBRk05U0grVUVRVTZrTUFNdFYwdzVFUG9DbkF1bS9yRjR1R3JMbFlRMHJxUXdDQTRsV0NGZ0FFNG9RM2tWemJCa0NXMXVOTFRLc2xZODk5VzY4c0dMS25Qb1NVMUljQUFNV3JCQzBBQ0VSOUNKRUlXZ0JrcHVtRzhVT0JLM3NsY3k2MlFCbGN0Q0FsRnkwQWdPSlZnaFlBUkxHNXZCblBHMyt6RUlJNFBidS8xVHNMa0plMWZaSzVoNzZ0aFplaEFGTTkwTE5kazhpcHdRTUFDRm9BRUlzL0RCUEpaOXNBeUVQVERlUDM5SFBySkhPdVdVQloxSWVRVE5NTnJsb0FBTVVUdEFBZ0VrRUxJbEVmQXBDQnBodU9YTE9nQUgvMmJlMURWeWlMK2hCU09qTjlBS0IwZ2hZQWhLRStoR0RVaHdEa1lYemxmMnlYWkd4YlZkWEtncUU0d2xXazVLSUZBRkE4UVFzQW9uSFZna2pVaHdETVdOTU5KK29VS01DcWIrc1hpNGF5OUczdG9nVXB1V2dCQUJSUDBBS0FhQVF0aU1TSGN3RHpOcjd5WDlnaEdYdnUyMW8xRHBUcndlNUo1SGlxWndNQUtKYWdCUUNocUE4aG1PT3orMXN2ZFFCbXFPbUc4YVQxbGQyUnVXc0xocUtwRHlFbDlTRUFRTkVFTFFDSTZJdXRFSWdQTUFEbWFXVnZaTzZiNmdBb25xQUZLWG1VQUFBVVRkQUNnSEEybHpkamZjaldaZ2ppczBVQXpFdlREV05JN3R6YXlKeUtNMERZaXBSY3RBQUFpaVpvQVVCVVgyMkdJTlNIQU16STFCZSt0ak15ZDl1MzlaTWxROW1tN3dNZUtaQ0tVQ3NBVURSQkN3Q2lFclFnRXZVaEFQTXh2dkpmMkJjWjJ3b1RBWC9ocWdYSk5OM2dVUUlBVUN4QkN3QkNVaDlDTU9wREFHYWc2WWFUcXFwdTdJck1MZnUyZnJGa1lMSXhDQklTdEFBQWlpVm9BVUJrcmxvUXhWZ2ZJbXdCRUo5WC91VHVzVy9yTDdZTS9JV0xGcVIwWWZvQVFLa0VMUUNJVE5DQ1NBUXRBQUpydW1IOFEvOG5PeUp6U3dzRy9xcHZhMEVMVW5MUkFnQW9scUFGQUdHcER5RVlRUXVBMkZ5eklIZDNQbEFGZnVMUllFamsxT0FCZ0ZJSldnQVFuYXNXUkxGUUh3SVFVOU1OUzMvb0ozTmorSGhseWNCUGJBeUdWS2FyWWdBQXhSRzBBQ0E2SGRSRUltZ0JFRXpURFVjK2dLWUE2NzZ0bnl3YStBblhia2hKMEFJQUtKS2dCUUNoYlM1dnhqOFlQZHNTUVFoYUFNUXpoaXdXOWtMR25sWGpBTC9ob2dVcG5aaytBRkFpUVFzQTVrQjlDRkdvRHdFSXBPbUdrNnFxL3JBVE1yZnEyL3JGa29HZjZkdDZNMVVNUVFxQ0ZnQkFrUVF0QUpnRDlTRkVjbTBiQUdINEdZSGNQZlJ0N2VzY2VBMVhMVWpsZUFxL0FnQVVSZEFDZ1BBMmx6Y2I5U0VFOHVucy92YklRZ0RTYXJwaHZEQjBiZzFrYm1uQndDdDlOeWdTY3RVQ0FDaU9vQVVBYzZFK2hFalVod0NrdDdZRE1uYzMxUUVBdklidkY2UjBZZm9BUUdrRUxRQ1lDeWVUaVVUUUFpQ2hwaHRXNDVscU95QmpXOWNzZ0RkeTBZS1VYTFFBQUlvamFBSEFMS2dQSVJqMUlRQ0pOTjF3NUFOb0NyRHUyL3JGb29IWG1yNW4rSjJaVk5TNUFRREZFYlFBWUU3VWh4Q0pxeFlBYVl5VklRdXpKMlBQZlZ1dkxCaDRCMWN0U0ticEJsY3RBSUNpQ0ZvQU1DZnFRNGhFMEFMZ3dLWS80RitaTzVsenNRVjRyNDNKa2RDRjRRTUFKUkcwQUdBMnB2cVFSeHNqaUxFKzVNUXlBQTVxYmR4azdxRnZhMWZjZ1BjU3RDQWxGeTBBZ0tJSVdnQXdONjVhRUltckZnQUgwblREdGY1dkNuQnR5Y0I3OVcydE9vU1VCQzBBZ0tJSVdnQXdOMTc0RVlrUFF3QU9vT21HbzZxcVZtWk41djdzMi9ySmtvRVBlakJBRWptZGZtWURBQ2lDb0FVQXM3SzV2SGxTSDBJZ3ArcERBQTVpV1ZYVnNWR1RzYTB3RWJBajZrTkl5VlVMQUtBWWdoWUF6Skg2RUNKUkh3S3dSMDAzbkV4QkM4alpxbS9yRnhzR2RrQjlDQ2xkbUQ0QVVBcEJDd0RtU0gwSWthZ1BBZGl2OFpYL3dvekoyR1BmMW1zTEJuYkVSUXRTY3RFQ0FDaUdvQVVBczZNK2hHRFVod0RzU2RNTjQ2dklLL01sY3k2MkFEdlR0L1g0Ky9LemlaS0lpeFlBUURFRUxRQ1lLL1VoUk9JREVvRDk4TXFmM0gzcjI5cVpmMkRYWExVZ2xjVlUrd1lBa0QxQkN3RG1TbjBJa1h5MkRZRGRhcnBockdZNk5WWXlKNndKN0lPZ0JTbTVhZ0VBRkVIUUFvQlptdXBEdnRrZVFSeWYzZC9xb2dYWWthWWJqbHl6b0FDMzA0bC9nRjF6S1llVS9HNE1BQlJCMEFLQU9YUFZna2l1YlFOZ1o4WlgvZ3ZqSkdOYllTSmdYMVFTa1ppZ0JRQlFCRUVMQU9aTTBJSkkxSWNBN01EVTYzMWpsbVJ1MmJmMWl5VURlL1JvdUNSeWJ2QUFRQWtFTFFDWXJjM2x6WXY2RUFKUkh3S3dHMS9Na2N3OTlHM3Q2eHpZTjFjdFNLYnBoZ3ZUQndCeUoyZ0J3Tnk1YWtFazZrTUFQbUQ2bzd4WGtPUnVaY1BBQVd3TW1ZUThRZ0FBc2lkb0FjRGNDVm9RaWFBRndNZDQ1VS91N3ZxMjlzb2NPQVJCQzFJU3RBQUFzaWRvQWNDc3FROGhtTVhaL2UxblN3RjR1NllibG1NTms5R1JzYTFyRnNDaDlHMjltYjd2UUFxcVF3Q0E3QWxhQUpBRFZ5MklSTkFDNEkyYWJqanlBVFFGV1BkdC9XVFJ3QUc1YWtFcXg5UFBkd0FBMlJLMEFDQUhnaFpFSW1nQjhIWmp5R0poYm1Uc2VReGFXREJ3WUtxS1NFbDlDQUNRTlVFTEFHWnZxZys1czBtQ1VCOEM4QVpOTjR4L2hQL0R6TWpjc20vckYwc0dEa3pRZ3BUVWh3QUFXUk8wQUNBWHJsb1FpYUFGd090NTVVL3VIdnEyOXJNcWtJTHFFRklTdEFBQXNpWm9BVUFXTnBjMzR4K3Z0N1pKRUlJV0FLL1FkTVA0L2ZMY3JNamMwb0tCRktaTE9zK0dUeUtxUXdDQXJBbGFBSkFUTHdXSllxd1B1YllOZ045eXpZTGMzZlZ0N1VVNWtKTDZFRkpaVEJWeEFBQlpFclFBSUNlQ0ZrVGlxZ1hBTHpUZHNLcXE2dGlNeU5qV05Rc2dBR0V2VWhLMEFBQ3lKV2dCUURiVWh4RE1wN1A3MnlOTEFmaW5waHVPZkFCTkFkYlQyWDZBbEZ5MElDVkJDd0FnVzRJV0FPVEdWUXNpY2RVQzRNZkd5cENGMlpDeDU3NnRWeFlNcEthK2lNUXVMQUFBeUpXZ0JRQzVFYlFnRWtFTGdQL1JkTVA0Qi9jcmN5RnoxeFlNQlBKZ0dTUnlhdkFBUUs0RUxRRElpdm9RZ2xFZkF2QlBYdm1UdTRlK3JaM3FCeUp4MVlKa3BwQXRBRUIyQkMwQXlKR3JGa1RpcWdYQXBPbUc4WlgvdVhtUU9kY3NnR2lFdjBqcHpQUUJnQndKV2dDUW83V3RFc2pTTWdEK0ZiSTRjczJDQXZ6WnQvV1RSUVBCdUdoQlNpNWFBQUJaRXJRQUlEdWJ5NXZ4ajBqUE5rc1FwMmYzdHllV0FmQ3Y0Tm14TVpDeHJUQVJFTkVVQVBNN01xbTRhQUVBWkVuUUFvQmNxUThoRXZVaFFOR2Fiamh4NFljQ3JQcTJmckZvSUNoWExVamxlUHBaRUFBZ0s0SVdBT1RxaTgwU2lLNTJvSFJqcmRlaTlDR1F0Y2UrcmRYWEFaRjl0eDBTY3RVQ0FNaU9vQVVBV1ZJZlFqRHFRNEJpTmQwdzluSi84aFZBNWx4c0FhSnowWUtVQkMwQWdPd0lXZ0NRTS9VaFJLSStCQ2lWVi83azdsdmYxbDZLQTZINVBrVmlGeFlBQU9SRzBBS0FuS2tQSVJMMUlVQnhtbTRZdi9lZDJqeVpjODBDbUl0SG15S1JjNE1IQUhJamFBRkF0dFNIRU14WUgrSmNLbENNcGh1T1hMT2dBTGQ5V3o5Wk5EQVRybHFRVE5NTmZoOEdBTElpYUFGQTd0U0hFSW1yRmtCSlZsVlZMV3ljakQwTEV3RXpzN0V3RWhLMEFBQ3lJbWdCUU83ODhadElQdHNHVUlLbUcwNnFxdnJEc3NuY3FtL3JGMHNHWnNSRkMxSzZNSDBBSUNlQ0ZnQmtiWE41ODZTSGxrQ08xWWNBaGZoaTBXVHVvVzlyWCtmQXJFeFZSMXRiSXhHL0N3TUFXUkcwQUtBRS9naE9KT3BEZ0t3MTNUQytWankzWlRLM3NtQmdwdFNIa01xcHlRTUFPUkcwQUtBRVgyMlpRTlNIQUxrVGNDUjNkMzFiTzc4UHpKWHZYeVF6QlhJQkFMSWdhQUZBOXRTSEVJejZFQ0JiVFRjc3grOXpOa3pHdHE1WkFETW5hRUZLZ2hZQVFEWUVMUUFvaGRlMVJMSzBEU0EzVFRjYytRQ2FBcXo3dG42eWFHREdWSWVRa2tjSEFFQTJCQzBBS0lYNkVDSlJId0xrYUYxVjFjSm15ZGh6MzliQ1JNQ3M5VzM5NHVJakNibG9BUUJrUTlBQ2dDS29EeUdZeGRuOXJiQUZrSTJtRzhiWGlWYzJTdVpjcEFKeTRhb0ZxU3lhYmpneGZRQWdCNElXQUpSRWZRaVJDRm9BT1ZuYkpwbDc2TnZhaFRRZ0Y0SVdwS1ErQkFESWdxQUZBQ1h4eDNFaUViUUFzdEIwdy9qOTdOdzJ5WnhyRmtCT3Z0c21DYWtQQVFDeUlHZ0JRREdtK3BCdk5rNFE2a09BWExobVFlN3UrcmIyK2h2SWh1OXBKT2FpQlFDUUJVRUxBRXJqcWdXUkNGb0FzOVowdzZxcXFtTmJKR05iMXl5QVREMVlMSW00aEFZQVpFSFFBb0RTQ0ZvUXlkWFovZTJSalFCejFIVERpUStnS2NDcWIrc1hpd1l5cEQ2RVpKcHVjTlVDQUpnOVFRc0Fpcks1dkhsUkgwSXdybG9BY3pWZXMxallIaGw3N3R0YU5RNlFLL1VocEhSaCtnREEzQWxhQUZBaVZ5MklSTkFDbUoybUc4WS9qbC9aSEptN3RtQWdZeTVha0pLTEZnREE3QWxhQUZBaVFRc2krYVErQkppaGxhV1J1WWUrclgwSUNXUnJxa1Y2dG1FU0ViUUFBR1pQMEFLQTRxZ1BJU0JYTFlEWmFMcGhmT1YvYm1Oa3pqVUxvQVRxUTBqbHRPa0dEdzRBZ0ZrVHRBQ2dWSzVhRUltZ0JUQUwweC9FWGJNZ2QzLzJiZjFreTBBQlhPNGhKVmN0QUlCWkU3UUFvRlNDRmtTaVBnU1lpMlZWVmNlMlJjYTJ3a1JBUVZ5MElLVUwwd2NBNWt6UUFvQWlUZlVoZDdaUElFNlVBNkUxM1hCU1ZkV05MWkc1WmQvV0w1WU1sS0J2YXhjdFNFblFBZ0NZTlVFTEFFcm1xZ1dSQ0ZvQTBhMXRpTXc5OW0zOXhaS0J3anhZT0ltb0RnRUFaazNRQW9CaWJTNXZ2azdub1NHQzA3UDcyeE9iQUNKcXVtRjhjZmpKY3NqYzBvS0JBcWtQSVpYRmRERU5BR0NXQkMwQUtKMnJGa1R5MlRhQW9GeXpJSGZmbk5BSENpVm9RVXJxUXdDQTJSSzBBS0IwZ2haRW9qNEVDS2ZwaHZHVi82bk5rTEd0YXhaQXdZVE1TRWw5Q0FBd1c0SVdBQlJOZlFqQnFBOEJRbW02NGFpcXFwV3RrTGwxMzlaUGxneVVhUHIrNTNkaVVoRzBBQUJtUzlBQ0FGeTFJQmIxSVVBa1k4aGlZU05rN0ZrMURvQ3JGaVJ6YnZRQXdGd0pXZ0NBb0FXeE9GME9oTkIwdzNoaDV3L2JJSE9ydnExZkxCa28zS2IwQVpCTzB3MFh4ZzhBekpHZ0JRREZVeDlDTU1kbjk3Zk9wd0lSZkxFRk12ZlF0N1d2Y3dBWExVakw3NzhBd0N3SldnREF2L2tqTzVGYzJ3YVEwdlN5MENsbmNyZXlZWUIvY2RHQ2xBUXRBSUJaRXJRQWdIOFR0Q0NTejdZQkpPYi9SWEozMTdlMUY5d0FWVlZORlVxUFprRWlxa01BZ0ZrU3RBQ0FmOWVIakM5NG5zMkNJTlNIQU1rMDNUQys4aisyQVRJMlZzWXRMUmpnYjF5MUlKWGpwaHVPVEI4QW1CdEJDd0Q0cjY5bVFTRHFRNENEbS83STdRTm9jcmVlWG04RDhGK3UvSkNTcXhZQXdPd0lXZ0RBZnptVFRpVHFRNEFVMWxWVkxVeWVqRDMzYmIyeVlJQi9jTkdDbEZ4MEJBQm1SOUFDQUNicVF3aG1yQS94cWdjNG1LWWJ4ajl3WDVrNG1YT3hCZUFIK3JZV3RDQWx2L3NDQUxNamFBRUFmNmMraEVqVWh3Q0h0RFp0TXZmUXQ3V2Y5UUIrN3NGc1NNUkZDd0JnZGdRdEFPRHYxSWNRaWZvUTRDQ2FiaGkvMzV5Yk5wbHp6UUxnMTc2YkQ0a3NwdXRxQUFDeklXZ0JBSCtoUG9SZ0ZtZjN0OElXd0Y0MTNYRGttZ1VGK05OWmZJRGY4bjJTbEFRdEFJQlpFYlFBZ0g5eTFZSklCQzJBZlJ0ZitSK2JNaG5iVmxXMXNtQ0EzM0xSZ3BRRUxRQ0FXUkcwQUlCL0VyUWdFa0VMWUcrYWJqaFJwMEFCVm4xYnYxZzB3SzlOM3l0ZGVDU1ZDNU1IQU9aRTBBSUEvc2ZtOHVhcHFxcEhjeUVJOVNIQVBvMnYvQmNtVE1hZSs3WldqUVB3ZXE1YWtNcXB5UU1BY3lKb0FRQS81cW9Ga1Z6YkJyQnJUVGVNcndhdkRKYk0rVDhVNEcwMjVrVXEwOCtuQUFDeklHZ0JBRC8yMVZ3STVOUFovZTJSaFFBN3RqSlFNdmZRdDdXWDJRQnZJMmhCU21lbUR3RE1oYUFGQVB5QStoQUNVaDhDN0V6VERlTXIvM01USlhPdVdRQzhrWUFhaWJsb0FRRE1ocUFGQVB5YytoQWlFYlFBZHFMcGh2RkN6dG8weWR4dDM5WlBsZ3p3TGcvR1JpSXVXZ0FBc3lGb0FRQS9wejZFU05TSEFMdXlyS3BxWVpwa2JDdE1CUEFoNmtOSTVianBoaFBUQndEbVFOQUNBSDVDZlFnQnVXb0JmTWowaCtzYlV5Unp5NzZ0WHl3WjROM1VoNUNTcXhZQXdDd0lXZ0RBcjZrUElSSkJDK0NqdlBJbmQ0OTlXL3Y1RGVCalhMUWdKVUVMQUdBV0JDMEE0TmY4b1o1SXh2b1FaMVNCZDJtNjRXTDhQbUo2Wkc1cHdRQWYwN2YxMDFUREJDbGNtRG9BTUFlQ0ZnRHdDNXZMbS9IczlEY3pJaEJYTFlEM0VoNGtkM2Q5V3p0M0Q3QWJ2cCtTeXJuSkF3QnpJR2dCQUwvMzFZd0k1Tm95Z0xkcXVtRjg1WDlzY0dSc2ZIbTlzbUNBblZFZlFqSk5ONmdQQVFEQ0U3UUFnTjhUdENDU1UvVWh3RnMwM1hEa0EyZ0tzSjVPM1FPd0d5NWFrSktnQlFBUW5xQUZBUHlHK2hBQ1VoOEN2TVVZc2xpWUdCbDdIb01XRmd5d082cVlTT3pDQWdDQTZBUXRBT0IxWExVZ0V2VWh3S3MwM1RCZXdQbkR0TWpjcW0vckYwc0cyTGxISXlVUkZ5MEFnUEFFTFFEZ2RRUXRpRVI5Q1BCYVgweUt6RDMwYmUzckhHQS9OdVpLSXFkVC9SMEFRRmlDRmdEd0N1cERDTWhWQytDWG1tNFlhNGJPVFluTUxTMFlZRy9VaDVDU3F4WUFRR2lDRmdEd2VxNWFFSW1nQmZBN2F4TWljM2Q5VzN0dERiQS92c2VTMG9YcEF3Q1JDVm9Bd091TlFZdXRlUkhFOGRuOXJSYyt3QTgxM2JBYXYwK1lEaG5idW1ZQnNGOVRtTTN2d0tUaTkxMEFJRFJCQ3dCNHBhayt4RlVMSW5IVkF2aUhxYy9hQjlEa2J0MjM5WXN0QSt5ZHF4YWs0cUlGQUJDYW9BVUF2STJnQlpGOHRnM2dCOGJLa0lYQmtMSG52cTFYRmd4d0VOK05tVVFXVFRlY0dENEFFSldnQlFDOHdlYnlSbjBJa2FnUEFmNm02WWJ4ZThLVnFaQTVGMXNBRHNkRkMxTHkreTRBRUphZ0JRQzhuYXNXUktJK0JQaXJ0V21RdVdJd3JuNEFBQ0FBU1VSQlZJZStyZjBzQm5BNExscVFrdm9RQUNBc1FRc0FlRHQvM0NjU1FRdmdYNXB1R0w4Zm5Kc0dtZlAvSHNBQjlXMzlNbFkybVRtSnVHZ0JBSVFsYUFFQWI2UStoR0FXWi9lM255MEZ5dFowdzFGVlZhdlM1MEQyL3V6YitzbWFBUTdPVlF0U0VTSUdBTUlTdEFDQTkzSFZna2dFTFlCbFZWWEh4VStCbkcyRmlRQ1MyUmc5cVRUZDRLb0ZBQkNTb0FVQXZJK2dCWkVJV2tEQm1tNDRtWUlXa0xQVmRMNGVnTU56MFlLVUxrd2ZBSWhJMEFJQTNrRjlDTUdvRDRHeWphLzhGNlVQZ2F3OTkyMjl0bUtBTlBxMmR0R0NsRnkwQUFCQ0VyUUFnUGY3WW5ZRUltZ0JCV3E2WVh6aGQyWDNaTzdhZ2dHU2U3QUNFbkhSQWdBSVNkQUNBTjVQMElKSUJDMmdURjc1azd0dmZWczdXUStRbnFzV3BITGNkTU9SNlFNQTBRaGFBTUE3YlM1dnhqODBQWnNmUVl6MUlWNzhRa0dhYmhqL3paL2FPWmxiV2pCQUNFSnZwS1ErQkFBSVI5QUNBRDdtcS9rUmlLc1dVSWpwVlo5ckZ1VHV0bS9ySjFzR0NNRkZDMUpTSHdJQWhDTm9BUUFmb3o2RVNENmQzZDg2cVFwbEdGLzVMK3lhakcyRmlRRGltSUp2TGpxU2lxQUZBQkNPb0FVQWZJRDZFQUp5MVFJeTEzVERTVlZWTi9aTTVwWjlXNzlZTWtBb3JscVFpdW9RQUNBY1FRc0ErRGoxSVVRaWFBSDVjMDJKM0QzMmJlM3JIQ0FlUVF0U1dVeGhZd0NBTUFRdEFPRGpmQkJBSk9wRElHTk5ONHhuazgvdG1Nd3RMUmdncE8vV1FrTHFRd0NBVUFRdEFPQ0QxSWNRa0tzV2tDL2hQbkozMTdlMUQvSUFBdkw5bWNUVWh3QUFvUWhhQU1CdXJNMlJRSzR0QS9MVGRNUDR5di9ZYXNuWXRxcXFsUVVEaFBab1BTUWlhQUVBaENKb0FRQzc4ZFVjQ2VUODdQNVdmeTFrcE9tR0l4OUFVNEIxMzlaUEZnMFFtcXNXcEtJK0R3QUlSZEFDQUhaZ2Mzbno1R1VQd2FnUGdieU1JWXVGblpLeFp4ZkNBR1poWTAyazBuVERoZUVEQUZFSVdnREE3dWpOSnhMMUlaQ0pwaHZHTThsLzJDZVpXL1p0L1dMSkFPRUpXcENTK2hBQUlBeEJDd0RZSGZVaFJIS3FQZ1N5NFpVL3VYdm8yOXJQVVFBejBMZjFHTFRZMmhXSnVHZ0JBSVFoYUFFQU82SStoSURVaDhETU5kM3dXUjgxQlZoYU1zQ3N1R3BCS2k1YUFBQmhDRm9Bd0c2cER5RVM5U0V3ZjY1WmtMdTc2WFUwQVBQeDNhNUk1TGpwaGlQREJ3QWlFTFFBZ04xeTlwcEl4dm9RTDM1Z3BwcHVXSTEvVExZL01yWjF6UUpnbGdRdFNFbDlDQUFRZ3FBRkFPeVEraEFDY3RVQ1ptaDZxZWNEYUhLMzd0djZ4WllCWnNjbElsTHltQUFBQ0VIUUFnQjJ6NWwzSXZsc0d6Qkw0LzhsQzZzalk4OTlXNjhzR0dCK3BwQ2NCd2FrNHFJRkFCQ0NvQVVBN0o3NkVDSTVWaDhDODlKMHcvakg0eXRySTNNdUxnSE1tNnNXcE9MM1d3QWdCRUVMQU5peHplWE4rTHJubTdrU2lBK3pZRjY4OGlkM0QzMWI2L2NIbURkQkMxSlpOTjBnYkFFQUpDZG9BUUQ3NGFvRmthZ1BnWmxvdW1FTVJwM2JGNWtUQUFTWVA0RTVVaEswQUFDU0U3UUFnUDBRdENBUzlTRXdBMDAzSExsbVFRSCs3TnY2eWFJQjVxMXZheGN0U01udnR3QkFjb0lXQUxBSDZrTUlhR2twRU43NDcvVFltc2pZVnBnSUlDc1Axa2tpRndZUEFLUW1hQUVBKytPcUJaR29ENEhBbW00NHFhcnF4bzdJM0twdjZ4ZExCc2lHK2hCU09UVjVBQ0ExUVFzQTJCOUJDeUpabk4zZkNsdEFYR3U3SVhPUGZWdjdPZ2ZJaS9vUWttbTZ3VlVMQUNBcFFRc0EyQlAxSVFRa2FBRUJUWDhrL21RM1pFNkZGVUIrQkMxSVNkQUNBRWhLMEFJQTlzdFZDeUlSdElDWXZQSW5kOS82dG5aZUhpQXpmVnMvVlZYMWJLOGtjbWJ3QUVCS2doWUFzRWVieTVzdlZWVnR6WmdnMUlkQU1FMDNYT3VZcGdDdVdRRGt5MVVMVWhHMEFBQ1NFclFBZ1AxejFZSklCQzBnaUtZYmpseXpvQUMzMDR0bkFQTGtZaEdwSERmZGNHTDZBRUFxZ2hZQXNIK0NGa1R5K2V6KzlzaEdJSVRWZUduR0tzallWcGdJSUhzdVdwQ1NxeFlBUURLQ0ZnQ3daNXZMbTYvcVF3aGs0YW9GcERlOXZ2dkRLc2pjc20vckYwc0d5RmZmMWk1YWtKS2dCUUNRaktBRkFCeUdxeFpFSW1nQjZYMnhBekwzMExlMXIzT0FNanphTTRsY0dEd0FrSXFnQlFBY2hxQUZrWHhTSHdMcE5OMHcva0g0M0FySTNNcUNBWXJocWdXcCtKa2FBRWhHMEFJQURrQjlDQUc1YWdIcGVPVlA3dTZja2djb3lzYTZTYVhwQnZVaEFFQVNnaFlBY0RpdVdoQ0pvQVVrMEhURHNxcXFZN01uWTF2WExBQ0tJMXhIU3VwREFJQWtCQzBBNEhBRUxZaEVmUWdjV05NTlJ6NkFwZ0RydnEyZkxCcWdITlAzZlJjY1NjVkZDd0FnQ1VFTEFEaVFxVDdrMmJ3SnhGVUxPS3gxVlZVTE15ZGp6MzFiQ3hNQmxFbDlDS2tJV2dBQVNRaGFBTUJodVdwQkpFdmJnTU9ZdXFPdmpKdk0rWDhGb0Z6cVEwamxkTG9jQndCd1VJSVdBSEJZWDh5YlFFN1A3bTlQTEFRT1ltM01aTzZoYjJ1QlVvQnlDVnFRa3FzV0FNREJDVm9Bd0FGdExtODI2a01JUm4wSTdGblREZU8vczNOekpuT3VXUUNVVFhVSUtWMllQZ0J3YUlJV0FIQjRYbnNTeWJWdHdONjVaa0h1N3ZxMjlnRWJRTUg2dG42cHF1clIxd0NKdUdnQkFCeWNvQVVBSEo3NkVDSlJId0o3MUhURHFxcXFZek1tWTF2WExBQ1lDTjJSaW9zV0FNREJDVm9Bd0lHcER5RWc5U0d3QjAwM25QZ0FtZ0tzcGxmTUFQQzkrQW1ReW1MNjJSc0E0R0FFTFFBZ0RmVWhST0tEWU5pUDhackZ3bXpKMkhQZjFxcHhBUGgvTGxxUWt2b1FBT0NnQkMwQUlBMzFJVVJ5ZkhaLzY0OVNzRU5OTjR6bmk2L01sTXhkV3pBQS82OXZhMEVMVWxJZkFnQWNsS0FGQUNTZ1BvU0FmRmdHdTdVeVR6TDMwTGUxRS9FQS9LOEhFeUVSandjQWdJTVN0QUNBZEp6YUpwTFB0Z0c3MFhUREdGdzZOMDR5SjZBSHdJOEk0WkdLbjc4QmdJTVN0QUNBZEw2YVBZR29ENEVkYUxyaFNKQ09BdnpadC9XVFJRUHdBK3BEU0dhcTd3TUFPQWhCQ3dCSVpITjVNMzVBOFdqK0JPSjFNbnpjc3FxcWhUbVNzYTFxSEFCK3dVVUxVdko0QUFBNEdFRUxBRWpyaS9rVGlQb1ErSUNtRzA2cXFyb3hRekszN052NnhaSUIrSkhwLzRobnd5RVJRUXNBNEdBRUxRQWdMZlVoUkRMV2h6aTFDdStuTW9UY1BmWnRMU1FLd08rb0R5RVZ2ODhDQUFjamFBRUFDYWtQSVNEMUlmQU9VeC8wSjdNamMwc0xCdUFWMUllUXluSFREVWVtRHdBY2dxQUZBS1RuWlNpUnFBK0I5M0hOZ3R4OTY5dmFCMmNBdklhTEZxU2tQZ1FBT0FoQkN3QklUMzBJa1N6TzdtK0ZMZUFObW00WVgvbWZtaG1aYzgwQ2dGY1J6Q014OVNFQXdFRUlXZ0JBWWxOOXlJTTlFSWlnQmJ6U2RKcDRaVjVrN3JadjZ5ZExCdUFOL0k1TEtvSVdBTUJCQ0ZvQVFBenFRNGhFMEFKZWJ3eFpMTXlMakQycnhnSGdIZFNIa0lycUVBRGdJQVF0QUNBRzlTRkVvajRFWHFIcGhwT3FxdjR3S3pLMzZ0djZ4WklCZUNOQkMxSlpURCtuQXdEc2xhQUZBQVN3dWJ3WlA4RDRaaGNFSW1nQnYrY2FFYmw3Nk52YTF6a0E3L0hkMUVoSWZRZ0FzSGVDRmdBUWg2c1dSSEoxZG45N1pDUHdZMDAzakgrOFBUY2VNcmV5WUFEZW8yL3JwNnFxdG9aSEl1cERBSUM5RTdRQWdEZ0VMWWpHVlF2NE9hLzh5ZDFkMzlaZUl3UHdFZjRmSVJVWExRQ0F2Uk8wQUlBZzFJY1FrS0FGL0VEVERlTXIvMk96SVdQakMrU2xCUVB3UVJzREpKRlRnd2NBOWszUUFnQmljZFdDU0Q2cEQ0Ry9hN3JoeUFmUUZHRGR0L1dMUlFQd1FTNWFrTXhVOVFjQXNEZUNGZ0FRaTZBRjBiaHFBWCszcnFwcVlTWms3TGx2NjVVRkEvQlJLcWhJN013Q0FJQjlFclFBZ0VEVWh4Q1FvQVZNbW00WS8xaDdaUjVrenNVV0FIYnAwVFJKeEVVTEFHQ3ZCQzBBSUo0dmRrSWdZMzNJaVlYQXY2eU5nY3c5OUczdHVoWUF1N1F4VFJKeDBRSUEyQ3RCQ3dBSVpuTjVNMzdBc2JVWEFuSFZndUkxM1REK096Z3ZmUTVrenpVTEFIWk5mUWlwSERmZGNHVDZBTUMrQ0ZvQVFFeGVreExKdFcxUXN1a1B0SzVaa0xzLys3YjI2aGlBWGZOL0N5bXBEd0VBOWtiUUFnQmlFclFna2xQMUlSUnVmT1YvWFBvUXlOcDRTV3RseFFEczJoVGljN0dSVk5TSEFBQjdJMmdCQUFHcER5RWc5U0VVcWVtR0UzVUtGR0RWdC9XTFJRT3dKNjVha0lxTEZnREEzZ2hhQUVCY3Jsb1FpZm9RU2pXKzhsL1lQaGw3N3R0YU5RNEErL1RkZEVuRVJRc0FZRzhFTFFBZ0xrRUxJbEVmUW5HYWJoaGZ3RjNaUEprVHBBTmczMXkwSUpWRjB3M0NGZ0RBWGdoYUFFQlE2a01JeUlkeGxNWXJmM0wzMExlMVY4WUE3SnYvYTBoSjBBSUEyQXRCQ3dDSXpWVUxJaEcwb0JoTk40eGY3NmMyVHVaOFh3ZGc3L3EyZmhtcnFreWFSQzRNSGdEWUIwRUxBSWp0aS8wUXlQSFovYTNYUUdTdjZZWWoxeXdvd0czZjFrOFdEY0NCdUdwQktuNkhCUUQyUXRBQ0FBTGJYTjU4OS9LSFlMeCtwZ1RMc2MvWnBzbllWcGdJZ0FQYkdEaUp1RklIQU95Rm9BVUF4S2MraEVnKzJ3WTVhN3JocEtxcUcwc21jOHZwakRzQUhJcWdCY2swM2FBK0JBRFlPVUVMQUloUGZRaVJxQThoZDE3NWs3dkh2cTM5YkFIQVFmVnRyVHFFbEFRdEFJQ2RFN1FBZ09BMmx6Y2I5U0VFb3o2RUxFMHYzVDdaTHBsYldqQUFpVHdZUElsNExBQUE3SnlnQlFETWcvb1FJbEVmUXE2ODhpZDMzN3dvQmlBaDlTR2tJbWdCQU95Y29BVUF6SU1QLzRoa3JBOFJ0aUFyVFRlTXIveVBiWldNYlYyekFDQXhZVDlTT1c2NjRjVDBBWUJkRXJRQWdCbFFIMEpBZ2haa28rbUdvNnFxVmpaSzV0WjlXejlaTWdBSnVXaEJTcTVhQUFBN0pXZ0JBUE9oUG9SSUJDM0l5Uml5V05nb0dSdkRtbXNMQmlDbEtmQzN0UVFTRWJRQUFIWkswQUlBNWtOOUNKRXMxSWVRZyttRThCK1dTZVpXZlZ1L1dESUFBYWdQSVpVTGt3Y0Fka25RQWdCbVlxb1BlYlF2QWhHMElBZENiT1R1b1c5clgrY0FSS0UraEZUT1RSNEEyQ1ZCQ3dDWUZ4K1VFSW1nQmJQV2RNTm5mM0NsQUN0TEJpQVFGeTFJcHVrRzlTRUF3TTRJV2dEQXZIeTFMd0laNjBPdUxZUVpXMXNlbWJ2cjI5b0hXZ0NFNGY4bEVsTWZBZ0RzaktBRkFNekk1dkxtU1gwSXdiaHF3U3cxM1RDKzhqKzJQVEsycmFwcWFjRUFCT1IzV2xKeDBRSUEyQmxCQ3dDWUgvVWhSUExwN1A3MnlFYVlrNlliam53QVRRSFdmVnUvV0RRQUFibHFRU3FDRmdEQXpnaGFBTUQ4cUE4aEdsY3RtSnV4TW1SaGEyVHN1Vy9ybFFVREVOVEdZa2prZEFwZEF3QjhtS0FGQU15TStoQUNFclJnTnBwdUdGK3hYZGtZbVhPeEJZRElCQzFJeVZVTEFHQW5CQzBBWUo3VWh4Q0oraERtWkcxYlpPNmhiMnZYcndBSXEyL3JNV2l4dFNFU3VUQjRBR0FYQkMwQVlKNThnRUkwcmxvUVh0TU4xMVZWbmRzVW1idTJZQUJtd0ZVTFVuSFJBZ0RZQ1VFTEFKaWhxVDdrd2U0SXhBZDdoRFoxTWE5c2ljejkyYmYxa3lVRE1BUGZMWWxFWExRQUFIWkMwQUlBNWt0OUNKR2NuOTNmbnRnSWdTMnJxanEySURLMkZTWUNZRVlFTFVobDBYU0QzMTBCZ0E4VHRBQ0ErVklmUWpUcVF3aHAra1BxMG5iSTNLcHY2eGRMQm1BbVZJZVFrcXNXQU1DSENWb0F3RXh0TG0vR0QxTysyUitCcUE4aHF2WDRjczEyeU5oejM5WnJDd1pnTHFadzRMT0ZrY2lad1FNQUh5Vm9BUUR6NXFvRmtaeXFEeUdhcGh2RzEycWZMSWJNQ2JvQk1FZnFRMGhGMEFJQStEQkJDd0NZTjBFTG9sRWZRalJlK1pPN2IzMWIrNkFLZ0RsU0gwSXE1eVlQQUh5VW9BVUF6Smo2RUFMeXFwb3dtbTRZdng1UGJZVE1MUzBZZ0prU0ZDU1o2ZklkQU1DN0NWb0F3UHk1YWtFazZrTUlvZW1HSTljc0tNQnQzOVpQRmczQUhQVnQ3YUlGS2FrUEFRQStSTkFDQU9aUDBJSm92SzRtZ3ZIcmNHRVRaR3dyVEFSQUJoNHNrVVFFTFFDQUR4RzBBSUNabStwRDd1eVJRRDViQmlrMTNUQmVWYm14QkRLMzdOdjZ4WklCbURsWExVaEZkUWdBOENHQ0ZnQ1FCMWN0aU9UNDdQN1c2eUJTK21MNlpPNnhiMnRmNXdEazRMc3Rrc2p4VkRjSUFQQXVnaFlBa0lITjVjM1g2WVE0UkhGdEU2VFFkTVA0TXUzYzhNbWNpaVlBY3VHaUJTbDVJQUFBdkp1Z0JRRGt3MVVMSWxFZlFpcGUrWk83dTc2dHZmNEZJQXQ5V3o5VlZmVnNteVNpUGdRQWVEZEJDd0RJaDZBRmthZ1A0ZUNhYmhoZitSK2JQQmticjFldExCaUF6TGhxUVNxQ0ZnREF1d2xhQUVBbTFJY1FrUG9RRG1icVYvWUJOTGxiVHk5L0FTQW5naGFrNG5FQUFQQnVnaFlBa0JkWExZaEUwSUpER2tNV0N4TW5ZK05aOWJVRkE1QWhsVmlrc21pNlFkZ0NBSGdYUVFzQXlJdWdCWkVzenU1dlA5c0kremI5Y2ZRUGd5WnpxNzZ0WHl3WmdOejBiUzFvUVVxQ0ZnREF1d2hhQUVCRzFJY1FrS0FGaCtDVlA3bDc2TnY2aXkwRGtMRkh5eVVSUVFzQTRGMEVMUUFnUDY1YUVJbWdCWHZWZE1QNE5YWnV5bVJ1YWNFQVpNNVZDMUs1TUhrQTREMEVMUUFnUDE1MkU0bjZFUGJOOXp4eWQ5ZTM5Y2FXQWNpYy8rdEk1ZFRrQVlEM0VMUUFnTXhzTG0vR1AxQTkyeXVCQ0Zxd0YwMDNyS3FxT2paZE1yWjF6UUtBUXJob1FUSk5ON2hxQVFDOG1hQUZBT1JKZlFpUmZENjd2ejJ5RVhhcDZZWVRIMEJUZ0hYZjFpOFdEVUR1K3JaK21nS0drTUtacVFNQWJ5Vm9BUUI1K21LdkJMSncxWUk5V0UxZlc1Q3I1NzZ0VjdZTFFFSFVoNUNLaXhZQXdKc0pXZ0JBaHRTSEVKQ2dCVHN6bmZhOU1sRXlkMjNCQUJSR2ZRaXB1R2dCQUx5Wm9BVUE1RXQ5Q0pGOFVoL0NEbm5sVCs0ZStyYjJZUk1BcGZGL0g2a2NOOTNnOTFVQTRFMEVMUUFnWCtwRGlNWlZDejZzNllieGxmKzVTWkk1MXl3QUtKSHFFRkpTSHdJQXZJbWdCUUJrU24wSUFRbGE4Q0hUS3pQWExNamRuMzFiUDlreUFLWHAyL3FscXFwSGl5Y1I5U0VBd0pzSVdnQkEzdFNIRUluNkVENXFPWjcxTlVVeXRoVW1BcUJ3cmxxUWlvc1dBTUNiQ0ZvQVFON1c5a3N3cmxyd0xrMDNuRlJWZFdONlpHNDF2ZVlGZ0ZJSldwQ0tla0lBNEUwRUxRQWdZNXZMbXllblZ3bG1hU0c4aytBWXVYdnMyOXJYT1FDbCsxNzZBRWluNlFiMUlRREFxd2xhQUVEK3Z0Z3hnWnllM2QrZVdBaHYwWFREZU1iM2s2R1JPVUUwQUlyWHQ3V0xGcVFrYUFFQXZKcWdCUURrNzZzZEU0ejZFTjdLSzM5eTk2MXZheTk0QWVEZkhzeUJSQzRNSGdCNExVRUxBTWljK2hBQ3VyWVVYcXZwaHZIcjVkVEF5SnhyRmdEd1g4S0hwT0tpQlFEd2FvSVdBRkFHOVNGRW9qNkVWMm02NGNnMUN3cHcyN2YxazBVRHdIK29EeUVWQVc4QTROVUVMUUNnRE9wRGlFWjlDSyt4cXFwcVlWSmtiQ3RNQkFEL0lHaEJNazAzcUE4QkFGNUYwQUlBQ3FBK2hJRFVoL0JMVFRlTVYwLytNQ1V5dCt6YitzV1NBZUMvcGt0UHowWkNJb0lXQU1DckNGb0FRRG5VaHhESldCK2kvNVpmOFQyTDNEMzBiZTNySEFCK3pGVUxVdkY3S2dEd0tvSVdBRkFPOVNGRTQ2b0ZQelNkNnowM0hUSzNzbUFBK0tudlJrTWlnaFlBd0tzSVdnQkFJYWI2a0cvMlRTQ2ZMWU9mOE1xZjNOMzFiZTBESkFENE9SY3RTT1Y0cWpFRUFQZ2xRUXNBS0l1ckZrUnlyRDZFLzlWMHczTDgyakFZTXJaMXpRSUFmazBna2NUOG5nb0EvSmFnQlFDVVJkQ0NhTlNIOEI5Tk54ejVBSm9DclB1MmZySm9BUGl0QnlNaWtRdURCd0IrUjlBQ0FBcXl1Yng1VVI5Q01PcEQrS3QxVlZVTEV5Rmp6OVBYT1FEd2UrcERTTVZGQ3dEZ3R3UXRBS0E4cmxvUXlWZ2Y0clVRNHpXTDhZK1pWeVpCNXBaOVc3OVlNZ0M4aXFBRnFaeWJQQUR3TzRJV0FGQWVRUXVpVVI5QzVaVS9CWGpvMjlyL3dRRHdldC9OaWxTbUlEZ0F3RThKV2dCQVlkU0hFSkQ2a01JMTNmRFpxekVLc0xSa0FIaTl2cTJmcXFyYUdobUp1THdJQVB5U29BVUFsTW1MV2lKWm5OM2ZDbHNVcXVtR0k5Y3NLTUJkMzliT253UEEyN2xxUVNvdVdnQUF2eVJvQVFCbEVyUWdHa0dMY28ydi9JOUxId0paMjdwbUFRRHZKcWhJS29JV0FNQXZDVm9BUUlHbStwQTd1eWNRUVlzQ05kMXc0Z05vQ3JEcTIvckZvZ0hnWFZ5MElKWFQ2Zm9lQU1BUENWb0FRTGxjdFNBUzlTRmxXbzI3TDMwSVpPMjViMnZWT0FEd2ZpNWFrSktyRmdEQVR3bGFBRUNoTnBjM1g2ZHo1aENGb0VWQm1tNjRxS3JxcXZRNWtMMXJLd2FBOTV1dVFqMGFJWWxjR0R3QThET0NGZ0JRTmxjdGlPVHE3UDdXYWRaeXJFb2ZBTmw3Nk52YXVYTUErRGhYTFVqRlJRc0E0S2NFTFFDZ2JJSVcvQjk3OTVNVFNaTHRDOWp5NmM3aHJnQjZCVW12QUhyaTArVEpKVmZNTW1vRkZiV0NDRlp3eVJVMHpGSXV1UjVNZmRMSkN0cFpRY0VLTHF3Z25yemJzNXFxeWorUVJJU2JtMzJmaEhyU1VrV2NFOGtmdDUrZEV4dFRMVEpRMW0xL3kvODQ5enFRUE5Nc0FHQXpCQmNaaTRrV0FNQlhDVm9BUU1hc0R5RkNnaGFKSyt1Mm4xcHlubnNkU042SHBpcnV0QmtBTnNKRUM4YXlWOWJ0b2VvREFGOGlhQUVBbUdwQlRONVpINUs4UmYvQU12Y2lrTFJIcTNFQVlIT2FxaEMwWUV5bVdnQUFYeVJvQVFBSVdoQWJVeTBTTmR3R1crWmVCNUszYUtyaVFac0JZS051bEpPUkhDazhBUEFsZ2hZQWtEbnJRNGlRb0VXNnJBd2hkYmROVlZ6b01nQnMzQ2NsWlNTQ0ZnREFGd2xhQUFEQlZBc2lZMzFJZ3NxNjdVZnV2c3U5RGlSdm9jVUFzQlhXaHpDV1k1VUhBTDVFMEFJQUNHNlpFNkc1cGlUSDl4bFNkOTFVaGR1MkFMQWRmc1l5bWlFMERnRHdPNElXQUVDL1BxUy9IWFN2RWtSRTBDSWhaZDMydC96ZjVsNEhrbWVhQlFCc1NWTVZELzVtWlVUV2h3QUFmeUpvQVFCOFpuMElNWGw3OVBIc1VFZW1yNnpiZmczTUt2YzZrTHl6cGlydXRCa0F0c3I2RU1ZaWFBRUEvSW1nQlFEdzJZVktFSmxURFVsQ0g3TFl5NzBJSk8zZWFod0EyQW5yUXhpTDFTRUF3SjhJV2dBQS8ySjlDQkd5UG1UaXlycnRwNUw4bkhzZFNONXFHR2NPQUd5WGlSYU01V0NZMUFjQThCdEJDd0RnS2V0RGlJbjFJZE5uVWc2cHUybXF3dWNjQUhhZ3FRb1RMUmlUcVJZQXdPOElXZ0FBVHprc0lqYW1Xa3hVV2JmOWc4amozT3RBOGxaYURBQTdkYVBjak9SSTRRR0Fwd1F0QUlEZldCOUNoQVF0cGt0d2k5UmR1bGtMQUR0bmZRaGpNZEVDQVBnZFFRc0E0SThjamhLVE85Mlluckp1KzF2K0I3blhnYVE5bW1ZQkFLTVFjbVFzSmxvQUFMOGphQUVBL0pHZ0JiSG9EekpQZFdOYXlycmREeUVzY3E4RHlUdHZxa0lRREFCMnowUUx4ckpYMXEyd0JRRHdHMEVMQU9CM3V0bXlQemk2VlJVaWNOck5sZzhhTVRubi9VUEkzSXRBMHU2YnFqRE5BZ0JHTUFRZEg5V2VrUWhhQUFDL0ViUUFBTDdFVkF2R2R0Yk5sc1lDVDh4d3crdDk3blVnZVNhMkFNQzQvSjNBV0FRdEFJRGZDRm9BQUY5eXBTcU02S2FiTGQwV242YnozQXRBOG02YXF2QXpFZ0RHWlgwSVl6bFJlUURnTTBFTEFPQlByQTloUlAwWTRGTU5tSjZ5YnVjaGhPUGM2MER5VExNQWdQR1phTUZZM3FvOEFQQ1pvQVVBOERYV2h6Q0cwMjYyZkZENWFTbnJkaitFWUFvSnFidHNxc0lOV2dBWVdWTVZnaGFNcHF4YlV5MEFnSDhSdEFBQXZzWm9kSGJ0ckpzdFBUU2RwdjZXLzBIdVJTQnBqNlpaQUVCVVRHQmtMRWNxRHdBRVFRc0E0R3VzRDJISGJyclowa1NFQ1Nycjl0QUJOQmxZTlZWaDJnNEF4TU9VS2NaaW9nVUE4QytDRmdEQXQxZ2Z3aTcwTjhWUFZYcXkrb0RNWHU1RklHbjNUVldjYXpFQVJNVWtQTVppb2dVQThDK0NGZ0RBdHdoYXNBdW4zV3pwcHZnRURmdUozK2RlQjVJMzEySUFpSTZKRm96bFlKanFCd0JrVHRBQ0FQaXE0ZkQ3V29YWW9yTnV0blFiYmJyYzhpZDFOMDFWK0I0RkFKRnBxcUliSnVQQkdFeTFBQUFFTFFDQTc3cFNJcmJrcHBzdFY0bzdUV1hkOXJmODMrWmVCNUpubWdVQXhNdFVDOFlpYUFFQUNGb0FBTjhsYU1FMjlMZlBUbFYybXNxNjNUZk5nZ3ljTlZWeHA5RUFFQzFUcHhqTGljb0RBSUlXQU1BM1dSL0NscHdPbnkybWFSRkMyTk03RXZZb1RBUUEwVFBSZ3JFY3F6d0FJR2dCQUR5SHFSWnMwb2R1dG5UN2JLTEt1ajBNSVN4enJ3UEpXelJWSVF3R0FISHpOd1dqS2V2VytoQUF5SnlnQlFEd0hJSVdiTXB0TjFzdVZIUFMzUEluZGJkTlZWem9NZ0RFYlFoRjNtc1RJeEcwQUlETUNWb0FBTjlsZlFnYjBvL2lQMVhNNlNycnR0OUYvQzczT3BBOFlUQUFtQTVUTFJqTGljb0RRTjRFTFFDQTV6TFZndGVhZDdQbG5TcE9tbHYrcE82NnFRb0hOZ0F3SFoxZU1SSVRMUUFnYzRJV0FNQnpYUTBUQ2VCSGZPaG1TMkdkQ1N2cnRyL2xmNUI3SFVqYW8ya1dBREE1QXBLTTVhM0tBMERlQkMwQWdHY1oxb2M0S09kSDNIYXpwY1BMQ1N2cmRqK0VzTXE5RGlUdnZLa0tVM2NBWUVLYXFqRFJndEVNcXhVQmdFd0pXZ0FBTHlGb3dVdjFOOFJQVlczeStwREZYdTVGSUduM2ZkQkNpd0Zna202MGpaRUlXZ0JBeGdRdEFJQm5HMVkvV0IvQ1M4eTcyZElOOFFrcjYvWXdoUEJ6N25VZ2VhdW1LaDYwR1FBbXlWUUx4bktrOGdDUUwwRUxBT0NsVExYZ3VUNE00UnltN1VML1NOeE5VeFUrNXdBd1haLzBqcEdZYUFFQUdSTzBBQUJleXNFNXozSGJ6WllMbFpxMnNtNzd0Uy9IdWRlQjVLMjBHQUFtelVRTHhySTNUQUFFQURJa2FBRUF2SWoxSVR4RC8vazRWYWdrbk9kZUFKSjMyVlNGVzdBQU1HRk5WZlNyQ3UvMWtKRllId0lBbVJLMEFBQitoS2tXZk11OG15M3ZWR2pheXJydGIva2Y1RjRIa3RhSHdremVBWUEwbUdyQldLd1BBWUJNQ1ZvQUFEOUMwSUt2K1RCTVBXSEN5cnJkZHdCTkJzNmJxbmpRYUFCSWdxQUZZekhSQWdBeUpXZ0JBTHlZOVNGOHhXMDNXenFjVDBPL01tUXY5eUtRdFB1bUtsWmFEQURKc0FxTXNSeXJQQURrU2RBQ0FQaFJGeXJIRTMzdzVsUkJwcStzMjM3MDdmdmM2MER5aE1JQUlDRk5WUWhhTUpxeWJrMjFBSUFNQ1ZvQUFEOUswSUtuNXQxc2VhY2lTWERMbjlUZE5GVmh4UkVBcE9kV1R4bkppY0lEUUg0RUxRQ0FIOUxObHYwTzNIdlZJNFR3WVZnbnc4U1ZkVHMzK3BZTW1HWUJBR2t5MVlLeG1HZ0JBQmtTdEFBQVhzUGhPcmZkYk9uUU1nRmwzZTZiWmtFR1BqUlYwV2swQUNUSnozakdJbWdCQUJrU3RBQUFYc1A2a0x3OWhoQk9jeTlDUXZyQXpFSHVSU0JwajhKRUFKQTBRUXZHOG5ZSXJnTUFHUkcwQUFCK21QVWgyWnQzcytWZDdrVklRVm0zaDlZcGtJRlZVeFVQR2cwQWFScW1WajFxTHlNeDFRSUFNaU5vQVFDOGx2VWhlZnJRelpaNm40N3pFTUplN2tVZ2FmZE5WWnhyTVFBa3oxUUx4bktpOGdDUUYwRUxBT0MxckEvSnoyMDNXNXAra0lpeWJ2c0hndTl5cndQSm0yc3hBR1Roa3pZekVrRUxBTWlNb0FVQThDcldoMlNuSDhWN21uc1JFdU9XUDZtN2JxckNvUXNBNU1IUGZNWmlkUWdBWkViUUFnRFlCRk10OGpIdlpzdTczSXVRaXJKdSsxditiM092QThremdRY0E4bUYxQ0dQWksrdjJVUFVCSUIrQ0ZnREFKZ2hhNU9GRE4xdGU1VjZFVkpSMXUyK2FCUms0YTZwQ09Bd0FNdEZVeFlPSmk0ekkraEFBeUlpZ0JRRHdhc09FZzF1VlROcHROMXU2Rlo2V3ZwOTd1UmVCcEQwS0V3RkFscXdQWVN6V2h3QkFSZ1F0QUlCTk1kVWlYZjFoNVdudVJVakpNTkoybVhzZFNONWl1TlVLQU9URitoREdJbWdCQUJrUnRBQUFOc1ZLaVhUTmg2a2xwRU13aXRUZE5sWGhjdzRBZVRMUmdyRWNxendBNUVQUUFnRFlDT3REa3ZXaG15MkZhQkpTMXUySkI0Qmt3S29qQU1oVVV4VW11MVNVN1FBQUlBQkpSRUZVV2pDYTRlOHRBQ0FEZ2hZQXdDYTVQWnlXMjI2MmRGaVpIdjlPU2QxbFV4VnVzZ0pBM201eUx3Q2pzVDRFQURJaGFBRUFiSkxKQitsNERDR2M1bDZFMUpSMTJ3ZG5Ebkt2QTBucnYzZXR0QmdBc2lkMHlWaE10QUNBVEFoYUFBQWJZMzFJVXVaRFAwbEVXYmY3RHFESndIbFRGYjUzQVFEV2h6QVdFeTBBSUJPQ0ZnREFwcDJyNk9SOTZHWkwwMG5TMC8vYjNNdTlDQ1R0M3M4Z0FHQWdhTUZZRG9hUU93Q1FPRUVMQUdEVEhOQlAyMjAzV3k1eUwwSnF5cnJ0YjFXOXo3ME9KRy9WVk1XRE5nTUF3NFNyKyt3THdWaXNEd0dBREFoYUFBQWIxYzJXL1NIWHRhcE8wbU1JNFRUM0lpVEtMWDlTZDlOVXhZVXVBd0JQbUdyQldLd1BBWUFNQ0ZvQUFOdGdxc1UwemJ2WjhpNzNJcVNtck5zK1BIT2NleDFJbmtrOEFNQWZmVklSUm1LaUJRQmtRTkFDQU5nR1FZdnArZERObHZxV0p0TXNTTjFsVXhWdXJBSUFmK1QzQThaaW9nVUFaRURRQWdEWU9PdERKdWUybXkzZEJrOVFXYmVyRU1KQjduVWdhWSttV1FBQVg5SlVoWWtXakdXdnJGdGhDd0JJbktBRkFMQXRwaU5NUTM5SWVacDdFVkpVMXUyaEEyZ3ljTjVVeFlOR0F3QmZjYXN3akVUUUFnQVNKMmdCQUd5TG9NVTB6THZaOGk3M0lpU3FuMmF4bDNzUlNOcDlVeFVyTFFZQXZzRlVDOFlpYUFFQWlSTzBBQUMyd3ZxUVNmalF6WllDTVFrcTYvWWtoUEErOXpxUVBCTmJBSUR2NlZTSWtad29QQUNrVGRBQ0FOZ21oL2p4dXUxbVM0ZVU2WExMbjlUZE5GWGhad3dBOEQyQ0ZvemxyY29EUU5vRUxRQ0FyZWxteTRzUXdxTUtSNmZ2eVdudVJVaFZXYmZ6RU1KeDduVWdlWE10QmdDK3BLemIvZjUzNHJKdSsxRG1QeFdKc1F5VEJnR0FSUDJYeGdJQVczWmxoVUYwNXQxc2VaZDdFVkxVUDFRMnpZSU1mR2lxd3Zjd0FPQTN3Ky9CcDhQWE81VWhFbjNRNHBObUFFQ2FCQzBBZ0cwVHRJakxoMjYyTkc0L1hmMDZtSVBjaTBEU0hvV0pBSUFnWE1FMEhPa1RBS1RyelhxOTFsNEFZS3VPUHA0OWhCRDJWSGwwdDkxczZVRlBvc3E2UFF3aC9KcDdIVWplTDAxVm5Hc3pBT1JKdUlLSnVXK3E0bERUQUNCTi8wZGZBWUFkTUVGaGZQMHQ4SG51UlVpY3cyZFNkeXRrQVFCNUt1djJ0S3piL3UvSy93MGgvRjNJZ29rNEdBTHhBRUNDckE0QkFIYkIrcER4TGJyWnNzdTlDS2txNi9iRXcyWXlzTkJrQU1oSEg2NTRNcjNDaEVTbXFwOHFlYWQ3QUpBZVFRc0FZT3U2MmZMcTZPUFpvNGRqbzduc1pzdUxUTjk3THR6eUozWFhUVlY4MG1VQVNKdHdCUWs2TXVVVEFOSWthQUVBN0lxcEZ1TzRkUXM4YldYZDlpdGgzdVplQjVMbit4Z0FKRXE0Z3NTZGFEQUFwRW5RQWdEWUZVR0wzZXVuaU15NzJmSWh0emVlaTdKdTkwMnpJQU5uVFZVWXR3d0FDUkd1SUNQSG1nMEFhWHF6WHErMUZnRFlpYU9QWi8xQjJZRnE3OHhQVm9ha3JhemJQbVR4Yys1MUlHbDlZT3l3cVFxQk1RQ1lPT0VLTXZiWHBpbzZId0FBU0l1SkZnREFMbDA1Rk42WlN5R0x0SlYxZStqZkV4bFlDRmtBd0hRSlY4Qy9ISVVRQkMwQUlER0NGZ0RBTGwwNEdONkoyLzV3TW9QM21UdEJHbEozMDFTRnp6a0FUSXh3QmZ6SmliL2ZBQ0E5Z2hZQXdNNTBzMlYzOVBIczN2cVFyZXJIN00rNzJkSU44SVNWZFh0aTF5OFpXR2t5QUV4RFdiZjlqZjM1RUs3dzl4NzgzcEY2QUVCNkJDMEFnRjJ6UG1TN0ZuMmdKZVUzeUwrNERVWHFMcHVxK0tUTEFCQXY0UXA0dHJkbDNlNWJpUWNBYVJHMEFBQjJ6ZnFRN2Juc1prc0g4SWtyNjNibFFUYUplelROQWdEaUpGd0JQNnovdHlOSURBQUpFYlFBQUhiSytwQ3R1ZTJuV1NUNjNoajB0NkQwbVF5Y04xVnhwOUVBRUFmaEN0aUlFMEVMQUVpTG9BVUFNQWJyUXphcnYvMDk3MlpMWTBqVGR4NUMyTXU5Q0NUdGZ2aWNBd0FqRXE2QWpUdFNVZ0JJaTZBRkFEQUc2ME0yYTlGUENrbnBEZkZudzhQdTkwcEQ0aFoyVndQQU9JUXJZS3RPbEJjQTB2Sm12VjVyS1FDd2MwY2Z6KzQ4dk51SXkyNjJuQ2Z3UHZpT3NtNzdNYlBINmtUQ2JwcXE4QUFhQUhaSXVBSjI2aTlXNUFGQU9reTBBQURHMG8rRy94L1ZmNVhiL3ZiM2hGOC96MVRXN2FtUUJSbncvUXdBZHFDczI4TWhYREVYcm9DZDZvTk5naFlBa0FoQkN3QmdMRmVDRnEveTJEOFk3V1pMSS9ZVFY5YnQvaEJNZ3BSZE5sVmhCUklBYk1rUXJqZ2R3aFZ2MVJsR2NUSThDd0VBRWlCb0FRQ01vcHN0NzQ0K250MTZ5UGZERnQxczZWQXlEd3MzRFVuY28ya1dBTEI1d2hVUW5TTXRBWUIwQ0ZvQUFHTzZNTlhpaDF4MnMrWEZCRjgzTHpROEhIY0FUZXJPbTZvd25RY0FOa0M0QXFKbUhTUUFKRVRRQWdBWWsvVWhMM2ZyNEQwcnF4RENYdTVGSUduM1RWV3N0QmdBZnB4d0JVeEhXYmRIVnVZQlFCb0VMUUNBMFZnZjhtTDllUDE1TjF1NitaMkJzbTc3L2IzdmM2OER5WnRyTVFDOG5IQUZURmIvZDU2Z0JRQWtRTkFDQUJpYjlTSFB0K2htU3c5azh1R1dQNm03YWFyaWt5NER3UE1JVjBBU2pyUVJBTklnYUFFQWpNMzZrT2U1N0diTGl5bThVRjZ2ck51NS9iMWt3RFFMQVBnTzRRcEl6b21XQWtBYTNxelhhNjBFQUVaMTlQR3M4OUR3bS9yMUtpZFdodVNock52OUVNSmRDR0V2OTFxUXRBOU5WU3kwR0FEK2JQaDk4SFQ0ZXFkRWtKei9icXJDMy9jQU1IRW1XZ0FBTVRnUElmeGRKNzdvc2IrOUptU1JsWVdRQllsN3RCb0hBSDVQdUFLeTBxOFBzVUlQQUNaTzBBSUFpTUdWb01WWExiclpzb3YwdGJGaHcyam9wYnFTdUpVYmZBQWdYQUVaT3hHMEFJRHBFN1FBQUViWFQyczQrbmgyN2VIaW4xeDJzK1ZGWksrSjdUcFhYeEozMjFTRnp6a0EyUkt1QUlhZ0JRQXdjWUlXQUVBc3JqeG8vSjNiWVlVRW1TanI5c1MvQVRMZyt4b0EyUkd1QVA3Z1NFRUFZUG9FTFFDQVdGZ2Y4aCtQSVlSNVAra2psaGZFVHJqbFQrcXVtNm93SWhtQUxBaFhBTit3MTYrTmJLcmlUcEVBWUxvRUxRQ0FLRmdmOGp1TGJyYnNJbm85YkZsWnQvMHQvN2ZxVE9KTXN3QWdhY0lWd0F2MEV3MnRDZ1dBQ1JPMEFBQmlZbjFJQ0pmZGJPbGhTMGFHQi9LcjNPdEE4czdjMkFNZ1ZXWGQ5c0dLdWI5bGdCZXdQZ1FBSms3UUFnQ0lTZTdyUTI3ZCtNNVNIN0xZeTcwSUpPM2VhaHdBVWpPRUt6NS8rVjBPZUNsQkN3Q1l1RGZyOVZvUEFZQm9ISDA4eTNXcXhXTS9PdFRLa0x6MGUzbERDTC9tWGdlUzkxTlRGU2IxQURCNXdoWEFKalZWOFVaQkFXQzZUTFFBQUdKemtXblFZaUZra1NXSHo2VHVSc2dDZ0NrVHJnQzJwYXpiazZZcVBpa3dBRXlUb0FVQUVKVnV0cnc2K25qMm1ObER6TXR1dG5RUW1abmhvZjF4N25VZ2VTc3RCbUJxaEN1QUhlblhod2hhQU1CRUNWb0FBREhxMTRlOHo2UXp0LzAwaXdoZUI3dDNydVlrN3RJTlBRQ21RcmdDR01HSnZ3c0JZTG9FTFFDQUdPVVN0T2duZDh5NzJmSWhndGZDRHBWMTI5L3lQMUJ6RXZab21nVUFzUk91QUVaMnBBRUFNRjF2MXV1MTlnRUEwVG42ZVBhUXdjUE9uNndNeVU5WnQvc2hoRHNQODBuY1dWTVZnaFlBUktlczIvNEcrVnk0QW9qRWZ6ZFY0ZklGQUV5UWlSWUFRS3hTbjJweEtXU1JyWE1QOVVuY3ZaQUZBREVwNi9ib1NiakNWREVnSmlmRDh3OEFZR0lFTFFDQVdLVWN0TGdOSVN3aWVCM3MyUENRUDRlMU9PVE45emNBUmlkY0FVekVrYUFGQUV5VG9BVUFFS1Z1dHJ3NituajJtT0ROLy80OXpidlowbWpRUEozblhnQ1NkOU5VaFFmRkFJeEN1QUtZb0JOTkE0QnBFclFBQUdLVzRsU0xSVGRiZGhHOERuYXNyTnYrb2YreHVwTTQweXdBMkNuaENtRGlqalFRQUticHpYcTkxam9BSUVwSEg4LzZteDMvU0tnN2w5MXNPWS9nZGJCalpkM3U5NE5hUFB3bmNaZE5WZmdlQjhEV0NWY0FpZmxyVXhVdVpBREF4SmhvQVFCRXE1c3RQeDE5UEx0UDVPSHByWnZlV1ZzNEJDQnhqNzdIQWJCTndoVkF3bzZHWUQ0QU1DR0NGZ0JBN1ByMUlUOVB2RXY5QWVTOG15MGZJbmd0N0ZoWnQ0Y09vTW5BcXFrSzMrTUEyQ2poQ2lBVC9UVFBDODBHZ0drUnRBQUFZbmVSUU5CaTBjMldicWZrYXhWQzJNdTlDQ1R0dnFtS2N5MEdZQk9Ha09xcGlXQkFSbzQwR3dDbVI5QUNBSWhhSDFDWStQcVE2MjYyZERNbFUyWGQ5amVUM3VkZUI1STMxMklBWHVOSnVLTC9tZkpXTVlITStMNEhBQk1rYUFFQVRNRlUxNGZjTzRETW5sditwTzZtcVlwUHVnekFTd2xYQVB4SEg5TDNlelVBVEl1Z0JRQXdCVk5kSDNMYXpaWVBFYndPUmxEV3JVTURjaUJNQnNDekNWY0FmRlUvRFZIUUFnQW1STkFDQUlqZVJOZUgvTksvN2doZUJ5TW82M2JmTkFzeThLR3BpanVOQnVCYmhDc0FudVZJbVFCZ1dnUXRBSUNwbU5MNmtPdHV0blRJbnJkRkNHRXY5eUtRdE1jUXdrcUxBZmdTNFFxQUZ4TzBBSUNKRWJRQUFLWmlLdXREN28zU3o5dHdzTERNdlE0a2I5RlVoZFZJQVB4R3VBTGdWUTc2NzZNbXhnSEFkQWhhQUFDVE1Ld1B1WjNBUTl2VGJyWjArSmczMDB4STNXMVRGUmU2RE1Dd0xtMHVYQUd3RWYxVUMwRUxBSmdJUVFzQVlFcjZnNzMvaWZqMS90SUhRaUo0SFl5a3JOdVRFTUk3OVNkeEN3MEd5TmNRcmpnZHZ2emVBN0E1UjhQYVZBQmdBZ1F0QUlBcHVZbzRhSEhkelpZbUdlQ1dQNm03YnFyaWt5NEQ1RVc0QW1BblRwUVpBS1pEMEFJQW1JeHV0cnlMZEgzSS9UQXVtWXlWZGR2ZjhqL3dHU0JoajZaWkFPUkR1QUpnNTQ2VkhBQ21ROUFDQUppYUdOZUhuSGF6NVVNRXI0T1JEQWNSSy9VbmNlZE5WZGdaRFpBdzRRcUFjWlYxZTlSVWhaV2tBREFCZ2hZQXdOVEV0ajdrbDI2MjlCQ0VQbVN4bDMwVlNGay91Y2Q2SklBRUNWY0FSS1ZmSCtJWkF3Qk1nS0FGQURBcGthMFB1ZTVtU3dlUG1ldHZISVVRZnM2OURpUnYxVlNGeVQwQWlSQ3VBSWpXa2RZQXdEUUlXZ0FBVXhURCtwRCtkdmQ4bXVWanc0UnRTTjFOVXhVWHVnd3diY0lWQUpNZ2FBRUFFL0YvTkFvQW1LQ3JDRjd5YVRkYnV0MmR1Ykp1KzRPSzQ5enJRUEpXV2d3d1hmM3ZLMlhkOW9HNS93MGgvRjNJQWlCcWI0ZGdIQUFRT1JNdEFJREpHZGFIM0l4NHdQMUxOMXZhbVVvd3pZSU1YRFpWOFVtakFhWmxDSU4rL3RyVFBvQko2YWRhK0IwY0FDSW5hQUVBVE5YRlNFR0w2MjYyZExoT2Y0RFIzL0kvVUFrUzloaENXR2d3d0RRSVZ3QWs0MFRRQWdEaUoyZ0JBRXpWMVRENmVKZnVRd2h6bnhpR1VhNE9vRW5kZVZNVlZpUUJSRXk0QWlCSlI5b0tBUEVUdEFBQUpxbWJMUitPUHA1ZDczakg5R24vMy9XSllWZ1o0akNEbE4wM1ZiSFNZWUQ0Q0ZjQUpPOUVpd0VnZm9JV0FNQ1VYZTB3YVBGTE4xdDJQaTJVZGRzLzlIcWZmU0ZJbllrdEFCRVJyZ0RJeWw1WnQ0ZE5WZHhwT3dERVM5QUNBSml5WGEwUHVlNW15M09mRkFadStaTzZtNllxcm5RWllGekNGUUJaNjllSENGb0FRTVFFTFFDQXlkclIrcEQ3RU1MY3A0VHc3d09QL3JOd3JCZ2t6alFMZ0pFTWs3UG13aFVBMlRzWkxwY0FBSkVTdEFBQXBtN2I2ME5PKzBDSFR3bGwzZTZiWmtFR1BqUlZZVTBTd0E2VmRYdjBKRnh4b1BZQURCTXRBSUNJQ1ZvQUFGTzN6ZlVodjNTenBRTkhQbHM0L0NCeGo4SkVBTHNoWEFIQWQ1aWtDQUNSZTdOZXIvVUlBSmkwbzQ5bjI1aHFjZDNObHFjK0dZUi9INFljOXR0cWpQQW1jYjgwVlhHdXlRRGJJVndCd0F2OXJhbUtUNG9HQUhFeTBRSUFTTUhGaG9NVzk4TkRjUGpzWE1pQ3hOMExXUUJzbm5BRkFLL1Evd3dSdEFDQVNBbGFBQUNUMTgyV1YwY2Z6eDQzZUJCKzJzMldEejRaaEg4ZmtKeHNZV0lLeEVhNERHQkRoQ3NBMkpBamhRU0FlQWxhQUFDcDZOZUh2Ti9BZS9tbG15MDdud3FlY011ZjFOMFlTUXp3T3NJVkFHekJpYUlDUUx3RUxRQ0FWR3dpYUhIZHpaWU8xZmxOV2JmOWdjbGJGU0Z4cGxrQS9BRGhDZ0MyN0tDczIvMm1La3pjQklBSUNWb0FBRW5Zd1BxUWU0ZU5QTlUvMERMTmdneWNOVlZ4cDlFQXoxUFc3ZUVRckZnSVZ3Q3dBMzJvei9RNUFJaVFvQVVBa0pMWFRMVTQ3V1pMdDBSNGF2V0s0QTVNd2FNd0VjRDNQUWxYbUhRRndLNmRDRm9BUUp3RUxRQ0FsUHhvME9LWGJyYnNmQkw0YkRoUStWbEJTTnpDR0dLQUx4T3VBQ0FTSnhvQkFIRVN0QUFBa3ZHRDYwT3V1OW5Talc3KzZFSkZTTnh0VXhVKzV3QlBDRmNBRUtFalRRR0FPQWxhQUFDcGVjbFVpL3ZoUVRyOHBxemIvc2JRc1lxUXVJVUdBd2hYQUJDOXZmNW5WVk1WZDFvRkFIRVJ0QUFBVW5QK2dxREZhVGRiR3B2UEg3bmxUK291bTZxdzV4bklsbkFGQUJOejR1OVVBSWlQb0FVQWtKUnV0dXlPUHA3MWt5b092dk8rZnVuL3Y3clBVMlhkTHA3eDJZRXA2OWNyclhRUXlJMXdCUUFUWm4wSUFFUkkwQUlBU0ZHL1B1VG5iN3l2NjI2MlBOZDVuaXJyZHQ4Qk5CazROM1lZeU1Yd3MvMTBXSmNrWEFIQVZKM29IQURFUjlBQ0FFalJ4VGVDRnZmRFRVYjRvejU4czZjcUpPeCsrSndESk90SnVLTC9lcWZUQUNSQVdCQUFJdlJtdlY3ckN3Q1FuS09QWjNkZldRSHhWeXREK0tPeWJ2dFJyUDlVR0JMM1UxTVZkanNEeVJHdUFDQURmMnVxNHBOR0EwQThUTFFBQUZMMXBmVWh2d2haOEJWdStaTzZHeUVMSUNYQ0ZRQmtwcjhjSUdnQkFCRVJ0QUFBVXZYSDlTSFgzV3pwTUowL0tldTJQNkE1VmhrU3Q5QmdZT3FFS3dESTJJa0xBZ0FRRjBFTEFDQkovZVNLbzQ5bjk4UDZrUDUvNXpyTlYzaFlSZW91bTZvd3pRZVlKT0VLQVBpWEkyVUFnTGdJV2dBQUtmdThQdVMwbXkwZmRKby9LdXQyTllSeElGV1BwbGtBVXlOY0FRQi9jdEQvZkd5cXdyTU5BSWlFb0FVQWtMSitVa0hYVDdmUVpmNm9yTnREQjlCazROekRXR0FLaENzQTRMdE9oZ3NsQUVBRUJDMEFnR1IxcytWZENPRkNoL21LZnByRm51S1FzUHVtS2xZYURNU3NyTnZUSndFTFA1Y0I0T3VPQkMwQUlCNkNGZ0FBWktlczIvNG0wSHVkSjNFbXRnQlJFcTRBZ0I5eW9td0FFQTlCQ3dBQWN1U1dQNm03YWFyQ2JUY2dHc0lWQVBCcVIwb0lBUEY0czE2dnRRTUFnR3lVZFRzUElmeGR4MG5jWDVxcXVOTmtZRXpDRlFDd2NYOXRxcUpUVmdBWW40a1dBQUJrbzZ6YmZkTXN5TUFISVF0Z0xNSVZBTEJWL1ZRTFFRc0FpSUNnQlFBQU9WbUVFQTUwbklROUNoTUJ1eVpjQVFBN2N4SkN1RkJ1QUJpZm9BVUFBRmtvNi9Zd2hMRFViUkszYXFyaVFaT0JiUk91QUlCUkhDazdBTVJCMEFJQWdGeWM2elNKdTIrcXd1Y2MySnF5Ym8rRzZWRENGUUF3anJmcURnQnhFTFFBQUNCNVpkMzI0MVhmNlRTSm0yc3dzR2xEdUdJK2hDdXMzd0tBY2QzM1A1dWJxdWowQVFER0pXZ0JBRUFPM1BJbmRkZE5WWHpTWldBVGhDc0FJQ3IzSVlTckVNS0ZnQVVBeEVQUUFnQ0FwSlYxdXpCZWxRd3NOQmw0RGVFS0FJaUtjQVVBUkU3UUFnQ0FaSlYxdXg5Q1dPa3dpVHRycXVKT2s0R1hFcTRBZ0tnSVZ3REFoQWhhQUFDUXNqNWtzYWZESk96UmFoemdKWVFyQUNBcWowL0NGVllCQXNDRXZGbXYxL29GQUVCeXlybzlEQ0g4cXJNazdxZW1LaTQwR2ZnVzRRb0FpTXJuY01WVlV4VlhXZ01BMDJTaUJRQUFxWEw0VE9wdWhTeUFyeEd1QUlDb0NGY0FRR0lFTFFBQVNFNVp0eWNoaEdPZEpYRUxEUWFlR3FZNW5RNEJpN2VLQXdDakVxNEFnSVFKV2dBQWtDSzMvRW5kcFIzT1FCQ3VBSURZQ0ZjQVFDWUVMUUFBU0VwWnR5c2owa2xjLy9CMnBjbVFMK0VLQUlpS2NBVUFaRWpRQWdDQVpKUjF1MitkQWhrNGI2cmlUcU1oTDhJVkFCQVY0UW9BeUp5Z0JRQUFLVGtQSWV6cEtBbTdIejduUUFhRUt3QWdPdGY5cWtyaENnRGd6WHE5enI0SUFBQk1YMW0zUnlHRWYyb2xpZnUvSHVwQzJvUXJBQ0E2MTArbVZ6eG9Ed0FRVExRQUFDQWhidm1UdWhzaEMwaVRjQVVBUkVlNEFnRDRKa0VMQUFBbXI2emIvbkRxV0NkSjNFS0RJUjFsM2U0UDRZcUZjQVVBUkVHNEFnQjROa0VMQUFBbWJUaW9NczJDMUYwMlZkSHBNa3piazNCRi8vVk9Pd0ZnZE1JVkFNQVBFYlFBQUdEcStwdkFCN3BJd2g1TnM0RHBFcTRBZ09nSVZ3QUFyeVpvQVFEQVpBMDc3UjFBazdwekQ0QmhXb1FyQUNBNndoVUF3RVlKV2dBQU1HV3JFTUtlRHBLdys2WXFWaG9NOFJPdUFJRG9DRmNBQUZzamFBRUF3Q1NWZFhzU1FuaXZleVJ1cnNFUUwrRUtBSWpPYlFqaFlnaFgzR2tQQUxBdGdoWUFBRXlWVy82azdxYXBpays2REhFUnJnQ0E2QWhYQUFBN0oyZ0JBTURrbEhYYjMvSS8xamtTWjVvRlJHVDQyU05jQVFCeEVLNEFBRVlsYUFFQXdLUU1ONG5QZFkzRWZmREFHTVpYMXUzcGsra1ZlMW9DQUtNU3JnQUFvaUZvQVFEQTFDd2NkcEc0UjZ0eFlEekNGUUFRRmVFS0FDQktiOWJydGM0QUFEQUpaZDBlaGhCKzFTMFM5MHRURmFhMndBNEpWd0JBVklRckFJRG9tV2dCQU1DVU9Id21kYmRDRnJBYndoVUFFQlhoQ2dCZ1VnUXRBQUNZaExKdVQwSUk3M1NMeEMwMEdMWkh1QUlBb25MZkJ5djZRTDF3QlFBd05ZSVdBQUJNeFlWT2tianJwaW8rYVRKc2xuQUZBRVRsYzdqaW9xbUtUbXNBZ0trU3RBQUFJSHBsM2ZhMy9BOTBpc1NaWmdFYklsd0JBRkVScmdBQWtpTm9BUUJBMU1xNjNROGhySFNKeEowWmx3eXZVOWJ0VVFoaFBud0pWd0RBdUlRckFJQ2tDVm9BQUJDN2xRTXpFdmZZNzZYV1pIaTVKK0dLVTVPUEFHQjB3aFVBUURiZXJOZHIzUVlBSUVwbDNSNkdFSDdWSFJMM1UxTVZGNW9NenlOY0FRQlJFYTRBQUxKa29nVUFBREZ6K0V6cWJvUXM0UHVFS3dBZ0tzSVZBRUQyQkMwQUFJaFNXYmY5WWRxeDdwQzRsUWJEbHdsWEFFQlVIb2NndkhBRkFKQzlJR2dCQUVERXpqV0h4RjAyVmZGSmsrRS9oQ3NBSUNxUHcrU0txNllxcnJRR0FPQS9CQzBBQUloT1diY3JCMndrN3RFMEMvZzM0UW9BaUlwd0JRREFNd2hhQUFBUWxiSnU5ME1JQzEwaGNlZE5WZHhwTXJrcTYvWncrRjR2WEFFQTR4T3VBQUI0SVVFTEFBQmkwNjhNMmRNVkVuYmZWSVZwRm1SbkNGZWNEdE1yM3ZvRUFNQ29oQ3NBQUY1QjBBSUFnR2dNNCtQZjZ3aUpNN0dGYkFoWEFFQlVoQ3NBQURaRTBBSUFnSmljNndhSnUvRlFtOVFKVndCQVZJUXJBQUMyUU5BQ0FJQW9sSFhiSDhnZDZ3YUpNODJDSkFsWEFFQlVoQ3NBQUxaTTBBSUFnTkdWZGJzZlFsanBCSW03YktxaTAyUlNJVndCQU5HNWZoS3dlTkFlQUlEdEViUUFBQ0FHL1MzL0E1MGdZWSttV1pBQzRRb0FpSTV3QlFEQUNBUXRBQUFZMVhCbzV3Q2ExSzA4K0dhcWhDc0FJRHJDRlFBQUl4TzBBQUJnYlAzS2tEMWRJR0gzVFZXY2F6QlRNcXgwK2h5dU9OWThBQmlkY0FVQVFFUUVMUUFBR0UxWnR5Y2hoUGM2UU9MbUdzd1VQQWxYOUYvdk5BMEFSaWRjQVFBUUtVRUxBQURHNUpZL3FidHBxdUtUTGhNcjRRb0FpSTV3QlFEQUJBaGFBQUF3aXJKdTdmb25CNlpaRUIzaENnQ0lqbkFGQU1ERUNGb0FBTEJ6d3lHZmFSYWs3a05URlhlNlRBeUVLd0FnT3JmRDMwVENGUUFBRXlSb0FRREFHQlloaEQyVkoyR1BJWVNWQmpNbTRRb0FpRTRmcnJnWXdoVUN1UUFBRS9abXZWN3JId0FBTzFQVzdXRUk0VmNWSjNFL05WVnhvY25zbW5BRkFFUkh1QUlBSUVFbVdnQUFzR3NPbjBuZHJaQUZ1MWJXN1Z5NEFnQ2lJVndCQUpBNFFRc0FBSGFtck51VEVNS3hpcE80aFFhekMyWGRuajZaWG1FZEV3Q01TN2dDQUNBamdoWUFBT3lTVy82azdycXBpays2ekxZSVZ3QkFWSVFyQUFBeUpXZ0JBTUJPbEhYYjMvSS9VRzBTWjVvRkd5ZGNBUUJSRWE0QUFFRFFBZ0NBN1N2cmRqK0VzRkpxRW5mbVlUdWJJbHdCQUZHNUR5R2NDMWNBQVBDWm9BVUFBTHV3Y2xCSTRqNC9mSWNmSmx3QkFGSHBmNys3NnFkWE5GWFJhUTBBQUUrOVdhL1hDZ0lBd05hVWRYc1VRdmluQ3BPNG41cXF1TkJrWGtxNEFnQ2lJbHdCQU1Dem1HZ0JBTUMydWVWUDZtNkVMSGdKNFFvQWlJcHdCUUFBTHlab0FRREExZ3lIaWNjcVRPSldHc3ozRE5OOTVzT1hjQVVBakV1NEFnQ0FWeEcwQUFCZ20weXpJSFdYVFZWODBtVys1RW00b2crZEhTZ1NBSXhLdUFJQWdJMFJ0QUFBWUN2S3VsMDVXQ1J4ajZaWjhFZkNGUUFRRmVFS0FBQzJRdEFDQUlDTksrdDJQNFN3VUZrU2Q5NVV4WjBtSTF3QkFGRVJyZ0FBWU9zRUxRQUEySVorWmNpZXlwS3crNllxVExQSW1IQUZBRVRsOFVtNHdsbzNBQUMyVHRBQ0FJQ05LdXYySklUd1hsVkpuSWt0R1JLdUFJQ29mQTVYWERWVmNhVTFBQURza3FBRkFBQ2I1cFkvcWJ2eE1EOGZ3aFVBRUJYaENnQUFvaUJvQVFEQXhwUjEyeDlHSHFzb2lUUE5JbkZsM1I0T2ZSYXVBSUR4Q1ZjQUFCQWRRUXNBQURhaXJOdDkweXpJd0llbUtqcU5UczhRcmpnZHBsZTh6YjBlQURBeTRRb0FBS0ltYUFFQXdLWXMzUHdtY1kvQ1JHa1JyZ0NBcUFoWEFBQXdHVy9XNjdWdUFRRHdLc05oNWErcVNPSithYXJpWEpPblRiZ0NBS0lpWEFFQXdDU1phQUVBd0NZNGZDWjE5MElXMHlWY0FRRFJ1UTRoWEFoWEFBQXdWWUlXQUFDOFNsbTNKeUdFZDZwSTR1WWFQQzNDRlFBUW5lc24weXNldEFjQWdDa1R0QUFBNExYYzhpZDFOMDFWZk5MbCtBbFhBRUIwaENzQUFFaVNvQVVBQUQrc3JGdUhtZVRBTkl1SWxYVzcveVJjY1p4N1BRQWdBc0lWQUFBa1Q5QUNBSUFmTWh4dW1tWkI2czZhcXJqVDViZzhDVmVjV2wwRUFGRVFyZ0FBSUN1Q0ZnQUEvS2hWQ0dGUDlVallvekJSUElRckFDQTZ3aFVBQUdUcnpYcTkxbjBBQUY2a3JOdkRFTUt2cWtiaWZtcXE0a0tUeHlOY0FRRFJFYTRBQUNCN3dVUUxBQUIra01OblVuY3JaREVPNFFvQWlJNXdCUUFBL0lHZ0JRQUFMMUxXN1VrSTRWalZTTnhDZzNkSHVBSUFvbk03aEt2N2NNV2Q5Z0FBd084SldnQUE4Rkp1K1pPNjY2WXFQdW55OXBWMU94ZXVBSUJvQ0ZjQUFNQXpDVm9BQVBCc1pkMzJ0L3dQVkl5RVBacG1zVjFsM1o0K21WNnhsL0o3QllBSkVLNEFBSUFmSUdnQkFNQ3pES1A5VjZwRjRzNGRNbXllY0FVQVJFVzRBZ0FBWGtuUUFnQ0E1enAzUUVyaTdvZlBPUnNnWEFFQVVSR3VBQUNBRFJLMEFBRGd1OHE2UFFvaHZGY3BFcmRxcXVKQmszK2NjQVVBUkVXNEFnQUF0a1RRQWdDQTUzRExuOVRkTkZWeG9jc3ZKMXdCQUZFUnJnQUFnQjBRdEFBQTRKdUdROVJqVlNKeEt3MStQdUVLQUloS3YvN3NxZzlIQzFjQUFNQnVDRm9BQVBBOXBsbVF1c3VtS2o3cDhyY05LNFRtdzVkd0JRQ002M080NHFLcGlrNHZBQUJndHdRdEFBRDRxckp1KzF2K0J5cEV3aDVEQ0FzTi9ySW40WXBUM3dzQVlIVENGUUFBRUFsQkN3QUF2cWlzMjBNSDBHU2dIN0g5b05IL0lWd0JBRkVScmdBQWdBZ0pXZ0FBOERVcjZ3RkkzSDFURlN0TkZxNEFnTWdJVndBQVFPUUVMUUFBK0pPeWJrOUNDTzlWaHNSbFBiRkZ1QUlBb2lKY0FRQUFFeUpvQVFEQWw3amxUK3B1bXFxNHlxM0x3aFVBRUJYaENnQUFtQ2hCQ3dBQWZxZXMyLzRROWxoVlNOdzhsd2FYZFhzNFRPOFFyZ0NBOFQzMndRcmhDZ0FBbURaQkN3QUFmbFBXN1g0STRWeEZTTnlIcGlydVVuNkxRN2ppZEFpVXZJM2dKUUZBemg2SHlSVlhPVTdVQWdDQUZBbGFBQUR3VkgvcmZVOUZTTmhqcXF0eGhDc0FJQ3JDRlFBQWtMQTM2L1ZhZndFQStIeEkrNnRLa0xoZm1xcElabXFMY0FVQVJFVzRBZ0FBTW1HaUJRQUFuMWtaUXVydVV3aFpDRmNBUUZTRUt3QUFJRU9DRmdBQTlBZTNKeUdFZHlxVzFJVm9BQUFnQUVsRVFWUkI0dVpUZlh2Q0ZRQVFGZUVLQUFESW5LQUZBQURCTkFzeWNOMVV4YWNwdlUzaENnQ0lpbkFGQUFEd0cwRUxBSURNbFhXN2NJaExCaFpUZUl0bDNlNFA0WXBUVTJZQVlIVENGUUFBd0JjSldnQUFaR3c0MUYzNURKQzRzNllxN21KOWk4SVZBQkNkNnlGY2NhRTFBQURBbHdoYUFBRGtyUTlaN09WZUJKTDJHT05xSE9FS0FJak85WlBwRlEvYUF3QUFmTXViOVhxdFFBQUFHU3JyOWpDRThLdmVrN2lmWXJtTktsd0JBTkVScmdBQUFINklpUllBQVBreUNwblUzWTRkc2hDdUFJRG9DRmNBQUFDdkptZ0JBSkNoc201UFFnakhlay9pRm1POFBlRUtBSWlPY0FVQUFMQlJnaFlBQUhreXpZTFVYVFpWOFdtWDc3R3MyejVZTVJldUFJQW9DRmNBQUFCYkkyZ0JBSkNac201WElZUURmU2RoanlHRTFTN2UzaEN1K1B5MTUwTUZBS01TcmdBQUFIWkMwQUlBSUNQRFNvTlIxaW5BRHAwM1ZYRzNyZitjY0FVQVJPVzIvOWt2WEFFQUFPeVNvQVVBUUY3T0hReVR1UHZoYzc1UndoVUFFSlhiWVJYZTFUYkRsUUFBQUY4amFBRUFrSW15Ym85Q0NPLzFtOFN0Tm5XYlZiZ0NBS0lpWEFFQUFFUkQwQUlBSUI4YnYrVVBrYmxwcXVMaU5TOUp1QUlBb2lKY0FRQUFSRW5RQWdBZ0E4UGg4YkZlazdqRmo3dzk0UW9BaUlwd0JRQUFFRDFCQ3dDQXhKVjF1MithQlJtNGJLcWllKzdiSEZicHpJZHd4WUVQQ0FDTVNyZ0NBQUNZRkVFTEFJRDBMUndrazdqSDUweXpFSzRBZ0tnSVZ3QUFBSk1sYUFFQWtMQ3liZzkvZEowQ1RNaDVVeFVQWDNxNXdoVUFFQlhoQ2dBQUlBbUNGZ0FBYVZ1RkVQYjBtSVRkTjFXeGV2cjJoQ3NBSUNyM2ZiQ2lEMWk4Wk0wWEFBQkF6QVF0QUFBU1ZkYnRTUWpodmY2U3VENVFJVndCQUhFUnJnQUFBSkwyWnIxZTZ6QUFRSUxLdXUwZmFyL1ZXeEwyK1JCSHVBSUF4aWRjQVFBQVpFUFFBZ0FnUVdYZDlqZjcvNjYzQUFCc2tYQUZBQUNRSlVFTEFJREVsSFc3SDBLNEN5SHM2UzBBQUJzbVhBRUFBR1R2djNJdkFBQkFnaFpDRmdBQWJKQndCUUFBd0JNbVdnQUFKS1NzMjhNUXdxOTZDZ0RBS3dsWEFBQUFmSVdKRmdBQWFUblhUd0FBZnREamszREZKMFVFQUFENE1oTXRBQUFTVWRidFNRamhIL29KQU1BTGZBNVhYRFZWY2FWd0FBQUEzMmVpQlFCQU9pNzBFZ0NBWnhDdUFBQUFlQVZCQ3dDQUJKUjF1d2doSE9nbEFBQmZJVndCQUFDd0lWYUhBQUJNWEZtMyt5R0V1eERDbmw0Q0FQQ0VjQVVBQU1BV21HZ0JBREI5S3lFTEFBQUd3aFVBQUFCYlpxSUZBTUNFbFhWN0dFTDRWUThCQUxJbVhBRUFBTEJESmxvQUFFemJoZjRCQUdUclVyZ0NBQUJnOXdRdEFBQW1xcXpiMHhEQ3NmNEJBR1RsK3NuMGlnZXRCd0FBMkQxQkN3Q0E2VHJYT3dDQUxBaFhBQUFBUkVUUUFnQmdnc3E2WFlVUUR2UU9BQ0Jad2hVQUFBQ1Jlck5lci9VR0FHQkN5cnJkRHlIY2hSRDI5QTBBSUNuQ0ZRQUFBQk5nb2dVQXdQU2NDMWtBQUNSRHVBSUFBR0JpVExRQUFKaVFzbTZQUWdqLzFETUFnRWtUcmdBQUFKZ3dFeTBBQUtibFhMOEFBQ1pKdUFJQUFDQVJnaFlBQUJOUjF1MDhoSENzWHdBQWszRWJRcmdZd2hWMzJnWUFBSkFHcTBNQUFDYWdyTnY5RUVJWFFqalFMd0NBcUFsWEFBQUFKTTVFQ3dDQWFWZ0lXUUFBUkV1NEFnQUFJQ01tV2dBQVJLNnMyOE5obXNXZVhnRUFSRU80QWdBQUlGTW1XZ0FBeE85Y3lBSUFJQXJDRlFBQUFKaG9BUUFRczdKdVQwSUkvOUFrQUlEUkNGY0FBQUR3T3laYUFBREU3VngvQUFCMlRyZ0NBQUNBcnhLMEFBQ0lWRm0zOHhEQ1cvMEJBTmlKK3lGY2NTRmNBUUFBd0xkWUhRSUFFS0d5YnZkRENQMEQvajM5QVFEWW1qNWNjVFdFS3pwbEJnQUE0RGxNdEFBQWlOTkN5QUlBWUN1RUt3QUFBSGdWRXkwQUFDSlQxdTFoQ09GWGZRRUEyQmpoQ2dBQUFEYkdSQXNBZ1BoYzZBa0F3S3NKVndBQUFMQVZnaFlBQUJFcDYvWWtoSENzSndBQVAwUzRBZ0FBZ0swVHRBQUFpSXRwRmdBQUx5TmNBUUFBd0U0SldnQUFSS0tzMjBVSTRVQS9BQUMrNjNFSXFBcFhBQUFBc0hOdjF1dTFxZ01Bakt5czIvMFF3bDBJWVU4dkFBQys2SEdZWEhIVlZNV1ZFZ0VBQURBV0V5MEFBT0t3RXJJQUFQZ1Q0UW9BQUFDaVk2SUZBTURJeXJvOUNpSDhVeDlJM0ZrSTRUQ0VjQ3BVQk1CM0NGY0FBQUFRTlJNdEFBREdkNjRISk82bXFZclY1N2RZMXUzcEVMZ1F1Z0RnUy9xZkRROGhoRStxQXdBQVFJeE10QUFBR05GdzRQei85SURFL2JXcGl1NlBiN0dzMi8wbmdZdDNQZ1FBL01GOUNHSGVWSVhBQlFBQUFGRVJ0QUFBR0ZGWnQzY2hoQU05SUdHWFRWWE12L2YyaEM0QStJWVBJWVJWVXhVUGlnUUFBRUFNQkMwQUFFWlMxbTIvU21HcC9pU3MzN0YvK05LRHNTZWhpMFVJNGEwUENBQWhoTnRodXNXZkppUUJBQURBcmdsYUFBQ01vS3pid3hCQ04rd2doMVNkTlZXeGVzMTdHLzZ0OUtHTHVkQUZBSnY0MlFJQUFBQ3ZKV2dCQURDQ3NtNHZRZ2p2MVo2RTNUZFZjYmpKdHlkMEFjREFkQXNBQUFCR0pXZ0JBTEJqWmQyZWhCRCtvZTRrN205TlZYemExbHNzNi9ab0NGejB3WXNESHlhQTdQVHJxVlpOVlp4clBRQUFBTHNtYUFFQXNHTmwzZmFIejhmcVRzSnVtcW80MmRYYkU3b0F5TnJOTU4zaUx2ZENBQUFBc0R1Q0ZnQUFPMVRXYlg4WS9IYzFKM0YvR2V2QWF3aGRMSWJReFo0UEdrQVdIb2V3eFpWMkF3QUFzQXVDRmdBQU8xTFc3WDRJb1hQam5zUjlhS3BpRWNOYkxPdjJkQWhjQ0YwQTVPRjZDRnc4NkRjQUFBRGJKR2dCQUxBalpkMnVRZ2hMOVNaaC9ZM2l3eGdQdUlRdUFMSnhQNFF0UG1rNUFBQUEyeUpvQVFDd0EyWGRIb1lRZmxWckV2ZExVeFhuTWIvRlliTE01OERGdXdoZUVnRGJjZFpVeFVwdEFRQUEyQVpCQ3dDQUhTanI5c3FoTG9tN2JhcmlhRXB2VWVnQ0lIbTMvZmY0cGlydXRCb0FBSUJORXJRQUFOaXlzbTVQUWdqL1VHY1M5N2NwajJsL0VycVloeENPSTNoSkFHeEd2OVpxMFZURmhYb0NBQUN3S1lJV0FBQmJWdFp0RjBKNHE4NGs3THFwaXROVTN0Nnc2dWR6Nk1LL1hZQTBYUGZmMTV1cWVOQlBBQUFBWGt2UUFnQmdpOHE2N1E5cS82N0dKTzR2cVk1bEY3b0FTTXI5c0VxazAxWUFBQUJlUTlBQ0FHQkxobFVFL2VIem5ocVRzTE9tS2xZNU5IZ0lYU3lHNE1WQkJDOEpnQitUemM4dUFBQUF0a1BRQWdCZ1M4cTZQUThoL0t5K0pLemZlMytZNHhqMnNtNlBoaWtYUWhjQTAzUXpUTGV3U2dRQUFJQVhFN1FBQU5pQzRlYjdyMnBMNG41cXF1SWk5eVkvQ1YzTVRiQUJtSlRISVd6eFNkc0FBQUI0Q1VFTEFJQXRLT3UyZjJCL3JMWWs3S2FwaWhNTi9yMnliaytIS1JlblFoY0FrMkdWQ0FBQUFDOGlhQUVBc0dGbDNmYUh6LzlRVnhMM056ZUF2MDNvQW1CU3JCSUJBQURnMlFRdEFBQTJyS3pidXhEQ2dicVNzTXVtS3VZYS9IeGwzYzZId01XN3FieG1nQXhaSlFJQUFNQ3pDRm9BQUd4UVdiZUxFTUwvcUNrSjZ3K2hqcHFxdU5Qa2x5dnJkdi9KbEF1aEM0QTRXU1VDQUFEQU53bGFBQUJzeUhDQWVtZEZBSWx6K0xRaFQwSVgvYlNMNHlUZUZFQTZyQklCQUFEZ3F3UXRnUC9QM3QwazFYR2tiUU5PTzk0NWZDc0FyMEI0QmFCSlRvV2pJbklxdEFMakZmaG9CWTFXWUpoV1JNWXJUWE5pV0VIRENndzdFQ3ZnaS9LYmRzdHVTUzUrRHFjcTY3b2l2SUIrYmhxT1R0NzVKQUJQcE92TGFRamh0WG5Tc0p1YzRxNkFuMTdYbDkxUFNoY3ZXdnZmQnpCVG5oSUJBQURnc3hRdEFBQ2VRTmVYdlJEQ3Y4MlN4djJRVTN3djVQVlN1Z0NZbko5eWlpZGlBUUFBNEErS0ZnQUFUNkRyeTduVi96VHVJcWQ0SU9UblZVc1h4N1Y0c2JPay8rMEFFL05oS01CNVNnUUFBSUNnYUFFQThIaGRYNFlEMFA4MVNocjNmVTd4VXNpYlV6Zm5IQ2xkQUd6TVRYMUt4TjlEQUFDQWhWTzBBQUI0aEs0djJ5R0VTNGVlTk80c3AzZ2s1T240cEhReC9MZTE5SGtBUEtQYllkTlFUdkhVMEFFQUFKWkwwUUlBNEJHNnZxeENDRCtiSVEwYkRwUjJyVXFmcnJwVjU0Ly9sQzRBbnNlN25PS3hXUU1BQUN5VG9nVUF3QU4xZmRtdDJ5d2NiTkt5bjNLS0p4S2VCNlVMZ0dkMUZVSTRVRVlFQUFCWUhrVUxBSUFINnZveXJJeCtiWDQwN0NhbnVDdmdlZXI2Y2xRTEY2K1dQZ3VBTmJxdFpZdExRd1lBQUZnT1JRc0FnQWZvK25JUVF2alY3R2pjeTV6aXVaRG5yZXZMOWlkYkxwUXVBTmJqVFU3eDFHd0JBQUNXUWRFQ0FPQUJ1cjRNaDgvN1prZkRMbktLQndKdXl5ZWxpeU8vd3dDZTNGbE84Y2hZQVFBQTJxZG9BUUJ3VDNVZC95L21SdU8reXlsZUM3bGRYVjkyUHlsZHZGajZQQUNleUZWOVN1U2pnUUlBQUxSTDBRSUE0QjdxYmZEaDhIbkwzR2pZdTV6aXNZQ1hRK2tDNEVuZDFyTEZwYkVDQUFDMFNkRUNBT0FldXI2c1FnZy9teGtOR3c2SGR0M0VYYTVhdWppdXhZdWRwYzhENElHR3Y2ZkhPY1ZUQXdRQUFHaVBvZ1VBd0VqMThQRTM4Nkp4Ynh3SzhZZXVMM3QxeTRYU0JjREQyQklGQUFEUUlFVUxBSUNSdXI2OER5RzhNaThhZHBWVDNCTXduL05KNmVMSTgwa0E5L0poK04xcFd4UUFBRUE3RkMwQUFFYm8rbklRUXZqVnJHamN5NXppdVpENUoxMWZEdXVXaTBPbEM0QlJyb2JmbVRuRmErTUNBQUNZUDBVTEFJQVJ1cjVjaGhCZW1CVU4rNUJUUEJRdzk2VjBBVERhYlFqaElLZDRhV1FBQUFEenBtZ0JBUEFQdXI0TTcyci95NXhvMkhEd3MrZVdMWS9WOWVXb0ZpNDhzd1R3Wlc5eWlxZm1Bd0FBTUYrS0ZnQUFYOUgxWlR1RWNPMldObzE3bTFOY0NabW5VbjkzSGlwZEFIelJ1NXppc2ZFQUFBRE1rNklGQU1CWGRIMDVDU0g4YUVZMDdLWnVzL2dvWk5aQjZRTGdpejZFRUk3OERRWUFBSmdmUlFzQWdDL28rckliUXZqTmZHaWM5ZVU4bS9wN2RTaGNERStNdkRCNWdIQVZRamhRdGdBQUFKZ1hSUXNBZ0MvbytuSWVRdGczSHhwMmtWTThFRENib0hRQjhLZmJXcmE0TkJJQUFJQjVVTFFBQVBpTXJpL0Q0Zk92WmtQalh1WVV6NFhNcHRYU3hWSDliMGNnd0FJTlpZdERmNWNCQUFEbVFkRUNBT0F6dXI1Y08reWpjV2M1eFNNaE16VmRYL1pxNGVMUTcyRmdnVHpwQlFBQU1BT0tGZ0FBZjlQMVpSVkMrTmxjYU5od2EzYlhlL0JNbmRJRnNGRHZjb3JId2djQUFKZ3VSUXNBZ0U5MGZka09JUXpiTExiTWhZYTl6U211Qk15Y2RIMDVySVdMUTcramdRV3dlUW9BQUdEQ0ZDMEFBRDdSOVdWWTFmemFUR2pZVFU1eFY4RE1tZElGc0JCWElZUURHNmdBQUFDbVI5RUNBS0NxSytyL2JSNDA3b2VjNG5zaDA0cGF1aGh1ZmI4U0t0QWdaUXNBQUlBSlVyUUFBS2k2dnB5SEVQYk5nNFpkNUJRUEJFeUw2dE5QZjJ5NVVMb0FXbkpieXhhWFVnVUFBSmdHUlFzQWdQODdvQnR1US85aUZqVHVlNGMwTElIU0JkQWdaUXNBQUlBSlViUUFBQmF2SHNnTlgxcnZMSDBXTk8xZFR2Rll4Q3hOMTVmZFdyZ1lDblV2L0FBQU16YVVMWTV6aXFkQ0JBQUEyQ3hGQ3dCZzhicStyRUlJUHk5OURqUnRPSmpaOWI0N1M2ZDBBVFRpamJJRkFBREFaaWxhQUFDTFZnL2RobTBXVzB1ZkJVMzdLYWQ0SW1MNGovcjcvNmorWjZNUk1EZit0Z01BQUd5UW9nVUFzR2hkWDRiYmdLK1hQZ2VhZHBOVDNCVXhmRm5YbDcxYXVEaFV1Z0JtNUN5bmVDUXdBQUNBNTZkb0FRQXNWdGVYZ3hEQ3IzNENhTnpMbk9LNWtHRWNwUXRnWnBRdEFBQUFOa0RSQWdCWXJLNHZsOTdvcDNFWE9jVURJY1BEZEgwNXJJV0xRMDlNQVJOMkZrSTR6aWwrRkJJQUFNRHpVTFFBQUJhcDY4dHc4KzhYNmRPNDczS0sxMEtHeDZ0L040YkN4U3ZqQkNib0tvUndvR3dCQUFEd1BCUXRBSURGNmZxeUhVSzRkanVaeHIzTkthNkVERStyL2cwNXF2L1ppZ1JNaWJJRkFBREFNMUcwQUFBV3ArdkxjUGo4cytScDJHMElZZGRCQzZ4WDE1ZTlUMG9YeW52QUZDaGJBQUFBUEFORkN3QmdVYnErN0lZUWZwTTZqWHVUVXp3Vk1qd2ZUNHNBRTZKc0FRQUFzR2FLRmdEQW9uUjllZThRak1aZDVSVDNoQXliVVF0OWYyeTUyQkVEc0NHM3RXeHhLUUFBQUlDbnAyZ0JBQ3hHMTVlREVNS3ZFcWR4TDNPSzUwS0d6ZXY2Y2xnTEZ3cCt3Q1lvV3dBQUFLeUpvZ1VBc0JoZFg2N2RMcVp4WnpuRkl5SER0Tmh5QVd5UXNnVUFBTUFhS0ZvQUFJdlE5ZVU0aFBBdmFkT3c0U0JsTDZkNExXU1lycTR2ZnhRdTlzVUVQQk5sQ3dBQWdDZW1hQUVBTksvcnkzWUlZVGg4M3BJMkRYdWJVMXdKR09haGJyazRycVVMZjUrQWRWTzJBQUFBZUVMZkdpWUFzQUFyaDFnMDdpYUVjQ0prbUk5aCsweE9jU2hhRElXTE4vWC94d0RyTW53V1B1LzZzbWZDQUFBQWoyZWpCUURRdEhwaitEY3AwN2czT2NWVEljTzhkWDA1cUZzdVhva1NXQk9iTFFBQUFKNkFvZ1VBMExTdUwrZmV3YWR4RnpuRkF5RkRPendyQXF5WnNnVUFBTUFqS1ZvQUFNM3ErbklZUXZoZkNkTzQ3eDJVUUp1NnZteUhFQTdyRTFnN1lnYWVrTElGQUFEQUl5aGFBQURONnZweTdXQ0t4cDNsRkkrRURPMnI1Y0ZqVzVxQUo2UnNBUUFBOEVEZkdod0EwS0t1TDI3LzBycmJldWdLTEVCTzhYMTlKdWo3b1dRbGMrQUpERThUblhkOTJUTk1BQUNBKzdIUkFnQm9UbDIxZnUxZGV4cjNOcWU0RWpJc1U5ZVgzUkRDVVMxYytYc0hQSWJORmdBQUFQZWthQUVBTktmcnkya0k0YlZrYWRoTlRuRlh3RUF0Ri81UnVMREpDWGdvWlFzQUFJQjdVTFFBQUpyUzlXVllxLzZyVkduY0Q4TXpBa0lHUHRYMTVZL0N4UXVEQVI1QTJRSUFBR0FrUlFzQW9DbGRYODVEQ1B0U3BXRVhPY1VEQVFOZlVrdUhLMzhQZ1FkUXRnQUFBQmhCMFFJQWFFYTl5ZnVMUkduY2R6bkZheUVELzZRV0xvNDhwd1hjMDFDMjJQTjVBd0FBNE1zVUxRQ0FKdFEzNmkrOVQwL2ozdVVVajRVTTNFZlhsOTI2NFVMaEFoanJxbTYyK0doaUFBQUEvKzFiTXdFQUduR3NaRUhqYnV0QktjQzlETGZTYzRyRFpvdnZRZ2huOWZjSndOZThDQ0djMXpJekFBQUFmMk9qQlFBd2UvV203ckROWWt1YU5PeW5uT0tKZ0lISHFnZW54L1UvZnp1QnI3SFpBZ0FBNERNVUxRQ0EyZXY2OGo2RThFcVNOT3dxcDdnbllPQXBLVndBSXlsYkFBQUEvSTJuUXdDQVdldjZjcUJrd1FJY0N4bDRhc09oYVU1eGVKSm8yQXoxMXBNaXdCY016NGpZcWdVQUFQQUpHeTBBZ0ZucituSlp2L3lGVm4zSUtSNUtGMWczR3k2QWYzQ1dVend5SkFBQUFFVUxBR0RHdXI0TVgvVCtJa01hOTExTzhWckl3SE9waFl1aDREVnN1OWd4ZU9BVHloWUFBTURpQlVVTEFHQ3U2aUhRdFJ1M05PNXRYZXNQc0JHMTFLaHdBWHpxcDV5aXAwUUFBSUJGKzNicEF3QUFac3RhYzFwMzZ6MTBZTk55aXFjNXhkMmgrRlYvTHdIOHE1YXdBQUFBRnN0R0N3QmdkcnErREFjK3YwbU94cjBaRGppRkRFeEYzU1oxck93SVZENnJBQUFBaTZWb0FRRE1UdGVYOHhEQ3Z1Um8yRVZPOFVEQXdCUXBYQURWc09YbUlLZDRhU0FBQU1EU0tGb0FBTFBTOVdVNGZQNVZhalR1WlU3eFhNakFsQ2xjQU1vV0FBREFVaWxhQUFDejB2WGxPb1N3SXpVYWRwWlQ5TzQ1TUJzS0Y3QjROeUdFdlp6aXg2VVBBZ0FBV0k1dlpRMEF6RVhYbDJNbEN4bzMzQXBkQ1JtWWsrRndOYWM0L083YUc4cGl3b1BGR1Q2Zm45ZlNGUUFBd0NJb1dnQUFzMUMvdUhVQVRldE9jb3JYVWdibWFQajlWVGZ5Zktkd0FZdnpJb1J3S25ZQUFHQXBGQzBBZ0xsWVdVZE80NGExMnlkQ0J1YnViNFdMQzRIQ1lyenErcUpzQVFBQUxNSTNkM2Qza2dZQUpxM3J5N0NLL045U29uRS81QlRmQ3hsb1RkZVhnMXFZM0JjdUxNSlBPVVhsVVFBQW9HbUtGZ0RBNUhWOU9YYzRRK011Y29vSFFnWmExdlhsc0c3dTJSRTBOTzlOVHRGMkN3QUFvRm1LRmdEQXBOVkRtZitWRW8zN1BxZDRLV1JnQ2JxK0hOVU5Gd29YMEs3YkVNS0J6emNBQUVDcnZwVXNBREJ4MWc3VHVqT0hFTUNTMUZ2dXc3TmdiK3RoTE5DZXJSRENlZGVYWGRrQ0FBQXRVclFBQUNhcjY0dmJyclJ1T0dBOGxqS3dORG5GanpuRjRlLzhjQWg3NWdjQW1qU1VMZDUzZmRrV0x3QUEwQnBGQ3dCZ2t1cnROd2ZRdEc0MUhEWktHVmlxV3JnWW5oTDVMb1J3NFFjQm12TWloSEFxVmdBQW9EV0tGZ0RBVkszcUxUaG8xVTFPMGRNNEFQOVh1TGpPS1I2RUVGNkdFSzdNQkpyeXF1dUx6endBQUVCVEZDMEFnTW5wK2pJY3RMeVdESTA3RWpEQVgrVVV6M09LZXlHRU4wTWh6WGlnR1Q5MmZmSFpCd0FBYUlhaUJRQXdSU3VwMExpTDRUQlJ5QUNmbDFNY25ob1lDaGR2UXdpM3hnUk4rS1hyeTU0b0FRQ0FGbnh6ZDNjblNBQmdNdXBOdDE4a1F1TytHOWJrQ3huZ24zVjkyYTBsVE51dVlQNkc0dFJ1VHZHakxBRUFnRG16MFFJQW1JeXVMOXUyV2JBQTc1UXNBTVliZm1mbUZJY2k1c3RoSTVEUndheHRoUkJzOVFJQUFHWlAwUUlBbUpMakVNS09SR2pZclRJUndNTU1UeTdsRkE5Q0NHODhKd0t6OXFMcnk2a0lBUUNBT2ZOMENBQXdDWFV0K0cvU29IRnZjb29PRmdBZXFXN0JHZ3FhUDVzbHpOWlBPY1VUOFFFQUFIT2thQUVBVEVMWGwvY2hoRmZTb0dGWE9jVTlBUU04blZyVUhBcHMrOFlLcy9SeTJGWWpPZ0FBWUc0VUxRQ0FqZXY2TXF3Qi8xVVNOTTVCQXNDYTFNOFNwNTRnZzlrWm5nSGF6U2wrRkIwQUFEQW4zMG9MQUpnQUs0TnAzUWNsQzREMUdYN0g1aFNIN1JadjY4RXRNQTliSVFTZmtRQUFnTmxSdEFBQU5xcnJ5MUVJNFlVVWFOeXhnQUhXTDZlNENpRU16elI5TUc2WWpSZGRYMDdGQlFBQXpJbW5Rd0NBamVuNnNoMUN1SzQzMmFCVmIrdkJId0RQeUhNaU1EdHZjb29LRndBQXdDellhQUVBYk5KS3lZTEczWGdhQjJBei92YWNDREI5SjExZjl1UUVBQURNZ1kwV0FNQkdkSDBaRGo1K00zMGE1Mlltd0FUVXp4M0Q3K045ZWNDa1hZVVFEbktLSDhVRUFBQk1tWTBXQU1DbU9IeW1kUmRLRmdEVGtGTzh6aWtPVDRuOFVMY05BZFAwd3I4VEFBQ0FPVkMwQUFDZVhYMHozWTFTV3JlU01NQzA1QlRmaHhDR3B3bmVpUVltNjFYWGwyUHhBQUFBVSticEVBRGcyWFY5dVE0aDdKZzhEVHZMS1I0SkdHQzZ1cjdzMVp2ekw4UUVrL1I5VHZGU05BQUF3QlRaYUFFQVBLdXVMeXNsQ3hwM2E1c0Z3UFFOQjdnNXhhRnM4VlA5M1ExTXkvdXVMOXN5QVFBQXBralJBZ0I0TnZXTFVtdUFhZDFKVHZGYXlnRHprRk04cWMrSlhJZ01KbVduYnAwQkFBQ1lIRStIQUFEUHB1dkw4RVhwYXhPbllUYzV4VjBCQTh4VDE1ZkRlckM3SlVLWWpEYzVSWVVMQUFCZ1VteTBBQUNlUlgwSFhjbUMxdG5ZQWpCak9jWDNJWVNoTUhjbVI1aU1rL3B2Q1FBQWdNbXcwUUlBZUJaZFg4NURDUHVtVGNNdWNvb0hBZ1pvUTllWGc3cmRZa2Vrc0hGWElZU0RuT0pIVVFBQUFGTmdvd1VBc0haMURiZVNCYTJ6elFLZ0lUbkZvU1E2M0tKL0oxZll1QmNoaEpVWUFBQ0FxYkRSQWdCWXE2NHYyeUdFUzdkQmFkeTduS0tpQlVDamJMZUF5ZmloUHZFREFBQ3dVVFphQUFEcmR1eFFnc2JkdW1FSjBMWlB0bHU4RlRWczFHa3RjZ01BQUd5VW9nVUFzRFpkWDNZOXA4QUNyTHdYRHRDKzRYZDlUbkVvMW4wZlFyZ1NPV3pFVmdqQlJnc0FBR0RqRkMwQWdIVmExUzlEb1ZVM09jVVQ2UUlzUjA3eE1xZG91d1Zzem43WEYyVnVBQUJnb3hRdEFJQzFxRytadnpaZEduY2tZSUJsc3QwQ05tclY5V1ZQQkFBQXdLWW9XZ0FBNjdJeVdScDNVZC9zQjJDaGh1MFdJWVNoWFByT3p3QThxMkZyM3FtUkF3QUFtNkpvQVFBOHVhNHZ3eTMvZlpPbGNiWlpBRENVTFQ3bUZJZG5ERjRPVDBxWkNEeWJGMTFmbExzQkFJQ04rT2J1N3M3a0FZQW4wL1ZsTzRSd1hXK1pRYXZlMXBYeEFQQ24ramxvK1B2d282bkFzL20rYnBjQkFBQjROalphQUFCUDdWakpnc2JkaGhCT2hBekEzMzJ5M2VLSCt2Y0NXRDlQaUFBQUFNOU8wUUlBZURKZFgzWkRDRCtiS0kwN0hnN1NoQXpBbCtRVTM0Y1FoczlGSHd3SjFtNTRRa1FKRmdBQWVGYWVEZ0VBbmt6WGwrRlE0WldKMHJDcm5PS2VnQUVZcSt2TGNYMU94TVl2V0srWE9jVnpNd1lBQUo2RGpSWUF3SlBvK25LZ1pNRUNIQXNaZ1B2SUtRNDM3WWVTM3BYQndWcWRkbjNaTm1JQUFPQTVLRm9BQUUvRnVsNWE5OEV0U1FBZUlxZDRYVGNpdlRWQVdKc2RwVmdBQU9DNUtGb0FBSTlXVjJLL01Fa2FkdXVMZXdBZUs2YzRQQ0h5TW9Sd1k1aXdGajkzZmZITUd3QUFzSGFLRmdEQW85VDF2Q3RUcEhFbncyMWtJUVB3V0hVNzBuQVEvTUV3WVMxT2pSVUFBRmczUlFzQTRMR0drc1dXS2RLd0cwL2pBUENVY29vZmM0cUhJWVEzZFdzUzhIUmVkSDFSQkFjQUFOYnFtN3U3T3hNR0FCNms2OHR1Q09FMzA2TnhiM0tLYmtZQ3NCYjFtWU5UejdEQmsvcytwM2hwckFBQXdEcllhQUVBUEliRFoxcDNvV1FCd0RyVmcrQ0RFTUk3ZzRZblpTTVpBQUN3Tm9vV0FNQ0RkSDBaMWwzdm14Nk5PeFl3QU90V254SVovdWI4NENrUmVETDdYVjk4bGdNQUFOWkMwUUlBZUNnM3hHamRtWFhUQUR5bm5PTDdFTUx3bE1pVndjT1RXSFY5MlRaS0FBRGdxU2xhQUFEMzF2VmxGVUxZTVRrYWRtdWJCUUNia0ZPOHppbnVlVW9FbnNTVzV3NEJBSUIxVUxRQUFPNmwzZ2h6QUUzclRvWTE3bElHWUZNOEpRSlA1bFhYbHdQakJBQUFucEtpQlFCd1h5ZjFaaGkwNmlhbnVKSXVBSnRXbnhJNThKUUlQTnFwSjBRQUFJQ25wR2dCQUl6VzlXVllZLzNheEdpY2pTMEFURVpPOGJLV0xjNmtBZysyNHpNZUFBRHdsTDY1dTdzelVBQmdsSzR2NXlHRWZkT2lZUmM1UmF1bEFaaWtyaTlISVlSZnBBTVA5bDFPOGRyNEFBQ0F4N0xSQWdBWXBYNnhyMlJCNjQ0a0RNQlU1UlJQUXdqZkQ4OWNDUWtlNU5UWUFBQ0FwNkJvQVFEOG8vcWU4Y3FrYU53N054d0JtTHI2bE1qd25OdUZzT0RlOXJ1K0hCb2JBQUR3V0lvV0FNQVl4L1ZkWTJqVnJUSVJBSE9SVS94WW43cDZKelM0dDVOYUpBY0FBSGd3UlFzQTRLdTZ2dXpXb2dXMGJEVWNXa2tZZ0RuSktRNmYwZDdVd2lBd3pvNS8zd0FBQUkrbGFBRUEvSlBobHYrV0tkR3dtNXppaVlBQm1LT2M0bWtJWWRodWNTTkFHTzI0RnNvQkFBQWVSTkVDQVBpaXJpL0RsL2F2VFlqR0hRa1lnRG5MS1Y2R0VQWkNDQmVDaEZHMlBCc0hBQUE4aHFJRkFQQTFidm5UdWc4NXhYTXBBekIzd3hOWU9jV2hKUHRPbURESzYxb3NCd0FBdURkRkN3RGdzN3ErRExmOFg1Z09qZk0rTndCTnlTa09mOXZlU0JWR3NkVUNBQUI0RUVVTEFPQy9kSDNadHMyQ0JYaWJVN3dXTkFDdHlTbWVoaEMrRHlIY0NoZSthcjhXekFFQUFPNUYwUUlBK0p6aittNHh0T3BXbVFpQWx1VVVMME1JZXlHRUswSERWNjFxMFJ3QUFHQTBSUXNBNEMrNnZ1eUdFSDQyRlJwM1BMeGxMMlFBV2xZM054MkVFRDRJR3I1b3gzTnlBQURBZlgxemQzZG5hQURBbjdxK3ZBOGh2RElSR25hVlU5d1RNQUJMMHZWbDJPVDBvOURoczRadFo3dUt1QUFBd0ZnMldnQUFmK3I2Y3FCa3dRSzRzUWpBNHVRVWg3OS9ieVFQbjdYbFdUa0FBT0ErRkMwQWdFK2RtZ2FOTzhzcG5nc1pnQ1hLS1E2ZjlWN1cyL3ZBWDcydXp5Z0NBQUQ4STBVTEFPQjNYVitPNi92RTBLcmhVR2tsWFFDV3JCWU9ENVF0NExOOFZnUUFBRVpSdEFBQWhwTEZ0aThWV1lDVG5PSzFvQUZZdXB6aVpRaGh1TGwvdGZSWndOKzhyczhwQWdBQWZKV2lCUUFRYXNsaXl5Um8yQlpOSHRVQUFCYlVTVVJCVkkxM3R3SGdQM0tLSCt0bWl3dGpnYjlRUUFjQUFQN1JOM2QzZDZZRUFBdlc5V1V2aFBCdlB3TTA3b2VjNG5zaEE4Qi82L3B5T3R6a054cjQwOHY2ekE0QUFNQm4yV2dCQUxqbFQrc3VsQ3dBNE10eWlrY2hoSGRHQkgreTFRSUFBUGdxUlFzQVdMQ3VMNGNoaEgwL0F6VHVXTUFBOEhVNXhlSHY1UnRqZ3QvdDEzOHJBUUFBZkphaUJRQXNtMjBXdE80c3AzZ3BaUUQ0WnpuRlUyVUwrSk4vS3dFQUFGK2thQUVBQzlYMVpWaUh1eU4vR25acm13VUEzRTh0VzN4Zi80N0NrdTEwZlRueUV3QUFBSHlPb2dVQUxGRFhsMjBIMEN6QVNVN3hvNkFCNEg3cU5xZ0RaUXNJS3lNQUFBQStSOUVDQUpacFdJTzdKWHNhZHBOVDlNVTRBRHlRc2dYOHpsWUxBQURnc3hRdEFHQmh1cjRNWDVpL2xqdU44NFU0QUR4U0xWdnNoUkN1ekpJRlU5NEZBQUQraTZJRkFDeVBMd3BwM1VWTzhWektBUEI0T2NYcnV0bEMyWUtsc3RVQ0FBRDRMNG9XQUxBZzlRdkNmWm5UT0YrRUE4QVR5aWwrVkxaZzRaVFZBUUNBdjFDMEFJQ0Y2UHF5N1F0Q0Z1QmR2WGtMQUR3aFpRc1didGhxY2JqMElRQUFBUCtoYUFFQXkzRThmRUVvYnhwMnEwd0VBT3VqYk1IQ0hTOTlBQUFBd0g4b1dnREFBblI5MmZYRklBdXdxZ2RBQU1DYUtGdXdZUHRkWHc3OEFBQUFBRUhSQWdBVzR5U0VzQ1Z1R25hVlV6d1JNQUNzM3lkbGl3L0d6Y0xZbmdZQUFQeE8wUUlBR2xkdlhiMlNNNDJ6c1FVQW50RlF0c2dwSG9ZUXpzeWRCYkhWQWdBQStKMmlCUUMwenkxL1d2Y2hwM2d1WlFCNGZqbkZJMlVMRnVaSTRBQUFnS0lGQURTczY4dndKZUFMR2RNNDJ5d0FZSU5xMmVKQ0JpekU2NjR2dThJR0FJQmxVN1FBZ0VaMWZkbTJ6WUlGZUp0VHZCWTBBR3pjOEl6SWxSaFlpSldnQVFCZzJSUXRBS0Jkd3kzL0xmblNzQnRsSWdDWWhweml4eERDZ2JJRkMvRzZGdHNCQUlDRlVyUUFnQWJWVmJZL3k1YkdyZXFoRGdBd0Fjb1dMSXpuNndBQVlNRVVMUUNnVGFkeXBYRVhPVVUvNXdBd01jb1dMSWlpQlFBQUxKaWlCUUEwcHV2TDhNWDJ2bHhwbkhleEFXQ2lsQzFZaUsydUwwZkNCZ0NBWlZLMEFJRDJ1T1ZQNjg1eWl1ZFNCb0RwK3FSc2NTTW1HbWFyQlFBQUxKU2lCUUEwcE92TDhFWGZqa3hwMksxdEZnQXdEN1ZzY1ZqL2ZrT0xYdFNOZ2dBQXdNSW9XZ0JBSTdxK2JEdUFaZ0ZPY29yWGdnYUFlY2dwWHRiTkZzb1d0TXJ6SVFBQXNFQ0tGZ0RRanBQaG5XQjUwckNibktJeUVRRE1qTElGalh2ZDlXVlh5QUFBc0N5S0ZnRFFnSzR2ZThNWGZMS2tjZDdBQm9DWnFtVUxmOHRwbGEwV0FBQ3dNSW9XQU5DR0V6blN1SXVjNG5zaEE4Qjg1UlJQUXdodlJFaURsSWdBQUdCaEZDMEFZT2E2dmh5R0VQYmxTT044ZVEwQURhaGxpN2V5cERGYlhWOXN0UUFBZ0FWUnRBQ0ErYlBOZ3RhZDFYWGpBRUFEY29xcjRlKzdMR21Nb2dVQUFDeUlvZ1VBekZqWGwrRkw2aDBaMHJCYjJ5d0FvRDA1eGVGUStrSzBOR1MvNjh1dVFBRUFZQmtVTFFCZ3B1cVhlQTZnYWQwcXAvaFJ5Z0RRcE9FSnZDdlIwaEQvUGdNQWdJVlF0QUNBK1JxMldXekpqNGJkNUJROWpRTUFqYXBseXNPNndRcGE0UGtRQUFCWUNFVUxBSmlocmk4SElZVFhzcU54dnFnR2dNYmxGSzlEQ0FmS0ZqUmlxK3VMejdBQUFMQUFpaFlBTUU4cnVkRzRpNXppdVpBQm9IMDV4VXRQTHRDUVEyRUNBRUQ3RkMwQVlHYnFEYWw5dWRFNE53RUJZRUZ5aXFjaGhMY3lwd0d2dXI3c0NoSUFBTnFtYUFFQU05TDFaZHMyQ3hiZ1hWMGpEZ0FzU0U1eCtKejdRZVkwd0ZZTEFBQm9uS0lGQU16THNGSjVSMlkwN0ZhWkNBQVdiZGhxZGJYMElUQjduc0lCQUlER2ZYTjNkeWRqQUppQnVuNzJOMW5SdURkMWRUZ0FzRkQxYys5bENHSEx6d0F6OW4xTzhWS0FBQURRSmhzdEFHQStUbVJGNDY2VUxBQ0Erb1NZcHhlWU8xc3RBQUNnWVlvV0FEQURYVjhPUWdpdlpFWGpmQmtOQVB3dXAzZ2VRdmpKTkpneFpTRUFBR2lZb2dVQXpJTnRGclR1UXoxUUFRRDRYVTV4K0F6OHdUU1lxYTJ1TDhvV0FBRFFLRVVMQUppNHJpL0RMZjhYY3FKaHQ3WlpBQUJmY0JSQ3VERWNaa3JSQWdBQUd2WE4zZDJkYkFGZ29ycStiSWNRaGplcXQyUkV3OTdtRkZjQ0JnQStwK3ZMWGdqaDNHZGladXIvNVJRL0NnOEFBTnBpb3dVQVROdktGOG8wN3NiVE9BREExK1FVTDIyL1lzWnN0UUFBZ0FZcFdnREFSSFY5MlEwaC9DZ2ZHcmR5d3c4QStDYzV4ZE1Rd3BsQk1VT0tGZ0FBMENCRkN3Q1lybFBaMExpTGVtZ0NBRERHc05YaXlxU1ltVmYxU1VnQUFLQWhpaFlBTUVGZFh3NUNDUHV5b1hFckFRTUFZOVV0V0VjR3hnelphZ0VBQUkxUnRBQ0FhWExMbjlhZDVSVFBwUXdBM0VkTzhUS0U4Sk9oTVRPS0ZnQUEwQmhGQ3dDWW1LNHZ3eTMvSGJuUXNOdTYraHNBNE41eWlpZkRFMlFteDR4NFBnUUFBQnFqYUFFQUUxSy9mSE1BVGV0TzZ1cHZBSUNIT3F6bFRaZ0xXeTBBQUtBaGloWUFNQzNEN2J3dG1kQ3dtNXppU3NBQXdHUFUwdWFSSVRJamloWUFBTkFRUlFzQW1JaXVMM3NoaE5meW9IRTJ0Z0FBVHlLbitENkVjR2Fhek1RclFRRUFRRHNVTFFCZ09rNWtRZU11Nm9FSUFNQlRHVXFjTjZiSkhIUjlzZFVDQUFBYW9XZ0JBQk5RdjNEYmx3V05zODBDQUhoU25oQmhaaFF0QUFDZ0VZb1dBTEJoWFYrMmJiTmdBZDdsRkM4RkRRQTh0WnppK2ZCWncyQ1pBVVVMQUFCb2hLSUZBR3plY010L1J3NDA3RGFFc0JJd0FMQkdLMCtJTUFOYlhWLzJCQVVBQVBPbmFBRUFHOVQxWmRkekNpekFxcTcxQmdCWUMwK0lNQ08yV2dBQVFBTVVMUUJnczRhYmQxc3lvR0UzT1VWUDR3QUFhMWVmRURremFTWk8wUUlBQUJxZ2FBRUFHOUwxNVNDRThOcjhhWnlicFFEQWN6cXV6NWJCVkwyb213MEJBSUFaVTdRQWdNMVptVDJOKzFCdmxnSUFQSXY2aElpbitaaTZBd2tCQU1DOEtWb0F3QVowZlJsdStlK2JQWTF6eUFFQVBMdWM0bWtJNGNMa21UQkZDd0FBbURsRkN3QjRabDFmdGtNSUorWk80OTdtRksrRkRBQnNpT2ZMbUxKRDZRQUF3THdwV2dEQTh4dHUrVytaT3cyN1ZTWUNBRGFwRmo3ZkNvR0oydXI2c2ljY0FBQ1lMMFVMQUhoR1hWOTJRd2cvbXptTk82N3Zvd01BYk5KUS9MeVJBQlBsK1JBQUFKZ3hSUXNBZUY1dStkTzZxL291T2dEQVJ0WGk1N0VVbUNqUGh3QUF3SXdwV2dEQU0rbjZNdHhZZW1YZU5NNWhCZ0F3R1RuRjl5R0VDNGt3UWZ0Q0FRQ0ErVkswQUlEbjQ1WS9yVHZMS1o1TEdRQ1lHRVZRSnFtVzhRRUFnQmxTdEFDQVo5RDFaZmh5ZDhlc2FkaHRDR0VsWUFCZ2FuS0tsME1oVkRCTWtLSUZBQURNbEtJRkFLeFoxNWR0QjlBc3dFbE84VnJRQU1CRXJXb3hGS1pFMFFJQUFHWkswUUlBMW0vNFVuZkxuR25ZelZDMEVEQUFNRlcxRU9yekNsT3pMeEVBQUpnblJRc0FXS091TDdzaGhCL05tTWF0Y29vZmhRd0FUTnlKclJaTVRkY1hXeTBBQUdDR0ZDMEFZTDFPelpmR1hlUVUvWndEQUpOWGk2RzJXakExaWhZQUFEQkRpaFlBc0NaZFh3NnRnbVVCam9VTUFNeEZUbkZWbnoyRHFWQzBBQUNBR1ZLMEFJRDFjVnVPMXAzbEZDK2xEQURNekVwZ1RJaHlQZ0FBekpDaUJRQ3NRZGVYNGN2YkhiT2xZYmUyV1FBQWMxU2ZQYlBWZ3NubytySW5EUUFBbUJkRkN3QjRZbDFmdGgxQXN3QW45WjF6QUlBNXN0V0NLZkY4Q0FBQXpJeWlCUUE4dmVISmtDMXpwV0UzOVgxekFJQlpzdFdDaWJIUkFnQUFaa2JSQWdDZVVGMzUrdHBNYWR5UmdBR0FCaWlPTWhVMldnQUF3TXdvV2dEQTB6b3hUeHAza1ZNOEZ6SUFNSGQxcThXdElKbUFuZm9FSlFBQU1CT0tGZ0R3UkxxK0RMZjg5ODJUeHRsbUFRQzBSRkdhcWZCOENBQUF6SWlpQlFBOGdYcjd5T3BoV3ZjdXAzZ3RaUUNnSVNlMldqQVJuZzhCQUlBWlViUUFnS2R4UEt4N05Vc2FkcXRNQkFDMEpxZjRNWVJ3S2xnbXdFWUxBQUNZRVVVTEFIaWtyaSs3dFdnQkxWdlZnd2dBZ05aNFBvUXBVTFFBQUlBWlViUUFnTWNiYnZsdm1TTU51OG9wT29BQUFKcFVuMFk3a3k0YnRsT2ZwQVFBQUdaQTBRSUFIcUhyeS9DTzdtc3pwSEUydGdBQXJmTjhDRk5ncXdVQUFNeUVvZ1VBUEk1Yi9yVHVRMDd4WE1vQVFNdnE1NTBySWJOaEJ3SUFBSUI1VUxRQWdBZnErbklVUW5oaGZqVE9OZ3NBWUNtVXFObTBYUWtBQU1BOEtGb0F3QVBVdDNOOUVVdnIzdFkzeXdFQWx1QjlDT0ZXMG15UXAwTUFBR0FtRkMwQTRHR0dXLzViWmtmRGJwV0pBSUFseVNsK3JHVUwyQlFiRXdFQVlDWVVMUURnbnJxK0RPdGNmelkzR25kY0R4c0FBSlpFMFpTTjZ2cGlxd1VBQU15QW9nVUEzTitwbWRHNGk1eWluM01BWUhGeWlwY2hoQnZKczBHN2hnOEFBTk9uYUFFQTk5RDE1U0NFc0c5bU5HNGxZQUJnd1d5MVlKTnN0QUFBZ0JsUXRBQ0ErM0hMbjlhZDVSVFBwUXdBTE5oNzRiTkJpaFlBQURBRGloWUFNRkxYbCtNUXdvNTUwYkJiMnl3QWdLWExLVjZIRUQ0c2ZRNXNqS2REQUFCZ0JoUXRBR0NFcmkvYkRxQlpnSk42c0FBQXNIUzJXckFwTDB3ZUFBQ21UOUVDQU1ZWlNoWmJaa1hEYnJ4SERnRHdKMFVMTnFicmk2MFdBQUF3Y1lvV0FQQVB1cjRNYitUK2FFNDA3amluK0ZISUFBQy9QeC95MGZNaGJKQ2lCUUFBVEp5aUJRRDhNN2Y4YWQxRlR0R3RUUUNBdi9MNWlFMDVNSGtBQUpnMlJRc0ErSXF1TDRjaGhIMHpvbkhIQWdZQStDK0tGbXpLdHNrREFNQzBLVm9Bd05mWlprSHJ6bktLbDFJR0FQaXIrbnpJbGJHd0FYdUdEZ0FBMDZab0FRQmYwUFZsRlVMWU1SOGFkbXViQlFEQVY1MGFEeHRnb3dVQUFFemMvd2dJQUw3b09vVHcxbmhvMkdXOXFRa0F3T2U5ZCtnTkFBREEzMzF6ZDNkbktBQUFBQUFBQUFBQUkzZzZCQUFBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQUFBQUFBQUFBQUVaU3RBQUFBQUFBQUFBQUdFblJBZ0FBQUFBQUFBQmdKRVVMQUFBQUFBQUFBSUNSRkMwQStQL3Qyb0VBQUFBQXc2RDdVMTloQU1VUkFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFFSWtXQUFBQUFBQUFBQUNSYUFFQUFBQUFBQUFBRUlrV0FBQUFBQUFBQUFDUmFBRUFBQUFBQUFBQUVJa1dBQUFBQUFBQUFBQ1JhQUVBQUFBQUFBQUFVR3c3RWJ6eERWYU85cE1BQUFBQVNVVk9SSzVDWUlJPSc7CmZ1bmN0aW9uIGFwcGx5QnJhbmRpbmcoKSB7CiAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnLmJyYW5kLWxvZ28nKS5mb3JFYWNoKChpbWcpID0+IHsgaW1nLnNyYyA9IExPR09fREFUQV9VUkk7IH0pOwogIGNvbnN0IGZhdmljb24gPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmF2aWNvbkxpbmsnKTsKICBpZiAoZmF2aWNvbikgZmF2aWNvbi5ocmVmID0gTE9HT19EQVRBX1VSSTsKfQoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIEFwaSDigJQgdGhpbiBmZXRjaCB3cmFwcGVycyBhcm91bmQgdGhlIFJFU1QgQVBJLgogICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KY29uc3QgQXBpID0gKCgpID0+IHsKICBhc3luYyBmdW5jdGlvbiByZXF1ZXN0KHBhdGgsIG9wdGlvbnMpIHsKICAgIGNvbnN0IHJlcyA9IGF3YWl0IGZldGNoKHBhdGgsIG9wdGlvbnMpOwogICAgbGV0IGJvZHk7CiAgICB0cnkgewogICAgICBib2R5ID0gYXdhaXQgcmVzLmpzb24oKTsKICAgIH0gY2F0Y2ggewogICAgICBib2R5ID0gbnVsbDsKICAgIH0KICAgIGlmIChyZXMuc3RhdHVzID09PSA0MDEgJiYgIXBhdGguc3RhcnRzV2l0aCgnL2FwaS9hdXRoLycpKSB7CiAgICAgIHdpbmRvdy5kaXNwYXRjaEV2ZW50KG5ldyBDdXN0b21FdmVudCgnbHJzOnNpZ25lZC1vdXQnKSk7CiAgICB9CiAgICBpZiAoIXJlcy5vaykgewogICAgICBjb25zdCBtZXNzYWdlID0gKGJvZHkgJiYgYm9keS5lcnJvcikgfHwgYFJlcXVlc3QgZmFpbGVkICgke3Jlcy5zdGF0dXN9KWA7CiAgICAgIHRocm93IG5ldyBFcnJvcihtZXNzYWdlKTsKICAgIH0KICAgIHJldHVybiBib2R5OwogIH0KCiAgZnVuY3Rpb24gcXMocGFyYW1zKSB7CiAgICBjb25zdCB1c3AgPSBuZXcgVVJMU2VhcmNoUGFyYW1zKCk7CiAgICBPYmplY3QuZW50cmllcyhwYXJhbXMgfHwge30pLmZvckVhY2goKFtrLCB2XSkgPT4gewogICAgICBpZiAodiAhPT0gdW5kZWZpbmVkICYmIHYgIT09IG51bGwgJiYgdiAhPT0gJycpIHVzcC5zZXQoaywgdik7CiAgICB9KTsKICAgIGNvbnN0IHMgPSB1c3AudG9TdHJpbmcoKTsKICAgIHJldHVybiBzID8gYD8ke3N9YCA6ICcnOwogIH0KCiAgcmV0dXJuIHsKICAgIGF1dGhNZTogKCkgPT4gcmVxdWVzdCgnL2FwaS9hdXRoL21lJyksCiAgICBhdXRoTG9naW46IChjb2RlKSA9PgogICAgICByZXF1ZXN0KCcvYXBpL2F1dGgvbG9naW4nLCB7IG1ldGhvZDogJ1BPU1QnLCBoZWFkZXJzOiB7ICdDb250ZW50LVR5cGUnOiAnYXBwbGljYXRpb24vanNvbicgfSwgYm9keTogSlNPTi5zdHJpbmdpZnkoeyBjb2RlIH0pIH0pLAogICAgYXV0aExvZ291dDogKCkgPT4gcmVxdWVzdCgnL2FwaS9hdXRoL2xvZ291dCcsIHsgbWV0aG9kOiAnUE9TVCcgfSksCgogICAgZmlsdGVyT3B0aW9uczogKCkgPT4gcmVxdWVzdCgnL2FwaS9hbmFseXRpY3MvZmlsdGVyLW9wdGlvbnMnKSwKICAgIGtwaXM6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL2twaXMke3FzKHBhcmFtcyl9YCksCiAgICBwbGF0Zm9ybUJyZWFrZG93bjogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9hbmFseXRpY3MvcGxhdGZvcm0tYnJlYWtkb3duJHtxcyhwYXJhbXMpfWApLAogICAgY2FtcGFpZ25CcmVha2Rvd246IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL2NhbXBhaWduLWJyZWFrZG93biR7cXMocGFyYW1zKX1gKSwKICAgIGNvbnRlbnRUeXBlQnJlYWtkb3duOiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9jb250ZW50LXR5cGUtYnJlYWtkb3duJHtxcyhwYXJhbXMpfWApLAogICAgbWV0cmljT3B0aW9uczogKHBsYXRmb3JtKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9tZXRyaWMtb3B0aW9ucyR7cXMoeyBwbGF0Zm9ybSB9KX1gKSwKICAgIG1ldHJpY1N1bW1hcnk6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL21ldHJpYy1zdW1tYXJ5JHtxcyhwYXJhbXMpfWApLAogICAgdHJlbmQ6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL3RyZW5kJHtxcyhwYXJhbXMpfWApLAogICAgdG9wUG9zdHM6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL3RvcC1wb3N0cyR7cXMocGFyYW1zKX1gKSwKICAgIGNvbXBhcmU6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL2NvbXBhcmUke3FzKHBhcmFtcyl9YCksCiAgICBtb250aGx5OiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9tb250aGx5JHtxcyhwYXJhbXMpfWApLAogICAgcXVhcnRlcmx5OiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9xdWFydGVybHkke3FzKHBhcmFtcyl9YCksCiAgICB5dGQ6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL3l0ZCR7cXMocGFyYW1zKX1gKSwKICAgIHBsYXRmb3JtUmVwb3J0OiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2FuYWx5dGljcy9wbGF0Zm9ybS1yZXBvcnQke3FzKHBhcmFtcyl9YCksCiAgICBwbGF0Zm9ybVN1bW1hcnk6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvYW5hbHl0aWNzL3BsYXRmb3JtLXN1bW1hcnkke3FzKHBhcmFtcyl9YCksCgogICAgcHJldmlld1VwbG9hZDogKGZpbGUpID0+IHsKICAgICAgY29uc3QgZm9ybSA9IG5ldyBGb3JtRGF0YSgpOwogICAgICBmb3JtLmFwcGVuZCgnZmlsZScsIGZpbGUpOwogICAgICByZXR1cm4gcmVxdWVzdCgnL2FwaS91cGxvYWRzL3ByZXZpZXcnLCB7IG1ldGhvZDogJ1BPU1QnLCBib2R5OiBmb3JtIH0pOwogICAgfSwKICAgIGNvbW1pdFVwbG9hZDogKHBheWxvYWQpID0+CiAgICAgIHJlcXVlc3QoJy9hcGkvdXBsb2Fkcy9jb21taXQnLCB7CiAgICAgICAgbWV0aG9kOiAnUE9TVCcsCiAgICAgICAgaGVhZGVyczogeyAnQ29udGVudC1UeXBlJzogJ2FwcGxpY2F0aW9uL2pzb24nIH0sCiAgICAgICAgYm9keTogSlNPTi5zdHJpbmdpZnkocGF5bG9hZCksCiAgICAgIH0pLAogICAgdXBsb2FkSGlzdG9yeTogKCkgPT4gcmVxdWVzdCgnL2FwaS91cGxvYWRzL2hpc3RvcnknKSwKICAgIHVwbG9hZEVycm9yczogKGlkKSA9PiByZXF1ZXN0KGAvYXBpL3VwbG9hZHMvJHtpZH0vZXJyb3JzYCksCiAgICB1cGxvYWRSYXdSb3dzOiAoaWQpID0+IHJlcXVlc3QoYC9hcGkvdXBsb2Fkcy8ke2lkfS9yYXctcm93c2ApLAoKICAgIGxpc3RSZWNvcmRzOiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL3JlY29yZHMke3FzKHBhcmFtcyl9YCksCiAgICByZWNvcmRzVGFibGU6IChwYXJhbXMpID0+IHJlcXVlc3QoYC9hcGkvcmVjb3Jkcy90YWJsZSR7cXMocGFyYW1zKX1gKSwKICAgIGdldFJlY29yZDogKGlkKSA9PiByZXF1ZXN0KGAvYXBpL3JlY29yZHMvJHtpZH1gKSwKICAgIHVwZGF0ZVJlY29yZDogKGlkLCB2YWx1ZXMpID0+CiAgICAgIHJlcXVlc3QoYC9hcGkvcmVjb3Jkcy8ke2lkfWAsIHsKICAgICAgICBtZXRob2Q6ICdQVVQnLAogICAgICAgIGhlYWRlcnM6IHsgJ0NvbnRlbnQtVHlwZSc6ICdhcHBsaWNhdGlvbi9qc29uJyB9LAogICAgICAgIGJvZHk6IEpTT04uc3RyaW5naWZ5KHsgdmFsdWVzIH0pLAogICAgICB9KSwKICAgIGRlbGV0ZVJlY29yZFBvc3Q6IChwb3N0SWQpID0+IHJlcXVlc3QoYC9hcGkvcmVjb3Jkcy9wb3N0LyR7cG9zdElkfWAsIHsgbWV0aG9kOiAnREVMRVRFJyB9KSwKICAgIGRlbGV0ZVJlY29yZFBsYXRmb3JtOiAocG9zdElkLCBwbGF0Zm9ybSkgPT4KICAgICAgcmVxdWVzdChgL2FwaS9yZWNvcmRzL3Bvc3QvJHtwb3N0SWR9L3BsYXRmb3JtLyR7cGxhdGZvcm19YCwgeyBtZXRob2Q6ICdERUxFVEUnIH0pLAoKICAgIHJlc3RvcmVCYWNrdXA6IChmb3JtKSA9PiByZXF1ZXN0KCcvYXBpL2JhY2t1cC9yZXN0b3JlJywgeyBtZXRob2Q6ICdQT1NUJywgYm9keTogZm9ybSB9KSwKCiAgICBsaXN0Rm9sbG93ZXJzOiAocGFyYW1zKSA9PiByZXF1ZXN0KGAvYXBpL2ZvbGxvd2VycyR7cXMocGFyYW1zKX1gKSwKICAgIGZvbGxvd2Vyc0dyb3d0aDogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9mb2xsb3dlcnMvZ3Jvd3RoJHtxcyhwYXJhbXMpfWApLAogICAgZm9sbG93ZXJzS3BpczogKHBhcmFtcykgPT4gcmVxdWVzdChgL2FwaS9mb2xsb3dlcnMva3BpcyR7cXMocGFyYW1zKX1gKSwKICAgIHNhdmVGb2xsb3dlcnM6IChwYXlsb2FkKSA9PgogICAgICByZXF1ZXN0KCcvYXBpL2ZvbGxvd2VycycsIHsgbWV0aG9kOiAnUE9TVCcsIGhlYWRlcnM6IHsgJ0NvbnRlbnQtVHlwZSc6ICdhcHBsaWNhdGlvbi9qc29uJyB9LCBib2R5OiBKU09OLnN0cmluZ2lmeShwYXlsb2FkKSB9KSwKICAgIHVwZGF0ZUZvbGxvd2VyczogKGlkLCBwYXlsb2FkKSA9PgogICAgICByZXF1ZXN0KGAvYXBpL2ZvbGxvd2Vycy8ke2lkfWAsIHsgbWV0aG9kOiAnUFVUJywgaGVhZGVyczogeyAnQ29udGVudC1UeXBlJzogJ2FwcGxpY2F0aW9uL2pzb24nIH0sIGJvZHk6IEpTT04uc3RyaW5naWZ5KHBheWxvYWQpIH0pLAogICAgZGVsZXRlRm9sbG93ZXJzOiAoaWQpID0+IHJlcXVlc3QoYC9hcGkvZm9sbG93ZXJzLyR7aWR9YCwgeyBtZXRob2Q6ICdERUxFVEUnIH0pLAogIH07Cn0pKCk7CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgU3RhdGUgLyBGb3JtYXQgLyBUb2FzdCDigJQgc2hhcmVkIGFwcCBzdGF0ZSArIHNtYWxsIHV0aWxpdGllcy4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmNvbnN0IFN0YXRlID0gKCgpID0+IHsKICBjb25zdCB0b2RheSA9IG5ldyBEYXRlKCk7CiAgY29uc3QgaXNvID0gKGQpID0+IGQudG9JU09TdHJpbmcoKS5zbGljZSgwLCAxMCk7CiAgY29uc3QgdGhpcnR5RGF5c0FnbyA9IG5ldyBEYXRlKHRvZGF5KTsKICB0aGlydHlEYXlzQWdvLnNldERhdGUodGhpcnR5RGF5c0Fnby5nZXREYXRlKCkgLSAyOSk7CgogIGNvbnN0IGZpbHRlcnMgPSB7CiAgICBkYXRlRnJvbTogaXNvKHRoaXJ0eURheXNBZ28pLAogICAgZGF0ZVRvOiBpc28odG9kYXkpLAogICAgcGxhdGZvcm06ICdhbGwnLAogICAgY2FtcGFpZ25UeXBlOiAnYWxsJywKICAgIGNvbnRlbnRUeXBlOiAnYWxsJywKICB9OwoKICBjb25zdCBsaXN0ZW5lcnMgPSBbXTsKCiAgcmV0dXJuIHsKICAgIGdldEZpbHRlcnM6ICgpID0+ICh7IC4uLmZpbHRlcnMgfSksCiAgICBzZXRGaWx0ZXJzKHBhcnRpYWwpIHsKICAgICAgT2JqZWN0LmFzc2lnbihmaWx0ZXJzLCBwYXJ0aWFsKTsKICAgICAgbGlzdGVuZXJzLmZvckVhY2goKGZuKSA9PiBmbih0aGlzLmdldEZpbHRlcnMoKSkpOwogICAgfSwKICAgIG9uQ2hhbmdlKGZuKSB7CiAgICAgIGxpc3RlbmVycy5wdXNoKGZuKTsKICAgIH0sCiAgfTsKfSkoKTsKCmNvbnN0IEZvcm1hdCA9IHsKICBudW1iZXIobikgewogICAgaWYgKG4gPT09IG51bGwgfHwgbiA9PT0gdW5kZWZpbmVkKSByZXR1cm4gJ+KAlCc7CiAgICByZXR1cm4gTWF0aC5yb3VuZChuKS50b0xvY2FsZVN0cmluZygnZW4tVVMnKTsKICB9LAogIGNvbXBhY3QobikgewogICAgaWYgKG4gPT09IG51bGwgfHwgbiA9PT0gdW5kZWZpbmVkKSByZXR1cm4gJ+KAlCc7CiAgICBjb25zdCBhYnMgPSBNYXRoLmFicyhuKTsKICAgIGlmIChhYnMgPj0gMV8wMDBfMDAwKSByZXR1cm4gYCR7KG4gLyAxXzAwMF8wMDApLnRvRml4ZWQoMSkucmVwbGFjZSgvXC4wJC8sICcnKX1NYDsKICAgIGlmIChhYnMgPj0gMV8wMDApIHJldHVybiBgJHsobiAvIDFfMDAwKS50b0ZpeGVkKDEpLnJlcGxhY2UoL1wuMCQvLCAnJyl9S2A7CiAgICByZXR1cm4gYCR7TWF0aC5yb3VuZChuKX1gOwogIH0sCiAgLyoqIERhc2hib2FyZC13aWRlICJwcm9mZXNzaW9uYWwiIG51bWJlciBmb3JtYXQ6IHBsYWluIHVuZGVyIDEsMDAwOyBjb21tYS1ncm91cGVkCiAgICAgIHVwIHRvIDEwLDAwMDsgYWJicmV2aWF0ZWQgKEsvTSkgYmV5b25kIHRoYXQg4oCUIGUuZy4gODUwLCAxLDI1MCwgMTIuNUssIDE1NkssIDEuMjVNLiAqLwogIHNtYXJ0KG4pIHsKICAgIGlmIChuID09PSBudWxsIHx8IG4gPT09IHVuZGVmaW5lZCkgcmV0dXJuICfigJQnOwogICAgY29uc3QgYWJzID0gTWF0aC5hYnMobik7CiAgICBpZiAoYWJzIDwgMTAwMCkgcmV0dXJuIGAke01hdGgucm91bmQobil9YDsKICAgIGlmIChhYnMgPCAxMDAwMCkgcmV0dXJuIE1hdGgucm91bmQobikudG9Mb2NhbGVTdHJpbmcoJ2VuLVVTJyk7CiAgICBpZiAoYWJzIDwgMV8wMDBfMDAwKSByZXR1cm4gYCR7KG4gLyAxMDAwKS50b0ZpeGVkKDEpLnJlcGxhY2UoL1wuMCQvLCAnJyl9S2A7CiAgICByZXR1cm4gYCR7KG4gLyAxXzAwMF8wMDApLnRvRml4ZWQoMikucmVwbGFjZSgvXC4/MCskLywgJycpfU1gOwogIH0sCiAgcGVyY2VudChuKSB7CiAgICBpZiAobiA9PT0gbnVsbCB8fCBuID09PSB1bmRlZmluZWQpIHJldHVybiAn4oCUJzsKICAgIHJldHVybiBgJHtOdW1iZXIobikudG9GaXhlZCgxKS5yZXBsYWNlKC9cLjAkLywgJycpfSVgOwogIH0sCiAgcGN0KG4pIHsKICAgIGlmIChuID09PSBudWxsIHx8IG4gPT09IHVuZGVmaW5lZCkgcmV0dXJuICfigJQnOwogICAgY29uc3Qgc2lnbiA9IG4gPiAwID8gJysnIDogJyc7CiAgICByZXR1cm4gYCR7c2lnbn0ke24udG9GaXhlZCgxKX0lYDsKICB9LAogIGRhdGUoaXNvXykgewogICAgaWYgKCFpc29fKSByZXR1cm4gJ+KAlCc7CiAgICBjb25zdCBbeSwgbSwgZF0gPSBpc29fLnNwbGl0KCctJykubWFwKE51bWJlcik7CiAgICByZXR1cm4gbmV3IERhdGUoeSwgbSAtIDEsIGQpLnRvTG9jYWxlRGF0ZVN0cmluZygnZW4tVVMnLCB7IG1vbnRoOiAnc2hvcnQnLCBkYXk6ICdudW1lcmljJywgeWVhcjogJ251bWVyaWMnIH0pOwogIH0sCiAgZHVyYXRpb24oc2Vjb25kcykgewogICAgaWYgKHNlY29uZHMgPT09IG51bGwgfHwgc2Vjb25kcyA9PT0gdW5kZWZpbmVkKSByZXR1cm4gJ+KAlCc7CiAgICBjb25zdCBzID0gTWF0aC5yb3VuZChzZWNvbmRzKTsKICAgIGlmIChzIDwgNjApIHJldHVybiBgJHtzfXNgOwogICAgaWYgKHMgPCAzNjAwKSByZXR1cm4gYCR7TWF0aC5mbG9vcihzIC8gNjApfW0gJHtzICUgNjB9c2A7CiAgICBjb25zdCBoID0gTWF0aC5mbG9vcihzIC8gMzYwMCk7CiAgICBjb25zdCBtID0gTWF0aC5yb3VuZCgocyAlIDM2MDApIC8gNjApOwogICAgcmV0dXJuIGAke2h9aCAke219bWA7CiAgfSwKICBkZWx0YUNsYXNzKG4pIHsKICAgIGlmIChuID09PSBudWxsIHx8IG4gPT09IHVuZGVmaW5lZCkgcmV0dXJuICdmbGF0JzsKICAgIGlmIChuID4gMC41KSByZXR1cm4gJ3VwJzsKICAgIGlmIChuIDwgLTAuNSkgcmV0dXJuICdkb3duJzsKICAgIHJldHVybiAnZmxhdCc7CiAgfSwKfTsKCmNvbnN0IFRvYXN0ID0gewogIHNob3cobWVzc2FnZSwgdHlwZSA9ICdzdWNjZXNzJykgewogICAgY29uc3Qgcm9vdCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd0b2FzdFJvb3QnKTsKICAgIGNvbnN0IGVsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBlbC5jbGFzc05hbWUgPSBgdG9hc3QgJHt0eXBlfWA7CiAgICBlbC50ZXh0Q29udGVudCA9IG1lc3NhZ2U7CiAgICByb290LmFwcGVuZENoaWxkKGVsKTsKICAgIHNldFRpbWVvdXQoKCkgPT4gZWwucmVtb3ZlKCksIDUwMDApOwogIH0sCn07CgovKiogU2FmZWx5IGJ1aWxkcyBET00gdGV4dCBub2RlcyBmb3IgdW50cnVzdGVkIHN0cmluZ3MgKGNhcHRpb25zLCBmaWxlbmFtZXMsIHBsYXRmb3JtIGxhYmVscyBmcm9tIGRhdGEpLiAqLwpmdW5jdGlvbiB0ZXh0RWwodGFnLCB0ZXh0LCBjbGFzc05hbWUpIHsKICBjb25zdCBlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQodGFnKTsKICBpZiAoY2xhc3NOYW1lKSBlbC5jbGFzc05hbWUgPSBjbGFzc05hbWU7CiAgZWwuYXBwZW5kQ2hpbGQoZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUodGV4dCA/PyAnJykpOwogIHJldHVybiBlbDsKfQoKLyoqIEEgcHJlbWl1bSBlbXB0eSBzdGF0ZTogaWNvbiArIGV4cGxhbmF0aW9uICsgb3B0aW9uYWwgYWN0aW9uLCBpbnN0ZWFkIG9mIGEgYmxhbmsgYXJlYS4KICAgIEljb25zIHJlbmRlciB2aWEgdGhlIHBhZ2Utd2lkZSBNdXRhdGlvbk9ic2VydmVyIHRoYXQgY2FsbHMgbHVjaWRlLmNyZWF0ZUljb25zKCkgKHNlZSBib290c3RyYXApLiAqLwpmdW5jdGlvbiBlbXB0eVN0YXRlKHsgaWNvbiA9ICdpbmJveCcsIHRpdGxlLCBtZXNzYWdlLCBhY3Rpb25MYWJlbCwgb25BY3Rpb24gfSkgewogIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICB3cmFwLmNsYXNzTmFtZSA9ICdlbXB0eS1zdGF0ZSc7CiAgY29uc3QgaWNvbldyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICBpY29uV3JhcC5jbGFzc05hbWUgPSAnZW1wdHktaWNvbic7CiAgaWNvbldyYXAuaW5uZXJIVE1MID0gYDxpIGRhdGEtbHVjaWRlPSIke2ljb259IiBzdHlsZT0id2lkdGg6MjJweDtoZWlnaHQ6MjJweDsiPjwvaT5gOwogIHdyYXAuYXBwZW5kQ2hpbGQoaWNvbldyYXApOwogIGlmICh0aXRsZSkgd3JhcC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIHRpdGxlLCAnZW1wdHktdGl0bGUnKSk7CiAgaWYgKG1lc3NhZ2UpIHdyYXAuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBtZXNzYWdlLCAnZW1wdHktbWVzc2FnZScpKTsKICBpZiAoYWN0aW9uTGFiZWwgJiYgb25BY3Rpb24pIHsKICAgIGNvbnN0IGJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgYnRuLmNsYXNzTmFtZSA9ICdidG4gcHJpbWFyeSc7CiAgICBidG4udGV4dENvbnRlbnQgPSBhY3Rpb25MYWJlbDsKICAgIGJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIG9uQWN0aW9uKTsKICAgIHdyYXAuYXBwZW5kQ2hpbGQoYnRuKTsKICB9CiAgcmV0dXJuIHdyYXA7Cn0KCi8qKiBBIDxidXR0b24+IHdpdGggYSBzbWFsbCBsZWFkaW5nIEx1Y2lkZSBpY29uIGJlZm9yZSBpdHMgbGFiZWwgKGxhYmVsIGlzIGFsd2F5cyBhIHN0YXRpYywgZGV2ZWxvcGVyLXN1cHBsaWVkIHN0cmluZyBhdCBjYWxsIHNpdGVzLCBuZXZlciB1c2VyIGRhdGEg4oCUIGluc2VydGVkIHZpYSBjcmVhdGVUZXh0Tm9kZSByZWdhcmRsZXNzKS4gKi8KZnVuY3Rpb24gaWNvbkJ0bihjbGFzc05hbWUsIGljb25OYW1lLCBsYWJlbCkgewogIGNvbnN0IGJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogIGJ0bi5jbGFzc05hbWUgPSBjbGFzc05hbWU7CiAgY29uc3QgaWNvbiA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2knKTsKICBpY29uLnNldEF0dHJpYnV0ZSgnZGF0YS1sdWNpZGUnLCBpY29uTmFtZSk7CiAgaWNvbi5zdHlsZS53aWR0aCA9ICcxM3B4JzsKICBpY29uLnN0eWxlLmhlaWdodCA9ICcxM3B4JzsKICBidG4uYXBwZW5kQ2hpbGQoaWNvbik7CiAgYnRuLmFwcGVuZENoaWxkKGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKGAgJHtsYWJlbH1gKSk7CiAgcmV0dXJuIGJ0bjsKfQoKLyoqIFNoaW1tZXJpbmcgcGxhY2Vob2xkZXJzIHNob3duIHRoZSBpbnN0YW50IGEgc2VjdGlvbiBzdGFydHMgbG9hZGluZywgc3dhcHBlZCBmb3IgcmVhbAogICAgY29udGVudCAob3IgYW4gZW1wdHkgc3RhdGUpIG9uY2UgdGhlIGZldGNoIHJlc29sdmVzIOKAlCBubyBibGFuayBhcmVhcyB3aGlsZSB3YWl0aW5nLiAqLwpmdW5jdGlvbiBza2VsZXRvblN0YXRHcmlkKGNvdW50ID0gNikgewogIGNvbnN0IGdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICBncmlkLmNsYXNzTmFtZSA9ICdza2VsZXRvbi1zdGF0LWdyaWQnOwogIGZvciAobGV0IGkgPSAwOyBpIDwgY291bnQ7IGkgKz0gMSkgewogICAgY29uc3QgdGlsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGlsZS5jbGFzc05hbWUgPSAnc2tlbGV0b24gc2tlbGV0b24tdGlsZSc7CiAgICBncmlkLmFwcGVuZENoaWxkKHRpbGUpOwogIH0KICByZXR1cm4gZ3JpZDsKfQpmdW5jdGlvbiBza2VsZXRvbkNoYXJ0KCkgewogIGNvbnN0IGRpdiA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogIGRpdi5jbGFzc05hbWUgPSAnc2tlbGV0b24gc2tlbGV0b24tY2hhcnQnOwogIHJldHVybiBkaXY7Cn0KZnVuY3Rpb24gc2tlbGV0b25Sb3dzKGNvdW50ID0gNikgewogIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICBmb3IgKGxldCBpID0gMDsgaSA8IGNvdW50OyBpICs9IDEpIHsKICAgIGNvbnN0IHJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcm93LmNsYXNzTmFtZSA9ICdza2VsZXRvbiBza2VsZXRvbi1yb3cnOwogICAgd3JhcC5hcHBlbmRDaGlsZChyb3cpOwogIH0KICByZXR1cm4gd3JhcDsKfQoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIFNoYXJlZCBhbmltYXRpb24gcHJpbWl0aXZlcyDigJQgYSBjb3VudC11cCBmb3IgS1BJIG51bWJlcnMgYW5kIGEKICAgQ1NTIHdpZHRoLXRyYW5zaXRpb24gYmFyLCBib3RoIHJldXNlZCBhY3Jvc3MgdGhlIERhc2hib2FyZCBhbmQKICAgQ29tcGFyaXNvbnMgcGFnZXMuIEJvdGggcmVzcGVjdCBwcmVmZXJzLXJlZHVjZWQtbW90aW9uIChndWFyZGVkCiAgIGluIENTUywgc2VlIC5iYXItZmlsbCAvIHRoZSBhbmltYXRlQ291bnQgZHVyYXRpb24gY2hlY2sgYmVsb3cpLgogICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KY29uc3QgUFJFRkVSU19SRURVQ0VEX01PVElPTiA9IHdpbmRvdy5tYXRjaE1lZGlhICYmIHdpbmRvdy5tYXRjaE1lZGlhKCcocHJlZmVycy1yZWR1Y2VkLW1vdGlvbjogcmVkdWNlKScpLm1hdGNoZXM7CgovKiogQW5pbWF0ZXMgYSBudW1iZXIgZnJvbSBgZnJvbWAgdG8gYHRvYCBpbnNpZGUgYGVsYCBvdmVyIGBkdXJhdGlvbmBtcywgZm9ybWF0dGluZyBlYWNoIGZyYW1lIHdpdGggYGZvcm1hdGAgKGRlZmF1bHRzIHRvIGEgcGxhaW4gcm91bmRlZCBpbnRlZ2VyKS4gU2tpcHMgc3RyYWlnaHQgdG8gdGhlIGZpbmFsIHZhbHVlIHVuZGVyIHByZWZlcnMtcmVkdWNlZC1tb3Rpb24uICovCmZ1bmN0aW9uIGFuaW1hdGVDb3VudChlbCwgZnJvbSwgdG8sIGR1cmF0aW9uID0gOTAwLCBmb3JtYXQpIHsKICBpZiAoIWVsKSByZXR1cm47CiAgY29uc3QgZm10ID0gZm9ybWF0IHx8ICgodikgPT4gTWF0aC5yb3VuZCh2KS50b0xvY2FsZVN0cmluZygnZW4tVVMnKSk7CiAgaWYgKFBSRUZFUlNfUkVEVUNFRF9NT1RJT04gfHwgZnJvbSA9PT0gdG8gfHwgIU51bWJlci5pc0Zpbml0ZShmcm9tKSB8fCAhTnVtYmVyLmlzRmluaXRlKHRvKSkgewogICAgZWwudGV4dENvbnRlbnQgPSBmbXQodG8pOwogICAgcmV0dXJuOwogIH0KICBjb25zdCBzdGFydCA9IHBlcmZvcm1hbmNlLm5vdygpOwogIGZ1bmN0aW9uIHRpY2sobm93KSB7CiAgICBjb25zdCBlbGFwc2VkID0gbm93IC0gc3RhcnQ7CiAgICBjb25zdCBwcm9ncmVzcyA9IE1hdGgubWluKDEsIGVsYXBzZWQgLyBkdXJhdGlvbik7CiAgICBjb25zdCBlYXNlZCA9IDEgLSBNYXRoLnBvdygxIC0gcHJvZ3Jlc3MsIDMpOyAvLyBlYXNlT3V0Q3ViaWMKICAgIGVsLnRleHRDb250ZW50ID0gZm10KGZyb20gKyAodG8gLSBmcm9tKSAqIGVhc2VkKTsKICAgIGlmIChwcm9ncmVzcyA8IDEpIHJlcXVlc3RBbmltYXRpb25GcmFtZSh0aWNrKTsKICB9CiAgcmVxdWVzdEFuaW1hdGlvbkZyYW1lKHRpY2spOwp9CgovKiogQSBsYWJlbGVkIGhvcml6b250YWwgYmFyIHRoYXQgYW5pbWF0ZXMgaXRzIHdpZHRoIGluIG9uIGluc2VydGlvbiDigJQgdXNlZCBmb3IgdGhlIENvbXBhcmlzb25zIHBhZ2UncyBwYWlyZWQgUmFuZ2UgQS9CIGJhcnMuIGB2YWx1ZWAvYG1heGAgZHJpdmUgdGhlIGZpbGwgcGVyY2VudGFnZTsgYGNvbG9yVmFyYCBpcyBhIENTUyBjdXN0b20gcHJvcGVydHkgbmFtZSAoZS5nLiAnLS1zZXJpZXMtMScpLiAqLwpmdW5jdGlvbiBidWlsZEJhcih7IGxhYmVsLCB2YWx1ZSwgbWF4LCBjb2xvclZhciwgZm9ybWF0VmFsdWUgfSkgewogIGNvbnN0IHJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogIHJvdy5jbGFzc05hbWUgPSAnYmFyLXJvdyc7CiAgY29uc3QgbGFiZWxFbCA9IHRleHRFbCgnZGl2JywgbGFiZWwsICdiYXItbGFiZWwnKTsKICBjb25zdCB0cmFjayA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogIHRyYWNrLmNsYXNzTmFtZSA9ICdiYXItdHJhY2snOwogIGNvbnN0IGZpbGwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICBmaWxsLmNsYXNzTmFtZSA9ICdiYXItZmlsbCc7CiAgZmlsbC5zdHlsZS5iYWNrZ3JvdW5kID0gY29sb3JWYXIgPyBgdmFyKCR7Y29sb3JWYXJ9KWAgOiAndmFyKC0tc2VyaWVzLTEpJzsKICB0cmFjay5hcHBlbmRDaGlsZChmaWxsKTsKICBjb25zdCB2YWx1ZUVsID0gdGV4dEVsKCdkaXYnLCBmb3JtYXRWYWx1ZSA/IGZvcm1hdFZhbHVlKHZhbHVlKSA6IFN0cmluZyh2YWx1ZSksICdiYXItdmFsdWUnKTsKICByb3cuYXBwZW5kKGxhYmVsRWwsIHRyYWNrLCB2YWx1ZUVsKTsKICBjb25zdCBwY3QgPSBtYXggPiAwID8gTWF0aC5taW4oMTAwLCBNYXRoLnJvdW5kKCh2YWx1ZSAvIG1heCkgKiAxMDAwKSAvIDEwKSA6IDA7CiAgcmVxdWVzdEFuaW1hdGlvbkZyYW1lKCgpID0+IHsgZmlsbC5zdHlsZS53aWR0aCA9IGAke3BjdH0lYDsgfSk7CiAgcmV0dXJuIHJvdzsKfQoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIENoYXJ0cyDigJQgQ2hhcnQuanMgYnVpbGRlcnMgKHZhbGlkYXRlZCBjYXRlZ29yaWNhbCBwYWxldHRlLAogICBoYWlybGluZSByZWNlc3NpdmUgZ3JpZGxpbmVzLCBzaW5nbGUgYXhpcywgbGVnZW5kIGFsd2F5cwogICBwcmVzZW50IGZvciAyKyBzZXJpZXMsIGluZGV4LW1vZGUgdG9vbHRpcHMpLgogICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KaWYgKHdpbmRvdy5DaGFydERhdGFMYWJlbHMpIENoYXJ0LnJlZ2lzdGVyKHdpbmRvdy5DaGFydERhdGFMYWJlbHMpOwoKY29uc3QgQ2hhcnRzID0gKCgpID0+IHsKICBjb25zdCByZWdpc3RyeSA9IG5ldyBNYXAoKTsgLy8gY2FudmFzSWQgLT4gQ2hhcnQgaW5zdGFuY2UsIHNvIHJlLXJlbmRlcnMgZGVzdHJveSB0aGUgb2xkIG9uZSBmaXJzdAogIGNvbnN0IE1BWF9MQUJFTEVEX0lURU1TID0gMjA7IC8vIGJleW9uZCB0aGlzLCBwZXItaXRlbSB2YWx1ZSBsYWJlbHMgd291bGQgb3ZlcmxhcCDigJQgcmVseSBvbiB0b29sdGlwcyBpbnN0ZWFkCgogIGZ1bmN0aW9uIGNzc1ZhcihuYW1lKSB7CiAgICByZXR1cm4gZ2V0Q29tcHV0ZWRTdHlsZShkb2N1bWVudC5kb2N1bWVudEVsZW1lbnQpLmdldFByb3BlcnR5VmFsdWUobmFtZSkudHJpbSgpOwogIH0KCiAgY29uc3QgU0VSSUVTX1ZBUlMgPSBbJy0tc2VyaWVzLTEnLCAnLS1zZXJpZXMtMicsICctLXNlcmllcy0zJywgJy0tc2VyaWVzLTQnLCAnLS1zZXJpZXMtNScsICctLXNlcmllcy02JywgJy0tc2VyaWVzLTcnLCAnLS1zZXJpZXMtOCddOwogIGZ1bmN0aW9uIHNlcmllc0NvbG9yKGluZGV4KSB7CiAgICByZXR1cm4gY3NzVmFyKFNFUklFU19WQVJTW2luZGV4ICUgU0VSSUVTX1ZBUlMubGVuZ3RoXSk7CiAgfQoKICBmdW5jdGlvbiBiYXNlR3JpZCgpIHsKICAgIHJldHVybiB7CiAgICAgIGNvbG9yOiBjc3NWYXIoJy0tZ3JpZGxpbmUnKSwKICAgICAgZHJhd1RpY2tzOiBmYWxzZSwKICAgIH07CiAgfQogIGZ1bmN0aW9uIGJhc2VUaWNrcygpIHsKICAgIHJldHVybiB7IGNvbG9yOiBjc3NWYXIoJy0tdGV4dC1tdXRlZCcpLCBmb250OiB7IHNpemU6IDExIH0gfTsKICB9CiAgZnVuY3Rpb24gYmFzZVRvb2x0aXAoKSB7CiAgICByZXR1cm4gewogICAgICBiYWNrZ3JvdW5kQ29sb3I6IGNzc1ZhcignLS1zdXJmYWNlLTEnKSwKICAgICAgdGl0bGVDb2xvcjogY3NzVmFyKCctLXRleHQtcHJpbWFyeScpLAogICAgICBib2R5Q29sb3I6IGNzc1ZhcignLS10ZXh0LXNlY29uZGFyeScpLAogICAgICBib3JkZXJDb2xvcjogY3NzVmFyKCctLWJvcmRlcicpLAogICAgICBib3JkZXJXaWR0aDogMSwKICAgICAgY29ybmVyUmFkaXVzOiAxMCwKICAgICAgcGFkZGluZzogMTIsCiAgICAgIGJveFBhZGRpbmc6IDQsCiAgICAgIHRpdGxlRm9udDogeyBzaXplOiAxMiwgd2VpZ2h0OiAnNzAwJyB9LAogICAgICBib2R5Rm9udDogeyBzaXplOiAxMiB9LAogICAgfTsKICB9CiAgZnVuY3Rpb24gbGFiZWxDb2xvcigpIHsKICAgIHJldHVybiBjc3NWYXIoJy0tdGV4dC1wcmltYXJ5Jyk7CiAgfQogIC8qKiBTbmFwcHksIHN1YnRsZSBtb3Rpb24g4oCUIGluIHRoZSAxNTAtMzAwbXMgcmFuZ2UgdGhlIHJlZGVzaWduIGNhbGxzIGZvciwgbmV2ZXIgYm91bmN5LiAqLwogIGZ1bmN0aW9uIGJhc2VBbmltYXRpb24oKSB7CiAgICByZXR1cm4geyBkdXJhdGlvbjogMjgwLCBlYXNpbmc6ICdlYXNlT3V0UXVhcnQnIH07CiAgfQoKICBmdW5jdGlvbiBkZXN0cm95KGNhbnZhc0lkKSB7CiAgICBpZiAocmVnaXN0cnkuaGFzKGNhbnZhc0lkKSkgewogICAgICByZWdpc3RyeS5nZXQoY2FudmFzSWQpLmRlc3Ryb3koKTsKICAgICAgcmVnaXN0cnkuZGVsZXRlKGNhbnZhc0lkKTsKICAgIH0KICB9CgogIC8qKiBNdWx0aS1zZXJpZXMgbGluZSBjaGFydCAoZS5nLiB3ZWVrbHkgdHJlbmQgcGVyIHBsYXRmb3JtKS4gT25lIHNlcmllcyBuZWVkcyBubyBsZWdlbmQgYm94LgogICAgICBQZXItcG9pbnQgdmFsdWUgbGFiZWxzIGFyZSBzaG93biBvbmx5IGZvciBhIHNpbmdsZSBzZXJpZXMg4oCUIHdpdGggc2V2ZXJhbCBzZXJpZXMgb3ZlcmxhaWQsCiAgICAgIGxhYmVsaW5nIGV2ZXJ5IHBvaW50IHdvdWxkIG92ZXJsYXAsIHNvIHRob3NlIHJlbHkgb24gdGhlIChzdGlsbC1wcmVzZW50KSBob3ZlciB0b29sdGlwLiAqLwogIGZ1bmN0aW9uIHRyZW5kQ2hhcnQoY2FudmFzSWQsIHsgbGFiZWxzLCBzZXJpZXMsIGZvcm1hdFZhbHVlIH0pIHsKICAgIGRlc3Ryb3koY2FudmFzSWQpOwogICAgY29uc3QgY3R4ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoY2FudmFzSWQpOwogICAgaWYgKCFjdHgpIHJldHVybiBudWxsOwogICAgY29uc3QgZm10ID0gZm9ybWF0VmFsdWUgfHwgKCh2KSA9PiBGb3JtYXQuc21hcnQodikpOwogICAgY29uc3Qgc2hvd0xhYmVscyA9IHNlcmllcy5sZW5ndGggPT09IDEgJiYgbGFiZWxzLmxlbmd0aCA8PSBNQVhfTEFCRUxFRF9JVEVNUzsKCiAgICBjb25zdCBkYXRhc2V0cyA9IHNlcmllcy5tYXAoKHMsIGkpID0+ICh7CiAgICAgIGxhYmVsOiBzLmxhYmVsLAogICAgICBkYXRhOiBzLmRhdGEsCiAgICAgIGJvcmRlckNvbG9yOiBzLmNvbG9yIHx8IHNlcmllc0NvbG9yKGkpLAogICAgICBiYWNrZ3JvdW5kQ29sb3I6IHMuY29sb3IgfHwgc2VyaWVzQ29sb3IoaSksCiAgICAgIGJvcmRlcldpZHRoOiAyLAogICAgICBwb2ludFJhZGl1czogc2hvd0xhYmVscyA/IDMgOiAwLAogICAgICBwb2ludEhvdmVyUmFkaXVzOiA0LAogICAgICBwb2ludEhpdFJhZGl1czogMTIsCiAgICAgIHRlbnNpb246IDAuMjUsCiAgICAgIGZpbGw6IGZhbHNlLAogICAgfSkpOwoKICAgIGNvbnN0IGNoYXJ0ID0gbmV3IENoYXJ0KGN0eCwgewogICAgICB0eXBlOiAnbGluZScsCiAgICAgIGRhdGE6IHsgbGFiZWxzLCBkYXRhc2V0cyB9LAogICAgICBvcHRpb25zOiB7CiAgICAgICAgcmVzcG9uc2l2ZTogdHJ1ZSwKICAgICAgICBtYWludGFpbkFzcGVjdFJhdGlvOiBmYWxzZSwKICAgICAgICBpbnRlcmFjdGlvbjogeyBtb2RlOiAnaW5kZXgnLCBpbnRlcnNlY3Q6IGZhbHNlIH0sCiAgICAgICAgbGF5b3V0OiB7IHBhZGRpbmc6IHsgdG9wOiBzaG93TGFiZWxzID8gMjAgOiA4IH0gfSwKICAgICAgICBhbmltYXRpb246IGJhc2VBbmltYXRpb24oKSwKICAgICAgICBwbHVnaW5zOiB7CiAgICAgICAgICBsZWdlbmQ6IHsKICAgICAgICAgICAgZGlzcGxheTogc2VyaWVzLmxlbmd0aCA+IDEsCiAgICAgICAgICAgIHBvc2l0aW9uOiAnYm90dG9tJywKICAgICAgICAgICAgbGFiZWxzOiB7IGNvbG9yOiBjc3NWYXIoJy0tdGV4dC1zZWNvbmRhcnknKSwgdXNlUG9pbnRTdHlsZTogdHJ1ZSwgcG9pbnRTdHlsZTogJ2xpbmUnLCBib3hXaWR0aDogMTYsIHBhZGRpbmc6IDE2LCBmb250OiB7IHNpemU6IDExIH0gfSwKICAgICAgICAgIH0sCiAgICAgICAgICB0b29sdGlwOiB7IC4uLmJhc2VUb29sdGlwKCksIHVzZVBvaW50U3R5bGU6IHRydWUgfSwKICAgICAgICAgIGRhdGFsYWJlbHM6IHNob3dMYWJlbHMKICAgICAgICAgICAgPyB7IGFsaWduOiAndG9wJywgYW5jaG9yOiAnZW5kJywgY29sb3I6IGxhYmVsQ29sb3IoKSwgZm9udDogeyBzaXplOiAxMSwgd2VpZ2h0OiAnNjAwJyB9LCBmb3JtYXR0ZXI6ICh2KSA9PiBmbXQodikgfQogICAgICAgICAgICA6IHsgZGlzcGxheTogZmFsc2UgfSwKICAgICAgICB9LAogICAgICAgIHNjYWxlczogewogICAgICAgICAgeDogeyBncmlkOiB7IGRpc3BsYXk6IGZhbHNlIH0sIHRpY2tzOiBiYXNlVGlja3MoKSB9LAogICAgICAgICAgeTogeyBncmlkOiBiYXNlR3JpZCgpLCB0aWNrczogYmFzZVRpY2tzKCksIGJvcmRlcjogeyBkaXNwbGF5OiBmYWxzZSB9LCBiZWdpbkF0WmVybzogdHJ1ZSB9LAogICAgICAgIH0sCiAgICAgIH0sCiAgICB9KTsKICAgIHJlZ2lzdHJ5LnNldChjYW52YXNJZCwgY2hhcnQpOwogICAgcmV0dXJuIGNoYXJ0OwogIH0KCiAgLyoqIFNpbmdsZS1tZXRyaWMgYmFyIGNoYXJ0IGFjcm9zcyBwbGF0Zm9ybXMgKGlkZW50aXR5IGVuY29kaW5nIOKAlCBlYWNoIGJhciBJUyBhIHBsYXRmb3JtKS4gKi8KICBmdW5jdGlvbiBwbGF0Zm9ybUJhckNoYXJ0KGNhbnZhc0lkLCB7IGxhYmVscywgZGF0YSwgY29sb3JzLCBmb3JtYXRWYWx1ZSB9KSB7CiAgICBkZXN0cm95KGNhbnZhc0lkKTsKICAgIGNvbnN0IGN0eCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGNhbnZhc0lkKTsKICAgIGlmICghY3R4KSByZXR1cm4gbnVsbDsKICAgIGNvbnN0IGZtdCA9IGZvcm1hdFZhbHVlIHx8ICgodikgPT4gRm9ybWF0LnNtYXJ0KHYpKTsKICAgIGNvbnN0IHNob3dMYWJlbHMgPSBsYWJlbHMubGVuZ3RoIDw9IE1BWF9MQUJFTEVEX0lURU1TOwoKICAgIGNvbnN0IGNoYXJ0ID0gbmV3IENoYXJ0KGN0eCwgewogICAgICB0eXBlOiAnYmFyJywKICAgICAgZGF0YTogewogICAgICAgIGxhYmVscywKICAgICAgICBkYXRhc2V0czogWwogICAgICAgICAgewogICAgICAgICAgICBkYXRhLAogICAgICAgICAgICBiYWNrZ3JvdW5kQ29sb3I6IGNvbG9ycywKICAgICAgICAgICAgYm9yZGVyUmFkaXVzOiA0LAogICAgICAgICAgICBtYXhCYXJUaGlja25lc3M6IDI4LAogICAgICAgICAgICBib3JkZXJTa2lwcGVkOiAnYm90dG9tJywKICAgICAgICAgIH0sCiAgICAgICAgXSwKICAgICAgfSwKICAgICAgb3B0aW9uczogewogICAgICAgIHJlc3BvbnNpdmU6IHRydWUsCiAgICAgICAgbWFpbnRhaW5Bc3BlY3RSYXRpbzogZmFsc2UsCiAgICAgICAgbGF5b3V0OiB7IHBhZGRpbmc6IHsgdG9wOiBzaG93TGFiZWxzID8gMjAgOiA4IH0gfSwKICAgICAgICBhbmltYXRpb246IGJhc2VBbmltYXRpb24oKSwKICAgICAgICBwbHVnaW5zOiB7CiAgICAgICAgICBsZWdlbmQ6IHsgZGlzcGxheTogZmFsc2UgfSwKICAgICAgICAgIHRvb2x0aXA6IGJhc2VUb29sdGlwKCksCiAgICAgICAgICBkYXRhbGFiZWxzOiBzaG93TGFiZWxzCiAgICAgICAgICAgID8geyBhbGlnbjogJ2VuZCcsIGFuY2hvcjogJ2VuZCcsIGNvbG9yOiBsYWJlbENvbG9yKCksIGZvbnQ6IHsgc2l6ZTogMTEsIHdlaWdodDogJzYwMCcgfSwgZm9ybWF0dGVyOiAodikgPT4gZm10KHYpIH0KICAgICAgICAgICAgOiB7IGRpc3BsYXk6IGZhbHNlIH0sCiAgICAgICAgfSwKICAgICAgICBzY2FsZXM6IHsKICAgICAgICAgIHg6IHsgZ3JpZDogeyBkaXNwbGF5OiBmYWxzZSB9LCB0aWNrczogYmFzZVRpY2tzKCkgfSwKICAgICAgICAgIHk6IHsgZ3JpZDogYmFzZUdyaWQoKSwgdGlja3M6IGJhc2VUaWNrcygpLCBib3JkZXI6IHsgZGlzcGxheTogZmFsc2UgfSwgYmVnaW5BdFplcm86IHRydWUgfSwKICAgICAgICB9LAogICAgICB9LAogICAgfSk7CiAgICByZWdpc3RyeS5zZXQoY2FudmFzSWQsIGNoYXJ0KTsKICAgIHJldHVybiBjaGFydDsKICB9CgogIC8qKiBQaWUgY2hhcnQgKGEgaGFuZGZ1bCBvZiBjYXRlZ29yaWVzIG9ubHkg4oCUIGUuZy4gQ2FtcGFpZ24gUGVyZm9ybWFuY2UncyBBZHMvT3JnYW5pYyBzcGxpdCkuCiAgICAgIFNsaWNlIGxhYmVscyBzaG93IGJvdGggc2hhcmUtb2Ytd2hvbGUgYW5kIHRoZSBhY3R1YWwgdmFsdWUsIHBlciB0aGUgIm5vIGhvdmVyIHJlcXVpcmVkIiBnb2FsLiAqLwogIGZ1bmN0aW9uIHBpZUNoYXJ0KGNhbnZhc0lkLCB7IGxhYmVscywgZGF0YSwgY29sb3JzLCBmb3JtYXRWYWx1ZSB9KSB7CiAgICBkZXN0cm95KGNhbnZhc0lkKTsKICAgIGNvbnN0IGN0eCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGNhbnZhc0lkKTsKICAgIGlmICghY3R4KSByZXR1cm4gbnVsbDsKICAgIGNvbnN0IGZtdCA9IGZvcm1hdFZhbHVlIHx8ICgodikgPT4gRm9ybWF0LnNtYXJ0KHYpKTsKICAgIGNvbnN0IHRvdGFsID0gZGF0YS5yZWR1Y2UoKHN1bSwgdikgPT4gc3VtICsgKHYgfHwgMCksIDApOwoKICAgIGNvbnN0IGNoYXJ0ID0gbmV3IENoYXJ0KGN0eCwgewogICAgICB0eXBlOiAncGllJywKICAgICAgZGF0YTogewogICAgICAgIGxhYmVscywKICAgICAgICBkYXRhc2V0czogW3sgZGF0YSwgYmFja2dyb3VuZENvbG9yOiBjb2xvcnMsIGJvcmRlckNvbG9yOiBjc3NWYXIoJy0tc3VyZmFjZS0xJyksIGJvcmRlcldpZHRoOiAyIH1dLAogICAgICB9LAogICAgICBvcHRpb25zOiB7CiAgICAgICAgcmVzcG9uc2l2ZTogdHJ1ZSwKICAgICAgICBtYWludGFpbkFzcGVjdFJhdGlvOiBmYWxzZSwKICAgICAgICBhbmltYXRpb246IGJhc2VBbmltYXRpb24oKSwKICAgICAgICBwbHVnaW5zOiB7CiAgICAgICAgICBsZWdlbmQ6IHsgZGlzcGxheTogdHJ1ZSwgcG9zaXRpb246ICdib3R0b20nLCBsYWJlbHM6IHsgY29sb3I6IGNzc1ZhcignLS10ZXh0LXNlY29uZGFyeScpLCBib3hXaWR0aDogMTIsIHBhZGRpbmc6IDE2LCBmb250OiB7IHNpemU6IDExIH0gfSB9LAogICAgICAgICAgdG9vbHRpcDogYmFzZVRvb2x0aXAoKSwKICAgICAgICAgIGRhdGFsYWJlbHM6IHsKICAgICAgICAgICAgY29sb3I6ICcjZmZmJywKICAgICAgICAgICAgZm9udDogeyBzaXplOiAxMiwgd2VpZ2h0OiAnNzAwJyB9LAogICAgICAgICAgICBmb3JtYXR0ZXI6ICh2KSA9PiB7CiAgICAgICAgICAgICAgY29uc3QgcGN0ID0gdG90YWwgPyBNYXRoLnJvdW5kKCh2IC8gdG90YWwpICogMTAwMCkgLyAxMCA6IDA7CiAgICAgICAgICAgICAgcmV0dXJuIGAke3BjdH0lXG4ke2ZtdCh2KX1gOwogICAgICAgICAgICB9LAogICAgICAgICAgfSwKICAgICAgICB9LAogICAgICB9LAogICAgfSk7CiAgICByZWdpc3RyeS5zZXQoY2FudmFzSWQsIGNoYXJ0KTsKICAgIHJldHVybiBjaGFydDsKICB9CgogIGZ1bmN0aW9uIGRlc3Ryb3lBbGwoKSB7CiAgICBbLi4ucmVnaXN0cnkua2V5cygpXS5mb3JFYWNoKGRlc3Ryb3kpOwogIH0KCiAgcmV0dXJuIHsgdHJlbmRDaGFydCwgcGxhdGZvcm1CYXJDaGFydCwgcGllQ2hhcnQsIHNlcmllc0NvbG9yLCBkZXN0cm95LCBkZXN0cm95QWxsIH07Cn0pKCk7CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgRGFzaGJvYXJkIHRhYjogYSBtZXRyaWMtZm9jdXNlZCBwcmVtaXVtIEJJIGRhc2hib2FyZC4gQSBzaW5nbGUKICAgTWV0cmljIHNlbGVjdG9yIChkeW5hbWljYWxseSBwb3B1bGF0ZWQgZnJvbSB3aGF0ZXZlciB0aGUKICAgc2VsZWN0ZWQgcGxhdGZvcm0ncyBkYXRhIGFjdHVhbGx5IGhhcyDigJQgbmV2ZXIgaGFyZGNvZGVkKSBkcml2ZXMKICAgdGhlIEtQSSBjYXJkcywgd2Vla2x5IHRyZW5kLCBwbGF0Zm9ybS9jYW1wYWlnbi9jb250ZW50LXR5cGUKICAgYnJlYWtkb3ducywgYW5kIHRoZSBUb3AgUGVyZm9ybWluZyBQb3N0cyByYW5raW5nIHRvZ2V0aGVyOwogICBQbGF0Zm9ybS9kYXRlL2NhbXBhaWduL2NvbnRlbnQtdHlwZSBmaWx0ZXJpbmcgY29tZXMgZnJvbSB0aGUKICAgc2hhcmVkIGZpbHRlciBiYXIuIEV2ZXJ5IGNoYXJ0IHNob3dzIGl0cyB2YWx1ZXMgZGlyZWN0bHkgKHZpYQogICBjaGFydGpzLXBsdWdpbi1kYXRhbGFiZWxzKSBzbyBub3RoaW5nIHJlcXVpcmVzIGEgaG92ZXIgdG8gcmVhZC4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmNvbnN0IERhc2hib2FyZCA9ICgoKSA9PiB7CiAgbGV0IHJvb3Q7CiAgbGV0IG1ldHJpYyA9ICd2aWV3cyc7CiAgbGV0IG1ldHJpY09wdGlvbnMgPSBbXTsKCiAgZnVuY3Rpb24gb3B0aW9uRm9yKGtleSkgewogICAgcmV0dXJuIG1ldHJpY09wdGlvbnMuZmluZCgobSkgPT4gbS5rZXkgPT09IGtleSk7CiAgfQogIGZ1bmN0aW9uIG1ldHJpY0xhYmVsKGtleSkgewogICAgY29uc3Qgb3B0ID0gb3B0aW9uRm9yKGtleSk7CiAgICByZXR1cm4gb3B0ID8gb3B0LmxhYmVsIDoga2V5OwogIH0KICBmdW5jdGlvbiBtZXRyaWNVbml0KGtleSkgewogICAgY29uc3Qgb3B0ID0gb3B0aW9uRm9yKGtleSk7CiAgICByZXR1cm4gb3B0ID8gb3B0LnVuaXQgOiAnbnVtYmVyJzsKICB9CiAgZnVuY3Rpb24gZm9ybWF0TWV0cmljVmFsdWUoa2V5LCB2YWx1ZSkgewogICAgY29uc3QgdW5pdCA9IG1ldHJpY1VuaXQoa2V5KTsKICAgIGlmICh2YWx1ZSA9PT0gbnVsbCB8fCB2YWx1ZSA9PT0gdW5kZWZpbmVkKSByZXR1cm4gJ+KAlCc7CiAgICBpZiAodW5pdCA9PT0gJ2R1cmF0aW9uJykgcmV0dXJuIEZvcm1hdC5kdXJhdGlvbih2YWx1ZSk7CiAgICByZXR1cm4gRm9ybWF0LnNtYXJ0KHZhbHVlKTsKICB9CgogIGZ1bmN0aW9uIHNoZWxsKCkgewogICAgcm9vdC5pbm5lckhUTUwgPSAnJzsKCiAgICBjb25zdCBjb250cm9scyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY29udHJvbHMuY2xhc3NOYW1lID0gJ2Rhc2hib2FyZC1jb250cm9scyc7CiAgICBjb25zdCBsYWJlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2xhYmVsJyk7CiAgICBsYWJlbC50ZXh0Q29udGVudCA9ICdNZXRyaWMnOwogICAgbGFiZWwuc2V0QXR0cmlidXRlKCdmb3InLCAnZGFzaGJvYXJkTWV0cmljU2VsZWN0Jyk7CiAgICBjb25zdCBtZXRyaWNTZWxlY3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzZWxlY3QnKTsKICAgIG1ldHJpY1NlbGVjdC5pZCA9ICdkYXNoYm9hcmRNZXRyaWNTZWxlY3QnOwogICAgbWV0cmljT3B0aW9ucy5mb3JFYWNoKChtKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOwogICAgICBvcHQudmFsdWUgPSBtLmtleTsKICAgICAgb3B0LnRleHRDb250ZW50ID0gbS5sYWJlbDsKICAgICAgaWYgKG0ua2V5ID09PSBtZXRyaWMpIG9wdC5zZWxlY3RlZCA9IHRydWU7CiAgICAgIG1ldHJpY1NlbGVjdC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgICBtZXRyaWNTZWxlY3QuYWRkRXZlbnRMaXN0ZW5lcignY2hhbmdlJywgKCkgPT4gewogICAgICBtZXRyaWMgPSBtZXRyaWNTZWxlY3QudmFsdWU7CiAgICAgIHJlZnJlc2hGb3JNZXRyaWMoKTsKICAgIH0pOwogICAgY29udHJvbHMuYXBwZW5kKGxhYmVsLCBtZXRyaWNTZWxlY3QpOwogICAgcm9vdC5hcHBlbmRDaGlsZChjb250cm9scyk7CgogICAgY29uc3Qgc3VtbWFyeUJhciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgc3VtbWFyeUJhci5jbGFzc05hbWUgPSAncGxhdGZvcm0tc3VtbWFyeS1iYXInOwogICAgc3VtbWFyeUJhci5pZCA9ICdwbGF0Zm9ybVN1bW1hcnlCYXInOwogICAgcm9vdC5hcHBlbmRDaGlsZChzdW1tYXJ5QmFyKTsKCiAgICBjb25zdCBrcGlUaXRsZSA9IHRleHRFbCgnZGl2JywgJ0tleSBwZXJmb3JtYW5jZSBpbmRpY2F0b3JzJywgJ3NlY3Rpb24tdGl0bGUnKTsKICAgIGNvbnN0IGtwaUdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGtwaUdyaWQuY2xhc3NOYW1lID0gJ3N0YXQtZ3JpZCc7CiAgICBrcGlHcmlkLmlkID0gJ2twaUdyaWQnOwogICAgcm9vdC5hcHBlbmQoa3BpVGl0bGUsIGtwaUdyaWQpOwoKICAgIGNvbnN0IGNoYXJ0c1RpdGxlID0gdGV4dEVsKCdkaXYnLCAnVHJlbmQgJiBwZXJmb3JtYW5jZSBicmVha2Rvd24nLCAnc2VjdGlvbi10aXRsZScpOwogICAgcm9vdC5hcHBlbmQoY2hhcnRzVGl0bGUpOwoKICAgIGNvbnN0IHRyZW5kQ2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdHJlbmRDYXJkLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGNvbnN0IHRyZW5kSGVhZGVyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0cmVuZEhlYWRlci5jbGFzc05hbWUgPSAnY2FyZC1oZWFkZXInOwogICAgdHJlbmRIZWFkZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMycsICdXZWVrbHkgcGVyZm9ybWFuY2UnKSk7CiAgICB0cmVuZEhlYWRlci5maXJzdENoaWxkLmlkID0gJ3RyZW5kQ2FyZFRpdGxlJzsKICAgIHRyZW5kQ2FyZC5hcHBlbmRDaGlsZCh0cmVuZEhlYWRlcik7CiAgICBjb25zdCB0cmVuZENoYXJ0V3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdHJlbmRDaGFydFdyYXAuY2xhc3NOYW1lID0gJ2NoYXJ0LXdyYXAgdGFsbCc7CiAgICB0cmVuZENoYXJ0V3JhcC5pZCA9ICd0cmVuZENoYXJ0V3JhcCc7CiAgICB0cmVuZENoYXJ0V3JhcC5pbm5lckhUTUwgPSAnPGNhbnZhcyBpZD0idHJlbmRDYW52YXMiPjwvY2FudmFzPic7CiAgICB0cmVuZENhcmQuYXBwZW5kQ2hpbGQodHJlbmRDaGFydFdyYXApOwogICAgcm9vdC5hcHBlbmRDaGlsZCh0cmVuZENhcmQpOwoKICAgIGNvbnN0IGdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGdyaWQuY2xhc3NOYW1lID0gJ2NhcmQtZ3JpZCBldmVuJzsKICAgIGdyaWQuc3R5bGUubWFyZ2luVG9wID0gJzE2cHgnOwoKICAgIGNvbnN0IGJyZWFrZG93bkNhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGJyZWFrZG93bkNhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgY29uc3QgYnJlYWtkb3duSGVhZGVyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBicmVha2Rvd25IZWFkZXIuY2xhc3NOYW1lID0gJ2NhcmQtaGVhZGVyJzsKICAgIGJyZWFrZG93bkhlYWRlci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2gzJywgJycpKTsKICAgIGJyZWFrZG93bkhlYWRlci5maXJzdENoaWxkLmlkID0gJ2JyZWFrZG93bkNhcmRUaXRsZSc7CiAgICBicmVha2Rvd25DYXJkLmFwcGVuZENoaWxkKGJyZWFrZG93bkhlYWRlcik7CiAgICBjb25zdCBicmVha2Rvd25XcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBicmVha2Rvd25XcmFwLmNsYXNzTmFtZSA9ICdjaGFydC13cmFwJzsKICAgIGJyZWFrZG93bldyYXAuaWQgPSAnYnJlYWtkb3duQ2hhcnRXcmFwJzsKICAgIGJyZWFrZG93bldyYXAuaW5uZXJIVE1MID0gJzxjYW52YXMgaWQ9ImJyZWFrZG93bkNhbnZhcyI+PC9jYW52YXM+JzsKICAgIGJyZWFrZG93bkNhcmQuYXBwZW5kQ2hpbGQoYnJlYWtkb3duV3JhcCk7CgogICAgY29uc3QgY29udGVudFR5cGVDYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjb250ZW50VHlwZUNhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgY29uc3QgY29udGVudFR5cGVIZWFkZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNvbnRlbnRUeXBlSGVhZGVyLmNsYXNzTmFtZSA9ICdjYXJkLWhlYWRlcic7CiAgICBjb250ZW50VHlwZUhlYWRlci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2gzJywgJycpKTsKICAgIGNvbnRlbnRUeXBlSGVhZGVyLmZpcnN0Q2hpbGQuaWQgPSAnY29udGVudFR5cGVDYXJkVGl0bGUnOwogICAgY29udGVudFR5cGVDYXJkLmFwcGVuZENoaWxkKGNvbnRlbnRUeXBlSGVhZGVyKTsKICAgIGNvbnN0IGNvbnRlbnRUeXBlV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY29udGVudFR5cGVXcmFwLmNsYXNzTmFtZSA9ICdjaGFydC13cmFwJzsKICAgIGNvbnRlbnRUeXBlV3JhcC5pZCA9ICdjb250ZW50VHlwZUNoYXJ0V3JhcCc7CiAgICBjb250ZW50VHlwZVdyYXAuaW5uZXJIVE1MID0gJzxjYW52YXMgaWQ9ImNvbnRlbnRUeXBlQ2FudmFzIj48L2NhbnZhcz4nOwogICAgY29udGVudFR5cGVDYXJkLmFwcGVuZENoaWxkKGNvbnRlbnRUeXBlV3JhcCk7CgogICAgZ3JpZC5hcHBlbmQoYnJlYWtkb3duQ2FyZCwgY29udGVudFR5cGVDYXJkKTsKICAgIHJvb3QuYXBwZW5kQ2hpbGQoZ3JpZCk7CgogICAgY29uc3QgdG9wVGl0bGUgPSB0ZXh0RWwoJ2RpdicsICdUb3AtcGVyZm9ybWluZyBwb3N0cycsICdzZWN0aW9uLXRpdGxlJyk7CiAgICBjb25zdCB0b3BDYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0b3BDYXJkLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGNvbnN0IHRvcEhlYWRlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdG9wSGVhZGVyLmNsYXNzTmFtZSA9ICdjYXJkLWhlYWRlcic7CiAgICB0b3BIZWFkZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMycsICdSYW5rZWQgYnkgc2VsZWN0ZWQgbWV0cmljJykpOwogICAgdG9wQ2FyZC5hcHBlbmRDaGlsZCh0b3BIZWFkZXIpOwogICAgY29uc3QgdGFibGVXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0YWJsZVdyYXAuY2xhc3NOYW1lID0gJ3RhYmxlLXNjcm9sbCc7CiAgICB0YWJsZVdyYXAuaWQgPSAndG9wUG9zdHNUYWJsZSc7CiAgICB0b3BDYXJkLmFwcGVuZENoaWxkKHRhYmxlV3JhcCk7CiAgICByb290LmFwcGVuZCh0b3BUaXRsZSwgdG9wQ2FyZCk7CiAgfQoKICAvKiogQWxsIHBvc3RzIHRpZWQgYXQgdGhlIHRvcCB2YWx1ZSDigJQgbm90IGp1c3Qgb25lLiBUaWVzIGFyZSByYXJlLCBzbyB0aGlzCiAgICAgIHVzdWFsbHkgcmVuZGVycyBleGFjdGx5IGxpa2UgYSBzaW5nbGUtcG9zdCB0aWxlOyB3aGVuIHRoZXJlIElTIGEgdGllLAogICAgICBldmVyeSB0aWVkIHBvc3QgaXMgbGlzdGVkICh0aGUgdmFsdWUgaXMgc2hvd24gb25jZSwgc2luY2UgdGllcyBzaGFyZSBpdCkuICovCiAgZnVuY3Rpb24gYmVzdFBvc3RzVGlsZShsYWJlbCwgcG9zdHMsIGN1cnJlbnRGb2xsb3dlcnMpIHsKICAgIGNvbnN0IHRpbGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRpbGUuY2xhc3NOYW1lID0gJ3N0YXQtdGlsZSBwb3N0LXRpbGUnOwogICAgdGlsZS5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIGxhYmVsLCAnc3RhdC1sYWJlbCcpKTsKICAgIGlmICghcG9zdHMgfHwgIXBvc3RzLmxlbmd0aCkgewogICAgICB0aWxlLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgJ05vIGRhdGEgeWV0JywgJ3N0YXQtdmFsdWUnKSk7CiAgICAgIGFwcGVuZEZvbGxvd2Vyc0Zvb3Rlcih0aWxlLCBjdXJyZW50Rm9sbG93ZXJzKTsKICAgICAgcmV0dXJuIHRpbGU7CiAgICB9CiAgICBpZiAocG9zdHMubGVuZ3RoID4gMSkgewogICAgICB0aWxlLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgYCR7cG9zdHMubGVuZ3RofSBwb3N0cyB0aWVkIGF0IHRoZSB0b3BgLCAncG9zdC1tZXRhJykpOwogICAgfQogICAgY29uc3QgcGxhdGZvcm1PcHRpb25zID0gKHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZSB8fCB7IHBsYXRmb3JtczogW10gfSkucGxhdGZvcm1zOwogICAgcG9zdHMuZm9yRWFjaCgocG9zdCkgPT4gewogICAgICBjb25zdCBtZXRhID0gcGxhdGZvcm1PcHRpb25zLmZpbmQoKHApID0+IHAuaWQgPT09IHBvc3QucGxhdGZvcm0pIHx8IHsgbGFiZWw6IHBvc3QucGxhdGZvcm0gfTsKICAgICAgY29uc3QgY2FwdGlvbiA9IHBvc3QuY2FwdGlvbiB8fCAnKG5vIGNhcHRpb24pJzsKICAgICAgY29uc3QgZW50cnkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgZW50cnkuY2xhc3NOYW1lID0gJ3Bvc3QtdGlsZS1lbnRyeSc7CiAgICAgIGNvbnN0IHZhbHVlRWwgPSB0ZXh0RWwoJ2RpdicsIGNhcHRpb24ubGVuZ3RoID4gNzAgPyBgJHtjYXB0aW9uLnNsaWNlKDAsIDcwKX3igKZgIDogY2FwdGlvbiwgJ3N0YXQtdmFsdWUnKTsKICAgICAgdmFsdWVFbC50aXRsZSA9IGNhcHRpb247CiAgICAgIGVudHJ5LmFwcGVuZENoaWxkKHZhbHVlRWwpOwogICAgICBlbnRyeS5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIGAke21ldGEubGFiZWx9IMK3ICR7Rm9ybWF0LmRhdGUocG9zdC5wdWJsaXNoX2RhdGUpfWAsICdwb3N0LW1ldGEnKSk7CiAgICAgIHRpbGUuYXBwZW5kQ2hpbGQoZW50cnkpOwogICAgfSk7CiAgICB0aWxlLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgZm9ybWF0TWV0cmljVmFsdWUobWV0cmljLCBwb3N0c1swXS52YWx1ZSksICdwb3N0LW1ldHJpYy12YWx1ZScpKTsKICAgIGFwcGVuZEZvbGxvd2Vyc0Zvb3Rlcih0aWxlLCBjdXJyZW50Rm9sbG93ZXJzKTsKICAgIHJldHVybiB0aWxlOwogIH0KCiAgLyoqICJDdXJyZW50IEZvbGxvd2VyczogTiIgZm9vdGVyIGxpbmUgb24gdGhlIEJlc3QgUGVyZm9ybWluZyBQb3N0IGNhcmQsIHNvdXJjZWQgZnJvbSB0aGUgc2FtZSBmb2xsb3dlcnNLcGlzKCkgZmlndXJlIHRoZSBLUEkgY2FyZCBhYm92ZSBzaG93cyDigJQgb21pdHRlZCBlbnRpcmVseSB3aGVuIHRoZXJlJ3Mgbm8gZm9sbG93ZXIgZGF0YSBhdCBhbGwsIHJhdGhlciB0aGFuIHNob3dpbmcgYSBjb25mdXNpbmcgIkN1cnJlbnQgRm9sbG93ZXJzOiDigJQiLiAqLwogIGZ1bmN0aW9uIGFwcGVuZEZvbGxvd2Vyc0Zvb3Rlcih0aWxlLCBjdXJyZW50Rm9sbG93ZXJzKSB7CiAgICBpZiAoY3VycmVudEZvbGxvd2VycyA9PT0gbnVsbCB8fCBjdXJyZW50Rm9sbG93ZXJzID09PSB1bmRlZmluZWQpIHJldHVybjsKICAgIGNvbnN0IGZvb3RlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZm9vdGVyLmNsYXNzTmFtZSA9ICdwb3N0LXRpbGUtZm9vdGVyJzsKICAgIGZvb3Rlci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3NwYW4nLCAnQ3VycmVudCBGb2xsb3dlcnMnLCAncG9zdC10aWxlLWZvb3Rlci1sYWJlbCcpKTsKICAgIGZvb3Rlci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3NwYW4nLCBGb3JtYXQubnVtYmVyKGN1cnJlbnRGb2xsb3dlcnMpLCAncG9zdC10aWxlLWZvb3Rlci12YWx1ZScpKTsKICAgIHRpbGUuYXBwZW5kQ2hpbGQoZm9vdGVyKTsKICB9CgogIGZ1bmN0aW9uIHN0YXRUaWxlKGxhYmVsLCB2YWx1ZSwgZm9ybWF0Rm4pIHsKICAgIGNvbnN0IHRpbGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRpbGUuY2xhc3NOYW1lID0gJ3N0YXQtdGlsZSc7CiAgICBjb25zdCB2YWx1ZUVsID0gdGV4dEVsKCdkaXYnLCAnJywgJ3N0YXQtdmFsdWUnKTsKICAgIHRpbGUuYXBwZW5kKHRleHRFbCgnZGl2JywgbGFiZWwsICdzdGF0LWxhYmVsJyksIHZhbHVlRWwpOwogICAgY29uc3QgZm10ID0gZm9ybWF0Rm4gfHwgKCh2KSA9PiBGb3JtYXQubnVtYmVyKHYpKTsKICAgIGlmICh0eXBlb2YgdmFsdWUgPT09ICdudW1iZXInICYmIE51bWJlci5pc0Zpbml0ZSh2YWx1ZSkpIHsKICAgICAgYW5pbWF0ZUNvdW50KHZhbHVlRWwsIDAsIHZhbHVlLCA5MDAsIGZtdCk7CiAgICB9IGVsc2UgewogICAgICB2YWx1ZUVsLnRleHRDb250ZW50ID0gZm10KHZhbHVlKTsKICAgIH0KICAgIHJldHVybiB0aWxlOwogIH0KCiAgLyoqICJGb2xsb3dlcnMgR3Jvd3RoIiB0aWxlOiBhbiBhYnNvbHV0ZS1kaWZmZXJlbmNlIHN0YXQtdmFsdWUgcGx1cyBhIHBlcmNlbnRhZ2UgZGVsdGEgbGluZSAoYXJyb3cgKyBjb2xvciBkcml2ZW4gYnkgRm9ybWF0LmRlbHRhQ2xhc3MsIHNhbWUgY29udmVudGlvbiBhcyB0aGUgQ29tcGFyaXNvbnMgcGFnZSdzIHN0YXQgdGlsZXMpLiAqLwogIGZ1bmN0aW9uIGZvbGxvd2Vyc0dyb3d0aFRpbGUoY2hhbmdlLCBjaGFuZ2VQY3QpIHsKICAgIGNvbnN0IHRpbGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRpbGUuY2xhc3NOYW1lID0gJ3N0YXQtdGlsZSc7CiAgICBjb25zdCB2YWx1ZUVsID0gdGV4dEVsKCdkaXYnLCAnJywgJ3N0YXQtdmFsdWUnKTsKICAgIHRpbGUuYXBwZW5kKHRleHRFbCgnZGl2JywgJ0ZvbGxvd2VycyBHcm93dGgnLCAnc3RhdC1sYWJlbCcpLCB2YWx1ZUVsKTsKICAgIGlmIChjaGFuZ2UgPT09IG51bGwgfHwgY2hhbmdlID09PSB1bmRlZmluZWQpIHsKICAgICAgdmFsdWVFbC50ZXh0Q29udGVudCA9ICfigJQnOwogICAgfSBlbHNlIHsKICAgICAgYW5pbWF0ZUNvdW50KHZhbHVlRWwsIDAsIGNoYW5nZSwgOTAwLCAodikgPT4gYCR7diA+IDAgPyAnKycgOiAnJ30ke0Zvcm1hdC5udW1iZXIoTWF0aC5yb3VuZCh2KSl9YCk7CiAgICB9CiAgICBjb25zdCBkZWx0YVRleHQgPSBjaGFuZ2VQY3QgPT09IG51bGwgfHwgY2hhbmdlUGN0ID09PSB1bmRlZmluZWQgPyAn4oCUJyA6IEZvcm1hdC5wY3QoY2hhbmdlUGN0KTsKICAgIHRpbGUuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBkZWx0YVRleHQsIGBzdGF0LWRlbHRhICR7Rm9ybWF0LmRlbHRhQ2xhc3MoY2hhbmdlUGN0KX1gKSk7CiAgICByZXR1cm4gdGlsZTsKICB9CgogIC8qKiAiTmV3IEZvbGxvd2VycyIgdGlsZTogZm9sbG93ZXJzIGdhaW5lZCB3aXRoaW4gdGhlIGN1cnJlbnRseSBzZWxlY3RlZCBkYXRlIHJhbmdlIOKAlCBzaG93cyAiTm8gZm9sbG93ZXIgdXBkYXRlIiByYXRoZXIgdGhhbiAwIHdoZW4gbm90aGluZyBpcyBjb21wdXRhYmxlIGZvciB0aGUgcmFuZ2UgKHBlciBzcGVjKSwgd2hpY2ggaXMgZGlmZmVyZW50IGZyb20gYSBnZW51aW5lIHplcm8uICovCiAgZnVuY3Rpb24gbmV3Rm9sbG93ZXJzVGlsZShuZXdGb2xsb3dlcnMpIHsKICAgIGNvbnN0IHRpbGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRpbGUuY2xhc3NOYW1lID0gJ3N0YXQtdGlsZSc7CiAgICBjb25zdCB2YWx1ZUVsID0gdGV4dEVsKCdkaXYnLCAnJywgJ3N0YXQtdmFsdWUnKTsKICAgIHRpbGUuYXBwZW5kKHRleHRFbCgnZGl2JywgJ05ldyBGb2xsb3dlcnMnLCAnc3RhdC1sYWJlbCcpLCB2YWx1ZUVsKTsKICAgIGlmIChuZXdGb2xsb3dlcnMgPT09IG51bGwgfHwgbmV3Rm9sbG93ZXJzID09PSB1bmRlZmluZWQpIHsKICAgICAgdmFsdWVFbC50ZXh0Q29udGVudCA9ICdObyBmb2xsb3dlciB1cGRhdGUnOwogICAgICB2YWx1ZUVsLmNsYXNzTGlzdC5hZGQoJ3N0YXQtdmFsdWUtbXV0ZWQnKTsKICAgIH0gZWxzZSB7CiAgICAgIGFuaW1hdGVDb3VudCh2YWx1ZUVsLCAwLCBuZXdGb2xsb3dlcnMsIDkwMCwgKHYpID0+IGAke3YgPiAwID8gJysnIDogJyd9JHtGb3JtYXQubnVtYmVyKE1hdGgucm91bmQodikpfWApOwogICAgfQogICAgcmV0dXJuIHRpbGU7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJLcGlzKHN1bW1hcnksIGZvbGxvd2VycykgewogICAgY29uc3QgZ3JpZCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdrcGlHcmlkJyk7CiAgICBpZiAoIWdyaWQpIHJldHVybjsKICAgIGdyaWQuaW5uZXJIVE1MID0gJyc7CgogICAgZ3JpZC5hcHBlbmRDaGlsZChzdGF0VGlsZSgnSGlnaGVzdCBWYWx1ZScsIHN1bW1hcnkuaGlnaGVzdCwgKHYpID0+IGZvcm1hdE1ldHJpY1ZhbHVlKG1ldHJpYywgdikpKTsKICAgIGdyaWQuYXBwZW5kQ2hpbGQoc3RhdFRpbGUoJ0F2ZXJhZ2UgVmFsdWUnLCBzdW1tYXJ5LmF2ZXJhZ2UsICh2KSA9PiBmb3JtYXRNZXRyaWNWYWx1ZShtZXRyaWMsIHYpKSk7CiAgICBncmlkLmFwcGVuZENoaWxkKHN0YXRUaWxlKCdUb3RhbCBWYWx1ZScsIHN1bW1hcnkudG90YWwsICh2KSA9PiBmb3JtYXRNZXRyaWNWYWx1ZShtZXRyaWMsIHYpKSk7CiAgICBncmlkLmFwcGVuZENoaWxkKHN0YXRUaWxlKCdOdW1iZXIgb2YgUG9zdHMnLCBzdW1tYXJ5LnBvc3RDb3VudCwgKHYpID0+IEZvcm1hdC5udW1iZXIodikpKTsKICAgIGdyaWQuYXBwZW5kQ2hpbGQoYmVzdFBvc3RzVGlsZSgnQmVzdCBQZXJmb3JtaW5nIFBvc3QnLCBzdW1tYXJ5LmJlc3RQb3N0cywgZm9sbG93ZXJzLmN1cnJlbnRGb2xsb3dlcnMpKTsKICAgIGdyaWQuYXBwZW5kQ2hpbGQoc3RhdFRpbGUoJ0N1cnJlbnQgRm9sbG93ZXJzJywgZm9sbG93ZXJzLmN1cnJlbnRGb2xsb3dlcnMsICh2KSA9PiAodiA9PT0gbnVsbCB8fCB2ID09PSB1bmRlZmluZWQgPyAn4oCUJyA6IEZvcm1hdC5udW1iZXIodikpKSk7CiAgICBncmlkLmFwcGVuZENoaWxkKGZvbGxvd2Vyc0dyb3d0aFRpbGUoZm9sbG93ZXJzLmZvbGxvd2Vyc0NoYW5nZSwgZm9sbG93ZXJzLmZvbGxvd2Vyc0NoYW5nZVBjdCkpOwogICAgZ3JpZC5hcHBlbmRDaGlsZChuZXdGb2xsb3dlcnNUaWxlKGZvbGxvd2Vycy5uZXdGb2xsb3dlcnMpKTsKICB9CgogIC8qKiBDb21wYWN0IGhvcml6b250YWwgc3RyaXAgYWJvdmUgdGhlIEtQSSBncmlkIOKAlCBvbmUgY2FyZCBwZXIgcGxhdGZvcm0gd2l0aCBhbnkgZGF0YSwgc2hvd2luZyBGb2xsb3dlcnMvUG9zdHMvVmlld3MvUmVhY2gvRW5nYWdlbWVudC4gUmVmcmVzaGVzIG9uIHRoZSBzYW1lIGZpbHRlci1jaGFuZ2UgY3ljbGUgYXMgZXZlcnl0aGluZyBlbHNlIG9uIHRoaXMgdGFiLiAqLwogIGFzeW5jIGZ1bmN0aW9uIHJlbmRlclBsYXRmb3JtU3VtbWFyeUJhcihmaWx0ZXJzKSB7CiAgICBjb25zdCBiYXIgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncGxhdGZvcm1TdW1tYXJ5QmFyJyk7CiAgICBpZiAoIWJhcikgcmV0dXJuOwogICAgY29uc3Qgcm93cyA9IGF3YWl0IEFwaS5wbGF0Zm9ybVN1bW1hcnkoZmlsdGVycyk7CiAgICBiYXIuaW5uZXJIVE1MID0gJyc7CiAgICBpZiAoIXJvd3MubGVuZ3RoKSB7CiAgICAgIGJhci5zdHlsZS5kaXNwbGF5ID0gJ25vbmUnOwogICAgICByZXR1cm47CiAgICB9CiAgICBiYXIuc3R5bGUuZGlzcGxheSA9ICcnOwogICAgcm93cy5mb3JFYWNoKChyb3cpID0+IHsKICAgICAgY29uc3QgY2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBjYXJkLmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1zdW1tYXJ5LWNhcmQnOwogICAgICBjb25zdCBoZWFkZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgaGVhZGVyLmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1zdW1tYXJ5LWhlYWRlcic7CiAgICAgIGNvbnN0IGRvdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgICAgZG90LmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1kb3QnOwogICAgICBkb3Quc3R5bGUuYmFja2dyb3VuZCA9IHJvdy5jb2xvcjsKICAgICAgaGVhZGVyLmFwcGVuZChkb3QsIHRleHRFbCgnc3BhbicsIHJvdy5sYWJlbCwgJ3BsYXRmb3JtLXN1bW1hcnktbmFtZScpKTsKICAgICAgY2FyZC5hcHBlbmRDaGlsZChoZWFkZXIpOwoKICAgICAgY29uc3QgbWV0cmljc1JvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBtZXRyaWNzUm93LmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1zdW1tYXJ5LW1ldHJpY3MnOwogICAgICBjb25zdCBlbnRyaWVzID0gWwogICAgICAgIFsnRm9sbG93ZXJzJywgcm93LmZvbGxvd2VycyA9PT0gbnVsbCA/ICfigJQnIDogRm9ybWF0LnNtYXJ0KHJvdy5mb2xsb3dlcnMpXSwKICAgICAgICBbJ1Bvc3RzJywgRm9ybWF0Lm51bWJlcihyb3cucG9zdHMpXSwKICAgICAgICBbJ1ZpZXdzJywgRm9ybWF0LnNtYXJ0KHJvdy52aWV3cyldLAogICAgICAgIFsnUmVhY2gnLCBGb3JtYXQuc21hcnQocm93LnJlYWNoKV0sCiAgICAgICAgWydFbmdhZ2VtZW50JywgRm9ybWF0LnNtYXJ0KHJvdy5lbmdhZ2VtZW50KV0sCiAgICAgIF07CiAgICAgIGVudHJpZXMuZm9yRWFjaCgoW2xhYmVsLCB2YWx1ZV0pID0+IHsKICAgICAgICBjb25zdCBjZWxsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgICAgY2VsbC5jbGFzc05hbWUgPSAncGxhdGZvcm0tc3VtbWFyeS1tZXRyaWMnOwogICAgICAgIGNlbGwuYXBwZW5kKHRleHRFbCgnc3BhbicsIGxhYmVsLCAncHNtLWxhYmVsJyksIHRleHRFbCgnc3BhbicsIHZhbHVlLCAncHNtLXZhbHVlJykpOwogICAgICAgIG1ldHJpY3NSb3cuYXBwZW5kQ2hpbGQoY2VsbCk7CiAgICAgIH0pOwogICAgICBjYXJkLmFwcGVuZENoaWxkKG1ldHJpY3NSb3cpOwogICAgICBiYXIuYXBwZW5kQ2hpbGQoY2FyZCk7CiAgICB9KTsKICB9CgogIC8qKiBTd2FwcyBhIGNoYXJ0IGNhcmQncyBjYW52YXMgZm9yIGFuIGVtcHR5LXN0YXRlIG1lc3NhZ2UsIG9yIHJlc3RvcmVzIHRoZSBjYW52YXMg4oCUIHNpbmNlCiAgICAgIHJlLXJlbmRlcmluZyBhIENoYXJ0LmpzIGluc3RhbmNlIG5lZWRzIGEgbGl2ZSA8Y2FudmFzPiwgbm90IHdoYXRldmVyIHRoZSBsYXN0IHJlbmRlciBsZWZ0IHRoZXJlLiAqLwogIGZ1bmN0aW9uIGNoYXJ0T3JFbXB0eSh3cmFwSWQsIGNhbnZhc0lkLCBoYXNEYXRhLCBlbXB0eU1lc3NhZ2UsIHJlbmRlckZuKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQod3JhcElkKTsKICAgIGlmICghd3JhcCkgcmV0dXJuOwogICAgQ2hhcnRzLmRlc3Ryb3koY2FudmFzSWQpOwogICAgaWYgKCFoYXNEYXRhKSB7CiAgICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQoZW1wdHlTdGF0ZSh7IGljb246ICdiYXItY2hhcnQtMycsIG1lc3NhZ2U6IGVtcHR5TWVzc2FnZSB9KSk7CiAgICAgIHJldHVybjsKICAgIH0KICAgIHdyYXAuaW5uZXJIVE1MID0gYDxjYW52YXMgaWQ9IiR7Y2FudmFzSWR9Ij48L2NhbnZhcz5gOwogICAgcmVuZGVyRm4oKTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIHJlbmRlclRyZW5kKGZpbHRlcnMpIHsKICAgIGNvbnN0IHBsYXRmb3JtT3B0aW9ucyA9ICh3aW5kb3cuX19maWx0ZXJPcHRpb25zQ2FjaGUgfHwgeyBwbGF0Zm9ybXM6IFtdIH0pLnBsYXRmb3JtczsKICAgIGNvbnN0IG1MYWJlbCA9IG1ldHJpY0xhYmVsKG1ldHJpYyk7CiAgICBjb25zdCBwbGF0Zm9ybXNUb0ZldGNoID0gZmlsdGVycy5wbGF0Zm9ybSA9PT0gJ2FsbCcgPyBwbGF0Zm9ybU9wdGlvbnMubWFwKChwKSA9PiBwLmlkKSA6IFtmaWx0ZXJzLnBsYXRmb3JtXTsKICAgIGNvbnN0IHRyZW5kUmVzcG9uc2VzID0gYXdhaXQgUHJvbWlzZS5hbGwoCiAgICAgIHBsYXRmb3Jtc1RvRmV0Y2gubWFwKChwKSA9PgogICAgICAgIEFwaS50cmVuZCh7IGRhdGVGcm9tOiBmaWx0ZXJzLmRhdGVGcm9tLCBkYXRlVG86IGZpbHRlcnMuZGF0ZVRvLCBwbGF0Zm9ybTogcCwgY2FtcGFpZ25UeXBlOiBmaWx0ZXJzLmNhbXBhaWduVHlwZSwgY29udGVudFR5cGU6IGZpbHRlcnMuY29udGVudFR5cGUgfSkKICAgICAgKQogICAgKTsKICAgIGNvbnN0IHdlZWtTZXQgPSBuZXcgU2V0KCk7CiAgICB0cmVuZFJlc3BvbnNlcy5mb3JFYWNoKChyb3dzKSA9PiByb3dzLmZvckVhY2goKHIpID0+IHdlZWtTZXQuYWRkKHIucGVyaW9kKSkpOwogICAgY29uc3Qgd2Vla3MgPSBbLi4ud2Vla1NldF0uc29ydCgpOwogICAgY29uc3Qgc2VyaWVzID0gcGxhdGZvcm1zVG9GZXRjaC5tYXAoKHAsIGkpID0+IHsKICAgICAgY29uc3QgbWV0YSA9IHBsYXRmb3JtT3B0aW9ucy5maW5kKChwbCkgPT4gcGwuaWQgPT09IHApIHx8IHsgbGFiZWw6IHAgfTsKICAgICAgY29uc3QgYnlXZWVrID0gT2JqZWN0LmZyb21FbnRyaWVzKHRyZW5kUmVzcG9uc2VzW2ldLm1hcCgocikgPT4gW3IucGVyaW9kLCByW21ldHJpY11dKSk7CiAgICAgIHJldHVybiB7IGxhYmVsOiBtZXRhLmxhYmVsLCBjb2xvcjogbWV0YS5jb2xvciwgZGF0YTogd2Vla3MubWFwKCh3KSA9PiAoYnlXZWVrW3ddID09PSB1bmRlZmluZWQgPyBudWxsIDogYnlXZWVrW3ddKSkgfTsKICAgIH0pOwoKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd0cmVuZENhcmRUaXRsZScpLnRleHRDb250ZW50ID0KICAgICAgZmlsdGVycy5wbGF0Zm9ybSA9PT0gJ2FsbCcgPyBgV2Vla2x5ICR7bUxhYmVsfSBieSBQbGF0Zm9ybWAgOiBgJHttTGFiZWx9IFRyZW5kYDsKCiAgICBjaGFydE9yRW1wdHkoJ3RyZW5kQ2hhcnRXcmFwJywgJ3RyZW5kQ2FudmFzJywgd2Vla3MubGVuZ3RoID4gMCwgJ05vIGRhdGEgaW4gdGhpcyByYW5nZSB5ZXQuJywgKCkgPT4gewogICAgICBDaGFydHMudHJlbmRDaGFydCgndHJlbmRDYW52YXMnLCB7IGxhYmVsczogd2Vla3MubWFwKEZvcm1hdC5kYXRlKSwgc2VyaWVzLCBmb3JtYXRWYWx1ZTogKHYpID0+IGZvcm1hdE1ldHJpY1ZhbHVlKG1ldHJpYywgdikgfSk7CiAgICB9KTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIHJlbmRlckJyZWFrZG93bihmaWx0ZXJzKSB7CiAgICBjb25zdCBtTGFiZWwgPSBtZXRyaWNMYWJlbChtZXRyaWMpOwogICAgY29uc3QgdGl0bGVFbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdicmVha2Rvd25DYXJkVGl0bGUnKTsKCiAgICBpZiAoZmlsdGVycy5wbGF0Zm9ybSA9PT0gJ2FsbCcpIHsKICAgICAgdGl0bGVFbC50ZXh0Q29udGVudCA9IGBQbGF0Zm9ybSBDb21wYXJpc29uIOKAlCAke21MYWJlbH1gOwogICAgICBjb25zdCBicmVha2Rvd24gPSBhd2FpdCBBcGkucGxhdGZvcm1CcmVha2Rvd24oZmlsdGVycyk7CiAgICAgIGNvbnN0IHNvcnRlZCA9IGJyZWFrZG93bi5maWx0ZXIoKHApID0+IHBbbWV0cmljXSAhPT0gbnVsbCAmJiBwW21ldHJpY10gIT09IHVuZGVmaW5lZCkuc29ydCgoYSwgYikgPT4gYlttZXRyaWNdIC0gYVttZXRyaWNdKTsKICAgICAgY2hhcnRPckVtcHR5KCdicmVha2Rvd25DaGFydFdyYXAnLCAnYnJlYWtkb3duQ2FudmFzJywgc29ydGVkLmxlbmd0aCA+IDAsICdObyBkYXRhIGluIHRoaXMgcmFuZ2UgeWV0LicsICgpID0+IHsKICAgICAgICBDaGFydHMucGxhdGZvcm1CYXJDaGFydCgnYnJlYWtkb3duQ2FudmFzJywgewogICAgICAgICAgbGFiZWxzOiBzb3J0ZWQubWFwKChwKSA9PiBwLmxhYmVsKSwKICAgICAgICAgIGRhdGE6IHNvcnRlZC5tYXAoKHApID0+IHBbbWV0cmljXSksCiAgICAgICAgICBjb2xvcnM6IHNvcnRlZC5tYXAoKHApID0+IHAuY29sb3IpLAogICAgICAgICAgZm9ybWF0VmFsdWU6ICh2KSA9PiBmb3JtYXRNZXRyaWNWYWx1ZShtZXRyaWMsIHYpLAogICAgICAgIH0pOwogICAgICB9KTsKICAgIH0gZWxzZSB7CiAgICAgIHRpdGxlRWwudGV4dENvbnRlbnQgPSBgQ2FtcGFpZ24gUGVyZm9ybWFuY2Ug4oCUICR7bUxhYmVsfWA7CiAgICAgIGNvbnN0IGNhbXBhaWducyA9IGF3YWl0IEFwaS5jYW1wYWlnbkJyZWFrZG93bihmaWx0ZXJzKTsKICAgICAgY29uc3Qgd2l0aFZhbHVlID0gY2FtcGFpZ25zLmZpbHRlcigoYykgPT4gY1ttZXRyaWNdICE9PSBudWxsICYmIGNbbWV0cmljXSAhPT0gdW5kZWZpbmVkICYmIGNbbWV0cmljXSA+IDApOwogICAgICBjaGFydE9yRW1wdHkoJ2JyZWFrZG93bkNoYXJ0V3JhcCcsICdicmVha2Rvd25DYW52YXMnLCB3aXRoVmFsdWUubGVuZ3RoID4gMCwgJ05vIGNhbXBhaWduIGRhdGEgaW4gdGhpcyByYW5nZSB5ZXQuJywgKCkgPT4gewogICAgICAgIENoYXJ0cy5waWVDaGFydCgnYnJlYWtkb3duQ2FudmFzJywgewogICAgICAgICAgbGFiZWxzOiB3aXRoVmFsdWUubWFwKChjKSA9PiBjLmNhbXBhaWduX3R5cGUpLAogICAgICAgICAgZGF0YTogd2l0aFZhbHVlLm1hcCgoYykgPT4gY1ttZXRyaWNdKSwKICAgICAgICAgIGNvbG9yczogd2l0aFZhbHVlLm1hcCgoXywgaSkgPT4gQ2hhcnRzLnNlcmllc0NvbG9yKGkpKSwKICAgICAgICAgIGZvcm1hdFZhbHVlOiAodikgPT4gZm9ybWF0TWV0cmljVmFsdWUobWV0cmljLCB2KSwKICAgICAgICB9KTsKICAgICAgfSk7CiAgICB9CiAgfQoKICBhc3luYyBmdW5jdGlvbiByZW5kZXJDb250ZW50VHlwZUJyZWFrZG93bihmaWx0ZXJzKSB7CiAgICBjb25zdCBtTGFiZWwgPSBtZXRyaWNMYWJlbChtZXRyaWMpOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2NvbnRlbnRUeXBlQ2FyZFRpdGxlJykudGV4dENvbnRlbnQgPSBgQ29udGVudCBUeXBlIFBlcmZvcm1hbmNlIOKAlCAke21MYWJlbH1gOwogICAgY29uc3Qgcm93cyA9IGF3YWl0IEFwaS5jb250ZW50VHlwZUJyZWFrZG93bihmaWx0ZXJzKTsKICAgIGNvbnN0IHNvcnRlZCA9IHJvd3MuZmlsdGVyKChjKSA9PiBjW21ldHJpY10gIT09IG51bGwgJiYgY1ttZXRyaWNdICE9PSB1bmRlZmluZWQpLnNvcnQoKGEsIGIpID0+IGJbbWV0cmljXSAtIGFbbWV0cmljXSk7CiAgICBjaGFydE9yRW1wdHkoJ2NvbnRlbnRUeXBlQ2hhcnRXcmFwJywgJ2NvbnRlbnRUeXBlQ2FudmFzJywgc29ydGVkLmxlbmd0aCA+IDAsICdObyBkYXRhIGluIHRoaXMgcmFuZ2UgeWV0LicsICgpID0+IHsKICAgICAgQ2hhcnRzLnBsYXRmb3JtQmFyQ2hhcnQoJ2NvbnRlbnRUeXBlQ2FudmFzJywgewogICAgICAgIGxhYmVsczogc29ydGVkLm1hcCgoYykgPT4gYy5jb250ZW50X3R5cGUpLAogICAgICAgIGRhdGE6IHNvcnRlZC5tYXAoKGMpID0+IGNbbWV0cmljXSksCiAgICAgICAgY29sb3JzOiBzb3J0ZWQubWFwKChfLCBpKSA9PiBDaGFydHMuc2VyaWVzQ29sb3IoaSkpLAogICAgICAgIGZvcm1hdFZhbHVlOiAodikgPT4gZm9ybWF0TWV0cmljVmFsdWUobWV0cmljLCB2KSwKICAgICAgfSk7CiAgICB9KTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIHJlbmRlclRvcFBvc3RzKGZpbHRlcnMpIHsKICAgIGNvbnN0IHBvc3RzID0gYXdhaXQgQXBpLnRvcFBvc3RzKHsgLi4uZmlsdGVycywgc29ydEJ5OiBtZXRyaWMsIGxpbWl0OiAxMCB9KTsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndG9wUG9zdHNUYWJsZScpOwogICAgaWYgKCF3cmFwKSByZXR1cm47CiAgICBpZiAoIXBvc3RzLmxlbmd0aCkgewogICAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgICB3cmFwLmFwcGVuZENoaWxkKGVtcHR5U3RhdGUoewogICAgICAgIGljb246ICd0cm9waHknLAogICAgICAgIHRpdGxlOiAnTm8gcG9zdHMgaW4gdGhpcyByYW5nZSB5ZXQnLAogICAgICAgIG1lc3NhZ2U6ICdVcGxvYWQgYSB3ZWVrbHkgZXhwb3J0LCBvciB3aWRlbiB0aGUgZGF0ZSByYW5nZSwgdG8gc2VlIHRvcCBwZXJmb3JtZXJzIGhlcmUuJywKICAgICAgICBhY3Rpb25MYWJlbDogJ1VwbG9hZCBkYXRhJywKICAgICAgICBvbkFjdGlvbjogKCkgPT4gZG9jdW1lbnQucXVlcnlTZWxlY3RvcignLnRhYi1idG5bZGF0YS10YWI9InVwbG9hZCJdJyk/LmNsaWNrKCksCiAgICAgIH0pKTsKICAgICAgcmV0dXJuOwogICAgfQogICAgY29uc3QgcGxhdGZvcm1PcHRpb25zID0gKHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZSB8fCB7IHBsYXRmb3JtczogW10gfSkucGxhdGZvcm1zOwoKICAgIGNvbnN0IHRhYmxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGFibGUnKTsKICAgIHRhYmxlLmNsYXNzTmFtZSA9ICdkYXRhLXRhYmxlJzsKICAgIGNvbnN0IHRoZWFkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGhlYWQnKTsKICAgIGNvbnN0IGhlYWRUciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICBoZWFkVHIuYXBwZW5kKAogICAgICB0ZXh0RWwoJ3RoJywgJ1JhbmsnKSwKICAgICAgdGV4dEVsKCd0aCcsICdEYXRlJyksCiAgICAgIHRleHRFbCgndGgnLCAnUGxhdGZvcm0nKSwKICAgICAgdGV4dEVsKCd0aCcsICdDYW1wYWlnbicpLAogICAgICB0ZXh0RWwoJ3RoJywgJ0NvbnRlbnQgVHlwZScpLAogICAgICB0ZXh0RWwoJ3RoJywgJ0NhcHRpb24nKSwKICAgICAgdGV4dEVsKCd0aCcsIG1ldHJpY0xhYmVsKG1ldHJpYyksICdudW0nKQogICAgKTsKICAgIGhlYWRUci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3RoJywgJycpKTsKICAgIHRoZWFkLmFwcGVuZENoaWxkKGhlYWRUcik7CgogICAgY29uc3QgdGJvZHkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0Ym9keScpOwogICAgcG9zdHMuZm9yRWFjaCgocCwgaSkgPT4gewogICAgICBjb25zdCB0ciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAgIGNvbnN0IG1ldGEgPSBwbGF0Zm9ybU9wdGlvbnMuZmluZCgocGwpID0+IHBsLmlkID09PSBwLnBsYXRmb3JtKSB8fCB7IGxhYmVsOiBwLnBsYXRmb3JtLCBjb2xvcjogJyM5OTknIH07CiAgICAgIGNvbnN0IHBsYXRmb3JtVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBjb25zdCBwaWxsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICBwaWxsLmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1waWxsJzsKICAgICAgY29uc3QgZG90ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICBkb3QuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLWRvdCc7CiAgICAgIGRvdC5zdHlsZS5iYWNrZ3JvdW5kID0gbWV0YS5jb2xvcjsKICAgICAgcGlsbC5hcHBlbmQoZG90LCBkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShtZXRhLmxhYmVsKSk7CiAgICAgIHBsYXRmb3JtVGQuYXBwZW5kQ2hpbGQocGlsbCk7CgogICAgICBjb25zdCBjYXB0aW9uID0gcC5jYXB0aW9uIHx8ICcobm8gY2FwdGlvbiknOwogICAgICBjb25zdCB0cnVuY2F0ZWQgPSBjYXB0aW9uLmxlbmd0aCA+IDYwID8gYCR7Y2FwdGlvbi5zbGljZSgwLCA2MCl94oCmYCA6IGNhcHRpb247CiAgICAgIGNvbnN0IGNhcHRpb25UZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIGlmIChwLnBvc3RpbmdfbGluaykgewogICAgICAgIGNvbnN0IGxpbmsgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdhJyk7CiAgICAgICAgbGluay5jbGFzc05hbWUgPSAnY2FwdGlvbi1saW5rJzsKICAgICAgICBsaW5rLmhyZWYgPSBwLnBvc3RpbmdfbGluazsKICAgICAgICBsaW5rLnRhcmdldCA9ICdfYmxhbmsnOwogICAgICAgIGxpbmsucmVsID0gJ25vb3BlbmVyIG5vcmVmZXJyZXInOwogICAgICAgIGxpbmsudGl0bGUgPSBjYXB0aW9uOwogICAgICAgIGxpbmsuYXBwZW5kQ2hpbGQoZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUodHJ1bmNhdGVkKSk7CiAgICAgICAgY2FwdGlvblRkLmFwcGVuZENoaWxkKGxpbmspOwogICAgICB9IGVsc2UgewogICAgICAgIGNhcHRpb25UZC5hcHBlbmRDaGlsZChkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZSh0cnVuY2F0ZWQpKTsKICAgICAgICBjYXB0aW9uVGQudGl0bGUgPSBjYXB0aW9uOwogICAgICB9CgogICAgICB0ci5hcHBlbmQoCiAgICAgICAgdGV4dEVsKCd0ZCcsIGAjJHtpICsgMX1gKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgRm9ybWF0LmRhdGUocC5wdWJsaXNoX2RhdGUpKSwKICAgICAgICBwbGF0Zm9ybVRkLAogICAgICAgIHRleHRFbCgndGQnLCBwLmNhbXBhaWduX3R5cGUgfHwgJ+KAlCcpLAogICAgICAgIHRleHRFbCgndGQnLCBwLmNvbnRlbnRfdHlwZSB8fCAn4oCUJyksCiAgICAgICAgY2FwdGlvblRkLAogICAgICAgIHRleHRFbCgndGQnLCBmb3JtYXRNZXRyaWNWYWx1ZShtZXRyaWMsIHAubWV0cmljX3ZhbHVlKSwgJ251bScpCiAgICAgICk7CgogICAgICBjb25zdCBhY3Rpb25UZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIGNvbnN0IHZpZXdCdG4gPSBpY29uQnRuKCdidG4nLCAnZXllJywgJ1ZpZXcgRGV0YWlscycpOwogICAgICB2aWV3QnRuLmRpc2FibGVkID0gIXAucmF3X3Jvd19pZDsKICAgICAgdmlld0J0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IFJlY29yZHMub3BlblZpZXcocC5yYXdfcm93X2lkKSk7CiAgICAgIGFjdGlvblRkLmFwcGVuZENoaWxkKHZpZXdCdG4pOwogICAgICB0ci5hcHBlbmRDaGlsZChhY3Rpb25UZCk7CgogICAgICB0Ym9keS5hcHBlbmRDaGlsZCh0cik7CiAgICB9KTsKICAgIHRhYmxlLmFwcGVuZCh0aGVhZCwgdGJvZHkpOwogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGFibGUpOwogIH0KCiAgLyoqIE1ldHJpYyAob3IgYW55IGZpbHRlcikgY2hhbmdlZCBidXQgdGhlIHBsYXRmb3JtIOKAlCBhbmQgdGhlcmVmb3JlIHRoZSBhdmFpbGFibGUgbWV0cmljIGxpc3Qg4oCUIGRpZG4ndDogbm8gbmVlZCB0byByZS1mZXRjaCBtZXRyaWMtb3B0aW9ucyBvciByZWJ1aWxkIHRoZSBzaGVsbCwganVzdCByZWZyZXNoIHRoZSBkYXRhLiAqLwogIGFzeW5jIGZ1bmN0aW9uIHJlZnJlc2hGb3JNZXRyaWMoKSB7CiAgICBjb25zdCBmaWx0ZXJzID0gU3RhdGUuZ2V0RmlsdGVycygpOwogICAgY29uc3QgW3N1bW1hcnksIGZvbGxvd2Vyc10gPSBhd2FpdCBQcm9taXNlLmFsbChbCiAgICAgIEFwaS5tZXRyaWNTdW1tYXJ5KHsgLi4uZmlsdGVycywgbWV0cmljIH0pLAogICAgICBBcGkuZm9sbG93ZXJzS3BpcyhmaWx0ZXJzKSwKICAgIF0pOwogICAgcmVuZGVyS3BpcyhzdW1tYXJ5LCBmb2xsb3dlcnMpOwogICAgYXdhaXQgUHJvbWlzZS5hbGwoWwogICAgICByZW5kZXJQbGF0Zm9ybVN1bW1hcnlCYXIoZmlsdGVycyksCiAgICAgIHJlbmRlclRyZW5kKGZpbHRlcnMpLCByZW5kZXJCcmVha2Rvd24oZmlsdGVycyksIHJlbmRlckNvbnRlbnRUeXBlQnJlYWtkb3duKGZpbHRlcnMpLCByZW5kZXJUb3BQb3N0cyhmaWx0ZXJzKSwKICAgIF0pOwogIH0KCiAgZnVuY3Rpb24gc2hvd1NrZWxldG9ucygpIHsKICAgIGNvbnN0IHN1bW1hcnlCYXIgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncGxhdGZvcm1TdW1tYXJ5QmFyJyk7CiAgICBpZiAoc3VtbWFyeUJhcikgeyBzdW1tYXJ5QmFyLmlubmVySFRNTCA9ICcnOyBzdW1tYXJ5QmFyLmFwcGVuZENoaWxkKHNrZWxldG9uU3RhdEdyaWQoNCkpOyB9CiAgICBjb25zdCBrcGlHcmlkID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2twaUdyaWQnKTsKICAgIGlmIChrcGlHcmlkKSB7IGtwaUdyaWQuaW5uZXJIVE1MID0gJyc7IGtwaUdyaWQuYXBwZW5kQ2hpbGQoc2tlbGV0b25TdGF0R3JpZCg4KSk7IH0KICAgIFsndHJlbmRDaGFydFdyYXAnLCAnYnJlYWtkb3duQ2hhcnRXcmFwJywgJ2NvbnRlbnRUeXBlQ2hhcnRXcmFwJ10uZm9yRWFjaCgoaWQpID0+IHsKICAgICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGlkKTsKICAgICAgaWYgKHdyYXApIHsgd3JhcC5pbm5lckhUTUwgPSAnJzsgd3JhcC5hcHBlbmRDaGlsZChza2VsZXRvbkNoYXJ0KCkpOyB9CiAgICB9KTsKICAgIGNvbnN0IHRvcFBvc3RzID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3RvcFBvc3RzVGFibGUnKTsKICAgIGlmICh0b3BQb3N0cykgeyB0b3BQb3N0cy5pbm5lckhUTUwgPSAnJzsgdG9wUG9zdHMuYXBwZW5kQ2hpbGQoc2tlbGV0b25Sb3dzKDYpKTsgfQogIH0KCiAgYXN5bmMgZnVuY3Rpb24gcmVuZGVyKCkgewogICAgcm9vdCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd2aWV3LWRhc2hib2FyZCcpOwogICAgY29uc3QgZmlsdGVycyA9IFN0YXRlLmdldEZpbHRlcnMoKTsKICAgIGNvbnN0IHsgb3B0aW9ucyB9ID0gYXdhaXQgQXBpLm1ldHJpY09wdGlvbnMoZmlsdGVycy5wbGF0Zm9ybSk7CiAgICBtZXRyaWNPcHRpb25zID0gb3B0aW9uczsKICAgIGlmICghbWV0cmljT3B0aW9ucy5zb21lKChtKSA9PiBtLmtleSA9PT0gbWV0cmljKSkgewogICAgICBtZXRyaWMgPSBtZXRyaWNPcHRpb25zLmxlbmd0aCA/IG1ldHJpY09wdGlvbnNbMF0ua2V5IDogJ3ZpZXdzJzsKICAgIH0KICAgIHNoZWxsKCk7CiAgICBzaG93U2tlbGV0b25zKCk7CiAgICBhd2FpdCByZWZyZXNoRm9yTWV0cmljKCk7CiAgfQoKICByZXR1cm4geyByZW5kZXIgfTsKfSkoKTsKCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBEYXRhIFJlY29yZHMgdGFiOiBhIENSTS1zdHlsZSwgcGxhdGZvcm0tZ3JvdXBlZCBicm93c2VyIGJhY2tlZAogICBieSBwb3N0cy9wb3N0X21ldHJpY3MgKHRoZSBzYW1lIG5vcm1hbGl6ZWQgZGF0YSB0aGUgZGFzaGJvYXJkLAogICBjb21wYXJpc29ucywgYW5kIHJlcG9ydHMgcmVhZCkg4oCUICJBbGwgUGxhdGZvcm1zIiBzaG93cyBhIGNvbW1vbgogICBjcm9zcy1wbGF0Zm9ybSBzdW1tYXJ5LCBhIHNwZWNpZmljIHBsYXRmb3JtIHNob3dzIG9ubHkgdGhhdAogICBwbGF0Zm9ybSdzIGN1cmF0ZWQgbWV0cmljcy4gRXZlcnkgZmllbGQgb2YgYSByZWNvcmQgKGV4YWN0bHkgYXMKICAgaW1wb3J0ZWQpIGlzIGFsd2F5cyByZWFjaGFibGUgdmlhIFZpZXcvRWRpdCByZWdhcmRsZXNzIG9mIHRoZQogICB0YWJsZSdzIGN1cmF0aW9uLCB3aGljaCByZWFkcyB0aGUgcmF3X3Jvd3MgbWlycm9yIGFuZCwgb24gc2F2ZSwKICAgcmUtc3luY3MgcG9zdHMvcG9zdF9tZXRyaWNzIHNvIGV2ZXJ5IHZpZXcgcmVmbGVjdHMgdGhlIGNoYW5nZQogICBpbW1lZGlhdGVseS4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmNvbnN0IFJlY29yZHMgPSAoKCkgPT4gewogIGxldCByb290OwogIGxldCBwYWdlID0gMTsKICBjb25zdCBwYWdlU2l6ZSA9IDI1OwogIGxldCBzZWFyY2hWYWx1ZSA9ICcnOwogIGxldCBzZWFyY2hEZWJvdW5jZSA9IG51bGw7CiAgbGV0IG1vZGFsU3RhdGUgPSBudWxsOyAvLyB7IHJlY29yZCwgdmFsdWVzOiBbLi4uXSB9IOKAlCBFZGl0IG1vZGFsIG9ubHkKICBsZXQgY3VycmVudFJlc3VsdCA9IG51bGw7IC8vIGxhc3QtbG9hZGVkIHBhZ2UsIGtlcHQgc28gc29ydGluZyBjYW4gcmUtcmVuZGVyIHdpdGhvdXQgYSBuZXR3b3JrIHJvdW5kLXRyaXAKICBsZXQgc29ydFN0YXRlID0geyBrZXk6IG51bGwsIGRpcjogJ2FzYycsIHR5cGU6ICdzdHJpbmcnIH07CgogIC8qKiBTb3J0cyBhIGNvcHkgb2YgYHJvd3NgIGJ5IGEgKHBvc3NpYmx5IGRvdHRlZCwgZS5nLiAibWV0cmljcy5yZWFjaCIpIGtleSBwYXRoLiBOdWxscyBhbHdheXMgc29ydCBsYXN0IHJlZ2FyZGxlc3Mgb2YgZGlyZWN0aW9uLiAqLwogIGZ1bmN0aW9uIHNvcnRSb3dzKHJvd3MsIGtleSwgZGlyLCB0eXBlKSB7CiAgICBjb25zdCBmYWN0b3IgPSBkaXIgPT09ICdhc2MnID8gMSA6IC0xOwogICAgY29uc3QgcmVhZCA9IChyb3cpID0+IGtleS5zcGxpdCgnLicpLnJlZHVjZSgobywgaykgPT4gKG8gPT09IG51bGwgfHwgbyA9PT0gdW5kZWZpbmVkID8gdW5kZWZpbmVkIDogb1trXSksIHJvdyk7CiAgICByZXR1cm4gWy4uLnJvd3NdLnNvcnQoKGEsIGIpID0+IHsKICAgICAgY29uc3QgYXYgPSByZWFkKGEpOwogICAgICBjb25zdCBidiA9IHJlYWQoYik7CiAgICAgIGNvbnN0IGFNaXNzaW5nID0gYXYgPT09IG51bGwgfHwgYXYgPT09IHVuZGVmaW5lZCB8fCBhdiA9PT0gJyc7CiAgICAgIGNvbnN0IGJNaXNzaW5nID0gYnYgPT09IG51bGwgfHwgYnYgPT09IHVuZGVmaW5lZCB8fCBidiA9PT0gJyc7CiAgICAgIGlmIChhTWlzc2luZyAmJiBiTWlzc2luZykgcmV0dXJuIDA7CiAgICAgIGlmIChhTWlzc2luZykgcmV0dXJuIDE7CiAgICAgIGlmIChiTWlzc2luZykgcmV0dXJuIC0xOwogICAgICBpZiAodHlwZSA9PT0gJ251bWJlcicpIHJldHVybiAoYXYgLSBidikgKiBmYWN0b3I7CiAgICAgIHJldHVybiBTdHJpbmcoYXYpLmxvY2FsZUNvbXBhcmUoU3RyaW5nKGJ2KSkgKiBmYWN0b3I7CiAgICB9KTsKICB9CgogIC8qKiBBIDx0aD4gdGhhdCB0b2dnbGVzIGFzY2VuZGluZy9kZXNjZW5kaW5nIG9uIGNsaWNrIGFuZCBzaG93cyBhbiBhcnJvdyBvbiB3aGljaGV2ZXIgY29sdW1uIGlzIGFjdGl2ZSDigJQgc29ydHMgdGhlIGFscmVhZHktbG9hZGVkIHBhZ2UgaW5zdGFudGx5LCBubyByZWxvYWQuICovCiAgZnVuY3Rpb24gc29ydGFibGVIZWFkZXIobGFiZWwsIGtleSwgdHlwZSkgewogICAgY29uc3QgdGggPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0aCcpOwogICAgaWYgKHR5cGUgPT09ICdudW1iZXInKSB0aC5jbGFzc05hbWUgPSAnbnVtJzsKICAgIHRoLmNsYXNzTGlzdC5hZGQoJ3NvcnRhYmxlLXRoJyk7CiAgICBjb25zdCBpc0FjdGl2ZSA9IHNvcnRTdGF0ZS5rZXkgPT09IGtleTsKICAgIHRoLmFwcGVuZENoaWxkKGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKGxhYmVsKSk7CiAgICB0aC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3NwYW4nLCBpc0FjdGl2ZSA/IChzb3J0U3RhdGUuZGlyID09PSAnYXNjJyA/ICcg4oaRJyA6ICcg4oaTJykgOiAnIOKGlScsICdzb3J0LWFycm93JykpOwogICAgdGguYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7CiAgICAgIHNvcnRTdGF0ZSA9IHsga2V5LCBkaXI6IHNvcnRTdGF0ZS5rZXkgPT09IGtleSAmJiBzb3J0U3RhdGUuZGlyID09PSAnYXNjJyA/ICdkZXNjJyA6ICdhc2MnLCB0eXBlIH07CiAgICAgIGlmIChjdXJyZW50UmVzdWx0KSByZW5kZXJUYWJsZShjdXJyZW50UmVzdWx0KTsKICAgIH0pOwogICAgcmV0dXJuIHRoOwogIH0KCiAgZnVuY3Rpb24gcGxhdGZvcm1NZXRhKCkgewogICAgcmV0dXJuICh3aW5kb3cuX19maWx0ZXJPcHRpb25zQ2FjaGUgfHwgeyBwbGF0Zm9ybXM6IFtdIH0pLnBsYXRmb3JtczsKICB9CgogIGZ1bmN0aW9uIHBsYXRmb3JtTGFiZWwoaWQpIHsKICAgIGNvbnN0IG0gPSBwbGF0Zm9ybU1ldGEoKS5maW5kKChwKSA9PiBwLmlkID09PSBpZCk7CiAgICByZXR1cm4gbSA/IG0ubGFiZWwgOiBpZDsKICB9CgogIGZ1bmN0aW9uIHNoZWxsKCkgewogICAgcm9vdC5pbm5lckhUTUwgPSAnJzsKICAgIHJvb3QuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCAnRGF0YSBSZWNvcmRzJywgJ3NlY3Rpb24tdGl0bGUnKSk7CiAgICByb290LmFwcGVuZENoaWxkKHRleHRFbCgKICAgICAgJ2RpdicsCiAgICAgICdCcm93c2UgYnkgcGxhdGZvcm0gdG8gc2VlIG9ubHkgaXRzIG1ldHJpY3MsIG9yIHN0YXkgb24gQWxsIFBsYXRmb3JtcyBmb3IgYSBjcm9zcy1wbGF0Zm9ybSBzdW1tYXJ5LiBFdmVyeSByZWNvcmQgc3RheXMgZnVsbHkgZWRpdGFibGUg4oCUIFZpZXcgb3IgRWRpdCBhbHdheXMgb3BlbnMgZXZlcnkgZmllbGQgaW1wb3J0ZWQgZnJvbSB0aGUgc3ByZWFkc2hlZXQsIG5vdCBqdXN0IHdoYXTigJlzIGluIHRoZSB0YWJsZS4nLAogICAgICAnbXV0ZWQnCiAgICApKTsKCiAgICBjb25zdCB0b29sYmFyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0b29sYmFyLmNsYXNzTmFtZSA9ICdyZWNvcmRzLXRvb2xiYXInOwogICAgY29uc3QgcGlsbHMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHBpbGxzLmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1maWx0ZXItcGlsbHMnOwogICAgcGlsbHMuaWQgPSAncmVjb3Jkc1BsYXRmb3JtUGlsbHMnOwogICAgY29uc3Qgc2VhcmNoID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBzZWFyY2guY2xhc3NOYW1lID0gJ3JlY29yZHMtc2VhcmNoJzsKICAgIGNvbnN0IHNlYXJjaElucHV0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsKICAgIHNlYXJjaElucHV0LnR5cGUgPSAnc2VhcmNoJzsKICAgIHNlYXJjaElucHV0LnBsYWNlaG9sZGVyID0gJ1NlYXJjaCBjYXB0aW9ucywgY2FtcGFpZ25zLCBjb250ZW50IHR5cGXigKYnOwogICAgc2VhcmNoSW5wdXQuaWQgPSAncmVjb3Jkc1NlYXJjaElucHV0JzsKICAgIHNlYXJjaElucHV0LnZhbHVlID0gc2VhcmNoVmFsdWU7CiAgICBzZWFyY2hJbnB1dC5hZGRFdmVudExpc3RlbmVyKCdpbnB1dCcsICgpID0+IHsKICAgICAgY2xlYXJUaW1lb3V0KHNlYXJjaERlYm91bmNlKTsKICAgICAgc2VhcmNoRGVib3VuY2UgPSBzZXRUaW1lb3V0KCgpID0+IHsKICAgICAgICBzZWFyY2hWYWx1ZSA9IHNlYXJjaElucHV0LnZhbHVlOwogICAgICAgIHBhZ2UgPSAxOwogICAgICAgIGxvYWQoKTsKICAgICAgfSwgMzAwKTsKICAgIH0pOwogICAgc2VhcmNoLmFwcGVuZENoaWxkKHNlYXJjaElucHV0KTsKICAgIHRvb2xiYXIuYXBwZW5kKHBpbGxzLCBzZWFyY2gpOwogICAgcm9vdC5hcHBlbmRDaGlsZCh0b29sYmFyKTsKCiAgICBjb25zdCBjYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjYXJkLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGNvbnN0IHRhYmxlV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGFibGVXcmFwLmNsYXNzTmFtZSA9ICd0YWJsZS1zY3JvbGwnOwogICAgdGFibGVXcmFwLmlkID0gJ3JlY29yZHNUYWJsZVdyYXAnOwogICAgY2FyZC5hcHBlbmRDaGlsZCh0YWJsZVdyYXApOwogICAgY29uc3QgcGFnZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHBhZ2VyLmNsYXNzTmFtZSA9ICdwYWdpbmF0aW9uLXJvdyc7CiAgICBwYWdlci5pZCA9ICdyZWNvcmRzUGFnZXInOwogICAgY2FyZC5hcHBlbmRDaGlsZChwYWdlcik7CiAgICByb290LmFwcGVuZENoaWxkKGNhcmQpOwoKICAgIHJlbmRlclBpbGxzKCk7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJQaWxscygpIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncmVjb3Jkc1BsYXRmb3JtUGlsbHMnKTsKICAgIGlmICghd3JhcCkgcmV0dXJuOwogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgIGNvbnN0IGN1cnJlbnQgPSBTdGF0ZS5nZXRGaWx0ZXJzKCkucGxhdGZvcm0gfHwgJ2FsbCc7CiAgICBjb25zdCBvcHRpb25zID0gW3sgaWQ6ICdhbGwnLCBsYWJlbDogJ0FsbCBQbGF0Zm9ybXMnLCBjb2xvcjogbnVsbCB9LCAuLi5wbGF0Zm9ybU1ldGEoKV07CiAgICBvcHRpb25zLmZvckVhY2goKG9wdCkgPT4gewogICAgICBjb25zdCBidG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgICAgYnRuLnR5cGUgPSAnYnV0dG9uJzsKICAgICAgYnRuLmNsYXNzTGlzdC50b2dnbGUoJ2lzLWFjdGl2ZScsIGN1cnJlbnQgPT09IG9wdC5pZCk7CiAgICAgIGlmIChvcHQuY29sb3IpIHsKICAgICAgICBjb25zdCBkb3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICAgICAgZG90LmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1kb3QnOwogICAgICAgIGRvdC5zdHlsZS5iYWNrZ3JvdW5kID0gb3B0LmNvbG9yOwogICAgICAgIGJ0bi5hcHBlbmRDaGlsZChkb3QpOwogICAgICB9CiAgICAgIGJ0bi5hcHBlbmRDaGlsZChkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShvcHQubGFiZWwpKTsKICAgICAgYnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gewogICAgICAgIGlmIChjdXJyZW50ID09PSBvcHQuaWQpIHJldHVybjsKICAgICAgICBjb25zdCBmaWx0ZXJTZWxlY3QgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmlsdGVyUGxhdGZvcm0nKTsKICAgICAgICBpZiAoZmlsdGVyU2VsZWN0KSBmaWx0ZXJTZWxlY3QudmFsdWUgPSBvcHQuaWQ7CiAgICAgICAgcGFnZSA9IDE7CiAgICAgICAgU3RhdGUuc2V0RmlsdGVycyh7IHBsYXRmb3JtOiBvcHQuaWQgfSk7CiAgICAgIH0pOwogICAgICB3cmFwLmFwcGVuZENoaWxkKGJ0bik7CiAgICB9KTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIGxvYWQoKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3JlY29yZHNUYWJsZVdyYXAnKTsKICAgIGlmICh3cmFwKSB7IHdyYXAuaW5uZXJIVE1MID0gJyc7IHdyYXAuYXBwZW5kQ2hpbGQoc2tlbGV0b25Sb3dzKDgpKTsgfQogICAgY29uc3QgZmlsdGVycyA9IFN0YXRlLmdldEZpbHRlcnMoKTsKICAgIGNvbnN0IHJlc3VsdCA9IGF3YWl0IEFwaS5yZWNvcmRzVGFibGUoeyAuLi5maWx0ZXJzLCBzZWFyY2g6IHNlYXJjaFZhbHVlLCBwYWdlLCBwYWdlU2l6ZSB9KTsKICAgIHJlbmRlclRhYmxlKHJlc3VsdCk7CiAgICByZW5kZXJQYWdlcihyZXN1bHQpOwogIH0KCiAgZnVuY3Rpb24gY29sdW1uTGFiZWxzRm9yKHJlY29yZCkgewogICAgcmV0dXJuIHJlY29yZC5oZWFkZXJzICYmIHJlY29yZC5oZWFkZXJzLmxlbmd0aAogICAgICA/IHJlY29yZC5oZWFkZXJzLm1hcCgoaCkgPT4gKGggJiYgaC50cmltKCkgPyBoIDogJyh1bmxhYmVsZWQgY29sdW1uKScpKQogICAgICA6IHJlY29yZC52YWx1ZXMubWFwKChfLCBpKSA9PiBgQ29sdW1uICR7aSArIDF9YCk7CiAgfQoKICAvKiogR3JvdXBzIGEgcmF3IHJlY29yZCdzIGZpZWxkcyBieSB0aGUgcXVhbGlmaWVkIGhlYWRlcidzIHBsYXRmb3JtLWdyb3VwIHByZWZpeAogICAgICAoZS5nLiAiRkFDRUJPT0sg4oCUIFZpZXdzIiksIHNvIHRoZSBWaWV3L0VkaXQgcG9wdXAgcmVhZHMgYXMgc2VjdGlvbnMgaW5zdGVhZAogICAgICBvZiBvbmUgbG9uZyBmbGF0IGxpc3Qg4oCUIGZhbGxzIGJhY2sgdG8gYSBzaW5nbGUgIkRldGFpbHMiIHNlY3Rpb24gZm9yCiAgICAgIGlkZW50aWZpZXIgY29sdW1ucyBhbmQgZm9yIHRoZSBzaW1wbGUgKG9uZS1wbGF0Zm9ybS1wZXItcm93KSBmb3JtYXQuICovCiAgZnVuY3Rpb24gZ3JvdXBGaWVsZFJvd3MobGFiZWxzLCB2YWx1ZXMpIHsKICAgIGNvbnN0IGdyb3VwcyA9IFtdOwogICAgY29uc3QgaW5kZXggPSBuZXcgTWFwKCk7CiAgICBsYWJlbHMuZm9yRWFjaCgobGFiZWwsIGlkeCkgPT4gewogICAgICBjb25zdCBzZXBJZHggPSBsYWJlbC5pbmRleE9mKCcg4oCUICcpOwogICAgICBjb25zdCBncm91cE5hbWUgPSBzZXBJZHggPj0gMCA/IGxhYmVsLnNsaWNlKDAsIHNlcElkeCkgOiAnRGV0YWlscyc7CiAgICAgIGNvbnN0IGZpZWxkTGFiZWwgPSBzZXBJZHggPj0gMCA/IGxhYmVsLnNsaWNlKHNlcElkeCArIDMpIDogbGFiZWw7CiAgICAgIGlmICghaW5kZXguaGFzKGdyb3VwTmFtZSkpIHsKICAgICAgICBpbmRleC5zZXQoZ3JvdXBOYW1lLCB7IGdyb3VwOiBncm91cE5hbWUsIGZpZWxkczogW10gfSk7CiAgICAgICAgZ3JvdXBzLnB1c2goaW5kZXguZ2V0KGdyb3VwTmFtZSkpOwogICAgICB9CiAgICAgIGluZGV4LmdldChncm91cE5hbWUpLmZpZWxkcy5wdXNoKHsgaWR4LCBsYWJlbDogZmllbGRMYWJlbCB8fCBgQ29sdW1uICR7aWR4ICsgMX1gLCB2YWx1ZTogdmFsdWVzW2lkeF0gfSk7CiAgICB9KTsKICAgIHJldHVybiBncm91cHM7CiAgfQoKICBmdW5jdGlvbiBwbGF0Zm9ybUJhZGdlcyhpZHMpIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHdyYXAuc3R5bGUuZGlzcGxheSA9ICdmbGV4JzsKICAgIHdyYXAuc3R5bGUuZmxleFdyYXAgPSAnd3JhcCc7CiAgICB3cmFwLnN0eWxlLmdhcCA9ICc0cHgnOwogICAgaWYgKCFpZHMubGVuZ3RoKSByZXR1cm4gdGV4dEVsKCdzcGFuJywgJ+KAlCcsICdtdXRlZCcpOwogICAgY29uc3QgbWV0YSA9IHBsYXRmb3JtTWV0YSgpOwogICAgaWRzLmZvckVhY2goKGlkKSA9PiB7CiAgICAgIGNvbnN0IG0gPSBtZXRhLmZpbmQoKHApID0+IHAuaWQgPT09IGlkKSB8fCB7IGxhYmVsOiBpZCwgY29sb3I6ICcjOTk5JyB9OwogICAgICBjb25zdCBwaWxsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICBwaWxsLmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1waWxsJzsKICAgICAgY29uc3QgZG90ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICBkb3QuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLWRvdCc7CiAgICAgIGRvdC5zdHlsZS5iYWNrZ3JvdW5kID0gbS5jb2xvcjsKICAgICAgcGlsbC5hcHBlbmQoZG90LCBkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShtLmxhYmVsKSk7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQocGlsbCk7CiAgICB9KTsKICAgIHJldHVybiB3cmFwOwogIH0KCiAgZnVuY3Rpb24gc3RhdHVzUGlsbChzdGF0dXMpIHsKICAgIGNvbnN0IHNwYW4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICBzcGFuLmNsYXNzTmFtZSA9IGBzdGF0dXMtcGlsbCAke3N0YXR1c31gOwogICAgc3Bhbi50ZXh0Q29udGVudCA9IHN0YXR1cyA9PT0gJ2VkaXRlZCcgPyAnRWRpdGVkJyA6ICdPcmlnaW5hbCc7CiAgICByZXR1cm4gc3BhbjsKICB9CgogIGZ1bmN0aW9uIG1ldHJpY0NlbGwoa2V5LCB2YWx1ZSkgewogICAgaWYgKGtleSA9PT0gJ3Bvc3RpbmdfbGluaycpIHsKICAgICAgY29uc3QgdGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICB0ZC5jbGFzc05hbWUgPSAnbGluay1jZWxsJzsKICAgICAgaWYgKHZhbHVlKSB7CiAgICAgICAgY29uc3QgYSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2EnKTsKICAgICAgICBhLmhyZWYgPSB2YWx1ZTsKICAgICAgICBhLnRhcmdldCA9ICdfYmxhbmsnOwogICAgICAgIGEucmVsID0gJ25vb3BlbmVyIG5vcmVmZXJyZXInOwogICAgICAgIGEudGV4dENvbnRlbnQgPSAnT3BlbiDihpcnOwogICAgICAgIHRkLmFwcGVuZENoaWxkKGEpOwogICAgICB9IGVsc2UgewogICAgICAgIHRkLmFwcGVuZENoaWxkKGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKCfigJQnKSk7CiAgICAgIH0KICAgICAgcmV0dXJuIHRkOwogICAgfQogICAgY29uc3QgZGlzcGxheSA9IGtleSA9PT0gJ3dhdGNoX3RpbWVfc2Vjb25kcycgPyBGb3JtYXQuZHVyYXRpb24odmFsdWUpIDogRm9ybWF0Lm51bWJlcih2YWx1ZSk7CiAgICByZXR1cm4gdGV4dEVsKCd0ZCcsIGRpc3BsYXksICdudW0nKTsKICB9CgogIGZ1bmN0aW9uIGFjdGlvbkJ1dHRvbnMocm93LCBwbGF0Zm9ybSkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgd3JhcC5jbGFzc05hbWUgPSAncm93LWFjdGlvbnMnOwogICAgY29uc3Qgdmlld0J0biA9IGljb25CdG4oJ2J0bicsICdleWUnLCAnVmlldycpOwogICAgdmlld0J0bi5kaXNhYmxlZCA9ICFyb3cucmF3Um93SWQ7CiAgICB2aWV3QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gb3BlblZpZXcocm93LnJhd1Jvd0lkKSk7CiAgICBjb25zdCBlZGl0QnRuID0gaWNvbkJ0bignYnRuJywgJ3BlbmNpbCcsICdFZGl0Jyk7CiAgICBlZGl0QnRuLmRpc2FibGVkID0gIXJvdy5yYXdSb3dJZDsKICAgIGVkaXRCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBvcGVuRWRpdG9yKHJvdy5yYXdSb3dJZCkpOwogICAgY29uc3QgZGVsZXRlQnRuID0gaWNvbkJ0bignYnRuIGRhbmdlcicsICd0cmFzaC0yJywgJ0RlbGV0ZScpOwogICAgZGVsZXRlQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gaGFuZGxlRGVsZXRlKHJvdywgcGxhdGZvcm0pKTsKICAgIHdyYXAuYXBwZW5kKHZpZXdCdG4sIGVkaXRCdG4sIGRlbGV0ZUJ0bik7CiAgICByZXR1cm4gd3JhcDsKICB9CgogIGZ1bmN0aW9uIGNhcHRpb25DZWxsKGNhcHRpb24pIHsKICAgIGNvbnN0IHRleHQgPSBjYXB0aW9uIHx8ICcobm8gY2FwdGlvbiknOwogICAgcmV0dXJuIHRleHRFbCgndGQnLCB0ZXh0Lmxlbmd0aCA+IDcwID8gYCR7dGV4dC5zbGljZSgwLCA3MCl94oCmYCA6IHRleHQpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyU3VtbWFyeVRhYmxlKHJlc3VsdCkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdyZWNvcmRzVGFibGVXcmFwJyk7CiAgICBpZiAoIXJlc3VsdC5yb3dzLmxlbmd0aCkgewogICAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgICB3cmFwLmFwcGVuZENoaWxkKGVtcHR5U3RhdGUoewogICAgICAgIGljb246ICdkYXRhYmFzZScsCiAgICAgICAgdGl0bGU6ICdObyByZWNvcmRzIG1hdGNoIHRoZXNlIGZpbHRlcnMgeWV0JywKICAgICAgICBtZXNzYWdlOiAnVXBsb2FkIGEgd2Vla2x5IGV4cG9ydCwgb3Igd2lkZW4gdGhlIGRhdGUgcmFuZ2UsIHRvIHNlZSByZWNvcmRzIGhlcmUuJywKICAgICAgICBhY3Rpb25MYWJlbDogJ1VwbG9hZCBkYXRhJywKICAgICAgICBvbkFjdGlvbjogKCkgPT4gZG9jdW1lbnQucXVlcnlTZWxlY3RvcignLnRhYi1idG5bZGF0YS10YWI9InVwbG9hZCJdJyk/LmNsaWNrKCksCiAgICAgIH0pKTsKICAgICAgcmV0dXJuOwogICAgfQogICAgY29uc3QgdGFibGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0YWJsZScpOwogICAgdGFibGUuY2xhc3NOYW1lID0gJ2RhdGEtdGFibGUnOwogICAgY29uc3QgdGhlYWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0aGVhZCcpOwogICAgY29uc3QgaGVhZFRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgIGhlYWRUci5hcHBlbmQoCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdEYXRlJywgJ3B1Ymxpc2hEYXRlJywgJ3N0cmluZycpLAogICAgICBzb3J0YWJsZUhlYWRlcignUGxhdGZvcm1zJywgJ3BsYXRmb3JtSWRzLjAnLCAnc3RyaW5nJyksCiAgICAgIHRleHRFbCgndGgnLCAnQ2FwdGlvbicpLAogICAgICB0ZXh0RWwoJ3RoJywgJ0NhbXBhaWduJyksCiAgICAgIHRleHRFbCgndGgnLCAnQ29udGVudCBUeXBlJyksCiAgICAgIHRleHRFbCgndGgnLCAnU3RhdHVzJyksCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdMYXN0IFVwZGF0ZWQnLCAndXBkYXRlZEF0JywgJ3N0cmluZycpLAogICAgICB0ZXh0RWwoJ3RoJywgJ0FjdGlvbnMnKQogICAgKTsKICAgIHRoZWFkLmFwcGVuZENoaWxkKGhlYWRUcik7CiAgICBjb25zdCB0Ym9keSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3Rib2R5Jyk7CiAgICBjb25zdCByb3dzID0gc29ydFN0YXRlLmtleSA/IHNvcnRSb3dzKHJlc3VsdC5yb3dzLCBzb3J0U3RhdGUua2V5LCBzb3J0U3RhdGUuZGlyLCBzb3J0U3RhdGUudHlwZSkgOiByZXN1bHQucm93czsKICAgIHJvd3MuZm9yRWFjaCgocikgPT4gewogICAgICBjb25zdCB0ciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAgIGNvbnN0IHBsYXRmb3Jtc1RkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgcGxhdGZvcm1zVGQuYXBwZW5kQ2hpbGQocGxhdGZvcm1CYWRnZXMoci5wbGF0Zm9ybUlkcykpOwogICAgICBjb25zdCBzdGF0dXNUZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIHN0YXR1c1RkLmFwcGVuZENoaWxkKHN0YXR1c1BpbGwoci5zdGF0dXMpKTsKICAgICAgY29uc3QgYWN0aW9uc1RkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGQnKTsKICAgICAgYWN0aW9uc1RkLmFwcGVuZENoaWxkKGFjdGlvbkJ1dHRvbnMociwgJ2FsbCcpKTsKICAgICAgdHIuYXBwZW5kKAogICAgICAgIHRleHRFbCgndGQnLCBGb3JtYXQuZGF0ZShyLnB1Ymxpc2hEYXRlKSksCiAgICAgICAgcGxhdGZvcm1zVGQsCiAgICAgICAgY2FwdGlvbkNlbGwoci5jYXB0aW9uKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgci5jYW1wYWlnblR5cGUgfHwgJ+KAlCcpLAogICAgICAgIHRleHRFbCgndGQnLCByLmNvbnRlbnRUeXBlIHx8ICfigJQnKSwKICAgICAgICBzdGF0dXNUZCwKICAgICAgICB0ZXh0RWwoJ3RkJywgci51cGRhdGVkQXQpLAogICAgICAgIGFjdGlvbnNUZAogICAgICApOwogICAgICB0Ym9keS5hcHBlbmRDaGlsZCh0cik7CiAgICB9KTsKICAgIHRhYmxlLmFwcGVuZCh0aGVhZCwgdGJvZHkpOwogICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGFibGUpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyUGxhdGZvcm1UYWJsZShyZXN1bHQpIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncmVjb3Jkc1RhYmxlV3JhcCcpOwogICAgaWYgKCFyZXN1bHQucm93cy5sZW5ndGgpIHsKICAgICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgICAgd3JhcC5hcHBlbmRDaGlsZChlbXB0eVN0YXRlKHsKICAgICAgICBpY29uOiAnZGF0YWJhc2UnLAogICAgICAgIHRpdGxlOiBgTm8gJHtwbGF0Zm9ybUxhYmVsKHJlc3VsdC5wbGF0Zm9ybSl9IHJlY29yZHMgbWF0Y2ggdGhlc2UgZmlsdGVycyB5ZXRgLAogICAgICAgIG1lc3NhZ2U6ICdUcnkgYSBkaWZmZXJlbnQgcGxhdGZvcm0sIG9yIHdpZGVuIHRoZSBkYXRlIHJhbmdlLicsCiAgICAgIH0pKTsKICAgICAgcmV0dXJuOwogICAgfQogICAgY29uc3QgdGFibGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0YWJsZScpOwogICAgdGFibGUuY2xhc3NOYW1lID0gJ2RhdGEtdGFibGUnOwogICAgY29uc3QgdGhlYWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0aGVhZCcpOwogICAgY29uc3QgaGVhZFRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgIGhlYWRUci5hcHBlbmQoc29ydGFibGVIZWFkZXIoJ0RhdGUnLCAncHVibGlzaERhdGUnLCAnc3RyaW5nJyksIHRleHRFbCgndGgnLCAnQ2FwdGlvbicpLCB0ZXh0RWwoJ3RoJywgJ0NhbXBhaWduJyksIHRleHRFbCgndGgnLCAnQ29udGVudCBUeXBlJykpOwogICAgcmVzdWx0LmNvbHVtbnMuZm9yRWFjaCgoYykgPT4gewogICAgICBpZiAoYy5rZXkgPT09ICdwb3N0aW5nX2xpbmsnKSB7CiAgICAgICAgaGVhZFRyLmFwcGVuZENoaWxkKHRleHRFbCgndGgnLCBjLmxhYmVsKSk7CiAgICAgIH0gZWxzZSB7CiAgICAgICAgaGVhZFRyLmFwcGVuZENoaWxkKHNvcnRhYmxlSGVhZGVyKGMubGFiZWwsIGBtZXRyaWNzLiR7Yy5rZXl9YCwgJ251bWJlcicpKTsKICAgICAgfQogICAgfSk7CiAgICBoZWFkVHIuYXBwZW5kKHRleHRFbCgndGgnLCAnU3RhdHVzJyksIHRleHRFbCgndGgnLCAnQWN0aW9ucycpKTsKICAgIHRoZWFkLmFwcGVuZENoaWxkKGhlYWRUcik7CiAgICBjb25zdCB0Ym9keSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3Rib2R5Jyk7CiAgICBjb25zdCByb3dzID0gc29ydFN0YXRlLmtleSA/IHNvcnRSb3dzKHJlc3VsdC5yb3dzLCBzb3J0U3RhdGUua2V5LCBzb3J0U3RhdGUuZGlyLCBzb3J0U3RhdGUudHlwZSkgOiByZXN1bHQucm93czsKICAgIHJvd3MuZm9yRWFjaCgocikgPT4gewogICAgICBjb25zdCB0ciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAgIHRyLmFwcGVuZCh0ZXh0RWwoJ3RkJywgRm9ybWF0LmRhdGUoci5wdWJsaXNoRGF0ZSkpLCBjYXB0aW9uQ2VsbChyLmNhcHRpb24pLCB0ZXh0RWwoJ3RkJywgci5jYW1wYWlnblR5cGUgfHwgJ+KAlCcpLCB0ZXh0RWwoJ3RkJywgci5jb250ZW50VHlwZSB8fCAn4oCUJykpOwogICAgICByZXN1bHQuY29sdW1ucy5mb3JFYWNoKChjKSA9PiB0ci5hcHBlbmRDaGlsZChtZXRyaWNDZWxsKGMua2V5LCByLm1ldHJpY3NbYy5rZXldKSkpOwogICAgICBjb25zdCBzdGF0dXNUZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIHN0YXR1c1RkLmFwcGVuZENoaWxkKHN0YXR1c1BpbGwoci5zdGF0dXMpKTsKICAgICAgdHIuYXBwZW5kQ2hpbGQoc3RhdHVzVGQpOwogICAgICBjb25zdCBhY3Rpb25zVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBhY3Rpb25zVGQuYXBwZW5kQ2hpbGQoYWN0aW9uQnV0dG9ucyhyLCByZXN1bHQucGxhdGZvcm0pKTsKICAgICAgdHIuYXBwZW5kQ2hpbGQoYWN0aW9uc1RkKTsKICAgICAgdGJvZHkuYXBwZW5kQ2hpbGQodHIpOwogICAgfSk7CiAgICB0YWJsZS5hcHBlbmQodGhlYWQsIHRib2R5KTsKICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CiAgICB3cmFwLmFwcGVuZENoaWxkKHRhYmxlKTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlclRhYmxlKHJlc3VsdCkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdyZWNvcmRzVGFibGVXcmFwJyk7CiAgICBpZiAoIXdyYXApIHJldHVybjsKICAgIGN1cnJlbnRSZXN1bHQgPSByZXN1bHQ7CiAgICBpZiAocmVzdWx0LnBsYXRmb3JtID09PSAnYWxsJykgcmVuZGVyU3VtbWFyeVRhYmxlKHJlc3VsdCk7CiAgICBlbHNlIHJlbmRlclBsYXRmb3JtVGFibGUocmVzdWx0KTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlclBhZ2VyKHJlc3VsdCkgewogICAgY29uc3QgcGFnZXIgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncmVjb3Jkc1BhZ2VyJyk7CiAgICBpZiAoIXBhZ2VyKSByZXR1cm47CiAgICBwYWdlci5pbm5lckhUTUwgPSAnJzsKICAgIGNvbnN0IHRvdGFsUGFnZXMgPSBNYXRoLm1heCgxLCBNYXRoLmNlaWwocmVzdWx0LnRvdGFsIC8gcmVzdWx0LnBhZ2VTaXplKSk7CiAgICBjb25zdCBwcmV2QnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBwcmV2QnRuLmNsYXNzTmFtZSA9ICdidG4nOwogICAgcHJldkJ0bi50ZXh0Q29udGVudCA9ICdQcmV2aW91cyc7CiAgICBwcmV2QnRuLmRpc2FibGVkID0gcmVzdWx0LnBhZ2UgPD0gMTsKICAgIHByZXZCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7IHBhZ2UgLT0gMTsgbG9hZCgpOyB9KTsKICAgIGNvbnN0IG5leHRCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIG5leHRCdG4uY2xhc3NOYW1lID0gJ2J0bic7CiAgICBuZXh0QnRuLnRleHRDb250ZW50ID0gJ05leHQnOwogICAgbmV4dEJ0bi5kaXNhYmxlZCA9IHJlc3VsdC5wYWdlID49IHRvdGFsUGFnZXM7CiAgICBuZXh0QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4geyBwYWdlICs9IDE7IGxvYWQoKTsgfSk7CiAgICBwYWdlci5hcHBlbmQocHJldkJ0biwgdGV4dEVsKCdzcGFuJywgYFBhZ2UgJHtyZXN1bHQucGFnZX0gb2YgJHt0b3RhbFBhZ2VzfSDigJQgJHtyZXN1bHQudG90YWx9IHJlY29yZChzKWApLCBuZXh0QnRuKTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIGhhbmRsZURlbGV0ZShyb3csIHBsYXRmb3JtKSB7CiAgICBjb25zdCBjYXB0aW9uID0gKHJvdy5jYXB0aW9uIHx8ICcobm8gY2FwdGlvbiknKS5zbGljZSgwLCA2MCk7CiAgICBjb25zdCBtZXNzYWdlID0gcGxhdGZvcm0gPT09ICdhbGwnCiAgICAgID8gYERlbGV0ZSB0aGlzIGVudGlyZSByZWNvcmQg4oCUICIke2NhcHRpb259IiDigJQgYWNyb3NzIGV2ZXJ5IHBsYXRmb3JtPyBJdHMgb3JpZ2luYWwgaW1wb3J0IHN0YXlzIGluIFVwbG9hZCBIaXN0b3J5LCBidXQgaXQgd2lsbCBkaXNhcHBlYXIgZnJvbSB0aGUgZGFzaGJvYXJkLCBjb21wYXJpc29ucywgYW5kIHJlcG9ydHMuYAogICAgICA6IGBSZW1vdmUgdGhpcyByZWNvcmQncyAke3BsYXRmb3JtTGFiZWwocGxhdGZvcm0pfSBkYXRhIOKAlCAiJHtjYXB0aW9ufSI/IElmIHRoaXMgaXMgaXRzIG9ubHkgcGxhdGZvcm0sIHRoZSB3aG9sZSByZWNvcmQgd2lsbCBiZSByZW1vdmVkIGZyb20gdGhlIGRhc2hib2FyZC5gOwogICAgaWYgKCF3aW5kb3cuY29uZmlybShtZXNzYWdlKSkgcmV0dXJuOwogICAgdHJ5IHsKICAgICAgaWYgKHBsYXRmb3JtID09PSAnYWxsJykgYXdhaXQgQXBpLmRlbGV0ZVJlY29yZFBvc3Qocm93LnBvc3RJZCk7CiAgICAgIGVsc2UgYXdhaXQgQXBpLmRlbGV0ZVJlY29yZFBsYXRmb3JtKHJvdy5wb3N0SWQsIHBsYXRmb3JtKTsKICAgICAgVG9hc3Quc2hvdygnUmVjb3JkIGRlbGV0ZWQuJywgJ3N1Y2Nlc3MnKTsKICAgICAgYXdhaXQgbG9hZCgpOwogICAgICB3aW5kb3cuZGlzcGF0Y2hFdmVudChuZXcgQ3VzdG9tRXZlbnQoJ2xyczpkYXRhLXVwZGF0ZWQnKSk7CiAgICB9IGNhdGNoIChlcnIpIHsKICAgICAgVG9hc3Quc2hvdyhlcnIubWVzc2FnZSwgJ2Vycm9yJyk7CiAgICB9CiAgfQoKICBmdW5jdGlvbiByZW1vdmVFeGlzdGluZ092ZXJsYXkoKSB7CiAgICBjb25zdCBvdmVybGF5ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3JlY29yZE1vZGFsT3ZlcmxheScpOwogICAgaWYgKG92ZXJsYXkpIG92ZXJsYXkucmVtb3ZlKCk7CiAgfQoKICBmdW5jdGlvbiBjbG9zZU1vZGFsKCkgewogICAgcmVtb3ZlRXhpc3RpbmdPdmVybGF5KCk7CiAgICBtb2RhbFN0YXRlID0gbnVsbDsKICB9CgogIC8vIE9ubHkgY2xlYXJzIHRoZSBzdGFsZSBET00gbm9kZSDigJQgTk9UIG1vZGFsU3RhdGUuIHJlbmRlckVkaXRNb2RhbCByZWFkcwogIC8vIG1vZGFsU3RhdGUgcmlnaHQgYWZ0ZXIgY2FsbGluZyB0aGlzIHRvIGJ1aWxkIHRoZSBmb3JtOyBpZiB0aGlzIGNhbGxlZAogIC8vIHRoZSByZWFsIGNsb3NlTW9kYWwoKSAoYXMgaXQgdXNlZCB0byksIHRoYXQgcmVzZXQgbW9kYWxTdGF0ZSB0byBudWxsIG91dAogIC8vIGZyb20gdW5kZXIgaXQgYmVmb3JlIHRoZSByZWFkLCB3aGljaCBpcyBleGFjdGx5IHdoeSBFZGl0IHdhcyBicm9rZW4uCiAgZnVuY3Rpb24gbW9kYWxTaGVsbCh0aXRsZVRleHQpIHsKICAgIHJlbW92ZUV4aXN0aW5nT3ZlcmxheSgpOwogICAgY29uc3Qgb3ZlcmxheSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgb3ZlcmxheS5jbGFzc05hbWUgPSAnbW9kYWwtb3ZlcmxheSc7CiAgICBvdmVybGF5LmlkID0gJ3JlY29yZE1vZGFsT3ZlcmxheSc7CiAgICBvdmVybGF5LmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKGUpID0+IHsgaWYgKGUudGFyZ2V0ID09PSBvdmVybGF5KSBjbG9zZU1vZGFsKCk7IH0pOwogICAgY29uc3QgcGFuZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHBhbmVsLmNsYXNzTmFtZSA9ICdtb2RhbC1wYW5lbCB3aWRlJzsKICAgIHBhbmVsLmFwcGVuZENoaWxkKHRleHRFbCgnaDInLCB0aXRsZVRleHQpKTsKICAgIG92ZXJsYXkuYXBwZW5kQ2hpbGQocGFuZWwpOwogICAgcmV0dXJuIHsgb3ZlcmxheSwgcGFuZWwgfTsKICB9CgogIGZ1bmN0aW9uIHJlY29yZFN1YnRpdGxlKHIpIHsKICAgIHJldHVybiBgU2hlZXQgIiR7ci5zaGVldE5hbWV9Iiwgcm93ICR7ci5yb3dOdW1iZXJ9JHtyLnBvc3RJZCA/IGAg4oCUIGxpbmtlZCB0byBkYXNoYm9hcmQgcG9zdCAjJHtyLnBvc3RJZH1gIDogJyDigJQgbm90IHBhcnQgb2YgdGhlIGRhc2hib2FyZCAoZS5nLiBuZWVkcyBhIHZhbGlkIGRhdGUpJ31gOwogIH0KCiAgLy8gLS0tLS0tLS0tLSBWaWV3IHBvcHVwOiByZWFkLW9ubHksIGV2ZXJ5IGZpZWxkLCBncm91cGVkIGludG8gc2VjdGlvbnMgLS0tLS0tLS0tLQogIGFzeW5jIGZ1bmN0aW9uIG9wZW5WaWV3KGlkKSB7CiAgICBjb25zdCByZWNvcmQgPSBhd2FpdCBBcGkuZ2V0UmVjb3JkKGlkKTsKICAgIGNvbnN0IHsgb3ZlcmxheSwgcGFuZWwgfSA9IG1vZGFsU2hlbGwoJ1JlY29yZCBkZXRhaWxzJyk7CiAgICBwYW5lbC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIHJlY29yZFN1YnRpdGxlKHJlY29yZCksICdtb2RhbC1zdWInKSk7CgogICAgY29uc3QgZ3JvdXBzID0gZ3JvdXBGaWVsZFJvd3MoY29sdW1uTGFiZWxzRm9yKHJlY29yZCksIHJlY29yZC52YWx1ZXMpOwogICAgZ3JvdXBzLmZvckVhY2goKGcpID0+IHsKICAgICAgY29uc3Qgc2VjdGlvbiA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBzZWN0aW9uLmNsYXNzTmFtZSA9ICdyZWNvcmQtc2VjdGlvbic7CiAgICAgIHNlY3Rpb24uYXBwZW5kQ2hpbGQodGV4dEVsKCdoNCcsIGcuZ3JvdXApKTsKICAgICAgY29uc3QgZ3JpZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBncmlkLmNsYXNzTmFtZSA9ICdmb3JtLWdyaWQnOwogICAgICBnLmZpZWxkcy5mb3JFYWNoKChmKSA9PiB7CiAgICAgICAgY29uc3QgZmllbGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgICBmaWVsZC5jbGFzc05hbWUgPSAndmlldy1maWVsZCc7CiAgICAgICAgZmllbGQuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBmLmxhYmVsLCAndmlldy1sYWJlbCcpKTsKICAgICAgICBjb25zdCB2YWwgPSBmLnZhbHVlID09PSB1bmRlZmluZWQgfHwgZi52YWx1ZSA9PT0gbnVsbCB8fCBmLnZhbHVlID09PSAnJyA/ICfigJQnIDogU3RyaW5nKGYudmFsdWUpOwogICAgICAgIGZpZWxkLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgdmFsLCAndmlldy12YWx1ZScpKTsKICAgICAgICBncmlkLmFwcGVuZENoaWxkKGZpZWxkKTsKICAgICAgfSk7CiAgICAgIHNlY3Rpb24uYXBwZW5kQ2hpbGQoZ3JpZCk7CiAgICAgIHBhbmVsLmFwcGVuZENoaWxkKHNlY3Rpb24pOwogICAgfSk7CgogICAgY29uc3QgYWN0aW9ucyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYWN0aW9ucy5jbGFzc05hbWUgPSAnbW9kYWwtYWN0aW9ucyc7CiAgICBjb25zdCBidG5Sb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGJ0blJvdy5jbGFzc05hbWUgPSAnYnRuLXJvdyc7CiAgICBjb25zdCBjbG9zZUJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgY2xvc2VCdG4uY2xhc3NOYW1lID0gJ2J0bic7CiAgICBjbG9zZUJ0bi50ZXh0Q29udGVudCA9ICdDbG9zZSc7CiAgICBjbG9zZUJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIGNsb3NlTW9kYWwpOwogICAgY29uc3QgZWRpdEJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgZWRpdEJ0bi5jbGFzc05hbWUgPSAnYnRuIHByaW1hcnknOwogICAgZWRpdEJ0bi50ZXh0Q29udGVudCA9ICdFZGl0IHRoaXMgcmVjb3JkJzsKICAgIGVkaXRCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBvcGVuRWRpdG9yKHJlY29yZC5pZCkpOwogICAgYnRuUm93LmFwcGVuZChjbG9zZUJ0biwgZWRpdEJ0bik7CiAgICBhY3Rpb25zLmFwcGVuZENoaWxkKGJ0blJvdyk7CiAgICBwYW5lbC5hcHBlbmRDaGlsZChhY3Rpb25zKTsKCiAgICBkb2N1bWVudC5ib2R5LmFwcGVuZENoaWxkKG92ZXJsYXkpOwogIH0KCiAgLy8gLS0tLS0tLS0tLSBFZGl0IHBvcHVwOiBldmVyeSBmaWVsZCwgZ3JvdXBlZCBpbnRvIHNlY3Rpb25zLCBhbGwgZWRpdGFibGUgLS0tLS0tLS0tLQogIGFzeW5jIGZ1bmN0aW9uIG9wZW5FZGl0b3IoaWQpIHsKICAgIGNvbnN0IHJlY29yZCA9IGF3YWl0IEFwaS5nZXRSZWNvcmQoaWQpOwogICAgbW9kYWxTdGF0ZSA9IHsgcmVjb3JkLCB2YWx1ZXM6IFsuLi5yZWNvcmQudmFsdWVzXSB9OwogICAgcmVuZGVyRWRpdE1vZGFsKCk7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJFZGl0TW9kYWwoKSB7CiAgICBjb25zdCByID0gbW9kYWxTdGF0ZS5yZWNvcmQ7CiAgICBjb25zdCB7IG92ZXJsYXksIHBhbmVsIH0gPSBtb2RhbFNoZWxsKCdFZGl0IHJlY29yZCcpOwogICAgcGFuZWwuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCByZWNvcmRTdWJ0aXRsZShyKSwgJ21vZGFsLXN1YicpKTsKCiAgICBjb25zdCBncm91cHMgPSBncm91cEZpZWxkUm93cyhjb2x1bW5MYWJlbHNGb3IociksIG1vZGFsU3RhdGUudmFsdWVzKTsKICAgIGdyb3Vwcy5mb3JFYWNoKChnKSA9PiB7CiAgICAgIGNvbnN0IHNlY3Rpb24gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgc2VjdGlvbi5jbGFzc05hbWUgPSAncmVjb3JkLXNlY3Rpb24nOwogICAgICBzZWN0aW9uLmFwcGVuZENoaWxkKHRleHRFbCgnaDQnLCBnLmdyb3VwKSk7CiAgICAgIGNvbnN0IGdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgZ3JpZC5jbGFzc05hbWUgPSAnZm9ybS1ncmlkJzsKICAgICAgZy5maWVsZHMuZm9yRWFjaCgoZikgPT4gewogICAgICAgIGNvbnN0IGZpZWxkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgICAgZmllbGQuY2xhc3NOYW1lID0gJ2Zvcm0tZmllbGQnOwogICAgICAgIGZpZWxkLmFwcGVuZENoaWxkKHRleHRFbCgnbGFiZWwnLCBmLmxhYmVsKSk7CiAgICAgICAgY29uc3Qgc3RyVmFsID0gZi52YWx1ZSA9PT0gdW5kZWZpbmVkIHx8IGYudmFsdWUgPT09IG51bGwgPyAnJyA6IFN0cmluZyhmLnZhbHVlKTsKICAgICAgICBjb25zdCBpc0xvbmcgPSBzdHJWYWwubGVuZ3RoID4gODA7CiAgICAgICAgY29uc3QgaW5wdXQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KGlzTG9uZyA/ICd0ZXh0YXJlYScgOiAnaW5wdXQnKTsKICAgICAgICBpZiAoIWlzTG9uZykgaW5wdXQudHlwZSA9ICd0ZXh0JzsKICAgICAgICBlbHNlIGZpZWxkLnN0eWxlLmdyaWRDb2x1bW4gPSAnMSAvIC0xJzsKICAgICAgICBpbnB1dC52YWx1ZSA9IHN0clZhbDsKICAgICAgICBpbnB1dC5hZGRFdmVudExpc3RlbmVyKCdpbnB1dCcsICgpID0+IHsgbW9kYWxTdGF0ZS52YWx1ZXNbZi5pZHhdID0gaW5wdXQudmFsdWU7IH0pOwogICAgICAgIGZpZWxkLmFwcGVuZENoaWxkKGlucHV0KTsKICAgICAgICBncmlkLmFwcGVuZENoaWxkKGZpZWxkKTsKICAgICAgfSk7CiAgICAgIHNlY3Rpb24uYXBwZW5kQ2hpbGQoZ3JpZCk7CiAgICAgIHBhbmVsLmFwcGVuZENoaWxkKHNlY3Rpb24pOwogICAgfSk7CgogICAgY29uc3QgYWN0aW9ucyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYWN0aW9ucy5jbGFzc05hbWUgPSAnbW9kYWwtYWN0aW9ucyc7CiAgICBjb25zdCBlcnJvck1zZyA9IHRleHRFbCgnc3BhbicsICcnLCAnbXV0ZWQnKTsKICAgIGVycm9yTXNnLmlkID0gJ21vZGFsRXJyb3JNc2cnOwogICAgY29uc3QgYnRuUm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBidG5Sb3cuY2xhc3NOYW1lID0gJ2J0bi1yb3cnOwogICAgY29uc3QgY2FuY2VsQnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICBjYW5jZWxCdG4uY2xhc3NOYW1lID0gJ2J0bic7CiAgICBjYW5jZWxCdG4udGV4dENvbnRlbnQgPSAnQ2FuY2VsJzsKICAgIGNhbmNlbEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIGNsb3NlTW9kYWwpOwogICAgY29uc3Qgc2F2ZUJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgc2F2ZUJ0bi5jbGFzc05hbWUgPSAnYnRuIHByaW1hcnknOwogICAgc2F2ZUJ0bi50ZXh0Q29udGVudCA9ICdTYXZlIGNoYW5nZXMnOwogICAgc2F2ZUJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHNhdmVFZGl0KHNhdmVCdG4pKTsKICAgIGJ0blJvdy5hcHBlbmQoY2FuY2VsQnRuLCBzYXZlQnRuKTsKICAgIGFjdGlvbnMuYXBwZW5kKGVycm9yTXNnLCBidG5Sb3cpOwogICAgcGFuZWwuYXBwZW5kQ2hpbGQoYWN0aW9ucyk7CgogICAgZG9jdW1lbnQuYm9keS5hcHBlbmRDaGlsZChvdmVybGF5KTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIHNhdmVFZGl0KGJ0bikgewogICAgY29uc3QgZXJyb3JFbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdtb2RhbEVycm9yTXNnJyk7CiAgICBlcnJvckVsLnRleHRDb250ZW50ID0gJyc7CiAgICBidG4uZGlzYWJsZWQgPSB0cnVlOwogICAgYnRuLnRleHRDb250ZW50ID0gJ1NhdmluZ+KApic7CiAgICB0cnkgewogICAgICBhd2FpdCBBcGkudXBkYXRlUmVjb3JkKG1vZGFsU3RhdGUucmVjb3JkLmlkLCBtb2RhbFN0YXRlLnZhbHVlcyk7CiAgICAgIFRvYXN0LnNob3coJ1JlY29yZCB1cGRhdGVkLicsICdzdWNjZXNzJyk7CiAgICAgIGNsb3NlTW9kYWwoKTsKICAgICAgYXdhaXQgbG9hZCgpOwogICAgICB3aW5kb3cuZGlzcGF0Y2hFdmVudChuZXcgQ3VzdG9tRXZlbnQoJ2xyczpkYXRhLXVwZGF0ZWQnKSk7CiAgICB9IGNhdGNoIChlcnIpIHsKICAgICAgZXJyb3JFbC50ZXh0Q29udGVudCA9IGVyci5tZXNzYWdlOwogICAgICBlcnJvckVsLnN0eWxlLmNvbG9yID0gJ3ZhcigtLXN0YXR1cy1jcml0aWNhbCknOwogICAgICBidG4uZGlzYWJsZWQgPSBmYWxzZTsKICAgICAgYnRuLnRleHRDb250ZW50ID0gJ1NhdmUgY2hhbmdlcyc7CiAgICB9CiAgfQoKICBhc3luYyBmdW5jdGlvbiByZW5kZXIoKSB7CiAgICByb290ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3ZpZXctcmVjb3JkcycpOwogICAgcGFnZSA9IDE7CiAgICBzaGVsbCgpOwogICAgYXdhaXQgbG9hZCgpOwogIH0KCiAgcmV0dXJuIHsgcmVuZGVyLCByZWxvYWQ6IGxvYWQsIG9wZW5WaWV3IH07Cn0pKCk7CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgQ29tcGFyaXNvbnMgdGFiOiB3ZWVrLXZzLXdlZWssIGN1c3RvbSByYW5nZSwgbW9udGhseSwKICAgcXVhcnRlcmx5LCBZVEQuCiAgID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwpjb25zdCBDb21wYXJpc29uID0gKCgpID0+IHsKICBjb25zdCBNT0RFUyA9IFsKICAgIHsga2V5OiAncGxhdGZvcm1zJywgbGFiZWw6ICdBbGwgUGxhdGZvcm1zJyB9LAogICAgeyBrZXk6ICd3ZWVrJywgbGFiZWw6ICdXZWVrIHZzIFdlZWsnIH0sCiAgICB7IGtleTogJ2N1c3RvbScsIGxhYmVsOiAnQ3VzdG9tIFJhbmdlJyB9LAogICAgeyBrZXk6ICdtb250aCcsIGxhYmVsOiAnTW9udGhseScgfSwKICAgIHsga2V5OiAncXVhcnRlcicsIGxhYmVsOiAnUXVhcnRlcmx5JyB9LAogICAgeyBrZXk6ICd5dGQnLCBsYWJlbDogJ1llYXIgdG8gRGF0ZScgfSwKICBdOwogIGNvbnN0IE1FVFJJQ19ST1dTID0gWwogICAgeyBrZXk6ICd2aWV3cycsIGxhYmVsOiAnVmlld3MnIH0sCiAgICB7IGtleTogJ3JlYWNoJywgbGFiZWw6ICdSZWFjaCcgfSwKICAgIHsga2V5OiAnaW1wcmVzc2lvbnMnLCBsYWJlbDogJ0ltcHJlc3Npb25zJyB9LAogICAgeyBrZXk6ICdlbmdhZ2VtZW50JywgbGFiZWw6ICdFbmdhZ2VtZW50JyB9LAogICAgeyBrZXk6ICdjbGlja3MnLCBsYWJlbDogJ0NsaWNrcycgfSwKICAgIHsga2V5OiAnZm9sbG93ZXJzX2dhaW5lZCcsIGxhYmVsOiAnRm9sbG93ZXJzIEdhaW5lZCcgfSwKICAgIHsga2V5OiAnd2F0Y2hfdGltZV9zZWNvbmRzJywgbGFiZWw6ICdXYXRjaCBUaW1lJyB9LAogICAgeyBrZXk6ICdzaGFyZXMnLCBsYWJlbDogJ1NoYXJlcycgfSwKICAgIHsga2V5OiAnY29tbWVudHMnLCBsYWJlbDogJ0NvbW1lbnRzJyB9LAogICAgeyBrZXk6ICdzYXZlcycsIGxhYmVsOiAnU2F2ZXMnIH0sCiAgXTsKCiAgbGV0IG1vZGUgPSAncGxhdGZvcm1zJzsKICBsZXQgcm9vdDsKICBsZXQgcGxhdGZvcm1DaGFydE1ldHJpYyA9ICdlbmdhZ2VtZW50JzsKICBsZXQgY2FyZFNvcnRNb2RlID0gJ292ZXJhbGwnOwogIGxldCBjYXJkUGxhdGZvcm1GaWx0ZXIgPSAnYWxsJzsKCiAgZnVuY3Rpb24gbW9uZGF5T2YoZGF0ZVN0cikgewogICAgY29uc3QgZCA9IG5ldyBEYXRlKGRhdGVTdHIpOwogICAgY29uc3QgZGF5ID0gZC5nZXREYXkoKTsKICAgIGNvbnN0IGRpZmYgPSBkYXkgPT09IDAgPyA2IDogZGF5IC0gMTsKICAgIGQuc2V0RGF0ZShkLmdldERhdGUoKSAtIGRpZmYpOwogICAgcmV0dXJuIGQudG9JU09TdHJpbmcoKS5zbGljZSgwLCAxMCk7CiAgfQogIGZ1bmN0aW9uIGFkZERheXMoZGF0ZVN0ciwgbikgewogICAgY29uc3QgZCA9IG5ldyBEYXRlKGRhdGVTdHIpOwogICAgZC5zZXREYXRlKGQuZ2V0RGF0ZSgpICsgbik7CiAgICByZXR1cm4gZC50b0lTT1N0cmluZygpLnNsaWNlKDAsIDEwKTsKICB9CgogIGZ1bmN0aW9uIHNoZWxsKCkgewogICAgcm9vdC5pbm5lckhUTUwgPSAnJzsKCiAgICBjb25zdCB0YWJzID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB0YWJzLmNsYXNzTmFtZSA9ICdtb2RlLXRhYnMnOwogICAgTU9ERVMuZm9yRWFjaCgobSkgPT4gewogICAgICBjb25zdCBidG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgICAgYnRuLnRleHRDb250ZW50ID0gbS5sYWJlbDsKICAgICAgYnRuLnR5cGUgPSAnYnV0dG9uJzsKICAgICAgaWYgKG0ua2V5ID09PSBtb2RlKSBidG4uY2xhc3NMaXN0LmFkZCgnaXMtYWN0aXZlJyk7CiAgICAgIGJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsKICAgICAgICBtb2RlID0gbS5rZXk7CiAgICAgICAgc2hlbGwoKTsKICAgICAgfSk7CiAgICAgIHRhYnMuYXBwZW5kQ2hpbGQoYnRuKTsKICAgIH0pOwogICAgcm9vdC5hcHBlbmRDaGlsZCh0YWJzKTsKCiAgICBjb25zdCBjb250cm9scyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY29udHJvbHMuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgY29udHJvbHMuaWQgPSAnY29tcGFyaXNvbkNvbnRyb2xzJzsKICAgIHJvb3QuYXBwZW5kQ2hpbGQoY29udHJvbHMpOwoKICAgIGNvbnN0IHJlc3VsdHMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHJlc3VsdHMuaWQgPSAnY29tcGFyaXNvblJlc3VsdHMnOwogICAgcm9vdC5hcHBlbmRDaGlsZChyZXN1bHRzKTsKCiAgICByZW5kZXJDb250cm9scygpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyQ29udHJvbHMoKSB7CiAgICBjb25zdCBjb250cm9scyA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdjb21wYXJpc29uQ29udHJvbHMnKTsKICAgIGNvbnRyb2xzLmlubmVySFRNTCA9ICcnOwogICAgY29uc3Qgcm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICByb3cuY2xhc3NOYW1lID0gJ2J0bi1yb3cnOwogICAgcm93LnN0eWxlLmFsaWduSXRlbXMgPSAnZW5kJzsKCiAgICBjb25zdCB0b2RheSA9IG5ldyBEYXRlKCkudG9JU09TdHJpbmcoKS5zbGljZSgwLCAxMCk7CiAgICBjb25zdCB0aGlzWWVhciA9IG5ldyBEYXRlKCkuZ2V0RnVsbFllYXIoKTsKCiAgICBpZiAobW9kZSA9PT0gJ3BsYXRmb3JtcycpIHsKICAgICAgY29uc3QgZkZyb20gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOyBmRnJvbS50eXBlID0gJ2RhdGUnOyBmRnJvbS5pZCA9ICdwbGF0Zm9ybVJlcG9ydEZyb20nOwogICAgICBjb25zdCBmVG8gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOyBmVG8udHlwZSA9ICdkYXRlJzsgZlRvLmlkID0gJ3BsYXRmb3JtUmVwb3J0VG8nOwogICAgICBjb25zdCBhcHBseUJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgICBhcHBseUJ0bi5jbGFzc05hbWUgPSAnYnRuIHByaW1hcnknOwogICAgICBhcHBseUJ0bi50eXBlID0gJ2J1dHRvbic7CiAgICAgIGFwcGx5QnRuLnRleHRDb250ZW50ID0gJ0FwcGx5IFJhbmdlJzsKICAgICAgYXBwbHlCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBsb2FkUGxhdGZvcm1SZXBvcnQoeyBkYXRlRnJvbTogZkZyb20udmFsdWUsIGRhdGVUbzogZlRvLnZhbHVlIH0pKTsKICAgICAgY29uc3QgY2xlYXJCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgICAgY2xlYXJCdG4uY2xhc3NOYW1lID0gJ2J0bic7CiAgICAgIGNsZWFyQnRuLnR5cGUgPSAnYnV0dG9uJzsKICAgICAgY2xlYXJCdG4udGV4dENvbnRlbnQgPSAnQWxsIFRpbWUnOwogICAgICBjbGVhckJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsgZkZyb20udmFsdWUgPSAnJzsgZlRvLnZhbHVlID0gJyc7IGxvYWRQbGF0Zm9ybVJlcG9ydCh7fSk7IH0pOwogICAgICByb3cuYXBwZW5kKAogICAgICAgIGxhYmVsZWQoJ0Zyb20gKG9wdGlvbmFsKScsIGZGcm9tKSwKICAgICAgICBsYWJlbGVkKCdUbyAob3B0aW9uYWwpJywgZlRvKSwKICAgICAgICBhcHBseUJ0biwKICAgICAgICBjbGVhckJ0bgogICAgICApOwogICAgICBjb250cm9scy5hcHBlbmRDaGlsZChyb3cpOwogICAgICBsb2FkUGxhdGZvcm1SZXBvcnQoe30pOwogICAgICByZXR1cm47CiAgICB9IGVsc2UgaWYgKG1vZGUgPT09ICd3ZWVrJykgewogICAgICBjb25zdCB3QSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7IHdBLnR5cGUgPSAnZGF0ZSc7IHdBLnZhbHVlID0gbW9uZGF5T2YodG9kYXkpOwogICAgICBjb25zdCB3QiA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7IHdCLnR5cGUgPSAnZGF0ZSc7IHdCLnZhbHVlID0gbW9uZGF5T2YoYWRkRGF5cyh0b2RheSwgLTcpKTsKICAgICAgcm93LmFwcGVuZChsYWJlbGVkKCdXZWVrIEEgKGFueSBkYXkgaW4gd2VlayknLCB3QSksIGxhYmVsZWQoJ1dlZWsgQiAoYW55IGRheSBpbiB3ZWVrKScsIHdCKSwgcnVuQnRuKCgpID0+IHsKICAgICAgICBjb25zdCByYW5nZUEgPSB7IGZyb206IG1vbmRheU9mKHdBLnZhbHVlKSwgdG86IGFkZERheXMobW9uZGF5T2Yod0EudmFsdWUpLCA2KSB9OwogICAgICAgIGNvbnN0IHJhbmdlQiA9IHsgZnJvbTogbW9uZGF5T2Yod0IudmFsdWUpLCB0bzogYWRkRGF5cyhtb25kYXlPZih3Qi52YWx1ZSksIDYpIH07CiAgICAgICAgcnVuQ29tcGFyZShyYW5nZUEsIHJhbmdlQiwgYFdlZWsgb2YgJHtGb3JtYXQuZGF0ZShyYW5nZUEuZnJvbSl9YCwgYFdlZWsgb2YgJHtGb3JtYXQuZGF0ZShyYW5nZUIuZnJvbSl9YCk7CiAgICAgIH0pKTsKICAgIH0gZWxzZSBpZiAobW9kZSA9PT0gJ2N1c3RvbScpIHsKICAgICAgY29uc3QgZkEgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOyBmQS50eXBlID0gJ2RhdGUnOyBmQS52YWx1ZSA9IGFkZERheXModG9kYXksIC0xMyk7CiAgICAgIGNvbnN0IHRBID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsgdEEudHlwZSA9ICdkYXRlJzsgdEEudmFsdWUgPSB0b2RheTsKICAgICAgY29uc3QgZkIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOyBmQi50eXBlID0gJ2RhdGUnOyBmQi52YWx1ZSA9IGFkZERheXModG9kYXksIC0yNyk7CiAgICAgIGNvbnN0IHRCID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsgdEIudHlwZSA9ICdkYXRlJzsgdEIudmFsdWUgPSBhZGREYXlzKHRvZGF5LCAtMTQpOwogICAgICByb3cuYXBwZW5kKAogICAgICAgIGxhYmVsZWQoJ1JhbmdlIEEgZnJvbScsIGZBKSwgbGFiZWxlZCgndG8nLCB0QSksCiAgICAgICAgbGFiZWxlZCgnUmFuZ2UgQiBmcm9tJywgZkIpLCBsYWJlbGVkKCd0bycsIHRCKSwKICAgICAgICBydW5CdG4oKCkgPT4gcnVuQ29tcGFyZSh7IGZyb206IGZBLnZhbHVlLCB0bzogdEEudmFsdWUgfSwgeyBmcm9tOiBmQi52YWx1ZSwgdG86IHRCLnZhbHVlIH0sICdSYW5nZSBBJywgJ1JhbmdlIEInKSkKICAgICAgKTsKICAgIH0gZWxzZSBpZiAobW9kZSA9PT0gJ21vbnRoJykgewogICAgICBjb25zdCB5ID0geWVhclNlbGVjdCh0aGlzWWVhcik7IGNvbnN0IG0gPSBtb250aFNlbGVjdChuZXcgRGF0ZSgpLmdldE1vbnRoKCkgKyAxKTsKICAgICAgY29uc3QgdG9nZ2xlID0gcGVyaW9kVG9nZ2xlKCk7CiAgICAgIHJvdy5hcHBlbmQobGFiZWxlZCgnWWVhcicsIHkpLCBsYWJlbGVkKCdNb250aCcsIG0pLCB0b2dnbGUuZWwsIHJ1bkJ0bihhc3luYyAoKSA9PiB7CiAgICAgICAgY29uc3QgcmVwb3J0ID0gYXdhaXQgQXBpLm1vbnRobHkoeyB5ZWFyOiB5LnZhbHVlLCBtb250aDogbS52YWx1ZSwgLi4uU3RhdGUuZ2V0RmlsdGVycygpIH0pOwogICAgICAgIHJlbmRlclBlcmlvZFJlcG9ydChyZXBvcnQsIHRvZ2dsZS5nZXQoKSk7CiAgICAgIH0pKTsKICAgIH0gZWxzZSBpZiAobW9kZSA9PT0gJ3F1YXJ0ZXInKSB7CiAgICAgIGNvbnN0IHkgPSB5ZWFyU2VsZWN0KHRoaXNZZWFyKTsgY29uc3QgcSA9IHF1YXJ0ZXJTZWxlY3QoKTsKICAgICAgY29uc3QgdG9nZ2xlID0gcGVyaW9kVG9nZ2xlKCk7CiAgICAgIHJvdy5hcHBlbmQobGFiZWxlZCgnWWVhcicsIHkpLCBsYWJlbGVkKCdRdWFydGVyJywgcSksIHRvZ2dsZS5lbCwgcnVuQnRuKGFzeW5jICgpID0+IHsKICAgICAgICBjb25zdCByZXBvcnQgPSBhd2FpdCBBcGkucXVhcnRlcmx5KHsgeWVhcjogeS52YWx1ZSwgcXVhcnRlcjogcS52YWx1ZSwgLi4uU3RhdGUuZ2V0RmlsdGVycygpIH0pOwogICAgICAgIHJlbmRlclBlcmlvZFJlcG9ydChyZXBvcnQsIHRvZ2dsZS5nZXQoKSk7CiAgICAgIH0pKTsKICAgIH0gZWxzZSBpZiAobW9kZSA9PT0gJ3l0ZCcpIHsKICAgICAgY29uc3QgeSA9IHllYXJTZWxlY3QodGhpc1llYXIpOwogICAgICByb3cuYXBwZW5kKGxhYmVsZWQoJ1llYXInLCB5KSwgcnVuQnRuKGFzeW5jICgpID0+IHsKICAgICAgICBjb25zdCByZXBvcnQgPSBhd2FpdCBBcGkueXRkKHsgeWVhcjogeS52YWx1ZSwgLi4uU3RhdGUuZ2V0RmlsdGVycygpIH0pOwogICAgICAgIHJlbmRlclBlcmlvZFJlcG9ydChyZXBvcnQsICd2c0xhc3RZZWFyJyk7CiAgICAgIH0pKTsKICAgIH0KCiAgICBjb250cm9scy5hcHBlbmRDaGlsZChyb3cpOwogIH0KCiAgZnVuY3Rpb24gbGFiZWxlZChsYWJlbCwgZWwpIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHdyYXAuY2xhc3NOYW1lID0gJ2ZpZWxkLWlubGluZSc7CiAgICB3cmFwLmFwcGVuZCh0ZXh0RWwoJ2xhYmVsJywgbGFiZWwpLCBlbCk7CiAgICByZXR1cm4gd3JhcDsKICB9CiAgZnVuY3Rpb24gcnVuQnRuKG9uQ2xpY2spIHsKICAgIGNvbnN0IGJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgYnRuLmNsYXNzTmFtZSA9ICdidG4gcHJpbWFyeSc7CiAgICBidG4udGV4dENvbnRlbnQgPSAnQ29tcGFyZSc7CiAgICBidG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBvbkNsaWNrKCkpOwogICAgcmV0dXJuIGJ0bjsKICB9CiAgZnVuY3Rpb24geWVhclNlbGVjdChkZWZhdWx0WWVhcikgewogICAgY29uc3Qgc2VsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc2VsZWN0Jyk7CiAgICBmb3IgKGxldCB5ID0gZGVmYXVsdFllYXIgLSAzOyB5IDw9IGRlZmF1bHRZZWFyICsgMTsgeSArPSAxKSB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOyBvcHQudmFsdWUgPSB5OyBvcHQudGV4dENvbnRlbnQgPSB5OwogICAgICBpZiAoeSA9PT0gZGVmYXVsdFllYXIpIG9wdC5zZWxlY3RlZCA9IHRydWU7CiAgICAgIHNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfQogICAgcmV0dXJuIHNlbDsKICB9CiAgZnVuY3Rpb24gbW9udGhTZWxlY3QoZGVmYXVsdE1vbnRoKSB7CiAgICBjb25zdCBuYW1lcyA9IFsnSmFudWFyeScsJ0ZlYnJ1YXJ5JywnTWFyY2gnLCdBcHJpbCcsJ01heScsJ0p1bmUnLCdKdWx5JywnQXVndXN0JywnU2VwdGVtYmVyJywnT2N0b2JlcicsJ05vdmVtYmVyJywnRGVjZW1iZXInXTsKICAgIGNvbnN0IHNlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NlbGVjdCcpOwogICAgbmFtZXMuZm9yRWFjaCgobiwgaSkgPT4gewogICAgICBjb25zdCBvcHQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsgb3B0LnZhbHVlID0gaSArIDE7IG9wdC50ZXh0Q29udGVudCA9IG47CiAgICAgIGlmIChpICsgMSA9PT0gZGVmYXVsdE1vbnRoKSBvcHQuc2VsZWN0ZWQgPSB0cnVlOwogICAgICBzZWwuYXBwZW5kQ2hpbGQob3B0KTsKICAgIH0pOwogICAgcmV0dXJuIHNlbDsKICB9CiAgZnVuY3Rpb24gcXVhcnRlclNlbGVjdCgpIHsKICAgIGNvbnN0IGN1cnJlbnRRID0gTWF0aC5mbG9vcihuZXcgRGF0ZSgpLmdldE1vbnRoKCkgLyAzKSArIDE7CiAgICBjb25zdCBzZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzZWxlY3QnKTsKICAgIFsxLCAyLCAzLCA0XS5mb3JFYWNoKChxKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOyBvcHQudmFsdWUgPSBxOyBvcHQudGV4dENvbnRlbnQgPSBgUSR7cX1gOwogICAgICBpZiAocSA9PT0gY3VycmVudFEpIG9wdC5zZWxlY3RlZCA9IHRydWU7CiAgICAgIHNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CiAgICByZXR1cm4gc2VsOwogIH0KICBmdW5jdGlvbiBwZXJpb2RUb2dnbGUoKSB7CiAgICBjb25zdCBzZWwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzZWxlY3QnKTsKICAgIFtbJ3ZzUHJldmlvdXNQZXJpb2QnLCAndnMgUHJldmlvdXMgUGVyaW9kJ10sIFsndnNMYXN0WWVhcicsICd2cyBTYW1lIFBlcmlvZCBMYXN0IFllYXInXV0uZm9yRWFjaCgoW3YsIGxdKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOyBvcHQudmFsdWUgPSB2OyBvcHQudGV4dENvbnRlbnQgPSBsOwogICAgICBzZWwuYXBwZW5kQ2hpbGQob3B0KTsKICAgIH0pOwogICAgcmV0dXJuIHsgZWw6IGxhYmVsZWQoJ0NvbXBhcmUnLCBzZWwpLCBnZXQ6ICgpID0+IHNlbC52YWx1ZSB9OwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gcnVuQ29tcGFyZShyYW5nZUEsIHJhbmdlQiwgbGFiZWxBLCBsYWJlbEIpIHsKICAgIGNvbnN0IGZpbHRlcnMgPSBTdGF0ZS5nZXRGaWx0ZXJzKCk7CiAgICBjb25zdCByZXN1bHQgPSBhd2FpdCBBcGkuY29tcGFyZSh7CiAgICAgIGZyb21BOiByYW5nZUEuZnJvbSwgdG9BOiByYW5nZUEudG8sIGZyb21COiByYW5nZUIuZnJvbSwgdG9COiByYW5nZUIudG8sCiAgICAgIHBsYXRmb3JtOiBmaWx0ZXJzLnBsYXRmb3JtLCBjYW1wYWlnblR5cGU6IGZpbHRlcnMuY2FtcGFpZ25UeXBlLCBjb250ZW50VHlwZTogZmlsdGVycy5jb250ZW50VHlwZSwKICAgIH0pOwogICAgcmVuZGVyQ29tcGFyZVJlc3VsdChyZXN1bHQsIGxhYmVsQSwgbGFiZWxCKTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlclBlcmlvZFJlcG9ydChyZXBvcnQsIHdoaWNoKSB7CiAgICBjb25zdCBjbXAgPSByZXBvcnRbd2hpY2hdOwogICAgY29uc3QgbGFiZWxBID0gJ0N1cnJlbnQgcGVyaW9kJzsKICAgIGNvbnN0IGxhYmVsQiA9IHdoaWNoID09PSAndnNMYXN0WWVhcicgPyAnU2FtZSBwZXJpb2QgbGFzdCB5ZWFyJyA6ICdQcmV2aW91cyBwZXJpb2QnOwogICAgcmVuZGVyQ29tcGFyZVJlc3VsdChjbXAsIGxhYmVsQSwgbGFiZWxCLCByZXBvcnQucmFuZ2UpOwogIH0KCiAgZnVuY3Rpb24gc3RhdFRpbGUobGFiZWwsIGN1cnJlbnQsIHByZXZpb3VzLCBncm93dGgsIGlzRHVyYXRpb24pIHsKICAgIGNvbnN0IHRpbGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRpbGUuY2xhc3NOYW1lID0gJ3N0YXQtdGlsZSc7CiAgICBjb25zdCBjdXJEaXNwbGF5ID0gaXNEdXJhdGlvbiA/IEZvcm1hdC5kdXJhdGlvbihjdXJyZW50KSA6IEZvcm1hdC5jb21wYWN0KGN1cnJlbnQpOwogICAgY29uc3QgcHJldkRpc3BsYXkgPSBpc0R1cmF0aW9uID8gRm9ybWF0LmR1cmF0aW9uKHByZXZpb3VzKSA6IEZvcm1hdC5jb21wYWN0KHByZXZpb3VzKTsKICAgIHRpbGUuYXBwZW5kKAogICAgICB0ZXh0RWwoJ2RpdicsIGxhYmVsLCAnc3RhdC1sYWJlbCcpLAogICAgICB0ZXh0RWwoJ2RpdicsIGN1ckRpc3BsYXksICdzdGF0LXZhbHVlJyksCiAgICAgIHRleHRFbCgnZGl2JywgYCR7Rm9ybWF0LnBjdChncm93dGgpfSDCtyB3YXMgJHtwcmV2RGlzcGxheX1gLCBgc3RhdC1kZWx0YSAke0Zvcm1hdC5kZWx0YUNsYXNzKGdyb3d0aCl9YCkKICAgICk7CiAgICByZXR1cm4gdGlsZTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlckNvbXBhcmVSZXN1bHQocmVzdWx0LCBsYWJlbEEsIGxhYmVsQiwgaGVhZGxpbmUpIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnY29tcGFyaXNvblJlc3VsdHMnKTsKICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CgogICAgY29uc3QgdGl0bGUgPSB0ZXh0RWwoJ2RpdicsIGhlYWRsaW5lCiAgICAgID8gYCR7Rm9ybWF0LmRhdGUocmVzdWx0LnJhbmdlQS5mcm9tKX0g4oCTICR7Rm9ybWF0LmRhdGUocmVzdWx0LnJhbmdlQS50byl9YAogICAgICA6IGAke2xhYmVsQX06ICR7Rm9ybWF0LmRhdGUocmVzdWx0LnJhbmdlQS5mcm9tKX0g4oCTICR7Rm9ybWF0LmRhdGUocmVzdWx0LnJhbmdlQS50byl9ICB2cyAgJHtsYWJlbEJ9OiAke0Zvcm1hdC5kYXRlKHJlc3VsdC5yYW5nZUIuZnJvbSl9IOKAkyAke0Zvcm1hdC5kYXRlKHJlc3VsdC5yYW5nZUIudG8pfWAsCiAgICAgICdzZWN0aW9uLXRpdGxlJyk7CiAgICB3cmFwLmFwcGVuZENoaWxkKHRpdGxlKTsKCiAgICBjb25zdCBncmlkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBncmlkLmNsYXNzTmFtZSA9ICdzdGF0LWdyaWQnOwogICAgZ3JpZC5hcHBlbmRDaGlsZChzdGF0VGlsZSgnUG9zdHMnLCByZXN1bHQucmFuZ2VBLnRvdGFscy5wb3N0X2NvdW50LCByZXN1bHQucmFuZ2VCLnRvdGFscy5wb3N0X2NvdW50LCByZXN1bHQuZ3Jvd3RoLnBvc3RfY291bnQsIGZhbHNlKSk7CiAgICBNRVRSSUNfUk9XUy5mb3JFYWNoKChtKSA9PiB7CiAgICAgIGdyaWQuYXBwZW5kQ2hpbGQoc3RhdFRpbGUobS5sYWJlbCwgcmVzdWx0LnJhbmdlQS50b3RhbHNbbS5rZXldLCByZXN1bHQucmFuZ2VCLnRvdGFsc1ttLmtleV0sIHJlc3VsdC5ncm93dGhbbS5rZXldLCBtLmtleSA9PT0gJ3dhdGNoX3RpbWVfc2Vjb25kcycpKTsKICAgIH0pOwogICAgd3JhcC5hcHBlbmRDaGlsZChncmlkKTsKCiAgICByZW5kZXJQbGF0Zm9ybUNvbXBhcmlzb25DYXJkcyh3cmFwLCByZXN1bHQsIGxhYmVsQSwgbGFiZWxCKTsKICB9CgogIC8qKgogICAqICJBbGwgUGxhdGZvcm1zIiByZXBvcnQg4oCUIHRoZSBoZWFkbGluZSBDb21wYXJpc29ucyB2aWV3LiBVbmxpa2UgdGhlCiAgICogd2Vlay9jdXN0b20vbW9udGgvcXVhcnRlci95dGQgdG9vbHMgYWJvdmUsIHRoaXMgaWdub3JlcyB0aGUgc2hhcmVkCiAgICogcGxhdGZvcm0vY2FtcGFpZ24vY29udGVudC10eXBlIGZpbHRlciBiYXIgZW50aXJlbHkgYW5kIG5lZWRzIG5vIGRhdGUKICAgKiByYW5nZTogaXQgYWx3YXlzIGNvdmVycyBldmVyeSBwbGF0Zm9ybSB3aXRoIGFueSBkYXRhICh1cGxvYWRlZCBwb3N0cwogICAqIGFuZC9vciBtYW51YWxseS1lbnRlcmVkIEZvbGxvd2VycyBEYXRhIFJlY29yZCBoaXN0b3J5KS4KICAgKi8KICBhc3luYyBmdW5jdGlvbiBsb2FkUGxhdGZvcm1SZXBvcnQocGFyYW1zKSB7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2NvbXBhcmlzb25SZXN1bHRzJyk7CiAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgd3JhcC5hcHBlbmRDaGlsZChza2VsZXRvblN0YXRHcmlkKDIpKTsKICAgIHdyYXAuYXBwZW5kQ2hpbGQoc2tlbGV0b25DaGFydCgpKTsKICAgIGNvbnN0IGhhc0V4cGxpY2l0UmFuZ2UgPSBwYXJhbXMgJiYgcGFyYW1zLmRhdGVGcm9tICYmIHBhcmFtcy5kYXRlVG87CiAgICBjb25zdCByZXBvcnQgPSBhd2FpdCBBcGkucGxhdGZvcm1SZXBvcnQoaGFzRXhwbGljaXRSYW5nZSA/IHBhcmFtcyA6IHt9KTsKICAgIHJlbmRlclBsYXRmb3JtUmVwb3J0KHJlcG9ydCk7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJQbGF0Zm9ybVJlcG9ydChyZXBvcnQpIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnY29tcGFyaXNvblJlc3VsdHMnKTsKICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CgogICAgaWYgKCFyZXBvcnQucGxhdGZvcm1zLmxlbmd0aCkgewogICAgICB3cmFwLmFwcGVuZENoaWxkKGVtcHR5U3RhdGUoewogICAgICAgIGljb246ICdnaXQtY29tcGFyZScsCiAgICAgICAgdGl0bGU6ICdObyBwbGF0Zm9ybSBkYXRhIHlldCcsCiAgICAgICAgbWVzc2FnZTogJ1VwbG9hZCBwb3N0cyBvciBhZGQgRm9sbG93ZXJzIERhdGEgUmVjb3JkIGVudHJpZXMgdG8gc2VlIGEgY3Jvc3MtcGxhdGZvcm0gY29tcGFyaXNvbiBoZXJlLicsCiAgICAgIH0pKTsKICAgICAgcmV0dXJuOwogICAgfQoKICAgIGNvbnN0IHJhbmdlTGFiZWwgPSByZXBvcnQucmFuZ2UuaXNFeHBsaWNpdAogICAgICA/IGAke0Zvcm1hdC5kYXRlKHJlcG9ydC5yYW5nZS5mcm9tKX0g4oCTICR7Rm9ybWF0LmRhdGUocmVwb3J0LnJhbmdlLnRvKX1gCiAgICAgIDogYEFsbCB0aW1lICgke0Zvcm1hdC5kYXRlKHJlcG9ydC5yYW5nZS5mcm9tKX0g4oCTICR7Rm9ybWF0LmRhdGUocmVwb3J0LnJhbmdlLnRvKX0pYDsKICAgIHdyYXAuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBgUGxhdGZvcm0gQ29tcGFyaXNvbiBSZXBvcnQg4oCUICR7cmFuZ2VMYWJlbH1gLCAnc2VjdGlvbi10aXRsZScpKTsKCiAgICBjb25zdCBiZXN0UCA9IHJlcG9ydC5wbGF0Zm9ybXMuZmluZCgocCkgPT4gcC5wbGF0Zm9ybSA9PT0gcmVwb3J0LmJlc3RQbGF0Zm9ybSk7CiAgICBjb25zdCB3b3JzdFAgPSByZXBvcnQucGxhdGZvcm1zLmZpbmQoKHApID0+IHAucGxhdGZvcm0gPT09IHJlcG9ydC53b3JzdFBsYXRmb3JtKTsKICAgIGlmIChiZXN0UCB8fCB3b3JzdFApIHsKICAgICAgY29uc3QgaGlnaGxpZ2h0R3JpZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBoaWdobGlnaHRHcmlkLmNsYXNzTmFtZSA9ICdzdGF0LWdyaWQnOwogICAgICBpZiAoYmVzdFApIHsKICAgICAgICBjb25zdCB0aWxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgICAgdGlsZS5jbGFzc05hbWUgPSAnc3RhdC10aWxlJzsKICAgICAgICB0aWxlLmFwcGVuZCgKICAgICAgICAgIHRleHRFbCgnZGl2JywgJ0Jlc3QtUGVyZm9ybWluZyBQbGF0Zm9ybScsICdzdGF0LWxhYmVsJyksCiAgICAgICAgICB0ZXh0RWwoJ2RpdicsIGJlc3RQLmxhYmVsLCAnc3RhdC12YWx1ZScpLAogICAgICAgICAgdGV4dEVsKCdkaXYnLCBgUmVhY2ggJHtGb3JtYXQuc21hcnQoYmVzdFAudG90YWxzLnJlYWNoKX0gwrcgRW5nYWdlbWVudCAke0Zvcm1hdC5zbWFydChiZXN0UC50b3RhbHMuZW5nYWdlbWVudCl9YCwgJ3Bvc3QtbWV0YScpCiAgICAgICAgKTsKICAgICAgICBoaWdobGlnaHRHcmlkLmFwcGVuZENoaWxkKHRpbGUpOwogICAgICB9CiAgICAgIGlmICh3b3JzdFApIHsKICAgICAgICBjb25zdCB0aWxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgICAgdGlsZS5jbGFzc05hbWUgPSAnc3RhdC10aWxlJzsKICAgICAgICB0aWxlLmFwcGVuZCgKICAgICAgICAgIHRleHRFbCgnZGl2JywgJ0xvd2VzdC1QZXJmb3JtaW5nIFBsYXRmb3JtJywgJ3N0YXQtbGFiZWwnKSwKICAgICAgICAgIHRleHRFbCgnZGl2Jywgd29yc3RQLmxhYmVsLCAnc3RhdC12YWx1ZScpLAogICAgICAgICAgdGV4dEVsKCdkaXYnLCBgUmVhY2ggJHtGb3JtYXQuc21hcnQod29yc3RQLnRvdGFscy5yZWFjaCl9IMK3IEVuZ2FnZW1lbnQgJHtGb3JtYXQuc21hcnQod29yc3RQLnRvdGFscy5lbmdhZ2VtZW50KX1gLCAncG9zdC1tZXRhJykKICAgICAgICApOwogICAgICAgIGhpZ2hsaWdodEdyaWQuYXBwZW5kQ2hpbGQodGlsZSk7CiAgICAgIH0KICAgICAgd3JhcC5hcHBlbmRDaGlsZChoaWdobGlnaHRHcmlkKTsKICAgIH0KCiAgICBjb25zdCB0YWJsZUNhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRhYmxlQ2FyZC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICB0YWJsZUNhcmQuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMycsICdQbGF0Zm9ybSBSYW5raW5nJykpOwogICAgY29uc3QgdGFibGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0YWJsZScpOwogICAgdGFibGUuY2xhc3NOYW1lID0gJ2RhdGEtdGFibGUnOwogICAgY29uc3QgdGhlYWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0aGVhZCcpOwogICAgY29uc3QgaGVhZFJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICBbJ1JhbmsnLCAnUGxhdGZvcm0nLCAnUG9zdHMnLCAnUmVhY2gnLCAnRW5nYWdlbWVudCcsICdJbXByZXNzaW9ucycsICdGb2xsb3dlciBHcm93dGgnXS5mb3JFYWNoKChsYWJlbCwgaSkgPT4gewogICAgICBjb25zdCB0aCA9IHRleHRFbCgndGgnLCBsYWJlbCk7CiAgICAgIGlmIChpID49IDIpIHRoLmNsYXNzTGlzdC5hZGQoJ251bScpOwogICAgICBoZWFkUm93LmFwcGVuZENoaWxkKHRoKTsKICAgIH0pOwogICAgdGhlYWQuYXBwZW5kQ2hpbGQoaGVhZFJvdyk7CiAgICB0YWJsZS5hcHBlbmRDaGlsZCh0aGVhZCk7CiAgICBjb25zdCB0Ym9keSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3Rib2R5Jyk7CiAgICByZXBvcnQucGxhdGZvcm1zLmZvckVhY2goKHApID0+IHsKICAgICAgY29uc3QgdHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgICB0ci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3RkJywgcC5vdmVyYWxsUmFuayA/IGAjJHtwLm92ZXJhbGxSYW5rfWAgOiAn4oCUJykpOwogICAgICBjb25zdCBwbGF0VGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBjb25zdCBwaWxsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOyBwaWxsLmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1waWxsJzsKICAgICAgY29uc3QgZG90ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOyBkb3QuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLWRvdCc7IGRvdC5zdHlsZS5iYWNrZ3JvdW5kID0gcC5jb2xvcjsKICAgICAgcGlsbC5hcHBlbmQoZG90LCBkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShwLmxhYmVsKSk7CiAgICAgIHBsYXRUZC5hcHBlbmRDaGlsZChwaWxsKTsKICAgICAgdHIuYXBwZW5kQ2hpbGQocGxhdFRkKTsKICAgICAgdHIuYXBwZW5kQ2hpbGQodGV4dEVsKCd0ZCcsIEZvcm1hdC5udW1iZXIocC5wb3N0Q291bnQpLCAnbnVtJykpOwogICAgICB0ci5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3RkJywgRm9ybWF0LnNtYXJ0KHAudG90YWxzLnJlYWNoKSwgJ251bScpKTsKICAgICAgdHIuYXBwZW5kQ2hpbGQodGV4dEVsKCd0ZCcsIEZvcm1hdC5zbWFydChwLnRvdGFscy5lbmdhZ2VtZW50KSwgJ251bScpKTsKICAgICAgdHIuYXBwZW5kQ2hpbGQodGV4dEVsKCd0ZCcsIEZvcm1hdC5zbWFydChwLnRvdGFscy5pbXByZXNzaW9ucyksICdudW0nKSk7CiAgICAgIGNvbnN0IGZvbGxvd2VyVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBmb2xsb3dlclRkLmNsYXNzTmFtZSA9ICdudW0nOwogICAgICBpZiAocC5mb2xsb3dlcnMuY2hhbmdlID09PSBudWxsKSB7CiAgICAgICAgZm9sbG93ZXJUZC5hcHBlbmRDaGlsZChkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShwLmZvbGxvd2Vycy5sYXRlc3QgIT09IG51bGwgPyBGb3JtYXQubnVtYmVyKHAuZm9sbG93ZXJzLmxhdGVzdCkgOiAn4oCUJykpOwogICAgICB9IGVsc2UgewogICAgICAgIGNvbnN0IGZvbGxvd2VyV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgICAgICBmb2xsb3dlcldyYXAuc3R5bGUuZGlzcGxheSA9ICdpbmxpbmUtZmxleCc7CiAgICAgICAgZm9sbG93ZXJXcmFwLnN0eWxlLmFsaWduSXRlbXMgPSAnY2VudGVyJzsKICAgICAgICBmb2xsb3dlcldyYXAuc3R5bGUuZ2FwID0gJzZweCc7CiAgICAgICAgZm9sbG93ZXJXcmFwLmFwcGVuZENoaWxkKHRleHRFbCgnc3BhbicsIGAke3AuZm9sbG93ZXJzLmNoYW5nZSA+IDAgPyAnKycgOiAnJ30ke0Zvcm1hdC5udW1iZXIocC5mb2xsb3dlcnMuY2hhbmdlKX1gLCBgc3RhdC1kZWx0YSAke0Zvcm1hdC5kZWx0YUNsYXNzKHAuZm9sbG93ZXJzLmNoYW5nZSl9YCkpOwogICAgICAgIGlmIChwLmZvbGxvd2Vycy5jaGFuZ2VQY3QgIT09IG51bGwpIGZvbGxvd2VyV3JhcC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ3NwYW4nLCBgKCR7Rm9ybWF0LnBjdChwLmZvbGxvd2Vycy5jaGFuZ2VQY3QpfSlgLCAncG9zdC1tZXRhJykpOwogICAgICAgIGZvbGxvd2VyVGQuYXBwZW5kQ2hpbGQoZm9sbG93ZXJXcmFwKTsKICAgICAgfQogICAgICB0ci5hcHBlbmRDaGlsZChmb2xsb3dlclRkKTsKICAgICAgdGJvZHkuYXBwZW5kQ2hpbGQodHIpOwogICAgfSk7CiAgICB0YWJsZS5hcHBlbmRDaGlsZCh0Ym9keSk7CiAgICBjb25zdCB0YWJsZVNjcm9sbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGFibGVTY3JvbGwuY2xhc3NOYW1lID0gJ3RhYmxlLXNjcm9sbCc7CiAgICB0YWJsZVNjcm9sbC5hcHBlbmRDaGlsZCh0YWJsZSk7CiAgICB0YWJsZUNhcmQuYXBwZW5kQ2hpbGQodGFibGVTY3JvbGwpOwogICAgd3JhcC5hcHBlbmRDaGlsZCh0YWJsZUNhcmQpOwoKICAgIGNvbnN0IGNoYXJ0Q2FyZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgY2hhcnRDYXJkLmNsYXNzTmFtZSA9ICdjYXJkJzsKICAgIGNvbnN0IGNoYXJ0SGVhZGVyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjaGFydEhlYWRlci5jbGFzc05hbWUgPSAnY2FyZC1oZWFkZXInOwogICAgY2hhcnRIZWFkZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMycsICdNZXRyaWMgQ29tcGFyaXNvbicpKTsKICAgIGNvbnN0IG1ldHJpY1NlbGVjdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NlbGVjdCcpOwogICAgTUVUUklDX1JPV1MuZm9yRWFjaCgobSkgPT4gewogICAgICBjb25zdCBvcHQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsgb3B0LnZhbHVlID0gbS5rZXk7IG9wdC50ZXh0Q29udGVudCA9IG0ubGFiZWw7CiAgICAgIGlmIChtLmtleSA9PT0gcGxhdGZvcm1DaGFydE1ldHJpYykgb3B0LnNlbGVjdGVkID0gdHJ1ZTsKICAgICAgbWV0cmljU2VsZWN0LmFwcGVuZENoaWxkKG9wdCk7CiAgICB9KTsKICAgIG1ldHJpY1NlbGVjdC5hZGRFdmVudExpc3RlbmVyKCdjaGFuZ2UnLCAoKSA9PiB7CiAgICAgIHBsYXRmb3JtQ2hhcnRNZXRyaWMgPSBtZXRyaWNTZWxlY3QudmFsdWU7CiAgICAgIGRyYXdQbGF0Zm9ybVJlcG9ydENoYXJ0KHJlcG9ydCk7CiAgICB9KTsKICAgIGNoYXJ0SGVhZGVyLmFwcGVuZENoaWxkKG1ldHJpY1NlbGVjdCk7CiAgICBjb25zdCBjaGFydFdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNoYXJ0V3JhcC5jbGFzc05hbWUgPSAnY2hhcnQtd3JhcCB0YWxsJzsKICAgIGNoYXJ0V3JhcC5pbm5lckhUTUwgPSAnPGNhbnZhcyBpZD0icGxhdGZvcm1SZXBvcnRDYW52YXMiPjwvY2FudmFzPic7CiAgICBjaGFydENhcmQuYXBwZW5kKGNoYXJ0SGVhZGVyLCBjaGFydFdyYXApOwogICAgd3JhcC5hcHBlbmRDaGlsZChjaGFydENhcmQpOwogICAgZHJhd1BsYXRmb3JtUmVwb3J0Q2hhcnQocmVwb3J0KTsKCiAgICBjb25zdCB3aXRoRm9sbG93ZXJzID0gcmVwb3J0LnBsYXRmb3Jtcy5maWx0ZXIoKHApID0+IHAuZm9sbG93ZXJzLmxhdGVzdCAhPT0gbnVsbCk7CiAgICBpZiAod2l0aEZvbGxvd2Vycy5sZW5ndGgpIHsKICAgICAgY29uc3QgZm9sbG93ZXJDYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIGZvbGxvd2VyQ2FyZC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICAgIGZvbGxvd2VyQ2FyZC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2gzJywgJ0ZvbGxvd2VyIEdyb3d0aCBieSBQbGF0Zm9ybScpKTsKICAgICAgY29uc3QgZkNoYXJ0V3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBmQ2hhcnRXcmFwLmNsYXNzTmFtZSA9ICdjaGFydC13cmFwIHRhbGwnOwogICAgICBmQ2hhcnRXcmFwLmlubmVySFRNTCA9ICc8Y2FudmFzIGlkPSJwbGF0Zm9ybUZvbGxvd2VyQ2FudmFzIj48L2NhbnZhcz4nOwogICAgICBmb2xsb3dlckNhcmQuYXBwZW5kQ2hpbGQoZkNoYXJ0V3JhcCk7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQoZm9sbG93ZXJDYXJkKTsKICAgICAgQ2hhcnRzLnBsYXRmb3JtQmFyQ2hhcnQoJ3BsYXRmb3JtRm9sbG93ZXJDYW52YXMnLCB7CiAgICAgICAgbGFiZWxzOiB3aXRoRm9sbG93ZXJzLm1hcCgocCkgPT4gcC5sYWJlbCksCiAgICAgICAgZGF0YTogd2l0aEZvbGxvd2Vycy5tYXAoKHApID0+IHAuZm9sbG93ZXJzLmxhdGVzdCB8fCAwKSwKICAgICAgICBjb2xvcnM6IHdpdGhGb2xsb3dlcnMubWFwKChwKSA9PiBwLmNvbG9yKSwKICAgICAgICBmb3JtYXRWYWx1ZTogKHYpID0+IEZvcm1hdC5zbWFydCh2KSwKICAgICAgfSk7CiAgICB9CgogICAgY29uc3QgaW5zaWdodHNDYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBpbnNpZ2h0c0NhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgaW5zaWdodHNDYXJkLmFwcGVuZENoaWxkKHRleHRFbCgnaDMnLCAnSW5zaWdodHMgJiBTdW1tYXJ5JykpOwogICAgY29uc3QgbGlzdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3VsJyk7CiAgICBsaXN0LmNsYXNzTmFtZSA9ICdpbnNpZ2h0cy1saXN0JzsKICAgIHJlcG9ydC5pbnNpZ2h0cy5mb3JFYWNoKChsaW5lKSA9PiBsaXN0LmFwcGVuZENoaWxkKHRleHRFbCgnbGknLCBsaW5lKSkpOwogICAgaW5zaWdodHNDYXJkLmFwcGVuZENoaWxkKGxpc3QpOwogICAgd3JhcC5hcHBlbmRDaGlsZChpbnNpZ2h0c0NhcmQpOwogIH0KCiAgZnVuY3Rpb24gZHJhd1BsYXRmb3JtUmVwb3J0Q2hhcnQocmVwb3J0KSB7CiAgICBDaGFydHMucGxhdGZvcm1CYXJDaGFydCgncGxhdGZvcm1SZXBvcnRDYW52YXMnLCB7CiAgICAgIGxhYmVsczogcmVwb3J0LnBsYXRmb3Jtcy5tYXAoKHApID0+IHAubGFiZWwpLAogICAgICBkYXRhOiByZXBvcnQucGxhdGZvcm1zLm1hcCgocCkgPT4gcC50b3RhbHNbcGxhdGZvcm1DaGFydE1ldHJpY10gfHwgMCksCiAgICAgIGNvbG9yczogcmVwb3J0LnBsYXRmb3Jtcy5tYXAoKHApID0+IHAuY29sb3IpLAogICAgICBmb3JtYXRWYWx1ZTogKHYpID0+IChwbGF0Zm9ybUNoYXJ0TWV0cmljID09PSAnd2F0Y2hfdGltZV9zZWNvbmRzJyA/IEZvcm1hdC5kdXJhdGlvbih2KSA6IEZvcm1hdC5zbWFydCh2KSksCiAgICB9KTsKICB9CgogIC8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgICBQbGF0Zm9ybSBQZXJmb3JtYW5jZSBDb21wYXJpc29uIOKAlCByZXBsYWNlcyB0aGUgb2xkIGdyb3VwZWQKICAgICAiUmFuZ2UgQSB2cyBSYW5nZSBCIGJ5IFBsYXRmb3JtIiBjaGFydC4gT25lIGNhcmQgcGVyIHBsYXRmb3JtCiAgICAgd2l0aCBhbnkgZGF0YSBpbiBlaXRoZXIgcmFuZ2UsIGJ1aWx0IGVudGlyZWx5IGZyb20gdGhlIHNhbWUKICAgICBjb21wYXJlUmFuZ2VzKCkgcmVzcG9uc2UgdGhlIHN0YXQtdGlsZSBncmlkIGFib3ZlIGFscmVhZHkKICAgICB1c2VzIChyZXN1bHQucmFuZ2VBLnBsYXRmb3JtcyAvIHJlc3VsdC5yYW5nZUIucGxhdGZvcm1zKSDigJQgbm8KICAgICBleHRyYSBmZXRjaC4KICAgICA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCiAgY29uc3QgQUxMX0NBUkRfTUVUUklDUyA9IFt7IGtleTogJ3Bvc3RfY291bnQnLCBsYWJlbDogJ1Bvc3RzJyB9LCAuLi5NRVRSSUNfUk9XU107CiAgY29uc3QgQ0FSRF9TT1JUX01PREVTID0gWwogICAgeyBrZXk6ICdvdmVyYWxsJywgbGFiZWw6ICdPdmVyYWxsIFBlcmZvcm1hbmNlJyB9LAogICAgeyBrZXk6ICdncm93dGgnLCBsYWJlbDogJ0hpZ2hlc3QgR3Jvd3RoJyB9LAogICAgeyBrZXk6ICdlbmdhZ2VtZW50JywgbGFiZWw6ICdIaWdoZXN0IEVuZ2FnZW1lbnQnIH0sCiAgICB7IGtleTogJ2ZvbGxvd2VycycsIGxhYmVsOiAnTW9zdCBGb2xsb3dlcnMnIH0sCiAgICB7IGtleTogJ3Bvc3RzJywgbGFiZWw6ICdNb3N0IFBvc3RzJyB9LAogICAgeyBrZXk6ICdhbHBoYScsIGxhYmVsOiAnQWxwaGFiZXRpY2FsJyB9LAogIF07CgogIC8qKiBQZXItbWV0cmljIHthLCBiLCBkaWZmLCBwY3REaWZmfSBhY3Jvc3MgYm90aCByYW5nZXMgZm9yIG9uZSBwbGF0Zm9ybSwgc2tpcHBpbmcgYW55IG1ldHJpYyB0aGF0J3MgemVybyBpbiBib3RoIOKAlCBhIHBsYXRmb3JtJ3MgY2FyZCBzaG91bGQgb25seSBldmVyIHNob3cgbWV0cmljcyBpdCBhY3R1YWxseSBoYXMuICovCiAgZnVuY3Rpb24gY29tcHV0ZUNhcmRNZXRyaWNzKHBsYXRmb3JtQSwgcGxhdGZvcm1CKSB7CiAgICBjb25zdCBtZXRyaWNzID0gW107CiAgICBBTExfQ0FSRF9NRVRSSUNTLmZvckVhY2goKHsga2V5LCBsYWJlbCB9KSA9PiB7CiAgICAgIGNvbnN0IGEgPSAocGxhdGZvcm1BICYmIHBsYXRmb3JtQVtrZXldKSB8fCAwOwogICAgICBjb25zdCBiID0gKHBsYXRmb3JtQiAmJiBwbGF0Zm9ybUJba2V5XSkgfHwgMDsKICAgICAgaWYgKGEgPT09IDAgJiYgYiA9PT0gMCkgcmV0dXJuOwogICAgICBjb25zdCBkaWZmID0gYSAtIGI7CiAgICAgIGNvbnN0IHBjdERpZmYgPSBiID8gTWF0aC5yb3VuZCgoZGlmZiAvIGIpICogMTAwMCkgLyAxMCA6IChhID4gMCA/IG51bGwgOiAwKTsKICAgICAgbWV0cmljcy5wdXNoKHsga2V5LCBsYWJlbCwgYSwgYiwgZGlmZiwgcGN0RGlmZiwgaXNEdXJhdGlvbjoga2V5ID09PSAnd2F0Y2hfdGltZV9zZWNvbmRzJyB9KTsKICAgIH0pOwogICAgcmV0dXJuIG1ldHJpY3M7CiAgfQoKICAvKiogQSBzaW5nbGUgImhvdyBkaWQgdGhpcyBwbGF0Zm9ybSBkbyBvdmVyYWxsIiBudW1iZXI6IHRoZSBhdmVyYWdlICUgY2hhbmdlIGFjcm9zcyBldmVyeSBtZXRyaWMgdGhhdCBoYXMgYSBjb21wdXRhYmxlIHBlcmNlbnRhZ2UgKGEgbWV0cmljIGdvaW5nIGZyb20gMCB0byBzb21ldGhpbmcgaGFzIG5vIHBlcmNlbnRhZ2Ug4oCUICJuZXciLCBub3QgY291bnRlZCBlaXRoZXIgd2F5KS4gKi8KICBmdW5jdGlvbiBvdmVyYWxsUGN0Q2hhbmdlKG1ldHJpY3MpIHsKICAgIGNvbnN0IHdpdGhQY3QgPSBtZXRyaWNzLmZpbHRlcigobSkgPT4gbS5wY3REaWZmICE9PSBudWxsKTsKICAgIGlmICghd2l0aFBjdC5sZW5ndGgpIHJldHVybiBudWxsOwogICAgcmV0dXJuIE1hdGgucm91bmQoKHdpdGhQY3QucmVkdWNlKChzdW0sIG0pID0+IHN1bSArIG0ucGN0RGlmZiwgMCkgLyB3aXRoUGN0Lmxlbmd0aCkgKiAxMCkgLyAxMDsKICB9CgogIGZ1bmN0aW9uIGJlc3RXZWFrZXN0TWV0cmljKG1ldHJpY3MpIHsKICAgIGNvbnN0IHdpdGhQY3QgPSBtZXRyaWNzLmZpbHRlcigobSkgPT4gbS5wY3REaWZmICE9PSBudWxsKTsKICAgIGlmICghd2l0aFBjdC5sZW5ndGgpIHJldHVybiB7IGJlc3Q6IG51bGwsIHdlYWtlc3Q6IG51bGwgfTsKICAgIGNvbnN0IGJlc3QgPSB3aXRoUGN0LnJlZHVjZSgoYSwgYikgPT4gKGIucGN0RGlmZiA+IGEucGN0RGlmZiA/IGIgOiBhKSk7CiAgICBjb25zdCB3ZWFrZXN0ID0gd2l0aFBjdC5yZWR1Y2UoKGEsIGIpID0+IChiLnBjdERpZmYgPCBhLnBjdERpZmYgPyBiIDogYSkpOwogICAgcmV0dXJuIHsgYmVzdCwgd2Vha2VzdCB9OwogIH0KCiAgZnVuY3Rpb24gdHJlbmREaXJlY3Rpb24ocGN0KSB7CiAgICBpZiAocGN0ID09PSBudWxsIHx8IHBjdCA9PT0gdW5kZWZpbmVkKSByZXR1cm4gJ2ZsYXQnOwogICAgaWYgKHBjdCA+IDAuNSkgcmV0dXJuICd1cCc7CiAgICBpZiAocGN0IDwgLTAuNSkgcmV0dXJuICdkb3duJzsKICAgIHJldHVybiAnZmxhdCc7CiAgfQoKICBmdW5jdGlvbiBidWlsZFBsYXRmb3JtQ2FyZHMocmVzdWx0KSB7CiAgICBjb25zdCBwbGF0Zm9ybU9wdGlvbnMgPSAod2luZG93Ll9fZmlsdGVyT3B0aW9uc0NhY2hlIHx8IHsgYWxsUGxhdGZvcm1zOiBbXSB9KS5hbGxQbGF0Zm9ybXM7CiAgICBjb25zdCBpZHMgPSBbLi4ubmV3IFNldChbLi4ucmVzdWx0LnJhbmdlQS5wbGF0Zm9ybXMsIC4uLnJlc3VsdC5yYW5nZUIucGxhdGZvcm1zXS5tYXAoKHApID0+IHAucGxhdGZvcm0pKV07CiAgICBjb25zdCBieUlkQSA9IE9iamVjdC5mcm9tRW50cmllcyhyZXN1bHQucmFuZ2VBLnBsYXRmb3Jtcy5tYXAoKHApID0+IFtwLnBsYXRmb3JtLCBwXSkpOwogICAgY29uc3QgYnlJZEIgPSBPYmplY3QuZnJvbUVudHJpZXMocmVzdWx0LnJhbmdlQi5wbGF0Zm9ybXMubWFwKChwKSA9PiBbcC5wbGF0Zm9ybSwgcF0pKTsKCiAgICByZXR1cm4gaWRzCiAgICAgIC5tYXAoKGlkKSA9PiB7CiAgICAgICAgY29uc3QgbWV0YSA9IHBsYXRmb3JtT3B0aW9ucy5maW5kKChwKSA9PiBwLmlkID09PSBpZCkgfHwgeyBpZCwgbGFiZWw6IGlkLCBjb2xvcjogJ3ZhcigtLXNlcmllcy0xKScgfTsKICAgICAgICBjb25zdCBhID0gYnlJZEFbaWRdIHx8IG51bGw7CiAgICAgICAgY29uc3QgYiA9IGJ5SWRCW2lkXSB8fCBudWxsOwogICAgICAgIGNvbnN0IG1ldHJpY3MgPSBjb21wdXRlQ2FyZE1ldHJpY3MoYSwgYik7CiAgICAgICAgY29uc3QgeyBiZXN0LCB3ZWFrZXN0IH0gPSBiZXN0V2Vha2VzdE1ldHJpYyhtZXRyaWNzKTsKICAgICAgICByZXR1cm4gewogICAgICAgICAgcGxhdGZvcm06IGlkLAogICAgICAgICAgbGFiZWw6IG1ldGEubGFiZWwsCiAgICAgICAgICBjb2xvcjogbWV0YS5jb2xvciwKICAgICAgICAgIG1ldHJpY3MsCiAgICAgICAgICBvdmVyYWxsOiBvdmVyYWxsUGN0Q2hhbmdlKG1ldHJpY3MpLAogICAgICAgICAgYmVzdCwKICAgICAgICAgIHdlYWtlc3QsCiAgICAgICAgICBmb2xsb3dlcnNHYWluZWQ6IChhID8gYS5mb2xsb3dlcnNfZ2FpbmVkIHx8IDAgOiAwKSArIChiID8gYi5mb2xsb3dlcnNfZ2FpbmVkIHx8IDAgOiAwKSwKICAgICAgICAgIHBvc3RzOiAoYSA/IGEucG9zdF9jb3VudCB8fCAwIDogMCkgKyAoYiA/IGIucG9zdF9jb3VudCB8fCAwIDogMCksCiAgICAgICAgICBlbmdhZ2VtZW50VG90YWw6IChhID8gYS5lbmdhZ2VtZW50IHx8IDAgOiAwKSArIChiID8gYi5lbmdhZ2VtZW50IHx8IDAgOiAwKSwKICAgICAgICB9OwogICAgICB9KQogICAgICAuZmlsdGVyKChjYXJkKSA9PiBjYXJkLm1ldHJpY3MubGVuZ3RoID4gMCk7CiAgfQoKICBmdW5jdGlvbiBzb3J0Q2FyZHMoY2FyZHMsIHNvcnRNb2RlKSB7CiAgICBjb25zdCBhcnIgPSBbLi4uY2FyZHNdOwogICAgaWYgKHNvcnRNb2RlID09PSAnZW5nYWdlbWVudCcpIHJldHVybiBhcnIuc29ydCgoeCwgeSkgPT4geS5lbmdhZ2VtZW50VG90YWwgLSB4LmVuZ2FnZW1lbnRUb3RhbCk7CiAgICBpZiAoc29ydE1vZGUgPT09ICdmb2xsb3dlcnMnKSByZXR1cm4gYXJyLnNvcnQoKHgsIHkpID0+IHkuZm9sbG93ZXJzR2FpbmVkIC0geC5mb2xsb3dlcnNHYWluZWQpOwogICAgaWYgKHNvcnRNb2RlID09PSAncG9zdHMnKSByZXR1cm4gYXJyLnNvcnQoKHgsIHkpID0+IHkucG9zdHMgLSB4LnBvc3RzKTsKICAgIGlmIChzb3J0TW9kZSA9PT0gJ2FscGhhJykgcmV0dXJuIGFyci5zb3J0KCh4LCB5KSA9PiB4LmxhYmVsLmxvY2FsZUNvbXBhcmUoeS5sYWJlbCkpOwogICAgLy8gJ292ZXJhbGwnIGFuZCAnZ3Jvd3RoJyBib3RoIHJhbmsgYnkgdGhlIHNhbWUgY29tcG9zaXRlICUgY2hhbmdlIOKAlCB0aGUgdHdvIGxhYmVscwogICAgLy8gcmVhZCBkaWZmZXJlbnRseSBvbiB0aGUgc2FtZSB1bmRlcmx5aW5nIG51bWJlciwgcGVyIHRoZSByZXF1ZXN0ZWQgb3B0aW9uIGxpc3QuCiAgICByZXR1cm4gYXJyLnNvcnQoKHgsIHkpID0+ICh5Lm92ZXJhbGwgPz8gLUluZmluaXR5KSAtICh4Lm92ZXJhbGwgPz8gLUluZmluaXR5KSk7CiAgfQoKICBmdW5jdGlvbiBidWlsZE1ldHJpY1JvdyhtKSB7CiAgICBjb25zdCByb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHJvdy5jbGFzc05hbWUgPSAncGNjLW1ldHJpYy1yb3cnOwogICAgY29uc3QgaGVhZGVyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBoZWFkZXIuY2xhc3NOYW1lID0gJ3BjYy1tZXRyaWMtaGVhZGVyJzsKICAgIGNvbnN0IGZtdCA9ICh2KSA9PiAobS5pc0R1cmF0aW9uID8gRm9ybWF0LmR1cmF0aW9uKHYpIDogRm9ybWF0LnNtYXJ0KHYpKTsKICAgIGNvbnN0IGRpZmZUZXh0ID0gbS5wY3REaWZmID09PSBudWxsCiAgICAgID8gYCR7bS5kaWZmID4gMCA/ICcrJyA6ICcnfSR7Zm10KG0uZGlmZil9IChuZXcpYAogICAgICA6IGAke20uZGlmZiA+IDAgPyAnKycgOiAnJ30ke2ZtdChtLmRpZmYpfSAoJHtGb3JtYXQucGN0KG0ucGN0RGlmZil9KWA7CiAgICBoZWFkZXIuYXBwZW5kKAogICAgICB0ZXh0RWwoJ3NwYW4nLCBtLmxhYmVsLCAncGNjLW1ldHJpYy1sYWJlbCcpLAogICAgICB0ZXh0RWwoJ3NwYW4nLCBkaWZmVGV4dCwgYHBjYy1tZXRyaWMtZGlmZiAke0Zvcm1hdC5kZWx0YUNsYXNzKG0ucGN0RGlmZil9YCkKICAgICk7CiAgICByb3cuYXBwZW5kQ2hpbGQoaGVhZGVyKTsKICAgIGNvbnN0IG1heCA9IE1hdGgubWF4KG0uYSwgbS5iLCAxKTsKICAgIHJvdy5hcHBlbmRDaGlsZChidWlsZEJhcih7IGxhYmVsOiAnQScsIHZhbHVlOiBtLmEsIG1heCwgY29sb3JWYXI6ICctLXNlcmllcy0xJywgZm9ybWF0VmFsdWU6IGZtdCB9KSk7CiAgICByb3cuYXBwZW5kQ2hpbGQoYnVpbGRCYXIoeyBsYWJlbDogJ0InLCB2YWx1ZTogbS5iLCBtYXgsIGNvbG9yVmFyOiAnLS10ZXh0LW11dGVkJywgZm9ybWF0VmFsdWU6IGZtdCB9KSk7CiAgICByZXR1cm4gcm93OwogIH0KCiAgZnVuY3Rpb24gYnVpbGRDYXJkRm9vdGVyKGNhcmQpIHsKICAgIGNvbnN0IGZvb3RlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZm9vdGVyLmNsYXNzTmFtZSA9ICdwY2MtZm9vdGVyJzsKICAgIGNvbnN0IGRpciA9IHRyZW5kRGlyZWN0aW9uKGNhcmQub3ZlcmFsbCk7CiAgICBjb25zdCByZXN1bHRUZXh0ID0gY2FyZC5vdmVyYWxsID09PSBudWxsCiAgICAgID8gJ05vdCBlbm91Z2ggZGF0YSB0byBjb21wYXJlJwogICAgICA6IGAke2RpciA9PT0gJ3VwJyA/ICdJbXByb3ZlZCcgOiBkaXIgPT09ICdkb3duJyA/ICdEZWNsaW5lZCcgOiAnTm8gc2lnbmlmaWNhbnQgY2hhbmdlJ30ke2RpciAhPT0gJ2ZsYXQnID8gYCBieSAke01hdGguYWJzKGNhcmQub3ZlcmFsbCl9JWAgOiAnJ31gOwogICAgZm9vdGVyLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgJ092ZXJhbGwgUmVzdWx0JywgJ3BjYy1mb290ZXItbGFiZWwnKSk7CiAgICBmb290ZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCByZXN1bHRUZXh0LCBgcGNjLWZvb3Rlci12YWx1ZSAke2Rpcn1gKSk7CiAgICBpZiAoY2FyZC5iZXN0KSBmb290ZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBgQmVzdCBNZXRyaWM6ICR7Y2FyZC5iZXN0LmxhYmVsfSAoJHtGb3JtYXQucGN0KGNhcmQuYmVzdC5wY3REaWZmKX0pYCwgJ3BjYy1mb290ZXItZGV0YWlsJykpOwogICAgaWYgKGNhcmQud2Vha2VzdCAmJiBjYXJkLndlYWtlc3QgIT09IGNhcmQuYmVzdCkgewogICAgICBmb290ZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBgV2Vha2VzdCBNZXRyaWM6ICR7Y2FyZC53ZWFrZXN0LmxhYmVsfSAoJHtGb3JtYXQucGN0KGNhcmQud2Vha2VzdC5wY3REaWZmKX0pYCwgJ3BjYy1mb290ZXItZGV0YWlsJykpOwogICAgfQogICAgcmV0dXJuIGZvb3RlcjsKICB9CgogIC8qKiBTZWxmLWNvbnRhaW5lZCBtb2RhbCBmb3IgIlZpZXcgRnVsbCBDb21wYXJpc29uIiDigJQgYSBzZXBhcmF0ZSBvdmVybGF5IGlkIGZyb20gdGhlIERhdGEgUmVjb3JkcyBFZGl0IG1vZGFsIChSZWNvcmRzLm1vZGFsU2hlbGwgaXMgYSBwcml2YXRlIGNsb3N1cmUgb2YgdGhhdCBtb2R1bGUsIG5vdCBzaGFyZWQgc3RhdGUpLCBzYW1lIHZpc3VhbCBsYW5ndWFnZSAoLm1vZGFsLW92ZXJsYXkgLyAubW9kYWwtcGFuZWwpIHNvIGl0IGxvb2tzIGlkZW50aWNhbC4gKi8KICBmdW5jdGlvbiBjbG9zZUNhcmRNb2RhbCgpIHsKICAgIGNvbnN0IG92ZXJsYXkgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnY29tcGFyaXNvbk1vZGFsT3ZlcmxheScpOwogICAgaWYgKG92ZXJsYXkpIG92ZXJsYXkucmVtb3ZlKCk7CiAgfQoKICBmdW5jdGlvbiBvcGVuQ2FyZE1vZGFsKGNhcmQsIGxhYmVsQSwgbGFiZWxCKSB7CiAgICBjbG9zZUNhcmRNb2RhbCgpOwogICAgY29uc3Qgb3ZlcmxheSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgb3ZlcmxheS5jbGFzc05hbWUgPSAnbW9kYWwtb3ZlcmxheSc7CiAgICBvdmVybGF5LmlkID0gJ2NvbXBhcmlzb25Nb2RhbE92ZXJsYXknOwogICAgb3ZlcmxheS5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIChlKSA9PiB7IGlmIChlLnRhcmdldCA9PT0gb3ZlcmxheSkgY2xvc2VDYXJkTW9kYWwoKTsgfSk7CiAgICBjb25zdCBwYW5lbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcGFuZWwuY2xhc3NOYW1lID0gJ21vZGFsLXBhbmVsIHdpZGUnOwogICAgcGFuZWwuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMicsIGAke2NhcmQubGFiZWx9IOKAlCBGdWxsIENvbXBhcmlzb25gKSk7CiAgICBwYW5lbC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIGAke2xhYmVsQX0gdnMgJHtsYWJlbEJ9YCwgJ21vZGFsLXN1YicpKTsKCiAgICBjb25zdCB0YWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICB0YWJsZS5jbGFzc05hbWUgPSAnZGF0YS10YWJsZSc7CiAgICBjb25zdCB0aGVhZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RoZWFkJyk7CiAgICB0aGVhZC5pbm5lckhUTUwgPSAnPHRyPjx0aD5NZXRyaWM8L3RoPjx0aCBjbGFzcz0ibnVtIj5SYW5nZSBBPC90aD48dGggY2xhc3M9Im51bSI+UmFuZ2UgQjwvdGg+PHRoIGNsYXNzPSJudW0iPkRpZmZlcmVuY2U8L3RoPjx0aCBjbGFzcz0ibnVtIj4lIERpZmZlcmVuY2U8L3RoPjx0aD5UcmVuZDwvdGg+PC90cj4nOwogICAgY29uc3QgdGJvZHkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0Ym9keScpOwogICAgY2FyZC5tZXRyaWNzLmZvckVhY2goKG0pID0+IHsKICAgICAgY29uc3QgZm10ID0gKHYpID0+IChtLmlzRHVyYXRpb24gPyBGb3JtYXQuZHVyYXRpb24odikgOiBGb3JtYXQuc21hcnQodikpOwogICAgICBjb25zdCB0ciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAgIGNvbnN0IHRyZW5kRWwgPSB0ZXh0RWwoJ3NwYW4nLCB0cmVuZERpcmVjdGlvbihtLnBjdERpZmYpID09PSAndXAnID8gJ+KWsicgOiB0cmVuZERpcmVjdGlvbihtLnBjdERpZmYpID09PSAnZG93bicgPyAn4pa8JyA6ICfigJQnLCBgc3RhdC1kZWx0YSAke0Zvcm1hdC5kZWx0YUNsYXNzKG0ucGN0RGlmZil9YCk7CiAgICAgIGNvbnN0IHRyZW5kVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICB0cmVuZFRkLmFwcGVuZENoaWxkKHRyZW5kRWwpOwogICAgICB0ci5hcHBlbmQoCiAgICAgICAgdGV4dEVsKCd0ZCcsIG0ubGFiZWwpLAogICAgICAgIHRleHRFbCgndGQnLCBmbXQobS5hKSwgJ251bScpLAogICAgICAgIHRleHRFbCgndGQnLCBmbXQobS5iKSwgJ251bScpLAogICAgICAgIHRleHRFbCgndGQnLCBgJHttLmRpZmYgPiAwID8gJysnIDogJyd9JHtmbXQobS5kaWZmKX1gLCAnbnVtJyksCiAgICAgICAgdGV4dEVsKCd0ZCcsIG0ucGN0RGlmZiA9PT0gbnVsbCA/ICduZXcnIDogRm9ybWF0LnBjdChtLnBjdERpZmYpLCAnbnVtJyksCiAgICAgICAgdHJlbmRUZAogICAgICApOwogICAgICB0Ym9keS5hcHBlbmRDaGlsZCh0cik7CiAgICB9KTsKICAgIHRhYmxlLmFwcGVuZCh0aGVhZCwgdGJvZHkpOwogICAgY29uc3QgdGFibGVTY3JvbGwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRhYmxlU2Nyb2xsLmNsYXNzTmFtZSA9ICd0YWJsZS1zY3JvbGwnOwogICAgdGFibGVTY3JvbGwuYXBwZW5kQ2hpbGQodGFibGUpOwogICAgcGFuZWwuYXBwZW5kQ2hpbGQodGFibGVTY3JvbGwpOwoKICAgIGNvbnN0IGFjdGlvbnMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGFjdGlvbnMuY2xhc3NOYW1lID0gJ21vZGFsLWFjdGlvbnMnOwogICAgY29uc3QgY2xvc2VCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIGNsb3NlQnRuLmNsYXNzTmFtZSA9ICdidG4nOwogICAgY2xvc2VCdG4udHlwZSA9ICdidXR0b24nOwogICAgY2xvc2VCdG4udGV4dENvbnRlbnQgPSAnQ2xvc2UnOwogICAgY2xvc2VCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBjbG9zZUNhcmRNb2RhbCk7CiAgICBhY3Rpb25zLmFwcGVuZENoaWxkKGNsb3NlQnRuKTsKICAgIHBhbmVsLmFwcGVuZENoaWxkKGFjdGlvbnMpOwoKICAgIG92ZXJsYXkuYXBwZW5kQ2hpbGQocGFuZWwpOwogICAgZG9jdW1lbnQuYm9keS5hcHBlbmRDaGlsZChvdmVybGF5KTsKICB9CgogIGZ1bmN0aW9uIGJ1aWxkUGxhdGZvcm1DYXJkKGNhcmQsIGxhYmVsQSwgbGFiZWxCKSB7CiAgICBjb25zdCBlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZWwuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLWNvbXBhcmUtY2FyZCc7CgogICAgY29uc3QgaGVhZGVyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBoZWFkZXIuY2xhc3NOYW1lID0gJ3BjYy1oZWFkZXInOwogICAgY29uc3QgbmFtZVdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIG5hbWVXcmFwLmNsYXNzTmFtZSA9ICdwY2MtaGVhZGVyLW5hbWUnOwogICAgY29uc3QgZG90ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgZG90LmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1kb3QnOwogICAgZG90LnN0eWxlLmJhY2tncm91bmQgPSBjYXJkLmNvbG9yOwogICAgbmFtZVdyYXAuYXBwZW5kKGRvdCwgdGV4dEVsKCdzcGFuJywgY2FyZC5sYWJlbCwgJ3BjYy1uYW1lJykpOwogICAgaGVhZGVyLmFwcGVuZENoaWxkKG5hbWVXcmFwKTsKICAgIGNvbnN0IGRpciA9IHRyZW5kRGlyZWN0aW9uKGNhcmQub3ZlcmFsbCk7CiAgICBoZWFkZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCBjYXJkLm92ZXJhbGwgPT09IG51bGwgPyAn4oCUJyA6IGAke2RpciA9PT0gJ3VwJyA/ICfilrInIDogZGlyID09PSAnZG93bicgPyAn4pa8JyA6ICfigJQnfSAke0Zvcm1hdC5wY3QoY2FyZC5vdmVyYWxsKX1gLCBgcGNjLWJhZGdlICR7ZGlyfWApKTsKICAgIGVsLmFwcGVuZENoaWxkKGhlYWRlcik7CgogICAgY29uc3QgY2FwdGlvbiA9IGNhcmQub3ZlcmFsbCA9PT0gbnVsbAogICAgICA/ICdOb3QgZW5vdWdoIGRhdGEgdG8gY29tcGFyZSB5ZXQnCiAgICAgIDogZGlyID09PSAndXAnID8gJ0ltcHJvdmVkIGNvbXBhcmVkIHRvIHByZXZpb3VzIHBlcmlvZCcKICAgICAgOiBkaXIgPT09ICdkb3duJyA/ICdMb3dlciB0aGFuIHByZXZpb3VzIHBlcmlvZCcKICAgICAgOiAnQWJvdXQgdGhlIHNhbWUgYXMgdGhlIHByZXZpb3VzIHBlcmlvZCc7CiAgICBlbC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIGNhcHRpb24sICdwY2MtY2FwdGlvbicpKTsKCiAgICBjb25zdCBtZXRyaWNzV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgbWV0cmljc1dyYXAuY2xhc3NOYW1lID0gJ3BjYy1tZXRyaWNzJzsKICAgIGNhcmQubWV0cmljcy5mb3JFYWNoKChtKSA9PiBtZXRyaWNzV3JhcC5hcHBlbmRDaGlsZChidWlsZE1ldHJpY1JvdyhtKSkpOwogICAgZWwuYXBwZW5kQ2hpbGQobWV0cmljc1dyYXApOwoKICAgIGVsLmFwcGVuZENoaWxkKGJ1aWxkQ2FyZEZvb3RlcihjYXJkKSk7CgogICAgY29uc3Qgdmlld0xpbmsgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIHZpZXdMaW5rLnR5cGUgPSAnYnV0dG9uJzsKICAgIHZpZXdMaW5rLmNsYXNzTmFtZSA9ICdwY2Mtdmlldy1saW5rJzsKICAgIHZpZXdMaW5rLnRleHRDb250ZW50ID0gJ1ZpZXcgRnVsbCBDb21wYXJpc29uIOKGkic7CiAgICB2aWV3TGluay5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IG9wZW5DYXJkTW9kYWwoY2FyZCwgbGFiZWxBLCBsYWJlbEIpKTsKICAgIGVsLmFwcGVuZENoaWxkKHZpZXdMaW5rKTsKCiAgICByZXR1cm4gZWw7CiAgfQoKICBmdW5jdGlvbiByZW5kZXJQbGF0Zm9ybUNvbXBhcmlzb25DYXJkcyh3cmFwLCByZXN1bHQsIGxhYmVsQSwgbGFiZWxCKSB7CiAgICBjb25zdCBhbGxDYXJkcyA9IGJ1aWxkUGxhdGZvcm1DYXJkcyhyZXN1bHQpOwoKICAgIGNvbnN0IHNlY3Rpb24gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHNlY3Rpb24uY2xhc3NOYW1lID0gJ3BjYy1zZWN0aW9uJzsKICAgIHNlY3Rpb24uYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCAnUGxhdGZvcm0gUGVyZm9ybWFuY2UgQ29tcGFyaXNvbicsICdzZWN0aW9uLXRpdGxlJykpOwoKICAgIGlmICghYWxsQ2FyZHMubGVuZ3RoKSB7CiAgICAgIHNlY3Rpb24uYXBwZW5kQ2hpbGQoZW1wdHlTdGF0ZSh7CiAgICAgICAgaWNvbjogJ2dpdC1jb21wYXJlJywKICAgICAgICB0aXRsZTogJ05vIGRhdGEgYXZhaWxhYmxlIGZvciB0aGUgc2VsZWN0ZWQgZGF0ZSByYW5nZXMuJywKICAgICAgICBtZXNzYWdlOiAnVHJ5IGEgd2lkZXIgcmFuZ2UsIG9yIGNoZWNrIHRoYXQgcG9zdHMgZXhpc3QgZm9yIGF0IGxlYXN0IG9uZSBwbGF0Zm9ybSBpbiBSYW5nZSBBIG9yIFJhbmdlIEIuJywKICAgICAgfSkpOwogICAgICB3cmFwLmFwcGVuZENoaWxkKHNlY3Rpb24pOwogICAgICByZXR1cm47CiAgICB9CgogICAgY29uc3QgY29udHJvbHMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNvbnRyb2xzLmNsYXNzTmFtZSA9ICdwY2MtY29udHJvbHMnOwoKICAgIGNvbnN0IHNvcnRTZWxlY3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzZWxlY3QnKTsKICAgIENBUkRfU09SVF9NT0RFUy5mb3JFYWNoKChtKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOyBvcHQudmFsdWUgPSBtLmtleTsgb3B0LnRleHRDb250ZW50ID0gbS5sYWJlbDsKICAgICAgaWYgKG0ua2V5ID09PSBjYXJkU29ydE1vZGUpIG9wdC5zZWxlY3RlZCA9IHRydWU7CiAgICAgIHNvcnRTZWxlY3QuYXBwZW5kQ2hpbGQob3B0KTsKICAgIH0pOwogICAgc29ydFNlbGVjdC5hZGRFdmVudExpc3RlbmVyKCdjaGFuZ2UnLCAoKSA9PiB7IGNhcmRTb3J0TW9kZSA9IHNvcnRTZWxlY3QudmFsdWU7IHJlbmRlckNhcmRHcmlkKCk7IH0pOwogICAgY29udHJvbHMuYXBwZW5kQ2hpbGQobGFiZWxlZCgnU29ydCBCeScsIHNvcnRTZWxlY3QpKTsKCiAgICBjb25zdCBmaWx0ZXJQaWxscyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZmlsdGVyUGlsbHMuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLWZpbHRlci1waWxscyc7CiAgICBjb25zdCBhbGxCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIGFsbEJ0bi50eXBlID0gJ2J1dHRvbic7CiAgICBhbGxCdG4uZGF0YXNldC5maWx0ZXIgPSAnYWxsJzsKICAgIGFsbEJ0bi50ZXh0Q29udGVudCA9ICdBbGwgUGxhdGZvcm1zJzsKICAgIGFsbEJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsgY2FyZFBsYXRmb3JtRmlsdGVyID0gJ2FsbCc7IHJlbmRlckNhcmRHcmlkKCk7IH0pOwogICAgZmlsdGVyUGlsbHMuYXBwZW5kQ2hpbGQoYWxsQnRuKTsKICAgIGFsbENhcmRzLmZvckVhY2goKGNhcmQpID0+IHsKICAgICAgY29uc3QgYnRuID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7CiAgICAgIGJ0bi50eXBlID0gJ2J1dHRvbic7CiAgICAgIGJ0bi5kYXRhc2V0LmZpbHRlciA9IGNhcmQucGxhdGZvcm07CiAgICAgIGNvbnN0IGRvdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgICAgZG90LmNsYXNzTmFtZSA9ICdwbGF0Zm9ybS1kb3QnOwogICAgICBkb3Quc3R5bGUuYmFja2dyb3VuZCA9IGNhcmQuY29sb3I7CiAgICAgIGJ0bi5hcHBlbmQoZG90LCBkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShjYXJkLmxhYmVsKSk7CiAgICAgIGJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsgY2FyZFBsYXRmb3JtRmlsdGVyID0gY2FyZC5wbGF0Zm9ybTsgcmVuZGVyQ2FyZEdyaWQoKTsgfSk7CiAgICAgIGZpbHRlclBpbGxzLmFwcGVuZENoaWxkKGJ0bik7CiAgICB9KTsKICAgIGNvbnRyb2xzLmFwcGVuZENoaWxkKGZpbHRlclBpbGxzKTsKICAgIHNlY3Rpb24uYXBwZW5kQ2hpbGQoY29udHJvbHMpOwoKICAgIGNvbnN0IGdyaWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGdyaWQuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLWNvbXBhcmUtZ3JpZCc7CiAgICBncmlkLmlkID0gJ3BsYXRmb3JtQ29tcGFyZUdyaWQnOwogICAgc2VjdGlvbi5hcHBlbmRDaGlsZChncmlkKTsKICAgIHdyYXAuYXBwZW5kQ2hpbGQoc2VjdGlvbik7CgogICAgZnVuY3Rpb24gcmVuZGVyQ2FyZEdyaWQoKSB7CiAgICAgIGNvbnN0IGdyaWRFbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdwbGF0Zm9ybUNvbXBhcmVHcmlkJyk7CiAgICAgIGlmICghZ3JpZEVsKSByZXR1cm47CiAgICAgIGdyaWRFbC5pbm5lckhUTUwgPSAnJzsKICAgICAgY29uc3QgdmlzaWJsZSA9IGNhcmRQbGF0Zm9ybUZpbHRlciA9PT0gJ2FsbCcgPyBhbGxDYXJkcyA6IGFsbENhcmRzLmZpbHRlcigoYykgPT4gYy5wbGF0Zm9ybSA9PT0gY2FyZFBsYXRmb3JtRmlsdGVyKTsKICAgICAgY29uc3Qgc29ydGVkID0gc29ydENhcmRzKHZpc2libGUsIGNhcmRTb3J0TW9kZSk7CiAgICAgIGlmICghc29ydGVkLmxlbmd0aCkgewogICAgICAgIGdyaWRFbC5hcHBlbmRDaGlsZChlbXB0eVN0YXRlKHsgaWNvbjogJ2dpdC1jb21wYXJlJywgbWVzc2FnZTogJ05vIGRhdGEgZm9yIHRoaXMgcGxhdGZvcm0gaW4gdGhlIHNlbGVjdGVkIGRhdGUgcmFuZ2VzLicgfSkpOwogICAgICB9IGVsc2UgewogICAgICAgIHNvcnRlZC5mb3JFYWNoKChjYXJkKSA9PiBncmlkRWwuYXBwZW5kQ2hpbGQoYnVpbGRQbGF0Zm9ybUNhcmQoY2FyZCwgbGFiZWxBLCBsYWJlbEIpKSk7CiAgICAgIH0KICAgICAgZmlsdGVyUGlsbHMucXVlcnlTZWxlY3RvckFsbCgnYnV0dG9uJykuZm9yRWFjaCgoYnRuKSA9PiB7CiAgICAgICAgYnRuLmNsYXNzTGlzdC50b2dnbGUoJ2lzLWFjdGl2ZScsIGJ0bi5kYXRhc2V0LmZpbHRlciA9PT0gY2FyZFBsYXRmb3JtRmlsdGVyKTsKICAgICAgfSk7CiAgICB9CiAgICByZW5kZXJDYXJkR3JpZCgpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyKCkgewogICAgcm9vdCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd2aWV3LWNvbXBhcmlzb24nKTsKICAgIHNoZWxsKCk7CiAgfQoKICByZXR1cm4geyByZW5kZXIgfTsKfSkoKTsKCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBVcGxvYWQgdGFiOiBkcmFnLWRyb3AsIHZhbGlkYXRpb24gcHJldmlldywgcGVyLXdlZWsgY29uZmxpY3QKICAgcmVzb2x1dGlvbiwgY29tbWl0IOKAlCBwbHVzIHRoZSBVcGxvYWQgSGlzdG9yeSB0YWIuCiAgID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwpjb25zdCBVcGxvYWQgPSAoKCkgPT4gewogIGxldCByb290OwogIGxldCBjdXJyZW50UHJldmlldyA9IG51bGw7IC8vIHsgZmlsZVBhdGgsIG9yaWdpbmFsTmFtZSwgZHVwbGljYXRlcywgaXNzdWVzLCBzYW1wbGUsIC4uLiB9CiAgY29uc3QgZHVwbGljYXRlQWN0aW9uT3ZlcnJpZGVzID0ge307CgogIGZ1bmN0aW9uIHNoZWxsKCkgewogICAgcm9vdC5pbm5lckhUTUwgPSAnJzsKCiAgICBjb25zdCBpbnRybyA9IHRleHRFbCgnZGl2JywgJ1VwbG9hZCBhIHdlZWtseSBleHBvcnQnLCAnc2VjdGlvbi10aXRsZScpOwogICAgcm9vdC5hcHBlbmRDaGlsZChpbnRybyk7CgogICAgY29uc3QgZHJvcHpvbmUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGRyb3B6b25lLmNsYXNzTmFtZSA9ICdkcm9wem9uZSc7CiAgICBkcm9wem9uZS5pZCA9ICdkcm9wem9uZSc7CiAgICBkcm9wem9uZS5pbm5lckhUTUwgPSBgCiAgICAgIDxkaXYgY2xhc3M9ImVtcHR5LWljb24iIHN0eWxlPSJtYXJnaW46IDAgYXV0byAxNHB4OyI+PGkgZGF0YS1sdWNpZGU9InVwbG9hZC1jbG91ZCIgc3R5bGU9IndpZHRoOjIycHg7aGVpZ2h0OjIycHg7Ij48L2k+PC9kaXY+CiAgICAgIDxoMz5EcmFnICZhbXA7IGRyb3AgeW91ciAuY3N2IG9yIC54bHN4IGZpbGUgaGVyZTwvaDM+CiAgICAgIDxwPm9yIGNsaWNrIHRvIGJyb3dzZSDigJQgZmlsZXMgYXJlIHZhbGlkYXRlZCBiZWZvcmUgYW55dGhpbmcgaXMgc2F2ZWQ8L3A+CiAgICAgIDxpbnB1dCB0eXBlPSJmaWxlIiBpZD0iZmlsZUlucHV0IiBhY2NlcHQ9Ii5jc3YsLnhsc3gsLnhscyIgLz4KICAgIGA7CiAgICByb290LmFwcGVuZENoaWxkKGRyb3B6b25lKTsKCiAgICBjb25zdCBwcmV2aWV3QXJlYSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcHJldmlld0FyZWEuaWQgPSAncHJldmlld0FyZWEnOwogICAgcm9vdC5hcHBlbmRDaGlsZChwcmV2aWV3QXJlYSk7CgogICAgd2lyZURyb3B6b25lKGRyb3B6b25lKTsKICB9CgogIGZ1bmN0aW9uIHdpcmVEcm9wem9uZShkcm9wem9uZSkgewogICAgY29uc3QgaW5wdXQgPSBkcm9wem9uZS5xdWVyeVNlbGVjdG9yKCcjZmlsZUlucHV0Jyk7CiAgICBkcm9wem9uZS5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IGlucHV0LmNsaWNrKCkpOwogICAgaW5wdXQuYWRkRXZlbnRMaXN0ZW5lcignY2hhbmdlJywgKCkgPT4gewogICAgICBpZiAoaW5wdXQuZmlsZXNbMF0pIGhhbmRsZUZpbGUoaW5wdXQuZmlsZXNbMF0pOwogICAgfSk7CiAgICBbJ2RyYWdlbnRlcicsICdkcmFnb3ZlciddLmZvckVhY2goKGV2dCkgPT4KICAgICAgZHJvcHpvbmUuYWRkRXZlbnRMaXN0ZW5lcihldnQsIChlKSA9PiB7IGUucHJldmVudERlZmF1bHQoKTsgZHJvcHpvbmUuY2xhc3NMaXN0LmFkZCgnaXMtZHJhZycpOyB9KQogICAgKTsKICAgIFsnZHJhZ2xlYXZlJywgJ2Ryb3AnXS5mb3JFYWNoKChldnQpID0+CiAgICAgIGRyb3B6b25lLmFkZEV2ZW50TGlzdGVuZXIoZXZ0LCAoZSkgPT4geyBlLnByZXZlbnREZWZhdWx0KCk7IGRyb3B6b25lLmNsYXNzTGlzdC5yZW1vdmUoJ2lzLWRyYWcnKTsgfSkKICAgICk7CiAgICBkcm9wem9uZS5hZGRFdmVudExpc3RlbmVyKCdkcm9wJywgKGUpID0+IHsKICAgICAgY29uc3QgZmlsZSA9IGUuZGF0YVRyYW5zZmVyLmZpbGVzWzBdOwogICAgICBpZiAoZmlsZSkgaGFuZGxlRmlsZShmaWxlKTsKICAgIH0pOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gaGFuZGxlRmlsZShmaWxlKSB7CiAgICBjb25zdCBhcmVhID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3ByZXZpZXdBcmVhJyk7CiAgICBhcmVhLmlubmVySFRNTCA9ICcnOwogICAgYXJlYS5hcHBlbmRDaGlsZChyb3dXaXRoU3Bpbm5lcignVmFsaWRhdGluZyBmaWxl4oCmJykpOwogICAgT2JqZWN0LmtleXMoZHVwbGljYXRlQWN0aW9uT3ZlcnJpZGVzKS5mb3JFYWNoKChrKSA9PiBkZWxldGUgZHVwbGljYXRlQWN0aW9uT3ZlcnJpZGVzW2tdKTsKICAgIHRyeSB7CiAgICAgIGN1cnJlbnRQcmV2aWV3ID0gYXdhaXQgQXBpLnByZXZpZXdVcGxvYWQoZmlsZSk7CiAgICAgIHJlbmRlclByZXZpZXcoKTsKICAgIH0gY2F0Y2ggKGVycikgewogICAgICBhcmVhLmlubmVySFRNTCA9ICcnOwogICAgICBhcmVhLmFwcGVuZENoaWxkKGVycm9yQmFubmVyKGVyci5tZXNzYWdlKSk7CiAgICB9CiAgfQoKICBmdW5jdGlvbiByb3dXaXRoU3Bpbm5lcih0ZXh0KSB7CiAgICBjb25zdCBlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZWwuY2xhc3NOYW1lID0gJ2xvYWRpbmctcm93JzsKICAgIGNvbnN0IHNwaW5uZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICBzcGlubmVyLmNsYXNzTmFtZSA9ICdzcGlubmVyJzsKICAgIGVsLmFwcGVuZChzcGlubmVyLCBkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShgICR7dGV4dH1gKSk7CiAgICByZXR1cm4gZWw7CiAgfQogIGZ1bmN0aW9uIGVycm9yQmFubmVyKG1lc3NhZ2UpIHsKICAgIGNvbnN0IGVsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBlbC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICBlbC5zdHlsZS5ib3JkZXJMZWZ0ID0gJzNweCBzb2xpZCB2YXIoLS1zdGF0dXMtY3JpdGljYWwpJzsKICAgIGVsLmFwcGVuZENoaWxkKHRleHRFbCgnZGl2JywgYENvdWxkIG5vdCByZWFkIHRoaXMgZmlsZTogJHttZXNzYWdlfWAsICdtdXRlZCcpKTsKICAgIHJldHVybiBlbDsKICB9CgogIGZ1bmN0aW9uIHJlbmRlclByZXZpZXcoKSB7CiAgICBjb25zdCBhcmVhID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3ByZXZpZXdBcmVhJyk7CiAgICBhcmVhLmlubmVySFRNTCA9ICcnOwogICAgY29uc3QgcCA9IGN1cnJlbnRQcmV2aWV3OwoKICAgIGNvbnN0IHN1bW1hcnlUaXRsZSA9IHRleHRFbCgnZGl2JywgJ1ZhbGlkYXRpb24gc3VtbWFyeScsICdzZWN0aW9uLXRpdGxlJyk7CiAgICBjb25zdCBzdW1tYXJ5R3JpZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgc3VtbWFyeUdyaWQuY2xhc3NOYW1lID0gJ3N0YXQtZ3JpZCc7CiAgICBzdW1tYXJ5R3JpZC5hcHBlbmQoCiAgICAgIHN0YXRUaWxlKCdGaWxlJywgcC5vcmlnaW5hbE5hbWUpLAogICAgICBzdGF0VGlsZSgnU2hlZXRzIGZvdW5kJywgcC5zaGVldHMubGVuZ3RoKSwKICAgICAgc3RhdFRpbGUoJ1RvdGFsIHJvd3MgKGFsbCBzaGVldHMpJywgcC50b3RhbERhdGFSb3dzKSwKICAgICAgc3RhdFRpbGUoJ05ldyByZWNvcmRzJywgcC5uZXdSZWNvcmRzQ291bnQpLAogICAgICBzdGF0VGlsZSgnRXhhY3QgZHVwbGljYXRlcyBmb3VuZCcsIHAuZHVwbGljYXRlcy5sZW5ndGgpLAogICAgICBzdGF0VGlsZSgnRHVwbGljYXRlIHJvd3MgaW4gZmlsZScsIHAuZHVwbGljYXRlUm93c0luRmlsZSksCiAgICAgIHN0YXRUaWxlKCdSb3dzIHdpdGggZXJyb3JzJywgcC5lcnJvclJvd3MpCiAgICApOwogICAgYXJlYS5hcHBlbmQoc3VtbWFyeVRpdGxlLCBzdW1tYXJ5R3JpZCk7CgogICAgaWYgKHAuc2hlZXRzLmxlbmd0aCkgewogICAgICBjb25zdCBzaGVldHNUaXRsZSA9IHRleHRFbCgnZGl2JywgJ1NoZWV0IGJyZWFrZG93bicsICdzZWN0aW9uLXRpdGxlJyk7CiAgICAgIGNvbnN0IHNoZWV0c1RhYmxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGFibGUnKTsKICAgICAgc2hlZXRzVGFibGUuY2xhc3NOYW1lID0gJ2RhdGEtdGFibGUnOwogICAgICBzaGVldHNUYWJsZS5pbm5lckhUTUwgPSAnPHRoZWFkPjx0cj48dGg+U2hlZXQ8L3RoPjx0aD5MYXlvdXQgZGV0ZWN0ZWQ8L3RoPjx0aCBjbGFzcz0ibnVtIj5Sb3dzPC90aD48dGggY2xhc3M9Im51bSI+VmFsaWQ8L3RoPjx0aCBjbGFzcz0ibnVtIj5FcnJvcnM8L3RoPjwvdHI+PC90aGVhZD4nOwogICAgICBjb25zdCBzaGVldHNCb2R5ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGJvZHknKTsKICAgICAgcC5zaGVldHMuZm9yRWFjaCgocykgPT4gewogICAgICAgIGNvbnN0IHRyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgICBjb25zdCBsYXlvdXRMYWJlbCA9IHMuZm9ybWF0ID09PSAnYWdlbmRhJyA/ICdMUlMgYWdlbmRhIHRyYWNrZXInIDogcy5mb3JtYXQgPT09ICdzaW1wbGUnID8gJ1NpbXBsZSBwbGF0Zm9ybSB0YWJsZScgOiAnTm90IHJlY29nbml6ZWQg4oCUIHNhdmVkIGFzIHJhdyBkYXRhIG9ubHknOwogICAgICAgIHRyLmFwcGVuZCgKICAgICAgICAgIHRleHRFbCgndGQnLCBzLm5hbWUpLAogICAgICAgICAgdGV4dEVsKCd0ZCcsIGxheW91dExhYmVsKSwKICAgICAgICAgIHRleHRFbCgndGQnLCBTdHJpbmcocy50b3RhbFJvd3MpLCAnbnVtJyksCiAgICAgICAgICB0ZXh0RWwoJ3RkJywgU3RyaW5nKHMudmFsaWRSb3dzKSwgJ251bScpLAogICAgICAgICAgdGV4dEVsKCd0ZCcsIFN0cmluZyhzLmVycm9yUm93cyksICdudW0nKQogICAgICAgICk7CiAgICAgICAgc2hlZXRzQm9keS5hcHBlbmRDaGlsZCh0cik7CiAgICAgIH0pOwogICAgICBzaGVldHNUYWJsZS5hcHBlbmRDaGlsZChzaGVldHNCb2R5KTsKICAgICAgY29uc3Qgc2hlZXRzV3JhcCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICBzaGVldHNXcmFwLmNsYXNzTmFtZSA9ICd0YWJsZS1zY3JvbGwnOwogICAgICBzaGVldHNXcmFwLmFwcGVuZENoaWxkKHNoZWV0c1RhYmxlKTsKICAgICAgYXJlYS5hcHBlbmQoc2hlZXRzVGl0bGUsIHNoZWV0c1dyYXApOwogICAgfQoKICAgIGlmIChwLmR1cGxpY2F0ZXMubGVuZ3RoKSB7CiAgICAgIGNvbnN0IGR1cFRpdGxlID0gdGV4dEVsKCdkaXYnLCBgRXhhY3QgZHVwbGljYXRlcyBmb3VuZCAoJHtwLmR1cGxpY2F0ZXMubGVuZ3RofSlgLCAnc2VjdGlvbi10aXRsZScpOwogICAgICBhcmVhLmFwcGVuZENoaWxkKGR1cFRpdGxlKTsKICAgICAgYXJlYS5hcHBlbmRDaGlsZCh0ZXh0RWwoCiAgICAgICAgJ2RpdicsCiAgICAgICAgJ0VhY2ggb2YgdGhlc2Ugcm93cyBpcyBieXRlLWZvci1ieXRlIGlkZW50aWNhbCB0byBhbiBhbHJlYWR5LXNhdmVkIHJlY29yZCDigJQgZXZlcnkgZmllbGQgbWF0Y2hlcywgaW5jbHVkaW5nIGV2ZXJ5IG1ldHJpYywgbm90IGp1c3QgdGhlIGRhdGUvY2FwdGlvbi9wbGF0Zm9ybS4gQ2hvb3NlIHdoYXQgdG8gZG8gd2l0aCBlYWNoIOKAlCBvciBzZXQgYSBkZWZhdWx0IGZvciBhbGwgb2YgdGhlbS4gKEEgcm93IHRoYXQgc2hhcmVzIHRoZSBzYW1lIGRhdGUvY2FwdGlvbi9wbGF0Zm9ybSBidXQgaGFzIGRpZmZlcmVudCBudW1iZXJzIGlzIG5vdCBzaG93biBoZXJlIOKAlCBpdOKAmXMgaW1wb3J0ZWQgYXV0b21hdGljYWxseSBhcyBpdHMgb3duIG5ldyByZWNvcmQsIHNpbmNlIGl0cyBhbmFseXRpY3MgY2hhbmdlZC4pJywKICAgICAgICAnbXV0ZWQnCiAgICAgICkpOwogICAgICBjb25zdCBkZWZhdWx0Um93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIGRlZmF1bHRSb3cuY2xhc3NOYW1lID0gJ2ZpZWxkLWlubGluZSc7CiAgICAgIGRlZmF1bHRSb3cuc3R5bGUubWFyZ2luID0gJzEwcHggMCc7CiAgICAgIGNvbnN0IGRlZmF1bHRTZWxlY3QgPSBhY3Rpb25TZWxlY3QoJ3NraXAnKTsKICAgICAgZGVmYXVsdFNlbGVjdC5pZCA9ICdkZWZhdWx0RHVwbGljYXRlQWN0aW9uU2VsZWN0JzsKICAgICAgZGVmYXVsdFNlbGVjdC5hZGRFdmVudExpc3RlbmVyKCdjaGFuZ2UnLCAoKSA9PiB7CiAgICAgICAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnLmNvbmZsaWN0LXJvdyBzZWxlY3RbZGF0YS1oYXNoXScpLmZvckVhY2goKHNlbCkgPT4gewogICAgICAgICAgaWYgKCFkdXBsaWNhdGVBY3Rpb25PdmVycmlkZXNbc2VsLmRhdGFzZXQuaGFzaF0pIHNlbC52YWx1ZSA9IGRlZmF1bHRTZWxlY3QudmFsdWU7CiAgICAgICAgfSk7CiAgICAgIH0pOwogICAgICBkZWZhdWx0Um93LmFwcGVuZCh0ZXh0RWwoJ2xhYmVsJywgJ0RlZmF1bHQgYWN0aW9uIGZvciBhbGwgbWF0Y2hlcycpLCBkZWZhdWx0U2VsZWN0KTsKICAgICAgYXJlYS5hcHBlbmRDaGlsZChkZWZhdWx0Um93KTsKCiAgICAgIGNvbnN0IGxpc3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgbGlzdC5jbGFzc05hbWUgPSAnY29uZmxpY3QtbGlzdCc7CiAgICAgIHAuZHVwbGljYXRlcy5mb3JFYWNoKChkKSA9PiB7CiAgICAgICAgY29uc3Qgcm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgICAgcm93LmNsYXNzTmFtZSA9ICdjb25mbGljdC1yb3cnOwogICAgICAgIGNvbnN0IGxlZnQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgICBsZWZ0LmFwcGVuZCgKICAgICAgICAgIHRleHRFbCgnZGl2JywgYCR7Rm9ybWF0LmRhdGUoZC5wdWJsaXNoRGF0ZSl9IOKAlCAkeyhkLmNhcHRpb24gfHwgJyhubyBjYXB0aW9uKScpLnNsaWNlKDAsIDcwKX1gLCAnd2Vlay1sYWJlbCcpLAogICAgICAgICAgdGV4dEVsKCdkaXYnLCBgRXhhY3QgbWF0Y2ggb2YgZXhpc3RpbmcgcmVjb3JkICMke2QuZXhpc3RpbmcucG9zdElkfSAobGFzdCB1cGRhdGVkICR7ZC5leGlzdGluZy51cGRhdGVkQXR9KWAsICd3ZWVrLW1ldGEnKQogICAgICAgICk7CiAgICAgICAgcm93LmFwcGVuZENoaWxkKGxlZnQpOwogICAgICAgIGNvbnN0IHNlbCA9IGFjdGlvblNlbGVjdCgnc2tpcCcpOwogICAgICAgIHNlbC5kYXRhc2V0Lmhhc2ggPSBkLmhhc2g7CiAgICAgICAgc2VsLmFkZEV2ZW50TGlzdGVuZXIoJ2NoYW5nZScsICgpID0+IHsgZHVwbGljYXRlQWN0aW9uT3ZlcnJpZGVzW2QuaGFzaF0gPSBzZWwudmFsdWU7IH0pOwogICAgICAgIHJvdy5hcHBlbmRDaGlsZChzZWwpOwogICAgICAgIGxpc3QuYXBwZW5kQ2hpbGQocm93KTsKICAgICAgfSk7CiAgICAgIGFyZWEuYXBwZW5kQ2hpbGQobGlzdCk7CiAgICB9CgogICAgY29uc3Qgbm90ZXNGaWVsZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgbm90ZXNGaWVsZC5jbGFzc05hbWUgPSAnZm9ybS1maWVsZCc7CiAgICBub3Rlc0ZpZWxkLnN0eWxlLm1hcmdpbiA9ICcxMnB4IDAnOwogICAgbm90ZXNGaWVsZC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2xhYmVsJywgJ1VwbG9hZCBub3RlcyAob3B0aW9uYWwpJykpOwogICAgY29uc3Qgbm90ZXNJbnB1dCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2lucHV0Jyk7CiAgICBub3Rlc0lucHV0LnR5cGUgPSAndGV4dCc7CiAgICBub3Rlc0lucHV0LmlkID0gJ3VwbG9hZE5vdGVzSW5wdXQnOwogICAgbm90ZXNJbnB1dC5wbGFjZWhvbGRlciA9ICdlLmcuICJXZWVrIDMgZXhwb3J0LCBpbmNsdWRlcyBjb3JyZWN0ZWQgVGlrVG9rIG51bWJlcnMiJzsKICAgIG5vdGVzRmllbGQuYXBwZW5kQ2hpbGQobm90ZXNJbnB1dCk7CiAgICBhcmVhLmFwcGVuZENoaWxkKG5vdGVzRmllbGQpOwoKICAgIGlmIChwLmlzc3Vlcy5sZW5ndGgpIHsKICAgICAgY29uc3QgaXNzdWVzVGl0bGUgPSB0ZXh0RWwoJ2RpdicsIGBSb3dzIHNraXBwZWQgb3IgZmxhZ2dlZCAoJHtwLmlzc3Vlcy5sZW5ndGh9KWAsICdzZWN0aW9uLXRpdGxlJyk7CiAgICAgIGNvbnN0IGlzc3Vlc0NhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgaXNzdWVzQ2FyZC5jbGFzc05hbWUgPSAnaXNzdWVzLWxpc3QnOwogICAgICBwLmlzc3Vlcy5mb3JFYWNoKChpc3N1ZSkgPT4gewogICAgICAgIGNvbnN0IHJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgICAgIHJvdy5jbGFzc05hbWUgPSAnaXNzdWUtcm93JzsKICAgICAgICBpZiAoaXNzdWUucm93TnVtYmVyKSByb3cuYXBwZW5kQ2hpbGQodGV4dEVsKCdzcGFuJywgYFJvdyAke2lzc3VlLnJvd051bWJlcn1gLCAncm93LW5vJykpOwogICAgICAgIHJvdy5hcHBlbmRDaGlsZChkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShpc3N1ZS5tZXNzYWdlKSk7CiAgICAgICAgaXNzdWVzQ2FyZC5hcHBlbmRDaGlsZChyb3cpOwogICAgICB9KTsKICAgICAgYXJlYS5hcHBlbmQoaXNzdWVzVGl0bGUsIGlzc3Vlc0NhcmQpOwogICAgfQoKICAgIGlmIChwLnNhbXBsZS5sZW5ndGgpIHsKICAgICAgY29uc3Qgc2FtcGxlVGl0bGUgPSB0ZXh0RWwoJ2RpdicsICdTYW1wbGUgb2YgcGFyc2VkIHJvd3MnLCAnc2VjdGlvbi10aXRsZScpOwogICAgICBjb25zdCB0YWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICAgIHRhYmxlLmNsYXNzTmFtZSA9ICdkYXRhLXRhYmxlJzsKICAgICAgdGFibGUuaW5uZXJIVE1MID0gJzx0aGVhZD48dHI+PHRoPkRhdGU8L3RoPjx0aD5DYXB0aW9uPC90aD48dGg+VHlwZTwvdGg+PHRoPkNhbXBhaWduPC90aD48dGg+UGxhdGZvcm1zPC90aD48L3RyPjwvdGhlYWQ+JzsKICAgICAgY29uc3QgdGJvZHkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0Ym9keScpOwogICAgICBwLnNhbXBsZS5mb3JFYWNoKChzKSA9PiB7CiAgICAgICAgY29uc3QgdHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgICAgIHRyLmFwcGVuZCgKICAgICAgICAgIHRleHRFbCgndGQnLCBGb3JtYXQuZGF0ZShzLnB1Ymxpc2hEYXRlKSksCiAgICAgICAgICB0ZXh0RWwoJ3RkJywgcy5jYXB0aW9uIHx8ICfigJQnKSwKICAgICAgICAgIHRleHRFbCgndGQnLCBzLmNvbnRlbnRUeXBlIHx8ICfigJQnKSwKICAgICAgICAgIHRleHRFbCgndGQnLCBzLmNhbXBhaWduVHlwZSB8fCAnVW5zcGVjaWZpZWQnKSwKICAgICAgICAgIHRleHRFbCgndGQnLCBzLnBsYXRmb3Jtcy5qb2luKCcsICcpKQogICAgICAgICk7CiAgICAgICAgdGJvZHkuYXBwZW5kQ2hpbGQodHIpOwogICAgICB9KTsKICAgICAgdGFibGUuYXBwZW5kQ2hpbGQodGJvZHkpOwogICAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIHdyYXAuY2xhc3NOYW1lID0gJ3RhYmxlLXNjcm9sbCc7CiAgICAgIHdyYXAuYXBwZW5kQ2hpbGQodGFibGUpOwogICAgICBhcmVhLmFwcGVuZChzYW1wbGVUaXRsZSwgd3JhcCk7CiAgICB9CgogICAgY29uc3QgYWN0aW9ucyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgYWN0aW9ucy5jbGFzc05hbWUgPSAnYnRuLXJvdyc7CiAgICBhY3Rpb25zLnN0eWxlLm1hcmdpblRvcCA9ICcxNnB4JzsKICAgIGNvbnN0IGNvbW1pdEJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgY29tbWl0QnRuLmNsYXNzTmFtZSA9ICdidG4gcHJpbWFyeSc7CiAgICBjb21taXRCdG4udGV4dENvbnRlbnQgPSBwLnZhbGlkUm93cyA+IDAgPyBgSW1wb3J0ICR7cC52YWxpZFJvd3N9IHJvdyhzKWAgOiAnTm90aGluZyB0byBpbXBvcnQnOwogICAgY29tbWl0QnRuLmRpc2FibGVkID0gcC52YWxpZFJvd3MgPT09IDA7CiAgICBjb21taXRCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBjb21taXQoY29tbWl0QnRuKSk7CiAgICBjb25zdCBjYW5jZWxCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTsKICAgIGNhbmNlbEJ0bi5jbGFzc05hbWUgPSAnYnRuJzsKICAgIGNhbmNlbEJ0bi50ZXh0Q29udGVudCA9ICdDYW5jZWwnOwogICAgY2FuY2VsQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4geyBjdXJyZW50UHJldmlldyA9IG51bGw7IHNoZWxsKCk7IH0pOwogICAgYWN0aW9ucy5hcHBlbmQoY29tbWl0QnRuLCBjYW5jZWxCdG4pOwogICAgYXJlYS5hcHBlbmRDaGlsZChhY3Rpb25zKTsKICB9CgogIGZ1bmN0aW9uIHN0YXRUaWxlKGxhYmVsLCB2YWx1ZSkgewogICAgY29uc3QgdGlsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgdGlsZS5jbGFzc05hbWUgPSAnc3RhdC10aWxlJzsKICAgIHRpbGUuYXBwZW5kKHRleHRFbCgnZGl2JywgbGFiZWwsICdzdGF0LWxhYmVsJyksIHRleHRFbCgnZGl2JywgU3RyaW5nKHZhbHVlKSwgJ3N0YXQtdmFsdWUnKSk7CiAgICByZXR1cm4gdGlsZTsKICB9CiAgZnVuY3Rpb24gYWN0aW9uU2VsZWN0KGRlZmF1bHRWYWwpIHsKICAgIGNvbnN0IHNlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NlbGVjdCcpOwogICAgW1snc2tpcCcsICdTa2lwIChrZWVwIGV4aXN0aW5nIHJlY29yZCB1bmNoYW5nZWQpJ10sIFsndXBkYXRlJywgJ1VwZGF0ZSBleGlzdGluZyByZWNvcmQnXSwgWydjcmVhdGUnLCAnQ3JlYXRlIGFzIGEgbmV3LCBzZXBhcmF0ZSByZWNvcmQnXV0uZm9yRWFjaCgoW3YsIGxdKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOyBvcHQudmFsdWUgPSB2OyBvcHQudGV4dENvbnRlbnQgPSBsOwogICAgICBpZiAodiA9PT0gZGVmYXVsdFZhbCkgb3B0LnNlbGVjdGVkID0gdHJ1ZTsKICAgICAgc2VsLmFwcGVuZENoaWxkKG9wdCk7CiAgICB9KTsKICAgIHJldHVybiBzZWw7CiAgfQoKICBhc3luYyBmdW5jdGlvbiBjb21taXQoYnRuKSB7CiAgICBidG4uZGlzYWJsZWQgPSB0cnVlOwogICAgYnRuLnRleHRDb250ZW50ID0gJ0ltcG9ydGluZ+KApic7CiAgICBjb25zdCBkZWZhdWx0RHVwbGljYXRlQWN0aW9uID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2RlZmF1bHREdXBsaWNhdGVBY3Rpb25TZWxlY3QnKT8udmFsdWUgfHwgJ3NraXAnOwogICAgY29uc3Qgbm90ZXMgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndXBsb2FkTm90ZXNJbnB1dCcpPy52YWx1ZSB8fCBudWxsOwogICAgdHJ5IHsKICAgICAgY29uc3QgcmVzdWx0ID0gYXdhaXQgQXBpLmNvbW1pdFVwbG9hZCh7CiAgICAgICAgZmlsZVBhdGg6IGN1cnJlbnRQcmV2aWV3LmZpbGVQYXRoLAogICAgICAgIG9yaWdpbmFsTmFtZTogY3VycmVudFByZXZpZXcub3JpZ2luYWxOYW1lLAogICAgICAgIGRlZmF1bHREdXBsaWNhdGVBY3Rpb24sCiAgICAgICAgZHVwbGljYXRlQWN0aW9uczogZHVwbGljYXRlQWN0aW9uT3ZlcnJpZGVzLAogICAgICAgIG5vdGVzLAogICAgICB9KTsKICAgICAgVG9hc3Quc2hvdygKICAgICAgICBgSW1wb3J0ZWQ6ICR7cmVzdWx0LmltcG9ydGVkUm93c30gbmV3LCAke3Jlc3VsdC51cGRhdGVkUm93c30gdXBkYXRlZCwgJHtyZXN1bHQuc2tpcHBlZFJvd3N9IHNraXBwZWQuYCwKICAgICAgICByZXN1bHQuZXJyb3JDb3VudCA+IDAgPyAnZXJyb3InIDogJ3N1Y2Nlc3MnCiAgICAgICk7CiAgICAgIGN1cnJlbnRQcmV2aWV3ID0gbnVsbDsKICAgICAgc2hlbGwoKTsKICAgICAgd2luZG93LmRpc3BhdGNoRXZlbnQobmV3IEN1c3RvbUV2ZW50KCdscnM6ZGF0YS11cGRhdGVkJykpOwogICAgfSBjYXRjaCAoZXJyKSB7CiAgICAgIFRvYXN0LnNob3coZXJyLm1lc3NhZ2UsICdlcnJvcicpOwogICAgICBidG4uZGlzYWJsZWQgPSBmYWxzZTsKICAgICAgYnRuLnRleHRDb250ZW50ID0gJ1JldHJ5IGltcG9ydCc7CiAgICB9CiAgfQoKICBmdW5jdGlvbiByZW5kZXIoKSB7CiAgICByb290ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3ZpZXctdXBsb2FkJyk7CiAgICBzaGVsbCgpOwogIH0KCiAgcmV0dXJuIHsgcmVuZGVyIH07Cn0pKCk7Cgpjb25zdCBIaXN0b3J5ID0gKCgpID0+IHsKICBsZXQgcm9vdDsKCiAgZnVuY3Rpb24gYmFkZ2VDbGFzcyhzdGF0dXMpIHsKICAgIGlmIChzdGF0dXMgPT09ICdzdWNjZXNzJykgcmV0dXJuICdzdWNjZXNzJzsKICAgIGlmIChzdGF0dXMgPT09ICdwYXJ0aWFsJykgcmV0dXJuICdwYXJ0aWFsJzsKICAgIHJldHVybiAnZmFpbGVkJzsKICB9CgogIGZ1bmN0aW9uIGJ1aWxkQmFja3VwQ2FyZCgpIHsKICAgIGNvbnN0IGNhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGNhcmQuY2xhc3NOYW1lID0gJ2NhcmQnOwogICAgY2FyZC5zdHlsZS5tYXJnaW5Cb3R0b20gPSAnMjBweCc7CiAgICBjb25zdCBoZWFkZXIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGhlYWRlci5jbGFzc05hbWUgPSAnY2FyZC1oZWFkZXInOwogICAgaGVhZGVyLmFwcGVuZENoaWxkKHRleHRFbCgnaDMnLCAnQmFja3VwICYgUmVzdG9yZScpKTsKICAgIGNhcmQuYXBwZW5kQ2hpbGQoaGVhZGVyKTsKICAgIGNhcmQuYXBwZW5kQ2hpbGQodGV4dEVsKAogICAgICAnZGl2JywKICAgICAgJ0Rvd25sb2FkIGEgZnVsbCBzbmFwc2hvdCBvZiB0aGUgZGF0YWJhc2UgYW55IHRpbWUuIFJlc3RvcmluZyByZXBsYWNlcyBBTEwgY3VycmVudCBkYXRhIHdpdGggdGhlIHVwbG9hZGVkIGJhY2t1cCBhbmQgcmVzdGFydHMgdGhlIHNlcnZlciDigJQgdGhpcyBjYW5ub3QgYmUgdW5kb25lLicsCiAgICAgICdtdXRlZCcKICAgICkpOwoKICAgIGNvbnN0IGFjdGlvbnMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGFjdGlvbnMuY2xhc3NOYW1lID0gJ2J0bi1yb3cnOwogICAgYWN0aW9ucy5zdHlsZS5tYXJnaW5Ub3AgPSAnMTRweCc7CgogICAgY29uc3QgZG93bmxvYWRCdG4gPSBpY29uQnRuKCdidG4gcHJpbWFyeScsICdkb3dubG9hZCcsICdEb3dubG9hZCBCYWNrdXAnKTsKICAgIGRvd25sb2FkQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4geyB3aW5kb3cubG9jYXRpb24uaHJlZiA9ICcvYXBpL2JhY2t1cC9leHBvcnQnOyB9KTsKCiAgICBjb25zdCByZXN0b3JlSW5wdXQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOwogICAgcmVzdG9yZUlucHV0LnR5cGUgPSAnZmlsZSc7CiAgICByZXN0b3JlSW5wdXQuYWNjZXB0ID0gJy5kYic7CiAgICByZXN0b3JlSW5wdXQuc3R5bGUuZGlzcGxheSA9ICdub25lJzsKCiAgICBjb25zdCByZXN0b3JlQnRuID0gaWNvbkJ0bignYnRuIGRhbmdlcicsICd1cGxvYWQnLCAnUmVzdG9yZSBmcm9tIEJhY2t1cCcpOwogICAgcmVzdG9yZUJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHJlc3RvcmVJbnB1dC5jbGljaygpKTsKCiAgICByZXN0b3JlSW5wdXQuYWRkRXZlbnRMaXN0ZW5lcignY2hhbmdlJywgYXN5bmMgKCkgPT4gewogICAgICBjb25zdCBmaWxlID0gcmVzdG9yZUlucHV0LmZpbGVzWzBdOwogICAgICBpZiAoIWZpbGUpIHJldHVybjsKICAgICAgY29uc3Qgc3VyZSA9IHdpbmRvdy5jb25maXJtKAogICAgICAgICdSZXN0b3Jpbmcgd2lsbCBSRVBMQUNFIGFsbCBjdXJyZW50IGRhdGEgd2l0aCB0aGlzIGJhY2t1cCBmaWxlIGFuZCByZXN0YXJ0IHRoZSBzZXJ2ZXIuIFRoaXMgY2Fubm90IGJlIHVuZG9uZS4gQ29udGludWU/JwogICAgICApOwogICAgICBpZiAoIXN1cmUpIHsKICAgICAgICByZXN0b3JlSW5wdXQudmFsdWUgPSAnJzsKICAgICAgICByZXR1cm47CiAgICAgIH0KICAgICAgcmVzdG9yZUJ0bi5kaXNhYmxlZCA9IHRydWU7CiAgICAgIHRyeSB7CiAgICAgICAgY29uc3QgZm9ybSA9IG5ldyBGb3JtRGF0YSgpOwogICAgICAgIGZvcm0uYXBwZW5kKCdmaWxlJywgZmlsZSk7CiAgICAgICAgY29uc3QgcmVzdWx0ID0gYXdhaXQgQXBpLnJlc3RvcmVCYWNrdXAoZm9ybSk7CiAgICAgICAgVG9hc3Quc2hvdyhyZXN1bHQubWVzc2FnZSB8fCAnQmFja3VwIHJlc3RvcmVkLiBUaGUgc2VydmVyIGlzIHJlc3RhcnRpbmcuJywgJ3N1Y2Nlc3MnKTsKICAgICAgfSBjYXRjaCAoZXJyKSB7CiAgICAgICAgVG9hc3Quc2hvdyhlcnIubWVzc2FnZSwgJ2Vycm9yJyk7CiAgICAgICAgcmVzdG9yZUJ0bi5kaXNhYmxlZCA9IGZhbHNlOwogICAgICB9IGZpbmFsbHkgewogICAgICAgIHJlc3RvcmVJbnB1dC52YWx1ZSA9ICcnOwogICAgICB9CiAgICB9KTsKCiAgICBhY3Rpb25zLmFwcGVuZChkb3dubG9hZEJ0biwgcmVzdG9yZUJ0biwgcmVzdG9yZUlucHV0KTsKICAgIGNhcmQuYXBwZW5kQ2hpbGQoYWN0aW9ucyk7CiAgICByZXR1cm4gY2FyZDsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIHJlbmRlcigpIHsKICAgIHJvb3QgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndmlldy1oaXN0b3J5Jyk7CiAgICByb290LmlubmVySFRNTCA9ICcnOwogICAgcm9vdC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsICdVcGxvYWQgaGlzdG9yeScsICdzZWN0aW9uLXRpdGxlJykpOwogICAgcm9vdC5hcHBlbmRDaGlsZChidWlsZEJhY2t1cENhcmQoKSk7CgogICAgY29uc3QgdXBsb2FkcyA9IGF3YWl0IEFwaS51cGxvYWRIaXN0b3J5KCk7CiAgICBpZiAoIXVwbG9hZHMubGVuZ3RoKSB7CiAgICAgIHJvb3QuYXBwZW5kQ2hpbGQoZW1wdHlTdGF0ZSh7CiAgICAgICAgaWNvbjogJ3VwbG9hZC1jbG91ZCcsCiAgICAgICAgdGl0bGU6ICdObyB1cGxvYWRzIHlldCcsCiAgICAgICAgbWVzc2FnZTogJ0ltcG9ydCB5b3VyIGZpcnN0IHdlZWtseSBleHBvcnQgdG8gc3RhcnQgc2VlaW5nIGRhdGEgYWNyb3NzIHRoZSBhcHAuJywKICAgICAgICBhY3Rpb25MYWJlbDogJ1VwbG9hZCBkYXRhJywKICAgICAgICBvbkFjdGlvbjogKCkgPT4gZG9jdW1lbnQucXVlcnlTZWxlY3RvcignLnRhYi1idG5bZGF0YS10YWI9InVwbG9hZCJdJyk/LmNsaWNrKCksCiAgICAgIH0pKTsKICAgICAgcmV0dXJuOwogICAgfQoKICAgIGNvbnN0IHRhYmxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGFibGUnKTsKICAgIHRhYmxlLmNsYXNzTmFtZSA9ICdkYXRhLXRhYmxlJzsKICAgIHRhYmxlLmlubmVySFRNTCA9ICc8dGhlYWQ+PHRyPjx0aD5GaWxlPC90aD48dGg+VXBsb2FkZWQ8L3RoPjx0aD5TdGF0dXM8L3RoPjx0aCBjbGFzcz0ibnVtIj5JbXBvcnRlZDwvdGg+PHRoIGNsYXNzPSJudW0iPlVwZGF0ZWQ8L3RoPjx0aCBjbGFzcz0ibnVtIj5Ta2lwcGVkPC90aD48dGggY2xhc3M9Im51bSI+RXJyb3JzPC90aD48dGg+V2Vla3M8L3RoPjx0aD5Ob3RlczwvdGg+PC90cj48L3RoZWFkPic7CiAgICBjb25zdCB0Ym9keSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3Rib2R5Jyk7CiAgICB1cGxvYWRzLmZvckVhY2goKHUpID0+IHsKICAgICAgY29uc3QgdHIgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgICB0ci5zdHlsZS5jdXJzb3IgPSAncG9pbnRlcic7CiAgICAgIGNvbnN0IGJhZGdlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgICBiYWRnZS5jbGFzc05hbWUgPSBgYmFkZ2UgJHtiYWRnZUNsYXNzKHUuc3RhdHVzKX1gOwogICAgICBiYWRnZS50ZXh0Q29udGVudCA9IHUuc3RhdHVzOwogICAgICBjb25zdCBzdGF0dXNUZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIHN0YXR1c1RkLmFwcGVuZENoaWxkKGJhZGdlKTsKICAgICAgdHIuYXBwZW5kKAogICAgICAgIHRleHRFbCgndGQnLCB1LmZpbGVuYW1lKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgdS51cGxvYWRlZF9hdCksCiAgICAgICAgc3RhdHVzVGQsCiAgICAgICAgdGV4dEVsKCd0ZCcsIFN0cmluZyh1LmltcG9ydGVkX3Jvd3MpLCAnbnVtJyksCiAgICAgICAgdGV4dEVsKCd0ZCcsIFN0cmluZyh1LnVwZGF0ZWRfcm93cyksICdudW0nKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgU3RyaW5nKHUuc2tpcHBlZF9yb3dzKSwgJ251bScpLAogICAgICAgIHRleHRFbCgndGQnLCBTdHJpbmcodS5lcnJvcl9jb3VudCksICdudW0nKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgdS53ZWVrc19hZmZlY3RlZC5tYXAoKHcpID0+IEZvcm1hdC5kYXRlKHcpKS5qb2luKCcsICcpIHx8ICfigJQnKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgdS5ub3RlcyB8fCAn4oCUJykKICAgICAgKTsKICAgICAgdHIuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB0b2dnbGVFcnJvcnModS5pZCwgdHIpKTsKICAgICAgdGJvZHkuYXBwZW5kQ2hpbGQodHIpOwogICAgfSk7CiAgICB0YWJsZS5hcHBlbmRDaGlsZCh0Ym9keSk7CiAgICBjb25zdCB3cmFwID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICB3cmFwLmNsYXNzTmFtZSA9ICd0YWJsZS1zY3JvbGwnOwogICAgd3JhcC5hcHBlbmRDaGlsZCh0YWJsZSk7CiAgICByb290LmFwcGVuZENoaWxkKHdyYXApOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gdG9nZ2xlRXJyb3JzKHVwbG9hZElkLCB0cikgewogICAgY29uc3QgZXhpc3RpbmcgPSB0ci5uZXh0RWxlbWVudFNpYmxpbmc7CiAgICBpZiAoZXhpc3RpbmcgJiYgZXhpc3RpbmcuY2xhc3NMaXN0LmNvbnRhaW5zKCdlcnJvci1sb2ctcm93JykpIHsKICAgICAgZXhpc3RpbmcucmVtb3ZlKCk7CiAgICAgIHJldHVybjsKICAgIH0KICAgIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJy5lcnJvci1sb2ctcm93JykuZm9yRWFjaCgoZWwpID0+IGVsLnJlbW92ZSgpKTsKICAgIGNvbnN0IGVycm9ycyA9IGF3YWl0IEFwaS51cGxvYWRFcnJvcnModXBsb2FkSWQpOwogICAgY29uc3QgbG9nUm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgIGxvZ1Jvdy5jbGFzc05hbWUgPSAnZXJyb3ItbG9nLXJvdyc7CiAgICBjb25zdCB0ZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICB0ZC5jb2xTcGFuID0gOTsKICAgIGlmICghZXJyb3JzLmxlbmd0aCkgewogICAgICB0ZC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsICdObyBpc3N1ZXMgbG9nZ2VkIGZvciB0aGlzIHVwbG9hZC4nLCAnbXV0ZWQnKSk7CiAgICB9IGVsc2UgewogICAgICBjb25zdCBsaXN0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICAgIGxpc3QuY2xhc3NOYW1lID0gJ2lzc3Vlcy1saXN0JzsKICAgICAgZXJyb3JzLmZvckVhY2goKGUpID0+IHsKICAgICAgICBjb25zdCByb3cgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgICByb3cuY2xhc3NOYW1lID0gJ2lzc3VlLXJvdyc7CiAgICAgICAgY29uc3QgYmFkZ2UgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICAgICAgYmFkZ2UuY2xhc3NOYW1lID0gYGJhZGdlICR7ZS5zZXZlcml0eX0tc2V2YDsKICAgICAgICBiYWRnZS50ZXh0Q29udGVudCA9IGUuc2V2ZXJpdHk7CiAgICAgICAgcm93LmFwcGVuZChiYWRnZSwgZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUoYCAke2Uucm93X251bWJlciA/IGBSb3cgJHtlLnJvd19udW1iZXJ9OiBgIDogJyd9JHtlLm1lc3NhZ2V9YCkpOwogICAgICAgIGxpc3QuYXBwZW5kQ2hpbGQocm93KTsKICAgICAgfSk7CiAgICAgIHRkLmFwcGVuZENoaWxkKGxpc3QpOwogICAgfQoKICAgIGNvbnN0IHJhd0J0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgcmF3QnRuLmNsYXNzTmFtZSA9ICdidG4nOwogICAgcmF3QnRuLnN0eWxlLm1hcmdpblRvcCA9ICcxMHB4JzsKICAgIHJhd0J0bi50ZXh0Q29udGVudCA9ICdWaWV3IGV2ZXJ5IHJhdyBzb3VyY2Ugcm93IGZyb20gdGhpcyB1cGxvYWQnOwogICAgcmF3QnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gbG9hZFJhd1Jvd3ModXBsb2FkSWQsIHJhd0J0bikpOwogICAgdGQuYXBwZW5kQ2hpbGQocmF3QnRuKTsKICAgIGNvbnN0IHJhd1dyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHJhd1dyYXAuaWQgPSBgcmF3V3JhcC0ke3VwbG9hZElkfWA7CiAgICB0ZC5hcHBlbmRDaGlsZChyYXdXcmFwKTsKCiAgICBsb2dSb3cuYXBwZW5kQ2hpbGQodGQpOwogICAgdHIuYWZ0ZXIobG9nUm93KTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIGxvYWRSYXdSb3dzKHVwbG9hZElkLCBidG4pIHsKICAgIGNvbnN0IHdyYXAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZChgcmF3V3JhcC0ke3VwbG9hZElkfWApOwogICAgaWYgKCF3cmFwKSByZXR1cm47CiAgICBpZiAod3JhcC5kYXRhc2V0LmxvYWRlZCkgewogICAgICB3cmFwLnN0eWxlLmRpc3BsYXkgPSB3cmFwLnN0eWxlLmRpc3BsYXkgPT09ICdub25lJyA/ICdibG9jaycgOiAnbm9uZSc7CiAgICAgIHJldHVybjsKICAgIH0KICAgIGJ0bi50ZXh0Q29udGVudCA9ICdMb2FkaW5n4oCmJzsKICAgIGNvbnN0IHsgcm93cywgdG90YWwgfSA9IGF3YWl0IEFwaS51cGxvYWRSYXdSb3dzKHVwbG9hZElkKTsKICAgIHdyYXAuZGF0YXNldC5sb2FkZWQgPSAnMSc7CiAgICBidG4udGV4dENvbnRlbnQgPSBgU2hvd2luZyAke3Jvd3MubGVuZ3RofSBvZiAke3RvdGFsfSByYXcgcm93KHMpYDsKCiAgICBjb25zdCBieVNoZWV0ID0gbmV3IE1hcCgpOwogICAgcm93cy5mb3JFYWNoKChyKSA9PiB7CiAgICAgIGlmICghYnlTaGVldC5oYXMoci5zaGVldF9uYW1lKSkgYnlTaGVldC5zZXQoci5zaGVldF9uYW1lLCBbXSk7CiAgICAgIGJ5U2hlZXQuZ2V0KHIuc2hlZXRfbmFtZSkucHVzaChyKTsKICAgIH0pOwoKICAgIHdyYXAuaW5uZXJIVE1MID0gJyc7CiAgICB3cmFwLnN0eWxlLm1hcmdpblRvcCA9ICcxMHB4JzsKICAgIGJ5U2hlZXQuZm9yRWFjaCgoc2hlZXRSb3dzLCBzaGVldE5hbWUpID0+IHsKICAgICAgd3JhcC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2RpdicsIGBTaGVldDogJHtzaGVldE5hbWV9ICgke3NoZWV0Um93cy5sZW5ndGh9IHJvdyhzKSlgLCAnc3RhdC1sYWJlbCcpKTsKICAgICAgY29uc3QgaGVhZGVycyA9IHNoZWV0Um93c1swXS5oZWFkZXJzOwogICAgICBjb25zdCB0YWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RhYmxlJyk7CiAgICAgIHRhYmxlLmNsYXNzTmFtZSA9ICdkYXRhLXRhYmxlJzsKICAgICAgY29uc3QgdGhlYWQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0cicpOwogICAgICB0aGVhZC5hcHBlbmQodGV4dEVsKCd0aCcsICdSb3cgIycpLCB0ZXh0RWwoJ3RoJywgJ0xpbmtlZCB0byBwb3N0JykpOwogICAgICBjb25zdCBjb2xDb3VudCA9IGhlYWRlcnMgPyBoZWFkZXJzLmxlbmd0aCA6IE1hdGgubWF4KC4uLnNoZWV0Um93cy5tYXAoKHIpID0+IHIucmF3Lmxlbmd0aCkpOwogICAgICBmb3IgKGxldCBpID0gMDsgaSA8IGNvbENvdW50OyBpICs9IDEpIHRoZWFkLmFwcGVuZENoaWxkKHRleHRFbCgndGgnLCBoZWFkZXJzICYmIGhlYWRlcnNbaV0gPyBTdHJpbmcoaGVhZGVyc1tpXSkgOiBgQ29sICR7aSArIDF9YCkpOwogICAgICBjb25zdCB0aGVhZFdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0aGVhZCcpOwogICAgICB0aGVhZFdyYXAuYXBwZW5kQ2hpbGQodGhlYWQpOwogICAgICBjb25zdCB0Ym9keSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3Rib2R5Jyk7CiAgICAgIHNoZWV0Um93cy5mb3JFYWNoKChyKSA9PiB7CiAgICAgICAgY29uc3QgdHIyID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndHInKTsKICAgICAgICB0cjIuYXBwZW5kKHRleHRFbCgndGQnLCBTdHJpbmcoci5yb3dfbnVtYmVyKSksIHRleHRFbCgndGQnLCByLnBvc3RfaWQgPyBgIyR7ci5wb3N0X2lkfWAgOiAn4oCUJykpOwogICAgICAgIGZvciAobGV0IGkgPSAwOyBpIDwgY29sQ291bnQ7IGkgKz0gMSkgewogICAgICAgICAgY29uc3QgdmFsID0gci5yYXdbaV07CiAgICAgICAgICB0cjIuYXBwZW5kQ2hpbGQodGV4dEVsKCd0ZCcsIHZhbCA9PT0gdW5kZWZpbmVkIHx8IHZhbCA9PT0gbnVsbCA/ICcnIDogU3RyaW5nKHZhbCkuc2xpY2UoMCwgNjApKSk7CiAgICAgICAgfQogICAgICAgIHRib2R5LmFwcGVuZENoaWxkKHRyMik7CiAgICAgIH0pOwogICAgICB0YWJsZS5hcHBlbmQodGhlYWRXcmFwLCB0Ym9keSk7CiAgICAgIGNvbnN0IHNjcm9sbFdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgc2Nyb2xsV3JhcC5jbGFzc05hbWUgPSAndGFibGUtc2Nyb2xsJzsKICAgICAgc2Nyb2xsV3JhcC5zdHlsZS5tYXJnaW5Cb3R0b20gPSAnMTZweCc7CiAgICAgIHNjcm9sbFdyYXAuYXBwZW5kQ2hpbGQodGFibGUpOwogICAgICB3cmFwLmFwcGVuZENoaWxkKHNjcm9sbFdyYXApOwogICAgfSk7CiAgfQoKICByZXR1cm4geyByZW5kZXIgfTsKfSkoKTsKCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBGb2xsb3dlcnMgRGF0YSB0YWI6IG1hbnVhbCB3ZWVrbHkgZm9sbG93ZXItY291bnQgZW50cnkgcGVyCiAgIHBsYXRmb3JtIOKAlCBlbnRpcmVseSBpbmRlcGVuZGVudCBvZiBzcHJlYWRzaGVldCB1cGxvYWRzIChpdHMgb3duCiAgIHRhYmxlLCBpdHMgb3duIEFQSSwgbmV2ZXIgdG91Y2hlZCBieSB0aGUgaW1wb3J0IHBpcGVsaW5lKS4gUG93ZXJzCiAgIEZvbGxvd2VyIEdyb3d0aCBjaGFydHMvY29tcGFyaXNvbnMgZWxzZXdoZXJlIGluIHRoZSBhcHAuCiAgID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwpjb25zdCBGb2xsb3dlcnMgPSAoKCkgPT4gewogIGxldCByb290OwogIGxldCBlZGl0aW5nSWQgPSBudWxsOyAvLyBub24tbnVsbCB3aGlsZSB0aGUgZm9ybSBpcyBlZGl0aW5nIGFuIGV4aXN0aW5nIGVudHJ5IHJhdGhlciB0aGFuIGFkZGluZyBhIG5ldyBvbmUKICBsZXQgc29ydFN0YXRlID0geyBrZXk6ICdlbnRyeV9kYXRlJywgZGlyOiAnZGVzYycsIHR5cGU6ICdzdHJpbmcnIH07CiAgbGV0IGN1cnJlbnRSb3dzID0gW107CgogIGZ1bmN0aW9uIGFsbFBsYXRmb3JtcygpIHsKICAgIHJldHVybiAod2luZG93Ll9fZmlsdGVyT3B0aW9uc0NhY2hlIHx8IHsgYWxsUGxhdGZvcm1zOiBbXSB9KS5hbGxQbGF0Zm9ybXMgfHwgW107CiAgfQoKICBmdW5jdGlvbiBwbGF0Zm9ybU1ldGFGb3IoaWQpIHsKICAgIHJldHVybiBhbGxQbGF0Zm9ybXMoKS5maW5kKChwKSA9PiBwLmlkID09PSBpZCkgfHwgeyBsYWJlbDogaWQsIGNvbG9yOiAnIzk5OScgfTsKICB9CgogIGZ1bmN0aW9uIHNoZWxsKCkgewogICAgcm9vdC5pbm5lckhUTUwgPSAnJzsKICAgIHJvb3QuYXBwZW5kQ2hpbGQodGV4dEVsKCdkaXYnLCAnRm9sbG93ZXJzIERhdGEgUmVjb3JkJywgJ3NlY3Rpb24tdGl0bGUnKSk7CiAgICByb290LmFwcGVuZENoaWxkKHRleHRFbCgKICAgICAgJ2RpdicsCiAgICAgICdNYW51YWxseSBsb2cgZWFjaCBwbGF0Zm9ybeKAmXMgdG90YWwgZm9sbG93ZXIgY291bnQgb25jZSBhIHdlZWsuIFRoaXMgaXMgaW5kZXBlbmRlbnQgb2Ygc3ByZWFkc2hlZXQgdXBsb2FkcyDigJQgaXQgcG93ZXJzIEZvbGxvd2VyIEdyb3d0aCBjaGFydHMgYW5kIGNvbXBhcmlzb25zIGVsc2V3aGVyZSBpbiB0aGUgYXBwLicsCiAgICAgICdtdXRlZCcKICAgICkpOwoKICAgIGNvbnN0IGZvcm1DYXJkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBmb3JtQ2FyZC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICBmb3JtQ2FyZC5zdHlsZS5tYXJnaW5Cb3R0b20gPSAnMjBweCc7CiAgICBmb3JtQ2FyZC5pZCA9ICdmb2xsb3dlcnNGb3JtQ2FyZCc7CiAgICByb290LmFwcGVuZENoaWxkKGZvcm1DYXJkKTsKCiAgICBjb25zdCB0YWJsZUNhcmQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRhYmxlQ2FyZC5jbGFzc05hbWUgPSAnY2FyZCc7CiAgICBjb25zdCB0YWJsZVdyYXAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIHRhYmxlV3JhcC5jbGFzc05hbWUgPSAndGFibGUtc2Nyb2xsJzsKICAgIHRhYmxlV3JhcC5pZCA9ICdmb2xsb3dlcnNUYWJsZVdyYXAnOwogICAgdGFibGVDYXJkLmFwcGVuZENoaWxkKHRhYmxlV3JhcCk7CiAgICByb290LmFwcGVuZENoaWxkKHRhYmxlQ2FyZCk7CgogICAgcmVuZGVyRm9ybSgpOwogIH0KCiAgZnVuY3Rpb24gcmVuZGVyRm9ybSgpIHsKICAgIGNvbnN0IGNhcmQgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZm9sbG93ZXJzRm9ybUNhcmQnKTsKICAgIGlmICghY2FyZCkgcmV0dXJuOwogICAgY2FyZC5pbm5lckhUTUwgPSAnJzsKICAgIGNvbnN0IGhlYWRlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgaGVhZGVyLmNsYXNzTmFtZSA9ICdjYXJkLWhlYWRlcic7CiAgICBoZWFkZXIuYXBwZW5kQ2hpbGQodGV4dEVsKCdoMycsIGVkaXRpbmdJZCAhPT0gbnVsbCA/ICdFZGl0IGVudHJ5JyA6ICdBZGQgYSB3ZWVrbHkgZW50cnknKSk7CiAgICBjYXJkLmFwcGVuZENoaWxkKGhlYWRlcik7CgogICAgY29uc3QgZ3JpZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZ3JpZC5jbGFzc05hbWUgPSAnZm9ybS1ncmlkJzsKCiAgICBjb25zdCBwbGF0Zm9ybUZpZWxkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBwbGF0Zm9ybUZpZWxkLmNsYXNzTmFtZSA9ICdmb3JtLWZpZWxkJzsKICAgIHBsYXRmb3JtRmllbGQuYXBwZW5kQ2hpbGQodGV4dEVsKCdsYWJlbCcsICdQbGF0Zm9ybScpKTsKICAgIGNvbnN0IHBsYXRmb3JtU2VsZWN0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc2VsZWN0Jyk7CiAgICBwbGF0Zm9ybVNlbGVjdC5pZCA9ICdmb2xsb3dlcnNQbGF0Zm9ybUlucHV0JzsKICAgIGFsbFBsYXRmb3JtcygpLmZvckVhY2goKHApID0+IHsKICAgICAgY29uc3Qgb3B0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7CiAgICAgIG9wdC52YWx1ZSA9IHAuaWQ7CiAgICAgIG9wdC50ZXh0Q29udGVudCA9IHAubGFiZWw7CiAgICAgIHBsYXRmb3JtU2VsZWN0LmFwcGVuZENoaWxkKG9wdCk7CiAgICB9KTsKICAgIHBsYXRmb3JtRmllbGQuYXBwZW5kQ2hpbGQocGxhdGZvcm1TZWxlY3QpOwoKICAgIGNvbnN0IGRhdGVGaWVsZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgZGF0ZUZpZWxkLmNsYXNzTmFtZSA9ICdmb3JtLWZpZWxkJzsKICAgIGRhdGVGaWVsZC5hcHBlbmRDaGlsZCh0ZXh0RWwoJ2xhYmVsJywgJ1dlZWsgLyBEYXRlJykpOwogICAgY29uc3QgZGF0ZUlucHV0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsKICAgIGRhdGVJbnB1dC50eXBlID0gJ2RhdGUnOwogICAgZGF0ZUlucHV0LmlkID0gJ2ZvbGxvd2Vyc0RhdGVJbnB1dCc7CiAgICBkYXRlRmllbGQuYXBwZW5kQ2hpbGQoZGF0ZUlucHV0KTsKCiAgICBjb25zdCBjb3VudEZpZWxkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBjb3VudEZpZWxkLmNsYXNzTmFtZSA9ICdmb3JtLWZpZWxkJzsKICAgIGNvdW50RmllbGQuYXBwZW5kQ2hpbGQodGV4dEVsKCdsYWJlbCcsICdGb2xsb3dlcnMgQ291bnQnKSk7CiAgICBjb25zdCBjb3VudElucHV0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsKICAgIGNvdW50SW5wdXQudHlwZSA9ICdudW1iZXInOwogICAgY291bnRJbnB1dC5taW4gPSAnMCc7CiAgICBjb3VudElucHV0LnN0ZXAgPSAnMSc7CiAgICBjb3VudElucHV0LmlkID0gJ2ZvbGxvd2Vyc0NvdW50SW5wdXQnOwogICAgY291bnRGaWVsZC5hcHBlbmRDaGlsZChjb3VudElucHV0KTsKCiAgICBncmlkLmFwcGVuZChwbGF0Zm9ybUZpZWxkLCBkYXRlRmllbGQsIGNvdW50RmllbGQpOwogICAgY2FyZC5hcHBlbmRDaGlsZChncmlkKTsKCiAgICBjb25zdCBlZGl0Um93ID0gZWRpdGluZ0lkICE9PSBudWxsID8gY3VycmVudFJvd3MuZmluZCgocikgPT4gci5pZCA9PT0gZWRpdGluZ0lkKSA6IG51bGw7CiAgICBpZiAoZWRpdFJvdykgewogICAgICBwbGF0Zm9ybVNlbGVjdC52YWx1ZSA9IGVkaXRSb3cucGxhdGZvcm07CiAgICAgIGRhdGVJbnB1dC52YWx1ZSA9IGVkaXRSb3cuZW50cnlfZGF0ZTsKICAgICAgY291bnRJbnB1dC52YWx1ZSA9IFN0cmluZyhlZGl0Um93LmZvbGxvd2Vyc19jb3VudCk7CiAgICB9IGVsc2UgewogICAgICBkYXRlSW5wdXQudmFsdWUgPSBuZXcgRGF0ZSgpLnRvSVNPU3RyaW5nKCkuc2xpY2UoMCwgMTApOwogICAgfQoKICAgIGNvbnN0IGFjdGlvbnMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgIGFjdGlvbnMuY2xhc3NOYW1lID0gJ21vZGFsLWFjdGlvbnMnOwogICAgY29uc3QgZXJyb3JFbCA9IHRleHRFbCgnc3BhbicsICcnLCAnbXV0ZWQnKTsKICAgIGVycm9yRWwuaWQgPSAnZm9sbG93ZXJzRm9ybUVycm9yJzsKICAgIGVycm9yRWwuc3R5bGUuY29sb3IgPSAndmFyKC0tc3RhdHVzLWNyaXRpY2FsKSc7CgogICAgY29uc3QgYnRuUm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICBidG5Sb3cuY2xhc3NOYW1lID0gJ2J0bi1yb3cnOwogICAgY29uc3Qgc2F2ZUJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgc2F2ZUJ0bi5jbGFzc05hbWUgPSAnYnRuIHByaW1hcnknOwogICAgc2F2ZUJ0bi50ZXh0Q29udGVudCA9IGVkaXRpbmdJZCAhPT0gbnVsbCA/ICdTYXZlIGNoYW5nZXMnIDogJ0FkZCBlbnRyeSc7CiAgICBzYXZlQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gc3VibWl0Rm9ybShzYXZlQnRuKSk7CiAgICBidG5Sb3cuYXBwZW5kQ2hpbGQoc2F2ZUJ0bik7CiAgICBpZiAoZWRpdGluZ0lkICE9PSBudWxsKSB7CiAgICAgIGNvbnN0IGNhbmNlbEJ0biA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpOwogICAgICBjYW5jZWxCdG4uY2xhc3NOYW1lID0gJ2J0bic7CiAgICAgIGNhbmNlbEJ0bi50ZXh0Q29udGVudCA9ICdDYW5jZWwnOwogICAgICBjYW5jZWxCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7IGVkaXRpbmdJZCA9IG51bGw7IHJlbmRlckZvcm0oKTsgfSk7CiAgICAgIGJ0blJvdy5hcHBlbmRDaGlsZChjYW5jZWxCdG4pOwogICAgfQogICAgYWN0aW9ucy5hcHBlbmQoZXJyb3JFbCwgYnRuUm93KTsKICAgIGNhcmQuYXBwZW5kQ2hpbGQoYWN0aW9ucyk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiBzdWJtaXRGb3JtKGJ0bikgewogICAgY29uc3QgZXJyb3JFbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmb2xsb3dlcnNGb3JtRXJyb3InKTsKICAgIGVycm9yRWwudGV4dENvbnRlbnQgPSAnJzsKICAgIGNvbnN0IHBsYXRmb3JtID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZvbGxvd2Vyc1BsYXRmb3JtSW5wdXQnKS52YWx1ZTsKICAgIGNvbnN0IGVudHJ5RGF0ZSA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmb2xsb3dlcnNEYXRlSW5wdXQnKS52YWx1ZTsKICAgIGNvbnN0IGZvbGxvd2Vyc0NvdW50ID0gTnVtYmVyKGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmb2xsb3dlcnNDb3VudElucHV0JykudmFsdWUpOwogICAgYnRuLmRpc2FibGVkID0gdHJ1ZTsKICAgIHRyeSB7CiAgICAgIGlmIChlZGl0aW5nSWQgIT09IG51bGwpIHsKICAgICAgICBhd2FpdCBBcGkudXBkYXRlRm9sbG93ZXJzKGVkaXRpbmdJZCwgeyBwbGF0Zm9ybSwgZW50cnlEYXRlLCBmb2xsb3dlcnNDb3VudCB9KTsKICAgICAgICBUb2FzdC5zaG93KCdFbnRyeSB1cGRhdGVkLicsICdzdWNjZXNzJyk7CiAgICAgIH0gZWxzZSB7CiAgICAgICAgYXdhaXQgQXBpLnNhdmVGb2xsb3dlcnMoeyBwbGF0Zm9ybSwgZW50cnlEYXRlLCBmb2xsb3dlcnNDb3VudCB9KTsKICAgICAgICBUb2FzdC5zaG93KCdFbnRyeSBzYXZlZC4nLCAnc3VjY2VzcycpOwogICAgICB9CiAgICAgIGVkaXRpbmdJZCA9IG51bGw7CiAgICAgIGF3YWl0IGxvYWQoKTsKICAgICAgd2luZG93LmRpc3BhdGNoRXZlbnQobmV3IEN1c3RvbUV2ZW50KCdscnM6ZGF0YS11cGRhdGVkJykpOwogICAgfSBjYXRjaCAoZXJyKSB7CiAgICAgIGVycm9yRWwudGV4dENvbnRlbnQgPSBlcnIubWVzc2FnZTsKICAgICAgYnRuLmRpc2FibGVkID0gZmFsc2U7CiAgICB9CiAgfQoKICBmdW5jdGlvbiBzdGFydEVkaXQocm93KSB7CiAgICBlZGl0aW5nSWQgPSByb3cuaWQ7CiAgICByZW5kZXJGb3JtKCk7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZm9sbG93ZXJzRm9ybUNhcmQnKS5zY3JvbGxJbnRvVmlldyh7IGJlaGF2aW9yOiAnc21vb3RoJywgYmxvY2s6ICdzdGFydCcgfSk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiBoYW5kbGVEZWxldGUocm93KSB7CiAgICBjb25zdCBzdXJlID0gd2luZG93LmNvbmZpcm0oYERlbGV0ZSB0aGUgJHtwbGF0Zm9ybU1ldGFGb3Iocm93LnBsYXRmb3JtKS5sYWJlbH0gZW50cnkgZm9yICR7Rm9ybWF0LmRhdGUocm93LmVudHJ5X2RhdGUpfT9gKTsKICAgIGlmICghc3VyZSkgcmV0dXJuOwogICAgdHJ5IHsKICAgICAgYXdhaXQgQXBpLmRlbGV0ZUZvbGxvd2Vycyhyb3cuaWQpOwogICAgICBUb2FzdC5zaG93KCdFbnRyeSBkZWxldGVkLicsICdzdWNjZXNzJyk7CiAgICAgIGlmIChlZGl0aW5nSWQgPT09IHJvdy5pZCkgZWRpdGluZ0lkID0gbnVsbDsKICAgICAgYXdhaXQgbG9hZCgpOwogICAgICB3aW5kb3cuZGlzcGF0Y2hFdmVudChuZXcgQ3VzdG9tRXZlbnQoJ2xyczpkYXRhLXVwZGF0ZWQnKSk7CiAgICB9IGNhdGNoIChlcnIpIHsKICAgICAgVG9hc3Quc2hvdyhlcnIubWVzc2FnZSwgJ2Vycm9yJyk7CiAgICB9CiAgfQoKICBmdW5jdGlvbiBzb3J0YWJsZUhlYWRlcihsYWJlbCwga2V5LCB0eXBlKSB7CiAgICBjb25zdCB0aCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RoJyk7CiAgICBpZiAodHlwZSA9PT0gJ251bWJlcicpIHRoLmNsYXNzTmFtZSA9ICdudW0nOwogICAgdGguY2xhc3NMaXN0LmFkZCgnc29ydGFibGUtdGgnKTsKICAgIGNvbnN0IGlzQWN0aXZlID0gc29ydFN0YXRlLmtleSA9PT0ga2V5OwogICAgdGguYXBwZW5kQ2hpbGQoZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUobGFiZWwpKTsKICAgIHRoLmFwcGVuZENoaWxkKHRleHRFbCgnc3BhbicsIGlzQWN0aXZlID8gKHNvcnRTdGF0ZS5kaXIgPT09ICdhc2MnID8gJyDihpEnIDogJyDihpMnKSA6ICcg4oaVJywgJ3NvcnQtYXJyb3cnKSk7CiAgICB0aC5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsKICAgICAgc29ydFN0YXRlID0geyBrZXksIGRpcjogc29ydFN0YXRlLmtleSA9PT0ga2V5ICYmIHNvcnRTdGF0ZS5kaXIgPT09ICdhc2MnID8gJ2Rlc2MnIDogJ2FzYycsIHR5cGUgfTsKICAgICAgcmVuZGVyVGFibGUoKTsKICAgIH0pOwogICAgcmV0dXJuIHRoOwogIH0KCiAgZnVuY3Rpb24gc29ydGVkUm93cygpIHsKICAgIGNvbnN0IHsga2V5LCBkaXIsIHR5cGUgfSA9IHNvcnRTdGF0ZTsKICAgIGNvbnN0IGZhY3RvciA9IGRpciA9PT0gJ2FzYycgPyAxIDogLTE7CiAgICByZXR1cm4gWy4uLmN1cnJlbnRSb3dzXS5zb3J0KChhLCBiKSA9PiB7CiAgICAgIGNvbnN0IGF2ID0gYVtrZXldOwogICAgICBjb25zdCBidiA9IGJba2V5XTsKICAgICAgaWYgKGF2ID09PSBudWxsIHx8IGF2ID09PSB1bmRlZmluZWQpIHJldHVybiAxOwogICAgICBpZiAoYnYgPT09IG51bGwgfHwgYnYgPT09IHVuZGVmaW5lZCkgcmV0dXJuIC0xOwogICAgICBpZiAodHlwZSA9PT0gJ251bWJlcicpIHJldHVybiAoYXYgLSBidikgKiBmYWN0b3I7CiAgICAgIHJldHVybiBTdHJpbmcoYXYpLmxvY2FsZUNvbXBhcmUoU3RyaW5nKGJ2KSkgKiBmYWN0b3I7CiAgICB9KTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlclRhYmxlKCkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmb2xsb3dlcnNUYWJsZVdyYXAnKTsKICAgIGlmICghd3JhcCkgcmV0dXJuOwogICAgaWYgKCFjdXJyZW50Um93cy5sZW5ndGgpIHsKICAgICAgd3JhcC5pbm5lckhUTUwgPSAnJzsKICAgICAgd3JhcC5hcHBlbmRDaGlsZChlbXB0eVN0YXRlKHsKICAgICAgICBpY29uOiAndXNlcnMnLAogICAgICAgIHRpdGxlOiAnTm8gZm9sbG93ZXIgZW50cmllcyB5ZXQnLAogICAgICAgIG1lc3NhZ2U6ICdBZGQgeW91ciBmaXJzdCB3ZWVrbHkgZm9sbG93ZXIgY291bnQgYWJvdmUgZm9yIGFueSBwbGF0Zm9ybS4nLAogICAgICB9KSk7CiAgICAgIHJldHVybjsKICAgIH0KICAgIGNvbnN0IHRhYmxlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGFibGUnKTsKICAgIHRhYmxlLmNsYXNzTmFtZSA9ICdkYXRhLXRhYmxlJzsKICAgIGNvbnN0IHRoZWFkID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgndGhlYWQnKTsKICAgIGNvbnN0IGhlYWRUciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICBoZWFkVHIuYXBwZW5kKAogICAgICBzb3J0YWJsZUhlYWRlcignUGxhdGZvcm0nLCAncGxhdGZvcm0nLCAnc3RyaW5nJyksCiAgICAgIHNvcnRhYmxlSGVhZGVyKCdXZWVrIC8gRGF0ZScsICdlbnRyeV9kYXRlJywgJ3N0cmluZycpLAogICAgICBzb3J0YWJsZUhlYWRlcignRm9sbG93ZXJzIENvdW50JywgJ2ZvbGxvd2Vyc19jb3VudCcsICdudW1iZXInKSwKICAgICAgdGV4dEVsKCd0aCcsICdMYXN0IFVwZGF0ZWQnKSwKICAgICAgdGV4dEVsKCd0aCcsICdBY3Rpb25zJykKICAgICk7CiAgICB0aGVhZC5hcHBlbmRDaGlsZChoZWFkVHIpOwogICAgY29uc3QgdGJvZHkgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0Ym9keScpOwogICAgc29ydGVkUm93cygpLmZvckVhY2goKHJvdykgPT4gewogICAgICBjb25zdCB0ciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RyJyk7CiAgICAgIGNvbnN0IHBsYXRmb3JtVGQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCd0ZCcpOwogICAgICBjb25zdCBtZXRhID0gcGxhdGZvcm1NZXRhRm9yKHJvdy5wbGF0Zm9ybSk7CiAgICAgIGNvbnN0IHBpbGwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICAgIHBpbGwuY2xhc3NOYW1lID0gJ3BsYXRmb3JtLXBpbGwnOwogICAgICBjb25zdCBkb3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdzcGFuJyk7CiAgICAgIGRvdC5jbGFzc05hbWUgPSAncGxhdGZvcm0tZG90JzsKICAgICAgZG90LnN0eWxlLmJhY2tncm91bmQgPSBtZXRhLmNvbG9yOwogICAgICBwaWxsLmFwcGVuZChkb3QsIGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKG1ldGEubGFiZWwpKTsKICAgICAgcGxhdGZvcm1UZC5hcHBlbmRDaGlsZChwaWxsKTsKCiAgICAgIGNvbnN0IGFjdGlvbnNUZCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3RkJyk7CiAgICAgIGNvbnN0IHJvd0FjdGlvbnMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTsKICAgICAgcm93QWN0aW9ucy5jbGFzc05hbWUgPSAncm93LWFjdGlvbnMnOwogICAgICBjb25zdCBlZGl0QnRuID0gaWNvbkJ0bignYnRuJywgJ3BlbmNpbCcsICdFZGl0Jyk7CiAgICAgIGVkaXRCdG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiBzdGFydEVkaXQocm93KSk7CiAgICAgIGNvbnN0IGRlbGV0ZUJ0biA9IGljb25CdG4oJ2J0biBkYW5nZXInLCAndHJhc2gtMicsICdEZWxldGUnKTsKICAgICAgZGVsZXRlQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gaGFuZGxlRGVsZXRlKHJvdykpOwogICAgICByb3dBY3Rpb25zLmFwcGVuZChlZGl0QnRuLCBkZWxldGVCdG4pOwogICAgICBhY3Rpb25zVGQuYXBwZW5kQ2hpbGQocm93QWN0aW9ucyk7CgogICAgICB0ci5hcHBlbmQoCiAgICAgICAgcGxhdGZvcm1UZCwKICAgICAgICB0ZXh0RWwoJ3RkJywgRm9ybWF0LmRhdGUocm93LmVudHJ5X2RhdGUpKSwKICAgICAgICB0ZXh0RWwoJ3RkJywgRm9ybWF0Lm51bWJlcihyb3cuZm9sbG93ZXJzX2NvdW50KSwgJ251bScpLAogICAgICAgIHRleHRFbCgndGQnLCByb3cudXBkYXRlZF9hdCksCiAgICAgICAgYWN0aW9uc1RkCiAgICAgICk7CiAgICAgIHRib2R5LmFwcGVuZENoaWxkKHRyKTsKICAgIH0pOwogICAgdGFibGUuYXBwZW5kKHRoZWFkLCB0Ym9keSk7CiAgICB3cmFwLmlubmVySFRNTCA9ICcnOwogICAgd3JhcC5hcHBlbmRDaGlsZCh0YWJsZSk7CiAgfQoKICBhc3luYyBmdW5jdGlvbiBsb2FkKCkgewogICAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmb2xsb3dlcnNUYWJsZVdyYXAnKTsKICAgIGlmICh3cmFwKSB7IHdyYXAuaW5uZXJIVE1MID0gJyc7IHdyYXAuYXBwZW5kQ2hpbGQoc2tlbGV0b25Sb3dzKDQpKTsgfQogICAgY3VycmVudFJvd3MgPSBhd2FpdCBBcGkubGlzdEZvbGxvd2Vycyh7fSk7CiAgICByZW5kZXJUYWJsZSgpOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gcmVuZGVyKCkgewogICAgcm9vdCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd2aWV3LWZvbGxvd2VycycpOwogICAgZWRpdGluZ0lkID0gbnVsbDsKICAgIHNoZWxsKCk7CiAgICBhd2FpdCBsb2FkKCk7CiAgfQoKICByZXR1cm4geyByZW5kZXIgfTsKfSkoKTsKCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBBcHAgYm9vdHN0cmFwOiB0YWIgcm91dGluZywgZmlsdGVyIGJhciB3aXJpbmcsIHRoZW1lIHRvZ2dsZS4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCigoKSA9PiB7CiAgY29uc3QgVklFV1MgPSB7CiAgICBkYXNoYm9hcmQ6IERhc2hib2FyZCwKICAgIHJlY29yZHM6IFJlY29yZHMsCiAgICBmb2xsb3dlcnM6IEZvbGxvd2VycywKICAgIGNvbXBhcmlzb246IENvbXBhcmlzb24sCiAgICB1cGxvYWQ6IFVwbG9hZCwKICAgIGhpc3Rvcnk6IEhpc3RvcnksCiAgfTsKCiAgbGV0IGFjdGl2ZVRhYiA9ICdkYXNoYm9hcmQnOwoKICBmdW5jdGlvbiBzd2l0Y2hUYWIodGFiKSB7CiAgICBhY3RpdmVUYWIgPSB0YWI7CiAgICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcudGFiLWJ0bicpLmZvckVhY2goKGJ0bikgPT4gewogICAgICBjb25zdCBpc0FjdGl2ZSA9IGJ0bi5kYXRhc2V0LnRhYiA9PT0gdGFiOwogICAgICBidG4uY2xhc3NMaXN0LnRvZ2dsZSgnaXMtYWN0aXZlJywgaXNBY3RpdmUpOwogICAgICBidG4uc2V0QXR0cmlidXRlKCdhcmlhLXNlbGVjdGVkJywgU3RyaW5nKGlzQWN0aXZlKSk7CiAgICB9KTsKICAgIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJy52aWV3JykuZm9yRWFjaCgodmlldykgPT4gewogICAgICB2aWV3LmNsYXNzTGlzdC50b2dnbGUoJ2lzLWFjdGl2ZScsIHZpZXcuaWQgPT09IGB2aWV3LSR7dGFifWApOwogICAgfSk7CiAgICAvLyBGaWx0ZXJzIGFwcGx5IHRvIERhc2hib2FyZCBhbmQgRGF0YSBSZWNvcmRzIChDb21wYXJpc29ucyBoYXMgaXRzIG93biByYW5nZSBjb250cm9scykuCiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmlsdGVyQmFyJykuc3R5bGUuZGlzcGxheSA9ICh0YWIgPT09ICdkYXNoYm9hcmQnIHx8IHRhYiA9PT0gJ3JlY29yZHMnKSA/ICdmbGV4JyA6ICdub25lJzsKICAgIHJlbmRlckFjdGl2ZVZpZXcoKTsKICB9CgogIGZ1bmN0aW9uIHJlbmRlckFjdGl2ZVZpZXcoKSB7CiAgICBjb25zdCB2aWV3ID0gVklFV1NbYWN0aXZlVGFiXTsKICAgIGlmICh2aWV3ICYmIHZpZXcucmVuZGVyKSB2aWV3LnJlbmRlcigpOwogIH0KCiAgYXN5bmMgZnVuY3Rpb24gbG9hZEZpbHRlck9wdGlvbnMoKSB7CiAgICBjb25zdCBvcHRpb25zID0gYXdhaXQgQXBpLmZpbHRlck9wdGlvbnMoKTsKICAgIHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZSA9IG9wdGlvbnM7CgogICAgY29uc3QgcGxhdGZvcm1TZWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmlsdGVyUGxhdGZvcm0nKTsKICAgIHBsYXRmb3JtU2VsLmxlbmd0aCA9IDE7CiAgICBvcHRpb25zLnBsYXRmb3Jtcy5mb3JFYWNoKChwKSA9PiB7CiAgICAgIGNvbnN0IG9wdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ29wdGlvbicpOwogICAgICBvcHQudmFsdWUgPSBwLmlkOwogICAgICBvcHQudGV4dENvbnRlbnQgPSBwLmxhYmVsOwogICAgICBwbGF0Zm9ybVNlbC5hcHBlbmRDaGlsZChvcHQpOwogICAgfSk7CgogICAgY29uc3QgY2FtcGFpZ25TZWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmlsdGVyQ2FtcGFpZ24nKTsKICAgIGNhbXBhaWduU2VsLmxlbmd0aCA9IDE7CiAgICBvcHRpb25zLmNhbXBhaWduVHlwZXMuZm9yRWFjaCgoYykgPT4gewogICAgICBjb25zdCBvcHQgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdvcHRpb24nKTsKICAgICAgb3B0LnZhbHVlID0gYzsKICAgICAgb3B0LnRleHRDb250ZW50ID0gYzsKICAgICAgY2FtcGFpZ25TZWwuYXBwZW5kQ2hpbGQob3B0KTsKICAgIH0pOwoKICAgIGNvbnN0IGNvbnRlbnRTZWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmlsdGVyQ29udGVudFR5cGUnKTsKICAgIGNvbnRlbnRTZWwubGVuZ3RoID0gMTsKICAgIG9wdGlvbnMuY29udGVudFR5cGVzLmZvckVhY2goKGMpID0+IHsKICAgICAgY29uc3Qgb3B0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnb3B0aW9uJyk7CiAgICAgIG9wdC52YWx1ZSA9IGM7CiAgICAgIG9wdC50ZXh0Q29udGVudCA9IGM7CiAgICAgIGNvbnRlbnRTZWwuYXBwZW5kQ2hpbGQob3B0KTsKICAgIH0pOwogIH0KCiAgZnVuY3Rpb24gd2lyZUZpbHRlckJhcigpIHsKICAgIGNvbnN0IGRhdGVGcm9tID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZpbHRlckRhdGVGcm9tJyk7CiAgICBjb25zdCBkYXRlVG8gPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmlsdGVyRGF0ZVRvJyk7CiAgICBjb25zdCBwbGF0Zm9ybSA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmaWx0ZXJQbGF0Zm9ybScpOwogICAgY29uc3QgY2FtcGFpZ24gPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmlsdGVyQ2FtcGFpZ24nKTsKICAgIGNvbnN0IGNvbnRlbnRUeXBlID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZpbHRlckNvbnRlbnRUeXBlJyk7CiAgICBjb25zdCBmID0gU3RhdGUuZ2V0RmlsdGVycygpOwogICAgZGF0ZUZyb20udmFsdWUgPSBmLmRhdGVGcm9tOwogICAgZGF0ZVRvLnZhbHVlID0gZi5kYXRlVG87CgogICAgZnVuY3Rpb24gYXBwbHkoKSB7CiAgICAgIFN0YXRlLnNldEZpbHRlcnMoewogICAgICAgIGRhdGVGcm9tOiBkYXRlRnJvbS52YWx1ZSwKICAgICAgICBkYXRlVG86IGRhdGVUby52YWx1ZSwKICAgICAgICBwbGF0Zm9ybTogcGxhdGZvcm0udmFsdWUsCiAgICAgICAgY2FtcGFpZ25UeXBlOiBjYW1wYWlnbi52YWx1ZSwKICAgICAgICBjb250ZW50VHlwZTogY29udGVudFR5cGUudmFsdWUsCiAgICAgIH0pOwogICAgfQogICAgW2RhdGVGcm9tLCBkYXRlVG8sIHBsYXRmb3JtLCBjYW1wYWlnbiwgY29udGVudFR5cGVdLmZvckVhY2goKGVsKSA9PiBlbC5hZGRFdmVudExpc3RlbmVyKCdjaGFuZ2UnLCBhcHBseSkpOwoKICAgIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJyNmaWx0ZXJQcmVzZXRzIGJ1dHRvbicpLmZvckVhY2goKGJ0bikgPT4gewogICAgICBidG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7CiAgICAgICAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnI2ZpbHRlclByZXNldHMgYnV0dG9uJykuZm9yRWFjaCgoYikgPT4gYi5jbGFzc0xpc3QucmVtb3ZlKCdpcy1hY3RpdmUnKSk7CiAgICAgICAgYnRuLmNsYXNzTGlzdC5hZGQoJ2lzLWFjdGl2ZScpOwogICAgICAgIGNvbnN0IHByZXNldCA9IGJ0bi5kYXRhc2V0LnByZXNldDsKICAgICAgICBjb25zdCB0b2RheSA9IG5ldyBEYXRlKCk7CiAgICAgICAgY29uc3QgdG8gPSB0b2RheS50b0lTT1N0cmluZygpLnNsaWNlKDAsIDEwKTsKICAgICAgICBsZXQgZnJvbTsKICAgICAgICBpZiAocHJlc2V0ID09PSAnYWxsJykgewogICAgICAgICAgY29uc3QgbWluID0gKHdpbmRvdy5fX2ZpbHRlck9wdGlvbnNDYWNoZSAmJiB3aW5kb3cuX19maWx0ZXJPcHRpb25zQ2FjaGUuZGF0ZVJhbmdlLm1pbikgfHwgdG87CiAgICAgICAgICBmcm9tID0gbWluOwogICAgICAgIH0gZWxzZSB7CiAgICAgICAgICBjb25zdCBkID0gbmV3IERhdGUodG9kYXkpOwogICAgICAgICAgZC5zZXREYXRlKGQuZ2V0RGF0ZSgpIC0gKE51bWJlcihwcmVzZXQpIC0gMSkpOwogICAgICAgICAgZnJvbSA9IGQudG9JU09TdHJpbmcoKS5zbGljZSgwLCAxMCk7CiAgICAgICAgfQogICAgICAgIGRhdGVGcm9tLnZhbHVlID0gZnJvbTsKICAgICAgICBkYXRlVG8udmFsdWUgPSB0bzsKICAgICAgICBhcHBseSgpOwogICAgICB9KTsKICAgIH0pOwoKICAgIFN0YXRlLm9uQ2hhbmdlKCgpID0+IHsKICAgICAgaWYgKGFjdGl2ZVRhYiA9PT0gJ2Rhc2hib2FyZCcpIERhc2hib2FyZC5yZW5kZXIoKTsKICAgICAgaWYgKGFjdGl2ZVRhYiA9PT0gJ3JlY29yZHMnKSBSZWNvcmRzLnJlbmRlcigpOwogICAgfSk7CiAgfQoKICBmdW5jdGlvbiB3aXJlVGFicygpIHsKICAgIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJy50YWItYnRuJykuZm9yRWFjaCgoYnRuKSA9PiB7CiAgICAgIGJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHN3aXRjaFRhYihidG4uZGF0YXNldC50YWIpKTsKICAgIH0pOwogIH0KCiAgZnVuY3Rpb24gd2lyZVRoZW1lKCkgewogICAgY29uc3QgdG9nZ2xlID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3RoZW1lVG9nZ2xlJyk7CiAgICBjb25zdCBpY29uU2xvdCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd0aGVtZVRvZ2dsZUljb24nKTsKICAgIGNvbnN0IHNldEljb24gPSAobmFtZSkgPT4geyBpY29uU2xvdC5pbm5lckhUTUwgPSBgPGkgZGF0YS1sdWNpZGU9IiR7bmFtZX0iIHN0eWxlPSJ3aWR0aDoxNnB4O2hlaWdodDoxNnB4OyI+PC9pPmA7IH07CiAgICBjb25zdCBzdG9yZWQgPSBsb2NhbFN0b3JhZ2UuZ2V0SXRlbSgnbHJzLXRoZW1lJyk7CiAgICBpZiAoc3RvcmVkKSB7CiAgICAgIGRvY3VtZW50LmRvY3VtZW50RWxlbWVudC5zZXRBdHRyaWJ1dGUoJ2RhdGEtdGhlbWUnLCBzdG9yZWQpOwogICAgICBzZXRJY29uKHN0b3JlZCA9PT0gJ2RhcmsnID8gJ3N1bicgOiAnbW9vbicpOwogICAgfQogICAgdG9nZ2xlLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gewogICAgICBjb25zdCBwcmVmZXJzRGFyayA9IHdpbmRvdy5tYXRjaE1lZGlhKCcocHJlZmVycy1jb2xvci1zY2hlbWU6IGRhcmspJykubWF0Y2hlczsKICAgICAgY29uc3QgY3VycmVudCA9IGRvY3VtZW50LmRvY3VtZW50RWxlbWVudC5nZXRBdHRyaWJ1dGUoJ2RhdGEtdGhlbWUnKSB8fCAocHJlZmVyc0RhcmsgPyAnZGFyaycgOiAnbGlnaHQnKTsKICAgICAgY29uc3QgbmV4dCA9IGN1cnJlbnQgPT09ICdkYXJrJyA/ICdsaWdodCcgOiAnZGFyayc7CiAgICAgIGRvY3VtZW50LmRvY3VtZW50RWxlbWVudC5zZXRBdHRyaWJ1dGUoJ2RhdGEtdGhlbWUnLCBuZXh0KTsKICAgICAgbG9jYWxTdG9yYWdlLnNldEl0ZW0oJ2xycy10aGVtZScsIG5leHQpOwogICAgICBzZXRJY29uKG5leHQgPT09ICdkYXJrJyA/ICdzdW4nIDogJ21vb24nKTsKICAgICAgQ2hhcnRzLmRlc3Ryb3lBbGwoKTsKICAgICAgcmVuZGVyQWN0aXZlVmlldygpOwogICAgfSk7CiAgfQoKICB3aW5kb3cuYWRkRXZlbnRMaXN0ZW5lcignbHJzOmRhdGEtdXBkYXRlZCcsIGFzeW5jICgpID0+IHsKICAgIGF3YWl0IGxvYWRGaWx0ZXJPcHRpb25zKCk7CiAgICByZW5kZXJBY3RpdmVWaWV3KCk7CiAgfSk7CgogIC8vIC0tLS0tLS0tLS0gQXV0aCBzY3JlZW4gLS0tLS0tLS0tLQogIGxldCBhcHBJbml0aWFsaXplZCA9IGZhbHNlOwoKICBmdW5jdGlvbiBzaG93QXV0aFNjcmVlbigpIHsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhdXRoU2NyZWVuJykuc3R5bGUuZGlzcGxheSA9ICdmbGV4JzsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhcHBTaGVsbCcpLnN0eWxlLmRpc3BsYXkgPSAnbm9uZSc7CiAgICBjb25zdCBjb2RlSW5wdXQgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXV0aENvZGUnKTsKICAgIGNvZGVJbnB1dC52YWx1ZSA9ICcnOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2F1dGhFcnJvcicpLnRleHRDb250ZW50ID0gJyc7CiAgICBjb2RlSW5wdXQuZm9jdXMoKTsKICB9CgogIGFzeW5jIGZ1bmN0aW9uIHNob3dBcHAoKSB7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXV0aFNjcmVlbicpLnN0eWxlLmRpc3BsYXkgPSAnbm9uZSc7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXBwU2hlbGwnKS5zdHlsZS5kaXNwbGF5ID0gJyc7CiAgICBpZiAoIWFwcEluaXRpYWxpemVkKSB7CiAgICAgIGFwcEluaXRpYWxpemVkID0gdHJ1ZTsKICAgICAgd2lyZVRhYnMoKTsKICAgICAgd2lyZVRoZW1lKCk7CiAgICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdsb2dvdXRCdG4nKS5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIGFzeW5jICgpID0+IHsKICAgICAgICBhd2FpdCBBcGkuYXV0aExvZ291dCgpOwogICAgICAgIGFwcEluaXRpYWxpemVkID0gZmFsc2U7CiAgICAgICAgc2hvd0F1dGhTY3JlZW4oKTsKICAgICAgfSk7CiAgICAgIGF3YWl0IGxvYWRGaWx0ZXJPcHRpb25zKCk7CiAgICAgIHdpcmVGaWx0ZXJCYXIoKTsKICAgICAgc3dpdGNoVGFiKCdkYXNoYm9hcmQnKTsKICAgIH0gZWxzZSB7CiAgICAgIGF3YWl0IGxvYWRGaWx0ZXJPcHRpb25zKCk7CiAgICAgIHJlbmRlckFjdGl2ZVZpZXcoKTsKICAgIH0KICB9CgogIGFzeW5jIGZ1bmN0aW9uIHN1Ym1pdEF1dGgoKSB7CiAgICBjb25zdCBlcnJvckVsID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2F1dGhFcnJvcicpOwogICAgY29uc3QgYnRuID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2F1dGhTdWJtaXRCdG4nKTsKICAgIGNvbnN0IGNvZGVJbnB1dCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhdXRoQ29kZScpOwogICAgZXJyb3JFbC50ZXh0Q29udGVudCA9ICcnOwogICAgYnRuLmRpc2FibGVkID0gdHJ1ZTsKICAgIGJ0bi50ZXh0Q29udGVudCA9ICdDaGVja2luZ+KApic7CiAgICB0cnkgewogICAgICBhd2FpdCBBcGkuYXV0aExvZ2luKGNvZGVJbnB1dC52YWx1ZSk7CiAgICAgIGF3YWl0IHNob3dBcHAoKTsKICAgIH0gY2F0Y2ggKGVycikgewogICAgICBlcnJvckVsLnRleHRDb250ZW50ID0gZXJyLm1lc3NhZ2U7CiAgICB9IGZpbmFsbHkgewogICAgICBidG4uZGlzYWJsZWQgPSBmYWxzZTsKICAgICAgYnRuLmlubmVySFRNTCA9ICc8aSBkYXRhLWx1Y2lkZT0iYXJyb3ctcmlnaHQiIHN0eWxlPSJ3aWR0aDoxNHB4O2hlaWdodDoxNHB4OyI+PC9pPiBFbnRlcic7CiAgICB9CiAgfQoKICBmdW5jdGlvbiB3aXJlQXV0aEZvcm0oKSB7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXV0aFN1Ym1pdEJ0bicpLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgc3VibWl0QXV0aCk7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYXV0aENvZGUnKS5hZGRFdmVudExpc3RlbmVyKCdrZXlkb3duJywgKGUpID0+IHsgaWYgKGUua2V5ID09PSAnRW50ZXInKSBzdWJtaXRBdXRoKCk7IH0pOwogIH0KCiAgd2luZG93LmFkZEV2ZW50TGlzdGVuZXIoJ2xyczpzaWduZWQtb3V0JywgKCkgPT4gewogICAgYXBwSW5pdGlhbGl6ZWQgPSBmYWxzZTsKICAgIHNob3dBdXRoU2NyZWVuKCk7CiAgfSk7CgogIGFzeW5jIGZ1bmN0aW9uIGluaXQoKSB7CiAgICBhcHBseUJyYW5kaW5nKCk7CiAgICB3aXJlQXV0aEZvcm0oKTsKICAgIGNvbnN0IHsgYXV0aGVudGljYXRlZCB9ID0gYXdhaXQgQXBpLmF1dGhNZSgpOwogICAgaWYgKGF1dGhlbnRpY2F0ZWQpIGF3YWl0IHNob3dBcHAoKTsKICAgIGVsc2Ugc2hvd0F1dGhTY3JlZW4oKTsKICB9CgogIC8vIEljb25zIGFyZSBwbGFjZWQgYXMgPGkgZGF0YS1sdWNpZGU9Ii4uLiI+IHBsYWNlaG9sZGVycyB0aHJvdWdob3V0IHRoZSBkeW5hbWljYWxseQogIC8vIHJlbmRlcmVkIFVJOyBMdWNpZGUgcmVwbGFjZXMgZWFjaCB3aXRoIGFuIGlubGluZSBTVkcuIFJhdGhlciB0aGFuIHJlbWVtYmVyaW5nIHRvIGNhbGwKICAvLyB0aGlzIGFmdGVyIGV2ZXJ5IHNpbmdsZSByZW5kZXIsIG9uZSBvYnNlcnZlciBjYXRjaGVzIGV2ZXJ5IERPTSBjaGFuZ2UgdGhhdCBjb3VsZCBoYXZlCiAgLy8gaW50cm9kdWNlZCBhIG5ldyBwbGFjZWhvbGRlci4KICBpZiAod2luZG93Lmx1Y2lkZSkgewogICAgd2luZG93Lmx1Y2lkZS5jcmVhdGVJY29ucygpOwogICAgLy8gY3JlYXRlSWNvbnMoKSByZXBsYWNlcyA8aSBkYXRhLWx1Y2lkZT4gcGxhY2Vob2xkZXJzIHdpdGggPHN2Zz4g4oCUIGl0c2VsZiBhIERPTQogICAgLy8gbXV0YXRpb24uIFdpdGhvdXQgZGlzY29ubmVjdGluZyBmaXJzdCwgdGhhdCB3cml0ZSByZS10cmlnZ2VycyB0aGlzIHNhbWUgb2JzZXJ2ZXIKICAgIC8vIGZvcmV2ZXIgKGFuIGluZmluaXRlIG11dGF0ZS9vYnNlcnZlIGxvb3AgdGhhdCBwZWdzIHRoZSBDUFUgYW5kIGNyYXNoZXMgdGhlIHRhYikuCiAgICAvLyBEaXNjb25uZWN0aW5nIGJlZm9yZSBlYWNoIHBhc3MgYW5kIHJlY29ubmVjdGluZyBhZnRlciwgcGx1cyBiYXRjaGluZyBidXJzdHMgb2YKICAgIC8vIG11dGF0aW9ucyBpbnRvIGEgc2luZ2xlIG1pY3JvdGFzaywgYnJlYWtzIHRoZSBjeWNsZS4KICAgIGxldCBpY29uc1NjaGVkdWxlZCA9IGZhbHNlOwogICAgY29uc3QgaWNvbk9ic2VydmVyID0gbmV3IE11dGF0aW9uT2JzZXJ2ZXIoKCkgPT4gewogICAgICBpZiAoaWNvbnNTY2hlZHVsZWQpIHJldHVybjsKICAgICAgaWNvbnNTY2hlZHVsZWQgPSB0cnVlOwogICAgICBxdWV1ZU1pY3JvdGFzaygoKSA9PiB7CiAgICAgICAgaWNvbnNTY2hlZHVsZWQgPSBmYWxzZTsKICAgICAgICBpY29uT2JzZXJ2ZXIuZGlzY29ubmVjdCgpOwogICAgICAgIHdpbmRvdy5sdWNpZGUuY3JlYXRlSWNvbnMoKTsKICAgICAgICBpY29uT2JzZXJ2ZXIub2JzZXJ2ZShkb2N1bWVudC5ib2R5LCB7IGNoaWxkTGlzdDogdHJ1ZSwgc3VidHJlZTogdHJ1ZSB9KTsKICAgICAgfSk7CiAgICB9KTsKICAgIGljb25PYnNlcnZlci5vYnNlcnZlKGRvY3VtZW50LmJvZHksIHsgY2hpbGRMaXN0OiB0cnVlLCBzdWJ0cmVlOiB0cnVlIH0pOwogIH0KCiAgaW5pdCgpOwp9KSgpOwo8L3NjcmlwdD4KPC9ib2R5Pgo8L2h0bWw+Cg==';
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
