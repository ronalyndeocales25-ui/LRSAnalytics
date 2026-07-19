# LRS Analytics Dashboard

A self-hosted social media analytics reporting system for Ligon-Razon Solutions.
Enter the access code, upload a weekly CSV/Excel export, and it validates the
data, resolves duplicates record-by-record, and refreshes every chart and
KPI — all in a single-file Node.js + SQLite app (`app.js`) you run yourself,
with no Google Cloud setup and no external services.

## Requirements

- Node.js **22.5 or newer** (this app uses Node's built-in `node:sqlite`
  module — no native build tools / Python required, unlike `better-sqlite3`).

## Getting started

```bash
npm install
npm start
```

Then open **http://localhost:4000**. The port can be changed with the `PORT`
environment variable.

Data is stored in `data/lrs.db` (a single SQLite file — back it up like any
other file). Uploaded files are held in `data/uploads/` only long enough to
be parsed, then deleted.

### Access code

The whole app sits behind a single shared access code — there are no
individual accounts, and nothing is attributed to a specific person (uploads
and edits are anonymous as far as the app is concerned). The first time the
server runs it creates `data/access-code.txt` with a default code of
**`LRS2026`**; open that file in any text editor to change it (then restart
the server for the new code to take effect — the server also prints the
current code to the console on startup). A signed-in session persists for 30
days via an httpOnly cookie.

## How it works

1. **Enter the access code.**
2. **Upload** a `.csv` or `.xlsx` file on the Upload tab (drag-and-drop or
   click to browse).
3. The server **parses and validates** it without writing anything yet, and
   shows you: how many rows are valid, which rows had errors (with the
   reason), duplicate rows within the file, and — for every row that matches
   a record already in the database — a per-record duplicate list.
4. For each matched duplicate, choose **Skip** (leave the existing record
   untouched), **Update** (overwrite it with the new row's values), or
   **Create** (keep both as separate records) — individually per record, or
   set one default action for every duplicate in the file. Rows that don't
   match anything existing are always imported automatically; you're only
   ever asked about genuine repeats. An optional note can be attached to the
   whole upload (e.g. "Week 3 corrections").
5. Click **Import**. Every dashboard, chart, and KPI reflects the new data
   immediately — no separate refresh step.

### Supported file layouts

The parser auto-detects two layouts, so you're not locked into one export
format:

- **LRS agenda tracker** — the wide format with a platform-group header row
  (FACEBOOK / INSTAGRAM / TIKTOK / LINKED IN / THREADS / YouTube) followed by
  a per-column metric header row. This is the format of the reference
  `LRS Agenda Tracking 2026` sheet.
- **Simple table** — one row per platform per post, with a `platform` column
  and metric columns named things like `views`, `reach`, `engagement`,
  `impressions`, `clicks`, `followers gained`, `watch time`, `shares`,
  `comments`, `saves`, `posting link`.

Column headers are matched by **synonym**, not fixed position — see the
`PLATFORMS`/`METRIC_SYNONYMS`/`IDENTIFIER_COLUMN_SYNONYMS` constants near the
top of `app.js`. That means:

- **Adding a platform** later (e.g. Snapchat) is a one-line addition to the
  `PLATFORMS` array — no parser changes.
- **Adding a metric** (e.g. a real `shares`/`comments`/`saves` column, which
  the reference sheet doesn't currently break out) is a one-line addition to
  `METRIC_SYNONYMS` — the database, API, and dashboard already support all
  ten canonical metrics end to end.

### Deduplication & conflict resolution

Each source row gets a fingerprint of **every** imported field — not just
date/caption/platform, but every metric (Views, Reach, Engagement, Clicks,
Followers Gained, etc.) for every platform on the row. A row is only ever
flagged as a duplicate if it's byte-for-byte identical to an already-saved
record; those are offered to you individually — nothing is overwritten or
discarded without an explicit **Skip / Update / Create** decision (or the
file-wide default you set before importing). A row with the same
date/caption/platform as an existing record but even one different metric
value (the normal case for a recurring weekly export where a post's numbers
keep climbing) is **not** a duplicate — it's imported automatically as its
own new record, so no analytics update is ever silently lost. Brand-new
records never require a decision either; they're imported automatically too.

### Every sheet, every row — nothing is silently dropped

- **Multi-sheet workbooks**: an `.xlsx` file with multiple tabs has *every*
  tab parsed, not just the first. Each sheet is independently detected as
  agenda / simple / unrecognized, and the Upload preview shows a per-sheet
  breakdown (layout detected, row/valid/error counts).
- **Raw preservation**: every non-blank row from every sheet — including
  ones that fail to parse (bad dates, unrecognized platforms) or come from a
  sheet whose layout isn't recognized at all — is stored verbatim in a
  `raw_rows` table, linked back to the post it became (when it became one).
  Nothing is modified or discarded; a row that can't be charted is still
  fully recoverable. Expand any entry on the **Upload History** tab and
  click "View every raw source row from this upload" to see it, grouped by
  sheet with the original headers.
- Rows that are 100% blank (every cell empty) are the only thing not stored
  individually, since there is nothing in them to preserve.

## Dashboards

- **Dashboard** — a metric-focused BI view. A **Metric** dropdown sits next
  to the shared **Platform** filter; its options are read live from
  whatever the selected platform's data actually contains (via
  `/api/analytics/metric-options`) — never a hardcoded per-platform list, so
  a metric only appears once a real column for it has been imported. It
  also offers one derived metric, **Engagement Rate** (engagement ÷ reach,
  or ÷ impressions where reach isn't tracked), whenever both halves of that
  ratio exist for the selection. Everything below the selectors — six KPI
  cards (Highest/Average/Total/Number of Posts/Best Performing Post/Lowest
  Performing Post), the weekly trend chart, a Platform Comparison chart
  (Campaign Performance instead, once a single platform is selected — since
  comparing one platform to itself isn't useful), a Content Type
  Performance chart, and the Top Performing Posts ranking (with a **View
  Details** button opening the same full-record popup as Data Records) —
  recomputes for the selected platform + metric + date range + campaign +
  content type. Every bar, line, and pie chart shows its actual values
  directly on the chart (via `chartjs-plugin-datalabels`), not just on
  hover, formatted the way a presentation deck would (`850`, `1,250`,
  `12.5K`, `156K`, `1.25M`) — point/bar labels are skipped only when a chart
  has enough items that labeling every one would overlap, in which case the
  hover tooltip still has the exact value.
- **Data Records** — a CRM-style, platform-grouped browser instead of a
  giant spreadsheet-in-a-table. A platform-pill filter (**All Platforms**,
  Facebook, Instagram, TikTok, LinkedIn, Threads, YouTube) plus a search box
  and the same date/campaign/content-type filters as the Dashboard sit above
  the table:
  - **All Platforms** shows one row per record with only the common fields
    (Date, Platforms, Caption, Campaign, Content Type, Status, Last Updated).
  - Selecting a **specific platform** narrows the table to records that
    have data on that platform and swaps in its own curated metric columns
    (e.g. Facebook shows Views/Reach/Engagement/Link; YouTube shows
    Views/Watch Time/Subscribers/Impressions/Link) — only what that
    platform's block in the source sheet actually captures, so no table is
    ever stretched across every platform's metrics at once.
  - Every row has **View** (a read-only popup with every imported field,
    grouped into sections by platform), **Edit** (the same, but every field
    editable), and **Delete**. Edit/View always load the complete source
    row — not just the columns visible in the table. Saving an edit re-syncs
    the underlying dashboard post + platform metrics in the same
    transaction, so the table, Dashboard, Comparisons, and every report
    reflect the change immediately; a **Status** badge flags whether a
    record has been edited since import. Delete removes just that
    platform's data when used from a platform-specific view (or the whole
    record, across every platform, from All Platforms) — the original
    import is always kept in Upload History's raw-row viewer regardless.
- **Comparisons** — Week vs. Week, a fully custom date-range vs. date-range,
  Monthly (vs. previous month or vs. the same month last year), Quarterly,
  and Year-to-Date — each showing per-metric growth % and a platform-by
  -platform comparison chart.
- **Upload History** — every past import with its status, row counts, any
  note left at import time, and a per-row error/skip log you can expand.

## Project layout

```
app.js                     The entire application — one file, run with `node app.js`
data/                      Runtime data (not source): lrs.db, access-code.txt, uploads/ staging
sample-data/               Small fixture files used by scripts/smoke-test.js
scripts/smoke-test.js      Optional end-to-end API test (upload -> merge -> compare)
scripts/make-multisheet-test.js  Generates sample-data/multisheet-test.xlsx (2-tab fixture
                           used to verify multi-sheet parsing + raw-row preservation)
_pre-single-file-backup/   The previous server/ + public/ split, kept for reference/rollback
                           only — nothing reads from this directory; safe to delete once
                           you're comfortable app.js is working the way you expect.
```

The whole application — SQLite schema, access-code auth, the spreadsheet
parser, import/records/analytics logic, every Express route, **and** the
frontend page itself — lives in `app.js`. Internally it's organized into the
same clearly-commented sections a multi-file layout would have (schema,
platform/metric config, access code, utils, services, routes, ...), just
concatenated into one file instead of split across `server/`.

The frontend (markup, CSS, and JS, all originally written as one
`public/index.html`) is embedded in `app.js` as a base64 string and served
verbatim on `GET /` — base64 specifically so the page's own inline
`<script>`, which uses backtick template literals extensively, never
conflicts with wrapping it in a JS string literal here. If you ever need to
edit the frontend by hand, decode it, edit, then re-encode: see
`_pre-single-file-backup/public/index.html` for the last known-good
plain-text copy, or extract the current one straight from `app.js`:

```bash
node -e "
const fs = require('fs');
const src = fs.readFileSync('app.js', 'utf8');
const b64 = src.match(/const INDEX_HTML_BASE64 = '([^']*)'/)[1];
fs.writeFileSync('index.extracted.html', Buffer.from(b64, 'base64'));
"
```

No build step either way — editing the extracted HTML and re-embedding it
(reverse of the above, replacing the base64 string in `app.js`) is the whole
workflow.

Four external dependencies are loaded from a CDN inside the embedded page:
Chart.js and its official `chartjs-plugin-datalabels` add-on (prints values
directly on the Dashboard's charts), the Inter font (Google Fonts), and
`lucide` (an icon library — icons are placed as `<i data-lucide="...">`
placeholders and a single `MutationObserver` swaps every one of them for its
inline SVG as soon as it appears, so no render path has to remember to call
it). The frontend's own internal logic is still organized into the same
sections it always was (Api, State/Format/Toast, Charts, Dashboard,
Comparison, Records, Upload/History, app bootstrap).

### Look and feel

The UI follows a glassmorphism design language — semi-transparent, blurred
surfaces (`backdrop-filter`) for the topbar, filter bar, cards, KPI tiles,
tables, modals, and inputs, layered over a soft gray gradient page
background, with the existing LRS blue as the one accent color (kept as
`--series-1`, unchanged from before). Every visual value lives in CSS custom
properties under `:root` (and its dark-mode overrides), so retheming means
editing tokens in one place, not hunting through component rules:
`--surface-1`/`--surface-2` (glass backgrounds), `--glass-blur`, `--border`,
`--radius-sm/md/lg` (10/14/18px), and `--shadow-card`/`--shadow-hover`/
`--shadow-modal`. Buttons, inputs, tables, cards, tabs, and modals all
transition on hover/press/open (150-300ms, no bounce); tables have sticky
glass headers, alternating row tint, and row-hover highlight instead of a
spreadsheet look; empty states show an icon, an explanation, and — where
there's an obvious next step — an "Upload data" button, never a blank area;
sections that are still loading show a shimmering skeleton placeholder
instead of a blank gap. All of this was built and verified without a
browser available in this environment (structural/JS/CSS validation only,
via `node --check` and brace-balance checks plus the smoke test) — it's
worth a visual pass once you're at a keyboard.

## Optional: run the smoke test

With the server running in another terminal:

```bash
node scripts/smoke-test.js
```

This signs in with whatever access code is in `data/access-code.txt`, uploads
the fixture files in `sample-data/`, exercises the Skip/Update/Create
duplicate-resolution path, and prints the resulting KPIs — useful after
changing the parser, import, or auth logic. It writes real rows into
`data/lrs.db`; delete that file (and its `-shm`/`-wal` siblings) afterward if
you want to start from a clean database again.

## Known limitations

- The Dashboard's metric list is limited to what the app actually stores per
  post (Views, Reach, Impressions, Engagement, Clicks, Followers Gained,
  Watch Time, Shares, Comments, Saves, plus the derived Engagement Rate).
  Metrics some platforms' native analytics exports offer but this project's
  source sheet doesn't break out — e.g. Facebook Reactions as distinct from
  Engagement, CTR, TikTok Completion Rate, Instagram Accounts
  Reached/Engaged or Profile Visits, Threads Replies/Reposts/Quotes — aren't
  available and won't appear in the Metric dropdown until a future upload
  actually contains that column (see `METRIC_SYNONYMS` in `app.js` to teach
  the parser a new one). There's also no
  cumulative "Followers Growth over time" chart yet — that needs the
  separate weekly-follower-count tracking described as a future phase, not
  the per-post "Followers Gained" metric this dashboard already charts.
- The uploaded file is fully re-parsed on both the preview step and the
  commit step (not cached in the database between the two). For the sheet
  sizes this is built for (weekly exports), this is not noticeable, but a
  very large file would parse twice.
- "Top posts" ranks individual **platform-post** rows (e.g. a Reel's TikTok
  performance and its Instagram performance are ranked separately), since a
  cross-post total doesn't make sense when platforms report different
  metrics.
- Excel parsing uses `exceljs` rather than the more common `xlsx` package,
  specifically because the published `xlsx` package has unpatched
  prototype-pollution/ReDoS advisories that matter when parsing
  user-uploaded files.
