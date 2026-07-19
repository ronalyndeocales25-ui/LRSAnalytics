/* One-off smoke test driving auth + the upload + analytics API end-to-end. Not part of the app; safe to delete. */
const fs = require('fs');
const path = require('path');

const BASE = 'http://localhost:4000';
let sessionCookie = '';

function authHeaders(extra = {}) {
  return { ...extra, ...(sessionCookie ? { Cookie: sessionCookie } : {}) };
}

async function request(url, options = {}) {
  const res = await fetch(url, { ...options, headers: authHeaders(options.headers) });
  const setCookie = res.headers.get('set-cookie');
  if (setCookie) sessionCookie = setCookie.split(';')[0];
  const json = await res.json().catch(() => null);
  if (!res.ok) throw new Error(`${url} failed (${res.status}): ${JSON.stringify(json)}`);
  return json;
}

async function ensureSignedIn() {
  const code = fs.readFileSync(path.join(__dirname, '..', 'data', 'access-code.txt'), 'utf8').trim();
  await request(`${BASE}/api/auth/login`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ code }),
  });
  console.log('Signed in with access code.');
}

async function previewFile(filePath) {
  const form = new FormData();
  const buf = fs.readFileSync(filePath);
  form.append('file', new Blob([buf]), path.basename(filePath));
  return request(`${BASE}/api/uploads/preview`, { method: 'POST', body: form });
}

async function commit(filePath, originalName, defaultDuplicateAction, duplicateActions = {}, notes = null) {
  return request(`${BASE}/api/uploads/commit`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ filePath, originalName, defaultDuplicateAction, duplicateActions, notes }),
  });
}

async function get(urlPath) {
  return request(`${BASE}${urlPath}`);
}

async function main() {
  console.log('--- Step 0: register + sign in ---');
  await ensureSignedIn();

  console.log('\n--- Step 1: preview + commit week1 sample ---');
  const p1 = await previewFile('sample-data/LRS-sample-week1.csv');
  console.log('preview1:', JSON.stringify({ validRows: p1.validRows, errorRows: p1.errorRows, newRecordsCount: p1.newRecordsCount, duplicates: p1.duplicates.length }, null, 2));
  const c1 = await commit(p1.filePath, p1.originalName, 'skip');
  console.log('commit1:', c1);

  console.log('\n--- Step 2: preview + commit edge-case sample ---');
  const p2 = await previewFile('sample-data/LRS-sample-edgecases.csv');
  console.log('preview2:', JSON.stringify({ validRows: p2.validRows, errorRows: p2.errorRows, newRecordsCount: p2.newRecordsCount }, null, 2));
  const c2 = await commit(p2.filePath, p2.originalName, 'skip');
  console.log('commit2:', c2);

  console.log('\n--- Step 3: analytics sanity checks ---');
  console.log('filter-options:', await get('/api/analytics/filter-options'));
  console.log('kpis Jan:', JSON.stringify(await get('/api/analytics/kpis?dateFrom=2026-01-01&dateTo=2026-01-31'), null, 2));

  console.log('\n--- Step 4: resubmit week1 (duplicates expected) — preview should list them ---');
  const p3 = await previewFile('sample-data/LRS-sample-week1-resubmit.csv');
  console.log('preview3 duplicates:', JSON.stringify(p3.duplicates, null, 2));
  console.log('preview3 newRecordsCount:', p3.newRecordsCount);

  console.log('\n--- Step 5: commit resubmit with defaultDuplicateAction="update" — expect 1 new + 1 updated ---');
  const c3 = await commit(p3.filePath, p3.originalName, 'update');
  console.log('commit3 (update):', c3);

  console.log('\n--- Step 6: verify facebook views updated for Jan 1 week ---');
  const breakdown = await get('/api/analytics/platform-breakdown?dateFrom=2026-01-01&dateTo=2026-01-02');
  console.log('platform-breakdown Jan1-2:', breakdown);

  console.log('\n--- Step 7: upload history ---');
  const history = await get('/api/uploads/history');
  console.log(history.map((u) => ({ id: u.id, filename: u.filename, imported: u.imported_rows, updated: u.updated_rows, skipped: u.skipped_rows })));

  console.log('\n--- Step 8: monthly / quarterly / ytd reports ---');
  console.log('monthly Jan 2026 totals:', (await get('/api/analytics/monthly?year=2026&month=1')).totals);
  console.log('quarterly Q1 2026 totals:', (await get('/api/analytics/quarterly?year=2026&quarter=1')).totals);
  console.log('ytd 2026 totals:', (await get('/api/analytics/ytd?year=2026')).totals);

  console.log('\nAll smoke-test steps completed.');
}

main().catch((err) => {
  console.error('SMOKE TEST FAILED:', err);
  process.exit(1);
});
