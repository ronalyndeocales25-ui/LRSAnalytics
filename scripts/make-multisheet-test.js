/* Builds a small 2-sheet .xlsx to verify multi-sheet parsing + raw-only preservation. */
const ExcelJS = require('exceljs');

async function main() {
  const wb = new ExcelJS.Workbook();

  const s1 = wb.addWorksheet('JanuaryData');
  s1.addRow(['', 'IDENTIFIERS', '', '', '', '', 'FACEBOOK', '', '', '', 'INSTAGRAM', '', '', '']);
  s1.addRow(['Ads/Organic', 'Post', 'FORMAT', 'Publish Date', 'Posting Time', 'Platforms', 'Views', 'Reach', 'Engagement', 'Posting link', 'Views', 'Reach', 'Interactions', 'Posting link']);
  s1.addRow(['Choose', 'Multi-sheet test post', 'REEL', '2026-08-03', '10:00 AM', 'FACEBOOK', 1500, 400, 60, 'https://fb.com/x', '', '', '', '']);
  // A row with an unparseable date -- should still be preserved raw, but not become a post.
  s1.addRow(['Organic', 'Bad date row', 'REEL', 'not-a-date', '9:00 AM', 'FACEBOOK', 200, 50, 10, 'https://fb.com/y', '', '', '', '']);

  const s2 = wb.addWorksheet('RandomNotes');
  s2.addRow(['Note', 'Author', 'Date']);
  s2.addRow(['Remember to renew domain', 'Mico', '2026-08-01']);
  s2.addRow(['Client call at 3pm', 'Patrick', '2026-08-02']);

  await wb.xlsx.writeFile('sample-data/multisheet-test.xlsx');
  console.log('wrote sample-data/multisheet-test.xlsx');
}
main();
