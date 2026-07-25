/**
 * End-to-end check in a real browser.
 *
 * The unit tests and the oracle prove the engine agrees with LibreOffice on
 * values. This proves the other half: that the values reach the screen. It loads
 * each sample, reads the banner and chips, extracts the rendered text of named
 * cells, and fails on any console error.
 *
 *   node tools/verify/drive.mjs            # against a running dev server
 */

import { chromium } from 'playwright-core';

const URL = process.env.URL ?? 'http://localhost:5176/';

const SAMPLES = [
  'Financial model (recalculated)',
  'Same model, no cached values',
  'Hardcoded total (audit demo)',
  'Formula tour',
];

const errors = [];
let failures = 0;

function check(label, ok, detail = '') {
  console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${label}${detail ? `  ${detail}` : ''}`);
  if (!ok) failures++;
}

const browser = await chromium.launch({
  executablePath: process.env.CHROME ?? '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
  headless: true,
});
const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
page.on('console', (m) => {
  if (m.type() === 'error') errors.push(m.text());
});
page.on('pageerror', (e) => errors.push(String(e)));

await page.goto(URL, { waitUntil: 'networkidle' });

/**
 * Wait for the *named* file to be on screen. Waiting for "computed" to appear
 * anywhere is satisfied by the previous file's chips, which made every reading
 * below a race against the loader.
 */
async function show(sample) {
  await page.getByRole('button', { name: sample, exact: true }).click();
  await page.waitForFunction(
    (label) => document.querySelector('.chip.file')?.textContent?.trim() === label,
    sample,
    { timeout: 20000 },
  );
  await page.waitForSelector('.xl-table td');
}

for (const sample of SAMPLES) {
  console.log(`\n── ${sample}`);
  await show(sample);

  const banner = (await page.locator('.banner').first().innerText()).replace(/\s+/g, ' ');
  const chips = (await page.locator('.chips').innerText()).replace(/\s+/g, ' ');
  console.log(`  banner: ${banner}`);
  console.log(`  chips:  ${chips}`);

  // The trap this project exists to fix: a formula cell that renders nothing
  // because the file carried no value for it. A cell that legitimately computes
  // to "" still renders blank and is not a defect, so the check keys on the
  // uncomputed marker in the tooltip rather than on emptiness.
  const stranded = await page.evaluate(() =>
    [...document.querySelectorAll('.xl-table td')].filter(
      (td) => td.title?.includes('not evaluated') || td.title?.includes('needs a recalc'),
    ).length,
  );
  check('no formula cell is left uncomputed', stranded === 0, `stranded: ${stranded}`);

  const banner2 = await page.locator('.banner').first().innerText();
  check('banner reports a computed count', /computed|could not be computed/i.test(banner2));

  const unsupported = await page.locator('.xl-unsupported').count();
  console.log(`  marked ⚠ (refused, never guessed): ${unsupported}`);
}

// The two model files must agree cell for cell: one carries LibreOffice's cached
// values, the other carries none and is computed here. Same numbers or the whole
// premise fails.
console.log('\n── recalculated vs computed: same numbers?');
const grab = async (sample) => {
  await show(sample);
  const sheets = await page.locator('.tab').count();
  const out = [];
  for (let i = 0; i < sheets; i++) {
    await page.locator('.tab').nth(i).click();
    out.push(
      await page.evaluate(() =>
        [...document.querySelectorAll('.xl-table td')].map((td) => td.textContent.trim()),
      ),
    );
  }
  return out;
};

const cached = await grab('Financial model (recalculated)');
const computed = await grab('Same model, no cached values');

let compared = 0;
let differing = [];
for (let s = 0; s < Math.min(cached.length, computed.length); s++) {
  const a = cached[s];
  const b = computed[s];
  for (let i = 0; i < Math.min(a.length, b.length); i++) {
    compared++;
    if (a[i] !== b[i]) differing.push(`sheet ${s} cell ${i}: "${a[i]}" vs "${b[i]}"`);
  }
}
check(
  `all ${compared} rendered cells match between the cached and the computed file`,
  differing.length === 0,
  differing.length ? `\n      ${differing.slice(0, 10).join('\n      ')}` : '',
);

// The audit case: the hardcoded Q4 total must be reported, not smoothed over.
console.log('\n── audit diff');
await show('Hardcoded total (audit demo)');
await page.getByText('diff vs file').click();
await page.waitForSelector('.gaps.bad li', { timeout: 5000 }).catch(() => {});
const findings = await page.locator('.gaps.bad li').allInnerTexts();
check('the hardcoded total is flagged', findings.length > 0, findings.join(' | '));

console.log(`\nconsole errors: ${errors.length}`);
for (const e of errors.slice(0, 10)) console.log(`  ${e}`);
if (errors.length) failures++;

await page.screenshot({ path: process.env.SHOT ?? '/tmp/xlsx-preview.png', fullPage: false });
await browser.close();

console.log(failures === 0 ? '\nALL CHECKS PASSED' : `\n${failures} CHECK(S) FAILED`);
process.exit(failures === 0 ? 0 : 1);
