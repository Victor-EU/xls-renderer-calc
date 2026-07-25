/** Screenshot one sample, for eyeballing the render. `node tools/verify/shot.mjs "Formula tour" out.png` */
import { chromium } from 'playwright-core';

const sample = process.argv[2] ?? 'Formula tour';
const out = process.argv[3] ?? '/tmp/xlsx-shot.png';
const tab = process.argv[4];

const browser = await chromium.launch({
  executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
  headless: true,
});
const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
await page.goto(process.env.URL ?? 'http://localhost:5176/', { waitUntil: 'networkidle' });
await page.getByRole('button', { name: sample, exact: true }).click();
await page.waitForFunction(
  (label) => document.querySelector('.chip.file')?.textContent?.trim() === label,
  sample,
);
await page.waitForSelector('.xl-table td');
if (tab) {
  await page.locator('.tab', { hasText: tab }).first().click();
  await page.waitForTimeout(200);
}
await page.getByText('diff vs file').click();
await page.waitForTimeout(250);
await page.screenshot({ path: out });
await browser.close();
console.log(`wrote ${out}`);
