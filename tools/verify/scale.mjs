/**
 * How big a model can this preview take before it should move off the main thread?
 *
 * The design budgets < 30 ms for 1k formulas and < 300 ms for 10k. This measures
 * the real numbers so the worker decision is made on data rather than on a
 * feeling that "it might be slow".
 *
 *   node tools/verify/scale.mjs
 */
import { chromium } from 'playwright-core';

const browser = await chromium.launch({
  executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
  headless: true,
});
const page = await browser.newPage();
await page.goto(process.env.URL ?? 'http://localhost:5176/', { waitUntil: 'networkidle' });

// Driven through the real UI rather than by importing the module, so the
// numbers include everything a user actually waits for.
const rows = [];
for (const size of ['scale-1k.xlsx', 'scale-10k.xlsx', 'scale-50k.xlsx']) {
  const ok = await page.evaluate(async (file) => {
    const res = await fetch('/' + file);
    return res.ok;
  }, size);
  if (!ok) {
    console.log(`  ${size}: not generated, skipping`);
    continue;
  }
  await page.setInputFiles('.filebtn input', `apps/demo/public/${size}`);
  await page.waitForFunction((f) => document.querySelector('.chip.file')?.textContent?.trim() === f, size, {
    timeout: 120000,
  });
  const chips = await page.locator('.chips').innerText();
  const parse = /parse (\d+) ms/.exec(chips)?.[1];
  const evalMs = /eval (\d+) ms/.exec(chips)?.[1];
  const computed = /computed\s+(\d+)/.exec(chips)?.[1];
  rows.push({ file: size, computed, parse, evalMs });
  console.log(`  ${size.padEnd(16)} ${String(computed).padStart(6)} formulas   parse ${parse} ms   eval ${evalMs} ms`);
}

await browser.close();
