/**
 * Recalculating a workbook on a server, with no browser and no UI framework.
 *
 *   node examples/node-headless.mjs path/to/workbook.xlsx
 *
 * The browser is where this library is aimed, but nothing about the pipeline
 * requires one: bytes in, values out. The three things a server usually wants
 * are all here — decide whether a file is worth rendering, compute it, and get
 * HTML back — and none of them touch the DOM.
 *
 * Everything below imports the packages by their published names, so this file
 * is also a check that the published surface is enough to do real work.
 */

import { readFileSync } from 'node:fs';
import { argv } from 'node:process';

import { inspectXlsx, loadXlsx, renderToHtml, plainText } from '@xlscalc/xlsx-preview';

const path = argv[2];
if (!path) {
  console.error('usage: node examples/node-headless.mjs <file.xlsx>');
  process.exit(2);
}

const bytes = readFileSync(path);
const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);

// ── 1. Ask before paying ─────────────────────────────────────────────────────
// A dry run: reads the OOXML and parses every formula, but evaluates nothing and
// never loads the styling dependency. On a 138,000-formula workbook this is
// under a second, against seven and a half for the full load.

const found = inspectXlsx(buf);
console.log(`${path}`);
console.log(`  saved by       ${found.writer ?? '(the file does not say)'}`);
console.log(`  formulas       ${found.formulas.toLocaleString()} (${found.uncached.toLocaleString()} with no cached value)`);
console.log(`  vocabulary     ${found.functions.length} distinct functions`);

if (!found.fullyCovered) {
  const worst = found.unsupported.slice(0, 5).map((f) => `${f.name}×${f.count}`).join(', ');
  console.log(`  NOT COVERED    ${worst}`);
  console.log(`                 at least ${found.unsupportedCells} cells use these, and they poison`);
  console.log(`                 whatever reads them — this may be a server-side job instead.`);
}
if (found.iterative) {
  console.log(`  iterative      this workbook asks Excel to resolve cycles by iterating`);
}

// ── 2. Compute ───────────────────────────────────────────────────────────────
// `now` is pinned so NOW()/TODAY() do not make two runs of the same file
// disagree — worth doing anywhere the output is cached, diffed or tested.

const doc = await loadXlsx(buf, { now: 45000 });
const s = doc.model.report.stats;
console.log(`\n  computed       ${s.computed} · from the file ${s.cached} · refused ${s.unsupported} · cycles ${s.circular}`);
console.log(`  timing         parse ${doc.parseMs}ms · layout ${doc.layoutMs}ms · evaluate ${doc.evalMs}ms`);

if (doc.stylesError) {
  // The values are unaffected — they come from the other parse — but say so.
  console.log(`  styling        could not be read (${doc.stylesError}); rendering in default fonts`);
}

// ── 3. Use the values ────────────────────────────────────────────────────────
// Every cell carries where its value came from. `unsupported` never carries a
// number, so there is no path from "we could not compute this" to a figure that
// looks computed.

for (const gap of doc.model.report.gaps.slice(0, 5)) {
  console.log(`  gap            ${gap.subject} ×${gap.count} — ${gap.reason}`);
}

for (const m of doc.model.report.mismatches.slice(0, 5)) {
  console.log(`  disagrees      ${m.address}: the file says ${plainText(m.cached)}, the formula gives ${plainText(m.computed)}`);
}

// ── 4. Render, without React ─────────────────────────────────────────────────
// The stylesheet is read from the package rather than copied, so there is one
// source of truth for what a refused cell looks like.

const css = readFileSync(
  new URL(import.meta.resolve('@xlscalc/xlsx-preview/view/style.css')),
  'utf8',
);
const html = renderToHtml(doc, 0, { document: true, css, showProvenance: true });
console.log(`\n  html           ${html.length.toLocaleString()} bytes for sheet "${doc.sheets[0]?.name}"`);
