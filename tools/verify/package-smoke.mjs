/**
 * Does the thing we would publish actually work?
 *
 * The test suite runs against `src` through Vite aliases, which is the right
 * trade for iteration speed and proves nothing at all about the artefact a
 * consumer installs. Everything between those two — the build, the `exports`
 * map, the file list, whether a relative import kept its `.js`, whether the
 * stylesheet shipped — is invisible to it.
 *
 * So this script deliberately does what the suite cannot: it imports the
 * packages *by their published specifiers* in plain Node ESM, with no bundler,
 * no alias and no TypeScript, and puts a real workbook through them. Node
 * resolves `@xlscalc/xlsx-preview` through the workspace symlink into the
 * package's own `exports`, which is the same code path `npm install` produces.
 *
 * Run it with `npm run smoke`, or as part of `npm run verify`.
 */

import { readFileSync, existsSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '../..');
const require = createRequire(import.meta.url);

let failures = 0;
const check = (label, ok, detail) => {
  if (ok) {
    console.log(`  ok    ${label}`);
  } else {
    failures++;
    console.log(`  FAIL  ${label}${detail ? ` — ${detail}` : ''}`);
  }
};

const read = (name) => {
  const b = readFileSync(join(ROOT, 'apps/demo/public', name));
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength);
};

console.log('\nresolving the packages the way npm install would');

const engine = await import('@xlscalc/formula-engine');
const preview = await import('@xlscalc/xlsx-preview');
const view = await import('@xlscalc/xlsx-preview/view');
const worker = await import('@xlscalc/xlsx-preview/worker');

check('@xlscalc/formula-engine imports', typeof engine.Workbook === 'function');
check('@xlscalc/xlsx-preview imports', typeof preview.loadXlsx === 'function');
check('@xlscalc/xlsx-preview/view imports', typeof view.ExcelView === 'function');
check('@xlscalc/xlsx-preview/worker imports', typeof worker.createPreviewWorker === 'function');

// The engine is the piece anyone can use server-side, and its whole claim is
// that it costs nothing to install. If a dependency ever appears, this is where
// it is noticed rather than in someone's bundle report.
const enginePkg = require('@xlscalc/formula-engine/package.json');
check('the engine still has zero dependencies', enginePkg.dependencies === undefined,
  JSON.stringify(enginePkg.dependencies));

console.log('\nthe files a consumer receives');

const previewPkg = require('@xlscalc/xlsx-preview/package.json');
const previewDir = dirname(require.resolve('@xlscalc/xlsx-preview/package.json'));

for (const [subpath, target] of Object.entries(previewPkg.exports)) {
  const file = typeof target === 'string' ? target : target.default;
  if (!file) continue;
  check(`exports "${subpath}" → ${file}`, existsSync(join(previewDir, file)));
  if (typeof target === 'object' && target.types) {
    check(`  its types exist`, existsSync(join(previewDir, target.types)));
  }
}

check('LICENSE ships with the package', existsSync(join(previewDir, 'LICENSE')));
check('the engine ships its LICENSE too',
  existsSync(join(dirname(require.resolve('@xlscalc/formula-engine/package.json')), 'LICENSE')));

// The stylesheet is not decoration: `⚠` is how a refused cell announces itself,
// and an unstyled one is a value the user does not notice is missing. So the
// class names the renderers emit must all exist in the stylesheet that ships
// beside them — otherwise the two halves of that guarantee drift apart.
const cssPath = join(previewDir, 'dist/view/style.css');
const css = existsSync(cssPath) ? readFileSync(cssPath, 'utf8') : '';
check('the view stylesheet ships', css.length > 0);

const emitted = new Set();
for (const file of ['dist/view/ExcelView.js', 'dist/html.js']) {
  const src = readFileSync(join(previewDir, file), 'utf8');
  for (const m of src.matchAll(/["'`](xl-[a-z-]+)["'`\s]/g)) emitted.add(m[1]);
}
check(`the renderers emit ${emitted.size} xl- classes`, emitted.size >= 4, [...emitted].join(' '));
for (const cls of [...emitted].sort()) {
  check(`  .${cls} is styled`, css.includes(`.${cls}`));
}

console.log('\nputting a real workbook through the published build');

const doc = await preview.loadXlsx(read('financial-model-nocache.xlsx'));
check('loads a workbook with no cached values', doc.sheets.length > 0);
check('every sheet has a layout', doc.layouts.length === doc.sheets.length);
check('it computed the formulas the file left empty',
  doc.model.report.stats.computed > 0,
  `computed=${doc.model.report.stats.computed}`);
check('nothing was rendered as a guessed number',
  doc.model.report.stats.formulas ===
    doc.model.report.stats.computed +
      doc.model.report.stats.cached +
      doc.model.report.stats.unsupported +
      doc.model.report.stats.circular +
      doc.model.report.stats.volatile,
  JSON.stringify(doc.model.report.stats));

const html = preview.renderToHtml(doc, 0);
check('renders to HTML with no React involved', html.startsWith('<table class="xl-table">'));

const restored = preview.fromSnapshot(structuredClone(preview.toSnapshot(doc)));
check('survives a structured clone, so the Worker path is real',
  preview.renderToHtml(restored, 0) === html);

const found = preview.inspectXlsx(read('financial-model-nocache.xlsx'));
check('inspects without evaluating', found.formulas === doc.model.report.stats.formulas,
  `${found.formulas} vs ${doc.model.report.stats.formulas}`);
check('the dry run never over-claims refusals',
  found.unsupportedCells <=
    doc.model.report.stats.unsupported + doc.model.report.stats.circular);

console.log(
  failures === 0
    ? '\nthe published packages resolve, import and work.\n'
    : `\n${failures} check(s) failed.\n`,
);
process.exit(failures === 0 ? 0 : 1);
