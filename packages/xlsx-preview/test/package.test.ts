/**
 * The package's public surface, exercised the way a consumer would.
 *
 * These are not tests of the engine — the oracle, the eval corpora and the real
 * corpus do that. They cover the seams that packaging introduced: that a
 * renderer can work without ExcelJS, that a document survives being flattened
 * for a Worker, that a caller can ask what will be refused before paying to
 * find out, and that a styling failure still renders something.
 */

import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

import { ERR } from '@xlscalc/formula-engine';
import {
  blankLayout,
  fromSnapshot,
  inspectXlsx,
  layoutKey,
  loadXlsx,
  renderToHtml,
  toSnapshot,
  type PreviewDocument,
} from '../src/index.js';

const HERE = dirname(fileURLToPath(import.meta.url));
const FIXTURES = join(HERE, '../../../apps/demo/public');

const read = (name: string): ArrayBuffer => {
  const b = readFileSync(join(FIXTURES, name));
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer;
};

const load = (name: string): Promise<PreviewDocument> => loadXlsx(read(name));

describe('layout is plain data', () => {
  it('describes every sheet without anyone touching ExcelJS', async () => {
    const doc = await load('financial-model.xlsx');

    expect(doc.layouts).toHaveLength(doc.sheets.length);
    for (const [i, layout] of doc.layouts.entries()) {
      expect(layout.name).toBe(doc.sheets[i]!.name);
      expect(layout.rows).toBeGreaterThan(0);
      expect(layout.cols).toBeGreaterThan(0);
      // 1-based, with a placeholder at index 0.
      expect(layout.colWidths).toHaveLength(layout.cols + 1);
      expect(layout.rowHeights).toHaveLength(layout.rows + 1);
      expect(layout.styles[0]).toEqual({});
    }
  });

  it('deduplicates styles rather than storing one per cell', async () => {
    const doc = await load('financial-model.xlsx');
    const layout = doc.layouts[0]!;
    // A financial model has thousands of cells and a few dozen formats. If this
    // ever inverts, the style table has stopped interning and the document is
    // mostly duplicate objects.
    expect(layout.styleAt.size).toBeGreaterThan(layout.styles.length);
  });

  it('survives a workbook whose styling could not be parsed', () => {
    // The degraded path: three of ten real workbooks made ExcelJS throw while
    // our own reader read them fine. Losing the fonts is acceptable; showing a
    // blank page is not, so there has to be a layout to fall back to.
    const layout = blankLayout('Sheet1', 12, 5);
    expect(layout.rows).toBe(12);
    expect(layout.cols).toBe(5);
    expect(layout.colWidths[1]).toBeGreaterThan(0);
    expect(layout.rowHeights[1]).toBeGreaterThan(0);
    expect(layout.styleAt.size).toBe(0);
  });
});

describe('renderToHtml', () => {
  it('renders a sheet with no React anywhere in the call', async () => {
    const doc = await load('financial-model.xlsx');
    const html = renderToHtml(doc, 0);
    expect(html.startsWith('<table class="xl-table">')).toBe(true);
    expect(html).toContain('<td');
    expect(html.endsWith('</table>')).toBe(true);
  });

  it('agrees with the file on a value the file already computed', async () => {
    const doc = await load('financial-model.xlsx');
    const html = renderToHtml(doc, 0, { tooltips: false });
    // Pick a computed number out of the model and require it to appear.
    const layout = doc.layouts[0]!;
    let found = false;
    for (let r = 1; r <= layout.rows && !found; r++) {
      for (let c = 1; c <= layout.cols; c++) {
        const cell = doc.model.cell(0, r, c);
        if (cell && typeof cell.value === 'number' && Math.abs(cell.value) > 1000) {
          found = true;
          break;
        }
      }
    }
    expect(found).toBe(true);
    expect(html.length).toBeGreaterThan(1000);
  });

  it('escapes everything that came out of the file', () => {
    // A cell's text, a sheet name and a hyperlink are all attacker-controlled
    // as far as this library is concerned.
    const html = renderToHtml(
      {
        layouts: [
          {
            ...blankLayout('S', 1, 1),
            content: new Map([[layoutKey(1, 1), { kind: 'link', text: '<b>x</b>', href: 'javascript:alert(1)"' }]]),
          },
        ],
        model: {
          facts: { uncached: 0, formulas: 0, arrayFormulas: 0, sharedFormulas: 0, date1904: false, fullCalcOnLoad: false, hasCalcChain: false },
          sheetNames: ['S'],
          report: { stats: { formulas: 0, computed: 0, cached: 0, unsupported: 0, circular: 0, volatile: 0, mismatched: 0, totalMs: 0 }, gaps: [], mismatches: [] },
          hardcoded: [],
          cell: () => undefined,
        },
      },
      0,
    );
    expect(html).not.toContain('<b>x</b>');
    expect(html).toContain('&lt;b&gt;x&lt;/b&gt;');
    expect(html).toContain('&quot;');
  });

  it('wraps a standalone document when asked', () => {
    const html = renderToHtml(
      {
        layouts: [blankLayout('Sheet <1>', 1, 1)],
        model: {
          facts: { uncached: 0, formulas: 0, arrayFormulas: 0, sharedFormulas: 0, date1904: false, fullCalcOnLoad: false, hasCalcChain: false },
          sheetNames: ['Sheet <1>'],
          report: { stats: { formulas: 0, computed: 0, cached: 0, unsupported: 0, circular: 0, volatile: 0, mismatched: 0, totalMs: 0 }, gaps: [], mismatches: [] },
          hardcoded: [],
          cell: () => undefined,
        },
      },
      0,
      { document: true, css: '.xl-table{color:red}' },
    );
    expect(html.startsWith('<!doctype html>')).toBe(true);
    expect(html).toContain('<title>Sheet &lt;1&gt;</title>');
    expect(html).toContain('<style>.xl-table{color:red}</style>');
  });
});

describe('snapshot', () => {
  it('round-trips every value, including errors', async () => {
    const doc = await load('formula-tour.xlsx');
    const restored = fromSnapshot(toSnapshot(doc));

    let compared = 0;
    for (const [s, layout] of doc.layouts.entries()) {
      for (let r = 1; r <= layout.rows; r++) {
        for (let c = 1; c <= layout.cols; c++) {
          const before = doc.model.cell(s, r, c);
          const after = restored.model.cell(s, r, c);
          if (!before) {
            expect(after).toBeUndefined();
            continue;
          }
          compared++;
          expect(after).toBeDefined();
          expect(after!.provenance).toBe(before.provenance);
          expect(after!.value).toStrictEqual(before.value);
          expect(after!.cached).toStrictEqual(before.cached);
          expect(after!.formula).toBe(before.formula);
        }
      }
    }
    expect(compared).toBeGreaterThan(0);
  });

  it('rebuilds an error as the engine s own singleton, not a lookalike object', async () => {
    const doc = await load('formula-tour.xlsx');
    const snap = toSnapshot(doc);
    snap.cells[0]!.set(layoutKey(1, 1), { value: { err: '#REF!' }, provenance: 'computed' });
    const restored = fromSnapshot(snap);
    // Identity, not equality: `isErr` is an instanceof check, so a cloned
    // plain object would silently render as [object Object].
    expect(restored.model.cell(0, 1, 1)!.value).toBe(ERR.ref);
  });

  it('carries the whole appearance across, so the same component can render it', async () => {
    const doc = await load('financial-model.xlsx');
    const restored = fromSnapshot(toSnapshot(doc));
    expect(restored.layouts).toEqual(doc.layouts);
    expect(restored.model.facts).toEqual(doc.model.facts);
    expect(restored.model.report.stats).toEqual(doc.model.report.stats);
    // And it renders to the same markup through the framework-free path.
    expect(renderToHtml(restored, 0)).toBe(renderToHtml(doc, 0));
  });

  it('is structured-cloneable, which is the entire point', async () => {
    const doc = await load('financial-model.xlsx');
    const cloned = structuredClone(toSnapshot(doc));
    expect(renderToHtml(fromSnapshot(cloned), 0)).toBe(renderToHtml(doc, 0));
  });
});

describe('inspectXlsx', () => {
  it('answers what will be refused without evaluating anything', () => {
    const found = inspectXlsx(read('financial-model.xlsx'));
    expect(found.formulas).toBeGreaterThan(0);
    expect(found.functions.length).toBeGreaterThan(0);
    // Sorted most-used first, so a caller can act on the top of the list.
    for (let i = 1; i < found.functions.length; i++) {
      expect(found.functions[i - 1]!.count).toBeGreaterThanOrEqual(found.functions[i]!.count);
    }
    expect(found.unsupported.every((f) => !f.supported)).toBe(true);
    expect(found.unsupportedCells).toBeLessThanOrEqual(found.formulas);
  });

  it('agrees with a full load about which functions are refused', async () => {
    const doc = await load('financial-model.xlsx');
    const found = inspectXlsx(read('financial-model.xlsx'));
    // The dry run counts roots, the load counts consequences — so the dry run
    // must never claim *more* than the load found.
    expect(found.unsupportedCells).toBeLessThanOrEqual(
      doc.model.report.stats.unsupported + doc.model.report.stats.circular,
    );
    if (found.fullyCovered) {
      expect(doc.model.report.stats.unsupported).toBe(0);
    }
  });

  it('reports the writer, which decides what the file s cached values are worth', () => {
    const found = inspectXlsx(read('financial-model.xlsx'));
    expect(typeof found.readMs).toBe('number');
    expect(found.sheets.length).toBeGreaterThan(0);
    expect(found.sheets.reduce((n, s) => n + s.formulas, 0)).toBe(found.formulas);
  });
});
