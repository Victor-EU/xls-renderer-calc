# @xlscalc/xlsx-preview

A client-side, **view-only** `.xlsx` preview that computes formula values in the
browser. Nothing is uploaded.

```bash
npm install @xlscalc/xlsx-preview
```

## The problem it solves

A workbook written by anything other than Excel — an agent, a script, a
reporting job — emits formulas with an empty value cache:

```xml
<c r="B7"><f>SUM(B4:B6)</f><v></v></c>
```

Excel computes on open. Every other consumer reads the cache and shows nothing,
so a generated financial model renders as a labelled skeleton with no numbers in
it. This package computes those values locally, and **marks anything it could
not compute rather than guessing**.

## Render it

```tsx
import { useEffect, useState } from 'react';
import { loadXlsx, type PreviewDocument } from '@xlscalc/xlsx-preview';
import { ExcelView } from '@xlscalc/xlsx-preview/view';
import '@xlscalc/xlsx-preview/view/style.css';   // required — see below

function Preview({ file }: { file: File }) {
  const [doc, setDoc] = useState<PreviewDocument | null>(null);

  useEffect(() => {
    void file.arrayBuffer().then(loadXlsx).then(setDoc);
  }, [file]);

  if (!doc) return null;
  return <ExcelView doc={doc} sheet={0} />;
}
```

That is the whole integration. `doc.sheets` gives you the tabs, and
`doc.model.report.stats` tells you what happened:

```ts
const { formulas, computed, cached, unsupported, circular } = doc.model.report.stats;
```

**The stylesheet is not optional.** `⚠` is how this library says it refused to
compute a cell rather than inventing a number, and unstyled it is text the eye
slides past. Everything worth changing is a custom property — `--xl-warn`,
`--xl-paper`, `--xl-ink`, `--xl-mismatch-outline` — so restyling never means
copying selectors.

## Do it off the main thread

Loading is synchronous once it starts. On a 138,000-formula workbook that is
about 7.5 seconds, and a preview that freezes the tab for 7 seconds is not a
preview.

```tsx
import { createPreviewWorker } from '@xlscalc/xlsx-preview/worker';

const preview = createPreviewWorker();
const doc = await preview.load(await file.arrayBuffer());

// Same component, same props.
<ExcelView doc={doc} sheet={0} />
```

Vite and webpack both bundle the worker entry automatically. The engine stays in
the worker, so tracing what a cell reads is a question rather than a property:

```ts
const cells = await preview.precedents(sheet, row, col);
```

This is also where the dependency weight goes. In the demo's production build,
the main-thread chunk is 256 kB — React, the renderer and number formatting —
and everything that reads OOXML, ExcelJS included, is in the 1.0 MB worker
chunk, off the critical path.

## Ask before you pay

Whether to render a file at all is a decision better made *before* loading it.

```ts
import { inspectXlsx } from '@xlscalc/xlsx-preview';

const found = inspectXlsx(buf);      // 0.8 s where loadXlsx takes 7.5

found.fullyCovered;   // false
found.unsupported;    // [{ name: 'OFFSET', count: 384, supported: false }, …]
found.iterative;      // true — this workbook expects Excel to iterate cycles
found.writer;         // 'Microsoft Excel'
```

This reads the OOXML and parses every formula; it evaluates nothing and never
loads the styling dependency. It reports where refusal *starts* — read the note
in `inspect.ts` about why that is a floor rather than a prediction.

## Without React

```ts
import { renderToHtml } from '@xlscalc/xlsx-preview';

const html = renderToHtml(doc, 0, { showProvenance: true });
```

Same grid, same styles, as a string — for a server, a Vue or Svelte host, an
email, a PDF pipeline, or a test that wants to assert on markup. It shares its
style composition with the React component, so the two cannot disagree about a
cell. `examples/node-headless.mjs` in the repository is a working server-side
script.

Nothing in the root entry imports React. It is a peer dependency of
`/view` only, and an optional one.

## What to expect, honestly

**On real workbooks, about 36% of formula cells render `⚠` rather than a
number** — measured over 202,795 formula cells from ten workbooks nobody wrote
for this project.

Almost none of that is a missing function. Refusal propagates by design, so a
handful of cells using something unsupported can darken a large fraction of a
model: in the worst file, 33 such cells left 64,809 warning. Whether that is
acceptable depends entirely on your files, which is what `inspectXlsx` is for.

What you get in exchange, on the 128,976 cells it does answer: **zero
unexplained disagreements**. 3,050 of those cells deliberately differ from the
value the file itself stored — the writer was Google Sheets, which reads a blank
reference differently from Excel, and we follow Excel — and each falls under a
named rule with an exact expected count, gated in both directions.

Other limits worth knowing before you adopt:

- **View-only.** No editing, no write-back, no charts, no pivot tables.
- **The grid is drawn in full — there is no virtualisation.** Every cell of a
  sheet's extent becomes a `<td>`, and a sheet's extent runs to the last row with
  anything in it, so one stray value in the bottom-right corner of a sheet
  declares a million rows. Both renderers stop at `DEFAULT_RENDER_LIMITS`
  (~150,000 cells) and caption what they left out rather than freezing the tab.
  Pass `limits` to raise or lower it.
- **ExcelJS brings a large dependency tree**, and with it `npm audit` findings —
  currently high-severity ones in `minimatch` and `brace-expansion`, reached
  through ExcelJS's Node-side zip writing, which this package never calls. They
  are not exploitable through anything exported here, but they will show up in
  your install and in CI gates that fail on severity. `@xlscalc/formula-engine`
  on its own has no dependencies at all.
- **Styling comes from ExcelJS**, which is a large dependency and occasionally
  cannot read a file that our own reader can. When that happens the values are
  unaffected, `doc.stylesError` is set, and the sheet renders in default fonts —
  a document without its formatting is honest, a blank page is not.
- **No iterative calculation.** A workbook with `iterate="1"` has deliberate
  cycles; they are refused and reported as cycles, not as a broken model.
- **Custom themes** embedded in `xl/theme/theme1.xml` fall back to the default
  Office palette.
- **Hyperlinks are followed only for `http`, `https` and `mailto`.** A `.xlsx`
  can name any target it likes, `javascript:` included, and escaping the URL does
  nothing about the scheme. Anything else renders as marked text with the target
  on hover — the cell still says what it said, it just is not a link. Relative
  file paths are in that group too: they point at the author's disk, not at
  anything a browser can reach.

## API

| export | what it does |
|---|---|
| `loadXlsx(buf, opts?)` | read, evaluate and lay out a workbook |
| `inspectXlsx(buf, opts?)` | what will be refused, without evaluating |
| `renderToHtml(doc, sheet, opts?)` | the grid as an HTML string |
| `toSnapshot` / `fromSnapshot` | flatten a document so it survives `postMessage` |
| `layoutSheet`, `blankLayout`, `mergeMap`, `paneOffsets` | the layout model |
| `renderExtent`, `truncationNote`, `DEFAULT_RENDER_LIMITS` | how much of a sheet a renderer will draw, and what to say about the rest |
| `isSafeHref` | whether a hyperlink target may be navigated to |
| `findHardcoded` | literals whose neighbours' formulas contradict them |
| `readXlsx` | the raw OOXML read, with no evaluation |
| `hasCachedValue` | whether a raw cell carries a value, or has never been computed |
| `/view` → `ExcelView`, `resolveContent`, `cellCss` | the React renderer |
| `/worker` → `createPreviewWorker` | the off-thread path |

`loadXlsx` accepts `{ now }` to pin `NOW()`/`TODAY()`, which makes a render
reproducible — worth doing anywhere the output is cached, diffed or tested.

## Requirements

Node 18+ or any modern browser. ESM only. TypeScript declarations included.
React ≥18 is an optional peer dependency, needed only for `/view`. MIT.
