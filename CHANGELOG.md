# Changelog

Both packages are released together at the same version. See
[VERSIONING.md](VERSIONING.md) — in particular, why adding a function is a
breaking change to what appears on screen.

## 0.2.0

The packaging release. `0.1.0` was never installable: both packages pointed
`main` at `./src/index.ts`, which works inside this repository and nowhere else.

### Breaking

- **`ExcelView` takes `doc` and `sheet` instead of `worksheet`, `model` and
  `sheetIndex`.** The `worksheet` prop was an `ExcelJS.Worksheet`, which put a
  22 MB dependency in the public type surface and made every caller reproduce
  `styled.getWorksheet(name) ?? styled.worksheets[i]` — a lookup that is
  non-obvious, easy to get wrong, and ours to do.

  ```diff
  - <ExcelView worksheet={ws} model={doc.model} sheetIndex={i} />
  + <ExcelView doc={doc} sheet={i} />
  ```

- **`resolveContent`'s `styled` option is now `content`**, taking a
  `CellContentData` from the layout rather than a raw ExcelJS value.
- **`format.ts`, `theme.ts` moved out of `view/`** and are re-exported from both
  the root and `/view`. Nothing under `view/` is imported by the root entry any
  more, so the base package pulls in no React.
- `PreviewDocument` gains `layouts` and `layoutMs`.

### Added

- **A build.** Both packages compile to `dist/` as ESM with declarations, and
  are reachable through an `exports` map. `npm run smoke` imports them by their
  published specifiers in plain Node — no bundler, no alias, no TypeScript — and
  puts a workbook through them.
- **`LICENSE`.** Both manifests claimed MIT with no licence text anywhere.
- **`@xlscalc/xlsx-preview/view/style.css`.** The stylesheet used to live in the
  demo app, which is private and never published — so an installed copy rendered
  `⚠` as unstyled text the eye slides past. The one mark that says "this is not
  a number we were willing to guess" was invisible to every consumer.
  Restyling is via custom properties (`--xl-warn`, `--xl-paper`, …).
- **`SheetLayout`** — a sheet's appearance as plain data: widths, heights,
  merges, frozen panes, and a deduplicated style table. On the largest workbook
  of the real corpus, 232,211 styled cells collapse to 912 distinct styles.
  Built once at load rather than per render, where it used to be one
  `getCell()` per coordinate on every React render.
- **`@xlscalc/xlsx-preview/worker`** — `createPreviewWorker()`, which runs the
  whole load off the main thread. This is what the layout work was for: a live
  `ExcelJS.Workbook`, the engine, and `model.cell` (a function) cannot cross a
  `postMessage`, and now nothing needs to. The 138,421-formula workbook spends
  7.5 s in `loadXlsx`; that is a frozen tab, and it is now a frozen worker.
  `precedents()` is asked over the same channel, since the engine stays there.
- **`toSnapshot` / `fromSnapshot`** — the flattening the worker uses, available
  on its own for caching a computed document. Errors are encoded rather than
  cloned: an `ExcelError` that crosses a `postMessage` loses its prototype, so
  `isErr()` would answer false and `#REF!` would render as `[object Object]`.
- **`inspectXlsx(buf)`** — what will be refused, without evaluating and without
  loading ExcelJS. 0.8 s against 7.5 s on the same workbook. It reports where
  refusal *starts*; the note in `inspect.ts` is explicit that this is a floor
  and not a prediction, because 33 such cells produced 64,809 warnings.
- **`renderToHtml(doc, sheet, opts)`** — the same grid as a string, for hosts
  that are not React. Shares `cellCss` with the component so the two cannot
  drift, and escapes everything, since every string in a workbook is input this
  library did not write.
- **`examples/node-headless.mjs`** — inspect, compute and render on a server.
- **`CHANGELOG.md`, `VERSIONING.md`**, per-package READMEs, `engines`,
  `publishConfig`, `keywords`, and `prepack` so a stale `dist/` cannot ship.

### Fixed

- **A workbook whose styling could not be parsed rendered nothing.** The load
  had been made non-fatal — three of ten real workbooks make ExcelJS throw while
  our own reader reads them fine — but the *renderer* still had no worksheet to
  draw from, so it drew nothing. `blankLayout()` gives it Excel's default
  sizing, and the extent is floored by the cells our own reader found. A
  document in default fonts is honest; a blank page is not.
- **`packages/formula-engine/src/parser.ts` contained two literal NUL bytes**,
  used as a sentinel for an omitted argument. It worked, and it made the file
  `data` rather than text to `file`, `grep`, `diff` and anything else that
  sniffs — an odd property for a parser to have. Written as `'\0missing'` now,
  and the duplicate literal replaced by the `MISSING_ARG` constant.
- **The demo typechecked against `dist/`**, so a breaking API change looked fine
  until someone rebuilt. It now resolves to source, like the tests, and
  `npm run typecheck` covers it.

### Fixed — a workbook could take the whole render with it

Ten bugs, found by a crash hunt rather than by the corpora, which is the point:
133 tests, an oracle at 100% and two eval corpora were all green while these were
live. **A green suite proves values, not survival.** Every one of them lived in a
path no fixture reached — a rare argument shape, an empty column, a file nobody
would write on purpose.

- **One ordinary formula rendered the entire workbook as nothing.**
  `evaluateAll` rethrew anything that was not an `Unsupported` or a
  `ParseError`, so a defect anywhere unwound the whole pass — the exact opposite
  of per-cell refusal. Two live routes into it: `=VLOOKUP(A1,T,2,)` (a trailing
  comma is an omitted argument, whose sentinel is a `Symbol`, which
  `v.toUpperCase` is not fond of) and `=MIN(A:A)` over more than ~125,000 rows
  (`Math.min(...ns)` is one JS argument per element, and V8 stops at about
  that). There is now a floor under both: an unforeseen throw is caught at the
  cell, reported in `gaps` as `INTERNAL`, and named as a defect rather than a
  decision.
- **`=NETWORKDAYS(1,50000000)` read as a hang.** Both it and `WORKDAY` walk one
  day at a time; a serial off the calendar is tens of millions of iterations.
  They refuse `#NUM!` now, which is what Excel does.
- **Three cells that were silently the wrong number**, all of which change output
  and so are named here as VERSIONING.md requires: `SUMIFS`, `AVERAGEIFS`,
  `MAXIFS` and `MINIFS` did not reshape a mis-sized aggregate range the way
  `SUMIF` does, so Excel's sum-range-is-a-corner rule applied to one of five
  siblings; `MATCH(">90 days", …, 0)` read an exact-match key as a criteria
  expression, so an ageing bucket matched the wrong row; and a reversed range
  written `$B2:A$1` left its `$` behind when normalised, which is invisible
  until a shared formula translates it.
- **`XLOOKUP(k,a,b,,0)` stored a `Symbol` in a cell**, where it survived until
  `structuredClone` refused it and the Worker's entire load rejected.
- The renderer had a second General-number implementation, and it was the wrong
  one: `1e20` displayed as `1.00000000000000e+2`. It delegates to the engine now,
  so there is one.
- A malformed numeric XML entity (`&#99999999;`) threw out of the *fatal* half of
  the load, failing a whole file for one bad character in one label.

### Fixed — what an untrusted file could do to the viewer

Found by a pre-publish review. Both are reachable from an ordinary `.xlsx`.

- **A `javascript:` hyperlink target reached the anchor.** Every string from the
  file was escaped at every insertion point — which stops attribute breakout and
  does nothing at all about the scheme, so one click ran script in the host's
  origin. Hyperlink targets are now judged by an allowlist (`http`, `https`,
  `mailto`) in `layout.ts`, where the file's bytes become our data, so a Worker
  clone and a third-party renderer inherit it. A blocked target still shows its
  text, marked, with the target on hover.
- **One stray cell in the last row of a sheet crashed the renderer.** A sheet's
  extent is the last row with anything in it, so a single accidental value in the
  bottom-right corner declares 1,048,576 × 16,384 — and both renderers paint
  every cell of the extent. Seventeen billion of them is an out-of-memory crash,
  measured, not a slow render. They now stop at `DEFAULT_RENDER_LIMITS` and
  caption what they left out: a truncated sheet that looks complete would be the
  same failure as a guessed number, one row further down.

### Fixed — packaging

- **`CAPABILITY.md` now ships inside `@xlscalc/formula-engine`.** VERSIONING.md
  makes diffing it between two versions the supported way to see what changed
  about what renders, and it was in the repository only — a mechanism nobody who
  installed the package could reach.
- **Source maps pointed at sources that were not in the tarball.** Both packages
  ship `src` now (minus tests), so a debugger lands on real code.
- `repository`, `bugs` and `homepage` on all three manifests.
- `examples/node-headless.mjs` used `import.meta.resolve`, which is unflagged
  only from Node 20.6 while both packages claim `>=18`.
- Doc corrections that a reader could have caught us on: the function count said
  197 where the registry says 204, the worst-case refusal root count said 204
  where the generated report says 33, the test count was three commits stale, and
  "MIT throughout" was true of the direct dependencies and not of ExcelJS's
  transitive tree — which is now described, advisories included.
