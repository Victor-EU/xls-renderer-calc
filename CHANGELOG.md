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
  and not a prediction, because 204 such cells produced 64,809 warnings.
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
