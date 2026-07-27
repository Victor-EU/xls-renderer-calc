# Changelog

Both packages — `@xlscalc/formula-engine` and `@xlscalc/xlsx-preview` — are
released together at the same version. See [VERSIONING.md](VERSIONING.md); in
particular, why adding a function is a minor bump and not a patch.

Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

---

## [0.4.0] — 2026-07-27

A minor bump for a bug fix, because the rule in [VERSIONING.md](VERSIONING.md)
is about what lands on someone's screen rather than about which file changed.
No computed value moved; a host that renders audit findings will show far fewer
of them.

### Fixed

- **`findHardcoded` reported findings that were not there — 900+ on the real
  corpus, of which essentially none were real.** "Neighbours" was never
  implemented: *every* formula cell in the literal's row or column counted as a
  candidate, so any shape shared by two of them was extrapolated over the whole
  line. On `apps/demo/public/financial-model-nocache.xlsx` — a file in this
  repository — that produced 33 findings and not one was correct. `B6`
  (`=B4+B5`) and `B18` (`=B16+B17`) are twelve rows apart and share a shape, so
  "each cell is the sum of the two above" was generalised down the column and
  applied *upward* onto the model's hand-entered inputs, announcing that a date
  serial in `B3` should have held `B1+B2` — which is 0, because `B1` and `B2`
  are empty title cells.

  A finding here names a cell and a number, so acting on a false one means
  editing a model that was correct. Three rules now stand in the way. A shape
  must be a **band**: its members occupy at least half their own span, nothing
  foreign sits inside that span, and the literal is inside the band or
  immediately *after* it — never immediately before, because the cell in front
  of a run of formulas is where models seed a series (an opening balance, a
  period-zero column) rather than where they get one wrong. And the literal must
  be in the same **series**: a formula that would read raw inputs, sitting beside
  formulas that read computed ones, is the first cell of another column of data,
  not a missing copy of its neighbours. That last rule also kills reconstructions
  whose inputs do not exist, which is how `=B1+B2` over two empty cells became a
  disagreement in the first place.

  On the ten-workbook real corpus: **900+ before (capped at 200 on four of the
  ten files), 23 after** — about two per workbook, with the right shape at last
  (a total that misses its own column sum by 167k; a stated `6731.67` where the
  formula gives `6731.69`; round `-20000` plugs where an average computes
  `-16849`).

  **This is not a precision guarantee.** Nobody has ground truth for those 23,
  and a separate corpus of machine-generated models measured every finding as
  false — generated models routinely stack a linked-data block above a
  typed-judgment block in one column, and treating a column as homogeneous is the
  assumption this whole detector rests on. Treat the output as a place to look,
  not as a defect list. The detector also had **no test at all**, which is how it
  shipped at that precision; it now has fourteen, including both demo workbooks.

  Values are unaffected — nothing about evaluation changed. What changes is which
  cells `PreviewModel.hardcoded` reports, so a host that renders those findings
  will show far fewer of them.

## [0.3.0] — 2026-07-27

### Added

- `hasCachedValue(cell)` — whether a raw cell carries a value or has never been
  computed. Exported beside `readXlsx` for anyone reading raw cells themselves.

### Fixed

- **`inspectXlsx(buf).uncached` reported more never-computed cells than the load
  it exists to predict.** An empty `<v>` means never computed *unless* `t="str"`,
  which says the formula ran and returned `""`. `bind` applied that rule and
  `inspect` did not, so every `IF(...,"",…)` was counted differently by the two.
  Both now call `hasCachedValue`, and a test holds them to the same answer on the
  same file. Over-reported only — no file was ever described as *more* computed
  than it was.
- **A module deleted from `src` could still ship in `dist`.** `tsc` never removes
  outputs whose sources are gone, so `prepack` alone did not deliver the "no
  stale `dist`" guarantee claimed in 0.2.0. The build now cleans first. Eight
  orphaned files were found in a local pack; they never reached the registry,
  because CI publishes from a fresh checkout.

---

## [0.2.0] — 2026-07-26

First release on npm. `0.1.0` was tagged in the repository but never published,
and could not have been: both packages pointed `main` at `./src/index.ts`, which
resolves inside this repository and nowhere else.

### Added

- **A build.** Both packages compile to `dist/` as ESM with declarations, behind
  an `exports` map. `npm run smoke` imports them by their published specifiers in
  plain Node — no bundler, no alias, no TypeScript.
- **`LICENSE`.** Both manifests claimed MIT with no licence text anywhere.
- **`@xlscalc/xlsx-preview/view/style.css`.** The stylesheet lived in the demo
  app, which is private, so an installed copy rendered `⚠` as unstyled text.
  Restyle via custom properties (`--xl-warn`, `--xl-paper`, …).
- **`SheetLayout`** — a sheet's appearance as plain data: widths, heights,
  merges, frozen panes, deduplicated styles. Built once at load rather than one
  `getCell()` per coordinate per render.
- **`@xlscalc/xlsx-preview/worker`** — `createPreviewWorker()` runs the whole
  load off the main thread. The 138,421-formula workbook takes 7.5 s in
  `loadXlsx`; that is now a frozen worker rather than a frozen tab.
- **`toSnapshot` / `fromSnapshot`** — flatten a computed document so it survives
  `postMessage`. Errors are encoded, not cloned: an `ExcelError` loses its
  prototype in transit, so `#REF!` used to arrive as `[object Object]`.
- **`inspectXlsx(buf)`** — what will be refused, without evaluating and without
  loading ExcelJS. 0.8 s against 7.5 s on the same workbook. Reports where
  refusal *starts*; it is a floor, not a prediction.
- **`renderToHtml(doc, sheet, opts)`** — the same grid as a string, for hosts
  that are not React. Shares `cellCss` with the component so the two cannot
  drift.
- `examples/node-headless.mjs` — inspect, compute and render on a server.
- `CHANGELOG.md`, `VERSIONING.md`, per-package READMEs, `engines`,
  `publishConfig`, `keywords`, and `prepack`.

### Changed — breaking

- **`ExcelView` takes `doc` and `sheet`** instead of `worksheet`, `model` and
  `sheetIndex`. The old `worksheet` prop was an `ExcelJS.Worksheet`, which put a
  22 MB dependency in the public type surface.

  ```diff
  - <ExcelView worksheet={ws} model={doc.model} sheetIndex={i} />
  + <ExcelView doc={doc} sheet={i} />
  ```

- **`resolveContent`'s `styled` option is now `content`**, taking a
  `CellContentData` from the layout rather than a raw ExcelJS value.
- **`format.ts` and `theme.ts` moved out of `view/`.** Nothing under `view/` is
  imported by the root entry any more, so the base package pulls in no React.
- `PreviewDocument` gains `layouts` and `layoutMs`.

### Fixed — wrong numbers

Each of these changed a rendered value, and VERSIONING.md requires naming them.

- **`SUMIFS`, `AVERAGEIFS`, `MAXIFS`, `MINIFS` did not reshape a mis-sized
  aggregate range** the way `SUMIF` does, so Excel's sum-range-is-a-corner rule
  applied to one sibling in five.
- **`MATCH(">90 days", …, 0)` read an exact-match key as a criteria expression**,
  so an ageing bucket matched the wrong row.
- **A reversed range written `$B2:A$1` lost its `$` when normalised** — invisible
  until a shared formula translates it.
- **The renderer had a second General-number implementation**, and it was the
  wrong one: `1e20` displayed as `1.00000000000000e+2`. It delegates to the
  engine now.

### Fixed — a workbook could take the whole render with it

Found by a crash hunt, not by the corpora: 133 tests, an oracle at 100% and two
eval corpora were all green while these were live.

- **One ordinary formula could render the entire workbook as nothing.**
  `evaluateAll` rethrew anything that was not an `Unsupported` or `ParseError`,
  so a defect anywhere unwound the whole pass. Reachable via
  `=VLOOKUP(A1,T,2,)` (a trailing comma is an omitted argument) and `=MIN(A:A)`
  over ~125,000+ rows (`Math.min(...ns)` exhausts the argument limit). An
  unforeseen throw is now caught at the cell and reported as `INTERNAL`.
- **`=NETWORKDAYS(1,50000000)` read as a hang.** It and `WORKDAY` walk one day at
  a time. Both refuse `#NUM!` now, as Excel does.
- **`XLOOKUP(k,a,b,,0)` stored a `Symbol` in a cell**, where it survived until
  `structuredClone` refused it and the Worker's entire load rejected.
- **A malformed numeric XML entity (`&#99999999;`) failed a whole file** for one
  bad character in one label.
- **A workbook whose styling could not be parsed rendered nothing.** Three of ten
  real workbooks make ExcelJS throw. `blankLayout()` now gives the renderer
  Excel's default sizing — a document in default fonts is honest, a blank page is
  not.

### Security

Both reachable from an ordinary `.xlsx`; found by a pre-publish review.

- **A `javascript:` hyperlink target reached the anchor.** Every string was
  escaped at every insertion point, which stops attribute breakout and does
  nothing about the scheme. Targets are now judged by an allowlist (`http`,
  `https`, `mailto`) in `layout.ts`, where the file's bytes become our data, so
  the Worker path and third-party renderers inherit it. A blocked target still
  shows its text, marked, with the target on hover.
- **One stray cell in the last row of a sheet crashed the renderer.** A single
  value in the bottom-right corner declares an extent of 1,048,576 × 16,384, and
  both renderers painted every cell of it — a measured out-of-memory crash, not a
  slow render. They stop at `DEFAULT_RENDER_LIMITS` and caption what they left
  out.

### Fixed — packaging

- **`CAPABILITY.md` now ships inside `@xlscalc/formula-engine`.** VERSIONING.md
  makes diffing it the supported way to see what changed about what renders, and
  it was in the repository only.
- **Source maps pointed at sources that were not in the tarball.** Both packages
  ship `src` now (minus tests).
- **`packages/formula-engine/src/parser.ts` contained two literal NUL bytes**,
  used as a sentinel for an omitted argument. Written as `'\0missing'` now, so
  the file is text to `grep`, `diff` and anything else that sniffs.
- **The demo typechecked against `dist/`**, so a breaking API change looked fine
  until someone rebuilt.
- `examples/node-headless.mjs` used `import.meta.resolve`, unflagged only from
  Node 20.6 while both packages claim `>=18`.
- `repository`, `bugs` and `homepage` on all three manifests.
- Doc corrections a reader could have caught us on: the function count said 197
  where the registry says 204, the worst-case refusal count said 204 where the
  generated report says 33, the test count was three commits stale, and "MIT
  throughout" was true of the direct dependencies but not of ExcelJS's transitive
  tree.

---

## 0.1.0 — never published

The registry has 0.2.0 and 0.3.0 only.

[0.3.0]: https://github.com/Victor-EU/xls-renderer-calc/releases/tag/v0.3.0
[0.2.0]: https://github.com/Victor-EU/xls-renderer-calc/releases/tag/v0.2.0
