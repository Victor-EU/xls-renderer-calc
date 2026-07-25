# Recalculate & render — the calculation layer for the client-side Excel renderer

**Design date:** 2026-07-25 · **Built:** 2026-07-25 (commit `2e5822d`)
**Status: BUILT.** This document describes the system as it exists and why it is
shaped that way. `README.md` covers what it does and how to run it;
`CAPABILITY.md` (generated) lists exactly what it computes.

Marker convention: **✓** = verified against real files or measured ·
**✗** = the original design was wrong here, corrected with the reason ·
**?** = still open.

> **Read this first if you knew the earlier draft.** Three of its load-bearing
> decisions did not survive contact with the build, and each is corrected in
> place below rather than quietly edited away:
>
> | Original | What happened |
> |---|---|
> | §0/§5/App. A — build a **closed `{SUM, IF}` grammar** for our own files | ✗ The vocabulary is not closed. Built broad (197 functions) and gated by an oracle instead — §2.1, §5 |
> | §8.2 — range dependencies are "**not a v1 concern**" | ✗ Materialising the graph is O(rows²) and ran out of memory at 45k formulas. The graph is no longer materialised — §8 |
> | §12.5 — LibreOffice-vs-Excel divergence is a **financial-function** risk | ✗ Real, but the largest class was **booleans** — §12.5 |
>
> Everything else — overlay-not-mutation, the fail-loud doctrine, the raw-`t`
> probe, the audit diff — was built as specified and holds up.

---

## 0. The decision, as it landed

Build the evaluator **from scratch**, make it **broad**, and gate it with a
**differential oracle**. Do not adopt a wasm engine. Keep LibreOffice headless as
an explicit, user-consented escalation for files this engine refuses.

The original reasoning reached "from scratch" via a different route — that our
own generator's vocabulary was closed and tiny, so a few hundred lines would
cover it completely, while an adopted engine would handle the open tail of
arbitrary user files. The first half turned out to be false (§2.1), which removes
the argument for a *small* engine but not the argument for our own one. What
actually decides it:

1. **Fail-loud requires introspection.** The doctrine in §3 is worth nothing
   unless the engine can say precisely what it could not do. An adopted engine
   returns a value; it cannot tell you which of its answers you should not trust.
   That property has to be built in, not wrapped around.
2. **Breadth is safe when refusal is the default.** A large hand-written library
   would be reckless if a gap produced a wrong number. Because a gap produces ⚠
   instead, breadth costs coverage risk, not correctness risk — and the oracle
   turns coverage risk into a measured number.
3. **Zero dependencies, zero bundle surprise.** The engine has no dependencies at
   all, runs in Node, a worker or the browser, and is debuggable in the same
   language as its caller.

The `formualizer` bake-off (§4) was **not run**. It was designed as the gate
before writing our own engine, and skipping a planned measurement deserves to be
stated plainly rather than glossed: the decision was made on the grounds above,
which are about the shape of the requirement rather than about the candidate's
quality. The interface it was meant to protect (§6.1) exists, so the measurement
can still be made later if bundle weight or coverage ever argues for it.

---

## 1. The problem, mechanically

*(Unchanged from the original — it was right.)*

A cell carries a formula and, separately, a **cached result**:

```xml
<c r="B7" s="19" t="n"><f>B6/B4</f><v>0.59004854368932</v></c>
```

`<f>` is the formula. `<v>` is what Excel computed when it last saved. Every
consumer that is not a spreadsheet application reads `<v>`.

Agent-generated files emit the formula with an **empty** cache ✓:

```xml
<c r="B7" s="19"><f>B6/B4</f><v></v></c>
```

Excel and LibreOffice do not care — they compute on open. Our renderer, and every
other cache-reading consumer, shows blank. In a financial model nearly every
meaningful number is a formula, so the document renders as a labelled skeleton.

The previous fix, `soffice --headless --convert-to xlsx`, works and stays as a
fallback, but costs the preview its defining properties:

| Property | LibreOffice recalc | Browser recalc |
|---|---|---|
| No server | ✗ needs a binary and a process | ✓ |
| Nothing leaves the browser | ✗ the file must be uploaded | ✓ |
| Fast | ✗ ~1–3 s per file | ✓ ~100 ms for 25,000 formulas ✓ measured |

---

## 2. What measurement established

### 2.1 ✗ The formula vocabulary is **not** closed

The original draft probed the two sample files and found their entire function
vocabulary to be `SUM` and `IF` — 23 calls across 76 formulas — and built its
central decision on that: a closed grammar, 100 % coverage by construction, an
allowlist published into the generator's prompt.

It flagged its own caveat ("n=1 model, written by one generator") and then treated
the number as load-bearing anyway. It should not have been. The generator path in
the product is explicit that structured export writes *values*, and that a
formula-driven workbook comes from **`python_execute` + openpyxl written
free-hand by a model** (`apps/api/src/tools/excel-export.ts:70`). The vocabulary
is therefore whatever an LLM types: `ROUND`, `IFERROR`, `NPV`, `IRR`, `XIRR`,
`SUMIF`, `INDEX`/`MATCH`, `EOMONTH`, `TEXT`. A `{SUM, IF}` engine would have
rendered ⚠ on the first real model.

**What replaced it:** 197 implemented functions, chosen for what a generated
financial model reaches for, with every one of them measured against LibreOffice
(§12). The closed-grammar *mechanism* survived in a better form —
`CAPABILITY.md` is generated from the registry and can still be published into
the generator's prompt as an allowlist. The floor is declared; it is just wider,
and it is derived from the code rather than from a sample of one.

The general lesson is worth keeping: a histogram over one artefact measures the
artefact, not the population. The corpus is still synthetic and this remains the
weakest evidence in the project (§16).

### 2.2 ✓ The per-cell ambiguity is resolvable — and the heuristic is gone

The old renderer carried a documented limitation: a formula that legitimately
computes to `""` is indistinguishable, per cell, from one that was never
computed, so it judged at file level with `RECALC_THRESHOLD = 0.25`.

True of ExcelJS, not of the file. The raw `t` attribute decides it ✓:

| Case | Raw XML | `t` |
|---|---|---|
| Computed to an empty string | `<c r="G9" t="str"><f>IF(G9=0,"",F9/G9-1)</f><v></v></c>` | **`str`** |
| Never computed | `<c r="F4"><f>SUM(B4:E4)</f><v></v></c>` | **absent** |

`packages/xlsx-preview/src/ooxml.ts` reads the sheet parts directly and decides
per cell. The heuristic is deleted, and the file-level count it approximated
(`facts.uncached`) is now exact.

The same pass earns its keep twice more. It recovers **shared formulas**
correctly — OOXML stores `<f t="shared" ref="B2:E2" si="0">` once and siblings
reference it by `si`, so a reader that returns the master text verbatim reports
four cells as computing the first cell's numbers; here the master's AST is
translated by each sibling's offset. And it surfaces array-formula masters and
the `date1904` flag, both of which ExcelJS obscures.

### 2.3 ✓ `fullCalcOnLoad`, and who sets it

| File | `<calcPr>` |
|---|---|
| `-nocache.xlsx` (openpyxl) | `<calcPr calcId="124519" fullCalcOnLoad="1"/>` |
| `financial-model.xlsx` (after LO) | `<calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.0001"/>` |

openpyxl already sets it, so a user who *downloads* the un-recalced file and
opens it in Excel sees correct numbers. The blanks were our renderer's problem
specifically. `readXlsx` reports the flag (`facts.fullCalcOnLoad`) but does not
act on it; it matters only for the write-back path (§13.2), which is not built.

### 2.4 ✓ Precision, as actually written

LibreOffice wrote `0.59004854368932` for `B6/B4` — 14 significant digits, where
the IEEE-754 double carries 17. This sets the comparison policy everywhere:
floats are compared with a **relative** epsilon of `1e-9`, never `===`. It is
used identically by the oracle comparator, the computed-vs-cached diff
(`valuesAgree` in `workbook.ts`) and the audit pass.

### 2.5 ✓ No `calcChain.xml` in either file

Neither sample has one. `readXlsx` reports `facts.hasCalcChain` so the write-back
path can strip a stale one rather than author it.

### 2.6 ✓ Measured, after the build

| | Before | After |
|---|---|---|
| Nocache sample | 75/75 formulas blank | 75/75 computed, 0 unsupported, ~1 ms |
| Cached vs computed render | — | **all 232 rendered cells identical** ✓ browser-verified |
| Parse (ExcelJS + raw pass) | 74 ms | 39–89 ms |
| 25,000 local-arithmetic formulas | — | ~100 ms |

---

## 3. Failure modes, ranked

*(Unchanged. This ordering is the ethical content of the design, and everything
else is subordinate to it.)*

| # | Failure | Severity | Mitigation, as built |
|---|---|---|---|
| 1 | **Renders a wrong number confidently** | **Catastrophic** — invisible, and it is a financial figure | §10 doctrine; refusals throw rather than return; oracle gates false confidence at zero (§12) |
| 2 | Renders blank | Bad | The whole point of this work — fixed |
| 3 | Renders ⚠ where Excel has a value | Acceptable | Honest. Counted, surfaced, drives the backlog |
| 4 | Slow or blocks the UI | Annoying | §9 — measured; the worker is deferred with a stated reason |
| 5 | Bundle bloat for files that need no calc | Annoying | The engine is dependency-free; no wasm was adopted |

The asymmetry between #1 and #3 is the entire argument. An engine that says "I
don't know" is a tool. An engine that guesses is a liability, and in a model the
guess is unfalsifiable by eye.

---

## 4. Engine landscape (verified 2026-07-25, not re-checked since)

| Package | License ✓ | Version ✓ | Notes |
|---|---|---|---|
| `formualizer` | **MIT OR Apache-2.0** | 0.7.1 | Rust→wasm, parser + evaluator + graph + xlsx ingest. The lead candidate if one were adopted |
| `@ironcalc/wasm` | MIT/Apache-2.0 | 0.7.0 | Rust→wasm; README self-describes as early-stage |
| `hyperformula` | **GPL-3.0-only** | 3.3.0 | The mature option, **disqualified by license** — not by quality |
| `@formulajs/formulajs` | MIT | 4.6.0 | ~500 Excel functions as plain JS. **No parser, no refs, no graph.** An ingredient |
| `fast-formula-parser` | MIT | 1.0.19 | Right architecture, last published ≈6 years ago |

The table stands as a record. Nothing from it was adopted; `@formulajs/formulajs`
was not used even as a source of function bodies, because the value in this
engine is the coercion and argument-mode rules *around* the functions (§7), not
the arithmetic inside them.

---

## 5. ✗ The reframe that collapsed

The original framed this as two problems with opposite answers: **A**, our own
generator's output, closed and declarable, answered by a small from-scratch
evaluator; and **B**, arbitrary user files, open and unbounded, answered by an
adopted engine.

The distinction dissolved when §2.1 showed Problem A's vocabulary is open too.
Both are the same problem, and it is not "which functions" — it is **which
behaviours**. What survives, and is what the build is actually organised around:

> *Guarantee the floor; don't cap the ceiling.* The floor is declared
> (`CAPABILITY.md`), enforced by refusal rather than by hope, and can be
> published into the generator's prompt. The ceiling stays open through the
> LibreOffice escalation (§13.1) for anything the floor does not cover.

---

## 6. Architecture

### 6.1 Overlay, not mutation ✓ built as designed

```
buf ─┬─► ExcelJS ─────────────► styles, fonts, fills, borders, merges,
     │   (parse.ts)             column widths, frozen panes
     │
     └─► ooxml.ts ────────────► formulas, raw `t`, shared-formula
         (fflate + xml.ts)      translation, defined names, date1904
                    │
                    ▼
              bind.ts ──► Workbook (formula-engine) ──► evaluateAll()
                    │
                    ▼
              ValueOverlay: value + provenance + the file's own value
                    │
                    ▼
              ExcelView.tsx / audit.ts
```

**Computed values never overwrite the file's cached ones.** Three things depend
on that, and all three shipped:

1. **Provenance.** Every rendered number knows where it came from. In a financial
   preview that is the difference between a view and a claim.
2. **The diff.** Cells with both a cached and a computed value can be compared —
   a genuine audit feature (§11), impossible if you overwrite.
3. **Swappability.** Engines stay interchangeable because none of them share
   mutable state with the renderer.

**Two passes over the same bytes**, each doing what it is best at. ExcelJS's
style fidelity is proven — it is the renderer the earlier spike validated cell by
cell against a real model — and rewriting `styles.xml` handling would be
regression risk with no upside. But it drops the `t` attribute, so it is not
trusted for values. Both passes get their own copy of the buffer: ExcelJS
detaches the ArrayBuffer it is handed, and a detached buffer read afterwards
yields zero bytes rather than an error.

### 6.2 Types, as built

```ts
type Provenance =
  | 'literal'      // straight from <v>, no formula
  | 'cached'       // read from the file's formula cache; not recomputed
  | 'computed'     // we evaluated it
  | 'unsupported'  // outside capability — renders ⚠, never a number
  | 'circular'     // part of, or downstream of, a reference cycle
  | 'volatile';    // computed, but depends on NOW/TODAY/RAND

interface CellRecord {
  value: Scalar;         // number | string | boolean | ExcelError | null(blank)
  provenance: Provenance;
  cached?: Scalar;       // retained when the file had one → feeds §11
  reason?: string;       // why unsupported — the backlog, verbatim
}
```

The original sketched a `RecalcEngine` interface with `load` / `evaluateAll` /
`setValue` / `setFormula` / `trace`. The built shape is the `Workbook` class
carrying the same surface — `evaluateAll()`, `precedents()` for trace,
`tryEvaluate()` for speculative evaluation — with `setValue`/`setFormula` as
authoring rather than incremental-edit entry points, because incremental recalc
(§9.2) is not built.

One type decision the original did not call out, and that carries more weight
than any function: **blank is `null`, distinct from `""` and from `0`.** An empty
cell is 0 in arithmetic, equals *both* 0 and `""` in comparison, and is invisible
to `COUNT`. `""` does none of that. Collapsing them is a whole class of silent
wrong answers — and keeping them apart is exactly what makes the sample model's
own `IF(G9=0,"",F9/G9-1)` guard work over a column with no prior-year data.

### 6.3 Module layout, as built

```
packages/formula-engine/src/
  values.ts       the value model — blank, errors, Matrix, RefValue
  errors.ts       Unsupported (a refusal) vs ExcelError (a value)
  a1.ts           addressing, 1-based throughout, id packing
  lexer.ts        tokenizer
  parser.ts       Pratt parser; Excel precedence incl. the two traps
  ast.ts          nodes, walk, translate (shared formulas), unparse
  coerce.ts       Excel's coercion rules — the highest-risk surface
  interpreter.ts  evaluation, error propagation, array broadcasting
  workbook.ts     storage, evaluation order, cycles, provenance
  functions/      registry + math, stats, logical, text, info,
                  lookup, conditional, dates, financial, format-hook

packages/xlsx-preview/src/
  xml.ts          minimal XML pull parser
  ooxml.ts        raw formula reader — the `t` attribute, shared formulas
  bind.ts         file → engine → ValueOverlay
  audit.ts        the hardcoded-cell detector
  parse.ts        orchestration (ExcelJS + raw)
  view/           ExcelView.tsx, format.ts (numfmt), theme.ts

apps/demo/        the preview at :5176
tools/oracle/     suites.py, generate.py, divergences.json, oracle.test.ts
tools/verify/     drive.mjs (browser E2E), scale.mjs, shot.mjs
```

---

## 7. Evaluation semantics — where compatibility actually lives

Function *count* is a vanity metric; compatibility lives in coercion, precision
and error propagation. This section was the original's best contribution and is
kept as the checklist it was, annotated with what shipped.

### 7.1 ✓ Operator precedence, including the oddities

Implemented in `parser.ts` and asserted twice — in `parser.test.ts` for structure
and in the oracle for value:

- **`-2^2 = 4`** — unary minus binds *tighter* than `^`, unlike mathematics and
  unlike every C-family language.
- **`2^3^2 = 64`** — `^` is **left**-associative in Excel.

Both produce a plausible wrong number rather than an error when got wrong, which
is why they are pinned in two places.

Intersection (space) and union (comma) were slated for rejection; both are
**implemented** instead, because once references are a first-class value type
they are a few lines each and `#NULL!` becomes expressible.

### 7.2 ✓ Type coercion — the highest-risk area

All implemented in `coerce.ts` and probed by the `coercion` suite:

| Situation | Excel | Note |
|---|---|---|
| `"5"+1` | `6` | arithmetic **does** coerce numeric text |
| `"5"=5` | `FALSE` | comparison does **not** — different rules per context |
| `TRUE+1` | `2` | booleans coerce in arithmetic |
| `TRUE=1` | `FALSE` | …but not in comparison |
| sort/compare order | number < text < boolean | cross-type comparison is by type rank |
| `"a"+1`, `""+1` | `#VALUE!` | an empty *string* is not an empty *cell* |

The boolean rows are also where LibreOffice disagrees with Excel (§12.5), which
made them the most consequential rows in the table.

### 7.3 ✓ Empty-cell semantics

Implemented, and the reason the model's `IF(G9=0,"",F9/G9-1)` renders a clean
blank instead of twelve `#DIV/0!`. One addition the original missed: a formula's
*final* value is never blank — `=J1` over an empty J1 is `0` in Excel, so blank
collapses to 0 at the cell boundary while staying distinct inside an expression.
That distinction was found by the oracle, not by reasoning.

### 7.4 ✓ Numbers and precision

- All arithmetic is IEEE-754 double.
- **`ROUND` rounds the decimal value, not the binary double.** `ROUND(2.675,2)`
  is `2.68`; in IEEE-754 2.675 is really 2.67499999999999982, so
  `Math.round(2.675*100)/100` gives 2.67 — a one-cent disagreement in any model
  with rounded currency. `excelRound` rounds the 15-significant-digit decimal
  string instead. The original correctly declined to emulate Excel's cosmetic
  near-zero rounding, and that decision stands.
- **General-format text switches to scientific at `1E16` and `1E-15`** ✓
  measured by bracketing probes, not assumed. The original had no position on
  this; a guess would have been wrong by ten orders of magnitude.

### 7.5 ✓ Dates

Serials, with the **1900 leap-year bug reproduced deliberately** — serial 60 is a
phantom 1900-02-29, and an engine that "fixes" it disagrees with Excel on every
date in the file. The **1904 system is supported**, not refused as the original
proposed: reading the flag is trivial and the failure mode of ignoring it is a
silent four-year error, which is a failure-mode-#1 outcome for a one-line saving.
Serial → `Date` conversion happens only at the render boundary.

### 7.6 ✓ Errors and propagation

Nine error values, propagating through operators and trapped by
`IFERROR`/`IFNA`/`ISERROR`/`ISNA`. The critical rule holds throughout:
**an Excel error is a legitimate computed value, not an engine failure.**
`#DIV/0!` renders as `#DIV/0!` with provenance `computed`. Only *our* inability
to evaluate produces `unsupported`. `errors.ts` makes the distinction structural —
a refusal is a thrown `Unsupported`, which cannot accidentally become a cell
value.

### 7.7 References

| Shape | Status |
|---|---|
| Relative / absolute / mixed | ✓ implemented |
| Range, cross-sheet with quoted names | ✓ implemented |
| Whole column / row (`A:A`, `1:1`) | ✓ implemented — clamped to the used range |
| Defined names | ✓ implemented, read from `xl/workbook.xml`, self-reference guarded |
| Intersection, union | ✓ implemented (originally slated for rejection) |
| 3D refs, structured table refs, external links | ✗ refused with a reason |
| `INDIRECT` / `OFFSET` | ✗ refused — they break the static dependency graph |

Whole-column refs and defined names were "v2" in the original; both are cheap
once the binder exists, and `SUM(A:A)` is common enough in generated models that
refusing it would have been a visible gap.

### 7.8 ✓ Function-level semantics that bite

- **`SUM` ignores text and booleans inside a range but coerces a direct
  argument.** Encoded once in `collectNumbers` so every aggregate shares it,
  rather than re-derived per function.
- **`IF` is lazily evaluated** — as are `IFERROR`, `IFNA`, `IFS`, `CHOOSE`,
  `SWITCH`. `IFERROR` is lazy for a second reason the original did not
  anticipate: an *unsupported* fallback must not poison a cell whose primary
  value computed cleanly.
- **Criteria strings** (`">100"`, `"<>x"`, wildcards) are a miniature expression
  language ✓ implemented once in `makeCriteria` and shared by the whole `*IF`
  family.
- **`VLOOKUP` approximate match** — the original said ship exact-only and reject
  approximate. Built differently: the *sorted contract is implemented exactly*,
  and **unsorted input is refused**. Approximate match is well-defined on sorted
  data and common in real models (rate tables, grade bands), so refusing it
  wholesale would have been a large gap; but on unsorted data Excel returns
  whatever its binary search's probe order lands on, and reproducing an arbitrary
  answer is failure mode #1. Refusing exactly the undefined case is strictly
  better than refusing the whole function.

### 7.9 Shared and array formulas

- **Shared formulas** ✓ handled in `ooxml.ts` by translating the master AST by
  each sibling's offset. This closes risk **R7** — rather than testing whether
  ExcelJS translates them, the raw reader does the translation itself, so the
  question no longer arises.
- **Array formulas and dynamic arrays** ✗ refused. A saved file already contains
  whatever the writer spilled, so refusing costs little; the masters are counted
  in `facts.arrayFormulas`.

### 7.10 ✓ Cycles

Detected and marked `circular`; **never iterated**. Excel refuses them by default
and iterative calc is opt-in, so a renderer that silently iterated would show
numbers nobody asked for. Both true cycles and self-references are caught (§8).

### 7.11 ✓ Volatile functions

`NOW`/`TODAY`/`RAND`/`RANDBETWEEN` are implemented and marked `volatile`. The
clock is read **once per evaluation run** and passed in, so a run is
reproducible. One policy the original did not specify: when a volatile cell has a
cached value, the preview shows the *file's* value. A preview should show what
the generator produced; recomputing `NOW()` would make it disagree with the file
for a reason that is not a finding. `OFFSET`/`INDIRECT` are refused outright.

---

## 8. ✗ Dependency graph — the design was wrong here

### 8.1 What the original specified

Intern addresses to integers; keep **forward edges** (precedent → dependents) and
**reverse edges**; Kahn topological sort; Tarjan for cycles. On range
dependencies it said: expanding `B4:E4` to four cells is fine, `SUM(A:A)` would be
a million edges so bound it to the used range, and block-bucketed range nodes are
"**not** a v1 concern — note it and move on."

### 8.2 Why that fails

It is not whole-column references that break it. It is the **running total**, and
every financial model has one:

```
D2 =SUM($B$1:B2)          2 edges
D3 =SUM($B$1:B3)          3 edges
…
D5000 =SUM($B$1:B5000)    5000 edges
```

One edge per precedent *cell* is O(rows²). At 5,000 rows that is 12.5 million Set
entries — hundreds of megabytes — and it took a 45,000-formula workbook from slow
to **out of memory** ✓ observed. The used-range bound does nothing here: every one
of those ranges is legitimately inside the used range.

The mistake was treating range expansion as a *size* problem to be capped, when
it is a *representation* problem. A range is one dependency, not N.

### 8.3 What was built

**The graph is never materialised.** For each formula, the precedent
*rectangles* are kept (`RefValue[]`, one per reference in the AST). Evaluation
order comes from an **iterative depth-first search** that walks those rectangles
lazily; its post-order is a topological order.

- **Memory is O(formulas)**, not O(edges). The 45k workbook now completes.
- **The walk allocates nothing per edge** — it advances a cursor
  `{rectIndex, row, col}` inside a frame.
- **Cycle detection falls out of the colouring**: grey-on-stack means a cycle,
  and the members are exactly the frames from that node upward. Tarjan is not
  needed for what the UI wants.

The search is **iterative rather than recursive** for the reason the original
gave for preferring a topological sort over lazy recursion in the first place —
a row-to-row chain (`revenue_t = revenue_{t-1} * growth`) is as deep as the model
is long, and 5,000 nested evaluations overflow the call stack. That reasoning was
right; it just applies to the DFS too. There is now a 5,000-deep chain test.

Time is still O(total precedent cells) — the running total genuinely reads 12.5 M
cells — but that is inherent to the formula and costs ~1.1 s rather than failing.

### 8.4 The dynamic-dependency escape hatch ✓ unchanged

`INDIRECT`/`OFFSET` create edges knowable only at evaluation time, so they are
refused at parse. That is exactly why refusing at *parse* time matters: the
precedent rectangles stay a static, provable object.

---

## 9. Real time

### 9.1 Measured, not budgeted

The original set targets; here are the numbers, by formula *shape* rather than by
count, because shape is what determines cost (`perf.test.ts`):

| Shape | Size | Evaluate | Original target |
|---|---:|---:|---|
| Local arithmetic — the normal case | 25,000 formulas | **~100 ms** | < 300 ms for 10k ✓ beaten |
| Row-to-row chain, 5,000 deep | 5,000 formulas | ~10 ms | — |
| Running total `SUM($A$1:A_n)` per row | 5,000 formulas | ~1.1 s | quadratic by construction |
| Full-table `VLOOKUP` per row | 2,000 formulas | ~0.39 s | quadratic by construction |
| Parse (ExcelJS + raw pass) | 3-sheet model | 39–89 ms | ~74 ms baseline ✓ held |

The bottom two are O(rows²) *cell reads* for any evaluator, Excel included. Two
hot-path fixes came out of measuring them: aggregates read ranges directly
instead of materialising a `Matrix` and allocating a wrapper object per cell, and
`VLOOKUP` reads only its search column instead of copying the whole table on
every call.

### 9.2 ✗ Not built: worker, incremental recalc, streaming

The original's §9 specified all three. None is built, and the honest reason
differs per item:

- **Web Worker.** Not a wrapper. `ExcelView` reads the ExcelJS worksheet
  directly for style, so moving parse and evaluation off-thread requires first
  flattening the styled grid into a transferable snapshot — a real refactor with
  fidelity-regression risk against the cell-by-cell match that is currently the
  project's strongest evidence. At ~100 ms for a 25,000-formula model the
  user-visible cost of not doing it is small; it becomes worth doing when a
  pathological sheet blocks for a second or more.
- **Incremental recalc.** The primitive exists (`tryEvaluate` runs a formula at a
  position without storing it) and the rectangles support dirty-set propagation,
  but nothing drives it because there is no edit surface.
- **Streaming.** Free once incremental exists, as the original said. Still true,
  still unbuilt.

### 9.3 ✓ Lazy loading

Moot as built: the engine has no dependencies and adds no wasm, so there is no
bundle weight to gate behind detection. The detection step (§2.2) was worth
having on its own merits.

---

## 10. The honesty layer ✓ built

### 10.1 Doctrine

> **A cell we cannot compute renders ⚠, never a number.**
> Unsupported function, refused ref shape, cycle, unsupported feature, or an
> uncomputable precedent → `unsupported` provenance, counted at file level,
> reason recorded verbatim.

Corollary from §7.6: genuine Excel errors are *computed values*. `#DIV/0!`
renders as `#DIV/0!`.

One rule the original did not state and the build required:
**uncomputable poisons downstream.** If a cell cannot be computed, every cell
that reads it is marked too. Without this, a `SUM` over a range containing one ⚠
cell would silently omit it and render a confident subtotal that is short one
input — failure mode #1 arriving through the back door of an otherwise honest
system. It is enforced at the read: `cellValue` throws while evaluating.

### 10.2 UX surface ✓

The banner changed meaning rather than shape, as designed:

| Before | After |
|---|---|
| ⚠ **This file needs a recalc.** 75 formula cells have no cached result | ✓ **Computed live in your browser.** 75 of 75 formulas · 0 unsupported · 1 ms |

Plus: status chips (computed / from file / unsupported / circular / disagrees
with file / parse ms / eval ms); a **Not computed** list aggregating gaps by
function with a sample address; per-cell hover showing formula, provenance,
computed value and the file's value when they differ; a computed/diff toggle;
and click-to-trace precedent highlighting via `Workbook.precedents()`.

---

## 11. Recalc as audit ✓ built, and extended

The original called this the sleeper feature and it was right — but it specified
only half of it.

**What it specified: the value diff.** A formula cell whose stored `<v>`
disagrees with what the formula computes. Built, with volatile cells and
within-epsilon differences excluded so it does not cry wolf.

**What it missed: on the case that matters most there is nothing to diff.** The
original listed "the agent hardcoded a total instead of summing it" first among
the findings and noted it was "structural, detectable without evaluating" — then
built no mechanism for it. A hardcoded cell has *no formula*, so no value
comparison can ever find it.

`audit.ts` detects it from the **shape of the neighbourhood**: a literal number
sitting where its row or column neighbours hold formulas. It infers what the cell
would have held by translating a neighbour's formula to that position and
evaluating it, so the finding is specific rather than a suspicion. It requires at
least two neighbours sharing one shape — a single neighbouring formula is a
pattern of one and would fire on every summary row. On the bundled sample:

> `Revenue!E6` is typed in as **761.2**, but the rest of the row (B6, C6, D6)
> uses `=SUM(E4:E5)`, which is **721.2**

This is the characteristic failure of a generated model: the writer computed one
figure in its head and typed it in. The number is plausible, in the right place,
formatted like its neighbours. Nothing about the rendered sheet reveals it.

**Trace** (`Workbook.precedents()`) is built as the one-hop version — click a
cell, its precedents highlight. The multi-hop chain narrative the original
sketched is not built.

---

## 12. The oracle ✓ built — and it earned its keep

### 12.1 The loop

```
suites.py ──► build/<suite>.xlsx ──► soffice --headless ──► expected.json
     │                                                            │
     └──────► spec.json ──► the engine builds the same workbook ──┘
                                          │
                                    per-probe diff ──► buckets
```

Both sides are driven from one spec, so a disagreement is a semantic difference
rather than a setup difference.

### 12.2 Buckets and the gate

| Bucket | Meaning |
|---|---|
| match | the two agree |
| unsupported | we refused, and said why — acceptable, tracked |
| divergence | they disagree and we deliberately follow Excel — declared |
| **MISMATCH** | they disagree and we did not know — **hard gate at zero** |

Current: **399 probes, 13 suites, coverage 100 %, accuracy 100 %, false
confidence 0**, with 29 declared divergences.

The original proposed coverage 100 % / accuracy 100 % / false-confidence 0 as v1
gates against a closed grammar. Only the third is a real gate — the first two
describe capability against *this corpus* and would drop the moment a suite
widens. False confidence is the one that describes trustworthiness.

### 12.3 ✓ Tolerance policy, unchanged

Relative epsilon `1e-9` for numbers; exact for strings, booleans and errors.
Blank and `""` are distinct outcomes and the comparator must not conflate them —
"the comparator is itself the first place that bug would hide" was a good call.
(One limit found in practice: openpyxl cannot distinguish a formula that produced
`""` from one that produced nothing, so on the oracle side both satisfy a blank
expectation. The engine still distinguishes them internally.)

### 12.4 Fixture hazards worth knowing

- **openpyxl writes formula text verbatim**, so post-2007 functions need the
  `_xlfn.` namespace or LibreOffice returns `#NAME?` — which looks exactly like a
  divergence and is really a fixture bug. `generate.py` adds it.
- LibreOffice applies a date or time *format* to `DATE()`/`TIME()` results, so
  openpyxl hands them back as datetimes; they are converted back to serials
  before comparison.

### 12.5 ✗ LibreOffice is not Excel — and the gap is not where we expected

The original anticipated divergence on "financial functions, edge-case rounding,
and newer functions", and recommended hand-verified Excel fixtures for `IRR`,
`XIRR`, `NPV`, `PMT`.

The financial functions were almost entirely **fine**. The largest divergence
class is **booleans**: LibreOffice treats them as plain numbers, so `TRUE=1` is
TRUE, `SUM` over a range containing TRUE includes it, `COUNT` counts it, and
`TRUE&""` is `"1"`. Excel gives booleans their own type, ranking above text and
above every number. Following the oracle here would have shipped
Excel-incompatible arithmetic into every model that uses a boolean anywhere.

The other declared classes: LibreOffice has no phantom 1900-02-29, so every date
before 1900-03-01 sits one day off Excel's serials; `DATE(26,1,1)` is 1926 by
Microsoft's documented "add 1900" rule, not 2026; `SQRT(-1)` and `LN(0)` are
`#NUM!` in Excel and `#VALUE!` in LibreOffice; LibreOffice pads scientific
exponents to three digits.

`divergences.json` records each with its reason, and the gate is symmetric: a
declared divergence that starts **matching** also fails, because that means
either LibreOffice changed or we drifted onto its side of the argument.

The deeper lesson is about oracles generally: an oracle is a *second opinion*,
not ground truth. The value came from the disagreements, and every one of them
had to be adjudicated by hand against Excel's documented behaviour.

### 12.6 ? Corpus breadth is the binding constraint

Unchanged from the original, and still the weakest evidence here: 399 synthetic
probes plus four sample workbooks. The oracle is a transform of the input, so it
can only teach us about shapes we thought to write down. Widening it with **real
agent output** is the highest-value next step (§16).

---

## 13. Fallbacks and write-back

### 13.1 The ladder ✓ as designed

| Tier | Path | When | Status |
|---|---|---|---|
| 0 | Trust `<v>` | Cache present | ✓ |
| 1 | **This engine** | Everything it covers | ✓ built |
| 2 | An adopted wasm engine | Open-vocabulary files | not adopted (§0) |
| 3 | LibreOffice headless | Tier 1 reports unsupported cells; user opts in | available, not wired into the app |
| 4 | Honest blank + ⚠ | Everything else | ✓ |

Tier 3 is deliberately an *escalation*, not the default, which preserves the
no-egress property for everything else.

### 13.2 Write-back — not built, and correctly gated

Injecting computed values as `<v>` so downstream consumers see numbers is
designed and unbuilt. The details still hold: set `t="str"` for string results
and `t="n"` for numeric; **omit `calcChain.xml`** and strip a stale one; keep
`fullCalcOnLoad="1"` so Excel recomputes anyway; never write back a cell whose
provenance is `unsupported` or `circular`.

And the reason to be careful stands: our `<v>` is *our* computation. If it
disagrees with Excel we have written our disagreement into the file permanently.
Gate it on a green oracle score and stamp provenance somewhere in the file.

### 13.3 Generator-side ✓ still the right move

`CAPABILITY.md` is generated from the registry and can be published into the
generation prompt as an allowlist, so the renderer never meets a formula it
cannot compute *by construction*. This is the original §5.1 idea, intact — only
the floor is wider, and the list is derived rather than hand-written.

---

## 14. Phase status

| P | Deliverable | Status |
|---|---|---|
| **0** | Eval harness + raw `t` probe; retire the 0.25 heuristic | ✓ done |
| **1** | Vocabulary histogram across real agent output | ✗ **not done** — superseded by §2.1, but the underlying need (a real corpus) is now §16's top item |
| **2** | `formualizer` bake-off | ✗ not run — decided on other grounds (§0) |
| **3** | Evaluator | ✓ done, broad rather than closed: 197 functions |
| **4** | Overlay + provenance UI | ✓ done |
| **5** | Worker + incremental + live edit | ✗ not done (§9.2) |
| **6** | Audit diff | ✓ done, and extended with the structural detector (§11) |
| **7** | Write-back | ✗ not done (§13.2) |

---

## 15. Risks, and what became of them

| # | Risk | Outcome |
|---|---|---|
| **R1** | **Silent wrong numbers** | Contained. Fail-loud throughout; false confidence is a hard gate at zero across 399 probes. Not *eliminated* — the corpus is synthetic (R10) |
| R2 | Single-maintainer wasm deps | Avoided entirely — nothing adopted |
| R3 | Wasm bundle weight | Moot — the engine is dependency-free |
| **R4** | **Closed grammar frozen on n=1** | **Materialised.** Caught before shipping, but only by checking the generator path. See §2.1 |
| R5 | LO oracle ≠ Excel | Materialised, differently than predicted (§12.5). Managed by a declared-divergence registry with a symmetric gate |
| R6 | Scope creep into "we built a spreadsheet app" | Held. View + audit, no authoring: no formula bar, no formatting UI, no save-in-place |
| R7 | ExcelJS may not translate shared formulas | Dissolved — the raw reader translates them itself (§7.9) |
| R8 | Write-back propagates our arithmetic into files | Not incurred; write-back unbuilt |
| **R9** | **New: O(n²) graph memory** | Materialised as an OOM at 45k formulas. Fixed by not materialising the graph (§8) |
| **R10** | **New: the corpus is synthetic** | Open. The gates are real, but they measure agreement on probes we wrote |

---

## 16. Open questions

1. **? Widen the corpus with real agent output.** The single highest-value next
   step. Every correctness claim here is conditional on a corpus we invented.
   Phase 1 was cut for the wrong reason — the closed grammar it was meant to
   freeze turned out not to exist — but the need for real data did not go away
   with it.
2. **? Which surface consumes this.** Still open. The packaging hedges: the
   engine is framework-free, the renderer is a React component, the demo is
   separable. If it lands in the product's Excel preview slot, the worker
   question (§9.2) gets sharper.
3. **? Is view-only still the contract.** Unresolved, and now cheap to resolve:
   the primitives for what-if exist, so this is a product decision rather than an
   engineering one.
4. **✓ Excel-parity bar.** Answered by the build: LibreOffice is good enough as a
   *second opinion*, provided every disagreement is adjudicated by hand and
   recorded. It is not good enough as ground truth.
5. **✓ Wasm in the bundle.** Answered: none, and none needed.
6. **✓ Does the generator hand-write OOXML or use openpyxl.** openpyxl, free-hand
   via `python_execute` — which is what invalidated §2.1.

---

## Appendix A — the capability floor

Superseded by **`CAPABILITY.md`**, generated from the function registry by
`capability.test.ts` and asserted on every test run. A hand-maintained list of
what an engine supports drifts within a week and then actively misleads — which,
for a fail-loud engine, is the one thing the documentation must not do.

Summary at time of writing: **197 implemented, 26 refused by name**, plus two
conditional refusals (`SUBTOTAL(101–111)`, and approximate-match lookups over
unsorted data).

The original's Appendix A specified a `{SUM, IF}` EBNF grammar with everything
else rejected at parse. The *grammar* is still there in `parser.ts` — Excel's
full expression syntax, with unsupported shapes rejected at parse rather than at
eval, exactly as intended. What changed is only the function set it admits.

---

## Appendix B — evidence

**Uncached vs cached, same cell, same model** ✓:
```xml
nocache : <c r="B7" s="19"><f>B6/B4</f><v></v></c>
recalc'd: <c r="B7" s="19" t="n"><f aca="false">B6/B4</f><v>0.59004854368932</v></c>
```

**Legitimately-empty vs never-computed — the `t` attribute** ✓:
```xml
recalc'd: <c r="G9" t="str"><f>IF(G9=0,"",F9/G9-1)</f><v></v></c>   ← computed to ""
nocache : <c r="F4"><f>SUM(B4:E4)</f><v></v></c>                     ← never computed
```

**`calcPr`** ✓:
```xml
nocache (openpyxl): <calcPr calcId="124519" fullCalcOnLoad="1"/>
recalc'd (LO):      <calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.0001"/>
```

**Original vocabulary probe** ✓ *(accurate about the sample, wrong as a basis for
freezing a grammar — §2.1)*: 76 formulas, 5 uncached, 0 shared-only, 19
cross-sheet; functions `IF`×12, `SUM`×11.

**Oracle, after the build** ✓:
```
399 probes · 13 suites · coverage 100.0% · accuracy 100.0% · false confidence 0
29 declared LibreOffice-vs-Excel divergences (tools/oracle/divergences.json)
largest class: booleans — LO treats them as numbers, Excel does not
```

**Browser end-to-end** ✓ (`tools/verify/drive.mjs`, headless Chrome):
```
Financial model (recalculated)   76 of 76 formulas · 0 unsupported · 1 ms
Same model, no cached values     75 of 75 formulas · 0 unsupported · 1 ms
all 232 rendered cells match between the cached and the computed file
the hardcoded total is flagged: Revenue!E6 typed in as 761.2; SUM(E4:E5) is 721.2
console errors: 0
```

**Performance by shape** ✓ (`perf.test.ts`):
```
local arithmetic    25000 formulas   ~100 ms
running total        5000 formulas   ~1.1 s    (12.5M cell reads, quadratic by construction)
full-table VLOOKUP   2000 formulas   ~0.39 s   (4M comparisons)
5000-deep chain      5000 formulas   ~10 ms    (guards against call-stack recursion)
```

**The graph failure that changed §8** ✓: with one edge materialised per precedent
cell, a 45,000-formula workbook containing a running-total column exhausted
memory and killed the test worker. Without materialising it: completes.

**npm registry, 2026-07-25** ✓: `formualizer@0.7.1` MIT OR Apache-2.0 ·
`@ironcalc/wasm@0.7.0` MIT/Apache-2.0 · `hyperformula@3.3.0` GPL-3.0-only ·
`@formulajs/formulajs@4.6.0` MIT · `fast-formula-parser@1.0.19` MIT (last publish
≈6 y ago).

---

## Appendix C — sources

- HyperFormula licensing — https://hyperformula.handsontable.com/docs/guide/licensing.html
- Formualizer — https://github.com/psu3d0/formualizer · https://www.formualizer.dev/
- IronCalc — https://github.com/ironcalc/IronCalc · https://nlnet.nl/project/IronCalc/
- Formula.js — https://www.npmjs.com/package/@formulajs/formulajs
- fast-formula-parser — https://github.com/LesterLyu/fast-formula-parser
- Predecessor spike: `/Users/vz/OSS excel render tests` — `SPIKE_FINDINGS.md`,
  `LEARNINGS.md`, and the `src/excel/` renderer this project's view descends from
- Generator path that invalidated §2.1:
  `financial-analyst-agent/apps/api/src/tools/excel-export.ts:70`
