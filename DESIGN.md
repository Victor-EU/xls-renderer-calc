# Recalculate & render — a calculation layer for the client-side Excel renderer

**Date:** 2026-07-25
**Subject repo:** `/Users/vz/OSS excel render tests` (ExcelJS + numfmt, client-side, view-only, MIT)
**Status: BUILT** — see `README.md` for what shipped and what did not.

> **Build outcome, 2026-07-25.** The architecture below was implemented as
> specified: overlay-not-mutation, fail-loud provenance, the raw-`t` probe
> replacing the 0.25 heuristic, the LibreOffice oracle, the computed-vs-cached
> audit. Phases 0–4 and 6 are done; the worker (P5) and write-back (P7) are not.
>
> **Three things the build contradicted, and why:**
>
> 1. **§2.1 / Appendix A: the closed `{SUM, IF}` grammar is wrong.** The
>    histogram was n=1. The actual generator path in the product is
>    `python_execute` + openpyxl written free-hand by a model
>    (`apps/api/src/tools/excel-export.ts:70`), so the vocabulary is open by
>    construction and a two-function engine would render ⚠ on the first real
>    model. Built broad (~180 functions) and gated by the oracle instead of
>    frozen against one sample. The fail-loud architecture is unchanged — it is
>    what makes a broad library safe.
>
> 2. **§8.2 "range dependencies: not a v1 concern" is wrong.** Materialising one
>    edge per precedent cell is O(rows²); a 5,000-row model with a running total
>    is 12.5 M edges and ran out of memory at 45,000 formulas. The graph is not
>    materialised at all now — precedent *rectangles* are walked lazily by an
>    iterative DFS whose post-order is the topological order.
>
> 3. **§8.3 "prefer explicit topo sort" over lazy recursion — for the right
>    reason, but the reason applies to the DFS too.** The search is iterative
>    because a row-to-row chain is as deep as the model is long. That is now a
>    test (`perf.test.ts`, 5,000-deep chain).
>
> **§12.5 was right and underrated.** LibreOffice-is-not-Excel produced 29
> declared divergences, and the biggest class was not financial functions but
> *booleans*: LibreOffice treats them as plain numbers, so `TRUE=1`, `SUM` over a
> range containing TRUE, and `COUNT` all differ. Following the oracle there would
> have shipped Excel-incompatible arithmetic.

Marker convention used throughout:
**✓** = verified today against real files/registries · **⚠** = vendor claim, not yet verified · **?** = open question.

---

## 0. Decision up front

Build the calc layer **from scratch for our own generated files**, and **adopt an
OSS wasm engine for everything else**, both behind one swappable interface.

The reason is measurement, not preference. The three-sheet sample model contains
76 formulas whose *entire* function vocabulary is **`SUM` and `IF`** plus bare
arithmetic (§2). A from-scratch evaluator covering that is small, zero-dependency,
provably fail-loud, and adds nothing to the bundle. Meanwhile an arbitrary
user-dropped `.xlsx` has an unbounded tail that no hand-written evaluator should
pretend to cover — and there is now a genuinely permissive engine for it
(`formualizer`, MIT/Apache-2.0, Rust→wasm ✓).

What must **not** happen: a partially-correct evaluator rendering a
plausible-but-wrong number into a financial model. That is strictly worse than
today's blank cell. Every design choice below is subordinate to that.

---

## 1. The problem, mechanically

### 1.1 What an agent writes

A cell in a worksheet part carries a formula and, separately, a **cached result**:

```xml
<c r="B7" s="19" t="n"><f>B6/B4</f><v>0.59004854368932</v></c>
```

`<f>` is the formula. `<v>` is the value Excel computed last time it saved.
Every consumer that isn't a full spreadsheet application reads `<v>`.

Agent-generated files (openpyxl, or hand-written OOXML) emit the formula with an
**empty** cache: ✓ verified, from `public/financial-model-nocache.xlsx`:

```xml
<c r="B7" s="19"><f>B6/B4</f><v></v></c>
```

Excel and LibreOffice don't care — they compute on open. Our renderer, and every
other cache-reading consumer, shows blank. In a financial model nearly every
meaningful number is a formula, so the document renders as a labelled skeleton.

### 1.2 The current fix, and why it's not enough

Today: `soffice --headless --convert-to xlsx` recalculates and caches every
formula. It works and it should stay (§13, tier 3). But it costs the project its
three defining properties:

| Property | LibreOffice recalc | Browser recalc |
|---|---|---|
| No server | ✗ needs a binary + process | ✓ |
| Nothing leaves the browser | ✗ file must be uploaded | ✓ |
| Real-time | ✗ ~1–3 s per call, per file | ✓ ms, per keystroke |

The third is the one that changes the product rather than just fixing a bug (§9).

### 1.3 What "real time" has to mean

Three distinct requirements hide behind the phrase. Design for all three; they
share one dependency graph.

1. **On load** — compute everything, once, before first paint. Fixes the blanks.
2. **On edit** — change an assumption, the model reflows. Turns a viewer into a
   what-if tool. *This is the one with product value.*
3. **On stream** — the agent is still writing the sheet; cells arrive as tool
   calls and the render recomputes per patch. "Watch the model get built."

---

## 2. What we verified today (the empirical anchor)

Probes against the two sample files in `public/`. **All ✓.**

### 2.1 Formula vocabulary — startlingly small

| | `financial-model.xlsx` (recalc'd) | `-nocache.xlsx` |
|---|---|---|
| Formula cells | 76 | 75 |
| Uncached | **5** (6.6% — legitimately empty) | **75** (100%) |
| Shared-formula-only cells | 0 | 0 |
| Cross-sheet refs | 19 cells | 19 cells |
| **Functions used** | **`IF`×12, `SUM`×11** | **`IF`×12, `SUM`×11** |
| Ref shapes | arithmetic ×57, cross-sheet ×19, string literal ×12, range `A1:B2` ×11, absolute `$` ×11 | identical |

**Two functions.** Everything else is `+ - * /`, cross-sheet references,
ranges, and absolute refs. This is the single most important number in the
document: it says the closed-vocabulary evaluator (§5, §7) is a few hundred lines
that reaches **100% coverage** on the corpus we actually care about — not an
aspiration.

Caveat before over-generalising: n=1 model, written by one generator
(`gen_model.py`). Phase 0 must widen this histogram across real agent output
before the closed set is frozen (§14, §15 R4).

### 2.2 The per-cell ambiguity is resolvable — `parse.ts` can stop guessing

`src/excel/parse.ts` carries a documented limitation:

> A formula that legitimately computes to an empty string is indistinguishable,
> per-cell, from one that was never computed (ExcelJS surfaces both as `{formula}`
> with no result). So we judge at the file level […] `RECALC_THRESHOLD = 0.25`.

That is true *of ExcelJS*, but **not true of the file**. ✓ The raw XML
distinguishes them by the `t` (cell type) attribute:

| Case | Raw XML | `t` |
|---|---|---|
| Computed to empty string | `<c r="G9" t="str"><f>IF(G9=0,"",F9/G9-1)</f><v></v></c>` | **`str`** |
| Never computed | `<c r="F4"><f>SUM(B4:E4)</f><v></v></c>` | **absent** |

A formula cell with an empty `<v>` and **no `t`** was never computed. With
`t="str"` it computed to `""`. ExcelJS drops `t` on the floor, which is why the
0.25 heuristic exists.

Two consequences:
- A ~30-line raw-XML probe (unzip `xl/worksheets/*.xml`, regex the `<c>` tags for
  `r`, `t`, presence of `<f>`, emptiness of `<v>`) replaces the heuristic with
  **exact per-cell truth** — independently of the whole calc layer. Cheap,
  standalone, ships first.
- Once we can *evaluate*, the question dissolves anyway: compute the cell; if the
  answer is `""`, it's legitimately empty. The heuristic is retired twice over.

### 2.3 `fullCalcOnLoad` is already set — by openpyxl, and LibreOffice strips it

I expected this to be a missing one-line win. It isn't, and the direction is the
opposite of intuition. ✓

| File | `<calcPr>` |
|---|---|
| `-nocache.xlsx` (openpyxl) | `<calcPr calcId="124519" fullCalcOnLoad="1"/>` |
| `financial-model.xlsx` (after LO) | `<calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.0001"/>` |

openpyxl already sets `fullCalcOnLoad="1"`, so a user who *downloads* the
un-recalced file and opens it in Excel sees correct numbers. The blanks are
**our renderer's problem specifically**, not a defect in the file.

So the recommendation narrows: `fullCalcOnLoad` is only a real gap for a
**hand-written OOXML** path that omits `<calcPr>` — worth a one-line assertion in
whatever writer the agent uses, and worth re-adding after a LibreOffice round-trip
(LO drops it, which is harmless only because LO also cached every value).

Incidentally, LO's `calcPr` hands us Excel's iterative-calculation defaults for
free: `iterateCount=100`, `iterateDelta=0.0001`, `iterate=false`. Use exactly
those when implementing circular-reference policy (§7.10).

### 2.4 Precision, as actually written

LO wrote `0.59004854368932` for `B6/B4` — 14 significant digits. The IEEE-754
double is `0.5900485436893203…`. Excel stores full precision and *displays* 15
significant digits. This sets the oracle-diff tolerance policy (§12.3): compare
floats with a **relative** epsilon around `1e-9`, never `===`.

### 2.5 No `calcChain.xml` in either file

✓ Neither sample has one. `calcChain.xml` is an optional Excel optimisation
recording evaluation order; a stale one causes Excel to complain. Our write-back
path (§13.2) should therefore **not** attempt to author it — just omit it (and
delete it, plus its content-type override and relationship, if a source file has
one). Excel rebuilds it.

### 2.6 Live baseline of the existing renderer

Measured in Chrome today: parse 12.4 KB / 3 sheets / 219 cells / 76 formulas in
**74 ms**; nocache file **75/75 un-cached**, trap banner fires; zero console
errors. Whatever recalc costs, it is measured against a 74 ms parse.

---

## 3. Failure modes, ranked

Design against these in order. The ranking is what makes this a *finance*
renderer rather than a generic one.

| # | Failure | Severity | Mitigation |
|---|---|---|---|
| 1 | **Renders a wrong number confidently** (bad coercion, unimplemented function silently treated as 0, wrong lookup semantics) | **Catastrophic** — invisible, and it's a financial figure | Fail-loud doctrine §10. Unsupported → `#UNSUPPORTED`, never a guess. Closed grammar rejects at *parse* time, not eval time |
| 2 | Renders blank (today) | Bad | The whole point of this work |
| 3 | Renders `#UNSUPPORTED` where Excel has a value | Acceptable | Honest. Counted, surfaced, drives the coverage backlog |
| 4 | Slow / blocks the UI | Annoying | Worker + incremental §9 |
| 5 | Bundle bloat for files that don't need calc | Annoying | Lazy `import()` gated on detection §8.4 |

The asymmetry between #1 and #3 is the entire ethical content of this design.
An engine that says "I don't know" is a tool. An engine that guesses is a
liability, and in a model the guess is unfalsifiable by eye.

---

## 4. Engine landscape (verified 2026-07-25)

`npm view` metadata is ✓; capability claims are ⚠ from READMEs.

| Package | License ✓ | Version ✓ | Maintainers ✓ | Capabilities ⚠ |
|---|---|---|---|---|
| **`formualizer`** | **MIT OR Apache-2.0** | 0.7.1 | 1 (`psu3d0`) | Rust→wasm. `Workbook.fromXlsxBytes()`, `setValue/setFormula`, `evaluateCell/evaluateAll`, `readRange`, `tokenize/parse`. 400+ fns; Arrow-backed storage; incremental dep tracking + cycle detection; topological scheduling; dynamic arrays w/ spill (`FILTER`/`UNIQUE`/`SORT`/`XLOOKUP`); xlsx r/w via calamine+umya. Self-described "production-grade", v0.6+; "span evaluation experimental, opt-in". Pitched at "high-trust agent workflows" |
| **`@ironcalc/wasm`** | MIT/Apache-2.0 | 0.7.0 | 1 (`n.hatcher`) | Rust→wasm, xlsx reader+writer, `model.evaluate()`. README self-describes as **early-stage WIP**. EC-grant funded; IronCalc GmbH (Berlin, early 2026) |
| `hyperformula` | **GPL-3.0-only** (or paid proprietary) | 3.3.0 | 5 (Handsontable) | The mature option: 400+ fns, CRUD, undo/redo, incremental recalc, named expressions |
| `@formulajs/formulajs` | MIT | 4.6.0 | 1 | ~500 Excel *functions* as plain JS. **No parser, no cell refs, no dep graph.** An ingredient |
| `fast-formula-parser` | MIT | 1.0.19 | 1 | Parser + evaluator w/ ref callbacks, ~300 fns. **Last publish ≈6 years ago** |

### Reading of the table

- **HyperFormula is disqualified by license,** not by quality. GPL-3.0-only
  means either the consuming product goes open source or a commercial licence is
  purchased. Worth knowing the paid door exists; not the default path.
- **`formualizer` is the lead adoption candidate.** It is the only permissive
  option that ships parser + evaluator + dependency graph + xlsx ingest + wasm in
  one package, and its stated design goals (determinism, agent workflows) line up
  with this use case unusually well. Its risks are maturity (v0.7.x) and
  bus factor (1) — mitigated by MIT + Rust: it can be vendored and forked.
- **`@formulajs/formulajs` remains useful even in the from-scratch path** as a
  vetted, MIT source of *individual function* implementations (`NPV`, `IRR`,
  `PMT`, `EOMONTH`…). Borrow the function bodies; keep our own parser, graph, and
  coercion rules — which is where correctness actually lives.
- `fast-formula-parser` is stale. Its architecture (evaluate with `onCell`/
  `onRange` callbacks) is nonetheless the right shape and worth reading for design.

Both wasm candidates need a **bundle-size measurement** before adoption — the
current app is 360 KB gzip total, and a Rust spreadsheet engine could rival that.
Unmeasured ⚠; §12 answers it.

---

## 5. The reframe: two problems wearing one hat

"Recalculate xlsx" is two problems with opposite correct answers. Conflating them
is what makes the build-vs-adopt question feel hard.

|  | **Problem A — our own agent's output** | **Problem B — arbitrary user `.xlsx`** |
|---|---|---|
| Formula vocabulary | **Closed and declarable.** Measured: `SUM`, `IF`, arithmetic (§2.1) | Open, unbounded tail |
| Who controls it | We do — we write the generator prompt | Nobody |
| Right answer | **From scratch.** ~600–900 LOC, zero deps, no bundle cost, provably fail-loud because anything outside the grammar is rejected at parse time | **Adopt an engine.** The tail isn't functions, it's coercion semantics (§7) — hand-rolling that is how you get failure mode #1 |
| Coverage target | **100%** of the declared grammar, gate-enforced | Best-effort + honest gaps + LO fallback |
| Volume | The common case, every render | Occasional (drag-and-drop) |

### 5.1 The closed loop this unlocks

Because we control the generator, the evaluator's supported set can be
**published into the generation prompt as a formula allowlist**. The renderer
then never meets a formula it can't compute, *by construction*, and the guarantee
is enforceable in CI: parse every generated model, assert every formula parses
under the closed grammar.

This inverts the usual dependency — instead of chasing Excel's surface area, we
declare a floor and hold the generator to it. Widening the floor is then a
deliberate, tested act (add function → implement → oracle-diff → add to prompt),
not a bug report from a user staring at a blank cell.

*Guarantee the floor; don't cap the ceiling.* Problem B keeps the ceiling open via
the adopted engine and the LO fallback — a user's hand-built model with
`XLOOKUP` and dynamic arrays still renders, just through a different tier.

---

## 6. Architecture

### 6.1 Overlay, not mutation

```
        ┌───────────────────────────────────────────────────────────┐
buf ──► │ parse.ts        ExcelJS workbook (unchanged, source of    │
        │                 truth for style/format/merge/layout)      │
        └──────────┬────────────────────────────────────────────────┘
                   │ formulas + cached values + raw `t` probe (§2.2)
                   ▼
        ┌───────────────────────────────────────────────────────────┐
        │ recalc.ts       RecalcEngine (swappable)                  │
        │                   ├─ ClosedGrammarEngine  (ours, §7)      │
        │                   ├─ FormualizerEngine    (wasm, §4)      │
        │                   └─ NullEngine           (cached-only)   │
        └──────────┬────────────────────────────────────────────────┘
                   │ ValueOverlay
                   ▼
        ┌───────────────────────────────────────────────────────────┐
        │ ExcelView.tsx   renders overlay[addr] ?? cachedValue,     │
        │                 badges by provenance, ⚠ on unsupported    │
        └───────────────────────────────────────────────────────────┘
```

**Never write computed values back into the ExcelJS workbook.** The overlay is a
side table. Three things depend on that choice:

1. **Provenance.** Every rendered number knows where it came from. In a financial
   preview that is not a nicety; it's the difference between a view and a claim.
2. **The diff.** Cells with *both* a cached and a computed value can be compared
   (§11) — which is a genuine audit feature, and impossible if you overwrite.
3. **Swappability.** Engines are interchangeable and A/B-comparable against each
   other and against the oracle if none of them mutate shared state. Same
   discipline as keeping the PDF parser swappable.

### 6.2 Types

```ts
type Addr = string;                     // "Income Statement!F4" — interned to int internally

type CellValue =
  | { t: 'num';   v: number }
  | { t: 'str';   v: string }
  | { t: 'bool';  v: boolean }
  | { t: 'date';  v: Date }             // serial → Date at the boundary, not inside eval
  | { t: 'err';   v: ExcelError };      // '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?'
                                        // | '#NUM!' | '#N/A' | '#NULL!' | '#SPILL!' | '#CALC!'

type Provenance =
  | 'cached'        // came from <v> in the file — we did not touch it
  | 'computed'      // we evaluated it
  | 'unsupported'   // formula outside engine capability — rendered as ⚠, never guessed
  | 'circular'      // participates in a reference cycle
  | 'volatile';     // computed, but depends on NOW/TODAY/RAND — frozen at load

type OverlayCell = {
  value: CellValue | null;
  src: Provenance;
  cached?: CellValue;     // retained when it existed and differs → feeds §11
  reason?: string;        // why unsupported — the coverage backlog, verbatim
  ms?: number;            // per-cell eval cost, for profiling hot spots
};

type ValueOverlay = {
  cells: Map<Addr, OverlayCell>;
  stats: {
    formulas: number; computed: number; cached: number;
    unsupported: number; circular: number; mismatched: number;
    engine: string; totalMs: number;
  };
  gaps: Array<{ fn?: string; shape?: string; count: number; sample: Addr }>;
};

interface RecalcEngine {
  readonly name: string;
  readonly capabilities: { functions: Set<string>; features: Set<Feature> };
  load(wb: ExcelJS.Workbook, raw?: RawProbe): Promise<void>;
  evaluateAll(): Promise<ValueOverlay>;
  setValue(addr: Addr, v: CellValue): Promise<CellPatch[]>;   // incremental — §9.2
  setFormula(addr: Addr, f: string): Promise<CellPatch[]>;    // streaming — §9.3
  trace(addr: Addr): { precedents: Addr[]; dependents: Addr[] };  // §11.2
}
```

`gaps` is deliberately part of the contract. An engine must be able to say what
it couldn't do, aggregated, or the fail-loud doctrine has no reporting surface.

### 6.3 Module layout in the subject repo

```
src/excel/
  parse.ts          (exists) + raw `t` probe → exact per-cell uncached truth §2.2
  format.ts         (exists) — isUncachedFormula() becomes overlay-aware
  theme.ts          (exists) — untouched
  ExcelView.tsx     (exists) + provenance badges, diff mode, trace highlight
  recalc/
    index.ts        RecalcEngine interface + engine selection/fallback ladder
    overlay.ts      ValueOverlay construction, stats, gap aggregation
    graph.ts        dependency graph, topo order, SCC cycle detection  §8
    worker.ts       off-main-thread host + message protocol            §9.1
    closed/         ── ours, Problem A ──
      grammar.ts    declared closed grammar (Appendix A) — the contract
      lex.ts        tokenizer
      parse.ts      Pratt parser → AST, rejects anything off-grammar
      refs.ts       A1 parsing, ranges, cross-sheet, absolute, defined names
      coerce.ts     Excel type-coercion rules                          §7.2
      fns/          SUM, IF, … one file per group, each with oracle tests
      eval.ts       AST walk over the graph
    formualizer/    ── adopted, Problem B ── thin adapter to the wasm engine
```

---

## 7. Evaluation semantics — the part that decides correctness

Function *count* is a vanity metric. Excel compatibility lives almost entirely in
coercion, precision, and error propagation. This is the catalogue a from-scratch
evaluator must consciously accept or reject; each item is either implemented and
oracle-tested, or **outside the declared grammar and rejected at parse**.

### 7.1 Operator precedence, including Excel's oddities

Highest → lowest: `:` (range) → ` ` (intersection) → `,` (union) → unary `-` →
`%` (postfix) → `^` → `*` `/` → `+` `-` → `&` → comparisons (`=` `<>` `<` `>` `<=` `>=`).

Two traps worth writing tests for immediately:
- **`-2^2 = 4`** in Excel. Unary minus binds *tighter* than `^`, unlike ordinary
  mathematical convention and unlike most programming languages.
- **`2^3^2 = 64`** — `^` is **left**-associative in Excel (`(2^3)^2`), not right.

Intersection (space) and union (comma) operators: out of the closed grammar,
rejected at parse. They're vanishingly rare in generated models and easy to
mis-parse into something plausible.

### 7.2 Type coercion — the highest-risk area

| Situation | Excel result | Note |
|---|---|---|
| `"5"+1` | `6` | Arithmetic **does** coerce numeric text |
| `"5"=5` | `FALSE` | Comparison does **not**. Different rules per context |
| `TRUE+1` | `2` | Booleans coerce in arithmetic |
| `TRUE=1` | `FALSE` | …but not in comparison |
| Sort/compare order | number < text < boolean | Cross-type comparison is by type rank |
| `"a"+1` | `#VALUE!` | Non-numeric text in arithmetic errors |
| `""+1` | `#VALUE!` | Empty *string* ≠ empty *cell* |

### 7.3 Empty-cell semantics — and why this model already depends on it

An empty cell is **not** an empty string. In arithmetic it is `0`; in comparison
it equals **both** `0` and `""`.

This is not academic. The sample model's own formula is:

```
IF(G9=0,"",F9/G9-1)      ← G9 is empty (no prior-year column)
```

`G9=0` → **TRUE** for an empty G9 → returns `""` → LO writes `t="str"` with an
empty `<v>` (✓ §2.2). An evaluator lacking the empty-cell rule computes `G9=0`
as FALSE, evaluates `F9/G9`, and renders **`#DIV/0!`** where Excel renders a
clean blank. One rule, twelve wrong cells in a 76-formula model.

This is the concrete case for "semantics over function count", and it's the
first unit test to write.

### 7.4 Numbers and precision

- All arithmetic is IEEE-754 double.
- Excel *displays* 15 significant digits; it stores full precision. Rendering
  goes through `numfmt` already, so display is handled — but the **oracle diff**
  must use a relative epsilon (§12.3), never `===`. ✓ §2.4 shows LO writing 14
  digits for a value whose double has 17.
- Excel applies a small "cosmetic rounding" when a subtraction result is very
  near zero (the classic `0.1+0.2-0.3` case). Decide explicitly: **do not**
  emulate it in v1; record it as a known divergence with a test that documents
  the difference rather than asserting Excel's behaviour.

### 7.5 Dates

- Dates are serial numbers; times are the fractional part.
- **The 1900 leap-year bug:** serial 60 is a phantom 1900-02-29. Any serial→date
  conversion must reproduce it to stay aligned with Excel.
- **The 1904 date system** exists (`<workbookPr date1904="1"/>`, legacy Mac).
  Read the flag; if set and we don't support it, that's a **file-level**
  `unsupported`, not a silent 4-year error.
- Keep serials internally; convert to `Date` only at the render boundary.
  `numfmt` already takes a JS `Date` directly for codes like `mmm-yyyy`.

### 7.6 Errors and propagation

Nine error values (`#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`,
`#NULL!`, `#SPILL!`, `#CALC!`). Rules: errors propagate through operators and
most functions; `IFERROR`/`IFNA`/`ISERROR`/`ISNA` trap them; `#N/A` has distinct
handling from the rest. Critically — **an Excel error is a legitimate computed
value, not an engine failure.** `#DIV/0!` from a genuine divide-by-zero renders as
`#DIV/0!` with provenance `computed`. Only *our* inability to evaluate produces
`unsupported`. Conflating these two would make the honesty layer meaningless.

### 7.7 References

| Shape | Example | Closed grammar? |
|---|---|---|
| Relative / absolute | `B4`, `$B$4`, `B$4` | ✓ (✓ present: 11 cells) |
| Range | `B4:E4` | ✓ (✓ present: 11 cells) |
| Cross-sheet, quoted names | `'Income Statement'!F4` | ✓ (✓ present: 19 cells — **quoting is mandatory**, the sheet name has a space) |
| Whole column / row | `A:A`, `1:1` | v2 — needs a used-range bound to avoid 1M-cell expansion |
| Defined names | `TaxRate` | v2 — parse `xl/workbook.xml` `<definedNames>` |
| 3D refs | `Sheet1:Sheet3!A1` | ✗ reject |
| Structured table refs | `Table1[Revenue]` | ✗ reject (v3 — needs table parts) |
| `INDIRECT` / `OFFSET` | `INDIRECT("B"&n)` | ✗ **reject — breaks the static dependency graph** (§8.4) |

### 7.8 Function-level semantics that bite

Even inside a small function set:
- **`SUM` ignores text and booleans** in ranges but errors on a direct text
  argument. Range vs scalar argument handling differs.
- **`IF` is lazily evaluated** — the untaken branch must not be evaluated, or
  `IF(A1=0,"",1/A1)` throws on the very case it's guarding.
- **Criteria strings** (`SUMIF`/`COUNTIF`: `">100"`, `"<>x"`, wildcards `*` `?`
  `~`) are a miniature expression language. When those functions are added,
  they're a *parser* task, not a one-liner.
- **`VLOOKUP` approximate match** (4th arg omitted/TRUE) requires sorted data and
  returns the largest value ≤ lookup — the single most misimplemented Excel
  behaviour. `MATCH` types `1`/`0`/`-1` likewise. Ship these with `FALSE`/exact
  only in v1 and reject approximate match rather than guessing.

### 7.9 Shared and array formulas

- **Shared formulas** — `<f t="shared" ref="B2:E2" si="0"/>` stores the formula
  once; sibling cells reference `si` and must have refs **translated** by their
  row/column offset. ✓ Neither sample uses them (`sharedOnly=0`), but Excel
  itself emits them constantly, so any file a user drops in will have them.
  ExcelJS surfaces `sharedFormula`; verify whether it translates refs or hands
  back the master formula verbatim — **?** open, must be tested before trusting.
- **Array formulas** (`t="array"`) and modern **dynamic arrays** with spill
  ranges, `_xlfn.` prefixes, and `#SPILL!` — reject in v1, whole-file flag.

### 7.10 Cycles

Detect strongly-connected components (Tarjan) on the dependency graph. Default
Excel behaviour is to refuse and warn; iterative calc is opt-in with
`iterateCount=100`, `iterateDelta=0.0001` — ✓ exactly the values LO wrote into
our own sample (§2.3). v1: detect, mark provenance `circular`, render `⚠`, count
at file level. Do not silently iterate.

### 7.11 Volatile functions

`NOW`, `TODAY`, `RAND`, `RANDBETWEEN`, `OFFSET`, `INDIRECT`, `CELL`, `INFO`.
Two problems: they make our output differ from the file's cache for legitimate
reasons (so §11's diff must exclude their descendants), and `OFFSET`/`INDIRECT`
make dependencies dynamic, which a static graph cannot express. Policy: freeze
`NOW`/`TODAY` at load time (one clock read, passed in — never read the clock
mid-evaluation, so runs are reproducible); mark cells and their dependents
`volatile`; reject `OFFSET`/`INDIRECT` entirely in the closed grammar.

---

## 8. Dependency graph

### 8.1 Representation

- Intern every address to an `i32` (`sheetIdx << 40 | row << 20 | col`). Maps
  keyed by string are the obvious first-draft perf mistake.
- **Forward edges** precedent → dependents (drives dirty propagation on edit).
- **Reverse edges** dependent → precedents (drives evaluation and `trace()`).
- Extract refs during parse — the AST already walks them; no second pass.

### 8.2 Range dependencies

A formula reading `B4:E4` depends on 4 cells; naively expanding is fine here but
`SUM(A:A)` is 1,048,576 edges. Bound whole-column refs to the sheet's used range,
and if range-heavy models appear later, add block-bucketed range nodes (a range
node depends on the 64×64 blocks it overlaps). **Not** a v1 concern — note it and
move on. ✓ Current corpus: 11 small ranges.

### 8.3 Evaluation order

Kahn topological sort over reverse edges → evaluate once, in order, no recursion,
no stack-depth risk. Tarjan SCC first to peel out cycles (§7.10). Alternative
(lazy recursive memoised eval with in-progress marking) is simpler to write but
risks deep recursion on long dependency chains — a 200-row model with
row-to-row links is a 200-deep chain, and cross-sheet KPI cards make it worse.
**Prefer explicit topo sort.**

### 8.4 The dynamic-dependency escape hatch

`INDIRECT`/`OFFSET` create edges only knowable at evaluation time. Any engine
claiming to support them needs either re-graphing after evaluation or
conservative over-approximation. We reject them (§7.7) — which is exactly why
rejecting at *parse* time matters: the graph stays a static, provable object.

---

## 9. Real time

### 9.1 Off the main thread

Parse and evaluation move into a Web Worker. The overlay comes back as a
structured-cloneable payload; if it ever gets big, flatten to parallel typed
arrays (`Int32Array` addrs + `Float64Array` values + tag array) rather than a
Map of objects.

```
main ──► { type:'load', buf: ArrayBuffer }              (transferable)
     ◄── { type:'overlay', cells, stats, gaps }
main ──► { type:'setValue',   addr, value }
     ◄── { type:'patch', cells: CellPatch[] }           (only dirty cells)
main ──► { type:'setFormula', addr, formula }           (streaming §9.3)
     ◄── { type:'patch', … }
main ──► { type:'trace', addr }
     ◄── { type:'trace', precedents, dependents }
```

Note the existing parse already needs `buf.slice(0)` because ExcelJS detaches the
buffer — with a worker, transfer semantics make that explicit rather than
incidental. (There's a scar from the PDF rasterizer on exactly this class of bug:
buffer detachment that unit tests don't catch.)

### 9.2 Incremental recalculation

On `setValue(addr)`: BFS the forward edges to collect transitive dependents →
that's the dirty set → re-evaluate in restricted topological order → return only
changed cells. Add volatile nodes to every dirty set unconditionally.

React should patch **only** touched cells. Re-rendering the whole `<table>` on
each keystroke will dominate the cost and make a 2 ms recalc feel like 200 ms.

### 9.3 Streaming agent output

If the generating agent emits cells incrementally, each arrival is
`setFormula`/`setValue` + a dirty-set recalc. The renderer becomes a live view of
model construction. Free, once §9.2 exists.

### 9.4 Budgets (targets, to be validated)

| Operation | Target | Reference |
|---|---|---|
| Parse (existing) | ~74 ms ✓ measured, 219 cells | today |
| Full evaluate, ≤1k formulas | < 30 ms | ~13× current formula count |
| Full evaluate, ≤10k formulas | < 300 ms | worker, so non-blocking anyway |
| Single-cell edit ripple | < 5 ms | dirty set is typically tens of cells |
| First paint regression | 0 ms | lazy-load, cached fast path |

### 9.5 Lazy loading

The engine module loads via dynamic `import()` **only** when the parse probe
finds uncached formulas. Files with a complete cache — the common case for
LO-recalc'd or Excel-saved files — pay nothing. This is the mitigation for a wasm
engine's bundle weight, and it's why the detection step (§2.2) is worth having
independently of the engine choice.

---

## 10. The honesty layer

### 10.1 Doctrine

> **A cell we cannot compute renders `⚠`, never a number.**
> Unsupported function, unsupported ref shape, uncertain coercion, cycle,
> unsupported date system, unresolved external link → `unsupported` provenance,
> counted at file level, reason recorded verbatim.

Corollary from §7.6: genuine Excel errors are *computed values*, not failures.
`#DIV/0!` renders as `#DIV/0!`.

### 10.2 UX surface

The existing banner + per-cell `⚠` chassis is exactly right and changes meaning
rather than shape — from *diagnosis* to *remedy*:

| Today | After |
|---|---|
| ⚠ **This file needs a recalc.** 75 formula cells have no cached result — they render blank | ✓ **Computed live in your browser.** 75/75 formulas · 0 unsupported · 4 ms |
| — | ⚠ **73/75 computed.** 2 cells use `XLOOKUP`, which this engine doesn't support — shown as ⚠ |

Additions:
- Status chips: `computed 75` · `cached 0` · `unsupported 0` · `engine closed-v1`.
- Cell hover: formula, computed value, cached value if present and different,
  provenance, precedent list.
- A three-way toggle **computed / cached / diff** (§11).
- Click a cell → highlight precedents (`trace()`), the audit gesture people
  actually want in a model.

---

## 11. The sleeper feature: recalc as audit

### 11.1 Computed-vs-cached diff

Once both values exist for a cell, disagreement is a **finding**:

- The agent **hardcoded** a total instead of summing it (`<v>1827.6</v>` with no
  `<f>` where siblings have formulas) — structural, detectable without evaluating.
- The stated total **≠** the sum of its parts (cached `500`, computed `450`).
- Stale cache: file edited by a tool that didn't recalc.

For an agent-generated financial model this is arguably worth more than the
rendering fix. "This model's numbers are internally consistent" is a claim
nobody can currently make cheaply, and it's exactly the kind of check a reviewer
of an LLM-produced model wants. It falls out of the overlay for free.

Exclusions to avoid crying wolf: volatile descendants (§7.11), and floats within
the relative epsilon (§12.3).

### 11.2 Trace / explain

With a dependency graph, "explain this number" is a graph walk:
`Net income ← Pretax − Tax ← EBIT − Interest ← Gross profit − Opex ← …`.
Render it as a chain. This is the same capability Excel's Trace Precedents
gives, and it's the natural pairing with an agent that *wrote* the model:
the agent's rationale and the graph's actual structure can be compared.

---

## 12. Eval harness — build this first

We already own a ground-truth oracle and it was sitting in `public/` the whole
time: **the same model twice**, pre- and post-LibreOffice-recalc.

### 12.1 The loop

```
financial-model-nocache.xlsx ──► engine ──► ValueOverlay
                                                │
financial-model.xlsx (LO-recalc'd) ──► expected values
                                                │
                                          per-cell diff ──► score + buckets
```

### 12.2 Metrics

| Metric | Definition | v1 gate |
|---|---|---|
| **Coverage** | cells the engine *attempted* / formula cells | 100% (closed grammar) |
| **Accuracy** | cells matching the oracle / attempted | 100% |
| **Unsupported** | count + `{function, shape}` buckets | 0 on our corpus |
| **False-confidence** | computed ≠ oracle **and** not flagged | **0 — hard gate, non-negotiable** |
| Latency | full evaluate, ms | §9.4 |
| Bundle | added KB gzip | measured, lazy-loaded |

False-confidence is the metric that matters. The others describe capability; that
one describes trustworthiness, and failure mode #1 is precisely a nonzero value here.

### 12.3 Tolerance policy

- Numbers: `|a−b| <= 1e-9 * max(1,|a|,|b|)` — relative, never `===` (✓ §2.4).
- Strings, booleans, errors: exact.
- Empty: `""` (`t="str"`) and never-computed are **distinct** outcomes (✓ §2.2)
  and must not be conflated by the comparator — the comparator is itself the
  first place that bug would hide.

### 12.4 Widening the corpus

1. Parametrise `gen_model.py` to emit variants (more sheets, deeper chains,
   shared formulas, whole-column refs, dates, lookups, cycles).
2. Add **real** agent-generated workbooks — the vocabulary histogram of §2.1 is
   n=1 and must not be frozen on that basis.
3. Recalc each with LO → `(input, expected)` pairs, committed as fixtures.
4. `npm run eval` prints the markdown table; CI gates on the two 100%s and the
   zero.

### 12.5 The oracle's own limits

LibreOffice is not Excel. On financial functions, edge-case rounding, and newer
functions they diverge. Mitigations: keep a small **hand-verified-in-Excel**
fixture set for anything financial (`IRR`, `XIRR`, `NPV`, `PMT`), and record
known LO↔Excel divergences as documented expectations rather than silent
failures. Also worth noting: the oracle is a **transform of the input** — if
`gen_model.py` never emits shared formulas, LO's output can't teach us about them.
Corpus breadth is the binding constraint on the oracle's value, not tolerance.

---

## 13. Fallbacks and write-back

### 13.1 The ladder

| Tier | Path | When | Cost |
|---|---|---|---|
| 0 | Trust `<v>` | Cache present (Excel/LO-saved files) | 0 |
| 1 | **Closed-grammar engine (ours)** | Our agent's output — every formula in-grammar | ~ms, no bundle |
| 2 | **`formualizer` wasm** | Open-vocabulary files, if the bake-off clears | ~ms, lazy KB |
| 3 | LibreOffice headless recalc | Tier 1+2 report unsupported cells; user opts in | server + upload |
| 4 | Honest blank + ⚠ | Everything else | 0 |

Tier 3 is the existing, proven capability and the eval oracle. Keep it — but it's
now an explicit, user-consented escalation ("compute this on the server?"), not
the default path, which preserves the no-egress property for tiers 0–2.

### 13.2 Write-back: "Save recalculated .xlsx"

Inject computed values as `<v>` (with correct `t`), so downstream consumers see
numbers. Details that matter: set `t="str"` for string results and `t="n"` for
numeric (✓ the distinction §2.2 depends on); **omit `calcChain.xml`** and strip
it if present (✓ §2.5 — neither sample has one; a stale one makes Excel
complain); keep `fullCalcOnLoad="1"` (✓ §2.3) so Excel still recomputes from the
formulas and our cache is only a convenience for non-Excel readers; never write
back cells with provenance `unsupported` or `circular`.

Note the asymmetry this creates and accept it deliberately: our `<v>` is *our*
computation. If it disagrees with Excel, we've now written our disagreement into
the file. That's an argument for gating write-back on a green eval score, and for
stamping provenance somewhere in the file (a custom property or a doc-info note).

### 13.3 Generator-side

- Assert `<calcPr fullCalcOnLoad="1"/>` if the agent hand-writes OOXML (openpyxl
  already does it ✓ §2.3).
- Publish the closed grammar into the generation prompt as an allowlist (§5.1),
  and CI-assert every generated formula parses under it.
- Prefer formulas over hardcoded values in generation — with a working evaluator,
  formulas are now *strictly better* (they render, and they're auditable via
  §11), which removes the incentive to hardcode.

---

## 14. Phased plan

Each phase is independently shippable and independently useful.

| P | Deliverable | Acceptance | Est. |
|---|---|---|---|
| **0** | **Eval harness + raw `t` probe.** Oracle diff runner, fixture pairs, markdown report. Retire `RECALCD_THRESHOLD` in favour of exact per-cell truth (§2.2) | Harness scores an engine end-to-end; `NullEngine` baseline = 0% computed / 0 false-confidence. `parse.ts` reports exact uncached counts with no heuristic | ~½ day |
| **1** | **Vocabulary histogram** across real agent output (§12.4) | Function/shape histogram over ≥10 real workbooks → the closed grammar is frozen against data, not against n=1 | ~½ day |
| **2** | **`formualizer` bake-off** — `fromXlsxBytes` → `evaluateAll` → `readRange`, scored by P0 | A number: accuracy, coverage, latency, added KB gzip. Decision recorded either way | ~½ day |
| **3** | **Closed-grammar engine v1** — lexer, Pratt parser, refs, coercion (§7.2, §7.3), `SUM`/`IF`/arithmetic, graph + topo eval | 100% coverage & accuracy on the corpus, **0 false-confidence**, `-2^2=4` and empty-cell tests green | ~3–5 days |
| **4** | **Overlay + provenance UI** — badges, chips, hover, banner rewrite (§10.2) | Nocache sample renders identically to the recalc'd sample, labelled "computed live" | ~1 day |
| **5** | **Worker + incremental + live edit** (§9) | Edit an assumption → dependent cells update < 5 ms; UI never blocks | ~2 days |
| **6** | **Audit diff mode** (§11) | A deliberately hardcoded-total fixture is flagged | ~1 day |
| **7** | **Write-back** (§13.2), gated on green eval | Round-trip: our output opens in Excel with correct values and no repair prompt | ~1 day |

P0–P2 are ~1.5 days and **de-risk everything after**: they tell us whether P3 is
even necessary, and they're the regression gate if it is. Do not start P3 before
P2 reports a number.

---

## 15. Risks

| # | Risk | Mitigation |
|---|---|---|
| **R1** | **Silent wrong numbers** (failure mode #1) | Fail-loud doctrine §10; closed grammar rejects at parse; false-confidence is a hard CI gate §12.2 |
| R2 | Single-maintainer wasm deps (`formualizer`, `@ironcalc/wasm` ✓) | MIT + Rust → vendorable/forkable. Engine behind an interface; tiers 1/3 survive its abandonment |
| R3 | Wasm bundle weight vs 360 KB baseline | Lazy `import()` gated on detection §9.5; measured in P2 before adoption |
| R4 | Closed grammar frozen on n=1 (§2.1) | P1 histogram over real output before freezing |
| R5 | LO oracle ≠ Excel (§12.5) | Hand-verified Excel fixtures for financial fns; divergences documented, not silent |
| R6 | **Scope creep into "we built a spreadsheet app"** | Explicit line: **view + what-if, not authoring.** No formula bar, no formatting UI, no save-in-place. Editing exists to drive recalc, not to replace Excel |
| R7 | ExcelJS may not translate shared-formula refs (§7.9 **?**) | Test explicitly in P1; if it hands back the master formula verbatim, we translate offsets ourselves — or fall to tier 2/3 for such files |
| R8 | Write-back propagates *our* arithmetic into files (§13.2) | Gate on green eval; stamp provenance; keep `fullCalcOnLoad` so Excel recomputes anyway |

---

## 16. Open questions

1. **Which surface consumes this?** The spike is standalone. Does the calc layer
   land in a product preview path, or stay a library? Changes how much the
   what-if editing (§9.2) matters versus pure on-load rendering.
2. **Is view-only still the contract?** §9.2 is where the value is, and it is
   also the door to R6. Worth an explicit decision before P5.
3. **Excel-parity bar** — is LibreOffice good enough as truth (§12.5), or does
   some class of numbers need Excel-verified fixtures from the start?
4. **Wasm in the bundle: acceptable?** If a hard no, tier 2 becomes
   LibreOffice-only and the closed grammar carries more weight (widen it faster).
5. **Does the generator hand-write OOXML or use openpyxl?** Determines whether
   §13.3's `fullCalcOnLoad` assertion is needed at all (openpyxl already ✓).

---

## Appendix A — closed grammar v1 (the declared floor)

Everything outside this is **rejected at parse time** with a reason string, and
renders `⚠`. Widening is a deliberate, oracle-tested act (§5.1).

```ebnf
formula     = expr ;
expr        = compare ;
compare     = concat { ("=" | "<>" | "<" | ">" | "<=" | ">=") concat } ;
concat      = addsub { "&" addsub } ;
addsub      = muldiv { ("+" | "-") muldiv } ;
muldiv      = power { ("*" | "/") power } ;
power       = postfix { "^" postfix } ;              (* LEFT-assoc — §7.1 *)
postfix     = unary [ "%" ] ;
unary       = [ "-" | "+" ] primary ;                (* binds tighter than ^ — §7.1 *)
primary     = number | string | bool | error
            | reference | funcall | "(" expr ")" ;

reference   = [ sheet "!" ] cellref [ ":" cellref ] ;
sheet       = ident | "'" any-but-quote "'" ;        (* quoting mandatory w/ spaces *)
cellref     = [ "$" ] col [ "$" ] row ;

funcall     = fname "(" [ arglist ] ")" ;
arglist     = expr { "," expr } ;
fname       = "SUM" | "IF" ;                         (* v1 — the measured set, §2.1 *)
```

**v1 semantics required:** empty-cell = 0 and = `""` (§7.3) · arithmetic coerces
numeric text, comparison does not (§7.2) · `IF` lazy in the untaken branch
(§7.8) · `SUM` skips text/booleans in ranges (§7.8) · errors propagate, are
values not failures (§7.6) · unary-minus/`^` precedence and `^` left-assoc (§7.1).

**Explicitly rejected in v1:** `INDIRECT`, `OFFSET`, whole-column refs, defined
names, 3D refs, structured table refs, intersection/union operators, array &
dynamic-array formulas, shared formulas we can't translate, the 1904 date system,
external links, iterative calculation of cycles.

**v2 candidates, in the order the histogram will probably demand them:**
`IFERROR`, `ROUND`, `AVERAGE`, `MIN`, `MAX`, `ABS`, `SUMIF(S)`, `COUNTIF(S)`,
`INDEX`/`MATCH` (exact only), `NPV`, `IRR`, `PMT`, `POWER`, `EOMONTH`, `DATE`,
`YEAR`, `TEXT`, whole-column refs, defined names.

---

## Appendix B — raw evidence (all ✓, 2026-07-25)

**Uncached vs cached, same cell, same model:**
```xml
nocache : <c r="B7" s="19"><f>B6/B4</f><v></v></c>
recalc'd: <c r="B7" s="19" t="n"><f aca="false">B6/B4</f><v>0.59004854368932</v></c>
```

**Legitimately-empty vs never-computed (the `t` attribute — §2.2):**
```xml
recalc'd: <c r="G9" t="str"><f>IF(G9=0,"",F9/G9-1)</f><v></v></c>   ← computed to ""
nocache : <c r="F4"><f>SUM(B4:E4)</f><v></v></c>                     ← never computed
```

**`calcPr`:**
```xml
nocache (openpyxl): <calcPr calcId="124519" fullCalcOnLoad="1"/>
recalc'd (LO):      <calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.0001"/>
```

**Vocabulary histogram:** `financial-model.xlsx` — 76 formulas, 5 uncached
(legit-empty), 0 shared-only, 19 cross-sheet cells; functions `IF`×12, `SUM`×11;
shapes arithmetic ×57, cross-sheet ×19, string literal ×12, range ×11,
absolute ×11.

**`calcChain.xml`:** absent from both files.

**Renderer baseline (Chrome):** 12.4 KB, 3 sheets, 219 cells, 76 formulas,
parse 74 ms, 0 console errors; nocache → 75/75 uncached, banner fires.

**npm registry, 2026-07-25:** `formualizer@0.7.1` MIT OR Apache-2.0 ·
`@ironcalc/wasm@0.7.0` MIT/Apache-2.0 · `hyperformula@3.3.0` GPL-3.0-only ·
`@formulajs/formulajs@4.6.0` MIT · `fast-formula-parser@1.0.19` MIT
(last publish ≈6 y ago).

---

## Appendix C — sources

- HyperFormula licensing — https://hyperformula.handsontable.com/docs/guide/licensing.html
- Formualizer — https://github.com/psu3d0/formualizer · wasm bindings: https://github.com/PSU3D0/formualizer/blob/main/bindings/wasm/README.md · https://www.formualizer.dev/
- IronCalc — https://github.com/ironcalc/IronCalc · https://nlnet.nl/project/IronCalc/
- Formula.js — https://www.npmjs.com/package/@formulajs/formulajs
- fast-formula-parser — https://github.com/LesterLyu/fast-formula-parser
- Subject repo: `/Users/vz/OSS excel render tests` — `SPIKE_FINDINGS.md`, `LEARNINGS.md`, `src/excel/{parse,format,theme}.ts`, `ExcelView.tsx`
