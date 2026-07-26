# xlsx renderer & calculator

A client-side, view-only `.xlsx` preview that **computes formula values in the
browser**. No server, nothing uploaded, MIT.

**The job:** a user asked an agent for a financial model, the agent produced an
`.xlsx`, and the user needs to look at it and decide whether it is right —
without downloading it, without opening Excel, and without the file leaving their
machine. Review, not editing; editing belongs in Excel.
([Full scope in `DESIGN.md`.](./DESIGN.md#scope-and-the-job-to-be-done))

The problem it exists for: an agent writing a workbook emits formulas with an
empty value cache —

```xml
<c r="F4"><f>SUM(B4:E4)</f><v></v></c>
```

Excel computes on open. Every other consumer reads `<v>` and shows nothing, so a
generated financial model renders as a labelled skeleton. The previous fix was a
server-side LibreOffice recalc, which costs the preview its two defining
properties: no server, and nothing leaves the browser.

**Status: built and verified.** 153 tests green; 399 oracle probes, a
**37,098-cell synthetic corpus of ten whole workbooks**, and a
**202,795-cell corpus of ten *real* workbooks** — including a 138,421-formula
business plan last saved by Excel itself — with **zero unexplained
disagreements** anywhere and every deliberate divergence declared and counted;
and a real-browser check confirming that a file with *no* cached values renders
identical numbers to the same file after a LibreOffice recalc.

```
npm install
npm run verify                 # typecheck, test, build, and smoke the published packages
npm run dev                    # the preview at http://localhost:5176
python3 tools/oracle/generate.py   # rebuild oracle fixtures (needs LibreOffice)
python3 eval/build.py              # rebuild the eval corpus (needs LibreOffice)
npm run eval:real                  # the real-workbook corpus — private, not in this repo
node tools/verify/drive.mjs        # end-to-end check in a real browser
```

---

## Using it

Two packages, published separately, MIT.

```bash
npm install @xlscalc/xlsx-preview          # load, compute, render
npm install @xlscalc/formula-engine        # just the evaluator, zero dependencies
```

```tsx
import { loadXlsx } from '@xlscalc/xlsx-preview';
import { ExcelView } from '@xlscalc/xlsx-preview/view';
import '@xlscalc/xlsx-preview/view/style.css';   // required: ⚠ is invisible without it

const doc = await loadXlsx(await file.arrayBuffer());
<ExcelView doc={doc} sheet={0} />
```

For a workbook of any size, run it off the main thread — same component, same
props:

```ts
import { createPreviewWorker } from '@xlscalc/xlsx-preview/worker';
const doc = await createPreviewWorker().load(await file.arrayBuffer());
```

And before committing to any of it, ask what will be refused. This parses every
formula but evaluates nothing and never loads the styling dependency — 0.8 s
where a full load of the same file takes 7.5:

```ts
import { inspectXlsx } from '@xlscalc/xlsx-preview';
const { fullyCovered, unsupported, iterative } = inspectXlsx(buf);
```

There is also `renderToHtml(doc, sheet)` for hosts that are not React, and
`examples/node-headless.mjs` for servers. Per-package documentation is in
[packages/xlsx-preview](./packages/xlsx-preview/README.md) and
[packages/formula-engine](./packages/formula-engine/README.md); release policy —
including why adding a function is a breaking change to what appears on screen —
is in [VERSIONING.md](./VERSIONING.md).

**The number to know before adopting: about 36% of formula cells in the real
corpus render ⚠.** Almost none of that is a missing function; refusal
propagates, so in the worst workbook 33 cells using `OFFSET` and `CELL` left
64,809 warning. Whether that is acceptable depends on your files, which is what
`inspectXlsx` is for. On the 128,976 cells it does answer there are **zero
unexplained disagreements**; 3,050 of them differ from the value the file itself
stored, and each falls under a documented rule with an exact expected count.

---

## The one rule

> **A cell we cannot compute renders ⚠, never a number.**

Everything else is subordinate to that. In a financial model a wrong number is
invisible and unfalsifiable by eye, so an engine that guesses is worse than one
that shows a blank. Unsupported functions, refused reference shapes, cycles and
uncertain coercion all become a marked cell with a stated reason, counted at file
level and listed in the UI.

Two corollaries that are easy to get wrong and are enforced in the code:

- **An Excel error is a computed value, not a failure.** `#DIV/0!` from a genuine
  divide-by-zero renders as `#DIV/0!` with provenance `computed`. Only *our*
  inability to evaluate produces ⚠. Conflating the two would make the honesty
  layer meaningless — a user could not tell "your model divides by zero here"
  from "this renderer gave up here".
- **Uncomputable poisons downstream.** If a cell cannot be computed, every cell
  that reads it is marked too. Letting a `SUM` silently treat an uncomputable
  precedent as zero is exactly the confident wrong subtotal this project exists
  to prevent.

---

## Layout

```
packages/formula-engine   zero-dependency Excel formula parser + evaluator
packages/xlsx-preview     xlsx binding, value overlay, layout, audit pass, renderers
apps/demo                 the preview app on :5176 — loads through the Worker
examples                  server-side script and the typechecked README quickstart
tools/oracle              LibreOffice differential harness + sample generators
tools/verify              real-browser end-to-end, screenshots, published-package smoke
```

`formula-engine` has **no dependencies at all** and runs in Node, a worker or the
browser. `xlsx-preview` adds ExcelJS (styles), numfmt (number formats) and fflate
(unzip). All three are MIT; ExcelJS brings a further ~80 transitive packages
whose licences are permissive but not uniformly MIT — Apache-2.0, ISC, BSD-3,
one Unlicense, one dual MIT-or-GPL. It also brings their advisories: `npm audit`
reports high-severity findings in `minimatch` and friends, none of them on a code
path this library calls. That tree is the price of ExcelJS's style fidelity, and
it is the single largest thing to weigh before installing `xlsx-preview`; the
engine on its own has none of it.

Inside `xlsx-preview` the boundary that matters is `layout.ts`: it reads a
sheet's whole appearance out of ExcelJS once, at load, into plain data. Nothing
downstream of it knows ExcelJS exists, which is what lets the renderer take a
document instead of a worksheet, lets the result cross a `postMessage`, and
keeps the styling dependency out of the main-thread bundle entirely.

---

## How it works

```
.xlsx ──┬─► ExcelJS ──────────► styles, fonts, merges, widths, panes
        └─► ooxml.ts ─────────► formulas, raw `t`, shared-formula translation
                                          │
                                   formula-engine
                                          │
                                    ValueOverlay ──► ExcelView
                                    (value + provenance + the file's own value)
```

**Two passes, each doing what it is best at.** ExcelJS's style fidelity is proven
and rewriting `styles.xml` handling would be risk without upside. But ExcelJS
drops the raw `t` attribute, and that single attribute decides the question the
whole recalc story turns on:

```xml
<c r="G9" t="str"><f>IF(G9=0,"",F9/G9-1)</f><v></v></c>   computed to ""
<c r="F4"><f>SUM(B4:E4)</f><v></v></c>                     never computed
```

Per cell those are indistinguishable without `t`, which is why the earlier
renderer had to guess at file level ("more than 25 % of formulas look uncached →
assume the file needs a recalc"). Reading the part directly replaces the guess
with the fact, and the heuristic is gone.

The same pass fixes shared formulas. OOXML stores `<f t="shared" ref="B2:E2"
si="0">` once and siblings reference it by `si`; a reader that returns the master
text verbatim reports four cells as computing the first cell's numbers. Here the
master's AST is translated by each sibling's offset.

**The overlay is a side table, never a mutation.** Computed values never
overwrite the file's cached ones, so every rendered number knows where it came
from, and the two can be compared.

---

## Correctness: the oracle

The engine is scored against **LibreOffice**, which recalculates the same probe
workbook and writes its answers into the fixture. Every probe lands in exactly
one bucket:

| bucket | meaning |
|---|---|
| **match** | the two agree |
| **unsupported** | we refused, and said why — acceptable, tracked |
| **divergence** | they disagree and we deliberately follow Excel — declared in `divergences.json` |
| **MISMATCH** | they disagree and we did not know — **hard gate at zero** |

Current run — 399 probes across 13 suites:

```
coverage 100.0%   accuracy 100.0%   false confidence 0
```

The probes target where an engine goes *quietly* wrong rather than covering
functions by count: coercion, blank-versus-empty-string, decimal rounding
boundaries, sign conventions, error propagation, date serials.

**LibreOffice is not Excel**, and pretending otherwise would launder its
behaviour into ours. The 29 declared divergences are the interesting output of
this harness, not noise. The largest class: *LibreOffice treats booleans as plain
numbers.* Excel gives them their own type, ranking above text and above every
number, so `TRUE=1` is FALSE, `SUM` ignores booleans inside a range, and `COUNT`
does not count them. Others: LibreOffice has no phantom 1900-02-29, so every date
before 1900-03-01 sits one day off; `DATE(26,1,1)` is 1926 per Microsoft's
documented rule, not 2026; `SQRT(-1)` is `#NUM!`, not `#VALUE!`.

A declared divergence that starts *matching* also fails the gate — that means
either LibreOffice changed or we drifted onto its side of the argument. This is
not pedantry: it is what caught the largest bug the corpus has found so far. A
knife-edge cell had been declared an unavoidable oracle-precision divergence,
with a plausible paragraph explaining why nobody could do better; an unrelated
one-ulp fix in `ROUND` made it match, the symmetric gate failed on the
now-false declaration, and the explanation turned out to be an excuse for a real
bug affecting 67 cells.

### The second harness: ten whole workbooks

The probe suite answers *is this function right*. It cannot answer *does a whole
workbook come out right*, and every probe in it was chosen by the same person
who wrote the engine. [`eval/`](eval/README.md) is the corpus that can surprise
its author: ten realistic models — budget, DCF, LBO, three-statement, approval
workflow, 2,000-row sales ledger, cohort triangle, 360-month amortisation,
inventory plan, and an adversarial sheet — written as domain models without
consulting the capability list, then compared cell by cell against LibreOffice.

```
37,098 formula cells   coverage 100.0%   accuracy 100.0%   false confidence 0
```

Its first run, before anything was tuned, found four engine bugs and two missing
functions. The most instructive: `COUNTIF(range,">0")` was counting text and
empty strings, because the `*IF` family was using Excel's *general* comparison
ordering — where `"label" > 0` really is TRUE — instead of Excel's type-scoped
criteria comparison. Every conditional sum over a column with a header or an
`""` in it was inflated, plausibly, invisibly.

### The third harness: workbooks nobody wrote for us

A synthetic corpus is a much better sample than hand-picked probes and is still
not production output — which [`eval/`](eval/README.md) said about itself. The
third harness is that sample: ten workbooks written by other people and other
tools, graded against the answers their own applications computed. Provenance is
load-bearing, because it decides what a cache is worth. One of them was last
saved by Microsoft Excel, which makes its 138,421 cached values the best ground
truth this project has had.

```
202,795 formula cells   answered 63.6%   unexplained 0   declared divergences 3,050
```

> **This corpus is not in the repository, and will not be.** They are real
> businesses' budgets, board packs and business plans, shared for a technical
> purpose and not for publication — so the workbooks, the harness that names
> them, the per-cell dumps and the report are all private. What is published is
> what this section says: aggregates, and the bugs they found.
>
> That is a real weakening of the evidence and it should be read as one. Every
> number above is *reported* rather than reproducible by a reader, and you have
> only this repository's word for it. The two harnesses that **are** reproducible
> — the 399-probe oracle and the 37,098-cell synthetic corpus, both scored
> against LibreOffice — are the ones to judge the engine on if you would rather
> not take a claim on trust. `npm run oracle` and `npm run eval` regenerate both
> from source, given a LibreOffice install.

It found eight more bugs across two passes, most of which the synthetic corpus
could not have found in principle. The worst: **every formula in a shared range was reconstructed
without its parentheses.** OOXML stores `<f t="shared">` once and offsets it for
each sibling; that path goes parse → translate → unparse, and `unparse` emitted
binary operators with no brackets, because the grammar has no parenthesis node.
`O19*(1-$E$17)` came back as `O19*1-$E$17` — no error, a number two and a half
times too large. openpyxl writes every formula out in full, so no synthetic
fixture could ever have exercised it.

The finding that is not a bug: **36.4 % of that corpus renders ⚠**, and almost
none of it is a missing function. In the business plan, **33 unsupported roots —
21 `OFFSET`, 12 `CELL` — darken 64,809 cells**, just under half the workbook.
Poisoning downstream is the right design; 1,964 dark cells per unsupported root
is the number that decides whether it is usable.

A second pass over the corpus fixed the reference model behind one of those
warnings. A whole-column reference like `A:A` names 1,048,576 rows; the engine
clamps it to the used range because iterating the rest is unaffordable and
pointless. That clamp is a property of *iteration*, and it had leaked into the
functions that report an extent or index by position — `ROWS(A:A)` answered 80
instead of 1,048,576, and `INDEX('Sheet'!$D:$O, 63, 1)` answered `#REF!` where
Excel reads an empty cell. A reference now carries the extent it was *declared*
with alongside the rectangle we will actually read.

Three behaviours the oracle pinned that would otherwise have been guesses:
General-format text switches to scientific notation at `1E16` and `1E-15`
(measured by bracketing, not assumed), and `ROUND(2.675,2)` is `2.68` — which
requires rounding the *decimal* value, since 2.675 is really 2.67499999999999982
in IEEE-754 and `Math.round(2.675*100)/100` gives 2.67.

---

## What it computes

**204 functions** — the exact list is in [`CAPABILITY.md`](./CAPABILITY.md),
generated from the function registry and asserted on every test run, so it cannot
drift from the code. It doubles as an allowlist that can be published into the
generator's prompt, so the renderer never meets a formula it cannot compute.

They were chosen for what a generated financial model actually
reaches for: math and rounding, statistics, the `*IF`/`*IFS` family with the
criteria mini-language and wildcards, logical (with genuinely lazy `IF`,
`IFERROR`, `IFS`, `CHOOSE`, `SWITCH`), text, dates including the 1900 leap-year
bug and the 1904 system, lookups, and the financial set (`NPV`, `XNPV`, `IRR`,
`XIRR`, `MIRR`, `PMT`/`IPMT`/`PPMT`, `PV`/`FV`/`NPER`/`RATE`, `CUMIPMT`,
`CUMPRINC`, depreciation, `EFFECT`, `RRI`, `PDURATION`).

Operators follow Excel exactly, including the two traps that produce a plausible
wrong number rather than an error:

```
-2^2  = 4      unary minus binds tighter than ^ — unlike mathematics
                and unlike every C-family language
2^3^2 = 64     ^ is LEFT-associative in Excel, not right
```

### What it deliberately refuses

26 recognised functions plus two conditional cases. Each renders ⚠ with the
reason and is listed in the UI:

- `INDIRECT` and `OFFSET` — they build references at evaluation time, which
  breaks the static dependency graph.
- **Approximate-match lookups over unsorted data.** `VLOOKUP`'s omitted 4th
  argument is a binary search, defined only on sorted input; on unsorted data
  Excel returns whatever its probe sequence lands on. The sorted contract is
  implemented exactly, and unsorted input is refused rather than answered with a
  plausible wrong row.
- `SUBTOTAL(101–111)` — they exclude manually hidden rows, which the engine
  cannot see. Answering would mean assuming nothing is hidden.
- Dynamic arrays (`FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`) and other spilling
  functions — they change the grid, not one cell's value.
- External workbook links, structured table references, 3D references.
- `TEXT()` when no number formatter is installed. The engine asks its host for
  one rather than shipping a partial format implementation that would silently
  mis-render `[Red](#,##0)`; the renderer installs numfmt, so in the app it works.

---

## The audit pass

Once both the computed and the stated value exist, disagreement is a **finding**.
For a model an agent wrote, this is arguably worth more than the rendering fix:
"these numbers are internally consistent" is a claim nobody could previously make
cheaply.

Two shapes, and only one of them needs a value comparison:

- **Stale cache** — a formula cell whose stored `<v>` disagrees with what the
  formula computes.
- **Hardcoded cell** — a literal number sitting where its neighbours hold
  formulas. This is the characteristic failure of a generated model: the writer
  computed one figure in its head and typed it in. There is no formula to
  disagree with, so no value comparison can find it — it is visible only in the
  *shape* of the neighbourhood.

The detector infers what the cell would have held by translating a neighbour's
formula to its position and evaluating it, so the finding is specific rather than
a vague suspicion. It requires at least two neighbours sharing one shape, because
a single neighbouring formula is a pattern of one and would fire on every summary
row. On the bundled sample:

> `Revenue!E6` is typed in as **761.2**, but the rest of the row (B6, C6, D6) uses
> `=SUM(E4:E5)`, which is **721.2**

---

## Performance

Measured, by formula *shape* rather than by count (`packages/formula-engine/src/perf.test.ts`):

| shape | size | evaluate |
|---|---:|---:|
| local arithmetic (the normal case) | 25,000 formulas | **~100 ms** |
| running total `SUM($A$1:A_n)` per row | 5,000 formulas | ~1.1 s |
| full-table `VLOOKUP` per row | 2,000 formulas | ~0.4 s |
| 5,000-deep row-to-row chain | 5,000 formulas | ~10 ms |

The last two rows are quadratic *by construction* — a running total reads every
earlier row — so they cost O(rows²) cell reads for any evaluator, Excel included.

Two findings from getting there, both recorded in the code:

- **Materialising the dependency graph is a trap.** One edge per precedent cell
  is O(rows²) Set entries; at 5,000 rows with a running total that is 12.5 M
  entries, which took a 45,000-formula workbook from slow to out-of-memory. The
  precedent *rectangles* are kept instead and walked lazily by an iterative
  depth-first search whose post-order is a topological order. Memory is
  O(formulas).
- **The search is iterative, not recursive**, because a row-to-row chain is as
  deep as the model is long and would overflow the call stack. The 5,000-deep
  chain test guards that.

---

## Known gaps

Stated plainly rather than buried:

- **The grid is drawn in full — there is no virtualisation.** Every cell of a
  sheet's extent becomes a `<td>`, which is fine for a financial model and fatal
  for a stray: `rowCount` is the last row with *anything* in it, so one
  accidental value in the bottom-right corner declares 1,048,576 × 16,384 and
  seventeen billion cells is an out-of-memory crash, not a slow render. Both
  renderers therefore stop at `DEFAULT_RENDER_LIMITS` (~150,000 cells) and
  caption what they left out; a host that knows its files can pass its own
  `limits`. Windowing the grid instead is the real fix and is not built.
- **The Worker is opt-in, and `loadXlsx` on its own still blocks.** The
  flattening this needed — `layout.ts` and `snapshot.ts` — is done, and
  `createPreviewWorker()` moves the whole load off-thread with no change to the
  renderer. But a caller who reaches for `loadXlsx` directly gets the
  synchronous path, which on a 138,000-formula workbook is 7.5 seconds of frozen
  tab. Making the off-thread path the default would mean the base entry point
  could no longer be used server-side, so for now the choice is the caller's and
  the demo makes it the visible one.
- **No what-if editing.** The engine has the primitive (`tryEvaluate` runs a
  formula at a position without storing it) and the graph supports dirty-set
  propagation, but no edit UI is wired up. View-only is the current contract.
- **No write-back.** Injecting computed values as `<v>` so downstream consumers
  see numbers is designed (`DESIGN.md` §13.2) but not built. It should stay gated
  on a green oracle score, because writing our arithmetic into a file makes any
  disagreement with Excel permanent.
- **Custom themes are not read.** `theme`/`indexed` colours resolve against the
  default Office palette; a workbook with an embedded custom theme in
  `xl/theme/theme1.xml` will render its accent colours slightly off.
- **In-sheet charts are not drawn.**
- **The real corpus is ten files, only one of them Excel-authored — and it is
  private.** It closed the "no production output" gap — 202,795 cells written by
  other people and other tools — but ten workbooks are wide, not random.
  Everything graded against a Google Sheets export or a generator is graded
  against a second opinion; only the business plan Excel saved is ground truth. A
  file's cache can also be stale, recording what its application last computed
  rather than what it would compute today. Nothing here looked stale; nothing
  rules it out. And because the workbooks are other people's confidential
  business data, none of that is in this repository for you to check — see the
  note above.
- **No iterative calculation.** A workbook that sets `iterate="1"` is asking for
  its circular references to be converged, which is the ordinary shape of
  interest on an average balance. We refuse them. The refusal is correct for an
  engine that does not iterate, but it is a capability gap and not a broken
  model — 4,194 cells in one real workbook.
- **The oracle carries only 15 significant digits.** LibreOffice writes at most
  15 into the `.xlsx`; Excel writes up to 17. A model with a knife-edge — a
  `ROUND` on a half-way boundary, a criterion built by concatenating a computed
  number — can therefore disagree with the oracle for reasons that are nobody's
  bug. The eval harness separates that class out rather than hiding it.

---

## Why from scratch

`DESIGN.md` weighed adopting a wasm engine (`formualizer`, MIT/Apache-2.0) and
concluded the closed-vocabulary path was viable for our own files. Two things
decided the build:

1. **Fail-loud needs introspection.** The doctrine is worth nothing unless the
   engine can say precisely what it could not do. An adopted engine reports a
   value; it cannot tell you which of its answers you should not trust.
2. **The vocabulary is not closed.** `DESIGN.md` §2.1 measured a 76-formula
   sample and found exactly two functions, `SUM` and `IF`. But the generator
   path is `python_execute` + openpyxl written free-hand by a model, so the
   vocabulary is whatever an LLM types. A `{SUM, IF}` engine would have rendered
   ⚠ on the first real model. That is why the library is broad and gated by an
   oracle rather than frozen against one sample.

LibreOffice headless stays available as the tier-3 escalation for files this
engine refuses — now an explicit, user-consented step rather than the default
path, which preserves the no-egress property for everything else.
