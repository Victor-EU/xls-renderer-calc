# xlsx renderer & calculator

A client-side, view-only `.xlsx` preview that **computes formula values in the
browser**. No server, nothing uploaded, MIT throughout.

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

**Status: built and verified.** 80 tests green, 399 oracle probes at 100 %
accuracy and zero false confidence, and a real-browser check confirming that a
file with *no* cached values renders identical numbers to the same file after a
LibreOffice recalc.

```
npm install
npm test                       # engine + oracle
npm run dev                    # the preview at http://localhost:5176
python3 tools/oracle/generate.py   # rebuild oracle fixtures (needs LibreOffice)
node tools/verify/drive.mjs        # end-to-end check in a real browser
```

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
packages/xlsx-preview     xlsx binding, value overlay, audit pass, React renderer
apps/demo                 the preview app on :5176
tools/oracle              LibreOffice differential harness + sample generators
tools/verify              real-browser end-to-end and screenshot drivers
```

`formula-engine` has **no dependencies at all** and runs in Node, a worker or the
browser. `xlsx-preview` adds ExcelJS (styles), numfmt (number formats) and fflate
(unzip) — all MIT.

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
either LibreOffice changed or we drifted onto its side of the argument.

Three behaviours the oracle pinned that would otherwise have been guesses:
General-format text switches to scientific notation at `1E16` and `1E-15`
(measured by bracketing, not assumed), and `ROUND(2.675,2)` is `2.68` — which
requires rounding the *decimal* value, since 2.675 is really 2.67499999999999982
in IEEE-754 and `Math.round(2.675*100)/100` gives 2.67.

---

## What it computes

**197 functions** — the exact list is in [`CAPABILITY.md`](./CAPABILITY.md),
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

- **No Web Worker yet.** Parse and evaluation run on the main thread. At the
  measured speeds that is imperceptible for a normal model (a 25,000-formula
  workbook evaluates in ~100 ms), but a pathological one can block for a second
  or more. Moving it off-thread is not a wrapper: the renderer reads the ExcelJS
  worksheet directly, so the styled grid would first have to be flattened into a
  transferable snapshot. That is a real refactor with fidelity regression risk,
  and it was left undone rather than done shakily.
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
- **The oracle corpus is synthetic.** 399 probes plus four sample workbooks. It
  should be widened with real agent output before the function coverage is
  treated as settled — the vocabulary claim is the weakest evidence here.

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
