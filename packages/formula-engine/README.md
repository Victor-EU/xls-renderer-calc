# @xlscalc/formula-engine

An Excel formula parser and evaluator with **no dependencies** and one rule: it
never renders a number it is not sure of.

```bash
npm install @xlscalc/formula-engine
```

```ts
import { Workbook } from '@xlscalc/formula-engine';

const wb = new Workbook();
wb.addSheet('Revenue');

wb.setValue(0, 1, 1, 100);           // A1 = 100
wb.setValue(0, 2, 1, 250);           // A2 = 250
wb.setFormula(0, 3, 1, 'SUM(A1:A2)*1.2');

const report = wb.evaluateAll();

wb.record(0, 3, 1).value;            // 420
wb.record(0, 3, 1).provenance;       // 'computed'
report.stats;                        // { formulas: 1, computed: 1, unsupported: 0, … }
```

Rows and columns are 1-based, as they are in Excel. Sheets are 0-based indices
in the order they were added.

## Why it refuses

Every cell that comes back carries a **provenance**, and that is the whole
design rather than a diagnostic:

| provenance | meaning |
|---|---|
| `literal` | straight from the file; we did not touch it |
| `cached` | the file's own computed value, reused |
| `computed` | we evaluated it |
| `volatile` | we evaluated it, and it depends on the clock |
| `circular` | part of a reference cycle |
| `unsupported` | outside capability — **no value at all** |

`unsupported` never carries a number. There is no code path from "we could not
compute this" to a figure, and refusal propagates: a `SUM` over a cell we
declined to compute is itself refused rather than quietly dropping the input.
That is deliberate, and it is expensive — in the largest workbook of the
real-data corpus, 33 cells using `OFFSET` and `CELL` leave 64,809 cells
warning rather than numbered. A total that silently omits a term it could not
evaluate is precisely the failure this library exists to prevent.

Excel's *own* error values are a separate thing entirely. `#DIV/0!` is a
computed result and renders as itself; the two are never conflated.

## What it knows

**204 functions.** The list is not prose that can go stale — ask the library:

```ts
import { implementedFunctions, refusedFunctions } from '@xlscalc/formula-engine';

implementedFunctions();   // ['ABS', 'ACOS', 'AND', … ] — 204 of them
refusedFunctions();       // ['AGGREGATE', 'CELL', 'FILTER', … ] — 26
```

`CAPABILITY.md` ships inside this package and is generated from the same
registry, so diffing it between two installed versions shows exactly what
changed about what renders.

26 further functions are *known and deliberately refused* — `OFFSET`,
`INDIRECT`, `CELL`, `INFO`, the dynamic-array family — each with a stated
reason that reaches the tooltip. Refusing something by name is better than not
recognising it, because the reason can be specific.

Note that **adding a function changes what existing users see** — a workbook
that showed a screen of warnings starts showing numbers, with no change on the
caller's side. So capability growth is never a patch release; the repository's
`VERSIONING.md` states the policy.

## How it is checked

Not by hand-picked probes. Three harnesses, each answering a question the last
one could not:

1. **An oracle** — thousands of generated formulas, every answer compared
   against LibreOffice.
2. **Ten whole synthetic workbooks** — agent-written models in the shape of real
   ones, since a formula in isolation is not a formula in a model.
3. **Ten real workbooks** — 202,795 formula cells nobody wrote for this project,
   graded against the values their own applications computed. On the 128,976
   cells it answers there are **zero unexplained disagreements**. 3,050 of them
   deliberately differ from the value the file stored — a Google Sheets export
   treats a blank reference differently from Excel, and we follow Excel — and
   each falls under a named rule with an exact expected count, gated in both
   directions so a rule that starts explaining more cells, or fewer, fails the
   build.

The real corpus found five bugs on its first run and three more on a second
pass — most of them things the synthetic corpus could not have found in
principle. Those workbooks are other people's confidential business data, so
they are not published and neither is the harness that reads them; the two
harnesses you *can* rerun are the oracle and the synthetic corpus, both in the
repository.

## Beyond evaluating

```ts
import { parseFormula, unparse, walk } from '@xlscalc/formula-engine';

const ast = parseFormula('=SUM(A1:A3)*(1-$B$1)');
walk(ast, (n) => { if (n.k === 'fn') console.log(n.name); });   // SUM
unparse(ast);                                                   // SUM(A1:A3)*(1-$B$1)
```

`parseFormula(unparse(ast))` is asserted to give back the same tree. That is not
decoration: `unparse` is how shared formulas are reconstructed, and when it and
the parser held separate precedence tables they drifted — `O19*(1-$E$17)` came
back as `O19*1-$E$17`, no error, a number two and a half times too large.

## Requirements

Node 18+ or any modern browser. ESM only. TypeScript declarations included.
MIT.
