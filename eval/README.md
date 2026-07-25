# Eval corpus

Ten synthetic workbooks, written the way an agent writes them, loaded the way a
user loads them, and checked cell by cell against LibreOffice.

```
python3 eval/build.py     # generate the corpus and recalculate it (needs LibreOffice)
npm run eval              # compare, and write REPORT.md
```

The current scoreboard is in [REPORT.md](REPORT.md). It is committed, so a change
in behaviour shows up as a diff rather than as a number nobody wrote down.

## Why this exists

`tools/oracle` probes one formula at a time against a shared grid and answers
*is this function right*. Useful, and not sufficient: every probe in it was
chosen by the same person who wrote the engine, which makes it a test of
intentions. It cannot fail in a way its author did not already think of.

This corpus answers a different question — *does a whole workbook come out
right* — and it is built to be able to surprise its author:

- The models are **domain models first**. They were written as a budget, an LBO,
  a cohort triangle, with the formulas that job needs, and were not checked
  against `CAPABILITY.md` while being written. Writing toward the supported list
  guarantees a green run and measures nothing.
- They exercise the **whole system**, not the library: the zip reader, the raw
  `t` attribute, shared-formula translation, cross-sheet binding, evaluation
  order, and only then the functions.
- They are emitted by **openpyxl with no cached values**, which is exactly the
  artefact this project exists to render.

## The ten models

| | model | what it is | what it is really testing |
|---|---|---|---|
| M01 | `m01_budget` | Annual budget, actuals, variance | Cross-sheet refs, running YTD totals, sign conventions, text flags |
| M02 | `m02_valuation_dcf` | DCF with a 35-cell sensitivity grid | NPV/XNPV/XIRR, `^` precedence, mixed anchoring (`$A12` × `C$4`) |
| M03 | `m03_lbo` | LBO, debt schedule, cash sweep | MIN/MAX waterfalls, two sheets referencing each other, IRR twice over |
| M04 | `m04_three_statement` | Linked IS / BS / CF | Evaluation order — the balance check only passes if ordering is right |
| M05 | `m05_workflow_approvals` | 160-row approval register | Nested IF, IFS, SWITCH, AND/OR, WORKDAY, COUNTIFS, lazy branches |
| M06 | `m06_sales_data` | 2,000 transactions + analysis | Scale (24k formulas), criteria strings, wildcards, date ranges |
| M07 | `m07_cohort_retention` | 18×18 cohort triangle | Blank vs `""` vs zero — half the grid is empty by construction |
| M08 | `m08_loan_amortization` | 360-month mortgage | A 360-link chain in one column; PMT/IPMT/CUMIPMT/RATE/NPER |
| M09 | `m09_inventory_planning` | 150 SKUs, safety stock, ABC | The statistical end of the library, which finance models never reach |
| M10 | `m10_edge_cases` | 136 labelled probes | Semantics, errors, and things that *should* refuse |

M10 is the odd one out and deliberately so. Its three groups are judged
differently: **semantics** must match exactly, **errors** like `#DIV/0!` are
computed values that must be produced rather than refused, and **refusals** like
`INDIRECT` or a circular reference are cells where ⚠ is the right answer and a
number is the failure. The point is that those three must not leak into each
other.

## How a cell is judged

Every formula cell lands in exactly one bucket:

| bucket | meaning |
|---|---|
| `match` | We and LibreOffice agree. |
| `inherited` | We disagree, but our formula run against *the oracle's* inputs reproduces the oracle's answer — so this cell computed correctly and the difference arrived from upstream. |
| `refused` | We declined, and said why. Correct for `INDIRECT`; a coverage gap anywhere else. Never a wrong number. |
| `circular` | We detected a cycle. |
| `volatile` | Depends on `NOW`/`TODAY`, so it cannot agree with an oracle computed at another moment. Excluded, not excused. |
| `divergence` | We disagree and we follow Excel on purpose. Declared in [`divergences.json`](divergences.json). |
| **`MISMATCH`** | We disagree and we did not know. **Hard gate at zero.** |

The divergence gate is **symmetric**: a declared divergence that starts
*matching* also fails. That is not pedantry — it is what caught the largest
finding below, after the fix that made the declaration obsolete.

`divergences.json` carries a `confidence` on every entry, because the honest
position differs case by case. Two entries are marked `unverified`: we believe
we follow Excel but have not run the cell in Excel itself. They are declared so
the gate stays useful and flagged so the uncertainty stays visible.

## What the first run found

The corpus was run for the first time before any of it was tuned. 37,098 formula
cells, 139 disagreements, and they resolved into five causes — four engine bugs
and one methodology limit.

**1. Numeric criteria matched non-numeric cells.** `COUNTIF(range,">0")` counted
text and empty strings, because the `*IF` family was using Excel's *general*
comparison ordering, where number < text < boolean and so `"label" > 0` is
genuinely TRUE. Excel's criteria comparison is type-scoped instead. Every
conditional sum over a column containing a header, a label or an `""` was
inflated. In M07 it silently overstated every retention denominator — 57 cells,
all plausible, none visibly wrong.

**2. `COUNTIF` did not broadcast over an array criteria.** Handed a range where
it expects one criterion, Excel evaluates once per element and returns an array;
we collapsed to the first cell. That breaks `SUMPRODUCT(1/COUNTIF(r,r))`, the
standard distinct-count idiom, into a confident wrong number rather than an
error.

**3. Arithmetic overflow produced `Infinity`.** `=1E+308*10` rendered as a
number. Excel gives `#NUM!`. `^` already had the guard; `+ - * /` did not.

**4. `ROUND` reconstructed its result one ulp off.** `716 * 10**-1` is
71.60000000000001, not 71.6. Individually invisible; in a debt schedule where
one rounded line feeds the next, the error walked downstream until a later
`ROUND` sat on a half-way boundary and flipped by a full 0.1. **This single ulp
accounted for 67 of the 139 disagreements** — the entire M03 cluster.

That last one is why the symmetric gate matters. The knife-edge cell had been
declared as an unavoidable oracle-precision divergence, with a plausible
paragraph explaining why nobody could do better. Fixing the ulp made it match,
the symmetric gate failed on the now-false declaration, and the explanation
turned out to be an excuse for a real bug.

**5. Two missing functions blanked 44% of a model.** `NORM.S.INV` and
`FORECAST.LINEAR` were not implemented. 186 refusals, poisoning 1,301 cells —
seven cells dark for every one actually missing, because uncomputable poisons
downstream by design. Neither function is exotic; both are what you reach for
the moment a model is about operations rather than finance. The library was
shaped like the tests written for it. Both are now implemented, with the CDF
accurate to ~1e-15 into the tail, where service levels actually live.

The methodology limit, worth knowing before trusting any number here:
**LibreOffice writes at most 15 significant digits into the `.xlsx`.** Excel
writes up to 17. So the oracle cannot carry full double precision, and any model
with a knife-edge — a `ROUND` on a half-way boundary, a criterion built by
concatenating a computed number — can disagree for reasons that are nobody's
bug. The `inherited` bucket exists to separate that class out by re-running our
formula against the oracle's own inputs.

## Adding a model

Drop a `models/mNN_name.py` exporting `build() -> Workbook`, add it to `MODELS`
in `build.py` and `eval.test.ts`, and rebuild. Two rules:

1. **Write the model, not the test.** Reach for the formula the job needs. If
   the engine cannot do it, that is the finding.
2. **Never tune a fixture to go green.** A model that has to be adjusted until
   it passes has stopped being evidence.

## Known limits

- **The corpus is still synthetic.** These are workbooks written by an agent in
  the shape of real ones, which is a much better sample than hand-picked probes
  and is still not a sample of production output. Widening it with real
  generated models remains the highest-value next step.
- **The oracle is LibreOffice, not Excel.** It is a second opinion, not ground
  truth. Where the two are known to differ, we follow Excel and declare it.
- **`RAND`/`RANDBETWEEN` are absent** by construction: they cannot agree with an
  oracle, so including them would add noise rather than coverage.
